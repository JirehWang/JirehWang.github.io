/**
 * 教會行事曆 - 批量操作 + AI 解析（Phase 2 擴充）
 */

// ─────────────────────────────────────────────────────────────
//  批量新增事項
//  data: { events: [{typeId, date, title, values}] }
// ─────────────────────────────────────────────────────────────
function cal_addEventsBatch(data) {
  if (!data || !Array.isArray(data.events)) throw new Error('events 陣列必填');
  if (data.events.length === 0) return { success: true, created: 0, message: '無事項可建立' };

  const types = _cal_readAll(_CAL_SHEET.TYPES);
  const typeIds = new Set(types.map(t => t.typeId));

  const errors = [];
  const validEvents = [];
  data.events.forEach((e, idx) => {
    if (!e.typeId)         { errors.push(`第 ${idx+1} 筆：缺 typeId`); return; }
    if (!typeIds.has(e.typeId)) { errors.push(`第 ${idx+1} 筆：類型不存在`); return; }
    if (!e.date)           { errors.push(`第 ${idx+1} 筆：缺日期`); return; }
    validEvents.push(e);
  });

  if (validEvents.length === 0) {
    return { success: false, created: 0, errors: errors, message: '無有效事項可建立' };
  }

  const now = new Date();
  const eventRows = [];
  const valueRows = [];

  validEvents.forEach(e => {
    const eventId = Utilities.getUuid();
    const title = (e.title || '').toString().trim()
      || _cal_autoTitle(e.values, e.typeId, types)
      || '(無標題)';
    eventRows.push([eventId, e.typeId, e.date, title, e.createdBy || 'batch', now, now]);
    if (e.values && typeof e.values === 'object') {
      Object.entries(e.values).forEach(([fid, v]) => {
        if (v !== undefined && v !== null && v !== '') {
          const valStr = (typeof v === 'object') ? JSON.stringify(v) : String(v);
          valueRows.push([eventId, fid, valStr]);
        }
      });
    }
  });

  if (eventRows.length > 0) {
    const sh = _cal_getSheet(_CAL_SHEET.EVENTS);
    sh.getRange(sh.getLastRow() + 1, 1, eventRows.length, eventRows[0].length).setValues(eventRows);
  }
  if (valueRows.length > 0) {
    const sh = _cal_getSheet(_CAL_SHEET.VALUES);
    sh.getRange(sh.getLastRow() + 1, 1, valueRows.length, 3).setValues(valueRows);
  }

  return {
    success: true,
    created: validEvents.length,
    skipped: errors.length,
    errors: errors,
    message: `批量建立完成：成功 ${validEvents.length}，跳過 ${errors.length}`
  };
}

// ─────────────────────────────────────────────────────────────
//  AI 解析（依事項類型的欄位 schema）
//  data: { rootTypeId, rawText, allowMultiple }
//
//  回傳：{ success, events: [{date, subTypeId, subTypeName, title, values}] }
// ─────────────────────────────────────────────────────────────
function cal_aiParseForType(data) {
  if (!data || !data.rootTypeId) throw new Error('rootTypeId 必填');
  if (!data.rawText || !data.rawText.toString().trim()) throw new Error('rawText 必填');

  const types = _cal_readAll(_CAL_SHEET.TYPES);
  const rootType = types.find(t => t.typeId === data.rootTypeId);
  if (!rootType) throw new Error('找不到頂層類型');

  // 取子類型 + 欄位定義
  const subTypes = types.filter(t => t.parentTypeId === data.rootTypeId);
  const fields = _cal_readAll(_CAL_SHEET.FIELDS)
    .filter(f => f.typeId === data.rootTypeId)
    .sort((a, b) => (Number(a.sortOrder) || 0) - (Number(b.sortOrder) || 0));

  // 組 prompt
  let fieldsDesc = fields.map(f => {
    let line = `  - "${f.fieldId}": ${f['顯示名稱']} (${f['欄位類型']})`;
    if (String(f['是否必填']).toUpperCase() === 'TRUE') line += ' [必填]';
    const opts = _cal_safeParseOptions(f['下拉選項']);
    if (opts.length > 0) line += ` 選項=[${opts.join(', ')}]`;
    return line;
  }).join('\n');

  let subTypesDesc = '';
  if (subTypes.length > 0) {
    subTypesDesc = `\n此類型有以下子類型，請判斷每筆資料屬於哪個子類型：\n` +
      subTypes.map(s => `  - "${s['名稱']}"`).join('\n');
  }

  const allowMultiple = data.allowMultiple !== false;
  const prompt = `你是教會行事曆助理。請從輸入文字提取「${rootType['名稱']}」事項資訊，輸出 JSON。

可用欄位 (key 必須用 fieldId)：
${fieldsDesc}
${subTypesDesc}

規則：
1. 日期一律 YYYY-MM-DD 格式
2. ${allowMultiple ? '若文字描述多筆，輸出多個 event' : '只輸出 1 個 event'}
3. 必填欄位若文字中無資訊，填空字串
4. 若有子類型，請填入 "subTypeName" 值需完全符合上述清單
5. 標題（title）留空即可，系統會自動產生

請只輸出 JSON，不要其他文字、不要 markdown：
{
  "events": [
    {
      "date": "2026-01-05",
      "subTypeName": "${subTypes[0] ? subTypes[0]['名稱'] : ''}",
      "title": "",
      "values": {
        ${fields.slice(0,2).map(f => `"${f.fieldId}": ""`).join(', ')}
      }
    }
  ]
}`;

  // 呼叫 Gemini
  const aiResult = callGeminiApi(prompt, data.rawText);

  // 解析 AI 回傳
  let parsed;
  try {
    parsed = JSON.parse(aiResult);
  } catch (e) {
    // 容錯：嘗試擷取第一個 { 到最後一個 } 之間的內容
    const m = aiResult.match(/\{[\s\S]*\}/);
    if (m) { try { parsed = JSON.parse(m[0]); } catch (e2) {} }
    if (!parsed) throw new Error('AI 回傳無法解析為 JSON：' + aiResult.substring(0, 200));
  }

  const events = Array.isArray(parsed.events) ? parsed.events : [];

  // 把 subTypeName → subTypeId
  const subTypesByName = {};
  subTypes.forEach(s => subTypesByName[s['名稱']] = s.typeId);

  const enriched = events.map(ev => {
    const subTypeId = ev.subTypeName ? subTypesByName[ev.subTypeName] : null;
    return {
      date:        ev.date || '',
      subTypeName: ev.subTypeName || '',
      subTypeId:   subTypeId || (subTypes.length === 0 ? rootType.typeId : ''),
      title:       ev.title || '',
      values:      ev.values || {}
    };
  });

  return {
    success: true,
    rootTypeId: data.rootTypeId,
    rootTypeName: rootType['名稱'],
    events: enriched,
    hasSubTypes: subTypes.length > 0
  };
}
