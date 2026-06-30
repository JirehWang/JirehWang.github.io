/**
 * 教會行事曆 - 事項 CRUD（Phase 2）
 *
 * 事項本身存在「事項」sheet，欄位值用 long-format 存在「事項欄位值」sheet
 * 寫入 / 更新會自動同步兩張表
 */

// ─────────────────────────────────────────────────────────────
//  取得事項清單（含類型 metadata + 欄位值）
//  參數：{ startDate, endDate, typeIds[] }
// ─────────────────────────────────────────────────────────────
function cal_getEvents(data) {
  const startDate = data && data.startDate;
  const endDate   = data && data.endDate;
  const typeIds   = data && Array.isArray(data.typeIds) ? data.typeIds : null;

  const events = _cal_readAll(_CAL_SHEET.EVENTS);
  const types  = _cal_readAll(_CAL_SHEET.TYPES);
  const fields = _cal_readAll(_CAL_SHEET.FIELDS);
  const values = _cal_readAll(_CAL_SHEET.VALUES);

  const typeById  = {};
  types.forEach(t => typeById[t.typeId] = t);
  const fieldById = {};
  fields.forEach(f => fieldById[f.fieldId] = f);

  // 把欄位值按 eventId 分組
  const valuesByEvent = {};
  values.forEach(v => {
    if (!valuesByEvent[v.eventId]) valuesByEvent[v.eventId] = [];
    valuesByEvent[v.eventId].push(v);
  });

  // 推導每個事項顯示用 metadata（找頂層類型補圖示/顏色）
  const result = events
    .filter(e => {
      if (startDate && _cal_dateStr(e['日期']) < startDate) return false;
      if (endDate && _cal_dateStr(e['日期']) > endDate) return false;
      if (typeIds && typeIds.length > 0 && typeIds.indexOf(e.typeId) === -1) return false;
      return true;
    })
    .map(e => {
      const type = typeById[e.typeId];
      // 找頂層父類型（取圖示 / 同步設定）
      let rootType = type;
      while (rootType && rootType.parentTypeId) {
        rootType = typeById[rootType.parentTypeId] || rootType;
        if (rootType.parentTypeId === '' || !rootType.parentTypeId) break;
      }

      const evValues = (valuesByEvent[e.eventId] || []).map(v => {
        const f = fieldById[v.fieldId];
        return f ? {
          fieldId:    v.fieldId,
          fieldName:  f['顯示名稱'],
          fieldType:  f['欄位類型'],
          value:      v['值']
        } : null;
      }).filter(Boolean);

      return {
        eventId:      e.eventId,
        typeId:       e.typeId,
        typeName:     type ? type['名稱'] : '(已刪除類型)',
        typeFullName: type ? (rootType && rootType.typeId !== type.typeId
          ? `${rootType['名稱']} - ${type['名稱']}` : type['名稱']) : '',
        typeIcon:     (type && type.icon) || (rootType && rootType.icon) || '',
        typeColor:    (type && type.color) || (rootType && rootType.color) || '#667eea',
        date:         _cal_dateStr(e['日期']),
        title:        e['顯示標題'] || '',
        createdAt:    e['建立時間'],
        updatedAt:    e['最後更新時間'],
        values:       evValues
      };
    })
    .sort((a, b) => (a.date < b.date ? -1 : 1));

  return result;
}

// 取單一事項詳情（給編輯 modal 用）
function cal_getEvent(data) {
  if (!data || !data.eventId) throw new Error('eventId 必填');
  const all = cal_getEvents({});
  const found = all.find(e => e.eventId === data.eventId);
  if (!found) throw new Error('找不到事項');
  return found;
}

// 新增事項
function cal_addEvent(data) {
  if (!data) throw new Error('資料必填');
  if (!data.typeId) throw new Error('typeId 必填');
  if (!data.date)   throw new Error('日期必填');

  // 驗證類型存在
  const types = _cal_readAll(_CAL_SHEET.TYPES);
  const type = types.find(t => t.typeId === data.typeId);
  if (!type) throw new Error('查無類型：' + data.typeId);

  const eventId = Utilities.getUuid();
  const now = new Date();
  const title = (data.title || '').toString().trim()
    || _cal_autoTitle(data.values, data.typeId, types)
    || '(無標題)';

  // 寫事項本表
  _cal_getSheet(_CAL_SHEET.EVENTS).appendRow([
    eventId, data.typeId, data.date, title, data.createdBy || '', now, now
  ]);

  // 寫欄位值
  _cal_writeEventValues(eventId, data.values || {});

  return { success: true, eventId: eventId, message: '已新增事項' };
}

// 更新事項
function cal_updateEvent(data) {
  if (!data || !data.eventId) throw new Error('eventId 必填');

  const sh = _cal_getSheet(_CAL_SHEET.EVENTS);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const colIdx = {};
  headers.forEach((h, i) => colIdx[h] = i);

  let row = -1;
  for (let r = 1; r < values.length; r++) {
    if (values[r][colIdx['eventId']] === data.eventId) { row = r; break; }
  }
  if (row === -1) throw new Error('找不到事項');

  // 更新主表欄位
  if (data.typeId !== undefined) values[row][colIdx['typeId']] = data.typeId;
  if (data.date !== undefined)   values[row][colIdx['日期']]   = data.date;
  if (data.title !== undefined) {
    const types = _cal_readAll(_CAL_SHEET.TYPES);
    values[row][colIdx['顯示標題']] = (data.title || '').toString().trim()
      || _cal_autoTitle(data.values, values[row][colIdx['typeId']], types)
      || '(無標題)';
  }
  values[row][colIdx['最後更新時間']] = new Date();
  sh.getRange(row + 1, 1, 1, values[row].length).setValues([values[row]]);

  // 更新欄位值（若有傳）
  if (data.values !== undefined) {
    // 先刪掉舊欄位值，再寫新的（簡單可靠）
    _cal_deleteRowsByIdSet(_CAL_SHEET.VALUES, 'eventId', new Set([data.eventId]));
    _cal_writeEventValues(data.eventId, data.values || {});
  }

  return { success: true, message: '已更新' };
}

// 刪除事項
function cal_deleteEvent(data) {
  if (!data || !data.eventId) throw new Error('eventId 必填');
  _cal_deleteRowsByIdSet(_CAL_SHEET.EVENTS, 'eventId', new Set([data.eventId]));
  _cal_deleteRowsByIdSet(_CAL_SHEET.VALUES, 'eventId', new Set([data.eventId]));
  return { success: true, message: '已刪除' };
}

// ─────────────────────────────────────────────────────────────
//  helpers
// ─────────────────────────────────────────────────────────────
function _cal_dateStr(d) {
  if (!d) return '';
  if (typeof d === 'string') return d;
  if (d instanceof Date) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }
  return String(d);
}

function _cal_writeEventValues(eventId, valuesObj) {
  // valuesObj: { fieldId: value, ... }
  const rows = [];
  Object.entries(valuesObj).forEach(([fieldId, v]) => {
    if (v === undefined || v === null || v === '') return;
    const valStr = (typeof v === 'object') ? JSON.stringify(v) : String(v);
    rows.push([eventId, fieldId, valStr]);
  });
  if (rows.length === 0) return;
  const sh = _cal_getSheet(_CAL_SHEET.VALUES);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
}

// 沒給 title 時自動產生：找該類型「第一個欄位」的值
function _cal_autoTitle(valuesObj, typeId, allTypes) {
  if (!valuesObj || Object.keys(valuesObj).length === 0) return '';
  // 找頂層
  let cur = allTypes.find(t => t.typeId === typeId);
  while (cur && cur.parentTypeId) {
    const p = allTypes.find(t => t.typeId === cur.parentTypeId);
    if (!p) break;
    cur = p;
  }
  if (!cur) return '';
  const fields = _cal_readAll(_CAL_SHEET.FIELDS)
    .filter(f => f.typeId === cur.typeId)
    .sort((a, b) => (Number(a.sortOrder) || 0) - (Number(b.sortOrder) || 0));
  for (const f of fields) {
    const v = valuesObj[f.fieldId];
    if (v) return String(v).substring(0, 60);
  }
  return '';
}
