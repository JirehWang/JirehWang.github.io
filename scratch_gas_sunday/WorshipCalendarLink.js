/**
 * WorshipCalendarLink.js — 敬拜團 ↔ 教會行事曆 串接
 *
 * 用途：
 *   1. 公佈欄要顯示行事曆的「講道資訊」（依預設子類型 + 日期覆寫）
 *   2. 公佈欄的「聚會名稱」要對應到行事曆事項標題，把細節拉過來
 *
 * 設定儲存於敬拜團 SS 的「行事曆連結設定」sheet：
 *   ├─ 列 1: KEY=defaultSermonSubTypeId, VALUE=<typeId>
 *   ├─ 列 2: KEY=__OVERRIDES__, VALUE=（無用，標題列）
 *   └─ 列 3+ : KEY=<YYYY-MM-DD>, VALUE=<typeId>   ← 日期覆寫
 *
 * 跨 SS 讀取：用 CacheService 4 小時快取避免每次都打 openById
 */

const _CAL_LINK_SHEET = '行事曆連結設定';

// ─────────────────────────────────────────────────────────────
//  確保 sheet 存在
// ─────────────────────────────────────────────────────────────
function _ensureCalLinkSheet() {
  const ss = getWorshipSS();
  let sh = ss.getSheetByName(_CAL_LINK_SHEET);
  if (!sh) {
    sh = ss.insertSheet(_CAL_LINK_SHEET);
    sh.appendRow(['KEY', 'VALUE', '說明']);
    sh.getRange(1, 1, 1, 3)
      .setBackground('#667eea').setFontColor('#ffffff')
      .setFontWeight('bold');
    sh.setFrozenRows(1);
    // ⚠️ A 欄（KEY）強制純文字格式，避免「2026-01-05」被 Sheets 自動轉 Date 物件
    //    若被轉成 Date，String(d) 會變成 "Sun Jan 05 2026 ..." 導致日期 key 識別失敗
    sh.getRange('A:A').setNumberFormat('@');
    // 種入預設值
    sh.appendRow(['defaultSermonSubTypeId', '', '預設的講道子類型 ID（從行事曆「講道資訊」下面的子類型挑一個）']);
  } else {
    // 既有 sheet：補強制純文字格式（重複設定無害）
    try { sh.getRange('A:A').setNumberFormat('@'); } catch (e) {}
  }
  return sh;
}

// 把 sheet 讀成 key → value map（自動把 Date 型別的 key 還原成 YYYY-MM-DD）
function _readCalLinkSettings() {
  const sh = _ensureCalLinkSheet();
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return {};
  const data = sh.getRange(2, 1, lastRow - 1, 2).getValues();
  const map = {};
  data.forEach(row => {
    const k = _normalizeKey(row[0]);
    if (k) map[k] = row[1] !== undefined ? String(row[1]).trim() : '';
  });
  return map;
}

// 把 cell 內可能是 Date 物件的 key 標準化成 YYYY-MM-DD 字串
function _normalizeKey(raw) {
  if (raw instanceof Date && !isNaN(raw.getTime())) {
    return Utilities.formatDate(raw, _getTz(), 'yyyy-MM-dd');
  }
  return raw ? String(raw).trim() : '';
}

// 寫單一 key（找不到則新增）
function _setCalLinkSetting(key, value, comment) {
  const sh = _ensureCalLinkSheet();
  const lastRow = sh.getLastRow();
  const targetKey = String(key).trim();
  if (lastRow > 1) {
    const keys = sh.getRange(2, 1, lastRow - 1, 1).getValues().map(r => _normalizeKey(r[0]));
    const idx = keys.indexOf(targetKey);
    if (idx !== -1) {
      // 順手把 A 欄 cell 強制純文字 + 重寫一次 key（修補可能是 Date 的舊資料）
      sh.getRange(idx + 2, 1).setNumberFormat('@').setValue(targetKey);
      sh.getRange(idx + 2, 2).setValue(value);
      if (comment) sh.getRange(idx + 2, 3).setValue(comment);
      return;
    }
  }
  // 新增列：先 append，再強制 A 欄純文字 + 重寫 key 以對抗 Sheets 自動 Date 化
  sh.appendRow([targetKey, value, comment || '']);
  const newRow = sh.getLastRow();
  sh.getRange(newRow, 1).setNumberFormat('@').setValue(targetKey);
}

// 刪一個 key
function _deleteCalLinkSetting(key) {
  const sh = _ensureCalLinkSheet();
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return;
  const targetKey = String(key).trim();
  const keys = sh.getRange(2, 1, lastRow - 1, 1).getValues().map(r => _normalizeKey(r[0]));
  // 可能因之前的 bug 有多筆重複 → 倒著刪全部
  for (let i = keys.length - 1; i >= 0; i--) {
    if (keys[i] === targetKey) sh.deleteRow(i + 2);
  }
}

// ─────────────────────────────────────────────────────────────
//  跨 SS 讀行事曆（含 4h 快取，已優化為呼叫 getCalendarSS()）
// ─────────────────────────────────────────────────────────────
function _readCalendarSheet(sheetName) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'calss_' + sheetName;
  const cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { /* fall through */ }
  }
  try {
    const calSs = getCalendarSS();
    const sh = calSs.getSheetByName(sheetName);
    if (!sh) return [];
    const values = sh.getDataRange().getValues();
    if (values.length <= 1) return [];
    const headers = values[0];
    const rows = values.slice(1).filter(r => r[0]).map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        let v = row[i];
        if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        obj[h] = v;
      });
      return obj;
    });
    try { cache.put(cacheKey, JSON.stringify(rows), 14400); } catch (e) {}
    return rows;
  } catch (e) {
    Logger.log('[_readCalendarSheet] 失敗：' + e.message);
    return [];
  }
}

// 清掉行事曆 SS 的快取（管理員手動觸發時用）
function clearCalendarLinkCache() {
  const cache = CacheService.getScriptCache();
  ['事項類型', '欄位定義', '事項', '事項欄位值'].forEach(name => {
    cache.remove('calss_' + name);
  });
  return { success: true, message: '✅ 行事曆快取已清除' };
}

// ─────────────────────────────────────────────────────────────
//  輕量 API：取得「服事表總表」已建立的所有日期
//  用途：admin 設定覆寫時，從這份清單下拉選擇日期（不用憑空輸入）
// ─────────────────────────────────────────────────────────────
function getScheduleDates() {
  const ss = getWorshipSS();
  const sheet = ss.getSheetByName('服事表總表');
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const dateIdx = headers.indexOf('日期');
  const nameIdx = headers.indexOf('聚會名稱');
  const typeIdx = headers.indexOf('聚會類別');
  const yearIdx = headers.indexOf('年度');
  const quarterIdx = headers.indexOf('季度');
  if (dateIdx === -1) return [];

  const tz = _getTz();
  const seen = new Set(); // 避免同日期重覆
  const result = [];
  const entries = [];
  for (let i = 1; i < data.length; i++) {
    let d = data[i][dateIdx];
    if (!d) continue;
    if (d instanceof Date) d = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    else d = String(d).substring(0, 10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) continue;
    if (seen.has(d)) continue;
    seen.add(d);
    
    const mName = nameIdx !== -1 ? String(data[i][nameIdx] || '').trim() : '';
    result.push({
      date: d,
      name: mName,
      type: typeIdx !== -1 ? String(data[i][typeIdx] || '').trim() : '',
      year:    yearIdx    !== -1 ? String(data[i][yearIdx]    || '').trim() : '',
      quarter: quarterIdx !== -1 ? String(data[i][quarterIdx] || '').trim() : ''
    });
    entries.push({ date: d, meetingName: mName });
  }

  // 🌟 合併行事曆資料以補齊「聚會名稱」與「聚會類別」的空白 fallback
  try {
    if (entries.length > 0 && typeof getCalendarDataForDates === 'function') {
      const calData = getCalendarDataForDates({ entries: entries });
      const cfg = getCalendarLinkConfig();
      const subTypeNameById = {};
      (cfg.sermonSubTypes || []).forEach(t => { subTypeNameById[t.typeId] = t.name; });

      result.forEach(row => {
        const d = row.date;
        const meetingName = row.name;
        const cd = calData[meetingName ? `${d}|${meetingName}` : d] || calData[d] || {};

        // 聚會名稱 fallback
        if (cd.namedEvent && cd.namedEvent.title) {
          row.name = String(cd.namedEvent.title).trim();
        }

        // 聚會類別 fallback
        if (cd.sermon && cd.sermon.typeName) {
          row.type = String(cd.sermon.typeName).trim();
        } else {
          const effSubName = subTypeNameById[cd.effectiveSermonSubTypeId];
          if (effSubName) {
            row.type = effSubName;
          } else if (!row.type) {
            row.type = '主日';
          }
        }
      });
    }
  } catch (e) {
    Logger.log('getScheduleDates 整合行事曆失敗：' + e.toString());
  }

  // 按日期降冪（最近的在最上面）
  result.sort((a, b) => (a.date < b.date ? 1 : -1));
  return result;
}

// ─────────────────────────────────────────────────────────────
//  對外 API：取得行事曆連結設定（給 admin UI 用）
// ─────────────────────────────────────────────────────────────
function getCalendarLinkConfig() {
  const settings = _readCalLinkSettings();
  const types = _readCalendarSheet('事項類型');

  // 找「講道資訊」頂層 + 其子類型
  const sermonRoot = types.find(t => !t.parentTypeId && t['名稱'] === '講道資訊');
  const sermonSubTypes = sermonRoot
    ? types.filter(t => t.parentTypeId === sermonRoot.typeId && t['名稱'] !== '台語')
        .sort((a, b) => (Number(a.sortOrder) || 0) - (Number(b.sortOrder) || 0))
        .map(t => ({ typeId: t.typeId, name: t['名稱'], icon: t.icon || '', color: t.color || '#5b8def' }))
    : [];

  // 列出所有頂層類型（給 future 擴充用）
  const rootTypes = types
    .filter(t => !t.parentTypeId)
    .map(t => ({ typeId: t.typeId, name: t['名稱'], icon: t.icon || '' }));

  // 從 settings map 中拆出 defaultSermonSubTypeId 與 overrides
  const defaultSermonSubTypeId = settings['defaultSermonSubTypeId'] || '';
  const overrides = {};
  Object.entries(settings).forEach(([k, v]) => {
    // 日期格式才當作 override
    if (/^\d{4}-\d{2}-\d{2}$/.test(k)) overrides[k] = v;
  });

  // 顯示驗證：default subType 有效嗎？
  const defaultIsValid = sermonSubTypes.some(t => t.typeId === defaultSermonSubTypeId);

  // 行事曆可達：有事項類型表即可，或簡易格式下事項表直接有資料
  const events4cfg = _readCalendarSheet('事項');
  return {
    sermonRootName:       sermonRoot ? sermonRoot['名稱'] : '',
    sermonSubTypes:       sermonSubTypes,
    rootTypes:            rootTypes,
    defaultSermonSubTypeId: defaultSermonSubTypeId,
    defaultIsValid:       defaultIsValid,
    overrides:            overrides,
    calendarReachable:    types.length > 0 || events4cfg.length > 0
  };
}

// 設定預設子類型
function setDefaultSermonSubType(data) {
  if (!data) throw new Error('資料必填');
  _setCalLinkSetting('defaultSermonSubTypeId', data.typeId || '',
    '預設的講道子類型 ID（從行事曆「講道資訊」下面的子類型挑一個）');
  return { success: true, message: '已更新預設講道子類型' };
}

// 設定某日期的覆寫（typeId 空 → 移除覆寫）
function setDateOverride(data) {
  if (!data || !data.date) throw new Error('date 必填');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(data.date)) throw new Error('日期格式應為 YYYY-MM-DD');
  if (!data.typeId) {
    _deleteCalLinkSetting(data.date);
    return { success: true, message: '已移除 ' + data.date + ' 的覆寫' };
  }
  _setCalLinkSetting(data.date, data.typeId, '日期覆寫');
  return { success: true, message: '已設定 ' + data.date + ' 的覆寫' };
}

// ─────────────────────────────────────────────────────────────
//  公佈欄合併用：依「日期 + 聚會名稱」組合取「該日期講道資訊 + 同名事項」
// ─────────────────────────────────────────────────────────────
function getCalendarDataForDates(data) {
  if (!data) throw new Error('data 必填');

  // 將輸入正規化為 entries 格式 [{date, meetingName}]
  let entries;
  if (Array.isArray(data.entries)) {
    entries = data.entries.map(e => ({
      date: String(e.date || '').trim(),
      meetingName: String(e.meetingName || '').trim()
    })).filter(e => e.date);
  } else if (Array.isArray(data.dates)) {
    const map = data.meetingNamesByDate || {};
    entries = data.dates.map(d => ({
      date: String(d).trim(),
      meetingName: String(map[d] || '').trim()
    })).filter(e => e.date);
  } else {
    throw new Error('entries 或 dates 至少要傳一個');
  }

  const settings = _readCalLinkSettings();
  const defaultSubId = settings['defaultSermonSubTypeId'] || '';

  // 跨 SS 讀全部資料（4h cache）
  const types = _readCalendarSheet('事項類型');
  const fields = _readCalendarSheet('欄位定義');
  const events = _readCalendarSheet('事項');
  const values = _readCalendarSheet('事項欄位值');

  // index 化
  const typeById = {};
  types.forEach(t => typeById[t.typeId] = t);
  const fieldById = {};
  fields.forEach(f => fieldById[f.fieldId] = { name: f['顯示名稱'], type: f['欄位類型'] });
  const valuesByEvent = {};
  values.forEach(v => {
    if (!valuesByEvent[v.eventId]) valuesByEvent[v.eventId] = {};
    const f = fieldById[v.fieldId];
    if (f) valuesByEvent[v.eventId][f.name] = v['值']; // 以欄位「顯示名稱」為 key
  });
  // 🔧 Schema 適配：支援簡易格式（ID | 日期 | 聚會名稱 | 聚會類別）
  events.forEach(e => {
    if (!e['顯示標題'] && e['聚會名稱']) {
      e['顯示標題'] = String(e['聚會名稱']).trim();
    }
    if (!e['typeId'] && e['聚會類別']) {
      e['_typeName'] = String(e['聚會類別']).trim();
    }
  });

  const eventsByDate = {};
  events.forEach(e => {
    const d = String(e['日期']).substring(0, 10);
    if (!eventsByDate[d]) eventsByDate[d] = [];
    eventsByDate[d].push(e);
  });

  // 找「講道資訊」頂層
  const sermonRoot = types.find(t => !t.parentTypeId && t['名稱'] === '講道資訊');
  const sermonSubTypes = sermonRoot ? types.filter(t => t.parentTypeId === sermonRoot.typeId && t['名稱'] !== '台語') : [];
  const sermonSubIds = new Set(sermonSubTypes.map(t => t.typeId));

  // 同日期 → sermon event 共用（不分聚會名稱）
  const sermonCacheByDate = {};
  function resolveSermonForDate(date) {
    if (sermonCacheByDate.hasOwnProperty(date)) return sermonCacheByDate[date];
    const dayEvents = eventsByDate[date] || [];
    const effectiveSubId = (settings[date] && settings[date].trim()) || defaultSubId;
    let sermonEvent = null;
    if (effectiveSubId) {
      sermonEvent = dayEvents.find(e => e.typeId === effectiveSubId) || null;
    }
    if (!sermonEvent && sermonSubIds.size > 0) {
      sermonEvent = dayEvents.find(e => sermonSubIds.has(e.typeId)) || null;
    }
    const obj = { effectiveSubId, sermonEvent };
    sermonCacheByDate[date] = obj;
    return obj;
  }

  const result = {};
  entries.forEach(entry => {
    const date = entry.date;
    const meetingName = entry.meetingName;
    const dayEvents = eventsByDate[date] || [];

    // 1. 講道資訊：per-date
    const { effectiveSubId, sermonEvent } = resolveSermonForDate(date);

    // 2. namedEvent：per-(date, meetingName)
    let namedEvent = null;
    if (meetingName) {
      namedEvent = dayEvents.find(e => {
        const t = (e['顯示標題'] || '').toString().trim();
        // 雙向 includes：標題包含名稱、或名稱包含標題（容錯講道資訊標題常帶括號）
        return t && (t.indexOf(meetingName) !== -1 || meetingName.indexOf(t) !== -1);
      }) || null;
    } else {
      // 🔧 如果 meetingName 為空（代表本地尚未填寫聚會名稱），則嘗試在 dayEvents 中尋找類型為「聚會名稱」的事項！
      const meetingNameType = types.find(t => t['名稱'] === '聚會名稱');
      if (meetingNameType) {
        namedEvent = dayEvents.find(e => e.typeId === meetingNameType.typeId) || null;
      }
      if (!namedEvent && dayEvents.length > 0) {
        // 排除屬於 sermonSubIds 的講道事項，避免非預期的講道 fallback
        namedEvent = dayEvents.find(e => !sermonSubIds.has(e.typeId)) || null;
      }
    }

    const key = meetingName ? `${date}|${meetingName}` : date;
    result[key] = {
      effectiveSermonSubTypeId: effectiveSubId,
      sermon: sermonEvent ? {
        eventId:   sermonEvent.eventId,
        typeId:    sermonEvent.typeId,
        typeName:  typeById[sermonEvent.typeId] ? typeById[sermonEvent.typeId]['名稱'] : '',
        title:     sermonEvent['顯示標題'],
        values:    valuesByEvent[sermonEvent.eventId] || {}
      } : null,
      namedEvent: namedEvent ? {
        eventId:   namedEvent.eventId || namedEvent['ID'] || '',
        typeId:    namedEvent.typeId || '',
        typeName:  typeById[namedEvent.typeId] ? typeById[namedEvent.typeId]['名稱'] : (namedEvent['_typeName'] || namedEvent['聚會類別'] || ''),
        title:     namedEvent['顯示標題'] || namedEvent['聚會名稱'] || '',
        values:    valuesByEvent[namedEvent.eventId] || {}
      } : null
    };
    // 同時加入只用 date 為 key 的後備
    result[date] = result[key];
  });

  return result;
}
