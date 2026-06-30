/**
 * 教會行事曆 - 新版資料結構（Phase 1）
 *
 * 4 張 sheet：
 *   1. 事項類型 (EventTypes)
 *   2. 欄位定義 (EventFields)
 *   3. 事項 (Events)
 *   4. 事項欄位值 (EventValues)
 *
 * 結構特性：
 *   - 類型支援階層（parentTypeId）：頂層 + 子類型
 *   - 欄位定義在「頂層類型」上，子類型繼承
 *   - 事項欄位值用長格式儲存（long-format），方便動態擴充
 *   - 每個頂層類型可設密碼（保護公開頁查看）
 */

// ─────────────────────────────────────────────────────────────
//  Sheet 名稱（不要改，前端會用到）
// ─────────────────────────────────────────────────────────────
const _CAL_SHEET = {
  TYPES:  '事項類型',
  FIELDS: '欄位定義',
  EVENTS: '事項',
  VALUES: '事項欄位值'
};

const _CAL_HEADERS = {
  [_CAL_SHEET.TYPES]: [
    'typeId', 'parentTypeId', '名稱', 'icon', 'color', 'sortOrder',
    'syncToAttendance', 'syncToMinistry', 'syncToWorship',
    'password', 'hidden', 'createdAt', 'excludedFieldIds'
  ],
  [_CAL_SHEET.FIELDS]: [
    'fieldId', 'typeId', '顯示名稱', '欄位類型', '是否必填',
    '下拉選項', 'sortOrder', 'createdAt'
  ],
  [_CAL_SHEET.EVENTS]: [
    'eventId', 'typeId', '日期', '顯示標題', '建立者', '建立時間', '最後更新時間'
  ],
  [_CAL_SHEET.VALUES]: [
    'eventId', 'fieldId', '值'
  ]
};

// 欄位類型列舉（給前端參考）
const _CAL_FIELD_TYPES = ['text', 'longtext', 'date', 'time', 'select', 'multiselect', 'number', 'url'];

// ─────────────────────────────────────────────────────────────
//  建立 schema（重複呼叫安全）
// ─────────────────────────────────────────────────────────────
function cal_setupSchema() {
  const ss = getCalendarSS();
  Object.entries(_CAL_HEADERS).forEach(([name, headers]) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(headers);
      sh.getRange(1, 1, 1, headers.length)
        .setBackground('#667eea').setFontColor('#ffffff')
        .setFontWeight('bold').setHorizontalAlignment('center');
      sh.setFrozenRows(1);
      sh.autoResizeColumns(1, headers.length);
    } else {
      // 已存在 → 檢查是否缺欄位（往後加新欄位的 migration）
      _cal_ensureHeaders(sh, headers);
    }
  });

  // 若類型表為空，種入預設 4 個頂層類型 + 講道資訊的 3 個子類型 + 主日學 A/B 班
  const typeSheet = ss.getSheetByName(_CAL_SHEET.TYPES);
  if (typeSheet.getLastRow() <= 1) {
    _cal_seedDefaultTypes();
  }
  return { success: true, message: 'Schema 已建立完成' };
}

// 確保 sheet 的 header 含預期欄位（缺則在最右側補上）
function _cal_ensureHeaders(sh, expectedHeaders) {
  const lastCol = sh.getLastColumn();
  const existing = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(s => String(s).trim());
  const toAdd = expectedHeaders.filter(h => existing.indexOf(h) === -1);
  if (toAdd.length === 0) return;
  toAdd.forEach((h, i) => {
    sh.getRange(1, lastCol + 1 + i).setValue(h)
      .setBackground('#667eea').setFontColor('#ffffff')
      .setFontWeight('bold').setHorizontalAlignment('center');
  });
}

function _cal_seedDefaultTypes() {
  // 頂層
  const sermonId = _cal_addTypeRow({
    parentTypeId: '', name: '講道資訊', icon: '📖', color: '#5b8def', sortOrder: 1,
    syncToAttendance: true, syncToMinistry: true, syncToWorship: true
  });
  const ssId = _cal_addTypeRow({
    parentTypeId: '', name: '主日學', icon: '🎒', color: '#48bb78', sortOrder: 2,
    syncToAttendance: true, syncToMinistry: false, syncToWorship: false
  });
  _cal_addTypeRow({
    parentTypeId: '', name: '其他課程', icon: '📚', color: '#ed8936', sortOrder: 3,
    syncToAttendance: false, syncToMinistry: false, syncToWorship: false
  });
  _cal_addTypeRow({
    parentTypeId: '', name: '會議', icon: '💼', color: '#718096', sortOrder: 4,
    syncToAttendance: false, syncToMinistry: false, syncToWorship: false
  });

  // 子層 - 講道資訊（台/華/聯合-台語/聯合-華語）
  _cal_addTypeRow({ parentTypeId: sermonId, name: '台語', icon: '🌾', color: '#5b8def', sortOrder: 1 });
  _cal_addTypeRow({ parentTypeId: sermonId, name: '華語', icon: '🌏', color: '#5b8def', sortOrder: 2 });
  _cal_addTypeRow({ parentTypeId: sermonId, name: '聯合-台語', icon: '🤝', color: '#5b8def', sortOrder: 3 });
  _cal_addTypeRow({ parentTypeId: sermonId, name: '聯合-華語', icon: '🤝', color: '#5b8def', sortOrder: 4 });

  // 子層 - 主日學（A/B 班）
  _cal_addTypeRow({ parentTypeId: ssId, name: 'A 班', icon: '🅰️', color: '#48bb78', sortOrder: 1 });
  _cal_addTypeRow({ parentTypeId: ssId, name: 'B 班', icon: '🅱️', color: '#48bb78', sortOrder: 2 });

  // 為「講道資訊」頂層種入預設欄位（子類型繼承）
  _cal_addFieldRow({ typeId: sermonId, name: '講題',   type: 'text',     required: true,  sortOrder: 1 });
  _cal_addFieldRow({ typeId: sermonId, name: '講員',   type: 'text',     required: true,  sortOrder: 2 });
  _cal_addFieldRow({ typeId: sermonId, name: '經文',   type: 'text',     required: false, sortOrder: 3 });
  _cal_addFieldRow({ typeId: sermonId, name: '宣召',   type: 'text',     required: false, sortOrder: 4 });
  _cal_addFieldRow({ typeId: sermonId, name: '金句',   type: 'longtext', required: false, sortOrder: 5 });
  _cal_addFieldRow({ typeId: sermonId, name: '詩歌',   type: 'longtext', required: false, sortOrder: 6 });
  _cal_addFieldRow({ typeId: sermonId, name: '備註',   type: 'longtext', required: false, sortOrder: 7 });

  // 為「主日學」頂層種入預設欄位
  _cal_addFieldRow({ typeId: ssId, name: '老師',       type: 'text',     required: false, sortOrder: 1 });
  _cal_addFieldRow({ typeId: ssId, name: '課表內容',   type: 'longtext', required: false, sortOrder: 2 });
}

// ─────────────────────────────────────────────────────────────
//  共用：取 sheet（找不到就建；找到也檢查 header 是否齊全 → 自動 migrate）
// ─────────────────────────────────────────────────────────────
function _cal_getSheet(name) {
  const ss = getCalendarSS();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    cal_setupSchema();
    sh = ss.getSheetByName(name);
  } else {
    // 快速檢查：若欄位數比預期少 → 跑 ensureHeaders 自動補
    const expected = _CAL_HEADERS[name];
    if (expected && sh.getLastColumn() < expected.length) {
      _cal_ensureHeaders(sh, expected);
    }
  }
  return sh;
}

// ─────────────────────────────────────────────────────────────
//  共用：讀整張表為 array of objects（依 header 動態映射）
// ─────────────────────────────────────────────────────────────
function _cal_readAll(sheetName) {
  const sh = _cal_getSheet(sheetName);
  if (sh.getLastRow() <= 1) return [];
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  return values.slice(1)
    .filter(row => row[0]) // 第一欄（id）必須有值
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        let v = row[i];
        if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
        obj[h] = v;
      });
      return obj;
    });
}

// ─────────────────────────────────────────────────────────────
//  內部：直接 append 一列類型 / 欄位（種子用，不檢查重複）
// ─────────────────────────────────────────────────────────────
function _cal_addTypeRow(p) {
  const id = Utilities.getUuid();
  _cal_getSheet(_CAL_SHEET.TYPES).appendRow([
    id,
    p.parentTypeId || '',
    p.name || '',
    p.icon || '',
    p.color || '#667eea',
    p.sortOrder || 0,
    p.syncToAttendance ? 'TRUE' : 'FALSE',
    p.syncToMinistry   ? 'TRUE' : 'FALSE',
    p.syncToWorship    ? 'TRUE' : 'FALSE',
    p.password || '',
    p.hidden ? 'TRUE' : 'FALSE',
    new Date(),
    p.excludedFieldIds
      ? (typeof p.excludedFieldIds === 'string' ? p.excludedFieldIds : JSON.stringify(p.excludedFieldIds))
      : '[]'
  ]);
  return id;
}

function _cal_addFieldRow(p) {
  const id = Utilities.getUuid();
  _cal_getSheet(_CAL_SHEET.FIELDS).appendRow([
    id,
    p.typeId || '',
    p.name || '',
    p.type || 'text',
    p.required ? 'TRUE' : 'FALSE',
    p.options ? (typeof p.options === 'string' ? p.options : JSON.stringify(p.options)) : '',
    p.sortOrder || 0,
    new Date()
  ]);
  return id;
}
