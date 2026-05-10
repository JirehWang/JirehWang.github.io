// Google Apps Script — 週報草稿雲端同步後端
// 部署為 Web App：執行身分=我，存取權=所有人
//
// 部署後將 Web App 網址填入 js/config.js 的 GAS_SYNC_URL

const SHEET_ID   = 'https://docs.google.com/spreadsheets/d/1fdImvQ2vMS9N95DGem602EUUYzUbRoZKOdbAii2QHIs/edit?usp=sharing';          // ← 填入你的 Google Spreadsheet ID
const SHEET_NAME = '週報草稿';

// ----------------------------------------------------------
// 入口：GET（讀取）
// ?action=list             → 列出所有草稿
// ?action=load&key=XXX     → 載入指定草稿
// ----------------------------------------------------------
function doGet(e) {
  const action = (e.parameter.action || '').trim();

  if (action === 'list') {
    const rows = _getSheet().getDataRange().getValues().slice(1);
    const drafts = rows
      .filter(r => r[0])
      .map(r => ({ key: r[0], updatedAt: r[1], preview: r[2] || '' }));
    return _json({ success: true, drafts });
  }

  if (action === 'load') {
    const key = e.parameter.key || '';
    const row = _findRow(key);
    if (!row) return _json({ success: false, error: '草稿不存在' });
    return _json({ success: true, data: JSON.parse(row[3]) });
  }

  return _json({ success: false, error: '未知操作' });
}

// ----------------------------------------------------------
// 入口：POST（寫入 / 刪除）
// body JSON: { key, data }           → 儲存草稿
// body JSON: { action:'delete', key } → 刪除草稿
// ----------------------------------------------------------
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    if (body.action === 'delete') {
      _deleteRow(body.key);
      return _json({ success: true });
    }

    // 儲存
    const { key, data } = body;
    if (!key || !data) return _json({ success: false, error: '缺少 key 或 data' });

    const now     = new Date().toISOString();
    const preview = (data.taiwanese && data.taiwanese.sermonTitle)
                  || (data.mandarin  && data.mandarin.sermonTitle)
                  || '';
    const jsonStr = JSON.stringify(data);
    const sheet   = _getSheet();
    const rowIdx  = _findRowIndex(key);

    if (rowIdx > 0) {
      sheet.getRange(rowIdx, 1, 1, 4).setValues([[key, now, preview, jsonStr]]);
    } else {
      sheet.appendRow([key, now, preview, jsonStr]);
    }

    return _json({ success: true, key, updatedAt: now });

  } catch (err) {
    return _json({ success: false, error: err.message });
  }
}

// ----------------------------------------------------------
// 內部工具
// ----------------------------------------------------------
function _getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['draft_key', 'updated_at', 'preview', 'data_json']);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(4, 600);
  }
  return sheet;
}

function _findRow(key) {
  const rows = _getSheet().getDataRange().getValues().slice(1);
  return rows.find(r => r[0] === key) || null;
}

function _findRowIndex(key) {
  const rows = _getSheet().getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === key) return i + 1; // 1-based
  }
  return -1;
}

function _deleteRow(key) {
  const sheet = _getSheet();
  const rows  = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (rows[i][0] === key) { sheet.deleteRow(i + 1); break; }
  }
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}