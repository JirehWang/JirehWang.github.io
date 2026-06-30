const WORSHIP_SPREADSHEET_ID = '1Yh5Kw1xpxFB73AXxLpNTpeTuRbFG-wthRBcNlHm0Wx8';

// --- 快取機制：單次執行內共用同一個 Spreadsheet 物件與時區字串 ---
let _worshipSsCache = null;
let _worshipTzCache = null;

function getWorshipSS() {
  if (!_worshipSsCache) _worshipSsCache = SpreadsheetApp.openById(WORSHIP_SPREADSHEET_ID);
  return _worshipSsCache;
}

function _getTz() {
  if (!_worshipTzCache) _worshipTzCache = getWorshipSS().getSpreadsheetTimeZone();
  return _worshipTzCache;
}
