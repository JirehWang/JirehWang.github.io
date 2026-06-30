// ==========================================
// WorshipTeamMembers.js - 敬拜團員名單 API
// ==========================================

const TEAM_MEMBER_SHEET_NAME = '敬拜團員名單';

// ─────────────────────────────────────────────────────────────
//  讀主日會友名單 → 給 datalist 自動完成用 (已由 getCachedMembers 最佳化)
// ─────────────────────────────────────────────────────────────
function worship_getMemberSuggestions() {
  // 最佳化：直接呼叫主系統已快取的 getCachedMembers()，省去跨試算表 openById 耗時
  const members = getCachedMembers();
  if (!members || members.length === 0) return [];

  const result = members.map(row => ({
    name:   String(row[0] || '').trim(),
    uid:    String(row[7] || '').trim(),
    gender: String(row[1] || '').trim()
  })).filter(m => m.name && m.uid);

  // 排序：照姓名
  result.sort((a, b) => a.name.localeCompare(b.name));
  return result;
}

// ─────────────────────────────────────────────────────────────
//  讀敬拜團員名單
// ─────────────────────────────────────────────────────────────
function getTeamMembers() {
  const sheet = getWorshipSS().getSheetByName(TEAM_MEMBER_SHEET_NAME);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(row => ({
    name:     row[0] ? String(row[0]).trim() : '',
    uid:      row[1] ? String(row[1]).trim() : '',
    status:   row[2] ? String(row[2]).trim() : '正式',
    joinDate: row[3] || ''
  })).filter(m => m.name);
}

// ─────────────────────────────────────────────────────────────
//  儲存敬拜團員名單（整份覆寫）
// ─────────────────────────────────────────────────────────────
function saveTeamMembers(members) {
  const sheet = getWorshipSS().getSheetByName(TEAM_MEMBER_SHEET_NAME);
  if (!sheet) throw new Error('找不到「' + TEAM_MEMBER_SHEET_NAME + '」工作表，請先執行 setupDatabase()');

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 4).clearContent();

  if (members && members.length > 0) {
    const rows = members
      .filter(m => m && m.name && m.name.trim())
      .map(m => [
        m.name.trim(),
        m.uid || '',
        (m.status === '實習') ? '實習' : '正式',
        m.joinDate || new Date()
      ]);
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 4).setValues(rows);
    }
  }

  return { message: '敬拜團員名單已儲存（共 ' + (members ? members.length : 0) + ' 位）' };
}
