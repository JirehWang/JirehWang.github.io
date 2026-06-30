/**
 * GroupCore.js — 小組點名系統的核心
 *
 * 整合自原 小組點名_測試版/Core.js
 * 設計原則：
 *  1. 小組資料仍住在原本的小組試算表（GROUP_SHEET_ID），不搬遷
 *  2. 同一份 GAS 同時操作兩個試算表（主日 + 小組），消除跨 GAS UrlFetch
 *  3. 共用 CacheManager 的會友名單快取（getCachedMembers），不再 HTTP 取資料
 */

// ── 常數區 ───────────────────────────────────────────────────
const GROUP_SHEET_ID = '1yErbbQUXmnOGga-CyyEnC1E3sM6JrdNZYn7PkOe8jp0';
// const ADMIN_CODE     = 'LK31'; // Moved to the bottom to avoid redeclaration error
const SECRET_TOKEN   = 'ChurchApp-2026'; // 與前端 _AUTH_TOKEN 一致

// ── 小組試算表快取（同一次請求內共用）──────────────────────────
let _groupSsCache = null;
function getGroupSS() {
  if (!_groupSsCache) _groupSsCache = SpreadsheetApp.openById(GROUP_SHEET_ID);
  return _groupSsCache;
}

/**
 * 取小組試算表的 sheet（容錯版：抓不到時走全表掃描）
 */
function getGroupSheet(targetName) {
  const ss = getGroupSS();
  const name = String(targetName).trim();
  let sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  const allSheets = ss.getSheets();
  for (let i = 0; i < allSheets.length; i++) {
    if (String(allSheets[i].getName()).trim() === name) return allSheets[i];
  }
  return null;
}

// ── 小組清單 CacheService 快取 ────────────────────────────────
const GROUPS_CACHE_KEY = 'GROUPS_LIST_V1';
const GROUPS_CACHE_TTL = 19800; // 5.5 小時；寫入即時重建，keepWarm 每 4 小時兜底

function _ensureGroupListSchema(sheet) {
  const expected = ["名稱", "狀態", "代碼", "建立日期", "UUID", "類型", "關聯常設小組"];
  const lastCol = sheet.getLastColumn();
  let cur = [];
  if (lastCol > 0) {
    cur = sheet.getRange(1, 1, 1, Math.min(lastCol, expected.length)).getValues()[0];
  }
  expected.forEach((h, i) => {
    if (!cur[i] || String(cur[i]).trim() !== h) {
      sheet.getRange(1, i + 1).setValue(h);
    }
  });
}

function _readGroupsFromSheet() {
  const sheet = getGroupSheet('小組清單');
  if (!sheet) return [];
  _ensureGroupListSchema(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(2, 1, lastRow - 1, Math.max(7, lastCol)).getValues();
  const headers = sheet.getRange(1, 1, 1, Math.max(7, lastCol)).getValues()[0];
  const distIdx = headers.indexOf('district');
  const clustIdx = headers.indexOf('cluster');
  return data
    .filter(row => row[0] && row[1] !== '隱藏')
    .map(row => ({
      name: String(row[0]).trim(),
      status: String(row[1] || '顯示').trim(),
      code: String(row[2] || '').trim(),
      date: row[3],
      uuid: String(row[4] || '').trim(),
      type: String(row[5] || '一般小組').trim(),
      associatedGroup: String(row[6] || '').trim(),
      district: distIdx !== -1 ? String(row[distIdx] || '').trim() : '',
      cluster: clustIdx !== -1 ? String(row[clustIdx] || '').trim() : ''
    }));
}

function _rebuildGroupsCache() {
  const groups = _readGroupsFromSheet();
  CacheService.getScriptCache().put(GROUPS_CACHE_KEY, JSON.stringify(groups), GROUPS_CACHE_TTL);
  return groups;
}

function refreshGroupCaches() {
  try {
    CacheService.getScriptCache().remove(GROUPS_CACHE_KEY);
    const groups = _rebuildGroupsCache();
    firebaseInvalidate(['getGroups', 'getAdminGroupsList', 'ministry_getGroups']);
    return { success: true, groups: groups, message: 'Group caches refreshed' };
  } catch (e) {
    return { success: false, message: 'Group cache refresh failed: ' + e.message };
  }
}

// ── 統一 JSON 回應 ─────────────────────────────────────────────
function _groupResponseJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Keep-warm：每 4 小時無條件重建三個快取 ───────────────────
/**
 * 由 time-based trigger 呼叫（每 4 小時）
 * 無條件重建 3 個 cache（會友名單 / 小組清單 / 點名系統清單）
 *
 * 設計考量：
 *   - Cache TTL 設 5.5 小時（接近 CacheService 6 小時上限）
 *   - 每 4 小時重建，留 1.5 小時餘裕，確保快取永不掉
 *   - 名單異動時另由 invalidateAndRebuildMemberCache() 立即同步
 *   - 一天僅執行 6 次，相對舊版「每 5 分鐘 + 每 10 分鐘」雙 trigger 大幅省 quota
 */
function keepWarm() {
  Logger.log('[keepWarm] ' + new Date().toISOString());

  try { _rebuildMemberCache(); Logger.log('[keepWarm] member cache rebuilt'); }
  catch (e) { Logger.log('[keepWarm] member cache rebuild failed: ' + e.message); }

  try { _rebuildGroupsCache(); Logger.log('[keepWarm] groups cache rebuilt'); }
  catch (e) { Logger.log('[keepWarm] groups cache rebuild failed: ' + e.message); }

  try { _rebuildGroupConfigCache(); Logger.log('[keepWarm] group config cache rebuilt'); }
  catch (e) { Logger.log('[keepWarm] group config cache rebuild failed: ' + e.message); }
}

/**
 * 手動執行一次，建立（或重設）4 小時 keep-warm trigger
 * 一併清除舊版 syncMemberCacheFromSheet trigger（已過時）
 * 在 GAS 編輯器選擇此函式後按 ▶ 執行即可
 */
function setupKeepWarmTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'keepWarm' || fn === 'syncMemberCacheFromSheet') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('keepWarm')
    .timeBased()
    .everyHours(4)
    .create();
  Logger.log('✅ keepWarm trigger 已建立（每 4 小時，並清除舊版 syncMemberCacheFromSheet trigger）');
}

// ═══════════════════════════════════════════════════════════════
//  小組管理 API
// ═══════════════════════════════════════════════════════════════

function getGroups() {
  const raw = CacheService.getScriptCache().get(GROUPS_CACHE_KEY);
  if (raw !== null) {
    try { return { success: true, groups: JSON.parse(raw) }; } catch (e) { /* 損毀，重建 */ }
  }
  const sheet = getGroupSheet('小組清單');
  if (!sheet) return { success: false, message: "找不到 '小組清單' 分頁" };
  return { success: true, groups: _rebuildGroupsCache() };
}

function verifyGroup(groupName, groupCode) {
  const decryptedCode = decryptGroupCode(groupCode).trim().toUpperCase();
  if (decryptedCode === ADMIN_CODE) {
    return { success: true, message: '管理員授權', isAdmin: true, encryptedCode: encryptGroupCode(ADMIN_CODE) };
  }
  const res = findGroupByCode(groupCode);
  if (res.success && res.groupName === String(groupName).trim()) {
    return { success: true, message: '驗證成功', encryptedCode: encryptGroupCode(groupCode) };
  }
  if (res.success && res.isAdmin) {
    return { success: true, message: '管理員驗證成功', encryptedCode: encryptGroupCode(ADMIN_CODE) };
  }
  return { success: false, message: res.message || '驗證失敗' };
}

function findGroupByCode(groupCode) {
  const decryptedCode = decryptGroupCode(groupCode).trim().toUpperCase();
  if (decryptedCode === ADMIN_CODE) {
    return { success: true, groupName: "ADMIN", isAdmin: true, encryptedCode: encryptGroupCode(ADMIN_CODE) };
  }

  if (!decryptedCode) return { success: false, message: '請輸入代碼' };

  const sheet = getGroupSheet("小組清單");
  if (!sheet) return { success: false, message: "找不到小組清單分頁" };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowCode = String(data[i][2] || "").trim(); // 原始大小寫
    const rowCodeUpper = rowCode.toUpperCase();
    
    if (rowCodeUpper === decryptedCode) {
      return {
        success: true,
        groupName: data[i][0] ? String(data[i][0]).trim() : "",
        isAdmin: false,
        encryptedCode: encryptGroupCode(rowCode) // 🌟 加密保留原始大小寫的 rowCode
      };
    }
  }
  return { success: false, message: '查無此代碼，請檢查大小寫或空格' };
}

function createGroup(groupName, groupCode, groupType, associatedGroup) {
  try {
    const sheet = getGroupSheet("小組清單");
    _ensureGroupListSchema(sheet);
    const dateStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd");
    const uuid = Utilities.getUuid();
    const type = groupType ? String(groupType).trim() : "一般小組";
    const assoc = associatedGroup ? String(associatedGroup).trim() : "";
    sheet.appendRow([String(groupName).trim(), '顯示', String(groupCode).trim(), dateStr, uuid, type, assoc]);
    
    // 如果是幸福小組，且選擇了繼承小組，則立刻自動初始化同工名單與紀錄表
    if (type === "幸福小組" && assoc) {
      try {
        const allMembers = getCachedMembers();
        const inheritedMembers = allMembers
          .filter(m => memberInGroup(m[8], assoc))
          .map(m => ({
            name: m[0] ? String(m[0]).trim() : "",
            role: "同工" // 預設為同工
          }))
          .filter(m => m.name);
        
        if (inheritedMembers.length > 0) {
          initGroup(groupName, inheritedMembers);
        }
      } catch (ex) {
        Logger.log("自動初始化繼承同工名單失敗: " + ex.message);
      }
    }

    _rebuildGroupsCache();
    firebaseInvalidate(['getGroups', 'getAdminGroupsList', 'ministry_getGroups']);
    return { success: true, message: '小組創建成功！', groupUuid: uuid };
  } catch (e) {
    return { success: false, message: "創建失敗: " + e.message };
  }
}

function getAdminGroupsList(authCode) {
  try {
    const cleanCode = String(authCode).trim();
    const isAdmin = (cleanCode === ADMIN_CODE);

    const sheet = getGroupSheet("小組清單");
    if (!sheet) return { success: false, message: "找不到小組清單" };
    _ensureGroupListSchema(sheet);

    const data = sheet.getDataRange().getValues();
    let groups = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const name = row[0] ? String(row[0]).trim() : "";
      if (!name) continue;

      const status = row[1] ? String(row[1]).trim() : "";
      const code   = row[2] ? String(row[2]).trim() : "";
      const date   = row[3] ? row[3] : "";
      let uuid     = row[4] ? String(row[4]).trim() : "";

      // 自動補全 UUID
      if (!uuid) {
        uuid = Utilities.getUuid();
        sheet.getRange(i + 1, 5).setValue(uuid);
      }

      const type   = row[5] ? String(row[5]).trim() : "一般小組";
      const associatedGroup = row[6] ? String(row[6]).trim() : "";

      groups.push({ name, status, code, date, uuid, type, associatedGroup });
    }

    if (isAdmin) {
      return { success: true, groups: groups, isAdmin: true };
    } else {
      groups = groups.filter(g => g.code === cleanCode);
      if (groups.length === 0) {
        return { success: false, message: "權限不足或輸入代碼錯誤" };
      }
      groups = groups.map(g => ({ name: g.name, status: g.status, code: g.code, uuid: g.uuid, type: g.type, associatedGroup: g.associatedGroup }));
      return { success: true, groups: groups, isAdmin: false };
    }
  } catch (e) {
    return { success: false, message: "清單讀取發生錯誤：" + e.message };
  }
}

/**
 * 取得所有會友的「姓名 + UID」列表（給小組系統 datalist 用）
 * 任何登入小組的人都可呼叫（用 token 驗證即可，不需 admin）
 *
 * 回傳每位會友：{ name, uid }
 * 用途：在「管理名單與身分」modal 的新增輸入框做自動完成下拉
 */
function getMemberSuggestions() {
  const members = getCachedMembers();
  return {
    success: true,
    data: members
      .map(m => ({
        name: m[0] ? String(m[0]).trim() : "",
        uid:  m[7] ? String(m[7]).trim() : ""
      }))
      .filter(m => m.name && m.uid)
      .sort((a, b) => a.name.localeCompare(b.name))
  };
}

/**
 * 取得「所有有屬小組的會友」清單（管理員專用）
 *
 * 回傳每位會友：姓名 / 性別 / 系統編號 / 所屬小組 / 身分
 * 只列出 所屬小組 非空的會友（沒有歸組的不顯示）
 *
 * 權限驗證：authCode 必須是 ADMIN_CODE
 */
function getAllGroupMembers(authCode) {
  const cleanCode = String(authCode || "").trim();
  if (cleanCode !== ADMIN_CODE) {
    return { success: false, message: "無權限存取總成員清單" };
  }

  const allMembers = getCachedMembers();
  const data = allMembers
    .filter(m => m[8] && String(m[8]).trim())   // 只列出有所屬小組的
    .map(m => ({
      name:   m[0] ? String(m[0]).trim() : "",
      gender: m[1] ? String(m[1]).trim() : "",
      uid:    m[7] ? String(m[7]).trim() : "",
      group:  m[8] ? String(m[8]).trim() : "",
      role:   m[9] ? String(m[9]).trim() : "小羊"
    }))
    .filter(m => m.name);

  return { success: true, data: data };
}

function updateGroupInfo(uuid, oldName, newName, newCode, newStatus) {
  var lock = LockService.getScriptLock();
  var hasLock = false;
  try {
    lock.waitLock(10000);
    hasLock = true;

    if (!uuid) return { success: false, message: "缺少小組系統識別碼 (UUID)" };

    const listSheet = getGroupSheet("小組清單");
    if (!listSheet) return { success: false, message: "找不到小組清單" };

    const data = listSheet.getDataRange().getValues();
    let targetRowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      if (data[i][4] && String(data[i][4]).trim() === String(uuid).trim()) {
        targetRowIndex = i + 1;
        break;
      }
    }

    if (targetRowIndex === -1) return { success: false, message: "系統錯誤：查無此小組的系統識別碼" };

    const cleanNewName = String(newName).trim();
    const cleanOldName = String(oldName).trim();

    if (cleanOldName !== cleanNewName) {
      for (let i = 1; i < data.length; i++) {
        if (i + 1 !== targetRowIndex && data[i][0] && String(data[i][0]).trim() === cleanNewName) {
          return { success: false, message: "新名稱已與其他小組重複，請換一個名字！" };
        }
      }
    }

    const finalStatus = (newStatus !== undefined && newStatus !== null)
      ? String(newStatus).trim()
      : data[targetRowIndex - 1][1];
    listSheet.getRange(targetRowIndex, 1, 1, 3).setValues([[
      cleanNewName,
      finalStatus,
      String(newCode).trim()
    ]]);

    if (cleanOldName !== cleanNewName) {
      const nameSheet = getGroupSheet(cleanOldName + "_名單");
      if (nameSheet) nameSheet.setName(cleanNewName + "_名單");

      const recordSheet = getGroupSheet(cleanOldName + "_點名紀錄");
      if (recordSheet) recordSheet.setName(cleanNewName + "_點名紀錄");
    }

    _rebuildGroupsCache();
    firebaseInvalidate(['getGroups', 'getAdminGroupsList', 'ministry_getGroups']);
    return { success: true };
  } catch (e) {
    if (!hasLock) {
      return { success: false, message: "伺服器繁忙，請稍後再試..." };
    }
    return { success: false, message: "後端執行錯誤：" + e.message };
  } finally {
    if (hasLock) {
      lock.releaseLock();
    }
  }
}

const OBFUSCATION_KEY = "LKC-Secure-2026";
const ENC_PREFIX = "enc_";
const ADMIN_CODE = "LK31";

function decryptGroupCode(str) {
  const safeStr = String(str || "");
  if (!safeStr) return "";
  if (safeStr.indexOf(ENC_PREFIX) !== 0) return safeStr;
  try {
    var hex = safeStr.substring(ENC_PREFIX.length);
    var plainText = "";
    for (var i = 0; i < hex.length; i += 2) {
      var charCode = parseInt(hex.substring(i, i + 2), 16);
      var decCharCode = charCode ^ OBFUSCATION_KEY.charCodeAt((i / 2) % OBFUSCATION_KEY.length);
      plainText += String.fromCharCode(decCharCode);
    }
    return plainText;
  } catch (e) {
    return safeStr;
  }
}

function encryptGroupCode(code) {
  // 強制轉換為字串，防止非字串類型（例如純數字）傳入時引發 Crash
  const safeCode = String(code || "");
  if (!safeCode) return "";
  if (safeCode.indexOf(ENC_PREFIX) === 0) return safeCode;
  
  var cipherText = "";
  for (var i = 0; i < safeCode.length; i++) {
    var charCode = safeCode.charCodeAt(i) ^ OBFUSCATION_KEY.charCodeAt(i % OBFUSCATION_KEY.length);
    cipherText += charCode.toString(16).padStart(2, '0');
  }
  return ENC_PREFIX + cipherText;
}
