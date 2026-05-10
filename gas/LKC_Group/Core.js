/**
 * core.gs - 系統核心路由與基礎管理 (安全加固 & 中央路由版)
 */
const SHEET_ID = '1opAz6PFYveCF4oP9d4c_Dppd2SWBgE_d4zk0a0dRIoU';
const SECRET_TOKEN = 'ChurchApp-2026'; // 🌟 必須與 config.js 的內容一致
const ADMIN_CODE = 'LK31'; // 定義最高權限密碼

// --- 快取機制 ---
let _ssCache = null;
function getSs() {
  if (!_ssCache) _ssCache = SpreadsheetApp.openById(SHEET_ID);
  return _ssCache;
}

/**
 * 處理 POST 請求 (API 主要入口)
 */
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    const token = params.token; 
    const data = params.data || {};
    
    // 🛡️ 安全檢查：驗證 Token
    if (token !== SECRET_TOKEN) {
      return responseJSON({ success: false, message: "Unauthorized: 金鑰驗證失敗" });
    }

    let result = {};

    // 🚀 中央分流邏輯
    switch (action) {
      case 'getGroups': result = getGroups(); break;
      case 'verifyGroup': result = verifyGroup(data.groupName, data.groupCode); break;
      case 'createGroup': result = createGroup(data.groupName, data.groupCode); break;
      case 'findGroupByCode': result = findGroupByCode(data.groupCode); break;
      case 'checkGroupStatus': result = checkGroupStatus(data.groupName); break;
      case 'initGroup': result = initGroup(data.groupName, data.members); break;
      case 'submitAttendance': result = submitAttendance(data.groupName, data.date, data.present, data.absent, data.newFriends); break;
      case 'updateMemberList': result = updateMemberList(data.groupName, data.members); break;
      case 'updateAttendanceRecord': result = updateAttendanceRecord(data.groupName, data.originalDate, data.newDate, data.present, data.absent, data.newFriends); break;
      case 'deleteAttendanceRecord': result = deleteAttendanceRecord(data.groupName, data.originalDate); break;
      
      case 'getStats': result = getStats(data.groupName, data.groupCode, data.startDate, data.endDate); break;
      
      case 'getAllGroupsStats': 
        // 嚴格比對最高權限密碼 (不轉大寫)
        const checkAdminCode = String(data.groupCode || data.authCode || "").trim();
        if (checkAdminCode !== ADMIN_CODE) {
            return responseJSON({success: false, message: "無權限存取全小組統計"});
        }
        result = getAllGroupsStats(data.startDate, data.endDate); 
        break;

      // ⚙️ 後台參數管理 (支援分級過濾與 UUID)
      case 'getAdminGroupsList': 
        result = getAdminGroupsList(data.authCode || data.groupCode); 
        break;
        
      case 'updateGroupInfo': 
        result = updateGroupInfo(data.uuid, data.oldName, data.newName, data.newCode, data.newStatus); 
        break;  
      case 'getWeeklyReport':
        result = getWeeklyReport();  // ✅ 改成賦值，不要直接 return
        break;
      default: result = { success: false, message: '未知操作: ' + action };
    }
    
    return responseJSON(result);
    
  } catch (error) {
    return responseJSON({ success: false, message: "核心後端錯誤: " + error.message });
  }
}

function doGet(e) {
  return ContentService.createTextOutput("✅ API 運作中... 中央路由系統已就緒 (Status: UUID Integrated)").setMimeType(ContentService.MimeType.TEXT);
}

// --- 業務邏輯實作 ---

function getGroups() {
  const sheet = getSheetSafely("小組清單");
  if (!sheet) return { success: false, message: "找不到 '小組清單' 分頁" };
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, groups: [] };

  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const groups = data
    .filter(row => row[0] && row[1] !== '隱藏')
    .map(row => ({ name: String(row[0]).trim() }));
    
  return { success: true, groups: groups };
}

function verifyGroup(groupName, groupCode) {
  const res = findGroupByCode(groupCode);
  if (res.success && res.groupName === String(groupName).trim()) {
    return { success: true, message: '驗證成功' };
  }
  return { success: false, message: res.message || '驗證失敗' };
}

/**
 * 修正後的 findGroupByCode - 第一關驗證
 */
function findGroupByCode(groupCode) {
  const searchCode = String(groupCode || "").trim();
  if (!searchCode) return { success: false, message: '請輸入代碼' };

  // 1. 優先判定管理員
  if (searchCode === ADMIN_CODE) {
    return { success: true, groupName: "ADMIN", isAdmin: true };
  }

  // 2. 搜尋資料表
  const sheet = getSheetSafely("小組清單");
  if (!sheet) return { success: false, message: "找不到小組清單分頁" };

  // 🌟 改用 getDataRange() 確保抓到所有內容，不論有沒有隱藏列
  const data = sheet.getDataRange().getValues();
  
  // 從第 1 列 (索引 1) 開始跑，避開標題列
  for (let i = 1; i < data.length; i++) {
    // 🌟 關鍵比對：強制轉為字串並去除前後空白
    const rowCode = String(data[i][2] || "").trim(); // C 欄 (索引 2)
    
    if (rowCode === searchCode) {
      // 找到了！回傳小組名稱 (A 欄)
      return { 
        success: true, 
        groupName: String(data[i][0]).trim(), 
        isAdmin: false 
      };
    }
  }

  // 如果跑完迴圈都沒找到
  return { success: false, message: '查無此代碼，請檢查大小寫或空格' };
}

function createGroup(groupName, groupCode) {
  try {
    const sheet = getSheetSafely("小組清單");
    const dateStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd");
    const uuid = Utilities.getUuid(); // 自動生成 UUID
    
    // 寫入5個欄位：名稱, 狀態, 代碼, 日期, UUID
    sheet.appendRow([String(groupName).trim(), '顯示', String(groupCode).trim(), dateStr, uuid]);
    return { success: true, message: '小組創建成功！' };
  } catch (e) {
    return { success: false, message: "創建失敗: " + e.message };
  }
}

function getSheetSafely(targetName) {
  const ss = getSs();
  const name = String(targetName).trim();
  let sheet = ss.getSheetByName(name);
  if (sheet) return sheet;

  const allSheets = ss.getSheets();
  for (let i = 0; i < allSheets.length; i++) {
    if (String(allSheets[i].getName()).trim() === name) return allSheets[i];
  }
  return null;
}

function responseJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * =========================================
 * 後台管理專用 API (管理員 & 小組長分級限定)
 * =========================================
 */

function getAdminGroupsList(authCode) {
  try {
    // 保持與後台資料完全一致的嚴格比對
    const cleanCode = String(authCode).trim(); 
    const isAdmin = (cleanCode === ADMIN_CODE); 

    const sheet = getSheetSafely("小組清單");
    if (!sheet) return { success: false, message: "找不到小組清單" };
    
    const data = sheet.getDataRange().getValues();
    let groups = [];
    
    // 從第 2 列開始讀取 (避開標題列)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const name = row[0] ? String(row[0]).trim() : "";
      
      if (!name) continue; // 略過空白列
      
      const status = row[1] ? String(row[1]).trim() : "";
      const code = row[2] ? String(row[2]).trim() : "";
      const date = row[3] ? row[3] : "";
      let uuid = row[4] ? String(row[4]).trim() : "";

      // 🌟 自動補全 UUID 機制：如果發現這組沒有 UUID，自動生成並補寫回 Sheet
      if (!uuid) {
        uuid = Utilities.getUuid();
        sheet.getRange(i + 1, 5).setValue(uuid); // 寫入第 5 欄 (E欄)
      }

      groups.push({
        name: name,
        status: status,
        code: code,
        date: date,
        uuid: uuid
      });
    }

    // 💡 權限分流判斷（修復版）
    if (isAdmin) {
        // ✅ 最高權限：回傳所有小組完整資訊
        return { success: true, groups: groups, isAdmin: true };
    } else {
        // 🔒 非最高權限：嚴格比對小組編號
        groups = groups.filter(g => g.code === cleanCode); 
        if (groups.length === 0) {
            return { success: false, message: "權限不足或輸入代碼錯誤" };
        }
        // 依照需求：只提供小組名稱、狀態、小組編號(隱藏建立日期、UUID等其餘資訊)
        groups = groups.map(g => ({ name: g.name, status: g.status, code: g.code, uuid: g.uuid }));
        return { success: true, groups: groups, isAdmin: false };
    }
  } catch (e) {
    return { success: false, message: "清單讀取發生錯誤：" + e.message };
  }
}

function updateGroupInfo(uuid, oldName, newName, newCode, newStatus) {
  try {
    if (!uuid) return { success: false, message: "缺少小組系統識別碼 (UUID)" };

    const listSheet = getSheetSafely("小組清單");
    if (!listSheet) return { success: false, message: "找不到小組清單" };

    const data = listSheet.getDataRange().getValues();
    let targetRowIndex = -1;

    // 🚀 使用 UUID 尋找目標列，完全不怕名字或代碼被改掉
    for (let i = 1; i < data.length; i++) {
      if (data[i][4] && String(data[i][4]).trim() === String(uuid).trim()) {
        targetRowIndex = i + 1; // 加上標題列的偏移
        break;
      }
    }

    if (targetRowIndex === -1) return { success: false, message: "系統錯誤：查無此小組的系統識別碼" };

    const cleanNewName = String(newName).trim();
    const cleanOldName = String(oldName).trim(); // 用於改分頁名稱
    
    // 檢查新名稱是否與「其他」小組重複
    if (cleanOldName !== cleanNewName) {
       for (let i = 1; i < data.length; i++) {
          if (i + 1 !== targetRowIndex && data[i][0] && String(data[i][0]).trim() === cleanNewName) {
             return { success: false, message: "新名稱已與其他小組重複，請換一個名字！" };
          }
       }
    }

    // ✏️ 更新資料表 
    listSheet.getRange(targetRowIndex, 1).setValue(cleanNewName); // Col 1: 名稱
    if (newStatus !== undefined && newStatus !== null) {
      listSheet.getRange(targetRowIndex, 2).setValue(String(newStatus).trim()); // Col 2: 狀態
    }
    listSheet.getRange(targetRowIndex, 3).setValue(String(newCode).trim()); // Col 3: 編號 (嚴格區分大小寫)

    // 🔄 同步修改相關的分頁名稱
    if (cleanOldName !== cleanNewName) {
       const nameSheet = getSheetSafely(cleanOldName + "_名單");
       if (nameSheet) nameSheet.setName(cleanNewName + "_名單");

       const recordSheet = getSheetSafely(cleanOldName + "_點名紀錄");
       if (recordSheet) recordSheet.setName(cleanNewName + "_點名紀錄");
    }

    return { success: true };
  } catch(e) {
    return { success: false, message: "後端執行錯誤：" + e.message };
  }
}