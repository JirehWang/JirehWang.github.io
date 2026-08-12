// 🌟 1. 設定安全金鑰 (必須與 config.js 裡面的 atob 解碼後一致)
const SECRET_TOKEN = "ChurchApp-2026"; 

/**
 * 🚀 入口導航：處理前端所有的 POST 請求
 */
function doPost(e) {
  let response; 
  
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error("無效的請求資料");

    const params = JSON.parse(e.postData.contents);
    
    // 🔒 核心安全攔截：檢查 Token
    if (params.token !== SECRET_TOKEN) {
      Logger.log("⚠️ 偵測到非法存取請求，已攔截。");
      return createJsonResponse({ status: 'error', message: '未經授權的存取：憑證無效' });
    }

    const action = params.action;
    Logger.log(`🔒 安全驗證通過，執行 Action: ${action}`);
    
    switch(action) {
      case 'getSchedule':
        response = { status: 'success', data: getMergedSchedule(params.data.year, params.data.quarter) };
        break;

      case 'getScheduleByDateRange':
        response = getScheduleByDateRange(params.data); 
        break;

      case 'saveSchedule':
        response = saveScheduleData(params.data.scheduleData);
        break;

      case 'getPositions':
        response = { status: 'success', data: getPositions() };
        break;

      case 'savePositions':
        response = { status: 'success', data: savePositions(params.data.positionsData) };
        break;

      // 🎵 新增：敬拜曲目
      case 'getSongs':
        response = getSongs(params.data);
        break;

      case 'saveSongs':
        response = saveSongs(params.data);
        break;

      default:
        response = { status: 'error', message: '未知的指令: ' + action };
    }
  } catch (error) {
    Logger.log(`❌ doPost 錯誤: ${error.toString()}`);
    response = { status: 'error', message: error.toString() };
  }
  
  return createJsonResponse(response);
}

/**
 * 🌐 GET 請求 (用於佈告欄讀取資料)
 */
function doGet(e) {
  const token = e.parameter.token;
  if (token !== SECRET_TOKEN) {
    return ContentService.createTextOutput("Unauthorized: Invalid Token")
      .setMimeType(ContentService.MimeType.TEXT);
  }

  const action = e.parameter.action;
  if (action === "getAggregatedReport") {
     // 彙整表邏輯（如有需要）
  }

  return ContentService.createTextOutput("API 運作中... 安全防護已啟動")
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * 🛠️ 輔助工具：統一輸出 JSON 格式
 */
function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}