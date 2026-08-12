/**
 * 教會行事曆系統 - Google Apps Script 後端 (中央路由安全版)
 */

// [請在此處保留您原本的 SPREADSHEET_ID 與 GEMINI_API_KEY 常數]
const SECRET_TOKEN = 'ChurchApp-2026'; // 🌟 必須與前端 config.js 一致

// --- 快取機制：減少重複開啟試算表的耗時 ---
let _ssCache = null;
function getSs() {
  if (!_ssCache) _ssCache = SpreadsheetApp.openById(SPREADSHEET_ID);
  return _ssCache;
}

/**
 * 🛡️ 處理 POST 請求 (API 唯一入口)
 * 支援 action: load, save, ai_parse
 */
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    const token = params.token; // 接收前端傳來的 Token
    const data = params.data || {};
    
    // 安全檢查：驗證 Token
    if (token !== SECRET_TOKEN) {
      return responseJSON({ success: false, message: "Unauthorized: 安全驗證失敗" });
    }

    let result = {};

    switch (action) {
      // 1. 載入資料 (由原本的 doGet 移至此處以提升安全性)
      case 'load':
        result = { success: true, events: loadEvents() };
        break;

      // 2. 儲存資料
      case 'save':
        saveEvents(data); // 此處 data 為前端傳來的 events 陣列
        result = { success: true, message: '雲端資料已成功儲存' };
        break;
      
      // 3. 處理 AI 講道解析
      case 'ai_parse':
        const aiParseResult = callGeminiApi(data.prompt, data.rawText);
        result = { success: true, result: aiParseResult };
        break;
      
      default:
        result = { success: false, message: '後端未定義此操作: ' + action };
    }
    
    return responseJSON(result);

  } catch (error) {
    return responseJSON({ success: false, message: "後端核心錯誤: " + error.toString() });
  }
}

/**
 * 處理 GET 請求 (僅供檢查 API 狀態)
 */
function doGet(e) {
  return ContentService.createTextOutput("✅ 行事曆系統 API 運作中... 中央路由已就緒").setMimeType(ContentService.MimeType.TEXT);
}

/**
 * 核心：呼叫 Gemini API 進行講道解析 (具重試機制)
 */
function callGeminiApi(prompt, rawText) {
  if (typeof GEMINI_API_KEY === 'undefined' || !GEMINI_API_KEY) throw new Error('未設定 GEMINI_API_KEY');

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${GEMINI_API_KEY}`;
  const payload = {
    "contents": [{ "parts": [{ "text": `${prompt}\n\n待解析文字：\n${rawText}` }] }],
    "generationConfig": { "temperature": 0.1 }
  };
  const options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true };

  let attempts = 0;
  while (attempts < 3) {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    
    if (response.getResponseCode() === 200 && !json.error) {
      return json.candidates[0].content.parts[0].text.replace(/```json/g, '').replace(/```/g, '').trim();
    }
    
    // 處理 503 或 429 錯誤的重試
    if (response.getResponseCode() === 503 || response.getResponseCode() === 429) {
      attempts++;
      Utilities.sleep(2000 * attempts);
      continue;
    }
    throw new Error(`API 錯誤: ${json.error ? json.error.message : '未知錯誤'}`);
  }
}

/**
 * 核心：儲存資料到三張獨立工作表
 */
function saveEvents(events) {
  const ss = getSs();
  
  // 初始化工作表並清除舊資料 (保留標題列)
  let mainSheet = ss.getSheetByName('聚會資料') || ss.insertSheet('聚會資料');
  let ministrySheet = ss.getSheetByName('事工細項') || ss.insertSheet('事工細項');
  let sermonSheet = ss.getSheetByName('講道資訊') || ss.insertSheet('講道資訊');

  mainSheet.clear().appendRow(['ID', '日期', '聚會名稱', '聚會類別', '最後更新時間']);
  ministrySheet.clear().appendRow(['聚會ID', '細項ID', '籌備日期', '事工內容']);
  sermonSheet.clear().appendRow(['聚會ID', '講道ID', '講道類別', '講題', '講員', '經文', '宣召', '金句', '詩歌', '備註']);
  
  const now = new Date();
  
  // 遍歷所有事件寫入資料
  events.forEach(event => {
    mainSheet.appendRow([event.id, event.date, event.name, event.category, now]);
    
    if (event.ministryItems && event.ministryItems.length > 0) {
      event.ministryItems.forEach(min => {
        ministrySheet.appendRow([event.id, min.id, min.date, min.content]);
      });
    }
    
    if (event.sermons && event.sermons.length > 0) {
      event.sermons.forEach(sermon => {
        sermonSheet.appendRow([event.id, sermon.id, sermon.type, sermon.title, sermon.speaker, sermon.scripture, sermon.callToWorship, sermon.goldenVerse, sermon.hymns, sermon.description]);
      });
    }
  });
  
  formatSheet(mainSheet);
  formatSheet(ministrySheet);
  formatSheet(sermonSheet);
}

/**
 * 核心：從三張工作表加強讀取並整合
 */
function loadEvents() {
  const ss = getSs();
  const mainSheet = ss.getSheetByName('聚會資料');
  if (!mainSheet) return [];
  
  const events = [];
  const mainData = mainSheet.getDataRange().getValues();
  
  // 1. 讀取主表
  for (let i = 1; i < mainData.length; i++) {
    if (!mainData[i][0]) continue;
    events.push({
      id: mainData[i][0],
      date: formatDate(mainData[i][1]),
      name: mainData[i][2],
      category: mainData[i][3],
      ministryItems: [],
      sermons: [],
      showSub: false
    });
  }
  
  // 2. 讀取事工表
  const ministrySheet = ss.getSheetByName('事工細項');
  if (ministrySheet && ministrySheet.getLastRow() > 1) {
    const minData = ministrySheet.getDataRange().getValues();
    for (let i = 1; i < minData.length; i++) {
      if (!minData[i][0]) continue;
      const event = events.find(e => String(e.id) === String(minData[i][0]));
      if (event) {
        event.ministryItems.push({ id: minData[i][1], date: formatDate(minData[i][2]), content: minData[i][3] });
      }
    }
  }
  
  // 3. 讀取講道表
  const sermonSheet = ss.getSheetByName('講道資訊');
  if (sermonSheet && sermonSheet.getLastRow() > 1) {
    const sermonData = sermonSheet.getDataRange().getValues();
    for (let i = 1; i < sermonData.length; i++) {
      if (!sermonData[i][0]) continue;
      const event = events.find(e => String(e.id) === String(sermonData[i][0]));
      if (event) {
        event.sermons.push({
          id: sermonData[i][1], type: sermonData[i][2], title: sermonData[i][3],
          speaker: sermonData[i][4], scripture: sermonData[i][5], callToWorship: sermonData[i][6],
          goldenVerse: sermonData[i][7], hymns: sermonData[i][8], description: sermonData[i][9]
        });
      }
    }
  }
  
  return events;
}

// --- 通用工具函式 ---

function formatDate(date) {
  if (!date) return '';
  if (typeof date === 'string') return date;
  const d = new Date(date);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

function responseJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow === 0) return;
  const headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange.setBackground('#667eea').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  for (let i = 1; i <= lastCol; i++) sheet.autoResizeColumn(i);
  sheet.setFrozenRows(1);
}