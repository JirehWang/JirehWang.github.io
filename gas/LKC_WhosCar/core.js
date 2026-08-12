/**
 * 核心：接收請求
 */
function doGet(e) {
  var action = e.parameter.action;
  var result;
  try {
    // 🌟 新增：讓「極速版」前端能一次下載所有資料庫 (這是最關鍵的一段！)
    if (action === 'getAllCarData') {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('車牌資料');
      // slice(1) 是為了扣除第一行的標題列
      result = sheet.getDataRange().getValues().slice(1);
    } 
    // 以下保留您原本的所有功能，確保向下相容
    else if (action === 'getCarOptions') {
      result = getCarOptions();
    } else if (action === 'getSearchResult') {
      result = getSearchResult(e.parameter.query);
    } else if (action === 'savePlate') {
      // 伺服器端驗證：車牌與姓名不可為空
      if (!e.parameter.plate || !e.parameter.name) {
        return ContentService.createTextOutput(JSON.stringify({status: "error", message: "❌ 儲存失敗：車號與姓名為必填項目"}))
          .setMimeType(ContentService.MimeType.JSON);
      }
      result = addOrUpdatePlate(e.parameter.plate, e.parameter.name, e.parameter.tel || ""); // 電話若無則傳空字串
    }
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({error: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 模糊搜尋邏輯 (保留)
function getSearchResult(inputData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('車牌資料'); 
  var data = sheet.getDataRange().getValues();
  var cleanInput = String(inputData).toUpperCase().replace(/[^A-Z0-9]/g, "");
  var matches = [];

  for (var i = 1; i < data.length; i++) {
    var dbPlate = String(data[i][1]).toUpperCase().replace(/[^A-Z0-9]/g, "");
    if (dbPlate === cleanInput) {
      matches.push({ data: data[i], score: 1.0 });
      continue;
    }
    var score = calculateSimilarity(dbPlate, cleanInput);
    if (score >= 0.6) { 
      matches.push({ data: data[i], score: score });
    }
  }
  matches.sort((a, b) => b.score - a.score);
  return matches.slice(0, 3).map(m => m.data);
}

// 新增或更新資料 (B:車號, C:姓名, D:電話)
function addOrUpdatePlate(plate, name, tel) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('車牌資料');
  var data = sheet.getDataRange().getValues();
  var cleanPlate = String(plate).toUpperCase().replace(/[^A-Z0-9]/g, "");
  
  for (var i = 1; i < data.length; i++) {
    var dbPlate = String(data[i][1]).toUpperCase().replace(/[^A-Z0-9]/g, "");
    if (dbPlate === cleanPlate) {
      sheet.getRange(i + 1, 3).setValue(name); 
      sheet.getRange(i + 1, 4).setValue(tel);  
      return { status: "success", message: "✅ 資料已成功更新" };
    }
  }
  sheet.appendRow([new Date(), plate, name, tel]);
  return { status: "success", message: "🆕 已建立全新車牌資料" };
}

// 相似度演算法 (Levenshtein Distance)
function calculateSimilarity(s1, s2) {
  var longer = s1.length >= s2.length ? s1 : s2;
  var shorter = s1.length < s2.length ? s1 : s2;
  if (longer.length === 0) return 1.0;
  var editDist = (function(a, b) {
    var tmp; if (a.length === 0) return b.length; if (b.length === 0) return a.length;
    if (a.length > b.length) { tmp = a; a = b; b = tmp; }
    var row = Array.from({length: a.length + 1}, (_, i) => i);
    for (var i = 1; i <= b.length; i++) {
      var prev = i;
      for (var j = 1; j <= a.length; j++) {
        var val = (b[j-1] === a[j-1]) ? row[j-1] : Math.min(row[j-1] + 1, prev + 1, row[j] + 1);
        row[j-1] = prev; prev = val;
      }
      row[a.length] = prev;
    }
    return row[a.length];
  })(longer, shorter);
  return (longer.length - editDist) / longer.length;
}

// 取得車牌下拉選單選項
function getCarOptions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('車牌資料');
  var data = sheet.getDataRange().getValues();
  var options = data.slice(1).map(r => String(r[1]).trim()).filter(r => r !== "");
  return [...new Set(options)];
}

/**
 * 🌟 Hugging Face 防休眠定時鬧鐘 (首頁穿甲彈版)
 * (每4小時執行一次，修改請從觸發條件修改)
 */
function wakeUpHuggingFace() {
  const timestamp = new Date().getTime();
  
  // 🌟 改回敲擊首頁，但掛上時間戳確保穿透 CDN 快取
  const url = "https://jirehwang-ocr.hf.space/?t=" + timestamp; 
  
  const options = {
    muteHttpExceptions: true,
    headers: {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
      "Cache-Control": "no-cache",
      "Pragma": "no-cache"
    }
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    
    // 這次因為是敲首頁，只要拿到 200 就是成功把 Flask 叫起來了！
    if (code === 200) {
      Logger.log("【防休眠鬧鐘】成功穿透並喚醒！狀態碼：" + code);
    } else {
      Logger.log("【防休眠鬧鐘】遇到預期外的狀態碼：" + code);
    }
  } catch (e) {
    Logger.log("【防休眠鬧鐘】連線失敗：" + e.message);
  }
}