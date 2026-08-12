// ==========================================
// 📖 1. 讀取講員資訊 (含 UUID 與區間過濾機制)api_sermon.gs
// ==========================================
/**
 * 🌟 講道登錄專用：從外部「聚會資料」抓出日期框架，再填入已存檔的講道資訊
 */
function getSermonInfo(payload) {
  const startDate = payload.startDate;
  const endDate = payload.endDate;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // 1. 讀取已經存檔的「牧師題目經文」資料 (Local)
  const sermonSheet = ss.getSheetByName('牧師題目經文');
  let localSermonDict = {};
  
  if (sermonSheet) {
    const data = sermonSheet.getDataRange().getValues();
    if (data.length > 1) {
      const headers = data[0];
      const uuidIdx = headers.indexOf('UUID');
      const dIdx = headers.indexOf('日期');
      const typeIdx = headers.indexOf('聚會類別');
      const pastorIdx = headers.indexOf('牧師');
      const titleIdx = headers.indexOf('題目');
      const scriptureIdx = headers.indexOf('經文');
      
      data.slice(1).forEach(row => {
        const dNum = toDateNum(row[dIdx]);
        const type = row[typeIdx] ? String(row[typeIdx]).trim() : "主日";
        if (dNum > 0) {
          // 建立字典以便後續配對
          localSermonDict[`${dNum}_${type}`] = {
            'UUID': uuidIdx > -1 ? row[uuidIdx] : '',
            '牧師': pastorIdx > -1 ? row[pastorIdx] : '',
            '題目': titleIdx > -1 ? row[titleIdx] : '',
            '經文': scriptureIdx > -1 ? row[scriptureIdx] : ''
          };
        }
      });
    }
  }

  // 2. 讀取外部「聚會資料」建立基礎框架 (External)
  const extUrl = 'https://docs.google.com/spreadsheets/d/1tKI5k7HwI9S2bTRV6RuKrzxBFGzROYHytsFYHuH8H7E/edit';
  let extData = [];
  try {
    const extSs = SpreadsheetApp.openByUrl(extUrl);
    const extSheet = extSs.getSheetByName('聚會資料');
    if (extSheet) extData = extSheet.getDataRange().getValues();
  } catch (e) {
    throw new Error("無法讀取外部聚會資料，請確認網址與權限。");
  }

  if (extData.length <= 1) return [];

  const extHeaders = extData[0];
  const extDateIdx = extHeaders.indexOf('日期');
  const extNameIdx = extHeaders.indexOf('聚會名稱');
  const extTypeIdx = extHeaders.indexOf('聚會類別');

  // 若有選取區間，則設定過濾條件
  const startNum = startDate ? toDateNum(startDate) : 0;
  const endNum = endDate ? toDateNum(endDate) : 99999999;

  let result = [];
  const tz = ss.getSpreadsheetTimeZone();

  extData.slice(1).forEach(row => {
    const dNum = toDateNum(row[extDateIdx]);
    
    // 只保留符合區間的外部日期框架
    if (dNum >= startNum && dNum <= endNum) {
      let rawDate = row[extDateIdx];
      let dateStr = "";
      if (rawDate instanceof Date) {
        dateStr = Utilities.formatDate(rawDate, tz, "yyyy-MM-dd");
      } else {
        dateStr = String(rawDate).trim();
      }

      const meetingName = row[extNameIdx] ? String(row[extNameIdx]).trim() : "";
      const meetingType = row[extTypeIdx] ? String(row[extTypeIdx]).trim() : "主日";
      
      // 🌟 關鍵：用日期+類別去剛才建立的字典找找看，是不是已經有講道資料了？
      const localInfo = localSermonDict[`${dNum}_${meetingType}`] || {};

      result.push({
        'UUID': localInfo['UUID'] || '',
        '日期': dateStr,
        '聚會名稱': meetingName,
        '聚會類別': meetingType,
        '牧師': localInfo['牧師'] || '',
        '題目': localInfo['題目'] || '',
        '經文': localInfo['經文'] || ''
      });
    }
  });

  return result; // 回傳以「外部日期」為主的完整清單給前端
}
// ==========================================
// 📖 2. 儲存講員資訊 (依 UUID 精準更新與刪除 + 瘦身)
// ==========================================
function saveSermonInfo(payload) {
  const SHEET_NAME = '牧師題目經文';
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  let sermonData = payload.sermonData || [];
  let deletedUUIDs = payload.deletedUUIDs || [];

  const data = sheet.getDataRange().getValues();
  let headers = data.length > 0 ? data[0] : ['UUID', '日期', '聚會類別', '牧師', '題目', '經文'];
  
  if (headers.indexOf('UUID') === -1) headers.push('UUID');

  let dbMap = new Map();
  if (data.length > 1) {
    for (let i = 1; i < data.length; i++) {
      let row = data[i];
      let obj = {};
      headers.forEach((h, idx) => obj[h] = row[idx]);

      if (!obj['UUID']) obj['UUID'] = Utilities.getUuid();

      if (!deletedUUIDs.includes(obj['UUID'])) {
        dbMap.set(obj['UUID'], obj);
      }
    }
  }

  sermonData.forEach(item => {
    let id = item['UUID'];
    if (!id) id = Utilities.getUuid();

    dbMap.set(id, {
      'UUID': id,
      '日期': item['日期'],
      '聚會類別': item['聚會類別'] || '主日',
      '牧師': item['牧師'] || '',
      '題目': item['題目'] || '',
      '經文': item['經文'] || ''
    });
  });

  const today = new Date();
  const cutoffDate = new Date();
  cutoffDate.setFullYear(today.getFullYear() - 1);
  
  // 🌟 修正：改用 SpreadsheetTimeZone 確保刪除期限精準
  const cutoffDateStr = Utilities.formatDate(cutoffDate, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");

  let finalObjects = [];
  dbMap.forEach((obj, uuid) => {
    // 確保存入時的日期也是標準格式
    let objDateStr = obj['日期'];
    if (objDateStr instanceof Date) {
      objDateStr = Utilities.formatDate(objDateStr, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
    }

    if (objDateStr >= cutoffDateStr) {
      finalObjects.push(obj);
    }
  });

  finalObjects.sort((a, b) => {
    let dComp = String(a['日期']).localeCompare(String(b['日期']));
    if (dComp !== 0) return dComp;
    return String(a['聚會類別']).localeCompare(String(b['聚會類別']));
  });

  let finalRows = finalObjects.map(obj => headers.map(h => obj[h] || ''));

  sheet.clearContents();
  sheet.appendRow(headers);
  if (finalRows.length > 0) {
    sheet.getRange(2, 1, finalRows.length, headers.length).setValues(finalRows);
  }

  return { status: "success", message: "資料已成功同步，並透過唯一碼精準更新！" };
}