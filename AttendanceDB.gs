/** * 1. 取得智慧名單 (支援回傳男女新朋友人數與編號) 
 * 邏輯：讀取今日紀錄，將男女新朋友人數分別回傳給前端
 **/
function getSmartAttendanceList(type, userId) {
  const ss = getSS();
  try {
    const memberSheet = ss.getSheetByName("會友名單");
    if (!memberSheet) throw new Error("找不到工作表：'會友名單'");

    const members = memberSheet.getDataRange().getValues();
    const headers = members[0];
    const todayStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy/M/d");
    
    // ✅ 1. 新增抓取「編號」的欄位索引
    const nameIdx = headers.indexOf("姓名");
    const genderIdx = headers.indexOf("性別");
    const dateIdx = headers.indexOf("建立日期");
    const excludeIdx = headers.indexOf("不列入統計");
    const idIdx = headers.indexOf("編號"); // <--- 新增這行

    // 抓取今日已點名名單與男女新朋友人數
    const attInfo = getTodayAttendanceInfo(ss, type, todayStr);
    const permanentSet = new Set(attInfo.names);

    // 抓取 SYNC_TEMP 暫存 (多人協作鎖定)
    const attendanceMap = getAttendanceCountMap(ss, type);
    const syncTempData = getSyncTempData(ss, type);

    let activeList = [];
    let excludedNames = [];

    members.slice(1).forEach(row => {
      const name = row[nameIdx] ? row[nameIdx].toString().trim() : "";
      if (!name) return;
      
      // ✅ 2. 把編號讀出來
      const memberId = (idIdx !== -1 && row[idIdx]) ? row[idIdx].toString().trim() : "";
      
      if (row[excludeIdx] === true || row[excludeIdx] === "TRUE") {
        excludedNames.push(name);
      } else {
        const isSubmitted = permanentSet.has(name);
        const temp = syncTempData[name] || { checked: false, operatorId: "" };
        
        activeList.push({
          id: memberId,   // <--- ✅ 3. 將編號一起打包傳給前端
          name: name,
          gender: (genderIdx !== -1) ? (row[genderIdx] || "未知") : "未知",
          createDate: (dateIdx !== -1 && row[dateIdx] instanceof Date) ? row[dateIdx].getTime() : 0,
          count: attendanceMap[name] || 0,
          isChecked: isSubmitted || temp.checked,
          isSubmitted: isSubmitted,
          operatorId: temp.operatorId
        });
      }
    });

    // 排序：出席率高 -> 新加入
    activeList.sort((a, b) => (b.count - a.count) || (b.createDate - a.createDate));

    return { 
      activeList: activeList, 
      excludedNames: excludedNames,
      nfMale: attInfo.nfMale,     
      nfFemale: attInfo.nfFemale  
    };

  } catch (e) {
    throw new Error(e.toString());
  }
}

/** 2. 輕量化輪詢 API **/
function getQuickSyncData(type, userId) {
  return getSmartAttendanceList(type, userId);
}

/** 3. 同步至暫存區 (維持多人協作鎖定) **/
function syncClickToServer(name, isChecked, type, userId) {
  const ss = getSS();
  let sheet = ss.getSheetByName("SYNC_TEMP") || ss.insertSheet("SYNC_TEMP");
  if (sheet.getLastRow() === 0) sheet.appendRow(["姓名", "狀態", "類別", "時間", "操作者"]);

  const data = sheet.getDataRange().getValues();
  let foundRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === name && data[i][2] === type) { foundRow = i + 1; break; }
  }
  
  if (isChecked) {
    if (foundRow !== -1) {
      sheet.getRange(foundRow, 2).setValue("checked");
      sheet.getRange(foundRow, 4).setValue(new Date());
      sheet.getRange(foundRow, 5).setValue(userId);
    } else {
      sheet.appendRow([name, "checked", type, new Date(), userId]);
    }
  } else if (foundRow !== -1) {
    sheet.deleteRow(foundRow);
  }
  return "OK";
}

/** 4. 撤銷已送出的名單 (效能優化版：只讀底部 30 行) **/
function revokeAttendance(name, type, userId) {
  const ss = getSS();
  const todayStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy/M/d");
  const sheet = ss.getSheetByName(type + "點名紀錄");
  if (!sheet) return "錯誤：找不到紀錄表";

  const lastRow = sheet.getLastRow();
  let rowIndex = -1;

  if (lastRow > 0) {
    const numRows = Math.min(30, lastRow);
    const startRow = lastRow - numRows + 1;
    const data = sheet.getRange(startRow, 1, numRows, 4).getValues();
    
    for (let i = data.length - 1; i >= 0; i--) {
      if (startRow + i === 1) continue; 
      const d = (data[i][0] instanceof Date) ? Utilities.formatDate(data[i][0], "GMT+8", "yyyy/M/d") : data[i][0].toString();
      if (d === todayStr) { rowIndex = startRow + i; break; }
    }
  }

  if (rowIndex !== -1) {
    const rowData = sheet.getRange(rowIndex, 1, 1, 4).getValues()[0];
    let originalList = rowData[1].toString().split(/[,，、]\s*/);
    let newList = originalList.filter(item => item.replace(/（/g, '(').split('(')[0].trim() !== name.trim());
    
    const nfMale = Number(rowData[2] || 0);
    const nfFemale = Number(rowData[3] || 0);

    // 更新 B 欄 (名單)
    if (newList.length === 0 && nfMale === 0 && nfFemale === 0) {
      sheet.deleteRow(rowIndex);
    } else {
      sheet.getRange(rowIndex, 2).setValue(newList.join(", "));
    }
    syncClickToServer(name, true, type, userId);
    return "OK";
  }
  return "找不到紀錄";
}


/** * 5. 正式送出 (效能優化版：只讀底部 30 行 + 拒絕空紀錄防呆)
 * 邏輯：如果送出完全空的資料，直接不新增或刪除該列
 **/
function saveAttendance(date, presentList, type, nfMale, nfFemale) {
  const ss = getSS();
  const sheetName = type + "點名紀錄";
  let sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  
  if (sheet.getLastRow() === 0) sheet.appendRow(["出席日", "名單", "新朋友(男)", "新朋友(女)"]);

  const lastRow = sheet.getLastRow();
  let rowIndex = -1;
  let existingListStr = "";

  if (lastRow > 1) {
    const numRows = Math.min(30, lastRow);
    const startRow = lastRow - numRows + 1;
    const data = sheet.getRange(startRow, 1, numRows, 2).getValues();
    
    for (let i = data.length - 1; i >= 0; i--) {
      if (startRow + i === 1) continue; 
      const d = (data[i][0] instanceof Date) ? Utilities.formatDate(data[i][0], "GMT+8", "yyyy/M/d") : data[i][0];
      if (d === date) { 
        rowIndex = startRow + i; 
        existingListStr = data[i][1] || ""; 
        break; 
      }
    }
  }

  let finalNames;
  if (rowIndex !== -1 && existingListStr !== "") {
    let combinedSet = new Set([...existingListStr.split(/[,，、]\s*/), ...presentList]);
    let arr = Array.from(combinedSet).filter(n => n && n.trim() !== "");
    finalNames = arr.join(", ");
  } else {
    let arr = presentList.filter(n => n && n.trim() !== "");
    finalNames = arr.join(", ");
  }

  // 🛡️ 雙重防呆：判斷是否為「全空資料」
  const isEmptyRecord = (finalNames === "" && nfMale === 0 && nfFemale === 0);

  if (rowIndex !== -1) {
    // 如果本來有資料，但被改成全空 -> 直接刪除該列
    if (isEmptyRecord) {
      sheet.deleteRow(rowIndex);
    } else {
      sheet.getRange(rowIndex, 2).setValue(finalNames);
      sheet.getRange(rowIndex, 3).setValue(nfMale); 
      sheet.getRange(rowIndex, 4).setValue(nfFemale); 
    }
  } else {
    // 如果是新的一天，但送出了全空資料 -> 甚麼都不做 (不產生空列)
    if (!isEmptyRecord) {
      sheet.appendRow([date, finalNames, nfMale, nfFemale]);
    }
  }

  clearTempAfterSubmit(type, presentList);
  return isEmptyRecord ? `✅ 已清除當天空白紀錄` : `✅ 同步成功 (新朋友: 男 ${nfMale} 人, 女 ${nfFemale} 人)`;
}

// ==========================================
//  [會友系統支援] 更新與新增
// ==========================================

/** 6. [會友系統] 更新會友資料 */
function updateMember(oldName, newData) {
  const ss = getSS();
  const sheet = ss.getSheetByName("會友名單");
  const dataRange = sheet.getDataRange();
  const list = dataRange.getValues();
  
  if (!oldName) {
    return addMember(newData);
  }

  let rowIndex = -1;
  for (let i = 1; i < list.length; i++) {
    if (list[i][0] == oldName) { 
      rowIndex = i + 1; 
      break;
    }
  }

  if (rowIndex === -1) return "❌ 找不到原始資料，無法更新";

  sheet.getRange(rowIndex, 1).setValue(newData.name);       
  sheet.getRange(rowIndex, 2).setValue(newData.gender);     
  sheet.getRange(rowIndex, 4).setValue(newData.note);       
  sheet.getRange(rowIndex, 5).setValue(newData.isExcluded); 

  return "✅ 資料更新成功！";
}

/** 7. [點名系統] 快速新增會友 */
function addMember(memberData) {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName("會友名單");
    const nameStr = memberData.name.trim();
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const nameIdx = headers.indexOf("姓名");
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][nameIdx].toString().trim() === nameStr) return "⚠️ 該姓名已存在！";
    }

    let newRow = new Array(headers.length).fill("");
    newRow[headers.indexOf("姓名")] = nameStr;
    newRow[headers.indexOf("性別")] = memberData.gender;
    newRow[headers.indexOf("備註")] = memberData.note || "";
    newRow[headers.indexOf("建立日期")] = new Date();
    newRow[headers.indexOf("不列入統計")] = memberData.isExcluded || false;

    sheet.appendRow(newRow);
    return "✅ 成功新增會友 " + nameStr;
  } catch (e) {
    return "❌ 失敗: " + e.message;
  }
}

// --- 輔助函式庫 ---

/** * 讀取今日出席資訊 (效能優化版：只讀底部 30 行) */
function getTodayAttendanceInfo(ss, type, todayStr) {
  const sheet = ss.getSheetByName(type + "點名紀錄");
  if (!sheet) return { names: [], nfMale: 0, nfFemale: 0 };
  
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) return { names: [], nfMale: 0, nfFemale: 0 };

  const numRows = Math.min(30, lastRow);
  const startRow = lastRow - numRows + 1;
  const data = sheet.getRange(startRow, 1, numRows, 4).getValues();
  
  for (let i = data.length - 1; i >= 0; i--) {
    if (startRow + i === 1) continue; 
    let d = (data[i][0] instanceof Date) ? Utilities.formatDate(data[i][0], "GMT+8", "yyyy/M/d") : data[i][0].toString();
    
    if (d === todayStr) {
      let names = [];
      if (data[i][1]) {
        data[i][1].toString().split(/[,，、]\s*/).forEach(e => {
            const n = e.split('(')[0].trim();
            if (n) names.push(n);
        });
      }
      
      return { 
        names: names, 
        nfMale: Number(data[i][2] || 0),   // C 欄
        nfFemale: Number(data[i][3] || 0)  // D 欄
      };
    }
  }
  return { names: [], nfMale: 0, nfFemale: 0 };
}

function getSyncTempData(ss, type) {
  const tempSheet = ss.getSheetByName("SYNC_TEMP");
  const result = {};
  if (!tempSheet) return result;
  const data = tempSheet.getDataRange().getValues();
  const NOW = new Date().getTime();
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === type) {
      const isExpired = (NOW - new Date(data[i][3]).getTime()) > (5 * 60 * 1000);
      result[data[i][0].toString().trim()] = { checked: data[i][1] === "checked", operatorId: isExpired ? "" : data[i][4] };
    }
  }
  return result;
}

function clearTempAfterSubmit(type, names) {
  const ss = getSS();
  const sheet = ss.getSheetByName("SYNC_TEMP");
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const nameSet = new Set(names.map(n => n.split('(')[0].trim()));
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][2] === type && nameSet.has(data[i][0].toString().trim())) sheet.deleteRow(i + 1);
  }
}

/** * 統計近 90 天的出席次數 (效能大絕招：加入 CacheService 快取) */
function getAttendanceCountMap(ss, type) {
  const cache = CacheService.getScriptCache();
  const todayDateStr = Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd");
  const cacheKey = "ATT_MAP_" + type + "_" + todayDateStr;
  
  const cachedData = cache.get(cacheKey);
  if (cachedData) {
    return JSON.parse(cachedData);
  }

  const counts = {};
  const now = new Date();
  const cutoffDate = new Date();
  cutoffDate.setDate(now.getDate() - 90);
  const cutoffTime = cutoffDate.getTime();

  let targetSheets = (type === '聯合') ? ["台語點名紀錄", "華語點名紀錄", "聯合點名紀錄"] : [type + "點名紀錄"];

  targetSheets.forEach(sheetName => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;
    const data = sh.getDataRange().getValues();
    if (data.length <= 1) return;

    data.slice(1).forEach(row => {
      const rowDate = row[0];
      if (rowDate instanceof Date && rowDate.getTime() >= cutoffTime) {
        if (row[1]) {
          const listStr = row[1].toString().replace(/（/g, '(').replace(/）/g, ')');
          const names = listStr.split(/[,，、]\s*/);
          names.forEach(entry => {
            const name = entry.split('(')[0].trim();
            if (name) counts[name] = (counts[name] || 0) + 1;
          });
        }
      }
    });
  });
  
  cache.put(cacheKey, JSON.stringify(counts), 21600); // 存入快取，保存 6 小時
  return counts;
}

function getSS() {
  if (typeof SPREADSHEET_ID !== 'undefined') return SpreadsheetApp.openById(SPREADSHEET_ID);
  // 請填入你的 ID
  const MY_ID = "請在此填入你的試算表ID"; 
  return SpreadsheetApp.openById(MY_ID);
}
