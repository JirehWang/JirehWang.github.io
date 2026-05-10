/**
 * 1. 取得智慧名單
 */
function getSmartAttendanceList(type, userId) {
  const ss = getSS();
  try {
    const memberSheet = ss.getSheetByName("會友名單");
    if (!memberSheet) throw new Error("找不到工作表：'會友名單'");
    const members = memberSheet.getDataRange().getValues();
    const headers = members[0];
    const todayStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy/M/d");
    const nameIdx = headers.indexOf("姓名");
    const genderIdx = headers.indexOf("性別");
    const dateIdx = headers.indexOf("建立日期");
    const excludeIdx = headers.indexOf("不列入統計");
    const idIdx = headers.indexOf("系統編號");
    const attInfo = getTodayAttendanceInfo(ss, type, todayStr);
    const permanentSet = new Set(attInfo.names);
    const attendanceMap = getAttendanceCountMap(ss, type);
    const syncTempData = getSyncTempData(ss, type);
    let activeList = [], excludedNames = [];

    members.slice(1).forEach(row => {
      const name = row[nameIdx] ? row[nameIdx].toString().trim() : "";
      if (!name) return;
      const memberId = (idIdx !== -1 && row[idIdx]) ? row[idIdx].toString().trim() : "";
      if (row[excludeIdx] === true || row[excludeIdx] === "TRUE") {
        excludedNames.push(name);
      } else {
        const isSubmitted = permanentSet.has(name);
        const temp = syncTempData[name] || { checked: false, operatorId: "" };
        activeList.push({
          id: memberId, name: name,
          gender: (genderIdx !== -1) ? (row[genderIdx] || "未知") : "未知",
          createDate: (dateIdx !== -1 && row[dateIdx] instanceof Date) ? row[dateIdx].getTime() : 0,
          count: attendanceMap[name] || 0,
          isChecked: isSubmitted || temp.checked,
          isSubmitted: isSubmitted,
          operatorId: temp.operatorId
        });
      }
    });
    activeList.sort((a, b) => (b.count - a.count) || (b.createDate - a.createDate));
    return { activeList, excludedNames, nfMale: attInfo.nfMale, nfFemale: attInfo.nfFemale };
  } catch (e) {
    throw new Error(e.toString());
  }
}

/**
 * 2. 輕量化輪詢
 */
function getQuickSyncData(type, userId) {
  return getSmartAttendanceList(type, userId);
}

/**
 * 3. 同步暫存區
 */
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

/**
 * 4. 撤銷已送出名單
 */
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
    if (newList.length === 0 && nfMale === 0 && nfFemale === 0) {
      sheet.deleteRow(rowIndex);
    } else {
      sheet.getRange(rowIndex, 2).setValue(newList.join(", "));
    }
    return "OK";
  }
  return "找不到紀錄";
}

/**
 * 5. 正式送出點名
 */
function saveAttendance(date, presentList, type, nfMale, nfFemale) {
  const ss = getSS();
  const sheetName = type + "點名紀錄";
  let sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  if (sheet.getLastRow() === 0) sheet.appendRow(["出席日", "名單", "新朋友(男)", "新朋友(女)"]);
  const lastRow = sheet.getLastRow();
  let rowIndex = -1, existingListStr = "";
  if (lastRow > 1) {
    const numRows = Math.min(30, lastRow);
    const startRow = lastRow - numRows + 1;
    const data = sheet.getRange(startRow, 1, numRows, 2).getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      if (startRow + i === 1) continue;
      const d = (data[i][0] instanceof Date) ? Utilities.formatDate(data[i][0], "GMT+8", "yyyy/M/d") : data[i][0];
      if (d === date) { rowIndex = startRow + i; existingListStr = data[i][1] || ""; break; }
    }
  }
  let finalNames;
  if (rowIndex !== -1 && existingListStr !== "") {
    let arr = Array.from(new Set([...existingListStr.split(/[,，、]\s*/), ...presentList])).filter(n => n && n.trim() !== "");
    finalNames = arr.join(", ");
  } else {
    finalNames = presentList.filter(n => n && n.trim() !== "").join(", ");
  }
  const isEmptyRecord = (finalNames === "" && nfMale === 0 && nfFemale === 0);
  if (rowIndex !== -1) {
    if (isEmptyRecord) { sheet.deleteRow(rowIndex); }
    else {
      sheet.getRange(rowIndex, 2).setValue(finalNames);
      sheet.getRange(rowIndex, 3).setValue(nfMale);
      sheet.getRange(rowIndex, 4).setValue(nfFemale);
    }
  } else {
    if (!isEmptyRecord) sheet.appendRow([date, finalNames, nfMale, nfFemale]);
  }
  clearTempAfterSubmit(type, presentList);
  return isEmptyRecord ? `✅ 已清除當天空白紀錄` : `✅ 同步成功 (新朋友: 男 ${nfMale} 人, 女 ${nfFemale} 人)`;
}

/**
 * 6. 取得分類結構
 */
function getGroupConfig() {
  const ss = getSS();
  let listSheet = ss.getSheetByName("點名系統清單");
  if (!listSheet) {
    listSheet = ss.insertSheet("點名系統清單");
    listSheet.appendRow(["點名類別", "群組名稱"]);
    listSheet.getRange("A1:B1").setFontWeight("bold").setBackground("#f3f3f3");
    listSheet.setFrozenRows(1);
    const defaultData = [
      ["禮拜", "台語"], ["禮拜", "華語"], ["禮拜", "聯合"],
      ["主日學", "主日學A班"], ["主日學", "主日學B班"]
    ];
    listSheet.getRange(2, 1, defaultData.length, 2).setValues(defaultData);
  }
  const data = listSheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < data.length; i++) {
    const category = data[i][0] ? data[i][0].toString().trim() : "";
    const groupName = data[i][1] ? data[i][1].toString().trim() : "";
    if (category && groupName) {
      if (!config[category]) config[category] = [];
      config[category].push(groupName);
    }
  }
  return config;
}

/**
 * 7. 建立新群組
 */
function createAttendanceGroup(category, groupName) {
  if (!groupName || groupName.trim() === "") throw new Error("群組名稱不能為空！");
  const cleanName = groupName.trim();
  const sheetName = cleanName + "點名紀錄";
  const ss = getSS();
  if (ss.getSheetByName(sheetName)) throw new Error("⚠️ 建立失敗：[" + cleanName + "] 的分頁已經存在囉！");
  const newSheet = ss.insertSheet(sheetName);
  newSheet.appendRow(["日期", "點名單", "新朋友(男)", "新朋友(女)"]);
  newSheet.getRange("A1:D1").setFontWeight("bold").setBackground("#f3f3f3");
  newSheet.setFrozenRows(1);
  let listSheet = ss.getSheetByName("點名系統清單");
  if (!listSheet) { getGroupConfig(); listSheet = ss.getSheetByName("點名系統清單"); }
  const data = listSheet.getDataRange().getValues();
  let isExist = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === category && data[i][1] === cleanName) { isExist = true; break; }
  }
  if (!isExist) listSheet.appendRow([category, cleanName]);
  return getGroupConfig();
}

// ==========================================
//  輔助函式
// ==========================================

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
      return { names, nfMale: Number(data[i][2] || 0), nfFemale: Number(data[i][3] || 0) };
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

function getAttendanceCountMap(ss, type) {
  const cache = CacheService.getScriptCache();
  const todayDateStr = Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd");
  const cacheKey = "ATT_MAP_" + type + "_" + todayDateStr;
  const cachedData = cache.get(cacheKey);
  if (cachedData) return JSON.parse(cachedData);
  const counts = {};
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - 90);
  const cutoffTime = cutoffDate.getTime();
  let targetSheets = (type === '聯合') ? ["台語點名紀錄", "華語點名紀錄", "聯合點名紀錄"] : [type + "點名紀錄"];
  targetSheets.forEach(sheetName => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;
    const data = sh.getDataRange().getValues();
    if (data.length <= 1) return;
    data.slice(1).forEach(row => {
      if (row[0] instanceof Date && row[0].getTime() >= cutoffTime && row[1]) {
        row[1].toString().replace(/（/g, '(').replace(/）/g, ')').split(/[,，、]\s*/).forEach(entry => {
          const name = entry.split('(')[0].trim();
          if (name) counts[name] = (counts[name] || 0) + 1;
        });
      }
    });
  });
  cache.put(cacheKey, JSON.stringify(counts), 21600);
  return counts;
}