/**
 * 日期轉數字工具
 */
function toDateNum(dateObj) {
  if (!dateObj) return 0;
  const d = new Date(dateObj);
  if (isNaN(d.getTime())) return 0;
  const y = d.getFullYear();
  const m = ("0" + (d.getMonth() + 1)).slice(-2);
  const day = ("0" + d.getDate()).slice(-2);
  return parseInt(y + m + day, 10);
}

/**
 * 統計主入口
 */
function getAttendanceStats(req) {
  const ss = getSS();
  const type = req.type;
  const baseSheet = req.baseSheet || "會友名單";
  let data;

  if (type.indexOf('合計') !== -1) {
    const targetGroups = req.targetGroups || [];
    if (targetGroups.length === 0) return { presentCount: 0, newFriends: 0, nfMale: 0, nfFemale: 0, presentMale: 0, presentFemale: 0, details: [] };
    if (req.mode === 'single') data = _getCombinedSingleStats(ss, req.date, baseSheet, targetGroups);
    else data = _getCombinedRangeStats(ss, req.start, req.end, baseSheet, targetGroups);
  } else {
    if (req.mode === 'single') data = _getSingleDayStats(ss, type, req.date, baseSheet);
    else data = _getRangeStats(ss, type, req.start, req.end, baseSheet);
  }

  const groupMembersMap = getSmallGroupMembersList();
  if (data && data.details && data.details.length > 0) {
    data.details.forEach(row => { row.inGroup = !!groupMembersMap[row.name]; });
  }
  return data;
}

/**
 * 1. 單日統計
 */
function _getSingleDayStats(ss, type, dateStr, baseSheetName) {
  const targetDateNum = toDateNum(dateStr);
  const sheet = ss.getSheetByName(type + "點名紀錄");
  const result = { presentCount: 0, newFriends: 0, nfMale: 0, nfFemale: 0, presentMale: 0, presentFemale: 0, details: [] };
  if (!sheet) return result;
  const data = sheet.getDataRange().getValues();
  let presentNames = new Set(), listStr = "";
  for (let i = 1; i < data.length; i++) {
    if (toDateNum(data[i][0]) === targetDateNum) {
      listStr = data[i][1] ? data[i][1].toString() : "";
      if (listStr) listStr.split(/[,，、]\s*/).forEach(n => { const c = n.split('(')[0].trim(); if (c) presentNames.add(c); });
      result.nfMale = Number(data[i][2] || 0);
      result.nfFemale = Number(data[i][3] || 0);
      break;
    }
  }
  result.presentMale = (listStr.match(/\(男\)/g) || []).length;
  result.presentFemale = (listStr.match(/\(女\)/g) || []).length;
  result.presentCount = presentNames.size;
  result.newFriends = result.nfMale + result.nfFemale;
  const memberSheet = ss.getSheetByName(baseSheetName);
  if (!memberSheet) return result;
  const memData = memberSheet.getDataRange().getValues();
  const nameIdx = memData[0].indexOf("姓名");
  const genderIdx = memData[0].indexOf("性別");
  const excludeIdx = memData[0].indexOf("不列入統計");
  for (let i = 1; i < memData.length; i++) {
    const name = nameIdx !== -1 ? memData[i][nameIdx] : memData[i][0];
    const gender = genderIdx !== -1 ? memData[i][genderIdx] : "";
    const isExcluded = excludeIdx !== -1 ? (memData[i][excludeIdx] === true || memData[i][excludeIdx] === "TRUE") : false;
    const attended = presentNames.has(name);
    if (!isExcluded || attended) result.details.push({ name, gender, count: attended ? 1 : 0, attended, rate: 0 });
  }
  result.details.sort((a, b) => (b.attended ? 1 : 0) - (a.attended ? 1 : 0));
  return result;
}

/**
 * 2. 區間統計
 */
function _getRangeStats(ss, type, startStr, endStr, baseSheetName) {
  const startNum = toDateNum(startStr), endNum = toDateNum(endStr);
  const sheet = ss.getSheetByName(type + "點名紀錄");
  const result = { presentCount: 0, newFriends: 0, nfMale: 0, nfFemale: 0, presentMale: 0, presentFemale: 0, avgCount: 0, details: [] };
  if (!sheet) return result;
  const data = sheet.getDataRange().getValues();
  let validDays = 0, attendanceMap = {}, sumMemberCounts = 0, sumTotalCounts = 0, totalPresentMale = 0, totalPresentFemale = 0;
  for (let i = 1; i < data.length; i++) {
    const rowDateNum = toDateNum(data[i][0]);
    if (rowDateNum >= startNum && rowDateNum < endNum) {
      let listStr = data[i][1] ? data[i][1].toString().trim() : "";
      const dayMale = Number(data[i][2] || 0), dayFemale = Number(data[i][3] || 0);
      if (listStr === "" && dayMale === 0 && dayFemale === 0) continue;
      validDays++;
      totalPresentMale += (listStr.match(/\(男\)/g) || []).length;
      totalPresentFemale += (listStr.match(/\(女\)/g) || []).length;
      let dayMemberCount = 0;
      if (listStr) {
        listStr.split(/[,，、]\s*/).forEach(entry => {
          const name = entry.split('(')[0].trim();
          if (name) { attendanceMap[name] = (attendanceMap[name] || 0) + 1; dayMemberCount++; }
        });
      }
      result.nfMale += dayMale; result.nfFemale += dayFemale;
      sumMemberCounts += dayMemberCount;
      sumTotalCounts += (dayMemberCount + dayMale + dayFemale);
    }
  }
  result.newFriends = result.nfMale + result.nfFemale;
  if (validDays > 0) {
    result.avgCount = Math.round(sumTotalCounts / validDays);
    result.presentCount = Math.round(sumMemberCounts / validDays);
    result.presentMale = Math.round(totalPresentMale / validDays);
    result.presentFemale = Math.round(totalPresentFemale / validDays);
  }
  const memberSheet = ss.getSheetByName(baseSheetName);
  if (!memberSheet) return result;
  const memData = memberSheet.getDataRange().getValues();
  const nameIdx = memData[0].indexOf("姓名"), genderIdx = memData[0].indexOf("性別"), excludeIdx = memData[0].indexOf("不列入統計");
  for (let i = 1; i < memData.length; i++) {
    const name = nameIdx !== -1 ? memData[i][nameIdx] : memData[i][0];
    const gender = genderIdx !== -1 ? memData[i][genderIdx] : "";
    const isExcluded = excludeIdx !== -1 ? (memData[i][excludeIdx] === true || memData[i][excludeIdx] === "TRUE") : false;
    if (!isExcluded) {
      const count = attendanceMap[name] || 0;
      result.details.push({ name, gender, count, rate: validDays > 0 ? Math.round((count / validDays) * 100) : 0 });
    }
  }
  result.details.sort((a, b) => b.rate - a.rate);
  return result;
}

/**
 * 3. 合計區間統計
 */
function _getCombinedRangeStats(ss, startStr, endStr, baseSheetName, targetTypes) {
  const startNum = toDateNum(startStr), endNum = toDateNum(endStr);
  let serviceDates = new Set(), memberDatesMap = {}, uniqueDailyAttendance = {}, nfMaleTotal = 0, nfFemaleTotal = 0;
  targetTypes.forEach(type => {
    const sheet = ss.getSheetByName(type + "點名紀錄");
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const rowDateNum = toDateNum(data[i][0]);
      if (rowDateNum >= startNum && rowDateNum < endNum) {
        let listStr = data[i][1] ? data[i][1].toString().trim() : "";
        let dMale = Number(data[i][2] || 0), dFemale = Number(data[i][3] || 0);
        if (listStr === "" && dMale === 0 && dFemale === 0) continue;
        const dateKey = rowDateNum.toString();
        serviceDates.add(dateKey); nfMaleTotal += dMale; nfFemaleTotal += dFemale;
        if (!uniqueDailyAttendance[dateKey]) uniqueDailyAttendance[dateKey] = {};
        if (listStr) {
          listStr.split(/[,，、]\s*/).forEach(n => {
            const name = n.split('(')[0].trim();
            const genderMatch = n.match(/\((男|女)\)/);
            const gender = genderMatch ? genderMatch[1] : "未知";
            if (name) {
              if (!memberDatesMap[name]) memberDatesMap[name] = new Set();
              memberDatesMap[name].add(dateKey);
              uniqueDailyAttendance[dateKey][name] = gender;
            }
          });
        }
      }
    }
  });
  const validDays = serviceDates.size;
  let details = [], sumAttendance = 0, totalMale = 0, totalFemale = 0;
  for (const date in uniqueDailyAttendance) {
    for (const name in uniqueDailyAttendance[date]) {
      if (uniqueDailyAttendance[date][name] === '男') totalMale++;
      if (uniqueDailyAttendance[date][name] === '女') totalFemale++;
    }
  }
  const memberSheet = ss.getSheetByName(baseSheetName);
  if (!memberSheet) return { presentCount: 0, newFriends: 0, nfMale: 0, nfFemale: 0, presentMale: 0, presentFemale: 0, avgCount: 0, details: [] };
  const memData = memberSheet.getDataRange().getValues();
  const nameIdx = memData[0].indexOf("姓名"), genderIdx = memData[0].indexOf("性別"), excludeIdx = memData[0].indexOf("不列入統計");
  for (let i = 1; i < memData.length; i++) {
    const name = nameIdx !== -1 ? memData[i][nameIdx] : memData[i][0];
    const gender = genderIdx !== -1 ? memData[i][genderIdx] : "";
    const isExcluded = excludeIdx !== -1 ? (memData[i][excludeIdx] === true || memData[i][excludeIdx] === "TRUE") : false;
    if (!isExcluded) {
      const count = memberDatesMap[name] ? memberDatesMap[name].size : 0;
      sumAttendance += count;
      details.push({ name, gender, count, rate: validDays > 0 ? Math.round((count / validDays) * 100) : 0 });
    }
  }
  details.sort((a, b) => b.rate - a.rate);
  return {
    presentCount: validDays > 0 ? Math.round(sumAttendance / validDays) : 0,
    newFriends: nfMaleTotal + nfFemaleTotal,
    nfMale: nfMaleTotal, nfFemale: nfFemaleTotal,
    presentMale: validDays > 0 ? Math.round(totalMale / validDays) : 0,
    presentFemale: validDays > 0 ? Math.round(totalFemale / validDays) : 0,
    avgCount: validDays > 0 ? Math.round((sumAttendance + nfMaleTotal + nfFemaleTotal) / validDays) : 0,
    details
  };
}

/**
 * 4. 合計單日統計
 */
function _getCombinedSingleStats(ss, dateStr, baseSheetName, targetTypes) {
  const targetDateNum = toDateNum(dateStr);
  let uniqueAttendees = {}, nfMaleTotal = 0, nfFemaleTotal = 0;
  targetTypes.forEach(type => {
    const sheet = ss.getSheetByName(type + "點名紀錄");
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (toDateNum(data[i][0]) === targetDateNum) {
        if (data[i][1]) {
          data[i][1].toString().split(/[,，、]\s*/).forEach(n => {
            const name = n.split('(')[0].trim();
            const genderMatch = n.match(/\((男|女)\)/);
            if (name) uniqueAttendees[name] = genderMatch ? genderMatch[1] : "未知";
          });
        }
        nfMaleTotal += Number(data[i][2] || 0);
        nfFemaleTotal += Number(data[i][3] || 0);
        break;
      }
    }
  });
  let presentMale = 0, presentFemale = 0;
  for (const name in uniqueAttendees) {
    if (uniqueAttendees[name] === '男') presentMale++;
    if (uniqueAttendees[name] === '女') presentFemale++;
  }
  const result = { presentCount: Object.keys(uniqueAttendees).length, newFriends: nfMaleTotal + nfFemaleTotal, nfMale: nfMaleTotal, nfFemale: nfFemaleTotal, presentMale, presentFemale, details: [] };
  const memberSheet = ss.getSheetByName(baseSheetName);
  if (!memberSheet) return result;
  const memData = memberSheet.getDataRange().getValues();
  const nameIdx = memData[0].indexOf("姓名"), genderIdx = memData[0].indexOf("性別"), excludeIdx = memData[0].indexOf("不列入統計");
  for (let i = 1; i < memData.length; i++) {
    const name = nameIdx !== -1 ? memData[i][nameIdx] : memData[i][0];
    const gender = genderIdx !== -1 ? memData[i][genderIdx] : "";
    const isExcluded = excludeIdx !== -1 ? (memData[i][excludeIdx] === true || memData[i][excludeIdx] === "TRUE") : false;
    const attended = uniqueAttendees.hasOwnProperty(name);
    if (!isExcluded || attended) result.details.push({ name, gender, count: attended ? 1 : 0, attended, rate: 0 });
  }
  result.details.sort((a, b) => (b.attended ? 1 : 0) - (a.attended ? 1 : 0));
  return result;
}

/**
 * 小組名單查詢
 */
function getSmallGroupMembersList() {
  const GROUP_SS_ID = "1opAz6PFYveCF4oP9d4c_Dppd2SWBgE_d4zk0a0dRIoU";
  let groupMembers = {};
  try {
    const ss = SpreadsheetApp.openById(GROUP_SS_ID);
    ss.getSheets().forEach(sheet => {
      if (sheet.getName().indexOf("_名單") !== -1) {
        const lastRow = sheet.getLastRow();
        if (lastRow >= 2) {
          sheet.getRange(2, 1, lastRow - 1, 1).getValues().forEach(row => {
            if (row[0]) { const name = row[0].toString().trim(); if (name) groupMembers[name] = true; }
          });
        }
      }
    });
  } catch (e) {
    console.error("無法讀取小組名單: " + e.toString());
  }
  return groupMembers;
}

/**
 * 出席頻率變化分析
 */
function getAttendanceTrend(req) {
  const ss = getSS();
  const baseSheetName = req.baseSheet || "會友名單";
  const startNum = toDateNum(req.start);
  const endNum   = toDateNum(req.end);

  // --- 1. 決定要查的點名表 ---
  let sheetNames = [];
  if (req.targetGroups && req.targetGroups.length > 0) {
      sheetNames = req.targetGroups.map(function(t) { return t + "點名紀錄"; });
  } else {
      sheetNames = [req.type + "點名紀錄"];
  }


  // --- 2. 收集場次資料：每場日期 -> 出席人名 Set ---
  const sessionMap = {};  

  sheetNames.forEach(function(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const rowDateNum = toDateNum(data[i][0]);
      if (rowDateNum < startNum || rowDateNum >= endNum) continue;

      const listStr = data[i][1] ? data[i][1].toString().trim() : "";
      const nfMale  = Number(data[i][2] || 0);
      const nfFemale = Number(data[i][3] || 0);

      if (listStr === "" && nfMale === 0 && nfFemale === 0) continue;

      const dateKey = rowDateNum.toString();

      if (!sessionMap[dateKey]) {
        sessionMap[dateKey] = { names: new Set(), nfCount: 0 };
      }

      if (listStr) {
        listStr.split(/[,，、]\s*/).forEach(function(entry) {
          const name = entry.split('(')[0].trim();
          if (name) sessionMap[dateKey].names.add(name);
        });
      }
      sessionMap[dateKey].nfCount = Math.max(
        sessionMap[dateKey].nfCount,
        nfMale + nfFemale
      );
    }
  });

  // --- 3. 整理場次清單（按日期排序）---
  const allDates = Object.keys(sessionMap).sort();
  const N = allDates.length;

  if (N === 0) {
    throw new Error("在指定日期前找不到任何有效場次紀錄。");
  }

  // --- 4. 載入會友名單 ---
  const memberSheet = ss.getSheetByName(baseSheetName);
  if (!memberSheet) throw new Error("找不到名單：" + baseSheetName);

  const memData   = memberSheet.getDataRange().getValues();
  const nameIdx   = memData[0].indexOf("姓名");
  const genderIdx = memData[0].indexOf("性別");
  const excludeIdx = memData[0].indexOf("不列入統計");

  const members = {};  
  for (let i = 1; i < memData.length; i++) {
    const name   = nameIdx   !== -1 ? memData[i][nameIdx]   : memData[i][0];
    const gender = genderIdx !== -1 ? memData[i][genderIdx] : "";
    const isExcluded = excludeIdx !== -1
      ? (memData[i][excludeIdx] === true || memData[i][excludeIdx] === "TRUE")
      : false;
    const nameStr = name ? name.toString().trim() : "";
    if (nameStr && !isExcluded) {
      members[nameStr] = { gender: gender || "-" };
    }
  }

  // --- 5. 設定時間窗：近期 56 天 (八週)，歷史最長 365 天 ---
  const RECENT_WINDOW_DAYS = 56;
  const MAX_HISTORY_DAYS = 365;

  function parseDateKey(dKey) {
    const y = parseInt(dKey.substring(0, 4), 10);
    const m = parseInt(dKey.substring(4, 6), 10) - 1;
    const d = parseInt(dKey.substring(6, 8), 10);
    return new Date(y, m, d);
  }

  const latestDateKey = allDates[allDates.length - 1];
  const latestDate = parseDateKey(latestDateKey);

  const cutoffDate = new Date(latestDate);
  cutoffDate.setDate(cutoffDate.getDate() - RECENT_WINDOW_DAYS);

  const maxHistoryDate = new Date(cutoffDate);
  maxHistoryDate.setDate(maxHistoryDate.getDate() - MAX_HISTORY_DAYS);

  const recentDates = [];
  const generalHistoryDates = [];

  allDates.forEach(function(d) {
    const dDate = parseDateKey(d);
    if (dDate >= cutoffDate) {
      recentDates.push(d);
    } else if (dDate >= maxHistoryDate) {
      generalHistoryDates.push(d);
    }
  });

  const recentCount = recentDates.length;
  const details = [];

  // --- 6. 計算衰退指標與綜合分數 ---
  Object.keys(members).forEach(function(name) {
    let firstAppearanceDateStr = null;

    for (let i = 0; i < allDates.length; i++) {
      if (sessionMap[allDates[i]].names.has(name)) {
        firstAppearanceDateStr = allDates[i];
        break;
      }
    }

    if (!firstAppearanceDateStr) return; 

    const firstAppDate = parseDateKey(firstAppearanceDateStr);
    const individualStartDate = firstAppDate > maxHistoryDate ? firstAppDate : maxHistoryDate;

    const historyTimeDiff = cutoffDate.getTime() - individualStartDate.getTime();
    const historyDays = Math.ceil(historyTimeDiff / (1000 * 3600 * 24));

    if (historyDays < RECENT_WINDOW_DAYS) return;

    let historyAttended = 0;
    let individualHistoryCount = 0;

    generalHistoryDates.forEach(function(d) {
      const dDate = parseDateKey(d);
      if (dDate >= individualStartDate) {
        individualHistoryCount++;
        if (sessionMap[d].names.has(name)) {
          historyAttended++;
        }
      }
    });

    let recentAttended = 0;
    recentDates.forEach(function(d) {
      if (sessionMap[d].names.has(name)) {
        recentAttended++;
      }
    });

    let consecutiveMisses = 0;
    let lastAttendedDateKey = null;
    for (let i = allDates.length - 1; i >= 0; i--) {
      if (!sessionMap[allDates[i]].names.has(name)) {
        consecutiveMisses++;
      } else {
        lastAttendedDateKey = allDates[i];
        break;
      }
    }

    let missingThreeWeeks = false;
    if (lastAttendedDateKey) {
      const lastDate = parseDateKey(lastAttendedDateKey);
      const diffDays = Math.floor((latestDate.getTime() - lastDate.getTime()) / (1000 * 3600 * 24));
      if (diffDays >= 21) {
        missingThreeWeeks = true;
      }
    } else if (consecutiveMisses > 0) {
      missingThreeWeeks = true; 
    }

    const historyRate = individualHistoryCount > 0 ? Math.round((historyAttended / individualHistoryCount) * 100) : 0;
    const recentRate = recentCount > 0 ? Math.round((recentAttended / recentCount) * 100) : 0;

    const rateDrop = historyRate - recentRate; 
    const dropScore = rateDrop > 0 ? rateDrop : 0;

    details.push({
      name: name,
      gender: members[name].gender,
      historyRate: historyRate,           
      recentRate: recentRate,             
      rateDrop: rateDrop,                 
      consecutiveMisses: consecutiveMisses, 
      dropScore: dropScore,
      missingThreeWeeks: missingThreeWeeks,
      warningDesc: missingThreeWeeks ? "⚠️ 已連續三週未出席" : ""
    });
  });

  details.sort(function(a, b) {
    return b.dropScore - a.dropScore;
  });

  // --- 7. 格式化日期用於回傳 ---
  function fmt(dateKey) {
    var y = dateKey.substring(0, 4);
    var m = dateKey.substring(4, 6);
    var d = dateKey.substring(6, 8);
    return y + "/" + m + "/" + d;
  }

  return {
    periodHistory: generalHistoryDates.length > 0 ? (fmt(generalHistoryDates[0]) + " ~ " + fmt(generalHistoryDates[generalHistoryDates.length - 1])) : "無",
    periodRecent: recentDates.length > 0 ? (fmt(recentDates[0]) + " ~ " + fmt(recentDates[recentDates.length - 1])) : "無",
    sessionsHistory: generalHistoryDates.length,
    sessionsRecent: recentDates.length,
    details: details
  };
}