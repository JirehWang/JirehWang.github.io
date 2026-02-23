// ==========================================
// ⚠️ 若上方已設定過 SPREADSHEET_ID 則可忽略
// const SPREADSHEET_ID = "你的試算表ID"; 
// ==========================================

/** 統計功能主入口 */
function getAttendanceStats(req) {
  const ss = getSS(); 
  const type = req.type; 
  
  if (type === '合計') {
    if (req.mode === 'single') return _getCombinedSingleStats(ss, req.date);
    else return _getCombinedRangeStats(ss, req.start, req.end);
  }

  if (req.mode === 'single') {
    return _getSingleDayStats(ss, type, req.date);
  } else {
    return _getRangeStats(ss, type, req.start, req.end);
  }
}

// ==========================================
//  1. [單一堂會] 單日統計
// ==========================================
function _getSingleDayStats(ss, type, dateStr) {
  const targetDate = new Date(dateStr);
  const formattedDate = Utilities.formatDate(targetDate, "GMT+8", "yyyy/M/d");
  const sheet = ss.getSheetByName(type + "點名紀錄");
  
  const result = { presentCount: 0, newFriends: 0, nfMale: 0, nfFemale: 0, presentMale: 0, presentFemale: 0, details: [] };
  if (!sheet) return result;

  const data = sheet.getDataRange().getValues();
  let presentNames = new Set();
  let listStr = ""; 
  
  for (let i = 1; i < data.length; i++) {
    const d = (data[i][0] instanceof Date) ? Utilities.formatDate(data[i][0], "GMT+8", "yyyy/M/d") : data[i][0];
    if (d === formattedDate) {
      listStr = data[i][1] ? data[i][1].toString() : "";
      if (listStr) {
        listStr.split(/[,，、]\s*/).forEach(n => {
           const cleanName = n.split('(')[0].trim();
           if(cleanName) presentNames.add(cleanName);
        });
      }
      // 讀取 C 欄與 D 欄 (新朋友)
      result.nfMale = Number(data[i][2] || 0);
      result.nfFemale = Number(data[i][3] || 0);
      break;
    }
  }
  
  // ✅ 直接算字串裡面的 (男) 和 (女) 作為名單出席人數
  result.presentMale = (listStr.match(/\(男\)/g) || []).length;
  result.presentFemale = (listStr.match(/\(女\)/g) || []).length;
  
  result.presentCount = presentNames.size;
  result.newFriends = result.nfMale + result.nfFemale;

  const memberSheet = ss.getSheetByName("會友名單");
  const memData = memberSheet.getDataRange().getValues();
  const nameIdx = memData[0].indexOf("姓名");
  const genderIdx = memData[0].indexOf("性別"); 
  const excludeIdx = memData[0].indexOf("不列入統計");

  for (let i = 1; i < memData.length; i++) {
    const name = memData[i][nameIdx];
    const gender = (genderIdx !== -1) ? memData[i][genderIdx] : "";
    const isExcluded = (memData[i][excludeIdx] === true || memData[i][excludeIdx] === "TRUE");
    const attended = presentNames.has(name);
    
    if (!isExcluded || attended) {
      result.details.push({
        name: name, gender: gender, count: attended ? 1 : 0, attended: attended, rate: 0 
      });
    }
  }

  result.details.sort((a, b) => (b.attended ? 1 : 0) - (a.attended ? 1 : 0));
  return result;
}

// ==========================================
//  2. [單一堂會] 區間統計 (加入空紀錄過濾)
// ==========================================
function _getRangeStats(ss, type, startStr, endStr) {
  const start = new Date(startStr).getTime();
  const end = new Date(endStr).getTime();
  const sheet = ss.getSheetByName(type + "點名紀錄");
  
  const result = { presentCount: 0, newFriends: 0, nfMale: 0, nfFemale: 0, presentMale: 0, presentFemale: 0, avgCount: 0, details: [] };
  if (!sheet) return result;

  const data = sheet.getDataRange().getValues();
  let validDays = 0, attendanceMap = {}, sumMemberCounts = 0, sumTotalCounts = 0;
  let totalPresentMale = 0, totalPresentFemale = 0;

  for (let i = 1; i < data.length; i++) {
    const rowDate = data[i][0] instanceof Date ? data[i][0] : new Date(data[i][0]);
    if (rowDate.getTime() >= start && rowDate.getTime() <= end) {
      
      let listStr = data[i][1] ? data[i][1].toString().trim() : "";
      const dayMale = Number(data[i][2] || 0);
      const dayFemale = Number(data[i][3] || 0);

      // 🛡️ 核心防呆：如果沒有任何名單，也沒有新朋友，這天就不算有效聚會場次
      if (listStr === "" && dayMale === 0 && dayFemale === 0) continue;

      validDays++; // 確定有資料，才將「總場次」加 1
      
      totalPresentMale += (listStr.match(/\(男\)/g) || []).length;
      totalPresentFemale += (listStr.match(/\(女\)/g) || []).length;

      let dayMemberCount = 0;
      if (listStr) {
        listStr.split(/[,，、]\s*/).forEach(entry => {
          const name = entry.split('(')[0].trim();
          if (name) { attendanceMap[name] = (attendanceMap[name] || 0) + 1; dayMemberCount++; }
        });
      }
      
      result.nfMale += dayMale;
      result.nfFemale += dayFemale;
      
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

  const memberSheet = ss.getSheetByName("會友名單");
  const memData = memberSheet.getDataRange().getValues();
  const nameIdx = memData[0].indexOf("姓名");
  const genderIdx = memData[0].indexOf("性別");
  const excludeIdx = memData[0].indexOf("不列入統計");
  
  for (let i = 1; i < memData.length; i++) {
    const name = memData[i][nameIdx];
    const gender = (genderIdx !== -1) ? memData[i][genderIdx] : "";
    const isExcluded = (memData[i][excludeIdx] === true || memData[i][excludeIdx] === "TRUE");
    
    if (!isExcluded) {
      const count = attendanceMap[name] || 0;
      result.details.push({
        name: name, gender: gender, count: count, rate: validDays > 0 ? Math.round((count / validDays) * 100) : 0
      });
    }
  }
  
  result.details.sort((a, b) => b.rate - a.rate);
  return result;
}

// ==========================================
//  3. [合計] 區間統計 (聯集) (加入空紀錄過濾)
// ==========================================
function _getCombinedRangeStats(ss, startStr, endStr) {
  const start = new Date(startStr).getTime();
  const end = new Date(endStr).getTime();
  const targetTypes = ['台語', '華語', '聯合'];

  let serviceDates = new Set(); 
  let memberDatesMap = {}; 
  let uniqueDailyAttendance = {}; 
  let nfMaleTotal = 0, nfFemaleTotal = 0; 
  
  targetTypes.forEach(type => {
    const sheet = ss.getSheetByName(type + "點名紀錄");
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const rowDateObj = data[i][0] instanceof Date ? data[i][0] : new Date(data[i][0]);
      const time = rowDateObj.getTime();

      if (time >= start && time <= end) {
        
        let listStr = data[i][1] ? data[i][1].toString().trim() : "";
        let dMale = Number(data[i][2] || 0);
        let dFemale = Number(data[i][3] || 0);

        // 🛡️ 核心防呆：只要整列是空的，就不計入這天的統計
        if (listStr === "" && dMale === 0 && dFemale === 0) continue;

        const dateStr = Utilities.formatDate(rowDateObj, "GMT+8", "yyyy/M/d");
        serviceDates.add(dateStr); // 確保這天有真實聚會，才算進場次中

        nfMaleTotal += dMale;
        nfFemaleTotal += dFemale;

        if (!uniqueDailyAttendance[dateStr]) uniqueDailyAttendance[dateStr] = {};

        if (listStr) {
           listStr.split(/[,，、]\s*/).forEach(n => {
             const name = n.split('(')[0].trim();
             const genderMatch = n.match(/\((男|女)\)/);
             const gender = genderMatch ? genderMatch[1] : "未知";

             if(name) {
               if(!memberDatesMap[name]) memberDatesMap[name] = new Set();
               memberDatesMap[name].add(dateStr);
               uniqueDailyAttendance[dateStr][name] = gender;
             }
           });
        }
      }
    }
  });

  const validDays = serviceDates.size; 
  let details = [];
  let sumAttendance = 0; 
  let totalMale = 0, totalFemale = 0;

  for (const date in uniqueDailyAttendance) {
    for (const name in uniqueDailyAttendance[date]) {
      if (uniqueDailyAttendance[date][name] === '男') totalMale++;
      if (uniqueDailyAttendance[date][name] === '女') totalFemale++;
    }
  }

  const memberSheet = ss.getSheetByName("會友名單");
  const memData = memberSheet.getDataRange().getValues();
  const nameIdx = memData[0].indexOf("姓名");
  const genderIdx = memData[0].indexOf("性別");
  const excludeIdx = memData[0].indexOf("不列入統計");

  for (let i = 1; i < memData.length; i++) {
    const name = memData[i][nameIdx];
    const gender = (genderIdx !== -1) ? memData[i][genderIdx] : "";
    const isExcluded = (memData[i][excludeIdx] === true || memData[i][excludeIdx] === "TRUE");

    if (!isExcluded) {
      const count = memberDatesMap[name] ? memberDatesMap[name].size : 0;
      sumAttendance += count;
      details.push({
        name: name, gender: gender, count: count, rate: validDays > 0 ? Math.round((count / validDays) * 100) : 0
      });
    }
  }

  details.sort((a, b) => b.rate - a.rate);
  const totalNewFriends = nfMaleTotal + nfFemaleTotal;

  return {
    presentCount: validDays > 0 ? Math.round(sumAttendance / validDays) : 0, 
    newFriends: totalNewFriends, 
    nfMale: nfMaleTotal,
    nfFemale: nfFemaleTotal,
    presentMale: validDays > 0 ? Math.round(totalMale / validDays) : 0,
    presentFemale: validDays > 0 ? Math.round(totalFemale / validDays) : 0,
    avgCount: validDays > 0 ? Math.round((sumAttendance + totalNewFriends) / validDays) : 0, 
    details: details
  };
}

// ==========================================
//  4. [合計] 單日統計 (聯集)
// ==========================================
function _getCombinedSingleStats(ss, dateStr) {
  const targetDate = new Date(dateStr);
  const formattedDate = Utilities.formatDate(targetDate, "GMT+8", "yyyy/M/d");
  const targetTypes = ['台語', '華語', '聯合'];

  let uniqueAttendees = {}; 
  let nfMaleTotal = 0, nfFemaleTotal = 0;

  targetTypes.forEach(type => {
    const sheet = ss.getSheetByName(type + "點名紀錄");
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const d = (data[i][0] instanceof Date) ? Utilities.formatDate(data[i][0], "GMT+8", "yyyy/M/d") : data[i][0];
      if (d === formattedDate) {
         if (data[i][1]) {
           data[i][1].toString().split(/[,，、]\s*/).forEach(n => {
             const name = n.split('(')[0].trim();
             const genderMatch = n.match(/\((男|女)\)/);
             if(name) {
                // 存入物件自動去重，並記錄字串抓到的性別
                uniqueAttendees[name] = genderMatch ? genderMatch[1] : "未知"; 
             }
           });
         }
         nfMaleTotal += Number(data[i][2] || 0);
         nfFemaleTotal += Number(data[i][3] || 0);
         break; 
      }
    }
  });

  // 計算去重後的男女
  let presentMale = 0, presentFemale = 0;
  for (const name in uniqueAttendees) {
    if (uniqueAttendees[name] === '男') presentMale++;
    if (uniqueAttendees[name] === '女') presentFemale++;
  }

  const result = { 
    presentCount: Object.keys(uniqueAttendees).length, 
    newFriends: nfMaleTotal + nfFemaleTotal, 
    nfMale: nfMaleTotal,
    nfFemale: nfFemaleTotal,
    presentMale: presentMale,
    presentFemale: presentFemale,
    details: [] 
  };

  const memberSheet = ss.getSheetByName("會友名單");
  const memData = memberSheet.getDataRange().getValues();
  const nameIdx = memData[0].indexOf("姓名");
  const genderIdx = memData[0].indexOf("性別");
  const excludeIdx = memData[0].indexOf("不列入統計");

  for (let i = 1; i < memData.length; i++) {
    const name = memData[i][nameIdx];
    const gender = (genderIdx !== -1) ? memData[i][genderIdx] : "";
    const isExcluded = (memData[i][excludeIdx] === true || memData[i][excludeIdx] === "TRUE");
    const attended = uniqueAttendees.hasOwnProperty(name);

    if (!isExcluded || attended) {
      result.details.push({
        name: name, gender: gender, count: attended ? 1 : 0, attended: attended, rate: 0
      });
    }
  }
  
  result.details.sort((a, b) => (b.attended ? 1 : 0) - (a.attended ? 1 : 0));
  return result;
}

function getSS() {
  if (typeof SPREADSHEET_ID !== 'undefined') return SpreadsheetApp.openById(SPREADSHEET_ID);
  const MY_ID = "請在此填入你的試算表ID"; 
  return SpreadsheetApp.openById(MY_ID);
}
