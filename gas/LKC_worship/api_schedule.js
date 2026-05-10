// ==========================================
// api_schedule.gs (服事表核心邏輯：外部框架驅動版)
// ==========================================

function toDateNum(dateObj) {
  if (!dateObj) return 0;
  if (dateObj instanceof Date) {
    if (isNaN(dateObj.getTime())) return 0;
    const tz = SpreadsheetApp.openById(SPREADSHEET_ID).getSpreadsheetTimeZone();
    return parseInt(Utilities.formatDate(dateObj, tz, "yyyyMMdd"), 10);
  }
  const str = String(dateObj).trim();
  const parts = str.split('-');
  if (parts.length === 3) return parseInt(parts[0] + parts[1].padStart(2, '0') + parts[2].padStart(2, '0'), 10);
  return 0;
}

/**
 * 🌟 核心升級：以「外部聚會資料」為主體框架，疊加本地排班
 * 🔧 修正：改用「日期_聚會名稱」複合 key，避免同日期多個相同聚會類別造成資料覆蓋
 */
function getMergedSchedule(year, quarter) {
  const reqYear = String(year).trim();
  const reqQuarter = String(quarter).trim();

  // 計算該季度的數字區間 (例如 Q1 -> 20260100 ~ 20260331)
  let startMonth = (parseInt(reqQuarter.replace('Q', '')) - 1) * 3 + 1;
  let startDateNum = parseInt(reqYear) * 10000 + startMonth * 100;
  let endDateNum = parseInt(reqYear) * 10000 + (startMonth + 2) * 100 + 31;

  let externalMeetings = []; 
  let externalSermonDict = {};
  const extUrl = 'https://docs.google.com/spreadsheets/d/1tKI5k7HwI9S2bTRV6RuKrzxBFGzROYHytsFYHuH8H7E/edit';

  try {
    const extSs = SpreadsheetApp.openByUrl(extUrl);
    
    // 1. 抓取外部聚會框架
    const meetingSheet = extSs.getSheetByName('聚會資料');
    let meetingMap = {};
    if (meetingSheet) {
      const mData = meetingSheet.getDataRange().getValues();
      const headers = mData[0];
      const idIdx = headers.indexOf('ID');
      const dateIdx = headers.indexOf('日期');
      const nameIdx = headers.indexOf('聚會名稱');
      const typeIdx = headers.indexOf('聚會類別');

      mData.slice(1).forEach(row => {
        if (idIdx > -1 && row[idIdx]) {
          let dObj = row[dateIdx];
          let dNum = toDateNum(dObj);
          let dateStr = dObj instanceof Date ? Utilities.formatDate(dObj, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(dObj).trim();
          let type = typeIdx > -1 ? String(row[typeIdx]).trim() : "主日";
          let name = nameIdx > -1 ? String(row[nameIdx]).trim() : "";

          meetingMap[row[idIdx]] = { dateNum: dNum, dateStr: dateStr, name: name, type: type };

          // 若落在查詢季度內，則建立外部基礎框架
          if (dNum >= startDateNum && dNum <= endDateNum) {
            externalMeetings.push({ '年度': reqYear, '季度': reqQuarter, '日期': dateStr, '聚會名稱': name, '聚會類別': type });
          }
        }
      });
    }

    // 2. 抓取外部講道資訊
    // 🔧 修正：改用「日期_聚會名稱」複合 key
    const sermonSheet = extSs.getSheetByName('講道資訊');
    if (sermonSheet) {
      const sData = sermonSheet.getDataRange().getValues();
      const sHeaders = sData[0];
      const mIdIdx = sHeaders.indexOf('聚會ID');
      const typeIdx = sHeaders.indexOf('講道類別');
      const speakerIdx = sHeaders.indexOf('講員');
      const titleIdx = sHeaders.indexOf('講題');
      const scriptIdx = sHeaders.indexOf('經文');

      sData.slice(1).forEach(row => {
        if (mIdIdx > -1 && row[mIdIdx] && meetingMap[row[mIdIdx]]) {
          const mInfo = meetingMap[row[mIdIdx]];
          // 🔧 修正：key 改為「日期_聚會名稱」
          externalSermonDict[`${mInfo.dateNum}_${mInfo.name}`] = {
            '聚會名稱': mInfo.name, '牧師': speakerIdx > -1 ? row[speakerIdx] : '',
            '題目': titleIdx > -1 ? row[titleIdx] : '', '經文': scriptIdx > -1 ? row[scriptIdx] : ''
          };
        }
      });
    }
  } catch (e) { Logger.log("讀取外部資料失敗：" + e.toString()); }

  // 3. 取得本地已排班資料
  const localData = getScheduleData(reqYear, reqQuarter);

  // 🔧 修正：localDict 改用「日期_聚會名稱」複合 key，避免同日期相同聚會類別互蓋
  let localDict = {};
  localData.forEach(row => {
    let key = `${toDateNum(row['日期'])}_${String(row['聚會名稱']).trim()}`;
    localDict[key] = row;
  });

  let finalSchedule = [];
  let processedKeys = new Set();

  // 4. 合併：以外部框架為底，填入本地排班與外部講道
  // 🔧 修正：key 統一改為「日期_聚會名稱」，嚴格比對，找不到給空物件
  externalMeetings.forEach(extRow => {
    let key = `${toDateNum(extRow['日期'])}_${String(extRow['聚會名稱']).trim()}`;
    processedKeys.add(key);

    // ✅ 嚴格比對，找不到就給空物件，不再做模糊 fallback
    let localRow = localDict[key] || {};
    let sermonInfo = externalSermonDict[key] || {};

    finalSchedule.push({
      ...extRow, ...localRow,
      '聚會名稱': sermonInfo['聚會名稱'] || extRow['聚會名稱'],
      '牧師': sermonInfo['牧師'] || '',
      '題目': sermonInfo['題目'] || '',
      '經文': sermonInfo['經文'] || ''
    });
  });

  // 5. 將本地額外新增的聚會也塞入清單
  localData.forEach(row => {
    let key = `${toDateNum(row['日期'])}_${String(row['聚會名稱']).trim()}`;
    if (!processedKeys.has(key)) {
      let sermonInfo = externalSermonDict[key] || {};
      finalSchedule.push({
        ...row,
        '聚會名稱': sermonInfo['聚會名稱'] || row['聚會名稱'] || '',
        '牧師': sermonInfo['牧師'] || '',
        '題目': sermonInfo['題目'] || '',
        '經文': sermonInfo['經文'] || ''
      });
    }
  });

  finalSchedule.sort((a, b) => toDateNum(a['日期']) - toDateNum(b['日期']));

  // 6. 🎵 合併敬拜曲目
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const songsSheet = ss.getSheetByName('敬拜曲目');
    if (songsSheet) {
      const sData = songsSheet.getDataRange().getValues();
      if (sData.length > 1) {
        const sHeaders = sData[0].map(h => String(h).trim());
        const sdIdx = sHeaders.indexOf('日期');
        const snIdx = sHeaders.indexOf('聚會名稱');
        const ssIdx = sHeaders.indexOf('敬拜曲目');
        let songsDict = {};
        sData.slice(1).forEach(row => {
          const dNum = toDateNum(row[sdIdx]);
          const mName = String(row[snIdx] || '').trim();
          if (dNum > 0) songsDict[`${dNum}_${mName}`] = String(row[ssIdx] || '').trim();
        });
        finalSchedule.forEach(row => {
          const key = `${toDateNum(row['日期'])}_${String(row['聚會名稱'] || '').trim()}`;
          row['敬拜曲目'] = songsDict[key] || '';
        });
      }
    }
  } catch(e) { Logger.log("讀取曲目資料失敗：" + e.toString()); }

  return finalSchedule;
}

// 下方保留您原本的 getScheduleByDateRange, getScheduleData, saveScheduleData
function getScheduleByDateRange(payload) {
  const startDate = payload.startDate;
  const endDate = payload.endDate;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('服事表總表');
  if (!sheet) return { status: 'error', message: '找不到服事表總表' };
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { status: 'success', data: [] };
  const headers = data[0].map(h => String(h).trim());
  const dateIndex = headers.indexOf('日期');
  const startNum = toDateNum(startDate), endNum = toDateNum(endDate);
  const result = data.slice(1).filter(row => {
    const rowDateNum = toDateNum(row[dateIndex]);
    return rowDateNum >= startNum && rowDateNum <= endNum;
  }).map(row => {
    let obj = {};
    headers.forEach((h, i) => {
      let val = row[i];
      if (val instanceof Date) val = Utilities.formatDate(val, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
      obj[h] = val;
    });
    return obj;
  });
  return { status: 'success', data: result };
}

function getScheduleData(year, quarter) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('服事表總表');
  if (!sheet) return []; 
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; 
  const headers = data[0];
  const yearIndex = headers.indexOf('年度'), quarterIndex = headers.indexOf('季度'), dateIndex = headers.indexOf('日期');
  
  const filteredRows = data.slice(1).filter(row => String(row[yearIndex]).trim() === String(year).trim() && String(row[quarterIndex]).trim() === String(quarter).trim());
  let result = filteredRows.map(row => {
    let rowData = {};
    headers.forEach((header, index) => {
      let val = row[index];
      if (val instanceof Date) val = Utilities.formatDate(val, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
      rowData[header] = val;
    });
    return rowData;
  });
  
  result.sort((a, b) => toDateNum(a['日期']) - toDateNum(b['日期']));
  const posHeaders = headers.filter(h => !['年度', '季度', '日期', '聚會類別', '聚會名稱'].includes(h));
  let prevLeader = null;
  result.forEach(row => {
    let warnings = [];
    let curLeader = row['主領'];
    if (curLeader && curLeader === prevLeader) warnings.push(`主領 [${curLeader}] 連續服事`);
    let names = [];
    posHeaders.forEach(h => {
      let p = String(row[h]).trim();
      if (p && p !== "" && p !== "【待定】") {
        if (names.includes(p)) warnings.push(`[${p}] 重複`);
        else names.push(p);
      }
    });
    row.hasWarning = warnings.length > 0;
    row.warningMessage = row.hasWarning ? "⚠️ " + warnings.join(" | ") : "";
    prevLeader = curLeader;
  });
  return result;
}

function saveScheduleData(scheduleData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('服事表總表') || ss.insertSheet('服事表總表');
  if (!scheduleData.length) return { status: 'success' };
  let dateNums = scheduleData.map(d => toDateNum(d['日期'])).sort((a,b) => a-b);
  const minNum = dateNums[0], maxNum = dateNums[dateNums.length - 1];
  const cutoffDate = new Date();
  cutoffDate.setFullYear(cutoffDate.getFullYear() - 1);
  const cutoffNum = toDateNum(cutoffDate);
  const oldRaw = sheet.getDataRange().getValues();
  let headers = oldRaw.length > 0 ? oldRaw[0] : ['年度', '季度', '日期', '聚會名稱', '聚會類別'];
  let hSet = new Set(headers);
  scheduleData.forEach(r => Object.keys(r).forEach(k => hSet.add(k)));
  let finalHeaders = ['年度', '季度', '日期', '聚會名稱', '聚會類別', ...Array.from(hSet).filter(h => !['年度', '季度', '日期', '聚會名稱', '聚會類別', 'leaves'].includes(h))];
  
  let retained = [];
  if (oldRaw.length > 1) {
    const dIdx = headers.indexOf('日期');
    oldRaw.slice(1).forEach(row => {
      const dNum = toDateNum(row[dIdx]);
      if (dNum >= cutoffNum && (dNum < minNum || dNum > maxNum)) {
        let obj = {};
        headers.forEach((h, i) => obj[h] = row[i]);
        retained.push(obj);
      }
    });
  }
  let finalData = retained.concat(scheduleData).sort((a,b) => toDateNum(a['日期']) - toDateNum(b['日期']));
  let finalRows = finalData.map(obj => finalHeaders.map(h => obj[h] || ''));
  sheet.clearContents();
  sheet.appendRow(finalHeaders);
  if (finalRows.length > 0) sheet.getRange(2, 1, finalRows.length, finalHeaders.length).setValues(finalRows);
  return { status: "success", message: "發佈成功！" };
}