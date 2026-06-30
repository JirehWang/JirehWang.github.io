// ==========================================
// WorshipSchedule.js (服事表核心邏輯：外部框架驅動版)
// ==========================================

function worship_toDateNum(dateObj) {
  if (!dateObj) return 0;
  if (dateObj instanceof Date) {
    if (isNaN(dateObj.getTime())) return 0;
    // 改用快取時區，避免每筆 Date 都觸發一次 openById + getSpreadsheetTimeZone
    return parseInt(Utilities.formatDate(dateObj, _getTz(), "yyyyMMdd"), 10);
  }
  const str = String(dateObj).trim();
  const parts = str.split('-');
  if (parts.length === 3) return parseInt(parts[0] + parts[1].padStart(2, '0') + parts[2].padStart(2, '0'), 10);
  return 0;
}

// 把任意輸入轉成「YYYY-MM-DD」字串
function worship_formatDateStr(d) {
  if (!d) return '';
  if (d instanceof Date) return Utilities.formatDate(d, _getTz(), 'yyyy-MM-dd');
  const s = String(d).trim();
  return /^\d{4}-\d{2}-\d{2}/.test(s) ? s.substring(0, 10) : s;
}

/**
 * 🌟 核心升級：以「外部聚會資料」為主體框架，疊加本地排班
 * 🔧 修正：改用「日期_聚會名稱」複合 key，避免同日期多個相同聚會類別造成資料覆蓋
 */
function getMergedSchedule(year, quarter) {
  const reqYear = String(year).trim();
  const reqQuarter = String(quarter).trim();

  // 1. 本地服事表（同工排班的核心來源）
  let localData = getScheduleData(reqYear, reqQuarter);

  // 🌟 如果本地服事表為空，則自動從「教會行事曆」的「講道資訊」或「聚會名稱」事項產生該季度的初始排班框架
  if (localData.length === 0) {
    try {
      const cfg = getCalendarLinkConfig();
      if (cfg.calendarReachable && cfg.sermonSubTypes.length > 0) {
        // 算出該季度的日期區間
        let startMonth, endMonth, endDay;
        if (reqQuarter === 'Q1') { startMonth = '01'; endMonth = '03'; endDay = '31'; }
        else if (reqQuarter === 'Q2') { startMonth = '04'; endMonth = '06'; endDay = '30'; }
        else if (reqQuarter === 'Q3') { startMonth = '07'; endMonth = '09'; endDay = '30'; }
        else if (reqQuarter === 'Q4') { startMonth = '10'; endMonth = '12'; endDay = '31'; }
        
        if (startMonth) {
          const startDate = `${reqYear}-${startMonth}-01`;
          const endDate = `${reqYear}-${endMonth}-${endDay}`;
          
          // 讀取行事曆事項與類型
          const events = _readCalendarSheet('事項') || [];
          const types = _readCalendarSheet('事項類型') || [];
          
          const subTypeNameById = {};
          (cfg.sermonSubTypes || []).forEach(t => { subTypeNameById[t.typeId] = t.name; });
          const sermonSubIds = new Set((cfg.sermonSubTypes || []).map(t => t.typeId));
          
          const meetingNameType = types.find(t => t['名稱'] === '聚會名稱');
          const meetingNameTypeId = meetingNameType ? meetingNameType.typeId : '';
          
          // 取得位置清單以利初始化欄位
          const positions = getPositions() || [];
          
          // 按日期分組，避免同日期因多個事件產生重複列
          const eventsByDateMap = {};
          events.forEach(e => {
            const d = worship_formatDateStr(e['日期']);
            if (d && d >= startDate && d <= endDate) {
              if (!eventsByDateMap[d]) eventsByDateMap[d] = [];
              eventsByDateMap[d].push(e);
            }
          });
          
          const initRows = [];
          Object.entries(eventsByDateMap).forEach(([d, dayEvts]) => {
            const sermonEvt = dayEvts.find(e => sermonSubIds.has(e.typeId)) || null;
            const namedEvt = dayEvts.find(e => e.typeId === meetingNameTypeId) || null;
            
            if (!sermonEvt && !namedEvt) return;
            
            let meetingName = '';
            if (namedEvt && namedEvt['顯示標題']) {
              meetingName = String(namedEvt['顯示標題']).trim();
            } else if (sermonEvt && sermonEvt['顯示標題']) {
              meetingName = String(sermonEvt['顯示標題']).trim();
            }
            
            let typeName = '';
            if (sermonEvt) {
              typeName = subTypeNameById[sermonEvt.typeId] || '主日';
            } else {
              const defaultSubId = (cfg.defaultSermonSubTypeId || '').trim();
              typeName = subTypeNameById[defaultSubId] || '';
            }
            
            const row = {
              '年度': reqYear,
              '季度': reqQuarter,
              '日期': d,
              '聚會名稱': meetingName,
              '聚會類別': typeName,
              '牧師': '',
              '題目': '',
              '經文': '',
              '敬拜曲目': '',
              'leaves': []
            };
            
            positions.forEach(pos => {
              row[pos.positionName] = pos.isRequired === '是' ? '【待定】' : '';
            });
            
            initRows.push(row);
          });
          
          localData = initRows;
          localData.sort((a, b) => worship_toDateNum(a['日期']) - worship_toDateNum(b['日期']));
        }
      }
    } catch (e) {
      Logger.log('從行事曆初始化季度框架失敗：' + e.toString());
    }
  }

  let finalSchedule = localData.map(row => {
    return Object.assign({}, row, {
      '年度': row['年度'] || reqYear,
      '季度': row['季度'] || reqQuarter
    });
  });

  // 2. 從新版行事曆 schema 抓資料（透過 getCalendarDataForDates，含 4h 快取）
  //    ⚠️ 同日期可能有多筆不同聚會名稱 → 用 (日期, 聚會名稱) 組合查詢，避免互相覆蓋
  try {
    const entries = [];
    finalSchedule.forEach(row => {
      const d = worship_formatDateStr(row['日期']);
      if (!d) return;
      entries.push({ date: d, meetingName: String(row['聚會名稱'] || '').trim() });
    });

    if (entries.length > 0 && typeof getCalendarDataForDates === 'function') {
      const calData = getCalendarDataForDates({ entries: entries });
      const cfg = getCalendarLinkConfig();
      const subTypeNameById = {};
      (cfg.sermonSubTypes || []).forEach(t => { subTypeNameById[t.typeId] = t.name; });

      finalSchedule = finalSchedule.map(row => {
        const d = worship_formatDateStr(row['日期']);
        const meetingName = String(row['聚會名稱'] || '').trim();
        // 優先用「日期|聚會名稱」key，找不到再 fallback 用 date
        const cd = calData[meetingName ? `${d}|${meetingName}` : d] || calData[d] || {};

        // 「聚會名稱」：行事曆有對應 row 名稱的事項 → 用行事曆標題
        if (cd.namedEvent && cd.namedEvent.title) {
          row['聚會名稱'] = String(cd.namedEvent.title).trim();
        }

        // 「聚會類別」：優先使用行事曆講道事項的類型名稱（如「華語」、「聯合-華語」、「聯合-台語」）
        if (cd.sermon && cd.sermon.typeName) {
          row['聚會類別'] = String(cd.sermon.typeName).trim();
        } else {
          // fallback: 沒有講道事項時，採用設定的子類型名稱，最後保留原值
          const effSubName = subTypeNameById[cd.effectiveSermonSubTypeId];
          if (effSubName) {
            row['聚會類別'] = effSubName;
          } else if (!row['聚會類別']) {
            row['聚會類別'] = '主日';
          }
        }

        // 講道資訊 → 帶入「牧師、題目、經文」（覆蓋已存在欄位）
        if (cd.sermon && cd.sermon.values) {
          const v = cd.sermon.values;
          if (v['講員']) row['牧師'] = String(v['講員']);
          if (v['講題']) row['題目'] = String(v['講題']);
          if (v['經文']) row['經文'] = String(v['經文']);
          if (v['詩歌'] && !row['敬拜曲目']) row['敬拜曲目'] = String(v['詩歌']);
        } else {
          if (row['牧師'] === undefined) row['牧師'] = '';
          if (row['題目'] === undefined) row['題目'] = '';
          if (row['經文'] === undefined) row['經文'] = '';
        }
        return row;
      });
    }
  } catch (e) {
    Logger.log('整合行事曆資料失敗：' + e.toString());
  }

  finalSchedule.sort((a, b) => worship_toDateNum(a['日期']) - worship_toDateNum(b['日期']));

  // 6. 🎵 合併敬拜曲目
  try {
    const ss = getWorshipSS();
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
          const dNum = worship_toDateNum(row[sdIdx]);
          const mName = String(row[snIdx] || '').trim();
          if (dNum > 0) songsDict[`${dNum}_${mName}`] = String(row[ssIdx] || '').trim();
        });
        finalSchedule.forEach(row => {
          const key = `${worship_toDateNum(row['日期'])}_${String(row['聚會名稱'] || '').trim()}`;
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
  const ss = getWorshipSS();
  const sheet = ss.getSheetByName('服事表總表');
  if (!sheet) return { status: 'error', message: '找不到服事表總表' };
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { status: 'success', data: [] };
  const headers = data[0].map(h => String(h).trim());
  const dateIndex = headers.indexOf('日期');
  const startNum = worship_toDateNum(startDate), endNum = worship_toDateNum(endDate);
  const tz = _getTz(); // 提到迴圈外，避免每個欄位都呼叫一次
  const result = data.slice(1).filter(row => {
    const rowDateNum = worship_toDateNum(row[dateIndex]);
    return rowDateNum >= startNum && rowDateNum <= endNum;
  }).map(row => {
    let obj = {};
    headers.forEach((h, i) => {
      if (!h) return; // 跳過空標題，避免產生空 key（會炸 Firebase RTDB cache）
      let val = row[i];
      if (val instanceof Date) val = Utilities.formatDate(val, tz, "yyyy-MM-dd");
      obj[h] = val;
    });
    return obj;
  });
  return { status: 'success', data: result };
}

function getScheduleData(year, quarter) {
  const ss = getWorshipSS();
  const sheet = ss.getSheetByName('服事表總表');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  const yearIndex = headers.indexOf('年度'), quarterIndex = headers.indexOf('季度'), dateIndex = headers.indexOf('日期');

  const filteredRows = data.slice(1).filter(row => String(row[yearIndex]).trim() === String(year).trim() && String(row[quarterIndex]).trim() === String(quarter).trim());
  const tz = _getTz(); // 提到迴圈外，避免每個欄位都呼叫一次
  let result = filteredRows.map(row => {
    let rowData = {};
    headers.forEach((header, index) => {
      if (!header) return; // 跳過空標題，避免空 key 炸 Firebase RTDB cache
      let val = row[index];
      if (val instanceof Date) val = Utilities.formatDate(val, tz, "yyyy-MM-dd");
      rowData[header] = val;
    });
    return rowData;
  });
  
  result.sort((a, b) => worship_toDateNum(a['日期']) - worship_toDateNum(b['日期']));
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
  const ss = getWorshipSS();
  let sheet = ss.getSheetByName('服事表總表') || ss.insertSheet('服事表總表');
  if (!scheduleData.length) return { status: 'success' };
  let dateNums = scheduleData.map(d => worship_toDateNum(d['日期'])).sort((a,b) => a-b);
  const minNum = dateNums[0], maxNum = dateNums[dateNums.length - 1];
  const cutoffDate = new Date();
  cutoffDate.setFullYear(cutoffDate.getFullYear() - 1);
  const cutoffNum = worship_toDateNum(cutoffDate);
  const oldRaw = sheet.getDataRange().getValues();
  let headers = oldRaw.length > 0 ? oldRaw[0] : ['年度', '季度', '日期', '聚會名稱', '聚會類別'];
  let hSet = new Set(headers);
  scheduleData.forEach(r => Object.keys(r).forEach(k => hSet.add(k)));
  let finalHeaders = ['年度', '季度', '日期', '聚會名稱', '聚會類別', ...Array.from(hSet).filter(h => !['年度', '季度', '日期', '聚會名稱', '聚會類別', 'leaves'].includes(h))];
  
  let retained = [];
  if (oldRaw.length > 1) {
    const dIdx = headers.indexOf('日期');
    oldRaw.slice(1).forEach(row => {
      const dNum = worship_toDateNum(row[dIdx]);
      if (dNum >= cutoffNum && (dNum < minNum || dNum > maxNum)) {
        let obj = {};
        headers.forEach((h, i) => obj[h] = row[i]);
        retained.push(obj);
      }
    });
  }
  let finalData = retained.concat(scheduleData).sort((a,b) => worship_toDateNum(a['日期']) - worship_toDateNum(b['日期']));
  let finalRows = finalData.map(obj => finalHeaders.map(h => obj[h] || ''));
  sheet.clearContents();
  sheet.appendRow(finalHeaders);
  if (finalRows.length > 0) sheet.getRange(2, 1, finalRows.length, finalHeaders.length).setValues(finalRows);
  return { status: "success", message: "發佈成功！" };
}

// ==========================================
// 測試函數：驗證 Q3 空季度初始化邏輯
// ==========================================
function testMergedScheduleQ3() {
  const res = getMergedSchedule("2026", "Q3");
  Logger.log("Q3 初始行數：" + res.length);
  if (res.length > 0) {
    Logger.log("第一列資料日期：" + res[0]['日期'] + "，名稱：" + res[0]['聚會名稱'] + "，類別：" + res[0]['聚會類別']);
  }
}

