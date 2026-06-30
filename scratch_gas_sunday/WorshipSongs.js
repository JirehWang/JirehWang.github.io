// ==========================================
// WorshipSongs.js - 敬拜曲目管理
// ==========================================

/**
 * 讀取曲目資料
 * payload: { startDate, endDate } 或 { year, quarter }
 */
function getSongs(payload) {
  const ss = getWorshipSS();
  const sheet = ss.getSheetByName('敬拜曲目');
  
  // 建立本地曲目字典：key = 日期_聚會名稱
  let localSongsDict = {};
  if (sheet) {
    const data = sheet.getDataRange().getValues();
    if (data.length > 1) {
      const headers = data[0].map(h => String(h).trim());
      const dateIdx = headers.indexOf('日期');
      const nameIdx = headers.indexOf('聚會名稱');
      const songsIdx = headers.indexOf('敬拜曲目');
      data.slice(1).forEach(row => {
        const dNum = worship_toDateNum(row[dateIdx]);
        const mName = String(row[nameIdx] || '').trim();
        if (dNum > 0) {
          localSongsDict[`${dNum}_${mName}`] = String(row[songsIdx] || '').trim();
        }
      });
    }
  }

  // 取得聚會框架（帶主領）
  let scheduleRows = [];
  if (payload.year && payload.quarter) {
    scheduleRows = getMergedSchedule(payload.year, payload.quarter);
  } else if (payload.startDate && payload.endDate) {
    const result = getScheduleByDateRange(payload);
    scheduleRows = result.data || [];
  }

  // 組合結果
  const result = scheduleRows.map(row => {
    const dNum = worship_toDateNum(row['日期']);
    const mName = String(row['聚會名稱'] || '').trim();
    const key = `${dNum}_${mName}`;
    return {
      '日期': row['日期'] || '',
      '聚會名稱': row['聚會名稱'] || '',
      '聚會類別': row['聚會類別'] || '',
      '主領': row['主領'] || '',
      '敬拜曲目': localSongsDict[key] || ''
    };
  });

  return { status: 'success', data: result };
}

/**
 * 儲存曲目資料
 * payload: { songsData: [{ 日期, 聚會名稱, 聚會類別, 主領, 敬拜曲目 }] }
 */
function saveSongs(payload) {
  const ss = getWorshipSS();
  let sheet = ss.getSheetByName('敬拜曲目');
  if (!sheet) sheet = ss.insertSheet('敬拜曲目');

  const songsData = payload.songsData || [];
  if (!songsData.length) return { status: 'success', message: '無資料可儲存' };

  const headers = ['日期', '聚會名稱', '聚會類別', '主領', '敬拜曲目'];
  const tz = ss.getSpreadsheetTimeZone(); // 提到迴圈外，下方多處重用

  // 讀取現有資料，建立字典
  const oldData = sheet.getDataRange().getValues();
  let existingDict = {};
  if (oldData.length > 1) {
    const oldHeaders = oldData[0].map(h => String(h).trim());
    const dIdx = oldHeaders.indexOf('日期');
    const nIdx = oldHeaders.indexOf('聚會名稱');
    const sIdx = oldHeaders.indexOf('敬拜曲目');
    const tIdx = oldHeaders.indexOf('聚會類別');
    const lIdx = oldHeaders.indexOf('主領');
    oldData.slice(1).forEach(row => {
      const dNum = worship_toDateNum(row[dIdx]);
      const mName = String(row[nIdx] || '').trim();
      if (dNum > 0) {
        existingDict[`${dNum}_${mName}`] = {
          '日期': row[dIdx] instanceof Date
            ? Utilities.formatDate(row[dIdx], tz, "yyyy-MM-dd")
            : String(row[dIdx]).trim(),
          '聚會名稱': mName,
          '聚會類別': String(row[tIdx] || '').trim(),
          '主領': String(row[lIdx] || '').trim(),
          '敬拜曲目': String(row[sIdx] || '').trim()
        };
      }
    });
  }

  // 用新資料覆蓋對應 key
  songsData.forEach(item => {
    const dNum = worship_toDateNum(item['日期']);
    const mName = String(item['聚會名稱'] || '').trim();
    const key = `${dNum}_${mName}`;
    // 將任意分隔符號統一轉成「、」
    const rawSongs = String(item['敬拜曲目'] || '').trim();
    const normalizedSongs = rawSongs
      ? rawSongs.split(/[,，、\/\\|；;]+/).map(s => s.trim()).filter(s => s).join('、')
      : '';
    existingDict[key] = {
      '日期': item['日期'],
      '聚會名稱': mName,
      '聚會類別': String(item['聚會類別'] || '').trim(),
      '主領': String(item['主領'] || '').trim(),
      '敬拜曲目': normalizedSongs
    };
  });

  // 清除過期資料（1年前）
  const cutoffDate = new Date();
  cutoffDate.setFullYear(cutoffDate.getFullYear() - 1);
  const cutoffStr = Utilities.formatDate(cutoffDate, tz, "yyyy-MM-dd");

  let finalRows = [];
  Object.values(existingDict).forEach(obj => {
    let dateStr = obj['日期'];
    if (dateStr instanceof Date) dateStr = Utilities.formatDate(dateStr, tz, "yyyy-MM-dd");
    if (String(dateStr).trim() >= cutoffStr) finalRows.push(obj);
  });

  finalRows.sort((a, b) => worship_toDateNum(a['日期']) - worship_toDateNum(b['日期']));

  sheet.clearContents();
  sheet.appendRow(headers);
  if (finalRows.length > 0) {
    const rows = finalRows.map(obj => headers.map(h => obj[h] || ''));
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  return { status: 'success', message: '曲目資料已成功儲存！' };
}
