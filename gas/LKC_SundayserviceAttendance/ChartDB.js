/**
 * 1. 取得單一群組圖表資料
 */
function getDailyChartData(type) {
  const ss = getSS();
  const sheet = ss.getSheetByName(type + "點名紀錄");
  if (!sheet) return { error: "⚠️ 找不到 [" + type + "] 的紀錄表，請先進行幾次點名！" };
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { error: "📊 目前資料量不足，無法繪製圖表！" };

  let chartRows = [];
  for (let i = 1; i < data.length; i++) {
    const dateVal = data[i][0];
    if (!dateVal) continue;
    let d = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
    if (isNaN(d.getTime())) continue;
    const dateKey = Utilities.formatDate(d, "GMT+8", "MM/dd");
    const namesStr = data[i][1] ? data[i][1].toString() : "";
    const membersCount = namesStr ? namesStr.split(/[,，、]\s*/).filter(n => n.trim() !== "").length : 0;
    const nfMale = Number(data[i][2]) || 0;
    const nfFemale = Number(data[i][3]) || 0;
    chartRows.push({
      date: dateKey, fullDate: d.getTime(),
      members: membersCount, newFriends: nfMale + nfFemale,
      total: membersCount + nfMale + nfFemale
    });
  }
  chartRows.sort((a, b) => a.fullDate - b.fullDate);
  const recentRows = chartRows.slice(-15);
  return {
    labels: recentRows.map(r => r.date),
    members: recentRows.map(r => r.members),
    newFriends: recentRows.map(r => r.newFriends),
    totals: recentRows.map(r => r.total)
  };
}

/**
 * 2. 取得跨群組分類圖表資料（支援堆疊圖與日期篩選）
 */
function getCategoryChartData(category, startDateStr, endDateStr) {
  const ss = getSS();
  const config = getGroupConfig();
  const groups = config[category] || [];
  if (groups.length === 0) return { error: "⚠️ 找不到 [" + category + "] 類別下的任何群組設定！" };

  const filterStartTime = startDateStr ? new Date(startDateStr.replace(/-/g, '/')).setHours(0, 0, 0, 0) : null;
  const filterEndTime = endDateStr ? new Date(endDateStr.replace(/-/g, '/')).setHours(23, 59, 59, 999) : null;

  const dateMap = {};

  groups.forEach(group => {
    const sheet = ss.getSheetByName(group + "點名紀錄");
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const dateVal = data[i][0];
      if (!dateVal) continue;
      let d = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
      if (isNaN(d.getTime())) continue;
      const timestamp = d.getTime();
      if (filterStartTime && timestamp < filterStartTime) continue;
      if (filterEndTime && timestamp > filterEndTime) continue;
      const dateKey = Utilities.formatDate(d, "GMT+8", "MM/dd");
      const namesStr = data[i][1] ? data[i][1].toString() : "";
      const membersCount = namesStr ? namesStr.split(/[,，、]\s*/).filter(n => n.trim() !== "").length : 0;
      const nfMale = Number(data[i][2]) || 0;
      const nfFemale = Number(data[i][3]) || 0;
      const newFriendsCount = nfMale + nfFemale;
      if (!dateMap[dateKey]) {
        dateMap[dateKey] = { timestamp, catMembers: 0, catNewFriends: 0 };
        groups.forEach(g => dateMap[dateKey][g] = 0);
      }
      dateMap[dateKey].catMembers += membersCount;
      dateMap[dateKey].catNewFriends += newFriendsCount;
      dateMap[dateKey][group] += (membersCount + newFriendsCount);
    }
  });

  const sortedDates = Object.keys(dateMap).sort((a, b) => dateMap[a].timestamp - dateMap[b].timestamp);
  const hasCustomDate = !!(startDateStr || endDateStr);
  const targetDates = hasCustomDate ? sortedDates : sortedDates.slice(-15);
  if (targetDates.length === 0) return { error: "⚠️ 該日期區間內沒有找到任何點名紀錄喔！" };

  const resultDatasets = { "MembersTotal": [], "NewFriendsTotal": [] };
  groups.forEach(g => resultDatasets[g] = []);
  targetDates.forEach(date => {
    resultDatasets["MembersTotal"].push(dateMap[date].catMembers);
    resultDatasets["NewFriendsTotal"].push(dateMap[date].catNewFriends);
    groups.forEach(g => resultDatasets[g].push(dateMap[date][g] || 0));
  });

  return { labels: targetDates, datasets: resultDatasets, groups, isCustomRange: hasCustomDate };
}