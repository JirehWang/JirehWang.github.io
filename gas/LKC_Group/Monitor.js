/**
 * 資源監控模組 - 自動紀錄至工作表
 */
const QUOTA_LIMIT_SEC = 5400; // 90分鐘 = 5400秒
const MONITOR_SHEET_NAME = "系統監控";

/**
 * 紀錄單次執行的時間 (此函式仍需放在 doGet 中)
 */
function recordUsage(startTime) {
  const props = PropertiesService.getScriptProperties();
  const now = new Date();
  const todayKey = "USAGE_DATE_" + now.toLocaleDateString();
  const usageKey = "USAGE_TIME_SEC";

  // 檢查是否換天，換天就重置
  const lastDate = props.getProperty("LAST_MONITOR_DATE");
  let currentUsage = 0;
  
  if (lastDate !== todayKey) {
    props.setProperty("LAST_MONITOR_DATE", todayKey);
    props.setProperty(usageKey, "0");
  } else {
    currentUsage = parseFloat(props.getProperty(usageKey)) || 0;
  }

  const duration = (new Date().getTime() - startTime) / 1000;
  props.setProperty(usageKey, (currentUsage + duration).toFixed(2));
}

/**
 * 每小時執行一次：將當天累計數據更新至工作表
 */
function updateDailyUsageLog() {
  const ss = getSS();
  const sheet = ss.getSheetByName(MONITOR_SHEET_NAME);
  const props = PropertiesService.getScriptProperties();
  
  const now = new Date();
  const todayStr = Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd");
  const usedSec = parseFloat(props.getProperty("USAGE_TIME_SEC")) || 0;
  const percent = ((usedSec / QUOTA_LIMIT_SEC) * 100).toFixed(2) + "%";
  const lastUpdate = Utilities.formatDate(now, "GMT+8", "HH:mm:ss");

  const data = sheet.getDataRange().getValues();
  let targetRow = -1;

  // 尋找是否已有今天的紀錄
  for (let i = 1; i < data.length; i++) {
    const rowDate = data[i][0] instanceof Date ? 
                    Utilities.formatDate(data[i][0], "GMT+8", "yyyy-MM-dd") : data[i][0];
    if (rowDate === todayStr) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow !== -1) {
    // 更新今日現有列
    sheet.getRange(targetRow, 2, 1, 3).setValues([[usedSec.toFixed(1), percent, lastUpdate]]);
  } else {
    // 新增今日新列
    sheet.appendRow([todayStr, usedSec.toFixed(1), percent, lastUpdate]);
  }
}