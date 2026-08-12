/**
 * 取得會友編號→姓名對照表（有快取，速度很快）
 */
function getMemberMap() {
  const cache = CacheService.getScriptCache();
  const cacheKey = "MEMBER_DATA_MAP";
  const cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }
  const ss = getSS();
  const memberSheet = ss.getSheetByName('會友名單');
  const data = memberSheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const name = data[i][0];   // A 欄：姓名
    const qrCode = data[i][7]; // H 欄：系統編號
    if (qrCode) map[qrCode.toString().trim()] = name;
  }
  cache.put(cacheKey, JSON.stringify(map), 600); // 快取 10 分鐘
  return map;
}

/**
 * 處理 QR 掃描請求
 */
function handleQrScanRequest(e) {
  const qrContent = (e.parameter.name || "").toString().trim();
  const userId = e.parameter.userId;

  const memberMap = getMemberMap();
  const foundName = memberMap[qrContent];

  if (!foundName) {
    return { status: "error", message: "找不到編號 " + qrContent + " 對應的姓名" };
  }

  const props = PropertiesService.getScriptProperties();
  const activeMode = props.getProperty('MODE_' + userId) || '台語';

  try {
    syncClickToServer(foundName, true, activeMode, userId);
    return { status: "success", name: foundName, mode: activeMode };
  } catch (err) {
    return { status: "error", message: "勾選失敗：" + err.toString() };
  }
}

/**
 * 更新裝置的點名模式
 */
function updateDeviceMode(userId, mode) {
  PropertiesService.getScriptProperties().setProperty('MODE_' + userId, mode);
}