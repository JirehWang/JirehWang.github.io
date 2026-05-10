/**
 * 全域變數設定
 * ⚠️ 請把下面的 ID 換成你自己的試算表 ID
 */
const SPREADSHEET_ID = '1tS_RSufp4Y1F1G84oUxzXCp4g5S5bg8dODzGwi3O_wE';
const MEMBER_SHEET = '會友名單';

/**
 * 統一的 doPost 入口點
 * GitHub Pages 前端發送的所有請求都經過這裡
 */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const payload = body.payload;

    let result;

    switch (action) {
      // ===== 會友管理 =====
      case 'getAllMembers':
        result = getAllMembers(); break;
      case 'updateMember':
        result = updateMember(payload.oldName || payload[0], payload.newData || payload[1]); break;
      case 'deleteMember':
        result = deleteMember(payload.name || payload); break;
      case 'addMember':
        result = addMember(payload); break;
      case 'previewMemberCard':
        result = previewMemberCard(payload); break;
      case 'generateMemberCard':
        result = generateMemberCard(payload); break;

      // ===== 點名系統 =====
      case 'getGroupConfig':
        result = getGroupConfig(); break;
      case 'getSmartAttendanceList':
        result = getSmartAttendanceList(payload.type || payload[0], payload.userId || payload[1]); break;
      case 'syncClickToServer':
        result = syncClickToServer(payload.name || payload[0], payload.isChecked || payload[1], payload.type || payload[2], payload.userId || payload[3]); break;
      case 'saveAttendance':
        result = saveAttendance(payload.date || payload[0], payload.presentList || payload[1], payload.type || payload[2], payload.nfMale || payload[3], payload.nfFemale || payload[4]); break;
      case 'revokeAttendance':
        result = revokeAttendance(payload.name || payload[0], payload.type || payload[1], payload.userId || payload[2]); break;
      case 'createAttendanceGroup':
        result = createAttendanceGroup(payload.category || payload[0], payload.groupName || payload[1]); break;
      case 'updateDeviceMode':
        result = updateDeviceMode(payload.userId || payload[0], payload.mode || payload[1]); break;
      case 'getQuickSyncData':
        result = getQuickSyncData(payload.type || payload[0], payload.userId || payload[1]); break;

      // ===== 統計查詢 =====
      case 'getAttendanceStats':
        result = getAttendanceStats(payload); break;
      case 'getAttendanceTrend':
        result = getAttendanceTrend(payload); break;

      // ===== 趨勢分析 =====
      case 'getCategoryChartData':
        result = getCategoryChartData(payload.category || payload[0], payload.startDate || payload[1], payload.endDate || payload[2]); break;

      default:
        throw new Error('未知操作：' + action);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ data: result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * doGet：處理 QR 掃描請求（保留原有功能）
 */
function doGet(e) {
  try {
    const action = e.parameter.action;
    const cat = e.parameter.cat;
    const grp = e.parameter.grp;

    // 1. 處理 QR 掃描點名請求
    if (action === 'syncClickToServer') {
      const result = handleQrScanRequest(e);
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 2. 處理場次跳轉請求（手機掃描場次 QR Code）
    if (cat && grp) {
      const githubUrl = "https://jirehwang.github.io/LKC_-SundayserviceAttendance.github.io/";
      const redirectUrl = githubUrl + "?cat=" + encodeURIComponent(cat) + "&grp=" + encodeURIComponent(grp);
      
      return HtmlService.createHtmlOutput(
        '<meta http-equiv="refresh" content="0; URL=' + redirectUrl + '">'
      )
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    // 3. 其他請求
    return ContentService.createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 取得試算表實例
 */
function getSS() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/**
 * 引入 HTML 檔案（保留舊有支援）
 */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (err) {
    return '';
  }
}