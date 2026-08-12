/**
 * 自動產生下一個 LK 編號
 */
function generateNextUID() {
  const sheet = getMemberSheet();
  const data = sheet.getDataRange().getValues();
  let maxNum = 0;
  for (let i = 1; i < data.length; i++) {
    const uid = data[i][7] ? data[i][7].toString().trim() : "";
    if (uid.startsWith("LK")) {
      const num = parseInt(uid.replace("LK", ""), 10);
      if (!isNaN(num) && num > maxNum) maxNum = num;
    }
  }
  return "LK" + (maxNum + 1).toString().padStart(5, '0');
}

/**
 * 1. 取得會友名單工作表
 */
function getMemberSheet() {
  const ss = getSS();
  let sheet = ss.getSheetByName(MEMBER_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(MEMBER_SHEET);
    sheet.appendRow(["姓名", "性別", "建立日期", "備註", "不列入統計", "異動日期", "異動紀錄", "系統編號", "QR Code"]);
  } else {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headers.length < 8 || headers[7] !== "系統編號") sheet.getRange("H1").setValue("系統編號");
    if (headers.length < 9 || headers[8] !== "QR Code") sheet.getRange("I1").setValue("QR Code");
  }
  return sheet;
}

/**
 * 2. 取得所有會友資料
 */
function getAllMembers() {
  const sheet = getMemberSheet();
  const fullData = sheet.getDataRange().getValues();
  if (fullData.length <= 1) return [];

  const headers = fullData[0];
  const rows = fullData.slice(1);
  const targetHeaders = ["姓名", "性別", "建立日期", "備註", "不列入統計", "異動日期", "異動紀錄", "系統編號"];

  const colIndex = {};
  targetHeaders.forEach(h => { colIndex[h] = headers.indexOf(h); });

  return rows.map(row => {
    return targetHeaders.map(h => {
      let cell = (colIndex[h] !== -1 && colIndex[h] < row.length) ? row[colIndex[h]] : "";
      if (cell instanceof Date) return Utilities.formatDate(cell, "GMT+8", "yyyy/MM/dd");
      return cell;
    });
  });
}

/**
 * 3. 新增會友
 */
function addMember(member) {
  const sheet = getMemberSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == member.name) return "⚠️ 新增失敗：姓名 [" + member.name + "] 已存在！";
  }
  const now = new Date();
  const uid = generateNextUID(); // ✅ 改為 LK 自動編號
  sheet.appendRow([member.name, member.gender, now, member.note, member.isExcluded, now, "初始建立", uid]);
  return "✅ 新增成功！編號：" + uid;
}

/**
 * 4. 編輯會友
 */
function updateMember(oldName, newData) {
  const sheet = getMemberSheet();
  const data = sheet.getDataRange().getValues();
  const now = new Date();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == oldName) {
      const rowIndex = i + 1;
      let changeLog = [];
      if (data[i][0] != newData.name) changeLog.push("姓名: " + data[i][0] + "->" + newData.name);
      if (data[i][1] != newData.gender) changeLog.push("性別: " + data[i][1] + "->" + newData.gender);
      if (data[i][3] != newData.note) changeLog.push("備註異動");
      const oldExcluded = (data[i][4] === true || data[i][4] === "TRUE");
      const newExcluded = (newData.isExcluded === true || newData.isExcluded === "TRUE");
      if (oldExcluded !== newExcluded) changeLog.push("統計狀態變更");

      if (changeLog.length > 0) {
        sheet.getRange(rowIndex, 1).setValue(newData.name);
        sheet.getRange(rowIndex, 2).setValue(newData.gender);
        sheet.getRange(rowIndex, 4).setValue(newData.note);
        sheet.getRange(rowIndex, 5).setValue(newData.isExcluded);
        sheet.getRange(rowIndex, 6).setValue(now);
        sheet.getRange(rowIndex, 7).setValue(changeLog.join(" | "));
        return "✅ 更新成功！";
      } else {
        return "⚠️ 資料無異動";
      }
    }
  }
  return "❌ 找不到原始資料，無法更新";
}

/**
 * 5. 刪除會友
 */
function deleteMember(name) {
  try {
    const sheet = getMemberSheet();
    const data = sheet.getDataRange().getValues();
    const targetName = name.toString().trim();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0].toString().trim() === targetName) {
        sheet.deleteRow(i + 1);
        return "🗑️ 成功刪除會友: " + targetName;
      }
    }
    return "❌ 找不到會友 [" + targetName + "]";
  } catch (e) {
    return "❌ 刪除過程出錯: " + e.toString();
  }
}

/**
 * 6. 補齊舊會友系統編號（GAS 編輯器手動執行）
 * ✅ 改為從目前最大 LK 編號往下續編
 */
function batchGenerateMissingQRCodes() {
  const sheet = getMemberSheet();
  const data = sheet.getDataRange().getValues();
  let updatedCount = 0;
  for (let i = 1; i < data.length; i++) {
    if (!data[i][7] || data[i][7].toString().trim() === "") {
      const uid = generateNextUID();
      sheet.getRange(i + 1, 8).setValue(uid);
      updatedCount++;
      Utilities.sleep(100);
    }
  }
  Logger.log("✅ 共補齊了 " + updatedCount + " 位會友的系統編號。");
}

// ==========================================
//  卡片製作
// ==========================================

const CARD_CONFIG = {
  TEMPLATE_SLIDE_ID: '13hmS49cHuIgX-P0lYxa-rJ_o29xNnNExr3v2BaLcdKw',
  OUTPUT_FOLDER_ID: '1XzD7WujBRisXYkp2WFtPoqcrVzHloT6T',
  QR_API: "https://quickchart.io/qr?size=300x300&text="
};

/**
 * 7. 產生個人卡片（儲存 JPG 到雲端硬碟）
 */
function generateMemberCard(member) {
  if (!member) return { success: false, msg: "未接收到資料" };
  let tempCopy = null;
  try {
    const name = String(member.name || "").trim();
    const uid  = String(member.uid  || "").trim();

    const templateFile = DriveApp.getFileById(CARD_CONFIG.TEMPLATE_SLIDE_ID);
    const folder       = DriveApp.getFolderById(CARD_CONFIG.OUTPUT_FOLDER_ID);

    // 1. 複製模板（暫存用）
    tempCopy = templateFile.makeCopy(`[暫存]_${name}`);
    const tempId     = tempCopy.getId();
    const presentation = SlidesApp.openById(tempId);
    const slide        = presentation.getSlides()[0];

    // 2. 替換文字
    slide.replaceAllText('{{NAME}}', name);
    slide.replaceAllText('{{UID}}',  uid);

    // 3. 插入 QR Code
    const qrUrl      = CARD_CONFIG.QR_API + encodeURIComponent(uid);
    const qrResponse = UrlFetchApp.fetch(qrUrl, { muteHttpExceptions: true });
    if (qrResponse.getResponseCode() === 200) {
      const qrImage = slide.insertImage(qrResponse.getBlob());
      qrImage.setLeft(35).setTop(35).setWidth(525).setHeight(525);
    }
    presentation.saveAndClose();

    // 4. 等待 Google 處理完成
    Utilities.sleep(1500);

    // 5. 匯出第一頁為 JPEG
    const exportUrl = `https://docs.google.com/presentation/d/${tempId}/export/jpeg`;
    const response  = UrlFetchApp.fetch(exportUrl, {
      headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) throw new Error("匯出圖片失敗，HTTP " + response.getResponseCode());

    // 6. 將 JPG 存到指定資料夾
    const jpgBlob = response.getBlob().setName(`${name}_個人名卡.jpg`);
    const jpgFile = folder.createFile(jpgBlob);

    // 7. 刪除暫存 Slide
    tempCopy.setTrashed(true);

    return { success: true, fileUrl: jpgFile.getUrl(), fileName: jpgFile.getName() };

  } catch (e) {
    if (tempCopy) { try { tempCopy.setTrashed(true); } catch(ex) {} }
    return { success: false, msg: e.toString() };
  }
}

/**
 * 8. 預覽卡片（回傳 JPG base64）
 */
function previewMemberCard(member) {
  if (!member) return { success: false, msg: "未接收到會友資料" };
  let tempCopy = null;
  try {
    const name = String(member.name || "").trim() || "空姓名";
    const uid = String(member.uid || "").trim();
    if (!uid || uid === "無編號") throw new Error("無效的系統編號");
    const templateFile = DriveApp.getFileById(CARD_CONFIG.TEMPLATE_SLIDE_ID);
    tempCopy = templateFile.makeCopy(`[暫存預覽]_${name}`);
    const tempId = tempCopy.getId();
    const presentation = SlidesApp.openById(tempId);
    const slide = presentation.getSlides()[0];
    slide.replaceAllText('{{NAME}}', name);
    slide.replaceAllText('{{UID}}', uid);
    const qrUrl = CARD_CONFIG.QR_API + encodeURIComponent(uid);
    const qrResponse = UrlFetchApp.fetch(qrUrl, { muteHttpExceptions: true });
    if (qrResponse.getResponseCode() === 200) {
      const qrImage = slide.insertImage(qrResponse.getBlob());
      qrImage.setLeft(35).setTop(35).setWidth(525).setHeight(525);
    }
    presentation.saveAndClose();
    Utilities.sleep(1500);
    const exportUrl = `https://docs.google.com/presentation/d/${tempId}/export/jpeg`;
    const response = UrlFetchApp.fetch(exportUrl, {
      headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) throw new Error("匯出圖片失敗");
    const base64Str = Utilities.base64Encode(response.getBlob().getBytes());
    tempCopy.setTrashed(true);
    return { success: true, base64: "data:image/jpeg;base64," + base64Str };
  } catch (e) {
    if (tempCopy) { try { tempCopy.setTrashed(true); } catch(ex) {} }
    return { success: false, msg: e.toString() };
  }
}

/**
 * 批量產生所有會友名卡（支援斷點續跑）
 * 每次手動執行會從上次中斷處繼續，全部完成後自動清除進度
 */
function batchGenerateAllMemberCards() {
  const MAX_RUNTIME_MS = 5 * 60 * 1000;
  const startTime = Date.now();

  const props = PropertiesService.getScriptProperties();
  const sheet = getMemberSheet();
  const data = sheet.getDataRange().getValues();

  let startRow     = parseInt(props.getProperty("BATCH_CARD_START_ROW") || "1", 10);
  let successCount = parseInt(props.getProperty("BATCH_CARD_SUCCESS")   || "0", 10);
  let skipCount    = parseInt(props.getProperty("BATCH_CARD_SKIP")      || "0", 10);
  let failList     = JSON.parse(props.getProperty("BATCH_CARD_FAILS")   || "[]");

  Logger.log("▶️ 從第 " + startRow + " 列開始，共 " + (data.length - 1) + " 人");

  for (let i = startRow; i < data.length; i++) {

    // 逾時檢查
    if (Date.now() - startTime > MAX_RUNTIME_MS) {
      saveProgress(props, i, successCount, skipCount, failList);
      const msg = "⏸️ 已暫停（接近時間上限）\n目前進度：第 " + i + " / " + (data.length - 1) + " 人\n✅ 成功：" + successCount + "　❌ 失敗：" + failList.length + "\n\n請再次執行繼續";
      Logger.log(msg);
      SpreadsheetApp.getUi().alert(msg);
      return;
    }

    const name = data[i][0] ? data[i][0].toString().trim() : "";
    const uid  = data[i][7] ? data[i][7].toString().trim() : "";

    if (!name || !uid) {
      skipCount++;
      // ✅ 跳過也立刻存進度
      saveProgress(props, i + 1, successCount, skipCount, failList);
      continue;
    }

    try {
      const result = generateMemberCard({ name: name, uid: uid });
      if (result.success) {
        successCount++;
        Logger.log("✅ [" + i + "] " + name + " (" + uid + ")");
      } else {
        failList.push(name + "：" + result.msg);
        Logger.log("❌ [" + i + "] " + name + " 失敗：" + result.msg);
      }
    } catch (e) {
      failList.push(name + "：" + e.toString());
      Logger.log("❌ [" + i + "] " + name + " 例外：" + e.toString());
    }

    // ✅ 每處理完一人，無論成功失敗，立刻儲存進度到下一列
    saveProgress(props, i + 1, successCount, skipCount, failList);

    Utilities.sleep(1500);
  }

  // 全部完成，清除進度
  props.deleteProperty("BATCH_CARD_START_ROW");
  props.deleteProperty("BATCH_CARD_SUCCESS");
  props.deleteProperty("BATCH_CARD_SKIP");
  props.deleteProperty("BATCH_CARD_FAILS");

  const summary = [
    "===== 全部完成 =====",
    "✅ 成功：" + successCount + " 人",
    "⏭️ 跳過：" + skipCount + " 人",
    "❌ 失敗：" + failList.length + " 人",
    failList.length > 0 ? "失敗名單：\n" + failList.join("\n") : ""
  ].join("\n");

  Logger.log(summary);
  SpreadsheetApp.getUi().alert(summary);
}

function forceSetStartRow() {
  PropertiesService.getScriptProperties().setProperty("BATCH_CARD_START_ROW", "77");
  Logger.log("✅ 已設定從第 77 列開始");
}
/**
 * 統一的進度儲存函式
 */
function saveProgress(props, nextRow, successCount, skipCount, failList) {
  props.setProperty("BATCH_CARD_START_ROW", String(nextRow));
  props.setProperty("BATCH_CARD_SUCCESS",   String(successCount));
  props.setProperty("BATCH_CARD_SKIP",      String(skipCount));
  props.setProperty("BATCH_CARD_FAILS",     JSON.stringify(failList));
}

/**
 * 手動重置進度（如果想從頭重跑時執行）
 */
function resetBatchCardProgress() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty("BATCH_CARD_START_ROW");
  props.deleteProperty("BATCH_CARD_SUCCESS");
  props.deleteProperty("BATCH_CARD_SKIP");
  props.deleteProperty("BATCH_CARD_FAILS");
  SpreadsheetApp.getUi().alert("✅ 進度已重置，下次執行將從頭開始");
}