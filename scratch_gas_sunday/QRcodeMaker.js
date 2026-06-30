/**
 * 批量匯出會友 QR 碼到 Google Drive
 * 在 GAS 編輯器手動執行這個函式即可
 */
function exportQRCodes() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("會友名單");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const values = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  const folderId = "1XzD7WujBRisXYkp2WFtPoqcrVzHloT6T";
  const folder = DriveApp.getFolderById(folderId);

  // 先讀取資料夾內已有的檔案，避免重複生成
  const existingFiles = new Set();
  const files = folder.getFiles();
  while (files.hasNext()) existingFiles.add(files.next().getName());

  let skipCount = 0, successCount = 0;
  let hasUpdates = false;
  const startTime = new Date().getTime();

  // 準備批次更新的狀態陣列 (J 欄，index 9)
  const statusUpdates = values.map(v => [v[9]]);

  for (let i = 0; i < values.length; i++) {
    // 執行時間保護（5.5 分鐘）
    if (new Date().getTime() - startTime > 330000) {
      // 提早退出前，先批次寫回目前進度
      if (hasUpdates) sheet.getRange(2, 10, statusUpdates.length, 1).setValues(statusUpdates);
      SpreadsheetApp.getUi().alert("執行時間接近上限，請再點一次『執行』繼續處理剩餘資料。");
      return;
    }

    const name = values[i][0];   // A 欄：姓名
    const id = values[i][7];     // H 欄：系統編號
    const status = values[i][9]; // J 欄：是否已完成
    const fileName = `${name}_${id}.png`;

    // 已完成的跳過
    if (status === "✅ 已完成" || existingFiles.has(fileName)) {
      if (status !== "✅ 已完成") {
        statusUpdates[i][0] = "✅ 已完成";
        hasUpdates = true;
      }
      skipCount++;
      continue;
    }

    if (id) {
      const qrUrl = `https://quickchart.io/qr?text=${encodeURIComponent(id)}&size=500&ecLevel=H`;
      try {
        const response = UrlFetchApp.fetch(qrUrl, { muteHttpExceptions: true });
        if (response.getResponseCode() === 200) {
          const blob = response.getBlob();
          blob.setName(fileName);
          folder.createFile(blob);
          
          statusUpdates[i][0] = "✅ 已完成";
          hasUpdates = true;
          successCount++;
        }
      } catch (e) {
        console.log(`第 ${i + 2} 列出錯: ${e.message}`);
      }
    }
  }

  // 正常執行完畢，批次寫回所有進度
  if (hasUpdates) sheet.getRange(2, 10, statusUpdates.length, 1).setValues(statusUpdates);

  SpreadsheetApp.getUi().alert(`執行完畢！\n跳過已存在：${skipCount} 筆\n本次新匯出：${successCount} 筆`);
}