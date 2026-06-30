// ==========================================
// 讀取位置與人員設定api_positions
// ==========================================
function getPositions() {
  const SHEET_NAME = '位置人員清單';
  const sheet = getWorshipSS().getSheetByName(SHEET_NAME);
  
  if (!sheet) throw new Error(`找不到 "${SHEET_NAME}" 工作表`);
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; 
  
  const result = data.slice(1).map(row => {
    return {
      positionName: row[0] || "",
      personnel: row[1] || "",
      isRequired: row[2] || "是"  // 新增：讀取第三欄的「是否必排」，若空白則預設為「是」
    };
  });
  
  return result;
}

// ==========================================
// 儲存位置與人員設定
// ==========================================
function savePositions(positionsData) {
  const SHEET_NAME = '位置人員清單';
  const sheet = getWorshipSS().getSheetByName(SHEET_NAME);
  
  if (!sheet) throw new Error(`找不到 "${SHEET_NAME}" 工作表`);
  
  // 修改：將讀取標題的範圍從 2 欄改為 3 欄 (A到C欄)
  // 💡 請確保您的 Google Sheet 第一列有三個標題，例如：['位置名稱', '人員名單', '是否必排']
  const headers = sheet.getRange(1, 1, 1, 3).getValues()[0];
  let newData = [headers];
  
  positionsData.forEach(item => {
    if (item.positionName && item.positionName.trim() !== '') {
      // 修改：寫入時把第三個欄位 item.isRequired 一併裝進去
      newData.push([
        item.positionName.trim(), 
        item.personnel.trim(), 
        item.isRequired ? item.isRequired.trim() : '是'
      ]);
    }
  });
  
  sheet.clearContents();
  // 修改：將寫入範圍從 2 欄改為 3 欄
  sheet.getRange(1, 1, newData.length, 3).setValues(newData);
  
  return { message: "位置與人員設定已成功儲存！" };
}
