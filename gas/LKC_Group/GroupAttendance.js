/**
 * GroupAttendance.gs - 處理與小組點名相關的業務 (支援 C 欄身分管理)
 */

// 1. 檢查是否初始化名單 (優化：讀取 A 到 C 欄)
function checkGroupStatus(groupName) {
  const mSheet = getSheetSafely(groupName + "_名單");
  if (!mSheet) return { isInitialized: false };
  
  const lastRow = mSheet.getLastRow();
  if (lastRow <= 1) return { isInitialized: true, members: [] };

  // 💡 讀取 A、B、C 三欄 (1:姓名, 2:日期, 3:身分)
  const data = mSheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const members = [];
  
  data.forEach(row => {
    const name = row[0] ? row[0].toString().trim() : "";
    if (name) {
      members.push({
        name: name,
        // 如果 C 欄有值就讀取，沒有的話就預設為小羊，確保舊資料相容
        role: row[2] ? row[2].toString().trim() : "小羊"
      });
    }
  });
                        
  return { isInitialized: true, members: members };
}

// 2. 初始化小組分頁 (寫入 A、B、C 欄)
function initGroup(groupName, members) {
  const ss = getSs(); 
  
  if (getSheetSafely(groupName + "_名單")) {
    return { success: false, message: "此小組已存在名單" };
  }

  const now = new Date();
  // 💡 前端傳來的是物件陣列 [{name: "張三", role: "小羊"}, ...]
  // 整理成 2D 陣列 [[姓名, 日期, 身分], ...]
  const memberRows = members
    .filter(m => m.name && m.name.trim() !== "")
    .map(m => [m.name.trim(), now, m.role || "小羊"]);

  // 1. 建立名單分頁並批次寫入
  const mSheet = ss.insertSheet(groupName + "_名單");
  mSheet.appendRow(["姓名", "建立日期", "身分"]); // 標題加上 C 欄
  
  if (memberRows.length > 0) {
    // 寫入 3 個欄位
    mSheet.getRange(2, 1, memberRows.length, 3).setValues(memberRows);
  }
  
  // 2. 建立點名紀錄分頁 (點名紀錄不受影響，維持原樣)
  const rSheet = ss.insertSheet(groupName + "_點名紀錄");
  rSheet.appendRow(["日期", "出席人員", "缺席人員", "新朋友", "實到人數"]);
  
  return { success: true, message: "初始化成功" };
}

// 3. 送出點名結果 (這段完全不影響，因為傳來的 present 本來就只有純名字)
function submitAttendance(groupName, date, present, absent, newFriends) {
  const rSheet = getSheetSafely(groupName + "_點名紀錄");
  if (!rSheet) return { success: false, message: "找不到紀錄表，請重新初始化" };
  
  const nfList = newFriends ? newFriends.split(/[,，]/).filter(n => n.trim()) : [];
  const totalCount = present.length + nfList.length;
  
  rSheet.appendRow([
    date,
    present.join(", "),
    absent.join(", "),
    newFriends,
    totalCount
  ]);
  
  return { success: true, message: "點名成功" };
}

// 4. 更新小組成員名單 (保護 B 欄日期，更新 A 欄與 C 欄)
function updateMemberList(groupName, members) {
  const mSheet = getSheetSafely(groupName + "_名單");
  if (!mSheet) return { success: false, message: "找不到名單分頁，無法更新" };
  
  const lastRow = mSheet.getLastRow();
  const dateMap = {};

  // 1. 取得現有名單的日期 (用來保護當初加入的時間)
  if (lastRow > 1) {
    const oldData = mSheet.getRange(2, 1, lastRow - 1, 2).getValues(); // 只需要讀前兩欄來存日期
    oldData.forEach(row => {
      const oldName = row[0] ? row[0].toString().trim() : "";
      if (oldName) dateMap[oldName] = row[1];
    });
    // 2. 💡 清空 A、B、C 三欄的內容
    mSheet.getRange(2, 1, lastRow - 1, 3).clearContent();
  }

  // 3. 整理新名單
  const now = new Date();
  const newRows = members
    .filter(m => m.name && m.name.trim() !== "")
    .map(m => {
      const cleanName = m.name.trim();
      // 陣列結構：[A欄: 姓名, B欄: 舊日期或今天, C欄: 身分]
      return [cleanName, dateMap[cleanName] || now, m.role || "小羊"];
    });

  // 4. 批次寫入新名單 (寫入 3 個欄位)
  if (newRows.length > 0) {
    mSheet.getRange(2, 1, newRows.length, 3).setValues(newRows);
  }

  return { success: true, message: "名單更新成功！" };
}

// 5. 修改歷史點名紀錄
function updateAttendanceRecord(groupName, originalDate, newDate, present, absent, newFriends) {
  const rSheet = getSheetSafely(groupName + "_點名紀錄");
  if (!rSheet) return { success: false, message: "找不到紀錄表" };

  const data = rSheet.getDataRange().getValues();
  let targetRowIndex = -1;

  // 尋找符合原始日期的那一列
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    let d = new Date(data[i][0]);
    let dStr = Utilities.formatDate(d, "GMT+8", "yyyy-MM-dd");
    if (dStr === originalDate) {
      targetRowIndex = i + 1; // +1 是因為試算表列數從 1 開始計算
      break;
    }
  }

  if (targetRowIndex === -1) {
    return { success: false, message: "找不到該日期的紀錄，無法修改" };
  }

  const nfList = newFriends ? newFriends.split(/[,，]/).filter(n => n.trim()) : [];
  const totalCount = present.length + nfList.length;

  // 直接覆寫該列的 5 個欄位
  rSheet.getRange(targetRowIndex, 1, 1, 5).setValues([[
    newDate,
    present.join(", "),
    absent.join(", "),
    newFriends,
    totalCount
  ]]);

  return { success: true, message: "紀錄修改成功" };
}

// 6. 刪除整筆點名紀錄
function deleteAttendanceRecord(groupName, originalDate) {
  const rSheet = getSheetSafely(groupName + "_點名紀錄");
  if (!rSheet) return { success: false, message: "找不到紀錄表" };

  const data = rSheet.getDataRange().getValues();
  let targetRowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    let d = new Date(data[i][0]);
    let dStr = Utilities.formatDate(d, "GMT+8", "yyyy-MM-dd");
    if (dStr === originalDate) {
      targetRowIndex = i + 1;
      break;
    }
  }

  if (targetRowIndex !== -1) {
    rSheet.deleteRow(targetRowIndex); // 直接刪除該列
    return { success: true, message: "紀錄刪除成功" };
  }
  return { success: false, message: "找不到該日期的紀錄" };
}