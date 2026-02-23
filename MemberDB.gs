// ==========================================
//  會友系統後端核心 (MemberDB.gs)
// ==========================================


/**
 * 0. 取得試算表物件 (共用工具)
 */
function getSS() {
  if (typeof SPREADSHEET_ID !== 'undefined') return SpreadsheetApp.openById(SPREADSHEET_ID);
  // 請填入你的 ID
  const MY_ID = "請在此填入你的試算表ID"; 
  return SpreadsheetApp.openById(MY_ID);
}

/**
 * 1. 取得會友名單工作表
 */
function getMemberSheet() {
  const ss = getSS();
  let sheet = ss.getSheetByName(MEMBER_SHEET);
  if (!sheet) {
    // 如果找不到表，自動建立並寫入標題
    sheet = ss.insertSheet(MEMBER_SHEET);
    sheet.appendRow(["姓名", "性別", "建立日期", "備註", "不列入統計", "異動日期", "異動紀錄"]);
  }
  return sheet;
}

/**
 * 2. 取得所有會友資料 (供前端列表顯示)
 */
function getAllMembers() {
  const sheet = getMemberSheet();
  const fullData = sheet.getDataRange().getValues();
  if (fullData.length <= 1) return []; // 只有標題或沒資料

  const headers = fullData[0]; // 第一列：標題列
  const rows = fullData.slice(1); // 剩下的資料列

  // 定義我們想要抓取的標題名稱 (對應前端表格順序)
  // 0:姓名, 1:性別, 2:建立日期, 3:備註, 4:不列入統計, 5:異動日期, 6:異動紀錄
  const targetHeaders = ["姓名", "性別", "建立日期", "備註", "不列入統計", "異動日期", "異動紀錄"];
  
  // 建立標題與索引的對照表
  const colIndex = {};
  targetHeaders.forEach(h => {
    colIndex[h] = headers.indexOf(h);
  });

  // 開始按照標題映射資料
  return rows.map(row => {
    return targetHeaders.map(h => {
      let cell = (colIndex[h] !== -1) ? row[colIndex[h]] : ""; 
      
      // 日期處理：轉成易讀格式
      if (cell instanceof Date) {
        return Utilities.formatDate(cell, "GMT+8", "yyyy/MM/dd");
      }
      return cell;
    });
  });
}

/**
 * 3. 新增會友
 * 結構：姓名, 性別, 建立日期, 備註, 不列入統計, 異動日期, 異動紀錄
 */
function addMember(member) {
  const sheet = getMemberSheet();
  const data = sheet.getDataRange().getValues();
  
  // 檢查重複 (防呆)
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == member.name) {
      return "⚠️ 新增失敗：姓名 [" + member.name + "] 已存在！";
    }
  }

  const now = new Date();
  sheet.appendRow([
    member.name, 
    member.gender, 
    now, 
    member.note, 
    member.isExcluded, 
    now, 
    "初始建立"
  ]);
  return "✅ 新增成功！";
}

/**
 * 4. 編輯會友（含自動異動追蹤）
 */
function updateMember(oldName, newData) {
  const sheet = getMemberSheet();
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  
  // 標題索引對照 (假設順序固定，若怕順序變動可用 indexOf 動態抓)
  // 0:姓名, 1:性別, 2:建立日, 3:備註, 4:不統計, 5:異動日, 6:異動紀錄
  
  let rowIndex = -1;
  
  // 尋找該會友在哪一行
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == oldName) {
      rowIndex = i + 1; // 轉成 Sheet 的行號 (從1開始)
      
      let changeLog = [];
      const oldData = data[i]; 
      
      // 比對差異
      if (oldData[0] != newData.name) changeLog.push(`姓名: ${oldData[0]}->${newData.name}`);
      if (oldData[1] != newData.gender) changeLog.push(`性別: ${oldData[1]}->${newData.gender}`);
      if (oldData[3] != newData.note) changeLog.push(`備註異動`);
      
      // 處理 boolean 值的比對
      const oldExcluded = (oldData[4] === true || oldData[4] === "TRUE");
      const newExcluded = (newData.isExcluded === true || newData.isExcluded === "TRUE");
      if (oldExcluded !== newExcluded) changeLog.push(`統計狀態變更`);
      
      if (changeLog.length > 0) {
        // 有變更才寫入
        sheet.getRange(rowIndex, 1).setValue(newData.name);
        sheet.getRange(rowIndex, 2).setValue(newData.gender);
        // 建立日期(Col 3) 不動
        sheet.getRange(rowIndex, 4).setValue(newData.note);
        sheet.getRange(rowIndex, 5).setValue(newData.isExcluded);
        
        // 更新異動資訊
        sheet.getRange(rowIndex, 6).setValue(now);
        
        // 異動紀錄(Col 7) 採用「累加」方式，或是「覆蓋顯示最新」
        // 這裡示範：覆蓋顯示最新異動內容
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
 * 5. 刪除會友 (嚴謹版)
 */
function deleteMember(name) {
  try {
    const sheet = getMemberSheet();
    const data = sheet.getDataRange().getValues();
    const targetName = name.toString().trim();

    // 從後面往前面刪，避免索引跑掉 (雖然這裡只刪一筆，但習慣上從後刪較安全)
    for (let i = data.length - 1; i >= 1; i--) {
      const currentCellName = data[i][0].toString().trim();

      if (currentCellName === targetName) {
        sheet.deleteRow(i + 1);
        return "🗑️ 成功刪除會友: " + targetName;
      }
    }
    return "❌ 找不到會友 [" + targetName + "]"; 
  } catch (e) {
    return "❌ 刪除過程出錯: " + e.toString();
  }
}
