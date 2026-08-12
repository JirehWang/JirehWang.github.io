function setupDatabase() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  const sheetsConfig = [
    {
      name: '服事表總表', 
      headers: ['年度', '季度', '日期', '聚會名稱', '聚會類別', '主領', '配唱1', '配唱2', '配唱3', '吉他', 'BASS', 'Keyboard', '鼓', '其它', '音控', '投影']
    },
    {
      name: '位置人員清單', 
      headers: ['位置名稱', '人員名單', '是否必排']
    },
    {
      name: '牧師題目經文', 
      headers: ['UUID', '日期', '聚會類別', '牧師', '題目', '經文']
    },
    {
      name: '不可排紀錄', 
      headers: ['日期', '人員姓名']
    },
    {
      // 🎵 新增：敬拜曲目工作表
      name: '敬拜曲目',
      headers: ['日期', '聚會名稱', '聚會類別', '主領', '敬拜曲目']
    }
  ];

  sheetsConfig.forEach(config => {
    let sheet = ss.getSheetByName(config.name);
    if (!sheet) {
      sheet = ss.insertSheet(config.name);
      sheet.appendRow(config.headers);
      sheet.setFrozenRows(1);
      const headerRange = sheet.getRange(1, 1, 1, config.headers.length);
      headerRange.setFontWeight('bold').setBackground('#f3f3f3');
    }
  });

  // 刪除預設空白頁
  ['工作表1', 'Sheet1'].forEach(name => {
    let s = ss.getSheetByName(name);
    if (s && ss.getSheets().length > 1 && s.getLastRow() === 0) ss.deleteSheet(s);
  });
}