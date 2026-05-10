/**
 * 💡 批量建立小組分頁並寫入 Config
 * 執行前請確認：
 * 1. 試算表中已存在名為「小組聚會表模板」的分頁
 * 2. 試算表中已存在名為「Config」的分頁
 */
function batchInitializeGroups() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  
  if (!configSheet) {
    throw new Error("找不到『Config』分頁，請先建立它！");
  }

  // 待處理的原始資料
  const rawData = [
    ["VN01", "葡萄樹A組", "小組聚會表模板", "啟用"],
    ["VN02", "葡萄樹B組", "小組聚會表模板", "啟用"],
    ["PM01", "棕樹A組", "小組聚會表模板", "啟用"],
    ["PM02", "棕樹B組", "小組聚會表模板", "啟用"],
    ["SD01", "種子小組", "小組聚會表模板", "啟用"],
    ["TM01", "提摩太小組", "小組聚會表模板", "啟用"],
    ["OL01", "橄欖樹小組", "小組聚會表模板", "啟用"],
    ["GR01", "恩典團契", "小組聚會表模板", "啟用"],
    ["YT01", "學青小組", "小組聚會表模板", "啟用"],
    ["ES01", "以斯帖小組", "小組聚會表模板", "啟用"],
    ["SN01", "松年團契", "小組聚會表模板", "啟用"],
    ["CD01", "香柏樹小組", "小組聚會表模板", "啟用"]
  ];

  console.log("🏁 開始批量初始化...");

  rawData.forEach(function(row) {
    const id = row[0];
    const groupName = row[1];
    const templateName = row[2];
    const status = row[3];

    // 1. 檢查分頁是否已存在，避免重複建立報錯
    if (ss.getSheetByName(groupName)) {
      console.warn("⚠️ 分頁「" + groupName + "」已存在，跳過建立動作。");
    } else {
      // 2. 尋找模板並複製
      const templateSheet = ss.getSheetByName(templateName);
      if (templateSheet) {
        templateSheet.copyTo(ss).setName(groupName);
        console.log("✅ 已建立分頁：" + groupName);
      } else {
        console.error("❌ 找不到模板分頁：「" + templateName + "」，請檢查名稱是否正確。");
        return; // 跳過此筆
      }
    }

    // 3. 檢查 Config 是否已有此 ID，若無則寫入
    const configData = configSheet.getDataRange().getValues();
    const existingIds = configData.map(function(r) { return r[0].toString().trim(); });
    
    if (existingIds.indexOf(id) === -1) {
      configSheet.appendRow([id, groupName, templateName, status]);
      console.log("📝 已將「" + groupName + "」寫入 Config 表。");
    } else {
      console.warn("ℹ️ Config 已有 ID「" + id + "」，跳過寫入。");
    }
  });

  console.log("🎊 批量初始化完成！");
  SpreadsheetApp.getUi().alert("系統初始化完成！共處理 " + rawData.length + " 筆資料。");
}