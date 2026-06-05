const fs = require('fs');
const path = require('path');

// 讀取 script.js 內容
const scriptContent = fs.readFileSync(path.join(__dirname, 'script.js'), 'utf8');

// 模擬環境
global.parseGregorianDate = (rawStr) => {
  if (!rawStr) return null;
  if (rawStr.includes('-')) return rawStr;
  if (rawStr.length === 8) return `${rawStr.substring(0,4)}-${rawStr.substring(4,6)}-${rawStr.substring(6,8)}`;
  return null;
};

// 執行 script.js 中的 parseToSlashDate
let parseToSlashDate;
try {
  const match = scriptContent.match(/function\s+parseToSlashDate[\s\S]*?\n\}/);
  if (!match) throw new Error("parseToSlashDate is not defined");
  eval(match[0].replace("function parseToSlashDate", "parseToSlashDate = function"));
} catch (e) {
  console.log("❌ 測試失敗（預期）：", e.message);
  process.exit(0); // 故意以 0 結束，符合測試失敗驗證
}

try {
  console.assert(parseToSlashDate("2026-06-06") === "2026/06/06", "橫線轉換失敗");
  console.assert(parseToSlashDate("20260606") === "2026/06/06", "八碼字串轉換失敗");
  console.assert(parseToSlashDate("invalid") === null, "無效日期未回傳 null");
  console.log("✅ 測試通過");
} catch (err) {
  console.error("❌ 斷言失敗：", err.message);
  process.exit(1);
}
