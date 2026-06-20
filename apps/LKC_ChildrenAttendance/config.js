/**
 * config.js — 橋接設定檔
 *
 * 使用中央路由系統 (central config.js) 中的 window.GAS_URL
 * api.js 期望 window.GAS_CONFIG.apiUrl，此檔案將其對應上去。
 *
 * 由於 <script> 同步載入順序保證中央 config.js 在此檔之前已執行，
 * window.GAS_URL 此時已就緒；無需輪詢。
 */
(function() {
  const fallbackUrl = 'https://script.google.com/macros/s/AKfycbxBOFeLiXu23kBMGU8iSvRyJci6fruTfk7HdahhcQFY777sCPSgasuNM7Z1CeuzuS-r/exec';
  if (window.GAS_URL) {
    window.GAS_CONFIG = { apiUrl: window.GAS_URL };
    console.log("✅ GAS_CONFIG 已橋接至中央路由：", window.GAS_URL);
  } else {
    window.GAS_CONFIG = { apiUrl: fallbackUrl };
    console.warn("⚠️ 中央 config.js 未提供 GAS_URL，使用備用 URL");
  }
})();
