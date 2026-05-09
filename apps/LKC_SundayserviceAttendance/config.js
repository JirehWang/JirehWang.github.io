/**
 * config.js — 橋接設定檔
 *
 * 使用中央路由系統 (central config.js) 中的 window.GAS_URL
 * api.js 期望 window.GAS_CONFIG.apiUrl，此檔案將其對應上去
 */
(function() {
  // 等待中央 config.js 載入並設定 window.GAS_URL
  let retries = 0;
  const maxRetries = 30;

  function checkAndBridge() {
    if (typeof window.GAS_URL !== 'undefined' && window.GAS_URL) {
      window.GAS_CONFIG = {
        apiUrl: window.GAS_URL
      };
      console.log("✅ GAS_CONFIG 已橋接至中央路由：", window.GAS_URL);
    } else if (retries < maxRetries) {
      retries++;
      setTimeout(checkAndBridge, 100);
    } else {
      // 逾時，使用備用 URL
      window.GAS_CONFIG = {
        apiUrl: 'https://script.google.com/macros/s/AKfycbyJbzjHIeFFRbqT-Ttk2OAPYfF-qDKYES8dJiu4sJCR4t2Fq9PTtbALwuiJDBxh55kR/exec'
      };
      console.warn("⚠️ 中央 config.js 載入超時，使用備用 URL");
    }
  }

  checkAndBridge();
})();
