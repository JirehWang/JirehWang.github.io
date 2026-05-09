// 📦 中央安全路由設定 (多專案版)
//
// 使用方式：HTML 在載入此檔之前先宣告自己的 key，例如：
//   <script>window._GAS_KEY = 'LKC_worship';</script>
//   <script src="https://jirehwang.github.io/LKC1958_June_1.github.io/config.js"></script>
//
// 若沒有宣告 _GAS_KEY，會 fallback 到 pathname / hostname 推測。
(function() {
  // 📝 子系統 → GAS 部署網址 對應表
  // key 必須與 app 目錄名稱一致
  const _URL_ROUTER = {
    "LKC_worship":                 "https://script.google.com/macros/s/AKfycbyk_6tUucVg-U4rRQjYHvk632teZyxufDkNX_X1WRUXPMGgsTaemVXD_mv9kBDjuSwOnA/exec",
    "LKC_MasterSchedule":          "https://script.google.com/macros/s/AKfycbwiYYWgKxmLRAEaE_pbp_kWyAzlRPcwYVQfvmJVamRJvosvt5wTTkvwebbFBkP8rMqX/exec",
    "LKC_MinistrySchedule":        "https://script.google.com/macros/s/AKfycbx4268IkgwQm2Es0gjDHLU_U9nKJrRMR1-xzbbtuaq08lePLgAQ2wnDRrCeHdy9jNhh/exec",
    "LKC_Group":                   "https://script.google.com/macros/s/AKfycbzfaWh_ooRTGijLV_7lYFUHFm83oL6DvYt9rt6ze5mDXhtwLv8ymxLX_PGuDTHzmNwe/exec",
    "LKC_WhosCar":                 "https://script.google.com/macros/s/AKfycbxOkoaNquIx_V8n_7eS_5ULmoqxPVly_Bezx9_QsmWSzNOcojrCI9Oa6UNd5hOD2euS/exec",
    "LKC_SundayserviceAttendance": "https://script.google.com/macros/s/AKfycbyJbzjHIeFFRbqT-Ttk2OAPYfF-qDKYES8dJiu4sJCR4t2Fq9PTtbALwuiJDBxh55kR/exec",
  };

  const _AUTH_TOKEN = "ChurchApp-2026";

  // 🌟 路由判斷：_GAS_KEY 優先，其次 pathname / hostname
  let rawPath = window.location.pathname.split('/')[1] || "";
  let repoName = rawPath.replace(/\.github\.io$/i, '');
  const hostname = window.location.hostname.split('.')[0];

  let currentKey = null;
  if (window._GAS_KEY && _URL_ROUTER[window._GAS_KEY]) {
    currentKey = window._GAS_KEY;
  } else if (_URL_ROUTER[repoName]) {
    currentKey = repoName;
  } else if (_URL_ROUTER[hostname]) {
    currentKey = hostname;
  }

  // 🛡️ 防呆檢查
  if (!currentKey) {
    console.error(`🚨 [路由錯誤] 找不到對應的 GAS：未宣告 window._GAS_KEY，且 pathname='${repoName}' / hostname='${hostname}' 都不在 _URL_ROUTER 中`);
    window.GAS_URL = null;
  } else {
    window.GAS_URL = _URL_ROUTER[currentKey];
    console.log(`✅ [${currentKey}] 中央路由系統已就緒`);
  }

  window.AUTH_TOKEN = _AUTH_TOKEN;

  // 🚀 中央 API 呼叫
  window.churchAPI = async function(action, data = {}) {
    if (!window.GAS_URL) {
      throw new Error("系統尚未就緒：GAS_URL 為空");
    }

    try {
      const resp = await fetch(window.GAS_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ action: action, token: window.AUTH_TOKEN, data: data })
      });
      return await resp.json();
    } catch (err) {
      console.error("📡 API 通訊失敗:", err);
      throw err;
    }
  };
})();
