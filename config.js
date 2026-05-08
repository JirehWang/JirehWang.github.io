// 📦 中央安全路由設定 (支援多專案 + 嚴格防呆版)
(function() {
  function safeAtob(base64Str) {
    try {
      return atob(base64Str.replace(/-/g, '+').replace(/_/g, '/'));
    } catch (e) {
      console.error("❌ Base64 解碼失敗:", e);
      return null;
    }
  }

  // 📝 你的系統對應表 (新增 WhosCar)
  const _URL_ROUTER = {
    "LKworship": "aHR0cHM6Ly9zY3JpcHQuZ29vZ2xlLmNvbS9tYWNyb3Mvcy9BS2Z5Y2J5a182dFV1Y1ZnLVU0clJRallIdms2MzJ0ZVp5eHVmRGtOWF9YMVdSVVhQTUdnc1RhZW1WWERfbXY5a0JEanVTd09uQS9leGVj",
    "LKCschedule": "aHR0cHM6Ly9zY3JpcHQuZ29vZ2xlLmNvbS9tYWNyb3Mvcy9BS2Z5Y2J3aVlZV2dLeG1MUkFFYUVfcGJwX2tXeUF6bFJQY3dZVlFmdm1KVmFtUkp2b3N2dDV3VFRrdndlYmJGQmtQOHJNcVgvZXhlYw==",
    "LKC1958_June_1": "aHR0cHM6Ly9zY3JpcHQuZ29vZ2xlLmNvbS9tYWNyb3Mvcy9BS2Z5Y2J4NDI2OElrZ3dRbTJFczBnakRITFVfVTluS0pyUk1SMS14emJidHVhcTA4bGVQTGdBUTJ3bkRSckNlSGR5OWpOaGgvZXhlYw==",
    "LKGroup": "aHR0cHM6Ly9zY3JpcHQuZ29vZ2xlLmNvbS9tYWNyb3Mvcy9BS2Z5Y2J6ZmFXaF9vb1JUR2lqTFZfN2xZRlVIRm04M29MNkR2WXQ5cnQ2emU1bURYaHR3THY4eW14TFhfUEd1RFRIem1Od2UvZXhlYw==",
    // 🚗 車牌管理系統專屬路由
    "WhosCar": "aHR0cHM6Ly9zY3JpcHQuZ29vZ2xlLmNvbS9tYWNyb3Mvcy9BS2Z5Y2J4T2tvYU5xdUl4X1Y4bl83ZVM1VUxtb3F4UFZseV9CZXp4OV9Rc21XU3pOT2NvanJDSTlvYTZVTmQ1aE9EMmV1Uy9leGVj",
    // 🛡️ DEFAULT 改為空值，作為防呆機制
    "DEFAULT": "" 
  };

  const _TOKEN_BASE64 = "Q2h1cmNoQXBwLTIwMjY=";

  // 🌟 智慧判斷邏輯
  let rawPath = window.location.pathname.split('/')[1] || "";
  let repoName = rawPath.replace(/\.github\.io$/i, ''); 
  const hostname = window.location.hostname.split('.')[0];

  // 找對應的 Key (會優先比對資料夾名稱，再比對網址前綴)
  const currentKey = _URL_ROUTER[repoName] ? repoName : (_URL_ROUTER[hostname] ? hostname : "DEFAULT");
  const targetEncodedUrl = _URL_ROUTER[currentKey];

  // 🛡️ 防呆檢查
  if (currentKey === "DEFAULT" || !targetEncodedUrl) {
    console.error(`🚨 [路由錯誤] 找不到專案 '${repoName}' 或 '${hostname}' 的後端設定！`);
    window.GAS_URL = null; 
  } else {
    window.GAS_URL = safeAtob(targetEncodedUrl);
    console.log(`✅ [${currentKey}] 中央路由系統已就緒`);
  }
  
  window.AUTH_TOKEN = safeAtob(_TOKEN_BASE64);

  // 🚀 真正的中央路由
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
