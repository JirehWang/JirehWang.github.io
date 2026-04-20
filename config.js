// ============================================================
//  🔐 教會服事管理系統 — 安全設定與 API 路由 (config.js)
//  修正版本 v2.0
// ============================================================

// 🌍 全域設定
const CHURCH_CONFIG = {
  // 🔗 Google Apps Script 部署 URL
  // 請將此替換為你的 Apps Script 部署網址
  // 取得方式：GAS 編輯器 → 部署 → 新建部署 → 選擇類型「網頁應用」
  GAS_DEPLOY_URL: "https://script.google.com/macros/d/{DEPLOYMENT_ID}/usercontent", 
  
  // 🔒 Security Token（必須與後端 core.gs 的 SECRET_TOKEN 一致）
  SECRET_TOKEN: "ChurchApp-2026",
  
  // ⏱️ API 超時時間（毫秒）
  API_TIMEOUT_MS: 30000, // 30 秒
  
  // 🔄 重試機制
  MAX_RETRIES: 3,
  RETRY_DELAY_MS: 1000,
  
  // 📊 Session 過期時間（毫秒）
  SESSION_TTL_MS: 3600000, // 1 小時
};


// ============================================================
//  🛡️ 中央 API 路由函式
//  所有前端的 API 呼叫都經由此函式，統一處理：
//  - Token 注入
//  - 超時保護
//  - 重試邏輯
//  - 錯誤分類
// ============================================================
async function churchAPI(action, params = {}) {
  // 驗證參數
  if (!action || typeof action !== 'string') {
    throw new Error("無效的 action 參數");
  }

  // 注入 Token（如果 params 中已有 token 則保留，否則自動加入）
  const payload = {
    action: action,
    token: params.token || CHURCH_CONFIG.SECRET_TOKEN,
    data: params.data || {},
    ...(params.id && { id: params.id }), // 部分 action 需要 id 參數
    ...(params.type && { type: params.type }) // 部分 action 需要 type 參數
  };

  // 執行 API 呼叫（含重試邏輯）
  return await fetchAPIWithRetry(payload);
}


// ============================================================
//  🔄 API 呼叫與重試機制
// ============================================================
async function fetchAPIWithRetry(payload, attempt = 1) {
  try {
    // 使用 POST 請求（GET 已被棄用用於隱私保護）
    const response = await fetchWithTimeout(
      fetch(CHURCH_CONFIG.GAS_DEPLOY_URL, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Accept': 'application/json',
          'X-Request-ID': generateRequestId() // 追蹤 ID，便於後端日誌
        },
        body: JSON.stringify(payload),
        mode: 'cors',
        credentials: 'omit'
      }),
      CHURCH_CONFIG.API_TIMEOUT_MS
    );

    // HTTP 狀態碼檢查
    if (!response.ok) {
      throw new APIError(
        `HTTP ${response.status}: ${response.statusText}`,
        response.status,
        'HTTP_ERROR'
      );
    }

    // 解析 JSON 回應
    const json = await response.json();

    // 檢查後端是否回傳錯誤
    if (json.status === 'error') {
      throw new APIError(
        json.message || '伺服器錯誤',
        null,
        classifyErrorType(json.message)
      );
    }

    // 成功
    return json;

  } catch (err) {
    // 分類錯誤並決定是否重試
    const errorType = err instanceof APIError ? err.type : classifyErrorType(err.message);

    // 可重試的錯誤類型
    const retryableErrors = ['TIMEOUT', 'NETWORK_ERROR', 'RATE_LIMIT', 'SERVER_ERROR'];

    if (retryableErrors.includes(errorType) && attempt < CHURCH_CONFIG.MAX_RETRIES) {
      // 指數退避重試
      const delay = CHURCH_CONFIG.RETRY_DELAY_MS * attempt;
      console.warn(`⚠️ [${errorType}] 將在 ${delay}ms 後進行第 ${attempt + 1} 次重試...`);
      
      await sleep(delay);
      return fetchAPIWithRetry(payload, attempt + 1);
    }

    // 不可重試或已達最大重試次數
    throw err;
  }
}


// ============================================================
//  ⏱️ 超時包裝函式
// ============================================================
function fetchWithTimeout(fetchPromise, timeoutMs) {
  return Promise.race([
    fetchPromise,
    new Promise((_, reject) =>
      setTimeout(
        () => reject(new APIError('請求逾時，請檢查網路連線', null, 'TIMEOUT')),
        timeoutMs
      )
    )
  ]);
}


// ============================================================
//  🎯 自訂 APIError 類別
// ============================================================
class APIError extends Error {
  constructor(message, httpStatus = null, type = 'UNKNOWN_ERROR') {
    super(message);
    this.name = 'APIError';
    this.httpStatus = httpStatus;
    this.type = type; // 用於前端區分處理邏輯
  }
}


// ============================================================
//  🔍 錯誤類型分類
// ============================================================
function classifyErrorType(message) {
  if (!message) return 'UNKNOWN_ERROR';

  const msg = message.toLowerCase();

  // 認證與授權
  if (msg.includes('無效的憑證') || msg.includes('unauthorized')) return 'AUTH_ERROR';
  if (msg.includes('權限不足')) return 'PERMISSION_ERROR';

  // 連線問題
  if (msg.includes('逾時') || msg.includes('timeout')) return 'TIMEOUT';
  if (msg.includes('網路') || msg.includes('network') || msg.includes('failed to fetch'))
    return 'NETWORK_ERROR';

  // 速率限制
  if (msg.includes('已達上限') || msg.includes('rate limit') || msg.includes('429'))
    return 'RATE_LIMIT';
  if (msg.includes('已達上限')) return 'AI_DAILY_LIMIT';

  // 伺服器錯誤
  if (msg.includes('503') || msg.includes('伺服器忙線')) return 'SERVER_BUSY';
  if (msg.includes('500') || msg.includes('伺服器錯誤')) return 'SERVER_ERROR';

  // 資料錯誤
  if (msg.includes('找不到') || msg.includes('not found')) return 'NOT_FOUND';
  if (msg.includes('無效') || msg.includes('invalid')) return 'VALIDATION_ERROR';

  return 'UNKNOWN_ERROR';
}


// ============================================================
//  🛠️ 工具函式
// ============================================================

/**
 * 延遲指定毫秒數
 */
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * 生成唯一的請求 ID，用於追蹤
 */
function generateRequestId() {
  return `req_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
}

/**
 * Session 管理：儲存編輯狀態
 */
const sessionManager = {
  setUnlocked(groupId) {
    const timestamp = Date.now();
    const data = { unlocked: true, timestamp };
    sessionStorage.setItem(`session_${groupId}`, JSON.stringify(data));
  },

  isUnlocked(groupId) {
    const stored = sessionStorage.getItem(`session_${groupId}`);
    if (!stored) return false;

    try {
      const data = JSON.parse(stored);
      const age = Date.now() - data.timestamp;
      
      // 超過 1 小時則自動過期
      if (age > CHURCH_CONFIG.SESSION_TTL_MS) {
        sessionStorage.removeItem(`session_${groupId}`);
        return false;
      }

      return data.unlocked === true;
    } catch (e) {
      return false;
    }
  },

  clear(groupId) {
    sessionStorage.removeItem(`session_${groupId}`);
  }
};

/**
 * 使用者通知（改進的 UI 反饋）
 */
const userNotification = {
  // 成功訊息（綠色 toast）
  success(message, duration = 3000) {
    this.show(message, 'success', duration);
  },

  // 警告訊息（黃色 toast）
  warning(message, duration = 5000) {
    this.show(message, 'warning', duration);
  },

  // 錯誤訊息（紅色 toast）
  error(message, duration = 5000) {
    this.show(message, 'danger', duration);
  },

  // 資訊訊息（藍色 toast）
  info(message, duration = 3000) {
    this.show(message, 'info', duration);
  },

  // 載入指示器
  showLoading(message = "處理中...") {
    const el = document.getElementById('globalLoading');
    if (el) {
      el.innerText = message;
      el.classList.remove('hidden');
    }
  },

  hideLoading() {
    const el = document.getElementById('globalLoading');
    if (el) {
      el.classList.add('hidden');
    }
  },

  // 顯示 toast 訊息
  show(message, type = 'info', duration = 3000) {
    const bgColorMap = {
      success: 'bg-success',
      warning: 'bg-warning text-dark',
      danger: 'bg-danger',
      info: 'bg-info'
    };

    const toast = document.createElement('div');
    toast.className = `position-fixed bottom-0 end-0 p-3`;
    toast.style.zIndex = '1050';
    toast.innerHTML = `
      <div class="toast show align-items-center text-white ${bgColorMap[type] || 'bg-info'} border-0">
        <div class="d-flex">
          <div class="toast-body">${message}</div>
          <button type="button" class="btn-close btn-close-white me-2 m-auto" data-bs-dismiss="toast"></button>
        </div>
      </div>
    `;
    document.body.appendChild(toast);

    if (duration > 0) {
      setTimeout(() => toast.remove(), duration);
    }
  }
};


// ============================================================
//  🎨 UI 狀態管理
// ============================================================
const uiState = {
  _lockStates: {}, // { actionName: bool }

  // 鎖定按鈕防止重複提交
  lock(actionName) {
    this._lockStates[actionName] = true;
  },

  unlock(actionName) {
    this._lockStates[actionName] = false;
  },

  isLocked(actionName) {
    return this._lockStates[actionName] === true;
  }
};

// ============================================================
//  🌍 全域暴露
// ============================================================
// 確保這些函式在全域作用域中可用
window.churchAPI = churchAPI;
window.sessionManager = sessionManager;
window.userNotification = userNotification;
window.uiState = uiState;
window.APIError = APIError;

console.log("✅ 教會服事管理系統 — 安全設定已加載");
