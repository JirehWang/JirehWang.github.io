/**
 * GeminiHelper.js — 共用 Gemini AI helper
 *
 * 整合自原事工管理 core.js（callGeminiAPIWithRetry）
 * 設計：
 *   - API key 從 PropertiesService 讀取（不要硬編碼到程式碼裡）
 *   - 含重試機制、每日限額、錯誤分類
 *   - 結果可選擇加 cache（同 prompt + 同輸入 → 同結果）
 */

const GEMINI_DAILY_LIMIT = 200;
const GEMINI_MAX_RETRIES = 2;
const GEMINI_RETRY_DELAY_MS = 3000;
const GEMINI_RESULT_CACHE_TTL = 14400; // 4 小時

/**
 * 統一從「LKC系統設定」試算表讀取任意 AI 配置 (具 10 分鐘快取機制)
 * @param {string} keyName 配置名稱 (如 "GEMINI_API_KEY"、"OPENROUTER_API_KEY")
 * @param {string} defaultValue 讀取失敗或未設定時的預設值
 */
function _getAiConfig(keyName, defaultValue) {
  const cache = CacheService.getScriptCache();
  const cacheKey = "GLOBAL_AI_CONFIG_" + keyName;
  const cachedVal = cache.get(cacheKey);
  
  if (cachedVal !== null) {
    return cachedVal; // 若快取內有，直接回傳
  }
  
  const SPREADSHEET_ID = "1kkRbpjXGdwv7ggzM7ojPnKLUwkV7CgoxzL7ZzOwZ2l4";
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("AI_Config");
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for (let i = 0; i < data.length; i++) {
        if (data[i][0] === keyName) {
          const val = data[i][1].toString().trim();
          cache.put(cacheKey, val, 600); // 快取 10 分鐘
          return val;
        }
      }
    }
  } catch (e) {
    console.error("讀取全域設定 [" + keyName + "] 失敗: " + e.message);
  }
  return defaultValue || "";
}

const _AI_PROVIDER_BY_API_KEY = {
  GEMINI_API_KEY: "Gemini",
  NVIDIA_API_KEY: "NVIDIA",
  OPENROUTER_API_KEY: "OpenRouter"
};

/**
 * 依 AI_Config 中 *_API_KEY 的列出順序決定 AI 備援順序。
 * 試算表未提供有效順序時，才回到既有預設順序。
 */
function _getAiProviderOrder() {
  const fallback = ["Gemini", "NVIDIA", "OpenRouter"];
  try {
    const sheet = SpreadsheetApp.openById("1kkRbpjXGdwv7ggzM7ojPnKLUwkV7CgoxzL7ZzOwZ2l4")
      .getSheetByName("AI_Config");
    if (!sheet) return fallback;

    const order = [];
    sheet.getDataRange().getValues().forEach(row => {
      const keyName = String(row && row[0] || "").trim().toUpperCase();
      const provider = _AI_PROVIDER_BY_API_KEY[keyName];
      if (provider && order.indexOf(provider) === -1) order.push(provider);
    });
    return order.length ? order : fallback;
  } catch (e) {
    console.warn("讀取 AI_Config 供應商順序失敗，改用預設順序：" + e);
    return fallback;
  }
}

function _getGeminiApiKey() {
  return _getAiConfig("GEMINI_API_KEY", "");
}

function _getGeminiModel() {
  return _getAiConfig("GEMINI_MODEL", "gemini-3.1-flash-lite");
}

function _getGeminiApiUrl() {
  const key = _getGeminiApiKey();
  const model = _getGeminiModel();
  return "https://generativelanguage.googleapis.com/v1beta/models/" + model + ":generateContent?key=" + key;
}


/**
 * 呼叫 Gemini API，附重試與限額
 * @param {string} systemPrompt 完整 system prompt（已含規則、表頭、名單等）
 * @param {string} userText     使用者輸入
 * @param {Object} opts         { useCache: true/false }
 * @returns {Array|Object}      解析後的 JSON
 */
function callGemini(systemPrompt, userText, opts) {
  opts = opts || {};

  // ── 1. 結果 cache（同樣的 prompt + 輸入 → 直接回傳）──
  const cache = CacheService.getScriptCache();
  let cacheKey = "";
  if (opts.useCache) {
    const hashSrc = systemPrompt + "\n---\n" + userText;
    cacheKey = "GEMINI_RESULT_" + Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, hashSrc)
      .map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
    const cached = cache.get(cacheKey);
    if (cached) {
      try { return JSON.parse(cached); } catch (e) { /* fall through */ }
    }
  }

  // ── 2. 每日限額檢查 ──
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const countKey = "AI_USAGE_COUNT_" + today;
  const currentCount = parseInt(props.getProperty(countKey) || "0", 10);
  if (currentCount >= GEMINI_DAILY_LIMIT) {
    throw new Error("今日 AI 使用次數已達上限 (" + GEMINI_DAILY_LIMIT + " 次)，請明日再試。");
  }
  props.setProperty(countKey, (currentCount + 1).toString());

  // ── 3. 依 AI_Config 的 *_API_KEY 列出順序進行備援呼叫 ──
  const priorityList = _getAiProviderOrder();
  let lastError = "";

  for (let i = 0; i < priorityList.length; i++) {
    const provider = priorityList[i];
    try {
      let res;
      if (provider === "Gemini") {
        res = _tryGeminiDirect(systemPrompt, userText);
      } else if (provider === "NVIDIA") {
        res = _tryNvidiaDirect(systemPrompt, userText);
      } else if (provider === "OpenRouter") {
        res = _tryOpenRouterDirect(systemPrompt, userText);
      } else {
        continue;
      }

      // 寫 cache
      if (opts.useCache && cacheKey) {
        try { cache.put(cacheKey, JSON.stringify(res), GEMINI_RESULT_CACHE_TTL); } catch (e) { /* skip */ }
      }
      return res;

    } catch (err) {
      lastError += provider + " 失敗: " + err.toString() + "; ";
      console.warn(provider + " 呼叫失敗，轉向備援 " + (priorityList[i+1] || "無") + "。錯誤: " + err.toString());
    }
  }

  throw new Error("所有 AI 備援模型均呼叫失敗。詳細錯誤紀錄: " + lastError);
}

// ── 4. 輔助連線函式 ──

function _tryGeminiDirect(systemPrompt, userText) {
  const apiKey = _getGeminiApiKey();
  if (!apiKey) throw new Error("未設定 GEMINI_API_KEY");
  const url = _getGeminiApiUrl();
  const requestPayload = {
    contents: [{ parts: [{ text: "系統指令：" + systemPrompt + "\n\n使用者輸入：\n" + userText }] }],
    generationConfig: { responseMimeType: "application/json" }
  };
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(requestPayload),
    muteHttpExceptions: true
  };

  let lastErr = "";
  for (let attempt = 1; attempt <= GEMINI_MAX_RETRIES; attempt++) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();
      if (code === 429 || code === 503) {
        lastErr = "HTTP " + code;
        if (attempt < GEMINI_MAX_RETRIES) { Utilities.sleep(GEMINI_RETRY_DELAY_MS * attempt); continue; }
        break;
      }
      if (code !== 200) {
        throw new Error("HTTP " + code + ": " + response.getContentText().substring(0, 150));
      }
      const json = JSON.parse(response.getContentText());
      if (json.error) throw new Error(json.error.message);
      
      const aiText = json.candidates[0].content.parts[0].text;
      return JSON.parse(aiText);
    } catch (err) {
      lastErr = err.toString();
      if (attempt < GEMINI_MAX_RETRIES) Utilities.sleep(GEMINI_RETRY_DELAY_MS * attempt);
    }
  }
  throw new Error(lastErr);
}

function _tryNvidiaDirect(systemPrompt, userText) {
  const apiKey = _getAiConfig("NVIDIA_API_KEY", "");
  if (!apiKey) throw new Error("未設定 NVIDIA_API_KEY");
  const model = _getAiConfig("NVIDIA_MODEL", "google/gemma-3n-e2b-it");
  const url = "https://integrate.api.nvidia.com/v1/chat/completions";
  const payload = {
    model: model,
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: userText }
    ],
    response_format: { type: "json_object" }
  };
  const options = {
    method: "post",
    contentType: "application/json",
    headers: { "Authorization": "Bearer " + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  if (code !== 200) {
    throw new Error("HTTP " + code + ": " + response.getContentText().substring(0, 150));
  }
  const json = JSON.parse(response.getContentText());
  if (json.error) throw new Error(json.error.message);
  
  const aiText = json.choices[0].message.content;
  return JSON.parse(aiText);
}

function _tryOpenRouterDirect(systemPrompt, userText) {
  const apiKey = _getAiConfig("OPENROUTER_API_KEY", "");
  if (!apiKey) throw new Error("未設定 OPENROUTER_API_KEY");
  const model = _getAiConfig("OPENROUTER_MODEL", "deepseek/deepseek-chat");
  const url = "https://openrouter.ai/api/v1/chat/completions";
  const payload = {
    model: model,
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: userText }
    ],
    response_format: { type: "json_object" }
  };
  const options = {
    method: "post",
    contentType: "application/json",
    headers: { "Authorization": "Bearer " + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  if (code !== 200) {
    throw new Error("HTTP " + code + ": " + response.getContentText().substring(0, 150));
  }
  const json = JSON.parse(response.getContentText());
  if (json.error) throw new Error(json.error.message);
  
  const aiText = json.choices[0].message.content;
  return JSON.parse(aiText);
}
