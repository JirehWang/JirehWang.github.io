// ============================================================
//  🏛️ 教會服事管理系統 — 後端核心 (core.gs)
//  修正版本 v3.0
//  v2.0 修正項目：
//    1. SECRET_TOKEN 改用 Script Properties（安全性）
//    2. 補充缺失的 GEMINI_API_URL 常數與設定集中管理
//    3. getSystemPromptTemplate() 完整實作
//    4. AI 呼叫加入重試機制與逾時處理
//    5. 彙整報表加入快取優化
//    6. 自動快取失效檢查（不再依賴手動觸發器）
//    7. 加入審計日誌記錄
//    8. 統一設定集中管理 CONFIG 物件
//  v3.0 新增項目：
//    9. 整合 autoSyncSmallGroups（UUID 核心版）
//       - Config 欄位升級為 7 欄，以 UUID 為穩定識別碼
//       - 自動建立、改名、同步狀態、同步小組編號
//       - 外部清單若無 UUID 自動補寫
//    10. 全部函式的欄位索引對應新結構更新
//
//  ⚠️ Config 分頁欄位結構（v3.0）：
//    A(0) = UUID（唯一識別碼，由外部系統或自動產生）
//    B(1) = 小組編號 ID（如 VN01，可改變）
//    C(2) = 小組名稱（如 葡萄樹A組，可改變）
//    D(3) = 模板類型（如 小組聚會表模板）
//    E(4) = 狀態（啟用 / 停用）
//    F(5) = AI 規則（groupPrompt）
//    G(6) = 同工名單（JSON 字串）
//
//  前端 URL 傳遞的 ?id= 為「小組編號」(欄 B)，非 UUID
// ============================================================

// ============================================================
//  🔧 統一設定中心（所有常數都從這裡取，不寫死在程式碼）
// ============================================================
function getConfig() {
  var props = PropertiesService.getScriptProperties();
  return {
    // 🔒 安全金鑰：請在 GAS 後台 → 專案設定 → 指令碼屬性 中設定
    //    key: SECRET_TOKEN  value: 你的金鑰字串
    SECRET_TOKEN: props.getProperty("SECRET_TOKEN") || "ChurchApp-2026",

    // 🤖 Gemini API 設定（請在指令碼屬性設定 GEMINI_API_KEY）
    GEMINI_API_KEY: props.getProperty("GEMINI_API_KEY") || "",
    get GEMINI_API_URL() {
      return getGeminiApiUrl();  // 呼叫 config.gs 的函式
    },

    // 📊 外部試算表 ID
    EXTERNAL_EVENTS_SHEET_ID: "1tKI5k7HwI9S2bTRV6RuKrzxBFGzROYHytsFYHuH8H7E",
    EXTERNAL_MEMBERS_SHEET_ID: "1opAz6PFYveCF4oP9d4c_Dppd2SWBgE_d4zk0a0dRIoU",

    // ⚡ 快取設定（秒）
    CACHE_MEMBERS_TTL:  21600, // 6 小時
    CACHE_EVENTS_TTL:   21600, // 6 小時
    CACHE_REPORT_TTL:    3600, // 1 小時

    // 🤖 AI 限制
    AI_DAILY_LIMIT: 200,
    AI_MAX_RETRIES: 2,
    AI_RETRY_DELAY: 3000  // ms
  };
}


// ============================================================
//  🔐 Token 驗證輔助
// ============================================================
function isValidToken(token) {
  return token === getConfig().SECRET_TOKEN;
}


// ============================================================
//  🔧 Config 分頁自動偵測與遷移（舊版 6 欄 → 新版 7 欄）
//
//  舊版欄位（6欄）：A=ID, B=名稱, C=模板, D=狀態, E=規則, F=名單
//  新版欄位（7欄）：A=UUID, B=ID, C=名稱, D=模板, E=狀態, F=規則, G=名單
//
//  偵測：讀第一筆資料的 A 欄，判斷是否為 UUID 格式
//  快取：用 ScriptProperties 記住「已是新版」，避免每次 request 都掃描
// ============================================================
function ensureConfigSchemaV3() {
  var props = PropertiesService.getScriptProperties();

  // 已確認是新版結構，直接跳過
  if (props.getProperty("CONFIG_SCHEMA_V3") === "true") return;

  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName('Config');

  if (!configSheet || configSheet.getLastRow() < 2) {
    // 空表示新建，標記為新版
    props.setProperty("CONFIG_SCHEMA_V3", "true");
    return;
  }

  // 讀取第一筆資料列（row 2）的 A 欄
  var firstA = configSheet.getRange(2, 1).getValue().toString().trim();

  // UUID 格式：含 "-" 且長度 > 20（如 4eeb78c9-1969-4f1f-ab4e-5b2401d492a2）
  var looksLikeUUID = firstA.indexOf("-") !== -1 && firstA.length > 20;

  if (looksLikeUUID) {
    props.setProperty("CONFIG_SCHEMA_V3", "true");
    return;
  }

  // ── 舊版，執行一次性遷移 ──
  var lastRow = configSheet.getLastRow();
  console.log("🔄 偵測到舊版 Config（" + lastRow + " 列），自動遷移至新版 7 欄格式...");

  // 在 A 欄前插入一欄
  configSheet.insertColumnBefore(1);

  // 每列填入新 UUID（第 1 列到最後一列）
  for (var row = 1; row <= lastRow; row++) {
    configSheet.getRange(row, 1).setValue(Utilities.getUuid());
  }

  invalidateReportCache();
  props.setProperty("CONFIG_SCHEMA_V3", "true");

  console.log("✅ Config 遷移完成，共處理 " + lastRow + " 列");
  writeAuditLog("system", "configMigrated", { rows: lastRow, from: "6col", to: "7col" });
}


// ============================================================
//  📥 doGet — 向下相容，保留 GET 讀取能力
// ============================================================
function doGet(e) {
  var token = e.parameter.token;
  var action = e.parameter.action;

  if (!isValidToken(token)) {
    return ContentService.createTextOutput("未授權的存取：無效的憑證").setMimeType(ContentService.MimeType.TEXT);
  }

  try {
    if (action === "getGroups")          return createJsonResponse({ status: "success", data: getExistingGroups() });
    if (action === "getTemplates")       return createJsonResponse({ status: "success", data: getTemplates() });
    if (action === "getPageConfig" && e.parameter.id)
                                         return createJsonResponse({ status: "success", data: getPageConfig(e.parameter.id) });
    if (action === "getAggregatedReport" && e.parameter.type)
                                         return createJsonResponse({ status: "success", data: getAggregatedReport(e.parameter.type) });

    return ContentService.createTextOutput("✅ API 運作中... 中央安全路由與快取系統已連線！").setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    writeAuditLog("system", "doGet_error", { action: action, error: err.toString() });
    return createJsonResponse({ status: "error", message: err.toString() });
  }
}


// ============================================================
//  📤 doPost — 主要入口點，統一處理所有讀寫動作
// ============================================================
function doPost(e) {
  try {
    // 同時支援兩種 Content-Type：
    //   application/json              → e.postData.contents 是 JSON 字串
    //   application/x-www-form-urlencoded → e.postData.contents 是 payload=...
    //   （後者用於繞過瀏覽器 CORS preflight）
    var raw = e.postData.contents;
    var payload;

    if (raw.indexOf("payload=") === 0) {
      // form-urlencoded 格式：payload=<URL encoded JSON>
      payload = JSON.parse(decodeURIComponent(raw.slice("payload=".length)));
    } else {
      // 原始 JSON 格式
      payload = JSON.parse(raw);
    }

    var action = payload.action;
    var token  = payload.token;
    var data   = payload.data || {};

    // 🔒 Token 驗證
    if (!isValidToken(token)) {
      return createJsonResponse({ status: "error", message: "未授權的存取：無效的憑證" });
    }

    // 🔧 確保 Config 欄位結構是新版（只在第一次或舊版時執行）
    ensureConfigSchemaV3();

    // 📖 讀取類
    if (action === "getGroups")          return createJsonResponse({ status: "success", data: getExistingGroups() });
    if (action === "getTemplates")       return createJsonResponse({ status: "success", data: getTemplates() });
    if (action === "getPageConfig")      return createJsonResponse({ status: "success", data: getPageConfig(data.id) });
    if (action === "getAggregatedReport")return createJsonResponse({ status: "success", data: getAggregatedReport(data.type) });

    // ✍️ 寫入類
    if (action === "saveSheetData")      return createJsonResponse({ status: "success", message: saveSheetData(data) });
    if (action === "createGroup")        return createJsonResponse({ status: "success", message: createGroup(data).msg });
    if (action === "toggleGroupStatus")  return createJsonResponse({ status: "success", message: toggleGroupStatus(data).msg });
    if (action === "saveGroupPrompt")    return createJsonResponse({ status: "success", message: saveGroupPromptData(data).msg });
    if (action === "saveGroupMembers")   return createJsonResponse({ status: "success", message: saveGroupMembersData(data).msg });

    // 🤖 AI 排班
    if (action === "parseWithAI") {
      var result = callGeminiAPIWithRetry(data.text, data.headers, data.members, data.groupPrompt, data.template);
      return createJsonResponse({ status: "success", data: result });
    }

    return createJsonResponse({ status: "error", message: "無效的 POST 請求: 未知的 action -> " + action });

  } catch (err) {
    writeAuditLog("system", "doPost_error", { error: err.toString() });
    return createJsonResponse({ status: "error", message: err.toString() });
  }
}


// ============================================================
//  🤖 Gemini AI — 附重試機制與逾時保護
// ============================================================
function callGeminiAPIWithRetry(rawText, headers, members, groupPrompt, template) {
  var cfg = getConfig();

  // 驗證輸入
  if (!headers || !Array.isArray(headers) || headers.length === 0) {
    throw new Error("後端未接收到有效的表頭清單 (headers 為空)");
  }
  if (!cfg.GEMINI_API_KEY) {
    throw new Error("Gemini API Key 未設定，請至 GAS 指令碼屬性設定 GEMINI_API_KEY");
  }

  // 每日使用量計數
  var props   = PropertiesService.getScriptProperties();
  var today   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var countKey = "AI_USAGE_COUNT_" + today;
  var currentCount = parseInt(props.getProperty(countKey) || "0", 10);

  if (currentCount >= cfg.AI_DAILY_LIMIT) {
    throw new Error("今日 AI 使用次數已達上限 (" + cfg.AI_DAILY_LIMIT + " 次)，請明日再試。");
  }
  props.setProperty(countKey, (currentCount + 1).toString());

  // 組建 Prompt
  var systemPrompt = getSystemPromptTemplate(headers, members, groupPrompt, template);
  var requestPayload = {
    contents: [{
      parts: [{ text: "系統指令：" + systemPrompt + "\n\n使用者輸入：\n" + rawText }]
    }],
    generationConfig: { responseMimeType: "application/json" }
  };

  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(requestPayload),
    muteHttpExceptions: true
  };

  // 重試迴圈
  var lastError = "";
  for (var attempt = 1; attempt <= cfg.AI_MAX_RETRIES; attempt++) {
    try {
      var response = UrlFetchApp.fetch(cfg.GEMINI_API_URL, options);
      var code = response.getResponseCode();

      // 服務忙線（429）→ 等待後重試
      if (code === 429 || code === 503) {
        lastError = "伺服器忙線中 (HTTP " + code + ")";
        if (attempt < cfg.AI_MAX_RETRIES) {
          Utilities.sleep(cfg.AI_RETRY_DELAY * attempt);
          continue;
        }
        break;
      }

      if (code !== 200) {
        lastError = "HTTP " + code + ": " + response.getContentText().substring(0, 200);
        break;
      }

      var json = JSON.parse(response.getContentText());
      if (json.error) throw new Error("Gemini 回傳錯誤: " + json.error.message);

      var aiText = json.candidates[0].content.parts[0].text;
      try {
        var parsed = JSON.parse(aiText);
        writeAuditLog("system", "parseWithAI_success", { headers: headers, rowCount: parsed.length });
        return parsed;
      } catch (parseErr) {
        throw new Error("AI 回傳格式錯誤，請重試一次。(原始: " + aiText.substring(0, 100) + ")");
      }

    } catch (fetchErr) {
      lastError = fetchErr.toString();
      if (attempt < cfg.AI_MAX_RETRIES) {
        Utilities.sleep(cfg.AI_RETRY_DELAY * attempt);
      }
    }
  }

  throw new Error("AI 服務故障，已重試 " + cfg.AI_MAX_RETRIES + " 次。最後錯誤：" + lastError);
}


// ============================================================
//  📝 AI 系統提示模板（完整實作）
// ============================================================
function getSystemPromptTemplate(headers, members, groupPrompt) {
  var headersStr  = headers.join("、");
  var membersStr  = (members && members.length > 0) ? members.join("、") : "（未提供名單）";
  var customRules = groupPrompt ? "\n\n【小組專屬規則】\n" + groupPrompt : "";

  return [
    "你是一個教會排班助手，專門處理台灣教會小組的聚會服事排班。",
    "",
    "【任務說明】",
    "根據使用者輸入的文字（可能是排班條件描述，或是已打好的排班文字），",
    "輸出一個符合以下欄位結構的 JSON 陣列。",
    "",
    "【欄位結構】",
    "每筆資料必須包含這些欄位：" + headersStr,
    "- 日期欄位格式必須是 yyyy-MM-dd（例如 2026-05-04）",
    "- 沒有資訊的欄位請填入空字串 \"\"，不可省略欄位",
    "",
    "【可用人員名單】",
    membersStr,
    customRules,
    "",
    "【輸出規則】",
    "1. 只輸出純 JSON 陣列，不加任何說明文字或 markdown 格式",
    "2. 每個元素代表一個聚會場次",
    "3. 所有欄位鍵名必須與【欄位結構】完全一致（包含中文）",
    "4. 若使用者要求自動排班，請從可用人員名單中均衡分配，並遵守專屬規則",
    "5. 範例格式：[{\"日期\":\"2026-05-04\",\"主題\":\"信心的勇氣\",\"帶領\":\"小明\"}]"
  ].join("\n");
}


// ============================================================
//  📊 彙整報表（含快取優化）
// ============================================================
function getAggregatedReport(type) {
  var cfg   = getConfig();
  var cache = CacheService.getScriptCache();
  var cacheKey = "AGGREGATED_REPORT_" + type;

  // 嘗試從快取讀取
  var cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch(e) { /* 快取損壞，重新生成 */ }
  }

  // 重新生成報表
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var config = ss.getSheetByName('Config').getDataRange().getValues();

  var targetGroups = [];
  // 欄位：A=UUID(0), B=ID(1), C=名稱(2), D=模板(3), E=狀態(4)
  for (var i = 1; i < config.length; i++) {
    var gName   = config[i][2] ? config[i][2].toString().trim() : "";  // C欄：名稱
    var gTemp   = config[i][3] ? config[i][3].toString().trim() : "";  // D欄：模板
    var gStatus = config[i][4] ? config[i][4].toString().trim() : "";  // E欄：狀態

    if (!gName || gStatus !== "啟用") continue;

    if (type === "smallGroup" && gTemp === "小組聚會表模板") {
      targetGroups.push({ name: gName, template: gTemp });
    } else if (type === "others" && gTemp !== "小組聚會表模板" && gTemp !== "") {
      targetGroups.push({ name: gName, template: gTemp });
    }
  }

  var allHeaders = ["分頁名稱", "模板類型"];
  var rawData    = [];

  targetGroups.forEach(function(g) {
    var sheet = ss.getSheetByName(g.name);
    if (!sheet) return;
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    var headers = data[0].map(function(h) { return h.toString().trim(); });
    headers.forEach(function(h) {
      if (h && allHeaders.indexOf(h) === -1) allHeaders.push(h);
    });

    for (var r = 1; r < data.length; r++) {
      var row     = data[r];
      var isEmpty = true;
      var rowObj  = {};

      for (var c = 0; c < headers.length; c++) {
        var h   = headers[c];
        var val = row[c];
        if (val !== "") {
          isEmpty = false;
          if (Object.prototype.toString.call(val) === '[object Date]') {
            val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }
          rowObj[h] = String(val).replace(/[\r\n]+/g, " ");
        }
      }
      if (!isEmpty) {
        rawData.push({ groupName: g.name, template: g.template, dataObj: rowObj });
      }
    }
  });

  var matrix = [allHeaders];
  rawData.forEach(function(rowItem) {
    var finalRow = [];
    allHeaders.forEach(function(h) {
      if (h === "分頁名稱")   finalRow.push(rowItem.groupName);
      else if (h === "模板類型") finalRow.push(rowItem.template);
      else finalRow.push(rowItem.dataObj[h] || "");
    });
    matrix.push(finalRow);
  });

  // 按日期排序
  var dateIdx = allHeaders.indexOf("日期");
  if (dateIdx !== -1 && matrix.length > 1) {
    var headersRow = matrix.shift();
    matrix.sort(function(a, b) {
      return (a[dateIdx] || "9999-99-99").localeCompare(b[dateIdx] || "9999-99-99");
    });
    matrix.unshift(headersRow);
  }

  // 存入快取
  try {
    var serialized = JSON.stringify(matrix);
    if (serialized.length < 100000) { // 快取上限約 100KB
      cache.put(cacheKey, serialized, cfg.CACHE_REPORT_TTL);
    }
  } catch(e) { /* 資料太大，略過快取 */ }

  return matrix;
}


// ============================================================
//  📄 getPageConfig（含自動快取刷新）
//  以「小組編號 ID」(欄 B) 查找，與前端 URL 參數一致
// ============================================================
function getPageConfig(id) {
  var cfg    = getConfig();
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var config = ss.getSheetByName('Config').getDataRange().getValues();

  var targetName    = "";
  var templateName  = "";
  var groupPrompt   = "";
  var customMembers = [];

  // 欄位：A=UUID(0), B=ID(1), C=名稱(2), D=模板(3), E=狀態(4), F=規則(5), G=名單(6)
  for (var i = 1; i < config.length; i++) {
    if (config[i][1].toString().trim() === id) {          // B欄：小組編號
      targetName   = config[i][2].toString().trim();       // C欄：名稱
      templateName = config[i][3] ? config[i][3].toString().trim() : "";  // D欄
      groupPrompt  = config[i][5] ? config[i][5].toString().trim() : "";  // F欄
      var membersStr = config[i][6] ? config[i][6].toString().trim() : "[]"; // G欄
      try { customMembers = JSON.parse(membersStr); } catch(e) { customMembers = []; }
      break;
    }
  }
  if (!targetName) throw new Error("找不到 ID 對應的分頁：" + id);

  var sheet = ss.getSheetByName(targetName);
  if (!sheet) throw new Error("找不到分頁：" + targetName);

  var vals = sheet.getRange(1, 1, 50, 16).getValues();

  var result = {
    groupName:     targetName,
    template:      templateName,
    matrix:        vals.map(function(row) {
      return row.map(function(cell) {
        if (!cell) return "";
        if (Object.prototype.toString.call(cell) === '[object Date]')
          return Utilities.formatDate(cell, Session.getScriptTimeZone(), "yyyy-MM-dd");
        return String(cell).replace(/[\r\n]+/g, " ");
      });
    }),
    members:       [],
    coreMembers:   [],
    customMembers: customMembers,
    groupPrompt:   groupPrompt,
    autoRoleRules: "",
    eventData:     []
  };

  var cache = CacheService.getScriptCache();

  // ⚡ 聚會資料快取（自動刷新）
  var cachedEvents = cache.get("SYNC_GLOBAL_EVENTS");
  if (cachedEvents) {
    result.eventData = JSON.parse(cachedEvents);
  } else {
    syncAllExternalResources();
    var freshEvents = cache.get("SYNC_GLOBAL_EVENTS");
    result.eventData = freshEvents ? JSON.parse(freshEvents) : [];
  }

  // ⚡ 小組名單快取（小組聚會表模板才需要）
  if (templateName === "小組聚會表模板" || templateName === "團契聚會表模板") {
    var cachedMembers = cache.get("EXT_MEMBERS_" + id);
    if (cachedMembers) {
      var mData = JSON.parse(cachedMembers);
      result.members       = mData.members;
      result.coreMembers   = mData.coreMembers;
      result.autoRoleRules = mData.autoRoleRules;
    } else {
      fetchAndCacheMembers(id, result);
    }
  }

  return result;
}


// ============================================================
//  👥 外部名單抓取與快取
// ============================================================
function fetchAndCacheMembers(id, result) {
  var cfg = getConfig();
  try {
    var extSs      = SpreadsheetApp.openById(cfg.EXTERNAL_MEMBERS_SHEET_ID);
    var listSheet  = extSs.getSheetByName("小組清單");
    if (!listSheet) return;

    var listData   = listSheet.getDataRange().getValues();
    var extGroupName = "";

    for (var r = 0; r < listData.length; r++) {
      if (listData[r][2] && listData[r][2].toString().trim() === id) {
        extGroupName = listData[r][0].toString().trim();
        break;
      }
    }
    if (!extGroupName) return;

    var memberSheet = extSs.getSheetByName(extGroupName + "_名單");
    if (!memberSheet) return;

    var memberData       = memberSheet.getRange("A:C").getValues();
    var members          = [];  // AI 排班用（不含陪伴同工）
    var coreWorkers      = [];  // datalist 候選（含陪伴同工）
    var generalWorkers   = [];
    var companionWorkers = [];  // 陪伴同工獨立記錄

    for (var m = 1; m < memberData.length; m++) {
      var name   = memberData[m][0] ? memberData[m][0].toString().trim() : "";
      var status = memberData[m][2] ? memberData[m][2].toString().trim() : "";
      if (!name) continue;

      // 小羊：完全排除
      if (status === "小羊") continue;

      // 陪伴同工：加入 coreMembers（手動可選）但不加入 members（AI 排班名單）
      if (status === "陪伴同工") {
        companionWorkers.push(name);
        if (coreWorkers.indexOf(name) === -1) coreWorkers.push(name);
        continue;
      }

      // 核心同工 / 一般同工：正常加入 members
      if (members.indexOf(name) === -1) members.push(name);
      if (status === "核心同工")      coreWorkers.push(name);
      else if (status === "一般同工") generalWorkers.push(name);
    }

    var autoRules = "【系統自動判斷之名單權限】：\n";
    if (coreWorkers.length > 0)
      autoRules += "- 核心同工 (" + coreWorkers.join(", ") + ")：可排「破冰、敬拜、話語分享」。\n";
    if (generalWorkers.length > 0)
      autoRules += "- 一般同工 (" + generalWorkers.join(", ") + ")：僅可排「破冰、敬拜」。\n";
    if (companionWorkers.length > 0)
      autoRules += "- 陪伴同工 (" + companionWorkers.join(", ") + ")：具核心同工權限，不列入自動排班，僅供手動彈性調整。\n";
    autoRules += "- 小羊：不可列入排班。\n";

    result.members       = members;
    result.coreMembers   = coreWorkers;  // 含陪伴同工，供 datalist 手動選用
    result.autoRoleRules = autoRules;

    CacheService.getScriptCache().put(
      "EXT_MEMBERS_" + id,
      JSON.stringify({ members: members, coreMembers: coreWorkers, autoRoleRules: autoRules }),
      cfg.CACHE_MEMBERS_TTL
    );
  } catch (e) {
    console.log("外部名單抓取失敗：" + e.message);
  }
}


// ============================================================
//  📅 外部聚會資料同步（可手動或定時觸發）
// ============================================================
function syncAllExternalResources() {
  var cfg   = getConfig();
  var cache = CacheService.getScriptCache();

  try {
    var extEventSs = SpreadsheetApp.openById(cfg.EXTERNAL_EVENTS_SHEET_ID);
    var eventSheet = extEventSs.getSheetByName("聚會資料");
    if (!eventSheet) return;

    var eData  = eventSheet.getDataRange().getValues();
    if (eData.length === 0) return;

    var hRow = eData[0];
    var dIdx = hRow.indexOf("日期");
    var nIdx = hRow.indexOf("聚會名稱");
    var cIdx = hRow.indexOf("聚會類別");
    var events = [];

    for (var e = 1; e < eData.length; e++) {
      if (dIdx === -1 || !eData[e][dIdx]) continue;
      var dVal = eData[e][dIdx];
      if (Object.prototype.toString.call(dVal) === '[object Date]') {
        dVal = Utilities.formatDate(dVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
        dVal = String(dVal).trim();
      }
      if (!dVal) continue;
      events.push({
        date:     dVal,
        name:     (nIdx > -1 && eData[e][nIdx]) ? String(eData[e][nIdx]).trim() : "",
        category: (cIdx > -1 && eData[e][cIdx]) ? String(eData[e][cIdx]).trim() : ""
      });
    }

    cache.put("SYNC_GLOBAL_EVENTS", JSON.stringify(events), cfg.CACHE_EVENTS_TTL);
    console.log("✅ 全局聚會資源同步完成，共 " + events.length + " 筆");
  } catch (e) {
    console.log("同步聚會資料失敗：" + e.message);
  }
}

function clearMembersCache() {
  var cache = CacheService.getScriptCache();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = ss.getSheetByName('Config').getDataRange().getValues();
  
  // 清除所有小組的名單快取
  for (var i = 1; i < config.length; i++) {
    var id = config[i][1] ? config[i][1].toString().trim() : "";
    if (id) cache.remove("EXT_MEMBERS_" + id);
  }
  console.log("✅ 所有小組名單快取已清除，下次載入將重新從外部試算表抓取");
}

// ============================================================
//  💾 資料寫入操作
// ============================================================
function saveSheetData(payload) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(payload.groupName);
  if (!sheet) throw new Error("分頁不存在：" + payload.groupName);

  var data = payload.matrix;
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  // 清除此分頁相關的彙整快取（資料已更新）
  invalidateReportCache();

  writeAuditLog("system", "saveSheetData", { groupName: payload.groupName, rows: data.length });
  return "✅ 儲存成功！";
}

function createGroup(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName(data.template).copyTo(ss).setName(data.name);
  // A=UUID(自動產生), B=ID, C=名稱, D=模板, E=狀態, F=規則, G=名單
  var newUuid = Utilities.getUuid();
  ss.getSheetByName('Config').appendRow([newUuid, data.id, data.name, data.template, "啟用", "", "[]"]);
  writeAuditLog("system", "createGroup", { uuid: newUuid, id: data.id, name: data.name });
  return { msg: "建立成功" };
}

function toggleGroupStatus(data) {
  var s          = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  var configData = s.getDataRange().getValues();

  // 以小組編號 ID (欄 B, index 1) 查找；狀態在欄 E (index 4)
  for (var i = 1; i < configData.length; i++) {
    if (configData[i][1].toString().trim() === data.id) {
      var newStatus = (data.status === "啟用") ? "停用" : "啟用";
      s.getRange(i + 1, 5).setValue(newStatus);             // E欄
      writeAuditLog("system", "toggleGroupStatus", { id: data.id, status: newStatus });
      return { msg: "已設為" + newStatus };
    }
  }
  throw new Error("找不到該分頁 ID：" + data.id);
}

function saveGroupPromptData(data) {
  var s          = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  var configData = s.getDataRange().getValues();

  // 以小組編號 ID (欄 B, index 1) 查找；規則在欄 F (index 5)
  for (var i = 1; i < configData.length; i++) {
    if (configData[i][1].toString().trim() === data.id) {
      s.getRange(i + 1, 6).setValue(data.prompt);           // F欄
      writeAuditLog("system", "saveGroupPrompt", { id: data.id });
      return { msg: "規則儲存成功" };
    }
  }
  throw new Error("找不到該分頁 ID：" + data.id);
}

function saveGroupMembersData(data) {
  var s          = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  var configData = s.getDataRange().getValues();

  // 以小組編號 ID (欄 B, index 1) 查找；名單在欄 G (index 6)
  for (var i = 1; i < configData.length; i++) {
    if (configData[i][1].toString().trim() === data.id) {
      s.getRange(i + 1, 7).setValue(JSON.stringify(data.members));  // G欄

      // 清除此組的名單快取（下次讀取時重新抓取）
      CacheService.getScriptCache().remove("EXT_MEMBERS_" + data.id);

      writeAuditLog("system", "saveGroupMembers", { id: data.id, count: data.members.length });
      return { msg: "名單儲存成功" };
    }
  }
  throw new Error("找不到該分頁 ID：" + data.id);
}


// ============================================================
//  🔄 自動全域同步外部小組清單（UUID 核心版）
//
//  Config 欄位：A=UUID, B=小組編號ID, C=名稱, D=模板, E=狀態, F=規則, G=名單
//  外部清單欄位：A=UUID, B=名稱, C=狀態, D=小組編號ID
//
//  同步行為：
//    情況 A：外部有、本地沒有 → 自動建立分頁 + Config 列
//    情況 B：UUID 吻合 → 比對並同步 ID、名稱、狀態（名稱變動時同步分頁名稱）
//    防呆：外部缺 UUID 時自動補寫
// ============================================================
function autoSyncSmallGroups() {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName('Config');
  if (!configSheet) return;

  var configData = configSheet.getDataRange().getValues();

  // 1. 建立本地索引（以 UUID 為 key）
  //    欄位：A=UUID(0), B=ID(1), C=名稱(2), D=模板(3), E=狀態(4)
  var localByUuid = {};
  for (var i = 1; i < configData.length; i++) {
    var uuid = configData[i][0] ? configData[i][0].toString().trim() : "";
    if (!uuid) continue;
    localByUuid[uuid] = {
      rowIndex: i + 1,        // 在 Sheet 中的實際列號（1-based）
      id:       configData[i][1] ? configData[i][1].toString().trim() : "",
      name:     configData[i][2] ? configData[i][2].toString().trim() : "",
      template: configData[i][3] ? configData[i][3].toString().trim() : "",
      status:   configData[i][4] ? configData[i][4].toString().trim() : ""
    };
  }

  var cfg = getConfig();
  try {
    var extSs      = SpreadsheetApp.openById(cfg.EXTERNAL_MEMBERS_SHEET_ID);
    var listSheet  = extSs.getSheetByName("小組清單");
    if (!listSheet) return;

    var extData       = listSheet.getDataRange().getValues();
    var templateSheet = ss.getSheetByName("小組聚會表模板");

    // ⚠️ 外部「小組清單」實際欄位順序（以讀取到的資料為準）：
    //   A(0) = 小組名稱
    //   B(1) = 狀態（值為「顯示」或「隱藏」，對應本地的「啟用」/「停用」）
    //   C(2) = 小組編號 ID（如 VN01）
    //   D(3) = 建立日期
    //   E(4) = UUID
    //
    // 本地 Config 狀態值「啟用」對應外部「顯示」，「停用」對應其他值
    for (var r = 1; r < extData.length; r++) {
      var extName    = extData[r][0] ? extData[r][0].toString().trim() : "";
      var extRawStatus = extData[r][1] ? extData[r][1].toString().trim() : "";
      var extGroupId = extData[r][2] ? extData[r][2].toString().trim() : "";
      // D(3) = 建立日期，略過不使用
      var extUuid    = extData[r][4] ? extData[r][4].toString().trim() : "";

      // 將外部狀態「顯示」→「啟用」，其他（隱藏/空白）→「停用」
      var extStatus = (extRawStatus === "顯示") ? "啟用" : "停用";

      if (!extName) continue;

      // 🛡️ 防呆：外部沒有 UUID → 自動補寫至 E 欄（第 5 欄）
      if (!extUuid) {
        extUuid = Utilities.getUuid();
        listSheet.getRange(r + 1, 5).setValue(extUuid);
      }

      if (!localByUuid[extUuid]) {
        // ─── 情況 A：本地沒有此 UUID → 新建（僅當狀態為啟用時）───
        if (extStatus === "啟用" && templateSheet) {
          try {
            var newSheet = templateSheet.copyTo(ss);
            newSheet.setName(extName);
            // A=UUID, B=ID, C=名稱, D=模板, E=狀態, F=規則, G=名單
            configSheet.appendRow([extUuid, extGroupId, extName, "小組聚會表模板", "啟用", "", "[]"]);
            writeAuditLog("system", "autoSync_created", { uuid: extUuid, id: extGroupId, name: extName });
          } catch (copyErr) {
            console.log("autoSync 建立分頁失敗（可能已存在）: " + extName + " / " + copyErr);
          }
        }
      } else {
        // ─── 情況 B：UUID 吻合 → 比對差異並同步 ───
        var local = localByUuid[extUuid];
        var changed = false;

        // B欄(欄號2)：小組編號 ID 是否變動
        if (local.id !== extGroupId) {
          configSheet.getRange(local.rowIndex, 2).setValue(extGroupId);
          writeAuditLog("system", "autoSync_idChanged",
            { uuid: extUuid, oldId: local.id, newId: extGroupId });
          changed = true;
        }

        // C欄(欄號3)：名稱是否變動（同步更新實體分頁名稱）
        if (local.name !== extName) {
          var targetSheet = ss.getSheetByName(local.name);
          if (targetSheet) {
            try { targetSheet.setName(extName); } catch (renameErr) {
              console.log("分頁改名失敗: " + local.name + " → " + extName);
            }
          }
          configSheet.getRange(local.rowIndex, 3).setValue(extName);
          writeAuditLog("system", "autoSync_nameChanged",
            { uuid: extUuid, oldName: local.name, newName: extName });
          changed = true;
        }

        // E欄(欄號5)：狀態是否變動（外部「顯示」→ 本地「啟用」）
        if (local.status !== extStatus) {
          configSheet.getRange(local.rowIndex, 5).setValue(extStatus);
          writeAuditLog("system", "autoSync_statusChanged",
            { uuid: extUuid, name: extName, extRaw: extRawStatus, mapped: extStatus });
          changed = true;
        }

        if (changed) {
          // 同步後清除相關快取
          invalidateReportCache();
        }
      }
    }

    console.log("✅ autoSyncSmallGroups 完成");
  } catch (e) {
    console.log("autoSyncSmallGroups 失敗: " + e.toString());
  }
}


// ============================================================
//  📋 讀取操作
// ============================================================

/**
 * 讀取所有小組列表
 * 呼叫前先執行 autoSyncSmallGroups 確保資料一致
 * 回傳的 id 為「小組編號」(欄 B)，前端以此作為 URL 參數
 */
function getExistingGroups() {
  // 每次載入儀表板前先對齊外部資料
  autoSyncSmallGroups();

  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!s || s.getLastRow() < 2) return [];

  // 欄位：A=UUID(0), B=ID(1), C=名稱(2), D=模板(3), E=狀態(4)
  return s.getDataRange().getValues().slice(1).map(function(r) {
    return {
      id:       r[1] ? r[1].toString().trim() : "",   // 小組編號，前端 URL 用
      name:     r[2] ? r[2].toString().trim() : "",
      template: r[3] ? r[3].toString().trim() : "其他",
      status:   r[4] ? r[4].toString().trim() : "啟用"
    };
  }).filter(function(i) { return i.id; });
}

function getTemplates() {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('模板名稱');
  if (!s || s.getLastRow() < 2) return [];
  return s.getRange(2, 1, s.getLastRow() - 1, 1).getValues()
    .map(function(r) { return r[0].toString().trim(); })
    .filter(String);
}


// ============================================================
//  🗑️ 快取清除輔助
// ============================================================
function invalidateReportCache() {
  var cache = CacheService.getScriptCache();
  cache.remove("AGGREGATED_REPORT_smallGroup");
  cache.remove("AGGREGATED_REPORT_others");
}


// ============================================================
//  📋 審計日誌
// ============================================================
function writeAuditLog(userId, action, details) {
  try {
    var ss       = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName('審計日誌');

    // 如果沒有日誌分頁，自動建立
    if (!logSheet) {
      logSheet = ss.insertSheet('審計日誌');
      logSheet.appendRow(["時間", "操作者", "動作", "詳細內容"]);
      logSheet.setFrozenRows(1);
    }

    logSheet.appendRow([
      new Date(),
      userId || "system",
      action,
      JSON.stringify(details)
    ]);
  } catch (e) {
    // 日誌寫入失敗不影響主流程
    console.log("審計日誌寫入失敗：" + e.message);
  }
}


// ============================================================
//  ⏰ 定時觸發設定
//
//  用法（在 GAS 編輯器執行一次）：
//    1. 打開 core.gs
//    2. 選擇函式 setupWeeklySync
//    3. 按執行 ▶️
//    4. 授權一次
//    5. 之後每週五 3:00 AM 自動執行 autoSyncSmallGroups()
//
//  檢查觸發器：Google Apps Script 編輯器左側「觸發器」面板
//  取消觸發器：在同一面板點右上角的三點選單 → 刪除
// ============================================================

/**
 * 🛠️ 一次性設定函式：建立「每週五凌晨 3 點」的定時觸發器
 * 在 GAS 編輯器中手動執行一次即可
 */
function setupWeeklySync() {
  // 刪除舊的同名觸發器（如果存在）
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'autoSyncSmallGroups') {
      ScriptApp.deleteTrigger(triggers[i]);
      console.log("已刪除舊觸發器");
    }
  }

  // 建立新觸發器：每週五凌晨 3 點
  ScriptApp.newTrigger('autoSyncSmallGroups')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)  // 週五
    .atHour(3)                             // 凌晨 3 點
    .create();

  console.log("✅ 定時觸發器已設定：每週五凌晨 3:00 AM 執行 autoSyncSmallGroups()");
  console.log("檢查觸發器：Google Apps Script 編輯器左側「觸發器」面板");
}

/**
 * 🧪 手動測試函式：立即執行一次同步
 * 用於測試同步邏輯是否正常
 */
function testSyncNow() {
  console.log("🧪 開始手動同步測試...");
  autoSyncSmallGroups();
  console.log("✅ 同步測試完成，請檢查審計日誌");
}


// ============================================================
//  🛠️ 工具函式
// ============================================================
function createJsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}