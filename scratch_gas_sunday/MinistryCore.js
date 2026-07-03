/**
 * MinistryCore.js — 教會服事管理（整合至主 GAS 版本）
 *
 * 整合自原 事工管理_GAS/core.js
 * 主要變化：
 *   ① getActiveSpreadsheet() → getMinistrySS()（用 openById）
 *   ② fetchAndCacheMembers 改用 getCachedMembers()，依 所屬小組 + 身分 過濾
 *      → 完全消除跨 SS openById 對小組試算表的依賴
 *      → 主日為單一真實來源，不再有同步落差
 *   ③ syncAllExternalResources 仍 openById 教會行事曆 SS（讀取，未整合）
 *   ④ 共用 GeminiHelper.callGemini，刪除重複的 callGeminiAPIWithRetry
 *   ⑤ getExistingGroups 加 cache（4h TTL，寫入時清除）
 *   ⑥ autoSyncSmallGroups 不再每次 dashboard 都跑（改由定時 trigger）
 *   ⑦ writeAuditLog 改批次寫入（PropertiesService 暫存 + flushAuditLog）
 *   ⑧ getRange("A:C") 修正為精確範圍
 *
 * 路由：所有 action 名稱以 ministry_ 前綴呼叫，避免與其他子系統衝突
 */

// ── 常數區 ───────────────────────────────────────────────────
const MINISTRY_SHEET_ID = '1Q-_GnXaUhhUnLhVJC1w6o0ty-XKTqzkPyaRFLArfLew'; // 依據使用者提供之正式/測試表 ID
const CALENDAR_SHEET_ID = '1tKI5k7HwI9S2bTRV6RuKrzxBFGzROYHytsFYHuH8H7E'; // 教會行事曆 prod（事工只讀）

// Cache key
const MIN_GROUPS_CACHE_KEY     = 'MIN_GROUPS_V1';
const MIN_GROUPS_CACHE_TTL     = 14400;   // 4 小時
const MIN_EVENTS_CACHE_KEY     = 'MIN_EVENTS_V1';
const MIN_EVENTS_CACHE_TTL     = 14400;   // 4 小時
const MIN_REPORT_CACHE_PREFIX  = 'MIN_REPORT_V3_';
const MIN_REPORT_CACHE_TTL     = 14400;   // 4 小時

const MINISTRY_INITIAL_FIELD_TEMPLATES = {
  "聚會型模板": {
    defaultFields: ["日期", "主題", "經文", "地點", "敬拜", "話語分享"],
    requiredFields: ["日期"]
  },
  "事工型模板": {
    defaultFields: ["日期", "地點"],
    requiredFields: ["日期"]
  }
};

function _ministryNormalizeScheduleMode(mode) {
  return mode === "membersOnly" ? "membersOnly" : "schedule";
}

function _ministryRequiresSchedule(templateName, pageFieldConfig) {
  if (templateName === "小組聚會表模板" || templateName === "團契聚會表模板") return true;
  return _ministryNormalizeScheduleMode(pageFieldConfig && pageFieldConfig.scheduleMode) !== "membersOnly";
}

function _ministryFieldTemplateType(templateName) {
  if (templateName === "小組聚會表模板" || templateName === "團契聚會表模板" || templateName === "聚會型模板") {
    return "聚會型模板";
  }
  return "事工型模板";
}

function _ministryDefaultPageFieldConfig(pageId, templateName) {
  const fieldTemplateType = _ministryFieldTemplateType(templateName);
  const template = MINISTRY_INITIAL_FIELD_TEMPLATES[fieldTemplateType] || MINISTRY_INITIAL_FIELD_TEMPLATES["事工型模板"];
  return {
    pageId: pageId || "",
    fieldTemplateType: fieldTemplateType,
    scheduleMode: "schedule",
    fields: template.defaultFields.map(name => ({
      name: name,
      enabled: true,
      custom: false
    })),
    requiredFields: template.requiredFields.slice(),
    customFields: [],
    updatedAt: new Date().toISOString()
  };
}

function _ministryNormalizePageFieldConfig(config, pageId, templateName) {
  const fallback = _ministryDefaultPageFieldConfig(pageId, templateName);
  if (!config || typeof config !== "object") return fallback;

  const fieldTemplateType = config.fieldTemplateType || fallback.fieldTemplateType;
  const template = MINISTRY_INITIAL_FIELD_TEMPLATES[fieldTemplateType] || MINISTRY_INITIAL_FIELD_TEMPLATES["事工型模板"];
  const requiredFields = Array.from(new Set([].concat(config.requiredFields || [], template.requiredFields)));
  const sourceFields = Array.isArray(config.fields) && config.fields.length ? config.fields : fallback.fields;
  const seen = {};
  const fields = [];

  sourceFields.forEach(field => {
    const name = typeof field === "string" ? field : field && field.name;
    if (!name || seen[name]) return;
    seen[name] = true;
    fields.push({
      name: name,
      enabled: requiredFields.indexOf(name) !== -1 ? true : !(field && field.enabled === false),
      custom: !!(field && field.custom)
    });
  });

  requiredFields.forEach(name => {
    if (!seen[name]) {
      fields.unshift({ name: name, enabled: true, custom: false });
      seen[name] = true;
    }
  });

  return {
    pageId: pageId || "",
    fieldTemplateType: fieldTemplateType,
    scheduleMode: _ministryNormalizeScheduleMode(config.scheduleMode || fallback.scheduleMode),
    fields: fields,
    requiredFields: requiredFields,
    customFields: fields.filter(f => f.custom).map(f => f.name),
    updatedAt: new Date().toISOString()
  };
}

// ── 試算表存取（請求內快取）─────────────────────────────────
let _ministrySsCache = null;
let _ministryCalendarSsCache = null;
function getMinistrySS() {
  if (!_ministrySsCache) _ministrySsCache = SpreadsheetApp.openById(MINISTRY_SHEET_ID);
  return _ministrySsCache;
}
function getCalendarSS() {
  if (!_ministryCalendarSsCache) _ministryCalendarSsCache = SpreadsheetApp.openById(CALENDAR_SHEET_ID);
  return _ministryCalendarSsCache;
}

// 請求內 Config sheet 資料快取（避免同次請求重複 getDataRange）
let _configDataCache = null;
function _getConfigData() {
  if (_configDataCache) return _configDataCache;
  const sheet = getMinistrySS().getSheetByName('Config');
  if (!sheet) return [];
  _configDataCache = sheet.getDataRange().getValues();
  return _configDataCache;
}
function _invalidateConfigDataCache() { _configDataCache = null; }

// ── Config 結構：
//   A=UUID(0) B=ID(1) C=名稱(2) D=模板(3) E=狀態(4) F=規則(5)
//   G=名單(6) H=講道設定(7) I=pageFieldConfig(8)
function ensureConfigSchemaV3() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty("CONFIG_SCHEMA_V3") === "true") return;

  const ss = getMinistrySS();
  const configSheet = ss.getSheetByName('Config');
  if (!configSheet || configSheet.getLastRow() < 2) {
    props.setProperty("CONFIG_SCHEMA_V3", "true");
    return;
  }

  const firstA = configSheet.getRange(2, 1).getValue().toString().trim();
  const looksLikeUUID = firstA.indexOf("-") !== -1 && firstA.length > 20;
  if (looksLikeUUID) {
    props.setProperty("CONFIG_SCHEMA_V3", "true");
    return;
  }

  // 舊版遷移：插入 UUID 欄
  const lastRow = configSheet.getLastRow();
  configSheet.insertColumnBefore(1);
  const newUuids = [];
  for (let row = 1; row <= lastRow; row++) newUuids.push([Utilities.getUuid()]);
  configSheet.getRange(1, 1, lastRow, 1).setValues(newUuids);
  invalidateMinistryReportCache();
  props.setProperty("CONFIG_SCHEMA_V3", "true");
  _enqueueAuditLog("system", "configMigrated", { rows: lastRow });
}

// ═══════════════════════════════════════════════════════════
//  📋 讀取類 API
// ═══════════════════════════════════════════════════════════

/**
 * 取得小組列表（cache-first）
 * 注意：不再每次都跑 autoSyncSmallGroups（改由定時 trigger）
 */
function ministry_getGroups() {
  ensureConfigSchemaV3();

  const cache = CacheService.getScriptCache();
  const raw = cache.get(MIN_GROUPS_CACHE_KEY);
  if (raw) {
    try { return JSON.parse(raw); } catch (e) { /* fall through */ }
  }

  const data = _getConfigData();
  if (data.length < 2) return [];

  const groups = data.slice(1).map(r => ({
    id:       r[1] ? String(r[1]).trim() : "",
    name:     r[2] ? String(r[2]).trim() : "",
    template: r[3] ? String(r[3]).trim() : "其他",
    status:   r[4] ? String(r[4]).trim() : "啟用"
  })).filter(g => g.id);

  try { cache.put(MIN_GROUPS_CACHE_KEY, JSON.stringify(groups), MIN_GROUPS_CACHE_TTL); } catch (e) {}
  return groups;
}

function ministry_getTemplates() {
  const sheet = getMinistrySS().getSheetByName('模板名稱');
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
    .map(r => r[0].toString().trim())
    .filter(String);
}

/**
 * 單一小組的頁面設定（含名單與聚會資料）
 *  - 名單：直接從主日 getCachedMembers 過濾此組（單一真實來源）
 *  - 聚會資料：從教會行事曆 SS 讀（cache-first）
 */
function ministry_getPageConfig(id, autoCreate) {
  id = decryptGroupCode(id);
  ensureConfigSchemaV3();

  const data = _getConfigData();
  let targetName = "", templateName = "", groupPrompt = "", customMembers = [], pageFieldConfig = null;
  let sermonSettings = { useSermon: false, sermonType: "華語/聯合" };

  for (let i = 1; i < data.length; i++) {
    if (data[i][1].toString().trim() === id) {
      targetName   = data[i][2].toString().trim();
      templateName = data[i][3] ? data[i][3].toString().trim() : "";
      groupPrompt  = data[i][5] ? data[i][5].toString().trim() : "";
      const membersStr = data[i][6] ? data[i][6].toString().trim() : "[]";
      try { customMembers = JSON.parse(membersStr); } catch (e) { customMembers = []; }

      // Read sermonSettings from column H (index 7) if it exists
      if (data[i].length > 7 && data[i][7]) {
        try {
          sermonSettings = JSON.parse(data[i][7].toString().trim());
        } catch (e) {
          sermonSettings = { useSermon: false, sermonType: "華語/聯合" };
        }
      }

      if (data[i].length > 8 && data[i][8]) {
        try {
          pageFieldConfig = JSON.parse(data[i][8].toString().trim());
        } catch (e) {
          pageFieldConfig = null;
        }
      }
      break;
    }
  }
  if (!targetName) {
    console.log("targetName not found for id: " + id + ", autoCreate: " + autoCreate + " (type: " + typeof autoCreate + ")");
    if (autoCreate === true || autoCreate === "true" || autoCreate === 1 || autoCreate === "1") {
      const groupInfo = findGroupByCode(id);
      console.log("findGroupByCode result for " + id + ": " + JSON.stringify(groupInfo));
      if (groupInfo && groupInfo.success && groupInfo.groupName) {
        const defaultTmplName = "小組聚會表模板";
        const defaultTmplType = "聚會型模板";
        const templateConfig = MINISTRY_INITIAL_FIELD_TEMPLATES[defaultTmplType];
        const firstConfig = _ministryNormalizePageFieldConfig({
          fields: templateConfig.defaultFields.map(name => ({ name: name, enabled: true, custom: false })),
          requiredFields: templateConfig.requiredFields,
          fieldTemplateType: defaultTmplType
        }, id, defaultTmplName);

        ministry_createGroup({
          id: id,
          name: groupInfo.groupName,
          template: defaultTmplName,
          fieldTemplateType: defaultTmplType,
          pageFieldConfig: firstConfig
        });

        // 遞迴重新讀取剛剛建立的設定 (autoCreate 傳入 false 避免潛在無窮迴圈)
        return ministry_getPageConfig(id, false);
      }
    }
    throw new Error("找不到 ID 對應的分頁：" + id);
  }

  const sheet = getMinistrySS().getSheetByName(targetName);
  if (!sheet) throw new Error("找不到分頁：" + targetName);

  pageFieldConfig = _ministryNormalizePageFieldConfig(pageFieldConfig, id, templateName);
  const enabledFields = pageFieldConfig.fields.filter(f => f.enabled !== false).map(f => f.name);
  const readColCount = Math.max(16, enabledFields.length);
  const lastRow = Math.max(50, sheet.getLastRow());
  const vals = sheet.getRange(1, 1, lastRow, readColCount).getValues();
  const result = {
    groupName:     targetName,
    template:      templateName,
    matrix:        vals.map(row => row.map(cell => {
      if (!cell) return "";
      if (Object.prototype.toString.call(cell) === '[object Date]') {
        return Utilities.formatDate(cell, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      return String(cell).replace(/[\r\n]+/g, " ");
    })),
    members:       [],
    coreMembers:   [],
    customMembers: customMembers,
    groupPrompt:   groupPrompt,
    autoRoleRules: "",
    eventData:     [],
    sermonSettings: sermonSettings,
    scheduleMode:   pageFieldConfig.scheduleMode,
    pageFieldConfig: pageFieldConfig
  };

  // 聚會資料：cache-first + 動態過濾
  const allEvents = ministry_getEvents();
  let filteredEvents = [];
  const activeSermonType = (sermonSettings && sermonSettings.sermonType) ? sermonSettings.sermonType : "";
  if (activeSermonType.indexOf('台語') !== -1) {
    filteredEvents = allEvents.filter(e => e.category.indexOf('台語') !== -1 || e.category.indexOf('聯合') !== -1);
  } else if (activeSermonType.indexOf('華語') !== -1) {
    filteredEvents = allEvents.filter(e => e.category.indexOf('華語') !== -1 || e.category.indexOf('聯合') !== -1);
  } else if (targetName.indexOf('台語') !== -1) {
    filteredEvents = allEvents.filter(e => e.category.indexOf('台語') !== -1 || e.category.indexOf('聯合') !== -1);
  } else {
    // 預設 (華語) 或分頁有名稱匹配
    filteredEvents = allEvents.filter(e => e.category.indexOf('華語') !== -1 || e.category.indexOf('聯合') !== -1);
  }
  result.eventData = filteredEvents;

  // 名單：小組/團契聚會表才需要 — 直接從主日 cache 取此組成員
  if (templateName === "小組聚會表模板" || templateName === "團契聚會表模板") {
    _attachGroupMembersFromMaster(targetName, result);
  }

  return result;
}

/**
 * 從主日 cache 取此組成員，建立 members/coreMembers/autoRoleRules
 * （取代原 fetchAndCacheMembers — 不再 openById 跨小組試算表）
 */
function _attachGroupMembersFromMaster(groupName, result) {
  try {
    const allMembers = getCachedMembers();
    // 支援多組格式：用 memberInGroup + getRoleForGroup 過濾並取出此組的身分
    const groupMembers = allMembers.filter(m => memberInGroup(m[8], groupName));

    const members = [];           // AI 排班用（不含小羊、不含陪伴同工）
    const coreWorkers = [];       // datalist 候選（含陪伴同工）
    const generalWorkers = [];
    const companionWorkers = [];

    groupMembers.forEach(m => {
      const name = m[0] ? String(m[0]).trim() : "";
      const role = getRoleForGroup(m[9], groupName);
      if (!name) return;

      if (role === "小羊") return;          // 完全排除

      if (role === "陪伴同工") {
        companionWorkers.push(name);
        if (coreWorkers.indexOf(name) === -1) coreWorkers.push(name);
        return;
      }

      if (members.indexOf(name) === -1) members.push(name);
      if (role === "核心同工")        coreWorkers.push(name);
      else if (role === "一般同工")   generalWorkers.push(name);
    });

    let autoRules = "【系統自動判斷之名單權限】：\n";
    if (coreWorkers.length > 0)
      autoRules += "- 核心同工 (" + coreWorkers.join(", ") + ")：可排「破冰、敬拜、話語分享」。\n";
    if (generalWorkers.length > 0)
      autoRules += "- 一般同工 (" + generalWorkers.join(", ") + ")：僅可排「破冰、敬拜」。\n";
    if (companionWorkers.length > 0)
      autoRules += "- 陪伴同工 (" + companionWorkers.join(", ") + ")：具核心同工權限，不列入自動排班。\n";
    autoRules += "- 小羊：不可列入排班。\n";

    result.members       = members;
    result.coreMembers   = coreWorkers;
    result.autoRoleRules = autoRules;
  } catch (e) {
    console.log("讀取主日名單失敗：" + e.message);
  }
}

/**
 * 從教會行事曆 SS 取聚會資料（cache-first）
 */
function ministry_getEvents() {
  const cache = CacheService.getScriptCache();
  const raw = cache.get(MIN_EVENTS_CACHE_KEY);
  if (raw) {
    try { return JSON.parse(raw); } catch (e) { /* fall through */ }
  }
  return _rebuildMinistryEventsCache();
}

function _rebuildMinistryEventsCache() {
  try {
    const calSs = getCalendarSS();
    const typeSheet = calSs.getSheetByName("事項類型");
    const eventSheet = calSs.getSheetByName("事項");
    if (!typeSheet || !eventSheet) return [];
    
    // 讀取所有事項類型
    const tData = typeSheet.getDataRange().getValues();
    if (tData.length < 2) return [];
    const tHeaders = tData[0];
    const typeById = {};
    const types = [];
    for (let i = 1; i < tData.length; i++) {
      const obj = {};
      tHeaders.forEach((h, idx) => obj[h] = tData[i][idx]);
      types.push(obj);
      typeById[obj.typeId] = obj;
    }
    
    // 找出「講道資訊」的子節點
    const sermonRoot = types.find(t => !t.parentTypeId && t['名稱'] === '講道資訊');
    if (!sermonRoot) return [];
    const sermonSubIds = new Set(types.filter(t => t.parentTypeId === sermonRoot.typeId).map(t => t.typeId));

    // 讀取欄位定義，映射 fieldId 到顯示名稱
    const fieldSheet = calSs.getSheetByName("欄位定義");
    const fieldById = {};
    if (fieldSheet && fieldSheet.getLastRow() > 1) {
      const fData = fieldSheet.getDataRange().getValues();
      const fHeaders = fData[0];
      const fIdIdx = fHeaders.indexOf("fieldId");
      const fNameIdx = fHeaders.indexOf("顯示名稱");
      if (fIdIdx !== -1 && fNameIdx !== -1) {
        for (let i = 1; i < fData.length; i++) {
          const fid = fData[i][fIdIdx];
          const fname = fData[i][fNameIdx];
          if (fid) fieldById[fid] = fname;
        }
      }
    }

    // 讀取事項欄位值，映射 eventId 到欄位值字典
    const valueSheet = calSs.getSheetByName("事項欄位值");
    const valuesByEvent = {};
    if (valueSheet && valueSheet.getLastRow() > 1) {
      const vData = valueSheet.getDataRange().getValues();
      const vHeaders = vData[0];
      const evIdIdx = vHeaders.indexOf("eventId");
      const fIdIdx = vHeaders.indexOf("fieldId");
      const valIdx = vHeaders.indexOf("值");
      if (evIdIdx !== -1 && fIdIdx !== -1 && valIdx !== -1) {
        for (let i = 1; i < vData.length; i++) {
          const evId = vData[i][evIdIdx];
          const fId = vData[i][fIdIdx];
          const val = vData[i][valIdx];
          if (evId && fId) {
            if (!valuesByEvent[evId]) valuesByEvent[evId] = {};
            const fieldName = fieldById[fId];
            if (fieldName) {
              valuesByEvent[evId][fieldName] = val;
            }
          }
        }
      }
    }
    
    // 讀取所有事項
    const eData = eventSheet.getDataRange().getValues();
    if (eData.length < 2) return [];
    const eHeaders = eData[0];
    const events = [];
    for (let i = 1; i < eData.length; i++) {
      const obj = {};
      eHeaders.forEach((h, idx) => obj[h] = eData[i][idx]);
      
      // 只保留屬於講道資訊的事項
      if (!sermonSubIds.has(obj.typeId)) continue;
      
      let dVal = obj['日期'];
      if (Object.prototype.toString.call(dVal) === '[object Date]') {
        dVal = Utilities.formatDate(dVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
        dVal = String(dVal).trim().substring(0, 10);
      }
      if (!dVal || dVal.length < 10) continue;
      
      const typeName = typeById[obj.typeId] ? typeById[obj.typeId]['名稱'] : "";
      const title = obj['顯示標題'] ? String(obj['顯示標題']).trim() : typeName;

      const evVals = valuesByEvent[obj.eventId] || {};
      const sermonTitle = evVals['講題'] || '';
      const sermonSpeaker = evVals['講員'] || '';
      const sermonScripture = evVals['經文'] || '';
      
      events.push({
        date: dVal,
        name: title,
        category: typeName, // e.g. "台語", "華語", "聯合"
        sermons: [{
          type: typeName,
          title: sermonTitle,
          speaker: sermonSpeaker,
          scripture: sermonScripture
        }]
      });
    }
    
    try { CacheService.getScriptCache().put(MIN_EVENTS_CACHE_KEY, JSON.stringify(events), MIN_EVENTS_CACHE_TTL); } catch (e) {}
    return events;
  } catch (e) {
    console.log("讀取聚會資料失敗：" + e.message);
    return [];
  }
}

/**
 * 彙整報表（cache-first）
 */
function ministry_getAggregatedReport(type) {
  const cache = CacheService.getScriptCache();
  const cacheKey = MIN_REPORT_CACHE_PREFIX + type;
  const cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { /* regen */ }
  }

  const ss = getMinistrySS();
  const config = _getConfigData();
  const targetGroups = [];

  for (let i = 1; i < config.length; i++) {
    const gName   = config[i][2] ? config[i][2].toString().trim() : "";
    const gTemp   = config[i][3] ? config[i][3].toString().trim() : "";
    const gStatus = config[i][4] ? config[i][4].toString().trim() : "";
    if (!gName || gStatus !== "啟用") continue;

    let pageFieldConfig = {};
    if (config[i].length > 8 && config[i][8]) {
      try { pageFieldConfig = JSON.parse(config[i][8].toString().trim()); } catch (e) { pageFieldConfig = {}; }
    }
    const requiresSchedule = _ministryRequiresSchedule(gTemp, pageFieldConfig);
    if (type === "smallGroup" && (gTemp === "小組聚會表模板" || gTemp === "團契聚會表模板")) {
      targetGroups.push({ name: gName, template: gTemp });
    } else if (type === "others" && gTemp !== "小組聚會表模板" && gTemp !== "團契聚會表模板" && gTemp !== "" && requiresSchedule) {
      targetGroups.push({ name: gName, template: gTemp });
    }
  }

  let allHeaders = ["分頁名稱", "模板類型"];
  const rawData = [];

  targetGroups.forEach(g => {
    const sheet = ss.getSheetByName(g.name);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    const headers = data[0].map(h => h.toString().trim());
    headers.forEach(h => { if (h && allHeaders.indexOf(h) === -1) allHeaders.push(h); });

    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      let isEmpty = true;
      const rowObj = {};
      for (let c = 0; c < headers.length; c++) {
        const h = headers[c];
        let val = row[c];
        if (val !== "") {
          isEmpty = false;
          if (Object.prototype.toString.call(val) === '[object Date]') {
            val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }
          rowObj[h] = String(val).replace(/[\r\n]+/g, " ");
        }
      }
      if (!isEmpty) rawData.push({ groupName: g.name, template: g.template, dataObj: rowObj });
    }
  });

  const matrix = [allHeaders];
  rawData.forEach(rowItem => {
    const finalRow = [];
    allHeaders.forEach(h => {
      if (h === "分頁名稱")        finalRow.push(rowItem.groupName);
      else if (h === "模板類型")   finalRow.push(rowItem.template);
      else                          finalRow.push(rowItem.dataObj[h] || "");
    });
    matrix.push(finalRow);
  });

  // 按日期排序
  const dateIdx = allHeaders.indexOf("日期");
  if (dateIdx !== -1 && matrix.length > 1) {
    const headersRow = matrix.shift();
    matrix.sort((a, b) => (a[dateIdx] || "9999-99-99").localeCompare(b[dateIdx] || "9999-99-99"));
    matrix.unshift(headersRow);
  }

  try {
    const serialized = JSON.stringify(matrix);
    if (serialized.length < 90000) cache.put(cacheKey, serialized, MIN_REPORT_CACHE_TTL);
  } catch (e) {}
  return matrix;
}

// ═══════════════════════════════════════════════════════════
//  ✍️ 寫入類 API（寫入後清相關 cache）
// ═══════════════════════════════════════════════════════════

function ministry_saveSheetData(payload) {
  const sheet = getMinistrySS().getSheetByName(payload.groupName);
  if (!sheet) throw new Error("分頁不存在：" + payload.groupName);
  const data = payload.matrix;
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  invalidateMinistryReportCache();
  firebaseInvalidate(['ministry_getAggregatedReport', 'ministry_getPageConfig']);
  _enqueueAuditLog("system", "saveSheetData", { groupName: payload.groupName, rows: data.length });
  return "✅ 儲存成功！";
}

function ministry_createGroup(data) {
  const ss = getMinistrySS();
  const pageFieldConfig = _ministryNormalizePageFieldConfig(
    data.pageFieldConfig,
    data.id,
    data.fieldTemplateType || data.template
  );
  const enabledFields = pageFieldConfig.fields.filter(f => f.enabled !== false).map(f => f.name);
  const tmpl = ss.getSheetByName(data.template);
  let newSheet;
  if (tmpl) {
    newSheet = tmpl.copyTo(ss).setName(data.name);
    if (enabledFields.length) {
      newSheet.getRange(1, 1, 1, enabledFields.length).setValues([enabledFields]);
    }
  } else {
    newSheet = ss.insertSheet(data.name);
    newSheet.getRange(1, 1, 1, enabledFields.length).setValues([enabledFields]);
  }
  const newUuid = Utilities.getUuid();
  ss.getSheetByName('Config').appendRow([
    newUuid,
    data.id,
    data.name,
    data.template || pageFieldConfig.fieldTemplateType,
    "啟用",
    "",
    "[]",
    "",
    JSON.stringify(pageFieldConfig)
  ]);
  _invalidateConfigDataCache();
  _invalidateMinistryGroupsCache();
  firebaseInvalidate(['ministry_getGroups', 'ministry_getAggregatedReport']);
  _enqueueAuditLog("system", "createGroup", { uuid: newUuid, id: data.id, name: data.name });
  return { msg: "建立成功" };
}

function ministry_savePageFieldConfig(data) {
  data.id = decryptGroupCode(data.id);
  const s = getMinistrySS().getSheetByName('Config');
  const configData = s.getDataRange().getValues();
  for (let i = 1; i < configData.length; i++) {
    if (configData[i][1].toString().trim() === data.id) {
      const templateName = configData[i][3] ? configData[i][3].toString().trim() : "";
      const pageFieldConfig = _ministryNormalizePageFieldConfig(data.pageFieldConfig, data.id, templateName);
      s.getRange(i + 1, 9).setValue(JSON.stringify(pageFieldConfig));
      _invalidateConfigDataCache();
      invalidateMinistryReportCache();
      firebaseInvalidate(['ministry_getPageConfig', 'ministry_getAggregatedReport', 'memberStatus_getMembers', 'memberStatus_getProfile', 'memberStatus_getServiceIndex']);
      _enqueueAuditLog("system", "savePageFieldConfig", { id: data.id, fields: pageFieldConfig.fields.length });
      return { msg: "欄位設定已儲存" };
    }
  }
  throw new Error("找不到該分頁 ID：" + data.id);
}

function ministry_toggleGroupStatus(data) {
  data.id = decryptGroupCode(data.id);
  const s = getMinistrySS().getSheetByName('Config');
  const configData = s.getDataRange().getValues();
  for (let i = 1; i < configData.length; i++) {
    if (configData[i][1].toString().trim() === data.id) {
      const newStatus = (data.status === "啟用") ? "停用" : "啟用";
      s.getRange(i + 1, 5).setValue(newStatus);
      _invalidateConfigDataCache();
      _invalidateMinistryGroupsCache();
      firebaseInvalidate(['ministry_getGroups']);
      _enqueueAuditLog("system", "toggleGroupStatus", { id: data.id, status: newStatus });
      return { msg: "已設為" + newStatus };
    }
  }
  throw new Error("找不到該分頁 ID：" + data.id);
}

function ministry_saveGroupPrompt(data) {
  data.id = decryptGroupCode(data.id);
  const s = getMinistrySS().getSheetByName('Config');
  const configData = s.getDataRange().getValues();
  for (let i = 1; i < configData.length; i++) {
    if (configData[i][1].toString().trim() === data.id) {
      s.getRange(i + 1, 6).setValue(data.prompt);
      _invalidateConfigDataCache();
      firebaseInvalidate(['ministry_getPageConfig']);
      _enqueueAuditLog("system", "saveGroupPrompt", { id: data.id });
      return { msg: "規則儲存成功" };
    }
  }
  throw new Error("找不到該分頁 ID：" + data.id);
}

function ministry_saveGroupMembers(data) {
  data.id = decryptGroupCode(data.id);
  const s = getMinistrySS().getSheetByName('Config');
  const configData = s.getDataRange().getValues();
  for (let i = 1; i < configData.length; i++) {
    if (configData[i][1].toString().trim() === data.id) {
      s.getRange(i + 1, 7).setValue(JSON.stringify(data.members));
      _invalidateConfigDataCache();
      firebaseInvalidate(['ministry_getPageConfig']);
      _enqueueAuditLog("system", "saveGroupMembers", { id: data.id, count: data.members.length });
      return { msg: "名單儲存成功" };
    }
  }
  throw new Error("找不到該分頁 ID：" + data.id);
}

function ministry_saveSermonSettings(data) {
  data.id = decryptGroupCode(data.id);
  const s = getMinistrySS().getSheetByName('Config');
  const configData = s.getDataRange().getValues();
  for (let i = 1; i < configData.length; i++) {
    if (configData[i][1].toString().trim() === data.id) {
      s.getRange(i + 1, 8).setValue(JSON.stringify(data.sermonSettings));
      _invalidateConfigDataCache();
      invalidateMinistryEventsCache();
      firebaseInvalidate(['ministry_getPageConfig']);
      _enqueueAuditLog("system", "saveSermonSettings", { id: data.id });
      return { msg: "講道設定儲存成功" };
    }
  }
  throw new Error("找不到該分頁 ID：" + data.id);
}

function ministry_forceRefreshEvents() {
  invalidateMinistryEventsCache();
  const freshEvents = _rebuildMinistryEventsCache();
  const count = freshEvents ? freshEvents.length : 0;
  _enqueueAuditLog("system", "forceRefreshEvents", { count: count });
  return { count: count };
}

function ministry_refreshCaches() {
  try {
    _invalidateMinistryGroupsCache();
    invalidateMinistryEventsCache();
    invalidateMinistryReportCache();
    _invalidateConfigDataCache();
    firebaseInvalidate([
      'ministry_getGroups',
      'ministry_getPageConfig',
      'ministry_getAggregatedReport',
      'ministry_getGroupMembers',
      'ministry_getTemplates'
    ]);
    return { success: true, message: 'Ministry caches refreshed' };
  } catch (e) {
    return { success: false, message: 'Ministry cache refresh failed: ' + e.message };
  }
}

// ═══════════════════════════════════════════════════════════
//  🤖 AI 排班（共用 GeminiHelper）
// ═══════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════
//  👥 小組成員身分管理（事工管理內部呼叫小組系統的橋接 API）
//
//  避免使用者在事工 / 小組系統間跳來跳去；
//  事工系統管理 小組聚會表模板 / 團契聚會表模板 時，
//  可直接編輯該小組成員的身分（核心同工/一般同工/小羊/陪伴同工），
//  異動會寫回主日的會友名單（單一真實來源），所有系統同步生效。
// ═══════════════════════════════════════════════════════════

/**
 * 取得指定小組的成員（含姓名/UID/身分/暱稱），給事工系統開 modal 編輯用
 * 內部直接呼叫小組系統的 checkGroupStatus
 */
function ministry_getGroupMembersForRoleEdit(groupName) {
  const result = checkGroupStatus(groupName);
  return result; // { isInitialized, members: [{name, uid, role, nickname}] }
}

/**
 * 把編輯後的小組成員身分寫回 — 內部呼叫 updateMemberList
 * updateMemberList 會走主日 master CRUD，並順便維護排序+暱稱
 */
function ministry_updateGroupMemberRoles(groupName, members) {
  return updateMemberList(groupName, members);
}

function ministry_parseWithAI(data) {
  if (!data.headers || !Array.isArray(data.headers) || data.headers.length === 0) {
    throw new Error("後端未接收到有效的表頭清單");
  }
  const systemPrompt = _ministryGetSystemPrompt(data.headers, data.members, data.groupPrompt, data.template);
  const result = callGemini(systemPrompt, data.text, { useCache: false });
  _enqueueAuditLog("system", "parseWithAI_success", { headers: data.headers, rowCount: Array.isArray(result) ? result.length : 0 });
  return result;
}

function _ministryGetSystemPrompt(headers, members, groupPrompt, template) {
  let templateRules = "";
  if (template === "團契聚會表模板") {
    templateRules = "\n\n【團契聚會表模板專屬規則】\n" +
      "- 「司會」欄位只能由核心同工擔任，一般同工不可填入此欄位。\n" +
      "- 「講員」、「主題」、「經文」、「地點」欄位請填入空字串 \"\"，由人工手動填寫。";
  }
  return `你是一個專業的教會行政排班大腦。你的唯一輸出格式是「純 JSON 陣列」。

【系統已知資源】
- 表格需填入的標題欄位：${JSON.stringify(headers)}
- 本小組可用名單（若有）：${JSON.stringify(members || [])}
- 本小組專屬排班規則：${groupPrompt || "無特殊規則"}

【系統強制規則（最高優先，不可違反）】
- 「陪伴同工」不列入自動排班，請勿將其填入任何欄位。
- 只能從【本小組可用名單】中挑選人員，名單外的人員一律不可填入。${templateRules}

【任務執行邏輯】
🔍 條件判斷：
如果輸入包含大量縮排、純文字列表或明顯的表格特徵 → 模式一：資料萃取
如果輸入包含對話指令、時間週期要求 → 模式二：智慧排班

🔴 模式一：資料萃取 — 以標題模糊匹配，將凌亂資料填入對應欄位。找不到的欄位填空字串 ""。
🔵 模式二：智慧排班 — 自動推算日期序列，從可用名單中挑人，遵守專屬規則。

【嚴格輸出格式】
JSON Key 必須 100% 完全等於標題欄位名稱。日期統一為 YYYY-MM-DD。直接輸出 JSON 陣列。`;
}

// ═══════════════════════════════════════════════════════════
//  🔄 自動同步外部小組清單（改成定時 trigger 跑，不再每次儀表板觸發）
// ═══════════════════════════════════════════════════════════

function ministry_autoSyncSmallGroups() {
  ensureConfigSchemaV3();

  const ss = getMinistrySS();
  const configSheet = ss.getSheetByName('Config');
  if (!configSheet) return;

  const configData = configSheet.getDataRange().getValues();

  const localByUuid = {};
  for (let i = 1; i < configData.length; i++) {
    const uuid = configData[i][0] ? configData[i][0].toString().trim() : "";
    if (!uuid) continue;
    localByUuid[uuid] = {
      rowIndex: i + 1,
      id:       configData[i][1] ? configData[i][1].toString().trim() : "",
      name:     configData[i][2] ? configData[i][2].toString().trim() : "",
      template: configData[i][3] ? configData[i][3].toString().trim() : "",
      status:   configData[i][4] ? configData[i][4].toString().trim() : ""
    };
  }

  try {
    // 直接讀本 GAS 的小組試算表（getGroupSS 由 GroupCore.js 提供）
    const listSheet = getGroupSS().getSheetByName("小組清單");
    if (!listSheet) return;
    const extData = listSheet.getDataRange().getValues();
    const templateSheet = ss.getSheetByName("小組聚會表模板");
    let hasConfigChanged = false;

    // 收集需要批次補的 UUID（寫入小組試算表 E 欄）
    const uuidPatches = []; // [{ row, uuid }]

    for (let r = 1; r < extData.length; r++) {
      const extName    = extData[r][0] ? extData[r][0].toString().trim() : "";
      const extRawStatus = extData[r][1] ? extData[r][1].toString().trim() : "";
      const extGroupId = extData[r][2] ? extData[r][2].toString().trim() : "";
      let extUuid      = extData[r][4] ? extData[r][4].toString().trim() : "";
      const extStatus  = (extRawStatus === "顯示") ? "啟用" : "停用";

      if (!extName) continue;

      if (!extUuid) {
        extUuid = Utilities.getUuid();
        uuidPatches.push({ row: r + 1, uuid: extUuid });
      }

      if (!localByUuid[extUuid]) {
        if (extStatus === "啟用" && templateSheet) {
          try {
            const newSheet = templateSheet.copyTo(ss);
            newSheet.setName(extName);
            configSheet.appendRow([extUuid, extGroupId, extName, "小組聚會表模板", "啟用", "", "[]"]);
            _enqueueAuditLog("system", "autoSync_created", { uuid: extUuid, id: extGroupId, name: extName });
            configData.push([extUuid, extGroupId, extName, "小組聚會表模板", "啟用", "", "[]"]);
          } catch (copyErr) {
            console.log("autoSync 建立分頁失敗: " + extName + " / " + copyErr);
          }
        }
      } else {
        const local = localByUuid[extUuid];
        let changed = false;
        if (local.id !== extGroupId)   { configData[local.rowIndex - 1][1] = extGroupId; changed = true; }
        if (local.name !== extName) {
          const targetSheet = ss.getSheetByName(local.name);
          if (targetSheet) { try { targetSheet.setName(extName); } catch (e) {} }
          configData[local.rowIndex - 1][2] = extName;
          changed = true;
        }
        if (local.status !== extStatus) { configData[local.rowIndex - 1][4] = extStatus; changed = true; }
        if (changed) hasConfigChanged = true;
      }
    }

    // 批次寫回小組試算表 UUID（一次 setValues 取代逐格 setValue）
    if (uuidPatches.length > 0) {
      // 因為 row 不連續，仍需逐格寫，但收集後可以最佳化（這裡保持簡潔）
      uuidPatches.forEach(p => listSheet.getRange(p.row, 5).setValue(p.uuid));
    }

    if (hasConfigChanged) {
      configSheet.getRange(1, 1, configData.length, configData[0].length).setValues(configData);
      invalidateMinistryReportCache();
      _invalidateConfigDataCache();
      _invalidateMinistryGroupsCache();
      firebaseInvalidate(['ministry_getGroups', 'ministry_getAggregatedReport', 'ministry_getPageConfig']);
    }
  } catch (e) {
    console.log("ministry_autoSyncSmallGroups 失敗: " + e.toString());
  }
}

/**
 * 設定 ministry_autoSyncSmallGroups 的定時 trigger（每小時跑一次）
 * 在 GAS 編輯器手動執行一次
 */
function setupMinistryAutoSyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'ministry_autoSyncSmallGroups') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('ministry_autoSyncSmallGroups')
    .timeBased()
    .everyHours(1)
    .create();
  Logger.log('✅ ministry_autoSyncSmallGroups trigger 已建立（每 1 小時）');
}

// ═══════════════════════════════════════════════════════════
//  📋 批次審計日誌（取代每次寫 Sheet）
// ═══════════════════════════════════════════════════════════

function _enqueueAuditLog(userId, action, details) {
  try {
    const props = PropertiesService.getScriptProperties();
    const queueRaw = props.getProperty('AUDIT_LOG_QUEUE') || '[]';
    const queue = JSON.parse(queueRaw);
    queue.push({ ts: new Date().toISOString(), userId: userId || 'system', action: action, details: details });
    // 上限保護：超過 500 筆強制 flush
    if (queue.length >= 500) {
      flushAuditLog();
      return;
    }
    props.setProperty('AUDIT_LOG_QUEUE', JSON.stringify(queue));
  } catch (e) { /* logging 失敗不影響業務 */ }
}

/**
 * 把暫存的審計日誌寫到 Sheet（每分鐘 trigger 觸發）
 */
function flushAuditLog() {
  const props = PropertiesService.getScriptProperties();
  const queueRaw = props.getProperty('AUDIT_LOG_QUEUE') || '[]';
  const queue = JSON.parse(queueRaw);
  if (queue.length === 0) return;

  // 立刻清空暫存（避免 race）
  props.deleteProperty('AUDIT_LOG_QUEUE');

  try {
    const ss = getMinistrySS();
    let logSheet = ss.getSheetByName('審計日誌');
    if (!logSheet) {
      logSheet = ss.insertSheet('審計日誌');
      logSheet.appendRow(["時間", "操作者", "動作", "詳細內容"]);
      logSheet.setFrozenRows(1);
    }
    const rows = queue.map(q => [new Date(q.ts), q.userId, q.action, JSON.stringify(q.details)]);
    logSheet.getRange(logSheet.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
  } catch (e) {
    // 寫入失敗：把資料塞回去等下次再試
    props.setProperty('AUDIT_LOG_QUEUE', queueRaw);
    console.log('flushAuditLog failed: ' + e.message);
  }
}

function setupAuditLogFlushTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'flushAuditLog') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('flushAuditLog').timeBased().everyMinutes(5).create();
  Logger.log('✅ flushAuditLog trigger 已建立（每 5 分鐘）');
}

// ═══════════════════════════════════════════════════════════
//  🗑️ Cache 清除輔助
// ═══════════════════════════════════════════════════════════

function invalidateMinistryReportCache() {
  const cache = CacheService.getScriptCache();
  cache.remove(MIN_REPORT_CACHE_PREFIX + 'smallGroup');
  cache.remove(MIN_REPORT_CACHE_PREFIX + 'others');
}

function _invalidateMinistryGroupsCache() {
  CacheService.getScriptCache().remove(MIN_GROUPS_CACHE_KEY);
}

function invalidateMinistryEventsCache() {
  CacheService.getScriptCache().remove(MIN_EVENTS_CACHE_KEY);
}


function ministry_verifyPageId(id, code) {
  var decryptedId = decryptGroupCode(id);
  var cleanCode = code ? code.toString().trim().toUpperCase() : "";
  var isMaster = (cleanCode === ADMIN_CODE);
  var isMatch = (cleanCode === decryptedId.toUpperCase());
  return { success: isMaster || isMatch };
}
