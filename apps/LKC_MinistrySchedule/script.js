// ============================================================
//  📋 教會服事管理系統 — 前端邏輯 (script.js)
//  修正版本 v3.0
//  v2.0 修正項目：
//    1. 使用 config.js 的 churchAPI() 與改進的錯誤處理
//    2. 加入防重複提交鎖定
//    3. 改用 userNotification 替代原生 alert
//    4. Session 管理改進（加入時間戳記）
//    5. 錯誤分類處理（不同錯誤顯示不同訊息）
//    6. AI 狀態提示改進
//  v3.0 修正項目：
//    7. 修正 getPageConfig 參數傳遞方式
//       { id: currentId } → { data: { id: currentId } }
//    8. 修正 getAggregatedReport 參數傳遞方式
//       { type: type } → { data: { type: type } }
//       （對齊後端 doPost 讀取 data.id / data.type 的方式）
// ============================================================

// userNotification / uiState / sessionManager / APIError 由中央 config.js 提供。
// 提供 getNotifier / getUIState shim 以避免大規模呼叫端改寫。
const getNotifier   = () => window.userNotification;
const getUIState    = () => window.uiState;
const getSessionMgr = () => window.sessionManager;

// ============================================================

const encryptGroupCode = window.encryptGroupCode || ((s) => s);
const decryptGroupCode = window.decryptGroupCode || ((s) => s);
const ENC_PREFIX = "enc_";

// 取得原始 ID 並立即進行混淆/加密處理，以防在網址列暴露明文 ID
let rawUrlId = new URLSearchParams(window.location.search).get('id') || "";
if (rawUrlId && rawUrlId.indexOf(ENC_PREFIX) !== 0) {
  const encryptedId = encryptGroupCode(rawUrlId);
  const urlParams = new URLSearchParams(window.location.search);
  urlParams.set('id', encryptedId);
  const newUrl = window.location.pathname + '?' + urlParams.toString();
  window.history.replaceState({}, '', newUrl);
  rawUrlId = encryptedId;
}
const currentId = rawUrlId;
let activeGroupName = "";
let currentTableHeaders = [];

let currentGroupMembers = [];
let currentCoreMembers = [];
let currentGroupPrompt = "";
let currentAutoRoleRules = "";

// 鎖定狀態與 Modal 實體
let isEditorUnlocked = false;
let bulletinModalInstance = null;

// 名單與聚會資料變數
let localCustomMembers = [];
let currentTemplate = "";
let currentEventData = [];
let currentSermonSettings = { useSermon: false, sermonType: "華語/聯合" };
let currentPageFieldConfig = null;
let fieldSettingsDraft = null;
let availableMinistryTemplates = [];

// 預覽佈告欄 modal 目前的篩選後矩陣（給下載 Excel 用）
let _currentBulletinFiltered = null;

const initialFieldTemplates = {
  "聚會型模板": {
    defaultFields: ["日期", "主題", "經文", "地點", "敬拜", "話語分享"],
    requiredFields: ["日期"]
  },
  "事工型模板": {
    defaultFields: ["日期", "地點"],
    requiredFields: ["日期"]
  }
};

const fieldTemplateBackendMap = {
  "聚會型模板": "小組聚會表模板",
  "事工型模板": "事工型模板"
};

function getBackendTemplateForFieldType(fieldTemplateType) {
  const preferred = fieldTemplateBackendMap[fieldTemplateType] || fieldTemplateType;
  if (!availableMinistryTemplates.length || availableMinistryTemplates.includes(preferred)) return preferred;
  if (fieldTemplateType === "聚會型模板") {
    return availableMinistryTemplates.find(t => t.includes("小組") || t.includes("團契")) || availableMinistryTemplates[0];
  }
  return availableMinistryTemplates.find(t => !t.includes("小組") && !t.includes("團契")) || availableMinistryTemplates[0];
}

function getFieldTemplateType(templateName) {
  if (templateName === "小組聚會表模板" || templateName === "團契聚會表模板" || templateName === "聚會型模板") {
    return "聚會型模板";
  }
  return "事工型模板";
}

function getFieldConfigStorageKey(pageId = currentId) {
  return `ministry.pageFieldConfig.${pageId || "new"}`;
}

function getEnabledFieldsFromConfig(config) {
  if (!config || !Array.isArray(config.fields)) return [];
  return config.fields.filter(field => field && field.enabled !== false).map(field => field.name);
}

function getRequiredFields(config) {
  return (config && Array.isArray(config.requiredFields) && config.requiredFields.length)
    ? config.requiredFields
    : ["日期"];
}

function normalizeFieldConfig(rawConfig, templateType, pageId) {
  const template = initialFieldTemplates[templateType] || initialFieldTemplates["事工型模板"];
  const requiredFields = Array.from(new Set([...(rawConfig && rawConfig.requiredFields || []), ...template.requiredFields]));
  const sourceFields = Array.isArray(rawConfig && rawConfig.fields) && rawConfig.fields.length
    ? rawConfig.fields
    : template.defaultFields.map(name => ({ name, enabled: true }));

  const seen = new Set();
  const fields = [];
  sourceFields.forEach(field => {
    const name = typeof field === "string" ? field : field && field.name;
    if (!name || seen.has(name)) return;
    seen.add(name);
    fields.push({
      name,
      enabled: requiredFields.includes(name) ? true : (typeof field === "object" && field.enabled === false ? false : true),
      custom: typeof field === "object" ? field.custom === true : !template.defaultFields.includes(name)
    });
  });
  requiredFields.forEach(name => {
    if (!seen.has(name)) {
      seen.add(name);
      fields.unshift({ name, enabled: true, custom: false });
    }
  });

  return {
    pageId: pageId || "",
    fieldTemplateType: templateType,
    fields,
    requiredFields,
    customFields: fields.filter(field => field.custom).map(field => field.name),
    updatedAt: new Date().toISOString()
  };
}

function buildPageFieldConfig(data, rawHeaders) {
  const templateType = getFieldTemplateType(data.template || "");
  const storageKey = getFieldConfigStorageKey();
  const stored = localStorage.getItem(storageKey);
  if (stored) {
    try {
      return normalizeFieldConfig(JSON.parse(stored), templateType, currentId);
    } catch (e) {
      localStorage.removeItem(storageKey);
    }
  }

  if (data.pageFieldConfig) {
    return normalizeFieldConfig(data.pageFieldConfig, data.pageFieldConfig.fieldTemplateType || templateType, currentId);
  }

  const existingHeaders = (rawHeaders || []).map(h => String(h || "").trim()).filter(Boolean);
  const template = initialFieldTemplates[templateType] || initialFieldTemplates["事工型模板"];
  const fields = existingHeaders.length ? existingHeaders : template.defaultFields;
  return normalizeFieldConfig({
    fields: fields.map(name => ({ name, enabled: true, custom: !template.defaultFields.includes(name) })),
    requiredFields: template.requiredFields
  }, templateType, currentId);
}

function savePageFieldConfigLocally(config) {
  currentPageFieldConfig = normalizeFieldConfig(config, config.fieldTemplateType || getFieldTemplateType(currentTemplate), currentId);
  localStorage.setItem(getFieldConfigStorageKey(), JSON.stringify(currentPageFieldConfig));
}


// ============================================================
//  🛡️ API 呼叫核心
//  config.js 已載入並提供 window.GAS_URL / window.AUTH_TOKEN，
//  此處的 fetchAPI 與中央 churchAPI 並存，差異在於：
//    - churchAPI 用 text/plain；fetchAPI 用 form-urlencoded（避開 preflight）
//    - fetchAPI 帶有 timeout/retry 邏輯
//  payload 結構：{ action, token, data: { ...params } }
// ============================================================

const _API_TIMEOUT_MS = 120000;

function normalizeMinistryAction(action) {
  return action.indexOf('ministry_') === 0 ? action : 'ministry_' + action;
}

function isMalformedCachedResult(result) {
  return result &&
    result.status &&
    result.status !== 'success' &&
    !result.message &&
    !Object.prototype.hasOwnProperty.call(result, 'data');
}

/**
 * 事工管理 API 呼叫
 *
 * 優先走中央 churchAPI（自動加 ministry_ 前綴 + Firebase RTDB 快取）
 * 若 churchAPI 不可用，回退至直接呼叫 GAS（手動加前綴）
 *
 * 回傳格式：後端統一回 { status, data, message? }
 * 此函式只回傳 data 部分；非 success 就拋 APIError
 */
async function fetchAPI(action, data = {}) {
  // ── 優先路徑：中央 churchAPI（含 Firebase cache + 自動加前綴） ──
  if (typeof window.churchAPI === 'function') {
    try {
      const result = await window.churchAPI(action, data);
      if (result && result.status === 'success') {
        return result.data;
      }
      if (isMalformedCachedResult(result)) {
        console.warn('[ministry-api] malformed cached result, fallback to direct GAS:', action, result);
        if (typeof window.churchAPIInvalidate === 'function') {
          window.churchAPIInvalidate(normalizeMinistryAction(action)).catch(err => {
            console.warn('[ministry-api] cache invalidation failed:', err);
          });
        }
        return await fetchDirectGAS(action, data);
      }
      if (!result || result.status !== 'success') {
        throw new APIError(
          (result && result.message) || '伺服器錯誤',
          null,
          classifyError(result && result.message)
        );
      }
    } catch (err) {
      if (err instanceof APIError) throw err;
      throw new APIError(err.message || '網路錯誤', null, classifyError(err.message));
    }
  }

  // ── 回退路徑：直接打 GAS（保留 retry / timeout 邏輯） ──
  // 手動加 ministry_ 前綴
  return await fetchDirectGAS(action, data);
}

async function fetchDirectGAS(action, data = {}) {
  const realAction = normalizeMinistryAction(action);
  const payload = {
    action: realAction,
    token:  window.AUTH_TOKEN,
    data:   data
  };

  const gasUrl = window.GAS_URL;
  if (!gasUrl) {
    throw new APIError("GAS 部署網址尚未設定", null, 'CONFIG_ERROR');
  }

  let lastError;
  const maxRetries = 3;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), _API_TIMEOUT_MS);

      // 改用 text/plain + 純 JSON body（與整合後的 doPost 一致）
      const response = await fetch(gasUrl, {
        method:   'POST',
        headers:  { 'Content-Type': 'text/plain;charset=utf-8' },
        body:     JSON.stringify(payload),
        signal:   controller.signal,
        redirect: 'follow'
      });
      clearTimeout(timer);

      if (!response.ok) throw new APIError(`HTTP ${response.status}`, response.status, 'HTTP_ERROR');

      const json = await response.json();
      if (json.status !== 'success') {
        throw new APIError(json.message || '伺服器錯誤', null, classifyError(json.message));
      }
      return json.data;

    } catch (err) {
      lastError = err;
      const type = err instanceof APIError ? err.type : classifyError(err.message);
      const retryable = ['HTTP_ERROR', 'SERVER_ERROR', 'SERVER_BUSY'].includes(type);
      if (retryable && attempt < maxRetries) {
        await new Promise(r => setTimeout(r, 1000 * attempt));
        continue;
      }
      break;
    }
  }
  throw lastError;
}

function classifyError(msg) {
  if (!msg) return 'UNKNOWN_ERROR';
  const m = msg.toLowerCase();
  if (m.includes('未授權') || m.includes('無效的憑證')) return 'AUTH_ERROR';
  if (m.includes('逾時') || m.includes('abort'))        return 'TIMEOUT';
  if (m.includes('failed to fetch') || m.includes('networkerror')) return 'NETWORK_ERROR';
  if (m.includes('已達上限'))                           return 'AI_DAILY_LIMIT';
  if (m.includes('503') || m.includes('忙線'))          return 'SERVER_BUSY';
  if (m.includes('找不到'))                             return 'NOT_FOUND';
  return 'UNKNOWN_ERROR';
}


// ============================================================
//  🎨 改進的錯誤處理
// ============================================================
function handleAPIError(err) {
  console.error("API 錯誤", err);

  if (!(err instanceof APIError)) {
    // 非 API 錯誤（例如 JSON 解析錯誤）
    getNotifier().error("發生未預期的錯誤：" + err.message);
    return;
  }

  const errorType = err.type;
  const message = err.message;

  switch (errorType) {
    case 'AUTH_ERROR':
      getNotifier().error("❌ 權限不足，請重新整理並輸入正確的 ID");
      break;

    case 'PERMISSION_ERROR':
      getNotifier().error("❌ 您沒有權限執行此操作");
      break;

    case 'TIMEOUT':
      getNotifier().error("⏱️ 請求逾時，請檢查網路連線並重試");
      break;

    case 'NETWORK_ERROR':
      getNotifier().error("🌐 網路連線失敗，請檢查網路狀態");
      break;

    case 'RATE_LIMIT':
      getNotifier().warning("⚠️ 請求過於頻繁，請稍候再試");
      break;

    case 'AI_DAILY_LIMIT':
      getNotifier().error("🤖 今日 AI 使用次數已達上限，請明日再試");
      break;

    case 'SERVER_BUSY':
      getNotifier().warning("⚠️ 伺服器忙線中，將自動重試...");
      break;

    case 'NOT_FOUND':
      getNotifier().error("❌ 找不到相關資料：" + message);
      break;

    case 'VALIDATION_ERROR':
      getNotifier().error("❌ 資料驗證失敗：" + message);
      break;

    default:
      getNotifier().error("❌ 錯誤：" + message);
  }
}


// ============================================================
//  📥 頁面初始化
// ============================================================
window.onload = async () => {
  // 檢查編輯鎖定狀態（使用改進的 sessionManager）
  if (currentId && getSessionMgr().isUnlocked(currentId)) {
    isEditorUnlocked = true;
  }

  // 還原功能區收合狀態
  const isCollapsed = localStorage.getItem('ministry.primaryActionsCollapsed') === 'true';
  const container = document.querySelector('.ministry-primary-actions');
  const btn = document.getElementById('toggleActionsBtn');
  const arrow = document.getElementById('toggleActionsArrow');
  if (container && btn) {
    if (isCollapsed) {
      container.classList.add('collapsed');
      if (arrow) arrow.innerText = '▾';
      btn.classList.remove('active');
    } else {
      container.classList.remove('collapsed');
      if (arrow) arrow.innerText = '▴';
      btn.classList.add('active');
    }
  }

  if (!currentId) {
    showSection('adminMain');
    await loadAdminData();
  } else {
    showSection('reportSection');
    initDateQuickFilter();
    try {
      const data = await fetchAPI('getPageConfig', { id: currentId });
      renderTable(data);

      // 如果未解鎖，才顯示佈告欄 (預覽模式)
      if (!isEditorUnlocked) {
        showBulletinBoard();
      }
    } catch (err) {
      handleAPIError(err);
    }
  }
};


// ============================================================
//  📊 載入管理頁面資料
// ============================================================
async function loadAdminData() {
  try {
    getNotifier().showLoading("⏳ 整理儀表板中...");

    const [groups, templates] = await Promise.all([
      fetchAPI('getGroups', {}),
      fetchAPI('getTemplates', {})
    ]);
    availableMinistryTemplates = Array.isArray(templates) ? templates : [];

    const div = document.getElementById('groupButtons');
    const base = window.location.href.split('?')[0];

    const grouped = groups.reduce((acc, g) => {
      const cat = g.template || "未分類項目";
      if (!acc[cat]) acc[cat] = [];
      acc[cat].push(g);
      return acc;
    }, {});

    let html = "";
    for (const cat in grouped) {
      html += `<div class="col-12 category-section" data-cat="${cat}"><div class="category-header">📦 ${cat} <span class="category-badge">${grouped[cat].length}</span></div></div>`;

      html += grouped[cat].map(g => {
        const isEnabled = g.status !== "停用";
        const shareUrl = `${base}?id=${encryptGroupCode(g.id)}`;

        return `
          <div class="col-12 col-md-4 group-item" data-search="${g.name} ${g.template}">
            <div class="card group-card h-100 shadow-sm d-flex flex-column" style="opacity: ${isEnabled ? '1' : '0.5'}; border-left: 5px solid ${isEnabled ? '#0d6efd' : '#ced4da'};">
              <div class="card-body p-3 flex-grow-1">
                <div class="d-flex align-items-center justify-content-between mb-2">
                  <a href="${isEnabled ? shareUrl : 'javascript:void(0)'}" class="group-link m-0" style="${isEnabled ? '' : 'pointer-events: none; cursor: default;'}">
                    <h5 class="card-title ${isEnabled ? 'text-dark' : 'text-muted'} m-0" style="${isEnabled ? '' : 'text-decoration: line-through;'}">${g.name}</h5>
                  </a>
                  <div class="form-check form-switch m-0 ms-3">
                    <input class="form-check-input" type="checkbox" role="switch" ${isEnabled ? 'checked' : ''} onchange="toggleStatus('${g.id}', '${g.status}')">
                  </div>
                </div>
                ${isEnabled ? `<button class="btn btn-sm btn-outline-success w-100 mt-2 fw-bold" onclick="copyShareLink('${shareUrl}')">🔗 複製分享網址</button>` : '<p class="text-muted small mt-2 m-0 text-center">已停用</p>'}
              </div>
            </div>
          </div>
        `;
      }).join('');
    }

    div.innerHTML = html || '<p class="text-center text-muted">目前尚無資料</p>';
    document.getElementById('templateSelect').innerHTML = `
      <option value="" disabled selected>選擇表格類型</option>
      <option value="聚會型模板">聚會型模板</option>
      <option value="事工型模板">事工型模板</option>
    `;

    getNotifier().success("✅ 儀表板已載入");
  } catch (err) {
    handleAPIError(err);
    document.getElementById('groupButtons').innerHTML = '<p class="text-danger">載入失敗，請重試</p>';
  } finally {
    getNotifier().hideLoading();
  }
}


// ============================================================
//  🔗 複製分享網址
// ============================================================
function copyShareLink(url) {
  navigator.clipboard.writeText(url)
    .then(() => getNotifier().success("✅ 專屬網址已複製！"))
    .catch(() => {
      getNotifier().warning("⚠️ 複製失敗，請手動複製：\n" + url);
    });
}


// ============================================================
//  🔍 搜尋小組
// ============================================================
function filterGroups() {
  const val = document.getElementById('groupSearch').value.toLowerCase();
  document.querySelectorAll('.group-item').forEach(el => {
    el.style.display = el.dataset.search.toLowerCase().includes(val) ? "" : "none";
  });
  document.querySelectorAll('.category-section').forEach(header => {
    let hasVisible = false;
    let next = header.nextElementSibling;
    while (next && next.classList.contains('group-item')) {
      if (next.style.display !== "none") hasVisible = true;
      next = next.nextElementSibling;
    }
    header.style.display = hasVisible ? "" : "none";
  });
}

function remapRowToCurrentHeaders(row, sourceHeaders) {
  const mapped = currentTableHeaders.map(header => {
    const idx = sourceHeaders.indexOf(header);
    return idx !== -1 ? (row[idx] || "") : "";
  });
  return mapped;
}


// ============================================================
//  📄 渲染排班表單
// ============================================================
function renderTable(data) {
  activeGroupName = data.groupName;
  document.getElementById('groupTitle').innerText = data.groupName;

  currentGroupMembers = data.members || [];
  currentCoreMembers = data.coreMembers || [];
  currentGroupPrompt = data.groupPrompt || "";
  currentAutoRoleRules = data.autoRoleRules || "";
  currentEventData = (data.eventData || []).map(ev => {
    if (ev.date) {
      const normalized = parseGregorianDate(String(ev.date));
      if (normalized) {
        ev.date = normalized;
      }
    }
    return ev;
  });

  currentTemplate = data.template || "";
  localCustomMembers = data.customMembers || [];
  updateTemplateSpecificLabels();

  const memberBtn = document.getElementById('manageMembersBtn');
  const groupRoleBtn = document.getElementById('manageGroupRolesBtn');
  const isGroupOrFellowship = (currentTemplate === "小組聚會表模板" || currentTemplate === "團契聚會表模板");

  if (memberBtn && !isGroupOrFellowship) {
    memberBtn.classList.remove('hidden');
    currentGroupMembers = localCustomMembers.map(m => m.name);

    if (currentTemplate === "新家人服事表模板") {
      const parentNames = localCustomMembers.filter(m => m.role === "小家長").map(m => m.name).join(", ");
      const normalNames = localCustomMembers.filter(m => m.role === "一般同工").map(m => m.name).join(", ");
      currentAutoRoleRules = `【系統強制權限】：\n小家長 (${parentNames})：可排所有服事。\n一般同工 (${normalNames})：不可排特定帶領服事。`;
    }
  } else if (memberBtn) {
    memberBtn.classList.add('hidden');
  }

  // 小組/團契模板才顯示「設定組員身分」按鈕與「講道連動設定」按鈕
  if (groupRoleBtn) {
    groupRoleBtn.classList.toggle('hidden', !isGroupOrFellowship);
  }
  const sermonSettingsSection = document.getElementById('sermonSettingsSection');
  if (sermonSettingsSection) {
    sermonSettingsSection.classList.toggle('hidden', !isGroupOrFellowship);
  }

  if (isGroupOrFellowship) {
    currentSermonSettings = data.sermonSettings || { useSermon: false, sermonType: "華語/聯合" };
    const useSermonToggle = document.getElementById('useSermonToggle');
    if (useSermonToggle) {
      useSermonToggle.checked = currentSermonSettings.useSermon === true;
      toggleSermonTypeSelect();
    }
    const sermonTypeSelect = document.getElementById('sermonTypeSelect');
    if (sermonTypeSelect) {
      sermonTypeSelect.value = currentSermonSettings.sermonType || "華語/聯合";
    }
  }

  const promptInput = document.getElementById('groupPromptInput');
  if (promptInput) promptInput.value = currentGroupPrompt;

  if (!data.matrix || !Array.isArray(data.matrix) || data.matrix.length === 0) {
    const templateType = getFieldTemplateType(currentTemplate);
    data.matrix = [initialFieldTemplates[templateType].defaultFields.slice()];
  }

  let rawHeaders = data.matrix[0].map(h => h.toString().trim());
  let validColCount = rawHeaders.length;
  while (validColCount > 0 && rawHeaders[validColCount - 1] === "") validColCount--;
  currentPageFieldConfig = buildPageFieldConfig(data, rawHeaders.slice(0, validColCount));
  rawHeaders = getEnabledFieldsFromConfig(currentPageFieldConfig);
  validColCount = rawHeaders.length;

  // 自動補齊「套用講道」欄位
  if (isGroupOrFellowship) {
    const hasSermonLinkHeader = rawHeaders.slice(0, validColCount).some(h => h === "套用講道");
    if (!hasSermonLinkHeader) {
      rawHeaders[validColCount] = "套用講道";
      validColCount++;
    }
  }

  currentTableHeaders = rawHeaders.slice(0, validColCount);

  let datalistHTML = "";
  if (currentGroupMembers.length > 0)
    datalistHTML += `<datalist id="allMembersList">` + currentGroupMembers.map(m => `<option value="${m}">`).join('') + `</datalist>`;
  if (currentCoreMembers.length > 0)
    datalistHTML += `<datalist id="coreMembersList">` + currentCoreMembers.map(m => `<option value="${m}">`).join('') + `</datalist>`;

  if (currentTemplate !== "小組聚會表模板") {
    if (currentTemplate === "新家人服事表模板") {
      const normalNames = localCustomMembers.filter(m => m.role === "一般同工").map(m => m.name);
      const parentNames = localCustomMembers.filter(m => m.role === "小家長").map(m => m.name);
      datalistHTML += `<datalist id="normalMembersList">` + normalNames.map(m => `<option value="${m}">`).join('') + `</datalist>`;
      datalistHTML += `<datalist id="parentMembersList">` + parentNames.map(m => `<option value="${m}">`).join('') + `</datalist>`;
    } else {
      const customNames = localCustomMembers.map(m => m.name);
      datalistHTML += `<datalist id="customMembersList">` + customNames.map(m => `<option value="${m}">`).join('') + `</datalist>`;
    }
  }

  const gridTemplate = buildRecordGridTemplate(validColCount);

  let html = datalistHTML;
  html += `<div class="record-grid-header fw-bold text-muted mb-2" style="display: grid; grid-template-columns: ${gridTemplate};">`;
  currentTableHeaders.forEach(h => html += `<div class="record-cell record-header-cell">${h}</div>`);
  html += `<div class="record-cell record-header-cell record-delete-header"><button type="button" class="record-delete-header-btn" onclick="deleteSelectedRows()">刪除</button></div></div>`;
  html += `<div id="rowsContainer" class="d-flex flex-column gap-2">`;

  const sourceHeaders = data.matrix[0].map(h => h.toString().trim());
  const rows = data.matrix.slice(1);
  let validRows = rows
    .filter(r => r.some(cell => cell.toString().trim() !== ""))
    .map(row => remapRowToCurrentHeaders(row, sourceHeaders));

  // 補齊行矩陣長度與預設值
  const sermonLinkColIdx = currentTableHeaders.indexOf("套用講道");
  validRows.forEach(row => {
    while (row.length < validColCount) {
      row.push("");
    }
    if (isGroupOrFellowship && sermonLinkColIdx !== -1 && !row[sermonLinkColIdx]) {
      row[sermonLinkColIdx] = "N";
    }
  });

  const dateColIdx = currentTableHeaders.findIndex(h => h.includes("日期"));
  const nameColIdx = currentTableHeaders.findIndex(h => h.includes("聚會名稱"));
  const catColIdx = currentTableHeaders.findIndex(h => h.includes("聚會類別"));

  // 1. 若為非小組模板，先將 eventData 中缺少的日期補入
  if (dateColIdx !== -1 && currentTemplate !== "小組聚會表模板" && currentTemplate !== "團契聚會表模板" && currentEventData.length > 0) {
    const existingDates = validRows.map(r => r[dateColIdx]);
    currentEventData.forEach(event => {
      if (!existingDates.includes(event.date)) {
        let newRow = new Array(validColCount).fill("");
        newRow[dateColIdx] = event.date;
        if (nameColIdx !== -1) newRow[nameColIdx] = event.name;
        if (catColIdx !== -1) newRow[catColIdx] = event.category;
        if (sermonLinkColIdx !== -1) newRow[sermonLinkColIdx] = "N";
        validRows.push(newRow);
      }
    });
  }

  // 2. 對於所有有效的 rows 中的日期，一律先格式化為 yyyy/mm/dd
  if (dateColIdx !== -1) {
    validRows.forEach(row => {
      if (row[dateColIdx]) {
        const slashDate = parseToSlashDate(row[dateColIdx]);
        if (slashDate) {
          row[dateColIdx] = slashDate;
        }
      }
    });

    // 3. 不限模板，全域按日期由小到大排序 (空白排在最下方)
    validRows.sort((a, b) => {
      let dateA = a[dateColIdx] || "9999/99/99";
      let dateB = b[dateColIdx] || "9999/99/99";
      if (!a[dateColIdx] || a[dateColIdx].trim() === "") dateA = "9999/99/99";
      if (!b[dateColIdx] || b[dateColIdx].trim() === "") dateB = "9999/99/99";
      return dateA.localeCompare(dateB);
    });
  }

  if (validRows.length === 0) {
    const emptyRow = new Array(validColCount).fill("");
    if (sermonLinkColIdx !== -1) emptyRow[sermonLinkColIdx] = "N";
    validRows.push(emptyRow);
  }

  validRows.forEach((rowData) => html += createRowHTML(rowData, gridTemplate));

  html += `</div>`;
  html += `<button type="button" class="btn btn-outline-primary w-100 mt-3 border border-2 border-primary border-opacity-50" style="border-style: dashed !important;" onclick="addNewRow()">➕ 新增一筆空白列</button>`;

  document.getElementById('dynamicFormContainer').innerHTML = html;
  initGridInteraction();

  // 載入時，針對所有有連動講道的列，進行一次講道資料的自動套用初始化
  document.querySelectorAll('.record-row').forEach(rowDiv => {
    if (sermonLinkColIdx !== -1) {
      const checkbox = rowDiv.querySelector(`input.sermon-link-checkbox[data-c="${sermonLinkColIdx}"]`);
      const sw = rowDiv.querySelector(`input.sermon-lang-switch[data-c="${sermonLinkColIdx}"]`);
      if (checkbox && checkbox.checked && dateColIdx !== -1) {
        const dVal = rowDiv.querySelector(`input.grid-input[data-c="${dateColIdx}"]`).value.trim();
        const langVal = sw && sw.checked ? "台語/聯合" : "華語/聯合";
        updateRowSermonState(rowDiv, langVal, dVal);
      }
    }
  });
}


// ============================================================
//  🧩 建立表單列 HTML
// ============================================================
function createRowHTML(rowData, gridTemplate) {
  if (!gridTemplate) gridTemplate = buildRecordGridTemplate(currentTableHeaders.length);
  let rowHtml = `<div class="record-row align-items-center" style="display: grid; grid-template-columns: ${gridTemplate};">`;

  const sermonLinkColIdx = currentTableHeaders.indexOf("套用講道");
  const rowLinkVal = sermonLinkColIdx !== -1 ? String(rowData[sermonLinkColIdx] || "N").trim() : "N";
  const isRowSermonLinked = rowLinkVal !== "N" && rowLinkVal !== "";

  currentTableHeaders.forEach((header, cIdx) => {
    let val = rowData[cIdx] || "";
    if (header === "經文" && (currentTemplate === "小組聚會表模板" || currentTemplate === "團契聚會表模板") && window.BibleFormatter) {
      val = window.BibleFormatter.format(val);
    }
    if (header === "套用講道") {
      const currentLinkVal = String(rowData[cIdx] || "N").trim();
      const isChecked = currentLinkVal !== "N" && currentLinkVal !== "";
      
      let langVal = currentLinkVal;
      if (currentLinkVal === "Y" || currentLinkVal === "true") {
        langVal = currentSermonSettings.sermonType;
      }
      if (langVal !== "華語/聯合" && langVal !== "台語/聯合") {
        langVal = currentSermonSettings.sermonType || "華語/聯合";
      }
      
      const isTaiwanese = langVal === "台語/聯合";
      const leftOpacity = (isChecked && !isTaiwanese) ? "1" : "0.4";
      const rightOpacity = (isChecked && isTaiwanese) ? "1" : "0.4";
      const switchDisabledAttr = isChecked ? "" : "disabled";
      
      const cellHtml = `
        <div class="d-flex align-items-center justify-content-center gap-1" style="width: 100%; height: 100%; font-size: 0.8rem; user-select: none;">
          <input type="checkbox" class="grid-checkbox sermon-link-checkbox" data-c="${cIdx}" ${isChecked ? 'checked' : ''} onchange="onSermonCheckboxChange(this)" style="margin-right: 2px;">
          <span class="text-primary fw-bold" id="lang-label-left-${cIdx}" style="font-size: 0.72rem; opacity: ${leftOpacity}; transition: opacity 0.2s;">華</span>
          <div class="form-check form-switch m-0 p-0 d-flex align-items-center" style="min-height: auto;">
            <input class="form-check-input sermon-lang-switch" type="checkbox" role="switch" data-c="${cIdx}" ${isTaiwanese ? 'checked' : ''} ${switchDisabledAttr} onchange="onSermonSwitchChange(this)" style="cursor: pointer; margin: 0; width: 1.6em; height: 0.95em; transition: all 0.2s;">
          </div>
          <span class="fw-bold" id="lang-label-right-${cIdx}" style="color: #fd7e14 !important; font-size: 0.72rem; opacity: ${rightOpacity}; transition: opacity 0.2s;">台</span>
        </div>
      `;
      rowHtml += `<div class="record-cell d-flex align-items-center justify-content-center">${cellHtml}</div>`;
      return;
    }

    let listAttr = "";
    let extraClass = "";
    let inputType = "text";
    // 表格內部的日期改用 text 輸入框以支援 yyyy/mm/dd 格式的手動輸入與顯示
    if (header.includes("日期")) inputType = "text";

    const isSermonField = header === "主題" || header === "經文";
    const readonlyAttr = (isRowSermonLinked && isSermonField) ? "readonly" : "";

    if (currentTemplate === "團契聚會表模板") {
      const allDropdownCols = ["破冰", "敬拜"];
      const coreDropdownCols = ["司會"];
      const isAllCol = allDropdownCols.some(c => header.includes(c));
      const isCoreCol = coreDropdownCols.some(c => header.includes(c));

      if (isCoreCol) {
        listAttr = `list="coreMembersList"`;
        extraClass = `datalist-input`;
      } else if (isAllCol) {
        listAttr = `list="allMembersList"`;
        extraClass = `datalist-input`;
      }

    } else if (currentTemplate !== "小組聚會表模板") {
      if (header.includes("日期") || header.includes("聚會名稱") || header.includes("聚會類別")) {
        listAttr = "";
        extraClass = "";
      } else {
        if (currentTemplate === "新家人服事表模板" && header.includes("小家長")) {
          listAttr = `list="parentMembersList"`;
          extraClass = `datalist-input`;
        } else if (currentTemplate === "新家人服事表模板" && header.includes("新家人同工")) {
          listAttr = `list="normalMembersList"`;
          extraClass = `datalist-input`;
        } else {
          listAttr = `list="customMembersList"`;
          extraClass = `datalist-input`;
        }
      }

    } else {
      const allDropdownCols = ["破冰", "敬拜", "分享"];
      const coreDropdownCols = ["話語", "領會", "主領", "帶領"];
      const isAllCol = allDropdownCols.some(c => header.includes(c));
      const isCoreCol = coreDropdownCols.some(c => header.includes(c));

      if (isCoreCol) {
        listAttr = `list="coreMembersList"`;
        extraClass = `datalist-input`;
      } else if (isAllCol) {
        listAttr = `list="allMembersList"`;
        extraClass = `datalist-input`;
      }
    }

    rowHtml += `<div class="record-cell"><input type="${inputType}" class="grid-input ${extraClass}" data-c="${cIdx}" value="${val}" title="${val}" ${listAttr} ${readonlyAttr}></div>`;
  });

  rowHtml += `<div class="record-cell d-flex justify-content-center"><input type="checkbox" class="form-check-input row-delete-checkbox" title="勾選後可批次刪除"></div></div>`;
  return rowHtml;
}

function buildRecordGridTemplate(columnCount) {
  const isNarrow = window.matchMedia && window.matchMedia('(max-width: 768px)').matches;
  const inputMinWidth = isNarrow ? 172 : 148;
  const actionWidth = isNarrow ? 64 : 48;
  return `repeat(${columnCount}, ${inputMinWidth}px) ${actionWidth}px`;
}

function sortRowsByDate() {
  const dateColIdx = currentTableHeaders.findIndex(h => h.includes("日期"));
  if (dateColIdx === -1) return;

  // 收集目前畫面上所有有效的 rows
  const matrix = collectVisibleMatrix();
  const headers = matrix[0];
  const rows = matrix.slice(1);

  // 排序，確保空白日期排在最下方
  rows.sort((a, b) => {
    let dateA = a[dateColIdx] || "9999/99/99";
    let dateB = b[dateColIdx] || "9999/99/99";
    if (!a[dateColIdx] || a[dateColIdx].trim() === "") dateA = "9999/99/99";
    if (!b[dateColIdx] || b[dateColIdx].trim() === "") dateB = "9999/99/99";
    return dateA.localeCompare(dateB);
  });

  // 重新渲染表格
  rerenderWithMatrix([headers, ...rows]);
}
window.sortRowsByDate = sortRowsByDate;

// ============================================================
//  ➕ 新增列 / 🗑️ 刪除列
// ============================================================
function addNewRow() {
  const container = document.getElementById('rowsContainer');
  const tempDiv = document.createElement('div');
  
  const defaultRow = Array(currentTableHeaders.length).fill("");
  const sermonLinkColIdx = currentTableHeaders.indexOf("套用講道");
  if (sermonLinkColIdx !== -1) {
    defaultRow[sermonLinkColIdx] = "N";
  }
  
  tempDiv.innerHTML = createRowHTML(defaultRow);
  container.prepend(tempDiv.firstElementChild);
}

function deleteRow(btnElement) {
  if (confirm("確定要刪除這筆排班資料嗎？")) {
    btnElement.parentElement.remove();
  }
}

function deleteSelectedRows() {
  const selected = Array.from(document.querySelectorAll('.row-delete-checkbox:checked'));
  if (selected.length === 0) {
    getNotifier().warning("⚠️ 請先勾選要刪除的列");
    return;
  }
  if (!confirm(`確定要刪除 ${selected.length} 列資料嗎？`)) return;
  selected.forEach(checkbox => checkbox.closest('.record-row')?.remove());
  getNotifier().success(`✅ 已刪除 ${selected.length} 列`);
}

function collectVisibleMatrix() {
  const matrix = [currentTableHeaders.slice()];
  document.querySelectorAll('.record-row').forEach(rowDiv => {
    const row = [];
    currentTableHeaders.forEach((header, cIdx) => {
      const input = rowDiv.querySelector(`[data-c="${cIdx}"]`);
      if (!input) {
        row.push("");
      } else if (input.type === 'checkbox') {
        row.push(input.checked ? "Y" : "N");
      } else {
        row.push(input.value.trim());
      }
    });
    matrix.push(row);
  });
  return matrix;
}

function rerenderWithMatrix(matrix) {
  renderTable({
    groupName: activeGroupName,
    template: currentTemplate,
    matrix,
    members: currentGroupMembers,
    coreMembers: currentCoreMembers,
    customMembers: localCustomMembers,
    groupPrompt: currentGroupPrompt,
    autoRoleRules: currentAutoRoleRules,
    eventData: currentEventData,
    sermonSettings: currentSermonSettings
  });
}

function isMeetingTemplate() {
  return currentTemplate === "小組聚會表模板" || currentTemplate === "團契聚會表模板";
}

function updateTemplateSpecificLabels() {
  const serviceMode = !isMeetingTemplate();
  const quarterTitle = serviceMode ? "新增一季服事表" : "新增一季聚會";
  const settingsTitle = serviceMode ? "服事表設定" : "聚會表設定";

  const quarterActionTitle = document.getElementById('quarterActionTitle');
  const quarterModalTitle = document.getElementById('quarterModalTitle');
  const settingsActionTitle = document.getElementById('scheduleSettingsActionTitle');
  const settingsModalTitle = document.getElementById('scheduleSettingsModalTitle');

  if (quarterActionTitle) quarterActionTitle.innerText = quarterTitle;
  if (quarterModalTitle) quarterModalTitle.innerText = quarterTitle;
  if (settingsActionTitle) settingsActionTitle.innerText = settingsTitle;
  if (settingsModalTitle) settingsModalTitle.innerText = settingsTitle;
}

function openQuarterModal() {
  const yearInput = document.getElementById('quarterYear');
  const quarterSelect = document.getElementById('quarterNumber');
  const sermonCheckbox = document.getElementById('quarterUseSermon');
  const sermonRow = document.getElementById('quarterUseSermonRow');
  const now = new Date();
  if (yearInput && !yearInput.value) yearInput.value = now.getFullYear();
  if (quarterSelect && !quarterSelect.value) quarterSelect.value = String(Math.floor(now.getMonth() / 3) + 1);
  const meetingMode = isMeetingTemplate();
  if (sermonRow) sermonRow.classList.toggle('hidden', !meetingMode);
  if (sermonCheckbox) sermonCheckbox.checked = meetingMode && currentSermonSettings.useSermon === true;
  updateTemplateSpecificLabels();
  new bootstrap.Modal(document.getElementById('quarterModal')).show();
}

function generateQuarterRows() {
  const year = Number(document.getElementById('quarterYear').value);
  const quarter = Number(document.getElementById('quarterNumber').value);
  const weekday = Number(document.getElementById('quarterWeekday').value);
  const useAI = document.getElementById('quarterUseAI').checked;
  const useSermonForNewRows = isMeetingTemplate() && document.getElementById('quarterUseSermon').checked;
  const dateColIdx = currentTableHeaders.findIndex(h => h.includes("日期"));
  const sermonLinkColIdx = currentTableHeaders.indexOf("套用講道");
  if (!year || dateColIdx === -1) {
    getNotifier().warning("⚠️ 需要年度與日期欄位才能產生季度資料");
    return;
  }

  const startMonth = (quarter - 1) * 3;
  const start = new Date(year, startMonth, 1);
  const end = new Date(year, startMonth + 3, 0);
  const dates = [];
  for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
    if (d.getDay() === weekday) {
      const y = d.getFullYear();
      const m = String(d.getMonth() + 1).padStart(2, "0");
      const day = String(d.getDate()).padStart(2, "0");
      // 改成 yyyy/mm/dd 斜線格式
      dates.push(`${y}/${m}/${day}`);
    }
  }

  const existingDates = new Set();
  document.querySelectorAll('.record-row').forEach(rowDiv => {
    const input = rowDiv.querySelector(`.grid-input[data-c="${dateColIdx}"]`);
    if (input && input.value.trim()) existingDates.add(input.value.trim());
  });

  let added = 0;
  const addedDates = [];
  dates.forEach(date => {
    if (existingDates.has(date)) return;
    const row = Array(currentTableHeaders.length).fill("");
    row[dateColIdx] = date;
    if (sermonLinkColIdx !== -1) row[sermonLinkColIdx] = useSermonForNewRows ? currentSermonSettings.sermonType : "N";
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = createRowHTML(row);
    const rowEl = tempDiv.firstElementChild;
    document.getElementById('rowsContainer').appendChild(rowEl);
    if (sermonLinkColIdx !== -1 && useSermonForNewRows) {
      updateRowSermonState(rowEl, currentSermonSettings.sermonType, date);
    }
    addedDates.push(date);
    added++;
  });

  bootstrap.Modal.getInstance(document.getElementById('quarterModal'))?.hide();
  getNotifier().success(`✅ 已新增 ${added} 筆季度聚會日期`);
  sortRowsByDate();

  if (useAI && addedDates.length > 0) {
    const aiBox = document.getElementById('aiRawText');
    if (aiBox) {
      aiBox.value = [
        `請依照目前儲存的班表規則，為 ${year} 年第 ${quarter} 季以下日期進行排班。`,
        "請保留日期，不要新增或刪除日期。",
        "",
        addedDates.map(date => `- ${date}`).join("\n")
      ].join("\n");
      processAI();
    }
  }
}

function openAiScheduleModal() {
  const modal = new bootstrap.Modal(document.getElementById('aiScheduleModal'));
  modal.show();
  setTimeout(() => {
    const box = document.getElementById('aiRawText');
    if (box) box.focus();
  }, 180);
}

function openScheduleRuleModal() {
  const modal = new bootstrap.Modal(document.getElementById('scheduleRuleModal'));
  modal.show();
  setTimeout(() => {
    const input = document.getElementById('groupPromptInput');
    if (input) input.focus();
  }, 180);
}

function openScheduleSettingsModal() {
  new bootstrap.Modal(document.getElementById('scheduleSettingsModal')).show();
}

function focusPasteBox() {
  openAiScheduleModal();
}

function focusAiRawText() {
  const box = document.getElementById('aiRawText');
  if (!box) return;
  box.scrollIntoView({ behavior: 'smooth', block: 'center' });
  box.focus();
}

function openFieldSettingsModal() {
  fieldSettingsDraft = JSON.parse(JSON.stringify(currentPageFieldConfig || normalizeFieldConfig(null, getFieldTemplateType(currentTemplate), currentId)));
  renderFieldSettingsList();
  new bootstrap.Modal(document.getElementById('fieldSettingsModal')).show();
}

function renderFieldSettingsList() {
  const list = document.getElementById('fieldSettingsList');
  const required = getRequiredFields(fieldSettingsDraft);
  list.innerHTML = fieldSettingsDraft.fields.map((field, idx) => {
    const isRequired = required.includes(field.name);
    return `
      <div class="field-settings-row">
        <div class="field-settings-name">${field.name}${isRequired ? ' <span class="field-settings-required">必要</span>' : ''}</div>
        <div class="form-check form-switch m-0">
          <input class="form-check-input" type="checkbox" ${field.enabled !== false ? 'checked' : ''} ${isRequired ? 'disabled' : ''} onchange="toggleDraftField(${idx}, this.checked)">
        </div>
        <div class="text-muted small">${field.custom ? '自訂' : '模板'}</div>
        <div class="field-settings-actions">
          <button class="btn btn-sm btn-outline-secondary" type="button" onclick="moveDraftField(${idx}, -1)" ${idx === 0 ? 'disabled' : ''}>↑</button>
          <button class="btn btn-sm btn-outline-secondary" type="button" onclick="moveDraftField(${idx}, 1)" ${idx === fieldSettingsDraft.fields.length - 1 ? 'disabled' : ''}>↓</button>
          <button class="btn btn-sm btn-outline-danger" type="button" onclick="removeDraftField(${idx})" ${isRequired ? 'disabled' : ''}>刪除</button>
        </div>
      </div>
    `;
  }).join('');
}

function toggleDraftField(idx, checked) {
  const required = getRequiredFields(fieldSettingsDraft);
  if (required.includes(fieldSettingsDraft.fields[idx].name)) return;
  fieldSettingsDraft.fields[idx].enabled = checked;
}

function moveDraftField(idx, delta) {
  const next = idx + delta;
  if (next < 0 || next >= fieldSettingsDraft.fields.length) return;
  const copy = fieldSettingsDraft.fields.slice();
  [copy[idx], copy[next]] = [copy[next], copy[idx]];
  fieldSettingsDraft.fields = copy;
  renderFieldSettingsList();
}

function removeDraftField(idx) {
  const required = getRequiredFields(fieldSettingsDraft);
  if (required.includes(fieldSettingsDraft.fields[idx].name)) return;
  fieldSettingsDraft.fields.splice(idx, 1);
  renderFieldSettingsList();
}

function addCustomField() {
  const input = document.getElementById('newFieldName');
  const name = input.value.trim();
  if (!name) return;
  if (fieldSettingsDraft.fields.some(field => field.name === name)) {
    getNotifier().warning("⚠️ 此欄位已存在");
    return;
  }
  fieldSettingsDraft.fields.push({ name, enabled: true, custom: true });
  input.value = "";
  renderFieldSettingsList();
}

function applyInitialTemplateFields() {
  const templateType = fieldSettingsDraft.fieldTemplateType || getFieldTemplateType(currentTemplate);
  const template = initialFieldTemplates[templateType];
  const existing = new Set(fieldSettingsDraft.fields.map(field => field.name));
  let added = 0;
  template.defaultFields.forEach(name => {
    if (existing.has(name)) return;
    fieldSettingsDraft.fields.push({ name, enabled: true, custom: false });
    existing.add(name);
    added++;
  });
  fieldSettingsDraft.requiredFields = Array.from(new Set([
    ...(fieldSettingsDraft.requiredFields || []),
    ...template.requiredFields
  ]));
  renderFieldSettingsList();
  getNotifier().success(added ? `✅ 已補齊 ${added} 個初始欄位` : "✅ 初始模板欄位已完整");
}

function saveFieldSettings() {
  const previousHeaders = currentTableHeaders.slice();
  const previousRows = collectVisibleMatrix().slice(1);
  savePageFieldConfigLocally(fieldSettingsDraft);
  const nextHeaders = getEnabledFieldsFromConfig(currentPageFieldConfig);
  currentTableHeaders = nextHeaders;
  const nextRows = previousRows.map(row => remapRowToCurrentHeaders(row, previousHeaders));
  rerenderWithMatrix([nextHeaders, ...nextRows]);
  bootstrap.Modal.getInstance(document.getElementById('fieldSettingsModal'))?.hide();
  getNotifier().success("✅ 欄位設定已套用，請記得儲存變更");
  if (currentId) {
    fetchAPI("savePageFieldConfig", { id: currentId, pageFieldConfig: currentPageFieldConfig })
      .catch(err => console.warn("pageFieldConfig 暫存於瀏覽器，後端尚未儲存：", err));
  }
}


// ============================================================
//  🎯 網格互動（複製貼上等）
// ============================================================
function initGridInteraction() {
  const container = document.getElementById('rowsContainer');
  if (!container) return;

  // 監聽日期變更：若是連動講道的列，日期變更後自動重新抓取講道資訊
  container.addEventListener('change', (e) => {
    const target = e.target;
    if (target.classList.contains('grid-input')) {
      const cIdx = parseInt(target.dataset.c);
      const header = currentTableHeaders[cIdx];
      if (header && header.includes("日期")) {
        const dVal = target.value.trim();
        if (dVal !== "") {
          const slashDate = parseToSlashDate(dVal);
          if (slashDate) {
            target.value = slashDate;
            target.title = slashDate;
            
            // 講道連動邏輯
            const rowDiv = target.closest('.record-row');
            const sermonLinkColIdx = currentTableHeaders.indexOf("套用講道");
            if (sermonLinkColIdx !== -1) {
              const checkbox = rowDiv.querySelector(`input.sermon-link-checkbox[data-c="${sermonLinkColIdx}"]`);
              if (checkbox && checkbox.checked) {
                const sw = rowDiv.querySelector(`input.sermon-lang-switch[data-c="${sermonLinkColIdx}"]`);
                const langVal = sw && sw.checked ? "台語/聯合" : "華語/聯合";
                updateRowSermonState(rowDiv, langVal, slashDate);
              }
            }
            
            // 日期變更後，觸發即時排序
            sortRowsByDate();
          } else {
            getNotifier().error("❌ 日期不符格式，請按照yyyy/mm/dd進行建立");
            target.value = "";
            target.title = "";
            target.focus();
          }
        }
      } else if (header && header === "經文" && (currentTemplate === "小組聚會表模板" || currentTemplate === "團契聚會表模板") && window.BibleFormatter) {
        const val = target.value.trim();
        if (val !== "") {
          const formatted = window.BibleFormatter.format(val);
          if (formatted !== val) {
            target.value = formatted;
            target.title = formatted;
          }
        }
      }
    }
  });

  container.addEventListener('paste', (e) => {
    const target = e.target;
    if (!target.classList.contains('grid-input')) return;
    const pasteData = (e.clipboardData || window.clipboardData).getData('text');
    if (pasteData.includes('\t') || pasteData.includes('\n')) {
      e.preventDefault();
      const startC = parseInt(target.dataset.c);
      const currentRowDiv = target.closest('.record-row');
      let currentRowIndex = Array.from(container.children).indexOf(currentRowDiv);
      const rows = pasteData.split(/\r?\n/);
      if (rows[rows.length - 1] === "") rows.pop();

      for (let i = 0; i < rows.length; i++) {
        if (currentRowIndex + i >= container.children.length) addNewRow();
        const targetRowDiv = container.children[currentRowIndex + i];
        const cols = rows[i].split('\t');
        for (let j = 0; j < cols.length; j++) {
          const c = startC + j;
          const input = targetRowDiv.querySelector(`[data-c="${c}"]`);
          if (input) {
            if (input.classList.contains('sermon-link-checkbox')) {
              let val = cols[j].trim();
              const isChecked = val !== 'N' && val !== 'false' && val !== '0' && val !== '';
              input.checked = isChecked;
              
              let langVal = currentSermonSettings.sermonType;
              if (val === '華語/聯合' || val === '台語/聯合') {
                langVal = val;
              }
              const sw = targetRowDiv.querySelector(`input.sermon-lang-switch[data-c="${c}"]`);
              if (sw) {
                sw.checked = langVal === "台語/聯合";
              }
              onSermonCheckboxChange(input);
            } else if (input.type === 'checkbox') {
              input.checked = (cols[j] === 'Y' || cols[j] === 'true' || cols[j] === true);
              onSermonLinkChange(input);
            } else {
              let val = cols[j];
              if (currentTableHeaders[c] === "經文" && (currentTemplate === "小組聚會表模板" || currentTemplate === "團契聚會表模板") && window.BibleFormatter) {
                val = window.BibleFormatter.format(val);
              }
              input.value = val;
              input.title = val;
              input.classList.add('highlight');
              setTimeout(() => input.classList.remove('highlight'), 2000);
              
              // 貼上日期時，如果同列的講道連動是啟用狀態，觸發重算講道
              if (currentTableHeaders[c] && currentTableHeaders[c].includes("日期")) {
                const sermonLinkColIdx = currentTableHeaders.indexOf("套用講道");
                if (sermonLinkColIdx !== -1) {
                  const checkbox = targetRowDiv.querySelector(`input.sermon-link-checkbox[data-c="${sermonLinkColIdx}"]`);
                  if (checkbox && checkbox.checked) {
                    const sw = targetRowDiv.querySelector(`input.sermon-lang-switch[data-c="${sermonLinkColIdx}"]`);
                    const langVal = sw && sw.checked ? "台語/聯合" : "華語/聯合";
                    updateRowSermonState(targetRowDiv, langVal, val.trim());
                  }
                }
              }
            }
          }
        }
      }
      
      // paste 處理完所有行後，最後呼叫即時排序
      sortRowsByDate();
    }
  });
}


// ============================================================
//  📅 日期篩選
// ============================================================
function initDateQuickFilter() {
  const yearSelect = document.getElementById('dateQuickYear');
  const quarterSelect = document.getElementById('dateQuickQuarter');
  if (!yearSelect || !quarterSelect) return;

  const now = new Date();
  const currentYear = now.getFullYear();
  const currentQuarter = Math.floor(now.getMonth() / 3) + 1;
  const years = [];
  for (let year = currentYear - 2; year <= currentYear + 3; year++) years.push(year);

  yearSelect.innerHTML = years
    .map(year => `<option value="${year}" ${year === currentYear ? 'selected' : ''}>${year}</option>`)
    .join('');
  quarterSelect.value = String(currentQuarter);
}

function quarterDateRange(year, quarter) {
  const startMonth = (quarter - 1) * 3;
  const start = new Date(year, startMonth, 1);
  const end = new Date(year, startMonth + 3, 0);
  return {
    start: formatDateInputValue(start),
    end: formatDateInputValue(end)
  };
}

function formatDateInputValue(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function applyQuarterDateFilter() {
  const year = Number(document.getElementById('dateQuickYear').value);
  const quarter = Number(document.getElementById('dateQuickQuarter').value);
  const range = quarterDateRange(year, quarter);
  document.getElementById('startDate').value = range.start;
  document.getElementById('endDate').value = range.end;
  filterByDate();
}

function applyYearDateFilter() {
  const year = Number(document.getElementById('dateQuickYear').value);
  document.getElementById('startDate').value = `${year}-01-01`;
  document.getElementById('endDate').value = `${year}-12-31`;
  filterByDate();
}

function filterByDate() {
  if (window.event) window.event.preventDefault();
  const start = document.getElementById('startDate').value;
  const end = document.getElementById('endDate').value;
  const recordRows = document.querySelectorAll('.record-row');
  const dateColIdx = currentTableHeaders.findIndex(h => h.includes("日期"));

  if (dateColIdx === -1) {
    getNotifier().warning("⚠️ 找不到包含「日期」的欄位。");
    return;
  }

  let visibleCount = 0;
  recordRows.forEach(rowDiv => {
    const inputs = rowDiv.querySelectorAll('.grid-input');
    if (inputs.length === 0) return;
    const dateVal = inputs[dateColIdx].value.trim();
    let show = true;
    if (!start && !end) show = true;
    else if (!dateVal) show = false;
    else {
      // 將表格的斜線日期暫時轉為橫線，以便與原生的 YYYY-MM-DD input 值進行正確的大小比對
      const compareDate = dateVal.replace(/\//g, "-");
      if (start && compareDate < start) show = false;
      if (end && compareDate > end) show = false;
    }
    if (show) {
      rowDiv.classList.remove('hidden');
      visibleCount++;
    } else {
      rowDiv.classList.add('hidden');
    }
  });

  if (start || end) {
    getNotifier().success(`✅ 已篩選出 ${visibleCount} 筆資料`);
  }
}

function clearDateFilter() {
  if (window.event) window.event.preventDefault();
  document.getElementById('startDate').value = "";
  document.getElementById('endDate').value = "";
  document.querySelectorAll('.record-row').forEach(rowDiv => rowDiv.classList.remove('hidden'));
}


// ============================================================
//  🤖 AI 排班處理（改進的狀態提示）
// ============================================================
async function processAI() {
  if (window.event) window.event.preventDefault();

  // 防止重複提交
  if (getUIState().isLocked('processAI')) return;
  getUIState().lock('processAI');

  const rawText = document.getElementById('aiRawText').value.trim();
  if (!rawText) {
    getNotifier().warning("⚠️ 請貼上排班文字或輸入排班條件");
    getUIState().unlock('processAI');
    return;
  }

  const submitBtn = document.querySelector('#aiScheduleModal .btn-success');
  const textarea = document.getElementById('aiRawText');
  if (submitBtn) submitBtn.disabled = true;
  if (textarea) textarea.disabled = true;

  getNotifier().showLoading("🤖 AI 運算中，請稍候...");
  document.getElementById('aiStatus').innerText = "⏳ 處理中...";

  try {
    const resData = await fetchAPI("parseWithAI", {
      text: rawText,
      headers: currentTableHeaders,
      members: currentGroupMembers,
      groupPrompt: currentGroupPrompt + "\n" + currentAutoRoleRules,
      template: currentTemplate
    });

    fillTableWithData(resData);
    getNotifier().success("✅ AI 排班完成！");
    document.getElementById('aiStatus').innerText = "✅ 解析/排班完成！";
    document.getElementById('aiRawText').value = "";
  } catch (err) {
    handleAPIError(err);
    document.getElementById('aiStatus').innerText = "❌ 處理失敗，請重試";
  } finally {
    getNotifier().hideLoading();
    getUIState().unlock('processAI');
    if (submitBtn) submitBtn.disabled = false;
    if (textarea) textarea.disabled = false;
  }
}


// ============================================================
//  📝 填充表單資料
// ============================================================
function fillTableWithData(parsedRows) {
  const container = document.getElementById('rowsContainer');
  const dateColIdx = currentTableHeaders.findIndex(h => h.includes("日期"));

  // 預先快取目前所有 row 與其 inputsMap (cIdx -> input/checkbox/select)
  const rowCache = Array.from(container.querySelectorAll('.record-row')).map(rowDiv => ({
    rowDiv,
    inputsMap: Array.from(rowDiv.querySelectorAll('input, select')).reduce((map, input) => {
      const c = input.dataset.c;
      if (c !== undefined) map[c] = input;
      return map;
    }, {})
  }));

  parsedRows.forEach(rowData => {
    let target = null;
    let aiDate = rowData["日期"] || rowData[currentTableHeaders[dateColIdx]];

    // 檢查並格式化為 yyyy/mm/dd
    if (aiDate && dateColIdx !== -1) {
      const slashDate = parseToSlashDate(String(aiDate).trim());
      if (slashDate) {
        rowData["日期"] = slashDate;
        if (currentTableHeaders[dateColIdx] !== "日期") {
          rowData[currentTableHeaders[dateColIdx]] = slashDate;
        }
        aiDate = slashDate; // 同步更新供後續比對使用
      } else {
        getNotifier().error(`❌ 日期 "${aiDate}" 不符格式，請按照yyyy/mm/dd進行建立`);
        return; // 跳過此筆無效資料的填充
      }
    }

    // 先嘗試比對日期
    if (aiDate && dateColIdx !== -1) {
      target = rowCache.find(r => {
        const di = r.inputsMap[dateColIdx];
        return di && di.value.trim() === aiDate;
      });
    }

    // 找不到日期相符的 → 找完全空白的列
    if (!target) {
      target = rowCache.find(r => {
        return Object.values(r.inputsMap).every(input => {
          if (input.type === 'checkbox') return true; // 勾選框不視為內容填寫
          return input.value.trim() === "";
        });
      });
    }

    // 都沒有 → 新增一列並加入 cache
    if (!target) {
      addNewRow();
      const rowDiv = container.lastElementChild;
      target = {
        rowDiv,
        inputsMap: Array.from(rowDiv.querySelectorAll('input, select')).reduce((map, input) => {
          const c = input.dataset.c;
          if (c !== undefined) map[c] = input;
          return map;
        }, {})
      };
      rowCache.push(target);
    }

    currentTableHeaders.forEach((header, colIdx) => {
      const val = rowData[header];
      if (val !== undefined && val !== null && val !== "") {
        const input = target.rowDiv.querySelector(`input.sermon-link-checkbox[data-c="${colIdx}"]`) || target.inputsMap[colIdx];
        if (input) {
          if (input.classList.contains('sermon-link-checkbox')) {
            const isChecked = val !== 'N' && val !== 'false' && val !== '0' && val !== '';
            input.checked = isChecked;
            let langVal = currentSermonSettings.sermonType;
            if (val === '華語/聯合' || val === '台語/聯合') {
              langVal = val;
            }
            const sw = target.rowDiv.querySelector(`input.sermon-lang-switch[data-c="${colIdx}"]`);
            if (sw) {
              sw.checked = langVal === "台語/聯合";
            }
            onSermonCheckboxChange(input);
          } else if (input.type === 'checkbox') {
            input.checked = (val === 'Y' || val === 'true' || val === true);
            onSermonLinkChange(input);
          } else {
            let finalVal = val;
            if (header === "經文" && (currentTemplate === "小組聚會表模板" || currentTemplate === "團契聚會表模板") && window.BibleFormatter) {
              finalVal = window.BibleFormatter.format(val);
            }
            input.value = finalVal;
            input.title = finalVal;
            input.classList.add('highlight');
            setTimeout(() => input.classList.remove('highlight'), 2000);
          }
        }
      }
    });

    // 填充完後，若是連動講道的列，觸發重算/代入講道資訊
    const sermonLinkColIdx = currentTableHeaders.indexOf("套用講道");
    if (sermonLinkColIdx !== -1 && dateColIdx !== -1) {
      const checkbox = target.rowDiv.querySelector(`input.sermon-link-checkbox[data-c="${sermonLinkColIdx}"]`);
      const dInput = target.inputsMap[dateColIdx];
      if (checkbox && checkbox.checked && dInput && dInput.value) {
        const sw = target.rowDiv.querySelector(`input.sermon-lang-switch[data-c="${sermonLinkColIdx}"]`);
        const langVal = sw && sw.checked ? "台語/聯合" : "華語/聯合";
        updateRowSermonState(target.rowDiv, langVal, dInput.value.trim());
      }
    }
  });

  sortRowsByDate();
}


// ============================================================
//  💾 儲存資料
// ============================================================
async function saveData() {
  if (window.event) window.event.preventDefault();

  // 防止重複提交
  if (getUIState().isLocked('saveData')) return;
  getUIState().lock('saveData');

  getNotifier().showLoading("💾 儲存中...");

  try {
    const matrix = [currentTableHeaders];
    document.querySelectorAll('.record-row').forEach(rowDiv => {
      const row = [];
      currentTableHeaders.forEach((header, cIdx) => {
        if (header === "套用講道") {
          const checkbox = rowDiv.querySelector(`input.sermon-link-checkbox[data-c="${cIdx}"]`);
          const sw = rowDiv.querySelector(`input.sermon-lang-switch[data-c="${cIdx}"]`);
          const isChecked = checkbox && checkbox.checked;
          const langVal = sw && sw.checked ? "台語/聯合" : "華語/聯合";
          row.push(isChecked ? langVal : "N");
        } else {
          const input = rowDiv.querySelector(`input.grid-input[data-c="${cIdx}"]`);
          row.push(input ? input.value : "");
        }
      });
      // 只要有任何非「套用講道」的欄位有內容，即視為有效列
      if (row.some((v, idx) => currentTableHeaders[idx] !== "套用講道" && v.trim() !== "")) {
        matrix.push(row);
      }
    });

    while (matrix.length <= 50) matrix.push(Array(currentTableHeaders.length).fill(""));

    await fetchAPI("saveSheetData", { groupName: activeGroupName, matrix: matrix });

    getNotifier().success("✅ 儲存成功！");
  } catch (err) {
    handleAPIError(err);
  } finally {
    getNotifier().hideLoading();
    getUIState().unlock('saveData');
  }
}


// ============================================================
//  🔄 切換狀態
// ============================================================
async function toggleStatus(groupId, currentStatus) {
  if (window.event) window.event.preventDefault();

  if (getUIState().isLocked('toggleStatus')) return;
  getUIState().lock('toggleStatus');

  getNotifier().showLoading("🔄 更新狀態中...");

  try {
    await fetchAPI("toggleGroupStatus", { id: groupId, status: currentStatus });
    await loadAdminData();
  } catch (err) {
    handleAPIError(err);
  } finally {
    getNotifier().hideLoading();
    getUIState().unlock('toggleStatus');
  }
}


// ============================================================
//  🎨 UI 元件控制
// ============================================================
function showSection(id) {
  document.querySelectorAll('.card-custom').forEach(el => el.classList.add('hidden'));
  document.getElementById(id).classList.remove('hidden');
}


// ============================================================
//  ➕ 建立新分頁表單
// ============================================================
const createForm = document.getElementById('createGroupForm');
if (createForm) {
  createForm.onsubmit = async function(e) {
    e.preventDefault();

    if (getUIState().isLocked('createGroup')) return;
    getUIState().lock('createGroup');

    getNotifier().showLoading("建立中...");
    try {
      const fieldTemplateType = document.getElementById('templateSelect').value;
      const nextId = document.getElementById('newId').value.trim();
      const template = initialFieldTemplates[fieldTemplateType] || initialFieldTemplates["事工型模板"];
      const firstConfig = normalizeFieldConfig({
        fields: template.defaultFields.map(name => ({ name, enabled: true, custom: false })),
        requiredFields: template.requiredFields
      }, fieldTemplateType, nextId);
      await fetchAPI("createGroup", {
        id: nextId,
        name: document.getElementById('newName').value,
        template: getBackendTemplateForFieldType(fieldTemplateType),
        fieldTemplateType,
        pageFieldConfig: firstConfig
      });
      localStorage.setItem(getFieldConfigStorageKey(nextId), JSON.stringify(firstConfig));
      location.reload();
    } catch (err) {
      handleAPIError(err);
    } finally {
      getNotifier().hideLoading();
      getUIState().unlock('createGroup');
    }
  };
}


// ============================================================
//  💾 儲存小組規則
// ============================================================
async function saveGroupPrompt() {
  if (window.event) window.event.preventDefault();

  if (getUIState().isLocked('saveGroupPrompt')) return;
  getUIState().lock('saveGroupPrompt');

  const newPrompt = document.getElementById('groupPromptInput').value.trim();
  getNotifier().showLoading("💾 儲存規則中...");

  try {
    await fetchAPI("saveGroupPrompt", { id: currentId, prompt: newPrompt });
    currentGroupPrompt = newPrompt;
    getNotifier().success("✅ 專屬規則儲存成功！");
  } catch (err) {
    handleAPIError(err);
  } finally {
    getNotifier().hideLoading();
    getUIState().unlock('saveGroupPrompt');
  }
}


// ============================================================
//  📅 共用日期篩選器（modal 用：年度 + 季度 + 滾動近 3 個月）
// ============================================================
const _MS_FILTER_QUARTERS = [1, 2, 3, 4];

// 應該被視為「小組類聚會」的模板（小組總表 / 各小組佈告欄會包含這些）
const _MS_FELLOWSHIP_TEMPLATES = ['小組聚會表模板', '團契聚會表模板'];
// 各項服事總表合併同日期時，要丟掉的「來源」欄位
const _MS_META_COLS = ['分頁名稱', '模板類型', '聚會名稱', '聚會類別'];

// ---- 矩陣 / 物件互轉 ----
function _ms_matrixToObjects(matrix) {
  if (!matrix || matrix.length < 2) return [];
  const headers = matrix[0];
  return matrix.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function _ms_objectsToMatrix(objects, headerOrder, opts = {}) {
  // strict: 只用 headerOrder 列出的欄位，忽略 objects 中其他欄位
  const strict = opts.strict === true;
  if (!objects || objects.length === 0) return headerOrder ? [headerOrder.slice()] : [];
  let headers;
  if (strict && headerOrder) {
    headers = headerOrder.slice();
  } else {
    const seen = new Set();
    headers = [];
    if (headerOrder) headerOrder.forEach(h => { if (!seen.has(h)) { seen.add(h); headers.push(h); } });
    objects.forEach(obj => Object.keys(obj).forEach(k => { if (!seen.has(k)) { seen.add(k); headers.push(k); } }));
  }
  const rows = objects.map(obj => headers.map(h => obj[h] == null ? '' : obj[h]));
  return [headers, ...rows];
}

function _ms_filterMatrix(matrix, predicate) {
  if (!matrix || matrix.length < 2) return matrix ? matrix.slice() : [];
  const filtered = _ms_matrixToObjects(matrix).filter(predicate);
  return _ms_objectsToMatrix(filtered, matrix[0]);
}

// 合併兩個（或多個）矩陣，headers 取 union 並依首次出現的順序排列
function _ms_mergeMatrices(...matrices) {
  const objs = matrices.flatMap(m => _ms_matrixToObjects(m || []));
  const seen = new Set();
  const headerOrder = [];
  matrices.forEach(m => {
    if (m && m[0]) m[0].forEach(h => { if (!seen.has(h)) { seen.add(h); headerOrder.push(h); } });
  });
  return _ms_objectsToMatrix(objs, headerOrder);
}

// 同日期的多列合併成一列：每欄不同值用「\n」串接，並加上 (來源) 標記
// dropColumns 中的欄位會被移除（預設為「分頁名稱/模板類型/聚會名稱/聚會類別」）
function _ms_collapseByDate(matrix, opts = {}) {
  const dropCols = opts.dropColumns || _MS_META_COLS;
  if (!matrix || matrix.length < 2) return matrix ? matrix.slice() : [];

  const objects = _ms_matrixToObjects(matrix);
  const byDate = new Map();
  const sourceOf = obj => obj['分頁名稱'] || obj['聚會名稱'] || '';

  objects.forEach(obj => {
    const date = obj['日期'] || '';
    if (!date) return;
    if (!byDate.has(date)) byDate.set(date, new Map()); // colName -> Map<value, Set<source>>
    const cols = byDate.get(date);
    const src = sourceOf(obj);
    Object.keys(obj).forEach(k => {
      if (k === '日期' || dropCols.includes(k)) return;
      const v = obj[k];
      if (v == null || v === '' || v === '-') return;
      if (!cols.has(k)) cols.set(k, new Map());
      const valMap = cols.get(k);
      const sv = String(v);
      if (!valMap.has(sv)) valMap.set(sv, new Set());
      if (src) valMap.get(sv).add(src);
    });
  });

  const merged = [];
  Array.from(byDate.keys()).sort().forEach(date => {
    const obj = { '日期': date };
    byDate.get(date).forEach((valMap, k) => {
      const lines = Array.from(valMap.entries()).map(([value, sources]) => {
        const srcs = Array.from(sources).filter(Boolean);
        return srcs.length > 0 ? `${value} (${srcs.join('、')})` : value;
      });
      obj[k] = lines.join('\n');
    });
    merged.push(obj);
  });

  // 欄位順序：日期優先，後接原始 matrix 的欄位（去掉 drop 跟日期）
  const headerOrder = ['日期', ...matrix[0].filter(h => h !== '日期' && !dropCols.includes(h))];
  return _ms_objectsToMatrix(merged, headerOrder);
}

// ---- getAggregatedReport 雙桶快取（30 秒內重複開 modal 不重複 fetch）----
let _ms_aggCache = { sm: null, oth: null, ts: 0 };
const _MS_AGG_TTL = 30 * 1000;

async function _ms_fetchBothAggregated() {
  const now = Date.now();
  if (_ms_aggCache.sm && (now - _ms_aggCache.ts) < _MS_AGG_TTL) return _ms_aggCache;
  // 用 allSettled 允許單邊失敗：例如 'others' 失敗時，'小組總表' 仍能顯示（只是少了 團契 那塊）
  const [smRes, othRes] = await Promise.allSettled([
    fetchAPI('getAggregatedReport', { type: 'smallGroup' }),
    fetchAPI('getAggregatedReport', { type: 'others' })
  ]);
  const sm  = smRes.status === 'fulfilled'  ? (smRes.value  || []) : [];
  const oth = othRes.status === 'fulfilled' ? (othRes.value || []) : [];
  if (smRes.status === 'rejected')  console.warn('[MinistrySchedule] 小組總表抓取失敗：', smRes.reason);
  if (othRes.status === 'rejected') console.warn('[MinistrySchedule] 各項總表抓取失敗：', othRes.reason);
  _ms_aggCache = { sm, oth, ts: now };
  return _ms_aggCache;
}

function _ms_getDateColIdx(headers) {
  if (!Array.isArray(headers)) return -1;
  return headers.findIndex(h => String(h || '').includes('日期'));
}

function _ms_yearsFromMatrix(matrix, dateColIdx) {
  if (!matrix || matrix.length < 2 || dateColIdx < 0) return [new Date().getFullYear()];
  const ys = new Set();
  for (let i = 1; i < matrix.length; i++) {
    const d = new Date(matrix[i][dateColIdx]);
    if (!isNaN(d)) ys.add(d.getFullYear());
  }
  const arr = Array.from(ys).sort((a, b) => b - a);
  const cur = new Date().getFullYear();
  if (!arr.includes(cur)) arr.unshift(cur);
  return arr;
}

function _ms_currentQuarter() {
  return Math.floor(new Date().getMonth() / 3) + 1;
}

function _ms_rollingWindow() {
  // 預設範圍：往前 1 個月、往後 2 個月
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const start = new Date(today);
  start.setMonth(start.getMonth() - 1);
  const end = new Date(today);
  end.setMonth(end.getMonth() + 2);
  end.setHours(23, 59, 59, 999);
  return [start, end];
}

function _ms_applyDateFilter(matrix, dateColIdx, mode, year, quarter) {
  if (!matrix || matrix.length < 2) return matrix ? matrix.slice() : [];
  if (dateColIdx < 0) return matrix.slice();
  const headers = matrix[0];
  let predicate;
  if (mode === 'rolling') {
    const [start, end] = _ms_rollingWindow();
    predicate = d => d >= start && d <= end;
  } else {
    const m1 = (quarter - 1) * 3 + 1;
    predicate = d => d.getFullYear() === year && (d.getMonth() + 1) >= m1 && (d.getMonth() + 1) <= m1 + 2;
  }
  const filtered = matrix.slice(1).filter(row => {
    const d = new Date(row[dateColIdx]);
    if (isNaN(d)) return false;
    return predicate(d);
  });
  return [headers, ...filtered];
}

function _ms_buildTableHtml(matrix, opts = {}) {
  const minWidth = opts.minWidth || 800;
  if (!matrix || matrix.length <= 1) {
    return '<p class="text-center text-muted my-4">此範圍內沒有資料，請改選其他季度</p>';
  }
  let html = `<table class="table table-bordered table-hover text-center align-middle m-0" style="min-width: ${minWidth}px;"><thead><tr>`;
  matrix[0].forEach(h => html += `<th class="bg-light" style="position: sticky; top: 0; z-index: 10; outline: 1px solid #dee2e6;">${h}</th>`);
  html += '</tr></thead><tbody>';
  // td 用 white-space: pre-line 讓合併日期時的「\n 換行」能正確呈現多行
  for (let i = 1; i < matrix.length; i++) {
    html += '<tr>';
    matrix[i].forEach(cell => html += `<td style="white-space: pre-line; vertical-align: top;">${cell || "-"}</td>`);
    html += '</tr>';
  }
  html += '</tbody></table>';
  return html;
}

function _ms_buildCardsHtml(matrix) {
  if (!matrix || matrix.length <= 1) {
    return '<p class="text-center text-muted my-4">此範圍內沒有資料，請改選其他季度</p>';
  }
  const headers = matrix[0];
  const dateColIdx = _ms_getDateColIdx(headers);
  const topicColIdx = headers.findIndex(h => h === '主題' || h === '聚會名稱');
  const verseColIdx = headers.findIndex(h => h === '經文');
  const locColIdx = headers.findIndex(h => h === '地點');
  const sermonColIdx = headers.findIndex(h => h === '講道連動');

  const excludeFields = ['日期', '主題', '聚會名稱', '經文', '地點', '講道連動', '分頁名稱', '模板類型', '聚會類別'];

  let cardsHtml = '<div class="glass-board-container">';

  for (let i = 1; i < matrix.length; i++) {
    const row = matrix[i];
    const dateVal = dateColIdx >= 0 ? row[dateColIdx] : '';

    const dateObj = new Date(dateVal);
    let day = '';
    let yearMonth = '';
    let weekDay = '';
    if (!isNaN(dateObj)) {
      day = String(dateObj.getDate()).padStart(2, '0');
      yearMonth = `${dateObj.getFullYear()}年 ${String(dateObj.getMonth() + 1).padStart(2, '0')}月`;
      weekDay = '星期' + ['日', '一', '二', '三', '四', '五', '六'][dateObj.getDay()];
    } else {
      day = '📅';
      yearMonth = dateVal || '聚會日';
      weekDay = '聚會日';
    }

    const topic = topicColIdx >= 0 ? row[topicColIdx] : '';
    const verse = verseColIdx >= 0 ? row[verseColIdx] : '';
    const location = locColIdx >= 0 ? row[locColIdx] : '';
    const hasSermon = sermonColIdx >= 0 && (String(row[sermonColIdx]).toUpperCase() === 'TRUE' || row[sermonColIdx] === true || String(row[sermonColIdx]) === '1');

    const duties = [];
    headers.forEach((h, idx) => {
      if (excludeFields.includes(h)) return;
      const val = row[idx];
      if (val && val !== '-' && val !== '') {
        duties.push({ role: h, name: val });
      }
    });

    const dutyItemsHtml = duties.map(d => `
      <div class="glass-duty-item">
        <span class="glass-duty-role">👤 ${d.role}</span>
        <span class="glass-duty-name">${d.name}</span>
      </div>
    `).join('');

    const metaItems = [];
    if (verse) metaItems.push(`<span>📖 <b>經文：</b>${verse}</span>`);
    if (location) metaItems.push(`<span>📍 <b>地點：</b>${location}</span>`);
    const metaHtml = metaItems.length > 0 ? `<div class="glass-meta">${metaItems.join('')}</div>` : '';

    cardsHtml += `
      <div class="glass-card color-ramp-${(i - 1) % 5}">
        <div class="glass-date">
          <div class="day">${day}</div>
          <div class="month-year">${yearMonth}</div>
          <div class="weekday">${weekDay}</div>
        </div>
        <div class="glass-content">
          <div class="glass-topic-line">
            <h3 class="glass-topic">${topic || (currentTemplate === '團契聚會表模板' ? '團契聚會' : '小組聚會')}</h3>
          </div>
          <div class="glass-duties">
            ${dutyItemsHtml || '<div class="text-muted small">一般聚會，無需特別服事</div>'}
          </div>
          ${metaHtml}
        </div>
      </div>
    `;
  }

  cardsHtml += '</div>';
  return cardsHtml;
}

/**
 * 在指定容器內渲染「年度+季度」篩選器 + 表格或卡片。
 */
function _ms_renderFilterableTable({ container, fullMatrix, tableMinWidth, renderMode = 'table', onFilteredChange }) {
  const headers = (fullMatrix && fullMatrix[0]) || [];
  const dateColIdx = _ms_getDateColIdx(headers);
  const noDateCol = dateColIdx < 0;
  const years = _ms_yearsFromMatrix(fullMatrix, dateColIdx);
  const today = new Date();
  const state = {
    mode: 'rolling',
    year: years.includes(today.getFullYear()) ? today.getFullYear() : years[0],
    quarter: _ms_currentQuarter()
  };

  const filterBar = noDateCol ? '' : `
    <div class="ms-filter-toolbar d-flex align-items-center gap-2 mb-2 flex-wrap p-2 bg-light rounded border">
      <span class="text-muted small fw-bold">📅 顯示範圍：</span>
      <button type="button" id="ms-filter-rolling" class="btn btn-sm btn-primary">近 3 個月</button>
      <span class="text-muted">|</span>
      <select id="ms-filter-year" class="form-select form-select-sm" style="width: auto;">
        ${years.map(y => `<option value="${y}" ${y === state.year ? 'selected' : ''}>${y} 年</option>`).join('')}
      </select>
      <select id="ms-filter-quarter" class="form-select form-select-sm" style="width: auto;">
        ${_MS_FILTER_QUARTERS.map(q => `<option value="${q}" ${q === state.quarter ? 'selected' : ''}>Q${q} (${(q - 1) * 3 + 1}~${q * 3}月)</option>`).join('')}
      </select>
      <span id="ms-filter-status" class="text-muted small ms-auto"></span>
    </div>`;

  container.innerHTML = filterBar + `<div class="table-responsive" id="ms-filter-table" style="max-height: 60vh; overflow-y: auto;"></div>`;

  function rerender() {
    const filtered = noDateCol
      ? fullMatrix.slice()
      : _ms_applyDateFilter(fullMatrix, dateColIdx, state.mode, state.year, state.quarter);
    
    if (renderMode === 'cards') {
      document.getElementById('ms-filter-table').innerHTML = _ms_buildCardsHtml(filtered);
    } else {
      document.getElementById('ms-filter-table').innerHTML = _ms_buildTableHtml(filtered, { minWidth: tableMinWidth });
    }
    
    const statusEl = document.getElementById('ms-filter-status');
    if (statusEl) {
      const recordCount = Math.max(filtered.length - 1, 0);
      statusEl.innerText = state.mode === 'rolling'
        ? `共 ${recordCount} 筆（近 3 個月）`
        : `共 ${recordCount} 筆（${state.year} Q${state.quarter}）`;
    }
    const rollingBtn = document.getElementById('ms-filter-rolling');
    if (rollingBtn) {
      rollingBtn.classList.toggle('btn-primary', state.mode === 'rolling');
      rollingBtn.classList.toggle('btn-outline-primary', state.mode !== 'rolling');
    }
    if (typeof onFilteredChange === 'function') onFilteredChange(filtered);
  }

  if (!noDateCol) {
    document.getElementById('ms-filter-rolling').onclick = () => { state.mode = 'rolling'; rerender(); };
    document.getElementById('ms-filter-year').onchange = e => { state.mode = 'quarter'; state.year = +e.target.value; rerender(); };
    document.getElementById('ms-filter-quarter').onchange = e => { state.mode = 'quarter'; state.quarter = +e.target.value; rerender(); };
  }

  rerender();
}


// ============================================================
//  📋 預覽佈告欄
// ============================================================
function showBulletinBoard() {
  if (window.event) window.event.preventDefault();

  // 從目前的編輯表單擷取完整 matrix（仍尊重 .hidden 過濾，例如編輯模式的日期區間）
  const matrix = [currentTableHeaders];
  document.querySelectorAll('.record-row').forEach(rowDiv => {
    if (rowDiv.classList.contains('hidden')) return;
    const row = Array.from(rowDiv.querySelectorAll('.grid-input')).map(i => i.value.trim());
    if (row.some(v => v !== "")) matrix.push(row);
  });

  _currentBulletinFiltered = matrix;
  _ms_renderFilterableTable({
    container: document.getElementById('bulletinContent'),
    fullMatrix: matrix,
    tableMinWidth: 800,
    renderMode: 'cards',
    onFilteredChange: filtered => { _currentBulletinFiltered = filtered; }
  });

  document.getElementById('bulletinModalLabel').innerText = `📋 ${activeGroupName} - 排班佈告欄`;

  const closeBtn = document.getElementById('modalCloseBtn');
  if (isEditorUnlocked) {
    closeBtn.innerText = "✖ 關閉預覽";
    closeBtn.classList.replace('btn-warning', 'btn-secondary');
  }

  if (!bulletinModalInstance) {
    bulletinModalInstance = new bootstrap.Modal(document.getElementById('bulletinModal'), {
      backdrop: 'static',
      keyboard: false
    });
  }
  bulletinModalInstance.show();
}


// ============================================================
//  🔓 解鎖編輯模式 (本地驗證與解密版)
// ============================================================

let unlockVerifyModalInstance = null;

async function closeModalOrUnlock() {
  if (window.event) window.event.preventDefault();
  if (isEditorUnlocked) {
    bulletinModalInstance.hide();
  } else {
    if (!unlockVerifyModalInstance) {
      unlockVerifyModalInstance = new bootstrap.Modal(document.getElementById('unlockVerifyModal'), {
        backdrop: 'static',
        keyboard: false
      });
      // 註冊 Enter 鍵監聽
      document.getElementById('unlockVerifyCode').addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          submitUnlockVerifyCode();
        }
      });
    }
    document.getElementById('unlockVerifyCode').value = '';
    document.getElementById('unlockVerifyError').classList.add('hidden');
    unlockVerifyModalInstance.show();
    
    // 延遲聚焦以支援 CSS 動畫完成
    setTimeout(() => {
      document.getElementById('unlockVerifyCode').focus();
    }, 500);
  }
}

async function submitUnlockVerifyCode() {
  const pwd = document.getElementById('unlockVerifyCode').value.trim();
  if (!pwd) {
    getNotifier().warning("⚠️ 請輸入專屬 ID");
    return;
  }

  const errorEl = document.getElementById('unlockVerifyError');
  errorEl.classList.add('hidden');

  const decryptedId = decryptGroupCode(currentId);
  const inputCode = pwd.toUpperCase();
  const isMaster = (inputCode === "LK31"); // ADMIN_CODE
  const isMatch = (inputCode === decryptedId.toUpperCase());

  if (isMaster || isMatch) {
    isEditorUnlocked = true;
    getSessionMgr().setUnlocked(currentId);
    
    if (unlockVerifyModalInstance) unlockVerifyModalInstance.hide();
    if (bulletinModalInstance) bulletinModalInstance.hide();
    
    getNotifier().success("✅ 編輯模式已啟用");
  } else {
    errorEl.classList.remove('hidden');
    const inputEl = document.getElementById('unlockVerifyCode');
    inputEl.value = '';
    inputEl.focus();
  }
}

window.submitUnlockVerifyCode = submitUnlockVerifyCode;


// ============================================================
//  📥 下載 Excel
// ============================================================
function downloadExcel() {
  if (window.event) window.event.preventDefault();

  // 若 modal 開啟過且已套用篩選，下載篩選後的內容；否則取整份編輯表單的可見資料
  let matrix;
  if (Array.isArray(_currentBulletinFiltered) && _currentBulletinFiltered.length > 1) {
    matrix = _currentBulletinFiltered;
  } else {
    matrix = [currentTableHeaders];
    document.querySelectorAll('.record-row').forEach(rowDiv => {
      if (rowDiv.classList.contains('hidden')) return;
      const row = Array.from(rowDiv.querySelectorAll('.grid-input')).map(i => i.value.trim());
      if (row.some(v => v !== "")) matrix.push(row);
    });
  }
  if (matrix.length === 1) {
    getNotifier().warning("⚠️ 目前沒有資料可以下載！");
    return;
  }

  const ws = XLSX.utils.aoa_to_sheet(matrix);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "佈告欄");

  const today = new Date().toISOString().slice(0, 10).replace(/-/g, "");
  XLSX.writeFile(wb, `${activeGroupName}_排班表_${today}.xlsx`);
  getNotifier().success("✅ Excel 已下載");
}


// ============================================================
//  👥 管理同工名單
// ============================================================
function openMemberModal() {
  if (window.event) window.event.preventDefault();
  const roleSelect = document.getElementById('newMemberRole');
  if (currentTemplate === "新家人服事表模板") {
    roleSelect.classList.remove('hidden');
  } else {
    roleSelect.classList.add('hidden');
  }
  renderMemberList();
  new bootstrap.Modal(document.getElementById('memberModal')).show();
}

function renderMemberList() {
  const listEl = document.getElementById('memberList');
  listEl.innerHTML = localCustomMembers.map((m, idx) => `
    <li class="list-group-item d-flex justify-content-between align-items-center">
      <div>
        <span class="fw-bold">${m.name}</span>
        ${currentTemplate === "新家人服事表模板" ? `<span class="badge bg-secondary ms-2">${m.role}</span>` : ''}
      </div>
      <button type="button" class="btn btn-sm btn-danger" onclick="deleteMember(${idx})">刪除</button>
    </li>
  `).join('');
}

function addMember() {
  if (window.event) window.event.preventDefault();
  const nameInput = document.getElementById('newMemberName');
  const roleSelect = document.getElementById('newMemberRole');
  const rawText = nameInput.value.trim();
  const role = roleSelect.value;

  if (!rawText) {
    getNotifier().warning("⚠️ 請輸入姓名！");
    return;
  }

  const names = rawText.split(/[\n,，、\s]+/).filter(n => n.trim() !== "");
  let addedCount = 0;
  let dupCount = 0;

  names.forEach(name => {
    if (localCustomMembers.find(m => m.name === name)) {
      dupCount++;
    } else {
      localCustomMembers.push({
        name: name,
        role: currentTemplate === "新家人服事表模板" ? role : "一般同工"
      });
      addedCount++;
    }
  });

  nameInput.value = "";
  renderMemberList();

  if (addedCount > 1) {
    getNotifier().success(`✅ 成功批量新增 ${addedCount} 筆名單！` + (dupCount > 0 ? `\n⚠️ 另有 ${dupCount} 筆已存在被自動略過。` : ""));
  } else if (addedCount === 0 && dupCount > 0) {
    getNotifier().warning("⚠️ 您輸入的名字都已經在名單中囉！");
  } else {
    getNotifier().success("✅ 已新增");
  }
}

function deleteMember(idx) {
  if (window.event) window.event.preventDefault();
  localCustomMembers.splice(idx, 1);
  renderMemberList();
}

async function saveMembersToServer() {
  if (window.event) window.event.preventDefault();

  if (getUIState().isLocked('saveMembersToServer')) return;
  getUIState().lock('saveMembersToServer');

  getNotifier().showLoading("💾 儲存名單中...");

  try {
    await fetchAPI("saveGroupMembers", { id: currentId, members: localCustomMembers });
    getNotifier().success("✅ 名單儲存成功！");

    const memberModalEl = document.getElementById('memberModal');
    if (memberModalEl) {
      const memberModal = bootstrap.Modal.getInstance(memberModalEl);
      if (memberModal) memberModal.hide();
    }

    getNotifier().showLoading("🔄 更新畫面中...");
    const freshConfig = await fetchAPI('getPageConfig', { id: currentId });
    renderTable(freshConfig);
    getNotifier().success("✅ 畫面已更新");
  } catch (err) {
    handleAPIError(err);
  } finally {
    getNotifier().hideLoading();
    getUIState().unlock('saveMembersToServer');
  }
}


// ============================================================
//  🧑‍🤝‍🧑 設定小組組員身分（小組/團契模板專用，跳過小組系統直接編輯 master）
// ============================================================
let _groupRoleEditingMembers = [];

async function openGroupRoleModal() {
  if (window.event) window.event.preventDefault();
  if (!activeGroupName) {
    getNotifier().warning("⚠️ 尚未載入小組資料");
    return;
  }

  document.getElementById('groupRoleModalTitle').innerText = activeGroupName;
  const listEl = document.getElementById('groupRoleList');
  listEl.innerHTML = '<li class="list-group-item text-center text-muted"><div class="spinner-border spinner-border-sm me-2"></div>載入中...</li>';

  new bootstrap.Modal(document.getElementById('groupRoleModal')).show();

  try {
    const data = await fetchAPI('getGroupMembers', { groupName: activeGroupName });
    if (!data || !data.isInitialized) {
      listEl.innerHTML = '<li class="list-group-item text-center text-warning py-3">⚠️ 此小組尚未初始化名單<br><small class="text-muted">請先去小組系統初始化</small></li>';
      _groupRoleEditingMembers = [];
      return;
    }
    _groupRoleEditingMembers = (data.members || []).map(m => ({ ...m }));
    renderGroupRoleList();
  } catch (err) {
    handleAPIError(err);
    listEl.innerHTML = `<li class="list-group-item text-center text-danger py-3">❌ 載入失敗</li>`;
  }
}

function renderGroupRoleList() {
  const listEl = document.getElementById('groupRoleList');
  if (_groupRoleEditingMembers.length === 0) {
    listEl.innerHTML = '<li class="list-group-item text-center text-muted py-3">此小組沒有任何成員</li>';
    return;
  }
  listEl.innerHTML = _groupRoleEditingMembers.map((m, idx) => {
    const roleClasses = {
      '核心同工': 'role-core',
      '一般同工': 'role-active',
      '陪伴同工': 'role-companion',
      '小羊':     'role-sheep'
    };
    const selClasses = {
      '核心同工': 'sel-core',
      '一般同工': 'sel-active',
      '陪伴同工': 'sel-companion',
      '小羊':     'sel-sheep'
    };
    const bClass = roleClasses[m.role] || 'role-sheep';
    const sClass = selClasses[m.role] || 'sel-sheep';
    const nickname = (m.nickname || '').trim();
    return `
      <li class="list-group-item d-flex justify-content-between align-items-center role-item ${bClass}">
        <div>
          <span class="fw-bold">${m.name}</span>
          ${nickname ? `<small class="text-muted ms-2">(${nickname})</small>` : ''}
        </div>
        <select class="form-select form-select-sm role-select ${sClass}" style="width: 150px;"
                onchange="updateGroupRoleByIdx(${idx}, this.value)">
          <option value="核心同工" ${m.role === '核心同工' ? 'selected' : ''}>⭐ 核心同工</option>
          <option value="一般同工" ${m.role === '一般同工' ? 'selected' : ''}>👤 一般同工</option>
          <option value="小羊"     ${m.role === '小羊'     ? 'selected' : ''}>🐑 小羊</option>
          <option value="陪伴同工" ${m.role === '陪伴同工' ? 'selected' : ''}>👥 陪伴同工</option>
        </select>
      </li>
    `;
  }).join('');
}

function updateGroupRoleByIdx(idx, newRole) {
  if (_groupRoleEditingMembers[idx]) {
    _groupRoleEditingMembers[idx].role = newRole;
    // 即時更新左側 border 顏色
    renderGroupRoleList();
  }
}

async function saveGroupRoles() {
  if (window.event) window.event.preventDefault();
  if (getUIState().isLocked('saveGroupRoles')) return;
  getUIState().lock('saveGroupRoles');

  getNotifier().showLoading("💾 儲存身分中...");
  try {
    await fetchAPI('updateGroupMemberRoles', {
      groupName: activeGroupName,
      members: _groupRoleEditingMembers
    });
    getNotifier().success("✅ 身分已更新！小組系統與 AI 排班會即時同步");

    const modal = bootstrap.Modal.getInstance(document.getElementById('groupRoleModal'));
    if (modal) modal.hide();

    // 重新讀取頁面設定（讓 AI 規則更新到新身分）
    getNotifier().showLoading("🔄 更新畫面中...");
    const freshConfig = await fetchAPI('getPageConfig', { id: currentId });
    renderTable(freshConfig);
  } catch (err) {
    handleAPIError(err);
  } finally {
    getNotifier().hideLoading();
    getUIState().unlock('saveGroupRoles');
  }
}


// ============================================================
//  📊 彙整報表
// ============================================================
async function showAggregatedReport(type) {
  if (window.event) window.event.preventDefault();

  if (getUIState().isLocked('showAggregatedReport')) return;
  getUIState().lock('showAggregatedReport');

  getNotifier().showLoading("📊 彙整資料中，這可能需要幾秒鐘...");

  try {
    // 同時抓 smallGroup + others，用模板類型重新分桶（團契歸到小組那邊）
    const { sm: smRaw, oth: othRaw } = await _ms_fetchBothAggregated();

    let matrix;
    if (type === 'smallGroup') {
      // 小組聚會總表 = 後端 smallGroup + others 中模板類型 = 團契聚會表模板 的列
      const fellowshipFromOthers = _ms_filterMatrix(othRaw, obj => obj['模板類型'] === '團契聚會表模板');
      const merged = _ms_mergeMatrices(smRaw, fellowshipFromOthers);

      if (merged.length > 1) {
        // 使用者指定的欄位白名單（其他欄位全部丟掉）
        // 「話語分享(講員)」是把小組的「話語分享」與團契的「講員」合併成同一欄
        const finalHeaders = ['日期', '分頁名稱', '破冰', '敬拜', '主題', '經文', '地點', '話語分享(講員)', '司會'];

        const objs = _ms_matrixToObjects(merged).map(obj => {
          const speak = (obj['話語分享'] == null ? '' : String(obj['話語分享'])).trim();
          const speaker = (obj['講員'] == null ? '' : String(obj['講員'])).trim();
          // 同列通常只會其中一個有值（小組 vs 團契），兩個都有時用 / 串接
          const combined = speak && speaker && speak !== speaker ? `${speak} / ${speaker}` : (speak || speaker);
          return { ...obj, '話語分享(講員)': combined };
        });

        // 依日期升冪排序（用 Date 解析以容忍不同日期格式）
        objs.sort((a, b) => {
          const da = new Date(a['日期'] || 0).getTime() || 0;
          const db = new Date(b['日期'] || 0).getTime() || 0;
          return da - db;
        });

        matrix = _ms_objectsToMatrix(objs, finalHeaders, { strict: true });
      } else {
        matrix = merged;
      }
    } else {
      // 各項服事總表 = others 排除團契後，依日期合併（同欄不同值用換行串接 + 來源標記）
      const withoutFellowship = _ms_filterMatrix(othRaw, obj => obj['模板類型'] !== '團契聚會表模板');
      matrix = _ms_collapseByDate(withoutFellowship);
    }

    if (!matrix || matrix.length <= 1) {
      getNotifier().warning("⚠️ 目前還沒有建立任何資料，或是資料都是空的喔！");
      return;
    }

    const title = type === 'smallGroup' ? '📊 所有小組聚會總表' : '📊 教會各項服事總表';
    let currentFiltered = matrix;

    _ms_renderFilterableTable({
      container: document.getElementById('aggregatedReportContent'),
      fullMatrix: matrix,
      tableMinWidth: 1200,
      onFilteredChange: filtered => { currentFiltered = filtered; }
    });

    document.getElementById('aggregatedReportModalLabel').innerText = title;
    document.getElementById('downloadAggregatedBtn').onclick = () => downloadAggregatedExcel(currentFiltered, title);

    new bootstrap.Modal(document.getElementById('aggregatedReportModal')).show();
  } catch (err) {
    handleAPIError(err);
  } finally {
    getNotifier().hideLoading();
    getUIState().unlock('showAggregatedReport');
  }
}


// ============================================================
//  📥 下載彙整報表 Excel
// ============================================================
function downloadAggregatedExcel(matrix, fileName) {
  const ws = XLSX.utils.aoa_to_sheet(matrix);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "彙整總表");

  const today = new Date().toISOString().slice(0, 10).replace(/-/g, "");
  XLSX.writeFile(wb, `${fileName}_${today}.xlsx`);
  getNotifier().success("✅ Excel 已下載");
}


// ============================================================
//  📥 匯出空白 Excel 模板
// ============================================================
function exportBlankTemplate() {
  if (!currentTableHeaders || currentTableHeaders.length === 0) {
    getNotifier().warning("⚠️ 找不到表格標題，請先載入排班表！");
    return;
  }
  const blankRow = currentTableHeaders.map(() => "");
  const data = [currentTableHeaders, blankRow, blankRow, blankRow, blankRow, blankRow];
  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "填寫模板");
  XLSX.writeFile(wb, `${activeGroupName}_Excel填寫模板.xlsx`);
  getNotifier().success("✅ Excel 模板已下載！請填寫日期後匯入。");
}


// ============================================================
//  📅 容錯日期解析（支援民國曆、西元、兩位年份、僅月日）
// ============================================================
function parseGregorianDate(rawStr) {
  if (!rawStr || typeof rawStr !== 'string') return null;

  const s = rawStr
    .replace(/[（(][一二三四五六日][）)]/g, '')
    .replace(/[（(][A-Za-z]{3}[）)]/gi, '')
    .replace(/星期[一二三四五六日]/g, '')
    .trim();

  const parts = s.match(/\d+/g);
  if (!parts) return null;

  let year, month, day;

  if (parts.length >= 3) {
    const p1 = parseInt(parts[0]);
    const p2 = parseInt(parts[1]);
    const p3 = parseInt(parts[2]);

    let rawYear, rawMonth, rawDay;
    if (p1 > 31) {
      rawYear = p1; rawMonth = p2; rawDay = p3;
    } else if (p3 > 31) {
      rawYear = p3; rawMonth = p2; rawDay = p1;
    } else {
      rawYear = p1; rawMonth = p2; rawDay = p3;
    }

    // <= 99: 2-digit Gregorian abbrev (26 → 2026)
    // 100-200: ROC (民國) year (115 → 2026)
    // > 200: 4-digit Gregorian
    if (rawYear <= 99) {
      year = 2000 + rawYear;
    } else if (rawYear <= 200) {
      year = rawYear + 1911;
    } else {
      year = rawYear;
    }
    month = rawMonth;
    day = rawDay;

  } else if (parts.length === 2) {
    year = new Date().getFullYear();
    month = parseInt(parts[0]);
    day = parseInt(parts[1]);
  } else {
    return null;
  }

  if (!year || !month || !day) return null;
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;

  return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}


// ============================================================
//  📤 匯入 Excel 填寫
// ============================================================
async function importExcelFile(input) {
  const file = input.files[0];
  if (!file) return;

  getNotifier().showLoading("⏳ 正在解析 Excel...");

  try {
    const buffer = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = e => resolve(e.target.result);
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });

    const wb = XLSX.read(buffer, { type: "array", cellDates: true, cellNF: false, cellText: false });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

    if (!rows || rows.length < 2) {
      getNotifier().warning("⚠️ Excel 檔案沒有資料列！");
      return;
    }

    // Normalize headers for case-insensitive, space-agnostic matching
    const normalize = h => String(h || "").trim().toLowerCase().replace(/\s/g, "");
    const excelHeaders = rows[0].map(normalize);
    const localHeaders = currentTableHeaders.map(normalize);

    // Build mapping: excelColIdx → localColIdx
    const colMap = {};
    excelHeaders.forEach((eh, ei) => {
      const li = localHeaders.findIndex(lh => lh === eh);
      if (li !== -1) colMap[ei] = li;
    });

    const dateLocalIdx = currentTableHeaders.findIndex(h => h.includes("日期"));
    if (dateLocalIdx === -1) {
      getNotifier().error("❌ 找不到「日期」欄位！");
      return;
    }

    // Find which Excel column maps to the local date column
    let dateExcelIdx = -1;
    for (const [ei, li] of Object.entries(colMap)) {
      if (parseInt(li) === dateLocalIdx) { dateExcelIdx = parseInt(ei); break; }
    }

    const parsedRows = [];
    let skippedCount = 0;

    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || row.every(cell => cell === "" || cell == null)) continue;

      // Parse date value
      let dateStr = null;
      if (dateExcelIdx !== -1) {
        const rawDate = row[dateExcelIdx];
        if (rawDate instanceof Date) {
          const y = rawDate.getFullYear();
          const m = String(rawDate.getMonth() + 1).padStart(2, '0');
          const d = String(rawDate.getDate()).padStart(2, '0');
          dateStr = `${y}-${m}-${d}`;
        } else if (rawDate !== "" && rawDate != null) {
          dateStr = parseGregorianDate(String(rawDate));
        }
      }

      if (!dateStr) {
        skippedCount++;
        getNotifier().error(`❌ 日期 "${row[dateExcelIdx] || ''}" 不符格式，請按照yyyy/mm/dd進行建立`);
        console.warn(`[importExcel] 第 ${r + 1} 列日期無效或缺失，已略過`, row);
        continue;
      }
      
      const slashDate = parseToSlashDate(dateStr);
      if (!slashDate) {
        skippedCount++;
        getNotifier().error(`❌ 日期 "${row[dateExcelIdx] || ''}" 不符格式，請按照yyyy/mm/dd進行建立`);
        continue;
      }
      dateStr = slashDate;

      // Build row object keyed by local header names
      const rowObj = {};
      for (const [ei, li] of Object.entries(colMap)) {
        const header = currentTableHeaders[li];
        let val = row[parseInt(ei)];
        if (val instanceof Date) {
          const y = val.getFullYear();
          const m = String(val.getMonth() + 1).padStart(2, '0');
          const d = String(val.getDate()).padStart(2, '0');
          val = `${y}-${m}-${d}`;
        } else {
          val = val == null ? "" : String(val);
        }
        rowObj[header] = val;
      }
      // Ensure date is in canonical YYYY-MM-DD format
      rowObj[currentTableHeaders[dateLocalIdx]] = dateStr;
      parsedRows.push(rowObj);
    }

    if (parsedRows.length === 0) {
      getNotifier().warning(`⚠️ 沒有可匯入的資料！（${skippedCount} 筆因日期無效被略過）`);
      return;
    }

    fillTableWithData(parsedRows);

    if (skippedCount > 0) {
      getNotifier().warning(`⚠️ 已匯入 ${parsedRows.length} 筆，另有 ${skippedCount} 筆因日期無效被略過。`);
    } else {
      getNotifier().success(`✅ 已成功匯入 ${parsedRows.length} 筆排班資料！`);
    }
  } catch (err) {
    console.error("[importExcel] 解析失敗：", err);
    getNotifier().error("❌ Excel 解析失敗：" + err.message);
  } finally {
    getNotifier().hideLoading();
    input.value = "";
  }
}

// ============================================================
//  📢 講道資訊連動設定與連動邏輯
// ============================================================
function toggleSermonTypeSelect() {
  const useSermonToggle = document.getElementById('useSermonToggle');
  const sermonTypeCol = document.getElementById('sermonTypeCol');
  if (useSermonToggle && sermonTypeCol) {
    sermonTypeCol.style.opacity = useSermonToggle.checked ? "1" : "0.5";
    sermonTypeCol.style.pointerEvents = useSermonToggle.checked ? "auto" : "none";
  }
}

async function saveSermonSettings() {
  if (getUIState().isLocked('saveSermonSettings')) return;
  getUIState().lock('saveSermonSettings');

  const useSermon = document.getElementById('useSermonToggle').checked;
  const sermonType = document.getElementById('sermonTypeSelect').value;

  getNotifier().showLoading("💾 儲存設定中...");
  try {
    await fetchAPI("saveSermonSettings", {
      id: currentId,
      sermonSettings: { useSermon, sermonType }
    });

    currentSermonSettings = { useSermon, sermonType };
    getNotifier().success("✅ 講道資訊連動設定已更新！各列的「套用講道」勾選請透過新增一季聚會或匯入功能設定。");
  } catch (err) {
    handleAPIError(err);
  } finally {
    getNotifier().hideLoading();
    getUIState().unlock('saveSermonSettings');
  }
}

function onSermonLinkChange(checkbox) {
  const rowDiv = checkbox.closest('.record-row');
  const dateColIdx = currentTableHeaders.findIndex(h => h.includes("日期"));
  if (dateColIdx === -1) return;
  const dateInput = rowDiv.querySelector(`input[data-c="${dateColIdx}"]`);
  const dateVal = dateInput ? dateInput.value.trim() : "";

  updateRowSermonState(rowDiv, checkbox.checked, dateVal);
}

function updateRowSermonState(rowDiv, linkType, dateStr) {
  // 同步更新「套用講道」勾選框與語言滑動開關狀態，確保畫面顯示與連動狀態一致
  const sermonLinkColIdx = currentTableHeaders.indexOf("套用講道");
  if (sermonLinkColIdx !== -1) {
    const checkbox = rowDiv.querySelector(`input.sermon-link-checkbox[data-c="${sermonLinkColIdx}"]`);
    const sw = rowDiv.querySelector(`input.sermon-lang-switch[data-c="${sermonLinkColIdx}"]`);
    const lblLeft = rowDiv.querySelector(`#lang-label-left-${sermonLinkColIdx}`);
    const lblRight = rowDiv.querySelector(`#lang-label-right-${sermonLinkColIdx}`);
    
    const isLinked = linkType !== "N" && linkType !== "";
    
    if (checkbox && checkbox.checked !== isLinked) {
      checkbox.checked = isLinked;
    }
    
    if (sw) {
      let langVal = linkType === "Y" ? currentSermonSettings.sermonType : linkType;
      if (langVal !== "華語/聯合" && langVal !== "台語/聯合") {
        let isSwChecked = sw.checked;
        langVal = isSwChecked ? "台語/聯合" : "華語/聯合";
      }
      
      const isTaiwanese = langVal === "台語/聯合";
      if (sw.checked !== isTaiwanese) {
        sw.checked = isTaiwanese;
      }
      sw.disabled = !isLinked;
      
      // 更新標籤透明度
      if (lblLeft && lblRight) {
        if (isLinked) {
          lblLeft.style.opacity = isTaiwanese ? "0.4" : "1";
          lblRight.style.opacity = isTaiwanese ? "1" : "0.4";
        } else {
          lblLeft.style.opacity = "0.4";
          lblRight.style.opacity = "0.4";
        }
      }
    }
  }

  // 找出這一列中所有 grid-input
  const inputsMap = Array.from(rowDiv.querySelectorAll('.grid-input')).reduce((map, input) => {
    const c = input.dataset.c;
    if (c !== undefined) map[c] = input;
    return map;
  }, {});

  // 只連動「主題」和「經文」，話語分享不受講道連動影響
  const fields = ["主題", "經文"];
  const fieldIndices = {};
  fields.forEach(f => {
    fieldIndices[f] = currentTableHeaders.indexOf(f);
  });

  console.log(`[SermonLink] updateRowSermonState: linkType="${linkType}", dateStr="${dateStr}", sermonType="${currentSermonSettings ? currentSermonSettings.sermonType : 'undefined'}"`);

  const isLinked = linkType !== "N" && linkType !== "";

  if (isLinked) {
    // 設為唯讀
    fields.forEach(f => {
      const idx = fieldIndices[f];
      if (idx !== -1 && inputsMap[idx]) {
        inputsMap[idx].readOnly = true;
      }
    });

    // 尋找對應講道資訊並套用
    if (dateStr) {
      const activeSermonType = (linkType === "Y") ? currentSermonSettings.sermonType : linkType;
      const sermon = findSermonForDate(dateStr, activeSermonType);
      console.log(`[SermonLink] findSermonForDate returned:`, sermon);
      if (sermon) {
        fields.forEach(f => {
          const idx = fieldIndices[f];
          if (idx !== -1 && inputsMap[idx]) {
            let val = "";
            if (f === "主題") val = sermon.title || "";
            else if (f === "經文") val = sermon.scripture || "";
            inputsMap[idx].value = val;
            inputsMap[idx].title = val;
            inputsMap[idx].classList.add('highlight');
            setTimeout(() => inputsMap[idx].classList.remove('highlight'), 1000);
          }
        });
        return;
      }
    }
    // 未設定日期或查無講道，則清空欄位值
    fields.forEach(f => {
      const idx = fieldIndices[f];
      if (idx !== -1 && inputsMap[idx]) {
        inputsMap[idx].value = "";
        inputsMap[idx].title = "";
      }
    });
  } else {
    // 取消或啟用為不可連動狀態：設為可編輯
    fields.forEach(f => {
      const idx = fieldIndices[f];
      if (idx !== -1 && inputsMap[idx]) {
        inputsMap[idx].readOnly = false;
      }
    });
  }
}

function findSermonForDate(dateStr, sermonType) {
  console.log(`[SermonLink] findSermonForDate details - dateStr: "${dateStr}", sermonType: "${sermonType}". currentEventData size: ${currentEventData ? currentEventData.length : 0}`);
  if (!dateStr || !currentEventData || currentEventData.length === 0) return null;

  // 1. 往前推到可作為小組分享主題的講道週日。
  //    若聚會日期本身是週日，使用再前一個週日，避免同一天上午講道被下午小組直接套用。
  const canonicalDate = parseGregorianDate(dateStr);
  if (!canonicalDate) {
    console.warn(`[SermonLink] parseGregorianDate returned null for dateStr: "${dateStr}"`);
    return null;
  }

  const parts = canonicalDate.split("-");
  const year = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10) - 1; // Date.UTC month is 0-11
  const day = parseInt(parts[2], 10);

  const dateObj = new Date(Date.UTC(year, month, day));
  if (isNaN(dateObj.getTime())) {
    console.error(`[SermonLink] Invalid Date constructed from:`, parts);
    return null;
  }

  // getUTCDay() 回傳 0-6 (星期天為 0)。
  // 週日聚會要取前一週；其他日子取當週往前最近的週日。
  const daysSinceSunday = dateObj.getUTCDay();
  const offsetDays = daysSinceSunday === 0 ? 7 : daysSinceSunday;
  dateObj.setUTCDate(dateObj.getUTCDate() - offsetDays);
  
  const y = dateObj.getUTCFullYear();
  const m = String(dateObj.getUTCMonth() + 1).padStart(2, '0');
  const d = String(dateObj.getUTCDate()).padStart(2, '0');
  const sundayStr = `${y}-${m}-${d}`;
  console.log(`[SermonLink] Computed Sunday: "${sundayStr}"`);

  // 2. 篩選出該週日的所有行事曆活動
  const sundayEvents = currentEventData.filter(ev => ev.date === sundayStr);
  console.log(`[SermonLink] Filtered Sunday events for "${sundayStr}":`, sundayEvents);
  if (sundayEvents.length === 0) return null;

  // 3. 收集所有活動中的講道資訊
  const sermons = [];
  sundayEvents.forEach(ev => {
    if (ev.sermons && ev.sermons.length > 0) {
      sermons.push(...ev.sermons);
    }
  });
  console.log(`[SermonLink] Sermons list for "${sundayStr}":`, sermons);

  if (sermons.length === 0) return null;

  // 4. 根據 sermonType 連動對應類別的講道 (例如 "華語/聯合" 拆成 "華語" 與 "聯合" 依序尋找)
  const targetTypes = (sermonType || "").split('/');
  let match = null;
  for (const t of targetTypes) {
    const trimmed = t.trim();
    if (!trimmed) continue;
    match = sermons.find(s => s.type && s.type.indexOf(trimmed) !== -1);
    if (match) break;
  }
  
  // 若都沒有，回傳 null（即空白），不套用第一筆講道
  if (!match) return null;
  console.log(`[SermonLink] Selected sermon match:`, match);

  return match;
}

async function forceSyncSermonData() {
  if (getUIState().isLocked('forceSyncSermonData')) return;
  getUIState().lock('forceSyncSermonData');

  getNotifier().showLoading("🔄 正在重新同步外部講道行事曆，請稍候...");
  try {
    const res = await fetchAPI("forceRefreshEvents", {});
    const count = (res && res.count !== undefined) ? res.count : 0;
    
    // 重新載入當前頁面的 PageConfig 以更新前端的 currentEventData
    const pageData = await fetchAPI('getPageConfig', { id: currentId });
    currentEventData = (pageData.eventData || []).map(ev => {
      if (ev.date) {
        const normalized = parseGregorianDate(String(ev.date));
        if (normalized) {
          ev.date = normalized;
        }
      }
      return ev;
    });
    
    // 重新整理目前畫面上所有勾選了「套用講道」的列
    const dateColIdx = currentTableHeaders.findIndex(h => h.includes("日期"));
    const sermonLinkColIdx = currentTableHeaders.indexOf("套用講道");
    
    document.querySelectorAll('.record-row').forEach(rowDiv => {
      if (sermonLinkColIdx !== -1) {
        const checkbox = rowDiv.querySelector(`input.sermon-link-checkbox[data-c="${sermonLinkColIdx}"]`);
        if (checkbox && checkbox.checked && dateColIdx !== -1) {
          const dVal = rowDiv.querySelector(`input.grid-input[data-c="${dateColIdx}"]`).value.trim();
          const sw = rowDiv.querySelector(`input.sermon-lang-switch[data-c="${sermonLinkColIdx}"]`);
          const langVal = sw && sw.checked ? "台語/聯合" : "華語/聯合";
          updateRowSermonState(rowDiv, langVal, dVal);
        }
      }
    });

    getNotifier().success(`✅ 同步完成！已更新 ${count} 筆講道日期資料。`);
  } catch (err) {
    handleAPIError(err);
  } finally {
    getNotifier().hideLoading();
    getUIState().unlock('forceSyncSermonData');
  }
}

function onSermonCheckboxChange(checkbox) {
  const rowDiv = checkbox.closest('.record-row');
  const dateColIdx = currentTableHeaders.findIndex(h => h.includes("日期"));
  if (dateColIdx === -1) return;
  const dateInput = rowDiv.querySelector(`input[data-c="${dateColIdx}"]`);
  const dateVal = dateInput ? dateInput.value.trim() : "";

  const sw = rowDiv.querySelector(`input.sermon-lang-switch[data-c="${checkbox.dataset.c}"]`);
  const lblLeft = rowDiv.querySelector(`#lang-label-left-${checkbox.dataset.c}`);
  const lblRight = rowDiv.querySelector(`#lang-label-right-${checkbox.dataset.c}`);

  if (sw) {
    sw.disabled = !checkbox.checked;
  }

  const isTaiwanese = sw ? sw.checked : false;
  if (lblLeft && lblRight) {
    if (checkbox.checked) {
      lblLeft.style.opacity = isTaiwanese ? "0.4" : "1";
      lblRight.style.opacity = isTaiwanese ? "1" : "0.4";
    } else {
      lblLeft.style.opacity = "0.4";
      lblRight.style.opacity = "0.4";
    }
  }

  const langVal = isTaiwanese ? "台語/聯合" : "華語/聯合";
  const linkType = checkbox.checked ? langVal : "N";
  updateRowSermonState(rowDiv, linkType, dateVal);
}

function onSermonSwitchChange(sw) {
  const rowDiv = sw.closest('.record-row');
  const dateColIdx = currentTableHeaders.findIndex(h => h.includes("日期"));
  if (dateColIdx === -1) return;
  const dateInput = rowDiv.querySelector(`input[data-c="${dateColIdx}"]`);
  const dateVal = dateInput ? dateInput.value.trim() : "";

  const lblLeft = rowDiv.querySelector(`#lang-label-left-${sw.dataset.c}`);
  const lblRight = rowDiv.querySelector(`#lang-label-right-${sw.dataset.c}`);

  const isTaiwanese = sw.checked;
  if (lblLeft && lblRight) {
    lblLeft.style.opacity = isTaiwanese ? "0.4" : "1";
    lblRight.style.opacity = isTaiwanese ? "1" : "0.4";
  }

  const langVal = isTaiwanese ? "台語/聯合" : "華語/聯合";
  updateRowSermonState(rowDiv, langVal, dateVal);
}

// 註冊至全域 window，確保 inline HTML 呼叫無誤
window.toggleSermonTypeSelect = toggleSermonTypeSelect;
window.saveSermonSettings = saveSermonSettings;
window.onSermonLinkChange = onSermonLinkChange;
window.onSermonCheckboxChange = onSermonCheckboxChange;
window.onSermonSwitchChange = onSermonSwitchChange;
window.forceSyncSermonData = forceSyncSermonData;

function togglePrimaryActions() {
  const container = document.querySelector('.ministry-primary-actions');
  const btn = document.getElementById('toggleActionsBtn');
  const arrow = document.getElementById('toggleActionsArrow');
  if (!container || !btn) return;
  
  const isCollapsed = container.classList.toggle('collapsed');
  
  localStorage.setItem('ministry.primaryActionsCollapsed', isCollapsed ? 'true' : 'false');
  
  if (isCollapsed) {
    if (arrow) arrow.innerText = '▾';
    btn.classList.remove('active');
  } else {
    if (arrow) arrow.innerText = '▴';
    btn.classList.add('active');
  }
}
window.togglePrimaryActions = togglePrimaryActions;

function parseToSlashDate(rawStr) {
  if (!rawStr) return null;
  const hypenDate = parseGregorianDate(String(rawStr).trim());
  if (!hypenDate) return null;
  return hypenDate.replace(/-/g, "/");
}
window.parseToSlashDate = parseToSlashDate;

