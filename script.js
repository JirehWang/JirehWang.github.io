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

// ============================================================
//  🛡️ 安全存取 config.js 提供的工具（防止載入競速問題）
//  config.js 從外部 CDN 載入，可能比 script.js 晚完成，
//  這裡提供備援實作，確保即使 config.js 尚未就緒也不會拋錯。
// ============================================================
function getNotifier() {
  return window.userNotification || {
    success: (msg) => console.log("✅", msg),
    error:   (msg) => { console.error("❌", msg); alert(msg); },
    warning: (msg) => { console.warn("⚠️", msg); alert(msg); },
    info:    (msg) => console.log("ℹ️", msg),
    showLoading: (msg) => {
      const el = document.getElementById('globalLoading');
      if (el) { el.innerText = msg || "⏳ 處理中..."; el.classList.remove('hidden'); }
    },
    hideLoading: () => {
      const el = document.getElementById('globalLoading');
      if (el) el.classList.add('hidden');
    }
  };
}

function getUIState() {
  return window.uiState || {
    _locks: {},
    lock:     function(k) { this._locks[k] = true; },
    unlock:   function(k) { this._locks[k] = false; },
    isLocked: function(k) { return !!this._locks[k]; }
  };
}

function getSessionMgr() {
  return window.sessionManager || {
    setUnlocked: (id) => sessionStorage.setItem(`unlocked_${id}`, 'true'),
    isUnlocked:  (id) => sessionStorage.getItem(`unlocked_${id}`) === 'true',
    clear:       (id) => sessionStorage.removeItem(`unlocked_${id}`)
  };
}

// ============================================================
//  🌐 全域變數保底宣告
//  不管 config.js 是新版或舊版，這些名稱在 script.js 內都保證存在。
//  若 config.js 已提供 window.userNotification 等，則 getNotifier() 會優先使用它。
// ============================================================
const userNotification = {
  success:     (msg, d) => getNotifier().success(msg, d),
  error:       (msg, d) => getNotifier().error(msg, d),
  warning:     (msg, d) => getNotifier().warning(msg, d),
  info:        (msg, d) => getNotifier().info(msg, d),
  showLoading: (msg)    => getNotifier().showLoading(msg),
  hideLoading: ()       => getNotifier().hideLoading()
};

const uiState = {
  lock:     (k) => getUIState().lock(k),
  unlock:   (k) => getUIState().unlock(k),
  isLocked: (k) => getUIState().isLocked(k)
};

const sessionManager = {
  setUnlocked: (id) => getSessionMgr().setUnlocked(id),
  isUnlocked:  (id) => getSessionMgr().isUnlocked(id),
  clear:       (id) => getSessionMgr().clear(id)
};

// APIError fallback（舊版 config.js 未定義時在此補充）
if (typeof APIError === 'undefined') {
  window.APIError = class APIError extends Error {
    constructor(message, httpStatus = null, type = 'UNKNOWN_ERROR') {
      super(message);
      this.name = 'APIError';
      this.httpStatus = httpStatus;
      this.type = type;
    }
  };
}

// ============================================================

const currentId = new URLSearchParams(window.location.search).get('id');
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


// ============================================================
//  🛡️ API 呼叫核心（自給自足，不依賴 config.js 的實作細節）
//
//  設定來源優先順序：
//    1. window.CHURCH_CONFIG（新版 config.js 提供）
//    2. window._GAS_URL / window._GAS_TOKEN（手動注入）
//    3. 內建預設值（需手動修改 FALLBACK_GAS_URL）
//
//  payload 結構統一為：{ action, token, data: { ...params } }
//  後端 doPost 讀取：var data = payload.data || {}
// ============================================================

// ⚙️ 如果沒有新版 config.js，在這裡直接填入你的部署網址與 Token
const _FALLBACK_GAS_URL   = "https://script.google.com/macros/s/AKfycbx4268IkgwQm2Es0gjDHLU_U9nKJrRMR1-xzbbtuaq08lePLgAQ2wnDRrCeHdy9jNhh/exec";
const _FALLBACK_GAS_TOKEN = "ChurchApp-2026";
const _API_TIMEOUT_MS     = 30000;

function _getGasUrl() {
  return (window.CHURCH_CONFIG && window.CHURCH_CONFIG.GAS_DEPLOY_URL) ||
         window._GAS_URL ||
         _FALLBACK_GAS_URL;
}

function _getGasToken() {
  return (window.CHURCH_CONFIG && window.CHURCH_CONFIG.SECRET_TOKEN) ||
         window._GAS_TOKEN ||
         _FALLBACK_GAS_TOKEN;
}

async function fetchAPI(action, data = {}) {
  // data 就是要放進 payload.data 的內容，呼叫端直接傳物件即可
  // 例：fetchAPI('getPageConfig', { id: currentId })
  //     fetchAPI('saveSheetData', { groupName: '...', matrix: [...] })

  const payload = {
    action: action,
    token:  _getGasToken(),
    data:   data
  };

  const gasUrl = _getGasUrl();
  if (!gasUrl || gasUrl.includes("YOUR_DEPLOYMENT_ID")) {
    throw new APIError(
      "GAS 部署網址尚未設定，請修改 config.js 或 script.js 的 _FALLBACK_GAS_URL",
      null, 'CONFIG_ERROR'
    );
  }

  let lastError;
  const maxRetries = 3;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), _API_TIMEOUT_MS);

      // GAS 不支援 CORS preflight（OPTIONS），必須用 no-cors + redirect:follow
      // no-cors 模式下 response 是 opaque，無法直接 .json()
      // 解法：先用 no-cors 觸發 GAS 執行，再用 GET 取得結果
      // 實務上最簡單的方式是改用 application/x-www-form-urlencoded
      // 這樣瀏覽器不會發 preflight，GAS 也能正常接收
      const formBody = "payload=" + encodeURIComponent(JSON.stringify(payload));

      const response = await fetch(gasUrl, {
        method:      'POST',
        headers:     { 'Content-Type': 'application/x-www-form-urlencoded' },
        body:        formBody,
        signal:      controller.signal,
        redirect:    'follow'
      });
      clearTimeout(timer);

      if (!response.ok) {
        throw new APIError(`HTTP ${response.status}`, response.status, 'HTTP_ERROR');
      }

      const json = await response.json();

      if (json.status !== 'success') {
        throw new APIError(json.message || '伺服器錯誤', null,
          classifyError(json.message));
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

  if (!currentId) {
    showSection('adminMain');
    await loadAdminData();
  } else {
    showSection('reportSection');
    try {
      const data = await fetchAPI('getPageConfig', { id: currentId });
      renderTable(data);

      // 如果未解鎖，才顯示布告欄 (預覽模式)
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
        const shareUrl = `${base}?id=${g.id}`;

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
    document.getElementById('templateSelect').innerHTML = '<option value="" disabled selected>選擇模板</option>' + templates.map(t => `<option value="${t}">${t}</option>`).join('');

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
  currentEventData = data.eventData || [];

  currentTemplate = data.template || "";
  localCustomMembers = data.customMembers || [];

  const memberBtn = document.getElementById('manageMembersBtn');
  if (memberBtn && currentTemplate !== "小組聚會表模板" && currentTemplate !== "團契聚會表模板") {
    memberBtn.classList.remove('hidden');
    currentGroupMembers = localCustomMembers.map(m => m.name);

    if (currentTemplate === "新家人服事表模板") {
      currentCoreMembers = localCustomMembers.filter(m => m.role === "小家長").map(m => m.name);
      let parentNames = currentCoreMembers.join(", ");
      let normalNames = localCustomMembers.filter(m => m.role === "一般同工").map(m => m.name).join(", ");
      currentAutoRoleRules = `【系統強制權限】：\n小家長 (${parentNames})：可排所有服事。\n一般同工 (${normalNames})：不可排特定帶領服事。`;
    }
  }

  const promptInput = document.getElementById('groupPromptInput');
  if (promptInput) promptInput.value = currentGroupPrompt;

  let rawHeaders = data.matrix[0].map(h => h.toString().trim());
  let validColCount = rawHeaders.length;
  while (validColCount > 0 && rawHeaders[validColCount - 1] === "") validColCount--;
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

  const gridTemplate = `repeat(${validColCount}, minmax(130px, 1fr)) 40px`;

  let html = datalistHTML;
  html += `<div class="record-grid-header fw-bold text-muted mb-2 px-1" style="display: grid; grid-template-columns: ${gridTemplate}; gap: 10px;">`;
  currentTableHeaders.forEach(h => html += `<div>${h}</div>`);
  html += `<div class="text-center">操作</div></div>`;
  html += `<div id="rowsContainer" class="d-flex flex-column gap-2">`;

  const rows = data.matrix.slice(1);
  let validRows = rows.filter(r => r.some(cell => cell.toString().trim() !== ""));

  const dateColIdx = currentTableHeaders.findIndex(h => h.includes("日期"));
  const nameColIdx = currentTableHeaders.findIndex(h => h.includes("聚會名稱"));
  const catColIdx = currentTableHeaders.findIndex(h => h.includes("聚會類別"));

  if (dateColIdx !== -1 && currentTemplate !== "小組聚會表模板" && currentEventData.length > 0) {
    const existingDates = validRows.map(r => r[dateColIdx]);

    currentEventData.forEach(event => {
      if (!existingDates.includes(event.date)) {
        let newRow = new Array(validColCount).fill("");
        newRow[dateColIdx] = event.date;
        if (nameColIdx !== -1) newRow[nameColIdx] = event.name;
        if (catColIdx !== -1) newRow[catColIdx] = event.category;
        validRows.push(newRow);
      }
    });

    validRows.sort((a, b) => {
      let dateA = a[dateColIdx] || "9999-99-99";
      let dateB = b[dateColIdx] || "9999-99-99";
      return dateA.localeCompare(dateB);
    });
  }

  if (validRows.length === 0) validRows.push(new Array(validColCount).fill(""));

  validRows.forEach((rowData) => html += createRowHTML(rowData, gridTemplate));

  html += `</div>`;
  html += `<button type="button" class="btn btn-outline-primary w-100 mt-3 border border-2 border-primary border-opacity-50" style="border-style: dashed !important;" onclick="addNewRow()">➕ 新增一筆空白列</button>`;

  document.getElementById('dynamicFormContainer').innerHTML = html;
  initGridInteraction();
}


// ============================================================
//  🧩 建立表單列 HTML
// ============================================================
function createRowHTML(rowData, gridTemplate) {
  if (!gridTemplate) gridTemplate = `repeat(${currentTableHeaders.length}, minmax(130px, 1fr)) 40px`;
  let rowHtml = `<div class="record-row align-items-center" style="display: grid; grid-template-columns: ${gridTemplate}; gap: 10px;">`;

  currentTableHeaders.forEach((header, cIdx) => {
    let listAttr = "";
    let extraClass = "";
    let inputType = "text";
    if (header.includes("日期")) inputType = "date";

    if (currentTemplate === "團契聚會表模板") {
      // 團契聚會表模板：破冰敬拜用全體名單，司會用核心名單，其餘手填
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
      // 新家人服事表模板 / 其他自訂模板
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
      // 小組聚會表模板
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

    const val = rowData[cIdx] || "";
    rowHtml += `<input type="${inputType}" class="grid-input ${extraClass}" data-c="${cIdx}" value="${val}" title="${val}" ${listAttr}>`;
  });

  rowHtml += `<button type="button" class="btn btn-sm btn-outline-danger" onclick="deleteRow(this)" title="刪除此列">✖</button></div>`;
  return rowHtml;
}

// ============================================================
//  ➕ 新增列 / 🗑️ 刪除列
// ============================================================
function addNewRow() {
  const container = document.getElementById('rowsContainer');
  const tempDiv = document.createElement('div');
  tempDiv.innerHTML = createRowHTML([]);
  container.appendChild(tempDiv.firstElementChild);
}

function deleteRow(btnElement) {
  if (confirm("確定要刪除這筆排班資料嗎？")) {
    btnElement.parentElement.remove();
  }
}


// ============================================================
//  🎯 網格互動（複製貼上等）
// ============================================================
function initGridInteraction() {
  const container = document.getElementById('rowsContainer');
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
        const inputs = targetRowDiv.querySelectorAll('.grid-input');
        const cols = rows[i].split('\t');
        for (let j = 0; j < cols.length; j++) {
          const c = startC + j;
          if (c < inputs.length) {
            inputs[c].value = cols[j];
            inputs[c].classList.add('highlight');
            setTimeout(() => inputs[c].classList.remove('highlight'), 2000);
          }
        }
      }
    }
  });
}


// ============================================================
//  📅 日期篩選
// ============================================================
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
      if (start && dateVal < start) show = false;
      if (end && dateVal > end) show = false;
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
  }
}


// ============================================================
//  📝 填充表單資料
// ============================================================
function fillTableWithData(parsedRows) {
  const container = document.getElementById('rowsContainer');
  const dateColIdx = currentTableHeaders.findIndex(h => h.includes("日期"));

  parsedRows.forEach(rowData => {
    let targetRowDiv = null;
    const aiDate = rowData["日期"] || rowData[currentTableHeaders[dateColIdx]];

    if (aiDate && dateColIdx !== -1) {
      const allRowDivs = container.querySelectorAll('.record-row');
      for (let rowDiv of allRowDivs) {
        const dateInput = rowDiv.querySelectorAll('.grid-input')[dateColIdx];
        if (dateInput && dateInput.value.trim() === aiDate) {
          targetRowDiv = rowDiv;
          break;
        }
      }
    }

    if (!targetRowDiv) {
      const allRowDivs = container.querySelectorAll('.record-row');
      for (let rowDiv of allRowDivs) {
        const inputs = Array.from(rowDiv.querySelectorAll('.grid-input'));
        if (inputs.every(input => input.value.trim() === "")) {
          targetRowDiv = rowDiv;
          break;
        }
      }
    }

    if (!targetRowDiv) {
      addNewRow();
      targetRowDiv = container.lastElementChild;
    }

    const inputs = targetRowDiv.querySelectorAll('.grid-input');
    currentTableHeaders.forEach((header, colIdx) => {
      const val = rowData[header];
      if (val && val !== "") {
        inputs[colIdx].value = val;
        inputs[colIdx].classList.add('highlight');
        setTimeout(() => inputs[colIdx].classList.remove('highlight'), 2000);
      }
    });
  });
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
      const row = Array.from(rowDiv.querySelectorAll('.grid-input')).map(i => i.value);
      if (row.some(v => v.trim() !== "")) matrix.push(row);
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
      await fetchAPI("createGroup", {
        id: document.getElementById('newId').value,
        name: document.getElementById('newName').value,
        template: document.getElementById('templateSelect').value
      });
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
    document.getElementById('promptSettings').classList.add('hidden');
  } catch (err) {
    handleAPIError(err);
  } finally {
    getNotifier().hideLoading();
    getUIState().unlock('saveGroupPrompt');
  }
}


// ============================================================
//  📋 預覽布告欄
// ============================================================
function showBulletinBoard() {
  if (window.event) window.event.preventDefault();
  const matrix = [currentTableHeaders];
  document.querySelectorAll('.record-row').forEach(rowDiv => {
    if (rowDiv.classList.contains('hidden')) return;
    const row = Array.from(rowDiv.querySelectorAll('.grid-input')).map(i => i.value.trim());
    if (row.some(v => v !== "")) matrix.push(row);
  });

  let tableHtml = '<table class="table table-bordered table-hover text-center align-middle m-0" style="min-width: 800px;"><thead><tr>';
  matrix[0].forEach(h => tableHtml += `<th class="bg-light" style="position: sticky; top: 0; z-index: 10; outline: 1px solid #dee2e6;">${h}</th>`);
  tableHtml += '</tr></thead><tbody>';

  for (let i = 1; i < matrix.length; i++) {
    tableHtml += '<tr>';
    matrix[i].forEach(cell => tableHtml += `<td>${cell || "-"}</td>`);
    tableHtml += '</tr>';
  }
  tableHtml += '</tbody></table>';

  if (matrix.length === 1) tableHtml = '<p class="text-center text-muted my-4">目前沒有資料可顯示</p>';

  document.getElementById('bulletinContent').innerHTML = `<div class="table-responsive" style="max-height: 65vh; overflow-y: auto;">${tableHtml}</div>`;
  document.getElementById('bulletinModalLabel').innerText = `📋 ${activeGroupName} - 排班布告欄`;

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
//  🔓 解鎖編輯模式
// ============================================================
function closeModalOrUnlock() {
  if (window.event) window.event.preventDefault();
  if (isEditorUnlocked) {
    bulletinModalInstance.hide();
  } else {
    const pwd = prompt(`🔒 編輯需要權限\n請輸入專屬 ID`);
    if (pwd === null) return;

    if (pwd.trim() === currentId) {
      isEditorUnlocked = true;
      getSessionMgr().setUnlocked(currentId);
      bulletinModalInstance.hide();
      getNotifier().success("✅ 編輯模式已啟用");
    } else {
      getNotifier().error("❌ ID 輸入錯誤！無法進入編輯模式。");
    }
  }
}


// ============================================================
//  📥 下載 Excel
// ============================================================
function downloadExcel() {
  if (window.event) window.event.preventDefault();
  const matrix = [currentTableHeaders];
  document.querySelectorAll('.record-row').forEach(rowDiv => {
    if (rowDiv.classList.contains('hidden')) return;
    const row = Array.from(rowDiv.querySelectorAll('.grid-input')).map(i => i.value.trim());
    if (row.some(v => v !== "")) matrix.push(row);
  });
  if (matrix.length === 1) {
    getNotifier().warning("⚠️ 目前沒有資料可以下載！");
    return;
  }

  const ws = XLSX.utils.aoa_to_sheet(matrix);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "布告欄");

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
//  📊 彙整報表
// ============================================================
async function showAggregatedReport(type) {
  if (window.event) window.event.preventDefault();

  if (getUIState().isLocked('showAggregatedReport')) return;
  getUIState().lock('showAggregatedReport');

  getNotifier().showLoading("📊 彙整資料中，這可能需要幾秒鐘...");

  try {
    const matrix = await fetchAPI('getAggregatedReport', { type: type });

    if (!matrix || matrix.length <= 1) {
      getNotifier().warning("⚠️ 目前還沒有建立任何資料，或是資料都是空的喔！");
      return;
    }

    let tableHtml = '<table class="table table-bordered table-hover text-center align-middle m-0" style="min-width: 1200px;"><thead><tr>';
    matrix[0].forEach(h => tableHtml += `<th class="bg-light" style="position: sticky; top: 0; z-index: 10; outline: 1px solid #dee2e6;">${h}</th>`);
    tableHtml += '</tr></thead><tbody>';

    for (let i = 1; i < matrix.length; i++) {
      tableHtml += '<tr>';
      matrix[i].forEach(cell => tableHtml += `<td>${cell || "-"}</td>`);
      tableHtml += '</tr>';
    }
    tableHtml += '</tbody></table>';

    document.getElementById('aggregatedReportContent').innerHTML = `<div class="table-responsive" style="max-height: 65vh; overflow-y: auto;">${tableHtml}</div>`;

    const title = type === 'smallGroup' ? '📊 所有小組聚會總表' : '📊 教會各項服事總表';
    document.getElementById('aggregatedReportModalLabel').innerText = title;
    document.getElementById('downloadAggregatedBtn').onclick = () => downloadAggregatedExcel(matrix, title);

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
