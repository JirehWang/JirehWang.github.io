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

// 預覽布告欄 modal 目前的篩選後矩陣（給下載 Excel 用）
let _currentBulletinFiltered = null;


// ============================================================
//  🛡️ API 呼叫核心
//  config.js 已載入並提供 window.GAS_URL / window.AUTH_TOKEN，
//  此處的 fetchAPI 與中央 churchAPI 並存，差異在於：
//    - churchAPI 用 text/plain；fetchAPI 用 form-urlencoded（避開 preflight）
//    - fetchAPI 帶有 timeout/retry 邏輯
//  payload 結構：{ action, token, data: { ...params } }
// ============================================================

const _API_TIMEOUT_MS = 120000;

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
      if (!result || result.status !== 'success') {
        throw new APIError(
          (result && result.message) || '伺服器錯誤',
          null,
          classifyError(result && result.message)
        );
      }
      return result.data;
    } catch (err) {
      if (err instanceof APIError) throw err;
      throw new APIError(err.message || '網路錯誤', null, classifyError(err.message));
    }
  }

  // ── 回退路徑：直接打 GAS（保留 retry / timeout 邏輯） ──
  // 手動加 ministry_ 前綴
  const realAction = action.indexOf('ministry_') === 0 ? action : 'ministry_' + action;
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
  const groupRoleBtn = document.getElementById('manageGroupRolesBtn');
  const isGroupOrFellowship = (currentTemplate === "小組聚會表模板" || currentTemplate === "團契聚會表模板");

  if (memberBtn && !isGroupOrFellowship) {
    memberBtn.classList.remove('hidden');
    currentGroupMembers = localCustomMembers.map(m => m.name);

    if (currentTemplate === "新家人服事表模板") {
      currentCoreMembers = localCustomMembers.filter(m => m.role === "小家長").map(m => m.name);
      let parentNames = currentCoreMembers.join(", ");
      let normalNames = localCustomMembers.filter(m => m.role === "一般同工").map(m => m.name).join(", ");
      currentAutoRoleRules = `【系統強制權限】：\n小家長 (${parentNames})：可排所有服事。\n一般同工 (${normalNames})：不可排特定帶領服事。`;
    }
  } else if (memberBtn) {
    memberBtn.classList.add('hidden');
  }

  // 小組/團契模板才顯示「設定組員身分」按鈕（直接編輯 master）
  if (groupRoleBtn) {
    groupRoleBtn.classList.toggle('hidden', !isGroupOrFellowship);
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

  if (dateColIdx !== -1 && currentTemplate !== "小組聚會表模板" && currentTemplate !== "團契聚會表模板" && currentEventData.length > 0) {
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

  // 預先快取目前所有 row 與其 inputs，避免每筆 parsedRow 都重新 querySelectorAll → O(N²)
  const rowCache = Array.from(container.querySelectorAll('.record-row')).map(rowDiv => ({
    rowDiv,
    inputs: rowDiv.querySelectorAll('.grid-input')
  }));

  parsedRows.forEach(rowData => {
    let target = null;
    const aiDate = rowData["日期"] || rowData[currentTableHeaders[dateColIdx]];

    // 先嘗試比對日期
    if (aiDate && dateColIdx !== -1) {
      target = rowCache.find(r => {
        const di = r.inputs[dateColIdx];
        return di && di.value.trim() === aiDate;
      });
    }

    // 找不到日期相符的 → 找完全空白的列
    if (!target) {
      target = rowCache.find(r => {
        for (let i = 0; i < r.inputs.length; i++) {
          if (r.inputs[i].value.trim() !== "") return false;
        }
        return true;
      });
    }

    // 都沒有 → 新增一列並加入 cache
    if (!target) {
      addNewRow();
      const rowDiv = container.lastElementChild;
      target = { rowDiv, inputs: rowDiv.querySelectorAll('.grid-input') };
      rowCache.push(target);
    }

    currentTableHeaders.forEach((header, colIdx) => {
      const val = rowData[header];
      if (val && val !== "") {
        const input = target.inputs[colIdx];
        input.value = val;
        input.classList.add('highlight');
        setTimeout(() => input.classList.remove('highlight'), 2000);
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
//  📅 共用日期篩選器（modal 用：年度 + 季度 + 滾動近 3 個月）
// ============================================================
const _MS_FILTER_QUARTERS = [1, 2, 3, 4];

// 應該被視為「小組類聚會」的模板（小組總表 / 各小組布告欄會包含這些）
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

/**
 * 在指定容器內渲染「年度+季度」篩選器 + 表格。
 * - 開啟時預設用 rolling（近 3 個月）
 * - 改選 年/季度 → 切到 quarter 模式
 * - 點「近 3 個月」按鈕 → 切回 rolling
 * onFilteredChange(filteredMatrix) 在每次重渲染時觸發，外部用來同步下載按鈕。
 */
function _ms_renderFilterableTable({ container, fullMatrix, tableMinWidth, onFilteredChange }) {
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
    <div class="d-flex align-items-center gap-2 mb-2 flex-wrap p-2 bg-light rounded border">
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
    document.getElementById('ms-filter-table').innerHTML = _ms_buildTableHtml(filtered, { minWidth: tableMinWidth });
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
//  📋 預覽布告欄
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
    onFilteredChange: filtered => { _currentBulletinFiltered = filtered; }
  });

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
    const roleColors = {
      '核心同工': 'border-start border-primary border-4',
      '一般同工': 'border-start border-info border-4',
      '陪伴同工': 'border-start border-secondary border-4',
      '小羊':     'border-start border-light border-4'
    };
    const bClass = roleColors[m.role] || roleColors['小羊'];
    const nickname = (m.nickname || '').trim();
    return `
      <li class="list-group-item d-flex justify-content-between align-items-center ${bClass}">
        <div>
          <span class="fw-bold">${m.name}</span>
          ${nickname ? `<small class="text-muted ms-2">(${nickname})</small>` : ''}
        </div>
        <select class="form-select form-select-sm" style="width: 150px;"
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
