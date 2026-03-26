const GAS_URL = "https://script.google.com/macros/s/AKfycbx4268IkgwQm2Es0gjDHLU_U9nKJrRMR1-xzbbtuaq08lePLgAQ2wnDRrCeHdy9jNhh/exec"; 
const currentId = new URLSearchParams(window.location.search).get('id');
let activeGroupName = "";
let currentTableHeaders = [];

let currentGroupMembers = []; 
let currentCoreMembers = [];  
let currentGroupPrompt = "";  
let currentAutoRoleRules = ""; 

// 🌟 新增：暫存從後端抓回來的原始資料，供唯讀網頁與編輯模式切換使用
let currentGroupData = null; 

const showLoading = (msg) => { const el = document.getElementById('globalLoading'); el.innerText = msg; el.classList.remove('hidden'); };
const hideLoading = () => document.getElementById('globalLoading').classList.add('hidden');

async function fetchAPI(action, params = {}) {
  let url = new URL(GAS_URL);
  url.searchParams.append('action', action);
  for (let key in params) url.searchParams.append(key, params[key]);
  const response = await fetch(url);
  const result = await response.json();
  if (result.status !== 'success') throw new Error(result.message);
  return result.data;
}

// 🌟 核心改變：有 ID 的網址，一律先顯示「唯讀分享網頁」
window.onload = async () => {
  if (!currentId) { 
    showSection('adminMain'); 
    await loadAdminData(); 
  } else { 
    showSection('publicViewSection'); // 先顯示唯讀區塊
    try { 
      currentGroupData = await fetchAPI('getPageConfig', { id: currentId }); 
      renderPublicView(currentGroupData); // 畫出唯讀表格
    } catch(e){
      alert(e.message);
    } 
  }
};

// ==========================================
// 🌟 系統大廳：加入「複製連結」功能
// ==========================================
async function loadAdminData() {
  try {
    showLoading("⏳ 整理儀表板中...");
    const [groups, templates] = await Promise.all([fetchAPI('getGroups'), fetchAPI('getTemplates')]);
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
        const shareUrl = `${base}?id=${g.id}`; // 產生該小組的專屬網址
        
        return `
          <div class="col-12 col-md-4 group-item" data-search="${g.name} ${g.template}">
            <div class="card group-card h-100 shadow-sm d-flex flex-column" style="opacity: ${isEnabled ? '1' : '0.5'}; border-left: 5px solid ${isEnabled ? '#0d6efd' : '#ced4da'};">
              <div class="card-body p-3 flex-grow-1">
                <div class="d-flex align-items-center justify-content-between mb-2">
                  <h5 class="card-title ${isEnabled ? 'text-dark' : 'text-muted'} m-0" style="${isEnabled ? '' : 'text-decoration: line-through;'}">${g.name}</h5>
                  <div class="form-check form-switch m-0 ms-3">
                    <input class="form-check-input" type="checkbox" role="switch" ${isEnabled ? 'checked' : ''} onchange="toggleStatus('${g.id}', '${g.status}')">
                  </div>
                </div>
                ${isEnabled ? `
                  <div class="mt-3 d-flex gap-2">
                    <a href="${shareUrl}" target="_blank" class="btn btn-sm btn-outline-primary w-50 fw-bold">👀 查看網頁</a>
                    <button class="btn btn-sm btn-success w-50 fw-bold" onclick="copyShareLink('${shareUrl}')">🔗 複製連結</button>
                  </div>
                ` : '<p class="text-muted small mt-2 m-0">已停用</p>'}
              </div>
            </div>
          </div>
        `;
      }).join('');
    }
    div.innerHTML = html || '<p class="text-center text-muted">目前尚無資料</p>';
    document.getElementById('templateSelect').innerHTML = '<option value="" disabled selected>選擇模板</option>' + templates.map(t=>`<option value="${t}">${t}</option>`).join('');
  } catch (err) { div.innerHTML = '<p class="text-danger">載入失敗</p>'; } finally { hideLoading(); }
}

function copyShareLink(url) {
  navigator.clipboard.writeText(url).then(() => {
    alert("✅ 專屬網址已複製！您可以直接貼到 LINE 群組給小組員看了。");
  }).catch(err => {
    alert("複製失敗，請手動複製此網址：\n" + url);
  });
}

// ==========================================
// 🌟 全新：唯讀分享網頁邏輯
// ==========================================
function renderPublicView(data) {
  activeGroupName = data.groupName;
  document.getElementById('publicViewTitle').innerText = `📋 ${data.groupName} - 排班布告欄`;
  
  const matrix = data.matrix;
  if (!matrix || matrix.length === 0) return;

  let tableHtml = '<table class="table table-bordered table-hover text-center align-middle mb-0" style="min-width: 800px;"><thead><tr class="table-light">';
  matrix[0].forEach(h => {
    if (h.trim() !== "") tableHtml += `<th>${h}</th>`;
  });
  tableHtml += '</tr></thead><tbody>';

  const rows = matrix.slice(1);
  const validRows = rows.filter(r => r.some(cell => cell.toString().trim() !== ""));
  
  if (validRows.length === 0) {
    tableHtml += `<tr><td colspan="${matrix[0].filter(h => h.trim() !== "").length}" class="text-muted py-4">目前尚無排班資料</td></tr>`;
  } else {
    validRows.forEach(row => {
      tableHtml += '<tr>';
      for(let i = 0; i < matrix[0].filter(h => h.trim() !== "").length; i++) {
        tableHtml += `<td>${row[i] || "-"}</td>`;
      }
      tableHtml += '</tr>';
    });
  }
  tableHtml += '</tbody></table>';

  document.getElementById('publicViewContent').innerHTML = tableHtml;
}

function enterEditMode() {
  const userInput = prompt(`🔒 編輯需要權限\n請輸入本小組的專屬 ID 進行解鎖：\n(提示：網址最後面的英文數字)`);
  if (userInput === null) return; 
  if (userInput.trim() === currentId) {
    showSection('reportSection'); // 切換到編輯畫面
    renderTable(currentGroupData); // 畫出可編輯的動態格子
  } else {
    alert("❌ ID 輸入錯誤，無法進入編輯模式！");
  }
}

function downloadPublicExcel() {
  if (typeof XLSX === 'undefined') return alert("載入套件中，請稍後再試...");
  if (!currentGroupData || !currentGroupData.matrix) return;
  
  const validMatrix = [currentGroupData.matrix[0].filter(h => h.trim() !== "")];
  const rows = currentGroupData.matrix.slice(1).filter(r => r.some(cell => cell.toString().trim() !== ""));
  
  if (rows.length === 0) return alert("目前沒有資料可以下載！");
  
  rows.forEach(row => {
    validMatrix.push(row.slice(0, validMatrix[0].length).map(cell => cell || ""));
  });

  const ws = XLSX.utils.aoa_to_sheet(validMatrix);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "布告欄");
  const today = new Date().toISOString().slice(0, 10).replace(/-/g, "");
  XLSX.writeFile(wb, `${activeGroupName}_排班表_${today}.xlsx`);
}

// ==========================================
// 🌟 原本的編輯模式與AI邏輯 (完全保留您的版本)
// ==========================================
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

function renderTable(data) {
  activeGroupName = data.groupName;
  document.getElementById('groupTitle').innerText = data.groupName;
  
  currentGroupMembers = data.members || [];
  currentCoreMembers = data.coreMembers || []; 
  currentGroupPrompt = data.groupPrompt || "";
  currentAutoRoleRules = data.autoRoleRules || ""; 

  const promptInput = document.getElementById('groupPromptInput');
  if (promptInput) promptInput.value = currentGroupPrompt;
  
  let rawHeaders = data.matrix[0].map(h => h.toString().trim());
  let validColCount = rawHeaders.length;
  while (validColCount > 0 && rawHeaders[validColCount - 1] === "") {
    validColCount--;
  }
  currentTableHeaders = rawHeaders.slice(0, validColCount);

  let datalistHTML = "";
  if (currentGroupMembers.length > 0) {
    datalistHTML += `<datalist id="allMembersList">` + currentGroupMembers.map(m => `<option value="${m}">`).join('') + `</datalist>`;
  }
  if (currentCoreMembers.length > 0) {
    datalistHTML += `<datalist id="coreMembersList">` + currentCoreMembers.map(m => `<option value="${m}">`).join('') + `</datalist>`;
  }
  
  const gridTemplate = `repeat(${validColCount}, minmax(130px, 1fr)) 40px`;
  
  let html = datalistHTML;
  html += `<div class="record-grid-header fw-bold text-muted mb-2 px-1" style="display: grid; grid-template-columns: ${gridTemplate}; gap: 10px;">`;
  currentTableHeaders.forEach(h => html += `<div>${h}</div>`);
  html += `<div class="text-center">操作</div></div>`;
  html += `<div id="rowsContainer" class="d-flex flex-column gap-2">`;
  
  const rows = data.matrix.slice(1);
  const validRows = rows.filter(r => r.some(cell => cell.toString().trim() !== ""));
  if (validRows.length === 0) {
    validRows.push([], [], []);
  }

  validRows.forEach((rowData) => {
    html += createRowHTML(rowData, gridTemplate);
  });
  
  html += `</div>`;
  html += `<button class="btn btn-outline-primary w-100 mt-3 border border-2 border-primary border-opacity-50" style="border-style: dashed !important;" onclick="addNewRow()">➕ 新增一筆空白列</button>`;
  document.getElementById('dynamicFormContainer').innerHTML = html;
  initGridInteraction();
}

function createRowHTML(rowData, gridTemplate) {
  const allDropdownCols = ["破冰", "敬拜"];
  const coreDropdownCols = ["話語", "分享"]; 

  if (!gridTemplate) gridTemplate = `repeat(${currentTableHeaders.length}, minmax(130px, 1fr)) 40px`;

  let rowHtml = `<div class="record-row align-items-center" style="display: grid; grid-template-columns: ${gridTemplate}; gap: 10px;">`;
  
  currentTableHeaders.forEach((header, cIdx) => {
    const isAllCol = (currentGroupMembers.length > 0) && allDropdownCols.some(c => header.includes(c));
    const isCoreCol = (currentCoreMembers.length > 0) && coreDropdownCols.some(c => header.includes(c));
    
    let listAttr = "";
    let extraClass = "";

    if (isCoreCol) {
      listAttr = `list="coreMembersList"`;
      extraClass = `datalist-input`;
    } 
    else if (isAllCol) {
      listAttr = `list="allMembersList"`;
      extraClass = `datalist-input`;
    }
    
    const val = rowData[cIdx] || "";
    rowHtml += `<input type="text" class="grid-input ${extraClass}" data-c="${cIdx}" value="${val}" title="${val}" ${listAttr}>`;
  });
  
  rowHtml += `<button class="btn btn-sm btn-outline-danger" onclick="deleteRow(this)" title="刪除此列">✖</button>`;
  rowHtml += `</div>`;
  return rowHtml;
}

function addNewRow() {
  const container = document.getElementById('rowsContainer');
  const tempDiv = document.createElement('div');
  tempDiv.innerHTML = createRowHTML([]);
  container.appendChild(tempDiv.firstElementChild);
}

function deleteRow(btnElement) {
  if(confirm("確定要刪除這筆排班資料嗎？")) {
    btnElement.parentElement.remove();
  }
}

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
        if (currentRowIndex + i >= container.children.length) {
          addNewRow();
        }
        
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

  container.addEventListener('keydown', (e) => {
    if (e.key === 'Delete' || e.key === 'Backspace') {
      const selected = document.querySelectorAll('.grid-input.selected');
      if (selected.length > 1) { 
        e.preventDefault();
        selected.forEach(input => { input.value = ""; input.style.backgroundColor = "transparent"; });
      }
    }
  });
}

function filterByDate() {
  const start = document.getElementById('startDate').value;
  const end = document.getElementById('endDate').value;
  const recordRows = document.querySelectorAll('.record-row');
  const dateColIdx = currentTableHeaders.findIndex(h => h.includes("日期")); 

  if (dateColIdx === -1) {
    alert("⚠️ 找不到包含「日期」的欄位，無法進行篩選。");
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
    const toast = document.createElement('div');
    toast.className = 'position-fixed bottom-0 end-0 p-3';
    toast.style.zIndex = '1050';
    toast.innerHTML = `<div class="toast show align-items-center text-white bg-success border-0"><div class="d-flex"><div class="toast-body">✅ 已篩選出 ${visibleCount} 筆資料</div></div></div>`;
    document.body.appendChild(toast);
    setTimeout(() => toast.remove(), 2500);
  }
}

function clearDateFilter() {
  document.getElementById('startDate').value = "";
  document.getElementById('endDate').value = "";
  document.querySelectorAll('.record-row').forEach(rowDiv => rowDiv.classList.remove('hidden'));
}

async function processAI() {
  const rawText = document.getElementById('aiRawText').value.trim();
  if (!rawText) return alert("請貼上文字");
  showLoading("🤖 AI 運算中，請稍候...");
  try {
    const response = await fetch(GAS_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ 
        action: "parseWithAI", 
        data: { 
          text: rawText, 
          headers: currentTableHeaders,
          members: currentGroupMembers,   
          groupPrompt: currentGroupPrompt + "\n" + currentAutoRoleRules 
        } 
      })
    });
    const resJson = await response.json();
    if (resJson.status !== 'success') throw new Error(resJson.message);
    fillTableWithData(resJson.data);
    document.getElementById('aiStatus').innerText = "✅ 解析/排班完成！";
    document.getElementById('aiRawText').value = ""; 
  } catch (err) { 
    alert("解析失敗：" + err.message); 
    document.getElementById('aiStatus').innerText = "❌ 發生錯誤，請重試。";
  } finally { hideLoading(); }
}

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

async function saveData() {
  showLoading("💾 儲存中...");
  try {
    const matrix = [currentTableHeaders];
    
    document.querySelectorAll('.record-row').forEach(rowDiv => {
      const row = Array.from(rowDiv.querySelectorAll('.grid-input')).map(i => i.value);
      if (row.some(v => v.trim() !== "")) matrix.push(row);
    });
    
    while(matrix.length <= 50) matrix.push(Array(currentTableHeaders.length).fill(""));
    
    await fetch(GAS_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: "saveSheetData", data: { groupName: activeGroupName, matrix: matrix } })
    });
    alert("✅ 儲存成功！");
  } catch (e) { alert("儲存失敗"); } finally { hideLoading(); }
}

async function toggleStatus(groupId, currentStatus) {
  showLoading("🔄 更新狀態中...");
  try {
    const response = await fetch(GAS_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: "toggleGroupStatus", data: { id: groupId, status: currentStatus } })
    });
    if ((await response.json()).status === 'success') await loadAdminData();
  } catch (e) { alert("更新失敗"); } finally { hideLoading(); }
}

function showSection(id) {
  document.querySelectorAll('.card-custom').forEach(el => el.classList.add('hidden'));
  document.getElementById(id).classList.remove('hidden');
}

const createForm = document.getElementById('createGroupForm');
if (createForm) {
  createForm.onsubmit = async function(e) {
    e.preventDefault(); showLoading("建立中...");
    try {
      await fetch(GAS_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ action: "createGroup", data: { id: document.getElementById('newId').value, name: document.getElementById('newName').value, template: document.getElementById('templateSelect').value } }) });
      location.reload();
    } catch (e) { alert("失敗"); } finally { hideLoading(); }
  };
}

async function saveGroupPrompt() {
  const newPrompt = document.getElementById('groupPromptInput').value.trim();
  showLoading("💾 儲存規則中...");
  try {
    await fetch(GAS_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: "saveGroupPrompt", data: { id: currentId, prompt: newPrompt } })
    });
    currentGroupPrompt = newPrompt; 
    alert("✅ 專屬規則儲存成功！下次 AI 排班將嚴格遵守此規則。");
    document.getElementById('promptSettings').classList.add('hidden'); 
  } catch (e) { alert("儲存失敗：" + e.message); } finally { hideLoading(); }
}

// 編輯模式內的布告欄預覽(Modal)
function showBulletinBoard() {
  const matrix = [currentTableHeaders];
  
  document.querySelectorAll('.record-row').forEach(rowDiv => {
    if (rowDiv.classList.contains('hidden')) return;
    const row = Array.from(rowDiv.querySelectorAll('.grid-input')).map(i => i.value.trim());
    if (row.some(v => v !== "")) matrix.push(row);
  });

  let tableHtml = '<table class="table table-bordered table-hover text-center align-middle" style="min-width: 800px;"><thead><tr>';
  matrix[0].forEach(h => tableHtml += `<th class="bg-light">${h}</th>`);
  tableHtml += '</tr></thead><tbody>';

  for (let i = 1; i < matrix.length; i++) {
    tableHtml += '<tr>';
    matrix[i].forEach(cell => tableHtml += `<td>${cell || "-"}</td>`); 
    tableHtml += '</tr>';
  }
  tableHtml += '</tbody></table>';

  if (matrix.length === 1) tableHtml = '<p class="text-center text-muted my-4">目前沒有資料可顯示</p>';

  document.getElementById('bulletinContent').innerHTML = `<div class="table-responsive">${tableHtml}</div>`;
  document.getElementById('bulletinModalLabel').innerText = `📋 ${activeGroupName} - 排班布告欄`;

  const bulletinModal = new bootstrap.Modal(document.getElementById('bulletinModal'));
  bulletinModal.show();
}

// 編輯模式內的 Excel 下載 (Modal)
function downloadExcel() {
  const matrix = [currentTableHeaders];
  
  document.querySelectorAll('.record-row').forEach(rowDiv => {
    if (rowDiv.classList.contains('hidden')) return;
    const row = Array.from(rowDiv.querySelectorAll('.grid-input')).map(i => i.value.trim());
    if (row.some(v => v !== "")) matrix.push(row);
  });

  if (matrix.length === 1) return alert("目前沒有資料可以下載！");

  const ws = XLSX.utils.aoa_to_sheet(matrix);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "布告欄");

  const today = new Date().toISOString().slice(0, 10).replace(/-/g, "");
  const filename = `${activeGroupName}_排班表_${today}.xlsx`;

  XLSX.writeFile(wb, filename);
}
