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
let currentEventData = []; // 接收物件陣列

const showLoading = (msg) => { const el = document.getElementById('globalLoading'); el.innerText = msg; el.classList.remove('hidden'); };
const hideLoading = () => document.getElementById('globalLoading').classList.add('hidden');

// 🌟 核心修改：全面改用 config.js 的安全路由
async function fetchAPI(action, params = {}) {
  if (typeof window.churchAPI !== 'function') {
    throw new Error("安全路由尚未載入，請確認 config.js 是否正常運作。");
  }
  const result = await window.churchAPI(action, params);
  if (result.status !== 'success') throw new Error(result.message || "發生未知錯誤");
  return result.data;
}

window.onload = async () => {
  // 🌟 新增：檢查該分頁是否曾經解鎖過 (機制 1: 刷新頁面保持解鎖)
  if (sessionStorage.getItem(`isUnlocked_${currentId}`) === 'true') {
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
      
      // 🌟 修改：如果「未解鎖」，才顯示布告欄 (預覽模式)，已解鎖就直接留在編輯區
      if (!isEditorUnlocked) {
        showBulletinBoard(); 
      }
    } catch(e){ alert(e.message); } 
  }
};

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
    document.getElementById('templateSelect').innerHTML = '<option value="" disabled selected>選擇模板</option>' + templates.map(t=>`<option value="${t}">${t}</option>`).join('');
  } catch (err) { div.innerHTML = '<p class="text-danger">載入失敗</p>'; } finally { hideLoading(); }
}

function copyShareLink(url) {
  navigator.clipboard.writeText(url).then(() => { alert("✅ 專屬網址已複製！"); }).catch(err => { alert("複製失敗，請手動複製此網址：\n" + url); });
}

function filterGroups() {
  const val = document.getElementById('groupSearch').value.toLowerCase();
  document.querySelectorAll('.group-item').forEach(el => { el.style.display = el.dataset.search.toLowerCase().includes(val) ? "" : "none"; });
  document.querySelectorAll('.category-section').forEach(header => {
    let hasVisible = false; let next = header.nextElementSibling;
    while (next && next.classList.contains('group-item')) { if (next.style.display !== "none") hasVisible = true; next = next.nextElementSibling; }
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
  currentEventData = data.eventData || []; // 外部聚會物件

  currentTemplate = data.template || "";
  localCustomMembers = data.customMembers || [];
  
  const memberBtn = document.getElementById('manageMembersBtn');
  if (memberBtn && currentTemplate !== "小組聚會表模板") {
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
  if (currentGroupMembers.length > 0) datalistHTML += `<datalist id="allMembersList">` + currentGroupMembers.map(m => `<option value="${m}">`).join('') + `</datalist>`;
  if (currentCoreMembers.length > 0) datalistHTML += `<datalist id="coreMembersList">` + currentCoreMembers.map(m => `<option value="${m}">`).join('') + `</datalist>`;

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
  currentTableHeaders.forEach(h => html += `<div>${h}</div>`); html += `<div class="text-center">操作</div></div>`;
  html += `<div id="rowsContainer" class="d-flex flex-column gap-2">`;
  
  const rows = data.matrix.slice(1);
  let validRows = rows.filter(r => r.some(cell => cell.toString().trim() !== ""));

  // 🌟 核心修改：一進分頁，自動比對並產出「還沒建立」的外部聚會列
  const dateColIdx = currentTableHeaders.findIndex(h => h.includes("日期"));
  const nameColIdx = currentTableHeaders.findIndex(h => h.includes("聚會名稱"));
  const catColIdx = currentTableHeaders.findIndex(h => h.includes("聚會類別"));

  if (dateColIdx !== -1 && currentTemplate !== "小組聚會表模板" && currentEventData.length > 0) {
    const existingDates = validRows.map(r => r[dateColIdx]);

    currentEventData.forEach(event => {
      // 如果這個日期還不在表格裡，就自動生出一列並填入前三欄
      if (!existingDates.includes(event.date)) {
        let newRow = new Array(validColCount).fill("");
        newRow[dateColIdx] = event.date;
        if (nameColIdx !== -1) newRow[nameColIdx] = event.name;
        if (catColIdx !== -1) newRow[catColIdx] = event.category;
        validRows.push(newRow);
      }
    });

    // 自動依照日期幫表格排好順序
    validRows.sort((a, b) => {
      let dateA = a[dateColIdx] || "9999-99-99";
      let dateB = b[dateColIdx] || "9999-99-99";
      return dateA.localeCompare(dateB);
    });
  }

  if (validRows.length === 0) validRows.push(new Array(validColCount).fill(""));

  validRows.forEach((rowData) => html += createRowHTML(rowData, gridTemplate));
  
  html += `</div>`;
  html += `<button class="btn btn-outline-primary w-100 mt-3 border border-2 border-primary border-opacity-50" style="border-style: dashed !important;" onclick="addNewRow()">➕ 新增一筆空白列</button>`;
  
  document.getElementById('dynamicFormContainer').innerHTML = html;
  initGridInteraction();
}

function createRowHTML(rowData, gridTemplate) {
  if (!gridTemplate) gridTemplate = `repeat(${currentTableHeaders.length}, minmax(130px, 1fr)) 40px`;
  let rowHtml = `<div class="record-row align-items-center" style="display: grid; grid-template-columns: ${gridTemplate}; gap: 10px;">`;
  
  currentTableHeaders.forEach((header, cIdx) => {
    let listAttr = ""; let extraClass = "";
    
    let inputType = "text";
    if (header.includes("日期")) inputType = "date";

    if (currentTemplate !== "小組聚會表模板") {
      if (header.includes("日期") || header.includes("聚會名稱") || header.includes("聚會類別")) {
        listAttr = ""; extraClass = "";
      } else {
        if (currentTemplate === "新家人服事表模板" && header.includes("小家長")) {
          listAttr = `list="parentMembersList"`; extraClass = `datalist-input`;
        } else if (currentTemplate === "新家人服事表模板" && header.includes("新家人同工")) {
          listAttr = `list="normalMembersList"`; extraClass = `datalist-input`;
        } else {
          listAttr = `list="customMembersList"`; extraClass = `datalist-input`;
        }
      }
    }
    else {
      const allDropdownCols = ["破冰", "敬拜", "分享"]; 
      const coreDropdownCols = ["話語", "領會", "主領", "帶領"]; 
      const isAllCol = allDropdownCols.some(c => header.includes(c));
      const isCoreCol = coreDropdownCols.some(c => header.includes(c));

      if (isCoreCol) { listAttr = `list="coreMembersList"`; extraClass = `datalist-input`; } 
      else if (isAllCol) { listAttr = `list="allMembersList"`; extraClass = `datalist-input`; }
    }
    
    const val = rowData[cIdx] || "";
    rowHtml += `<input type="${inputType}" class="grid-input ${extraClass}" data-c="${cIdx}" value="${val}" title="${val}" ${listAttr}>`;
  });
  
  rowHtml += `<button class="btn btn-sm btn-outline-danger" onclick="deleteRow(this)" title="刪除此列">✖</button></div>`;
  return rowHtml;
}

function addNewRow() { const container = document.getElementById('rowsContainer'); const tempDiv = document.createElement('div'); tempDiv.innerHTML = createRowHTML([]); container.appendChild(tempDiv.firstElementChild); }
function deleteRow(btnElement) { if(confirm("確定要刪除這筆排班資料嗎？")) btnElement.parentElement.remove(); }

function initGridInteraction() {
  const container = document.getElementById('rowsContainer');
  container.addEventListener('paste', (e) => {
    const target = e.target; if (!target.classList.contains('grid-input')) return;
    const pasteData = (e.clipboardData || window.clipboardData).getData('text');
    if (pasteData.includes('\t') || pasteData.includes('\n')) {
      e.preventDefault(); 
      const startC = parseInt(target.dataset.c); const currentRowDiv = target.closest('.record-row'); let currentRowIndex = Array.from(container.children).indexOf(currentRowDiv);
      const rows = pasteData.split(/\r?\n/); if (rows[rows.length - 1] === "") rows.pop(); 
      for (let i = 0; i < rows.length; i++) {
        if (currentRowIndex + i >= container.children.length) addNewRow();
        const targetRowDiv = container.children[currentRowIndex + i]; const inputs = targetRowDiv.querySelectorAll('.grid-input'); const cols = rows[i].split('\t');
        for (let j = 0; j < cols.length; j++) {
          const c = startC + j;
          if (c < inputs.length) { inputs[c].value = cols[j]; inputs[c].classList.add('highlight'); setTimeout(() => inputs[c].classList.remove('highlight'), 2000); }
        }
      }
    }
  });
  container.addEventListener('keydown', (e) => {
    if (e.key === 'Delete' || e.key === 'Backspace') {
      const selected = document.querySelectorAll('.grid-input.selected');
      if (selected.length > 1) { e.preventDefault(); selected.forEach(input => { input.value = ""; input.style.backgroundColor = "transparent"; }); }
    }
  });
}

function filterByDate() {
  const start = document.getElementById('startDate').value; const end = document.getElementById('endDate').value;
  const recordRows = document.querySelectorAll('.record-row'); const dateColIdx = currentTableHeaders.findIndex(h => h.includes("日期")); 
  if (dateColIdx === -1) return alert("⚠️ 找不到包含「日期」的欄位。");
  let visibleCount = 0;
  recordRows.forEach(rowDiv => {
    const inputs = rowDiv.querySelectorAll('.grid-input'); if (inputs.length === 0) return;
    const dateVal = inputs[dateColIdx].value.trim(); let show = true;
    if (!start && !end) show = true; else if (!dateVal) show = false; 
    else { if (start && dateVal < start) show = false; if (end && dateVal > end) show = false; }
    if (show) { rowDiv.classList.remove('hidden'); visibleCount++; } else { rowDiv.classList.add('hidden'); }
  });
  if (start || end) {
    const toast = document.createElement('div'); toast.className = 'position-fixed bottom-0 end-0 p-3'; toast.style.zIndex = '1050';
    toast.innerHTML = `<div class="toast show align-items-center text-white bg-success border-0"><div class="d-flex"><div class="toast-body">✅ 已篩選出 ${visibleCount} 筆資料</div></div></div>`;
    document.body.appendChild(toast); setTimeout(() => toast.remove(), 2500);
  }
}

function clearDateFilter() { document.getElementById('startDate').value = ""; document.getElementById('endDate').value = ""; document.querySelectorAll('.record-row').forEach(rowDiv => rowDiv.classList.remove('hidden')); }

async function processAI() {
  const rawText = document.getElementById('aiRawText').value.trim(); if (!rawText) return alert("請貼上文字");
  showLoading("🤖 AI 運算中，請稍候...");
  try {
    const resData = await fetchAPI("parseWithAI", { 
      text: rawText, 
      headers: currentTableHeaders, 
      members: currentGroupMembers, 
      groupPrompt: currentGroupPrompt + "\n" + currentAutoRoleRules 
    });
    fillTableWithData(resData); 
    document.getElementById('aiStatus').innerText = "✅ 解析/排班完成！"; 
    document.getElementById('aiRawText').value = ""; 
  } catch (err) { 
    let errorMsg = err.message;
    if (errorMsg.includes("high demand") || errorMsg.includes("503")) {
      errorMsg = "伺服器忙線中，請稍後再試。";
      alert(errorMsg);
      document.getElementById('aiStatus').innerText = "❌ " + errorMsg;
    } else {
      alert("解析失敗：" + errorMsg); 
      document.getElementById('aiStatus').innerText = "❌ 發生錯誤，請重試。"; 
    }
  } finally { 
    hideLoading(); 
  }
}

function fillTableWithData(parsedRows) {
  const container = document.getElementById('rowsContainer'); const dateColIdx = currentTableHeaders.findIndex(h => h.includes("日期"));
  parsedRows.forEach(rowData => {
    let targetRowDiv = null; const aiDate = rowData["日期"] || rowData[currentTableHeaders[dateColIdx]];
    if (aiDate && dateColIdx !== -1) { const allRowDivs = container.querySelectorAll('.record-row'); for (let rowDiv of allRowDivs) { const dateInput = rowDiv.querySelectorAll('.grid-input')[dateColIdx]; if (dateInput && dateInput.value.trim() === aiDate) { targetRowDiv = rowDiv; break; } } }
    if (!targetRowDiv) { const allRowDivs = container.querySelectorAll('.record-row'); for (let rowDiv of allRowDivs) { const inputs = Array.from(rowDiv.querySelectorAll('.grid-input')); if (inputs.every(input => input.value.trim() === "")) { targetRowDiv = rowDiv; break; } } }
    if (!targetRowDiv) { addNewRow(); targetRowDiv = container.lastElementChild; }
    const inputs = targetRowDiv.querySelectorAll('.grid-input');
    currentTableHeaders.forEach((header, colIdx) => {
      const val = rowData[header]; if (val && val !== "") { inputs[colIdx].value = val; inputs[colIdx].classList.add('highlight'); setTimeout(() => inputs[colIdx].classList.remove('highlight'), 2000); }
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
    
    await fetchAPI("saveSheetData", { groupName: activeGroupName, matrix: matrix });
    
    alert("✅ 儲存成功！");
  } catch (e) { alert("儲存失敗"); } finally { hideLoading(); }
}

async function toggleStatus(groupId, currentStatus) {
  showLoading("🔄 更新狀態中...");
  try {
    await fetchAPI("toggleGroupStatus", { id: groupId, status: currentStatus });
    await loadAdminData();
  } catch (e) { alert("更新失敗"); } finally { hideLoading(); }
}

function showSection(id) { document.querySelectorAll('.card-custom').forEach(el => el.classList.add('hidden')); document.getElementById(id).classList.remove('hidden'); }

const createForm = document.getElementById('createGroupForm');
if (createForm) {
  createForm.onsubmit = async function(e) {
    e.preventDefault(); showLoading("建立中...");
    try { 
      await fetchAPI("createGroup", { 
        id: document.getElementById('newId').value, 
        name: document.getElementById('newName').value, 
        template: document.getElementById('templateSelect').value 
      });
      location.reload(); 
    } catch (e) { alert("失敗"); } finally { hideLoading(); }
  };
}

async function saveGroupPrompt() {
  const newPrompt = document.getElementById('groupPromptInput').value.trim(); showLoading("💾 儲存規則中...");
  try { 
    await fetchAPI("saveGroupPrompt", { id: currentId, prompt: newPrompt });
    currentGroupPrompt = newPrompt; 
    alert("✅ 專屬規則儲存成功！"); 
    document.getElementById('promptSettings').classList.add('hidden'); 
  } catch (e) { alert("儲存失敗：" + e.message); } finally { hideLoading(); }
}

function showBulletinBoard() {
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

function closeModalOrUnlock() {
  if (isEditorUnlocked) {
    bulletinModalInstance.hide();
  } else {
    const pwd = prompt(`🔒 編輯需要權限\n請輸入專屬 ID `);
    if (pwd === null) return; 
    
    if (pwd.trim() === currentId) {
      isEditorUnlocked = true;
      // 🌟 新增：將解鎖狀態寫入分頁短期記憶 (機制 1: 刷新頁面保持解鎖)
      sessionStorage.setItem(`isUnlocked_${currentId}`, 'true');
      bulletinModalInstance.hide();
    } else {
      alert("❌ ID 輸入錯誤！無法進入編輯模式。");
    }
  }
}

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
  XLSX.writeFile(wb, `${activeGroupName}_排班表_${today}.xlsx`);
}

function openMemberModal() {
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
      <button class="btn btn-sm btn-danger" onclick="deleteMember(${idx})">刪除</button>
    </li>
  `).join('');
}

function addMember() {
  const nameInput = document.getElementById('newMemberName');
  const roleSelect = document.getElementById('newMemberRole');
  const rawText = nameInput.value.trim();
  const role = roleSelect.value;
  
  if (!rawText) return alert("請輸入姓名！");
  
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
    alert(`✅ 成功批量新增 ${addedCount} 筆名單！` + (dupCount > 0 ? `\n⚠️ 另有 ${dupCount} 筆已存在被自動略過。` : ""));
  } else if (addedCount === 0 && dupCount > 0) {
    alert("⚠️ 您輸入的名字都已經在名單中囉！");
  }
}

function deleteMember(idx) {
  localCustomMembers.splice(idx, 1);
  renderMemberList();
}

// 🌟 替換儲存同工名單功能 (機制 2: 無縫局部更新版)
async function saveMembersToServer() {
  showLoading("💾 儲存名單中...");
  try {
    await fetchAPI("saveGroupMembers", { id: currentId, members: localCustomMembers });
    alert("✅ 名單儲存成功！");
    
    // 1. 關閉人員管理的 Modal (不再使用 location.reload())
    const memberModalEl = document.getElementById('memberModal');
    if (memberModalEl) {
      const memberModal = bootstrap.Modal.getInstance(memberModalEl);
      if (memberModal) memberModal.hide();
    }

    // 2. 背景重新抓取最新的 Config 以更新下拉選單與畫面，達到「無縫銜接」
    showLoading("🔄 更新下拉選單與畫面中...");
    const freshConfig = await fetchAPI('getPageConfig', { id: currentId });
    renderTable(freshConfig); // 重繪表格，新名字就會出現在下拉選單裡了

  } catch (e) { alert("儲存失敗：" + e.message); } finally { hideLoading(); }
}

async function showAggregatedReport(type) {
  showLoading("📊 彙整資料中，這可能需要幾秒鐘...");
  try {
    const matrix = await fetchAPI('getAggregatedReport', { type: type });
    
    if (!matrix || matrix.length <= 1) {
      alert("目前還沒有建立任何資料，或是資料都是空的喔！");
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
  } catch (e) {
    alert("彙整失敗：" + e.message);
  } finally {
    hideLoading();
  }
}

function downloadAggregatedExcel(matrix, fileName) {
  const ws = XLSX.utils.aoa_to_sheet(matrix);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "彙整總表");

  const today = new Date().toISOString().slice(0, 10).replace(/-/g, "");
  XLSX.writeFile(wb, `${fileName}_${today}.xlsx`);
}
