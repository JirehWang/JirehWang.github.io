const GAS_URL = "https://script.google.com/macros/s/AKfycbx4268IkgwQm2Es0gjDHLU_U9nKJrRMR1-xzbbtuaq08lePLgAQ2wnDRrCeHdy9jNhh/exec"; 
const currentId = new URLSearchParams(window.location.search).get('id');
let activeGroupName = "";
let currentTableHeaders = [];

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

window.onload = async () => {
  if (!currentId) { showSection('adminMain'); await loadAdminData(); } 
  else { showSection('reportSection'); try { const data = await fetchAPI('getPageConfig', { id: currentId }); renderTable(data); } catch(e){alert(e.message);} }
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
        return `
          <div class="col-12 col-md-4 group-item" data-search="${g.name} ${g.template}">
            <div class="card group-card h-100 shadow-sm" style="opacity: ${isEnabled ? '1' : '0.5'}; border-left: 5px solid ${isEnabled ? '#0d6efd' : '#ced4da'};">
              <div class="card-body p-3 d-flex align-items-center justify-content-between">
                
                <a href="${isEnabled ? base + '?id=' + g.id : 'javascript:void(0)'}" 
                   class="group-link" 
                   style="${isEnabled ? '' : 'pointer-events: none; cursor: default;'}">
                  <h5 class="card-title ${isEnabled ? 'text-dark' : 'text-muted'}" 
                      style="${isEnabled ? '' : 'text-decoration: line-through;'}">${g.name}</h5>
                </a>

                <div class="form-check form-switch m-0 ms-3">
                  <input class="form-check-input" type="checkbox" role="switch" 
                         ${isEnabled ? 'checked' : ''} 
                         onchange="toggleStatus('${g.id}', '${g.status}')">
                </div>

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
  
  let rawHeaders = data.matrix[0].map(h => h.toString().trim());
  let validColCount = rawHeaders.length;
  while (validColCount > 0 && rawHeaders[validColCount - 1] === "") {
    validColCount--;
  }
  currentTableHeaders = rawHeaders.slice(0, validColCount);

  const theadHTML = currentTableHeaders.map(h => `<th style="width: 150px; position: relative;">${h}<div class="resizer"></div></th>`).join('');
  document.getElementById('tableHead').innerHTML = `<tr>${theadHTML}</tr>`;
  
  let datalistHTML = "";
  if (data.members && data.members.length > 0) {
    datalistHTML = `<datalist id="groupMembers">` + data.members.map(m => `<option value="${m}">`).join('') + `</datalist>`;
  }
  
  const rows = data.matrix.slice(1);
  let html = datalistHTML; 
  const dropdownCols = ["破冰", "敬拜", "話語", "分享"]; 
  
  for (let i = 0; i < 50; i++) {
    const rowData = rows[i] || [];
    html += `<tr>`;
    
    for (let j = 0; j < currentTableHeaders.length; j++) {
      const header = currentTableHeaders[j] || "";
      const isDropdownCol = (data.members && data.members.length > 0) && dropdownCols.some(c => header.includes(c));
      
      const listAttr = isDropdownCol ? `list="groupMembers"` : "";
      const extraClass = isDropdownCol ? `datalist-input` : "";
      
      html += `<td><input type="text" class="grid-input ${extraClass}" data-r="${i}" data-c="${j}" value="${rowData[j] || ""}" title="${rowData[j] || ""}" ${listAttr}></td>`;
    }
    html += `</tr>`;
  }
  document.getElementById('tableBody').innerHTML = html;
  
  initResizable();
  if(typeof initGridInteraction === 'function') initGridInteraction();
}

function initResizable() {
  const table = document.getElementById('mainTable');
  const cols = table.querySelectorAll('th');
  [].forEach.call(cols, (col) => {
    const resizer = col.querySelector('.resizer');
    if (!resizer) return;
    let x = 0, w = 0;
    const onMouseDown = (e) => {
      x = e.clientX; w = parseInt(window.getComputedStyle(col).width, 10);
      document.addEventListener('mousemove', onMouseMove); document.addEventListener('mouseup', onMouseUp);
      resizer.classList.add('resizing');
    };
    const onMouseMove = (e) => { const dx = e.clientX - x; col.style.width = `${w + dx}px`; };
    const onMouseUp = () => { document.removeEventListener('mousemove', onMouseMove); document.removeEventListener('mouseup', onMouseUp); resizer.classList.remove('resizing'); };
    resizer.addEventListener('mousedown', onMouseDown);
  });
}

function initGridInteraction() {
  const tbody = document.getElementById('tableBody');
  let startCell = null;

  tbody.addEventListener('mousedown', (e) => {
    if (!e.target.classList.contains('grid-input')) return;
    const r = parseInt(e.target.dataset.r);
    const c = parseInt(e.target.dataset.c);

    if (e.shiftKey && startCell) {
      e.preventDefault();
      clearSelection();
      const minR = Math.min(startCell.r, r), maxR = Math.max(startCell.r, r);
      const minC = Math.min(startCell.c, c), maxC = Math.max(startCell.c, c);

      for (let i = minR; i <= maxR; i++) {
        for (let j = minC; j <= maxC; j++) {
          const input = document.querySelector(`.grid-input[data-r="${i}"][data-c="${j}"]`);
          if (input) input.classList.add('selected');
        }
      }
    } else {
      clearSelection();
      startCell = { r, c };
      e.target.classList.add('selected');
    }
  });

  function clearSelection() {
    document.querySelectorAll('.grid-input.selected').forEach(el => el.classList.remove('selected'));
  }

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Delete' || e.key === 'Backspace') {
      const selected = document.querySelectorAll('.grid-input.selected');
      if (selected.length > 1) { 
        e.preventDefault();
        selected.forEach(input => {
          input.value = "";
          input.style.backgroundColor = "transparent";
        });
      }
    }
  });

  tbody.addEventListener('paste', (e) => {
    const target = e.target;
    if (!target.classList.contains('grid-input')) return;

    const pasteData = (e.clipboardData || window.clipboardData).getData('text');
    if (pasteData.includes('\t') || pasteData.includes('\n')) {
      e.preventDefault(); 
      clearSelection();

      const startR = parseInt(target.dataset.r);
      const startC = parseInt(target.dataset.c);
      
      const rows = pasteData.split(/\r?\n/);
      if (rows[rows.length - 1] === "") rows.pop(); 

      for (let i = 0; i < rows.length; i++) {
        const cols = rows[i].split('\t');
        for (let j = 0; j < cols.length; j++) {
          const r = startR + i;
          const c = startC + j;
          const input = document.querySelector(`.grid-input[data-r="${r}"][data-c="${c}"]`);
          if (input) {
            input.value = cols[j];
            input.style.backgroundColor = "#fff3cd"; 
            input.classList.add('selected'); 
          }
        }
      }
    }
  });
}

function filterByDate() {
  const start = document.getElementById('startDate').value;
  const end = document.getElementById('endDate').value;
  const tableRows = document.querySelectorAll('#tableBody tr');
  const dateColIdx = currentTableHeaders.findIndex(h => h.includes("日期")); 

  if (dateColIdx === -1) {
    alert("⚠️ 找不到包含「日期」的欄位，無法進行篩選。");
    return;
  }

  let visibleCount = 0;

  tableRows.forEach(tr => {
    const inputs = tr.querySelectorAll('input');
    if (inputs.length === 0) return;
    
    const dateVal = inputs[dateColIdx].value.trim();

    if (!start && !end) {
      tr.style.display = "";
      return;
    }

    if (!dateVal) {
      tr.style.display = "none";
      return;
    }

    let show = true;
    if (start && dateVal < start) show = false;
    if (end && dateVal > end) show = false;

    tr.style.display = show ? "" : "none";
    if (show) visibleCount++;
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
  document.querySelectorAll('#tableBody tr').forEach(tr => tr.style.display = "");
}

async function processAI() {
  const rawText = document.getElementById('aiRawText').value.trim();
  if (!rawText) return alert("請貼上文字");
  showLoading("🤖 AI 解析中...");
  try {
    const response = await fetch(GAS_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: "parseWithAI", data: { text: rawText, headers: currentTableHeaders } })
    });
    const resJson = await response.json();
    if (resJson.status !== 'success') throw new Error(resJson.message);
    fillTableWithData(resJson.data);
    document.getElementById('aiStatus').innerText = "✅ 解析完成！";
    document.getElementById('aiRawText').value = ""; 
  } catch (err) { alert("解析失敗：" + err.message); } finally { hideLoading(); }
}

function fillTableWithData(parsedRows) {
  const tableRows = document.querySelectorAll('#tableBody tr');
  const dateColIdx = currentTableHeaders.indexOf("日期");
  const dateMap = {};
  tableRows.forEach(tr => { const d = tr.querySelectorAll('input')[dateColIdx]?.value.trim(); if(d) dateMap[d] = tr; });
  parsedRows.forEach(rowData => {
    let targetRow = dateMap[rowData["日期"]];
    if (!targetRow) for (let tr of tableRows) if (tr.querySelectorAll('input')[0].value.trim() === "") { targetRow = tr; break; }
    if (targetRow) {
      const inputs = targetRow.querySelectorAll('input');
      currentTableHeaders.forEach((header, colIdx) => {
        const val = rowData[header];
        if (val && val !== "") { inputs[colIdx].value = val; inputs[colIdx].style.backgroundColor = "#fff3cd"; }
      });
    }
  });
}

async function saveData() {
  showLoading("💾 儲存中...");
  try {
    const matrix = [currentTableHeaders];
    document.querySelectorAll('#tableBody tr').forEach(tr => {
      const row = Array.from(tr.querySelectorAll('input')).map(i => i.value);
      if (row.some(v => v.trim() !== "")) matrix.push(row);
    });
    while(matrix.length <= 50) matrix.push(Array(currentTableHeaders.length).fill(""));
    await fetch(GAS_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: "saveSheetData", data: { groupName: activeGroupName, matrix: matrix } })
    });
    alert("✅ 儲存成功！");
    document.querySelectorAll('.grid-input').forEach(i => i.style.backgroundColor = 'transparent');
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
