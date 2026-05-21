const visibleColumns = [
  '新家人姓名',
  '參加的聚會是',
  '手機',
  '關懷同工',
  '邀約人',
  '日期',
  '落戶狀態'
];

const form = document.getElementById('newFamilyForm');
const formNotice = document.getElementById('formNotice');
const trackingNotice = document.getElementById('trackingNotice');
const submitBtn = document.getElementById('submitBtn');
const refreshBtn = document.getElementById('refreshBtn');
const closeBtn = document.getElementById('closeBtn');
const trackingContent = document.getElementById('trackingContent');
const caseCount = document.getElementById('caseCount');
const dateField = document.getElementById('date');

dateField.valueAsDate = new Date();

document.querySelectorAll('.tab').forEach(button => {
  button.addEventListener('click', () => switchTab(button.dataset.tab));
});

form.addEventListener('submit', async event => {
  event.preventDefault();
  setNotice(formNotice, '送出中...');
  submitBtn.disabled = true;

  try {
    const result = await callApi('submitNewFamily', Object.fromEntries(new FormData(form).entries()));
    setNotice(formNotice, `${result.message}，表單號：${result.formNumber}`, 'success');
    form.reset();
    dateField.valueAsDate = new Date();
  } catch (error) {
    setNotice(formNotice, error.message || String(error), 'error');
  } finally {
    submitBtn.disabled = false;
  }
});

refreshBtn.addEventListener('click', loadTrackingCases);
closeBtn.addEventListener('click', closeSelectedCases);

async function callApi(action, data = {}) {
  const apiUrl = window.NEW_FAMILY_API_URL || '';
  const token = window.NEW_FAMILY_AUTH_TOKEN || '';

  if (!apiUrl) {
    throw new Error('尚未設定 GAS Web App URL，請先填入 api-config.js');
  }

  const response = await fetch(apiUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action, token, data })
  });

  const result = await response.json();
  if (!result.success) {
    throw new Error(result.message || '操作失敗');
  }
  return result;
}

function switchTab(tabName) {
  document.querySelectorAll('.tab').forEach(button => {
    button.classList.toggle('active', button.dataset.tab === tabName);
  });

  document.getElementById('formPanel').hidden = tabName !== 'form';
  document.getElementById('trackingPanel').hidden = tabName !== 'tracking';

  if (tabName === 'tracking') loadTrackingCases();
}

async function loadTrackingCases() {
  setNotice(trackingNotice, '');
  trackingContent.className = 'empty';
  trackingContent.textContent = '載入中...';
  refreshBtn.disabled = true;
  closeBtn.disabled = true;

  try {
    const result = await callApi('getTrackingCases');
    renderTrackingCases(result.data || []);
  } catch (error) {
    trackingContent.className = 'empty';
    trackingContent.textContent = '讀取失敗';
    setNotice(trackingNotice, error.message || String(error), 'error');
  } finally {
    refreshBtn.disabled = false;
    closeBtn.disabled = false;
  }
}

function renderTrackingCases(rows) {
  caseCount.textContent = `共 ${rows.length} 筆`;

  if (!rows.length) {
    trackingContent.className = 'empty';
    trackingContent.textContent = '目前沒有追蹤中的資料';
    return;
  }

  const table = document.createElement('table');
  const thead = document.createElement('thead');
  const tbody = document.createElement('tbody');
  const headRow = document.createElement('tr');

  headRow.innerHTML = `<th class="check-cell">結案</th>${visibleColumns.map(column => `<th>${escapeHtml(column)}</th>`).join('')}`;
  thead.appendChild(headRow);

  rows.forEach(item => {
    const row = document.createElement('tr');
    const checkboxCell = document.createElement('td');
    checkboxCell.className = 'check-cell';
    checkboxCell.innerHTML = `<input type="checkbox" value="${item.rowNumber}" aria-label="勾選 ${escapeHtml(item['新家人姓名'] || '此筆資料')} 結案">`;
    row.appendChild(checkboxCell);

    visibleColumns.forEach(column => {
      const cell = document.createElement('td');
      cell.textContent = item[column] || '';
      row.appendChild(cell);
    });

    tbody.appendChild(row);
  });

  table.appendChild(thead);
  table.appendChild(tbody);

  trackingContent.className = 'table-wrap';
  trackingContent.textContent = '';
  trackingContent.appendChild(table);
}

async function closeSelectedCases() {
  const selected = Array.from(trackingContent.querySelectorAll('input[type="checkbox"]:checked'))
    .map(input => Number(input.value));

  if (!selected.length) {
    setNotice(trackingNotice, '請先勾選要結案的資料', 'error');
    return;
  }

  if (!confirm(`確認將 ${selected.length} 筆資料移到「已結案」嗎？`)) return;

  setNotice(trackingNotice, '結案處理中...');
  closeBtn.disabled = true;

  try {
    const result = await callApi('closeCases', { rowNumbers: selected });
    setNotice(trackingNotice, result.message, 'success');
    await loadTrackingCases();
  } catch (error) {
    setNotice(trackingNotice, error.message || String(error), 'error');
  } finally {
    closeBtn.disabled = false;
  }
}

function setNotice(element, message, type) {
  element.textContent = message;
  element.className = `notice ${type || ''}`.trim();
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
