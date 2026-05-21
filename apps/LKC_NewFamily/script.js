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
const meetingSelect = document.getElementById('meeting');
const settlementStatusSelect = document.getElementById('settlementStatus');
const closedNotice = document.getElementById('closedNotice');
const closedContent = document.getElementById('closedContent');
const closedCount = document.getElementById('closedCount');
const closedSearchBtn = document.getElementById('closedSearchBtn');

dateField.valueAsDate = new Date();
loadMeetingOptions();
loadSettlementStatusOptions();

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
closedSearchBtn.addEventListener('click', loadClosedCases);

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

async function callSundayAttendanceApi(action, data = {}) {
  const apiUrl = window.SUNDAY_ATTENDANCE_API_URL || '';
  const token = window.NEW_FAMILY_AUTH_TOKEN || '';

  if (!apiUrl) {
    throw new Error('尚未設定主日出席 API URL');
  }

  const response = await fetch(apiUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action, token, data })
  });

  const result = await response.json();
  if (result.status === 'error' || result.error) {
    throw new Error(result.message || result.error || '聚會清單讀取失敗');
  }
  return result.data || result;
}

async function callGroupAttendanceApi(action, data = {}) {
  const apiUrl = window.GROUP_ATTENDANCE_API_URL || '';
  const token = window.NEW_FAMILY_AUTH_TOKEN || '';

  if (!apiUrl) {
    throw new Error('尚未設定小組點名 API URL');
  }

  const response = await fetch(apiUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action, token, data })
  });

  const result = await response.json();
  if (!result.success) {
    throw new Error(result.message || '小組清單讀取失敗');
  }
  return result;
}

async function loadMeetingOptions() {
  meetingSelect.disabled = true;
  meetingSelect.innerHTML = '<option value="">載入聚會清單中...</option>';

  try {
    const groupConfig = await callSundayAttendanceApi('getGroupConfig');
    const options = flattenMeetingOptions(groupConfig);

    meetingSelect.innerHTML = '<option value="">請選擇</option>';
    options.forEach(item => {
      const option = document.createElement('option');
      option.value = item.name;
      option.textContent = `${item.category} / ${item.name}`;
      meetingSelect.appendChild(option);
    });
  } catch (error) {
    meetingSelect.innerHTML = '<option value="">聚會清單讀取失敗</option>';
    setNotice(formNotice, error.message || String(error), 'error');
  } finally {
    meetingSelect.disabled = false;
  }
}

function flattenMeetingOptions(groupConfig) {
  return Object.entries(groupConfig || {}).flatMap(([category, names]) => {
    if (!Array.isArray(names)) return [];
    return names
      .filter(Boolean)
      .map(name => ({ category, name: String(name).trim() }))
      .filter(item => item.name);
  });
}

async function loadSettlementStatusOptions() {
  settlementStatusSelect.disabled = true;
  settlementStatusSelect.innerHTML = '<option value="">載入小組清單中...</option>';

  try {
    const result = await callGroupAttendanceApi('getGroups');
    const groupNames = (result.groups || [])
      .map(group => String(group.name || '').trim())
      .filter(Boolean);

    settlementStatusSelect.innerHTML = '<option value="">請選擇</option>';
    groupNames.forEach(name => {
      const option = document.createElement('option');
      option.value = name;
      option.textContent = name;
      settlementStatusSelect.appendChild(option);
    });

    const visitOption = document.createElement('option');
    visitOption.value = '請安拜訪';
    visitOption.textContent = '請安拜訪';
    settlementStatusSelect.appendChild(visitOption);
  } catch (error) {
    settlementStatusSelect.innerHTML = '<option value="請安拜訪">請安拜訪</option>';
    setNotice(formNotice, error.message || String(error), 'error');
  } finally {
    settlementStatusSelect.disabled = false;
  }
}

function switchTab(tabName) {
  document.querySelectorAll('.tab').forEach(button => {
    button.classList.toggle('active', button.dataset.tab === tabName);
  });

  document.getElementById('formPanel').hidden = tabName !== 'form';
  document.getElementById('trackingPanel').hidden = tabName !== 'tracking';
  document.getElementById('closedPanel').hidden = tabName !== 'closed';

  if (tabName === 'tracking') loadTrackingCases();
  if (tabName === 'closed') loadClosedCases();
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

  trackingContent.className = 'table-wrap';
  trackingContent.textContent = '';
  trackingContent.appendChild(buildCaseTable(rows, true));
}

async function loadClosedCases() {
  setNotice(closedNotice, '');
  closedContent.className = 'empty';
  closedContent.textContent = '載入中...';
  closedSearchBtn.disabled = true;

  try {
    const result = await callApi('getClosedCases', {
      name: document.getElementById('closedName').value,
      startDate: document.getElementById('closedStartDate').value,
      endDate: document.getElementById('closedEndDate').value
    });
    renderClosedCases(result.data || []);
  } catch (error) {
    closedContent.className = 'empty';
    closedContent.textContent = '讀取失敗';
    setNotice(closedNotice, error.message || String(error), 'error');
  } finally {
    closedSearchBtn.disabled = false;
  }
}

function renderClosedCases(rows) {
  closedCount.textContent = `共 ${rows.length} 筆`;

  if (!rows.length) {
    closedContent.className = 'empty';
    closedContent.textContent = '沒有符合條件的已結案資料';
    return;
  }

  closedContent.className = 'table-wrap';
  closedContent.textContent = '';
  closedContent.appendChild(buildCaseTable(rows, false));
}

function buildCaseTable(rows, selectable) {
  const table = document.createElement('table');
  const thead = document.createElement('thead');
  const tbody = document.createElement('tbody');
  const headRow = document.createElement('tr');

  headRow.innerHTML = `${selectable ? '<th class="check-cell">結案</th>' : ''}${visibleColumns.map(column => `<th>${escapeHtml(column)}</th>`).join('')}`;
  thead.appendChild(headRow);

  rows.forEach(item => {
    const row = document.createElement('tr');

    if (selectable) {
      const checkboxCell = document.createElement('td');
      checkboxCell.className = 'check-cell';
      checkboxCell.innerHTML = `<input type="checkbox" value="${item.rowNumber}" aria-label="勾選 ${escapeHtml(item['新家人姓名'] || '此筆資料')} 結案">`;
      row.appendChild(checkboxCell);
    }

    visibleColumns.forEach(column => {
      const cell = document.createElement('td');
      cell.textContent = item[column] || '';
      row.appendChild(cell);
    });

    tbody.appendChild(row);
  });

  table.appendChild(thead);
  table.appendChild(tbody);
  return table;
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
