const trackingColumns = [
  '新家人姓名',
  '參加的聚會是',
  '手機',
  '關懷同工',
  '邀約人',
  '日期',
  '落戶狀態',
  '備註',
  '會友名單狀態',
  '點名系統代碼',
  '主日點名小組'
];

const closedColumns = trackingColumns.filter(column => column !== '主日點名小組');

const editFields = [
  { name: '新家人姓名', label: '新家人姓名', required: true },
  { name: '參加的聚會是', label: '參加的聚會是', type: 'meeting', required: true },
  { name: '新家人性別', label: '新家人性別', type: 'select', options: ['男', '女'] },
  { name: '新家人工作', label: '新家人工作' },
  { name: '年齡 -', label: '年齡 -' },
  { name: '有參加過基督教的崇拜嗎 ?', label: '有參加過基督教的崇拜嗎 ?', type: 'select', options: ['有', '沒有', '不確定'] },
  { name: '今天為什麼來到林口教會的呢 ?\n(朋友介紹請於其他填入朋友的姓名)', label: '今天為什麼來到林口教會的呢 ?', type: 'textarea', full: true },
  { name: '表單號', label: '表單號' },
  { name: '關懷同工', label: '關懷同工' },
  { name: '地址', label: '地址', full: true },
  { name: '市話', label: '市話' },
  { name: '手機', label: '手機' },
  { name: '日期', label: '日期', inputType: 'date' },
  { name: '落戶狀態', label: '落戶狀態', type: 'settlement' },
  { name: '邀約人', label: '邀約人' },
  { name: '備註', label: '備註', type: 'textarea', full: true },
  { name: '會友名單狀態', label: '會友名單狀態', type: 'select', options: ['已加入', '已存在'] },
  { name: '點名系統代碼', label: '點名系統代碼' }
];

const form = document.getElementById('newFamilyForm');
const formNotice = document.getElementById('formNotice');
const trackingNotice = document.getElementById('trackingNotice');
const submitBtn = document.getElementById('submitBtn');
const trackingSearchBtn = document.getElementById('trackingSearchBtn');
const addMembersBtn = document.getElementById('addMembersBtn');
const closeBtn = document.getElementById('closeBtn');
const trackingContent = document.getElementById('trackingContent');
const caseCount = document.getElementById('caseCount');
const dateField = document.getElementById('date');
const meetingSelect = document.getElementById('meeting');
const closedNotice = document.getElementById('closedNotice');
const closedContent = document.getElementById('closedContent');
const closedCount = document.getElementById('closedCount');
const closedSearchBtn = document.getElementById('closedSearchBtn');
const editModal = document.getElementById('editModal');
const editCaseForm = document.getElementById('editCaseForm');
const editFieldContainer = document.getElementById('editFields');
const editNotice = document.getElementById('editNotice');
const editSaveBtn = document.getElementById('editSaveBtn');
const editSubtitle = document.getElementById('editSubtitle');
const analysisCount = document.getElementById('analysisCount');
const analysisOpenBtn = document.getElementById('analysisOpenBtn');
const analysisPreview = document.getElementById('analysisPreview');
const analysisNotice = document.getElementById('analysisNotice');
const analysisYear = document.getElementById('analysisYear');
const analysisStartDate = document.getElementById('analysisStartDate');
const analysisEndDate = document.getElementById('analysisEndDate');
const analysisStatusFilter = document.getElementById('analysisStatusFilter');
const analysisModal = document.getElementById('analysisModal');
const analysisSubtitle = document.getElementById('analysisSubtitle');
const analysisModalContent = document.getElementById('analysisModalContent');

let meetingOptions = [];
let settlementOptions = ['請安拜訪'];
let editingCase = null;
let trackingCases = [];
let firebaseCacheModulePromise = null;
let memberDirectoryPromise = null;

const newFamilyCacheTtl = 19800;
const newFamilyListActions = new Set(['getTrackingCases', 'getClosedCases']);

dateField.valueAsDate = new Date();
analysisYear.value = new Date().getFullYear();
loadMeetingOptions();
loadSettlementStatusOptions();
setAnalysisRange('year');

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

trackingSearchBtn.addEventListener('click', loadTrackingCases);
addMembersBtn.addEventListener('click', addSelectedMembers);
closeBtn.addEventListener('click', closeSelectedCases);
closedSearchBtn.addEventListener('click', loadClosedCases);
analysisOpenBtn.addEventListener('click', openAnalysisModal);
analysisYear.addEventListener('change', () => setAnalysisRange('year'));
analysisStartDate.addEventListener('change', refreshAnalysisPreview);
analysisEndDate.addEventListener('change', refreshAnalysisPreview);
analysisStatusFilter.addEventListener('change', refreshAnalysisPreview);
document.querySelectorAll('[data-analysis-range]').forEach(button => {
  button.addEventListener('click', () => setAnalysisRange(button.dataset.analysisRange));
});
document.getElementById('editCloseBtn').addEventListener('click', closeEditModal);
document.getElementById('editCancelBtn').addEventListener('click', closeEditModal);
document.getElementById('analysisCloseBtn').addEventListener('click', closeAnalysisModal);
document.getElementById('analysisDoneBtn').addEventListener('click', closeAnalysisModal);
editModal.addEventListener('click', event => {
  if (event.target === editModal) closeEditModal();
});
analysisModal.addEventListener('click', event => {
  if (event.target === analysisModal) closeAnalysisModal();
});
editCaseForm.addEventListener('submit', saveTrackingCase);

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

async function callCachedListApi(action, data = {}) {
  if (!newFamilyListActions.has(action)) {
    return callApi(action, data);
  }

  try {
    const cache = await getFirebaseCacheModule();
    return await cache.cacheGetOrFetch(
      action,
      '_default',
      () => callApi(action),
      newFamilyCacheTtl
    );
  } catch (error) {
    console.warn('[new-family-cache] fallback to GAS', error);
    return callApi(action, data);
  }
}

function getFirebaseCacheModule() {
  if (!firebaseCacheModulePromise) {
    firebaseCacheModulePromise = import('../../firebase/firebase-cache.js');
  }
  return firebaseCacheModulePromise;
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

async function callSundayAttendancePayloadApi(action, payload) {
  const apiUrl = window.SUNDAY_ATTENDANCE_API_URL || '';

  if (!apiUrl) {
    throw new Error('尚未設定主日出席 API URL');
  }

  const response = await fetch(apiUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action, payload })
  });

  const result = await response.json();
  if (result.error) {
    throw new Error(result.error);
  }
  return result.data;
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
    meetingOptions = options;

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
  try {
    const result = await callGroupAttendanceApi('getGroups');
    const groupNames = (result.groups || [])
      .map(group => String(group.name || '').trim())
      .filter(Boolean);
    settlementOptions = Array.from(new Set([...groupNames, '請安拜訪']));
  } catch (error) {
    settlementOptions = ['請安拜訪'];
    setNotice(formNotice, error.message || String(error), 'error');
  }
}

function switchTab(tabName) {
  document.querySelectorAll('.tab').forEach(button => {
    button.classList.toggle('active', button.dataset.tab === tabName);
  });

  document.getElementById('formPanel').hidden = tabName !== 'form';
  document.getElementById('trackingPanel').hidden = tabName !== 'tracking';
  document.getElementById('closedPanel').hidden = tabName !== 'closed';
  document.getElementById('analysisPanel').hidden = tabName !== 'analysis';

  if (tabName === 'tracking') loadTrackingCases();
  if (tabName === 'closed') loadClosedCases();
  if (tabName === 'analysis') refreshAnalysisPreview();
}

async function loadTrackingCases() {
  setNotice(trackingNotice, '');
  trackingContent.className = 'empty';
  trackingContent.textContent = '載入中...';
  trackingSearchBtn.disabled = true;
  addMembersBtn.disabled = true;
  closeBtn.disabled = true;

  try {
    const filters = {
      name: document.getElementById('trackingName').value,
      startDate: document.getElementById('trackingStartDate').value,
      endDate: document.getElementById('trackingEndDate').value
    };
    const result = await callCachedListApi('getTrackingCases');
    const rows = await enrichRowsWithSundayGroups(filterCases(result.data || [], filters));
    renderTrackingCases(rows);
  } catch (error) {
    trackingContent.className = 'empty';
    trackingContent.textContent = '讀取失敗';
    setNotice(trackingNotice, error.message || String(error), 'error');
  } finally {
    trackingSearchBtn.disabled = false;
    addMembersBtn.disabled = false;
    closeBtn.disabled = false;
  }
}

function renderTrackingCases(rows) {
  trackingCases = rows;
  caseCount.textContent = `共 ${rows.length} 筆`;

  if (!rows.length) {
    trackingContent.className = 'empty';
    trackingContent.textContent = '目前沒有追蹤中的資料';
    return;
  }

  trackingContent.className = 'table-wrap';
  trackingContent.textContent = '';
  trackingContent.appendChild(buildCaseTable(rows, true, trackingColumns));
}

async function addSelectedMembers() {
  const selectedCases = getSelectedTrackingCases();

  if (!selectedCases.length) {
    setNotice(trackingNotice, '請先勾選要加入會友名單的資料', 'error');
    return;
  }

  const selectedNames = selectedCases
    .map(item => item['新家人姓名'])
    .filter(Boolean);
  if (!selectedNames.length) {
    setNotice(trackingNotice, '勾選資料沒有可加入的姓名', 'error');
    return;
  }

  if (!confirm(`確認將 ${selectedNames.length} 位加入會友名單嗎？`)) return;

  addMembersBtn.disabled = true;
  closeBtn.disabled = true;
  setNotice(trackingNotice, '加入會友名單中...');

  try {
    const results = [];

    for (const item of selectedCases) {
      const name = String(item['新家人姓名'] || '').trim();
      if (!name) {
        results.push({ ok: false, name: '未填姓名', message: '略過未填姓名資料' });
        continue;
      }

      const message = String(await callSundayAttendancePayloadApi('addMember', {
        name,
        gender: item['新家人性別'] || '',
        note: item['備註'] || '',
        isExcluded: false
      }) || '');
      let memberCode = extractMemberCode(message);
      const duplicate = message.includes('已存在');
      if (message.includes('成功')) {
        memberDirectoryPromise = null;
      }
      const existingMember = await findMemberRecord(name, memberCode);
      if (!memberCode && duplicate) memberCode = existingMember.memberCode;
      const sundayGroup = existingMember.sundayGroup || '';
      results.push({
        ok: message.includes('成功'),
        duplicate,
        rowNumber: item.rowNumber,
        memberCode,
        sundayGroup,
        name,
        message
      });
    }

    const successCount = results.filter(item => item.ok).length;
    const duplicateCount = results.filter(item => item.duplicate).length;
    const memberStatuses = results
      .filter(item => item.ok || item.duplicate)
      .map(item => ({
        rowNumber: item.rowNumber,
        status: item.ok ? '已加入' : '已存在',
        memberCode: item.memberCode || '',
        sundayGroup: item.sundayGroup || ''
      }));

    if (memberStatuses.length) {
      await callApi('markTrackingMemberStatuses', { items: memberStatuses });
      await loadTrackingCases();
    }

    const failed = results.filter(item => !item.ok && !item.duplicate);
    const suffix = failed.length
      ? `；未加入：${failed.map(item => `${item.name} ${item.message}`).join('、')}`
      : '';
    const duplicateText = duplicateCount ? `，已存在 ${duplicateCount} 位` : '';
    const codeText = results
      .filter(item => (item.ok || item.duplicate) && item.memberCode)
      .map(item => `${item.name} ${item.memberCode}`)
      .join('、');
    const codeSuffix = codeText ? `；代碼：${codeText}` : '';
    setNotice(trackingNotice, `已加入會友名單 ${successCount} 位${duplicateText}${codeSuffix}${suffix}`, failed.length ? 'error' : 'success');
  } catch (error) {
    setNotice(trackingNotice, error.message || String(error), 'error');
  } finally {
    addMembersBtn.disabled = false;
    closeBtn.disabled = false;
  }
}

function extractMemberCode(message) {
  const match = String(message || '').match(/編號[:：]\s*([A-Z]+\d+)/i);
  return match ? match[1].toUpperCase() : '';
}

async function getMemberDirectory() {
  if (!memberDirectoryPromise) {
    memberDirectoryPromise = callSundayAttendancePayloadApi('getAllMembers', {})
      .then(members => {
        const byName = new Map();
        const byCode = new Map();
        (Array.isArray(members) ? members : []).forEach(member => {
          const name = String(member[0] || '').trim();
          const memberCode = String(member[7] || '').trim();
          const sundayGroup = String(member[8] || '').trim();
          const record = { name, memberCode, sundayGroup };
          if (name) byName.set(name, record);
          if (memberCode) byCode.set(memberCode, record);
        });
        return { byName, byCode };
      })
      .catch(error => {
        memberDirectoryPromise = null;
        throw error;
      });
  }
  return memberDirectoryPromise;
}

async function findMemberRecord(name, memberCode) {
  try {
    const directory = await getMemberDirectory();
    const targetName = String(name || '').trim();
    const targetCode = String(memberCode || '').trim();
    return directory.byCode.get(targetCode) || directory.byName.get(targetName) || {
      name: targetName,
      memberCode: targetCode,
      sundayGroup: ''
    };
  } catch (error) {
    console.warn('[new-family] member lookup failed', error);
    return {
      name: String(name || '').trim(),
      memberCode: String(memberCode || '').trim(),
      sundayGroup: ''
    };
  }
}

async function enrichRowsWithSundayGroups(rows) {
  try {
    const directory = await getMemberDirectory();
    return rows.map(item => {
      const memberCode = String(item['點名系統代碼'] || '').trim();
      const name = String(item['新家人姓名'] || '').trim();
      const member = directory.byCode.get(memberCode) || directory.byName.get(name);
      if (!member || !member.sundayGroup) return item;
      return {
        ...item,
        '主日點名小組': member.sundayGroup
      };
    });
  } catch (error) {
    console.warn('[new-family] sunday group enrich skipped', error);
    return rows;
  }
}

async function loadClosedCases() {
  setNotice(closedNotice, '');
  closedContent.className = 'empty';
  closedContent.textContent = '載入中...';
  closedSearchBtn.disabled = true;

  try {
    const filters = {
      name: document.getElementById('closedName').value,
      startDate: document.getElementById('closedStartDate').value,
      endDate: document.getElementById('closedEndDate').value
    };
    const result = await callCachedListApi('getClosedCases');
    renderClosedCases(filterCases(result.data || [], filters));
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
  closedContent.appendChild(buildCaseTable(rows, false, closedColumns));
}

async function refreshAnalysisPreview() {
  setNotice(analysisNotice, '');
  analysisPreview.className = 'empty';
  analysisPreview.textContent = '載入分析資料中...';
  analysisOpenBtn.disabled = true;

  try {
    const dateRows = await getAnalysisDateRows();
    populateAnalysisStatusFilter(dateRows);
    const rows = filterAnalysisRowsByStatus(dateRows);
    analysisCount.textContent = `共 ${rows.length} 筆`;

    if (!rows.length) {
      analysisPreview.textContent = '這個範圍沒有已結案的新朋友資料';
      return;
    }

    analysisPreview.className = 'table-wrap analysis-table';
    analysisPreview.textContent = '';
    analysisPreview.appendChild(buildAnalysisDetailTable(rows));
  } catch (error) {
    analysisPreview.textContent = '分析資料讀取失敗';
    setNotice(analysisNotice, error.message || String(error), 'error');
  } finally {
    analysisOpenBtn.disabled = false;
  }
}

async function openAnalysisModal() {
  setNotice(analysisNotice, '');
  analysisOpenBtn.disabled = true;

  try {
    const dateRows = await getAnalysisDateRows();
    const rows = filterAnalysisRowsByStatus(dateRows);
    const pivot = buildSettlementPivot(rows);

    analysisSubtitle.textContent = `${analysisStartDate.value || '不限'} 至 ${analysisEndDate.value || '不限'}`;
    analysisModalContent.textContent = '';
    analysisModalContent.appendChild(buildAnalysisSummary(rows, pivot));

    const pivotWrap = document.createElement('div');
    pivotWrap.className = 'table-wrap analysis-table';
    pivotWrap.appendChild(buildAnalysisPivotTable(pivot, rows.length));
    analysisModalContent.appendChild(pivotWrap);

    analysisModal.hidden = false;
  } catch (error) {
    setNotice(analysisNotice, error.message || String(error), 'error');
  } finally {
    analysisOpenBtn.disabled = false;
  }
}

function closeAnalysisModal() {
  analysisModal.hidden = true;
  analysisModalContent.textContent = '';
}

async function getAnalysisDateRows() {
  const result = await callCachedListApi('getClosedCases');
  const filters = {
    startDate: analysisStartDate.value,
    endDate: analysisEndDate.value
  };
  return filterCases(result.data || [], filters);
}

function filterAnalysisRowsByStatus(rows) {
  const status = analysisStatusFilter.value;
  return status
    ? rows.filter(item => normalizeSettlementStatus(item['落戶狀態']) === status)
    : rows;
}

function setAnalysisRange(rangeKey) {
  const year = Number(analysisYear.value) || new Date().getFullYear();
  const ranges = {
    year: ['01-01', '12-31'],
    h1: ['01-01', '06-30'],
    h2: ['07-01', '12-31'],
    q1: ['01-01', '03-31'],
    q2: ['04-01', '06-30'],
    q3: ['07-01', '09-30'],
    q4: ['10-01', '12-31']
  };
  const range = ranges[rangeKey] || ranges.year;
  analysisStartDate.value = `${year}-${range[0]}`;
  analysisEndDate.value = `${year}-${range[1]}`;

  if (!document.getElementById('analysisPanel').hidden) {
    refreshAnalysisPreview();
  }
}

function populateAnalysisStatusFilter(rows) {
  const current = analysisStatusFilter.value;
  const statuses = Array.from(new Set(rows.map(item => normalizeSettlementStatus(item['落戶狀態'])))).sort();
  analysisStatusFilter.innerHTML = '<option value="">全部</option>';
  statuses.forEach(status => {
    const option = document.createElement('option');
    option.value = status;
    option.textContent = status;
    analysisStatusFilter.appendChild(option);
  });
  if (statuses.includes(current)) {
    analysisStatusFilter.value = current;
  }
}

function buildSettlementPivot(rows) {
  const counts = new Map();
  rows.forEach(item => {
    const status = normalizeSettlementStatus(item['落戶狀態']);
    counts.set(status, (counts.get(status) || 0) + 1);
  });
  return Array.from(counts.entries())
    .map(([status, count]) => ({ status, count }))
    .sort((a, b) => b.count - a.count || a.status.localeCompare(b.status));
}

function normalizeSettlementStatus(value) {
  return String(value || '').trim() || '未填落戶狀態';
}

function buildAnalysisSummary(rows, pivot) {
  const summary = document.createElement('div');
  summary.className = 'analysis-summary';
  summary.innerHTML = `
    <div class="summary-item"><span>總清單人數</span><strong>${rows.length}</strong></div>
    <div class="summary-item"><span>落戶狀態種類</span><strong>${pivot.length}</strong></div>
  `;
  return summary;
}

function buildAnalysisPivotTable(pivot, total) {
  const table = document.createElement('table');
  const body = pivot.length
    ? pivot.map(item => {
      const percent = total ? Math.round((item.count / total) * 1000) / 10 : 0;
      return `<tr><td>${escapeHtml(item.status)}</td><td>${item.count}</td><td>${percent}%</td></tr>`;
    }).join('')
    : '<tr><td colspan="3">沒有符合範圍的資料</td></tr>';
  table.innerHTML = `
    <thead><tr><th>落戶狀態</th><th>人數</th><th>比例</th></tr></thead>
    <tbody>${body}</tbody>
  `;
  return table;
}

function buildAnalysisDetailTable(rows) {
  const table = document.createElement('table');
  const body = rows.length
    ? rows.map(item => `
      <tr>
        <td>${escapeHtml(item['新家人姓名'] || '')}</td>
        <td>${escapeHtml(item['日期'] || '')}</td>
        <td>${escapeHtml(normalizeSettlementStatus(item['落戶狀態']))}</td>
        <td>${escapeHtml(item['點名系統代碼'] || '')}</td>
      </tr>
    `).join('')
    : '<tr><td colspan="4">沒有符合範圍的明細</td></tr>';
  table.innerHTML = `
    <thead><tr><th>姓名</th><th>日期</th><th>落戶狀態</th><th>點名系統代碼</th></tr></thead>
    <tbody>${body}</tbody>
  `;
  return table;
}

function filterCases(rows, filters) {
  const name = String(filters.name || '').trim().toLowerCase();
  const startDate = String(filters.startDate || '').trim();
  const endDate = String(filters.endDate || '').trim();

  return rows.filter(item => {
    const caseName = String(item['新家人姓名'] || '').toLowerCase();
    const caseDate = String(item['日期'] || '').trim();

    if (name && !caseName.includes(name)) return false;
    if (startDate && (!caseDate || caseDate < startDate)) return false;
    if (endDate && (!caseDate || caseDate > endDate)) return false;
    return true;
  });
}

function buildCaseTable(rows, selectable, columns) {
  const table = document.createElement('table');
  const thead = document.createElement('thead');
  const tbody = document.createElement('tbody');
  const headRow = document.createElement('tr');

  headRow.innerHTML = `${selectable ? '<th class="check-cell">結案</th><th class="action-cell">編輯</th>' : ''}${columns.map(column => `<th>${escapeHtml(getColumnLabel(column))}</th>`).join('')}`;
  thead.appendChild(headRow);

  rows.forEach(item => {
    const row = document.createElement('tr');

    if (selectable) {
      const checkboxCell = document.createElement('td');
      checkboxCell.className = 'check-cell';
      checkboxCell.innerHTML = `<input type="checkbox" value="${item.rowNumber}" aria-label="勾選 ${escapeHtml(item['新家人姓名'] || '此筆資料')} 結案">`;
      row.appendChild(checkboxCell);

      const editCell = document.createElement('td');
      editCell.className = 'action-cell';
      const editButton = document.createElement('button');
      editButton.type = 'button';
      editButton.className = 'btn secondary';
      editButton.textContent = '編輯';
      editButton.addEventListener('click', () => openEditModal(item));
      editCell.appendChild(editButton);
      row.appendChild(editCell);
    }

    columns.forEach(column => {
      const cell = document.createElement('td');
      if (column === '主日點名小組') {
        cell.appendChild(buildSundayGroupTag(item[column] || ''));
      } else {
        cell.textContent = item[column] || '';
      }
      row.appendChild(cell);
    });

    tbody.appendChild(row);
  });

  table.appendChild(thead);
  table.appendChild(tbody);
  return table;
}

function buildSundayGroupTag(value) {
  const groupName = String(value || '').trim();
  const tag = document.createElement('button');
  tag.type = 'button';
  tag.className = groupName ? 'group-tag' : 'group-tag empty-tag';
  tag.textContent = groupName ? shortenGroupName(groupName) : '未';
  tag.title = groupName || '主日點名尚無小組';
  tag.setAttribute('aria-label', groupName ? `主日點名小組：${groupName}` : '主日點名尚無小組');
  tag.addEventListener('click', () => tag.classList.toggle('expanded'));
  return tag;
}

function getColumnLabel(column) {
  return column === '主日點名小組' ? '現行小組' : column;
}

function shortenGroupName(groupName) {
  const normalized = String(groupName || '').trim();
  return normalized.length > 4 ? `${normalized.slice(0, 3)}...` : normalized;
}

function openEditModal(item) {
  editingCase = item;
  editSubtitle.textContent = item['新家人姓名']
    ? `${item['新家人姓名']}，表單號 ${item['表單號'] || '未填'}`
    : `表單號 ${item['表單號'] || '未填'}`;
  setNotice(editNotice, '');
  editFieldContainer.textContent = '';

  editFields.forEach((field, index) => {
    const wrapper = document.createElement('div');
    wrapper.className = `field ${field.full ? 'full' : ''}`.trim();

    const label = document.createElement('label');
    label.textContent = field.label;
    label.htmlFor = `edit-field-${index}`;

    const control = createEditControl(field, item[field.name] || '');
    control.id = `edit-field-${index}`;
    control.name = field.name;
    if (field.required) control.required = true;

    wrapper.appendChild(label);
    wrapper.appendChild(control);
    editFieldContainer.appendChild(wrapper);
  });

  editModal.hidden = false;
}

function closeEditModal() {
  editingCase = null;
  editModal.hidden = true;
  editCaseForm.reset();
  editFieldContainer.textContent = '';
  setNotice(editNotice, '');
}

function createEditControl(field, value) {
  if (field.type === 'textarea') {
    const textarea = document.createElement('textarea');
    textarea.value = value;
    return textarea;
  }

  if (field.type === 'meeting') {
    return createSelectControl(
      meetingOptions.map(item => ({ value: item.name, text: `${item.category} / ${item.name}` })),
      value
    );
  }

  if (field.type === 'settlement') {
    return createSelectControl(
      settlementOptions.map(name => ({ value: name, text: name })),
      value
    );
  }

  if (field.type === 'select') {
    return createSelectControl(
      field.options.map(option => ({ value: option, text: option })),
      value
    );
  }

  const input = document.createElement('input');
  input.type = field.inputType || 'text';
  input.value = value;
  return input;
}

function createSelectControl(options, value) {
  const select = document.createElement('select');
  const blank = document.createElement('option');
  blank.value = '';
  blank.textContent = '請選擇';
  select.appendChild(blank);

  const values = new Set();
  options.forEach(item => {
    if (!item.value || values.has(item.value)) return;
    values.add(item.value);
    const option = document.createElement('option');
    option.value = item.value;
    option.textContent = item.text;
    select.appendChild(option);
  });

  if (value && !values.has(value)) {
    const current = document.createElement('option');
    current.value = value;
    current.textContent = value;
    select.appendChild(current);
  }

  select.value = value;
  return select;
}

async function saveTrackingCase(event) {
  event.preventDefault();
  if (!editingCase) return;

  setNotice(editNotice, '儲存中...');
  editSaveBtn.disabled = true;

  try {
    const result = await callApi('updateTrackingCase', {
      rowNumber: editingCase.rowNumber,
      values: Object.fromEntries(new FormData(editCaseForm).entries())
    });
    setNotice(trackingNotice, result.message, 'success');
    closeEditModal();
    await loadTrackingCases();
  } catch (error) {
    setNotice(editNotice, error.message || String(error), 'error');
  } finally {
    editSaveBtn.disabled = false;
  }
}

async function closeSelectedCases() {
  const selected = getSelectedTrackingCases().map(item => item.rowNumber);

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

function getSelectedTrackingCases() {
  const selectedRows = new Set(Array.from(trackingContent.querySelectorAll('input[type="checkbox"]:checked'))
    .map(input => Number(input.value)));
  return trackingCases.filter(item => selectedRows.has(Number(item.rowNumber)));
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
