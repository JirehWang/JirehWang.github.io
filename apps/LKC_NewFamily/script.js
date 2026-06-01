const trackingColumns = [
  '新家人姓名',
  '新家人性別',
  '參加的聚會是',
  '手機',
  '關懷同工',
  '邀約人',
  '日期',
  '落戶狀態',
  '備註',
  '會友名單狀態',
  '點名系統代碼'
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
const analysisExportDetailBtn = document.getElementById('analysisExportDetailBtn');
const analysisExportSummaryBtn = document.getElementById('analysisExportSummaryBtn');

let meetingOptions = [];
let settlementOptions = ['請安拜訪'];
let editingCase = null;
let trackingCases = [];
let firebaseCacheModulePromise = null;
let memberDirectoryPromise = null;
let currentAnalysisRows = [];
let currentAnalysisPivot = [];

const newFamilyCacheTtl = 19800;
const newFamilyListActions = new Set(['getTrackingCases', 'getClosedCases']);
const stoppedAttendanceStatus = '停止聚會';

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

  if (!confirm('確認要送出此筆新家人資料嗎？')) {
    return;
  }

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
analysisExportDetailBtn.addEventListener('click', exportAnalysisDetail);
analysisExportSummaryBtn.addEventListener('click', exportAnalysisSummary);
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

document.addEventListener('click', () => {
  document.querySelectorAll('.action-menu').forEach(menu => {
    menu.hidden = true;
  });
});

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
    renderTrackingCases(filterCases(result.data || [], filters));
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
          const isExcluded = normalizeBoolean(member[4]);
          const record = { name, memberCode, sundayGroup, isExcluded };
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
      sundayGroup: '',
      isExcluded: false
    };
  } catch (error) {
    console.warn('[new-family] member lookup failed', error);
    return {
      name: String(name || '').trim(),
      memberCode: String(memberCode || '').trim(),
      sundayGroup: '',
      isExcluded: false
    };
  }
}

async function enrichRowsWithSundayMemberData(rows) {
  try {
    const directory = await getMemberDirectory();
    return rows.map(item => {
      const memberCode = String(item['點名系統代碼'] || '').trim();
      const name = String(item['新家人姓名'] || '').trim();
      const member = directory.byCode.get(memberCode) || directory.byName.get(name);
      if (!member) return item;
      return {
        ...item,
        '主日點名小組': member.sundayGroup || item['主日點名小組'] || '',
        displaySettlementStatus: member.isExcluded
          ? stoppedAttendanceStatus
          : normalizeSettlementStatus(item['落戶狀態'])
      };
    });
  } catch (error) {
    console.warn('[new-family] sunday member enrich skipped', error);
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
    const rows = await enrichRowsWithSundayMemberData(filterCases(result.data || [], filters));
    renderClosedCases(rows);
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
  analysisExportDetailBtn.disabled = true;

  try {
    const dateRows = await getAnalysisDateRows();
    populateAnalysisStatusFilter(dateRows);
    const rows = filterAnalysisRowsByStatus(dateRows);
    currentAnalysisRows = rows;
    currentAnalysisPivot = buildSettlementPivot(rows);
    analysisCount.textContent = `共 ${rows.length} 筆`;

    if (!rows.length) {
      analysisPreview.textContent = '這個範圍沒有已結案的新朋友資料';
      return;
    }

    analysisPreview.className = 'table-wrap analysis-table';
    analysisPreview.textContent = '';
    analysisPreview.appendChild(buildAnalysisDetailTable(rows));
  } catch (error) {
    currentAnalysisRows = [];
    currentAnalysisPivot = [];
    analysisPreview.textContent = '分析資料讀取失敗';
    setNotice(analysisNotice, error.message || String(error), 'error');
  } finally {
    analysisOpenBtn.disabled = false;
    analysisExportDetailBtn.disabled = !currentAnalysisRows.length;
  }
}

async function openAnalysisModal() {
  setNotice(analysisNotice, '');
  analysisOpenBtn.disabled = true;

  try {
    const dateRows = await getAnalysisDateRows();
    const rows = filterAnalysisRowsByStatus(dateRows);
    const pivot = buildSettlementPivot(rows);
    currentAnalysisRows = rows;
    currentAnalysisPivot = pivot;

    analysisSubtitle.textContent = `${analysisStartDate.value || '不限'} 至 ${analysisEndDate.value || '不限'}`;
    analysisModalContent.textContent = '';
    analysisModalContent.appendChild(buildAnalysisSummary(rows, pivot));

    const pivotWrap = document.createElement('div');
    pivotWrap.className = 'table-wrap analysis-table';
    pivotWrap.appendChild(buildAnalysisPivotTable(pivot, rows.length));
    analysisModalContent.appendChild(pivotWrap);

    analysisModal.hidden = false;
    analysisExportSummaryBtn.disabled = !rows.length;
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

async function exportAnalysisDetail() {
  await exportCombinedWorkbook();
}

async function exportAnalysisSummary() {
  await exportCombinedWorkbook();
}

async function exportCombinedWorkbook() {
  try {
    const rows = currentAnalysisRows.length
      ? currentAnalysisRows
      : filterAnalysisRowsByStatus(await getAnalysisDateRows());
    if (!rows.length) {
      setNotice(analysisNotice, '沒有可匯出的資料', 'error');
      return;
    }

    // Determine unique column groups dynamically
    const colGroupsMap = new Map();
    rows.forEach(item => {
      const yq = getYearQuarter(item['日期']);
      if (yq) {
        colGroupsMap.set(`${yq.year}_${yq.quarter}`, yq);
      }
    });

    const sortedYqKeys = Array.from(colGroupsMap.keys()).sort((a, b) => {
      const [ay, aq] = a.split('_');
      const [by, bq] = b.split('_');
      if (ay !== by) return ay - by;
      return aq.localeCompare(bq);
    });

    const columnGroups = [];
    const yearsPresent = Array.from(new Set(sortedYqKeys.map(k => k.split('_')[0])));

    yearsPresent.forEach(year => {
      const qKeys = sortedYqKeys.filter(k => k.startsWith(year + '_'));
      if (qKeys.length > 1) {
        columnGroups.push({ year: parseInt(year, 10), quarter: 'Q1-Q4' });
      }
      qKeys.forEach(k => {
        columnGroups.push(colGroupsMap.get(k));
      });
    });

    const numGroups = columnGroups.length;
    const numCols = 15 + 3 * numGroups;
    
    // Initialize Sheet 1: Pivot Table matrix
    const matrix = [];
    for (let r = 0; r < 35; r++) {
      matrix.push(new Array(numCols).fill(null));
    }

    const groups = [
      '葡萄樹',
      '以斯帖',
      '松年團契',
      '棕樹',
      '芥菜種',
      '香柏樹',
      '橄欖樹',
      '種子',
      '提摩太',
      '恩典團契',
      '尚未落戶'
    ];

    function mapGroup(status) {
      const s = String(status || '').trim();
      if (s === '松年' || s === '松年團契') return '松年團契';
      if (s === '恩典' || s === '恩典團契') return '恩典團契';
      if (groups.includes(s)) return s;
      return '尚未落戶';
    }

    function getYearQuarter(dateStr) {
      if (!dateStr) return null;
      const match = dateStr.match(/^(\d{4})-(\d{2})-\d{2}$/);
      if (!match) return null;
      const year = parseInt(match[1], 10);
      const month = parseInt(match[2], 10);
      let quarter = '';
      if (month >= 1 && month <= 3) quarter = 'Q1';
      else if (month >= 4 && month <= 6) quarter = 'Q2';
      else if (month >= 7 && month <= 9) quarter = 'Q3';
      else if (month >= 10 && month <= 12) quarter = 'Q4';
      return { year, quarter };
    }

    function getServiceType(meetingName) {
      const name = String(meetingName || '');
      if (name.includes('聯合')) return '聯合';
      if (name.includes('台語')) return '台語';
      if (name.includes('華語')) return '華語';
      return '華語';
    }

    const years = [2024, 2025, 2026];
    years.forEach((year, yIdx) => {
      const colIdx = 1 + yIdx * 2;
      const yrRows = rows.filter(item => {
        const yq = getYearQuarter(item['日期']);
        return yq && yq.year === year;
      });

      const totalCount = yrRows.length;
      let dateRangeStr = '';
      if (yrRows.length > 0) {
        const dates = yrRows.map(item => item['日期']).filter(Boolean).sort();
        if (dates.length > 0) {
          const minDate = dates[0].replace(/-/g, '/');
          const maxDate = dates[dates.length - 1].replace(/-/g, '/');
          dateRangeStr = `（${minDate}-${maxDate}）`;
        }
      }

      const groupCounts = {};
      groups.forEach(g => { groupCounts[g] = 0; });
      let stoppedCount = 0;
      let visitCount = 0;

      yrRows.forEach(item => {
        const status = String(item['落戶狀態'] || '').trim();
        if (status === '停止聚會') {
          stoppedCount++;
        } else if (status === '請安拜訪') {
          visitCount++;
        } else {
          const mapped = mapGroup(status);
          groupCounts[mapped] = (groupCounts[mapped] || 0) + 1;
        }
      });

      const validCount = Object.values(groupCounts).reduce((a, b) => a + b, 0);

      matrix[1][colIdx] = `${year}年度新家人落戶說明：`;
      matrix[3][colIdx] = `📌 留名卡總數：${totalCount} 筆`;
      matrix[4][colIdx] = dateRangeStr;
      matrix[6][colIdx] = `✅ 有效名單：${validCount} 位`;
      matrix[7][colIdx] = `目前落戶情況如下：`;

      const groupEmojis = {
        '葡萄樹': '🍇',
        '以斯帖': '👑',
        '松年團契': '🌿',
        '棕樹': '🌴',
        '芥菜種': '🌱',
        '香柏樹': '🌲',
        '橄欖樹': '🫒',
        '種子': '🌾',
        '提摩太': '📖',
        '恩典團契': '💒',
        '尚未落戶': '🕊️'
      };

      groups.forEach((g, gIdx) => {
        const emoji = groupEmojis[g] || '';
        matrix[8 + gIdx][colIdx] = `• ${emoji} ${g}：${groupCounts[g]} 位`;
      });

      matrix[20][colIdx] = `❌ 停止聚會名單：${stoppedCount} 位`;
      matrix[21][colIdx] = `（曾落戶小組後因故離開）`;
      matrix[23][colIdx] = `📋 請安拜訪名單：${visitCount} 位`;
    });

    const allValidCount = rows.filter(item => {
      const status = String(item['落戶狀態'] || '').trim();
      return status !== '停止聚會' && status !== '請安拜訪';
    }).length;
    matrix[25][1] = `** 統計表數字同各年度新家人落戶說明 (參照初始資料 - 以${allValidCount}筆有效資料分析)`;
    matrix[27][1] = `**2025/03/16 三樓禮拜堂啟用`;
    matrix[28][1] = `**2025/10/19 一樓禮拜堂啟用`;
    matrix[29][1] = `**2026/03/01 台華語同步禮拜10:00`;

    // Populate Pivot Table Headers
    matrix[1][7] = '新家人落戶統計';
    columnGroups.forEach((group, groupIdx) => {
      const startCol = 9 + 3 * groupIdx;
      matrix[1][startCol] = group.year;
      matrix[2][startCol] = group.quarter;
      matrix[3][startCol] = '聯合';
      matrix[3][startCol + 1] = '台語';
      matrix[3][startCol + 2] = '華語';
    });

    const grandTotalStart = 9 + 3 * numGroups;
    matrix[1][grandTotalStart] = '總計';
    matrix[3][grandTotalStart] = '聯合';
    matrix[3][grandTotalStart + 1] = '台語';
    matrix[3][grandTotalStart + 2] = '華語';

    const pctCol = 12 + 3 * numGroups;
    matrix[1][pctCol] = '%';

    matrix[3][7] = '小組別';
    matrix[3][8] = '是否受邀';

    // Helper functions for column names
    function colIndex(colLetter) {
      let index = 0;
      for (let i = 0; i < colLetter.length; i++) {
        index = index * 26 + (colLetter.charCodeAt(i) - 64);
      }
      return index - 1;
    }

    const activeColLetters = [];
    columnGroups.forEach((group, groupIdx) => {
      if (group.quarter !== 'Q1-Q4') {
        activeColLetters.push(columnName(9 + 3 * groupIdx));
      }
    });

    // Populate Pivot Table Data Rows
    groups.forEach((g, gIdx) => {
      const rIdx = 4 + 2 * gIdx; // Excel row R = rIdx + 1
      const R_invited = rIdx + 1;
      const R_notInvited = rIdx + 2;

      matrix[rIdx][7] = g;
      matrix[rIdx][8] = '受邀';
      matrix[rIdx + 1][8] = '非受邀';

      // Counts for each Column Group
      columnGroups.forEach((group, groupIdx) => {
        const startCol = 9 + 3 * groupIdx;
        const matchingCases = rows.filter(item => {
          if (mapGroup(item['落戶狀態']) !== g) return false;
          const yq = getYearQuarter(item['日期']);
          if (!yq || yq.year !== group.year) return false;
          if (group.quarter !== 'Q1-Q4' && yq.quarter !== group.quarter) return false;
          return true;
        });

        // Split by invited and service
        const invitedCases = matchingCases.filter(item => item['邀約人'] && String(item['邀約人']).trim());
        const notInvitedCases = matchingCases.filter(item => !item['邀約人'] || !String(item['邀約人']).trim());

        function countService(cases, type) {
          const count = cases.filter(item => getServiceType(item['參加的聚會是']) === type).length;
          return count > 0 ? count : null;
        }

        matrix[rIdx][startCol] = countService(invitedCases, '聯合');
        matrix[rIdx][startCol + 1] = countService(invitedCases, '台語');
        matrix[rIdx][startCol + 2] = countService(invitedCases, '華語');

        matrix[rIdx + 1][startCol] = countService(notInvitedCases, '聯合');
        matrix[rIdx + 1][startCol + 1] = countService(notInvitedCases, '台語');
        matrix[rIdx + 1][startCol + 2] = countService(notInvitedCases, '華語');
      });

      // Sum formulas for grand total columns
      const colAE = columnName(grandTotalStart);
      const colAF = columnName(grandTotalStart + 1);
      const colAG = columnName(grandTotalStart + 2);

      matrix[rIdx][grandTotalStart] = '=' + activeColLetters.map(col => `${col}${R_invited}`).join('+');
      matrix[rIdx][grandTotalStart + 1] = '=' + activeColLetters.map(col => `${columnName(colIndex(col) + 1)}${R_invited}`).join('+');
      matrix[rIdx][grandTotalStart + 2] = '=' + activeColLetters.map(col => `${columnName(colIndex(col) + 2)}${R_invited}`).join('+');

      matrix[rIdx + 1][grandTotalStart] = '=' + activeColLetters.map(col => `${col}${R_notInvited}`).join('+');
      matrix[rIdx + 1][grandTotalStart + 1] = '=' + activeColLetters.map(col => `${columnName(colIndex(col) + 1)}${R_notInvited}`).join('+');
      matrix[rIdx + 1][grandTotalStart + 2] = '=' + activeColLetters.map(col => `${columnName(colIndex(col) + 2)}${R_notInvited}`).join('+');

      // Percentage formula
      matrix[rIdx][pctCol] = `=SUM(${colAE}${R_invited}:${colAG}${R_invited})/SUM(${colAE}${R_invited}:${colAG}${R_notInvited})`;
      matrix[rIdx + 1][pctCol] = `=SUM(${colAE}${R_notInvited}:${colAG}${R_notInvited})/SUM(${colAE}${R_invited}:${colAG}${R_notInvited})`;
    });

    // Row 27 & 28 Totals
    matrix[26][7] = '總計';
    matrix[26][8] = '受邀';
    matrix[27][8] = '非受邀';

    const invitedRows = [5, 7, 9, 11, 13, 15, 17, 19, 21, 23, 25];
    const notInvitedRows = [6, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26];

    for (let c = 9; c <= grandTotalStart + 2; c++) {
      const col = columnName(c);
      matrix[26][c] = '=' + invitedRows.map(r => `${col}${r}`).join('+');
      matrix[27][c] = '=' + notInvitedRows.map(r => `${col}${r}`).join('+');
    }

    // Row 29 & 30 Percentages for Column Groups
    matrix[28][7] = '%';
    matrix[28][8] = '受邀';
    matrix[29][8] = '非受邀';

    columnGroups.forEach((group, groupIdx) => {
      const startCol = 9 + 3 * groupIdx;
      const colA = columnName(startCol);
      const colB = columnName(startCol + 1);
      const colC = columnName(startCol + 2);
      
      for (let c = startCol; c <= startCol + 2; c++) {
        const col = columnName(c);
        matrix[28][c] = `=${col}27/SUM(${colA}27:${colC}28)`;
        matrix[29][c] = `=${col}28/SUM(${colA}27:${colC}28)`;
      }
    });

    // Percentages for Grand Total column group
    const colAE = columnName(grandTotalStart);
    const colAF = columnName(grandTotalStart + 1);
    const colAG = columnName(grandTotalStart + 2);
    for (let c = grandTotalStart; c <= grandTotalStart + 2; c++) {
      const col = columnName(c);
      matrix[28][c] = `=${col}27/SUM(${colAE}27:${colAG}28)`;
      matrix[29][c] = `=${col}28/SUM(${colAE}27:${colAG}28)`;
    }

    // Row 32 (index 31): 主日禮拜人數
    matrix[31][7] = '主日禮拜人數';
    for (let c = 9; c <= grandTotalStart + 2; c++) {
      matrix[31][c] = '-';
    }
    matrix[31][pctCol] = '-';

    // Sheet 2: Detail table
    const detailHeaders = ['新家人姓名', '參加的聚會是', '表單號', '關懷同工', '關懷狀態', '落戶狀態', '邀約人', '立案日', '結案日', '家長備註欄'];
    const detailRows = [
      detailHeaders,
      ...rows.map(item => [
        item['新家人姓名'] || '',
        item['參加的聚會是'] || '',
        item['表單號'] ? Number(item['表單號']) : '',
        item['關懷同工'] || '',
        '結案',
        getDisplaySettlementStatus(item) || '',
        item['邀約人'] || '',
        item['日期'] || '',
        item['日期'] || '',
        item['備註'] || ''
      ])
    ];

    const sheets = [
      { name: '新家人落戶分析', rows: matrix },
      { name: '最新新家人名單&落戶狀態', rows: detailRows }
    ];

    exportWorkbook(sheets, `新家人留名卡紀錄_截至${(analysisEndDate.value || '').replace(/-/g, '')}`);
  } catch (error) {
    setNotice(analysisNotice, error.message || String(error), 'error');
  }
}

function exportWorkbook(sheets, filenameBase) {
  const blob = createXlsxBlob(sheets);
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `${sanitizeFilename(filenameBase)}.xlsx`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function getAnalysisRangeLabel() {
  return `${analysisStartDate.value || 'all'}_${analysisEndDate.value || 'all'}`;
}

function sanitizeFilename(value) {
  return String(value || 'export')
    .replace(/[\\/:*?"<>|]/g, '-')
    .replace(/\s+/g, '_');
}

function createXlsxBlob(sheets) {
  const files = {};
  const safeSheets = sheets.map((sheet, index) => ({
    name: sanitizeSheetName(sheet.name || `Sheet${index + 1}`),
    rows: rowsToMatrix(sheet.rows || [])
  }));

  files['[Content_Types].xml'] = buildContentTypesXml(safeSheets.length);
  files['_rels/.rels'] = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
    '</Relationships>'
  ].join('');
  files['xl/workbook.xml'] = buildWorkbookXml(safeSheets);
  files['xl/_rels/workbook.xml.rels'] = buildWorkbookRelsXml(safeSheets.length);

  safeSheets.forEach((sheet, index) => {
    files[`xl/worksheets/sheet${index + 1}.xml`] = buildWorksheetXml(sheet.rows);
  });

  return new Blob([buildZip(files)], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
}

function rowsToMatrix(rows) {
  if (!rows.length) return [[]];
  if (Array.isArray(rows[0])) return rows;
  const headers = Object.keys(rows[0]);
  return [
    headers,
    ...rows.map(row => headers.map(header => row[header] === undefined || row[header] === null ? '' : row[header]))
  ];
}

function buildContentTypesXml(sheetCount) {
  const sheetOverrides = Array.from({ length: sheetCount }, (_, index) =>
    `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
  ).join('');
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
    sheetOverrides,
    '</Types>'
  ].join('');
}

function buildWorkbookXml(sheets) {
  const sheetTags = sheets.map((sheet, index) =>
    `<sheet name="${escapeXmlAttribute(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`
  ).join('');
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    `<sheets>${sheetTags}</sheets>`,
    '</workbook>'
  ].join('');
}

function buildWorkbookRelsXml(sheetCount) {
  const relations = Array.from({ length: sheetCount }, (_, index) =>
    `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`
  ).join('');
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    relations,
    '</Relationships>'
  ].join('');
}

function buildWorksheetXml(rows) {
  const rowXml = rows.map((row, rowIndex) => {
    const cellXml = row.map((value, columnIndex) => buildCellXml(value, columnIndex, rowIndex)).join('');
    return `<row r="${rowIndex + 1}">${cellXml}</row>`;
  }).join('');
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    `<sheetData>${rowXml}</sheetData>`,
    '</worksheet>'
  ].join('');
}

function buildCellXml(value, columnIndex, rowIndex) {
  const reference = `${columnName(columnIndex)}${rowIndex + 1}`;
  if (typeof value === 'string' && value.startsWith('=')) {
    return `<c r="${reference}"><f>${escapeXmlText(value.slice(1))}</f></c>`;
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    return `<c r="${reference}"><v>${value}</v></c>`;
  }
  return `<c r="${reference}" t="inlineStr"><is><t>${escapeXmlText(value)}</t></is></c>`;
}

function columnName(index) {
  let name = '';
  let value = index + 1;
  while (value > 0) {
    const remainder = (value - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

function sanitizeSheetName(value) {
  return String(value || 'Sheet')
    .replace(/[:\\/?*\[\]]/g, ' ')
    .slice(0, 31)
    .trim() || 'Sheet';
}

function escapeXmlText(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function escapeXmlAttribute(value) {
  return escapeXmlText(value).replace(/"/g, '&quot;');
}

function buildZip(files) {
  const encoder = new TextEncoder();
  const chunks = [];
  const centralDirectory = [];
  let offset = 0;
  const entries = Object.entries(files).map(([name, content]) => ({
    name,
    nameBytes: encoder.encode(name),
    data: encoder.encode(content)
  }));

  entries.forEach(entry => {
    const crc = crc32(entry.data);
    const localHeader = zipLocalHeader(entry, crc);
    chunks.push(localHeader, entry.nameBytes, entry.data);
    centralDirectory.push({ entry, crc, offset });
    offset += localHeader.length + entry.nameBytes.length + entry.data.length;
  });

  const centralStart = offset;
  centralDirectory.forEach(item => {
    const centralHeader = zipCentralHeader(item.entry, item.crc, item.offset);
    chunks.push(centralHeader, item.entry.nameBytes);
    offset += centralHeader.length + item.entry.nameBytes.length;
  });

  const centralSize = offset - centralStart;
  chunks.push(zipEndRecord(entries.length, centralSize, centralStart));
  return concatUint8Arrays(chunks);
}

function zipLocalHeader(entry, crc) {
  const buffer = new ArrayBuffer(30);
  const view = new DataView(buffer);
  view.setUint32(0, 0x04034b50, true);
  view.setUint16(4, 20, true);
  view.setUint16(6, 0, true);
  view.setUint16(8, 0, true);
  view.setUint16(10, 0, true);
  view.setUint16(12, 0, true);
  view.setUint32(14, crc, true);
  view.setUint32(18, entry.data.length, true);
  view.setUint32(22, entry.data.length, true);
  view.setUint16(26, entry.nameBytes.length, true);
  view.setUint16(28, 0, true);
  return new Uint8Array(buffer);
}

function zipCentralHeader(entry, crc, localOffset) {
  const buffer = new ArrayBuffer(46);
  const view = new DataView(buffer);
  view.setUint32(0, 0x02014b50, true);
  view.setUint16(4, 20, true);
  view.setUint16(6, 20, true);
  view.setUint16(8, 0, true);
  view.setUint16(10, 0, true);
  view.setUint16(12, 0, true);
  view.setUint16(14, 0, true);
  view.setUint32(16, crc, true);
  view.setUint32(20, entry.data.length, true);
  view.setUint32(24, entry.data.length, true);
  view.setUint16(28, entry.nameBytes.length, true);
  view.setUint16(30, 0, true);
  view.setUint16(32, 0, true);
  view.setUint16(34, 0, true);
  view.setUint16(36, 0, true);
  view.setUint32(38, 0, true);
  view.setUint32(42, localOffset, true);
  return new Uint8Array(buffer);
}

function zipEndRecord(entryCount, centralSize, centralStart) {
  const buffer = new ArrayBuffer(22);
  const view = new DataView(buffer);
  view.setUint32(0, 0x06054b50, true);
  view.setUint16(4, 0, true);
  view.setUint16(6, 0, true);
  view.setUint16(8, entryCount, true);
  view.setUint16(10, entryCount, true);
  view.setUint32(12, centralSize, true);
  view.setUint32(16, centralStart, true);
  view.setUint16(20, 0, true);
  return new Uint8Array(buffer);
}

function concatUint8Arrays(parts) {
  const total = parts.reduce((sum, part) => sum + part.length, 0);
  const output = new Uint8Array(total);
  let offset = 0;
  parts.forEach(part => {
    output.set(part, offset);
    offset += part.length;
  });
  return output;
}

function crc32(data) {
  let crc = -1;
  for (let index = 0; index < data.length; index++) {
    crc = (crc >>> 8) ^ crc32Table[(crc ^ data[index]) & 0xff];
  }
  return (crc ^ -1) >>> 0;
}

const crc32Table = (() => {
  const table = new Uint32Array(256);
  for (let index = 0; index < 256; index++) {
    let value = index;
    for (let bit = 0; bit < 8; bit++) {
      value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
    }
    table[index] = value >>> 0;
  }
  return table;
})();

async function getAnalysisDateRows() {
  const result = await callCachedListApi('getClosedCases');
  const filters = {
    startDate: analysisStartDate.value,
    endDate: analysisEndDate.value
  };
  return enrichRowsWithSundayMemberData(filterCases(result.data || [], filters));
}

function filterAnalysisRowsByStatus(rows) {
  const status = analysisStatusFilter.value;
  return status
    ? rows.filter(item => getDisplaySettlementStatus(item) === status)
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
  const statuses = Array.from(new Set(rows.map(item => getDisplaySettlementStatus(item)))).sort();
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
    const status = getDisplaySettlementStatus(item);
    counts.set(status, (counts.get(status) || 0) + 1);
  });
  return Array.from(counts.entries())
    .map(([status, count]) => ({ status, count }))
    .sort((a, b) => b.count - a.count || a.status.localeCompare(b.status));
}

function normalizeSettlementStatus(value) {
  return String(value || '').trim() || '未填落戶狀態';
}

function getDisplaySettlementStatus(item) {
  return item && item.displaySettlementStatus
    ? item.displaySettlementStatus
    : normalizeSettlementStatus(item && item['落戶狀態']);
}

function normalizeBoolean(value) {
  if (value === true) return true;
  const text = String(value || '').trim().toLowerCase();
  return ['true', 'yes', 'y', '1', '是'].includes(text);
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
  table.innerHTML = `
    <thead><tr><th>姓名</th><th>日期</th><th>落戶狀態</th><th>點名系統代碼</th><th>現行小組</th></tr></thead>
    <tbody></tbody>
  `;
  const tbody = table.querySelector('tbody');
  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="5">沒有符合範圍的明細</td></tr>';
    return table;
  }

  rows.forEach(item => {
    const row = document.createElement('tr');
    ['新家人姓名', '日期', '落戶狀態', '點名系統代碼'].forEach(column => {
      const cell = document.createElement('td');
      cell.textContent = column === '落戶狀態'
        ? getDisplaySettlementStatus(item)
        : item[column] || '';
      row.appendChild(cell);
    });
    const groupCell = document.createElement('td');
    groupCell.appendChild(buildSundayGroupTag(item['主日點名小組'] || ''));
    row.appendChild(groupCell);
    tbody.appendChild(row);
  });
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

  headRow.innerHTML = `${selectable ? '<th class="check-cell">結案</th><th class="action-cell"></th>' : ''}${columns.map(column => `<th>${escapeHtml(getColumnLabel(column))}</th>`).join('')}`;
  thead.appendChild(headRow);

  rows.forEach(item => {
    const row = document.createElement('tr');

    if (selectable) {
      const checkboxCell = document.createElement('td');
      checkboxCell.className = 'check-cell';
      checkboxCell.innerHTML = `<input type="checkbox" value="${item.rowNumber}" aria-label="勾選 ${escapeHtml(item['新家人姓名'] || '此筆資料')} 結案">`;
      row.appendChild(checkboxCell);

      const actionCell = document.createElement('td');
      actionCell.className = 'action-cell';
      
      const dropdown = document.createElement('div');
      dropdown.className = 'action-dropdown';

      const toggleButton = document.createElement('button');
      toggleButton.type = 'button';
      toggleButton.className = 'btn secondary action-toggle-btn';
      toggleButton.textContent = '操作';
      dropdown.appendChild(toggleButton);

      const menu = document.createElement('div');
      menu.className = 'action-menu';
      menu.hidden = true;

      const editItem = document.createElement('button');
      editItem.type = 'button';
      editItem.className = 'menu-item';
      editItem.textContent = '編輯';
      editItem.addEventListener('click', () => {
        menu.hidden = true;
        openEditModal(item);
      });
      menu.appendChild(editItem);

      const deleteItem = document.createElement('button');
      deleteItem.type = 'button';
      deleteItem.className = 'menu-item danger';
      deleteItem.textContent = '刪除';
      deleteItem.addEventListener('click', () => {
        menu.hidden = true;
        deleteSingleCase(item);
      });
      menu.appendChild(deleteItem);

      dropdown.appendChild(menu);
      actionCell.appendChild(dropdown);
      row.appendChild(actionCell);

      toggleButton.addEventListener('click', event => {
        event.stopPropagation();
        document.querySelectorAll('.action-menu').forEach(m => {
          if (m !== menu) m.hidden = true;
        });
        menu.hidden = !menu.hidden;
      });
    }

    columns.forEach(column => {
      const cell = document.createElement('td');
      if (column === '主日點名小組') {
        cell.appendChild(buildSundayGroupTag(item[column] || ''));
      } else if (column === '落戶狀態') {
        cell.textContent = getDisplaySettlementStatus(item);
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

async function deleteSingleCase(item) {
  const name = item['新家人姓名'] || '此筆資料';
  if (!confirm(`確認要永久刪除新朋友「${name}」的追蹤資料嗎？此操作將無法復原。`)) {
    return;
  }

  setNotice(trackingNotice, '刪除中...');
  
  try {
    const result = await callApi('deleteTrackingCase', { rowNumber: item.rowNumber });
    setNotice(trackingNotice, result.message, 'success');
    await loadTrackingCases();
  } catch (error) {
    setNotice(trackingNotice, error.message || String(error), 'error');
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
