/* script.js - 敬拜團服事管理系統 (外部框架驅動 + 列專屬請假版) */

// --- 全域變數 ---
let currentPositions = [];
let generatedScheduleData = [];
let uniquePersonnel = [];
let sortablePositions = null;
let _worshipTeamCache = null; // 敬拜團員名單快取（供「位置與同工」下拉使用）
let loadedDashboardData = []; // 供團員查詢服事天數使用

// ==========================================
// 🔍 共用：可搜尋浮動下拉選單元件
// ==========================================
function _hideFloatingDropdown() {
  const dd = document.getElementById('_floatingDropdown');
  if (dd) dd.remove();
  document.removeEventListener('mousedown', _floatingDropdownOutsideClick, { capture: true });
}
function _floatingDropdownOutsideClick(e) {
  const dd = document.getElementById('_floatingDropdown');
  if (dd && !dd.contains(e.target)) _hideFloatingDropdown();
}

/**
 * 顯示一個可搜尋的浮動下拉清單
 * @param {HTMLElement} anchorEl - 錨點元素（決定位置）
 * @param {Array<{label:string, subLabel?:string, value:any, disabled?:boolean}>} items
 * @param {Function} onPick - 回呼 (item) => void
 * @param {Object} [opts] - { placeholder, emptyText, width }
 */
function _showFloatingDropdown(anchorEl, items, onPick, opts = {}) {
  _hideFloatingDropdown();
  const rect = anchorEl.getBoundingClientRect();
  const width = opts.width || Math.max(rect.width, 240);

  const dd = document.createElement('div');
  dd.id = '_floatingDropdown';
  dd.className = 'shadow-lg border rounded bg-white';
  dd.style.cssText = `position:fixed; top:${rect.bottom + 4}px; left:${rect.left}px;
                      width:${width}px; max-height:300px; overflow:hidden;
                      z-index:2050; display:flex; flex-direction:column;`;

  // 搜尋框
  const searchWrap = document.createElement('div');
  searchWrap.className = 'p-2 border-bottom bg-light';
  const search = document.createElement('input');
  search.type = 'text';
  search.className = 'form-control form-control-sm';
  search.placeholder = opts.placeholder || '🔍 輸入關鍵字搜尋...';
  searchWrap.appendChild(search);
  dd.appendChild(searchWrap);

  // 清單區
  const list = document.createElement('div');
  list.style.cssText = 'overflow-y:auto; flex:1;';
  dd.appendChild(list);

  function render(kw = '') {
    const f = (kw || '').toLowerCase().trim();
    const filtered = items.filter(it =>
      !f || (it.label || '').toLowerCase().includes(f) ||
      (it.subLabel || '').toLowerCase().includes(f)
    );
    if (filtered.length === 0) {
      list.innerHTML = `<div class="p-3 text-center text-muted small">${opts.emptyText || '查無相符'}</div>`;
      return;
    }
    list.innerHTML = filtered.map(it => {
      const i = items.indexOf(it);
      const disabledCls = it.disabled ? 'text-muted' : '';
      const cursor = it.disabled ? 'not-allowed' : 'pointer';
      return `<div class="ss-item px-3 py-2 border-bottom d-flex justify-content-between align-items-center ${disabledCls}"
                style="cursor:${cursor};" data-i="${i}">
        <span><strong>${it.label}</strong></span>
        ${it.subLabel ? `<small class="text-muted">${it.subLabel}</small>` : ''}
      </div>`;
    }).join('');
    list.querySelectorAll('.ss-item').forEach(el => {
      const it = items[parseInt(el.dataset.i)];
      if (it.disabled) return;
      el.onmouseenter = () => el.style.backgroundColor = '#e7f1ff';
      el.onmouseleave = () => el.style.backgroundColor = '';
      el.onmousedown = (e) => {
        e.preventDefault();
        e.stopPropagation();
        onPick(it);
        _hideFloatingDropdown();
      };
    });
  }

  search.addEventListener('input', () => render(search.value));
  render();
  document.body.appendChild(dd);
  setTimeout(() => search.focus(), 30);

  // 外部點擊關閉 (不再監聽 scroll/resize，防止任何行動裝置或虛擬鍵盤彈起導致選單被收合)
  setTimeout(() => {
    document.addEventListener('mousedown', _floatingDropdownOutsideClick, { capture: true });
  }, 0);
}

// 確保敬拜團員名單已載入（快取）
async function ensureWorshipTeamLoaded(force = false) {
  if (_worshipTeamCache && !force) return _worshipTeamCache;
  try {
    const res = await callAPI('getTeamMembers');
    _worshipTeamCache = (res && res.data) ? res.data : [];
  } catch (e) {
    console.warn('載入敬拜團員名單失敗', e);
    _worshipTeamCache = [];
  }
  return _worshipTeamCache;
}

// 主日會友下拉（敬拜團員名單分頁用）
async function openMainMemberDropdown(anchorEl) {
  // 1. 若快取為空 → 自動嘗試重新抓（包含第一次點擊就出現的情境）
  if (!_mainMemberSuggestionsCache || _mainMemberSuggestionsCache.length === 0) {
    // 先在浮動下拉顯示 loading（並非 alert，不會死循環）
    _showFloatingDropdown(anchorEl, [{ label: '⏳ 正在抓主日會友清單...', subLabel: '', value: null, disabled: true }],
      () => {}, { placeholder: '載入中...', emptyText: '載入中...' });
    try {
      const sugRes = await callAPI('getMemberSuggestions');
      // GAS 端 throw 會走到 doPost catch → {status:'error', message}
      if (sugRes && sugRes.status === 'error') {
        throw new Error(sugRes.message || 'GAS 回傳錯誤');
      }
      _mainMemberSuggestionsCache = (sugRes && sugRes.data) ? sugRes.data : [];
      buildTeamMemberDatalist();
    } catch (err) {
      _showFloatingDropdown(anchorEl, [
        { label: `❌ 載入失敗：${err.message || err}`, subLabel: '', value: null, disabled: true },
        { label: '👉 請按右方「重試」', subLabel: '', value: '__retry__' }
      ], (it) => { if (it.value === '__retry__') openMainMemberDropdown(anchorEl); },
        { placeholder: '載入失敗' });
      return;
    }
  }

  // 2. 真的空（GAS 回傳空陣列）→ 顯示在 dropdown 內，附「重試」項
  if (!_mainMemberSuggestionsCache || _mainMemberSuggestionsCache.length === 0) {
    _showFloatingDropdown(anchorEl, [
      { label: '⚠️ 主日會友清單為空', subLabel: 'GAS 可能讀不到主日測試試算表', value: null, disabled: true },
      { label: '🔄 重新嘗試', subLabel: '', value: '__retry__' }
    ], (it) => {
      if (it.value === '__retry__') {
        _mainMemberSuggestionsCache = null;
        openMainMemberDropdown(anchorEl);
      }
    }, { placeholder: '無資料' });
    return;
  }

  // 3. 過濾已加入的人 → 渲染
  const existingUids = new Set(_editingTeamMembers.map(m => m.uid).filter(x => x));
  const nameCount = {};
  _mainMemberSuggestionsCache.forEach(m => {
    if (!existingUids.has(m.uid)) nameCount[m.name] = (nameCount[m.name] || 0) + 1;
  });
  const available = _mainMemberSuggestionsCache.filter(m => !existingUids.has(m.uid));
  if (available.length === 0) {
    _showFloatingDropdown(anchorEl, [
      { label: 'ℹ️ 所有主日會友都已在敬拜團員名單中', subLabel: '', value: null, disabled: true }
    ], () => {}, { placeholder: '都加入了' });
    return;
  }
  const items = available
    .sort((a, b) => (a.name || '').localeCompare(b.name || ''))
    .map(m => {
      const label = nameCount[m.name] > 1 ? `${m.name} (${m.uid})` : m.name;
      return { label: m.name, subLabel: m.uid || '', value: label };
    });
  _showFloatingDropdown(anchorEl, items, (it) => {
    const input = document.getElementById('newTeamMemberInput');
    input.value = it.value;
    input.focus();
  }, { placeholder: '🔍 輸入姓名或編號搜尋...', emptyText: '查無此會友' });
}

// --- 初始化 ---
window.onload = () => {
  const syncTimeEl = document.getElementById('syncTime');
  if (syncTimeEl) syncTimeEl.innerText = new Date().toLocaleTimeString();

  // 根據當前日期設定預設季度
  const today = new Date();
  const currentYear = today.getFullYear();
  const currentMonth = today.getMonth(); // 0-11
  const currentQuarter = 'Q' + (Math.floor(currentMonth / 3) + 1);

  const yearSelect = document.getElementById('yearSelect');
  const quarterSelect = document.getElementById('quarterSelect');
  if (yearSelect && Array.from(yearSelect.options).some(opt => opt.value === String(currentYear))) {
    yearSelect.value = String(currentYear);
  }
  if (quarterSelect && Array.from(quarterSelect.options).some(opt => opt.value === currentQuarter)) {
    quarterSelect.value = currentQuarter;
  }

  loadDashboard();
};

function formatDateSafe(dateObj) {
  if (!dateObj || isNaN(dateObj.getTime())) return "";
  const y = dateObj.getFullYear();
  const m = String(dateObj.getMonth() + 1).padStart(2, '0');
  const d = String(dateObj.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function parseDateSafe(dateStr) {
  if (!dateStr) return new Date();
  const parts = dateStr.split('-');
  if (parts.length !== 3) return new Date();
  return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
}

// --- 🌟 安全網橋接設定 ---
async function callAPI(action, payload) {
  try {
    if (typeof window.churchAPI !== 'function') {
      throw new Error("中央安全設定檔 (config.js) 尚未載入，請檢查 HTML 引用路徑。");
    }
    return await window.churchAPI(action, payload);
  } catch (error) {
    console.error("🚫 API 通訊失敗:", error);
    throw error;
  }
}

async function ensurePositionsLoaded() {
  if (currentPositions && currentPositions.length > 0) return;
  const result = await callAPI('getPositions', {});
  if (result && result.status === 'success') {
    currentPositions = result.data || [];
    // Populate uniquePersonnel from currentPositions
    let nameSet = new Set();
    currentPositions.forEach(pos => (pos.personnel || '').split(',').forEach(n => n.trim() && nameSet.add(n.trim())));
    uniquePersonnel = Array.from(nameSet).sort();
  }
}

function switchTab(tabId) {
  const content = document.getElementById(tabId);
  if (!content) return;
  document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.nav-link').forEach(el => el.classList.remove('active'));
  content.classList.add('active');
  const activeLink = document.querySelector(`a[onclick="switchTab('${tabId}')"]`);
  if (activeLink) activeLink.classList.add('active');

  if(tabId === 'dashboard') loadDashboard();
  if(tabId === 'settings') loadPositions();
  if(tabId === 'schedule') initScheduleTab();
  if(tabId === 'teamMembers') loadTeamMembers();
  if(tabId === 'calLink') loadCalLinkSettings();
}

// ==========================================
// 📅 行事曆連結設定
// ==========================================
let _calLinkConfig = null;
let _scheduleDates = null; // [{date, name, type, year, quarter}, ...]

async function loadCalLinkSettings() {
  document.getElementById('calLinkLoadingArea').style.display = '';
  document.getElementById('calLinkMainArea').style.display = 'none';
  try {
    const [cfgRes, datesRes] = await Promise.all([
      callAPI('getCalendarLinkConfig', {}),
      callAPI('getScheduleDates', {})
    ]);
    if (cfgRes.status !== 'success') throw new Error(cfgRes.message || '載入連結設定失敗');
    if (datesRes.status !== 'success') throw new Error(datesRes.message || '載入服事表日期失敗');
    _calLinkConfig = cfgRes.data;
    _scheduleDates = datesRes.data || [];
    renderCalLinkUI();
    renderScheduleDateList();
  } catch (err) {
    document.getElementById('calLinkLoadingArea').innerHTML =
      `<div class="alert alert-danger">❌ 載入失敗：${err.message}</div>`;
  }
}

function renderCalLinkUI() {
  if (!_calLinkConfig) return;
  document.getElementById('calLinkLoadingArea').style.display = 'none';
  document.getElementById('calLinkMainArea').style.display = '';

  const subTypes = _calLinkConfig.sermonSubTypes || [];

  // 預設子類型下拉
  const sel = document.getElementById('defaultSermonSubType');
  if (!_calLinkConfig.calendarReachable) {
    sel.innerHTML = '<option value="">⚠️ 讀不到行事曆資料（請確認跨 SS 授權）</option>';
    sel.disabled = true;
  } else if (subTypes.length === 0) {
    sel.innerHTML = '<option value="">⚠️ 行事曆中還沒有「講道資訊」的子類型，請先到行事曆建立</option>';
    sel.disabled = true;
  } else {
    sel.innerHTML = '<option value="">-- 不指定 --</option>' +
      subTypes.map(t => `<option value="${t.typeId}" ${t.typeId === _calLinkConfig.defaultSermonSubTypeId ? 'selected' : ''}>${t.icon || ''} ${t.name}</option>`).join('');
    sel.disabled = false;
  }

  // 狀態顯示
  const statusEl = document.getElementById('defaultSubTypeStatus');
  if (_calLinkConfig.defaultSermonSubTypeId && !_calLinkConfig.defaultIsValid) {
    statusEl.innerHTML = '<span class="text-danger">⚠️ 目前儲存的預設子類型在行事曆中找不到（可能已被刪除），請重選</span>';
  } else if (_calLinkConfig.defaultSermonSubTypeId) {
    statusEl.innerHTML = '<span class="text-success">✓ 已設定</span>';
  } else {
    statusEl.innerText = '尚未設定，沒設的話公佈欄的講道欄位會空白';
  }
}

// 渲染服事表日期清單（取代舊「日期覆寫清單」）
function renderScheduleDateList() {
  const container = document.getElementById('scheduleDateListContainer');
  if (!container) return;

  if (!_scheduleDates || _scheduleDates.length === 0) {
    container.innerHTML = '<div class="alert alert-light text-muted text-center m-2">服事表還沒有任何日期</div>';
    return;
  }

  const subTypes = (_calLinkConfig && _calLinkConfig.sermonSubTypes) || [];
  const overrides = (_calLinkConfig && _calLinkConfig.overrides) || {};
  const defaultId = (_calLinkConfig && _calLinkConfig.defaultSermonSubTypeId) || '';

  if (subTypes.length === 0) {
    container.innerHTML = '<div class="alert alert-warning m-2">⚠️ 行事曆中沒有「講道資訊」子類型，無法做覆寫對應</div>';
    return;
  }

  // 篩選
  const mode = document.getElementById('dateFilterSelect').value;
  let rows = _scheduleDates;
  if (mode === 'overrides') {
    rows = rows.filter(r => !!overrides[r.date]);
  }
  if (rows.length === 0) {
    container.innerHTML = '<div class="alert alert-light text-muted text-center m-2">無符合條件的日期</div>';
    return;
  }

  // 表格
  let html = `<table class="table table-sm table-bordered align-middle mb-0">
    <thead class="table-light text-center sticky-top" style="top:0; z-index:5;">
      <tr>
        <th style="width: 14%;">日期</th>
        <th style="width: 20%;">聚會名稱</th>
        <th style="width: 12%;">類別</th>
        <th style="width: 14%;">實際採用</th>
        <th style="width: 32%;">覆寫子類型</th>
        <th style="width: 8%;">狀態</th>
      </tr>
    </thead><tbody>`;
  rows.forEach(r => {
    const ov = overrides[r.date] || '';
    const effective = ov || defaultId;
    const effObj = subTypes.find(t => t.typeId === effective);
    const isOverridden = !!ov;
    const stateBadge = isOverridden
      ? `<span class="badge bg-warning text-dark">覆寫</span>`
      : (defaultId
          ? `<span class="badge bg-secondary">預設</span>`
          : `<span class="badge bg-danger">無</span>`);

    let effectiveLabel;
    if (effObj) effectiveLabel = `${effObj.icon || ''} ${effObj.name}`;
    else if (effective) effectiveLabel = `<span class="text-danger">查無</span>`;
    else effectiveLabel = `<span class="text-muted">－</span>`;

    const selOpts = `<option value="">使用預設</option>` +
      subTypes.map(t => `<option value="${t.typeId}" ${t.typeId === ov ? 'selected' : ''}>${t.icon || ''} ${t.name}</option>`).join('');

    html += `<tr ${isOverridden ? 'class="table-warning"' : ''}>
      <td class="text-center fw-bold"><small>${r.date}</small></td>
      <td><small>${escapeHtmlSafe(r.name || '-')}</small></td>
      <td class="text-center"><small class="text-muted">${escapeHtmlSafe(r.type || '-')}</small></td>
      <td class="text-center"><small>${effectiveLabel}</small></td>
      <td>
        <select class="form-select form-select-sm" onchange="changeDateOverride('${r.date}', this.value, this)">
          ${selOpts}
        </select>
      </td>
      <td class="text-center">${stateBadge}</td>
    </tr>`;
  });
  html += '</tbody></table>';
  container.innerHTML = html;
}

function escapeHtmlSafe(s) {
  return String(s == null ? '' : s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

// 表格內下拉改變時觸發
async function changeDateOverride(date, typeId, selectEl) {
  // 先樂觀更新（UI 不會 reload 整個表）
  const old = (_calLinkConfig.overrides && _calLinkConfig.overrides[date]) || '';
  if (selectEl) selectEl.disabled = true;
  try {
    const res = await callAPI('setDateOverride', { date, typeId });
    if (res.status !== 'success') throw new Error(res.message);
    if (!_calLinkConfig.overrides) _calLinkConfig.overrides = {};
    if (typeId) _calLinkConfig.overrides[date] = typeId;
    else delete _calLinkConfig.overrides[date];
    // 只重渲染清單（保留下拉狀態）
    renderScheduleDateList();
  } catch (err) {
    alert('❌ ' + err.message);
    if (selectEl) selectEl.value = old; // 回復
  } finally {
    if (selectEl) selectEl.disabled = false;
  }
}

async function saveDefaultSubType() {
  const typeId = document.getElementById('defaultSermonSubType').value;
  try {
    const res = await callAPI('setDefaultSermonSubType', { typeId });
    if (res.status !== 'success') throw new Error(res.message);
    if (_calLinkConfig) _calLinkConfig.defaultSermonSubTypeId = typeId;
    renderCalLinkUI();
    renderScheduleDateList(); // 預設變了，每列「實際採用」也要更新
  } catch (err) {
    alert('❌ ' + err.message);
  }
}

async function reloadScheduleDates() {
  try {
    const res = await callAPI('getScheduleDates', {});
    if (res.status !== 'success') throw new Error(res.message);
    _scheduleDates = res.data || [];
    renderScheduleDateList();
  } catch (err) {
    alert('❌ ' + err.message);
  }
}

async function clearCalLinkCacheBtn() {
  try {
    // 1. 清 GAS 端的跨 SS 行事曆 cache（CacheService）
    const res = await callAPI('clearCalendarLinkCache', {});
    if (res.status !== 'success') throw new Error(res.message);
    // 2. 連帶清 Firebase 的 getSchedule / getScheduleByDateRange（避免舊資料卡住）
    if (typeof window.churchAPIInvalidate === 'function') {
      await window.churchAPIInvalidate('getSchedule');
      await window.churchAPIInvalidate('getScheduleByDateRange');
    }
    alert('✅ 已清空所有快取（GAS + Firebase），重新載入中...');
    await loadCalLinkSettings();
  } catch (err) {
    alert('❌ ' + err.message);
  }
}

// ==========================================
// 👥 敬拜團員名單管理
// ==========================================
let _editingTeamMembers = [];
let _mainMemberSuggestionsCache = null;

async function loadTeamMembers() {
  const tbody = document.getElementById('teamMembersTbody');
  tbody.innerHTML = '<tr><td colspan="4" class="text-center text-muted p-4"><div class="spinner-border spinner-border-sm me-2"></div>載入中...</td></tr>';

  try {
    const [sugRes, teamRes] = await Promise.all([
      callAPI('getMemberSuggestions'),
      callAPI('getTeamMembers')
    ]);
    if (sugRes && sugRes.status === 'error') {
      console.warn('[loadTeamMembers] getMemberSuggestions error：', sugRes.message);
      _mainMemberSuggestionsCache = []; // 仍可繼續，使用者按 ▼ 時會重抓並顯示錯誤
    } else {
      _mainMemberSuggestionsCache = (sugRes && sugRes.data) ? sugRes.data : [];
    }
    _editingTeamMembers = ((teamRes && teamRes.data) ? teamRes.data : []).map(m => ({ ...m }));

    buildTeamMemberDatalist();
    renderTeamMembersTable();
  } catch (err) {
    tbody.innerHTML = `<tr><td colspan="4" class="text-center text-danger p-4">❌ 載入失敗：${err.message}</td></tr>`;
  }
}

function buildTeamMemberDatalist() {
  const datalist = document.getElementById('teamMemberSuggestions');
  if (!datalist || !_mainMemberSuggestionsCache) return;

  const existingUids = new Set(_editingTeamMembers.map(m => m.uid));
  // 算同名次數，決定 datalist 顯示格式
  const nameCount = {};
  _mainMemberSuggestionsCache.forEach(m => {
    if (!existingUids.has(m.uid)) {
      nameCount[m.name] = (nameCount[m.name] || 0) + 1;
    }
  });

  datalist.innerHTML = _mainMemberSuggestionsCache
    .filter(m => !existingUids.has(m.uid))
    .map(m => {
      const label = nameCount[m.name] > 1 ? `${m.name} (${m.uid})` : m.name;
      return `<option value="${label}"></option>`;
    }).join('');
}

function renderTeamMembersTable() {
  const tbody = document.getElementById('teamMembersTbody');
  if (_editingTeamMembers.length === 0) {
    tbody.innerHTML = '<tr><td colspan="4" class="text-center text-muted p-4">尚無團員，請從上方加入</td></tr>';
    _updateTeamMembersCount();
    return;
  }

  // 排序：正式在前、姓名次之
  const sorted = _editingTeamMembers.slice().sort((a, b) => {
    if (a.status !== b.status) return a.status === '正式' ? -1 : 1;
    return (a.name || '').localeCompare(b.name || '');
  });

  tbody.innerHTML = sorted.map(m => {
    const idx = _editingTeamMembers.indexOf(m);
    const statusBadge = m.status === '實習'
      ? '<span class="badge bg-warning text-dark">🎓 實習</span>'
      : '<span class="badge bg-primary">⭐ 正式</span>';
    return `
      <tr>
        <td><strong>${m.name}</strong></td>
        <td><small class="text-muted font-monospace">${m.uid || '-'}</small></td>
        <td class="text-center">
          ${statusBadge}
          <select class="form-select form-select-sm d-inline-block ms-2" style="width: 110px;"
                  onchange="updateTeamMemberStatus(${idx}, this.value)">
            <option value="正式" ${m.status === '正式' ? 'selected' : ''}>正式</option>
            <option value="實習" ${m.status === '實習' ? 'selected' : ''}>實習</option>
          </select>
        </td>
        <td class="text-center">
          <button class="btn btn-sm btn-outline-danger" onclick="removeTeamMember(${idx})">移除</button>
        </td>
      </tr>
    `;
  }).join('');

  _updateTeamMembersCount();
}

function _updateTeamMembersCount() {
  const total = _editingTeamMembers.length;
  const formal = _editingTeamMembers.filter(m => m.status === '正式').length;
  const intern = _editingTeamMembers.filter(m => m.status === '實習').length;
  const totalEl  = document.getElementById('teamMembersCount');
  const formalEl = document.getElementById('teamMembersFormalCount');
  const internEl = document.getElementById('teamMembersInternCount');
  if (totalEl)  totalEl.innerText = total;
  if (formalEl) formalEl.innerText = formal;
  if (internEl) internEl.innerText = intern;
}

function addTeamMember() {
  const input = document.getElementById('newTeamMemberInput');
  const statusSel = document.getElementById('newTeamMemberStatus');
  const raw = (input.value || '').trim();
  const status = statusSel ? statusSel.value : '正式';
  if (!raw) { alert('⚠️ 請輸入姓名'); return; }

  // 解析「姓名 (LK00001)」格式
  const m = raw.match(/^(.+?)\s*\((LK\d+)\)\s*$/i);
  let name = m ? m[1].trim() : raw;
  let uid  = m ? m[2].trim().toUpperCase() : '';

  // 純姓名 → 從主日候選清單自動帶 UID（唯一同名才自動帶）
  if (!uid && _mainMemberSuggestionsCache) {
    const matched = _mainMemberSuggestionsCache.filter(x => x.name === name);
    if (matched.length === 1) {
      uid = matched[0].uid;
    } else if (matched.length > 1) {
      alert(`⚠️ 主日有多位「${name}」，請從下拉選單點選正確的一位（會顯示編號區別）`);
      return;
    } else {
      // 主日沒有 → 提示先去主日建檔
      if (!confirm(`⚠️ 主日會友名單中查無「${name}」，仍要加入嗎？\n（建議先去主日會友管理建檔以取得系統編號）`)) {
        return;
      }
    }
  }

  // 重複檢查（同 UID 或同姓名）
  const dup = _editingTeamMembers.some(em =>
    (uid && em.uid === uid) || (!uid && em.name === name && !em.uid)
  );
  if (dup) { alert('此人已經在團員名單中'); return; }

  _editingTeamMembers.push({
    name: name,
    uid: uid,
    status: status,
    joinDate: new Date().toISOString()
  });

  input.value = '';
  renderTeamMembersTable();
  buildTeamMemberDatalist();
}

function updateTeamMemberStatus(idx, status) {
  if (_editingTeamMembers[idx]) {
    _editingTeamMembers[idx].status = status === '實習' ? '實習' : '正式';
    renderTeamMembersTable();
  }
}

function removeTeamMember(idx) {
  const m = _editingTeamMembers[idx];
  if (!m) return;
  if (confirm(`確定要將【${m.name}】從敬拜團員名單中移除嗎？`)) {
    _editingTeamMembers.splice(idx, 1);
    renderTeamMembersTable();
    buildTeamMemberDatalist();
  }
}

async function saveTeamMembersToServer() {
  try {
    const res = await callAPI('saveTeamMembers', { members: _editingTeamMembers });
    if (res && res.status === 'success') {
      _worshipTeamCache = null; // 失效：「位置與同工」下次開啟時會重新拉
      alert('✅ ' + (res.message || '已儲存'));
    } else {
      alert('❌ 儲存失敗：' + (res && res.message ? res.message : '未知錯誤'));
    }
  } catch (err) {
    alert('❌ 儲存失敗：' + err.message);
  }
}

// ==========================================
// 1. 公佈欄邏輯
// ==========================================
let currentViewMode = 'cards'; // 預設為玻璃卡片

function _findClosestCardIndex(data) {
  if (!data || data.length === 0) return -1;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const todayMs = today.getTime();

  let closestIdx = -1;
  let minDiffFuture = Infinity;
  let closestIdxPast = -1;
  let minDiffPast = Infinity;

  for (let i = 0; i < data.length; i++) {
    const dStr = data[i]['日期'];
    if (!dStr) continue;
    const d = parseDateSafe(dStr);
    d.setHours(0, 0, 0, 0);
    const dMs = d.getTime();
    
    const diff = dMs - todayMs;
    if (diff >= 0) {
      if (diff < minDiffFuture) {
        minDiffFuture = diff;
        closestIdx = i;
      }
    } else {
      const absDiff = Math.abs(diff);
      if (absDiff < minDiffPast) {
        minDiffPast = absDiff;
        closestIdxPast = i;
      }
    }
  }

  return closestIdx !== -1 ? closestIdx : closestIdxPast;
}

function switchView(mode) {
  currentViewMode = mode;
  const cardsContainer = document.getElementById('dashboardCardsContainer');
  const tableContainer = document.getElementById('dashboardTableContainer');
  const btnCards = document.getElementById('viewBtnCards');
  const btnTable = document.getElementById('viewBtnTable');
  
  if (mode === 'cards') {
    if (cardsContainer) cardsContainer.style.display = 'block';
    if (tableContainer) tableContainer.style.display = 'none';
    if (btnCards) btnCards.classList.add('active');
    if (btnTable) btnTable.classList.remove('active');
    
    // 自動平滑滾動至最接近今日日期的卡片
    setTimeout(() => {
      const closestCard = document.getElementById('worship-closest-date-item');
      if (closestCard) {
        closestCard.scrollIntoView({ behavior: 'smooth', block: 'center' });
      }
    }, 150);
  } else {
    if (cardsContainer) cardsContainer.style.display = 'none';
    if (tableContainer) tableContainer.style.display = 'block';
    if (btnCards) btnCards.classList.remove('active');
    if (btnTable) btnTable.classList.add('active');
  }
}

async function loadDashboard() {
  const container = document.getElementById('dashboardContainer');
  const yearSelect = document.getElementById('yearSelect');
  const quarterSelect = document.getElementById('quarterSelect');
  if (!container || !quarterSelect) return;

  let year, quarter;
  if (yearSelect) {
    year = yearSelect.value;
    quarter = quarterSelect.value;
  } else {
    const parts = quarterSelect.value.split('-');
    year = parts[0];
    quarter = parts[1];
  }
  container.innerHTML = `<div class="text-center p-5 text-primary"><div class="spinner-border"></div><div class="mt-2">同步 ${year}-${quarter} 資料中...</div></div>`;

  try {
    await ensurePositionsLoaded();
    const result = await callAPI('getSchedule', { year, quarter });
    if (result.status === 'success') {
      const syncTimeEl = document.getElementById('syncTime');
      if (syncTimeEl) syncTimeEl.innerText = new Date().toLocaleTimeString();
      renderDashboardTable(result.data);
    } else {
      container.innerHTML = `<div class="alert alert-warning text-center m-4">⚠️ ${result.message || '查無資料'}</div>`;
    }
  } catch (error) {
    container.innerHTML = `<div class="alert alert-danger text-center m-4">❌ 連線失敗<br><small>${error.message}</small></div>`;
  }
}

function renderDashboardTable(data) {
  const container = document.getElementById('dashboardContainer');
  loadedDashboardData = data || []; // 儲存已載入的資料供團員查詢使用
  
  // 載入新季度資料時，如果查詢框有內容則自動更新查詢結果，否則清除結果
  const memberSearchInput = document.getElementById('memberSearchInput');
  if (memberSearchInput && memberSearchInput.value.trim() === '') {
    clearMemberSearch();
  } else if (memberSearchInput) {
    queryMemberSchedule();
  }

  if (!data || data.length === 0) {
    container.innerHTML = '<div class="alert alert-light text-center m-4">📋 本季度暫無排班資料。</div>';
    return;
  }
  
  // 固定欄位：日期、聚會名稱、聚會類別、牧師、題目、經文、敬拜曲目
  // 然後才是服事同工（主領、配唱...等動態欄位）
  const fixedHeaders = ['日期', '聚會名稱', '聚會類別', '牧師', '題目', '經文', '敬拜曲目'];
  const allKeys = Object.keys(data[0]);
  const excludedHeaders = ['hasWarning', 'warningMessage', '年度', '季度', 'leaves'];
  const positionHeaders = currentPositions
    .map(p => p.positionName)
    .filter(role => role && data.some(row => Object.prototype.hasOwnProperty.call(row, role)));
  const fallbackHeaders = allKeys.filter(k =>
    !fixedHeaders.includes(k) &&
    !excludedHeaders.includes(k) &&
    !positionHeaders.includes(k)
  );
  const dynamicHeaders = [...positionHeaders, ...fallbackHeaders];
  const roleIcons = {
    '主領': '🎙️',
    '配唱1': '🎶',
    '配唱2': '🎶',
    '配唱3': '🎶',
    '吉他': '🎸',
    'BASS': '🎸',
    'Keyboard': '🎹',
    '鼓': '🥁',
    '其它': '✨',
    '音控': '🎚️',
    '投影': '🖥️'
  };

  // 建立雙視圖的容器
  container.innerHTML = `
    <div id="dashboardCardsContainer" class="view-container"></div>
    <div id="dashboardTableContainer" class="view-container" style="display: none;"></div>
  `;
  
  const cardsContainer = document.getElementById('dashboardCardsContainer');
  const tableContainer = document.getElementById('dashboardTableContainer');

  // 找出最接近今日的聚會卡片索引
  const closestIdx = _findClosestCardIndex(data);

  // --- 1. 渲染卡片視圖 ---
  let cardsHtml = `<div class="dashboard-grid">`;
  
  data.forEach((row, idx) => {
    const hasWarning = row.hasWarning || false;
    const warningMsg = row.warningMessage || '';
    
    // 卡片背景微調 (最接近今日的聚會卡片加上 id="worship-closest-date-item" 與 closest-date-card 樣式類別)
    const isClosest = (idx === closestIdx);
    const cardIdAttr = isClosest ? 'id="worship-closest-date-item"' : '';
    const closestClass = isClosest ? 'closest-date-card' : '';
    cardsHtml += `<div ${cardIdAttr} class="dashboard-card ${closestClass}" style="${hasWarning ? 'background: rgba(255, 243, 205, 0.3) !important;' : ''}">
      <div>
        <!-- 卡片頂部 -->
        <div class="card-top">
          <div class="card-date-info">
            <div class="card-date-row">
              <span class="card-date" style="${hasWarning ? 'color: #9a4005;' : ''}">${row['日期'] || ''}</span>
              ${hasWarning ? `<span class="date-warning-badge" title="${warningMsg}">⚠️ ${warningMsg}</span>` : ''}
            </div>
            <span class="card-meeting-name">${row['聚會名稱'] || ''}</span>
          </div>
          <span class="badge-g">${row['聚會類別'] || '聚會'}</span>
        </div>`;

    // 渲染服事同工
    cardsHtml += `<div class="card-section-title">👥 服事同工</div>
        <div class="personnel-grid">`;
    dynamicHeaders.forEach(role => {
      let person = (row[role] || '').trim();
      let personBadge = '';
      if (!person || person === '【待定】' || person === '待定') {
        personBadge = `<span class="badge-pending">【待定】</span>`;
      } else {
        personBadge = `<span class="badge-b">${person}</span>`;
      }
      cardsHtml += `<div class="personnel-item">
        <span class="personnel-role">${roleIcons[role] || '🙋'} ${role}：</span>${personBadge}
      </div>`;
    });
    cardsHtml += `</div>`;

    // 渲染敬拜曲目
    const songs = row['敬拜曲目'] || '';
    cardsHtml += `<div class="card-section-title">🎵 敬拜曲目</div>`;
    if (songs && songs !== '-' && songs !== '【待定】') {
      const songList = songs.split(/[\n\r,，、\/\\|；;]+/).map(s => s.trim()).filter(x => x);
      cardsHtml += `<div class="song-badges-container">`;
      songList.forEach(song => {
        cardsHtml += `<span class="song-badge-item">🎵 ${song}</span>`;
      });
      cardsHtml += `</div>`;
    } else {
      cardsHtml += `<div class="song-badges-container" style="color: #adb5bd; font-style: italic; justify-content: center; font-size: 0.82rem;">
        📋 尚未填入敬拜曲目
      </div>`;
    }

    cardsHtml += `</div>`;

    // 渲染講道資訊框
    const preacher = row['牧師'] || '';
    const title = row['題目'] || '';
    const scripture = row['經文'] || '';
    
    cardsHtml += `<div class="sermon-box">
      🎙️ <strong>講道：</strong>${preacher || '-'}<br>
      📖 <strong>題目：</strong>${title || '-'}<br>
      📜 <strong>經文：</strong>${formatScriptureLink(scripture)}
    </div>`;

    cardsHtml += `</div>`; // dashboard-card 結束
  });

  cardsHtml += `</div>`;
  if (cardsContainer) cardsContainer.innerHTML = cardsHtml;

  // --- 2. 渲染表格視圖 ---
  let tableHtml = `
    <div class="dashboard-table-container">
      <table class="dashboard-table">
        <thead>
          <tr>
            <th style="width: 125px;">日期</th>
            <th style="width: 100px;">聚會</th>
            <th style="width: 80px;">類別</th>
            <th style="min-width: 180px;">講道資訊</th>
  `;
  
  // 動態欄位表頭
  dynamicHeaders.forEach(role => {
    tableHtml += `<th>${roleIcons[role] || '🙋'} ${role}</th>`;
  });
  
  tableHtml += `
            <th style="min-width: 220px;">敬拜曲目</th>
          </tr>
        </thead>
        <tbody>
  `;
  
  data.forEach(row => {
    const hasWarning = row.hasWarning || false;
    const warningMsg = row.warningMessage || '';
    
    // 日期欄位內容 (請假警示加在日期下方)
    let dateCellContent = `<span style="font-weight:bold; color: ${hasWarning ? '#9a4005' : '#006030'};">${row['日期'] || ''}</span>`;
    if (hasWarning) {
      dateCellContent += `<br><div class="date-warning-badge" style="margin-top:4px;">⚠️ ${warningMsg}</div>`;
    }
    
    // 講道資訊整合
    const preacher = row['牧師'] || '';
    const title = row['題目'] || '';
    const scripture = row['經文'] || '';
    let sermonInfo = '';
    if (preacher) sermonInfo += `🎙️ <strong>${preacher}</strong><br>`;
    if (title && title !== '-') sermonInfo += `📖 <span class="text-muted">${title}</span><br>`;
    if (scripture && scripture !== '-') sermonInfo += `📜 <span class="text-muted" style="font-size:0.82rem;">${formatScriptureLink(scripture)}</span>`;
    if (!sermonInfo) sermonInfo = '-';
    
    tableHtml += `<tr class="${hasWarning ? 'warning-row' : ''}">
      <td>${dateCellContent}</td>
      <td><strong>${row['聚會名稱'] || ''}</strong></td>
      <td><span class="badge-g">${row['聚會類別'] || '聚會'}</span></td>
      <td style="text-align: left; vertical-align: top;">${sermonInfo}</td>
    `;
    
    // 動態服事人員
    dynamicHeaders.forEach(role => {
      let person = (row[role] || '').trim();
      let personBadge = '';
      if (!person || person === '【待定】' || person === '待定') {
        personBadge = `<span class="badge-pending">【待定】</span>`;
      } else {
        personBadge = `<span class="badge-b">${person}</span>`;
      }
      tableHtml += `<td>${personBadge}</td>`;
    });
    
    // 敬拜曲目
    const songs = row['敬拜曲目'] || '';
    let songsContent = '';
    if (songs && songs !== '-' && songs !== '【待定】') {
      const songList = songs.split(/[\n\r,，、\/\\|；;]+/).map(s => s.trim()).filter(x => x);
      songsContent = `<div style="display:flex; flex-wrap:wrap; gap:4px; justify-content:center;">`;
      songList.forEach(song => {
        songsContent += `<span class="song-badge">🎵 ${song}</span>`;
      });
      songsContent += `</div>`;
    } else {
      songsContent = `<span style="color:#888; font-style:italic; font-size:0.82rem;">尚未填入曲目</span>`;
    }
    
    tableHtml += `
      <td>${songsContent}</td>
    </tr>`;
  });
  
  tableHtml += `
        </tbody>
      </table>
    </div>
  `;
  if (tableContainer) tableContainer.innerHTML = tableHtml;

  // 切換至目前的視圖模式 (同步按鈕狀態與顯隱)
  switchView(currentViewMode);
}

// ==========================================
// 2. 位置與人員設定
// ==========================================
async function loadPositions() {
  const tbody = document.getElementById('positionsTbody');
  if (!tbody) return;

  // 同步載入：敬拜團員名單（給標籤式選人用）+ 位置資料
  await ensureWorshipTeamLoaded(true); // 強制刷新，確保是最新名單

  const result = await callAPI('getPositions', {});
  tbody.innerHTML = '';
  if (result.status === 'success') {
    result.data.length === 0
      ? addPositionRow('主領', '', '是')
      : result.data.forEach(i => addPositionRow(i.positionName, i.personnel, i.isRequired || '是'));
    if (sortablePositions) sortablePositions.destroy();
    sortablePositions = new Sortable(tbody, {
      handle: '.drag-handle',
      animation: 150,
      filter: 'input, select, button, .personnel-picker, .badge, .btn-close',
      preventOnFilter: false
    });
  }
}

function addPositionRow(posName, personnel, isRequired = "是") {
  const tbody = document.getElementById('positionsTbody');
  const tr = document.createElement('tr');
  tr.innerHTML = `
    <td class="text-center align-middle drag-handle" style="cursor: grab; color: #adb5bd;">☰</td>
    <td><input type="text" class="form-control form-control-sm pos-name text-center" value="${posName}" onclick="this.select()"></td>
    <td class="personnel-cell"></td>
    <td><select class="form-select form-select-sm pos-required"><option value="是" ${isRequired === "是" ? "selected" : ""}>必排</option><option value="否" ${isRequired === "否" ? "selected" : ""}>非必排</option></select></td>
    <td class="text-center"><button class="btn btn-outline-danger btn-sm" onclick="this.closest('tr').remove()">x</button></td>
  `;
  tbody.appendChild(tr);
  renderPersonnelTagPicker(tr.querySelector('.personnel-cell'), personnel || '');
}

/**
 * 標籤式多選人員（從敬拜團員名單挑選 + 內建關鍵字搜尋）
 * 仍以隱藏欄位 .pos-personnel 儲存「逗號分隔字串」，保持後端相容
 */
function renderPersonnelTagPicker(td, currentValue) {
  const selected = (currentValue || '').split(/[,、]/).map(s => s.trim()).filter(x => x);
  td.innerHTML = `
    <div class="personnel-picker d-flex flex-wrap gap-1 align-items-center p-1 border rounded bg-white"
         style="min-height:34px;">
      <div class="tags-container d-flex flex-wrap gap-1 flex-grow-1"></div>
      <button type="button" class="btn btn-sm btn-outline-primary add-personnel-btn px-2 py-0"
              style="font-size:0.78rem; white-space:nowrap;">＋ 加入</button>
    </div>
    <input type="hidden" class="pos-personnel" value="${selected.join(',')}">
  `;

  const tagsContainer = td.querySelector('.tags-container');
  const hiddenInput = td.querySelector('.pos-personnel');
  const addBtn = td.querySelector('.add-personnel-btn');

  function refreshHidden() {
    const names = Array.from(tagsContainer.querySelectorAll('.tag-name'))
      .map(el => el.dataset.name);
    hiddenInput.value = names.join(',');
  }

  function addTag(name, opts = {}) {
    name = (name || '').trim();
    if (!name) return;
    // 重複跳過
    const exists = Array.from(tagsContainer.querySelectorAll('.tag-name'))
      .some(el => el.dataset.name === name);
    if (exists) return;

    const isInTeam = _worshipTeamCache && _worshipTeamCache.some(m => m.name === name);
    const status = isInTeam ? (_worshipTeamCache.find(m => m.name === name).status || '正式') : '';
    const bgClass = !isInTeam ? 'bg-secondary'
                  : status === '實習' ? 'bg-warning text-dark' : 'bg-primary';

    const tag = document.createElement('span');
    tag.className = `badge ${bgClass} tag-name d-inline-flex align-items-center`;
    tag.dataset.name = name;
    tag.style.cssText = 'font-size:0.78rem; padding:0.35em 0.55em; gap:0.3em;';
    tag.title = isInTeam ? `敬拜團員（${status}）` : '⚠️ 此人不在敬拜團員名單中';
    tag.innerHTML = `
      ${!isInTeam ? '⚠️ ' : ''}${name}
      <button type="button" class="btn-close" style="font-size:0.5rem;" aria-label="移除"></button>
    `;
    tag.querySelector('button').onclick = (e) => {
      e.stopPropagation();
      tag.remove();
      refreshHidden();
    };
    tagsContainer.appendChild(tag);
    refreshHidden();
  }

  selected.forEach(n => addTag(n));

  addBtn.onclick = async (e) => {
    e.stopPropagation();
    const team = await ensureWorshipTeamLoaded();
    if (!team || team.length === 0) {
      alert('⚠️ 敬拜團員名單為空\n請先到「👥 敬拜團員名單」分頁新增成員');
      return;
    }
    const existingNames = new Set(
      Array.from(tagsContainer.querySelectorAll('.tag-name')).map(el => el.dataset.name)
    );
    // 正式優先 → 實習 → 已選的（顯示為 disabled）
    const items = team
      .slice()
      .sort((a, b) => {
        if (a.status !== b.status) return a.status === '正式' ? -1 : 1;
        return (a.name || '').localeCompare(b.name || '');
      })
      .map(m => ({
        label: m.name,
        subLabel: (m.status === '實習' ? '🎓 實習' : '⭐ 正式') + (m.uid ? `　${m.uid}` : ''),
        value: m.name,
        disabled: existingNames.has(m.name)
      }));
    _showFloatingDropdown(addBtn, items, (it) => addTag(it.value), {
      placeholder: '🔍 輸入姓名或編號搜尋敬拜團員...',
      emptyText: '查無相符的敬拜團員',
      width: 320
    });
  };
}

async function savePositionsToServer() {
  const rows = document.querySelectorAll('#positionsTbody tr');
  let positionsData = [];
  rows.forEach(tr => {
    const name = tr.querySelector('.pos-name').value.trim();
    if (name) positionsData.push({ positionName: name, personnel: tr.querySelector('.pos-personnel').value.trim(), isRequired: tr.querySelector('.pos-required').value });
  });
  const btn = document.querySelector('button[onclick="savePositionsToServer()"]');
  btn.disabled = true;
  await callAPI('savePositions', { positionsData });
  userNotification.success("✅ 位置設定儲存成功！");
  btn.disabled = false;
}

// ==========================================
// 3. 服事安排 (外部框架 + 智慧填補)
// ==========================================

let currentRowIndexForLeave = -1;

async function initScheduleTab() {
  const result = await callAPI('getPositions', {});
  if (result.status === 'success') {
    currentPositions = result.data;
    let nameSet = new Set();
    currentPositions.forEach(pos => (pos.personnel || '').split(',').forEach(n => n.trim() && nameSet.add(n.trim())));
    uniquePersonnel = Array.from(nameSet).sort();
  }
}

async function loadScheduleByQuarter() {
  const yearSelect = document.getElementById('editYearSelect');
  const quarterSelect = document.getElementById('editQuarterSelect');
  
  let year, quarter;
  if (yearSelect && quarterSelect) {
    year = yearSelect.value;
    quarter = quarterSelect.value;
  } else {
    const select = document.getElementById('editQuarterSelect');
    const parts = select.value.split('-');
    year = parts[0];
    quarter = parts[1];
  }
  
  document.getElementById('previewContainer').style.display = 'none';
  document.getElementById('saveScheduleBtn').style.display = 'none';
  const placeholder = document.getElementById('previewPlaceholder');
  placeholder.style.display = 'block'; 
  
  placeholder.innerHTML = `<div class="p-4 text-center text-success"><div class="spinner-border spinner-border-sm"></div> 從外部載入 ${year} ${quarter} 框架中...</div>`;
  
  try {
    const result = await callAPI('getSchedule', { year, quarter });
    if (result.status === 'success' && result.data.length > 0) {
      generatedScheduleData = result.data;
      renderPreviewTable(generatedScheduleData);
    } else {
      placeholder.innerHTML = `<div class="alert alert-warning m-4">查無 ${year} ${quarter} 資料，且外部也無此季度的聚會紀錄。</div>`;
    }
  } catch (error) { 
    placeholder.innerHTML = `<div class="alert alert-danger m-4">❌ 讀取失敗，請確認網路連線。</div>`;
  }
}

async function loadScheduleByDateRange() {
  const start = document.getElementById('queryStartDate').value;
  const end = document.getElementById('queryEndDate').value;
  if (!start || !end) return userNotification.warning("請先設定起訖日期");

  document.getElementById('previewContainer').style.display = 'none';
  document.getElementById('saveScheduleBtn').style.display = 'none';
  const placeholder = document.getElementById('previewPlaceholder');
  placeholder.style.display = 'block'; 

  placeholder.innerHTML = '<div class="p-4 text-center text-primary"><div class="spinner-border spinner-border-sm"></div> 區間資料讀取中...</div>';
  
  try {
    const result = await callAPI('getScheduleByDateRange', { startDate: start, endDate: end });
    if (result.status === 'success' && result.data && result.data.length > 0) {
      generatedScheduleData = result.data;
      renderPreviewTable(generatedScheduleData);
    } else {
      placeholder.innerHTML = `<div class="alert alert-info m-4">${start} 至 ${end} 無存檔資料。</div>`;
    }
  } catch (error) { 
    placeholder.innerHTML = `<div class="alert alert-danger m-4">❌ 區間讀取失敗。</div>`;
  }
}

// 🌟 建立聚會日期（單一 / 批量）
let _addMeetingMode = 'single'; // 'single' | 'batch'

function openAddExtraModal() {
  // 重置為單一模式
  switchAddMeetingMode('single');
  document.getElementById('extraDate').value = '';
  document.getElementById('extraName').value = '';
  document.getElementById('extraType').value = '';
  document.getElementById('batchStartDate').value = '';
  document.getElementById('batchEndDate').value = '';
  // 預設勾選「日」（週日）
  document.querySelectorAll('#batchWeekdayPicker input[type=checkbox]').forEach(cb => cb.checked = (cb.value === '0'));
  document.getElementById('batchPreviewArea').style.display = 'none';
  // 初始化年份下拉
  _initBatchYearSelect();
  bootstrap.Modal.getOrCreateInstance(document.getElementById('extraMeetingModal')).show();
}

// 初始化年份選單（當年 -1 ~ +2）
function _initBatchYearSelect() {
  const sel = document.getElementById('batchYearSelect');
  if (!sel) return;
  const cur = new Date().getFullYear();
  sel.innerHTML = '';
  for (let y = cur - 1; y <= cur + 2; y++) {
    const opt = document.createElement('option');
    opt.value = y;
    opt.textContent = y;
    if (y === cur) opt.selected = true;
    sel.appendChild(opt);
  }
}

// 套用季度快捷：填入日期區間並觸發預覽
function applyQuarterShortcut(q) {
  const year = document.getElementById('batchYearSelect').value;
  const ranges = {
    1: [`${year}-01-01`, `${year}-03-31`],
    2: [`${year}-04-01`, `${year}-06-30`],
    3: [`${year}-07-01`, `${year}-09-30`],
    4: [`${year}-10-01`, `${year}-12-31`]
  };
  const [start, end] = ranges[q];
  document.getElementById('batchStartDate').value = start;
  document.getElementById('batchEndDate').value   = end;
  // 高亮被選中的季度鈕
  document.querySelectorAll('[onclick^="applyQuarterShortcut"]').forEach(btn => {
    btn.classList.toggle('active', btn.getAttribute('onclick') === `applyQuarterShortcut(${q})`);
  });
  updateBatchPreview();
}

function switchAddMeetingMode(mode) {
  _addMeetingMode = mode;
  document.getElementById('addMode-single').style.display = mode === 'single' ? '' : 'none';
  document.getElementById('addMode-batch').style.display  = mode === 'batch'  ? '' : 'none';
  document.getElementById('tab-single').classList.toggle('active', mode === 'single');
  document.getElementById('tab-batch').classList.toggle('active', mode === 'batch');
}

// 批量模式：即時預覽展開的日期清單
function _expandBatchDates() {
  const start = document.getElementById('batchStartDate').value;
  const end   = document.getElementById('batchEndDate').value;
  const weekdays = new Set(
    Array.from(document.querySelectorAll('#batchWeekdayPicker input:checked')).map(cb => parseInt(cb.value))
  );
  if (!start || !end || weekdays.size === 0) return [];
  const result = [];
  const cur = new Date(start + 'T00:00:00');
  const last = new Date(end   + 'T00:00:00');
  while (cur <= last) {
    if (weekdays.has(cur.getDay())) {
      const y = cur.getFullYear();
      const m = String(cur.getMonth() + 1).padStart(2, '0');
      const d = String(cur.getDate()).padStart(2, '0');
      result.push(`${y}-${m}-${d}`);
    }
    cur.setDate(cur.getDate() + 1);
  }
  return result;
}

// 批量預覽（綁定在 input change 時呼叫，HTML 用 oninput 觸發）
function updateBatchPreview() {
  const dates = _expandBatchDates();
  const area = document.getElementById('batchPreviewArea');
  if (dates.length === 0) { area.style.display = 'none'; return; }
  area.style.display = '';
  area.innerHTML = `共 <b>${dates.length}</b> 個日期：` + dates.map(d => `<span class="badge bg-secondary me-1">${d}</span>`).join('');
}

async function _hydrateRowsFromCalendar(rows) {
  const entries = rows
    .map(row => ({ date: row['日期'], meetingName: row['聚會名稱'] || '' }))
    .filter(e => e.date);
  if (entries.length === 0) return;

  try {
    const res = await callAPI('getCalendarDataForDates', { entries });
    const calData = res && res.status === 'success' ? (res.data || {}) : (res || {});
    rows.forEach(row => {
      const date = row['日期'];
      const key = row['聚會名稱'] ? `${date}|${row['聚會名稱']}` : date;
      const cd = calData[key] || calData[date] || {};
      if (!row['聚會名稱'] && cd.namedEvent && cd.namedEvent.title) {
        row['聚會名稱'] = String(cd.namedEvent.title).trim();
      }
      if (!row['聚會類別'] && cd.sermon && cd.sermon.typeName) {
        row['聚會類別'] = String(cd.sermon.typeName).trim();
      }
      if (cd.sermon && cd.sermon.values) {
        row['牧師'] = String(cd.sermon.values['講員'] || row['牧師'] || '');
        row['題目'] = String(cd.sermon.values['講題'] || row['題目'] || '');
        row['經文'] = String(cd.sermon.values['經文'] || row['經文'] || '');
      }
    });
  } catch (err) {
    console.warn('新增日期時帶入行事曆資料失敗:', err);
    userNotification.warning('日期已新增，但行事曆資料暫時帶入失敗，儲存後重新讀取可再同步。');
  }
}

async function confirmAddExtraMeeting() {
  const rowsToHydrate = [];
  if (_addMeetingMode === 'single') {
    // ── 單一模式 ──
    const date = document.getElementById('extraDate').value;
    const name = document.getElementById('extraName').value.trim();
    const type = document.getElementById('extraType').value.trim(); // 留空由行事曆帶入
    if (!date) return userNotification.warning("請選擇日期！");

    const existing = generatedScheduleData.find(r => r['日期'] === date);
    if (existing) return userNotification.warning(`${date} 已存在，請勿重複新增。`);

    const row = {
      '年度': date.substring(0, 4),
      '季度': `Q${Math.ceil((new Date(date + 'T00:00:00').getMonth() + 1) / 3)}`,
      '日期': date,
      '聚會名稱': name,   // 留空由行事曆帶入
      '聚會類別': type,   // 留空由行事曆帶入
      'leaves': []
    };
    generatedScheduleData.push(row);
    rowsToHydrate.push(row);

  } else {
    // ── 批量模式 ──
    const dates = _expandBatchDates();
    if (dates.length === 0) return userNotification.warning("請設定日期區間並至少勾選一個週幾！");

    const existingDates = new Set(generatedScheduleData.map(r => r['日期']));
    let added = 0;
    dates.forEach(date => {
      if (existingDates.has(date)) return; // 跳過已存在
      const row = {
        '年度': date.substring(0, 4),
        '季度': `Q${Math.ceil((new Date(date + 'T00:00:00').getMonth() + 1) / 3)}`,
        '日期': date,
        '聚會名稱': '',   // 留空，由行事曆帶入
        '聚會類別': '',   // 留空，由行事曆帶入
        'leaves': []
      };
      generatedScheduleData.push(row);
      rowsToHydrate.push(row);
      added++;
    });
    if (added === 0) return userNotification.warning("所選區間的日期均已存在，無需重複新增。");
    userNotification.success(`已批量新增 ${added} 個聚會日期`);
  }

  await _hydrateRowsFromCalendar(rowsToHydrate);
  generatedScheduleData.sort((a, b) => parseDateSafe(a['日期']) - parseDateSafe(b['日期']));
  renderPreviewTable(generatedScheduleData);
  bootstrap.Modal.getOrCreateInstance(document.getElementById('extraMeetingModal')).hide();
}



// 🌟 每一列專屬的請假設定
function openRowLeaveModal(idx) {
  currentRowIndexForLeave = idx;
  const currentLeaves = generatedScheduleData[idx].leaves || [];
  let html = '<div class="row g-2">';
  uniquePersonnel.forEach(name => {
    const isChecked = currentLeaves.includes(name) ? 'checked' : '';
    html += `<div class="col-6 col-sm-4"><div class="form-check"><input class="form-check-input leave-checkbox" type="checkbox" value="${name}" id="chk_${name}" ${isChecked}><label class="form-check-label" for="chk_${name}">${name}</label></div></div>`;
  });
  document.getElementById('leaveModalBody').innerHTML = html + '</div>';
  bootstrap.Modal.getOrCreateInstance(document.getElementById('leaveModal')).show();
}

function confirmLeaveSelection() {
  const selected = Array.from(document.querySelectorAll('.leave-checkbox:checked')).map(cb => cb.value);
  generatedScheduleData[currentRowIndexForLeave].leaves = selected;
  bootstrap.Modal.getOrCreateInstance(document.getElementById('leaveModal')).hide();
  renderPreviewTable(generatedScheduleData);
}

// 🌟 智慧填補
function smartGenerateSchedule() {
  if (generatedScheduleData.length === 0) return userNotification.warning("請先載入季度框架或新增日期！");

  let leaderPool = [], previousLeader = null, consecutive = {};
  currentPositions.forEach(p => consecutive[p.positionName] = {});

  generatedScheduleData.forEach((row) => {
    let leaves = row.leaves || [];
    let assigned = [];

    currentPositions.forEach(pos => {
      let name = pos.positionName;
      if (!row[name] || row[name] === '【待定】') {
        let candidates = (pos.personnel || '').split(',').map(s => s.trim()).filter(x => x);
        let pick = "";
        
        if (name === '主領') {
          if (!leaderPool.length) leaderPool = [...candidates];
          let valid = leaderPool.filter(p => !leaves.includes(p) && p !== previousLeader);
          if (valid.length) { pick = valid[Math.floor(Math.random()*valid.length)]; leaderPool = leaderPool.filter(p => p !== pick); }
        } else {
          let valid = candidates.filter(p => !leaves.includes(p) && !assigned.includes(p) && (consecutive[name][p]||0) < 2);
          if (valid.length) pick = valid[Math.floor(Math.random()*valid.length)];
        }
        
        row[name] = pick || (pos.isRequired === '是' ? "【待定】" : "");
      }

      let finalPick = row[name];
      if (finalPick && finalPick !== "【待定】") assigned.push(finalPick);
      if (name === '主領') previousLeader = finalPick;
      
      let allCands = (pos.personnel || '').split(',').map(s => s.trim()).filter(x => x);
      allCands.forEach(c => consecutive[name][c] = (c === finalPick ? (consecutive[name][c]||0)+1 : 0));
    });
  });
  
  renderPreviewTable(generatedScheduleData);
}

// ==========================================
// 🌟 預覽表格渲染 - 修正版（可讀性 + 橫向捲動）
// ==========================================
function renderPreviewTable(data) {
  const container = document.getElementById('previewContainer');
  const thead = document.getElementById('previewThead');
  const tbody = document.getElementById('previewTbody');
  thead.innerHTML = ''; 
  tbody.innerHTML = '';
  if (!data.length) return;

  // 欄位定義：固定寬度讓每欄都看得清楚
  const colConfig = {
    '請假/狀態': '110px',
    '日期':      '100px',
    '聚會名稱':  '150px',
    '聚會類別':  '100px',
  };
  // 職位欄預設寬度
  const posColWidth = '110px';

  let headers = ['請假/狀態', '日期', '聚會名稱', '聚會類別', ...currentPositions.map(p => p.positionName)];

  // --- 表頭 ---
  let trH = document.createElement('tr');
  headers.forEach(h => {
    let th = document.createElement('th');
    th.innerText = h;
    th.style.minWidth = colConfig[h] || posColWidth;
    th.style.whiteSpace = 'nowrap';
    th.style.padding = '12px 10px';
    th.style.textAlign = 'center';
    th.style.fontSize = '0.9rem';
    trH.appendChild(th);
  });
  thead.appendChild(trH);

  // --- 資料列 ---
  data.forEach((row, idx) => {
    let tr = document.createElement('tr');

    headers.forEach(h => {
      let td = document.createElement('td');
      td.style.padding = '8px 10px';
      td.style.verticalAlign = 'middle';
      td.style.textAlign = 'center';
      td.style.minWidth = colConfig[h] || posColWidth;

      if (h === '請假/狀態') {
        // 請假按鈕 + 徽章
        let leaveBadges = (row.leaves || [])
          .map(n => `<span class="badge bg-danger me-1 mt-1" style="font-size:0.75rem;">${n}</span>`)
          .join('');
        td.innerHTML = `
          <button class="btn btn-sm btn-outline-secondary py-0 px-2 mb-1" 
                  style="font-size:0.78rem; white-space:nowrap;" 
                  onclick="openRowLeaveModal(${idx})">設請假</button>
          <div style="max-width:105px; white-space:normal; margin:0 auto;">${leaveBadges}</div>`;

      } else if (h === '日期') {
        td.innerHTML = `<span class="badge bg-secondary" style="font-size:0.85rem; white-space:nowrap;">${row[h] || ''}</span>`;

      } else if (h === '聚會名稱' || h === '聚會類別') {
        let input = document.createElement('input');
        input.type = 'text';
        input.className = 'form-control form-control-sm text-center border-0 bg-transparent fw-bold';
        input.style.minWidth = colConfig[h];
        input.style.fontSize = '0.88rem';
        input.style.color = h === '聚會名稱' ? '#198754' : '#0d6efd';
        input.value = row[h] || '';
        input.onclick = function() { this.select(); };
        input.onchange = (e) => { generatedScheduleData[idx][h] = e.target.value.trim(); };
        td.appendChild(input);

      } else {
        // 職位下拉選單
        let pos = currentPositions.find(p => p.positionName === h);
        let cands = (pos?.personnel || '').split(',').map(s => s.trim()).filter(x => x);
        let sel = document.createElement('select');
        sel.className = 'form-select form-select-sm text-center';
        sel.style.minWidth = posColWidth;
        sel.style.fontSize = '0.88rem';
        sel.style.border = '1px solid #dee2e6';
        sel.style.borderRadius = '6px';
        sel.style.backgroundColor = row[h] === '【待定】' ? '#ffebee' : '#fff';
        sel.style.color = row[h] === '【待定】' ? '#c62828' : 'inherit';
        sel.style.fontWeight = row[h] === '【待定】' ? 'bold' : 'normal';

        // 預設選項
        const defaultVal = pos?.isRequired === '是' ? '【待定】' : '';
        const defaultLabel = pos?.isRequired === '是' ? '【待定】' : '－';
        sel.innerHTML = `<option value="${defaultVal}">${defaultLabel}</option>` 
          + cands.map(c => `<option value="${c}" ${row[h] === c ? 'selected' : ''}>${c}</option>`).join('');

        sel.onchange = function() {
          generatedScheduleData[idx][h] = this.value;
          validateCellSelection(idx, h, this.value, this, true);
        };
        validateCellSelection(idx, h, row[h], sel, false);
        td.appendChild(sel);
      }

      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });

  // 顯示容器，確保橫向可捲動
  container.style.overflowX = 'auto';
  container.style.WebkitOverflowScrolling = 'touch';
  document.getElementById('previewPlaceholder').style.display = 'none';
  container.style.display = 'block';
  document.getElementById('saveScheduleBtn').style.display = 'inline-block';
}

async function saveGeneratedSchedule() {
  const btn = document.getElementById('saveScheduleBtn'); 
  btn.disabled = true; 
  btn.innerText = "儲存中...";
  const result = await callAPI('saveSchedule', { scheduleData: generatedScheduleData });
  if (result.status === 'success') { 
    userNotification.success("🎉 排班表已成功存檔！");
    loadDashboard(); 
    switchTab('dashboard'); 
  }
  btn.disabled = false; 
  btn.innerText = "💾 儲存並發佈";
}

// 🌟 驗證手動選擇同工是否違反規則（請假、重複、連續三週以上服事）
function validateCellSelection(idx, positionName, value, selectEl, triggerAlert = false) {
  if (!value || value === '【待定】' || value === '待定') {
    selectEl.style.backgroundColor = value === '【待定】' ? '#ffebee' : '#fff';
    selectEl.style.color = value === '【待定】' ? '#c62828' : 'inherit';
    selectEl.style.fontWeight = value === '【待定】' ? 'bold' : 'normal';
    selectEl.title = '';
    return;
  }

  const row = generatedScheduleData[idx];

  // 1. 該人員請假
  const leaves = row.leaves || [];
  const hasLeave = leaves.includes(value);

  // 2. 同日重複
  const duplicatePositions = currentPositions
    .map(p => p.positionName)
    .filter(posName => posName !== positionName && row[posName] === value);
  const hasDuplicate = duplicatePositions.length > 0;

  // 3. 連續三週以上服事
  const inRow = (r) => {
    if (!r) return false;
    return currentPositions.some(p => r[p.positionName] === value);
  };
  const inPrev1 = idx >= 1 && inRow(generatedScheduleData[idx - 1]);
  const inPrev2 = idx >= 2 && inRow(generatedScheduleData[idx - 2]);
  const inNext1 = idx < generatedScheduleData.length - 1 && inRow(generatedScheduleData[idx + 1]);
  const inNext2 = idx < generatedScheduleData.length - 2 && inRow(generatedScheduleData[idx + 2]);
  const hasConsecutive = (inPrev1 && inPrev2) || (inPrev1 && inNext1) || (inNext1 && inNext2);

  const warnings = [];
  if (hasLeave) warnings.push(`[${value}] 此日請假`);
  if (hasDuplicate) warnings.push(`[${value}] 此日重複服事 (${duplicatePositions.join('、')})`);
  if (hasConsecutive) warnings.push(`[${value}] 已連續三週以上服事`);

  if (hasLeave || hasDuplicate) {
    // 嚴重違規：紅色標示 (Bootstrap Danger)
    const warnMsg = warnings.join(' | ');
    selectEl.style.backgroundColor = '#f8d7da'; // light red
    selectEl.style.color = '#842029';           // dark red
    selectEl.style.fontWeight = 'bold';
    selectEl.title = `❌ 錯誤：${warnMsg}`;
    
    // 彈出通知 (僅限使用者手動切換時，避免渲染時彈出一大堆)
    if (triggerAlert && typeof userNotification !== 'undefined' && userNotification.warning) {
      userNotification.warning(`❌ 錯誤：${warnMsg}`);
    }
  } else if (hasConsecutive) {
    // 警示違規：黃色標示 (Bootstrap Warning)
    const warnMsg = warnings.join(' | ');
    selectEl.style.backgroundColor = '#fff3cd'; // light yellow
    selectEl.style.color = '#664d03';           // dark gold
    selectEl.style.fontWeight = 'bold';
    selectEl.title = `⚠️ 警告：${warnMsg}`;
    
    // 彈出通知 (僅限使用者手動切換時，避免渲染時彈出一大堆)
    if (triggerAlert && typeof userNotification !== 'undefined' && userNotification.warning) {
      userNotification.warning(`⚠️ 警告：${warnMsg}`);
    }
  } else {
    // 無違規：還原樣式
    selectEl.style.backgroundColor = '#fff';
    selectEl.style.color = 'inherit';
    selectEl.style.fontWeight = 'normal';
    selectEl.title = '';
  }
}

// 🔍 團員個人班表查詢與統計邏輯
function queryMemberSchedule() {
  const input = document.getElementById('memberSearchInput');
  const resultDiv = document.getElementById('memberSearchResult');
  if (!input || !resultDiv) return;

  const kw = input.value.trim();
  if (!kw) {
    resultDiv.innerHTML = '';
    resultDiv.style.display = 'none';
    return;
  }

  if (!loadedDashboardData || loadedDashboardData.length === 0) {
    resultDiv.innerHTML = '<div class="alert alert-light text-center mb-0">無本季排班資料可供查詢</div>';
    resultDiv.style.display = 'block';
    return;
  }

  // 尋找此同工在所有職位中的排班
  const matches = [];
  loadedDashboardData.forEach(row => {
    const positions = [];
    currentPositions.forEach(pos => {
      const role = pos.positionName;
      if (row[role] && String(row[role]).trim() === kw) {
        positions.push(role);
      }
    });

    if (positions.length > 0) {
      matches.push({
        date: row['日期'],
        meetingName: row['聚會名稱'] || '主日崇拜',
        meetingType: row['聚會類別'] || '華語',
        roles: positions
      });
    }
  });

  resultDiv.style.display = 'block';
  if (matches.length === 0) {
    resultDiv.innerHTML = `<div class="alert alert-warning text-center mb-0">⚠️ 查無 <strong>${kw}</strong> 在本季度的服事安排。</div>`;
  } else {
    let html = `
      <div class="card border-success shadow-sm" style="background: rgba(25, 135, 84, 0.02); border-radius: 12px;">
        <div class="card-body py-3">
          <h6 class="card-title fw-bold text-success mb-2 d-flex justify-content-between align-items-center">
            <span>🎉 查詢結果：<strong>${kw}</strong> 本季服事統計</span>
            <span class="badge bg-success fs-6 rounded-pill">共 ${matches.length} 天服事</span>
          </h6>
          <div class="table-responsive mt-2">
            <table class="table table-sm table-bordered bg-white align-middle text-center mb-0" style="font-size: 0.88rem; border-radius: 8px; overflow: hidden; border: 1px solid rgba(0,0,0,0.08);">
              <thead class="table-light">
                <tr>
                  <th style="width: 25%;">日期</th>
                  <th style="width: 35%;">聚會名稱</th>
                  <th style="width: 20%;">聚會類別</th>
                  <th style="width: 20%;">擔任位置</th>
                </tr>
              </thead>
              <tbody>
    `;

    matches.forEach(m => {
      const rolesBadge = m.roles.map(r => `<span class="badge bg-primary rounded-pill me-1" style="font-size:0.75rem; padding: 4px 8px;">${r}</span>`).join('');
      html += `
        <tr>
          <td><span class="badge bg-secondary rounded-pill" style="font-size:0.82rem;">${m.date}</span></td>
          <td><strong>${m.meetingName}</strong></td>
          <td><span class="badge-g" style="font-size:0.78rem;">${m.meetingType}</span></td>
          <td>${rolesBadge}</td>
        </tr>
      `;
    });

    html += `
              </tbody>
            </table>
          </div>
        </div>
      </div>
    `;
    resultDiv.innerHTML = html;
  }
}

function clearMemberSearch() {
  const input = document.getElementById('memberSearchInput');
  const resultDiv = document.getElementById('memberSearchResult');
  if (input) input.value = '';
  if (resultDiv) {
    resultDiv.innerHTML = '';
    resultDiv.style.display = 'none';
  }
}

// 🌟 顯示團員個人班表查詢的可搜尋模糊下拉選單
function showMemberSearchDropdown(anchorEl) {
  if (!uniquePersonnel || uniquePersonnel.length === 0) {
    let nameSet = new Set();
    currentPositions.forEach(pos => (pos.personnel || '').split(',').forEach(n => n.trim() && nameSet.add(n.trim())));
    uniquePersonnel = Array.from(nameSet).sort();
  }

  const items = uniquePersonnel.map(name => ({
    label: name,
    value: name
  }));

  _showFloatingDropdown(anchorEl, items, (item) => {
    anchorEl.value = item.value;
    queryMemberSchedule();
    _hideFloatingDropdown();
  }, {
    placeholder: '🔍 輸入關鍵字模糊搜尋...',
    emptyText: '查無此團員',
    width: anchorEl.offsetWidth
  });
}

// 經文超連結腳本，點擊後進入 PPT 產生器自動查詢
function formatScriptureLink(scriptureText) {
  if (!scriptureText || scriptureText.trim() === '-' || scriptureText.trim() === '') {
    return '-';
  }
  const text = scriptureText.trim();
  const url = `../LKC_ppt_generator/index.html?query=${encodeURIComponent(text)}&lang=zh&ver=unv&auto=1`;
  return `<a href="${url}" target="_blank" class="scripture-link" title="點擊進入自動製作 PPT">${text}</a>`;
}
