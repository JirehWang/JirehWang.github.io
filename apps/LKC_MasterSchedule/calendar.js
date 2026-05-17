/**
 * 教會行事曆 - 月曆視圖（Phase 2）
 * 依賴：FullCalendar 6 + 中央 churchAPI
 */

let _calendar = null;
let _types = { tree: [], flat: [] }; // {tree:[]rootTypes(with children), flat:[]allTypes}
let _selectableTypes = []; // 可被選為事項類型的（葉子節點 或 無子的頂層）
let _activeTypeIds = new Set(); // 當前篩選顯示的 typeIds
let _fieldsByType = {}; // typeId → fields (cache)
let _currentDetailEvent = null;

async function callAPI(action, data) {
  if (typeof window.churchAPI !== 'function') throw new Error('config.js 尚未載入');
  return await window.churchAPI(action, data || {});
}

window.addEventListener('DOMContentLoaded', async () => {
  initCalendar();
  await loadTypesAndChips();
});

// ─────────────────────────────────────────────────────────────
// 1. 初始化 FullCalendar
// ─────────────────────────────────────────────────────────────
function initCalendar() {
  const el = document.getElementById('calendar');
  _calendar = new FullCalendar.Calendar(el, {
    initialView: 'dayGridMonth',
    locale: 'zh-tw',
    height: 'auto',
    firstDay: 0, // 週日
    headerToolbar: {
      left: 'prev,next today',
      center: 'title',
      right: 'dayGridMonth,timeGridWeek,listMonth'
    },
    buttonText: { today: '今天', month: '月', week: '週', list: '列表' },
    dayMaxEvents: 4,
    moreLinkText: '+ 還有',

    datesSet: function(info) {
      // 翻頁時重新拉
      loadEventsForRange(info.startStr.substring(0,10), info.endStr.substring(0,10));
    },

    dateClick: function(info) {
      // 點空白日期 → 新增事項，預填日期
      openAddEventModal(info.dateStr);
    },

    eventClick: function(info) {
      openEventDetail(info.event.extendedProps.raw);
    },

    eventContent: function(arg) {
      // 自訂事件渲染：圖示 + 標題 + 子類型 badge
      const ev = arg.event.extendedProps.raw;
      const iconHtml = ev.typeIcon ? `<span class="me-1">${ev.typeIcon}</span>` : '';
      return {
        html: `<div class="d-flex align-items-center" style="overflow:hidden;">
                 ${iconHtml}
                 <span class="text-truncate">${escapeHtml(ev.title || ev.typeName)}</span>
               </div>`
      };
    }
  });
  _calendar.render();
}

// ─────────────────────────────────────────────────────────────
// 2. 載入類型 + 渲染篩選 chips
// ─────────────────────────────────────────────────────────────
async function loadTypesAndChips() {
  try {
    const res = await callAPI('cal_getTypes');
    if (!res.success) throw new Error(res.message);
    _types = res.data;
  } catch (err) {
    document.getElementById('typeChipsContainer').innerHTML =
      `<div class="alert alert-danger m-0 p-2 small">類型載入失敗：${err.message}</div>`;
    return;
  }

  // 算出「可被選為事項類型」的清單：葉子節點 OR 無子的頂層
  _selectableTypes = [];
  _types.flat.forEach(t => {
    const hasChildren = _types.flat.some(c => c.parentTypeId === t.typeId);
    if (!hasChildren) _selectableTypes.push(t);
  });

  renderTypeChips();
  populateTypeSelect();
  // 預設全選
  _activeTypeIds = new Set(_selectableTypes.map(t => t.typeId));
  // 觸發載入當前月事項
  if (_calendar) {
    const view = _calendar.view;
    loadEventsForRange(view.activeStart.toISOString().substring(0,10), view.activeEnd.toISOString().substring(0,10));
  }
}

function renderTypeChips() {
  const container = document.getElementById('typeChipsContainer');
  if (_selectableTypes.length === 0) {
    container.innerHTML = '<span class="text-muted small">尚無可用類型（請先到「事項類型管理」建立）</span>';
    return;
  }
  // 按頂層分組
  const byRoot = {};
  _selectableTypes.forEach(t => {
    let root = t;
    while (root.parentTypeId) {
      root = _types.flat.find(p => p.typeId === root.parentTypeId) || root;
      if (!root.parentTypeId) break;
    }
    const key = root.typeId;
    if (!byRoot[key]) byRoot[key] = { root, items: [] };
    byRoot[key].items.push(t);
  });

  container.innerHTML = Object.values(byRoot).map(group => {
    const chipsHtml = group.items.map(t => {
      const isParent = t.typeId === group.root.typeId;
      const label = isParent ? `${t.icon || ''} ${t['名稱']}` : `${t.icon || ''} ${t['名稱']}`;
      return `<span class="badge filter-chip" style="background-color:${t.color || '#667eea'};color:#fff;"
                data-tid="${t.typeId}" onclick="toggleChip('${t.typeId}')">${label}</span>`;
    }).join('');
    return `<div class="d-flex align-items-center gap-1 me-3">
              <small class="text-muted">${group.root.icon || ''} ${group.root['名稱']}：</small>
              ${chipsHtml}
            </div>`;
  }).join('');
}

function toggleChip(typeId) {
  if (_activeTypeIds.has(typeId)) _activeTypeIds.delete(typeId);
  else _activeTypeIds.add(typeId);
  document.querySelectorAll(`.filter-chip[data-tid="${typeId}"]`).forEach(el => {
    el.classList.toggle('off', !_activeTypeIds.has(typeId));
  });
  // 重新載當前範圍
  const view = _calendar.view;
  loadEventsForRange(view.activeStart.toISOString().substring(0,10), view.activeEnd.toISOString().substring(0,10));
}

function toggleAllChips(on) {
  if (on) _activeTypeIds = new Set(_selectableTypes.map(t => t.typeId));
  else    _activeTypeIds.clear();
  document.querySelectorAll('.filter-chip').forEach(el => {
    el.classList.toggle('off', !_activeTypeIds.has(el.dataset.tid));
  });
  const view = _calendar.view;
  loadEventsForRange(view.activeStart.toISOString().substring(0,10), view.activeEnd.toISOString().substring(0,10));
}

// ─────────────────────────────────────────────────────────────
// 3. 拉事項 → 餵給 FullCalendar
// ─────────────────────────────────────────────────────────────
async function loadEventsForRange(startDate, endDate) {
  try {
    const req = { startDate, endDate };
    if (_activeTypeIds.size > 0 && _activeTypeIds.size < _selectableTypes.length) {
      req.typeIds = Array.from(_activeTypeIds);
    } else if (_activeTypeIds.size === 0) {
      // 全不選 → 清空
      _calendar.removeAllEvents();
      return;
    }
    const res = await callAPI('cal_getEvents', req);
    if (!res.success) throw new Error(res.message);
    const events = res.data || [];

    _calendar.removeAllEvents();
    events.forEach(ev => {
      _calendar.addEvent({
        id: ev.eventId,
        title: ev.title || ev.typeName,
        start: ev.date,
        backgroundColor: ev.typeColor,
        borderColor: ev.typeColor,
        textColor: '#fff',
        extendedProps: { raw: ev }
      });
    });
  } catch (err) {
    console.error('載入事項失敗', err);
  }
}

// ─────────────────────────────────────────────────────────────
// 4. 事項 Modal — 新增 / 編輯
// ─────────────────────────────────────────────────────────────
function populateTypeSelect() {
  const sel = document.getElementById('evf_typeId');
  if (_selectableTypes.length === 0) {
    sel.innerHTML = '<option value="">尚無可用類型，請先到「事項類型管理」建立</option>';
    return;
  }
  // 按頂層分組
  const byRoot = {};
  _selectableTypes.forEach(t => {
    let root = t;
    while (root.parentTypeId) {
      root = _types.flat.find(p => p.typeId === root.parentTypeId) || root;
      if (!root.parentTypeId) break;
    }
    if (!byRoot[root.typeId]) byRoot[root.typeId] = { root, items: [] };
    byRoot[root.typeId].items.push(t);
  });
  let html = '<option value="">-- 選擇類型 --</option>';
  Object.values(byRoot).forEach(group => {
    if (group.items.length === 1 && group.items[0].typeId === group.root.typeId) {
      // 頂層自己就是葉子
      html += `<option value="${group.root.typeId}">${group.root.icon || ''} ${group.root['名稱']}</option>`;
    } else {
      html += `<optgroup label="${group.root.icon || ''} ${group.root['名稱']}">`;
      group.items.forEach(t => {
        html += `<option value="${t.typeId}">${t.icon || ''} ${t['名稱']}</option>`;
      });
      html += '</optgroup>';
    }
  });
  sel.innerHTML = html;
}

function openAddEventModal(dateStr) {
  document.getElementById('eventModalTitle').innerText = '新增事項';
  document.getElementById('evf_eventId').value = '';
  document.getElementById('evf_typeId').value = '';
  document.getElementById('evf_date').value = dateStr || new Date().toISOString().substring(0, 10);
  document.getElementById('evf_title').value = '';
  document.getElementById('evf_fieldsContainer').innerHTML =
    '<div class="text-muted text-center py-4">請先選擇事項類型，下方會顯示對應欄位</div>';
  document.getElementById('evf_deleteBtn').style.display = 'none';
  bootstrap.Modal.getOrCreateInstance(document.getElementById('eventModal')).show();
}

async function openEditEventModal(event) {
  document.getElementById('eventModalTitle').innerText = '編輯事項';
  document.getElementById('evf_eventId').value = event.eventId;
  document.getElementById('evf_typeId').value = event.typeId;
  document.getElementById('evf_date').value = event.date;
  document.getElementById('evf_title').value = event.title || '';
  document.getElementById('evf_deleteBtn').style.display = '';

  // 載入欄位 + 預填值
  await renderFieldsForType(event.typeId, event.values);
  bootstrap.Modal.getOrCreateInstance(document.getElementById('eventModal')).show();
}

async function onTypeChanged() {
  const tid = document.getElementById('evf_typeId').value;
  if (!tid) {
    document.getElementById('evf_fieldsContainer').innerHTML =
      '<div class="text-muted text-center py-4">請先選擇事項類型</div>';
    return;
  }
  await renderFieldsForType(tid, []);
}

async function renderFieldsForType(typeId, existingValues) {
  const container = document.getElementById('evf_fieldsContainer');
  container.innerHTML = '<div class="text-center py-3"><div class="spinner-border spinner-border-sm"></div></div>';

  let fields;
  if (_fieldsByType[typeId]) {
    fields = _fieldsByType[typeId];
  } else {
    try {
      const res = await callAPI('cal_getFields', { typeId });
      if (!res.success) throw new Error(res.message);
      fields = res.data.fields;
      _fieldsByType[typeId] = fields;
    } catch (err) {
      container.innerHTML = `<div class="alert alert-danger">欄位載入失敗：${err.message}</div>`;
      return;
    }
  }

  if (fields.length === 0) {
    container.innerHTML = '<div class="alert alert-light border text-muted">此類型沒有定義任何欄位</div>';
    return;
  }

  // existingValues: array of {fieldId, value}
  const valMap = {};
  (existingValues || []).forEach(v => valMap[v.fieldId] = v.value);

  container.innerHTML = fields.map(f => _renderFieldInput(f, valMap[f.fieldId] || '')).join('');
}

function _renderFieldInput(f, value) {
  const fid = f.fieldId;
  const label = escapeHtml(f['顯示名稱']);
  const req = f.required ? '<span class="text-danger">*</span>' : '';
  const v = escapeAttr(value);

  let inputHtml = '';
  switch (f['欄位類型']) {
    case 'longtext':
      inputHtml = `<textarea class="form-control field-input" rows="3" data-fid="${fid}" data-req="${f.required}">${escapeHtml(value)}</textarea>`;
      break;
    case 'date':
      inputHtml = `<input type="date" class="form-control field-input" value="${v}" data-fid="${fid}" data-req="${f.required}">`;
      break;
    case 'time':
      inputHtml = `<input type="time" class="form-control field-input" value="${v}" data-fid="${fid}" data-req="${f.required}">`;
      break;
    case 'number':
      inputHtml = `<input type="number" class="form-control field-input" value="${v}" data-fid="${fid}" data-req="${f.required}">`;
      break;
    case 'url':
      inputHtml = `<input type="url" class="form-control field-input" value="${v}" data-fid="${fid}" data-req="${f.required}" placeholder="https://...">`;
      break;
    case 'select': {
      const opts = (Array.isArray(f['下拉選項']) ? f['下拉選項'] : []);
      inputHtml = `<select class="form-select field-input" data-fid="${fid}" data-req="${f.required}">
        <option value="">-- 請選擇 --</option>
        ${opts.map(o => `<option value="${escapeAttr(o)}" ${o === value ? 'selected' : ''}>${escapeHtml(o)}</option>`).join('')}
      </select>`;
      break;
    }
    case 'multiselect': {
      const opts = (Array.isArray(f['下拉選項']) ? f['下拉選項'] : []);
      const selected = String(value || '').split(',').map(s => s.trim());
      inputHtml = `<div class="border rounded p-2 bg-light" data-fid="${fid}" data-req="${f.required}" data-multi="1">
        ${opts.map(o => `
          <div class="form-check form-check-inline">
            <input class="form-check-input" type="checkbox" value="${escapeAttr(o)}" id="ms_${fid}_${escapeAttr(o)}" ${selected.includes(o) ? 'checked' : ''}>
            <label class="form-check-label" for="ms_${fid}_${escapeAttr(o)}">${escapeHtml(o)}</label>
          </div>
        `).join('')}
      </div>`;
      break;
    }
    case 'text':
    default:
      inputHtml = `<input type="text" class="form-control field-input" value="${v}" data-fid="${fid}" data-req="${f.required}">`;
  }

  return `<div class="mb-2">
    <label class="form-label fw-bold small mb-1">${label} ${req}</label>
    ${inputHtml}
  </div>`;
}

async function saveEvent() {
  const eventId = document.getElementById('evf_eventId').value;
  const typeId = document.getElementById('evf_typeId').value;
  const date = document.getElementById('evf_date').value;
  const title = document.getElementById('evf_title').value.trim();

  if (!typeId) { alert('請選擇事項類型'); return; }
  if (!date)   { alert('請選擇日期'); return; }

  // 收集欄位值
  const valuesObj = {};
  const inputs = document.querySelectorAll('#evf_fieldsContainer [data-fid]');
  let missing = null;
  inputs.forEach(el => {
    const fid = el.dataset.fid;
    const required = el.dataset.req === 'true';
    let val;
    if (el.dataset.multi === '1') {
      val = Array.from(el.querySelectorAll('input[type="checkbox"]:checked')).map(c => c.value).join(',');
    } else {
      val = el.value;
    }
    if (required && !val.toString().trim()) missing = el;
    if (val) valuesObj[fid] = val;
  });
  if (missing) {
    missing.focus();
    alert('有必填欄位尚未填寫');
    return;
  }

  const data = { typeId, date, title, values: valuesObj };
  try {
    let res;
    if (eventId) {
      data.eventId = eventId;
      res = await callAPI('cal_updateEvent', data);
    } else {
      res = await callAPI('cal_addEvent', data);
    }
    if (!res.success) throw new Error(res.message);
    bootstrap.Modal.getOrCreateInstance(document.getElementById('eventModal')).hide();
    // 重新拉當前範圍
    const view = _calendar.view;
    loadEventsForRange(view.activeStart.toISOString().substring(0,10), view.activeEnd.toISOString().substring(0,10));
  } catch (err) {
    alert('❌ ' + err.message);
  }
}

async function deleteEvent() {
  const eventId = document.getElementById('evf_eventId').value;
  if (!eventId) return;
  if (!confirm('確定刪除此事項？')) return;
  try {
    const res = await callAPI('cal_deleteEvent', { eventId });
    if (!res.success) throw new Error(res.message);
    bootstrap.Modal.getOrCreateInstance(document.getElementById('eventModal')).hide();
    const view = _calendar.view;
    loadEventsForRange(view.activeStart.toISOString().substring(0,10), view.activeEnd.toISOString().substring(0,10));
  } catch (err) {
    alert('❌ ' + err.message);
  }
}

// ─────────────────────────────────────────────────────────────
// 5. 事項詳情 Modal（點月曆 chip 開啟）
// ─────────────────────────────────────────────────────────────
function openEventDetail(event) {
  _currentDetailEvent = event;
  const header = document.getElementById('eventDetailHeader');
  header.style.background = event.typeColor + '20';
  header.style.borderBottom = `3px solid ${event.typeColor}`;
  document.getElementById('eventDetailTitle').innerHTML = `${event.typeIcon || ''} ${escapeHtml(event.title || event.typeName)}`;

  const valuesHtml = event.values && event.values.length > 0
    ? event.values.map(v => `<div class="field-row-display">
        <span class="label">${escapeHtml(v.fieldName)}：</span>
        <span>${v.fieldType === 'longtext' ? escapeHtml(v.value).replace(/\n/g, '<br>') : escapeHtml(v.value)}</span>
      </div>`).join('')
    : '<div class="text-muted text-center py-2">沒有填寫任何欄位</div>';

  document.getElementById('eventDetailBody').innerHTML = `
    <div class="event-meta mb-3">
      📅 <b>${event.date}</b>
      🏷️ <span class="badge" style="background:${event.typeColor};">${event.typeFullName}</span>
    </div>
    ${valuesHtml}
  `;
  bootstrap.Modal.getOrCreateInstance(document.getElementById('eventDetailModal')).show();
}

function editFromDetail() {
  bootstrap.Modal.getOrCreateInstance(document.getElementById('eventDetailModal')).hide();
  setTimeout(() => openEditEventModal(_currentDetailEvent), 300);
}

async function confirmDeleteFromDetail() {
  if (!_currentDetailEvent) return;
  if (!confirm(`確定刪除「${_currentDetailEvent.title || _currentDetailEvent.typeName}」？`)) return;
  try {
    const res = await callAPI('cal_deleteEvent', { eventId: _currentDetailEvent.eventId });
    if (!res.success) throw new Error(res.message);
    bootstrap.Modal.getOrCreateInstance(document.getElementById('eventDetailModal')).hide();
    const view = _calendar.view;
    loadEventsForRange(view.activeStart.toISOString().substring(0,10), view.activeEnd.toISOString().substring(0,10));
  } catch (err) {
    alert('❌ ' + err.message);
  }
}

// ═════════════════════════════════════════════════════════════
// 📅 批量新增（手動）
// ═════════════════════════════════════════════════════════════
let _batchDates = []; // ['2026-01-05', ...]

function openBatchModal() {
  _batchDates = [];
  document.getElementById('batch_startDate').value = '';
  document.getElementById('batch_endDate').value = '';
  document.querySelectorAll('.batch-wd').forEach(b => b.classList.remove('active', 'btn-info', 'text-white'));
  document.getElementById('batch_dateChips').innerHTML = '<span class="text-muted small">尚未產生</span>';
  document.getElementById('batch_summary').innerText = '';
  // 重用單筆 modal 的類型 select
  document.getElementById('batch_typeId').innerHTML = document.getElementById('evf_typeId').innerHTML;
  document.getElementById('batch_typeId').value = '';
  document.getElementById('batch_fieldsContainer').innerHTML = '<div class="text-muted text-center py-3">請先選擇類型</div>';

  // 星期幾按鈕互動
  document.querySelectorAll('.batch-wd').forEach(btn => {
    btn.onclick = () => {
      btn.classList.toggle('active');
      btn.classList.toggle('btn-info');
      btn.classList.toggle('text-white');
    };
  });

  bootstrap.Modal.getOrCreateInstance(document.getElementById('batchModal')).show();
}

function batchGenerateDates() {
  const start = document.getElementById('batch_startDate').value;
  const end = document.getElementById('batch_endDate').value;
  if (!start || !end) { alert('請選擇起訖日期'); return; }
  if (start > end) { alert('開始日期不可大於結束日期'); return; }

  const selectedWds = Array.from(document.querySelectorAll('.batch-wd.active')).map(b => parseInt(b.dataset.wd));
  const sd = new Date(start), ed = new Date(end);
  const dates = [];
  for (let d = new Date(sd); d <= ed; d.setDate(d.getDate() + 1)) {
    if (selectedWds.length === 0 || selectedWds.includes(d.getDay())) {
      const y = d.getFullYear();
      const m = String(d.getMonth() + 1).padStart(2, '0');
      const day = String(d.getDate()).padStart(2, '0');
      dates.push(`${y}-${m}-${day}`);
    }
  }
  _batchDates = dates;
  renderBatchDateChips();
}

function renderBatchDateChips() {
  const c = document.getElementById('batch_dateChips');
  if (_batchDates.length === 0) {
    c.innerHTML = '<span class="text-muted small">尚未產生</span>';
    document.getElementById('batch_summary').innerText = '';
    return;
  }
  c.innerHTML = _batchDates.map((d, i) =>
    `<span class="badge bg-info text-white">${d} <button class="btn-close btn-close-white ms-1" style="font-size:.5rem;" onclick="removeBatchDate(${i})"></button></span>`
  ).join('');
  document.getElementById('batch_summary').innerText = `將建立 ${_batchDates.length} 筆事項`;
}

function removeBatchDate(idx) {
  _batchDates.splice(idx, 1);
  renderBatchDateChips();
}

async function onBatchTypeChanged() {
  const tid = document.getElementById('batch_typeId').value;
  const container = document.getElementById('batch_fieldsContainer');
  if (!tid) {
    container.innerHTML = '<div class="text-muted text-center py-3">請先選擇類型</div>';
    return;
  }
  container.innerHTML = '<div class="text-center py-3"><div class="spinner-border spinner-border-sm"></div></div>';
  let fields;
  if (_fieldsByType[tid]) fields = _fieldsByType[tid];
  else {
    try {
      const res = await callAPI('cal_getFields', { typeId: tid });
      if (!res.success) throw new Error(res.message);
      fields = res.data.fields;
      _fieldsByType[tid] = fields;
    } catch (err) {
      container.innerHTML = `<div class="alert alert-danger">${err.message}</div>`;
      return;
    }
  }
  if (fields.length === 0) {
    container.innerHTML = '<div class="alert alert-light border text-muted">此類型沒有欄位</div>';
    return;
  }
  // 給 batch 用，特別前綴 fid 避免與單筆 modal 衝突
  container.innerHTML = fields.map(f => _renderFieldInput(f, ''))
    .join('').replace(/data-fid="/g, 'data-bfid="');
}

async function confirmBatchAdd() {
  if (_batchDates.length === 0) { alert('請先產生日期'); return; }
  const typeId = document.getElementById('batch_typeId').value;
  if (!typeId) { alert('請選擇類型'); return; }

  // 收集欄位值（套用到每一筆）
  const valuesObj = {};
  const inputs = document.querySelectorAll('#batch_fieldsContainer [data-bfid]');
  let missing = null;
  inputs.forEach(el => {
    const fid = el.dataset.bfid;
    const required = el.dataset.req === 'true';
    let val;
    if (el.dataset.multi === '1') {
      val = Array.from(el.querySelectorAll('input[type="checkbox"]:checked')).map(c => c.value).join(',');
    } else {
      val = el.value;
    }
    if (required && !val.toString().trim()) missing = el;
    if (val) valuesObj[fid] = val;
  });
  if (missing) { missing.focus(); alert('有必填欄位尚未填寫'); return; }

  const events = _batchDates.map(d => ({ typeId, date: d, values: valuesObj, title: '' }));

  try {
    const res = await callAPI('cal_addEventsBatch', { events });
    if (!res.success) throw new Error(res.message || '建立失敗');
    alert(`✅ ${res.message}`);
    bootstrap.Modal.getOrCreateInstance(document.getElementById('batchModal')).hide();
    const view = _calendar.view;
    loadEventsForRange(view.activeStart.toISOString().substring(0,10), view.activeEnd.toISOString().substring(0,10));
  } catch (err) {
    alert('❌ ' + err.message);
  }
}

// ═════════════════════════════════════════════════════════════
// 🤖 AI 解析
// ═════════════════════════════════════════════════════════════
let _aiParseResult = null;

function openAiModal() {
  // 只列頂層類型（AI 會自動判斷子類型）
  const sel = document.getElementById('ai_rootTypeId');
  const roots = _types.tree;
  sel.innerHTML = '<option value="">-- 選擇頂層類型 --</option>' +
    roots.map(t => `<option value="${t.typeId}">${t.icon || ''} ${t['名稱']}</option>`).join('');
  document.getElementById('ai_rawText').value = '';
  document.getElementById('ai_resultPreview').innerHTML = '';
  document.getElementById('ai_confirmBtn').style.display = 'none';
  _aiParseResult = null;
  bootstrap.Modal.getOrCreateInstance(document.getElementById('aiModal')).show();
}

async function runAiParse() {
  const rootTypeId = document.getElementById('ai_rootTypeId').value;
  const rawText = document.getElementById('ai_rawText').value.trim();
  if (!rootTypeId) { alert('請選擇預期類型'); return; }
  if (!rawText) { alert('請貼上要解析的文字'); return; }

  const preview = document.getElementById('ai_resultPreview');
  preview.innerHTML = '<div class="text-center py-3"><div class="spinner-border text-warning"></div><div class="small text-muted mt-2">AI 解析中...（可能需要 5-15 秒）</div></div>';

  try {
    const res = await callAPI('cal_aiParseForType', { rootTypeId, rawText, allowMultiple: true });
    if (!res.success) throw new Error(res.message || 'AI 解析失敗');
    _aiParseResult = res;
    renderAiPreview(res);
  } catch (err) {
    preview.innerHTML = `<div class="alert alert-danger">❌ ${err.message}</div>`;
  }
}

function renderAiPreview(res) {
  const events = res.events || [];
  if (events.length === 0) {
    document.getElementById('ai_resultPreview').innerHTML = '<div class="alert alert-warning">AI 沒有解析出任何事項</div>';
    return;
  }
  let html = `<div class="alert alert-success py-2 mb-2">✅ 解析出 ${events.length} 個事項，請確認後按下方「批量建立」</div>`;
  html += `<div class="table-responsive"><table class="table table-sm table-bordered"><thead class="table-light"><tr>
    <th>#</th><th>日期</th><th>${res.hasSubTypes ? '子類型' : '類型'}</th><th>主要內容</th>
  </tr></thead><tbody>`;
  events.forEach((ev, i) => {
    const dateValid = /^\d{4}-\d{2}-\d{2}$/.test(ev.date);
    const subOk = res.hasSubTypes ? !!ev.subTypeId : true;
    const valuesPreview = Object.entries(ev.values || {}).slice(0,3)
      .map(([k,v]) => escapeHtml(String(v).substring(0,40))).join(' / ');
    html += `<tr ${(!dateValid || !subOk) ? 'class="table-warning"' : ''}>
      <td>${i+1}</td>
      <td>${escapeHtml(ev.date)} ${dateValid ? '' : '<small class="text-danger">⚠ 日期格式錯</small>'}</td>
      <td>${escapeHtml(ev.subTypeName || res.rootTypeName)} ${(res.hasSubTypes && !subOk) ? '<small class="text-danger">⚠ 找不到</small>' : ''}</td>
      <td><small>${valuesPreview}</small></td>
    </tr>`;
  });
  html += '</tbody></table></div>';
  document.getElementById('ai_resultPreview').innerHTML = html;
  document.getElementById('ai_confirmBtn').style.display = '';
}

async function confirmAiBatchAdd() {
  if (!_aiParseResult || !_aiParseResult.events) return;
  const events = _aiParseResult.events
    .filter(ev => /^\d{4}-\d{2}-\d{2}$/.test(ev.date)) // 日期格式對才建立
    .map(ev => ({
      typeId: ev.subTypeId || _aiParseResult.rootTypeId, // 子類型優先
      date: ev.date,
      title: ev.title || '',
      values: ev.values || {}
    }))
    .filter(ev => ev.typeId);

  if (events.length === 0) { alert('沒有有效事項可建立'); return; }

  try {
    const res = await callAPI('cal_addEventsBatch', { events });
    if (!res.success) throw new Error(res.message);
    alert(`✅ ${res.message}`);
    bootstrap.Modal.getOrCreateInstance(document.getElementById('aiModal')).hide();
    const view = _calendar.view;
    loadEventsForRange(view.activeStart.toISOString().substring(0,10), view.activeEnd.toISOString().substring(0,10));
  } catch (err) {
    alert('❌ ' + err.message);
  }
}

// ═════════════════════════════════════════════════════════════
// 📤 Excel 上傳 → 預覽 → 建立
// ═════════════════════════════════════════════════════════════
let _excelImportRows = null; // [{typeId, date, title, values, errors}]

async function handleExcelUpload(ev) {
  const file = ev.target.files[0];
  if (!file) return;

  try {
    const data = await file.arrayBuffer();
    const wb = XLSX.read(data);
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    ev.target.value = ''; // reset，方便重傳同一檔

    if (rows.length === 0) { alert('Excel 沒有資料'); return; }

    // 找到對應類型（用 Sheet 名稱比對頂層類型）
    const rootType = _types.flat.find(t => !t.parentTypeId && t['名稱'] === sheetName);
    if (!rootType) {
      alert(`❌ 找不到對應的類型「${sheetName}」\n請確認 Excel sheet 名稱與某個頂層類型完全相符\n（建議直接使用「欄位管理」→「匯出模板」下載的範本）`);
      return;
    }

    // 取該類型的欄位定義
    let fields;
    if (_fieldsByType[rootType.typeId]) fields = _fieldsByType[rootType.typeId];
    else {
      const res = await callAPI('cal_getFields', { typeId: rootType.typeId });
      if (!res.success) throw new Error(res.message);
      fields = res.data.fields;
      _fieldsByType[rootType.typeId] = fields;
    }
    // 子類型清單（依名稱對應）
    const subTypesByName = {};
    _types.flat.filter(t => t.parentTypeId === rootType.typeId).forEach(t => subTypesByName[t['名稱']] = t.typeId);

    // 欄位名 → fieldId
    const fieldByName = {};
    fields.forEach(f => fieldByName[f['顯示名稱']] = f);
    const requiredFieldNames = fields.filter(f => f.required).map(f => f['顯示名稱']);

    // 逐列轉換
    _excelImportRows = rows.map((row, idx) => {
      const errors = [];
      // 日期
      let date = row['日期'] || row['date'];
      if (date instanceof Date) {
        date = `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,'0')}-${String(date.getDate()).padStart(2,'0')}`;
      } else if (typeof date === 'number') {
        // Excel 序列日期
        const d = XLSX.SSF.parse_date_code(date);
        if (d) date = `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`;
      } else {
        date = String(date || '').trim();
      }
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) errors.push('日期格式錯');

      // 子類型
      let typeId = rootType.typeId;
      const subTypeName = String(row['子類型'] || '').trim();
      if (subTypeName) {
        if (subTypesByName[subTypeName]) typeId = subTypesByName[subTypeName];
        else if (Object.keys(subTypesByName).length > 0) errors.push(`子類型「${subTypeName}」不存在`);
      } else if (Object.keys(subTypesByName).length > 0) {
        // 有子類型但沒指定 — 用 root（提示）
        // 改為錯誤可能太嚴格，這裡只警告（不擋）
      }

      // 標題
      const title = String(row['顯示標題'] || row['標題'] || '').trim();

      // 欄位值
      const values = {};
      Object.entries(row).forEach(([col, v]) => {
        if (['日期', 'date', '子類型', '顯示標題', '標題'].indexOf(col) !== -1) return;
        const f = fieldByName[col];
        if (f && v !== '' && v !== null && v !== undefined) {
          values[f.fieldId] = (v instanceof Date) ? v.toISOString().substring(0,10) : String(v);
        }
      });

      // 必填檢查
      requiredFieldNames.forEach(fn => {
        const f = fieldByName[fn];
        if (f && !values[f.fieldId]) errors.push(`欠必填「${fn}」`);
      });

      return { idx: idx + 2, typeId, date, title, values, errors, subTypeName };
    });

    renderExcelPreview(rootType);
    bootstrap.Modal.getOrCreateInstance(document.getElementById('excelPreviewModal')).show();
  } catch (err) {
    alert('❌ Excel 解析失敗：' + err.message);
  }
}

function renderExcelPreview(rootType) {
  const ok = _excelImportRows.filter(r => r.errors.length === 0).length;
  const bad = _excelImportRows.length - ok;
  document.getElementById('excelPreviewSummary').innerText = `共 ${_excelImportRows.length} 列：可建立 ${ok}，問題 ${bad}`;
  let html = `<div class="alert alert-info py-2 mb-2">📂 對應類型：<b>${escapeHtml(rootType['名稱'])}</b>　🟢 可建立 ${ok} 筆，🟡 問題 ${bad} 筆（有問題的列不會建立）</div>`;
  html += '<div class="table-responsive"><table class="table table-sm table-bordered"><thead class="table-light"><tr>';
  html += '<th>Excel 列</th><th>日期</th><th>子類型</th><th>標題</th><th>欄位值（前 3 個）</th><th>檢查</th></tr></thead><tbody>';
  _excelImportRows.forEach(r => {
    const valuesPreview = Object.entries(r.values).slice(0, 3)
      .map(([fid, v]) => escapeHtml(String(v).substring(0,30))).join(' / ');
    const cls = r.errors.length > 0 ? 'table-warning' : '';
    const errs = r.errors.length > 0
      ? `<span class="text-danger small">⚠ ${r.errors.join('，')}</span>`
      : '<span class="text-success small">✓</span>';
    html += `<tr class="${cls}">
      <td class="text-center">${r.idx}</td>
      <td>${escapeHtml(r.date)}</td>
      <td>${escapeHtml(r.subTypeName || '(root)')}</td>
      <td>${escapeHtml(r.title)}</td>
      <td><small>${valuesPreview}</small></td>
      <td>${errs}</td>
    </tr>`;
  });
  html += '</tbody></table></div>';
  document.getElementById('excelPreviewBody').innerHTML = html;
  document.getElementById('excelConfirmBtn').disabled = ok === 0;
}

async function confirmExcelImport() {
  if (!_excelImportRows) return;
  const events = _excelImportRows
    .filter(r => r.errors.length === 0)
    .map(r => ({ typeId: r.typeId, date: r.date, title: r.title, values: r.values }));
  if (events.length === 0) { alert('沒有可建立的列'); return; }

  try {
    const res = await callAPI('cal_addEventsBatch', { events });
    if (!res.success) throw new Error(res.message);
    alert(`✅ ${res.message}`);
    bootstrap.Modal.getOrCreateInstance(document.getElementById('excelPreviewModal')).hide();
    const view = _calendar.view;
    loadEventsForRange(view.activeStart.toISOString().substring(0,10), view.activeEnd.toISOString().substring(0,10));
  } catch (err) {
    alert('❌ ' + err.message);
  }
}

// ─────────────────────────────────────────────────────────────
// helpers
// ─────────────────────────────────────────────────────────────
function escapeHtml(s) {
  return String(s == null ? '' : s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}
function escapeAttr(s) {
  return String(s == null ? '' : s).replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}
