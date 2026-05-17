/**
 * 事項類型管理 - Phase 1 前端腳本
 *
 * 使用中央 churchAPI（config.js 提供）
 */

let _calTypesFlat = []; // 平坦清單（含 children）
let _calTypesTree = []; // 樹（root 陣列）
let _currentFieldsRootTypeId = null;
let _currentFieldsList = [];

async function callAPI(action, data) {
  if (typeof window.churchAPI !== 'function') {
    throw new Error('config.js 尚未載入');
  }
  return await window.churchAPI(action, data || {});
}

window.addEventListener('DOMContentLoaded', loadTypes);

async function loadTypes() {
  const container = document.getElementById('typesContainer');
  container.innerHTML = `<div class="empty-hint"><div class="spinner-border text-primary mb-2"></div><div>載入中...</div></div>`;
  try {
    const res = await callAPI('cal_getTypes');
    if (!res || !res.success) throw new Error((res && res.message) || '載入失敗');
    _calTypesTree = res.data.types;
    _calTypesFlat = res.data.flat;
    renderTypes();
  } catch (err) {
    container.innerHTML = `<div class="alert alert-danger">❌ 載入失敗：${err.message}
      <hr>若為「找不到分頁」類型錯誤，請先進到 GAS 編輯器手動執行一次 <code>cal_setupSchema</code> 即可建立分頁。</div>`;
  }
}

function renderTypes() {
  const container = document.getElementById('typesContainer');
  if (_calTypesTree.length === 0) {
    container.innerHTML = `<div class="empty-hint">
      還沒有任何類型。<br>
      <button class="btn btn-primary mt-3" onclick="openAddTypeModal('')">＋ 新增第一個頂層類型</button>
    </div>`;
    return;
  }

  container.innerHTML = _calTypesTree.map(t => _renderTypeNode(t, false)).join('');
}

function _renderTypeNode(type, isChild) {
  const syncPills = [
    type.syncToAttendance ? '<span class="pill sync-pill me-1">📋 主日點名</span>' : '',
    type.syncToMinistry   ? '<span class="pill sync-pill me-1">🧑‍🤝‍🧑 事工</span>'   : '',
    type.syncToWorship    ? '<span class="pill sync-pill me-1">🎵 敬拜團</span>'   : ''
  ].join('');

  const lockPill   = type.hasPassword ? '<span class="pill lock-pill me-1">🔒 有密碼</span>' : '';
  const hiddenPill = type.hidden ? '<span class="pill hidden-pill me-1">👁️ 隱藏</span>' : '';

  const childRender = (type.children && type.children.length > 0)
    ? type.children.map(c => _renderTypeNode(c, true)).join('')
    : '';

  return `
    <div class="card type-card mb-2 ${isChild ? 'type-child' : ''}" style="border-left-color: ${type.color || '#667eea'};">
      <div class="card-body py-2 px-3">
        <div class="d-flex align-items-center flex-wrap gap-2">
          <div class="fs-5 fw-bold" style="color:${type.color || '#333'};">
            ${type.icon || ''} ${type['名稱']}
          </div>
          ${isChild ? '<span class="badge bg-light text-muted">子類型</span>' : ''}
          <div class="ms-auto d-flex gap-1 flex-wrap">
            ${syncPills}${lockPill}${hiddenPill}
          </div>
        </div>
        <div class="mt-2 d-flex gap-2 flex-wrap">
          ${!isChild ? `<button class="btn btn-sm btn-outline-success" onclick="openAddTypeModal('${type.typeId}')">＋ 新增子類型</button>` : ''}
          ${!isChild ? `<button class="btn btn-sm btn-outline-primary" onclick="openFieldsModal('${type.typeId}', '${escapeAttr(type['名稱'])}')">📝 欄位管理</button>` : ''}
          <button class="btn btn-sm btn-outline-secondary" onclick='openEditTypeModal(${JSON.stringify(type).replace(/'/g, "&#39;")})'>✏️ 編輯</button>
          <button class="btn btn-sm btn-outline-danger" onclick="confirmDeleteType('${type.typeId}', '${escapeAttr(type['名稱'])}')">🗑️ 刪除</button>
        </div>
      </div>
    </div>
    ${childRender}
  `;
}

function escapeAttr(s) {
  return String(s || '').replace(/'/g, '&#39;').replace(/"/g, '&quot;');
}

// ─── 類型 Modal ───
function openAddTypeModal(parentTypeId) {
  _resetTypeForm();
  document.getElementById('typeModalTitle').innerText = parentTypeId ? '新增子類型' : '新增頂層類型';
  document.getElementById('typeForm_parentTypeId').value = parentTypeId || '';
  // 子類型不顯示「同步/密碼/隱藏」區塊（繼承父）
  document.getElementById('typeForm_parentOnlySection').style.display = parentTypeId ? 'none' : '';
  bootstrap.Modal.getOrCreateInstance(document.getElementById('typeModal')).show();
}

function openEditTypeModal(type) {
  _resetTypeForm();
  document.getElementById('typeModalTitle').innerText = '編輯類型：' + (type['名稱'] || '');
  document.getElementById('typeForm_typeId').value = type.typeId || '';
  document.getElementById('typeForm_parentTypeId').value = type.parentTypeId || '';
  document.getElementById('typeForm_name').value = type['名稱'] || '';
  document.getElementById('typeForm_icon').value = type.icon || '';
  document.getElementById('typeForm_color').value = type.color || '#667eea';
  document.getElementById('typeForm_sortOrder').value = type.sortOrder || 0;
  document.getElementById('typeForm_syncToAttendance').checked = !!type.syncToAttendance;
  document.getElementById('typeForm_syncToMinistry').checked = !!type.syncToMinistry;
  document.getElementById('typeForm_syncToWorship').checked = !!type.syncToWorship;
  document.getElementById('typeForm_hidden').checked = !!type.hidden;
  // 編輯時不顯示密碼（避免被偷看），使用者要改才填新值
  document.getElementById('typeForm_password').value = '';
  document.getElementById('typeForm_password').placeholder = type.hasPassword ? '已設密碼（留空 = 不改 / 輸入空白以外字串才會更新）' : '留空 = 公開';
  document.getElementById('typeForm_parentOnlySection').style.display = type.parentTypeId ? 'none' : '';
  bootstrap.Modal.getOrCreateInstance(document.getElementById('typeModal')).show();
}

function _resetTypeForm() {
  ['typeForm_typeId','typeForm_parentTypeId','typeForm_name','typeForm_icon','typeForm_sortOrder','typeForm_password']
    .forEach(id => document.getElementById(id).value = '');
  document.getElementById('typeForm_color').value = '#667eea';
  ['typeForm_syncToAttendance','typeForm_syncToMinistry','typeForm_syncToWorship','typeForm_hidden']
    .forEach(id => document.getElementById(id).checked = false);
  document.getElementById('typeForm_password').placeholder = '留空 = 公開';
}

async function saveType() {
  const id = document.getElementById('typeForm_typeId').value;
  const isEdit = !!id;
  const data = {
    typeId: id || undefined,
    parentTypeId: document.getElementById('typeForm_parentTypeId').value,
    name: document.getElementById('typeForm_name').value.trim(),
    icon: document.getElementById('typeForm_icon').value.trim(),
    color: document.getElementById('typeForm_color').value,
    sortOrder: parseInt(document.getElementById('typeForm_sortOrder').value) || 0,
    syncToAttendance: document.getElementById('typeForm_syncToAttendance').checked,
    syncToMinistry:   document.getElementById('typeForm_syncToMinistry').checked,
    syncToWorship:    document.getElementById('typeForm_syncToWorship').checked,
    hidden:           document.getElementById('typeForm_hidden').checked
  };
  // 密碼：編輯時若空白則不改；新增則直接套用
  const pwd = document.getElementById('typeForm_password').value;
  if (isEdit) {
    if (pwd.trim() !== '') data.password = pwd; // 含空白也算改成空
  } else {
    if (pwd) data.password = pwd;
  }

  if (!data.name) { alert('請輸入名稱'); return; }

  try {
    const res = await callAPI(isEdit ? 'cal_updateType' : 'cal_addType', data);
    if (!res.success) throw new Error(res.message || '失敗');
    bootstrap.Modal.getOrCreateInstance(document.getElementById('typeModal')).hide();
    await loadTypes();
  } catch (err) {
    alert('❌ ' + err.message);
  }
}

async function confirmDeleteType(typeId, name) {
  if (!confirm(`確定要刪除「${name}」嗎？\n⚠️ 連同子類型、欄位、事項都會一起刪除（不可復原）`)) return;
  try {
    const res = await callAPI('cal_deleteType', { typeId });
    if (!res.success) throw new Error(res.message);
    alert('✅ ' + res.message);
    await loadTypes();
  } catch (err) {
    alert('❌ ' + err.message);
  }
}

// ─── 欄位 Modal ───
async function openFieldsModal(typeId, typeName) {
  document.getElementById('fieldsModalTypeName').innerText = typeName;
  _currentFieldsRootTypeId = typeId;
  document.getElementById('fieldsList').innerHTML = '<div class="text-center p-4"><div class="spinner-border spinner-border-sm"></div></div>';
  bootstrap.Modal.getOrCreateInstance(document.getElementById('fieldsModal')).show();
  try {
    const res = await callAPI('cal_getFields', { typeId });
    if (!res.success) throw new Error(res.message);
    _currentFieldsList = res.data.fields;
    _currentFieldsRootTypeId = res.data.rootTypeId;
    renderFieldsList();
  } catch (err) {
    document.getElementById('fieldsList').innerHTML = `<div class="alert alert-danger">❌ ${err.message}</div>`;
  }
}

function renderFieldsList() {
  const container = document.getElementById('fieldsList');
  if (_currentFieldsList.length === 0) {
    container.innerHTML = '<div class="text-muted text-center py-4">此類型還沒有欄位</div>';
    return;
  }
  container.innerHTML = _currentFieldsList.map(f => _renderFieldRow(f)).join('');
}

function _renderFieldRow(f) {
  const fid = f.fieldId;
  const opts = Array.isArray(f['下拉選項']) ? f['下拉選項'].join('，') : '';
  return `
    <div class="field-row" data-fid="${fid}">
      <div class="row g-2 align-items-center">
        <div class="col-md-3">
          <input type="text" class="form-control form-control-sm" value="${escapeAttr(f['顯示名稱'])}" data-k="name">
        </div>
        <div class="col-md-3">
          <select class="form-select form-select-sm" data-k="type">
            ${['text','longtext','date','time','select','multiselect','number','url']
              .map(t => `<option value="${t}" ${t === f['欄位類型'] ? 'selected' : ''}>${_fieldTypeLabel(t)}</option>`).join('')}
          </select>
        </div>
        <div class="col-md-3">
          <input type="text" class="form-control form-control-sm" placeholder="下拉選項用，逗號分隔" value="${escapeAttr(opts)}" data-k="options">
        </div>
        <div class="col-md-1">
          <div class="form-check pt-1">
            <input type="checkbox" class="form-check-input" ${f.required ? 'checked' : ''} data-k="required" title="必填">
            <label class="form-check-label small">必</label>
          </div>
        </div>
        <div class="col-md-2 text-end">
          <button class="btn btn-sm btn-success me-1" onclick="saveFieldRow('${fid}')">💾</button>
          <button class="btn btn-sm btn-outline-danger" onclick="deleteFieldRow('${fid}', '${escapeAttr(f['顯示名稱'])}')">🗑️</button>
        </div>
      </div>
    </div>
  `;
}

function _fieldTypeLabel(t) {
  return { text:'文字', longtext:'長文字', date:'日期', time:'時間', select:'單選', multiselect:'多選', number:'數字', url:'網址' }[t] || t;
}

function openAddFieldRow() {
  if (!_currentFieldsRootTypeId) return;
  const tempId = '_new_' + Date.now();
  _currentFieldsList.push({
    fieldId: tempId, typeId: _currentFieldsRootTypeId,
    '顯示名稱': '', '欄位類型': 'text', required: false, '下拉選項': [], sortOrder: _currentFieldsList.length + 1
  });
  renderFieldsList();
  // 焦點到新列
  setTimeout(() => {
    const row = document.querySelector(`.field-row[data-fid="${tempId}"]`);
    if (row) row.querySelector('input[data-k="name"]').focus();
  }, 50);
}

async function saveFieldRow(fid) {
  const row = document.querySelector(`.field-row[data-fid="${fid}"]`);
  if (!row) return;
  const get = k => row.querySelector(`[data-k="${k}"]`);
  const data = {
    name: get('name').value.trim(),
    type: get('type').value,
    required: get('required').checked,
    options: get('options').value.trim() // 後端會自動解析逗號分隔
  };
  if (!data.name) { alert('請輸入欄位名稱'); return; }

  try {
    let res;
    if (String(fid).startsWith('_new_')) {
      data.typeId = _currentFieldsRootTypeId;
      res = await callAPI('cal_addField', data);
      if (!res.success) throw new Error(res.message);
    } else {
      data.fieldId = fid;
      res = await callAPI('cal_updateField', data);
      if (!res.success) throw new Error(res.message);
    }
    // 重新拉
    const r = await callAPI('cal_getFields', { typeId: _currentFieldsRootTypeId });
    _currentFieldsList = r.data.fields;
    renderFieldsList();
  } catch (err) {
    alert('❌ ' + err.message);
  }
}

async function deleteFieldRow(fid, name) {
  if (String(fid).startsWith('_new_')) {
    _currentFieldsList = _currentFieldsList.filter(f => f.fieldId !== fid);
    renderFieldsList();
    return;
  }
  if (!confirm(`確定刪除欄位「${name}」？\n⚠️ 所有事項中此欄位的值也會一併清除`)) return;
  try {
    const res = await callAPI('cal_deleteField', { fieldId: fid });
    if (!res.success) throw new Error(res.message);
    _currentFieldsList = _currentFieldsList.filter(f => f.fieldId !== fid);
    renderFieldsList();
  } catch (err) {
    alert('❌ ' + err.message);
  }
}

// ─── 遷移 ───
async function runMigration() {
  if (!confirm(`即將進行：
1. 把舊「聚會資料 / 事工細項 / 講道資訊」備份成 _Backup_* 分頁
2. 依新結構轉成事項

舊資料不會被刪除，可以重複跑。是否繼續？`)) return;
  try {
    const res = await callAPI('cal_migrateOldData');
    if (!res.success) throw new Error(res.message);
    alert(res.message + '\n\n備份分頁：\n' + (res.backups || []).join('\n') + '\n\n' + (res.note || ''));
  } catch (err) {
    alert('❌ ' + err.message);
  }
}
