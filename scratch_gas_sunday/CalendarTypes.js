/**
 * 教會行事曆 - 事項類型 CRUD
 *
 * 階層結構：parentTypeId 空字串 = 頂層；非空 = 子層
 * 密碼：頂層類型可設密碼；子類型自動繼承父類型密碼
 *      （前端 verify 時用 cal_verifyTypePassword 通常傳「頂層 typeId」）
 */

// 取得所有類型（含 children 陣列，方便前端渲染樹）
function cal_getTypes() {
  const all = _cal_readAll(_CAL_SHEET.TYPES);

  // 把 boolean 字串轉回實際 bool
  all.forEach(t => {
    t.syncToAttendance = String(t.syncToAttendance).toUpperCase() === 'TRUE';
    t.syncToMinistry   = String(t.syncToMinistry).toUpperCase()   === 'TRUE';
    t.syncToWorship    = String(t.syncToWorship).toUpperCase()    === 'TRUE';
    t.hidden           = String(t.hidden).toUpperCase()           === 'TRUE';
    t.hasPassword      = !!(t.password && String(t.password).trim());
    delete t.password; // ⚠️ 永遠不要把密碼回傳前端
    // excludedFieldIds: 字串 JSON → array
    try {
      t.excludedFieldIds = t.excludedFieldIds ? JSON.parse(t.excludedFieldIds) : [];
    } catch (e) { t.excludedFieldIds = []; }
    if (!Array.isArray(t.excludedFieldIds)) t.excludedFieldIds = [];
  });

  // 排序：sortOrder 升冪
  all.sort((a, b) => (Number(a.sortOrder) || 0) - (Number(b.sortOrder) || 0));

  // 建立 children
  const byId = {};
  all.forEach(t => { t.children = []; byId[t.typeId] = t; });
  const roots = [];
  all.forEach(t => {
    if (t.parentTypeId && byId[t.parentTypeId]) {
      byId[t.parentTypeId].children.push(t);
    } else {
      roots.push(t);
    }
  });

  return { types: roots, flat: all };
}

// 新增類型
function cal_addType(data) {
  if (!data || !data.name) throw new Error('類型名稱必填');
  // 子類型不要設密碼（會被父覆蓋）
  if (data.parentTypeId) data.password = '';
  const id = _cal_addTypeRow({
    parentTypeId:     data.parentTypeId || '',
    name:             String(data.name).trim(),
    icon:             data.icon || '',
    color:            data.color || '#667eea',
    sortOrder:        Number(data.sortOrder) || _cal_nextSortOrder(data.parentTypeId || ''),
    syncToAttendance: !!data.syncToAttendance,
    syncToMinistry:   !!data.syncToMinistry,
    syncToWorship:    !!data.syncToWorship,
    password:         data.password || '',
    hidden:           !!data.hidden
  });
  return { success: true, typeId: id, message: '已新增類型' };
}

// 更新類型
function cal_updateType(data) {
  if (!data || !data.typeId) throw new Error('typeId 必填');
  const sh = _cal_getSheet(_CAL_SHEET.TYPES);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const colIdx = {};
  headers.forEach((h, i) => colIdx[h] = i);

  for (let r = 1; r < values.length; r++) {
    if (values[r][colIdx['typeId']] === data.typeId) {
      // 子類型不能改 parentTypeId 為自己或自己的後代（避免循環）
      if (data.parentTypeId !== undefined) {
        if (data.parentTypeId === data.typeId) throw new Error('不能把自己設為父類型');
        if (_cal_isDescendant(data.typeId, data.parentTypeId)) throw new Error('不能把後代設為父類型');
      }

      const setIf = (key, val, transform) => {
        if (val !== undefined && colIdx[key] !== undefined) {
          values[r][colIdx[key]] = transform ? transform(val) : val;
        }
      };
      setIf('parentTypeId', data.parentTypeId);
      setIf('名稱', data.name && String(data.name).trim());
      setIf('icon', data.icon);
      setIf('color', data.color);
      setIf('sortOrder', data.sortOrder, v => Number(v) || 0);
      setIf('syncToAttendance', data.syncToAttendance, v => v ? 'TRUE' : 'FALSE');
      setIf('syncToMinistry',   data.syncToMinistry,   v => v ? 'TRUE' : 'FALSE');
      setIf('syncToWorship',    data.syncToWorship,    v => v ? 'TRUE' : 'FALSE');
      setIf('hidden',           data.hidden,           v => v ? 'TRUE' : 'FALSE');
      // 排除欄位清單（子類型用）
      if (data.excludedFieldIds !== undefined && colIdx['excludedFieldIds'] !== undefined) {
        values[r][colIdx['excludedFieldIds']] =
          Array.isArray(data.excludedFieldIds) ? JSON.stringify(data.excludedFieldIds) : String(data.excludedFieldIds || '[]');
      }
      // 密碼：空字串 = 清除；undefined = 不動
      if (data.password !== undefined) {
        values[r][colIdx['password']] = data.password || '';
      }

      sh.getRange(r + 1, 1, 1, values[r].length).setValues([values[r]]);
      return { success: true, message: '已更新' };
    }
  }
  throw new Error('找不到該類型');
}

// 刪除類型（連同子類型 + 欄位定義 + 事項一起刪）
function cal_deleteType(data) {
  if (!data || !data.typeId) throw new Error('typeId 必填');

  const typesAll = _cal_readAll(_CAL_SHEET.TYPES);
  const toDelete = _cal_collectSubtree(data.typeId, typesAll);
  if (toDelete.length === 0) throw new Error('找不到該類型');

  // 1. 刪事項 + 對應事項欄位值
  const eventsAll = _cal_readAll(_CAL_SHEET.EVENTS);
  const eventsToDelete = eventsAll.filter(e => toDelete.includes(e.typeId));
  const eventIdsToDelete = new Set(eventsToDelete.map(e => e.eventId));
  if (eventsToDelete.length > 0) {
    _cal_deleteRowsByIdSet(_CAL_SHEET.EVENTS, 'eventId', eventIdsToDelete);
    _cal_deleteRowsByIdSet(_CAL_SHEET.VALUES, 'eventId', eventIdsToDelete);
  }

  // 2. 刪欄位定義
  const fieldsAll = _cal_readAll(_CAL_SHEET.FIELDS);
  const fieldIdsToDelete = new Set(
    fieldsAll.filter(f => toDelete.includes(f.typeId)).map(f => f.fieldId)
  );
  if (fieldIdsToDelete.size > 0) {
    _cal_deleteRowsByIdSet(_CAL_SHEET.FIELDS, 'fieldId', fieldIdsToDelete);
  }

  // 3. 刪類型
  _cal_deleteRowsByIdSet(_CAL_SHEET.TYPES, 'typeId', new Set(toDelete));

  return {
    success: true,
    message: `已刪除（含 ${toDelete.length} 個類型、${fieldIdsToDelete.size} 個欄位、${eventsToDelete.length} 個事項）`
  };
}

// 驗證類型密碼（前端拿密碼比對。子類型 → 找頂層父類型的密碼）
function cal_verifyTypePassword(data) {
  if (!data || !data.typeId) return { success: false, message: '缺少 typeId' };
  const all = _cal_readAll(_CAL_SHEET.TYPES);
  const targetType = all.find(t => t.typeId === data.typeId);
  if (!targetType) return { success: false, message: '查無此類型' };

  // 找頂層
  let cur = targetType;
  while (cur.parentTypeId) {
    const parent = all.find(t => t.typeId === cur.parentTypeId);
    if (!parent) break;
    cur = parent;
  }
  const truePassword = cur.password || '';
  if (!truePassword) return { success: true, message: '此類型未設密碼' };

  if (String(data.password || '').trim() === String(truePassword).trim()) {
    return { success: true, message: '密碼正確' };
  }
  return { success: false, message: '密碼錯誤' };
}

// ─────────────────────────────────────────────────────────────
//  helpers
// ─────────────────────────────────────────────────────────────
function _cal_nextSortOrder(parentTypeId) {
  const all = _cal_readAll(_CAL_SHEET.TYPES);
  const siblings = all.filter(t => (t.parentTypeId || '') === (parentTypeId || ''));
  const max = siblings.reduce((m, t) => Math.max(m, Number(t.sortOrder) || 0), 0);
  return max + 1;
}

function _cal_collectSubtree(rootId, allTypes) {
  // 廣度優先收集 root + 所有後代的 typeId
  const result = [];
  const queue = [rootId];
  while (queue.length > 0) {
    const cur = queue.shift();
    result.push(cur);
    allTypes.filter(t => t.parentTypeId === cur).forEach(c => queue.push(c.typeId));
  }
  return result;
}

function _cal_isDescendant(ancestorId, candidateId) {
  if (!candidateId) return false;
  const all = _cal_readAll(_CAL_SHEET.TYPES);
  let cur = all.find(t => t.typeId === candidateId);
  while (cur && cur.parentTypeId) {
    if (cur.parentTypeId === ancestorId) return true;
    cur = all.find(t => t.typeId === cur.parentTypeId);
  }
  return false;
}

function _cal_deleteRowsByIdSet(sheetName, idColName, idSet) {
  const sh = _cal_getSheet(sheetName);
  if (sh.getLastRow() <= 1) return;
  const values = sh.getDataRange().getValues();
  const idIdx = values[0].indexOf(idColName);
  if (idIdx === -1) return;

  // 倒著刪（不會影響上面的 row 編號）
  for (let r = values.length - 1; r >= 1; r--) {
    if (idSet.has(values[r][idIdx])) {
      sh.deleteRow(r + 1);
    }
  }
}
