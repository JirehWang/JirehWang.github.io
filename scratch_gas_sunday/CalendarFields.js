/**
 * 教會行事曆 - 欄位定義 CRUD
 *
 * 重點：
 *   欄位定義在「頂層類型」上（取消子類型分別設定，避免欄位四散）。
 *   傳入子類型 typeId 時，會自動往上找頂層，把欄位掛到頂層。
 *   讀取時用 cal_getFields(typeId) 也會自動往上找頂層回傳。
 */

// ─────────────────────────────────────────────────────────────
//  取得指定類型的「有效欄位」清單
//
//  邏輯：
//    - 若 typeId 是頂層 → 回傳該頂層的所有欄位（own）
//    - 若 typeId 是子類型 → (繼承自頂層的欄位 - 此子類型已排除) + 此子類型專屬欄位
//
//  回傳：
//    {
//      rootTypeId, subTypeId(null if root),
//      fields: [...]      // 給「新增/編輯事項表單」用的有效欄位（排序後）
//      inheritedFields: [...] // 給欄位管理 UI 顯示「繼承自父」區塊（含 excluded 標記）
//      ownFields: [...]   // 給欄位管理 UI 顯示「專屬」區塊
//      excludedFieldIds: [...] // 該子類型已排除的繼承欄位 ID
//    }
// ─────────────────────────────────────────────────────────────
function cal_getFields(data) {
  const typeId = data && data.typeId;
  if (!typeId) throw new Error('typeId 必填');

  const allTypes = _cal_readAll(_CAL_SHEET.TYPES);
  const targetType = allTypes.find(t => t.typeId === typeId);
  if (!targetType) throw new Error('查無類型：' + typeId);

  let rootId, subTypeId;
  if (targetType.parentTypeId) {
    rootId = _cal_findRootTypeId(typeId);
    subTypeId = typeId;
  } else {
    rootId = typeId;
    subTypeId = null;
  }

  const allFields = _cal_readAll(_CAL_SHEET.FIELDS);
  const mapField = f => ({
    ...f,
    required: String(f['是否必填']).toUpperCase() === 'TRUE',
    下拉選項: _cal_safeParseOptions(f['下拉選項'])
  });
  const sortByOrder = (a, b) => (Number(a.sortOrder) || 0) - (Number(b.sortOrder) || 0);

  // root 的所有欄位
  const rootFields = allFields.filter(f => f.typeId === rootId).map(mapField);

  if (!subTypeId) {
    // 頂層：own = root
    return {
      rootTypeId: rootId,
      subTypeId: null,
      fields:           rootFields.slice().sort(sortByOrder),
      inheritedFields:  [],
      ownFields:        rootFields.slice().sort(sortByOrder),
      excludedFieldIds: []
    };
  }

  // 子類型：解析該子類型的排除清單
  let excludedIds = [];
  try {
    excludedIds = targetType.excludedFieldIds ? JSON.parse(targetType.excludedFieldIds) : [];
  } catch (e) {}
  if (!Array.isArray(excludedIds)) excludedIds = [];

  // 繼承欄位（標 source = inherited，excluded 標記用）
  const inheritedFields = rootFields.map(f => ({
    ...f,
    source: 'inherited',
    excluded: excludedIds.indexOf(f.fieldId) !== -1
  }));

  // 子類型專屬欄位
  const ownFields = allFields
    .filter(f => f.typeId === subTypeId)
    .map(mapField)
    .map(f => ({ ...f, source: 'own' }));

  // 「給表單用」的有效欄位 = 沒排除的繼承 + 全部專屬
  const effective = inheritedFields
    .filter(f => !f.excluded)
    .concat(ownFields)
    .sort(sortByOrder);

  return {
    rootTypeId: rootId,
    subTypeId: subTypeId,
    fields: effective,
    inheritedFields: inheritedFields.sort(sortByOrder),
    ownFields: ownFields.sort(sortByOrder),
    excludedFieldIds: excludedIds
  };
}

// 新增欄位（typeId 可以是頂層或子類型）
// - 頂層：欄位掛在頂層，所有子類型自動繼承
// - 子類型：欄位為該子類型專屬，不影響其他子類型 / 頂層
function cal_addField(data) {
  if (!data || !data.typeId) throw new Error('typeId 必填');
  if (!data.name)            throw new Error('欄位名稱必填');
  const ft = data.type || 'text';
  if (_CAL_FIELD_TYPES.indexOf(ft) === -1) {
    throw new Error('不支援的欄位類型：' + ft);
  }

  // 驗證類型存在
  const types = _cal_readAll(_CAL_SHEET.TYPES);
  const target = types.find(t => t.typeId === data.typeId);
  if (!target) throw new Error('查無類型：' + data.typeId);

  const id = _cal_addFieldRow({
    typeId:    data.typeId, // 直接掛在傳入的類型上（頂層或子類型）
    name:      String(data.name).trim(),
    type:      ft,
    required:  !!data.required,
    options:   data.options,
    sortOrder: Number(data.sortOrder) || _cal_nextFieldSortOrder(data.typeId)
  });
  return {
    success: true, fieldId: id,
    typeId: data.typeId,
    isRoot: !target.parentTypeId,
    message: '已新增欄位'
  };
}

// 更新欄位
function cal_updateField(data) {
  if (!data || !data.fieldId) throw new Error('fieldId 必填');
  const sh = _cal_getSheet(_CAL_SHEET.FIELDS);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const colIdx = {};
  headers.forEach((h, i) => colIdx[h] = i);

  for (let r = 1; r < values.length; r++) {
    if (values[r][colIdx['fieldId']] === data.fieldId) {
      const setIf = (key, val, transform) => {
        if (val !== undefined && colIdx[key] !== undefined) {
          values[r][colIdx[key]] = transform ? transform(val) : val;
        }
      };
      setIf('顯示名稱', data.name && String(data.name).trim());
      if (data.type !== undefined) {
        if (_CAL_FIELD_TYPES.indexOf(data.type) === -1) throw new Error('不支援的欄位類型：' + data.type);
        setIf('欄位類型', data.type);
      }
      setIf('是否必填', data.required, v => v ? 'TRUE' : 'FALSE');
      if (data.options !== undefined) {
        setIf('下拉選項', data.options, v =>
          (typeof v === 'string' ? v : JSON.stringify(v))
        );
      }
      setIf('sortOrder', data.sortOrder, v => Number(v) || 0);
      sh.getRange(r + 1, 1, 1, values[r].length).setValues([values[r]]);
      return { success: true, message: '已更新' };
    }
  }
  throw new Error('找不到該欄位');
}

// 刪除欄位（同時清掉所有事項對該欄位的值）
function cal_deleteField(data) {
  if (!data || !data.fieldId) throw new Error('fieldId 必填');
  _cal_deleteRowsByIdSet(_CAL_SHEET.FIELDS, 'fieldId', new Set([data.fieldId]));
  _cal_deleteRowsByIdSet(_CAL_SHEET.VALUES, 'fieldId', new Set([data.fieldId]));
  return { success: true, message: '已刪除欄位' };
}

// 重新排序欄位（傳整批新順序）
function cal_reorderFields(data) {
  if (!data || !Array.isArray(data.fieldIds)) throw new Error('fieldIds 陣列必填');
  const sh = _cal_getSheet(_CAL_SHEET.FIELDS);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idIdx = headers.indexOf('fieldId');
  const orderIdx = headers.indexOf('sortOrder');

  data.fieldIds.forEach((fid, i) => {
    for (let r = 1; r < values.length; r++) {
      if (values[r][idIdx] === fid) {
        values[r][orderIdx] = i + 1;
        sh.getRange(r + 1, orderIdx + 1).setValue(i + 1);
        break;
      }
    }
  });
  return { success: true, message: '已重新排序' };
}

// ─────────────────────────────────────────────────────────────
//  helpers
// ─────────────────────────────────────────────────────────────
function _cal_findRootTypeId(typeId) {
  const all = _cal_readAll(_CAL_SHEET.TYPES);
  let cur = all.find(t => t.typeId === typeId);
  if (!cur) throw new Error('查無類型：' + typeId);
  while (cur.parentTypeId) {
    const p = all.find(t => t.typeId === cur.parentTypeId);
    if (!p) break;
    cur = p;
  }
  return cur.typeId;
}

function _cal_nextFieldSortOrder(rootTypeId) {
  const all = _cal_readAll(_CAL_SHEET.FIELDS);
  const siblings = all.filter(f => f.typeId === rootTypeId);
  const max = siblings.reduce((m, f) => Math.max(m, Number(f.sortOrder) || 0), 0);
  return max + 1;
}

function _cal_safeParseOptions(raw) {
  if (!raw) return [];
  if (Array.isArray(raw)) return raw;
  try { return JSON.parse(raw); } catch (e) {
    // 容錯：用逗號分隔的字串
    return String(raw).split(/[,，]/).map(s => s.trim()).filter(x => x);
  }
}
