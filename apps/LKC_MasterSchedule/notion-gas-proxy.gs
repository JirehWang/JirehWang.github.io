/**
 * Notion proxy for LKC_MasterSchedule.
 *
 * Required Script Properties:
 * - NOTION_TOKEN
 * - NOTION_CALENDAR_DATABASE_ID
 */

const NOTION_VERSION = '2022-06-28';
const NOTION_ROOT_TYPES = ['聚會名稱', '講道資訊', '會議'];

const NOTION_PROP = {
  title: '活動名稱',
  date: '日期',
  type: '活動類型',
  subtype: '子類型',
  displayTitle: '顯示標題',
  location: '地點',
  ministry: '負責單位',
  owner: '負責人',
  public: '是否公開',
  status: '狀態',
  note: '備註'
};

function handleNotionCalendarAction(action, data) {
  switch (action) {
    case 'notion_cal_getTypes':
      return notion_cal_getTypes(data);
    case 'notion_cal_getFields':
      return notion_cal_getFields(data);
    case 'notion_cal_getEvents':
      return notion_cal_getEvents(data || {});
    case 'notion_cal_getEvent':
      return notion_cal_getEvent(data || {});
    case 'notion_cal_addEvent':
      return notion_cal_addEvent(data || {});
    case 'notion_cal_updateEvent':
      return notion_cal_updateEvent(data || {});
    case 'notion_cal_deleteEvent':
      return notion_cal_deleteEvent(data || {});
    case 'notion_cal_addEventsBatch':
      return notion_cal_addEventsBatch(data || {});
    case 'notion_cal_aiParseForType':
      return notion_cal_aiParseForType(data || {});
    case 'notion_cal_addType':
      return notion_cal_addType(data || {});
    case 'notion_cal_updateType':
      return notion_cal_updateType(data || {});
    case 'notion_cal_deleteType':
      return notion_cal_deleteType(data || {});
    case 'notion_cal_addField':
    case 'notion_cal_updateField':
    case 'notion_cal_deleteField':
    case 'notion_cal_reorderFields':
      return notion_cal_fieldsReadOnly_(action);
    default:
      return { success: false, message: 'Unknown Notion calendar action: ' + action };
  }
}

function notion_cal_getTypes() {
  const db = notionGetDatabase_();
  const typeOptions = (((db.properties || {})[NOTION_PROP.type] || {}).select || {}).options || [];
  const subTypeOptions = (((db.properties || {})[NOTION_PROP.subtype] || {}).select || {}).options || [];
  const subTypesByRoot = notionSubTypesByRoot_(subTypeOptions);
  const rootTypes = notionRootTypes_();
  const roots = typeOptions
    .filter(opt => rootTypes.indexOf(opt.name) !== -1)
    .map(opt => {
      const root = notionTypeFromOption_(opt, '');
      root.children = (subTypesByRoot[opt.name] || []).map(child => notionTypeFromOption_(child, opt.name));
      return root;
    });
  const flat = [];
  roots.forEach(root => {
    flat.push(root);
    (root.children || []).forEach(child => flat.push(child));
  });

  return {
    success: true,
    data: {
      types: roots,
      tree: roots,
      flat: flat
    }
  };
}

function notion_cal_getFields() {
  const db = notionGetDatabase_();
  const props = db.properties || {};
  const names = [
    NOTION_PROP.subtype,
    NOTION_PROP.displayTitle,
    NOTION_PROP.location,
    NOTION_PROP.ministry,
    NOTION_PROP.owner,
    NOTION_PROP.public,
    NOTION_PROP.status,
    NOTION_PROP.note
  ];

  const fields = names
    .filter(name => props[name])
    .map((name, idx) => notionFieldFromProperty_(name, props[name], idx + 1));

  return {
    success: true,
    data: {
      rootTypeId: '',
      subTypeId: '',
      inheritedFields: [],
      ownFields: fields,
      fields: fields,
      excludedFieldIds: []
    }
  };
}

function notion_cal_getEvents(data) {
  const payload = {
    page_size: 100,
    sorts: [{ property: NOTION_PROP.date, direction: 'ascending' }]
  };

  const filters = [];
  if (data && data.startDate) {
    filters.push({ property: NOTION_PROP.date, date: { on_or_after: data.startDate } });
  }
  if (data && data.endDate) {
    filters.push({ property: NOTION_PROP.date, date: { before: data.endDate } });
  }
  if (data && data.typeIds && data.typeIds.length) {
    const typeIds = data.typeIds.map(String);
    const rootTypes = typeIds.filter(id => NOTION_ROOT_TYPES.indexOf(id) !== -1);
    const subTypes = typeIds.filter(id => NOTION_ROOT_TYPES.indexOf(id) === -1);
    const orFilters = [];
    rootTypes.forEach(typeId => {
      orFilters.push({
        property: NOTION_PROP.type,
        select: { equals: typeId }
      });
    });
    subTypes.forEach(typeId => {
      const normalized = notionNormalizeTypeInput_(typeId);
      orFilters.push({
        property: NOTION_PROP.subtype,
        select: { equals: normalized.subType || typeId }
      });
    });
    filters.push({
      or: orFilters
    });
  }
  if (filters.length === 1) payload.filter = filters[0];
  if (filters.length > 1) payload.filter = { and: filters };

  const pages = [];
  let cursor = null;
  do {
    if (cursor) payload.start_cursor = cursor;
    const res = notionFetch_('/databases/' + notionDatabaseId_() + '/query', 'post', payload);
    (res.results || []).forEach(page => pages.push(page));
    cursor = res.has_more ? res.next_cursor : null;
  } while (cursor);

  return {
    success: true,
    data: pages.map(notionPageToCalendarEvent_)
  };
}

function notion_cal_getEvent(data) {
  if (!data || !data.eventId) {
    return { success: false, message: 'Missing eventId' };
  }
  const page = notionFetch_('/pages/' + data.eventId, 'get');
  return {
    success: true,
    data: notionPageToCalendarEvent_(page)
  };
}

function notion_cal_addEvent(data) {
  if (!data.typeId) return { success: false, message: 'Missing typeId' };
  if (!data.date) return { success: false, message: 'Missing date' };

  const payload = {
    parent: { database_id: notionDatabaseId_() },
    properties: notionBuildPageProperties_(data)
  };
  const page = notionFetch_('/pages', 'post', payload);
  return {
    success: true,
    eventId: page.id,
    message: '已新增到 Notion',
    data: notionPageToCalendarEvent_(page)
  };
}

function notion_cal_updateEvent(data) {
  if (!data.eventId) return { success: false, message: 'Missing eventId' };

  const payload = {
    properties: notionBuildPageProperties_(data)
  };
  const page = notionFetch_('/pages/' + data.eventId, 'patch', payload);
  return {
    success: true,
    message: '已更新到 Notion',
    data: notionPageToCalendarEvent_(page)
  };
}

function notion_cal_deleteEvent(data) {
  if (!data.eventId) return { success: false, message: 'Missing eventId' };

  notionFetch_('/pages/' + data.eventId, 'patch', { archived: true });
  return {
    success: true,
    message: '已從 Notion 封存'
  };
}

function notion_cal_addEventsBatch(data) {
  const events = Array.isArray(data.events) ? data.events : [];
  if (events.length === 0) return { success: false, message: 'No events to add' };

  const created = [];
  events.forEach(ev => {
    const res = notion_cal_addEvent(ev);
    if (res.success) created.push(res.eventId);
  });

  return {
    success: true,
    message: '已新增 ' + created.length + ' 筆到 Notion',
    eventIds: created
  };
}

function notion_cal_aiParseForType(data) {
  if (!data || !data.rootTypeId) return { success: false, message: 'Missing rootTypeId' };
  if (!data.rawText || !String(data.rawText).trim()) return { success: false, message: 'Missing rawText' };

  const typesRes = notion_cal_getTypes();
  const flatTypes = (typesRes.data && typesRes.data.flat) || [];
  const rootType = flatTypes.find(t => t.typeId === data.rootTypeId) ||
    flatTypes.find(t => !t.parentTypeId && t.name === data.rootTypeId);
  if (!rootType) return { success: false, message: '找不到指定的 Notion 活動類型' };

  const subTypes = flatTypes.filter(t => t.parentTypeId === rootType.typeId);
  const fieldsRes = notion_cal_getFields({ typeId: rootType.typeId });
  const fields = ((fieldsRes.data && fieldsRes.data.fields) || [])
    .filter(f => f.fieldId !== NOTION_PROP.subtype)
    .filter(f => f.fieldId !== NOTION_PROP.public)
    .filter(f => f.fieldId !== NOTION_PROP.status)
    .filter(f => f.fieldId !== NOTION_PROP.note);

  const fieldsDesc = fields.map(f => {
    const options = Array.isArray(f.options) && f.options.length ? ' options=[' + f.options.join(', ') + ']' : '';
    return '  - "' + f.fieldId + '": ' + f.name + ' (' + f.type + ')' + options;
  }).join('\n') || '  - "' + NOTION_PROP.displayTitle + '": 顯示標題 (text)';

  const subTypesDesc = subTypes.length
    ? '\n可用子類型，請只使用其中一個作為 subTypeName：\n' +
      subTypes.map(s => '  - "' + s.name + '"').join('\n')
    : '';

  const allowMultiple = data.allowMultiple !== false;
  const sampleSubType = subTypes[0] ? subTypes[0].name : '';
  const prompt = [
    '你是教會行事曆資料整理助手。請從使用者貼上的文字中，整理出可匯入 Notion 行事曆的 JSON。',
    '活動類型是：' + rootType.name,
    '可用欄位如下，values 的 key 必須使用欄位 id：',
    fieldsDesc,
    subTypesDesc,
    '',
    '規則：',
    '1. 日期一律輸出 YYYY-MM-DD。',
    '2. ' + (allowMultiple ? '如果文字中有多筆行程，請輸出多個 events。' : '只輸出一個 event。'),
    '3. 找不到的欄位請留空字串，不要捏造。',
    '4. title 是月曆上顯示的短標題；若不確定可留空。',
    '5. 只輸出 JSON，不要 markdown，不要解釋。',
    '',
    '格式：',
    '{',
    '  "events": [',
    '    {',
    '      "date": "2026-01-05",',
    '      "subTypeName": "' + sampleSubType + '",',
    '      "title": "",',
    '      "values": {',
    fields.slice(0, 2).map(f => '        "' + f.fieldId + '": ""').join(',\n'),
    '      }',
    '    }',
    '  ]',
    '}'
  ].join('\n');

  const aiResult = callGeminiApi(prompt, data.rawText);
  let parsed;
  try {
    parsed = JSON.parse(aiResult);
  } catch (err) {
    const match = String(aiResult || '').match(/\{[\s\S]*\}/);
    if (match) {
      try {
        parsed = JSON.parse(match[0]);
      } catch (innerErr) {}
    }
    if (!parsed) throw new Error('AI 回傳無法解析為 JSON：' + String(aiResult || '').substring(0, 200));
  }

  const subTypesByName = {};
  subTypes.forEach(s => subTypesByName[s.name] = s.typeId);
  const events = (Array.isArray(parsed.events) ? parsed.events : []).map(ev => {
    const subTypeId = ev.subTypeName ? subTypesByName[ev.subTypeName] : '';
    return {
      date: ev.date || '',
      subTypeName: ev.subTypeName || '',
      subTypeId: subTypeId || (subTypes.length === 0 ? rootType.typeId : ''),
      title: ev.title || '',
      values: ev.values || {}
    };
  });

  return {
    success: true,
    rootTypeId: rootType.typeId,
    rootTypeName: rootType.name,
    events: events,
    hasSubTypes: subTypes.length > 0
  };
}

function notion_cal_addType(data) {
  const name = String(data.name || '').trim();
  if (!name) return { success: false, message: '類型名稱不可空白' };

  const parentTypeId = String(data.parentTypeId || '').trim();
  if (!parentTypeId) {
    notionUpsertSelectOption_(NOTION_PROP.type, name, data.color);
    const groups = notionGetSubTypeGroups_();
    if (!groups[name]) groups[name] = [];
    notionSaveSubTypeGroups_(groups);
    return {
      success: true,
      typeId: name,
      message: '已新增到 Notion 活動類型'
    };
  }

  const parent = notionNormalizeTypeInput_(parentTypeId).rootType || parentTypeId;
  notionUpsertSelectOption_(NOTION_PROP.subtype, name, data.color);
  const groups = notionGetSubTypeGroups_();
  groups[parent] = groups[parent] || [];
  if (groups[parent].indexOf(name) === -1) groups[parent].push(name);
  notionSaveSubTypeGroups_(groups);

  return {
    success: true,
    typeId: parent + '::' + name,
    message: '已新增到 Notion 子類型'
  };
}

function notion_cal_updateType(data) {
  if (!data || !data.typeId) return { success: false, message: 'Missing typeId' };
  if (data.name === undefined && data.color === undefined && data.excludedFieldIds !== undefined) {
    return { success: true, message: 'Notion 欄位繼承設定已略過' };
  }

  const parsed = notionNormalizeTypeInput_(data.typeId);
  const oldRoot = parsed.rootType || data.typeId;
  const oldSub = parsed.subType || '';
  const nextName = String(data.name || oldSub || oldRoot).trim();
  if (!nextName) return { success: false, message: '類型名稱不可空白' };

  const groups = notionGetSubTypeGroups_();
  if (oldSub) {
    notionUpsertSelectOption_(NOTION_PROP.subtype, nextName, data.color);
    groups[oldRoot] = groups[oldRoot] || [];
    const idx = groups[oldRoot].indexOf(oldSub);
    if (idx === -1 && groups[oldRoot].indexOf(nextName) === -1) groups[oldRoot].push(nextName);
    if (idx !== -1) groups[oldRoot][idx] = nextName;
    notionSaveSubTypeGroups_(groups);
    return {
      success: true,
      typeId: oldRoot + '::' + nextName,
      message: '已更新 Notion 子類型'
    };
  }

  notionUpsertSelectOption_(NOTION_PROP.type, nextName, data.color);
  if (oldRoot !== nextName) {
    groups[nextName] = groups[oldRoot] || [];
    delete groups[oldRoot];
    notionSaveSubTypeGroups_(groups);
  }
  return {
    success: true,
    typeId: nextName,
    message: '已更新 Notion 活動類型；舊 select 選項會保留在 Notion 中'
  };
}

function notion_cal_deleteType(data) {
  if (!data || !data.typeId) return { success: false, message: 'Missing typeId' };

  const parsed = notionNormalizeTypeInput_(data.typeId);
  const root = parsed.rootType || data.typeId;
  const sub = parsed.subType || '';
  const groups = notionGetSubTypeGroups_();

  if (sub) {
    groups[root] = (groups[root] || []).filter(name => name !== sub);
    notionSaveSubTypeGroups_(groups);
    return {
      success: true,
      message: '已從 GitHub 事項管理隱藏；Notion 的 select 選項仍保留，避免影響既有資料'
    };
  }

  delete groups[root];
  notionSaveSubTypeGroups_(groups);
  return {
    success: true,
    message: '已從 GitHub 事項管理隱藏；Notion 的活動類型選項仍保留，避免影響既有資料'
  };
}

function notion_cal_fieldsReadOnly_(action) {
  return {
    success: false,
    message: '目前 Notion 欄位請直接在 Notion 資料庫調整；前端先支援類型與子類型管理。'
  };
}

function notionBuildPageProperties_(data) {
  const values = data.values || {};
  const normalized = notionNormalizeTypeInput_(data.typeId);
  const rootType = normalized.rootType || data.typeId;
  const subType = values[NOTION_PROP.subtype] || normalized.subType || '';
  const title = data.title || values[NOTION_PROP.displayTitle] || subType || rootType || '未命名活動';

  const props = {};
  props[NOTION_PROP.title] = { title: [{ text: { content: String(title) } }] };

  if (data.date !== undefined) {
    props[NOTION_PROP.date] = { date: data.date ? { start: String(data.date).substring(0, 10) } : null };
  }

  if (data.typeId !== undefined) {
    props[NOTION_PROP.type] = rootType ? { select: { name: String(rootType) } } : { select: null };
  }

  notionSetProp_(props, NOTION_PROP.subtype, subType, 'select');
  notionSetProp_(props, NOTION_PROP.displayTitle, values[NOTION_PROP.displayTitle], 'text');
  notionSetProp_(props, NOTION_PROP.location, values[NOTION_PROP.location], 'text');
  notionSetProp_(props, NOTION_PROP.ministry, values[NOTION_PROP.ministry], 'text');
  notionSetProp_(props, NOTION_PROP.owner, values[NOTION_PROP.owner], 'text');
  notionSetProp_(props, NOTION_PROP.public, values[NOTION_PROP.public], 'checkbox');
  notionSetProp_(props, NOTION_PROP.status, values[NOTION_PROP.status], 'select');
  notionSetProp_(props, NOTION_PROP.note, values[NOTION_PROP.note], 'text');

  return props;
}

function notionSetProp_(props, name, value, type) {
  if (value === undefined) return;
  if (type === 'select') {
    props[name] = value ? { select: { name: String(value) } } : { select: null };
    return;
  }
  if (type === 'checkbox') {
    props[name] = { checkbox: notionTruthy_(value) };
    return;
  }
  props[name] = { rich_text: value ? [{ text: { content: String(value) } }] : [] };
}

function notionTruthy_(value) {
  const text = String(value || '').trim().toLowerCase();
  return value === true || text === 'true' || text === '1' || text === 'yes' || text === '是' || text === '公開';
}

function notionPageToCalendarEvent_(page) {
  const props = page.properties || {};
  const title = notionPlainText_(props[NOTION_PROP.displayTitle]) ||
    notionPlainText_(props[NOTION_PROP.title]) ||
    notionSelectName_(props[NOTION_PROP.subtype]) ||
    notionSelectName_(props[NOTION_PROP.type]) ||
    '未命名活動';
  const typeName = notionSelectName_(props[NOTION_PROP.type]) || '其他';
  const subType = notionSelectName_(props[NOTION_PROP.subtype]);
  const date = notionDateStart_(props[NOTION_PROP.date]);
  const color = notionTypeColor_(typeName);

  return {
    eventId: page.id,
    typeId: typeName,
    typeName: subType || typeName,
    typeFullName: subType ? typeName + ' / ' + subType : typeName,
    typeIcon: notionTypeIcon_(typeName),
    typeColor: color,
    title: title,
    date: date,
    values: [
      notionValueRow_('子類型', subType, 'select'),
      notionValueRow_('地點', notionPlainText_(props[NOTION_PROP.location]), 'text'),
      notionValueRow_('負責單位', notionPlainText_(props[NOTION_PROP.ministry]), 'text'),
      notionValueRow_('負責人', notionPlainText_(props[NOTION_PROP.owner]), 'text'),
      notionValueRow_('是否公開', notionCheckbox_(props[NOTION_PROP.public]) ? '是' : '否', 'text'),
      notionValueRow_('狀態', notionSelectName_(props[NOTION_PROP.status]), 'select'),
      notionValueRow_('備註', notionPlainText_(props[NOTION_PROP.note]), 'longtext')
    ].filter(row => row.value)
  };
}

function notionTypeFromOption_(opt, parentTypeId) {
  const id = parentTypeId ? parentTypeId + '::' + opt.name : opt.name;
  const type = {
    typeId: id,
    parentTypeId: parentTypeId || '',
    icon: notionTypeIcon_(opt.name),
    color: notionTypeColor_(opt.name),
    children: [],
    sortOrder: 0,
    hidden: false
  };
  type['?迂'] = opt.name;
  type.name = opt.name;
  return type;
}

function notionSubTypesByRoot_(subTypeOptions) {
  const groups = notionGetSubTypeGroups_();
  const byName = {};
  subTypeOptions.forEach(opt => byName[opt.name] = opt);
  const result = {};
  Object.keys(groups).forEach(root => {
    result[root] = groups[root]
      .filter(name => byName[name])
      .map(name => byName[name]);
  });
  return result;
}

function notionNormalizeTypeInput_(typeId) {
  const text = String(typeId || '');
  if (text.indexOf('::') !== -1) {
    const parts = text.split('::');
    return { rootType: parts[0], subType: parts.slice(1).join('::') };
  }
  if (notionRootTypes_().indexOf(text) !== -1) return { rootType: text, subType: '' };
  const groups = notionGetSubTypeGroups_();
  for (const root in groups) {
    if ((groups[root] || []).indexOf(text) !== -1) return { rootType: root, subType: text };
  }
  return { rootType: '聚會名稱', subType: text };
}

function notionFieldFromProperty_(name, prop, sortOrder) {
  const fieldType = notionFieldType_(prop.type);
  let options = prop.type === 'select'
    ? ((prop.select || {}).options || []).map(opt => opt.name)
    : [];
  if (prop.type === 'checkbox') options = ['是', '否'];
  const field = {
    fieldId: name,
    typeId: '',
    required: false,
    sortOrder: sortOrder
  };
  field['憿舐內?迂'] = name;
  field['甈?憿?'] = fieldType;
  field['銝??賊?'] = options;
  field.name = name;
  field.type = fieldType;
  field.options = options;
  return field;
}

function notionFieldType_(type) {
  if (type === 'rich_text' || type === 'title') return 'text';
  if (type === 'date') return 'date';
  if (type === 'number') return 'number';
  if (type === 'url') return 'url';
  if (type === 'select' || type === 'status') return 'select';
  if (type === 'multi_select') return 'multiselect';
  if (type === 'checkbox') return 'select';
  return 'text';
}

function notionValueRow_(name, value, fieldType) {
  return {
    fieldId: name,
    fieldName: name,
    fieldType: fieldType || 'text',
    value: value || ''
  };
}

function notionRootTypes_() {
  const groups = notionGetSubTypeGroups_();
  const roots = NOTION_ROOT_TYPES.slice();
  Object.keys(groups).forEach(name => {
    if (roots.indexOf(name) === -1) roots.push(name);
  });
  return roots;
}

function notionGetSubTypeGroups_() {
  const defaults = {
    '聚會名稱': ['聚會名稱'],
    '講道資訊': ['台語', '華語', '聯合'],
    '會議': ['長執會', '小會', '幼兒園董事會', '同工會議', '牧區會議']
  };
  const raw = PropertiesService.getScriptProperties().getProperty('NOTION_SUBTYPE_GROUPS');
  if (!raw) return defaults;
  try {
    const parsed = JSON.parse(raw);
    Object.keys(defaults).forEach(root => {
      if (!Array.isArray(parsed[root])) parsed[root] = defaults[root];
    });
    return parsed;
  } catch (err) {
    return defaults;
  }
}

function notionSaveSubTypeGroups_(groups) {
  PropertiesService.getScriptProperties().setProperty('NOTION_SUBTYPE_GROUPS', JSON.stringify(groups || {}));
}

function notionUpsertSelectOption_(propertyName, optionName, color) {
  const db = notionGetDatabase_();
  const prop = (db.properties || {})[propertyName];
  if (!prop || prop.type !== 'select') {
    throw new Error('Notion property is not select: ' + propertyName);
  }

  const options = ((prop.select || {}).options || []).map(opt => ({
    id: opt.id,
    name: opt.name,
    color: opt.color || 'default'
  }));
  const existing = options.find(opt => opt.name === optionName);
  if (!existing) {
    options.push({
      name: optionName,
      color: notionOptionColor_(color)
    });
  }
  notionUpdateSelectOptions_(propertyName, options);
}

function notionUpdateSelectOptions_(propertyName, options) {
  const clean = options.map(opt => {
    const item = {
      name: opt.name,
      color: opt.color || 'default'
    };
    if (opt.id) item.id = opt.id;
    return item;
  });
  const payload = { properties: {} };
  payload.properties[propertyName] = { select: { options: clean } };
  notionFetch_('/databases/' + notionDatabaseId_(), 'patch', payload);
}

function notionOptionColor_(value) {
  const text = String(value || '').trim().toLowerCase();
  if (['default', 'gray', 'brown', 'orange', 'yellow', 'green', 'blue', 'purple', 'pink', 'red'].indexOf(text) !== -1) {
    return text;
  }
  if (/^#/.test(text)) {
    const hex = text.replace('#', '');
    const r = parseInt(hex.substring(0, 2), 16);
    const g = parseInt(hex.substring(2, 4), 16);
    const b = parseInt(hex.substring(4, 6), 16);
    if (Number.isNaN(r) || Number.isNaN(g) || Number.isNaN(b)) return 'default';
    if (r > 200 && g < 120 && b < 140) return 'red';
    if (r > 210 && g > 120 && b < 100) return 'orange';
    if (r > 190 && g > 170 && b < 110) return 'yellow';
    if (g > r && g > b) return 'green';
    if (b > r && b > g) return 'blue';
    if (r > 150 && b > 150) return 'purple';
    if (r > 180 && b > 130) return 'pink';
    return 'gray';
  }
  return 'default';
}

function notionGetDatabase_() {
  return notionFetch_('/databases/' + notionDatabaseId_(), 'get');
}

function notionFetch_(path, method, body) {
  const params = {
    method: method || 'get',
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + notionToken_(),
      'Notion-Version': NOTION_VERSION,
      'Content-Type': 'application/json'
    }
  };
  if (body) params.payload = JSON.stringify(body);

  const resp = UrlFetchApp.fetch('https://api.notion.com/v1' + path, params);
  const text = resp.getContentText();
  const status = resp.getResponseCode();
  const json = text ? JSON.parse(text) : {};
  if (status < 200 || status >= 300) {
    throw new Error('Notion API error ' + status + ': ' + (json.message || text));
  }
  return json;
}

function notionToken_() {
  const token = PropertiesService.getScriptProperties().getProperty('NOTION_TOKEN');
  if (!token) throw new Error('Missing Script Property: NOTION_TOKEN');
  return token;
}

function notionDatabaseId_() {
  return PropertiesService.getScriptProperties().getProperty('NOTION_CALENDAR_DATABASE_ID') ||
    '3871fade-5670-41d2-991f-af1612d48c0b';
}

function notionPlainText_(prop) {
  if (!prop) return '';
  const arr = prop.title || prop.rich_text;
  if (Array.isArray(arr)) return arr.map(x => x.plain_text || '').join('');
  if (prop.type === 'select') return notionSelectName_(prop);
  if (prop.type === 'checkbox') return prop.checkbox ? 'true' : '';
  if (prop.type === 'date') return notionDateStart_(prop);
  return '';
}

function notionSelectName_(prop) {
  if (!prop) return '';
  if (prop.select) return prop.select.name || '';
  if (prop.status) return prop.status.name || '';
  return '';
}

function notionDateStart_(prop) {
  return prop && prop.date && prop.date.start ? prop.date.start.substring(0, 10) : '';
}

function notionCheckbox_(prop) {
  return !!(prop && prop.checkbox);
}

function notionTypeColor_(typeName) {
  const map = {
    '聚會名稱': '#2563eb',
    '講道資訊': '#db2777',
    '會議': '#4b5563'
  };
  return map[typeName] || '#667eea';
}

function notionTypeIcon_(typeName) {
  const map = {
    '聚會名稱': '📅',
    '講道資訊': '📖',
    '會議': '📝'
  };
  return map[typeName] || '';
}
