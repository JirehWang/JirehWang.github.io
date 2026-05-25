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
    default:
      return { success: false, message: 'Unknown Notion calendar action: ' + action };
  }
}

function notion_cal_getTypes() {
  const db = notionGetDatabase_();
  const typeOptions = (((db.properties || {})[NOTION_PROP.type] || {}).select || {}).options || [];
  const roots = typeOptions
    .filter(opt => NOTION_ROOT_TYPES.indexOf(opt.name) !== -1)
    .map(opt => notionTypeFromOption_(opt, ''));

  return {
    success: true,
    data: {
      types: roots,
      tree: roots,
      flat: roots
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
    filters.push({
      or: data.typeIds.map(typeId => ({
        property: NOTION_PROP.type,
        select: { equals: typeId }
      }))
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

function notionBuildPageProperties_(data) {
  const values = data.values || {};
  const title = data.title || values[NOTION_PROP.displayTitle] || data.typeId || '未命名活動';

  const props = {};
  props[NOTION_PROP.title] = { title: [{ text: { content: String(title) } }] };

  if (data.date !== undefined) {
    props[NOTION_PROP.date] = { date: data.date ? { start: String(data.date).substring(0, 10) } : null };
  }

  if (data.typeId !== undefined) {
    props[NOTION_PROP.type] = data.typeId ? { select: { name: String(data.typeId) } } : { select: null };
  }

  notionSetProp_(props, NOTION_PROP.subtype, values[NOTION_PROP.subtype], 'select');
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
  const type = {
    typeId: opt.name,
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
