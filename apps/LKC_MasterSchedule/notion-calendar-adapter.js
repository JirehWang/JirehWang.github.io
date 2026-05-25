/*
 * Notion calendar adapter for LKC_MasterSchedule.
 *
 * This file keeps the existing calendar UI intact, but redirects calendar
 * reads to a GAS proxy that talks to the official Notion API. The Notion token
 * must live in Apps Script Properties, never in GitHub Pages.
 */
(function () {
  const cfg = window.LKC_NOTION_CALENDAR || {};

  if (!cfg.enabled) return;

  const originalChurchAPI = window.churchAPI;
  const readActions = new Set(['cal_getTypes', 'cal_getFields', 'cal_getEvents', 'cal_getEvent']);
  const writeActions = new Set([
    'cal_addEvent',
    'cal_updateEvent',
    'cal_deleteEvent',
    'cal_addEventsBatch',
    'cal_aiParseForType'
  ]);

  if (typeof originalChurchAPI !== 'function') {
    console.warn('[NotionCalendar] churchAPI is not ready; adapter skipped.');
    return;
  }

  window.churchAPI = async function notionCalendarChurchAPI(action, data) {
    if (readActions.has(action)) {
      const res = await originalChurchAPI('notion_' + action, data || {});
      return normalizeNotionCalendarResponse(action, res);
    }

    if (writeActions.has(action)) {
      if (action === 'cal_aiParseForType') {
        return {
          success: false,
          message: 'AI import to Notion is not enabled yet. Please use manual input first.'
        };
      }
      const res = await originalChurchAPI('notion_' + action, data || {});
      return normalizeNotionCalendarResponse(action, res);
    }

    return originalChurchAPI(action, data || {});
  };

  window.LKC_NOTION_CALENDAR_READY = true;
  console.info('[NotionCalendar] Read adapter enabled.');

  function normalizeNotionCalendarResponse(action, res) {
    if (!res || !res.success) return res;

    if (action === 'cal_getTypes' && res.data) {
      const flat = Array.isArray(res.data.flat) ? res.data.flat.map(normalizeType) : [];
      const tree = Array.isArray(res.data.tree || res.data.types)
        ? (res.data.tree || res.data.types).map(normalizeType)
        : flat;
      res.data.flat = flat;
      res.data.tree = tree;
      res.data.types = tree;
      return res;
    }

    if (action === 'cal_getFields' && res.data) {
      ['fields', 'ownFields', 'inheritedFields'].forEach(key => {
        if (Array.isArray(res.data[key])) res.data[key] = res.data[key].map(normalizeField);
      });
      return res;
    }

    if (action === 'cal_getEvents' && Array.isArray(res.data)) {
      res.data = res.data.map(normalizeEvent);
      return res;
    }

    if ((action === 'cal_getEvent' || action === 'cal_addEvent' || action === 'cal_updateEvent') && res.data) {
      res.data = normalizeEvent(res.data);
      return res;
    }

    return res;
  }

  function normalizeType(type) {
    const name = type.name || type.label || type.typeName || type['?迂'] || type.typeId || '';
    return {
      ...type,
      name,
      typeName: name,
      label: name,
      '名稱': name,
      '?迂': name,
      children: Array.isArray(type.children) ? type.children.map(normalizeType) : []
    };
  }

  function normalizeField(field) {
    const name = field.name || field.fieldName || field['憿舐內?迂'] || field.fieldId || '';
    const type = field.type || field.fieldType || field['甈?憿?'] || 'text';
    const options = field.options || field['銝??賊?'] || [];
    return {
      ...field,
      name,
      type,
      options,
      '顯示名稱': name,
      '欄位類型': type,
      '選項': options,
      '憿舐內?迂': name,
      '甈?憿?': type,
      '銝??賊?': options
    };
  }

  function normalizeEvent(event) {
    return {
      ...event,
      typeName: event.typeName || event.name || event.typeId || '',
      typeFullName: event.typeFullName || event.typeName || event.typeId || '',
      values: Array.isArray(event.values) ? event.values : []
    };
  }
})();
