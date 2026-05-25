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
      return originalChurchAPI('notion_' + action, data || {});
    }

    if (writeActions.has(action)) {
      if (action === 'cal_aiParseForType') {
        return {
          success: false,
          message: 'AI import to Notion is not enabled yet. Please use manual input first.'
        };
      }
      return originalChurchAPI('notion_' + action, data || {});
    }

    return originalChurchAPI(action, data || {});
  };

  window.LKC_NOTION_CALENDAR_READY = true;
  console.info('[NotionCalendar] Read adapter enabled.');
})();
