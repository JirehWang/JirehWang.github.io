(function() {
  const button = document.getElementById('calendar-load');
  button.onclick = async function() {
    const date = document.getElementById('service-date').value;
    if (!date) return status('請先選擇禮拜日期');
    try {
      button.disabled = true;
      status('正在讀取行事曆…');
      const result = await window.worshipReadAPI('cal_getEvents', { startDate: date, endDate: date });
      const events = Array.isArray(result && result.data) ? result.data : [];
      const event = window.TaiwaneseWorshipCalendarAdapter.selectTaiwaneseSermonEvent(events, date);
      if (!event) return status(`找不到 ${date} 的「講道資訊－台語」資料`);
      window.TaiwaneseWorshipCalendarAdapter.applyCalendarEvent(event, model);
      await window.generateCalendarContent();
      const libraryResults = await window.loadPptLibraryContent();
      render();
      const loadedPages = libraryResults.reduce((total, item) => total + (item.pageCount || 0), 0);
      const missing = libraryResults.filter(item => item.state === 'missing');
      status(`已帶入「講道資訊－台語」：${event.title || date}；資料庫 ${loadedPages} 頁${missing.length ? `，${missing.length} 項找不到` : ''}`);
    } catch (error) {
      status(`行事曆帶入失敗：${error.message}`);
    } finally {
      button.disabled = false;
    }
  };
})();
