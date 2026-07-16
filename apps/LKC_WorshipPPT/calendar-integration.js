(function() {
  const button = document.getElementById('calendar-load');
  button.onclick = async function() {
    const date = document.getElementById('service-date').value;
    if (!date) return status('請先選擇禮拜日期');
    try {
      button.disabled = true;
      status('正在讀取行事曆與週報資料…');
      const bulletinPromise = window.loadBulletinPptContent
        ? window.loadBulletinPptContent(date).catch(error => ({ error: error.message }))
        : Promise.resolve({ error: '週報讀取模組未載入' });
      const result = await window.worshipReadAPI('cal_getEvents', { startDate: date, endDate: date });
      const events = Array.isArray(result && result.data) ? result.data : [];
      const event = window.TaiwaneseWorshipCalendarAdapter.selectTaiwaneseSermonEvent(events, date);
      let calendarSummary = `找不到 ${date} 的「講道資訊－台語」資料`;
      let libraryResults = [];
      if (event) {
        window.TaiwaneseWorshipCalendarAdapter.applyCalendarEvent(event, model);
        await window.generateCalendarContent();
        libraryResults = await window.loadPptLibraryContent();
        calendarSummary = `已帶入「講道資訊－台語」：${event.title || date}`;
      }
      const bulletinResult = await bulletinPromise;
      render();
      const loadedPages = libraryResults.reduce((total, item) => total + (item.pageCount || 0), 0);
      const missing = libraryResults.filter(item => item.state === 'missing');
      const librarySummary = event ? `資料庫 ${loadedPages} 頁${missing.length ? `，${missing.length} 項找不到` : ''}` : '未載入聖詩／啟應文';
      const bulletinSummary = window.describeBulletinPptContent(bulletinResult);
      status(`${calendarSummary}；${librarySummary}；${bulletinSummary}`);
    } catch (error) {
      status(`行事曆帶入失敗：${error.message}`);
    } finally {
      button.disabled = false;
    }
  };
})();
