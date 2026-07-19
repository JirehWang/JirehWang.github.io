(function() {
  const previousEditor = editor;
  const profile = window.activeWorshipTemplateProfile || {};
  const generatedIds = new Set(Array.isArray(profile.bibleSections)
    ? profile.bibleSections.map(item => item.sectionId)
    : ['call', 'scripture', 'verse']);
  const portIds = new Set(Array.isArray(profile.librarySections)
    ? profile.librarySections.map(item => item[0])
    : ['pre-hymn-1', 'pre-hymn-2', 'hymn-1', 'hymn-2', 'response', 'prayer-song', 'offering', 'doxology', 'amen']);
  const hymnOpacityIds = new Set(window.hymnOpacitySectionIds || []);

  editor = function() {
    if (!generatedIds.has(active) && !portIds.has(active)) return previousEditor();
    const item = model[active];
    const form = document.getElementById('editor-form');
    const sourceLabel = generatedIds.has(active) ? '行事曆輸入值（經文範圍）' : '行事曆輸入值（資料庫索引）';
    const note = generatedIds.has(active)
      ? `此值只作為經文查詢條件；投影片內容由${Array.isArray(profile.bibleVersions) && profile.bibleVersions.length > 1 ? '台語／華語' : '台語'}聖經資料產生器建立。`
      : '此值只作為資料庫索引；按下方按鈕後會從雲端下載並解析原始 PPTX。';
    form.innerHTML = `<div class="inline-note">${note}</div>${field(sourceLabel, 'sourceValue', item.sourceValue || '')}`;
    if (generatedIds.has(active)) {
      form.insertAdjacentHTML('beforeend', '<button type="button" class="button" id="regenerate-section">依輸入值重新產生</button>');
    } else {
      form.insertAdjacentHTML('beforeend', `<button type="button" class="button" id="load-library-section">載入雲端 PPT 資料庫</button>${item.libraryError ? `<p class="inline-note">${item.libraryError}</p>` : ''}`);
      if (active !== 'response') form.insertAdjacentHTML('beforeend', `<label class="field"><span>聖詩頁白色色塊透明度</span><div class="range-wrap"><input id="library-image-opacity" type="range" min="40" max="80" value="${item.opacity || 60}"><output class="range-value">${item.opacity || 60}%</output></div><small>數值越高，背景越淡。</small></label>`);
    }
    form.querySelector('[data-key="sourceValue"]').addEventListener('input', event => {
      item.sourceValue = event.target.value;
    });
    const regenerate = document.getElementById('regenerate-section');
    if (regenerate) regenerate.onclick = async () => {
      try {
        regenerate.disabled = true;
        status('正在依輸入值產生投影片…');
        await window.generateCalendarContent();
        render();
        status('已重新產生投影片內容');
      } catch (error) {
        status(`內容產生失敗：${error.message}`);
      } finally {
        regenerate.disabled = false;
      }
    };
    const loadLibrary = document.getElementById('load-library-section');
    if (loadLibrary) loadLibrary.onclick = async () => {
      try {
        loadLibrary.disabled = true;
        status('正在下載並解析雲端 PPTX…');
        const result = await window.reloadCurrentPptLibrarySection();
        status(result && result.state === 'missing' ? result.message : `已載入 ${result.pageCount || 0} 頁`);
      } catch (error) {
        status(`資料庫載入失敗：${error.message}`);
      } finally {
        loadLibrary.disabled = false;
      }
    };
    const imageOpacity = document.getElementById('library-image-opacity');
    if (imageOpacity) imageOpacity.oninput = event => {
      window.TaiwaneseWorshipSlideProduction.applyHymnOpacity(model, window.hymnOpacitySectionIds, active, Number(event.target.value), hymnOpacityIds.has(active) && window.isHymnOpacitySyncEnabled());
      imageOpacity.nextElementSibling.textContent = `${item.opacity}%`;
      preview();
    };
  };
  render();
})();
