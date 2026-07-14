(function() {
  const previousEditor = editor;
  const generatedIds = new Set(['call', 'scripture', 'verse']);
  const portIds = new Set(['pre-hymn-1', 'pre-hymn-2', 'hymn-1', 'hymn-2', 'response', 'prayer-song', 'offering', 'doxology', 'amen']);
  const hymnOpacityIds = new Set(window.hymnOpacitySectionIds || []);

  editor = function() {
    if (!generatedIds.has(active) && !portIds.has(active)) return previousEditor();
    const item = model[active];
    const form = document.getElementById('editor-form');
    const sourceLabel = generatedIds.has(active) ? '行事曆輸入值（經文範圍）' : '行事曆輸入值（資料庫索引）';
    const note = generatedIds.has(active)
      ? '此值只作為經文查詢條件；投影片內容由台語聖經資料產生器建立。'
      : '此值只作為資料庫索引；按下方按鈕後會從雲端下載並解析原始 PPTX。';
    form.innerHTML = `<div class="inline-note">${note}</div>${field(sourceLabel, 'sourceValue', item.sourceValue || '')}`;
    if (generatedIds.has(active)) {
      form.insertAdjacentHTML('beforeend', '<button type="button" class="button" id="regenerate-section">依輸入值重新產生</button>');
    } else {
      form.insertAdjacentHTML('beforeend', `<button type="button" class="button" id="load-library-section">載入雲端 PPT 資料庫</button>${item.libraryError ? `<p class="inline-note">${item.libraryError}</p>` : ''}`);
      if (active !== 'response') form.insertAdjacentHTML('beforeend', `<label class="field"><span>譜面透明度</span><div class="range-wrap"><input id="library-image-opacity" type="range" min="40" max="80" value="${item.opacity || 60}"><output class="range-value">${item.opacity || 60}%</output></div></label>${hymnOpacityIds.has(active) ? `<label class="sync-option"><input id="sync-hymn-opacity" type="checkbox" ${window.isHymnOpacitySyncEnabled() ? 'checked' : ''}><span>所有聖詩相關頁面一併套用</span></label>` : ''}`);
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
    const syncOpacity = document.getElementById('sync-hymn-opacity');
    if (syncOpacity) syncOpacity.onchange = event => {
      window.setHymnOpacitySyncEnabled(event.target.checked);
      const globalSync = document.getElementById('sync-hymn-opacity-global');
      if (globalSync) globalSync.checked = event.target.checked;
      if (event.target.checked) window.TaiwaneseWorshipSlideProduction.applyHymnOpacity(model, window.hymnOpacitySectionIds, active, item.opacity, true);
      preview();
      status(event.target.checked ? '已啟用所有聖詩譜面透明度同步' : '已改為各聖詩分開調整透明度');
    };
  };
  render();
})();
