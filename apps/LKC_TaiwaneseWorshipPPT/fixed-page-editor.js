(function() {
  const baseEditor = editor;
  const escapeText = value => String(value || '').replace(/[&<>]/g, char => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;' })[char]);
  const fixedIds = new Set(['creed', 'lord-prayer']);
  editor = function() {
    const item = model[active];
    if (!Array.isArray(item.pptPages) || !fixedIds.has(active)) return baseEditor();
    const form = document.getElementById('editor-form');
    const templates = item.pptPages.map(page => typeof page === 'string' ? ({ body: page }) : ({ ...page }));
    const weights = templates.map(page => String(page.body || '').length || 1);
    form.innerHTML = `<div class="inline-note">使用單一全文欄位編輯；系統會依原 PPT 各頁文字量比例自動重排，右側立即顯示結果。</div><label class="field"><span>${escapeText(item.label)}全文</span><textarea data-fixed-source>${escapeText(item.body || '')}</textarea><small>空白行會作為優先分頁位置。</small></label>`;
    form.querySelector('[data-fixed-source]').addEventListener('input', event => {
      item.body = event.target.value;
      const pages = window.TaiwaneseWorshipSlideProduction.paginateFixedText(item.body, weights);
      item.pptPages = pages.map((page, index) => ({ ...templates[Math.min(index, templates.length - 1)], body: page.body }));
      previewPage = Math.min(previewPage, item.pptPages.length - 1);
      preview();
    });
  };
})();
