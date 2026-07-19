(function() {
  const baseEditor = editor;
  const escapeText = value => String(value || '').replace(/[&<>]/g, char => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;' })[char]);
  const fixedIds = new Set(['creed', 'lord-prayer']);
  editor = function() {
    const item = model[active];
    if (!Array.isArray(item.pptPages) || !fixedIds.has(active)) return baseEditor();
    const form = document.getElementById('editor-form');
    const templates = item.pptPages.map(page => typeof page === 'string' ? ({ body: page }) : ({ ...page }));
    if (item.type === 'dual-fixed') {
      const primaryWeights = templates.map(page => String(page.primaryBody || '').length || 1);
      const secondaryWeights = templates.map(page => String(page.secondaryBody || '').length || 1);
      form.innerHTML = `<div class="inline-note">台語與華語會分別寫入左右兩個內文框；兩欄各自依來源 PPT 的頁面文字量重排，不會合併成單一文字框。</div><div class="dual-fixed-editor"><label class="field"><span>台語全文</span><textarea data-fixed-primary>${escapeText(item.body || '')}</textarea><small>空白行會作為台語欄的優先分頁位置。</small></label><label class="field"><span>華語全文</span><textarea data-fixed-secondary>${escapeText(item.secondaryBody || '')}</textarea><small>空白行會作為華語欄的優先分頁位置。</small></label></div>`;
      const rebuild = () => {
        const primaryPages = window.TaiwaneseWorshipSlideProduction.paginateFixedText(item.body, primaryWeights);
        const secondaryPages = window.TaiwaneseWorshipSlideProduction.paginateFixedText(item.secondaryBody, secondaryWeights);
        item.pptPages = templates.map((template, index) => ({
          ...template,
          primaryBody: primaryPages[index] ? primaryPages[index].body : '',
          secondaryBody: secondaryPages[index] ? secondaryPages[index].body : ''
        }));
        previewPage = Math.min(previewPage, item.pptPages.length - 1);
        preview();
      };
      form.querySelector('[data-fixed-primary]').addEventListener('input', event => { item.body = event.target.value; rebuild(); });
      form.querySelector('[data-fixed-secondary]').addEventListener('input', event => { item.secondaryBody = event.target.value; rebuild(); });
      return;
    }
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
