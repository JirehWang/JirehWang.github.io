(function(root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.TaiwaneseWorshipContentGenerators = api;

  if (typeof document !== 'undefined') {
    async function generateBibleSection(sectionId, label) {
      const item = model[sectionId];
      if (!item || !item.sourceValue) return;
      const records = await api.queryBibleViaReadApi(item.sourceValue, root.FhlBibleService, root.worshipReadAPI);
      item.pptPages = root.TaiwaneseWorshipSlideProduction.buildBiblePages(
        sectionId,
        label,
        item.sourceValue,
        records,
        2
      );
    }

    root.generateCalendarContent = async function() {
      await Promise.all([
        generateBibleSection('call', '宣召'),
        generateBibleSection('scripture', '聖經'),
        generateBibleSection('verse', '聖經')
      ]);
      if (model.verse && Array.isArray(model.verse.pptPages)) {
        model.verse.pptPages.unshift({ id: 'verse:title', kind: 'section', title: '金句', body: '', layout: {} });
      }
    };
  }
})(typeof globalThis !== 'undefined' ? globalThis : this, function() {
  async function queryBibleViaReadApi(reference, bibleService, readApi) {
    if (!bibleService || typeof bibleService.parseQuery !== 'function') throw new Error('台語聖經解析器尚未載入');
    if (typeof readApi !== 'function') throw new Error('雲端讀取介面尚未載入');
    const queries = bibleService.parseQuery(reference);
    if (!queries.length) throw new Error(`無法識別經文格式：「${reference}」`);
    const responses = await Promise.all(queries.map(query => readApi('cal_queryBible', {
      book: query.short,
      chap: query.chap,
      sec: query.sec,
      version: 'tghg'
    })));
    return responses.flatMap(response => (response.records || []).map(record => ({
      ...record,
      bible_text: record.bible_text || record.text || ''
    })));
  }

  return { queryBibleViaReadApi };
});
