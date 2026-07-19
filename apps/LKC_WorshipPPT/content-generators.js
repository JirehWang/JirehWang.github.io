(function(root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.TaiwaneseWorshipContentGenerators = api;

  if (typeof document !== 'undefined') {
    async function generateBibleSection(config) {
      const sectionId = config.sectionId;
      const label = config.label;
      const item = model[sectionId];
      if (!item || !item.sourceValue) return;
      const profile = root.activeWorshipTemplateProfile || {};
      const versions = Array.isArray(config.versions) && config.versions.length
        ? config.versions
        : (profile.bibleVersions || ['tghg']);
      const recordsByVersion = await Promise.all(versions.map(version =>
        api.queryBibleViaReadApi(item.sourceValue, root.FhlBibleService, root.worshipReadAPI, version)
      ));
      item.pptPages = recordsByVersion.flatMap((records, versionIndex) =>
        root.TaiwaneseWorshipSlideProduction.buildBiblePages(
          sectionId,
          label,
          item.sourceValue,
          records,
          2,
          {
            languageLabel: Array.isArray(config.languageLabels) ? config.languageLabels[versionIndex] : '',
            bibleVersion: versions[versionIndex]
          }
        )
      ).map((page, index) => ({ ...page, id: `${sectionId}:${index + 1}` }));
      if (config.prependTitle) {
        item.pptPages.unshift({ id: `${sectionId}:title`, kind: 'section', title: config.prependTitle, body: '', layout: {} });
      }
    }

    root.generateCalendarContent = async function() {
      const profile = root.activeWorshipTemplateProfile || {};
      const configs = Array.isArray(profile.bibleSections) && profile.bibleSections.length
        ? profile.bibleSections
        : [
            { sectionId: 'call', label: '宣召', versions: ['tghg'] },
            { sectionId: 'scripture', label: '聖經', versions: ['tghg'] },
            { sectionId: 'verse', label: '聖經', versions: ['tghg'], prependTitle: '金句' }
          ];
      await Promise.all(configs.map(generateBibleSection));
    };
  }
})(typeof globalThis !== 'undefined' ? globalThis : this, function() {
  async function queryBibleViaReadApi(reference, bibleService, readApi, version = 'tghg') {
    if (!bibleService || typeof bibleService.parseQuery !== 'function') throw new Error('台語聖經解析器尚未載入');
    if (typeof readApi !== 'function') throw new Error('雲端讀取介面尚未載入');
    const queries = bibleService.parseQuery(reference);
    if (!queries.length) throw new Error(`無法識別經文格式：「${reference}」`);
    const results = await Promise.all(queries.map(async query => ({
      query,
      response: await readApi('cal_queryBible', {
        book: query.short,
        chap: query.chap,
        sec: query.sec,
        version
      })
    })));
    return results.flatMap(({ query, response }) => (response.records || []).map(record => ({
      ...record,
      bible_text: record.bible_text || record.text || '',
      ...(query.bookName ? {
        queryBookName: query.bookName,
        queryChap: query.chap,
        querySec: query.sec,
        queryGroupKey: `${query.bookName}_${query.chap}_${query.sec}`
      } : {})
    })));
  }

  return { queryBibleViaReadApi };
});
