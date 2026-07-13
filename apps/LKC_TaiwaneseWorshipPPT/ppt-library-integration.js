(function() {
  const library = window.TaiwaneseWorshipPptxLibrary;
  const fileCache = new Map();
  let indexPromise = null;

  async function getIndex() {
    if (!indexPromise) {
      indexPromise = (async function() {
        const result = await window.worshipReadAPI('cal_getPptLibraryIndex', {});
        if (!result || !Array.isArray(result.data)) throw new Error('PPT 資料庫索引格式不正確');
        return result.data;
      })().catch(error => {
        indexPromise = null;
        throw error;
      });
    }
    return indexPromise;
  }

  function pagesForEntry(entry) {
    if (!fileCache.has(entry.fileId)) {
      fileCache.set(entry.fileId, library.downloadAndParse(entry, window.JSZip, window.worshipReadAPI).catch(error => {
        fileCache.delete(entry.fileId);
        throw error;
      }));
    }
    return fileCache.get(entry.fileId);
  }

  async function loadSection(sectionId, kind, entries) {
    const item = model[sectionId];
    const number = library.normalizeLibraryNumber(item && item.sourceValue);
    if (!item || !number) return { sectionId, state: 'empty' };
    const entry = library.findLibraryEntry(entries, kind, number);
    if (!entry) {
      delete item.pptPages;
      item.libraryError = `資料庫找不到 ${kind === 'hymn' ? '聖詩' : '啟應文'} ${number}`;
      return { sectionId, state: 'missing', message: item.libraryError };
    }
    if (item.libraryFileId === entry.fileId && Array.isArray(item.pptPages) && item.pptPages.length) {
      return { sectionId, state: 'cached', pageCount: item.pptPages.length };
    }
    const pages = await pagesForEntry(entry);
    item.pptPages = pages.map((page, index) => ({ ...page, id: `${sectionId}:${index + 1}` }));
    item.libraryFileId = entry.fileId;
    item.libraryEntry = { kind: entry.kind, number: entry.number, title: entry.title, fileName: entry.fileName };
    item.title = kind === 'hymn' ? entry.fileName.replace(/\.pptx$/i, '') : entry.title;
    item.libraryError = '';
    return { sectionId, state: 'loaded', pageCount: pages.length, entry };
  }

  window.loadPptLibraryContent = async function(sectionIds) {
    const entries = await getIndex();
    const targets = [
      ['hymn-1', 'hymn'],
      ['hymn-2', 'hymn'],
      ['doxology', 'hymn'],
      ['response', 'response']
    ].filter(([sectionId]) => !sectionIds || sectionIds.includes(sectionId));
    return Promise.all(targets.map(([sectionId, kind]) => loadSection(sectionId, kind, entries)));
  };

  window.reloadCurrentPptLibrarySection = async function() {
    const result = await window.loadPptLibraryContent([active]);
    render();
    return result[0];
  };
})();
