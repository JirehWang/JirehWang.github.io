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
      fileCache.set(entry.fileId, library.downloadAndParse(entry, window.JSZip, window.worshipReadAPI).then(pages => library.rasterizeImportedPages(pages)).catch(error => {
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
    if (sectionId === 'offering') {
      item.title = '奉獻';
      item.kicker = '';
    } else if (sectionId === 'amen') {
      item.title = '阿們頌';
      item.kicker = '';
    } else if (sectionId === 'prayer-song') {
      item.title = entry.fileName.replace(/\.pptx$/i, '');
      item.kicker = '';
    } else if (sectionId === 'doxology') {
      item.title = `頌榮 – 第${entry.number}首`;
      item.kicker = entry.title || '';
    } else if (kind === 'hymn') {
      item.title = `聖詩 – 第 ${entry.number} 首`;
      item.kicker = entry.title || '';
    } else {
      item.title = entry.title;
    }
    item.libraryError = '';
    return { sectionId, state: 'loaded', pageCount: pages.length, entry };
  }

  window.loadPptLibraryContent = async function(sectionIds) {
    const entries = await getIndex();
    const targets = [
      ['pre-hymn-1', 'hymn'],
      ['pre-hymn-2', 'hymn'],
      ['hymn-1', 'hymn'],
      ['hymn-2', 'hymn'],
      ['doxology', 'hymn'],
      ['response', 'response'],
      ['prayer-song', 'hymn'],
      ['offering', 'hymn'],
      ['amen', 'hymn']
    ].filter(([sectionId]) => !sectionIds || sectionIds.includes(sectionId));
    return Promise.all(targets.map(([sectionId, kind]) => loadSection(sectionId, kind, entries)));
  };

  window.reloadCurrentPptLibrarySection = async function() {
    const result = await window.loadPptLibraryContent([active]);
    render();
    return result[0];
  };

  window.addEventListener('load', () => {
    window.loadPptLibraryContent(['prayer-song', 'offering', 'amen'])
      .then(() => render())
      .catch(error => console.warn('固定聖詩載入失敗：', error));
  });
})();
