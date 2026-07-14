(function(root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.TaiwaneseWorshipSlideProduction = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function() {
  const cleanText = value => String(value == null ? '' : value).replace(/<\/?[a-zA-Z0-9]+[^>]*>/g, '').trim();

  function normalizeColor(value, fallback = '#111111') {
    const color = String(value || '').trim().toLowerCase();
    if (/^#[0-9a-f]{6}$/.test(color)) return color;
    if (/^#[0-9a-f]{3}$/.test(color)) return `#${color.slice(1).split('').map(char => char + char).join('')}`;
    return normalizeColor(fallback, '#111111');
  }

  function isSupportedBackgroundImage(file) {
    if (!file) return false;
    const type = String(file.type || '').toLowerCase();
    const name = String(file.name || '').toLowerCase();
    return ['image/png', 'image/jpeg', 'image/webp'].includes(type)
      || (!type && /\.(png|jpe?g|webp)$/.test(name));
  }

  function normalizeBackgroundImageDataUrl(value) {
    const dataUrl = String(value || '').trim();
    return /^data:image\/(?:png|jpeg|webp);base64,[a-z0-9+/=\s]+$/i.test(dataUrl) ? dataUrl : '';
  }

  function buildBiblePages(sectionId, label, reference, records, versesPerPage = 2) {
    const safeRecords = Array.isArray(records) ? records : [];
    const pages = [];
    for (let index = 0; index < safeRecords.length; index += versesPerPage) {
      const pageRecords = safeRecords.slice(index, index + versesPerPage);
      pages.push({
        id: `${sectionId}:${pages.length + 1}`,
        kind: 'scripture',
        title: `${label}－${reference}`,
        body: pageRecords.map(record => `${record.sec} ${cleanText(record.bible_text || record.text)}`).join('\n\n'),
        layout: {}
      });
    }
    return pages;
  }

  function composeLibraryPages(item) {
    const pages = (Array.isArray(item && item.pptPages) ? item.pptPages : []).map(page =>
      typeof page === 'string' ? ({ kind: 'liturgical', body: page }) : ({ kind: 'liturgical', ...page })
    );
    return item && item.includeSectionTitle ? [{ kind: 'section' }, ...pages] : pages;
  }

  function applyFixedLibraryDefaults(model) {
    const fixed = [
      ['prayer-song', '261', false],
      ['offering', '306B', true],
      ['amen', '522', false]
    ];
    fixed.forEach(([sectionId, sourceValue, includeSectionTitle]) => {
      if (!model || !model[sectionId]) return;
      model[sectionId].sourceValue = sourceValue;
      model[sectionId].includeSectionTitle = includeSectionTitle;
    });
    return model;
  }

  function paginateFixedText(value, pageWeights) {
    const paragraphs = String(value || '').split(/\n\s*\n/).map(part => part.trim()).filter(Boolean);
    const weights = (Array.isArray(pageWeights) && pageWeights.length ? pageWeights : [1]).map(weight => Math.max(1, Number(weight) || 1));
    if (!paragraphs.length) return weights.map(() => ({ body: '' }));
    const pageCount = Math.min(weights.length, paragraphs.length);
    const activeWeights = weights.slice(0, pageCount);
    const totalWeight = activeWeights.reduce((sum, weight) => sum + weight, 0);
    const totalLength = paragraphs.join('\n\n').length;
    const pages = [];
    let paragraphIndex = 0;
    for (let pageIndex = 0; pageIndex < pageCount; pageIndex += 1) {
      const remainingPages = pageCount - pageIndex - 1;
      const target = totalLength * activeWeights[pageIndex] / totalWeight;
      const parts = [];
      let length = 0;
      while (paragraphIndex < paragraphs.length - remainingPages) {
        const paragraph = paragraphs[paragraphIndex];
        parts.push(paragraph);
        paragraphIndex += 1;
        length += paragraph.length + (parts.length > 1 ? 2 : 0);
        if (length >= target) break;
      }
      pages.push({ body: parts.join('\n\n') });
    }
    return pages;
  }

  function buildDeckEntries(sectionDecks) {
    let deckNumber = 0;
    return (sectionDecks || []).flatMap((section, sectionIndex) => (section.pages || []).map((page, pageIndex) => ({
      ...page,
      id: page.id || `${section.sectionId}:${pageIndex + 1}`,
      sectionId: section.sectionId,
      sectionLabel: section.label,
      sectionIndex,
      pageIndex,
      deckNumber: ++deckNumber
    })));
  }

  function ensureLayoutState(state) {
    if (!state.groups) state.groups = {};
    if (!state.pageAssignments) state.pageAssignments = {};
    return state;
  }

  function detachPagesFromLayoutGroup(state, pageIds) {
    ensureLayoutState(state);
    (pageIds || []).forEach(pageId => {
      const groupId = state.pageAssignments[pageId];
      if (!groupId || !state.groups[groupId]) return;
      state.groups[groupId].pageIds = state.groups[groupId].pageIds.filter(id => id !== pageId);
      delete state.pageAssignments[pageId];
    });
    return state;
  }

  function createLayoutGroup(state, groupId, pageIds, params) {
    ensureLayoutState(state);
    detachPagesFromLayoutGroup(state, pageIds);
    const previous = state.groups[groupId] || { id: groupId, name: groupId, pageIds: [], params: {} };
    const uniquePageIds = Array.from(new Set([...(previous.pageIds || []), ...(pageIds || [])]));
    state.groups[groupId] = { ...previous, id: groupId, pageIds: uniquePageIds, params: { ...(params || {}) } };
    uniquePageIds.forEach(pageId => { state.pageAssignments[pageId] = groupId; });
    return state.groups[groupId];
  }

  function updateLayoutGroup(state, groupId, params) {
    ensureLayoutState(state);
    if (!state.groups[groupId]) throw new Error(`找不到版面群組：${groupId}`);
    state.groups[groupId].params = { ...(params || {}) };
    return state.groups[groupId];
  }

  function layoutForPage(state, page) {
    ensureLayoutState(state);
    const groupId = state.pageAssignments[page.id];
    const groupParams = groupId && state.groups[groupId] ? state.groups[groupId].params : {};
    return { ...groupParams, ...(page.layout || {}) };
  }

  return {
    normalizeColor,
    isSupportedBackgroundImage,
    normalizeBackgroundImageDataUrl,
    buildBiblePages,
    composeLibraryPages,
    applyFixedLibraryDefaults,
    buildDeckEntries,
    paginateFixedText,
    createLayoutGroup,
    updateLayoutGroup,
    detachPagesFromLayoutGroup,
    layoutForPage
  };
});
