(function(root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.PrayerSlideProduction = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function() {
  const cleanText = value => String(value == null ? '' : value).replace(/<\/?[a-zA-Z0-9]+[^>]*>/g, '').trim();

  const DEFAULT_LAYOUT_PARAMS = {
    titleSize: 60,
    titleX: 8,
    titleY: 6,
    titleW: 84,
    titleH: 14,
    titleAlign: 'center',
    titleColor: '#111827',
    contentSize: 48,
    contentX: 8,
    contentY: 24,
    contentW: 84,
    contentH: 68,
    contentAlign: 'left',
    contentColor: '#1F2937',
    lineSpacing: 1.5
  };

  function normalizeColor(value, fallback = '#111111') {
    const color = String(value || '').trim().toLowerCase();
    if (/^#[0-9a-f]{6}$/.test(color)) return color;
    if (/^#[0-9a-f]{3}$/.test(color)) return `#${color.slice(1).split('').map(char => char + char).join('')}`;
    return fallback;
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

  function parseRecognizedSections(texts, model = {}) {
    const sectionsToUpdate = {};
    const numberToSectionMap = {
      1: 'silence',
      2: 'hymn-1',
      3: 'scripture',
      4: 'thanksgiving',
      5: 'repentance',
      6: 'world',
      7: 'nation',
      8: 'church',
      9: 'members',
      10: 'oneself',
      11: 'verse',
      12: 'hymn-2',
      13: 'benediction'
    };

    (Array.isArray(texts) ? texts : [texts]).forEach(sourceText => {
      let currentSection = null;
      let currentLines = [];
      let currentTitle = '';

      const collectCurrentSection = () => {
        if (!currentSection) return;
        const existing = sectionsToUpdate[currentSection];
        if (existing) {
          existing.lines.push(...currentLines);
          return;
        }
        const item = model[currentSection] || {};
        sectionsToUpdate[currentSection] = {
          title: currentTitle || item.title || item.label || '',
          lines: currentLines.slice()
        };
      };

      const sourceLines = String(sourceText || '').split('\n').map(line => line.trim()).filter(Boolean);
      if (sourceLines.length && /^\d+\.?$/.test(sourceLines[sourceLines.length - 1])) {
        sourceLines.pop();
      }

      sourceLines.forEach(trimmed => {

        // A section number must include a title. This excludes handwritten page footers such as "5.".
        const headerMatch = trimmed.match(/^(\d+)\.\s*(.+)/);
        if (headerMatch) {
          const sectionKey = numberToSectionMap[parseInt(headerMatch[1], 10)];
          if (sectionKey) {
            collectCurrentSection();
            const item = model[sectionKey] || {};
            currentSection = sectionKey;
            currentLines = [];
            currentTitle = headerMatch[2].trim() || item.label || item.title || '';
            return;
          }
        }

        if (currentSection) currentLines.push(trimmed);
      });

      // Reset at every image boundary so the next image's date/title cannot leak into this section.
      collectCurrentSection();
    });

    return sectionsToUpdate;
  }

  let textMeasureContext;
  const NATIVE_TEXT_WRAP_SAFETY = 0.92;

  function fallbackTextWidth(value) {
    return Array.from(String(value || '')).reduce((width, char) => {
      if (/\s/.test(char)) return width + 0.35;
      if (/[\u0000-\u00ff]/.test(char)) return width + 0.55;
      if (/[，。；：、？！）》」』]/.test(char)) return width + 0.55;
      return width + 1;
    }, 0);
  }

  function wrapTextForBox(value, options = {}) {
    const text = String(value == null ? '' : value);
    if (!text) return '';
    const fontSize = Math.max(1, Number(options.fontSize) || 48);
    const boxWidth = Math.max(1, Number(options.boxWidth) || 84);
    const fontFamily = String(options.fontFamily || 'Microsoft JhengHei');
    const bold = options.bold !== false;
    const browserMaxWidth = 1280 * boxWidth / 100 * NATIVE_TEXT_WRAP_SAFETY;
    const fallbackMaxWidth = 960 * boxWidth / 100 / fontSize * NATIVE_TEXT_WRAP_SAFETY;

    if (!textMeasureContext && typeof document !== 'undefined' && document.createElement) {
      const canvas = document.createElement('canvas');
      textMeasureContext = canvas.getContext && canvas.getContext('2d');
    }
    if (textMeasureContext) {
      textMeasureContext.font = `${bold ? 700 : 400} ${fontSize * 4 / 3}px "${fontFamily}"`;
    }
    const measure = candidate => textMeasureContext
      ? textMeasureContext.measureText(candidate).width
      : fallbackTextWidth(candidate);
    const maxWidth = textMeasureContext ? browserMaxWidth : fallbackMaxWidth;
    const prohibitedLineStarts = /[，。；：、？！）》」』]/;
    const lines = [];

    text.split('\n').forEach(sourceLine => {
      if (!sourceLine) {
        lines.push('');
        return;
      }
      let line = '';
      Array.from(sourceLine).forEach(char => {
        const candidate = line + char;
        if (line && measure(candidate) > maxWidth && !prohibitedLineStarts.test(char)) {
          lines.push(line);
          line = char;
        } else {
          line = candidate;
        }
      });
      lines.push(line);
    });
    return lines.join('\n');
  }

  // Split a text block into points (bullets)
  function splitIntoPoints(bodyText) {
    if (!bodyText) return [];
    const lines = bodyText.split('\n').map(line => line.trim()).filter(Boolean);
    const points = [];
    let currentPoint = '';

    lines.forEach(line => {
      // Matches: a., a), 1., 1), (1), (a), ①, etc.
      const isNewBullet = /^(?:[a-hA-H]\.|[a-hA-H]\)|[0-9]+\.|[0-9]+\)|(?:\(|（)[a-hA-H0-9]+(?:\)|）)|①|②|③|④|⑤|⑥|⑦|⑧|⑨|⑩)/.test(line);
      if (isNewBullet) {
        if (currentPoint) {
          points.push(currentPoint);
        }
        currentPoint = line;
      } else {
        if (currentPoint) {
          currentPoint += '\n' + line;
        } else {
          currentPoint = line;
        }
      }
    });

    if (currentPoint) {
      points.push(currentPoint);
    }
    return points;
  }

  // Paginate a single point so it fits in the max line limit
  function paginatePointText(pointText, maxLines = 5, contentW = 84, fontSize = 48) {
    const wrapped = wrapTextForBox(pointText, { fontSize, boxWidth: contentW });
    const lines = wrapped.split('\n');
    const pages = [];
    for (let i = 0; i < lines.length; i += maxLines) {
      pages.push(lines.slice(i, i + maxLines).join('\n'));
    }
    return pages.length ? pages : [''];
  }

  function paginateGroupedPoints(bodyText, maxLines = 5, contentW = 84, fontSize = 48) {
    const points = splitIntoPoints(bodyText);
    if (!points.length) return [];
    const pages = [];
    let currentLines = [];
    const flushCurrent = () => {
      if (!currentLines.length) return;
      pages.push(currentLines.join('\n'));
      currentLines = [];
    };

    points.forEach(point => {
      const pointLines = wrapTextForBox(point, { fontSize, boxWidth: contentW }).split('\n');
      if (pointLines.length > maxLines) {
        flushCurrent();
        for (let index = 0; index < pointLines.length; index += maxLines) {
          pages.push(pointLines.slice(index, index + maxLines).join('\n'));
        }
        return;
      }
      if (currentLines.length && currentLines.length + pointLines.length > maxLines) {
        flushCurrent();
      }
      currentLines.push(...pointLines);
    });

    flushCurrent();
    return pages;
  }

  // Build slide pages for a section based on its type
  function generateSectionPages(sectionId, item) {
    const pages = [];
    const title = item.title || item.label;

    if (item.type === 'bible') {
      // Dynamic Bible query
      const records = item.bibleRecords || [];
      // 2 verses per page is typical
      const versesPerPage = 2;
      let currentPage = [];
      records.forEach(record => {
        if (currentPage.length >= versesPerPage) {
          pages.push(currentPage);
          currentPage = [];
        }
        currentPage.push(record);
      });
      if (currentPage.length) pages.push(currentPage);

      return pages.map((pageRecords, index) => {
        // Construct reference label for this page
        const first = pageRecords[0];
        const last = pageRecords[pageRecords.length - 1];
        const ref = (first && last)
          ? `${first.chap}:${first.sec}${first.sec !== last.sec ? '-' + last.sec : ''}`
          : '';
        const pageTitle = first ? `${first.bookName} ${ref}` : title;
        
        return {
          id: `${sectionId}:${index + 1}`,
          kind: 'bible',
          title: pageTitle,
          body: pageRecords.map(r => `${r.sec} ${cleanText(r.bible_text || r.text)}`).join('\n\n'),
          layout: {}
        };
      });
    }

    if (item.type === 'list-bible') {
      // Bullet points with bible header
      const bibleText = item.bibleRecords && item.bibleRecords.map(r => `${r.sec} ${cleanText(r.bible_text || r.text)}`).join(' ') || '';
      // Page 1: The Scripture text
      if (bibleText) {
        const wrappedBible = wrapTextForBox(bibleText, { fontSize: 48, boxWidth: 84 });
        const bibleLines = wrappedBible.split('\n');
        const maxLines = 5;
        for (let i = 0; i < bibleLines.length; i += maxLines) {
          pages.push({
            kind: 'list-bible-scripture',
            title: `${title} (${item.bibleQuery || ''})`,
            body: bibleLines.slice(i, i + maxLines).join('\n'),
            layout: {}
          });
        }
      }

      // Subsequent pages: keep sub-points together under the same major section.
      if (item.body) {
        const paginatedSubpages = paginateGroupedPoints(item.body, 5, 84, 48);
        paginatedSubpages.forEach(subpage => {
          pages.push({
            kind: 'list-item',
            title: title,
            body: subpage,
            layout: {}
          });
        });
      }

      return pages.map((page, index) => ({
        ...page,
        id: `${sectionId}:${index + 1}`
      }));
    }

    if (item.type === 'list') {
      if (item.body) {
        const paginatedSubpages = paginateGroupedPoints(item.body, 5, 84, 48);
        paginatedSubpages.forEach(subpage => {
          pages.push({
            kind: 'list-item',
            title: title,
            body: subpage,
            layout: {}
          });
        });
      }

      if (!pages.length) {
        pages.push({
          kind: 'list-item',
          title: title,
          body: '',
          layout: {}
        });
      }

      return pages.map((page, index) => ({
        ...page,
        id: `${sectionId}:${index + 1}`
      }));
    }

    if (item.type === 'praise') {
      // Song pages split by blank line
      const segments = (item.body || '').split(/\n\s*\n/).map(s => s.trim()).filter(Boolean);
      segments.forEach(segment => {
        pages.push({
          kind: 'praise-lyrics',
          title: title,
          body: segment,
          layout: {}
        });
      });
      if (!pages.length) {
        pages.push({
          kind: 'praise-lyrics',
          title: title,
          body: '',
          layout: {}
        });
      }
      return pages.map((page, index) => ({
        ...page,
        id: `${sectionId}:${index + 1}`
      }));
    }

    // Default 'content' type (e.g. silence, benediction)
    const paginatedBody = paginatePointText(item.body || '', 5, 84, 48);
    paginatedBody.forEach(subpage => {
      pages.push({
        kind: 'content',
        title: title,
        body: subpage,
        layout: {}
      });
    });

    return pages.map((page, index) => ({
      ...page,
      id: `${sectionId}:${index + 1}`
    }));
  }

  function buildDeckEntries(sections, model) {
    let deckNumber = 0;
    let resolvedDecks = [];

    if (Array.isArray(sections) && sections.length > 0) {
      if (Array.isArray(sections[0])) {
        // Signature: buildDeckEntries(sections, model)
        resolvedDecks = sections.map(([sectionId, label]) => {
          const item = model[sectionId];
          const pages = generateSectionPages(sectionId, item);
          return { sectionId, label, pages };
        });
      } else if (typeof sections[0] === 'object' && sections[0] !== null) {
        // Signature: buildDeckEntries(sectionDecks)
        resolvedDecks = sections;
      }
    }

    return resolvedDecks.flatMap((section, sectionIndex) => (section.pages || []).map((page, pageIndex) => ({
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

  function layoutForPage(state, page) {
    ensureLayoutState(state);
    const groupId = state.pageAssignments[page.id];
    const groupParams = groupId && state.groups[groupId] ? state.groups[groupId].params : {};
    return { ...DEFAULT_LAYOUT_PARAMS, ...(page.layout || {}), ...groupParams };
  }

  return {
    DEFAULT_LAYOUT_PARAMS,
    normalizeColor,
    isSupportedBackgroundImage,
    normalizeBackgroundImageDataUrl,
    parseRecognizedSections,
    wrapTextForBox,
    splitIntoPoints,
    generateSectionPages,
    buildDeckEntries,
    layoutForPage
  };
});
