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

  // Keep the PrayerPPT layout contract compatible with the shared layout editor.
  // The editor renders against a 960px-wide 16:9 canvas, so 1cqw equals 9.6pt.
  function pointsToCanvasCqw(value) {
    return Number(value) / 9.6;
  }

  function canvasCqwToPoints(value) {
    return Number(value) * 9.6;
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

  const BIBLE_BOOK_NAMES = [
    '帖撒羅尼迦後書', '帖撒羅尼迦前書', '哥林多後書', '哥林多前書', '提摩太後書', '提摩太前書', '彼得後書', '彼得前書', '約翰三書', '約翰二書', '約翰一書',
    '歷代志下', '歷代志上', '撒母耳記下', '撒母耳記上', '列王紀下', '列王紀上', '以斯拉記', '尼希米記', '以斯帖記', '耶利米書', '以西結書', '何西阿書', '約珥書', '阿摩司書', '俄巴底亞書', '約拿書', '彌迦書', '那鴻書', '哈巴谷書', '西番雅書', '哈該書', '撒迦利亞書', '瑪拉基書',
    '創世記', '出埃及記', '利未記', '民數記', '申命記', '約書亞記', '士師記', '路得記', '約伯記', '詩篇', '箴言', '傳道書', '雅歌', '以賽亞書', '耶利米哀歌', '但以理書', '馬太福音', '馬可福音', '路加福音', '約翰福音', '使徒行傳', '羅馬書', '加拉太書', '以弗所書', '腓立比書', '歌羅西書', '提多書', '腓利門書', '希伯來書', '雅各書', '猶大書', '啟示錄',
    '帖後', '帖前', '帖撒後', '帖撒前', '林後', '林前', '提前', '提後', '彼後', '彼前', '約三', '約二', '約一', '創', '出', '利', '民', '申', '書', '士', '得', '撒下', '撒上', '王下', '王上', '代下', '代上', '拉', '尼', '斯', '伯', '詩', '箴', '傳', '歌', '賽', '耶', '哀', '結', '但', '何', '珥', '摩', '俄', '拿', '彌', '鴻', '哈', '番', '該', '亞', '瑪', '太', '可', '路', '約', '徒', '羅', '加', '弗', '腓', '西', '多', '門', '來', '雅', '猶', '啟'
  ].sort((left, right) => right.length - left.length);

  const BIBLE_REFERENCE_PATTERN = new RegExp(
    `(${BIBLE_BOOK_NAMES.join('|')})\\s*(\\d+)\\s*[:：]\\s*(\\d+(?:\\s*[-~～]\\s*\\d+)?)`,
    'g'
  );

  function extractBibleReferences(lines) {
    const references = [];
    const seen = new Set();
    (Array.isArray(lines) ? lines : [lines]).forEach(line => {
      const text = String(line || '');
      BIBLE_REFERENCE_PATTERN.lastIndex = 0;
      let match;
      while ((match = BIBLE_REFERENCE_PATTERN.exec(text))) {
        const reference = `${match[1]} ${match[2]}:${match[3].replace(/\\s+/g, '')}`;
        if (!seen.has(reference)) {
          seen.add(reference);
          references.push(reference);
        }
      }
    });
    return references;
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

  function detachPagesFromLayoutGroup(state, pageIds) {
    ensureLayoutState(state);
    (pageIds || []).forEach(pageId => {
      const groupId = state.pageAssignments[pageId];
      if (!groupId || !state.groups[groupId]) return;
      state.groups[groupId].pageIds = (state.groups[groupId].pageIds || []).filter(id => id !== pageId);
      delete state.pageAssignments[pageId];
    });
    return state;
  }

  function createLayoutGroup(state, groupId, pageIds, params) {
    ensureLayoutState(state);
    detachPagesFromLayoutGroup(state, pageIds);
    const previous = state.groups[groupId] || { id: groupId, name: groupId, pageIds: [], params: {} };
    const uniquePageIds = Array.from(new Set([...(previous.pageIds || []), ...(pageIds || [])]));
    state.groups[groupId] = {
      ...previous,
      id: groupId,
      pageIds: uniquePageIds,
      params: { ...(params || {}) }
    };
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
    return { ...DEFAULT_LAYOUT_PARAMS, ...(page.layout || {}), ...groupParams };
  }

  return {
    DEFAULT_LAYOUT_PARAMS,
    normalizeColor,
    pointsToCanvasCqw,
    canvasCqwToPoints,
    isSupportedBackgroundImage,
    normalizeBackgroundImageDataUrl,
    parseRecognizedSections,
    extractBibleReferences,
    wrapTextForBox,
    splitIntoPoints,
    generateSectionPages,
    buildDeckEntries,
    createLayoutGroup,
    updateLayoutGroup,
    detachPagesFromLayoutGroup,
    layoutForPage
  };
});
