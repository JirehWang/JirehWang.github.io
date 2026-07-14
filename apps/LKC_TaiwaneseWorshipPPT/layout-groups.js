(function() {
  const production = window.TaiwaneseWorshipSlideProduction;
  const layoutState = { groups: {}, pageAssignments: {} };
  const pendingSelection = new Set();
  let liveParams = null;
  window.worshipLayoutState = layoutState;

  try {
    const draft = JSON.parse(localStorage.getItem('lkc-taiwanese-worship-draft') || '{}');
    if (draft.layoutState) {
      layoutState.groups = draft.layoutState.groups || {};
      layoutState.pageAssignments = draft.layoutState.pageAssignments || {};
    }
  } catch (error) {
    console.warn('版面群組草稿讀取失敗', error);
  }

  function persistLayoutState() {
    try {
      const draft = JSON.parse(localStorage.getItem('lkc-taiwanese-worship-draft') || '{}');
      localStorage.setItem('lkc-taiwanese-worship-draft', JSON.stringify({ ...draft, layoutState }));
    } catch (error) {
      console.warn('版面群組保存失敗', error);
    }
  }

  function sectionDecks() {
    return sections.map(([sectionId, label]) => {
      const generatedPages = slidePages(model[sectionId], sectionId);
      const pages = generatedPages.length ? generatedPages : [{ kind: 'section', body: '' }];
      return { sectionId, label, pages: pages.map((page, index) => ({ ...page, id: page.id || `${sectionId}:${index + 1}` })) };
    });
  }

  function deckEntries() {
    return production.buildDeckEntries(sectionDecks());
  }

  function currentDeckEntry() {
    return deckEntries().find(entry => entry.sectionId === active && entry.pageIndex === previewPage) || deckEntries()[0];
  }

  function selectedIds() {
    return Array.from(pendingSelection);
  }

  function showDeckEntry(entry) {
    if (!entry) return;
    active = entry.sectionId;
    render();
    previewPage = entry.pageIndex;
    preview();
    updateDeckNavigator();
  }

  window.navigateDeck = function(delta) {
    const deck = deckEntries();
    const current = currentDeckEntry();
    const index = Math.max(0, deck.findIndex(entry => entry.id === current.id));
    showDeckEntry(deck[Math.max(0, Math.min(deck.length - 1, index + delta))]);
  };

  function groupLabel(pageId) {
    const groupId = layoutState.pageAssignments[pageId];
    const group = groupId && layoutState.groups[groupId];
    return group ? (group.name || group.id) : '未分組';
  }

  function renderDeckNavigator() {
    const decks = sectionDecks();
    document.getElementById('flow-list').innerHTML = `<div class="deck-chapters">${decks.map((section, sectionIndex) => `
      <details class="deck-chapter" data-deck-section="${section.sectionId}" ${section.sectionId === active ? 'open' : ''}>
        <summary><input type="checkbox" data-layout-section="${section.sectionId}" aria-label="勾選 ${section.label} 全章" ${section.pages.every(page => pendingSelection.has(page.id)) ? 'checked' : ''}><span><b>${String(sectionIndex + 1).padStart(2, '0')}</b>${section.label}</span><small>${section.pages.length} 頁</small></summary>
        <div class="deck-page-list">${section.pages.map((page, pageIndex) => `<label data-deck-page-row="${page.id}"><input type="checkbox" data-layout-page="${page.id}" ${pendingSelection.has(page.id) ? 'checked' : ''}><button type="button" data-deck-page="${page.id}">第 ${pageIndex + 1} 頁</button><small>${groupLabel(page.id)}</small></label>`).join('')}</div>
      </details>`).join('')}</div>`;

    document.querySelectorAll('[data-deck-page]').forEach(button => button.onclick = () => showDeckEntry(deckEntries().find(entry => entry.id === button.dataset.deckPage)));
    document.querySelectorAll('[data-layout-page]').forEach(box => box.onchange = () => {
      if (box.checked) pendingSelection.add(box.dataset.layoutPage); else pendingSelection.delete(box.dataset.layoutPage);
      syncSectionCheckbox(box.closest('[data-deck-section]'));
      if (box.checked) {
        showDeckEntry(deckEntries().find(entry => entry.id === box.dataset.layoutPage));
      }
    });
    document.querySelectorAll('[data-layout-section]').forEach(box => box.onchange = event => {
      event.stopPropagation();
      const chapter = box.closest('[data-deck-section]');
      chapter.querySelectorAll('[data-layout-page]').forEach(pageBox => {
        pageBox.checked = box.checked;
        if (box.checked) pendingSelection.add(pageBox.dataset.layoutPage); else pendingSelection.delete(pageBox.dataset.layoutPage);
      });
      if (box.checked) {
        chapter.open = true;
        const first = chapter.querySelector('[data-deck-page]');
        if (first) showDeckEntry(deckEntries().find(entry => entry.id === first.dataset.deckPage));
      }
    });
    document.querySelectorAll('[data-deck-section]').forEach(syncSectionCheckbox);
    updateDeckNavigator();
  }

  function syncSectionCheckbox(chapter) {
    if (!chapter) return;
    const sectionBox = chapter.querySelector('[data-layout-section]');
    const pageBoxes = Array.from(chapter.querySelectorAll('[data-layout-page]'));
    sectionBox.checked = pageBoxes.length > 0 && pageBoxes.every(box => box.checked);
    sectionBox.indeterminate = pageBoxes.some(box => box.checked) && !sectionBox.checked;
  }

  function updateDeckNavigator() {
    const deck = deckEntries();
    const current = currentDeckEntry();
    document.querySelectorAll('[data-deck-page-row]').forEach(row => row.classList.toggle('is-current', row.dataset.deckPageRow === current.id));
    document.querySelectorAll('[data-deck-section]').forEach(chapter => chapter.classList.toggle('is-current', chapter.dataset.deckSection === current.sectionId));
    const count = document.getElementById('slide-count');
    if (count && current) count.textContent = `${current.deckNumber} / ${deck.length}`;
    const chapterName = document.getElementById('preview-name');
    if (chapterName && current) chapterName.textContent = `${current.sectionLabel}・第 ${current.pageIndex + 1} 頁`;
  }

  function numberValue(id, fallback) {
    const input = document.getElementById(id);
    if (!input || input.value === '') return fallback;
    return Number(input.value);
  }

  function paramsFromForm() {
    return {
      titleSize: numberValue('lg-title-size', 60), titleX: numberValue('lg-title-x', 10), titleY: numberValue('lg-title-y', 6), titleW: numberValue('lg-title-w', 80), titleH: numberValue('lg-title-h', 16), titleAlign: document.getElementById('lg-title-align').value, titleColor: production.normalizeColor(document.getElementById('lg-title-color').value, '#111111'),
      contentSize: numberValue('lg-content-size', 48), contentX: numberValue('lg-content-x', 8), contentY: numberValue('lg-content-y', 24), contentW: numberValue('lg-content-w', 84), contentH: numberValue('lg-content-h', 68), contentAlign: document.getElementById('lg-content-align').value, contentColor: production.normalizeColor(document.getElementById('lg-content-color').value, '#111111'), lineSpacing: numberValue('lg-line-spacing', 1.5)
    };
  }

  function computedColor(value, fallback = '#111111') {
    const match = String(value || '').match(/rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i);
    if (!match) return production.normalizeColor(value, fallback);
    return `#${match.slice(1, 4).map(channel => Math.max(0, Math.min(255, Number(channel))).toString(16).padStart(2, '0')).join('')}`;
  }

  function canvasParams() {
    const frame = document.querySelector('.slide-frame');
    const content = document.getElementById('slide-content');
    if (!frame || !content) return null;
    const frameRect = frame.getBoundingClientRect();
    if (!frameRect.width || !frameRect.height) return null;
    const measure = (element, prefix, fallback) => {
      if (!element) return fallback;
      const rect = element.getBoundingClientRect();
      const style = getComputedStyle(element);
      const fontPx = parseFloat(style.fontSize) || 0;
      const linePx = parseFloat(style.lineHeight) || fontPx;
      const rawX = (rect.left - frameRect.left) / frameRect.width * 100;
      const rawW = rect.width / frameRect.width * 100;
      const edgeBuffer = 1.2;
      const bufferedX = style.textAlign === 'right' ? rawX - edgeBuffer * 2 : style.textAlign === 'center' ? rawX - edgeBuffer : rawX;
      const stableX = Math.max(0, bufferedX);
      const stableW = Math.min(100 - stableX, rawW + edgeBuffer * 2);
      return {
        [`${prefix}Size`]: Number((fontPx / frameRect.width * 100 * 7.2).toFixed(1)),
        [`${prefix}X`]: Number(stableX.toFixed(1)),
        [`${prefix}Y`]: Number(((rect.top - frameRect.top) / frameRect.height * 100).toFixed(1)),
        [`${prefix}W`]: Number(stableW.toFixed(1)),
        [`${prefix}H`]: Number((rect.height / frameRect.height * 100).toFixed(1)),
        [`${prefix}Align`]: style.textAlign || fallback[`${prefix}Align`],
        [`${prefix}Color`]: computedColor(style.color, fallback[`${prefix}Color`]),
        ...(prefix === 'content' ? { lineSpacing: Number((linePx / Math.max(fontPx, 1)).toFixed(2)) } : {})
      };
    };
    const titleFallback = { titleSize: 60, titleX: 10, titleY: 6, titleW: 80, titleH: 16, titleAlign: 'center', titleColor: '#111111' };
    const contentFallback = { contentSize: 48, contentX: 8, contentY: 24, contentW: 84, contentH: 68, contentAlign: 'left', contentColor: '#111111', lineSpacing: 1.5 };
    const importedObjects = Array.from(content.querySelectorAll('.ppt-object-text'));
    if (importedObjects.length) {
      const measureImported = (role, prefix, fallback) => {
        const objects = importedObjects.filter(element => element.dataset.pptRole === role);
        if (!objects.length) return fallback;
        const rects = objects.map(element => element.getBoundingClientRect());
        const left = Math.min(...rects.map(rect => rect.left));
        const top = Math.min(...rects.map(rect => rect.top));
        const right = Math.max(...rects.map(rect => rect.right));
        const bottom = Math.max(...rects.map(rect => rect.bottom));
        const style = getComputedStyle(objects[0]);
        const fontPx = parseFloat(style.fontSize) || 0;
        return {
          [`${prefix}Size`]: Number((fontPx / frameRect.width * 100 * 7.2).toFixed(1)),
          [`${prefix}X`]: Number(((left - frameRect.left) / frameRect.width * 100).toFixed(1)),
          [`${prefix}Y`]: Number(((top - frameRect.top) / frameRect.height * 100).toFixed(1)),
          [`${prefix}W`]: Number(((right - left) / frameRect.width * 100).toFixed(1)),
          [`${prefix}H`]: Number(((bottom - top) / frameRect.height * 100).toFixed(1)),
          [`${prefix}Align`]: style.textAlign || fallback[`${prefix}Align`],
          [`${prefix}Color`]: computedColor(style.color, fallback[`${prefix}Color`]),
          ...(prefix === 'content' ? { lineSpacing: Number((parseFloat(style.lineHeight) / Math.max(fontPx, 1)).toFixed(2)) || 1.05 } : {})
        };
      };
      return {
        ...measureImported('title', 'title', titleFallback),
        ...measureImported('content', 'content', contentFallback)
      };
    }
    return {
      ...measure(content.querySelector('h1'), 'title', titleFallback),
      ...measure(content.querySelector('.body, p'), 'content', contentFallback)
    };
  }

  function populateForm(params) {
    Object.entries(params || {}).forEach(([key, value]) => {
      const input = document.getElementById('lg-' + key.replace(/[A-Z]/g, letter => '-' + letter.toLowerCase()));
      if (input && value != null && value !== '') input.value = value;
    });
  }

  function parameterFields() {
    return `<div class="layout-parameter-tabs"><button type="button" class="is-active" data-layout-tab="title">標題</button><button type="button" data-layout-tab="content">內文</button></div>
      <div class="layout-params" data-layout-pane="title"><label>字級<input id="lg-title-size" type="number" value="60"></label><label>X<input id="lg-title-x" type="number" value="10"></label><label>Y<input id="lg-title-y" type="number" value="6"></label><label>寬<input id="lg-title-w" type="number" value="80"></label><label>高<input id="lg-title-h" type="number" value="16"></label><label>對齊<select id="lg-title-align"><option value="center">置中</option><option value="left">靠左</option><option value="right">靠右</option></select></label><label>文字顏色<input id="lg-title-color" type="color" value="#111111"></label></div>
      <div class="layout-params is-hidden" data-layout-pane="content"><label>字級<input id="lg-content-size" type="number" value="48"></label><label>X<input id="lg-content-x" type="number" value="8"></label><label>Y<input id="lg-content-y" type="number" value="24"></label><label>寬<input id="lg-content-w" type="number" value="84"></label><label>高<input id="lg-content-h" type="number" value="68"></label><label>對齊<select id="lg-content-align"><option value="left">靠左</option><option value="center">置中</option><option value="right">靠右</option></select></label><label>行距<input id="lg-line-spacing" type="number" value="1.5" step="0.1"></label><label>文字顏色<input id="lg-content-color" type="color" value="#111111"></label></div>`;
  }

  function renderFloatingPanel() {
    const panel = document.getElementById('layout-floating-panel');
    const groups = Object.values(layoutState.groups);
    panel.innerHTML = `<header><div><small>版面參數</small><strong>調整勾選頁面</strong></div><button type="button" id="layout-panel-close" aria-label="關閉版面參數">×</button></header>
      <div class="floating-group-fields"><label>群組名稱<input id="layout-group-name" placeholder="例如：經文頁"></label><label>載入群組<select id="layout-group-existing"><option value="">新增群組</option>${groups.map(group => `<option value="${group.id}">${group.name || group.id}</option>`).join('')}</select></label></div>
      ${parameterFields()}
      <footer><button type="button" class="button quiet" id="layout-detach">解除群組</button><button type="button" class="button primary" id="layout-save-group">儲存參數組</button></footer>`;

    document.getElementById('layout-panel-close').onclick = () => panel.classList.add('is-hidden');
    enablePanelDragging(panel);
    document.querySelectorAll('[data-layout-tab]').forEach(button => button.onclick = () => {
      document.querySelectorAll('[data-layout-tab]').forEach(item => item.classList.toggle('is-active', item === button));
      document.querySelectorAll('[data-layout-pane]').forEach(pane => pane.classList.toggle('is-hidden', pane.dataset.layoutPane !== button.dataset.layoutTab));
    });
    document.querySelectorAll('.layout-params input, .layout-params select').forEach(input => input.addEventListener('input', () => { liveParams = paramsFromForm(); preview(); }));
    document.getElementById('layout-group-existing').onchange = event => loadGroup(event.target.value);
    document.getElementById('layout-save-group').onclick = saveGroup;
    document.getElementById('layout-detach').onclick = detachSelection;
  }

  function openFloatingPanel(syncWithCanvas = true) {
    const panel = document.getElementById('layout-floating-panel');
    panel.classList.remove('is-hidden');
    if (syncWithCanvas) {
      liveParams = null;
      populateForm(canvasParams());
    }
  }

  function enablePanelDragging(panel) {
    const handle = panel.querySelector('header');
    handle.onpointerdown = event => {
      if (event.target.closest('button')) return;
      event.preventDefault();
      const rect = panel.getBoundingClientRect();
      const offsetX = event.clientX - rect.left;
      const offsetY = event.clientY - rect.top;
      panel.style.left = `${rect.left}px`;
      panel.style.top = `${rect.top}px`;
      panel.style.right = 'auto';
      panel.style.bottom = 'auto';
      handle.setPointerCapture(event.pointerId);
      handle.onpointermove = moveEvent => {
        const maxLeft = Math.max(0, window.innerWidth - panel.offsetWidth);
        const maxTop = Math.max(0, window.innerHeight - panel.offsetHeight);
        panel.style.left = `${Math.max(0, Math.min(maxLeft, moveEvent.clientX - offsetX))}px`;
        panel.style.top = `${Math.max(0, Math.min(maxTop, moveEvent.clientY - offsetY))}px`;
      };
      handle.onpointerup = endEvent => {
        handle.releasePointerCapture(endEvent.pointerId);
        handle.onpointermove = null;
        handle.onpointerup = null;
      };
    };
  }

  function loadGroup(groupId) {
    const group = layoutState.groups[groupId];
    if (!group) return;
    document.getElementById('layout-group-name').value = group.name || group.id;
    pendingSelection.clear();
    group.pageIds.forEach(pageId => pendingSelection.add(pageId));
    renderDeckNavigator();
    showDeckEntry(deckEntries().find(entry => entry.id === group.pageIds[0]));
    openFloatingPanel(false);
    populateForm(group.params || {});
    liveParams = { ...(group.params || {}) };
    preview();
  }

  function saveGroup() {
    const pageIds = selectedIds();
    const existingId = document.getElementById('layout-group-existing').value;
    const name = document.getElementById('layout-group-name').value.trim();
    if (!name || pageIds.length === 0) return status('請輸入群組名稱並勾選至少一頁');
    const group = production.createLayoutGroup(layoutState, existingId || `layout-${Date.now()}`, pageIds, paramsFromForm());
    group.name = name;
    liveParams = null;
    persistLayoutState();
    renderDeckNavigator();
    renderFloatingPanel();
    openFloatingPanel();
    preview();
    status(`已儲存版面群組：${name}`);
  }

  function detachSelection() {
    production.detachPagesFromLayoutGroup(layoutState, selectedIds());
    liveParams = null;
    persistLayoutState();
    renderDeckNavigator();
    preview();
    status('已解除所選頁面的版面群組');
  }

  window.applyPageLayoutToPreview = function(content, page) {
    if (!page.id) page.id = `${active}:${previewPage + 1}`;
    const stored = production.layoutForPage(layoutState, page);
    const params = selectedIds().includes(page.id) && liveParams ? { ...stored, ...liveParams } : stored;
    const importedObjects = Array.from(content.querySelectorAll('.ppt-object-text'));
    if (importedObjects.length) {
      const applyImportedRole = (role, prefix) => {
        const objects = importedObjects.filter(element => element.dataset.pptRole === role);
        if (!objects.length || params[`${prefix}X`] == null) return;
        const sourceRects = objects.map(element => ({
          element,
          x: Number(element.dataset.sourceX), y: Number(element.dataset.sourceY),
          w: Number(element.dataset.sourceW), h: Number(element.dataset.sourceH)
        }));
        const sourceX = Math.min(...sourceRects.map(rect => rect.x));
        const sourceY = Math.min(...sourceRects.map(rect => rect.y));
        const sourceRight = Math.max(...sourceRects.map(rect => rect.x + rect.w));
        const sourceBottom = Math.max(...sourceRects.map(rect => rect.y + rect.h));
        const scaleX = Number(params[`${prefix}W`]) / Math.max(sourceRight - sourceX, 0.01);
        const scaleY = Number(params[`${prefix}H`]) / Math.max(sourceBottom - sourceY, 0.01);
        sourceRects.forEach(rect => {
          rect.element.style.left = `${Number(params[`${prefix}X`]) + (rect.x - sourceX) * scaleX}%`;
          rect.element.style.top = `${Number(params[`${prefix}Y`]) + (rect.y - sourceY) * scaleY}%`;
          rect.element.style.width = `${rect.w * scaleX}%`;
          rect.element.style.height = `${rect.h * scaleY}%`;
          rect.element.style.textAlign = params[`${prefix}Align`] || rect.element.style.textAlign;
          if (prefix === 'content' && params.lineSpacing) rect.element.style.lineHeight = params.lineSpacing;
          rect.element.querySelectorAll('[data-source-font-size]').forEach(run => {
            const sourceBaseSize = Number(rect.element.dataset.sourceFontSize) || 18;
            const relativeSize = (Number(run.dataset.sourceFontSize) || sourceBaseSize) / sourceBaseSize;
            run.style.fontSize = `${Number(params[`${prefix}Size`]) / 7.2 * relativeSize}cqw`;
            if (params[`${prefix}Color`]) run.style.color = production.normalizeColor(params[`${prefix}Color`], '#111111');
          });
        });
      };
      applyImportedRole('title', 'title');
      applyImportedRole('content', 'content');
      return;
    }
    const title = content.querySelector('h1');
    const body = content.querySelector('.body, p');
    const place = (element, prefix) => {
      if (!element) return;
      if (params[`${prefix}Color`]) element.style.color = production.normalizeColor(params[`${prefix}Color`], '#111111');
      if (params[`${prefix}X`] == null) return;
      element.style.position = 'absolute';
      element.style.left = `${params[`${prefix}X`]}%`;
      element.style.top = `${params[`${prefix}Y`]}%`;
      element.style.width = `${params[`${prefix}W`]}%`;
      element.style.height = `${params[`${prefix}H`]}%`;
      element.style.margin = '0';
      element.style.textAlign = params[`${prefix}Align`] || '';
      if (params[`${prefix}Size`]) element.style.fontSize = `${params[`${prefix}Size`] / 7.2}cqw`;
    };
    place(title, 'title');
    place(body, 'content');
    if (body && params.lineSpacing) body.style.lineHeight = params.lineSpacing;
  };

  const basePreview = preview;
  preview = function() { basePreview(); updateDeckNavigator(); };

  document.getElementById('save-draft').onclick = () => {
    localStorage.setItem('lkc-taiwanese-worship-draft', JSON.stringify({ model, backgroundColor, backgroundImage, syncHymnOpacity: window.isHymnOpacitySyncEnabled(), layoutState }));
    status('已儲存內容與版面群組');
  };

  document.getElementById('layout-panel-open').onclick = openFloatingPanel;

  flow = renderDeckNavigator;
  renderFloatingPanel();
  render();
})();
