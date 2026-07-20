let previewPage = 0;
const safeHtml = value => String(value || '').replace(/[&<>]/g, char => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;' })[char]);
const safeAttr = value => safeHtml(value).replace(/"/g, '&quot;');

function slidePages(item, sectionId) {
  return window.PrayerSlideProduction.generateSectionPages(sectionId || active, item);
}

preview = function() {
  const item = model[active];
  const pages = slidePages(item);
  previewPage = Math.min(previewPage, Math.max(0, pages.length - 1));
  const page = pages[previewPage];
  
  const frame = document.querySelector('.slide-frame');
  const background = document.querySelector('.slide-background');
  
  document.getElementById('preview-name').textContent = item.label;
  document.getElementById('slide-count').textContent = `${previewPage + 1} / ${pages.length}`;
  
  background.style.backgroundImage = backgroundImage ? `url("${backgroundImage}")` : 'none';
  background.style.backgroundColor = backgroundColor;
  background.style.opacity = 1;
  
  const content = document.getElementById('slide-content');
  const title = safeHtml(page.title);
  
  content.innerHTML = `
    <h1 class="slide-title-elem">${title}</h1>
    <div class="slide-body-elem">${safeHtml(page.body).replace(/\n/g, '<br>')}</div>
  `;
  
  // Apply layout parameters dynamically
  const layout = window.PrayerSlideProduction.layoutForPage(window.worshipLayoutState, page);
  
  const titleElem = content.querySelector('.slide-title-elem');
  const bodyElem = content.querySelector('.slide-body-elem');
  
  if (titleElem) {
    titleElem.style.position = 'absolute';
    titleElem.style.fontSize = `${layout.titleSize / 9.6}cqw`;
    titleElem.style.left = `${layout.titleX}%`;
    titleElem.style.top = `${layout.titleY}%`;
    titleElem.style.width = `${layout.titleW}%`;
    titleElem.style.height = `${layout.titleH}%`;
    titleElem.style.textAlign = layout.titleAlign || 'center';
    titleElem.style.color = layout.titleColor || '#FFFFFF';
    titleElem.style.margin = '0';
    titleElem.style.padding = '0';
  }
  
  if (bodyElem) {
    bodyElem.style.position = 'absolute';
    bodyElem.style.fontSize = `${layout.contentSize / 9.6}cqw`;
    bodyElem.style.left = `${layout.contentX}%`;
    bodyElem.style.top = `${layout.contentY}%`;
    bodyElem.style.width = `${layout.contentW}%`;
    bodyElem.style.height = `${layout.contentH}%`;
    bodyElem.style.textAlign = layout.contentAlign || 'left';
    bodyElem.style.color = layout.contentColor || '#E0E0E0';
    bodyElem.style.lineHeight = layout.lineSpacing || 1.5;
    bodyElem.style.margin = '0';
    bodyElem.style.padding = '0';
    bodyElem.style.whiteSpace = 'pre-wrap';
  }
};

document.getElementById('page-prev').onclick = () => {
  if (window.navigateDeck) window.navigateDeck(-1);
  else { previewPage = Math.max(0, previewPage - 1); preview(); }
};

document.getElementById('page-next').onclick = () => {
  if (window.navigateDeck) window.navigateDeck(1);
  else { previewPage += 1; preview(); }
};

document.addEventListener('keydown', event => {
  const target = event.target;
  if (target && (target.matches('input, textarea, select, button, [contenteditable="true"]'))) return;
  if (event.key === 'ArrowLeft' || event.key === 'ArrowUp') {
    event.preventDefault();
    if (window.navigateDeck) window.navigateDeck(-1);
    else { previewPage = Math.max(0, previewPage - 1); preview(); }
  } else if (event.key === 'ArrowRight' || event.key === 'ArrowDown') {
    event.preventDefault();
    if (window.navigateDeck) window.navigateDeck(1);
    else { previewPage += 1; preview(); }
  }
});

const originalRender = render;
render = function() { previewPage = 0; originalRender(); };
render();
