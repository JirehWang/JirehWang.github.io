let previewPage = 0;
const safeHtml = value => String(value || '').replace(/[&<>]/g, char => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;' })[char]);
const safeAttr = value => safeHtml(value).replace(/"/g, '&quot;');
function renderImportedPptPage(page, item) {
  const imageOpacity = Math.max(0, Math.min(1, Number(item.opacity || 60) / 100));
  return `<div class="ppt-import-layer">${(page.objects || []).map((object, objectIndex) => {
    const geometry = `left:${object.x}%;top:${object.y}%;width:${object.w}%;height:${object.h}%`;
    if (object.type === 'image') {
      return `<img class="ppt-object-image" src="${safeAttr(object.src || '')}" alt="" style="${geometry};opacity:${imageOpacity}">`;
    }
    const runs = Array.isArray(object.runs) && object.runs.length ? object.runs : [{ text: object.text || '' }];
    const runHtml = runs.map(run => {
      const styles = [
        run.fontSize ? `font-size:${run.fontSize / 9.6}cqw` : '',
        run.fontFamily ? `font-family:${safeAttr(run.fontFamily)}` : '',
        run.color ? `color:${safeAttr(run.color)}` : '',
        run.bold ? 'font-weight:700' : '',
        run.italic ? 'font-style:italic' : '',
        run.underline ? 'text-decoration:underline' : ''
      ].filter(Boolean).join(';');
      return `<span data-source-font-size="${run.fontSize || object.fontSize || 18}" style="${styles}">${safeHtml(run.text)}</span>`;
    }).join('');
    return `<div class="ppt-object-text" data-ppt-role="${object.role || 'content'}" data-ppt-index="${objectIndex}" data-source-x="${object.x}" data-source-y="${object.y}" data-source-w="${object.w}" data-source-h="${object.h}" data-source-font-size="${object.fontSize || 18}" style="${geometry};text-align:${object.align || 'left'};align-items:${object.verticalAlign || 'start'};font-size:${(object.fontSize || 18) / 9.6}cqw;font-family:${safeAttr(object.fontFamily || 'Microsoft JhengHei')};color:${safeAttr(object.color || '#000000')};font-weight:${object.bold ? 700 : 400}"><span class="ppt-text-runs">${runHtml}</span></div>`;
  }).join('')}</div>`;
}
function slidePages(item) {
  if (Array.isArray(item.pptPages)) return item.pptPages.map(page => typeof page === 'string' ? ({ kind:'liturgical', body:page }) : ({ kind:'liturgical', ...page }));
  if (item.type === 'fixed') {
    const kind = active === 'creed' || active === 'lord-prayer' ? 'liturgical' : 'content';
    return (item.body || '').split(/\n\s*\n/).filter(Boolean).map(body => ({ kind, body }));
  }
  if (item.type === 'cover') return [{ kind:'cover' }];
  if (item.type === 'praise') {
    const lyrics = (item.body || '').split(/\n\s*\n/).filter(Boolean).map(body => ({ kind:'praise-lyrics', body }));
    return [{ kind:'praise-title' }, ...lyrics];
  }
  if (item.type === 'manual') {
    return (item.body || '').split(/\n\s*\n/).filter(Boolean).map(body => ({ kind:'content', body }));
  }
  if (item.type === 'hymn') return [{ kind:'section' }, { kind:'score' }];
  if (item.type === 'fixed-title') return [{ kind:'section' }, { kind:'score' }];
  if (item.type === 'title') return [{ kind:'section' }];
  return [{ kind:'content', body:item.body || '' }];
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
  background.style.backgroundImage = 'none';
  background.style.backgroundColor = backgroundColor;
  background.style.opacity = 1;
  const content = document.getElementById('slide-content');
  const title = safeHtml(page.title || item.title || item.label);
  const kicker = safeHtml(item.kicker || '');
  if (page.kind === 'ppt-import') {
    content.className = 'slide-content template-ppt-import';
    content.innerHTML = renderImportedPptPage(page, item);
  }
  else if (page.kind === 'cover') {
    const date = document.getElementById('service-date').value;
    const [year, month, day] = date ? date.split('-') : [];
    const formatted = date ? `主後${year}年${month}月${day}日` : '';
    content.className = 'slide-content template-cover';
    content.innerHTML = `<h1>台語主日禮拜</h1><p>${formatted}</p>`;
  }
  else if (page.kind === 'content') content.className = 'slide-content template-content', content.innerHTML = `<h1>${title}</h1><div class="body">${safeHtml(page.body)}</div>`;
  else if (page.kind === 'scripture') content.className = 'slide-content template-content template-scripture', content.innerHTML = `<h1>${title}</h1><div class="body">${safeHtml(page.body)}</div>`;
  else if (page.kind === 'liturgical') {
    const alignment = page.align === 'center' ? ' is-centered' : ' is-left';
    content.className = `slide-content template-liturgical${alignment}${page.showTitle === false ? ' no-title' : ''}`;
    content.innerHTML = `${page.showTitle === false ? '' : `<h1>${title}</h1>`}<div class="body">${safeHtml(page.body)}</div>`;
  }
  else if (page.kind === 'praise-title') content.className = 'slide-content template-section', content.innerHTML = `<h1>讚美</h1><p>${title}</p><p>${kicker}</p>`;
  else if (page.kind === 'praise-lyrics') content.className = 'slide-content template-praise', content.innerHTML = `<div class="body">${safeHtml(page.body)}</div>`;
  else if (page.kind === 'score') content.className = 'slide-content template-score', content.innerHTML = `<h1>${title}</h1><p>${kicker}</p><div class="score-slot"></div>`;
  else {
    const subtitles = { '會前領唱':'請準備心今天的禮拜', '靜默一分鐘':'請將手機關機或靜音', '後奏':'請後奏結束後再起身或交談', '平安禮':'請兄弟姊妹互相行平安禮' };
    content.className = 'slide-content template-section';
    content.innerHTML = `<h1>${title}</h1><p>${kicker || subtitles[item.label] || ''}</p>`;
  }
  if (window.applyPageLayoutToPreview) window.applyPageLayoutToPreview(content, page);
};
document.getElementById('page-prev').onclick = () => window.navigateDeck ? window.navigateDeck(-1) : (previewPage = Math.max(0, previewPage - 1), preview());
document.getElementById('page-next').onclick = () => window.navigateDeck ? window.navigateDeck(1) : (previewPage += 1, preview());
document.addEventListener('keydown', event => {
  const target = event.target;
  if (target && (target.matches('input, textarea, select, button, [contenteditable="true"]'))) return;
  if (event.key === 'ArrowLeft' || event.key === 'ArrowUp') {
    event.preventDefault();
    if (window.navigateDeck) window.navigateDeck(-1); else { previewPage = Math.max(0, previewPage - 1); preview(); }
  } else if (event.key === 'ArrowRight' || event.key === 'ArrowDown') {
    event.preventDefault();
    if (window.navigateDeck) window.navigateDeck(1); else { previewPage += 1; preview(); }
  }
});
const originalRender = render;
render = function() { previewPage = 0; originalRender(); };
render();
