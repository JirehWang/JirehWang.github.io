const templateProfiles = window.PrayerTemplateProfiles;
const activeWorshipTemplateProfile = templateProfiles.getTemplateProfile('prayer');
const sections = activeWorshipTemplateProfile.sections;
const model = templateProfiles.createTemplateModel();
const draftKey = activeWorshipTemplateProfile.draftKey;

window.activeWorshipTemplateId = 'prayer';
window.activeWorshipTemplateProfile = activeWorshipTemplateProfile;
window.worshipDraftKey = draftKey;
window.worshipLayoutState = { groups: {}, pageAssignments: {}, outputScale: { text: 100, image: 100 } };
window.TaiwaneseWorshipSlideProduction = window.PrayerSlideProduction;
window.hymnOpacitySectionIds = [];

let active = activeWorshipTemplateProfile.activeSectionId || sections[0][0];
let backgroundColor = activeWorshipTemplateProfile.defaultBackgroundColor || '#111111';
let backgroundImage = '';

// LocalStorage Draft Loading
try {
  const draft = JSON.parse(localStorage.getItem(draftKey) || '{}');
  backgroundColor = window.PrayerSlideProduction.normalizeColor(draft.backgroundColor, backgroundColor);
  backgroundImage = window.PrayerSlideProduction.normalizeBackgroundImageDataUrl(draft.backgroundImage);
  Object.entries(draft.model || {}).forEach(([id, saved]) => {
    if (!model[id] || !saved || typeof saved !== 'object') return;
    ['title', 'kicker', 'body', 'bibleQuery', 'bibleRecords'].forEach(key => {
      if (saved[key] !== undefined) model[id][key] = saved[key];
    });
  });
} catch (error) {
  console.warn('草稿讀取失敗', error);
}

const $ = selector => document.querySelector(selector);
const esc = value => String(value == null ? '' : value).replace(/[&<>]/g, character => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[character]));

function field(label, key, value = '', type = 'text', hint = '') {
  return `<label class="field"><span>${label}</span>${type === 'textarea'
    ? `<textarea data-key="${key}">${esc(value)}</textarea>`
    : `<input data-key="${key}" value="${esc(value)}">`}${hint ? `<small>${hint}</small>` : ''}</label>`;
}

function flow() {
  $('#flow-list').innerHTML = sections.map(([id, label], index) => {
    const item = model[id];
    let pageCount = 1;
    if (window.PrayerSlideProduction) {
      pageCount = window.PrayerSlideProduction.generateSectionPages(id, item).length;
    }
    return `<button class="flow-item ${id === active ? 'active' : ''}" data-id="${id}"><span class="flow-number">${String(index + 1).padStart(2, '0')}</span>${label}<small style="float:right;opacity:0.6">${pageCount} 頁</small></button>`;
  }).join('');
  document.querySelectorAll('[data-id]').forEach(button => {
    button.onclick = () => { active = button.dataset.id; render(); };
  });
}

function editor() {
  const item = model[active];
  const form = $('#editor-form');
  let html = '';

  if (item.type === 'bible') {
    html = `
      ${field('標題', 'title', item.title)}
      ${field('經文代號（多筆請用分號分割，如：羅 8:26; 弗 6:18）', 'bibleQuery', item.bibleQuery, 'text', '自動向信望愛 API 查詢台語漢字版經文')}
      <div style="margin-top: 10px;">
        <button class="button quiet" type="button" id="query-bible-btn">手動發起查詢</button>
      </div>
      <div class="bible-preview-box" style="margin-top: 15px; padding: 10px; background: #222; border-radius: 4px; font-size: 14px; max-height: 200px; overflow-y: auto;">
        <strong>API 經文載入結果：</strong>
        <p style="margin-top:5px; color:#aaa; white-space:pre-wrap;">${item.bibleRecords && item.bibleRecords.length ? item.bibleRecords.map(r => `(${r.sec}) ${r.text}`).join('\n') : '尚未查詢或無經文資料'}</p>
      </div>
    `;
  } else if (item.type === 'list-bible') {
    html = `
      ${field('標題', 'title', item.title)}
      ${field('開頭經文代號（如：路 5:32）', 'bibleQuery', item.bibleQuery, 'text', '自動向信望愛 API 查詢台語漢字版經文')}
      <div style="margin-top: 10px; margin-bottom: 15px;">
        <button class="button quiet" type="button" id="query-bible-btn">手動發起查詢</button>
      </div>
      ${field('禱告項目清單 (每點一行，自動以 a., b., c. 拆分投影頁)', 'body', item.body, 'textarea')}
    `;
  } else if (item.type === 'list') {
    html = `
      ${field('標題', 'title', item.title)}
      ${field('禱告項目清單 (每點一行，自動以 a., b., c. 拆分投影頁)', 'body', item.body, 'textarea')}
    `;
  } else if (item.type === 'praise') {
    html = `
      ${field('詩歌名稱', 'title', item.title)}
      ${field('歌詞（以空白行分頁）', 'body', item.body, 'textarea')}
    `;
  } else {
    // content (quiet waiting, Lord's prayer, etc.)
    html = `
      ${field('標題', 'title', item.title)}
      ${field('內容（自動分頁）', 'body', item.body, 'textarea')}
    `;
  }

  form.innerHTML = html;
  
  // Handlers
  form.querySelectorAll('[data-key]').forEach(element => {
    element.oninput = event => {
      item[event.target.dataset.key] = event.target.value;
      preview();
    };
  });

  const queryBtn = $('#query-bible-btn');
  if (queryBtn) {
    queryBtn.onclick = () => queryBibleForSection(active);
  }
}

async function queryBibleForSection(sectionId) {
  const item = model[sectionId];
  if (!item || !item.bibleQuery) return;
  status(`正在查詢經文 "${item.bibleQuery}"…`);
  try {
    const results = await window.FhlBibleService.query(item.bibleQuery, 'tghg');
    const allRecords = results.flatMap(res => res.records.map(r => ({
      ...r,
      bookName: res.queryObj.bookName
    })));
    item.bibleRecords = allRecords;
    preview();
    flow(); // Update page counts in left column
    status('經文查詢成功！');
  } catch (error) {
    status(`經文查詢失敗：${error.message}`);
  }
}

function render() {
  flow();
  editor();
  preview();
  $('#section-kind').textContent = model[active].type === 'bible' ? '聖經經文 (台語漢字)' : '流程內容';
  $('#section-title').textContent = model[active].label;
}

function status(message) {
  const target = $('#save-state');
  const text = String(message || '');
  const state = /正在|讀取中|載入中|下載中|解析中|產生中|儲存中|驗證中|更新中/.test(text)
    ? 'busy'
    : /失敗|錯誤|無法|PERMISSION_DENIED|找不到/.test(text) ? 'error' : /已|成功/.test(text) ? 'success' : 'idle';
  target.textContent = text;
  target.dataset.state = state;
  target.setAttribute('aria-busy', String(state === 'busy'));
  document.body.classList.toggle('is-busy', state === 'busy');
}

// Background Controls
const backgroundInput = $('#background-color');
backgroundInput.value = backgroundColor;
backgroundInput.oninput = event => {
  backgroundColor = window.PrayerSlideProduction.normalizeColor(event.target.value, '#111111');
  preview();
  status(`已套用純色背景：${backgroundColor}`);
};

const backgroundImageUpload = $('#background-image-upload');
backgroundImageUpload.onchange = event => {
  const file = event.target.files && event.target.files[0];
  if (!file) return;
  if (!window.PrayerSlideProduction.isSupportedBackgroundImage(file)) {
    status('背景圖僅支援 PNG、JPG 或 WebP');
    event.target.value = '';
    return;
  }
  const reader = new FileReader();
  status('正在讀取背景圖片…');
  reader.onload = () => {
    backgroundImage = window.PrayerSlideProduction.normalizeBackgroundImageDataUrl(reader.result);
    preview();
    status(`已套用背景圖：${file.name}`);
    event.target.value = '';
  };
  reader.onerror = () => status('背景圖讀取失敗，請重新選擇');
  reader.readAsDataURL(file);
};

// Drafts
$('#save-draft').onclick = () => {
  localStorage.setItem(draftKey, JSON.stringify({ model, backgroundColor, backgroundImage }));
  status('已儲存至此瀏覽器');
};

// Export
$('#export-ppt').onclick = async () => {
  try {
    status('正在匯出 PPTX 簡報檔…');
    await window.PrayerPptExport.exportPrayerPPTX({
      model,
      backgroundColor,
      backgroundImage,
      getDeckEntries: () => window.PrayerSlideProduction.buildDeckEntries(sections, model)
    });
    status('PPTX 簡報檔已成功下載！');
  } catch (error) {
    status(`簡報匯出失敗：${error.message}`);
    console.error(error);
  }
};

// Date
const localIsoDate = () => new Intl.DateTimeFormat('sv-SE', {
  timeZone: 'Asia/Taipei', year: 'numeric', month: '2-digit', day: '2-digit'
}).format(new Date());
$('#service-date').value = localIsoDate();

// Text Import & Parsing Dialog
const importBtn = document.createElement('button');
importBtn.className = 'button quiet';
importBtn.id = 'import-raw-text-btn';
importBtn.textContent = '智慧導入文字';
$('.topbar-actions').insertBefore(importBtn, $('#save-draft'));

const importDialog = document.createElement('dialog');
importDialog.id = 'import-text-dialog';
importDialog.style.padding = '20px';
importDialog.style.background = '#222';
importDialog.style.color = '#fff';
importDialog.style.border = '1px solid #444';
importDialog.style.borderRadius = '8px';
importDialog.style.width = '600px';
importDialog.style.maxWidth = '90%';
importDialog.innerHTML = `
  <form method="dialog" id="import-dialog-form">
    <h3 style="margin-top:0;">智慧解析手寫/AI提取文字</h3>
    <p style="color:#aaa; font-size:13px; margin-bottom:10px;">請將禱告會手寫稿所提取的完整文字（如 1. 2. 3. 直到 13.）貼入下方框中：</p>
    <textarea id="import-textarea" style="width:100%; height:300px; background:#111; color:#fff; border:1px solid #444; border-radius:4px; padding:8px; font-family:monospace; font-size:13px; resize:vertical;"></textarea>
    <div style="margin-top:15px; text-align:right;">
      <button class="button quiet" type="button" id="import-cancel" style="margin-right:8px;">取消</button>
      <button class="button primary" type="submit" id="import-submit">進行解析與導入</button>
    </div>
  </form>
`;
document.body.appendChild(importDialog);

importBtn.onclick = () => {
  importDialog.showModal();
};

$('#import-cancel').onclick = () => {
  importDialog.close();
};

$('#import-dialog-form').onsubmit = (e) => {
  e.preventDefault();
  const text = $('#import-textarea').value;
  importRawText(text);
  importDialog.close();
};

function importRawText(text) {
  if (!text) return;
  
  const lines = text.split('\n');
  let currentSection = null;
  let currentLines = [];

  // Reset model
  Object.keys(model).forEach(key => {
    model[key].body = '';
    if (model[key].bibleQuery !== undefined) {
      model[key].bibleQuery = '';
      model[key].bibleRecords = [];
    }
  });

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

  lines.forEach(line => {
    const trimmed = line.trim();
    if (!trimmed) return;

    // Matches main section headers: "1. 請安靜心...", "2. 詩歌...", "4. 獻上感謝...", "11. pray 金句..."
    const headerMatch = trimmed.match(/^(\d+)\.\s*(.*)/);
    if (headerMatch) {
      const num = parseInt(headerMatch[1], 10);
      const sectionKey = numberToSectionMap[num];
      if (sectionKey) {
        if (currentSection) {
          saveCurrentSectionContent(currentSection, currentLines);
        }
        currentSection = sectionKey;
        currentLines = [];
        const titleText = headerMatch[2].trim();
        model[sectionKey].title = titleText || model[sectionKey].label;
        return;
      }
    }

    if (currentSection) {
      currentLines.push(trimmed);
    }
  });

  if (currentSection) {
    saveCurrentSectionContent(currentSection, currentLines);
  }

  // Trigger Bible queries for all Bible sections
  const bibleSections = ['scripture', 'repentance', 'nation', 'verse'];
  status('正在透過 API 解析並查詢各段聖經經文…');
  Promise.all(bibleSections.map(sec => queryBibleForSection(sec))).then(() => {
    render();
    status('智慧文字導入與經文解析載入完成！');
  });
}

function saveCurrentSectionContent(sectionId, lines) {
  const item = model[sectionId];
  if (sectionId === 'scripture' || sectionId === 'verse') {
    const queries = lines.map(line => {
      return line.replace(/\(台語\)|\(華語\)|（台語）|（華語）/g, '').replace(/[\"\']/g, '').trim();
    }).filter(Boolean);
    item.bibleQuery = queries.join('; ');
  } else if (sectionId === 'repentance' || sectionId === 'nation') {
    const firstLine = lines[0] || '';
    const hasScripture = /^[a-zA-Z\u4e00-\u9fa5]+\s*\d+/.test(firstLine);
    if (hasScripture) {
      item.bibleQuery = firstLine.replace(/\(台語\)|\(華語\)|（台語）|（華語）/g, '').replace(/[\"\']/g, '').trim();
      item.body = lines.slice(1).join('\n');
    } else {
      item.bibleQuery = '';
      item.body = lines.join('\n');
    }
  } else {
    item.body = lines.join('\n');
  }
}

// Initial setup to bind coordinate-saving features (mocked for simplicity, or synchronized via window structure)
window.getDeckEntries = () => window.PrayerSlideProduction.buildDeckEntries(sections, model);
window.navigateDeck = null; // Defined by layout-groups.js if loaded

render();
