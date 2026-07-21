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
    const lines = (item.body || '').split('\n');
    const maxLen = lines.reduce((max, line) => Math.max(max, line.trim().length), 0);
    html = `
      ${field('詩歌名稱', 'title', item.title)}
      <label class="field">
        <span style="display:flex; justify-content:space-between; width:100%;">
          <span>歌詞（段落之間請空一行，以便 PPT 自動分頁）：</span>
          <span class="max-line-len-badge" style="font-weight:normal; opacity:0.7; font-size:13px; color:#bb86fc;">最長單行：${maxLen} 字</span>
        </span>
        <textarea data-key="body">${esc(item.body)}</textarea>
        <small>依你貼入的歌詞分頁，不自動生成歌詞。</small>
      </label>
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
      const key = event.target.dataset.key;
      item[key] = event.target.value;
      
      // Update max line length badge in real-time
      if (key === 'body' && item.type === 'praise') {
        const lines = event.target.value.split('\n');
        const maxLen = lines.reduce((max, l) => Math.max(max, l.trim().length), 0);
        const badge = $('.max-line-len-badge');
        if (badge) {
          badge.textContent = `最長單行：${maxLen} 字`;
        }
      }
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
  <div class="dialog-tabs">
    <button type="button" class="dialog-tab-btn active" id="tab-btn-ai">照片 AI 辨識</button>
    <button type="button" class="dialog-tab-btn" id="tab-btn-text">貼入純文字</button>
  </div>
  
  <!-- Tab 1: AI Image Upload -->
  <div class="dialog-tab-content active" id="tab-content-ai">
    <h3 style="margin-top:0;">上傳手寫草稿照片 (AI 圖文判定)</h3>
    <div class="dropzone" id="ai-image-dropzone">
      <span class="dropzone-icon">☁️</span>
      <span class="dropzone-text">將多張手寫相片拖曳至此，或點擊此處批量上傳</span>
      <span class="dropzone-subtext">(支援 PNG、JPG 或 WebP，可一次選取多張)</span>
      <input type="file" id="ai-image-file" accept="image/png,image/jpeg,image/webp" multiple style="display:none;">
    </div>
    
    <div class="batch-preview-list" id="ai-image-preview-list"></div>
    
    <div id="ai-parse-status" style="margin-top:15px; min-height:20px; font-size:13px; color:#bb86fc; white-space:pre-wrap;"></div>
    
    <div style="margin-top:15px; text-align:right;">
      <button class="button quiet" type="button" id="import-cancel-ai" style="margin-right:8px;">取消</button>
      <button class="button primary" type="button" id="ai-parse-btn">批量 AI 圖文判定</button>
    </div>
  </div>
  
  <!-- Tab 2: Raw Text Input -->
  <div class="dialog-tab-content" id="tab-content-text">
    <form method="dialog" id="import-dialog-form">
      <h3 style="margin-top:0;">智慧解析手寫/AI提取文字</h3>
      <p style="color:#aaa; font-size:13px; margin-bottom:10px;">請將禱告會手寫稿所提取的完整文字（如 1. 2. 3. 直到 13.）貼入下方框中：</p>
      <textarea id="import-textarea" style="width:100%; height:250px; background:#111; color:#fff; border:1px solid #444; border-radius:4px; padding:8px; font-family:monospace; font-size:13px; resize:vertical;"></textarea>
      <div style="margin-top:15px; text-align:right;">
        <button class="button quiet" type="button" id="import-cancel-text" style="margin-right:8px;">取消</button>
        <button class="button primary" type="submit" id="import-submit">進行解析與導入</button>
      </div>
    </form>
  </div>
`;
document.body.appendChild(importDialog);

// Tab Switching
const tabBtnAi = $('#tab-btn-ai');
const tabBtnText = $('#tab-btn-text');
const tabContentAi = $('#tab-content-ai');
const tabContentText = $('#tab-content-text');

tabBtnAi.onclick = () => {
  tabBtnAi.classList.add('active');
  tabBtnText.classList.remove('active');
  tabContentAi.classList.add('active');
  tabContentText.classList.remove('active');
};

tabBtnText.onclick = () => {
  tabBtnText.classList.add('active');
  tabBtnAi.classList.remove('active');
  tabContentText.classList.add('active');
  tabContentAi.classList.remove('active');
};

// Drag & Drop / Click Upload
const dropzone = $('#ai-image-dropzone');
const fileInput = $('#ai-image-file');
const previewList = $('#ai-image-preview-list');
const parseButton = $('#ai-parse-btn');
let selectedImageFiles = [];
let selectedPreviewUrls = [];

function clearSelectedImageFiles() {
  selectedPreviewUrls.forEach(url => URL.revokeObjectURL(url));
  selectedPreviewUrls = [];
  selectedImageFiles = [];
  fileInput.value = '';
  previewList.innerHTML = '';
  $('#ai-image-dropzone .dropzone-text').textContent = '將多張手寫相片拖曳至此，或點擊此處批量上傳';
}

function selectImageFiles(files) {
  const imageFiles = Array.from(files || []);
  if (!imageFiles.length) return;
  if (imageFiles.some(file => !['image/png', 'image/jpeg', 'image/webp'].includes(file.type))) {
    alert('僅支援 PNG、JPG 或 WebP 圖片。');
    return;
  }

  clearSelectedImageFiles();
  selectedImageFiles = imageFiles;
  selectedImageFiles.forEach(file => {
    const previewUrl = URL.createObjectURL(file);
    selectedPreviewUrls.push(previewUrl);
    const item = document.createElement('div');
    item.className = 'batch-preview-item';
    const image = document.createElement('img');
    image.src = previewUrl;
    image.alt = file.name;
    const label = document.createElement('small');
    label.textContent = file.name;
    label.title = file.name;
    item.append(image, label);
    previewList.appendChild(item);
  });
  $('#ai-image-dropzone .dropzone-text').textContent = `已選擇 ${selectedImageFiles.length} 張圖片`;
}

function readImageAsBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const dataUrl = String(reader.result || '');
      const commaIndex = dataUrl.indexOf(',');
      if (commaIndex === -1) {
        reject(new Error(`無法讀取 ${file.name}`));
        return;
      }
      resolve(dataUrl.substring(commaIndex + 1));
    };
    reader.onerror = () => reject(new Error(`圖片讀取失敗：${file.name}`));
    reader.readAsDataURL(file);
  });
}

dropzone.onclick = () => fileInput.click();

dropzone.ondragover = (e) => {
  e.preventDefault();
  dropzone.classList.add('dragover');
};

dropzone.ondragleave = () => {
  dropzone.classList.remove('dragover');
};

dropzone.ondrop = (e) => {
  e.preventDefault();
  dropzone.classList.remove('dragover');
  selectImageFiles(e.dataTransfer && e.dataTransfer.files);
};

fileInput.onchange = (e) => {
  selectImageFiles(e.target.files);
};

$('#import-cancel-ai').onclick = () => importDialog.close();
$('#import-cancel-text').onclick = () => importDialog.close();

// Trigger AI Parsing
parseButton.onclick = async () => {
  if (!selectedImageFiles.length) {
    alert('請先選擇或拖入要辨識的手寫稿照片！');
    return;
  }

  const statusDiv = $('#ai-parse-status');
  const recognizedTexts = [];
  parseButton.disabled = true;
  fileInput.disabled = true;
  statusDiv.style.color = 'var(--accent-color)';

  try {
    for (let index = 0; index < selectedImageFiles.length; index += 1) {
      const file = selectedImageFiles[index];
      statusDiv.textContent = `正在辨識第 ${index + 1}/${selectedImageFiles.length} 張：${file.name}`;
      const base64Data = await readImageAsBase64(file);
      const response = await window.churchAPI('cal_parsePrayerImage', { mimeType: file.type, base64Data });
      if (response && response.success && response.text) {
        recognizedTexts.push(response.text.trim());
      } else if (response && response.error) {
        throw new Error(`第 ${index + 1} 張辨識失敗：${response.error}`);
      } else {
        throw new Error(`第 ${index + 1} 張辨識失敗：${(response && response.message) || 'GAS 未回傳有效文字內容'}`);
      }
    }

    const combinedText = recognizedTexts.join('\n\n');
    $('#import-textarea').value = combinedText;
    importRecognizedTexts(recognizedTexts);
    statusDiv.textContent = `🎉 ${selectedImageFiles.length} 張照片辨識完成！已合併文字並解析經文。`;
    statusDiv.style.color = '#10b981';
    setTimeout(() => importDialog.close(), 1000);
  } catch (err) {
    statusDiv.textContent = `❌ 辨識失敗：${err.message}`;
    statusDiv.style.color = '#ef4444';
  } finally {
    parseButton.disabled = false;
    fileInput.disabled = false;
  }
};

importBtn.onclick = () => {
  // Clear file uploads and status on open
  clearSelectedImageFiles();
  $('#ai-parse-status').textContent = '';
  importDialog.showModal();
};

$('#import-dialog-form').onsubmit = (e) => {
  e.preventDefault();
  const text = $('#import-textarea').value;
  importRawText(text);
  importDialog.close();
};

function importRawText(text) {
  if (!text) return;
  importRecognizedTexts([text]);
}

function importRecognizedTexts(texts) {
  const sectionsToUpdate = window.PrayerSlideProduction.parseRecognizedSections(texts, model);

  // Update only the parsed sections (so single page scan increments nicely)
  const bibleSectionsToQuery = [];
  Object.entries(sectionsToUpdate).forEach(([sectionId, data]) => {
    model[sectionId].title = data.title;
    saveCurrentSectionContent(sectionId, data.lines);
    if (['scripture', 'repentance', 'nation', 'verse'].includes(sectionId)) {
      bibleSectionsToQuery.push(sectionId);
    }
  });

  status('正在透過 API 解析並查詢各段聖經經文…');
  Promise.all(bibleSectionsToQuery.map(sec => queryBibleForSection(sec))).then(() => {
    render();
    status('智慧文字導入與經文解析載入完成！');
  });
}

function saveCurrentSectionContent(sectionId, lines) {
  const item = model[sectionId];
  const bibleReferences = window.PrayerSlideProduction.extractBibleReferences(lines);
  if (sectionId === 'scripture' || sectionId === 'verse') {
    // OCR may include the handwritten verse itself; only send references to the Bible API.
    item.bibleQuery = bibleReferences.join('; ');
    item.body = '';
  } else if (sectionId === 'repentance' || sectionId === 'nation') {
    const referenceLineIndexes = lines.reduce((indexes, line, index) => {
      if (window.PrayerSlideProduction.extractBibleReferences([line]).length) indexes.push(index);
      return indexes;
    }, []);
    if (bibleReferences.length) {
      item.bibleQuery = bibleReferences[0];
      // The API owns Bible text. Keep only non-scripture prayer content in this list section.
      item.body = lines.filter((line, index) => !referenceLineIndexes.includes(index)).join('\n');
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
// Initial render is owned by ppt-format-preview.js after it defines preview().
