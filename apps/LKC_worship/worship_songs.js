/* worship_songs.js - 敬拜曲目管理頁面邏輯 */

let songsData = [];        // 目前載入的資料
let dirtyRows = new Set(); // 已修改但未存的列索引

// --- API 橋接 ---
async function callAPI(action, payload) {
  if (typeof window.churchAPI !== 'function') {
    throw new Error("中央安全設定檔 (config.js) 尚未載入");
  }
  return await window.churchAPI(action, payload);
}

// --- 讀取：季度 ---
async function loadByQuarter() {
  const val = document.getElementById('quarterSelect').value;
  const [year, quarter] = val.split('-');
  await loadSongs({ year, quarter });
}

// --- 讀取：日期區間 ---
async function loadByDateRange() {
  const start = document.getElementById('startDate').value;
  const end = document.getElementById('endDate').value;
  if (!start || !end) return userNotification.warning("請設定起訖日期");
  await loadSongs({ startDate: start, endDate: end });
}

// --- 核心讀取函式 ---
async function loadSongs(payload) {
  showPlaceholder('<div class="spinner-border spinner-border-sm text-primary me-2"></div> 讀取中...');
  dirtyRows.clear();
  updateSaveBtn();

  try {
    const result = await callAPI('getSongs', payload);
    if (result.status === 'success' && result.data.length > 0) {
      songsData = result.data;
      renderCards();
      document.getElementById('syncTime').innerText = new Date().toLocaleTimeString();
    } else {
      showPlaceholder('📋 此區間無聚會資料');
    }
  } catch (err) {
    showPlaceholder(`❌ 讀取失敗：${err.message}`);
  }
}

// --- 渲染卡片列表 ---
function renderCards() {
  const container = document.getElementById('songsCardsContainer');
  container.innerHTML = '';

  songsData.forEach((row, idx) => {
    const card = document.createElement('div');
    card.className = 'song-card';
    card.id = `card-${idx}`;

    const songs = row['敬拜曲目'] || '';
    const songItems = songs
      ? songs.split('、').map(s => `<div class="song-item">🎵 ${s.trim()}</div>`).join('')
      : '';

    card.innerHTML = `
      <div>
        <div class="card-top">
          <div class="card-date-info">
            <span class="card-date">${row['日期'] || ''}</span>
            <span class="card-meeting-name fw-bold">${row['聚會名稱'] || ''}</span>
          </div>
          <div class="badge-container">
            <span class="badge-g">${row['聚會類別'] || '聚會'}</span>
            <span class="badge-b ${row['主領'] ? '' : 'empty'}">主領：${row['主領'] || '待定'}</span>
          </div>
        </div>
        
        <div class="card-middle-content" id="songs-cell-${idx}">
          <div class="songs-display ${songs ? '' : 'empty'}" id="songs-display-${idx}">
            ${songs ? songItems : '尚未填入曲目'}
          </div>
          <textarea class="songs-input" id="songs-input-${idx}" 
                    style="display:none;" 
                    placeholder="輸入曲目，以任意符號分隔（如逗號、頓號、斜線）&#10;例：Amazing Grace,How Great Thou Art,主禱文"
          >${songs ? songs.split('、').join('\n') : ''}</textarea>
          <div class="hint-text" id="songs-hint-${idx}" style="display:none;">
            💡 每首歌一行，或以任意符號分隔，儲存後自動整理
          </div>
        </div>
      </div>
      
      <div class="card-bottom-actions">
        <div id="btn-group-${idx}">
          <button class="btn-edit-row" onclick="editRow(${idx})">✏️ 編輯曲目</button>
        </div>
      </div>
    `;
    container.appendChild(card);
  });

  document.getElementById('tableWrapper').style.display = 'block';
  document.getElementById('placeholder').style.display = 'none';
}

// --- 進入編輯模式 ---
function editRow(idx) {
  // 顯示 textarea，隱藏 display
  document.getElementById(`songs-display-${idx}`).style.display = 'none';
  document.getElementById(`songs-input-${idx}`).style.display = 'block';
  document.getElementById(`songs-hint-${idx}`).style.display = 'block';
  document.getElementById(`songs-cell-${idx}`).classList.add('edit-mode');
  document.getElementById(`songs-input-${idx}`).focus();

  // 換按鈕
  document.getElementById(`btn-group-${idx}`).innerHTML = `
    <button class="btn-cancel-row me-1" onclick="cancelRow(${idx})">取消</button>
    <button class="btn-save-row" onclick="saveRow(${idx})">💾 暫存修改</button>
  `;

  dirtyRows.add(idx);
  updateSaveBtn();
}

// --- 確認單列 ---
function saveRow(idx) {
  const raw = document.getElementById(`songs-input-${idx}`).value.trim();
  // 將任意分隔符號（含換行）統一轉成「、」
  const normalized = raw
    ? raw.split(/[\n\r,，、\/\\|；;]+/).map(s => s.trim()).filter(s => s).join('、')
    : '';

  songsData[idx]['敬拜曲目'] = normalized;

  // 更新 display
  const display = document.getElementById(`songs-display-${idx}`);
  const songItems = normalized
    ? normalized.split('、').map(s => `<div class="song-item">🎵 ${s}</div>`).join('')
    : '';
  display.innerHTML = normalized ? songItems : '尚未填入曲目';
  display.className = `songs-display ${normalized ? '' : 'empty'}`;
  display.style.display = 'block';

  document.getElementById(`songs-input-${idx}`).style.display = 'none';
  document.getElementById(`songs-hint-${idx}`).style.display = 'none';
  document.getElementById(`songs-cell-${idx}`).classList.remove('edit-mode');

  // 換回編輯按鈕
  document.getElementById(`btn-group-${idx}`).innerHTML = `
    <button class="btn-edit-row" onclick="editRow(${idx})">✏️ 編輯曲目</button>
  `;
}

// --- 取消編輯 ---
function cancelRow(idx) {
  const original = songsData[idx]['敬拜曲目'] || '';
  document.getElementById(`songs-input-${idx}`).value = original
    ? original.split('、').join('\n')
    : '';
  document.getElementById(`songs-input-${idx}`).style.display = 'none';
  document.getElementById(`songs-hint-${idx}`).style.display = 'none';
  document.getElementById(`songs-display-${idx}`).style.display = 'block';
  document.getElementById(`songs-cell-${idx}`).classList.remove('edit-mode');

  document.getElementById(`btn-group-${idx}`).innerHTML = `
    <button class="btn-edit-row" onclick="editRow(${idx})">✏️ 編輯曲目</button>
  `;

  dirtyRows.delete(idx);
  updateSaveBtn();
}

// --- 儲存全部 ---
async function saveAllSongs() {
  // 先確認所有還在編輯中的列
  dirtyRows.forEach(idx => {
    const input = document.getElementById(`songs-input-${idx}`);
    if (input && input.style.display !== 'none') saveRow(idx);
  });

  const btn = document.getElementById('saveAllBtn');
  btn.disabled = true;
  btn.innerText = '儲存中...';

  try {
    const result = await callAPI('saveSongs', { songsData: songsData });
    if (result.status === 'success') {
      userNotification.success('🎉 曲目資料已成功儲存！');
      dirtyRows.clear();
      updateSaveBtn();
      document.getElementById('syncTime').innerText = new Date().toLocaleTimeString();
    } else {
      userNotification.warning('⚠️ 儲存失敗：' + result.message);
    }
  } catch (err) {
    userNotification.error('❌ 儲存失敗：' + err.message);
  }

  btn.disabled = false;
  btn.innerText = '💾 儲存所有曲目';
}

// --- 更新固定儲存按鈕顯示 ---
function updateSaveBtn() {
  document.getElementById('saveAllBtn').style.display =
    songsData.length > 0 ? 'block' : 'none';
}

// --- 工具：顯示佔位訊息 ---
function showPlaceholder(msg) {
  document.getElementById('tableWrapper').style.display = 'none';
  const ph = document.getElementById('placeholder');
  ph.style.display = 'block';
  ph.innerHTML = `<div class="p-4">${msg}</div>`;
}

// --- 頁面初始化 ---
window.onload = () => {
  // 根據當前日期設定預設季度
  const today = new Date();
  const currentYear = today.getFullYear();
  const currentMonth = today.getMonth(); // 0-11
  const currentQuarter = 'Q' + (Math.floor(currentMonth / 3) + 1);
  const targetVal = `${currentYear}-${currentQuarter}`;

  const quarterSelect = document.getElementById('quarterSelect');
  if (quarterSelect && Array.from(quarterSelect.options).some(opt => opt.value === targetVal)) {
    quarterSelect.value = targetVal;
  }
};
