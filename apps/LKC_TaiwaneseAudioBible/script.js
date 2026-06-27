// 台語有聲聖經 - script.js

// 1. 聖經 66 卷書對照表與狀態變數
const BIBLE_BOOKS = FhlBibleService.BIBLE_BOOKS;
let activePlaylist = []; // 目前播放清單項目：[{ eng, chap, bookName, bid, label }]
let currentPlayIndex = -1;
let currentBibleVersion = 'tghg';
let currentAudioVersion = '1';
let currentSpeed = 1.0;

// 2. UI 元素
const bibleVerSelect = document.getElementById('bible-ver-select');
const audioVerSelect = document.getElementById('audio-ver-select');
const queryInput = document.getElementById('query-input');
const btnQuery = document.getElementById('btn-query');
const playlistCard = document.getElementById('playlist-card');
const playlistContainer = document.getElementById('playlist-container');

const currentChapterTitle = document.getElementById('current-chapter-title');
const audioSourceName = document.getElementById('audio-source-name');
const bibleAudio = document.getElementById('bible-audio');

const speedRange = document.getElementById('speed-range');
const speedVal = document.getElementById('speed-val');
const autoNextCheckbox = document.getElementById('auto-next');

const scriptureTitleDisplay = document.getElementById('scripture-title-display');
const scriptureInfoDisplay = document.getElementById('scripture-info-display');
const scriptureBody = document.getElementById('scripture-body');

// 3. 初始化事件繫結
document.addEventListener('DOMContentLoaded', () => {
    // 綁定查詢按鈕與輸入框 Enter
    btnQuery.addEventListener('click', performQuery);
    queryInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            performQuery();
        }
    });

    // 綁定譯本變更
    bibleVerSelect.addEventListener('change', () => {
        currentBibleVersion = bibleVerSelect.value;
        // 如果當前已有清單，重新查詢以刷新經文譯本
        if (queryInput.value.trim()) {
            performQuery();
        }
    });

    audioVerSelect.addEventListener('change', () => {
        currentAudioVersion = audioVerSelect.value;
        // 變更語音譯本時，若當前有正在播放的章節，則重新載入音訊
        if (currentPlayIndex >= 0 && activePlaylist[currentPlayIndex]) {
            loadAudioForChapter(activePlaylist[currentPlayIndex]);
        }
    });

    // 語速變更
    speedRange.addEventListener('input', () => {
        currentSpeed = parseFloat(speedRange.value);
        speedVal.textContent = currentSpeed.toFixed(1);
        bibleAudio.playbackRate = currentSpeed;
    });

    // 音訊播放結束自動下一首
    bibleAudio.addEventListener('ended', () => {
        if (autoNextCheckbox.checked) {
            playNextChapter();
        }
    });

    // 錯誤處理
    bibleAudio.addEventListener('error', (e) => {
        console.error("Audio playback error:", e);
        showToast("無法載入此章節的音訊檔案，可能該版本查無此章音檔。", "error");
    });

    // 解析 URL 參數進行自動查詢
    parseUrlParams();
});

// 4. 解析 URL 參數
function parseUrlParams() {
    const urlParams = new URLSearchParams(window.location.search);
    const queryParam = urlParams.get('query');
    const bibleVerParam = urlParams.get('bible_ver');
    const audioVerParam = urlParams.get('audio_ver');
    const speedParam = urlParams.get('speed');

    if (bibleVerParam) {
        bibleVerSelect.value = bibleVerParam;
        currentBibleVersion = bibleVerParam;
    }
    if (audioVerParam) {
        audioVerSelect.value = audioVerParam;
        currentAudioVersion = audioVerParam;
    }
    if (speedParam) {
        const speed = parseFloat(speedParam);
        if (speed >= 0.5 && speed <= 2.0) {
            speedRange.value = speed;
            currentSpeed = speed;
            speedVal.textContent = speed.toFixed(1);
        }
    }
    if (queryParam) {
        queryInput.value = queryParam;
        performQuery();
    }
}

// 5. 取得書卷的 bid (1-indexed, 1-66)
function getBidByBookEng(eng) {
    const idx = BIBLE_BOOKS.findIndex(b => b.eng.toLowerCase() === eng.toLowerCase());
    return idx >= 0 ? idx + 1 : -1;
}

// 6. 執行經文與播放清單查詢
async function performQuery() {
    const qStr = queryInput.value.trim();
    if (!qStr) {
        showToast("請輸入經文段落範圍", "error");
        return;
    }

    showLoadingState();

    try {
        // 使用共享服務進行經文查詢
        const results = await FhlBibleService.query(qStr, currentBibleVersion);
        
        if (!results || results.length === 0) {
            throw new Error("無法解析經文格式");
        }

        // 渲染右側經文與大標
        renderScripture(results);

        // 建立播放清單 (Playlist)
        buildPlaylist(results);

        showToast("經文查詢成功！", "success");
    } catch (err) {
        console.error(err);
        showErrorState(err.message || "經文查詢失敗，請檢查輸入格式或網路連線。");
    }
}

// 7. 顯示 Loading 狀態
function showLoadingState() {
    scriptureTitleDisplay.textContent = "讀經板";
    scriptureInfoDisplay.textContent = "載入中...";
    scriptureBody.innerHTML = `
        <div class="placeholder-msg">
            <div class="logo-icon" style="animation: spin 1s linear infinite; font-size: 28px; width: 44px; height: 44px;">🔄</div>
            <p>正在從信望愛 API 獲取經文與音訊資料...</p>
        </div>
    `;
    // 旋轉動畫
    if (!document.getElementById('spin-style')) {
        const style = document.createElement('style');
        style.id = 'spin-style';
        style.innerHTML = `@keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }`;
        document.head.appendChild(style);
    }
}

// 8. 顯示 Error 狀態
function showErrorState(msg) {
    scriptureTitleDisplay.textContent = "讀經板";
    scriptureInfoDisplay.textContent = "錯誤";
    scriptureBody.innerHTML = `
        <div class="placeholder-msg" style="color: var(--danger);">
            <span class="placeholder-icon">⚠️</span>
            <p>${msg}</p>
        </div>
    `;
    playlistCard.style.display = 'none';
}

// 9. 渲染經文列表到讀經板
function renderScripture(results) {
    scriptureBody.innerHTML = '';
    
    // 取得第一個查詢結果的書名與範圍作為 Badge 顯示
    const firstQObj = results[0].queryObj;
    const lastQObj = results[results.length - 1].queryObj;
    
    let rangeLabel = `${firstQObj.bookName} ${firstQObj.chap}`;
    if (firstQObj.sec) rangeLabel += `:${firstQObj.sec}`;
    if (results.length > 1 || (firstQObj.chap !== lastQObj.chap)) {
        rangeLabel += ` ~ ${lastQObj.bookName} ${lastQObj.chap}`;
        if (lastQObj.sec) rangeLabel += `:${lastQObj.sec}`;
    }

    scriptureTitleDisplay.textContent = "台語聖經經文";
    scriptureInfoDisplay.textContent = rangeLabel;

    // 循序輸出每一節
    results.forEach(res => {
        const qObj = res.queryObj;
        const records = res.records || [];
        
        // 增加一個章節子標題
        const sectionHeader = document.createElement('div');
        sectionHeader.className = 'scripture-section-header';
        sectionHeader.style.margin = '24px 0 12px 0';
        sectionHeader.style.paddingBottom = '8px';
        sectionHeader.style.borderBottom = '1px dashed rgba(255,255,255,0.06)';
        sectionHeader.style.fontSize = '1.1rem';
        sectionHeader.style.color = 'var(--text-secondary)';
        sectionHeader.style.fontWeight = '600';
        sectionHeader.textContent = `【${qObj.bookName} 第 ${qObj.chap} 章】`;
        scriptureBody.appendChild(sectionHeader);

        records.forEach(rec => {
            const verseDiv = document.createElement('div');
            verseDiv.className = 'verse-p';
            verseDiv.dataset.verseId = `${qObj.eng}_${rec.chap}_${rec.sec}`;
            
            const numSpan = document.createElement('span');
            numSpan.className = 'verse-num';
            numSpan.textContent = rec.sec;
            
            const textSpan = document.createElement('span');
            textSpan.className = 'verse-text';
            textSpan.innerHTML = rec.text;

            verseDiv.appendChild(numSpan);
            verseDiv.appendChild(textSpan);
            scriptureBody.appendChild(verseDiv);
        });
    });
}

// 10. 建立播放清單
function buildPlaylist(results) {
    activePlaylist = [];
    playlistContainer.innerHTML = '';

    // 整理不重複的 { eng, chap, bookName, bid }
    const seen = new Set();
    
    results.forEach(res => {
        const qObj = res.queryObj;
        const bid = getBidByBookEng(qObj.eng);
        const key = `${qObj.eng}_${qObj.chap}`;
        
        if (!seen.has(key) && bid !== -1) {
            seen.add(key);
            activePlaylist.push({
                eng: qObj.eng,
                chap: qObj.chap,
                bookName: qObj.bookName,
                bid: bid,
                label: `${qObj.bookName} 第 ${qObj.chap} 章`
            });
        }
    });

    if (activePlaylist.length === 0) {
        playlistCard.style.display = 'none';
        return;
    }

    playlistCard.style.display = 'flex';

    // 渲染清單項目
    activePlaylist.forEach((item, index) => {
        const playItem = document.createElement('div');
        playItem.className = 'playlist-item';
        playItem.dataset.index = index;
        
        const titleSpan = document.createElement('span');
        titleSpan.textContent = item.label;
        
        const iconSpan = document.createElement('span');
        iconSpan.className = 'playlist-item-play-icon';
        iconSpan.textContent = "▶";

        playItem.appendChild(titleSpan);
        playItem.appendChild(iconSpan);

        playItem.addEventListener('click', () => {
            playChapter(index);
        });

        playlistContainer.appendChild(playItem);
    });

    // 預設播放第一個章節
    playChapter(0);
}

// 11. 播放指定索引的章節
function playChapter(index) {
    if (index < 0 || index >= activePlaylist.length) return;
    
    currentPlayIndex = index;
    const chapterItem = activePlaylist[index];

    // 更新清單項目的 active 狀態
    const items = playlistContainer.querySelectorAll('.playlist-item');
    items.forEach(el => el.classList.remove('active'));
    
    const activeEl = playlistContainer.querySelector(`.playlist-item[data-index="${index}"]`);
    if (activeEl) {
        activeEl.classList.add('active');
        activeEl.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }

    // 載入音訊
    loadAudioForChapter(chapterItem);
}

// 12. 請求 API 並載入音訊
async function loadAudioForChapter(item) {
    currentChapterTitle.textContent = item.label;
    audioSourceName.textContent = "正在獲取語音...";

    try {
        const apiUrl = `https://bible.fhl.net/json/au.php?version=${currentAudioVersion}&bid=${item.bid}&chap=${item.chap}`;
        
        const response = await fetch(apiUrl);
        if (!response.ok) {
            throw new Error(`HTTP 錯誤: ${response.status}`);
        }
        
        const data = await response.json();
        if (data.status !== 'success' || !data.mp3) {
            throw new Error("查無此章節的音訊檔案");
        }

        // 更新音訊來源
        bibleAudio.src = data.mp3;
        audioSourceName.textContent = `版本: ${data.name || '台語有聲聖經'}`;
        
        // 載入並播放
        bibleAudio.load();
        
        // 設定播放語速
        bibleAudio.playbackRate = currentSpeed;
        
        // 延遲播放以避開瀏覽器限制
        setTimeout(() => {
            bibleAudio.play().catch(err => {
                console.log("Auto-play was blocked or interrupted:", err);
                audioSourceName.textContent += " (點擊播放按鈕開始)";
            });
        }, 150);

        // 高亮讀經板中對應當前播放章節的經文
        highlightActiveChapterText(item.eng, item.chap);

    } catch (err) {
        console.error(err);
        audioSourceName.textContent = "無法載入此章音訊";
        showToast(`語音載入失敗: ${err.message}`, "error");
    }
}

// 13. 高亮讀經板對應章節經文
function highlightActiveChapterText(eng, chap) {
    const allVerses = scriptureBody.querySelectorAll('.verse-p');
    allVerses.forEach(el => {
        const id = el.dataset.verseId;
        if (id && id.startsWith(`${eng}_${chap}_`)) {
            el.style.borderLeft = "3px solid var(--primary)";
            el.style.backgroundColor = "rgba(16, 185, 129, 0.03)";
        } else {
            el.style.borderLeft = "none";
            el.style.backgroundColor = "transparent";
        }
    });
}

// 14. 播放下一章
function playNextChapter() {
    if (currentPlayIndex >= 0 && currentPlayIndex < activePlaylist.length - 1) {
        playChapter(currentPlayIndex + 1);
    } else {
        showToast("已播放至清單最後一章", "success");
    }
}

// 15. Toast 訊息提示
function showToast(msg, type = "success") {
    const toast = document.getElementById('toast-message');
    toast.className = `toast ${type}`;
    toast.textContent = msg;
    toast.classList.add('show');

    setTimeout(() => {
        toast.classList.remove('show');
    }, 3000);
}
