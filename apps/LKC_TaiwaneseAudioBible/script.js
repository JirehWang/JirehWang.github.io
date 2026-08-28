// 台語有聲聖經 - script.js

// 1. 聖經 66 卷書對照表與狀態變數
const BIBLE_BOOKS = FhlBibleService.BIBLE_BOOKS;
const OLD_TESTAMENT_BOOKS = BIBLE_BOOKS.slice(0, 39); // 舊約 39 卷 (創世記 ~ 瑪拉基書)
const NEW_TESTAMENT_BOOKS = BIBLE_BOOKS.slice(39);    // 新約 27 卷 (馬太福音 ~ 啟示錄)

let currentBookIndex = 18; // 預設 詩篇 (索引 18，第 19 卷)
let currentChapter = 23;   // 預設 第 23 章
let currentCatalogTestament = 'ot'; // 視覺目錄目前頁籤：'ot' (舊約) 或 'nt' (新約)

let activePlaylist = []; // 目前播放清單項目：[{ eng, chap, bookName, bid, label }]
let currentPlayIndex = -1;
let currentBibleVersion = 'tghg';
let currentAudioVersion = '1';
let currentSpeed = 1.0;

// 2. UI 元素宣告
const bibleVerSelect = document.getElementById('bible-ver-select');
const audioVerSelect = document.getElementById('audio-ver-select');
const bookSelect = document.getElementById('book-select');
const chapterSelect = document.getElementById('chapter-select');
const btnGoChapter = document.getElementById('btn-go-chapter');

const btnToggleCatalog = document.getElementById('btn-toggle-catalog');
const catalogToggleText = document.getElementById('catalog-toggle-text');
const catalogPanel = document.getElementById('catalog-panel');
const catalogBooksContainer = document.getElementById('catalog-books-container');
const catalogChaptersSection = document.getElementById('catalog-chapters-section');
const catalogSelectedBookName = document.getElementById('catalog-selected-book-name');
const catalogChaptersGrid = document.getElementById('catalog-chapters-grid');
const testamentTabBtns = document.querySelectorAll('.tab-btn');

const manualQueryToggle = document.getElementById('manual-query-toggle');
const manualQueryBody = document.getElementById('manual-query-body');
const queryInput = document.getElementById('query-input');
const btnQuery = document.getElementById('btn-query');

const playlistCard = document.getElementById('playlist-card');
const playlistContainer = document.getElementById('playlist-container');

const currentChapterTitle = document.getElementById('current-chapter-title');
const audioSourceName = document.getElementById('audio-source-name');
const bibleAudio = document.getElementById('bible-audio');

const btnPrevChap = document.getElementById('btn-prev-chap');
const btnNextChap = document.getElementById('btn-next-chap');
const btnHeaderPrev = document.getElementById('btn-header-prev');
const btnHeaderNext = document.getElementById('btn-header-next');
const btnFooterPrev = document.getElementById('btn-footer-prev');
const btnFooterNext = document.getElementById('btn-footer-next');
const footerPrevLabel = document.getElementById('footer-prev-label');
const footerNextLabel = document.getElementById('footer-next-label');
const footerCurrentLabel = document.getElementById('footer-current-label');

const speedRange = document.getElementById('speed-range');
const speedVal = document.getElementById('speed-val');
const autoNextCheckbox = document.getElementById('auto-next');

const scriptureTitleDisplay = document.getElementById('scripture-title-display');
const scriptureInfoDisplay = document.getElementById('scripture-info-display');
const scriptureBody = document.getElementById('scripture-body');

// 3. 初始化事件繫結
document.addEventListener('DOMContentLoaded', () => {
    // 3.1 初始化書卷下拉選單 (分舊約/新約區塊) 與章數選單
    initBookAndChapterSelectors();

    // 3.2 初始化視覺目錄
    initVisualCatalog();

    // 3.3 綁定選單事件
    bookSelect.addEventListener('change', () => {
        const eng = bookSelect.value;
        const bIdx = BIBLE_BOOKS.findIndex(b => b.eng === eng);
        if (bIdx !== -1) {
            currentBookIndex = bIdx;
            updateChapterDropdown(currentBookIndex, 1);
        }
    });

    btnGoChapter.addEventListener('click', () => {
        const eng = bookSelect.value;
        const bIdx = BIBLE_BOOKS.findIndex(b => b.eng === eng);
        const chap = parseInt(chapterSelect.value, 10) || 1;
        if (bIdx !== -1) {
            loadBookChapter(bIdx, chap, true);
        }
    });

    chapterSelect.addEventListener('change', () => {
        const eng = bookSelect.value;
        const bIdx = BIBLE_BOOKS.findIndex(b => b.eng === eng);
        const chap = parseInt(chapterSelect.value, 10) || 1;
        if (bIdx !== -1) {
            loadBookChapter(bIdx, chap, true);
        }
    });

    // 3.4 手動查詢展開/收合切換
    manualQueryToggle.addEventListener('click', () => {
        const isHidden = manualQueryBody.style.display === 'none';
        manualQueryBody.style.display = isHidden ? 'flex' : 'none';
        const icon = manualQueryToggle.querySelector('.toggle-icon');
        if (icon) icon.textContent = isHidden ? '－' : '＋';
    });

    // 3.5 手動查詢按鈕與 Enter
    btnQuery.addEventListener('click', performQuery);
    queryInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            performQuery();
        }
    });

    // 3.6 上一章 / 下一章按鈕事件
    btnPrevChap.addEventListener('click', playPrevChapter);
    btnNextChap.addEventListener('click', playNextChapter);
    btnHeaderPrev.addEventListener('click', playPrevChapter);
    btnHeaderNext.addEventListener('click', playNextChapter);
    btnFooterPrev.addEventListener('click', playPrevChapter);
    btnFooterNext.addEventListener('click', playNextChapter);

    // 3.7 聖經與語音譯本變更
    bibleVerSelect.addEventListener('change', () => {
        currentBibleVersion = bibleVerSelect.value;
        // 重新載入當前章節
        loadBookChapter(currentBookIndex, currentChapter, false);
    });

    audioVerSelect.addEventListener('change', () => {
        currentAudioVersion = audioVerSelect.value;
        if (currentPlayIndex >= 0 && activePlaylist[currentPlayIndex]) {
            loadAudioForChapter(activePlaylist[currentPlayIndex], true);
        }
    });

    // 3.8 語速變更
    speedRange.addEventListener('input', () => {
        currentSpeed = parseFloat(speedRange.value);
        speedVal.textContent = currentSpeed.toFixed(1);
        bibleAudio.playbackRate = currentSpeed;
    });

    // 3.9 音訊播放結束自動下一首
    bibleAudio.addEventListener('ended', () => {
        if (autoNextCheckbox.checked) {
            playNextChapter();
        }
    });

    // 3.10 錯誤處理
    bibleAudio.addEventListener('error', (e) => {
        console.error("Audio playback error:", e);
        showToast("無法載入此章節的音訊檔案，可能該版本查無此章音檔。", "error");
    });

    // 3.11 鍵盤快捷鍵 (在非輸入框時，按左右鍵可切換章節)
    document.addEventListener('keydown', (e) => {
        if (['INPUT', 'SELECT', 'TEXTAREA'].includes(document.activeElement.tagName)) return;
        if (e.key === 'ArrowLeft') {
            e.preventDefault();
            playPrevChapter();
        } else if (e.key === 'ArrowRight') {
            e.preventDefault();
            playNextChapter();
        }
    });

    // 3.12 解析 URL 參數進行自動查詢或預設載入詩篇 23
    parseUrlParamsOrInitDefault();
});

// 4. 初始化書卷與章數選單 (區分為舊約與新約兩大 optgroup)
function initBookAndChapterSelectors() {
    bookSelect.innerHTML = '';

    // 4.1 舊約群組
    const otGroup = document.createElement('optgroup');
    otGroup.label = '📜 舊約聖經 (39卷)';
    OLD_TESTAMENT_BOOKS.forEach(b => {
        const opt = document.createElement('option');
        opt.value = b.eng;
        opt.textContent = `${b.full} (${b.chapters}章)`;
        otGroup.appendChild(opt);
    });
    bookSelect.appendChild(otGroup);

    // 4.2 新約群組
    const ntGroup = document.createElement('optgroup');
    ntGroup.label = '✝️ 新約聖經 (27卷)';
    NEW_TESTAMENT_BOOKS.forEach(b => {
        const opt = document.createElement('option');
        opt.value = b.eng;
        opt.textContent = `${b.full} (${b.chapters}章)`;
        ntGroup.appendChild(opt);
    });
    bookSelect.appendChild(ntGroup);

    // 設定初始選取書卷
    const defaultBook = BIBLE_BOOKS[currentBookIndex];
    bookSelect.value = defaultBook.eng;
    updateChapterDropdown(currentBookIndex, currentChapter);
}

// 5. 更新章數下拉選單
function updateChapterDropdown(bookIdx, selectedChap = 1) {
    const book = BIBLE_BOOKS[bookIdx];
    if (!book) return;

    chapterSelect.innerHTML = '';
    for (let c = 1; c <= book.chapters; c++) {
        const opt = document.createElement('option');
        opt.value = c;
        opt.textContent = `第 ${c} 章`;
        if (c === selectedChap) {
            opt.selected = true;
        }
        chapterSelect.appendChild(opt);
    }
}

// 6. 初始化與控制視覺目錄
function initVisualCatalog() {
    // 展開/收合目錄按鈕
    btnToggleCatalog.addEventListener('click', () => {
        const isHidden = catalogPanel.style.display === 'none';
        catalogPanel.style.display = isHidden ? 'flex' : 'none';
        btnToggleCatalog.classList.toggle('active', isHidden);
        catalogToggleText.textContent = isHidden ? '收合書卷目錄' : '展開目錄快速選章';

        if (isHidden) {
            // 開啟時根據當前書卷自動切換舊約/新約
            const isNT = currentBookIndex >= 39;
            switchCatalogTestament(isNT ? 'nt' : 'ot');
            renderCatalogChapters(BIBLE_BOOKS[currentBookIndex]);
        }
    });

    // 舊約/新約 頁籤切換
    testamentTabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const testament = btn.dataset.testament;
            switchCatalogTestament(testament);
        });
    });
}

// 7. 切換視覺目錄的舊約/新約
function switchCatalogTestament(testament) {
    currentCatalogTestament = testament;
    testamentTabBtns.forEach(b => {
        b.classList.toggle('active', b.dataset.testament === testament);
    });

    renderCatalogBooks(testament);
}

// 8. 渲染視覺目錄的書卷按鈕
function renderCatalogBooks(testament) {
    catalogBooksContainer.innerHTML = '';
    const books = testament === 'ot' ? OLD_TESTAMENT_BOOKS : NEW_TESTAMENT_BOOKS;

    books.forEach(book => {
        const bIdx = BIBLE_BOOKS.findIndex(b => b.eng === book.eng);
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'catalog-book-btn';
        btn.textContent = book.full;
        btn.title = `${book.full} (共 ${book.chapters} 章)`;

        if (bIdx === currentBookIndex) {
            btn.classList.add('active');
        }

        btn.addEventListener('click', () => {
            // 標記 active
            catalogBooksContainer.querySelectorAll('.catalog-book-btn').forEach(el => el.classList.remove('active'));
            btn.classList.add('active');

            // 展開章數
            renderCatalogChapters(book);
        });

        catalogBooksContainer.appendChild(btn);
    });
}

// 9. 渲染所選書卷的章數按鈕
function renderCatalogChapters(book) {
    if (!book) return;
    const bIdx = BIBLE_BOOKS.findIndex(b => b.eng === book.eng);

    catalogChaptersSection.style.display = 'flex';
    catalogSelectedBookName.textContent = `${book.full} (${book.chapters}章)`;
    catalogChaptersGrid.innerHTML = '';

    for (let c = 1; c <= book.chapters; c++) {
        const chapBtn = document.createElement('button');
        chapBtn.type = 'button';
        chapBtn.className = 'catalog-chap-btn';
        chapBtn.textContent = c;
        chapBtn.title = `${book.full} 第 ${c} 章`;

        if (bIdx === currentBookIndex && c === currentChapter) {
            chapBtn.classList.add('active');
        }

        chapBtn.addEventListener('click', () => {
            loadBookChapter(bIdx, c, true);
        });

        catalogChaptersGrid.appendChild(chapBtn);
    }
}

// 10. 核心章節載入與切換函式
async function loadBookChapter(bookIdx, chap, autoPlay = true) {
    if (bookIdx < 0 || bookIdx >= BIBLE_BOOKS.length) return;
    const book = BIBLE_BOOKS[bookIdx];
    if (chap < 1 || chap > book.chapters) chap = 1;

    currentBookIndex = bookIdx;
    currentChapter = chap;

    // 10.1 同步更新下拉選單
    bookSelect.value = book.eng;
    updateChapterDropdown(bookIdx, chap);

    // 10.2 同步更新手動查詢框
    queryInput.value = `${book.full} ${chap}`;

    // 10.3 更新前後章導航按鈕狀態與標籤
    updateNavButtonsState();

    // 10.4 若目錄已開啟，同步高亮
    if (catalogPanel.style.display !== 'none') {
        const isNT = bookIdx >= 39;
        if ((isNT && currentCatalogTestament !== 'nt') || (!isNT && currentCatalogTestament !== 'ot')) {
            switchCatalogTestament(isNT ? 'nt' : 'ot');
        } else {
            renderCatalogBooks(currentCatalogTestament);
        }
        renderCatalogChapters(book);
    }

    // 10.5 顯示 Loading 並請求經文資料
    showLoadingState();

    try {
        const queryObj = {
            eng: book.eng,
            short: book.short,
            chap: chap,
            sec: "",
            bookName: book.full
        };

        const res = await FhlBibleService.fetchScripture(queryObj, currentBibleVersion);
        
        const results = [{
            queryObj: queryObj,
            record: res.record,
            records: res.records
        }];

        // 渲染經文
        renderScripture(results);

        // 建立單章播放清單
        activePlaylist = [{
            eng: book.eng,
            chap: chap,
            bookName: book.full,
            bid: bookIdx + 1,
            label: `${book.full} 第 ${chap} 章`
        }];
        currentPlayIndex = 0;
        playlistCard.style.display = 'none';

        // 載入音訊
        loadAudioForChapter(activePlaylist[0], autoPlay);

    } catch (err) {
        console.error("載入經文失敗:", err);
        showErrorState(err.message || "經文載入失敗，請檢查網路連線。");
    }
}

// 11. 計算前一章目標 (支援全本聖經 66 卷跨卷銜接)
function getPrevChapterTarget(bIdx, chap) {
    if (chap > 1) {
        return { bookIndex: bIdx, chap: chap - 1 };
    } else if (bIdx > 0) {
        const prevBook = BIBLE_BOOKS[bIdx - 1];
        return { bookIndex: bIdx - 1, chap: prevBook.chapters };
    }
    return null; // 已是創世記第 1 章
}

// 12. 計算後一章目標 (支援全本聖經 66 卷跨卷銜接)
function getNextChapterTarget(bIdx, chap) {
    const currBook = BIBLE_BOOKS[bIdx];
    if (chap < currBook.chapters) {
        return { bookIndex: bIdx, chap: chap + 1 };
    } else if (bIdx < BIBLE_BOOKS.length - 1) {
        return { bookIndex: bIdx + 1, chap: 1 };
    }
    return null; // 已是啟示錄第 22 章
}

// 13. 更新上一章 / 下一章按鈕狀態與文字提示
function updateNavButtonsState() {
    const prevTarget = getPrevChapterTarget(currentBookIndex, currentChapter);
    const nextTarget = getNextChapterTarget(currentBookIndex, currentChapter);

    const currBook = BIBLE_BOOKS[currentBookIndex];
    footerCurrentLabel.textContent = `${currBook.full} 第 ${currentChapter} 章`;

    // 上一章設定
    if (prevTarget) {
        const prevBook = BIBLE_BOOKS[prevTarget.bookIndex];
        const prevText = `${prevBook.full} ${prevTarget.chap}章`;
        
        btnPrevChap.disabled = false;
        btnHeaderPrev.disabled = false;
        btnFooterPrev.disabled = false;

        btnPrevChap.title = `上一章：${prevText}`;
        btnHeaderPrev.title = `上一章：${prevText}`;
        footerPrevLabel.textContent = prevText;
    } else {
        btnPrevChap.disabled = true;
        btnHeaderPrev.disabled = true;
        btnFooterPrev.disabled = true;

        btnPrevChap.title = "已是聖經第一章";
        btnHeaderPrev.title = "已是聖經第一章";
        footerPrevLabel.textContent = "已是首章";
    }

    // 下一章設定
    if (nextTarget) {
        const nextBook = BIBLE_BOOKS[nextTarget.bookIndex];
        const nextText = `${nextBook.full} ${nextTarget.chap}章`;

        btnNextChap.disabled = false;
        btnHeaderNext.disabled = false;
        btnFooterNext.disabled = false;

        btnNextChap.title = `下一章：${nextText}`;
        btnHeaderNext.title = `下一章：${nextText}`;
        footerNextLabel.textContent = nextText;
    } else {
        btnNextChap.disabled = true;
        btnHeaderNext.disabled = true;
        btnFooterNext.disabled = true;

        btnNextChap.title = "已是聖經最後一章";
        btnHeaderNext.title = "已是聖經最後一章";
        footerNextLabel.textContent = "已是末章";
    }
}

// 14. 播放上一章
function playPrevChapter() {
    // 若在自訂多章播放清單中且非首項，切換清單上一首
    if (activePlaylist.length > 1 && currentPlayIndex > 0) {
        playChapter(currentPlayIndex - 1);
        return;
    }

    const prevTarget = getPrevChapterTarget(currentBookIndex, currentChapter);
    if (prevTarget) {
        loadBookChapter(prevTarget.bookIndex, prevTarget.chap, true);
        showToast(`切換至上一章：${BIBLE_BOOKS[prevTarget.bookIndex].full} 第 ${prevTarget.chap} 章`, "success");
    } else {
        showToast("已經是聖經第一卷第一章 (創世記 第 1 章)", "info");
    }
}

// 15. 播放下一章
function playNextChapter() {
    // 若在自訂多章播放清單中且非末項，切換清單下一首
    if (activePlaylist.length > 1 && currentPlayIndex < activePlaylist.length - 1) {
        playChapter(currentPlayIndex + 1);
        return;
    }

    const nextTarget = getNextChapterTarget(currentBookIndex, currentChapter);
    if (nextTarget) {
        loadBookChapter(nextTarget.bookIndex, nextTarget.chap, true);
        showToast(`切換至下一章：${BIBLE_BOOKS[nextTarget.bookIndex].full} 第 ${nextTarget.chap} 章`, "success");
    } else {
        showToast("已經是聖經最後一卷最後一章 (啟示錄 第 22 章)", "info");
    }
}

// 16. 請求 API 並載入音訊
async function loadAudioForChapter(item, autoPlay = true) {
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
        
        // 載入
        bibleAudio.load();
        
        // 設定播放語速
        bibleAudio.playbackRate = currentSpeed;
        
        // 自動播放
        if (autoPlay) {
            setTimeout(() => {
                bibleAudio.play().catch(err => {
                    console.log("Auto-play was blocked or interrupted:", err);
                    audioSourceName.textContent += " (點擊播放按鈕開始)";
                });
            }, 150);
        }

        // 高亮讀經板中對應當前播放章節的經文
        highlightActiveChapterText(item.eng, item.chap);

    } catch (err) {
        console.error(err);
        audioSourceName.textContent = "無法載入此章音訊";
        showToast(`語音載入失敗: ${err.message}`, "error");
    }
}

// 17. 執行手動經文範圍查詢
async function performQuery() {
    const qStr = queryInput.value.trim();
    if (!qStr) {
        showToast("請輸入經文段落範圍", "error");
        return;
    }

    showLoadingState();

    try {
        const results = await FhlBibleService.query(qStr, currentBibleVersion);
        
        if (!results || results.length === 0) {
            throw new Error("無法解析經文格式");
        }

        // 同步當前書卷與章數至第一筆結果
        const firstQObj = results[0].queryObj;
        const bIdx = BIBLE_BOOKS.findIndex(b => b.eng.toLowerCase() === firstQObj.eng.toLowerCase());
        if (bIdx !== -1) {
            currentBookIndex = bIdx;
            currentChapter = firstQObj.chap;
            bookSelect.value = BIBLE_BOOKS[bIdx].eng;
            updateChapterDropdown(bIdx, currentChapter);
            updateNavButtonsState();
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

// 18. 建立多章播放清單
function buildPlaylist(results) {
    activePlaylist = [];
    playlistContainer.innerHTML = '';

    const seen = new Set();
    
    results.forEach(res => {
        const qObj = res.queryObj;
        const bIdx = BIBLE_BOOKS.findIndex(b => b.eng.toLowerCase() === qObj.eng.toLowerCase());
        const bid = bIdx !== -1 ? bIdx + 1 : -1;
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

    if (activePlaylist.length <= 1) {
        playlistCard.style.display = 'none';
        if (activePlaylist.length === 1) {
            currentPlayIndex = 0;
            loadAudioForChapter(activePlaylist[0], true);
        }
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

// 19. 播放指定索引的清單章節
function playChapter(index) {
    if (index < 0 || index >= activePlaylist.length) return;
    
    currentPlayIndex = index;
    const chapterItem = activePlaylist[index];

    // 更新書卷與章數狀態
    const bIdx = BIBLE_BOOKS.findIndex(b => b.eng.toLowerCase() === chapterItem.eng.toLowerCase());
    if (bIdx !== -1) {
        currentBookIndex = bIdx;
        currentChapter = chapterItem.chap;
        bookSelect.value = BIBLE_BOOKS[bIdx].eng;
        updateChapterDropdown(bIdx, currentChapter);
        updateNavButtonsState();
    }

    // 更新清單項目的 active 狀態
    const items = playlistContainer.querySelectorAll('.playlist-item');
    items.forEach(el => el.classList.remove('active'));
    
    const activeEl = playlistContainer.querySelector(`.playlist-item[data-index="${index}"]`);
    if (activeEl) {
        activeEl.classList.add('active');
        activeEl.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }

    // 載入音訊
    loadAudioForChapter(chapterItem, true);
}

// 20. 渲染經文列表到讀經板
function renderScripture(results) {
    scriptureBody.innerHTML = '';
    
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

    results.forEach(res => {
        const qObj = res.queryObj;
        const records = res.records || [];
        
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

    // 捲動至讀經板頂部
    scriptureBody.scrollTop = 0;
}

// 21. 高亮讀經板對應章節經文
function highlightActiveChapterText(eng, chap) {
    const allVerses = scriptureBody.querySelectorAll('.verse-p');
    allVerses.forEach(el => {
        const id = el.dataset.verseId;
        if (id && id.startsWith(`${eng}_${chap}_`)) {
            el.style.borderLeft = "3px solid var(--primary)";
            el.style.backgroundColor = "rgba(16, 185, 129, 0.04)";
        } else {
            el.style.borderLeft = "none";
            el.style.backgroundColor = "transparent";
        }
    });
}

// 22. 顯示 Loading 與 Error 狀態
function showLoadingState() {
    scriptureTitleDisplay.textContent = "讀經板";
    scriptureInfoDisplay.textContent = "載入中...";
    scriptureBody.innerHTML = `
        <div class="placeholder-msg">
            <div class="logo-icon" style="animation: spin 1s linear infinite; font-size: 28px; width: 44px; height: 44px;">🔄</div>
            <p>正在獲取經文與音訊資料...</p>
        </div>
    `;
    if (!document.getElementById('spin-style')) {
        const style = document.createElement('style');
        style.id = 'spin-style';
        style.innerHTML = `@keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }`;
        document.head.appendChild(style);
    }
}

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

// 23. 解析 URL 參數或載入預設章節
function parseUrlParamsOrInitDefault() {
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
    } else {
        // 預設載入詩篇 23 (Psalms 23)
        loadBookChapter(18, 23, false);
    }
}

// 24. Toast 提示訊息
function showToast(msg, type = "success") {
    const toast = document.getElementById('toast-message');
    if (!toast) return;
    toast.className = `toast ${type}`;
    toast.textContent = msg;
    toast.classList.add('show');

    setTimeout(() => {
        toast.classList.remove('show');
    }, 3000);
}
