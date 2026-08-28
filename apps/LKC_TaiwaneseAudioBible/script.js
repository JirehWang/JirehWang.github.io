// ==========================================================================
// 台語有聲聖經 (LKC Taiwanese Audio Bible) - script.js
// 長輩友善與銀髮族無障礙核心邏輯（支援彈出視窗選經與設定）
// ==========================================================================

// 1. 聖經 66 卷書對照表與全域狀態變數
let BIBLE_BOOKS = (typeof FhlBibleService !== 'undefined' && FhlBibleService.BIBLE_BOOKS) ? FhlBibleService.BIBLE_BOOKS : [];
let OLD_TESTAMENT_BOOKS = BIBLE_BOOKS.slice(0, 39); // 舊約 39 卷 (創世記 ~ 瑪拉基書)
let NEW_TESTAMENT_BOOKS = BIBLE_BOOKS.slice(39);    // 新約 27 卷 (馬太福音 ~ 啟示錄)

let currentBookIndex = 18; // 預設 詩篇 (索引 18，第 19 卷)
let currentChapter = 23;   // 預設 第 23 章
let currentCatalogTestament = 'ot'; // 視覺目錄目前頁籤：'ot' (舊約) 或 'nt' (新約)

let activePlaylist = []; // 目前播放清單項目：[{ eng, chap, bookName, bid, label }]
let currentPlayIndex = -1;
let currentBibleVersion = 'tghg';
let currentAudioVersion = '1';
let currentSpeed = 1.0;
let currentFontSize = 'large';
let currentTheme = 'parchment';

// 2. DOM 元素變數
let modalPicker, modalSettings;
let btnOpenPicker, btnClosePicker, btnPlayerChangeBook;
let btnOpenSettings, btnCloseSettings, btnSaveSettings;
let bibleVerSelect, audioVerSelect, bookSelect, chapterSelect, btnGoChapter;
let catalogBooksContainer, catalogChaptersSection, catalogSelectedBookName, catalogChaptersGrid, testamentTabBtns;
let manualQueryToggle, manualQueryBody, queryInput, btnQuery;
let playlistCard, playlistContainer;
let seniorPlayerCard, currentChapterTitle, audioSourceName, bibleAudio;
let btnPlayPause, playPauseIcon, playPauseLabel, timeCurrent, timeTotal, audioSeekBar;
let btnPrevChap, btnNextChap, btnHeaderPrev, btnHeaderNext, btnFooterPrev, btnFooterNext;
let footerPrevLabel, footerNextLabel, footerCurrentLabel;
let speedPillBtns, autoNextCheckbox;
let fontSizePillBtns, themePillBtns;
let scriptureTitleDisplay, scriptureInfoDisplay, scriptureBody;

// 3. 初始化事件繫結
document.addEventListener('DOMContentLoaded', () => {
    // 確保 BIBLE_BOOKS 已載入
    if ((!BIBLE_BOOKS || BIBLE_BOOKS.length === 0) && typeof FhlBibleService !== 'undefined' && FhlBibleService.BIBLE_BOOKS) {
        BIBLE_BOOKS = FhlBibleService.BIBLE_BOOKS;
        OLD_TESTAMENT_BOOKS = BIBLE_BOOKS.slice(0, 39);
        NEW_TESTAMENT_BOOKS = BIBLE_BOOKS.slice(39);
    }

    // 獲取彈出視窗與控制按鈕
    modalPicker = document.getElementById('modal-scripture-picker');
    modalSettings = document.getElementById('modal-settings');
    btnOpenPicker = document.getElementById('btn-open-picker');
    btnClosePicker = document.getElementById('btn-close-picker');
    btnPlayerChangeBook = document.getElementById('btn-player-change-book');
    btnOpenSettings = document.getElementById('btn-open-settings');
    btnCloseSettings = document.getElementById('btn-close-settings');
    btnSaveSettings = document.getElementById('btn-save-settings');

    // 獲取選經表單元素
    bibleVerSelect = document.getElementById('bible-ver-select');
    audioVerSelect = document.getElementById('audio-ver-select');
    bookSelect = document.getElementById('book-select');
    chapterSelect = document.getElementById('chapter-select');
    btnGoChapter = document.getElementById('btn-go-chapter');

    // 視覺選經目錄
    catalogBooksContainer = document.getElementById('catalog-books-container');
    catalogChaptersSection = document.getElementById('catalog-chapters-section');
    catalogSelectedBookName = document.getElementById('catalog-selected-book-name');
    catalogChaptersGrid = document.getElementById('catalog-chapters-grid');
    testamentTabBtns = document.querySelectorAll('.tab-btn-lg');

    // 手動查詢
    manualQueryToggle = document.getElementById('manual-query-toggle');
    manualQueryBody = document.getElementById('manual-query-body');
    queryInput = document.getElementById('query-input');
    btnQuery = document.getElementById('btn-query');

    // 清單與播放器
    playlistCard = document.getElementById('playlist-card');
    playlistContainer = document.getElementById('playlist-container');

    seniorPlayerCard = document.getElementById('senior-player-card');
    currentChapterTitle = document.getElementById('current-chapter-title');
    audioSourceName = document.getElementById('audio-source-name');
    bibleAudio = document.getElementById('bible-audio');

    btnPlayPause = document.getElementById('btn-play-pause');
    playPauseIcon = document.getElementById('play-pause-icon');
    playPauseLabel = document.getElementById('play-pause-label');
    timeCurrent = document.getElementById('time-current');
    timeTotal = document.getElementById('time-total');
    audioSeekBar = document.getElementById('audio-seek-bar');

    btnPrevChap = document.getElementById('btn-prev-chap');
    btnNextChap = document.getElementById('btn-next-chap');
    btnHeaderPrev = document.getElementById('btn-header-prev');
    btnHeaderNext = document.getElementById('btn-header-next');
    btnFooterPrev = document.getElementById('btn-footer-prev');
    btnFooterNext = document.getElementById('btn-footer-next');
    footerPrevLabel = document.getElementById('footer-prev-label');
    footerNextLabel = document.getElementById('footer-next-label');
    footerCurrentLabel = document.getElementById('footer-current-label');

    speedPillBtns = document.querySelectorAll('.btn-speed-pill');
    autoNextCheckbox = document.getElementById('auto-next');

    fontSizePillBtns = document.querySelectorAll('.btn-tool-pill');
    themePillBtns = document.querySelectorAll('.btn-theme-pill');

    scriptureTitleDisplay = document.getElementById('scripture-title-display');
    scriptureInfoDisplay = document.getElementById('scripture-info-display');
    scriptureBody = document.getElementById('scripture-body');

    // 3.1 初始化無障礙設定（字級與主題）
    initAccessibilitySettings();

    // 3.2 初始化書卷下拉選單與視覺目錄
    initBookAndChapterSelectors();
    initVisualCatalog();

    // 3.3 初始化播放器自訂控制事件
    initSeniorAudioPlayer();

    // 3.4 彈出視窗（Modal）事件綁定
    initModalControls();

    // 3.5 綁定下拉選單事件
    if (bookSelect) {
        bookSelect.addEventListener('change', () => {
            const eng = bookSelect.value;
            const bIdx = BIBLE_BOOKS.findIndex(b => b.eng === eng);
            if (bIdx !== -1) {
                currentBookIndex = bIdx;
                updateChapterDropdown(currentBookIndex, 1);
            }
        });
    }

    if (btnGoChapter) {
        btnGoChapter.addEventListener('click', () => {
            const eng = bookSelect ? bookSelect.value : '';
            const bIdx = BIBLE_BOOKS.findIndex(b => b.eng === eng);
            const chap = chapterSelect ? (parseInt(chapterSelect.value, 10) || 1) : 1;
            if (bIdx !== -1) {
                closePickerModal();
                loadBookChapter(bIdx, chap, true);
            }
        });
    }

    if (chapterSelect) {
        chapterSelect.addEventListener('change', () => {
            const eng = bookSelect ? bookSelect.value : '';
            const bIdx = BIBLE_BOOKS.findIndex(b => b.eng === eng);
            const chap = parseInt(chapterSelect.value, 10) || 1;
            if (bIdx !== -1) {
                closePickerModal();
                loadBookChapter(bIdx, chap, true);
            }
        });
    }

    // 3.6 手動進階查詢展開/收合切換
    if (manualQueryToggle) {
        manualQueryToggle.addEventListener('click', () => {
            const isHidden = manualQueryBody.style.display === 'none';
            manualQueryBody.style.display = isHidden ? 'flex' : 'none';
            const icon = manualQueryToggle.querySelector('.toggle-icon');
            if (icon) icon.textContent = isHidden ? '－' : '＋';
        });
    }

    // 3.7 手動查詢按鈕與 Enter 鍵
    if (btnQuery) btnQuery.addEventListener('click', () => {
        closePickerModal();
        performQuery();
    });

    if (queryInput) {
        queryInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                closePickerModal();
                performQuery();
            }
        });
    }

    // 3.8 上一章 / 下一章按鈕事件
    if (btnPrevChap) btnPrevChap.addEventListener('click', playPrevChapter);
    if (btnNextChap) btnNextChap.addEventListener('click', playNextChapter);
    if (btnHeaderPrev) btnHeaderPrev.addEventListener('click', playPrevChapter);
    if (btnHeaderNext) btnHeaderNext.addEventListener('click', playNextChapter);
    if (btnFooterPrev) btnFooterPrev.addEventListener('click', playPrevChapter);
    if (btnFooterNext) btnFooterNext.addEventListener('click', playNextChapter);

    // 3.9 譯本切換
    if (bibleVerSelect) {
        bibleVerSelect.addEventListener('change', () => {
            currentBibleVersion = bibleVerSelect.value;
            loadBookChapter(currentBookIndex, currentChapter, false);
        });
    }

    if (audioVerSelect) {
        audioVerSelect.addEventListener('change', () => {
            currentAudioVersion = audioVerSelect.value;
            if (currentPlayIndex >= 0 && activePlaylist[currentPlayIndex]) {
                loadAudioForChapter(activePlaylist[currentPlayIndex], true);
            }
        });
    }

    // 3.10 鍵盤快捷鍵 (在非輸入框時，按左右鍵可切換章節，空白鍵可播放/暫停，ESC 關閉彈窗)
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
            closePickerModal();
            closeSettingsModal();
            return;
        }
        if (['INPUT', 'SELECT', 'TEXTAREA'].includes(document.activeElement.tagName)) return;
        if (e.key === 'ArrowLeft') {
            e.preventDefault();
            playPrevChapter();
        } else if (e.key === 'ArrowRight') {
            e.preventDefault();
            playNextChapter();
        } else if (e.key === ' ' || e.code === 'Space') {
            e.preventDefault();
            togglePlayPause();
        }
    });

    // 3.11 解析 URL 參數進行自動查詢或預設載入詩篇 23
    parseUrlParamsOrInitDefault();
});

// ==========================================================================
// 4. 彈出視窗（Modal）控制
// ==========================================================================
function initModalControls() {
    // 開啟選經彈窗
    if (btnOpenPicker) btnOpenPicker.addEventListener('click', openPickerModal);
    if (btnPlayerChangeBook) btnPlayerChangeBook.addEventListener('click', openPickerModal);

    // 關閉選經彈窗
    if (btnClosePicker) btnClosePicker.addEventListener('click', closePickerModal);
    if (modalPicker) {
        modalPicker.addEventListener('click', (e) => {
            if (e.target === modalPicker) closePickerModal();
        });
    }

    // 開啟設定彈窗
    if (btnOpenSettings) btnOpenSettings.addEventListener('click', openSettingsModal);

    // 關閉設定彈窗
    if (btnCloseSettings) btnCloseSettings.addEventListener('click', closeSettingsModal);
    if (btnSaveSettings) btnSaveSettings.addEventListener('click', () => {
        closeSettingsModal();
        showToast("設定已儲存！", "success");
    });
    if (modalSettings) {
        modalSettings.addEventListener('click', (e) => {
            if (e.target === modalSettings) closeSettingsModal();
        });
    }
}

function openPickerModal() {
    if (modalPicker) {
        modalPicker.style.display = 'flex';
        // 同步選中的書卷與章數
        const currBook = BIBLE_BOOKS[currentBookIndex];
        if (currBook) {
            renderCatalogChapters(currBook);
        }
    }
}

function closePickerModal() {
    if (modalPicker) modalPicker.style.display = 'none';
}

function openSettingsModal() {
    if (modalSettings) modalSettings.style.display = 'flex';
}

function closeSettingsModal() {
    if (modalSettings) modalSettings.style.display = 'none';
}

// ==========================================================================
// 5. 無障礙工具（字級大小與護眼主題記憶）
// ==========================================================================
function initAccessibilitySettings() {
    // 5.1 字級大小
    const savedFontSize = localStorage.getItem('lkc_audiobible_fontsize') || 'large';
    setFontSize(savedFontSize);

    fontSizePillBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const size = btn.dataset.size;
            setFontSize(size);
        });
    });

    // 5.2 主題設定
    const savedTheme = localStorage.getItem('lkc_audiobible_theme') || 'parchment';
    setTheme(['parchment', 'light', 'dark'].includes(savedTheme) ? savedTheme : 'parchment');

    themePillBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const theme = btn.dataset.theme;
            setTheme(theme);
        });
    });
}

function setFontSize(size) {
    currentFontSize = size;
    document.documentElement.setAttribute('data-fontsize', size);
    localStorage.setItem('lkc_audiobible_fontsize', size);

    fontSizePillBtns.forEach(btn => {
        btn.classList.toggle('active', btn.dataset.size === size);
    });
}

function setTheme(theme) {
    currentTheme = theme;
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('lkc_audiobible_theme', theme);

    themePillBtns.forEach(btn => {
        btn.classList.toggle('active', btn.dataset.theme === theme);
    });
}

// ==========================================================================
// 6. 初始化書卷與章數選單 (區分為舊約與新約兩大 optgroup)
// ==========================================================================
function initBookAndChapterSelectors() {
    if (!bookSelect) return;
    bookSelect.innerHTML = '';

    // 6.1 舊約群組
    const otGroup = document.createElement('optgroup');
    otGroup.label = '📜 舊約聖經 (39卷)';
    OLD_TESTAMENT_BOOKS.forEach(b => {
        const opt = document.createElement('option');
        opt.value = b.eng;
        opt.textContent = `${b.full} (${b.chapters}章)`;
        otGroup.appendChild(opt);
    });
    bookSelect.appendChild(otGroup);

    // 6.2 新約群組
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
    if (defaultBook) {
        bookSelect.value = defaultBook.eng;
        updateChapterDropdown(currentBookIndex, currentChapter);
    }
}

// 7. 更新章數下拉選單
function updateChapterDropdown(bookIdx, selectedChap = 1) {
    if (!chapterSelect) return;
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

// ==========================================================================
// 8. 視覺化選經目錄 (Visual Book & Chapter Picker)
// ==========================================================================
function initVisualCatalog() {
    // 舊約/新約 頁籤切換
    testamentTabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const testament = btn.dataset.testament;
            switchCatalogTestament(testament);
        });
    });

    // 預設渲染
    const isNT = currentBookIndex >= 39;
    switchCatalogTestament(isNT ? 'nt' : 'ot');
    if (BIBLE_BOOKS[currentBookIndex]) {
        renderCatalogChapters(BIBLE_BOOKS[currentBookIndex]);
    }
}

function switchCatalogTestament(testament) {
    currentCatalogTestament = testament;
    testamentTabBtns.forEach(b => {
        b.classList.toggle('active', b.dataset.testament === testament);
    });

    renderCatalogBooks(testament);
}

function renderCatalogBooks(testament) {
    if (!catalogBooksContainer) return;
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
            catalogBooksContainer.querySelectorAll('.catalog-book-btn').forEach(el => el.classList.remove('active'));
            btn.classList.add('active');
            renderCatalogChapters(book);
        });

        catalogBooksContainer.appendChild(btn);
    });
}

function renderCatalogChapters(book) {
    if (!book || !catalogChaptersSection || !catalogChaptersGrid) return;
    const bIdx = BIBLE_BOOKS.findIndex(b => b.eng === book.eng);

    catalogChaptersSection.style.display = 'flex';
    if (catalogSelectedBookName) {
        catalogSelectedBookName.textContent = `【${book.full}】`;
    }
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
            closePickerModal();
            loadBookChapter(bIdx, c, true);
        });

        catalogChaptersGrid.appendChild(chapBtn);
    }
}

// ==========================================================================
// 9. 自製長輩大按鈕語音播放器控制邏輯
// ==========================================================================
function initSeniorAudioPlayer() {
    if (!bibleAudio) return;

    // 9.1 巨大播放 / 暫停按鈕
    if (btnPlayPause) {
        btnPlayPause.addEventListener('click', togglePlayPause);
    }

    // 9.2 音訊播放中狀態更新
    bibleAudio.addEventListener('play', () => {
        updatePlayPauseUI(true);
    });

    bibleAudio.addEventListener('pause', () => {
        updatePlayPauseUI(false);
    });

    // 9.3 時間更新與進度條同步
    bibleAudio.addEventListener('timeupdate', () => {
        const curr = bibleAudio.currentTime;
        const dur = bibleAudio.duration;

        if (timeCurrent) timeCurrent.textContent = formatTime(curr);

        if (!isNaN(dur) && dur > 0 && audioSeekBar) {
            const percent = (curr / dur) * 100;
            audioSeekBar.value = percent;
        }
    });

    // 9.4 載入中與元資料就緒
    bibleAudio.addEventListener('loadedmetadata', () => {
        if (timeTotal) timeTotal.textContent = formatTime(bibleAudio.duration);
    });

    bibleAudio.addEventListener('durationchange', () => {
        if (timeTotal) timeTotal.textContent = formatTime(bibleAudio.duration);
    });

    // 9.5 進度條拖曳快轉
    if (audioSeekBar) {
        audioSeekBar.addEventListener('input', () => {
            const dur = bibleAudio.duration;
            if (!isNaN(dur) && dur > 0) {
                const seekTo = (audioSeekBar.value / 100) * dur;
                bibleAudio.currentTime = seekTo;
            }
        });
    }

    // 9.6 語速按鈕快速切換
    speedPillBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const speed = parseFloat(btn.dataset.speed);
            if (!isNaN(speed)) {
                currentSpeed = speed;
                bibleAudio.playbackRate = speed;
                speedPillBtns.forEach(b => b.classList.toggle('active', b.dataset.speed === btn.dataset.speed));
                showToast(`語速已切換為 ${speed}x`, "success");
            }
        });
    });

    // 9.7 音訊播放結束自動下一首
    bibleAudio.addEventListener('ended', () => {
        if (autoNextCheckbox && autoNextCheckbox.checked) {
            playNextChapter();
        }
    });

    // 9.8 錯誤處理
    bibleAudio.addEventListener('error', (e) => {
        console.error("Audio playback error:", e);
        showToast("此章節暫無音訊檔案，或網路連線異常。", "error");
        updatePlayPauseUI(false);
    });
}

function togglePlayPause() {
    if (!bibleAudio) return;
    if (bibleAudio.paused) {
        bibleAudio.play().catch(err => {
            console.log("Audio play blocked or error:", err);
            showToast("請點擊播放鈕開始收聽", "info");
        });
    } else {
        bibleAudio.pause();
    }
}

function updatePlayPauseUI(isPlaying) {
    if (playPauseIcon) playPauseIcon.textContent = isPlaying ? '⏸' : '▶';
    if (playPauseLabel) playPauseLabel.textContent = isPlaying ? '暫停' : '播放';
    if (btnPlayPause) {
        btnPlayPause.classList.toggle('is-playing', isPlaying);
        btnPlayPause.title = isPlaying ? '暫停' : '播放';
    }
}

function formatTime(seconds) {
    if (isNaN(seconds) || seconds < 0) return "00:00";
    const mins = Math.floor(seconds / 60);
    const secs = Math.floor(seconds % 60);
    return `${mins.toString().padStart(2, '0')}:${secs.toString().padStart(2, '0')}`;
}

// ==========================================================================
// 10. 核心章節載入與切換函式
// ==========================================================================
async function loadBookChapter(bookIdx, chap, autoPlay = true) {
    if (bookIdx < 0 || bookIdx >= BIBLE_BOOKS.length) return;
    const book = BIBLE_BOOKS[bookIdx];
    if (chap < 1 || chap > book.chapters) chap = 1;

    currentBookIndex = bookIdx;
    currentChapter = chap;

    // 10.1 同步更新下拉選單
    if (bookSelect) bookSelect.value = book.eng;
    updateChapterDropdown(bookIdx, chap);

    // 10.2 同步更新手動查詢框
    if (queryInput) queryInput.value = `${book.full} ${chap}`;

    // 10.3 更新前後章導航按鈕狀態與標籤
    updateNavButtonsState();

    // 10.4 同步更新視覺目錄的高亮狀態
    const isNT = bookIdx >= 39;
    if ((isNT && currentCatalogTestament !== 'nt') || (!isNT && currentCatalogTestament !== 'ot')) {
        switchCatalogTestament(isNT ? 'nt' : 'ot');
    } else {
        renderCatalogBooks(currentCatalogTestament);
    }
    renderCatalogChapters(book);

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
        if (playlistCard) playlistCard.style.display = 'none';

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
    if (footerCurrentLabel && currBook) {
        footerCurrentLabel.textContent = `${currBook.full} 第 ${currentChapter} 章`;
    }

    // 上一章設定
    if (prevTarget) {
        const prevBook = BIBLE_BOOKS[prevTarget.bookIndex];
        const prevText = `${prevBook.full} ${prevTarget.chap}章`;
        
        if (btnPrevChap) {
            btnPrevChap.disabled = false;
            btnPrevChap.title = `上一章：${prevText}`;
            const textSpan = btnPrevChap.querySelector('.btn-text');
            if (textSpan) textSpan.textContent = `上一章 (${prevText})`;
        }
        if (btnHeaderPrev) {
            btnHeaderPrev.disabled = false;
            btnHeaderPrev.title = `上一章：${prevText}`;
        }
        if (btnFooterPrev) btnFooterPrev.disabled = false;
        if (footerPrevLabel) footerPrevLabel.textContent = prevText;
    } else {
        if (btnPrevChap) {
            btnPrevChap.disabled = true;
            btnPrevChap.title = "已是聖經第一章";
            const textSpan = btnPrevChap.querySelector('.btn-text');
            if (textSpan) textSpan.textContent = "已是首章";
        }
        if (btnHeaderPrev) {
            btnHeaderPrev.disabled = true;
            btnHeaderPrev.title = "已是聖經第一章";
        }
        if (btnFooterPrev) btnFooterPrev.disabled = true;
        if (footerPrevLabel) footerPrevLabel.textContent = "已是首章";
    }

    // 下一章設定
    if (nextTarget) {
        const nextBook = BIBLE_BOOKS[nextTarget.bookIndex];
        const nextText = `${nextBook.full} ${nextTarget.chap}章`;

        if (btnNextChap) {
            btnNextChap.disabled = false;
            btnNextChap.title = `下一章：${nextText}`;
            const textSpan = btnNextChap.querySelector('.btn-text');
            if (textSpan) textSpan.textContent = `下一章 (${nextText})`;
        }
        if (btnHeaderNext) {
            btnHeaderNext.disabled = false;
            btnHeaderNext.title = `下一章：${nextText}`;
        }
        if (btnFooterNext) btnFooterNext.disabled = false;
        if (footerNextLabel) footerNextLabel.textContent = nextText;
    } else {
        if (btnNextChap) {
            btnNextChap.disabled = true;
            btnNextChap.title = "已是聖經最後一章";
            const textSpan = btnNextChap.querySelector('.btn-text');
            if (textSpan) textSpan.textContent = "已是末章";
        }
        if (btnHeaderNext) {
            btnHeaderNext.disabled = true;
            btnHeaderNext.title = "已是聖經最後一章";
        }
        if (btnFooterNext) btnFooterNext.disabled = true;
        if (footerNextLabel) footerNextLabel.textContent = "已是末章";
    }
}

// 14. 播放上一章
function playPrevChapter() {
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
    if (currentChapterTitle) currentChapterTitle.textContent = `【${item.label}】`;
    if (audioSourceName) audioSourceName.textContent = "正在獲取語音...";

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

        if (bibleAudio) {
            bibleAudio.src = data.mp3;
            if (audioSourceName) audioSourceName.textContent = `${data.name || '台語有聲聖經'}`;
            bibleAudio.load();
            bibleAudio.playbackRate = currentSpeed;
            
            if (autoPlay) {
                setTimeout(() => {
                    bibleAudio.play().then(() => {
                        updatePlayPauseUI(true);
                    }).catch(err => {
                        console.log("Auto-play was blocked or interrupted:", err);
                        updatePlayPauseUI(false);
                    });
                }, 150);
            } else {
                updatePlayPauseUI(false);
            }
        }

        highlightActiveChapterText(item.eng, item.chap);

    } catch (err) {
        console.error(err);
        if (audioSourceName) audioSourceName.textContent = "無法載入此章音訊";
        showToast(`語音載入失敗: ${err.message}`, "error");
        updatePlayPauseUI(false);
    }
}

// 17. 執行手動經文範圍查詢
async function performQuery() {
    const qStr = queryInput ? queryInput.value.trim() : '';
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

        const firstQObj = results[0].queryObj;
        const bIdx = BIBLE_BOOKS.findIndex(b => b.eng.toLowerCase() === firstQObj.eng.toLowerCase());
        if (bIdx !== -1) {
            currentBookIndex = bIdx;
            currentChapter = firstQObj.chap;
            if (bookSelect) bookSelect.value = BIBLE_BOOKS[bIdx].eng;
            updateChapterDropdown(bIdx, currentChapter);
            updateNavButtonsState();
        }

        renderScripture(results);
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
    if (playlistContainer) playlistContainer.innerHTML = '';

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
        if (playlistCard) playlistCard.style.display = 'none';
        if (activePlaylist.length === 1) {
            currentPlayIndex = 0;
            loadAudioForChapter(activePlaylist[0], true);
        }
        return;
    }

    if (playlistCard) playlistCard.style.display = 'flex';

    if (playlistContainer) {
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
    }

    playChapter(0);
}

// 19. 播放指定索引的清單章節
function playChapter(index) {
    if (index < 0 || index >= activePlaylist.length) return;
    
    currentPlayIndex = index;
    const chapterItem = activePlaylist[index];

    const bIdx = BIBLE_BOOKS.findIndex(b => b.eng.toLowerCase() === chapterItem.eng.toLowerCase());
    if (bIdx !== -1) {
        currentBookIndex = bIdx;
        currentChapter = chapterItem.chap;
        if (bookSelect) bookSelect.value = BIBLE_BOOKS[bIdx].eng;
        updateChapterDropdown(bIdx, currentChapter);
        updateNavButtonsState();
    }

    if (playlistContainer) {
        const items = playlistContainer.querySelectorAll('.playlist-item');
        items.forEach(el => el.classList.remove('active'));
        
        const activeEl = playlistContainer.querySelector(`.playlist-item[data-index="${index}"]`);
        if (activeEl) {
            activeEl.classList.add('active');
            activeEl.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        }
    }

    loadAudioForChapter(chapterItem, true);
}

// 20. 渲染經文列表到讀經板
function renderScripture(results) {
    if (!scriptureBody) return;
    scriptureBody.innerHTML = '';
    
    const firstQObj = results[0].queryObj;
    const lastQObj = results[results.length - 1].queryObj;
    
    let rangeLabel = `${firstQObj.bookName} ${firstQObj.chap}`;
    if (firstQObj.sec) rangeLabel += `:${firstQObj.sec}`;
    if (results.length > 1 || (firstQObj.chap !== lastQObj.chap)) {
        rangeLabel += ` ~ ${lastQObj.bookName} ${lastQObj.chap}`;
        if (lastQObj.sec) rangeLabel += `:${lastQObj.sec}`;
    }

    if (scriptureTitleDisplay) scriptureTitleDisplay.textContent = "台語聖經經文";
    if (scriptureInfoDisplay) scriptureInfoDisplay.textContent = rangeLabel;

    results.forEach(res => {
        const qObj = res.queryObj;
        const records = res.records || [];
        
        const sectionHeader = document.createElement('div');
        sectionHeader.className = 'scripture-section-header';
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

    window.scrollTo({ top: 0, behavior: 'smooth' });
}

// 21. 高亮讀經板對應章節經文
function highlightActiveChapterText(eng, chap) {
    if (!scriptureBody) return;
    const allVerses = scriptureBody.querySelectorAll('.verse-p');
    allVerses.forEach(el => {
        const id = el.dataset.verseId;
        if (id && id.startsWith(`${eng}_${chap}_`)) {
            el.style.borderLeft = "4px solid var(--primary)";
            el.style.backgroundColor = "var(--verse-highlight)";
        } else {
            el.style.borderLeft = "none";
            el.style.backgroundColor = "transparent";
        }
    });
}

// 22. 顯示 Loading 與 Error 狀態
function showLoadingState() {
    if (scriptureTitleDisplay) scriptureTitleDisplay.textContent = "讀經板";
    if (scriptureInfoDisplay) scriptureInfoDisplay.textContent = "載入中...";
    if (scriptureBody) {
        scriptureBody.innerHTML = `
            <div class="placeholder-msg">
                <div class="logo-icon" style="animation: spin 1s linear infinite; font-size: 28px; width: 48px; height: 48px;">🔄</div>
                <p style="font-size: 18px; font-weight: 700; color: var(--text-primary);">正在為您載入經文與音訊...</p>
            </div>
        `;
    }
    if (!document.getElementById('spin-style')) {
        const style = document.createElement('style');
        style.id = 'spin-style';
        style.innerHTML = `@keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }`;
        document.head.appendChild(style);
    }
}

function showErrorState(msg) {
    if (scriptureTitleDisplay) scriptureTitleDisplay.textContent = "讀經板";
    if (scriptureInfoDisplay) scriptureInfoDisplay.textContent = "錯誤";
    if (scriptureBody) {
        scriptureBody.innerHTML = `
            <div class="placeholder-msg" style="color: var(--danger);">
                <span class="placeholder-icon">⚠️</span>
                <p style="font-size: 18px; font-weight: 700;">${msg}</p>
            </div>
        `;
    }
    if (playlistCard) playlistCard.style.display = 'none';
}

// 23. 解析 URL 參數或載入預設章節
function parseUrlParamsOrInitDefault() {
    const urlParams = new URLSearchParams(window.location.search);
    const queryParam = urlParams.get('query');
    const bibleVerParam = urlParams.get('bible_ver');
    const audioVerParam = urlParams.get('audio_ver');
    const speedParam = urlParams.get('speed');

    if (bibleVerParam && bibleVerSelect) {
        bibleVerSelect.value = bibleVerParam;
        currentBibleVersion = bibleVerParam;
    }
    if (audioVerParam && audioVerSelect) {
        audioVerSelect.value = audioVerParam;
        currentAudioVersion = audioVerParam;
    }
    if (speedParam) {
        const speed = parseFloat(speedParam);
        if (speed >= 0.5 && speed <= 2.0) {
            currentSpeed = speed;
            if (bibleAudio) bibleAudio.playbackRate = speed;
            speedPillBtns.forEach(b => b.classList.toggle('active', parseFloat(b.dataset.speed) === speed));
        }
    }

    if (queryParam) {
        if (queryInput) queryInput.value = queryParam;
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

