// 聖經 PPT 產生器 - script.js

// 1. 聖經 66 卷書完整對照表
const BIBLE_BOOKS = [
    // 舊約 (39卷)
    { full: "創世記", short: "創", eng: "Gen", chapters: 50 },
    { full: "出埃及記", short: "出", eng: "Ex", chapters: 40 },
    { full: "利未記", short: "利", eng: "Lev", chapters: 27 },
    { full: "民數記", short: "民", eng: "Num", chapters: 36 },
    { full: "申命記", short: "申", eng: "Deut", chapters: 34 },
    { full: "約書亞記", short: "書", eng: "Josh", chapters: 24 },
    { full: "士師記", short: "士", eng: "Judg", chapters: 21 },
    { full: "路得記", short: "得", eng: "Ruth", chapters: 4 },
    { full: "撒母耳記上", short: "撒上", eng: "1Sam", chapters: 31 },
    { full: "撒母耳記下", short: "撒下", eng: "2Sam", chapters: 24 },
    { full: "列王紀上", short: "王上", eng: "1Kings", chapters: 22 },
    { full: "列王紀下", short: "王下", eng: "2Kings", chapters: 25 },
    { full: "歷代志上", short: "代上", eng: "1Chron", chapters: 29 },
    { full: "歷代志下", short: "代下", eng: "2Chron", chapters: 36 },
    { full: "以斯拉記", short: "拉", eng: "Ezra", chapters: 10 },
    { full: "尼希米記", short: "尼", eng: "Neh", chapters: 13 },
    { full: "以斯帖記", short: "斯", eng: "Esth", chapters: 10 },
    { full: "約伯記", short: "伯", eng: "Job", chapters: 42 },
    { full: "詩篇", short: "詩", eng: "Ps", chapters: 150 },
    { full: "箴言", short: "箴", eng: "Prov", chapters: 31 },
    { full: "傳道書", short: "傳", eng: "Eccles", chapters: 12 },
    { full: "雅歌", short: "歌", eng: "Song", chapters: 8 },
    { full: "以賽亞書", short: "賽", eng: "Isa", chapters: 66 },
    { full: "耶利米書", short: "耶", eng: "Jer", chapters: 52 },
    { full: "耶利米哀歌", short: "哀", eng: "Lam", chapters: 5 },
    { full: "以西結書", short: "結", eng: "Ezek", chapters: 48 },
    { full: "但以理書", short: "但", eng: "Dan", chapters: 12 },
    { full: "何西阿書", short: "何", eng: "Hos", chapters: 14 },
    { full: "約珥書", short: "珥", eng: "Joel", chapters: 3 },
    { full: "阿摩司書", short: "摩", eng: "Amos", chapters: 9 },
    { full: "俄巴底亞書", short: "俄", eng: "Obad", chapters: 1 },
    { full: "約拿書", short: "拿", eng: "Jonah", chapters: 4 },
    { full: "彌迦書", short: "彌", eng: "Mic", chapters: 7 },
    { full: "那鴻書", short: "鴻", eng: "Nah", chapters: 3 },
    { full: "哈巴谷書", short: "哈", eng: "Hab", chapters: 3 },
    { full: "西番雅書", short: "番", eng: "Zeph", chapters: 3 },
    { full: "哈該書", short: "該", eng: "Hag", chapters: 2 },
    { full: "撒迦利亞書", short: "亞", eng: "Zech", chapters: 14 },
    { full: "瑪拉基書", short: "瑪", eng: "Mal", chapters: 4 },

    // 新約 (27卷)
    { full: "馬太福音", short: "太", eng: "Matt", chapters: 28 },
    { full: "馬可福音", short: "可", eng: "Mark", chapters: 16 },
    { full: "路加福音", short: "路", eng: "Luke", chapters: 24 },
    { full: "約翰福音", short: "約", eng: "John", chapters: 21 },
    { full: "使徒行傳", short: "徒", eng: "Acts", chapters: 28 },
    { full: "羅馬書", short: "羅", eng: "Rom", chapters: 16 },
    { full: "哥林多前書", short: "林前", eng: "1Cor", chapters: 16 },
    { full: "哥林多後書", short: "林後", eng: "2Cor", chapters: 13 },
    { full: "加拉太書", short: "加", eng: "Gal", chapters: 6 },
    { full: "以弗所書", short: "弗", eng: "Eph", chapters: 6 },
    { full: "腓立比書", short: "腓", eng: "Phil", chapters: 4 },
    { full: "歌羅西書", short: "西", eng: "Col", chapters: 4 },
    { full: "帖撒羅尼迦前書", short: "帖前", eng: "1Thess", chapters: 5 },
    { full: "帖撒羅尼迦後書", short: "帖後", eng: "2Thess", chapters: 3 },
    { full: "提摩太前書", short: "提前", eng: "1Tim", chapters: 6 },
    { full: "提摩太後書", short: "提後", eng: "2Tim", chapters: 4 },
    { full: "提多書", short: "多", eng: "Titus", chapters: 3 },
    { full: "腓利門書", short: "門", eng: "Philem", chapters: 1 },
    { full: "希伯來書", short: "希", eng: "Heb", chapters: 13 },
    { full: "雅各書", short: "雅", eng: "Jas", chapters: 5 },
    { full: "彼得前書", short: "彼前", eng: "1Pet", chapters: 5 },
    { full: "彼得後書", short: "彼後", eng: "2Pet", chapters: 3 },
    { full: "約翰一書", short: "約一", eng: "1John", chapters: 5 },
    { full: "約翰二書", short: "約二", eng: "2John", chapters: 1 },
    { full: "約翰三書", short: "約三", eng: "3John", chapters: 1 },
    { full: "猶大書", short: "猶", eng: "Jude", chapters: 1 },
    { full: "啟示錄", short: "啟", eng: "Rev", chapters: 22 }
];

// 2. 當前應用狀態變數
let currentLanguage = 'tw';
let currentVersion = 'tghg';
let currentQueryObj = null; // { eng, chap, sec, bookName }
let fetchedVerses = []; // API 回傳的完整經節 record 列表
let selectedVerses = []; // 使用者勾選的經節
let slidePages = []; // 分割後的投影片頁面資料：[ [{sec, text, chap}, ...], ... ]
let currentPreviewPageIndex = 0; // 當前預覽的投影片頁面索引
let uploadedBgImageBase64 = null; // 使用者上傳背景圖片的 base64 資料

// 3. UI 元素宣告
const langSelect = document.getElementById('language-select');
const verSelect = document.getElementById('version-select');
const queryInput = document.getElementById('query-input');
const btnQuery = document.getElementById('btn-query');
const versesListContainer = document.getElementById('verses-list-container');
const versesList = document.getElementById('verses-list');
const btnSelectAll = document.getElementById('btn-select-all');

const bgColorInput = document.getElementById('bg-color');
const bgColorText = document.getElementById('bg-color-text');
const bgImageUpload = document.getElementById('bg-image-upload');
const bgImageFilename = document.getElementById('bg-image-filename');

const titleFont = document.getElementById('title-font');
const titleSize = document.getElementById('title-size');
const titleColor = document.getElementById('title-color');
const titleColorText = document.getElementById('title-color-text');
const titleAlign = document.getElementById('title-align');
const titleUnderline = document.getElementById('title-underline');
const titleX = document.getElementById('title-x');
const titleY = document.getElementById('title-y');
const titleW = document.getElementById('title-w');
const titleH = document.getElementById('title-h');
const titleBold = document.getElementById('title-bold');
const titleItalic = document.getElementById('title-italic');

const contentFont = document.getElementById('content-font');
const contentSize = document.getElementById('content-size');
const contentColor = document.getElementById('content-color');
const contentColorText = document.getElementById('content-color-text');
const contentAlign = document.getElementById('content-align');
const contentSpacing = document.getElementById('content-spacing');
const contentX = document.getElementById('content-x');
const contentY = document.getElementById('content-y');
const contentW = document.getElementById('content-w');
const contentH = document.getElementById('content-h');
const contentBold = document.getElementById('content-bold');
const contentItalic = document.getElementById('content-italic');

const layoutModes = document.getElementsByName('layout-mode');
const btnExportPptx = document.getElementById('btn-export-pptx');

const btnPrevSlide = document.getElementById('btn-prev-slide');
const btnNextSlide = document.getElementById('btn-next-slide');
const slidePageIndicator = document.getElementById('slide-page-indicator');
const slidePreviewBox = document.getElementById('slide-preview-box');
const previewTitle = document.getElementById('preview-title');
const previewContent = document.getElementById('preview-content');
const toastMessage = document.getElementById('toast-message');

// 4. 初始化事件綁定
document.addEventListener('DOMContentLoaded', () => {
    // 預設語言與譯本變更聯動
    langSelect.addEventListener('change', handleLanguageChange);
    verSelect.addEventListener('change', () => {
        currentVersion = verSelect.value;
    });

    // 查詢按鈕與 Enter 鍵綁定
    btnQuery.addEventListener('click', performQuery);
    queryInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            performQuery();
        }
    });

    // 全選與反選
    btnSelectAll.addEventListener('click', toggleSelectAll);

    // 顏色選取器雙向綁定
    bindColorInput(bgColorInput, bgColorText);
    bindColorInput(titleColor, titleColorText);
    bindColorInput(contentColor, contentColorText);

    // 背景圖片上傳處理
    bgImageUpload.addEventListener('change', handleBgImageUpload);

    // 模板變動即時重繪預覽
    const templateInputs = [
        bgColorInput, bgImageUpload, titleFont, titleSize, titleColor, 
        titleAlign, titleUnderline, titleBold, titleItalic, titleX, titleY, titleW, titleH,
        contentFont, contentSize, contentColor, contentAlign, contentSpacing,
        contentBold, contentItalic, contentX, contentY, contentW, contentH
    ];
    templateInputs.forEach(input => {
        input.addEventListener('input', updatePreview);
        input.addEventListener('change', updatePreview);
    });
    
    // 排版模式變更時重新排版並更新預覽
    layoutModes.forEach(radio => {
        radio.addEventListener('change', () => {
            recalculateLayout();
            updatePreview();
        });
    });

    // 監聽視窗大小改變以動態縮放預覽字型
    window.addEventListener('resize', updatePreview);

    // 簡報分頁導覽
    btnPrevSlide.addEventListener('click', () => {
        if (currentPreviewPageIndex > 0) {
            currentPreviewPageIndex--;
            updatePreview();
        }
    });
    btnNextSlide.addEventListener('click', () => {
        if (currentPreviewPageIndex < slidePages.length - 1) {
            currentPreviewPageIndex++;
            updatePreview();
        }
    });

    // 匯出按鈕
    btnExportPptx.addEventListener('click', exportToPPTX);

    // 解析網址傳遞的經文參數進行自動查詢
    const urlParams = new URLSearchParams(window.location.search);
    const queryParam = urlParams.get('query');
    const langParam = urlParams.get('lang');
    const verParam = urlParams.get('ver');
    const autoParam = urlParams.get('auto');

    console.log("PPT Generator parameter parser detected parameters:", { queryParam, langParam, verParam, autoParam });

    if (queryParam) {
        queryInput.value = queryParam;
        console.log("Set queryInput.value to:", queryInput.value);
        
        if (langParam) {
            langSelect.value = langParam;
            console.log("Set langSelect.value to:", langSelect.value);
            handleLanguageChange();
        }
        
        if (verParam) {
            verSelect.value = verParam;
            currentVersion = verParam;
            console.log("Set verSelect.value to:", verSelect.value, ", currentVersion set to:", currentVersion);
        }
        
        if (autoParam === '1' || autoParam === 'true') {
            console.log("Auto query execution triggered.");
            performQuery();
        }
    }
});

// 5. 語系與譯本選擇邏輯
function handleLanguageChange() {
    currentLanguage = langSelect.value;
    verSelect.innerHTML = '';
    
    if (currentLanguage === 'tw') {
        // 台語譯本選項
        const opt1 = new Option('台語：巴克禮台漢本 (tghg)', 'tghg');
        const opt2 = new Option('台語：現代台語2021漢字版 (ttvhl2021)', 'ttvhl2021');
        verSelect.add(opt1);
        verSelect.add(opt2);
    } else {
        // 華語譯本選項
        const opt1 = new Option('華語：和合本上帝版 (unv_god)', 'unv_god');
        const opt2 = new Option('華語：和合本神版 (unv)', 'unv');
        verSelect.add(opt1);
        verSelect.add(opt2);
    }
    currentVersion = verSelect.value;
}

// 雙向綁定 Color picker 與文字框
function bindColorInput(picker, textInput) {
    picker.addEventListener('input', () => {
        textInput.value = picker.value.toUpperCase();
        updatePreview();
    });
    textInput.addEventListener('change', () => {
        let val = textInput.value.trim();
        if (!val.startsWith('#')) {
            val = '#' + val;
        }
        if (/^#[0-9A-F]{6}$/i.test(val)) {
            picker.value = val;
            textInput.value = val.toUpperCase();
            updatePreview();
        }
    });
}

// 背景上傳
function handleBgImageUpload(e) {
    const file = e.target.files[0];
    if (!file) {
        bgImageFilename.textContent = '未選擇檔案';
        uploadedBgImageBase64 = null;
        updatePreview();
        return;
    }
    bgImageFilename.textContent = file.name;
    const reader = new FileReader();
    reader.onload = function(evt) {
        uploadedBgImageBase64 = evt.target.result;
        updatePreview();
    };
    reader.readAsDataURL(file);
}

// 6. 智能經文解析與數字轉換邏輯
function chineseToArabic(str) {
    const charMap = {
        '零': 0, '〇': 0,
        '一': 1, '二': 2, '兩': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9,
        '十': 10, '廿': 20, '卅': 30, '卌': 40, '百': 100
    };
    if (/^\d+$/.test(str)) {
        return parseInt(str, 10);
    }
    let val = 0;
    let temp = 0;
    
    for (let i = 0; i < str.length; i++) {
        const char = str[i];
        const num = charMap[char];
        if (num === undefined) continue;
        
        if (num === 10) {
            if (temp === 0) temp = 1;
            val += temp * 10;
            temp = 0;
        } else if (num === 20 || num === 30 || num === 40) {
            val += num;
            temp = 0;
        } else if (num === 100) {
            if (temp === 0) temp = 1;
            val += temp * 100;
            temp = 0;
        } else {
            temp = num;
        }
    }
    val += temp;
    return val;
}

function idToRegex(id) {
    let escaped = id.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
    escaped = escaped.replace(/\s+/g, '\\s*');
    const chars = id.replace(/\s+/g, '').split('');
    const pattern = chars.map(c => c + '\\s*').join('');
    if (/^[a-zA-Z0-9\s]+$/.test(id)) {
        return new RegExp('^' + pattern, 'i');
    } else {
        return new RegExp('^' + pattern);
    }
}

function parseChapterAndVerse(str) {
    let normalized = str.trim()
        .replace(/：/g, ':')
        .replace(/[~～－—至]/g, '-')
        .replace(/\s+/g, ' ');
    
    // 分流正則：
    // regex1: 章與節之間有明確分隔符 (冒號或空格)
    // regex2: 中文數字章直接連著阿拉伯數字節 (無分隔符，如「一1-11」)
    const regex1 = /^([零〇一二兩三四五六七八九十廿卅卌百]+|\d+)\s*[:\s]\s*([\d\-,\s]+)$/;
    const regex2 = /^([零〇一二兩三四五六七八九十廿卅卌百]+)\s*([\d\-,\s]+)$/;
    
    let match = normalized.match(regex1) || normalized.match(regex2);
    if (!match) return null;
    
    const chapStr = match[1];
    const secStr = match[2].replace(/\s+/g, '');
    
    const chap = chineseToArabic(chapStr);
    if (isNaN(chap) || chap <= 0) return null;
    
    return {
        chap: chap,
        sec: secStr
    };
}

function parseScriptureInput(inputStr) {
    const trimmed = inputStr.trim();
    if (!trimmed) return null;
    
    // 生成所有可能匹配的候選書卷，長度遞減排序
    const candidates = [];
    for (const book of BIBLE_BOOKS) {
        candidates.push({ book, id: book.full });
        candidates.push({ book, id: book.short });
        candidates.push({ book, id: book.eng });
    }
    candidates.sort((a, b) => b.id.length - a.id.length);
    
    // 遍歷匹配 (支援回溯)
    for (const cand of candidates) {
        const regex = idToRegex(cand.id);
        const match = trimmed.match(regex);
        if (match) {
            const matchedLength = match[0].length;
            const rest = trimmed.slice(matchedLength).trim();
            const parsedCV = parseChapterAndVerse(rest);
            if (parsedCV) {
                return {
                    eng: cand.book.eng,
                    short: cand.book.short,
                    chap: parsedCV.chap,
                    sec: parsedCV.sec,
                    bookName: cand.book.full
                };
            }
        }
    }
    return null;
}

// 7. 異步 FHL API 查詢與資料載入
async function performQuery() {
    const inputVal = queryInput.value.trim();
    if (!inputVal) {
        showToast('請輸入要查詢的經文，例如「以弗所書 5:1-4」', 'warning');
        return;
    }

    // 支援以分號分割多段查詢 (例如: 詩篇37:25-26; 以賽亞書32:17-18; 提摩太後書3:16)
    // 同時支援省略書卷名稱的同卷跨章節查詢 (例如: 創1:1-3; 4:1-6)
    const parts = inputVal.split(/[;；]/).map(p => p.trim()).filter(p => p.length > 0);
    const queries = [];
    let lastBookObj = null; // 記錄最近一次解析成功的書卷，以供後面省略書卷時繼承

    for (let part of parts) {
        let parsed = parseScriptureInput(part);
        
        if (!parsed) {
            // 如果解析失敗，看看是不是省略了書卷名稱（例如 "4:1-6" 或 "四1-6" 或 "4 1-6"）
            const parsedCV = parseChapterAndVerse(part);
            if (parsedCV && lastBookObj) {
                // 繼承前一段書卷
                parsed = {
                    eng: lastBookObj.eng,
                    short: lastBookObj.short,
                    chap: parsedCV.chap,
                    sec: parsedCV.sec,
                    bookName: lastBookObj.bookName
                };
            }
        }

        if (!parsed) {
            showToast(`無法識別其中一段經文格式: "${part}"，請檢查是否輸入如「以弗所書 5:1-4」`, 'warning');
            return;
        }

        // 記錄當前解析出的書卷以供後續繼承
        lastBookObj = BIBLE_BOOKS.find(b => b.eng === parsed.eng);
        queries.push(parsed);
    }

    // 將第一個 queryObj 設為全域作為 fallback
    currentQueryObj = queries[0];
    showToast(`正在向信望愛聖經 API 載入 ${queries.length} 段經文，請稍候...`, 'info');

    try {
        // 並行發送所有經文段落的 API 請求
        const fetchPromises = queries.map(async (queryObj) => {
            const fhlBook = queryObj.eng === 'Heb' ? '來' : queryObj.short;
            const qstr = `${fhlBook} ${queryObj.chap}:${queryObj.sec}`;
            const apiVersion = currentVersion === 'unv_god' ? 'unv' : currentVersion;
            const apiUrl = `https://bible.fhl.net/json/qsb.php?qstr=${encodeURIComponent(qstr)}&version=${apiVersion}&gb=0`;
            
            const response = await fetch(apiUrl);
            if (!response.ok) {
                throw new Error(`查詢 "${qstr}" 時發生錯誤: HTTP ${response.status}`);
            }
            const data = await response.json();
            
            if (data.status !== 'success') {
                throw new Error(`查詢 "${qstr}" 失敗: ${data.message || 'API 傳回錯誤'}`);
            }

            if (!data.record || data.record.length === 0) {
                throw new Error(`查無此段經文: "${qstr}"`);
            }

            let record = data.record;
            // 若為上帝版，替換「神」字為「上帝」
            if (currentVersion === 'unv_god') {
                record = record.map(v => ({
                    ...v,
                    bible_text: v.bible_text.replace(/(?:[ 　]+|^)神/g, '上帝')
                }));
            }

            // 標記這段經文所屬的書卷、章節與段落 Key，用於排版時跨段落強制分頁與生成正確標題
            return record.map(v => ({
                ...v,
                queryBookName: queryObj.bookName,
                queryChap: queryObj.chap,
                querySec: queryObj.sec,
                queryGroupKey: `${queryObj.bookName}_${queryObj.chap}_${queryObj.sec}`
            }));
        });

        // 等待所有請求完成
        const results = await Promise.all(fetchPromises);
        
        // 合併所有取得的經文
        fetchedVerses = results.flat();
        
        // 預設全選
        selectedVerses = [...fetchedVerses];

        // 渲染經文複選列表
        renderVersesChecklist();
        
        // 重新排版與預覽
        recalculateLayout();
        updatePreview();

        showToast(`成功載入共 ${fetchedVerses.length} 節經文！`, 'success');
        btnExportPptx.disabled = false;

    } catch (error) {
        console.error('查詢聖經 API 發生錯誤:', error);
        showToast(`讀取聖經失敗，原因: ${error.message}，請檢查拼寫後再試。`, 'danger');
    }
}

// 渲染查詢後的經節複選框列表
function renderVersesChecklist() {
    versesList.innerHTML = '';
    versesListContainer.style.display = 'block';

    fetchedVerses.forEach((verse, index) => {
        const item = document.createElement('div');
        item.className = 'verse-item selected';
        item.dataset.index = index;

        // 簡短修飾文字，神版去空白或保留
        const textPreview = verse.bible_text.trim();

        item.innerHTML = `
            <label class="checkbox-label">
                <input type="checkbox" checked value="${index}">
                <span class="checkbox-custom"></span>
                <strong>${verse.chap}:${verse.sec}</strong>
            </label>
            <span class="verse-text-preview">${textPreview}</span>
        `;

        // 綁定項目點擊與 Checkbox 連動
        const checkbox = item.querySelector('input[type="checkbox"]');
        checkbox.addEventListener('change', () => {
            handleVerseSelectionChange(index, checkbox.checked, item);
        });

        // 點擊文字也可切換
        item.addEventListener('click', (e) => {
            if (e.target !== checkbox && !checkbox.contains(e.target) && e.target.tagName !== 'LABEL') {
                checkbox.checked = !checkbox.checked;
                handleVerseSelectionChange(index, checkbox.checked, item);
            }
        });

        versesList.appendChild(item);
    });

    btnSelectAll.textContent = '取消全選';
}

function handleVerseSelectionChange(index, isChecked, itemElement) {
    const verse = fetchedVerses[index];
    if (isChecked) {
        itemElement.classList.add('selected');
        if (!selectedVerses.includes(verse)) {
            selectedVerses.push(verse);
        }
    } else {
        itemElement.classList.remove('selected');
        selectedVerses = selectedVerses.filter(v => v !== verse);
    }
    // 依據勾選結果，排序順序保持與 API 回傳一致
    selectedVerses.sort((a, b) => fetchedVerses.indexOf(a) - fetchedVerses.indexOf(b));
    
    // 重新計算排版與重繪預覽
    recalculateLayout();
    updatePreview();

    // 更新全選按鈕標籤
    const anyUnchecked = fetchedVerses.length !== selectedVerses.length;
    btnSelectAll.textContent = anyUnchecked ? '全選' : '取消全選';
    btnExportPptx.disabled = selectedVerses.length === 0;
}

function toggleSelectAll() {
    const checkboxes = versesList.querySelectorAll('input[type="checkbox"]');
    const items = versesList.querySelectorAll('.verse-item');
    const label = btnSelectAll.textContent;

    if (label === '取消全選') {
        checkboxes.forEach(cb => cb.checked = false);
        items.forEach(item => item.classList.remove('selected'));
        selectedVerses = [];
        btnSelectAll.textContent = '全選';
    } else {
        checkboxes.forEach(cb => cb.checked = true);
        items.forEach(item => item.classList.add('selected'));
        selectedVerses = [...fetchedVerses];
        btnSelectAll.textContent = '取消全選';
    }
    
    recalculateLayout();
    updatePreview();
    btnExportPptx.disabled = selectedVerses.length === 0;
}

// 8. 核心 Canvas 模擬排版與斷頁演算法
function recalculateLayout() {
    if (selectedVerses.length === 0) {
        slidePages = [];
        currentPreviewPageIndex = 0;
        return;
    }

    const layoutMode = document.querySelector('input[name="layout-mode"]:checked').value;

    if (layoutMode === 'single') {
        // 單節單頁模式
        slidePages = selectedVerses.map(v => [v]);
    } else {
        // 最大內文版面自動斷頁模式
        slidePages = paginateVersesAuto();
    }

    currentPreviewPageIndex = 0;
}

function paginateVersesAuto() {
    const pages = [];
    let currentPage = [];
    
    // 獲取內文邊界設定 (以英吋為單位)
    const cw = parseFloat(contentW.value) / 100 * 10; // 容器寬度 (英吋)
    const ch = parseFloat(contentH.value) / 100 * 5.625; // 容器高度限制 (英吋)
    const size = parseFloat(contentSize.value); // 內文字型大小 (pt)
    const font = contentFont.value; // 字型
    const spacing = parseFloat(contentSpacing.value); // 行距比例

    // 遍歷所有勾選的經文
    for (let i = 0; i < selectedVerses.length; i++) {
        const verse = selectedVerses[i];
        
        // 1. 如果當前頁面已經有經文，且即將加入的這一節屬於不同的 queryGroupKey，就必須強制分開（斷頁），避免不同段經文拼在同一頁上
        if (currentPage.length > 0 && currentPage[0].queryGroupKey !== verse.queryGroupKey) {
            pages.push(currentPage);
            currentPage = [verse];
            continue;
        }

        // 2. 否則進行常規排版模擬高度計算
        const testPage = [...currentPage, verse];
        const combinedText = buildCombinedText(testPage);

        // 透過 Canvas 計算總高度 (像素)
        const heightPx = calculateTextHeight(combinedText, cw, size, font, spacing);
        const heightInches = heightPx / 96; // 1 inch = 96 pixels

        if (heightInches > ch && currentPage.length > 0) {
            // 如果加了這一節會溢出，且當前頁不是空的，就代表這節該放下一頁
            pages.push(currentPage);
            currentPage = [verse];
        } else {
            // 還裝得下，直接裝入
            currentPage.push(verse);
        }
    }

    if (currentPage.length > 0) {
        pages.push(currentPage);
    }

    return pages;
}

// 將分組的經文編排成文字，例如： "5:3 經文內容\n5:4 經文內容"
function buildCombinedText(pageVerses) {
    return pageVerses.map(v => `${v.chap}:${v.sec} ${v.bible_text.trim()}`).join('\n');
}

// Canvas 計算文字高度 (Word-wrap 模擬)
function calculateTextHeight(text, containerWidthInches, fontSizePt, fontFace, lineSpacingMultiplier) {
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    
    // 1 pt = 96 / 72 pixels (約 1.3333 像素)
    const fontSizePx = fontSizePt * (96 / 72);
    ctx.font = `${fontSizePx}px "${fontFace}", "Microsoft JhengHei"`;
    
    // 邊界寬度轉換成像素 (1 inch = 96 pixels)
    const maxWidthPx = containerWidthInches * 96;
    
    // 切割每一行 (處理文字中已有的 \n)
    const rawLines = text.split('\n');
    let totalLineCount = 0;

    rawLines.forEach(lineText => {
        // 對每一行進行 Word-wrap 計算
        const chars = lineText.split('');
        let currentLine = '';

        for (let j = 0; j < chars.length; j++) {
            const char = chars[j];
            const testLine = currentLine + char;
            const metrics = ctx.measureText(testLine);

            if (metrics.width > maxWidthPx && currentLine.length > 0) {
                totalLineCount++;
                currentLine = char;
            } else {
                currentLine = testLine;
            }
        }
        if (currentLine.length > 0) {
            totalLineCount++;
        }
    });

    // 總高度 = 行數 * 單行高
    const singleLineHeightPx = fontSizePx * lineSpacingMultiplier;
    return totalLineCount * singleLineHeightPx;
}

// 9. 所見即所得預覽更新
function updatePreview() {
    if (slidePages.length === 0) {
        // 沒有資料
        slidePageIndicator.textContent = '第 0 / 0 頁';
        btnPrevSlide.disabled = true;
        btnNextSlide.disabled = true;
        previewTitle.textContent = '以弗所書 5:1-4';
        previewContent.textContent = '經文預覽內容將顯示於此...';
        
        // 套用基本黑色背景
        slidePreviewBox.style.backgroundColor = '#000000';
        slidePreviewBox.style.backgroundImage = 'none';
        return;
    }

    // 更新翻頁控制
    slidePageIndicator.textContent = `第 ${currentPreviewPageIndex + 1} / ${slidePages.length} 頁`;
    btnPrevSlide.disabled = currentPreviewPageIndex === 0;
    btnNextSlide.disabled = currentPreviewPageIndex === slidePages.length - 1;

    // 取得當前頁資料
    const currentPageVerses = slidePages[currentPreviewPageIndex];
    
    // 計算標題 (例如 "以弗所書 5:1-2" 或 "以弗所書 5:3")
    const titleText = buildTitleRangeText(currentPageVerses);
    const contentText = buildCombinedText(currentPageVerses);

    // 1. 套用背景
    slidePreviewBox.style.backgroundColor = bgColorInput.value;
    if (uploadedBgImageBase64) {
        slidePreviewBox.style.backgroundImage = `url(${uploadedBgImageBase64})`;
    } else {
        slidePreviewBox.style.backgroundImage = 'none';
    }

    // 計算預覽視窗縮放比例 (簡報在 96 dpi 下的標準尺寸為 960 x 540 像素)
    const previewBoxWidth = slidePreviewBox.offsetWidth || 960;
    const scale = previewBoxWidth / 960;
    const ptToPx = 96 / 72; // pt 轉 px 比例

    // 2. 標題預覽渲染
    previewTitle.textContent = titleText;
    previewTitle.style.left = `${titleX.value}%`;
    previewTitle.style.top = `${titleY.value}%`;
    previewTitle.style.width = `${titleW.value}%`;
    previewTitle.style.height = `${titleH.value}%`;
    previewTitle.style.fontFamily = `"${titleFont.value}", "Microsoft JhengHei"`;
    previewTitle.style.fontSize = `${parseFloat(titleSize.value) * ptToPx * scale}px`; // 動態等比例字型大小
    previewTitle.style.color = titleColor.value;
    previewTitle.style.textAlign = titleAlign.value;
    previewTitle.style.fontWeight = titleBold.checked ? 'bold' : 'normal';
    previewTitle.style.fontStyle = titleItalic.checked ? 'italic' : 'normal';
    previewTitle.style.borderBottom = titleUnderline.checked ? `${2 * scale}px solid ${titleColor.value}` : 'none';

    // 3. 內文預覽渲染
    previewContent.textContent = contentText;
    previewContent.style.left = `${contentX.value}%`;
    previewContent.style.top = `${contentY.value}%`;
    previewContent.style.width = `${contentW.value}%`;
    previewContent.style.height = `${contentH.value}%`;
    previewContent.style.fontFamily = `"${contentFont.value}", "Microsoft JhengHei"`;
    previewContent.style.fontSize = `${parseFloat(contentSize.value) * ptToPx * scale}px`; // 動態等比例字型大小
    previewContent.style.color = contentColor.value;
    previewContent.style.textAlign = contentAlign.value;
    previewContent.style.fontWeight = contentBold.checked ? 'bold' : 'normal';
    previewContent.style.fontStyle = contentItalic.checked ? 'italic' : 'normal';
    
    // 行高套用 (以像素為單位設定行高，確保行高也等比例縮放)
    const calculatedLineHeight = parseFloat(contentSize.value) * ptToPx * parseFloat(contentSpacing.value) * scale;
    previewContent.style.lineHeight = `${calculatedLineHeight}px`;
}

// 根據當前頁包含的經文，生成標題範圍，如 "以弗所書 5:1-2"
function buildTitleRangeText(pageVerses) {
    if (!pageVerses || pageVerses.length === 0) return '';
    
    // 直接從當前頁的第一個經節中取得其所屬的書卷和章資訊，支援多段跨書卷經文
    const bookName = pageVerses[0].queryBookName || (currentQueryObj ? currentQueryObj.bookName : '聖經');
    const chap = pageVerses[0].queryChap || pageVerses[0].chap;
    
    if (pageVerses.length === 1) {
        return `${bookName} ${chap}:${pageVerses[0].sec}`;
    }

    // 取得最小節與最大節
    const startSec = pageVerses[0].sec;
    const endSec = pageVerses[pageVerses.length - 1].sec;
    return `${bookName} ${chap}:${startSec}-${endSec}`;
}

// 10. PPTX 生成與下載匯出
function exportToPPTX() {
    try {
        if (typeof PptxGenJS === 'undefined') {
            alert('聖經 PPT 產生器錯誤：找不到 PptxGenJS 簡報繪圖庫。\n請檢查您的網路連線是否正常，若在教會投影等離線環境，建議將 PptxGenJS 庫下載存於本地使用。');
            return;
        }

        if (slidePages.length === 0) {
            showToast('無可匯出的投影片頁面。', 'warning');
            return;
        }

        showToast('正在生成 PPTX 檔案，請稍候...', 'info');

        // 初始化 PptxGenJS
        const pptx = new PptxGenJS();
        
        // 設定 Layout 比例為 16:9
        pptx.layout = 'LAYOUT_16x9'; // 10" x 5.625" (25.4cm x 14.28cm)
        const slideW = 10;
        const slideH = 5.625;

        // 讀取 UI 設定，做防禦性防 NaN 處理
        const bgFill = bgColorInput.value.replace('#', '');
        const tFont = titleFont.value;
        const tSize = parseInt(titleSize.value, 10) || 40;
        const tColor = titleColor.value.replace('#', '');
        const tAlign = titleAlign.value;
        const tUnder = titleUnderline.checked;
        const tBold = titleBold.checked;
        const tItalic = titleItalic.checked;
        
        const txVal = (parseFloat(titleX.value) || 0) / 100 * slideW;
        const tyVal = (parseFloat(titleY.value) || 0) / 100 * slideH;
        const twVal = (parseFloat(titleW.value) || 10) / 100 * slideW;
        const thVal = (parseFloat(titleH.value) || 5) / 100 * slideH;

        const cFont = contentFont.value;
        const cSize = parseInt(contentSize.value, 10) || 28;
        const cColor = contentColor.value.replace('#', '');
        const cAlign = contentAlign.value;
        const cSpacing = parseFloat(contentSpacing.value) || 1.3;
        const cBold = contentBold.checked;
        const cItalic = contentItalic.checked;

        const cxVal = (parseFloat(contentX.value) || 0) / 100 * slideW;
        const cyVal = (parseFloat(contentY.value) || 0) / 100 * slideH;
        const cwVal = (parseFloat(contentW.value) || 10) / 100 * slideW;
        const chVal = (parseFloat(contentH.value) || 5) / 100 * slideH;

        // 相容性判定：取得正確的 LINE ShapeType
        // 部分 PptxGenJS 舊版使用 pptx.shapes.LINE，而新版使用 pptx.ShapeType.line
        let lineShapeType = null;
        if (pptx.ShapeType && typeof pptx.ShapeType.line !== 'undefined') {
            lineShapeType = pptx.ShapeType.line;
        } else if (pptx.shapes && typeof pptx.shapes.LINE !== 'undefined') {
            lineShapeType = pptx.shapes.LINE;
        } else {
            lineShapeType = 'line'; // fallback 字串
        }

        // 遍歷所有投影片分頁進行繪製
        slidePages.forEach((pageVerses) => {
            const slide = pptx.addSlide();

            // 1. 背景處理
            if (uploadedBgImageBase64) {
                slide.background = { data: uploadedBgImageBase64 };
            } else {
                slide.background = { fill: bgFill };
            }

            // 2. 標題文字繪製
            const titleTextText = buildTitleRangeText(pageVerses);
            slide.addText(titleTextText, {
                x: txVal,
                y: tyVal,
                w: twVal,
                h: thVal,
                fontFace: tFont,
                fontSize: tSize,
                color: tColor,
                align: tAlign,
                valign: 'middle',
                bold: tBold,
                italic: tItalic,
                margin: 0
            });

            // 如果啟用了底線白線 (模擬高質感底線 shape，不緊貼文字基線)
            if (tUnder && lineShapeType) {
                try {
                    // 在標題文字框下方繪製一條直線
                    slide.addShape(lineShapeType, {
                        x: txVal,
                        y: tyVal + thVal + 0.05, // 往下一點點做橫線
                        w: twVal,
                        h: 0,
                        line: { color: tColor, width: 2 } // 線條粗細 2pt
                    });
                } catch (shapeErr) {
                    console.warn('繪製標題底線 shape 發生異常:', shapeErr);
                }
            }

            // 3. 內文文字繪製
            const bodyCombinedText = buildCombinedText(pageVerses);
            
            // PptxGenJS 行高計算
            const calculatedLineSpacingPt = Math.round(cSize * cSpacing);

            slide.addText(bodyCombinedText, {
                x: cxVal,
                y: cyVal,
                w: cwVal,
                h: chVal,
                fontFace: cFont,
                fontSize: cSize,
                color: cColor,
                align: cAlign,
                valign: 'top',
                lineSpacing: calculatedLineSpacingPt,
                bold: cBold,
                italic: cItalic,
                margin: 0
            });
        });

        // 下載檔案
        const fileName = `聖經_${buildTitleRangeText(slidePages[0])}等.pptx`;
        pptx.writeFile({ fileName: fileName })
            .then(() => {
                showToast('PPTX 簡報檔已成功產出並開始下載！', 'success');
            })
            .catch((err) => {
                console.error('PPTX 生成失敗:', err);
                alert(`簡報匯出失敗！錯誤原因: ${err.message || err}`);
                showToast(`簡報產出失敗: ${err.message}`, 'danger');
            });
    } catch (err) {
        console.error('匯出 PPTX 過程中發生致命例外錯誤:', err);
        alert(`聖經 PPT 產生器匯出異常！\n\n【錯誤訊息】\n${err.message}\n\n【詳細堆疊】\n${err.stack}`);
        showToast('簡報生成發生例外錯誤！', 'danger');
    }
}

// 11. 輔助 UI 函式：Toast 浮動卡片提示
function showToast(message, type = 'info') {
    toastMessage.textContent = message;
    
    // 重置樣式
    toastMessage.className = 'toast-card';
    
    // 依據 type 套用顏色
    if (type === 'success') {
        toastMessage.style.backgroundColor = 'rgba(16, 185, 129, 0.15)';
        toastMessage.style.borderColor = 'rgba(16, 185, 129, 0.3)';
        toastMessage.style.color = '#a7f3d0';
    } else if (type === 'warning') {
        toastMessage.style.backgroundColor = 'rgba(245, 158, 11, 0.15)';
        toastMessage.style.borderColor = 'rgba(245, 158, 11, 0.3)';
        toastMessage.style.color = '#fde68a';
    } else if (type === 'danger') {
        toastMessage.style.backgroundColor = 'rgba(239, 68, 68, 0.15)';
        toastMessage.style.borderColor = 'rgba(239, 68, 68, 0.3)';
        toastMessage.style.color = '#fca5a5';
    } else {
        // info
        toastMessage.style.backgroundColor = 'rgba(59, 130, 246, 0.15)';
        toastMessage.style.borderColor = 'rgba(59, 130, 246, 0.3)';
        toastMessage.style.color = '#bfdbfe';
    }
}
