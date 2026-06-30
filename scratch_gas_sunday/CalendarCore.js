const CALENDAR_SPREADSHEET_ID = '1tKI5k7HwI9S2bTRV6RuKrzxBFGzROYHytsFYHuH8H7E';

// --- 快取機制：減少重複開啟試算表的耗時 ---
let _calendarSsCache = null;
function getCalendarSS() {
  if (!_calendarSsCache) _calendarSsCache = SpreadsheetApp.openById(CALENDAR_SPREADSHEET_ID);
  return _calendarSsCache;
}

/**
 * 核心：呼叫 Gemini API 進行講道解析 (具重試機制)
 */
function callGeminiApi(prompt, rawText) {
  const key = _getGeminiApiKey(); // 整合至主主日專案的 GeminiHelper key
  if (!key) throw new Error('未設定 GEMINI_API_KEY，請至主專案指令碼屬性設定。');

  const model = PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || 'gemini-3.1-flash-lite';
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${key}`;
  const payload = {
    "contents": [{ "parts": [{ "text": `${prompt}\n\n待解析文字：\n${rawText}` }] }],
    "generationConfig": { "temperature": 0.1 }
  };
  const options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true };

  let attempts = 0;
  while (attempts < 3) {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    
    if (response.getResponseCode() === 200 && !json.error) {
      return json.candidates[0].content.parts[0].text.replace(/```json/g, '').replace(/```/g, '').trim();
    }
    
    // 處理 503 或 429 錯誤的重試
    if (response.getResponseCode() === 503 || response.getResponseCode() === 429) {
      attempts++;
      Utilities.sleep(2000 * attempts);
      continue;
    }
    throw new Error(`API 錯誤: ${json.error ? json.error.message : '未知錯誤'}`);
  }
}

/**
 * 核心：儲存資料到三張獨立工作表
 */
function saveEvents(events) {
  const ss = getCalendarSS();

  // 初始化工作表並清除舊資料 (保留標題列)
  let mainSheet = ss.getSheetByName('聚會資料') || ss.insertSheet('聚會資料');
  let ministrySheet = ss.getSheetByName('事工細項') || ss.insertSheet('事工細項');
  let sermonSheet = ss.getSheetByName('講道資訊') || ss.insertSheet('講道資訊');

  mainSheet.clear().appendRow(['ID', '日期', '聚會名稱', '聚會類別', '最後更新時間']);
  ministrySheet.clear().appendRow(['聚會ID', '細項ID', '籌備日期', '事工內容']);
  sermonSheet.clear().appendRow(['聚會ID', '講道ID', '講道類別', '講題', '講員', '經文', '宣召', '金句', '詩歌', '備註']);

  const now = new Date();

  // 先收集成 2D 陣列，再以 setValues 一次寫入，將 N 次 appendRow 降為每張表 1 次 RPC
  const mainRows = [];
  const ministryRows = [];
  const sermonRows = [];

  events.forEach(event => {
    mainRows.push([event.id, event.date, event.name, event.category, now]);

    if (event.ministryItems && event.ministryItems.length > 0) {
      event.ministryItems.forEach(min => {
        ministryRows.push([event.id, min.id, min.date, min.content]);
      });
    }

    if (event.sermons && event.sermons.length > 0) {
      event.sermons.forEach(sermon => {
        sermonRows.push([event.id, sermon.id, sermon.type, sermon.title, sermon.speaker, sermon.scripture, sermon.callToWorship, sermon.goldenVerse, sermon.hymns, sermon.description]);
      });
    }
  });

  if (mainRows.length > 0) {
    mainSheet.getRange(2, 1, mainRows.length, 5).setValues(mainRows);
  }
  if (ministryRows.length > 0) {
    ministrySheet.getRange(2, 1, ministryRows.length, 4).setValues(ministryRows);
  }
  if (sermonRows.length > 0) {
    sermonSheet.getRange(2, 1, sermonRows.length, 10).setValues(sermonRows);
  }

  _cal_formatSheet(mainSheet);
  _cal_formatSheet(ministrySheet);
  _cal_formatSheet(sermonSheet);
}

/**
 * 核心：從三張工作表加強讀取並整合
 */
function loadEvents() {
  try {
    const events = _cal_readAll(_CAL_SHEET.EVENTS);
    const types  = _cal_readAll(_CAL_SHEET.TYPES);
    const fields = _cal_readAll(_CAL_SHEET.FIELDS);
    const values = _cal_readAll(_CAL_SHEET.VALUES);

    const typeById  = {};
    types.forEach(t => typeById[t.typeId] = t);
    const fieldById = {};
    fields.forEach(f => fieldById[f.fieldId] = f);

    // 把欄位值按 eventId 分組
    const valuesByEvent = {};
    values.forEach(v => {
      if (!valuesByEvent[v.eventId]) valuesByEvent[v.eventId] = [];
      valuesByEvent[v.eventId].push(v);
    });

    // 找到「講道資訊」頂層類型的 ID
    const sermonRoot = types.find(t => !t.parentTypeId && t['名稱'] === '講道資訊');
    const sermonRootId = sermonRoot ? sermonRoot.typeId : null;

    // 我們只關心「講道資訊」及其子類型的事項
    const sermonEvents = events.filter(e => {
      const type = typeById[e.typeId];
      if (!type) return false;
      let rootType = type;
      while (rootType && rootType.parentTypeId) {
        rootType = typeById[rootType.parentTypeId] || rootType;
        if (rootType.parentTypeId === '' || !rootType.parentTypeId) break;
      }
      return rootType && rootType.typeId === sermonRootId;
    });

    // 按日期分組
    const eventsByDate = {};
    sermonEvents.forEach(e => {
      const dateStr = _cal_dateStr(e['日期']);
      if (!eventsByDate[dateStr]) eventsByDate[dateStr] = [];
      eventsByDate[dateStr].push(e);
    });

    // 構造舊格式的 events 陣列
    const oldEvents = Object.entries(eventsByDate).map(([dateStr, evList]) => {
      // 判斷這天是聯合聚會還是普通聚會
      let hasUnited = false;
      const sermons = evList.map(e => {
        const type = typeById[e.typeId];
        const typeName = type ? type['名稱'] : ''; // '台語', '華語', '聯合-台語', '聯合-華語'
        if (typeName.indexOf('聯合') !== -1) hasUnited = true;

        const evValues = valuesByEvent[e.eventId] || [];
        const sermonObj = {
          id: e.eventId,
          type: typeName, // '台語', '華語', '聯合-台語', '聯合-華語'
          title: '',
          speaker: '',
          scripture: '',
          callToWorship: '',
          goldenVerse: '',
          hymns: '',
          description: ''
        };

        evValues.forEach(v => {
          const f = fieldById[v.fieldId];
          if (!f) return;
          const fname = f['顯示名稱'];
          if (fname === '講題') sermonObj.title = v['值'] || '';
          else if (fname === '講員') sermonObj.speaker = v['值'] || '';
          else if (fname === '經文') sermonObj.scripture = v['值'] || '';
          else if (fname === '宣召') sermonObj.callToWorship = v['值'] || '';
          else if (fname === '金句') sermonObj.goldenVerse = v['值'] || '';
          else if (fname === '詩歌') sermonObj.hymns = v['值'] || '';
          else if (fname === '備註') sermonObj.description = v['值'] || '';
        });

        return sermonObj;
      });

      // 取得第一個事項作為主事項資訊
      const mainEvent = evList[0];
      const category = hasUnited ? '聯合聚會' : '台華語聚會';

      return {
        id: mainEvent.eventId,
        date: dateStr,
        name: mainEvent['顯示標題'] || (hasUnited ? '聯合禮拜' : '主日崇拜'),
        category: category,
        ministryItems: [], // 事工細項在新版獨立，此處回傳空陣列
        sermons: sermons,
        showSub: false
      };
    });

    return oldEvents.sort((a, b) => (a.date < b.date ? -1 : 1));
  } catch (err) {
    console.error('loadEvents error:', err);
    return [];
  }
}

function _cal_formatDate(date) {
  if (!date) return '';
  if (typeof date === 'string') return date;
  const d = new Date(date);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

function _cal_formatSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow === 0) return;
  const headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange.setBackground('#667eea').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.autoResizeColumns(1, lastCol);
  sheet.setFrozenRows(1);
}

/**
 * 查詢信望愛和合本聖經經文
 */
function cal_queryBible(data) {
  const book = data.book || '';
  const chap = data.chap || 1;
  const sec = data.sec || '';
  const version = data.version || 'unv';
  if (!book) return { success: false, message: '缺少經卷名稱 (book)' };
  return fetchBibleText(book, chap, sec, version);
}

/**
 * 查詢信望愛和合本聖經經文
 */
function fetchBibleText(book, chap, sec, version) {
  const url = 'https://bible.fhl.net/json/qb.php';
  const queryVersion = version || 'unv';
  
  let targetUrl = url + '?chineses=' + encodeURIComponent(book) + '&chap=' + chap + '&version=' + queryVersion;
  if (sec) {
    targetUrl += '&sec=' + sec;
  }
  
  try {
    const response = UrlFetchApp.fetch(targetUrl);
    const result = JSON.parse(response.getContentText());
    
    if (result.status === 'success') {
      return {
        success: true,
        version: result.version,
        v_name: result.v_name,
        records: result.record.map(function(r) {
          return {
            chap: r.chap,
            sec: r.sec,
            text: r.bible_text
          };
        })
      };
    } else {
      return { success: false, error: result.status || 'API 回傳失敗' };
    }
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}
