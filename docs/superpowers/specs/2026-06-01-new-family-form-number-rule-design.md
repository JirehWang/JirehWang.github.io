# 新家人系統 - 表單號編碼規則修改設計規格書 (New Family Form Number Rule Spec)

本規格書設計修改「新家人管理系統」的表單編號（表單號）產生規則，從原先基於送出時間的隨機代碼（`NF-yyyyMMdd-HHmmss`）變更為符合主日聚會語系分類的連續流水號編碼。

## 1. 變更範圍

本變更僅影響後端 Google Apps Script API 的表單號碼生成邏輯，不影響前端 UI 佈局，但前端的表單送出提示將顯示修改後的新編號。

*   **修改檔案**：
    *   [Code.js](file:///d:/program/LKC/%E6%96%B0%E5%AE%B6%E4%BA%BA%E7%AE%A1%E7%90%86%E7%B3%BB%E7%B5%B1/Code.js) (GAS 後端編碼產生邏輯)

---

## 2. 詳細設計

### 2.1 編碼規則定義
表單編號為 11 位數的數字字串，格式定義為：`[8位數日期] + [1位數聚會代碼] + [2位數每日流水號]`。

#### 1. 日期部分 (8 位數)
取自表單填寫的「日期」欄位（`yyyy-MM-dd`），移除橫線轉換為 `yyyymmdd` (例如：`20260601`)。

#### 2. 聚會分類代碼 (1 位數)
依據表單中「參加的聚會是」欄位值進行比對：
*   字串包含 **`聯合`** $\rightarrow$ 代碼為 **`0`**
*   字串包含 **`台語`** $\rightarrow$ 代碼為 **`1`**
*   字串包含 **`華語`** $\rightarrow$ 代碼為 **`2`**
*   *備註：若以上條件皆不符合，將拋出錯誤「無法識別的聚會分類，表單號產生失敗」，以嚴格確保資料完整性。*

#### 3. 每日流水號 (2 位數)
*   掃描試算表 **「追蹤中」** 與 **「已結案」** 分頁的「表單號」欄位。
*   尋找開頭為 `[8位數日期] + [1位數聚會代碼]` 的所有編號。
*   取最後兩碼為數值並尋找最大值，新表單號之流水號即為 **`最大值 + 1`**（不足兩位前導補零，例如 `01`）。

---

### 2.2 程式碼修改詳細設計 (Code.js)

#### 1. `normalizePayload_` 呼叫端修改
傳入表單中的「日期」與「參加的聚會是」給表單號產生函式：
```javascript
function normalizePayload_(formData, now) {
  const payload = {};
  HEADERS.forEach(header => {
    payload[header] = String(formData[header] || '').trim();
  });

  payload['日期'] = payload['日期'] || Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd');
  // 修改：傳入日期與聚會名稱
  payload['表單號'] = payload['表單號'] || createFormNumber_(payload['日期'], payload['參加的聚會是']);

  return payload;
}
```

#### 2. `createFormNumber_` 實作與輔助函式
```javascript
function createFormNumber_(dateStr, meetingStr) {
  const yyyymmdd = String(dateStr || '').replace(/-/g, '');
  if (yyyymmdd.length !== 8) {
    throw new Error('日期格式錯誤，無法產生表單號');
  }

  let categoryCode = '';
  const meeting = String(meetingStr || '');
  if (meeting.indexOf('聯合') !== -1) {
    categoryCode = '0';
  } else if (meeting.indexOf('台語') !== -1) {
    categoryCode = '1';
  } else if (meeting.indexOf('華語') !== -1) {
    categoryCode = '2';
  } else {
    throw new Error('無法識別的聚會分類：「' + meeting + '」，表單號產生失敗');
  }

  const prefix = yyyymmdd + categoryCode;
  const maxSerial = getMaxSerialNumber_(prefix);
  const nextSerial = maxSerial + 1;
  const serialStr = nextSerial < 10 ? '0' + nextSerial : String(nextSerial);

  return prefix + serialStr;
}

function getMaxSerialNumber_(prefix) {
  const ss = getSpreadsheet_();
  const trackingSheet = ensureSheet_(ss, TRACKING_SHEET_NAME);
  const closedSheet = ensureSheet_(ss, CLOSED_SHEET_NAME);
  
  let maxSerial = 0;
  const formNumberIndex = HEADERS.indexOf('表單號');
  if (formNumberIndex === -1) return 0;

  [trackingSheet, closedSheet].forEach(sheet => {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    const values = sheet.getRange(2, formNumberIndex + 1, lastRow - 1, 1).getValues();
    values.forEach(row => {
      const formNumber = String(row[0] || '').trim();
      if (formNumber.indexOf(prefix) === 0) {
        const serialStr = formNumber.substring(prefix.length);
        const serial = parseInt(serialStr, 10);
        if (!isNaN(serial) && serial > maxSerial) {
          maxSerial = serial;
        }
      }
    });
  });
  
  return maxSerial;
}
```

---

## 3. 驗證計劃 (Verification Plan)

### 3.1 手動功能測試 (在表單網頁填寫測試)
1. **聯合禮拜表單號產生測試**：
   - 選擇日期為 `2026-06-01`，選擇聚會包含 `聯合`（例如「主日聯合崇拜」），送出表單。
   - 預期：回傳的表單號開頭為 `20260601001`。
2. **台語禮拜表單號產生測試**：
   - 選擇日期為 `2026-06-01`，選擇聚會包含 `台語`（例如「台語禮拜」），送出表單。
   - 預期：回傳的表單號開頭為 `20260601101`。
3. **華語禮拜表單號產生與遞增測試**：
   - 選擇日期為 `2026-06-01`，選擇聚會包含 `華語`（例如「華語第一堂」），送出表單。
   - 預期：第一筆為 `20260601201`。
   - 緊接著在同一天再送出一筆「華語」個案。
   - 預期：第二筆自動遞增為 `20260601202`。
4. **跨天流水號重置測試**：
   - 選擇日期為 `2026-06-02`，選擇聚會為「華語禮拜」，送出表單。
   - 預期：回傳的表單號應重置為 `20260602201`。
5. **例外輸入阻斷測試**：
   - 刻意點選不含「聯合/台語/華語」的聚會（若有的話）或以開發者工具送出非法值。
   - 預期：後端拒絕寫入並回傳錯誤訊息「無法識別的聚會分類...」。
