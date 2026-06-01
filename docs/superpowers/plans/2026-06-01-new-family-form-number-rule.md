# 新家人系統 - 表單號編碼規則修改實作計劃

> **面向 AI 代理的工作者：** 必需子技能：使用 superpowers:subagent-driven-development（推薦）或 superpowers:executing-plans 逐任務實現此計劃。步驟使用複選框（`- [ ]`）語法來跟蹤進度。

**目標：** 在後端 Apps Script 中修改 `normalizePayload_` 與 `createFormNumber_`，使其產生的表單號符合 `[8位數日期] + [1位數聚會代碼] + [2位數每日分類流水號]` 的編碼規則。

**架構：**
1. 後端 `normalizePayload_` 將表單中的「日期」與「參加的聚會是」傳入 `createFormNumber_`。
2. 後端 `createFormNumber_` 解析日期為 `yyyymmdd`，並依據聚會名稱比對出分類代碼（聯合-0/台語-1/華語-2），接著呼叫 `getMaxSerialNumber_(prefix)`。
3. `getMaxSerialNumber_(prefix)` 查詢「追蹤中」與「已結案」試算表以獲得最大流水號，加一後補零為二位數，組裝為新的表單號。
4. 使用 `clasp push` 部署後端代碼。

**技術棧：** Google Apps Script, Clasp CLI

---

### 任務 1：修改後端 GAS 代碼與部署

**文件：**
- 修改：[Code.js](file:///d:/program/LKC/%E6%96%B0%E5%AE%B6%E4%BA%BA%E7%AE%A1%E7%90%86%E7%B3%BB%E7%B5%B1/Code.js)

- [ ] **步驟 1：修改 `normalizePayload_` 的 `createFormNumber_` 呼叫**
  尋找 [Code.js](file:///d:/program/LKC/%E6%96%B0%E5%AE%B6%E4%BA%BA%E7%AE%A1%E7%90%86%E7%B3%BB%E7%B5%B1/Code.js) 中 `normalizePayload_` 的 `payload['表單號']` 定義列，將原本僅傳入 `now` 的呼叫修改為傳入 `payload['日期']` 與 `payload['參加的聚會是']`。

  **修改目標代碼 (約第 397 行)：**
  ```javascript
  payload['表單號'] = payload['表單號'] || createFormNumber_(payload['日期'], payload['參加的聚會是']);
  ```

- [ ] **步驟 2：重構 `createFormNumber_` 函式並新增 `getMaxSerialNumber_` 輔助函式**
  尋找 [Code.js](file:///d:/program/LKC/%E6%96%B0%E5%AE%B6%E4%BA%BA%E7%AE%A1%E7%90%86%E7%B3%BB%E7%B5%B1/Code.js) 的 `createFormNumber_` 定義列（約第 402-404 行），將其修改為全新的依分類及流水號產生編號的邏輯，並在其後方新增 `getMaxSerialNumber_(prefix)` 函式。

  **修改目標代碼：**
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

- [ ] **步驟 3：使用 clasp 部署後端代碼**
  進入 GAS 目錄並執行 `npx clasp push`。

  運行命令：
  ```powershell
  npx clasp push
  ```
  預期輸出：顯示 `Pushed 2 files.`。

---

### 任務 2：Commit 本地變更

- [ ] **步驟 1：將所有修改 commit 至 Git**
  運行命令：
  ```powershell
  git add docs/superpowers/plans/2026-06-01-new-family-form-number-rule.md
  git commit -m "feat(new-family): add implementation plan for form number rule change"
  ```
