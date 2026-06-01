# 新家人系統 - 單筆資料刪除實作計劃

> **面向 AI 代理的工作者：** 必需子技能：使用 superpowers:subagent-driven-development（推薦）或 superpowers:executing-plans 逐任務實現此計劃。步驟使用複選框（`- [ ]`）語法來跟蹤進度。

**目標：** 在前端「追蹤中」表格之操作欄位新增「刪除」按鈕，並於後端 GAS 新增對應的刪除 API 以清除 Google Sheet 中對應的行。

**架構：** 
1. 前端 CSS 擴大操作欄寬度以並排容納「編輯」與「刪除」按鈕。
2. 前端 JS 表格渲染時新增「刪除」按鈕及其 `click` 事件監聽，並呼叫後端 API。
3. 後端 Apps Script (`Code.js`) 的 `doPost` 路由分流並實作 `deleteTrackingCase`。
4. 使用 `clasp push` 將後端代碼發佈至 Google Apps Script。

**技術棧：** Vanilla JavaScript, CSS, Google Apps Script, Clasp CLI

---

### 任務 1：微調前端樣式與佈局

**文件：**
- 修改：[style.css](file:///d:/program/Github/LKC1958_June_1.github.io/apps/LKC_NewFamily/style.css)

- [ ] **步驟 1：調整 `.action-cell` 寬度**
  尋找 [style.css](file:///d:/program/Github/LKC1958_June_1.github.io/apps/LKC_NewFamily/style.css) 的 `.action-cell` 樣式定義，將其寬度從 `92px` 修改為 `148px`。

  **修改目標代碼：**
  ```css
  .action-cell {
    width: 148px;
  }
  ```

---

### 任務 2：修改前端主要邏輯

**文件：**
- 修改：[script.js](file:///d:/program/Github/LKC1958_June_1.github.io/apps/LKC_NewFamily/script.js)

- [ ] **步驟 1：修改 `buildCaseTable` 以支援雙按鈕操作欄**
  尋找 [script.js](file:///d:/program/Github/LKC1958_June_1.github.io/apps/LKC_NewFamily/script.js) 的 `buildCaseTable` 函式，將表頭文字「編輯」修改為「操作」，並在 `editCell` (即 `actionCell`) 同時加入「編輯」與「刪除」按鈕。

  **修改目標代碼 (約第 1103 行及 1115-1124 行)：**
  ```javascript
  // 修改表頭
  headRow.innerHTML = `${selectable ? '<th class="check-cell">結案</th><th class="action-cell">操作</th>' : ''}${columns.map(column => `<th>${escapeHtml(getColumnLabel(column))}</th>`).join('')}`;
  
  // ...
  
  // 修改儲存格內容
  if (selectable) {
    const checkboxCell = document.createElement('td');
    checkboxCell.className = 'check-cell';
    checkboxCell.innerHTML = `<input type="checkbox" value="${item.rowNumber}" aria-label="勾選 ${escapeHtml(item['新家人姓名'] || '此筆資料')} 結案">`;
    row.appendChild(checkboxCell);

    const actionCell = document.createElement('td');
    actionCell.className = 'action-cell';
    
    // 編輯按鈕
    const editButton = document.createElement('button');
    editButton.type = 'button';
    editButton.className = 'btn secondary';
    editButton.textContent = '編輯';
    editButton.addEventListener('click', () => openEditModal(item));
    actionCell.appendChild(editButton);

    // 刪除按鈕
    const deleteButton = document.createElement('button');
    deleteButton.type = 'button';
    deleteButton.className = 'btn danger';
    deleteButton.textContent = '刪除';
    deleteButton.style.marginLeft = '6px';
    deleteButton.addEventListener('click', () => deleteSingleCase(item));
    actionCell.appendChild(deleteButton);
    
    row.appendChild(actionCell);
  }
  ```

- [ ] **步驟 2：實作 `deleteSingleCase` 函數**
  在 `script.js` 底部（例如 `closeSelectedCases` 函式下方）新增 `deleteSingleCase(item)` 函式，彈出確認提示並呼叫 `callApi` 進行刪除。

  **新增代碼：**
  ```javascript
  async function deleteSingleCase(item) {
    const name = item['新家人姓名'] || '此筆資料';
    if (!confirm(`確認要永久刪除新朋友「${name}」的追蹤資料嗎？此操作將無法復原。`)) {
      return;
    }

    setNotice(trackingNotice, '刪除中...');
    
    try {
      const result = await callApi('deleteTrackingCase', { rowNumber: item.rowNumber });
      setNotice(trackingNotice, result.message, 'success');
      await loadTrackingCases();
    } catch (error) {
      setNotice(trackingNotice, error.message || String(error), 'error');
    }
  }
  ```

---

### 任務 3：實作後端 GAS 邏輯與部署

**文件：**
- 修改：[Code.js](file:///d:/program/LKC/%E6%96%B0%E5%AE%B6%E4%BA%BA%E7%AE%A1%E7%90%86%E7%B3%BB%E7%B5%B1/Code.js)

- [ ] **步驟 1：在 `doPost` 中新增刪除路由分流**
  尋找 [Code.js](file:///d:/program/LKC/%E6%96%B0%E5%AE%B6%E4%BA%BA%E7%AE%A1%E7%90%86%E7%B3%BB%E7%B5%B1/Code.js) 的 `doPost` 函式，加入對 `deleteTrackingCase` action 的判斷。

  **修改目標代碼：**
  ```javascript
  if (action === 'closeCases') return createJsonResponse(closeCases(data.rowNumbers || data.rows || []));
  if (action === 'deleteTrackingCase') return createJsonResponse(deleteTrackingCase(data));
  ```

- [ ] **步驟 2：實作後端 `deleteTrackingCase` 函式**
  在 `Code.js` 底部新增 `deleteTrackingCase` 函式以刪除試算表列。

  **新增代碼：**
  ```javascript
  function deleteTrackingCase(data) {
    const rowNumber = Number(data && data.rowNumber);

    if (!Number.isInteger(rowNumber) || rowNumber < 2) {
      throw new Error('找不到要刪除的追蹤中資料');
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    try {
      const ss = getSpreadsheet_();
      const sheet = ensureSheet_(ss, TRACKING_SHEET_NAME);
      if (rowNumber > sheet.getLastRow()) {
        throw new Error('資料列已變更，請重新整理後再刪除');
      }

      sheet.deleteRow(rowNumber);
      refreshNewFamilyCaches_();
      return {
        success: true,
        message: '已刪除追蹤中資料'
      };
    } finally {
      lock.releaseLock();
    }
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

### 任務 4：Commit 本地變更

- [ ] **步驟 1：將所有修改 commit 至 Git**
  運行命令：
  ```powershell
  git add apps/LKC_NewFamily/style.css apps/LKC_NewFamily/script.js
  git commit -m "feat(new-family): add single case delete feature in tracking panel"
  ```
