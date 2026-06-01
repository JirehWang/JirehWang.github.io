# 新家人系統 - 單筆資料刪除功能設計規格書 (New Family Single Case Delete Spec)

本規格書設計在「新家人管理系統」的「追蹤中」分頁加入單筆資料刪除功能。當同工發現建立錯誤或重複的個案資料時，可以直接在該筆資料列點擊「刪除」按鈕，在二次確認後將資料自 Google Sheet 中永久刪除，並自動更新 Firebase 清單快取。

## 1. 變更範圍

本變更涉及後端 Google Apps Script API 以及前端網頁與 CSS 樣式。

*   **修改檔案**：
    *   [Code.js](file:///d:/program/LKC/%E6%96%B0%E5%AE%B6%E4%BA%BA%E7%AE%A1%E7%90%86%E7%B3%BB%E7%B5%B1/Code.js) (GAS 後端新增 API Action)
    *   [script.js](file:///d:/program/Github/LKC1958_June_1.github.io/apps/LKC_NewFamily/script.js) (前端渲染與 API 呼叫邏輯)
    *   [style.css](file:///d:/program/Github/LKC1958_June_1.github.io/apps/LKC_NewFamily/style.css) (操作欄位 CSS 寬度調整)

---

## 2. 詳細設計

### 2.1 後端 GAS 設計 (Code.js)
在後端新增一個 API action: `deleteTrackingCase`，接收要刪除的 `rowNumber`，以原子鎖（LockService）保護並將該列自「追蹤中」試算表刪除，隨後刷新快取。

#### `doPost` 路由新增：
```javascript
if (action === 'deleteTrackingCase') return createJsonResponse(deleteTrackingCase(data));
```

#### `deleteTrackingCase(data)` 實作：
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

---

### 2.2 前端 JavaScript 設計 (script.js)
1. 在 `buildCaseTable` 中，將表頭的 `<th class="action-cell">編輯</th>` 修改為 `<th class="action-cell">操作</th>`。
2. 在相同的 td（`action-cell`）中，除了原有的「編輯」按鈕外，新增一個「刪除」按鈕。
3. 實作 `deleteSingleCase(item)` 函數，點擊時彈出確認對話框，經確認後呼叫 API 進行刪除，並重新整理追蹤中清單。

#### `buildCaseTable` 修改預覽：
```javascript
// 修改前：
const editCell = document.createElement('td');
editCell.className = 'action-cell';
const editButton = document.createElement('button');
editButton.type = 'button';
editButton.className = 'btn secondary';
editButton.textContent = '編輯';
editButton.addEventListener('click', () => openEditModal(item));
editCell.appendChild(editButton);
row.appendChild(editCell);

// 修改後：
const actionCell = document.createElement('td');
actionCell.className = 'action-cell';

const editButton = document.createElement('button');
editButton.type = 'button';
editButton.className = 'btn secondary';
editButton.textContent = '編輯';
editButton.addEventListener('click', () => openEditModal(item));
actionCell.appendChild(editButton);

const deleteButton = document.createElement('button');
deleteButton.type = 'button';
deleteButton.className = 'btn danger';
deleteButton.textContent = '刪除';
deleteButton.style.marginLeft = '6px';
deleteButton.addEventListener('click', () => deleteSingleCase(item));
actionCell.appendChild(deleteButton);

row.appendChild(actionCell);
```

#### `deleteSingleCase` 實作：
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

### 2.3 CSS 佈局微調 (style.css)
將 `.action-cell` 寬度增加，以並排容納「編輯」與「刪除」兩個按鈕。

```css
.action-cell {
  width: 148px; /* 從原本的 92px 放大 */
}
```

---

## 3. 驗證計劃 (Verification Plan)

### 3.1 手動功能測試
1. **正常刪除流程**：
   - 開啟「追蹤中」分頁，選擇其中一筆個案（例如「測試員」）。
   - 點擊該列的「刪除」按鈕。
   - 預期：跳出瀏覽器確認框：「確認要永久刪除新朋友「測試員」的追蹤資料嗎？此操作將無法復原。」
   - 點擊「確認」。
   - 預期：顯示「刪除中...」，刪除成功後提示「已刪除追蹤中資料」，且列表重新載入，「測試員」已消失。
   - 檢查 Google Sheet，確認該列已被刪除。

2. **取消刪除流程**：
   - 點擊任何個案的「刪除」按鈕。
   - 在確認框中點擊「取消」。
   - 預期：確認框關閉，列表維持原樣，不發送任何 API 網路請求。
   - 檢查 Google Sheet，確認該個案資料依然存在。
