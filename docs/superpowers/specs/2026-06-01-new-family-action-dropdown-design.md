# 新家人系統 - 操作下拉選單功能設計規格書 (New Family Action Dropdown Spec)

本規格書設計將「追蹤中」表格的操作介面進行精簡。每一列個案將只會呈現一個「操作」按鈕。點擊該按鈕時，會於按鈕下方彈出包含「編輯」與「刪除」選項的浮動下拉選單。此外，表頭的「操作」文字欄位名稱將被移除以使版面更為乾淨。

## 1. 變更範圍

本變更僅影響前端網頁的 HTML 表格結構、JavaScript 互動邏輯與 CSS 樣式，不涉及 GAS 後端與資料庫。

*   **修改檔案**：
    *   [script.js](file:///d:/program/Github/LKC1958_June_1.github.io/apps/LKC_NewFamily/script.js) (前端按鈕渲染邏輯與事件綁定)
    *   [style.css](file:///d:/program/Github/LKC1958_June_1.github.io/apps/LKC_NewFamily/style.css) (新增下拉選單的浮動樣式)

---

## 2. 詳細設計

### 2.1 表格結構修改 (script.js)
1. 表頭：修改 `buildCaseTable` 的 `<th class="action-cell">操作</th>` 為 `<th class="action-cell"></th>`，隱藏欄位名稱。
2. 表列：改為渲染一個包含「操作」按鈕與隱藏選單（包含「編輯」與「刪除」）的容器。
3. 互動邏輯：
   - 點擊「操作」按鈕時切換該選單的顯示狀態，並同時收合其他列的選單。
   - 點擊頁面其他任何地方時，自動隱藏所有選單（Click Outside 收合）。

#### `buildCaseTable` 修改預覽：
```javascript
// 修改表頭 (移除欄位標題)
headRow.innerHTML = `${selectable ? '<th class="check-cell">結案</th><th class="action-cell"></th>' : ''}${columns.map(column => `<th>${escapeHtml(getColumnLabel(column))}</th>`).join('')}`;

// ...

// 修改資料列按鈕渲染 (改為單一下拉按鈕)
if (selectable) {
  const checkboxCell = document.createElement('td');
  checkboxCell.className = 'check-cell';
  checkboxCell.innerHTML = `<input type="checkbox" value="${item.rowNumber}" aria-label="勾選 ${escapeHtml(item['新家人姓名'] || '此筆資料')} 結案">`;
  row.appendChild(checkboxCell);

  const actionCell = document.createElement('td');
  actionCell.className = 'action-cell';
  
  // 下拉選單容器
  const dropdown = document.createElement('div');
  dropdown.className = 'action-dropdown';

  // 主要操作按鈕
  const toggleButton = document.createElement('button');
  toggleButton.type = 'button';
  toggleButton.className = 'btn secondary action-toggle-btn';
  toggleButton.textContent = '操作';
  dropdown.appendChild(toggleButton);

  // 選單選單內容
  const menu = document.createElement('div');
  menu.className = 'action-menu';
  menu.hidden = true;

  // 編輯選項
  const editItem = document.createElement('button');
  editItem.type = 'button';
  editItem.className = 'menu-item';
  editItem.textContent = '編輯';
  editItem.addEventListener('click', () => {
    menu.hidden = true;
    openEditModal(item);
  });
  menu.appendChild(editItem);

  // 刪除選項
  const deleteItem = document.createElement('button');
  deleteItem.type = 'button';
  deleteItem.className = 'menu-item danger';
  deleteItem.textContent = '刪除';
  deleteItem.addEventListener('click', () => {
    menu.hidden = true;
    deleteSingleCase(item);
  });
  menu.appendChild(deleteItem);

  dropdown.appendChild(menu);
  actionCell.appendChild(dropdown);
  row.appendChild(actionCell);

  // 點擊事件：切換本列選單並收合其他列
  toggleButton.addEventListener('click', event => {
    event.stopPropagation();
    document.querySelectorAll('.action-menu').forEach(m => {
      if (m !== menu) m.hidden = true;
    });
    menu.hidden = !menu.hidden;
  });
}
```

#### 全域點擊收合監聽器 (script.js)：
```javascript
// 在 script.js 頂部事件監聽器區塊新增
document.addEventListener('click', () => {
  document.querySelectorAll('.action-menu').forEach(menu => {
    menu.hidden = true;
  });
});
```

---

### 2.2 樣式設計 (style.css)
寬度調整回緊湊寬度，並加入絕對定位樣式確保選單浮動顯示。

```css
/* 調整欄位寬度 */
.action-cell {
  width: 84px;
}

/* 下拉選單容器 */
.action-dropdown {
  position: relative;
  display: inline-block;
}

.action-toggle-btn {
  width: 100%;
}

/* 浮動選單主要樣式 */
.action-menu {
  position: absolute;
  top: 100%;
  left: 0;
  z-index: 100;
  min-width: 90px;
  margin-top: 4px;
  background: var(--surface);
  border: 1px solid var(--line);
  border-radius: 6px;
  box-shadow: 0 4px 12px rgba(20, 31, 28, 0.12);
  display: flex;
  flex-direction: column;
  overflow: hidden;
}

.action-menu[hidden] {
  display: none;
}

/* 選單按鈕樣式 */
.action-menu .menu-item {
  border: 0;
  background: transparent;
  padding: 8px 12px;
  text-align: center;
  font-size: 13px;
  font-weight: 700;
  color: var(--text);
  cursor: pointer;
  width: 100%;
  border-radius: 0;
  min-height: auto;
}

.action-menu .menu-item:hover {
  background: var(--surface-2);
  color: var(--primary);
}

.action-menu .menu-item.danger {
  color: var(--danger);
}

.action-menu .menu-item.danger:hover {
  background: #fff5f5;
  color: var(--danger);
}
```

---

## 3. 驗證計劃 (Verification Plan)

1. **選單展開/收合驗證**：
   - 點擊任一列的「操作」按鈕。
   - 預期：按鈕下方出現含有「編輯」與「刪除」的浮動選單，且表格行高與寬度均無任何拉伸或變形。
2. **Click Outside 收合驗證**：
   - 展開某列選單後，點擊網頁空白處或其它欄位。
   - 預期：該選單自動隱藏。
3. **單一開啟驗證**：
   - 展開第 1 列選單後，隨即點擊第 2 列的「操作」按鈕。
   - 預期：第 1 列選單自動關閉，並同時顯示第 2 列選單。
4. **功能正確性驗證**：
   - 點擊選單內的「編輯」：選單隱藏，並彈出編輯個案 Modal。
   - 點擊選單內的「刪除」：選單隱藏，跳出確認視窗，點擊確定後資料成功刪除且列表重新整理。
