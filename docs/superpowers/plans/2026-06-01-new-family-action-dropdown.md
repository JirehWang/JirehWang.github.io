# 新家人系統 - 操作下拉選單實作計劃

> **面向 AI 代理的工作者：** 必需子技能：使用 superpowers:subagent-driven-development（推薦）或 superpowers:executing-plans 逐任務實現此計劃。步驟使用複選框（`- [ ]`）語法來跟蹤進度。

**目標：** 重構前端「追蹤中」表格，將操作欄的「編輯」與「刪除」收納至單一「操作」按鈕所觸發的浮動下拉選單中，並移除欄位名稱。

**架構：**
1. 修改 `style.css`：調整操作儲存格寬度為 `84px`，並加入絕對定位下拉選單樣式，使其浮動於表格上方，保證列高不被撐開。
2. 修改 `script.js`：
   - 全域新增一個 `click` 事件監聽器以收合所有下拉選單（Click Outside）。
   - 修改 `buildCaseTable`：移除操作欄表頭文字，改為渲染「操作」按鈕及包含編輯/刪除選單項的 `action-menu`，並綁定切換和單獨開啟的事件。
3. 驗證選單展開/關閉、編輯/刪除功能的可用性。

**技術棧：** Vanilla JavaScript, CSS

---

### 任務 1：修改前端樣式與佈局

**文件：**
- 修改：[style.css](file:///d:/program/Github/LKC1958_June_1.github.io/apps/LKC_NewFamily/style.css)

- [ ] **步驟 1：調整 `.action-cell` 寬度並新增下拉選單樣式**
  尋找 [style.css](file:///d:/program/Github/LKC1958_June_1.github.io/apps/LKC_NewFamily/style.css) 中的 `.action-cell`，將寬度修改為 `84px`，並在後面追加下拉選單、選單項目與浮動層的樣式代碼。

  **修改目標代碼 (約第 321 行)：**
  ```css
  .action-cell {
    width: 84px;
  }
  ```

  **新增代碼 (追加於 `.action-cell` 相關樣式下方，約第 330 行前)：**
  ```css
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

### 任務 2：重構前端互動邏輯

**文件：**
- 修改：[script.js](file:///d:/program/Github/LKC1958_June_1.github.io/apps/LKC_NewFamily/script.js)

- [ ] **步驟 1：全域註冊 Click Outside 收合事件監聽器**
  在 [script.js](file:///d:/program/Github/LKC1958_June_1.github.io/apps/LKC_NewFamily/script.js) 中加入全域點擊事件監聽器（例如在 `switchTab` 函式上方，約第 290 行）。

  **新增代碼：**
  ```javascript
  document.addEventListener('click', () => {
    document.querySelectorAll('.action-menu').forEach(menu => {
      menu.hidden = true;
    });
  });
  ```

- [ ] **步驟 2：重構 `buildCaseTable` 的按鈕渲染邏輯**
  尋找 [script.js](file:///d:/program/Github/LKC1958_June_1.github.io/apps/LKC_NewFamily/script.js) 中的 `buildCaseTable` 函式，將原本表頭的「操作」欄位標題改為空，並修改 `actionCell` 的按鈕渲染，將「編輯」與「刪除」包裝至單一按鈕觸發的下拉選單容器中。

  **修改目標代碼 1 (表頭欄位名稱變更，約第 1105-1108 行)：**
  ```javascript
  headRow.innerHTML = `${selectable ? '<th class="check-cell">結案</th><th class="action-cell"></th>' : ''}${columns.map(column => `<th>${escapeHtml(getColumnLabel(column))}</th>`).join('')}`;
  ```

  **修改目標代碼 2 (儲存格內按鈕渲染變更，約第 1118-1135 行)：**
  ```javascript
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

      // 選單內容容器
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

      // 點擊事件：開啟此列選單，收合其他列
      toggleButton.addEventListener('click', event => {
        event.stopPropagation();
        document.querySelectorAll('.action-menu').forEach(m => {
          if (m !== menu) m.hidden = true;
        });
        menu.hidden = !menu.hidden;
      });
  ```

---

### 任務 3：驗證、提交與推送

- [ ] **步驟 1：本地與手動驗證**
  在瀏覽器中手動操作表格列的「操作」下拉按鈕，確認是否僅顯示一個按鈕、點擊後正確彈出包含「編輯」與「刪除」的浮動選單，點擊空白處自動收合，且點擊選項功能正常。

- [ ] **步驟 2：Commit 並推送至 GitHub 倉庫**
  運行命令：
  ```powershell
  git add apps/LKC_NewFamily/style.css apps/LKC_NewFamily/script.js docs/superpowers/plans/2026-06-01-new-family-action-dropdown.md
  git commit -m "feat(new-family): wrap edit and delete buttons in single action dropdown menu"
  git push
  ```
