# 敬拜曲目管理頁面 (worship_songs.html) 視覺與佈局優化實現計劃

> **面向 AI 代理的工作者：** 必需子技能：使用 superpowers:subagent-driven-development（推薦）或 superpowers:executing-plans 逐任務實現此計劃。步驟使用複選框（`- [ ]`）語法來跟踪進度。

**目標：** 將 `worship_songs.html` 的表格版面替換為玻璃擬物卡片流（Scheme A），並實現高對比度、與 `admin.html` 一致的優質視覺樣式。

**建築：** 將 HTML 中的 table 替換為 cards-grid 的 div，重構 JS 渲染邏輯，生成支援響應式 Grid 的 `.song-card` DOM，並處理編輯與取消事件。

**技術棧：** HTML5, Vanilla CSS (with Backdrop Filter), Javascript, Bootstrap 5 (Layout utilities only).

---

### 任務 1：重構 HTML 結構與專屬樣式

**文件：**
- 修改：`apps/LKC_worship/worship_songs.html`

- [ ] **步驟 1：修改 style.css 與 html 中的樣式以支援卡片流與發光發亮效果**
  我們需要修改 `apps/LKC_worship/worship_songs.html` 的內置樣式，移除 table 相關 CSS，並加入卡片流（`.cards-grid`, `.song-card`）等樣式。
  
- [ ] **步驟 2：替換 HTML 的 Table 結構為卡片容器**
  將 `worship_songs.html` 的 `<div class="table-scroll-wrapper" id="tableWrapper">...</table></div>` 替換為：
  ```html
  <div class="table-scroll-wrapper" id="tableWrapper" style="display:none; border:none; background:transparent; box-shadow:none;">
    <div class="cards-grid" id="songsCardsContainer"></div>
  </div>
  ```

- [ ] **步驟 3：手動檢查 HTML 的結構**
  確認 `worship_songs.html` 中的容器 ID 正確無誤，並且無語法錯誤。

- [ ] **步驟 4：Commit**
  ```bash
  git add apps/LKC_worship/worship_songs.html
  git commit -m "style: restructure worship_songs.html for Scheme A card layout"
  ```

---

### 任務 2：重構 Javascript 渲染與交互邏輯

**文件：**
- 修改：`apps/LKC_worship/worship_songs.js`

- [ ] **步驟 1：修改 renderTable 函數為 renderCards 函數**
  將 `worship_songs.js` 的 `renderTable` 修改為 `renderCards`，動態生成 `.song-card` HTML 片段：
  - 日期與星期。
  - 聚會名稱與類別標籤（高對比 `#eaf5ee` 底色 + `#006030` 文字）。
  - 主領人標籤（高對比 `#e6f0f7` 底色 + `#30759f` 文字）。
  - 敬拜曲目列表（`.song-item`，綠字底色透明）。
  - 編輯狀態隱藏 `songs-display` 並顯示發光的 `textarea` 及提示文字。

- [ ] **步驟 2：修改 editRow、saveRow、cancelRow 函數以適配卡片結構**
  由於沒有 Table row，原本獲取 `tr` 的 `row-${idx}` 應改為 `card-${idx}`。
  編輯與取消的 DOM 查找需要適配新的卡片結構。

- [ ] **步驟 3：修正 loadSongs 呼叫 renderTable 的地方**
  將 `loadSongs` 內的 `renderTable()` 改成 `renderCards()`。

- [ ] **步驟 4：靜態程式碼檢查與無錯運行驗證**
  確保沒有引用錯誤或語法錯誤。

- [ ] **步驟 5：Commit**
  ```bash
  git add apps/LKC_worship/worship_songs.js
  git commit -m "feat: refactor worship_songs.js to render cards instead of table"
  ```

---

### 任務 3：全站與行動端視覺自檢驗收

- [ ] **步驟 1：在瀏覽器載入 worship_songs.html 驗收效果**
  開啟本機瀏覽器，手動測試讀取 2026-Q2 季度，確保：
  - 卡片呈現出漂亮的磨砂玻璃質感。
  - 標籤和文字的可讀性大幅提升。
  - 點擊「編輯」可順利打開輸入框且焦點高亮。
  - 修改後「儲存」或「暫存」功能完全正常，API 返回成功。

- [ ] **步驟 2：行動端尺寸模擬與響應式核對**
  在 Chrome 開發者工具中開啟行動端模擬器（如 iPhone SE / iPhone 12 Pro 尺寸），確保卡片寬度自適應 100%，對比度充足，且「儲存所有曲目」按鈕在右下角位置正確且不會擋住主要曲目內容。

- [ ] **步驟 3：Commit 並推送到 GitHub 遠端儲存庫**
  ```bash
  git push origin main
  ```
