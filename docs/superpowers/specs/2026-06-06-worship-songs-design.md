# 敬拜曲目管理頁面 (worship_songs.html) 視覺與佈局優化設計規格

此文件描述將 `worship_songs.html` 升級為玻璃擬物卡片流方案（Scheme A）的設計與實作規格。

## 1. 視覺與佈局目標 (Visual & Layout Goals)
- **色系與主題**：與 `admin.html` 完美契合，採用 Google Fonts Inter + Noto Sans TC 作為基礎字型，並在主體背景下套用毛玻璃質感與 LKC 綠藍漸層元素。
- **對比度與易讀性**：
  - 日期與聚會名稱加粗加大。
  - 綠色類別標籤採用明亮的 `#eaf5ee` 搭配 `#006030` 深綠色文字。
  - 藍色主領標籤採用 `#e6f0f7` 搭配 `#30759f` 藍色文字。
  - 手機端移除複雜的表格橫向滾動，改用單欄卡片流；桌面端以雙欄或三欄響應式網格（Grid）排列卡片。

## 2. 變更詳情 (Proposed Changes)

### 2.1 apps/LKC_worship/worship_songs.html
- **樣式重構**：
  - 移除原有的內嵌舊式 Table CSS。
  - 新增符合 Glassmorphism 的卡片樣式、漸層表頭、查詢列磨砂效果、以及 `textarea` 發光聚焦效果。
- **DOM 結構調整**：
  - 將 `<div class="table-scroll-wrapper" id="tableWrapper">` 表格包裝層改為 `<div class="cards-grid" id="cardsGrid">` 的卡片網格容器。
  - 移除原本 Table 表頭與表身的靜態 HTML。

### 2.2 apps/LKC_worship/worship_songs.js
- **渲染邏輯重構**：
  - 將 `renderTable` 重新命名並重構為 `renderCards`。
  - 產出的 DOM 結構為卡片（`.song-card`），每張卡片內置有「日期與標籤列」、「曲目預覽/編輯框」、「操作按鈕列」。
- **事件與操作**：
  - `editRow(idx)`：顯示對應卡片中的 `textarea` 及輸入提示，隱藏原本的 `songs-display` 曲目列表。將「編輯」按鈕替換為「暫存」與「取消」。
  - `saveRow(idx)`：提取 `textarea` 內容，過濾格式化並寫入 `songsData[idx]['敬拜曲目']`。重新以 `.song-item` 格式渲染列表並顯示。
  - `cancelRow(idx)`：還原 `textarea` 內容，隱藏編輯區並重新顯示原曲目列表。
  - `saveAllSongs()`：保存所有曲目時，檢查並完成所有正在進行中的編輯卡片，並顯示儲存中動畫。

## 3. 規格自檢 (Verification Plan)
- **手動驗證**：
  - 行動端響應式排版測試：在手機解析度下，卡片流應呈現為單欄，卡片寬度自適應 100%，對比度良好且不會產生橫向溢出。
  - 編輯狀態測試：點擊任何卡片的編輯按鈕後，輸入框正確發光，暫存後曲目渲染格式正確（逗號斜線自動轉換為頓號並折行呈現）。
  - 儲存全部測試：點擊右下角儲存按鈕，成功串接後端 API 並更新同步時間。
