# 敬拜團服事公佈欄 (index.html) 整體視覺優化設計規格

此文件描述將服事公佈欄（`index.html`）升級為 Scheme B（自適應日程卡片牆）的整體視覺優化規格。

## 1. 視覺與佈局目標 (Visual & Layout Goals)
- **卡片自適應 (Responsive Grid Layout)**：
  - 移除首頁傳統表格展示，改用響應式日程卡片網格（`.dashboard-grid`）。
  - 桌面端與大螢幕下，卡片自動以雙欄/三欄排列，提升螢幕利用率；手機端自動變為單欄（100% 寬度）。
- **警示整合優化**：
  - 若聚會存在崗位衝突或同工請假（`row.hasWarning` 為真），卡片背景仍維持預設的白色毛玻璃質感，但直接在日期的右側展示醒目的黃橙色警語標籤（`.date-warning-badge`），以最小空間發揮最大視覺警示。
  - 服事同工如果為 `【待定】` 或空白，展示為高對比橘黃色待定徽章（`.badge-pending`），已排定同工展示為 Portal 藍徽章（`.badge-b`）。
- **資訊豐富度與標籤化**：
  - 曲目列表（逗號分隔字串）轉為獨立的綠字音樂標籤（`.song-badge-item`），使版面更有活力。
  - 卡片底部以專屬淡藍色裝飾框包裹「講道牧師、題目與經文」，讓會友與團員迅速掌握當週主題。

## 2. 變更詳情 (Proposed Changes)

### 2.1 apps/LKC_worship/style.css
- 新增卡片牆網格與日程卡片專用樣式（`.dashboard-grid`, `.dashboard-card`）。
- 定義 `.date-warning-badge`、`.song-badge-item` 樣式。
- 為 `body` 加上精美的全域漸層背景，加強毛玻璃立體感。

### 2.2 apps/LKC_worship/index.html
- 調整標題 `⛪ 敬拜團服事公佈欄` 及季度查詢小盒的樣式，引入磨砂玻璃框。
- 修改 `#dashboardContainer` 以支援無邊框的卡片牆。

### 2.3 apps/LKC_worship/script.js
- 將 `renderDashboardTable(data)` 重構為渲染卡片牆，將所有資料動態渲染為卡片結構並插入到 `#dashboardContainer` 中。

## 3. 驗證計劃 (Verification Plan)
- **行動端與桌面端自適應測試**：在瀏覽器中模擬手機及桌面版面，確認卡片牆排列整齊，沒有表格溢出。
- **警示標籤核對**：確認請假人員的警告標籤確實出現在日期右側，且背景為白色玻璃卡片，標籤顏色與卡片背景有明顯對比。
- **曲目與待定崗位測試**：確認逗號分隔的曲目完美標籤化，待定同工顯示為黃橙色 Badge。
