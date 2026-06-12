# 服事公佈欄首頁 (index.html) 整體視覺優化實現計劃

> **面向 AI 代理的工作者：** 必需子技能：使用 superpowers:subagent-driven-development（推薦）或 superpowers:executing-plans 逐任務實現此計劃。步驟使用複選框（`- [ ]`）語法來跟踪進度。

**目標：** 將公佈欄首頁表格改為磨砂玻璃日程卡片牆（Scheme B），並加入日期右側高對比警示標籤、同工崗位徽章與曲目標籤化。

**建築：** 
- 在 `style.css` 新增 `.dashboard-grid` 與 `.dashboard-card` 樣式。
- 在 `index.html` 微調標題與查詢小盒。
- 在 `script.js` 重構 `renderDashboardTable` 函數以動態渲染卡片結構。

**技術棧：** HTML5, Vanilla CSS (with Backdrop Filter), Javascript, Bootstrap 5.

---

### 任務 1：新增 CSS 樣式支援

**文件：**
- 修改：`apps/LKC_worship/style.css`

- [ ] **步驟 1：調整全域背景漸層**
  將 `body` 的 `background-color` 改為漸層色：
  ```css
  body { 
      background: radial-gradient(circle at 100% 0%, rgba(48, 117, 159, 0.08) 0%, rgba(0, 96, 48, 0.03) 100%), #fafbfa !important; 
      ...
  }
  ```

- [ ] **步驟 2：新增卡片牆與日程卡片樣式**
  新增 `.dashboard-grid` 網格佈局，與 `.dashboard-card` 白色玻璃卡片樣式。
  
- [ ] **步驟 3：新增警語標籤與同工曲目標記樣式**
  新增 `.date-warning-badge` 黃橙色警語標籤，與 `.song-badge-item` 曲目標籤。
  
- [ ] **步驟 4：靜態 CSS 代碼確認**
  確認沒有重複的 Class 定義，且語法完全正確。

- [ ] **步驟 5：Commit**
  ```bash
  git add apps/LKC_worship/style.css
  git commit -m "style: add responsive grid cards and warning badge styles for dashboard"
  ```

---

### 任務 2：修改 index.html 佈局結構

**文件：**
- 修改：`apps/LKC_worship/index.html`

- [ ] **步驟 1：美化標題與季度選單區塊**
  將 `.quarter-box` 樣式升級為毛玻璃質感。
  
- [ ] **步驟 2：調整 #dashboardContainer 容器包裝**
  將原本 Table 外層的 `.table-scroll-container` 包裝去除邊框與底色，改為無邊框以適配卡片牆。

- [ ] **步驟 3：Commit**
  ```bash
  git add apps/LKC_worship/index.html
  git commit -m "style: polish index.html layout and glassmorphic header widget"
  ```

---

### 任務 3：重構 script.js 渲染邏輯

**文件：**
- 修改：`apps/LKC_worship/script.js`

- [ ] **步驟 1：重構 renderDashboardTable 函數**
  將原本輸出 `<table>` 的邏輯重構為輸出 `<div class="dashboard-grid">`，並在其中生成 `.dashboard-card`：
  - 日期與星期，並判斷若 `row.hasWarning` 為真，在日期右側添加 `<span class="date-warning-badge">⚠️ ${row.warningMessage}</span>`。
  - 同工崗位徽章化：動態職責欄位值為 `【待定】` 則生成 `.badge-pending` 徽章，已排定則生成 `.badge-b` 徽章。
  - 敬拜曲目徽章化：將逗號分隔字串拆分生成獨立的音樂小標籤 `.song-badge-item`。
  - 講道資訊區：在卡片底部以專屬帶有微光藍色的 `.sermon-box` 包裹「講道牧師、題目、經文」。

- [ ] **步驟 2：確認程式碼中動態欄位解析正確**
  確認 `finalHeaders`、`fixedHeaders` 與崗位資訊匹配無誤。

- [ ] **步驟 3：Commit**
  ```bash
  git add apps/LKC_worship/script.js
  git commit -m "feat: refactor script.js rendering logic to support Scheme B cards"
  ```

---

### 任務 4：全站自適應驗收與推行

- [ ] **步驟 1：本地瀏覽器測試驗收**
  讀取 2026-Q2 資料，核對：
  - 首頁大表格是否完全轉為響應式卡片牆。
  - 卡片呈現白色半透明毛玻璃質感。
  - 請假警告是否正確且有高對比地顯示在日期右側，與卡片底色有明顯區別。
  - 同工名字與待定、曲目是否都成功徽章化。

- [ ] **步驟 2：行動端尺寸模擬**
  使用 Chrome 行動端模擬器核實單欄呈現。

- [ ] **步驟 3：推送到 GitHub 遠端 main 分支**
  ```bash
  git push origin main
  ```
