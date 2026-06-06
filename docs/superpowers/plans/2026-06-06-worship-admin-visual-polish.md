# 敬拜團服事管理系統 — 網頁美編優化實現計劃

> **面向 AI 代理的工作者：** 必需子技能：使用 superpowers:subagent-driven-development（推薦）或 superpowers:executing-plans 逐任務實現此計劃。步驟使用複選框（`- [ ]`）語法來跟踪進度。

**目標：** 依據 Option A 現代微光摩登風 (Modern Glassmorphic) 優化管理後台的網頁美編，同時保持原有 LKC 綠 (#006030) 與 Portal 藍 (#30759f) 色系。

**架構：** 本次修改集中於前端樣式，更新 `style.css` 中的變數與 CSS 類別（含毛玻璃卡片、漸層按鈕、新字體、陰影等效果），並小幅調整 `admin.html` 頭部以載入 Google Fonts。

**技術棧：** HTML5, CSS3, Google Fonts (Inter & Noto Sans TC), Bootstrap 5

---

### 任務 1：字型集成與 HTML 設定

**文件：**
- 修改：`apps/LKC_worship/admin.html:1-15`

- [ ] **步驟 1：修改 `admin.html` 以導入 Google Fonts**
  在 `<head>` 中引入 Google Fonts 連結，放在 Bootstrap CSS 之前或之後。
  
  ```html
  <head>
    <meta charset="UTF-8">
    <title>敬拜團服事排班系統 - 管理後台</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Noto+Sans+TC:wght@400;500;700&display=swap" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
  ```

- [ ] **步驟 2：運行本地伺服器手動驗證**
  在 `d:\program\Github\LKC1958_June_1.github.io` 啟動一個 Python 輕量級伺服器：
  `python -m http.server 8000`
  打開瀏覽器造訪：`http://localhost:8000/apps/LKC_worship/admin.html`
  確認頁面無載入錯誤。

- [ ] **步驟 3：Commit**
  ```bash
  git add apps/LKC_worship/admin.html
  git commit -m "style: import google fonts in admin.html"
  ```

---

### 任務 2：CSS 核心設計系統與容器重構

**文件：**
- 修改：`apps/LKC_worship/style.css:1-20`
- 修改：`apps/LKC_worship/style.css:124-210`

- [ ] **步驟 1：於 `style.css` 開頭引入字型與宣告設計系統 CSS 變數**
  更新 `body` 的 `background-color` 與 `font-family`。
  
  ```css
  /* style.css - 敬拜團服事管理系統 (視覺優化版) */
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Noto+Sans+TC:wght@400;500;700&display=swap');
  
  :root {
      --lkc-green: #006030;
      --portal-blue: #30759f;
      --bg-gradient: linear-gradient(135deg, #006030 0%, #30759f 100%);
      --card-bg: rgba(255, 255, 255, 0.85);
      --border-subtle: rgba(0, 96, 48, 0.08);
      --shadow-subtle: 0 8px 32px 0 rgba(0, 96, 48, 0.04);
  }
  
  /* --- 基礎樣式 --- */
  body { 
      background-color: #f4f7f6 !important; 
      font-family: 'Inter', 'Noto Sans TC', -apple-system, BlinkMacSystemFont, "Microsoft JhengHei", sans-serif !important; 
      padding: 15px; 
      -webkit-text-size-adjust: 100%;
  }
  
  /* 容器樣式 */
  .container, .main-card { 
      background: var(--card-bg) !important;
      backdrop-filter: blur(12px);
      -webkit-backdrop-filter: blur(12px);
      padding: 30px !important; 
      border-radius: 16px !important; 
      border: 1px solid var(--border-subtle) !important;
      box-shadow: var(--shadow-subtle) !important; 
      max-width: 1400px; 
      margin: auto; 
  }
  ```

- [ ] **步驟 2：更新導覽列、卡片、按鈕樣式**
  更新 `style.css` 中第 124-210 行左右的樣式設定，升級卡片圓角、陰影與按鈕漸層 Hover 動態效果。
  
  ```css
  /* Bootstrap overrides for LKC Green and Portal Blue theme */
  .text-primary {
    color: var(--lkc-green) !important;
  }
  .border-primary {
    border-color: var(--lkc-green) !important;
  }
  .card {
    background: rgba(255, 255, 255, 0.9) !important;
    border: 1px solid var(--border-subtle) !important;
    border-radius: 12px !important;
    box-shadow: 0 4px 12px rgba(0, 96, 48, 0.02) !important;
  }
  .btn-primary, .btn-success {
    background: var(--bg-gradient) !important;
    border: none !important;
    color: white !important;
    font-weight: 600 !important;
    box-shadow: 0 4px 12px rgba(0, 96, 48, 0.15) !important;
  }
  .btn-primary:hover, .btn-success:hover {
    transform: translateY(-2px);
    box-shadow: 0 6px 18px rgba(0, 96, 48, 0.25) !important;
  }
  .btn-outline-primary {
    color: var(--lkc-green) !important;
    border: 1.5px solid var(--lkc-green) !important;
  }
  .btn-outline-primary:hover {
    background: var(--bg-gradient) !important;
    color: white !important;
    border-color: transparent !important;
  }
  .bg-success {
    background-color: var(--lkc-green) !important;
  }
  .text-success {
    color: var(--lkc-green) !important;
  }
  .bg-primary {
    background-color: var(--lkc-green) !important;
  }
  .form-select, .form-control {
    border: 1.5px solid rgba(0, 96, 48, 0.15) !important;
    border-radius: 8px !important;
  }
  .form-select:focus, .form-control:focus {
    border-color: var(--portal-blue) !important;
    box-shadow: 0 0 0 3px rgba(48, 117, 159, 0.15) !important;
  }
  .nav-pills .nav-link.active, .nav-pills .show > .nav-link {
    background: var(--bg-gradient) !important;
    color: white !important;
    box-shadow: 0 4px 12px rgba(0, 96, 48, 0.18) !important;
  }
  .nav-link {
    color: var(--portal-blue) !important;
  }
  .nav-link:hover {
    color: var(--lkc-green) !important;
  }
  ```

- [ ] **步驟 3：手動網頁驗證**
  重整網頁 `http://localhost:8000/apps/LKC_worship/admin.html`。
  確認卡片及主按鈕已成功呈現毛玻璃與漸層效果，焦點輸入框顯示藍色光暈。

- [ ] **步驟 4：Commit**
  ```bash
  git add apps/LKC_worship/style.css
  git commit -m "style: apply core glassmorphic containers and dynamic button gradient styles"
  ```

---

### 任務 3：狀態標籤與表格美化

**文件：**
- 修改：`apps/LKC_worship/style.css:50-104`
- 修改：`apps/LKC_worship/style.css:295-353`

- [ ] **步驟 1：美化表格設計與凍結欄位陰影**
  修改表格相關樣式，優化 `.modern-table`，使其表頭具有漸層色彩，並將凍結欄位邊界加入柔和陰影。
  
  ```css
  /* --- 🌟 表格滾動與凍結窗格 --- */
  .table-scroll-container { 
      background: white; 
      border-radius: 12px; 
      overflow-x: auto; 
      border: 1px solid var(--border-subtle); 
      position: relative; 
      -webkit-overflow-scrolling: touch;
      box-shadow: 0 4px 20px rgba(0, 0, 0, 0.02);
  }
  
  .modern-table, .bulletin-table { 
      width: max-content; min-width: 100%; border-collapse: separate; border-spacing: 0; 
  }
  
  /* 表頭樣式 */
  .modern-table th, .bulletin-table th { 
      background: linear-gradient(180deg, #006030 0%, #004d26 100%) !important;
      color: white; 
      text-align: center; 
      padding: 15px 10px; 
      white-space: nowrap; 
      position: sticky; 
      top: 0; 
      z-index: 10; 
      border-bottom: 2px solid rgba(0, 96, 48, 0.15) !important;
  }
  
  /* 單元格通用樣式 */
  .modern-table td, .bulletin-table td { 
      vertical-align: middle; 
      text-align: center; 
      padding: 12px 10px; 
      border-bottom: 1px solid rgba(0, 96, 48, 0.04) !important;
      max-width: 180px; 
      white-space: normal; 
      word-break: break-all; 
      line-height: 1.5; 
      font-size: 0.95rem;
      background: rgba(255, 255, 255, 0.7);
      transition: background 0.2s ease;
  }
  
  .modern-table tr:hover td, .bulletin-table tr:hover td {
      background: rgba(0, 96, 48, 0.02) !important;
  }
  
  /* 🌟 修正：凍結第一欄（日期） */
  .modern-table th:nth-child(1), .modern-table td:nth-child(1),
  .bulletin-table th:nth-child(1), .bulletin-table td:nth-child(1) { 
      position: sticky; 
      left: 0; 
      width: 100px;
      z-index: 5; 
      background: rgba(255, 255, 255, 0.9) !important; 
      white-space: nowrap !important;
      padding-left: 8px !important;
      padding-right: 8px !important;
      font-size: 0.9rem;
      border-right: 1px solid rgba(0, 96, 48, 0.05) !important;
  }
  
  /* 🌟 修正：凍結第二欄（聚會類別） */
  .modern-table th:nth-child(2), .modern-table td:nth-child(2),
  .bulletin-table th:nth-child(2), .bulletin-table td:nth-child(2) { 
      position: sticky; 
      left: 100px;
      width: 90px; 
      z-index: 5; 
      background: rgba(255, 255, 255, 0.9) !important; 
      box-shadow: 4px 0 10px rgba(0, 96, 48, 0.04) !important; 
      white-space: nowrap !important;
      border-right: 1.5px solid rgba(0, 96, 48, 0.08) !important;
  }
  ```

- [ ] **步驟 2：美化狀態標籤與徽章 (Badges)**
  修改 `.badge.bg-dark`、正式與實習同工的狀態標記。
  
  ```css
  /* --- 標籤與徽章 (Badges) --- */
  .badge.bg-dark {
    background: linear-gradient(135deg, #1e3a2f 0%, #006030 100%) !important;
    border-radius: 99px !important;
    font-weight: bold !important;
    box-shadow: 0 2px 8px rgba(0, 96, 48, 0.15) !important;
  }
  
  /* 實習狀態 */
  .badge.bg-warning, .text-warning {
    background-color: rgba(180, 83, 9, 0.08) !important;
    color: #b45309 !important;
    border: 1px solid rgba(180, 83, 9, 0.15) !important;
    border-radius: 99px !important;
    padding: 4px 12px !important;
    font-weight: 600 !important;
  }
  
  /* 正式狀態 */
  .badge.bg-primary, .text-primary {
    background-color: rgba(19, 115, 51, 0.08) !important;
    color: #137333 !important;
    border: 1px solid rgba(19, 115, 51, 0.15) !important;
    border-radius: 99px !important;
    padding: 4px 12px !important;
    font-weight: 600 !important;
  }
  ```

- [ ] **步驟 3：手動網頁驗證**
  在 `http://localhost:8000/apps/LKC_worship/admin.html` 切換至各個分頁。
  確認表格表頭漸層自然、同工名單處的「正式 / 實習」狀態標籤圓角好看且底色柔和，排版無錯位。

- [ ] **步驟 4：Commit**
  ```bash
  git add apps/LKC_worship/style.css
  git commit -m "style: update tables layout with gradient header and polished status badges"
  ```

---

## 驗證計劃

### 自動化測試
本項目為靜態網頁搭配 Google Apps Script 的純前端頁面，無配置本地前端測試框架（例如 Jest/Vitest），因而使用人工視覺與互動驗證。

### 手動驗證
1. 啟動本機伺服器：`python -m http.server 8000`。
2. 造訪管理後台：`http://localhost:8000/apps/LKC_worship/admin.html`。
3. 檢查以下項目：
   - 頁面字體是否加載並切換為 `Inter` 與 `Noto Sans TC`（檢查開發者工具的 Network 與 Computed 樣式）。
   - 切換 Tab Pill，確認活動頁籤的漸層與投影效果正常，無跑版。
   - 點擊「👥 敬拜團員名單」與「⚙️ 位置與同工」頁面，確認狀態標記與輸入框正常。
   - 點擊「🗓️ 服事安排預覽」，讀取季度班表後，確認凍結的「日期」與「類別」欄位黏滯正常，且具有精緻邊框與邊緣投影。
