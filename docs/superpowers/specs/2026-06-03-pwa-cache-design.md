# 設計文檔：PWA Manifest 與 Service Worker 離線快取機制

- **日期**：2026-06-03
- **狀態**：已批准 (Approved)
- **專案**：LKC 教會管理系統 (GitHub Pages)

## 1. 目的
為 LKC 教會管理系統導入 PWA (Progressive Web App) 離線快取機制，提供離線可用性與加快二次載入速度，同時確保動態資料 API (Google Apps Script 與 Firebase RTDB) 繞過快取以保持即時性。

## 2. 架構與配置

### 2.1 Web App Manifest (`manifest.json`)
在專案根目錄下建立 `manifest.json`，配置如下：
- `name`: "LKC 教會管理系統"
- `short_name`: "LKC系統"
- `start_url`: "/LKC1958_June_1.github.io/index.html"
- `display`: "standalone"
- `orientation`: "portrait"
- `background_color`: "#1a202c"
- `theme_color`: "#667eea"
- `icons`: 包含 192x192 及 512x512 圖示（指向 `https://jirehwang.github.io/LKC1958_June_1.github.io/docs/images/`）

### 2.2 HTML 引入
在以下主要子應用的 HTML `<head>` 中引入 manifest：
- `apps/LKC_Group/index.html`
- `apps/LKC_SundayserviceAttendance/index.html`
- `apps/LKC_worship/admin.html`

代碼：
```html
<link rel="manifest" href="/LKC1958_June_1.github.io/manifest.json">
```

### 2.3 Service Worker (`service-worker.js`)
在專案根目錄建立 `service-worker.js`，包含：
- **預快取資源 (Precaching)**: 快取 `index.html`, `config.js`, `manifest.json`。
- **快取策略**: `Stale-While-Revalidate`（快取優先，背景拉取更新）。
- **安全防護 (不快取)**:
  - 繞過所有 `POST` 請求。
  - 繞過網址包含 `script.google.com` (GAS) 與 `firebasedatabase.app` (Firebase) 的請求。

### 2.4 自動註冊 (`config.js`)
在 `config.js` 的立即執行函數 (IIFE) 尾端加入 Service Worker 註冊代碼，並動態判定環境：
- GitHub Pages 環境：Scope 設為 `/LKC1958_June_1.github.io/`，SW 路徑為 `/LKC1958_June_1.github.io/service-worker.js`。
- 本地開發環境：Scope 設為 `/`，SW 路徑為 `/service-worker.js`。

## 3. 測試與驗證計畫
1. 靜態語法檢查：確保 JS 及 JSON 語法無誤。
2. 檢查 Scope 設定：確保在 GitHub Pages 與 localhost 環境下註冊成功。
3. 快取驗證：在瀏覽器 Application 面板檢查 Cache Storage 中是否成功快取靜態資源，並確認 GAS/Firebase 請求皆繞過快取（顯示為 Network 請求）。
