# PWA Manifest 與 Service Worker 離線快取機制 實現計畫

> **面向 AI 代理的工作者：** 必需子技能：使用 superpowers:subagent-driven-development（推薦）或 superpowers:executing-plans 逐任務實現此計畫。步驟使用複選框（`- [ ]`）語法來跟踪進度。

**目標：** 為 LKC 教會管理系統導入 PWA Manifest 與 Service Worker 離線快取。

**建築：** 採用 Stale-While-Revalidate 快取策略，並針對 GAS 與 Firebase 請求進行安全防護（繞過快取）。在 config.js 中自動註冊 SW 並動態判定 Scope。

**技術棧：** Vanilla JavaScript, Web App Manifest, Service Worker API.

---

## 任務 1：建立 PWA Manifest 配置

**文件：**
- 創建：`manifest.json`

- [ ] **步驟 1：建立 `manifest.json` 檔案**
  內容：
  ```json
  {
    "short_name": "LKC系統",
    "name": "LKC 教會管理系統",
    "icons": [
      {
        "src": "https://jirehwang.github.io/LKC1958_June_1.github.io/docs/images/logo_192.png",
        "type": "image/png",
        "sizes": "192x192"
      },
      {
        "src": "https://jirehwang.github.io/LKC1958_June_1.github.io/docs/images/logo_512.png",
        "type": "image/png",
        "sizes": "512x512"
      }
    ],
    "start_url": "/LKC1958_June_1.github.io/index.html",
    "background_color": "#1a202c",
    "theme_color": "#667eea",
    "display": "standalone",
    "orientation": "portrait"
  }
  ```
- [ ] **步驟 2：驗證 JSON 格式**
  確認 JSON 語法無誤。
- [ ] **步驟 3：Commit**
  ```bash
  git add manifest.json
  git commit -m "feat: add PWA manifest.json"
  ```

---

## 任務 2：在 HTML 檔案中引用 manifest

**文件：**
- 修改：`apps/LKC_Group/index.html`
- 修改：`apps/LKC_SundayserviceAttendance/index.html`
- 修改：`apps/LKC_worship/admin.html`

- [ ] **步驟 1：修改 `apps/LKC_Group/index.html`**
  在 `<head>` 中插入：
  ```html
  <link rel="manifest" href="/LKC1958_June_1.github.io/manifest.json">
  ```
- [ ] **步驟 2：修改 `apps/LKC_SundayserviceAttendance/index.html`**
  在 `<head>` 中插入：
  ```html
  <link rel="manifest" href="/LKC1958_June_1.github.io/manifest.json">
  ```
- [ ] **步驟 3：修改 `apps/LKC_worship/admin.html`**
  在 `<head>` 中插入：
  ```html
  <link rel="manifest" href="/LKC1958_June_1.github.io/manifest.json">
  ```
- [ ] **步驟 4：確認引用正確性**
  檢查 HTML 是否在合適的 `<head>` 區段中引用。
- [ ] **步驟 5：Commit**
  ```bash
  git add apps/LKC_Group/index.html apps/LKC_SundayserviceAttendance/index.html apps/LKC_worship/admin.html
  git commit -m "feat: add manifest link in apps HTML"
  ```

---

## 任務 3：編寫 PWA Service Worker 核心快取邏輯

**文件：**
- 創建：`service-worker.js`

- [ ] **步驟 1：建立 `service-worker.js` 檔案**
  內容：
  ```javascript
  const CACHE_NAME = 'lkc-pwa-cache-v1';
  const PRECACHE_ASSETS = [
    '/LKC1958_June_1.github.io/index.html',
    '/LKC1958_June_1.github.io/config.js',
    '/LKC1958_June_1.github.io/manifest.json'
  ];

  // 1. 安裝事件：預先快取核心資源
  self.addEventListener('install', event => {
    event.waitUntil(
      caches.open(CACHE_NAME)
        .then(cache => cache.addAll(PRECACHE_ASSETS))
        .then(() => self.skipWaiting())
    );
  });

  // 2. 啟用事件：清理舊版快取
  self.addEventListener('activate', event => {
    event.waitUntil(
      caches.keys().then(keys => {
        return Promise.all(
          keys.map(key => {
            if (key !== CACHE_NAME) {
              return caches.delete(key);
            }
          })
        );
      }).then(() => self.clients.claim())
    );
  });

  // 3. 攔截請求事件：Stale-While-Revalidate 策略
  self.addEventListener('fetch', event => {
    const url = new URL(event.request.url);

    // 🛡️ 安全防護：GAS 請求、Firebase 認證與所有 POST 請求一律直接過網，不快取
    if (event.request.method === 'POST' || 
        url.hostname.includes('script.google.com') || 
        url.hostname.includes('firebasedatabase.app')) {
      return; // 讓瀏覽器直接去網路請求
    }

    // 僅快取與本站相關的 GET 靜態資源
    if (event.request.method === 'GET') {
      event.respondWith(
        caches.open(CACHE_NAME).then(cache => {
          return cache.match(event.request).then(cachedResponse => {
            // 發送網路請求做背景更新
            const fetchPromise = fetch(event.request).then(networkResponse => {
              if (networkResponse && networkResponse.status === 200) {
                // 快取新版資源
                cache.put(event.request, networkResponse.clone());
              }
              return networkResponse;
            }).catch(err => {
              console.warn('[SW] 背景更新失敗（可能處於離線狀態）:', err);
            });

            // 快取命中則秒回快取；否則等待網路回傳
            return cachedResponse || fetchPromise;
          });
        })
      );
    }
  });
  ```
- [ ] **步驟 2：驗證 JS 語法**
  檢查 service-worker.js 無語法錯誤。
- [ ] **步驟 3：Commit**
  ```bash
  git add service-worker.js
  git commit -m "feat: add service-worker.js with Stale-While-Revalidate caching"
  ```

---

## 任務 4：於前端 config.js 實現 Service Worker 自動註冊

**文件：**
- 修改：`config.js`

- [ ] **步驟 1：修改 `config.js`，在 IIFE 尾端加入註冊邏輯**
  在 `config.js` 的 IIFE (立即執行函數) 結束前（約 608 行附近），在對外暴露與 dispatchEvent 之前插入：
  ```javascript
  // ─────────────────────────────────────────────────────────────
  //  PWA Service Worker 自動註冊（Stale-While-Revalidate 靜默更新版）
  // ─────────────────────────────────────────────────────────────
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', function() {
      // 依據 hostname 動態判定 scope 路徑
      // 本地環境可能是 "/"，而 GitHub Pages 是 "/LKC1958_June_1.github.io/"
      let scopePath = '/';
      let swPath = '/service-worker.js';

      if (window.location.hostname.indexOf('github.io') !== -1) {
        scopePath = '/LKC1958_June_1.github.io/';
        swPath = '/LKC1958_June_1.github.io/service-worker.js';
      }

      navigator.serviceWorker.register(swPath, { scope: scopePath })
        .then(function(reg) {
          console.log('✅ [PWA] ServiceWorker 註冊成功，Scope: ', reg.scope);
        }).catch(function(err) {
          console.warn('❌ [PWA] ServiceWorker 註冊失敗: ', err);
        });
    });
  }
  ```
- [ ] **步驟 2：檢查註冊邏輯與語法**
  確保代碼不破壞 `config.js` 本身的 IIFE 結構。
- [ ] **步驟 3：Commit**
  ```bash
  git add config.js
  git commit -m "feat: integrate automatic PWA SW registration in config.js"
  ```
