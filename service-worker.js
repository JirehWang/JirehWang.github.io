const CACHE_NAME = 'lkc-pwa-cache-v20260606_green_v3';
const PRECACHE_ASSETS = [
  'config.js',
  'manifest.json'
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

  // 🛡️ 僅處理 http 和 https 協議，避免 chrome-extension 等協議報錯
  if (url.protocol !== 'http:' && url.protocol !== 'https:') {
    return;
  }

  // 🛡️ 安全防護：GAS 請求、Firebase 認證與所有 POST 請求一律直接過網，不快取
  if (event.request.method === 'POST' || 
      url.hostname.includes('script.google.com') || 
      url.hostname.includes('firebasedatabase.app')) {
    return; // 讓瀏覽器直接去網路請求
  }

  // 🛡️ 如果網址參數帶有版本或強制刷新標記（v= 或 nocache=），直接向伺服器請求，不快取
  if (url.search.includes('v=') || url.search.includes('nocache=')) {
    return;
  }

  // 🛡️ 所有 HTML 網頁採用「即時響應」，絕不快取，直接過網，防止 HTML 快取鎖死前端更新
  if (url.pathname === '/' || url.pathname.endsWith('.html') || url.pathname.endsWith('/')) {
    return;
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
          });

          // ⚠️ 關鍵修正：如果快取命中，直接回傳；並在背景處理 fetch 更新
          if (cachedResponse) {
            fetchPromise.catch(err => {
              console.warn('[SW] 背景更新失敗（可能處於離線狀態）:', err);
            });
            return cachedResponse;
          }

          // 如果快取未命中，則等待網路返回。若網路也失敗，則直接向瀏覽器拋出錯誤
          return fetchPromise;
        });
      })
    );
  }
});
