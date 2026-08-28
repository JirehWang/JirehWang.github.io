const CACHE_NAME = 'lkc-pwa-safe-v20260826';
const STATIC_CACHE_EXTENSIONS = /\.(?:css|png|jpg|jpeg|webp|gif|svg|ico|woff2?)$/i;
const isGitHub = typeof self !== 'undefined' && self.location && self.location.hostname.includes('github.io');
const STATIC_CACHE_PATHS = isGitHub ? ['/LKC1958_June_1.github.io/manifest.json'] : ['/manifest.json'];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(STATIC_CACHE_PATHS).catch(err => console.warn('Cache addAll ignored:', err)))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(
        keys
          .filter(key => key !== CACHE_NAME && key.indexOf('lkc-') === 0)
          .map(key => caches.delete(key))
      ))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('message', event => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});

self.addEventListener('fetch', event => {
  const request = event.request;
  const url = new URL(request.url);

  if (request.method !== 'GET') return;
  if (url.protocol !== 'http:' && url.protocol !== 'https:') return;
  if (url.hostname.includes('script.google.com')) return;
  if (url.hostname.includes('firebasedatabase.app')) return;
  if (url.search.includes('nocache=')) return;

  const isNavigation = request.mode === 'navigate';
  const isHtml = isNavigation || url.pathname.endsWith('.html') || url.pathname.endsWith('/');
  const isConfig = url.pathname.endsWith('/config.js');
  const isAppScript = url.pathname.includes('/apps/') && url.pathname.endsWith('.js');

  if (isHtml || isConfig || isAppScript) {
    event.respondWith(fetch(request));
    return;
  }

  const shouldCacheStatic =
    STATIC_CACHE_EXTENSIONS.test(url.pathname) ||
    url.pathname.endsWith('/manifest.json');

  if (!shouldCacheStatic) return;

  event.respondWith(
    caches.open(CACHE_NAME).then(cache => {
      return cache.match(request).then(cachedResponse => {
        const fetchPromise = fetch(request).then(networkResponse => {
          if (networkResponse && networkResponse.status === 200) {
            cache.put(request, networkResponse.clone());
          }
          return networkResponse;
        });
        return cachedResponse || fetchPromise;
      });
    })
  );
});
