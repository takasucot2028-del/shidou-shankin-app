const CACHE_NAME = 'shidou-report-v4';

const PRECACHE_FILES = [
  './',
  './index.html',
  './admin.html',
  './css/common.css',
  './css/report.css',
  './css/admin.css',
  './js/config.js',
  './js/report.js',
  './js/admin.js',
  './manifest.json',
  './icons/icon-192.svg',
  './icons/icon-512.svg',
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(PRECACHE_FILES))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(
        keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))
      ))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // GAS APIはネットワーク優先（オフライン時はエラーレスポンスを返す）
  if (url.hostname.includes('script.google.com')) {
    event.respondWith(
      fetch(event.request).catch(() =>
        new Response(
          JSON.stringify({ success: false, error: 'オフラインのため通信できません' }),
          { headers: { 'Content-Type': 'application/json' } }
        )
      )
    );
    return;
  }

  // 同一オリジンのファイル（HTML/CSS/JS等）はネットワーク優先。
  // 常に最新を取得し、成功時はキャッシュ更新。オフライン時のみキャッシュへフォールバック。
  // ※キャッシュ優先だと古いindex.html/JSを返し続け、更新が反映されないため。
  if (event.request.method === 'GET' && url.origin === self.location.origin) {
    event.respondWith(
      fetch(event.request).then(response => {
        if (response && response.status === 200 && response.type === 'basic') {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        }
        return response;
      }).catch(() => caches.match(event.request))
    );
    return;
  }

  // それ以外（CDN等のクロスオリジン）はキャッシュ優先、なければネットワーク
  event.respondWith(
    caches.match(event.request).then(cached => cached || fetch(event.request))
  );
});
