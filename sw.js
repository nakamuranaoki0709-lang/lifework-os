const CACHE_NAME = 'lifework-fast-v20261121100731';
const CORE_ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './icon.svg'
];

self.addEventListener('install', (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(CORE_ASSETS))
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;

  const url = new URL(req.url);
  // GASへのデータ同期通信はキャッシュせずそのまま通す
  if (url.hostname.includes('script.google.com') || url.hostname.includes('googleusercontent.com')) {
    return;
  }

  // アプリ本体とUIライブラリ（Tailwind / FontAwesome）はキャッシュから即返却（0.01秒）
  const isAppAsset = (url.origin === self.location.origin) ||
                     url.hostname.includes('cdn.tailwindcss.com') ||
                     url.hostname.includes('cdnjs.cloudflare.com');

  if (isAppAsset) {
    event.respondWith(
      caches.match(req, { ignoreSearch: url.origin === self.location.origin }).then((cached) => {
        const networkFetch = fetch(req).then((fresh) => {
          if (fresh && (fresh.ok || fresh.type === 'opaque')) {
            caches.open(CACHE_NAME).then((c) => c.put(req, fresh.clone()));
          }
          return fresh;
        }).catch(() => null);

        if (cached) {
          // 既にキャッシュがあれば0.01秒で即返し、裏側で最新化
          return cached;
        }
        return networkFetch.then((res) => res || caches.match('./index.html'));
      })
    );
  }
});