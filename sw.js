// 端末内の画面を即座に返す（通信を待たない）。更新は裏で取得し、次回の起動から反映する
const CACHE = 'lifework-24859844d80a';
const ASSETS = ["./index.html","./manifest.json","./icon.svg","./vendor/icons.css","./vendor/chart.umd.js","./vendor/fa-solid-900.woff2","./vendor/fa-regular-400.woff2"];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE)
      .then(c => c.addAll(ASSETS.map(u => new Request(u, { cache: 'reload' }))))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k.startsWith('lifework-') && k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;
  event.respondWith(
    caches.open(CACHE)
      .then(c => c.match(req.mode === 'navigate' ? './index.html' : req, { ignoreSearch: true }))
      .then(hit => hit || fetch(req))
  );
});
