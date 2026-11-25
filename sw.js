// 端末内の画面を即座に返す（通信を待たない）。更新は裏で取得し、次回の起動から反映する
const CACHE = 'lifework-4935fead7f36';
const ASSETS = ["./index.html","./manifest.json","./icon.svg","./vendor/icons.css","./vendor/fa-solid-900.woff2","./vendor/fa-regular-400.woff2"];

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

// 共有の受け取り: 内容を端末内に置いてから画面を開く（内容はどこにも送らない）
const SHARE_INBOX = 'share-inbox';
async function receiveShare(req) {
  const f = await req.formData();
  const data = { title: f.get('title') || '', text: f.get('text') || '', url: f.get('url') || '', at: Date.now() };
  const c = await caches.open(SHARE_INBOX);
  await c.put(new URL('./__shared', self.registration.scope).href, new Response(JSON.stringify(data), { headers: { 'content-type': 'application/json' } }));
  return Response.redirect(new URL('./index.html?share=1', self.registration.scope).href, 303);
}

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method === 'POST' && new URL(req.url).pathname.endsWith('/share-target')) {
    event.respondWith(receiveShare(req));
    return;
  }
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;
  event.respondWith(
    caches.open(CACHE)
      .then(c => c.match(req.mode === 'navigate' ? './index.html' : req, { ignoreSearch: true }))
      .then(hit => hit || fetch(req))
  );
});
