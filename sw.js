// 更新圖示或其他檔案後，把版本號加 1，手機上的舊快取就會被換掉
const VERSION = 'v1';
const CACHE = 'rehab-demo-' + VERSION;
const SHELL = [
  './',
  'index.html',
  'manifest.webmanifest',
  'icons/icon-192.png',
  'icons/icon-512.png',
  'icons/icon-maskable-512.png',
  'icons/apple-touch-icon.png'
];
const FONT_HOSTS = ['fonts.googleapis.com', 'fonts.gstatic.com'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k.startsWith('rehab-demo-') && k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // 頁面：先抓網路上的最新版，沒網路時用快取，所以內容一改就會生效
  if (req.mode === 'navigate') {
    e.respondWith(
      fetch(req)
        .then(res => {
          const copy = res.clone();
          caches.open(CACHE).then(c => c.put('index.html', copy));
          return res;
        })
        .catch(() => caches.match('index.html').then(r => r || caches.match('./')))
    );
    return;
  }

  // 字型：先用快取，背景再更新
  if (FONT_HOSTS.includes(url.hostname)) {
    e.respondWith(
      caches.open(CACHE).then(c =>
        c.match(req).then(hit => {
          const net = fetch(req)
            .then(res => { if (res.ok || res.type === 'opaque') c.put(req, res.clone()); return res; })
            .catch(() => hit || new Response('', { status: 504 }));
          return hit || net;
        })
      )
    );
    return;
  }

  // 同網域的圖示等檔案：先用快取
  if (url.origin === self.location.origin) {
    e.respondWith(caches.match(req).then(hit => hit || fetch(req)));
  }
});
