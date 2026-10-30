/* Service worker: lưu app vào bộ nhớ đệm để mở được khi mất mạng.
   Tăng VERSION mỗi lần cập nhật app để máy người dùng nhận bản mới. */
const VERSION = 'tkh-v1';
const SHELL = [
  './', 'index.html', 'app.css', 'app.js', 'manifest.webmanifest',
  'icons/icon-192.png', 'icons/icon-512.png', 'icons/apple-touch-icon.png',
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(VERSION).then((c) => Promise.all(SHELL.map((u) => c.add(u).catch(() => {})))).then(() => self.skipWaiting()));
});
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== VERSION).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.pathname.includes('/rest/v1/')) return; // dữ liệu đám mây: luôn lấy trực tiếp
  // App của mình: lấy mạng trước để luôn có bản mới, mất mạng thì dùng bản đệm
  if (url.origin === location.origin) {
    e.respondWith(fetch(req.mode === 'navigate' ? req.url : req, { cache: 'no-cache' }).then((res) => { const copy = res.clone(); caches.open(VERSION).then((c) => c.put(req, copy)); return res; })
      .catch(() => caches.match(req).then((r) => r || caches.match('index.html'))));
    return;
  }
  // Font, thư viện CDN: dùng bản đệm trước
  e.respondWith(caches.match(req).then((hit) => hit || fetch(req).then((res) => {
    if (res.ok || res.type === 'opaque') { const copy = res.clone(); caches.open(VERSION).then((c) => c.put(req, copy)); }
    return res;
  })));
});
