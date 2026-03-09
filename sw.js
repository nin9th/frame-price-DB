const CACHE = 'wood-price-v2';
const SHELL = ['./index.html', './manifest.json', './icons/icon-192.png', './icons/icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  // Always network-first for fonts
  if (e.request.url.includes('fonts.googleapis') || e.request.url.includes('fonts.gstatic')) {
    e.respondWith(fetch(e.request).catch(() => caches.match(e.request)));
    return;
  }
  // Cache-first for app shell
  e.respondWith(caches.match(e.request).then(cached => cached || fetch(e.request)));
});
