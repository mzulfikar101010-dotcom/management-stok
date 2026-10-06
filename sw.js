// Service worker minimal — cukup untuk syarat "installable" PWA.
// Tidak melakukan caching agresif, supaya data selalu fresh dari Google Sheets.
self.addEventListener('install', (e) => {
  self.skipWaiting();
});

self.addEventListener('activate', (e) => {
  self.clients.claim();
});

self.addEventListener('fetch', (e) => {
  // pass-through biasa, tidak meng-cache apapun
  e.respondWith(fetch(e.request));
});
