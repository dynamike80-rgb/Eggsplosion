// Minimale service worker - vereist voor PWA-installatie, geen offline-caching.
// Elke request wordt gewoon normaal doorgestuurd naar het netwerk.
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', (event) => {
  event.waitUntil(self.clients.claim());
});
self.addEventListener('fetch', (event) => {
  event.respondWith(fetch(event.request));
});
