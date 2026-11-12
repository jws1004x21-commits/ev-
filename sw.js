// v4.30: Excel daily log download and close-up MCCB/meter photo fill
self.addEventListener('install', (event) => {
    self.skipWaiting();
});

self.addEventListener('activate', (event) => {
    event.waitUntil(self.clients.claim());
});
