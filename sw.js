// v4.29: read only the MCCB amp rating and the meter number from the panel photo
self.addEventListener('install', (event) => {
    self.skipWaiting();
});

self.addEventListener('activate', (event) => {
    event.waitUntil(self.clients.claim());
});
