// v4.45: clear MCCB/meter fields on photo delete; prefer 02xxxxx manufacturing no.
self.addEventListener('install', (event) => {
    self.skipWaiting();
});

self.addEventListener('activate', (event) => {
    event.waitUntil(self.clients.claim());
});
