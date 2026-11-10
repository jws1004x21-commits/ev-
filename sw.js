// v4.12: drive count, inspector required, meter sheet cols, already-done, update banner
self.addEventListener('install', (event) => {
    self.skipWaiting();
});

self.addEventListener('activate', (event) => {
    event.waitUntil(self.clients.claim());
});
