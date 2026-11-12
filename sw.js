// v4.48: faster meter OCR feedback; home self-check with sample Excel; map tile fallback
self.addEventListener('install', (event) => {
    self.skipWaiting();
});

self.addEventListener('activate', (event) => {
    event.waitUntil(self.clients.claim());
});
