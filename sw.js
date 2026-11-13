// v4.73: MCCB and meter numbers read only from close-up photos
const SW_VER = '4.73';
self.addEventListener('install', (event) => {
    self.skipWaiting();
});

self.addEventListener('activate', (event) => {
    event.waitUntil((async () => {
        const keys = await caches.keys();
        await Promise.all(keys.map((k) => caches.delete(k)));
        await self.clients.claim();
    })());
});

self.addEventListener('fetch', (event) => {
    const req = event.request;
    if (req.method !== 'GET') return;
    const url = new URL(req.url);
    const bust = url.pathname.endsWith('/') || url.pathname.endsWith('.html') || url.pathname.endsWith('version.json') || url.pathname.endsWith('sw.js');
    const freshReq = bust
        ? new Request(req.url, { cache: 'reload', credentials: req.credentials, mode: req.mode, headers: req.headers })
        : new Request(req.url, { cache: 'no-store', credentials: req.credentials, mode: req.mode, headers: req.headers });
    event.respondWith(fetch(freshReq).catch(() => fetch(req)));
});
