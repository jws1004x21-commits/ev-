// v4.57: emergency mode panel + hard cache bust
const SW_VER = '4.57';
self.addEventListener('install', (event) => {
    self.skipWaiting();
});

self.addEventListener('activate', (event) => {
    event.waitUntil((async () => {
        const keys = await caches.keys();
        await Promise.all(keys.map((k) => caches.delete(k)));
        await self.clients.claim();
        const clients = await self.clients.matchAll({ type: 'window' });
        clients.forEach((c) => {
            try { c.postMessage({ type: 'SW_ACTIVATED', version: SW_VER }); } catch (err) {}
        });
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
