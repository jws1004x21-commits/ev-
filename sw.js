// v4.49: drop Google Maps dir (no API); Kakao/Naver navigation only
self.addEventListener('install', (event) => {
    self.skipWaiting();
});

self.addEventListener('activate', (event) => {
    event.waitUntil(self.clients.claim());
});
