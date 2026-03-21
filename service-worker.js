const CACHE_NAME = 'arch-toolkit-v1';
const ASSETS = [
    './',
    './index.html',
    './style.css',
    './app.js',
    './assets/icon.png',
    'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.0/exceljs.min.js',
    'https://fonts.googleapis.com/css2?family=Outfit:wght@400;600;700&display=swap'
];

self.addEventListener('install', event => {
    event.waitUntil(
        caches.open(CACHE_NAME).then(cache => {
            return cache.addAll(ASSETS);
        })
    );
});

self.addEventListener('fetch', event => {
    event.respondWith(
        caches.match(event.request).then(response => {
            return response || fetch(event.request);
        })
    );
});
