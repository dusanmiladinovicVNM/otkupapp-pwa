const CACHE_NAME = 'otkupapp-v33';
const ASSETS = [
    './index.html',
    './manifest.json',
    'https://unpkg.com/html5-qrcode@2.3.8/html5-qrcode.min.js'
];

// Install: cache app shell
self.addEventListener('install', (event) => {
    event.waitUntil(
        caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS))
    );
    self.skipWaiting();
});

// Activate: clean old caches
self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys().then(keys =>
            Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
        )
    );
    self.clients.claim();
});

// Fetch: network-first for HTML, cache-first for libraries
self.addEventListener('fetch', (event) => {
    const url = new URL(event.request.url);
    
    // API calls: always network, no SW interception
    if (url.hostname === 'script.google.com') {
        return;
    }
    
    // HTML pages: network-first (so updates arrive immediately)
    if (event.request.destination === 'document') {
        event.respondWith(
            fetch(event.request).then(response => {
                const clone = response.clone();
                caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
                return response;
            }).catch(() => {
                return caches.match(event.request) || caches.match('./index.html');
            })
        );
        return;
    }
    
    // Everything else: cache first, fallback to network
    event.respondWith(
        caches.match(event.request).then(cached => {
            return cached || fetch(event.request).then(response => {
                if (response.status === 200) {
                    const clone = response.clone();
                    caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
                }
                return response;
            });
        }).catch(() => {
            if (event.request.destination === 'document') {
                return caches.match('./index.html');
            }
        })
    );
});
