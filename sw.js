const CACHE_NAME = 'AgriX-v8';
const ASSETS = [
    './index.html',
    './manifest.json',

        // Styles
    './src/styles/auth.css',
    './src/styles/base.css',
    './src/styles/components.css',
    './src/styles/features-kooperant.css',
    './src/styles/features-management.css',
    './src/styles/features-otkup.css',
    './src/styles/features-vozac.css',
    './src/styles/layout.css',
    './src/styles/print.css',

    // Utils
    './src/js/utils/dom.js',
    './src/js/utils/storage.js',
    './src/js/utils/sanitize.js',
    './src/js/utils/format.js',
    './src/js/utils/async.js',
    './src/js/utils/merge.js',

    // Config + State
    './src/js/config.js',
    './src/js/state.js',

    // Services
    './src/js/services/db.js',
    './src/js/services/api.js',
    './src/js/services/auth.js',
    './src/js/services/qr.js',

    // UI
    './src/js/ui/toast.js',
    './src/js/ui/signatures.js',
    './src/js/ui/tabs.js',
    './src/js/ui/role-nav.js',

    // Features — kooperant
    './src/js/features/kooperant/pregled.js',
    './src/js/features/kooperant/kartica.js',
    './src/js/features/kooperant/koopinfo.js',
    './src/js/features/kooperant/parcele.js',
    './src/js/features/kooperant/sync.js',
    './src/js/features/kooperant/agromere.js',
    './src/js/features/kooperant/knjiga-polja.js',
    './src/js/features/kooperant/fiskalni.js',
    './src/js/features/kooperant/bottom-nav.js',

    // Features — otkup
    './src/js/features/otkup/otkup-form.js',
    './src/js/features/otkup/otkup-pregled.js',
    './src/js/features/otkup/otkupni-list.js',
    './src/js/features/otkup/otpremnice.js',
    './src/js/features/otkup/otkup-more.js',
    './src/js/features/otkup/sync.js',

    // Features — vozac
    './src/js/features/vozac/zbirna.js',
    './src/js/features/vozac/transport.js',

    // Features — management
    './src/js/features/management/kooperanti.js',
    './src/js/features/management/stanice.js',
    './src/js/features/management/kupci.js',
    './src/js/features/management/agrohemija.js',
    './src/js/features/management/dispecer.js',
    './src/js/features/management/mgmt-shell-v2.js',

    // App bootstrap
    './src/js/app.js',

    // Icons
    './icons/icon-192x192.png',
    './icons/icon-512x512.png',
    
    'https://unpkg.com/html5-qrcode@2.3.8/html5-qrcode.min.js',
    'https://unpkg.com/jspdf@2.5.2/dist/jspdf.umd.min.js',
    'https://unpkg.com/leaflet@1.9.4/dist/leaflet.css',
    'https://unpkg.com/leaflet@1.9.4/dist/leaflet.js',
    'https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js'
];

// Install: cache app shell
self.addEventListener('install', (event) => {
    event.waitUntil((async () => {
        const cache = await caches.open(CACHE_NAME);

        await Promise.allSettled(
            ASSETS.map(async (url) => {
                try {
                    await cache.add(url);
                } catch (err) {
                    console.warn('[SW] asset cache failed:', url, err);
                }
            })
        );
    })());

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
