//sw.js
const CACHE_NAME = 'AgriX-v20-C001';
const ASSETS = [
    './index.html',
    './manifest.json',

    // Styles
    './src/styles/auth.css',
    './src/styles/base.css',
    './src/styles/components.css',
    './src/styles/components_v2.css',
    './src/styles/features-kooperant.css',
    './src/styles/features-management.css',
    './src/styles/features-otkup.css',
    './src/styles/features-vozac.css',
    './src/styles/layout.css',
    './src/styles/print.css',

    // Utils
    './src/js/utils/dom.js',
    './src/js/utils/storage.js',
    './src/js/utils/lazy.js',
    './src/js/utils/sanitize.js',
    './src/js/utils/format.js',
    './src/js/utils/async.js',
    './src/js/utils/merge.js',
    './src/js/utils/sync-engine.js',
    './src/js/utils/master-sync-guard.js',

    // Config + State
    './src/js/config.js',
    './src/js/state.js',

    // Services
    './src/js/services/db.js',
    './src/js/services/api.js',
    './src/js/services/auth.js',
    './src/js/services/qr.js',

    // UI
    './src/js/ui/agrix-icons.js',
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
    './src/js/features/otkup/intercom-field.js',

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
    './src/js/features/management/intercom-monitor.js',

    // App bootstrap
    './src/js/app.js',

    // Icons
    './icons/icon-192x192.png',
    './icons/icon-512x512.png',

    // Vendor libs (self-hosted)
    './vendor/jspdf.umd.min.js',
    './vendor/leaflet.css',
    './vendor/chart.umd.min.js',
    './vendor/firebase-app-compat.js',
    './vendor/firebase-auth-compat.js',
    './vendor/firebase-database-compat.js',
    './vendor/images/marker-icon.png',
    './vendor/images/marker-icon-2x.png',
    './vendor/images/marker-shadow.png',

    // Fonts (Faza 1 — self-hosted Cormorant + DM Sans)
    './src/styles/fonts.css',
    './vendor/fonts/dm-sans-400.woff2',
    './vendor/fonts/dm-sans-400-ext.woff2',
    './vendor/fonts/dm-sans-500.woff2',
    './vendor/fonts/dm-sans-500-ext.woff2',
    './vendor/fonts/dm-sans-600.woff2',
    './vendor/fonts/dm-sans-600-ext.woff2',
    './vendor/fonts/dm-sans-700.woff2',
    './vendor/fonts/dm-sans-700-ext.woff2',
    './vendor/fonts/cormorant-400.woff2',
    './vendor/fonts/cormorant-400-ext.woff2',
    './vendor/fonts/cormorant-400-italic.woff2',
    './vendor/fonts/cormorant-400-italic-ext.woff2',
    './vendor/fonts/cormorant-600.woff2',
    './vendor/fonts/cormorant-600-ext.woff2',
    './vendor/fonts/cormorant-600-italic.woff2',
    './vendor/fonts/cormorant-600-italic-ext.woff2',
    './vendor/fonts/cormorant-700.woff2',
    './vendor/fonts/cormorant-700-ext.woff2'
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
    // Bypass non-GET (POST/PUT/DELETE ne mogu u Cache API)
    if (event.request.method !== 'GET') return;
    
    const url = new URL(event.request.url);
    
    // API/external calls: always network, no SW interception
    if (
        url.hostname === 'script.google.com' ||
        url.hostname === 'script.googleusercontent.com' ||
        url.hostname === 'identitytoolkit.googleapis.com' ||
        url.hostname === 'securetoken.googleapis.com' ||
        url.hostname.endsWith('.firebaseio.com') ||
        url.hostname.endsWith('.firebasedatabase.app') ||
        url.hostname === 'api.open-meteo.com' ||
        url.hostname === 'api.qrserver.com' ||
        url.hostname === 'server.arcgisonline.com' ||
        url.hostname.endsWith('.tile.openstreetmap.org')
    ) {
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
            return fetch(event.request);
        })
    );
});
