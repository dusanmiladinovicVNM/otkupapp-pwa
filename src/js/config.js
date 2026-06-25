const SESSION_EXPIRES_AT = getLs('authExpiresAt', '');

function isStoredAuthExpired(expiresAt) {
    if (!expiresAt) return false;

    const expMs = Date.parse(expiresAt);
    if (!Number.isFinite(expMs)) return true;

    return Date.now() > expMs;
}

if (isStoredAuthExpired(SESSION_EXPIRES_AT)) {
    removeLs(['userRole', 'otkupacID', 'entityID', 'entityName', 'authToken', 'authExpiresAt', 'username']);
}

const SESSION_ROLE = getLs('userRole', '');
const SESSION_ENTITY_ID = getLs('entityID', '') || getLs('otkupacID', '');
const SESSION_AUTH_EXPIRES_AT = getLs('authExpiresAt', '');

// ============================================================
// TENANT REGISTRACIJA (silo model: 1 GAS + 1 Drive po klijentu)
// ------------------------------------------------------------
// Svaki klijent ima svoj GAS Web App `/exec` URL i svoj IndexedDB.
// Frontend je jedan (app.agrix.rs); klijent se bira URL-om PRE logina:
//
//     app.agrix.rs/bucaijoca/  ->  /?t=bucaijoca  ->  TENANTS.bucaijoca  (vidi /bucaijoca/index.html)
//
// Slug (kljuc mape) MORA da se poklapa sa imenom subfoldera i sa `?t=`.
// Izbor se persistuje (localStorage 'tenant') da PWA install (start_url "/")
// i naredna otvaranja ostanu na istom klijentu dok se ne klikne drugi link.
//
// >> POPUNI: zameni PLACEHOLDER_GAS_* sa stvarnim `/exec` URL-ovima 3 GAS-a.
//    k1 je trenutni produkcioni URL (ostaje da postojeci klijent radi bez prekida).
// ============================================================
const TENANTS = {
    bucaijoca: {
        name: 'Buca i Joca',
        API_URL: 'https://script.google.com/macros/s/AKfycbw8lxyA-NaYZtfhsCPqs4LxUNK4V-iH9ZFVOwmWTGSu_CdOeqcJJ9_b2DKCg4iqu-DM/exec'
    },
    venivo: {
        name: 'Venivo',
        API_URL: 'https://script.google.com/macros/s/AKfycbzwiFihoCyPPq78unIRlu8KlZc9qwovEUkUabC4U593RoU4VkqbNIN5mY1AS0vDr8ay/exec'
    },
    bukovik: {
        name: 'Bukovik',
        API_URL: 'https://script.google.com/macros/s/AKfycbyrg-RfpFwxVwkWVx9gfuUZQC91yAY2ghzNxeD8YmFarMTzxGb-7_ulRRtWHdOc7Mfjqw/exec'
    }
};

// Postojeci/primarni klijent. Drzi legacy DB ime (`OtkupAppDB`) da mu se
// offline baza ne preimenuje pri prelasku na multitenant. Postavi na slug
// klijenta koji vec radi u produkciji.
const DEFAULT_TENANT = 'bucaijoca';

function resolveTenant() {
    let t = '';

    // 1) ?t=slug (per-klijent bookmark/redirect) — prioritet, persistuje se
    try {
        t = (new URLSearchParams(location.search).get('t') || '').trim();
    } catch (e) {
        t = '';
    }

    // 2) prvi segment putanje (npr. otvoren /k1/ direktno) kao alternativa
    if (!TENANTS[t]) {
        const seg = (location.pathname.split('/').filter(Boolean)[0] || '').trim();
        if (TENANTS[seg]) t = seg;
    }

    // 3) zapamceni izbor
    if (!TENANTS[t]) t = getLs('tenant', '');

    // 4) default (postojeci klijent)
    if (!TENANTS[t]) t = DEFAULT_TENANT;

    setLs('tenant', t);
    return t;
}

const ACTIVE_TENANT = resolveTenant();
const TENANT_CFG = TENANTS[ACTIVE_TENANT] || TENANTS[DEFAULT_TENANT];

// ============================================================
// CONFIG
// ============================================================
window.CONFIG = {
    API_URL: TENANT_CFG.API_URL,
    TENANT: ACTIVE_TENANT,
    TENANT_NAME: TENANT_CFG.name,
    AUTH_EXPIRES_AT: SESSION_AUTH_EXPIRES_AT,
    OTKUPAC_ID: SESSION_ROLE === 'Otkupac' ? SESSION_ENTITY_ID : '',
    USER_ROLE: SESSION_ROLE,
    ENTITY_ID: SESSION_ENTITY_ID,
    ENTITY_NAME: getLs('entityName', ''),
    USERNAME: getLs('username', ''),
    TOKEN: getLs('authToken', ''),
    // Izolovan IndexedDB po klijentu. Primarni klijent zadrzava legacy ime
    // da se postojeci offline podaci ne osirote pri prelasku na multitenant.
    DB_NAME: ACTIVE_TENANT === DEFAULT_TENANT ? 'OtkupAppDB' : ('OtkupAppDB_' + ACTIVE_TENANT),
    DB_VERSION: 6,
    STORE_NAME: 'otkupi',
    STAMM_STORE: 'stammdaten',
    APP_VERSION: '1.0.0-C001',
    FIREBASE_API_KEY:    'AIzaSyAh-OhV1qAYl3blAPrvt3Kg9TUjeaNSlMw',
    FIREBASE_PROJECT_ID: 'agrix-25e20',
    FIREBASE_APP_ID:     '1:154375753183:web:ff37154f6c8ce10526486a',
    FIREBASE_RTDB_URL:   'https://agrix-25e20-default-rtdb.europe-west1.firebasedatabase.app/'
};
