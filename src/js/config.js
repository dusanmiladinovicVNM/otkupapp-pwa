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
// CONFIG
// ============================================================
window.CONFIG = {
    API_URL: 'https://script.google.com/macros/s/AKfycbyus4FqtQ9iZjHpqGiwTAgq_dkx5wqQ1x3-WriqlXIVgZ_-dNzoAvGYHejvlr6InXkrUg/exec',
    AUTH_EXPIRES_AT: SESSION_AUTH_EXPIRES_AT,
    OTKUPAC_ID: SESSION_ROLE === 'Otkupac' ? SESSION_ENTITY_ID : '',
    USER_ROLE: SESSION_ROLE,
    ENTITY_ID: SESSION_ENTITY_ID,
    ENTITY_NAME: getLs('entityName', ''),
    TOKEN: getLs('authToken', ''),
    DB_NAME: 'OtkupAppDB',
    DB_VERSION: 6,
    STORE_NAME: 'otkupi',
    STAMM_STORE: 'stammdaten',
    APP_VERSION: '2.0.1',
    FIREBASE_API_KEY:    'AIzaSyAh-OhV1qAYl3blAPrvt3Kg9TUjeaNSlMw',
    FIREBASE_PROJECT_ID: 'agrix-25e20',
    FIREBASE_APP_ID:     '1:154375753183:web:ff37154f6c8ce10526486a',
    FIREBASE_RTDB_URL:   'https://agrix-25e20-default-rtdb.europe-west1.firebasedatabase.app/'
};
