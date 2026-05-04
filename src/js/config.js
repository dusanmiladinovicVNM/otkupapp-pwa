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
    APP_VERSION: '2.0.1'
};
