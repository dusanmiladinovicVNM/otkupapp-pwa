const SESSION_ROLE = getLs('userRole', '');
const SESSION_ENTITY_ID = getLs('entityID', '') || getLs('otkupacID', '');

// ============================================================
// CONFIG
// ============================================================
window.CONFIG = {
    API_URL: 'https://script.google.com/macros/s/AKfycbyus4FqtQ9iZjHpqGiwTAgq_dkx5wqQ1x3-WriqlXIVgZ_-dNzoAvGYHejvlr6InXkrUg/exec',
    OTKUPAC_ID: SESSION_ROLE === 'Otkupac' ? SESSION_ENTITY_ID : '',
    USER_ROLE: SESSION_ROLE,
    ENTITY_ID: SESSION_ENTITY_ID,
    ENTITY_NAME: getLs('entityName', ''),
    TOKEN: getLs('authToken', ''),
    DB_NAME: 'OtkupAppDB',
    DB_VERSION: 5,
    STORE_NAME: 'otkupi',
    STAMM_STORE: 'stammdaten',
    AGRO_STORE: 'agromere',
    APP_VERSION: '2.0.1'
};
