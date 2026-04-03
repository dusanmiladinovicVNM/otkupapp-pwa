// ============================================================
// CONFIG
// ============================================================
const CONFIG = {
    API_URL: 'https://script.google.com/macros/s/AKfycbyus4FqtQ9iZjHpqGiwTAgq_dkx5wqQ1x3-WriqlXIVgZ_-dNzoAvGYHejvlr6InXkrUg/exec',
    OTKUPAC_ID: localStorage.getItem('otkupacID') || '',
    USER_ROLE: localStorage.getItem('userRole') || '',
    ENTITY_ID: localStorage.getItem('entityID') || '',
    ENTITY_NAME: localStorage.getItem('entityName') || '',
    TOKEN: localStorage.getItem('authToken') || '',
    DB_NAME: 'OtkupAppDB',
    DB_VERSION: 4,
    STORE_NAME: 'otkupi',
    STAMM_STORE: 'stammdaten',
    AGRO_STORE: 'agromere'
};
