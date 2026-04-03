window.openDB = function () {
    return new Promise((resolve, reject) => {
        const req = indexedDB.open(CONFIG.DB_NAME, CONFIG.DB_VERSION);

        req.onupgradeneeded = (e) => {
            const d = e.target.result;

            if (!d.objectStoreNames.contains(CONFIG.STORE_NAME)) {
                const s = d.createObjectStore(CONFIG.STORE_NAME, { keyPath: 'clientRecordID' });
                s.createIndex('syncStatus', 'syncStatus', { unique: false });
                s.createIndex('datum', 'datum', { unique: false });
            }

            if (!d.objectStoreNames.contains(CONFIG.STAMM_STORE)) {
                d.createObjectStore(CONFIG.STAMM_STORE, { keyPath: 'key' });
            }

            if (!d.objectStoreNames.contains(CONFIG.AGRO_STORE)) {
                const a = d.createObjectStore(CONFIG.AGRO_STORE, { keyPath: 'clientRecordID' });
                a.createIndex('syncStatus', 'syncStatus', { unique: false });
            }

            if (!d.objectStoreNames.contains('zbirne')) {
                const z = d.createObjectStore('zbirne', { keyPath: 'clientRecordID' });
                z.createIndex('syncStatus', 'syncStatus', { unique: false });
            }

            if (!d.objectStoreNames.contains('tretmani')) {
                const t = d.createObjectStore('tretmani', { keyPath: 'clientRecordID' });
                t.createIndex('syncStatus', 'syncStatus', { unique: false });
                t.createIndex('datum', 'datum', { unique: false });
                t.createIndex('parcelaID', 'parcelaID', { unique: false });
            }
        };

        req.onsuccess = (e) => resolve(e.target.result);
        req.onerror = (e) => reject(e.target.error);
    });
};

window.dbPut = function (db, storeName, data) {
    return new Promise((resolve, reject) => {
        const tx = db.transaction(storeName, 'readwrite');
        tx.objectStore(storeName).put(data);
        tx.oncomplete = () => resolve();
        tx.onerror = (e) => reject(e.target.error);
    });
};

window.dbGet = function (db, storeName, key) {
    return new Promise((resolve, reject) => {
        const tx = db.transaction(storeName, 'readonly');
        const req = tx.objectStore(storeName).get(key);
        req.onsuccess = () => resolve(req.result);
        req.onerror = (e) => reject(e.target.error);
    });
};

window.dbGetAll = function (db, storeName) {
    return new Promise((resolve, reject) => {
        const tx = db.transaction(storeName, 'readonly');
        const req = tx.objectStore(storeName).getAll();
        req.onsuccess = () => resolve(req.result);
        req.onerror = (e) => reject(e.target.error);
    });
};

window.dbGetByIndex = function (db, storeName, indexName, value) {
    return new Promise((resolve, reject) => {
        const tx = db.transaction(storeName, 'readonly');
        const idx = tx.objectStore(storeName).index(indexName);
        const req = idx.getAll(value);
        req.onsuccess = () => resolve(req.result);
        req.onerror = (e) => reject(e.target.error);
    });
};

window.dbDelete = function (db, storeName, key) {
    return new Promise((resolve, reject) => {
        const tx = db.transaction(storeName, 'readwrite');
        const req = tx.objectStore(storeName).delete(key);
        req.onsuccess = () => resolve();
        req.onerror = (e) => reject(e.target.error);
    });
};
