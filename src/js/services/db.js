(function () {
    const DB_OPEN_TIMEOUT_MS = 8000;

    function getDbName() {
        return CONFIG.DB_NAME;
    }

    function getDbVersion() {
        return CONFIG.DB_VERSION;
    }

    function getMainStoreName() {
        return CONFIG.STORE_NAME;
    }

    function getStammStoreName() {
        return CONFIG.STAMM_STORE;
    }

    function buildDbSchema() {
        return [
            {
                name: getMainStoreName(),
                options: { keyPath: 'clientRecordID' },
                indexes: [
                    { name: 'syncStatus', keyPath: 'syncStatus', options: { unique: false } },
                    { name: 'datum', keyPath: 'datum', options: { unique: false } }
                ]
            },
            {
                name: getStammStoreName(),
                options: { keyPath: 'key' },
                indexes: []
            },
            {
                name: 'zbirne',
                options: { keyPath: 'clientRecordID' },
                indexes: [
                    { name: 'syncStatus', keyPath: 'syncStatus', options: { unique: false } }
                ]
            },
            {
                name: 'tretmani',
                options: { keyPath: 'clientRecordID' },
                indexes: [
                    { name: 'syncStatus', keyPath: 'syncStatus', options: { unique: false } },
                    { name: 'datum', keyPath: 'datum', options: { unique: false } },
                    { name: 'parcelaID', keyPath: 'parcelaID', options: { unique: false } }
                ]
            },
            {
                name: 'troskovi',
                options: { keyPath: 'clientRecordID' },
                indexes: [
                    { name: 'syncStatus', keyPath: 'syncStatus', options: { unique: false } },
                    { name: 'datum', keyPath: 'datum', options: { unique: false } }
                ]
            }
        ];
    }

    function ensureObjectStore(db, schemaItem) {
        let store;

        if (!db.objectStoreNames.contains(schemaItem.name)) {
            store = db.createObjectStore(schemaItem.name, schemaItem.options || {});
        } else {
            store = null;
        }

        if (store) {
            (schemaItem.indexes || []).forEach((idx) => {
                store.createIndex(idx.name, idx.keyPath, idx.options || { unique: false });
            });
        }
    }

    function runDbMigrations(db, oldVersion, newVersion) {
        console.log('[DB] upgrade', { oldVersion, newVersion });

        // Legacy cleanup: agromere store više nije aktivan store.
        if (db.objectStoreNames.contains('agromere')) {
            db.deleteObjectStore('agromere');
        }

        buildDbSchema().forEach((schemaItem) => {
            ensureObjectStore(db, schemaItem);
        });
    }

    function attachDbLifecycleHandlers(db) {
        if (!db) return db;

        db.onversionchange = function () {
            console.warn('[DB] versionchange detected, closing current connection');
            try {
                db.close();
            } catch (_) {}
        };

        db.onclose = function () {
            console.warn('[DB] connection closed');
        };

        db.onerror = function (event) {
            console.error('[DB] database error:', event && event.target && event.target.error);
        };

        return db;
    }

    function isRecoverableDbError(err) {
        const msg = String((err && err.message) || err || '').toLowerCase();

        return (
            msg.includes('version') ||
            msg.includes('upgrade') ||
            msg.includes('object store') ||
            msg.includes('index') ||
            msg.includes('notfounderror') ||
            msg.includes('invalidstateerror') ||
            msg.includes('aborterror') ||
            msg.includes('blocked') ||
            msg.includes('timeout')
        );
    }

    function tryOpenDb(isRecoveryAttempt) {
        return new Promise((resolve, reject) => {
            let settled = false;

            const timeoutId = setTimeout(() => {
                if (settled) return;
                settled = true;
                reject(new Error('IndexedDB open timeout'));
            }, DB_OPEN_TIMEOUT_MS);

            let req;
            try {
                req = indexedDB.open(getDbName(), getDbVersion());
            } catch (err) {
                clearTimeout(timeoutId);
                reject(err);
                return;
            }

            req.onupgradeneeded = function (event) {
                const db = event.target.result;
                const oldVersion = event.oldVersion || 0;
                const newVersion = event.newVersion || getDbVersion();

                runDbMigrations(db, oldVersion, newVersion);
            };

            req.onblocked = function () {
                clearTimeout(timeoutId);
                if (settled) return;
                settled = true;
                reject(new Error(isRecoveryAttempt ? 'IndexedDB blocked during recovery' : 'IndexedDB blocked'));
            };

            req.onerror = function (event) {
                clearTimeout(timeoutId);
                if (settled) return;
                settled = true;
                reject(event && event.target ? event.target.error : new Error('IndexedDB open failed'));
            };

            req.onsuccess = function (event) {
                clearTimeout(timeoutId);
                if (settled) return;
                settled = true;

                const db = event.target.result;
                resolve(attachDbLifecycleHandlers(db));
            };
        });
    }

    function closeDbSilently(db) {
        try {
            if (db && typeof db.close === 'function') {
                db.close();
            }
        } catch (_) {}
    }

    function deleteDatabaseSafe() {
        return new Promise((resolve, reject) => {
            let req;
            try {
                req = indexedDB.deleteDatabase(getDbName());
            } catch (err) {
                reject(err);
                return;
            }

            req.onblocked = function () {
                reject(new Error('IndexedDB delete blocked'));
            };

            req.onerror = function (event) {
                reject(event && event.target ? event.target.error : new Error('IndexedDB delete failed'));
            };

            req.onsuccess = function () {
                resolve();
            };
        });
    }

    async function resetDbAndReopen() {
        closeDbSilently(window.db);

        try {
            await deleteDatabaseSafe();
        } catch (err) {
            console.error('[DB] deleteDatabaseSafe failed:', err);
            throw err;
        }

        const db = await tryOpenDb(true);
        window.db = db;
        return db;
    }

    function hasStore(db, storeName) {
        return !!(db && db.objectStoreNames && db.objectStoreNames.contains(storeName));
    }

    function assertStoreExists(db, storeName) {
        if (!hasStore(db, storeName)) {
            throw new Error('Missing object store: ' + storeName);
        }
    }

    function assertIndexExists(store, storeName, indexName) {
        if (!store.indexNames || !store.indexNames.contains(indexName)) {
            throw new Error('Missing index "' + indexName + '" on store "' + storeName + '"');
        }
    }

    function withStore(db, storeName, mode, executor) {
        return new Promise((resolve, reject) => {
            try {
                assertStoreExists(db, storeName);

                const tx = db.transaction(storeName, mode);
                const store = tx.objectStore(storeName);
                const result = executor(store, tx);

                tx.oncomplete = function () {
                    resolve(result);
                };

                tx.onerror = function (event) {
                    reject(event && event.target ? event.target.error : new Error('IndexedDB transaction failed'));
                };

                tx.onabort = function (event) {
                    reject(event && event.target ? event.target.error : new Error('IndexedDB transaction aborted'));
                };
            } catch (err) {
                reject(err);
            }
        });
    }

    window.openDB = async function openDB() {
        try {
            const db = await tryOpenDb(false);
            window.db = db;
            return db;
        } catch (err) {
            console.error('[DB] open failed:', err);

            if (!isRecoverableDbError(err)) {
                throw err;
            }

            console.warn('[DB] attempting recovery reset...');
            return await resetDbAndReopen();
        }
    };

    window.resetIndexedDb = async function resetIndexedDb() {
        return await resetDbAndReopen();
    };

    window.dbPut = function dbPut(db, storeName, data) {
        return new Promise((resolve, reject) => {
            try {
                assertStoreExists(db, storeName);

                const tx = db.transaction(storeName, 'readwrite');
                tx.objectStore(storeName).put(data);

                tx.oncomplete = function () {
                    resolve();
                };

                tx.onerror = function (event) {
                    reject(event && event.target ? event.target.error : new Error('dbPut failed'));
                };

                tx.onabort = function (event) {
                    reject(event && event.target ? event.target.error : new Error('dbPut aborted'));
                };
            } catch (err) {
                reject(err);
            }
        });
    };

    window.dbGet = function dbGet(db, storeName, key) {
        return new Promise((resolve, reject) => {
            try {
                assertStoreExists(db, storeName);

                const tx = db.transaction(storeName, 'readonly');
                const req = tx.objectStore(storeName).get(key);

                req.onsuccess = function () {
                    resolve(req.result);
                };

                req.onerror = function (event) {
                    reject(event && event.target ? event.target.error : new Error('dbGet failed'));
                };
            } catch (err) {
                reject(err);
            }
        });
    };

    window.dbGetAll = function dbGetAll(db, storeName) {
        return new Promise((resolve, reject) => {
            try {
                assertStoreExists(db, storeName);

                const tx = db.transaction(storeName, 'readonly');
                const req = tx.objectStore(storeName).getAll();

                req.onsuccess = function () {
                    resolve(req.result || []);
                };

                req.onerror = function (event) {
                    reject(event && event.target ? event.target.error : new Error('dbGetAll failed'));
                };
            } catch (err) {
                reject(err);
            }
        });
    };

    window.dbGetByIndex = function dbGetByIndex(db, storeName, indexName, value) {
        return new Promise((resolve, reject) => {
            try {
                assertStoreExists(db, storeName);

                const tx = db.transaction(storeName, 'readonly');
                const store = tx.objectStore(storeName);
                assertIndexExists(store, storeName, indexName);

                const req = store.index(indexName).getAll(value);

                req.onsuccess = function () {
                    resolve(req.result || []);
                };

                req.onerror = function (event) {
                    reject(event && event.target ? event.target.error : new Error('dbGetByIndex failed'));
                };
            } catch (err) {
                reject(err);
            }
        });
    };

    window.dbDelete = function dbDelete(db, storeName, key) {
        return new Promise((resolve, reject) => {
            try {
                assertStoreExists(db, storeName);

                const tx = db.transaction(storeName, 'readwrite');
                const req = tx.objectStore(storeName).delete(key);

                req.onsuccess = function () {
                    resolve();
                };

                req.onerror = function (event) {
                    reject(event && event.target ? event.target.error : new Error('dbDelete failed'));
                };

                tx.onabort = function (event) {
                    reject(event && event.target ? event.target.error : new Error('dbDelete aborted'));
                };
            } catch (err) {
                reject(err);
            }
        });
    };
})();
