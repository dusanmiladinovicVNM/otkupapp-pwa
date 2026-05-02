// ============================================================
// SYNC ENGINE
//
// Single source of truth for all role-specific sync operations.
// All sync functions (syncQueue / syncTretmani / syncTroskovi /
// syncZbirne) MUST go through this engine.
//
// Contract:
//   syncStore({
//     storeName,        // IndexedDB store
//     action,           // GAS action name
//     inFlightKey,      // key in runtime.sync.{key}
//     entityIdField,    // name of the field GAS expects (e.g. 'kooperantID')
//     successLabel,     // toast text on success
//     onResultRecord,   // optional per-record post-success hook
//     showToasts        // default true; set false for silent runs
//   }) -> SyncResult
// ============================================================

(function () {
    function getRuntimeSync() {
        const runtime = window.appRuntime || {};
        if (!runtime.sync) {
            runtime.sync = {
                queueInFlight: false,
                otkupacInFlight: false,
                tretmaniInFlight: false,
                troskoviInFlight: false,
                zbirnaInFlight: false
            };
        }
        return runtime.sync;
    }

    function buildSyncResult(overrides) {
        return Object.assign({
            ok: false,
            synced: 0,
            failed: 0,
            results: [],
            reason: '',
            code: '',
            partial: false
        }, overrides || {});
    }

    function isFeatureDisabled(json) {
        return !!(json && (
            json.code === 'FEATURE_DISABLED' ||
            (json.error && /not\s*enabled|nije\s*aktivan/i.test(String(json.error)))
        ));
    }

    function isAuthError(json) {
        return !!(json && (json.code === 401 || json.code === 'unauthorized'));
    }

    function isSuccessStatus(result) {
        return !!(
            result && (
                result.success === true ||
                result.status === 'synced' ||
                result.status === 'duplicate' ||
                result.status === 'existing' ||
                result.status === 'inserted' ||
                result.status === 'updated'
            )
        );
    }

    async function markPendingAsSyncing(storeName, pending) {
        const ts = new Date().toISOString();
        for (const r of pending) {
            r.syncStatus = 'syncing';
            r.syncAttemptAt = ts;
            await dbPut(db, storeName, r);
        }
    }

    async function rollbackPendingFromError(storeName, pending, errorMessage, serverStatus) {
        for (const r of pending) {
            try {
                if (r.syncStatus === 'syncing') {
                    r.syncStatus = 'pending';
                    r.lastSyncError = errorMessage;
                    r.lastServerStatus = serverStatus || 'request-failed';
                    r.syncAttempts = (r.syncAttempts || 0) + 1;
                    await dbPut(db, storeName, r);
                }
            } catch (rbErr) {
                console.error('[sync-engine] rollback failed:', rbErr);
            }
        }
    }

    async function rollbackPendingAsFeatureDisabled(storeName, pending) {
        // FEATURE_DISABLED: revert syncing -> pending without polluting attempts/errors.
        for (const r of pending) {
            try {
                if (r.syncStatus === 'syncing') {
                    r.syncStatus = 'pending';
                    r.lastServerStatus = 'feature-disabled';
                    await dbPut(db, storeName, r);
                }
            } catch (_) {}
        }
    }

    async function applyServerResults(storeName, pending, results, onResultRecord) {
        const byClientId = new Map(pending.map(r => [r.clientRecordID, r]));
        const mentionedIds = new Set(results.map(r => r.clientRecordID).filter(Boolean));

        let syncedCount = 0;
        let failedCount = 0;

        for (const result of results) {
            const record = byClientId.get(result.clientRecordID);
            if (!record) continue;

            record.syncAttempts = (record.syncAttempts || 0) + 1;

            if (isSuccessStatus(result)) {
                record.syncStatus = 'synced';
                record.lastSyncError = '';
                record.syncedAt = new Date().toISOString();
                record.serverRecordID = result.serverRecordID || record.serverRecordID || '';
                record.updatedAtServer = result.updatedAtServer || record.updatedAtServer || '';
                record.lastServerStatus = result.status || 'synced';

                if (typeof onResultRecord === 'function') {
                    try { onResultRecord(record, result); } catch (_) {}
                }
                syncedCount++;
            } else {
                record.syncStatus = 'pending';
                record.lastSyncError = result.error || 'Sync stavke neuspešan';
                record.lastServerStatus = result.status || 'failed';
                failedCount++;
            }

            await dbPut(db, storeName, record);
        }

        for (const record of pending) {
            if (!mentionedIds.has(record.clientRecordID)) {
                record.syncStatus = 'pending';
                record.lastSyncError = 'Nema potvrde sa servera';
                record.lastServerStatus = 'missing-result';
                record.syncAttempts = (record.syncAttempts || 0) + 1;
                failedCount++;
                await dbPut(db, storeName, record);
            }
        }

        return { syncedCount, failedCount };
    }

    async function recoverStaleSyncingRecords(storeName) {
        let syncing = [];

        try {
            syncing = await dbGetByIndex(db, storeName, 'syncStatus', 'syncing');
        } catch (err) {
            console.error('[sync-engine] stale syncing lookup failed:', err);
            return 0;
        }

        if (!Array.isArray(syncing) || syncing.length === 0) {
            return 0;
        }

        let recovered = 0;

        for (const record of syncing) {
            try {
                record.syncStatus = 'pending';
                record.lastServerStatus = record.lastServerStatus || 'stale-syncing-recovered';

                if (!record.lastSyncError) {
                    record.lastSyncError = 'Recovered from stale syncing state';
                }

                await dbPut(db, storeName, record);
                recovered++;
            } catch (err) {
                console.error('[sync-engine] stale syncing recovery failed:', err);
            }
        }

        return recovered;
    }

    window.syncStore = async function syncStore(options) {
        const {
            storeName,
            action,
            inFlightKey,
            entityIdField = 'kooperantID',
            successLabel = 'Sinhronizovano',
            onResultRecord,
            showToasts = true
        } = options;

        if (!db) return buildSyncResult({ reason: 'db-not-ready' });
        if (!navigator.onLine) return buildSyncResult({ reason: 'offline' });

        const runtimeSync = getRuntimeSync();
        if (runtimeSync[inFlightKey]) {
            return buildSyncResult({ reason: 'already-running' });
        }
        runtimeSync[inFlightKey] = true;

        let pending = [];
        const toast = (msg, kind) => { if (showToasts) showToast(msg, kind); };

        try {
            await recoverStaleSyncingRecords(storeName);
            
            pending = await dbGetByIndex(db, storeName, 'syncStatus', 'pending');
            if (!Array.isArray(pending) || pending.length === 0) {
                return buildSyncResult({ ok: true, reason: 'no-pending' });
            }

            await markPendingAsSyncing(storeName, pending);

            const entityID = CONFIG.ENTITY_ID || CONFIG.OTKUPAC_ID || '';

            if (!entityID) {
                await rollbackPendingFromError(
                    storeName,
                    pending,
                    'Nedostaje entityID za sync',
                    'missing-entity-id'
                );

                return buildSyncResult({
                    failed: pending.length,
                    reason: 'missing-entity-id',
                    code: 'MISSING_ENTITY_ID'
                });
            }

            const payload = { records: pending };
            payload[entityIdField] = entityID;

            const json = await apiPost(action, payload);

            if (!json) {
                await rollbackPendingFromError(
                    storeName,
                    pending,
                    'Prazan odgovor servera',
                    'empty-response'
                );

            toast('Greška pri sinhronizaciji', 'error');

                return buildSyncResult({
                    failed: pending.length,
                    reason: 'empty-response'
                });
            }

            if (isFeatureDisabled(json)) {
                await rollbackPendingAsFeatureDisabled(storeName, pending);

                return buildSyncResult({
                    reason: 'feature-disabled',
                    code: json.code || 'FEATURE_DISABLED'
                });
            }

            if (isAuthError(json)) {
                await rollbackPendingFromError(
                    storeName,
                    pending,
                    'Sesija istekla',
                    'auth-error'
                );

                return buildSyncResult({
                    failed: pending.length,
                    reason: 'auth-error',
                    code: json.code || 401
                });
            }

            // GAS batch contract:
            // success=false može i dalje imati results[].
            // Zato results[] mora da se obradi pre generic success=false grane.
            if (Array.isArray(json.results)) {
                const { syncedCount, failedCount } = await applyServerResults(
                    storeName, pending, json.results, onResultRecord
                );

                if (syncedCount > 0 && failedCount === 0) {
                    toast(successLabel + ': ' + syncedCount, 'success');
                } else if (syncedCount > 0 && failedCount > 0) {
                    toast(successLabel + ': ' + syncedCount + ' uspešno, ' + failedCount + ' neuspešno', 'info');
                } else {
                    toast(successLabel + ' nisu sinhronizovane', 'error');
                }

                return buildSyncResult({
                    ok: failedCount === 0,
                    synced: syncedCount,
                    failed: failedCount,
                    results: json.results,
                    reason: failedCount > 0
                        ? (json.code === 'BATCH_FAILED' ? 'batch-failed' : 'partial-failure')
                        : '',
                    code: json.code || '',
                    partial: syncedCount > 0 && failedCount > 0
                });
            }

            if (json.success === false) {
                const errorMessage = json.error || 'Sync neuspešan';

                await rollbackPendingFromError(
                    storeName,
                    pending,
                    errorMessage,
                    'request-failed'
                );

                toast(successLabel + ' nije uspeo', 'error');

                return buildSyncResult({
                    failed: pending.length,
                    reason: 'server-failed',
                    code: json.code || ''
                });
            }

            // Legacy fallback (no results[] array): treat all as success.
            for (const record of pending) {
                record.syncStatus = 'synced';
                record.lastSyncError = '';
                record.syncedAt = new Date().toISOString();
                record.lastServerStatus = 'legacy-success';
                record.syncAttempts = (record.syncAttempts || 0) + 1;
                await dbPut(db, storeName, record);
            }
            toast(successLabel + ': ' + pending.length, 'success');
            return buildSyncResult({ ok: true, synced: pending.length, reason: 'legacy-success' });

        } catch (err) {
            console.error('[sync-engine] ' + action + ' failed:', err);
            await rollbackPendingFromError(
                storeName, pending,
                (err && err.message) || 'Greška pri sync-u',
                'exception'
            );
            toast('Greška pri sinhronizaciji', 'error');
            return buildSyncResult({
                failed: pending.length || 0,
                reason: 'exception',
                code: (err && err.name) || ''
            });
        } finally {
            runtimeSync[inFlightKey] = false;
        }
    };

    window.buildSyncResult = buildSyncResult;

    window.isAnySyncInFlight = function isAnySyncInFlight() {
        const s = (window.appRuntime || {}).sync || {};
        return !!(s.queueInFlight || s.otkupacInFlight || s.tretmaniInFlight ||
                  s.troskoviInFlight || s.zbirnaInFlight);
    };
})();
