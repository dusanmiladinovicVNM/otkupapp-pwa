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
    const STALE_SYNCING_MAX_AGE_MS = 2 * 60 * 1000;
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
        const records = Array.isArray(pending) ? pending : [];

        for (const record of records) {
            try {
                if (!record || !record.clientRecordID) continue;

                // Request-level failure means every record from the current batch
                // that did not receive a confirmed server result must become retryable.
                record.syncStatus = 'pending';
                record.lastSyncError = errorMessage || 'Sync neuspešan';
                record.lastServerStatus = serverStatus || 'request-failed';
                record.syncAttempts = (parseInt(record.syncAttempts, 10) || 0) + 1;

                await dbPut(db, storeName, record);
            } catch (rbErr) {
                console.error('[sync-engine] rollback failed:', rbErr);
            }
        }
    }

    async function rollbackPendingAsFeatureDisabled(storeName, pending) {
        // FEATURE_DISABLED: revert current batch to pending without polluting attempts/errors.
        const records = Array.isArray(pending) ? pending : [];

        for (const record of records) {
            try {
                if (!record || !record.clientRecordID) continue;

                record.syncStatus = 'pending';
                record.lastServerStatus = 'feature-disabled';

                await dbPut(db, storeName, record);
            } catch (rbErr) {
                console.error('[sync-engine] feature-disabled rollback failed:', rbErr);
            }
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

    function isStaleSyncingRecord(record, nowMs) {
        if (!record || record.syncStatus !== 'syncing') return false;

        const rawTs =
            record.syncAttemptAt ||
            record.updatedAtClient ||
            record.createdAtClient ||
            '';

        if (!rawTs) return true;

        const ts = new Date(rawTs).getTime();

        if (!Number.isFinite(ts)) return true;

        return nowMs - ts > STALE_SYNCING_MAX_AGE_MS;
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

        const nowMs = Date.now();
        let recovered = 0;

        for (const record of syncing) {
            try {
                if (!isStaleSyncingRecord(record, nowMs)) {
                    continue;
                }

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
        
        // Izračunaj entityID PRE master sync check-a (umesto posle pending fetch-a)
        const entityID = CONFIG.ENTITY_ID || CONFIG.OTKUPAC_ID || '';

        if (typeof window.ensureMasterSyncNotActive === 'function') {
            // Za otkup sync (action='sync'), entityID je otkupacID-stanica.
            // Za druge sync action-e (action='syncZbirna', itd.), stanicaID se ne prosleđuje.
            const lockOpts = { showToast: showToasts };
            if (action === 'sync' && entityID) {
                lockOpts.stanicaID = entityID;
            }
        
            const allowed = await window.ensureMasterSyncNotActive('sync:' + action, lockOpts);

            if (!allowed) {
                return buildSyncResult({
                    ok: true,
                    synced: 0,
                    failed: 0,
                    reason: 'master-sync-active',
                    code: 'MASTER_SYNC_ACTIVE'
                });
            }
        }

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

            const isTemporarySyncLock = json && (
                json.code === 'MASTER_SYNC_ACTIVE' ||
                json.code === 'STANICA_LOCK_ACTIVE'
            );
            if (isTemporarySyncLock) {
                const lockReason = json.lockKind === 'stanica'
                    ? 'stanica-lock-active'
                    : 'master-sync-active';
                
                for (const record of pending) {
                    try {
                        if (!record || !record.clientRecordID) continue;

                        record.syncStatus = 'pending';
                        record.lastServerStatus = 'master-sync-active';
                        record.lastSyncError = '';      // OBAVEZNO clear — soft lock nije greška
                        record.syncAttemptAt = '';
                        // syncAttempts SE NE INKREMENTIRA — to bi pojelo backoff threshold za prave greske

                        await dbPut(db, storeName, record);
                    } catch (rbErr) {
                        console.error('[sync-engine] soft-lock rollback failed:', rbErr);
                    }
                }

                toast(
                    json.message || json.error ||
                        (json.lockKind === 'stanica'
                            ? 'Stanica je u kancelarijskoj obradi. Unos je sačuvan i biće poslat kasnije.'
                            : 'Master sync je u toku. Unos je sačuvan lokalno i biće poslat kasnije.'),
                    'warning'
                );

                return buildSyncResult({
                    ok: true,
                    synced: 0,
                    failed: 0,
                    reason: lockReason,
                    code: json.code || 'MASTER_SYNC_ACTIVE'
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
            
                if (typeof window.reportClientError === 'function') {
                window.reportClientError(err, {
                    source: 'sync-engine',
                    errorAction: 'sync:' + action,
                    details: 'storeName=' + storeName +
                        '\ninFlightKey=' + inFlightKey +
                        '\npending=' + (Array.isArray(pending) ? pending.length : 0)
                });
            }
            
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

    window.recoverStaleSyncingRecords = recoverStaleSyncingRecords;

    window.recoverStaleSyncingStores = async function recoverStaleSyncingStores(storeNames) {
        const stores = Array.isArray(storeNames) ? storeNames : [];
        const results = [];

        for (const storeName of stores) {
            if (!storeName) continue;

            try {
                const recovered = await recoverStaleSyncingRecords(storeName);
                results.push({
                    storeName,
                    recovered
                });
            } catch (err) {
                console.error('[sync-engine] bootstrap stale recovery failed:', storeName, err);
                results.push({
                    storeName,
                    recovered: 0,
                    error: err && err.message ? err.message : String(err || '')
                });
            }
        }

        return results;
    };

    window.buildSyncResult = buildSyncResult;

    window.isAnySyncInFlight = function isAnySyncInFlight() {
        const s = (window.appRuntime || {}).sync || {};
        return !!(s.queueInFlight || s.otkupacInFlight || s.tretmaniInFlight ||
                  s.troskoviInFlight || s.zbirnaInFlight);
    };
})();
