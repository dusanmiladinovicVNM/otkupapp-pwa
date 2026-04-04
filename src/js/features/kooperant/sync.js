// ============================================================
// KOOPERANT SYNC
// Covers:
// - CONFIG.AGRO_STORE  -> action: syncAgromere
// - 'tretmani'         -> action: syncTretman
//
// Note:
// This file assumes the following globals already exist:
// db, CONFIG, dbGetByIndex, dbPut, apiPost, showToast
// ============================================================

(function () {
    function ensureKooperantRuntime() {
        window.appRuntime = window.appRuntime || {};

        if (!window.appRuntime.kooperantSync) {
            window.appRuntime.kooperantSync = {
                agromereInFlight: false,
                tretmaniInFlight: false
            };
        }

        return window.appRuntime.kooperantSync;
    }

    async function syncEntityStore(options) {
        const {
            storeName,
            action,
            inFlightKey,
            successLabel
        } = options;

        if (!db) return { ok: false, reason: 'db-not-ready' };
        if (!navigator.onLine) return { ok: false, reason: 'offline' };

        const runtime = ensureKooperantRuntime();

        if (runtime[inFlightKey]) {
            return { ok: false, reason: 'already-running' };
        }

        runtime[inFlightKey] = true;

        let pending = [];

        try {
            pending = await dbGetByIndex(db, storeName, 'syncStatus', 'pending');

            if (!Array.isArray(pending) || pending.length === 0) {
                return { ok: true, synced: 0, failed: 0 };
            }

            for (const record of pending) {
                record.syncStatus = 'syncing';
                record.syncAttemptAt = new Date().toISOString();
                await dbPut(db, storeName, record);
            }

            const json = await apiPost(action, {
                kooperantID: CONFIG.ENTITY_ID,
                records: pending
            });

            if (!json || json.success === false) {
                for (const record of pending) {
                    record.syncStatus = 'pending';
                    record.lastSyncError = json && json.error ? json.error : 'Sync neuspešan';
                    record.lastServerStatus = 'request-failed';
                    record.syncAttempts = (record.syncAttempts || 0) + 1;
                    await dbPut(db, storeName, record);
                }

                showToast(successLabel + ' sync nije uspeo', 'error');
                return { ok: false, synced: 0, failed: pending.length };
            }

            if (Array.isArray(json.results)) {
                const byClientId = new Map(
                    pending.map(r => [r.clientRecordID, r])
                );

                const mentionedIds = new Set(
                    json.results.map(r => r.clientRecordID).filter(Boolean)
                );

                let syncedCount = 0;
                let failedCount = 0;

                for (const result of json.results) {
                    const record = byClientId.get(result.clientRecordID);
                    if (!record) continue;

                    record.syncAttempts = (record.syncAttempts || 0) + 1;

                    const isSuccess =
                        !!result.success ||
                        result.status === 'synced' ||
                        result.status === 'duplicate' ||
                        result.status === 'existing' ||
                        result.status === 'inserted' ||
                        result.status === 'updated';

                    if (isSuccess) {
                        record.syncStatus = 'synced';
                        record.lastSyncError = '';
                        record.syncedAt = new Date().toISOString();
                        record.serverRecordID = result.serverRecordID || record.serverRecordID || '';
                        record.updatedAtServer = result.updatedAtServer || record.updatedAtServer || '';
                        record.lastServerStatus = result.status || 'synced';
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

                if (syncedCount > 0 && failedCount === 0) {
                    showToast(successLabel + ': ' + syncedCount, 'success');
                } else if (syncedCount > 0 && failedCount > 0) {
                    showToast(successLabel + ': ' + syncedCount + ' uspešno, ' + failedCount + ' neuspešno', 'info');
                } else {
                    showToast(successLabel + ' nisu sinhronizovane', 'error');
                }

                return { ok: failedCount === 0, synced: syncedCount, failed: failedCount };
            }

            // Legacy fallback for older backend responses: { success: true }
            for (const record of pending) {
                record.syncStatus = 'synced';
                record.lastSyncError = '';
                record.syncedAt = new Date().toISOString();
                record.lastServerStatus = 'legacy-success';
                record.syncAttempts = (record.syncAttempts || 0) + 1;
                await dbPut(db, storeName, record);
            }

            showToast(successLabel + ': ' + pending.length, 'success');
            return { ok: true, synced: pending.length, failed: 0 };
        } catch (err) {
            console.error(action + ' failed:', err);

            for (const record of pending) {
                try {
                    if (record.syncStatus === 'syncing') {
                        record.syncStatus = 'pending';
                        record.lastSyncError = err && err.message ? err.message : 'Greška pri sync-u';
                        record.lastServerStatus = 'exception';
                        record.syncAttempts = (record.syncAttempts || 0) + 1;
                        await dbPut(db, storeName, record);
                    }
                } catch (rollbackErr) {
                    console.error('rollback failed:', rollbackErr);
                }
            }

            showToast('Greška pri sinhronizaciji', 'error');
            return { ok: false, synced: 0, failed: pending.length || 0 };
        } finally {
            runtime[inFlightKey] = false;
        }
    }

    // ------------------------------------------------------------
    // AGROMERE (legacy/simple agro store)
    // Backend action: syncAgromere
    // Store: CONFIG.AGRO_STORE
    // ------------------------------------------------------------
    window.syncAgromere = async function syncAgromere() {
        return syncEntityStore({
            storeName: CONFIG.AGRO_STORE,
            action: 'syncAgromere',
            inFlightKey: 'agromereInFlight',
            successLabel: 'Agromere sinhronizovane'
        });
    };

    // ------------------------------------------------------------
    // TRETMANI (digitalni agronom)
    // Backend action: syncTretman
    // Store: 'tretmani'
    // ------------------------------------------------------------
    window.syncTretmani = async function syncTretmani() {
        return syncEntityStore({
            storeName: 'tretmani',
            action: 'syncTretman',
            inFlightKey: 'tretmaniInFlight',
            successLabel: 'Tretmani sinhronizovani'
        });
    };

    // ------------------------------------------------------------
    // Optional convenience helper
    // ------------------------------------------------------------
    window.syncKooperantNow = async function syncKooperantNow() {
        const results = [];

        // syncAgromere only if store exists in config and function is meaningful in your app
        if (CONFIG && CONFIG.AGRO_STORE) {
            try {
                results.push({ type: 'agromere', ...(await window.syncAgromere()) });
            } catch (e) {
                results.push({ type: 'agromere', ok: false, error: e.message || 'syncAgromere failed' });
            }
        }

        try {
            results.push({ type: 'tretmani', ...(await window.syncTretmani()) });
        } catch (e) {
            results.push({ type: 'tretmani', ok: false, error: e.message || 'syncTretmani failed' });
        }

        return results;
    };
})();
