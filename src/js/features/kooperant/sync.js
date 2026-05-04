// ============================================================
// KOOPERANT SYNC
//
// Thin wrappers around the shared sync engine (utils/sync-engine.js).
// Add new entity stores here by registering a new wrapper.
// ============================================================

(function () {
    // ------------------------------------------------------------
    // TRETMANI (digitalni agronom)
    // Backend action: syncTretman
    // Store: 'tretmani'
    // ------------------------------------------------------------
    window.syncTretmani = async function syncTretmani() {
        return syncStore({
            storeName: 'tretmani',
            action: 'syncTretman',
            inFlightKey: 'tretmaniInFlight',
            entityIdField: 'kooperantID',
            successLabel: 'Tretmani sinhronizovani'
        });
    };

    // ------------------------------------------------------------
    // TROSKOVI (knjiga polja)
    // Backend action: syncTrosak
    // Store: 'troskovi'

    // Troškovi koriste isti shared sync engine kao tretmani.
    // Backend action syncTrosak je aktivan; FEATURE_DISABLED ostaje
    // podržan samo kao generic fallback u sync-engine.js.
    // ------------------------------------------------------------
    window.syncTroskovi = async function syncTroskovi() {
        return syncStore({
            storeName: 'troskovi',
            action: 'syncTrosak',
            inFlightKey: 'troskoviInFlight',
            entityIdField: 'kooperantID',
            successLabel: 'Troškovi sinhronizovani'
        });
    };

    // ------------------------------------------------------------
    // Run all kooperant syncs in sequence
    // ------------------------------------------------------------
    window.syncKooperantNow = async function syncKooperantNow() {
        const unitResults = [];

        const syncs = [
            ['tretmani', window.syncTretmani],
            ['troskovi', window.syncTroskovi]
        ];

        for (const [type, fn] of syncs) {
            if (typeof fn !== 'function') {
                unitResults.push({
                    type,
                    ok: false,
                    synced: 0,
                    failed: 0,
                    results: [],
                    reason: 'missing-sync-function',
                    code: '',
                    partial: false
                });
                continue;
            }

            try {
                unitResults.push(normalizeKooperantSyncUnit(type, await fn()));
            } catch (e) {
                unitResults.push({
                    type,
                    ok: false,
                    synced: 0,
                    failed: 0,
                    results: [],
                    reason: (e && e.message) || (type + ' sync failed'),
                    code: (e && e.name) || '',
                    partial: false
                });
            }
        }

        return aggregateKooperantSyncResult(unitResults);
    };

    window.requestKooperantSync = async function requestKooperantSync(reason) {
        const runtime = typeof getAppRuntime === 'function'
            ? getAppRuntime()
            : (window.appRuntime || {});

        const runtimeSync = runtime.sync || (runtime.sync = {});
        const syncReason = reason || 'manual';

        if (!navigator.onLine) {
            return {
                ok: false,
                role: 'Kooperant',
                synced: 0,
                failed: 0,
                results: [],
                reason: 'offline',
                code: 'OFFLINE',
                partial: false
            };
        }

        if (runtimeSync.kooperantRequestInFlight || runtimeSync.tretmaniInFlight || runtimeSync.troskoviInFlight) {
            runtimeSync.kooperantSyncRequested = true;
            runtimeSync.kooperantSyncRequestedReason = syncReason;

            return {
                ok: true,
                role: 'Kooperant',
                synced: 0,
                failed: 0,
                results: [],
                reason: 'already-running',
                code: 'ALREADY_RUNNING',
                partial: false
            };
        }

        runtimeSync.kooperantRequestInFlight = true;
        runtimeSync.kooperantLastSyncReason = syncReason;

        try {
            let result = await window.syncKooperantNow();

            // Ako je stigao novi zahtev dok je prvi sync trajao, uradi još jedan serijski pass.
            if (runtimeSync.kooperantSyncRequested) {
                runtimeSync.kooperantSyncRequested = false;
                runtimeSync.kooperantLastSyncReason = runtimeSync.kooperantSyncRequestedReason || 'requested';
                runtimeSync.kooperantSyncRequestedReason = '';

                result = await window.syncKooperantNow();
            }

            return result;
        } finally {
            runtimeSync.kooperantRequestInFlight = false;

            try {
                if (typeof updateSyncBadge === 'function') await updateSyncBadge();
            } catch (_) {}
        }
    };
    
    function normalizeKooperantSyncUnit(type, result) {
        const r = result && typeof result === 'object' ? result : {};

        const results = Array.isArray(r.results) ? r.results : [];

        let synced = Number.isFinite(r.synced)
            ? r.synced
            : results.filter(x => x && (x.success === true || x.status === 'synced')).length;

        let failed = Number.isFinite(r.failed)
            ? r.failed
            : results.filter(x => x && x.success === false).length;

        if (!Number.isFinite(synced)) synced = 0;
        if (!Number.isFinite(failed)) failed = 0;

        return {
            type,
            ok: typeof r.ok === 'boolean' ? r.ok : failed === 0,
            synced,
            failed,
            results,
            reason: r.reason || r.error || '',
            code: r.code || '',
            partial: !!r.partial || (synced > 0 && failed > 0)
        };
    }

    function aggregateKooperantSyncResult(unitResults) {
        const synced = unitResults.reduce((sum, r) => sum + (parseInt(r.synced, 10) || 0), 0);
        const failed = unitResults.reduce((sum, r) => sum + (parseInt(r.failed, 10) || 0), 0);

        const blockingFailures = unitResults.filter(r =>
            r.ok === false &&
            r.reason !== 'no-pending' &&
            r.reason !== 'feature-disabled'
        );

        const featureDisabled = unitResults.filter(r => r.reason === 'feature-disabled');

        let reason = '';
        let code = '';

        if (blockingFailures.length > 0) {
            reason = blockingFailures[0].reason || 'partial-failure';
            code = blockingFailures[0].code || '';
        } else if (featureDisabled.length > 0 && synced === 0) {
            reason = 'feature-disabled';
            code = featureDisabled[0].code || 'FEATURE_DISABLED';
        } else if (unitResults.every(r => r.reason === 'no-pending' || r.reason === 'feature-disabled')) {
            reason = 'no-pending';
        }

        return {
            ok: blockingFailures.length === 0,
            role: 'Kooperant',
            synced,
            failed,
            results: unitResults,
            reason,
            code,
            partial: blockingFailures.length > 0 || unitResults.some(r => r.partial)
        };
    }
})();
