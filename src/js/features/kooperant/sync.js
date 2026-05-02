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
    //
    // NOTE: GAS endpoint currently returns FEATURE_DISABLED.
    // Engine handles this gracefully — records stay 'pending'
    // without polluting attempt counters or error fields.
    // When GAS implements the endpoint, this will start working
    // automatically with no frontend changes required.
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
        const results = [];

        const syncs = [
            ['tretmani', window.syncTretmani],
            ['troskovi', window.syncTroskovi]
        ];

        for (const [type, fn] of syncs) {
            try {
                results.push({ type, ...(await fn()) });
            } catch (e) {
                results.push({
                    type,
                    ok: false,
                    error: (e && e.message) || (type + ' sync failed')
                });
            }
        }

        return results;
    };
})();
