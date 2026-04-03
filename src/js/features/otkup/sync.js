async function syncQueue() {
    const pending = await dbGetByIndex(db, CONFIG.STORE_NAME, 'syncStatus', 'pending');
    if (pending.length === 0) return;

    updateSyncBadge('syncing');

    const json = await apiPost('sync', {
        otkupacID: CONFIG.OTKUPAC_ID,
        records: pending
    });

    if (json && json.success) {
        for (const r of pending) {
            r.syncStatus = 'synced';
            await dbPut(db, CONFIG.STORE_NAME, r);
        }
        showToast('Sinhr: ' + pending.length, 'success');
    } else if (json) {
        showToast('Greška: ' + (json.error || ''), 'error');
    }

    updateSyncBadge();
    updateStats();
}

async function syncAgromere() {
    const pending = await dbGetByIndex(db, CONFIG.AGRO_STORE, 'syncStatus', 'pending');
    if (pending.length === 0) return;

    const json = await apiPost('syncAgromere', {
        kooperantID: CONFIG.ENTITY_ID,
        records: pending
    });

    if (json && json.success) {
        for (const r of pending) {
            r.syncStatus = 'synced';
            await dbPut(db, CONFIG.AGRO_STORE, r);
        }
        showToast('Agromere sinhr: ' + pending.length, 'success');
    }
}

async function syncNow() {
    if (!navigator.onLine) { showToast('Nema konekcije', 'error'); return; }
    await syncQueue(); renderQueueList();
}
