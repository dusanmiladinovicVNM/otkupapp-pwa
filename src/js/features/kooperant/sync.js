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
