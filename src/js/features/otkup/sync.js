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

async function syncNow() {
    if (!navigator.onLine) { showToast('Nema konekcije', 'error'); return; }
    await syncQueue(); renderQueueList();
}

async function updateSyncBadge(status) {
    const badge = byId('syncBadge');
    if (!badge) return;
    if (status === 'syncing') {
        setText(badge, 'SYNC...');
        badge.className = 'sync-badge sync-pending';
        return;
    }
    const pending = await dbGetByIndex(db, CONFIG.STORE_NAME, 'syncStatus', 'pending');
    if (!navigator.onLine) {
        setText(badge, 'OFFLINE' + (pending.length > 0 ? ' (' + pending.length + ')' : ''));
        badge.className = 'sync-badge sync-offline';
    } else if (pending.length > 0) {
        setText(badge, 'ČEKA: ' + pending.length);
        badge.className = 'sync-badge sync-pending';
    } else {
        setText(badge, 'ONLINE');
        badge.className = 'sync-badge sync-online';
    }
}
