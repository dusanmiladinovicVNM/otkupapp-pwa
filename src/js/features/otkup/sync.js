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

async function updateStats() {
    const all = await dbGetAll(db, CONFIG.STORE_NAME);
    const today = new Date().toISOString().split('T')[0];
    const t = all.filter(r => r.datum === today);
    document.getElementById('statPending').textContent = t.filter(r => r.syncStatus === 'pending').length;
    document.getElementById('statSynced').textContent = t.filter(r => r.syncStatus === 'synced').length;
}

async function renderQueueList() {
    const pending = await dbGetByIndex(db, CONFIG.STORE_NAME, 'syncStatus', 'pending');
    const list = byId('queueList');
    if (!list) return;

    if (pending.length === 0) {
        setHtml(list, '<p style="text-align:center;color:var(--text-muted);padding:40px;">Nema stavki za sinhronizaciju</p>');
        return;
    }

    setHtml(list, pending.map(r =>
        `<div class="queue-item"><div class="qi-header"><span class="qi-koop">${escapeHtml(r.kooperantName)}</span><span class="qi-time">${escapeHtml(new Date(r.createdAtClient).toLocaleTimeString('sr'))}</span></div>
            <div class="qi-detail">${escapeHtml(r.vrstaVoca)} ${escapeHtml(r.klasa)} | ${r.kolicina} kg × ${r.cena} RSD</div></div>`
    ).join(''));
}

