// ============================================================
// OTKUPAC SYNC
// Thin wrapper around the shared sync engine (utils/sync-engine.js).
// ============================================================

async function syncQueue() {
    await updateSyncBadge('syncing');

    const result = await syncStore({
        storeName: CONFIG.STORE_NAME,
        action: 'sync',
        inFlightKey: 'otkupacInFlight',
        entityIdField: 'otkupacID',
        successLabel: 'Sinhronizovano'
    });

    try { await updateSyncBadge(); } catch (_) {}
    try { await updateStats(); } catch (_) {}
    try { await renderQueueList(); } catch (_) {}

    return result;
}

async function syncNow() {
    if (!navigator.onLine) {
        showToast('Nema konekcije', 'error');
        return buildSyncResult({ reason: 'offline' });
    }

    const runtimeSync = (window.appRuntime || {}).sync || {};
    if (runtimeSync.otkupacInFlight) {
        showToast('Sinhronizacija je već u toku', 'info');
        return buildSyncResult({ reason: 'already-running' });
    }

    return await syncQueue();
}

async function updateSyncBadge(status) {
    const badge = byId('syncBadge');
    if (!badge || !db) return;

    if (status === 'syncing') {
        setText(badge, 'SYNC...');
        badge.className = 'sync-badge sync-pending';
        return;
    }

    let pending = [];
    let syncing = [];

    try {
        pending = await dbGetByIndex(db, CONFIG.STORE_NAME, 'syncStatus', 'pending');
    } catch (e) {
        pending = [];
    }

    try {
        syncing = await dbGetByIndex(db, CONFIG.STORE_NAME, 'syncStatus', 'syncing');
    } catch (e) {
        syncing = [];
    }

    const waitCount = (pending?.length || 0) + (syncing?.length || 0);

    if (!navigator.onLine) {
        setText(badge, 'OFFLINE' + (waitCount > 0 ? ' (' + waitCount + ')' : ''));
        badge.className = 'sync-badge sync-offline';
    } else if ((syncing?.length || 0) > 0 || isAnySyncInFlight()) {
        setText(badge, 'SYNC...');
        badge.className = 'sync-badge sync-pending';
    } else if (waitCount > 0) {
        setText(badge, 'ČEKA: ' + waitCount);
        badge.className = 'sync-badge sync-pending';
    } else {
        setText(badge, 'ONLINE');
        badge.className = 'sync-badge sync-online';
    }
}

async function updateStats() {
    if (!db) return;

    const statPending = document.getElementById('statPending');
    const statSynced = document.getElementById('statSynced');

    if (!statPending || !statSynced) return;

    try {
        const all = await dbGetAll(db, CONFIG.STORE_NAME);
        const today = new Date().toISOString().split('T')[0];
        const t = (all || []).filter(r => r.datum === today);

        statPending.textContent = t.filter(r =>
            r.syncStatus === 'pending' || r.syncStatus === 'syncing'
        ).length;

        statSynced.textContent = t.filter(r =>
            r.syncStatus === 'synced'
        ).length;
    } catch (err) {
        console.error('updateStats failed:', err);
    }
}

async function renderQueueList() {
    const list = byId('queueList');
    if (!list || !db) return;

    try {
        const pending = await dbGetByIndex(db, CONFIG.STORE_NAME, 'syncStatus', 'pending');
        const syncing = await dbGetByIndex(db, CONFIG.STORE_NAME, 'syncStatus', 'syncing');
        const items = [...(syncing || []), ...(pending || [])];

        if (items.length === 0) {
            setHtml(list, '<p style="text-align:center;color:var(--text-muted);padding:40px;">Nema stavki za sinhronizaciju</p>');
            return;
        }

        setHtml(list, items.map(r => `
            <div class="queue-item">
                <div class="qi-header">
                    <span class="qi-koop">${escapeHtml(r.kooperantName || '')}</span>
                    <span class="qi-time">${escapeHtml(formatQueueTime(r.createdAtClient))}</span>
                </div>
                <div class="qi-detail">
                    ${escapeHtml(r.vrstaVoca || '')}
                    ${escapeHtml(r.klasa || '')}
                    | ${escapeHtml(String(r.kolicina || 0))} kg × ${escapeHtml(String(r.cena || 0))} RSD
                </div>
                ${r.syncStatus === 'syncing'
                    ? '<div class="qi-status" style="font-size:12px;color:var(--text-muted);margin-top:6px;">Sinhronizacija u toku...</div>'
                    : ''}
                ${r.lastSyncError
                    ? '<div class="qi-status" style="font-size:12px;color:#b42318;margin-top:6px;">' + escapeHtml(r.lastSyncError) + '</div>'
                    : ''}
            </div>
        `).join(''));
    } catch (err) {
        console.error('renderQueueList failed:', err);
        setHtml(list, '<p style="text-align:center;color:var(--text-muted);padding:40px;">Greška pri učitavanju reda</p>');
    }
}

function formatQueueTime(value) {
    if (!value) return '';
    try {
        return new Date(value).toLocaleTimeString('sr-RS');
    } catch (e) {
        return '';
    }
}
