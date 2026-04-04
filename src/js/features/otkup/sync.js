// ============================================================
// SYNC
// ============================================================

async function syncQueue() {
    if (!db) {
        console.warn('syncQueue: db not ready');
        return { ok: false, reason: 'db-not-ready' };
    }

    if (!navigator.onLine) {
        updateSyncBadge();
        return { ok: false, reason: 'offline' };
    }

    if (window.appRuntime && window.appRuntime.syncInFlight) {
        return { ok: false, reason: 'already-running' };
    }

    if (window.appRuntime) {
        window.appRuntime.syncInFlight = true;
    }

    let pending = [];

    try {
        pending = await dbGetByIndex(db, CONFIG.STORE_NAME, 'syncStatus', 'pending');

        if (!Array.isArray(pending) || pending.length === 0) {
            await updateSyncBadge();
            await updateStats();
            return { ok: true, synced: 0, failed: 0 };
        }

        await updateSyncBadge('syncing');

        // mark as syncing before request
        for (const record of pending) {
            record.syncStatus = 'syncing';
            record.syncAttemptAt = new Date().toISOString();
            await dbPut(db, CONFIG.STORE_NAME, record);
        }

        const json = await apiPost('sync', {
            otkupacID: CONFIG.OTKUPAC_ID,
            records: pending
        });

        // network / hard failure
        if (!json || json.success === false) {
            for (const record of pending) {
                record.syncStatus = 'pending';
                record.lastSyncError = json && json.error ? String(json.error) : 'Sync neuspešan';
                await dbPut(db, CONFIG.STORE_NAME, record);
            }

            if (json && json.error) {
                showToast('Greška: ' + json.error, 'error');
            } else {
                showToast('Sinhronizacija nije uspela', 'error');
            }

            return { ok: false, synced: 0, failed: pending.length };
        }

        // --------------------------------------------
        // 1) NEW BACKEND MODE:
        // backend can return per-record results
        // expected examples:
        // { success:true, results:[{clientRecordID:'x', success:true, serverRecordID:123}] }
        // { success:true, synced:[...ids], failed:[...ids] }
        // --------------------------------------------
        let syncedCount = 0;
        let failedCount = 0;

        if (Array.isArray(json.results)) {
            const byClientId = new Map(
                pending.map(r => [r.clientRecordID, r])
            );

            for (const result of json.results) {
                const record = byClientId.get(result.clientRecordID);
                if (!record) continue;

                if (result.success) {
                    record.syncStatus = 'synced';
                    record.lastSyncError = '';
                    record.syncedAt = new Date().toISOString();

                    if (result.serverRecordID) {
                        record.serverRecordID = result.serverRecordID;
                    }
                    if (result.updatedAtServer) {
                        record.updatedAtServer = result.updatedAtServer;
                    }

                    syncedCount++;
                } else {
                    record.syncStatus = 'pending';
                    record.lastSyncError = result.error || 'Sync stavke neuspešan';
                    failedCount++;
                }

                await dbPut(db, CONFIG.STORE_NAME, record);
            }

            // any pending record not mentioned by backend -> return to pending
            for (const record of pending) {
                const mentioned = json.results.some(x => x.clientRecordID === record.clientRecordID);
                if (!mentioned) {
                    record.syncStatus = 'pending';
                    record.lastSyncError = 'Nema potvrde sa servera';
                    failedCount++;
                    await dbPut(db, CONFIG.STORE_NAME, record);
                }
            }
        }

        // --------------------------------------------
        // 2) ALT BACKEND MODE:
        // { success:true, synced:["id1"], failed:["id2"] }
        // --------------------------------------------
        else if (Array.isArray(json.synced) || Array.isArray(json.failed)) {
            const syncedSet = new Set(Array.isArray(json.synced) ? json.synced : []);
            const failedSet = new Set(Array.isArray(json.failed) ? json.failed : []);

            for (const record of pending) {
                if (syncedSet.has(record.clientRecordID)) {
                    record.syncStatus = 'synced';
                    record.lastSyncError = '';
                    record.syncedAt = new Date().toISOString();
                    syncedCount++;
                } else if (failedSet.has(record.clientRecordID)) {
                    record.syncStatus = 'pending';
                    record.lastSyncError = 'Server odbio zapis';
                    failedCount++;
                } else {
                    record.syncStatus = 'pending';
                    record.lastSyncError = 'Nema potvrde sa servera';
                    failedCount++;
                }

                await dbPut(db, CONFIG.STORE_NAME, record);
            }
        }

        // --------------------------------------------
        // 3) LEGACY BACKEND MODE:
        // { success:true } => treat all as synced
        // --------------------------------------------
        else {
            for (const record of pending) {
                record.syncStatus = 'synced';
                record.lastSyncError = '';
                record.syncedAt = new Date().toISOString();
                await dbPut(db, CONFIG.STORE_NAME, record);
                syncedCount++;
            }
        }

        if (syncedCount > 0 && failedCount === 0) {
            showToast('Sinhronizovano: ' + syncedCount, 'success');
        } else if (syncedCount > 0 && failedCount > 0) {
            showToast('Sinhronizovano: ' + syncedCount + ', neuspešno: ' + failedCount, 'info');
        } else if (failedCount > 0) {
            showToast('Nijedna stavka nije sinhronizovana', 'error');
        }

        return { ok: failedCount === 0, synced: syncedCount, failed: failedCount };
    } catch (err) {
        console.error('syncQueue failed:', err);

        // rollback syncing -> pending
        for (const record of pending) {
            try {
                if (record.syncStatus === 'syncing') {
                    record.syncStatus = 'pending';
                    record.lastSyncError = err && err.message ? err.message : 'Greška pri sync-u';
                    await dbPut(db, CONFIG.STORE_NAME, record);
                }
            } catch (rollbackErr) {
                console.error('sync rollback failed:', rollbackErr);
            }
        }

        showToast('Greška pri sinhronizaciji', 'error');
        return { ok: false, synced: 0, failed: pending.length || 0 };
    } finally {
        if (window.appRuntime) {
            window.appRuntime.syncInFlight = false;
        }

        try { await updateSyncBadge(); } catch (_) {}
        try { await updateStats(); } catch (_) {}
        try {
            if (typeof renderQueueList === 'function') {
                await renderQueueList();
            }
        } catch (_) {}
    }
}

async function syncNow() {
    if (!navigator.onLine) {
        showToast('Nema konekcije', 'error');
        return;
    }

    if (window.appRuntime && window.appRuntime.syncInFlight) {
        showToast('Sinhronizacija je već u toku', 'info');
        return;
    }

    await syncQueue();
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
    } else if ((syncing?.length || 0) > 0 || (window.appRuntime && window.appRuntime.syncInFlight)) {
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
