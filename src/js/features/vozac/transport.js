async function loadVozacTransport() {
    const list = document.getElementById('transportList');
    if (!list) return;

    let local = [];
    let server = [];

    try {
        local = await dbGetAll(db, 'zbirne');
    } catch (err) {
        console.error('loadVozacTransport local failed:', err);
    }

    const json = await safeAsync(async () => {
        return await apiFetch('action=getVozacZbirne');
    }, 'Greška pri učitavanju transporta');

    if (json && json.success && Array.isArray(json.records)) {
        server = json.records.map(r => ({
            clientRecordID: r.ClientRecordID || '',
            serverRecordID: r.ServerRecordID || '',
            createdAtClient: r.CreatedAtClient || '',
            updatedAtClient: r.UpdatedAtClient || r.CreatedAtClient || '',
            updatedAtServer: r.UpdatedAtServer || r.ReceivedAt || '',

            datum: fmtDate(r.Datum),
            kupacID: r.KupacID || '',
            kupacName: r.KupacName || r.KupacID || '',
            vrstaVoca: r.VrstaVoca || '',
            sortaVoca: r.SortaVoca || '',
            kolicinaKlI: parseFloat(r.KolicinaKlI) || 0,
            kolicinaKlII: parseFloat(r.KolicinaKlII) || 0,
            kolAmbalaze: parseInt(r.KolAmbalaze, 10) || 0,
            tipAmbalaze: r.TipAmbalaze || '',
            klasa: r.Klasa || '',
            otkupRecordIDs: r.OtkupRecordIDs || '',

            syncStatus: 'synced',
            lastSyncError: '',
            lastServerStatus: 'server'
        }));
    }

    const merged = mergeTransportZbirne(local, server);

    if (merged.length === 0) {
        list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema transporta</p>';
        return;
    }

    list.innerHTML = merged.map(r => {
        const totalKg = (r.kolicinaKlI || 0) + (r.kolicinaKlII || 0);
        const syncIcon =
            r.syncStatus === 'syncing' ? '🔄' :
            r.syncStatus === 'pending' ? '⏳' : '✅';

        return `<div class="queue-item">
            <div class="qi-header">
                <span class="qi-koop">🏭 ${escapeHtml(r.kupacName || r.kupacID || '')}</span>
                <span class="qi-time">${escapeHtml(r.datum || '')}</span>
            </div>
            <div class="qi-detail">
                ${totalKg.toLocaleString('sr')} kg | Amb: ${r.kolAmbalaze || 0} | ${syncIcon}
            </div>
            ${r.serverRecordID
                ? `<div class="qi-detail" style="font-size:11px;color:var(--text-muted);">${escapeHtml(r.serverRecordID)}</div>`
                : ''}
            ${r.lastSyncError
                ? `<div class="qi-detail" style="font-size:11px;color:#b42318;">${escapeHtml(r.lastSyncError)}</div>`
                : ''}
        </div>`;
    }).join('');
}

function mergeTransportZbirne(local, server) {
    return mergeOfflineRecords(local, server, normalizeLocalZbirnaRecord);
}
