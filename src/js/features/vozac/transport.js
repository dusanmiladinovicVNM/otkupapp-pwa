async function loadVozacTransport() {
    const list = document.getElementById('transportList');
    if (!list) return;

    if (typeof initVozacStatus === 'function') initVozacStatus();

    // Fetch plans and zbirne in parallel
    const [planJson, zbirneJson] = await Promise.all([
        safeAsync(() => apiFetch('action=getVozacPlans'), 'Greška pri učitavanju planova'),
        safeAsync(() => apiFetch('action=getVozacZbirne'), 'Greška pri učitavanju transporta')
    ]);

    renderVozacPlans(planJson);

    let local = [];
    let server = [];

    try {
        local = await dbGetAll(db, 'zbirne');
    } catch (err) {
        console.error('loadVozacTransport local failed:', err);
    }

    if (zbirneJson && zbirneJson.success && Array.isArray(zbirneJson.records)) {
        server = zbirneJson.records.map(r => ({
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
            brojZbirne: r.BrojZbirne || '',
            otkupRecordIDs: r.OtkupRecordIDs || '',

            syncStatus: 'synced',
            lastSyncError: '',
            lastServerStatus: 'server'
        }));
    }

    const merged = mergeTransportZbirne(local, server);
    const cnt = document.getElementById('transportCount');

    if (merged.length === 0) {
        list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:24px 0;">Nema izvršenih transporta</p>';
        if (cnt) cnt.textContent = '0';
        return;
    }

    if (cnt) cnt.textContent = String(merged.length);

    list.innerHTML = merged.map(r => {
        const totalKg = (r.kolicinaKlI || 0) + (r.kolicinaKlII || 0);
        const isPending = r.syncStatus === 'pending' || r.syncStatus === 'syncing';

        const trMod = isPending ? 'tr--load' : '';
        const bdgClass = r.syncStatus === 'syncing'  ? 'tr__bdg--road' :
                         r.syncStatus === 'pending'   ? 'tr__bdg--pending' :
                                                        'tr__bdg--done';
        const bdgLabel = r.syncStatus === 'syncing'  ? 'Sync...' :
                         r.syncStatus === 'pending'   ? 'Na čekanju' :
                                                        'Sinhronizovano';
        const syncIcon = r.syncStatus === 'syncing'  ? '🔄' :
                         r.syncStatus === 'pending'   ? '⏳' : '✅';

        const fruitLine = [r.vrstaVoca, r.sortaVoca].filter(Boolean).join(' ');
        const detLines = [
            fruitLine ? escapeHtml(fruitLine) : '',
            r.kolAmbalaze ? 'Amb: <strong>' + r.kolAmbalaze + '</strong>' : '',
            r.brojZbirne ? escapeHtml(r.brojZbirne) : ''
        ].filter(Boolean).join(' · ');

        return `<div class="tr ${trMod}">
            <div class="tr__top">
                <div>
                    <div class="tr__dest">${escapeHtml(r.kupacName || r.kupacID || '—')}</div>
                    ${fruitLine ? `<div class="tr__from">${escapeHtml(fruitLine)}</div>` : ''}
                </div>
                <div class="tr__time">${escapeHtml(r.datum || '')}</div>
            </div>
            <div class="tr__main">
                <div class="tr__kg">${totalKg.toLocaleString('sr')}<span class="u"> kg</span></div>
                <div class="tr__det">${detLines}</div>
            </div>
            <div class="tr__bot">
                <span class="tr__bdg ${bdgClass}">${bdgLabel}</span>
                <span class="tr__sync">${syncIcon}</span>
            </div>
            ${r.lastSyncError ? `<div class="tr__err">${escapeHtml(r.lastSyncError)}</div>` : ''}
        </div>`;
    }).join('');
}

function renderVozacPlans(json) {
    const planiList = document.getElementById('vozacPlaniList');
    const planiCnt = document.getElementById('vozacPlaniCount');
    if (!planiList) return;

    const plans = (json && json.success && Array.isArray(json.records)) ? json.records : [];

    if (planiCnt) planiCnt.textContent = plans.length ? String(plans.length) : '0';

    if (plans.length === 0) {
        planiList.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:16px 0 4px;">Nema planiranih ruta za danas</p>';
        return;
    }

    const statusLabels = {
        planned: 'Planirano',
        u_toku: 'U toku'
    };

    planiList.innerHTML = plans.map(p => {
        const kg = parseInt(p.PlannedKg || 0);
        const st = String(p.Status || 'planned');
        const isActive = st === 'u_toku';
        const trMod = isActive ? 'tr--load' : '';
        const bdgClass = isActive ? 'tr__bdg--load' : 'tr__bdg--road';
        const bdgLabel = statusLabels[st] || escapeHtml(st);

        const route = [escapeHtml(p.StanicaName || p.StanicaID || ''), escapeHtml(p.KupacName || p.KupacID || '')].filter(Boolean).join(' → ');

        return `<div class="tr ${trMod}">
            <div class="tr__top">
                <div>
                    <div class="tr__dest">${escapeHtml(p.KupacName || p.KupacID || '—')}</div>
                    <div class="tr__from">sa stanice ${escapeHtml(p.StanicaName || p.StanicaID || '?')}</div>
                </div>
                <div class="tr__time">${isActive ? '▶' : '·'}</div>
            </div>
            <div class="tr__main">
                <div class="tr__kg">${kg.toLocaleString('sr')}<span class="u"> kg</span></div>
                <div class="tr__det">${route}</div>
            </div>
            <div class="tr__bot">
                <span class="tr__bdg ${bdgClass}">${bdgLabel}</span>
                <span class="tr__sync">${isActive ? '🚛' : '📋'}</span>
            </div>
        </div>`;
    }).join('');
}

function mergeTransportZbirne(local, server) {
    return mergeOfflineRecords(local, server, normalizeLocalZbirnaRecord);
}
