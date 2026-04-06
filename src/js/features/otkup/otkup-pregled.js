// ============================================================
// OTKUP PREGLED
// ============================================================
async function loadOtkupPregled() {
    const list = document.getElementById('pregledList');
    const fldOd = document.getElementById('fldPregledOd');
    const fldDo = document.getElementById('fldPregledDo');

    if (!list) return;

    list.innerHTML = '<p style="text-align:center;padding:20px;color:var(--text-muted);">Učitavanje...</p>';

    const od = fldOd ? fldOd.value : '';
    const doo = fldDo ? fldDo.value : '';

    let local = [];
    let server = [];

    try {
        if (db) {
            local = await dbGetAll(db, CONFIG.STORE_NAME);
        }
    } catch (err) {
        console.error('loadOtkupPregled local failed:', err);
    }

    if (navigator.onLine) {
        const json = await safeAsync(async () => {
            return await apiFetch('action=getOtkupi&otkupacID=' + encodeURIComponent(CONFIG.OTKUPAC_ID));
        }, 'Greška pri učitavanju otkupa');

        if (json && json.success && Array.isArray(json.records)) {
            server = json.records.map(mapServerOtkupRecord);
        }
    }

    let all = mergeOtkupPregledRecords(local, server);

    if (od) all = all.filter(r => (r.datum || '') >= od);
    if (doo) all = all.filter(r => (r.datum || '') <= doo);

    all.sort(compareOtkupPregledRecordsDesc);

    renderOtkupPregledStats(all);
    renderOtkupPregledList(list, all);
}

function mapServerOtkupRecord(r) {
    return {
        clientRecordID: r.ClientRecordID || '',
        serverRecordID: r.ServerRecordID || '',
        createdAtClient: normalizeIso(r.CreatedAtClient),
        updatedAtClient: normalizeIso(r.UpdatedAtClient || r.CreatedAtClient),
        updatedAtServer: normalizeIso(r.UpdatedAtServer || r.ReceivedAt),
        syncedAt: normalizeIso(r.UpdatedAtServer || r.ReceivedAt),

        datum: fmtDate(r.Datum),
        kooperantID: r.KooperantID || '',
        kooperantName: r.KooperantName || r.KooperantID || '',
        vrstaVoca: r.VrstaVoca || '',
        sortaVoca: r.SortaVoca || '',
        klasa: r.Klasa || 'I',
        kolicina: parseFloat(r.Kolicina) || 0,
        cena: parseFloat(r.Cena) || 0,
        kolAmbalaze: parseInt(r.KolAmbalaze, 10) || 0,
        parcelaID: r.ParcelaID || '',
        vozacID: r.VozacID || r.VozaciID || '',
        napomena: r.Napomena || '',

        syncStatus: 'synced',
        syncAttempts: 0,
        lastSyncError: '',
        lastServerStatus: 'server'
    };
}

function mergeOtkupPregledRecords(local, server) {
    return mergeOfflineRecords(local, server, normalizeLocalPregledRecord);
}

function normalizeLocalPregledRecord(r) {
    return {
        clientRecordID: r.clientRecordID || '',
        serverRecordID: r.serverRecordID || '',
        createdAtClient: normalizeIso(r.createdAtClient),
        updatedAtClient: normalizeIso(r.updatedAtClient || r.createdAtClient),
        updatedAtServer: normalizeIso(r.updatedAtServer),
        syncedAt: normalizeIso(r.syncedAt),

        datum: r.datum || '',
        kooperantID: r.kooperantID || '',
        kooperantName: r.kooperantName || r.kooperantID || '',
        vrstaVoca: r.vrstaVoca || '',
        sortaVoca: r.sortaVoca || '',
        klasa: r.klasa || 'I',
        kolicina: parseFloat(r.kolicina) || 0,
        cena: parseFloat(r.cena) || 0,
        kolAmbalaze: parseInt(r.kolAmbalaze, 10) || 0,
        parcelaID: r.parcelaID || '',
        vozacID: r.vozacID || '',
        napomena: r.napomena || '',

        syncStatus: r.syncStatus || 'pending',
        syncAttempts: parseInt(r.syncAttempts, 10) || 0,
        lastSyncError: r.lastSyncError || '',
        lastServerStatus: r.lastServerStatus || ''
    };
}

function compareOtkupPregledRecordsDesc(a, b) {
    const aTime = a.updatedAtClient || a.createdAtClient || a.updatedAtServer || '';
    const bTime = b.updatedAtClient || b.createdAtClient || b.updatedAtServer || '';

    const byDate = (b.datum || '').localeCompare(a.datum || '');
    if (byDate !== 0) return byDate;

    const byTime = String(bTime).localeCompare(String(aTime));
    if (byTime !== 0) return byTime;

    return String(b.clientRecordID || '').localeCompare(String(a.clientRecordID || ''));
}

function renderOtkupPregledStats(all) {
    const countEl = document.getElementById('statPregledCount');
    const kgEl = document.getElementById('statPregledKg');
    const vrednostEl = document.getElementById('statPregledVrednost');
    const koopEl = document.getElementById('statPregledKoop');

    const kg = all.reduce((s, r) => s + (parseFloat(r.kolicina) || 0), 0);
    const vr = all.reduce((s, r) => s + ((parseFloat(r.kolicina) || 0) * (parseFloat(r.cena) || 0)), 0);
    const koopCount = new Set(all.map(r => r.kooperantID).filter(Boolean)).size;

    if (countEl) countEl.textContent = String(all.length);
    if (kgEl) kgEl.textContent = kg.toLocaleString('sr-RS');
    if (vrednostEl) vrednostEl.textContent = vr.toLocaleString('sr-RS');
    if (koopEl) koopEl.textContent = String(koopCount);
}

function renderOtkupPregledList(list, all) {
    if (!all.length) {
        list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:40px;">Nema otkupa</p>';
        return;
    }

    list.innerHTML = all.map(r => {
        const vrednost = ((parseFloat(r.kolicina) || 0) * (parseFloat(r.cena) || 0)).toLocaleString('sr-RS');
        const meta = buildPregledStatusMeta(r);

        return `
            <div class="queue-item" style="border-left-color:${meta.color};">
                <div class="qi-header">
                    <span class="qi-koop">${escapeHtml(r.kooperantName || '')}</span>
                    <span class="qi-time">${escapeHtml(r.datum || '')}</span>
                </div>
                <div class="qi-detail">
                    ${escapeHtml(r.vrstaVoca || '')}
                    ${escapeHtml(r.sortaVoca || '')}
                    ${escapeHtml(r.klasa || '')}
                    | ${escapeHtml(String(r.kolicina || 0))} kg × ${escapeHtml(String(r.cena || 0))}
                    = <strong>${escapeHtml(vrednost)} RSD</strong>
                    ${r.parcelaID ? ' | ' + escapeHtml(r.parcelaID) : ''}
                </div>
                <div style="margin-top:6px;font-size:12px;color:var(--text-muted);">
                    ${escapeHtml(meta.label)}
                    ${r.lastSyncError ? ' | ' + escapeHtml(r.lastSyncError) : ''}
                </div>
            </div>
        `;
    }).join('');
}

function buildPregledStatusMeta(r) {
    if (r.syncStatus === 'syncing') {
        return {
            color: 'var(--warning)',
            label: 'Sinhronizacija u toku'
        };
    }

    if (r.syncStatus === 'pending') {
        return {
            color: 'var(--warning)',
            label: 'Čeka sinhronizaciju'
        };
    }

    return {
        color: 'var(--success)',
        label: r.serverRecordID ? ('Sinhronizovano • ' + r.serverRecordID) : 'Sinhronizovano'
    };
}
