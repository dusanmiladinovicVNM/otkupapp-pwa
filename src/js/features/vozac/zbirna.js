// ============================================================
// VOZAC: ZBIRNA
// ============================================================
let vozacOtkupi = [];
let _lastMergedZbirne = null;

async function loadVozacData() {
    vozacOtkupi = [];

    const json = await safeAsync(async () => {
        return await apiFetch('action=getVozacOtkupi');
    }, 'Greška pri učitavanju podataka vozača');

    if (json && json.success && Array.isArray(json.records)) {
        vozacOtkupi = json.records.map(r => ({
            clientRecordID: r.ClientRecordID || '',
            serverRecordID: r.ServerRecordID || '',
            datum: fmtDate(r.Datum),
            kooperantName: r.KooperantName || r.KooperantID || '',
            kooperantID: r.KooperantID || '',
            vrstaVoca: r.VrstaVoca || '',
            sortaVoca: r.SortaVoca || '',
            klasa: r.Klasa || 'I',
            kolicina: parseFloat(r.Kolicina) || 0,
            cena: parseFloat(r.Cena) || 0,
            tipAmbalaze: r.TipAmbalaze || '',
            kolAmbalaze: parseInt(r.KolAmbalaze, 10) || 0,
            stanicaID: r.OtkupacID || extractStanicaIdFromSource(r._source) || '',
            vozacID: r.VozacID || '',
            updatedAtServer: r.UpdatedAtServer || r.ReceivedAt || '',
            syncStatus: 'synced'
        }));
    }

    // Jedan fetch za zbirne — koristi se i za filter i za renderovanje
    const zbirne = await getMergedZbirneForVozac();
    _lastMergedZbirne = zbirne;

    const consumedIds = getConsumedOtkupIdsFromZbirne(zbirne);
    vozacOtkupi = vozacOtkupi.filter(r => !consumedIds.has(r.clientRecordID));

    renderVozacOtpremnice();
    renderVozacZbirneFromData(zbirne);
}

async function loadVozacZbirne() {
    // Standalone poziv — kad se zove van loadVozacData (npr. posle cancelZbirna)
    const zbirne = await getMergedZbirneForVozac();
    _lastMergedZbirne = zbirne;
    renderVozacZbirneFromData(zbirne);
}

function extractStanicaIdFromSource(source) {
    const s = String(source || '');
    if (s.startsWith('OTK-')) return s.substring(4);
    return s;
}

function renderVozacOtpremnice() {
    const today = getTodayIsoDate();
    const todayOtkupi = vozacOtkupi.filter(r => r.datum === today);

    const list = document.getElementById('vozacOtpremniceList');
    if (!list) return;

    if (todayOtkupi.length === 0) {
        list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema otpremnica za danas</p>';
        const btn = document.getElementById('btnNovaZbirna');
        if (btn) btn.style.display = 'none';
        return;
    }

    const btn = document.getElementById('btnNovaZbirna');
    if (btn) btn.style.display = '';

    // Group by stanica
    const grouped = {};
    todayOtkupi.forEach(r => {
        const s = r.stanicaID || '?';
        if (!grouped[s]) grouped[s] = { items: [], kg: 0, amb: 0 };
        grouped[s].items.push(r);
        grouped[s].kg += r.kolicina || 0;
        grouped[s].amb += r.kolAmbalaze || 0;
    });

    list.innerHTML = Object.entries(grouped).map(([sta, g]) =>
        `<div class="vblk-sta">${escapeHtml(fmtStanica(sta))} · ${g.kg.toLocaleString('sr')} kg · ${g.amb} amb</div>` +
        g.items.map(r =>
            `<div class="vblk sel">
                <div class="vblk__c">✓</div>
                <div>
                    <div class="vblk__nm">${escapeHtml(r.kooperantName)}</div>
                    <div class="vblk__s">${escapeHtml(r.vrstaVoca)} ${escapeHtml(r.klasa)} · ${r.kolAmbalaze} amb</div>
                </div>
                <div class="vblk__kg">${(r.kolicina || 0).toLocaleString('sr')}<span class="u"> kg</span></div>
            </div>`
        ).join('')
    ).join('');
}

async function startZbirnaCreation() {
    document.getElementById('zbirnaMainView').style.display = 'none';
    document.getElementById('zbirnaCreateView').style.display = 'block';
    
    const sel = document.getElementById('fldZbirnaKupac');
    sel.innerHTML = '<option value="">-- Izaberi kupca --</option>';
    
    // Populate from stammdaten
    (stammdaten.kupci || []).forEach(k => {
        const o = document.createElement('option');
        o.value = k.KupacID;
        o.textContent = k.Naziv + ' (' + k.KupacID + ')';
        sel.appendChild(o);
    });
    
    // Optional fallback from mgmtData
    if (mgmtData && mgmtData.saldoKupci) {
        mgmtData.saldoKupci.forEach(k => {
            const value = k.KupacID || k.Kupac;
            if (!value) return;

            const exists = Array.from(sel.options).some(opt => opt.value === value);
            if (exists) return;

            const o = document.createElement('option');
            o.value = value;
            o.textContent = k.Kupac || k.KupacID;
            sel.appendChild(o);
        });
    }
    
    renderZbirnaSummary();
}

function renderZbirnaSummary() {
    const today = getTodayIsoDate();
    const todayOtkupi = vozacOtkupi.filter(r => r.datum === today);

    let totalKgI = 0, totalKgII = 0, totalAmb = 0;
    todayOtkupi.forEach(r => {
        if (r.klasa === 'II') totalKgII += r.kolicina || 0;
        else totalKgI += r.kolicina || 0;
        totalAmb += r.kolAmbalaze || 0;
    });

    const totalKg = totalKgI + totalKgII;
    const stanicaCount = new Set(todayOtkupi.map(r => r.stanicaID)).size;

    const summaryEl = document.getElementById('zbirnaOtkupiSummary');
    if (summaryEl) {
        summaryEl.innerHTML = `<div class="vsum">
            <div class="vsum__total">${totalKg.toLocaleString('sr')}<span class="u"> kg</span></div>
            <div class="vsum__detail">
                ${totalKgII > 0
                    ? 'Kl. I: ' + totalKgI.toLocaleString('sr') + ' kg · Kl. II: ' + totalKgII.toLocaleString('sr') + ' kg'
                    : 'Klasa I'}
                · Amb: ${totalAmb}
            </div>
            <div class="vsum__sub">${todayOtkupi.length} otkupa · ${stanicaCount} ${stanicaCount === 1 ? 'stanica' : 'stanice'}</div>
        </div>`;
    }

    const countEl = document.getElementById('zbirnaOtkupiCount');
    if (countEl) countEl.textContent = String(todayOtkupi.length);

    // List individual otkupi as vblk cards (read-only, all selected)
    const listEl = document.getElementById('zbirnaOtkupiList');
    if (listEl) {
        listEl.innerHTML = todayOtkupi.map(r =>
            `<div class="vblk sel">
                <div class="vblk__c">✓</div>
                <div>
                    <div class="vblk__nm">${escapeHtml(r.kooperantName)}</div>
                    <div class="vblk__s">${escapeHtml(r.vrstaVoca)} ${escapeHtml(r.klasa)} · ${escapeHtml(fmtStanica(r.stanicaID))}</div>
                </div>
                <div class="vblk__kg">${(r.kolicina || 0).toLocaleString('sr')}<span class="u"> kg</span></div>
            </div>`
        ).join('');
    }
}

function mapServerZbirnaRecord(r) {
    return {
        clientRecordID: r.ClientRecordID || '',
        serverRecordID: r.ServerRecordID || '',
        brojZbirne: r.BrojZbirne || '',
        createdAtClient: normalizeIso(r.CreatedAtClient),
        updatedAtClient: normalizeIso(r.UpdatedAtClient || r.CreatedAtClient),
        updatedAtServer: normalizeIso(r.UpdatedAtServer || r.ReceivedAt),
        syncedAt: normalizeIso(r.UpdatedAtServer || r.ReceivedAt),

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
        syncAttempts: 0,
        lastSyncError: '',
        lastServerStatus: 'server'
    };
}

function normalizeLocalZbirnaRecord(r) {
    return {
        clientRecordID: r.clientRecordID || '',
        serverRecordID: r.serverRecordID || '',
        brojZbirne: r.brojZbirne || '',
        createdAtClient: normalizeIso(r.createdAtClient),
        updatedAtClient: normalizeIso(r.updatedAtClient || r.createdAtClient),
        updatedAtServer: normalizeIso(r.updatedAtServer),
        syncedAt: normalizeIso(r.syncedAt),

        datum: r.datum || '',
        kupacID: r.kupacID || '',
        kupacName: r.kupacName || r.kupacID || '',
        vrstaVoca: r.vrstaVoca || '',
        sortaVoca: r.sortaVoca || '',
        kolicinaKlI: parseFloat(r.kolicinaKlI) || 0,
        kolicinaKlII: parseFloat(r.kolicinaKlII) || 0,
        kolAmbalaze: parseInt(r.kolAmbalaze, 10) || 0,
        tipAmbalaze: r.tipAmbalaze || '',
        klasa: r.klasa || '',
        otkupRecordIDs: r.otkupRecordIDs || '',

        syncStatus: r.syncStatus || 'pending',
        syncAttempts: parseInt(r.syncAttempts, 10) || 0,
        lastSyncError: r.lastSyncError || '',
        lastServerStatus: r.lastServerStatus || ''
    };
}

function mergeZbirneRecords(local, server) {
    return mergeOfflineRecords(local, server, normalizeLocalZbirnaRecord);
}

function renderVozacZbirneFromData(allZbirne) {
    const all = (allZbirne || [])
        .filter(r => !r.deleted)
        .sort((a, b) => {
            const byDate = (b.datum || '').localeCompare(a.datum || '');
            if (byDate !== 0) return byDate;

            const byTime = String(b.updatedAtClient || b.createdAtClient || b.updatedAtServer || '')
                .localeCompare(String(a.updatedAtClient || a.createdAtClient || a.updatedAtServer || ''));
            if (byTime !== 0) return byTime;

            return String(b.clientRecordID || '').localeCompare(String(a.clientRecordID || ''));
        });

    const list = document.getElementById('vozacZbirneList');
    if (!list) return;

    if (all.length === 0) {
        list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:12px;">Nema kreiranih zbirnih</p>';
        return;
    }

    list.innerHTML = all.map(r => {
        const totalKg = (r.kolicinaKlI || 0) + (r.kolicinaKlII || 0);
        const isPending = r.syncStatus === 'pending' || r.syncStatus === 'syncing';
        const trMod = isPending ? 'tr--load' : '';
        const bdgClass = r.syncStatus === 'syncing' ? 'tr__bdg--road' :
                         r.syncStatus === 'pending'  ? 'tr__bdg--pending' :
                                                       'tr__bdg--done';
        const bdgLabel = r.syncStatus === 'syncing' ? 'Sync...' :
                         r.syncStatus === 'pending'  ? 'Na čekanju' :
                                                       'Sinhronizovano';
        const syncIcon = r.syncStatus === 'syncing' ? '🔄' :
                         r.syncStatus === 'pending'  ? '⏳' : '✅';

        const klaseLine = r.kolicinaKlII > 0
            ? 'Kl.I: ' + (r.kolicinaKlI || 0).toLocaleString('sr') + ' · Kl.II: ' + (r.kolicinaKlII || 0).toLocaleString('sr') + ' kg'
            : '';

        return `<div class="tr ${trMod}">
            <div class="tr__top">
                <div>
                    <div class="tr__dest">${escapeHtml(r.kupacName)}</div>
                    ${r.vrstaVoca ? `<div class="tr__from">${escapeHtml(r.vrstaVoca)}</div>` : ''}
                </div>
                <div class="tr__time">${escapeHtml(r.datum)}</div>
            </div>
            <div class="tr__main">
                <div class="tr__kg">${totalKg.toLocaleString('sr')}<span class="u"> kg</span></div>
                <div class="tr__det">
                    Amb: <strong>${r.kolAmbalaze || 0}</strong>
                    ${r.brojZbirne ? ' · ' + escapeHtml(r.brojZbirne) : ''}
                    ${klaseLine ? '<br>' + klaseLine : ''}
                </div>
            </div>
            <div class="tr__bot">
                <span class="tr__bdg ${bdgClass}">${bdgLabel}</span>
                <span class="tr__sync">${syncIcon}</span>
            </div>
            ${r.lastSyncError ? `<div class="tr__err">${escapeHtml(r.lastSyncError)}</div>` : ''}
        </div>`;
    }).join('');
}

function cancelZbirna() {
    document.getElementById('zbirnaCreateView').style.display = 'none';
    document.getElementById('zbirnaMainView').style.display = 'block';
    loadVozacZbirne();
}

async function confirmZbirna() {
    if (typeof withSubmitLock !== 'function') {
        if (typeof window.ensureMasterSyncNotActive === 'function') {
            const allowed = await window.ensureMasterSyncNotActive('confirmZbirna', {
                showToast: true
            });

            if (!allowed) return;
        }

        return confirmZbirnaUnlocked();
    }

    return withSubmitLock('zbirna:confirm', confirmZbirnaUnlocked, {
        action: 'confirm-zbirna',
        reason: 'confirmZbirna',
        alreadyMessage: 'Kreiranje zbirne je već u toku'
    });
}

async function confirmZbirnaUnlocked() {
    const kupacSel = document.getElementById('fldZbirnaKupac');
    if (!kupacSel || !kupacSel.value) {
        showToast('Izaberite kupca', 'error');
        return;
    }

    const kupacID = kupacSel.value;
    const today = getTodayIsoDate();

    // Koristi cached zbirne umesto novog API call-a
    const zbirne = _lastMergedZbirne || await getMergedZbirneForVozac();
    const consumedIds = getConsumedOtkupIdsFromZbirne(zbirne);

    const todayOtkupi = (vozacOtkupi || []).filter(r =>
        r.datum === today && !consumedIds.has(r.clientRecordID)
    );

    if (todayOtkupi.length === 0) {
        showToast('Nema otkupa za danas', 'error');
        return;
    }

    let totalKgI = 0;
    let totalKgII = 0;
    let totalAmb = 0;
    const vrste = new Set();
    const sorte = new Set();

    todayOtkupi.forEach(r => {
        if (r.klasa === 'II') totalKgII += r.kolicina || 0;
        else totalKgI += r.kolicina || 0;

        totalAmb += r.kolAmbalaze || 0;
        if (r.vrstaVoca) vrste.add(r.vrstaVoca);
        if (r.sortaVoca) sorte.add(r.sortaVoca);
    });

    const kupacName = kupacSel.selectedOptions[0]
        ? kupacSel.selectedOptions[0].textContent
        : kupacID;

    const nowIso = new Date().toISOString();

    // === BrojZbirne PWA-side generation ===
    const vozacBrojX = parseInt(String(CONFIG.ENTITY_ID || '').replace(/\D/g, ''), 10);
    if (!vozacBrojX || isNaN(vozacBrojX)) {
        showToast('Greška: VozacID nije validan za generaciju broja zbirne', 'error');
        return;
    }

    const ddmmyy = formatDdmmyy(today);

    // Sequence: count današnjih ne-deleted zbirni iz cached merged set
    const todayZbirneCount = (zbirne || [])
        .filter(z => z.datum === today)
        .length;

    const seq = todayZbirneCount + 1;
    const brojZbirne = (seq === 1)
        ? `${vozacBrojX}/${ddmmyy}`
        : `${vozacBrojX}/${ddmmyy}-${seq}`;
    // === end BrojZbirne ===

    const record = {
        clientRecordID: (window.crypto && typeof window.crypto.randomUUID === 'function')
            ? window.crypto.randomUUID()
            : ('zbr-' + Date.now() + '-' + Math.floor(Math.random() * 1000000)),
        serverRecordID: '',
        brojZbirne: brojZbirne,
        createdAtClient: nowIso,
        updatedAtClient: nowIso,
        updatedAtServer: '',
        syncedAt: '',

        vozacID: CONFIG.ENTITY_ID,
        datum: today,
        kupacID: kupacID,
        kupacName: kupacName,
        vrstaVoca: Array.from(vrste).join(', '),
        sortaVoca: Array.from(sorte).join(', '),
        kolicinaKlI: totalKgI,
        kolicinaKlII: totalKgII,
        tipAmbalaze: todayOtkupi[0].tipAmbalaze || '',
        kolAmbalaze: totalAmb,
        klasa: totalKgII > 0 ? 'I+II' : 'I',
        otkupRecordIDs: todayOtkupi.map(r => r.clientRecordID).join(','),

        syncStatus: 'pending',
        syncAttempts: 0,
        syncAttemptAt: '',
        lastSyncError: '',
        lastServerStatus: '',
        deleted: false,
        entityType: 'zbirna',
        schemaVersion: 1
    };

    try {
        await dbPut(db, 'zbirne', record);
    } catch (err) {
        console.error('confirmZbirna dbPut failed:', err);
        showToast('Greška pri čuvanju zbirne', 'error');
        return;
    }

    showToast('Zbirna kreirana!', 'success');
    cancelZbirna();

    if (navigator.onLine && typeof syncQueueSafe === 'function') {
        syncQueueSafe('post-save');
    }

    try {
        await loadVozacData();
    } catch (err) {
        console.error('confirmZbirna loadVozacData failed:', err);
    }
}

async function syncZbirne() {
    const result = await syncStore({
        storeName: 'zbirne',
        action: 'syncZbirna',
        inFlightKey: 'zbirnaInFlight',
        entityIdField: 'vozacID',
        successLabel: 'Zbirna sinhronizovana',
        onResultRecord: (record, serverResult) => {
            // Zbirna-specific: backend returns brojZbirne on success
            if (serverResult.brojZbirne) {
                record.brojZbirne = serverResult.brojZbirne;
            }
        }
    });

    try { await loadVozacZbirne(); } catch (_) {}

    return result;
}

function getConsumedOtkupIdsFromZbirne(zbirne) {
    const used = new Set();

    (zbirne || []).forEach(z => {
        const raw = String(z.otkupRecordIDs || '').trim();
        if (!raw) return;

        raw.split(',')
            .map(x => x.trim())
            .filter(Boolean)
            .forEach(id => used.add(id));
    });

    return used;
}

async function getMergedZbirneForVozac() {
    let local = [];
    let server = [];

    try {
        local = await dbGetAll(db, 'zbirne');
    } catch (err) {
        console.error('getMergedZbirneForVozac local failed:', err);
    }

    const json = await safeAsync(async () => {
        return await apiFetch('action=getVozacZbirne');
    }, 'Greška pri učitavanju zbirnih');

    if (json && json.success && Array.isArray(json.records)) {
        server = json.records.map(mapServerZbirnaRecord);
    }

    return dedupeRecordsForRender(mergeZbirneRecords(local, server))
    .filter(r => !r.deleted);
}

function formatDdmmyy(isoDate) {
    // '2026-05-06' → '060526'
    if (!isoDate || typeof isoDate !== 'string') return '';
    const parts = isoDate.split('-');
    if (parts.length !== 3) return '';
    const [y, m, d] = parts;
    return d + m + y.slice(2);
}
