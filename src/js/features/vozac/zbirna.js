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
    const today = new Date().toISOString().split('T')[0];
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
        `<div style="background:white;border-radius:var(--radius);padding:14px;margin-bottom:12px;box-shadow:0 1px 3px rgba(0,0,0,0.08);border-left:4px solid var(--primary);">
            <div style="display:flex;justify-content:space-between;margin-bottom:6px;">
                <strong style="color:var(--primary);">${escapeHtml(fmtStanica(sta))}</strong>
                <span style="font-weight:600;">${g.kg.toLocaleString('sr')} kg</span>
            </div>
            <div style="font-size:12px;color:var(--text-muted);margin-bottom:6px;">${g.items.length} otkupa | Amb: ${g.amb}</div>
            ${g.items.map(r => `<div style="padding:3px 0;font-size:12px;border-top:1px solid #eee;">
               ${escapeHtml(r.kooperantName)} | ${escapeHtml(r.vrstaVoca)} ${escapeHtml(r.klasa)} | ${r.kolicina} kg | ${r.kolAmbalaze} amb
            </div>`).join('')}
        </div>`).join('');
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
    const today = new Date().toISOString().split('T')[0];
    const todayOtkupi = vozacOtkupi.filter(r => r.datum === today);
    
    let totalKgI = 0, totalKgII = 0, totalAmb = 0;
    todayOtkupi.forEach(r => {
        if (r.klasa === 'II') totalKgII += r.kolicina || 0;
        else totalKgI += r.kolicina || 0;
        totalAmb += r.kolAmbalaze || 0;
    });
    
    document.getElementById('zbirnaOtkupiSummary').innerHTML = 
        `<div style="font-size:16px;font-weight:700;">Ukupno: ${(totalKgI + totalKgII).toLocaleString('sr')} kg</div>
         <div style="font-size:13px;opacity:0.9;">Kl. I: ${totalKgI.toLocaleString('sr')} kg | Kl. II: ${totalKgII.toLocaleString('sr')} kg | Amb: ${totalAmb}</div>
         <div style="font-size:12px;opacity:0.7;">${todayOtkupi.length} otkupa sa ${new Set(todayOtkupi.map(r => r.stanicaID)).size} stanica</div>`;
    
    // List individual otkupi
    document.getElementById('zbirnaOtkupiList').innerHTML = todayOtkupi.map(r => {
        const vr = ((r.kolicina || 0) * (r.cena || 0)).toLocaleString('sr');
        return `<div class="queue-item">
             <div class="qi-header"><span class="qi-koop">${escapeHtml(r.kooperantName)}</span><span class="qi-time">${escapeHtml(fmtStanica(r.stanicaID))}</span></div>
             <div class="qi-detail">${escapeHtml(r.vrstaVoca)} ${escapeHtml(r.klasa)} | ${r.kolicina} kg × ${r.cena} = ${vr} RSD | Amb: ${r.kolAmbalaze}</div>
        </div>`;
    }).join('');
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
        const bc = (r.syncStatus === 'pending' || r.syncStatus === 'syncing')
            ? 'var(--warning)'
            : 'var(--success)';

        const syncText =
            r.syncStatus === 'syncing' ? ' | sync...' :
            r.syncStatus === 'pending' ? ' | pending' :
            (r.brojZbirne ? ' | ' + r.brojZbirne :
             r.serverRecordID ? ' | ' + r.serverRecordID : '');

        return `<div class="queue-item" style="border-left-color:${bc};">
            <div class="qi-header">
                <span class="qi-koop">🏭 ${escapeHtml(r.kupacName)}</span>
                <span class="qi-time">${escapeHtml(r.datum)}</span>
            </div>
            <div class="qi-detail">
                ${escapeHtml(r.vrstaVoca)} | ${totalKg.toLocaleString('sr')} kg | Amb: ${r.kolAmbalaze || 0}${escapeHtml(syncText)}
            </div>
            ${r.kolicinaKlII > 0
                ? '<div class="qi-detail" style="font-size:11px;">Kl.I: ' + (r.kolicinaKlI || 0).toLocaleString('sr') + ' kg | Kl.II: ' + (r.kolicinaKlII || 0).toLocaleString('sr') + ' kg</div>'
                : ''}
            ${r.lastSyncError
                ? '<div class="qi-detail" style="font-size:11px;color:#b42318;">' + escapeHtml(r.lastSyncError) + '</div>'
                : ''}
        </div>`;
    }).join('');
}

function cancelZbirna() {
    document.getElementById('zbirnaCreateView').style.display = 'none';
    document.getElementById('zbirnaMainView').style.display = 'block';
    loadVozacZbirne();
}

async function confirmZbirna() {
    const kupacSel = document.getElementById('fldZbirnaKupac');
    if (!kupacSel || !kupacSel.value) {
        showToast('Izaberite kupca', 'error');
        return;
    }

    const kupacID = kupacSel.value;
    const today = new Date().toISOString().split('T')[0];

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

    const record = {
        clientRecordID: (window.crypto && typeof window.crypto.randomUUID === 'function')
            ? window.crypto.randomUUID()
            : ('zbr-' + Date.now() + '-' + Math.floor(Math.random() * 1000000)),
        serverRecordID: '',
        brojZbirne: '',
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

    if (navigator.onLine) {
        if (typeof syncZbirne === 'function') {
            try { await syncZbirne(); } catch (_) {}
        }
    }
    
    try {
        await loadVozacData();
    } catch (err) {
        console.error('confirmZbirna loadVozacData failed:', err);
    }
}

function getVozacRuntime() {
    return window.appRuntime || {};
}

async function syncZbirne() {
    const runtime = getVozacRuntime();
    
    if (!db) return { ok: false, reason: 'db-not-ready' };
    if (!navigator.onLine) return { ok: false, reason: 'offline' };

    if (runtime.zbirnaSyncInFlight) {
        return { ok: false, reason: 'already-running' };
    }
    runtime.zbirnaSyncInFlight = true;

    let pending = [];

    try {
        pending = await dbGetByIndex(db, 'zbirne', 'syncStatus', 'pending');

        if (!Array.isArray(pending) || pending.length === 0) {
            return { ok: true, synced: 0, failed: 0 };
        }

        for (const record of pending) {
            record.syncStatus = 'syncing';
            record.syncAttemptAt = new Date().toISOString();
            await dbPut(db, 'zbirne', record);
        }

        const json = await apiPost('syncZbirna', {
            vozacID: CONFIG.ENTITY_ID,
            records: pending
        });

        if (!json || json.success === false) {
            for (const record of pending) {
                record.syncStatus = 'pending';
                record.lastSyncError = json && json.error ? json.error : 'Sync neuspešan';
                record.lastServerStatus = 'request-failed';
                record.syncAttempts = (record.syncAttempts || 0) + 1;
                await dbPut(db, 'zbirne', record);
            }

            showToast('Zbirna sync nije uspeo', 'error');
            return { ok: false, synced: 0, failed: pending.length };
        }

        if (Array.isArray(json.results)) {
            const byClientId = new Map(
                pending.map(r => [r.clientRecordID, r])
            );

            const mentionedIds = new Set(
                json.results.map(x => x.clientRecordID).filter(Boolean)
            );

            let syncedCount = 0;
            let failedCount = 0;

            for (const result of json.results) {
                const record = byClientId.get(result.clientRecordID);
                if (!record) continue;

                record.syncAttempts = (record.syncAttempts || 0) + 1;

                const isSuccess =
                    !!result.success ||
                    result.status === 'synced' ||
                    result.status === 'duplicate' ||
                    result.status === 'existing' ||
                    result.status === 'inserted' ||
                    result.status === 'updated';

                if (isSuccess) {
                    record.syncStatus = 'synced';
                    record.lastSyncError = '';
                    record.syncedAt = new Date().toISOString();
                    record.serverRecordID = result.serverRecordID || record.serverRecordID || '';
                    record.updatedAtServer = result.updatedAtServer || record.updatedAtServer || '';
                    record.brojZbirne = result.brojZbirne || record.brojZbirne || '';
                    record.lastServerStatus = result.status || 'synced';
                    syncedCount++;
                } else {
                    record.syncStatus = 'pending';
                    record.lastSyncError = result.error || 'Sync stavke neuspešan';
                    record.lastServerStatus = result.status || 'failed';
                    failedCount++;
                }

                await dbPut(db, 'zbirne', record);
            }

            for (const record of pending) {
                if (!mentionedIds.has(record.clientRecordID)) {
                    record.syncStatus = 'pending';
                    record.lastSyncError = 'Nema potvrde sa servera';
                    record.lastServerStatus = 'missing-result';
                    record.syncAttempts = (record.syncAttempts || 0) + 1;
                    failedCount++;
                    await dbPut(db, 'zbirne', record);
                }
            }

            if (syncedCount > 0 && failedCount === 0) {
                showToast('Zbirna sinhronizovana: ' + syncedCount, 'success');
            } else if (syncedCount > 0) {
                showToast('Zbirna sync: ' + syncedCount + ' uspešno, ' + failedCount + ' neuspešno', 'info');
            } else {
                showToast('Zbirne nisu sinhronizovane', 'error');
            }

            return { ok: failedCount === 0, synced: syncedCount, failed: failedCount };
        }

        // legacy fallback
        for (const record of pending) {
            record.syncStatus = 'synced';
            record.lastSyncError = '';
            record.syncedAt = new Date().toISOString();
            record.lastServerStatus = 'legacy-success';
            record.syncAttempts = (record.syncAttempts || 0) + 1;
            await dbPut(db, 'zbirne', record);
        }

        showToast('Zbirna sinhronizovana', 'success');
        return { ok: true, synced: pending.length, failed: 0 };
    } catch (err) {
        console.error('syncZbirne failed:', err);

        for (const record of pending) {
            try {
                if (record.syncStatus === 'syncing') {
                    record.syncStatus = 'pending';
                    record.lastSyncError = err.message || 'Greška pri sync-u';
                    record.lastServerStatus = 'exception';
                    record.syncAttempts = (record.syncAttempts || 0) + 1;
                    await dbPut(db, 'zbirne', record);
                }
            } catch (_) {}
        }

        showToast('Greška pri sinhronizaciji zbirnih', 'error');
        return { ok: false, synced: 0, failed: pending.length || 0 };
    } finally {
       runtime.zbirnaSyncInFlight = false;
        try { await loadVozacZbirne(); } catch (_) {}
    }
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

    return mergeZbirneRecords(local, server).filter(r => !r.deleted);
}
