// ============================================================
// VOZAC: ZBIRNA
// ============================================================
let vozacOtkupi = [];

async function loadVozacData() {
    vozacOtkupi = [];

    const json = await safeAsync(async () => {
        return await apiFetch('action=getVozacOtkupi');
    }, 'Greška pri učitavanju podataka vozača');

    if (json && json.success && json.records) {
        vozacOtkupi = json.records.map(r => ({
            clientRecordID: r.ClientRecordID || '',
            datum: fmtDate(r.Datum),
            kooperantName: r.KooperantName || r.KooperantID || '',
            kooperantID: r.KooperantID || '',
            vrstaVoca: r.VrstaVoca || '',
            sortaVoca: r.SortaVoca || '',
            klasa: r.Klasa || 'I',
            kolicina: parseFloat(r.Kolicina) || 0,
            cena: parseFloat(r.Cena) || 0,
            tipAmbalaze: r.TipAmbalaze || '',
            kolAmbalaze: parseInt(r.KolAmbalaze) || 0,
            stanicaID: r.OtkupacID || r._source || '',
            zbirnaID: r._zbirnaID || ''
        }));
    }

    renderVozacOtpremnice();
    loadVozacZbirne();
}

function renderVozacOtpremnice() {
    const today = new Date().toISOString().split('T')[0];
    const todayOtkupi = vozacOtkupi.filter(r => r.datum === today);
    
    const list = document.getElementById('vozacOtpremniceList');
    if (todayOtkupi.length === 0) {
        list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema otpremnica za danas</p>';
        document.getElementById('btnNovaZbirna').style.display = 'none';
        return;
    }
    document.getElementById('btnNovaZbirna').style.display = '';
    
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

async function confirmZbirna() {
    const kupacID = document.getElementById('fldZbirnaKupac').value;
    if (!kupacID) { showToast('Izaberite kupca!', 'error'); return; }
    
    const today = new Date().toISOString().split('T')[0];
    const todayOtkupi = vozacOtkupi.filter(r => r.datum === today);
    
    if (todayOtkupi.length === 0) { showToast('Nema otkupa za danas', 'error'); return; }
    
    let totalKgI = 0, totalKgII = 0, totalAmb = 0;
    const vrste = new Set(), sorte = new Set();
    todayOtkupi.forEach(r => {
        if (r.klasa === 'II') totalKgII += r.kolicina || 0;
        else totalKgI += r.kolicina || 0;
        totalAmb += r.kolAmbalaze || 0;
        if (r.vrstaVoca) vrste.add(r.vrstaVoca);
        if (r.sortaVoca) sorte.add(r.sortaVoca);
    });
    
    const kupacName = document.getElementById('fldZbirnaKupac').selectedOptions[0].textContent;
    
    const record = {
        clientRecordID: crypto.randomUUID(),
        createdAtClient: new Date().toISOString(),
        vozacID: CONFIG.ENTITY_ID,
        datum: today,
        kupacID: kupacID,
        kupacName: kupacName,
        vrstaVoca: [...vrste].join(', '),
        sortaVoca: [...sorte].join(', '),
        kolicinaKlI: totalKgI,
        kolicinaKlII: totalKgII,
        tipAmbalaze: todayOtkupi[0].tipAmbalaze || '',
        kolAmbalaze: totalAmb,
        klasa: totalKgII > 0 ? 'I+II' : 'I',
        otkupRecordIDs: todayOtkupi.map(r => r.clientRecordID).join(','),
        syncStatus: 'pending'
    };
    
    await dbPut(db, 'zbirne', record);
    showToast('Zbirna kreirana!', 'success');
    cancelZbirna();
    
    // Sync immediately
    if (navigator.onLine) syncZbirne();
}

function cancelZbirna() {
    document.getElementById('zbirnaCreateView').style.display = 'none';
    document.getElementById('zbirnaMainView').style.display = 'block';
    loadVozacZbirne();
}

async function loadVozacZbirne() {
    const local = await dbGetAll(db, 'zbirne');

    let server = [];
    try {
        const json = await apiFetch('action=getVozacZbirne');
        if (json && json.success && json.records) server = json.records.map(r => ({
            clientRecordID: r.ClientRecordID || '',
            datum: fmtDate(r.Datum),
            kupacName: r.KupacName || r.KupacID || '',
            kolicinaKlI: parseFloat(r.KolicinaKlI) || 0,
            kolicinaKlII: parseFloat(r.KolicinaKlII) || 0,
            kolAmbalaze: parseInt(r.KolAmbalaze) || 0,
            vrstaVoca: r.VrstaVoca || '',
            syncStatus: 'synced'
        }));
    } catch (e) {}

    const serverIDs = new Set(server.map(r => r.clientRecordID));
    const all = [...server, ...local.filter(r => r.syncStatus === 'pending' && !serverIDs.has(r.clientRecordID))];

    const list = document.getElementById('vozacZbirneList');
    if (all.length === 0) {
        list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:12px;">Nema kreiranih zbirnih</p>';
        return;
    }

    list.innerHTML = all.map(r => {
        const totalKg = (r.kolicinaKlI || 0) + (r.kolicinaKlII || 0);
        const bc = r.syncStatus === 'pending' ? 'var(--warning)' : 'var(--success)';
        return `<div class="queue-item" style="border-left-color:${bc};">
            <div class="qi-header"><span class="qi-koop">🏭 ${escapeHtml(r.kupacName)}</span><span class="qi-time">${escapeHtml(r.datum)}</span></div>
            <div class="qi-detail">${escapeHtml(r.vrstaVoca)} | ${totalKg.toLocaleString('sr')} kg | Amb: ${r.kolAmbalaze || 0}</div>
            ${r.kolicinaKlII > 0 ? '<div class="qi-detail" style="font-size:11px;">Kl.I: ' + (r.kolicinaKlI||0).toLocaleString('sr') + ' kg | Kl.II: ' + (r.kolicinaKlII||0).toLocaleString('sr') + ' kg</div>' : ''}
        </div>`;
    }).join('');
}

async function syncZbirne() {
    const pending = await dbGetByIndex(db, 'zbirne', 'syncStatus', 'pending');
    if (pending.length === 0) return;

    const json = await apiPost('syncZbirna', {
        vozacID: CONFIG.ENTITY_ID,
        records: pending
    });

    if (json && json.success) {
        for (const r of pending) {
            r.syncStatus = 'synced';
            await dbPut(db, 'zbirne', r);
        }
        showToast('Zbirna sinhronizovana', 'success');
    }
}
