// ============================================================
// MANAGEMENT: KUPCI
// ============================================================
async function loadMgmtFakture() {
    const kupacID = document.getElementById('mgmtFaktureKupac').value;
    if (!kupacID) { document.getElementById('mgmtFaktureList').innerHTML = ''; return; }
    
    let records = [];
    if (mgmtData && mgmtData.fakture) {
        records = mgmtData.fakture.filter(r => 
            String(r.KupacID) === kupacID || String(r.Kupac) === kupacID
        );
    } else {
        document.getElementById('mgmtFaktureList').innerHTML = '<p style="text-align:center;padding:20px;color:var(--text-muted);">Učitavanje...</p>';
        try {
            const json = await apiFetch('action=getMgmtFakture&kupacID=' + encodeURIComponent(kupacID));
            if (json && json.success && json.records) records = json.records;
        } catch (e) {}
    }
    
    const list = document.getElementById('mgmtFaktureList');
    if (records.length === 0) { list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema faktura</p>'; return; }
    list.innerHTML = records.map(r => {
        const iznos = parseFloat(r.Iznos) || 0;
        const placeno = parseFloat(r.Placeno) || 0;
        const saldo = parseFloat(r.Saldo) || 0;
        const bc = saldo <= 0 ? 'var(--success)' : 'var(--danger)';
        return `<div class="queue-item" style="border-left-color:${bc};cursor:pointer;" onclick="toggleFakturaStavke('${r.FakturaID}', this)">
            <div class="qi-header"><span class="qi-koop">${r.BrojFakture || r.FakturaID}</span><span class="qi-time">${fmtDate(r.Datum)}</span></div>
            <div class="qi-detail">Iznos: <strong>${iznos.toLocaleString('sr')}</strong> | Plaćeno: ${placeno.toLocaleString('sr')} | Saldo: <strong>${saldo.toLocaleString('sr')}</strong></div>
            <div class="qi-detail" style="font-size:11px;margin-top:2px;">${r.Status || ''}${r.SEFStatus ? ' | SEF: ' + r.SEFStatus : ''}</div>
            <div class="faktura-stavke" id="stavke-${r.FakturaID}" style="display:none;margin-top:8px;padding-top:8px;border-top:1px solid #eee;"></div>
        </div>`;
    }).join('');
}

async function toggleFakturaStavke(fakturaID, parentEl) {
    const div = document.getElementById('stavke-' + fakturaID);
    if (!div) return;
    if (div.style.display === 'block') { div.style.display = 'none'; return; }
    div.style.display = 'block';
    
    let stavke = [];
    if (mgmtData && mgmtData.fakturaStavke) {
        stavke = mgmtData.fakturaStavke.filter(r => String(r.FakturaID) === fakturaID);
    } else {
        div.innerHTML = '<span style="font-size:12px;color:var(--text-muted);">Učitavanje...</span>';
        try {
            const json = await apiFetch('action=getMgmtFakturaStavke&fakturaID=' + encodeURIComponent(fakturaID));
            if (json && json.success && json.records) stavke = json.records;
        } catch (e) {}
    }
    
    if (stavke.length === 0) { div.innerHTML = '<span style="font-size:12px;color:var(--text-muted);">Nema stavki</span>'; return; }
    div.innerHTML = `<table style="width:100%;font-size:11px;border-collapse:collapse;">
        <tr style="color:var(--text-muted);"><td>Prijemnica</td><td>Zbirna</td><td>Klasa</td><td style="text-align:right;">Kg</td><td style="text-align:right;">Cena</td><td style="text-align:right;">Iznos</td></tr>
        ${stavke.map(s => `<tr style="border-top:1px solid #f0f0f0;">
            <td>${s.BrojPrijemnice || s.PrijemnicaID || ''}</td>
            <td>${s.BrojZbirne || ''}</td>
            <td>${s.Klasa || ''}</td>
            <td style="text-align:right;">${(parseFloat(s.Kolicina) || 0).toLocaleString('sr')}</td>
            <td style="text-align:right;">${(parseFloat(s.Cena) || 0).toLocaleString('sr')}</td>
            <td style="text-align:right;font-weight:600;">${(parseFloat(s.Iznos) || 0).toLocaleString('sr')}</td>
        </tr>`).join('')}
    </table>`;
}

function loadMgmtKupci() {
    const records = (mgmtData && mgmtData.saldoKupci) ? mgmtData.saldoKupci : [];
    const list = document.getElementById('mgmtKupciList');
    if (records.length === 0) { list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema podataka</p>'; return; }
    list.innerHTML = records.map(r => {
        const saldo = parseFloat(r.Saldo)||0;
        const bc = saldo > 0 ? 'var(--danger)' : 'var(--success)';
        return `<div class="queue-item" style="border-left-color:${bc};">
            <div class="qi-header"><span class="qi-koop">${r.Kupac||r.KupacID||''}</span><span class="qi-time">Saldo: ${saldo.toLocaleString('sr')} RSD</span></div>
            <div class="qi-detail">Fakturisano: ${(parseFloat(r.Fakturisano)||0).toLocaleString('sr')} | Placeno: ${(parseFloat(r.Placeno)||0).toLocaleString('sr')}</div></div>`;
    }).join('');
}

function loadMgmtPredato() {
    const records = (mgmtData && mgmtData.predatoPoKupcu) ? mgmtData.predatoPoKupcu : [];
    const list = document.getElementById('mgmtPredatoList');
    if (records.length === 0) { list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema podataka</p>'; return; }
    const grouped = {};
    records.forEach(r => {
        const k = r.KupacID||'?';
        if (!grouped[k]) grouped[k] = { items: [], totalKg: 0, totalAmb: 0, totalVr: 0 };
        const kg = parseFloat(r.Kolicina)||0, amb = parseFloat(r.Ambalaza)||0, vr = parseFloat(r.Vrednost)||0;
        grouped[k].items.push(r); grouped[k].totalKg += kg; grouped[k].totalAmb += amb; grouped[k].totalVr += vr;
    });
    list.innerHTML = Object.entries(grouped).map(([kupac, g]) =>
        `<div style="background:white;border-radius:var(--radius);padding:14px;margin-bottom:12px;box-shadow:0 1px 3px rgba(0,0,0,0.08);border-left:4px solid var(--accent);">
            <div style="display:flex;justify-content:space-between;margin-bottom:8px;"><strong style="color:var(--primary);font-size:16px;">${kupac}</strong><span style="font-size:13px;font-weight:600;">${g.totalKg.toLocaleString('sr')} kg</span></div>
            <div style="font-size:12px;color:var(--text-muted);margin-bottom:8px;">Amb: ${g.totalAmb.toLocaleString('sr')} | Vrednost: ${g.totalVr.toLocaleString('sr')} RSD</div>
            ${g.items.map(r => { const kg=parseFloat(r.Kolicina)||0,amb=parseFloat(r.Ambalaza)||0,vr=parseFloat(r.Vrednost)||0; return `<div style="padding:4px 0;font-size:12px;border-top:1px solid #eee;display:flex;justify-content:space-between;"><span>${r.VrstaVoca} ${r.Klasa}</span><span>${kg.toLocaleString('sr')} kg | ${amb.toLocaleString('sr')} amb | ${vr.toLocaleString('sr')} RSD | ${r.BrojPrijemnica||0} prij.</span></div>`; }).join('')}
        </div>`).join('');
}

