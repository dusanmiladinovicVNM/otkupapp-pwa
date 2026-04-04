// ============================================================
// MANAGEMENT: STANICE
// ============================================================
async function loadMgmtOtkupi() {
    const stanicaID = document.getElementById('mgmtOtkupiStanica').value;
    if (!stanicaID) { document.getElementById('mgmtOtkupiList').innerHTML = ''; return; }
    const od = document.getElementById('mgmtOtkupiOd').value, doo = document.getElementById('mgmtOtkupiDo').value;
    
    let records = [];
    if (mgmtData && mgmtData.otkupiAll) {
        records = mgmtData.otkupiAll
            .filter(r => (r._sheetName === 'OTK-' + stanicaID) || (r.OtkupacID === stanicaID))
            .map(r => ({
                datum: fmtDate(r.Datum), kooperantName: r.KooperantName || r.KooperantID || '',
                vrstaVoca: r.VrstaVoca || '', klasa: r.Klasa || 'I',
                kolicina: parseFloat(r.Kolicina) || 0, cena: parseFloat(r.Cena) || 0
            }));
    } else {
        try {
            const json = await apiFetch('action=getMgmtOtkupiByStanica&stanicaID=' + encodeURIComponent(stanicaID));
            if (json && json.success && json.records) records = json.records.map(r => ({
                datum: fmtDate(r.Datum), kooperantName: r.KooperantName || r.KooperantID || '',
                vrstaVoca: r.VrstaVoca || '', klasa: r.Klasa || 'I',
                kolicina: parseFloat(r.Kolicina) || 0, cena: parseFloat(r.Cena) || 0
            }));
        } catch (e) {}
    }
    
    if (od) records = records.filter(r => r.datum >= od);
    if (doo) records = records.filter(r => r.datum <= doo);
    records.sort((a, b) => b.datum.localeCompare(a.datum));
    const kg = records.reduce((s, r) => s + (r.kolicina || 0), 0);
    const vr = records.reduce((s, r) => s + (r.kolicina || 0) * (r.cena || 0), 0);
    document.getElementById('mgmtOtkupiCount').textContent = records.length;
    document.getElementById('mgmtOtkupiKg').textContent = kg.toLocaleString('sr');
    document.getElementById('mgmtOtkupiVrednost').textContent = vr.toLocaleString('sr');
    document.getElementById('mgmtOtkupiKoop').textContent = new Set(records.map(r => r.kooperantName)).size;
    const list = document.getElementById('mgmtOtkupiList');
    if (records.length === 0) { list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema otkupa</p>'; return; }
    list.innerHTML = records.map(r => {
        const v = ((r.kolicina || 0) * (r.cena || 0)).toLocaleString('sr');
        return `<div class="queue-item"><div class="qi-header"><span class="qi-koop">${r.kooperantName}</span><span class="qi-time">${r.datum}</span></div>
            <div class="qi-detail">${r.vrstaVoca} ${r.klasa} | ${r.kolicina} kg × ${r.cena} = <strong>${v} RSD</strong></div></div>`;
    }).join('');
}

function loadMgmtSaldoOM() {
    const records = (mgmtData && mgmtData.saldoOM) ? mgmtData.saldoOM : [];
    const list = document.getElementById('mgmtSaldoOMList');
    if (records.length === 0) { list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema podataka</p>'; return; }
    list.innerHTML = records.map(r => {
        const saldo = parseFloat(r.Saldo)||0;
        const bc = saldo > 0 ? 'var(--warning)' : 'var(--success)';
        return `<div class="queue-item" style="border-left-color:${bc};">
            <div class="qi-header"><span class="qi-koop">${fmtStanica(r.StanicaID || r.Stanica)}</span><span class="qi-time">Saldo: ${saldo.toLocaleString('sr')} RSD</span></div>
            <div class="qi-detail">Avans: ${(parseFloat(r.Avans)||0).toLocaleString('sr')} | Isplaceno: ${(parseFloat(r.Isplaceno)||0).toLocaleString('sr')}</div></div>`;
    }).join('');
}

function loadMgmtOtkupPoOM() {
    const records = (mgmtData && mgmtData.otkupPoOM) ? mgmtData.otkupPoOM : [];
    const list = document.getElementById('mgmtOtkupPoOMList');
    if (records.length === 0) { list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema podataka</p>'; return; }
    const grouped = {};
    records.forEach(r => {
        const s = r.StanicaID||'?';
        if (!grouped[s]) grouped[s] = { items: [], totalKg: 0, totalAmb: 0, totalVr: 0 };
        const kg = parseFloat(r.Kolicina)||0, amb = parseFloat(r.Ambalaza)||0, vr = parseFloat(r.Vrednost)||0;
        grouped[s].items.push(r); grouped[s].totalKg += kg; grouped[s].totalAmb += amb; grouped[s].totalVr += vr;
    });
    list.innerHTML = Object.entries(grouped).map(([stanica, g]) =>
        `<div style="background:white;border-radius:var(--radius);padding:14px;margin-bottom:12px;box-shadow:0 1px 3px rgba(0,0,0,0.08);border-left:4px solid var(--primary);">
            <div style="display:flex;justify-content:space-between;margin-bottom:8px;"><strong style="color:var(--primary);font-size:16px;">${fmtStanica(stanica)}</strong><span style="font-size:13px;font-weight:600;">${g.totalKg.toLocaleString('sr')} kg</span></div>
            <div style="font-size:12px;color:var(--text-muted);margin-bottom:8px;">Amb: ${g.totalAmb.toLocaleString('sr')} | Vrednost: ${g.totalVr.toLocaleString('sr')} RSD</div>
            ${g.items.map(r => { const kg=parseFloat(r.Kolicina)||0,amb=parseFloat(r.Ambalaza)||0,vr=parseFloat(r.Vrednost)||0; return `<div style="padding:4px 0;font-size:12px;border-top:1px solid #eee;display:flex;justify-content:space-between;"><span>${r.VrstaVoca} ${r.Klasa}</span><span>${kg.toLocaleString('sr')} kg | ${amb.toLocaleString('sr')} amb | ${vr.toLocaleString('sr')} RSD | ${r.BrojOtkupa||0} otk.</span></div>`; }).join('')}
        </div>`).join('');
}
