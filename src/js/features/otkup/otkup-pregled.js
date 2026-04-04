// ============================================================
// OTKUP PREGLED
// ============================================================
async function loadOtkupPregled() {
    document.getElementById('pregledList').innerHTML = '<p style="text-align:center;padding:20px;color:var(--text-muted);">Učitavanje...</p>';
    const od = document.getElementById('fldPregledOd').value, doo = document.getElementById('fldPregledDo').value;
    const local = await dbGetAll(db, CONFIG.STORE_NAME);
    let server = [];
    try {
        const json = await apiFetch('action=getOtkupi&otkupacID=' + encodeURIComponent(CONFIG.OTKUPAC_ID));
        if (json && json.success && json.records) server = json.records.map(r => ({
            clientRecordID: r.ClientRecordID||'', datum: fmtDate(r.Datum), kooperantID: r.KooperantID||'',
            kooperantName: r.KooperantName||r.KooperantID||'', vrstaVoca: r.VrstaVoca||'', sortaVoca: r.SortaVoca||'',
            klasa: r.Klasa||'I', kolicina: parseFloat(r.Kolicina)||0, cena: parseFloat(r.Cena)||0,
            kolAmbalaze: parseInt(r.KolAmbalaze)||0, parcelaID: r.ParcelaID||'', syncStatus: r.SyncStatus||'synced'
        }));
    } catch (e) {}
    const serverIDs = new Set(server.map(r => r.clientRecordID));
    let all = [...server, ...local.filter(r => r.syncStatus === 'pending' && !serverIDs.has(r.clientRecordID))];
    if (od) all = all.filter(r => r.datum >= od);
    if (doo) all = all.filter(r => r.datum <= doo);
    all.sort((a, b) => b.datum.localeCompare(a.datum));
    const kg = all.reduce((s, r) => s + (r.kolicina||0), 0);
    const vr = all.reduce((s, r) => s + (r.kolicina||0)*(r.cena||0), 0);
    document.getElementById('statPregledCount').textContent = all.length;
    document.getElementById('statPregledKg').textContent = kg.toLocaleString('sr');
    document.getElementById('statPregledVrednost').textContent = vr.toLocaleString('sr');
    document.getElementById('statPregledKoop').textContent = new Set(all.map(r => r.kooperantID)).size;
    const list = document.getElementById('pregledList');
    if (all.length === 0) { list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:40px;">Nema otkupa</p>'; return; }
    list.innerHTML = all.map(r => {
        const v = ((r.kolicina||0)*(r.cena||0)).toLocaleString('sr');
        const bc = r.syncStatus==='pending' ? 'var(--warning)' : 'var(--success)';
        return `<div class="queue-item" style="border-left-color:${bc};"><div class="qi-header"><span class="qi-koop">${escapeHtml(r.kooperantName)}</span><span class="qi-time">${escapeHtml(r.datum)}</span></div><div class="qi-detail">${escapeHtml(r.vrstaVoca)} ${escapeHtml(r.sortaVoca||'')} ${escapeHtml(r.klasa)} | ${r.kolicina} kg × ${r.cena} = <strong>${v} RSD</strong>${r.parcelaID ? ' | ' + escapeHtml(r.parcelaID) : ''}</div></div>`;
    }).join('');
}
