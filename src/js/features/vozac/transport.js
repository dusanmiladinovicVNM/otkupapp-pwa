    async function loadVozacTransport() {
    const list = document.getElementById('transportList');
    const local = await dbGetAll(db, 'zbirne');
    if (local.length === 0) {
        list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema transporta</p>';
        return;
    }
    list.innerHTML = local.map(r => {
        const totalKg = (r.kolicinaKlI || 0) + (r.kolicinaKlII || 0);
        return `<div class="queue-item">
            <div class="qi-header"><span class="qi-koop">🏭 ${r.kupacName || r.kupacID}</span><span class="qi-time">${r.datum}</span></div>
            <div class="qi-detail">${totalKg.toLocaleString('sr')} kg | Amb: ${r.kolAmbalaze || 0} | ${r.syncStatus === 'synced' ? '✅' : '⏳'}</div>
        </div>`;
    }).join('');
}
