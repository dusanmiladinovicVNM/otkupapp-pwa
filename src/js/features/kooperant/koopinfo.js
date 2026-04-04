// ============================================================
// KOOPERANT: INFO
// ============================================================
async function loadKoopInfo() {
    const config = stammdaten.config || [];
    const gv = k => { const c = config.find(c => c.Parameter === k); return c ? c.Vrednost : '-'; };
    document.getElementById('koopInfoContent').innerHTML = `
        <div style="background:white;border-radius:var(--radius);padding:16px;margin-bottom:12px;box-shadow:0 1px 3px rgba(0,0,0,0.08);">
            <h3 style="color:var(--primary);margin-bottom:12px;">Otkup informacije</h3>
            <table style="width:100%;">
                <tr><td style="padding:8px;color:var(--text-muted);">Status:</td><td style="padding:8px;font-weight:600;">${gv('OtkupAktivan')==='Da'?'🟢 Aktivan':'🔴 Neaktivan'}</td></tr>
                <tr><td style="padding:8px;color:var(--text-muted);">Radno vreme:</td><td style="padding:8px;">${escapeHtml(gv('RadnoVremeOd'))} - ${escapeHtml(gv('RadnoVremeDo'))}</td></tr>
                <tr><td style="padding:8px;color:var(--text-muted);">Sezona:</td><td style="padding:8px;">${escapeHtml(gv('SezonaOd'))} - ${escapeHtml(gv('SezonaDo'))}</td></tr>
            </table>
        </div>
        <div style="background:white;border-radius:var(--radius);padding:16px;box-shadow:0 1px 3px rgba(0,0,0,0.08);">
            <h3 style="color:var(--primary);margin-bottom:12px;">Aktuelne cene</h3>
            <table style="width:100%;">
                ${config.filter(c => c.Parameter && c.Parameter.startsWith('Cena')).map(c =>
                    '<tr><td style="padding:8px;color:var(--text-muted);">'+escapeHtml(c.Parameter.replace('Cena',''))+':</td><td style="padding:8px;font-weight:600;">'+escapeHtml(c.Vrednost)+' RSD/kg</td></tr>').join('')}
            </table>
        </div>`;
}
