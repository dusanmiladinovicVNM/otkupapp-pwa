// ============================================================
// KOOPERANT: KARTICA
// ============================================================
let karticaCache = null;

async function loadKartica() {
    document.getElementById('karticaName').textContent = CONFIG.ENTITY_NAME;
    document.getElementById('karticaID').textContent = CONFIG.ENTITY_ID;

    if (karticaCache) {
        renderKartica(karticaCache);
        return;
    }

    document.getElementById('karticaList').innerHTML = '<p style="text-align:center;padding:20px;color:var(--text-muted);">Učitavanje...</p>';

    let records = [];
    const json = await safeAsync(async () => {
        return await apiFetch('action=getKartica&kooperantID=' + encodeURIComponent(CONFIG.ENTITY_ID));
    }, 'Greška pri učitavanju kartice');

    if (json && json.success && json.records) {
        records = json.records.filter(r => r.Opis !== 'UKUPNO');
    }

    karticaCache = records;
    renderKartica(records);
}

function renderKartica(records) {
    if (records.length === 0) {
        document.getElementById('karticaList').innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:40px;">Kartica nije dostupna</p>';
        ['karticaZaduzenje','karticaRazduzenje','karticaSaldo'].forEach(id => document.getElementById(id).textContent = '0');
        return;
    }
    let zad = 0, raz = 0;
    document.getElementById('karticaList').innerHTML = records.map(r => {
        const z = parseFloat(r.Zaduzenje)||0, ra = parseFloat(r.Razduzenje)||0, s = parseFloat(r.Saldo)||0;
        zad += z; raz += ra;
        return `<div class="queue-item" style="border-left-color:${z>0?'var(--danger)':'var(--success)'};">
            <div class="qi-header"><span class="qi-koop">${escapeHtml(r.BrojDok||'')}</span><span class="qi-time">${escapeHtml(fmtDate(r.Datum))}</span></div>
            <div class="qi-detail">${escapeHtml(r.Opis||'')}</div>
            <div class="qi-detail" style="font-size:12px;margin-top:2px;">
                ${z>0?'<span style="color:var(--danger);">Zaduž: '+z.toLocaleString('sr')+'</span> ':''}
                ${ra>0?'<span style="color:var(--success);">Razduž: '+ra.toLocaleString('sr')+'</span> ':''}
                | Saldo: <strong>${s.toLocaleString('sr')}</strong></div></div>`;
    }).join('');
    document.getElementById('karticaZaduzenje').textContent = zad.toLocaleString('sr');
    document.getElementById('karticaRazduzenje').textContent = raz.toLocaleString('sr');
    document.getElementById('karticaSaldo').textContent = (zad - raz).toLocaleString('sr');
}
