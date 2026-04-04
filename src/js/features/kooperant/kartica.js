// ============================================================
// KOOPERANT: KARTICA
// ============================================================
let karticaCache = null;

async function loadKartica() {
    const nameEl = document.getElementById('karticaName');
    const idEl = document.getElementById('karticaID');
    const listEl = document.getElementById('karticaList');

    if (nameEl) nameEl.textContent = CONFIG.ENTITY_NAME;
    if (idEl) idEl.textContent = CONFIG.ENTITY_ID;
    if (!listEl) return;

    if (karticaCache) {
        renderKartica(karticaCache);
        return;
    }

    listEl.innerHTML = '<p style="text-align:center;padding:20px;color:var(--text-muted);">Učitavanje...</p>';

    let records = [];
    const json = await safeAsync(async () => {
        return await apiFetch('action=getKartica&kooperantID=' + encodeURIComponent(CONFIG.ENTITY_ID));
    }, 'Greška pri učitavanju kartice');

    if (json && json.success && Array.isArray(json.records)) {
        records = json.records.filter(r => r.Opis !== 'UKUPNO');
    }

    karticaCache = records;
    renderKartica(records);
}

function renderKartica(records) {
    const listEl = document.getElementById('karticaList');
    const zadEl = document.getElementById('karticaZaduzenje');
    const razEl = document.getElementById('karticaRazduzenje');
    const saldoEl = document.getElementById('karticaSaldo');

    if (!listEl) return;

    if (!records.length) {
        listEl.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:40px;">Kartica nije dostupna</p>';
        if (zadEl) zadEl.textContent = '0';
        if (razEl) razEl.textContent = '0';
        if (saldoEl) saldoEl.textContent = '0';
        return;
    }

    let zad = 0;
    let raz = 0;
    let saldo = 0;

    listEl.innerHTML = records.map(r => {
        const z = parseFloat(r.Zaduzenje) || 0;
        const ra = parseFloat(r.Razduzenje) || 0;
        const s = parseFloat(r.Saldo) || 0;

        zad += z;
        raz += ra;
        saldo = s;

        return `<div class="queue-item" style="border-left-color:${z > 0 ? 'var(--danger)' : 'var(--success)'};">
            <div class="qi-header">
                <span class="qi-koop">${escapeHtml(r.BrojDok || '')}</span>
                <span class="qi-time">${escapeHtml(fmtDate(r.Datum))}</span>
            </div>
            <div class="qi-detail">${escapeHtml(r.Opis || '')}</div>
            <div class="qi-detail" style="font-size:12px;margin-top:2px;">
                ${z > 0 ? '<span style="color:var(--danger);">Zaduž: ' + z.toLocaleString('sr-RS') + '</span> ' : ''}
                ${ra > 0 ? '<span style="color:var(--success);">Razduž: ' + ra.toLocaleString('sr-RS') + '</span> ' : ''}
                | Saldo: <strong>${s.toLocaleString('sr-RS')}</strong>
            </div>
        </div>`;
    }).join('');

    if (zadEl) zadEl.textContent = zad.toLocaleString('sr-RS');
    if (razEl) razEl.textContent = raz.toLocaleString('sr-RS');
    if (saldoEl) saldoEl.textContent = saldo.toLocaleString('sr-RS');
}

function invalidateKarticaCache() {
    karticaCache = null;
}
