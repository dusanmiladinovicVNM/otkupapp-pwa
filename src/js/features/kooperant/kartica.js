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

    listEl.innerHTML = '<div class="kartica-loading">Učitavanje kartice...</div>';

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
        listEl.innerHTML = `
            <div class="kartica-empty">
                <div class="kartica-empty-title">Nema stavki kartice</div>
                <div class="kartica-empty-text">Trenutno nema dostupnih zaduženja ni razduženja.</div>
            </div>
        `;
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

        return `
            <div class="kartica-row">
                <div class="kartica-docblock">
                    <div class="kartica-doc">${escapeHtml(r.BrojDok || '-')}</div>
                    <div class="kartica-date">${escapeHtml(fmtDate(r.Datum))}</div>
                </div>

                <div class="kartica-opis">${escapeHtml(r.Opis || '-')}</div>

                <div class="kartica-amount kartica-amount--zad ${z > 0 ? 'is-danger' : ''}">
                    ${z > 0 ? z.toLocaleString('sr-RS') : '—'}
                </div>

                <div class="kartica-amount kartica-amount--raz ${ra > 0 ? 'is-success' : ''}">
                    ${ra > 0 ? ra.toLocaleString('sr-RS') : '—'}
                </div>

                <div class="kartica-amount kartica-amount--saldo is-saldo">
                    ${s.toLocaleString('sr-RS')}
                </div>
            </div>
        `;
    }).join('');

    if (zadEl) zadEl.textContent = zad.toLocaleString('sr-RS');
    if (razEl) razEl.textContent = raz.toLocaleString('sr-RS');
    if (saldoEl) saldoEl.textContent = saldo.toLocaleString('sr-RS');
}

function invalidateKarticaCache() {
    karticaCache = null;
}
