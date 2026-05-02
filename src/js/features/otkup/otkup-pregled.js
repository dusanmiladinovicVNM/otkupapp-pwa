// ============================================================
// OTKUP PREGLED / DANAS
// Finalna očišćena verzija
// ============================================================

const otkupPregledState = {
    quickFilter: 'danas',
    rows: [],
    detailRecord: null,
    eventsBound: false
};

async function loadOtkupPregled() {
    const listEl = byId('pregledList');
    if (!listEl) return;

    ensurePregledDefaultDates();
    bindPregledEventsOnce();

    setHtml(
        listEl,
        '<p style="text-align:center;padding:20px;color:var(--text-muted);">Učitavanje...</p>'
    );

    let localRows = [];
    let serverRows = [];

    try {
        if (db) {
            localRows = await dbGetAll(db, CONFIG.STORE_NAME);
        }
    } catch (err) {
        console.error('loadOtkupPregled local failed:', err);
    }

    if (navigator.onLine) {
        const json = await safeAsync(async () => {
            return await apiFetch('action=getOtkupi&otkupacID=' + encodeURIComponent(CONFIG.OTKUPAC_ID));
        }, 'Greška pri učitavanju otkupa');

        if (json && json.success && Array.isArray(json.records)) {
            serverRows = json.records.map(mapServerOtkupRecord);
        }
    }

    const mergedRows = mergePregledRecords(localRows, serverRows)
        .map(enrichPregledRecord)
        .sort(comparePregledRowsDesc);

    otkupPregledState.rows = mergedRows;

    rerenderPregled();
}

function ensurePregledDefaultDates() {
    const fldOd = byId('fldPregledOd');
    const fldDo = byId('fldPregledDo');
    const today = getTodayIsoDate();

    if (fldOd && !fldOd.value) fldOd.value = today;
    if (fldDo && !fldDo.value) fldDo.value = today;
}

function bindPregledEventsOnce() {
    if (otkupPregledState.eventsBound) return;

    const listEl = byId('pregledList');
    if (listEl) {
        listEl.addEventListener('click', function (e) {
            const card = e.target.closest('.danas-card');
            if (!card) return;

            const recordKey = card.getAttribute('data-record-key') || '';
            if (!recordKey) return;

            openPregledDetail(recordKey);
        });
    }

    otkupPregledState.eventsBound = true;
}

function onPregledDateChange() {
    otkupPregledState.quickFilter = 'custom';
    updatePregledQuickFiltersUI();
    rerenderPregled();
}

function setPregledQuickFilter(filterName, btn) {
    otkupPregledState.quickFilter = filterName;

    const fldOd = byId('fldPregledOd');
    const fldDo = byId('fldPregledDo');

    const today = getTodayIsoDate();
    const juce = getRelativeIsoDate(-1);

    if (filterName === 'danas') {
        if (fldOd) fldOd.value = today;
        if (fldDo) fldDo.value = today;
    } else if (filterName === 'juce') {
        if (fldOd) fldOd.value = juce;
        if (fldDo) fldDo.value = juce;
    } else if (filterName === 'sve') {
        if (fldOd) fldOd.value = '';
        if (fldDo) fldDo.value = '';
    } else if (filterName === 'bez_vozaca') {
        if (fldOd) fldOd.value = today;
        if (fldDo) fldDo.value = today;
    } else if (filterName === 'problemi') {
        if (fldOd) fldOd.value = today;
        if (fldDo) fldDo.value = today;
    }

    updatePregledQuickFiltersUI(btn);
    rerenderPregled();
}

function updatePregledQuickFiltersUI(btn) {
    qsa('#pregledQuickFilters .danas-pill').forEach(el => {
        removeClass(el, 'active');
    });

    if (btn && btn.classList) {
        addClass(btn, 'active');
        return;
    }

    const activeBtn = qs('#pregledQuickFilters .danas-pill[data-filter="' + otkupPregledState.quickFilter + '"]');
    if (activeBtn) addClass(activeBtn, 'active');
}

function rerenderPregled() {
    const listEl = byId('pregledList');
    if (!listEl) return;

    const rows = getFilteredPregledRows();
    renderOtkupPregledStats(rows);
    renderOtkupPregledList(listEl, rows);
}

function getFilteredPregledRows() {
    const fldOd = byId('fldPregledOd');
    const fldDo = byId('fldPregledDo');

    const od = fldOd ? fldOd.value : '';
    const doo = fldDo ? fldDo.value : '';

    let rows = otkupPregledState.rows.slice();

    if (od) rows = rows.filter(r => (r.datum || '') >= od);
    if (doo) rows = rows.filter(r => (r.datum || '') <= doo);

    if (otkupPregledState.quickFilter === 'bez_vozaca') {
        rows = rows.filter(r => !r.vozacID);
    }

    if (otkupPregledState.quickFilter === 'problemi') {
        rows = rows.filter(isPregledProblem);
    }

    return rows;
}

function mapServerOtkupRecord(r) {
    return {
        clientRecordID: r.ClientRecordID || '',
        serverRecordID: r.ServerRecordID || '',
        createdAtClient: normalizeIso(r.CreatedAtClient),
        updatedAtClient: normalizeIso(r.UpdatedAtClient || r.CreatedAtClient),
        updatedAtServer: normalizeIso(r.UpdatedAtServer || r.ReceivedAt),
        syncedAt: normalizeIso(r.UpdatedAtServer || r.ReceivedAt),

        datum: toIsoDateOnly(r.Datum),
        datumLabel: fmtDate(r.Datum),

        kooperantID: r.KooperantID || '',
        kooperantName: r.KooperantName || r.KooperantID || '',
        vrstaVoca: r.VrstaVoca || '',
        sortaVoca: r.SortaVoca || '',
        klasa: r.Klasa || 'I',
        kolicina: parseFloat(r.Kolicina) || 0,
        cena: parseFloat(r.Cena) || 0,
        kolAmbalaze: parseInt(r.KolAmbalaze, 10) || 0,
        tipAmbalaze: r.TipAmbalaze || '',
        parcelaID: r.ParcelaID || '',
        vozacID: r.VozacID || r.VozaciID || '',
        napomena: r.Napomena || '',

        syncStatus: 'synced',
        syncAttempts: 0,
        lastSyncError: '',
        lastServerStatus: 'server',
        deleted: false
    };
}

function normalizeLocalPregledRecord(r) {
    return {
        clientRecordID: r.clientRecordID || '',
        serverRecordID: r.serverRecordID || '',
        createdAtClient: normalizeIso(r.createdAtClient),
        updatedAtClient: normalizeIso(r.updatedAtClient || r.createdAtClient),
        updatedAtServer: normalizeIso(r.updatedAtServer),
        syncedAt: normalizeIso(r.syncedAt),

        datum: toIsoDateOnly(r.datum || ''),
        datumLabel: fmtDate(r.datum || ''),

        kooperantID: r.kooperantID || '',
        kooperantName: r.kooperantName || r.kooperantID || '',
        vrstaVoca: r.vrstaVoca || '',
        sortaVoca: r.sortaVoca || '',
        klasa: r.klasa || 'I',
        kolicina: parseFloat(r.kolicina) || 0,
        cena: parseFloat(r.cena) || 0,
        kolAmbalaze: parseInt(r.kolAmbalaze, 10) || 0,
        tipAmbalaze: r.tipAmbalaze || '',
        parcelaID: r.parcelaID || '',
        vozacID: r.vozacID || '',
        napomena: r.napomena || '',

        syncStatus: r.syncStatus || 'pending',
        syncAttempts: parseInt(r.syncAttempts, 10) || 0,
        lastSyncError: r.lastSyncError || '',
        lastServerStatus: r.lastServerStatus || '',
        deleted: !!r.deleted
    };
}

function mergePregledRecords(localRows, serverRows) {
    const map = new Map();

    serverRows.forEach(row => {
        const key = getPregledRecordKey(row);
        if (!key) return;
        map.set(key, row);
    });

    localRows
        .map(normalizeLocalPregledRecord)
        .forEach(row => {
            const key = getPregledRecordKey(row);
            if (!key) return;

            const existing = map.get(key);

            if (!existing) {
                map.set(key, row);
                return;
            }

            // Ako lokalni zapis nije sinhronizovan ili ima grešku, lokalni ima prednost
            if (row.syncStatus !== 'synced' || row.lastSyncError) {
                map.set(key, row);
                return;
            }

            // Ako je lokalni noviji od server rendera, uzmi lokalni
            if ((row.updatedAtClient || '') > (existing.updatedAtClient || '')) {
                map.set(key, row);
            }
        });

    return Array.from(map.values()).filter(r => !r.deleted);
}

function getPregledRecordKey(r) {
    if (r.serverRecordID) return 'srv:' + r.serverRecordID;
    if (r.clientRecordID) return 'cli:' + r.clientRecordID;
    return '';
}

function enrichPregledRecord(r) {
    const vrednost = (parseFloat(r.kolicina) || 0) * (parseFloat(r.cena) || 0);

    return {
        ...r,
        datum: toIsoDateOnly(r.datum || ''),
        datumLabel: r.datumLabel || fmtDate(r.datum || ''),
        vrednost: vrednost,
        statusMeta: buildPregledStatusMeta(r)
    };
}

function comparePregledRowsDesc(a, b) {
    const byDate = String(b.datum || '').localeCompare(String(a.datum || ''));
    if (byDate !== 0) return byDate;

    const aTime = a.updatedAtClient || a.createdAtClient || a.updatedAtServer || '';
    const bTime = b.updatedAtClient || b.createdAtClient || b.updatedAtServer || '';
    const byTime = String(bTime).localeCompare(String(aTime));
    if (byTime !== 0) return byTime;

    return String(b.clientRecordID || '').localeCompare(String(a.clientRecordID || ''));
}

function renderOtkupPregledStats(rows) {
    const countEl = byId('statPregledCount');
    const kgEl = byId('statPregledKg');
    const vrednostEl = byId('statPregledVrednost');
    const koopEl = byId('statPregledKoop');

    const count = rows.length;
    const kg = rows.reduce((sum, r) => sum + (parseFloat(r.kolicina) || 0), 0);
    const vrednost = rows.reduce((sum, r) => sum + (parseFloat(r.vrednost) || 0), 0);
    const koopCount = new Set(rows.map(r => r.kooperantID).filter(Boolean)).size;

    setText(countEl, String(count));
    setText(kgEl, kg.toLocaleString('sr-RS'));
    setText(vrednostEl, vrednost.toLocaleString('sr-RS'));
    setText(koopEl, String(koopCount));
}

function renderOtkupPregledList(listEl, rows) {
    if (!rows.length) {
        setHtml(listEl, `
            <div class="danas-empty">
                <div class="danas-empty-title">Nema otkupa za izabrani prikaz</div>
                <div class="danas-empty-text">Promeni filter ili datum da bi video stavke.</div>
            </div>
        `);
        return;
    }

    const sections = buildPregledSections(rows)
        .filter(section => section.items.length > 0);

    setHtml(listEl, sections.map(section => `
        <section class="danas-section">
            <div class="danas-section-head">
                <div class="danas-section-title">${escapeHtml(section.title)}</div>
                <div class="danas-section-count">${section.items.length}</div>
            </div>

            <div class="danas-cards">
                ${section.items.map(renderPregledCard).join('')}
            </div>
        </section>
    `).join(''));
}

function buildPregledSections(rows) {
    const problemi = rows.filter(isPregledProblem);
    const bezVozaca = rows.filter(r => !r.vozacID && !isPregledProblem(r));
    const saVozacem = rows.filter(r => !!r.vozacID && !isPregledProblem(r));

    return [
        { key: 'bez_vozaca', title: 'Danas bez vozača', items: bezVozaca },
        { key: 'sa_vozacem', title: 'Danas sa vozačem', items: saVozacem },
        { key: 'problemi', title: 'Sync problemi', items: problemi }
    ];
}

function renderPregledCard(r) {
    const key = getPregledRecordKey(r);
    const status = r.statusMeta || buildPregledStatusMeta(r);
    const sortaPart = r.sortaVoca ? ' / ' + r.sortaVoca : '';
    const parcelaPart = r.parcelaID ? ' • Parcela ' + r.parcelaID : '';

    return `
        <button type="button" class="danas-card" data-record-key="${escapeHtml(key)}">
            <div class="danas-card-top">
                <div class="danas-card-koop">${escapeHtml(r.kooperantName || '-')}</div>
                <div class="danas-card-date">${escapeHtml(r.datumLabel || r.datum || '-')}</div>
            </div>

            <div class="danas-card-main">
                <div class="danas-card-line">
                    ${escapeHtml(r.vrstaVoca || '-')}
                    ${escapeHtml(sortaPart)}
                    <span class="danas-card-class">Kl. ${escapeHtml(r.klasa || 'I')}</span>
                </div>

                <div class="danas-card-line danas-card-line--muted">
                    ${escapeHtml(formatKg(r.kolicina))} × ${escapeHtml(formatMoney(r.cena))}
                    ${escapeHtml(parcelaPart)}
                </div>
            </div>

            <div class="danas-card-bottom">
                <span class="danas-badge danas-badge--${escapeHtml(status.kind)}">${escapeHtml(status.label)}</span>
                <strong class="danas-card-value">${escapeHtml(formatMoney(r.vrednost))}</strong>
            </div>
        </button>
    `;
}

function buildPregledStatusMeta(r) {
    if (isPregledProblem(r)) {
        return { kind: 'problem', label: 'Sync problem' };
    }

    if (!r.vozacID) {
        return { kind: 'warning', label: 'Bez vozača' };
    }

    return { kind: 'success', label: 'Dodeljen' };
}

function isPregledProblem(r) {
    return !!(
        (r.syncStatus && r.syncStatus !== 'synced') ||
        (r.lastSyncError && String(r.lastSyncError).trim() !== '')
    );
}

function openPregledDetail(recordKey) {
    const row = otkupPregledState.rows.find(r => getPregledRecordKey(r) === recordKey);
    if (!row) return;

    const modal = byId('pregledDetailModal');
    const dateEl = byId('pregledDetailDate');
    const gridEl = byId('pregledDetailGrid');
    if (!modal || !dateEl || !gridEl) return;

    otkupPregledState.detailRecord = row;

    setText(dateEl, row.datumLabel || row.datum || '-');

    setHtml(gridEl, [
        ['Kooperant', row.kooperantName || '-'],
        ['Kooperant ID', row.kooperantID || '-'],
        ['Vrsta', row.vrstaVoca || '-'],
        ['Sorta', row.sortaVoca || '-'],
        ['Klasa', row.klasa || '-'],
        ['Količina', formatKg(row.kolicina)],
        ['Cena', formatMoney(row.cena)],
        ['Vrednost', formatMoney(row.vrednost)],
        ['Vozač', row.vozacID || 'Nije dodeljen'],
        ['Ambalaža', row.kolAmbalaze ? row.kolAmbalaze + ' kom' : '-'],
        ['Tip ambalaže', row.tipAmbalaze || '-'],
        ['Parcela', row.parcelaID || '-'],
        ['Sync', (row.statusMeta || buildPregledStatusMeta(row)).label],
        ['Napomena', row.napomena || '-']
    ].map(([label, value]) => `
        <div class="danas-detail-item">
            <span>${escapeHtml(label)}</span>
            <strong>${escapeHtml(String(value))}</strong>
        </div>
    `).join(''));

    addClass(modal, 'visible');
}

function closePregledDetail() {
    const modal = byId('pregledDetailModal');
    if (modal) removeClass(modal, 'visible');
    otkupPregledState.detailRecord = null;
}

function openPregledDetailOtkupniList() {
    if (!otkupPregledState.detailRecord) return;
    if (typeof showOtkupniList === 'function') {
        showOtkupniList(otkupPregledState.detailRecord);
    }
}
