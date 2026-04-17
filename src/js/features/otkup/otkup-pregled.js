// ============================================================
// OTKUP PREGLED / DANAS
// ============================================================

const otkupPregledState = {
    quickFilter: 'danas',
    rows: [],
    detailRecord: null
};

async function loadOtkupPregled() {
    const list = document.getElementById('pregledList');
    if (!list) return;

    ensurePregledDefaultDates();
    bindPregledEventsOnce();

    list.innerHTML = '<p style="text-align:center;padding:20px;color:var(--text-muted);">Učitavanje...</p>';

    let local = [];
    let server = [];

    try {
        if (db) {
            local = await dbGetAll(db, CONFIG.STORE_NAME);
        }
    } catch (err) {
        console.error('loadOtkupPregled local failed:', err);
    }

    if (navigator.onLine) {
        const json = await safeAsync(async () => {
            return await apiFetch('action=getOtkupi&otkupacID=' + encodeURIComponent(CONFIG.OTKUPAC_ID));
        }, 'Greška pri učitavanju otkupa');

        if (json && json.success && Array.isArray(json.records)) {
            server = json.records.map(mapServerOtkupRecord);
        }
    }

    const merged = mergeOtkupPregledRecords(local, server)
        .map(enrichPregledRecord)
        .sort(compareOtkupPregledRecordsDesc);

    otkupPregledState.rows = merged;

    const filtered = getFilteredPregledRows();
    renderOtkupPregledStats(filtered);
    renderOtkupPregledList(list, filtered);
    updatePregledQuickFiltersUI();
}

function ensurePregledDefaultDates() {
    const fldOd = document.getElementById('fldPregledOd');
    const fldDo = document.getElementById('fldPregledDo');
    const today = getTodayIsoDate();

    if (fldOd && !fldOd.value) fldOd.value = today;
    if (fldDo && !fldDo.value) fldDo.value = today;
}

function bindPregledEventsOnce() {
    const root = document.getElementById('tab-pregled');
    if (!root || root.dataset.boundPregled === '1') return;

    root.dataset.boundPregled = '1';
}

function onPregledDateChange() {
    otkupPregledState.quickFilter = 'custom';
    updatePregledQuickFiltersUI();
    rerenderPregled();
}

function setPregledQuickFilter(filterName, btn) {
    otkupPregledState.quickFilter = filterName;

    const fldOd = document.getElementById('fldPregledOd');
    const fldDo = document.getElementById('fldPregledDo');

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

function rerenderPregled() {
    const list = document.getElementById('pregledList');
    if (!list) return;

    const filtered = getFilteredPregledRows();
    renderOtkupPregledStats(filtered);
    renderOtkupPregledList(list, filtered);
}

function updatePregledQuickFiltersUI(btn) {
    document.querySelectorAll('#pregledQuickFilters .danas-pill').forEach(el => {
        el.classList.remove('active');
    });

    if (btn && btn.classList) {
        btn.classList.add('active');
        return;
    }

    const active = document.querySelector(`#pregledQuickFilters .danas-pill[data-filter="${otkupPregledState.quickFilter}"]`);
    if (active) active.classList.add('active');
}

function getFilteredPregledRows() {
    const fldOd = document.getElementById('fldPregledOd');
    const fldDo = document.getElementById('fldPregledDo');

    const od = fldOd ? fldOd.value : '';
    const doo = fldDo ? fldDo.value : '';

    let all = [...otkupPregledState.rows];

    if (od) all = all.filter(r => (r.datum || '') >= od);
    if (doo) all = all.filter(r => (r.datum || '') <= doo);

    if (otkupPregledState.quickFilter === 'bez_vozaca') {
        all = all.filter(r => !r.vozacID);
    }

    if (otkupPregledState.quickFilter === 'problemi') {
        all = all.filter(r => isPregledProblem(r));
    }

    return all;
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
        datumLabel: safeFmtDate(r.Datum),

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
        lastServerStatus: 'server'
    };
}

function mergeOtkupPregledRecords(local, server) {
    return mergeOfflineRecords(local, server, normalizeLocalPregledRecord);
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
        datumLabel: safeFmtDate(r.datum || ''),

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
        lastServerStatus: r.lastServerStatus || ''
    };
}

function enrichPregledRecord(r) {
    const vrednost = ((parseFloat(r.kolicina) || 0) * (parseFloat(r.cena) || 0));
    return {
        ...r,
        datum: toIsoDateOnly(r.datum || ''),
        datumLabel: r.datumLabel || safeFmtDate(r.datum || ''),
        vrednost,
        statusMeta: buildPregledStatusMeta(r)
    };
}

function compareOtkupPregledRecordsDesc(a, b) {
    const aTime = a.updatedAtClient || a.createdAtClient || a.updatedAtServer || '';
    const bTime = b.updatedAtClient || b.createdAtClient || b.updatedAtServer || '';

    const byDate = String(b.datum || '').localeCompare(String(a.datum || ''));
    if (byDate !== 0) return byDate;

    const byTime = String(bTime).localeCompare(String(aTime));
    if (byTime !== 0) return byTime;

    return String(b.clientRecordID || '').localeCompare(String(a.clientRecordID || ''));
}

function renderOtkupPregledStats(all) {
    const countEl = document.getElementById('statPregledCount');
    const kgEl = document.getElementById('statPregledKg');
    const vrednostEl = document.getElementById('statPregledVrednost');
    const koopEl = document.getElementById('statPregledKoop');

    const kg = all.reduce((s, r) => s + (parseFloat(r.kolicina) || 0), 0);
    const vr = all.reduce((s, r) => s + (parseFloat(r.vrednost) || 0), 0);
    const koopCount = new Set(all.map(r => r.kooperantID).filter(Boolean)).size;

    if (countEl) countEl.textContent = String(all.length);
    if (kgEl) kgEl.textContent = kg.toLocaleString('sr-RS');
    if (vrednostEl) vrednostEl.textContent = vr.toLocaleString('sr-RS');
    if (koopEl) koopEl.textContent = String(koopCount);
}

function renderOtkupPregledList(list, all) {
    if (!all.length) {
        list.innerHTML = `
            <div class="danas-empty">
                <div class="danas-empty-title">Nema otkupa za izabrani prikaz</div>
                <div class="danas-empty-text">Promenite filter ili datum da biste videli stavke.</div>
            </div>
        `;
        return;
    }

    const sections = buildPregledSections(all);

    list.innerHTML = sections
        .filter(section => section.items.length > 0)
        .map(section => `
            <section class="danas-section">
                <div class="danas-section-head">
                    <div class="danas-section-title">${escapeHtml(section.title)}</div>
                    <div class="danas-section-count">${section.items.length}</div>
                </div>

                <div class="danas-cards">
                    ${section.items.map(renderPregledCard).join('')}
                </div>
            </section>
        `).join('');
}

function buildPregledSections(all) {
    const problemi = all.filter(isPregledProblem);
    const bezVozaca = all.filter(r => !r.vozacID && !isPregledProblem(r));
    const saVozacem = all.filter(r => !!r.vozacID && !isPregledProblem(r));

    return [
        { key: 'bez_vozaca', title: 'Danas bez vozača', items: bezVozaca },
        { key: 'sa_vozacem', title: 'Danas sa vozačem', items: saVozacem },
        { key: 'problemi', title: 'Sync problemi', items: problemi }
    ];
}

function renderPregledCard(r) {
    const sortaPart = r.sortaVoca ? ` / ${r.sortaVoca}` : '';
    const parcelaPart = r.parcelaID ? ` • Parcela ${r.parcelaID}` : '';
    const status = r.statusMeta || buildPregledStatusMeta(r);

    return `
        <button type="button" class="danas-card" onclick="openPregledDetail('${escapeHtmlJs(r.clientRecordID || r.serverRecordID || '')}')">
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
                <span class="danas-badge danas-badge--${status.kind}">${escapeHtml(status.label)}</span>
                <strong class="danas-card-value">${escapeHtml(formatMoney(r.vrednost))}</strong>
            </div>
        </button>
    `;
}

function buildPregledStatusMeta(r) {
    if (isPregledProblem(r)) {
        return { kind: 'problem', label: 'Sync problem', color: '#c62828' };
    }

    if (!r.vozacID) {
        return { kind: 'warning', label: 'Bez vozača', color: '#ef6c00' };
    }

    return { kind: 'success', label: 'Dodeljen', color: '#2e7d32' };
}

function isPregledProblem(r) {
    return !!(
        (r.syncStatus && r.syncStatus !== 'synced') ||
        (r.lastSyncError && String(r.lastSyncError).trim() !== '')
    );
}

function openPregledDetail(recordKey) {
    const modal = document.getElementById('pregledDetailModal');
    const dateEl = document.getElementById('pregledDetailDate');
    const grid = document.getElementById('pregledDetailGrid');

    if (!modal || !dateEl || !grid) return;

    const row = otkupPregledState.rows.find(r =>
        (r.clientRecordID && r.clientRecordID === recordKey) ||
        (r.serverRecordID && r.serverRecordID === recordKey)
    );

    if (!row) return;

    otkupPregledState.detailRecord = row;

    dateEl.textContent = row.datumLabel || row.datum || '-';

    grid.innerHTML = [
        ['Kooperant', row.kooperantName || '-'],
        ['Kooperant ID', row.kooperantID || '-'],
        ['Vrsta', row.vrstaVoca || '-'],
        ['Sorta', row.sortaVoca || '-'],
        ['Klasa', row.klasa || '-'],
        ['Količina', formatKg(row.kolicina)],
        ['Cena', formatMoney(row.cena)],
        ['Vrednost', formatMoney(row.vrednost)],
        ['Vozač', row.vozacID || 'Nije dodeljen'],
        ['Ambalaža', row.kolAmbalaze ? `${row.kolAmbalaze} kom` : '-'],
        ['Tip ambalaže', row.tipAmbalaze || '-'],
        ['Parcela', row.parcelaID || '-'],
        ['Sync', (row.statusMeta || buildPregledStatusMeta(row)).label],
        ['Napomena', row.napomena || '-']
    ].map(([label, value]) => `
        <div class="danas-detail-item">
            <span>${escapeHtml(label)}</span>
            <strong>${escapeHtml(String(value))}</strong>
        </div>
    `).join('');

    modal.classList.add('visible');
}

function closePregledDetail() {
    const modal = document.getElementById('pregledDetailModal');
    if (modal) modal.classList.remove('visible');
    otkupPregledState.detailRecord = null;
}

function openPregledDetailOtkupniList() {
    if (!otkupPregledState.detailRecord) return;
    if (typeof showOtkupniList === 'function') {
        showOtkupniList(otkupPregledState.detailRecord);
    }
}

function getTodayIsoDate() {
    return new Date().toISOString().slice(0, 10);
}

function getRelativeIsoDate(offsetDays) {
    const d = new Date();
    d.setDate(d.getDate() + offsetDays);
    return d.toISOString().slice(0, 10);
}

function toIsoDateOnly(input) {
    if (!input) return '';

    const s = String(input).trim();

    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

    if (/^\d{2}\.\d{2}\.\d{4}\.?$/.test(s)) {
        const clean = s.replace(/\.$/, '');
        const [dd, mm, yyyy] = clean.split('.');
        return `${yyyy}-${mm}-${dd}`;
    }

    const d = new Date(s);
    if (!isNaN(d.getTime())) {
        return d.toISOString().slice(0, 10);
    }

    return s;
}

function safeFmtDate(input) {
    try {
        return typeof fmtDate === 'function' ? fmtDate(input) : (toIsoDateOnly(input) || '');
    } catch (_) {
        return toIsoDateOnly(input) || '';
    }
}

function formatMoney(v) {
    return `${(parseFloat(v) || 0).toLocaleString('sr-RS')} RSD`;
}

function formatKg(v) {
    return `${(parseFloat(v) || 0).toLocaleString('sr-RS')} kg`;
}

function escapeHtmlJs(value) {
    return String(value || '')
        .replaceAll('\\', '\\\\')
        .replaceAll("'", "\\'")
        .replaceAll('"', '\\"');
}
