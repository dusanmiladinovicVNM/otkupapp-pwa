// ============================================================
// OTPREMA
// UI first + osnovna lokalna logika dodele
// ============================================================

const otpremaState = {
    rows: [],
    selectedVozac: null,
    selectedKeys: new Set(),
    successRows: [],
    eventsBound: false
};

async function loadOtpremaOverview() {
    bindOtpremaEventsOnce();
    populateOtpremaFallbackDrivers();
    showOtpremaRootView();

    const rootEl = byId('otpremaRootSections');
    if (rootEl) {
        setHtml(rootEl, '<p style="text-align:center;padding:20px;color:var(--text-muted);">Učitavanje...</p>');
    }

    let localRows = [];
    let serverRows = [];

    try {
        if (db) {
            localRows = await dbGetAll(db, CONFIG.STORE_NAME);
        }
    } catch (err) {
        console.error('loadOtpremaOverview local failed:', err);
    }

    if (navigator.onLine) {
        const json = await safeAsync(async () => {
            return await apiFetch('action=getOtkupi&otkupacID=' + encodeURIComponent(CONFIG.OTKUPAC_ID));
        }, 'Greška pri učitavanju otpreme');

        if (json && json.success && Array.isArray(json.records)) {
            serverRows = json.records.map(mapServerOtpremaRecord);
        }
    }

    const mergedRows = mergeOtpremaRecords(localRows, serverRows)
        .map(enrichOtpremaRecord)
        .sort(compareOtpremaRowsDesc);

    otpremaState.rows = mergedRows;
    renderOtpremaRoot();
}

function bindOtpremaEventsOnce() {
    if (otpremaState.eventsBound) return;

    const rootSections = byId('otpremaRootSections');
    if (rootSections) {
        rootSections.addEventListener('click', function (e) {
            const card = e.target.closest('.otprema-card');
            if (!card) return;

            const key = card.getAttribute('data-record-key') || '';
            if (!key) return;

            openOtpremaDetail(key);
        });
    }

    const assignSections = byId('otpremaAssignSections');
    if (assignSections) {
        assignSections.addEventListener('change', function (e) {
            const checkbox = e.target.closest('.otprema-check');
            if (!checkbox) return;

            const key = checkbox.getAttribute('data-record-key') || '';
            if (!key) return;

            if (checkbox.checked) {
                otpremaState.selectedKeys.add(key);
            } else {
                otpremaState.selectedKeys.delete(key);
            }

            updateOtpremaAssignSummary();
        });
    }

    otpremaState.eventsBound = true;
}

function populateOtpremaFallbackDrivers() {
    const sel = byId('fldOtpremaFallbackVozac');
    if (!sel) return;

    const current = sel.value || '';
    sel.innerHTML = '<option value="">-- Izaberi vozača --</option>';

    (stammdaten.vozaci || []).forEach(v => {
        const id = v.VozacID || v.ID || '';
        const name = v.ImePrezime || v.Naziv || v.Ime || id;
        if (!id) return;

        const opt = document.createElement('option');
        opt.value = id;
        opt.textContent = name + ' (' + id + ')';
        sel.appendChild(opt);
    });

    if (current) sel.value = current;
}

function toggleOtpremaFallback() {
    const panel = byId('otpremaFallbackPanel');
    if (!panel) return;

    panel.style.display = panel.style.display === 'none' ? 'block' : 'none';
}

function applyOtpremaFallbackDriver() {
    const sel = byId('fldOtpremaFallbackVozac');
    if (!sel || !sel.value) {
        showToast('Izaberi vozača', 'error');
        return;
    }

    const id = sel.value;
    const vozac = (stammdaten.vozaci || []).find(v => (v.VozacID || v.ID) === id);
    const name = vozac ? (vozac.ImePrezime || vozac.Naziv || vozac.Ime || id) : id;

    setOtpremaVozac(id, name);
}

function startOtpremaVozacQRScan() {
    const readerDiv = byId('qr-reader-otprema-vozac');
    if (!readerDiv) return;

    readerDiv.style.display = 'block';

    const scanner = new Html5Qrcode('qr-reader-otprema-vozac');
    scanner.start(
        { facingMode: 'environment' },
        { fps: 10, qrbox: { width: 260, height: 260 } },
        (decodedText) => {
            onOtpremaVozacQRScanned(decodedText);
            scanner.stop().then(() => {
                readerDiv.style.display = 'none';
            }).catch(() => {
                readerDiv.style.display = 'none';
            });
        },
        () => {}
    ).catch(err => {
        showToast('Kamera nije dostupna: ' + err, 'error');
        readerDiv.style.display = 'none';
    });
}

function onOtpremaVozacQRScanned(text) {
    try {
        const data = JSON.parse(text);
        if (data.type === 'VOZ' && data.id) {
            setOtpremaVozac(data.id, data.name || data.id);
            return;
        }
    } catch (_) {}

    if (String(text).startsWith('VOZ-')) {
        const vozac = (stammdaten.vozaci || []).find(v => (v.VozacID || v.ID) === text);
        const name = vozac ? (vozac.ImePrezime || vozac.Naziv || vozac.Ime || text) : text;
        setOtpremaVozac(text, name);
        return;
    }

    showToast('Nije QR vozača', 'error');
}

function setOtpremaVozac(id, name) {
    otpremaState.selectedVozac = { id, name };
    otpremaState.selectedKeys.clear();
    openOtpremaAssignView();
}

function cancelOtpremaAssign() {
    otpremaState.selectedKeys.clear();
    otpremaState.selectedVozac = null;
    showOtpremaRootView();
    renderOtpremaRoot();
}

function showOtpremaRootView() {
    showEl(byId('otpremaRootView'), '');
    hideEl(byId('otpremaAssignView'));
    hideEl(byId('otpremaSuccessView'));
}

function openOtpremaAssignView() {
    hideEl(byId('otpremaRootView'));
    showEl(byId('otpremaAssignView'), '');
    hideEl(byId('otpremaSuccessView'));

    renderOtpremaAssignView();
}

function showOtpremaSuccessView() {
    hideEl(byId('otpremaRootView'));
    hideEl(byId('otpremaAssignView'));
    showEl(byId('otpremaSuccessView'), '');
}

function renderOtpremaRoot() {
    const sectionsEl = byId('otpremaRootSections');
    if (!sectionsEl) return;

    const sections = buildOtpremaRootSections();
    renderOtpremaSummary(sections);

    const html = `
        ${renderOtpremaSection('Današnji bez vozača', sections.todayUnassigned, true)}
        ${renderOtpremaSection('Raniji bez vozača', sections.olderUnassigned, true)}
        ${renderOtpremaAssignedGroups(sections.todayAssignedGroups)}
    `;

    setHtml(sectionsEl, html);
}

function renderOtpremaSummary(sections) {
    setText(byId('otpremaTodayUnassignedCount'), String(sections.todayUnassigned.length));
    setText(byId('otpremaTodayUnassignedKg'), formatOtpremaKg(sumOtpremaKg(sections.todayUnassigned)));

    setText(byId('otpremaOlderUnassignedCount'), String(sections.olderUnassigned.length));
    setText(byId('otpremaOlderUnassignedKg'), formatOtpremaKg(sumOtpremaKg(sections.olderUnassigned)));

    const todayAssignedRows = sections.todayAssignedGroups.flatMap(g => g.items);
    setText(byId('otpremaTodayAssignedCount'), String(todayAssignedRows.length));
    setText(byId('otpremaTodayAssignedKg'), formatOtpremaKg(sumOtpremaKg(todayAssignedRows)));
}

function buildOtpremaRootSections() {
    const today = getTodayIsoDate();

    const todayUnassigned = otpremaState.rows.filter(r =>
        !r.vozacID &&
        r.datum === today
    );

    const olderUnassigned = otpremaState.rows.filter(r =>
        !r.vozacID &&
        r.datum !== today
    );

    const todayAssigned = otpremaState.rows.filter(r =>
        !!r.vozacID &&
        getOtpremaAssignedDate(r) === today
    );

    const groupsMap = new Map();

    todayAssigned.forEach(row => {
        const key = row.vozacID || 'NEPOZNAT';

        if (!groupsMap.has(key)) {
            groupsMap.set(key, {
                vozacID: row.vozacID || '',
                vozacName: resolveVozacName(row.vozacID),
                items: []
            });
        }

        groupsMap.get(key).items.push(row);
    });

    const todayAssignedGroups = Array.from(groupsMap.values()).sort((a, b) =>
        String(a.vozacName || '').localeCompare(String(b.vozacName || ''))
    );

    return {
        todayUnassigned,
        olderUnassigned,
        todayAssignedGroups
    };
}

function renderOtpremaSection(title, items, showWarnings) {
    if (!items.length) {
        return `
            <section class="otprema-section">
                <div class="otprema-section-head">
                    <div class="otprema-section-title">${escapeHtml(title)}</div>
                    <div class="otprema-section-count">0</div>
                </div>
                <div class="otprema-empty">Nema stavki</div>
            </section>
        `;
    }

    return `
        <section class="otprema-section">
            <div class="otprema-section-head">
                <div class="otprema-section-title">${escapeHtml(title)}</div>
                <div class="otprema-section-count">${items.length}</div>
            </div>
            <div class="otprema-cards">
                ${items.map(row => renderOtpremaCard(row, showWarnings)).join('')}
            </div>
        </section>
    `;
}

function renderOtpremaAssignedGroups(groups) {
    if (!groups.length) {
        return `
            <section class="otprema-section">
                <div class="otprema-section-head">
                    <div class="otprema-section-title">Danas otpremljeno</div>
                    <div class="otprema-section-count">0</div>
                </div>
                <div class="otprema-empty">Nema današnjih dodela</div>
            </section>
        `;
    }

    return `
        <section class="otprema-section">
            <div class="otprema-section-head">
                <div class="otprema-section-title">Danas otpremljeno</div>
                <div class="otprema-section-count">${groups.reduce((s, g) => s + g.items.length, 0)}</div>
            </div>

            <div class="otprema-groups">
                ${groups.map(group => `
                    <div class="otprema-driver-group">
                        <div class="otprema-driver-group-head">
                            <div>
                                <div class="otprema-driver-group-name">${escapeHtml(group.vozacName)}</div>
                                <div class="otprema-driver-group-sub">${escapeHtml(group.vozacID)} • ${group.items.length} stavki • ${escapeHtml(formatOtpremaKg(sumOtpremaKg(group.items)))}</div>
                            </div>
                        </div>
                        <div class="otprema-cards">
                            ${group.items.map(row => renderOtpremaCard(row, false, true)).join('')}
                        </div>
                    </div>
                `).join('')}
            </div>
        </section>
    `;
}

function renderOtpremaCard(row, showWarning, isAssigned) {
    const key = getOtpremaRecordKey(row);
    const note = row.napomena ? `<div class="otprema-card-note">${escapeHtml(row.napomena)}</div>` : '';
    const warning = showWarning && row.datum !== getTodayIsoDate()
        ? `<span class="otprema-badge otprema-badge--warning">Raniji otkup</span>`
        : '';

    const statusBadges = [];

    if (isAssigned) {
        statusBadges.push(`<span class="otprema-badge otprema-badge--success">Dodeljen</span>`);

        if (row.syncStatus && row.syncStatus !== 'synced') {
            statusBadges.push(`<span class="otprema-badge otprema-badge--pending">Čeka sync</span>`);
        }

        if (row.lastSyncError) {
            statusBadges.push(`<span class="otprema-badge otprema-badge--error">Sync greška</span>`);
        }
    } else {
        statusBadges.push(`<span class="otprema-badge otprema-badge--pending">Bez vozača</span>`);
    }

    if (warning) statusBadges.push(warning);

    return `
        <div class="otprema-card" data-record-key="${escapeHtml(key)}">
            <div class="otprema-card-top">
                <div class="otprema-card-koop">${escapeHtml(row.kooperantName || '-')}</div>
                <div class="otprema-card-date">${escapeHtml(row.datum || '-')}</div>
            </div>

            <div class="otprema-card-main">
                <div class="otprema-card-line">
                    ${escapeHtml(row.vrstaVoca || '-')}
                    ${row.sortaVoca ? ' / ' + escapeHtml(row.sortaVoca) : ''}
                    <span class="otprema-card-class">Kl. ${escapeHtml(row.klasa || 'I')}</span>
                </div>

                <div class="otprema-card-line otprema-card-line--muted">
                    ${escapeHtml(formatOtpremaKg(row.kolicina))} • ${escapeHtml(formatOtpremaAmbalaza(row))}
                </div>

                ${note}
            </div>

            <div class="otprema-card-bottom">
                <div class="otprema-card-badges">
                    ${statusBadges.join('')}
                </div>
            </div>
        </div>
    `;
}

function renderOtpremaAssignView() {
    const driverCard = byId('otpremaAssignDriverCard');
    const sectionsEl = byId('otpremaAssignSections');
    if (!driverCard || !sectionsEl || !otpremaState.selectedVozac) return;

    setHtml(driverCard, `
        <div class="otprema-driver-card-label">Vozač</div>
        <div class="otprema-driver-card-name">${escapeHtml(otpremaState.selectedVozac.name)}</div>
        <div class="otprema-driver-card-sub">${escapeHtml(otpremaState.selectedVozac.id)}</div>
    `);

    const sections = buildOtpremaAssignSections();

    setHtml(sectionsEl, `
        ${renderOtpremaAssignSection('Današnji bez vozača', sections.todayUnassigned)}
        ${renderOtpremaAssignSection('Raniji bez vozača', sections.olderUnassigned, true)}
    `);

    updateOtpremaAssignSummary();
}

function buildOtpremaAssignSections() {
    const today = getTodayIsoDate();

    return {
        todayUnassigned: otpremaState.rows.filter(r => !r.vozacID && r.datum === today),
        olderUnassigned: otpremaState.rows.filter(r => !r.vozacID && r.datum !== today)
    };
}

function renderOtpremaAssignSection(title, items, showWarning) {
    if (!items.length) {
        return `
            <section class="otprema-section">
                <div class="otprema-section-head">
                    <div class="otprema-section-title">${escapeHtml(title)}</div>
                    <div class="otprema-section-count">0</div>
                </div>
                <div class="otprema-empty">Nema stavki</div>
            </section>
        `;
    }

    return `
        <section class="otprema-section">
            <div class="otprema-section-head">
                <div class="otprema-section-title">${escapeHtml(title)}</div>
                <div class="otprema-section-count">${items.length}</div>
            </div>

            <div class="otprema-check-cards">
                ${items.map(row => renderOtpremaAssignCard(row, showWarning)).join('')}
            </div>
        </section>
    `;
}

function renderOtpremaAssignCard(row, showWarning) {
    const key = getOtpremaRecordKey(row);
    const checked = otpremaState.selectedKeys.has(key) ? 'checked' : '';
    const note = row.napomena ? `<div class="otprema-card-note">${escapeHtml(row.napomena)}</div>` : '';
    const warning = showWarning
        ? `<span class="otprema-badge otprema-badge--warning">Raniji otkup</span>`
        : '';

    return `
        <label class="otprema-check-card">
            <div class="otprema-check-col">
                <input class="otprema-check" type="checkbox" data-record-key="${escapeHtml(key)}" ${checked}>
            </div>

            <div class="otprema-check-body">
                <div class="otprema-card-top">
                    <div class="otprema-card-koop">${escapeHtml(row.kooperantName || '-')}</div>
                    <div class="otprema-card-date">${escapeHtml(row.datum || '-')}</div>
                </div>

                <div class="otprema-card-main">
                    <div class="otprema-card-line">
                        ${escapeHtml(row.vrstaVoca || '-')}
                        ${row.sortaVoca ? ' / ' + escapeHtml(row.sortaVoca) : ''}
                        <span class="otprema-card-class">Kl. ${escapeHtml(row.klasa || 'I')}</span>
                    </div>

                    <div class="otprema-card-line otprema-card-line--muted">
                        ${escapeHtml(formatOtpremaKg(row.kolicina))} • ${escapeHtml(formatOtpremaAmbalaza(row))}
                    </div>

                    ${note}
                </div>

                <div class="otprema-card-bottom">
                    <div class="otprema-card-badges">${warning}</div>
                </div>
            </div>
        </label>
    `;
}

function selectAllOtpremaToday() {
    const todayRows = otpremaState.rows.filter(r => !r.vozacID && r.datum === getTodayIsoDate());
    otpremaState.selectedKeys = new Set(todayRows.map(getOtpremaRecordKey));
    renderOtpremaAssignView();
}

function clearOtpremaSelection() {
    otpremaState.selectedKeys.clear();
    renderOtpremaAssignView();
}

function updateOtpremaAssignSummary() {
    const selectedRows = getSelectedOtpremaRows();
    setText(byId('otpremaSelectedCount'), String(selectedRows.length));
    setText(byId('otpremaSelectedKg'), formatOtpremaKg(sumOtpremaKg(selectedRows)));
}

async function confirmOtpremaAssign() {
    if (!otpremaState.selectedVozac) {
        showToast('Prvo izaberi vozača', 'error');
        return;
    }

    const selectedRows = getSelectedOtpremaRows();
    if (!selectedRows.length) {
        showToast('Izaberi najmanje jednu stavku', 'error');
        return;
    }

    const nowIso = new Date().toISOString();
    const updatedRows = selectedRows.map(row =>
        buildUpdatedOtpremaRecord(row, otpremaState.selectedVozac, nowIso)
    );

    try {
        for (const row of updatedRows) {
            await dbPut(db, CONFIG.STORE_NAME, row);
        }

        otpremaState.successRows = updatedRows;
        renderOtpremaSuccessView(updatedRows, otpremaState.selectedVozac);
        showOtpremaSuccessView();

        if (typeof updateSyncBadge === 'function') updateSyncBadge();

        if (navigator.onLine && typeof syncQueueSafe === 'function') {
            syncQueueSafe();
        }

        // osveži lokalni state posle success prikaza
        otpremaState.rows = otpremaState.rows.map(row => {
            const match = updatedRows.find(u => getOtpremaRecordKey(u) === getOtpremaRecordKey(row));
            return match || row;
        });

    } catch (err) {
        console.error('confirmOtpremaAssign failed:', err);
        showToast('Greška pri potvrdi otpreme', 'error');
    }
}

function buildUpdatedOtpremaRecord(row, vozac, nowIso) {
    if (!row.clientRecordID) {
        throw new Error('Otprema zahteva postojeći clientRecordID');
    }

    return {
        clientRecordID: row.clientRecordID,
        serverRecordID: row.serverRecordID || '',
        createdAtClient: row.createdAtClient || nowIso,
        updatedAtClient: nowIso,
        updatedAtServer: row.updatedAtServer || '',
        syncedAt: '',
        deviceID: typeof getDeviceID === 'function' ? getDeviceID() : '',

        otkupacID: row.otkupacID || CONFIG.OTKUPAC_ID,
        datum: row.datum || getTodayIsoDate(),

        kooperantID: row.kooperantID || '',
        kooperantName: row.kooperantName || '',
        vrstaVoca: row.vrstaVoca || '',
        sortaVoca: row.sortaVoca || '',
        klasa: row.klasa || 'I',
        kolicina: parseFloat(row.kolicina) || 0,
        cena: parseFloat(row.cena) || 0,

        tipAmbalaze: row.tipAmbalaze || '',
        kolAmbalaze: parseInt(row.kolAmbalaze, 10) || 0,

        parcelaID: row.parcelaID || '',
        napomena: row.napomena || '',
        vozacID: vozac.id,
        vozacName: vozac.name,

        syncStatus: 'pending',
        syncAttempts: 0,
        syncAttemptAt: '',
        lastSyncError: '',
        lastServerStatus: '',
        deleted: !!row.deleted,
        entityType: row.entityType || 'otkup',
        schemaVersion: row.schemaVersion || 1
    };
}

function renderOtpremaSuccessView(rows, vozac) {
    setText(byId('otpremaSuccessDriver'), vozac.name + ' (' + vozac.id + ')');
    setText(byId('otpremaSuccessCount'), String(rows.length));
    setText(byId('otpremaSuccessKg'), formatOtpremaKg(sumOtpremaKg(rows)));

    const listEl = byId('otpremaSuccessList');
    if (!listEl) return;

    setHtml(listEl, `
        <section class="otprema-section">
            <div class="otprema-section-head">
                <div class="otprema-section-title">Dodeljene stavke</div>
                <div class="otprema-section-count">${rows.length}</div>
            </div>
            <div class="otprema-cards">
                ${rows.map(row => renderOtpremaCard(row, false, true)).join('')}
            </div>
        </section>
    `);
}

function openOtpremaDetail(recordKey) {
    const row = otpremaState.rows.find(r => getOtpremaRecordKey(r) === recordKey);
    if (!row) return;

    const modal = byId('otpremaDetailModal');
    const body = byId('otpremaDetailBody');
    const title = byId('otpremaDetailTitle');

    if (!modal || !body || !title) return;

    setText(title, row.kooperantName || row.kooperantID || 'Detalj otpreme');

    const badges = [];
    if (row.vozacID) badges.push('Dodeljen: ' + resolveVozacName(row.vozacID));
    if (row.syncStatus && row.syncStatus !== 'synced') badges.push('Čeka sync');
    if (row.lastSyncError) badges.push('Sync greška');

    setHtml(body, `
        <div class="otprema-detail-grid">
            <div><strong>Datum:</strong> ${escapeHtml(row.datum || '-')}</div>
            <div><strong>Kooperant:</strong> ${escapeHtml(row.kooperantName || row.kooperantID || '-')}</div>
            <div><strong>Roba:</strong> ${escapeHtml(row.vrstaVoca || '-')} ${row.sortaVoca ? '/ ' + escapeHtml(row.sortaVoca) : ''}</div>
            <div><strong>Klasa:</strong> ${escapeHtml(row.klasa || 'I')}</div>
            <div><strong>Količina:</strong> ${escapeHtml(formatOtpremaKg(row.kolicina))}</div>
            <div><strong>Ambalaža:</strong> ${escapeHtml(formatOtpremaAmbalaza(row))}</div>
            <div><strong>Vozač:</strong> ${escapeHtml(resolveVozacName(row.vozacID) || 'Nije dodeljen')}</div>
            ${row.napomena ? `<div><strong>Napomena:</strong> ${escapeHtml(row.napomena)}</div>` : ''}
            ${badges.length ? `<div><strong>Status:</strong> ${escapeHtml(badges.join(' • '))}</div>` : ''}
        </div>
    `);

    addClass(modal, 'visible');
}

function closeOtpremaDetail() {
    const modal = byId('otpremaDetailModal');
    if (modal) removeClass(modal, 'visible');
}

function backToOtpremaRoot() {
    otpremaState.selectedKeys.clear();
    otpremaState.selectedVozac = null;
    otpremaState.successRows = [];
    showOtpremaRootView();
    renderOtpremaRoot();
}

function mapServerOtpremaRecord(r) {
    return {
        clientRecordID: r.ClientRecordID || '',
        serverRecordID: r.ServerRecordID || '',
        createdAtClient: normalizeIso(r.CreatedAtClient),
        updatedAtClient: normalizeIso(r.UpdatedAtClient || r.CreatedAtClient),
        updatedAtServer: normalizeIso(r.UpdatedAtServer || r.ReceivedAt),
        syncedAt: normalizeIso(r.UpdatedAtServer || r.ReceivedAt),

        datum: toIsoDateOnly(r.Datum),
        kooperantID: r.KooperantID || '',
        kooperantName: r.KooperantName || r.KooperantID || '',
        vrstaVoca: r.VrstaVoca || '',
        sortaVoca: r.SortaVoca || '',
        klasa: r.Klasa || 'I',
        kolicina: parseFloat(r.Kolicina) || 0,
        cena: parseFloat(r.Cena) || 0,
        tipAmbalaze: r.TipAmbalaze || '',
        kolAmbalaze: parseInt(r.KolAmbalaze, 10) || 0,
        parcelaID: r.ParcelaID || '',
        napomena: r.Napomena || '',
        vozacID: r.VozacID || r.VozaciID || '',
        vozacName: r.VozacName || '',

        syncStatus: 'synced',
        lastSyncError: '',
        deleted: false
    };
}

function normalizeLocalOtpremaRecord(r) {
    return {
        clientRecordID: r.clientRecordID || '',
        serverRecordID: r.serverRecordID || '',
        createdAtClient: normalizeIso(r.createdAtClient),
        updatedAtClient: normalizeIso(r.updatedAtClient || r.createdAtClient),
        updatedAtServer: normalizeIso(r.updatedAtServer),
        syncedAt: normalizeIso(r.syncedAt),

        datum: toIsoDateOnly(r.datum || ''),
        kooperantID: r.kooperantID || '',
        kooperantName: r.kooperantName || r.kooperantID || '',
        vrstaVoca: r.vrstaVoca || '',
        sortaVoca: r.sortaVoca || '',
        klasa: r.klasa || 'I',
        kolicina: parseFloat(r.kolicina) || 0,
        cena: parseFloat(r.cena) || 0,
        tipAmbalaze: r.tipAmbalaze || '',
        kolAmbalaze: parseInt(r.kolAmbalaze, 10) || 0,
        parcelaID: r.parcelaID || '',
        napomena: r.napomena || '',
        vozacID: r.vozacID || '',
        vozacName: r.vozacName || '',

        syncStatus: r.syncStatus || 'pending',
        lastSyncError: r.lastSyncError || '',
        deleted: !!r.deleted
    };
}

function mergeOtpremaRecords(localRows, serverRows) {
    const map = new Map();

    serverRows.forEach(row => {
        const key = getOtpremaRecordKey(row);
        if (!key) return;
        map.set(key, row);
    });

    localRows
        .map(normalizeLocalOtpremaRecord)
        .forEach(row => {
            const key = getOtpremaRecordKey(row);
            if (!key) return;

            const existing = map.get(key);

            if (!existing) {
                map.set(key, row);
                return;
            }

            if (row.syncStatus !== 'synced' || row.lastSyncError) {
                map.set(key, row);
                return;
            }

            if ((row.updatedAtClient || '') > (existing.updatedAtClient || '')) {
                map.set(key, row);
            }
        });

    return Array.from(map.values()).filter(r => !r.deleted);
}

function enrichOtpremaRecord(r) {
    return {
        ...r,
        datum: toIsoDateOnly(r.datum || ''),
        vozacName: resolveVozacName(r.vozacID) || r.vozacName || ''
    };
}

function compareOtpremaRowsDesc(a, b) {
    const byDate = String(b.datum || '').localeCompare(String(a.datum || ''));
    if (byDate !== 0) return byDate;

    const aTime = a.updatedAtClient || a.createdAtClient || a.updatedAtServer || '';
    const bTime = b.updatedAtClient || b.createdAtClient || b.updatedAtServer || '';
    return String(bTime).localeCompare(String(aTime));
}

function getOtpremaRecordKey(r) {
    if (r.serverRecordID) return 'srv:' + r.serverRecordID;
    if (r.clientRecordID) return 'cli:' + r.clientRecordID;
    return '';
}

function resolveVozacName(vozacID) {
    if (!vozacID) return '';

    const vozac = (stammdaten.vozaci || []).find(v => (v.VozacID || v.ID) === vozacID);
    return vozac ? (vozac.ImePrezime || vozac.Naziv || vozac.Ime || vozacID) : vozacID;
}

function getOtpremaAssignedDate(row) {
    return toIsoDateOnly(
        row.updatedAtClient ||
        row.updatedAtServer ||
        row.syncedAt ||
        ''
    );
}

function sumOtpremaKg(rows) {
    return rows.reduce((sum, r) => sum + (parseFloat(r.kolicina) || 0), 0);
}

function formatOtpremaKg(value) {
    return (parseFloat(value) || 0).toLocaleString('sr-RS') + ' kg';
}

function formatOtpremaAmbalaza(row) {
    const kom = parseInt(row.kolAmbalaze, 10) || 0;
    const tip = row.tipAmbalaze || '';
    if (!kom && !tip) return 'Bez ambalaže';
    if (!kom) return tip;
    if (!tip) return kom + ' kom';
    return kom + ' kom • ' + tip;
}

function toIsoDateOnly(input) {
    if (!input) return '';

    const s = String(input).trim();

    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

    if (/^\d{2}\.\d{2}\.\d{4}\.?$/.test(s)) {
        const clean = s.replace(/\.$/, '');
        const parts = clean.split('.');
        return parts[2] + '-' + parts[1] + '-' + parts[0];
    }

    try {
        const d = new Date(s);
        if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
    } catch (_) {}

    return s;
}

function getTodayIsoDate() {
    return new Date().toISOString().slice(0, 10);
}

function getSelectedOtpremaRows() {
    return otpremaState.rows.filter(r =>
        otpremaState.selectedKeys.has(getOtpremaRecordKey(r)) &&
        !r.vozacID
    );
}
