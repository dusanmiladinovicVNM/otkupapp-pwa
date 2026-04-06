// ============================================================
// OTPREMA (dispatch)
// ============================================================
let otpremaVozacID = '';
let otpremaUnassigned = [];

async function getOtpremaOtkupiCached(forceRefresh) {
    if (
        !forceRefresh &&
        _otpremaServerCache.data &&
        (Date.now() - _otpremaServerCache.ts < OTPREMA_CACHE_TTL)
    ) {
        return _otpremaServerCache.data;
    }

    let local = [];
    let server = [];

    try {
        local = await dbGetAll(db, CONFIG.STORE_NAME);
    } catch (err) {
        console.error('getOtpremaOtkupiCached local failed:', err);
    }

    if (navigator.onLine) {
        try {
            const json = await apiFetch('action=getOtkupi&otkupacID=' + encodeURIComponent(CONFIG.OTKUPAC_ID));
            if (json && json.success && Array.isArray(json.records)) {
                server = json.records.map(mapServerOtpremaRecord);
            }
        } catch (e) {
            console.error('getOtpremaOtkupiCached server failed:', e);
        }
    }

    const all = mergeOtpremaRecords(local, server);

    _otpremaServerCache = { data: all, ts: Date.now() };
    return all;
}

function invalidateOtpremaCache() {
    _otpremaServerCache = { data: null, ts: 0 };
}

function startOtpremaVozacScan() {
    const readerDiv = document.getElementById('qr-reader-otprema');
    readerDiv.style.display = 'block';
    const scanner = new Html5Qrcode('qr-reader-otprema');
    scanner.start(
        { facingMode: 'environment' },
        { fps: 10, qrbox: { width: 250, height: 250 } },
        (decodedText) => {
            scanner.stop().then(() => { readerDiv.style.display = 'none'; });
            onOtpremaVozacScanned(decodedText);
        },
        () => {}
    ).catch(err => { showToast('Kamera nije dostupna', 'error'); readerDiv.style.display = 'none'; });
}

function onOtpremaVozacScanned(text) {
    let vozID = '', vozName = '';
    try {
        const data = JSON.parse(text);
        if (data.type === 'VOZ' && data.id) { vozID = data.id; vozName = data.name || data.id; }
    } catch (e) {}
    if (!vozID && text.startsWith('VOZ-')) { vozID = text; vozName = text; }
    if (!vozID) { showToast('Nije QR vozača', 'error'); return; }

    otpremaVozacID = vozID;
    document.getElementById('otpremaVozacName').textContent = vozName;
    document.getElementById('otpremaVozacId').textContent = vozID;
    showOtpremaAssignView();
}

async function showOtpremaAssignView() {
    const mainView = document.getElementById('otpremaMainView');
    const assignView = document.getElementById('otpremaAssignView');

    if (mainView) mainView.style.display = 'none';
    if (assignView) assignView.style.display = 'block';

    const all = await getOtpremaOtkupiCached(true);

    otpremaUnassigned = all.filter(r => !r.vozacID);
    otpremaUnassigned.sort((a, b) => {
        const byDate = (b.datum || '').localeCompare(a.datum || '');
        if (byDate !== 0) return byDate;

        const byTime = String(b.updatedAtClient || b.createdAtClient || '').localeCompare(
            String(a.updatedAtClient || a.createdAtClient || '')
        );
        if (byTime !== 0) return byTime;

        return String(b.clientRecordID || '').localeCompare(String(a.clientRecordID || ''));
    });

    renderOtpremaCheckboxes();
}

function renderOtpremaCheckboxes() {
    const list = document.getElementById('otpremaOtkupList');
    if (!list) return;

    if (otpremaUnassigned.length === 0) {
        list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema neraspoređenih otkupa za danas</p>';
        return;
    }

    list.innerHTML = otpremaUnassigned.map((r, i) => {
        const vr = ((r.kolicina || 0) * (r.cena || 0)).toLocaleString('sr');
        return `<div class="queue-item" style="cursor:pointer;" onclick="toggleOtpremaItem(${i})">
            <div style="display:flex;align-items:center;gap:10px;">
                <input type="checkbox" id="otpChk${i}" style="width:20px;height:20px;flex-shrink:0;" onclick="event.stopPropagation();updateOtpremaSummary();">
                <div style="flex:1;">
                    <div class="qi-header"><span class="qi-koop">${escapeHtml(r.kooperantName)}</span><span class="qi-time">${escapeHtml(r.datum)}</span></div>
                    <div class="qi-detail" style="font-size:11px;color:var(--text-muted);">${escapeHtml(r.klasa || 'I')}</div>
                    <div class="qi-detail">${escapeHtml(r.vrstaVoca)} ${escapeHtml(r.sortaVoca || '')} | ${r.kolicina} kg × ${r.cena} = ${vr} RSD</div>
                </div>
            </div>
        </div>`;
    }).join('');

    updateOtpremaSummary();
}
function toggleOtpremaItem(index) {
    const chk = document.getElementById('otpChk' + index);
    chk.checked = !chk.checked;
    updateOtpremaSummary();
}

function toggleSelectAll() {
    const checkboxes = document.querySelectorAll('[id^="otpChk"]');
    const allChecked = Array.from(checkboxes).every(c => c.checked);
    checkboxes.forEach(c => c.checked = !allChecked);
    updateOtpremaSummary();
}

function updateOtpremaSummary() {
    const div = document.getElementById('otpremaSummary');
    const text = document.getElementById('otpremaSummaryText');
    if (!div || !text) return;

    let kg = 0;
    let count = 0;

    otpremaUnassigned.forEach((r, i) => {
        const chk = document.getElementById('otpChk' + i);
        if (chk && chk.checked) {
            kg += r.kolicina || 0;
            count++;
        }
    });

    if (count > 0) {
        div.style.display = 'block';
        text.textContent = 'Izabrano: ' + count + ' otkupa | ' + kg.toLocaleString('sr') + ' kg';
    } else {
        div.style.display = 'none';
    }
}

async function confirmOtprema() {
    if (!otpremaVozacID) {
        showToast('Nema vozača', 'error');
        return;
    }

    const selected = [];

    for (let i = 0; i < otpremaUnassigned.length; i++) {
        const chk = document.getElementById('otpChk' + i);
        if (chk && chk.checked) {
            selected.push(otpremaUnassigned[i]);
        }
    }

    if (selected.length === 0) {
        showToast('Izaberite bar jedan otkup', 'error');
        return;
    }

    try {
        const nowIso = new Date().toISOString();

        for (const item of selected) {
            item.vozacID = otpremaVozacID;
            item.updatedAtClient = nowIso;
            item.syncStatus = 'pending';
            item.lastSyncError = '';
            item.lastServerStatus = '';
            item.syncAttemptAt = '';
            await dbPut(db, CONFIG.STORE_NAME, item);
        }

        showToast(selected.length + ' otkupa dodeljeno vozaču', 'success');

        invalidateOtpremaCache();
        cancelOtprema();

        if (navigator.onLine) {
            if (typeof syncQueueSafe === 'function') {
                await syncQueueSafe();
            } else if (typeof syncQueue === 'function') {
                await syncQueue();
            }
        }
    } catch (err) {
        console.error('confirmOtprema failed:', err);
        showToast('Greška pri dodeli otkupa vozaču', 'error');
    }
}

async function cancelOtprema() {
    otpremaVozacID = '';
    otpremaUnassigned = [];

    const assignView = document.getElementById('otpremaAssignView');
    const mainView = document.getElementById('otpremaMainView');

    if (assignView) assignView.style.display = 'none';
    if (mainView) mainView.style.display = 'block';

    await loadOtpremaOverview();
}

async function loadOtpremaOverview() {
    const all = await getOtpremaOtkupiCached(false);

    const unassigned = all.filter(r => !r.vozacID);
    unassigned.sort((a, b) => (b.datum || '').localeCompare(a.datum || ''));

    const assigned = all.filter(r => r.vozacID);

    const uList = document.getElementById('otpremaUnassignedList');
    if (uList) {
        if (unassigned.length === 0) {
            uList.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:12px;font-size:13px;">Svi otkupi su raspoređeni</p>';
        } else {
            uList.innerHTML = unassigned.map(r => {
                const vr = ((r.kolicina || 0) * (r.cena || 0)).toLocaleString('sr-RS');
                const statusText =
                    r.syncStatus === 'syncing' ? ' | sync...' :
                    r.syncStatus === 'pending' ? ' | pending' : '';

                return `<div class="queue-item" style="border-left-color:var(--warning);">
                    <div class="qi-header"><span class="qi-koop">${escapeHtml(r.kooperantName)}</span><span class="qi-time">${escapeHtml(r.datum)}</span></div>
                    <div class="qi-detail">${escapeHtml(r.vrstaVoca)} ${escapeHtml(r.sortaVoca || '')} ${escapeHtml(r.klasa || 'I')} | ${r.kolicina} kg | ${vr} RSD${statusText}</div>
                </div>`;
            }).join('');
        }
    }

    const aList = document.getElementById('otpremaAssignedList');
    if (aList) {
        if (assigned.length === 0) {
            aList.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:12px;font-size:13px;">Nema otprema za danas</p>';
        } else {
            const grouped = {};

            assigned.forEach(r => {
                const v = r.vozacID;
                if (!grouped[v]) grouped[v] = { items: [], kg: 0 };
                grouped[v].items.push(r);
                grouped[v].kg += r.kolicina || 0;
            });

            aList.innerHTML = Object.entries(grouped).map(([vozID, g]) =>
                `<div style="background:white;border-radius:var(--radius);padding:14px;margin-bottom:12px;box-shadow:0 1px 3px rgba(0,0,0,0.08);border-left:4px solid var(--success);">
                    <div style="display:flex;justify-content:space-between;margin-bottom:6px;">
                        <strong style="color:var(--primary);">🚛 ${escapeHtml(vozID)}</strong>
                        <span style="font-weight:600;">${g.kg.toLocaleString('sr-RS')} kg | ${g.items.length} otk.</span>
                    </div>
                    ${g.items.map(r => {
                        const st =
                            r.syncStatus === 'syncing' ? ' (sync...)' :
                            r.syncStatus === 'pending' ? ' (pending)' : '';
                        return `<div style="padding:3px 0;font-size:12px;border-top:1px solid #eee;">${escapeHtml(r.kooperantName)} | ${escapeHtml(r.vrstaVoca)} ${escapeHtml(r.klasa || '')} | ${r.kolicina} kg${escapeHtml(st)}</div>`;
                    }).join('')}
                </div>`
            ).join('');
        }
    }
}


//HELPERS
function mapServerOtpremaRecord(r) {
    return {
        clientRecordID: r.ClientRecordID || '',
        serverRecordID: r.ServerRecordID || '',
        createdAtClient: normalizeIso(r.CreatedAtClient),
        updatedAtClient: normalizeIso(r.UpdatedAtClient || r.CreatedAtClient),
        updatedAtServer: normalizeIso(r.UpdatedAtServer || r.ReceivedAt),
        syncedAt: normalizeIso(r.UpdatedAtServer || r.ReceivedAt),

        datum: fmtDate(r.Datum),
        kooperantID: r.KooperantID || '',
        kooperantName: r.KooperantName || r.KooperantID || '',
        vrstaVoca: r.VrstaVoca || '',
        sortaVoca: r.SortaVoca || '',
        klasa: r.Klasa || 'I',
        kolicina: parseFloat(r.Kolicina) || 0,
        cena: parseFloat(r.Cena) || 0,
        kolAmbalaze: parseInt(r.KolAmbalaze, 10) || 0,
        parcelaID: r.ParcelaID || '',
        vozacID: r.VozacID || r.VozaciID || '',
        napomena: r.Napomena || '',

        syncStatus: 'synced',
        syncAttempts: 0,
        lastSyncError: '',
        lastServerStatus: 'server'
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

        datum: r.datum || '',
        kooperantID: r.kooperantID || '',
        kooperantName: r.kooperantName || r.kooperantID || '',
        vrstaVoca: r.vrstaVoca || '',
        sortaVoca: r.sortaVoca || '',
        klasa: r.klasa || 'I',
        kolicina: parseFloat(r.kolicina) || 0,
        cena: parseFloat(r.cena) || 0,
        kolAmbalaze: parseInt(r.kolAmbalaze, 10) || 0,
        parcelaID: r.parcelaID || '',
        vozacID: r.vozacID || '',
        napomena: r.napomena || '',

        syncStatus: r.syncStatus || 'pending',
        syncAttempts: parseInt(r.syncAttempts, 10) || 0,
        lastSyncError: r.lastSyncError || '',
        lastServerStatus: r.lastServerStatus || ''
    };
}

function mergeOtpremaRecords(local, server) {
    return mergeOfflineRecords(local, server, normalizeLocalOtpremaRecord);
}

