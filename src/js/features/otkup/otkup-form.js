const OTKUP_TIP_AMBALAZE_OPTIONS = ['12/1', '6/1', '2/1'];

function initOtkupFormUI(options) {
    const opts = options || {};
    const preserveSelection = !!opts.preserveSelection;

    const prev = preserveSelection ? {
        kooperantID: getFieldValue('fldKooperantID'),
        vrsta: getFieldValue('fldVrsta'),
        sorta: getFieldValue('fldSorta'),
        klasa: getFieldValue('fldKlasa'),
        tipAmbalaze: getFieldValue('fldTipAmbalaze'),
        parcelaID: getFieldValue('fldParcela')
    } : null;

    populateVrstaDropdown();
    populateTipAmbalazeDropdown(prev ? prev.tipAmbalaze : '');
    applyDefaults();

    if (preserveSelection && prev) {
        if (prev.vrsta) {
            setFieldValue('fldVrsta', prev.vrsta);
            onVrstaChange();
        }
        if (prev.sorta) setFieldValue('fldSorta', prev.sorta);
        if (prev.klasa) setFieldValue('fldKlasa', prev.klasa);
        if (prev.tipAmbalaze) setFieldValue('fldTipAmbalaze', prev.tipAmbalaze);

        if (prev.kooperantID) {
            populateParcelaDropdown(prev.kooperantID);
            if (prev.parcelaID) setFieldValue('fldParcela', prev.parcelaID);
        }
    }

    bindOtkupFormUIEvents();
    clearOtkupValidation();
    updateTipAmbalazeHint();
}

// UX flow na mobile:
// 1 -> 2 nakon kooperanta
// 2 -> 3 nakon parcele
// 3 -> 4 nakon količine
// blok 4 se smatra završenim tek nakon ambalaže, bez auto-scroll dalje
function bindOtkupFormUIEvents() {
    const root = document.getElementById('tab-otkup');
    if (!root || root.dataset.otkupUiBound === '1') return;

    const fldVrsta = document.getElementById('fldVrsta');
    const fldTipAmbalaze = document.getElementById('fldTipAmbalaze');
    const fldKooperantManual = document.getElementById('fldKooperantManual');
    const fldKolicina = document.getElementById('fldKolicina');
    const fldCena = document.getElementById('fldCena');
    const fldAmbalaza = document.getElementById('fldAmbalaza');
    const fldParcela = document.getElementById('fldParcela');
    
    if (fldVrsta) {
        fldVrsta.addEventListener('change', () => {
            clearOtkupError('errVrsta', 'fldVrsta');
            clearOtkupError('errTipAmbalaze', 'fldTipAmbalaze');
            updateTipAmbalazeHint();
        });
    }

    if (fldTipAmbalaze) {
        fldTipAmbalaze.addEventListener('change', () => {
            clearOtkupError('errTipAmbalaze', 'fldTipAmbalaze');
            updateTipAmbalazeHint();
        });
    }

    if (fldKooperantManual) {
        fldKooperantManual.addEventListener('change', () => {
            clearOtkupError('errKooperant', 'fldKooperantManual');
        });
    }
    
    if (fldParcela) {
        fldParcela.addEventListener('change', () => {
            scrollToOtkupStep('otkupStep3Roba');
        });
    }

    if (fldKolicina) {
        fldKolicina.addEventListener('input', () => {
            clearOtkupError('errKolicina', 'fldKolicina');
        });

        fldKolicina.addEventListener('change', () => {
            if (!fldKolicina.value) return;
            scrollToOtkupStep('otkupStep4CenaAmbalaza');
        });

        fldKolicina.addEventListener('blur', () => {
            if (!fldKolicina.value) return;
            scrollToOtkupStep('otkupStep4CenaAmbalaza');
        });
    }

    if (fldCena) {
        fldCena.addEventListener('input', () => {
            clearOtkupError('errCena', 'fldCena');
        });
    }

    if (fldAmbalaza) {
        fldAmbalaza.addEventListener('input', () => {
            clearOtkupError('errAmbalaza', 'fldAmbalaza');
        });

        fldAmbalaza.addEventListener('change', () => {
            if (!fldAmbalaza.value) return;

            // blok 4 je tek sada kompletan
            // za sada nema auto-scroll ka Napredno
        });

        fldAmbalaza.addEventListener('blur', () => {
            if (!fldAmbalaza.value) return;

            // blok 4 je tek sada kompletan
            // za sada nema auto-scroll ka Napredno
        });
    }

    root.dataset.otkupUiBound = '1';
}

function populateVrstaDropdown() {
    const sel = document.getElementById('fldVrsta');
    if (!sel) return;

    sel.innerHTML = '<option value="">-- Izaberi --</option>';

    const vrsteSet = new Set();
    (stammdaten.kulture || []).forEach(k => {
        if (!k || !k.VrstaVoca || vrsteSet.has(k.VrstaVoca)) return;

        vrsteSet.add(k.VrstaVoca);

        const opt = document.createElement('option');
        opt.value = k.VrstaVoca;
        opt.textContent = k.VrstaVoca;
        sel.appendChild(opt);
    });

    populateKooperantDropdown();
}

function populateKooperantDropdown() {
    const sel = document.getElementById('fldKooperantManual');
    if (!sel) return;

    sel.innerHTML = '<option value="">-- Izaberi --</option>';

    (stammdaten.kooperanti || [])
        .filter(k => k && k.StanicaID === CONFIG.OTKUPAC_ID)
        .forEach(k => {
            const opt = document.createElement('option');
            opt.value = k.KooperantID;
            opt.textContent = `${k.Ime || ''} ${k.Prezime || ''} (${k.KooperantID})`.trim();
            sel.appendChild(opt);
        });
}

function onManualKooperantChange() {
    const fld = document.getElementById('fldKooperantManual');
    if (!fld) return;

    const koopID = fld.value;
    if (!koopID) return;

    const koop = (stammdaten.kooperanti || []).find(k => k.KooperantID === koopID);
    setKooperant(koopID, koop ? `${koop.Ime || ''} ${koop.Prezime || ''}`.trim() : koopID);
    clearOtkupError('errKooperant', 'fldKooperantManual');
    scrollToOtkupStep('otkupStep2ParcelaVozac');
}

function populateParcelaDropdown(kooperantID) {
    const sel = document.getElementById('fldParcela');
    const group = document.getElementById('parcelaGroup');

    if (!sel || !group) return;

    sel.innerHTML = '<option value="">-- Bez parcele --</option>';

    const parcele = (stammdaten.parcele || []).filter(p => p && p.KooperantID === kooperantID);

    if (parcele.length === 0) {
        group.style.display = 'none';
        return;
    }

    group.style.display = 'block';

    parcele.forEach(p => {
        const opt = document.createElement('option');
        opt.value = p.ParcelaID;
        opt.textContent = `${p.KatBroj || ''} - ${p.Kultura || ''} (${p.ParcelaID})`;
        sel.appendChild(opt);
    });
}

function populateTipAmbalazeDropdown(selectedValue) {
    const sel = document.getElementById('fldTipAmbalaze');
    if (!sel) return;

    const current = selectedValue || sel.value || '';
    sel.innerHTML = '<option value="">-- Izaberi --</option>';

    OTKUP_TIP_AMBALAZE_OPTIONS.forEach(val => {
        const opt = document.createElement('option');
        opt.value = val;
        opt.textContent = val;
        sel.appendChild(opt);
    });

    if (current && OTKUP_TIP_AMBALAZE_OPTIONS.includes(current)) {
        sel.value = current;
    }
}

function normalizeVrsta(vrsta) {
    return String(vrsta || '')
        .trim()
        .toLowerCase()
        .replaceAll('š', 's')
        .replaceAll('đ', 'dj')
        .replaceAll('č', 'c')
        .replaceAll('ć', 'c')
        .replaceAll('ž', 'z');
}

function getDefaultTipAmbalazeForVrsta(vrsta) {
    const v = normalizeVrsta(vrsta);

    if (v === 'visnja' || v === 'sljiva') {
        return '12/1';
    }

    return '6/1';
}

function updateTipAmbalazeHint() {
    const hint = document.getElementById('tipAmbalazeHint');
    const fldVrsta = document.getElementById('fldVrsta');
    const fldTip = document.getElementById('fldTipAmbalaze');
    if (!hint || !fldVrsta || !fldTip) return;

    const vrsta = fldVrsta.value || '';
    const selected = fldTip.value || '';

    if (!vrsta) {
        hint.textContent = 'Izaberi vrstu voća da se predloži podrazumevani tip ambalaže.';
        return;
    }

    const def = getDefaultTipAmbalazeForVrsta(vrsta);
    hint.textContent = `Podrazumevano za ${vrsta}: ${def}. Trenutno izabrano: ${selected || 'nije izabrano'}.`;
}

function syncTipAmbalazeWithVrsta(force) {
    const fldVrsta = document.getElementById('fldVrsta');
    const fldTip = document.getElementById('fldTipAmbalaze');
    if (!fldVrsta || !fldTip) return;

    const vrsta = fldVrsta.value || '';
    if (!vrsta) {
        if (force) fldTip.value = '';
        updateTipAmbalazeHint();
        return;
    }

    const defaultTip = getDefaultTipAmbalazeForVrsta(vrsta);
    if (force || !fldTip.value) {
        fldTip.value = defaultTip;
    }

    updateTipAmbalazeHint();
}

function applyDefaults() {
    const config = stammdaten.config || [];

    populateTipAmbalazeDropdown();

    const dv = config.find(c => c.Parameter === 'DefaultVrsta');
    if (dv && dv.Vrednost) {
        const fldVrsta = document.getElementById('fldVrsta');
        if (fldVrsta) fldVrsta.value = dv.Vrednost;

        onVrstaChange();

        const ds = config.find(c => c.Parameter === 'DefaultSorta');
        if (ds && ds.Vrednost) {
            const fldSorta = document.getElementById('fldSorta');
            if (fldSorta) fldSorta.value = ds.Vrednost;
        }
    } else {
        syncTipAmbalazeWithVrsta(true);
    }

    applyDefaultCena();
}

function applyDefaultCena() {
    const fldVrsta = document.getElementById('fldVrsta');
    const fldCena = document.getElementById('fldCena');
    if (!fldVrsta || !fldCena) return;

    const vrsta = fldVrsta.value;
    if (!vrsta) return;

    const cc = (stammdaten.config || []).find(c => c.Parameter === 'Cena' + vrsta);
    if (cc && cc.Vrednost && (!fldCena.value || fldCena.value === '0')) {
        fldCena.value = cc.Vrednost;
    }
}

function onVrstaChange() {
    const fldVrsta = document.getElementById('fldVrsta');
    const sel = document.getElementById('fldSorta');
    if (!fldVrsta || !sel) return;

    const vrsta = fldVrsta.value;
    sel.innerHTML = '<option value="">-- Izaberi --</option>';

    const sorteSet = new Set();

    (stammdaten.kulture || [])
        .filter(k => k && k.VrstaVoca === vrsta && k.SortaVoca)
        .forEach(k => {
            if (sorteSet.has(k.SortaVoca)) return;
            sorteSet.add(k.SortaVoca);

            const opt = document.createElement('option');
            opt.value = k.SortaVoca;
            opt.textContent = k.SortaVoca;
            sel.appendChild(opt);
        });

    applyDefaultCena();
    syncTipAmbalazeWithVrsta(true);
    clearOtkupError('errVrsta', 'fldVrsta');
    clearOtkupError('errTipAmbalaze', 'fldTipAmbalaze');
}

// ============================================================
// SAVE OTKUP
// ============================================================

async function saveOtkup() {
    try {
        const input = readOtkupForm();
        const validationError = validateOtkupInput(input);

        if (validationError) {
            showToast(validationError, 'error');
            return;
        }

        const record = buildOtkupRecord(input);

        await dbPut(db, CONFIG.STORE_NAME, record);

        showToast('Otkup sačuvan! ' + escapeHtml(String(record.kolicina)) + ' kg', 'success');

        if (typeof showOtkupniList === 'function') {
            showOtkupniList(record);
        }

        resetForm();
        await safeRefreshAfterSave();

        if (navigator.onLine) {
            setTimeout(() => {
                if (typeof syncQueueSafe === 'function') {
                    syncQueueSafe();
                } else if (typeof syncQueue === 'function') {
                    syncQueue();
                }
            }, 0);
        }
    } catch (err) {
        console.error('saveOtkup failed:', err);
        showToast('Greška pri čuvanju otkupa', 'error');
    }
}

function readOtkupForm() {
    const kooperantID = getFieldValue('fldKooperantID');
    const kooperantName = getTextValue('koopName');
    const vrstaVoca = getFieldValue('fldVrsta');
    const sortaVoca = getFieldValue('fldSorta');
    const klasa = getFieldValue('fldKlasa') || 'I';
    const kolicina = parseFloat(getFieldValue('fldKolicina')) || 0;
    const cena = parseFloat(getFieldValue('fldCena')) || 0;
    const tipAmbalaze = getFieldValue('fldTipAmbalaze') || getDefaultTipAmbalazeForVrsta(vrstaVoca);
    const kolAmbalaze = parseInt(getFieldValue('fldAmbalaza'), 10) || 0;
    const parcelaID = getFieldValue('fldParcela') || '';
    const napomena = getFieldValue('fldNapomena') || '';
    const vozacID = getFieldValue('fldVozacID') || '';

    return {
        kooperantID,
        kooperantName,
        vrstaVoca,
        sortaVoca,
        klasa,
        kolicina,
        cena,
        tipAmbalaze,
        kolAmbalaze,
        parcelaID,
        napomena,
        vozacID
    };
}

function validateOtkupInput(input) {
    clearOtkupValidation();

    if (!input.kooperantID) {
        setOtkupError('errKooperant', 'fldKooperantManual', 'Skenirajte ili izaberite kooperanta');
        return 'Skenirajte ili izaberite kooperanta';
    }

    if (!input.vrstaVoca) {
        setOtkupError('errVrsta', 'fldVrsta', 'Izaberite vrstu voća');
        return 'Izaberite vrstu voća';
    }

    if (input.kolicina <= 0) {
        setOtkupError('errKolicina', 'fldKolicina', 'Unesite količinu');
        return 'Unesite količinu';
    }

    if (input.cena <= 0) {
        setOtkupError('errCena', 'fldCena', 'Unesite cenu');
        return 'Unesite cenu';
    }

    if (!input.tipAmbalaze) {
        setOtkupError('errTipAmbalaze', 'fldTipAmbalaze', 'Izaberite tip ambalaže');
        return 'Izaberite tip ambalaže';
    }

    if (input.kolAmbalaze <= 0) {
        setOtkupError('errAmbalaza', 'fldAmbalaza', 'Unesite broj komada ambalaže');
        return 'Unesite broj komada ambalaže';
    }

    return '';
}

function buildOtkupRecord(input) {
    const nowIso = new Date().toISOString();
    const today = nowIso.split('T')[0];

    return {
        clientRecordID: generateClientRecordID(),
        serverRecordID: '',
        createdAtClient: nowIso,
        updatedAtClient: nowIso,
        updatedAtServer: '',
        syncedAt: '',
        deviceID: safeGetDeviceID(),

        otkupacID: CONFIG.OTKUPAC_ID,
        datum: today,

        kooperantID: input.kooperantID,
        kooperantName: input.kooperantName || input.kooperantID,

        vrstaVoca: input.vrstaVoca,
        sortaVoca: input.sortaVoca || '',
        klasa: input.klasa || 'I',
        kolicina: input.kolicina,
        cena: input.cena,

        tipAmbalaze: input.tipAmbalaze,
        kolAmbalaze: input.kolAmbalaze,

        parcelaID: input.parcelaID || '',
        napomena: input.napomena || '',
        vozacID: input.vozacID || '',

        syncStatus: 'pending',
        syncAttempts: 0,
        syncAttemptAt: '',
        lastSyncError: '',
        lastServerStatus: '',
        deleted: false,
        entityType: 'otkup',
        schemaVersion: 1
    };
}

async function safeRefreshAfterSave() {
    try {
        if (typeof updateStats === 'function') {
            await updateStats();
        }
    } catch (e) {
        console.error('updateStats after save failed:', e);
    }

    try {
        if (typeof renderQueueList === 'function') {
            await renderQueueList();
        }
    } catch (e) {
        console.error('renderQueueList after save failed:', e);
    }

    try {
        if (typeof updateSyncBadge === 'function') {
            await updateSyncBadge();
        }
    } catch (e) {
        console.error('updateSyncBadge after save failed:', e);
    }
}

function resetForm() {
    clearOtkupValidation();

    setFieldValue('fldKooperantID', '');
    setFieldValue('fldKooperantManual', '');

    const koopDisplay = document.getElementById('koopDisplay');
    if (koopDisplay) koopDisplay.classList.remove('visible');

    const fldParcela = document.getElementById('fldParcela');
    if (fldParcela) {
        fldParcela.innerHTML = '<option value="">-- Bez parcele --</option>';
    }

    const parcelaGroup = document.getElementById('parcelaGroup');
    if (parcelaGroup) parcelaGroup.style.display = 'none';

    setFieldValue('fldVrsta', '');

    const fldSorta = document.getElementById('fldSorta');
    if (fldSorta) {
        fldSorta.innerHTML = '<option value="">-- Izaberi --</option>';
    }

    setFieldValue('fldKlasa', 'I');
    setFieldValue('fldKolicina', '');
    setFieldValue('fldCena', '');
    setFieldValue('fldTipAmbalaze', '');
    setFieldValue('fldAmbalaza', '');
    setFieldValue('fldNapomena', '');
    setFieldValue('fldVozacID', '');

    const vozacDisplay = document.getElementById('vozacDisplay');
    if (vozacDisplay) vozacDisplay.classList.remove('visible');

    populateTipAmbalazeDropdown();
    applyDefaults();
    updateTipAmbalazeHint();
}

// ============================================================
// VALIDATION HELPERS
// ============================================================

function clearOtkupValidation() {
    document.querySelectorAll('#tab-otkup .otk-field-error').forEach(el => {
        el.hidden = true;
        el.textContent = '';
    });

    document.querySelectorAll('#tab-otkup .is-invalid').forEach(el => {
        el.classList.remove('is-invalid');
    });
}

function setOtkupError(errorId, fieldId, message) {
    const err = document.getElementById(errorId);
    const field = document.getElementById(fieldId);

    if (err) {
        err.hidden = false;
        err.textContent = message;
    }

    if (field) {
        field.classList.add('is-invalid');
        try { field.focus(); } catch (_) {}
    }
}

function clearOtkupError(errorId, fieldId) {
    const err = document.getElementById(errorId);
    const field = document.getElementById(fieldId);

    if (err) {
        err.hidden = true;
        err.textContent = '';
    }

    if (field) {
        field.classList.remove('is-invalid');
    }
}

// ============================================================
// HELPERS
// ============================================================

function getFieldValue(id) {
    const el = document.getElementById(id);
    return el ? el.value : '';
}

function setFieldValue(id, value) {
    const el = document.getElementById(id);
    if (el) el.value = value;
}

function getTextValue(id) {
    const el = document.getElementById(id);
    return el ? (el.textContent || '') : '';
}

function safeGetDeviceID() {
    try {
        return typeof getDeviceID === 'function' ? getDeviceID() : '';
    } catch (e) {
        return '';
    }
}

function generateClientRecordID() {
    if (window.crypto && typeof window.crypto.randomUUID === 'function') {
        return window.crypto.randomUUID();
    }

    return 'loc-' + Date.now() + '-' + Math.floor(Math.random() * 1000000);
}

function isMobileViewport() {
    return window.matchMedia('(max-width: 900px)').matches;
}

function scrollToOtkupStep(stepId) {
    if (!isMobileViewport()) return;

    const run = () => {
        const el = document.getElementById(stepId);
        if (!el) return;

        const rect = el.getBoundingClientRect();
        const absoluteTop = window.scrollY + rect.top;
        const targetY = absoluteTop - 72;

        window.scrollTo({
            top: Math.max(0, targetY),
            behavior: 'smooth'
        });
    };

    requestAnimationFrame(() => {
        requestAnimationFrame(run);
    });
}
