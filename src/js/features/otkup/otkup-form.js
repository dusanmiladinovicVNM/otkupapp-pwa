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

    // ============================================================
    // Picker handlers (Faza 3.1)
    // Sync visual class-picker / pkg-picker → hidden inputs.
    // ============================================================

    const klasaPicker = document.getElementById('klasaPicker');
    if (klasaPicker) {
        klasaPicker.addEventListener('click', function(e) {
            const btn = e.target.closest('.class-pick');
            if (!btn) return;
            klasaPicker.querySelectorAll('.class-pick').forEach(function(b) {
                b.classList.remove('is-sel');
            });
            btn.classList.add('is-sel');
            const hidden = document.getElementById('fldKlasa');
            if (hidden) {
                hidden.value = btn.dataset.klasa || 'I';
                hidden.dispatchEvent(new Event('change', { bubbles: true }));
            }
        });
    }

    const tipAmbalazePicker = document.getElementById('tipAmbalazePicker');
    if (tipAmbalazePicker) {
        const hiddenSelect = document.getElementById('fldTipAmbalaze');

        function syncTipAmbalazePickerFromHidden() {
            if (!hiddenSelect) return;
            const val = hiddenSelect.value || '12/1';
            tipAmbalazePicker.querySelectorAll('.pkg-pick').forEach(function(b) {
                b.classList.toggle('is-sel', b.dataset.tip === val);
            });
        }

        tipAmbalazePicker.addEventListener('click', function(e) {
            const btn = e.target.closest('.pkg-pick');
            if (!btn) return;
            tipAmbalazePicker.querySelectorAll('.pkg-pick').forEach(function(b) {
                b.classList.remove('is-sel');
            });
            btn.classList.add('is-sel');
            const val = btn.dataset.tip || '12/1';
            if (hiddenSelect) {
                let opt = hiddenSelect.querySelector('option[value="' + val + '"]');
                if (!opt) {
                    opt = document.createElement('option');
                    opt.value = val;
                    opt.textContent = val;
                    hiddenSelect.appendChild(opt);
                }
                hiddenSelect.value = val;
                hiddenSelect.dispatchEvent(new Event('change', { bubbles: true }));
            }
        });

        // Sync pills kad populateTipAmbalazeDropdown promeni select.value
        if (hiddenSelect) {
            hiddenSelect.addEventListener('change', syncTipAmbalazePickerFromHidden);
        }

        // Inicijalni sync
        syncTipAmbalazePickerFromHidden();
    }

    bindOtkupFormStateListeners(); 
    
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
    if (typeof withSubmitLock !== 'function') {
        if (typeof window.ensureMasterSyncNotActive === 'function') {
            const allowed = await window.ensureMasterSyncNotActive('saveOtkup', {
                showToast: true
            });

            if (!allowed) return;
        }

        return saveOtkupUnlocked();
    }

    return withSubmitLock('otkup:save', saveOtkupUnlocked, {
        action: 'save-otkup',
        reason: 'saveOtkup',
        alreadyMessage: 'Čuvanje otkupa je već u toku'
    });
}

async function saveOtkupUnlocked() {
    try {
        const input = readOtkupForm();
        const validationError = validateOtkupInput(input);

        if (validationError) {
            showToast(validationError, 'error');
            return;
        }

        // PWA-first BrojDokumenta — kanon "x/ddmmyy[-rb]"
        const brojDokumenta = await generateBrojDokumenta();
        if (!brojDokumenta) {
            showToast('Greška: nije moguće generisati broj dokumenta', 'error');
            return;
        }

        const record = buildOtkupRecord(input, brojDokumenta);

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
                    syncQueueSafe('post-save');
                }
            }, 0);
        }
    } catch (err) {
        console.error('saveOtkup failed:', err);
        showToast('Greška pri čuvanju otkupa', 'error');
    }
}

// ============================================================
// PWA-first BrojDokumenta — kanon "x/ddmmyy[-rb]"
//
// Sequence source: max iz merged local IDB + server (OTK-{otkupacID} sheet,
// koji uključuje i VBA-bulk-pushed brojeve posle VBA stanica unlock-a).
//
// Ako je stanica locked iz VBA, server fetch može da vrati lock signal —
// u tom slučaju koristimo samo local count. To je prihvatljiv kompromis:
// VBA bulk-push-uje na unlock i posle toga PWA vidi sve brojeve. Tokom lock-a
// PWA save u svakom slučaju neće preći sync (sync-engine ga blokira), pa
// privremeni "možda zastareli" broj se reconcile-uje pri sync-u.
// ============================================================
async function generateBrojDokumenta() {
    const today = getTodayIsoDate();
    const otkupacID = CONFIG.OTKUPAC_ID || '';

    const stanicaBrojX = parseInt(String(otkupacID).replace(/\D/g, ''), 10);
    if (!stanicaBrojX || isNaN(stanicaBrojX)) {
        console.error('generateBrojDokumenta: OtkupacID nije validan numerički', otkupacID);
        return '';
    }

    // Lokalni IDB scan
    let localToday = [];
    try {
        const all = await dbGetAll(db, CONFIG.STORE_NAME);
        localToday = (all || []).filter(r =>
            r.datum === today &&
            (r.otkupacID || '') === otkupacID &&
            !r.deleted
        );
    } catch (err) {
        console.error('generateBrojDokumenta local read failed:', err);
    }

    // Server merged scan — getOtkupiForOtkupac uključuje OtkupiAll master +
    // OTK-{otkupacID} live. VBA bulk-push redovi su u live posle unlock-a.
    let serverToday = [];
    if (navigator.onLine) {
        const json = await safeAsync(async () => {
            return await apiFetch(
                'action=getOtkupi&otkupacID=' + encodeURIComponent(otkupacID)
            );
        }, '');  // empty toast — već dovoljno hendlovano

        if (json && json.success && Array.isArray(json.records)) {
            serverToday = json.records.filter(r => {
                const d = toIsoDateOnly(r.Datum);
                return d === today;
            });
        }
        // Ako je GAS vratio lock error, json nije .success → serverToday ostaje [].
        // To je OK: pad-back na local count + dalje sa local-only sequence.
    }

    // Dedupe + max seq
    const seenCrid = new Set();
    let maxSeq = 0;

    function extractSeq(broj) {
        const s = String(broj || '').trim();
        if (!s) return 0;
        const slashPos = s.indexOf('/');
        if (slashPos === -1) return 0;
        const dashPos = s.lastIndexOf('-');
        if (dashPos === -1 || dashPos < slashPos) return 1;
        const tail = s.substring(dashPos + 1);
        const n = parseInt(tail, 10);
        return isNaN(n) ? 0 : n;
    }

    serverToday.forEach(r => {
        const crid = String(r.ClientRecordID || '').trim();
        if (crid) seenCrid.add(crid);
        const s = extractSeq(r.BrojDokumenta);
        if (s > maxSeq) maxSeq = s;
    });

    localToday.forEach(r => {
        const crid = String(r.clientRecordID || '').trim();
        if (crid && seenCrid.has(crid)) return;
        const s = extractSeq(r.brojDokumenta);
        if (s > maxSeq) maxSeq = s;
    });

    const seq = maxSeq + 1;

    const parts = today.split('-');
    if (parts.length !== 3) return '';
    const ddmmyy = parts[2] + parts[1] + parts[0].slice(2);

    return (seq === 1)
        ? `${stanicaBrojX}/${ddmmyy}`
        : `${stanicaBrojX}/${ddmmyy}-${seq}`;
}

function readOtkupForm() {
    const kooperantID = getFieldValue('fldKooperantID');
    const kooperantName = getTextValue('koopName');
    const vrstaVoca = getFieldValue('fldVrsta');
    const sortaVoca = getFieldValue('fldSorta');
    const klasa = getFieldValue('fldKlasa') || 'I';
    const kolicina = parseDecimalInput(getFieldValue('fldKolicina'));
    const cena = parseDecimalInput(getFieldValue('fldCena'));
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

function buildOtkupRecord(input, brojDokumenta) {
    const nowIso = new Date().toISOString();
    const today = getTodayIsoDate();

    return {
        clientRecordID: generateClientRecordID(),
        serverRecordID: '',
        brojDokumenta: brojDokumenta || '',   // ← NOVO polje
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
        const targetY = absoluteTop - 96;

        window.scrollTo({
            top: Math.max(0, targetY),
            behavior: 'smooth'
        });
    };

    requestAnimationFrame(() => {
        requestAnimationFrame(run);
    });
}

// ============================================================
// OTKUP FORM STATE MACHINE (Faza 3.2)
//
// Evaluira trenutno stanje forme i primenjuje:
//   - .is-disabled na step elementima koji čekaju prerequisite
//   - .is-done na step__num kada je step popunjen
//   - btn save disabled state + label / hint toggle
//   - live total kalkulacija u sticky bar-u
//
// Pozove se na svaki promenu polja koja utiče na state.
// ============================================================

function evaluateOtkupFormState() {
    const koopID = (document.getElementById('fldKooperantID') || {}).value || '';
    const vrsta = (document.getElementById('fldVrsta') || {}).value || '';
    const kolicina = parseDecimalInput((document.getElementById('fldKolicina') || {}).value);
    const cena = parseDecimalInput((document.getElementById('fldCena') || {}).value);
    const klasa = (document.getElementById('fldKlasa') || {}).value || '';
    const tipAmb = (document.getElementById('fldTipAmbalaze') || {}).value || '';

    return {
        step1Done: !!koopID,
        step3Done: !!vrsta && !!klasa,
        step4Done: kolicina > 0 && cena > 0 && !!tipAmb,
        canSave: !!koopID && !!vrsta && !!klasa && kolicina > 0 && cena > 0 && !!tipAmb,
        liveTotal: kolicina * cena
    };
}

function applyOtkupFormState() {
    const state = evaluateOtkupFormState();

    // Step disabled state
    const step2 = document.getElementById('otkupStep2ParcelaVozac');
    const step3 = document.getElementById('otkupStep3Roba');
    const step4 = document.getElementById('otkupStep4CenaAmbalaza');
    const step5 = document.getElementById('otkupStep5Napredno');

    [step2, step3, step4, step5].forEach(function(el) {
        if (!el) return;
        el.classList.toggle('is-disabled', !state.step1Done);
    });

    // Step completion check marks
    const setStepDone = function(stepId, done) {
        const step = document.getElementById(stepId);
        if (!step) return;
        const num = step.querySelector('.step__num');
        if (!num) return;
        num.classList.toggle('is-done', !!done);
    };

    setStepDone('otkupStep1Kooperant', state.step1Done);
    setStepDone('otkupStep3Roba', state.step3Done);
    setStepDone('otkupStep4CenaAmbalaza', state.step4Done);

    // Save button state + label
    const btnSave = document.getElementById('btnSaveOtkup');
    const btnLabel = document.getElementById('btnSaveOtkupLabel');
    const btnHint = document.getElementById('btnSaveOtkupHint');

    if (btnSave) {
        btnSave.classList.toggle('is-disabled', !state.canSave);
        btnSave.disabled = !state.canSave;
    }

    if (btnHint) {
        if (state.step1Done) {
            btnHint.style.display = 'none';
        } else {
            btnHint.style.display = '';
            btnHint.textContent = '(skenirajte prvo)';
        }
    }

    if (btnLabel) {
        if (state.canSave && state.liveTotal > 0) {
            const totalFormatted = Math.round(state.liveTotal).toLocaleString('sr-RS');
            btnLabel.textContent = 'Sačuvaj · ' + totalFormatted + ' RSD';
        } else {
            btnLabel.textContent = 'Sačuvaj otkup';
        }
    }
}

function bindOtkupFormStateListeners() {
    const ids = [
        'fldKooperantID',
        'fldKooperantManual',
        'fldVrsta',
        'fldKolicina',
        'fldCena',
        'fldAmbalaza',
        'fldKlasa',
        'fldTipAmbalaze'
    ];

    ids.forEach(function(id) {
        const el = document.getElementById(id);
        if (!el) return;
        el.addEventListener('input', applyOtkupFormState);
        el.addEventListener('change', applyOtkupFormState);
    });

    // Inicijalna primena state-a (sve disabled posle reset-a)
    applyOtkupFormState();

    // Polling fallback za hidden inputs koji se setuju programatski
    setInterval(applyOtkupFormState, 500);
}

// Eksportuj na window da ostali moduli mogu da pozovu nakon QR scan-a
// ili reset-a forme
window.applyOtkupFormState = applyOtkupFormState;
