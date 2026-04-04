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

function applyDefaults() {
    const config = stammdaten.config || [];

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
            if (typeof syncQueueSafe === 'function') {
                await syncQueueSafe();
            } else if (typeof syncQueue === 'function') {
                await syncQueue();
            }
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
        kolAmbalaze,
        parcelaID,
        napomena,
        vozacID
    };
}

function validateOtkupInput(input) {
    if (!input.kooperantID) return 'Skenirajte ili izaberite kooperanta';
    if (!input.vrstaVoca) return 'Izaberite vrstu voća';
    if (input.kolicina <= 0) return 'Unesite količinu';
    if (input.cena <= 0) return 'Unesite cenu';

    return '';
}

function buildOtkupRecord(input) {
    const nowIso = new Date().toISOString();
    const today = nowIso.split('T')[0];

    return {
        clientRecordID: crypto.randomUUID(),
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

        tipAmbalaze: '12/1',
        kolAmbalaze: input.kolAmbalaze,

        parcelaID: input.parcelaID || '',
        napomena: input.napomena || '',
        vozacID: input.vozacID || '',

        syncStatus: 'pending',
        syncAttempts: 0,
        lastSyncError: '',
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
    setFieldValue('fldAmbalaza', '');
    setFieldValue('fldNapomena', '');
    setFieldValue('fldVozacID', '');

    const vozacDisplay = document.getElementById('vozacDisplay');
    if (vozacDisplay) vozacDisplay.classList.remove('visible');

    applyDefaults();
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
