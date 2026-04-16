// ============================================================
// DIGITALNI AGRONOM — Kooperant Agromere Tab
// PATCHED FOR OFFLINE-FIRST + SYNC CONSISTENCY
// ============================================================

let agroState = {
    parcelaID: '',
    parcelaData: null,
    mera: '',
    artikalID: '',
    artikalData: null,
    kolicina: 0,
    dozaPreporucena: 0,
    opremaTraktor: '',
    opremaPrskalica: '',
    opremaOstalo: '',
    napomena: '',
    timerStart: null,
    timerInterval: null,
    timerResult: null,
    geoStart: null,
    geoEnd: null,
    geoAutoDetect: false,
    meteoOverride: false,
    meteoSnapshot: null,
    karencaDana: 0,
    lager: [],
    opremaList: [],
    geoWatchId: null
};

let _tretmaniCache = { data: null, ts: 0 };
const TRETMANI_CACHE_TTL = 30000; // 30s

// --- Baza čestih naziva opreme za autocomplete ---
const OPREMA_PREDLOZI = {
    Traktor: ['IMT 533', 'IMT 539', 'IMT 542', 'IMT 560', 'IMT 577',
              'John Deere', 'New Holland', 'Massey Ferguson', 'Zetor', 'Ursus',
              'Torpedo', 'Rakovica 65', 'Rakovica 76', 'Belarus', 'Tomo Vinković'],
    Prskalica: ['Agrip 200L', 'Agrip 400L', 'Agrip 600L', 'Morava 440',
                'Holder', 'Stihl SR', 'Atomizer Cifarelli', 'Leđna prskalica',
                'Turbo atomizer', 'Vučena prskalica']
};

// ============================================================
// INIT — poziva se iz showTab('agromere')
// ============================================================
async function loadAgronom() {
    await safeAsync(async () => {
        agroResetState();
        agroLoadLager();
        await agroLoadOprema();
        agroPopulateParcele();
        agroStartGeo();
        await agroLoadIstorija();

        const step1 = document.getElementById('agroStep1');
        const step2 = document.getElementById('agroStep2');
        if (step1) step1.style.display = 'block';
        if (step2) step2.style.display = 'none';

        // background sync if online
        if (navigator.onLine && typeof syncTretmani === 'function') {
            syncTretmani().catch(err => console.error('syncTretmani background failed:', err));
        }
    }, 'Greška pri učitavanju digitalnog agronoma');
}

function agroResetState() {
    if (agroState.timerInterval) clearInterval(agroState.timerInterval);
    if (agroState.geoWatchId) navigator.geolocation.clearWatch(agroState.geoWatchId);

    agroState = {
        parcelaID: '',
        parcelaData: null,
        mera: '',
        artikalID: '',
        artikalData: null,
        kolicina: 0,
        dozaPreporucena: 0,
        opremaTraktor: '',
        opremaPrskalica: '',
        opremaOstalo: '',
        napomena: '',
        timerStart: null,
        timerInterval: null,
        timerResult: null,
        geoStart: null,
        geoEnd: null,
        geoAutoDetect: false,
        meteoOverride: false,
        meteoSnapshot: null,
        karencaDana: 0,
        lager: agroState.lager || [],
        opremaList: agroState.opremaList || [],
        geoWatchId: null
    };

    document.querySelectorAll('.agro-mera-btn').forEach(btn => btn.classList.remove('active'));

    const idsToClear = [
        'agroParcelaSel', 'agroArtikal', 'agroKolicina', 'agroNapomena',
        'agroTraktor', 'agroPrskalica', 'agroOpremaOstalo',
        'agroTraktorNovi', 'agroPrskalicaNovi',
        'agroDozaCalc', 'agroJM', 'agroPreporukaCalc', 'agroPreporukaDetail'
    ];

    idsToClear.forEach(id => {
        const el = document.getElementById(id);
        if (!el) return;
        if ('value' in el) el.value = '';
        else el.textContent = '—';
    });

    const sticky = document.getElementById('agroTimerSticky');
    if (sticky) sticky.style.display = 'none';

    const timerPanel = document.getElementById('agroTimerPanel');
    if (timerPanel) timerPanel.style.display = 'none';

    const btnStart = document.getElementById('agroBtnStart');
    const btnStop = document.getElementById('agroBtnStop');
    if (btnStart) btnStart.style.display = 'none';
    if (btnStop) btnStop.style.display = 'none';
}

// ============================================================
// LAGER — iz stammdaten.magacinkoop
// ============================================================
function agroLoadLager() {
    const koopID = CONFIG.ENTITY_ID;
    const mk = stammdaten.magacinkoop || [];
    agroState.lager = mk.filter(r => r.KooperantID === koopID && parseFloat(r.Stanje) > 0)
        .map(r => ({
            artikalID: r.ArtikalID,
            naziv: r.ArtikalNaziv || r.Naziv || r.ArtikalID,
            tip: r.Tip || '',
            jm: r.JedinicaMere || 'kg',
            cena: parseFloat(r.CenaPoJedinici) || 0,
            dozaPoHa: parseFloat(String(r.DozaPoHa || '0').replace(',', '.')) || 0,
            pakovanje: parseFloat(String(r.Pakovanje || '0').replace(',', '.')) || 0,
            karencaDana: parseInt(r.KarencaDana) || 0,
            primljeno: parseFloat(r.Primljeno) || 0,
            utroseno: parseFloat(r.Utroseno) || 0,
            stanje: parseFloat(r.Stanje) || 0
        }));
}

// ============================================================
// OPREMA
// Keeps current direct-save behavior, but hardens local merge
// ============================================================
async function agroLoadOprema() {
    let serverList = [];

    const json = await safeAsync(async () => {
        return await apiFetch('action=getOprema&kooperantID=' + encodeURIComponent(CONFIG.ENTITY_ID));
    }, 'Greška pri učitavanju opreme');

    if (json && json.success && Array.isArray(json.records)) {
        serverList = json.records.map(r => ({
            clientRecordID: r.ClientRecordID || '',
            naziv: r.Naziv || '',
            tip: r.Tip || '',
            createdAt: normalizeIso(r.CreatedAt),
            source: 'server'
        }));
    }

    // lokalni fallback preko stammdaten / predloga / već unetih
    const merged = new Map();

    serverList.forEach(item => {
        const key = (item.tip + '|' + item.naziv).trim().toLowerCase();
        merged.set(key, item);
    });

    (agroState.opremaList || []).forEach(item => {
        const key = ((item.tip || '') + '|' + (item.naziv || '')).trim().toLowerCase();
        if (!merged.has(key) && item.naziv) merged.set(key, item);
    });

    ['Traktor', 'Prskalica'].forEach(tip => {
        (OPREMA_PREDLOZI[tip] || []).forEach(naziv => {
            const key = (tip + '|' + naziv).trim().toLowerCase();
            if (!merged.has(key)) {
                merged.set(key, { naziv, tip, source: 'preset' });
            }
        });
    });

    agroState.opremaList = Array.from(merged.values());
    agroRenderOpremaSelects();
}

function agroRenderOpremaSelects() {
    const traktorSel = document.getElementById('agroTraktor');
    const prskalicaSel = document.getElementById('agroPrskalica');

    if (!traktorSel || !prskalicaSel) return;

    // --- Traktor ---
    traktorSel.innerHTML = '<option value="">-- Bez traktora --</option>';

    const traktori = (agroState.opremaList || [])
        .filter(o => o.tip === 'Traktor')
        .sort((a, b) => String(a.naziv || '').localeCompare(String(b.naziv || '')));

    const tNames = new Set(traktori.map(o => String(o.naziv || '').toLowerCase()));

    // Kooperantova oprema (server + lokalna)
    traktori.forEach(o => {
        const opt = document.createElement('option');
        opt.value = o.naziv;
        opt.textContent = o.naziv;
        traktorSel.appendChild(opt);
    });

    // Predlozi — samo oni koje kooperant još nema
    const tPredlozi = (OPREMA_PREDLOZI.Traktor || []).filter(n =>
        !tNames.has(String(n).toLowerCase())
    );
    if (tPredlozi.length > 0) {
        const og = document.createElement('optgroup');
        og.label = '— Česti modeli —';
        tPredlozi.forEach(n => {
            const opt = document.createElement('option');
            opt.value = n;
            opt.textContent = n;
            og.appendChild(opt);
        });
        traktorSel.appendChild(og);
    }

    // --- Prskalica ---
    prskalicaSel.innerHTML = '<option value="">-- Bez prskalice --</option>';

    const prskalice = (agroState.opremaList || [])
        .filter(o => o.tip === 'Prskalica' || o.tip === 'Atomizer')
        .sort((a, b) => String(a.naziv || '').localeCompare(String(b.naziv || '')));

    const pNames = new Set(prskalice.map(o => String(o.naziv || '').toLowerCase()));

    prskalice.forEach(o => {
        const opt = document.createElement('option');
        opt.value = o.naziv;
        opt.textContent = o.naziv;
        prskalicaSel.appendChild(opt);
    });

    const pPredlozi = (OPREMA_PREDLOZI.Prskalica || []).filter(n =>
        !pNames.has(String(n).toLowerCase())
    );
    if (pPredlozi.length > 0) {
        const og = document.createElement('optgroup');
        og.label = '— Česti modeli —';
        pPredlozi.forEach(n => {
            const opt = document.createElement('option');
            opt.value = n;
            opt.textContent = n;
            og.appendChild(opt);
        });
        prskalicaSel.appendChild(og);
    }
}

async function agroSaveNovaOprema(tip, value) {
    const naziv = String(value || '').trim();
    if (!naziv) return;

    const exists = (agroState.opremaList || []).some(o =>
        String(o.tip || '').toLowerCase() === String(tip || '').toLowerCase() &&
        String(o.naziv || '').toLowerCase() === naziv.toLowerCase()
    );

    if (!exists) {
        agroState.opremaList.push({
            clientRecordID: agroGenerateClientRecordID('oprema'),
            naziv,
            tip,
            createdAt: new Date().toISOString(),
            source: 'local'
        });
        agroRenderOpremaSelects();
    }

    if (tip === 'Traktor') {
        const sel = document.getElementById('agroTraktor');
        if (sel) sel.value = naziv;
        agroState.opremaTraktor = naziv;
    } else if (tip === 'Prskalica') {
        const sel = document.getElementById('agroPrskalica');
        if (sel) sel.value = naziv;
        agroState.opremaPrskalica = naziv;
    }

    // backend save stays direct for now
    if (navigator.onLine) {
        try {
            await apiPost('syncOprema', {
                kooperantID: CONFIG.ENTITY_ID,
                records: [{
                    clientRecordID: agroGenerateClientRecordID('oprema'),
                    naziv,
                    tip
                }]
            });
        } catch (e) {
            console.error('agroSaveNovaOprema sync failed:', e);
        }
    }
}

// ============================================================
// OPTIONAL: call after field changes if needed
// ============================================================
function agroMarkDraftDirty(record) {
    if (!record) return record;
    record.updatedAtClient = new Date().toISOString();
    record.syncStatus = 'pending';
    record.lastSyncError = '';
    record.lastServerStatus = '';
    return record;
}

// ============================================================
// PARCELA
// ============================================================
function agroPopulateParcele() {
    const sel = document.getElementById('agroParcelaSel');
    if (!sel) return;
    sel.innerHTML = '<option value="">-- Izaberi parcelu --</option>';
    (stammdaten.parcele || []).filter(p => p.KooperantID === CONFIG.ENTITY_ID).forEach(p => {
        const ha = parseFloat(String(p.PovrsinaHa || '0').replace(',', '.')) || 0;
        const o = document.createElement('option');
        o.value = p.ParcelaID;
        o.textContent = (p.KatBroj || p.ParcelaID) + ' — ' + (p.Kultura || '?') + ' (' + ha.toFixed(2) + ' ha)';
        sel.appendChild(o);
    });
}

function onAgroParcelaChange() {
    const pid = document.getElementById('agroParcelaSel').value;
    agroState.parcelaID = pid;
    agroState.parcelaData = (stammdaten.parcele || []).find(p => p.ParcelaID === pid) || null;

    // Meteo strip
    if (pid) {
        loadAgroMeteoStrip(pid);
        checkAgroKarenca(pid);
    } else {
        document.getElementById('agroMeteoStrip').style.display = 'none';
        document.getElementById('agroKarencaWarn').classList.remove('visible');
    }

    // Reset mera
    agroState.mera = '';
    document.querySelectorAll('.agro-mera-btn').forEach(b => b.classList.remove('selected'));
    document.getElementById('agroPreparatSection').classList.remove('visible');
    document.getElementById('agroBtnStart').style.display = pid ? 'none' : 'none';
}

async function loadAgroMeteoStrip(parcelaID) {
    window.meteoCache = window.meteoCache || {};
    const strip = document.getElementById('agroMeteoStrip');
    if (!strip) return;

    strip.style.display = 'flex';

    let data = null;

    // Iz cache-a ako je svež
    if (window.meteoCache[parcelaID] && (Date.now() - window.meteoCache[parcelaID]._ts < 3600000)) {
        data = window.meteoCache[parcelaID];
    } else {
        const json = await safeAsync(async () => {
            return await apiFetch('action=getParcelMeteo&parcelaId=' + encodeURIComponent(parcelaID));
        }, 'Greška pri učitavanju meteo podataka');

        if (json && json.success) {
            json._ts = Date.now();
            window.meteoCache[parcelaID] = json;
            data = json;
        }
    }

    if (!data || !data.current) {
        strip.innerHTML = '<span>Nema meteo podataka</span>';
        agroState.meteoSnapshot = null;
        return;
    }

    const c = data.current;
    const temp = Number(c.temperature || c.Temp || 0);
    const wind = Number(c.windSpeed || c.Wind || 0);
    const hum = Number(c.humidity || c.Humidity || 0);
    const precip = Number(c.precipitation || c.Precip || 0);

    agroState.meteoSnapshot = { temp: temp, wind: wind, humidity: hum };

    const tempEl = document.getElementById('agroMeteoTemp');
    const windEl = document.getElementById('agroMeteoWind');
    const humEl = document.getElementById('agroMeteoHumidity');
    const precipEl = document.getElementById('agroMeteoPrecip');

    if (tempEl) tempEl.textContent = '🌡️ ' + temp.toFixed(1) + '°C';
    if (windEl) windEl.textContent = '💨 ' + wind.toFixed(0) + ' km/h';
    if (humEl) humEl.textContent = '💧 ' + hum + '%';
    if (precipEl) precipEl.textContent = precip > 0 ? '🌧️ ' + precip.toFixed(1) + 'mm' : '☀️ Suvo';
}

// ============================================================
// KARENCA CHECK — za izabranu parcelu
// ============================================================

async function getTretmaniCached(forceRefresh) {
    if (
        !forceRefresh &&
        _tretmaniCache.data &&
        (Date.now() - _tretmaniCache.ts < TRETMANI_CACHE_TTL)
    ) {
        return _tretmaniCache.data;
    }

    let local = [];
    let server = [];

    try {
        local = await dbGetAll(db, 'tretmani');
    } catch (e) {
        console.error('getTretmaniCached local failed:', e);
    }

    const json = await safeAsync(async () => {
        return await apiFetch('action=getTretmani&kooperantID=' + encodeURIComponent(CONFIG.ENTITY_ID));
    }, 'Greška pri učitavanju tretmana');

    if (json && json.success && Array.isArray(json.records)) {
        server = json.records.map(agroMapServerTretman);
    }

    const merged = agroMergeTretmani(local, server);

    _tretmaniCache = { data: merged, ts: Date.now() };
    return merged;
}

function invalidateTretmaniCache() {
    _tretmaniCache = { data: null, ts: 0 };
}

async function checkAgroKarenca(parcelaID) {
    const warn = document.getElementById('agroKarencaWarn');
    const berbaBtn = document.getElementById('agroBerbaBtn');
    if (!warn || !berbaBtn) return;

    warn.classList.remove('visible');
    berbaBtn.classList.remove('disabled');

    const tretmani = await getTretmaniCached(false);

    const parcelTretmani = tretmani.filter(t =>
        t.parcelaID === parcelaID && parseInt(t.karencaDana, 10) > 0 && !t.deleted
    );

    if (parcelTretmani.length === 0) return;

    parcelTretmani.sort((a, b) => (b.datum || '').localeCompare(a.datum || ''));
    const last = parcelTretmani[0];
    const datum = last.datum;
    const karDana = parseInt(last.karencaDana, 10) || 0;
    const prepNaziv = last.artikalNaziv || '?';

    const tretmanDate = new Date(datum);
    const berbeDate = new Date(tretmanDate.getTime() + karDana * 24 * 60 * 60 * 1000);
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    if (berbeDate > today) {
        const ostalo = Math.ceil((berbeDate - today) / (24 * 60 * 60 * 1000));
        warn.classList.add('visible');
        document.getElementById('agroKarencaText').innerHTML =
            '<strong>' + escapeHtml(prepNaziv) + '</strong> — tretman ' + escapeHtml(datum) +
            '<br>Berba dozvoljena: <strong>' + escapeHtml(berbeDate.toLocaleDateString('sr')) +
            '</strong> (još ' + escapeHtml(String(ostalo)) + ' dana)';
        berbaBtn.classList.add('disabled');
    }
}

// ============================================================
// MERA SELECTION
// ============================================================
function selectAgroMera(btn, mera) {
    if (btn.classList.contains('disabled')) {
        showToast('Karenca aktivna — berba nije dozvoljena', 'error');
        return;
    }

    if (!agroState.parcelaID) { showToast('Izaberite parcelu', 'error'); return; }

    document.querySelectorAll('.agro-mera-btn').forEach(b => b.classList.remove('selected'));
    btn.classList.add('selected');
    agroState.mera = mera;

    // Preparat section — samo za Zastita/Prihrana
    const prepSec = document.getElementById('agroPreparatSection');
    if (mera === 'Zastita' || mera === 'Prihrana') {
        prepSec.classList.add('visible');
        agroPopulatePreparati(mera);

        // Meteo check za Zastita
        if (mera === 'Zastita') {
            agroCheckMeteo();
        } else {
            document.getElementById('agroMeteoWarn').classList.remove('visible');
        }
    } else {
        prepSec.classList.remove('visible');
        document.getElementById('agroMeteoWarn').classList.remove('visible');
        document.getElementById('agroPreporuka').classList.remove('visible');
    }

    // Prikaži Start dugme
    document.getElementById('agroBtnStart').style.display = 'block';
    document.getElementById('agroBtnStop').style.display = 'none';
}

// ============================================================
// PREPARAT — filtrirano po tipu + lager
// ============================================================
function agroPopulatePreparati(mera) {
    const sel = document.getElementById('agroPreparatSel');
    sel.innerHTML = '<option value="">-- Izaberi preparat --</option>';
    document.getElementById('agroPreporuka').classList.remove('visible');
    const karInfo = document.getElementById('agroKarencaInfo');
    if (karInfo) karInfo.style.display = 'none';

    const tipFilter = mera === 'Zastita' ? 'Pesticid' : 'Djubrivo';

    const available = agroState.lager.filter(a => a.tip === tipFilter && a.stanje > 0);

    if (available.length === 0) {
        sel.innerHTML = '<option value="">Nema preparata na lageru</option>';
        return;
    }

    available.forEach(a => {
        const o = document.createElement('option');
        o.value = a.artikalID;
        o.textContent = a.naziv + ' — ' + a.stanje.toLocaleString('sr') + ' ' + a.jm + ' na lageru';
        sel.appendChild(o);
    });
}

function onAgroPreparatChange() {
    const artID = document.getElementById('agroPreparatSel').value;
    if (!artID) {
        document.getElementById('agroPreporuka').classList.remove('visible');
        document.getElementById('agroKarencaInfo').style.display = 'none';
        agroState.artikalID = '';
        agroState.artikalData = null;
        return;
    }

    const art = agroState.lager.find(a => a.artikalID === artID);
    agroState.artikalID = artID;
    agroState.artikalData = art;

    // JM
    document.getElementById('agroJM').value = art ? art.jm : '';

    // Karenca info
    const karInfo = document.getElementById('agroKarencaInfo');
    if (art && art.karencaDana > 0) {
        karInfo.style.display = 'block';
        const berbeDate = new Date();
        berbeDate.setDate(berbeDate.getDate() + art.karencaDana);
        document.getElementById('agroKarencaInfoText').innerHTML =
            '⏱️ Karenca: <strong>' + escapeHtml(art.karencaDana) + ' dana</strong> — berba dozvoljena od ' +
            escapeHtml(berbeDate.toLocaleDateString('sr'));
        agroState.karencaDana = art.karencaDana;
    } else {
        karInfo.style.display = 'none';
        agroState.karencaDana = 0;
    }

    // Smart dosage
    agroCalcPreporuka();
}

// ============================================================
// SMART DOSAGE
// ============================================================
function agroCalcPreporuka() {
    const panel = document.getElementById('agroPreporuka');
    const art = agroState.artikalData;
    const parcela = agroState.parcelaData;

    if (!art || !parcela || art.dozaPoHa <= 0) {
        panel.classList.remove('visible');
        return;
    }

    const ha = parseFloat(String(parcela.PovrsinaHa || '0').replace(',', '.')) || 0;
    if (ha <= 0) { panel.classList.remove('visible'); return; }

    const rawQty = art.dozaPoHa * ha;
    let finalQty = rawQty;
    let pakInfo = '';

    if (art.pakovanje > 0) {
        const pakCount = Math.ceil(rawQty / art.pakovanje);
        finalQty = pakCount;
        pakInfo = pakCount + ' × ' + art.pakovanje + ' ' + art.jm + ' (pakovanje)';
    }

    // Proveri da li ima dovoljno na lageru
    let lagerWarn = '';
    const needed = art.pakovanje > 0 ? finalQty * art.pakovanje : finalQty;
    if (needed > art.stanje) {
        lagerWarn = '<br><span style="color:var(--danger);font-weight:600;">⚠️ Nedovoljno na lageru! Imate: ' +
            art.stanje.toLocaleString('sr') + ' ' + art.jm + '</span>';
    }

    panel.classList.add('visible');
    document.getElementById('agroPreporukaCalc').innerHTML =
        '<strong>' + escapeHtml(finalQty.toLocaleString('sr')) + (art.pakovanje > 0 ? ' pak.' : ' ' + escapeHtml(art.jm)) + '</strong> — ' + escapeHtml(art.naziv);
    document.getElementById('agroPreporukaDetail').innerHTML =
        escapeHtml(art.dozaPoHa) + ' ' + escapeHtml(art.jm) + '/ha × ' + escapeHtml(ha.toFixed(2)) + ' ha = ' +
        escapeHtml(rawQty.toLocaleString('sr', {maximumFractionDigits:2})) + ' ' + escapeHtml(art.jm) +
        (pakInfo ? '<br>' + escapeHtml(pakInfo) : '') + lagerWarn;

    agroState.dozaPreporucena = finalQty;
    panel._finalQty = finalQty;
}

function agroPrimeniPreporuku() {
    const panel = document.getElementById('agroPreporuka');
    if (!panel || !panel._finalQty) return;
    document.getElementById('agroKolicina').value = panel._finalQty;
    showToast('Količina: ' + panel._finalQty.toLocaleString('sr'), 'success');
}

// ============================================================
// METEO VALIDATION (samo za Zastita)
// ============================================================
function agroCheckMeteo() {
    const warn = document.getElementById('agroMeteoWarn');
    warn.classList.remove('visible', 'danger');
    agroState.meteoOverride = false;

    if (!agroState.meteoSnapshot) return;

    const m = agroState.meteoSnapshot;
    const parcela = agroState.parcelaData;
    const kultura = parcela ? (parcela.Kultura || '') : '';

    // Pragovi iz CROP_THRESHOLDS (hardcoded za sad, može iz config-a)
    const thresholds = {
        'Visnja': { windMax: 15 }, 'Jabuka': { windMax: 15 }, 'Sljiva': { windMax: 15 },
        'Kruska': { windMax: 15 }, 'Breskva': { windMax: 12 }, 'Malina': { windMax: 12 },
        '_default': { windMax: 15 }
    };
    const th = thresholds[kultura] || thresholds['_default'];

    const warnings = [];

    if (m.wind > th.windMax) {
        warnings.push({ level: 'danger', text: 'Vetar ' + m.wind.toFixed(0) + ' km/h premašuje dozvoljenih ' + th.windMax + ' km/h za ' + (kultura || 'ovu kulturu') });
    }
    if (m.temp < 5) {
        warnings.push({ level: 'danger', text: 'Temperatura ' + m.temp.toFixed(1) + '°C — prenisko za prskanje (min 5°C)' });
    }
    if (m.temp > 35) {
        warnings.push({ level: 'danger', text: 'Temperatura ' + m.temp.toFixed(1) + '°C — previsoko za prskanje (max 35°C)' });
    }
    if (m.humidity > 90) {
        warnings.push({ level: 'warning', text: 'Vlažnost ' + m.humidity + '% — smanjena efikasnost preparata' });
    }

    if (warnings.length === 0) return;

    const isDanger = warnings.some(w => w.level === 'danger');
    warn.classList.add('visible');
    if (isDanger) warn.classList.add('danger');

    document.getElementById('agroMeteoWarnTitle').textContent = isDanger ? '🚫 BLOKADA — Nepovoljni uslovi' : '⚠️ Upozorenje';
    document.getElementById('agroMeteoWarnText').innerHTML = warnings.map(w => w.text).join('<br>');
}

function agroMeteoOverride() {
    agroState.meteoOverride = true;
    document.getElementById('agroMeteoWarn').classList.remove('visible');
    showToast('Meteo override — nastavak na sopstvenu odgovornost', 'info');
}

// ============================================================
// GEOFENCING
// ============================================================
let _lastGeoCheck = 0;
function agroStartGeo() {
    if (!navigator.geolocation) return;
    agroState.geoWatchId = navigator.geolocation.watchPosition(
        pos => {
            const now = Date.now();
            if (now - _lastGeoCheck < 5000) return; // max jednom u 5s
            _lastGeoCheck = now;
            agroState.geoStart = { lat: pos.coords.latitude, lng: pos.coords.longitude };
            agroCheckParcelaProximity(pos.coords.latitude, pos.coords.longitude);
        },
        () => {},
        { enableHighAccuracy: true, maximumAge: 30000, timeout: 15000 }
    );
}

function agroCheckParcelaProximity(lat, lng) {
    // Ako je parcela već izabrana — ne diraj
    if (agroState.parcelaID) return;
    
    const parcele = (stammdaten.parcele || []).filter(p => p.KooperantID === CONFIG.ENTITY_ID);
    const banner = document.getElementById('agroGeoBanner');
    let bestMatch = null, bestDist = Infinity;

    for (const p of parcele) {
        const pLat = parseFloat(String(p.Lat || '').replace(',', '.'));
        const pLng = parseFloat(String(p.Lng || '').replace(',', '.'));
        if (!pLat || !pLng || isNaN(pLat) || isNaN(pLng)) continue;

        // Point-in-polygon check
        if (p.PolygonGeoJSON) {
            try {
                const geom = JSON.parse(p.PolygonGeoJSON);
                if (agroPointInPolygon(lat, lng, geom)) {
                    bestMatch = p; bestDist = 0; break;
                }
            } catch(e) {}
        }

        // Haversine distance
        const dist = agroHaversine(lat, lng, pLat, pLng);
        if (dist < bestDist) { bestDist = dist; bestMatch = p; }
    }

    if (!bestMatch) return;

    const sel = document.getElementById('agroParcelaSel');
    const ha = parseFloat(String(bestMatch.PovrsinaHa || '0').replace(',', '.')) || 0;

    if (bestDist <= 50) {
        banner.className = 'agro-geo-banner detected';
        banner.innerHTML = '📍 Detektovana parcela: <strong>' + escapeHtml(bestMatch.KatBroj || bestMatch.ParcelaID) + '</strong> ' +
             escapeHtml(bestMatch.Kultura || '') + ' (' + escapeHtml(ha.toFixed(2)) + ' ha)';
        sel.value = bestMatch.ParcelaID;
        agroState.geoAutoDetect = true;
        onAgroParcelaChange();
    } else if (bestDist <= 200) {
        banner.className = 'agro-geo-banner nearby';
        banner.innerHTML =
            '📍 Blizu parcele: <strong>' + escapeHtml(bestMatch.KatBroj || bestMatch.ParcelaID) + '</strong> (' +
            escapeHtml(String(Math.round(bestDist))) + 'm) — ' +
            '<a href="#" onclick="document.getElementById(\'agroParcelaSel\').value=\'' +
            String(bestMatch.ParcelaID).replace(/'/g, "\\'") +
            '\';onAgroParcelaChange();return false;" style="color:#92400e;font-weight:700;">Izaberi</a>';
    }
}

function agroHaversine(lat1, lng1, lat2, lng2) {
    const R = 6371000;
    const dLat = (lat2 - lat1) * Math.PI / 180;
    const dLng = (lng2 - lng1) * Math.PI / 180;
    const a = Math.sin(dLat/2)**2 + Math.cos(lat1*Math.PI/180) * Math.cos(lat2*Math.PI/180) * Math.sin(dLng/2)**2;
    return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
}

function agroPointInPolygon(lat, lng, geometry) {
    const coords = geometry.coordinates ? geometry.coordinates[0] : geometry[0];
    if (!coords) return false;
    let inside = false;
    for (let i = 0, j = coords.length - 1; i < coords.length; j = i++) {
        const xi = coords[i][1], yi = coords[i][0];
        const xj = coords[j][1], yj = coords[j][0];
        if (((yi > lng) !== (yj > lng)) && (lat < (xj - xi) * (lng - yi) / (yj - yi) + xi)) {
            inside = !inside;
        }
    }
    return inside;
}

// ============================================================
// TIMER
// ============================================================
function agroStartRad() {
    if (!agroState.parcelaID) { showToast('Izaberite parcelu', 'error'); return; }
    if (!agroState.mera) { showToast('Izaberite meru', 'error'); return; }

    // Za Zastita/Prihrana — proveri da li je preparat izabran
    if ((agroState.mera === 'Zastita' || agroState.mera === 'Prihrana') && !agroState.artikalID) {
        showToast('Izaberite preparat', 'error'); return;
    }

    // Snimi količinu
    if (agroState.artikalID) {
        agroState.kolicina = parseFloat(document.getElementById('agroKolicina').value) || 0;
    }

    // Snimi opremu
    agroState.opremaTraktor = document.getElementById('agroTraktor').value || document.getElementById('agroTraktorNovi').value || '';
    agroState.opremaPrskalica = document.getElementById('agroPrskalica').value || document.getElementById('agroPrskalicaNovi').value || '';
    agroState.opremaOstalo = document.getElementById('agroOpremaOstalo').value || '';
    agroState.napomena = document.getElementById('agroNapomena').value || '';

    // GPS start
    if (navigator.geolocation) {
        navigator.geolocation.getCurrentPosition(pos => {
            agroState.geoStart = { lat: pos.coords.latitude, lng: pos.coords.longitude };
        }, () => {}, { enableHighAccuracy: true });
    }

    // Start timer
    agroState.timerStart = new Date();
    agroState.timerInterval = setInterval(agroUpdateTimer, 1000);
    sessionStorage.setItem('agroTimerStart', agroState.timerStart.toISOString());

    // UI
    document.getElementById('agroTimerPanel').style.display = 'block';
    document.getElementById('agroTimerLabel').textContent =
        (agroState.parcelaData ? agroState.parcelaData.KatBroj : agroState.parcelaID) + ' — ' + agroState.mera;
    document.getElementById('agroBtnStart').style.display = 'none';
    document.getElementById('agroBtnStop').style.display = 'block';
    document.getElementById('agroTimerSticky').classList.add('active');

    showToast('Tajmer pokrenut', 'success');
}

function agroUpdateTimer() {
    if (!agroState.timerStart) return;
    const elapsed = Math.floor((new Date() - agroState.timerStart) / 1000);
    const h = String(Math.floor(elapsed / 3600)).padStart(2, '0');
    const m = String(Math.floor((elapsed % 3600) / 60)).padStart(2, '0');
    const s = String(elapsed % 60).padStart(2, '0');
    const display = h + ':' + m + ':' + s;
    document.getElementById('agroTimerDisplay').textContent = display;
    document.getElementById('agroTimerStickyText').textContent = '⏱️ ' + display + ' — ' + agroState.mera;
}

function agroStopRad() {
    clearInterval(agroState.timerInterval);
    const end = new Date();
    const trajanjeMin = Math.round((end - agroState.timerStart) / 60000);
    sessionStorage.removeItem('agroTimerStart');

    // GPS end
    if (navigator.geolocation) {
        navigator.geolocation.getCurrentPosition(pos => {
            agroState.geoEnd = { lat: pos.coords.latitude, lng: pos.coords.longitude };
        }, () => {}, { enableHighAccuracy: true });
    }

    agroState.timerResult = {
        pocetakISO: agroState.timerStart.toISOString(),
        zavrsetakISO: end.toISOString(),
        trajanjeMinuta: trajanjeMin
    };

    // Prikaži potvrdu
    document.getElementById('agroBtnStop').style.display = 'none';
    document.getElementById('agroTimerSticky').classList.remove('active');
    agroShowConfirm();
}

// ============================================================
// POTVRDA
// ============================================================
function agroShowConfirm() {
    document.getElementById('agroStep1').style.display = 'none';
    document.getElementById('agroStep2').style.display = 'block';

    const p = agroState.parcelaData;
    const art = agroState.artikalData;
    const timer = agroState.timerResult;

    const rows = [];
    rows.push(['Parcela', (p ? p.KatBroj : agroState.parcelaID) + ' — ' + (p ? p.Kultura : '')]);
    rows.push(['Površina', p ? (parseFloat(String(p.PovrsinaHa || '0').replace(',', '.')) || 0).toFixed(2) + ' ha' : '?']);
    rows.push(['Mera', agroState.mera]);

    if (art) {
        rows.push(['Preparat', art.naziv]);
        rows.push(['Količina', agroState.kolicina + ' ' + art.jm]);
        if (agroState.karencaDana > 0) {
            const berbeDate = new Date();
            berbeDate.setDate(berbeDate.getDate() + agroState.karencaDana);
            rows.push(['Karenca', agroState.karencaDana + ' dana → berba od ' + berbeDate.toLocaleDateString('sr')]);
        }
    }

    if (agroState.opremaTraktor) rows.push(['Traktor', agroState.opremaTraktor]);
    if (agroState.opremaPrskalica) rows.push(['Prskalica', agroState.opremaPrskalica]);
    if (agroState.opremaOstalo) rows.push(['Ostala oprema', agroState.opremaOstalo]);

    if (timer) {
        const h = Math.floor(timer.trajanjeMinuta / 60);
        const m = timer.trajanjeMinuta % 60;
        rows.push(['Trajanje', (h > 0 ? h + 'h ' : '') + m + ' min']);
    }

    if (agroState.meteoSnapshot) {
        rows.push(['Meteo', agroState.meteoSnapshot.temp.toFixed(1) + '°C, vetar ' + agroState.meteoSnapshot.wind.toFixed(0) + ' km/h, vlažnost ' + agroState.meteoSnapshot.humidity + '%']);
        if (agroState.meteoOverride) rows.push(['Meteo override', '⚠️ Da — nastavljeno uprkos upozorenju']);
    }

    if (agroState.napomena) rows.push(['Napomena', agroState.napomena]);
    if (agroState.geoAutoDetect) rows.push(['GPS', '📍 Auto-detect']);

    document.getElementById('agroConfirmPanel').innerHTML = rows.map(r =>
         '<div class="agro-confirm-row"><span class="label">' + escapeHtml(r[0]) + '</span><span class="value">' + escapeHtml(r[1]) + '</span></div>'
    ).join('');
}

function agroBackToStep1() {
    document.getElementById('agroStep1').style.display = 'block';
    document.getElementById('agroStep2').style.display = 'none';
}

// ============================================================
// SAVE TRETMAN
// ============================================================
async function agroSaveTretman() {
    const art = agroState.artikalData;
    const timer = agroState.timerResult;
    const now = new Date();
    const nowIso = now.toISOString();

    let datumBerbeDozvoljeno = '';
    if (agroState.karencaDana > 0) {
        const d = new Date();
        d.setDate(d.getDate() + agroState.karencaDana);
        datumBerbeDozvoljeno = d.toISOString().split('T')[0];
    }

    const record = {
        clientRecordID: agroGenerateClientRecordID('tretman'),
        serverRecordID: '',
        createdAtClient: nowIso,
        updatedAtClient: nowIso,
        updatedAtServer: '',
        syncedAt: '',

        kooperantID: CONFIG.ENTITY_ID,
        parcelaID: agroState.parcelaID,
        datum: nowIso.split('T')[0],
        mera: agroState.mera,

        artikalID: art ? art.artikalID : '',
        artikalNaziv: art ? art.naziv : '',
        kolicinaUpotrebljena: agroState.kolicina || '',
        jedinicaMere: art ? art.jm : '',
        dozaPreporucena: agroState.dozaPreporucena || '',
        dozaPrimenjena: agroState.kolicina || '',

        opremaTraktor: agroState.opremaTraktor,
        opremaPrskalica: agroState.opremaPrskalica,
        opremaOstalo: agroState.opremaOstalo,

        karencaDana: agroState.karencaDana || '',
        datumBerbeDozvoljeno: datumBerbeDozvoljeno,

        vremePocetka: timer ? timer.pocetakISO : '',
        vremeZavrsetka: timer ? timer.zavrsetakISO : '',
        trajanjeMinuta: timer ? timer.trajanjeMinuta : '',

        geoLatStart: agroState.geoStart ? agroState.geoStart.lat : '',
        geoLngStart: agroState.geoStart ? agroState.geoStart.lng : '',
        geoLatEnd: agroState.geoEnd ? agroState.geoEnd.lat : '',
        geoLngEnd: agroState.geoEnd ? agroState.geoEnd.lng : '',
        geoAutoDetect: agroState.geoAutoDetect ? 'Da' : '',

        meteoTemp: agroState.meteoSnapshot ? agroState.meteoSnapshot.temp : '',
        meteoWind: agroState.meteoSnapshot ? agroState.meteoSnapshot.wind : '',
        meteoHumidity: agroState.meteoSnapshot ? agroState.meteoSnapshot.humidity : '',
        meteoOverride: agroState.meteoOverride ? 'Da' : '',

        napomena: agroState.napomena,

        syncStatus: 'pending',
        syncAttempts: 0,
        syncAttemptAt: '',
        lastSyncError: '',
        lastServerStatus: '',
        deleted: false,
        entityType: 'tretman',
        schemaVersion: 1
    };

    await dbPut(db, 'tretmani', record);
    showToast('Tretman sačuvan!', 'success');

    // Invalidate cache jer smo dodali nov record
    invalidateTretmaniCache();

    // Sync ako smo online
    if (navigator.onLine && typeof syncTretmani === 'function') {
        try {
            await syncTretmani();
            invalidateTretmaniCache(); // ponovo invalidate jer sync menja statuse
        } catch (e) {
            console.error('syncTretmani after save failed:', e);
        }
    }

    // Reset UI
    agroResetState();
    agroPopulateParcele();

    // Jedan poziv istorije (ne dva kao pre)
    await agroLoadIstorija();

    const step1 = document.getElementById('agroStep1');
    const step2 = document.getElementById('agroStep2');
    if (step1) step1.style.display = 'block';
    if (step2) step2.style.display = 'none';
}
// ============================================================
// ISTORIJA
// ============================================================
async function agroLoadIstorija() {
    const all = (await getTretmaniCached(false))
        .filter(r => !r.deleted)
        .sort((a, b) => {
            const byDate = (b.datum || '').localeCompare(a.datum || '');
            if (byDate !== 0) return byDate;

            const byTime = String(b.updatedAtClient || b.createdAtClient || b.updatedAtServer || '')
                .localeCompare(String(a.updatedAtClient || a.createdAtClient || a.updatedAtServer || ''));
            if (byTime !== 0) return byTime;

            return String(b.clientRecordID || '').localeCompare(String(a.clientRecordID || ''));
        });

    const list = document.getElementById('agroTretmaniList');
    const icons = { Zastita: '🛡️', Prihrana: '🌱', Rezidba: '✂️', Zalivanje: '💧', Berba: '🍎' };

    if (!list) return;

    if (all.length === 0) {
        list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema tretmana</p>';
        return;
    }

    list.innerHTML = all.map(r => {
        const bc = (r.syncStatus === 'pending' || r.syncStatus === 'syncing')
            ? 'var(--warning)'
            : 'var(--success)';

        const min = parseInt(r.trajanjeMinuta, 10) || 0;
        const timeStr = min > 0
            ? (Math.floor(min / 60) > 0 ? Math.floor(min / 60) + 'h ' : '') + (min % 60) + 'min'
            : '';

        const syncText =
            r.syncStatus === 'syncing' ? ' | sync...' :
            r.syncStatus === 'pending' ? ' | pending' :
            (r.serverRecordID ? ' | ' + r.serverRecordID : '');

        return `<div class="queue-item" style="border-left-color:${bc};">
            <div class="qi-header">
                <span class="qi-koop">${icons[r.mera] || ''} ${escapeHtml(r.mera || '')}</span>
                <span class="qi-time">${escapeHtml(r.datum || '')}</span>
            </div>
            <div class="qi-detail">
                ${escapeHtml(r.parcelaID || '')}
                ${r.artikalNaziv ? ' | ' + escapeHtml(r.artikalNaziv) + ' ' + escapeHtml(String(r.kolicinaUpotrebljena || '')) + ' ' + escapeHtml(r.jedinicaMere || '') : ''}
                ${timeStr ? ' | ⏱️ ' + escapeHtml(timeStr) : ''}
                ${r.opremaTraktor ? ' | 🚜 ' + escapeHtml(r.opremaTraktor) : ''}
                ${r.karencaDana ? ' | Karenca ' + escapeHtml(String(r.karencaDana)) + 'd' : ''}
                ${syncText}
            </div>
            ${r.lastSyncError ? `<div style="margin-top:6px;font-size:12px;color:#b42318;">${escapeHtml(r.lastSyncError)}</div>` : ''}
        </div>`;
    }).join('');
}


// ============================================================
// HELPERS
// ============================================================
function agroGenerateClientRecordID(prefix) {
    if (window.crypto && typeof window.crypto.randomUUID === 'function') {
        return window.crypto.randomUUID();
    }
    return (prefix || 'agro') + '-' + Date.now() + '-' + Math.floor(Math.random() * 1000000);
}

function agroMapServerTretman(r) {
    return {
        clientRecordID: r.ClientRecordID || '',
        serverRecordID: r.ServerRecordID || '',
        createdAtClient: normalizeIso(r.CreatedAtClient),
        updatedAtClient: normalizeIso(r.UpdatedAtClient || r.CreatedAtClient),
        updatedAtServer: normalizeIso(r.UpdatedAtServer || r.ReceivedAt),
        syncedAt: normalizeIso(r.UpdatedAtServer || r.ReceivedAt),

        kooperantID: r.KooperantID || '',
        parcelaID: r.ParcelaID || '',
        datum: fmtDate(r.Datum),
        mera: r.Mera || '',

        artikalID: r.ArtikalID || '',
        artikalNaziv: r.ArtikalNaziv || '',
        kolicinaUpotrebljena: r.KolicinaUpotrebljena || '',
        jedinicaMere: r.JedinicaMere || '',
        dozaPreporucena: r.DozaPreporucena || '',
        dozaPrimenjena: r.DozaPrimenjena || '',

        opremaTraktor: r.OpremaTraktor || '',
        opremaPrskalica: r.OpremaPrskalica || '',
        opremaOstalo: r.OpremaOstalo || '',

        karencaDana: r.KarencaDana || '',
        datumBerbeDozvoljeno: r.DatumBerbeDozvoljeno || '',

        vremePocetka: r.VremePocetka || '',
        vremeZavrsetka: r.VremeZavrsetka || '',
        trajanjeMinuta: r.TrajanjeMinuta || '',

        geoLatStart: r.GeoLatStart || '',
        geoLngStart: r.GeoLngStart || '',
        geoLatEnd: r.GeoLatEnd || '',
        geoLngEnd: r.GeoLngEnd || '',
        geoAutoDetect: r.GeoAutoDetect || '',

        meteoTemp: r.MeteoTemp || '',
        meteoWind: r.MeteoWind || '',
        meteoHumidity: r.MeteoHumidity || '',
        meteoOverride: r.MeteoOverride || '',

        napomena: r.Napomena || '',

        syncStatus: 'synced',
        syncAttempts: 0,
        syncAttemptAt: '',
        lastSyncError: '',
        lastServerStatus: 'server',
        deleted: false,
        entityType: 'tretman',
        schemaVersion: 1
    };
}

function agroNormalizeLocalTretman(r) {
    return {
        clientRecordID: r.clientRecordID || '',
        serverRecordID: r.serverRecordID || '',
        createdAtClient: normalizeIso(r.createdAtClient),
        updatedAtClient: normalizeIso(r.updatedAtClient || r.createdAtClient),
        updatedAtServer: normalizeIso(r.updatedAtServer),
        syncedAt: normalizeIso(r.syncedAt),

        kooperantID: r.kooperantID || '',
        parcelaID: r.parcelaID || '',
        datum: r.datum || '',
        mera: r.mera || '',

        artikalID: r.artikalID || '',
        artikalNaziv: r.artikalNaziv || '',
        kolicinaUpotrebljena: r.kolicinaUpotrebljena || '',
        jedinicaMere: r.jedinicaMere || '',
        dozaPreporucena: r.dozaPreporucena || '',
        dozaPrimenjena: r.dozaPrimenjena || '',

        opremaTraktor: r.opremaTraktor || '',
        opremaPrskalica: r.opremaPrskalica || '',
        opremaOstalo: r.opremaOstalo || '',

        karencaDana: r.karencaDana || '',
        datumBerbeDozvoljeno: r.datumBerbeDozvoljeno || '',

        vremePocetka: r.vremePocetka || '',
        vremeZavrsetka: r.vremeZavrsetka || '',
        trajanjeMinuta: r.trajanjeMinuta || '',

        geoLatStart: r.geoLatStart || '',
        geoLngStart: r.geoLngStart || '',
        geoLatEnd: r.geoLatEnd || '',
        geoLngEnd: r.geoLngEnd || '',
        geoAutoDetect: r.geoAutoDetect || '',

        meteoTemp: r.meteoTemp || '',
        meteoWind: r.meteoWind || '',
        meteoHumidity: r.meteoHumidity || '',
        meteoOverride: r.meteoOverride || '',

        napomena: r.napomena || '',

        syncStatus: r.syncStatus || 'pending',
        syncAttempts: parseInt(r.syncAttempts, 10) || 0,
        syncAttemptAt: r.syncAttemptAt || '',
        lastSyncError: r.lastSyncError || '',
        lastServerStatus: r.lastServerStatus || '',
        deleted: !!r.deleted,
        entityType: r.entityType || 'tretman',
        schemaVersion: r.schemaVersion || 1
    };
}

function agroMergeTretmani(local, server) {
    return mergeOfflineRecords(local, server, agroNormalizeLocalTretman);
}

function showRadoviSection(name, btn) {
    document.querySelectorAll('.radovi-section').forEach(el => el.classList.remove('active'));
    document.querySelectorAll('.radovi-subnav-btn').forEach(el => el.classList.remove('active'));

    const section = document.getElementById('radovi-section-' + name);
    if (section) section.classList.add('active');

    if (btn && btn.classList) {
        btn.classList.add('active');
    }
}

function scrollRadoviFormIntoView() {
    const el = document.getElementById('radoviFormAnchor');
    if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' });
}
