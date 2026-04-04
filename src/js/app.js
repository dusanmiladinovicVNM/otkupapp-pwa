// ============================================================
// STATE
// ============================================================
let db = null;
let qrScanner = null;
let stammdaten = { kooperanti: [], kulture: [], config: [], parcele: [], stanice: [], kupci: [], vozaci: [] };
let selectedMera = '';
let mgmtData = null;
let parcelExpertOpen = {};



// ============================================================
// INIT
// ============================================================
document.addEventListener('DOMContentLoaded', async () => {
    if (!getLs('authToken', '') || !getLs('otkupacID', '')) {
        showLoginScreen();
        return;
    }
    db = await openDB();
    await loadStammdaten();
    applyRoleVisibility();
    document.getElementById('headerInfo').textContent = CONFIG.USER_ROLE + ': ' + CONFIG.ENTITY_NAME;

    if (CONFIG.USER_ROLE === 'Otkupac') {
        populateVrstaDropdown();
        applyDefaults();
        const today = new Date().toISOString().split('T')[0];
        document.getElementById('fldPregledOd').value = today;
        document.getElementById('fldPregledDo').value = today;
        const otpDatumEl = document.getElementById('fldOtpremniceDatum');
        if (otpDatumEl) otpDatumEl.value = today;
    }
    if (CONFIG.USER_ROLE === 'Kooperant') {
        populateAgroParcele();
    }
    if (CONFIG.USER_ROLE === 'Management') {
        populateMgmtStanice();
        const today = new Date().toISOString().split('T')[0];
        document.getElementById('mgmtOtkupiOd').value = today;
        document.getElementById('mgmtOtkupiDo').value = today;
        prefetchMgmtData().then(() => { 
            populateMgmtKupciDropdown();
            showTab('dispecer');
        });
    }
    updateSyncBadge();
    window.addEventListener('online', () => { updateSyncBadge(); if (CONFIG.USER_ROLE === 'Otkupac') syncQueue(); });
    window.addEventListener('offline', () => updateSyncBadge());
    setInterval(() => { if (navigator.onLine && CONFIG.USER_ROLE === 'Otkupac') syncQueue(); }, 60000);
});



// ============================================================
// QR SCANNER
// ============================================================
function startQRScan() {
    const readerDiv = document.getElementById('qr-reader');
    readerDiv.style.display = 'block';
    if (qrScanner) qrScanner.clear();
    qrScanner = new Html5Qrcode('qr-reader');
    qrScanner.start(
        { facingMode: 'environment' },
        { fps: 10, qrbox: { width: 250, height: 250 } },
        (decodedText) => { onQRScanned(decodedText); qrScanner.stop().then(() => { readerDiv.style.display = 'none'; }); },
        () => {}
    ).catch(err => { showToast('Kamera nije dostupna: ' + err, 'error'); readerDiv.style.display = 'none'; });
}

function onQRScanned(text) {
    try { const data = JSON.parse(text); if (data.id) { setKooperant(data.id, data.name || data.id); return; } } catch (e) {}
    if (text.startsWith('KOOP-')) {
        const koop = stammdaten.kooperanti.find(k => k.KooperantID === text);
        setKooperant(text, koop ? (koop.Ime + ' ' + koop.Prezime) : text);
        return;
    }
    showToast('Nepoznat QR kod', 'error');
}

function setKooperant(id, name) {
    document.getElementById('fldKooperantID').value = id;
    document.getElementById('koopName').textContent = name;
    document.getElementById('koopId').textContent = id;
    document.getElementById('koopDisplay').classList.add('visible');
    showToast('Kooperant: ' + name, 'success');
    populateParcelaDropdown(id);
}

function startVozacQRScan() {
    const readerDiv = document.getElementById('qr-reader-vozac');
    readerDiv.style.display = 'block';
    const scanner = new Html5Qrcode('qr-reader-vozac');
    scanner.start(
        { facingMode: 'environment' },
        { fps: 10, qrbox: { width: 250, height: 250 } },
        (decodedText) => {
            onVozacQRScanned(decodedText);
            scanner.stop().then(() => { readerDiv.style.display = 'none'; });
        },
        () => {}
    ).catch(err => { showToast('Kamera nije dostupna: ' + err, 'error'); readerDiv.style.display = 'none'; });
}

function onVozacQRScanned(text) {
    try {
        const data = JSON.parse(text);
        if (data.type === 'VOZ' && data.id) { setVozac(data.id, data.name || data.id); return; }
    } catch (e) {}
    if (text.startsWith('VOZ-')) {
        setVozac(text, text);
        return;
    }
    showToast('Nije QR vozača', 'error');
}

function setVozac(id, name) {
    document.getElementById('fldVozacID').value = id;
    document.getElementById('vozacName').textContent = name;
    document.getElementById('vozacId').textContent = id;
    document.getElementById('vozacDisplay').classList.add('visible');
    showToast('Vozač: ' + name, 'success');
}

function clearVozac() {
    document.getElementById('fldVozacID').value = '';
    document.getElementById('vozacDisplay').classList.remove('visible');
}
    
// ============================================================
// QR PROFILE
// ============================================================
function showQRProfile() {
    const modal = document.getElementById('qrProfileModal');
    document.getElementById('qrProfileName').textContent = CONFIG.ENTITY_NAME;
    document.getElementById('qrProfileRole').textContent = CONFIG.USER_ROLE;
    document.getElementById('qrProfileID').textContent = CONFIG.ENTITY_ID;
    modal.style.display = 'flex';
    
    generateQRCode('qrProfileCanvas', JSON.stringify({
        type: CONFIG.USER_ROLE === 'Kooperant' ? 'KOOP' : CONFIG.USER_ROLE === 'Otkupac' ? 'OTK' : CONFIG.USER_ROLE === 'Vozac' ? 'VOZ' : 'MGMT',
        id: CONFIG.ENTITY_ID,
        name: CONFIG.ENTITY_NAME
    }));
}

function generateQRCode(canvasId, text) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    
    // Koristi QR generisanje iz eksternog CDN-a
    const img = new Image();
    img.onload = function() {
        canvas.width = 250;
        canvas.height = 250;
        const ctx = canvas.getContext('2d');
        ctx.fillStyle = 'white';
        ctx.fillRect(0, 0, 250, 250);
        ctx.drawImage(img, 0, 0, 250, 250);
    };
    img.onerror = function() {
        // Fallback: prikaži tekst ako API ne radi
        canvas.width = 250;
        canvas.height = 250;
        const ctx = canvas.getContext('2d');
        ctx.fillStyle = 'white';
        ctx.fillRect(0, 0, 250, 250);
        ctx.fillStyle = '#1a5e2a';
        ctx.font = 'bold 16px sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText(CONFIG.ENTITY_ID, 125, 125);
    };
    img.src = 'https://api.qrserver.com/v1/create-qr-code/?size=250x250&data=' + encodeURIComponent(text);
}
    
// ============================================================
// STAMMDATEN
// ============================================================
async function loadStammdaten() {
    try {
        const cached = await dbGetAll(db, CONFIG.STAMM_STORE);
        const obj = cached.find(c => c.key === 'all');
        if (obj) stammdaten = obj.data;
    } catch (e) {}

    if (navigator.onLine) {
        try {
            const json = await apiFetch('action=getStammdaten');
            if (json && json.success && json.data) {
                stammdaten = json.data;
                await dbPut(db, CONFIG.STAMM_STORE, {
                    key: 'all',
                    data: stammdaten,
                    updatedAt: new Date().toISOString()
                });
            }
        } catch (e) {}
    }
}
    
function fmtStanica(stanicaID) {
    if (!stanicaID) return '';
    const s = (stammdaten.stanice || []).find(s => s.StanicaID === stanicaID);
    const name = s ? (s.Naziv || s.Mesto || stanicaID) : stanicaID;
    if (name === stanicaID) return stanicaID;
    return name + ' <span style="font-size:11px;color:var(--text-muted);">' + stanicaID + '</span>';
}


function initSignaturePad(canvasId) {
    const canvas = document.getElementById(canvasId); if (!canvas) return;
    const ctx = canvas.getContext('2d');
    const rect = canvas.getBoundingClientRect();
    const scaleX = canvas.width / rect.width;
    const scaleY = canvas.height / rect.height;
    ctx.scale(scaleX, scaleY);
    ctx.strokeStyle = '#1a1a1a';
    ctx.lineWidth = 2;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    let drawing = false, lastX = 0, lastY = 0;
    function getPos(e) { const r = canvas.getBoundingClientRect(); const touch = e.touches ? e.touches[0] : e; return { x: touch.clientX - r.left, y: touch.clientY - r.top }; }
    function startDraw(e) { e.preventDefault(); drawing = true; const p = getPos(e); lastX = p.x; lastY = p.y; }
    function draw(e) { if (!drawing) return; e.preventDefault(); const p = getPos(e); ctx.beginPath(); ctx.moveTo(lastX, lastY); ctx.lineTo(p.x, p.y); ctx.stroke(); lastX = p.x; lastY = p.y; }
    function stopDraw(e) { if (e) e.preventDefault(); drawing = false; }
    canvas.addEventListener('mousedown', startDraw); canvas.addEventListener('mousemove', draw);
    canvas.addEventListener('mouseup', stopDraw); canvas.addEventListener('mouseleave', stopDraw);
    canvas.addEventListener('touchstart', startDraw, { passive: false });
    canvas.addEventListener('touchmove', draw, { passive: false });
    canvas.addEventListener('touchend', stopDraw, { passive: false });
}

function clearSignature(canvasId) {
    const canvas = document.getElementById(canvasId); if (!canvas) return;
    canvas.getContext('2d').clearRect(0, 0, canvas.width, canvas.height);
}

function getSignatureData(canvasId) {
    const canvas = document.getElementById(canvasId); if (!canvas) return '';
    const data = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data;
    if (!data.some((val, i) => i % 4 === 3 && val > 0)) return '';
    return canvas.toDataURL('image/png');
}
    
// ============================================================
// OTPREMA (dispatch)
// ============================================================
let otpremaVozacID = '';
let otpremaUnassigned = [];

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
    document.getElementById('otpremaMainView').style.display = 'none';
    document.getElementById('otpremaAssignView').style.display = 'block';

    const today = new Date().toISOString().split('T')[0];
    
    const local = await dbGetAll(db, CONFIG.STORE_NAME);
    
    let server = [];
    try {
        const json = await apiFetch('action=getOtkupi&otkupacID=' + encodeURIComponent(CONFIG.OTKUPAC_ID));
        if (json && json.success && json.records) server = json.records.map(r => ({
            clientRecordID: r.ClientRecordID || '', datum: fmtDate(r.Datum),
            kooperantID: r.KooperantID || '', kooperantName: r.KooperantName || r.KooperantID || '',
            vrstaVoca: r.VrstaVoca || '', sortaVoca: r.SortaVoca || '',
            klasa: r.Klasa || 'I', kolicina: parseFloat(r.Kolicina) || 0,
            cena: parseFloat(r.Cena) || 0, kolAmbalaze: parseInt(r.KolAmbalaze) || 0,
            parcelaID: r.ParcelaID || '', vozacID: r.VozacID || r.VozaciID || '',
            napomena: r.Napomena || '', syncStatus: 'synced'
        }));
    } catch (e) {}
    
    const serverIDs = new Set(server.map(r => r.clientRecordID));
    const all = [...server, ...local.filter(r => r.syncStatus === 'pending' && !serverIDs.has(r.clientRecordID))];
    
    otpremaUnassigned = all.filter(r => !r.vozacID);
    otpremaUnassigned.sort((a, b) => (b.datum || '').localeCompare(a.datum || ''));
    renderOtpremaCheckboxes();
}

function renderOtpremaCheckboxes() {
    const list = document.getElementById('otpremaOtkupList');
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
                    <div class="qi-header"><span class="qi-koop">${r.kooperantName}</span><span class="qi-time">${r.datum}</span></div>
                    <div class="qi-detail" style="font-size:11px;color:var(--text-muted);">${r.klasa || 'I'}</div>
                    <div class="qi-detail">${r.vrstaVoca} ${r.sortaVoca || ''} | ${r.kolicina} kg × ${r.cena} = ${vr} RSD</div>
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
    let kg = 0, count = 0;
    otpremaUnassigned.forEach((r, i) => {
        const chk = document.getElementById('otpChk' + i);
        if (chk && chk.checked) { kg += r.kolicina || 0; count++; }
    });
    const div = document.getElementById('otpremaSummary');
    if (count > 0) {
        div.style.display = 'block';
        document.getElementById('otpremaSummaryText').textContent = 
            'Izabrano: ' + count + ' otkupa | ' + kg.toLocaleString('sr') + ' kg';
    } else {
        div.style.display = 'none';
    }
}

async function confirmOtprema() {
    if (!otpremaVozacID) { showToast('Nema vozača', 'error'); return; }
    
    let count = 0;
    for (let i = 0; i < otpremaUnassigned.length; i++) {
        const chk = document.getElementById('otpChk' + i);
        if (chk && chk.checked) {
            otpremaUnassigned[i].vozacID = otpremaVozacID;
            otpremaUnassigned[i].syncStatus = 'pending';
            await dbPut(db, CONFIG.STORE_NAME, otpremaUnassigned[i]);
            count++;
        }
    }
    
    if (count === 0) { showToast('Izaberite bar jedan otkup', 'error'); return; }
    
    showToast(count + ' otkupa dodeljeno vozaču', 'success');
    cancelOtprema();
    loadOtpremaOverview();
    if (navigator.onLine) syncQueue();
}

function cancelOtprema() {
    otpremaVozacID = '';
    otpremaUnassigned = [];
    document.getElementById('otpremaAssignView').style.display = 'none';
    document.getElementById('otpremaMainView').style.display = 'block';
    loadOtpremaOverview();
}

async function loadOtpremaOverview() {
    const today = new Date().toISOString().split('T')[0];
    
    const local = await dbGetAll(db, CONFIG.STORE_NAME);
    
    let server = [];
    try {
        const json = await apiFetch('action=getOtkupi&otkupacID=' + encodeURIComponent(CONFIG.OTKUPAC_ID));
        if (json && json.success && json.records) server = json.records.map(r => ({
            clientRecordID: r.ClientRecordID || '', datum: fmtDate(r.Datum),
            kooperantID: r.KooperantID || '', kooperantName: r.KooperantName || r.KooperantID || '',
            vrstaVoca: r.VrstaVoca || '', sortaVoca: r.SortaVoca || '',
            klasa: r.Klasa || 'I', kolicina: parseFloat(r.Kolicina) || 0,
            cena: parseFloat(r.Cena) || 0, kolAmbalaze: parseInt(r.KolAmbalaze) || 0,
            parcelaID: r.ParcelaID || '', vozacID: r.VozacID || r.VozaciID || '',
            napomena: r.Napomena || '', syncStatus: 'synced'
        }));
    } catch (e) {}
    
    const serverIDs = new Set(server.map(r => r.clientRecordID));
    const all = [...server, ...local.filter(r => r.syncStatus === 'pending' && !serverIDs.has(r.clientRecordID))];
    const unassigned = all.filter(r => !r.vozacID);
    unassigned.sort((a, b) => (b.datum || '').localeCompare(a.datum || ''));
    const assigned = all.filter(r => r.vozacID);
    
    const uList = document.getElementById('otpremaUnassignedList');
    if (unassigned.length === 0) {
        uList.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:12px;font-size:13px;">Svi otkupi su raspoređeni</p>';
    } else {
        uList.innerHTML = unassigned.map(r => {
            const vr = ((r.kolicina || 0) * (r.cena || 0)).toLocaleString('sr');
            return `<div class="queue-item" style="border-left-color:var(--warning);">
                <div class="qi-header"><span class="qi-koop">${r.kooperantName}</span><span class="qi-time">${r.datum}</span></div>
                <div class="qi-detail">${r.vrstaVoca} ${r.sortaVoca || ''} ${r.klasa || 'I'} | ${r.kolicina} kg | ${vr} RSD</div>
            </div>`;
        }).join('');
    }
    
    const aList = document.getElementById('otpremaAssignedList');
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
                    <strong style="color:var(--primary);">🚛 ${vozID}</strong>
                    <span style="font-weight:600;">${g.kg.toLocaleString('sr')} kg | ${g.items.length} otk.</span>
                </div>
                ${g.items.map(r => `<div style="padding:3px 0;font-size:12px;border-top:1px solid #eee;">${r.kooperantName} | ${r.vrstaVoca} ${r.klasa || ''} | ${r.kolicina} kg</div>`).join('')}
            </div>`).join('');
    }
}

// ============================================================
// KOOPERANT: AGROMERE
// ============================================================
// ============================================================
// DIGITALNI AGRONOM — Kooperant Agromere Tab
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
    lager: [],        // artikli na lageru ovog kooperanta
    opremaList: [],   // oprema ovog kooperanta
    geoWatchId: null
};

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
    // Reset state
    agroResetState();

    // Load lager (iz stammdaten.magacinkoop)
    agroLoadLager();

    // Load oprema (sa servera + lokalna)
    await agroLoadOprema();

    // Populate parcele dropdown
    agroPopulateParcele();

    // Start GPS
    agroStartGeo();

    // Load istorija
    agroLoadIstorija();

    // Show step 1
    document.getElementById('agroStep1').style.display = 'block';
    document.getElementById('agroStep2').style.display = 'none';
}

function agroResetState() {
    if (agroState.timerInterval) clearInterval(agroState.timerInterval);
    if (agroState.geoWatchId) navigator.geolocation.clearWatch(agroState.geoWatchId);

    agroState = {
        parcelaID: '', parcelaData: null, mera: '', artikalID: '', artikalData: null,
        kolicina: 0, dozaPreporucena: 0, opremaTraktor: '', opremaPrskalica: '',
        opremaOstalo: '', napomena: '', timerStart: null, timerInterval: null,
        timerResult: null, geoStart: null, geoEnd: null, geoAutoDetect: false,
        meteoOverride: false, meteoSnapshot: null, karencaDana: 0,
        lager: agroState.lager || [], opremaList: agroState.opremaList || [],
        geoWatchId: null
    };

    // Reset UI
    document.querySelectorAll('.agro-mera-btn').forEach(b => b.classList.remove('selected'));
    document.getElementById('agroPreparatSection').classList.remove('visible');
    document.getElementById('agroPreporuka').classList.remove('visible');
    document.getElementById('agroKarencaWarn').classList.remove('visible');
    document.getElementById('agroMeteoWarn').classList.remove('visible');
    document.getElementById('agroTimerPanel').style.display = 'none';
    document.getElementById('agroBtnStart').style.display = 'none';
    document.getElementById('agroBtnStop').style.display = 'none';
    document.getElementById('agroTimerSticky').classList.remove('active');
    const karInfo = document.getElementById('agroKarencaInfo');
    if (karInfo) karInfo.style.display = 'none';
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
// OPREMA — sa servera + lokalna
// ============================================================
async function agroLoadOprema() {
    agroState.opremaList = [];
    try {
        const json = await apiFetch('action=getOprema&kooperantID=' + encodeURIComponent(CONFIG.ENTITY_ID));
        if (json && json.success && json.records) {
            agroState.opremaList = json.records.map(r => ({
                naziv: r.Naziv || '', tip: r.Tip || ''
            }));
        }
    } catch(e) {}

    agroPopulateOpremaDropdowns();
}

function agroPopulateOpremaDropdowns() {
    const tSel = document.getElementById('agroTraktor');
    const pSel = document.getElementById('agroPrskalica');
    if (!tSel || !pSel) return;

    tSel.innerHTML = '<option value="">-- Bez traktora --</option>';
    pSel.innerHTML = '<option value="">-- Bez prskalice --</option>';

    // Kooperantova oprema (sa servera)
    const traktori = agroState.opremaList.filter(o => o.tip === 'Traktor');
    const prskalice = agroState.opremaList.filter(o => o.tip === 'Prskalica' || o.tip === 'Atomizer');

    traktori.forEach(o => {
        const op = document.createElement('option'); op.value = o.naziv; op.textContent = o.naziv;
        tSel.appendChild(op);
    });

    prskalice.forEach(o => {
        const op = document.createElement('option'); op.value = o.naziv; op.textContent = o.naziv;
        pSel.appendChild(op);
    });

    // Predlozi (samo oni koje kooperant još nema)
    const tNames = new Set(traktori.map(o => o.naziv));
    const pNames = new Set(prskalice.map(o => o.naziv));

    if (traktori.length === 0) {
        const og = document.createElement('optgroup'); og.label = '— Česti modeli —';
        OPREMA_PREDLOZI.Traktor.forEach(n => {
            if (!tNames.has(n)) { const op = document.createElement('option'); op.value = n; op.textContent = n; og.appendChild(op); }
        });
        if (og.children.length > 0) tSel.appendChild(og);
    }

    if (prskalice.length === 0) {
        const og = document.createElement('optgroup'); og.label = '— Česti modeli —';
        OPREMA_PREDLOZI.Prskalica.forEach(n => {
            if (!pNames.has(n)) { const op = document.createElement('option'); op.value = n; op.textContent = n; og.appendChild(op); }
        });
        if (og.children.length > 0) pSel.appendChild(og);
    }
}

async function agroSaveNovaOprema(tip, naziv) {
    if (!naziv || !naziv.trim()) return;
    naziv = naziv.trim();

    // Sačuvaj na server
    try {
        await fetch(CONFIG.API_URL, {
            method: 'POST', headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify({
                action: 'syncOprema', token: CONFIG.TOKEN,
                kooperantID: CONFIG.ENTITY_ID,
                records: [{ clientRecordID: crypto.randomUUID(), naziv: naziv, tip: tip }]
            })
        });
    } catch(e) {}

    // Dodaj u lokalni niz i osveži dropdown
    agroState.opremaList.push({ naziv: naziv, tip: tip });
    agroPopulateOpremaDropdowns();

    // Auto-select
    if (tip === 'Traktor') document.getElementById('agroTraktor').value = naziv;
    if (tip === 'Prskalica') document.getElementById('agroPrskalica').value = naziv;

    showToast('Oprema sačuvana: ' + naziv, 'success');
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
    const strip = document.getElementById('agroMeteoStrip');
    strip.style.display = 'flex';

    // Iz meteoCache ako postoji
    let data = null;
    if (meteoCache[parcelaID] && (Date.now() - meteoCache[parcelaID]._ts < 3600000)) {
        data = meteoCache[parcelaID];
    } else {
        try {
            const url = CONFIG.API_URL + '?action=getParcelMeteo&parcelaId=' + encodeURIComponent(parcelaID);
            const resp = await fetch(url);
            const json = await resp.json();
            if (json && json.success) { json._ts = Date.now(); meteoCache[parcelaID] = json; data = json; }
        } catch(e) {}
    }

    if (!data || !data.current) {
        strip.innerHTML = '<span>Nema meteo podataka</span>';
        agroState.meteoSnapshot = null;
        return;
    }

    const c = data.current;
    const temp = (c.temperature || c.Temp || 0);
    const wind = (c.windSpeed || c.Wind || 0);
    const hum = (c.humidity || c.Humidity || 0);
    const precip = (c.precipitation || c.Precip || 0);

    agroState.meteoSnapshot = { temp: temp, wind: wind, humidity: hum };

    document.getElementById('agroMeteoTemp').textContent = '🌡️ ' + temp.toFixed(1) + '°C';
    document.getElementById('agroMeteoWind').textContent = '💨 ' + wind.toFixed(0) + ' km/h';
    document.getElementById('agroMeteoHumidity').textContent = '💧 ' + hum + '%';
    document.getElementById('agroMeteoPrecip').textContent = precip > 0 ? '🌧️ ' + precip.toFixed(1) + 'mm' : '☀️ Suvo';
}

// ============================================================
// KARENCA CHECK — za izabranu parcelu
// ============================================================
async function checkAgroKarenca(parcelaID) {
    const warn = document.getElementById('agroKarencaWarn');
    const berbaBtn = document.getElementById('agroBerbaBtn');
    warn.classList.remove('visible');
    berbaBtn.classList.remove('disabled');

    // Čitaj tretmane sa servera
    let tretmani = [];
    try {
        const json = await apiFetch('action=getTretmani&kooperantID=' + encodeURIComponent(CONFIG.ENTITY_ID));
        if (json && json.success) tretmani = json.records || [];
    } catch(e) {}

    // Dodaj lokalne pending
    try {
        const local = await dbGetAll(db, 'tretmani');
        tretmani = [...tretmani, ...local.filter(r => r.syncStatus === 'pending')];
    } catch(e) {}

    // Nađi poslednji tretman sa karencom za ovu parcelu
    const parcelTretmani = tretmani.filter(t =>
        t.ParcelaID === parcelaID && parseInt(t.KarencaDana) > 0
    );

    if (parcelTretmani.length === 0) return;

    // Sortiraj po datumu desc
    parcelTretmani.sort((a, b) => (b.Datum || b.datum || '').localeCompare(a.Datum || a.datum || ''));
    const last = parcelTretmani[0];
    const datum = last.Datum || last.datum;
    const karDana = parseInt(last.KarencaDana || last.karencaDana) || 0;
    const prepNaziv = last.ArtikalNaziv || last.artikalNaziv || '?';

    const tretmanDate = new Date(datum);
    const berbeDate = new Date(tretmanDate.getTime() + karDana * 24 * 60 * 60 * 1000);
    const today = new Date();
    today.setHours(0,0,0,0);

    if (berbeDate > today) {
        const ostalo = Math.ceil((berbeDate - today) / (24 * 60 * 60 * 1000));
        warn.classList.add('visible');
        document.getElementById('agroKarencaText').innerHTML =
            '<strong>' + prepNaziv + '</strong> — tretman ' + datum +
            '<br>Berba dozvoljena: <strong>' + berbeDate.toLocaleDateString('sr') + '</strong> (još ' + ostalo + ' dana)';
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
            '⏱️ Karenca: <strong>' + art.karencaDana + ' dana</strong> — berba dozvoljena od ' +
            berbeDate.toLocaleDateString('sr');
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
        '<strong>' + finalQty.toLocaleString('sr') + (art.pakovanje > 0 ? ' pak.' : ' ' + art.jm) + '</strong> — ' + art.naziv;
    document.getElementById('agroPreporukaDetail').innerHTML =
        art.dozaPoHa + ' ' + art.jm + '/ha × ' + ha.toFixed(2) + ' ha = ' +
        rawQty.toLocaleString('sr', {maximumFractionDigits:2}) + ' ' + art.jm +
        (pakInfo ? '<br>' + pakInfo : '') + lagerWarn;

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
function agroStartGeo() {
    if (!navigator.geolocation) return;

    agroState.geoWatchId = navigator.geolocation.watchPosition(
        pos => {
            const lat = pos.coords.latitude;
            const lng = pos.coords.longitude;
            agroState.geoStart = { lat, lng };
            agroCheckParcelaProximity(lat, lng);
        },
        () => { /* silent — ručni fallback */ },
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
        banner.innerHTML = '📍 Detektovana parcela: <strong>' + (bestMatch.KatBroj || bestMatch.ParcelaID) + '</strong> ' +
            (bestMatch.Kultura || '') + ' (' + ha.toFixed(2) + ' ha)';
        sel.value = bestMatch.ParcelaID;
        agroState.geoAutoDetect = true;
        onAgroParcelaChange();
    } else if (bestDist <= 200) {
        banner.className = 'agro-geo-banner nearby';
        banner.innerHTML = '📍 Blizu parcele: <strong>' + (bestMatch.KatBroj || bestMatch.ParcelaID) + '</strong> (' +
            Math.round(bestDist) + 'm) — <a href="#" onclick="document.getElementById(\'agroParcelaSel\').value=\'' +
            bestMatch.ParcelaID + '\';onAgroParcelaChange();return false;" style="color:#92400e;font-weight:700;">Izaberi</a>';
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
        '<div class="agro-confirm-row"><span class="label">' + r[0] + '</span><span class="value">' + r[1] + '</span></div>'
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

    let datumBerbeDozvoljeno = '';
    if (agroState.karencaDana > 0) {
        const d = new Date();
        d.setDate(d.getDate() + agroState.karencaDana);
        datumBerbeDozvoljeno = d.toISOString().split('T')[0];
    }

    const record = {
        clientRecordID: crypto.randomUUID(),
        createdAtClient: now.toISOString(),
        kooperantID: CONFIG.ENTITY_ID,
        parcelaID: agroState.parcelaID,
        datum: now.toISOString().split('T')[0],
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
        syncStatus: 'pending'
    };

    // Save to IndexedDB
    await dbPut(db, 'tretmani', record);
    showToast('Tretman sačuvan!', 'success');

    // Sync immediately
    if (navigator.onLine) {
        try {
            const resp = await fetch(CONFIG.API_URL, {
                method: 'POST', headers: { 'Content-Type': 'text/plain' },
                body: JSON.stringify({
                    action: 'syncTretman', token: CONFIG.TOKEN,
                    kooperantID: CONFIG.ENTITY_ID,
                    records: [record]
                })
            });
            const json = await resp.json();
            if (json.success) {
                record.syncStatus = 'synced';
                await dbPut(db, 'tretmani', record);
            }
        } catch(e) {}
    }

    // Reset
    agroResetState();
    agroPopulateParcele();
    agroLoadIstorija();
    document.getElementById('agroStep1').style.display = 'block';
    document.getElementById('agroStep2').style.display = 'none';
}

// ============================================================
// ISTORIJA
// ============================================================
async function agroLoadIstorija() {
    const local = await dbGetAll(db, 'tretmani');
    let server = [];
    try {
        const json = await apiFetch('action=getTretmani&kooperantID=' + encodeURIComponent(CONFIG.ENTITY_ID));
        if (json && json.success && json.records) {
            server = json.records.map(r => ({
                mera: r.Mera || '', datum: fmtDate(r.Datum),
                parcelaID: r.ParcelaID || '', artikalNaziv: r.ArtikalNaziv || '',
                kolicinaUpotrebljena: r.KolicinaUpotrebljena || '',
                jedinicaMere: r.JedinicaMere || '',
                trajanjeMinuta: r.TrajanjeMinuta || '',
                opremaTraktor: r.OpremaTraktor || '',
                karencaDana: r.KarencaDana || '',
                datumBerbeDozvoljeno: r.DatumBerbeDozvoljeno || '',
                syncStatus: 'synced'
            }));
        }
    } catch(e) {}

    const serverIDs = new Set(server.map(r => r.clientRecordID));
    const all = [...server, ...local.filter(r => r.syncStatus === 'pending' && !serverIDs.has(r.clientRecordID))
        .map(r => ({
            mera: r.mera, datum: r.datum, parcelaID: r.parcelaID,
            artikalNaziv: r.artikalNaziv, kolicinaUpotrebljena: r.kolicinaUpotrebljena,
            jedinicaMere: r.jedinicaMere, trajanjeMinuta: r.trajanjeMinuta,
            opremaTraktor: r.opremaTraktor, karencaDana: r.karencaDana,
            datumBerbeDozvoljeno: r.datumBerbeDozvoljeno, syncStatus: 'pending'
        }))
    ].sort((a, b) => (b.datum || '').localeCompare(a.datum || ''));

    const list = document.getElementById('agroTretmaniList');
    const icons = { Zastita: '🛡️', Prihrana: '🌱', Rezidba: '✂️', Zalivanje: '💧', Berba: '🍎' };

    if (all.length === 0) {
        list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema tretmana</p>';
        return;
    }

    list.innerHTML = all.map(r => {
        const bc = r.syncStatus === 'pending' ? 'var(--warning)' : 'var(--success)';
        const min = parseInt(r.trajanjeMinuta) || 0;
        const timeStr = min > 0 ? (Math.floor(min/60) > 0 ? Math.floor(min/60) + 'h ' : '') + (min%60) + 'min' : '';
        return `<div class="queue-item" style="border-left-color:${bc};">
            <div class="qi-header">
                <span class="qi-koop">${icons[r.mera] || ''} ${r.mera}</span>
                <span class="qi-time">${r.datum}</span>
            </div>
            <div class="qi-detail">${r.parcelaID}${r.artikalNaziv ? ' | ' + r.artikalNaziv + ' ' + (r.kolicinaUpotrebljena || '') + ' ' + (r.jedinicaMere || '') : ''}${timeStr ? ' | ⏱️ ' + timeStr : ''}${r.opremaTraktor ? ' | 🚜 ' + r.opremaTraktor : ''}${r.karencaDana ? ' | Karenca ' + r.karencaDana + 'd' : ''}</div>
        </div>`;
    }).join('');
}

// ============================================================
// KOOPERANT: INFO
// ============================================================
async function loadKoopInfo() {
    const config = stammdaten.config || [];
    const gv = k => { const c = config.find(c => c.Parameter === k); return c ? c.Vrednost : '-'; };
    document.getElementById('koopInfoContent').innerHTML = `
        <div style="background:white;border-radius:var(--radius);padding:16px;margin-bottom:12px;box-shadow:0 1px 3px rgba(0,0,0,0.08);">
            <h3 style="color:var(--primary);margin-bottom:12px;">Otkup informacije</h3>
            <table style="width:100%;">
                <tr><td style="padding:8px;color:var(--text-muted);">Status:</td><td style="padding:8px;font-weight:600;">${gv('OtkupAktivan')==='Da'?'🟢 Aktivan':'🔴 Neaktivan'}</td></tr>
                <tr><td style="padding:8px;color:var(--text-muted);">Radno vreme:</td><td style="padding:8px;">${gv('RadnoVremeOd')} - ${gv('RadnoVremeDo')}</td></tr>
                <tr><td style="padding:8px;color:var(--text-muted);">Sezona:</td><td style="padding:8px;">${gv('SezonaOd')} - ${gv('SezonaDo')}</td></tr>
            </table>
        </div>
        <div style="background:white;border-radius:var(--radius);padding:16px;box-shadow:0 1px 3px rgba(0,0,0,0.08);">
            <h3 style="color:var(--primary);margin-bottom:12px;">Aktuelne cene</h3>
            <table style="width:100%;">
                ${config.filter(c => c.Parameter && c.Parameter.startsWith('Cena')).map(c =>
                    '<tr><td style="padding:8px;color:var(--text-muted);">'+c.Parameter.replace('Cena','')+':</td><td style="padding:8px;font-weight:600;">'+c.Vrednost+' RSD/kg</td></tr>').join('')}
            </table>
        </div>`;
}

// ============================================================
// TAB NAVIGATION (non-management)
// ============================================================
function showTab(tabName) {
    qsa('.tab-content').forEach(t => removeClass(t, 'active'));
    qsa('.tab-btn').forEach(b => removeClass(b, 'active'));

    const tabEl = byId('tab-' + tabName);
    if (tabEl) addClass(tabEl, 'active');
    if (event && event.target) addClass(event.target, 'active');

    if (tabName === 'queue') { renderQueueList(); updateStats(); }
    if (tabName === 'pregled') loadOtkupPregled();
    if (tabName === 'otpremnice') loadOtpremaOverview();
    if (tabName === 'kartica') loadKartica();
    if (tabName === 'parcele') loadParcele();
    if (tabName === 'agromere') loadAgronom();
    if (tabName === 'koopinfo') loadKoopInfo();
    if (tabName === 'zbirna') loadVozacData();
    if (tabName === 'transport') loadVozacTransport();
    if (tabName === 'dispecer') loadDispecer();
}

// ============================================================
// HELPERS
// ============================================================
function showToast(msg, type = 'info') {
    const toast = byId('toast');
    setText(toast, msg);
    toast.className = 'toast show ' + type;
    setTimeout(() => { toast.className = 'toast'; }, 3000);
}

// ============================================================
// SERVICE WORKER
// ============================================================
if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('./sw.js').then(reg => {
        setInterval(() => reg.update(), 60000);
        reg.addEventListener('updatefound', () => {
            const nw = reg.installing;
            nw.addEventListener('statechange', () => { if (nw.state === 'activated') showToast('Nova verzija učitana', 'info'); });
        });
    }).catch(err => console.log('SW registration failed:', err));
}
