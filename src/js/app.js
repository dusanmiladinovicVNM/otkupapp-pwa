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
