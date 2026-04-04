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
