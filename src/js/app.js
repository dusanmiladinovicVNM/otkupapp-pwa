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
// src/js/app.js
document.addEventListener('DOMContentLoaded', bootstrapApp);

async function bootstrapApp() {
    AppState.patch('init', { domReady: true, bootError: null });

    try {
        if (!hasValidSession()) {
            showLoginScreen();
            return;
        }

        const dbInstance = await openDB();
        AppState.set('db', dbInstance);
        AppState.patch('init', { dbReady: true });

        await loadStammdatenFromCache();
        applyShellUI();
        applyRoleVisibilitySafe();

        await bootstrapRole();

        bindGlobalAppEvents();
        startBackgroundJobs();

        // network refresh after shell is stable
        refreshStammdatenInBackground();

        AppState.patch('init', {
            stammdatenReady: true,
            appReady: true
        });
    } catch (err) {
        console.error('bootstrap failed', err);
        AppState.patch('init', { bootError: err });
        showToast('Greška pri pokretanju aplikacije', 'error');
    }
}

function hasValidSession() {
    return !!getLs('authToken', '') && !!getLs('otkupacID', '');
}

function applyShellUI() {
    const headerInfo = document.getElementById('headerInfo');
    if (headerInfo) {
        headerInfo.textContent = CONFIG.USER_ROLE + ': ' + CONFIG.ENTITY_NAME;
    }
    updateSyncBadge();
}

function applyRoleVisibilitySafe() {
    applyRoleVisibility();
}

async function bootstrapRole() {
    const today = new Date().toISOString().split('T')[0];

    if (CONFIG.USER_ROLE === 'Otkupac') {
        populateVrstaDropdown();
        applyDefaults();

        const od = document.getElementById('fldPregledOd');
        const _do = document.getElementById('fldPregledDo');
        if (od) od.value = today;
        if (_do) _do.value = today;
    }

    if (CONFIG.USER_ROLE === 'Kooperant') {
        await guardStammdaten(agroPopulateParcele);
    }

    if (CONFIG.USER_ROLE === 'Management') {
        populateMgmtStanice();

        const od = document.getElementById('mgmtOtkupiOd');
        const _do = document.getElementById('mgmtOtkupiDo');
        if (od) od.value = today;
        if (_do) _do.value = today;

        await prefetchMgmtData();
        populateMgmtKupciDropdown();
        showTab('dispecer');
    }
}

function bindGlobalAppEvents() {
    window.addEventListener('online', handleOnline);
    window.addEventListener('offline', handleOffline);
}

function startBackgroundJobs() {
    if (window.__syncInterval) clearInterval(window.__syncInterval);

    window.__syncInterval = setInterval(() => {
        if (navigator.onLine && CONFIG.USER_ROLE === 'Otkupac') {
            syncQueueSafe();
        }
    }, 60000);
}

function handleOnline() {
    updateSyncBadge();
    if (CONFIG.USER_ROLE === 'Otkupac') syncQueueSafe();
}

function handleOffline() {
    updateSyncBadge();
}

async function syncQueueSafe() {
    if (AppState.get('sync.inFlight')) return;

    AppState.patch('sync', { inFlight: true });
    try {
        await syncQueue();
        AppState.patch('sync', { lastRunAt: new Date().toISOString() });
    } finally {
        AppState.patch('sync', { inFlight: false });
    }
}


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


    
// ============================================================
// STAMMDATEN
// ============================================================
async function loadStammdaten() {
    await safeAsync(async () => {
        const cached = await dbGetAll(db, CONFIG.STAMM_STORE);
        const obj = cached.find(c => c.key === 'all');
        if (obj) stammdaten = obj.data;
    });

    if (!navigator.onLine) return;

    await safeAsync(async () => {
        const json = await apiFetch('action=getStammdaten');
        if (json && json.success && json.data) {
            stammdaten = json.data;
            await dbPut(db, CONFIG.STAMM_STORE, {
                key: 'all',
                data: stammdaten,
                updatedAt: new Date().toISOString()
            });
        }
    }, 'Greška pri učitavanju šifarnika');
}

async function loadStammdatenFromCache() {
    const db = AppState.get('db');
    if (!db) throw new Error('DB nije inicijalizovan');

    await safeAsync(async () => {
        const cached = await dbGetAll(db, CONFIG.STAMM_STORE);
        const obj = cached.find(c => c.key === 'all');

        if (obj && obj.data) {
            AppState.set('stammdaten', obj.data);
        }
    }, 'Greška pri čitanju lokalnih šifarnika');
}

async function refreshStammdatenInBackground() {
    if (!navigator.onLine) return;

    await safeAsync(async () => {
        const json = await apiFetch('action=getStammdaten');
        if (!(json && json.success && json.data)) return;

        AppState.set('stammdaten', json.data);

        const db = AppState.get('db');
        await dbPut(db, CONFIG.STAMM_STORE, {
            key: 'all',
            data: json.data,
            updatedAt: new Date().toISOString()
        });

        window.dispatchEvent(new CustomEvent('stammdaten:updated', {
            detail: { source: 'network' }
        }));
    }, 'Greška pri osvežavanju šifarnika');
}



function fmtStanica(stanicaID) {
    if (!stanicaID) return '';
    const s = (stammdaten.stanice || []).find(s => s.StanicaID === stanicaID);
    const name = s ? (s.Naziv || s.Mesto || stanicaID) : stanicaID;
    if (name === stanicaID) return stanicaID;
    return name + ' <span style="font-size:11px;color:var(--text-muted);">' + stanicaID + '</span>';
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
