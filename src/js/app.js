// ============================================================
// INIT
// ============================================================
function getAppRuntime() {
    return window.appRuntime;
}

document.addEventListener('DOMContentLoaded', bootstrapApp);

async function bootstrapApp() {
    const runtime = getAppRuntime();
    
    if (runtime.initStarted) return;
    runtime.initStarted = true;

    try {
        if (!hasValidSession()) {
            hideLoader();
            showLoginScreen();
            return;
        }

        db = await openDB();

        await loadStammdatenFromCache();

        applyRoleVisibility();
        applyHeaderInfo();
        bindAppShellEvents();
        setDefaultDates();

        await bootstrapRole();

        if (typeof initRoleNavEngine === 'function') {
            initRoleNavEngine();
        }

        updateSyncBadge();
        bindConnectivityEvents();
        startBackgroundSync();

        runtime.stammdatenReady = true;
        runtime.appReady = true;

        refreshStammdatenInBackground();
    } catch (err) {
        console.error('bootstrapApp failed:', err);
        showToast('Greška pri pokretanju aplikacije', 'error');
    } finally {
        // UVEK sakrij loader — čak i ako boot pukne
        hideLoader();
    }
}

function hideLoader() {
    const loader = document.getElementById('appLoader');
    if (loader) loader.style.display = 'none';
}

function hasValidSession() {
    return !!getLs('authToken', '') && !!getSessionEntityID();
}

function getSessionEntityID() {
    return getLs('entityID', '') || getLs('otkupacID', '');
}

function applyHeaderInfo() {
    const el = document.getElementById('headerInfo');
    if (!el) return;
    el.textContent = CONFIG.USER_ROLE + ': ' + CONFIG.ENTITY_NAME;
    
    // Version u footer ili kao data atribut
    document.body.dataset.version = CONFIG.APP_VERSION;
}

function setDefaultDates() {
    const today = new Date().toISOString().split('T')[0];

    const fldPregledOd = document.getElementById('fldPregledOd');
    const fldPregledDo = document.getElementById('fldPregledDo');
    const fldOtpremniceDatum = document.getElementById('fldOtpremniceDatum');
    const mgmtOtkupiOd = document.getElementById('mgmtOtkupiOd');
    const mgmtOtkupiDo = document.getElementById('mgmtOtkupiDo');

    if (fldPregledOd && !fldPregledOd.value) fldPregledOd.value = today;
    if (fldPregledDo && !fldPregledDo.value) fldPregledDo.value = today;
    if (fldOtpremniceDatum && !fldOtpremniceDatum.value) fldOtpremniceDatum.value = today;
    if (mgmtOtkupiOd && !mgmtOtkupiOd.value) mgmtOtkupiOd.value = today;
    if (mgmtOtkupiDo && !mgmtOtkupiDo.value) mgmtOtkupiDo.value = today;
}

async function bootstrapRole() {
    if (CONFIG.USER_ROLE === 'Otkupac') {
        if (typeof initOtkupFormUI === 'function') {
            initOtkupFormUI();
        } else {
            if (typeof populateVrstaDropdown === 'function') populateVrstaDropdown();
            if (typeof applyDefaults === 'function') applyDefaults();
        }

        safeCall(() => showTab('otkup'));
        return;
    }

    if (CONFIG.USER_ROLE === 'Kooperant') {
        await guardStammdaten(async () => {
            if (typeof agroPopulateParcele === 'function') {
                await agroPopulateParcele();
            }
        });

        safeCall(() => showTab('home'));
        return;
    }

    if (CONFIG.USER_ROLE === 'Vozac') {
        safeCall(() => showTab('zbirna'));
        return;
    }

    if (CONFIG.USER_ROLE === 'Management') {
        if (typeof populateMgmtStanice === 'function') populateMgmtStanice();

        try {
            if (typeof prefetchMgmtData === 'function') {
                await prefetchMgmtData();
            }
        } catch (err) {
            console.error('prefetchMgmtData failed:', err);
        }

        if (typeof populateMgmtKupciDropdown === 'function') {
            populateMgmtKupciDropdown();
        }

        if (typeof mgmtShellInit === 'function') {
            safeCall(() => mgmtShellInit());
        } else {
            safeCall(() => showTab('dispecer'));
        }
        return;
    }
}

function bindAppShellEvents() {
    const qrProfileModal = document.getElementById('qrProfileModal');
    if (qrProfileModal && !qrProfileModal.dataset.bound) {
        qrProfileModal.addEventListener('click', () => {
            qrProfileModal.style.display = 'none';
        });
        qrProfileModal.dataset.bound = '1';
    }

    window.addEventListener('stammdaten:updated', handleStammdatenUpdated);
}

function bindConnectivityEvents() {
    if (window.__appConnectivityBound) return;
    window.__appConnectivityBound = true;

    window.addEventListener('online', async () => {
        updateSyncBadge();
        await syncQueueSafe();
        refreshStammdatenInBackground();
    });

    window.addEventListener('offline', () => {
        updateSyncBadge();
    });
}

function startBackgroundSync() {
    const runtime = getAppRuntime();
    
    if (runtime.syncIntervalId) {
        clearInterval(runtime.syncIntervalId);
    }

    runtime.syncIntervalId = setInterval(() => {
        if (!navigator.onLine) return;
        if (CONFIG.USER_ROLE === 'Management') return;
        syncQueueSafe();
    }, 60000);
}

// ============================================================
// STAMMDATEN
// ============================================================
async function loadStammdatenFromCache() {
    await safeAsync(async () => {
        const cached = await dbGetAll(db, CONFIG.STAMM_STORE);
        const obj = (cached || []).find(c => c.key === 'all');

        if (obj && obj.data) {
            stammdaten = normalizeStammdaten(obj.data);
        } else {
            stammdaten = normalizeStammdaten(null);
        }
    }, 'Greška pri čitanju lokalnih šifarnika');
}

async function refreshStammdatenInBackground() {
    const runtime = getAppRuntime();
    if (!navigator.onLine) return;
    if (runtime.stammdatenRefreshInFlight) return;

    runtime.stammdatenRefreshInFlight = true;

    try {
        await safeAsync(async () => {
            const result = await apiFetchSafe('action=getStammdaten');

            if (!result.ok || !(result.data && result.data.success && result.data.data)) {
                if (result.error) {
                    console.error('getStammdaten failed:', result.error, result);
                }
                return;
            }

            const nextData = normalizeStammdaten(result.data.data);
            stammdaten = nextData;

            await dbPut(db, CONFIG.STAMM_STORE, {
                key: 'all',
                data: nextData,
                updatedAt: new Date().toISOString()
            });

            window.dispatchEvent(new CustomEvent('stammdaten:updated', {
                detail: { source: 'network' }
            }));
        }, 'Greška pri učitavanju šifarnika');
    } finally {
        runtime.stammdatenRefreshInFlight = false;
    }
}

function normalizeStammdaten(data) {
    var src = data || {};
    var known = {
        kooperanti: [],
        kulture: [],
        config: [],
        parcele: [],
        stanice: [],
        kupci: [],
        vozaci: [],
        artikli: [],
        magacinkoop: [],
        meteoLatest: [],
        kartice: []
    };
    var result = Object.assign({}, src);
    Object.keys(known).forEach(function(k) {
        result[k] = Array.isArray(src[k]) ? src[k] : known[k];
    });

    if (!result.meteoLatest || !result.meteoLatest.length) {
        result.meteoLatest = Array.isArray(src.meteolatest) ? src.meteolatest : [];
    }

    return result;
}

function hasStammdaten() {
    return !!(
        stammdaten &&
        typeof stammdaten === 'object' &&
        Array.isArray(stammdaten.kooperanti) &&
        Array.isArray(stammdaten.parcele)
    );
}

async function guardStammdaten(fn) {
    if (!hasStammdaten()) {
        showToast('Šifarnici još nisu spremni', 'info');
        return;
    }

    try {
        return await fn();
    } catch (err) {
        console.error('guardStammdaten failed:', err);
        showToast('Greška u radu sa šifarnicima', 'error');
    }
}

function handleStammdatenUpdated() {
    try {
        // Invalidate caches koji zavise od stammdaten
        if (typeof invalidateKarticaCache === 'function') {
            invalidateKarticaCache();
        }

        if (typeof invalidateTretmaniCache === 'function') {
            invalidateTretmaniCache();
        }

        if (typeof invalidateOtpremaCache === 'function') {
            invalidateOtpremaCache();
        }

        if (typeof invalidateKpCache === 'function') { 
            invalidateKpCache();
        }
        
        // Repopulate dropdowns per role
        if (CONFIG.USER_ROLE === 'Kooperant') {
            if (typeof agroPopulateParcele === 'function') agroPopulateParcele();
        }

        if (CONFIG.USER_ROLE === 'Management') {
            if (typeof populateMgmtStanice === 'function') populateMgmtStanice();
            if (typeof populateMgmtKupciDropdown === 'function') populateMgmtKupciDropdown();
            if (typeof mgmtRenderOverview === 'function') mgmtRenderOverview();
        }

        if (CONFIG.USER_ROLE === 'Otkupac') {
            if (typeof initOtkupFormUI === 'function') {
                initOtkupFormUI({ preserveSelection: true });
            } else if (typeof populateVrstaDropdown === 'function') {
                populateVrstaDropdown();
            }
        }
    } catch (err) {
        console.error('handleStammdatenUpdated failed:', err);
    }
}

// ============================================================
// SYNC
// ============================================================
async function runRoleSync() {
    if (!navigator.onLine) return { ok: false, reason: 'offline' };

    if (CONFIG.USER_ROLE === 'Otkupac') {
        if (typeof syncQueue !== 'function') return { ok: false, reason: 'missing-syncQueue' };
        await syncQueue();
        return { ok: true, role: 'Otkupac' };
    }

    if (CONFIG.USER_ROLE === 'Kooperant') {
        if (typeof syncKooperantNow !== 'function') return { ok: false, reason: 'missing-syncKooperantNow' };
        await syncKooperantNow();
        return { ok: true, role: 'Kooperant' };
    }

    if (CONFIG.USER_ROLE === 'Vozac') {
        if (typeof syncZbirne !== 'function') return { ok: false, reason: 'missing-syncZbirne' };
        await syncZbirne();
        return { ok: true, role: 'Vozac' };
    }

    return { ok: false, reason: 'no-sync-for-role' };
}

async function syncQueueSafe() {
    const runtime = getAppRuntime();
    
    if (!navigator.onLine) return;
    if (CONFIG.USER_ROLE === 'Management') return;
    if (runtime.syncInFlight) return;

    runtime.syncInFlight = true;
    updateSyncBadge();

    try {
        await runRoleSync();
    } catch (err) {
        console.error('runRoleSync failed:', err);
    } finally {
        runtime.syncInFlight = false;
        updateSyncBadge();
    }
}

// ============================================================
// QR SCANNER
// ============================================================
function onQRScanned(text) {
    try {
        const data = JSON.parse(text);
        if (data.id) {
            setKooperant(data.id, data.name || data.id);
            return;
        }
    } catch (e) {}

    if (text.startsWith('KOOP-')) {
        const koop = (stammdaten.kooperanti || []).find(k => k.KooperantID === text);
        setKooperant(text, koop ? (koop.Ime + ' ' + koop.Prezime) : text);
        return;
    }

    showToast('Nepoznat QR kod', 'error');
}

function setKooperant(id, name) {
    const fldKooperantID = document.getElementById('fldKooperantID');
    const koopName = document.getElementById('koopName');
    const koopId = document.getElementById('koopId');
    const koopDisplay = document.getElementById('koopDisplay');

    if (fldKooperantID) fldKooperantID.value = id;
    if (koopName) koopName.textContent = name;
    if (koopId) koopId.textContent = id;
    if (koopDisplay) koopDisplay.classList.add('visible');

    showToast('Kooperant: ' + name, 'success');

    if (typeof populateParcelaDropdown === 'function') {
        populateParcelaDropdown(id);
    }
}

function startVozacQRScan() {
    const readerDiv = document.getElementById('qr-reader-vozac');
    if (!readerDiv) return;

    readerDiv.style.display = 'block';

    const scanner = new Html5Qrcode('qr-reader-vozac');
    scanner.start(
        { facingMode: 'environment' },
        { fps: 10, qrbox: { width: 250, height: 250 } },
        (decodedText) => {
            onVozacQRScanned(decodedText);
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

function onVozacQRScanned(text) {
    try {
        const data = JSON.parse(text);
        if (data.type === 'VOZ' && data.id) {
            setVozac(data.id, data.name || data.id);
            return;
        }
    } catch (e) {}

    if (text.startsWith('VOZ-')) {
        setVozac(text, text);
        return;
    }

    showToast('Nije QR vozača', 'error');
}

function setVozac(id, name) {
    const fldVozacID = document.getElementById('fldVozacID');
    const vozacName = document.getElementById('vozacName');
    const vozacId = document.getElementById('vozacId');
    const vozacDisplay = document.getElementById('vozacDisplay');

    if (fldVozacID) fldVozacID.value = id;
    if (vozacName) vozacName.textContent = name;
    if (vozacId) vozacId.textContent = id;
    if (vozacDisplay) vozacDisplay.classList.add('visible');

    showToast('Vozač: ' + name, 'success');
}

function clearVozac() {
    const fldVozacID = document.getElementById('fldVozacID');
    const vozacDisplay = document.getElementById('vozacDisplay');

    if (fldVozacID) fldVozacID.value = '';
    if (vozacDisplay) vozacDisplay.classList.remove('visible');
}

// ============================================================
// QR PROFILE
// ============================================================
function showQRProfile() {
    const modal = document.getElementById('qrProfileModal');
    const nameEl = document.getElementById('qrProfileName');
    const roleEl = document.getElementById('qrProfileRole');
    const idEl = document.getElementById('qrProfileID');

    if (!modal || !nameEl || !roleEl || !idEl) return;

    nameEl.textContent = CONFIG.ENTITY_NAME;
    roleEl.textContent = CONFIG.USER_ROLE;
    idEl.textContent = CONFIG.ENTITY_ID;
    modal.style.display = 'flex';

    generateQRCode('qrProfileCanvas', JSON.stringify({
        type:
            CONFIG.USER_ROLE === 'Kooperant' ? 'KOOP' :
            CONFIG.USER_ROLE === 'Otkupac' ? 'OTK' :
            CONFIG.USER_ROLE === 'Vozac' ? 'VOZ' : 'MGMT',
        id: CONFIG.ENTITY_ID,
        name: CONFIG.ENTITY_NAME
    }));
}

// ============================================================
// HELPERS
// ============================================================
function safeCall(fn) {
    try {
        return fn();
    } catch (err) {
        console.error('safeCall failed:', err);
    }
}

// ============================================================
// SERVICE WORKER
// ============================================================
if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('./sw.js').then(reg => {
        setInterval(() => reg.update(), 60000);

        reg.addEventListener('updatefound', () => {
            const nw = reg.installing;
            if (!nw) return;

            nw.addEventListener('statechange', () => {
                if (nw.state === 'activated') {
                    showToast('Nova verzija učitana', 'info');
                }
            });
        });
    }).catch(err => {
        console.log('SW registration failed:', err);
    });
}
