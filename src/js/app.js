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

    const pregledDetailCard = document.querySelector('#pregledDetailModal .danas-detail-card');
    if (pregledDetailCard && !pregledDetailCard.dataset.bound) {
        pregledDetailCard.addEventListener('click', (e) => e.stopPropagation());
        pregledDetailCard.dataset.bound = '1';
    }

    const otpremaDetailCard = document.querySelector('#otpremaDetailModal .otprema-detail-card');
    if (otpremaDetailCard && !otpremaDetailCard.dataset.bound) {
        otpremaDetailCard.addEventListener('click', (e) => e.stopPropagation());
        otpremaDetailCard.dataset.bound = '1';
    }

    const homeQuickActionsCard = document.querySelector('#homeQuickActionsModal .home-quick-sheet');
    if (homeQuickActionsCard && !homeQuickActionsCard.dataset.bound) {
        homeQuickActionsCard.addEventListener('click', (e) => e.stopPropagation());
        homeQuickActionsCard.dataset.bound = '1';
    }
    
    if (!window.__appShellDelegatedBound) {
        window.__appShellDelegatedBound = true;
        document.addEventListener('click', handleAppShellClick);
        document.addEventListener('change', handleAppShellChange);
        document.addEventListener('input', handleAppShellInput);
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

function handleAppShellClick(event) {
    const actionEl = event.target.closest('[data-action]');
    if (actionEl) {
        const action = actionEl.dataset.action;

        if (action === 'show-qr-profile') {
            showQRProfile();
            return;
        }

        if (action === 'logout') {
            doLogout();
            return;
        }

        if (action === 'start-qr-scan') {
            startQRScan();
            return;
        }

        if (action === 'start-vozac-qr-scan') {
            startVozacQRScan();
            return;
        }

        if (action === 'clear-vozac') {
            clearVozac();
            return;
        }

        if (action === 'reset-otkup-form') {
            resetForm();
            return;
        }

        if (action === 'save-otkup') {
            saveOtkup();
            return;
        }

        if (action === 'pregled-filter') {
            setPregledQuickFilter(actionEl.dataset.filter, actionEl);
            return;
        }

        if (action === 'close-pregled-detail') {
            closePregledDetail();
            return;
        }

        if (action === 'open-pregled-otkupni-list') {
            openPregledDetailOtkupniList();
            return;
        }

        if (action === 'start-otprema-vozac-qr-scan') {
            startOtpremaVozacQRScan();
            return;
        }

        if (action === 'toggle-otprema-fallback') {
            toggleOtpremaFallback();
            return;
        }

        if (action === 'apply-otprema-fallback-driver') {
            applyOtpremaFallbackDriver();
            return;
        }

        if (action === 'cancel-otprema-assign') {
            cancelOtpremaAssign();
            return;
        }

        if (action === 'select-all-otprema-today') {
            selectAllOtpremaToday();
            return;
        }

        if (action === 'clear-otprema-selection') {
            clearOtpremaSelection();
            return;
        }

        if (action === 'confirm-otprema-assign') {
            confirmOtpremaAssign();
            return;
        }

        if (action === 'back-to-otprema-root') {
            backToOtpremaRoot();
            return;
        }

        if (action === 'close-otprema-detail') {
            closeOtpremaDetail();
            return;
        }

        if (action === 'sync-otkupac-now') {
            syncOtkupacFromMore();
            return;
        }

        if (action === 'clear-otkupac-signature') {
            clearOtkupacSignature();
            return;
        }

        if (action === 'save-otkupac-signature') {
            saveOtkupacSignature();
            return;
        }

        if (action === 'open-home-quick-actions') {
            openHomeQuickActions();
            return;
        }

        if (action === 'home-show-alerts') {
            showHomeAlerts();
            return;
        }

        if (action === 'home-go-new-rad') {
            goToNewRad();
            return;
        }

        if (action === 'home-go-new-trosak') {
            goToNewTrosak();
            return;
        }

        if (action === 'home-go-scan-racun') {
            goToScanRacun();
            return;
        }

        if (action === 'close-home-quick-actions') {
            closeHomeQuickActions();
            return;
        }

        if (action === 'home-quick-new-rad') {
            closeHomeQuickActions();
            goToNewRad();
            return;
        }

        if (action === 'home-quick-new-trosak') {
            closeHomeQuickActions();
            goToNewTrosak();
            return;
        }

        if (action === 'home-quick-scan-racun') {
            closeHomeQuickActions();
            goToScanRacun();
            return;
        }

        if (action === 'home-quick-kartica') {
            closeHomeQuickActions();
            goToKartica();
            return;
        }

        if (action === 'home-quick-knjiga-polja') {
            closeHomeQuickActions();
            goToKnjigaPolja();
            return;
        }

        if (action === 'toggle-parcele-view') {
            toggleParceleView();
            return;
        }

        if (action === 'show-parcele-section') {
            showParceleSection(actionEl.dataset.section, actionEl);
            return;
        }

        if (action === 'close-parcela-detail') {
            closeParcelaDetail();
            return;
        }

        if (action === 'go-new-rad-from-parcela') {
            goToNewRadFromParcela();
            return;
        }

        if (action === 'go-new-trosak-from-parcela') {
            goToNewTrosakFromParcela();
            return;
        }

        if (action === 'show-parcela-detail-section') {
            showParcelaDetailSection(actionEl.dataset.section, actionEl);
            return;
        }

        if (action === 'radovi-open-new') {
            showRadoviSection('tretmani', document.querySelector('.radovi-subnav-btn'));
            scrollRadoviFormIntoView();
            return;
        }

        if (action === 'show-radovi-section') {
            showRadoviSection(actionEl.dataset.section, actionEl);
            return;
        }

        if (action === 'select-agro-mera') {
            selectAgroMera(actionEl, actionEl.dataset.mera);
            return;
        }

        if (action === 'agro-meteo-override') {
            agroMeteoOverride();
            return;
        }

        if (action === 'agro-primeni-preporuku') {
            agroPrimeniPreporuku();
            return;
        }

        if (action === 'agro-start-rad') {
            agroStartRad();
            return;
        }

        if (action === 'agro-stop-rad') {
            agroStopRad();
            return;
        }

        if (action === 'agro-save-tretman') {
            agroSaveTretman();
            return;
        }

        if (action === 'agro-back-to-step1') {
            agroBackToStep1();
            return;
        }

        if (action === 'radovi-open-tretmani') {
            showRadoviSection('tretmani', document.querySelector('.radovi-subnav-btn'));
            return;
        }

        if (action === 'knjiga-open-new-trosak') {
            showKnjigaSection(
                'troskovi',
                document.querySelector('.knjiga-subnav-btn[data-action="show-knjiga-section"][data-section="troskovi"]')
            );
            scrollKnjigaTrosakFormIntoView();
            return;
        }

        if (action === 'show-knjiga-section') {
            showKnjigaSection(actionEl.dataset.section, actionEl);
            return;
        }

        if (action === 'kp-save-trosak') {
            kpSaveTrosak();
            return;
        }

        if (action === 'start-fiskalni-scan') {
            startFiskalniScan();
            return;
        }

        if (action === 'fiskalni-save-to-lager') {
            fiskalniSaveToLager();
            return;
        }

        if (action === 'fiskalni-cancel') {
            fiskalniCancel();
            return;
        }

        if (action === 'mgmt-dash-period') {
            setMgmtDashboardPeriod(actionEl.dataset.period, actionEl);
            return;
        }

        if (action === 'dp-ok') {
            dpOK();
            return;
        }

        if (action === 'dp-x') {
            dpX();
            return;
        }

        if (action === 'dp-ad') {
            dpAD();
            return;
        }

        if (action === 'load-dispecer') {
            loadDispecer();
            return;
        }

        if (action === 'mgmt-otkup-sub') {
            showMgmtOtkupSub(actionEl.dataset.sub, actionEl);
            return;
        }

        if (action === 'mgmt-partner-segment') {
            showMgmtPartnerSegment(actionEl.dataset.segment, actionEl);
            return;
        }

        if (action === 'mgmt-koop-sub') {
            showMgmtKoopSub(actionEl.dataset.sub, actionEl);
            return;
        }

        if (action === 'mgmt-kup-sub') {
            showMgmtKupSub(actionEl.dataset.sub, actionEl);
            return;
        }

        if (action === 'mgmt-agro-sub') {
            showMgmtAgroSub(actionEl.dataset.sub, actionEl);
            return;
        }

        if (action === 'start-izd-koop-scan') {
            startIzdKoopScan();
            return;
        }

        if (action === 'izd-primeni-preporuku') {
            izdPrimeniPreporuku();
            return;
        }

        if (action === 'start-izd-barcode-scan') {
            startIzdBarcodeScan();
            return;
        }

        if (action === 'izd-dodaj-stavku') {
            izdDodajStavku();
            return;
        }

        if (action === 'izd-zavrsi') {
            izdZavrsi();
            return;
        }

        if (action === 'izd-reset') {
            izdReset();
            return;
        }

        if (action === 'confirm-zbirna') {
            confirmZbirna();
            return;
        }

        if (action === 'cancel-zbirna') {
            cancelZbirna();
            return;
        }

        if (action === 'start-zbirna-creation') {
            startZbirnaCreation();
            return;
        }

        if (action === 'more-fiskalni-racuni') {
            showTab('agromere', findTabBtnByTabName('agromere'));
            setTimeout(() => {
                if (typeof startFiskalniScan === 'function') startFiskalniScan();
            }, 250);
            return;
        }

        if (action === 'sync-kooperant-from-more') {
            syncKooperantFromMore();
            return;
        }

        if (action === 'role-nav-tab') {
            showRoleNavTab(actionEl.dataset.tab, actionEl);
            return;
        }

        if (action === 'open-parcela-detail') {
            openParcelaDetail(actionEl.dataset.parcelaId, actionEl.dataset.source || '');
            return;
        }

        if (action === 'focus-parcel') {
            focusParcel(actionEl.dataset.parcelaId);
            return;
        }

        if (action === 'toggle-expert-panel') {
            event.stopPropagation();
            toggleExpertPanel(actionEl.dataset.parcelaId);
            return;
        }

        if (action === 'agro-select-nearby-parcela') {
            const parcelaSel = document.getElementById('agroParcelaSel');
            if (parcelaSel) {
                parcelaSel.value = actionEl.dataset.parcelaId || '';
            }
            onAgroParcelaChange();
            return;
        }

        if (action === 'select-trosak-kat') {
            selectTrosakKat(actionEl, actionEl.dataset.kat || '');
            return;
        }

        if (action === 'toggle-kp-otkupi-group') {
            const index = parseInt(actionEl.dataset.index || '', 10);
            if (!isNaN(index)) {
                toggleKpOtkupiGroup(index);
            }
            return;
        }

        if (action === 'pregled-alert-click') {
            const index = parseInt(actionEl.dataset.index || '', 10);
            if (!isNaN(index)) {
                onPregledAlertClick(index);
            }
            return;
        }
    }

    const routeEl = event.target.closest('[data-route]');
    if (routeEl) {
        const routeType = routeEl.dataset.route;

        if (routeType === 'tab') {
            showTab(routeEl.dataset.tab, routeEl);
            return;
        }

        if (routeType === 'mgmt-root') {
            showMgmtRoot(routeEl.dataset.root, routeEl);
            return;
        }
    }

    if (event.target.id === 'pregledDetailModal') {
        closePregledDetail();
        return;
    }

    if (event.target.id === 'otpremaDetailModal') {
        closeOtpremaDetail();
        return;
    }

    if (event.target.id === 'homeQuickActionsModal') {
        closeHomeQuickActions();
        return;
    }
}

function handleAppShellChange(event) {
    const el = event.target;
    if (!el || !el.id) return;

    if (el.id === 'fldKooperantManual') {
        onManualKooperantChange();
        return;
    }

    if (el.id === 'fldVrsta') {
        onVrstaChange();
        return;
    }

    if (el.id === 'fldPregledOd' || el.id === 'fldPregledDo') {
        onPregledDateChange();
        return;
    }
    
    if (el.id === 'parceleKulturaFilter') {
        applyParceleFilters();
        return;
    }

    if (el.id === 'agroParcelaSel') {
        onAgroParcelaChange();
        return;
    }

    if (el.id === 'agroPreparatSel') {
        onAgroPreparatChange();
        return;
    }

    if (el.id === 'agroTraktor') {
        refreshRadoviOpremaInfo();
        return;
    }

    if (el.id === 'agroPrskalica') {
        refreshRadoviOpremaInfo();
        return;
    }

    if (el.id === 'kpParcelaSel' || el.id === 'kpSezona') {
        kpLoadBilans();
        return;
    }

    if (el.id === 'mgmtOtkupiStanica' || el.id === 'mgmtOtkupiOd' || el.id === 'mgmtOtkupiDo') {
        loadMgmtOtkupi();
        return;
    }

    if (el.id === 'mgmtPregledStanica') {
        renderMgmtKoopPregled();
        return;
    }

    if (el.id === 'mgmtStanica') {
        onMgmtStanicaChange();
        return;
    }

    if (el.id === 'mgmtKooperant') {
        onMgmtKooperantChange();
        return;
    }

    if (el.id === 'mgmtFaktureKupac') {
        loadMgmtFakture();
        return;
    }

    if (el.id === 'izdKooperant') {
        onIzdKooperantChange();
        return;
    }

    if (el.id === 'agroTraktorNovi') {
        agroSaveNovaOprema('Traktor', el.value);
        return;
    }

    if (el.id === 'agroPrskalicaNovi') {
        agroSaveNovaOprema('Prskalica', el.value);
        return;
    }
}

function handleAppShellInput(event) {
    const el = event.target;
    if (!el || !el.id) return;

    if (el.id === 'parceleSearch') {
        applyParceleFilters();
        return;
    }

    if (el.id === 'agroOpremaOstalo') {
        refreshRadoviOpremaInfo();
        return;
    }
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

    try {
        await runRoleSync();
    } catch (err) {
        console.error('runRoleSync failed:', err);
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
