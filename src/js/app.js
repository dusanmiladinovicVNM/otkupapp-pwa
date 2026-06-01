// ============================================================
// INIT
// ============================================================
function getAppRuntime() {
    return window.appRuntime;
}

// ── Bootstrap error boundary ──────────────────────────────
// Inline stilovi su namerni — radi i ako CSS fajlovi ne učitaju.
function showBootError(err) {
    const loader = document.getElementById('appLoader');
    if (loader) loader.style.display = 'none';

    const existing = document.getElementById('bootErrorScreen');
    if (existing) return;

    const msg = (err && err.message) ? err.message : String(err || 'Nepoznata greška');

    const screen = document.createElement('div');
    screen.id = 'bootErrorScreen';
    screen.style.cssText = 'position:fixed;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:32px;background:#F7F4EE;font-family:system-ui,sans-serif;z-index:9999;text-align:center;';
    screen.innerHTML =
        '<div style="font-size:32px;margin-bottom:16px;">⚠️</div>' +
        '<div style="font-size:17px;font-weight:700;color:#1E2D14;margin-bottom:8px;">Aplikacija nije mogla da se pokrene</div>' +
        '<div style="font-size:13px;color:#7A856E;margin-bottom:24px;max-width:320px;">' + msg + '</div>' +
        '<button id="bootErrReload" style="padding:12px 28px;background:#5EA135;color:white;border:none;border-radius:10px;font-size:15px;font-weight:600;cursor:pointer;margin-bottom:12px;width:100%;max-width:280px;">Pokušaj ponovo</button>' +
        '<button id="bootErrLogout" style="padding:12px 28px;background:none;color:#7A856E;border:1px solid #d0ccc4;border-radius:10px;font-size:14px;cursor:pointer;width:100%;max-width:280px;">Odjavi se i resetuj</button>';

    document.body.appendChild(screen);

    document.getElementById('bootErrReload').onclick = function () { window.location.reload(); };
    document.getElementById('bootErrLogout').onclick = function () {
        try { if (typeof doLogout === 'function') doLogout(); } catch (_) {}
        localStorage.clear();
        window.location.reload();
    };
}

// Global net — hvata neočekivane greške tokom boot-a (pre appReady)
window.addEventListener('error', function (e) {
    if (getAppRuntime().appReady) return;
    showBootError(e.error || e.message);
});
window.addEventListener('unhandledrejection', function (e) {
    if (getAppRuntime().appReady) return;
    showBootError(e.reason);
});
// ─────────────────────────────────────────────────────────

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

        await recoverStaleSyncingForCurrentRole('bootstrap');

        await loadStammdatenFromCache();

        applyRoleVisibility();
        applyHeaderInfo();
        applyAllTabEyebrows();
        applyOtkupHeaderDate();
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

        if (typeof window.reportClientError === 'function') {
            window.reportClientError(err, {
                source: 'app',
                errorAction: 'bootstrapApp'
            });
        }

        showBootError(err);
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

function applyAllTabEyebrows() {
    const station = ((window.CONFIG && CONFIG.ENTITY_NAME) || '').toUpperCase();
    const setEyebrow = (id, baseLabel) => {
        const el = document.getElementById(id);
        if (!el) return;
        el.textContent = station ? baseLabel + ' · ' + station : baseLabel;
    };

    setEyebrow('otkupRoleEyebrow', 'OTKUP');
    setEyebrow('pregledRoleEyebrow', 'PREGLED');
    setEyebrow('viseRoleEyebrow', 'PODEŠAVANJA');
    // Otprema već ima svoj per-render helper u otpremnice.js
}

function applyOtkupHeaderDate() {
    const el = document.getElementById('otkupHeaderDate');
    if (!el) return;

    try {
        const fmt = new Intl.DateTimeFormat('sr-Latn-RS', {
            weekday: 'long',
            day: 'numeric',
            month: 'long',
            year: 'numeric'
        });
        // npr. "petak, 25. maj 2026."
        el.textContent = fmt.format(new Date());
    } catch (_) {
        // Fallback ako sr-Latn-RS nije podržan u browser-u
        const d = new Date();
        el.textContent =
            String(d.getDate()).padStart(2, '0') + '.' +
            String(d.getMonth() + 1).padStart(2, '0') + '.' +
            d.getFullYear() + '.';
    }
}

function setDefaultDates() {
    const today = getTodayIsoDate();

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
        if (window.intercomField) {
            window.intercomField.init().catch(err =>
                console.warn('[app] intercom field init skipped:', err)
            );
        }
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
        if (typeof initVozacStatus === 'function') safeCall(() => initVozacStatus());
        safeCall(() => showTab('transport'));
        return;
    }

    if (CONFIG.USER_ROLE === 'Management') {
        if (typeof populateMgmtStanice === 'function') populateMgmtStanice();

        if (typeof prefetchMgmtData === 'function') {
            prefetchMgmtData()
                .then(() => {
                    if (typeof populateMgmtKupciDropdown === 'function') {
                        populateMgmtKupciDropdown();
                    }

                    if (typeof mgmtRenderOverview === 'function') {
                        mgmtRenderOverview();
                    }

                    if (
                        window.mgmtShellState &&
                        window.mgmtShellState.activeRoot === 'dashboard' &&
                        typeof mgmtRenderDashboard === 'function'
                    ) {
                        mgmtRenderDashboard();
                    }
                })
                .catch(err => {
                    console.error('prefetchMgmtData background failed:', err);
                });
        }

        if (typeof mgmtShellInit === 'function') {
            safeCall(() => mgmtShellInit());
        } else {
            safeCall(() => showTab('dispecer'));
        }

        if (window.intercomMonitor && stammdaten && Array.isArray(stammdaten.stanice)) {
            const activeStations = _intercomActiveStations(stammdaten.stanice);
            window.intercomMonitor.init(activeStations).catch(err =>
                console.warn('[app] intercom monitor init skipped:', err)
            );
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
        await syncQueueSafe('online');
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

        if (action === 'agro-novo-izdavanje') {
            if (typeof showMgmtAgroView === 'function') showMgmtAgroView('izdavanje');
            return;
        }

        if (action === 'agro-back-main') {
            if (typeof showMgmtAgroView === 'function') showMgmtAgroView('main');
            return;
        }

        if (action === 'agro-back-izdavanje') {
            if (typeof showMgmtAgroView === 'function') showMgmtAgroView('izdavanje');
            return;
        }

        if (action === 'sig-clear') {
            const target = actionEl.dataset.target;
            if (target && typeof clearSignature === 'function') clearSignature(target);
            return;
        }

        if (action === 'otp-zatvori-i-sacuvaj') {
            if (typeof otpZatvoriISacuvaj === 'function') otpZatvoriISacuvaj();
            return;
        }

        if (action === 'otp-stampa') {
            // stub — štampa u pripremi
            if (typeof showToast === 'function') showToast('Štampa je u pripremi.');
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

        if (action === 'vozac-fab-nova-zbirna') {
            showTab('zbirna');
            // loadVozacData is async — wait for it to populate vozacOtkupi before opening create view
            setTimeout(() => startZbirnaCreation(), 400);
            return;
        }

        if (action === 'vozac-set-status') {
            if (typeof setVozacStatus === 'function') setVozacStatus(actionEl.dataset.status);
            return;
        }

        if (action === 'vozac-scan-kupac') {
            showToast('QR skener nije još dostupan — izaberite kupca iz liste', 'info');
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
        if (action === 'intercom-request-permission') {
            if (window.intercomField) window.intercomField.handle(action);
            return;
        }
        if (action === 'intercom-listen') {
            if (window.intercomMonitor) window.intercomMonitor.handle(action, actionEl.dataset);
            return;
        }
        if (action === 'intercom-stop') {
            if (window.intercomMonitor) window.intercomMonitor.handle(action, actionEl.dataset);
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
        syncQueueSafe('interval');
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
    var result = (typeof structuredClone === 'function') ? structuredClone(src) : JSON.parse(JSON.stringify(src));
    Object.keys(known).forEach(function(k) {
        result[k] = Array.isArray(result[k]) ? result[k] : known[k];
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

        if (typeof window.reportClientError === 'function') {
            window.reportClientError(err, {
                source: 'app',
                errorAction: 'guardStammdaten'
            });
        }
        
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

function normalizeRoleSyncResult(result, roleName) {
    const role = roleName || (CONFIG && CONFIG.USER_ROLE) || '';

    if (!result || typeof result !== 'object') {
        return {
            ok: true,
            role,
            synced: 0,
            failed: 0,
            results: [],
            reason: 'completed',
            code: '',
            partial: false
        };
    }

    const results = Array.isArray(result.results) ? result.results : [];

    let synced = Number.isFinite(result.synced)
        ? result.synced
        : results.filter(r => r && (r.success === true || r.status === 'synced' || r.ok === true)).length;

    let failed = Number.isFinite(result.failed)
        ? result.failed
        : results.filter(r => r && (r.success === false || r.ok === false)).length;

    if (!Number.isFinite(synced)) synced = 0;
    if (!Number.isFinite(failed)) failed = 0;

    let ok;
    if (typeof result.ok === 'boolean') {
        ok = result.ok;
    } else if (typeof result.success === 'boolean') {
        ok = result.success;
    } else {
        ok = failed === 0;
    }

    return {
        ok,
        role: result.role || role,
        synced,
        failed,
        results,
        reason: result.reason || result.error || '',
        code: result.code || '',
        partial: !!result.partial || (synced > 0 && failed > 0)
    };
}

async function requestRoleSync(reason) {
    const role = CONFIG.USER_ROLE || '';
    const syncReason = reason || 'manual';

    if (role === 'Management') {
        return normalizeRoleSyncResult({
            ok: false,
            reason: 'no-sync-for-role',
            code: 'NO_SYNC_FOR_ROLE'
        }, role);
    }

    if (role === 'Otkupac') {
        if (typeof requestOtkupSync === 'function') {
            return normalizeRoleSyncResult(await requestOtkupSync(syncReason), 'Otkupac');
        }

        return normalizeRoleSyncResult({
            ok: false,
            reason: 'missing-requestOtkupSync',
            code: 'MISSING_SYNC_WRAPPER'
        }, 'Otkupac');
    }

    if (role === 'Kooperant') {
        if (typeof requestKooperantSync === 'function') {
            return normalizeRoleSyncResult(await requestKooperantSync(syncReason), 'Kooperant');
        }

        return normalizeRoleSyncResult({
            ok: false,
            reason: 'missing-requestKooperantSync',
            code: 'MISSING_SYNC_WRAPPER'
        }, 'Kooperant');
    }

    if (role === 'Vozac') {
        // Vozac ostaje kroz role-level gate dok ne uvedemo requestVozacSync u vozac/zbirna.js.
        // Bitno: i dalje ide kroz syncQueueSafe/requestRoleSync, ne direktno iz triggera.
        if (typeof syncZbirne === 'function') {
            return normalizeRoleSyncResult(await syncZbirne(), 'Vozac');
        }

        return normalizeRoleSyncResult({
            ok: false,
            reason: 'missing-syncZbirne',
            code: 'MISSING_SYNC_FUNCTION'
        }, 'Vozac');
    }

    return normalizeRoleSyncResult({
        ok: false,
        reason: 'no-sync-for-role',
        code: 'NO_SYNC_FOR_ROLE'
    }, role);
}

window.requestRoleSync = requestRoleSync;

async function runRoleSync(reason) {
    return requestRoleSync(reason || 'manual');
}

async function syncQueueSafe(reason) {
    const runtime = getAppRuntime();
    const runtimeSync = runtime.sync || (runtime.sync = {});

    const role = CONFIG.USER_ROLE || '';
    const syncReason = reason || 'manual';

    if (!navigator.onLine) {
        return normalizeRoleSyncResult({
            ok: false,
            reason: 'offline',
            code: 'OFFLINE'
        }, role);
    }

    if (role === 'Management') {
        return normalizeRoleSyncResult({
            ok: false,
            reason: 'no-sync-for-role',
            code: 'NO_SYNC_FOR_ROLE'
        }, 'Management');
    }

    // Global role-level guard: svi triggeri ulaze ovde.
    if (runtimeSync.queueInFlight) {
        runtimeSync.queueRequested = true;
        runtimeSync.queueRequestedReason = syncReason;

        return normalizeRoleSyncResult({
            ok: true,
            reason: 'already-running',
            code: 'ALREADY_RUNNING'
        }, role);
    }

    runtimeSync.queueInFlight = true;
    runtimeSync.queueReason = syncReason;

    try {
        if (typeof updateSyncBadge === 'function') {
            try { await updateSyncBadge('syncing'); } catch (_) {}
        }

        let result = await requestRoleSync(syncReason);

        // Ako je manual/online/interval/post-save stigao dok je sync trajao,
        // ne pokrećemo paralelno, nego radimo još jedan serijski pass.
        if (runtimeSync.queueRequested) {
            const nextReason = runtimeSync.queueRequestedReason || 'requested';

            runtimeSync.queueRequested = false;
            runtimeSync.queueRequestedReason = '';
            runtimeSync.queueReason = nextReason;

            result = await requestRoleSync(nextReason);
        }

        return normalizeRoleSyncResult(result, role);
    } catch (err) {
        console.error('syncQueueSafe failed:', err);

        if (typeof window.reportClientError === 'function') {
            window.reportClientError(err, {
                source: 'app',
                errorAction: 'syncQueueSafe',
                reason: syncReason,
                role
            });
        }

        return normalizeRoleSyncResult({
            ok: false,
            reason: (err && err.message) || 'sync-error',
            code: (err && err.name) || 'SYNC_ERROR'
        }, role);
    } finally {
        runtimeSync.queueInFlight = false;
        runtimeSync.queueReason = '';

        if (typeof updateSyncBadge === 'function') {
            try { await updateSyncBadge(); } catch (_) {}
        }
    }
}

async function recoverStaleSyncingForCurrentRole(reason) {
    if (!db || typeof recoverStaleSyncingStores !== 'function') {
        return [];
    }

    const role = CONFIG.USER_ROLE || '';
    const stores = [];

    if (role === 'Otkupac') {
        stores.push(CONFIG.STORE_NAME);
    } else if (role === 'Kooperant') {
        stores.push('tretmani', 'troskovi');
    } else if (role === 'Vozac') {
        stores.push('zbirne');
    } else {
        return [];
    }

    const uniqueStores = Array.from(new Set(stores.filter(Boolean)));

    try {
        const results = await recoverStaleSyncingStores(uniqueStores);

        const recoveredTotal = results.reduce((sum, r) => {
            return sum + (parseInt(r.recovered, 10) || 0);
        }, 0);

        const runtime = getAppRuntime();
        const runtimeSync = runtime.sync || (runtime.sync = {});

        runtimeSync.lastStaleRecoveryReason = reason || 'bootstrap';
        runtimeSync.lastStaleRecoveryAt = new Date().toISOString();
        runtimeSync.lastStaleRecoveryCount = recoveredTotal;
        runtimeSync.lastStaleRecoveryResults = results;

        if (recoveredTotal > 0) {
            console.info('[app] stale syncing recovered:', results);
        }

        return results;
    } catch (err) {
        console.error('recoverStaleSyncingForCurrentRole failed:', err);

        if (typeof window.reportClientError === 'function') {
            window.reportClientError(err, {
                source: 'app',
                errorAction: 'recoverStaleSyncingForCurrentRole',
                reason: reason || 'bootstrap',
                role
            });
        }

        return [];
    }
}

window.recoverStaleSyncingForCurrentRole = recoverStaleSyncingForCurrentRole;

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
                if (nw.state === 'installed' && navigator.serviceWorker.controller) {
                    // Nova verzija čeka — pitaj korisnika za reload
                    if (confirm('Nova verzija aplikacije je dostupna. Osvežiti stranicu?')) {
                        window.location.reload();
                    }
                }
            });
        });
    }).catch(err => {
        console.log('SW registration failed:', err);
    });
}

function _intercomActiveStations(stanice) {
    return (stanice || [])
        .filter(s => s && s.StanicaID)
        .map(s => ({
            entityID:  s.StanicaID,
            name:      s.Naziv || s.StanicaID,
            stationID: s.StanicaID
        }));
}
