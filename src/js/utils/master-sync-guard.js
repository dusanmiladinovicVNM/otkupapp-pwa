(function () {
    const CACHE_TTL_MS = 10000;
    const POLL_MS = 30000;
    const OFFLINE_LOCK_CACHE_MS = 10 * 60 * 1000;
    const FAIL_BACKOFF_MS = 45000;
    const REQUEST_TIMEOUT_MS = 12000;

    const cachedStates = new Map();   // key: stanicaID ('' za global), value: { state, at }
    let lastFailAt = 0;
    let inFlight = new Map();          // key: stanicaID, value: Promise
    let pollTimer = null;

    function buildState(overrides) {
        return Object.assign({
            success: true,
            locked: false,
            stale: false,
            unknown: false,
            offline: false,
            timeout: false,
            message: '',
            code: '',
            lockKind: '',   // ← DODATO: '' | 'master' | 'stanica'

            // EPOHA MASTER SYNC-a (review #390, deveti krug).
            //
            // Server je salje, a klijent ju je bacao -- pa se "da li se nesto
            // promenilo od mog poslednjeg snimka" nije imalo cime izmeriti.
            updatedAt: ''
        }, overrides || {});
    }

    function ensureMasterSyncOverlay() {
        let el = document.getElementById('masterSyncBlocker');
        if (el) return el;

        el = document.createElement('div');
        el.id = 'masterSyncBlocker';
        el.style.cssText = [
            'display:none',
            'position:fixed',
            'inset:0',
            'z-index:99999',
            'background:rgba(15,23,42,0.72)',
            'backdrop-filter:blur(2px)',
            'align-items:center',
            'justify-content:center',
            'padding:20px'
        ].join(';');

        el.innerHTML = `
            <div style="
                max-width:420px;
                width:100%;
                background:#fff;
                border-radius:16px;
                padding:22px;
                box-shadow:0 20px 50px rgba(0,0,0,.25);
                text-align:center;
                color:#0f172a;
            ">
                <div style="font-size:34px;margin-bottom:10px;">🔄</div>
                <div style="font-size:18px;font-weight:800;margin-bottom:8px;">
                    Sinhronizacija je u toku
                </div>
                <div id="masterSyncBlockerMessage" style="font-size:14px;line-height:1.45;color:#475569;">
                    Master računar trenutno sinhronizuje podatke. Sačekajte završetak.
                </div>
                <button id="masterSyncBlockerRefresh" type="button" style="
                    margin-top:16px;
                    border:0;
                    border-radius:10px;
                    padding:10px 14px;
                    font-weight:700;
                    background:#0f766e;
                    color:#fff;
                ">
                    Proveri ponovo
                </button>
            </div>
        `;

        document.body.appendChild(el);

        const btn = document.getElementById('masterSyncBlockerRefresh');
        if (btn) {
            btn.addEventListener('click', async () => {
                btn.disabled = true;
                btn.textContent = 'Proveravam...';

                try {
                    const state = await window.getMasterSyncStateSafe(true);
                    if (state && state.locked === true) showMasterSyncOverlay(state);
                    else await hideMasterSyncOverlay(state);
                } finally {
                    btn.disabled = false;
                    btn.textContent = 'Proveri ponovo';
                }
            });
        }

        return el;
    }

    function showMasterSyncOverlay(state) {
        const el = ensureMasterSyncOverlay();
        const msg = document.getElementById('masterSyncBlockerMessage');

        if (msg) {
            msg.textContent = (state && state.message) ||
                'Master računar trenutno sinhronizuje podatke. Sačekajte završetak.';
        }

        el.style.display = 'flex';
    }

    function sakrijOverlay() {
        const el = document.getElementById('masterSyncBlocker');
        if (el) el.style.display = 'none';
    }

    // KRITERIJUM JE EPOHA, NE VIDLJIVOST OVERLAY-a (review #390, deveti krug).
    //
    // Prethodna verzija je osvezavala samo ako je overlay BIO PRIKAZAN. To nije
    // isto sto i "master epoha se promenila": uredjaj koji je ceo lock interval
    // proveo u pozadini -- ili je ciklus prosao izmedju dva polling tick-a --
    // overlay nikad nije ni video, pa nije ni osvezavao. Ostajao je na
    // ZASTARELOM "free" i, posto je lock vec skinut, klik je bio dozvoljen.
    //
    // Sada se pita ono sto stvarno odlucuje: je li epoha ista kao ona za koju je
    // trenutni snimak potvrdjen. refreshOtpremaPosleLocka to i proverava, pa je
    // poziv jeftin kad se nista nije promenilo.
    //
    // VRACA: true = stanje je potvrdjeno sveze (upis sme), false = nije (overlay
    // ostaje). Pozivalac ne sme da tumaci "zavrsilo je" kao "bezbedno je".
    async function hideMasterSyncOverlay(state) {
        if (typeof window.refreshOtpremaPosleLocka === 'function') {
            try {
                await window.refreshOtpremaPosleLocka(
                    String((state && state.updatedAt) || '')
                );
            } catch (err) {
                console.error('refreshOtpremaPosleLocka failed:', err);
                postaviPorukuOverlay(
                    'Sinhronizacija je zavrsena, ali osvezavanje stanja nije uspelo. ' +
                    'Proveri vezu pa probaj ponovo -- do tada se predaja ne moze potvrditi.'
                );
                prikaziOverlay();
                return false;   // overlay OSTAJE
            }
        }

        sakrijOverlay();
        return true;
    }

    function prikaziOverlay() {
        const el = document.getElementById('masterSyncBlocker');
        if (el) el.style.display = 'flex';
    }

    function postaviPorukuOverlay(tekst) {
        const msg = document.getElementById('masterSyncBlockerMessage');
        if (msg) msg.textContent = tekst;
    }

    async function fetchMasterSyncState(force, stanicaID) {
        stanicaID = stanicaID || '';
        const now = Date.now();
        const cached = cachedStates.get(stanicaID);

        if (!force && cached && now - cached.at < CACHE_TTL_MS) {
            return cached.state;
        }

        if (!force && inFlight.has(stanicaID)) {
            return inFlight.get(stanicaID);
        }

        if (!force && lastFailAt && now - lastFailAt < FAIL_BACKOFF_MS) {
            return (cached && cached.state) || buildState({
                success: false,
                locked: false,
                unknown: true,
                code: 'MASTER_SYNC_STATE_BACKOFF'
            });
        }

        if (!navigator.onLine) {
            if (cached && cached.state.locked && now - cached.at < OFFLINE_LOCK_CACHE_MS) {
                return buildState({
                    locked: true,
                    offline: true,
                    message: cached.state.message ||
                        'Sinhronizacija je bila aktivna. Sačekajte konekciju za proveru.',
                    lockKind: cached.state.lockKind || '',
                    updatedAt: cached.state.updatedAt || ''
                });
            }
            return buildState({ locked: false, offline: true });
        }

        if (typeof window.apiFetchSafe !== 'function') {
            return buildState({
                success: false, locked: false, unknown: true,
                code: 'API_HELPER_MISSING'
            });
        }

        const request = (async function () {
            try {
                const params = stanicaID
                    ? 'action=getMasterSyncState&stanicaID=' + encodeURIComponent(stanicaID)
                    : 'action=getMasterSyncState';

                const result = await window.apiFetchSafe(params, {
                    timeoutMs: REQUEST_TIMEOUT_MS,
                    includeToken: false,
                    silent: true
                });

                if (!result || !result.ok || !result.data) {
                    lastFailAt = Date.now();
                    return buildState({
                        success: false, locked: false, unknown: true,
                        timeout: !!(result && result.isTimeout),
                        code: result && result.isTimeout
                            ? 'MASTER_SYNC_STATE_TIMEOUT'
                            : 'MASTER_SYNC_STATE_UNAVAILABLE'
                    });
                }

                const data = result.data;
                const state = buildState({
                    success: data.success !== false,
                    locked: data.success !== false && data.locked === true,
                    stale: data.stale === true,
                    unknown: data.success === false,
                    message: data.message || data.error || '',
                    code: data.code || '',
                    lockKind: data.lockKind || '',
                    updatedAt: String(data.updatedAt || '')
                });

                cachedStates.set(stanicaID, { state: state, at: Date.now() });
                lastFailAt = 0;
                return state;

            } catch (err) {
                lastFailAt = Date.now();
                return buildState({
                    success: false, locked: false, unknown: true,
                    code: 'MASTER_SYNC_STATE_EXCEPTION'
                });
            } finally {
                inFlight.delete(stanicaID);
            }
        })();

        inFlight.set(stanicaID, request);
        return request;
    }

    window.getMasterSyncStateSafe = async function getMasterSyncStateSafe(force, stanicaID) {
        try {
            return await fetchMasterSyncState(force === true, stanicaID || '');
        } catch (_) {
            return buildState({
                success: false,
                locked: false,
                unknown: true,
                code: 'MASTER_SYNC_STATE_ERROR'
            });
        }
    };

    window.ensureMasterSyncNotActive = async function ensureMasterSyncNotActive(context, options) {
        const opts = options || {};
        const stanicaID = opts.stanicaID || '';
        
        // Ne forsiraj network svaki put. Ako cache kaže unlocked, pusti.
        // Server-side GAS blocker je finalna zaštita za realan write.
        const state = await window.getMasterSyncStateSafe(false, stanicaID);

        if (state && state.locked === true) {
            showMasterSyncOverlay(state);

            if (opts.showToast !== false && typeof window.showToast === 'function') {
                const msg = state.message ||
                    (state.lockKind === 'stanica'
                        ? 'Stanica je trenutno u kancelarijskoj obradi. Sačekajte završetak.'
                        : 'Master sync je u toku. Unos je sačuvan lokalno i biće poslat kasnije.');
                window.showToast(msg, 'warning');
            }

            return false;
        }

        // Unknown/timeout nije potvrđen lock.
        // Soft-lock model: lokalni rad i pokušaj sync-a smeju dalje,
        // GAS će vratiti MASTER_SYNC_ACTIVE ako je lock stvarno aktivan.
        //
        // POVRATNA VREDNOST PRATI ISHOD (review #390, deveti krug).
        //
        // true ovde znaci "upis sme da krene". Ranije je stizao i kad strog
        // refresh padne, jer je hideMasterSyncOverlay gutao neuspeh -- pa bi
        // withSubmitLock pustio komandu nad stanjem koje nije potvrdjeno.
        const spremno = await hideMasterSyncOverlay(state);
        return spremno === true;
    };

    window.startMasterSyncGuardPolling = function startMasterSyncGuardPolling() {
        if (pollTimer) return;

        const tick = async function () {
            try {
                const state = await window.getMasterSyncStateSafe(false);

                if (state && state.locked === true) {
                    showMasterSyncOverlay(state);
                } else {
                    await hideMasterSyncOverlay(state);
                }
            } catch (_) {
                // fail-open za status-check
            } finally {
                pollTimer = setTimeout(tick, POLL_MS);
            }
        };

        pollTimer = setTimeout(tick, 3000);
    };

    document.addEventListener('visibilitychange', () => {
        if (document.hidden) return;

        window.getMasterSyncStateSafe(false).then(async state => {
            if (state && state.locked === true) showMasterSyncOverlay(state);
            else await hideMasterSyncOverlay(state);
        }).catch(() => {});
    });

    window.addEventListener('online', () => {
        window.getMasterSyncStateSafe(false).then(async state => {
            if (state && state.locked === true) showMasterSyncOverlay(state);
            else await hideMasterSyncOverlay(state);
        }).catch(() => {});
    });

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', () => {
            window.startMasterSyncGuardPolling();
        });
    } else {
        window.startMasterSyncGuardPolling();
    }
})();
