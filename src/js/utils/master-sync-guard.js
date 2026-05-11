(function () {
    const CACHE_TTL_MS = 5000;
    const POLL_MS = 30000;
    const OFFLINE_LOCK_CACHE_MS = 10 * 60 * 1000;

    let cachedState = null;
    let cachedAt = 0;
    let pollStarted = false;

    function buildState(overrides) {
        return Object.assign({
            success: true,
            locked: false,
            stale: false,
            unknown: false,
            offline: false,
            message: ''
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
                    if (state && state.locked) showMasterSyncOverlay(state);
                    else hideMasterSyncOverlay();
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

    function hideMasterSyncOverlay() {
        const el = document.getElementById('masterSyncBlocker');
        if (el) el.style.display = 'none';
    }

    async function fetchMasterSyncState() {
        if (!navigator.onLine) {
            if (
                cachedState &&
                cachedState.locked &&
                Date.now() - cachedAt < OFFLINE_LOCK_CACHE_MS
            ) {
                return buildState({
                    locked: true,
                    offline: true,
                    message: cachedState.message ||
                        'Master sync je bio aktivan. Sačekajte konekciju za proveru.'
                });
            }

            return buildState({
                locked: false,
                offline: true,
                message: ''
            });
        }

        if (typeof window.apiFetchSafe !== 'function') {
            return buildState({
                locked: true,
                unknown: true,
                message: 'Nije moguće proveriti sync status. Pokušajte ponovo za par sekundi.'
            });
        }

        const result = await window.apiFetchSafe('action=getMasterSyncState', {
            timeoutMs: 5000,
            includeToken: false
        });

        if (!result || !result.ok || !result.data) {
            return buildState({
                locked: true,
                unknown: true,
                message: 'Nije moguće proveriti master sync status. Pokušajte ponovo za par sekundi.'
            });
        }

        const data = result.data;

        return buildState({
            success: data.success !== false,
            locked: !!data.locked,
            stale: !!data.stale,
            unknown: data.success === false,
            message: data.message || data.error || ''
        });
    }

    window.getMasterSyncStateSafe = async function getMasterSyncStateSafe(force) {
        const now = Date.now();

        if (!force && cachedState && now - cachedAt < CACHE_TTL_MS) {
            return cachedState;
        }

        try {
            cachedState = await fetchMasterSyncState();
            cachedAt = now;
            return cachedState;
        } catch (err) {
            console.error('[master-sync-guard] check failed:', err);

            cachedState = buildState({
                locked: navigator.onLine,
                unknown: true,
                message: 'Nije moguće proveriti master sync status.'
            });
            cachedAt = now;

            return cachedState;
        }
    };

    window.ensureMasterSyncNotActive = async function ensureMasterSyncNotActive(context, options) {
        const opts = options || {};
        const state = await window.getMasterSyncStateSafe(true);

        if (state && state.locked) {
            showMasterSyncOverlay(state);

            if (opts.showToast !== false && typeof window.showToast === 'function') {
                window.showToast(
                    state.message || 'Master sync je u toku. Sačekajte završetak.',
                    'warning'
                );
            }

            return false;
        }

        hideMasterSyncOverlay();
        return true;
    };

    window.startMasterSyncGuardPolling = function startMasterSyncGuardPolling() {
        if (pollStarted) return;
        pollStarted = true;

        setInterval(async () => {
            try {
                const state = await window.getMasterSyncStateSafe(true);
                if (state && state.locked) showMasterSyncOverlay(state);
                else hideMasterSyncOverlay();
            } catch (_) {}
        }, POLL_MS);
    };

    document.addEventListener('visibilitychange', () => {
        if (!document.hidden && typeof window.getMasterSyncStateSafe === 'function') {
            window.getMasterSyncStateSafe(true).then(state => {
                if (state && state.locked) showMasterSyncOverlay(state);
                else hideMasterSyncOverlay();
            }).catch(() => {});
        }
    });

    window.addEventListener('online', () => {
        if (typeof window.getMasterSyncStateSafe === 'function') {
            window.getMasterSyncStateSafe(true).then(state => {
                if (state && state.locked) showMasterSyncOverlay(state);
                else hideMasterSyncOverlay();
            }).catch(() => {});
        }
    });

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', () => {
            window.startMasterSyncGuardPolling();
        });
    } else {
        window.startMasterSyncGuardPolling();
    }
})();
