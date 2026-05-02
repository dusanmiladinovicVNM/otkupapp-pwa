// src/js/state.js

(function () {
    const initialState = {
        db: null,
        stammdaten: {
            kooperanti: [],
            kulture: [],
            config: [],
            parcele: [],
            stanice: [],
            kupci: [],
            vozaci: []
        },
        mgmtData: null,
        qrScanner: null,
        selectedMera: '',
        parcelExpertOpen: {},
        runtime: {
            initStarted: false,
            appReady: false,
            stammdatenReady: false,
            syncInFlight: false,
            stammdatenRefreshInFlight: false,
            syncIntervalId: null,
            sync: {
                queueInFlight: false,
                otkupacInFlight: false,
                tretmaniInFlight: false,
                troskoviInFlight: false,
                zbirnaInFlight: false
        }
    }
};
    const state = structuredClone(initialState);
    const listeners = new Map();

    function emit(key, value) {
        if (!listeners.has(key)) return;
        listeners.get(key).forEach(fn => {
            try { fn(value, state); } catch (e) { console.error('state listener error', key, e); }
        });
    }

    function get(path) {
        if (!path) return state;
        return path.split('.').reduce((acc, part) => acc ? acc[part] : undefined, state);
    }

    function set(path, value) {
        const parts = path.split('.');
        let ref = state;

        for (let i = 0; i < parts.length - 1; i++) {
            const k = parts[i];
            if (typeof ref[k] !== 'object' || ref[k] === null) ref[k] = {};
            ref = ref[k];
        }

        ref[parts[parts.length - 1]] = value;
        emit(path, value);
        return value;
    }

    function patch(path, partial) {
        const curr = get(path) || {};
        const next = { ...curr, ...partial };
        set(path, next);
        return next;
    }

    function subscribe(path, fn) {
        if (!listeners.has(path)) listeners.set(path, new Set());
        listeners.get(path).add(fn);
        return () => listeners.get(path).delete(fn);
    }

    window.AppState = {
        get,
        set,
        patch,
        subscribe,
        reset() {
            Object.keys(initialState).forEach(k => {
                state[k] = structuredClone(initialState[k]);
            });
        }
    };

    // Compatibility layer za legacy kod
    Object.defineProperty(window, 'db', {
        get() { return window.AppState.get('db'); },
        set(v) { window.AppState.set('db', v); }
    });

    Object.defineProperty(window, 'stammdaten', {
        get() { return window.AppState.get('stammdaten'); },
        set(v) { window.AppState.set('stammdaten', v); }
    });

    Object.defineProperty(window, 'mgmtData', {
        get() { return window.AppState.get('mgmtData'); },
        set(v) { window.AppState.set('mgmtData', v); }
    });

    Object.defineProperty(window, 'qrScanner', {
        get() { return window.AppState.get('qrScanner'); },
        set(v) { window.AppState.set('qrScanner', v); }
    });

    Object.defineProperty(window, 'selectedMera', {
        get() { return window.AppState.get('selectedMera'); },
        set(v) { window.AppState.set('selectedMera', v); }
    });

    Object.defineProperty(window, 'parcelExpertOpen', {
        get() { return window.AppState.get('parcelExpertOpen'); },
        set(v) { window.AppState.set('parcelExpertOpen', v); }
    });

    Object.defineProperty(window, 'appRuntime', {
        get() { return window.AppState.get('runtime'); },
        set(v) { window.AppState.set('runtime', v); }
    });
})();
