(function () {
    const CLIENT_ERROR_MAX_PER_PAGE = 10;
    const CLIENT_ERROR_DEDUPE_MS = 10000;

    let clientErrorCount = 0;
    let lastClientErrorKey = '';
    let lastClientErrorAt = 0;

    function safeString(value, maxLen) {
        const s = String(value || '');
        const limit = parseInt(maxLen, 10) || 500;
        return s.length > limit ? s.substring(0, limit) : s;
    }

    function getErrorMessage(err) {
        if (!err) return 'Unknown error';
        if (err.message) return String(err.message);
        return String(err);
    }

    function getErrorStack(err) {
        if (!err) return '';
        if (err.stack) return String(err.stack);
        return '';
    }

    function getClientErrorMeta(context) {
        const ctx = context && typeof context === 'object' ? context : {};

        let role = '';
        let entityID = '';
        let appVersion = '';
        let token = '';

        try {
            role = (window.CONFIG && CONFIG.USER_ROLE) || '';
            entityID = (window.CONFIG && (CONFIG.ENTITY_ID || CONFIG.OTKUPAC_ID)) || '';
            appVersion = (window.CONFIG && CONFIG.APP_VERSION) || '';
            token = (window.CONFIG && CONFIG.TOKEN) || '';
        } catch (_) {}

        return {
            errorAction: ctx.errorAction || ctx.action || ctx.source || 'client-error',
            source: ctx.source || 'PWA',
            details: ctx.details || '',
            role,
            entityID: ctx.entityID || entityID,
            appVersion,
            token,
            url: window.location ? window.location.href : '',
            userAgent: navigator.userAgent || ''
        };
    }

    async function sendClientErrorPayload(payload) {
        // Prefer normalized API helper when available.
        if (typeof window.apiPostSafe === 'function') {
            await window.apiPostSafe('logClientError', payload, { timeoutMs: 8000 });
            return;
        }

        // Early-boot fallback: api.js may not be loaded yet.
        if (window.CONFIG && CONFIG.API_URL) {
            const body = Object.assign({ action: 'logClientError' }, payload);

            await fetch(CONFIG.API_URL, {
                method: 'POST',
                headers: {
                    'Content-Type': 'text/plain',
                    'Accept': 'application/json'
                },
                body: JSON.stringify(body)
            });

            return;
        }

        // No API available yet.
    }

    window.reportClientError = async function reportClientError(err, context) {
        try {
            if (clientErrorCount >= CLIENT_ERROR_MAX_PER_PAGE) return;

            const message = getErrorMessage(err);
            const stack = getErrorStack(err);
            const meta = getClientErrorMeta(context);

            const dedupeKey = [
                meta.errorAction,
                message,
                stack.substring(0, 200)
            ].join('|');

            const now = Date.now();

            if (
                dedupeKey === lastClientErrorKey &&
                now - lastClientErrorAt < CLIENT_ERROR_DEDUPE_MS
            ) {
                return;
            }

            lastClientErrorKey = dedupeKey;
            lastClientErrorAt = now;
            clientErrorCount++;

            const details = [
                meta.details ? 'details=' + meta.details : '',
                meta.role ? 'role=' + meta.role : '',
                meta.appVersion ? 'appVersion=' + meta.appVersion : '',
                meta.url ? 'url=' + meta.url : '',
                meta.userAgent ? 'ua=' + meta.userAgent : '',
                stack ? 'stack=' + stack : ''
            ].filter(Boolean).join('\n');

            await sendClientErrorPayload({
                token: meta.token,
                entityID: meta.entityID,
                errorAction: safeString(meta.errorAction, 120),
                message: safeString(message, 500),
                details: safeString(details, 1500),
                stack: safeString(stack, 1500)
            });
        } catch (reportErr) {
            // Logging must never break app runtime.
            console.warn('[reportClientError] failed:', reportErr);
        }
    };

    window.safeAsync = async function safeAsync(fn, onErrorToast, errorContext) {
        try {
            return await fn();
        } catch (err) {
            console.error('[safeAsync]', err);

            if (typeof window.reportClientError === 'function') {
                window.reportClientError(err, Object.assign({
                    source: 'safeAsync',
                    errorAction: onErrorToast || 'safeAsync'
                }, errorContext || {}));
            }

            if (onErrorToast) {
                showToast(onErrorToast, 'error');
            }

            return null;
        }
    };

    window.installGlobalErrorReporting = function installGlobalErrorReporting() {
        if (window.__clientErrorReportingBound) return;
        window.__clientErrorReportingBound = true;

        window.addEventListener('error', function (event) {
            if (typeof window.reportClientError !== 'function') return;

            window.reportClientError(event.error || event.message, {
                source: 'window.error',
                errorAction: 'window.error',
                details: [
                    event.filename || '',
                    event.lineno ? 'line=' + event.lineno : '',
                    event.colno ? 'col=' + event.colno : ''
                ].filter(Boolean).join(' ')
            });
        });

        window.addEventListener('unhandledrejection', function (event) {
            if (typeof window.reportClientError !== 'function') return;

            window.reportClientError(event.reason || 'Unhandled promise rejection', {
                source: 'window.unhandledrejection',
                errorAction: 'window.unhandledrejection'
            });
        });
    };

    window.installGlobalErrorReporting();
})();
