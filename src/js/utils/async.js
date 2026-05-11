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

    function redactSensitiveText(value) {
        let out = String(value || '');

        let currentToken = '';
        try {
            currentToken = (window.CONFIG && CONFIG.TOKEN) || '';
        } catch (_) {}

        out = out
            .replace(/(\b(?:token|authToken)\b\s*[:=]\s*["']?)[^"',\s&}]+/gi, '$1[REDACTED]')
            .replace(/((?:token|authToken)=)[^&\s]+/gi, '$1[REDACTED]');

        if (currentToken && currentToken.length >= 8) {
            out = out.split(currentToken).join('[REDACTED]');
        }

        return out;
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

        try {
            role = (window.CONFIG && CONFIG.USER_ROLE) || '';
            entityID = (window.CONFIG && (CONFIG.ENTITY_ID || CONFIG.OTKUPAC_ID)) || '';
            appVersion = (window.CONFIG && CONFIG.APP_VERSION) || '';
        } catch (_) {}

        return {
            errorAction: ctx.errorAction || ctx.action || ctx.source || 'client-error',
            source: ctx.source || 'PWA',
            details: ctx.details || '',
            role,
            entityID: ctx.entityID || entityID,
            appVersion,
            url: window.location ? window.location.href : '',
            userAgent: navigator.userAgent || ''
        };
    }

    async function sendClientErrorPayload(payload) {
        // Prefer normalized API helper when available.
        if (typeof window.apiPostSafe === 'function') {
            await window.apiPostSafe('logClientError', payload, {
                timeoutMs: 8000,
                includeToken: false
            });
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

            const message = redactSensitiveText(getErrorMessage(err));
            const stack = redactSensitiveText(getErrorStack(err));
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
                entityID: meta.entityID,
                errorAction: safeString(redactSensitiveText(meta.errorAction), 120),
                message: safeString(redactSensitiveText(message), 500),
                details: safeString(redactSensitiveText(details), 1500),
                stack: safeString(redactSensitiveText(stack), 1500)
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

        window.withSubmitLock = async function withSubmitLock(lockKey, fn, options) {
        const key = String(lockKey || 'submit');
        const opts = options || {};

        if (!opts.skipMasterSyncGuard && typeof window.ensureMasterSyncNotActive === 'function') {
            const allowed = await window.ensureMasterSyncNotActive(
                opts.reason || opts.action || key,
                { showToast: true }
            );

            if (!allowed) return null;
        }

        const runtime = typeof getAppRuntime === 'function'
            ? getAppRuntime()
            : (window.appRuntime || (window.appRuntime = {}));

        const locks = runtime.submitLocks || (runtime.submitLocks = {});

        if (locks[key]) {
            if (opts.alreadyMessage && typeof showToast === 'function') {
                showToast(opts.alreadyMessage, 'info');
            }
            return null;
        }

        locks[key] = {
            startedAt: new Date().toISOString(),
            reason: opts.reason || ''
        };

        const selector = opts.selector || (
            opts.action ? '[data-action="' + String(opts.action).replace(/"/g, '\\"') + '"]' : ''
        );

        const lockedElements = selector
            ? Array.from(document.querySelectorAll(selector)).map(el => ({
                el,
                disabled: !!el.disabled,
                ariaBusy: el.getAttribute('aria-busy'),
                ariaDisabled: el.getAttribute('aria-disabled')
            }))
            : [];

        lockedElements.forEach(item => {
            const el = item.el;
            el.classList.add('is-submitting');
            el.setAttribute('aria-busy', 'true');
            el.setAttribute('aria-disabled', 'true');

            if ('disabled' in el) {
                el.disabled = true;
            }
        });

        try {
            return await fn();
        } finally {
            delete locks[key];

            lockedElements.forEach(item => {
                const el = item.el;
                el.classList.remove('is-submitting');

                if (item.ariaBusy === null) el.removeAttribute('aria-busy');
                else el.setAttribute('aria-busy', item.ariaBusy);

                if (item.ariaDisabled === null) el.removeAttribute('aria-disabled');
                else el.setAttribute('aria-disabled', item.ariaDisabled);

                if ('disabled' in el) {
                    el.disabled = item.disabled;
                }
            });
        }
    };

    window.installGlobalErrorReporting();
})();
