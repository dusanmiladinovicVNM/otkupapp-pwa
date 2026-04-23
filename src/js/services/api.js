function apiHandleAuthFailure(json) {
    if (json && json.code === 401) {
        if (typeof showToast === 'function') showToast('Sesija istekla', 'error');
        if (typeof doLogout === 'function') doLogout();
        return true;
    }
    return false;
}

function apiNormalizePayload(actionOrParams, payload = {}) {
    let bodyPayload = {};

    if (typeof actionOrParams === 'string' && actionOrParams.indexOf('=') !== -1) {
        const search = new URLSearchParams(actionOrParams);
        for (const [key, value] of search.entries()) {
            bodyPayload[key] = value;
        }
    } else if (typeof actionOrParams === 'string') {
        bodyPayload.action = actionOrParams;
    }

    return {
        token: CONFIG.TOKEN,
        ...bodyPayload,
        ...payload
    };
}

function apiBuildResult(fields) {
    return {
        ok: !!fields.ok,
        status: typeof fields.status === 'number' ? fields.status : 0,
        data: typeof fields.data === 'undefined' ? null : fields.data,
        error: fields.error || '',
        code: typeof fields.code === 'undefined' ? null : fields.code,
        isTimeout: !!fields.isTimeout,
        isNetworkError: !!fields.isNetworkError,
        isAuthError: !!fields.isAuthError
    };
}

async function apiRequestSafe(actionOrParams, payload = {}, options = {}) {
    const timeoutMs = typeof options.timeoutMs === 'number' ? options.timeoutMs : 20000;

    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);

    try {
        const requestBody = apiNormalizePayload(actionOrParams, payload);

        const resp = await fetch(CONFIG.API_URL, {
            method: 'POST',
            headers: {
                'Content-Type': 'text/plain',
                'Accept': 'application/json'
            },
            body: JSON.stringify(requestBody),
            signal: controller.signal
        });

        const rawText = await resp.text();
        let json = null;

        try {
            json = rawText ? JSON.parse(rawText) : null;
        } catch (parseErr) {
            console.error('API JSON parse failed:', parseErr, rawText);
            return apiBuildResult({
                ok: false,
                status: resp.status,
                error: 'Neispravan JSON odgovor sa servera',
                code: 'bad_json'
            });
        }

        const isAuthError = apiHandleAuthFailure(json);

        if (!resp.ok) {
            console.error('API HTTP error:', resp.status, json || rawText);
            return apiBuildResult({
                ok: false,
                status: resp.status,
                data: json,
                error: (json && (json.error || json.message)) || ('HTTP ' + resp.status),
                code: json && typeof json.code !== 'undefined' ? json.code : resp.status,
                isAuthError: isAuthError
            });
        }

        if (json && json.success === false) {
            return apiBuildResult({
                ok: false,
                status: resp.status,
                data: json,
                error: json.error || json.message || 'API request failed',
                code: typeof json.code !== 'undefined' ? json.code : null,
                isAuthError: isAuthError
            });
        }

        return apiBuildResult({
            ok: true,
            status: resp.status,
            data: json,
            error: '',
            code: json && typeof json.code !== 'undefined' ? json.code : null,
            isAuthError: isAuthError
        });
    } catch (e) {
        if (e && e.name === 'AbortError') {
            console.error('API request timeout:', actionOrParams);
            return apiBuildResult({
                ok: false,
                status: 0,
                error: 'Request timeout',
                code: 'timeout',
                isTimeout: true
            });
        }

        console.error('API request failed:', actionOrParams, e);
        return apiBuildResult({
            ok: false,
            status: 0,
            error: e && e.message ? e.message : 'Network error',
            code: 'network_error',
            isNetworkError: true
        });
    } finally {
        clearTimeout(timer);
    }
}

window.apiBuildUrl = function () {
    return CONFIG.API_URL;
};

// backwards-compatible raw helpers
window.apiFetch = async function (params, options = {}) {
    const result = await apiRequestSafe(params, {}, options);

    if (result.isAuthError) return null;
    if (result.ok) return result.data;
    if (result.data) return result.data;

    return null;
};

window.apiPost = async function (action, payload = {}, options = {}) {
    const result = await apiRequestSafe(action, payload, options);

    if (result.isAuthError) return null;
    if (result.ok) return result.data;
    if (result.data) return result.data;

    return null;
};

// new normalized helpers
window.apiFetchSafe = async function (params, options = {}) {
    return apiRequestSafe(params, {}, options);
};

window.apiPostSafe = async function (action, payload = {}, options = {}) {
    return apiRequestSafe(action, payload, options);
};
