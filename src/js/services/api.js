function apiHandleAuthFailure(json) {
    if (json && json.code === 401) {
        if (typeof showToast === 'function') showToast('Sesija istekla', 'error');
        if (typeof doLogout === 'function') doLogout();
        return true;
    }
    return false;
}

async function apiRequest(actionOrParams, payload = {}, options = {}) {
    const timeoutMs = typeof options.timeoutMs === 'number' ? options.timeoutMs : 20000;

    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);

    try {
        let bodyPayload = {};

        if (typeof actionOrParams === 'string' && actionOrParams.indexOf('=') !== -1) {
            const search = new URLSearchParams(actionOrParams);
            for (const [key, value] of search.entries()) {
                bodyPayload[key] = value;
            }
        } else if (typeof actionOrParams === 'string') {
            bodyPayload.action = actionOrParams;
        } else {
            bodyPayload = {};
        }

        const resp = await fetch(CONFIG.API_URL, {
            method: 'POST',
            headers: {
                'Content-Type': 'text/plain',
                'Accept': 'application/json'
            },
            body: JSON.stringify({
                token: CONFIG.TOKEN,
                ...bodyPayload,
                ...payload
            }),
            signal: controller.signal
        });

        const rawText = await resp.text();
        let json = null;

        try {
            json = rawText ? JSON.parse(rawText) : null;
        } catch (parseErr) {
            console.error('API JSON parse failed:', parseErr, rawText);
            return null;
        }

        if (!resp.ok) {
            console.error('API HTTP error:', resp.status, json || rawText);
            if (apiHandleAuthFailure(json)) return null;
            return json || null;
        }

        if (apiHandleAuthFailure(json)) return null;

        return json;
    } catch (e) {
        if (e && e.name === 'AbortError') {
            console.error('API request timeout:', actionOrParams);
        } else {
            console.error('API request failed:', actionOrParams, e);
        }
        return null;
    } finally {
        clearTimeout(timer);
    }
}

window.apiBuildUrl = function () {
    return CONFIG.API_URL;
};

window.apiFetch = async function (params) {
    return apiRequest(params);
};

window.apiPost = async function (action, payload = {}) {
    return apiRequest(action, payload);
};
