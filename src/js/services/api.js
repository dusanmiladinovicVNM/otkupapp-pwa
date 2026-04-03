window.apiBuildUrl = function (params) {
    return CONFIG.API_URL + '?' + params + '&token=' + encodeURIComponent(CONFIG.TOKEN);
};

window.apiFetch = async function (params) {
    const url = apiBuildUrl(params);

    try {
        const resp = await fetch(url);
        const json = await resp.json();

        if (json && json.code === 401) {
            if (typeof showToast === 'function') showToast('Sesija istekla', 'error');
            if (typeof doLogout === 'function') doLogout();
            return null;
        }

        return json;
    } catch (e) {
        console.log('API fetch failed:', e);
        return null;
    }
};

window.apiPost = async function (action, payload = {}) {
    try {
        const resp = await fetch(CONFIG.API_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify({
                action,
                token: CONFIG.TOKEN,
                ...payload
            })
        });

        const json = await resp.json();

        if (json && json.code === 401) {
            if (typeof showToast === 'function') showToast('Sesija istekla', 'error');
            if (typeof doLogout === 'function') doLogout();
            return null;
        }

        return json;
    } catch (e) {
        console.log('API post failed:', e);
        return null;
    }
};
