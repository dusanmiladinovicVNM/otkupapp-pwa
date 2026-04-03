window.getLs = function (key, fallback = null) {
    try {
        const value = localStorage.getItem(key);
        return value === null ? fallback : value;
    } catch (e) {
        return fallback;
    }
};

window.setLs = function (key, value) {
    try {
        localStorage.setItem(key, value);
        return true;
    } catch (e) {
        return false;
    }
};

window.removeLs = function (keys) {
    try {
        if (Array.isArray(keys)) {
            keys.forEach(key => localStorage.removeItem(key));
        } else if (typeof keys === 'string') {
            localStorage.removeItem(keys);
        }
        return true;
    } catch (e) {
        return false;
    }
};

window.getDeviceID = function () {
    let id = getLs('deviceID', '');
    if (!id) {
        id = 'DEV-' + crypto.randomUUID().slice(0, 8);
        setLs('deviceID', id);
    }
    return id;
};
