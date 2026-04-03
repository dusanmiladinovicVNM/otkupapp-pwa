window.fmtDate = function (val) {
    if (!val) return '';
    if (typeof val === 'string' && val.length >= 10) return val.substring(0, 10);
    try {
        return new Date(val).toISOString().split('T')[0];
    } catch (e) {
        return String(val);
    }
};
