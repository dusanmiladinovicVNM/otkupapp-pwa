window.localIsoDateFromDate = function localIsoDateFromDate(date) {
    if (!(date instanceof Date) || isNaN(date.getTime())) return '';

    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');

    return year + '-' + month + '-' + day;
};

window.getTodayIsoDate = function getTodayIsoDate() {
    return window.localIsoDateFromDate(new Date());
};

window.getRelativeIsoDate = function getRelativeIsoDate(offsetDays) {
    const d = new Date();
    d.setDate(d.getDate() + (parseInt(offsetDays, 10) || 0));
    return window.localIsoDateFromDate(d);
};

window.toIsoDateOnly = function toIsoDateOnly(input) {
    if (!input) return '';

    if (input instanceof Date) {
        return window.localIsoDateFromDate(input);
    }

    const s = String(input).trim();

    // Already canonical date-only.
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
        return s;
    }

    // Serbian format: 02.05.2026. -> 2026-05-02
    if (/^\d{2}\.\d{2}\.\d{4}\.?$/.test(s)) {
        const clean = s.replace(/\.$/, '');
        const parts = clean.split('.');
        return parts[2] + '-' + parts[1] + '-' + parts[0];
    }

    // ISO timestamp or other parseable date.
    // Important: return LOCAL calendar date, not UTC date.
    try {
        const d = new Date(s);
        if (!isNaN(d.getTime())) {
            return window.localIsoDateFromDate(d);
        }
    } catch (_) {}

    return s;
};

window.fmtDate = function fmtDate(val) {
    return window.toIsoDateOnly(val);
};

window.fmtStanica = function fmtStanica(stanicaID) {
    if (!stanicaID) return '';
    const s = (stammdaten.stanice || []).find(x => x.StanicaID === stanicaID);
    const name = s ? (s.Naziv || s.Mesto || stanicaID) : stanicaID;
    if (name === stanicaID) return stanicaID;
    return name + ' (' + stanicaID + ')';
};

window.normalizeIso = function normalizeIso(value) {
    if (!value) return '';
    try {
        const d = new Date(value);
        if (isNaN(d.getTime())) return String(value);
        return d.toISOString();
    } catch (_) {
        return String(value);
    }
};
