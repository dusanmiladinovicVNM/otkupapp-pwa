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

window.formatNumber = function formatNumber(value, options) {
    const n = Number(value);
    const opts = options || {};

    if (!Number.isFinite(n)) {
        return opts.fallback || '0';
    }

    return n.toLocaleString('sr-RS', {
        minimumFractionDigits: opts.minimumFractionDigits || 0,
        maximumFractionDigits: typeof opts.maximumFractionDigits === 'number'
            ? opts.maximumFractionDigits
            : 2
    });
};

window.formatKg = function formatKg(value) {
    return window.formatNumber(value, {
        maximumFractionDigits: 2
    }) + ' kg';
};

window.formatMoney = function formatMoney(value) {
    return window.formatNumber(value, {
        maximumFractionDigits: 2
    }) + ' RSD';
};

/**
 * Parse decimal user input safely. Handles Serbian locale where users may
 * type comma "," as decimal separator (mobile keyboards, regional habit).
 *
 *   parseDecimalInput("120,50")  => 120.5
 *   parseDecimalInput("120.50")  => 120.5
 *   parseDecimalInput("  120 ")  => 120
 *   parseDecimalInput("")        => 0
 *   parseDecimalInput(null)      => 0
 *   parseDecimalInput("abc")     => 0
 */
window.parseDecimalInput = function parseDecimalInput(value) {
    const s = String(value == null ? '' : value).trim().replace(',', '.');
    const n = parseFloat(s);
    return Number.isFinite(n) ? n : 0;
};
