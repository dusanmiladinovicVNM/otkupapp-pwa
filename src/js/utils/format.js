window.fmtDate = function (val) {
    if (!val) return '';
    if (typeof val === 'string' && val.length >= 10) return val.substring(0, 10);
    try {
        return new Date(val).toISOString().split('T')[0];
    } catch (e) {
        return String(val);
    }
};

window.fmtStanica = function(stanicaID) {
    if (!stanicaID) return '';
    const s = (stammdaten.stanice || []).find(x => x.StanicaID === stanicaID);
    const name = s ? (s.Naziv || s.Mesto || stanicaID) : stanicaID;
    if (name === stanicaID) return stanicaID;
    return name + ' (' + stanicaID + ')';
};

window.normalizeIso = function (value) {
    if (!value) return '';
    try {
        const d = new Date(value);
        if (isNaN(d.getTime())) return String(value);
        return d.toISOString();
    } catch (_) {
        return String(value);
    }
};
