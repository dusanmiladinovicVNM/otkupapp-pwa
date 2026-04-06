// ============================================================
// TAB NAVIGATION (non-management)
// ============================================================
function showTab(tabName, btn) {
    qsa('.tab-content').forEach(t => removeClass(t, 'active'));
    qsa('.tab-btn').forEach(b => removeClass(b, 'active'));

    const tabEl = byId('tab-' + tabName);
    if (tabEl) addClass(tabEl, 'active');

    // btn dolazi iz onclick="showTab('xxx', this)"
    if (btn && btn.classList) {
        addClass(btn, 'active');
    }

    if (tabName === 'queue') { renderQueueList(); updateStats(); }
    if (tabName === 'pregled') loadOtkupPregled();
    if (tabName === 'otpremnice') loadOtpremaOverview();
    if (tabName === 'kartica') loadKartica();
    if (tabName === 'parcele') loadParcele();
    if (tabName === 'agromere') loadAgronom();
    if (tabName === 'koopinfo') loadKoopInfo();
    if (tabName === 'zbirna') loadVozacData();
    if (tabName === 'transport') loadVozacTransport();
    if (tabName === 'dispecer') loadDispecer();
}
