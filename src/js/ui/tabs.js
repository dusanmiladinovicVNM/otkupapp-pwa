// ============================================================
// TAB NAVIGATION (non-management)
// ============================================================
function showTab(tabName, btn) {
    // Cleanup agromere GPS kad se napusti tab
    if (tabName !== 'agromere' && agroState && agroState.geoWatchId) {
        navigator.geolocation.clearWatch(agroState.geoWatchId);
        agroState.geoWatchId = null;
    }
    qsa('.tab-content').forEach(t => removeClass(t, 'active'));
    qsa('.tab-btn').forEach(b => removeClass(b, 'active'));

    const tabEl = byId('tab-' + tabName);
    if (tabEl) addClass(tabEl, 'active');

    // btn dolazi iz onclick="showTab('xxx', this)"
    if (btn && btn.classList) {
        addClass(btn, 'active');
    }

    if (tabName === 'queue') {
        if (CONFIG.USER_ROLE === 'Otkupac' && typeof loadOtkupacMore === 'function') {
            loadOtkupacMore();
        } else {
            renderQueueList();
            updateStats();
        }
    }
    if (tabName === 'pregled') loadOtkupPregled();
    if (tabName === 'otpremnice') loadOtpremaOverview();
    if (tabName === 'home') loadPregled();
    if (tabName === 'kartica') loadKartica();
    if (tabName === 'parcele') loadParcele();
    if (tabName === 'agromere') loadAgronom();
    if (tabName === 'koopinfo') loadKoopInfo();
    if (tabName === 'zbirna') loadVozacData();
    if (tabName === 'transport') loadVozacTransport();
    if (tabName === 'dispecer') loadDispecer();
    if (tabName === 'knjigapolja') loadKnjigaPolja();
}
