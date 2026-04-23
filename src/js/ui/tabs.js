// ============================================================
// TAB NAVIGATION (non-management)
// ============================================================
function showTab(tabName, btn) {
    // Cleanup agromere GPS kad se napusti tab
    if (
        tabName !== 'agromere' &&
        typeof window.agroState !== 'undefined' &&
        window.agroState &&
        window.agroState.geoWatchId
    ) {
        navigator.geolocation.clearWatch(window.agroState.geoWatchId);
        window.agroState.geoWatchId = null;
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
    if (tabName === 'knjigapolja') loadKnjigaPolja();

    setTimeout(() => {
        if (typeof updateRoleNavActive === 'function') {
            updateRoleNavActive();
        }
    }, 0);
}
