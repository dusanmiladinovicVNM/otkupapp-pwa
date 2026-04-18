// ============================================================
// ROLE-AWARE BOTTOM NAV
// Kooperant + Otkupac:
// - mobile  = bottom nav
// - desktop = isti nav gore kao topnav
//
// Management ima svoj poseban shell/nav layer.
// Ovaj fajl NE SME da preuzima ownership nad MGMT navigacijom
// niti da patchuje globalni showTab.
// ============================================================

function getCurrentRole() {
    return String((CONFIG && CONFIG.USER_ROLE) || '').trim().toLowerCase();
}

function getActiveBottomNavConfig() {
    const role = getCurrentRole();

    if (role === 'kooperant') {
        return {
            role: 'kooperant',
            navId: 'koopBottomNav',
            bodyClass: 'has-koop-bottom-nav',
            tabMap: {
                'tab-home': 'home',
                'tab-parcele': 'parcele',
                'tab-agromere': 'agromere',
                'tab-knjigapolja': 'knjigapolja',
                'tab-more': 'more',
                'tab-kartica': 'more',
                'tab-koopinfo': 'more'
            }
        };
    }

    if (role === 'otkupac') {
        return {
            role: 'otkupac',
            navId: 'otkupBottomNav',
            bodyClass: 'has-otkup-bottom-nav',
            tabMap: {
                'tab-otkup': 'otkup',
                'tab-pregled': 'pregled',
                'tab-otpremnice': 'otpremnice',
                'tab-queue': 'queue'
            }
        };
    }

    return null;
}

function updateBottomNavVisibility() {
    const cfg = getActiveBottomNavConfig();
    const isMobile = window.innerWidth <= 900;

    const koopNav = document.getElementById('koopBottomNav');
    const otkupNav = document.getElementById('otkupBottomNav');

    // Reset oba nav-a
    if (koopNav) koopNav.classList.remove('visible');
    if (otkupNav) otkupNav.classList.remove('visible');

    document.body.classList.remove('has-koop-bottom-nav', 'has-otkup-bottom-nav');

    // Ovaj fajl upravlja samo Kooperant/Otkupac mobile bottom nav-om
    if (!cfg || !isMobile) return;

    const nav = document.getElementById(cfg.navId);
    if (nav) nav.classList.add('visible');
    document.body.classList.add(cfg.bodyClass);
}

function updateBottomNavButtons(tabName, btn, navId) {
    const root = document.getElementById(navId);
    if (!root) return;

    root.querySelectorAll('.bottom-nav-btn').forEach(b => {
        b.classList.remove('active');
    });

    if (btn && btn.classList) {
        btn.classList.add('active');
        return;
    }

    const activeBtn = root.querySelector(`.bottom-nav-btn[data-tab="${tabName}"]`);
    if (activeBtn) activeBtn.classList.add('active');
}

function updateBottomNavActive() {
    const cfg = getActiveBottomNavConfig();
    if (!cfg) return;

    const activeTab = document.querySelector('.tab-content.active');
    if (!activeTab) return;

    const id = activeTab.id || '';
    const navTab = cfg.tabMap[id];
    if (!navTab) return;

    updateBottomNavButtons(navTab, null, cfg.navId);
}

function findLegacyTabBtn(tabName) {
    const buttons = Array.from(document.querySelectorAll('.tab-btn'));
    return buttons.find(btn => {
        const onclick = String(btn.getAttribute('onclick') || '');
        return onclick.includes(`showTab('${tabName}'`) || onclick.includes(`showTab("${tabName}"`);
    }) || null;
}

function showBottomTab(tabName, btn) {
    const cfg = getActiveBottomNavConfig();
    if (!cfg) return;

    if (cfg.role === 'kooperant' && tabName === 'more') {
        qsa('.tab-content').forEach(t => removeClass(t, 'active'));
        qsa('.tab-btn').forEach(b => removeClass(b, 'active'));

        const tabEl = byId('tab-more');
        if (tabEl) addClass(tabEl, 'active');

        updateBottomNavButtons(tabName, btn, cfg.navId);
        return;
    }

    showTab(tabName, findLegacyTabBtn(tabName) || btn);
    updateBottomNavButtons(tabName, btn, cfg.navId);
}

async function syncKooperantFromMore() {
    if (typeof syncKooperantNow !== 'function') {
        showToast('Sync nije dostupan', 'error');
        return;
    }

    showToast('Pokrećem sinhronizaciju...', 'info');

    try {
        const result = await syncKooperantNow();
        console.log('syncKooperantNow', result);
        invalidatePregledCacheSafe();
        showToast('Sinhronizacija završena', 'success');
    } catch (e) {
        console.error(e);
        showToast('Greška pri sinhronizaciji', 'error');
    }
}

function invalidatePregledCacheSafe() {
    if (typeof invalidatePregledCache === 'function') {
        invalidatePregledCache();
    }
}

window.initBottomNav = function () {
    updateBottomNavVisibility();
    updateBottomNavActive();
};

window.updateBottomNavVisibility = updateBottomNavVisibility;
window.updateBottomNavActive = updateBottomNavActive;

window.addEventListener('resize', () => {
    updateBottomNavVisibility();
    updateBottomNavActive();
});
