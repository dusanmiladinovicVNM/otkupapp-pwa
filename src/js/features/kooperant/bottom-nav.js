// ============================================================
// ROLE-AWARE BOTTOM NAV
// Kooperant + Otkupac:
// - mobile  = bottom nav
// - desktop = isti nav gore kao topnav
// Vozac + Management koriste legacy tabBar
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
    const loginContainer = document.getElementById('loginContainer');

    const isLoginVisible = !!(
        loginContainer &&
        loginContainer.offsetParent !== null &&
        loginContainer.innerHTML.trim() !== '' &&
        getComputedStyle(loginContainer).display !== 'none'
    );

    const cfg = getActiveBottomNavConfig();

    ['koopBottomNav', 'otkupBottomNav', 'mgmtBottomNav'].forEach(id => {
        const nav = document.getElementById(id);
        if (nav) nav.classList.remove('visible');
    });
    document.body.classList.remove('has-koop-bottom-nav', 'has-otkup-bottom-nav', 'has-mgmt-bottom-nav');

    const tabBar = document.getElementById('tabBar');

    // ako nema role-nav ili je login ekran, vrati legacy tabBar
    if (!cfg || isLoginVisible) {
        if (tabBar) tabBar.style.display = '';
        return;
    }

    const nav = document.getElementById(cfg.navId);

    // fallback: ako nav ne postoji, vrati legacy tabBar
    if (!nav) {
        if (tabBar) tabBar.style.display = '';
        return;
    }

    // za Kooperant/Otkupac koristimo role nav umesto legacy tabBar
    if (tabBar) tabBar.style.display = 'none';

    nav.classList.add('visible');
    document.body.classList.add(cfg.bodyClass);
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

(function attachBottomNavHook() {
    const originalShowTab = window.showTab;
    if (typeof originalShowTab !== 'function') return;
    if (window.__bottomNavHookAttached) return;

    window.__bottomNavHookAttached = true;

    window.showTab = function patchedShowTab(tabName, btn) {
        originalShowTab(tabName, btn);
        updateBottomNavVisibility();
        updateBottomNavActive();
    };
})();

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
