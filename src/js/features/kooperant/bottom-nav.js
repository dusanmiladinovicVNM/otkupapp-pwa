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
    const isMobile = window.innerWidth <= 900;
    const role = CONFIG.USER_ROLE;

    const koopNav = document.getElementById('koopBottomNav');
    const otkupNav = document.getElementById('otkupBottomNav');
    const mgmtNav = document.getElementById('mgmtBottomNav');

    if (koopNav) koopNav.classList.toggle('visible', role === 'Kooperant' && isMobile);
    if (otkupNav) otkupNav.classList.toggle('visible', role === 'Otkupac' && isMobile);
    if (mgmtNav) mgmtNav.classList.toggle('visible', role === 'Management' && isMobile);

    document.body.classList.toggle('has-koop-bottom-nav', role === 'Kooperant' && isMobile);
    document.body.classList.toggle('has-otkup-bottom-nav', role === 'Otkupac' && isMobile);
    document.body.classList.toggle('has-mgmt-bottom-nav', role === 'Management' && isMobile);
}

function updateBottomNavActive() {
    if (CONFIG.USER_ROLE === 'Management' && typeof updateMgmtBottomNavActive === 'function') {
        updateMgmtBottomNavActive();
        return;
    }

    if (CONFIG.USER_ROLE === 'Kooperant') {
        const active = document.querySelector('.tab-content.active');
        const activeId = active ? active.id.replace('tab-', '') : '';
        document.querySelectorAll('#koopBottomNav .bottom-nav-btn').forEach(btn => {
            btn.classList.toggle('active', btn.dataset.tab === activeId);
        });
        return;
    }

    if (CONFIG.USER_ROLE === 'Otkupac') {
        const active = document.querySelector('.tab-content.active');
        const activeId = active ? active.id.replace('tab-', '') : '';
        document.querySelectorAll('#otkupBottomNav .bottom-nav-btn').forEach(btn => {
            btn.classList.toggle('active', btn.dataset.tab === activeId);
        });
    }
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
