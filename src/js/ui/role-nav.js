function getNormalizedRole() {
    return String((CONFIG && CONFIG.USER_ROLE) || '').trim().toLowerCase();
}

function getRoleNavConfig() {
    const role = getNormalizedRole();

    if (role === 'kooperant') {
        return {
            navId: 'koopBottomNav',
            bodyClass: 'has-koop-bottom-nav',
            type: 'showTab',
            defaultTab: 'home',
            tabMap: {
                home: 'home',
                parcele: 'parcele',
                agromere: 'agromere',
                knjigapolja: 'knjigapolja',
                more: 'more',
                kartica: 'more',
                koopinfo: 'more'
            }
        };
    }

    if (role === 'otkupac') {
        return {
            navId: 'otkupBottomNav',
            bodyClass: 'has-otkup-bottom-nav',
            type: 'showTab',
            defaultTab: 'otkup',
            tabMap: {
                otkup: 'otkup',
                pregled: 'pregled',
                otpremnice: 'otpremnice',
                queue: 'queue'
            }
        };
    }

    if (role === 'management') {
        return {
            navId: 'mgmtBottomNav',
            bodyClass: 'has-mgmt-bottom-nav',
            type: 'showMgmtRoot',
            defaultTab: 'pregled',
            tabMap: {
                pregled: 'pregled',
                dispecer: 'dispecer',
                otkup: 'otkup',
                partneri: 'partneri',
                agro: 'agro'
            }
        };
    }

    if (role === 'vozac') {
        return {
            navId: 'vozacBottomNav',
            bodyClass: 'has-vozac-bottom-nav',
            type: 'showTab',
            defaultTab: 'zbirna',
            tabMap: {
                zbirna: 'zbirna',
                transport: 'transport'
            }
        };
    }

    return null;
}

function updateRoleNavVisibility() {
    const cfg = getRoleNavConfig();

    ['koopBottomNav', 'otkupBottomNav', 'mgmtBottomNav', 'vozacBottomNav'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.classList.remove('visible');
    });

    document.body.classList.remove(
        'has-koop-bottom-nav',
        'has-otkup-bottom-nav',
        'has-mgmt-bottom-nav',
        'has-vozac-bottom-nav'
    );

    if (!cfg) return;

    const nav = document.getElementById(cfg.navId);
    if (!nav) return;

    nav.classList.add('visible');

    if (window.innerWidth <= 900) {
        document.body.classList.add(cfg.bodyClass);
    }
}

function updateRoleNavActive(activeKey) {
    const cfg = getRoleNavConfig();
    if (!cfg) return;

    const nav = document.getElementById(cfg.navId);
    if (!nav) return;

    nav.querySelectorAll('.bottom-nav-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.tab === activeKey);
    });
}

function resolveActiveNavKeyFromDom() {
    const cfg = getRoleNavConfig();
    if (!cfg) return null;

    const activeTab = document.querySelector('.tab-content.active');
    if (!activeTab) return cfg.defaultTab;

    const id = (activeTab.id || '').replace('tab-', '');
    return cfg.tabMap[id] || id || cfg.defaultTab;
}

function syncRoleNavActiveFromDom() {
    const key = resolveActiveNavKeyFromDom();
    if (key) updateRoleNavActive(key);
}

function showRoleNavTab(tabName, btn) {
    const cfg = getRoleNavConfig();
    if (!cfg) return;

    if (cfg.type === 'showMgmtRoot') {
        if (typeof showMgmtRoot === 'function') {
            showMgmtRoot(tabName);
        }
        updateRoleNavActive(tabName);
        return;
    }

    if (typeof showTab === 'function') {
        showTab(tabName);
    }
    updateRoleNavActive(tabName);
}

window.updateRoleNavVisibility = updateRoleNavVisibility;
window.updateRoleNavActive = syncRoleNavActiveFromDom;
window.showRoleNavTab = showRoleNavTab;

window.addEventListener('resize', () => {
    updateRoleNavVisibility();
    syncRoleNavActiveFromDom();
});
