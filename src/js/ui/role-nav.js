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

function updateRoleNavActive(targetKey) {
    const cfg = getRoleNavConfig();
    if (!cfg) return;

    const nav = document.getElementById(cfg.navId);
    if (!nav) return;

    const finalKey = targetKey || resolveActiveRoleNavKey() || cfg.defaultKey;

    // Prvo očisti active sa svih role nav-ova
    clearAllRoleNavActiveStates();

    // Onda aktiviraj samo odgovarajući button za aktivnu rolu
    nav.querySelectorAll('.bottom-nav-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.tab === finalKey);
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

function showRoleNavTab(tabKey, btn) {
    const cfg = getRoleNavConfig();
    if (!cfg) return;

    // Odmah prebaci aktivno stanje u nav-u
    updateRoleNavActive(tabKey);

    if (cfg.showMode === 'showMgmtRoot') {
        if (typeof showMgmtRoot === 'function') {
            showMgmtRoot(tabKey);
        }

        // Posle rendera ponovo sinhronizuj iz stvarnog state-a
        setTimeout(() => updateRoleNavActive(), 0);
        return;
    }

    if (typeof showTab === 'function') {
        showTab(tabKey);
    }

    // Posle promene taba ponovo sinhronizuj iz DOM-a
    setTimeout(() => updateRoleNavActive(), 0);
}

function clearAllRoleNavActiveStates() {
    getAllRoleNavIds().forEach(id => {
        const nav = document.getElementById(id);
        if (!nav) return;

        nav.querySelectorAll('.bottom-nav-btn').forEach(btn => {
            btn.classList.remove('active');
        });
    });
}

function initRoleNavEngine() {
    updateRoleNavVisibility();

    // Ukloni svaki statički active iz HTML-a pa sinhronizuj iz stvarnog state-a
    clearAllRoleNavActiveStates();
    updateRoleNavActive();
}

window.updateRoleNavVisibility = updateRoleNavVisibility;
window.updateRoleNavActive = syncRoleNavActiveFromDom;
window.showRoleNavTab = showRoleNavTab;

window.addEventListener('resize', () => {
    updateRoleNavVisibility();
    syncRoleNavActiveFromDom();
});
