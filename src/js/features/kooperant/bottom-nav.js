// ============================================================
// KOOPERANT — BOTTOM NAV
// Koristi postojeći showTab(tabName, btn)
// ============================================================

function initKoopBottomNav() {
    updateKoopBottomNavVisibility();
    updateKoopBottomNavActive();
}

function updateKoopBottomNavVisibility() {
    const nav = document.getElementById('koopBottomNav');
    if (!nav) return;

    const loginContainer = document.getElementById('loginContainer');
    const appContainer = document.getElementById('appContainer');

    const role = String((CONFIG && CONFIG.USER_ROLE) || '').trim().toLowerCase();
    const isKooperant = role === 'kooperant';

    const loginVisible = !!(
        loginContainer &&
        loginContainer.offsetParent !== null &&
        loginContainer.innerHTML.trim() !== '' &&
        getComputedStyle(loginContainer).display !== 'none'
    );

    const appVisible = !!(
        appContainer &&
        appContainer.offsetParent !== null &&
        getComputedStyle(appContainer).display !== 'none'
    );

    const shouldShow = isKooperant && appVisible && !loginVisible;

    nav.classList.toggle('visible', shouldShow);
    document.body.classList.toggle('has-koop-bottom-nav', shouldShow);

    console.log('BOTTOM NAV', {
        role,
        isKooperant,
        appVisible,
        loginVisible,
        shouldShow,
        navClass: nav.className
    });
}

function showBottomTab(tabName, btn) {
    if (tabName === 'more') {
        qsa('.tab-content').forEach(t => removeClass(t, 'active'));
        qsa('.tab-btn').forEach(b => removeClass(b, 'active'));

        const tabEl = byId('tab-more');
        if (tabEl) addClass(tabEl, 'active');

        updateBottomNavButtons(tabName, btn);
        return;
    }

    showTab(tabName, findLegacyTabBtn(tabName) || btn);
    updateBottomNavButtons(tabName, btn);
}

function updateBottomNavButtons(tabName, btn) {
    document.querySelectorAll('#koopBottomNav .bottom-nav-btn').forEach(b => {
        b.classList.remove('active');
    });

    if (btn && btn.classList) {
        btn.classList.add('active');
        return;
    }

    const activeBtn = document.querySelector(`#koopBottomNav .bottom-nav-btn[data-tab="${tabName}"]`);
    if (activeBtn) activeBtn.classList.add('active');
}

function updateKoopBottomNavActive() {
    const activeTab = document.querySelector('.tab-content.active');
    if (!activeTab) return;

    const id = activeTab.id || '';
    const map = {
        'tab-home': 'home',
        'tab-parcele': 'parcele',
        'tab-agromere': 'agromere',
        'tab-knjigapolja': 'knjigapolja',
        'tab-more': 'more',
        'tab-kartica': 'more',
        'tab-koopinfo': 'more'
    };

    const navTab = map[id];
    if (!navTab) return;

    updateBottomNavButtons(navTab);
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

// Hook na postojeći showTab tok
(function attachKoopBottomNavHook() {
    const originalShowTab = window.showTab;
    if (typeof originalShowTab !== 'function') return;

    window.showTab = function patchedShowTab(tabName, btn) {
        originalShowTab(tabName, btn);
        updateKoopBottomNavActive();
    };
})();

document.addEventListener('DOMContentLoaded', function () {
    initKoopBottomNav();
});
