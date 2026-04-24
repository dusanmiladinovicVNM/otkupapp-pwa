// ============================================================
// bottom-nav.js
// ============================================================

function findTabBtnByTabName(tabName) {
    return document.querySelector('.tab-btn[data-route="tab"][data-tab="' + tabName + '"]') || null;
}

async function syncKooperantFromMore() {
    if (typeof syncKooperantNow !== 'function') {
        showToast('Sync nije dostupan', 'error');
        return;
    }

    showToast('Pokrećem sinhronizaciju...', 'info');

    try {
        const result = await syncKooperantNow();

        if (!result || result.ok === false || result.success === false) {
            showToast('Greška pri sinhronizaciji', 'error');
            return;
        }

        invalidatePregledCacheSafe();

        if (window.__APP_DEBUG) {
            console.log('syncKooperantNow', result);
        }

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
