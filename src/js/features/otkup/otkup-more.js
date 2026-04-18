// ============================================================
// OTKUPAC - VISE TAB
// Potpis Otkupca + sync status
// ============================================================

const otkupacMoreState = {
    uiBound: false,
    signaturePadBound: false,
    signatureDirty: false,
    signatureCtx: null,
    signatureCanvas: null
};

function getOtkupacSignatureStorageKey() {
    return 'otkupac-signature:' + String((CONFIG && CONFIG.ENTITY_ID) || 'unknown');
}

function getSavedOtkupacSignature() {
    try {
        return localStorage.getItem(getOtkupacSignatureStorageKey()) || '';
    } catch (_) {
        return '';
    }
}

function hasSavedOtkupacSignature() {
    return !!getSavedOtkupacSignature();
}

async function loadOtkupacMore() {
    bindOtkupacMoreEventsOnce();
    renderOtkupacMoreProfile();
    renderOtkupacSignaturePad();
    await renderOtkupacMoreSyncStats();
}

function bindOtkupacMoreEventsOnce() {
    if (otkupacMoreState.uiBound) return;

    const canvas = byId('otkupacSignatureCanvas');
    if (canvas && typeof initSignaturePad === 'function') {
        initSignaturePad('otkupacSignatureCanvas');
    }

    otkupacMoreState.uiBound = true;
}


function renderOtkupacMoreProfile() {
    setText(byId('otkMoreProfileName'), (CONFIG && CONFIG.ENTITY_NAME) || '-');
    setText(byId('otkMoreProfileRole'), (CONFIG && CONFIG.USER_ROLE) || '-');
    setText(byId('otkMoreProfileEntity'), fmtStanica((CONFIG && CONFIG.ENTITY_ID) || '') || ((CONFIG && CONFIG.ENTITY_ID) || '-'));
}

function renderOtkupacSignaturePad() {
    const badge = byId('otkMoreSignatureBadge');
    const saved = getSavedOtkupacSignature();

    if (badge) {
        badge.textContent = saved ? 'Sačuvan' : 'Nije unet';
        toggleClass(badge, 'is-ready', !!saved);
    }

    const canvas = byId('otkupacSignatureCanvas');
    if (!canvas || typeof initSignaturePad !== 'function') return;

    initSignaturePad('otkupacSignatureCanvas');

    if (saved) {
        drawSavedOtkupacSignature(saved);
    } else if (typeof clearSignature === 'function') {
        clearSignature('otkupacSignatureCanvas');
    }

    otkupacMoreState.signatureDirty = false;
}

function drawSavedOtkupacSignature(dataUrl) {
    const canvas = byId('otkupacSignatureCanvas');
    if (!canvas || !dataUrl) return;

    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    const dpr = Math.max(window.devicePixelRatio || 1, 1);
    const rect = canvas.getBoundingClientRect();
    const cssWidth = Math.max(1, Math.round(rect.width || canvas.clientWidth || 300));
    const cssHeight = Math.max(1, Math.round(rect.height || canvas.clientHeight || 150));

    const img = new Image();
    img.onload = function () {
        if (typeof ctx.resetTransform === 'function') {
            ctx.resetTransform();
        } else {
            ctx.setTransform(1, 0, 0, 1, 0, 0);
        }

        ctx.clearRect(0, 0, canvas.width, canvas.height);
        ctx.scale(dpr, dpr);

        const pad = 12;
        ctx.drawImage(img, pad, pad, cssWidth - pad * 2, cssHeight - pad * 2);

        ctx.strokeStyle = '#1a1a1a';
        ctx.lineWidth = 2;
        ctx.lineCap = 'round';
        ctx.lineJoin = 'round';
    };
    img.src = dataUrl;
}

async function renderOtkupacMoreSyncStats() {
    const listEl = byId('otkMoreQueueList');
    if (listEl) {
        setHtml(listEl, '<div class="otk-more-empty">Učitavanje sync statusa...</div>');
    }

    let rows = [];
    try {
        if (db) {
            rows = await dbGetAll(db, CONFIG.STORE_NAME);
        }
    } catch (err) {
        console.error('renderOtkupacMoreSyncStats failed:', err);
    }

    rows = (rows || []).filter(r => !r.deleted);

    const pending = rows.filter(r => (r.syncStatus || 'pending') !== 'synced');
    const errors = rows.filter(r => !!String(r.lastSyncError || '').trim());
    const synced = rows.filter(r => (r.syncStatus || '') === 'synced');

    setText(byId('otkMorePendingCount'), String(pending.length));
    setText(byId('otkMoreErrorCount'), String(errors.length));
    setText(byId('otkMoreSyncedCount'), String(synced.length));

    const lastSync = getLastSyncStamp(rows);
    setText(
        byId('otkMoreLastSync'),
        'Poslednja sinhronizacija: ' + (lastSync ? formatOtkMoreDateTime(lastSync) : '—')
    );

    renderOtkupacMoreQueueList(pending, errors);
}

function renderOtkupacMoreQueueList(pending, errors) {
    const listEl = byId('otkMoreQueueList');
    if (!listEl) return;

    if (!pending.length && !errors.length) {
        setHtml(listEl, '<div class="otk-more-empty">Nema lokalnih stavki koje čekaju sinhronizaciju.</div>');
        return;
    }

    const unique = new Map();

    [...errors, ...pending].forEach(r => {
        const key = getOtkMoreRecordKey(r);
        if (!key) return;
        if (!unique.has(key)) unique.set(key, r);
    });

    const rows = Array.from(unique.values()).sort((a, b) => {
        const aTime = String(a.updatedAtClient || a.createdAtClient || '');
        const bTime = String(b.updatedAtClient || b.createdAtClient || '');
        return bTime.localeCompare(aTime);
    });

    setHtml(listEl, rows.map(renderOtkMoreQueueCard).join(''));
}

function renderOtkMoreQueueCard(row) {
    const hasError = !!String(row.lastSyncError || '').trim();

    return `
        <div class="otk-more-queue-card">
            <div class="otk-more-queue-top">
                <div class="otk-more-queue-title">${escapeHtml(row.kooperantName || row.kooperantID || 'Otkup')}</div>
                <span class="otk-more-queue-badge ${hasError ? 'is-error' : 'is-pending'}">
                    ${hasError ? 'Greška' : 'Na čekanju'}
                </span>
            </div>

            <div class="otk-more-queue-line">
                ${escapeHtml(row.vrstaVoca || '-')}
                ${row.sortaVoca ? ' / ' + escapeHtml(row.sortaVoca) : ''}
                <span class="otk-more-queue-muted"> • ${escapeHtml(formatOtkMoreKg(row.kolicina || 0))}</span>
            </div>

            <div class="otk-more-queue-line otk-more-queue-muted">
                ${escapeHtml(fmtDate(row.datum || ''))}
                ${row.vozacID ? ' • Vozač ' + escapeHtml(row.vozacID) : ''}
            </div>

            ${hasError ? `<div class="otk-more-queue-error">${escapeHtml(String(row.lastSyncError || ''))}</div>` : ''}
        </div>
    `;
}

async function syncOtkupacFromMore() {
    if (typeof syncQueueSafe !== 'function') {
        showToast('Sync nije dostupan', 'error');
        return;
    }

    showToast('Pokrećem sinhronizaciju...', 'info');

    try {
        await syncQueueSafe();
        await renderOtkupacMoreSyncStats();
        showToast('Sinhronizacija završena', 'success');
    } catch (err) {
        console.error(err);
        showToast('Greška pri sinhronizaciji', 'error');
    }
}

function saveOtkupacSignature() {
    if (typeof getSignatureData !== 'function') {
        showToast('Potpis nije dostupan', 'error');
        return;
    }

    try {
        const dataUrl = getSignatureData('otkupacSignatureCanvas');
        if (!dataUrl) {
            showToast('Prvo unesite potpis', 'error');
            return;
        }

        localStorage.setItem(getOtkupacSignatureStorageKey(), dataUrl);
        renderOtkupacSignaturePad();
        showToast('Potpis je sačuvan', 'success');
    } catch (err) {
        console.error(err);
        showToast('Greška pri čuvanju potpisa', 'error');
    }
}

function clearOtkupacSignature() {
    try {
        localStorage.removeItem(getOtkupacSignatureStorageKey());

        if (typeof clearSignature === 'function') {
            clearSignature('otkupacSignatureCanvas');
        }

        renderOtkupacSignaturePad();
        showToast('Potpis je obrisan', 'info');
    } catch (err) {
        console.error(err);
        showToast('Greška pri brisanju potpisa', 'error');
    }
}

function getLastSyncStamp(rows) {
    const stamps = (rows || [])
        .map(r => r.syncedAt || r.updatedAtServer || '')
        .filter(Boolean)
        .sort();

    return stamps.length ? stamps[stamps.length - 1] : '';
}

function getOtkMoreRecordKey(row) {
    if (row.serverRecordID) return 'srv:' + row.serverRecordID;
    if (row.clientRecordID) return 'cli:' + row.clientRecordID;
    return '';
}

function formatOtkMoreKg(value) {
    return (parseFloat(value) || 0).toLocaleString('sr-RS') + ' kg';
}

function formatOtkMoreDateTime(value) {
    if (!value) return '—';
    try {
        const d = new Date(value);
        if (isNaN(d.getTime())) return String(value);
        return d.toLocaleString('sr-RS');
    } catch (_) {
        return String(value);
    }
}

function debounce(fn, wait) {
    let t = null;
    return function () {
        const args = arguments;
        clearTimeout(t);
        t = setTimeout(() => fn.apply(this, args), wait);
    };
}
