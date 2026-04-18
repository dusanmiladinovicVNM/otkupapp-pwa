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
    if (canvas) {
        initOtkupacSignaturePad(canvas);
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
    if (badge) {
        badge.textContent = hasSavedOtkupacSignature() ? 'Sačuvan' : 'Nije unet';
        toggleClass(badge, 'is-ready', hasSavedOtkupacSignature());
    }

    const canvas = byId('otkupacSignatureCanvas');
    if (!canvas) return;

    resizeOtkupacSignatureCanvas();

    const saved = getSavedOtkupacSignature();
    if (saved) {
        drawSignatureImageToCanvas(saved);
    } else {
        clearSignatureCanvasVisualOnly();
    }

    otkupacMoreState.signatureDirty = false;
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

function initOtkupacSignaturePad(canvas) {
    if (!canvas || otkupacMoreState.signaturePadBound) return;

    otkupacMoreState.signatureCanvas = canvas;
    canvas.style.touchAction = 'none';

    const begin = (e) => {
        const point = getCanvasPoint(canvas, e);
        const ctx = otkupacMoreState.signatureCtx;
        if (!ctx) return;

        ctx.beginPath();
        ctx.moveTo(point.x, point.y);
        canvas.dataset.drawing = '1';
        otkupacMoreState.signatureDirty = true;
    };

    const move = (e) => {
        if (canvas.dataset.drawing !== '1') return;
        const point = getCanvasPoint(canvas, e);
        const ctx = otkupacMoreState.signatureCtx;
        if (!ctx) return;

        ctx.lineTo(point.x, point.y);
        ctx.stroke();
    };

    const end = () => {
        canvas.dataset.drawing = '0';
    };

    canvas.addEventListener('pointerdown', begin);
    canvas.addEventListener('pointermove', move);
    canvas.addEventListener('pointerup', end);
    canvas.addEventListener('pointerleave', end);
    canvas.addEventListener('pointercancel', end);

    window.addEventListener('resize', debounceOtkupacSignatureResize);

    otkupacMoreState.signaturePadBound = true;
}

const debounceOtkupacSignatureResize = debounce(function () {
    const activeTab = document.querySelector('#tab-queue.tab-content.active');
    if (!activeTab) return;
    renderOtkupacSignaturePad();
}, 150);

function resizeOtkupacSignatureCanvas() {
    const canvas = otkupacMoreState.signatureCanvas || byId('otkupacSignatureCanvas');
    if (!canvas) return;

    const dpr = Math.max(window.devicePixelRatio || 1, 1);
    const cssWidth = Math.max(canvas.clientWidth || 320, 320);
    const cssHeight = 210;

    const previous = canvas.toDataURL('image/png');

    canvas.width = Math.floor(cssWidth * dpr);
    canvas.height = Math.floor(cssHeight * dpr);
    canvas.style.height = cssHeight + 'px';

    const ctx = canvas.getContext('2d');
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.lineWidth = 2.2;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.strokeStyle = '#1f3f1f';

    otkupacMoreState.signatureCtx = ctx;
    otkupacMoreState.signatureCanvas = canvas;

    clearSignatureCanvasVisualOnly();

    const saved = getSavedOtkupacSignature();
    if (saved) {
        drawSignatureImageToCanvas(saved);
    } else if (previous && previous !== 'data:,') {
        drawSignatureImageToCanvas(previous);
    }
}

function clearSignatureCanvasVisualOnly() {
    const canvas = otkupacMoreState.signatureCanvas;
    const ctx = otkupacMoreState.signatureCtx;
    if (!canvas || !ctx) return;

    ctx.clearRect(0, 0, canvas.width, canvas.height);

    ctx.save();
    ctx.setTransform(1, 0, 0, 1, 0, 0);

    const cssWidth = canvas.clientWidth || 320;
    const cssHeight = 210;

    ctx.strokeStyle = 'rgba(0,0,0,0.12)';
    ctx.lineWidth = 1;
    ctx.strokeRect(0.5, 0.5, cssWidth - 1, cssHeight - 1);

    ctx.setLineDash([6, 6]);
    ctx.beginPath();
    ctx.moveTo(18, cssHeight - 36);
    ctx.lineTo(cssWidth - 18, cssHeight - 36);
    ctx.stroke();

    ctx.setLineDash([]);
    ctx.fillStyle = 'rgba(0,0,0,0.35)';
    ctx.font = '12px sans-serif';
    ctx.fillText('Potpis Otkupca', 18, cssHeight - 42);

    ctx.restore();
}

function drawSignatureImageToCanvas(dataUrl) {
    const canvas = otkupacMoreState.signatureCanvas;
    const ctx = otkupacMoreState.signatureCtx;
    if (!canvas || !ctx || !dataUrl) return;

    const img = new Image();
    img.onload = function () {
        clearSignatureCanvasVisualOnly();

        const cssWidth = canvas.clientWidth || 320;
        const cssHeight = 210;
        const pad = 14;

        const drawWidth = cssWidth - pad * 2;
        const drawHeight = cssHeight - pad * 2;

        ctx.drawImage(img, pad, pad, drawWidth, drawHeight);
    };
    img.src = dataUrl;
}

function saveOtkupacSignature() {
    const canvas = byId('otkupacSignatureCanvas');
    if (!canvas) return;

    try {
        const dataUrl = canvas.toDataURL('image/png');
        localStorage.setItem(getOtkupacSignatureStorageKey(), dataUrl);
        otkupacMoreState.signatureDirty = false;
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
        otkupacMoreState.signatureDirty = false;
        renderOtkupacSignaturePad();
        showToast('Potpis je obrisan', 'info');
    } catch (err) {
        console.error(err);
        showToast('Greška pri brisanju potpisa', 'error');
    }
}

function getCanvasPoint(canvas, event) {
    const rect = canvas.getBoundingClientRect();
    return {
        x: event.clientX - rect.left,
        y: event.clientY - rect.top
    };
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
