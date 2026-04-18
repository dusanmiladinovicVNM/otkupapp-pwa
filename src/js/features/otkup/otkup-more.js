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
    } else {
        if (typeof clearSignature === 'function') {
            clearSignature('otkupacSignatureCanvas');
        }
        drawEmptyOtkupacSignaturePlaceholder();
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

        const pad = 16;
        const boxW = cssWidth - pad * 2;
        const boxH = cssHeight - pad * 2;

        const ratio = Math.min(boxW / img.width, boxH / img.height);
        const drawW = img.width * ratio;
        const drawH = img.height * ratio;
        const x = pad + (boxW - drawW) / 2;
        const y = pad + (boxH - drawH) / 2;

        ctx.drawImage(img, x, y, drawW, drawH);
    };
    img.src = dataUrl;
}

function drawEmptyOtkupacSignaturePlaceholder() {
    const canvas = byId('otkupacSignatureCanvas');
    if (!canvas) return;

    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    const dpr = Math.max(window.devicePixelRatio || 1, 1);
    const rect = canvas.getBoundingClientRect();
    const cssWidth = Math.max(1, Math.round(rect.width || canvas.clientWidth || 300));
    const cssHeight = Math.max(1, Math.round(rect.height || canvas.clientHeight || 150));

    if (typeof ctx.resetTransform === 'function') {
        ctx.resetTransform();
    } else {
        ctx.setTransform(1, 0, 0, 1, 0, 0);
    }

    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.scale(dpr, dpr);

    ctx.strokeStyle = 'rgba(0,0,0,0.15)';
    ctx.lineWidth = 1;
    ctx.setLineDash([6, 6]);
    ctx.beginPath();
    ctx.moveTo(18, cssHeight - 36);
    ctx.lineTo(cssWidth - 18, cssHeight - 36);
    ctx.stroke();
    ctx.setLineDash([]);

    ctx.fillStyle = 'rgba(0,0,0,0.35)';
    ctx.font = '12px sans-serif';
    ctx.fillText('Potpiši se ovde', 18, cssHeight - 42);
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
    const canvas = byId('otkupacSignatureCanvas');
    if (!canvas) {
        showToast('Potpis nije dostupan', 'error');
        return;
    }

    try {
        const trimmed = exportTrimmedSignature(canvas);
        if (!trimmed) {
            showToast('Prvo unesite potpis', 'error');
            return;
        }

        localStorage.setItem(getOtkupacSignatureStorageKey(), trimmed);
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

function exportTrimmedSignature(canvas) {
    const ctx = canvas.getContext('2d');
    if (!ctx) return '';

    const width = canvas.width;
    const height = canvas.height;
    const imageData = ctx.getImageData(0, 0, width, height).data;

    let minX = width;
    let minY = height;
    let maxX = -1;
    let maxY = -1;

    for (let y = 0; y < height; y++) {
        for (let x = 0; x < width; x++) {
            const i = (y * width + x) * 4;
            const alpha = imageData[i + 3];

            if (alpha > 10) {
                minX = Math.min(minX, x);
                minY = Math.min(minY, y);
                maxX = Math.max(maxX, x);
                maxY = Math.max(maxY, y);
            }
        }
    }

    if (maxX < 0 || maxY < 0) return '';

    const pad = 24;
    minX = Math.max(0, minX - pad);
    minY = Math.max(0, minY - pad);
    maxX = Math.min(width - 1, maxX + pad);
    maxY = Math.min(height - 1, maxY + pad);

    const cropW = maxX - minX + 1;
    const cropH = maxY - minY + 1;

    const out = document.createElement('canvas');
    out.width = cropW;
    out.height = cropH;

    const outCtx = out.getContext('2d');
    outCtx.clearRect(0, 0, cropW, cropH);
    outCtx.drawImage(canvas, minX, minY, cropW, cropH, 0, 0, cropW, cropH);

    return out.toDataURL('image/png');
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
