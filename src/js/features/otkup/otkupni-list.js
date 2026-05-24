// ============================================================
// OTKUPNI LIST + SIGNATURE PAD
// ============================================================

let _otpremaServerCache = { data: null, ts: 0 };
const OTPREMA_CACHE_TTL = 15000; // 15s

function showOtkupniList(record) {
    const config = stammdaten.config || [];
    const gv = k => {
        const c = config.find(c => c.Parameter === k);
        return c ? c.Vrednost : '';
    };

    const koop = (stammdaten.kooperanti || []).find(k => k.KooperantID === record.kooperantID) || {};
    const koopFullName = ((koop.Ime || '') + ' ' + (koop.Prezime || '')).trim()
        || record.kooperantName || 'Kooperant';
    const koopInitials = ((koop.Ime || 'K').charAt(0) + (koop.Prezime || '').charAt(0)).toUpperCase();

    const vrednostNum = record.kolicina * record.cena;
    const pdvStopa = parseFloat(gv('OtkupPDVStopa')) || 8;
    const pdvIznos = Math.round(vrednostNum * pdvStopa / 100);
    const ukupno = vrednostNum + pdvIznos;
    const ukupnoFormatted = ukupno.toLocaleString('sr-RS');

    const otkupBroj = record.brojDokumenta
        || record.serverRecordID
        || (record.clientRecordID ? 'OL-' + String(record.clientRecordID).slice(0, 8) : 'OL');

    const datumFormatted = (function() {
        if (!record.datum) return '';
        try {
            const d = new Date(record.datum);
            if (isNaN(d.getTime())) return record.datum;
            return d.toLocaleDateString('sr-RS', { day: 'numeric', month: 'long', year: 'numeric' });
        } catch (_) { return record.datum; }
    })();

    const savedOtkupacSignature = (typeof getSavedOtkupacSignature === 'function')
        ? getSavedOtkupacSignature()
        : '';

    let modal = document.getElementById('otkupniListModal');
    if (!modal) {
        modal = document.createElement('div');
        modal.id = 'otkupniListModal';
        modal.className = 'ol-modal';
        document.body.appendChild(modal);
    } else {
        modal.className = 'ol-modal';
    }

    modal.innerHTML = `
        <div class="ol-shell">

            <!-- Forest success header -->
            <div class="ol-hd">
                <div class="ol-hd__row">
                    <div class="ol-hd__role">Otkupac · ${escapeHtml(gv('STATION_NAME') || CONFIG.ENTITY_NAME || '')}</div>
                    <button type="button" class="ol-hd__close" data-action="close-otkupni-list-modal" aria-label="Zatvori">×</button>
                </div>
                <div class="ol-hd__success">
                    <div class="ol-hd__check">✓</div>
                    <div class="ol-hd__success-text">
                        <div class="ol-hd__title">Otkup sačuvan</div>
                        <div class="ol-hd__sub">${escapeHtml(koopFullName)} · ${record.kolicina} kg · Klasa ${escapeHtml(record.klasa || '')}</div>
                    </div>
                </div>
                <div class="ol-hd__date">${escapeHtml(datumFormatted)}</div>
            </div>

            <div class="ol-body">

                <!-- Summary accent card -->
                <div class="card card--accent ol-summary">
                    <div class="ol-summary__head">
                        <div class="ol-summary__eyebrow">Otkupni list spreman</div>
                        <div class="ol-summary__kg">
                            <span class="ol-summary__kg-val">${record.kolicina}</span><span class="ol-summary__kg-unit">kg</span>
                        </div>
                    </div>
                    <div class="ol-summary__broj">${escapeHtml(otkupBroj)}</div>
                </div>

                <!-- Info grid 2x2 -->
                <div class="ol-info-grid">
                    <div class="ol-info">
                        <div class="ol-info__label">Kooperant</div>
                        <div class="ol-info__value">${escapeHtml(koopFullName)}</div>
                    </div>
                    <div class="ol-info">
                        <div class="ol-info__label">Vrednost</div>
                        <div class="ol-info__value">${ukupnoFormatted} RSD</div>
                    </div>
                    <div class="ol-info">
                        <div class="ol-info__label">Klasa</div>
                        <div class="ol-info__value">${escapeHtml(record.klasa || '')}</div>
                    </div>
                    <div class="ol-info">
                        <div class="ol-info__label">Ambalaža</div>
                        <div class="ol-info__value">${record.kolAmbalaze || 0} × ${escapeHtml(record.tipAmbalaze || '')}</div>
                    </div>
                </div>

                <!-- Receipt details (expandable) -->
                <details class="ol-receipt">
                    <summary class="ol-receipt__summary">Detalji otkupnog lista</summary>
                    <div class="ol-receipt__body">
                        <div class="ol-receipt__seller">
                            <strong>${escapeHtml(gv('SELLER_NAME'))}</strong>
                            <div>${escapeHtml(gv('SELLER_STREET'))}, ${escapeHtml(gv('SELLER_CITY'))} ${escapeHtml(gv('SELLER_POSTAL_CODE'))}</div>
                            <div>PIB: ${escapeHtml(gv('SELLER_PIB'))} · MB: ${escapeHtml(gv('SELLER_MATICNI_BROJ'))}</div>
                            <div>TR: ${escapeHtml(gv('SELLER_ACCOUNT'))}</div>
                        </div>
                        <div class="ol-receipt__koop">
                            <strong>${escapeHtml(koopFullName)}</strong>
                            <div>${escapeHtml(koop.Adresa || '')}, ${escapeHtml(koop.Mesto || '')}</div>
                            <div>JMBG: ${escapeHtml(koop.JMBG || '________')} · BPG: ${escapeHtml(koop.BPGBroj || '________')}</div>
                        </div>
                        <table class="ol-receipt__table">
                            <tr><td>Datum</td><td>${escapeHtml(datumFormatted)}</td></tr>
                            <tr><td>Proizvod</td><td>${escapeHtml(record.vrstaVoca)} ${escapeHtml(record.sortaVoca || '')}</td></tr>
                            <tr><td>Cena</td><td>${record.cena} RSD/kg</td></tr>
                            <tr><td>Vrednost</td><td>${vrednostNum.toLocaleString('sr-RS')} RSD</td></tr>
                            ${pdvStopa > 0 ? `<tr><td>PDV naknada (${pdvStopa}%)</td><td>${pdvIznos.toLocaleString('sr-RS')} RSD</td></tr>` : ''}
                            <tr class="ol-receipt__total"><td>Za isplatu</td><td><strong>${ukupnoFormatted} RSD</strong></td></tr>
                            ${record.parcelaID ? `<tr><td>Parcela</td><td>${escapeHtml(record.parcelaID)}</td></tr>` : ''}
                            <tr><td>Rok isplate</td><td>${escapeHtml(gv('OtkupRokIsplate') || 'Po dogovoru')}</td></tr>
                        </table>
                    </div>
                </details>

                <!-- Kooperant signature pad -->
                <div class="ol-signature">
                    <div class="ol-signature__head">
                        <div class="ol-signature__label">Potpis kooperanta</div>
                        <button type="button" class="ol-signature__clear" data-action="otkupni-clear-signature">Obriši</button>
                    </div>
                    <canvas id="sigKooperant" width="720" height="200" class="ol-signature__canvas"></canvas>
                </div>

                <!-- Saved otkupac signature -->
                ${savedOtkupacSignature
                    ? `<div class="ol-signature ol-signature--saved">
                           <div class="ol-signature__label">Potpis otkupca</div>
                           <div class="ol-signature__saved-wrap">
                               <img src="${savedOtkupacSignature}" alt="Potpis otkupca" class="ol-signature__saved-img">
                           </div>
                       </div>`
                    : `<div class="ol-signature ol-signature--missing">
                           <div class="ol-signature__warn">⚠ Potpis otkupca nije unet u tabu Više</div>
                       </div>`}
            </div>

            <!-- Sticky action bar -->
            <div class="ol-actions">
                <button type="button" class="btn-v2 btn-v2--primary" data-action="otkupni-confirm"
                        data-client-record-id="${escapeHtml(String(record.clientRecordID || ''))}">
                    Potvrdi i sačuvaj
                </button>
                <div class="ol-actions__row">
                    <button type="button" class="btn-v2 btn-v2--secondary" data-action="otkupni-print">
                        Štampaj
                    </button>
                    <button type="button" class="btn-v2 btn-v2--secondary" data-action="otkupni-save-pdf"
                            data-client-record-id="${escapeHtml(String(record.clientRecordID || ''))}">
                        Sačuvaj PDF
                    </button>
                </div>
            </div>
        </div>
    `;

    if (!modal.dataset.bound) {
        modal.addEventListener('click', function (event) {
            const actionEl = event.target.closest('[data-action]');
            if (!actionEl) return;

            const action = actionEl.dataset.action;

            if (action === 'otkupni-clear-signature') {
                clearSignature('sigKooperant');
                return;
            }

            if (action === 'otkupni-confirm') {
                saveOtkupniListWithSignatures(actionEl.dataset.clientRecordId || '');
                return;
            }

            if (action === 'otkupni-print') {
                window.print();
                return;
            }

            if (action === 'otkupni-save-pdf') {
                savePdfToDrive(actionEl.dataset.clientRecordId || '');
                return;
            }

            if (action === 'close-otkupni-list-modal') {
                closeOtkupniListModal();
            }
        });

        modal.dataset.bound = '1';
    }

    modal.style.display = 'block';

    setTimeout(() => {
        initSignaturePad('sigKooperant');
    }, 100);
}

async function saveOtkupniListWithSignatures(clientRecordID) {
    const sigK = getSignatureData('sigKooperant');

    if (!sigK) {
        showToast('Kooperant mora da se potpiše!', 'error');
        return;
    }

    try {
        const r = await dbGet(db, CONFIG.STORE_NAME, clientRecordID);

        if (!r) {
            showToast('Zapis nije pronađen', 'error');
            return;
        }

        r.sigKooperant = sigK;
        r.signedAt = new Date().toISOString();

        await dbPut(db, CONFIG.STORE_NAME, r);

        showToast('Otkupni list potpisan!', 'success');

        const modal = document.getElementById('otkupniListModal');
        if (modal) modal.style.display = 'none';
    } catch (e) {
        console.error('saveOtkupniListWithSignatures failed:', e);
        showToast('Greška pri čuvanju potpisa', 'error');
    }
}

async function savePdfToDrive(clientRecordID) {
    const record = await dbGet(db, CONFIG.STORE_NAME, clientRecordID);
    if (!record) { showToast('Zapis nije pronađen', 'error'); return; }

    const config = stammdaten.config || [];
    const gv = k => { const c = config.find(c => c.Parameter === k); return c ? c.Vrednost : ''; };
    const koop = (stammdaten.kooperanti || []).find(k => k.KooperantID === record.kooperantID) || {};
    const vrednostNum = record.kolicina * record.cena;
    const pdvStopa = parseFloat(gv('OtkupPDVStopa')) || 8;
    const pdvIznos = Math.round(vrednostNum * pdvStopa / 100);
    const ukupno = vrednostNum + pdvIznos;

    const sigOtkupac =
        (typeof getSavedOtkupacSignature === 'function' && getSavedOtkupacSignature())
            ? getSavedOtkupacSignature()
            : (record.sigOtkupac || '');

    const sigKooperant = getSignatureData('sigKooperant') || (record.sigKooperant || '');

    showToast('Generisanje PDF-a...', 'info');

    try {
        const jsPDF = (window.jspdf && window.jspdf.jsPDF) || window.jsPDF;
        if (!jsPDF) { showToast('PDF biblioteka nije učitana', 'error'); return; }
        const doc = new jsPDF({ format: 'a5', unit: 'mm' });
        const w = doc.internal.pageSize.getWidth();
        let y = 10;

        doc.setFontSize(13);
        doc.setFont(undefined, 'bold');
        doc.text(gv('SELLER_NAME'), w / 2, y, { align: 'center' });
        y += 5;
        doc.setFontSize(8);
        doc.setFont(undefined, 'normal');
        doc.text(gv('SELLER_STREET') + ', ' + gv('SELLER_CITY') + ' ' + gv('SELLER_POSTAL_CODE'), w / 2, y, { align: 'center' });
        y += 4;
        doc.text('PIB: ' + gv('SELLER_PIB') + ' | MB: ' + gv('SELLER_MATICNI_BROJ'), w / 2, y, { align: 'center' });
        y += 4;
        doc.text('TR: ' + gv('SELLER_ACCOUNT'), w / 2, y, { align: 'center' });
        y += 3;
        doc.setLineWidth(0.5);
        doc.line(10, y, w - 10, y);
        y += 6;

        doc.setFontSize(14);
        doc.setFont(undefined, 'bold');
        doc.text('OTKUPNI LIST', w / 2, y, { align: 'center' });
        y += 7;

        doc.setFillColor(240, 240, 234);
        doc.rect(10, y, w - 20, 14, 'F');
        doc.setFontSize(10);
        doc.setFont(undefined, 'bold');
        doc.text((koop.Ime || '') + ' ' + (koop.Prezime || ''), 12, y + 4);
        doc.setFontSize(8);
        doc.setFont(undefined, 'normal');
        doc.text((koop.Adresa || '') + ', ' + (koop.Mesto || ''), 12, y + 8);
        doc.text('JMBG: ' + (koop.JMBG || '________') + '  |  BPG: ' + (koop.BPGBroj || '________'), 12, y + 12);
        y += 18;

        const lx = 12;
        const vx = 60;
        doc.setFontSize(9);

        function addRow(label, value, bold, line) {
            if (line) { doc.setLineWidth(0.3); doc.line(lx, y, w - 12, y); y += 1; }
            doc.setFont(undefined, 'normal');
            doc.setTextColor(100);
            doc.text(label, lx, y + 4);
            doc.setTextColor(0);
            if (bold) doc.setFont(undefined, 'bold');
            doc.text(String(value), vx, y + 4);
            doc.setFont(undefined, 'normal');
            y += 6;
        }

        addRow('Datum:', record.datum, true, false);
        addRow('Proizvod:', record.vrstaVoca + ' ' + (record.sortaVoca || ''), false, false);
        addRow('Klasa:', record.klasa, false, false);
        addRow('Količina:', record.kolicina + ' kg', true, false);
        addRow('Cena:', record.cena + ' RSD/kg', false, false);
        addRow('Vrednost:', vrednostNum.toLocaleString('sr') + ' RSD', true, true);
        if (pdvStopa > 0) {
            addRow('PDV naknada (' + pdvStopa + '%):', pdvIznos.toLocaleString('sr') + ' RSD', false, false);
        }

        doc.setLineWidth(0.5);
        doc.line(lx, y, w - 12, y);
        y += 1;
        doc.setFontSize(11);
        doc.setFont(undefined, 'bold');
        doc.setTextColor(100);
        doc.text('ZA ISPLATU:', lx, y + 5);
        doc.setTextColor(0);
        doc.text(ukupno.toLocaleString('sr') + ' RSD', vx, y + 5);
        y += 8;
        doc.setFontSize(9);
        doc.setFont(undefined, 'normal');

        addRow('Ambalaža:', record.kolAmbalaze + ' kom', false, false);
        if (record.parcelaID) addRow('Parcela:', record.parcelaID, false, false);
        addRow('Rok isplate:', gv('OtkupRokIsplate') || 'Po dogovoru', false, false);

        y += 4;

        const sigW = (w - 30) / 2;
        const sigH = 20;

        doc.setFontSize(7);
        doc.setTextColor(100);
        doc.text('Potpis otkupljivača:', 12, y);
        doc.text('Potpis kooperanta:', 17 + sigW, y);
        y += 2;

        doc.setDrawColor(200);
        doc.rect(12, y, sigW, sigH);
        doc.rect(17 + sigW, y, sigW, sigH);

        if (sigOtkupac) {
            try {
                doc.addImage(sigOtkupac, 'PNG', 13, y + 1, sigW - 2, sigH - 2);
            } catch (e) {
                if (typeof reportClientError === 'function') {
                    reportClientError(e, {
                        source: 'otkupni-list',
                        errorAction: 'pdf-add-signature-otkupac'
                    });
                }
            }
        }

        if (sigKooperant) {
            try {
                doc.addImage(sigKooperant, 'PNG', 18 + sigW, y + 1, sigW - 2, sigH - 2);
            } catch (e) {
                if (typeof reportClientError === 'function') {
                    reportClientError(e, {
                        source: 'otkupni-list',
                        errorAction: 'pdf-add-signature-kooperant'
                    });
                }
            }
        }

        y += sigH + 5;
        doc.setFontSize(6);
        doc.setTextColor(150);
        doc.text('Generisano: ' + new Date().toISOString().substring(0, 19).replace('T', ' '), w / 2, y, { align: 'center' });

        const pdfBase64 = doc.output('datauristring').split(',')[1];
        const fileName = 'OtkupniList_' + record.kooperantID + '_' + record.datum + '_' + clientRecordID.substring(0, 8);

        const json = await apiPost('uploadPdf', {
            fileName: fileName,
            pdfBase64: pdfBase64
        });
        
        if (json.success) { showToast('PDF sačuvan na Drive!', 'success'); }
        else { showToast('Greška: ' + (json.error || ''), 'error'); }
    } catch (e) {
        console.error('PDF error:', e);

        if (typeof reportClientError === 'function') {
            reportClientError(e, {
                source: 'otkupni-list',
                errorAction: 'savePdfToDrive'
            });
        }

        showToast('Greška pri generisanju PDF-a', 'error');
    }
}

function closeOtkupniListModal() {
    destroySignaturePad('sigKooperant');
    const modal = document.getElementById('otkupniListModal');
    if (modal) modal.style.display = 'none';
}
