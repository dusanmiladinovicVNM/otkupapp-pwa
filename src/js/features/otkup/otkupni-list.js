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
    const vrednostNum = record.kolicina * record.cena;
    const pdvStopa = parseFloat(gv('OtkupPDVStopa')) || 8;
    const pdvIznos = Math.round(vrednostNum * pdvStopa / 100);
    const ukupno = vrednostNum + pdvIznos;

    const savedOtkupacSignature =
        (typeof getSavedOtkupacSignature === 'function')
            ? getSavedOtkupacSignature()
            : '';

    const otkupacSignatureHtml = savedOtkupacSignature
        ? `<img src="${savedOtkupacSignature}" alt="Potpis Otkupca" style="display:block;width:100%;height:80px;object-fit:contain;border:1px solid #ccc;border-radius:6px;background:#fff;">`
        : `<div style="display:flex;align-items:center;justify-content:center;width:100%;height:80px;border:1px dashed #ccc;border-radius:6px;background:#fafaf7;color:#777;font-size:12px;text-align:center;padding:8px;">Potpis Otkupca nije unet u tabu Više</div>`;

    let modal = document.getElementById('otkupniListModal');
    if (!modal) {
        modal = document.createElement('div');
        modal.id = 'otkupniListModal';
        modal.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:white;z-index:9999;overflow-y:auto;';
        document.body.appendChild(modal);
    }

    modal.innerHTML = `
        <div style="padding:16px;max-width:420px;margin:0 auto;font-family:sans-serif;">
            <div style="text-align:center;border-bottom:2px solid #333;padding-bottom:10px;margin-bottom:12px;">
                <div style="font-size:18px;font-weight:700;">${gv('SELLER_NAME')}</div>
                <div style="font-size:12px;color:#666;">${escapeHtml(gv('SELLER_STREET'))}, ${escapeHtml(gv('SELLER_CITY'))} ${escapeHtml(gv('SELLER_POSTAL_CODE'))}</div>
                <div style="font-size:12px;color:#666;">PIB: ${escapeHtml(gv('SELLER_PIB'))} | MB: ${escapeHtml(gv('SELLER_MATICNI_BROJ'))}</div>
                <div style="font-size:12px;color:#666;">TR: ${escapeHtml(gv('SELLER_ACCOUNT'))}</div>
            </div>

            <h2 style="text-align:center;margin-bottom:14px;font-size:20px;">OTKUPNI LIST</h2>

            <div style="background:#f5f5f0;padding:10px;border-radius:8px;margin-bottom:12px;font-size:13px;">
                <div><strong>${koop.Ime || ''} ${koop.Prezime || ''}</strong></div>
                <div>${escapeHtml(koop.Adresa || '')}, ${escapeHtml(koop.Mesto || '')}</div>
                <div>JMBG: ${escapeHtml(koop.JMBG || '________')} | BPG: ${escapeHtml(koop.BPGBroj || '________')}</div>
            </div>

            <table style="width:100%;border-collapse:collapse;font-size:14px;">
                <tr><td style="padding:6px;color:#666;width:40%;">Datum:</td><td style="padding:6px;font-weight:600;">${escapeHtml(record.datum)}</td></tr>
                <tr><td style="padding:6px;color:#666;">Proizvod:</td><td style="padding:6px;">${escapeHtml(record.vrstaVoca)} ${escapeHtml(record.sortaVoca || '')}</td></tr>
                <tr><td style="padding:6px;color:#666;">Klasa:</td><td style="padding:6px;">${escapeHtml(record.klasa)}</td></tr>
                <tr><td style="padding:6px;color:#666;">Količina:</td><td style="padding:6px;font-weight:600;">${record.kolicina} kg</td></tr>
                <tr><td style="padding:6px;color:#666;">Cena:</td><td style="padding:6px;">${record.cena} RSD/kg</td></tr>
                <tr style="border-top:1px solid #ccc;"><td style="padding:6px;color:#666;">Vrednost:</td><td style="padding:6px;font-weight:600;">${vrednostNum.toLocaleString('sr')} RSD</td></tr>
                ${pdvStopa > 0 ? '<tr><td style="padding:6px;color:#666;">PDV naknada (' + pdvStopa + '%):</td><td style="padding:6px;">' + pdvIznos.toLocaleString('sr') + ' RSD</td></tr>' : ''}
                <tr style="border-top:2px solid #333;"><td style="padding:8px;color:#666;">ZA ISPLATU:</td><td style="padding:8px;font-weight:700;font-size:18px;">${ukupno.toLocaleString('sr')} RSD</td></tr>
                <tr><td style="padding:6px;color:#666;">Ambalaža:</td><td style="padding:6px;">${record.kolAmbalaze} kom</td></tr>
                ${record.parcelaID ? '<tr><td style="padding:6px;color:#666;">Parcela:</td><td style="padding:6px;">' + escapeHtml(record.parcelaID) + '</td></tr>' : ''}
                <tr><td style="padding:6px;color:#666;">Rok isplate:</td><td style="padding:6px;">${escapeHtml(gv('OtkupRokIsplate') || 'Po dogovoru')}</td></tr>
            </table>

            <div style="margin-top:20px;">
                <div style="margin-bottom:16px;">
                    <div style="font-size:12px;color:#666;margin-bottom:4px;">Potpis otkupljivača:</div>
                    ${otkupacSignatureHtml}
                </div>

                <div style="margin-bottom:16px;">
                    <div style="font-size:12px;color:#666;margin-bottom:4px;">Potpis kooperanta:</div>
                    <canvas id="sigKooperant" width="720" height="200" style="border:1px solid #ccc;border-radius:6px;width:100%;height:80px;touch-action:none;"></canvas>
                </div>
            </div>

            <div style="text-align:center;margin-top:16px;display:flex;gap:8px;">
                <button onclick="clearSignature('sigKooperant')" style="flex:1;padding:12px;font-size:14px;background:#f5f5f0;color:#666;border:1px solid #ccc;border-radius:8px;">Obriši</button>
                <button onclick="saveOtkupniListWithSignatures('${record.clientRecordID}')" style="flex:1;padding:12px;font-size:14px;background:var(--primary);color:white;border:none;border-radius:8px;">Potvrdi</button>
                <button onclick="window.print()" style="flex:1;padding:12px;font-size:14px;background:var(--accent);color:white;border:none;border-radius:8px;">Štampaj</button>
            </div>

            <button onclick="savePdfToDrive('${record.clientRecordID}')" style="width:100%;padding:12px;margin-top:8px;font-size:14px;background:#2196F3;color:white;border:none;border-radius:8px;">📄 Sačuvaj PDF na Drive</button>
            <button onclick="closeOtkupniListModal()" style="width:100%;padding:10px;margin-top:8px;font-size:14px;background:none;color:#666;border:1px solid #ccc;border-radius:8px;">Zatvori</button>
        </div>
    `;

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

    console.log('SIG OTK length:', sigOtkupac ? sigOtkupac.length : 0);
    console.log('SIG KOOP length:', sigKooperant ? sigKooperant.length : 0);

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
            console.log('Adding sigOtkupac to PDF');
            try { doc.addImage(sigOtkupac, 'PNG', 13, y + 1, sigW - 2, sigH - 2); console.log('sigOtk OK'); } catch (e) { console.log('sigOtk ERROR:', e); }
        }
        if (sigKooperant) {
            console.log('Adding sigKooperant to PDF');
            try { doc.addImage(sigKooperant, 'PNG', 18 + sigW, y + 1, sigW - 2, sigH - 2); console.log('sigKoop OK'); } catch (e) { console.log('sigKoop ERROR:', e); }
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
        console.log('PDF error:', e);
        showToast('Greška pri generisanju PDF-a', 'error');
    }
}

function closeOtkupniListModal() {
    destroySignaturePad('sigKooperant');
    const modal = document.getElementById('otkupniListModal');
    if (modal) modal.style.display = 'none';
}
