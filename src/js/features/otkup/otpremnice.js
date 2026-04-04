// ============================================================
// OTKUPNI LIST + SIGNATURE PAD
// ============================================================
function showOtkupniList(record) {
    const config = stammdaten.config || [];
    const gv = k => { const c = config.find(c => c.Parameter === k); return c ? c.Vrednost : ''; };
    const koop = (stammdaten.kooperanti || []).find(k => k.KooperantID === record.kooperantID) || {};
    const vrednostNum = record.kolicina * record.cena;
    const pdvStopa = parseFloat(gv('OtkupPDVStopa')) || 8;
    const pdvIznos = Math.round(vrednostNum * pdvStopa / 100);
    const ukupno = vrednostNum + pdvIznos;

    let modal = document.getElementById('otkupniListModal');
    if (!modal) { modal = document.createElement('div'); modal.id = 'otkupniListModal'; modal.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:white;z-index:9999;overflow-y:auto;'; document.body.appendChild(modal); }

    modal.innerHTML = `<div style="padding:16px;max-width:420px;margin:0 auto;font-family:sans-serif;">
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
            <div style="margin-bottom:16px;"><div style="font-size:12px;color:#666;margin-bottom:4px;">Potpis otkupljivača:</div><canvas id="sigOtkupac" width="720" height="200" style="border:1px solid #ccc;border-radius:6px;width:100%;height:80px;touch-action:none;"></canvas></div>
            <div style="margin-bottom:16px;"><div style="font-size:12px;color:#666;margin-bottom:4px;">Potpis kooperanta:</div><canvas id="sigKooperant" width="720" height="200" style="border:1px solid #ccc;border-radius:6px;width:100%;height:80px;touch-action:none;"></canvas></div>
        </div>
        <div style="text-align:center;margin-top:16px;display:flex;gap:8px;">
            <button onclick="clearSignature('sigOtkupac');clearSignature('sigKooperant')" style="flex:1;padding:12px;font-size:14px;background:#f5f5f0;color:#666;border:1px solid #ccc;border-radius:8px;">Obriši</button>
            <button onclick="saveOtkupniListWithSignatures('${record.clientRecordID}')" style="flex:1;padding:12px;font-size:14px;background:var(--primary);color:white;border:none;border-radius:8px;">Potvrdi</button>
            <button onclick="window.print()" style="flex:1;padding:12px;font-size:14px;background:var(--accent);color:white;border:none;border-radius:8px;">Štampaj</button>
        </div>
        <button onclick="savePdfToDrive('${record.clientRecordID}')" style="width:100%;padding:12px;margin-top:8px;font-size:14px;background:#2196F3;color:white;border:none;border-radius:8px;">📄 Sačuvaj PDF na Drive</button>
        <button onclick="document.getElementById('otkupniListModal').style.display='none'" style="width:100%;padding:10px;margin-top:8px;font-size:14px;background:none;color:#666;border:1px solid #ccc;border-radius:8px;">Zatvori</button>
    </div>`;
    modal.style.display = 'block';
    setTimeout(() => { initSignaturePad('sigOtkupac'); initSignaturePad('sigKooperant'); }, 100);
}

async function saveOtkupniListWithSignatures(clientRecordID) {
    const sigO = getSignatureData('sigOtkupac');
    const sigK = getSignatureData('sigKooperant');

    if (!sigO) {
        showToast('Potpišite se kao otkupljivač!', 'error');
        return;
    }

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

        r.sigOtkupac = sigO;
        r.sigKooperant = sigK;
        r.signedAt = new Date().toISOString();

        // ne diraj syncStatus ako su potpisi samo lokalni artefakt
        // ako želiš da i potpisi idu na server, ovde bi išlo:
        // r.updatedAtClient = new Date().toISOString();
        // r.syncStatus = 'pending';
        // r.lastSyncError = '';
        // r.lastServerStatus = '';

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

    let sigOtkupac = getSignatureData('sigOtkupac') || (record.sigOtkupac || '');
    let sigKooperant = getSignatureData('sigKooperant') || (record.sigKooperant || '');
    console.log('SIG OTK length:', sigOtkupac.length);
    console.log('SIG KOOP length:', sigKooperant.length);

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

// ============================================================
// OTPREMA (dispatch)
// ============================================================
let otpremaVozacID = '';
let otpremaUnassigned = [];

function startOtpremaVozacScan() {
    const readerDiv = document.getElementById('qr-reader-otprema');
    readerDiv.style.display = 'block';
    const scanner = new Html5Qrcode('qr-reader-otprema');
    scanner.start(
        { facingMode: 'environment' },
        { fps: 10, qrbox: { width: 250, height: 250 } },
        (decodedText) => {
            scanner.stop().then(() => { readerDiv.style.display = 'none'; });
            onOtpremaVozacScanned(decodedText);
        },
        () => {}
    ).catch(err => { showToast('Kamera nije dostupna', 'error'); readerDiv.style.display = 'none'; });
}

function onOtpremaVozacScanned(text) {
    let vozID = '', vozName = '';
    try {
        const data = JSON.parse(text);
        if (data.type === 'VOZ' && data.id) { vozID = data.id; vozName = data.name || data.id; }
    } catch (e) {}
    if (!vozID && text.startsWith('VOZ-')) { vozID = text; vozName = text; }
    if (!vozID) { showToast('Nije QR vozača', 'error'); return; }

    otpremaVozacID = vozID;
    document.getElementById('otpremaVozacName').textContent = vozName;
    document.getElementById('otpremaVozacId').textContent = vozID;
    showOtpremaAssignView();
}

async function showOtpremaAssignView() {
    const mainView = document.getElementById('otpremaMainView');
    const assignView = document.getElementById('otpremaAssignView');

    if (mainView) mainView.style.display = 'none';
    if (assignView) assignView.style.display = 'block';

    let local = [];
    let server = [];

    try {
        local = await dbGetAll(db, CONFIG.STORE_NAME);
    } catch (err) {
        console.error('showOtpremaAssignView local failed:', err);
    }

    if (navigator.onLine) {
        try {
            const json = await apiFetch('action=getOtkupi&otkupacID=' + encodeURIComponent(CONFIG.OTKUPAC_ID));
            if (json && json.success && Array.isArray(json.records)) {
                server = json.records.map(mapServerOtpremaRecord);
            }
        } catch (e) {
            console.error('showOtpremaAssignView server failed:', e);
        }
    }

    const all = mergeOtpremaRecords(local, server);

    otpremaUnassigned = all.filter(r => !r.vozacID);
    otpremaUnassigned.sort((a, b) => {
        const byDate = (b.datum || '').localeCompare(a.datum || '');
        if (byDate !== 0) return byDate;

        const byTime = String(b.updatedAtClient || b.createdAtClient || '').localeCompare(
            String(a.updatedAtClient || a.createdAtClient || '')
        );
        if (byTime !== 0) return byTime;

        return String(b.clientRecordID || '').localeCompare(String(a.clientRecordID || ''));
    });

    renderOtpremaCheckboxes();
}

function renderOtpremaCheckboxes() {
    const list = document.getElementById('otpremaOtkupList');
    if (!list) return;

    if (otpremaUnassigned.length === 0) {
        list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema neraspoređenih otkupa za danas</p>';
        return;
    }

    list.innerHTML = otpremaUnassigned.map((r, i) => {
        const vr = ((r.kolicina || 0) * (r.cena || 0)).toLocaleString('sr');
        return `<div class="queue-item" style="cursor:pointer;" onclick="toggleOtpremaItem(${i})">
            <div style="display:flex;align-items:center;gap:10px;">
                <input type="checkbox" id="otpChk${i}" style="width:20px;height:20px;flex-shrink:0;" onclick="event.stopPropagation();updateOtpremaSummary();">
                <div style="flex:1;">
                    <div class="qi-header"><span class="qi-koop">${escapeHtml(r.kooperantName)}</span><span class="qi-time">${escapeHtml(r.datum)}</span></div>
                    <div class="qi-detail" style="font-size:11px;color:var(--text-muted);">${escapeHtml(r.klasa || 'I')}</div>
                    <div class="qi-detail">${escapeHtml(r.vrstaVoca)} ${escapeHtml(r.sortaVoca || '')} | ${r.kolicina} kg × ${r.cena} = ${vr} RSD</div>
                </div>
            </div>
        </div>`;
    }).join('');

    updateOtpremaSummary();
}
function toggleOtpremaItem(index) {
    const chk = document.getElementById('otpChk' + index);
    chk.checked = !chk.checked;
    updateOtpremaSummary();
}

function toggleSelectAll() {
    const checkboxes = document.querySelectorAll('[id^="otpChk"]');
    const allChecked = Array.from(checkboxes).every(c => c.checked);
    checkboxes.forEach(c => c.checked = !allChecked);
    updateOtpremaSummary();
}

function updateOtpremaSummary() {
    const div = document.getElementById('otpremaSummary');
    const text = document.getElementById('otpremaSummaryText');
    if (!div || !text) return;

    let kg = 0;
    let count = 0;

    otpremaUnassigned.forEach((r, i) => {
        const chk = document.getElementById('otpChk' + i);
        if (chk && chk.checked) {
            kg += r.kolicina || 0;
            count++;
        }
    });

    if (count > 0) {
        div.style.display = 'block';
        text.textContent = 'Izabrano: ' + count + ' otkupa | ' + kg.toLocaleString('sr') + ' kg';
    } else {
        div.style.display = 'none';
    }
}

async function confirmOtprema() {
    if (!otpremaVozacID) {
        showToast('Nema vozača', 'error');
        return;
    }

    const selected = [];

    for (let i = 0; i < otpremaUnassigned.length; i++) {
        const chk = document.getElementById('otpChk' + i);
        if (chk && chk.checked) {
            selected.push(otpremaUnassigned[i]);
        }
    }

    if (selected.length === 0) {
        showToast('Izaberite bar jedan otkup', 'error');
        return;
    }

    try {
        const nowIso = new Date().toISOString();

        for (const item of selected) {
            item.vozacID = otpremaVozacID;
            item.updatedAtClient = nowIso;
            item.syncStatus = 'pending';
            item.lastSyncError = '';
            item.lastServerStatus = '';
            item.syncAttemptAt = '';
            await dbPut(db, CONFIG.STORE_NAME, item);
        }

        showToast(selected.length + ' otkupa dodeljeno vozaču', 'success');

        cancelOtprema();
        await loadOtpremaOverview();

        if (navigator.onLine) {
            if (typeof syncQueueSafe === 'function') {
                await syncQueueSafe();
            } else if (typeof syncQueue === 'function') {
                await syncQueue();
            }
        }
    } catch (err) {
        console.error('confirmOtprema failed:', err);
        showToast('Greška pri dodeli otkupa vozaču', 'error');
    }
}

async function cancelOtprema() {
    otpremaVozacID = '';
    otpremaUnassigned = [];

    const assignView = document.getElementById('otpremaAssignView');
    const mainView = document.getElementById('otpremaMainView');

    if (assignView) assignView.style.display = 'none';
    if (mainView) mainView.style.display = 'block';

    await loadOtpremaOverview();
}

async function loadOtpremaOverview() {
    let local = [];
    let server = [];

    try {
        local = await dbGetAll(db, CONFIG.STORE_NAME);
    } catch (err) {
        console.error('loadOtpremaOverview local failed:', err);
    }

    if (navigator.onLine) {
        const json = await safeAsync(async () => {
            return await apiFetch('action=getOtkupi&otkupacID=' + encodeURIComponent(CONFIG.OTKUPAC_ID));
        }, 'Greška pri učitavanju otpreme');

        if (json && json.success && Array.isArray(json.records)) {
            server = json.records.map(mapServerOtpremaRecord);
        }
    }

    const all = mergeOtpremaRecords(local, server);

    const unassigned = all.filter(r => !r.vozacID);
    unassigned.sort((a, b) => (b.datum || '').localeCompare(a.datum || ''));

    const assigned = all.filter(r => r.vozacID);

    const uList = document.getElementById('otpremaUnassignedList');
    if (uList) {
        if (unassigned.length === 0) {
            uList.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:12px;font-size:13px;">Svi otkupi su raspoređeni</p>';
        } else {
            uList.innerHTML = unassigned.map(r => {
                const vr = ((r.kolicina || 0) * (r.cena || 0)).toLocaleString('sr-RS');
                const statusText =
                    r.syncStatus === 'syncing' ? ' | sync...' :
                    r.syncStatus === 'pending' ? ' | pending' : '';

                return `<div class="queue-item" style="border-left-color:var(--warning);">
                    <div class="qi-header"><span class="qi-koop">${escapeHtml(r.kooperantName)}</span><span class="qi-time">${escapeHtml(r.datum)}</span></div>
                    <div class="qi-detail">${escapeHtml(r.vrstaVoca)} ${escapeHtml(r.sortaVoca || '')} ${escapeHtml(r.klasa || 'I')} | ${r.kolicina} kg | ${vr} RSD${statusText}</div>
                </div>`;
            }).join('');
        }
    }

    const aList = document.getElementById('otpremaAssignedList');
    if (aList) {
        if (assigned.length === 0) {
            aList.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:12px;font-size:13px;">Nema otprema za danas</p>';
        } else {
            const grouped = {};

            assigned.forEach(r => {
                const v = r.vozacID;
                if (!grouped[v]) grouped[v] = { items: [], kg: 0 };
                grouped[v].items.push(r);
                grouped[v].kg += r.kolicina || 0;
            });

            aList.innerHTML = Object.entries(grouped).map(([vozID, g]) =>
                `<div style="background:white;border-radius:var(--radius);padding:14px;margin-bottom:12px;box-shadow:0 1px 3px rgba(0,0,0,0.08);border-left:4px solid var(--success);">
                    <div style="display:flex;justify-content:space-between;margin-bottom:6px;">
                        <strong style="color:var(--primary);">🚛 ${escapeHtml(vozID)}</strong>
                        <span style="font-weight:600;">${g.kg.toLocaleString('sr-RS')} kg | ${g.items.length} otk.</span>
                    </div>
                    ${g.items.map(r => {
                        const st =
                            r.syncStatus === 'syncing' ? ' (sync...)' :
                            r.syncStatus === 'pending' ? ' (pending)' : '';
                        return `<div style="padding:3px 0;font-size:12px;border-top:1px solid #eee;">${escapeHtml(r.kooperantName)} | ${escapeHtml(r.vrstaVoca)} ${escapeHtml(r.klasa || '')} | ${r.kolicina} kg${escapeHtml(st)}</div>`;
                    }).join('')}
                </div>`
            ).join('');
        }
    }
}


//HELPERS
function mapServerOtpremaRecord(r) {
    return {
        clientRecordID: r.ClientRecordID || '',
        serverRecordID: r.ServerRecordID || '',
        createdAtClient: normalizeDateTime(r.CreatedAtClient),
        updatedAtClient: normalizeDateTime(r.UpdatedAtClient || r.CreatedAtClient),
        updatedAtServer: normalizeDateTime(r.UpdatedAtServer || r.ReceivedAt),
        syncedAt: normalizeDateTime(r.UpdatedAtServer || r.ReceivedAt),

        datum: fmtDate(r.Datum),
        kooperantID: r.KooperantID || '',
        kooperantName: r.KooperantName || r.KooperantID || '',
        vrstaVoca: r.VrstaVoca || '',
        sortaVoca: r.SortaVoca || '',
        klasa: r.Klasa || 'I',
        kolicina: parseFloat(r.Kolicina) || 0,
        cena: parseFloat(r.Cena) || 0,
        kolAmbalaze: parseInt(r.KolAmbalaze, 10) || 0,
        parcelaID: r.ParcelaID || '',
        vozacID: r.VozacID || r.VozaciID || '',
        napomena: r.Napomena || '',

        syncStatus: 'synced',
        syncAttempts: 0,
        lastSyncError: '',
        lastServerStatus: 'server'
    };
}

function normalizeLocalOtpremaRecord(r) {
    return {
        clientRecordID: r.clientRecordID || '',
        serverRecordID: r.serverRecordID || '',
        createdAtClient: normalizeDateTime(r.createdAtClient),
        updatedAtClient: normalizeDateTime(r.updatedAtClient || r.createdAtClient),
        updatedAtServer: normalizeDateTime(r.updatedAtServer),
        syncedAt: normalizeDateTime(r.syncedAt),

        datum: r.datum || '',
        kooperantID: r.kooperantID || '',
        kooperantName: r.kooperantName || r.kooperantID || '',
        vrstaVoca: r.vrstaVoca || '',
        sortaVoca: r.sortaVoca || '',
        klasa: r.klasa || 'I',
        kolicina: parseFloat(r.kolicina) || 0,
        cena: parseFloat(r.cena) || 0,
        kolAmbalaze: parseInt(r.kolAmbalaze, 10) || 0,
        parcelaID: r.parcelaID || '',
        vozacID: r.vozacID || '',
        napomena: r.napomena || '',

        syncStatus: r.syncStatus || 'pending',
        syncAttempts: parseInt(r.syncAttempts, 10) || 0,
        lastSyncError: r.lastSyncError || '',
        lastServerStatus: r.lastServerStatus || ''
    };
}

function mergeOtpremaRecords(local, server) {
    const merged = new Map();

    (server || []).forEach(r => {
        if (r && r.clientRecordID) merged.set(r.clientRecordID, r);
    });

    (local || []).forEach(r => {
        if (!r || !r.clientRecordID) return;

        const localNorm = normalizeLocalOtpremaRecord(r);
        const existing = merged.get(localNorm.clientRecordID);

        if (!existing) {
            merged.set(localNorm.clientRecordID, localNorm);
            return;
        }

        if (localNorm.syncStatus === 'pending' || localNorm.syncStatus === 'syncing') {
            merged.set(localNorm.clientRecordID, localNorm);
            return;
        }

        const localUpdated = localNorm.updatedAtClient || localNorm.createdAtClient || '';
        const serverUpdated = existing.updatedAtServer || existing.updatedAtClient || existing.createdAtClient || '';

        if (localUpdated && serverUpdated && localUpdated > serverUpdated) {
            merged.set(localNorm.clientRecordID, localNorm);
        }
    });

    return Array.from(merged.values());
}

function normalizeDateTime(value) {
    if (!value) return '';
    try {
        const d = new Date(value);
        if (isNaN(d.getTime())) return String(value);
        return d.toISOString();
    } catch (_) {
        return String(value);
    }
}
