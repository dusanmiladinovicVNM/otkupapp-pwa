// ============================================================
// STATE
// ============================================================
let db = null;
let qrScanner = null;
let stammdaten = { kooperanti: [], kulture: [], config: [], parcele: [], stanice: [], kupci: [], vozaci: [] };
let selectedMera = '';
let mgmtData = null;
let parcelExpertOpen = {};

// ============================================================
// MANAGEMENT SUB-TAB DEFINITIONS
// ============================================================
const MGMT_SUBS = {
    kooperanti: [
        { id: 'koop-kartica', label: 'Kartica', load: null },
        { id: 'koop-saldo', label: 'Saldo', load: loadMgmtKoopSaldo },
        { id: 'koop-pregled', label: 'Pregled', load: loadMgmtKoopPregled }
    ],
    stanice: [
        { id: 'sta-otkupi', label: 'Otkupi', load: null },
        { id: 'sta-saldo', label: 'Saldo OM', load: loadMgmtSaldoOM },
        { id: 'sta-roba', label: 'Roba', load: loadMgmtOtkupPoOM }
    ],
    kupci: [
        { id: 'kup-fakture', label: 'Fakture', load: null },
        { id: 'kup-saldo', label: 'Saldo', load: loadMgmtKupci },
        { id: 'kup-roba', label: 'Roba', load: loadMgmtPredato }
    ],
    agrohemija: [
        { id: 'agro-izdavanje', label: 'Izdavanje', load: function() { populateIzdDropdowns(); } },
        { id: 'agro-stanje', label: 'Stanje', load: loadMgmtAgroStanje }
    ]
};

// ============================================================
// INIT
// ============================================================
document.addEventListener('DOMContentLoaded', async () => {
    if (!getLs('authToken', '') || !getLs('otkupacID', '')) {
        showLoginScreen();
        return;
    }
    db = await openDB();
    await loadStammdaten();
    applyRoleVisibility();
    document.getElementById('headerInfo').textContent = CONFIG.USER_ROLE + ': ' + CONFIG.ENTITY_NAME;

    if (CONFIG.USER_ROLE === 'Otkupac') {
        populateVrstaDropdown();
        applyDefaults();
        const today = new Date().toISOString().split('T')[0];
        document.getElementById('fldPregledOd').value = today;
        document.getElementById('fldPregledDo').value = today;
        const otpDatumEl = document.getElementById('fldOtpremniceDatum');
        if (otpDatumEl) otpDatumEl.value = today;
    }
    if (CONFIG.USER_ROLE === 'Kooperant') {
        populateAgroParcele();
    }
    if (CONFIG.USER_ROLE === 'Management') {
        populateMgmtStanice();
        const today = new Date().toISOString().split('T')[0];
        document.getElementById('mgmtOtkupiOd').value = today;
        document.getElementById('mgmtOtkupiDo').value = today;
        prefetchMgmtData().then(() => { 
            populateMgmtKupciDropdown();
            showTab('dispecer');
        });
    }
    updateSyncBadge();
    window.addEventListener('online', () => { updateSyncBadge(); if (CONFIG.USER_ROLE === 'Otkupac') syncQueue(); });
    window.addEventListener('offline', () => updateSyncBadge());
    setInterval(() => { if (navigator.onLine && CONFIG.USER_ROLE === 'Otkupac') syncQueue(); }, 60000);
});

// ============================================================
// PREFETCH
// ============================================================
async function prefetchMgmtData() {
    try {
        const json = await apiFetch('action=getMgmtAll');
        if (json && json.success) { mgmtData = json; }
    } catch (e) {}
}

function populateMgmtKupciDropdown() {
    const sel = document.getElementById('mgmtFaktureKupac');
    if (!sel) return;
    sel.innerHTML = '<option value="">-- Izaberi --</option>';
    const kupci = mgmtData ? (mgmtData.saldoKupci || []) : [];
    kupci.forEach(k => {
        const o = document.createElement('option');
        o.value = k.KupacID || k.Kupac;
        o.textContent = k.Kupac || k.KupacID;
        sel.appendChild(o);
    });
}

// ============================================================
// QR SCANNER
// ============================================================
function startQRScan() {
    const readerDiv = document.getElementById('qr-reader');
    readerDiv.style.display = 'block';
    if (qrScanner) qrScanner.clear();
    qrScanner = new Html5Qrcode('qr-reader');
    qrScanner.start(
        { facingMode: 'environment' },
        { fps: 10, qrbox: { width: 250, height: 250 } },
        (decodedText) => { onQRScanned(decodedText); qrScanner.stop().then(() => { readerDiv.style.display = 'none'; }); },
        () => {}
    ).catch(err => { showToast('Kamera nije dostupna: ' + err, 'error'); readerDiv.style.display = 'none'; });
}

function onQRScanned(text) {
    try { const data = JSON.parse(text); if (data.id) { setKooperant(data.id, data.name || data.id); return; } } catch (e) {}
    if (text.startsWith('KOOP-')) {
        const koop = stammdaten.kooperanti.find(k => k.KooperantID === text);
        setKooperant(text, koop ? (koop.Ime + ' ' + koop.Prezime) : text);
        return;
    }
    showToast('Nepoznat QR kod', 'error');
}

function setKooperant(id, name) {
    document.getElementById('fldKooperantID').value = id;
    document.getElementById('koopName').textContent = name;
    document.getElementById('koopId').textContent = id;
    document.getElementById('koopDisplay').classList.add('visible');
    showToast('Kooperant: ' + name, 'success');
    populateParcelaDropdown(id);
}

function startVozacQRScan() {
    const readerDiv = document.getElementById('qr-reader-vozac');
    readerDiv.style.display = 'block';
    const scanner = new Html5Qrcode('qr-reader-vozac');
    scanner.start(
        { facingMode: 'environment' },
        { fps: 10, qrbox: { width: 250, height: 250 } },
        (decodedText) => {
            onVozacQRScanned(decodedText);
            scanner.stop().then(() => { readerDiv.style.display = 'none'; });
        },
        () => {}
    ).catch(err => { showToast('Kamera nije dostupna: ' + err, 'error'); readerDiv.style.display = 'none'; });
}

function onVozacQRScanned(text) {
    try {
        const data = JSON.parse(text);
        if (data.type === 'VOZ' && data.id) { setVozac(data.id, data.name || data.id); return; }
    } catch (e) {}
    if (text.startsWith('VOZ-')) {
        setVozac(text, text);
        return;
    }
    showToast('Nije QR vozača', 'error');
}

function setVozac(id, name) {
    document.getElementById('fldVozacID').value = id;
    document.getElementById('vozacName').textContent = name;
    document.getElementById('vozacId').textContent = id;
    document.getElementById('vozacDisplay').classList.add('visible');
    showToast('Vozač: ' + name, 'success');
}

function clearVozac() {
    document.getElementById('fldVozacID').value = '';
    document.getElementById('vozacDisplay').classList.remove('visible');
}
    
// ============================================================
// QR PROFILE
// ============================================================
function showQRProfile() {
    const modal = document.getElementById('qrProfileModal');
    document.getElementById('qrProfileName').textContent = CONFIG.ENTITY_NAME;
    document.getElementById('qrProfileRole').textContent = CONFIG.USER_ROLE;
    document.getElementById('qrProfileID').textContent = CONFIG.ENTITY_ID;
    modal.style.display = 'flex';
    
    generateQRCode('qrProfileCanvas', JSON.stringify({
        type: CONFIG.USER_ROLE === 'Kooperant' ? 'KOOP' : CONFIG.USER_ROLE === 'Otkupac' ? 'OTK' : CONFIG.USER_ROLE === 'Vozac' ? 'VOZ' : 'MGMT',
        id: CONFIG.ENTITY_ID,
        name: CONFIG.ENTITY_NAME
    }));
}

function generateQRCode(canvasId, text) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    
    // Koristi QR generisanje iz eksternog CDN-a
    const img = new Image();
    img.onload = function() {
        canvas.width = 250;
        canvas.height = 250;
        const ctx = canvas.getContext('2d');
        ctx.fillStyle = 'white';
        ctx.fillRect(0, 0, 250, 250);
        ctx.drawImage(img, 0, 0, 250, 250);
    };
    img.onerror = function() {
        // Fallback: prikaži tekst ako API ne radi
        canvas.width = 250;
        canvas.height = 250;
        const ctx = canvas.getContext('2d');
        ctx.fillStyle = 'white';
        ctx.fillRect(0, 0, 250, 250);
        ctx.fillStyle = '#1a5e2a';
        ctx.font = 'bold 16px sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText(CONFIG.ENTITY_ID, 125, 125);
    };
    img.src = 'https://api.qrserver.com/v1/create-qr-code/?size=250x250&data=' + encodeURIComponent(text);
}
    
// ============================================================
// STAMMDATEN
// ============================================================
async function loadStammdaten() {
    try {
        const cached = await dbGetAll(db, CONFIG.STAMM_STORE);
        const obj = cached.find(c => c.key === 'all');
        if (obj) stammdaten = obj.data;
    } catch (e) {}

    if (navigator.onLine) {
        try {
            const json = await apiFetch('action=getStammdaten');
            if (json && json.success && json.data) {
                stammdaten = json.data;
                await dbPut(db, CONFIG.STAMM_STORE, {
                    key: 'all',
                    data: stammdaten,
                    updatedAt: new Date().toISOString()
                });
            }
        } catch (e) {}
    }
}
    
function fmtStanica(stanicaID) {
    if (!stanicaID) return '';
    const s = (stammdaten.stanice || []).find(s => s.StanicaID === stanicaID);
    const name = s ? (s.Naziv || s.Mesto || stanicaID) : stanicaID;
    if (name === stanicaID) return stanicaID;
    return name + ' <span style="font-size:11px;color:var(--text-muted);">' + stanicaID + '</span>';
}

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
            <div style="font-size:12px;color:#666;">${gv('SELLER_STREET')}, ${gv('SELLER_CITY')} ${gv('SELLER_POSTAL_CODE')}</div>
            <div style="font-size:12px;color:#666;">PIB: ${gv('SELLER_PIB')} | MB: ${gv('SELLER_MATICNI_BROJ')}</div>
            <div style="font-size:12px;color:#666;">TR: ${gv('SELLER_ACCOUNT')}</div>
        </div>
        <h2 style="text-align:center;margin-bottom:14px;font-size:20px;">OTKUPNI LIST</h2>
        <div style="background:#f5f5f0;padding:10px;border-radius:8px;margin-bottom:12px;font-size:13px;">
            <div><strong>${koop.Ime || ''} ${koop.Prezime || ''}</strong></div>
            <div>${koop.Adresa || ''}, ${koop.Mesto || ''}</div>
            <div>JMBG: ${koop.JMBG || '________'} | BPG: ${koop.BPGBroj || '________'}</div>
        </div>
        <table style="width:100%;border-collapse:collapse;font-size:14px;">
            <tr><td style="padding:6px;color:#666;width:40%;">Datum:</td><td style="padding:6px;font-weight:600;">${record.datum}</td></tr>
            <tr><td style="padding:6px;color:#666;">Proizvod:</td><td style="padding:6px;">${record.vrstaVoca} ${record.sortaVoca || ''}</td></tr>
            <tr><td style="padding:6px;color:#666;">Klasa:</td><td style="padding:6px;">${record.klasa}</td></tr>
            <tr><td style="padding:6px;color:#666;">Količina:</td><td style="padding:6px;font-weight:600;">${record.kolicina} kg</td></tr>
            <tr><td style="padding:6px;color:#666;">Cena:</td><td style="padding:6px;">${record.cena} RSD/kg</td></tr>
            <tr style="border-top:1px solid #ccc;"><td style="padding:6px;color:#666;">Vrednost:</td><td style="padding:6px;font-weight:600;">${vrednostNum.toLocaleString('sr')} RSD</td></tr>
            ${pdvStopa > 0 ? '<tr><td style="padding:6px;color:#666;">PDV naknada (' + pdvStopa + '%):</td><td style="padding:6px;">' + pdvIznos.toLocaleString('sr') + ' RSD</td></tr>' : ''}
            <tr style="border-top:2px solid #333;"><td style="padding:8px;color:#666;">ZA ISPLATU:</td><td style="padding:8px;font-weight:700;font-size:18px;">${ukupno.toLocaleString('sr')} RSD</td></tr>
            <tr><td style="padding:6px;color:#666;">Ambalaža:</td><td style="padding:6px;">${record.kolAmbalaze} kom</td></tr>
            ${record.parcelaID ? '<tr><td style="padding:6px;color:#666;">Parcela:</td><td style="padding:6px;">' + record.parcelaID + '</td></tr>' : ''}
            <tr><td style="padding:6px;color:#666;">Rok isplate:</td><td style="padding:6px;">${gv('OtkupRokIsplate') || 'Po dogovoru'}</td></tr>
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

function initSignaturePad(canvasId) {
    const canvas = document.getElementById(canvasId); if (!canvas) return;
    const ctx = canvas.getContext('2d');
    const rect = canvas.getBoundingClientRect();
    const scaleX = canvas.width / rect.width;
    const scaleY = canvas.height / rect.height;
    ctx.scale(scaleX, scaleY);
    ctx.strokeStyle = '#1a1a1a';
    ctx.lineWidth = 2;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    let drawing = false, lastX = 0, lastY = 0;
    function getPos(e) { const r = canvas.getBoundingClientRect(); const touch = e.touches ? e.touches[0] : e; return { x: touch.clientX - r.left, y: touch.clientY - r.top }; }
    function startDraw(e) { e.preventDefault(); drawing = true; const p = getPos(e); lastX = p.x; lastY = p.y; }
    function draw(e) { if (!drawing) return; e.preventDefault(); const p = getPos(e); ctx.beginPath(); ctx.moveTo(lastX, lastY); ctx.lineTo(p.x, p.y); ctx.stroke(); lastX = p.x; lastY = p.y; }
    function stopDraw(e) { if (e) e.preventDefault(); drawing = false; }
    canvas.addEventListener('mousedown', startDraw); canvas.addEventListener('mousemove', draw);
    canvas.addEventListener('mouseup', stopDraw); canvas.addEventListener('mouseleave', stopDraw);
    canvas.addEventListener('touchstart', startDraw, { passive: false });
    canvas.addEventListener('touchmove', draw, { passive: false });
    canvas.addEventListener('touchend', stopDraw, { passive: false });
}

function clearSignature(canvasId) {
    const canvas = document.getElementById(canvasId); if (!canvas) return;
    canvas.getContext('2d').clearRect(0, 0, canvas.width, canvas.height);
}

function getSignatureData(canvasId) {
    const canvas = document.getElementById(canvasId); if (!canvas) return '';
    const data = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data;
    if (!data.some((val, i) => i % 4 === 3 && val > 0)) return '';
    return canvas.toDataURL('image/png');
}

async function saveOtkupniListWithSignatures(clientRecordID) {
    const sigO = getSignatureData('sigOtkupac'), sigK = getSignatureData('sigKooperant');
    if (!sigO) { showToast('Potpišite se kao otkupljivač!', 'error'); return; }
    if (!sigK) { showToast('Kooperant mora da se potpiše!', 'error'); return; }
    try {
        const r = await dbGet(db, CONFIG.STORE_NAME, clientRecordID);
        if (r) {
            r.sigOtkupac = sigO;
            r.sigKooperant = sigK;
            r.signedAt = new Date().toISOString();
            await dbPut(db, CONFIG.STORE_NAME, r);
        }
        req.onsuccess = () => { const r = req.result; if (r) { r.sigOtkupac = sigO; r.sigKooperant = sigK; r.signedAt = new Date().toISOString(); store.put(r); } };
    } catch (e) {}
    showToast('Otkupni list potpisan!', 'success');
    document.getElementById('otkupniListModal').style.display = 'none';
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
// SYNC
// ============================================================
async function syncQueue() {
    const pending = await dbGetByIndex(db, CONFIG.STORE_NAME, 'syncStatus', 'pending');
    if (pending.length === 0) return;

    updateSyncBadge('syncing');

    const json = await apiPost('sync', {
        otkupacID: CONFIG.OTKUPAC_ID,
        records: pending
    });

    if (json && json.success) {
        for (const r of pending) {
            r.syncStatus = 'synced';
            await dbPut(db, CONFIG.STORE_NAME, r);
        }
        showToast('Sinhr: ' + pending.length, 'success');
    } else if (json) {
        showToast('Greška: ' + (json.error || ''), 'error');
    }

    updateSyncBadge();
    updateStats();
}

async function syncAgromere() {
    const pending = await dbGetByIndex(db, CONFIG.AGRO_STORE, 'syncStatus', 'pending');
    if (pending.length === 0) return;

    const json = await apiPost('syncAgromere', {
        kooperantID: CONFIG.ENTITY_ID,
        records: pending
    });

    if (json && json.success) {
        for (const r of pending) {
            r.syncStatus = 'synced';
            await dbPut(db, CONFIG.AGRO_STORE, r);
        }
        showToast('Agromere sinhr: ' + pending.length, 'success');
    }
}

async function syncNow() {
    if (!navigator.onLine) { showToast('Nema konekcije', 'error'); return; }
    await syncQueue(); renderQueueList();
}

// ============================================================
// OTKUP PREGLED
// ============================================================
async function loadOtkupPregled() {
    document.getElementById('pregledList').innerHTML = '<p style="text-align:center;padding:20px;color:var(--text-muted);">Učitavanje...</p>';
    const od = document.getElementById('fldPregledOd').value, doo = document.getElementById('fldPregledDo').value;
    const local = await dbGetAll(db, CONFIG.STORE_NAME);
    let server = [];
    try {
        const json = await apiFetch('action=getOtkupi&otkupacID=' + encodeURIComponent(CONFIG.OTKUPAC_ID));
        if (json && json.success && json.records) server = json.records.map(r => ({
            clientRecordID: r.ClientRecordID||'', datum: fmtDate(r.Datum), kooperantID: r.KooperantID||'',
            kooperantName: r.KooperantName||r.KooperantID||'', vrstaVoca: r.VrstaVoca||'', sortaVoca: r.SortaVoca||'',
            klasa: r.Klasa||'I', kolicina: parseFloat(r.Kolicina)||0, cena: parseFloat(r.Cena)||0,
            kolAmbalaze: parseInt(r.KolAmbalaze)||0, parcelaID: r.ParcelaID||'', syncStatus: r.SyncStatus||'synced'
        }));
    } catch (e) {}
    const serverIDs = new Set(server.map(r => r.clientRecordID));
    let all = [...server, ...local.filter(r => r.syncStatus === 'pending' && !serverIDs.has(r.clientRecordID))];
    if (od) all = all.filter(r => r.datum >= od);
    if (doo) all = all.filter(r => r.datum <= doo);
    all.sort((a, b) => b.datum.localeCompare(a.datum));
    const kg = all.reduce((s, r) => s + (r.kolicina||0), 0);
    const vr = all.reduce((s, r) => s + (r.kolicina||0)*(r.cena||0), 0);
    document.getElementById('statPregledCount').textContent = all.length;
    document.getElementById('statPregledKg').textContent = kg.toLocaleString('sr');
    document.getElementById('statPregledVrednost').textContent = vr.toLocaleString('sr');
    document.getElementById('statPregledKoop').textContent = new Set(all.map(r => r.kooperantID)).size;
    const list = document.getElementById('pregledList');
    if (all.length === 0) { list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:40px;">Nema otkupa</p>'; return; }
    list.innerHTML = all.map(r => {
        const v = ((r.kolicina||0)*(r.cena||0)).toLocaleString('sr');
        const bc = r.syncStatus==='pending' ? 'var(--warning)' : 'var(--success)';
        return `<div class="queue-item" style="border-left-color:${bc};"><div class="qi-header"><span class="qi-koop">${r.kooperantName}</span><span class="qi-time">${r.datum}</span></div><div class="qi-detail">${r.vrstaVoca} ${r.sortaVoca||''} ${r.klasa} | ${r.kolicina} kg × ${r.cena} = <strong>${v} RSD</strong>${r.parcelaID?' | '+r.parcelaID:''}</div></div>`;
    }).join('');
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
    document.getElementById('otpremaMainView').style.display = 'none';
    document.getElementById('otpremaAssignView').style.display = 'block';

    const today = new Date().toISOString().split('T')[0];
    
    const local = await dbGetAll(db, CONFIG.STORE_NAME);
    
    let server = [];
    try {
        const json = await apiFetch('action=getOtkupi&otkupacID=' + encodeURIComponent(CONFIG.OTKUPAC_ID));
        if (json && json.success && json.records) server = json.records.map(r => ({
            clientRecordID: r.ClientRecordID || '', datum: fmtDate(r.Datum),
            kooperantID: r.KooperantID || '', kooperantName: r.KooperantName || r.KooperantID || '',
            vrstaVoca: r.VrstaVoca || '', sortaVoca: r.SortaVoca || '',
            klasa: r.Klasa || 'I', kolicina: parseFloat(r.Kolicina) || 0,
            cena: parseFloat(r.Cena) || 0, kolAmbalaze: parseInt(r.KolAmbalaze) || 0,
            parcelaID: r.ParcelaID || '', vozacID: r.VozacID || r.VozaciID || '',
            napomena: r.Napomena || '', syncStatus: 'synced'
        }));
    } catch (e) {}
    
    const serverIDs = new Set(server.map(r => r.clientRecordID));
    const all = [...server, ...local.filter(r => r.syncStatus === 'pending' && !serverIDs.has(r.clientRecordID))];
    
    otpremaUnassigned = all.filter(r => !r.vozacID);
    otpremaUnassigned.sort((a, b) => (b.datum || '').localeCompare(a.datum || ''));
    renderOtpremaCheckboxes();
}

function renderOtpremaCheckboxes() {
    const list = document.getElementById('otpremaOtkupList');
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
                    <div class="qi-header"><span class="qi-koop">${r.kooperantName}</span><span class="qi-time">${r.datum}</span></div>
                    <div class="qi-detail" style="font-size:11px;color:var(--text-muted);">${r.klasa || 'I'}</div>
                    <div class="qi-detail">${r.vrstaVoca} ${r.sortaVoca || ''} | ${r.kolicina} kg × ${r.cena} = ${vr} RSD</div>
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
    let kg = 0, count = 0;
    otpremaUnassigned.forEach((r, i) => {
        const chk = document.getElementById('otpChk' + i);
        if (chk && chk.checked) { kg += r.kolicina || 0; count++; }
    });
    const div = document.getElementById('otpremaSummary');
    if (count > 0) {
        div.style.display = 'block';
        document.getElementById('otpremaSummaryText').textContent = 
            'Izabrano: ' + count + ' otkupa | ' + kg.toLocaleString('sr') + ' kg';
    } else {
        div.style.display = 'none';
    }
}

async function confirmOtprema() {
    if (!otpremaVozacID) { showToast('Nema vozača', 'error'); return; }
    
    let count = 0;
    for (let i = 0; i < otpremaUnassigned.length; i++) {
        const chk = document.getElementById('otpChk' + i);
        if (chk && chk.checked) {
            otpremaUnassigned[i].vozacID = otpremaVozacID;
            otpremaUnassigned[i].syncStatus = 'pending';
            await dbPut(db, CONFIG.STORE_NAME, otpremaUnassigned[i]);
            count++;
        }
    }
    
    if (count === 0) { showToast('Izaberite bar jedan otkup', 'error'); return; }
    
    showToast(count + ' otkupa dodeljeno vozaču', 'success');
    cancelOtprema();
    loadOtpremaOverview();
    if (navigator.onLine) syncQueue();
}

function cancelOtprema() {
    otpremaVozacID = '';
    otpremaUnassigned = [];
    document.getElementById('otpremaAssignView').style.display = 'none';
    document.getElementById('otpremaMainView').style.display = 'block';
    loadOtpremaOverview();
}

async function loadOtpremaOverview() {
    const today = new Date().toISOString().split('T')[0];
    
    const local = await dbGetAll(db, CONFIG.STORE_NAME);
    
    let server = [];
    try {
        const json = await apiFetch('action=getOtkupi&otkupacID=' + encodeURIComponent(CONFIG.OTKUPAC_ID));
        if (json && json.success && json.records) server = json.records.map(r => ({
            clientRecordID: r.ClientRecordID || '', datum: fmtDate(r.Datum),
            kooperantID: r.KooperantID || '', kooperantName: r.KooperantName || r.KooperantID || '',
            vrstaVoca: r.VrstaVoca || '', sortaVoca: r.SortaVoca || '',
            klasa: r.Klasa || 'I', kolicina: parseFloat(r.Kolicina) || 0,
            cena: parseFloat(r.Cena) || 0, kolAmbalaze: parseInt(r.KolAmbalaze) || 0,
            parcelaID: r.ParcelaID || '', vozacID: r.VozacID || r.VozaciID || '',
            napomena: r.Napomena || '', syncStatus: 'synced'
        }));
    } catch (e) {}
    
    const serverIDs = new Set(server.map(r => r.clientRecordID));
    const all = [...server, ...local.filter(r => r.syncStatus === 'pending' && !serverIDs.has(r.clientRecordID))];
    const unassigned = all.filter(r => !r.vozacID);
    unassigned.sort((a, b) => (b.datum || '').localeCompare(a.datum || ''));
    const assigned = all.filter(r => r.vozacID);
    
    const uList = document.getElementById('otpremaUnassignedList');
    if (unassigned.length === 0) {
        uList.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:12px;font-size:13px;">Svi otkupi su raspoređeni</p>';
    } else {
        uList.innerHTML = unassigned.map(r => {
            const vr = ((r.kolicina || 0) * (r.cena || 0)).toLocaleString('sr');
            return `<div class="queue-item" style="border-left-color:var(--warning);">
                <div class="qi-header"><span class="qi-koop">${r.kooperantName}</span><span class="qi-time">${r.datum}</span></div>
                <div class="qi-detail">${r.vrstaVoca} ${r.sortaVoca || ''} ${r.klasa || 'I'} | ${r.kolicina} kg | ${vr} RSD</div>
            </div>`;
        }).join('');
    }
    
    const aList = document.getElementById('otpremaAssignedList');
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
                    <strong style="color:var(--primary);">🚛 ${vozID}</strong>
                    <span style="font-weight:600;">${g.kg.toLocaleString('sr')} kg | ${g.items.length} otk.</span>
                </div>
                ${g.items.map(r => `<div style="padding:3px 0;font-size:12px;border-top:1px solid #eee;">${r.kooperantName} | ${r.vrstaVoca} ${r.klasa || ''} | ${r.kolicina} kg</div>`).join('')}
            </div>`).join('');
    }
}

// ============================================================
// VOZAC: ZBIRNA
// ============================================================
let vozacOtkupi = [];

async function loadVozacData() {
    // Load otkupi assigned to this vozac
    vozacOtkupi = [];
    try {
        const json = await apiFetch('action=getVozacOtkupi');
        if (json && json.success && json.records) {
            vozacOtkupi = json.records.map(r => ({
                clientRecordID: r.ClientRecordID || '',
                datum: fmtDate(r.Datum),
                kooperantName: r.KooperantName || r.KooperantID || '',
                kooperantID: r.KooperantID || '',
                vrstaVoca: r.VrstaVoca || '',
                sortaVoca: r.SortaVoca || '',
                klasa: r.Klasa || 'I',
                kolicina: parseFloat(r.Kolicina) || 0,
                cena: parseFloat(r.Cena) || 0,
                tipAmbalaze: r.TipAmbalaze || '',
                kolAmbalaze: parseInt(r.KolAmbalaze) || 0,
                stanicaID: r.OtkupacID || r._source || '',
                zbirnaID: r._zbirnaID || ''
            }));
        }
    } catch (e) {}
    renderVozacOtpremnice();
    loadVozacZbirne();
}

function renderVozacOtpremnice() {
    const today = new Date().toISOString().split('T')[0];
    const todayOtkupi = vozacOtkupi.filter(r => r.datum === today);
    
    const list = document.getElementById('vozacOtpremniceList');
    if (todayOtkupi.length === 0) {
        list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema otpremnica za danas</p>';
        document.getElementById('btnNovaZbirna').style.display = 'none';
        return;
    }
    document.getElementById('btnNovaZbirna').style.display = '';
    
    // Group by stanica
    const grouped = {};
    todayOtkupi.forEach(r => {
        const s = r.stanicaID || '?';
        if (!grouped[s]) grouped[s] = { items: [], kg: 0, amb: 0 };
        grouped[s].items.push(r);
        grouped[s].kg += r.kolicina || 0;
        grouped[s].amb += r.kolAmbalaze || 0;
    });
    
    list.innerHTML = Object.entries(grouped).map(([sta, g]) =>
        `<div style="background:white;border-radius:var(--radius);padding:14px;margin-bottom:12px;box-shadow:0 1px 3px rgba(0,0,0,0.08);border-left:4px solid var(--primary);">
            <div style="display:flex;justify-content:space-between;margin-bottom:6px;">
                <strong style="color:var(--primary);">${fmtStanica(sta)}</strong>
                <span style="font-weight:600;">${g.kg.toLocaleString('sr')} kg</span>
            </div>
            <div style="font-size:12px;color:var(--text-muted);margin-bottom:6px;">${g.items.length} otkupa | Amb: ${g.amb}</div>
            ${g.items.map(r => `<div style="padding:3px 0;font-size:12px;border-top:1px solid #eee;">
                ${r.kooperantName} | ${r.vrstaVoca} ${r.klasa} | ${r.kolicina} kg | ${r.kolAmbalaze} amb
            </div>`).join('')}
        </div>`).join('');
}

async function startZbirnaCreation() {
    document.getElementById('zbirnaMainView').style.display = 'none';
    document.getElementById('zbirnaCreateView').style.display = 'block';
    
    const sel = document.getElementById('fldZbirnaKupac');
    sel.innerHTML = '<option value="">-- Izaberi kupca --</option>';
    
    // Populate from stammdaten
    (stammdaten.kupci || []).forEach(k => {
        const o = document.createElement('option');
        o.value = k.KupacID;
        o.textContent = k.Naziv + ' (' + k.KupacID + ')';
        sel.appendChild(o);
    });
    
    // Optional fallback from mgmtData
    if (mgmtData && mgmtData.saldoKupci) {
        mgmtData.saldoKupci.forEach(k => {
            const value = k.KupacID || k.Kupac;
            if (!value) return;

            const exists = Array.from(sel.options).some(opt => opt.value === value);
            if (exists) return;

            const o = document.createElement('option');
            o.value = value;
            o.textContent = k.Kupac || k.KupacID;
            sel.appendChild(o);
        });
    }
    
    renderZbirnaSummary();
}

function renderZbirnaSummary() {
    const today = new Date().toISOString().split('T')[0];
    const todayOtkupi = vozacOtkupi.filter(r => r.datum === today);
    
    let totalKgI = 0, totalKgII = 0, totalAmb = 0;
    todayOtkupi.forEach(r => {
        if (r.klasa === 'II') totalKgII += r.kolicina || 0;
        else totalKgI += r.kolicina || 0;
        totalAmb += r.kolAmbalaze || 0;
    });
    
    document.getElementById('zbirnaOtkupiSummary').innerHTML = 
        `<div style="font-size:16px;font-weight:700;">Ukupno: ${(totalKgI + totalKgII).toLocaleString('sr')} kg</div>
         <div style="font-size:13px;opacity:0.9;">Kl. I: ${totalKgI.toLocaleString('sr')} kg | Kl. II: ${totalKgII.toLocaleString('sr')} kg | Amb: ${totalAmb}</div>
         <div style="font-size:12px;opacity:0.7;">${todayOtkupi.length} otkupa sa ${new Set(todayOtkupi.map(r => r.stanicaID)).size} stanica</div>`;
    
    // List individual otkupi
    document.getElementById('zbirnaOtkupiList').innerHTML = todayOtkupi.map(r => {
        const vr = ((r.kolicina || 0) * (r.cena || 0)).toLocaleString('sr');
        return `<div class="queue-item">
            <div class="qi-header"><span class="qi-koop">${r.kooperantName}</span><span class="qi-time">${fmtStanica(r.stanicaID)}</span></div>
            <div class="qi-detail">${r.vrstaVoca} ${r.klasa} | ${r.kolicina} kg × ${r.cena} = ${vr} RSD | Amb: ${r.kolAmbalaze}</div>
        </div>`;
    }).join('');
}

async function confirmZbirna() {
    const kupacID = document.getElementById('fldZbirnaKupac').value;
    if (!kupacID) { showToast('Izaberite kupca!', 'error'); return; }
    
    const today = new Date().toISOString().split('T')[0];
    const todayOtkupi = vozacOtkupi.filter(r => r.datum === today);
    
    if (todayOtkupi.length === 0) { showToast('Nema otkupa za danas', 'error'); return; }
    
    let totalKgI = 0, totalKgII = 0, totalAmb = 0;
    const vrste = new Set(), sorte = new Set();
    todayOtkupi.forEach(r => {
        if (r.klasa === 'II') totalKgII += r.kolicina || 0;
        else totalKgI += r.kolicina || 0;
        totalAmb += r.kolAmbalaze || 0;
        if (r.vrstaVoca) vrste.add(r.vrstaVoca);
        if (r.sortaVoca) sorte.add(r.sortaVoca);
    });
    
    const kupacName = document.getElementById('fldZbirnaKupac').selectedOptions[0].textContent;
    
    const record = {
        clientRecordID: crypto.randomUUID(),
        createdAtClient: new Date().toISOString(),
        vozacID: CONFIG.ENTITY_ID,
        datum: today,
        kupacID: kupacID,
        kupacName: kupacName,
        vrstaVoca: [...vrste].join(', '),
        sortaVoca: [...sorte].join(', '),
        kolicinaKlI: totalKgI,
        kolicinaKlII: totalKgII,
        tipAmbalaze: todayOtkupi[0].tipAmbalaze || '',
        kolAmbalaze: totalAmb,
        klasa: totalKgII > 0 ? 'I+II' : 'I',
        otkupRecordIDs: todayOtkupi.map(r => r.clientRecordID).join(','),
        syncStatus: 'pending'
    };
    
    await dbPut(db, 'zbirne', record);
    showToast('Zbirna kreirana!', 'success');
    cancelZbirna();
    
    // Sync immediately
    if (navigator.onLine) syncZbirne();
}

function cancelZbirna() {
    document.getElementById('zbirnaCreateView').style.display = 'none';
    document.getElementById('zbirnaMainView').style.display = 'block';
    loadVozacZbirne();
}

async function loadVozacZbirne() {
    const local = await dbGetAll(db, 'zbirne');

    let server = [];
    try {
        const json = await apiFetch('action=getVozacZbirne');
        if (json && json.success && json.records) server = json.records.map(r => ({
            clientRecordID: r.ClientRecordID || '',
            datum: fmtDate(r.Datum),
            kupacName: r.KupacName || r.KupacID || '',
            kolicinaKlI: parseFloat(r.KolicinaKlI) || 0,
            kolicinaKlII: parseFloat(r.KolicinaKlII) || 0,
            kolAmbalaze: parseInt(r.KolAmbalaze) || 0,
            vrstaVoca: r.VrstaVoca || '',
            syncStatus: 'synced'
        }));
    } catch (e) {}

    const serverIDs = new Set(server.map(r => r.clientRecordID));
    const all = [...server, ...local.filter(r => r.syncStatus === 'pending' && !serverIDs.has(r.clientRecordID))];

    const list = document.getElementById('vozacZbirneList');
    if (all.length === 0) {
        list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:12px;">Nema kreiranih zbirnih</p>';
        return;
    }

    list.innerHTML = all.map(r => {
        const totalKg = (r.kolicinaKlI || 0) + (r.kolicinaKlII || 0);
        const bc = r.syncStatus === 'pending' ? 'var(--warning)' : 'var(--success)';
        return `<div class="queue-item" style="border-left-color:${bc};">
            <div class="qi-header"><span class="qi-koop">🏭 ${r.kupacName}</span><span class="qi-time">${r.datum}</span></div>
            <div class="qi-detail">${r.vrstaVoca} | ${totalKg.toLocaleString('sr')} kg | Amb: ${r.kolAmbalaze || 0}</div>
            ${r.kolicinaKlII > 0 ? '<div class="qi-detail" style="font-size:11px;">Kl.I: ' + (r.kolicinaKlI||0).toLocaleString('sr') + ' kg | Kl.II: ' + (r.kolicinaKlII||0).toLocaleString('sr') + ' kg</div>' : ''}
        </div>`;
    }).join('');
}

async function syncZbirne() {
    const pending = await dbGetByIndex(db, 'zbirne', 'syncStatus', 'pending');
    if (pending.length === 0) return;

    const json = await apiPost('syncZbirna', {
        vozacID: CONFIG.ENTITY_ID,
        records: pending
    });

    if (json && json.success) {
        for (const r of pending) {
            r.syncStatus = 'synced';
            await dbPut(db, 'zbirne', r);
        }
        showToast('Zbirna sinhronizovana', 'success');
    }
}

    async function loadVozacTransport() {
    const list = document.getElementById('transportList');
    const local = await dbGetAll(db, 'zbirne');
    if (local.length === 0) {
        list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema transporta</p>';
        return;
    }
    list.innerHTML = local.map(r => {
        const totalKg = (r.kolicinaKlI || 0) + (r.kolicinaKlII || 0);
        return `<div class="queue-item">
            <div class="qi-header"><span class="qi-koop">🏭 ${r.kupacName || r.kupacID}</span><span class="qi-time">${r.datum}</span></div>
            <div class="qi-detail">${totalKg.toLocaleString('sr')} kg | Amb: ${r.kolAmbalaze || 0} | ${r.syncStatus === 'synced' ? '✅' : '⏳'}</div>
        </div>`;
    }).join('');
}
    
// ============================================================
// KOOPERANT: KARTICA
// ============================================================
let karticaCache = null;

async function loadKartica() {
    document.getElementById('karticaName').textContent = CONFIG.ENTITY_NAME;
    document.getElementById('karticaID').textContent = CONFIG.ENTITY_ID;
    
    if (karticaCache) {
        renderKartica(karticaCache);
        return;
    }
    
    document.getElementById('karticaList').innerHTML = '<p style="text-align:center;padding:20px;color:var(--text-muted);">Učitavanje...</p>';
    
    let records = [];
    try {
        const json = await apiFetch('action=getKartica&kooperantID=' + encodeURIComponent(CONFIG.ENTITY_ID));
        if (json && json.success && json.records) records = json.records.filter(r => r.Opis !== 'UKUPNO');
    } catch (e) {}
    
    karticaCache = records;
    renderKartica(records);
}

function renderKartica(records) {
    if (records.length === 0) {
        document.getElementById('karticaList').innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:40px;">Kartica nije dostupna</p>';
        ['karticaZaduzenje','karticaRazduzenje','karticaSaldo'].forEach(id => document.getElementById(id).textContent = '0');
        return;
    }
    let zad = 0, raz = 0;
    document.getElementById('karticaList').innerHTML = records.map(r => {
        const z = parseFloat(r.Zaduzenje)||0, ra = parseFloat(r.Razduzenje)||0, s = parseFloat(r.Saldo)||0;
        zad += z; raz += ra;
        return `<div class="queue-item" style="border-left-color:${z>0?'var(--danger)':'var(--success)'};">
            <div class="qi-header"><span class="qi-koop">${r.BrojDok||''}</span><span class="qi-time">${fmtDate(r.Datum)}</span></div>
            <div class="qi-detail">${r.Opis||''}</div>
            <div class="qi-detail" style="font-size:12px;margin-top:2px;">
                ${z>0?'<span style="color:var(--danger);">Zaduž: '+z.toLocaleString('sr')+'</span> ':''}
                ${ra>0?'<span style="color:var(--success);">Razduž: '+ra.toLocaleString('sr')+'</span> ':''}
                | Saldo: <strong>${s.toLocaleString('sr')}</strong></div></div>`;
    }).join('');
    document.getElementById('karticaZaduzenje').textContent = zad.toLocaleString('sr');
    document.getElementById('karticaRazduzenje').textContent = raz.toLocaleString('sr');
    document.getElementById('karticaSaldo').textContent = (zad - raz).toLocaleString('sr');
}

// ============================================================
// KOOPERANT: PARCELE
// ============================================================
let parcelMapInstance = null;

const kooperantParcelStyle = {
    color: '#ffd60a',
    weight: 3,
    opacity: 1,
    fillColor: '#ffd60a',
    fillOpacity: 0.18
};

const kooperantSelectedParcelStyle = {
    color: '#ff2d55',
    weight: 4,
    opacity: 1,
    fillColor: '#ff2d55',
    fillOpacity: 0.22
};

let activeKooperantParcelLayer = null;

function resetKooperantParcelHighlight() {
    if (activeKooperantParcelLayer && activeKooperantParcelLayer.setStyle) {
        activeKooperantParcelLayer.setStyle(kooperantParcelStyle);
    }
    activeKooperantParcelLayer = null;
}

function highlightKooperantParcelLayer(layer) {
    resetKooperantParcelHighlight();
    if (layer && layer.setStyle) {
        layer.setStyle(kooperantSelectedParcelStyle);
        activeKooperantParcelLayer = layer;
    }
}

function buildKooperantParcelPopup(p) {
    return `
        <div>
            <div style="font-size:18px;font-weight:700;margin-bottom:6px;">
                ${p.KatBroj || p.ParcelaID}
            </div>
            <div><b>Kultura:</b> ${p.Kultura || '-'}</div>
            <div><b>Površina:</b> ${p.PovrsinaHa || '?'} ha</div>
            <div><b>KO:</b> ${p.KatOpstina || '-'}</div>
            <div><b>GGAP:</b> ${p.GGAPStatus || '-'}</div>
            <div style="margin-top:6px;color:#666;">${p.ParcelaID}</div>
        </div>
    `;
}
    
async function loadParcele() {
    const parcele = (stammdaten.parcele || []).filter(p => p.KooperantID === CONFIG.ENTITY_ID);
    const list = document.getElementById('parceleList');
    const mapDiv = document.getElementById('parceleMap');
    
    if (parcele.length === 0) {
        list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:40px;">Nema parcela</p>';
        mapDiv.style.display = 'none';
        return;
    }
    
    // Render list with loading placeholders
    list.innerHTML = parcele.map(p =>
        `<div id="parcel-card-${p.ParcelaID}" style="background:white;border-radius:var(--radius);padding:14px;margin-bottom:8px;box-shadow:0 1px 3px rgba(0,0,0,0.08);border-left:4px solid var(--primary);cursor:pointer;" onclick="focusParcel('${p.ParcelaID}')">
            <div style="display:flex;justify-content:space-between;margin-bottom:4px;">
                <strong>${p.KatBroj || p.ParcelaID}</strong>
                <span style="font-size:12px;color:var(--text-muted);">${p.ParcelaID}</span>
            </div>
            <div style="font-size:13px;color:var(--text-muted);margin-bottom:6px;">${p.Kultura || ''} | ${p.PovrsinaHa || '?'} ha | ${p.KatOpstina || ''}${p.GGAPStatus ? ' | GGAP: ' + p.GGAPStatus : ''}</div>
            <div id="parcel-meteo-${p.ParcelaID}" style="font-size:12px;color:var(--text-muted);">⏳ Meteo...</div>
        </div>`).join('');
    
    // Init map
    if (parcelMapInstance) { parcelMapInstance.remove(); parcelMapInstance = null; }
    parcelMapInstance = L.map(mapDiv).setView([43.28, 21.72], 13);
    L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', {
        maxZoom: 22,
        attribution: 'Esri, Maxar, Earthstar Geographics'
    }).addTo(parcelMapInstance);
    
    // Load all parcel geo data + meteo
    const allBounds = [];
    window._parcelLayers = {};
    
    for (const p of parcele) {
        try {
            const resp = await fetch(CONFIG.API_URL + '?action=getParcelGeo&parcelaId=' + encodeURIComponent(p.ParcelaID));
            const json = await resp.json();
            
            if (json && json.success && json.parcel) {
                const geo = json.parcel;
                const lat = parseFloat(String(geo.Lat).replace(',', '.'));
                const lng = parseFloat(String(geo.Lng).replace(',', '.'));
                const popupHtml = buildKooperantParcelPopup(p);
                
                if (geo.PolygonGeoJSON) {
                    const geometry = JSON.parse(geo.PolygonGeoJSON);
                    const feature = {type: 'Feature', properties: p, geometry: geometry};
                    const layer = L.geoJSON(feature, { style: kooperantParcelStyle }).addTo(parcelMapInstance);
                    layer.eachLayer(l => {
                        l.bindTooltip(`${p.KatBroj || p.ParcelaID}`, { permanent: true, direction: 'center', className: 'parcel-label' });
                        l.bindPopup(popupHtml);
                        l.on('click', () => { highlightKooperantParcelLayer(l); });
                        const bounds = l.getBounds();
                        if (bounds.isValid()) allBounds.push(bounds);
                        window._parcelLayers[p.ParcelaID] = l;
                    });
                } else if (lat && lng && !isNaN(lat) && !isNaN(lng)) {
                    const marker = L.circleMarker([lat, lng], {
                        radius: 8, color: '#ffd60a', weight: 3, fillColor: '#ffd60a', fillOpacity: 0.85
                    }).addTo(parcelMapInstance);
                    marker.bindTooltip(`${p.KatBroj || p.ParcelaID}`, { permanent: true, direction: 'top', className: 'parcel-label' });
                    marker.bindPopup(popupHtml);
                    allBounds.push(L.latLngBounds([marker.getLatLng(), marker.getLatLng()]));
                    window._parcelLayers[p.ParcelaID] = marker;
                }
            }
        } catch (e) {}
        
        // Load meteo for this parcel (non-blocking per parcel)
        loadParcelMeteoInline(p.ParcelaID, p.Kultura || '');
    }

    if (allBounds.length > 0) {
        let combined = allBounds[0];
        for (let i = 1; i < allBounds.length; i++) {
            combined.extend(allBounds[i]);
        }
        parcelMapInstance.fitBounds(combined.pad(0.2));
    }
}


// ============================================================
// KOOPERANT: PARCELA METEO + RISK
// ============================================================
let meteoCache = {};

async function loadParcelMeteo(parcelaId, kultura) {
    const panel = document.getElementById('parceleMeteo');
    panel.style.display = 'block';
    panel.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:12px;">Učitavanje meteo podataka...</p>';
    
    // Check local cache
    if (meteoCache[parcelaId] && (Date.now() - meteoCache[parcelaId]._ts < 3600000)) {
        renderMeteoPanel(meteoCache[parcelaId]);
        return;
    }
    
    try {
        const url = CONFIG.API_URL + '?action=getParcelMeteo&parcelaId=' + encodeURIComponent(parcelaId);
        const resp = await fetch(url);
        const json = await resp.json();
        
        if (json && json.success) {
            json._ts = Date.now();
            meteoCache[parcelaId] = json;
            renderMeteoPanel(json);
        } else {
            panel.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:12px;">' + (json.error || 'Nema meteo podataka') + '</p>';
        }
    } catch (e) {
        panel.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:12px;">Greška pri učitavanju</p>';
    }
}

function renderMeteoPanel(data) {
    const panel = document.getElementById('parceleMeteo');
    const c = data.current || {};
    const risk = data.risk || {};
    const spray = data.sprayWindow || [];
    const daily = data.daily || data.ForecastDaily || [];
    const parcelaId = data.parcelaId || '';

    if (panel.dataset) panel.dataset.currentParcelaId = parcelaId;

    panel.innerHTML = `
        <div class="meteo-panel">
            <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">
                <div style="font-size:14px;font-weight:700;color:var(--primary);">
                    ${data.katBroj || data.parcelaId} — ${data.kultura || ''}
                </div>
                <div style="font-size:10px;color:var(--text-muted);">
                    ${new Date(data.fetchedAt).toLocaleTimeString('sr', {hour:'2-digit',minute:'2-digit'})}
                </div>
            </div>
            
            <div class="meteo-current">
                <div>
                    <div class="meteo-temp">${Number(c.temperature || 0).toFixed(1)}°</div>
                </div>
                <div class="meteo-details">
                    <div>${weatherCodeText(c.weatherCode || 0)}</div>
                    <div>💧 Vlažnost: ${c.humidity || 0}%</div>
                    <div>💨 Vetar: ${Number(c.windSpeed || 0).toFixed(1)} km/h (udari: ${Number(c.windGusts || 0).toFixed(1)})</div>
                    ${Number(c.precipitation || 0) > 0 ? '<div>🌧️ Padavine: ' + Number(c.precipitation).toFixed(1) + ' mm</div>' : ''}
                </div>
            </div>
            
            ${renderRiskSection(risk)}
            ${renderSpraySection(spray, data.kultura)}
            ${renderForecast(daily, ['Ned', 'Pon', 'Uto', 'Sre', 'Čet', 'Pet', 'Sub'])}
            ${renderExpertInfo(parcelaId, c)}
        </div>
    `;
}
function renderRiskSection(risk) {
    if (!risk || !risk.items || risk.items.length === 0) {
        return '<div class="meteo-risk ok">✅ Nema rizika — uslovi su povoljni</div>';
    }
    
    return '<div style="margin-bottom:10px;">' +
        '<div style="font-size:12px;font-weight:600;color:var(--text-muted);margin-bottom:6px;">UPOZORENJA</div>' +
        risk.items.map(r =>
            `<div class="meteo-risk ${r.level}">
                <span style="font-size:18px;">${r.icon}</span>
                <span>${r.message}</span>
            </div>`
        ).join('') +
    '</div>';
}

function renderSpraySection(windows, kultura) {
    let html = '<div style="margin-bottom:10px;">';
    html += '<div style="font-size:12px;font-weight:600;color:var(--text-muted);margin-bottom:6px;">PROZOR ZA PRSKANJE</div>';
    
    if (!windows || windows.length === 0) {
        html += '<div class="spray-window" style="background:#fef3c7;border-color:#fcd34d;">⚠️ Nema pogodnog termina za prskanje u naredna 72h</div>';
    } else {
        windows.forEach((w, i) => {
            const start = new Date(w.start);
            const end = new Date(w.end);
            const startStr = start.toLocaleDateString('sr', {weekday:'short', day:'numeric', month:'short'}) + ' ' +
                           start.toLocaleTimeString('sr', {hour:'2-digit', minute:'2-digit'});
            const endStr = end.toLocaleTimeString('sr', {hour:'2-digit', minute:'2-digit'});
            
            html += `<div class="spray-window">
                <div class="spray-time">${i === 0 ? '✅ ' : ''}${startStr} — ${endStr} (${w.hours}h)</div>
                <div class="spray-details">Temp: ${w.avgTemp}°C | Vetar: ${w.avgWind} km/h | Vlažnost: ${w.avgHumidity}%</div>
            </div>`;
        });
    }
    
    html += '</div>';
    return html;
}

function renderForecast(daily, dayNames) {
    if (!daily || daily.length === 0) return '';

    const first3 = daily.slice(0, 3);

    let html = '<div style="font-size:12px;font-weight:600;color:var(--text-muted);margin-bottom:6px;">PROGNOZA — NAREDNA 3 DANA</div>';
    html += '<div class="meteo-forecast-3d">';

    first3.forEach((d, i) => {
        const date = new Date(d.date);
        const dayName = i === 0 ? 'Danas' : dayNames[date.getDay()];
        const icon = weatherCodeIcon(d.weatherCode);

        html += `
            <div class="meteo-day-3d">
                <div class="day-name">${dayName}</div>
                <div class="day-icon">${icon}</div>
                <div class="day-temp">${Math.round(d.tempMax)}°</div>
                <div class="day-temp-min">${Math.round(d.tempMin)}°</div>
                <div class="day-rain">${d.precipSum > 0 ? d.precipSum.toFixed(1) + ' mm' : '&nbsp;'}</div>
            </div>
        `;
    });

    html += '</div>';
    return html;
}

function weatherCodeText(code) {
    const codes = {
        0: 'Vedro', 1: 'Pretežno vedro', 2: 'Delimično oblačno', 3: 'Oblačno',
        45: 'Magla', 48: 'Magla sa mrazom',
        51: 'Slaba kiša', 53: 'Umerena kiša', 55: 'Jaka kiša',
        61: 'Slaba kiša', 63: 'Umerena kiša', 65: 'Jaka kiša',
        71: 'Slab sneg', 73: 'Umeren sneg', 75: 'Jak sneg',
        80: 'Pljuskovi', 81: 'Umereni pljuskovi', 82: 'Jaki pljuskovi',
        95: 'Grmljavina', 96: 'Grmljavina sa gradom', 99: 'Jaka grmljavina'
    };
    return codes[code] || 'Nepoznato';
}

function weatherCodeIcon(code) {
    if (code === 0) return '☀️';
    if (code <= 3) return '⛅';
    if (code <= 48) return '🌫️';
    if (code <= 65) return '🌧️';
    if (code <= 75) return '❄️';
    if (code <= 82) return '🌦️';
    if (code >= 95) return '⛈️';
    return '🌤️';
}

async function loadParcelMeteoInline(parcelaId, kultura) {
    const el = document.getElementById('parcel-meteo-' + parcelaId);
    if (!el) return;
    
    // Check local cache
    if (meteoCache[parcelaId] && (Date.now() - meteoCache[parcelaId]._ts < 3600000)) {
        el.innerHTML = renderMeteoInline(meteoCache[parcelaId]);
        return;
    }
    
    try {
        const url = CONFIG.API_URL + '?action=getParcelMeteo&parcelaId=' + encodeURIComponent(parcelaId);
        const resp = await fetch(url);
        const json = await resp.json();
        
        if (json && json.success) {
            json._ts = Date.now();
            meteoCache[parcelaId] = json;
            el.innerHTML = renderMeteoInline(json);
        } else {
            el.innerHTML = '<span style="color:var(--text-muted);">Nema meteo podataka</span>';
        }
    } catch (e) {
        el.innerHTML = '<span style="color:var(--text-muted);">—</span>';
    }
}

function renderMeteoInline(data) {
    const c = data.current || {};
    const risk = data.risk || {};
    const spray = data.sprayWindow || data.SprayWindows || [];
    const daily = data.daily || data.ForecastDaily || [];

    const temp = Number(c.temperature || c.Temp || 0).toFixed(1);
    const hum = Number(c.humidity || c.Humidity || 0).toFixed(0);
    const wind = Number(c.windSpeed || c.Wind || 0).toFixed(0);
    const icon = weatherCodeIcon(c.weatherCode || c.WeatherCode || 0);

    let riskHtml = '<span class="parcel-chip ok">✅ Bez rizika</span>';
    const riskItems = risk ? (risk.items || risk.RiskItems || []) : [];
    if (riskItems.length > 0) {
        const first = riskItems[0];
        const cls = first.level === 'danger' ? 'danger' : 'warn';
        riskHtml = `<span class="parcel-chip ${cls}">${first.icon} ${first.message}</span>`;
    }

    let sprayHtml = '<span class="parcel-chip warn">⚠️ Nema termina za prskanje</span>';
    if (spray.length > 0) {
        const w = spray[0];
        const start = new Date(w.start);
        const end = new Date(w.end);
        const dayNames = ['Ned','Pon','Uto','Sre','Čet','Pet','Sub'];
        const dayStr = dayNames[start.getDay()];
        const startTime = start.toLocaleTimeString('sr', { hour:'2-digit', minute:'2-digit' });
        const endTime = end.toLocaleTimeString('sr', { hour:'2-digit', minute:'2-digit' });
        sprayHtml = `<span class="parcel-chip ok">🎯 ${dayStr} ${startTime}-${endTime} (${w.hours}h)</span>`;
    }

    let forecastHtml = '';
    if (daily.length > 0) {
        const dayNames = ['Ned','Pon','Uto','Sre','Čet','Pet','Sub'];
        forecastHtml = `
            <div class="parcel-forecast-inline">
                ${daily.slice(0, 3).map((d, i) => {
                    const dt = new Date(d.date);
                    const name = i === 0 ? 'Danas' : dayNames[dt.getDay()];
                    return `
                        <span class="parcel-forecast-day">
                            <strong>${name}</strong>
                            ${weatherCodeIcon(d.weatherCode || 0)}
                            ${Math.round(Number(d.tempMax || 0))}°/${Math.round(Number(d.tempMin || 0))}°
                            ${Number(d.precipSum || 0) > 0 ? '💧' + Number(d.precipSum).toFixed(1) : ''}
                        </span>
                    `;
                }).join('')}
            </div>
        `;
    }

    return `
        <div class="parcel-meteo-compact">
            <div class="parcel-meteo-topline">
                <span class="parcel-chip temp">${temp}°C</span>
                <span class="parcel-chip">${icon}</span>
                <span class="parcel-chip">💧 ${hum}%</span>
                <span class="parcel-chip">💨 ${wind} km/h</span>
                ${riskHtml}
            </div>
            <div class="parcel-meteo-midline">
                ${sprayHtml}
            </div>
            <div class="parcel-meteo-bottomline">
                ${forecastHtml}
            </div>
        </div>
    `;
}

function renderExpertInfo(parcelId, current) {
    const soilMoist = current.soilMoist_0_1 ?? current.SoilMoist_0_1cm ?? null;
    const soilTemp = current.soilTemp_0 ?? current.SoilTemp_0cm ?? null;
    const et0 = current.et0 ?? current.ET0 ?? null;
    const uv = current.uvIndex ?? current.UVIndex ?? null;
    const solar = current.solarRadiation ?? current.SolarRadiation ?? null;

    const hasExpert =
        soilMoist !== null ||
        soilTemp !== null ||
        et0 !== null ||
        uv !== null ||
        solar !== null;

    if (!hasExpert) return '';

    const isOpen = !!parcelExpertOpen[parcelId];

    let html = `
        <button class="parcel-expert-toggle" onclick="toggleParcelExpert('${parcelId}')">
            <div>
                <div class="parcel-expert-title">
                    <span>🧪 Expert info</span>
                </div>
                <div class="parcel-expert-sub">Zemljište, ET₀, UV i dodatni agro podaci</div>
            </div>
            <span class="parcel-expert-chevron ${isOpen ? 'open' : ''}">⌄</span>
        </button>
    `;

    if (!isOpen) return html;

    html += `
        <div class="parcel-expert-panel">
            <div class="parcel-expert-grid">
                ${soilTemp !== null ? `
                    <div class="parcel-expert-item">
                        <div class="parcel-expert-k">🌡️ Temperatura zemljišta</div>
                        <div class="parcel-expert-v">${Number(soilTemp).toFixed(1)}°C</div>
                    </div>
                ` : ''}

                ${soilMoist !== null ? `
                    <div class="parcel-expert-item">
                        <div class="parcel-expert-k">🌱 Vlažnost zemljišta</div>
                        <div class="parcel-expert-v">${(Number(soilMoist) * 100).toFixed(0)}%</div>
                    </div>
                ` : ''}

                ${et0 !== null ? `
                    <div class="parcel-expert-item">
                        <div class="parcel-expert-k">💦 ET₀</div>
                        <div class="parcel-expert-v">${Number(et0).toFixed(1)} mm</div>
                    </div>
                ` : ''}

                ${uv !== null && Number(uv) > 0 ? `
                    <div class="parcel-expert-item">
                        <div class="parcel-expert-k">☀️ UV indeks</div>
                        <div class="parcel-expert-v">${Number(uv).toFixed(1)}</div>
                    </div>
                ` : ''}

                ${solar !== null && Number(solar) > 0 ? `
                    <div class="parcel-expert-item">
                        <div class="parcel-expert-k">🔆 Solarno zračenje</div>
                        <div class="parcel-expert-v">${Number(solar).toFixed(0)} W/m²</div>
                    </div>
                ` : ''}
            </div>
        </div>
    `;

    return html;
}

function toggleParcelExpert(parcelId) {
    parcelExpertOpen[parcelId] = !parcelExpertOpen[parcelId];

    const panel = document.getElementById('parceleMeteo');
    if (panel && panel.dataset && panel.dataset.currentParcelaId === parcelId) {
        const cached = meteoCache[parcelId];
        if (cached) renderMeteoPanel(cached);
    }

    const parcela = (stammdaten.parcele || []).find(p => p.ParcelaID === parcelId);
    if (parcela) {
        loadParcelMeteoInline(parcelId, parcela.Kultura || '');
    }
}

function focusParcel(parcelaID) {
    if (!parcelMapInstance || !window._parcelLayers || !window._parcelLayers[parcelaID]) return;

    const layer = window._parcelLayers[parcelaID];

    if (layer.getBounds) {
        parcelMapInstance.fitBounds(layer.getBounds().pad(0.3));
        highlightKooperantParcelLayer(layer);
    } else if (layer.getLatLng) {
        parcelMapInstance.setView(layer.getLatLng(), 17);
        resetKooperantParcelHighlight();
    }

    if (layer.openPopup) layer.openPopup();

    document.getElementById('parceleMap').scrollIntoView({ behavior: 'smooth' });
}
// ============================================================
// KOOPERANT: AGROMERE
// ============================================================
// ============================================================
// DIGITALNI AGRONOM — Kooperant Agromere Tab
// ============================================================
let agroState = {
    parcelaID: '',
    parcelaData: null,
    mera: '',
    artikalID: '',
    artikalData: null,
    kolicina: 0,
    dozaPreporucena: 0,
    opremaTraktor: '',
    opremaPrskalica: '',
    opremaOstalo: '',
    napomena: '',
    timerStart: null,
    timerInterval: null,
    timerResult: null,
    geoStart: null,
    geoEnd: null,
    geoAutoDetect: false,
    meteoOverride: false,
    meteoSnapshot: null,
    karencaDana: 0,
    lager: [],        // artikli na lageru ovog kooperanta
    opremaList: [],   // oprema ovog kooperanta
    geoWatchId: null
};

// --- Baza čestih naziva opreme za autocomplete ---
const OPREMA_PREDLOZI = {
    Traktor: ['IMT 533', 'IMT 539', 'IMT 542', 'IMT 560', 'IMT 577',
              'John Deere', 'New Holland', 'Massey Ferguson', 'Zetor', 'Ursus',
              'Torpedo', 'Rakovica 65', 'Rakovica 76', 'Belarus', 'Tomo Vinković'],
    Prskalica: ['Agrip 200L', 'Agrip 400L', 'Agrip 600L', 'Morava 440',
                'Holder', 'Stihl SR', 'Atomizer Cifarelli', 'Leđna prskalica',
                'Turbo atomizer', 'Vučena prskalica']
};

// ============================================================
// INIT — poziva se iz showTab('agromere')
// ============================================================
async function loadAgronom() {
    // Reset state
    agroResetState();

    // Load lager (iz stammdaten.magacinkoop)
    agroLoadLager();

    // Load oprema (sa servera + lokalna)
    await agroLoadOprema();

    // Populate parcele dropdown
    agroPopulateParcele();

    // Start GPS
    agroStartGeo();

    // Load istorija
    agroLoadIstorija();

    // Show step 1
    document.getElementById('agroStep1').style.display = 'block';
    document.getElementById('agroStep2').style.display = 'none';
}

function agroResetState() {
    if (agroState.timerInterval) clearInterval(agroState.timerInterval);
    if (agroState.geoWatchId) navigator.geolocation.clearWatch(agroState.geoWatchId);

    agroState = {
        parcelaID: '', parcelaData: null, mera: '', artikalID: '', artikalData: null,
        kolicina: 0, dozaPreporucena: 0, opremaTraktor: '', opremaPrskalica: '',
        opremaOstalo: '', napomena: '', timerStart: null, timerInterval: null,
        timerResult: null, geoStart: null, geoEnd: null, geoAutoDetect: false,
        meteoOverride: false, meteoSnapshot: null, karencaDana: 0,
        lager: agroState.lager || [], opremaList: agroState.opremaList || [],
        geoWatchId: null
    };

    // Reset UI
    document.querySelectorAll('.agro-mera-btn').forEach(b => b.classList.remove('selected'));
    document.getElementById('agroPreparatSection').classList.remove('visible');
    document.getElementById('agroPreporuka').classList.remove('visible');
    document.getElementById('agroKarencaWarn').classList.remove('visible');
    document.getElementById('agroMeteoWarn').classList.remove('visible');
    document.getElementById('agroTimerPanel').style.display = 'none';
    document.getElementById('agroBtnStart').style.display = 'none';
    document.getElementById('agroBtnStop').style.display = 'none';
    document.getElementById('agroTimerSticky').classList.remove('active');
    const karInfo = document.getElementById('agroKarencaInfo');
    if (karInfo) karInfo.style.display = 'none';
}

// ============================================================
// LAGER — iz stammdaten.magacinkoop
// ============================================================
function agroLoadLager() {
    const koopID = CONFIG.ENTITY_ID;
    const mk = stammdaten.magacinkoop || [];
    agroState.lager = mk.filter(r => r.KooperantID === koopID && parseFloat(r.Stanje) > 0)
        .map(r => ({
            artikalID: r.ArtikalID,
            naziv: r.ArtikalNaziv || r.Naziv || r.ArtikalID,
            tip: r.Tip || '',
            jm: r.JedinicaMere || 'kg',
            cena: parseFloat(r.CenaPoJedinici) || 0,
            dozaPoHa: parseFloat(String(r.DozaPoHa || '0').replace(',', '.')) || 0,
            pakovanje: parseFloat(String(r.Pakovanje || '0').replace(',', '.')) || 0,
            karencaDana: parseInt(r.KarencaDana) || 0,
            primljeno: parseFloat(r.Primljeno) || 0,
            utroseno: parseFloat(r.Utroseno) || 0,
            stanje: parseFloat(r.Stanje) || 0
        }));
}

// ============================================================
// OPREMA — sa servera + lokalna
// ============================================================
async function agroLoadOprema() {
    agroState.opremaList = [];
    try {
        const json = await apiFetch('action=getOprema&kooperantID=' + encodeURIComponent(CONFIG.ENTITY_ID));
        if (json && json.success && json.records) {
            agroState.opremaList = json.records.map(r => ({
                naziv: r.Naziv || '', tip: r.Tip || ''
            }));
        }
    } catch(e) {}

    agroPopulateOpremaDropdowns();
}

function agroPopulateOpremaDropdowns() {
    const tSel = document.getElementById('agroTraktor');
    const pSel = document.getElementById('agroPrskalica');
    if (!tSel || !pSel) return;

    tSel.innerHTML = '<option value="">-- Bez traktora --</option>';
    pSel.innerHTML = '<option value="">-- Bez prskalice --</option>';

    // Kooperantova oprema (sa servera)
    const traktori = agroState.opremaList.filter(o => o.tip === 'Traktor');
    const prskalice = agroState.opremaList.filter(o => o.tip === 'Prskalica' || o.tip === 'Atomizer');

    traktori.forEach(o => {
        const op = document.createElement('option'); op.value = o.naziv; op.textContent = o.naziv;
        tSel.appendChild(op);
    });

    prskalice.forEach(o => {
        const op = document.createElement('option'); op.value = o.naziv; op.textContent = o.naziv;
        pSel.appendChild(op);
    });

    // Predlozi (samo oni koje kooperant još nema)
    const tNames = new Set(traktori.map(o => o.naziv));
    const pNames = new Set(prskalice.map(o => o.naziv));

    if (traktori.length === 0) {
        const og = document.createElement('optgroup'); og.label = '— Česti modeli —';
        OPREMA_PREDLOZI.Traktor.forEach(n => {
            if (!tNames.has(n)) { const op = document.createElement('option'); op.value = n; op.textContent = n; og.appendChild(op); }
        });
        if (og.children.length > 0) tSel.appendChild(og);
    }

    if (prskalice.length === 0) {
        const og = document.createElement('optgroup'); og.label = '— Česti modeli —';
        OPREMA_PREDLOZI.Prskalica.forEach(n => {
            if (!pNames.has(n)) { const op = document.createElement('option'); op.value = n; op.textContent = n; og.appendChild(op); }
        });
        if (og.children.length > 0) pSel.appendChild(og);
    }
}

async function agroSaveNovaOprema(tip, naziv) {
    if (!naziv || !naziv.trim()) return;
    naziv = naziv.trim();

    // Sačuvaj na server
    try {
        await fetch(CONFIG.API_URL, {
            method: 'POST', headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify({
                action: 'syncOprema', token: CONFIG.TOKEN,
                kooperantID: CONFIG.ENTITY_ID,
                records: [{ clientRecordID: crypto.randomUUID(), naziv: naziv, tip: tip }]
            })
        });
    } catch(e) {}

    // Dodaj u lokalni niz i osveži dropdown
    agroState.opremaList.push({ naziv: naziv, tip: tip });
    agroPopulateOpremaDropdowns();

    // Auto-select
    if (tip === 'Traktor') document.getElementById('agroTraktor').value = naziv;
    if (tip === 'Prskalica') document.getElementById('agroPrskalica').value = naziv;

    showToast('Oprema sačuvana: ' + naziv, 'success');
}

// ============================================================
// PARCELA
// ============================================================
function agroPopulateParcele() {
    const sel = document.getElementById('agroParcelaSel');
    if (!sel) return;
    sel.innerHTML = '<option value="">-- Izaberi parcelu --</option>';
    (stammdaten.parcele || []).filter(p => p.KooperantID === CONFIG.ENTITY_ID).forEach(p => {
        const ha = parseFloat(String(p.PovrsinaHa || '0').replace(',', '.')) || 0;
        const o = document.createElement('option');
        o.value = p.ParcelaID;
        o.textContent = (p.KatBroj || p.ParcelaID) + ' — ' + (p.Kultura || '?') + ' (' + ha.toFixed(2) + ' ha)';
        sel.appendChild(o);
    });
}

function onAgroParcelaChange() {
    const pid = document.getElementById('agroParcelaSel').value;
    agroState.parcelaID = pid;
    agroState.parcelaData = (stammdaten.parcele || []).find(p => p.ParcelaID === pid) || null;

    // Meteo strip
    if (pid) {
        loadAgroMeteoStrip(pid);
        checkAgroKarenca(pid);
    } else {
        document.getElementById('agroMeteoStrip').style.display = 'none';
        document.getElementById('agroKarencaWarn').classList.remove('visible');
    }

    // Reset mera
    agroState.mera = '';
    document.querySelectorAll('.agro-mera-btn').forEach(b => b.classList.remove('selected'));
    document.getElementById('agroPreparatSection').classList.remove('visible');
    document.getElementById('agroBtnStart').style.display = pid ? 'none' : 'none';
}

async function loadAgroMeteoStrip(parcelaID) {
    const strip = document.getElementById('agroMeteoStrip');
    strip.style.display = 'flex';

    // Iz meteoCache ako postoji
    let data = null;
    if (meteoCache[parcelaID] && (Date.now() - meteoCache[parcelaID]._ts < 3600000)) {
        data = meteoCache[parcelaID];
    } else {
        try {
            const url = CONFIG.API_URL + '?action=getParcelMeteo&parcelaId=' + encodeURIComponent(parcelaID);
            const resp = await fetch(url);
            const json = await resp.json();
            if (json && json.success) { json._ts = Date.now(); meteoCache[parcelaID] = json; data = json; }
        } catch(e) {}
    }

    if (!data || !data.current) {
        strip.innerHTML = '<span>Nema meteo podataka</span>';
        agroState.meteoSnapshot = null;
        return;
    }

    const c = data.current;
    const temp = (c.temperature || c.Temp || 0);
    const wind = (c.windSpeed || c.Wind || 0);
    const hum = (c.humidity || c.Humidity || 0);
    const precip = (c.precipitation || c.Precip || 0);

    agroState.meteoSnapshot = { temp: temp, wind: wind, humidity: hum };

    document.getElementById('agroMeteoTemp').textContent = '🌡️ ' + temp.toFixed(1) + '°C';
    document.getElementById('agroMeteoWind').textContent = '💨 ' + wind.toFixed(0) + ' km/h';
    document.getElementById('agroMeteoHumidity').textContent = '💧 ' + hum + '%';
    document.getElementById('agroMeteoPrecip').textContent = precip > 0 ? '🌧️ ' + precip.toFixed(1) + 'mm' : '☀️ Suvo';
}

// ============================================================
// KARENCA CHECK — za izabranu parcelu
// ============================================================
async function checkAgroKarenca(parcelaID) {
    const warn = document.getElementById('agroKarencaWarn');
    const berbaBtn = document.getElementById('agroBerbaBtn');
    warn.classList.remove('visible');
    berbaBtn.classList.remove('disabled');

    // Čitaj tretmane sa servera
    let tretmani = [];
    try {
        const json = await apiFetch('action=getTretmani&kooperantID=' + encodeURIComponent(CONFIG.ENTITY_ID));
        if (json && json.success) tretmani = json.records || [];
    } catch(e) {}

    // Dodaj lokalne pending
    try {
        const local = await dbGetAll(db, 'tretmani');
        tretmani = [...tretmani, ...local.filter(r => r.syncStatus === 'pending')];
    } catch(e) {}

    // Nađi poslednji tretman sa karencom za ovu parcelu
    const parcelTretmani = tretmani.filter(t =>
        t.ParcelaID === parcelaID && parseInt(t.KarencaDana) > 0
    );

    if (parcelTretmani.length === 0) return;

    // Sortiraj po datumu desc
    parcelTretmani.sort((a, b) => (b.Datum || b.datum || '').localeCompare(a.Datum || a.datum || ''));
    const last = parcelTretmani[0];
    const datum = last.Datum || last.datum;
    const karDana = parseInt(last.KarencaDana || last.karencaDana) || 0;
    const prepNaziv = last.ArtikalNaziv || last.artikalNaziv || '?';

    const tretmanDate = new Date(datum);
    const berbeDate = new Date(tretmanDate.getTime() + karDana * 24 * 60 * 60 * 1000);
    const today = new Date();
    today.setHours(0,0,0,0);

    if (berbeDate > today) {
        const ostalo = Math.ceil((berbeDate - today) / (24 * 60 * 60 * 1000));
        warn.classList.add('visible');
        document.getElementById('agroKarencaText').innerHTML =
            '<strong>' + prepNaziv + '</strong> — tretman ' + datum +
            '<br>Berba dozvoljena: <strong>' + berbeDate.toLocaleDateString('sr') + '</strong> (još ' + ostalo + ' dana)';
        berbaBtn.classList.add('disabled');
    }
}

// ============================================================
// MERA SELECTION
// ============================================================
function selectAgroMera(btn, mera) {
    if (btn.classList.contains('disabled')) {
        showToast('Karenca aktivna — berba nije dozvoljena', 'error');
        return;
    }

    if (!agroState.parcelaID) { showToast('Izaberite parcelu', 'error'); return; }

    document.querySelectorAll('.agro-mera-btn').forEach(b => b.classList.remove('selected'));
    btn.classList.add('selected');
    agroState.mera = mera;

    // Preparat section — samo za Zastita/Prihrana
    const prepSec = document.getElementById('agroPreparatSection');
    if (mera === 'Zastita' || mera === 'Prihrana') {
        prepSec.classList.add('visible');
        agroPopulatePreparati(mera);

        // Meteo check za Zastita
        if (mera === 'Zastita') {
            agroCheckMeteo();
        } else {
            document.getElementById('agroMeteoWarn').classList.remove('visible');
        }
    } else {
        prepSec.classList.remove('visible');
        document.getElementById('agroMeteoWarn').classList.remove('visible');
        document.getElementById('agroPreporuka').classList.remove('visible');
    }

    // Prikaži Start dugme
    document.getElementById('agroBtnStart').style.display = 'block';
    document.getElementById('agroBtnStop').style.display = 'none';
}

// ============================================================
// PREPARAT — filtrirano po tipu + lager
// ============================================================
function agroPopulatePreparati(mera) {
    const sel = document.getElementById('agroPreparatSel');
    sel.innerHTML = '<option value="">-- Izaberi preparat --</option>';
    document.getElementById('agroPreporuka').classList.remove('visible');
    const karInfo = document.getElementById('agroKarencaInfo');
    if (karInfo) karInfo.style.display = 'none';

    const tipFilter = mera === 'Zastita' ? 'Pesticid' : 'Djubrivo';

    const available = agroState.lager.filter(a => a.tip === tipFilter && a.stanje > 0);

    if (available.length === 0) {
        sel.innerHTML = '<option value="">Nema preparata na lageru</option>';
        return;
    }

    available.forEach(a => {
        const o = document.createElement('option');
        o.value = a.artikalID;
        o.textContent = a.naziv + ' — ' + a.stanje.toLocaleString('sr') + ' ' + a.jm + ' na lageru';
        sel.appendChild(o);
    });
}

function onAgroPreparatChange() {
    const artID = document.getElementById('agroPreparatSel').value;
    if (!artID) {
        document.getElementById('agroPreporuka').classList.remove('visible');
        document.getElementById('agroKarencaInfo').style.display = 'none';
        agroState.artikalID = '';
        agroState.artikalData = null;
        return;
    }

    const art = agroState.lager.find(a => a.artikalID === artID);
    agroState.artikalID = artID;
    agroState.artikalData = art;

    // JM
    document.getElementById('agroJM').value = art ? art.jm : '';

    // Karenca info
    const karInfo = document.getElementById('agroKarencaInfo');
    if (art && art.karencaDana > 0) {
        karInfo.style.display = 'block';
        const berbeDate = new Date();
        berbeDate.setDate(berbeDate.getDate() + art.karencaDana);
        document.getElementById('agroKarencaInfoText').innerHTML =
            '⏱️ Karenca: <strong>' + art.karencaDana + ' dana</strong> — berba dozvoljena od ' +
            berbeDate.toLocaleDateString('sr');
        agroState.karencaDana = art.karencaDana;
    } else {
        karInfo.style.display = 'none';
        agroState.karencaDana = 0;
    }

    // Smart dosage
    agroCalcPreporuka();
}

// ============================================================
// SMART DOSAGE
// ============================================================
function agroCalcPreporuka() {
    const panel = document.getElementById('agroPreporuka');
    const art = agroState.artikalData;
    const parcela = agroState.parcelaData;

    if (!art || !parcela || art.dozaPoHa <= 0) {
        panel.classList.remove('visible');
        return;
    }

    const ha = parseFloat(String(parcela.PovrsinaHa || '0').replace(',', '.')) || 0;
    if (ha <= 0) { panel.classList.remove('visible'); return; }

    const rawQty = art.dozaPoHa * ha;
    let finalQty = rawQty;
    let pakInfo = '';

    if (art.pakovanje > 0) {
        const pakCount = Math.ceil(rawQty / art.pakovanje);
        finalQty = pakCount;
        pakInfo = pakCount + ' × ' + art.pakovanje + ' ' + art.jm + ' (pakovanje)';
    }

    // Proveri da li ima dovoljno na lageru
    let lagerWarn = '';
    const needed = art.pakovanje > 0 ? finalQty * art.pakovanje : finalQty;
    if (needed > art.stanje) {
        lagerWarn = '<br><span style="color:var(--danger);font-weight:600;">⚠️ Nedovoljno na lageru! Imate: ' +
            art.stanje.toLocaleString('sr') + ' ' + art.jm + '</span>';
    }

    panel.classList.add('visible');
    document.getElementById('agroPreporukaCalc').innerHTML =
        '<strong>' + finalQty.toLocaleString('sr') + (art.pakovanje > 0 ? ' pak.' : ' ' + art.jm) + '</strong> — ' + art.naziv;
    document.getElementById('agroPreporukaDetail').innerHTML =
        art.dozaPoHa + ' ' + art.jm + '/ha × ' + ha.toFixed(2) + ' ha = ' +
        rawQty.toLocaleString('sr', {maximumFractionDigits:2}) + ' ' + art.jm +
        (pakInfo ? '<br>' + pakInfo : '') + lagerWarn;

    agroState.dozaPreporucena = finalQty;
    panel._finalQty = finalQty;
}

function agroPrimeniPreporuku() {
    const panel = document.getElementById('agroPreporuka');
    if (!panel || !panel._finalQty) return;
    document.getElementById('agroKolicina').value = panel._finalQty;
    showToast('Količina: ' + panel._finalQty.toLocaleString('sr'), 'success');
}

// ============================================================
// METEO VALIDATION (samo za Zastita)
// ============================================================
function agroCheckMeteo() {
    const warn = document.getElementById('agroMeteoWarn');
    warn.classList.remove('visible', 'danger');
    agroState.meteoOverride = false;

    if (!agroState.meteoSnapshot) return;

    const m = agroState.meteoSnapshot;
    const parcela = agroState.parcelaData;
    const kultura = parcela ? (parcela.Kultura || '') : '';

    // Pragovi iz CROP_THRESHOLDS (hardcoded za sad, može iz config-a)
    const thresholds = {
        'Visnja': { windMax: 15 }, 'Jabuka': { windMax: 15 }, 'Sljiva': { windMax: 15 },
        'Kruska': { windMax: 15 }, 'Breskva': { windMax: 12 }, 'Malina': { windMax: 12 },
        '_default': { windMax: 15 }
    };
    const th = thresholds[kultura] || thresholds['_default'];

    const warnings = [];

    if (m.wind > th.windMax) {
        warnings.push({ level: 'danger', text: 'Vetar ' + m.wind.toFixed(0) + ' km/h premašuje dozvoljenih ' + th.windMax + ' km/h za ' + (kultura || 'ovu kulturu') });
    }
    if (m.temp < 5) {
        warnings.push({ level: 'danger', text: 'Temperatura ' + m.temp.toFixed(1) + '°C — prenisko za prskanje (min 5°C)' });
    }
    if (m.temp > 35) {
        warnings.push({ level: 'danger', text: 'Temperatura ' + m.temp.toFixed(1) + '°C — previsoko za prskanje (max 35°C)' });
    }
    if (m.humidity > 90) {
        warnings.push({ level: 'warning', text: 'Vlažnost ' + m.humidity + '% — smanjena efikasnost preparata' });
    }

    if (warnings.length === 0) return;

    const isDanger = warnings.some(w => w.level === 'danger');
    warn.classList.add('visible');
    if (isDanger) warn.classList.add('danger');

    document.getElementById('agroMeteoWarnTitle').textContent = isDanger ? '🚫 BLOKADA — Nepovoljni uslovi' : '⚠️ Upozorenje';
    document.getElementById('agroMeteoWarnText').innerHTML = warnings.map(w => w.text).join('<br>');
}

function agroMeteoOverride() {
    agroState.meteoOverride = true;
    document.getElementById('agroMeteoWarn').classList.remove('visible');
    showToast('Meteo override — nastavak na sopstvenu odgovornost', 'info');
}

// ============================================================
// GEOFENCING
// ============================================================
function agroStartGeo() {
    if (!navigator.geolocation) return;

    agroState.geoWatchId = navigator.geolocation.watchPosition(
        pos => {
            const lat = pos.coords.latitude;
            const lng = pos.coords.longitude;
            agroState.geoStart = { lat, lng };
            agroCheckParcelaProximity(lat, lng);
        },
        () => { /* silent — ručni fallback */ },
        { enableHighAccuracy: true, maximumAge: 30000, timeout: 15000 }
    );
}

function agroCheckParcelaProximity(lat, lng) {
    // Ako je parcela već izabrana — ne diraj
    if (agroState.parcelaID) return;
    
    const parcele = (stammdaten.parcele || []).filter(p => p.KooperantID === CONFIG.ENTITY_ID);
    const banner = document.getElementById('agroGeoBanner');
    let bestMatch = null, bestDist = Infinity;

    for (const p of parcele) {
        const pLat = parseFloat(String(p.Lat || '').replace(',', '.'));
        const pLng = parseFloat(String(p.Lng || '').replace(',', '.'));
        if (!pLat || !pLng || isNaN(pLat) || isNaN(pLng)) continue;

        // Point-in-polygon check
        if (p.PolygonGeoJSON) {
            try {
                const geom = JSON.parse(p.PolygonGeoJSON);
                if (agroPointInPolygon(lat, lng, geom)) {
                    bestMatch = p; bestDist = 0; break;
                }
            } catch(e) {}
        }

        // Haversine distance
        const dist = agroHaversine(lat, lng, pLat, pLng);
        if (dist < bestDist) { bestDist = dist; bestMatch = p; }
    }

    if (!bestMatch) return;

    const sel = document.getElementById('agroParcelaSel');
    const ha = parseFloat(String(bestMatch.PovrsinaHa || '0').replace(',', '.')) || 0;

    if (bestDist <= 50) {
        banner.className = 'agro-geo-banner detected';
        banner.innerHTML = '📍 Detektovana parcela: <strong>' + (bestMatch.KatBroj || bestMatch.ParcelaID) + '</strong> ' +
            (bestMatch.Kultura || '') + ' (' + ha.toFixed(2) + ' ha)';
        sel.value = bestMatch.ParcelaID;
        agroState.geoAutoDetect = true;
        onAgroParcelaChange();
    } else if (bestDist <= 200) {
        banner.className = 'agro-geo-banner nearby';
        banner.innerHTML = '📍 Blizu parcele: <strong>' + (bestMatch.KatBroj || bestMatch.ParcelaID) + '</strong> (' +
            Math.round(bestDist) + 'm) — <a href="#" onclick="document.getElementById(\'agroParcelaSel\').value=\'' +
            bestMatch.ParcelaID + '\';onAgroParcelaChange();return false;" style="color:#92400e;font-weight:700;">Izaberi</a>';
    }
}

function agroHaversine(lat1, lng1, lat2, lng2) {
    const R = 6371000;
    const dLat = (lat2 - lat1) * Math.PI / 180;
    const dLng = (lng2 - lng1) * Math.PI / 180;
    const a = Math.sin(dLat/2)**2 + Math.cos(lat1*Math.PI/180) * Math.cos(lat2*Math.PI/180) * Math.sin(dLng/2)**2;
    return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
}

function agroPointInPolygon(lat, lng, geometry) {
    const coords = geometry.coordinates ? geometry.coordinates[0] : geometry[0];
    if (!coords) return false;
    let inside = false;
    for (let i = 0, j = coords.length - 1; i < coords.length; j = i++) {
        const xi = coords[i][1], yi = coords[i][0];
        const xj = coords[j][1], yj = coords[j][0];
        if (((yi > lng) !== (yj > lng)) && (lat < (xj - xi) * (lng - yi) / (yj - yi) + xi)) {
            inside = !inside;
        }
    }
    return inside;
}

// ============================================================
// TIMER
// ============================================================
function agroStartRad() {
    if (!agroState.parcelaID) { showToast('Izaberite parcelu', 'error'); return; }
    if (!agroState.mera) { showToast('Izaberite meru', 'error'); return; }

    // Za Zastita/Prihrana — proveri da li je preparat izabran
    if ((agroState.mera === 'Zastita' || agroState.mera === 'Prihrana') && !agroState.artikalID) {
        showToast('Izaberite preparat', 'error'); return;
    }

    // Snimi količinu
    if (agroState.artikalID) {
        agroState.kolicina = parseFloat(document.getElementById('agroKolicina').value) || 0;
    }

    // Snimi opremu
    agroState.opremaTraktor = document.getElementById('agroTraktor').value || document.getElementById('agroTraktorNovi').value || '';
    agroState.opremaPrskalica = document.getElementById('agroPrskalica').value || document.getElementById('agroPrskalicaNovi').value || '';
    agroState.opremaOstalo = document.getElementById('agroOpremaOstalo').value || '';
    agroState.napomena = document.getElementById('agroNapomena').value || '';

    // GPS start
    if (navigator.geolocation) {
        navigator.geolocation.getCurrentPosition(pos => {
            agroState.geoStart = { lat: pos.coords.latitude, lng: pos.coords.longitude };
        }, () => {}, { enableHighAccuracy: true });
    }

    // Start timer
    agroState.timerStart = new Date();
    agroState.timerInterval = setInterval(agroUpdateTimer, 1000);
    sessionStorage.setItem('agroTimerStart', agroState.timerStart.toISOString());

    // UI
    document.getElementById('agroTimerPanel').style.display = 'block';
    document.getElementById('agroTimerLabel').textContent =
        (agroState.parcelaData ? agroState.parcelaData.KatBroj : agroState.parcelaID) + ' — ' + agroState.mera;
    document.getElementById('agroBtnStart').style.display = 'none';
    document.getElementById('agroBtnStop').style.display = 'block';
    document.getElementById('agroTimerSticky').classList.add('active');

    showToast('Tajmer pokrenut', 'success');
}

function agroUpdateTimer() {
    if (!agroState.timerStart) return;
    const elapsed = Math.floor((new Date() - agroState.timerStart) / 1000);
    const h = String(Math.floor(elapsed / 3600)).padStart(2, '0');
    const m = String(Math.floor((elapsed % 3600) / 60)).padStart(2, '0');
    const s = String(elapsed % 60).padStart(2, '0');
    const display = h + ':' + m + ':' + s;
    document.getElementById('agroTimerDisplay').textContent = display;
    document.getElementById('agroTimerStickyText').textContent = '⏱️ ' + display + ' — ' + agroState.mera;
}

function agroStopRad() {
    clearInterval(agroState.timerInterval);
    const end = new Date();
    const trajanjeMin = Math.round((end - agroState.timerStart) / 60000);
    sessionStorage.removeItem('agroTimerStart');

    // GPS end
    if (navigator.geolocation) {
        navigator.geolocation.getCurrentPosition(pos => {
            agroState.geoEnd = { lat: pos.coords.latitude, lng: pos.coords.longitude };
        }, () => {}, { enableHighAccuracy: true });
    }

    agroState.timerResult = {
        pocetakISO: agroState.timerStart.toISOString(),
        zavrsetakISO: end.toISOString(),
        trajanjeMinuta: trajanjeMin
    };

    // Prikaži potvrdu
    document.getElementById('agroBtnStop').style.display = 'none';
    document.getElementById('agroTimerSticky').classList.remove('active');
    agroShowConfirm();
}

// ============================================================
// POTVRDA
// ============================================================
function agroShowConfirm() {
    document.getElementById('agroStep1').style.display = 'none';
    document.getElementById('agroStep2').style.display = 'block';

    const p = agroState.parcelaData;
    const art = agroState.artikalData;
    const timer = agroState.timerResult;

    const rows = [];
    rows.push(['Parcela', (p ? p.KatBroj : agroState.parcelaID) + ' — ' + (p ? p.Kultura : '')]);
    rows.push(['Površina', p ? (parseFloat(String(p.PovrsinaHa || '0').replace(',', '.')) || 0).toFixed(2) + ' ha' : '?']);
    rows.push(['Mera', agroState.mera]);

    if (art) {
        rows.push(['Preparat', art.naziv]);
        rows.push(['Količina', agroState.kolicina + ' ' + art.jm]);
        if (agroState.karencaDana > 0) {
            const berbeDate = new Date();
            berbeDate.setDate(berbeDate.getDate() + agroState.karencaDana);
            rows.push(['Karenca', agroState.karencaDana + ' dana → berba od ' + berbeDate.toLocaleDateString('sr')]);
        }
    }

    if (agroState.opremaTraktor) rows.push(['Traktor', agroState.opremaTraktor]);
    if (agroState.opremaPrskalica) rows.push(['Prskalica', agroState.opremaPrskalica]);
    if (agroState.opremaOstalo) rows.push(['Ostala oprema', agroState.opremaOstalo]);

    if (timer) {
        const h = Math.floor(timer.trajanjeMinuta / 60);
        const m = timer.trajanjeMinuta % 60;
        rows.push(['Trajanje', (h > 0 ? h + 'h ' : '') + m + ' min']);
    }

    if (agroState.meteoSnapshot) {
        rows.push(['Meteo', agroState.meteoSnapshot.temp.toFixed(1) + '°C, vetar ' + agroState.meteoSnapshot.wind.toFixed(0) + ' km/h, vlažnost ' + agroState.meteoSnapshot.humidity + '%']);
        if (agroState.meteoOverride) rows.push(['Meteo override', '⚠️ Da — nastavljeno uprkos upozorenju']);
    }

    if (agroState.napomena) rows.push(['Napomena', agroState.napomena]);
    if (agroState.geoAutoDetect) rows.push(['GPS', '📍 Auto-detect']);

    document.getElementById('agroConfirmPanel').innerHTML = rows.map(r =>
        '<div class="agro-confirm-row"><span class="label">' + r[0] + '</span><span class="value">' + r[1] + '</span></div>'
    ).join('');
}

function agroBackToStep1() {
    document.getElementById('agroStep1').style.display = 'block';
    document.getElementById('agroStep2').style.display = 'none';
}

// ============================================================
// SAVE TRETMAN
// ============================================================
async function agroSaveTretman() {
    const art = agroState.artikalData;
    const timer = agroState.timerResult;
    const now = new Date();

    let datumBerbeDozvoljeno = '';
    if (agroState.karencaDana > 0) {
        const d = new Date();
        d.setDate(d.getDate() + agroState.karencaDana);
        datumBerbeDozvoljeno = d.toISOString().split('T')[0];
    }

    const record = {
        clientRecordID: crypto.randomUUID(),
        createdAtClient: now.toISOString(),
        kooperantID: CONFIG.ENTITY_ID,
        parcelaID: agroState.parcelaID,
        datum: now.toISOString().split('T')[0],
        mera: agroState.mera,
        artikalID: art ? art.artikalID : '',
        artikalNaziv: art ? art.naziv : '',
        kolicinaUpotrebljena: agroState.kolicina || '',
        jedinicaMere: art ? art.jm : '',
        dozaPreporucena: agroState.dozaPreporucena || '',
        dozaPrimenjena: agroState.kolicina || '',
        opremaTraktor: agroState.opremaTraktor,
        opremaPrskalica: agroState.opremaPrskalica,
        opremaOstalo: agroState.opremaOstalo,
        karencaDana: agroState.karencaDana || '',
        datumBerbeDozvoljeno: datumBerbeDozvoljeno,
        vremePocetka: timer ? timer.pocetakISO : '',
        vremeZavrsetka: timer ? timer.zavrsetakISO : '',
        trajanjeMinuta: timer ? timer.trajanjeMinuta : '',
        geoLatStart: agroState.geoStart ? agroState.geoStart.lat : '',
        geoLngStart: agroState.geoStart ? agroState.geoStart.lng : '',
        geoLatEnd: agroState.geoEnd ? agroState.geoEnd.lat : '',
        geoLngEnd: agroState.geoEnd ? agroState.geoEnd.lng : '',
        geoAutoDetect: agroState.geoAutoDetect ? 'Da' : '',
        meteoTemp: agroState.meteoSnapshot ? agroState.meteoSnapshot.temp : '',
        meteoWind: agroState.meteoSnapshot ? agroState.meteoSnapshot.wind : '',
        meteoHumidity: agroState.meteoSnapshot ? agroState.meteoSnapshot.humidity : '',
        meteoOverride: agroState.meteoOverride ? 'Da' : '',
        napomena: agroState.napomena,
        syncStatus: 'pending'
    };

    // Save to IndexedDB
    await dbPut(db, 'tretmani', record);
    showToast('Tretman sačuvan!', 'success');

    // Sync immediately
    if (navigator.onLine) {
        try {
            const resp = await fetch(CONFIG.API_URL, {
                method: 'POST', headers: { 'Content-Type': 'text/plain' },
                body: JSON.stringify({
                    action: 'syncTretman', token: CONFIG.TOKEN,
                    kooperantID: CONFIG.ENTITY_ID,
                    records: [record]
                })
            });
            const json = await resp.json();
            if (json.success) {
                record.syncStatus = 'synced';
                await dbPut(db, 'tretmani', record);
            }
        } catch(e) {}
    }

    // Reset
    agroResetState();
    agroPopulateParcele();
    agroLoadIstorija();
    document.getElementById('agroStep1').style.display = 'block';
    document.getElementById('agroStep2').style.display = 'none';
}

// ============================================================
// ISTORIJA
// ============================================================
async function agroLoadIstorija() {
    const local = await dbGetAll(db, 'tretmani');
    let server = [];
    try {
        const json = await apiFetch('action=getTretmani&kooperantID=' + encodeURIComponent(CONFIG.ENTITY_ID));
        if (json && json.success && json.records) {
            server = json.records.map(r => ({
                mera: r.Mera || '', datum: fmtDate(r.Datum),
                parcelaID: r.ParcelaID || '', artikalNaziv: r.ArtikalNaziv || '',
                kolicinaUpotrebljena: r.KolicinaUpotrebljena || '',
                jedinicaMere: r.JedinicaMere || '',
                trajanjeMinuta: r.TrajanjeMinuta || '',
                opremaTraktor: r.OpremaTraktor || '',
                karencaDana: r.KarencaDana || '',
                datumBerbeDozvoljeno: r.DatumBerbeDozvoljeno || '',
                syncStatus: 'synced'
            }));
        }
    } catch(e) {}

    const serverIDs = new Set(server.map(r => r.clientRecordID));
    const all = [...server, ...local.filter(r => r.syncStatus === 'pending' && !serverIDs.has(r.clientRecordID))
        .map(r => ({
            mera: r.mera, datum: r.datum, parcelaID: r.parcelaID,
            artikalNaziv: r.artikalNaziv, kolicinaUpotrebljena: r.kolicinaUpotrebljena,
            jedinicaMere: r.jedinicaMere, trajanjeMinuta: r.trajanjeMinuta,
            opremaTraktor: r.opremaTraktor, karencaDana: r.karencaDana,
            datumBerbeDozvoljeno: r.datumBerbeDozvoljeno, syncStatus: 'pending'
        }))
    ].sort((a, b) => (b.datum || '').localeCompare(a.datum || ''));

    const list = document.getElementById('agroTretmaniList');
    const icons = { Zastita: '🛡️', Prihrana: '🌱', Rezidba: '✂️', Zalivanje: '💧', Berba: '🍎' };

    if (all.length === 0) {
        list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema tretmana</p>';
        return;
    }

    list.innerHTML = all.map(r => {
        const bc = r.syncStatus === 'pending' ? 'var(--warning)' : 'var(--success)';
        const min = parseInt(r.trajanjeMinuta) || 0;
        const timeStr = min > 0 ? (Math.floor(min/60) > 0 ? Math.floor(min/60) + 'h ' : '') + (min%60) + 'min' : '';
        return `<div class="queue-item" style="border-left-color:${bc};">
            <div class="qi-header">
                <span class="qi-koop">${icons[r.mera] || ''} ${r.mera}</span>
                <span class="qi-time">${r.datum}</span>
            </div>
            <div class="qi-detail">${r.parcelaID}${r.artikalNaziv ? ' | ' + r.artikalNaziv + ' ' + (r.kolicinaUpotrebljena || '') + ' ' + (r.jedinicaMere || '') : ''}${timeStr ? ' | ⏱️ ' + timeStr : ''}${r.opremaTraktor ? ' | 🚜 ' + r.opremaTraktor : ''}${r.karencaDana ? ' | Karenca ' + r.karencaDana + 'd' : ''}</div>
        </div>`;
    }).join('');
}

// ============================================================
// KOOPERANT: INFO
// ============================================================
async function loadKoopInfo() {
    const config = stammdaten.config || [];
    const gv = k => { const c = config.find(c => c.Parameter === k); return c ? c.Vrednost : '-'; };
    document.getElementById('koopInfoContent').innerHTML = `
        <div style="background:white;border-radius:var(--radius);padding:16px;margin-bottom:12px;box-shadow:0 1px 3px rgba(0,0,0,0.08);">
            <h3 style="color:var(--primary);margin-bottom:12px;">Otkup informacije</h3>
            <table style="width:100%;">
                <tr><td style="padding:8px;color:var(--text-muted);">Status:</td><td style="padding:8px;font-weight:600;">${gv('OtkupAktivan')==='Da'?'🟢 Aktivan':'🔴 Neaktivan'}</td></tr>
                <tr><td style="padding:8px;color:var(--text-muted);">Radno vreme:</td><td style="padding:8px;">${gv('RadnoVremeOd')} - ${gv('RadnoVremeDo')}</td></tr>
                <tr><td style="padding:8px;color:var(--text-muted);">Sezona:</td><td style="padding:8px;">${gv('SezonaOd')} - ${gv('SezonaDo')}</td></tr>
            </table>
        </div>
        <div style="background:white;border-radius:var(--radius);padding:16px;box-shadow:0 1px 3px rgba(0,0,0,0.08);">
            <h3 style="color:var(--primary);margin-bottom:12px;">Aktuelne cene</h3>
            <table style="width:100%;">
                ${config.filter(c => c.Parameter && c.Parameter.startsWith('Cena')).map(c =>
                    '<tr><td style="padding:8px;color:var(--text-muted);">'+c.Parameter.replace('Cena','')+':</td><td style="padding:8px;font-weight:600;">'+c.Vrednost+' RSD/kg</td></tr>').join('')}
            </table>
        </div>`;
}

// ============================================================
// AGROHEMIJA IZDAVANJE — Barcode + Dropdown + Korpa
// ============================================================
let izdKorpa = []; // {artikalID, naziv, jm, cena, kolicina, vrednost}
let izdSelectedKoopID = '';
let izdSelectedKoopName = '';

// --- Populate ---
function populateIzdDropdowns() {
    console.log('stammdaten.artikli:', stammdaten.artikli);
    console.log('artikli length:', (stammdaten.artikli || []).length);
    // Kooperanti
    const kSel = document.getElementById('izdKooperant');
    if (!kSel || kSel.options.length > 1) return;
    kSel.innerHTML = '<option value="">-- Izaberi kooperanta --</option>';
    (stammdaten.kooperanti || []).forEach(k => {
        const o = document.createElement('option');
        o.value = k.KooperantID;
        o.textContent = k.Ime + ' ' + k.Prezime + ' (' + k.KooperantID + ')';
        kSel.appendChild(o);
    });

    // Artikli
    const aSel = document.getElementById('izdArtikal');
    if (!aSel || aSel.options.length > 1) return;
    aSel.innerHTML = '<option value="">-- Ili izaberi ručno --</option>';
    aSel.onchange = function() { izdCalcPreporuka(); };
    (stammdaten.artikli || []).forEach(a => {
        const o = document.createElement('option');
        o.value = a.ArtikalID;
        const cena = parseFloat(a.CenaPoJedinici) || 0;
        o.textContent = a.Naziv + ' (' + (a.JedinicaMere || 'kom') + ') — ' + cena.toLocaleString('sr') + ' RSD';
        aSel.appendChild(o);
    });
}

function onIzdKooperantChange() {
    const koopID = document.getElementById('izdKooperant').value;
    izdSelectedKoopID = koopID;
    const koop = (stammdaten.kooperanti || []).find(k => k.KooperantID === koopID);
    izdSelectedKoopName = koop ? (koop.Ime + ' ' + koop.Prezime) : koopID;

    const pList = document.getElementById('izdParceleList');
    const pGroup = document.getElementById('izdParcelaGroup');
    pList.innerHTML = '';
    document.getElementById('izdUkupnaHa').textContent = '0';
    izdHidePreporuka();

    if (!koopID) { pGroup.style.display = 'none'; return; }

    const parcele = (stammdaten.parcele || []).filter(p => p.KooperantID === koopID);
    if (parcele.length === 0) { pGroup.style.display = 'none'; return; }

    pGroup.style.display = 'block';
    pList.innerHTML = parcele.map((p, i) => {
        const ha = parseFloat(String(p.PovrsinaHa || '0').replace(',', '.')) || 0;
        return `<label class="izd-parcela-chk">
            <input type="checkbox" id="izdPChk${i}" value="${p.ParcelaID}" data-ha="${ha}" onchange="izdOnParceleChange()">
            <div class="parcela-info">${p.KatBroj || p.ParcelaID} — ${p.Kultura || '?'}</div>
            <div class="parcela-ha">${ha.toFixed(2)} ha</div>
        </label>`;
    }).join('');
}

// --- Parcele checkbox change → recalc ---
function izdOnParceleChange() {
    const checked = document.querySelectorAll('#izdParceleList input[type="checkbox"]:checked');
    let totalHa = 0;
    checked.forEach(chk => { totalHa += parseFloat(chk.dataset.ha) || 0; });
    document.getElementById('izdUkupnaHa').textContent = totalHa.toFixed(2);

    // Recalc preporuka ako je artikal izabran
    izdCalcPreporuka();
}

// --- Kalkulacija preporuke ---
function izdCalcPreporuka() {
    const artID = document.getElementById('izdArtikal').value;
    const panel = document.getElementById('izdPreporuka');

    if (!artID) { izdHidePreporuka(); return; }

    const art = (stammdaten.artikli || []).find(a => a.ArtikalID === artID);
    if (!art) { izdHidePreporuka(); return; }

    const dozaPoHa = parseFloat(String(art.DozaPoHa || '0').replace(',', '.')) || 0;
    if (dozaPoHa <= 0) { izdHidePreporuka(); return; }

    // Izračunaj ukupnu površinu iz izabranih parcela
    const checked = document.querySelectorAll('#izdParceleList input[type="checkbox"]:checked');
    let totalHa = 0;
    const parcelaNames = [];
    checked.forEach(chk => {
        totalHa += parseFloat(chk.dataset.ha) || 0;
        const pid = chk.value;
        const p = (stammdaten.parcele || []).find(x => x.ParcelaID === pid);
        parcelaNames.push(p ? (p.KatBroj || pid) : pid);
    });

    if (totalHa <= 0) { izdHidePreporuka(); return; }

    // Osnovna kalkulacija: DozaPoHa × ukupna površina
    const rawQty = dozaPoHa * totalHa;
    const jm = art.JedinicaMere || 'kg';
    const pakovanje = parseFloat(String(art.Pakovanje || '0').replace(',', '.')) || 0;

    let finalQty = rawQty;
    let pakInfo = '';

    // Pakovanje zaokruživanje (ceil na cela pakovanja)
    if (pakovanje > 0) {
        const pakCount = Math.ceil(rawQty / pakovanje);
        finalQty = pakCount ;
        pakInfo = pakCount + ' × ' + pakovanje + ' ' + jm + ' (pakovanje)';
    }

    // Prikaži
    panel.classList.add('visible');
    document.getElementById('izdPreporukaCalc').innerHTML =
        '<strong>' + finalQty.toLocaleString('sr') + ' ' + jm + '</strong>' +
        ' — ' + art.Naziv;
    document.getElementById('izdPreporukaDetail').innerHTML =
        dozaPoHa + ' ' + jm + '/ha × ' + totalHa.toFixed(2) + ' ha = ' +
        rawQty.toLocaleString('sr', {maximumFractionDigits: 2}) + ' ' + jm +
        (pakInfo ? '<br>' + pakInfo : '') +
        '<br>Parcele: ' + parcelaNames.join(', ');

    // Sačuvaj za "Primeni"
    panel._finalQty = finalQty;
}

function izdHidePreporuka() {
    const panel = document.getElementById('izdPreporuka');
    if (panel) {
        panel.classList.remove('visible');
        panel._finalQty = null;
    }
}

function izdPrimeniPreporuku() {
    const panel = document.getElementById('izdPreporuka');
    if (!panel || !panel._finalQty) return;
    document.getElementById('izdKolicina').value = panel._finalQty;
    showToast('Količina: ' + panel._finalQty.toLocaleString('sr'), 'success');
}

// --- QR Scan Kooperant ---
function startIzdKoopScan() {
    const readerDiv = document.getElementById('qr-reader-izd-koop');
    readerDiv.style.display = 'block';
    const scanner = new Html5Qrcode('qr-reader-izd-koop');
    scanner.start(
        { facingMode: 'environment' },
        { fps: 10, qrbox: { width: 250, height: 250 } },
        (text) => {
            scanner.stop().then(() => { readerDiv.style.display = 'none'; });
            let koopID = '';
            try { const d = JSON.parse(text); if (d.id) koopID = d.id; } catch(e) {}
            if (!koopID && text.startsWith('KOOP-')) koopID = text;
            if (!koopID) { showToast('Nije QR kooperanta', 'error'); return; }
            document.getElementById('izdKooperant').value = koopID;
            onIzdKooperantChange();
            showToast('Kooperant: ' + izdSelectedKoopName, 'success');
        },
        () => {}
    ).catch(() => { showToast('Kamera nije dostupna', 'error'); readerDiv.style.display = 'none'; });
}

// --- Barcode Scan Artikal ---
function startIzdBarcodeScan() {
    const readerDiv = document.getElementById('qr-reader-izd-barcode');
    readerDiv.style.display = 'block';
    const scanner = new Html5Qrcode('qr-reader-izd-barcode');
    scanner.start(
        { facingMode: 'environment' },
        { fps: 10, qrbox: { width: 300, height: 150 },
          formatsToSupport: [
              Html5QrcodeSupportedFormats.EAN_13,
              Html5QrcodeSupportedFormats.EAN_8,
              Html5QrcodeSupportedFormats.CODE_128,
              Html5QrcodeSupportedFormats.CODE_39,
              Html5QrcodeSupportedFormats.QR_CODE
          ]
        },
        (decodedText) => {
            // Nemoj zaustavljati scanner — ostaje otvoren za sledeće skeniranje
            onBarcodeScanned(decodedText);
        },
        () => {}
    ).catch(err => { showToast('Kamera nije dostupna', 'error'); readerDiv.style.display = 'none'; });
}

function stopIzdBarcodeScan() {
    const readerDiv = document.getElementById('qr-reader-izd-barcode');
    readerDiv.style.display = 'none';
    // Html5Qrcode nema globalni handle — ali sakrivanje div-a je dovoljan UX
}

let _lastBarcode = '';
let _lastBarcodeTime = 0;

function onBarcodeScanned(code) {
    // Debounce — isti barkod u roku od 2s ignorišemo
    const now = Date.now();
    if (code === _lastBarcode && (now - _lastBarcodeTime) < 2000) return;
    _lastBarcode = code;
    _lastBarcodeTime = now;

    const artikal = (stammdaten.artikli || []).find(a =>
        String(a.BarKod || '').trim() === String(code).trim()
    );

    if (!artikal) {
        showToast('Artikal nije pronađen: ' + code, 'error');
        return;
    }

    // Dodaj u korpu sa količinom 1
    izdDodajUKorpu(artikal.ArtikalID, 1);
    showToast('✅ ' + artikal.Naziv, 'success');
}

// --- Dodaj stavku (iz dropdown-a ili barkoda) ---
function izdDodajStavku() {
    const artID = document.getElementById('izdArtikal').value;
    if (!artID) { showToast('Izaberite artikal', 'error'); return; }
    const kol = parseFloat(document.getElementById('izdKolicina').value) || 1;
    if (kol <= 0) { showToast('Količina mora biti > 0', 'error'); return; }
    izdDodajUKorpu(artID, kol);
    document.getElementById('izdArtikal').value = '';
    document.getElementById('izdKolicina').value = '1';
}

function izdDodajUKorpu(artikalID, kolicina) {
    const art = (stammdaten.artikli || []).find(a => a.ArtikalID === artikalID);
    if (!art) return;

    // Proveri da li već postoji u korpi
    const existing = izdKorpa.find(s => s.artikalID === artikalID);
    if (existing) {
        existing.kolicina += kolicina;
        existing.vrednost = existing.kolicina * existing.cena;
    } else {
        const cena = parseFloat(art.CenaPoJedinici) || 0;
        izdKorpa.push({
            artikalID: artikalID,
            naziv: art.Naziv || artikalID,
            jm: art.JedinicaMere || 'kom',
            cena: cena,
            kolicina: kolicina,
            vrednost: kolicina * cena
        });
    }

    izdRenderKorpa();
}

// --- Render Korpa ---
function izdRenderKorpa() {
    const bd = document.getElementById('izdKorpaBd');
    const countEl = document.getElementById('izdKorpaCount');
    const totalEl = document.getElementById('izdUkupno');

    if (izdKorpa.length === 0) {
        bd.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;font-size:13px;">Skenirajte barkod ili izaberite artikal</p>';
        countEl.textContent = '0 stavki';
        totalEl.textContent = '0 RSD';
        return;
    }

    const ukupno = izdKorpa.reduce((s, r) => s + r.vrednost, 0);
    countEl.textContent = izdKorpa.length + ' stavki';
    totalEl.textContent = ukupno.toLocaleString('sr') + ' RSD';

    bd.innerHTML = izdKorpa.map((s, i) => `
        <div class="izd-row">
            <div class="izd-row-name">${s.naziv}</div>
            <div class="izd-row-qty">
                <input type="number" value="${s.kolicina}" inputmode="decimal"
                    style="width:50px;text-align:center;border:1px solid var(--border);border-radius:4px;padding:4px;font-size:13px;"
                    onchange="izdUpdateQty(${i}, this.value)">
            </div>
            <div class="izd-row-price">${s.cena.toLocaleString('sr')} /${s.jm}</div>
            <div class="izd-row-total">${s.vrednost.toLocaleString('sr')}</div>
            <button class="izd-row-del" onclick="izdRemoveStavka(${i})">✕</button>
        </div>
    `).join('');
}

function izdUpdateQty(index, val) {
    const kol = parseFloat(val) || 0;
    if (kol <= 0) { izdRemoveStavka(index); return; }
    izdKorpa[index].kolicina = kol;
    izdKorpa[index].vrednost = kol * izdKorpa[index].cena;
    izdRenderKorpa();
}

function izdRemoveStavka(index) {
    izdKorpa.splice(index, 1);
    izdRenderKorpa();
}

// --- Završi Izdavanje ---
async function izdZavrsi() {
    if (!izdSelectedKoopID) { showToast('Izaberite kooperanta', 'error'); return; }
    if (izdKorpa.length === 0) { showToast('Korpa je prazna', 'error'); return; }

    const ukupno = izdKorpa.reduce((s, r) => s + r.vrednost, 0);
    const checkedParcele = document.querySelectorAll('#izdParceleList input[type="checkbox"]:checked');
    const parcelaIDs = [];
    checkedParcele.forEach(chk => parcelaIDs.push(chk.value));
    const parcelaID = parcelaIDs.join(',');
    const napomena = document.getElementById('izdNapomena').value || '';

    // Prikaži otpremnicu za potpise
    izdShowOtpremnica({
        kooperantID: izdSelectedKoopID,
        kooperantName: izdSelectedKoopName,
        parcelaID: parcelaID,
        stavke: [...izdKorpa],
        ukupnaVrednost: ukupno,
        napomena: napomena,
        datum: new Date().toISOString().split('T')[0]
    });
}

// --- Otpremnica Modal (isti pattern kao Otkupni List) ---
function izdShowOtpremnica(data) {
    const config = stammdaten.config || [];
    const gv = k => { const c = config.find(c => c.Parameter === k); return c ? c.Vrednost : ''; };
    const koop = (stammdaten.kooperanti || []).find(k => k.KooperantID === data.kooperantID) || {};

    let modal = document.getElementById('izdOtpremnicaModal');
    if (!modal) {
        modal = document.createElement('div');
        modal.id = 'izdOtpremnicaModal';
        modal.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:white;z-index:9999;overflow-y:auto;';
        document.body.appendChild(modal);
    }

    const stavkeHtml = data.stavke.map((s, i) =>
        `<tr>
            <td style="padding:6px;font-size:13px;">${i+1}</td>
            <td style="padding:6px;font-size:13px;">${s.naziv}</td>
            <td style="padding:6px;font-size:13px;text-align:center;">${s.kolicina} ${s.jm}</td>
            <td style="padding:6px;font-size:13px;text-align:right;">${s.cena.toLocaleString('sr')}</td>
            <td style="padding:6px;font-size:13px;text-align:right;font-weight:600;">${s.vrednost.toLocaleString('sr')}</td>
        </tr>`
    ).join('');

    modal.innerHTML = `<div style="padding:16px;max-width:420px;margin:0 auto;font-family:sans-serif;">
        <div style="text-align:center;border-bottom:2px solid #333;padding-bottom:10px;margin-bottom:12px;">
            <div style="font-size:18px;font-weight:700;">${gv('SELLER_NAME')}</div>
            <div style="font-size:12px;color:#666;">${gv('SELLER_STREET')}, ${gv('SELLER_CITY')} ${gv('SELLER_POSTAL_CODE')}</div>
            <div style="font-size:12px;color:#666;">PIB: ${gv('SELLER_PIB')} | MB: ${gv('SELLER_MATICNI_BROJ')}</div>
        </div>
        <h2 style="text-align:center;margin-bottom:14px;font-size:18px;">OTPREMNICA AGROHEMIJE</h2>
        <div style="background:#f5f5f0;padding:10px;border-radius:8px;margin-bottom:12px;font-size:13px;">
            <div><strong>${koop.Ime || ''} ${koop.Prezime || ''}</strong></div>
            <div>${koop.Adresa || ''}, ${koop.Mesto || ''}</div>
            <div>JMBG: ${koop.JMBG || '________'} | BPG: ${koop.BPGBroj || '________'}</div>
            ${data.parcelaID ? '<div>Parcela: ' + data.parcelaID + '</div>' : ''}
        </div>
        <table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:12px;">
            <tr style="background:#f0f0eb;">
                <th style="padding:6px;text-align:left;font-size:11px;color:#666;">#</th>
                <th style="padding:6px;text-align:left;font-size:11px;color:#666;">Artikal</th>
                <th style="padding:6px;text-align:center;font-size:11px;color:#666;">Kol.</th>
                <th style="padding:6px;text-align:right;font-size:11px;color:#666;">Cena</th>
                <th style="padding:6px;text-align:right;font-size:11px;color:#666;">Vrednost</th>
            </tr>
            ${stavkeHtml}
        </table>
        <div style="display:flex;justify-content:space-between;padding:10px;background:#1a5e2a;color:white;border-radius:8px;font-size:16px;font-weight:700;margin-bottom:16px;">
            <span>UKUPNO:</span>
            <span>${data.ukupnaVrednost.toLocaleString('sr')} RSD</span>
        </div>
        <div style="font-size:12px;color:#666;margin-bottom:12px;">Datum: ${data.datum}${data.napomena ? ' | Napomena: ' + data.napomena : ''}</div>

        <div style="margin-top:16px;">
            <div style="margin-bottom:16px;">
                <div style="font-size:12px;color:#666;margin-bottom:4px;">Potpis izdavaoca:</div>
                <canvas id="sigIzdavalac" width="720" height="200" style="border:1px solid #ccc;border-radius:6px;width:100%;height:80px;touch-action:none;"></canvas>
            </div>
            <div style="margin-bottom:16px;">
                <div style="font-size:12px;color:#666;margin-bottom:4px;">Potpis primaoca (kooperant):</div>
                <canvas id="sigPrimalac" width="720" height="200" style="border:1px solid #ccc;border-radius:6px;width:100%;height:80px;touch-action:none;"></canvas>
            </div>
        </div>

        <div style="text-align:center;margin-top:16px;display:flex;gap:8px;">
            <button onclick="clearSignature('sigIzdavalac');clearSignature('sigPrimalac')" style="flex:1;padding:12px;font-size:14px;background:#f5f5f0;color:#666;border:1px solid #ccc;border-radius:8px;">Obriši</button>
            <button onclick="izdConfirmSave()" style="flex:1;padding:12px;font-size:14px;background:var(--primary);color:white;border:none;border-radius:8px;">✅ Potvrdi</button>
            <button onclick="window.print()" style="flex:1;padding:12px;font-size:14px;background:var(--accent);color:white;border:none;border-radius:8px;">🖨️ Štampaj</button>
        </div>
        <button onclick="izdSavePdf()" style="width:100%;padding:12px;margin-top:8px;font-size:14px;background:#2196F3;color:white;border:none;border-radius:8px;">📄 Sačuvaj PDF na Drive</button>
        <button onclick="document.getElementById('izdOtpremnicaModal').style.display='none'" style="width:100%;padding:10px;margin-top:8px;font-size:14px;background:none;color:#666;border:1px solid #ccc;border-radius:8px;">Zatvori</button>
    </div>`;

    modal.style.display = 'block';
    modal._data = data; // sačuvaj za PDF
    setTimeout(() => {
        initSignaturePad('sigIzdavalac');
        initSignaturePad('sigPrimalac');
    }, 100);
}

// --- Confirm + Save to Server ---
async function izdConfirmSave() {
    const sigI = getSignatureData('sigIzdavalac');
    const sigP = getSignatureData('sigPrimalac');
    if (!sigI) { showToast('Potpišite se kao izdavalac!', 'error'); return; }
    if (!sigP) { showToast('Kooperant mora da se potpiše!', 'error'); return; }

    const modal = document.getElementById('izdOtpremnicaModal');
    const data = modal._data;

    showToast('Čuvanje...', 'info');
    try {
        const json = await apiPost('saveIzdavanje', {
            kooperantID: data.kooperantID,
            kooperantName: data.kooperantName,
            parcelaID: data.parcelaID,
            stavke: data.stavke,
            ukupnaVrednost: data.ukupnaVrednost,
            izdaoUser: CONFIG.ENTITY_NAME,
            napomena: data.napomena
        });
        if (json.success) {
            showToast('Izdavanje sačuvano: ' + json.izdavanjeID, 'success');
            izdReset();
            modal.style.display = 'none';
        } else {
            showToast(json.error || 'Greška', 'error');
        }
    } catch(e) {
        showToast('Nema konekcije', 'error');
    }
}

// --- PDF Save (isti pattern kao Otkupni List) ---
async function izdSavePdf() {
    const modal = document.getElementById('izdOtpremnicaModal');
    const data = modal._data;
    if (!data) { showToast('Nema podataka', 'error'); return; }

    const config = stammdaten.config || [];
    const gv = k => { const c = config.find(c => c.Parameter === k); return c ? c.Vrednost : ''; };
    const koop = (stammdaten.kooperanti || []).find(k => k.KooperantID === data.kooperantID) || {};

    const sigI = getSignatureData('sigIzdavalac') || '';
    const sigP = getSignatureData('sigPrimalac') || '';

    showToast('Generisanje PDF-a...', 'info');

    try {
        const jsPDF = (window.jspdf && window.jspdf.jsPDF) || window.jsPDF;
        if (!jsPDF) { showToast('PDF biblioteka nije učitana', 'error'); return; }
        const doc = new jsPDF({ format: 'a5', unit: 'mm' });
        const w = doc.internal.pageSize.getWidth();
        let y = 10;

        // Header
        doc.setFontSize(13); doc.setFont(undefined, 'bold');
        doc.text(gv('SELLER_NAME'), w/2, y, { align: 'center' }); y += 5;
        doc.setFontSize(8); doc.setFont(undefined, 'normal');
        doc.text(gv('SELLER_STREET') + ', ' + gv('SELLER_CITY'), w/2, y, { align: 'center' }); y += 4;
        doc.text('PIB: ' + gv('SELLER_PIB') + ' | MB: ' + gv('SELLER_MATICNI_BROJ'), w/2, y, { align: 'center' }); y += 3;
        doc.setLineWidth(0.5); doc.line(10, y, w-10, y); y += 6;

        // Title
        doc.setFontSize(14); doc.setFont(undefined, 'bold');
        doc.text('OTPREMNICA AGROHEMIJE', w/2, y, { align: 'center' }); y += 7;

        // Kooperant
        doc.setFillColor(240, 240, 234); doc.rect(10, y, w-20, 12, 'F');
        doc.setFontSize(10); doc.setFont(undefined, 'bold');
        doc.text((koop.Ime||'') + ' ' + (koop.Prezime||''), 12, y+4); y += 5;
        doc.setFontSize(8); doc.setFont(undefined, 'normal');
        doc.text((koop.Adresa||'') + ', ' + (koop.Mesto||''), 12, y+4); y += 5;
        doc.text('JMBG: ' + (koop.JMBG||'________') + '  BPG: ' + (koop.BPGBroj||'________'), 12, y+4); y += 8;

        // Table header
        doc.setFontSize(8); doc.setFont(undefined, 'bold');
        doc.setTextColor(100);
        doc.text('#', 12, y+3); doc.text('Artikal', 18, y+3);
        doc.text('Kol.', 75, y+3); doc.text('Cena', 95, y+3);
        doc.text('Vrednost', w-12, y+3, { align: 'right' });
        doc.setLineWidth(0.3); doc.line(10, y+5, w-10, y+5); y += 7;

        // Table rows
        doc.setFont(undefined, 'normal'); doc.setTextColor(0);
        data.stavke.forEach((s, i) => {
            doc.setFontSize(8);
            doc.text(String(i+1), 12, y+3);
            doc.text(s.naziv.substring(0, 30), 18, y+3);
            doc.text(s.kolicina + ' ' + s.jm, 75, y+3);
            doc.text(s.cena.toLocaleString('sr'), 95, y+3);
            doc.text(s.vrednost.toLocaleString('sr'), w-12, y+3, { align: 'right' });
            y += 5;
            if (y > 180) { doc.addPage(); y = 15; }
        });

        // Total
        doc.setLineWidth(0.5); doc.line(10, y, w-10, y); y += 2;
        doc.setFontSize(11); doc.setFont(undefined, 'bold');
        doc.text('UKUPNO:', 12, y+5);
        doc.text(data.ukupnaVrednost.toLocaleString('sr') + ' RSD', w-12, y+5, { align: 'right' });
        y += 10;

        // Info
        doc.setFontSize(8); doc.setFont(undefined, 'normal'); doc.setTextColor(100);
        doc.text('Datum: ' + data.datum + (data.napomena ? '  |  ' + data.napomena : ''), 12, y); y += 6;

        // Signatures
        const sigW = (w - 30) / 2, sigH = 20;
        doc.setFontSize(7); doc.setTextColor(100);
        doc.text('Potpis izdavaoca:', 12, y); doc.text('Potpis primaoca:', 17+sigW, y); y += 2;
        doc.setDrawColor(200);
        doc.rect(12, y, sigW, sigH); doc.rect(17+sigW, y, sigW, sigH);
        if (sigI) try { doc.addImage(sigI, 'PNG', 13, y+1, sigW-2, sigH-2); } catch(e) {}
        if (sigP) try { doc.addImage(sigP, 'PNG', 18+sigW, y+1, sigW-2, sigH-2); } catch(e) {}
        y += sigH + 5;

        doc.setFontSize(6); doc.setTextColor(150);
        doc.text('Generisano: ' + new Date().toISOString().substring(0, 19).replace('T', ' '), w/2, y, { align: 'center' });

        // Upload
        const pdfBase64 = doc.output('datauristring').split(',')[1];
        const fileName = 'Otpremnica_Agro_' + data.kooperantID + '_' + data.datum;

        const json = await apiPost('uploadPdf', {
            fileName: fileName,
            pdfBase64: pdfBase64
        });
        if (json.success) { showToast('PDF sačuvan na Drive!', 'success'); }
        else { showToast('Greška: ' + (json.error || ''), 'error'); }
    } catch(e) {
        showToast('Greška pri generisanju PDF-a', 'error');
    }
}

// --- Reset ---
function izdReset() {
    izdKorpa = [];
    izdSelectedKoopID = '';
    izdSelectedKoopName = '';
    izdRenderKorpa();
    const koopSel = document.getElementById('izdKooperant');
    if (koopSel) koopSel.value = '';
    const pGroup = document.getElementById('izdParcelaGroup');
    if (pGroup) pGroup.style.display = 'none';
    const nap = document.getElementById('izdNapomena');
    if (nap) nap.value = '';
    _lastBarcode = '';
    
    // Reset parcele checkboxes
    document.querySelectorAll('#izdParceleList input[type="checkbox"]').forEach(chk => chk.checked = false);
    const haEl = document.getElementById('izdUkupnaHa');
    if (haEl) haEl.textContent = '0';
    izdHidePreporuka();
}
                 
// ============================================================
// MANAGEMENT: NAVIGATION
// ============================================================
function showMgmtMain(section) {
    qsa('.tab-content').forEach(t => removeClass(t, 'active'));
    qsa('.tab-btn').forEach(b => removeClass(b, 'active'));

    addClass(byId('tab-mgmt'), 'active');
    if (event && event.target) addClass(event.target, 'active');

    const subs = MGMT_SUBS[section];
    const bar = byId('mgmtSubBar');

    setHtml(bar, subs.map((s, i) =>
        `<button class="sub-tab-btn${i === 0 ? ' active' : ''}" onclick="showMgmtSub('${s.id}', this)">${s.label}</button>`
    ).join(''));

    qsa('.mgmt-sub').forEach(s => removeClass(s, 'active'));

    const firstEl = byId('mgmt-' + subs[0].id);
    if (firstEl) addClass(firstEl, 'active');
    if (subs[0].load) subs[0].load();
}

function showMgmtSub(subId, btn) {
    qsa('.mgmt-sub').forEach(s => removeClass(s, 'active'));
    qsa('.sub-tab-btn').forEach(b => removeClass(b, 'active'));

    const el = byId('mgmt-' + subId);
    if (el) addClass(el, 'active');
    if (btn) addClass(btn, 'active');

    const allSubs = Object.values(MGMT_SUBS).flat();
    const sub = allSubs.find(s => s.id === subId);
    if (sub && sub.load) sub.load();
}
    
// ============================================================
// DISPECER STATE
// ============================================================
let dpDem = [];
let dpPlans = [];
let dpSel = null;

// kapacitet po kamionu
let dpKap = JSON.parse(localStorage.getItem('dpKap') || '{}');

// status po kamionu
let dpKS = {};

// master lista kamiona za prikaz
let dpKamioni = [];

// ============================================================
// HELPERS
// ============================================================
function dpToday() {
    const d = new Date();
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
}

function dpSN(stanicaID) {
    if (!stanicaID) return '?';
    const s = (stammdaten.stanice || []).find(x => x.StanicaID === stanicaID);
    return s ? (s.Naziv || s.Mesto || stanicaID) : stanicaID;
}

function dpSaveKap() {
    localStorage.setItem('dpKap', JSON.stringify(dpKap || {}));
}

function dpSetK(vid, kg) {
    dpKap[vid] = parseInt(kg) || 0;
    dpSaveKap();
}

function dpSetS(vid, status, ruta) {
    dpKS[vid] = {
        status: status || 'slobodan',
        ruta: ruta || ''
    };
}

function dpCalcRuta(vid) {
    const plans = (dpPlans || [])
        .filter(p => p.VozacID === vid && (p.Status === 'planned' || p.Status === 'u_toku'));

    if (!plans.length) return '';

    const stanice = [];
    const kupci = [];

    plans.forEach(p => {
        const stanica = p.StanicaName || p.StanicaID || '';
        const kupac = p.KupacName || p.KupacID || '';

        if (stanica && !stanice.includes(stanica)) stanice.push(stanica);
        if (kupac && !kupci.includes(kupac)) kupci.push(kupac);
    });

    return [...stanice, ...kupci].join(' → ');
}



function dpGetSup() {
    const today = dpToday();
    return ((mgmtData && mgmtData.otkupiAll) || []).filter(r =>
        !(r.VozacID || r.VozaciID || '') &&
        fmtDate(r.Datum) === today
    );
}

function dpGetAsg() {
    const today = dpToday();
    return ((mgmtData && mgmtData.otkupiAll) || []).filter(r =>
        !!(r.VozacID || r.VozaciID || '') &&
        fmtDate(r.Datum) === today
    );
}

// ============================================================
// INIT / REFRESH
// ============================================================
async function dpInit() {
    dpKamioni = [];
    dpKS = {};

    // 0) master lista iz stammdaten + kapacitet iz Vozaci taba
    (stammdaten.vozaci || []).forEach(v => {
        const vid = v.VozacID || v.vozacID || v.ID || '';
        if (!vid) return;

        if (!dpKamioni.some(x => x.id === vid)) {
            dpKamioni.push({
                id: vid,
                name: v.Naziv || v.Ime || v.Vozac || vid
            });
        }

        const kap = parseInt(v.KapacitetKG) || 0;
        if (kap > 0 && !dpKap[vid]) {
            dpKap[vid] = kap;
        }
    });

    dpSaveKap();

    // 1) status kamiona sa servera
    try {
        const json = await apiFetch('action=getKamionStatus');
        if (json && json.success) {
            const records = json.records || [];

            records.forEach(r => {
                const vid = r.VozacID || r.vozacID || '';
                if (!vid) return;

                dpKS[vid] = {
                    status: r.Status || r.status || 'slobodan',
                    ruta: r.Ruta || r.ruta || ''
                };

                if (!dpKamioni.some(x => x.id === vid)) {
                    dpKamioni.push({
                        id: vid,
                        name: r.VozacName || r.Naziv || vid
                    });
                }
            });
        }
    } catch (e) {}

    // 2) dodaj i kamione iz assigned otkupa da ništa ne nestane
    dpGetAsg().forEach(r => {
        const vid = r.VozacID || r.VozaciID || '';
        if (!vid) return;
        if (!dpKamioni.some(x => x.id === vid)) {
            dpKamioni.push({ id: vid, name: vid });
        }
    });

    // 3) dodaj i kamione iz localStorage kapaciteta/statusa
    Object.keys(dpKap).forEach(vid => {
        if (!dpKamioni.some(x => x.id === vid)) {
            dpKamioni.push({ id: vid, name: vid });
        }
    });

    Object.keys(dpKS).forEach(vid => {
        if (!dpKamioni.some(x => x.id === vid)) {
            dpKamioni.push({ id: vid, name: vid });
        }
    });

    // 4) demand + plans
    dpDem = [];
    dpPlans = [];
    try {
        const j = await apiFetch('action=getDispecer');
        if (j && j.success) {
            dpDem = j.demand || [];
            dpPlans = j.plans || [];
        }
    } catch (e) {}

    // 5) uskladi rute po planovima
    const vozaciSaPlanovima = [...new Set((dpPlans || []).map(p => p.VozacID).filter(Boolean))];
    vozaciSaPlanovima.forEach(vid => {
        const ruta = dpCalcRuta(vid);
        if (!dpKS[vid]) dpKS[vid] = {};
        dpKS[vid].ruta = ruta;
        if (!dpKS[vid].status || dpKS[vid].status === 'slobodan') {
            dpKS[vid].status = 'utovar';
        }
    });

    dpPD();
    dpRS();
    dpRTr();
    dpRD();
    dpRP();
    dpRK();
    dpX();

    const rt = document.getElementById('dpRT');
    if (rt) {
        rt.textContent = 'Ažurirano: ' + new Date().toLocaleTimeString('sr', {
            hour: '2-digit',
            minute: '2-digit'
        });
    }
}
async function loadDispecer() {
    await dpInit();
}
// ============================================================
// DROPDOWNS
// ============================================================
function dpPD() {
    const k = document.getElementById('dpDK');
    if (k && k.options.length <= 1) {
        (stammdaten.kupci || []).forEach(x => {
            const o = document.createElement('option');
            o.value = x.KupacID;
            o.textContent = x.Naziv || x.KupacID;
            k.appendChild(o);
        });

        if (mgmtData && mgmtData.saldoKupci) {
            mgmtData.saldoKupci.forEach(x => {
                const v = x.KupacID || x.Kupac;
                if (!v || Array.from(k.options).some(o => o.value === v)) return;
                const o = document.createElement('option');
                o.value = v;
                o.textContent = x.Kupac || v;
                k.appendChild(o);
            });
        }
    }

    const vs = document.getElementById('dpDV');
    if (vs && vs.options.length <= 1) {
        const seen = new Set();
        (stammdaten.kulture || []).forEach(x => {
            if (x.VrstaVoca && !seen.has(x.VrstaVoca)) {
                seen.add(x.VrstaVoca);
                const o = document.createElement('option');
                o.value = x.VrstaVoca;
                o.textContent = x.VrstaVoca;
                vs.appendChild(o);
            }
        });
    }
}

// ============================================================
// SUPPLY
// ============================================================
function dpRS() {
    const b = document.getElementById('dpSB');
    if (!b) return;

    const sup = dpGetSup();
    const st = document.getElementById('dpST');
    if (st) {
        const totalKg = sup.reduce((s, r) => s + (parseFloat(r.Kolicina) || 0), 0);
        st.textContent = totalKg.toLocaleString('sr') + ' kg';
    }
    const g = {};

    sup.forEach(r => {
        const s = r.OtkupacID || (r._sheetName || '').replace('OTK-', '') || '?';
        if (!g[s]) g[s] = { kg: 0, n: 0, rows: [] };
        g[s].kg += parseFloat(r.Kolicina) || 0;
        g[s].n++;
        g[s].rows.push(r);
    });

    const ids = Object.keys(g).sort((a, b) => (g[b].kg || 0) - (g[a].kg || 0));

    if (!ids.length) {
        b.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema neraspoređene robe za danas</p>';
        return;
    }

    b.innerHTML = ids.map(sid => {
        const x = g[sid];
        const isSel = dpSel && dpSel.step >= 2 && dpSel.sid === sid;
        return `
            <div class="dp-card sup${isSel ? ' sel' : ''}" onclick="dpTS('${sid}')">
                <div style="display:flex;justify-content:space-between;align-items:center;">
                    <span style="font-weight:700;">📦 ${dpSN(sid)}</span>
                    <span style="font-weight:700;">${x.kg.toLocaleString('sr')} kg</span>
                </div>
                <div style="font-size:12px;color:var(--text-muted);margin-top:4px;">
                    ${x.n} otkupa
                </div>
            </div>
        `;
    }).join('');
}

// ============================================================
// TRANSPORT
// ============================================================
function dpRTr() {
    const b = document.getElementById('dpTB');
    if (!b) return;

    const asg = dpGetAsg();
    const vm = {};

    // 1) svi poznati kamioni
    (dpKamioni || []).forEach(v => {
        const vid = v.id || v.VozacID || v.vozacID || '';
        if (!vid) return;

        vm[vid] = {
            kg: 0,              // stvarno dodeljeno kroz otkupe
            plannedKg: 0,       // planirano kroz dispečer planove
            n: 0,               // broj stvarno dodeljenih otkupa
            st: new Set(),      // stanice iz realnih otkupa i planova
            name: v.name || v.Naziv || vid,
            plans: []           // svi aktivni planovi za kamion
        };
    });

    // 2) stvarno assigned otkupi
    asg.forEach(r => {
        const vid = r.VozacID || r.VozaciID || '';
        if (!vid) return;

        if (!vm[vid]) {
            vm[vid] = {
                kg: 0,
                plannedKg: 0,
                n: 0,
                st: new Set(),
                name: vid,
                plans: []
            };
        }

        vm[vid].kg += parseFloat(r.Kolicina) || 0;
        vm[vid].n++;

        const sid = r.OtkupacID || (r._sheetName || '').replace('OTK-', '');
        if (sid) vm[vid].st.add(sid);
    });

    // 3) aktivni planovi
    (dpPlans || []).forEach(p => {
        const vid = p.VozacID || '';
        if (!vid) return;
        if (p.Status !== 'planned' && p.Status !== 'u_toku') return;

        if (!vm[vid]) {
            vm[vid] = {
                kg: 0,
                plannedKg: 0,
                n: 0,
                st: new Set(),
                name: p.VozacName || vid,
                plans: []
            };
        }

        const pKg = parseFloat(p.PlannedKg) || 0;
        vm[vid].plannedKg += pKg;
        vm[vid].plans.push(p);

        const sid = p.StanicaID || '';
        if (sid) vm[vid].st.add(sid);
    });

    // 4) fallback iz kapaciteta
    Object.keys(dpKap).forEach(vid => {
        if (!vm[vid]) {
            vm[vid] = {
                kg: 0,
                plannedKg: 0,
                n: 0,
                st: new Set(),
                name: vid,
                plans: []
            };
        }
    });

    // 5) fallback iz statusa
    Object.keys(dpKS).forEach(vid => {
        if (!vm[vid]) {
            vm[vid] = {
                kg: 0,
                plannedKg: 0,
                n: 0,
                st: new Set(),
                name: vid,
                plans: []
            };
        }
    });

    // 6) sortiranje — najviše ukupno opterećeni prvi
    const ids = Object.keys(vm).sort((a, b) => {
        const ak = (vm[a].kg || 0) + (vm[a].plannedKg || 0);
        const bk = (vm[b].kg || 0) + (vm[b].plannedKg || 0);
        if (bk !== ak) return bk - ak;
        return (vm[a].name || a).localeCompare(vm[b].name || b);
    });

    const tt = document.getElementById('dpTT');
    if (tt) tt.textContent = ids.length;

    if (!ids.length) {
        b.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema kamiona za prikaz</p>';
        return;
    }

    const sl = {
        slobodan: 'Slobodan',
        utovar: 'Utovar',
        naputu: 'Na putu',
        istovar: 'Istovar'
    };

    b.innerHTML = ids.map(vid => {
        const x = vm[vid];
        const cap = parseInt(dpKap[vid]) || 0;

        const realKg = x.kg || 0;
        const plannedKg = x.plannedKg || 0;
        const loadKg = realKg + plannedKg;

        const pct = cap > 0 ? Math.min(100, Math.round((loadKg / cap) * 100)) : 0;
        const freeKg = cap > 0 ? Math.max(0, cap - loadKg) : 0;

        const bc =
            pct >= 95 ? 'var(--danger)' :
            pct >= 75 ? 'var(--accent)' :
            'var(--success)';

        const isSel = dpSel && dpSel.step >= 1 && dpSel.vid === vid;
        const st = (dpKS[vid] || {}).status || 'slobodan';
        const ruta = (dpKS[vid] || {}).ruta || '';

        // sve stanice kroz planove + assigned
        const stationNames = [...(x.st || [])].map(s => dpSN(s));

        // hladnjače iz aktivnih planova
        const planKupci = [...new Set(
            (x.plans || [])
                .map(p => p.KupacName || p.KupacID || '')
                .filter(Boolean)
        )];

        let rutaText = '';
        if (stationNames.length && planKupci.length) {
            rutaText = stationNames.join(' → ') + ' → ' + planKupci.join(', ');
        } else if (stationNames.length && ruta) {
            rutaText = ruta;
        } else if (ruta) {
            rutaText = ruta;
        }

        const planHtml = (x.plans || []).map(p => {
            const pKg = parseInt(p.PlannedKg || 0) || 0;
            return `
                <div style="font-size:11px;margin-top:3px;color:var(--success);font-weight:600;">
                    📋 Plan: ${p.StanicaName || p.StanicaID || '?'} → ${p.KupacName || p.KupacID || '?'} (${pKg.toLocaleString('sr')} kg)
                </div>
            `;
        }).join('');

        return `
            <div class="dp-card trn${isSel ? ' sel' : ''}" onclick="dpTK('${vid}')">
                <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:2px;">
                    <span style="font-weight:700;font-size:14px;">🚛 ${x.name || vid}</span>
                    <span style="font-size:14px;font-weight:700;">${loadKg.toLocaleString('sr')} kg</span>
                </div>

                <div style="font-size:12px;color:var(--text-muted);margin-top:2px;">
                    Kap: ${
                        cap > 0
                            ? cap.toLocaleString('sr') + ' kg'
                            : `<input type="number"
                                      inputmode="numeric"
                                      placeholder="kg"
                                      style="width:70px;padding:2px 4px;font-size:11px;border:1px solid var(--border);border-radius:4px;"
                                      onclick="event.stopPropagation()"
                                      onchange="dpSetK('${vid}', parseInt(this.value)||0); dpRTr()">`
                    }
                    · Popunjeno: <strong>${pct}%</strong>
                    ${cap > 0 ? ` · Slobodno: ${freeKg.toLocaleString('sr')} kg` : ''}
                </div>

                ${cap > 0 ? `
                    <div class="dp-bar" style="margin-top:6px;">
                        <div class="dp-bf" style="width:${pct}%;background:${bc};"></div>
                    </div>
                ` : ''}

                <div style="font-size:11px;color:var(--text-muted);margin-top:4px;">
                    ${realKg > 0 ? `Realno: ${realKg.toLocaleString('sr')} kg` : 'Realno: 0 kg'}
                    ${plannedKg > 0 ? ` · Planirano: ${plannedKg.toLocaleString('sr')} kg` : ''}
                    · ${x.n} otk.
                </div>

                ${rutaText ? `
                    <div style="font-size:11px;color:var(--text-muted);margin-top:4px;">
                        Ruta: ${rutaText}
                    </div>
                ` : ''}

                ${planHtml}

                <div style="display:flex;justify-content:space-between;align-items:center;margin-top:6px;">
                    <span class="dp-badge ${st}">${sl[st] || st}</span>
                </div>

                <div class="dp-stb" onclick="event.stopPropagation()">
                    ${['slobodan', 'utovar', 'naputu', 'istovar']
                        .map(s => `<button class="${st === s ? 'on' : ''}" onclick="dpCS('${vid}','${s}')">${sl[s]}</button>`)
                        .join('')}
                </div>
            </div>
        `;
    }).join('');
}
// ============================================================
// DEMAND
// ============================================================
function dpRD() {
    const l = document.getElementById('dpDL2');
    const t = document.getElementById('dpDT');
    if (!l || !t) return;

    const tot = dpDem.reduce((s, d) => s + (parseInt(d.Kg) || 0), 0);
    t.textContent = tot.toLocaleString('sr') + ' kg';

    if (!dpDem.length) {
        l.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema demand-a za danas</p>';
        return;
    }

    // saberi aktivne planove po DemandID
    const plannedByDemand = {};
    (dpPlans || []).forEach(p => {
        if (p.Status !== 'planned' && p.Status !== 'u_toku') return;

        const did = p.DemandID || '';
        if (!did) return;

        plannedByDemand[did] = (plannedByDemand[did] || 0) + (parseInt(p.PlannedKg) || 0);
    });

    l.innerHTML = dpDem.map(d => {
        const did = d.DemandID || d.demandID || '';
        const isSel = dpSel && dpSel.step >= 3 && dpSel.did === did;

        const kup = d.KupacName || d.KupacID || '?';
        const trazeno = parseInt(d.Kg) || 0;
        const primljeno = parseInt(d.Primljeno) || 0;
        const planirano = plannedByDemand[did] || 0;

        // koliko još nije ni planirano ni primljeno
        const preostalo = Math.max(0, trazeno - planirano - primljeno);

        // ukupno "pokriveno"
        const pokriveno = Math.min(trazeno, planirano + primljeno);
        const pct = trazeno > 0 ? Math.min(100, Math.round((pokriveno / trazeno) * 100)) : 0;

        const barColor =
            pct >= 100 ? 'var(--success)' :
            pct >= 70 ? 'var(--accent)' :
            '#1565c0';

        return `
            <div class="dp-card dem${isSel ? ' sel' : ''}" onclick="dpTD('${did}')">
                <div style="display:flex;justify-content:space-between;align-items:center;">
                    <strong>🏭 ${kup}</strong>
                    <strong>${trazeno.toLocaleString('sr')} kg</strong>
                </div>

                <div style="font-size:12px;color:var(--text-muted);margin-top:4px;">
                    ${d.Vrsta || ''} ${d.Klasa || ''}
                </div>

                <div class="dp-bar" style="margin-top:8px;">
                    <div class="dp-bf" style="width:${pct}%;background:${barColor};"></div>
                </div>

                <div style="font-size:11px;color:var(--text-muted);margin-top:6px;line-height:1.5;">
                    <div>Planirano: <strong style="color:#1565c0;">${planirano.toLocaleString('sr')} kg</strong></div>
                    <div>Primljeno: <strong style="color:var(--success);">${primljeno.toLocaleString('sr')} kg</strong></div>
                    <div>Preostalo: <strong style="color:${preostalo > 0 ? 'var(--danger)' : 'var(--success)'};">${preostalo.toLocaleString('sr')} kg</strong></div>
                </div>
            </div>
        `;
    }).join('');
}
// ============================================================
// PLANOVI
// ============================================================
function dpRP() {
    const b = document.getElementById('dpPlanList');
    if (!b) return;

    if (!dpPlans.length) {
        b.innerHTML = '';
        return;
    }

    b.innerHTML = dpPlans.map(p => `
        <div class="dp-plan-item">
            <div>
                <div class="dp-pi-route">🚛 ${p.VozacID} · ${p.StanicaName || p.StanicaID || '?'} → ${p.KupacName || p.KupacID || '?'}</div>
                <div style="font-size:11px;color:var(--text-muted);margin-top:2px;">
                    ${parseInt(p.PlannedKg || 0).toLocaleString('sr')} kg · status: ${p.Status || 'planned'}
                </div>
            </div>
            <div style="display:flex;gap:6px;">
                <button onclick="dpChgPlanSt('${p.PlanID}','u_toku')" title="U toku">▶</button>
                <button onclick="dpChgPlanSt('${p.PlanID}','zavrseno')" title="Završeno">✓</button>
                <button onclick="dpRmPlan('${p.PlanID}')" title="Obriši">✕</button>
            </div>
        </div>
    `).join('');
}

// ============================================================
// KPI
// ============================================================
function dpRK() {
    const k1 = document.getElementById('dpK1');
    const k2 = document.getElementById('dpK2');
    const k3 = document.getElementById('dpK3');
    const k4 = document.getElementById('dpK4');
    if (!k1 || !k2 || !k3 || !k4) return;

    k1.textContent = dpGetSup()
        .reduce((s, r) => s + (parseFloat(r.Kolicina) || 0), 0)
        .toLocaleString('sr');

    // ovde sada brojimo master listu kamiona, ne samo assigned
    const vz = new Set();
    (dpKamioni || []).forEach(v => vz.add(v.id || v.VozacID || v.vozacID));
    Object.keys(dpKap).forEach(v => vz.add(v));
    Object.keys(dpKS).forEach(v => vz.add(v));

    k2.textContent = vz.size;
    k3.textContent = dpDem.reduce((s, d) => s + (parseInt(d.Kg) || 0), 0).toLocaleString('sr');
    k4.textContent = dpPlans.filter(p => p.Status === 'planned' || p.Status === 'u_toku').length;
}

// ============================================================
// TAP TO PLAN
// ============================================================
function dpTK(vid) {
    if (dpSel && dpSel.vid === vid) {
        dpX();
        return;
    }
    dpSel = { step: 1, vid };
    dpBN('🚛 ' + vid + ' izabran', 'Korak 2: tapnite STANICU odakle treba pokupiti robu');
    dpHL();
}

function dpTS(sid) {
    if (!dpSel || dpSel.step < 1) {
        showToast('Prvo izaberite kamion', 'info');
        return;
    }
    if (dpSel.sid === sid) {
        dpSel.step = 1;
        dpSel.sid = null;
        dpBN('🚛 ' + dpSel.vid, 'Korak 2: tapnite stanicu');
        dpHL();
        return;
    }

    dpSel.step = 2;
    dpSel.sid = sid;

    const kg = dpGetSup()
        .filter(r => (r.OtkupacID || (r._sheetName || '').replace('OTK-', '')) === sid)
        .reduce((s, r) => s + (parseFloat(r.Kolicina) || 0), 0);

    dpBN(
        '🚛 ' + dpSel.vid + ' → 📦 ' + dpSN(sid) + ' (' + kg.toLocaleString('sr') + ' kg)',
        'Korak 3: tapnite HLADNJAČU gde ide roba'
    );
    dpHL();
}

function dpTD(did) {
    if (!dpSel || dpSel.step < 2) {
        showToast(dpSel && dpSel.step >= 1 ? 'Izaberite stanicu' : 'Prvo kamion, pa stanicu', 'info');
        return;
    }
    if (dpSel.did === did) {
        dpSel.step = 2;
        dpSel.did = null;
        dpBN('🚛 ' + dpSel.vid + ' → 📦 ' + dpSN(dpSel.sid), 'Korak 3: tapnite hladnjaču');
        dpHL();
        return;
    }

    dpSel.step = 3;
    dpSel.did = did;

    const d = dpDem.find(x => x.DemandID === did);
    const kg = dpGetSup()
        .filter(r => (r.OtkupacID || (r._sheetName || '').replace('OTK-', '')) === dpSel.sid)
        .reduce((s, r) => s + (parseFloat(r.Kolicina) || 0), 0);

    dpBN(
        '🚛 ' + dpSel.vid + ' → 📦 ' + dpSN(dpSel.sid) + ' (' + kg.toLocaleString('sr') + ' kg) → 🏭 ' + (d ? d.KupacName : '?'),
        'Tapnite SAČUVAJ PLAN'
    );
    dpHL();
}

function dpBN(t, s) {
    document.getElementById('dpBt').textContent = t;
    document.getElementById('dpBs').textContent = s;
    document.getElementById('dpBnr').classList.add('active');
}

function dpX() {
    dpSel = null;
    document.getElementById('dpBnr').classList.remove('active');
    dpHL();
}

function dpHL() {
    dpRS();
    dpRTr();
    dpRD();
}

// ============================================================
// SAVE PLAN
// ============================================================
async function dpOK() {
    if (!dpSel || dpSel.step < 3) {
        showToast('Završite sva 3 koraka', 'error');
        return;
    }

    const vid = dpSel.vid;
    const sid = dpSel.sid;
    const d = dpDem.find(x => x.DemandID === dpSel.did);
    const kupN = d ? (d.KupacName || d.KupacID) : '?';
    const kupID = d ? (d.KupacID || '') : '';

    const kg = dpGetSup()
        .filter(r => (r.OtkupacID || (r._sheetName || '').replace('OTK-', '')) === sid)
        .reduce((s, r) => s + (parseFloat(r.Kolicina) || 0), 0);

    showToast('Čuvanje plana.', 'info');

    try {
        const json = await apiPost('saveDispecer', {
            demandID: dpSel.did,
            vozacID: vid,
            stanicaID: sid,
            stanicaName: dpSN(sid),
            kupacID: kupID,
            kupacName: kupN,
            plannedKg: Math.round(kg)
        });

        if (json.success) {
            const newPlan = {
                PlanID: json.planID,
                DemandID: dpSel.did,
                VozacID: vid,
                StanicaID: sid,
                StanicaName: dpSN(sid),
                KupacID: kupID,
                KupacName: kupN,
                PlannedKg: Math.round(kg),
                Status: 'planned'
            };

            dpPlans.push(newPlan);

            const ruta = dpCalcRuta(vid);
            dpSetS(vid, 'utovar', ruta);

            fetch(CONFIG.API_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain' },
                body: JSON.stringify({
                    action: 'updateKamionStatus',
                    token: CONFIG.TOKEN,
                    vozacID: vid,
                    status: 'utovar',
                    ruta: ruta
                })
            }).catch(() => {});

            if (!dpKamioni.some(x => x.id === vid)) {
                dpKamioni.push({ id: vid, name: vid });
            }

            showToast('📋 Plan: ' + ruta, 'success');
        } else {
            showToast(json.error || 'Greška', 'error');
        }
    } catch (e) {
        showToast('Nema konekcije', 'error');
    }

    dpX();
    dpRS();
    dpRTr();
    dpRP();
    dpRK();
}

// ============================================================
// PLAN STATUS CHANGE
// ============================================================
async function dpChgPlanSt(planID, newStatus) {
    try {
        await apiPost('updateDispecer', {
            planID,
            status: newStatus
        });

        const p = dpPlans.find(x => x.PlanID === planID);
        const vid = p ? p.VozacID : '';

        if (p) p.Status = newStatus;
        if (newStatus === 'zavrseno') {
            dpPlans = dpPlans.filter(x => x.PlanID !== planID);
        }

        if (vid) {
            const ruta = dpCalcRuta(vid);
            const status = ruta ? (newStatus === 'zavrseno' ? ((dpKS[vid] && dpKS[vid].status) || 'utovar') : newStatus) : 'slobodan';

            dpSetS(vid, status, ruta);

            apiPost('updateKamionStatus', {
                vozacID: vid,
                status: status,
                ruta: ruta
            });
        }

        dpRP();
        dpRK();
        dpRS();
        dpRTr();

        showToast(newStatus === 'zavrseno' ? 'Plan završen' : 'Plan u toku', 'success');
    } catch (e) {}
}

async function dpRmPlan(planID) {
    try {
        const plan = dpPlans.find(x => x.PlanID === planID);
        const vid = plan ? plan.VozacID : '';

        await apiPost('removeDispecer', {
            planID
        });

        dpPlans = dpPlans.filter(x => x.PlanID !== planID);

        if (vid) {
            const ruta = dpCalcRuta(vid);
            const status = ruta ? ((dpKS[vid] && dpKS[vid].status) || 'utovar') : 'slobodan';

            dpSetS(vid, status, ruta);

            fetch(CONFIG.API_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain' },
                body: JSON.stringify({
                    action: 'updateKamionStatus',
                    token: CONFIG.TOKEN,
                    vozacID: vid,
                    status: status,
                    ruta: ruta
                })
            }).catch(() => {});
        }

        dpRP();
        dpRK();
        dpRS();
        dpRTr();

        showToast('Plan obrisan', 'info');
    } catch (e) {}
}
    

// ============================================================
// KAMION STATUS
// ============================================================
async function dpCS(vid, st) {
    dpKS[vid] = {
        status: st,
        ruta: (dpKS[vid] || {}).ruta || ''
    };

    if (!dpKamioni.some(x => x.id === vid)) {
        dpKamioni.push({ id: vid, name: vid });
    }

    dpRTr();

    apiPost('updateKamionStatus', {
        vozacID: vid,
        status: st,
        ruta: (dpKS[vid] || {}).ruta || ''
    });
}

async function dpAD() {
    const kupacID = document.getElementById('dpDK').value;
    const kg = parseInt(document.getElementById('dpDG').value) || 0;
    const vrsta = document.getElementById('dpDV').value || '';
    const klasa = document.getElementById('dpDL').value || '';

    if (!kupacID) {
        showToast('Izaberite kupca', 'error');
        return;
    }
    if (kg <= 0) {
        showToast('Unesite kg', 'error');
        return;
    }

    const kupacName =
        document.getElementById('dpDK').selectedOptions[0]?.textContent || kupacID;

    try {
        const json = await apiPost('saveWarRoomDemand', {
            kupacID,
            kupacName,
            kg,
            vrsta,
            klasa
        });

        if (!json.success) {
            showToast(json.error || 'Greška pri čuvanju', 'error');
            return;
        }

        dpDem.push({
            DemandID: json.demandID,
            KupacID: kupacID,
            KupacName: kupacName,
            Kg: kg,
            Vrsta: vrsta,
            Klasa: klasa,
            Primljeno: 0
        });

        document.getElementById('dpDG').value = '';
        document.getElementById('dpDV').value = '';
        document.getElementById('dpDL').value = '';

        dpRD();
        dpRK();

        showToast('Zahtev dodat', 'success');
    } catch (e) {
        showToast('Nema konekcije', 'error');
    }
}
// ============================================================
// MANAGEMENT: KOOPERANTI
// ============================================================
function populateMgmtStanice() {
    const stanice = stammdaten.stanice || [];
    const fallbackIDs = new Set();
    if (stanice.length === 0) {
        (stammdaten.kooperanti || []).forEach(k => { if (k.StanicaID) fallbackIDs.add(k.StanicaID); });
    }
    ['mgmtStanica', 'mgmtOtkupiStanica'].forEach(selId => {
        const sel = document.getElementById(selId);
        if (!sel) return;
        sel.innerHTML = '<option value="">-- Izaberi stanicu --</option>';
        if (stanice.length > 0) {
            stanice.forEach(s => {
                const o = document.createElement('option');
                o.value = s.StanicaID;
                o.textContent = (s.Naziv || s.Mesto || s.StanicaID) + ' (' + s.StanicaID + ')';
                sel.appendChild(o);
            });
        } else {
            fallbackIDs.forEach(id => {
                const o = document.createElement('option');
                o.value = id; o.textContent = id;
                sel.appendChild(o);
            });
        }
    });
}

function onMgmtStanicaChange() {
    const stanicaID = document.getElementById('mgmtStanica').value;
    const sel = document.getElementById('mgmtKooperant');
    sel.innerHTML = '<option value="">-- Izaberi kooperanta --</option>';
    document.getElementById('mgmtKarticaHeader').style.display = 'none';
    document.getElementById('mgmtKarticaList').innerHTML = '';
    if (!stanicaID) return;
    (stammdaten.kooperanti || []).filter(k => k.StanicaID === stanicaID).forEach(k => {
        const o = document.createElement('option'); o.value = k.KooperantID;
        o.textContent = k.Ime + ' ' + k.Prezime + ' (' + k.KooperantID + ')'; sel.appendChild(o);
    });
}

async function onMgmtKooperantChange() {
    const koopID = document.getElementById('mgmtKooperant').value;
    if (!koopID) { document.getElementById('mgmtKarticaHeader').style.display = 'none'; document.getElementById('mgmtKarticaList').innerHTML = ''; return; }
    const koop = (stammdaten.kooperanti || []).find(k => k.KooperantID === koopID);
    document.getElementById('mgmtKarticaName').textContent = koop ? koop.Ime + ' ' + koop.Prezime : koopID;
    document.getElementById('mgmtKarticaID').textContent = koopID;
    document.getElementById('mgmtKarticaHeader').style.display = 'block';

    let records = [];
    if (mgmtData && mgmtData.kartice) {
        records = mgmtData.kartice.filter(r => r.KooperantID === koopID && r.Opis !== 'UKUPNO');
    } else {
        try { const json = await apiFetch('action=getMgmtKartica&kooperantID=' + encodeURIComponent(koopID)); if (json && json.success && json.records) records = json.records.filter(r => r.Opis !== 'UKUPNO'); } catch (e) {}
    }
    if (records.length === 0) {
        document.getElementById('mgmtKarticaList').innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema podataka</p>';
        ['mgmtKarticaZad','mgmtKarticaRaz','mgmtKarticaSaldo'].forEach(id => document.getElementById(id).textContent = '0');
        return;
    }
    let zad = 0, raz = 0;
    document.getElementById('mgmtKarticaList').innerHTML = records.map(r => {
        const z = parseFloat(r.Zaduzenje)||0, ra = parseFloat(r.Razduzenje)||0, s = parseFloat(r.Saldo)||0;
        zad += z; raz += ra;
        return `<div class="queue-item" style="border-left-color:${z>0?'var(--danger)':'var(--success)'};">
            <div class="qi-header"><span class="qi-koop">${r.BrojDok||''}</span><span class="qi-time">${fmtDate(r.Datum)}</span></div>
            <div class="qi-detail">${r.Opis||''}</div>
            <div class="qi-detail" style="font-size:12px;margin-top:2px;">
                ${z>0?'<span style="color:var(--danger);">Zaduž: '+z.toLocaleString('sr')+'</span> ':''}
                ${ra>0?'<span style="color:var(--success);">Razduž: '+ra.toLocaleString('sr')+'</span> ':''}
                | Saldo: <strong>${s.toLocaleString('sr')}</strong></div></div>`;
    }).join('');
    document.getElementById('mgmtKarticaZad').textContent = zad.toLocaleString('sr');
    document.getElementById('mgmtKarticaRaz').textContent = raz.toLocaleString('sr');
    document.getElementById('mgmtKarticaSaldo').textContent = (zad - raz).toLocaleString('sr');
}

function loadMgmtKoopSaldo() {
    const kartice = (mgmtData && mgmtData.kartice) ? mgmtData.kartice : [];
    const list = document.getElementById('mgmtKoopSaldoList');
    const totals = kartice.filter(r => r.Opis === 'UKUPNO');
    if (totals.length === 0) { list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema podataka</p>'; return; }
    list.innerHTML = totals.map(r => {
        const koop = (stammdaten.kooperanti || []).find(k => k.KooperantID === r.KooperantID);
        const name = koop ? koop.Ime + ' ' + koop.Prezime : r.KooperantID;
        const saldo = parseFloat(r.Saldo)||0;
        const zad = parseFloat(r.Zaduzenje)||0, raz = parseFloat(r.Razduzenje)||0;
        const bc = saldo > 0 ? 'var(--warning)' : 'var(--success)';
        return `<div class="queue-item" style="border-left-color:${bc};">
            <div class="qi-header"><span class="qi-koop">${name}</span><span class="qi-time">Saldo: ${saldo.toLocaleString('sr')} RSD</span></div>
            <div class="qi-detail">Zaduž: ${zad.toLocaleString('sr')} | Razduž: ${raz.toLocaleString('sr')}</div></div>`;
    }).join('');
}

function loadMgmtKoopPregled() {
    // Popuni dropdown stanica
    const sel = document.getElementById('mgmtPregledStanica');
    if (sel.options.length <= 1) {
        const stanice = new Set();
        const data = (mgmtData && mgmtData.saldoOMDetail) ? mgmtData.saldoOMDetail : [];
        data.forEach(r => { if (r.StanicaID) stanice.add(r.StanicaID); });
        stanice.forEach(s => { const o = document.createElement('option'); o.value = s; const st = (stammdaten.stanice||[]).find(x => x.StanicaID === s); o.textContent = st ? (st.Naziv||st.Mesto||s)+' ('+s+')' : s; sel.appendChild(o); });
    }
    renderMgmtKoopPregled();
}

function renderMgmtKoopPregled() {
    const stanicaFilter = document.getElementById('mgmtPregledStanica').value;
    let records = (mgmtData && mgmtData.saldoOMDetail) ? mgmtData.saldoOMDetail : [];
    if (stanicaFilter) records = records.filter(r => r.StanicaID === stanicaFilter);
    
    const list = document.getElementById('mgmtKoopPregledList');
    if (records.length === 0) { list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema podataka</p>'; return; }
    
    // Sortiraj po stanici pa po imenu
    records.sort((a, b) => (a.StanicaID||'').localeCompare(b.StanicaID||'') || (a.Kooperant||'').localeCompare(b.Kooperant||''));
    
    // Totali
    let totKg = 0, totVr = 0, totIsp = 0, totAgro = 0, totSaldo = 0, totAmb = 0;
    records.forEach(r => {
        totKg += parseFloat(r.Kolicina)||0; totVr += parseFloat(r.Vrednost)||0;
        totIsp += parseFloat(r.Isplaceno)||0; totAgro += parseFloat(r.AgroZaduzenje)||0;
        totSaldo += parseFloat(r.Saldo)||0; totAmb += parseFloat(r.Ambalaza)||0;
    });
    
    list.innerHTML = `
        <div class="stats-grid" style="margin-bottom:12px;">
            <div class="stat-card"><div class="stat-value" style="font-size:20px;">${totKg.toLocaleString('sr')}</div><div class="stat-label">Ukupno kg</div></div>
            <div class="stat-card"><div class="stat-value" style="font-size:20px;">${totVr.toLocaleString('sr')}</div><div class="stat-label">Vrednost</div></div>
            <div class="stat-card"><div class="stat-value" style="font-size:20px;">${totIsp.toLocaleString('sr')}</div><div class="stat-label">Isplaćeno</div></div>
            <div class="stat-card"><div class="stat-value" style="font-size:20px;">${totSaldo.toLocaleString('sr')}</div><div class="stat-label">Saldo</div></div>
        </div>
        ${records.map(r => {
            const kg = parseFloat(r.Kolicina)||0;
            const vr = parseFloat(r.Vrednost)||0;
            const isp = parseFloat(r.Isplaceno)||0;
            const agro = parseFloat(r.AgroZaduzenje)||0;
            const saldo = parseFloat(r.Saldo)||0;
            const amb = parseFloat(r.Ambalaza)||0;
            const bc = saldo > 0 ? 'var(--warning)' : 'var(--success)';
            return `<div class="queue-item" style="border-left-color:${bc};">
                <div class="qi-header"><span class="qi-koop">${r.Kooperant||r.KooperantID}</span><span class="qi-time">${fmtStanica(r.StanicaID)}</span></div>
                <div class="qi-detail">
                    ${kg.toLocaleString('sr')} kg | Vrednost: ${vr.toLocaleString('sr')}
                </div>
                <div class="qi-detail" style="font-size:12px;margin-top:2px;">
                    Isplaćeno: ${isp.toLocaleString('sr')} | Agro: ${agro.toLocaleString('sr')} | Amb: ${amb.toLocaleString('sr')}
                </div>
                <div class="qi-detail" style="font-size:13px;margin-top:4px;font-weight:600;">
                    Saldo: <span style="color:${saldo>0?'var(--danger)':'var(--success)'};">${saldo.toLocaleString('sr')} RSD</span>
                </div>
            </div>`;
        }).join('')}`;
}
    
// ============================================================
// MANAGEMENT: STANICE
// ============================================================
async function loadMgmtOtkupi() {
    const stanicaID = document.getElementById('mgmtOtkupiStanica').value;
    if (!stanicaID) { document.getElementById('mgmtOtkupiList').innerHTML = ''; return; }
    const od = document.getElementById('mgmtOtkupiOd').value, doo = document.getElementById('mgmtOtkupiDo').value;
    
    let records = [];
    if (mgmtData && mgmtData.otkupiAll) {
        records = mgmtData.otkupiAll
            .filter(r => (r._sheetName === 'OTK-' + stanicaID) || (r.OtkupacID === stanicaID))
            .map(r => ({
                datum: fmtDate(r.Datum), kooperantName: r.KooperantName || r.KooperantID || '',
                vrstaVoca: r.VrstaVoca || '', klasa: r.Klasa || 'I',
                kolicina: parseFloat(r.Kolicina) || 0, cena: parseFloat(r.Cena) || 0
            }));
    } else {
        try {
            const json = await apiFetch('action=getMgmtOtkupiByStanica&stanicaID=' + encodeURIComponent(stanicaID));
            if (json && json.success && json.records) records = json.records.map(r => ({
                datum: fmtDate(r.Datum), kooperantName: r.KooperantName || r.KooperantID || '',
                vrstaVoca: r.VrstaVoca || '', klasa: r.Klasa || 'I',
                kolicina: parseFloat(r.Kolicina) || 0, cena: parseFloat(r.Cena) || 0
            }));
        } catch (e) {}
    }
    
    if (od) records = records.filter(r => r.datum >= od);
    if (doo) records = records.filter(r => r.datum <= doo);
    records.sort((a, b) => b.datum.localeCompare(a.datum));
    const kg = records.reduce((s, r) => s + (r.kolicina || 0), 0);
    const vr = records.reduce((s, r) => s + (r.kolicina || 0) * (r.cena || 0), 0);
    document.getElementById('mgmtOtkupiCount').textContent = records.length;
    document.getElementById('mgmtOtkupiKg').textContent = kg.toLocaleString('sr');
    document.getElementById('mgmtOtkupiVrednost').textContent = vr.toLocaleString('sr');
    document.getElementById('mgmtOtkupiKoop').textContent = new Set(records.map(r => r.kooperantName)).size;
    const list = document.getElementById('mgmtOtkupiList');
    if (records.length === 0) { list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema otkupa</p>'; return; }
    list.innerHTML = records.map(r => {
        const v = ((r.kolicina || 0) * (r.cena || 0)).toLocaleString('sr');
        return `<div class="queue-item"><div class="qi-header"><span class="qi-koop">${r.kooperantName}</span><span class="qi-time">${r.datum}</span></div>
            <div class="qi-detail">${r.vrstaVoca} ${r.klasa} | ${r.kolicina} kg × ${r.cena} = <strong>${v} RSD</strong></div></div>`;
    }).join('');
}

function loadMgmtSaldoOM() {
    const records = (mgmtData && mgmtData.saldoOM) ? mgmtData.saldoOM : [];
    const list = document.getElementById('mgmtSaldoOMList');
    if (records.length === 0) { list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema podataka</p>'; return; }
    list.innerHTML = records.map(r => {
        const saldo = parseFloat(r.Saldo)||0;
        const bc = saldo > 0 ? 'var(--warning)' : 'var(--success)';
        return `<div class="queue-item" style="border-left-color:${bc};">
            <div class="qi-header"><span class="qi-koop">${fmtStanica(r.StanicaID || r.Stanica)}</span><span class="qi-time">Saldo: ${saldo.toLocaleString('sr')} RSD</span></div>
            <div class="qi-detail">Avans: ${(parseFloat(r.Avans)||0).toLocaleString('sr')} | Isplaceno: ${(parseFloat(r.Isplaceno)||0).toLocaleString('sr')}</div></div>`;
    }).join('');
}

function loadMgmtOtkupPoOM() {
    const records = (mgmtData && mgmtData.otkupPoOM) ? mgmtData.otkupPoOM : [];
    const list = document.getElementById('mgmtOtkupPoOMList');
    if (records.length === 0) { list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema podataka</p>'; return; }
    const grouped = {};
    records.forEach(r => {
        const s = r.StanicaID||'?';
        if (!grouped[s]) grouped[s] = { items: [], totalKg: 0, totalAmb: 0, totalVr: 0 };
        const kg = parseFloat(r.Kolicina)||0, amb = parseFloat(r.Ambalaza)||0, vr = parseFloat(r.Vrednost)||0;
        grouped[s].items.push(r); grouped[s].totalKg += kg; grouped[s].totalAmb += amb; grouped[s].totalVr += vr;
    });
    list.innerHTML = Object.entries(grouped).map(([stanica, g]) =>
        `<div style="background:white;border-radius:var(--radius);padding:14px;margin-bottom:12px;box-shadow:0 1px 3px rgba(0,0,0,0.08);border-left:4px solid var(--primary);">
            <div style="display:flex;justify-content:space-between;margin-bottom:8px;"><strong style="color:var(--primary);font-size:16px;">${fmtStanica(stanica)}</strong><span style="font-size:13px;font-weight:600;">${g.totalKg.toLocaleString('sr')} kg</span></div>
            <div style="font-size:12px;color:var(--text-muted);margin-bottom:8px;">Amb: ${g.totalAmb.toLocaleString('sr')} | Vrednost: ${g.totalVr.toLocaleString('sr')} RSD</div>
            ${g.items.map(r => { const kg=parseFloat(r.Kolicina)||0,amb=parseFloat(r.Ambalaza)||0,vr=parseFloat(r.Vrednost)||0; return `<div style="padding:4px 0;font-size:12px;border-top:1px solid #eee;display:flex;justify-content:space-between;"><span>${r.VrstaVoca} ${r.Klasa}</span><span>${kg.toLocaleString('sr')} kg | ${amb.toLocaleString('sr')} amb | ${vr.toLocaleString('sr')} RSD | ${r.BrojOtkupa||0} otk.</span></div>`; }).join('')}
        </div>`).join('');
}

// ============================================================
// MANAGEMENT: KUPCI
// ============================================================
async function loadMgmtFakture() {
    const kupacID = document.getElementById('mgmtFaktureKupac').value;
    if (!kupacID) { document.getElementById('mgmtFaktureList').innerHTML = ''; return; }
    
    let records = [];
    if (mgmtData && mgmtData.fakture) {
        records = mgmtData.fakture.filter(r => 
            String(r.KupacID) === kupacID || String(r.Kupac) === kupacID
        );
    } else {
        document.getElementById('mgmtFaktureList').innerHTML = '<p style="text-align:center;padding:20px;color:var(--text-muted);">Učitavanje...</p>';
        try {
            const json = await apiFetch('action=getMgmtFakture&kupacID=' + encodeURIComponent(kupacID));
            if (json && json.success && json.records) records = json.records;
        } catch (e) {}
    }
    
    const list = document.getElementById('mgmtFaktureList');
    if (records.length === 0) { list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema faktura</p>'; return; }
    list.innerHTML = records.map(r => {
        const iznos = parseFloat(r.Iznos) || 0;
        const placeno = parseFloat(r.Placeno) || 0;
        const saldo = parseFloat(r.Saldo) || 0;
        const bc = saldo <= 0 ? 'var(--success)' : 'var(--danger)';
        return `<div class="queue-item" style="border-left-color:${bc};cursor:pointer;" onclick="toggleFakturaStavke('${r.FakturaID}', this)">
            <div class="qi-header"><span class="qi-koop">${r.BrojFakture || r.FakturaID}</span><span class="qi-time">${fmtDate(r.Datum)}</span></div>
            <div class="qi-detail">Iznos: <strong>${iznos.toLocaleString('sr')}</strong> | Plaćeno: ${placeno.toLocaleString('sr')} | Saldo: <strong>${saldo.toLocaleString('sr')}</strong></div>
            <div class="qi-detail" style="font-size:11px;margin-top:2px;">${r.Status || ''}${r.SEFStatus ? ' | SEF: ' + r.SEFStatus : ''}</div>
            <div class="faktura-stavke" id="stavke-${r.FakturaID}" style="display:none;margin-top:8px;padding-top:8px;border-top:1px solid #eee;"></div>
        </div>`;
    }).join('');
}

async function toggleFakturaStavke(fakturaID, parentEl) {
    const div = document.getElementById('stavke-' + fakturaID);
    if (!div) return;
    if (div.style.display === 'block') { div.style.display = 'none'; return; }
    div.style.display = 'block';
    
    let stavke = [];
    if (mgmtData && mgmtData.fakturaStavke) {
        stavke = mgmtData.fakturaStavke.filter(r => String(r.FakturaID) === fakturaID);
    } else {
        div.innerHTML = '<span style="font-size:12px;color:var(--text-muted);">Učitavanje...</span>';
        try {
            const json = await apiFetch('action=getMgmtFakturaStavke&fakturaID=' + encodeURIComponent(fakturaID));
            if (json && json.success && json.records) stavke = json.records;
        } catch (e) {}
    }
    
    if (stavke.length === 0) { div.innerHTML = '<span style="font-size:12px;color:var(--text-muted);">Nema stavki</span>'; return; }
    div.innerHTML = `<table style="width:100%;font-size:11px;border-collapse:collapse;">
        <tr style="color:var(--text-muted);"><td>Prijemnica</td><td>Zbirna</td><td>Klasa</td><td style="text-align:right;">Kg</td><td style="text-align:right;">Cena</td><td style="text-align:right;">Iznos</td></tr>
        ${stavke.map(s => `<tr style="border-top:1px solid #f0f0f0;">
            <td>${s.BrojPrijemnice || s.PrijemnicaID || ''}</td>
            <td>${s.BrojZbirne || ''}</td>
            <td>${s.Klasa || ''}</td>
            <td style="text-align:right;">${(parseFloat(s.Kolicina) || 0).toLocaleString('sr')}</td>
            <td style="text-align:right;">${(parseFloat(s.Cena) || 0).toLocaleString('sr')}</td>
            <td style="text-align:right;font-weight:600;">${(parseFloat(s.Iznos) || 0).toLocaleString('sr')}</td>
        </tr>`).join('')}
    </table>`;
}

function loadMgmtKupci() {
    const records = (mgmtData && mgmtData.saldoKupci) ? mgmtData.saldoKupci : [];
    const list = document.getElementById('mgmtKupciList');
    if (records.length === 0) { list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema podataka</p>'; return; }
    list.innerHTML = records.map(r => {
        const saldo = parseFloat(r.Saldo)||0;
        const bc = saldo > 0 ? 'var(--danger)' : 'var(--success)';
        return `<div class="queue-item" style="border-left-color:${bc};">
            <div class="qi-header"><span class="qi-koop">${r.Kupac||r.KupacID||''}</span><span class="qi-time">Saldo: ${saldo.toLocaleString('sr')} RSD</span></div>
            <div class="qi-detail">Fakturisano: ${(parseFloat(r.Fakturisano)||0).toLocaleString('sr')} | Placeno: ${(parseFloat(r.Placeno)||0).toLocaleString('sr')}</div></div>`;
    }).join('');
}

function loadMgmtPredato() {
    const records = (mgmtData && mgmtData.predatoPoKupcu) ? mgmtData.predatoPoKupcu : [];
    const list = document.getElementById('mgmtPredatoList');
    if (records.length === 0) { list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema podataka</p>'; return; }
    const grouped = {};
    records.forEach(r => {
        const k = r.KupacID||'?';
        if (!grouped[k]) grouped[k] = { items: [], totalKg: 0, totalAmb: 0, totalVr: 0 };
        const kg = parseFloat(r.Kolicina)||0, amb = parseFloat(r.Ambalaza)||0, vr = parseFloat(r.Vrednost)||0;
        grouped[k].items.push(r); grouped[k].totalKg += kg; grouped[k].totalAmb += amb; grouped[k].totalVr += vr;
    });
    list.innerHTML = Object.entries(grouped).map(([kupac, g]) =>
        `<div style="background:white;border-radius:var(--radius);padding:14px;margin-bottom:12px;box-shadow:0 1px 3px rgba(0,0,0,0.08);border-left:4px solid var(--accent);">
            <div style="display:flex;justify-content:space-between;margin-bottom:8px;"><strong style="color:var(--primary);font-size:16px;">${kupac}</strong><span style="font-size:13px;font-weight:600;">${g.totalKg.toLocaleString('sr')} kg</span></div>
            <div style="font-size:12px;color:var(--text-muted);margin-bottom:8px;">Amb: ${g.totalAmb.toLocaleString('sr')} | Vrednost: ${g.totalVr.toLocaleString('sr')} RSD</div>
            ${g.items.map(r => { const kg=parseFloat(r.Kolicina)||0,amb=parseFloat(r.Ambalaza)||0,vr=parseFloat(r.Vrednost)||0; return `<div style="padding:4px 0;font-size:12px;border-top:1px solid #eee;display:flex;justify-content:space-between;"><span>${r.VrstaVoca} ${r.Klasa}</span><span>${kg.toLocaleString('sr')} kg | ${amb.toLocaleString('sr')} amb | ${vr.toLocaleString('sr')} RSD | ${r.BrojPrijemnica||0} prij.</span></div>`; }).join('')}
        </div>`).join('');
}

// ============================================================
// MANAGEMENT: AGROHEMIJA (placeholder)
// ============================================================

function loadMgmtAgroStanje() {
    document.getElementById('mgmtAgroStanjeList').innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Pokrenite ExportMgmtReports iz Excela</p>';
}

// ============================================================
// TAB NAVIGATION (non-management)
// ============================================================
function showTab(tabName) {
    qsa('.tab-content').forEach(t => removeClass(t, 'active'));
    qsa('.tab-btn').forEach(b => removeClass(b, 'active'));

    const tabEl = byId('tab-' + tabName);
    if (tabEl) addClass(tabEl, 'active');
    if (event && event.target) addClass(event.target, 'active');

    if (tabName === 'queue') { renderQueueList(); updateStats(); }
    if (tabName === 'pregled') loadOtkupPregled();
    if (tabName === 'otpremnice') loadOtpremaOverview();
    if (tabName === 'kartica') loadKartica();
    if (tabName === 'parcele') loadParcele();
    if (tabName === 'agromere') loadAgronom();
    if (tabName === 'koopinfo') loadKoopInfo();
    if (tabName === 'zbirna') loadVozacData();
    if (tabName === 'transport') loadVozacTransport();
    if (tabName === 'dispecer') loadDispecer();
}

// ============================================================
// UI UPDATES
// ============================================================
async function updateSyncBadge(status) {
    const badge = byId('syncBadge');
    if (!badge) return;
    if (status === 'syncing') {
        setText(badge, 'SYNC...');
        badge.className = 'sync-badge sync-pending';
        return;
    }
    const pending = await dbGetByIndex(db, CONFIG.STORE_NAME, 'syncStatus', 'pending');
    if (!navigator.onLine) {
        setText(badge, 'OFFLINE' + (pending.length > 0 ? ' (' + pending.length + ')' : ''));
        badge.className = 'sync-badge sync-offline';
    } else if (pending.length > 0) {
        setText(badge, 'ČEKA: ' + pending.length);
        badge.className = 'sync-badge sync-pending';
    } else {
        setText(badge, 'ONLINE');
        badge.className = 'sync-badge sync-online';
    }
}

async function updateStats() {
    const all = await dbGetAll(db, CONFIG.STORE_NAME);
    const today = new Date().toISOString().split('T')[0];
    const t = all.filter(r => r.datum === today);
    document.getElementById('statPending').textContent = t.filter(r => r.syncStatus === 'pending').length;
    document.getElementById('statSynced').textContent = t.filter(r => r.syncStatus === 'synced').length;
}

async function renderQueueList() {
    const pending = await dbGetByIndex(db, CONFIG.STORE_NAME, 'syncStatus', 'pending');
    const list = byId('queueList');
    if (!list) return;

    if (pending.length === 0) {
        setHtml(list, '<p style="text-align:center;color:var(--text-muted);padding:40px;">Nema stavki za sinhronizaciju</p>');
        return;
    }

    setHtml(list, pending.map(r =>
        `<div class="queue-item"><div class="qi-header"><span class="qi-koop">${r.kooperantName}</span><span class="qi-time">${new Date(r.createdAtClient).toLocaleTimeString('sr')}</span></div>
            <div class="qi-detail">${r.vrstaVoca} ${r.klasa} | ${r.kolicina} kg × ${r.cena} RSD</div></div>`
    ).join(''));
}

// ============================================================
// HELPERS
// ============================================================
function showToast(msg, type = 'info') {
    const toast = byId('toast');
    setText(toast, msg);
    toast.className = 'toast show ' + type;
    setTimeout(() => { toast.className = 'toast'; }, 3000);
}

// ============================================================
// SERVICE WORKER
// ============================================================
if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('./sw.js').then(reg => {
        setInterval(() => reg.update(), 60000);
        reg.addEventListener('updatefound', () => {
            const nw = reg.installing;
            nw.addEventListener('statechange', () => { if (nw.state === 'activated') showToast('Nova verzija učitana', 'info'); });
        });
    }).catch(err => console.log('SW registration failed:', err));
}
