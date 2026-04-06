// ============================================================
// AGROHEMIJA IZDAVANJE — Barcode + Dropdown + Korpa
// ============================================================
let izdKorpa = []; // {artikalID, naziv, jm, cena, kolicina, vrednost}
let izdSelectedKoopID = '';
let izdSelectedKoopName = '';
let izdPreporukaQty = null;


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
            <input type="checkbox" id="izdPChk${i}" value="${escapeHtml(p.ParcelaID)}" data-ha="${ha}" onchange="izdOnParceleChange()">
            <div class="parcela-info">${escapeHtml(p.KatBroj || p.ParcelaID)} — ${escapeHtml(p.Kultura || '?')}</div>
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

    const rawQty = dozaPoHa * totalHa;
    const jm = art.JedinicaMere || 'kg';
    const pakovanje = parseFloat(String(art.Pakovanje || '0').replace(',', '.')) || 0;

    let finalQty = rawQty;
    let pakInfo = '';

    if (pakovanje > 0) {
        const pakCount = Math.ceil(rawQty / pakovanje);
        finalQty = pakCount;
        pakInfo = pakCount + ' × ' + pakovanje + ' ' + jm + ' (pakovanje)';
    }

    panel.classList.add('visible');
    document.getElementById('izdPreporukaCalc').innerHTML =
        '<strong>' + escapeHtml(finalQty.toLocaleString('sr')) + ' ' + escapeHtml(jm) + '</strong>' +
        ' — ' + escapeHtml(art.Naziv);

    document.getElementById('izdPreporukaDetail').innerHTML =
        escapeHtml(String(dozaPoHa)) + ' ' + escapeHtml(jm) + '/ha × ' + escapeHtml(totalHa.toFixed(2)) + ' ha = ' +
        escapeHtml(rawQty.toLocaleString('sr', { maximumFractionDigits: 2 })) + ' ' + escapeHtml(jm) +
        (pakInfo ? '<br>' + escapeHtml(pakInfo) : '') +
        '<br>Parcele: ' + escapeHtml(parcelaNames.join(', '));

    // Module scope umesto DOM property
    izdPreporukaQty = finalQty;
}

function izdHidePreporuka() {
    const panel = document.getElementById('izdPreporuka');
    if (panel) panel.classList.remove('visible');
    izdPreporukaQty = null;
}

function izdPrimeniPreporuku() {
    if (izdPreporukaQty === null) return;
    document.getElementById('izdKolicina').value = izdPreporukaQty;
    showToast('Količina: ' + izdPreporukaQty.toLocaleString('sr'), 'success');
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
            <div class="izd-row-name">${escapeHtml(s.naziv)}</div>
            <div class="izd-row-qty">
                <input type="number" value="${s.kolicina}" inputmode="decimal"
                    style="width:50px;text-align:center;border:1px solid var(--border);border-radius:4px;padding:4px;font-size:13px;"
                    onchange="izdUpdateQty(${i}, this.value)">
            </div>
            <div class="izd-row-price">${s.cena.toLocaleString('sr')} /${escapeHtml(s.jm)}</div>
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
            <td style="padding:6px;font-size:13px;">${escapeHtml(s.naziv)}</td>
            <td style="padding:6px;font-size:13px;text-align:center;">${escapeHtml(s.kolicina)} ${escapeHtml(s.jm)}</td>
            <td style="padding:6px;font-size:13px;text-align:right;">${s.cena.toLocaleString('sr')}</td>
            <td style="padding:6px;font-size:13px;text-align:right;font-weight:600;">${s.vrednost.toLocaleString('sr')}</td>
        </tr>`
    ).join('');

    modal.innerHTML = `<div style="padding:16px;max-width:420px;margin:0 auto;font-family:sans-serif;">
        <div style="text-align:center;border-bottom:2px solid #333;padding-bottom:10px;margin-bottom:12px;">
            <div style="font-size:18px;font-weight:700;">${gv('SELLER_NAME')}</div>
            <div style="font-size:12px;color:#666;">${escapeHtml(gv('SELLER_STREET'))}, ${escapeHtml(gv('SELLER_CITY'))} ${escapeHtml(gv('SELLER_POSTAL_CODE'))}</div>
            <div style="font-size:12px;color:#666;">PIB: ${escapeHtml(gv('SELLER_PIB'))} | MB: ${escapeHtml(gv('SELLER_MATICNI_BROJ'))}</div>
        </div>
        <h2 style="text-align:center;margin-bottom:14px;font-size:18px;">OTPREMNICA AGROHEMIJE</h2>
        <div style="background:#f5f5f0;padding:10px;border-radius:8px;margin-bottom:12px;font-size:13px;">
            <div><strong>${koop.Ime || ''} ${koop.Prezime || ''}</strong></div>
            <div>${escapeHtml(koop.Adresa || '')}, ${escapeHtml(koop.Mesto || '')}</div>
            <div>JMBG: ${escapeHtml(koop.JMBG || '________')} | BPG: ${escapeHtml(koop.BPGBroj || '________')}</div>
            ${data.parcelaID ? '<div>Parcela: ' + escapeHtml(data.parcelaID) + '</div>' : ''}
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
        <div style="font-size:12px;color:#666;margin-bottom:12px;">Datum: ${escapeHtml(data.datum)}${data.napomena ? ' | Napomena: ' + escapeHtml(data.napomena) : ''}</div>

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
        <button onclick="closeIzdOtpremnicaModal()" style="width:100%;padding:10px;margin-top:8px;font-size:14px;background:none;color:#666;border:1px solid #ccc;border-radius:8px;">Zatvori</button>
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
        if (!json) { showToast('Nema konekcije', 'error'); return; }
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
        if (!json) { showToast('Nema konekcije', 'error'); return; }
        if (json.success) { showToast('PDF sačuvan na Drive!', 'success'); }
        else { showToast('Greška: ' + (json.error || ''), 'error'); }
    } catch(e) {
        showToast('Greška pri generisanju PDF-a', 'error');
    }
}

function closeIzdOtpremnicaModal() {
    destroySignaturePad('sigIzdavalac');
    destroySignaturePad('sigPrimalac');
    const modal = document.getElementById('izdOtpremnicaModal');
    if (modal) modal.style.display = 'none';
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

function loadMgmtAgroStanje() {
    document.getElementById('mgmtAgroStanjeList').innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Pokrenite ExportMgmtReports iz Excela</p>';
}
