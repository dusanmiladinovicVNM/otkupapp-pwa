// ============================================================
// FISKALNI RAČUN → LAGER
// Depends on: CONFIG, stammdaten, apiFetch, apiPost,
//             safeAsync, escapeHtml, showToast, Html5Qrcode
// ============================================================

let fiskalniStavke = [];
let fiskalniMeta = {};

// ============================================================
// QR SCAN
// ============================================================
async function startFiskalniScan() {
    var readerDiv = document.getElementById('qr-reader-fiskalni');
    if (!readerDiv) return;

    // Proveri native podršku
    if (!('BarcodeDetector' in window)) {
        showToast('Ovaj uređaj ne podržava skeniranje. Koristite opciju "Slikaj".', 'error');
        readerDiv.style.display = 'block';
        readerDiv.innerHTML = `
            <div style="padding:12px;">
                <label class="btn-primary" style="background:var(--primary);text-align:center;cursor:pointer;display:block;">
                    📷 Slikaj QR kod računa
                    <input type="file" accept="image/*" capture="environment" style="display:none;" onchange="scanFiskalniFromPhoto(this)">
                </label>
            </div>
        `;
        return;
    }

    try {
        var stream = await navigator.mediaDevices.getUserMedia({
            video: {
                facingMode: 'environment',
                width: { ideal: 1920 },
                height: { ideal: 1080 }
            }
        });

        readerDiv.style.display = 'block';
        readerDiv.innerHTML = `
            <div style="position:relative;">
                <video id="fiskalniVideo" playsinline autoplay style="width:100%;border-radius:var(--radius);"></video>
                <div style="position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);width:250px;height:250px;border:3px solid #ffd60a;border-radius:12px;"></div>
                <button onclick="stopFiskalniScan()" style="position:absolute;top:8px;right:8px;background:rgba(0,0,0,0.6);color:white;border:none;border-radius:50%;width:36px;height:36px;font-size:18px;cursor:pointer;">✕</button>
            </div>
        `;

        var video = document.getElementById('fiskalniVideo');
        video.srcObject = stream;
        await video.play();

        var detector = new BarcodeDetector({ formats: ['qr_code'] });
        var scanning = true;

        // Sačuvaj referencu za cleanup
        window._fiskalniStream = stream;
        window._fiskalniScanning = true;

        async function scanFrame() {
            if (!window._fiskalniScanning) return;

            try {
                var barcodes = await detector.detect(video);
                if (barcodes.length > 0) {
                    window._fiskalniScanning = false;
                    stream.getTracks().forEach(function(t) { t.stop(); });
                    readerDiv.style.display = 'none';
                    onFiskalniScanned(barcodes[0].rawValue);
                    return;
                }
            } catch (e) {}

            if (window._fiskalniScanning) {
                requestAnimationFrame(scanFrame);
            }
        }

        requestAnimationFrame(scanFrame);

    } catch (err) {
        showToast('Kamera nije dostupna: ' + err.message, 'error');
        readerDiv.style.display = 'none';
    }
}

function stopFiskalniScan() {
    window._fiskalniScanning = false;
    if (window._fiskalniStream) {
        window._fiskalniStream.getTracks().forEach(function(t) { t.stop(); });
        window._fiskalniStream = null;
    }
    var readerDiv = document.getElementById('qr-reader-fiskalni');
    if (readerDiv) readerDiv.style.display = 'none';
}

async function scanFiskalniFromPhoto(input) {
    if (!input.files || !input.files[0]) return;

    showToast('Čitam QR sa slike...', 'info');

    try {
        // Smanji sliku na max 1024px pre slanja
        var base64 = await resizeImageForQR(input.files[0], 1024);

        var json = await safeAsync(async function() {
            return await apiPost('parseFiskalniImage', {
                kooperantID: CONFIG.ENTITY_ID,
                imageBase64: base64
            });
        }, 'Greška pri čitanju računa');

        if (!json) return;
        if (json.duplicate) { showToast('Ovaj račun je već skeniran', 'error'); return; }
        if (!json.success) { showToast(json.error || 'QR nije pronađen', 'error'); return; }

        document.getElementById('qr-reader-fiskalni').style.display = 'none';

        fiskalniMeta = {
            invoiceNumber: json.invoiceNumber,
            company: json.company,
            date: json.date,
            totalAmount: json.totalAmount,
            verificationUrl: json.verificationUrl || ''
        };
        fiskalniStavke = json.items || [];
        renderFiskalniResult();
    } catch (err) {
        showToast('Greška: ' + err.message, 'error');
    }

    input.value = '';
}

function resizeImageForQR(file, maxSize) {
    return new Promise(function(resolve, reject) {
        var img = new Image();
        img.onload = function() {
            var w = img.width;
            var h = img.height;

            // Smanji ako je veće od maxSize
            if (w > maxSize || h > maxSize) {
                if (w > h) {
                    h = Math.round(h * maxSize / w);
                    w = maxSize;
                } else {
                    w = Math.round(w * maxSize / h);
                    h = maxSize;
                }
            }

            var canvas = document.createElement('canvas');
            canvas.width = w;
            canvas.height = h;
            var ctx = canvas.getContext('2d');
            ctx.drawImage(img, 0, 0, w, h);

            // JPEG kvalitet 0.85 — dovoljno za QR, manja veličina
            var dataUrl = canvas.toDataURL('image/jpeg', 0.85);
            resolve(dataUrl.split(',')[1]);
        };
        img.onerror = reject;
        img.src = URL.createObjectURL(file);
    });
}


function fileToBase64(file) {
    return new Promise(function(resolve, reject) {
        var reader = new FileReader();
        reader.onload = function() { resolve(reader.result.split(',')[1]); };
        reader.onerror = reject;
        reader.readAsDataURL(file);
    });
}
// ============================================================
// PARSE
// ============================================================
async function onFiskalniScanned(text) {
    var url = text;
    if (!text.startsWith('http')) {
        try { url = decodeURIComponent(text); } catch (e) { url = text; }
    }

    if (!url.startsWith('http')) {
        showToast('Nije fiskalni QR kod', 'error');
        return;
    }

    showToast('Učitavanje fiskalnog...', 'info');

    var json = await safeAsync(async function() {
        return await apiPost('parseFiskalni', {
            kooperantID: CONFIG.ENTITY_ID,
            verificationUrl: url
        });
    }, 'Greška pri učitavanju fiskalnog računa');

    if (!json) return;
    if (json.duplicate) { showToast('Ovaj račun je već skeniran', 'error'); return; }
    if (!json.success) { showToast(json.error || 'Greška', 'error'); return; }

    document.getElementById('qr-reader-fiskalni').style.display = 'none';

    fiskalniMeta = {
        invoiceNumber: json.invoiceNumber,
        company: json.company,
        date: json.date,
        totalAmount: json.totalAmount,
        verificationUrl: url
    };
    fiskalniStavke = json.items || [];
    renderFiskalniResult();
}

// ============================================================
// RENDER
// ============================================================
function renderFiskalniResult() {
    const resultDiv = document.getElementById('fiskalniResult');
    const headerDiv = document.getElementById('fiskalniHeader');
    const stavkeDiv = document.getElementById('fiskalniStavke');

    if (!resultDiv || !headerDiv || !stavkeDiv) return;

    resultDiv.style.display = 'block';

    headerDiv.innerHTML =
        '<strong>' + escapeHtml(fiskalniMeta.company || '?') + '</strong> | ' +
        escapeHtml(fiskalniMeta.invoiceNumber || '') + ' | ' +
        escapeHtml(fiskalniMeta.date || '') + ' | ' +
        (fiskalniMeta.totalAmount || 0).toLocaleString('sr') + ' RSD';

    const artikli = stammdaten.artikli || [];

    stavkeDiv.innerHTML = `
        <table class="fis-table">
            <tr>
                <th>☑</th>
                <th>Naziv (fiskalni)</th>
                <th>Kol</th>
                <th>Cena</th>
                <th>Ukupno</th>
                <th>Artikal</th>
            </tr>
            ${fiskalniStavke.map((s, i) => {
                const isPrep = !!s.artikalID;
                const matchClass =
                    (s.matchConfidence === 'exact' || s.matchConfidence === 'mapped') ? 'fis-match-exact' :
                    s.matchConfidence === 'fuzzy' ? 'fis-match-fuzzy' : 'fis-match-none';

                let artikalCell = '';
                if (s.artikalID) {
                    artikalCell = '<span class="' + matchClass + '">✅ ' + escapeHtml(s.artikalNaziv) + '</span>';
                } else {
                    artikalCell = `
                        <select id="fisMap${i}" style="font-size:11px;padding:4px;max-width:120px;" onchange="onFiskalniMap(${i}, this.value)">
                            <option value="">❓ Izaberi...</option>
                            <option value="__NEW__">➕ Novi artikal</option>
                            ${artikli.map(a => '<option value="' + escapeHtml(a.ArtikalID) + '">' + escapeHtml(a.Naziv) + '</option>').join('')}
                        </select>
                        <div id="fisNewArt${i}" style="display:none;margin-top:4px;"></div>
                    `;
                }

                return `<tr>
                    <td><input type="checkbox" id="fisChk${i}" ${isPrep ? 'checked' : ''} style="width:18px;height:18px;"></td>
                    <td style="font-size:12px;">${escapeHtml(s.naziv)}</td>
                    <td>${s.kolicina}</td>
                    <td>${s.jedCena.toLocaleString('sr')}</td>
                    <td style="font-weight:600;">${s.ukupno.toLocaleString('sr')}</td>
                    <td>${artikalCell}</td>
                </tr>`;
            }).join('')}
        </table>
    `;
}

// ============================================================
// MAP ARTIKAL
// ============================================================
function onFiskalniMap(index, value) {
    const newArtDiv = document.getElementById('fisNewArt' + index);

    if (value === '__NEW__') {
        // Prikaži formu za novi artikal
        if (newArtDiv) {
            newArtDiv.style.display = 'block';
            newArtDiv.innerHTML = `
                <div style="background:#f8f8f4;padding:8px;border-radius:6px;font-size:11px;">
                    <input type="text" id="fisNewNaziv${index}" placeholder="Naziv artikla"
                        value="${escapeHtml(fiskalniStavke[index].naziv)}"
                        style="width:100%;padding:4px;font-size:12px;margin-bottom:4px;border:1px solid var(--border);border-radius:4px;">
                    <select id="fisNewTip${index}" style="width:100%;padding:4px;font-size:11px;margin-bottom:4px;border:1px solid var(--border);border-radius:4px;">
                        <option value="Pesticid">Pesticid</option>
                        <option value="Djubrivo">Đubrivo</option>
                        <option value="Ostalo">Ostalo</option>
                    </select>
                    <select id="fisNewJM${index}" style="width:100%;padding:4px;font-size:11px;margin-bottom:4px;border:1px solid var(--border);border-radius:4px;">
                        <option value="l">Litar (l)</option>
                        <option value="kg">Kilogram (kg)</option>
                        <option value="kom">Komad (kom)</option>
                    </select>
                    <button onclick="fiskalniCreateNewArtikal(${index})" class="btn-primary" style="width:100%;padding:6px;font-size:11px;margin-top:2px;">
                        ✅ Kreiraj artikal
                    </button>
                </div>
            `;
        }
        return;
    }

    // Sakrij new art formu ako je bila otvorena
    if (newArtDiv) newArtDiv.style.display = 'none';

    if (!value) return;

    const art = (stammdaten.artikli || []).find(a => a.ArtikalID === value);
    if (!art) return;

    fiskalniStavke[index].artikalID = value;
    fiskalniStavke[index].artikalNaziv = art.Naziv;
    fiskalniStavke[index].matchConfidence = 'manual';

    const chk = document.getElementById('fisChk' + index);
    if (chk) chk.checked = true;
}

async function fiskalniCreateNewArtikal(index) {
    const naziv = (document.getElementById('fisNewNaziv' + index).value || '').trim();
    const tip = document.getElementById('fisNewTip' + index).value || 'Ostalo';
    const jm = document.getElementById('fisNewJM' + index).value || 'kom';

    if (!naziv) {
        showToast('Unesite naziv artikla', 'error');
        return;
    }

    // Generiši privremeni ID
    const tempID = 'ART-NEW-' + Date.now() + '-' + index;

    showToast('Kreiranje artikla...', 'info');

    // Sačuvaj na server
    const json = await safeAsync(async function() {
        return await apiPost('createArtikal', {
            naziv: naziv,
            tip: tip,
            jedinicaMere: jm,
            cenaPoJedinici: fiskalniStavke[index].jedCena || 0
        });
    }, 'Greška pri kreiranju artikla');

    let artikalID = tempID;
    let artikalNaziv = naziv;

    if (json && json.success && json.artikalID) {
        artikalID = json.artikalID;

        // Dodaj u lokalne stammdaten da bude odmah dostupan
        stammdaten.artikli = stammdaten.artikli || [];
        stammdaten.artikli.push({
            ArtikalID: artikalID,
            Naziv: naziv,
            Tip: tip,
            JedinicaMere: jm,
            CenaPoJedinici: fiskalniStavke[index].jedCena || 0
        });

        showToast('Artikal kreiran: ' + escapeHtml(naziv), 'success');
    } else {
        // Offline ili error — koristi privremeni ID, sync će rešiti
        stammdaten.artikli = stammdaten.artikli || [];
        stammdaten.artikli.push({
            ArtikalID: tempID,
            Naziv: naziv,
            Tip: tip,
            JedinicaMere: jm,
            CenaPoJedinici: fiskalniStavke[index].jedCena || 0
        });

        showToast('Artikal sačuvan lokalno: ' + escapeHtml(naziv), 'info');
    }

    // Ažuriraj stavku
    fiskalniStavke[index].artikalID = artikalID;
    fiskalniStavke[index].artikalNaziv = artikalNaziv;
    fiskalniStavke[index].matchConfidence = 'new';

    // Auto-check
    const chk = document.getElementById('fisChk' + index);
    if (chk) chk.checked = true;

    // Sakrij formu, prikaži match
    const newArtDiv = document.getElementById('fisNewArt' + index);
    if (newArtDiv) {
        newArtDiv.style.display = 'block';
        newArtDiv.innerHTML = '<span class="fis-match-exact">✅ ' + escapeHtml(naziv) + ' (novi)</span>';
    }

    // Sačuvaj mapiranje za buduće
    apiPost('saveFiskalniMapiranje', {
        mappings: [{
            fiskalniNaziv: fiskalniStavke[index].naziv,
            artikalID: artikalID,
            artikalNaziv: naziv,
            kooperantID: CONFIG.ENTITY_ID
        }]
    }).catch(function() {});
}

// ============================================================
// SAVE TO LAGER
// ============================================================
async function fiskalniSaveToLager() {
    const selected = [];
    const newMappings = [];
    let hasError = false;

    fiskalniStavke.forEach((s, i) => {
        const chk = document.getElementById('fisChk' + i);
        if (!chk || !chk.checked) return;

        if (!s.artikalID) {
            showToast('Stavka "' + escapeHtml(s.naziv) + '" nema artikal — izaberite ili odčekirajte', 'error');
            hasError = true;
            return;
        }

        selected.push({
            clientRecordID: (window.crypto && typeof window.crypto.randomUUID === 'function')
                ? window.crypto.randomUUID()
                : ('fis-' + Date.now() + '-' + i),
            createdAtClient: new Date().toISOString(),
            naziv: s.naziv,
            artikalID: s.artikalID,
            artikalNaziv: s.artikalNaziv,
            kolicina: s.kolicina,
            jedCena: s.jedCena,
            ukupno: s.ukupno,
            pdvStopa: s.pdvStopa || ''
        });

        if (s.matchConfidence === 'manual' || s.matchConfidence === 'none') {
            newMappings.push({
                fiskalniNaziv: s.naziv,
                artikalID: s.artikalID,
                artikalNaziv: s.artikalNaziv,
                kooperantID: CONFIG.ENTITY_ID
            });
        }
    });

    if (hasError) return;
    if (!selected.length) { showToast('Nema čekiranih stavki', 'error'); return; }

    showToast('Čuvanje...', 'info');

    // Save stavke
    const json = await safeAsync(async () => {
        return await apiPost('saveFiskalni', {
            kooperantID: CONFIG.ENTITY_ID,
            invoiceNumber: fiskalniMeta.invoiceNumber,
            company: fiskalniMeta.company,
            date: fiskalniMeta.date,
            verificationUrl: fiskalniMeta.verificationUrl,
            stavke: selected
        });
    }, 'Greška pri čuvanju fiskalnog');

    if (!json) return;

    if (!json.success) {
        showToast(json.error || 'Greška', 'error');
        return;
    }

    // Save mappings — fire and forget
    if (newMappings.length > 0) {
        apiPost('saveFiskalniMapiranje', {
            mappings: newMappings
        }).catch(() => {});
    }

    showToast(selected.length + ' stavki preneseno u lager', 'success');
    fiskalniCancel();
}

// ============================================================
// CANCEL
// ============================================================
function fiskalniCancel() {
    fiskalniStavke = [];
    fiskalniMeta = {};

    const resultDiv = document.getElementById('fiskalniResult');
    if (resultDiv) resultDiv.style.display = 'none';
}

