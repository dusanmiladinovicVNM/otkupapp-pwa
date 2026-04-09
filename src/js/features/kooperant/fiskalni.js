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
function startFiskalniScan() {
    var readerDiv = document.getElementById('qr-reader-fiskalni');
    if (!readerDiv) return;

    readerDiv.style.display = 'block';
    readerDiv.innerHTML = `
        <div style="display:flex;flex-direction:column;gap:10px;padding:12px;">
            <button class="btn-primary" onclick="startFiskalniNative()" style="background:var(--primary);">📷 Skeniraj QR</button>
            <button class="btn-primary" onclick="document.getElementById('fiskalniFileInput').click()" style="background:var(--accent);">🖼️ Slika računa</button>
            <input type="file" id="fiskalniFileInput" accept="image/*" capture="environment" style="display:none;" onchange="scanFiskalniFromFile(this)">
            <video id="fiskalniVideo" playsinline style="display:none;width:100%;border-radius:var(--radius);"></video>
            <canvas id="fiskalniCanvas" style="display:none;"></canvas>
        </div>
    `;
}

async function startFiskalniNative() {
    // Opcija 1: Native BarcodeDetector (Chrome 83+, Safari 16.4+)
    if ('BarcodeDetector' in window) {
        try {
            var detector = new BarcodeDetector({ formats: ['qr_code'] });
            var stream = await navigator.mediaDevices.getUserMedia({
                video: { facingMode: 'environment', width: { ideal: 1920 }, height: { ideal: 1080 } }
            });

            var video = document.getElementById('fiskalniVideo');
            video.style.display = 'block';
            video.srcObject = stream;
            await video.play();

            var scanning = true;

            async function scanFrame() {
                if (!scanning) return;
                try {
                    var barcodes = await detector.detect(video);
                    if (barcodes.length > 0) {
                        scanning = false;
                        stream.getTracks().forEach(function(t) { t.stop(); });
                        video.style.display = 'none';
                        document.getElementById('qr-reader-fiskalni').style.display = 'none';
                        onFiskalniScanned(barcodes[0].rawValue);
                        return;
                    }
                } catch (e) {}
                if (scanning) requestAnimationFrame(scanFrame);
            }

            requestAnimationFrame(scanFrame);

            // Timeout posle 30s
            setTimeout(function() {
                if (scanning) {
                    scanning = false;
                    stream.getTracks().forEach(function(t) { t.stop(); });
                    video.style.display = 'none';
                    showToast('QR nije pronađen. Probajte opciju "Slika računa"', 'info');
                }
            }, 30000);

            return;
        } catch (err) {
            console.log('BarcodeDetector failed, fallback to Html5Qrcode:', err);
        }
    }

    // Opcija 2: Fallback na Html5Qrcode
    startFiskalniHtml5Qr();
}

function startFiskalniHtml5Qr() {
    var readerDiv = document.getElementById('qr-reader-fiskalni');
    var cameraDiv = document.createElement('div');
    cameraDiv.id = 'fiskalniCameraDiv';
    readerDiv.querySelector('div').appendChild(cameraDiv);

    var scanner = new Html5Qrcode('fiskalniCameraDiv');
    scanner.start(
        { facingMode: 'environment' },
        { fps: 15, qrbox: { width: 300, height: 300 } },
        function(text) {
            scanner.stop().then(function() {
                readerDiv.style.display = 'none';
            }).catch(function() {
                readerDiv.style.display = 'none';
            });
            onFiskalniScanned(text);
        },
        function() {}
    ).catch(function(err) {
        showToast('Kamera nije dostupna: ' + err, 'error');
    });
}

function scanFiskalniFromFile(input) {
    if (!input.files || !input.files[0]) return;

    showToast('Skeniram sliku...', 'info');

    var file = input.files[0];

    // Probaj native BarcodeDetector na slici
    if ('BarcodeDetector' in window) {
        createImageBitmap(file).then(function(bitmap) {
            var detector = new BarcodeDetector({ formats: ['qr_code'] });
            return detector.detect(bitmap);
        }).then(function(barcodes) {
            if (barcodes.length > 0) {
                document.getElementById('qr-reader-fiskalni').style.display = 'none';
                onFiskalniScanned(barcodes[0].rawValue);
            } else {
                // Fallback na Html5Qrcode file scan
                scanFiskalniFileHtml5(file);
            }
        }).catch(function() {
            scanFiskalniFileHtml5(file);
        });
    } else {
        scanFiskalniFileHtml5(file);
    }

    input.value = '';
}

function scanFiskalniFileHtml5(file) {
    var tempDiv = document.getElementById('fiskalniCameraDiv');
    if (!tempDiv) {
        tempDiv = document.createElement('div');
        tempDiv.id = 'fiskalniCameraDiv';
        tempDiv.style.display = 'none';
        document.getElementById('qr-reader-fiskalni').querySelector('div').appendChild(tempDiv);
    }

    var scanner = new Html5Qrcode('fiskalniCameraDiv');
    scanner.scanFile(file, true)
        .then(function(text) {
            document.getElementById('qr-reader-fiskalni').style.display = 'none';
            onFiskalniScanned(text);
        })
        .catch(function() {
            showToast('QR nije pronađen na slici. Pokušajte bližu i oštriju sliku.', 'error');
        });
}
// ============================================================
// PARSE
// ============================================================
async function onFiskalniScanned(text) {
    // DEBUG — prikaži šta je skenirano
    showToast('QR: ' + text.substring(0, 80), 'info');
    console.log('FISKALNI QR RAW:', text);
    
    let url = text;
    if (!text.includes('suf.purs.gov.rs') && !text.startsWith('https://')) {
        try { url = decodeURIComponent(text); } catch (e) { url = text; }
    }

    if (!url.startsWith('http')) {
        showToast('Nije fiskalni QR kod: ' + text.substring(0, 50), 'error');
        return;
    }

    showToast('Učitavanje fiskalnog...', 'info');

    const json = await safeAsync(async () => {
        return await apiPost('parseFiskalni', {
            kooperantID: CONFIG.ENTITY_ID,
            verificationUrl: url
        });
    }, 'Greška pri učitavanju fiskalnog računa');

    if (!json) return;

    if (json.duplicate) {
        showToast('Ovaj račun je već skeniran', 'error');
        return;
    }

    if (!json.success) {
        showToast(json.error || 'Greška', 'error');
        return;
    }

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

                const artikalCell = s.artikalID
                    ? '<span class="' + matchClass + '">✅ ' + escapeHtml(s.artikalNaziv) + '</span>'
                    : '<select id="fisMap' + i + '" style="font-size:11px;padding:4px;max-width:120px;" onchange="onFiskalniMap(' + i + ', this.value)">' +
                      '<option value="">❓ Izaberi...</option>' +
                      artikli.map(a => '<option value="' + escapeHtml(a.ArtikalID) + '">' + escapeHtml(a.Naziv) + '</option>').join('') +
                      '</select>';

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
function onFiskalniMap(index, artikalID) {
    if (!artikalID) return;
    const art = (stammdaten.artikli || []).find(a => a.ArtikalID === artikalID);
    if (!art) return;

    fiskalniStavke[index].artikalID = artikalID;
    fiskalniStavke[index].artikalNaziv = art.Naziv;
    fiskalniStavke[index].matchConfidence = 'manual';

    const chk = document.getElementById('fisChk' + index);
    if (chk) chk.checked = true;
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

