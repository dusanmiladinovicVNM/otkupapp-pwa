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
    var scanner = new Html5Qrcode('qr-reader-fiskalni');

    scanner.start(
        { facingMode: 'environment' },
        { 
            fps: 15,
            qrbox: { width: 300, height: 300 },
            aspectRatio: 1.0,
            disableFlip: false
        },
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
        readerDiv.style.display = 'none';
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

