// ============================================================
// KNJIGA POLJA — Bilans, Troškovi, Lager
// Depends on: CONFIG, AppState, stammdaten, db, apiFetch,
//             apiPost, safeAsync, dbPut, dbGetAll, dbGetByIndex,
//             fmtDate, escapeHtml, showToast
// ============================================================

const TROSAK_KATEGORIJE = [
    { id: 'gorivo', label: '⛽ Gorivo' },
    { id: 'popravka', label: '🔧 Popravka' },
    { id: 'osiguranje', label: '🛡️ Osiguranje' },
    { id: 'sertifikacija', label: '📋 Sertifikacija' },
    { id: 'analiza', label: '🔬 Analiza' },
    { id: 'navodnjavanje', label: '💧 Navodnjavanje' },
    { id: 'ambalaza', label: '📦 Ambalaža' },
    { id: 'radna_snaga', label: '👷 Radna snaga' },
    { id: 'zakup', label: '🏞️ Zakup' },
    { id: 'transport', label: '🚛 Transport' },
    { id: 'ostalo', label: '📌 Ostalo' }
];

let kpData = { proizvodnja: [], tretmani: [], troskovi: [], lager: [] };
let selectedTrosakKat = '';
let _kpLoaded = false;

// ============================================================
// INIT
// ============================================================
async function loadKnjigaPolja() {
    await safeAsync(async () => {
        kpPopulateDropdowns();
        await kpFetchAll();
        kpLoadBilans();
        _kpLoaded = true;
    }, 'Greška pri učitavanju knjige polja');
}

function kpPopulateDropdowns() {
    // Parcela filter
    const pSel = document.getElementById('kpParcelaSel');
    if (pSel && pSel.options.length <= 1) {
        pSel.innerHTML = '<option value="">-- Sve parcele --</option>';
        (stammdaten.parcele || [])
            .filter(p => p.KooperantID === CONFIG.ENTITY_ID)
            .forEach(p => {
                const o = document.createElement('option');
                o.value = p.ParcelaID;
                o.textContent = escapeHtml((p.KatBroj || p.ParcelaID) + ' — ' + (p.Kultura || '?'));
                pSel.appendChild(o);
            });
    }

    // Troškovi parcela dropdown
    const tpSel = document.getElementById('trosakParcela');
    if (tpSel && tpSel.options.length <= 1) {
        tpSel.innerHTML = '<option value="">-- Opšti trošak --</option>';
        (stammdaten.parcele || [])
            .filter(p => p.KooperantID === CONFIG.ENTITY_ID)
            .forEach(p => {
                const o = document.createElement('option');
                o.value = p.ParcelaID;
                o.textContent = escapeHtml((p.KatBroj || p.ParcelaID) + ' — ' + (p.Kultura || '?'));
                tpSel.appendChild(o);
            });
    }

    // Troškovi kategorije grid
    const katGrid = document.getElementById('trosakKatGrid');
    if (katGrid && !katGrid.children.length) {
        katGrid.innerHTML = TROSAK_KATEGORIJE.map(k =>
            `<button class="trosak-kat-btn" onclick="selectTrosakKat(this,'${k.id}')">${k.label}</button>`
        ).join('');
    }

    // Datum default
    const dEl = document.getElementById('trosakDatum');
    if (dEl && !dEl.value) dEl.value = new Date().toISOString().split('T')[0];
}

// ============================================================
// FETCH ALL DATA
// ============================================================
async function kpFetchAll() {
    // Proizvodnja
    const pJson = await safeAsync(async () => {
        return await apiFetch('action=getKooperantProizvodnja&kooperantID=' + encodeURIComponent(CONFIG.ENTITY_ID));
    });
    if (pJson && pJson.success) kpData.proizvodnja = pJson.records || [];

    // Tretmani — koristi getTretmaniCached ako postoji, inače API
    if (typeof getTretmaniCached === 'function') {
        const cached = await getTretmaniCached(false);
        kpData.tretmani = cached || [];
    } else {
        const tJson = await safeAsync(async () => {
            return await apiFetch('action=getTretmani&kooperantID=' + encodeURIComponent(CONFIG.ENTITY_ID));
        });
        if (tJson && tJson.success) kpData.tretmani = tJson.records || [];
    }

    // Troškovi — iz IndexedDB + server
    let localTroskovi = [];
    try {
        localTroskovi = await dbGetAll(db, 'troskovi');
    } catch (e) {}

    const trJson = await safeAsync(async () => {
        return await apiFetch('action=getTroskovi&kooperantID=' + encodeURIComponent(CONFIG.ENTITY_ID));
    });

    const serverTroskovi = (trJson && trJson.success) ? (trJson.records || []) : [];

    // Merge local pending + server
    kpData.troskovi = kpMergeTroskovi(localTroskovi, serverTroskovi);

    // Lager
    kpData.lager = (stammdaten.magacinkoop || []).filter(r => r.KooperantID === CONFIG.ENTITY_ID);

    // Auto-generated radna snaga troškovi
    kpCalcRadnaSnaga();
}

function kpMergeTroskovi(local, server) {
    if (typeof mergeOfflineRecords === 'function') {
        return mergeOfflineRecords(local, server, function (r) {
            return {
                clientRecordID: r.clientRecordID || '',
                createdAtClient: normalizeIso(r.createdAtClient),
                updatedAtClient: normalizeIso(r.updatedAtClient || r.createdAtClient),
                updatedAtServer: normalizeIso(r.updatedAtServer),
                kooperantID: r.kooperantID || r.KooperantID || '',
                parcelaID: r.parcelaID || r.ParcelaID || '',
                datum: r.datum || r.Datum || '',
                Datum: r.datum || r.Datum || '',
                Kategorija: r.kategorija || r.Kategorija || '',
                Opis: r.opis || r.Opis || '',
                Iznos: parseFloat(r.iznos || r.Iznos) || 0,
                DokumentBroj: r.dokumentBroj || r.DokumentBroj || '',
                Napomena: r.napomena || r.Napomena || '',
                ParcelaID: r.parcelaID || r.ParcelaID || '',
                syncStatus: r.syncStatus || 'synced'
            };
        });
    }

    // Fallback bez mergeOfflineRecords
    const merged = new Map();
    (server || []).forEach(r => {
        const id = r.ClientRecordID || r.clientRecordID || '';
        if (id) merged.set(id, r);
    });
    (local || []).forEach(r => {
        if (!r || !r.clientRecordID) return;
        if (r.syncStatus === 'pending') merged.set(r.clientRecordID, r);
        else if (!merged.has(r.clientRecordID)) merged.set(r.clientRecordID, r);
    });
    return Array.from(merged.values());
}

function kpCalcRadnaSnaga() {
    const config = stammdaten.config || [];
    const crsConf = config.find(c => c.Parameter === 'CenaRadaSat');
    const cenaRadaSat = crsConf ? parseFloat(crsConf.Vrednost) || 0 : 0;
    if (cenaRadaSat <= 0) return;

    (kpData.tretmani || []).forEach(t => {
        const min = parseInt(t.TrajanjeMinuta || t.trajanjeMinuta) || 0;
        if (min <= 0) return;
        const sati = min / 60;
        const iznos = Math.round(sati * cenaRadaSat);
        if (iznos <= 0) return;

        const tId = t.ClientRecordID || t.clientRecordID || 'x_no_match';
        const exists = kpData.troskovi.some(tr =>
            (tr.Kategorija || tr.kategorija) === 'radna_snaga' &&
            (tr.Napomena || tr.napomena || '').includes(tId)
        );
        if (exists) return;

        kpData.troskovi.push({
            _auto: true,
            Datum: t.Datum || t.datum || '',
            ParcelaID: t.ParcelaID || t.parcelaID || '',
            Kategorija: 'radna_snaga',
            Opis: (t.Mera || t.mera || '') + ' — ' + sati.toFixed(1) + 'h × ' + cenaRadaSat.toLocaleString('sr') + ' RSD/h',
            Iznos: iznos,
            Napomena: 'Auto: ' + tId
        });
    });
}

// ============================================================
// BILANS
// ============================================================
function kpLoadBilans() {
    const parcelaFilter = document.getElementById('kpParcelaSel') ?
        document.getElementById('kpParcelaSel').value : '';
    const parcele = (stammdaten.parcele || []).filter(p => p.KooperantID === CONFIG.ENTITY_ID);

    let proizvodnja = kpData.proizvodnja;
    let tretmani = kpData.tretmani;
    let troskovi = kpData.troskovi;

    if (parcelaFilter) {
        proizvodnja = proizvodnja.filter(r => r.ParcelaID === parcelaFilter);
        tretmani = tretmani.filter(r => (r.ParcelaID || r.parcelaID) === parcelaFilter);
        troskovi = troskovi.filter(r => (r.ParcelaID || r.parcelaID) === parcelaFilter);
    }

    // Calc
    const proizvodnjaTotal = proizvodnja.reduce((s, r) =>
        s + (parseFloat(r.Kolicina) || 0) * (parseFloat(r.Cena) || 0), 0);

    let agrohemijaTotal = 0;
    tretmani.forEach(t => {
        const artID = t.ArtikalID || t.artikalID || '';
        const kol = parseFloat(t.KolicinaUpotrebljena || t.kolicinaUpotrebljena) || 0;
        if (!artID || kol <= 0) return;
        const lagerItem = kpData.lager.find(l => l.ArtikalID === artID);
        const wac = lagerItem ? (parseFloat(lagerItem.WAC) || parseFloat(lagerItem.CenaPoJedinici) || 0) : 0;
        agrohemijaTotal += kol * wac;
    });

    const troskoviTotal = troskovi.reduce((s, r) => s + (parseFloat(r.Iznos || r.iznos) || 0), 0);

    let radniMinuti = 0;
    tretmani.forEach(t => { radniMinuti += parseInt(t.TrajanjeMinuta || t.trajanjeMinuta) || 0; });
    const radniSati = radniMinuti / 60;

    const rezultat = proizvodnjaTotal - agrohemijaTotal - troskoviTotal;

    // Render
    const rows = document.getElementById('kpBilansRows');
    if (rows) {
        rows.innerHTML = `
            <div class="kp-bilans-row"><span class="kp-label">📦 Proizvodnja</span><span class="kp-value kp-pos">+${proizvodnjaTotal.toLocaleString('sr')} RSD</span></div>
            <div class="kp-bilans-row"><span class="kp-label">🧪 Agrohemija</span><span class="kp-value kp-neg">-${agrohemijaTotal.toLocaleString('sr')} RSD</span></div>
            <div class="kp-bilans-row"><span class="kp-label">💰 Ostali troškovi</span><span class="kp-value kp-neg">-${troskoviTotal.toLocaleString('sr')} RSD</span></div>
            <div class="kp-bilans-row"><span class="kp-label">⏱️ Radni sati</span><span class="kp-value">${Math.floor(radniSati)}h ${Math.round(radniMinuti % 60)}min</span></div>
            <div class="kp-bilans-row kp-total"><span class="kp-label">📊 REZULTAT</span><span class="kp-value ${rezultat >= 0 ? 'kp-pos' : 'kp-neg'}">${rezultat >= 0 ? '+' : ''}${rezultat.toLocaleString('sr')} RSD</span></div>
        `;
    }

    kpRenderTretmani(tretmani);
    kpRenderOtkupi(proizvodnja);
    kpRenderTroskovi(troskovi);
    kpRenderLager();

    if (!parcelaFilter) kpRenderSummary(parcele);
    else {
        const sumDiv = document.getElementById('kpSummary');
        if (sumDiv) sumDiv.innerHTML = '';
    }
}

// ============================================================
// RENDER FUNCTIONS
// ============================================================
function kpRenderTretmani(tretmani) {
    const title = document.getElementById('kpTretmaniTitle');
    const list = document.getElementById('kpTretmaniList');
    if (!list) return;

    if (!tretmani.length) {
        if (title) title.style.display = 'none';
        list.innerHTML = '';
        return;
    }

    if (title) title.style.display = 'block';
    const icons = { Zastita: '🛡️', Prihrana: '🌱', Rezidba: '✂️', Zalivanje: '💧', Berba: '🍎' };

    list.innerHTML = tretmani.map(t => {
        const mera = t.Mera || t.mera || '';
        const art = t.ArtikalNaziv || t.artikalNaziv || '';
        const kol = t.KolicinaUpotrebljena || t.kolicinaUpotrebljena || '';
        const jm = t.JedinicaMere || t.jedinicaMere || '';
        const min = parseInt(t.TrajanjeMinuta || t.trajanjeMinuta) || 0;
        const timeStr = min > 0 ? Math.floor(min / 60) + 'h ' + (min % 60) + 'min' : '';

        return `<div class="queue-item">
            <div class="qi-header">
                <span class="qi-koop">${icons[mera] || ''} ${escapeHtml(mera)}</span>
                <span class="qi-time">${escapeHtml(fmtDate(t.Datum || t.datum))}</span>
            </div>
            <div class="qi-detail">
                ${escapeHtml(t.ParcelaID || t.parcelaID || '')}
                ${art ? ' | ' + escapeHtml(art) + ' ' + escapeHtml(String(kol)) + ' ' + escapeHtml(jm) : ''}
                ${timeStr ? ' | ⏱️ ' + escapeHtml(timeStr) : ''}
            </div>
        </div>`;
    }).join('');
}

function kpRenderOtkupi(otkupi) {
    const title = document.getElementById('kpOtkupiTitle');
    const list = document.getElementById('kpOtkupiList');
    if (!list) return;

    if (!otkupi.length) {
        if (title) title.style.display = 'none';
        list.innerHTML = '';
        return;
    }

    if (title) title.style.display = 'block';

    list.innerHTML = otkupi.map(r => {
        const vr = ((parseFloat(r.Kolicina) || 0) * (parseFloat(r.Cena) || 0)).toLocaleString('sr');
        return `<div class="queue-item" style="border-left-color:var(--success);">
            <div class="qi-header">
                <span class="qi-koop">${escapeHtml(r.VrstaVoca || '')} ${escapeHtml(r.Klasa || 'I')}</span>
                <span class="qi-time">${escapeHtml(fmtDate(r.Datum))}</span>
            </div>
            <div class="qi-detail">
                ${(parseFloat(r.Kolicina) || 0).toLocaleString('sr')} kg × ${(parseFloat(r.Cena) || 0).toLocaleString('sr')}
                = <strong>${vr} RSD</strong>
                ${r.ParcelaID ? ' | ' + escapeHtml(r.ParcelaID) : ''}
            </div>
        </div>`;
    }).join('');
}

function kpRenderTroskovi(troskovi) {
    const title = document.getElementById('kpTroskoviTitle');
    const list = document.getElementById('kpTroskoviList');
    if (!list) return;

    if (!troskovi.length) {
        if (title) title.style.display = 'none';
        list.innerHTML = '';
        return;
    }

    if (title) title.style.display = 'block';

    list.innerHTML = troskovi.map(r => {
        const iznos = parseFloat(r.Iznos || r.iznos) || 0;
        const isAuto = r._auto;
        const kat = r.Kategorija || r.kategorija || '';
        const opis = r.Opis || r.opis || '';
        const datum = r.Datum || r.datum || '';

        return `<div class="queue-item" style="border-left-color:var(--warning);">
            <div class="qi-header">
                <span class="qi-koop">${escapeHtml(kat)} ${isAuto ? '(auto)' : ''}</span>
                <span class="qi-time">${escapeHtml(fmtDate(datum))}</span>
            </div>
            <div class="qi-detail">${escapeHtml(opis)} | <strong>${iznos.toLocaleString('sr')} RSD</strong></div>
        </div>`;
    }).join('');
}

function kpRenderLager() {
    const list = document.getElementById('kpLagerList');
    if (!list) return;

    const lager = kpData.lager.filter(r => parseFloat(r.Stanje) > 0);

    if (!lager.length) {
        list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:12px;">Nema preparata na lageru</p>';
        return;
    }

    list.innerHTML = lager.map(r => {
        const stanje = parseFloat(r.Stanje) || 0;
        const wac = parseFloat(r.WAC) || parseFloat(r.CenaPoJedinici) || 0;
        const vrednost = stanje * wac;

        return `<div class="queue-item">
            <div class="qi-header">
                <span class="qi-koop">${escapeHtml(r.ArtikalNaziv || r.ArtikalID)}</span>
                <span class="qi-time">${stanje.toLocaleString('sr')} ${escapeHtml(r.JedinicaMere || 'kg')}</span>
            </div>
            <div class="qi-detail">
                WAC: ${wac.toLocaleString('sr')} RSD | Vrednost: <strong>${vrednost.toLocaleString('sr')} RSD</strong>
            </div>
        </div>`;
    }).join('');
}

function kpRenderSummary(parcele) {
    const div = document.getElementById('kpSummary');
    if (!div || !parcele.length) {
        if (div) div.innerHTML = '';
        return;
    }

    const rows = [];
    let grandProd = 0, grandAgro = 0, grandTros = 0, grandRes = 0;

    parcele.forEach(p => {
        const pid = p.ParcelaID;
        const prod = kpData.proizvodnja
            .filter(r => r.ParcelaID === pid)
            .reduce((s, r) => s + (parseFloat(r.Kolicina) || 0) * (parseFloat(r.Cena) || 0), 0);

        let agro = 0;
        kpData.tretmani
            .filter(t => (t.ParcelaID || t.parcelaID) === pid)
            .forEach(t => {
                const artID = t.ArtikalID || t.artikalID || '';
                const kol = parseFloat(t.KolicinaUpotrebljena || t.kolicinaUpotrebljena) || 0;
                if (!artID || kol <= 0) return;
                const li = kpData.lager.find(l => l.ArtikalID === artID);
                const wac = li ? (parseFloat(li.WAC) || parseFloat(li.CenaPoJedinici) || 0) : 0;
                agro += kol * wac;
            });

        const tros = kpData.troskovi
            .filter(r => (r.ParcelaID || r.parcelaID) === pid)
            .reduce((s, r) => s + (parseFloat(r.Iznos || r.iznos) || 0), 0);

        const res = prod - agro - tros;
        rows.push({ name: escapeHtml((p.KatBroj || pid) + ' ' + (p.Kultura || '')), prod, agro, tros, res });
        grandProd += prod; grandAgro += agro; grandTros += tros; grandRes += res;
    });

    // Troškovi bez parcele
    const bezParcele = kpData.troskovi
        .filter(r => !(r.ParcelaID || r.parcelaID))
        .reduce((s, r) => s + (parseFloat(r.Iznos || r.iznos) || 0), 0);

    if (bezParcele > 0) {
        rows.push({ name: 'Opšti troškovi', prod: 0, agro: 0, tros: bezParcele, res: -bezParcele });
        grandTros += bezParcele;
        grandRes -= bezParcele;
    }

    div.innerHTML = `
        <div class="kp-section-title">Sve parcele — pregled</div>
        <table class="kp-summary-table">
            <tr>
                <th>Parcela</th>
                <th style="text-align:right;">Proizv.</th>
                <th style="text-align:right;">Agroh.</th>
                <th style="text-align:right;">Trošk.</th>
                <th style="text-align:right;">Rezultat</th>
            </tr>
            ${rows.map(r => `<tr>
                <td>${r.name}</td>
                <td style="text-align:right;color:var(--success);">${r.prod.toLocaleString('sr')}</td>
                <td style="text-align:right;color:var(--danger);">${r.agro.toLocaleString('sr')}</td>
                <td style="text-align:right;color:var(--danger);">${r.tros.toLocaleString('sr')}</td>
                <td style="text-align:right;font-weight:600;color:${r.res >= 0 ? 'var(--success)' : 'var(--danger)'};">${r.res.toLocaleString('sr')}</td>
            </tr>`).join('')}
            <tr class="kp-row-total">
                <td>UKUPNO</td>
                <td style="text-align:right;">${grandProd.toLocaleString('sr')}</td>
                <td style="text-align:right;">${grandAgro.toLocaleString('sr')}</td>
                <td style="text-align:right;">${grandTros.toLocaleString('sr')}</td>
                <td style="text-align:right;color:${grandRes >= 0 ? 'var(--success)' : 'var(--danger)'};">${grandRes.toLocaleString('sr')}</td>
            </tr>
        </table>
    `;
}

// ============================================================
// TROŠKOVI — SAVE
// ============================================================
function selectTrosakKat(btn, kat) {
    document.querySelectorAll('.trosak-kat-btn').forEach(b => b.classList.remove('selected'));
    btn.classList.add('selected');
    selectedTrosakKat = kat;
}

async function kpSaveTrosak() {
    if (!selectedTrosakKat) { showToast('Izaberite kategoriju', 'error'); return; }

    const iznos = parseFloat(document.getElementById('trosakIznos').value) || 0;
    if (iznos <= 0) { showToast('Unesite iznos', 'error'); return; }

    const nowIso = new Date().toISOString();

    const record = {
        clientRecordID: (window.crypto && typeof window.crypto.randomUUID === 'function')
            ? window.crypto.randomUUID()
            : ('tros-' + Date.now() + '-' + Math.floor(Math.random() * 1000000)),
        createdAtClient: nowIso,
        updatedAtClient: nowIso,
        kooperantID: CONFIG.ENTITY_ID,
        parcelaID: document.getElementById('trosakParcela').value || '',
        datum: document.getElementById('trosakDatum').value || nowIso.split('T')[0],
        kategorija: selectedTrosakKat,
        opis: document.getElementById('trosakOpis').value || '',
        iznos: iznos,
        dokumentBroj: document.getElementById('trosakDokBroj').value || '',
        napomena: '',
        syncStatus: 'pending',
        syncAttempts: 0,
        lastSyncError: ''
    };

    await dbPut(db, 'troskovi', record);
    showToast('Trošak sačuvan: ' + iznos.toLocaleString('sr') + ' RSD', 'success');

    // Sync
    if (navigator.onLine) {
        const json = await safeAsync(async () => {
            return await apiPost('syncTrosak', {
                kooperantID: CONFIG.ENTITY_ID,
                records: [record]
            });
        });

        if (json && json.success) {
            record.syncStatus = 'synced';
            record.lastSyncError = '';
            await dbPut(db, 'troskovi', record);
        } else {
            record.lastSyncError = (json && json.error) ? json.error : 'Sync neuspešan';
            record.syncAttempts = 1;
            await dbPut(db, 'troskovi', record);
        }
    }

    // Reset forma
    document.getElementById('trosakOpis').value = '';
    document.getElementById('trosakIznos').value = '';
    document.getElementById('trosakDokBroj').value = '';
    document.querySelectorAll('.trosak-kat-btn').forEach(b => b.classList.remove('selected'));
    selectedTrosakKat = '';

    // Reload
    await kpFetchAll();
    kpLoadBilans();
}

// ============================================================
// INVALIDATE
// ============================================================
function invalidateKpCache() {
    _kpLoaded = false;
    kpData = { proizvodnja: [], tretmani: [], troskovi: [], lager: [] };
}

