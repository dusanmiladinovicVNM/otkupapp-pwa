// ============================================================
// VOZAC: ZBIRNA
// ============================================================
// VOZAC RADI SA OTPREMNICAMA (S5-4b-2).
//
// Do ovog reza je ovde stajao spisak OTKUPNIH redova, filtriran po
// Otkup.VozacID. Ta veza je pala u S5-4a -- ekran otpreme vise ne dira otkupni
// zapis -- pa je spisak bio prazan. Kanon kaze da vozac nosi otpremnice i da se
// zbirna sastavlja od njih, pa se i radi sa njima.
let vozacOtpremnice = [];
let _lastMergedZbirne = null;

// Kilaza i gajbe dolaze iz STAVKI. Zaglavlje otpremnice jos nosi legacy kolone,
// ali se one ne izvoze i ne citaju -- dokument je zaglavlje + stavke.
function otpKgKlase(o, klasa) {
    return (o.stavke || [])
        .filter(s => String(s.klasa || '') === klasa)
        .reduce((zbir, s) => zbir + (Number(s.kolicina) || 0), 0);
}

function otpKg(o) {
    return (o.stavke || []).reduce((zbir, s) => zbir + (Number(s.kolicina) || 0), 0);
}

function otpAmb(o) {
    return (o.stavke || []).reduce((zbir, s) => zbir + (Number(s.kolAmbalaze) || 0), 0);
}

function otpKlaseOpis(o) {
    const klase = Array.from(new Set((o.stavke || []).map(s => String(s.klasa || '')).filter(Boolean)));
    return klase.join('+');
}

async function loadVozacData() {
    const json = await safeAsync(async () => {
        return await apiFetch('action=getVozacOtpremnice');
    }, 'Greška pri učitavanju podataka vozača');

    // readModelChanging: master ciklus menja stanje, server namerno ne salje
    // spisak. Zadrzi poslednje poznato umesto da ga obrises praznim odgovorom.
    //
    // Zato se spisak NE prazni unapred (review #392, P3): ranije je prva linija
    // bila vozacOtpremnice = [], pa je komentar obecavao zadrzavanje koje runtime
    // nije radio.
    if (json && json.readModelChanging) {
        showToast(json.message || 'Stanje vožnji se trenutno menja', 'warning');
    } else if (json && json.success && Array.isArray(json.records)) {
        vozacOtpremnice = json.records.map(r => ({
            otpremnicaID: r.otpremnicaID || '',
            brojOtpremnice: r.brojOtpremnice || '',
            datum: fmtDate(r.datum),
            stanicaID: r.stanicaID || '',
            vozacID: r.vozacID || '',
            vrstaVoca: r.vrstaVoca || '',
            sortaVoca: r.sortaVoca || '',
            tipAmbalaze: r.tipAmbalaze || '',
            predajaID: r.predajaID || '',

            // Prazno = slobodna za zbirnu. Master racuna iz clanstva, pa se
            // posle storna zbirne otpremnica sama vraca u opticaj.
            zbirnaID: r.zbirnaID || '',

            stavke: Array.isArray(r.stavke) ? r.stavke : []
        }));
    } else {
        // Ni podatak ni imenovan razlog -- spisak se prazni, jer bi zadrzan
        // stari ovde bio tvrdnja bez pokrica.
        vozacOtpremnice = [];
    }

    // Jedan fetch za zbirne -- koristi se i za renderovanje
    const zbirne = await getMergedZbirneForVozac();
    _lastMergedZbirne = zbirne;

    // DVE ISTINE, NE JEDNA.
    //
    //   kanonska tekuca  -- server: Otpremnica.zbirnaID (posle master ciklusa)
    //   privremena       -- lokalna zbirna koju master jos nije razresio
    //
    // Prva sama nije dovoljna: postoji legitiman prozor u kom je zbirna vec
    // napravljena a master je jos nije video.
    const rezervisane = rezervisaneOtpremnice(zbirne);

    vozacOtpremnice = vozacOtpremnice.filter(o =>
        !o.zbirnaID && !rezervisane.has(o.otpremnicaID)
    );

    renderVozacOtpremnice();
    renderVozacZbirneFromData(zbirne);
}

async function loadVozacZbirne() {
    // Standalone poziv — kad se zove van loadVozacData (npr. posle cancelZbirna)
    const zbirne = await getMergedZbirneForVozac();
    _lastMergedZbirne = zbirne;
    renderVozacZbirneFromData(zbirne);
}

function renderVozacOtpremnice() {
    const today = getTodayIsoDate();
    const todayOtkupi = vozacOtpremnice.filter(r => r.datum === today);

    const list = document.getElementById('vozacOtpremniceList');
    if (!list) return;

    if (todayOtkupi.length === 0) {
        list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema otpremnica za danas</p>';
        const btn = document.getElementById('btnNovaZbirna');
        if (btn) btn.style.display = 'none';
        return;
    }

    const btn = document.getElementById('btnNovaZbirna');
    if (btn) btn.style.display = '';

    // Group by stanica
    const grouped = {};
    todayOtkupi.forEach(r => {
        const s = r.stanicaID || '?';
        if (!grouped[s]) grouped[s] = { items: [], kg: 0, amb: 0 };
        grouped[s].items.push(r);
        grouped[s].kg += otpKg(r);
        grouped[s].amb += otpAmb(r);
    });

    list.innerHTML = Object.entries(grouped).map(([sta, g]) =>
        `<div class="vblk-sta">${escapeHtml(fmtStanica(sta))} · ${g.kg.toLocaleString('sr')} kg · ${g.amb} amb</div>` +
        g.items.map(r =>
            `<div class="vblk sel">
                <div class="vblk__c">✓</div>
                <div>
                    <div class="vblk__nm">${escapeHtml(r.brojOtpremnice || r.otpremnicaID)}</div>
                    <div class="vblk__s">${escapeHtml(r.vrstaVoca)} ${escapeHtml(otpKlaseOpis(r))} · ${otpAmb(r)} amb</div>
                </div>
                <div class="vblk__kg">${otpKg(r).toLocaleString('sr')}<span class="u"> kg</span></div>
            </div>`
        ).join('')
    ).join('');
}

// Spisak otpremnica koji je vozac STVARNO video kad je otvorio ekran.
//
// Potvrda se meri prema njemu, ne prema "svemu sto je sada slobodno": osvezavanje
// izmedju otvaranja i klika sme da ZAUSTAVI komandu, ali ne sme tiho da promeni
// manifest koji je korisnik pregledao.
let zbirnaNamera = [];

async function startZbirnaCreation() {
    document.getElementById('zbirnaMainView').style.display = 'none';
    document.getElementById('zbirnaCreateView').style.display = 'block';

    const danas = getTodayIsoDate();
    zbirnaNamera = (vozacOtpremnice || [])
        .filter(r => r.datum === danas && !r.zbirnaID)
        .map(r => r.otpremnicaID);
    
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
    const today = getTodayIsoDate();
    const todayOtkupi = vozacOtpremnice.filter(r => r.datum === today);

    let totalKgI = 0, totalKgII = 0, totalAmb = 0;
    todayOtkupi.forEach(r => {
        totalKgI += otpKgKlase(r, 'I');
        totalKgII += otpKgKlase(r, 'II');
        totalAmb += otpAmb(r);
    });

    const totalKg = totalKgI + totalKgII;
    const stanicaCount = new Set(todayOtkupi.map(r => r.stanicaID)).size;

    const summaryEl = document.getElementById('zbirnaOtkupiSummary');
    if (summaryEl) {
        summaryEl.innerHTML = `<div class="vsum">
            <div class="vsum__total">${totalKg.toLocaleString('sr')}<span class="u"> kg</span></div>
            <div class="vsum__detail">
                ${totalKgII > 0
                    ? 'Kl. I: ' + totalKgI.toLocaleString('sr') + ' kg · Kl. II: ' + totalKgII.toLocaleString('sr') + ' kg'
                    : 'Klasa I'}
                · Amb: ${totalAmb}
            </div>
            <div class="vsum__sub">${todayOtkupi.length} ${todayOtkupi.length === 1 ? 'otpremnica' : 'otpremnice'} · ${stanicaCount} ${stanicaCount === 1 ? 'stanica' : 'stanice'}</div>
        </div>`;
    }

    const countEl = document.getElementById('zbirnaOtkupiCount');
    if (countEl) countEl.textContent = String(todayOtkupi.length);

    // List individual otkupi as vblk cards (read-only, all selected)
    const listEl = document.getElementById('zbirnaOtkupiList');
    if (listEl) {
        listEl.innerHTML = todayOtkupi.map(r =>
            `<div class="vblk sel">
                <div class="vblk__c">✓</div>
                <div>
                    <div class="vblk__nm">${escapeHtml(r.brojOtpremnice || r.otpremnicaID)}</div>
                    <div class="vblk__s">${escapeHtml(r.vrstaVoca)} ${escapeHtml(otpKlaseOpis(r))} · ${escapeHtml(fmtStanica(r.stanicaID))}</div>
                </div>
                <div class="vblk__kg">${otpKg(r).toLocaleString('sr')}<span class="u"> kg</span></div>
            </div>`
        ).join('');
    }
}

function mapServerZbirnaRecord(r) {
    return {
        clientRecordID: r.ClientRecordID || '',
        serverRecordID: r.ServerRecordID || '',
        brojZbirne: r.BrojZbirne || '',
        createdAtClient: normalizeIso(r.CreatedAtClient),
        updatedAtClient: normalizeIso(r.UpdatedAtClient || r.CreatedAtClient),
        updatedAtServer: normalizeIso(r.UpdatedAtServer || r.ReceivedAt),
        syncedAt: normalizeIso(r.UpdatedAtServer || r.ReceivedAt),

        datum: fmtDate(r.Datum),
        kupacID: r.KupacID || '',
        kupacName: r.KupacName || r.KupacID || '',
        vrstaVoca: r.VrstaVoca || '',
        sortaVoca: r.SortaVoca || '',
        kolicinaKlI: parseFloat(r.KolicinaKlI) || 0,
        kolicinaKlII: parseFloat(r.KolicinaKlII) || 0,
        kolAmbalaze: parseInt(r.KolAmbalaze, 10) || 0,
        tipAmbalaze: r.TipAmbalaze || '',
        klasa: r.Klasa || '',
        otpremnicaIDs: r.OtpremnicaIDs || '',

        // VERDIKT MASTERA, DOSLOVNO (review #392, P1).
        //
        // Do sada se gubio: mapper je upisivao fiksno 'synced', sto znaci samo
        // "GAS je primio red". Master svoj ishod pise NAZAD u VOZ list
        // (Synced>Master / Duplicate / SyncError:...), pa je to jedini signal po
        // kom klijent zna da je dokument stvarno razresen.
        masterStatus: r.SyncStatus || '',

        syncStatus: 'synced',
        syncAttempts: 0,
        lastSyncError: '',
        lastServerStatus: 'server'
    };
}

function normalizeLocalZbirnaRecord(r) {
    return {
        clientRecordID: r.clientRecordID || '',
        serverRecordID: r.serverRecordID || '',
        brojZbirne: r.brojZbirne || '',
        createdAtClient: normalizeIso(r.createdAtClient),
        updatedAtClient: normalizeIso(r.updatedAtClient || r.createdAtClient),
        updatedAtServer: normalizeIso(r.updatedAtServer),
        syncedAt: normalizeIso(r.syncedAt),

        datum: r.datum || '',
        kupacID: r.kupacID || '',
        kupacName: r.kupacName || r.kupacID || '',
        vrstaVoca: r.vrstaVoca || '',
        sortaVoca: r.sortaVoca || '',
        kolicinaKlI: parseFloat(r.kolicinaKlI) || 0,
        kolicinaKlII: parseFloat(r.kolicinaKlII) || 0,
        kolAmbalaze: parseInt(r.kolAmbalaze, 10) || 0,
        tipAmbalaze: r.tipAmbalaze || '',
        klasa: r.klasa || '',
        otpremnicaIDs: r.otpremnicaIDs || '',
        masterStatus: r.masterStatus || '',

        syncStatus: r.syncStatus || 'pending',
        syncAttempts: parseInt(r.syncAttempts, 10) || 0,
        lastSyncError: r.lastSyncError || '',
        lastServerStatus: r.lastServerStatus || ''
    };
}

function mergeZbirneRecords(local, server) {
    return mergeOfflineRecords(local, server, normalizeLocalZbirnaRecord);
}

function renderVozacZbirneFromData(allZbirne) {
    const all = (allZbirne || [])
        .filter(r => !r.deleted)
        .sort((a, b) => {
            const byDate = (b.datum || '').localeCompare(a.datum || '');
            if (byDate !== 0) return byDate;

            const byTime = String(b.updatedAtClient || b.createdAtClient || b.updatedAtServer || '')
                .localeCompare(String(a.updatedAtClient || a.createdAtClient || a.updatedAtServer || ''));
            if (byTime !== 0) return byTime;

            return String(b.clientRecordID || '').localeCompare(String(a.clientRecordID || ''));
        });

    const list = document.getElementById('vozacZbirneList');
    if (!list) return;

    if (all.length === 0) {
        list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:12px;">Nema kreiranih zbirnih</p>';
        return;
    }

    list.innerHTML = all.map(r => {
        const totalKg = (r.kolicinaKlI || 0) + (r.kolicinaKlII || 0);
        const isPending = r.syncStatus === 'pending' || r.syncStatus === 'syncing';
        const trMod = isPending ? 'tr--load' : '';
        const bdgClass = r.syncStatus === 'syncing' ? 'tr__bdg--road' :
                         r.syncStatus === 'pending'  ? 'tr__bdg--pending' :
                                                       'tr__bdg--done';
        const bdgLabel = r.syncStatus === 'syncing' ? 'Sync...' :
                         r.syncStatus === 'pending'  ? 'Na čekanju' :
                                                       'Sinhronizovano';
        const syncIcon = r.syncStatus === 'syncing' ? '🔄' :
                         r.syncStatus === 'pending'  ? '⏳' : '✅';

        const klaseLine = r.kolicinaKlII > 0
            ? 'Kl.I: ' + (r.kolicinaKlI || 0).toLocaleString('sr') + ' · Kl.II: ' + (r.kolicinaKlII || 0).toLocaleString('sr') + ' kg'
            : '';

        return `<div class="tr ${trMod}">
            <div class="tr__top">
                <div>
                    <div class="tr__dest">${escapeHtml(r.kupacName)}</div>
                    ${r.vrstaVoca ? `<div class="tr__from">${escapeHtml(r.vrstaVoca)}</div>` : ''}
                </div>
                <div class="tr__time">${escapeHtml(r.datum)}</div>
            </div>
            <div class="tr__main">
                <div class="tr__kg">${totalKg.toLocaleString('sr')}<span class="u"> kg</span></div>
                <div class="tr__det">
                    Amb: <strong>${r.kolAmbalaze || 0}</strong>
                    ${r.brojZbirne ? ' · ' + escapeHtml(r.brojZbirne) : ''}
                    ${klaseLine ? '<br>' + klaseLine : ''}
                </div>
            </div>
            <div class="tr__bot">
                <span class="tr__bdg ${bdgClass}">${bdgLabel}</span>
                <span class="tr__sync">${syncIcon}</span>
            </div>
            ${r.lastSyncError ? `<div class="tr__err">${escapeHtml(r.lastSyncError)}</div>` : ''}
        </div>`;
    }).join('');
}

function cancelZbirna() {
    document.getElementById('zbirnaCreateView').style.display = 'none';
    document.getElementById('zbirnaMainView').style.display = 'block';
    loadVozacZbirne();
}

async function confirmZbirna() {
    if (typeof withSubmitLock !== 'function') {
        if (typeof window.ensureMasterSyncNotActive === 'function') {
            const allowed = await window.ensureMasterSyncNotActive('confirmZbirna', {
                showToast: true
            });

            if (!allowed) return;
        }

        return confirmZbirnaUnlocked();
    }

    return withSubmitLock('zbirna:confirm', confirmZbirnaUnlocked, {
        action: 'confirm-zbirna',
        reason: 'confirmZbirna',
        alreadyMessage: 'Kreiranje zbirne je već u toku'
    });
}

async function confirmZbirnaUnlocked() {
    const kupacSel = document.getElementById('fldZbirnaKupac');
    if (!kupacSel || !kupacSel.value) {
        showToast('Izaberite kupca', 'error');
        return;
    }

    const kupacID = kupacSel.value;
    const today = getTodayIsoDate();

    // KOMANDNA KAPIJA (review #392, P1).
    //
    // Drugi uredjaj je mogao da potrosi iste otpremnice dok je ovaj ekran stajao
    // otvoren. Lock sam po sebi to ne hvata: master moze da zavrsi ciklus, lock
    // padne, a ovaj uredjaj i dalje drzi stari snimak u memoriji.
    //
    // Zato se pre upisa tazi SVEZ authoritative spisak. Endpoint iz #391 vec
    // garantuje da je snimak iz jedne objavljene generacije, pa se ovde meri samo
    // da li je NAMERA korisnika jos izvodljiva.
    //
    // OFFLINE ostaje offline-first: bez veze se radi nad poslednjim poznatim.
    let slobodneNaServeru = null;

    if (navigator.onLine) {
        const sveze = await safeAsync(async () => {
            return await apiFetch('action=getVozacOtpremnice');
        }, 'Greška pri proveri stanja vožnji');

        if (!sveze || sveze.success !== true || sveze.readModelChanging) {
            showToast(
                (sveze && sveze.message) ||
                'Stanje vožnji se ne može potvrditi - probaj ponovo',
                'error'
            );
            return;
        }

        slobodneNaServeru = new Set(
            (Array.isArray(sveze.records) ? sveze.records : [])
                .filter(r => !String((r && r.zbirnaID) || '').trim())
                .map(r => String((r && r.otpremnicaID) || '').trim())
        );
    }

    // DRUGA SVEZA KAPIJA: LOKALNE NERAZRESENE REZERVACIJE.
    //
    // Server sam nije dovoljan. Drugi tab (ista baza, isti uredjaj) mogao je
    // upravo da napravi zbirnu nad istim otpremnicama; master je jos ne vidi, pa
    // server sasvim tacno kaze "slobodna". withSubmitLock to ne hvata -- brava
    // zivi samo u memoriji OVOG taba.
    let rezervisaneSada;
    try {
        rezervisaneSada = rezervisaneOtpremnice(await getMergedZbirneForVozac());
    } catch (err) {
        console.error('confirmZbirna rezervacije failed:', err);
        showToast('Stanje zbirnih se ne može pročitati - probaj ponovo', 'error');
        return;
    }

    // TACNO ONAJ SKUP KOJI JE KORISNIK VIDEO.
    //
    // Ne uzima se "sve sto je sada slobodno": to bi tiho promenilo manifest koji
    // je upravo pregledan. Ako je ijedna otpremnica otisla -- serveru ili drugoj
    // lokalnoj zbirni -- komanda STAJE i ekran se precrtava.
    const izgubljene = (zbirnaNamera || []).filter(id =>
        (slobodneNaServeru && !slobodneNaServeru.has(id)) || rezervisaneSada.has(id)
    );

    if (!zbirnaNamera.length || izgubljene.length) {
        showToast(
            izgubljene.length === 1
                ? 'Jedna otpremnica je u međuvremenu već u zbirnoj - proveri spisak'
                : izgubljene.length + ' otpremnica je u međuvremenu već u zbirnoj - proveri spisak',
            'error'
        );
        await loadVozacData();
        cancelZbirna();
        return;
    }

    // NAMERA SE RAZRESAVA TACNO, ILI SE NE RAZRESAVA.
    //
    // Filter bi tiho napravio MANJI manifest od onog koji je korisnik pregledao:
    // ako lokalni spisak u medjuvremenu vise ne nosi neku otpremnicu, zbirna bi
    // nastala bez nje. Zato se broji -- razresen skup mora biti jednak nameri.
    const todayOtkupi = (zbirnaNamera || [])
        .map(id => (vozacOtpremnice || []).find(r => r.otpremnicaID === id))
        .filter(Boolean);

    if (todayOtkupi.length !== (zbirnaNamera || []).length) {
        showToast('Spisak otpremnica se promenio - proveri pa probaj ponovo', 'error');
        await loadVozacData();
        cancelZbirna();
        return;
    }

    let totalKgI = 0;
    let totalKgII = 0;
    let totalAmb = 0;
    const vrste = new Set();
    const sorte = new Set();

    todayOtkupi.forEach(r => {
        totalKgI += otpKgKlase(r, 'I');
        totalKgII += otpKgKlase(r, 'II');

        totalAmb += otpAmb(r);
        if (r.vrstaVoca) vrste.add(r.vrstaVoca);
        if (r.sortaVoca) sorte.add(r.sortaVoca);
    });

    const kupacName = kupacSel.selectedOptions[0]
        ? kupacSel.selectedOptions[0].textContent
        : kupacID;

    const nowIso = new Date().toISOString();

    const record = {
        clientRecordID: (window.crypto && typeof window.crypto.randomUUID === 'function')
            ? window.crypto.randomUUID()
            : ('zbr-' + Date.now() + '-' + Math.floor(Math.random() * 1000000)),
        serverRecordID: '',

        // BROJ DODELJUJE MASTER (A2: broj je labela, ne identitet).
        //
        // Klijentski racun vozacBroj/ddmmyy-seq je obrisan: redni broj se
        // izvodio iz zbirni koje BAS OVAJ uredjaj zna, pa su dva telefona istog
        // dana mogla smisliti isti broj. Prazno polje uvoz vec razume kao
        // "generisi lokalno" (GetBrojZbirneForIDStrict).
        brojZbirne: '',
        createdAtClient: nowIso,
        updatedAtClient: nowIso,
        updatedAtServer: '',
        syncedAt: '',

        vozacID: CONFIG.ENTITY_ID,
        datum: today,
        kupacID: kupacID,
        kupacName: kupacName,
        vrstaVoca: Array.from(vrste).join(', '),
        sortaVoca: Array.from(sorte).join(', '),
        kolicinaKlI: totalKgI,
        kolicinaKlII: totalKgII,
        tipAmbalaze: todayOtkupi[0].tipAmbalaze || '',
        kolAmbalaze: totalAmb,
        klasa: totalKgII > 0 ? 'I+II' : 'I',

        // IDENTITET PUTUJE (S5-4b-2). Master vise ne prevodi otkupne CRID-ove u
        // otpremnice -- dobija ih onakve kakve ih je vozac video.
        otpremnicaIDs: todayOtkupi.map(r => r.otpremnicaID).join(','),

        syncStatus: 'pending',
        syncAttempts: 0,
        syncAttemptAt: '',
        lastSyncError: '',
        lastServerStatus: '',
        deleted: false,
        entityType: 'zbirna',
        schemaVersion: 1
    };

    try {
        await dbPut(db, 'zbirne', record);
    } catch (err) {
        console.error('confirmZbirna dbPut failed:', err);
        showToast('Greška pri čuvanju zbirne', 'error');
        return;
    }

    showToast('Zbirna kreirana!', 'success');
    cancelZbirna();

    if (navigator.onLine && typeof syncQueueSafe === 'function') {
        syncQueueSafe('post-save');
    }

    try {
        await loadVozacData();
    } catch (err) {
        console.error('confirmZbirna loadVozacData failed:', err);
    }
}

async function syncZbirne() {
    const result = await syncStore({
        storeName: 'zbirne',
        action: 'syncZbirna',
        inFlightKey: 'zbirnaInFlight',
        entityIdField: 'vozacID',
        successLabel: 'Zbirna sinhronizovana',
        onResultRecord: (record, serverResult) => {
            // Zbirna-specific: backend returns brojZbirne on success
            if (serverResult.brojZbirne) {
                record.brojZbirne = serverResult.brojZbirne;
            }
        }
    });

    try { await loadVozacZbirne(); } catch (_) {}

    return result;
}

// Kodovi kojima server kaze "presudio sam i odbijam" -- retry ne menja ishod.
//
// Transportni neuspeh nema kod, pa ovde namerno nije nabrojan: takav zapis je
// jos na putu i mora da zadrzi svoje otpremnice.
const ZBIRNA_TRAJNO_ODBIJENA = new Set([
    'ZBIRNA_CONFLICT',
    'VALIDATION_ERROR',
    'CLIENT_RECORD_ID_MISSING'
]);

// Da li je master doneo odluku o ovoj zbirni.
//
// Isti skup koji GAS zove terminalnim: master je red video i presudio, pa od tog
// trenutka o otpremnicama govori KANONSKO stanje, ne vise lokalni dogadjaj.
function zbirnaRazresenaOdMastera(z) {
    const s = String((z && z.masterStatus) || '').trim();
    return s === 'Synced>Master' || s === 'Duplicate' || s.indexOf('SyncError') === 0;
}

// OTPREMNICE KOJE DRZI JOS NERAZRESENA ZBIRNA (review #392, P1).
//
// Server govori kanonsku tekucu istinu (Otpremnica.zbirnaID), ali je zna tek
// POSLE master ciklusa. Izmedju klika i tog ciklusa lokalna zbirna postoji, a
// server jos sasvim tacno kaze "slobodna" -- pa bi ista otpremnica odmah ponovo
// usla u izbor i vozac bi nad njom napravio DRUGU zbirnu. Master bi je kasnije
// odbio, ali komanda bi vec bila izvrsena kao ispravna.
//
// Lokalni 'synced' NIJE razresenje: znaci samo da je GAS primio red. Razresenje
// daje master, i cita se iz njegovog verdikta -- inace bi rezervacija trajala
// zauvek i posle storna zbirne otpremnica se nikad ne bi vratila u izbor.
function rezervisaneOtpremnice(zbirne) {
    const rez = new Set();

    (zbirne || []).forEach(z => {
        if (!z || z.deleted) return;

        // ODBIJEN DOGADJAJ OSLOBADJA, ALI SAMO TRAJNO ODBIJEN.
        //
        // Sync engine i transportni pad i poslovno odbijanje ostavlja kao
        // syncStatus 'pending' -- razlikuje ih tek KOD koji je server vratio.
        // Bez te razlike bi trajno odbijena zbirna zauvek drzala svoje
        // otpremnice, jer svaki retry pada iz istog razloga.
        //
        // Prazan kod = transportni neuspeh: zapis je i dalje na putu, pa
        // rezervacija OSTAJE.
        const lokalni = String(z.syncStatus || '').trim();
        if (lokalni === 'error' || lokalni === 'failed') return;
        if (ZBIRNA_TRAJNO_ODBIJENA.has(String(z.lastServerCode || '').trim())) return;

        if (zbirnaRazresenaOdMastera(z)) return;

        String(z.otpremnicaIDs || '')
            .split(',')
            .map(x => x.trim())
            .filter(Boolean)
            .forEach(id => rez.add(id));
    });

    return rez;
}

async function getMergedZbirneForVozac() {
    let local = [];
    let server = [];

    try {
        local = await dbGetAll(db, 'zbirne');
    } catch (err) {
        console.error('getMergedZbirneForVozac local failed:', err);
    }

    const json = await safeAsync(async () => {
        return await apiFetch('action=getVozacZbirne');
    }, 'Greška pri učitavanju zbirnih');

    if (json && json.success && Array.isArray(json.records)) {
        server = json.records.map(mapServerZbirnaRecord);
    }

    return dedupeRecordsForRender(mergeZbirneRecords(local, server))
    .filter(r => !r.deleted);
}

