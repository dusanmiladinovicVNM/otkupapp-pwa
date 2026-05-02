// ============================================================
// KOOPERANT: PREGLED
// Agregira postojeće module:
// - kartica.js
// - knjiga-polja.js
// - agromere.js
// - parcele.js
// - koopinfo.js
// - kooperant sync
// ============================================================

let pregledCache = null;
let pregledCacheTs = 0;
const PREGLED_CACHE_TTL = 30000; // 30s

async function loadPregled(forceRefresh) {
    const root = document.getElementById('tab-home');
    if (!root) return;

    const nameEl = document.getElementById('homeName');
    const idEl = document.getElementById('homeKoopId');

    if (nameEl) nameEl.textContent = CONFIG.ENTITY_NAME || 'Gazdinstvo';
    if (idEl) idEl.textContent = CONFIG.ENTITY_ID || '-';

    const now = Date.now();
    if (
        !forceRefresh &&
        pregledCache &&
        (now - pregledCacheTs < PREGLED_CACHE_TTL)
    ) {
        renderPregled(pregledCache);
        return;
    }

    renderPregledLoading();

    const data = await buildPregledData();
    pregledCache = data;
    pregledCacheTs = Date.now();

    renderPregled(data);
}

function invalidatePregledCache() {
    pregledCache = null;
    pregledCacheTs = 0;
}

async function buildPregledData() {
    const data = {
        hero: {
            entityName: CONFIG.ENTITY_NAME || 'Gazdinstvo',
            entityID: CONFIG.ENTITY_ID || '-'
        },
        kpi: {
            todayRadovi: 0,
            activeAlerts: 0,
            saldo: 0,
            parcele: 0
        },
        alerts: [],
        bilans: {
            proizvodnja: 0,
            agrohemija: 0,
            troskovi: 0,
            rezultat: 0,
            radniSati: 0
        },
        kartica: {
            zaduzenje: 0,
            razduzenje: 0,
            saldo: 0
        },
        info: {
            otkupAktivan: '-',
            radnoVreme: '-',
            sezona: '-'
        }
    };

    // --------------------------------------------------------
    // KOOP INFO
    // --------------------------------------------------------
    try {
        const config = stammdaten.config || [];
        const gv = (k) => {
            const c = config.find(x => x.Parameter === k);
            return c ? c.Vrednost : '-';
        };

        data.info.otkupAktivan = gv('OtkupAktivan');
        data.info.radnoVreme = `${gv('RadnoVremeOd')} - ${gv('RadnoVremeDo')}`;
        data.info.sezona = `${gv('SezonaOd')} - ${gv('SezonaDo')}`;
    } catch (e) {
        console.error('Pregled info build failed:', e);
    }

    // --------------------------------------------------------
    // PARCELE + METEO ALERTS
    // --------------------------------------------------------
    try {
        const parcele = (stammdaten.parcele || []).filter(p => p.KooperantID === CONFIG.ENTITY_ID);
        data.kpi.parcele = parcele.length;

        const meteoAlerts = [];

        parcele.forEach(p => {
            const pid = String(p.ParcelaID || '').trim();
            const katBroj = p.KatBroj || p.ParcelaID || pid;
            const kultura = p.Kultura || '';

            const meteo = (window.meteoCache && window.meteoCache[pid]) ? window.meteoCache[pid] : null;
            if (!meteo) return;

            const risk = meteo.risk || {};
            const items = Array.isArray(risk.items) ? risk.items : [];
            if (items.length > 0) {
                const first = items[0];
                meteoAlerts.push({
                    type: first.level === 'danger' ? 'danger' : 'warning',
                    icon: first.icon || '⚠️',
                    title: `${katBroj}${kultura ? ' — ' + kultura : ''}`,
                    text: first.message || 'Aktivno upozorenje na parceli',
                    action: 'parcele',
                    parcelaID: pid
                });
            } else {
                const spray = Array.isArray(meteo.sprayWindow) ? meteo.sprayWindow : [];
                if (!spray.length) {
                    meteoAlerts.push({
                        type: 'warning',
                        icon: '🌦️',
                        title: `${katBroj}${kultura ? ' — ' + kultura : ''}`,
                        text: 'Nema povoljnog termina za prskanje u narednom periodu.',
                        action: 'parcele',
                        parcelaID: pid
                    });
                }
            }
        });

        data.alerts.push(...meteoAlerts.slice(0, 3));
    } catch (e) {
        console.error('Pregled parcele/alerts build failed:', e);
    }

    // --------------------------------------------------------
    // RADOVI / TRETMANI
    // --------------------------------------------------------
    try {
        let tretmani = [];
        if (typeof getTretmaniCached === 'function') {
            tretmani = await getTretmaniCached(false) || [];
        }

        const todayIso = getTodayIsoDate();
        const todayRadovi = tretmani.filter(t => {
            const d = toIsoDateOnly(t.datum || t.Datum || '');
            return d === todayIso && !t.deleted;
        });

        data.kpi.todayRadovi = todayRadovi.length;

        const kasneRadove = tretmani.filter(t => {
            const mera = String(t.mera || t.Mera || '');
            const d = toIsoDateOnly(t.datum || t.Datum || '');
            return !!mera && d < todayIso && !t.deleted;
        });

        if (todayRadovi.length === 0) {
            const kulturaCounts = {};
            const parcele = (stammdaten.parcele || []).filter(p => p.KooperantID === CONFIG.ENTITY_ID);

            parcele.forEach(p => {
                const k = String(p.Kultura || '').trim();
                if (!k) return;
                kulturaCounts[k] = (kulturaCounts[k] || 0) + 1;
            });

            const dominantnaKultura = Object.keys(kulturaCounts)
                .sort((a, b) => kulturaCounts[b] - kulturaCounts[a])[0] || '';

            const dominantnaParcela = parcele.find(p => String(p.Kultura || '').trim() === dominantnaKultura);
            const predlog = getDynamicRadoviPredlog(
                dominantnaKultura,
                dominantnaParcela ? dominantnaParcela.ParcelaID : ''
            );

            data.alerts.push({
                type: 'ok',
                icon: predlog.icon,
                title: predlog.title,
                text: predlog.text,
                items: predlog.items,
                action: 'agromere'
            });
        }
    } catch (e) {
        console.error('Pregled tretmani build failed:', e);
    }

    // --------------------------------------------------------
    // KARTICA
    // --------------------------------------------------------
    try {
        let records = [];

        if (Array.isArray(karticaCache)) {
            records = karticaCache;
        } else {
            const json = await safeAsync(async () => {
                return await apiFetch('action=getKartica&kooperantID=' + encodeURIComponent(CONFIG.ENTITY_ID));
            });

            if (json && json.success && Array.isArray(json.records)) {
                records = json.records.filter(r => r.Opis !== 'UKUPNO');
            }
        }

        let zad = 0;
        let raz = 0;
        let saldo = 0;

        records.forEach(r => {
            const z = parseFloat(r.Zaduzenje) || 0;
            const ra = parseFloat(r.Razduzenje) || 0;
            const s = parseFloat(r.Saldo) || 0;
            zad += z;
            raz += ra;
            saldo = s;
        });

        data.kartica.zaduzenje = zad;
        data.kartica.razduzenje = raz;
        data.kartica.saldo = saldo;
        data.kpi.saldo = saldo;
    } catch (e) {
        console.error('Pregled kartica build failed:', e);
    }

    // --------------------------------------------------------
    // KNJIGA POLJA / BILANS
    // --------------------------------------------------------
    try {
        if (!_kpLoaded && typeof loadKnjigaPolja === 'function') {
            await loadKnjigaPolja();
        } else if (typeof kpFetchAll === 'function' && (!kpData || !Array.isArray(kpData.proizvodnja))) {
            await kpFetchAll();
        }

        const proizvodnja = Array.isArray(kpData?.proizvodnja) ? kpData.proizvodnja : [];
        const tretmani = Array.isArray(kpData?.tretmani) ? kpData.tretmani : [];
        const troskovi = Array.isArray(kpData?.troskovi) ? kpData.troskovi : [];
        const lager = Array.isArray(kpData?.lager) ? kpData.lager : [];

        const proizvodnjaTotal = proizvodnja.reduce((s, r) => {
            return s + (parseFloat(r.Vrednost) || ((parseFloat(r.Kolicina) || 0) * (parseFloat(r.Cena) || 0)));
        }, 0);

        let agrohemijaTotal = 0;
        tretmani.forEach(t => {
            const artID = t.ArtikalID || t.artikalID || '';
            const kol = parseFloat(t.KolicinaUpotrebljena || t.kolicinaUpotrebljena) || 0;
            if (!artID || kol <= 0) return;

            const lagerItem = lager.find(l => l.ArtikalID === artID);
            const wac = lagerItem ? (parseFloat(lagerItem.WAC) || parseFloat(lagerItem.CenaPoJedinici) || 0) : 0;
            agrohemijaTotal += kol * wac;
        });

        const troskoviTotal = troskovi.reduce((s, r) => s + (parseFloat(r.Iznos || r.iznos) || 0), 0);

        let radniMinuti = 0;
        tretmani.forEach(t => {
            radniMinuti += parseInt(t.TrajanjeMinuta || t.trajanjeMinuta) || 0;
        });

        const radniSati = radniMinuti / 60;
        const rezultat = proizvodnjaTotal - agrohemijaTotal - troskoviTotal;

        data.bilans.proizvodnja = proizvodnjaTotal;
        data.bilans.agrohemija = agrohemijaTotal;
        data.bilans.troskovi = troskoviTotal;
        data.bilans.rezultat = rezultat;
        data.bilans.radniSati = radniSati;
    } catch (e) {
        console.error('Pregled bilans build failed:', e);
    }

    // --------------------------------------------------------
    // SYNC ALERTS
    // --------------------------------------------------------
    try {
        if (db) {
            let pendingTretmani = [];
            let pendingAgromere = [];

            try {
                pendingTretmani = await dbGetByIndex(db, 'tretmani', 'syncStatus', 'pending');
            } catch (_) {}

            try {
                if (CONFIG && CONFIG.AGRO_STORE) {
                    pendingAgromere = await dbGetByIndex(db, CONFIG.AGRO_STORE, 'syncStatus', 'pending');
                }
            } catch (_) {}

            const totalPending = (pendingTretmani?.length || 0) + (pendingAgromere?.length || 0);
            if (totalPending > 0) {
                data.alerts.push({
                    type: 'warning',
                    icon: '🔄',
                    title: 'Sinhronizacija',
                    text: `${totalPending} stavki čeka sinhronizaciju.`,
                    action: 'sync'
                });
            } else {
                data.alerts.push({
                    type: 'ok',
                    icon: '✅',
                    title: 'Sync status',
                    text: 'Svi podaci su sinhronizovani.',
                    action: 'sync'
                });
            }
        }
    } catch (e) {
        console.error('Pregled sync build failed:', e);
    }

    // --------------------------------------------------------
    // OTKUP STATUS ALERT
    // --------------------------------------------------------
    try {
        if (data.info.otkupAktivan && data.info.otkupAktivan !== 'Da') {
            data.alerts.push({
                type: 'danger',
                icon: '⛔',
                title: 'Otkup',
                text: 'Otkup trenutno nije aktivan.',
                action: 'koopinfo'
            });
        }
    } catch (e) {
        console.error('Pregled otkup status build failed:', e);
    }

    // dedupe + count
    data.alerts = dedupePregledAlerts(data.alerts).slice(0, 5);
    data.kpi.activeAlerts = data.alerts.filter(a => a.type !== 'ok').length;

    return data;
}

function dedupePregledAlerts(alerts) {
    const seen = new Set();
    const out = [];

    (alerts || []).forEach(a => {
        const key = [a.type, a.title, a.text, a.action, a.parcelaID].join('|');
        if (seen.has(key)) return;
        seen.add(key);
        out.push(a);
    });

    return out;
}

function renderPregledLoading() {
    const alertsEl = document.getElementById('homeAlertsList');
    if (alertsEl) {
        alertsEl.innerHTML = '<div class="koop-loading">Učitavanje pregleda...</div>';
    }
}

function renderPregled(data) {
    setTextIfExists('homeName', data.hero.entityName);
    setTextIfExists('homeKoopId', data.hero.entityID);

    setTextIfExists('homeKpiRadovi', toSrNum(data.kpi.todayRadovi));
    setTextIfExists('homeKpiAlerts', toSrNum(data.kpi.activeAlerts));
    setTextIfExists('homeKpiSaldo', toSrNum(data.kpi.saldo));

    setTextIfExists('homeBilansProizvodnja', toSrNum(data.bilans.proizvodnja));
    setTextIfExists('homeBilansTroskovi', toSrNum(data.bilans.troskovi));
    setTextIfExists('homeBilansAgro', toSrNum(data.bilans.agrohemija));
    setTextIfExists('homeBilansRezultat', toSrNum(data.bilans.rezultat));

    setTextIfExists('homeKarticaZad', toSrNum(data.kartica.zaduzenje));
    setTextIfExists('homeKarticaRaz', toSrNum(data.kartica.razduzenje));
    setTextIfExists('homeKarticaSaldo', toSrNum(data.kartica.saldo));

    const alertsEl = document.getElementById('homeAlertsList');
    if (alertsEl) {
        if (!data.alerts.length) {
            alertsEl.innerHTML = `
                <div class="koop-empty">
                    <div class="koop-empty-title">Nema aktivnih upozorenja</div>
                    <div class="koop-empty-text">Trenutno nema meteo, sync ni operativnih upozorenja.</div>
                </div>
            `;
        } else {
            alertsEl.innerHTML = data.alerts.map((a, idx) => `
                <button class="home-alert-item is-${escapeHtml(a.type || 'ok')}" data-action="pregled-alert-click" data-index="${idx}" type="button">
                    <div class="home-alert-icon">${escapeHtml(a.icon || '•')}</div>
                    <div>
                        <div class="home-alert-title">${escapeHtml(a.title || '')}</div>
                        <div class="home-alert-text">${escapeHtml(a.text || '')}</div>
                        ${Array.isArray(a.items) && a.items.length ? `
                            <div class="home-alert-list">
                                ${a.items.map(item => `<div class="home-alert-list-item">• ${escapeHtml(item)}</div>`).join('')}
                        </div>
                    ` : ''}
                </div>
            </button>
        `).join('');
        }
    }

    window._pregledLastData = data;
}

function onPregledAlertClick(index) {
    const data = window._pregledLastData;
    if (!data || !Array.isArray(data.alerts)) return;

    const alert = data.alerts[index];
    if (!alert) return;

    if (alert.action === 'parcele') {
        showTab('parcele', findTabBtnByTabName('parcele'));
        if (alert.parcelaID) {
            setTimeout(() => {
                try { focusParcel(alert.parcelaID); } catch (_) {}
            }, 250);
        }
        return;
    }

    if (alert.action === 'agromere') {
        showTab('agromere', findTabBtnByTabName('agromere'));
        return;
    }

    if (alert.action === 'koopinfo') {
        showTab('koopinfo', findTabBtnByTabName('koopinfo'));
        return;
    }

    if (alert.action === 'sync') {
        showToast('Otvorite Radovi ili pokrenite sync kada bude dostupan u Pregledu', 'info');
        return;
    }
}

function openHomeQuickActions() {
    const modal = document.getElementById('homeQuickActionsModal');
    if (modal) modal.classList.add('visible');
}

function closeHomeQuickActions() {
    const modal = document.getElementById('homeQuickActionsModal');
    if (modal) modal.classList.remove('visible');
}

function goToNewRad() {
    showTab('agromere', findTabBtnByTabName('agromere'));
}

function goToNewTrosak() {
    showTab('knjigapolja', findTabBtnByTabName('knjigapolja'));
}

function goToScanRacun() {
    showTab('knjigapolja', findTabBtnByTabName('knjigapolja'));

    setTimeout(() => {
        if (typeof showKnjigaSection === 'function') {
            const lagerBtn = Array.from(document.querySelectorAll('.knjiga-subnav-btn'))
                .find(btn => btn.textContent.trim().toLowerCase().includes('lager'));
            showKnjigaSection('lager', lagerBtn || null);
        }

        if (typeof startFiskalniScan === 'function') {
            startFiskalniScan();
        }
    }, 250);
}

function goToKartica() {
    showTab('kartica', findTabBtnByTabName('kartica'));
}

function goToKnjigaPolja() {
    showTab('knjigapolja', findTabBtnByTabName('knjigapolja'));
}

function showHomeAlerts() {
    const data = window._pregledLastData;
    if (!data || !data.alerts || !data.alerts.length) {
        showToast('Nema aktivnih upozorenja', 'info');
        return;
    }
    showToast('Dodirnite pojedinačno upozorenje za otvaranje detalja', 'info');
}

// ------------------------------------------------------------
// Helpers
// ------------------------------------------------------------
function setTextIfExists(id, value) {
    const el = document.getElementById(id);
    if (el) el.textContent = value;
}

function toSrNum(value) {
    const n = Number(value || 0);
    return n.toLocaleString('sr-RS');
}

function findTabBtnByTabName(tabName) {
    return document.querySelector('.tab-btn[data-route="tab"][data-tab="' + tabName + '"]');
}

function getDynamicRadoviPredlog(kultura, parcelaId) {
    const k = String(kultura || '').trim().toLowerCase();
    const mesec = new Date().getMonth() + 1;

    const meteo = (window.meteoCache && parcelaId) ? window.meteoCache[parcelaId] : null;
    const riskItems = meteo && meteo.risk && Array.isArray(meteo.risk.items) ? meteo.risk.items : [];
    const sprayWindows = meteo && Array.isArray(meteo.sprayWindow) ? meteo.sprayWindow : [];

    const imaRizik = riskItems.length > 0;
    const imaDanger = riskItems.some(r => r.level === 'danger');
    const nemaSprayProzora = !sprayWindows.length;

    let base = {
        icon: '📅',
        title: 'Planirani radovi',
        text: 'Dostupan je automatski predlog sledećih radova.',
        items: ['Pregled parcele', 'Planiranje narednog rada']
    };

    // --------------------------------------------------------
    // JABUKA
    // --------------------------------------------------------
    if (k === 'jabuka') {
        if (mesec <= 2) {
            base = {
                icon: '🍎',
                title: 'Planirani radovi',
                text: 'Za jabuku je aktuelan zimski plan radova.',
                items: ['Rezidba', 'Pregled zasada', 'Plan zaštite']
            };
        } else if (mesec <= 4) {
            base = {
                icon: '🍎',
                title: 'Planirani radovi',
                text: 'Za jabuku su aktuelni prolećni radovi.',
                items: ['Zaštita zasada', 'Prihrana', 'Praćenje cvetanja i zametanja']
            };
        } else if (mesec <= 6) {
            base = {
                icon: '🍎',
                title: 'Planirani radovi',
                text: 'Za jabuku su aktuelni radovi intenzivnog porasta.',
                items: ['Zaštita zasada', 'Folijarna prihrana', 'Praćenje zdravstvenog stanja plodova']
            };
        } else if (mesec <= 8) {
            base = {
                icon: '🍎',
                title: 'Planirani radovi',
                text: 'Za jabuku je aktuelna letnja operativa.',
                items: ['Zaštita zasada', 'Navodnjavanje po potrebi', 'Priprema za berbu ranijih sorti']
            };
        } else if (mesec <= 10) {
            base = {
                icon: '🍎',
                title: 'Planirani radovi',
                text: 'Za jabuku je aktuelan period berbe i završnih radova.',
                items: ['Berba', 'Evidencija prinosa', 'Planiranje postberbenih radova']
            };
        } else {
            base = {
                icon: '🍎',
                title: 'Planirani radovi',
                text: 'Za jabuku su aktuelni jesenji i pripremni radovi.',
                items: ['Sanitarni pregled zasada', 'Jesenja prihrana po potrebi', 'Priprema za mirovanje']
            };
        }
    }

    // --------------------------------------------------------
    // VIŠNJA
    // --------------------------------------------------------
    if (k === 'visnja' || k === 'višnja') {
        if (mesec <= 2) {
            base = {
                icon: '🍒',
                title: 'Planirani radovi',
                text: 'Za višnju je aktuelan zimski plan radova.',
                items: ['Rezidba', 'Pregled zasada', 'Priprema zaštite za sezonu']
            };
        } else if (mesec <= 4) {
            base = {
                icon: '🍒',
                title: 'Planirani radovi',
                text: 'Za višnju su aktuelni prolećni radovi.',
                items: ['Zaštita zasada', 'Praćenje cvetanja', 'Prihrana po potrebi']
            };
        } else if (mesec <= 6) {
            base = {
                icon: '🍒',
                title: 'Planirani radovi',
                text: 'Za višnju je aktuelna priprema i organizacija berbe.',
                items: ['Zaštita pred berbu', 'Praćenje zrenja', 'Priprema berbe']
            };
        } else if (mesec === 7) {
            base = {
                icon: '🍒',
                title: 'Planirani radovi',
                text: 'Za višnju je aktuelan period berbe i završetka sezone.',
                items: ['Berba', 'Evidencija prinosa', 'Postberbeni pregled zasada']
            };
        } else if (mesec <= 9) {
            base = {
                icon: '🍒',
                title: 'Planirani radovi',
                text: 'Za višnju su aktuelni postberbeni radovi.',
                items: ['Postberbena zaštita', 'Dopunska prihrana', 'Sanitarni pregled zasada']
            };
        } else {
            base = {
                icon: '🍒',
                title: 'Planirani radovi',
                text: 'Za višnju su aktuelni jesenji i pripremni radovi.',
                items: ['Priprema za mirovanje', 'Pregled stabala', 'Plan rezidbe']
            };
        }
    }

    // --------------------------------------------------------
    // ŠLJIVA
    // --------------------------------------------------------
    if (k === 'sljiva' || k === 'šljiva') {
        if (mesec <= 2) {
            base = {
                icon: '🟣',
                title: 'Planirani radovi',
                text: 'Za šljivu je aktuelan zimski plan radova.',
                items: ['Rezidba', 'Pregled stabala', 'Plan zaštite']
            };
        } else if (mesec <= 4) {
            base = {
                icon: '🟣',
                title: 'Planirani radovi',
                text: 'Za šljivu su aktuelni prolećni radovi.',
                items: ['Zaštita zasada', 'Prihrana', 'Praćenje cvetanja i zametanja']
            };
        } else if (mesec <= 6) {
            base = {
                icon: '🟣',
                title: 'Planirani radovi',
                text: 'Za šljivu su aktuelni radovi razvoja ploda.',
                items: ['Zaštita zasada', 'Praćenje zdravstvenog stanja', 'Navodnjavanje po potrebi']
            };
        } else if (mesec <= 8) {
            base = {
                icon: '🟣',
                title: 'Planirani radovi',
                text: 'Za šljivu je aktuelan period pripreme i organizacije berbe.',
                items: ['Praćenje zrenja', 'Zaštita pred berbu', 'Planiranje berbe']
            };
        } else if (mesec <= 10) {
            base = {
                icon: '🟣',
                title: 'Planirani radovi',
                text: 'Za šljivu je aktuelna berba i završetak sezone.',
                items: ['Berba', 'Evidencija prinosa', 'Postberbeni pregled']
            };
        } else {
            base = {
                icon: '🟣',
                title: 'Planirani radovi',
                text: 'Za šljivu su aktuelni jesenji radovi i priprema zasada.',
                items: ['Sanitarni pregled', 'Jesenja prihrana po potrebi', 'Priprema za mirovanje']
            };
        }
    }

    // --------------------------------------------------------
    // METEO KOREKCIJA
    // --------------------------------------------------------
    if (imaDanger) {
        return {
            icon: '⛔',
            title: 'Planirani radovi',
            text: 'Meteo rizik je trenutno visok — preporučuje se odlaganje tretmana i fokus na pregled parcele.',
            items: [
                'Pregled parcele',
                'Praćenje upozorenja',
                'Sačekati povoljniji termin'
            ]
        };
    }

    if (imaRizik) {
        return {
            icon: '⚠️',
            title: base.title,
            text: base.text + ' Prisutan je meteo rizik, pa preporuke treba prilagoditi uslovima.',
            items: [
                base.items[0],
                'Praćenje meteo upozorenja',
                'Provera bezbednog termina za rad'
            ]
        };
    }

    if (nemaSprayProzora && base.items.some(i => i.toLowerCase().includes('zaštita'))) {
        return {
            icon: '🌦️',
            title: base.title,
            text: base.text + ' Trenutno nema dobrog prozora za prskanje.',
            items: [
                'Pregled zasada',
                'Priprema opreme',
                'Sačekati povoljan termin za zaštitu'
            ]
        };
    }

    return base;
}
