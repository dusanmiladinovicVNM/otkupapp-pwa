// ============================================================
// MANAGEMENT SHELL V2
// ============================================================

window.mgmtShellState = {
    activeRoot: 'pregled',
    partnerSegment: 'kooperanti',
    koopSub: 'pregled',
    kupSub: 'fakture',
    otkupSub: 'otkupi',
    agroSub: 'izdavanje',
    dashboardPeriod: '7d',
    mounted: false
};

// ============================================================
// MGMT DATA PREFETCH + BOOT HELPERS
// moved from legacy mgmt-shell.js
// ============================================================

async function prefetchMgmtData() {
    try {
        const json = await apiFetch('action=getMgmtOverview');

        if (json && json.success) {
            window.mgmtData = Object.assign({}, window.mgmtData || {}, json);
            mgmtData = window.mgmtData;
        }
    } catch (e) {
        console.error('prefetchMgmtData failed:', e);
    }
}

const mgmtSectionPromises = {};

async function ensureMgmtSection(section, action, params) {
    window.mgmtData = window.mgmtData || {};
    mgmtData = window.mgmtData;

    if (Object.prototype.hasOwnProperty.call(window.mgmtData, section)) {
        return Array.isArray(window.mgmtData[section])
            ? window.mgmtData[section]
            : [];
    }

    if (!mgmtSectionPromises[section]) {
        const query = 'action=' + encodeURIComponent(action) + (params ? '&' + params : '');

        mgmtSectionPromises[section] = apiFetch(query)
            .then(json => {
                if (json && json.success) {
                    window.mgmtData = Object.assign({}, window.mgmtData || {}, json);
                    mgmtData = window.mgmtData;
                }

                return Array.isArray(window.mgmtData[section])
                    ? window.mgmtData[section]
                    : [];
            })
            .catch(err => {
                console.error('ensureMgmtSection failed:', section, err);
                return [];
            })
            .finally(() => {
                delete mgmtSectionPromises[section];
            });
    }

    return mgmtSectionPromises[section];
}

window.ensureMgmtSection = ensureMgmtSection;

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

function mgmtLogError(scope, error) {
    console.error(`[mgmt-shell-v2] ${scope} failed:`, error);
}

function mgmtGetDataSlices() {
    const data = window.mgmtData || {};

    return {
        otkupiAll: Array.isArray(data.otkupiAll) ? data.otkupiAll : [],
        saldoKupci: Array.isArray(data.saldoKupci) ? data.saldoKupci : [],
        saldoOM: Array.isArray(data.saldoOM) ? data.saldoOM : [],
        kartice: Array.isArray(data.kartice) ? data.kartice : []
    };
}

function mgmtSetText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
}

function mgmtActivateButton(btn, fallbackSelector) {
    if (btn && btn.classList) {
        btn.classList.add('active');
        return;
    }

    const autoBtn = document.querySelector(fallbackSelector);
    if (autoBtn) autoBtn.classList.add('active');
}

function mgmtClearActive(selector) {
    document.querySelectorAll(selector).forEach(el => el.classList.remove('active'));
}

function mgmtHidePanels(ids) {
    ids.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.classList.remove('active');
    });
}

const MGMT_PARTNER_PANEL_IDS = [
    'mgmt-koop-pregled',
    'mgmt-koop-kartica',
    'mgmt-koop-saldo',
    'mgmt-kup-fakture',
    'mgmt-kup-saldo',
    'mgmt-kup-roba'
];

const MGMT_DASH_PERIOD_COPY = {
    today: {
        kgLabel: 'Otkup · danas',
        kgSub: 'kg danas',
        koopSub: 'jedinstveni danas',
        avgSub: 'ponderisana RSD/kg · danas'
    },
    season: {
        kgLabel: 'Otkup · sezona',
        kgSub: 'kg u sezoni',
        koopSub: 'jedinstveni u sezoni',
        avgSub: 'ponderisana RSD/kg · sezona'
    },
    '7d': {
        kgLabel: 'Otkup · 7 dana',
        kgSub: 'kg u poslednjih 7 dana',
        koopSub: 'jedinstveni u 7 dana',
        avgSub: 'ponderisana RSD/kg · 7 dana'
    }
};

async function mgmtShellInit() {
    if (typeof loadDispecer === 'function') {
        try {
            await loadDispecer();
        } catch (e) {
            mgmtLogError('loadDispecer in mgmtShellInit', e);
        }
    }

    await showMgmtRoot('pregled');

    if (typeof updateRoleNavVisibility === 'function') updateRoleNavVisibility();
    if (typeof updateRoleNavActive === 'function') updateRoleNavActive();
}

async function showMgmtRoot(root, btn) {
    window.mgmtShellState.activeRoot = root;

    // Sync mobile top bar title + drawer active state
    if (typeof mgmtMobileUpdateTbar === 'function') mgmtMobileUpdateTbar(root);

    document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));

    const rootMap = {
        dashboard:  'tab-mgmt-dashboard',
        pregled:    'tab-mgmt-pregled',
        dispecer:   'tab-mgmt-dispecer',
        uzivo:      'tab-mgmt-uzivo',
        otkup:      'tab-mgmt-otkup',
        partneri:   'tab-mgmt-partneri',
        kooperanti: 'tab-mgmt-partneri',
        kupci:      'tab-mgmt-partneri',
        fakture:    'tab-mgmt-partneri',
        izvodi:     'tab-mgmt-izvodi',
        agro:       'tab-mgmt-agro'
    };

    const targetId = rootMap[root];
    const target = document.getElementById(targetId);
    if (target) target.classList.add('active');

    document.querySelectorAll('.tab-btn.role-management').forEach(el => el.classList.remove('active'));
    if (btn && btn.classList) {
        btn.classList.add('active');
    } else {
        const autoBtn = document.querySelector('.tab-btn.role-management[data-root="' + root + '"]');
        if (autoBtn) autoBtn.classList.add('active');
    }

    if (root === 'dashboard') {
        await mgmtRenderDashboard();
    } else if (root === 'pregled') {
        mgmtRenderOverview();
    } else if (root === 'dispecer') {
        if (typeof loadDispecer === 'function') {
            try {
                await loadDispecer();
            } catch (e) {
                mgmtLogError('showMgmtRoot', e);
            }
        }
        mgmtRenderOverview();
    } else if (root === 'uzivo') {
        if (typeof loadMgmtOtkupUzivo === 'function') loadMgmtOtkupUzivo();
    } else if (root === 'otkup') {
        showMgmtOtkupSub(window.mgmtShellState.otkupSub || 'otkupi');
    } else if (root === 'kooperanti') {
        showMgmtPartnerSegment('kooperanti');
    } else if (root === 'kupci') {
        showMgmtPartnerSegment('kupci');
    } else if (root === 'fakture') {
        showMgmtPartnerSegment('kupci');
        showMgmtKupSub('fakture');
        if (typeof loadMgmtFakture === 'function') loadMgmtFakture();
    } else if (root === 'izvodi') {
        // placeholder — izvodi i isplate
    } else if (root === 'partneri') {
        showMgmtPartnerSegment(window.mgmtShellState.partnerSegment || 'kooperanti');
    } else if (root === 'agro') {
        if (typeof showMgmtAgroView === 'function') {
            showMgmtAgroView('main');
        }
    }

    setTimeout(() => {
        if (typeof updateRoleNavActive === 'function') {
            updateRoleNavActive();
        }
    }, 0);
}

function showMgmtOtkupSub(sub, btn) {
    window.mgmtShellState.otkupSub = sub;

    const map = {
        otkupi: 'mgmt-sta-otkupi',
        roba: 'mgmt-sta-roba',
        saldo: 'mgmt-sta-saldo'
    };

    mgmtHidePanels(['mgmt-sta-otkupi', 'mgmt-sta-roba', 'mgmt-sta-saldo']);

    const target = document.getElementById(map[sub]);
    if (target) target.classList.add('active');

    mgmtClearActive('#mgmtOtkupSubBar .sub-tab-btn');
    mgmtActivateButton(btn, `#mgmtOtkupSubBar .sub-tab-btn[data-sub="${sub}"]`);

    if (sub === 'otkupi' && typeof loadMgmtOtkupi === 'function') loadMgmtOtkupi();
    if (sub === 'roba' && typeof loadMgmtOtkupPoOM === 'function') loadMgmtOtkupPoOM();
    if (sub === 'saldo' && typeof loadMgmtSaldoOM === 'function') loadMgmtSaldoOM();
}

function showMgmtKoopSub(sub, btn) {
    window.mgmtShellState.koopSub = sub;

    const map = {
        pregled: 'mgmt-koop-pregled',
        saldo:   'mgmt-koop-saldo',
        kartica: 'mgmt-koop-kartica'
    };

    mgmtHidePanels(MGMT_PARTNER_PANEL_IDS);

    const target = document.getElementById(map[sub]);
    if (target) target.classList.add('active');

    mgmtClearActive('#mgmtPartnerKoopSubBar .sub-tab-btn');
    mgmtActivateButton(btn, `#mgmtPartnerKoopSubBar .sub-tab-btn[data-sub="${sub}"]`);

    if (sub === 'pregled' && typeof loadMgmtKoopPregled === 'function') loadMgmtKoopPregled();
    if (sub === 'saldo' && typeof loadMgmtKoopSaldo === 'function') loadMgmtKoopSaldo();
}

function showMgmtKupSub(sub, btn) {
    window.mgmtShellState.kupSub = sub;

    const map = {
        saldo:   'mgmt-kup-saldo',
        fakture: 'mgmt-kup-fakture',
        roba:    'mgmt-kup-roba'
    };

    mgmtHidePanels(MGMT_PARTNER_PANEL_IDS);

    const target = document.getElementById(map[sub]);
    if (target) target.classList.add('active');

    mgmtClearActive('#mgmtPartnerKupSubBar .sub-tab-btn');
    mgmtActivateButton(btn, `#mgmtPartnerKupSubBar .sub-tab-btn[data-sub="${sub}"]`);

    if (sub === 'fakture' && typeof loadMgmtFakture === 'function') loadMgmtFakture();
    if (sub === 'saldo' && typeof loadMgmtKupci === 'function') loadMgmtKupci();
    if (sub === 'roba' && typeof loadMgmtPredato === 'function') loadMgmtPredato();
}

// Kept as a shim — actual logic lives in showMgmtAgroView() in agrohemija.js
function showMgmtAgroSub(sub) {
    const viewMap = { izdavanje: 'izdavanje', stanje: 'main', otpremnice: 'main' };
    if (typeof showMgmtAgroView === 'function') {
        showMgmtAgroView(viewMap[sub] || 'main');
    }
}

function showMgmtPartnerSegment(segment, btn) {
    window.mgmtShellState.partnerSegment = segment;

    mgmtClearActive('#mgmtPartnerSegmentBar .mgmt-seg-btn');
    mgmtActivateButton(btn, `#mgmtPartnerSegmentBar .mgmt-seg-btn[data-segment="${segment}"]`);

    const koopBar = document.getElementById('mgmtPartnerKoopSubBar');
    const kupBar = document.getElementById('mgmtPartnerKupSubBar');
    const koopSection = document.getElementById('mgmt-koop-section');
    const kupSection = document.getElementById('mgmt-kup-section');

    if (koopBar) koopBar.style.display = segment === 'kooperanti' ? '' : 'none';
    if (kupBar) kupBar.style.display = segment === 'kupci' ? '' : 'none';
    if (koopSection) koopSection.style.display = segment === 'kooperanti' ? '' : 'none';
    if (kupSection) kupSection.style.display = segment === 'kupci' ? '' : 'none';

    // Update "+ Novi …" action text per active segment
    const newPartnerBtn = document.querySelector('[data-action="mgmt-novi-partner"]');
    if (newPartnerBtn) {
        newPartnerBtn.textContent = segment === 'kooperanti' ? '+ Novi kooperant' : '+ Novi kupac';
    }

    if (segment === 'kooperanti') {
        showMgmtKoopSub(window.mgmtShellState.koopSub || 'saldo');
        if (typeof loadMgmtKoopSaldo === 'function') loadMgmtKoopSaldo();
    } else {
        showMgmtKupSub(window.mgmtShellState.kupSub || 'saldo');
        if (typeof loadMgmtKupci === 'function') loadMgmtKupci();
    }
}

function mgmtRenderOverview() {
    const today = mgmtDashTodayISO();
    const data = window.mgmtData || mgmtData || {};
    const { otkupiAll, saldoKupci, saldoOM, kartice } = mgmtGetDataSlices();

    const danas = otkupiAll.filter(r => mgmtDashFmtDate(mgmtGetOtkupDate(r)) === today);

    const danasCount = otkupiAll.length
        ? danas.length
        : (parseInt(data.otkupDanasCount || 0, 10) || 0);

    const danasKg = otkupiAll.length
        ? danas.reduce((s, r) => s + mgmtGetOtkupKg(r), 0)
        : (parseFloat(data.otkupDanasKg || 0) || 0);

    let kgCeka = 0;
    if (typeof dpGetSup === 'function') {
        try {
            kgCeka = dpGetSup().reduce((s, r) => s + (parseFloat(r.Kolicina) || 0), 0);
        } catch (e) {
            mgmtLogError('mgmtRenderOverview u dpGetSup', e);
        }
    }

    let aktivniPlanovi = 0;
    if (Array.isArray(dpPlans)) {
        aktivniPlanovi = dpPlans.filter(p => p.Status === 'planned' || p.Status === 'u_toku').length;
    }

    let aktivniKamioni = 0;
    if (Array.isArray(dpKamioni)) {
        aktivniKamioni = dpKamioni.length;
    } else if (stammdaten && Array.isArray(stammdaten.vozaci)) {
        aktivniKamioni = stammdaten.vozaci.length;
    }

    let demandKg = 0;
    if (Array.isArray(dpDem)) {
        demandKg = dpDem.reduce((s, d) => s + (parseInt(d.Kg) || 0), 0);
    }

    mgmtSetText('mgmtOverviewOtkupi', danasCount.toLocaleString('sr'));
    mgmtSetText('mgmtOverviewKg', danasKg.toLocaleString('sr'));
    mgmtSetText('mgmtOverviewCeka', kgCeka.toLocaleString('sr'));
    mgmtSetText('mgmtOverviewPlanovi', aktivniPlanovi.toLocaleString('sr'));
    mgmtSetText('mgmtOverviewKamioni', aktivniKamioni.toLocaleString('sr'));
    mgmtSetText('mgmtOverviewDemand', demandKg.toLocaleString('sr'));

    const alerts = [];

    if (kgCeka > 0) alerts.push('Roba koja čeka: ' + kgCeka.toLocaleString('sr') + ' kg');
    if (demandKg > 0 && aktivniPlanovi === 0) alerts.push('Postoji demand bez aktivnih planova');
    if (saldoKupci.some(r => (parseFloat(r.Saldo) || 0) > 0)) alerts.push('Postoje kupci sa otvorenim saldom');
    if (saldoOM.some(r => (parseFloat(r.Saldo) || 0) > 0)) alerts.push('Postoje mesta sa otvorenim saldom');

    const koopTotals = kartice.filter(r => r.Opis === 'UKUPNO');

    const koopSaldo = kartice.length
        ? koopTotals.reduce((s, r) => s + (parseFloat(r.Saldo) || 0), 0)
        : (parseFloat(data.koopSaldoTotal || 0) || 0);

    const kupSaldo = saldoKupci.reduce((s, r) => s + (parseFloat(r.Saldo) || 0), 0);
    const omSaldo = saldoOM.reduce((s, r) => s + (parseFloat(r.Saldo) || 0), 0);

    const alertsEl = document.getElementById('mgmtOverviewAlerts');
    if (alertsEl) {
        alertsEl.innerHTML = alerts.length
            ? alerts.map(x => `<div class="queue-item"><div class="qi-detail">${escapeHtml(x)}</div></div>`).join('')
            : `<div class="queue-item"><div class="qi-detail">Nema kritičnih upozorenja za prikaz.</div></div>`;
    }

    const opsEl = document.getElementById('mgmtOverviewOps');
    if (opsEl) {
        opsEl.innerHTML = `
            <button class="queue-item" type="button" data-route="mgmt-root" data-root="dispecer" style="cursor:pointer;width:100%;text-align:left;border:0;">
                <div class="qi-header"><span class="qi-koop">Dispečer</span><span class="qi-time">${aktivniPlanovi}</span></div>
                <div class="qi-detail">Aktivni planovi i stanje transporta</div>
            </button>

            <button class="queue-item" type="button" data-route="mgmt-root" data-root="otkup" style="cursor:pointer;width:100%;text-align:left;border:0;">
                <div class="qi-header"><span class="qi-koop">Otkup</span><span class="qi-time">${danasKg.toLocaleString('sr')} kg</span></div>
                <div class="qi-detail">Današnji otkupi i roba po mestima</div>
            </button>

            <button class="queue-item" type="button" data-route="mgmt-root" data-root="agro" style="cursor:pointer;width:100%;text-align:left;border:0;">
                <div class="qi-header"><span class="qi-koop">Agrohemija</span><span class="qi-time">modul</span></div>
                <div class="qi-detail">Izdavanje i stanje magacina</div>
            </button>
        `;
    }

    const financeEl = document.getElementById('mgmtOverviewFinance');
    if (financeEl) {
        financeEl.innerHTML = `
            <div class="queue-item">
                <div class="qi-header"><span class="qi-koop">Kooperanti</span><span class="qi-time">${koopSaldo.toLocaleString('sr')} RSD</span></div>
                <div class="qi-detail">Ukupni otvoreni saldo prema kooperantima</div>
            </div>
            <div class="queue-item">
                <div class="qi-header"><span class="qi-koop">Kupci</span><span class="qi-time">${kupSaldo.toLocaleString('sr')} RSD</span></div>
                <div class="qi-detail">Ukupni otvoreni saldo kupaca</div>
            </div>
            <div class="queue-item">
                <div class="qi-header"><span class="qi-koop">Saldo mesta</span><span class="qi-time">${omSaldo.toLocaleString('sr')} RSD</span></div>
                <div class="qi-detail">Zbirni saldo otkupnih mesta</div>
            </div>
        `;
    }

    const quickEl = document.getElementById('mgmtOverviewQuickLinks');
    if (quickEl) {
        quickEl.innerHTML = `
            <div class="more-menu">
                <button class="more-menu-item" type="button" data-route="mgmt-root" data-root="dispecer">
                    <span class="more-menu-icon">${agIcon('clipboard', '20px')}</span>
                    <span class="more-menu-text"><strong>Dispečer</strong><small>Transport i planovi</small></span>
                </button>

                <button class="more-menu-item" type="button" data-route="mgmt-root" data-root="otkup">
                    <span class="more-menu-icon">${agIcon('coin', '20px')}</span>
                    <span class="more-menu-text"><strong>Otkup</strong><small>Otkupi, roba i saldo mesta</small></span>
                </button>

                <button class="more-menu-item" type="button" data-route="mgmt-root" data-root="partneri">
                    <span class="more-menu-icon">${agIcon('user', '20px')}</span>
                    <span class="more-menu-text"><strong>Partneri</strong><small>Kooperanti i kupci</small></span>
                </button>

                <button class="more-menu-item" type="button" data-route="mgmt-root" data-root="agro">
                    <span class="more-menu-icon">${agIcon('sprout', '20px')}</span>
                    <span class="more-menu-text"><strong>Agrohemija</strong><small>Izdavanje i stanje</small></span>
                </button>
            </div>
        `;
    }
}

// ============================================================
// CANONICAL FIELD ACCESSORS
// Svi mgmt dashboard/overview funkcije moraju koristiti ove
// umesto direktnog r.Datum / r.Kolicina pristupa.
// ============================================================

function mgmtGetOtkupDate(record) {
    return record?.Datum || record?.datum || '';
}

function mgmtGetOtkupKg(record) {
    return parseFloat(record?.Kolicina || record?.kolicina || 0) || 0;
}

// ============================================================
// DATE / NUMBER FORMATTING
// ============================================================

function mgmtDashNum(v) {
    return parseFloat(v) || 0;
}

function mgmtDashFmtInt(v) {
    return Math.round(mgmtDashNum(v)).toLocaleString('sr');
}

function mgmtDashFmtDec(v, digits = 1) {
    return mgmtDashNum(v).toLocaleString('sr', {
        minimumFractionDigits: digits,
        maximumFractionDigits: digits
    });
}

function mgmtDashExcelSerialToISO(value) {
    const serial = Number(value);
    if (!Number.isFinite(serial)) return '';

    const base = new Date(Date.UTC(1899, 11, 30));
    const date = new Date(base.getTime() + Math.round(serial * 86400000));

    const y = date.getUTCFullYear();
    const m = String(date.getUTCMonth() + 1).padStart(2, '0');
    const d = String(date.getUTCDate()).padStart(2, '0');

    return `${y}-${m}-${d}`;
}

function mgmtDashFmtDate(v) {
    if (v === null || v === undefined || v === '') return '';

    // Number → Excel serial
    if (typeof v === 'number') {
        return mgmtDashExcelSerialToISO(v);
    }

    // Date objekat
    if (v instanceof Date && !isNaN(v.getTime())) {
        return mgmtDashLocalISO(v);
    }

    const s = String(v).trim();
    if (!s) return '';

    // Excel/Sheets serial kao string (5-cifreni broj, opcioni decimali)
    if (/^\d{5}(\.\d+)?$/.test(s)) {
        return mgmtDashExcelSerialToISO(Number(s));
    }

    // ISO / ISO-like (2026-04-23 ili 2026-04-23T12:00:00)
    const isoMatch = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (isoMatch) {
        return `${isoMatch[1]}-${isoMatch[2]}-${isoMatch[3]}`;
    }

    // Srpski/evropski formati:
    // 23.04.2026 / 23.04.2026. / 23/04/2026 / 23-04-2026
    // plus opcioni time suffix (npr. 23.04.2026 14:30)
    const localMatch = s.match(
        /^(\d{1,2})[./-](\d{1,2})[./-](\d{4})(?:\.)?(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?$/
    );
    if (localMatch) {
        const dd = String(localMatch[1]).padStart(2, '0');
        const mm = String(localMatch[2]).padStart(2, '0');
        const yyyy = localMatch[3];
        return `${yyyy}-${mm}-${dd}`;
    }

    // Poslednji fallback: native Date parse, ali lokalni ISO (ne UTC)
    const parsed = new Date(s);
    if (!isNaN(parsed.getTime())) {
        return mgmtDashLocalISO(parsed);
    }

    return '';
}

function mgmtDashTodayISO() {
    return mgmtDashLocalISO(new Date());
}

function mgmtDashLocalISO(date) {
    if (!(date instanceof Date) || isNaN(date.getTime())) return '';

    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');

    return `${y}-${m}-${d}`;
}

function mgmtDashGetLastNDays(n) {
    const out = [];
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    for (let i = n - 1; i >= 0; i--) {
        const d = new Date(today);
        d.setDate(today.getDate() - i);
        out.push(mgmtDashLocalISO(d));
    }

    return out;
}

function mgmtDashBuildSeries(otkupiAll, period) {
    if (period === 'today') {
        const today = mgmtDashTodayISO();
        const grouped = {};

        (otkupiAll || []).forEach(r => {
            const day = mgmtDashFmtDate(mgmtGetOtkupDate(r));
            if (day !== today) return;

            const label = mgmtDashGetStationName(r);
            if (!grouped[label]) grouped[label] = 0;
            grouped[label] += mgmtGetOtkupKg(r);
        });

        return Object.keys(grouped)
            .map(label => ({
                label,
                value: grouped[label]
            }))
            .sort((a, b) => b.value - a.value);
    }

    const days = mgmtDashGetPeriodDays(period, otkupiAll);
    const sums = {};
    days.forEach(day => { sums[day] = 0; });

    (otkupiAll || []).forEach(r => {
        const day = mgmtDashFmtDate(mgmtGetOtkupDate(r));
        if (!Object.prototype.hasOwnProperty.call(sums, day)) return;
        sums[day] += mgmtGetOtkupKg(r);
    });

    return days.map(day => ({
        day,
        value: sums[day] || 0
    }));
}

function mgmtDashGetKoopSaldoFromKartice(kartice) {
    const lastByKoop = {};

    (kartice || []).forEach(r => {
        const koopId = r.KooperantID || '';
        if (!koopId) return;

        const currentTs = new Date(r.Datum || 0).getTime() || 0;
        const prev = lastByKoop[koopId];
        const prevTs = prev ? (new Date(prev.Datum || 0).getTime() || 0) : -1;

        if (!prev || currentTs >= prevTs) {
            lastByKoop[koopId] = r;
        }
    });

    return Object.values(lastByKoop).reduce((sum, r) => sum + mgmtDashNum(r.Saldo), 0);
}

function mgmtDashRenderList(elId, items) {
    const el = document.getElementById(elId);
    if (!el) return;

    if (!items.length) {
        el.innerHTML = `<div class="mgmt-dash-list-empty">Nema podataka za prikaz.</div>`;
        return;
    }

    el.innerHTML = items.map(item => `
        <div class="mgmt-dash-list-item">
            <div class="mgmt-dash-list-top">
                <span class="mgmt-dash-list-title">${escapeHtml(item.title || '')}</span>
                ${item.value ? `<span class="mgmt-dash-list-value">${escapeHtml(item.value)}</span>` : ''}
            </div>
            ${item.text ? `<div class="mgmt-dash-list-text">${escapeHtml(item.text)}</div>` : ''}
        </div>
    `).join('');
}

function mgmtDashRenderQuickLinks() {
    const el = document.getElementById('mgmtDashQuickLinks');
    if (!el) return;

    el.innerHTML = `
        <button class="mgmt-dash-quicklink" type="button" data-route="mgmt-root" data-root="pregled">
            <span class="mgmt-dash-quicklink-icon">${agIcon('home', '22px')}</span>
            <span class="mgmt-dash-quicklink-text">
                <strong>Pregled</strong>
                <small>Postojeći overview</small>
            </span>
        </button>

        <button class="mgmt-dash-quicklink" type="button" data-route="mgmt-root" data-root="dispecer">
            <span class="mgmt-dash-quicklink-icon">${agIcon('clipboard', '22px')}</span>
            <span class="mgmt-dash-quicklink-text">
                <strong>Dispečer</strong>
                <small>Transport i planovi</small>
            </span>
        </button>

        <button class="mgmt-dash-quicklink" type="button" data-route="mgmt-root" data-root="otkup">
            <span class="mgmt-dash-quicklink-icon">${agIcon('coin', '22px')}</span>
            <span class="mgmt-dash-quicklink-text">
                <strong>Otkup</strong>
                <small>Otkupi i roba</small>
            </span>
        </button>

        <button class="mgmt-dash-quicklink" type="button" data-route="mgmt-root" data-root="partneri">
            <span class="mgmt-dash-quicklink-icon">${agIcon('user', '22px')}</span>
            <span class="mgmt-dash-quicklink-text">
                <strong>Partneri</strong>
                <small>Kooperanti i kupci</small>
            </span>
        </button>

        <button class="mgmt-dash-quicklink" type="button" data-route="mgmt-root" data-root="agro">
            <span class="mgmt-dash-quicklink-icon">${agIcon('sprout', '22px')}</span>
            <span class="mgmt-dash-quicklink-text">
                <strong>Agrohemija</strong>
                <small>Izdavanje i stanje</small>
            </span>
        </button>
    `;
}

let mgmtDashChartInstance = null;

async function mgmtDashRenderChart(series) {
    try {
        await lazyLoadScript('/vendor/chart.umd.min.js');
    } catch (err) {
        console.error('Chart lazy load failed:', err);
        return;
    }
    
    const canvas = document.getElementById('mgmtDashChart');
    if (!canvas || typeof Chart === 'undefined') return;

    const period = window.mgmtShellState.dashboardPeriod || '7d';
    const isToday = period === 'today';

    const labels = isToday
        ? series.map(item => item.label)
        : series.map(item => {
            const d = new Date(item.day);
            return `${String(d.getDate()).padStart(2, '0')}.${String(d.getMonth() + 1).padStart(2, '0')}.`;
        });

    const data = series.map(item => Math.round(item.value || 0));

    if (mgmtDashChartInstance) {
        mgmtDashChartInstance.destroy();
        mgmtDashChartInstance = null;
    }

    const ctx = canvas.getContext('2d');

    mgmtDashChartInstance = new Chart(ctx, {
        type: isToday ? 'bar' : 'line',
        data: {
            labels,
            datasets: [
                isToday
                    ? {
                        label: 'Otkup kg',
                        data,
                        backgroundColor: 'rgba(26,94,42,0.85)',
                        borderRadius: 8,
                        borderSkipped: false,
                        barThickness: 18,
                        maxBarThickness: 22
                    }
                    : {
                        label: 'Otkup kg',
                        data,
                        borderColor: '#1a5e2a',
                        backgroundColor: 'rgba(26,94,42,0.10)',
                        fill: true,
                        tension: 0.35,
                        borderWidth: 3,
                        pointRadius: 4,
                        pointHoverRadius: 6,
                        pointBackgroundColor: '#1a5e2a',
                        pointBorderColor: '#1a5e2a',
                        pointBorderWidth: 0
                    }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            indexAxis: isToday ? 'y' : 'x',
            animation: {
                duration: 450
            },
            interaction: {
                mode: 'index',
                intersect: false
            },
            plugins: {
                legend: {
                    display: false
                },
                tooltip: {
                    backgroundColor: '#1f2937',
                    titleColor: '#ffffff',
                    bodyColor: '#ffffff',
                    displayColors: false,
                    padding: 10,
                    callbacks: {
                        label(context) {
                            const value = Number(context.raw || 0);
                            return `${value.toLocaleString('sr')} kg`;
                        }
                    }
                }
            },
            layout: {
                padding: {
                    top: 8,
                    right: 8,
                    bottom: 0,
                    left: 8
                }
            },
            scales: isToday
                ? {
                    x: {
                        beginAtZero: true,
                        grace: '8%',
                        grid: {
                            color: '#e8e3d6'
                        },
                        border: {
                            display: false
                        },
                        ticks: {
                            color: '#7b7b72',
                            font: {
                                size: 11
                            },
                            callback(value) {
                                return Number(value).toLocaleString('sr');
                            }
                        }
                    },
                    y: {
                        grid: {
                            display: false
                        },
                        border: {
                            display: false
                        },
                        ticks: {
                            color: '#7b7b72',
                            font: {
                                size: 11
                            }
                        }
                    }
                }
                : {
                    x: {
                        grid: {
                            display: false
                        },
                        border: {
                            display: false
                        },
                        ticks: {
                            color: '#7b7b72',
                            font: {
                                size: 11
                            }
                        }
                    },
                    y: {
                        beginAtZero: true,
                        grace: '8%',
                        grid: {
                            color: '#e8e3d6'
                        },
                        border: {
                            display: false
                        },
                        ticks: {
                            color: '#7b7b72',
                            font: {
                                size: 11
                            },
                            callback(value) {
                                return Number(value).toLocaleString('sr');
                            }
                        }
                    }
                }
        }
    });

    mgmtDashUpdateChartMeta(period, series);
}

function mgmtDashStatusLabel(status) {
    const map = {
        planned: 'Planirano',
        u_toku: 'U toku',
        zavrseno: 'Završeno'
    };
    return map[status] || status || '—';
}

function mgmtDashRenderDispatcher() {
    const el = document.getElementById('mgmtDashDispatcher');
    if (!el) return;

    const plans = Array.isArray(dpPlans)
        ? dpPlans.filter(p => p.Status === 'planned' || p.Status === 'u_toku')
        : [];

    if (!plans.length) {
        el.innerHTML = `<div class="mgmt-dash-list-empty">Nema aktivnih planova za prikaz.</div>`;
        return;
    }

    el.innerHTML = plans
        .sort((a, b) => mgmtDashNum(b.PlannedKg) - mgmtDashNum(a.PlannedKg))
        .slice(0, 6)
        .map(p => `
            <div class="mgmt-dash-dispatch-item">
                <div class="mgmt-dash-dispatch-top">
                    <div>
                        <div class="mgmt-dash-dispatch-name">${escapeHtml(mgmtDashGetDriverName(p))}</div>
                        <div class="mgmt-dash-dispatch-route">
                            ${escapeHtml(p.StanicaName || p.StanicaID || '?')} → ${escapeHtml(p.KupacName || p.KupacID || '?')}
                        </div>
                    </div>
                    <span class="mgmt-dash-status ${escapeHtml(p.Status || 'planned')}">${escapeHtml(mgmtDashStatusLabel(p.Status))}</span>
                </div>

                <div class="mgmt-dash-dispatch-meta">
                    <span><strong>${mgmtDashFmtInt(p.PlannedKg || 0)}</strong> kg</span>
                    <span>Demand: ${escapeHtml(p.DemandID || '-')}</span>
                </div>
            </div>
        `)
        .join('');
}

async function mgmtRenderDashboard() {
    if (!window.mgmtData || !window.mgmtData.saldoOM) {
        if (typeof prefetchMgmtData === 'function') {
            try {
                await prefetchMgmtData();
            } catch (e) {
                mgmtLogError('mgmtRenderDashboard u prefetchMgmtData', e);
            }
        }
    }

    if (typeof ensureMgmtSection === 'function') {
        await ensureMgmtSection('otkupiAll', 'getMgmtOtkupiAll');
        await ensureMgmtSection('kartice', 'getMgmtKartice');
    }

    if (typeof loadDispecer === 'function') {
        try {
            await loadDispecer();
        } catch (e) {
            mgmtLogError('mgmtRenderDashboard u loadDispecer', e);
        }
    }

    const { otkupiAll, saldoKupci, saldoOM, kartice } = mgmtGetDataSlices();

    const period = window.mgmtShellState?.dashboardPeriod || '7d';
    const stats = mgmtDashGetPeriodStats(otkupiAll, period);
    const series = mgmtDashBuildSeries(otkupiAll, period);

    let kgCeka = 0;
    if (typeof dpGetSup === 'function') {
        try {
            kgCeka = dpGetSup().reduce((s, r) => s + mgmtDashNum(r.Kolicina), 0);
        } catch (e) {
            mgmtLogError('mgmtRenderDashboard u dpGetSup', e);
        }
    }

    let demandKg = 0;
    if (Array.isArray(dpDem)) {
        demandKg = dpDem.reduce((s, d) => s + mgmtDashNum(d.Kg), 0);
    }

    const koopSaldo = mgmtDashGetKoopSaldoFromKartice(kartice);
    const kupSaldo = saldoKupci.reduce((s, r) => s + mgmtDashNum(r.Saldo), 0);
    const omSaldo = saldoOM.reduce((s, r) => s + mgmtDashNum(r.Saldo), 0);

    mgmtSetText('mgmtDashWeekKg', mgmtDashFmtInt(stats.totalKg));
    mgmtSetText('mgmtDashActiveKoops', mgmtDashFmtInt(stats.activeKoops));
    mgmtSetText('mgmtDashAvgPrice', mgmtDashFmtDec(stats.avgPrice, 1));
    mgmtSetText('mgmtDashGGAP', '86%');

    const updatedEl = document.getElementById('mgmtDashUpdatedAt');
    if (updatedEl) {
        updatedEl.textContent = 'Ažurirano: ' + new Date().toLocaleTimeString('sr', {
            hour: '2-digit',
            minute: '2-digit'
        });
    }

    const kgLabel = document.getElementById('mgmtDashKgLabel');
    const kgSub = document.querySelector('#mgmtDashWeekKg')?.nextElementSibling;
    const koopSub = document.querySelector('#mgmtDashActiveKoops')?.nextElementSibling;
    const avgSub = document.querySelector('#mgmtDashAvgPrice')?.nextElementSibling;

    const periodCopy = MGMT_DASH_PERIOD_COPY[period] || MGMT_DASH_PERIOD_COPY['7d'];

    if (kgLabel) kgLabel.textContent = periodCopy.kgLabel;
    if (kgSub) kgSub.textContent = periodCopy.kgSub;
    if (koopSub) koopSub.textContent = periodCopy.koopSub;
    if (avgSub) avgSub.textContent = periodCopy.avgSub;

    const alertItems = [];

    if (kgCeka > 0) {
        alertItems.push({
            title: 'Roba čeka',
            value: `${mgmtDashFmtInt(kgCeka)} kg`,
            text: 'Postoje neraspoređene količine koje čekaju transport.'
        });
    }

    if (demandKg > 0) {
        alertItems.push({
            title: 'Otvoren demand',
            value: `${mgmtDashFmtInt(demandKg)} kg`,
            text: 'Postoji aktivna tražnja kupaca u sistemu.'
        });
    }

    if (saldoKupci.some(r => mgmtDashNum(r.Saldo) > 0)) {
        alertItems.push({
            title: 'Kupci sa saldom',
            text: 'Postoje kupci sa otvorenim finansijskim stanjem.'
        });
    }

    if (saldoOM.some(r => mgmtDashNum(r.Saldo) > 0)) {
        alertItems.push({
            title: 'Otkupna mesta sa saldom',
            text: 'Postoje mesta sa otvorenim saldom.'
        });
    }

    const financeItems = [
        {
            title: 'Kooperanti',
            value: `${mgmtDashFmtInt(koopSaldo)} RSD`,
            text: 'Poslednji poznati saldo po kooperantu.'
        },
        {
            title: 'Kupci',
            value: `${mgmtDashFmtInt(kupSaldo)} RSD`,
            text: 'Ukupni otvoreni saldo kupaca.'
        },
        {
            title: 'Saldo mesta',
            value: `${mgmtDashFmtInt(omSaldo)} RSD`,
            text: 'Zbirni saldo otkupnih mesta.'
        }
    ];

    mgmtDashRenderList('mgmtDashAlerts', alertItems);
    mgmtDashRenderList('mgmtDashFinance', financeItems);
    mgmtDashRenderQuickLinks();
    mgmtDashRenderChart(series);
    mgmtDashRenderDispatcher();
}

function mgmtDashGetStationName(record) {
    const stationId = record?.OtkupacID || '';
    if (stationId && window.stammdaten && Array.isArray(stammdaten.stanice)) {
        const hit = stammdaten.stanice.find(s => s.StanicaID === stationId);
        if (hit) return hit.Naziv || hit.Mesto || hit.StanicaID || stationId;
    }

    if (stationId) return stationId;

    const sheetName = record?._sheetName || '';
    if (sheetName.startsWith('OTK-')) return sheetName.replace('OTK-', '');

    return 'Nepoznato';
}

function mgmtDashGetDriverName(plan) {
    if (!plan) return '?';

    const vozacID = plan.VozacID || '';
    const vozacName = plan.VozacName || '';

    if (vozacID && Array.isArray(dpKamioni)) {
        const hit = dpKamioni.find(v =>
            (v.id || v.VozacID || v.vozacID || '') === vozacID
        );
        if (hit && hit.name && hit.name !== vozacID) return hit.name;
    }

    if (window.stammdaten && Array.isArray(stammdaten.vozaci) && vozacID) {
        const hit = stammdaten.vozaci.find(v =>
            (v.VozacID || v.vozacID || v.ID || '') === vozacID
        );
        if (hit) {
            const fullName =
                [hit.Ime, hit.Prezime].filter(Boolean).join(' ').trim() ||
                hit.Naziv ||
                hit.Vozac ||
                '';
            if (fullName && fullName !== vozacID) return fullName;
        }
    }

    if (vozacName && vozacName !== vozacID) return vozacName;
    return vozacID || vozacName || '?';
}

function mgmtDashUpdateChartMeta(period, series) {
    const subtitleEl = document.getElementById('mgmtDashChartSubtitle');
    const legendEl = document.getElementById('mgmtDashChartLegend');

    if (subtitleEl) {
        if (period === 'today') {
            subtitleEl.textContent = 'Danas · kg po stanici';
        } else if (period === 'season') {
            subtitleEl.textContent = 'Sezona · kg po danu';
        } else {
            subtitleEl.textContent = 'Poslednjih 7 dana · kg po danu';
        }
    }

    if (legendEl) {
        const total = (series || []).reduce((sum, item) => sum + (item.value || 0), 0);

        if (period === 'today') {
            legendEl.textContent = `Ukupno danas: ${Math.round(total).toLocaleString('sr')} kg`;
        } else if (period === 'season') {
            legendEl.textContent = `Ukupno u sezoni: ${Math.round(total).toLocaleString('sr')} kg`;
        } else {
            legendEl.textContent = `Ukupno za 7 dana: ${Math.round(total).toLocaleString('sr')} kg`;
        }
    }
}

function setMgmtDashboardPeriod(period, btn) {
    window.mgmtShellState.dashboardPeriod = period;

    // Sync both the desktop period switch and the mobile action row copy
    document.querySelectorAll('.mgmt-dash-period-btn[data-period]')
        .forEach(el => el.classList.toggle('active', el.dataset.period === period));

    mgmtRenderDashboard();
}

function mgmtDashGetSeasonStart(otkupiAll) {
    const validDates = (otkupiAll || [])
        .map(r => mgmtDashFmtDate(mgmtGetOtkupDate(r)))
        .filter(Boolean)
        .sort();

    return validDates.length ? validDates[0] : mgmtDashTodayISO();
}

function mgmtDashGetPeriodDays(period, otkupiAll) {
    if (period === 'today') return [mgmtDashTodayISO()];
    if (period === '7d') return mgmtDashGetLastNDays(7);

    const start = mgmtDashGetSeasonStart(otkupiAll);
    const out = [];

    const from = new Date(start);
    const to = new Date();

    if (isNaN(from.getTime())) return [mgmtDashTodayISO()];

    from.setHours(0, 0, 0, 0);
    to.setHours(0, 0, 0, 0);

    const cursor = new Date(from);
    while (cursor <= to) {
        out.push(mgmtDashLocalISO(cursor));
        cursor.setDate(cursor.getDate() + 1);
    }

    return out;
}

function mgmtDashGetPeriodStats(otkupiAll, period) {
    const days = new Set(mgmtDashGetPeriodDays(period, otkupiAll));
    let totalKg = 0;
    let weightedSum = 0;
    const koops = new Set();

    (otkupiAll || []).forEach(r => {
        const day = mgmtDashFmtDate(mgmtGetOtkupDate(r));
        if (!days.has(day)) return;

        const kg = mgmtGetOtkupKg(r);
        const cena = mgmtDashNum(r.Cena);

        totalKg += kg;
        weightedSum += kg * cena;

        if (r.KooperantID) koops.add(r.KooperantID);
    });

    return {
        totalKg,
        activeKoops: koops.size,
        avgPrice: totalKg > 0 ? (weightedSum / totalKg) : 0
    };
}

// ============================================================
// MGMT DESKTOP SIDEBAR — AgriX Design System v2
// Injects a fixed sidebar for desktop (>= 900px) when the
// management role is active. Falls back gracefully on mobile.
// ============================================================

const MGMT_SIDEBAR_SVG = {
    chart:    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M3 21h18M6 17v-6M11 17V7M16 17v-4M20 17v-9"/></svg>',
    truck:    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M3 6h10v10H3zM13 10h5l3 3v3h-8"/><circle cx="7" cy="17" r="2"/><circle cx="17" cy="17" r="2"/></svg>',
    grid:     '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><path d="M14 14h3v3h-3z"/></svg>',
    building: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M4 21V5a2 2 0 0 1 2-2h12a2 2 0 0 1 2 2v16"/><path d="M3 21h18M8 7h2M14 7h2M8 11h2M14 11h2"/></svg>',
    person:   '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><circle cx="9" cy="8" r="3.5"/><path d="M2 21c0-3.5 3-6 7-6s7 2.5 7 6"/></svg>',
    people:   '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><circle cx="9" cy="8" r="3.5"/><path d="M2 21c0-3.5 3-6 7-6s7 2.5 7 6"/><circle cx="17" cy="7" r="2.5"/></svg>',
    invoice:  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M3 4h18v6H3zM3 14h18v6H3z"/></svg>',
    money:    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M15.5 8.5a5 5 0 1 0 0 7"/><path d="M8 11h6M8 13.5h6"/></svg>',
    leaf:     '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M12 20v-8"/><path d="M12 12c0-4 3-6 7-6-1 5-3 7-7 6Z"/><path d="M12 14c0-3-2-5-6-5 1 4 3 6 6 5Z"/></svg>',
    overview: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 21V9"/></svg>',
};

const MGMT_NAV_ITEMS = [
    { group: 'Operativa' },
    { root: 'pregled',    label: 'Pregled',       icon: 'chart'    },
    { root: 'dispecer',   label: 'Dispečer',      icon: 'truck'    },
    { root: 'uzivo',      label: 'Otkup uživo',   icon: 'grid'     },
    { root: 'otkup',      label: 'Stanice',       icon: 'building' },
    { group: 'Partneri' },
    { root: 'kooperanti', label: 'Kooperanti',    icon: 'person'   },
    { root: 'kupci',      label: 'Kupci',         icon: 'people'   },
    { root: 'fakture',    label: 'Fakture i SEF',    icon: 'invoice' },
    { root: 'izvodi',     label: 'Izvodi i isplate', icon: 'money'   },
    { group: 'Finansije' },
    { root: 'agro',       label: 'Agrohemija',       icon: 'leaf'    },
];

function mgmtGetUserInitials() {
    try {
        const u = firebase.auth().currentUser;
        if (!u) return '—';
        const d = u.displayName || u.email || '';
        const parts = d.trim().split(/\s+/);
        if (parts.length >= 2) return (parts[0][0] + parts[1][0]).toUpperCase();
        return d.slice(0, 2).toUpperCase();
    } catch (_) { return '—'; }
}

function mgmtGetUserName() {
    try {
        const u = firebase.auth().currentUser;
        if (!u) return 'Korisnik';
        return u.displayName || u.email || 'Korisnik';
    } catch (_) { return 'Korisnik'; }
}

function mgmtSidebarHTML(activeRoot) {
    const items = MGMT_NAV_ITEMS.map(item => {
        if (item.group) {
            return `<div class="mgmt-side__group">${item.group}</div>`;
        }
        const active = item.root === activeRoot ? ' is-active' : '';
        return `
            <button class="mgmt-side__item${active}"
                    data-action="mgmt-sidebar-nav"
                    data-root="${item.root}">
                ${MGMT_SIDEBAR_SVG[item.icon] || ''}
                ${item.label}
            </button>`;
    }).join('');

    const initials = mgmtGetUserInitials();
    const name = mgmtGetUserName();

    return `
        <div class="mgmt-side__brand">Agri<em>X</em></div>
        <div class="mgmt-side__role">Uprava</div>
        ${items}
        <div class="mgmt-side__user">
            <div class="av">${initials}</div>
            <div>
                <div class="nm">${escapeHtml(name)}</div>
                <div class="rl">Management</div>
            </div>
        </div>
    `;
}

function mgmtSidebarMount() {
    let sidebar = document.getElementById('mgmt-desktop-sidebar');
    if (!sidebar) {
        sidebar = document.createElement('div');
        sidebar.id = 'mgmt-desktop-sidebar';
        sidebar.className = 'mgmt-side';
        sidebar.setAttribute('role', 'navigation');
        sidebar.setAttribute('aria-label', 'Management navigacija');
        // Inject into body so it can be positioned via CSS
        document.body.appendChild(sidebar);
    }
    sidebar.innerHTML = mgmtSidebarHTML(
        (window.mgmtShellState && window.mgmtShellState.activeRoot) || 'pregled'
    );
}

function mgmtSidebarUpdateActive(root) {
    const sidebar = document.getElementById('mgmt-desktop-sidebar');
    if (!sidebar) return;
    sidebar.querySelectorAll('.mgmt-side__item').forEach(btn => {
        btn.classList.toggle('is-active', btn.dataset.root === root);
    });
}

// Handle sidebar nav clicks
document.addEventListener('click', function(e) {
    const btn = e.target.closest('[data-action="mgmt-sidebar-nav"]');
    if (!btn) return;
    const root = btn.dataset.root;
    if (root && typeof showMgmtRoot === 'function') {
        showMgmtRoot(root);
    }
});

// ============================================================
// MANAGEMENT MOBILE — tbar, drawer, wizard
// ============================================================

const MGMT_TBAR_LABELS = {
    dashboard:  { sub: 'Upravljanje', h: 'Dashboard' },
    pregled:    { sub: 'Upravljanje', h: 'Pregled' },
    dispecer:   { sub: 'Operativa',   h: 'Dispečer' },
    uzivo:      { sub: 'Otkup',       h: 'Uživo' },
    otkup:      { sub: 'Otkup',       h: 'Stanice' },
    partneri:   { sub: 'Partneri',    h: 'Kooperanti i kupci' },
    kooperanti: { sub: 'Partneri',    h: 'Kooperanti' },
    kupci:      { sub: 'Partneri',    h: 'Kupci' },
    fakture:    { sub: 'Partneri',    h: 'Fakture' },
    izvodi:     { sub: 'Finansije',   h: 'Izvodi' },
    agro:       { sub: 'Agrohemija',  h: 'Izdavanje' }
};

function mgmtMobileUpdateTbar(root) {
    const lbl = MGMT_TBAR_LABELS[root] || { sub: 'Uprava', h: root || 'Pregled' };
    const subEl = document.getElementById('mgmtTbarSub');
    const hEl   = document.getElementById('mgmtTbarH');
    if (subEl) subEl.textContent = lbl.sub;
    if (hEl)   hEl.textContent   = lbl.h;

    document.querySelectorAll('.mgmt-drawer__item[data-root]').forEach(btn => {
        btn.classList.toggle('is-active', btn.dataset.root === root);
    });
}

function mgmtMobileTbarSync() {
    const isMgmt   = window.CONFIG && String(CONFIG.USER_ROLE || '').toLowerCase() === 'management';
    const isMobile = window.innerWidth <= 900;
    const tbar     = document.getElementById('mgmtMobileTbar');
    const header   = document.querySelector('.header');

    if (!tbar) return;

    if (isMgmt && isMobile) {
        tbar.style.display = 'flex';
        if (header) header.style.display = 'none';
    } else {
        tbar.style.display = 'none';
        if (header) header.style.display = '';
    }
}

function mgmtDrawerOpen() {
    const drawer = document.getElementById('mgmtDrawer');
    if (drawer) drawer.classList.add('is-open');
    document.body.style.overflow = 'hidden';
}

function mgmtDrawerClose() {
    const drawer = document.getElementById('mgmtDrawer');
    if (drawer) drawer.classList.remove('is-open');
    document.body.style.overflow = '';
}

function mgmtDrawerPopulateUser() {
    const initials      = mgmtGetUserInitials();
    const name          = mgmtGetUserName();
    const avatarEl      = document.getElementById('mgmtTbarAvatar');
    const drawerAvEl    = document.getElementById('mgmtDrawerAvatar');
    const drawerNameEl  = document.getElementById('mgmtDrawerName');
    if (avatarEl)     avatarEl.textContent     = initials;
    if (drawerAvEl)   drawerAvEl.textContent   = initials;
    if (drawerNameEl) drawerNameEl.textContent = name;
}

// ── Agrohemija wizard mobile step navigation ──────────────────

let _izWzStep = 1;

function mgmtWizardStep(n) {
    _izWzStep = n;

    ['izdWzCol1', 'izdWzCol2', 'izdWzCol3'].forEach((id, i) => {
        const el = document.getElementById(id);
        if (el) el.classList.toggle('iz-wz-hidden', i + 1 !== n);
    });

    for (let i = 1; i <= 3; i++) {
        const stepEl = document.getElementById('izWpStep' + i);
        if (!stepEl) continue;
        const dot = stepEl.querySelector('.iz-wz-progress__dot');
        const lbl = stepEl.querySelector('.iz-wz-progress__lbl');
        if (dot) {
            dot.classList.remove('is-active', 'is-done');
            if (i < n)      dot.classList.add('is-done');
            else if (i === n) dot.classList.add('is-active');
        }
        if (lbl) lbl.classList.toggle('is-active', i === n);
    }

    for (let i = 1; i <= 2; i++) {
        const bar = document.getElementById('izWpBar' + i);
        if (bar) bar.classList.toggle('is-done', i < n);
    }

    const backBtn = document.getElementById('izWzBackBtn');
    const nextBtn = document.getElementById('izWzNextBtn');
    if (backBtn) backBtn.textContent = n === 1 ? 'Otkaži' : '← Nazad';
    const nextLabels = ['Dalje · Artikli →', 'Dalje · Korpa →', 'Izvrši izdavanje'];
    if (nextBtn) nextBtn.textContent = nextLabels[n - 1] || 'Dalje';
}

function mgmtWizardInitMobile() {
    if (window.innerWidth > 900) {
        ['izdWzCol1', 'izdWzCol2', 'izdWzCol3'].forEach(id => {
            const el = document.getElementById(id);
            if (el) el.classList.remove('iz-wz-hidden');
        });
        return;
    }
    mgmtWizardStep(1);
}

// ── Event delegation for mobile interactions ──────────────────

document.addEventListener('click', function(e) {
    // Hamburger button or "Više" bottom nav tab → open drawer
    if (e.target.closest('#mgmtDrawerBtn') || e.target.closest('[data-action="mgmt-drawer-toggle"]')) {
        mgmtDrawerOpen();
        return;
    }

    // Close button or dim overlay → close drawer
    if (e.target.closest('#mgmtDrawerClose') || e.target.closest('#mgmtDrawerDim')) {
        mgmtDrawerClose();
        return;
    }

    // Drawer nav item → navigate + close
    const drawerItem = e.target.closest('[data-action="mgmt-drawer-nav"]');
    if (drawerItem) {
        const root = drawerItem.dataset.root;
        const segment = drawerItem.dataset.segment;
        mgmtDrawerClose();
        if (root && typeof showMgmtRoot === 'function') {
            showMgmtRoot(root).then(() => {
                if (segment && typeof showMgmtPartnerSegment === 'function') {
                    showMgmtPartnerSegment(segment);
                }
            });
        }
        return;
    }

    // Wizard next / submit
    if (e.target.closest('#izWzNextBtn')) {
        if (_izWzStep < 3) {
            mgmtWizardStep(_izWzStep + 1);
        } else {
            const saveBtn = document.getElementById('izdSaveBtn');
            if (saveBtn) saveBtn.click();
        }
        return;
    }

    // Wizard back / cancel
    if (e.target.closest('#izWzBackBtn')) {
        if (_izWzStep > 1) {
            mgmtWizardStep(_izWzStep - 1);
        } else {
            if (typeof showMgmtAgroView === 'function') showMgmtAgroView('main');
        }
        return;
    }
});

window.addEventListener('resize', mgmtMobileTbarSync);

// Patch showMgmtRoot to also update sidebar active state
const _origShowMgmtRoot = window.showMgmtRoot || showMgmtRoot;
window.showMgmtRoot = async function(root, btn) {
    mgmtSidebarUpdateActive(root);
    return _origShowMgmtRoot.call(this, root, btn);
};

// Patch mgmtShellInit: mount sidebar then run mobile setup
const _origMgmtShellInit = window.mgmtShellInit || mgmtShellInit;
window.mgmtShellInit = async function() {
    mgmtSidebarMount();
    const result = await _origMgmtShellInit.call(this);
    mgmtMobileTbarSync();
    mgmtMobileUpdateTbar((window.mgmtShellState && window.mgmtShellState.activeRoot) || 'pregled');
    mgmtDrawerPopulateUser();
    return result;
};

// Patch showMgmtAgroView to init wizard on mobile when entering izdavanje
(function() {
    const _origAgroView = window.showMgmtAgroView;
    if (typeof _origAgroView !== 'function') return;
    window.showMgmtAgroView = function(view) {
        _origAgroView.call(this, view);
        if (view === 'izdavanje') {
            mgmtWizardInitMobile();
            if (window.innerWidth <= 900) {
                const subEl = document.getElementById('mgmtTbarSub');
                const hEl   = document.getElementById('mgmtTbarH');
                if (subEl) subEl.textContent = 'Agrohemija';
                if (hEl)   hEl.textContent   = 'Novo izdavanje';
            }
        } else if (view === 'main') {
            mgmtMobileUpdateTbar('agro');
        }
    };
})();
