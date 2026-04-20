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
    mounted: false
};

function mgmtShellInit() {
    mgmtMountLegacyBlocks();
    mgmtRenderOverview();
    showMgmtRoot('pregled');
    if (typeof updateRoleNavVisibility === 'function') updateRoleNavVisibility();
    if (typeof updateRoleNavActive === 'function') updateRoleNavActive();
}

function showMgmtRoot(root, btn) {
    window.mgmtShellState.activeRoot = root;

    document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));

    const rootMap = {
        dashboard: 'tab-mgmt-dashboard',
        pregled: 'tab-mgmt-pregled',
        dispecer: 'tab-mgmt-dispecer',
        otkup: 'tab-mgmt-otkup',
        partneri: 'tab-mgmt-partneri',
        agro: 'tab-mgmt-agro'
    };

    const targetId = rootMap[root];
    const target = document.getElementById(targetId);
    if (target) target.classList.add('active');

    document.querySelectorAll('.tab-btn.role-management').forEach(el => el.classList.remove('active'));
    if (btn && btn.classList) {
        btn.classList.add('active');
    } else {
        const autoBtn = document.querySelector('.tab-btn.role-management[data-mgmt-root="' + root + '"]');
        if (autoBtn) autoBtn.classList.add('active');
    }

    if (root === 'dashboard') {
        mgmtRenderDashboard();
    } else if (root === 'pregled') {
        mgmtRenderOverview();
    } else if (root === 'dispecer') {
        if (typeof loadDispecer === 'function') loadDispecer();
    } else if (root === 'otkup') {
        showMgmtOtkupSub(window.mgmtShellState.otkupSub || 'otkupi');
    } else if (root === 'partneri') {
        showMgmtPartnerSegment(window.mgmtShellState.partnerSegment || 'kooperanti');
    } else if (root === 'agro') {
        showMgmtAgroSub(window.mgmtShellState.agroSub || 'izdavanje');
    }

    setTimeout(() => {
        if (typeof updateRoleNavActive === 'function') {
            updateRoleNavActive();
        }
    }, 0);
}

function showMgmtBottomRoot(root, btn) {
    showMgmtRoot(root);
    if (typeof updateRoleNavActive === 'function') updateRoleNavActive(root);
}

function updateMgmtBottomNavActive() {
    const nav = document.getElementById('mgmtBottomNav');
    if (!nav) return;

    nav.querySelectorAll('.bottom-nav-btn').forEach(el => {
        el.classList.toggle('active', el.dataset.tab === window.mgmtShellState.activeRoot);
    });
}

function updateMgmtBottomNavVisibility() {
    const nav = document.getElementById('mgmtBottomNav');
    if (!nav) return;

    const isMgmt = CONFIG.USER_ROLE === 'Management';
    const isMobile = window.innerWidth <= 900;

    nav.classList.toggle('visible', !!(isMgmt && isMobile));
    document.body.classList.toggle('has-mgmt-bottom-nav', !!(isMgmt && isMobile));
}

function mgmtMountLegacyBlocks() {
    if (window.mgmtShellState.mounted) return;
    window.mgmtShellState.mounted = true;

    // Dispecer
    const oldDispecer = document.getElementById('tab-dispecer');
    const dispecerMount = document.getElementById('mgmtDispecerMount');
    if (oldDispecer && dispecerMount && oldDispecer !== dispecerMount) {
        while (oldDispecer.firstChild) {
            dispecerMount.appendChild(oldDispecer.firstChild);
        }
        oldDispecer.remove();
    }

    // Otkup
    const otkupMount = document.getElementById('mgmtOtkupMount');
    if (otkupMount) {
        const otkupi = document.getElementById('mgmt-sta-otkupi');
        const roba = document.getElementById('mgmt-sta-roba');
        const saldo = document.getElementById('mgmt-sta-saldo');

        if (otkupi) otkupMount.appendChild(otkupi);
        if (roba) otkupMount.appendChild(roba);
        if (saldo) otkupMount.appendChild(saldo);
    }

    // Partneri
    const partneriMount = document.getElementById('mgmtPartneriMount');
    if (partneriMount) {
        [
            'mgmt-koop-pregled',
            'mgmt-koop-kartica',
            'mgmt-koop-saldo',
            'mgmt-kup-fakture',
            'mgmt-kup-saldo',
            'mgmt-kup-roba'
        ].forEach(id => {
            const el = document.getElementById(id);
            if (el) partneriMount.appendChild(el);
        });
    }

    // Agro
    const agroMount = document.getElementById('mgmtAgroMount');
    if (agroMount) {
        const izd = document.getElementById('mgmt-agro-izdavanje');
        const stanje = document.getElementById('mgmt-agro-stanje');
        if (izd) agroMount.appendChild(izd);
        if (stanje) agroMount.appendChild(stanje);
    }

    // Sakrij legacy wrapper ako ostane
    const legacy = document.getElementById('tab-mgmt');
    if (legacy) legacy.style.display = 'none';
}

function showMgmtOtkupSub(sub, btn) {
    window.mgmtShellState.otkupSub = sub;

    const map = {
        otkupi: 'mgmt-sta-otkupi',
        roba: 'mgmt-sta-roba',
        saldo: 'mgmt-sta-saldo'
    };

    ['mgmt-sta-otkupi', 'mgmt-sta-roba', 'mgmt-sta-saldo'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.classList.remove('active');
    });

    const target = document.getElementById(map[sub]);
    if (target) target.classList.add('active');

    document.querySelectorAll('#mgmtOtkupSubBar .sub-tab-btn').forEach(el => el.classList.remove('active'));
    if (btn && btn.classList) {
        btn.classList.add('active');
    } else {
        const autoBtn = document.querySelector('#mgmtOtkupSubBar .sub-tab-btn[data-sub="' + sub + '"]');
        if (autoBtn) autoBtn.classList.add('active');
    }

    if (sub === 'otkupi' && typeof loadMgmtOtkupi === 'function') loadMgmtOtkupi();
    if (sub === 'roba' && typeof loadMgmtOtkupPoOM === 'function') loadMgmtOtkupPoOM();
    if (sub === 'saldo' && typeof loadMgmtSaldoOM === 'function') loadMgmtSaldoOM();
}

function showMgmtPartnerSegment(segment, btn) {
    window.mgmtShellState.partnerSegment = segment;

    document.querySelectorAll('#mgmtPartnerSegmentBar .sub-tab-btn').forEach(el => el.classList.remove('active'));
    if (btn && btn.classList) {
        btn.classList.add('active');
    } else {
        const autoBtn = document.querySelector('#mgmtPartnerSegmentBar .sub-tab-btn[data-segment="' + segment + '"]');
        if (autoBtn) autoBtn.classList.add('active');
    }

    const koopBar = document.getElementById('mgmtPartnerKoopSubBar');
    const kupBar = document.getElementById('mgmtPartnerKupSubBar');

    if (koopBar) koopBar.style.display = segment === 'kooperanti' ? '' : 'none';
    if (kupBar) kupBar.style.display = segment === 'kupci' ? '' : 'none';

    if (segment === 'kooperanti') {
        showMgmtKoopSub(window.mgmtShellState.koopSub || 'pregled');
    } else {
        showMgmtKupSub(window.mgmtShellState.kupSub || 'fakture');
    }
}

function showMgmtKoopSub(sub, btn) {
    window.mgmtShellState.koopSub = sub;

    const map = {
        pregled: 'mgmt-koop-pregled',
        kartica: 'mgmt-koop-kartica',
        saldo: 'mgmt-koop-saldo'
    };

    [
        'mgmt-koop-pregled',
        'mgmt-koop-kartica',
        'mgmt-koop-saldo',
        'mgmt-kup-fakture',
        'mgmt-kup-saldo',
        'mgmt-kup-roba'
    ].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.classList.remove('active');
    });

    const target = document.getElementById(map[sub]);
    if (target) target.classList.add('active');

    document.querySelectorAll('#mgmtPartnerKoopSubBar .sub-tab-btn').forEach(el => el.classList.remove('active'));
    if (btn && btn.classList) {
        btn.classList.add('active');
    } else {
        const autoBtn = document.querySelector('#mgmtPartnerKoopSubBar .sub-tab-btn[data-sub="' + sub + '"]');
        if (autoBtn) autoBtn.classList.add('active');
    }

    if (sub === 'pregled' && typeof loadMgmtKoopPregled === 'function') loadMgmtKoopPregled();
    if (sub === 'saldo' && typeof loadMgmtKoopSaldo === 'function') loadMgmtKoopSaldo();
}

function showMgmtKupSub(sub, btn) {
    window.mgmtShellState.kupSub = sub;

    const map = {
        fakture: 'mgmt-kup-fakture',
        saldo: 'mgmt-kup-saldo',
        roba: 'mgmt-kup-roba'
    };

    [
        'mgmt-koop-pregled',
        'mgmt-koop-kartica',
        'mgmt-koop-saldo',
        'mgmt-kup-fakture',
        'mgmt-kup-saldo',
        'mgmt-kup-roba'
    ].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.classList.remove('active');
    });

    const target = document.getElementById(map[sub]);
    if (target) target.classList.add('active');

    document.querySelectorAll('#mgmtPartnerKupSubBar .sub-tab-btn').forEach(el => el.classList.remove('active'));
    if (btn && btn.classList) {
        btn.classList.add('active');
    } else {
        const autoBtn = document.querySelector('#mgmtPartnerKupSubBar .sub-tab-btn[data-sub="' + sub + '"]');
        if (autoBtn) autoBtn.classList.add('active');
    }

    if (sub === 'fakture' && typeof loadMgmtFakture === 'function') loadMgmtFakture();
    if (sub === 'saldo' && typeof loadMgmtKupci === 'function') loadMgmtKupci();
    if (sub === 'roba' && typeof loadMgmtPredato === 'function') loadMgmtPredato();
}

function showMgmtAgroSub(sub, btn) {
    window.mgmtShellState.agroSub = sub;

    const map = {
        izdavanje: 'mgmt-agro-izdavanje',
        stanje: 'mgmt-agro-stanje'
    };

    ['mgmt-agro-izdavanje', 'mgmt-agro-stanje'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.classList.remove('active');
    });

    const target = document.getElementById(map[sub]);
    if (target) target.classList.add('active');

    document.querySelectorAll('#mgmtAgroSubBar .sub-tab-btn').forEach(el => el.classList.remove('active'));
    if (btn && btn.classList) {
        btn.classList.add('active');
    } else {
        const autoBtn = document.querySelector('#mgmtAgroSubBar .sub-tab-btn[data-sub="' + sub + '"]');
        if (autoBtn) autoBtn.classList.add('active');
    }

    if (sub === 'izdavanje' && typeof populateIzdDropdowns === 'function') populateIzdDropdowns();
    if (sub === 'stanje' && typeof loadMgmtAgroStanje === 'function') loadMgmtAgroStanje();
}

function mgmtRenderOverview() {
    const today = new Date().toISOString().split('T')[0];
    const otkupiAll = (window.mgmtData && mgmtData.otkupiAll) ? mgmtData.otkupiAll : [];
    const saldoKupci = (window.mgmtData && mgmtData.saldoKupci) ? mgmtData.saldoKupci : [];
    const saldoOM = (window.mgmtData && mgmtData.saldoOM) ? mgmtData.saldoOM : [];
    const kartice = (window.mgmtData && mgmtData.kartice) ? mgmtData.kartice : [];

    const danas = otkupiAll.filter(r => fmtDate(r.Datum) === today);
    const danasCount = danas.length;
    const danasKg = danas.reduce((s, r) => s + (parseFloat(r.Kolicina) || 0), 0);

    let kgCeka = 0;
    if (typeof dpGetSup === 'function') {
        try {
            kgCeka = dpGetSup().reduce((s, r) => s + (parseFloat(r.Kolicina) || 0), 0);
        } catch (e) {}
    }

    let aktivniPlanovi = 0;
    if (window.dpPlans && Array.isArray(dpPlans)) {
        aktivniPlanovi = dpPlans.filter(p => p.Status === 'planned' || p.Status === 'u_toku').length;
    }

    let aktivniKamioni = 0;
    if (window.dpKamioni && Array.isArray(dpKamioni)) {
        aktivniKamioni = dpKamioni.length;
    } else if (stammdaten && Array.isArray(stammdaten.vozaci)) {
        aktivniKamioni = stammdaten.vozaci.length;
    }

    let demandKg = 0;
    if (window.dpDem && Array.isArray(dpDem)) {
        demandKg = dpDem.reduce((s, d) => s + (parseInt(d.Kg) || 0), 0);
    }

    setTextSafe('mgmtOverviewOtkupi', danasCount.toLocaleString('sr'));
    setTextSafe('mgmtOverviewKg', danasKg.toLocaleString('sr'));
    setTextSafe('mgmtOverviewCeka', kgCeka.toLocaleString('sr'));
    setTextSafe('mgmtOverviewPlanovi', aktivniPlanovi.toLocaleString('sr'));
    setTextSafe('mgmtOverviewKamioni', aktivniKamioni.toLocaleString('sr'));
    setTextSafe('mgmtOverviewDemand', demandKg.toLocaleString('sr'));

    const alerts = [];

    if (kgCeka > 0) alerts.push('Roba koja čeka: ' + kgCeka.toLocaleString('sr') + ' kg');
    if (demandKg > 0 && aktivniPlanovi === 0) alerts.push('Postoji demand bez aktivnih planova');
    if (saldoKupci.some(r => (parseFloat(r.Saldo) || 0) > 0)) alerts.push('Postoje kupci sa otvorenim saldom');
    if (saldoOM.some(r => (parseFloat(r.Saldo) || 0) > 0)) alerts.push('Postoje mesta sa otvorenim saldom');

    const koopTotals = kartice.filter(r => r.Opis === 'UKUPNO');
    const koopSaldo = koopTotals.reduce((s, r) => s + (parseFloat(r.Saldo) || 0), 0);
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
            <div class="queue-item" onclick="showMgmtRoot('dispecer')" style="cursor:pointer;">
                <div class="qi-header"><span class="qi-koop">Dispečer</span><span class="qi-time">${aktivniPlanovi}</span></div>
                <div class="qi-detail">Aktivni planovi i stanje transporta</div>
            </div>
            <div class="queue-item" onclick="showMgmtRoot('otkup')" style="cursor:pointer;">
                <div class="qi-header"><span class="qi-koop">Otkup</span><span class="qi-time">${danasKg.toLocaleString('sr')} kg</span></div>
                <div class="qi-detail">Današnji otkupi i roba po mestima</div>
            </div>
            <div class="queue-item" onclick="showMgmtRoot('agro')" style="cursor:pointer;">
                <div class="qi-header"><span class="qi-koop">Agrohemija</span><span class="qi-time">modul</span></div>
                <div class="qi-detail">Izdavanje i stanje magacina</div>
            </div>
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
                <button class="more-menu-item" onclick="showMgmtRoot('dispecer')">
                    <span class="more-menu-icon">📋</span>
                    <span class="more-menu-text"><strong>Dispečer</strong><small>Transport i planovi</small></span>
                </button>
                <button class="more-menu-item" onclick="showMgmtRoot('otkup')">
                    <span class="more-menu-icon">◷</span>
                    <span class="more-menu-text"><strong>Otkup</strong><small>Otkupi, roba i saldo mesta</small></span>
                </button>
                <button class="more-menu-item" onclick="showMgmtRoot('partneri')">
                    <span class="more-menu-icon">◎</span>
                    <span class="more-menu-text"><strong>Partneri</strong><small>Kooperanti i kupci</small></span>
                </button>
                <button class="more-menu-item" onclick="showMgmtRoot('agro')">
                    <span class="more-menu-icon">✿</span>
                    <span class="more-menu-text"><strong>Agrohemija</strong><small>Izdavanje i stanje</small></span>
                </button>
            </div>
        `;
    }
}

function setTextSafe(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
}

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

function mgmtDashFmtDate(v) {
    if (!v) return '';
    const s = String(v);
    return s.length >= 10 ? s.slice(0, 10) : s;
}

function mgmtDashSetText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
}

function mgmtDashEscape(value) {
    return String(value ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function mgmtDashTodayISO() {
    return new Date().toISOString().split('T')[0];
}

function mgmtDashGetLastNDays(n) {
    const out = [];
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    for (let i = n - 1; i >= 0; i--) {
        const d = new Date(today);
        d.setDate(today.getDate() - i);
        out.push(d.toISOString().split('T')[0]);
    }

    return out;
}

function mgmtDashBuild7DaySeries(otkupiAll) {
    const days = mgmtDashGetLastNDays(7);
    const sums = {};
    days.forEach(day => { sums[day] = 0; });

    (otkupiAll || []).forEach(r => {
        const day = mgmtDashFmtDate(r.Datum);
        if (!Object.prototype.hasOwnProperty.call(sums, day)) return;
        sums[day] += mgmtDashNum(r.Kolicina);
    });

    return days.map(day => ({
        day,
        value: sums[day] || 0
    }));
}

function mgmtDashGetWeekStats(otkupiAll) {
    const days = new Set(mgmtDashGetLastNDays(7));
    let weekKg = 0;
    let weightedSum = 0;
    const koops = new Set();

    (otkupiAll || []).forEach(r => {
        const day = mgmtDashFmtDate(r.Datum);
        if (!days.has(day)) return;

        const kg = mgmtDashNum(r.Kolicina);
        const cena = mgmtDashNum(r.Cena);

        weekKg += kg;
        weightedSum += kg * cena;

        if (r.KooperantID) koops.add(r.KooperantID);
    });

    return {
        weekKg,
        activeKoops: koops.size,
        avgPrice: weekKg > 0 ? (weightedSum / weekKg) : 0
    };
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
                <span class="mgmt-dash-list-title">${mgmtDashEscape(item.title || '')}</span>
                ${item.value ? `<span class="mgmt-dash-list-value">${mgmtDashEscape(item.value)}</span>` : ''}
            </div>
            ${item.text ? `<div class="mgmt-dash-list-text">${mgmtDashEscape(item.text)}</div>` : ''}
        </div>
    `).join('');
}

function mgmtDashRenderQuickLinks() {
    const el = document.getElementById('mgmtDashQuickLinks');
    if (!el) return;

    el.innerHTML = `
        <button class="mgmt-dash-quicklink" type="button" onclick="showMgmtRoot('pregled')">
            <span class="mgmt-dash-quicklink-icon">⌂</span>
            <span class="mgmt-dash-quicklink-text">
                <strong>Pregled</strong>
                <small>Postojeći overview</small>
            </span>
        </button>

        <button class="mgmt-dash-quicklink" type="button" onclick="showMgmtRoot('dispecer')">
            <span class="mgmt-dash-quicklink-icon">📋</span>
            <span class="mgmt-dash-quicklink-text">
                <strong>Dispečer</strong>
                <small>Transport i planovi</small>
            </span>
        </button>

        <button class="mgmt-dash-quicklink" type="button" onclick="showMgmtRoot('otkup')">
            <span class="mgmt-dash-quicklink-icon">◷</span>
            <span class="mgmt-dash-quicklink-text">
                <strong>Otkup</strong>
                <small>Otkupi i roba</small>
            </span>
        </button>

        <button class="mgmt-dash-quicklink" type="button" onclick="showMgmtRoot('partneri')">
            <span class="mgmt-dash-quicklink-icon">◎</span>
            <span class="mgmt-dash-quicklink-text">
                <strong>Partneri</strong>
                <small>Kooperanti i kupci</small>
            </span>
        </button>

        <button class="mgmt-dash-quicklink" type="button" onclick="showMgmtRoot('agro')">
            <span class="mgmt-dash-quicklink-icon">✿</span>
            <span class="mgmt-dash-quicklink-text">
                <strong>Agrohemija</strong>
                <small>Izdavanje i stanje</small>
            </span>
        </button>
    `;
}

function mgmtDashRenderChart(series) {
    const svg = document.getElementById('mgmtDashChart');
    const legend = document.getElementById('mgmtDashChartLegend');
    if (!svg) return;

    const width = 680;
    const height = 260;
    const pad = { top: 16, right: 18, bottom: 34, left: 42 };
    const innerW = width - pad.left - pad.right;
    const innerH = height - pad.top - pad.bottom;
    const maxVal = Math.max(1, ...series.map(x => x.value));
    const stepX = series.length > 1 ? innerW / (series.length - 1) : innerW;

    const points = series.map((item, i) => {
        const x = pad.left + (i * stepX);
        const y = pad.top + innerH - ((item.value / maxVal) * innerH);
        return { ...item, x, y };
    });

    const polyline = points.map(p => `${p.x},${p.y}`).join(' ');
    const area = [
        `${pad.left},${pad.top + innerH}`,
        ...points.map(p => `${p.x},${p.y}`),
        `${pad.left + innerW},${pad.top + innerH}`
    ].join(' ');

    const grid = [];
    const ticks = 4;
    for (let i = 0; i <= ticks; i++) {
        const y = pad.top + ((innerH / ticks) * i);
        const value = Math.round(maxVal - ((maxVal / ticks) * i));
        grid.push(`
            <line x1="${pad.left}" y1="${y}" x2="${pad.left + innerW}" y2="${y}" stroke="#e8e3d6" stroke-width="1" />
            <text x="${pad.left - 8}" y="${y + 4}" text-anchor="end" font-size="11" fill="#7b7b72">${mgmtDashFmtInt(value)}</text>
        `);
    }

    const labels = points.map(p => {
        const d = new Date(p.day);
        const txt = `${String(d.getDate()).padStart(2, '0')}.${String(d.getMonth() + 1).padStart(2, '0')}.`;
        return `<text x="${p.x}" y="${height - 10}" text-anchor="middle" font-size="11" fill="#7b7b72">${txt}</text>`;
    });

    const dots = points.map(p => `<circle cx="${p.x}" cy="${p.y}" r="4" fill="#1a5e2a"></circle>`);

    svg.innerHTML = `
        <rect x="0" y="0" width="${width}" height="${height}" fill="transparent"></rect>
        ${grid.join('')}
        <polygon points="${area}" fill="rgba(26,94,42,0.10)"></polygon>
        <polyline points="${polyline}" fill="none" stroke="#1a5e2a" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"></polyline>
        ${dots.join('')}
        ${labels.join('')}
    `;

    if (legend) {
        const total = series.reduce((s, x) => s + x.value, 0);
        legend.textContent = `Ukupno za 7 dana: ${mgmtDashFmtInt(total)} kg`;
    }
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

    const plans = (window.dpPlans && Array.isArray(dpPlans))
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
                        <div class="mgmt-dash-dispatch-name">${mgmtDashEscape(p.VozacName || p.VozacID || '?')}</div>
                        <div class="mgmt-dash-dispatch-route">
                            ${mgmtDashEscape(p.StanicaName || p.StanicaID || '?')} → ${mgmtDashEscape(p.KupacName || p.KupacID || '?')}
                        </div>
                    </div>
                    <span class="mgmt-dash-status ${mgmtDashEscape(p.Status || 'planned')}">${mgmtDashEscape(mgmtDashStatusLabel(p.Status))}</span>
                </div>

                <div class="mgmt-dash-dispatch-meta">
                    <span><strong>${mgmtDashFmtInt(p.PlannedKg || 0)}</strong> kg</span>
                    <span>Demand: ${mgmtDashEscape(p.DemandID || '-')}</span>
                </div>
            </div>
        `)
        .join('');
}

async function mgmtRenderDashboard() {
    if (!window.mgmtData && typeof prefetchMgmtData === 'function') {
        try { await prefetchMgmtData(); } catch (e) {}
    }

    if (typeof loadDispecer === 'function' && (!window.dpPlans || !window.dpDem)) {
        try { await loadDispecer(); } catch (e) {}
    }

    const otkupiAll = (window.mgmtData && mgmtData.otkupiAll) ? mgmtData.otkupiAll : [];
    const saldoKupci = (window.mgmtData && mgmtData.saldoKupci) ? mgmtData.saldoKupci : [];
    const saldoOM = (window.mgmtData && mgmtData.saldoOM) ? mgmtData.saldoOM : [];
    const kartice = (window.mgmtData && mgmtData.kartice) ? mgmtData.kartice : [];

    let kgCeka = 0;
    if (typeof dpGetSup === 'function') {
        try {
            kgCeka = dpGetSup().reduce((s, r) => s + mgmtDashNum(r.Kolicina), 0);
        } catch (e) {}
    }

    let demandKg = 0;
    if (window.dpDem && Array.isArray(dpDem)) {
        demandKg = dpDem.reduce((s, d) => s + mgmtDashNum(d.Kg), 0);
    }

    const weekStats = mgmtDashGetWeekStats(otkupiAll);
    const series = mgmtDashBuild7DaySeries(otkupiAll);

    const koopSaldo = mgmtDashGetKoopSaldoFromKartice(kartice);
    const kupSaldo = saldoKupci.reduce((s, r) => s + mgmtDashNum(r.Saldo), 0);
    const omSaldo = saldoOM.reduce((s, r) => s + mgmtDashNum(r.Saldo), 0);

    mgmtDashSetText('mgmtDashWeekKg', mgmtDashFmtInt(weekStats.weekKg));
    mgmtDashSetText('mgmtDashActiveKoops', mgmtDashFmtInt(weekStats.activeKoops));
    mgmtDashSetText('mgmtDashAvgPrice', mgmtDashFmtDec(weekStats.avgPrice, 1));
    mgmtDashSetText('mgmtDashGGAP', '86%');

    const updatedEl = document.getElementById('mgmtDashUpdatedAt');
    if (updatedEl) {
        updatedEl.textContent = 'Ažurirano: ' + new Date().toLocaleTimeString('sr', {
            hour: '2-digit',
            minute: '2-digit'
        });
    }

    const alertItems = [];
    if (kgCeka > 0) alertItems.push({
        title: 'Roba čeka',
        value: `${mgmtDashFmtInt(kgCeka)} kg`,
        text: 'Postoje neraspoređene količine koje čekaju transport.'
    });
    if (demandKg > 0) alertItems.push({
        title: 'Otvoren demand',
        value: `${mgmtDashFmtInt(demandKg)} kg`,
        text: 'Postoji aktivna tražnja kupaca u sistemu.'
    });
    if (saldoKupci.some(r => mgmtDashNum(r.Saldo) > 0)) alertItems.push({
        title: 'Kupci sa saldom',
        text: 'Postoje kupci sa otvorenim finansijskim stanjem.'
    });
    if (saldoOM.some(r => mgmtDashNum(r.Saldo) > 0)) alertItems.push({
        title: 'Otkupna mesta sa saldom',
        text: 'Postoje mesta sa otvorenim saldom.'
    });

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
