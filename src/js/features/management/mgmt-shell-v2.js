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

    if (root === 'pregled') {
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
