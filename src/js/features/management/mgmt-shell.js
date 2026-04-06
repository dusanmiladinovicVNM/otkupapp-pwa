// ============================================================
// MANAGEMENT SUB-TAB DEFINITIONS
// ============================================================
const MGMT_SUBS = {
    kooperanti: [
        { id: 'koop-kartica', label: 'Kartica', load: null },
        { id: 'koop-saldo', label: 'Saldo', load: loadMgmtKoopSaldo },
        { id: 'koop-pregled', label: 'Pregled', load: loadMgmtKoopPregled }
    ],
    stanice: [
        { id: 'sta-otkupi', label: 'Otkupi', load: null },
        { id: 'sta-saldo', label: 'Saldo OM', load: loadMgmtSaldoOM },
        { id: 'sta-roba', label: 'Roba', load: loadMgmtOtkupPoOM }
    ],
    kupci: [
        { id: 'kup-fakture', label: 'Fakture', load: null },
        { id: 'kup-saldo', label: 'Saldo', load: loadMgmtKupci },
        { id: 'kup-roba', label: 'Roba', load: loadMgmtPredato }
    ],
    agrohemija: [
        { id: 'agro-izdavanje', label: 'Izdavanje', load: function() { populateIzdDropdowns(); } },
        { id: 'agro-stanje', label: 'Stanje', load: loadMgmtAgroStanje }
    ]
};

// ============================================================
// PREFETCH
// ============================================================
async function prefetchMgmtData() {
    try {
        const json = await apiFetch('action=getMgmtAll');
        if (json && json.success) { mgmtData = json; }
    } catch (e) {}
}

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

// ============================================================
// MANAGEMENT: NAVIGATION
// ============================================================
function showMgmtMain(section, btn) {
    qsa('.tab-content').forEach(t => removeClass(t, 'active'));
    qsa('.tab-btn').forEach(b => removeClass(b, 'active'));

    addClass(byId('tab-mgmt'), 'active');

    if (btn && btn.classList) {
        addClass(btn, 'active');
    }

    const subs = MGMT_SUBS[section];
    if (!subs) return;

    const bar = byId('mgmtSubBar');
    setHtml(bar, subs.map((s, i) =>
        `<button class="sub-tab-btn${i === 0 ? ' active' : ''}" onclick="showMgmtSub('${s.id}', this)">${s.label}</button>`
    ).join(''));

    qsa('.mgmt-sub').forEach(s => removeClass(s, 'active'));
    const firstEl = byId('mgmt-' + subs[0].id);
    if (firstEl) addClass(firstEl, 'active');
    if (subs[0].load) subs[0].load();
}

function showMgmtSub(subId, btn) {
    qsa('.mgmt-sub').forEach(s => removeClass(s, 'active'));
    qsa('.sub-tab-btn').forEach(b => removeClass(b, 'active'));

    const el = byId('mgmt-' + subId);
    if (el) addClass(el, 'active');
    if (btn) addClass(btn, 'active');

    const allSubs = Object.values(MGMT_SUBS).flat();
    const sub = allSubs.find(s => s.id === subId);
    if (sub && sub.load) sub.load();
}
