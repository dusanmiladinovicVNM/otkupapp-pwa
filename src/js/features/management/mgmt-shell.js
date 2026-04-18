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
    // Legacy compatibility only.
    // Novi MGMT shell koristi showMgmtRoot(...) iz mgmt-shell-v2.js
    if (typeof showMgmtRoot === 'function') {
        const map = {
            kooperanti: 'partneri',
            stanice: 'otkup',
            kupci: 'partneri',
            agrohemija: 'agro'
        };

        const targetRoot = map[section] || 'pregled';
        showMgmtRoot(targetRoot, btn);

        if (targetRoot === 'partneri') {
            if (section === 'kooperanti' && typeof showMgmtPartnerSegment === 'function') {
                showMgmtPartnerSegment('kooperanti');
            }
            if (section === 'kupci' && typeof showMgmtPartnerSegment === 'function') {
                showMgmtPartnerSegment('kupci');
            }
        }
        return;
    }
}

function showMgmtSub(subId, btn) {
    // Legacy compatibility only.
    if (typeof showMgmtRoot !== 'function') return;

    const koopMap = {
        'koop-pregled': 'pregled',
        'koop-kartica': 'kartica',
        'koop-saldo': 'saldo'
    };

    const kupMap = {
        'kup-fakture': 'fakture',
        'kup-saldo': 'saldo',
        'kup-roba': 'roba'
    };

    const otkupMap = {
        'sta-otkupi': 'otkupi',
        'sta-saldo': 'saldo',
        'sta-roba': 'roba'
    };

    const agroMap = {
        'agro-izdavanje': 'izdavanje',
        'agro-stanje': 'stanje'
    };

    if (koopMap[subId]) {
        showMgmtRoot('partneri');
        if (typeof showMgmtPartnerSegment === 'function') showMgmtPartnerSegment('kooperanti');
        if (typeof showMgmtKoopSub === 'function') showMgmtKoopSub(koopMap[subId], btn);
        return;
    }

    if (kupMap[subId]) {
        showMgmtRoot('partneri');
        if (typeof showMgmtPartnerSegment === 'function') showMgmtPartnerSegment('kupci');
        if (typeof showMgmtKupSub === 'function') showMgmtKupSub(kupMap[subId], btn);
        return;
    }

    if (otkupMap[subId]) {
        showMgmtRoot('otkup');
        if (typeof showMgmtOtkupSub === 'function') showMgmtOtkupSub(otkupMap[subId], btn);
        return;
    }

    if (agroMap[subId]) {
        showMgmtRoot('agro');
        if (typeof showMgmtAgroSub === 'function') showMgmtAgroSub(agroMap[subId], btn);
    }
}

document.addEventListener('DOMContentLoaded', function() {
    const oldBar = document.getElementById('mgmtSubBar');
    if (oldBar) oldBar.style.display = 'none';
});
