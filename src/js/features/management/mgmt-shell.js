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
