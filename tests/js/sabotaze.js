'use strict';

// ============================================================
// KATALOG SABOTAZA -- dvosmerni dokaz za JS sloj
// ============================================================
//
// Zelena suite koja nikad nije pokazana crvena ne dokazuje da isto meri
// (CLAUDE.md S5). Zato svaka tvrdnja ovde ima unos koji kvari TACNO onu
// produkcionu odluku koju ta tvrdnja opisuje, i self-test proverava da tvrdnja
// tada pukne PO IMENU.
//
// Ista sema kao tools/dokaz.py: (fajl, sidro, zamena, test). Razlika je sto se
// ovde nista ne upisuje na disk -- zamena se primenjuje na tekst u memoriji pre
// izvrsavanja, pa dokaz sme da se vrti paralelno sa izmenama.

const DB = 'src/js/services/db.js';
const ZBR = 'src/js/features/vozac/zbirna.js';
const GAS = 'gas/Code.gs';

module.exports = [
    {
        ime: 'claim-ne-prekida-transakciju',
        suite: 'db-claim',
        tvrdnja: 'claim odbija i NE upisuje kad provera vrati razlog',
        zasto: 'bez abort-a transakcija se zavrsi uspesno, pa odbijen claim izgleda kao prihvacen',
        fajl: DB,
        sidro: 'try { tx.abort(); } catch (_) {}',
        zamena: 'try { /* sabotaza: nema abort-a */ } catch (_) {}'
    },
    {
        ime: 'claim-citanje-tiho-prolazi',
        suite: 'db-claim',
        tvrdnja: 'claim BACA kad citanje padne, ne vraca ok:false',
        zasto: '"ne znam stanje" pretvoreno u "stanje dozvoljava" je lazan uspeh',
        fajl: DB,
        sidro: [
            '                req.onerror = function (event) {',
            '                    reject(event && event.target ? event.target.error',
            "                                                 : new Error('dbClaimInStore read failed'));",
            '                };'
        ].join('\n'),
        zamena: [
            '                req.onerror = function (event) {',
            "                    resolve({ ok: true, razlog: '' });   // sabotaza",
            '                };'
        ].join('\n')
    },
    {
        ime: 'rezervacija-ne-cita-master',
        suite: 'zbirna-rezervacije',
        tvrdnja: 'zbirna koju je master razresio ne drzi nista',
        zasto: 'rezervacija bez masterovog verdikta traje zauvek -- posle storna se otpremnica nikad ne vraca u izbor',
        fajl: ZBR,
        sidro: '        if (zbirnaRazresenaOdMastera(z)) return;',
        zamena: '        if (false) return;   // sabotaza'
    },
    {
        ime: 'rezervacija-pusta-transportni-pad',
        suite: 'zbirna-rezervacije',
        tvrdnja: 'transportni pad ZADRZAVA rezervaciju',
        zasto: 'ako pending sam po sebi oslobadja, prekid mreze pusta istu otpremnicu u drugu zbirnu',
        fajl: ZBR,
        sidro: "        if (ZBIRNA_TRAJNO_ODBIJENA.has(String(z.lastServerCode || '').trim())) return;",
        zamena: "        if (String(z.syncStatus || '').trim() === 'pending') return;   // sabotaza"
    },
    {
        ime: 'razresenje-ne-vidi-syncerror',
        suite: 'zbirna-rezervacije',
        tvrdnja: 'terminalni statusi mastera su prepoznati doslovno',
        zasto: 'SyncError je masterova odluka; ako se ne prizna, odbijena zbirna drzi otpremnice zauvek',
        fajl: ZBR,
        sidro: "    return s === 'Synced>Master' || s === 'Duplicate' || s.indexOf('SyncError') === 0;",
        zamena: "    return s === 'Synced>Master' || s === 'Duplicate';   // sabotaza"
    },
    {
        ime: 'skup-izvora-gleda-redosled',
        suite: 'gas-skup-izvora',
        tvrdnja: 'isti skup u drugom redosledu NIJE razlika',
        zasto: 'bez sortiranja isti manifest u drugom redu postaje ZBIRNA_CONFLICT, pa retry izgleda kao greska',
        fajl: GAS,
        sidro: '      .filter(Boolean)\n      .sort();',
        zamena: '      .filter(Boolean);   // sabotaza'
    }
];
