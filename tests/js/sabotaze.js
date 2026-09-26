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
const OTKS = 'src/js/features/otkup/otkup-stavke.js';

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
        // Centralna tvrdnja reza mora da ima svoj fault seam (review #393, P3).
        //
        // SABOTAZA JE LEGALNA, SEMANTIKA JE POKVARENA -- i to je cela poenta
        // (review #393, drugi krug). Prva verzija je odlagala store.put u
        // makrotask, ali taj `store` pripada transakciji koja se dotad zavrsila,
        // pa je ishod bio TransactionInactiveError iz timera -- izuzetak IZVAN
        // try/catch oko tvrdnje, koji je ubijao ceo prolaz. To nije dokaz nego
        // pad harness-a.
        //
        // Ovde upis ide u SVOJU, sasvim ispravnu readwrite transakciju. Time se
        // kvari tacno ona invarijanta koju tvrdnja meri -- provera i upis nisu
        // vise jedan potez:
        //
        //   TX A: read -> slobodno, commit          TX B: read -> slobodno, commit
        //   TX A': write                            TX B': write
        //
        // pa obe konekcije jave uspeh nad istom otpremnicom.
        ime: 'claim-upis-u-drugoj-transakciji',
        suite: 'db-claim',
        tvrdnja: 'dve konekcije nad istom bazom: tacno jedan claim prolazi',
        zasto: 'provera i upis u dva poteza nije kapija nego nada -- oba taba prodju',
        fajl: DB,
        sidro: '                    store.put(record);',
        zamena: ('                    var writeTx = db.transaction(storeName, ' +
                 "'readwrite');   // sabotaza\n" +
                 '                    writeTx.objectStore(storeName).put(record);')
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
    },
    {
        // Prazan skup stavki je greska, ne dokument od nula kilograma.
        ime: 'otk-stavke-prazan-skup-prolazi',
        suite: 'gas-otk-stavke',
        tvrdnja: 'otkup bez stavki se odbija PO IMENU',
        zasto: 'bez ove kapije zaglavlje bez stavki stize u master kao siroce koje sa terena nema kako da se popravi',
        fajl: GAS,
        sidro: "  if (!Array.isArray(stavke) || stavke.length === 0) {",
        zamena: "  if (false) {   // sabotaza: prazan skup prolazi"
    },
    {
        // Identitet reda je jedini nacin da retry ne udvoji robu.
        ime: 'otk-stavke-bez-identiteta-prolazi',
        suite: 'gas-otk-stavke',
        tvrdnja: 'stavka bez svog ClientRecordID se odbija',
        zasto: 'bez identiteta reda ponovljen sync ne razlikuje retry od druge stavke, pa se kolicina udvaja',
        fajl: GAS,
        sidro: "    if (!crid) {",
        zamena: "    if (false) {   // sabotaza: stavka bez identiteta prolazi"
    },
    {
        // Redosled NIJE tvrdnja o dokumentu: master dodeljuje RedniBroj po
        // kanonskom redu klasa.
        ime: 'otk-stavke-razlika-po-poziciji',
        suite: 'gas-otk-stavke',
        tvrdnja: 'iste stavke u drugom redosledu NISU razlika',
        zasto: 'poredjenje po poziciji pretvara retry u OTKUP_CONFLICT, pa klijent dobije gresku na ispravan zapis',
        fajl: GAS,
        sidro: "    if (ulaz[id] === undefined) {",
        zamena: "    if (ulaz[id] === undefined || Object.keys(ulaz)[k] !== id) {   // sabotaza"
    },
    {
        // Prazan tab je PRVI UPIS, ne konflikt.
        ime: 'otk-stavke-prazan-tab-je-konflikt',
        suite: 'gas-otk-stavke',
        tvrdnja: 'prazan tab NIJE razlika',
        zasto: 'bez ovog izlaza bi svaki prvi upis odmah bio OTKUP_CONFLICT, pa nov otkup nikad ne bi prosao',
        fajl: GAS,
        sidro: "  if (!uTabu || Object.keys(uTabu).length === 0) return '';",
        zamena: "  if (!uTabu) return '';   // sabotaza: prazan tab je razlika"
    },
    {
        // Vrednost dokumenta je zbir PO STAVCI. Proizvod zbirova (460 * 80) daje
        // 36800 umesto 21800 -- razlika izlazi na iznos za isplatu.
        ime: 'vrednost-kao-proizvod-zbirova',
        suite: 'otkup-stavke',
        tvrdnja: 'dve klase: vrednost je zbir PO STAVCI, ne proizvod zbirova',
        zasto: 'proizvod zbirova mnozi ukupne kilograme ukupnom cenom, pa dvoklasni dokument dobija iznos koji nijedna stavka ne tvrdi',
        fajl: OTKS,
        sidro: "        return z + Number(s && s.kolicina || 0) * Number(s && s.cena || 0);",
        zamena: "        return z + Number(s && s.kolicina || 0);   // sabotaza: cena ispada"
    },
    {
        // Dvoklasni dokument nema jednu cenu; prva cena bi bila pogresan podatak.
        ime: 'cena-prve-stavke-za-sve',
        suite: 'otkup-stavke',
        tvrdnja: 'dve klase NEMAJU jednu cenu',
        zasto: 'ako prikaz uzme prvu cenu, racun na dvoklasnom dokumentu tvrdi cenu koja ne vazi za svu robu',
        fajl: OTKS,
        sidro: "    if (stavke.length !== 1) return null;",
        zamena: "    if (stavke.length === 0) return null;   // sabotaza: prva cena za sve"
    },
    {
        // Oznaka klasa mora da ide kanonskim redom -- isto pravilo po kom master
        // dodeljuje RedniBroj. Inace se oznaka menja od redosleda sinhronizacije.
        ime: 'klase-po-redu-stizanja',
        suite: 'otkup-stavke',
        tvrdnja: 'oznaka klasa ide kanonskim redom, ne redom stizanja',
        zasto: 'bez kanonskog reda isti dokument dobija razlicitu oznaku klase zavisno od redosleda kojim su stavke stigle',
        fajl: OTKS,
        sidro: "            return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib);",
        zamena: "            return 0;   // sabotaza: stabilan sort zadrzava red stizanja"
    }
];
