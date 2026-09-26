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
        sidro: "    if (!crid) {\n      const err = new Error('Stavka bez ClientRecordID (' + oznaka + ')');",
        zamena: "    if (false) {   // sabotaza: stavka bez identiteta prolazi\n      const err = new Error('Stavka bez ClientRecordID (' + oznaka + ')');"
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
        // Prazan indeks je PRVI UPIS, ne konflikt. Sidro je zavrsni izlaz
        // uskladjivanja, pa sabotaza pogadja tacno degenerisan slucaj: tab u kom
        // ovaj dokument jos nema nijedan red.
        ime: 'otk-stavke-recovery-je-konflikt',
        suite: 'gas-otk-stavke',
        tvrdnja: 'prazan tab NIJE razlika',
        zasto: 'bez ovog izlaza bi svaki prvi upis odmah bio OTKUP_CONFLICT, pa nov otkup nikad ne bi prosao',
        fajl: GAS,
        sidro: "      if (!unos || unos.parent !== roditelj) {\n        return 'stavka ' + id + ' nije deo zavrsenog dokumenta';\n      }\n    }\n  }\n\n  return '';",
        zamena: "      if (!unos || unos.parent !== roditelj) {\n        return 'stavka ' + id + ' nije deo zavrsenog dokumenta';\n      }\n    }\n  }\n\n  return kljuceviTaba.length === 0 ? 'prazan tab' : '';   // sabotaza"
    },
    {
        // Medjujezicna granica: jedna zica, dva pisca, dva jezika, JEDAN raspored.
        ime: 'otk-stavke-raspored-kolona-drugaciji',
        suite: 'gas-otk-stavke',
        tvrdnja: 'ugovor kolona je DOSLOVNO isti kao u VBA',
        zasto: 'VBA citalac naslov taba poredi kolonu po kolonu i pada po imenu na prvu razliku, pa premestena kolona u GAS-u obara uvoz celog lista stanice',
        fajl: GAS,
        sidro: "  'OtkupStavkaID',\n  'OtkupID',",
        zamena: "  'OtkupID',\n  'OtkupStavkaID',   // sabotaza: zamenjen raspored"
    },
    {
        // P3 iz review-a #394. VBA citalac oba ova stanja odbija po imenu i
        // prekida uvoz celog lista; GAS je prvo preskakao a drugo tiho prepisivao.
        ime: 'otk-stavke-indeks-bez-identiteta-prolazi',
        suite: 'gas-otk-stavke',
        tvrdnja: 'indeks pada na red sa roditeljem a bez svog ClientRecordID',
        zasto: 'red koji GAS pusta a master odbija zaustavlja uvoz SVIH otkupa te stanice, ne samo spornog reda',
        fajl: GAS,
        sidro: "    if (!crid) {\n      const errBezId = new Error(",
        zamena: "    if (false) {   // sabotaza: red bez identiteta prolazi\n      const errBezId = new Error("
    },
    {
        ime: 'otk-stavke-indeks-dupli-prolazi',
        suite: 'gas-otk-stavke',
        tvrdnja: 'indeks pada na dva reda sa istim ClientRecordID stavke',
        zasto: 'tiho prepisivanje duplog item CRID-a sakriva red koji postoji u tabu, pa poredjenje skupa meri manje robe nego sto je upisano',
        fajl: GAS,
        sidro: "    if (izlaz[crid] !== undefined) {",
        zamena: "    if (false) {   // sabotaza: dupli item CRID se tiho prepisuje"
    },
    {
        // Review #394, drugi krug. Recovery sme SAMO pre completion marker-a; cim
        // zaglavlje postoji, skup je nepromenljiv. Bez te granice zavrsen -- cak i
        // Synced>Master -- dokument dobija novu stavku na retry-u, manifest ostaje
        // na starom broju, a master ga posle toga odbija kao neporavnat.
        ime: 'otk-stavke-zavrsen-prima-dopunu',
        suite: 'gas-otk-stavke',
        tvrdnja: 'zavrsen dokument NE PRIMA novu stavku',
        zasto: 'bez granice completion marker-a zavrsen otkup se mutira na retry-u: PWA misli da je ispravka prihvacena, master je nema, i razlika se ne vidi nigde',
        fajl: GAS,
        sidro: "  if (!dopustiDopunu) {",
        zamena: "  if (false) {   // sabotaza: zavrsen dokument prima dopunu"
    },
    {
        // Druga polovina NAMERNE ASIMETRIJE: dopuniti sto fali je dovrsavanje
        // upisa, ali zaboraviti sto u tabu postoji je druga tvrdnja o dokumentu.
        // Bez ovog izlaza bi klijent mogao da "skrati" dokument tihim izostavljanjem.
        ime: 'otk-stavke-zaboravljena-prolazi',
        suite: 'gas-otk-stavke',
        tvrdnja: 'stavka u tabu koju ulaz ne nosi JE razlika',
        zasto: 'stavka koja postoji u tabu a nije stigla znaci da klijent tvrdi manji dokument; tiho prihvatanje bi robu izbrisalo iz zaglavlja koje je vec upisano',
        fajl: GAS,
        sidro: "    if (ulaz[id] === undefined) {\n      return 'stavka ' + id + ' postoji u tabu a nije stigla';",
        zamena: "    if (false) {   // sabotaza: zaboravljena stavka prolazi\n      return 'stavka ' + id + ' postoji u tabu a nije stigla';"
    },
    {
        // P1 iz review-a #394. Preskakanje stavke po samom ID-u, bez poredjenja
        // sadrzaja, pravi hibridni dokument koji prolazi i manifest kapiju.
        ime: 'otk-stavke-partial-bez-poredjenja',
        suite: 'gas-otk-stavke',
        tvrdnja: 'partial upis sa izmenjenim sadrzajem JE konflikt, i kad zaglavlja nema',
        zasto: 'prvi pokusaj upise S1=100 i padne pre zaglavlja; retry posalje S1=120 i S2=60, pa u tabu ostane 100+60 -- skup koji nijedan klijent nije poslao',
        fajl: GAS,
        sidro: "    if (ulaz[id] !== unos.kljuc) {",
        zamena: "    if (false) {   // sabotaza: sadrzaj postojece stavke se ne poredi"
    },
    {
        // P2 iz review-a #394. Item CRID je globalan u VBA citaocu, pa GAS koji
        // gleda samo tekuceg roditelja prihvata stanje koje master odbija --
        // fail-closed nad CELIM listom stanice, ne nad jednim dokumentom.
        ime: 'otk-stavke-tudj-roditelj-prolazi',
        suite: 'gas-otk-stavke',
        tvrdnja: 'isti item CRID pod drugim otkupom JE konflikt',
        zasto: 'dva dokumenta koja dele identitet reda prolaze kroz GAS a obaraju ceo VBA uvoz te stanice',
        fajl: GAS,
        sidro: "    if (unos.parent !== roditelj) {",
        zamena: "    if (false) {   // sabotaza: tudj roditelj se ne razlikuje"
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
