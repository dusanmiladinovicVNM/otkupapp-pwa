'use strict';

// ============================================================
// OTK_STAVKE u GAS-u -- skup stavki je dokument (S5-5b)
// ============================================================
//
// Zaglavlje OTK lista od ovog reza ne nosi klasu, kolicinu, cenu ni ambalazu.
// Nosi MANIFEST. Zato GAS mora da razlikuje stvari koje su do sada bile jedna:
// "nema stavki", "iste stavke u drugom redosledu", "druga tvrdnja o istom
// dokumentu", "nedovrsen upis" i "tudji identitet reda".
//
// Mere se CISTI pomocnici: harness ne moze da podigne SpreadsheetApp (stubovi
// bacaju izuzetak sa imenom clana), pa I/O ostaje tanak a odluka merena.
//
// Prvi krug review-a #394 je prosao TACNO IZMEDJU starih tvrdnji: nijedna nije
// merila partial upis (retry kad zaglavlje jos ne postoji) ni koliziju item
// CRID-a izmedju dva dokumenta. Zato su te dve dodate imenom i obrazlozenjem.

const assert = require('node:assert');
const { ucitaj, procitaj } = require('../harness');

const FAJL = 'gas/Code.gs';

async function ucitajModul(opcije) {
    const ctx = ucitaj([FAJL], opcije || {});

    for (const ime of ['otkStavkeNormalizuj_', 'otkStavkeUskladiNedovrsen_',
                       'otkStavkeUskladiZavrsen_', 'otkStavkeIndeksIzRedova_',
                       'otkStavkeKoloneUgovor_', 'otkStavkaKljuc_',
                       'otkupZaglavljeRazlika_', 'otkupZaglavljePoljaSadrzaja_',
                       'otkupZaglavljeKoloneUgovor_', 'buildOtkupMergeKey_']) {
        assert.strictEqual(typeof ctx[ime], 'function',
            ime + ' nije vidljiva posle ucitavanja ' + FAJL);
    }

    return {
        normalizuj: ctx.otkStavkeNormalizuj_,

        // DVA IMENOVANA UGOVORA, ne jedan bez konteksta (review #394, drugi krug):
        // zaglavlje je completion marker, pa dopuna sme samo dok ga nema.
        nedovrsen: ctx.otkStavkeUskladiNedovrsen_,
        zavrsen: ctx.otkStavkeUskladiZavrsen_,

        // Cist deo indeksa taba: pravila fail-closed se mere bez SpreadsheetApp.
        indeksIzRedova: ctx.otkStavkeIndeksIzRedova_,

        // Kroz FUNKCIJU, ne kroz const: top-level const u vm kontekstu ne postaje
        // svojstvo globalnog objekta, pa bi `ctx.OTK_STAVKE_COLUMNS` bio undefined.
        kolone: ctx.otkStavkeKoloneUgovor_(),

        // Kanonski sadrzaj ZAGLAVLJA -- druga polovina istog identiteta.
        zaglavljeRazlika: ctx.otkupZaglavljeRazlika_,
        zaglavljePolja: ctx.otkupZaglavljePoljaSadrzaja_(),
        zaglavljeKolone: ctx.otkupZaglavljeKoloneUgovor_(),

        kljuc: ctx.otkStavkaKljuc_,
        mergeKljuc: ctx.buildOtkupMergeKey_
    };
}

const O1 = 'CRID-Z1';
const O2 = 'CRID-Z2';

// Jedna ispravna stavka; test menja tacno ono sto meri.
function stavka(nad) {
    return Object.assign({
        clientRecordID: 'CRID-S1',
        klasa: 'I',
        kolicina: 100,
        cena: 50,
        kolAmbalaze: 4
    }, nad || {});
}

// Cinjenice zaglavlja jednog ispravnog dokumenta; test menja tacno jedno polje.
const ZAGL = {
    OtkupacID: 'ST-1',
    Datum: '2026-09-26',
    KooperantID: 'K1',
    VrstaVoca: 'Malina',
    SortaVoca: 'Willamette',
    ParcelaID: 'P1',
    TipAmbalaze: '6/1',
    StavkeCount: 1
};

// Red zaglavlja u rasporedu UGOVORA + indeks po imenu, kako ih processRecord vidi.
// Raspored dolazi iz produkcije, pa preimenovana kolona ne moze da prodje.
function zaglavljeRed(kolone, polja) {
    const idx = {};
    kolone.forEach(function (k, i) { idx[k] = i; });

    return {
        red: kolone.map(function (k) {
            return polja[k] !== undefined ? polja[k] : '';
        }),
        idx: idx
    };
}

// Ugovor kolona PROCITAN IZ VBA IZVORA -- druga strana iste zice.
//
// `OtkStavkeKolone` u modMasterSync nabraja imena konstanti, a njihove vrednosti
// su u modConfig (COL_OKS_*) i u modMasterSync (OKS_WIRE_*), pa se razresavaju
// odavde. VBA citalac naslov poredi kolonu po kolonu i pada po imenu na prvu
// razliku, zato je i redosled deo ugovora.
function vbaUgovorKolona() {
    const izvori = [
        procitaj('src-vba/modMasterSync.bas'),
        procitaj('src-vba/modConfig.bas')
    ].join('\n');

    const konstante = {};
    const rxConst = /(?:Public|Private)\s+Const\s+(\w+)\s+As\s+String\s*=\s*"([^"]*)"/g;
    let m;
    while ((m = rxConst.exec(izvori)) !== null) {
        konstante[m[1]] = m[2];
    }

    const telo = /OtkStavkeKolone\s*=\s*Array\(([^)]*)\)/.exec(izvori);
    assert.ok(telo, 'OtkStavkeKolone = Array(...) nije nadjen u modMasterSync.bas');

    return telo[1]
        .replace(/_\s*\n/g, ' ')
        .split(',')
        .map(s => s.trim())
        .filter(Boolean)
        .map(function (ime) {
            assert.notStrictEqual(konstante[ime], undefined,
                'VBA konstanta bez vrednosti: ' + ime);
            return konstante[ime];
        });
}

// Red taba OTK_STAVKE, u rasporedu ugovora. Prazno polje = prazna celija, tacno
// kao sto ga pisac ostavi: push red nema wire kolone, PWA red nema OtkupStavkaID.
function wireRed(kolone, polja) {
    return kolone.map(function (k) {
        return polja[k] !== undefined ? polja[k] : '';
    });
}

// Indeks taba kakav ga otkStavkeIndeksTaba_ gradi: item CRID -> {parent, kljuc}.
// GLOBALAN je, pa fixture mora da nosi roditelja po redu -- bas to je granica
// koju P2 iz review-a #394 trazi.
function indeks(kljucFn, redovi) {
    const m = {};
    for (const [id, parent, s] of redovi) {
        m[id] = {
            parent: parent,
            kljuc: kljucFn(s.klasa, s.kolicina, s.cena, s.kolAmbalaze, s.brutoKg || 0)
        };
    }
    return m;
}

module.exports = {
    ime: 'gas-otk-stavke',
    ucitaj: ucitajModul,
    tvrdnje: {

        'otkup bez stavki se odbija PO IMENU': async function (m) {
            assert.throws(() => m.normalizuj([], O1), /Stavke missing/,
                'prazan skup je prosao -- dokument od nula kilograma');
            assert.throws(() => m.normalizuj(undefined, O1), /Stavke missing/,
                'nedostajuce stavke su prosle kao prazan skup');
        },

        'stavka bez svog ClientRecordID se odbija': async function (m) {
            // Bez identiteta reda se ponovljen sync ne razlikuje od druge stavke,
            // pa bi retry posle mreznog pada udvajao robu.
            assert.throws(
                () => m.normalizuj([stavka({ clientRecordID: '' })], O1),
                /Stavka bez ClientRecordID/,
                'stavka bez identiteta je prosla');
        },

        'dve stavke iste klase se odbijaju': async function (m) {
            // Bas bug koji header+stavke uklanja: jedan logicki dokument rasut
            // po redovima.
            assert.throws(
                () => m.normalizuj([
                    stavka({ clientRecordID: 'CRID-S1' }),
                    stavka({ clientRecordID: 'CRID-S2' })
                ], O1),
                /Dve stavke iste klase/,
                'dve stavke iste klase su prosle');
        },

        'dve stavke sa istim ClientRecordID se odbijaju': async function (m) {
            assert.throws(
                () => m.normalizuj([
                    stavka({ clientRecordID: 'CRID-S1', klasa: 'I' }),
                    stavka({ clientRecordID: 'CRID-S1', klasa: 'II' })
                ], O1),
                /istim ClientRecordID/,
                'dva reda istog identiteta su prosla kao dve stavke');
        },

        'normalizacija nosi redni broj po redu stizanja i bruto kad postoji':
        async function (m) {
            const out = m.normalizuj([
                stavka({ clientRecordID: 'CRID-S1', klasa: 'II', kolicina: 60 }),
                stavka({ clientRecordID: 'CRID-S2', klasa: 'I', brutoKg: 110 })
            ], O1);

            assert.strictEqual(out.length, 2, 'obe stavke nisu prosle');
            assert.strictEqual(out[0].redniBroj, 1, 'prva stavka nije dobila redni broj 1');
            assert.strictEqual(out[1].redniBroj, 2, 'druga stavka nije dobila redni broj 2');
            assert.strictEqual(out[0].otkupClientRecordID, O1,
                'stavka ne nosi CRID svog zaglavlja');
            assert.strictEqual(out[1].brutoKg, 110, 'bruto sa zice nije prenet');
        },

        // REDOSLED NIJE TVRDNJA O DOKUMENTU. Master dodeljuje RedniBroj po
        // kanonskom redu klasa, pa je redosled kojim je PWA slala nebitan.
        'iste stavke u drugom redosledu NISU razlika': async function (m) {
            const a = stavka({ clientRecordID: 'CRID-S1', klasa: 'I' });
            const b = stavka({ clientRecordID: 'CRID-S2', klasa: 'II', kolicina: 60 });

            const tab = indeks(m.kljuc, [['CRID-S1', O1, a], ['CRID-S2', O1, b]]);
            const stigle = m.normalizuj([b, a], O1);

            assert.strictEqual(m.nedovrsen(tab, stigle, O1), '',
                'drugi redosled je prijavljen kao konflikt -- retry bi postao greska');
        },

        'izmenjena kolicina JE razlika i imenuje stavku': async function (m) {
            const a = stavka({ clientRecordID: 'CRID-S1' });
            const tab = indeks(m.kljuc, [['CRID-S1', O1, a]]);
            const stigle = m.normalizuj(
                [stavka({ clientRecordID: 'CRID-S1', kolicina: 999 })], O1);

            const r = m.nedovrsen(tab, stigle, O1);
            assert.notStrictEqual(r, '', 'izmenjena kolicina nije prijavljena kao razlika');
            assert.match(String(r), /CRID-S1/, 'razlika ne imenuje koja stavka');
        },

        'stavka u tabu koju ulaz ne nosi JE razlika': async function (m) {
            const a = stavka({ clientRecordID: 'CRID-S1', klasa: 'I' });
            const b = stavka({ clientRecordID: 'CRID-S2', klasa: 'II', kolicina: 60 });

            const tab = indeks(m.kljuc, [['CRID-S1', O1, a], ['CRID-S2', O1, b]]);
            const stigle = m.normalizuj([a], O1);

            assert.match(String(m.nedovrsen(tab, stigle, O1)), /CRID-S2/,
                'zaboravljena stavka nije prijavljena kao razlika');
        },

        // PRAZAN TAB JE PRVI UPIS, ne razlika. Bez ovoga nijedan nov otkup ne bi
        // mogao da prodje -- svaki bi odmah bio OTKUP_CONFLICT.
        'prazan tab NIJE razlika': async function (m) {
            const stigle = m.normalizuj([stavka()], O1);
            assert.strictEqual(m.nedovrsen({}, stigle, O1), '',
                'prazan tab je prijavljen kao konflikt -- prvi upis bi pao');
        },

        // ============================================================
        // P1 iz review-a #394: PARTIAL UPIS + IZMENJEN RETRY
        // ============================================================
        //
        // Prvi pokusaj upise S1=100 i padne PRE zaglavlja. Retry posalje S1=120 i
        // S2=60. Dok je provera stajala samo u grani "zaglavlje postoji", ovaj put
        // je prolazio bez poredjenja sadrzaja: S1 se preskoci po ID-u, S2 se
        // dopise, zaglavlje kaze 2 -- i master uveze 100+60, skup koji NIJEDAN
        // klijent nije poslao. Manifest se pritom poklapa, pa ga ni VBA ne hvata.
        'partial upis sa izmenjenim sadrzajem JE konflikt, i kad zaglavlja nema':
        async function (m) {
            const prvi = stavka({ clientRecordID: 'CRID-S1', klasa: 'I', kolicina: 100 });
            const tab = indeks(m.kljuc, [['CRID-S1', O1, prvi]]);

            const retry = m.normalizuj([
                stavka({ clientRecordID: 'CRID-S1', klasa: 'I', kolicina: 120 }),
                stavka({ clientRecordID: 'CRID-S2', klasa: 'II', kolicina: 60 })
            ], O1);

            const r = m.nedovrsen(tab, retry, O1);
            assert.notStrictEqual(r, '',
                'partial upis sa izmenjenim sadrzajem je prosao -- nastao bi hibridni dokument');
            assert.match(String(r), /CRID-S1/, 'konflikt ne imenuje stavku koja se razlikuje');
        },

        // ============================================================
        // P2 iz review-a #394: ITEM CRID JE GLOBALAN
        // ============================================================
        //
        // VBA citalac drzi JEDAN skup vidjenih item CRID-ova za ceo tab i dva reda
        // istog ClientRecordID-a odbija bez obzira na roditelja. Dok je GAS gledao
        // samo redove tekuceg roditelja, prihvatao je stanje koje VBA kasnije
        // odbija -- i to fail-closed nad CELIM listom stanice, pa je jedan
        // pogresan item CRID zaustavljao uvoz svih otkupa te stanice.
        'isti item CRID pod drugim otkupom JE konflikt': async function (m) {
            const a = stavka({ clientRecordID: 'ITEM-X', klasa: 'I' });

            // ITEM-X je u tabu, ali kao stavka otkupa O1.
            const tab = indeks(m.kljuc, [['ITEM-X', O1, a]]);

            // Sada isti ITEM-X stize kao stavka otkupa O2, sa ISTIM sadrzajem --
            // pa ga poredjenje sadrzaja samo po sebi ne bi uhvatilo.
            const stigle = m.normalizuj([stavka({ clientRecordID: 'ITEM-X', klasa: 'I' })], O2);

            const r = m.nedovrsen(tab, stigle, O2);
            assert.notStrictEqual(r, '',
                'isti item CRID pod drugim otkupom je prosao -- VBA bi ga odbio i oborio ceo list');
            assert.match(String(r), /ITEM-X/, 'konflikt ne imenuje stavku');
            assert.match(String(r), new RegExp(O1), 'konflikt ne imenuje tudjeg roditelja');
        },

        // ASIMETRIJA JE NAMERNA, ALI SAMO PRE COMPLETION MARKER-A: dopisati sto
        // fali je dovrsavanje istog upisa, zaboraviti sto postoji je druga tvrdnja
        // o dokumentu. Cim zaglavlje postoji, ni dopuna nije dozvoljena -- to meri
        // tvrdnja 'zavrsen dokument NE PRIMA novu stavku'.
        'nedovrsen upis: stavka koja u tabu fali je RECOVERY': async function (m) {
            const a = stavka({ clientRecordID: 'CRID-S1', klasa: 'I' });
            const tab = indeks(m.kljuc, [['CRID-S1', O1, a]]);

            const stigle = m.normalizuj([
                stavka({ clientRecordID: 'CRID-S1', klasa: 'I' }),
                stavka({ clientRecordID: 'CRID-S2', klasa: 'II', kolicina: 60 })
            ], O1);

            assert.strictEqual(m.nedovrsen(tab, stigle, O1), '',
                'dopuna nedostajuce stavke je prijavljena kao konflikt -- prekinut upis se ne bi mogao dovrsiti');
        },

        // Stavke TUDJEG otkupa u tabu ne smeju da smetaju: tab je zajednicki, pa
        // bi inace svaki drugi dokument na listu pravio laznu razliku.
        'stavke drugog otkupa ne ulaze u poredjenje': async function (m) {
            const moja = stavka({ clientRecordID: 'CRID-S1', klasa: 'I' });
            const tudja = stavka({ clientRecordID: 'CRID-T1', klasa: 'I' });

            const tab = indeks(m.kljuc, [['CRID-S1', O1, moja], ['CRID-T1', O2, tudja]]);
            const stigle = m.normalizuj([stavka({ clientRecordID: 'CRID-S1', klasa: 'I' })], O1);

            assert.strictEqual(m.nedovrsen(tab, stigle, O1), '',
                'stavka drugog otkupa je prijavljena kao razlika');
        },

        // ============================================================
        // Review #394, drugi krug: ZAGLAVLJE JE COMPLETION MARKER
        // ============================================================
        //
        // Recovery pravilo je vazilo i za VEC ZAVRSEN dokument, jer je kapija bila
        // digunta iznad grananja na existingRow. Zavrsen otkup je tako na retry-u
        // dobijao novu stavku: GAS vrati success/existing, manifest ostane na
        // starom broju, a master isti dokument posle toga odbija kao neporavnat.
        // PWA misli da je ispravka prihvacena, master je nema, razlika se ne vidi.
        //
        // Gore od toga: to se desavalo PRE citanja terminalnog statusa, pa je i
        // Synced>Master dokument mogao da dobije stavku.
        'zavrsen dokument NE PRIMA novu stavku': async function (m) {
            const s1 = stavka({ clientRecordID: 'CRID-S1', klasa: 'I' });
            const tab = indeks(m.kljuc, [['CRID-S1', O1, s1]]);

            const retry = m.normalizuj([
                stavka({ clientRecordID: 'CRID-S1', klasa: 'I' }),
                stavka({ clientRecordID: 'CRID-S2', klasa: 'II', kolicina: 60 })
            ], O1);

            const r = m.zavrsen(tab, retry, O1);
            assert.notStrictEqual(r, '',
                'zavrsen dokument je primio novu stavku -- manifest bi ostao na starom broju');
            assert.match(String(r), /CRID-S2/, 'konflikt ne imenuje stavku koja se dodaje');

            // ISTI ULAZ pre completion marker-a MORA da prodje, inace tvrdnja meri
            // "sve se odbija" umesto "zavrsen dokument je nepromenljiv".
            assert.strictEqual(m.nedovrsen(tab, retry, O1), '',
                'isti ulaz je odbijen i kad zaglavlja nema -- prekinut upis se ne bi mogao dovrsiti');
        },

        // KONTROLA: doslovno isti payload nad zavrsenim dokumentom je idempotentan.
        // Bez nje bi gornja tvrdnja bila zelena i da ugovor odbija svaki retry.
        'zavrsen dokument: doslovno isti skup je idempotentan': async function (m) {
            const s1 = stavka({ clientRecordID: 'CRID-S1', klasa: 'I' });
            const s2 = stavka({ clientRecordID: 'CRID-S2', klasa: 'II', kolicina: 60 });

            const tab = indeks(m.kljuc, [['CRID-S1', O1, s1], ['CRID-S2', O1, s2]]);
            const stigle = m.normalizuj([s2, s1], O1);   // i u drugom redosledu

            assert.strictEqual(m.zavrsen(tab, stigle, O1), '',
                'ponovljen sync istog dokumenta je prijavljen kao konflikt');
        },

        // JEDNA ZICA, DVA PISCA, DVA JEZIKA -- i jedan raspored kolona.
        //
        // Bez ove tvrdnje je "doslovno isti ugovor" bio komentar. VBA citalac
        // naslov taba poredi kolonu po kolonu i pada po imenu na prvu razliku, pa
        // bi preimenovana ili premestena kolona u GAS-u oborila uvoz CELOG lista
        // stanice -- a nijedna kapija to ne bi javila pre produkcije.
        'ugovor kolona je DOSLOVNO isti kao u VBA': async function (m) {
            assert.deepStrictEqual(m.kolone, vbaUgovorKolona(),
                'raspored kolona OTK_STAVKE se razlikuje od modMasterSync.OtkStavkeKolone');
        },

        // ============================================================
        // Review #394, P3: INDEKS TABA PADA PO IMENU, kao VBA citalac
        // ============================================================
        //
        // Dok je GAS red bez item CRID-a preskakao a dupli CRID tiho prepisivao,
        // isti tab je za GAS bio ispravan a za master neispravan -- dva citaoca
        // istog ugovora sa razlicitom strogoscu.
        'indeks preskace red desktop push-a': async function (m) {
            // Push red ima svoj OtkupStavkaID i pravi OtkupID, a wire kolone su mu
            // prazne. Ne pripada imenskom prostoru item CRID-ova.
            const redovi = [
                wireRed(m.kolone, {
                    OtkupStavkaID: 'OKS-1', OtkupID: 'OTK-1', RedniBroj: 1,
                    Klasa: 'I', Kolicina: 100, Cena: 50, KolAmbalaze: 4
                })
            ];

            assert.deepStrictEqual(m.indeksIzRedova(m.kolone, redovi), {},
                'red desktop push-a je usao u indeks stavki sa terena');
        },

        'indeks nosi roditelja i sadrzaj po stavci': async function (m) {
            const redovi = [
                wireRed(m.kolone, {
                    RedniBroj: 1, Klasa: 'I', Kolicina: 100, Cena: 50, KolAmbalaze: 4,
                    ClientRecordID: 'ITEM-1', OtkupClientRecordID: O1
                })
            ];

            const idx = m.indeksIzRedova(m.kolone, redovi);
            assert.strictEqual(idx['ITEM-1'].parent, O1, 'indeks ne nosi roditelja stavke');
            assert.strictEqual(idx['ITEM-1'].kljuc, m.kljuc('I', 100, 50, 4, 0),
                'indeks ne nosi sadrzaj stavke');
        },

        'indeks pada na red sa roditeljem a bez svog ClientRecordID': async function (m) {
            const redovi = [
                wireRed(m.kolone, {
                    RedniBroj: 1, Klasa: 'I', Kolicina: 100, Cena: 50,
                    OtkupClientRecordID: O1
                })
            ];

            assert.throws(() => m.indeksIzRedova(m.kolone, redovi),
                /nema svoj ClientRecordID/,
                'red bez identiteta je prosao -- GAS bi ga pustio a master odbio ceo list');
        },

        'indeks pada na dva reda sa istim ClientRecordID stavke': async function (m) {
            const red = {
                RedniBroj: 1, Klasa: 'I', Kolicina: 100, Cena: 50,
                ClientRecordID: 'ITEM-1', OtkupClientRecordID: O1
            };
            const redovi = [wireRed(m.kolone, red), wireRed(m.kolone, red)];

            assert.throws(() => m.indeksIzRedova(m.kolone, redovi),
                /istim ClientRecordID/,
                'dupli item CRID je tiho prepisan -- master ga odbija fail-closed');
        },

        // ============================================================
        // Review #394, treci krug: IDENTITET OBUHVATA CEO PAYLOAD
        // ============================================================
        //
        // Stavke su strogo uskladjene, ali cinjenice ZAGLAVLJA nisu bile. Isti CRID
        // sa drugim kooperantom vracao je existing/success, pa promenjen retry
        // nikad ne stigne do mastera: zaglavlje se ne prepise, PwaIstiSadrzaj nema
        // sta da uporedi, i ispravka tiho nestaje.
        'zaglavlje: iste cinjenice su idempotentne': async function (m) {
            const p = zaglavljeRed(m.zaglavljeKolone, ZAGL);

            assert.strictEqual(m.zaglavljeRazlika(p.red, p.idx, ZAGL), '',
                'ponovljen sync istog zaglavlja je prijavljen kao konflikt');
        },

        'zaglavlje: drugi kooperant pod istim CRID-om JE konflikt': async function (m) {
            const p = zaglavljeRed(m.zaglavljeKolone, ZAGL);
            const drugi = Object.assign({}, ZAGL, { KooperantID: 'K2' });

            const r = m.zaglavljeRazlika(p.red, p.idx, drugi);
            assert.notStrictEqual(r, '',
                'promenjen kooperant je prosao kao duplikat -- ispravka bi tiho nestala');
            assert.match(String(r), /KooperantID/, 'konflikt ne imenuje polje');
        },

        'zaglavlje: drugi manifest JE konflikt': async function (m) {
            const p = zaglavljeRed(m.zaglavljeKolone, ZAGL);
            const drugi = Object.assign({}, ZAGL, { StavkeCount: 2 });

            assert.match(String(m.zaglavljeRazlika(p.red, p.idx, drugi)), /StavkeCount/,
                'promenjen manifest nije prijavljen kao razlika');
        },

        // TRANSPORTNA METADATA NIJE TVRDNJA O DOKUMENTU. Da ucestvuje, svaki retry
        // bi bio konflikt -- vremena i DeviceID se i OCEKUJE da se razlikuju.
        'zaglavlje: transportna metadata NE ucestvuje': async function (m) {
            const uListu = Object.assign({}, ZAGL, {
                CreatedAtClient: '2026-09-26T08:00:00.000Z',
                UpdatedAtClient: '2026-09-26T08:00:00.000Z',
                UpdatedAtServer: '2026-09-26T08:00:01.000Z',
                ReceivedAt: '2026-09-26T08:00:01.000Z',
                DeviceID: 'DEV-1',
                KooperantName: 'Kooperant Jedan',
                VozacID: 'VOZ-1'
            });
            const p = zaglavljeRed(m.zaglavljeKolone, uListu);

            // Ulaz nosi DRUGA vremena, drugi uredjaj, drugu labelu i drugog vozaca.
            const stiglo = Object.assign({}, ZAGL, {
                CreatedAtClient: '2026-09-26T09:30:00.000Z',
                UpdatedAtClient: '2026-09-26T09:30:00.000Z',
                UpdatedAtServer: '2026-09-26T09:30:01.000Z',
                ReceivedAt: '2026-09-26T09:30:01.000Z',
                DeviceID: 'DEV-2',
                KooperantName: 'K. Jedan',
                VozacID: 'VOZ-2'
            });

            assert.strictEqual(m.zaglavljeRazlika(p.red, p.idx, stiglo), '',
                'transportna metadata je usla u ugovor -- svaki retry bi bio konflikt');
        },

        // BROJ DOKUMENTA JE USLOVAN, doslovno kao u PwaIstiSadrzaj: prazan incoming
        // broj znaci da ga master generise lokalno, pa bi poredjenje prijavljivalo
        // konflikt tamo gde ga nema.
        'zaglavlje: BrojDokumenta ucestvuje samo kad ga PWA posalje': async function (m) {
            const p = zaglavljeRed(m.zaglavljeKolone,
                                   Object.assign({}, ZAGL, { BrojDokumenta: 'A' }));

            assert.strictEqual(
                m.zaglavljeRazlika(p.red, p.idx, Object.assign({}, ZAGL, { BrojDokumenta: '' })),
                '',
                'prazan incoming broj je prijavljen kao razlika -- lokalno generisan broj bi svaki retry rusio');

            assert.match(
                String(m.zaglavljeRazlika(p.red, p.idx,
                                         Object.assign({}, ZAGL, { BrojDokumenta: 'B' }))),
                /BrojDokumenta/,
                'poslat drugi broj dokumenta nije prijavljen kao razlika');
        },

        // Polje koje se poredi a nije u rasporedu kolona dalo bi undefined indeks:
        // svaka vrednost bi ispala prazna i SVAKI retry bi bio konflikt.
        'zaglavlje: sva poredjena polja postoje u ugovoru kolona': async function (m) {
            const kolone = m.zaglavljeKolone;

            m.zaglavljePolja.concat(['BrojDokumenta']).forEach(function (ime) {
                assert.ok(kolone.indexOf(ime) >= 0,
                    'polje ' + ime + ' se poredi a nije u rasporedu kolona zaglavlja');
            });
        },

        // Treca grana merge kljuca je obrisana: gradila je kljuc od atributa,
        // medju njima Kolicina i Cena, koje zaglavlje vise ne nosi. Takav kljuc
        // bi za SVE redove bez ID-jeva bio isti tekst od praznina.
        'red bez identiteta ne dobija sastavljen merge kljuc': async function (m) {
            const bezId = {
                OtkupacID: 'ST-1', Datum: '2026-09-26', KooperantID: 'KOOP-1',
                VrstaVoca: 'Malina', SortaVoca: 'Willamette', ParcelaID: 'P-1'
            };

            assert.strictEqual(m.mergeKljuc(bezId), '',
                'red bez identiteta je dobio kljuc sastavljen od atributa');

            assert.strictEqual(m.mergeKljuc({ ServerRecordID: 'OTK-1' }), 'SID:OTK-1',
                'red sa ServerRecordID nije adresiran po njemu');
            assert.strictEqual(m.mergeKljuc({ ClientRecordID: 'CRID-1' }), 'CRID:CRID-1',
                'red sa ClientRecordID nije adresiran po njemu');
        }
    }
};
