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
const { ucitaj } = require('../harness');

const FAJL = 'gas/Code.gs';

async function ucitajModul(opcije) {
    const ctx = ucitaj([FAJL], opcije || {});

    for (const ime of ['otkStavkeNormalizuj_', 'otkStavkeUskladi_',
                       'otkStavkaKljuc_', 'buildOtkupMergeKey_']) {
        assert.strictEqual(typeof ctx[ime], 'function',
            ime + ' nije vidljiva posle ucitavanja ' + FAJL);
    }

    return {
        normalizuj: ctx.otkStavkeNormalizuj_,
        uskladi: ctx.otkStavkeUskladi_,
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

            assert.strictEqual(m.uskladi(tab, stigle, O1), '',
                'drugi redosled je prijavljen kao konflikt -- retry bi postao greska');
        },

        'izmenjena kolicina JE razlika i imenuje stavku': async function (m) {
            const a = stavka({ clientRecordID: 'CRID-S1' });
            const tab = indeks(m.kljuc, [['CRID-S1', O1, a]]);
            const stigle = m.normalizuj(
                [stavka({ clientRecordID: 'CRID-S1', kolicina: 999 })], O1);

            const r = m.uskladi(tab, stigle, O1);
            assert.notStrictEqual(r, '', 'izmenjena kolicina nije prijavljena kao razlika');
            assert.match(String(r), /CRID-S1/, 'razlika ne imenuje koja stavka');
        },

        'stavka u tabu koju ulaz ne nosi JE razlika': async function (m) {
            const a = stavka({ clientRecordID: 'CRID-S1', klasa: 'I' });
            const b = stavka({ clientRecordID: 'CRID-S2', klasa: 'II', kolicina: 60 });

            const tab = indeks(m.kljuc, [['CRID-S1', O1, a], ['CRID-S2', O1, b]]);
            const stigle = m.normalizuj([a], O1);

            assert.match(String(m.uskladi(tab, stigle, O1)), /CRID-S2/,
                'zaboravljena stavka nije prijavljena kao razlika');
        },

        // PRAZAN TAB JE PRVI UPIS, ne razlika. Bez ovoga nijedan nov otkup ne bi
        // mogao da prodje -- svaki bi odmah bio OTKUP_CONFLICT.
        'prazan tab NIJE razlika': async function (m) {
            const stigle = m.normalizuj([stavka()], O1);
            assert.strictEqual(m.uskladi({}, stigle, O1), '',
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

            const r = m.uskladi(tab, retry, O1);
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

            const r = m.uskladi(tab, stigle, O2);
            assert.notStrictEqual(r, '',
                'isti item CRID pod drugim otkupom je prosao -- VBA bi ga odbio i oborio ceo list');
            assert.match(String(r), /ITEM-X/, 'konflikt ne imenuje stavku');
            assert.match(String(r), new RegExp(O1), 'konflikt ne imenuje tudjeg roditelja');
        },

        // ASIMETRIJA JE NAMERNA: dopisati sto fali je dovrsavanje istog upisa,
        // zaboraviti sto postoji je druga tvrdnja o dokumentu.
        //
        // Ova tvrdnja NEMA svoju sabotazu, i to je odluka: pravilo je odsustvo
        // provere, pa bi mu seam trebalo DODATI kod -- a odbrana napisana pre
        // merenja je vec jednom postala nalaz (#393, treci krug). Sabotaza
        // 'otk-stavke-recovery-je-konflikt' meri isti izlaz sa druge strane.
        'stavka koja u tabu fali je RECOVERY, ne konflikt': async function (m) {
            const a = stavka({ clientRecordID: 'CRID-S1', klasa: 'I' });
            const tab = indeks(m.kljuc, [['CRID-S1', O1, a]]);

            const stigle = m.normalizuj([
                stavka({ clientRecordID: 'CRID-S1', klasa: 'I' }),
                stavka({ clientRecordID: 'CRID-S2', klasa: 'II', kolicina: 60 })
            ], O1);

            assert.strictEqual(m.uskladi(tab, stigle, O1), '',
                'dopuna nedostajuce stavke je prijavljena kao konflikt -- prekinut upis se ne bi mogao dovrsiti');
        },

        // Stavke TUDJEG otkupa u tabu ne smeju da smetaju: tab je zajednicki, pa
        // bi inace svaki drugi dokument na listu pravio laznu razliku.
        'stavke drugog otkupa ne ulaze u poredjenje': async function (m) {
            const moja = stavka({ clientRecordID: 'CRID-S1', klasa: 'I' });
            const tudja = stavka({ clientRecordID: 'CRID-T1', klasa: 'I' });

            const tab = indeks(m.kljuc, [['CRID-S1', O1, moja], ['CRID-T1', O2, tudja]]);
            const stigle = m.normalizuj([stavka({ clientRecordID: 'CRID-S1', klasa: 'I' })], O1);

            assert.strictEqual(m.uskladi(tab, stigle, O1), '',
                'stavka drugog otkupa je prijavljena kao razlika');
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
