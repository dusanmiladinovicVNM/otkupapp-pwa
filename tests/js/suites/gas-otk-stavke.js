'use strict';

// ============================================================
// OTK_STAVKE u GAS-u -- skup stavki je dokument (S5-5b)
// ============================================================
//
// Zaglavlje OTK lista od ovog reza ne nosi klasu, kolicinu, cenu ni ambalazu.
// Nosi MANIFEST. Zato GAS mora da razlikuje tri stvari koje su do sada bile
// jedna: "nema stavki", "iste stavke u drugom redosledu" i "druga tvrdnja o
// istom dokumentu".
//
// Mere se CISTI pomocnici: harness ne moze da podigne SpreadsheetApp (stubovi
// bacaju izuzetak sa imenom clana), pa I/O ostaje tanak a odluka merena.

const assert = require('node:assert');
const { ucitaj } = require('../harness');

const FAJL = 'gas/Code.gs';

async function ucitajModul(opcije) {
    const ctx = ucitaj([FAJL], opcije || {});

    for (const ime of ['otkStavkeNormalizuj_', 'otkStavkeRazlika_',
                       'otkStavkaKljuc_', 'buildOtkupMergeKey_']) {
        assert.strictEqual(typeof ctx[ime], 'function',
            ime + ' nije vidljiva posle ucitavanja ' + FAJL);
    }

    return {
        normalizuj: ctx.otkStavkeNormalizuj_,
        razlika: ctx.otkStavkeRazlika_,
        kljuc: ctx.otkStavkaKljuc_,
        mergeKljuc: ctx.buildOtkupMergeKey_
    };
}

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

// Mapa "sta je u tabu", u obliku koji otkStavkeRazlika_ prima.
function uTabu(kljuc, parovi) {
    const m = {};
    for (const [id, s] of parovi) {
        m[id] = kljuc(s.klasa, s.kolicina, s.cena, s.kolAmbalaze, s.brutoKg || 0);
    }
    return m;
}

module.exports = {
    ime: 'gas-otk-stavke',
    ucitaj: ucitajModul,
    tvrdnje: {

        'otkup bez stavki se odbija PO IMENU': async function (m) {
            assert.throws(() => m.normalizuj([], 'CRID-Z1'), /Stavke missing/,
                'prazan skup je prosao -- dokument od nula kilograma');
            assert.throws(() => m.normalizuj(undefined, 'CRID-Z1'), /Stavke missing/,
                'nedostajuce stavke su prosle kao prazan skup');
        },

        'stavka bez svog ClientRecordID se odbija': async function (m) {
            // Bez identiteta reda se ponovljen sync ne razlikuje od druge stavke,
            // pa bi retry posle mreznog pada udvajao robu.
            assert.throws(
                () => m.normalizuj([stavka({ clientRecordID: '' })], 'CRID-Z1'),
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
                ], 'CRID-Z1'),
                /Dve stavke iste klase/,
                'dve stavke iste klase su prosle');
        },

        'dve stavke sa istim ClientRecordID se odbijaju': async function (m) {
            assert.throws(
                () => m.normalizuj([
                    stavka({ clientRecordID: 'CRID-S1', klasa: 'I' }),
                    stavka({ clientRecordID: 'CRID-S1', klasa: 'II' })
                ], 'CRID-Z1'),
                /istim ClientRecordID/,
                'dva reda istog identiteta su prosla kao dve stavke');
        },

        'normalizacija nosi redni broj po redu stizanja i bruto kad postoji':
        async function (m) {
            const out = m.normalizuj([
                stavka({ clientRecordID: 'CRID-S1', klasa: 'II', kolicina: 60 }),
                stavka({ clientRecordID: 'CRID-S2', klasa: 'I', brutoKg: 110 })
            ], 'CRID-Z1');

            assert.strictEqual(out.length, 2, 'obe stavke nisu prosle');
            assert.strictEqual(out[0].redniBroj, 1, 'prva stavka nije dobila redni broj 1');
            assert.strictEqual(out[1].redniBroj, 2, 'druga stavka nije dobila redni broj 2');
            assert.strictEqual(out[0].otkupClientRecordID, 'CRID-Z1',
                'stavka ne nosi CRID svog zaglavlja');
            assert.strictEqual(out[1].brutoKg, 110, 'bruto sa zice nije prenet');
        },

        // REDOSLED NIJE TVRDNJA O DOKUMENTU. Master dodeljuje RedniBroj po
        // kanonskom redu klasa, pa je redosled kojim je PWA slala nebitan.
        'iste stavke u drugom redosledu NISU razlika': async function (m) {
            const a = stavka({ clientRecordID: 'CRID-S1', klasa: 'I' });
            const b = stavka({ clientRecordID: 'CRID-S2', klasa: 'II', kolicina: 60 });

            const tab = uTabu(m.kljuc, [['CRID-S1', a], ['CRID-S2', b]]);
            const stigle = m.normalizuj([b, a], 'CRID-Z1');

            assert.strictEqual(m.razlika(tab, stigle), '',
                'drugi redosled je prijavljen kao konflikt -- retry bi postao greska');
        },

        'izmenjena kolicina JE razlika i imenuje stavku': async function (m) {
            const a = stavka({ clientRecordID: 'CRID-S1' });
            const tab = uTabu(m.kljuc, [['CRID-S1', a]]);
            const stigle = m.normalizuj([stavka({ clientRecordID: 'CRID-S1', kolicina: 999 })],
                                        'CRID-Z1');

            const r = m.razlika(tab, stigle);
            assert.notStrictEqual(r, '', 'izmenjena kolicina nije prijavljena kao razlika');
            assert.match(String(r), /CRID-S1/, 'razlika ne imenuje koja stavka');
        },

        'drugi broj stavki JE razlika': async function (m) {
            const a = stavka({ clientRecordID: 'CRID-S1', klasa: 'I' });
            const b = stavka({ clientRecordID: 'CRID-S2', klasa: 'II', kolicina: 60 });

            const tab = uTabu(m.kljuc, [['CRID-S1', a], ['CRID-S2', b]]);
            const stigle = m.normalizuj([a], 'CRID-Z1');

            assert.match(String(m.razlika(tab, stigle)), /broj stavki/,
                'izgubljena stavka nije prijavljena kao razlika');
        },

        // PRAZAN TAB JE PRVI UPIS, ne razlika. Bez ovoga bi svaki nov otkup
        // odmah bio OTKUP_CONFLICT.
        'prazan tab NIJE razlika': async function (m) {
            const stigle = m.normalizuj([stavka()], 'CRID-Z1');
            assert.strictEqual(m.razlika({}, stigle), '',
                'prazan tab je prijavljen kao konflikt -- prvi upis bi pao');
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
