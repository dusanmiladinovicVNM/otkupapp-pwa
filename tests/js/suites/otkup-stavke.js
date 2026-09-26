'use strict';

// ============================================================
// Stavke otkupa u PWA -- zbir je na JEDNOM mestu (S5-5b)
// ============================================================
//
// Zapis otkupa ne nosi klasu, kolicinu, cenu i ambalazu na sebi: nosi stavke[].
// Svaki ekran koji bi sam sabirao je jedna sansa da se dva prikaza iste robe
// raziðu -- i to se vidi tek na iznosu koji neko isplacuje.
//
// UI u ovom rezu unosi JEDNU klasu, pa tvrdnje namerno mere i N=1 i N=2: prva
// dokazuje da se prikaz nije promenio, druga da model vec radi.

const assert = require('node:assert');
const { ucitaj } = require('../harness');

const FAJL = 'src/js/features/otkup/otkup-stavke.js';

async function ucitajModul(opcije) {
    const ctx = ucitaj([FAJL], opcije || {});

    for (const ime of ['otkupStavke', 'otkupZbirKg', 'otkupZbirVrednosti',
                       'otkupZbirAmbalaze', 'otkupKlaseTekst', 'otkupCenaAkoJedna',
                       'otkupJednaStavka', 'novaStavkaOtkupa']) {
        assert.strictEqual(typeof ctx[ime], 'function',
            ime + ' nije vidljiva posle ucitavanja ' + FAJL);
    }

    return ctx;
}

function zapis(stavke) {
    return { clientRecordID: 'CRID-Z1', tipAmbalaze: 'Gajba', stavke: stavke };
}

module.exports = {
    ime: 'otkup-stavke',
    ucitaj: ucitajModul,
    tvrdnje: {

        'jedna klasa: zbirovi su isti kao stara ravna polja': async function (m) {
            const r = zapis([{ klasa: 'I', kolicina: 400, cena: 50, kolAmbalaze: 20 }]);

            assert.strictEqual(m.otkupZbirKg(r), 400, 'kg jedne stavke nije zbir');
            assert.strictEqual(m.otkupZbirVrednosti(r), 20000,
                'vrednost jedne stavke nije kolicina * cena');
            assert.strictEqual(m.otkupZbirAmbalaze(r), 20, 'gajbe jedne stavke nisu zbir');
            assert.strictEqual(m.otkupKlaseTekst(r), 'I', 'oznaka klase se promenila za N=1');
            assert.strictEqual(m.otkupCenaAkoJedna(r), 50, 'cena jedne klase nije vidljiva');
        },

        'dve klase: vrednost je zbir PO STAVCI, ne proizvod zbirova': async function (m) {
            // 400*50 + 60*30 = 21800. Proizvod zbirova bi dao 460*80 = 36800 --
            // razlika koja bi izasla na iznos za isplatu.
            const r = zapis([
                { klasa: 'I', kolicina: 400, cena: 50, kolAmbalaze: 20 },
                { klasa: 'II', kolicina: 60, cena: 30, kolAmbalaze: 4 }
            ]);

            assert.strictEqual(m.otkupZbirKg(r), 460, 'kg dve stavke nisu sabrane');
            assert.strictEqual(m.otkupZbirVrednosti(r), 21800,
                'vrednost nije zbir po stavci');
            assert.strictEqual(m.otkupZbirAmbalaze(r), 24, 'gajbe dve stavke nisu sabrane');
        },

        'dve klase NEMAJU jednu cenu': async function (m) {
            const r = zapis([
                { klasa: 'I', kolicina: 400, cena: 50, kolAmbalaze: 0 },
                { klasa: 'II', kolicina: 60, cena: 30, kolAmbalaze: 0 }
            ]);

            assert.strictEqual(m.otkupCenaAkoJedna(r), null,
                'dvoklasni dokument je dao jednu cenu -- prikaz bi tvrdio pogresan podatak');
            assert.strictEqual(m.otkupJednaStavka(r), null,
                'dvoklasni dokument je dao jednu stavku');
        },

        // KANONSKI RED KLASA, ne red stizanja: isto pravilo po kom master
        // dodeljuje RedniBroj, pa se oznaka ne menja od redosleda sinhronizacije.
        'oznaka klasa ide kanonskim redom, ne redom stizanja': async function (m) {
            const r = zapis([
                { klasa: 'II', kolicina: 60, cena: 30, kolAmbalaze: 0 },
                { klasa: 'I', kolicina: 400, cena: 50, kolAmbalaze: 0 }
            ]);

            assert.strictEqual(m.otkupKlaseTekst(r), 'I, II',
                'oznaka klasa prati red stizanja umesto kanonskog reda');
        },

        // ZAPIS BEZ STAVKI NE DOBIJA IZMISLJEN PODATAK. Pristupnik vraca nulu i
        // prazno; odluku sta to znaci donosi pozivalac, koji jedini zna kontekst.
        'zapis bez stavki daje nulu, ne izuzetak i ne pogodjen broj': async function (m) {
            assert.deepStrictEqual(m.otkupStavke({}), [], 'zapis bez stavki nije dao prazan niz');
            assert.deepStrictEqual(m.otkupStavke(null), [], 'null zapis nije dao prazan niz');
            assert.strictEqual(m.otkupZbirKg({}), 0, 'zapis bez stavki nije dao nula kg');
            assert.strictEqual(m.otkupKlaseTekst({}), '', 'zapis bez stavki je dao oznaku klase');
            assert.strictEqual(m.otkupCenaAkoJedna({}), null, 'zapis bez stavki je dao cenu');
        },

        'nova stavka nosi SVOJ identitet': async function (m) {
            // Bez identiteta reda ponovljen sync ne razlikuje retry od druge
            // stavke, pa bi se kolicina udvajala.
            const a = m.novaStavkaOtkupa({ klasa: 'I', kolicina: 400, cena: 50, kolAmbalaze: 20 });
            const b = m.novaStavkaOtkupa({ klasa: 'II', kolicina: 60, cena: 30, kolAmbalaze: 4 });

            assert.ok(a.clientRecordID, 'stavka je nastala bez ClientRecordID-a');
            assert.notStrictEqual(a.clientRecordID, b.clientRecordID,
                'dve stavke su dobile isti identitet');
        },

        // Bruto se cuva SAMO kad je unos bio bruto: nula nije "bruto = 0" nego
        // "nije bruto unos", pa pisac u masteru bruto tada i ne cuva.
        'bruto se nosi samo kad postoji': async function (m) {
            const bez = m.novaStavkaOtkupa({ klasa: 'I', kolicina: 400, cena: 50 });
            const sa = m.novaStavkaOtkupa({ klasa: 'I', kolicina: 400, cena: 50, brutoKg: 430 });

            assert.strictEqual('brutoKg' in bez, false,
                'bruto je upisan i kad ga unos nije imao');
            assert.strictEqual(sa.brutoKg, 430, 'bruto sa unosa nije prenet');
        }
    }
};
