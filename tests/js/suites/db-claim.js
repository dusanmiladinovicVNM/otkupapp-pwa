'use strict';

// ============================================================
// dbClaimInStore -- atomski claim (#392, treci krug)
// ============================================================
//
// Ovo je jedina tvrdnja u celom rezu koja se NE MOZE procitati iz koda: da li
// provera i upis stvarno padaju u jednu transakciju zavisi od ponasanja
// IndexedDB-a pod dva istovremena pozivaoca, ne od toga kako je funkcija
// napisana. Zato harness i postoji.

const assert = require('node:assert');
const { ucitaj } = require('../harness');

const FAJL = 'src/js/services/db.js';
const STORE = 'zbirne';

// STABILAN ULAZ: /auto instalira indexedDB i IDBKeyRange na globalThis i postoji
// u svakoj verziji paketa. Oblik named exporta se menjao izmedju verzija, pa se
// na njega ne oslanjamo.
require('fake-indexeddb/auto');

const IDB = globalThis.indexedDB;
const KEY_RANGE = globalThis.IDBKeyRange;

if (!IDB) {
    throw new Error('fake-indexeddb/auto nije instalirao globalThis.indexedDB');
}

// IZOLACIJA IDE PREKO IMENA BAZE, NE PREKO require kesa.
// fake-indexeddb drzi stanje u modulu; brisanje kesa je krhko (zavisi od
// exports mape paketa), a sveza baza po tvrdnji je i jasnija i sigurnija.
let brojac = 0;

// Otvaranje je razdvojeno od pravljenja NAMERNO (review #393, P1): tvrdnja o
// dva taba mora da otvori istu bazu DVAPUT, a ne da dva pozivaoca deli jedan
// IDBDatabase objekat.
function otvori(ime) {
    return new Promise(function (resolve, reject) {
        const req = IDB.open(ime, 1);
        req.onupgradeneeded = function () {
            req.result.createObjectStore(STORE, { keyPath: 'clientRecordID' });
        };
        req.onsuccess = function () { resolve(req.result); };
        req.onerror = function () { reject(req.error); };
    });
}

async function novaBaza() {
    brojac += 1;
    const ime = 'harness-zbirne-' + brojac;
    return { ime: ime, db: await otvori(ime) };
}

function svi(db) {
    return new Promise(function (resolve, reject) {
        const tx = db.transaction(STORE, 'readonly');
        const req = tx.objectStore(STORE).getAll();
        req.onsuccess = function () { resolve(req.result || []); };
        req.onerror = function () { reject(req.error); };
    });
}

function zbirna(crid, otpremnicaIDs) {
    return { clientRecordID: crid, otpremnicaIDs: otpremnicaIDs, syncStatus: 'pending' };
}

// Zauzeto = bilo koji zapis u store-u koji nosi ovaj OtpremnicaID.
function zauzima(zapisi, otpID) {
    return (zapisi || []).some(function (z) {
        return String((z && z.otpremnicaIDs) || '').split(',')
            .map(function (x) { return x.trim(); })
            .indexOf(otpID) >= 0;
    });
}

// Baza koja NE MOZE da procita: getAll odmah javlja gresku.
// "Ne znam stanje" i "stanje dozvoljava" nisu isto, i to je ceo predmet tvrdnje.
function bazaKojaNeCita() {
    return {
        objectStoreNames: { contains: function () { return true; } },
        transaction: function () {
            const tx = { abort: function () {} };
            const req = {};
            setTimeout(function () {
                if (typeof req.onerror === 'function') {
                    req.onerror({ target: { error: new Error('citanje je palo') } });
                }
            }, 0);
            tx.objectStore = function () {
                return { getAll: function () { return req; }, put: function () {} };
            };
            return tx;
        }
    };
}

async function ucitajModul(opcije) {
    const ctx = ucitaj([FAJL], Object.assign({
        globali: { indexedDB: IDB, IDBKeyRange: KEY_RANGE }
    }, opcije || {}));

    return { claim: ctx.window.dbClaimInStore };
}

module.exports = {
    ime: 'db-claim',
    ucitaj: ucitajModul,
    tvrdnje: {

        'claim upisuje kad provera ne vrati razlog': async function (m) {
            const db = (await novaBaza()).db;

            const r = await m.claim(db, STORE, zbirna('ZBR-A', 'OTP-1'), function () { return ''; });

            assert.strictEqual(r.ok, true, 'claim je odbijen bez razloga');
            const zapisi = await svi(db);
            assert.strictEqual(zapisi.length, 1, 'zapis nije upisan');
            assert.strictEqual(zapisi[0].clientRecordID, 'ZBR-A');
        },

        'claim odbija i NE upisuje kad provera vrati razlog': async function (m) {
            const db = (await novaBaza()).db;
            await m.claim(db, STORE, zbirna('ZBR-A', 'OTP-1'), function () { return ''; });

            const r = await m.claim(db, STORE, zbirna('ZBR-B', 'OTP-1'), function (zapisi) {
                return zauzima(zapisi, 'OTP-1') ? 'OTP-1 je vec u drugoj zbirnoj' : '';
            });

            assert.strictEqual(r.ok, false, 'claim je prosao preko zauzete otpremnice');
            assert.match(String(r.razlog), /OTP-1/, 'razlog ne imenuje otpremnicu');

            const zapisi = await svi(db);
            assert.strictEqual(zapisi.length, 1, 'odbijen claim je ipak upisao zapis');
            assert.strictEqual(zapisi[0].clientRecordID, 'ZBR-A');
        },

        'claim BACA kad citanje padne, ne vraca ok:false': async function (m) {
            let bacio = false;

            try {
                await m.claim(bazaKojaNeCita(), STORE, zbirna('ZBR-A', 'OTP-1'),
                              function () { return ''; });
            } catch (err) {
                bacio = true;
            }

            assert.strictEqual(bacio, true,
                'pad citanja je vracen kao odgovor umesto da bude bacen');
        },

        // DVE KONEKCIJE, NE DVA POZIVAOCA (review #393, P1).
        //
        // Prva verzija ove tvrdnje je oba claim-a pustila kroz ISTI IDBDatabase
        // objekat. To dokazuje slabiju stvar -- da se dve transakcije na jednoj
        // konekciji serijalizuju -- a #392 je uveo dbClaimInStore zbog DVA TABA,
        // a tab ima svoju konekciju. Zato se ista baza otvara dvaput.
        'dve konekcije nad istom bazom: tacno jedan claim prolazi': async function (m) {
            const prva = await novaBaza();
            const druga = await otvori(prva.ime);

            // Ako bi implementacija vracala kesiranu konekciju, tvrdnja bi tiho
            // skliznula nazad na slabiju verziju. Zato se to meri, ne pretpostavlja.
            assert.notStrictEqual(prva.db, druga,
                'druga konekcija je isti objekat -- tvrdnja bi merila jednu konekciju');

            function provera(zapisi) {
                return zauzima(zapisi, 'OTP-1') ? 'OTP-1 je vec u drugoj zbirnoj' : '';
            }

            // BEZ await izmedju: obe konekcije polaze iz istog stanja baze, kao
            // dva taba nad istim uredjajem. withSubmitLock ovo ne hvata -- brava
            // zivi u memoriji jednog taba.
            const [a, b] = await Promise.all([
                m.claim(prva.db, STORE, zbirna('ZBR-A', 'OTP-1'), provera),
                m.claim(druga, STORE, zbirna('ZBR-B', 'OTP-1'), provera)
            ]);

            const prosli = [a, b].filter(function (r) { return r && r.ok === true; });
            assert.strictEqual(prosli.length, 1,
                'ista otpremnica je usla u ' + prosli.length + ' zbirnih (ocekivano 1)');

            const zapisi = await svi(prva.db);
            assert.strictEqual(zapisi.length, 1, 'u bazi su dve zbirne nad istom otpremnicom');
        }
    }
};
