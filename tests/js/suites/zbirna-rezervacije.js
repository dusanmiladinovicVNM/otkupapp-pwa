'use strict';

// ============================================================
// rezervisaneOtpremnice -- lifecycle privremene istine (#392)
// ============================================================
//
// Dve istine odlucuju da li je otpremnica slobodna:
//
//   kanonska tekuca -- Otpremnica.zbirnaID, koju server zna tek POSLE master
//                      ciklusa
//   privremena      -- lokalna zbirna koju master jos nije razresio
//
// Ceo rizik je u tome KAD privremena prestaje da vazi. Prerano -- ista
// otpremnica ulazi u dve zbirne. Prekasno -- posle odbijanja ili storna se nikad
// ne vraca u izbor.

const assert = require('node:assert');
const { ucitaj } = require('../harness');

const FAJL = 'src/js/features/vozac/zbirna.js';

function zbirna(crid, otpID, dodatno) {
    return Object.assign({
        clientRecordID: crid,
        otpremnicaIDs: otpID,
        syncStatus: 'synced',
        masterStatus: '',
        lastServerCode: '',
        deleted: false
    }, dodatno || {});
}

// Jedan fixture za sve cetiri tvrdnje: svaka zbirna je u DRUGOJ fazi zivota, pa
// jedna sabotaza mora da obori tacno jednu tvrdnju.
function fixture() {
    return [
        // jos na putu, master je nije video -> DRZI
        zbirna('ZBR-1', 'OTP-1', { syncStatus: 'synced' }),
        // master presudio -> od tog trenutka govori kanon, ne lokalni dogadjaj
        zbirna('ZBR-2', 'OTP-2', { masterStatus: 'Synced>Master' }),
        // server presudio i odbio -> retry ne menja ishod, mora da OSLOBODI
        zbirna('ZBR-3', 'OTP-3', { syncStatus: 'pending', lastServerCode: 'ZBIRNA_CONFLICT' }),
        // mreza pukla, koda nema -> zapis je i dalje na putu, DRZI
        zbirna('ZBR-4', 'OTP-4', { syncStatus: 'pending', lastServerCode: '' })
    ];
}

async function ucitajModul(opcije) {
    const ctx = ucitaj([FAJL], opcije || {});

    assert.strictEqual(typeof ctx.rezervisaneOtpremnice, 'function',
        'rezervisaneOtpremnice nije vidljiva posle ucitavanja ' + FAJL);

    return {
        rezervisane: ctx.rezervisaneOtpremnice,
        razresena: ctx.zbirnaRazresenaOdMastera
    };
}

module.exports = {
    ime: 'zbirna-rezervacije',
    ucitaj: ucitajModul,
    tvrdnje: {

        'nerazresena lokalna zbirna DRZI svoju otpremnicu': async function (m) {
            const rez = m.rezervisane(fixture());
            assert.strictEqual(rez.has('OTP-1'), true,
                'OTP-1 je slobodna iako je zbirna koja je nosi jos nerazresena');
        },

        'zbirna koju je master razresio ne drzi nista': async function (m) {
            const rez = m.rezervisane(fixture());
            assert.strictEqual(rez.has('OTP-2'), false,
                'OTP-2 ostaje rezervisana i posle masterove odluke -- posle storna se nikad ne bi vratila u izbor');
        },

        'trajno odbijena zbirna OSLOBADJA otpremnicu': async function (m) {
            const rez = m.rezervisane(fixture());
            assert.strictEqual(rez.has('OTP-3'), false,
                'OTP-3 ostaje zakljucana odbijenom zbirnom -- svaki retry pada iz istog razloga');
        },

        'transportni pad ZADRZAVA rezervaciju': async function (m) {
            const rez = m.rezervisane(fixture());
            assert.strictEqual(rez.has('OTP-4'), true,
                'OTP-4 je oslobodjena zbog pada mreze -- zapis je jos na putu i moze da uspe');
        },

        'terminalni statusi mastera su prepoznati doslovno': async function (m) {
            assert.strictEqual(m.razresena({ masterStatus: 'Synced>Master' }), true);
            assert.strictEqual(m.razresena({ masterStatus: 'Duplicate' }), true);
            assert.strictEqual(m.razresena({ masterStatus: 'SyncError: nesto' }), true);
            assert.strictEqual(m.razresena({ masterStatus: '' }), false);
            assert.strictEqual(m.razresena({ masterStatus: 'pending' }), false);
        }
    }
};
