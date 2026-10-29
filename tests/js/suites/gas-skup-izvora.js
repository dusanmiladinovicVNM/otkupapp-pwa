'use strict';

// ============================================================
// skupIzvoraRazlika_ -- redosled nije tvrdnja (#392, prvi krug)
// ============================================================
//
// GAS mora da razlikuje "isti manifest, drugi redosled" od "druga tvrdnja o
// dokumentu". Prvo je retry, drugo je ZBIRNA_CONFLICT.

const assert = require('node:assert');
const { ucitaj } = require('../harness');

const FAJL = 'gas/Code.gs';

async function ucitajModul(opcije) {
    const ctx = ucitaj([FAJL], opcije || {});

    assert.strictEqual(typeof ctx.skupIzvoraRazlika_, 'function',
        'skupIzvoraRazlika_ nije vidljiva posle ucitavanja ' + FAJL);

    return { razlika: ctx.skupIzvoraRazlika_ };
}

module.exports = {
    ime: 'gas-skup-izvora',
    ucitaj: ucitajModul,
    tvrdnje: {

        'isti skup u drugom redosledu NIJE razlika': async function (m) {
            assert.strictEqual(m.razlika('OTP-1,OTP-2', 'OTP-2,OTP-1'), '',
                'drugi redosled je prijavljen kao konflikt -- retry bi postao greska');
        },

        'dodata otpremnica JE razlika i imenuje polje': async function (m) {
            const r = m.razlika('OTP-1', 'OTP-1,OTP-2');
            assert.notStrictEqual(r, '', 'dodata otpremnica nije prijavljena kao razlika');
            assert.match(String(r), /OtpremnicaIDs/, 'razlika ne imenuje polje');
        },

        'uklonjena otpremnica JE razlika': async function (m) {
            assert.notStrictEqual(m.razlika('OTP-1,OTP-2', 'OTP-1'), '',
                'uklonjena otpremnica nije prijavljena kao razlika');
        },

        // ISTI REDOSLED NAMERNO: ova tvrdnja meri trimovanje i prazan token, ne
        // redosled. Da meri oba, jedna sabotaza bi obarala dve tvrdnje i nijedna
        // ne bi imenovala svoj razlog.
        'prazni tokeni i razmaci se ne racunaju': async function (m) {
            assert.strictEqual(m.razlika('OTP-1, ,OTP-2', ' OTP-1,OTP-2 '), '',
                'razmak ili prazan token je prijavljen kao druga tvrdnja');
        }
    }
};
