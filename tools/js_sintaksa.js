'use strict';

// ============================================================
// SINTAKSNA KAPIJA ZA JS/GS -- jeftina, determinsticka, bez Excela
// ============================================================
//
//   node tools/js_sintaksa.js              provera svih fajlova
//   node tools/js_sintaksa.js --self-test  provera da kapija zaista grize
//
// Zasto postoji: do ovog reza je PWA/GAS sloj bio jedini deo repoa bez IJEDNE
// automatske provere. Neuravnotezena vitica u src/js se videla tek kao ekran
// koji se ne ucitava, a u gas/Code.gs kao deploy koji pukne. Sopstveni brojac
// vitica koji sam pisao u toku #392 je pomocno sredstvo, ne kapija.
//
// `node --check` se NE koristi: on tip modula bira po ekstenziji, pa .gs ne ume
// da proveri. vm.Script parsira sadrzaj bez obzira na ime fajla.

const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const KOREN = path.resolve(__dirname, '..');

const STABLA = [
    { dir: 'src/js', ext: ['.js'] },
    { dir: 'gas', ext: ['.gs', '.js'] },
    { dir: 'tests/js', ext: ['.js'] },
    { dir: 'tools', ext: ['.js'], plitko: true }
];

function nabroj(rel, ext, plitko) {
    const abs = path.join(KOREN, rel);
    if (!fs.existsSync(abs)) return [];

    const out = [];
    fs.readdirSync(abs, { withFileTypes: true }).forEach(function (e) {
        const pod = path.posix.join(rel, e.name);
        if (e.isDirectory()) {
            if (!plitko) out.push.apply(out, nabroj(pod, ext, false));
        } else if (ext.indexOf(path.extname(e.name)) >= 0) {
            out.push(pod);
        }
    });
    return out.sort();
}

function proveri(rel) {
    const src = fs.readFileSync(path.join(KOREN, rel), 'utf8');
    try {
        new vm.Script(src, { filename: rel });
        return '';
    } catch (err) {
        return (err && err.message) || String(err);
    }
}

function kapija() {
    let fajlova = [];
    STABLA.forEach(function (s) {
        fajlova = fajlova.concat(nabroj(s.dir, s.ext, !!s.plitko));
    });

    let palo = 0;
    fajlova.forEach(function (rel) {
        const greska = proveri(rel);
        if (greska) {
            palo += 1;
            console.log('::error file=' + rel + '::' + greska);
        }
    });

    console.log('js_sintaksa: ' + (fajlova.length - palo) + '/' + palo +
                ' (proslo/palo) od ' + fajlova.length + ' fajlova');
    return palo === 0;
}

// --- self-test ---------------------------------------------------------------
// Prva polovina: ulazi koji MORAJU da zapiste. Druga: legalan JS koji NE SME.
const MORA_DA_PADNE = [
    ['neuravnotezena vitica', 'function a() { if (1) { return 2; }'],
    ['visak zatvorene vitice', 'function a() { return 1; } }'],
    ['nedovrsen string', "const a = 'abc;"],
    ['rezervisana rec kao ime', 'const function = 1;'],
    ['zarez u pozivu bez argumenta', 'f(a, , b);'],
    ['top-level await u skripti', 'const x = await f();']
];

const NE_SME_DA_PADNE = [
    ['async/await u funkciji', 'async function a() { return await f(); }'],
    ['template literal', 'const a = `x ${y} z`;'],
    ['regex sa kosom crtom', 'const a = /a\/b/g;'],
    ['opciono lancanje', 'const a = b?.c?.d ?? e;'],
    ['getter u objektu', 'const o = { get a() { return 1; } };'],
    ['klasa sa privatnim poljem', 'class A { #x = 1; y() { return this.#x; } }']
];

function samoTest() {
    let greske = 0;

    function parsira(kod) {
        try {
            new vm.Script(kod, { filename: 'self-test' });
            return true;
        } catch (err) {
            return false;
        }
    }

    MORA_DA_PADNE.forEach(function (par) {
        if (parsira(par[1])) {
            console.log('::error::self-test: kapija je PROPUSTILA -- ' + par[0]);
            greske += 1;
        } else {
            console.log('OK: pada kako treba -- ' + par[0]);
        }
    });

    NE_SME_DA_PADNE.forEach(function (par) {
        if (!parsira(par[1])) {
            console.log('::error::self-test: kapija je LAZNO prijavila -- ' + par[0]);
            greske += 1;
        } else {
            console.log('OK: prolazi kako treba -- ' + par[0]);
        }
    });

    console.log('js_sintaksa self-test: ' +
                (MORA_DA_PADNE.length + NE_SME_DA_PADNE.length - greske) + '/' + greske);
    return greske === 0;
}

const ok = process.argv.indexOf('--self-test') >= 0 ? samoTest() : kapija();
process.exit(ok ? 0 : 1);
