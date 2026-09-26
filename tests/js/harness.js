'use strict';

// ============================================================
// UCITAVANJE PRODUKCIONOG JS-a U IZOLOVAN KONTEKST
// ============================================================
//
// Tvrdnja mora da ide kroz PRODUKCIONI fajl, ne kroz kopiju funkcije prepisanu u
// test. Zato se `src/js/**` i `gas/Code.gs` ucitavaju kakvi su, u vm kontekst sa
// laznim `window`-om: ono sto u pregledacu postane globalno, ovde postane
// svojstvo konteksta i moze da se pozove.
//
// Isti ulaz nosi i SABOTAZU: tekstualnu zamenu po sidru, pre izvrsavanja. Tako
// dvosmerni dokaz ne dira radno stablo -- za razliku od tools/dokaz.py, koji
// kvari fajl na disku (i zato ne sme da se vrti paralelno sa izmenama).

const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const KOREN = path.resolve(__dirname, '..', '..');

// KRAJ REDA SE NORMALIZUJE PRE SIDRA.
//
// Radno stablo na Windows-u je CRLF (core.autocrlf=true), a CI checkout na
// ubuntu-u je LF. Sidro napisano sa \n bi zato gadjalo samo jednu od dve masine,
// a kapija koja vazi na jednoj masini nije kapija -- isti razlog zbog kog
// .gitattributes drzi tools/*.sh na LF.
function normalizuj(t) {
    return String(t).replace(/\r\n/g, '\n');
}

function procitaj(rel) {
    return normalizuj(fs.readFileSync(path.join(KOREN, rel), 'utf8'));
}

// Apps Script globali. Nijedan se ne poziva pri ucitavanju (mereno: gas/Code.gs
// nema izvrsnih iskaza na nultoj koloni), pa stub postoji da GRESKA BUDE JASNA
// ako neka tvrdnja slucajno zaluta u Google API, umesto "undefined is not a
// function" iz sredine fajla.
function appsScriptStub(ime) {
    return new Proxy({}, {
        get: function (_t, prop) {
            throw new Error('Apps Script API nije dostupan u harness-u: ' +
                            ime + '.' + String(prop));
        }
    });
}

function napraviKontekst(dodatno) {
    // Kontekst NAMERNO ne dobija ugradjene tipove host realma (Object, Array, Set...):
    // vm mu daje svoje. Ubacivanje host verzija pravi cross-realm neslaganje,
    // gde `x instanceof Array` laze iako je x niz.
    const window = {};
    const osnova = {
        window: window,
        self: window,
        console: console,
        setTimeout: setTimeout,
        clearTimeout: clearTimeout,
        setInterval: setInterval,
        clearInterval: clearInterval,
        structuredClone: structuredClone,
        navigator: { onLine: true },
        location: { href: 'https://test.local/', origin: 'https://test.local' },
        SpreadsheetApp: appsScriptStub('SpreadsheetApp'),
        PropertiesService: appsScriptStub('PropertiesService'),
        CacheService: appsScriptStub('CacheService'),
        LockService: appsScriptStub('LockService'),
        UrlFetchApp: appsScriptStub('UrlFetchApp'),
        DriveApp: appsScriptStub('DriveApp'),
        Utilities: appsScriptStub('Utilities'),
        Logger: appsScriptStub('Logger'),
        Session: appsScriptStub('Session'),
        ScriptApp: appsScriptStub('ScriptApp'),
        ContentService: appsScriptStub('ContentService'),
        HtmlService: appsScriptStub('HtmlService'),
        MailApp: appsScriptStub('MailApp')
    };

    Object.keys(dodatno || {}).forEach(function (k) {
        osnova[k] = dodatno[k];
    });

    const ctx = vm.createContext(osnova);
    return ctx;
}

function primeni(izvor, rel, sabotaza) {
    let out = izvor;

    (sabotaza || []).forEach(function (s) {
        if (s.fajl !== rel) return;

        const sidro = normalizuj(s.sidro);
        const n = out.split(sidro).length - 1;

        // SIDRO ZASTAREVA. Izmena produkcionog koda pomeri tekst, a sabotaza
        // koja se tiho ne primeni ostavlja ZELEN self-test bez ijednog merenja.
        if (n !== 1) {
            throw new Error('SIDRO_ZASTARELO: ' + rel + ' :: "' +
                            sidro.slice(0, 70).replace(/\n/g, '\\n') +
                            '" -> ' + n + ' pogodaka, ocekivan tacno 1');
        }

        out = out.replace(sidro, normalizuj(s.zamena));
        s._primenjeno = (s._primenjeno || 0) + 1;
    });

    return out;
}

// fajlovi: spisak putanja relativnih na koren repoa, u redosledu ucitavanja.
// opcije.globali: dodatne globalne vrednosti konteksta.
// opcije.sabotaza: [{ fajl, sidro, zamena }]
function ucitaj(fajlovi, opcije) {
    opcije = opcije || {};

    const sabotaza = (opcije.sabotaza || []).map(function (s) {
        return { fajl: s.fajl, sidro: s.sidro, zamena: s.zamena, _primenjeno: 0 };
    });

    // globalThis se NAMERNO ne ubacuje: vm kontekst ga vec ima, a ubacivanje
    // svojstva istog imena zaklanja pravi globalni objekat.
    const ctx = napraviKontekst(opcije.globali);

    fajlovi.forEach(function (rel) {
        const izvor = primeni(procitaj(rel), rel, sabotaza);
        new vm.Script(izvor, { filename: rel }).runInContext(ctx);
    });

    // Katalog koji imenuje fajl koji suite ne ucitava nije sabotaza nego greska
    // u katalogu -- i mora da vristi, ne da prodje kao "nije oborilo test".
    sabotaza.forEach(function (s) {
        if (s._primenjeno !== 1) {
            throw new Error('SABOTAZA NIJE PRIMENJENA: ' + s.fajl +
                            ' nije medju ucitanim fajlovima (' + fajlovi.join(', ') + ')');
        }
    });

    return ctx;
}

module.exports = { ucitaj: ucitaj, procitaj: procitaj, KOREN: KOREN };
