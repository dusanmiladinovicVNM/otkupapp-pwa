'use strict';

// ============================================================
// JS HARNESS -- pokretac
// ============================================================
//
//   node tests/js/pokreni.js              zelena kapija (sve tvrdnje)
//   node tests/js/pokreni.js --self-test  dvosmerni dokaz (svaka sabotaza obara svoju tvrdnju)
//
// Namerno BEZ node:test runnera: isti oblik izlaza i isti --self-test kao svaka
// python kapija u .github/workflows/static.yml. Jedan mehanizam, ne dva.

const SABOTAZE = require('./sabotaze');

const SUITES = [
    require('./suites/db-claim'),
    require('./suites/zbirna-rezervacije'),
    require('./suites/gas-skup-izvora')
];

function poImenu(ime) {
    const s = SUITES.filter(function (x) { return x.ime === ime; })[0];
    if (!s) throw new Error('KATALOG IMENUJE NEPOZNAT SUITE: ' + ime);
    return s;
}

// Vraca spisak imena tvrdnji koje su PUKLE, i mapu poruka.
async function vrtiSuite(suite, opcije) {
    const pale = [];
    const poruke = {};
    const imena = Object.keys(suite.tvrdnje);

    // NEOCEKIVANA ASYNC GRESKA OBARA PROCES, I TAKO TREBA (review #393, treci krug).
    //
    // Ovde je nakratko stajao globalni uncaughtException/unhandledRejection
    // handler koji je zakasnele greske pripisivao tvrdnji "u toku". Dodat je zbog
    // sabotaze koja je gadjala mrtav objectStore; ta sabotaza je zamenjena
    // legalnom, pa je handler ostao samo kao RIZIK: greska koja stigne posle svih
    // tvrdnji bila bi progutana, a greska koja stigne tokom naredne pripisana
    // njoj. Za kapiju je "neocekivano -> crveno" tacnije od pogadjanja cija je
    // greska. Suite kojem stvarno treba pozadinski posao mora da ceka njegov
    // signal zavrsetka, eksplicitno.
    //
    // Modul se ucitava PO TVRDNJI: jedna tvrdnja ne sme da vidi stanje koje je
    // ostavila prethodna (baza, brojaci, modul-level promenljive ekrana).
    for (const ime of imena) {
        try {
            const modul = await suite.ucitaj(opcije);
            await suite.tvrdnje[ime](modul);
        } catch (err) {
            pale.push(ime);
            poruke[ime] = (err && err.message) || String(err);
        }
    }

    return { imena: imena, pale: pale, poruke: poruke };
}

async function zelenaKapija() {
    let ukupno = 0;
    let palo = 0;

    for (const suite of SUITES) {
        const r = await vrtiSuite(suite);
        ukupno += r.imena.length;

        for (const ime of r.imena) {
            if (r.pale.indexOf(ime) >= 0) {
                palo += 1;
                console.log('::error::' + suite.ime + ' :: ' + ime + ' -- ' + r.poruke[ime]);
            } else {
                console.log('OK: ' + suite.ime + ' :: ' + ime);
            }
        }
    }

    console.log('');
    console.log('js harness: ' + (ukupno - palo) + '/' + palo + ' (proslo/palo) od ' + ukupno);
    return palo === 0;
}

async function dvosmerniDokaz() {
    let greske = 0;
    const pokriveno = {};

    for (const s of SABOTAZE) {
        const suite = poImenu(s.suite);

        if (!suite.tvrdnje[s.tvrdnja]) {
            console.log('::error::' + s.ime + ': TVRDNJA NE POSTOJI u suite-u ' +
                        s.suite + ': "' + s.tvrdnja + '"');
            greske += 1;
            continue;
        }

        pokriveno[s.suite + ' :: ' + s.tvrdnja] = true;

        let r;
        try {
            r = await vrtiSuite(suite, {
                sabotaza: [{ fajl: s.fajl, sidro: s.sidro, zamena: s.zamena }]
            });
        } catch (err) {
            // SIDRO_ZASTARELO i SABOTAZA NIJE PRIMENJENA stizu ovde: to nije
            // "nije oborilo test" nego katalog koji vise ne gadja kod.
            console.log('::error::' + s.ime + ': ' + ((err && err.message) || String(err)));
            greske += 1;
            continue;
        }

        const oborena = r.pale.indexOf(s.tvrdnja) >= 0;

        if (!oborena) {
            console.log('::error::' + s.ime + ': NE OBARA SVOJ TEST -- "' + s.tvrdnja +
                        '" je ostala zelena pod sabotazom');
            greske += 1;
            continue;
        }

        if (r.pale.length === r.imena.length) {
            console.log('::error::' + s.ime + ': OBARA CEO SUITE (' + r.pale.length + '/' +
                        r.imena.length + ') -- verovatno je pokvarila ucitavanje, ne invarijantu');
            greske += 1;
            continue;
        }

        const ostale = r.pale.filter(function (x) { return x !== s.tvrdnja; });
        console.log('OK: ' + s.ime + ' -> crveno: "' + s.tvrdnja + '"' +
                    (ostale.length ? ' (uz ' + ostale.length + ' preklapajucih: ' +
                     ostale.join(' | ') + ')' : ''));
    }

    // Tvrdnja bez sabotaze nije greska (u rezu se dodaju samo NOVE sabotaze,
    // CLAUDE.md S5), ali mora da bude VIDLJIVA -- inace rupa raste necujno.
    const bez = [];
    SUITES.forEach(function (suite) {
        Object.keys(suite.tvrdnje).forEach(function (ime) {
            if (!pokriveno[suite.ime + ' :: ' + ime]) bez.push(suite.ime + ' :: ' + ime);
        });
    });

    console.log('');
    console.log('sabotaza: ' + (SABOTAZE.length - greske) + '/' + greske +
                ' (dokazano/palo) od ' + SABOTAZE.length);

    if (bez.length) {
        console.log('bez sabotaze (' + bez.length + '): ' + bez.join(' | '));
    }

    return greske === 0;
}

// IZLAZ IDE KROZ exitCode, NE KROZ process.exit (review #393, cetvrti krug).
//
// process.exit() nasilno prekida Node bez praznjenja event loop-a, pa bi greska
// iz zakasnelog timera nestala isto kao sto ju je pre gutao globalni handler --
// drugi mehanizam, ista klasa. Sa exitCode Node se zavrsi prirodno: sve
// zavrseno -> 0, pala tvrdnja -> 1, a zakasneo uncaught exception ili odbijen
// promise stvarno obori proces.
//
// Posledica ugovora: suite koji otvori pravi pozadinski posao mora sam da ga
// zatvori ili docekao -- inace prolaz visi umesto da se zavrsi.
(async function () {
    const selfTest = process.argv.indexOf('--self-test') >= 0;

    try {
        const ok = selfTest ? await dvosmerniDokaz() : await zelenaKapija();
        process.exitCode = ok ? 0 : 1;
    } catch (err) {
        console.log('::error::harness je pao van tvrdnje: ' + ((err && err.stack) || err));
        process.exitCode = 1;
    }
})();
