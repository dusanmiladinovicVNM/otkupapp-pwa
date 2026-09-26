'use strict';

// ============================================================
// STAVKE OTKUPA -- jedno mesto koje zna sta je linija dokumenta (S5-5b)
// ============================================================
//
// Zapis otkupa od ovog reza NE nosi klasu, kolicinu, cenu i ambalazu na sebi:
// nosi `stavke[]`, red po klasi. Oblik lokalnog zapisa je oblik ZICE, jer
// sync-engine salje zapise verbatim -- nema per-record transform hook-a, pa bi
// drugi lokalni oblik znacio nov sloj samo za privremeno stanje.
//
// UI u ovom rezu i dalje unosi JEDNU klasu (odluka: samo zica, N=1). Pristupnici
// zato rade i za N=1 i za N>1: prikaz se ne menja, a model je vec tacan. Kad
// forma dobije vise klasa, nijedan citalac se ne menja.
//
// CITAOCI NE SABIRAJU SAMI. Dva mesta koja sabiraju istu robu se razidju, i to
// se vidi tek na iznosu koji neko isplacuje. Zato su zbirovi ovde, jednom.

// Identitet zapisa. Zivi uz stavke jer i stavka ima SVOJ ClientRecordID: master
// joj ne zna OtkupStavkaID dok je ne upise (CreateOtkup_TX), pa je CRID reda
// jedini nacin da ponovljen sync ne udvoji robu.
function generateClientRecordID() {
    if (window.crypto && typeof window.crypto.randomUUID === 'function') {
        return window.crypto.randomUUID();
    }

    return 'loc-' + Date.now() + '-' + Math.floor(Math.random() * 1000000);
}

// Nova stavka sa svojim identitetom. brutoKg se nosi SAMO kad je unos bio bruto:
// nula nije "bruto = 0" nego "nije bruto unos", pa pisac u masteru bruto tada i
// ne cuva.
function novaStavkaOtkupa(polja) {
    const p = polja || {};
    const stavka = {
        clientRecordID: generateClientRecordID(),
        klasa: String(p.klasa || 'I').trim(),
        kolicina: Number(p.kolicina || 0),
        cena: Number(p.cena || 0),
        kolAmbalaze: Number(p.kolAmbalaze || 0)
    };

    const bruto = Number(p.brutoKg || 0);
    if (bruto > 0) stavka.brutoKg = bruto;

    return stavka;
}

// Stavke zapisa, uvek niz. Prazan niz znaci "ovaj zapis ih ne nosi", i pozivalac
// odlucuje sta to znaci -- pristupnik ne izmislja podatak.
function otkupStavke(rec) {
    if (!rec) return [];
    return Array.isArray(rec.stavke) ? rec.stavke : [];
}

function otkupZbirKg(rec) {
    return otkupStavke(rec).reduce(function (z, s) {
        return z + Number(s && s.kolicina || 0);
    }, 0);
}

function otkupZbirVrednosti(rec) {
    return otkupStavke(rec).reduce(function (z, s) {
        return z + Number(s && s.kolicina || 0) * Number(s && s.cena || 0);
    }, 0);
}

function otkupZbirAmbalaze(rec) {
    return otkupStavke(rec).reduce(function (z, s) {
        return z + Number(s && s.kolAmbalaze || 0);
    }, 0);
}

// Klase dokumenta kao tekst, kanonskim redom (I, II, III) a ne redom stizanja --
// isto pravilo po kom master dodeljuje RedniBroj.
function otkupKlaseTekst(rec) {
    const red = ['I', 'II', 'III'];

    return otkupStavke(rec)
        .map(function (s) { return String(s && s.klasa || '').trim(); })
        .filter(Boolean)
        .sort(function (a, b) {
            const ia = red.indexOf(a);
            const ib = red.indexOf(b);
            return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib);
        })
        .join(', ');
}

// Cena kad dokument ima TACNO JEDNU klasu; inace null.
//
// Vraca null namerno: dokument od dve klase ima dve cene, pa jedan broj tu ne
// postoji. Pozivalac koji prikazuje "Cena" mora da odluci sta ce sa tim -- tiho
// uzimanje prve cene bi na dvoklasnom dokumentu prikazalo pogresan podatak.
function otkupCenaAkoJedna(rec) {
    const stavke = otkupStavke(rec);
    if (stavke.length !== 1) return null;
    return Number(stavke[0].cena || 0);
}

// Jedina stavka, ili null. Ista logika kao otkupCenaAkoJedna.
function otkupJednaStavka(rec) {
    const stavke = otkupStavke(rec);
    return stavke.length === 1 ? stavke[0] : null;
}

window.generateClientRecordID = generateClientRecordID;
window.novaStavkaOtkupa = novaStavkaOtkupa;
window.otkupStavke = otkupStavke;
window.otkupZbirKg = otkupZbirKg;
window.otkupZbirVrednosti = otkupZbirVrednosti;
window.otkupZbirAmbalaze = otkupZbirAmbalaze;
window.otkupKlaseTekst = otkupKlaseTekst;
window.otkupCenaAkoJedna = otkupCenaAkoJedna;
window.otkupJednaStavka = otkupJednaStavka;
