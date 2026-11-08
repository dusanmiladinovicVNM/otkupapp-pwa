"""Pokretac provera: ucita svesku, pusti svaki modul, skupi nalaze. SAMO CITA."""

import importlib
import traceback
from collections import Counter

from citac import procitaj_svesku, sha256_fajla

from provera.zajednicko import GRESKA, INFO, SUMNJA, Kontekst

MODULI = ("prijemnica", "otpremnica_zbirna", "otkup", "novac", "amb_pal_fak", "storno")

# Kodovi razlicitih modula koji za ISTI red opisuju istu stvar. Ostaje prvi (po redu
# modula), uz napomenu ko je jos prijavio; ozbiljnost je najveca od spojenih.
ISTO = (
    frozenset(("ZBR_ZIVA_BEZ_PRIJEMNICE", "ZBR_TURA_BEZ_PRIJEMNICE")),
)
_RANG = {GRESKA: 0, SUMNJA: 1, INFO: 2}


def spoji_iste(nalazi):
    grupa = {kod: i for i, g in enumerate(ISTO) for kod in g}
    prvi, out = {}, []
    for x in nalazi:
        g = grupa.get(x.kod)
        if g is None:
            out.append(x)
            continue
        kljuc_ = (g, x.tabela, x.id)
        if kljuc_ not in prvi:
            prvi[kljuc_] = x
            out.append(x)
            continue
        z = prvi[kljuc_]
        z.opis += " [Isto prijavljuje i %s: %s %s.]" % (x.modul, x.kod, x.ozbiljnost)
        if _RANG[x.ozbiljnost] < _RANG[z.ozbiljnost]:
            z.ozbiljnost = x.ozbiljnost
        if not z.predlog and x.predlog:
            z.predlog = x.predlog
    return out


def pokreni_nad_tabelama(tabele, moduli=MODULI):
    """Kontekst sa nalazima svih modula. Modul koji padne ne obara ostale:
    njegov pad je nalaz MODUL_PAO (GRESKA) sa tracebackom u opisu. Nalazi
    razlicitih modula o istoj stvari se spajaju (ISTO)."""
    ctx = Kontekst(tabele)
    for ime in moduli:
        ctx.modul(ime)
        try:
            mod = importlib.import_module("provera." + ime)
        except ModuleNotFoundError as e:
            if e.name == "provera." + ime:
                continue
            raise
        try:
            mod.provere(ctx)
        except Exception:  # noqa: BLE001 -- pad jednog modula je nalaz, ne kraj provere
            ctx.nalaz("MODUL_PAO", GRESKA, "", ime, traceback.format_exc(limit=6))
    ctx.nalazi = spoji_iste(ctx.nalazi)
    return ctx


def pokreni(putanja, moduli=MODULI):
    sha_pre = sha256_fajla(putanja)
    tabele = procitaj_svesku(putanja)
    ctx = pokreni_nad_tabelama(tabele, moduli)
    if sha256_fajla(putanja) != sha_pre:
        raise RuntimeError("Sveska se promenila tokom citanja -- izvestaj ne vazi.")
    ctx.sha = sha_pre
    return ctx


def sazetak(ctx):
    c = Counter((x.ozbiljnost, x.kod, x.modul) for x in ctx.nalazi)
    red = {GRESKA: 0, SUMNJA: 1, INFO: 2}
    return sorted(((o, kod, mod, b) for (o, kod, mod), b in c.items()), key=lambda x: (red[x[0]], x[2], x[1]))
