"""Pokretac provera: ucita svesku, pusti svaki modul, skupi nalaze. SAMO CITA."""

import importlib
import traceback
from collections import Counter

from citac import procitaj_svesku, sha256_fajla

from provera.zajednicko import GRESKA, INFO, SUMNJA, Kontekst

MODULI = ("prijemnica", "otpremnica_zbirna", "otkup", "novac", "amb_pal_fak", "storno")


def pokreni_nad_tabelama(tabele, moduli=MODULI):
    """Kontekst sa nalazima svih modula. Modul koji padne ne obara ostale:
    njegov pad je nalaz MODUL_PAO (GRESKA) sa tracebackom u opisu."""
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
