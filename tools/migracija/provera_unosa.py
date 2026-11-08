"""Provera unosa stare sveske -- greske koje je unos ostavio u podacima. SAMO CITA.

Iz root-a repoa, u Git Bash-u:

    python tools/migracija/provera_unosa.py "C:/.../stara.xlsm"

Izvestaj ide u folder pored sveske: Migracija_pregled/<ime>_provera_unosa.xlsx
(nikad u repo -- nosi stvarne podatke). Sveska se ne menja: SHA256 pre i posle
citanja mora biti isti.

Ozbiljnost: GRESKA = podatak je sigurno pogresan; SUMNJA = potpis greske, ali
postoji i legitimno objasnjenje (operater proverava); INFO = skrece paznju.
Klase gresaka, njihov uzrok u starom kodu i potpis u podacima: provera/README.md.
"""

import argparse
import os
import sys

OVDE = os.path.dirname(os.path.abspath(__file__))
if OVDE not in sys.path:
    sys.path.insert(0, OVDE)

import pregled  # noqa: E402
from provera.pokretac import pokreni, sazetak  # noqa: E402
from provera.zajednicko import KOLONE_NALAZA  # noqa: E402


def izvestaj(putanja, ctx):
    folder = os.path.join(os.path.dirname(os.path.abspath(putanja)), "Migracija_pregled")
    os.makedirs(folder, exist_ok=True)
    cilj = os.path.join(folder, os.path.splitext(os.path.basename(putanja))[0] + "_provera_unosa.xlsx")
    sv = pregled.Sveska()
    sv.dodaj_list("Sazetak", ["Ozbiljnost", "Kod", "Modul", "Broj"], [list(x) for x in sazetak(ctx)],
                  tabela="tblSazetak")
    red = {"GRESKA": 0, "SUMNJA": 1, "INFO": 2}
    nalazi = sorted(ctx.nalazi, key=lambda x: (red[x.ozbiljnost], x.modul, x.kod, x.tabela, x.id))
    sv.dodaj_list("Nalazi", KOLONE_NALAZA, [x.red() for x in nalazi], tabela="tblNalazi")
    sv.snimi(cilj)
    return cilj


def main(argv=None):
    p = argparse.ArgumentParser(description="Provera unosa stare sveske (samo cita).")
    p.add_argument("sveska")
    a = p.parse_args(argv)
    ctx = pokreni(a.sveska)
    print("Sveska:        %s" % a.sveska)
    print("SHA256:        %s (nepromenjena posle citanja)" % ctx.sha)
    print("Nalaza ukupno: %d" % len(ctx.nalazi))
    print()
    for oz, kod, mod, b in sazetak(ctx):
        print("  %-7s %-48s %-18s %5d" % (oz, kod, mod, b))
    cilj = izvestaj(a.sveska, ctx)
    print()
    print("Izvestaj: %s" % cilj)
    return 0


if __name__ == "__main__":
    sys.exit(main())
