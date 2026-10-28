"""Migracija stare AgriX sveske u novu -- jedan alat umesto lanca MIG-01..MIG-22.

FAZA 1 (ovo): `pregled` -- SAMO CITA staru svesku i pravi Excel pregled:

    python tools/migracija/migracija.py pregled --izvor "C:/putanja/stara.xlsm"

  * stara knjiga ambalaze prevedena u AMB-10 kretanja (u memoriji);
  * svaki kooperant, stanica, vozac i kupac: stavka po stavka, novi i stari
    tekuci saldo, najmanji dnevni saldo;
  * --knjizenje ispravno (podrazumevano): dodato ULAZ_TUDJE za kooperante i
    kupce i NABAVKA za stanice, tako da nijedan nalog ni jednog dana nije
    ispod nule -- oblik u kome je nova sveska spremna za rad; --knjizenje
    verno: tacno kao stara sveska, sa predlogom preseka na kraju;
  * sta ide ZA ODLUKU (red bez pravila, nespojena noga, storniran dokument...);
  * dokumenti sa vise redova (odluka 1) i tabele naspram kanona.

Izvorna sveska se otvara samo za citanje; SHA256 pre i posle mora biti isti.
Izlaz ide u folder pored izvora (Migracija_pregled/...), NIKAD u git repo:
pregled nosi stvarne podatke klijenta.

FAZA 2 (upis u novu svesku iz gita) dolazi tek posle odluka 1-5 -- pre-flight
kapija (docs/DOMEN, AMB-INV-07/08, header+stavke) je za nju zatvorena.

    python tools/migracija/migracija.py --self-test     testovi bez Excela

Ne trazi Excel ni pywin32; radi i u Linux sesiji i u CI-ju.
"""

import argparse
import os
import sys
import time
import unittest

OVDE = os.path.dirname(os.path.abspath(__file__))
if OVDE not in sys.path:
    sys.path.insert(0, OVDE)

import ambalaza  # noqa: E402
import dokumenti  # noqa: E402
import pregled  # noqa: E402
from citac import GreskaCitanja, procitaj_svesku, sha256_fajla  # noqa: E402

IZUZECI = os.path.join(OVDE, "izuzeci.json")
KOREN_REPOA = os.path.dirname(os.path.dirname(OVDE))


def _u_repou(putanja):
    # normcase: na Windows-u su C:\Users i c:\users ista putanja.
    p = os.path.normcase(os.path.realpath(putanja))
    r = os.path.normcase(os.path.realpath(KOREN_REPOA))
    return p == r or p.startswith(r + os.sep)


def komanda_pregled(a):
    izvor = os.path.abspath(a.izvor)
    if not os.path.isfile(izvor):
        print("GRESKA: izvorna sveska ne postoji: %s" % izvor, file=sys.stderr)
        return 2
    vreme = time.strftime("%Y%m%d_%H%M%S")
    stem = os.path.splitext(os.path.basename(izvor))[0]
    izlaz = os.path.abspath(a.izlaz) if a.izlaz else os.path.join(
        os.path.dirname(izvor), "Migracija_pregled", "%s_%s" % (stem, vreme))
    if _u_repou(izlaz):
        print("GRESKA: izlaz je unutar git repoa (%s). Pregled nosi stvarne podatke klijenta --\n"
              "izaberi folder van repoa (--izlaz)." % izlaz, file=sys.stderr)
        return 2
    izuzeci_put = None if a.bez_izuzetaka else (a.izuzeci or IZUZECI)

    print("Citam: %s" % izvor)
    sha_pre = sha256_fajla(izvor)
    try:
        tabele = procitaj_svesku(izvor)
    except (GreskaCitanja, OSError) as e:
        print("GRESKA citanja: %s" % e, file=sys.stderr)
        print("Ako je sveska otvorena u Excelu, zatvori je i pokusaj ponovo.", file=sys.stderr)
        return 2
    print("  tabela: %d" % len(tabele))

    try:
        izuzeci = ambalaza.ucitaj_izuzetke(izuzeci_put)
        rez = ambalaza.prevedi(tabele, izuzeci, knjizenje=a.knjizenje)
    except (ambalaza.GreskaAmbalaze, ValueError, OSError) as e:
        # ValueError pokriva i pokvaren izuzeci.json (json.JSONDecodeError).
        print("GRESKA: %s" % e, file=sys.stderr)
        return 2
    pokrivenost = ambalaza.proveri_pokrivenost(rez)
    salda = ambalaza.izracunaj_salda(rez)
    invarijante = ambalaza.proveri_invarijante(rez)
    stavke = ambalaza.stavke_po_nalogu(rez)

    kanon = dokumenti.ucitaj_kanon()
    popis = dokumenti.popis_tabela(tabele, kanon)
    oblik_dok = dokumenti.oblik_dokumenata(tabele)
    dok_sazetak, dok_grupe = dokumenti.identitet_dokumenata(tabele)

    sha_posle = sha256_fajla(izvor)
    os.makedirs(izlaz, exist_ok=True)
    xlsx = os.path.join(izlaz, "pregled_migracije.xlsx")
    izvor_info = dict(ime=os.path.basename(izvor), sha=sha_pre, nepromenjen=sha_pre == sha_posle,
                      vreme=time.strftime("%Y-%m-%d %H:%M:%S"))
    saz = pregled.napravi_izvestaj(xlsx, izvor_info, rez, salda, stavke, invarijante, pokrivenost,
                                   popis, dok_sazetak, dok_grupe, oblik_dok)
    # Samo nazivi tabela i kolona i broj redova -- bez ijednog podatka, pa se
    # sme poslati dalje (ulaz za upis u novu svesku).
    with open(os.path.join(izlaz, "tabele.txt"), "w", encoding="utf-8") as f:
        for t in sorted(tabele.values(), key=lambda x: x.ime.casefold()):
            f.write("%s (%d redova): %s\n" % (t.ime, len(t.redovi), ", ".join(t.kolone)))
    txt = os.path.join(izlaz, "sazetak.txt")
    with open(txt, "w", encoding="utf-8") as f:
        for k, v in saz:
            f.write(("%-55s %s" % (k, v)).rstrip() + "\n")

    for k, v in saz:
        print(("%-55s %s" % (k, v)).rstrip())
    print("\nPregled: %s" % xlsx)
    if sha_pre != sha_posle:
        print("UPOZORENJE: izvorna sveska se promenila tokom citanja (neko je radi u Excelu?).",
              file=sys.stderr)
        return 2
    if any(pokrivenost.values()) or any(s["razlika"] != 0 for s in salda):
        print("GRESKA ALATA: pokrivenost ili salda nisu tacna -- pregled se ne sme koristiti.",
              file=sys.stderr)
        return 3
    return 0


def komanda_self_test():
    testovi = unittest.defaultTestLoader.discover(OVDE, pattern="test_*.py", top_level_dir=OVDE)
    r = unittest.TextTestRunner(verbosity=1).run(testovi)
    return 0 if r.wasSuccessful() else 1


def main(argv=None):
    p = argparse.ArgumentParser(description="Migracija stare AgriX sveske (faza 1: pregled, samo cita).")
    p.add_argument("--self-test", action="store_true", help="pokreni testove i izadji")
    sub = p.add_subparsers(dest="komanda")
    pp = sub.add_parser("pregled", help="napravi Excel pregled iz stare sveske (samo cita)")
    pp.add_argument("--izvor", required=True, help="stara .xlsm sveska")
    pp.add_argument("--izlaz", help="folder za pregled (podrazumevano: pored izvora)")
    pp.add_argument("--izuzeci", help="izuzeci.json (podrazumevano: tools/migracija/izuzeci.json)")
    pp.add_argument("--bez-izuzetaka", action="store_true", help="ne primenjuj nijedan izuzetak")
    pp.add_argument("--knjizenje", choices=(ambalaza.ISPRAVNO, ambalaza.VERNO), default=ambalaza.ISPRAVNO,
                    help="ispravno (podrazumevano): istorija + ULAZ_TUDJE za partnere i NABAVKA za "
                         "stanice, nijedan nalog ispod nule; verno: tacno kao stara sveska")
    a = p.parse_args(argv)
    if a.self_test:
        return komanda_self_test()
    if a.komanda == "pregled":
        return komanda_pregled(a)
    p.print_help()
    return 2


if __name__ == "__main__":
    sys.exit(main())
