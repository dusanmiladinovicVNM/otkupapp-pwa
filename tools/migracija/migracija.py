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
from collections import Counter
import sys
import time
import unittest

OVDE = os.path.dirname(os.path.abspath(__file__))
if OVDE not in sys.path:
    sys.path.insert(0, OVDE)

import ambalaza  # noqa: E402
import cilj  # noqa: E402
import dokumenti  # noqa: E402
import pregled  # noqa: E402
import provere  # noqa: E402
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

    kanon = dokumenti.ucitaj_kanon()
    try:
        izuzeci = ambalaza.ucitaj_izuzetke(izuzeci_put)
        if a.knjizenje == ambalaza.ISPRAVNO:
            # Ista knjiga koju pise upis: robni dokumenti odredjuju svoju knjigu
            # (MIG-ODL-14), pa pregled ide kroz cilj, ne samo kroz prevod nogu.
            rez = cilj.izgradi(tabele, kanon, izuzeci, knjizenje=a.knjizenje,
                               izuzeci_dok=cilj.ucitaj_izuzetke_dokumenata(izuzeci_put)).amb
            if rez is None:
                raise ambalaza.GreskaAmbalaze("Sveska nema tabelu tblAmbalaza")
        else:
            rez = ambalaza.prevedi(tabele, izuzeci, knjizenje=a.knjizenje)
    except (ambalaza.GreskaAmbalaze, ValueError, OSError) as e:
        # ValueError pokriva i pokvaren izuzeci.json (json.JSONDecodeError).
        print("GRESKA: %s" % e, file=sys.stderr)
        return 2
    pokrivenost = ambalaza.proveri_pokrivenost(rez)
    salda = ambalaza.izracunaj_salda(rez)
    invarijante = ambalaza.proveri_invarijante(rez)
    stavke = ambalaza.stavke_po_nalogu(rez)

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


def _izlazni_folder(izvor, izlaz, ime):
    if izlaz:
        return os.path.abspath(izlaz)
    stem = os.path.splitext(os.path.basename(izvor))[0]
    return os.path.join(os.path.dirname(izvor), "Migracija_pregled",
                        "%s_%s_%s" % (stem, ime, time.strftime("%Y%m%d_%H%M%S")))


def napravi_cilj(izvor, izuzeci_put, knjizenje):
    """Procitaj izvor i sagradi novu svesku u memoriji, sa proverama.
    Vrati (tabele izvora, Cilj, sha pre, sha posle)."""
    sha_pre = sha256_fajla(izvor)
    tabele = procitaj_svesku(izvor)
    izuzeci = ambalaza.ucitaj_izuzetke(izuzeci_put)
    izuzeci_dok = cilj.ucitaj_izuzetke_dokumenata(izuzeci_put)
    c = cilj.izgradi(tabele, dokumenti.ucitaj_kanon(), izuzeci, knjizenje=knjizenje,
                     izuzeci_dok=izuzeci_dok)
    provere.proveri(c)
    provere.proveri_isplate_iz_starog(c, tabele)
    return tabele, c, sha_pre, sha256_fajla(izvor)


def komanda_proba(a):
    """FAZA 2, korak 1: nova sveska u memoriji + provere -- NISTA SE NE PISE u svesku.

    Paket (proba_upisa.xlsx) nosi svaku tabelu nove sveske kao Excel tabelu i
    list Provere. BLOKIRA znaci: posle upisa bi main pao na citanju cele tabele.
    """
    izvor = os.path.abspath(a.izvor)
    if not os.path.isfile(izvor):
        print("GRESKA: izvorna sveska ne postoji: %s" % izvor, file=sys.stderr)
        return 2
    izlaz = _izlazni_folder(izvor, a.izlaz, "proba")
    if _u_repou(izlaz):
        print("GRESKA: izlaz je unutar git repoa (%s) -- izaberi folder van repoa." % izlaz,
              file=sys.stderr)
        return 2
    print("Citam: %s" % izvor)
    try:
        tabele, c, sha_pre, sha_posle = napravi_cilj(
            izvor, None if a.bez_izuzetaka else (a.izuzeci or IZUZECI), a.knjizenje)
    except (GreskaCitanja, ambalaza.GreskaAmbalaze, ValueError, OSError) as e:
        print("GRESKA: %s" % e, file=sys.stderr)
        return 2
    os.makedirs(izlaz, exist_ok=True)
    tezine = Counter(n["tezina"] for n in c.nalazi)
    pravila = Counter((n["tezina"], n["pravilo"]) for n in c.nalazi)
    saz = [("Izvor", os.path.basename(izvor)), ("Izvor SHA256", sha_pre),
           ("Izvor nepromenjen posle citanja", "DA" if sha_pre == sha_posle else "NE -- PROVERI"),
           ("Knjizenje ambalaze", a.knjizenje), ("", "")]
    for t in ("BLOKIRA", "ODLUKA", "UPOZORENJE", "INFO"):
        saz.append(("Nalaza %s" % t, tezine.get(t, 0)))
    saz.append(("", ""))
    for (t, p), n in sorted(pravila.items()):
        saz.append(("  %s  %s" % (t, p), n))
    saz.append(("", ""))
    saz.append(("Tabela nove sveske / redova / iz starih redova", ""))
    for ime in sorted(c.tabele):
        if c.tabele[ime]:
            saz.append(("  " + ime, "%d   (staro: %s)" % (len(c.tabele[ime]),
                                                        c.izvor_redova.get(ime, "-"))))
    sv = pregled.Sveska()
    sv.dodaj_list("Sazetak", ["Stavka", "Vrednost"], saz, sirine=[60, 70])
    red_n = {"BLOKIRA": 0, "ODLUKA": 1, "UPOZORENJE": 2, "INFO": 3}
    sv.dodaj_list("Provere", ["Tezina", "Pravilo", "Tabela", "ID", "Opis"],
                  [[n["tezina"], n["pravilo"], n["tabela"], n["id"], n["opis"]]
                   for n in sorted(c.nalazi, key=lambda x: (red_n.get(x["tezina"], 9), x["pravilo"]))],
                  sirine=[12, 44, 22, 38, 100], tabela="Provere")
    for ime in sorted(c.tabele):
        redovi = c.tabele[ime]
        if not redovi:
            continue
        kol = c.kanon[ime]
        sv.dodaj_list(ime[3:] if ime.startswith("tbl") else ime, kol,
                      [[r.get(k, "") for k in kol] for r in redovi], tabela=ime)
    paket = os.path.join(izlaz, "proba_upisa.xlsx")
    sv.snimi(paket)
    with open(os.path.join(izlaz, "proba_sazetak.txt"), "w", encoding="utf-8") as f:
        for k, v in saz:
            f.write(("%-60s %s" % (k, v)).rstrip() + "\n")
        # BLOKIRA i ODLUKA ceo tekst: o njima se odlucuje pre upisa, pa sazetak
        # mora da kaze KOJI dokument ili nalog, ne samo koliko.
        f.write("\nBLOKIRA i ODLUKA -- pojedinacno:\n")
        for n in c.nalazi:
            if n["tezina"] in (provere.BLOKIRA, provere.ODLUKA):
                f.write("%s %s %s %s %s\n" % (n["tezina"], n["pravilo"], n["tabela"], n["id"], n["opis"]))
    for k, v in saz:
        print(("%-60s %s" % (k, v)).rstrip())
    print("\nPaket: %s" % paket)
    if sha_pre != sha_posle:
        return 2
    return 0


def komanda_upis(a):
    """FAZA 2, korak 2: upis u NOVU svesku preko Excela (Windows + Excel + pywin32).

    Odbija ako proba ima BLOKIRA nalaz; trazi --prihvatam-odluke kad ima ODLUKA
    nalaza (operater ih je procitao u probi). Kad cilj ne postoji, pravi ga
    tools/make_dev_workbook.py iz gita. Posle upisa svaka celija se poredi.
    """
    izvor = os.path.abspath(a.izvor)
    meta = os.path.abspath(a.cilj)
    if not os.path.isfile(izvor):
        print("GRESKA: izvorna sveska ne postoji: %s" % izvor, file=sys.stderr)
        return 2
    if os.path.normcase(meta) == os.path.normcase(izvor):
        print("GRESKA: cilj je izvorna sveska -- upis u staru svesku nije dozvoljen.", file=sys.stderr)
        return 2
    if _u_repou(meta):
        print("GRESKA: cilj je unutar git repoa -- nova sveska nosi podatke klijenta.", file=sys.stderr)
        return 2
    print("Citam i gradim novu svesku u memoriji: %s" % izvor)
    try:
        tabele, c, sha_pre, sha_posle = napravi_cilj(
            izvor, None if a.bez_izuzetaka else (a.izuzeci or IZUZECI), a.knjizenje)
    except (GreskaCitanja, ambalaza.GreskaAmbalaze, ValueError, OSError) as e:
        print("GRESKA: %s" % e, file=sys.stderr)
        return 2
    tezine = Counter(n["tezina"] for n in c.nalazi)
    if tezine.get("BLOKIRA"):
        print("UPIS ODBIJEN: %d BLOKIRA nalaza -- pokreni `proba` i vidi list Provere." % tezine["BLOKIRA"],
              file=sys.stderr)
        return 3
    if tezine.get("ODLUKA") and not a.prihvatam_odluke:
        print("UPIS ZAUSTAVLJEN: %d ODLUKA nalaza. Procitaj ih u probi (list Provere), pa ponovi sa "
              "--prihvatam-odluke." % tezine["ODLUKA"], file=sys.stderr)
        for (t, p), n in sorted(Counter((x["tezina"], x["pravilo"]) for x in c.nalazi).items()):
            if t == "ODLUKA":
                print("  %-50s %d" % (p, n), file=sys.stderr)
        return 4
    if os.name != "nt":
        print("Upis trazi Windows + Excel + pywin32.", file=sys.stderr)
        return 2
    if not os.path.exists(meta):
        os.makedirs(os.path.dirname(meta), exist_ok=True)
        print("Pravim svezu svesku iz gita: %s" % meta)
        sys.path.insert(0, os.path.join(KOREN_REPOA, "tools"))
        import make_dev_workbook
        if make_dev_workbook.napravi(meta.replace("/", "\\")) != 0:
            print("GRESKA: make_dev_workbook nije uspeo.", file=sys.stderr)
            return 2
    import upis
    print("Upisujem u: %s" % meta)
    preskocene = upis.upisi(meta, c)
    if sha256_fajla(izvor) != sha_pre:
        print("UPOZORENJE: izvorna sveska se promenila tokom rada.", file=sys.stderr)
    print("Poredim upisanu svesku sa planom ...")
    razlike = upis.uporedi(c, meta, preskocene)
    amb_razlike = _uporedi_salda(c, meta)
    izv = os.path.splitext(meta)[0] + "_upis_izvestaj.txt"
    with open(izv, "w", encoding="utf-8") as f:
        f.write("Izvor: %s\nSHA256: %s\nCilj: %s\n\n" % (izvor, sha_pre, meta))
        for ime in sorted(c.tabele):
            if c.tabele[ime]:
                f.write("%-28s %6d\n" % (ime, len(c.tabele[ime])))
        f.write("\nRazlike posle upisa: %d\n" % len(razlike))
        f.writelines(r + "\n" for r in razlike)
        f.write("\nSalda ambalaze nova sveska vs prevod: %d razlika\n" % len(amb_razlike))
        f.writelines(r + "\n" for r in amb_razlike)
        f.write("\nNalazi probe:\n")
        f.writelines("%s %s %s %s %s\n" % (n["tezina"], n["pravilo"], n["tabela"], n["id"], n["opis"])
                     for n in c.nalazi)
    print("Razlike posle upisa: %d tabela   salda ambalaze: %d razlika"
          % (len({r.split(" ", 1)[0].rstrip(":") for r in razlike}), len(amb_razlike)))
    print("Izvestaj: %s" % izv)
    if razlike or amb_razlike:
        print("UPIS NIJE TACAN -- ne koristi novu svesku; posalji izvestaj.", file=sys.stderr)
        return 5
    print("\nSledece: otvori novu svesku, Alt+F11 -> Debug -> Compile VBAProject, pa pokreni program.")
    return 0


def _uporedi_salda(c, meta):
    """Salda iz upisane knjige (AMB-10 oblik) moraju biti ista kao iz prevoda."""
    if c.amb is None:
        return []
    nova = ambalaza.prevedi(procitaj_svesku(meta))
    a = {(s["nalog_tip"], ambalaza.kljuc(s["nalog_id"]), ambalaza.norm_tip(s["tip"])): s["novi"]
         for s in ambalaza.izracunaj_salda(c.amb)}
    b = {(s["nalog_tip"], ambalaza.kljuc(s["nalog_id"]), ambalaza.norm_tip(s["tip"])): s["novi"]
         for s in ambalaza.izracunaj_salda(nova)}
    return ["%s: prevod %s, nova sveska %s" % ("/".join(k), a.get(k, 0), b.get(k, 0))
            for k in sorted(set(a) | set(b)) if a.get(k, 0) != b.get(k, 0)]


def komanda_self_test():
    testovi = unittest.defaultTestLoader.discover(OVDE, pattern="test_*.py", top_level_dir=OVDE)
    r = unittest.TextTestRunner(verbosity=1).run(testovi)
    return 0 if r.wasSuccessful() else 1


def main(argv=None):
    p = argparse.ArgumentParser(description="Migracija stare AgriX sveske (faza 1: pregled, samo cita).")
    p.add_argument("--self-test", action="store_true", help="pokreni testove i izadji")
    sub = p.add_subparsers(dest="komanda")
    pr = sub.add_parser("proba", help="faza 2: nova sveska u memoriji + provere (samo cita)")
    pr.add_argument("--izvor", required=True, help="stara .xlsm sveska")
    pr.add_argument("--izlaz", help="folder za paket (podrazumevano: pored izvora)")
    pr.add_argument("--izuzeci", help="izuzeci.json (podrazumevano: tools/migracija/izuzeci.json)")
    pr.add_argument("--bez-izuzetaka", action="store_true", help="ne primenjuj nijedan izuzetak")
    pr.add_argument("--knjizenje", choices=(ambalaza.ISPRAVNO, ambalaza.VERNO), default=ambalaza.ISPRAVNO)
    pu = sub.add_parser("upis", help="faza 2: upis u NOVU svesku preko Excela (Windows)")
    pu.add_argument("--izvor", required=True, help="stara .xlsm sveska")
    pu.add_argument("--cilj", required=True, help="nova .xlsm (ako ne postoji, pravi se iz gita)")
    pu.add_argument("--izuzeci", help="izuzeci.json (podrazumevano: tools/migracija/izuzeci.json)")
    pu.add_argument("--bez-izuzetaka", action="store_true", help="ne primenjuj nijedan izuzetak")
    pu.add_argument("--knjizenje", choices=(ambalaza.ISPRAVNO, ambalaza.VERNO), default=ambalaza.ISPRAVNO)
    pu.add_argument("--prihvatam-odluke", action="store_true",
                    help="procitao sam ODLUKA nalaze probe i prihvatam podrazumevana pravila")
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
    if a.komanda == "upis":
        return komanda_upis(a)
    if a.komanda == "proba":
        return komanda_proba(a)
    if a.komanda == "pregled":
        return komanda_pregled(a)
    p.print_help()
    return 2


if __name__ == "__main__":
    sys.exit(main())
