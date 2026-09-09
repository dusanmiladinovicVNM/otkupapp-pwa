"""Poredi svesku sa registrom seme (src-vba/modSchema.bas). SAMO cita.

Odgovara na jedino pitanje koje se postavlja pre uvoza novog koda u zatecenu
svesku: da li ce EnsureAllTables nesto DODATI, i da li sveska ima nesto sto
registar ne zna.

    python tools/schema_diff.py "C:/.../AgriX_C002.xlsm"

Tri ishoda, i sva tri su normalna -- vazno je da se VIDE pre uvoza:

  FALI U SVESCI   registar ima, sveska nema.
                  EnsureAllTables ce to napraviti. Bezbedno (samo dodaje).
  VISAK U SVESCI  sveska ima, registar ne zna.
                  NISTA se ne brise -- ali registar je nepotpun, pa te kolone
                  VerifySchema nikad nece cuvati. Regenerisi registar iz OVE
                  sveske ako su legitimne.
  POZICIONI HAZARD  kanon i sveska se razilaze PRE kraja (premestena,
                  izbacena iz sredine, ili ubacena u sredinu). NIJE bezopasno.
                  modDataAccess.AppendRow pise POZICIONO, a pisci poput
                  modOtkup.SaveOtkup grade goli Array(...) sa 22 vrednosti.
                  Preraspored tiho salje vrednosti u pogresne kolone -- gore
                  od pada upisa. Citanje jeste po imenu (GetColumnIndex);
                  upis nije.

Exit: 0 = nema razlike, 1 = ima razlike (nalaz, odluci sam),
      2 = kvar ILI razlika redosleda (uvoz nebezbedan).

Windows + Excel + pywin32. Sveska se otvara read-only i nikad ne snima.
"""

import argparse
import os
import re
import sys

MSO_AUTOMATION_SECURITY_LOW = 1

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
KANON = os.path.join(ROOT, "schema", "schema.json")


def registar() -> dict:
    """tblIme -> (sheet, [kolone]) iz KANONA.

    Cita schema/schema.json, ne generisani modSchema.bas: kanon je izvor istine,
    a artefakt sme da bude zastareo (to hvata gen_schema_module.py --check).
    Alat koji poredi svesku mora da ide na izvor.
    """
    import json
    with open(KANON, encoding="utf-8") as fh:
        d = json.load(fh)
    return {t["table"]: (t["sheet"], t["columns"]) for t in d["tables"]}


def sveska(path: str) -> dict:
    try:
        import win32com.client as win32
    except ImportError:
        print("pywin32 nije instaliran: python -m pip install pywin32",
              file=sys.stderr)
        return None

    path = os.path.abspath(path)
    if not os.path.exists(path):
        print("Sveska ne postoji: " + path, file=sys.stderr)
        return None

    xl = win32.DispatchEx("Excel.Application")
    wb = None
    try:
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.AutomationSecurity = MSO_AUTOMATION_SECURITY_LOW
        xl.EnableEvents = False          # Workbook_Open / StartApp se NE pokrece
        wb = xl.Workbooks.Open(path, ReadOnly=True, UpdateLinks=0)

        out = {}
        for ws in wb.Worksheets:
            for lo in ws.ListObjects:
                out[lo.Name] = (ws.Name,
                                [c.Name for c in lo.ListColumns])
        return out
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        xl.Quit()


def main(argv) -> int:
    ap = argparse.ArgumentParser(description="Poredi svesku sa registrom seme.")
    ap.add_argument("sveska", help="putanja do .xlsm")
    ap.add_argument("--tiho", action="store_true",
                    help="ispisi samo zbir, bez spiska")
    a = ap.parse_args(argv[1:])

    reg = registar()
    if not reg:
        print("Registar je prazan -- modSchema.bas nije procitan.", file=sys.stderr)
        return 2

    wbs = sveska(a.sveska)
    if wbs is None:
        return 2

    fali_tab = sorted(set(reg) - set(wbs))
    visak_tab = sorted(set(wbs) - set(reg))
    fali_kol, visak_kol, redosled = [], [], []

    for tbl in sorted(set(reg) & set(wbs)):
        r_cols, w_cols = reg[tbl][1], wbs[tbl][1]
        for c in r_cols:
            if c not in w_cols:
                fali_kol.append((tbl, c))
        for c in w_cols:
            if c not in r_cols:
                visak_kol.append((tbl, c))

        # POZICIONA BEZBEDNOST je pitanje PREFIKSA, ne skupa.
        #
        # AppendRow pise poziciono, pa je bezbedno samo ako je jedan spisak
        # prefiks drugog:
        #   kanon A B C  |  sveska A B C X Y   -> OK (visak je samo REP)
        #   kanon A B C D|  sveska A B C       -> OK (fali samo REP; leci se)
        # Sve ostalo pomera bar jednu poziciju:
        #   kanon A B C D|  sveska A C D       -> B fali IZ SREDINE
        #   kanon A B C D|  sveska A X B C D   -> X ubacen U SREDINU
        # Ranija provera je gledala samo cistu permutaciju istog skupa, pa je
        # oba gornja slucaja propustala.
        n = min(len(r_cols), len(w_cols))
        if r_cols[:n] != w_cols[:n]:
            redosled.append(tbl)

    print("Registar: %d tabela, %d kolona"
          % (len(reg), sum(len(v[1]) for v in reg.values())))
    print("Sveska:   %d tabela, %d kolona"
          % (len(wbs), sum(len(v[1]) for v in wbs.values())))
    print()

    def sekcija(naslov, stavke, opis):
        print("%s (%d)%s" % (naslov, len(stavke), "" if stavke else " -- nema"))
        if stavke and not a.tiho:
            print("  " + opis)
            for s in stavke:
                print("    " + (s if isinstance(s, str) else "%s.%s" % s))
        print()

    sekcija("FALI U SVESCI -- tabela", fali_tab,
            "EnsureAllTables ce ih napraviti.")
    sekcija("FALI U SVESCI -- kolona", fali_kol,
            "EnsureAllTables ce ih dodati (samo dodaje, nista ne brise).")
    sekcija("VISAK U SVESCI -- tabela", visak_tab,
            "Registar ih ne zna: VerifySchema ih nece cuvati. Regenerisi registar "
            "iz OVE sveske ako su legitimne.")
    sekcija("VISAK U SVESCI -- kolona", visak_kol,
            "Isto: nista se ne brise, ali ih registar ne cuva.")
    sekcija("POZICIONI HAZARD -- BLOKIRA UVOZ", redosled,
            "Upis je POZICION (AppendRow): vrednosti bi otisle u pogresne kolone. "
            "Mora se resiti pre uvoza -- EnsureAllTables ovo NE popravlja, jer bi "
            "premestanje kolone pomerilo podatke.")

    ukupno = (len(fali_tab) + len(fali_kol) + len(visak_tab)
              + len(visak_kol) + len(redosled))
    if ukupno == 0:
        print("REZULTAT: sveska i registar se poklapaju.")
        return 0

    if redosled:
        print("REZULTAT: %d razlika, od toga %d POZICIONI HAZARD -- uvoz je "
              "NEBEZBEDAN dok se ne resi." % (ukupno, len(redosled)))
        return 2

    print("REZULTAT: %d razlika. Nije greska -- nalaz. Odluci pre uvoza." % ukupno)
    return 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
