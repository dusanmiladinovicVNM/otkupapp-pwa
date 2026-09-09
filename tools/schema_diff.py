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
  RAZLIKA REDOSLEDA  ista imena, drugi raspored. Bezobrazno bezopasno:
                  ceo kod cita kolone po IMENU (GetColumnIndex), ne po indeksu.

Exit: 0 = nema razlike, 1 = ima razlike (nije greska, nego nalaz), 2 = kvar.

Windows + Excel + pywin32. Sveska se otvara read-only i nikad ne snima.
"""

import argparse
import os
import re
import sys

MSO_AUTOMATION_SECURITY_LOW = 1

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCHEMA_BAS = os.path.join(ROOT, "src-vba", "modSchema.bas")
MODCONFIG = os.path.join(ROOT, "src-vba", "modConfig.bas")

SPEC_POC = re.compile(r"^Private Sub (Spec\w+)\(ByVal reg As Object\)")
KOL = re.compile(r'^\s*k\.Add "([^"]+)"')
REG_RED = re.compile(r'^\s*Reg reg, (TBL_\w+), "([^"]+)", k')
TBL_CONST = re.compile(r'^Public Const (TBL_\w+)\s+As String\s*=\s*"(\w+)"')


def registar() -> dict:
    """tblIme -> (sheet, [kolone]) iz modSchema.bas, bez Excela."""
    const2tbl = {}
    with open(MODCONFIG, encoding="ascii", errors="replace") as fh:
        for l in fh:
            m = TBL_CONST.match(l)
            if m:
                const2tbl[m.group(1)] = m.group(2)

    out, kolone = {}, []
    with open(SCHEMA_BAS, encoding="ascii", errors="replace") as fh:
        for l in fh:
            l = l.rstrip("\r\n")
            if SPEC_POC.match(l):
                kolone = []
                continue
            m = KOL.match(l)
            if m:
                kolone.append(m.group(1))
                continue
            m = REG_RED.match(l)
            if m:
                tbl = const2tbl.get(m.group(1), m.group(1))
                out[tbl] = (m.group(2), kolone)
                kolone = []
    return out


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
        if r_cols != w_cols and sorted(r_cols) == sorted(w_cols):
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
    sekcija("RAZLIKA REDOSLEDA", redosled,
            "Bezopasno: kod cita kolone po IMENU (GetColumnIndex), ne po indeksu.")

    ukupno = (len(fali_tab) + len(fali_kol) + len(visak_tab)
              + len(visak_kol) + len(redosled))
    if ukupno == 0:
        print("REZULTAT: sveska i registar se poklapaju.")
        return 0

    print("REZULTAT: %d razlika. Nije greska -- nalaz. Odluci pre uvoza." % ukupno)
    return 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
