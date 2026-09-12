"""Kapija nad markerom prekinutog VBA importa (pending + mutated).

Postoji zato sto ove invarijante suite NE MOZE da izmeri: `modVbaTools` je
SELF_MODULE i ni `run_vba.py` ni `ImportAllVBA` ga ne uvoze, pa nijedan test u
svesci ne vidi njegov kod. Isti razlog kao za kapije u `.claude/rules/testovi.md`
S8/S9 -- greska se ne vidi kao crven test nego kao sveska koja se vise ne moze
snimiti, ili kao tiho zabetonirano ostecenje.

Sta se cuva:

  RESET       BeginImportTransaction NE SME da resetuje "mutated". Nov pokusaj
              importa nije dokaz da je raniji popravljen -- izvrsava se PRE
              teardown-a i PRE ValidateFormDesigner, pa prolaz koji tu padne nije
              popravio nista. Reset bi zatecen dokaz stete proglasio bezbednim.
              (Tacno ta greska je bila u prvoj verziji, PR #313.)

  NASLEDJIVANJE
              mMutated NE SME da krene od False. FAIL grana radi "If Not
              mMutated Then ClearImportPhase2State", a to brise CELU sekciju --
              dakle i sticky "mutated" iz ranijeg prolaza. Prolaz koji padne pre
              sopstvene mutacije bi tako obrisao dokaz TUDJE, jos nepopravljene
              stete. Isti efekat kao RESET, samo napisan drugacije -- i prva
              verzija ove kapije ga NIJE videla, pa je bila zelena nad kodom koji
              krsi invarijantu.

  MUTACIJA    "mutated" = "1" mora da se upise na OBE tacke mutacije: u fazi 1
              (pre prve destruktivne operacije) i na ulasku u fazu 2.

  EVENTI      Od prve mutacije nadalje Application.EnableEvents mora biti ON.
              Jedina globalna kapija nad Save-om je ThisWorkbook.Workbook_BeforeSave;
              ako VBE prekine izvrsavanje dok su eventi ugaseni,
              RestoreRuntimeAfterImport se nikad ne izvrsi i Ctrl+S prolazi
              neometano nad mozda nepotpunim projektom.

  ODLUKA      modImportState.MarkerBlokira mora da trazi OBA kljuca. Sa samo
              "pending" kapija blokira i sveske koje nikad nisu dirnute --
              mereno 12.09.2026: 14 od 14 sekcija u registru je imalo pending=1.

Exit 0 = cisto, 2 = ima nalaza.
"""
import argparse
import io
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "src-vba")

VBATOOLS = "modVbaTools.bas"
IMPORTSTATE = "modImportState.bas"

RE_RESET = re.compile(r'SaveSetting\s+\w+\s*,\s*\w+\s*,\s*"mutated"\s*,\s*""')
RE_SET1 = re.compile(r'SaveSetting\s+\w+\s*,\s*[\w()]+\s*,\s*"mutated"\s*,\s*"1"')
RE_EVENTS_OFF = re.compile(r'^\s*Application\.EnableEvents\s*=\s*False', re.M)
RE_EVENTS_ON = re.compile(r'^\s*Application\.EnableEvents\s*=\s*True', re.M)


def bez_komentara(tekst):
    """Skini VBA komentare -- pravilo se ne sme zadovoljiti recenicom o pravilu."""
    out = []
    for red in tekst.split("\n"):
        golo = red.split("'")[0] if "'" in red else red
        out.append(golo)
    return "\n".join(out)


def telo_procedure(tekst, ime):
    """Telo procedure `ime` (bez komentara i BEZ potpisa), ili None ako je nema.

    Potpis se preskace namerno: imena parametara (`ByVal mutated As String`)
    inace zadovoljavaju provere o sadrzaju tela. Self-test je tu gresku i nasao --
    slucaj "odluka nazad na jedan kljuc" je prolazio cist iako je telo citalo
    samo `pending`.
    """
    m = re.search(r"^\s*(?:Public |Private )?(?:Sub|Function)\s+%s\b" % re.escape(ime),
                  tekst, re.M)
    if not m:
        return None
    nl = tekst.find("\n", m.end())
    if nl < 0:
        return None
    # nastavak potpisa prelomljen sa " _" pripada potpisu, ne telu
    while tekst[:nl].rstrip().endswith("_"):
        nl = tekst.find("\n", nl + 1)
        if nl < 0:
            return None
    kraj = re.search(r"^\s*End (?:Sub|Function)\s*$", tekst[nl:], re.M)
    if not kraj:
        return None
    return tekst[nl:nl + kraj.start()]


def proveri(vbatools_txt, importstate_txt):
    nalazi = []
    vt = bez_komentara(vbatools_txt)
    ist = bez_komentara(importstate_txt)

    # --- RESET -------------------------------------------------------------
    begin = telo_procedure(vt, "BeginImportTransaction")
    if begin is None:
        nalazi.append("RESET: nema procedure BeginImportTransaction -- kapija ne meri nista")
    elif RE_RESET.search(begin):
        nalazi.append(
            "RESET: BeginImportTransaction resetuje \"mutated\" na \"\". Nov pokusaj "
            "importa nije dokaz da je raniji popravljen -- prolaz koji padne pre "
            "sopstvene mutacije bi zatecenu stetu proglasio bezbednom.")

    # --- NASLEDJIVANJE -----------------------------------------------------
    # Drugi oblik istog reseta, koji prva verzija ove kapije NIJE videla: ona je
    # gledala samo doslovno SaveSetting "mutated","". Ali FAIL grana radi
    # "If Not mMutated Then ClearImportPhase2State", a to brise CELU sekciju --
    # pa "mMutated = False" na pocetku prolaza ima IDENTICAN efekat: prolaz koji
    # padne pre sopstvene mutacije obrise dokaz ranije, jos nepopravljene stete.
    if re.search(r"^\s*mMutated\s*=\s*False\s*$", vt, re.M):
        nalazi.append(
            "NASLEDJIVANJE: mMutated se inicijalizuje na False. FAIL grana tada "
            "brise CELU sekciju i kad je raniji prolaz stvarno mutirao projekat -- "
            "isti efekat kao reset \"mutated\", samo napisan drugacije. Mora da "
            "nasledi zateceno nerazreseno stanje (ImportNijeDovrsen).")
    elif not re.search(r"mMutated\s*=\s*[\w.]*ImportNijeDovrsen", vt):
        nalazi.append(
            "NASLEDJIVANJE: mMutated se ne inicijalizuje iz zatecenog stanja "
            "(ImportNijeDovrsen) -- FAIL grana ne moze da razlikuje benigni marker "
            "od dokaza ranije mutacije.")

    # --- MUTACIJA ----------------------------------------------------------
    n_set = len(RE_SET1.findall(vt))
    if n_set < 2:
        nalazi.append(
            "MUTACIJA: \"mutated\"=\"1\" se upisuje %d put(a), ocekivano najmanje 2 "
            "(faza 1 pre prve destruktivne operacije, i ulazak u fazu 2)." % n_set)

    # --- EVENTI ------------------------------------------------------------
    m = RE_SET1.search(vt)
    if m:
        posle = vt[m.end():m.end() + 1200]
        if not RE_EVENTS_ON.search(posle):
            nalazi.append(
                "EVENTI: posle prve mutacije nema Application.EnableEvents = True. "
                "Workbook_BeforeSave je jedina globalna kapija nad Save-om; sa "
                "ugasenim eventima Ctrl+S prolazi neometano.")

    faza2 = telo_procedure(vt, "ImportAllVBA_Phase2")
    if faza2 is None:
        nalazi.append("EVENTI: nema procedure ImportAllVBA_Phase2 -- kapija ne meri fazu 2")
    elif RE_EVENTS_OFF.search(faza2):
        nalazi.append(
            "EVENTI: faza 2 gasi Application.EnableEvents. Ona radi nad projektom "
            "kome je faza 1 vec uklonila komponente -- bas tada Save mora da moze "
            "da se zaustavi.")

    # --- ODLUKA ------------------------------------------------------------
    odluka = telo_procedure(ist, "MarkerBlokira")
    if odluka is None:
        nalazi.append("ODLUKA: nema modImportState.MarkerBlokira")
    else:
        if "mutated" not in odluka:
            nalazi.append(
                "ODLUKA: MarkerBlokira ne gleda \"mutated\". Samo \"pending\" znaci "
                "\"import je jednom pokrenut\", ne \"projekat je pokvaren\" -- kapija "
                "bi blokirala i sveske koje nikad nisu dirnute.")
        if "pending" not in odluka:
            nalazi.append("ODLUKA: MarkerBlokira ne gleda \"pending\"")

    return nalazi


# --- self-test: kapija koja nikad nije pokazana crvena ne dokazuje da meri ----

CIST_VBATOOLS = '''
Private Sub ImportAllVBA()
    mMutated = modImportState.ImportNijeDovrsen()
    If Not mMutated Then ClearImportPhase2State
    mMutated = True
    SaveSetting IMPORT_REG_APP, P2Section(), "mutated", "1"
    Application.EnableEvents = True
End Sub

Private Sub ImportAllVBA_Phase2()
    SaveSetting IMPORT_REG_APP, sec, "mutated", "1"
    Application.ScreenUpdating = False
End Sub

Private Function BeginImportTransaction() As Boolean
    SaveSetting IMPORT_REG_APP, sec, "pending", "1"
End Function
'''

CIST_STATE = '''
Public Function MarkerBlokira(ByVal pending As String, ByVal mutated As String) As Boolean
    MarkerBlokira = (pending = "1") And (mutated = "1")
End Function
'''

SLUCAJEVI = [
    ("cist", CIST_VBATOOLS, CIST_STATE, 0),
    ("reset u BeginImportTransaction",
     CIST_VBATOOLS.replace('SaveSetting IMPORT_REG_APP, sec, "pending", "1"',
                           'SaveSetting IMPORT_REG_APP, sec, "mutated", ""'),
     CIST_STATE, 1),
    ("samo jedna tacka mutacije",
     CIST_VBATOOLS.replace('    SaveSetting IMPORT_REG_APP, sec, "mutated", "1"\n', ""),
     CIST_STATE, 1),
    ("eventi ostaju ugaseni posle mutacije",
     CIST_VBATOOLS.replace("    Application.EnableEvents = True\n", ""),
     CIST_STATE, 1),
    ("faza 2 gasi evente",
     CIST_VBATOOLS.replace("Private Sub ImportAllVBA_Phase2()\n",
                           "Private Sub ImportAllVBA_Phase2()\n    Application.EnableEvents = False\n"),
     CIST_STATE, 1),
    ("mMutated krece od False -- FAIL brise tudji dokaz",
     CIST_VBATOOLS.replace("mMutated = modImportState.ImportNijeDovrsen()",
                           "mMutated = False"),
     CIST_STATE, 1),
    ("mMutated ne nasledjuje zateceno stanje",
     CIST_VBATOOLS.replace("mMutated = modImportState.ImportNijeDovrsen()",
                           "mMutated = (1 = 2)"),
     CIST_STATE, 1),
    ("odluka nazad na jedan kljuc",
     CIST_VBATOOLS,
     CIST_STATE.replace('(pending = "1") And (mutated = "1")', '(pending = "1")'), 1),
    ("komentar ne zadovoljava pravilo",
     CIST_VBATOOLS.replace("    Application.EnableEvents = True\n",
                           "    ' Application.EnableEvents = True\n"),
     CIST_STATE, 1),
]


def self_test():
    pali = []
    for ime, vt, ist, ocekivano in SLUCAJEVI:
        n = len(proveri(vt, ist))
        if ocekivano == 0 and n != 0:
            pali.append("'%s' je trebalo da bude cist, a dao je %d nalaza" % (ime, n))
        if ocekivano > 0 and n == 0:
            pali.append("'%s' je trebalo da zapisti, a prosao je cist" % ime)
    if pali:
        print("vba_import_marker_gate --self-test: PALO", file=sys.stderr)
        for x in pali:
            print("  " + x, file=sys.stderr)
        return 2
    print("vba_import_marker_gate --self-test: %d slucajeva, cisto" % len(SLUCAJEVI))
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--self-test", action="store_true")
    args = ap.parse_args(argv)

    if args.self_test:
        return self_test()

    putevi = [os.path.join(SRC, VBATOOLS), os.path.join(SRC, IMPORTSTATE)]
    for p in putevi:
        if not os.path.exists(p):
            print("nema fajla: " + p, file=sys.stderr)
            return 2

    tekstovi = [io.open(p, encoding="ascii", errors="replace", newline="").read()
                for p in putevi]
    nalazi = proveri(*tekstovi)

    if nalazi:
        print("vba_import_marker_gate: %d nalaza" % len(nalazi), file=sys.stderr)
        for x in nalazi:
            print("  " + x, file=sys.stderr)
        return 2

    print("vba_import_marker_gate: cisto (marker je sticky, eventi zive kroz "
          "destruktivni deo, odluka trazi oba kljuca).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
