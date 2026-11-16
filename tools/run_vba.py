"""Headless VBA runner za AgriX OtkupApp: import src-vba -> test suite (compile proba samo uz --compile-only).

Radi SAMO na Windows masini sa instaliranim Excelom (COM). Na macOS/Linux nema
VBProject COM interfejsa -- VBA sesije se vode na Windows kutiji.

    pip install pywin32
    File > Options > Trust Center > Trust Center Settings > Macro Settings
        -> "Trust access to the VBA project object model"

Upotreba:
    python tools/run_vba.py --compile-only          # samo import + compile (najbrze, najstabilnije)
    python tools/run_vba.py                         # + podrazumevani set suite-ova
    python tools/run_vba.py --suite RunIzvestajTests
    python tools/run_vba.py --suite RunAllTests --test T_ParseDatum_Ugovor   # samo ti testovi
    python tools/run_vba.py --all

Izbor testova (`--test`, moze vise puta) pusta samo imenovane testove JEDNE
suite, u njenom redosledu, kroz ulaznu tacku iz kljuca `izbor` u SUITES. Takav
prolaz NIJE dokaz cele suite: marker zelenog se ne upisuje (vba_gate). Uz svaki
prolaz se ispisuju najsporiji testovi (VREME), iz `<rezultat>_vreme.txt`.
    python tools/run_vba.py --workbook "C:\\putanja\\AgriX_OtkupApp.xlsm"

Sveska: bez `--workbook` koristi se `tests/fixtures/otkup_test.xlsm`. Ako ga nema,
skript napravi PRAZNU .xlsm -- dovoljno za compile, ali NE i za suite (prazna
sveska nema tabele). Pravi fixture se pravi sa `tools/make_fixture.py`; sadrzi
samo sinteticke podatke, a suite koje diraju tabele ionako seju sebi svoje
(SVT-*, BIT-*, TST-*) u transakciji koja se uvek ponistava -- prava radna sveska
im NIJE potrebna. Sveska se uvek kopira u temp, original se ne dira.
Detalji: docs/EXCEL_TEST_HARNESS.md.

Izlazni kod: 0 = zeleno, 2 = palo (pala suite, Excel ubijen -- FATAL, ili uz
--compile-only compile greska / NEJASNO).

Zasto ovaj skript NE zove `ImportAllVBA`:
  - `modVbaTools.ImportAllVBA` ima hardkodiran folder
    (C:\\Users\\Dusan\\Documents\\GitHub\\otkupapp-pwa\\src-vba\\) i zavrsava se
    `MsgBox`-om. Oboje je smrt za headless run. Import se ovde radi istom logikom
    preko COM-a, sa folderom izvedenim iz lokacije ovog fajla.
  - Poredak i pravila su ista kao u `ImportAllVBA`: .doccls se MERGE-uje u
    postojecu komponentu, .bas/.cls/.frm se uklone pa uvezu, .frm mora imati .frx
    par, `modVbaTools` se preskace (izvrsava se).
"""

from __future__ import annotations

import argparse
import importlib.util
import io
import json
import os
import re
import shutil
import sys
import tempfile
import threading
import time

# --- Suite katalog -----------------------------------------------------------
#
# `gate` = suite PODIZE gresku kada neka provera padne. Samo takve suite headless
# runner moze da vidi kao crvene -- kod ostalih rezultat postoji jedino u
# Immediate prozoru, pa ih runner prijavljuje kao "blind" (prosle bez greske,
# sto NIJE isto sto i "sve provere prosle").
#
# `dialogs` = suite otvara MsgBox (potvrda i/ili zavrsni rezime) -> oslanja se na
# watchdog koji dijalog zatvara podrazumevanim dugmetom.

SUITES = {
    # Verdikt NE dolazi iz toga da li Run() pukne (modTest hvata gresku po testu
    # da jedan pad ne obori ostale), nego iz last_run.txt pored sveske.
    #
    # `izbor` = ulazna tacka koja prima spisak testova ("T_A,T_B") i pusta samo
    # njih, u redosledu suite-a, sa istim rezultat-fajlom. Samo suite koje je
    # imaju primaju `--test`.
    "RunAllTests":              {"gate": True,  "dialogs": False, "default": True,
                                 "result_file": "last_run.txt",
                                 "izbor": "RunAllTestsIzbor"},
    "RunIzvestajTests":         {"gate": True,  "dialogs": False, "default": True},
    "RunSheetsJsonParserTests": {"gate": True,  "dialogs": True,  "default": True},
    # Detalji pada ove suite ne prezive COM granicu (opis greske stigne kao golo
    # "Exception occurred"), pa i ona pise rezultat pored sveske.
    "RunBankaImportTestSuite":  {"gate": True,  "dialogs": True,  "default": True,
                                 "result_file": "last_run_banka.txt"},
    "RunFakturaSmokeSuite":     {"gate": True,  "dialogs": True,  "default": True},
    "Test_StornoCentar_All":    {"gate": True,  "dialogs": False, "default": True},
    "TestLicense_All":          {"gate": True,  "dialogs": False, "default": True},
    # Nisu u podrazumevanom setu: traze mrezu, live SEF nalog ili duze rade.
    "RunGoogleSyncSmokeSuite":  {"gate": True,  "dialogs": True,  "default": False},
    "RunMasterSyncSmokeSuite":  {"gate": True,  "dialogs": True,  "default": False},
    "RunSEFTestSuite":          {"gate": True,  "dialogs": True,  "default": False},
    "RunStornoTestSuite":       {"gate": True,  "dialogs": True,  "default": True,
                                 "result_file": "last_run_storno.txt"},
    "RunPaleteTestSuite":       {"gate": True,  "dialogs": True,  "default": True},
    "RunNovacSmokeSuite":       {"gate": False, "dialogs": True,  "default": False},
    # Pad ove suite stize kao Err.Raise iz EndRun, pa opis ne prezivi COM
    # granicu -- ime pale tvrdnje se cita iz fajla, kao kod banke.
    "RunBusinessFlowProSuite":  {"gate": True,  "dialogs": True,  "default": True,
                                 "result_file": "last_run_bfp.txt",
                                 "izbor": "RunBusinessFlowProIzbor"},
    "RunAgrohemijaSmokeSuite":  {"gate": True,  "dialogs": True,  "default": True},
    # Golden poslovni scenariji: mere ISHOD, ne strukturu. Svaki radi u
    # transakciji i rollback-uje se, pa je drugi prolaz identican prvom.
    # Verdikt stize kao Err.Raise iz RunGoldenSuite -- gate: True.
    "RunGoldenSuite":           {"gate": True,  "dialogs": False, "default": True},
    "RunProductionHealthCheck": {"gate": False, "dialogs": True,  "default": False},
    "TestMonitoring_All":       {"gate": False, "dialogs": False, "default": False},
}

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_VBA = os.path.join(ROOT, "src-vba")
SELF_MODULE = "modVbaTools"          # isto kao modVbaTools.SELF_MODULE
COMPILE_CONTROL_ID = 578             # VBE: Debug > Compile VBAProject

VBE_TYPE = {".bas": 1, ".cls": 2, ".frm": 3, ".doccls": 100}

DEFAULT_FIXTURE = os.path.join(ROOT, "tests", "fixtures", "otkup_test.xlsm")
GOLDEN_DIR = os.path.join(ROOT, "tests", "golden")
XL_OPENXML_MACRO = 52                # xlOpenXMLWorkbookMacroEnabled (.xlsm)


def _copy_golden(src: str, dst: str) -> None:
    os.makedirs(dst, exist_ok=True)
    if not os.path.isdir(src):
        return
    for name in os.listdir(src):
        if name.lower().endswith(".txt"):
            shutil.copy2(os.path.join(src, name), os.path.join(dst, name))


def _read_test_results(wbdir: str, report: dict,
                       fname: str = "last_run.txt",
                       suite: str = "RunAllTests") -> int:
    """Verdikt iz fajla koji je suite upisala pored sveske.

    Nema fajla = pad, ne prolaz: to znaci da suite nije stigla do kraja
    (compile error, visenje, ubijen proces) -- ishod koji nije eksplicitno OK.

    Rezultat ide u SLOT PO SUITE-U. Dok je postojao jedan globalni
    report["tests"], druga suite sa rezultat-fajlom prepisivala je prvu: full run
    je umeo da zavrsi sa "SUITE FAIL RunAllTests" i "TESTS 196 ukupno, 0 palo",
    a ime palog testa je nestajalo. Izlaz koji sakriva crveno je tacno ono protiv
    cega je ceo ovaj alat.
    """
    slotovi = report.setdefault("suite_results", {})
    path = os.path.join(wbdir, fname)
    if not os.path.exists(path):
        slotovi[suite] = {"error": f"suite nije upisala {fname} "
                                   "(nije stigla do kraja)"}
        return 2

    with open(path, "r", encoding="utf-8", errors="replace") as fh:
        text = fh.read()

    lines = [ln.rstrip("\r") for ln in text.split("\n")]
    head = lines[0] if lines else ""
    total = failed = -1
    for token in head.split():
        if token.startswith("TESTS="):
            total = int(token[6:] or -1)
        elif token.startswith("FAIL="):
            failed = int(token[5:] or -1)

    detail = [ln for ln in lines[1:] if ln.strip()]
    slotovi[suite] = {"total": total, "failed": failed, "detail": detail}

    if total < 0 or failed < 0:
        slotovi[suite]["error"] = f"neocekivan prvi red: {head!r}"
        return 2
    return 2 if failed else 0


# --- Watchdog za modalne dijaloge --------------------------------------------
#
# Svaki MsgBox u nevidljivom Excelu je trajno visenje: COM poziv ne puca, samo
# stoji. Watchdog nadgleda prozore klase #32770 (standardni Windows dialog) koji
# pripadaju NASOJ Excel instanci, procita tekst i klikne prvo dugme (za MsgBox
# je to podrazumevano: OK kod vbOKOnly, Yes kod vbYesNo). Time se visenje
# pretvara u zabelezen dogadjaj -- ukljucujuci i "Compile error: ..." dijalog,
# koji je inace jedini nacin da VBE prijavi neuspeh kompajliranja.

class DialogWatchdog(threading.Thread):
    def __init__(self, pid: int, poll: float = 0.4, na_dijalog=None):
        super().__init__(daemon=True)
        self.pid = pid
        self.poll = poll
        self.seen: list[str] = []
        self._stop = threading.Event()
        # Javlja se tekst SVAKOG dijaloga (v. Prekid.dijalog): koban dijalog
        # skracuje cekanje na Excel, ostali ne menjaju nista.
        self.na_dijalog = na_dijalog

    def stop(self) -> None:
        self._stop.set()

    def run(self) -> None:
        import win32gui
        import win32process

        while not self._stop.is_set():
            try:
                self._sweep(win32gui, win32process)
            except Exception:
                pass
            self._stop.wait(self.poll)

    def _sweep(self, win32gui, win32process) -> None:
        targets: list[int] = []

        def on_window(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return
            if win32gui.GetClassName(hwnd) != "#32770":
                return
            if win32process.GetWindowThreadProcessId(hwnd)[1] != self.pid:
                return
            targets.append(hwnd)

        win32gui.EnumWindows(on_window, None)

        for hwnd in targets:
            text = self._dialog_text(win32gui, hwnd)
            self.seen.append(text)
            if self.na_dijalog is not None:
                try:
                    self.na_dijalog(text)
                except Exception:
                    pass
            self._click_default(win32gui, hwnd)

    @staticmethod
    def _dialog_text(win32gui, hwnd) -> str:
        parts = [win32gui.GetWindowText(hwnd)]

        def on_child(child, _):
            if win32gui.GetClassName(child) == "Static":
                s = win32gui.GetWindowText(child).strip()
                if s:
                    parts.append(s)

        try:
            win32gui.EnumChildWindows(hwnd, on_child, None)
        except Exception:
            pass
        return " | ".join(p for p in parts if p)

    @staticmethod
    def _click_default(win32gui, hwnd) -> None:
        buttons: list[tuple[int, str]] = []

        def on_child(child, _):
            if win32gui.GetClassName(child) == "Button":
                caption = win32gui.GetWindowText(child).replace("&", "").strip().lower()
                buttons.append((child, caption))

        try:
            win32gui.EnumChildWindows(hwnd, on_child, None)
        except Exception:
            pass

        BM_CLICK = 0x00F5
        WM_CLOSE = 0x0010
        if not buttons:
            win32gui.PostMessage(hwnd, WM_CLOSE, 0, 0)
            return

        # NIKAD "Debug": na VBA runtime-error dijalogu to ubacuje VBE u break
        # mode, posle cega COM pozivi vise ne odgovaraju i run visi zauvek.
        # (Stari kod je klikao buttons[0] naslepo -- a "Debug" ume da bude prvi.)
        forbidden = ("debug", "help", "pomoc")
        preferred = ("end", "ok", "u redu", "da", "yes", "cancel", "otkazi", "zavrsi")

        safe = [(h, c) for h, c in buttons if not any(f in c for f in forbidden)]
        if not safe:
            win32gui.PostMessage(hwnd, WM_CLOSE, 0, 0)
            return

        for want in preferred:
            for h, c in safe:
                if c == want or c.startswith(want):
                    win32gui.PostMessage(h, BM_CLICK, 0, 0)
                    return

        win32gui.PostMessage(safe[0][0], BM_CLICK, 0, 0)


# --- Import src-vba preko COM-a ----------------------------------------------

def _read_code_body(path: str) -> str:
    """Vrati samo kod, bez VBA header bloka -- za .doccls merge.

    Header izgleda ovako:

        VERSION 1.0 CLASS
        BEGIN
          MultiUse = -1  'True
        END
        Attribute VB_Name = "ThisWorkbook"
        Attribute VB_GlobalNameSpace = False
        ...
        <ovde pocinje kod>

    Header se mora skinuti POZICIJSKI, red po red. Raniji filter "preskoci sve
    sto lici na header" je puknuo na `MultiUse = -1  'True`: ta linija ne lici ni
    na sta iz liste, pa je proglasena pocetkom koda -- i `END` + sve `Attribute`
    linije su zavrsile U MODULU. `End` je u VBA naredba koja obara izvrsavanje,
    pa je uvezena sveska padala u break mode i run je visio zauvek.
    """
    with open(path, "r", encoding="ascii", errors="replace") as fh:
        lines = fh.read().splitlines()

    i, n = 0, len(lines)

    if i < n and lines[i].strip().upper().startswith("VERSION"):
        i += 1

    if i < n and lines[i].strip().upper().startswith("BEGIN"):
        depth = 0
        while i < n:
            token = lines[i].strip().upper()
            if token.startswith("BEGIN"):
                depth += 1
            elif token == "END":
                depth -= 1
            i += 1
            if depth == 0:
                break

    while i < n and lines[i].strip().startswith("Attribute "):
        i += 1

    return "\r\n".join(lines[i:])


def _has_vb_header(path: str) -> bool:
    with open(path, "r", encoding="ascii", errors="replace") as fh:
        first = fh.readline()
    return "Attribute VB_Name" in first or first.lstrip().upper().startswith("VERSION")


def report_orphan_components(wb, log: list[str], src: str = SRC_VBA) -> None:
    """Moduli koji postoje u svesci a NEMA ih u src-vba/ (ili u `--src`).

    Import prepisuje samo ono sto repo ima; zaostao modul iz donora ostaje i
    izvrsava se. Ako nosi Public ime koje postoji i u svezem kodu, VBA to vidi kao
    "Ambiguous name" i odbija da pokrene makro iz njega -- sto izgleda kao
    "Cannot run the macro", a ne kao compile greska. Tako je TestLicense_All bio
    mrtav a `vba_check` uredno zelen: duplikat nije bio u repou nego u svesci.

    Samo prijavljuje. Brisanje bi bilo agresivno prema svesci koju je operater
    prosledio kroz --workbook.
    """
    STD, CLS, FRM = 1, 2, 3          # vbext_ct_StdModule / ClassModule / MSForm
    have = {os.path.splitext(n)[0].lower()
            for n in os.listdir(src)
            if os.path.splitext(n)[1] in (".bas", ".cls", ".frm")}

    orphans = []
    for vbc in wb.VBProject.VBComponents:
        try:
            if int(vbc.Type) not in (STD, CLS, FRM):
                continue                      # document moduli se merge-uju, ne brisu
            name = str(vbc.Name)
        except Exception:
            continue
        if name.lower() == SELF_MODULE.lower():
            continue
        if name.lower() not in have:
            orphans.append(name)

    if orphans:
        log.append(f"ORPHAN {len(orphans)} modul(a) u svesci bez para u src-vba/ "
                   f"(mogu da daju 'Ambiguous name'): " + ", ".join(sorted(orphans)))


def import_src_vba(wb, log: list[str], src: str = SRC_VBA) -> None:
    """Uvezi izvor iz `src` (podrazumevano src-vba/).

    `src` drugaciji od src-vba daje dokaz.py (D-6): mutacija se primenjuje u
    njegovoj kopiji izvora, pa radno stablo tokom dokaza niko ne pise.
    """
    proj = wb.VBProject
    files = sorted(os.listdir(src))

    # 1) document moduli (.doccls) -- kod se MERGE-uje u postojecu komponentu
    #
    # Nedostajuce komponente se sabiraju u JEDNU liniju: nad praznim fixture-om
    # nedostaje svih 40+ listova, a 40 linija SKIP-a zatrpa nalaz zbog kog se
    # ovo i pokrece.
    missing: list[str] = []
    for name in files:
        base, ext = os.path.splitext(name)
        if ext != ".doccls":
            continue
        try:
            vbc = proj.VBComponents(base)
        except Exception:
            missing.append(base)
            continue
        cm = vbc.CodeModule
        if cm.CountOfLines > 0:
            cm.DeleteLines(1, cm.CountOfLines)
        cm.AddFromString(_read_code_body(os.path.join(src, name)))

    if missing:
        log.append(f"SKIP {len(missing)} .doccls (nema komponente u svesci): "
                   + ", ".join(missing))

    # 2) standardni / klasni / forme
    for name in files:
        base, ext = os.path.splitext(name)
        if ext not in (".bas", ".cls", ".frm"):
            continue
        if base.lower() == SELF_MODULE.lower():
            continue
        path = os.path.join(src, name)

        if ext == ".frm" and not os.path.exists(os.path.join(src, base + ".frx")):
            log.append(f"SKIP {name} (nema .frx para)")
            continue

        try:
            proj.VBComponents.Remove(proj.VBComponents(base))
        except Exception:
            pass  # komponenta ne postoji -- normalno za nov modul

        if _has_vb_header(path):
            proj.VBComponents.Import(path)
        elif ext in (".bas", ".cls"):
            vbc = proj.VBComponents.Add(VBE_TYPE[ext])
            vbc.Name = base
            vbc.CodeModule.AddFromFile(path)
        else:
            log.append(f"SKIP {name} (forma bez headera)")


def self_test() -> int:
    """Provere koje NE traze Excel -- rade i na Linux/macOS.

    Postoje zato sto se strip VBA header-a pokvario u tisini: header je zavrsio
    u kodu, sveska je pala u break mode, a to se videlo tek na Windows masini,
    kroz screenshot. Ova provera hvata istu klasu greske bez Excela.
    """
    leaks: list[str] = []
    checked = 0
    for name in sorted(os.listdir(SRC_VBA)):
        if not name.endswith(".doccls"):
            continue
        checked += 1
        body = _read_code_body(os.path.join(SRC_VBA, name))
        for line in body.splitlines()[:5]:
            s = line.strip()
            if (s.startswith("Attribute ") or s.startswith("MultiUse")
                    or s.upper() in ("BEGIN", "END") or s.upper().startswith("VERSION ")):
                leaks.append(f"{name}: header procureo u kod -> {s!r}")
                break

    # MARKER JE WIRE-UP, NE SAMO FUNKCIJA. zabelezi_prolaz je dokazan u
    # `vba_gate.py --self-test`; ovde se tvrdi da ga OVAJ alat zove, i da
    # mu prenosi `--no-import` (bez toga bi marker tvrdio da je izvrsen
    # src-vba, a izvrsen je kod zatecen u svesci). Pozivno mesto je u main()
    # posle Excela, pa se bez Excela ne moze IZVRSITI -- tvrdi se nad izvorom.
    with io.open(os.path.abspath(__file__), encoding="utf-8",
                 newline="") as fh:
        sopstveni = fh.read()
    # IGLA NE SME DA POGODI SEBE. Tekstovi koji se traze stoje i ovde, kao
    # string literali, pa je pretraga nad CELIM fajlom zadovoljena sopstvenim
    # navodnikom: obrisi pravi poziv i provera ostane zelena. Dvosmerni dokaz
    # je to i pokazao, pa se telo ovog self-testa izbacuje iz pretrage.
    sopstveni = sopstveni.replace("\r\n", "\n")
    poc = sopstveni.find("def self_test(")
    kraj = sopstveni.find("\ndef ", poc + 1)
    tudje = sopstveni[:poc] + sopstveni[kraj if kraj > 0 else len(sopstveni):]
    for tekst, opis in (
            ('_gate.snimi_kontekst(src_dir=src, sveska=fixture,\n'
             '                                               sveska_hes_iz=wbpath)',
             "main() ne snima kontekst dokaza PRE run-a, nad temp kopijom "
             "sveske i nad izvorom koji se stvarno uvozi (--src)"),
            ('_gate.zabelezi_prolaz(\n'
             '                report, rc, args.no_import, kontekst=kontekst_dokaza, src_dir=src)',
             "main() ne predaje SNIMLJEN kontekst, ili marker racuna otisak nad "
             "src-vba umesto nad uvezenim izvorom (--src)"),
            ('import_src_vba(wb, report["import"], src)\n'
             '            report_orphan_components(wb, report["import"], src)',
             "main() uvozi src-vba umesto izvora iz --src (D-6: dokaz bi merio "
             "radno stablo, ne svoju kopiju)"),
            ('    src, razlog = izvor_iz_argumenata(args)',
             "main() ne proverava --src pre Excela"),
            ('lines.append(f"GREEN   {report[\'green\']}")',
             "izvestaj ne ispisuje red GREEN -- upis markera bi bio nevidljiv"),
    ):
        if tekst not in tudje:
            leaks.append("WIRE-UP: " + opis)

    leaks.extend(_self_test_izbor(tudje))
    leaks.extend(_self_test_prekid(tudje))
    leaks.extend(_self_test_src())

    for line in leaks:
        print(line, file=sys.stderr)
    if leaks:
        print(f"\nself-test: {len(leaks)} nalaza od {checked} .doccls fajlova.", file=sys.stderr)
        return 2
    print(f"self-test: cisto ({checked} .doccls fajlova + wire-up markera).")
    return 0


def _self_test_prekid(tudje: str) -> list:
    """Prekid po fazi, koban dijalog, compile proba, istorija, --bez-markera.

    Sve bez Excela: Prekid dobija lazne tajmere i laznu funkciju ubijanja, pa se
    redosled (nova faza -> pun budzet; koban dijalog -> KOBAN_SEKUNDI, jednom)
    meri bez cekanja. Da ga main() zaista koristi tvrdi se nad izvorom.
    """
    nalazi = []

    def tvrdi(uslov, opis):
        if not uslov:
            nalazi.append("PREKID: " + opis)

    # Igla ne sme da pogodi sebe (v. self_test).
    poc = tudje.find("def _self_test_prekid(")
    kraj = tudje.find("\ndef ", poc + 1)
    tudje = tudje[:poc] + tudje[kraj if kraj > 0 else len(tudje):]

    # --- koban dijalog, u oba smera -------------------------------------------
    for tekst in ("Microsoft Visual Basic for Applications | Compile error: | Syntax error",
                  "Microsoft Visual Basic | Run-time error '9': | Subscript out of range",
                  "Microsoft Visual Basic for Applications | Compile error: | "
                  "Variable not defined"):
        tvrdi(koban_dijalog(tekst), "koban dijalog nije prepoznat: %r" % tekst)
    for tekst in ("AgriX | Business Flow Pro tests finished. | RunID=1 | Failed=0",
                  "Microsoft Excel | Want to save your changes to 'otkup_test.xlsm'?",
                  "AgriX | Pokrenuti test suite? Greske u podacima se vracaju.", ""):
        tvrdi(not koban_dijalog(tekst), "obican dijalog proglasen kobnim: %r" % tekst)

    # --- Prekid, sa laznim tajmerima ------------------------------------------
    svi = []

    class LazniTajmer:
        def __init__(self, sekundi, fn):
            self.sekundi, self.fn, self.otkazan, self.daemon = sekundi, fn, False, False
            svi.append(self)

        def start(self):
            pass

        def cancel(self):
            self.otkazan = True

    ubijeni = []

    def lazno_ubij(pid, flag):
        flag["fired"] = True
        ubijeni.append(pid)

    KOBAN = "Microsoft Visual Basic for Applications | Compile error: | Syntax error"
    pr = Prekid(77, 1800.0, ubij=lazno_ubij, tajmer=LazniTajmer)
    pr.naoruzaj("priprema")
    pr.naoruzaj("RunAllTests")
    tvrdi(len(svi) == 2 and svi[0].otkazan and not svi[1].otkazan
          and svi[1].sekundi == 1800.0,
          "nova faza ne otkazuje prethodni tajmer ili ne dobija pun budzet")
    pr.dijalog("AgriX | Business Flow Pro tests finished.")
    tvrdi(len(svi) == 2, "obican dijalog menja prekid")
    pr.dijalog(KOBAN)
    tvrdi(len(svi) == 3 and svi[1].otkazan and svi[2].sekundi == KOBAN_SEKUNDI,
          "koban dijalog ne skracuje cekanje na KOBAN_SEKUNDI")
    pr.dijalog(KOBAN)
    tvrdi(len(svi) == 3,
          "isti koban dijalog u sledecem obilasku pomera prekid (nikad ne bi okinuo)")
    if len(svi) >= 3:
        svi[2].fn()
    tvrdi(ubijeni == [77] and pr.stanje["fired"] and pr.stanje["faza"] == "RunAllTests"
          and "Syntax error" in pr.stanje["dijalog"],
          "okinut prekid ne ubija proces ili ne belezi fazu i dijalog")
    tvrdi("RunAllTests" in poruka_prekida(pr.stanje)
          and "Syntax error" in poruka_prekida(pr.stanje),
          "FATAL posle kobnog dijaloga ne imenuje fazu i dijalog")
    # Izmereno 16.11.2026: VBA tekst dijaloga nosi prelome ("Compile error:\n\n
    # Syntax error"), a dokaz.py cita samo red FATAL -- razlog mora biti u njemu.
    lomljen = dict(pr.stanje, dijalog="Microsoft Visual Basic for Applications | "
                                      "Compile error:\r\n\r\nSyntax error")
    tvrdi("\n" not in poruka_prekida(lomljen)
          and "Compile error: Syntax error" in poruka_prekida(lomljen),
          "FATAL sa lomljenim tekstom dijaloga nije u jednom redu")
    pr.naoruzaj("RunBusinessFlowProSuite")
    tvrdi(svi[-1].sekundi == 1800.0 and not svi[-1].otkazan,
          "posle kobnog dijaloga sledeca faza ne dobija pun budzet")
    svi[-1].fn()
    tvrdi(pr.stanje["faza"] == "RunAllTests" and "Syntax error" in pr.stanje["dijalog"],
          "drugo okidanje brise PRVI uzrok iz FATAL poruke")
    n = len(svi)
    pr.dijalog(KOBAN)
    tvrdi(len(svi) == n + 1 and svi[-1].sekundi == KOBAN_SEKUNDI,
          "u novoj fazi koban dijalog vise ne skracuje cekanje")
    pr.stop()
    tvrdi(svi[-1].otkazan, "stop ne otkazuje tajmer")
    # Zaustavljen prekid (kraj run-a) se ne naoruzava ni na koban dijalog. Meri
    # se nad SVEZIM prekidom: kod `pr` je oznaka kobnog dijaloga vec upaljena,
    # pa bi odbijanje dolazilo od nje, ne od zaustavljanja.
    pr2 = Prekid(78, 1800.0, ubij=lazno_ubij, tajmer=LazniTajmer)
    pr2.naoruzaj("zatvaranje")
    pr2.stop()
    n = len(svi)
    pr2.dijalog(KOBAN)
    tvrdi(len(svi) == n, "zaustavljen prekid se ponovo naoruzava na dijalog")
    istekao = {"fired": True, "faza": "RunBusinessFlowProSuite", "dijalog": "",
               "sekundi": 1800.0}
    tvrdi("RunBusinessFlowProSuite" in poruka_prekida(istekao)
          and "--timeout" in poruka_prekida(istekao),
          "FATAL za istekao budzet ne imenuje fazu")

    # --- istorija prolaza ----------------------------------------------------
    tmp = tempfile.mkdtemp(prefix="runvba_ist_")
    try:
        put = os.path.join(tmp, "istorija.jsonl")
        izv = {"faze": [{"ime": "uvoz src-vba", "sekundi": 6.5}],
               "suites": [{"name": "RunAllTests", "status": "OK", "seconds": 230.9,
                           "gate": True}]}
        for _ in range(2):
            dopisi_istoriju(izv, 0, ["--suite", "RunAllTests"], time.time() - 5, "abc",
                            put=put)
        with open(put, encoding="utf-8") as fh:
            redovi = fh.read().splitlines()
        poslednji = json.loads(redovi[-1]) if redovi else {}
        tvrdi(len(redovi) == 2 and poslednji.get("faze") == izv["faze"]
              and poslednji.get("izvor") == "abc"
              and poslednji.get("suite") == [{"name": "RunAllTests", "status": "OK",
                                              "seconds": 230.9}],
              "istorija ne dopisuje red sa fazama, suitama i izvorom")
        tvrdi(dopisi_istoriju(izv, 0, [], time.time(), "",
                              put=os.path.join(put, "pod_fajlom.jsonl")) != "",
              "neuspeo upis istorije se ne prijavljuje")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    tvrdi(parse_args(["--bez-markera"]).bez_markera, "nema argumenta --bez-markera")

    # --- wire-up u main() -----------------------------------------------------
    for tekst, opis in (
            ('        if args.compile_only:\n'
             '            report["compile"] = compile_project(xl, wb, watchdog)',
             "compile proba se ne vezuje za --compile-only"),
            ("na_dijalog=prekid.dijalog)", "watchdog ne javlja dijaloge prekidu"),
            ("                    self.na_dijalog(text)\n",
             "watchdog ne zove na_dijalog za vidjen dijalog"),
            ("                prekid.naoruzaj(suite)\n",
             "suite ne dobija svoj budzet (prekid bi merio ceo prolaz)"),
            ("        if args.bez_markera:\n", "--bez-markera ne sprecava upis markera"),
            ('report["fatal"] = poruka_prekida(prekid.stanje)',
             "FATAL ne dolazi iz Prekid.stanje"),
            ("    dopisi_istoriju(report, rc, argv, pocetak,", "main ne dopisuje istoriju"),
            ('        report["faze"] = faze\n', "izvestaj ne nosi faze"),
    ):
        tvrdi(tekst in tudje, opis)
    tvrdi(tudje.count("compile_project(") - tudje.count("def compile_project(") == 1,
          "compile_project se zove i van --compile-only")
    return nalazi


def _vba_telo(fajl: str, zaglavlje: str) -> list:
    """Logicki redovi tela procedure (bez komentara i praznih), ili [] ako je nema."""
    with io.open(os.path.join(SRC_VBA, fajl), encoding="ascii",
                 errors="replace", newline="") as fh:
        redovi = fh.read().replace("\r\n", "\n").split("\n")
    out, unutra = [], False
    for red in redovi:
        s = red.strip()
        if not unutra:
            unutra = s.startswith(zaglavlje)
            continue
        if s.startswith("End Sub"):
            break
        if s and not s.startswith("'"):
            out.append(s)
    return out


def _self_test_izbor(tudje: str) -> list:
    """Izbor testova (`--test`) i vreme po testu -- provere bez Excela.

    Najopasniji kvar izbora je TIH: izbor koji prezivi prolaz pretvara sledeci
    pun run u deo suite-a, zelen. VBA test (T_TestIzbor_TacnoImeIJednokratno)
    meri pravilo, ali ne moze da izmeri da ga ulazna tacka ZOVE -- radi usred
    tog istog prolaza. Zato se poziv tvrdi ovde, nad izvorom.
    """
    nalazi = []

    # Igla ne sme da pogodi sebe (v. self_test): tekstovi koji se traze stoje i
    # u OVOJ funkciji, kao literali, pa se njeno telo izbacuje iz pretrage.
    poc = tudje.find("def _self_test_izbor(")
    kraj = tudje.find("\ndef ", poc + 1)
    if poc >= 0:
        tudje = tudje[:poc] + (tudje[kraj:] if kraj > 0 else "")

    def tvrdi(uslov, opis):
        if not uslov:
            nalazi.append("IZBOR: " + opis)

    # --- kombinacije argumenata, svaka u oba smera ---------------------------
    def problem(*argv):
        # Pad je NALAZ po imenu, ne traceback: pukla provera bi u izlazu nosila
        # i tekst tvrdnje (iz ispisa izvornog reda), pa bi se cinilo da je pala
        # bas ona.
        try:
            return problem_izbora(parse_args(list(argv)))
        except Exception as exc:            # noqa: BLE001
            nalazi.append("IZBOR: problem_izbora puca za %s -- %r" % (argv, exc))
            return "PUKAO"

    tvrdi(problem() == "", "run bez --test ima problem izbora")
    tvrdi(problem("--suite", "RunAllTests", "--test", "T_A") == "",
          "valjan izbor (jedna suite sa `izbor`) se odbija")
    tvrdi("TACNO JEDNU" in problem("--test", "T_A"),
          "--test bez --suite se prihvata")
    tvrdi(problem("--suite", "RunAllTests", "--suite", "RunGoldenSuite",
                  "--test", "T_A") != "", "--test uz DVE suite se prihvata")
    tvrdi(problem("--all", "--suite", "RunAllTests", "--test", "T_A") != "",
          "--test uz --all se prihvata")
    tvrdi(problem("--suite", "RunIzvestajTests", "--test", "T_A") != "",
          "--test nad suite BEZ ulazne tacke izbora se prihvata")
    tvrdi(problem("--compile-only", "--suite", "RunAllTests", "--test", "T_A") != "",
          "--test uz --compile-only se prihvata")
    tvrdi(problem("--suite", "RunAllTests", "--test", " , ") != "",
          "--test bez ijednog imena se prihvata")
    tvrdi(izabrani_testovi(parse_args(["--test", "T_A, T_B", "--test", "T_A"]))
          == ["T_A", "T_B"], "spisak testova se ne cisti (razmaci, duplikati)")
    for ime, meta in SUITES.items():
        if meta.get("izbor") and not meta.get("result_file"):
            tvrdi(False, "suite %s ima `izbor` bez result_file -- izbor ne bi "
                         "imao gde da procita koji je test pao" % ime)

    # --- vreme po testu --------------------------------------------------------
    tvrdi(vreme_fajl("last_run_bfp.txt") == "last_run_bfp_vreme.txt",
          "ime fajla vremena ne prati rezultat-fajl")
    tmpd = tempfile.mkdtemp(prefix="vreme_")
    try:
        put = os.path.join(tmpd, "v.txt")
        with open(put, "w", encoding="utf-8", newline="") as fh:
            # Str$ u VBA pise ".123" (bez nule), red na kraju nema TAB-ove.
            fh.write(".123\tOK\tT_A\n12.5\tFAIL\tT_B\nnije broj\tOK\tT_C\nsmece\n")
        tvrdi(procitaj_vreme(put) == [(12.5, "FAIL", "T_B"), (0.123, "OK", "T_A")],
              "vreme se ne cita kako ga VBA pise (ili nije sortirano)")
        tvrdi(procitaj_vreme(os.path.join(tmpd, "nema.txt")) == [],
              "nepostojeci fajl vremena obara run")
    finally:
        shutil.rmtree(tmpd, ignore_errors=True)

    # --- wire-up u ovom alatu ---------------------------------------------------
    for tekst, opis in (
            ("razlog = problem_izbora(args)", "main() ne proverava kombinaciju --test"),
            ("xl.Run(f\"'{wb.Name}'!{meta['izbor']}\", \",\".join(testovi))",
             "main() ne zove ulaznu tacku izbora sa spiskom testova"),
            ('entry["izbor"] = testovi',
             "izvestaj ne oznacava izbor -- vba_gate bi upisao marker cele suite"),
    ):
        tvrdi(tekst in tudje, opis)

    # --- wire-up u VBA izvoru ---------------------------------------------------
    run_all = _vba_telo("modTest.bas", "Public Sub RunAllTests()")
    prvi_run = next((i for i, r in enumerate(run_all) if r.startswith("RunOne ")), -1)
    preuzmi = next((i for i, r in enumerate(run_all)
                    if r == "m_Izbor = PreuzmiIzbor(m_IzborZahtev)"), -1)
    tvrdi(0 <= preuzmi < prvi_run,
          "RunAllTests ne preuzima (i ne brise) izbor PRE prvog RunOne")
    tvrdi("If Not TestIzabran(nm) Then Exit Sub"
          in _vba_telo("modTest.bas", "Private Sub RunOne("),
          "RunOne ne pita TestIzabran -- izbor bi pustio sve testove")

    bfp = _vba_telo("modBusinessFlowProTests.bas",
                    "Public Sub RunBusinessFlowProSuite()")
    tvrdi(bfp[:1] == ["m_BfpIzbor = BfpPreuzmiIzbor(m_BfpIzborZahtev)"],
          "RunBusinessFlowProSuite ne preuzima izbor kao PRVU naredbu")
    # Telo suite je podeljeno na BfpDeo1..N (PROCEDURA_VELIKA). Deo koji suite
    # ne pozove TIHO izbaci stotinu testova -- VBA nepozvan Private Sub ne
    # prijavljuje, a suite ostaje zelena. Zato: svaki definisan deo je pozvan
    # tacno jednom, redom, a deo nosi samo pozive kroz BfpIzabran.
    with io.open(os.path.join(SRC_VBA, "modBusinessFlowProTests.bas"),
                 encoding="ascii", errors="replace") as fh:
        bfp_izvor = fh.read()
    definisani = sorted(int(n) for n in
                        re.findall(r"(?im)^Private Sub BfpDeo(\d+)\(\)", bfp_izvor))
    pozvani = [int(r[len("BfpDeo"):]) for r in bfp if re.fullmatch(r"BfpDeo\d+", r)]
    tvrdi(pozvani == definisani == list(range(1, len(definisani) + 1)),
          "RunBusinessFlowProSuite ne zove svaki BfpDeoN tacno jednom i redom "
          "(definisani %s, pozvani %s) -- nepozvan deo tiho izbaci svoje testove"
          % (definisani, pozvani))
    telo_delova = []
    for n in definisani:
        telo_delova += _vba_telo("modBusinessFlowProTests.bas",
                                 "Private Sub BfpDeo%d()" % n)
    tudje_u_delu = [r for r in telo_delova
                    if not re.fullmatch(r'If BfpIzabran\("(Test_\w+)"\) Then \1', r)]
    tvrdi(not tudje_u_delu,
          "BfpDeoN nosi nesto osim poziva `If BfpIzabran(\"Test_X\") Then "
          "Test_X`: %s" % ", ".join(tudje_u_delu[:3]))
    gole = [r for r in bfp + telo_delova if re.fullmatch(r"Test_\w+", r)]
    tvrdi(not gole,
          "RunBusinessFlowProSuite zove test mimo BfpIzabran (izbor ga ne bi "
          "preskocio, a vreme bi pripisao tudjem koraku): %s -- pisi "
          "`If BfpIzabran(\"Test_X\") Then Test_X`" % ", ".join(gole[:5]))
    # Ime u izboru i pozvan test moraju biti isti: `If BfpIzabran("Test_A")
    # Then Test_B` bi za izbor Test_A pustio Test_B, a Test_B ne bi pustio nikad.
    razliciti = []
    for r in bfp + telo_delova:
        m = re.fullmatch(r'If BfpIzabran\("(\w+)"\) Then (\w+)', r)
        if m and m.group(1) != m.group(2):
            razliciti.append(r)
    tvrdi(not razliciti,
          "BfpIzabran pita za jedan test, a zove drugi: %s" % ", ".join(razliciti[:3]))
    return nalazi


def _terminate_pid(pid: int, flag: dict) -> None:
    """Ubij Excel proces. Zove se iz Timer niti kad glavni tok stoji na COM-u."""
    flag["fired"] = True
    try:
        import win32api
        import win32con
        h = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, pid)
        win32api.TerminateProcess(h, 0)
    except Exception:
        pass


# --- tvrdi prekid PO FAZI (KI-096) -------------------------------------------
#
# Budzet `--timeout` vazi za svaku fazu posebno: priprema (otvaranje, uvoz,
# sema), pa svaka suite, pa zatvaranje. Jedan tajmer za ceo poziv je obarao
# zdrav pun prolaz (BFP 688 s posle RunAllTests-a od 228 s, a granica je bila
# 600 s), uz FATAL poruku koja upucuje na break mode.
#
# KOBAN DIJALOG skracuje cekanje. Posle "Compile error" ili "Run-time error"
# Excel se ili vrati za nekoliko sekundi (watchdog je kliknuo End ili OK), ili
# je VBE u break mode-u i nece se vratiti nikad. Bez ovoga je takav prolaz visio
# do kraja budzeta -- 06.10.2026 je to bilo 585 s za "Syntax error" bez mesta,
# a sa danasnjih 1800 s bilo bi pola sata -- i poruka nije imenovala dijalog.
_KOBAN_DIJALOG = re.compile(r"compile error|run-time error|syntax error|"
                            r"fehler beim kompilieren|laufzeitfehler", re.IGNORECASE)
KOBAN_SEKUNDI = 15.0


def koban_dijalog(tekst: str) -> bool:
    """Da li je dijalog znak da je VBA stao (compile ili runtime greska)."""
    return bool(_KOBAN_DIJALOG.search(tekst or ""))


class Prekid:
    """Tvrdi prekid sa budzetom po fazi i skracenim cekanjem posle kobnog dijaloga.

    `stanje` je i zastavica za _terminate_pid ("fired"), i zapis za FATAL
    poruku: u kojoj fazi, posle kog dijaloga i posle koliko sekundi.
    `ubij` i `tajmer` su parametri samo zbog self-testa (bez Excela i bez cekanja).
    """

    def __init__(self, pid, budzet: float, ubij=None, tajmer=None):
        self.pid = pid
        self.budzet = budzet
        self._ubij = ubij or _terminate_pid
        self._tajmer = tajmer or threading.Timer
        self._lock = threading.Lock()
        self._t = None
        self._koban = False
        self.faza = ""
        self.stanje = {"fired": False, "faza": "", "dijalog": "", "sekundi": 0.0}

    def _pokreni(self, sekundi: float, dijalog: str = "") -> None:
        if self._t is not None:
            self._t.cancel()
        faza = self.faza

        def okini():
            # Belezi se PRVI uzrok: posle ubijanja i sledeca faza moze da istekne,
            # a FATAL treba da imenuje fazu u kojoj je Excel stvarno stao.
            if not self.stanje["fired"]:
                self.stanje.update(faza=faza, dijalog=dijalog, sekundi=sekundi)
            self._ubij(self.pid, self.stanje)

        self._t = self._tajmer(sekundi, okini)
        self._t.daemon = True
        self._t.start()

    def naoruzaj(self, faza: str) -> None:
        """Nova faza dobija pun budzet; skraceno cekanje iz prethodne se brise."""
        with self._lock:
            self.faza = faza
            self._koban = False
            self._pokreni(self.budzet)

    def dijalog(self, tekst: str) -> None:
        """Zove ga watchdog za SVAKI dijalog. Skracuje cekanje samo za koban, i
        samo jednom po fazi: dijalog koji se ne zatvori vidi se u svakom obilasku,
        pa bi ponovno naoruzavanje pomeralo prekid unedogled."""
        if not koban_dijalog(tekst):
            return
        with self._lock:
            if self._t is None or self._koban:
                return
            self._koban = True
            self._pokreni(KOBAN_SEKUNDI, tekst)

    def stop(self) -> None:
        with self._lock:
            if self._t is not None:
                self._t.cancel()
            self._t = None


# --- Compile ------------------------------------------------------------------

COMPILE_PROBE_MODULE = "modZzCompileProbe"
COMPILE_PROBE_FUNC = "ZzCompileProbe"
COMPILE_PROBE_CODE = (
    "Option Explicit\r\n"
    "\r\n"
    "Public Function ZzCompileProbe() As Long\r\n"
    "    ZzCompileProbe = 42\r\n"
    "End Function\r\n"
)


def compile_project(xl, wb, watchdog) -> dict:
    """Vrati {"ok": True|False|None, "error": ..., "detail": {...}}.

    Sta je ovde vec bilo netacno (tri lazna "zeleno"):

    - `Application.Run` nad probe funkcijom NE kompajlira ceo projekat. VBA
      kompajlira lenjo -- prevede modul koji se zove i ono sto taj modul stvarno
      referencira. Probe modul ne referencira nista, pa je greska u `modConfig`
      prosla netaknuta.
    - `FindControl(...).Execute()` bez VIDLJIVOG I AKTIVNOG VBE prozora ne uradi
      nista: kontrola je disabled, `Execute` tiho prodje, dijalog se ne pojavi --
      pa je izostanak dijaloga izgledao kao uspeh. Minimizovan prozor se broji
      kao neaktivan.

    Verdikt (redom):
      1. compile dijalog uhvacen  -> PAO
      2. kontrola postala disabled -> PROSAO (projekat je kompajliran)
      3. inace                     -> NEPOZNATO, nikad zeleno

    `detail` nosi sirovo stanje svakog pokusaja, da sledeci pogresan ishod ne
    mora opet da se pogadja.
    """
    detail: dict = {}

    try:
        vbe = xl.VBE
    except Exception as exc:            # noqa: BLE001
        return {"ok": None, "error": f"Nema pristupa VBE ({exc}) -- ukljuci "
                                     '"Trust access to the VBA project object model".',
                "detail": detail}

    # VBE prozor mora biti vidljiv I AKTIVAN. Prethodni pokusaj ga je
    # minimizovao "da ne skace preko ekrana" -- minimizovan prozor nije aktivan,
    # pa je meni kontrola ostala mrtva i `Execute` je tiho prolazio.
    try:
        vbe.MainWindow.Visible = True
        vbe.MainWindow.WindowState = 0      # vbext_ws_Normal
        try:
            vbe.MainWindow.SetFocus()
        except Exception:
            pass
        detail["vbe_visible"] = bool(vbe.MainWindow.Visible)
        detail["vbe_caption"] = str(vbe.MainWindow.Caption)
    except Exception as exc:                # noqa: BLE001
        detail["vbe_visible"] = f"greska: {exc}"

    try:
        vbe.ActiveVBProject = wb.VBProject
        detail["active_project"] = str(vbe.ActiveVBProject.Name)
    except Exception as exc:                # noqa: BLE001
        detail["active_project"] = f"greska: {exc}"

    try:
        ctl = vbe.CommandBars.FindControl(Id=COMPILE_CONTROL_ID)
    except Exception as exc:                # noqa: BLE001
        return {"ok": None, "error": f"FindControl pao: {exc}", "detail": detail}

    if ctl is None:
        return {"ok": None, "error": f"Nema kontrole 'Compile VBAProject' (Id={COMPILE_CONTROL_ID}).",
                "detail": detail}

    try:
        detail["enabled_before"] = bool(ctl.Enabled)
    except Exception as exc:                # noqa: BLE001
        detail["enabled_before"] = f"greska: {exc}"

    # Odmah posle importa projekat NIJE kompajliran -- kontrola mora biti
    # enabled. Ako nije, meni nam ne govori nista i ne smemo da tvrdimo ishod.
    if detail.get("enabled_before") is False:
        return {"ok": None,
                "error": "Kontrola 'Compile' je disabled JOS PRE compile-a -- meni ne "
                         "reaguje, ishod se ne moze utvrditi.",
                "detail": detail}

    # Pokusaj 1: COM Execute nad meni kontrolom.
    _attempt_compile(ctl, watchdog, detail, "com", lambda: ctl.Execute())

    # Pokusaj 2: SendKeys u VBE prozor (Alt+D, C = Debug > Compile VBAProject).
    # Ide samo ako COM put nije nista uradio -- i samo ako AppActivate potvrdi da
    # je VBE prozor zaista aktivan, da tastatura ne odleti u tudju aplikaciju.
    if detail.get("com_enabled_after") is True and not detail.get("com_dialogs"):
        _attempt_compile(ctl, watchdog, detail, "keys",
                         lambda: _sendkeys_compile(win32_shell(), vbe, detail))

    stage = "keys" if "keys_enabled_after" in detail else "com"
    dialogs = list(detail.get("com_dialogs") or []) + list(detail.get("keys_dialogs") or [])
    enabled_after = detail.get(f"{stage}_enabled_after")

    bad = [d for d in dialogs
           if any(k in d.lower() for k in ("compile", "kompajl", "greska", "error", "fehler"))]

    # Probe ostaje kao dodatna kontrola -- NE kompajlira ceo projekat (VBA je
    # lenj), ali hvata projekat koji uopste nije izvrsiv.
    detail["probe"] = _run_probe(xl, wb)

    if bad:
        return {"ok": False, "error": " ;; ".join(bad), "detail": detail}
    if enabled_after is False:
        return {"ok": True, "error": None, "detail": detail}

    # Ovde se NE sme tvrditi ni "prosao" ni "pao": nijedan compile dijalog nije
    # video svetlo dana, a kontrola je ostala aktivna -- sto znaci da compile
    # verovatno uopste nije izvrsen. Nepoznat ishod je nalaz, ne zeleno.
    return {"ok": None,
            "error": "Compile nije izvrsen (nema dijaloga, kontrola ostala aktivna) "
                     "-- ishod NEPOZNAT. Uradi rucno: Alt+F11 > Debug > Compile VBAProject.",
            "detail": detail}


def win32_shell():
    import win32com.client as win32
    return win32.Dispatch("WScript.Shell")


def _attempt_compile(ctl, watchdog, detail, prefix: str, action) -> None:
    """Odradi jedan pokusaj compile-a i zapisi sirovo stanje pod `prefix`."""
    before = len(watchdog.seen)
    try:
        action()
        detail[f"{prefix}_exec"] = "ok"
    except Exception as exc:                # noqa: BLE001
        detail[f"{prefix}_exec"] = f"greska: {exc}"
    time.sleep(2.5)
    detail[f"{prefix}_dialogs"] = watchdog.seen[before:]
    try:
        detail[f"{prefix}_enabled_after"] = bool(ctl.Enabled)
    except Exception as exc:                # noqa: BLE001
        detail[f"{prefix}_enabled_after"] = f"greska: {exc}"


def _sendkeys_compile(shell, vbe, detail) -> None:
    """Alt+D, C u VBE prozoru. Salje se SAMO ako je prozor stvarno aktiviran."""
    caption = str(vbe.MainWindow.Caption)
    active = bool(shell.AppActivate(caption))
    detail["keys_appactivate"] = active
    if not active:
        raise RuntimeError(f"AppActivate('{caption}') nije uspeo -- SendKeys se preskace")
    time.sleep(0.5)
    shell.SendKeys("%d")
    time.sleep(0.4)
    shell.SendKeys("c")


def _run_probe(xl, wb):
    """Dodaj trivijalan modul, pozovi ga, obrisi. Vrati vrednost ili tekst greske."""
    proj = wb.VBProject
    try:
        vbc = proj.VBComponents.Add(VBE_TYPE[".bas"])
        vbc.Name = COMPILE_PROBE_MODULE
        vbc.CodeModule.AddFromString(COMPILE_PROBE_CODE)
    except Exception as exc:                # noqa: BLE001
        return f"probe modul nije dodat: {exc}"

    try:
        return int(xl.Run(f"'{wb.Name}'!{COMPILE_PROBE_FUNC}") or 0)
    except Exception as exc:                # noqa: BLE001
        return f"Run pao: {exc}"
    finally:
        try:
            proj.VBComponents.Remove(proj.VBComponents(COMPILE_PROBE_MODULE))
        except Exception:
            pass


# --- Fixture -----------------------------------------------------------------

def create_blank_fixture(path: str, win32) -> None:
    """Napravi praznu .xlsm svesku na `path`.

    Za `--compile-only` je prazna sveska dovoljna, i to nije srecna slucajnost:
    nijedan modul ne referencira sheet CodeName rano-vezano (svi `SHT_*` u
    `modConfig` su string literali), a nijedan `.doccls` nema `Public` clan koji
    bi neko spolja zvao. Sheet `.doccls`-ovi se zato uredno preskoce ("nema
    komponente") i compile i dalje pokriva ceo `src-vba/`.

    Za suite-ove prazna sveska NIJE dovoljna -- one citaju `tblOtkup`,
    `tblKooperanti` i ostale tabele, pa tamo ide prava radna sveska kroz
    `--workbook`.
    """
    import pythoncom

    os.makedirs(os.path.dirname(path), exist_ok=True)
    # Zaseban COM apartment: ovo se zove PRE glavnog CoInitialize u main().
    pythoncom.CoInitialize()
    xl = win32.DispatchEx("Excel.Application")
    try:
        xl.Visible = False
        xl.DisplayAlerts = False
        wb = xl.Workbooks.Add()
        wb.SaveAs(path, FileFormat=XL_OPENXML_MACRO)
        wb.Close(SaveChanges=False)
    finally:
        try:
            xl.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


# --- izvor (D-6) ---------------------------------------------------------------

def izvor_iz_argumenata(args) -> tuple:
    """(apsolutna putanja izvora, razlog odbijanja ili "").

    Bez `--src` je to src-vba/. Folder bez ijednog .bas fajla nije izvor nego
    greska u pozivu: uvoz bi prosao bez ijednog modula, a suite bi pala na
    "Cannot run the macro" -- razlog bi pokazivao na VBA, ne na poziv.
    """
    if not getattr(args, "src", None):
        return SRC_VBA, ""
    src = os.path.abspath(args.src)
    if not os.path.isdir(src):
        return src, "--src %s: folder ne postoji" % src
    if not any(n.lower().endswith(".bas") for n in os.listdir(src)):
        return src, "--src %s: nema nijednog .bas fajla" % src
    return src, ""


def _self_test_src() -> list:
    """--src (D-6): uvoz i trazenje siroceta citaju BAS zadati folder.

    Lazna sveska belezi putanje koje joj se uvoze; SRC_VBA za to vreme pokazuje
    na MAMAC sa istim imenima, pa uvoz iz pogresnog foldera pada po imenu.
    """
    nalazi = []
    global SRC_VBA
    uvezeno, tela = [], []

    class _Kod:
        CountOfLines = 0

        def DeleteLines(self, *_a):
            pass

        def AddFromString(self, s):
            tela.append(s)

        def AddFromFile(self, p):
            uvezeno.append(p)

    class _Komp:
        def __init__(self, ime, tip=1):
            self.Name, self.Type, self.CodeModule = ime, tip, _Kod()

    class _Komponente:
        def __init__(self, imena):
            self.d = {i.lower(): _Komp(i) for i in imena}
            # dokument modul (Type 100): .doccls se u njega MERGE-uje
            self.d["thisworkbook"] = _Komp("ThisWorkbook", 100)

        def __call__(self, ime):
            return self.d[ime.lower()]

        def __iter__(self):
            return iter(list(self.d.values()))

        def Remove(self, _k):
            pass

        def Import(self, p):
            uvezeno.append(p)

        def Add(self, _tip):
            return _Komp("novi")

    class _Sveska:
        def __init__(self, imena):
            self.VBProject = type("P", (), {})()
            self.VBProject.VBComponents = _Komponente(imena)

    pravi = SRC_VBA
    privremeni = tempfile.mkdtemp(prefix="run_vba_src_")
    try:
        cilj = os.path.join(privremeni, "cilj")
        mamac = os.path.join(privremeni, "mamac")
        # Mamac nosi DRUGACIJI skup modula (modStari umesto modB): sirocad
        # racunata prema pogresnom folderu tako daju drugaciji spisak, pa se
        # i ta zamena vidi -- sa istim imenima bi bila nevidljiva.
        for d, imena in ((cilj, ("modA.bas", "modB.bas")),
                         (mamac, ("modA.bas", "modStari.bas"))):
            os.makedirs(d)
            for ime in imena:
                with open(os.path.join(d, ime), "w", newline="\r\n") as fh:
                    fh.write('Attribute VB_Name = "%s"\n' % ime[:-4])
            # .doccls ide drugim putem (merge u postojecu komponentu), pa ga
            # telo razlikuje: uvoz iz pogresnog foldera nosi rec MAMAC.
            with open(os.path.join(d, "ThisWorkbook.doccls"), "w",
                      newline="\r\n") as fh:
                fh.write("' %s\n" % ("CILJ" if d == cilj else "MAMAC"))
        # Forma sa .frx parom SAMO u cilju: par trazen u pogresnom folderu
        # daje "nema .frx para" i forma ispada iz uvoza.
        with open(os.path.join(cilj, "frmA.frm"), "w", newline="\r\n") as fh:
            fh.write("VERSION 5.00\n")
        with open(os.path.join(cilj, "frmA.frx"), "wb") as fh:
            fh.write(b"\0")
        SRC_VBA = mamac
        log: list = []
        try:
            import_src_vba(_Sveska([]), log, cilj)
        except Exception as exc:            # noqa: BLE001 -- pad je nalaz, ne traceback
            nalazi.append("SRC: uvoz iz --src je pukao (%s: %s) -- cita fajlove "
                          "van zadatog foldera" % (type(exc).__name__, exc))
        report_orphan_components(_Sveska(["modA", "modStari"]), log, cilj)
        args = argparse.Namespace(src=cilj)
        izv, razlog = izvor_iz_argumenata(args)
        prazan = os.path.join(privremeni, "prazan")
        os.makedirs(prazan)
        _i, razlog_prazan = izvor_iz_argumenata(argparse.Namespace(src=prazan))
        _i, razlog_nema = izvor_iz_argumenata(
            argparse.Namespace(src=os.path.join(privremeni, "nema")))
    finally:
        SRC_VBA = pravi
        shutil.rmtree(privremeni, ignore_errors=True)
    bez, razlog_bez = izvor_iz_argumenata(argparse.Namespace(src=None))

    if tela != ["' CILJ"]:
        nalazi.append("SRC: .doccls se ne uvozi iz --src foldera (tela: %r)" % tela)
    if not uvezeno or any(os.path.dirname(p) != cilj for p in uvezeno):
        nalazi.append("SRC: uvoz ne cita zadati --src folder (uvezeno: %r)"
                      % [os.path.basename(os.path.dirname(p)) for p in uvezeno])
    if sorted(os.path.basename(p) for p in uvezeno) != ["frmA.frm", "modA.bas",
                                                         "modB.bas"]:
        nalazi.append("SRC: uvoz iz --src nije uvezao tacno modA, modB i frmA (%r)"
                      % sorted(os.path.basename(p) for p in uvezeno))
    sirocad = [r for r in log if r.startswith("ORPHAN")]
    if len(sirocad) != 1 or "modStari" not in sirocad[0] or "modA" in sirocad[0]:
        nalazi.append("SRC: sirocad se ne racunaju prema --src folderu (%r)" % sirocad)
    if razlog or izv != cilj:
        nalazi.append("SRC: ispravan --src je odbijen (%r)" % razlog)
    if not razlog_prazan or not razlog_nema:
        nalazi.append("SRC: --src bez .bas fajlova ili nepostojeci folder je prihvacen")
    if razlog_bez or bez != pravi:
        nalazi.append("SRC: bez --src izvor nije src-vba/")
    return nalazi


# --- Glavni tok --------------------------------------------------------------

def parse_args(argv: list[str]) -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="Headless import + compile + VBA suite")
    ap.add_argument("--workbook", help="putanja do .xlsm (podrazumevano tests/fixtures/otkup_test.xlsm)")
    ap.add_argument("--compile-only", action="store_true", help="stani posle compile-a")
    ap.add_argument("--no-import", action="store_true", help="ne uvozi src-vba (testiraj zatecen kod)")
    ap.add_argument("--suite", action="append", default=[], help="pokreni bas ovu suite (moze vise puta)")
    ap.add_argument("--test", action="append", default=[], metavar="IME",
                    help="pusti samo ovaj test (moze vise puta, ili \"T_A,T_B\"); "
                         "trazi tacno jednu --suite koja ima `izbor` u SUITES")
    ap.add_argument("--all", action="store_true", help="pokreni sve suite iz kataloga")
    ap.add_argument("--keep", action="store_true", help="ne brisi temp kopiju sveske")
    ap.add_argument("--timeout-dialog", type=float, default=0.4, help="interval watchdog-a u sekundama")
    # Prekid se meri PO FAZI (priprema, svaka suite, zatvaranje), v. Prekid.
    # Najduza faza od 11.2026 je BFP (~700 s posle RunAllTests-a); 1800 s i dalje
    # hvata pravo visenje, a koban dijalog ga skracuje na KOBAN_SEKUNDI.
    ap.add_argument("--timeout", type=float, default=1800.0,
                    help="tvrdi prekid u sekundama za SVAKU fazu posebno (priprema, "
                         "svaka suite, zatvaranje); ubija Excel proces; default 1800")
    ap.add_argument("--bez-markera", action="store_true",
                    help="ne upisuj marker zelenog (dokaz.py nad namerno pokvarenim "
                         "izvorom, KI-036)")
    ap.add_argument("--src", metavar="DIR",
                    help="uvezi izvor iz ovog foldera umesto iz src-vba/ (dokaz.py "
                         "daje svoju kopiju, D-6); marker nosi otisak SADRZAJA, pa "
                         "vazi za src-vba tacno kad je sadrzaj isti")
    ap.add_argument("--self-test", action="store_true",
                    help="provere koje ne traze Excel (radi i na Linux/macOS)")
    ap.add_argument("--ignore-fixture-sig", action="store_true",
                    help="ne staj na ustajao fixture (svesno testiranje nad starim podacima)")
    return ap.parse_args(argv)


# --- ustajao fixture ---------------------------------------------------------
#
# Fixture je gitignored, pa ga `git checkout` NE menja: posle prelaska na drugu
# granu na disku ostaje sveska prethodne. Testovi tada padaju NA PODACIMA, a pad
# izgleda kao regresija koda -- tacno taj pad je jednom pojeo pola sata trijaze
# (cetiri crvena testa, kod ispravan).
#
# make_fixture.py pored sveske ostavlja `.sig` sa hash-om posejanih podataka.
# Poredi se PRE podizanja Excela: neslaganje je stanje diska, ne koda.

def proveri_fixture_potpis(fixture: str, koriscen_workbook: bool,
                           ignorisi: bool) -> int:
    """0 = nastavi. 2 = stani (ustajao fixture)."""
    if koriscen_workbook:
        return 0            # tudja sveska, generator nema sta da tvrdi o njoj

    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    try:
        import make_fixture
    except ImportError:
        return 0

    zapisan = make_fixture.read_sig(fixture)
    tekuci = make_fixture.signature()
    if zapisan == tekuci:
        return 0

    # Bash oblik sa "/": korisnikova ljuska je Git Bash, gde backslash NESTAJE bez
    # greske -- "tools\\make_fixture.py" postane "toolsmake_fixture.py", a poruka
    # pokazuje na fajl koji ne postoji umesto na pravi uzrok (CLAUDE.md).
    # Postojeci fixture sme da bude svoj donor; --out je tada obavezan.
    #
    # PREMESTAJU SE DVA FAJLA. Potpis stoji u pratecem .sig (FIXTURE_SIG_EXT) sa
    # istim korenom imena, pa `mv` samo .xlsm ostavlja stari .sig -- i provera
    # ustajalosti puca PONOVO, nad svezom sveskom. Tako izgleda kao da
    # regeneracija nije radila, a radila je.
    komanda = ('python tools/make_fixture.py --donor tests/fixtures/otkup_test.xlsm '
               '--out tests/fixtures/otkup_test_new.xlsm --force\n'
               '  mv tests/fixtures/otkup_test_new.xlsm tests/fixtures/otkup_test.xlsm\n'
               '  mv tests/fixtures/otkup_test_new.sig tests/fixtures/otkup_test.sig')
    if not zapisan:
        # FAIL-CLOSED. Sveska od generatora pre ovog sistema NE moze da se
        # proveri, a to je bas prvi run na svakoj zatecenoj masini -- tacno
        # onaj u kome se incident i desio. Upozorenje bi ga propustilo jednom
        # po masini, sto je isto kao da provere nema.
        #
        # Prazna auto-sveska ovde ne stize: nju pravi grana IZNAD ovog poziva.
        print(f"FIXTURE BEZ POTPISA: {fixture}", file=sys.stderr)
        print(f"  Sveska je od generatora pre uvodjenja potpisa, pa se ne moze",
              file=sys.stderr)
        print(f"  utvrditi da li su podaci svezi. Regenerisi je jednom:", file=sys.stderr)
        print(f"\n  {komanda}\n", file=sys.stderr)
        if ignorisi:
            print("  --ignore-fixture-sig: nastavljam ipak.", file=sys.stderr)
            return 0
        return 2

    print(f"USTAJAO FIXTURE: {fixture}", file=sys.stderr)
    print(f"  posejan potpisom {zapisan}, generator sada trazi {tekuci}", file=sys.stderr)
    print(f"  Fixture je gitignored -- `git checkout` ga NE menja, pa je na disku", file=sys.stderr)
    print(f"  ostala sveska prethodne grane. Testovi bi pali na PODACIMA.", file=sys.stderr)
    print(f"\n  {komanda}\n", file=sys.stderr)
    if ignorisi:
        print("  --ignore-fixture-sig: nastavljam ipak.", file=sys.stderr)
        return 0
    return 2


def izabrani_testovi(args: argparse.Namespace) -> list[str]:
    """Imena iz `--test`, bez praznih i bez duplikata, u redosledu navodjenja.

    Redosled izvrsavanja ionako odredjuje suite (VBA pusta izabrane u svom
    redosledu); ovde se samo cisti spisak.
    """
    out: list[str] = []
    for arg in args.test:
        for ime in arg.split(","):
            ime = ime.strip()
            if ime and ime not in out:
                out.append(ime)
    return out


def problem_izbora(args: argparse.Namespace) -> str:
    """Razlog zbog koga `--test` ne moze da se izvrsi, ili "" kad moze.

    Izbor je deo JEDNE suite: ime testa ne kaze kojoj pripada, a ulazna tacka
    izbora postoji samo za suite sa kljucem `izbor`. Svaka druga kombinacija bi
    tiho pustila celu suite (ili nijednu) i izgledala kao izbor.
    """
    if not args.test:
        return ""
    if not izabrani_testovi(args):
        return "--test bez ijednog imena"
    if args.compile_only:
        return "--test uz --compile-only nema sta da pusti"
    if args.all or len(args.suite) != 1:
        return "--test trazi TACNO JEDNU --suite (ime testa ne kaze kojoj pripada)"
    meta = SUITES.get(args.suite[0])
    if meta is None:
        return ""                    # nepoznatu suite prijavljuje chosen_suites
    if not meta.get("izbor"):
        sa = sorted(k for k, v in SUITES.items() if v.get("izbor"))
        return ("suite %s nema ulaznu tacku izbora (`izbor` u SUITES); "
                "izbor primaju: %s" % (args.suite[0], ", ".join(sa)))
    return ""


def vreme_fajl(result_file: str) -> str:
    """`last_run.txt` -> `last_run_vreme.txt`: vremena po testu pored rezultata."""
    koren, nastavak = os.path.splitext(result_file)
    return koren + "_vreme" + nastavak


def procitaj_vreme(path: str) -> list:
    """[(sekunde, status, ime)] iz fajla vremena, najsporiji prvi.

    Dijagnostika, ne verdikt: fajla nema ili je red necitljiv -> taj red se
    preskace, a run ne pada. Verdikt nosi rezultat-fajl.
    """
    out = []
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as fh:
            redovi = fh.read().split("\n")
    except OSError:
        return out
    for red in redovi:
        delovi = red.rstrip("\r").split("\t")
        if len(delovi) != 3:
            continue
        try:
            sek = float(delovi[0])
        except ValueError:
            continue
        out.append((sek, delovi[1], delovi[2]))
    out.sort(key=lambda x: -x[0])
    return out


VREME_NAJSPORIJIH = 10


def chosen_suites(args: argparse.Namespace) -> list[str]:
    if args.suite:
        unknown = [s for s in args.suite if s not in SUITES]
        if unknown:
            print(f"Nepoznata suite: {', '.join(unknown)}", file=sys.stderr)
            print(f"Poznate: {', '.join(sorted(SUITES))}", file=sys.stderr)
            raise SystemExit(2)
        return args.suite
    if args.all:
        return list(SUITES)
    return [k for k, v in SUITES.items() if v["default"]]


def main(argv: list[str]) -> int:
    args = parse_args(argv)

    if args.self_test:
        return self_test()

    # Pre Excela i pre provere platforme: pogresna kombinacija je greska u
    # pozivu, ne u VBA kodu, i mora da se vidi i na masini bez Excela.
    razlog = problem_izbora(args)
    if razlog:
        print("IZBOR: " + razlog, file=sys.stderr)
        return 2
    src, razlog = izvor_iz_argumenata(args)
    if razlog:
        print("IZVOR: " + razlog, file=sys.stderr)
        return 2

    if os.name != "nt":
        print("run_vba.py radi samo na Windows-u (Excel COM).", file=sys.stderr)
        return 2

    import pythoncom
    import win32api
    import win32com.client as win32
    import win32con
    import win32process

    # Vreme po fazi poziva (M-2): bez njega se zna samo zbir, pa se ne vidi sta
    # od ~18 s rezije po pozivu je start Excela, uvoz ili gasenje.
    pocetak = time.time()
    faze: list = []
    oznaka = [pocetak]

    def zabelezi_fazu(ime: str) -> None:
        sada = time.time()
        faze.append({"ime": ime, "sekundi": round(sada - oznaka[0], 1)})
        oznaka[0] = sada

    fixture = args.workbook or DEFAULT_FIXTURE
    if not os.path.exists(fixture):
        if args.workbook:
            print(f"Sveska ne postoji: {fixture}", file=sys.stderr)
            return 2
        # Podrazumevani fixture se pravi sam -- prazna sveska je dovoljna za
        # compile (vidi create_blank_fixture). Ovo se desi tacno jednom.
        print(f"Fixture ne postoji, pravim praznu svesku: {fixture}")
        try:
            create_blank_fixture(fixture, win32)
        except Exception as exc:
            print(f"Ne mogu da napravim fixture: {exc}", file=sys.stderr)
            print('Prosledi postojecu svesku: --workbook "...\\AgriX_OtkupApp.xlsm"',
                  file=sys.stderr)
            return 2
        if not args.compile_only:
            print("UPOZORENJE: prazna sveska nema tabele -- suite ce pasti na podacima.",
                  file=sys.stderr)
            print('Za suite koristi --workbook "...\\AgriX_OtkupApp.xlsm".', file=sys.stderr)
    elif not args.compile_only:
        # Compile ne cita podatke, pa ga ustajao fixture ne dotice.
        rc_sig = proveri_fixture_potpis(fixture, bool(args.workbook),
                                        args.ignore_fixture_sig)
        if rc_sig:
            return rc_sig

    tmp = tempfile.mkdtemp(prefix="vbatest_")
    wbpath = os.path.join(tmp, os.path.basename(fixture))
    shutil.copy2(fixture, wbpath)

    # Golden fajlovi idu u temp pre rana i vracaju se posle: modTest ih trazi
    # pored sveske, a nov golden mora da zavrsi u repou (tests/golden) na
    # ljudski pregled -- inace bi ga sledeci run tiho napravio ponovo.
    _copy_golden(GOLDEN_DIR, os.path.join(tmp, "golden"))

    # SNIMAK KONTEKSTA DOKAZA se uzima SADA, pre nego sto Excel krene -- i to
    # nad TEMP KOPIJOM sveske, jer je ona ono sto Excel stvarno otvara.
    #
    # Otisci uzeti na KRAJU run-a opisali bi stanje diska POSLE testova: prolaz
    # traje 20-60 minuta, a razvoj ide paralelno, pa bi izmena src-vba, golden
    # fajla ili fixture-a tokom run-a bila pripisana kao dokazana. `vba_gate`
    # pre upisa tvrdi da se nista od toga nije promenilo, i ne upisuje ako se
    # promenilo -- bolje nema dokaza nego dokaz o stanju koje nije mereno.
    _gate, kontekst_dokaza, kontekst_greska = None, None, ""
    try:
        _spec = importlib.util.spec_from_file_location(
            "_vba_gate_za_run", os.path.join(ROOT, "tools", "vba_gate.py"))
        _gate = importlib.util.module_from_spec(_spec)
        _spec.loader.exec_module(_gate)
        kontekst_dokaza = _gate.snimi_kontekst(src_dir=src, sveska=fixture,
                                               sveska_hes_iz=wbpath)
    except Exception as exc:                # noqa: BLE001
        kontekst_greska = "snimak konteksta nije uzet -- %s" % exc

    report: dict = {"workbook": fixture, "import": [], "compile": None, "suites": [], "dialogs": []}
    if src != SRC_VBA:
        report["izvor_dir"] = src
    rc = 2
    xl = None
    pid = None
    watchdog = None
    prekid = None
    zabelezi_fazu("priprema bez Excela")

    pythoncom.CoInitialize()
    try:
        xl = win32.DispatchEx("Excel.Application")
        pid = win32process.GetWindowThreadProcessId(xl.Hwnd)[1]

        # Tvrdi prekid: ako Excel iz bilo kog razloga prestane da odgovara (break
        # mode, dijalog koji watchdog ne prepozna), COM poziv ne puca -- samo
        # stoji. Bez ovoga run visi dok ga neko ne ubije rukom. Budzet je po fazi.
        prekid = Prekid(pid, args.timeout)
        prekid.naoruzaj("priprema")

        watchdog = DialogWatchdog(pid, poll=args.timeout_dialog,
                                  na_dijalog=prekid.dijalog)
        watchdog.start()

        xl.Visible = False
        xl.DisplayAlerts = False
        xl.AutomationSecurity = 1      # msoAutomationSecurityLow -- makroi bez pitanja
        xl.EnableEvents = False        # KLJUCNO: Workbook_Open (StartApp/self-update) se ne pokrece

        zabelezi_fazu("start Excela")
        wb = xl.Workbooks.Open(wbpath, UpdateLinks=0)
        zabelezi_fazu("otvaranje sveske")

        if not args.no_import:
            import_src_vba(wb, report["import"], src)
            report_orphan_components(wb, report["import"], src)
            zabelezi_fazu("uvoz src-vba")

        # COMPILE PROBA SAMO UZ --compile-only.
        #
        # Uz suite je proba bila ritual: verdikt je skoro uvek NEJASNO, ne obara
        # run i ne ide u marker, a jedinu zabelezenu pravu compile gresku
        # (06.10.2026, "Syntax error") nije uhvatila -- prolaz je visio 585 s.
        # Kostala je 5,9 s fiksnog cekanja po pozivu (COM pa SendKeys, izmereno
        # 16.11.2026), otvarala VBE i otimala fokus, a posle nje nekvalifikovan
        # Application.Run prestaje da radi. Projekat koji se ne kompajlira sada
        # hvata koban dijalog tokom suite (Prekid.dijalog): prekid za
        # KOBAN_SEKUNDI, uz tekst dijaloga. --compile-only ostaje kakav je: tu je
        # proba jedini signal, i on priprema svesku za rucni Compile.
        if args.compile_only:
            report["compile"] = compile_project(xl, wb, watchdog)
            zabelezi_fazu("compile proba")
        else:
            report["compile"] = {"ok": None, "preskocen": True, "detail": {},
                                 "error": "preskocena uz suite -- rucni Compile "
                                          "ostaje kapija pred release"}

        compile_ok = report["compile"].get("ok")

        if compile_ok is False:
            # Eksplicitan "Compile error: ..." dijalog je pravi signal -- ostaje fatalan.
            rc = 2
        elif args.compile_only:
            # Bez suite-ova je probe jedini izvor istine, pa NEJASNO i dalje pada:
            # alat koji ne zna ishod mora da kaze da ne zna, glasno.
            rc = 0 if compile_ok is True else 2
        else:
            # NEJASNO vise ne obara run kad suite-ovi mogu da daju pravi odgovor.
            # Da bi se RunAllTests uopste pokrenuo, VBA mora da kompajlira modTest
            # i sve sto on referencira -- a to je bas kod pod testom. Verdikt probe
            # se i dalje racuna i ispisuje nepromenjen, samo vise nije prepreka.
            # Fixture dolazi iz starijeg donora (npr. 2.28.4), a kod je noviji --
            # kolone dodate u medjuvremenu ne postoje dok se ne pokrene schema
            # upgrade. Bez ovoga je RunBusinessFlowProSuite davao 147/310 palih
            # provera, a uzrok je izgledao kao regresija. Rutina je idempotentna
            # (EnsureColumnOnTable je no-op kad kolona postoji), pa se vrti uvek.
            # Isti redosled trazi i modTestStornoCentar (v. komentar na vrhu tog
            # modula). Mora POSLE importa -- schema pravila dolaze iz svezeg koda.
            try:
            # POZIV JE KVALIFIKOVAN IMENOM SVESKE, i to nije stil.
            #
            # Izmereno 08.10.2026: posle koraka detekcije compile-a (VBE prozor +
            # tastaturne komande, verdikt NEJASNO) NEKVALIFIKOVAN Application.Run
            # prestane da prolazi -- "Cannot run the macro 'EnsureRuntimeSchema'" u
            # 0.0s, i to na SVAKOJ suite. Taj potpis docs i memorija vode kao modul
            # koji se ne kompajlira, pa alat lazno prijavi kvar projekta: direktan
            # kvalifikovan poziv nad istom temp kopijom prolazi i daje 204/0.
            #
            # Probe (_run_probe) je OD POCETKA kvalifikovan i zato je prolazio --
            # razlika je bila u pozivu, ne u projektu.
                xl.Run(f"'{wb.Name}'!EnsureRuntimeSchema")
                report["schema"] = "OK"
            except Exception as exc:        # noqa: BLE001
                report["schema"] = f"FAIL {exc}"
            zabelezi_fazu("sema")

            failed = 0
            testovi = izabrani_testovi(args)
            for suite in chosen_suites(args):
                meta = SUITES[suite]
                entry = {"name": suite, "gate": meta["gate"]}
                prekid.naoruzaj(suite)
                t0 = time.time()
                try:
                    if testovi:
                        # Izbor nosi ime suite-a u izvestaju, ali i oznaku
                        # `izbor` -- po njoj vba_gate ne upisuje marker.
                        entry["izbor"] = testovi
                        xl.Run(f"'{wb.Name}'!{meta['izbor']}", ",".join(testovi))
                    else:
                        xl.Run(f"'{wb.Name}'!{suite}")
                except Exception as exc:    # noqa: BLE001
                    entry["status"] = "FAIL"
                    entry["error"] = str(exc)
                    failed += 1
                    # Suite koja pad prijavljuje GRESKOM (banka) svejedno je
                    # upisala detalje pored sveske. Bez ovoga od celog pada
                    # ostane samo "Exception occurred" -- pa se ne vidi KOJA
                    # provera je pala, nego samo da jeste.
                    if meta.get("result_file"):
                        _read_test_results(tmp, report, meta["result_file"], suite)
                else:
                    if meta.get("result_file"):
                        # Verdikt iz fajla suite, ne iz "Run() nije pukao".
                        if _read_test_results(tmp, report, meta["result_file"],
                                              suite) == 0:
                            entry["status"] = "OK"
                        else:
                            entry["status"] = "FAIL"
                            failed += 1
                    else:
                        entry["status"] = "OK" if meta["gate"] else "BLIND"
                entry["seconds"] = round(time.time() - t0, 1)
                if meta.get("result_file"):
                    vreme = procitaj_vreme(os.path.join(
                        tmp, vreme_fajl(meta["result_file"])))
                    if vreme:
                        report.setdefault("suite_results", {}).setdefault(
                            suite, {})["vreme"] = vreme
                report["suites"].append(entry)
                oznaka[0] = time.time()      # suite nije rezija -- meri se sama
            # Pala priprema seme = rezultati nisu merodavni, pa run pada i kad su
            # sve suite zelene. Neuspela pretpostavka se ne precutkuje.
            rc = 2 if (failed or report.get("schema", "OK") != "OK") else 0

    except Exception as exc:                # noqa: BLE001
        report["fatal"] = str(exc)
        rc = 2
    finally:
        # Sveska se zatvara PRE gasenja cuvara. Close() ume da podigne
        # "Want to save your changes?" -- dovoljno je da jedna suite u svom
        # ciscenju vrati Application.DisplayAlerts na True. Ranije su i watchdog
        # i killer vec bili ugaseni na tom mestu, pa je taj dijalog visio zauvek:
        # nije imao ko da ga klikne ni ko da ubije proces.
        if prekid is not None:
            prekid.naoruzaj("zatvaranje")
        try:
            xl.DisplayAlerts = False
        except Exception:
            pass
        try:
            # SaveChanges eksplicitno -- goli Workbooks.Close() pita.
            #
            # Uz --keep se SNIMA: temp kopija se zadrzava bas zato da bi se u njoj
            # gledao trag rana (log sheet-ovi koje pisu suite, npr.
            # BUSINESS_FLOW_PRO_TEST_LOG za tools/read_test_log.py). Bez snimanja
            # se zadrzi fajl u stanju PRE rana, pa trijaza cita tudji, stariji run
            # i ne zna da ga cita. Original se ni ovde ne dira -- radi se nad temp
            # kopijom.
            while int(xl.Workbooks.Count) > 0:
                xl.Workbooks(1).Close(SaveChanges=bool(args.keep))
        except Exception:
            pass

        if prekid is not None:
            prekid.stop()
            if prekid.stanje["fired"]:
                report["fatal"] = poruka_prekida(prekid.stanje)
                report["prekid"] = dict(prekid.stanje)
                rc = 2
        if watchdog is not None:
            time.sleep(1.0)
            watchdog.stop()
            report["dialogs"] = watchdog.seen
        try:
            xl.Quit()
        except Exception:
            pass
        del xl
        pythoncom.CoUninitialize()
        time.sleep(1.0)
        if pid:
            try:
                h = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, pid)
                win32api.TerminateProcess(h, 0)
            except Exception:
                pass
        # Nov golden nastaje u temp-u -- vrati ga u repo da ga covek pregleda i
        # commit-uje. Bez ovoga bi ga svaki sledeci run pravio iznova i test bi
        # zauvek padao istom porukom.
        _copy_golden(os.path.join(tmp, "golden"), GOLDEN_DIR)

        if args.keep:
            print(f"\nTemp kopija zadrzana: {tmp}")
        else:
            shutil.rmtree(tmp, ignore_errors=True)
        zabelezi_fazu("zatvaranje i gasenje")
        report["faze"] = faze

    # MARKER ZELENOG: "suite su bile zelene" je tvrdnja o NEKOM izvoru, pod NEKIM
    # test sistemom, nad NEKOM svescom -- i do sada se znalo samo iz recenice uz
    # PR. Ovde se vezuje za snimak konteksta uzet PRE run-a, pa
    # `vba_gate.py --require-green` ume da razlikuje "dokazano" od "dokazano
    # nesto drugo". Pravila -- sta se NE sme upisati (pao run, --no-import, BLIND
    # suite kao dokazana, kontekst promenjen tokom run-a) -- zive u vba_gate,
    # zajedno sa svojim dokazom.
    #
    # Compile se ovde NE upisuje ni kad headless javi OK: compile je rucna kapija
    # operatera (`--mark-compile`), a headless verdikt je cesto NEJASNO. Sirovo
    # stanje i dalje ide u izvestaj, kao i do sada.
    #
    # Nalaz alata ne sme da obori run: greska u markeru nije greska u VBA kodu.
    try:
        if args.bez_markera:
            # KI-036: dokaz.py pusta suite nad NAMERNO pokvarenim izvorom. Prolaz
            # u kome sabotaza ne obori nista je zelen, pa bi upisao marker nad
            # izvorom koji nije cist -- i posle dokaza bi --require-green trazio
            # ponovljen prolaz.
            report["green"] = ("marker nije upisan: --bez-markera (prolaz nad "
                               "namerno pokvarenim izvorom)")
        elif kontekst_greska:
            report["green"] = "marker nije upisan: " + kontekst_greska
        else:
            report["green"] = _gate.zabelezi_prolaz(
                report, rc, args.no_import, kontekst=kontekst_dokaza, src_dir=src)
    except Exception as exc:                # noqa: BLE001
        report["green"] = "marker nije upisan: %s" % exc

    _write_report(report, rc)
    dopisi_istoriju(report, rc, argv, pocetak,
                    (kontekst_dokaza or {}).get("izvor", ""))
    return rc


def jedan_red(tekst: str) -> str:
    """Tekst dijaloga u jednom redu: VBA ga lomi ("Compile error:\n\nSyntax error"),
    a citaci izvestaja (dokaz.py, FATAL) citaju red po red -- pa bi bez ovoga
    ostali bez razloga."""
    return " ".join((tekst or "").split())


def poruka_prekida(stanje: dict) -> str:
    """FATAL tekst iz Prekid.stanje: imenuje fazu, a uz koban dijalog i dijalog."""
    if stanje.get("dijalog"):
        return (f"Excel se nije vratio {stanje['sekundi']:g}s posle dijaloga u fazi "
                f"'{stanje['faza']}' -- proces je ubijen. Dijalog: "
                f"{jedan_red(stanje['dijalog'])}. "
                "Najcesci uzrok: kod se ne kompajlira ili je VBE u break mode-u.")
    return (f"Faza '{stanje['faza']}' je trajala duze od {stanje['sekundi']:g}s -- "
            "proces je ubijen. Uzrok je break mode u VBE-u ILI faza duza od --timeout "
            "(prekid se meri po fazi); vidi docs/EXCEL_TEST_HARNESS.md.")


ISTORIJA = os.path.join(ROOT, "tests", "run_vba_istorija.jsonl")


def dopisi_istoriju(report: dict, rc: int, argv: list, pocetak: float,
                    izvor: str, put: str = ISTORIJA) -> str:
    """Dopisi JEDAN red po pozivu (M-2). Vraca "" ili razlog zasto nije upisano.

    tests/last_run.* prepisuje svaki poziv, pa se trajanja do sada nisu nigde
    cuvala. Greska u upisu istorije ne sme da obori run: to nije nalaz o kodu.
    """
    red = {
        "pocetak": time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime(pocetak)),
        "sekundi": round(time.time() - pocetak, 1),
        "argv": list(argv),
        "rc": rc,
        "izvor": izvor,
        "faze": report.get("faze", []),
        "suite": [{k: s.get(k) for k in ("name", "status", "seconds")}
                  for s in report.get("suites", [])],
        "fatal": report.get("fatal", ""),
    }
    try:
        os.makedirs(os.path.dirname(put), exist_ok=True)
        with open(put, "a", encoding="utf-8") as fh:
            fh.write(json.dumps(red, ensure_ascii=True) + "\n")
        return ""
    except Exception as exc:                # noqa: BLE001
        return str(exc)


def _write_report(report: dict, rc: int) -> None:
    outdir = os.path.join(ROOT, "tests")
    os.makedirs(outdir, exist_ok=True)
    with open(os.path.join(outdir, "last_run.json"), "w", encoding="utf-8") as fh:
        json.dump(report, fh, ensure_ascii=False, indent=2)

    lines = []
    if report.get("izvor_dir"):
        lines.append(f"IZVOR   {report['izvor_dir']}")
    for msg in report["import"]:
        lines.append(f"IMPORT  {msg}")

    c = report.get("compile")
    if c and c.get("preskocen"):
        lines.append(f"COMPILE PRESKOCEN  {c['error']}")
    elif c:
        state = {True: "OK", False: "FAIL", None: "NEJASNO"}[c.get("ok")]
        lines.append(f"COMPILE {state}" + (f"  {c['error']}" if c.get("error") else ""))
        # Sirovo stanje ide u ispis uvek: tri puta je "zeleno" bilo netacno, pa
        # jedan run mora da nosi dovoljno podataka za dijagnozu.
        for key, value in (c.get("detail") or {}).items():
            lines.append(f"        {key} = {value!r}")

    schema = report.get("schema")
    if schema:
        lines.append(f"SCHEMA  {schema}")

    # Rezultat svake suite ide UZ NJU i nosi njeno ime. Jedan zajednicki blok
    # je znacio da poslednja suite prepise prethodnu.
    slotovi = report.get("suite_results", {})
    for s in report["suites"]:
        lines.append(f"SUITE   {s['status']:6} {s['name']} ({s['seconds']}s)"
                     + (f"  {s.get('error', '')}" if s["status"] == "FAIL" else ""))
        if s.get("izbor"):
            # Izbor mora da se vidi u istom bloku: "SUITE OK RunAllTests" bez
            # ovoga cita se kao cela suite.
            lines.append(f"IZBOR   {s['name']}: {len(s['izbor'])} testova -- "
                         + ", ".join(s["izbor"]))
        t = slotovi.get(s["name"])
        if not t:
            continue
        if t.get("error"):
            lines.append(f"TESTS   {s['name']}: {t['error']}")
        else:
            lines.append(f"TESTS   {s['name']}: {t['total']} ukupno, "
                         f"{t['failed']} palo")
        # Ime bas tog testa mora da se vidi u ispisu -- to je razlika izmedju
        # "nesto je palo" i upotrebljivog nalaza.
        for ln in t.get("detail", []):
            lines.append(f"        {ln}")
        # Najsporiji koraci. Red pocinje brojem, ne statusom, pa ga citaci
        # palih (dokaz.py: "FAIL ...", banka: "PAO ...") ne mogu procitati kao pad.
        vreme = t.get("vreme") or []
        if vreme:
            lines.append(f"VREME   {s['name']}: {sum(v[0] for v in vreme):.1f}s "
                         f"u {len(vreme)} koraka; najsporiji:")
            for sek, status, ime in vreme[:VREME_NAJSPORIJIH]:
                lines.append(f"        {sek:8.3f}s  {status:5} {ime}")

    # Rezija poziva po fazi (sve sto nije suite). Red pocinje sa FAZE, pa ga
    # citaci palih ne mogu procitati kao pad.
    faze = report.get("faze") or []
    if faze:
        lines.append("FAZE    " + " | ".join(f"{f['ime']} {f['sekundi']:.1f}s" for f in faze)
                     + f"  -> rezija {sum(f['sekundi'] for f in faze):.1f}s")

    for d in report.get("dialogs", []):
        lines.append(f"DIALOG  {jedan_red(d)}")

    if "fatal" in report:
        lines.append(f"FATAL   {report['fatal']}")

    if report.get("green"):
        lines.append(f"GREEN   {report['green']}")

    lines.append("")
    lines.append("REZULTAT: " + ("ZELENO" if rc == 0 else "PALO"))
    blind = [s["name"] for s in report["suites"] if s["status"] == "BLIND"]
    if blind:
        lines.append("BLIND (suite bez fail-gate-a -- prosla bez greske NIJE dokaz da su "
                     "sve provere prosle): " + ", ".join(blind))

    text = "\n".join(lines)
    with open(os.path.join(outdir, "last_run.txt"), "w", encoding="utf-8") as fh:
        fh.write(text + "\n")
    # Ispis ne sme da PUKNE zbog jednog znaka. Konzola je cp1252, a poruka o
    # padu nosi ono sto je test video -- pa je jedan ChrW(183) iz prikaznog
    # teksta parcele rusio ceo izvestaj: run bi zavrsio Traceback-om UMESTO
    # linijom "FAIL <test> -- <tvrdnja>". U petlji dvosmernog dokaza to izgleda
    # kao "sabotaza nije oborila nista", pa je jedna ispravna sabotaza
    # (parcela-tekst) bila prijavljena kao mrtva.
    try:
        print(text)
    except UnicodeEncodeError:
        enc = sys.stdout.encoding or "utf-8"
        print(text.encode(enc, errors="replace").decode(enc, errors="replace"))


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
