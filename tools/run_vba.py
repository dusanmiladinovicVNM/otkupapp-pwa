"""Headless VBA runner za AgriX OtkupApp: import src-vba -> compile -> test suite.

Radi SAMO na Windows masini sa instaliranim Excelom (COM). Na macOS/Linux nema
VBProject COM interfejsa -- VBA sesije se vode na Windows kutiji.

    pip install pywin32
    File > Options > Trust Center > Trust Center Settings > Macro Settings
        -> "Trust access to the VBA project object model"

Upotreba:
    python tools/run_vba.py --compile-only          # samo import + compile (najbrze, najstabilnije)
    python tools/run_vba.py                         # + podrazumevani set suite-ova
    python tools/run_vba.py --suite RunIzvestajTests
    python tools/run_vba.py --all
    python tools/run_vba.py --workbook "C:\\putanja\\AgriX_OtkupApp.xlsm"

Sveska: bez `--workbook` koristi se `tests/fixtures/otkup_test.xlsm`. Ako ga nema,
skript napravi PRAZNU .xlsm -- dovoljno za compile, ali NE i za suite (prazna
sveska nema tabele). Pravi fixture se pravi sa `tools/make_fixture.py`; sadrzi
samo sinteticke podatke, a suite koje diraju tabele ionako seju sebi svoje
(SVT-*, BIT-*, TST-*) u transakciji koja se uvek ponistava -- prava radna sveska
im NIJE potrebna. Sveska se uvek kopira u temp, original se ne dira.
Detalji: docs/EXCEL_TEST_HARNESS.md.

Izlazni kod: 0 = zeleno, 2 = palo (compile greska, pala suite, ili neocekivan dijalog).

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
import json
import os
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
    "RunAllTests":              {"gate": True,  "dialogs": False, "default": True,
                                 "result_file": True},
    "RunIzvestajTests":         {"gate": True,  "dialogs": False, "default": True},
    "RunSheetsJsonParserTests": {"gate": True,  "dialogs": True,  "default": True},
    "RunBankaImportTestSuite":  {"gate": True,  "dialogs": True,  "default": True},
    "RunFakturaSmokeSuite":     {"gate": True,  "dialogs": True,  "default": True},
    "Test_StornoCentar_All":    {"gate": True,  "dialogs": False, "default": True},
    "TestLicense_All":          {"gate": False, "dialogs": False, "default": True},
    # Nisu u podrazumevanom setu: traze mrezu, live SEF nalog ili duze rade.
    "RunGoogleSyncSmokeSuite":  {"gate": True,  "dialogs": True,  "default": False},
    "RunMasterSyncSmokeSuite":  {"gate": True,  "dialogs": True,  "default": False},
    "RunSEFTestSuite":          {"gate": True,  "dialogs": True,  "default": False},
    "RunStornoTestSuite":       {"gate": True,  "dialogs": True,  "default": True},
    "RunPaleteTestSuite":       {"gate": True,  "dialogs": True,  "default": True},
    "RunNovacSmokeSuite":       {"gate": False, "dialogs": True,  "default": False},
    "RunBusinessFlowProSuite":  {"gate": True,  "dialogs": True,  "default": True},
    "RunAgrohemijaSmokeSuite":  {"gate": True,  "dialogs": True,  "default": True},
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


def _read_test_results(wbdir: str, report: dict) -> int:
    """Verdikt iz last_run.txt koji je upisao modTest.RunAllTests.

    Nema fajla = pad, ne prolaz: to znaci da RunAllTests nije stigao do kraja
    (compile error, visenje, ubijen proces) -- ishod koji nije eksplicitno OK.
    """
    path = os.path.join(wbdir, "last_run.txt")
    if not os.path.exists(path):
        report["tests"] = {"error": "modTest nije upisao last_run.txt "
                                    "(RunAllTests nije stigao do kraja)"}
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
    report["tests"] = {"total": total, "failed": failed, "detail": detail}

    if total < 0 or failed < 0:
        report["tests"]["error"] = f"neocekivan prvi red: {head!r}"
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
    def __init__(self, pid: int, poll: float = 0.4):
        super().__init__(daemon=True)
        self.pid = pid
        self.poll = poll
        self.seen: list[str] = []
        self._stop = threading.Event()

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


def import_src_vba(wb, log: list[str]) -> None:
    proj = wb.VBProject
    files = sorted(os.listdir(SRC_VBA))

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
        cm.AddFromString(_read_code_body(os.path.join(SRC_VBA, name)))

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
        path = os.path.join(SRC_VBA, name)

        if ext == ".frm" and not os.path.exists(os.path.join(SRC_VBA, base + ".frx")):
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

    for line in leaks:
        print(line, file=sys.stderr)
    if leaks:
        print(f"\nself-test: {len(leaks)} nalaza od {checked} .doccls fajlova.", file=sys.stderr)
        return 2
    print(f"self-test: cisto ({checked} .doccls fajlova).")
    return 0


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


# --- Glavni tok --------------------------------------------------------------

def parse_args(argv: list[str]) -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="Headless import + compile + VBA suite")
    ap.add_argument("--workbook", help="putanja do .xlsm (podrazumevano tests/fixtures/otkup_test.xlsm)")
    ap.add_argument("--compile-only", action="store_true", help="stani posle compile-a")
    ap.add_argument("--no-import", action="store_true", help="ne uvozi src-vba (testiraj zatecen kod)")
    ap.add_argument("--suite", action="append", default=[], help="pokreni bas ovu suite (moze vise puta)")
    ap.add_argument("--all", action="store_true", help="pokreni sve suite iz kataloga")
    ap.add_argument("--keep", action="store_true", help="ne brisi temp kopiju sveske")
    ap.add_argument("--timeout-dialog", type=float, default=0.4, help="interval watchdog-a u sekundama")
    ap.add_argument("--timeout", type=float, default=600.0,
                    help="tvrdi prekid u sekundama (ubija Excel proces; default 600)")
    ap.add_argument("--self-test", action="store_true",
                    help="provere koje ne traze Excel (radi i na Linux/macOS)")
    return ap.parse_args(argv)


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

    if os.name != "nt":
        print("run_vba.py radi samo na Windows-u (Excel COM).", file=sys.stderr)
        return 2

    import pythoncom
    import win32api
    import win32com.client as win32
    import win32con
    import win32process

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

    tmp = tempfile.mkdtemp(prefix="vbatest_")
    wbpath = os.path.join(tmp, os.path.basename(fixture))
    shutil.copy2(fixture, wbpath)

    # Golden fajlovi idu u temp pre rana i vracaju se posle: modTest ih trazi
    # pored sveske, a nov golden mora da zavrsi u repou (tests/golden) na
    # ljudski pregled -- inace bi ga sledeci run tiho napravio ponovo.
    _copy_golden(GOLDEN_DIR, os.path.join(tmp, "golden"))

    report: dict = {"workbook": fixture, "import": [], "compile": None, "suites": [], "dialogs": []}
    rc = 2
    xl = None
    pid = None
    watchdog = None
    killer = None
    hard_stop: dict = {"fired": False}

    pythoncom.CoInitialize()
    try:
        xl = win32.DispatchEx("Excel.Application")
        pid = win32process.GetWindowThreadProcessId(xl.Hwnd)[1]

        watchdog = DialogWatchdog(pid, poll=args.timeout_dialog)
        watchdog.start()

        # Tvrdi prekid: ako Excel iz bilo kog razloga prestane da odgovara (break
        # mode, dijalog koji watchdog ne prepozna), COM poziv ne puca -- samo
        # stoji. Bez ovoga run visi dok ga neko ne ubije rukom.
        killer = threading.Timer(args.timeout, _terminate_pid, args=(pid, hard_stop))
        killer.daemon = True
        killer.start()

        xl.Visible = False
        xl.DisplayAlerts = False
        xl.AutomationSecurity = 1      # msoAutomationSecurityLow -- makroi bez pitanja
        xl.EnableEvents = False        # KLJUCNO: Workbook_Open (StartApp/self-update) se ne pokrece

        wb = xl.Workbooks.Open(wbpath, UpdateLinks=0)

        if not args.no_import:
            import_src_vba(wb, report["import"])

        # Compile se radi UVEK -- i uz --compile-only i pre suite-ova. To je
        # najjeftiniji gate koji hvata najcescu klasu kvara posle Edit/Write nad
        # src-vba (duple definicije, deklaracija posle prve procedure, ime koje
        # se poklapa sa rezervisanom reci).
        report["compile"] = compile_project(xl, wb, watchdog)

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
                xl.Run("EnsureRuntimeSchema")
                report["schema"] = "OK"
            except Exception as exc:        # noqa: BLE001
                report["schema"] = f"FAIL {exc}"

            failed = 0
            for suite in chosen_suites(args):
                meta = SUITES[suite]
                entry = {"name": suite, "gate": meta["gate"]}
                t0 = time.time()
                try:
                    xl.Run(suite)
                except Exception as exc:    # noqa: BLE001
                    entry["status"] = "FAIL"
                    entry["error"] = str(exc)
                    failed += 1
                else:
                    if meta.get("result_file"):
                        # Verdikt iz last_run.txt, ne iz "Run() nije pukao".
                        if _read_test_results(tmp, report) == 0:
                            entry["status"] = "OK"
                        else:
                            entry["status"] = "FAIL"
                            failed += 1
                    else:
                        entry["status"] = "OK" if meta["gate"] else "BLIND"
                entry["seconds"] = round(time.time() - t0, 1)
                report["suites"].append(entry)
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

        if killer is not None:
            killer.cancel()
        if hard_stop["fired"]:
            report["fatal"] = (f"Excel nije odgovarao {args.timeout:g}s -- proces je ubijen. "
                               "Najcesci uzrok: VBE je u break mode-u (vidi "
                               "docs/EXCEL_TEST_HARNESS.md).")
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

    _write_report(report, rc)
    return rc


def _write_report(report: dict, rc: int) -> None:
    outdir = os.path.join(ROOT, "tests")
    os.makedirs(outdir, exist_ok=True)
    with open(os.path.join(outdir, "last_run.json"), "w", encoding="utf-8") as fh:
        json.dump(report, fh, ensure_ascii=False, indent=2)

    lines = []
    for msg in report["import"]:
        lines.append(f"IMPORT  {msg}")

    c = report.get("compile")
    if c:
        state = {True: "OK", False: "FAIL", None: "NEJASNO"}[c.get("ok")]
        lines.append(f"COMPILE {state}" + (f"  {c['error']}" if c.get("error") else ""))
        # Sirovo stanje ide u ispis uvek: tri puta je "zeleno" bilo netacno, pa
        # jedan run mora da nosi dovoljno podataka za dijagnozu.
        for key, value in (c.get("detail") or {}).items():
            lines.append(f"        {key} = {value!r}")

    schema = report.get("schema")
    if schema:
        lines.append(f"SCHEMA  {schema}")

    for s in report["suites"]:
        lines.append(f"SUITE   {s['status']:6} {s['name']} ({s['seconds']}s)"
                     + (f"  {s.get('error', '')}" if s["status"] == "FAIL" else ""))

    t = report.get("tests")
    if t:
        if t.get("error"):
            lines.append(f"TESTS   {t['error']}")
        else:
            lines.append(f"TESTS   {t['total']} ukupno, {t['failed']} palo")
        # Ime bas tog testa mora da se vidi u ispisu -- to je razlika izmedju
        # "nesto je palo" i upotrebljivog nalaza.
        for ln in t.get("detail", []):
            lines.append(f"        {ln}")

    for d in report.get("dialogs", []):
        lines.append(f"DIALOG  {d}")

    if "fatal" in report:
        lines.append(f"FATAL   {report['fatal']}")

    lines.append("")
    lines.append("REZULTAT: " + ("ZELENO" if rc == 0 else "PALO"))
    blind = [s["name"] for s in report["suites"] if s["status"] == "BLIND"]
    if blind:
        lines.append("BLIND (suite bez fail-gate-a -- prosla bez greske NIJE dokaz da su "
                     "sve provere prosle): " + ", ".join(blind))

    text = "\n".join(lines)
    with open(os.path.join(outdir, "last_run.txt"), "w", encoding="utf-8") as fh:
        fh.write(text + "\n")
    print(text)


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
