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

Sveska: bez `--workbook` koristi se `tests/fixtures/otkup_test.xlsm`, a ako ga nema,
skript ga sam napravi kao PRAZNU .xlsm. Za compile je to dovoljno; suite-ovima
trebaju podaci, pa im prosledi pravu radnu svesku kroz `--workbook`. Sveska se
uvek kopira u temp -- original se ne dira. Detalji: docs/EXCEL_TEST_HARNESS.md.

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
    "RunIzvestajTests":         {"gate": True,  "dialogs": False, "default": True},
    "RunSheetsJsonParserTests": {"gate": True,  "dialogs": True,  "default": True},
    "RunBankaImportTestSuite":  {"gate": True,  "dialogs": True,  "default": True},
    "RunFakturaSmokeSuite":     {"gate": True,  "dialogs": True,  "default": True},
    "Test_StornoCentar_All":    {"gate": False, "dialogs": False, "default": True},
    "TestLicense_All":          {"gate": False, "dialogs": False, "default": True},
    # Nisu u podrazumevanom setu: traze mrezu, live SEF nalog ili duze rade.
    "RunGoogleSyncSmokeSuite":  {"gate": True,  "dialogs": True,  "default": False},
    "RunMasterSyncSmokeSuite":  {"gate": True,  "dialogs": True,  "default": False},
    "RunSEFTestSuite":          {"gate": True,  "dialogs": True,  "default": False},
    "RunStornoTestSuite":       {"gate": False, "dialogs": True,  "default": False},
    "RunPaleteTestSuite":       {"gate": False, "dialogs": True,  "default": False},
    "RunNovacSmokeSuite":       {"gate": False, "dialogs": True,  "default": False},
    "RunBusinessFlowProSuite":  {"gate": False, "dialogs": True,  "default": False},
    "RunAgrohemijaSmokeSuite":  {"gate": False, "dialogs": True,  "default": False},
    "RunProductionHealthCheck": {"gate": False, "dialogs": True,  "default": False},
    "TestMonitoring_All":       {"gate": False, "dialogs": False, "default": False},
}

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_VBA = os.path.join(ROOT, "src-vba")
SELF_MODULE = "modVbaTools"          # isto kao modVbaTools.SELF_MODULE
COMPILE_CONTROL_ID = 578             # VBE: Debug > Compile VBAProject

VBE_TYPE = {".bas": 1, ".cls": 2, ".frm": 3, ".doccls": 100}

DEFAULT_FIXTURE = os.path.join(ROOT, "tests", "fixtures", "otkup_test.xlsm")
XL_OPENXML_MACRO = 52                # xlOpenXMLWorkbookMacroEnabled (.xlsm)


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
        buttons: list[int] = []

        def on_child(child, _):
            if win32gui.GetClassName(child) == "Button":
                buttons.append(child)

        try:
            win32gui.EnumChildWindows(hwnd, on_child, None)
        except Exception:
            pass

        BM_CLICK = 0x00F5
        WM_CLOSE = 0x0010
        if buttons:
            win32gui.PostMessage(buttons[0], BM_CLICK, 0, 0)
        else:
            win32gui.PostMessage(hwnd, WM_CLOSE, 0, 0)


# --- Import src-vba preko COM-a ----------------------------------------------

def _read_code_body(path: str) -> str:
    """Vrati samo kod, bez VBA header bloka -- za .doccls merge."""
    with open(path, "r", encoding="ascii", errors="replace") as fh:
        lines = fh.read().splitlines()

    out, started = [], False
    for line in lines:
        if not started:
            stripped = line.strip()
            if stripped.startswith("VERSION") or stripped.startswith("Attribute VB_Name"):
                continue
            if stripped.startswith("BEGIN") or stripped == "END" or stripped.startswith("Attribute "):
                continue
            if not stripped:
                continue
            started = True
        out.append(line)
    return "\r\n".join(out)


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
    """Vrati {"ok": True|False|None, "error": ...} za compile celog projekta.

    NE oslanja se na `Enabled` stanje menija "Debug > Compile VBAProject". VBE
    osvezava enabled-stanje kontrola tek kad se meni iscrta, pa u nevidljivom
    Excelu ostaje `True` i posle uspesnog compile-a -- ta heuristika je davala
    "NEJASNO" uvek.

    Verdikt daje probe: u projekat se doda trivijalan modul sa funkcijom koja
    vraca 42, pa se ta funkcija pozove. VBA pred izvrsavanje BILO KOJE procedure
    kompajlira ceo projekat, pa greska u ma kom modulu obara `Application.Run`.
    Vracenih 42 znaci da se ceo projekat kompajlira -- to je tvrdo DA.
    """
    before = len(watchdog.seen)

    # 1) Forsiraj pun compile kroz meni. Greska ovde ide kroz modalni dijalog,
    #    koji watchdog zatvori i ostavi u `seen`.
    try:
        ctl = xl.VBE.CommandBars.FindControl(Id=COMPILE_CONTROL_ID)
        ctl.Execute()
        time.sleep(1.0)
    except Exception as exc:            # noqa: BLE001
        return {"ok": False, "error": f"Compile meni: {exc}"}

    dialogs = watchdog.seen[before:]
    bad = [d for d in dialogs
           if any(k in d.lower() for k in ("compile", "kompajl", "greska", "error"))]
    if bad:
        return {"ok": False, "error": " ;; ".join(bad)}

    # 2) Probe -- jedini korak koji daje tvrd DA.
    proj = wb.VBProject
    try:
        vbc = proj.VBComponents.Add(VBE_TYPE[".bas"])
        vbc.Name = COMPILE_PROBE_MODULE
        vbc.CodeModule.AddFromString(COMPILE_PROBE_CODE)
    except Exception as exc:            # noqa: BLE001
        return {"ok": None, "error": f"Ne mogu da dodam probe modul: {exc}"}

    try:
        value = xl.Run(f"'{wb.Name}'!{COMPILE_PROBE_FUNC}")
    except Exception as exc:            # noqa: BLE001
        return {"ok": False, "error": f"Probe pao -- compile greska u projektu: {exc}"}
    finally:
        try:
            proj.VBComponents.Remove(proj.VBComponents(COMPILE_PROBE_MODULE))
        except Exception:
            pass

    if int(value or 0) != 42:
        return {"ok": None, "error": f"Probe vratio {value!r} umesto 42."}
    return {"ok": True, "error": None}


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

    report: dict = {"workbook": fixture, "import": [], "compile": None, "suites": [], "dialogs": []}
    rc = 2
    xl = None
    pid = None
    watchdog = None

    pythoncom.CoInitialize()
    try:
        xl = win32.DispatchEx("Excel.Application")
        pid = win32process.GetWindowThreadProcessId(xl.Hwnd)[1]

        watchdog = DialogWatchdog(pid, poll=args.timeout_dialog)
        watchdog.start()

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

        if report["compile"].get("ok") is not True:
            # Sve osim eksplicitnog True je pad. "Nejasno" ne sme da prodje kao
            # zeleno -- alat koji ne zna ishod mora da kaze da ne zna, glasno.
            rc = 2
        elif args.compile_only:
            rc = 0
        else:
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
                    entry["status"] = "OK" if meta["gate"] else "BLIND"
                entry["seconds"] = round(time.time() - t0, 1)
                report["suites"].append(entry)
            rc = 2 if failed else 0

    except Exception as exc:                # noqa: BLE001
        report["fatal"] = str(exc)
        rc = 2
    finally:
        if watchdog is not None:
            time.sleep(1.0)
            watchdog.stop()
            report["dialogs"] = watchdog.seen
        try:
            xl.Workbooks.Close()
        except Exception:
            pass
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

    for s in report["suites"]:
        lines.append(f"SUITE   {s['status']:6} {s['name']} ({s['seconds']}s)"
                     + (f"  {s.get('error', '')}" if s["status"] == "FAIL" else ""))

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
