"""Headless VBA runner za AgriX OtkupApp: static -> import -> compile -> schema -> suite.

Excel deo radi SAMO na Windows masini sa instaliranim Excelom (COM). Na
macOS/Linux nema VBProject COM interfejsa -- tamo se ispise pun verdikt sa
`BEHAVIOR NOT_RUN`, sto je tacniji ishod od ranijeg "radi samo na Windows-u".

    pip install pywin32
    File > Options > Trust Center > Trust Center Settings > Macro Settings
        -> "Trust access to the VBA project object model"

Upotreba:
    python tools/run_vba.py                         # kapija 'dev' (isto sto i Stop hook)
    python tools/run_vba.py --gate pr               # + blind/spore lokalne suite
    python tools/run_vba.py --gate release          # + external (Google/MasterSync/SEF)
    python tools/run_vba.py --compile-only          # samo import + compile
    python tools/run_vba.py --suite RunIzvestajTests
    python tools/run_vba.py --category unit
    python tools/run_vba.py --shuffle --seed 12345  # nezavisnost od redosleda
    python tools/run_vba.py --repeat 3              # isti proces, vise prolaza
    python tools/run_vba.py --workbook "C:\\putanja\\AgriX_OtkupApp.xlsm"

Katalog suite-ova NIJE ovde nego u `tests/suite_manifest.json` -- jedini izvor
istine, koji cita i `vba_gate.py --manifest-check` (postoji li suite uopste) i
`release_gate.py`. Nova suite ulazi u kapiju time sto je upisana tamo.

Sveska: bez `--workbook` koristi se `tests/fixtures/otkup_test.xlsm`. Ako ga nema,
skript napravi PRAZNU .xlsm -- dovoljno za compile, ali NE i za suite (prazna
sveska nema tabele). Pravi fixture se pravi sa `tools/make_fixture.py`; sadrzi
samo sinteticke podatke, a suite koje diraju tabele ionako seju sebi svoje
(SVT-*, BIT-*, TST-*) u transakciji koja se uvek ponistava -- prava radna sveska
im NIJE potrebna. Sveska se uvek kopira u temp, original se ne dira.
Detalji: docs/EXCEL_TEST_HARNESS.md.

VERDIKT NIJE JEDAN. Ranije je run zavrsavao sa "REZULTAT: ZELENO", pa je
`COMPILE NEJASNO` nestajalo u toj jednoj reci. Sada svaka istina ima svoj red:

    STATIC    vba_check nad src-vba/            PASS | FAIL
    COMPILE   VBE Debug > Compile               PASS | UNKNOWN | FAIL
    SCHEMA    EnsureRuntimeSchema pre suite-ova PASS | FAIL | NOT_RUN
    BEHAVIOR  suite iz izabrane kapije          PASS | FAIL | NOT_RUN
    COUNTS    broj provera vs manifest          PASS | FAIL | NOT_RUN
    CLEANUP   tragovi test podataka posle run-a PASS | FAIL | NOT_RUN
    EXTERNAL  Google / MasterSync / SEF         PASS | FAIL | NOT_RUN

`BEHAVIOR PASS` uz `COMPILE UNKNOWN` je dozvoljen ishod -- ali mora tako i da
pise. Release kapija (`tools/release_gate.py`) je ta koja UNKNOWN ne prihvata.

Izlazni kod: 0 = sve kapije prosle, 2 = bar jedna pala.

Zasto ovaj skript NE zove `ImportAllVBA`:
  - `modVbaTools.ImportAllVBA` ima hardkodiran folder
    (C:\\Users\\Dusan\\Documents\\GitHub\\otkupapp-pwa\\src-vba\\) i zavrsava se
    `MsgBox`-om. Oboje je smrt za headless run. Import se ovde radi istom logikom
    preko COM-a, sa folderom izvedenim iz lokacije ovog fajla.
  - Poredak i pravila su ista kao u `ImportAllVBA`: .doccls se MERGE-uje u
    postojecu komponentu, .bas/.cls/.frm se uklone pa uvezu, .frm mora imati .frx
    par.

JEDNA RAZLIKA U ODNOSU NA `ImportAllVBA`: modVbaTools SE UVOZI.
    VBA procedura ne moze da ukloni modul iz kojeg se sama izvrsava, pa
    `ImportAllVBA` sebe preskace. Python nema taj problem. Ranije je i ovaj
    skript preskakao `modVbaTools` -- a `source_hash` ga UKLJUCUJE, pa je marker
    tvrdio da je testiran ceo `src-vba`, dok jedan modul nije bio ni u projektu.
    Hash i uvezeni kod moraju da opisuju istu stvar.
"""

from __future__ import annotations

import argparse
import io
import json
import os
import random
import shutil
import sys
import tempfile
import threading
import time
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import vba_check                      # noqa: E402  -- STATIC kapija
import vba_gate                       # noqa: E402  -- manifest, hash, marker

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_VBA = os.path.join(ROOT, "src-vba")
COMPILE_CONTROL_ID = 578             # VBE: Debug > Compile VBAProject
RUNNER_VERSION = 2                   # menja se kad se promeni oblik last_run.json

VBE_TYPE = {".bas": 1, ".cls": 2, ".frm": 3, ".doccls": 100}

DEFAULT_FIXTURE = os.path.join(ROOT, "tests", "fixtures", "otkup_test.xlsm")
GOLDEN_DIR = os.path.join(ROOT, "tests", "golden")
OUT_DIR = os.path.join(ROOT, "tests")
XL_OPENXML_MACRO = 52                # xlOpenXMLWorkbookMacroEnabled (.xlsm)

# Masinski citljiv izvestaj koji pisu suite kroz modTestRunner.TR_Report.
# Zivi PORED SVESKE (dakle u temp folderu), kao i last_run.txt koji pise modTest.
SUITE_RESULTS = "suite_results.txt"


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

    NE MESATI sa `suite_results.txt` (SUITE_RESULTS): ovaj fajl nosi IME PALOG
    TESTA, onaj nosi BROJ provera za COUNTS kapiju. Dva citaoca, dva zapisa.
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


def read_suite_counts(wbdir: str) -> dict:
    """`suite_results.txt` -> {suite: {pass, fail, notrun, seconds, message}}.

    Pise ga `modTestRunner.TR_Report`, jedan red po suite-u, dopunjavanjem
    (append) -- da run koji pukne na sedmom suite-u i dalje nosi brojeve prvih
    sest. Suite koji jos nije prikljucen na TR_Report se ovde ne pojavljuje i to
    se u ispisu vidi kao `nema broja`, ne kao nula.
    """
    path = os.path.join(wbdir, SUITE_RESULTS)
    out: dict = {}
    if not os.path.exists(path):
        return out
    with open(path, "r", encoding="utf-8", errors="replace") as fh:
        for raw in fh:
            parts = raw.rstrip("\r\n").split("|")
            if len(parts) < 6 or parts[0] != "SUITE":
                continue
            try:
                rec = {
                    "pass": int(parts[2] or 0),
                    "fail": int(parts[3] or 0),
                    "notrun": parts[4].strip() not in ("0", ""),
                    "seconds": float(parts[5] or 0),
                    "message": parts[6] if len(parts) > 6 else "",
                }
            except ValueError:
                continue
            # NOT_RUN je LEPLJIV. Suite ume da upise dva reda u istom prolazu:
            # TR_NotRun kaze "nije pokrenut", pa EH grana pozove ReportResults
            # koji kroz TR_Report upise 0/0 kao da je sve normalno. Da poslednji
            # red pobedjuje, "nije pokrenut" bi se izgubilo -- tacno ono stanje
            # zbog kojeg ova provera postoji.
            old = out.get(parts[1])
            if old and old.get("notrun") and not rec["notrun"]:
                rec["notrun"] = True
                rec["message"] = old.get("message") or rec["message"]
            out[parts[1]] = rec
    return out


def check_counts(manifest: dict, counts: dict, ran: list[str]) -> tuple[str, list[str]]:
    """Broj provera po suite-u vs `min_asserts` iz manifesta.

    Ovo je kapija protiv tihog pada pokrivenosti: ako Banka danas ima 189
    provera, sutrasnjih 120 je crveno dok neko ne spusti `min_asserts`
    eksplicitno, u commit-u koji se vidi.
    """
    lines: list[str] = []
    bad = False
    measured = 0

    for sid in ran:
        try:
            meta = vba_gate.suite_meta(manifest, sid)
        except vba_gate.ManifestError:
            continue
        want = int(meta.get("min_asserts") or 0)
        rec = counts.get(sid)

        if rec is None:
            if meta.get("reports_counts"):
                lines.append(f"{sid}: manifest kaze reports_counts=true, a suite nije "
                             f"upisao red u {SUITE_RESULTS} -- nije se pokrenuo do kraja")
                bad = True
            elif want:
                lines.append(f"{sid}: nema broja (jos nije na modTestRunner.TR_Report); "
                             f"ocekivano >= {want} ostaje nemereno")
            continue

        measured += 1
        got = rec["pass"] + rec["fail"]
        if rec["notrun"]:
            lines.append(f"{sid}: NOT_RUN -- {rec['message'] or 'bez razloga'}")
            bad = True
        elif rec["fail"] > 0:
            # Suite je SAMA prijavila pad. Ovo je jaci dokaz od "Run() nije
            # pukao": ako je Err.Raise u suite-u uklonjen ili pokvaren, VBA
            # zavrsi uredno, `raises: true` da status OK, a broj provera i dalje
            # zadovolji min_asserts -- i ceo run bi bio zelen iznad zapisanog
            # FAIL=1. Zato prijavljen pad ima prednost nad statusom iz Run().
            lines.append(f"{sid}: suite prijavila FAIL={rec['fail']} "
                         f"(PASS={rec['pass']}) -- pad je zapisan u {SUITE_RESULTS} "
                         f"i kad Run() nije pukao")
            bad = True
        elif want and got < want:
            lines.append(f"{sid}: {got} provera, a manifest trazi >= {want} "
                         f"-- pokrivenost je pala, spusti min_asserts svesno ili vrati testove")
            bad = True
        else:
            lines.append(f"{sid}: {got} provera (min {want})")

    # `bad` pre `measured`: suite koja je OBECALA broj (reports_counts) a nije ga
    # upisala je pad, ne "nema sta da se meri". Bez ovog redosleda bi se pad
    # tokom izvrsavanja migriranog suite-a prikazao kao NOT_RUN i propustio run.
    if bad:
        return "FAIL", lines
    if not measured:
        return "NOT_RUN", lines
    return "PASS", lines


def apply_reported_failures(suites: list[dict], counts: dict) -> None:
    """Status suite-a iz NJENOG izvestaja ima prednost nad ishodom `Run()`.

    `xl.Run(suite)` puca samo ako suite podigne gresku. Ako je taj `Err.Raise`
    uklonjen, pokvaren ili zaobidjen (rana grana, `On Error Resume Next` na
    pogresnom mestu), `Run()` zavrsi uredno i runner bi rekao `OK` -- iznad reda
    `SUITE|...|188|1|...` koji suite sama upisala. Ovde se to ispravlja.
    """
    for entry in suites:
        rec = counts.get(entry["name"])
        if not rec:
            continue
        if rec.get("notrun"):
            entry["status"] = "NOT_RUN"
            entry["error"] = rec.get("message") or "suite se nije pokrenuo"
        elif rec.get("fail", 0) > 0 and entry.get("status") == "OK":
            entry["status"] = "FAIL"
            entry["error"] = (f"suite prijavila FAIL={rec['fail']} PASS={rec['pass']} "
                              f"u {SUITE_RESULTS}, a Run() nije pukao")


# --- Watchdog za modalne dijaloge --------------------------------------------
#
# Svaki MsgBox u nevidljivom Excelu je trajno visenje: COM poziv ne puca, samo
# stoji. Watchdog nadgleda prozore klase #32770 (standardni Windows dialog) koji
# pripadaju NASOJ Excel instanci, procita tekst i klikne prvo dugme (za MsgBox
# je to podrazumevano: OK kod vbOKOnly, Yes kod vbYesNo). Time se visenje
# pretvara u zabelezen dogadjaj -- ukljucujuci i "Compile error: ..." dijalog,
# koji je inace jedini nacin da VBE prijavi neuspeh kompajliranja.
#
# Watchdog je sigurnosna mreza, ne model izvrsavanja: cilj je `dialogs: false` u
# manifestu za svaku suite (MsgBox iza modTestMode/IUserPrompt seam-a). Dok je
# `dialogs: true`, to je zapisan dug, ne osobina.

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


def report_orphan_components(wb, log: list[str], authoritative: bool) -> bool:
    """Moduli koji postoje u svesci a NEMA ih u src-vba/. Vraca True ako je nalaz.

    Import prepisuje samo ono sto repo ima; zaostao modul iz donora ostaje i
    izvrsava se. Ako nosi Public ime koje postoji i u svezem kodu, VBA to vidi kao
    "Ambiguous name" i odbija da pokrene makro iz njega -- sto izgleda kao
    "Cannot run the macro", a ne kao compile greska. Tako je TestLicense_All bio
    mrtav a `vba_check` uredno zelen: duplikat nije bio u repou nego u svesci.

    `authoritative` = radimo nad NASIM fixture-om (nema --workbook). Tada je
    ORPHAN tvrd pad: fixture pravi `make_fixture.py`, koji sav kod donora uklanja,
    pa zaostao modul znaci da je fixture pokvaren i rezultati nisu merodavni.
    Nad prosledjenom tudjom sveskom ostaje samo prijava.
    """
    STD, CLS, FRM = 1, 2, 3          # vbext_ct_StdModule / ClassModule / MSForm
    have = {os.path.splitext(n)[0].lower()
            for n in os.listdir(SRC_VBA)
            if os.path.splitext(n)[1] in (".bas", ".cls", ".frm")}

    orphans = []
    for vbc in wb.VBProject.VBComponents:
        try:
            if int(vbc.Type) not in (STD, CLS, FRM):
                continue                      # document moduli se merge-uju, ne brisu
            name = str(vbc.Name)
        except Exception:
            continue
        if name.lower() not in have:
            orphans.append(name)

    if not orphans:
        return False

    log.append(f"ORPHAN {len(orphans)} modul(a) u svesci bez para u src-vba/ "
               f"(mogu da daju 'Ambiguous name'): " + ", ".join(sorted(orphans)))
    if authoritative:
        log.append("ORPHAN je nad sopstvenim fixture-om TVRD PAD -- make_fixture.py "
                   "uklanja sav kod donora, pa zaostao modul znaci pokvaren fixture. "
                   "Napravi ga ponovo: python tools/make_fixture.py --donor <put> --force")
    return authoritative


def import_src_vba(wb, log: list[str]) -> list[str]:
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
    return missing

    # 2) standardni / klasni / forme
    for name in files:
        base, ext = os.path.splitext(name)
        if ext not in (".bas", ".cls", ".frm"):
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

    # Cist-Python deo runnera: izbor suite-ova i racunanje verdikta. Bez ovoga bi
    # se logika kapije mogla proveriti samo na Windows masini sa Excelom.
    return _self_test_pure()


def _self_test_pure() -> int:
    fails: list[str] = []

    def chk(cond: bool, label: str) -> None:
        if cond:
            print(f"  PASS  {label}")
        else:
            fails.append(label)
            print(f"  FAIL  {label}")

    man = vba_gate.load_manifest()

    dev = chosen_suites(parse_args([]), man)
    pr = chosen_suites(parse_args(["--gate", "pr"]), man)
    rel = chosen_suites(parse_args(["--gate", "release"]), man)
    chk(set(dev) <= set(pr) <= set(rel), "dev je podskup pr, pr je podskup release")
    chk(all(not vba_gate.suite_meta(man, s).get("external") for s in pr),
        "u 'pr' kapiji nema nijedne external suite")
    chk(any(vba_gate.suite_meta(man, s).get("external") for s in rel),
        "'release' kapija ukljucuje external suite")
    chk(all(vba_gate.suite_meta(man, s)["category"] != "health" for s in rel),
        "health provere nisu test kapija (RunProductionHealthCheck van gate-ova)")

    a = parse_args(["--shuffle", "--seed", "12345"])
    b = parse_args(["--shuffle", "--seed", "12345"])
    c = parse_args(["--shuffle", "--seed", "999"])
    chk(chosen_suites(a, man) == chosen_suites(b, man), "isti seed daje isti redosled")
    chk(sorted(chosen_suites(a, man)) == sorted(dev), "shuffle menja redosled, ne skup")
    chk(chosen_suites(c, man) != chosen_suites(a, man) or len(dev) < 3,
        "drugi seed daje drugi redosled")
    chk(len(chosen_suites(parse_args(["--repeat", "3"]), man)) == 3 * len(dev),
        "--repeat 3 vrti isti set tri puta")

    # COUNTS: tih pad broja provera mora biti crven. Ocekivani broj se CITA iz
    # manifesta -- zakucan broj bi se razisao sa njim cim neko doda proveru.
    want = int(vba_gate.suite_meta(man, "TestLicense_All")["min_asserts"])
    counts = {"TestLicense_All": {"pass": want, "fail": 0, "notrun": False,
                                  "seconds": 0.1, "message": ""}}
    st, _ = check_counts(man, counts, ["TestLicense_All"])
    chk(st == "PASS", f"COUNTS: {want} provera pri min_asserts={want} je PASS")

    counts["TestLicense_All"]["pass"] = want - 11
    st, lines = check_counts(man, counts, ["TestLicense_All"])
    chk(st == "FAIL" and any("min" in ln for ln in lines),
        f"COUNTS: pad sa {want} na {want - 11} provera je FAIL")

    counts["TestLicense_All"] = {"pass": 0, "fail": 0, "notrun": True,
                                 "seconds": 0.0, "message": "licenca ugasena"}
    st, _ = check_counts(man, counts, ["TestLicense_All"])
    chk(st == "FAIL", "COUNTS: NOT_RUN suite je FAIL, ne prolaz")

    st, _ = check_counts(man, {}, ["TestLicense_All"])
    chk(st == "FAIL", "COUNTS: reports_counts=true bez reda u izvestaju je FAIL")

    # NOT_RUN mora da prezivi kasniji "normalan" red iz iste suite.
    import tempfile as _tf
    with _tf.TemporaryDirectory() as d:
        with open(os.path.join(d, SUITE_RESULTS), "w", encoding="utf-8") as fh:
            fh.write("SUITE|RunBankaImportTestSuite|0|0|1|0.1|operater odustao\n")
            fh.write("SUITE|RunBankaImportTestSuite|0|1|0|0.2|\n")
        merged = read_suite_counts(d)
        chk(merged["RunBankaImportTestSuite"]["notrun"] is True,
            "NOT_RUN prezivljava kasniji red iste suite (EH grana ga ne prepise)")
        chk("odustao" in merged["RunBankaImportTestSuite"]["message"],
            "razlog za NOT_RUN se cuva")

    # --- P0: suite prijavila FAIL, a Run() nije pukao ------------------------
    # Ako je Err.Raise u suite-u uklonjen ili zaobidjen, VBA zavrsi uredno,
    # manifest `raises: true` da status OK, a 188+1=189 zadovolji min_asserts.
    # Bez ove provere je ceo run bio zelen iznad zapisanog FAIL=1.
    counts = {"RunBankaImportTestSuite": {"pass": 188, "fail": 1, "notrun": False,
                                          "seconds": 9.0, "message": ""}}
    st, lines = check_counts(man, counts, ["RunBankaImportTestSuite"])
    chk(st == "FAIL" and any("FAIL=1" in ln for ln in lines),
        "COUNTS: 188 pass + 1 fail je FAIL i kad je zbir >= min_asserts")

    rows = [{"name": "RunBankaImportTestSuite", "status": "OK", "gate": True}]
    apply_reported_failures(rows, counts)
    chk(rows[0]["status"] == "FAIL",
        "prijavljen FAIL pregazi status OK iz Run()")

    v = verdicts({"static": "PASS", "compile": {"ok": True}, "schema": "OK",
                  "suites": rows, "counts_status": st, "mode": "suites"},
                 enforce_cleanup=False)
    chk(rc_from(v) == 2, "188 pass + 1 fail bez VBA raise -> ceo run pada (exit 2)")

    rows = [{"name": "RunPaleteTestSuite", "status": "OK", "gate": True}]
    apply_reported_failures(rows, {"RunPaleteTestSuite": {
        "pass": 0, "fail": 0, "notrun": True, "seconds": 0.1,
        "message": "paletiranje iskljuceno"}})
    chk(rows[0]["status"] == "NOT_RUN", "prijavljen NOT_RUN pregazi status OK iz Run()")

    v = verdicts({"static": "PASS", "compile": {"ok": True}, "schema": "OK",
                  "suites": rows, "counts_status": "FAIL", "mode": "suites"},
                 enforce_cleanup=False)
    chk(v["BEHAVIOR"] == "FAIL" and rc_from(v) == 2,
        "NOT_RUN red daje BEHAVIOR FAIL, ne 'nista se nije desilo'")

    # --- rezultat PO SUITE-U (main #227): jedan globalni slot je sakrivao crveno
    # Ovo je ovde zato sto je regresija stvarno nastala: pri rebase-u na main je
    # ovaj port nestao, a nijedna provera ga nije cuvala. Sada ga cuva.
    with _tf.TemporaryDirectory() as d:
        with open(os.path.join(d, "last_run.txt"), "w", encoding="utf-8") as fh:
            fh.write("TESTS=17 FAIL=1\nFAIL T_ClearForm_Ugovor -- datum ostaje\n")
        with open(os.path.join(d, "last_run_banka.txt"), "w", encoding="utf-8") as fh:
            fh.write("TESTS=189 FAIL=0\n")
        rep = {}
        rc1 = _read_test_results(d, rep, "last_run.txt", "RunAllTests")
        rc2 = _read_test_results(d, rep, "last_run_banka.txt", "RunBankaImportTestSuite")
        chk(rc1 == 2 and rc2 == 0, "rezultat se cita iz fajla koji manifest imenuje")
        chk(set(rep.get("suite_results", {})) == {"RunAllTests", "RunBankaImportTestSuite"},
            "dve suite -> dva slota, druga NE prepisuje prvu")
        chk(rep["suite_results"]["RunAllTests"]["failed"] == 1
            and rep["suite_results"]["RunBankaImportTestSuite"]["failed"] == 0,
            "svaki slot nosi svoj broj palih")

        rep.update({"static": "PASS", "compile": {"ok": True}, "schema": "OK",
                    "mode": "suites", "counts_status": "PASS", "import": [], "dialogs": [],
                    "suites": [{"name": "RunAllTests", "status": "FAIL", "gate": True,
                                "seconds": 1.0},
                               {"name": "RunBankaImportTestSuite", "status": "OK",
                                "gate": True, "seconds": 2.0}]})
        _write_report(rep, verdicts(rep, False), 2, "t", out_dir=d)
        ispis = io.open(os.path.join(d, "t.txt"), encoding="utf-8").read()
        chk("TESTS   RunAllTests: 17 ukupno, 1 palo" in ispis,
            "ispis imenuje suite uz njen rezultat")
        chk("T_ClearForm_Ugovor" in ispis, "ime palog testa prezivi do ispisa")
        chk("TESTS   RunBankaImportTestSuite: 189 ukupno, 0 palo" in ispis,
            "zelena suite ne gura crvenu iz ispisa")

    # RunAllTests je jedini suite cija se ocekivana brojka moze izmeriti bez
    # Excela: modTest pise TESTS=<broj RunOne poziva>. Kad main doda testove, a
    # manifest ostane na staroj brojci, COUNTS kapija za taj suite tiho oslabi
    # (163 moze da padne na 18 i da i dalje "prodje"). Zato se poredi ovde.
    runone = 0
    mt = os.path.join(SRC_VBA, "modTest.bas")
    if os.path.exists(mt):
        with open(mt, "r", encoding="ascii", errors="replace") as fh:
            runone = sum(1 for ln in fh if ln.startswith("    RunOne "))
    want_all = int(vba_gate.suite_meta(man, "RunAllTests")["min_asserts"])
    chk(runone == 0 or want_all == runone,
        f"min_asserts RunAllTests ({want_all}) prati broj RunOne poziva ({runone})")

    manifest_rf = {x["id"]: x["result_file"] for x in man["suites"] if x["result_file"]}
    chk(manifest_rf.get("RunBankaImportTestSuite") == "last_run_banka.txt",
        "manifest nosi ime banka rezultat-fajla (detalj pada ne prezivi COM)")

    # --- P0: --no-import ne sme da vodi u last-green -------------------------
    def marks_green(argv: list[str]) -> bool:
        a = parse_args(argv)
        return (not a.suite and not a.category and not a.compile_only
                and not a.no_import and not a.no_mark_green)

    chk(marks_green(["--gate", "pr"]), "pun 'pr' run upisuje marker")
    chk(not marks_green(["--gate", "pr", "--no-import"]),
        "--no-import NIKAD ne upisuje marker (sveska nosi tudji kod)")
    chk(not marks_green(["--suite", "RunAllTests"]), "--suite ne upisuje marker")
    chk(not marks_green(["--category", "unit"]), "--category ne upisuje marker")
    chk(not marks_green(["--compile-only"]), "--compile-only ne upisuje marker")

    # --- prazan izbor i rezimi ----------------------------------------------
    try:
        chosen_suites(parse_args(["--category", "nepostojeca"]), man)
        chk(False, "prazan izbor suite-ova je tvrda greska")
    except SystemExit:
        chk(True, "prazan izbor suite-ova je tvrda greska")

    v = verdicts({"static": "PASS", "compile": {"ok": True}, "schema": "OK",
                  "suites": [], "counts_status": "NOT_RUN", "mode": "suites"},
                 enforce_cleanup=False)
    chk(rc_from(v) == 2, "run bez ijedne izvrsene suite ne moze da bude uspeh")

    v = verdicts({"static": "PASS", "compile": {"ok": True}, "schema": "OK",
                  "suites": [{"name": "H", "status": "BLIND", "gate": False}],
                  "counts_status": "PASS", "mode": "suites"}, enforce_cleanup=False)
    chk(rc_from(v) == 2, "run u kome su SVE suite blind ne daje verdikt -> exit 2")

    v = verdicts({"static": "PASS", "compile": {"ok": None}, "schema": None,
                  "suites": [], "counts_status": "NOT_RUN", "mode": "compile-only"},
                 enforce_cleanup=False)
    chk(rc_from(v) == 2, "--compile-only sa NEJASNO i dalje pada")
    v = verdicts({"static": "PASS", "compile": {"ok": True}, "schema": None,
                  "suites": [], "counts_status": "NOT_RUN", "mode": "compile-only"},
                 enforce_cleanup=False)
    chk(rc_from(v) == 0, "--compile-only sa COMPILE OK prolazi")

    # Verdikt: UNKNOWN compile ne obara run kad su suite prosle, ali se VIDI.
    v = verdicts({"static": "PASS", "compile": {"ok": None}, "schema": "OK",
                  "suites": [{"name": "X", "status": "OK", "gate": True}],
                  "counts_status": "PASS"}, enforce_cleanup=False)
    chk(v["COMPILE"] == "UNKNOWN" and v["BEHAVIOR"] == "PASS" and rc_from(v) == 0,
        "COMPILE UNKNOWN uz zelene suite: run prolazi, ali verdikt kaze UNKNOWN")

    v = verdicts({"static": "PASS", "compile": {"ok": False}, "schema": "OK",
                  "suites": [{"name": "X", "status": "OK", "gate": True}],
                  "counts_status": "PASS"}, enforce_cleanup=False)
    chk(rc_from(v) == 2, "COMPILE FAIL obara run i kad su suite zelene")

    v = verdicts({"static": "PASS", "compile": {"ok": True}, "schema": "FAIL nema rutine",
                  "suites": [{"name": "X", "status": "OK", "gate": True}],
                  "counts_status": "PASS"}, enforce_cleanup=False)
    chk(rc_from(v) == 2, "SCHEMA FAIL obara run i kad su suite zelene")

    v = verdicts({"static": "PASS", "compile": {"ok": True}, "schema": "OK",
                  "suites": [{"name": "X", "status": "BLIND", "gate": False}],
                  "counts_status": "NOT_RUN"}, enforce_cleanup=False)
    chk(v["BEHAVIOR"] == "NOT_RUN",
        "sama BLIND suite ne daje BEHAVIOR PASS ('proslo bez greske' nije dokaz)")

    v = verdicts({"static": "FAIL", "compile": None, "schema": None,
                  "suites": [], "counts_status": "NOT_RUN"}, enforce_cleanup=False)
    chk(rc_from(v) == 2, "STATIC FAIL obara run pre Excela")

    v = verdicts({"static": "PASS", "compile": {"ok": True}, "schema": "OK",
                  "suites": [{"name": "X", "status": "OK", "gate": True}],
                  "counts_status": "PASS", "cleanup": {"status": "FAIL", "detail": ["x"]}},
                 enforce_cleanup=False)
    chk(rc_from(v) == 0 and v["CLEANUP"] == "FAIL",
        "CLEANUP FAIL se prijavljuje, ali ne obara run bez --enforce-cleanup")
    v = verdicts({"static": "PASS", "compile": {"ok": True}, "schema": "OK",
                  "suites": [{"name": "X", "status": "OK", "gate": True}],
                  "counts_status": "PASS", "cleanup": {"status": "FAIL", "detail": ["x"]}},
                 enforce_cleanup=True)
    chk(rc_from(v) == 2, "--enforce-cleanup pretvara CLEANUP FAIL u pad")

    chk(not cleanup_enforced(parse_args([])), "CLEANUP je prijava u 'dev' kapiji")
    chk(cleanup_enforced(parse_args(["--gate", "pr"])),
        "CLEANUP je BLOKIRAJUCI u 'pr' kapiji")
    chk(cleanup_enforced(parse_args(["--gate", "release"])),
        "CLEANUP je BLOKIRAJUCI u 'release' kapiji")
    chk(not cleanup_enforced(parse_args(["--gate", "pr", "--no-enforce-cleanup"])),
        "--no-enforce-cleanup je izlaz dok detektor nije dokazan")

    print()
    if fails:
        print(f"self-test (cist Python deo): {len(fails)} nalaza", file=sys.stderr)
        return 2
    print("self-test (cist Python deo): cisto")
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


# --- Leak detektor -------------------------------------------------------------

def check_leaks(wb, manifest: dict) -> dict:
    """Da li je posle suite-ova ostao trag test podataka u svesci.

    "Transakcija bi trebalo da rollback-uje" nije dokaz da jeste. Ovo trazi
    zaostale redove sa test prefiksima (`leak_prefixes` iz manifesta) i listove
    koje su suite napravile (`_Test*`).

    NAMERNO NE OBARA RUN podrazumevano: provera nikad nije pokrenuta nad pravim
    Excelom, pa bi tvrd crven verdikt bio tvrdnja koju niko nije proverio.
    Ukljucuje se sa `--enforce-cleanup`, posle jednog Windows run-a koji pokaze
    oba smera (cist run = PASS, namerno ostavljen TST- red = FAIL).
    """
    prefixes = tuple(manifest.get("leak_prefixes") or [])
    detail: list[str] = []
    if not prefixes:
        return {"status": "NOT_RUN", "detail": ["manifest nema leak_prefixes"]}

    try:
        for ws in wb.Worksheets:
            try:
                sheet_name = str(ws.Name)
            except Exception:               # noqa: BLE001
                continue
            if sheet_name.lower().startswith("_test"):
                detail.append(f"list '{sheet_name}' je ostao posle run-a")
            for lo in ws.ListObjects:
                try:
                    if int(lo.ListRows.Count) == 0:
                        continue
                    values = lo.DataBodyRange.Value2
                except Exception:           # noqa: BLE001
                    continue
                hits = _count_prefix_hits(values, prefixes)
                if hits:
                    detail.append(f"{lo.Name}: {hits} red(ova) sa test prefiksom")
    except Exception as exc:                # noqa: BLE001
        return {"status": "NOT_RUN", "detail": [f"provera nije izvrsena: {exc}"]}

    return {"status": "FAIL" if detail else "PASS", "detail": detail}


def _count_prefix_hits(values, prefixes: tuple) -> int:
    """Broj REDOVA u kojima bar jedna celija pocinje test prefiksom."""
    if not isinstance(values, tuple):
        values = ((values,),)
    hits = 0
    for row in values:
        cells = row if isinstance(row, tuple) else (row,)
        for cell in cells:
            if isinstance(cell, str) and cell.startswith(prefixes):
                hits += 1
                break
    return hits


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
    ap = argparse.ArgumentParser(description="Headless static + import + compile + VBA suite")
    ap.add_argument("--workbook", help="putanja do .xlsm (podrazumevano tests/fixtures/otkup_test.xlsm)")
    ap.add_argument("--gate", choices=("dev", "pr", "release"), default="dev",
                    help="koju kapiju vrtimo (podrazumevano dev = ono sto pusta Stop hook)")
    ap.add_argument("--category", action="append", default=[],
                    help="samo suite date kategorije (unit/integration/contract/e2e/external/health)")
    ap.add_argument("--compile-only", action="store_true", help="stani posle compile-a")
    ap.add_argument("--no-import", action="store_true", help="ne uvozi src-vba (testiraj zatecen kod)")
    ap.add_argument("--suite", action="append", default=[], help="pokreni bas ovu suite (moze vise puta)")
    ap.add_argument("--all", action="store_true", help="pokreni sve suite iz kataloga")
    ap.add_argument("--shuffle", action="store_true",
                    help="nasumican redosled suite-ova (hvata test koji prolazi samo zato "
                         "sto ga je prethodni pripremio)")
    ap.add_argument("--seed", type=int, default=0, help="seed za --shuffle (isti seed = isti redosled)")
    ap.add_argument("--repeat", type=int, default=1,
                    help="koliko puta ponoviti izabran set u ISTOM Excel procesu "
                         "(hvata stale cache, static Boolean, license latch)")
    ap.add_argument("--keep", action="store_true", help="ne brisi temp kopiju sveske")
    ap.add_argument("--enforce-cleanup", action="store_true",
                    help="CLEANUP nalaz obara run (podrazumevano za --gate pr/release)")
    ap.add_argument("--no-enforce-cleanup", action="store_true",
                    help="CLEANUP samo prijavi, ne obaraj run (izlaz za slucaj laznog nalaza)")
    ap.add_argument("--no-mark-green", action="store_true",
                    help="ne upisuj last-green marker ni kad je sve zeleno")
    ap.add_argument("--timeout-dialog", type=float, default=0.4, help="interval watchdog-a u sekundama")
    ap.add_argument("--timeout", type=float, default=0.0,
                    help="tvrdi prekid celog run-a u sekundama; 0 = suma timeout_s iz manifesta + 120")
    ap.add_argument("--ignore-fixture-sig", action="store_true",
                    help="ne staj na ustajao fixture (svesno testiranje nad starim podacima)")
    ap.add_argument("--report-name", default="last_run",
                    help="osnova imena izvestaja u tests/ (podrazumevano last_run); "
                         "release kapija vrti dva runa pa im treba odvojen zapis")
    ap.add_argument("--self-test", action="store_true",
                    help="provere koje ne traze Excel (radi i na Linux/macOS)")
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

    try:
        import make_fixture
    except ImportError:
        return 0

    zapisan = make_fixture.read_sig(fixture)
    tekuci = make_fixture.signature()
    if zapisan == tekuci:
        return 0

    komanda = 'python tools\\make_fixture.py --donor "<put do .xlsm>" --force'
    if not zapisan:
        # FAIL-CLOSED. Sveska od generatora pre ovog sistema NE moze da se
        # proveri, a to je bas prvi run na svakoj zatecenoj masini -- tacno
        # onaj u kome se incident i desio. Upozorenje bi ga propustilo jednom
        # po masini, sto je isto kao da provere nema.
        #
        # Prazna auto-sveska ovde ne stize: nju pravi grana IZNAD ovog poziva.
        print(f"FIXTURE BEZ POTPISA: {fixture}", file=sys.stderr)
        print("  Sveska je od generatora pre uvodjenja potpisa, pa se ne moze",
              file=sys.stderr)
        print("  utvrditi da li su podaci svezi. Regenerisi je jednom:", file=sys.stderr)
        print(f"\n  {komanda}\n", file=sys.stderr)
        if ignorisi:
            print("  --ignore-fixture-sig: nastavljam ipak.", file=sys.stderr)
            return 0
        return 2

    print(f"USTAJAO FIXTURE: {fixture}", file=sys.stderr)
    print(f"  posejan potpisom {zapisan}, generator sada trazi {tekuci}", file=sys.stderr)
    print("  Fixture je gitignored -- `git checkout` ga NE menja, pa je na disku",
          file=sys.stderr)
    print("  ostala sveska prethodne grane. Testovi bi pali na PODACIMA.", file=sys.stderr)
    print(f"\n  {komanda}\n", file=sys.stderr)
    if ignorisi:
        print("  --ignore-fixture-sig: nastavljam ipak.", file=sys.stderr)
        return 0
    return 2

def cleanup_enforced(args: argparse.Namespace) -> bool:
    """CLEANUP je BLOKIRAJUCI u `pr` i `release`, prijava u `dev`.

    Razlog za razliku: dev kapija se vrti stalno i lazan nalaz bi je ucinio
    neupotrebljivom; `pr`/`release` se vrte namerno i tamo je zaostao TST- red
    ozbiljan -- svaka sledeca suite, ukljucujuci external ugovore, radila bi nad
    prljavom sveskom. `--no-enforce-cleanup` postoji kao izlaz dok se detektor ne
    dokaze u oba smera nad pravim Excelom.
    """
    if args.no_enforce_cleanup:
        return False
    if args.enforce_cleanup:
        return True
    return args.gate in ("pr", "release") and not args.suite and not args.category


def chosen_suites(args: argparse.Namespace, manifest: dict) -> list[str]:
    known = [s["id"] for s in manifest["suites"]]

    if args.suite:
        unknown = [s for s in args.suite if s not in known]
        if unknown:
            print(f"Nepoznata suite: {', '.join(unknown)}", file=sys.stderr)
            print(f"Poznate: {', '.join(sorted(known))}", file=sys.stderr)
            raise SystemExit(2)
        picked = list(args.suite)
    elif args.all:
        picked = list(known)
    elif args.category:
        want = {c.lower() for c in args.category}
        picked = [s["id"] for s in manifest["suites"]
                  if str(s.get("category", "")).lower() in want]
    else:
        picked = vba_gate.suites_for_gate(manifest, args.gate)

    if not picked:
        # Prazan izbor bi ranije dao run bez ijedne suite i exit 0 -- "nista nije
        # palo". Filter koji ne pogadja nijednu suite je greska u pozivu, ne prolaz.
        print("Izbor ne pogadja nijednu suite (gate="
              f"{args.gate}, category={args.category or '-'}).", file=sys.stderr)
        raise SystemExit(2)

    if args.shuffle:
        # Seed ide u ispis i u last_run.json: pad koji se vidi samo u jednom
        # redosledu mora da se moze ponoviti tacno tim seed-om.
        random.Random(args.seed).shuffle(picked)

    return picked * max(1, int(args.repeat))


def verdicts(report: dict, enforce_cleanup: bool) -> dict:
    """Sedam nezavisnih istina umesto jednog "ZELENO".

    Pravila koja nisu ocigledna:
      - `COMPILE UNKNOWN` ne obara run kad su suite isle: da bi se suite uopste
        pokrenula, VBA je morao da kompajlira nju i sve sto referencira. Ali
        UNKNOWN ostaje ispisan, i release kapija ga ne prihvata.
      - Sama BLIND suite ne daje `BEHAVIOR PASS`: "proslo bez greske" nije isto
        sto i "sve provere prosle".
      - `CLEANUP` je podrazumevano samo prijava (v. check_leaks).
    """
    v: dict = {}

    v["STATIC"] = report.get("static") or "NOT_RUN"

    c = report.get("compile")
    v["COMPILE"] = "NOT_RUN" if not c else {
        True: "PASS", False: "FAIL", None: "UNKNOWN"}[c.get("ok")]

    schema = report.get("schema")
    v["SCHEMA"] = "NOT_RUN" if schema is None else ("PASS" if schema == "OK" else "FAIL")

    suites = report.get("suites") or []
    local = [s for s in suites if not s.get("external")]
    ext = [s for s in suites if s.get("external")]

    def _status(rows: list[dict]) -> str:
        if not rows:
            return "NOT_RUN"
        # NOT_RUN red je PAD, ne odsustvo rezultata: suite koja se nije pokrenula
        # ostavlja rupu u pokrivenosti isto kao suite koja je pala.
        if any(r["status"] in ("FAIL", "TIMEOUT", "NOT_RUN") for r in rows):
            return "FAIL"
        if any(r["status"] == "OK" for r in rows):
            return "PASS"
        return "NOT_RUN"        # samo BLIND redovi -- prosle bez verdikta

    v["BEHAVIOR"] = _status(local)
    v["EXTERNAL"] = _status(ext)
    v["COUNTS"] = report.get("counts_status") or "NOT_RUN"
    v["CLEANUP"] = (report.get("cleanup") or {}).get("status", "NOT_RUN")
    v["_cleanup_enforced"] = enforce_cleanup
    v["_mode"] = report.get("mode") or "suites"
    return v


def rc_from(v: dict) -> int:
    """Uspeh trazi POZITIVAN dokaz, ne izostanak greske.

    Ranije je `NOT_RUN` prolazio kroz sve provere (nije `FAIL`), pa je run u kome
    se nista nije izvrsilo zavrsavao sa exit 0. "Nije pokrenuto" nije "proslo" --
    to pravilo vazi za suite, pa mora da vazi i za ceo run.
    """
    if v["STATIC"] != "PASS":
        return 2
    if v["COMPILE"] == "FAIL":
        return 2

    mode = v.get("_mode") or "suites"
    if mode == "compile-only":
        # Bez suite-ova je probe jedini izvor istine, pa NEJASNO i dalje pada.
        return 0 if v["COMPILE"] == "PASS" else 2
    if mode == "static-only":
        # Linux/macOS ili prekid pre Excela: nista se nije izvrsilo.
        return 2

    if v["SCHEMA"] != "PASS":
        return 2
    if v["BEHAVIOR"] == "FAIL" or v["EXTERNAL"] == "FAIL":
        return 2
    # Bar jedan skup mora da da PASS. Run u kome su sve suite BLIND (ili ih nema)
    # nije dao nijedan verdikt i ne sme da izadje kao uspeh.
    if v["BEHAVIOR"] != "PASS" and v["EXTERNAL"] != "PASS":
        return 2
    if v["COUNTS"] != "PASS":
        return 2
    if v.get("_cleanup_enforced") and v["CLEANUP"] == "FAIL":
        return 2
    return 0


def main(argv: list[str]) -> int:
    args = parse_args(argv)

    if args.self_test:
        return self_test()

    try:
        manifest = vba_gate.load_manifest()
    except vba_gate.ManifestError as exc:
        print(f"MANIFEST {exc}", file=sys.stderr)
        return 2

    fixture = args.workbook or DEFAULT_FIXTURE
    blank_fixture = False
    report: dict = {
        "runner_version": RUNNER_VERSION,
        "gate": args.gate,
        "shuffle": bool(args.shuffle),
        "seed": args.seed if args.shuffle else None,
        "repeat": args.repeat,
        "provenance": vba_gate.provenance(fixture),
        "workbook": fixture,
        # Da li je src-vba iz repoa STVARNO uvezen u svesku nad kojom su suite
        # isle. `False` znaci da source_hash u provenance-u opisuje repo, a ne
        # ono sto je izvrseno -- zato takav run nikad ne pise last-green marker.
        "source_imported": not args.no_import,
        "static": None,
        "import": [],
        "compile": None,
        "suites": [],
        "dialogs": [],
    }

    # --- STATIC kapija ---------------------------------------------------
    # Ide PRE Excela i vrti se svuda. Nema smisla dizati Excel nad izvorom koji
    # ima ne-ASCII bajt ili duplu Public definiciju -- to bi bio pad u compile-u,
    # samo trideset sekundi kasnije i uz nejasniju poruku.
    report["mode"] = "compile-only" if args.compile_only else "suites"

    static_rc = vba_check.main(["--hook"])
    report["static"] = "PASS" if static_rc == 0 else "FAIL"

    manifest_findings = vba_gate.manifest_check(manifest)
    if manifest_findings:
        report["static"] = "FAIL"
        report["manifest"] = manifest_findings

    if report["static"] == "FAIL":
        report["mode"] = "static-only"
        v = verdicts(report, cleanup_enforced(args))
        _write_report(report, v, 2, args.report_name)
        return 2

    if os.name != "nt":
        report["mode"] = "static-only"
        report["behavior_skip"] = ("Excel COM postoji samo na Windows-u -- testovi ponasanja "
                                   "se OVDE NE IZVRSAVAJU. To nije 'proslo'.")
        v = verdicts(report, cleanup_enforced(args))
        _write_report(report, v, 2, args.report_name)
        return 2

    import pythoncom
    import win32api
    import win32com.client as win32
    import win32con
    import win32process

    if not os.path.exists(fixture):
        if args.workbook:
            print(f"Sveska ne postoji: {fixture}", file=sys.stderr)
            return 2
        # Podrazumevani fixture se pravi sam -- prazna sveska je dovoljna za
        # compile (vidi create_blank_fixture). Ovo se desi tacno jednom.
        print(f"Fixture ne postoji, pravim praznu svesku: {fixture}")
        blank_fixture = True
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
        report["provenance"]["fixture_hash"] = vba_gate.file_hash(fixture)
    elif not args.compile_only:
        # Compile ne cita podatke, pa ga ustajao fixture ne dotice.
        rc_sig = proveri_fixture_potpis(fixture, bool(args.workbook),
                                        args.ignore_fixture_sig)
        if rc_sig:
            report["fatal"] = ("Ustajao ili nepotpisan fixture -- v. poruku iznad. "
                               "Rezultati nad starim podacima nisu merodavni.")
            v = verdicts(report, cleanup_enforced(args))
            _write_report(report, v, 2, args.report_name)
            return 2

    suites = chosen_suites(args, manifest)
    metas = {s: vba_gate.suite_meta(manifest, s) for s in set(suites)}

    tmp = tempfile.mkdtemp(prefix="vbatest_")
    wbpath = os.path.join(tmp, os.path.basename(fixture))
    shutil.copy2(fixture, wbpath)

    # Golden fajlovi idu u temp pre rana i vracaju se posle: modTest ih trazi
    # pored sveske, a nov golden mora da zavrsi u repou (tests/golden) na
    # ljudski pregled -- inace bi ga sledeci run tiho napravio ponovo.
    _copy_golden(GOLDEN_DIR, os.path.join(tmp, "golden"))

    hard_total = args.timeout or (sum(int(metas[s].get("timeout_s") or 300)
                                      for s in suites) + 120)

    rc = 2
    xl = None
    pid = None
    watchdog = None
    killer = None
    hard_stop: dict = {"fired": False, "suite": None}

    pythoncom.CoInitialize()
    try:
        xl = win32.DispatchEx("Excel.Application")
        pid = win32process.GetWindowThreadProcessId(xl.Hwnd)[1]

        watchdog = DialogWatchdog(pid, poll=args.timeout_dialog)
        watchdog.start()

        # Tvrdi prekid: ako Excel iz bilo kog razloga prestane da odgovara (break
        # mode, dijalog koji watchdog ne prepozna), COM poziv ne puca -- samo
        # stoji. Bez ovoga run visi dok ga neko ne ubije rukom. Ceo run ima svoj
        # tajmer; svaka suite dobija JOS jedan, svoj (v. petlju nize) -- inace
        # BFP koji odjednom traje 40s umesto 5s ne bi bio vidljiv dok ne pojede
        # ceo budzet run-a.
        killer = threading.Timer(hard_total, _terminate_pid, args=(pid, hard_stop))
        killer.daemon = True
        killer.start()

        xl.Visible = False
        xl.DisplayAlerts = False
        xl.AutomationSecurity = 1      # msoAutomationSecurityLow -- makroi bez pitanja
        xl.EnableEvents = False        # KLJUCNO: Workbook_Open (StartApp/self-update) se ne pokrece

        wb = xl.Workbooks.Open(wbpath, UpdateLinks=0)

        orphan_fatal = False
        source_gap = ""
        if not args.no_import:
            missing_doccls = import_src_vba(wb, report["import"])
            orphan_fatal = report_orphan_components(wb, report["import"],
                                                    authoritative=not args.workbook)
            # Nedostajuci .doccls znaci da kod tih listova NIJE ni uvezen, a
            # `source_hash` ih uracunava -- marker bi tvrdio da je testiran ceo
            # src-vba. Nad NASIM fixture-om (donor ima sve listove) to je pad.
            # Nad praznom automatski napravljenom sveskom nije: tamo nedostaje
            # svih 40+ listova po dizajnu, ali takav run sme samo `--compile-only`
            # i njegov rezultat se NE zove pun compile izvora.
            if missing_doccls and not args.workbook:
                if blank_fixture:
                    source_gap = (f"{len(missing_doccls)} .doccls nije uvezeno "
                                  f"(prazna sveska) -- ovo NIJE compile celog izvora")
                    report["import"].append("SOURCE  " + source_gap)
                else:
                    report["fatal"] = (
                        f"{len(missing_doccls)} .doccls modula nije uvezeno u fixture "
                        f"({', '.join(missing_doccls[:6])}) -- kod tih listova nije "
                        f"testiran, a source_hash ga uracunava. Napravi fixture "
                        f"ponovo: python tools/make_fixture.py --donor <put> --force")
        report["source_gap"] = source_gap

        # Compile se radi UVEK -- i uz --compile-only i pre suite-ova. To je
        # najjeftiniji gate koji hvata najcescu klasu kvara posle Edit/Write nad
        # src-vba (duple definicije, deklaracija posle prve procedure, ime koje
        # se poklapa sa rezervisanom reci).
        report["compile"] = compile_project(xl, wb, watchdog)

        compile_ok = report["compile"].get("ok")

        if "fatal" in report:
            pass                        # nedostajuci .doccls -- v. import granu
        elif orphan_fatal:
            report["fatal"] = ("ORPHAN moduli u sopstvenom fixture-u -- rezultati nisu "
                               "merodavni (v. IMPORT redove iznad).")
        elif compile_ok is False:
            pass                        # eksplicitan compile dijalog nosi verdikt
        elif args.compile_only:
            pass                        # bez suite-ova je probe jedini izvor istine
        else:
            # NEJASNO vise ne obara run kad suite-ovi mogu da daju pravi odgovor.
            # Da bi se suite uopste pokrenula, VBA mora da kompajlira nju i sve
            # sto ona referencira -- a to je bas kod pod testom. Verdikt probe se
            # i dalje racuna i ispisuje nepromenjen, samo vise nije prepreka.
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

            for idx, suite in enumerate(suites, start=1):
                meta = metas[suite]
                entry = {"name": suite, "gate": bool(meta.get("raises")),
                         "category": meta.get("category"),
                         "external": bool(meta.get("external")), "pass": idx}
                per_suite = threading.Timer(int(meta.get("timeout_s") or 300),
                                            _terminate_pid, args=(pid, hard_stop))
                per_suite.daemon = True
                hard_stop["suite"] = suite
                per_suite.start()
                t0 = time.time()
                try:
                    xl.Run(suite)
                except Exception as exc:    # noqa: BLE001
                    entry["status"] = "TIMEOUT" if hard_stop["fired"] else "FAIL"
                    entry["error"] = str(exc)
                    # Suite koja pad prijavljuje GRESKOM (banka, i svaka na
                    # modTestRunner-u) svejedno je upisala detalje pored sveske.
                    # Bez ovoga od celog pada ostane samo "Exception occurred" --
                    # pa se ne vidi KOJA provera je pala, nego samo da jeste.
                    if meta.get("result_file"):
                        _read_test_results(tmp, report, meta["result_file"], suite)
                else:
                    if meta.get("result_file"):
                        # Verdikt iz fajla suite, ne iz "Run() nije pukao".
                        entry["status"] = "OK" if _read_test_results(
                            tmp, report, meta["result_file"], suite) == 0 else "FAIL"
                    else:
                        entry["status"] = "OK" if meta.get("raises") else "BLIND"
                finally:
                    per_suite.cancel()
                entry["seconds"] = round(time.time() - t0, 1)
                report["suites"].append(entry)
                if hard_stop["fired"]:
                    report["fatal"] = (f"TIMEOUT na suite '{suite}' "
                                       f"({meta.get('timeout_s')}s) -- Excel proces je ubijen.")
                    break

            counts = read_suite_counts(tmp)
            report["counts"] = counts
            ran = [s["name"] for s in report["suites"]]
            report["counts_status"], report["counts_detail"] = check_counts(manifest, counts, ran)
            # Status suite-a se PREPRAVLJA iz njenog sopstvenog izvestaja: "Run()
            # nije pukao" je slabiji dokaz od "suite je zapisala FAIL=1".
            apply_reported_failures(report["suites"], counts)

            if not hard_stop["fired"]:
                report["cleanup"] = check_leaks(wb, manifest)

    except Exception as exc:                # noqa: BLE001
        report["fatal"] = str(exc)
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
        if hard_stop["fired"] and "fatal" not in report:
            report["fatal"] = (f"Excel nije odgovarao -- proces je ubijen "
                               f"(suite '{hard_stop['suite']}'). Najcesci uzrok: VBE je u "
                               f"break mode-u (vidi docs/EXCEL_TEST_HARNESS.md).")
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

    v = verdicts(report, cleanup_enforced(args))
    rc = 2 if "fatal" in report else rc_from(v)

    # Marker se pise SAMO za pun set izabrane kapije: `--suite X` je zelen za X,
    # ne za projekat, pa bi upis markera posle njega bio bas ono lazno zeleno
    # koje ceo ovaj sloj postoji da spreci. Isto vazi za `--category` i
    # `--compile-only`, a `--shuffle`/`--repeat` su dozvoljeni (isti skup).
    #
    # U marker ide i IME kapije: zelen `dev` run zadovoljava Stop hook, ali NE
    # zadovoljava release (koji trazi bar `pr`) -- `dev` ne vrti ni
    # RunNovacSmokeSuite ni external ugovore.
    # `--no-import` je OBAVEZAN deo uslova. Bez njega je postojala ova sekvenca:
    # repo src-vba = NOV kod, sveska = STARI kod, `--no-import --gate pr` prodje
    # nad starim kodom -- a marker dobije hash NOVOG izvora. To rusi centralni
    # invarijant cele ove arhitekture: hash mora da predstavlja kod koji je
    # zaista uvezen i izvrsen.
    full_set = (not args.suite and not args.category and not args.compile_only
                and not args.no_import and not args.no_mark_green)
    if rc == 0 and full_set:
        payload = dict(report["provenance"])
        payload["suites"] = sorted({s["name"] for s in report["suites"]})
        payload["gate"] = args.gate
        payload["note"] = f"run_vba.py --gate {args.gate}"
        vba_gate.write_green(payload)
        report["marked_green"] = payload["source_hash"]

    _write_report(report, v, rc, args.report_name)
    return rc


def _write_junit(report: dict, path: str) -> None:
    """JUnit XML -- da GitHub/IDE prikazu KOJI test je pao, a ne "Python exit 2"."""
    suites_el = ET.Element("testsuites", name="AgriX VBA")
    counts = report.get("counts") or {}
    for s in report.get("suites") or []:
        rec = counts.get(s["name"], {})
        n_pass = int(rec.get("pass") or 0)
        n_fail = int(rec.get("fail") or 0)
        total = n_pass + n_fail or 1
        el = ET.SubElement(suites_el, "testsuite", name=s["name"],
                           tests=str(total), failures=str(n_fail or (1 if s["status"] in
                                                                     ("FAIL", "TIMEOUT") else 0)),
                           skipped=str(1 if s["status"] == "BLIND" else 0),
                           time=str(s.get("seconds", 0)))
        case = ET.SubElement(el, "testcase", classname=s.get("category") or "vba",
                             name=s["name"], time=str(s.get("seconds", 0)))
        if s["status"] in ("FAIL", "TIMEOUT"):
            fail = ET.SubElement(case, "failure", type=s["status"])
            fail.text = s.get("error", s["status"])
        elif s["status"] == "BLIND":
            ET.SubElement(case, "skipped",
                          message="blind suite -- rezultat postoji samo u Immediate prozoru")
    ET.ElementTree(suites_el).write(path, encoding="utf-8", xml_declaration=True)


def _write_report(report: dict, v: dict, rc: int, name: str = "last_run",
                  out_dir: str = None) -> None:
    OUT_DIR_ = out_dir or OUT_DIR
    os.makedirs(OUT_DIR_, exist_ok=True)
    report["verdicts"] = {k: val for k, val in v.items() if not k.startswith("_")}
    with open(os.path.join(OUT_DIR_, f"{name}.json"), "w", encoding="utf-8") as fh:
        json.dump(report, fh, ensure_ascii=False, indent=2)
    try:
        _write_junit(report, os.path.join(OUT_DIR_, f"{name}.xml"))
    except Exception as exc:                # noqa: BLE001
        report.setdefault("import", []).append(f"JUnit XML nije upisan: {exc}")

    p = report.get("provenance") or {}
    lines = [
        f"IZVOR   src-vba {str(p.get('source_hash', ''))[:16]}..."
        + ("  (RADNO STABLO PRLJAVO)" if p.get("src_vba_dirty") else ""),
        f"        git {str(p.get('git_sha', ''))[:8]} {p.get('git_branch', '')} "
        f"{p.get('git_describe', '')}",
        f"        fixture {p.get('fixture')} {str(p.get('fixture_hash') or 'nema')[:16]}",
        f"        {p.get('platform', '')}  python {p.get('python', '')}  "
        f"locale {p.get('locale', '')}",
        f"KAPIJA  {report.get('gate')}"
        + (f"  shuffle seed={report.get('seed')}" if report.get("shuffle") else "")
        + (f"  repeat={report.get('repeat')}" if (report.get("repeat") or 1) > 1 else ""),
        "",
    ]

    for msg in report.get("manifest") or []:
        lines.append(f"MANIFEST {msg}")
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

    # Rezultat svake suite ide UZ NJU i nosi njeno ime. Jedan zajednicki blok je
    # znacio da poslednja suite sa rezultat-fajlom prepise prethodnu.
    slotovi = report.get("suite_results", {})
    for s in report["suites"]:
        lines.append(f"SUITE   {s['status']:7} {s['name']} ({s.get('seconds', 0)}s)"
                     + (f"  {s.get('error', '')}" if s["status"] in ("FAIL", "TIMEOUT") else ""))
        t = slotovi.get(s["name"])
        if not t:
            continue
        if t.get("error"):
            lines.append(f"TESTS   {s['name']}: {t['error']}")
        else:
            lines.append(f"TESTS   {s['name']}: {t['total']} ukupno, {t['failed']} palo")
        # Ime bas tog testa mora da se vidi u ispisu -- to je razlika izmedju
        # "nesto je palo" i upotrebljivog nalaza.
        for ln in t.get("detail", []):
            lines.append(f"        {ln}")

    for ln in report.get("counts_detail") or []:
        lines.append(f"ASSERTS {ln}")

    cleanup = report.get("cleanup")
    if cleanup:
        for ln in cleanup.get("detail") or []:
            lines.append(f"CLEANUP {ln}")

    for d in report.get("dialogs", []):
        lines.append(f"DIALOG  {d}")

    if report.get("source_gap"):
        lines.append(f"SOURCE  {report['source_gap']}")
    if report.get("behavior_skip"):
        lines.append(f"NAPOMENA {report['behavior_skip']}")
    if "fatal" in report:
        lines.append(f"FATAL   {report['fatal']}")

    lines.append("")
    lines.append("--- VERDIKT ------------------------------------------------")
    for key in ("STATIC", "COMPILE", "SCHEMA", "BEHAVIOR", "COUNTS", "CLEANUP", "EXTERNAL"):
        suffix = ""
        if key == "COMPILE" and v[key] == "UNKNOWN":
            suffix = "  (dozvoljeno uz zelen BEHAVIOR; release kapija ovo NE prihvata)"
        if key == "CLEANUP" and v[key] != "NOT_RUN":
            suffix = ("  (blokirajuce)" if v.get("_cleanup_enforced")
                      else "  (ne-blokirajuce u 'dev' kapiji; --enforce-cleanup da obara)")
        lines.append(f"{key:9} {v[key]}{suffix}")
    lines.append("------------------------------------------------------------")
    lines.append("ISHOD: " + ("PROSLO" if rc == 0 else "PALO"))

    blind = [s["name"] for s in report["suites"] if s["status"] == "BLIND"]
    if blind:
        lines.append("BLIND (suite bez fail-gate-a -- prosla bez greske NIJE dokaz da su "
                     "sve provere prosle): " + ", ".join(blind))
    if report.get("marked_green"):
        lines.append(f"ZELENO zapisano za izvor {report['marked_green'][:16]}... "
                     f"kapija={report.get('gate')} (vba_gate.py --status)")

    text = "\n".join(lines)
    with open(os.path.join(OUT_DIR_, f"{name}.txt"), "w", encoding="utf-8") as fh:
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
