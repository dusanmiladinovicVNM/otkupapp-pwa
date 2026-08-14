"""Release Gate: kapije koje moraju proci PRE nego sto tag i push uopste nastanu.

Ovo je odgovor na konkretan dogadjaj. `vba-v2.40.0` je merge-ovan i tag-ovan uz
posteno zapisanu napomenu da behavior runner nije ni pokrenut i da nov UI nema
nijedan behavior test. Bilo je posteno, i bilo je prekasno: tag je vec postojao.
Uzrok nije nepaznja nego redosled -- `release.ps1` je verziju, commit, tag i push
radio PRE nego sto operateru uopste kaze da uradi Import i Compile.

Zato se sada gate vrti kao poseban korak, a `release.ps1`/`release.sh` su svedeni
na git plumbing oko njega:

    STATIC     tools/vba_check.py + who_writes --check + manifest
    FIXTURE    fixture postoji i njegov hash ide u zapis
    BEHAVIOR   tools/run_vba.py --gate release nad OVIM izvorom
    GREEN      last-green marker se poklapa sa hashom src-vba KOJI SE TAG-UJE
    COMPILE    pravi Debug > Compile VBAProject -- UNKNOWN nije prihvatljivo

    python3 tools/release_gate.py --version 2.41.0
    python3 tools/release_gate.py --version 2.41.0 --dry-run
    python3 tools/release_gate.py --version 2.41.0 --waive external --reason "..."

Izlazni kod: 0 = sve kapije prosle (ili su izuzete uz zapisan razlog), 2 = stop.

ZASTO SE BUMP RADI PRE GATE-A (a ne posle)
    Gate mora da vrti TACNO ono sto ce biti tag-ovano. Ako se `APP_VERSION`
    promeni posle zelenog run-a, hash src-vba se menja i tag pokriva izvor koji
    niko nije testirao -- ista rupa, samo manja. Zato `release.ps1` prvo upise
    verziju u radno stablo (bez commita), pa pozove ovaj gate, pa tek onda
    commit-uje. Ako gate padne, bump se vraca i stablo ostaje cisto.

WAIVER
    `--waive X --reason "..."` je jedini nacin da kapija ne bude blokirajuca, i
    on ostavlja trag: ime kapije i razlog idu u `tests/last_release.json` i u
    poruku anotiranog taga. Waiver bez razloga se odbija. "Testovi nisu
    pokrenuti" vise ne moze da bude fusnota u release notes koju niko ne trazi.
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from datetime import datetime, timezone

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import vba_gate                       # noqa: E402

ROOT = vba_gate.ROOT
OUT = os.path.join(ROOT, "tests", "last_release.json")

GATES = ("static", "fixture", "behavior", "green", "compile", "external")


class Gate:
    def __init__(self, name: str):
        self.name = name
        self.status = "NOT_RUN"       # PASS | FAIL | NOT_RUN | WAIVED
        self.detail = ""

    def as_dict(self) -> dict:
        return {"gate": self.name, "status": self.status, "detail": self.detail}


def _run(cmd: list[str], timeout: float = 3600.0, echo: bool = False) -> tuple[int, str]:
    """Pokreni alat i vrati (rc, izlaz).

    `echo` ispisuje ceo izlaz na konzolu. Bez toga bi operater video samo skraceni
    `detail` iz verdikta -- a bas kod pale behavior kapije treba mu ime testa koji
    je pao, ne rezime.
    """
    try:
        p = subprocess.run(cmd, cwd=ROOT, capture_output=True, text=True, timeout=timeout)
    except FileNotFoundError as exc:
        return 127, str(exc)
    except subprocess.TimeoutExpired:
        return 124, f"timeout posle {timeout:g}s: {' '.join(cmd)}"
    out = (p.stdout or "") + (p.stderr or "")
    if echo and out.strip():
        print(f"--- {' '.join(cmd[1:])}")
        print(out.rstrip())
        print("---")
    return p.returncode, out


def _py() -> str:
    return sys.executable or "python3"


def _tail(text: str, n: int = 12) -> str:
    lines = [ln for ln in (text or "").splitlines() if ln.strip()]
    return " | ".join(lines[-n:])


# --- pojedinacne kapije --------------------------------------------------------

def gate_static(g: Gate, dry: bool) -> None:
    if dry:
        g.status, g.detail = "NOT_RUN", "dry-run"
        return
    rc1, out1 = _run([_py(), os.path.join("tools", "vba_check.py")], timeout=300)
    rc2, out2 = _run([_py(), os.path.join("tools", "who_writes.py"), "--check"], timeout=300)
    rc3, out3 = _run([_py(), os.path.join("tools", "vba_gate.py"), "--manifest-check"],
                     timeout=300)
    if rc1 == 0 and rc2 == 0 and rc3 == 0:
        g.status, g.detail = "PASS", "vba_check + who_writes + manifest"
        return
    g.status = "FAIL"
    bad = []
    if rc1:
        bad.append("vba_check: " + _tail(out1, 6))
    if rc2:
        bad.append("who_writes zastareo (python3 tools/who_writes.py --out): " + _tail(out2, 3))
    if rc3:
        bad.append("manifest: " + _tail(out3, 6))
    g.detail = " ;; ".join(bad)


def gate_fixture(g: Gate, fixture: str, dry: bool) -> None:
    if not os.path.exists(fixture):
        g.status = "FAIL"
        g.detail = (f"fixture ne postoji: {fixture} -- napravi ga: "
                    f'python tools/make_fixture.py --donor "<put do .xlsm>"')
        return
    h = vba_gate.file_hash(fixture)
    size = os.path.getsize(fixture)
    if size < 1024:
        g.status = "FAIL"
        g.detail = f"fixture je {size} B -- to nije prava sveska"
        return
    g.status, g.detail = "PASS", f"{os.path.basename(fixture)} sha256={h[:16]}... {size} B"


def gate_behavior(g: Gate, dry: bool, extra: list[str]) -> None:
    if dry:
        g.status, g.detail = "NOT_RUN", "dry-run"
        return
    cmd = [_py(), os.path.join("tools", "run_vba.py"), "--gate", "release"] + extra
    rc, out = _run(cmd, timeout=7200, echo=True)
    g.status = "PASS" if rc == 0 else "FAIL"
    g.detail = _tail(out, 4) + "  (pun ispis iznad i u tests/last_run.txt)"



def gate_green(g: Gate, dry: bool) -> None:
    """Trazi se marker iz bar `pr` kapije nad TACNO ovim izvorom.

    Zasto `pr` a ne `release`: external ugovori (Google/SEF) imaju svoju kapiju
    ispod i mogu legitimno da budu NOT_RUN ili WAIVED za konkretno izdanje. Ono
    sto ne sme da nedostaje je pun LOKALNI set nad kodom koji se tag-uje.
    """
    ok, cur, rec = vba_gate.is_green(need_gate="pr")
    if ok:
        g.status = "PASS"
        g.detail = (f"src-vba {cur[:16]}... zeleno "
                    f"{(rec or {}).get('timestamp_utc', '?')} "
                    f"kapija={(rec or {}).get('gate')} "
                    f"({len((rec or {}).get('suites') or [])} suite-ova)")
        return
    g.status = "FAIL"
    if rec is None:
        g.detail = (f"nema zapisa o zelenom run-u; tekuci izvor je {cur[:16]}... "
                    f"-- tag bi pokrio kod koji niko nije testirao")
    elif rec.get("source_hash") != cur:
        g.detail = (f"poslednji zeleni je {str(rec.get('source_hash', ''))[:16]}..., "
                    f"a tag-uje se {cur[:16]}... -- to nije isti kod")
    else:
        g.detail = (f"izvor jeste zelen, ali samo za kapiju '{rec.get('gate')}' "
                    f"-- release trazi bar 'pr' (dev ne vrti RunNovacSmokeSuite "
                    f"ni external ugovore)")


def _last_run() -> dict | None:
    path = os.path.join(ROOT, "tests", "last_run.json")
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:                       # noqa: BLE001
        return None


def gate_compile(g: Gate, dry: bool) -> None:
    """Release nema pravo na COMPILE UNKNOWN.

    Za svakodnevni Stop gate je best-effort dovoljan -- suite koja se izvrsila
    dokazuje da je VBA kompajlirao bar ono sto ona dodiruje. Za release to nije
    dovoljno: isporucuje se ceo projekat, ukljucujuci module koje nijedna suite
    ne dodiruje.

    Prvo se cita verdikt iz behavior run-a (on ionako radi compile). Drugi,
    namenski `--compile-only` prolaz ide SAMO kad taj verdikt nije jednoznacan --
    to je najstabilniji mod, jer je tamo probe jedini izvor istine pa NEJASNO i
    sam obara run.
    """
    if dry:
        g.status, g.detail = "NOT_RUN", "dry-run"
        return

    data = _last_run() or {}
    seen = (data.get("verdicts") or {}).get("COMPILE")
    if seen == "PASS":
        g.status, g.detail = "PASS", "behavior run: COMPILE OK"
        return
    if seen == "FAIL":
        g.status = "FAIL"
        g.detail = "behavior run: " + str((data.get("compile") or {}).get("error", "COMPILE FAIL"))
        return

    rc, out = _run([_py(), os.path.join("tools", "run_vba.py"), "--compile-only"],
                   timeout=1800, echo=True)
    if rc == 0:
        g.status, g.detail = "PASS", "run_vba.py --compile-only: COMPILE OK"
        return
    g.status = "FAIL"
    unknown = "NEJASNO" in out
    g.detail = _tail(out, 12) + (
        " ;; COMPILE NEJASNO -- headless probe nije mogao da utvrdi ishod. Uradi rucno "
        "Alt+F11 > Debug > Compile VBAProject; ako prodje, pusti gate sa "
        '--waive compile --reason "rucni Debug>Compile cist, <datum>"' if unknown else "")


def gate_external(g: Gate, dry: bool) -> None:
    """External ugovori (Google / MasterSync / SEF) nisu obavezni za svaki release,
    ali NE SMEJU da nestanu iz zapisa. Behavior gate ih vrti (`--gate release`),
    pa je ovde samo prepis njegovog ishoda: PASS, FAIL ili NOT_RUN."""
    if dry:
        g.status, g.detail = "NOT_RUN", "dry-run"
        return
    data = _last_run()
    if data is None:
        g.status, g.detail = "NOT_RUN", "nema tests/last_run.json"
        return
    ext = [s for s in (data.get("suites") or []) if s.get("external")]
    g.status = (data.get("verdicts") or {}).get("EXTERNAL", "NOT_RUN")
    g.detail = ", ".join(f"{s['name']}={s['status']}" for s in ext) or "nijedna nije isla"


# --- orkestracija ---------------------------------------------------------------

def run_gates(version: str, fixture: str, waivers: dict, dry: bool,
              extra_run_args: list[str]) -> tuple[int, dict]:
    order = [Gate(n) for n in GATES]
    by_name = {g.name: g for g in order}

    gate_static(by_name["static"], dry)
    gate_fixture(by_name["fixture"], fixture, dry)

    # BEHAVIOR pre GREEN: run_vba upisuje marker kad prodje, pa GREEN posle njega
    # proverava da je marker bas za OVAJ izvor (a ne za neki raniji zeleni run).
    gate_behavior(by_name["behavior"], dry, extra_run_args)
    gate_green(by_name["green"], dry)
    gate_compile(by_name["compile"], dry)
    gate_external(by_name["external"], dry)

    for name, reason in waivers.items():
        g = by_name[name]
        prev = (g.detail or "").strip()
        if len(prev) > 160:
            prev = prev[:157] + "..."
        g.detail = f"WAIVED ({reason}) -- bilo je: {g.status} {prev}".strip()
        g.status = "WAIVED"

    blocked = [g.name for g in order if g.status in ("FAIL", "NOT_RUN")
               and not (dry and g.status == "NOT_RUN")]
    # NOT_RUN blokira isto kao FAIL: "nije pokrenuto" nije "proslo". Jedini izlaz
    # je waiver, koji ostavlja trag.
    rc = 2 if blocked else 0

    record = {
        "version": version,
        "tag": f"vba-v{version}",
        "timestamp_utc": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "result": "PASS" if rc == 0 else "BLOCKED",
        "blocked_by": blocked,
        "provenance": vba_gate.provenance(fixture),
        "gates": [g.as_dict() for g in order],
        "dry_run": dry,
    }
    return rc, record


def verdict_text(record: dict) -> str:
    p = record["provenance"]
    lines = [
        f"AgriX release gate -- {record['tag']}",
        f"  src-vba   {p['source_hash']}",
        f"  git       {p['git_sha'][:12]} ({p['git_branch']})"
        + ("  RADNO STABLO PRLJAVO" if p["src_vba_dirty"] else ""),
        f"  fixture   {p['fixture']} {str(p['fixture_hash'] or 'nema')[:16]}",
        f"  masina    {p['platform']}  locale {p['locale']}",
        f"  vreme     {record['timestamp_utc']}",
        "",
    ]
    for g in record["gates"]:
        lines.append(f"  {g['gate'].upper():9} {g['status']:8} {g['detail']}")
    lines.append("")
    lines.append(f"  ISHOD     {record['result']}"
                 + (f"  (blokira: {', '.join(record['blocked_by'])})"
                    if record["blocked_by"] else ""))
    return "\n".join(lines)


def parse_args(argv: list[str]) -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="Release kapije pre taga i push-a.")
    ap.add_argument("--version", required=True, help="verzija koja se izdaje, npr. 2.41.0")
    ap.add_argument("--fixture", default=vba_gate.DEFAULT_FIXTURE)
    ap.add_argument("--dry-run", action="store_true",
                    help="ne pokreci Excel; pokazi redosled kapija i sta bi blokiralo")
    ap.add_argument("--waive", action="append", default=[], choices=list(GATES),
                    help="izuzmi kapiju (trazi --reason); zapisuje se u tag i metadata")
    ap.add_argument("--reason", default="", help="razlog za --waive")
    ap.add_argument("--run-arg", action="append", default=[],
                    help="dodatni argument za run_vba.py (npr. --run-arg --shuffle)")
    ap.add_argument("--tag-message", action="store_true",
                    help="ispisi SAMO tekst za poruku anotiranog taga (bez pokretanja kapija)")
    return ap.parse_args(argv)


def main(argv: list[str]) -> int:
    args = parse_args(argv)
    version = args.version.lstrip("v")

    if args.tag_message:
        if not os.path.exists(OUT):
            print(f"Nema {OUT} -- pokreni gate pre taga.", file=sys.stderr)
            return 2
        with open(OUT, "r", encoding="utf-8") as fh:
            print(verdict_text(json.load(fh)))
        return 0

    if args.waive and not args.reason.strip():
        print("--waive trazi --reason \"...\". Izuzeta kapija bez zapisanog razloga je "
              "tacno ono sto je proizvelo v2.40.0.", file=sys.stderr)
        return 2

    waivers = {name: args.reason.strip() for name in args.waive}
    rc, record = run_gates(version, args.fixture, waivers, args.dry_run, args.run_arg)

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as fh:
        json.dump(record, fh, ensure_ascii=False, indent=2)

    print(verdict_text(record))
    print(f"\nZapis: {OUT}")
    if rc != 0 and not args.dry_run:
        print("\nTAG I PUSH SE NE RADE. Popravi kapiju, ili je izuzmi svesno:\n"
              f"    python3 tools/release_gate.py --version {version} "
              f"--waive <kapija> --reason \"...\"", file=sys.stderr)
    return rc


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
