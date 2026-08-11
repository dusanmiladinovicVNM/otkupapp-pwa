"""Staticke provere nad src-vba -- mehanizovan CLAUDE.md sec.4 i sec.5.

Radi svuda (nema Excela, nema COM-a), pa je jedina verifikacija koja postoji i u
Claude Code sesiji na Linux/macOS masini. Namena: "verify > conclude" prestaje da
bude molba i postaje exit kod.

    python tools/vba_check.py                 # sve provere nad src-vba/
    python tools/vba_check.py fajl1.bas ...   # samo nad datim fajlovima
    python tools/vba_check.py --hook          # tiho kad je cisto (za PostToolUse hook)

Izlazni kod: 0 = cisto, 2 = ima nalaza (blokirajuce za hook).

Provere:
  1. ASCII        -- svaki VBA izvor mora ostati 100% ASCII (dijakritika ide kroz
                     modPoruke/ChrW). Ne-ASCII bajt = `ImportAllVBA` ucita smece.
  2. DEKLARACIJA  -- modul-level Const/promenljiva/Declare/Type/Enum posle prve
                     procedure. VBA to NE kompajlira, a to je prirodno mesto na
                     koje deklaracija padne kad se pise "uz funkciju koja je koristi".
  3. REZERVISANO  -- ime promenljive/konstante koje se case-insensitive poklapa sa
                     VBA kljucnom reci (`Dim eNum As Long` -> `Enum` -> compile error).
  4. DUPLIKAT     -- isti Public Sub/Function/Const u dva modula = "Ambiguous name"
                     posle merge-a.
  5. PORUKA       -- `Poruka("KLJUC")` bez para u `modPoruke.UpsertPoruke`.
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from collections import defaultdict

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_VBA = os.path.join(ROOT, "src-vba")
VBA_EXT = (".bas", ".cls", ".frm", ".doccls")

# Reci koje VBA NE prihvata kao ime promenljive -- compile-hard podskup.
# VBA je case-insensitive, pa `Dim eNum As Long` = `Enum` i obara compile (RF-06).
#
# CLAUDE.md sec.4 navodi siru listu (`name`, `line`, `text`, `date`, `base`,
# `time`, `mid`, `local`, `read`...). Te reci su STILSKA preporuka, ne compile
# greska: postojeci kod ih vec koristi kao imena promenljivih i kompajlira se
# (modDrive.bas `name`, modJournaling.bas `line`, modTheme.bas `text`,
# modBrojevi.bas `base`). Da checker ne bi vikao na 20 mesta zatecenog koda koje
# niko nece menjati, ovde je samo ono sto stvarno obara compile.
RESERVED = {
    "and", "as", "byref", "byval", "call", "case", "class", "const", "declare",
    "dim", "do", "each", "else", "elseif", "empty", "end", "enum", "eqv",
    "erase", "event", "exit", "false", "for", "friend", "function", "global",
    "gosub", "goto", "if", "imp", "implements", "in", "is", "let", "lib",
    "like", "loop", "me", "mod", "new", "next", "nothing", "not", "null", "on",
    "option", "optional", "or", "paramarray", "preserve", "private", "property",
    "public", "raiseevent", "redim", "rem", "resume", "select", "set", "static",
    "step", "stop", "sub", "then", "to", "true", "type", "until", "wend",
    "while", "with", "withevents", "xor",
    # imena tipova su takodje rezervisana
    "boolean", "byte", "currency", "double", "integer", "long", "longlong",
    "longptr", "object", "single", "string", "variant",
}

PROC_START = re.compile(
    r"^\s*(?:Public\s+|Private\s+|Friend\s+|Global\s+)?(?:Static\s+)?"
    r"(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+\w+", re.IGNORECASE)

MODULE_DECL = re.compile(
    r"^(Public|Private|Global)\s+"
    r"(Const\b|Declare\b|Type\b|Enum\b|WithEvents\b|\w+\s+As\b|\w+\s*\()", re.IGNORECASE)

PUBLIC_PROC = re.compile(
    r"^Public\s+(?:Static\s+)?(?:Sub|Function|Const)\s+(\w+)", re.IGNORECASE)

# Ime deklarisane promenljive/konstante -- modifikatori se preskacu, pa
# `Public Const FOO As String` daje FOO, a ne "Const".
DECL_NAMES = re.compile(
    r"^\s*(?:Public|Private|Global|Dim|ReDim|Static|Const)\s+"
    r"(?:(?:Const|Static|WithEvents|Preserve)\s+)*(\w+)", re.IGNORECASE)

# Linije koje NISU deklaracija promenljive -- ime posle kljucne reci je ime
# procedure/tipa, ne promenljive.
NOT_A_VAR = re.compile(
    r"^\s*(?:Public\s+|Private\s+|Friend\s+|Global\s+)?(?:Static\s+)?"
    r"(?:Declare\b|Sub\b|Function\b|Property\b|Type\b|Enum\b|Event\b)", re.IGNORECASE)

PARAM_NAMES = re.compile(r"(?:ByVal|ByRef)\s+(\w+)", re.IGNORECASE)

PORUKA_USE = re.compile(r'Poruka\(\s*"([A-Z0-9_]+)"\s*\)')
PORUKA_DEF = re.compile(r'UpsertRow\s+lo,\s*existing,\s*"([A-Z0-9_]+)"')


class Finding:
    def __init__(self, path: str, line: int, code: str, msg: str):
        self.path, self.line, self.code, self.msg = path, line, code, msg

    def __str__(self) -> str:
        rel = os.path.relpath(self.path, ROOT)
        return f"{rel}:{self.line}: {self.code}: {self.msg}"


def vba_files(paths: list[str]) -> list[str]:
    if paths:
        return [os.path.abspath(p) for p in paths if p.lower().endswith(VBA_EXT)]
    return [os.path.join(SRC_VBA, n) for n in sorted(os.listdir(SRC_VBA))
            if n.lower().endswith(VBA_EXT)]


def check_ascii(path: str, raw: bytes) -> list[Finding]:
    out = []
    for i, line in enumerate(raw.split(b"\n"), start=1):
        bad = [b for b in line if b > 0x7F]
        if bad:
            chars = "".join(f"\\x{b:02x}" for b in bad[:6])
            out.append(Finding(path, i, "ASCII",
                               f"ne-ASCII bajt ({chars}). Tekst sa dijakritikom ide kroz "
                               f'modPoruke.UpsertPoruke + Poruka("KLJUC"), ne u izvor.'))
    return out


def check_decl_after_proc(path: str, lines: list[str]) -> list[Finding]:
    out, first_proc = [], None
    for i, line in enumerate(lines, start=1):
        if first_proc is None and PROC_START.match(line):
            first_proc = i
            continue
        if first_proc is None:
            continue
        if PROC_START.match(line):
            continue
        m = MODULE_DECL.match(line)
        if m:
            out.append(Finding(path, i, "DEKLARACIJA",
                               f"modul-level deklaracija posle prve procedure (linija {first_proc}). "
                               f"VBA ovo NE kompajlira -- premesti u deklaracionu sekciju na vrh."))
    return out


def check_reserved(path: str, lines: list[str]) -> list[Finding]:
    out = []
    for i, line in enumerate(lines, start=1):
        stripped = line.strip()
        if stripped.startswith("'"):
            continue
        names = []
        if not NOT_A_VAR.match(line):
            m = DECL_NAMES.match(line)
            if m:
                names.append(m.group(1))
        names.extend(PARAM_NAMES.findall(line))
        for n in names:
            if n.lower() in RESERVED:
                out.append(Finding(path, i, "REZERVISANO",
                                   f"'{n}' se case-insensitive poklapa sa VBA kljucnom reci "
                                   f"-- compile error. Koristi konvenciju projekta "
                                   f"(errNum/errDesc/errSrc)."))
    return out


def collect_public(path: str, lines: list[str]) -> list[tuple[str, int]]:
    """Public Sub/Function/Const van `#If ... #End If` blokova.

    Uslovna kompilacija namerno definise isto ime u vise grana (modMouseWheel ima
    VBA7 implementaciju i pre-VBA7 no-op stubove) -- to NIJE "Ambiguous name",
    jer se u projekat kompajlira samo jedna grana.
    """
    out, cond_depth = [], 0
    for i, line in enumerate(lines, start=1):
        stripped = line.strip().lower()
        if stripped.startswith("#if"):
            cond_depth += 1
            continue
        if stripped.startswith("#end if"):
            cond_depth = max(0, cond_depth - 1)
            continue
        if cond_depth:
            continue
        m = PUBLIC_PROC.match(line)
        if m:
            out.append((m.group(1), i))
    return out


def check_poruke(files: list[str]) -> list[Finding]:
    poruke_path = os.path.join(SRC_VBA, "modPoruke.bas")
    if not os.path.exists(poruke_path):
        return []
    with open(poruke_path, "r", encoding="ascii", errors="replace") as fh:
        defined = set(PORUKA_DEF.findall(fh.read()))

    out = []
    for path in files:
        if os.path.basename(path) == "modPoruke.bas":
            continue
        with open(path, "r", encoding="ascii", errors="replace") as fh:
            for i, line in enumerate(fh, start=1):
                for key in PORUKA_USE.findall(line):
                    if key not in defined:
                        out.append(Finding(path, i, "PORUKA",
                                           f'"{key}" nema par u modPoruke.UpsertPoruke '
                                           f"(orphan kljuc -- prikazace se prazno)."))
    return out


def main(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(description="Staticke provere nad src-vba")
    ap.add_argument("paths", nargs="*", help="konkretni fajlovi (podrazumevano ceo src-vba/)")
    ap.add_argument("--hook", action="store_true", help="bez izlaza kad je cisto")
    args = ap.parse_args(argv)

    files = vba_files(args.paths)
    if not files:
        return 0

    findings: list[Finding] = []
    publics: dict[str, list[tuple[str, int]]] = defaultdict(list)

    for path in files:
        with open(path, "rb") as fh:
            raw = fh.read()
        findings += check_ascii(path, raw)

        lines = raw.decode("ascii", errors="replace").replace("\r\n", "\n").split("\n")
        findings += check_decl_after_proc(path, lines)
        findings += check_reserved(path, lines)
        # Samo standardni moduli (.bas) dele globalni imenski prostor. Public clan
        # forme ili klase (.frm/.cls/.doccls) je clan tog objekta, ne globalno ime,
        # pa isto ime u dve forme NIJE "Ambiguous name".
        if path.lower().endswith(".bas"):
            for name, ln in collect_public(path, lines):
                publics[name.lower()].append((path, ln))

    # Duplikate trazimo samo kad se gleda ceo src-vba -- na podskupu fajlova bi
    # nalaz bio lazno negativan i zbunjujuci.
    if not args.paths:
        for name, sites in sorted(publics.items()):
            if len(sites) > 1:
                where = ", ".join(f"{os.path.basename(p)}:{ln}" for p, ln in sites)
                findings.append(Finding(sites[0][0], sites[0][1], "DUPLIKAT",
                                        f"Public '{name}' definisan na vise mesta ({where}) "
                                        f'-- VBA "Ambiguous name detected".'))

    findings += check_poruke(files)

    if not findings:
        if not args.hook:
            print(f"vba_check: cisto ({len(files)} fajlova).")
        return 0

    by_code: dict[str, int] = defaultdict(int)
    for f in sorted(findings, key=lambda f: (f.path, f.line)):
        print(str(f), file=sys.stderr)
        by_code[f.code] += 1
    summary = ", ".join(f"{k}={v}" for k, v in sorted(by_code.items()))
    print(f"\nvba_check: {len(findings)} nalaza ({summary}).", file=sys.stderr)
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
