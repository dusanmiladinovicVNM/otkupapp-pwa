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
                     (izuzetak: ugovor ekrana `Scr_*` u `modScr*` -- v. SCR_UGOVOR)
                     posle merge-a.
  5. PORUKA       -- `Poruka("KLJUC")` bez para u `modPoruke.UpsertPoruke`.
  6. NEDEFINISAN  -- poziv procedure koja nigde u projektu nije definisana
                     ("Sub or Function not defined").
  7. ARNOST       -- poziv sa pogresnim brojem argumenata ("Wrong number of
                     arguments").
  8. DUPLIKAT_LOKALNI -- isto ime dva puta u ISTOM modulu (izuzetak: Property
                     Get/Let/Set trojka). Modul se ne kompajlira, a greska se
                     javlja kao "Cannot run the macro" na bilo kom makrou.

Provere 6 i 7 pokrivaju dve najcesce compile greske u ovom projektu -- one zbog
kojih je i pravljen headless compile gate koji se nije dao ukrotiti
(docs/EXCEL_TEST_HARNESS.md). Ovde se hvataju bez Excela, u milisekundama.
Ne pokrivaju: tipove, nedeklarisane promenljive, greske u .frm/.cls.
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

# --- izuzetak od DUPLIKAT-a: ugovor ekrana novog UI-ja ---------------------
#
# Ljuska `modOtkupUI` ne poznaje nijedan ekran po imenu: svaki ekranski modul
# (`modScr*`) implementira isti skup procedura, a ljuska ih zove ISKLJUCIVO
# kasno vezano -- `Application.Run "modScrDokumenti.Scr_Rows"`. Poziv je uvek
# kvalifikovan imenom modula, pa VBA nema sta da razresava i "Ambiguous name"
# ne nastaje (potvrdjeno: oba modula su u projektu i kompajlira se).
#
# Izuzetak je namerno uzak i vazi SAMO kad su SVI definicioni fajlovi ekranski
# moduli. Isto ime u bilo kom drugom modulu i dalje pada -- ukljucujuci slucaj
# kad neko ugovornu proceduru prekopira u obican modul pa je pozove nekvalifi-
# kovano, sto je bas greska koju ova provera treba da uhvati.
SCR_UGOVOR = {
    "scr_meta", "scr_build", "scr_layout", "scr_rows", "scr_event",
    "scr_save", "scr_resetcache", "scr_liste", "scr_lista", "scr_radnje",
    "scr_naslovdopuna",
}


def je_ekranski_modul(path: str) -> bool:
    return os.path.basename(path).lower().startswith("modscr")

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


# --- 6. NEDEFINISAN ----------------------------------------------------------
#
# "Sub or Function not defined" je, uz "Ambiguous name" (v. DUPLIKAT), najcesca
# compile greska u ovom projektu. Obe se vide iz izvora, pa Excel ovde uopste ne
# treba -- sto je dobro, jer se headless compile gate nije dao ukrotiti
# (docs/EXCEL_TEST_HARNESS.md).
#
# Provera je NAMERNO uska:
#   - gleda SAMO .bas module. U .frm/.cls se nasledjeni clanovi zovu bez
#     kvalifikatora (`Repaint`, `Show`, `SetFocus`), pa bi tamo lazni nalazi bili
#     pravilo, a ne izuzetak.
#   - gleda SAMO poziv u poziciji naredbe (`Foo`, `Foo a, b`, `Call Foo(a)`).
#     Izraz `x = Foo(1)` se ne dira: bez tipova se poziv funkcije ne razlikuje od
#     indeksiranja niza.
# Lazan nalaz u hook-u je gori od propustenog, pa je prag namerno visok.

PROC_DEF = re.compile(
    r"^\s*(?:Public\s+|Private\s+|Friend\s+|Global\s+)?(?:Static\s+)?"
    r"(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)", re.IGNORECASE)

DECLARE_DEF = re.compile(
    r"^\s*(?:Public\s+|Private\s+|Global\s+)?Declare\s+(?:PtrSafe\s+)?"
    r"(?:Sub|Function)\s+(\w+)", re.IGNORECASE)

CALL_STMT = re.compile(r"^(?:Call\s+)?([A-Za-z_]\w*)\s*(.*)$", re.IGNORECASE)

BLOCK_OPEN = re.compile(r"^\s*(?:Public\s+|Private\s+)?(Type|Enum)\s+\w+", re.IGNORECASE)
BLOCK_CLOSE = re.compile(r"^\s*End\s+(Type|Enum)\b", re.IGNORECASE)

# Reci koje na pocetku naredbe NISU poziv procedure: VBA naredbe, kljucne reci i
# ugradjene rutine koje se zovu bez tacke.
STMT_WORDS = {
    "if", "for", "next", "do", "loop", "while", "wend", "select", "case", "end",
    "exit", "on", "resume", "goto", "gosub", "return", "with", "set", "let",
    "dim", "redim", "const", "static", "public", "private", "friend", "global",
    "type", "enum", "declare", "sub", "function", "property", "option", "erase",
    "stop", "rem", "else", "elseif", "then", "call", "implements", "attribute",
    "raiseevent", "event", "open", "close", "print", "write", "input", "put",
    "get", "seek", "lock", "unlock", "width", "line", "name", "kill", "mkdir",
    "rmdir", "chdir", "chdrive", "setattr", "filecopy", "reset", "randomize",
    "beep", "doevents", "load", "unload", "msgbox", "debug", "err", "error",
    "date", "time", "sendkeys", "appactivate", "savesetting", "deletesetting",
    "lset", "rset", "mid", "midb", "version", "begin", "multiuse", "true",
    "false", "nothing", "me", "new", "each", "to", "step", "is", "and", "or",
    "not", "xor", "mod", "like", "imp", "eqv", "byval", "byref", "optional",
    "paramarray", "preserve", "in", "as", "lib", "alias", "withevents", "class",
    "application", "sleep", "shell",
}


def collect_definitions(files: list[str]) -> set[str]:
    """Sva imena procedura definisana bilo gde u projektu (sva rasirenja)."""
    names: set[str] = set()
    for path in files:
        with open(path, "r", encoding="ascii", errors="replace") as fh:
            for line in fh:
                for rx in (DECLARE_DEF, PROC_DEF):
                    m = rx.match(line)
                    if m:
                        names.add(m.group(1).lower())
                        break
    return names


def _strip_comment(text: str) -> str:
    """Odbaci prateci ' komentar, ali ne apostrof unutar stringa."""
    in_str = False
    for i, ch in enumerate(text):
        if ch == '"':
            in_str = not in_str
        elif ch == "'" and not in_str:
            return text[:i].rstrip()
    return text


def _split_top_level(text: str) -> list[str]:
    """Podeli po zarezima koji NISU unutar zagrada ili navodnika."""
    parts, depth, in_str, cur = [], 0, False, ""
    for ch in text:
        if ch == '"':
            in_str = not in_str
        elif not in_str and ch in "([":
            depth += 1
        elif not in_str and ch in ")]":
            depth -= 1
        elif not in_str and ch == "," and depth == 0:
            parts.append(cur.strip())
            cur = ""
            continue
        cur += ch
    if cur.strip():
        parts.append(cur.strip())
    return parts


def collect_arities(files: list[str]) -> dict[str, tuple[int, float]]:
    """Ime procedure -> (min, max) broj argumenata; max = inf uz ParamArray.

    Ime definisano na vise mesta sa razlicitom arnoscu se ISKLJUCUJE -- tu se bez
    razresavanja opsega ne moze tvrditi sta je pozvano.
    """
    seen: dict[str, set[tuple[int, float]]] = defaultdict(set)
    for path in files:
        with open(path, "r", encoding="ascii", errors="replace") as fh:
            lines = fh.read().replace("\r\n", "\n").split("\n")
        i = 0
        while i < len(lines):
            line = lines[i]
            while line.rstrip().endswith("_") and i + 1 < len(lines):
                line = line.rstrip()[:-1] + " " + lines[i + 1]
                i += 1
            m = PROC_DEF.match(line) or DECLARE_DEF.match(line)
            if m and "(" in line:
                params = _split_top_level(line[line.index("(") + 1:line.rindex(")")]
                                          if ")" in line else "")
                lo = sum(1 for p in params
                         if p and not re.match(r"^(Optional|ParamArray)\b", p, re.IGNORECASE))
                hi: float = float("inf") if any(
                    re.match(r"^ParamArray\b", p, re.IGNORECASE) for p in params) else len(params)
                seen[m.group(1).lower()].add((lo, hi))
            i += 1
    return {name: next(iter(v)) for name, v in seen.items() if len(v) == 1}


def check_undefined(path: str, lines: list[str], defined: set[str],
                    arities: dict) -> list[Finding]:
    if not path.lower().endswith(".bas"):
        return []

    out: list[Finding] = []
    block_depth = 0
    continued = False

    for i, raw in enumerate(lines, start=1):
        line = raw.rstrip()
        was_continued, continued = continued, line.endswith("_")

        if BLOCK_OPEN.match(line):
            block_depth += 1
            continue
        if BLOCK_CLOSE.match(line):
            block_depth = max(0, block_depth - 1)
            continue
        if block_depth or was_continued:
            continue

        stmt = _strip_comment(line.strip())
        if not stmt or stmt.startswith("'") or stmt.startswith("#") or ":" in stmt:
            continue        # komentar, uslovna kompilacija, labela ili vise naredbi
        if PROC_DEF.match(line) or DECLARE_DEF.match(line):
            continue

        explicit_call = bool(re.match(r"^Call\s", stmt, re.IGNORECASE))
        m = CALL_STMT.match(stmt)
        if not m:
            continue
        name, rest = m.group(1), m.group(2).lstrip()

        if name.lower() in STMT_WORDS or name.lower() in RESERVED:
            continue
        # `Foo = 1` (dodela), `Foo.Bar` (clan), `Foo As Long` (clan tipa) --
        # nista od toga nije poziv procedure.
        if rest.startswith(("=", ".", "!")) or re.match(r"^As\s", rest, re.IGNORECASE):
            continue
        # `Foo(kljuc).Add x` / `Foo(i) = 1` -- indeksiranje kolekcije ili niza.
        # Bez `Call` prefiksa, ime sa zagradom na pocetku naredbe je u ovom
        # kodu uvek indeks, ne poziv. (Sve 8 prvih laznih nalaza bilo je ovo.)
        if rest.startswith("(") and not explicit_call:
            continue
        if name.lower() not in defined:
            out.append(Finding(path, i, "NEDEFINISAN",
                               f"poziv '{name}' -- nigde u src-vba nije definisan "
                               f'Sub/Function/Property. VBA: "Sub or Function not defined".'))
            continue

        # Arnost -- druga polovina istog compile problema ("Wrong number of
        # arguments"). Proverava se samo kad je poziv cela naredba u jednoj
        # liniji i kad je ime jednoznacno definisano.
        span = arities.get(name.lower())
        if span is None or line.rstrip().endswith("_"):
            continue
        args = rest[1:rest.rindex(")")] if explicit_call and rest.startswith("(") else rest
        n_args = len(_split_top_level(args))
        lo, hi = span
        if n_args < lo or n_args > hi:
            ocekivano = f"{lo:g}" if lo == hi else (
                f"{lo:g}-{hi:g}" if hi != float("inf") else f"{lo:g}+")
            out.append(Finding(path, i, "ARNOST",
                               f"poziv '{name}' sa {n_args} argumenata, a deklarisano je "
                               f'{ocekivano}. VBA: "Wrong number of arguments".'))
    return out


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


# --- 8. DUPLIKAT_LOKALNI -----------------------------------------------------
#
# DUPLIKAT (provera 4) gleda GLOBALNI imenski prostor -- isto Public ime u dva
# modula. Duplo ime unutar JEDNOG modula mu je nevidljivo, a obara compile isto
# tako: dupli `Private Const FX_FAKTURA_BEZ_IZNOSA` u modTest.bas prosao je
# checker cist, a projekat se posle toga nije kompajlirao. Simptom nije bio
# "Ambiguous name" nego "Cannot run the macro" na SVAKOM makrou -- modul koji se
# ne kompajlira obara ceo projekat, pa greska izgleda kao da je bilo gde.
#
# VBA u jednom modulu ne trpi dva clana istog imena, sa TACNO JEDNIM izuzetkom:
# `Property Get/Let/Set X` je trojka nad istim imenom. Zato se procedura pamti sa
# vrstom, pa se trojka prepoznaje, a `Property Get X` dvaput i dalje pada.
#
# Za razliku od DUPLIKAT-a ova provera radi i nad .frm/.cls: ogranicenje na .bas
# je tamo bilo zato sto Public clan forme nije globalno ime -- unutar modula je
# sudar sudar bez obzira na vrstu fajla.
#
# Namerno se NE gleda:
#   - `Const`/`Dim` unutar procedure -- lokalni su, isto ime u dve procedure je
#     potpuno legalno i najcesci oblik u kodu;
#   - druga i dalja imena iz `Private a As Long, b As Long` -- promasaj, ne
#     lazan nalaz.

PROP_DEF = re.compile(
    r"^\s*(?:Public\s+|Private\s+|Friend\s+|Global\s+)?(?:Static\s+)?"
    r"Property\s+(Get|Let|Set)\s+(\w+)", re.IGNORECASE)


def collect_local_names(lines: list[str]) -> dict[str, list[tuple[str, int]]]:
    """ime -> [(vrsta, linija)] za sve clanove modula.

    Vrsta je "Get"/"Let"/"Set" za Property, inace "proc" ili "deklaracija".
    Uslovna kompilacija se preskace iz istog razloga kao u collect_public.
    """
    names: dict[str, list[tuple[str, int]]] = defaultdict(list)
    cond_depth = 0
    in_block = False
    first_proc = None

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

        # Clanovi Type/Enum bloka su u imenskom prostoru tog tipa, ne modula.
        if BLOCK_OPEN.match(line):
            in_block = True
            continue
        if in_block:
            if BLOCK_CLOSE.match(line):
                in_block = False
            continue

        m = PROP_DEF.match(line)
        if m:
            first_proc = first_proc or i
            names[m.group(2).lower()].append((m.group(1).capitalize(), i))
            continue

        m = PROC_DEF.match(line)
        if m:
            first_proc = first_proc or i
            names[m.group(1).lower()].append(("proc", i))
            continue

        m = DECLARE_DEF.match(line)
        if m:
            names[m.group(1).lower()].append(("proc", i))
            continue

        # Deklaracije samo IZNAD prve procedure -- ispod su lokalne (a modul-level
        # deklaracija na tom mestu je vec nalaz provere DEKLARACIJA).
        if first_proc is None and not NOT_A_VAR.match(line):
            m = DECL_NAMES.match(line)
            if m:
                names[m.group(1).lower()].append(("deklaracija", i))

    return names


def check_local_dupes(path: str, lines: list[str]) -> list[Finding]:
    out = []
    for name, sites in sorted(collect_local_names(lines).items()):
        if len(sites) < 2:
            continue
        kinds = [k for k, _ in sites]
        # Property trojka nad istim imenom: legalna dok je svaka vrsta jednom.
        if all(k in ("Get", "Let", "Set") for k in kinds) and len(set(kinds)) == len(kinds):
            continue
        where = ", ".join(f"{k}@{ln}" for k, ln in sites)
        out.append(Finding(path, sites[1][1], "DUPLIKAT_LOKALNI",
                           f"'{name}' definisan {len(sites)} puta u istom modulu ({where}) "
                           f"-- modul se NE kompajlira, a greska se javlja kao "
                           f'"Cannot run the macro" na bilo kom makrou.'))
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


# --- self-test: dokaz u oba smera, trajno ------------------------------------
#
# CLAUDE.md par.5 trazi dvosmerni dokaz kad se menja SAM CHECKER: zelena provera
# koja nikad nije pokazana crvena ne dokazuje da isla sta meri. Za DUPLIKAT_LOKALNI
# taj dokaz ne moze da bude sabotaza u src-vba (sabotaza.py obara modTest testove,
# a ovo je staticka provera), pa stoji ovde -- i vrti se u CI-ju, na svakom PR-u.
#
# Svaki slucaj je (naziv, ocekivan broj nalaza, izvor). Nula znaci "legalan VBA
# koji NE sme da zapisti" -- ta polovina je vaznija: lazan nalaz u PostToolUse
# hook-u je gori od propustenog, jer uci da se checker ignorise.

SELF_TEST_CASES = [
    # --- mora da zapisti ---
    ("dupli Private Const (zatecen incident)", 1, """Option Explicit
Private Const FX_FAKTURA As String = "FAK-TEST-0"
Private Const FX_DRUGO As String = "X"
Private Const FX_FAKTURA As String = "FAK-TEST-0"
Public Sub Radi()
End Sub
"""),
    ("Sub i Function istog imena", 1, """Option Explicit
Public Sub Obradi()
End Sub
Private Function Obradi() As Long
End Function
"""),
    ("dva Property Get istog imena", 1, """Option Explicit
Public Property Get Ime() As String
End Property
Public Property Get Ime() As String
End Property
"""),
    # --- ne sme da zapisti ---
    ("Property Get/Let/Set trojka", 0, """Option Explicit
Private mIme As String
Public Property Get Ime() As String
    Ime = mIme
End Property
Public Property Let Ime(ByVal v As String)
    mIme = v
End Property
Public Property Set Ime(ByVal v As Object)
End Property
"""),
    ("isto lokalno ime u dve procedure", 0, """Option Explicit
Public Sub Prva()
    Const LIMIT As Long = 10
    Dim i As Long
End Sub
Public Sub Druga()
    Const LIMIT As Long = 20
    Dim i As Long
End Sub
"""),
    ("uslovna kompilacija (modMouseWheel obrazac)", 0, """Option Explicit
#If VBA7 Then
Public Sub HookMouse()
End Sub
#Else
Public Sub HookMouse()
End Sub
#End If
"""),
    ("clan Type/Enum bloka nije clan modula", 0, """Option Explicit
Public Type TRed
    Naziv As String
End Type
Public Enum EStatus
    Naziv = 1
End Enum
Public Sub Naziv()
End Sub
"""),
]


def self_test() -> int:
    palo = []
    for naziv, ocekivano, izvor in SELF_TEST_CASES:
        lines = izvor.replace("\r\n", "\n").split("\n")
        dobijeno = len(check_local_dupes("<self-test>", lines))
        if dobijeno != ocekivano:
            palo.append(f"  {naziv}: ocekivano {ocekivano} nalaza, dobijeno {dobijeno}")

    for line in palo:
        print(line, file=sys.stderr)
    if palo:
        print(f"\nself-test: {len(palo)} od {len(SELF_TEST_CASES)} slucajeva palo.",
              file=sys.stderr)
        return 2
    print(f"self-test: cisto ({len(SELF_TEST_CASES)} slucajeva DUPLIKAT_LOKALNI).")
    return 0


def main(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(description="Staticke provere nad src-vba")
    ap.add_argument("paths", nargs="*", help="konkretni fajlovi (podrazumevano ceo src-vba/)")
    ap.add_argument("--hook", action="store_true", help="bez izlaza kad je cisto")
    ap.add_argument("--self-test", action="store_true",
                    help="dokazi da provere zaista grizu (ne cita src-vba)")
    args = ap.parse_args(argv)

    if args.self_test:
        return self_test()

    files = vba_files(args.paths)
    if not files:
        return 0

    findings: list[Finding] = []
    publics: dict[str, list[tuple[str, int]]] = defaultdict(list)

    # Definicije se UVEK skupljaju nad celim src-vba, i kad se proverava jedan
    # fajl (hook) -- inace bi svaki poziv van tog fajla izgledao nedefinisano.
    defined = collect_definitions(vba_files([]))
    arities = collect_arities(vba_files([]))

    for path in files:
        with open(path, "rb") as fh:
            raw = fh.read()
        findings += check_ascii(path, raw)

        lines = raw.decode("ascii", errors="replace").replace("\r\n", "\n").split("\n")
        findings += check_decl_after_proc(path, lines)
        findings += check_reserved(path, lines)
        findings += check_undefined(path, lines, defined, arities)
        findings += check_local_dupes(path, lines)
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
                # ugovor ekrana -- kasno vezan, uvek kvalifikovan (v. SCR_UGOVOR)
                if name in SCR_UGOVOR and all(je_ekranski_modul(p) for p, _ in sites):
                    continue
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
