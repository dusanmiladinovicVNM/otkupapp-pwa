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
  9. ZAKLONJENO   -- lokalni skalar (`Dim poruka As String`) koji se zove sa
                     zagradom. VBA to cita kao indeksiranje -> "Expected array".
                     Tipicno kad ime zaklanja istoimenu funkciju (`Poruka()`).
 10. MRTAV_LOG   -- `LogErr` posle `On Error` koje je vec obrisalo `Err`, pa
                     poziv ne upisuje NISTA u log.
 12. KOPIJA_NIZA -- `ByVal` na parametru koji se cita kao 2D niz. VBA kopira CEO
                     niz pri svakom pozivu; kod citaca po celiji to je kopija
                     tabele po procitanom polju (mereno: 1.8 ms po redu).
 11. ODSECEN     -- prazan fajl ili fajl bez `Attribute VB_Name`. Nije modul
                     nego ostatak neuspelog upisa; do sada je prolazio kao cist.

Provere 6 i 7 pokrivaju dve najcesce compile greske u ovom projektu -- one zbog
kojih je i pravljen headless compile gate koji se nije dao ukrotiti
(docs/EXCEL_TEST_HARNESS.md). Ovde se hvataju bez Excela, u milisekundama.
Ne pokrivaju: tipove, nedeklarisane promenljive, greske u .frm/.cls.
"""

from __future__ import annotations

import argparse
import contextlib
import io
import os
import re
import shutil
import sys
import tempfile
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
    "scr_naslovdopuna", "scr_brojac", "scr_cipovi",
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


def _logical_lines(lines: list[str]) -> list[tuple[str, int]]:
    """Spoji VBA nastavke reda (` _`) u jednu logicku liniju.

    Vraca (tekst, broj PRVOG fizickog reda) -- nalaz se prijavljuje na redu na
    kome poziv POCINJE, jer tamo operater i trazi.

    Postoji zato sto je provera arnosti preskakala svaki prelomljen poziv:
    `line.rstrip().endswith("_")` je znacio "ne gledaj". Bas tako je prosao
    RowsAktivni sa 8 argumenata za 9 parametara -- greska koja se videla tek
    kao [break] u VBE-u.
    """
    out: list[tuple[str, int]] = []
    i = 0
    while i < len(lines):
        prvi = i + 1
        tekst = lines[i].rstrip()
        while tekst.endswith("_") and i + 1 < len(lines):
            tekst = tekst[:-1].rstrip() + " " + lines[i + 1].strip()
            i += 1
        out.append((tekst, prvi))
        i += 1
    return out


def _split_statements(text: str) -> list[tuple[str, bool]]:
    """Podeli po dvotackama koje NISU u stringu ili zagradi.

    Vraca (naredba, je_labela). VBA dvotacka radi dve stvari: razdvaja naredbe
    (`Case 26: Foo`, `fokus = "x": Exit Function`) i zavrsava labelu (`EH:`).
    Ranije je checker preskakao SVAKU liniju sa dvotackom, pa mu je izmakao
    svaki poziv iza `Case N:` -- ukljucujuci ceo registar testova u modTest.
    """
    parts: list[tuple[str, bool]] = []
    depth, in_str, cur = 0, False, ""
    for i, ch in enumerate(text):
        if ch == '"':
            in_str = not in_str
        elif not in_str and ch in "([":
            depth += 1
        elif not in_str and ch in ")]":
            depth -= 1
        # `:=` je IMENOVAN ARGUMENT, ne kraj naredbe. Bez ove provere se
        # `Monitor_Error moduleName:="x", procedureName:="y"` cepa na prvom
        # dvotackom i ostane poziv sa jednim argumentom -- 101 lazan nalaz.
        elif not in_str and ch == ":" and depth == 0 and text[i + 1:i + 2] != "=":
            parts.append((cur.strip(), True))
            cur = ""
            continue
        cur += ch
    parts.append((cur.strip(), False))
    return [(p, lab) for p, lab in parts if p]


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
            # SAMO PRVA NAREDBA REDA.
            # `Function CLR_ERROR() As Long: CLR_ERROR = RGB(1,2,3): End Function`
            # je jedan red sa tri naredbe; bez ovoga bi lista parametara progutala
            # i argumente RGB-a, pa bi arnost bila 3 umesto 0 -- i svaki poziv
            # CLR() bio prijavljen kao pogresna arnost.
            glava = _split_statements(line)[0][0] if _split_statements(line) else line
            m = PROC_DEF.match(glava) or DECLARE_DEF.match(glava)
            if m and "(" in glava:
                params = _split_top_level(
                    glava[glava.index("(") + 1:glava.rindex(")")] if ")" in glava else "")
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

    for line, i in _logical_lines(lines):
        if BLOCK_OPEN.match(line):
            block_depth += 1
            continue
        if BLOCK_CLOSE.match(line):
            block_depth = max(0, block_depth - 1)
            continue
        if block_depth:
            continue

        cela = _strip_comment(line.strip())
        if not cela or cela.startswith("'") or cela.startswith("#"):
            continue        # komentar ili uslovna kompilacija
        if PROC_DEF.match(line) or DECLARE_DEF.match(line):
            continue

        for stmt, je_labela in _split_statements(cela):
            # `EH:` je labela, ne poziv -- bez nje bi svaki EH blok u projektu
            # bio prijavljen kao poziv nedefinisane procedure.
            if je_labela and re.fullmatch(r"[A-Za-z_]\w*", stmt):
                continue
            out += _proveri_naredbu(path, i, line, stmt, defined, arities)
    return out


def _proveri_naredbu(path: str, i: int, line: str, stmt: str,
                     defined: set[str], arities: dict) -> list[Finding]:
        out: list[Finding] = []
        explicit_call = bool(re.match(r"^Call\s", stmt, re.IGNORECASE))
        m = CALL_STMT.match(stmt)
        if not m:
            return out
        name, rest = m.group(1), m.group(2).lstrip()

        if name.lower() in STMT_WORDS or name.lower() in RESERVED:
            return out
        # `Foo = 1` (dodela), `Foo.Bar` (clan), `Foo As Long` (clan tipa) --
        # nista od toga nije poziv procedure.
        if rest.startswith(("=", ".", "!")) or re.match(r"^As\s", rest, re.IGNORECASE):
            return out
        # `Foo(kljuc).Add x` / `Foo(i) = 1` -- indeksiranje kolekcije ili niza.
        # Bez `Call` prefiksa, ime sa zagradom na pocetku naredbe je u ovom
        # kodu uvek indeks, ne poziv. (Sve 8 prvih laznih nalaza bilo je ovo.)
        if rest.startswith("(") and not explicit_call:
            return out
        if name.lower() not in defined:
            out.append(Finding(path, i, "NEDEFINISAN",
                               f"poziv '{name}' -- nigde u src-vba nije definisan "
                               f'Sub/Function/Property. VBA: "Sub or Function not defined".'))
            return out

        # Arnost -- druga polovina istog compile problema ("Wrong number of
        # arguments"). Proverava se samo kad je poziv cela naredba u jednoj
        # liniji i kad je ime jednoznacno definisano.
        span = arities.get(name.lower())
        if span is None:
            return out
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


# --- ZAKLONJENO: skalar koji se zove sa zagradom -----------------------------
#
# Zatecen incident (12 poziva u dve procedure, ziveli od v6-ui-119 do v6-ui-141):
#
#     Public Function StornoIzvrsi(..., ByRef poruka As String, ...)
#         ...
#         poruka = Poruka("STORNO_MSG_ZBIRNA_PRIJ")     ' Expected array
#
# VBA je case-insensitive, pa lokalno ime `poruka` zaklanja funkciju `Poruka()`.
# Poziv unutar te procedure zato NIJE poziv funkcije nego indeksiranje String
# promenljive -- compile error "Expected array".
#
# Zasto to nista drugo nije videlo:
#   - suite: VBA kompajlira proceduru TEK KAD SE POZOVE, a te dve je zvao samo UI;
#   - ARNOST/NEDEFINISAN: poziv je u poziciji IZRAZA (x = Foo(...)), sto je namerno
#     neproveravano (v. gore -- 406 laznih nalaza);
#   - CI: ne pokrece Excel.
#   Ostao je samo rucni Debug > Compile, i nasao ih je operater.
#
# Pravilo je uze od "ime zaklanja funkciju" i time sigurno: skalar EKSPLICITNOG
# tipa se u VBA ne moze indeksirati NIKAKO, pa je `ime(` uz `Dim ime As String`
# uvek greska -- nezavisno od toga da li nesto zaklanja.
#
# Namerno IZOSTAVLJENO, jer bi davalo lazne nalaze:
#   Variant            -- moze da nosi niz, pa je v(1) legalno
#   nizovi             -- Dim a(0 To 3) As String, a(1) je legalno
#   objekti            -- Dim d As Object, d("k") je default member
#   string literali    -- 14 od prvih 20 nalaza ovog obrasca bilo je ime unutar
#                         teksta ("...bez OtkupID (dokument: ...")
#   komentari          -- isto
SKALARNI_TIP = re.compile(
    r"^(String|Long|Integer|Double|Boolean|Date|Currency|Single|Byte|"
    r"Vb[A-Za-z]\w*|LongLong|LongPtr)$", re.IGNORECASE)

PROC_END = re.compile(r"^\s*End\s+(?:Sub|Function|Property)\b", re.IGNORECASE)
PARAM_DECL = re.compile(
    r"(?:^|[(,])\s*(?:Optional\s+)?(?:ByVal\s+|ByRef\s+)?(\w+)\s+As\s+(\w+)",
    re.IGNORECASE)
LOCAL_DECL = re.compile(r"^(?:Dim|Static)\s+(.*)$", re.IGNORECASE)
TIPIZOVANA_DEKL = re.compile(r"^(\w+)\s*(\([^)]*\))?\s+As\s+(\w+)", re.IGNORECASE)


def _strip_strings(text: str) -> str:
    """Isprazni string literale (sadrzaj, ne navodnike)."""
    out, in_str = [], False
    for ch in text:
        if ch == '"':
            in_str = not in_str
            out.append(ch)
        elif not in_str:
            out.append(ch)
    return "".join(out)


def _clean(text: str) -> str:
    return _strip_comment(_strip_strings(text))


def _scalar_names(header: str, body: list[str]) -> dict[str, str]:
    """ime -> tip, za parametre i lokalne skalare EKSPLICITNOG tipa."""
    names: dict[str, str] = {}
    inner = header[header.find("(") + 1:header.rfind(")")] if "(" in header else ""
    for m in PARAM_DECL.finditer(inner):
        names[m.group(1).lower()] = m.group(2)
    for line in body:
        m = LOCAL_DECL.match(_clean(line).strip())
        if not m:
            continue
        # Jedan Dim red nosi vise deklaracija. Citanje samo PRVE je tacno ono
        # zbog cega je drugi nalaz (StornoRedF8) prvi put promasen:
        #     Dim razlog As String, poruka As String, odg As VbMsgBoxResult
        for part in _split_top_level(m.group(1)):
            d = TIPIZOVANA_DEKL.match(part)
            if d and not d.group(2):
                names[d.group(1).lower()] = d.group(3)
    return {n: t for n, t in names.items() if SKALARNI_TIP.match(t)}


def check_scalar_call(path: str, lines: list[str]) -> list[Finding]:
    out: list[Finding] = []
    i, n = 0, len(lines)
    while i < n:
        m = PROC_DEF.match(lines[i])
        if not m:
            i += 1
            continue
        proc = m.group(1)
        # Zaglavlje se moze lomiti kroz ` _`, a parametri su u njemu.
        header, j = lines[i].rstrip(), i
        while header.endswith("_") and j + 1 < n:
            j += 1
            header = header[:-1] + " " + lines[j].strip()
            header = header.rstrip()
        k = j + 1
        while k < n and not PROC_END.match(lines[k]) and not PROC_DEF.match(lines[k]):
            k += 1
        body = lines[j + 1:k]
        for name, typ in _scalar_names(_strip_strings(header), body).items():
            hit = re.compile(r"(?<![\w.$])" + re.escape(name) + r"\s*\(", re.IGNORECASE)
            for off, line in enumerate(body):
                if hit.search(_clean(line)):
                    out.append(Finding(
                        path, j + 2 + off, "ZAKLONJENO",
                        f"'{name} As {typ}' je skalar, a zove se sa zagradom u '{proc}' "
                        f'-- VBA to cita kao indeksiranje: compile error "Expected array". '
                        f"Ako je ciljana istoimena procedura, pozovi je KVALIFIKOVANO "
                        f"(modPoruke.Poruka(...))."))
        i = k
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


# --- jedna putanja za sve provere nad jednim fajlom ---------------------------
#
# Postoji da bi self-test isao KROZ NJU, a ne pored nje. Da self-test zove
# check_local_dupes direktno, dokazivao bi samo da funkcija ume da nadje duplikat
# -- ne i da je CLI zaista zove. Brisanje jednog reda iz main() tada ostavlja i
# repo-run i self-test zelene, a checker prakticno iskljucen. To je ista klasa
# greske kao placebo test: zeleno, ali nije prikljuceno na produkcionu putanju.


# --- MRTAV_LOG: LogErr posle On Error koje je vec obrisalo Err ----------------
#
# `LogErr` pise SAMO kad je `Err.Number <> 0`. Svaka `On Error` naredba resetuje
# `Err` (dokazano modTest-om 68). Zato je ovaj redosled unutar handlera nem:
#
#     EH:
#         errDesc = Err.Description
#         On Error Resume Next               <- Err vise nije postavljen
#         LogErr "SaveOtpremnicaMulti_TX"     <- vidi 0, ne pise NISTA
#
# Osamdeset sedam takvih poziva je zivelo u kodu. Posledica je bila pad upisa
# BEZ IJEDNE linije u logu, pa se pravi uzrok (schema drift) trazio satima.
# Ispravno je pozvati LogErr PRE `On Error`, ili poslati opis izricito preko
# LogError SRC, errDesc, errNum.
#
# Provera gleda SAMO unutar handlera (posle labele EH:/ErrHandler:/Fin:/VRATI:).
# Van njega je `On Error Resume Next` legitimna priprema pred poziv koji sme da
# pukne -- tamo je Err posle poziva jos ziv i LogErr uredno pise.
# IGNORECASE zato sto je VBA case-insensitive: `eh:` i `EH:` su isti program, pa
# bi provera koja vidi samo drugi oblik cutala nad prvim -- checker zelen, a
# citava jedna legitimna sintaksa ga zaobilazi.
HANDLER_LABELA = re.compile(r"^(EH|ErrHandler|Fin|VRATI)\w*:\s*$",
                            re.IGNORECASE)
KRAJ_PROC = re.compile(r"^(Exit (Sub|Function|Property)|End (Sub|Function|Property))\b",
                       re.IGNORECASE)
ON_ERROR_BRISE = ("on error resume next", "on error goto 0")


def check_dead_log(path: str, lines: list[str]) -> list[Finding]:
    out = []
    u_handleru = False
    brisac = None
    for i, raw_line in enumerate(lines):
        t = _strip_comment(raw_line).strip()
        if not t:
            continue
        if HANDLER_LABELA.match(t):
            u_handleru, brisac = True, None
            continue
        if KRAJ_PROC.match(t):
            u_handleru, brisac = False, None
            continue
        if u_handleru and t.lower() in ON_ERROR_BRISE:
            brisac = t
            continue
        if u_handleru and brisac and re.match(r"^LogErr\b", t, re.IGNORECASE):
            out.append(Finding(
                path, i + 1, "MRTAV_LOG",
                "LogErr posle '%s' -- ta naredba resetuje Err, pa LogErr "
                "(koji pise samo kad je Err.Number <> 0) ne upisuje NISTA. "
                "Pozovi LogErr PRE nje, ili posalji opis izricito: "
                "LogError SRC, errDesc, errNum." % brisac))
            brisac = None
    return out

# --- ODSECEN: izvor bez zaglavlja koje VBA export UVEK pise -------------------
#
# `Attribute VB_Name = "..."` nosi SVAKI izvoz iz VBE-a -- svih 191 fajlova u
# src-vba, bez ijednog izuzetka. Fajl bez tog reda nije VBA modul nego ostatak
# neuspelog upisa.
#
# Rupa nije teorijska. Obrazac
#
#     io.open(P, "wb").write(s.encode("ascii"))
#
# otvara fajl PRE nego sto `encode` pukne: `open("wb")` je vec odsekao na nula
# bajtova kad izuzetak stigne. Tri puta je tako ostao prazan `.bas` -- i checker
# je svaki put prijavio CISTO, jer prazan fajl nema sta da prekrsi. Zelen izlaz
# nad izbrisanim modulom je gori od crvenog: `ImportAllVBA` ga uveze kao prazan
# i sve sto je u njemu bilo nestane, bez ijedne poruke.
#
# Bezbedan upis je `data = s.encode(...)` PA `io.open(P, "wb").write(data)`.
VB_NAME = re.compile(r'^Attribute VB_Name = "')


def check_truncated(path: str, raw: bytes, lines: list[str]) -> list[Finding]:
    if not raw.strip():
        return [Finding(path, 1, "ODSECEN",
                        "fajl je prazan. Najcesci uzrok: upis koji radi "
                        "open(P, 'wb') PRE nego sto encode pukne -- otvaranje vec "
                        "odsece fajl. Prvo `data = s.encode(...)`, pa upis.")]
    if not any(VB_NAME.match(ln) for ln in lines):
        return [Finding(path, 1, "ODSECEN",
                        "nema reda 'Attribute VB_Name = ...'. Svaki izvoz iz VBE-a "
                        "ga nosi, pa je ovo odsecen fajl, ne modul.")]
    return []

# --- KOPIJA_NIZA: `ByVal` na parametru koji se koristi kao 2D niz ---------------
#
# `ByVal` na `Variant`-u koji SADRZI niz tera VBA da kopira CEO niz pri svakom
# pozivu. Kad je procedura citac PO CELIJI -- a takve su sve pomocne rutine nad
# `GetTableData` nizom -- to je kopija cele tabele po procitanom polju.
#
# Mereno na terenu (v6-ui-147), `modPaletniList.SafeCell`:
#
#     1063 stavki, 1918 ms: citanje tabele 0, prolaz kroz stavke 1883 ms
#
# to jest 1.8 ms po redu za citanje dva polja iz niza koji je VEC u memoriji.
# Jedanaest takvih citaca je zivelo u kodu (v6-ui-157); provera postoji da
# dvanaesti ne bi usao neprimecen.
#
# NAMERNO USKA, iz dva razloga:
#
#   1. Trazi se DVOINDEKSNI pristup (`a(r, c)`). Jednoindeksni je isti trosak, ali
#      `Split()` rezultat, kolekcija i default-member poziv izgledaju isto, pa bi
#      lazni nalazi bili cesci od pravih -- ista odluka kao kod ARNOST-a.
#   2. Procedura koja u niz PISE mora ostati ByVal, inace bi menjala pozivaocev
#      niz. Zato se preskace svako telo koje sadrzi upis (`a = `, `a(...) = `,
#      `ReDim a`, `Erase a`). To je granica koja deli citaca od radnika.
#   3. Trazi se telo BEZ PETLJE i kratko (do CITAC_MAX_REDOVA naredbi). Rutina koja
#      niz primi pa ga sama iterira placa JEDNU kopiju za ceo prolaz -- to je
#      zanemarljivo i nije predmet ove provere. Skupo je samo ono sto se zove PO
#      CELIJI, a takva rutina nema svoju petlju i stane u nekoliko redova.
#      Bez ovog suzavanja provera je nad zatecenim kodom dala 51 nalaz, od cega
#      vecina bezopasnih -- a lazan nalaz uci da se checker ignorise.
CITAC_MAX_REDOVA = 12
PETLJA = re.compile(r"^(For|Do|While)\b", re.IGNORECASE)
#
# Potpis se sklapa preko nastavaka reda (`_`): visered potpis je cest bas kod ovih
# rutina (v. modStornoDok.KolicinaReda).
PARAM_BYVAL = re.compile(r"ByVal\s+(\w+)\s+As\s+Variant\b", re.IGNORECASE)


def _potpis_sa_nastavcima(lines: list[str], i: int) -> tuple[str, int]:
    """Ceo potpis od reda i, preko nastavaka `_`. Vraca (tekst, poslednji red)."""
    tekst = _strip_comment(lines[i]).rstrip()
    while tekst.endswith("_") and i + 1 < len(lines):
        i += 1
        tekst = tekst[:-1].rstrip() + " " + _strip_comment(lines[i]).strip()
    return tekst, i


def _telo_pise_u(ime: str, telo: list[str]) -> bool:
    """Da li telo MENJA parametar -- tada je ByVal moguce namerno, pa se ne dira."""
    upis = re.compile(r"^" + re.escape(ime) + r"\s*(\(.*\))?\s*=(?!=)", re.IGNORECASE)
    redim = re.compile(r"^(ReDim|Erase)\b.*\b" + re.escape(ime) + r"\b", re.IGNORECASE)
    for ln in telo:
        t = _strip_comment(ln).strip()
        if t and (upis.match(t) or redim.match(t)):
            return True
    return False


def check_array_copy(path: str, lines: list[str]) -> list[Finding]:
    out: list[Finding] = []
    i, n = 0, len(lines)
    while i < n:
        if not PROC_START.match(_strip_comment(lines[i])):
            i += 1
            continue
        potpis, kraj = _potpis_sa_nastavcima(lines, i)
        imena = PARAM_BYVAL.findall(potpis)
        if not imena:
            i = kraj + 1
            continue
        j = kraj + 1
        while j < n and not PROC_END.match(lines[j]):
            j += 1
        telo = lines[kraj + 1:j]
        naredbe = [t for t in (_strip_comment(x).strip() for x in telo) if t]
        if len(naredbe) > CITAC_MAX_REDOVA:
            i = j + 1
            continue
        if any(PETLJA.match(t) for t in naredbe):
            i = j + 1
            continue
        for ime in imena:
            dvoindeksni = re.compile(
                r"\b" + re.escape(ime) + r"\s*\([^()]*,[^()]*\)", re.IGNORECASE)
            if not any(dvoindeksni.search(_strip_comment(ln)) for ln in telo):
                continue
            if _telo_pise_u(ime, telo):
                continue
            out.append(Finding(
                path, i + 1, "KOPIJA_NIZA",
                f"parametar '{ime}' je ByVal, a koristi se kao 2D niz. VBA tada kopira "
                f"CEO niz pri SVAKOM pozivu -- kod citaca po celiji to je kopija cele "
                f"tabele po procitanom polju (mereno: 1.8 ms po redu). Telo niz samo "
                f"cita, pa je ByRef bez ijedne posledice po ponasanje."))
        i = j + 1
    return out


# --- REGISTAR: tri rucno odrzavana spiska testova moraju biti isti -----------
#
# modTest nosi test u TRI odvojena registra:
#
#   RunAllTests   `RunOne 114`                      -- sta se izvrsava
#   TestName      `Case 114: TestName = "T_Ime"`    -- pod kojim imenom se pad prijavljuje
#   InvokeTest    `Case 114: T_Ime`                 -- sta se stvarno zove
#
# Odrzavaju se rukom i vec su se razisli pri rebase-u. Svaki razlaz je nem:
#
#   nema u InvokeTest  ->  RunOne broji test, nista se ne izvrsi, suite je ZELENA
#   nema u TestName    ->  pad se prijavi kao "T_Nepoznat_114"
#   nema u RunOne      ->  test postoji i prolazi, ali se nikad ne pusta
#   pogresno ime       ->  Case 114 zove telo testa 113; oba "prolaze"
#
# Provera se okida SADRZAJEM, ne imenom fajla: trazi sva tri obrasca. Fajl koji
# nema takav registar je ne vidi.

_REG_RUNONE = re.compile(r"^\s*RunOne\s+(\d+)\s*$")
_REG_IME = re.compile(r"^\s*Case\s+(\d+)\s*:\s*TestName\s*=\s*\"([^\"]+)\"")
_REG_POZIV = re.compile(r"^\s*Case\s+(\d+)\s*:\s*([A-Za-z_][A-Za-z0-9_]*)\s*$")


_Reg = dict[int, tuple[int, str]]


def _registar_blokovi(lines: list[str]) -> tuple[_Reg, _Reg, _Reg,
                                                 list[tuple[str, int, int]], bool]:
    """(RunOne, TestName, InvokeTest, duplikati, je_registar).

    Duplikat se NE gubi: prvo vidjenje ostaje, a duplikati se skupljaju posebno.

    `je_registar` prati STRUKTURU -- postojanje sve tri procedure -- a NE broj
    parsiranih unosa. Razlika je bila propust: dok se okidalo na "sva tri
    recnika su neprazna", registar kome je CEO InvokeTest ostao bez ijedne Case
    grane prolazio je kao "ovo nije registar" i provera se gasila. Dakle bas
    najgori slucaj -- 114 testova se broji, nijedan se ne izvrsi -- nije bio
    pokriven.
    """
    runone: _Reg = {}
    imena: _Reg = {}
    pozivi: _Reg = {}
    dupli: list[tuple[str, int, int]] = []

    ima_runall = False
    ima_imena = False
    ima_pozivi = False

    u_runall = False
    u_imenu = False
    u_pozivu = False

    for i, ln in enumerate(lines, 1):
        gol = _strip_comment(ln)

        if re.match(r"^\s*(Public|Private)?\s*Sub\s+RunAllTests\s*\(", gol, re.I):
            ima_runall = True
            u_runall = True
            continue
        if re.match(r"^\s*Private\s+Function\s+TestName\s*\(", gol, re.I):
            ima_imena = True
            u_imenu = True
            continue
        if re.match(r"^\s*Private\s+Sub\s+InvokeTest\s*\(", gol, re.I):
            ima_pozivi = True
            u_pozivu = True
            continue
        if re.match(r"^\s*End\s+(Function|Sub)\s*$", gol, re.I):
            u_runall = False
            u_imenu = False
            u_pozivu = False
            continue

        # RunOne se broji SAMO u RunAllTests. Van njega bi helper koji negde
        # drugde zove "RunOne 17" izgledao kao dupli unos registra.
        if u_runall:
            m = _REG_RUNONE.match(gol)
            if m:
                n = int(m.group(1))
                if n in runone:
                    dupli.append(("RunOne", n, i))
                else:
                    runone[n] = (i, "")
                continue

        if u_imenu:
            m = _REG_IME.match(gol)
            if m:
                n = int(m.group(1))
                if n in imena:
                    dupli.append(("TestName", n, i))
                else:
                    imena[n] = (i, m.group(2))
            continue

        if u_pozivu:
            m = _REG_POZIV.match(gol)
            if m:
                n = int(m.group(1))
                if n in pozivi:
                    dupli.append(("InvokeTest", n, i))
                else:
                    pozivi[n] = (i, m.group(2))

    return runone, imena, pozivi, dupli, (ima_runall and ima_imena and ima_pozivi)


def check_test_registry(path: str, lines: list[str]) -> list[Finding]:
    runone, imena, pozivi, dupli, je_registar = _registar_blokovi(lines)

    # Okida se na STRUKTURU, ne na broj unosa: fajl koji ima sve tri procedure
    # JESTE registar, pa i onda kad je neki od spiskova ostao prazan -- to je
    # bas slucaj koji se najvise isplati prijaviti.
    if not je_registar:
        return []

    out = []

    for koji, n, ln in dupli:
        out.append(Finding(path, ln, "REGISTAR",
                           f"{koji} ima indeks {n} DVAPUT. Drugi se tiho gubi "
                           f"(Select Case uzima prvi), a RunAllTests ga broji."))

    svi = set(runone) | set(imena) | set(pozivi)
    for n in sorted(svi):
        gde = []
        if n not in runone:
            gde.append("RunOne")
        if n not in imena:
            gde.append("TestName")
        if n not in pozivi:
            gde.append("InvokeTest")
        if not gde:
            continue
        # Linija bilo kog registra koji ga IMA -- da nalaz vodi na mesto.
        ln = next((v[0] for v in (runone.get(n), imena.get(n), pozivi.get(n)) if v), 1)
        out.append(Finding(path, ln, "REGISTAR",
                           f"Test {n} nedostaje u: {', '.join(gde)}. Tri registra "
                           f"moraju da nose ISTI skup indeksa."))

    # Rupa u numeraciji: registar je 1..max, bez preskakanja. Preskocen broj
    # znaci da je test obrisan a ostala tri traga -- ili da je dodat pogresan.
    if svi:
        rupe = [n for n in range(1, max(svi) + 1) if n not in svi]
        if rupe:
            out.append(Finding(path, 1, "REGISTAR",
                               f"Rupa u numeraciji testova: {rupe}. Registar ide "
                               f"1..{max(svi)} bez preskakanja."))

    # Najtisi razlaz: Case N zove telo TUDJEG testa. Oba "prolaze", a jedan se
    # nikad ne izvrsi -- ime u izvestaju pripada jednom, telo drugom.
    for n in sorted(set(imena) & set(pozivi)):
        ime = imena[n][1]
        poziv = pozivi[n][1]
        if ime.lower() != poziv.lower():
            out.append(Finding(path, pozivi[n][0], "REGISTAR",
                               f"Test {n}: TestName kaze '{ime}', a InvokeTest zove "
                               f"'{poziv}'. Pad bi se prijavio pod tudjim imenom."))

    return out


def check_file(path: str, raw: bytes, lines: list[str],
               defined: set[str], arities: dict[str, tuple[int, float]]) -> list[Finding]:
    out = []
    out += check_ascii(path, raw)
    out += check_decl_after_proc(path, lines)
    out += check_reserved(path, lines)
    out += check_undefined(path, lines, defined, arities)
    out += check_local_dupes(path, lines)
    out += check_scalar_call(path, lines)
    out += check_dead_log(path, lines)
    out += check_truncated(path, raw, lines)
    out += check_array_copy(path, lines)
    out += check_test_registry(path, lines)
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

# --- ZAKLONJENO: skalar zvan sa zagradom ---
# Slucajevi 1 i 2 su REKONSTRUKCIJA zatecenog incidenta, ne izmisljeni primeri:
# oba oblika su stvarno postojala u modStornoDok.StornoIzvrsi i
# modScrDokumenti.StornoRedF8 od v6-ui-119 do v6-ui-141.
ZAKLONJENO_CASES = [
    # --- mora da zapisti ---
    ("parametar zaklanja funkciju (zatecen incident)", 1, """Option Explicit
Public Function StornoIzvrsi(ByVal tip As String, ByRef poruka As String) As Boolean
    poruka = Poruka("STORNO_MSG_OK")
End Function
"""),
    ("druga deklaracija u Dim redu (drugi zatecen nalaz)", 1, """Option Explicit
Private Function StornoRedF8(ByVal red As Long) As Boolean
    Dim razlog As String, poruka As String, odg As VbMsgBoxResult
    ShowToast Poruka("OTKUI_ERR_NEMA_REDA"), True
End Function
"""),
    ("prelomljeno zaglavlje -- parametar u nastavku reda", 1, """Option Explicit
Public Sub Radi(ByVal a As Long, _
                ByRef poruka As String)
    poruka = Poruka("KLJUC")
End Sub
"""),
    ("Optional parametar", 1, """Option Explicit
Public Sub Radi(Optional ByVal poruka As String = "")
    poruka = Poruka("KLJUC")
End Sub
"""),

    # --- ne sme da zapisti (ova polovina je vaznija) ---
    ("ime unutar STRING literala", 0, """Option Explicit
Public Sub Radi()
    Dim otkupID As String
    LogError "SRC", "Otvoren blok bez OtkupID (dokument: x)."
End Sub
"""),
    ("ime u KOMENTARU", 0, """Option Explicit
Public Sub Radi()
    Dim poruka As String
    ' ovde se poruka("KLJUC") samo opisuje
End Sub
"""),
    ("niz sme da se indeksira", 0, """Option Explicit
Public Sub Radi()
    Dim polje(0 To 3) As String
    polje(1) = "x"
End Sub
"""),
    ("Variant moze da nosi niz", 0, """Option Explicit
Public Sub Radi()
    Dim v As Variant
    v = Array(1, 2)
    Debug.Print v(1)
End Sub
"""),
    ("objekat -- default member", 0, """Option Explicit
Public Sub Radi()
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("k") = 1
End Sub
"""),
    ("kvalifikovan poziv uz istoimenu promenljivu", 0, """Option Explicit
Public Function StornoIzvrsi(ByRef poruka As String) As Boolean
    poruka = modPoruke.Poruka("STORNO_MSG_OK")
End Function
"""),
    ("promenljiva se prosledjuje, ne zove", 0, """Option Explicit
Public Sub Radi()
    Dim poruka As String
    If Len(poruka) = 0 Then Debug.Print InStr(1, poruka, "x")
End Sub
"""),
    ("istoimena promenljiva u DRUGOJ proceduri", 0, """Option Explicit
Public Sub Prva()
    Dim poruka As String
    poruka = "x"
End Sub
Public Sub Druga()
    Debug.Print Poruka("KLJUC")
End Sub
"""),
]


# Slucajevi za REGISTAR. Ovaj checker se okida SADRZAJEM (trazi sva tri spiska),
# pa slucaj koji nije registar mora da prodje bez nalaza -- ta polovina je
# vaznija: lazan nalaz u PostToolUse hook-u uci da se checker ignorise.
REGISTAR_CASES = [
    ('CEO InvokeTest prazan', 3, """Option Explicit
Public Sub RunAllTests()
    RunOne 1
    RunOne 2
    RunOne 3
End Sub

Private Function TestName(ByVal idx As Long) As String
    Select Case idx
        Case 1: TestName = "T_Test1"
        Case 2: TestName = "T_Test2"
        Case 3: TestName = "T_Test3"
        Case Else: TestName = "T_Nepoznat_" & idx
    End Select
End Function

Private Sub InvokeTest(ByVal idx As Long)
    Select Case idx
    End Select
End Sub
"""),
    ('CEO TestName prazan', 3, """Option Explicit
Public Sub RunAllTests()
    RunOne 1
    RunOne 2
    RunOne 3
End Sub

Private Function TestName(ByVal idx As Long) As String
    Select Case idx
        Case Else: TestName = "T_Nepoznat_" & idx
    End Select
End Function

Private Sub InvokeTest(ByVal idx As Long)
    Select Case idx
        Case 1: T_Test1
        Case 2: T_Test2
        Case 3: T_Test3
    End Select
End Sub
"""),
    ('RunAllTests bez ijednog RunOne', 3, """Option Explicit
Public Sub RunAllTests()
End Sub

Private Function TestName(ByVal idx As Long) As String
    Select Case idx
        Case 1: TestName = "T_Test1"
        Case 2: TestName = "T_Test2"
        Case 3: TestName = "T_Test3"
        Case Else: TestName = "T_Nepoznat_" & idx
    End Select
End Function

Private Sub InvokeTest(ByVal idx As Long)
    Select Case idx
        Case 1: T_Test1
        Case 2: T_Test2
        Case 3: T_Test3
    End Select
End Sub
"""),
    ('RunOne van RunAllTests nije registar', 0, """Option Explicit
Public Sub RunAllTests()
    RunOne 1
    RunOne 2
    RunOne 3
End Sub

Private Function TestName(ByVal idx As Long) As String
    Select Case idx
        Case 1: TestName = "T_Test1"
        Case 2: TestName = "T_Test2"
        Case 3: TestName = "T_Test3"
        Case Else: TestName = "T_Nepoznat_" & idx
    End Select
End Function

Private Sub InvokeTest(ByVal idx As Long)
    Select Case idx
        Case 1: T_Test1
        Case 2: T_Test2
        Case 3: T_Test3
    End Select
End Sub

Private Sub Pomocna()
    RunOne 2
End Sub
"""),
    ('zdrav registar', 0, """Option Explicit
Public Sub RunAllTests()
    RunOne 1
    RunOne 2
    RunOne 3
End Sub

Private Function TestName(ByVal idx As Long) As String
    Select Case idx
        Case 1: TestName = "T_Test1"
        Case 2: TestName = "T_Test2"
        Case 3: TestName = "T_Test3"
        Case Else: TestName = "T_Nepoznat_" & idx
    End Select
End Function

Private Sub InvokeTest(ByVal idx As Long)
    Select Case idx
        Case 1: T_Test1
        Case 2: T_Test2
        Case 3: T_Test3
    End Select
End Sub
"""),
    ('nema u InvokeTest', 1, """Option Explicit
Public Sub RunAllTests()
    RunOne 1
    RunOne 2
    RunOne 3
End Sub

Private Function TestName(ByVal idx As Long) As String
    Select Case idx
        Case 1: TestName = "T_Test1"
        Case 2: TestName = "T_Test2"
        Case 3: TestName = "T_Test3"
        Case Else: TestName = "T_Nepoznat_" & idx
    End Select
End Function

Private Sub InvokeTest(ByVal idx As Long)
    Select Case idx
        Case 1: T_Test1
        Case 2: T_Test2
    End Select
End Sub
"""),
    ('nema u RunOne', 1, """Option Explicit
Public Sub RunAllTests()
    RunOne 1
    RunOne 2
End Sub

Private Function TestName(ByVal idx As Long) As String
    Select Case idx
        Case 1: TestName = "T_Test1"
        Case 2: TestName = "T_Test2"
        Case 3: TestName = "T_Test3"
        Case Else: TestName = "T_Nepoznat_" & idx
    End Select
End Function

Private Sub InvokeTest(ByVal idx As Long)
    Select Case idx
        Case 1: T_Test1
        Case 2: T_Test2
        Case 3: T_Test3
    End Select
End Sub
"""),
    ('nema u TestName', 1, """Option Explicit
Public Sub RunAllTests()
    RunOne 1
    RunOne 2
    RunOne 3
End Sub

Private Function TestName(ByVal idx As Long) As String
    Select Case idx
        Case 1: TestName = "T_Test1"
        Case 2: TestName = "T_Test2"
        Case Else: TestName = "T_Nepoznat_" & idx
    End Select
End Function

Private Sub InvokeTest(ByVal idx As Long)
    Select Case idx
        Case 1: T_Test1
        Case 2: T_Test2
        Case 3: T_Test3
    End Select
End Sub
"""),
    ('dupli RunOne', 1, """Option Explicit
Public Sub RunAllTests()
    RunOne 1
    RunOne 2
    RunOne 3
    RunOne 3
End Sub

Private Function TestName(ByVal idx As Long) As String
    Select Case idx
        Case 1: TestName = "T_Test1"
        Case 2: TestName = "T_Test2"
        Case 3: TestName = "T_Test3"
        Case Else: TestName = "T_Nepoznat_" & idx
    End Select
End Function

Private Sub InvokeTest(ByVal idx As Long)
    Select Case idx
        Case 1: T_Test1
        Case 2: T_Test2
        Case 3: T_Test3
    End Select
End Sub
"""),
    ('rupa u numeraciji', 1, """Option Explicit
Public Sub RunAllTests()
    RunOne 1
    RunOne 2
    RunOne 4
End Sub

Private Function TestName(ByVal idx As Long) As String
    Select Case idx
        Case 1: TestName = "T_Test1"
        Case 2: TestName = "T_Test2"
        Case 4: TestName = "T_Test4"
        Case Else: TestName = "T_Nepoznat_" & idx
    End Select
End Function

Private Sub InvokeTest(ByVal idx As Long)
    Select Case idx
        Case 1: T_Test1
        Case 2: T_Test2
        Case 4: T_Test4
    End Select
End Sub
"""),
    ('Case N zove tudje telo', 1, """Option Explicit
Public Sub RunAllTests()
    RunOne 1
    RunOne 2
    RunOne 3
End Sub

Private Function TestName(ByVal idx As Long) As String
    Select Case idx
        Case 1: TestName = "T_Test1"
        Case 2: TestName = "T_Test2"
        Case 3: TestName = "T_Test3"
        Case Else: TestName = "T_Nepoznat_" & idx
    End Select
End Function

Private Sub InvokeTest(ByVal idx As Long)
    Select Case idx
        Case 1: T_Test1
        Case 2: T_Test1
        Case 3: T_Test3
    End Select
End Sub
"""),
    ('modul bez registra', 0, """Option Explicit
Public Sub Radi()
    Select Case 1
        Case 1: Nesto
    End Select
End Sub
"""),
    ('RunOne bez Select Case registara', 0, """Option Explicit
Public Sub RunAllTests()
    RunOne 1
    RunOne 2
End Sub
"""),
]


def _dupli_nalazi(findings: list[Finding]) -> int:
    return sum(1 for f in findings if f.code == "DUPLIKAT_LOKALNI")


# Slucajevi za NEDEFINISAN / ARNOST. Definicije se zadaju uz slucaj, jer se
# inace skupljaju nad celim src-vba -- a ovi izvori tamo ne postoje.
#
# Sve cetiri "mora da zapisti" stavke su greske koje su STVARNO prosle checker
# u ovoj sesiji i videle se tek kao [break] u VBE-u ili kao visenje harnessa.
SELF_TEST_POZIVI = [
    # --- mora da zapisti ---
    ("poziv iza Case N:", {"NEDEFINISAN": 1}, {"radi"}, {"radi": (2, 2)}, """Option Explicit
Public Sub Registar(ByVal idx As Long)
    Select Case idx
        Case 1: Radi 1, 2
        Case 2: NePostojiNikako
    End Select
End Sub
"""),
    ("arnost prelomljenog poziva", {"ARNOST": 1}, {"radi"}, {"radi": (2, 2)}, """Option Explicit
Public Sub Prelomljen()
    Radi 1, _
         2, _
         3
End Sub
"""),
    ("poziv iza dvotacke u istom redu", {"NEDEFINISAN": 1}, {"radi"}, {"radi": (2, 2)}, """Option Explicit
Public Sub Dve()
    Radi 1, 2: NemaMeNigde
End Sub
"""),
    ("arnost poziva sa premalo argumenata", {"ARNOST": 1}, {"radi"}, {"radi": (2, 2)}, """Option Explicit
Public Sub Malo()
    Radi 1
End Sub
"""),
    # --- ne sme da zapisti ---
    ("labela EH: nije poziv", {}, {"radi"}, {"radi": (2, 2)}, """Option Explicit
Public Sub SaLabelom()
    On Error GoTo EH
    Radi 1, 2
    Exit Sub
EH:
    Radi 1, 2
End Sub
"""),
    ("imenovani argumenti (`:=`)", {}, {"radi"}, {"radi": (2, 2)}, """Option Explicit
Public Sub Imenovani()
    Radi a:=1, b:=2
End Sub
"""),
    ("prelomljen poziv sa tacnom arnoscu", {}, {"radi"}, {"radi": (2, 2)}, """Option Explicit
Public Sub Dobar()
    Radi 1, _
         2
End Sub
"""),
    ("clan Type bloka nije poziv", {}, {"radi"}, {"radi": (2, 2)}, """Option Explicit
Public Type TRed
    Naziv As String
End Type
"""),
]


# --- MRTAV_LOG: LogErr koji ne moze da zapise ---
#
# Dokaz u oba smera. Druga polovina (0 nalaza) je vaznija: `On Error Resume
# Next` PRE poziva koji sme da pukne je legitiman obrazac i cest u kodu --
# lazan nalaz nad njim bi naucio da se checker ignorise.
MRTAV_LOG_CASES = [
    # --- MRTAV_LOG: mora da zapisti ---
    ("LogErr posle Resume Next u handleru", 1, """Option Explicit
Public Sub Radi()
    On Error GoTo EH
    Exit Sub
EH:
    Dim errDesc As String
    errDesc = Err.Description
    On Error Resume Next
    LogErr "Radi"
End Sub
"""),
    ("LogErr posle On Error GoTo 0 u handleru", 1, """Option Explicit
Public Sub Radi()
    On Error GoTo EH
    Exit Sub
EH:
    On Error GoTo 0
    LogErr "Radi"
End Sub
"""),
    ("LogErr posle Resume Next, labela malim slovima", 1, """Option Explicit
Public Sub Radi()
    On Error GoTo eh
    Exit Sub
eh:
    On Error Resume Next
    LogErr "Radi"
End Sub
"""),
    # --- MRTAV_LOG: NE sme da zapisti ---
    ("LogErr PRE Resume Next u handleru", 0, """Option Explicit
Public Sub Radi()
    On Error GoTo EH
    Exit Sub
EH:
    LogErr "Radi"
    On Error Resume Next
End Sub
"""),
    ("Resume Next PRE poziva koji puca -- Err je posle njega ziv", 0, """Option Explicit
Public Sub Radi()
    On Error Resume Next
    MozdaPukne
    If Err.Number <> 0 Then
        LogErr "Radi"
        Err.Clear
    End If
End Sub
"""),
]


# --- ODSECEN: dokaz u oba smera ---------------------------------------------
#
# Druga polovina (0 nalaza) drzi granicu: provera sme da zapisti SAMO na fajlu
# bez zaglavlja. Minimalan legalan modul u ovom repou ima 154 bajta i nema
# `Option Explicit` (v. modMeteo.bas), pa nista strozije od VB_Name ne prolazi
# nad zatecenim izvorom.
ODSECEN_CASES = [
    # --- mora da zapisti ---
    ("prazan fajl", 1, ""),
    ("samo beline", 1, "   \r\n\r\n"),
    ("kod bez VB_Name zaglavlja", 1, """Option Explicit
Public Sub Radi()
End Sub
"""),
    # --- NE sme da zapisti ---
    ("minimalan modul, bez Option Explicit", 0, '''Attribute VB_Name = "modMeteo"
' === modMeteo ===
' TODO: Implementierung
'''),
    ("forma: VB_Name dolazi tek posle Begin bloka", 0, '''VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmX 
   Caption         =   "UserForm1"
End
Attribute VB_Name = "frmX"
Option Explicit
'''),
]

# --- KOPIJA_NIZA: dokaz u oba smera ------------------------------------------
#
# Druga polovina (0 nalaza) nosi celu tezinu. Provera je NAMERNO uska, a granice
# su bas ovi slucajevi: rutina koja u niz PISE mora ostati ByVal, rutina koja niz
# SAMA iterira placa jednu kopiju za ceo prolaz, a jednoindeksni pristup se ne
# razlikuje od Split() rezultata ili kolekcije.
#
# Bez tog suzavanja je provera nad zatecenim kodom dala 51 nalaz umesto 26.
KOPIJA_NIZA_CASES = [
    # --- mora da zapisti ---
    ("citac po celiji, ByVal 2D niz", 1, """Option Explicit
Private Function Celija(ByVal d As Variant, ByVal r As Long, ByVal c As Long) As String
    If c > 0 Then Celija = Trim$(CStr(d(r, c)))
End Function
"""),
    ("isto, potpis preko dva reda", 1, """Option Explicit
Private Function Kolicina(ByVal d As Variant, ByVal r As Long, _
                          ByVal c As Long) As Double
    If c > 0 Then Kolicina = NzD(d(r, c))
End Function
"""),
    # --- NE sme da zapisti ---
    ("vec je ByRef", 0, """Option Explicit
Private Function Celija(ByRef d As Variant, ByVal r As Long, ByVal c As Long) As String
    If c > 0 Then Celija = Trim$(CStr(d(r, c)))
End Function
"""),
    ("telo PISE u niz -- ByVal je tu namerno", 0, """Option Explicit
Private Function Ocisti(ByVal d As Variant, ByVal r As Long, ByVal c As Long) As Variant
    d(r, c) = ""
    Ocisti = d
End Function
"""),
    ("telo PREUZIMA niz na sebe", 0, """Option Explicit
Private Function Kopija(ByVal d As Variant) As Variant
    ReDim Preserve d(1 To 2, 1 To 2)
    Kopija = d(1, 1)
End Function
"""),
    ("rutina koja niz SAMA iterira placa jednu kopiju", 0, """Option Explicit
Private Sub Ispisi(ByVal d As Variant, ByVal n As Long)
    Dim r As Long
    For r = 1 To n
        Debug.Print d(r, 1)
    Next r
End Sub
"""),
    ("jednoindeksni pristup se NE prijavljuje", 0, """Option Explicit
Private Function Prvi(ByVal d As Variant) As String
    Prvi = CStr(d(0))
End Function
"""),
    ("skalarni Variant nije niz", 0, """Option Explicit
Private Function Tekst(ByVal v As Variant) As String
    Tekst = Trim$(CStr(v))
End Function
"""),
]


def self_test() -> int:
    palo = []

    # collect_arities nad jednolinijskom funkcijom: parametri se citaju do PARNE
    # zatvorene zagrade i samo iz PRVE naredbe reda. Ranije je uzimao poslednju
    # zagradu u redu, pa je `Function F() As Long: F = RGB(1,2,3): End Function`
    # dobijao arnost 3 umesto 0 -- i svaki poziv F() bi bio prijavljen.
    tmp1 = tempfile.mkdtemp(prefix="vbacheck_ar_")
    try:
        put1 = os.path.join(tmp1, "modJednolinijska.bas")
        with open(put1, "w", encoding="ascii", newline="\r\n") as fh:
            fh.write("Option Explicit\n"
                     "Public Function CLR() As Long: CLR = RGB(1, 2, 3): End Function\n")
        dobijeno_ar = collect_arities([put1]).get("clr")
        if dobijeno_ar != (0, 0):
            palo.append(f"  arnost jednolinijske funkcije: ocekivano (0, 0), "
                        f"dobijeno {dobijeno_ar}")
    finally:
        shutil.rmtree(tmp1, ignore_errors=True)

    for naziv, ocekivano, defined, arities, izvor in SELF_TEST_POZIVI:
        lines = izvor.replace("\r\n", "\n").split("\n")
        nalazi = check_undefined("<self-test>.bas", lines, defined, arities)
        dobijeno: dict[str, int] = defaultdict(int)
        for f in nalazi:
            dobijeno[f.code] += 1
        if dict(dobijeno) != ocekivano:
            palo.append(f"  {naziv}: ocekivano {ocekivano or '{}'}, dobijeno {dict(dobijeno) or '{}'}")

    # Sve kroz check_file -- istu funkciju koju zove main(). Ostale provere se
    # ne racunaju, ali se izvrsavaju: slucaj koji bi im pao rusio bi i CLI.
    for naziv, ocekivano, izvor in SELF_TEST_CASES:
        lines = izvor.replace("\r\n", "\n").split("\n")
        raw = izvor.encode("ascii")
        dobijeno = _dupli_nalazi(check_file("<self-test>", raw, lines, set(), {}))
        if dobijeno != ocekivano:
            palo.append(f"  {naziv}: ocekivano {ocekivano} nalaza, dobijeno {dobijeno}")

    for naziv, ocekivano, izvor in KOPIJA_NIZA_CASES:
        lines = izvor.replace("\r\n", "\n").split("\n")
        raw = izvor.encode("ascii")
        nalazi = check_file("<self-test>", raw, lines, set(), {})
        dobijeno = sum(1 for f in nalazi if f.code == "KOPIJA_NIZA")
        if dobijeno != ocekivano:
            palo.append(f"  KOPIJA_NIZA/{naziv}: ocekivano {ocekivano} nalaza, "
                        f"dobijeno {dobijeno}")

    for naziv, ocekivano, izvor in REGISTAR_CASES:
        lines = izvor.replace("\r\n", "\n").split("\n")
        raw = izvor.encode("ascii")
        nalazi = check_file("<self-test>", raw, lines, set(), {})
        dobijeno = sum(1 for f in nalazi if f.code == "REGISTAR")
        if dobijeno != ocekivano:
            palo.append(f"  REGISTAR/{naziv}: ocekivano {ocekivano} nalaza, "
                        f"dobijeno {dobijeno}")

    for naziv, ocekivano, izvor in ODSECEN_CASES:
        lines = izvor.replace("\r\n", "\n").split("\n")
        raw = izvor.encode("ascii")
        nalazi = check_file("<self-test>", raw, lines, set(), {})
        dobijeno = sum(1 for f in nalazi if f.code == "ODSECEN")
        if dobijeno != ocekivano:
            palo.append(f"  ODSECEN/{naziv}: ocekivano {ocekivano} nalaza, "
                        f"dobijeno {dobijeno}")

    for naziv, ocekivano, izvor in MRTAV_LOG_CASES:
        lines = izvor.replace("\r\n", "\n").split("\n")
        raw = izvor.encode("ascii")
        nalazi = check_file("<self-test>", raw, lines, set(), {})
        dobijeno = sum(1 for f in nalazi if f.code == "MRTAV_LOG")
        if dobijeno != ocekivano:
            palo.append(f"  {naziv}: ocekivano {ocekivano} MRTAV_LOG, "
                        f"dobijeno {dobijeno}")

    for naziv, ocekivano, izvor in ZAKLONJENO_CASES:
        lines = izvor.replace("\r\n", "\n").split("\n")
        raw = izvor.encode("ascii")
        nalazi = check_file("<self-test>", raw, lines, set(), {})
        dobijeno = sum(1 for f in nalazi if f.code == "ZAKLONJENO")
        if dobijeno != ocekivano:
            palo.append(f"  ZAKLONJENO/{naziv}: ocekivano {ocekivano} nalaza, "
                        f"dobijeno {dobijeno}")

    # I jedan slucaj kroz CEO CLI, nad pravim fajlom na disku. check_file dokazuje
    # da provera radi; ovo dokazuje da je CLI zaista zove i da vraca exit 2.
    naziv = "ceo CLI nad pravim .bas fajlom"
    tmp = tempfile.mkdtemp(prefix="vbacheck_")
    try:
        put = os.path.join(tmp, "modSelfTest.bas")
        with open(put, "w", encoding="ascii", newline="\r\n") as fh:
            fh.write(SELF_TEST_CASES[0][2])
        buf = io.StringIO()
        with contextlib.redirect_stderr(buf):
            rc = main([put])
        izlaz = buf.getvalue()
        if rc != 2:
            palo.append(f"  {naziv}: ocekivan exit 2, dobijen {rc}")
        elif "DUPLIKAT_LOKALNI" not in izlaz:
            palo.append(f"  {naziv}: exit 2 je stigao, ali ne od DUPLIKAT_LOKALNI "
                        f"-- CLI mozda vise ne zove tu proveru")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    ukupno = (len(SELF_TEST_CASES) + len(SELF_TEST_POZIVI) + len(ZAKLONJENO_CASES)
              + len(MRTAV_LOG_CASES) + len(ODSECEN_CASES) + len(KOPIJA_NIZA_CASES)
              + len(REGISTAR_CASES) + 2)
    for line in palo:
        print(line, file=sys.stderr)
    if palo:
        print(f"\nself-test: {len(palo)} od {ukupno} slucajeva palo.", file=sys.stderr)
        return 2
    print(f"self-test: cisto ({ukupno} slucajeva: DUPLIKAT_LOKALNI, "
          f"NEDEFINISAN/ARNOST, ZAKLONJENO, i jedan kroz ceo CLI).")
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
        lines = raw.decode("ascii", errors="replace").replace("\r\n", "\n").split("\n")
        findings += check_file(path, raw, lines, defined, arities)
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
