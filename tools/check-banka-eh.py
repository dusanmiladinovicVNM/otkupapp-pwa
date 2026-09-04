#!/usr/bin/env python3
"""
check-banka-eh.py -- staticka provera AUD-054 obrasca na RF-09 (banka) putanji.

Povod (RF-09, review runda 5): `ImportBankaInbox_WithDrivePull` je u EH bloku
radio

    EH:
        LogErr SRC
        Err.Raise Err.Number, SRC, Err.description

`LogErr` interno ide kroz `On Error Resume Next` / `On Error GoTo 0`, sto RESETUJE
`Err`. Citanje `Err.Number` / `Err.description` POSLE njega vraca prazan (ili tudji)
uzrok, pa uvoz koji je uredno rollback-ovan na vrhu lanca gubi ZASTO je pao -- bas
ono ponasanje koje RF-09 zatvara.

Ova skripta je jeftina kapija koja hvata povratak obrasca. Trazi u banka modulima:
  (a) `Err.Raise Err.Number` / `Err.Raise Err.number, ...` bilo gde (uzrok se cita
      iz zive `Err` umesto iz uhvacene kopije);
  (b) citanje `Err.Number` / `Err.description` / `Err.Source` POSLE poziva `LogErr`
      unutar istog EH bloka;
  (c) `Err.Raise` dok je AKTIVAN `On Error Resume Next` -- takav raise se TIHO
      PROGUTA i procedura se normalno vrati pozivaocu.

(c) je dodato u review rundi 6: `AutoMapAllBankaImport_TX` je u EH bloku imao
`On Error Resume Next` (za best-effort logovanje/monitoring), pa `Err.Raise` na
kraju nije radio nista -- funkcija je vracala 0 i forma je opet prikazivala
"Automatski mapirano: 0" nad ponistenim batch-om. Provera (a)+(b) to nije videla,
pa je checker davao lazno zeleno na zivom blocker-u.

CI ne pokrece Excel, pa se ovakva regresija inace vidi tek kad operater dobije
gresku bez uzroka.

Izlaz: 0 = cisto, 1 = nadjen bar jedan problem (lista sa fajlom i linijom).
"""

import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
SRC = REPO / "src-vba"

# RF-09 teren: import + mapiranje + ekran.
#
# frmBankaImport.frm je otisao u koraku 6 (docs/UI_MIGRACIJA_KATALOG.md par.27.14),
# a njegovo mesto je uzeo modScrBankaUvoz: tamo je sada i rucno mapiranje i kuka
# ulaska koja pokrece uvoz, pa se ista EH disciplina trazi od njega.
FILES = [
    "modBankaImport.bas",
    "modBankaMapiranje.bas",
    "modParse.bas",
    "modScrBankaUvoz.bas",
    "modTestBanka.bas",
]

PROC_START = re.compile(
    r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?"
    r"(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)",
    re.IGNORECASE,
)
PROC_END = re.compile(r"^\s*End\s+(?:Sub|Function|Property)\b", re.IGNORECASE)
RAISE_LIVE_ERR = re.compile(r"Err\.Raise\s+Err\.Number", re.IGNORECASE)
LOGERR_CALL = re.compile(r"^\s*(?:Call\s+)?LogErr\b", re.IGNORECASE)
ERR_READ = re.compile(r"Err\.(Number|Description|Source)\b", re.IGNORECASE)
ERR_CLEAR = re.compile(r"^\s*Err\.Clear\b", re.IGNORECASE)
# `On Error Resume Next` + citanje Err odmah posle poziva je legitiman obrazac
# (npr. best-effort Drive pull), pa se resetuje "posle LogErr" stanje.
ON_ERROR = re.compile(r"^\s*On\s+Error\b", re.IGNORECASE)
ON_ERROR_RESUME = re.compile(r"^\s*On\s+Error\s+Resume\s+Next\b", re.IGNORECASE)
ON_ERROR_GOTO = re.compile(r"^\s*On\s+Error\s+GoTo\s+(\S+)", re.IGNORECASE)
ERR_RAISE = re.compile(r"\bErr\.Raise\b", re.IGNORECASE)


def strip_comment(line: str) -> str:
    out, in_str = [], False
    for ch in line:
        if ch == '"':
            in_str = not in_str
        if ch == "'" and not in_str:
            break
        out.append(ch)
    return "".join(out)


def check_file(path: Path):
    problems = []
    proc = "(module level)"
    seen_logerr = False
    resume_next = False   # da li je `On Error Resume Next` trenutno aktivan

    for lineno, raw in enumerate(path.read_text(encoding="latin-1").splitlines(), 1):
        code = strip_comment(raw)
        if not code.strip():
            continue

        m = PROC_START.match(code)
        if m:
            proc = m.group(1)
            seen_logerr = False
            resume_next = False
        elif PROC_END.match(code):
            proc = "(module level)"
            seen_logerr = False
            resume_next = False

        if RAISE_LIVE_ERR.search(code):
            problems.append(
                (lineno, proc, "Err.Raise cita ZIVI Err.Number "
                                "(uhvati errNum/errDesc/errSrc PRE logovanja)")
            )

        # `Err.Raise` pod aktivnim Resume Next se tiho proguta -> pozivalac
        # dobija "uspesan" povratak umesto greske.
        if ERR_RAISE.search(code) and resume_next:
            problems.append(
                (lineno, proc, "Err.Raise pod aktivnim On Error Resume Next "
                                "(raise se TIHO GUTA -- dodaj On Error GoTo 0 pre njega)")
            )

        if ON_ERROR_RESUME.match(code):
            resume_next = True
            seen_logerr = False
            continue

        if ON_ERROR_GOTO.match(code):
            resume_next = False
            seen_logerr = False
            continue

        if LOGERR_CALL.match(code):
            seen_logerr = True
            continue

        # `On Error ...` (ostali oblici) / `Err.Clear` zatvaraju prethodni EH kontekst.
        if ON_ERROR.match(code) or ERR_CLEAR.match(code):
            seen_logerr = False
            continue

        if seen_logerr and ERR_READ.search(code):
            problems.append(
                (lineno, proc, "citanje Err.* POSLE LogErr "
                                "(LogErr resetuje Err -> uzrok se gubi)")
            )

    return problems


def scan_all_swallowed_raise():
    """Informativno: `Err.Raise` pod Resume Next u CELOM src-vba.

    Ovaj obrazac je uvek bug (raise se guta), za razliku od (a)/(b) koji su deo
    sireg AUD-054 duga van RF-09 terena. Zato se skenira siroko, ali NE obara
    exit kod -- kapija ostaje RF-09 teren.
    """
    found = []

    for path in sorted(SRC.glob("*.bas")) + sorted(SRC.glob("*.frm")) + sorted(SRC.glob("*.cls")):
        if path.name in FILES:
            continue
        for lineno, proc, msg in check_file(path):
            if "TIHO GUTA" in msg:
                found.append((path.name, lineno, proc))

    return found


def main() -> int:
    total = 0

    for name in FILES:
        path = SRC / name
        if not path.exists():
            print(f"SKIP  {name} (ne postoji)")
            continue

        problems = check_file(path)
        if not problems:
            print(f"OK    {name}")
            continue

        total += len(problems)
        for lineno, proc, msg in problems:
            print(f"FAIL  {name}:{lineno}  [{proc}]  {msg}")

    # Sirok pregled samo za "progutan raise" (uvek bug, ne samo na RF-09 terenu).
    wide = scan_all_swallowed_raise()
    if wide:
        print()
        print("INFO  progutan Err.Raise van RF-09 terena (nije kapija ovog paketa):")
        for name, lineno, proc in wide:
            print(f"INFO    {name}:{lineno}  [{proc}]")

    print()
    if total:
        print(f"AUD-054 obrazac: {total} problem(a) na banka putanji.")
        return 1

    print("AUD-054 obrazac: cisto na banka putanji.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
