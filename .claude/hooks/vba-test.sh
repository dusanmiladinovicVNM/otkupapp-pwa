#!/usr/bin/env bash
# Stop hook: kad je u sesiji diran src-vba/, pusti BRZI set testova ponasanja.
#
# vba_check (PostToolUse) hvata sintaksu i compile-hard greske. Ovo hvata drugu
# klasu: izmenu koja se uredno kompajlira, a menja PONASANJE -- npr. brisanje
# datuma ili zbirne u ClearOtkupFields. Vidi .claude/rules/testovi.md.
#
# NAMERNO NIJE PUN GATE. Pun podrazumevani set je 11 suite-ova / ~1050 provera uz
# podizanje Excela; na svakom Stop-u je desktop sesiju cinio neupotrebljivom. Pun
# set je zato ostao NAMERNA komanda pred commit/release:
#
#     python tools\run_vba.py
#
# Exit 2 => stderr ide nazad Claude-u kao blokirajuci nalaz.
set -uo pipefail

# Brzi set: modTest -- 11 provera ponasanja (tri nad legacy formom, tri nad novim
# UI-jem, pet nad upisom zbirne i prijemnice). Sekunde uz Excel, ne minuti.
SUITE=RunAllTests

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$ROOT" || exit 0

# NE "command -v": na Windows-u PATH sadrzi Microsoft Store execution alias za
# python3 (~/AppData/Local/Microsoft/WindowsApps/python3). Fajl POSTOJI, pa ga
# command -v nadje, ali svaki poziv ispise "Python was not found" i izadje 49.
# $PY je tako bio interpreter koji ne radi: who_writes --check ispod je "padao"
# iz pogresnog razloga (lazno "WHO_WRITES.md je zastareo", exit 2 na svakom
# Stop-u), a do ziga i Excela se nikad nije stizalo. Proverava se IZVRSAVANJE.
PY=python3
"$PY" -c "" >/dev/null 2>&1 || PY=python
"$PY" -c "" >/dev/null 2>&1 || exit 0

# Sve dalje se pali samo kad je VBA izvor stvarno diran: ili u radnom stablu, ili
# u poslednjem commit-u (Claude cesto commit-uje pa tek onda stane -- tada radno
# stablo izgleda cisto).
changed=0
git diff --quiet HEAD -- src-vba/ 2>/dev/null || changed=1
if [ "$changed" -eq 0 ] && git rev-parse HEAD~1 >/dev/null 2>&1; then
    git diff --quiet HEAD~1 HEAD -- src-vba/ 2>/dev/null || changed=1
fi
[ "$changed" -eq 1 ] || exit 0

# Mapa vlasnistva se generise iz koda, pa istruli cim neko doda AddTableSnapshot
# ili AppendRow nad novom tabelom. Instant je i ne trazi Excel -- zato ide PRE
# i zig-a i win32com granice, da se proveri i u Linux sesiji i na svakom Stop-u.
if ! "$PY" tools/who_writes.py --check >/dev/null 2>&1; then
    echo "docs/DOMEN/WHO_WRITES.md je zastareo posle izmene u src-vba/." >&2
    echo "Regenerisi: python3 tools/who_writes.py --out" >&2
    exit 2
fi

# Isto stanje se ne proverava dvaput.
#
# Bez ovoga je grana "poslednji commit" iznad bila zamka: cim jednom commit-ujes
# izmenu u src-vba/, uslov ostaje tacan DO KRAJA SESIJE, pa se suite vrtela na
# svakom sledecem Stop-u -- i na turnovima gde nije dirnut nijedan fajl. Zig je
# HEAD + hash nekomitovanog diffa nad src-vba/, pa svaka nova izmena i svaki nov
# commit ponovo pale suite, a puko cavrljanje ne.
git_dir="$(git rev-parse --git-dir 2>/dev/null)"
stamp=""
[ -n "$git_dir" ] && stamp="$git_dir/vba-test-stamp"
token="$(git rev-parse HEAD 2>/dev/null)|$(git diff HEAD -- src-vba/ 2>/dev/null | git hash-object --stdin 2>/dev/null)"
if [ -n "$stamp" ] && [ -f "$stamp" ] && [ "$(cat "$stamp" 2>/dev/null)" = "$token" ]; then
    exit 0
fi

# Bez Excela nema sta da se vrti. Linux/macOS sesija (Claude Code na webu) prolazi
# TIHO -- tamo i dalje radi vba_check kroz PostToolUse i provera mape iznad. Hook
# koji bi tamo pao na svakom stop-u bio bi samo smetnja.
"$PY" -c "import win32com.client" >/dev/null 2>&1 || exit 0

# Fixture je lokalan artefakt (.gitignore) -- bez njega suite nema nad cim.
if [ ! -f tests/fixtures/otkup_test.xlsm ]; then
    echo "src-vba/ je menjan, ali tests/fixtures/otkup_test.xlsm ne postoji." >&2
    echo 'Napravi ga: python tools/make_fixture.py --donor "<put do .xlsm>"' >&2
    exit 2
fi

out="$("$PY" tools/run_vba.py --suite "$SUITE" 2>&1)"
rc=$?

if [ "$rc" -ne 0 ]; then
    echo "$out" >&2
    echo "" >&2
    echo "Brzi set ($SUITE) je pao posle izmene u src-vba/ -- vidi ime testa iznad." >&2
    echo "Detalji: .claude/rules/testovi.md" >&2
    exit 2
fi

# Zig se pise SAMO na zeleno: pao set mora da se ponovi na sledecem Stop-u.
[ -n "$stamp" ] && printf '%s' "$token" > "$stamp"

exit 0
