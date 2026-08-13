#!/usr/bin/env bash
# Stop hook: kad je u sesiji diran src-vba/, pusti behavior suite pre kraja.
#
# vba_check (PostToolUse) hvata sintaksu i compile-hard greske. Ovo hvata drugu
# klasu: izmenu koja se uredno kompajlira, a menja PONASANJE -- npr. brisanje
# datuma ili zbirne u ClearOtkupFields. Vidi .claude/rules/testovi.md.
#
# Exit 2 => stderr ide nazad Claude-u kao blokirajuci nalaz.
set -uo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$ROOT" || exit 0

PY=python3
command -v "$PY" >/dev/null 2>&1 || PY=python
command -v "$PY" >/dev/null 2>&1 || exit 0

# Bez Excela nema sta da se vrti. Linux/macOS sesija (Claude Code na webu) prolazi
# TIHO -- tamo i dalje radi vba_check kroz PostToolUse. Hook koji bi tamo pao na
# svakom stop-u bio bi samo smetnja.
"$PY" -c "import win32com.client" >/dev/null 2>&1 || exit 0

# Suite kosta ~15s uz podizanje Excela, pa se pali samo kad je VBA izvor stvarno
# diran: ili u radnom stablu, ili u poslednjem commit-u (Claude cesto commit-uje
# pa tek onda stane -- tada radno stablo izgleda cisto).
changed=0
git diff --quiet HEAD -- src-vba/ 2>/dev/null || changed=1
if [ "$changed" -eq 0 ] && git rev-parse HEAD~1 >/dev/null 2>&1; then
    git diff --quiet HEAD~1 HEAD -- src-vba/ 2>/dev/null || changed=1
fi
[ "$changed" -eq 1 ] || exit 0

# Fixture je lokalan artefakt (.gitignore) -- bez njega suite nema nad cim.
if [ ! -f tests/fixtures/otkup_test.xlsm ]; then
    echo "src-vba/ je menjan, ali tests/fixtures/otkup_test.xlsm ne postoji." >&2
    echo 'Napravi ga: python tools/make_fixture.py --donor "<put do .xlsm>"' >&2
    exit 2
fi

out="$("$PY" tools/run_vba.py --suite RunAllTests 2>&1)"
rc=$?

if [ "$rc" -ne 0 ]; then
    echo "$out" >&2
    echo "" >&2
    echo "Behavior suite je pala posle izmene u src-vba/ -- vidi ime testa iznad." >&2
    echo "Detalji: .claude/rules/testovi.md" >&2
    exit 2
fi

exit 0
