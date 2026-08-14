#!/usr/bin/env bash
# PostToolUse hook: posle Edit/Write nad VBA izvorom pusti tools/vba_check.py
# nad tim fajlom. Exit 2 => stderr ide nazad Claude-u kao blokirajuci nalaz.
#
# Namena: greske iz CLAUDE.md sec.4 (ne-ASCII bajt, Const posle prve procedure,
# rezervisana rec, orphan Poruka kljuc) do sada su se videle tek kad operater
# uradi ImportAllVBA + Compile. Sada se vide u istoj sekundi.
set -uo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"

# NE hardkodiran python3: na Windows-u PATH sadrzi Microsoft Store execution
# alias za python3 koji POSTOJI kao fajl, ali svaki poziv izadje sa 49. Parsiranje
# payload-a je zato tiho puklo, file_path je ostajao prazan i hook je izlazio 0 --
# provera se NIKAD nije izvrsila, bez ijedne poruke. Proverava se IZVRSAVANJE.
PY=python3
"$PY" -c "" >/dev/null 2>&1 || PY=python
"$PY" -c "" >/dev/null 2>&1 || exit 0

payload="$(cat)"

file_path="$(printf '%s' "$payload" | "$PY" -c '
import json, sys
try:
    d = json.load(sys.stdin)
except Exception:
    sys.exit(0)
print(d.get("tool_input", {}).get("file_path", "") or "")
')"

[ -n "$file_path" ] || exit 0

# .claude/settings*.json: fajl koji ne prodje JSON validaciju Claude Code odbacuje
# U CELINI -- ne to jedno pravilo, nego SVA permission pravila i OBA hook-a
# odjednom (2da7ee8e: jedan izostavljen zarez ih je drzao van snage pola dana).
# Odgovor mora da stigne u istom turnu, jer sledeci Edit vec radi bez ijednog
# pravila i bez ovog hook-a.
case "$file_path" in
  *.claude/settings.json|*.claude/settings.local.json)
    [ -f "$file_path" ] || exit 0
    err="$("$PY" -c 'import json,sys; json.load(open(sys.argv[1], encoding="utf-8"))' "$file_path" 2>&1)"
    rc=$?
    if [ "$rc" -ne 0 ]; then
      echo "$file_path nije validan JSON:" >&2
      printf '%s\n' "$err" | tail -n 1 >&2
      echo "Claude Code takav fajl odbacuje U CELINI -- ne vazi nijedno permission" >&2
      echo "pravilo ni jedan hook. Popravi pre nego sto nastavis." >&2
      exit 2
    fi
    exit 0
    ;;
esac

case "$file_path" in
  *.bas|*.cls|*.frm|*.doccls) ;;
  *) exit 0 ;;
esac

[ -f "$file_path" ] || exit 0

if ! "$PY" "$ROOT/tools/vba_check.py" --hook "$file_path"; then
  echo "" >&2
  echo "vba_check je pao nad $file_path -- popravi pre nego sto nastavis (.claude/rules/vba-izvor.md)." >&2
  exit 2
fi

exit 0
