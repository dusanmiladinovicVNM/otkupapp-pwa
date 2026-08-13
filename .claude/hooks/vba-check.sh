#!/usr/bin/env bash
# PostToolUse hook: posle Edit/Write nad VBA izvorom pusti tools/vba_check.py
# nad tim fajlom. Exit 2 => stderr ide nazad Claude-u kao blokirajuci nalaz.
#
# Namena: greske iz CLAUDE.md sec.4 (ne-ASCII bajt, Const posle prve procedure,
# rezervisana rec, orphan Poruka kljuc) do sada su se videle tek kad operater
# uradi ImportAllVBA + Compile. Sada se vide u istoj sekundi.
set -uo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"

payload="$(cat)"

file_path="$(printf '%s' "$payload" | python3 -c '
import json, sys
try:
    d = json.load(sys.stdin)
except Exception:
    sys.exit(0)
print(d.get("tool_input", {}).get("file_path", "") or "")
')"

[ -n "$file_path" ] || exit 0

case "$file_path" in
  *.bas|*.cls|*.frm|*.doccls) ;;
  *) exit 0 ;;
esac

[ -f "$file_path" ] || exit 0

if ! python3 "$ROOT/tools/vba_check.py" --hook "$file_path"; then
  echo "" >&2
  echo "vba_check je pao nad $file_path -- popravi pre nego sto nastavis (.claude/rules/vba-izvor.md)." >&2
  exit 2
fi

exit 0
