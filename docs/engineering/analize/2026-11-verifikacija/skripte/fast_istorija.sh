#!/usr/bin/env bash
# Meri FAST kapije nad istorijskim snimcima (git archive u scratch; nista se ne menja u repo stablima).
set -u
REPO=/home/user/otkupapp-pwa
W=/tmp/claude-0/-home-user-otkupapp-pwa/d6cefc99-c516-5b26-8cb4-97e3f89ec739/scratchpad/istorija-metrike/snimci
mkdir -p "$W"
t() { # ispisi min od 2 merenja (s) i rc
  local best=999999 rc=0
  for i in 1 2; do
    local a=$(date +%s.%N); "$@" >/dev/null 2>&1; rc=$?; local b=$(date +%s.%N)
    local d=$(python3 -c "print(round($b-$a,2))")
    best=$(python3 -c "print(min($best,$d))")
  done
  echo "$best(rc=$rc)"
}
echo "dan commit sabotaza vba_fajlova vba_check_full hook_modConfig proveri_sidra"
for d in "$@"; do
  c=$(git -C $REPO log --first-parent --until="$d 23:59:59" -1 --format=%H origin/main)
  dir="$W/$d"; rm -rf "$dir"; mkdir -p "$dir"
  git -C $REPO archive "$c" | tar -x -C "$dir"
  nf=$(ls "$dir/src-vba" 2>/dev/null | grep -cE '\.(bas|cls|frm|doccls)$')
  ns=$(awk -F, -v d="$d" '$1<=d{v=$3} END{print v}' /tmp/claude-0/-home-user-otkupapp-pwa/d6cefc99-c516-5b26-8cb4-97e3f89ec739/scratchpad/istorija-metrike/rast.csv)
  full=$(cd "$dir" && t python3 tools/vba_check.py)
  hook=$(cd "$dir" && t python3 tools/vba_check.py --hook src-vba/modConfig.bas)
  if [ -f "$dir/tools/sabotaza.py" ]; then sid=$(cd "$dir" && t python3 tools/sabotaza.py --proveri-sidra); else sid="-"; fi
  echo "$d ${c:0:8} $ns $nf $full $hook $sid"
done
