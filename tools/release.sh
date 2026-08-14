#!/usr/bin/env bash
# tools/release.sh -- release rutina sa OBAVEZNOM kapijom pre taga.
#
# Isti redosled kao tools/release.ps1 (koji je referenca na Windows build masini);
# obe skripte samo rade git plumbing oko tools/release_gate.py, pa logika kapije
# postoji na jednom mestu.
#
#   1  main + pull + cisto radno stablo
#   2  bump APP_VERSION U RADNO STABLO (bez commita)   <- gate vrti BAS to
#   3  release gate (static / fixture / behavior / green / compile / external)
#   4  commit + push                                    <- tek posle zelene kapije
#   5  anotiran tag sa verdiktom kapije + push
#   6  stamp build otisak
#   7  preostali Excel koraci
#
# Ako kapija padne, bump se vraca i radno stablo ostaje cisto.
#
# NAPOMENA: behavior i compile kapija traze Windows + Excel + pywin32. Iz Git
# Bash-a na build masini rade; iz Linux sesije obe padaju i release se ZAUSTAVLJA
# -- to je namerno, a ne kvar.
#
#   bash tools/release.sh 2.41.0
#   bash tools/release.sh 2.41.0 --dry-run
#   bash tools/release.sh 2.41.0 --waive external --reason "sandbox SEF nedostupan"
set -euo pipefail

ver=""
dry=""
waives=()
reason=""

while [ $# -gt 0 ]; do
  case "$1" in
    --dry-run) dry="--dry-run"; shift ;;
    --waive)   waives+=("--waive" "$2"); shift 2 ;;
    --reason)  reason="$2"; shift 2 ;;
    -*)        echo "Nepoznat argument: $1" >&2; exit 1 ;;
    *)         ver="$1"; shift ;;
  esac
done

if [ -z "$ver" ]; then
  echo "Upotreba: bash tools/release.sh <verzija> [--dry-run] [--waive <kapija> --reason \"...\"]" >&2
  exit 1
fi
ver="${ver#v}"                         # dozvoli i "v2.41.0"
tag="vba-v${ver}"
root="$(git rev-parse --show-toplevel)"
cfg="$root/src-vba/modConfig.bas"
mbi="src-vba/modBuildInfo.bas"

PY=python3
command -v "$PY" >/dev/null 2>&1 || PY=python

restore_bump() {
  git -C "$root" checkout -- 'src-vba/modConfig.bas'
  echo "   Bump vracen -- radno stablo je cisto."
}

echo "== 1/7 main + pull =="
git -C "$root" checkout main
git -C "$root" pull --ff-only origin main

# radni direktorijum mora biti cist (modBuildInfo.bas se ignorise -- njega stamp prepisuje)
if [ -n "$(git -C "$root" status --porcelain -- . ":(exclude)$mbi")" ]; then
  echo "Radni direktorijum nije cist (osim $mbi). Commit/oclisti pa ponovi." >&2
  git -C "$root" status --short >&2
  exit 1
fi

echo "== 2/7 bump APP_VERSION -> $ver (samo u radno stablo) =="
if ! grep -q 'APP_VERSION As String' "$cfg"; then
  echo "Ne nalazim APP_VERSION u $cfg" >&2; exit 1
fi
tmp="$(mktemp)"
sed -E 's/APP_VERSION As String = "[^"]*"/APP_VERSION As String = "'"$ver"'"/' "$cfg" > "$tmp" && mv "$tmp" "$cfg"

echo "== 3/7 RELEASE GATE =="
gate_args=("$root/tools/release_gate.py" --version "$ver")
[ ${#waives[@]} -gt 0 ] && gate_args+=("${waives[@]}")
[ -n "$reason" ] && gate_args+=(--reason "$reason")
[ -n "$dry" ] && gate_args+=("$dry")

if ! "$PY" "${gate_args[@]}"; then
  restore_bump
  echo "" >&2
  echo "Release kapija nije prosla -- nema commita, nema taga, nema push-a." >&2
  echo "Vidi ispis iznad i tests/last_release.json." >&2
  exit 2
fi

if [ -n "$dry" ]; then
  restore_bump
  echo
  echo "DRY RUN: kapije su samo izlistane, nista nije commit-ovano ni tag-ovano."
  exit 0
fi

echo "== 4/7 commit + push =="
if [ -n "$(git -C "$root" status --porcelain -- "$cfg")" ]; then
  git -C "$root" commit -m "release: vba v${ver}" -- "$cfg"
  git -C "$root" push origin main
else
  echo "   APP_VERSION je vec $ver -- nema commita."
fi

echo "== 5/7 anotiran tag $tag =="
# Tag nosi verdikt kapije u poruci: `git show vba-vX.Y.Z` kasnije tacno kaze sta
# je bilo zeleno, sta izuzeto i nad kojim hashom src-vba.
msgfile="$(mktemp)"
"$PY" "$root/tools/release_gate.py" --version "$ver" --tag-message > "$msgfile"

if git -C "$root" rev-parse -q --verify "refs/tags/$tag" >/dev/null; then
  echo "   Tag $tag vec postoji lokalno -- preskacem kreiranje."
else
  git -C "$root" tag -a "$tag" -F "$msgfile"
fi
git -C "$root" push origin "$tag"
rm -f "$msgfile"

echo "== 6/7 stamp build otisak =="
bash "$root/tools/stamp-build.sh"

echo "== 7/7 git deo gotov =="
echo
echo "------------------------------------------------------------"
echo "Kapija je prosla PRE taga -- Excel koraci ispod su isporuka, ne provera."
echo
echo "Sad u Excelu (prazan build-master .xlsm):"
echo "  1) Alt+F8 -> ImportAllVBA"
echo "  2) Debug -> Compile VBAProject   (mora bez greske)"
echo "  3) Alt+F8 -> AssertBlankBuild    (mora 'BLANKO OK')"
echo "  3b) Alt+F8 -> PublishReleaseToDrive  (kod + version.json u AgriX_Release)"
echo "  4) (opciono) VBE: Tools -> Digital Signature   (ako potpisujes)"
echo "  5) Snimi / Save As builds/AgriX_${ver}.xlsm"
echo
echo "Zatim vrati placeholder i isporuci:"
echo "  git checkout -- $mbi"
echo "  -> posalji .xlsm klijentima"
echo
echo "Verdikt kapije je u poruci taga:  git show $tag"
echo "Ne zaboravi: dopuni docs/RELEASE_NOTES.md (par recenica o ovom izdanju)."
echo
echo "Provera: GAS rebuildMonitoringFleet() / Fleet tab (ili auto-trigger na sat)."
echo "------------------------------------------------------------"
