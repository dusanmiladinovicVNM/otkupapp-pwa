#!/usr/bin/env bash
# tools/release.sh — release rutina (git deo): pull -> bump APP_VERSION -> commit -> tag -> push -> stamp.
# Pokreni iz repo klona NA BUILD MASINI (gde Excel modVbaTools.FOLDER pokazuje na isti src-vba):
#   bash tools/release.sh 2.2.2
# Posle ovoga ostaju 3 Excel klika + restore — skripta ih ispiše. Vidi docs/RELEASE_PROCEDURE.md.
set -euo pipefail

ver="${1:-}"
if [ -z "$ver" ]; then
  echo "Upotreba: bash tools/release.sh <verzija>   (npr. 2.2.2)" >&2
  exit 1
fi
ver="${ver#v}"                         # dozvoli i "v2.2.2"
tag="vba-v${ver}"
root="$(git rev-parse --show-toplevel)"
cfg="$root/src-vba/modConfig.bas"
mbi="src-vba/modBuildInfo.bas"

echo "== 1/5 main + pull =="
git -C "$root" checkout main
git -C "$root" pull --ff-only origin main

# radni direktorijum mora biti cist (modBuildInfo.bas se ignorise — njega stamp prepisuje)
if [ -n "$(git -C "$root" status --porcelain -- . ":(exclude)$mbi")" ]; then
  echo "Radni direktorijum nije cist (osim $mbi). Commit/oclisti pa ponovi." >&2
  git -C "$root" status --short >&2
  exit 1
fi

echo "== 2/5 bump APP_VERSION -> $ver =="
if ! grep -q 'APP_VERSION As String' "$cfg"; then
  echo "Ne nalazim APP_VERSION u $cfg" >&2; exit 1
fi
tmp="$(mktemp)"
sed -E 's/APP_VERSION As String = "[^"]*"/APP_VERSION As String = "'"$ver"'"/' "$cfg" > "$tmp" && mv "$tmp" "$cfg"

if [ -n "$(git -C "$root" status --porcelain -- "$cfg")" ]; then
  git -C "$root" commit -m "release: vba v${ver}" -- "$cfg"
  git -C "$root" push origin main
else
  echo "   APP_VERSION je vec $ver — nema commita."
fi

echo "== 3/5 tag $tag =="
if git -C "$root" rev-parse -q --verify "refs/tags/$tag" >/dev/null; then
  echo "   Tag $tag vec postoji lokalno — preskačem kreiranje."
else
  git -C "$root" tag "$tag"
fi
git -C "$root" push origin "$tag"

echo "== 4/5 stamp build otisak =="
bash "$root/tools/stamp-build.sh"

echo "== 5/5 git deo gotov =="
echo
echo "------------------------------------------------------------"
echo "Sad u Excelu (master .xlsm) — RUCNO:"
echo "  1) Alt+F8 -> ImportAllVBA"
echo "  2) Debug -> Compile VBAProject   (mora bez greske)"
echo "  3) (opciono) VBE: Tools -> Digital Signature   (ako potpisuješ)"
echo "  4) Snimi (Ctrl+S)"
echo
echo "Zatim vrati placeholder i isporuci:"
echo "  git checkout -- $mbi"
echo "  -> pošalji .xlsm klijentima"
echo
echo "Provera: GAS rebuildMonitoringFleet() / Fleet tab (ili auto-trigger na sat)."
echo "------------------------------------------------------------"
