#!/usr/bin/env bash
# tools/stamp-build.sh
# Upisuje trenutni git otisak u src-vba/modBuildInfo.bas PRE modVbaTools.ImportAllVBA.
# Pokreni iz bilo kog foldera repoa:  bash tools/stamp-build.sh
# Vidi docs/RELEASE_PROCEDURE.md.
set -euo pipefail

root="$(git rev-parse --show-toplevel)"
sha="$(git rev-parse --short HEAD)"
date="$(git show -s --format=%cI HEAD)"
# git describe = auto verzija iz gita: "vba-v2.2.1" na tagu, "vba-v2.2.1-3-gabc1234" +3 commita
ver="$(git describe --tags --always 2>/dev/null)"
[ -z "$ver" ] && ver="$sha"

# "dirty" ako ima necommit-ovanih izmena (osim samog modBuildInfo.bas koji upravo pišemo)
if [ -n "$(git status --porcelain -- . ':(exclude)src-vba/modBuildInfo.bas')" ]; then
  sha="${sha}+dirty"
  ver="${ver}+dirty"
fi

out="$root/src-vba/modBuildInfo.bas"
cat > "$out" <<EOF
Attribute VB_Name = "modBuildInfo"
Option Explicit

' AUTO-GENERISANO (tools/stamp-build). Ne edituj rucno; vidi docs/RELEASE_PROCEDURE.md.
Public Const BUILD_SHA As String = "$sha"
Public Const BUILD_VERSION As String = "$ver"
Public Const BUILD_DATE As String = "$date"
EOF

echo "Stamped $out -> $ver ($sha, $date)"
