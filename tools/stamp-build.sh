#!/usr/bin/env bash
# tools/stamp-build.sh
# Upisuje trenutni git otisak u src-vba/modBuildInfo.bas PRE modVbaTools.ImportAllVBA.
# Pokreni iz bilo kog foldera repoa:  bash tools/stamp-build.sh
# Vidi docs/RELEASE_PROCEDURE.md.
set -euo pipefail

root="$(git rev-parse --show-toplevel)"
sha="$(git rev-parse --short HEAD)"
date="$(git show -s --format=%cI HEAD)"

# "dirty" ako ima necommit-ovanih izmena (osim samog modBuildInfo.bas koji upravo pišemo)
if [ -n "$(git status --porcelain -- . ':(exclude)src-vba/modBuildInfo.bas')" ]; then
  sha="${sha}+dirty"
fi

out="$root/src-vba/modBuildInfo.bas"
cat > "$out" <<EOF
Attribute VB_Name = "modBuildInfo"
Option Explicit

' AUTO-GENERISANO (tools/stamp-build). Ne edituj rucno; vidi docs/RELEASE_PROCEDURE.md.
Public Const BUILD_SHA As String = "$sha"
Public Const BUILD_DATE As String = "$date"
EOF

echo "Stamped $out -> $sha ($date)"
