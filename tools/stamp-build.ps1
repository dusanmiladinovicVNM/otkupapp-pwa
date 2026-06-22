# tools/stamp-build.ps1
# Upisuje trenutni git otisak u src-vba\modBuildInfo.bas PRE modVbaTools.ImportAllVBA.
# Pokreni iz root-a repoa:
#   powershell -ExecutionPolicy Bypass -File tools\stamp-build.ps1
# Vidi docs/RELEASE_PROCEDURE.md.
$ErrorActionPreference = 'Stop'

$root = (git rev-parse --show-toplevel).Trim()
$sha  = (git rev-parse --short HEAD).Trim()
$date = (git show -s --format=%cI HEAD).Trim()

# "dirty" ako ima necommit-ovanih izmena (osim samog modBuildInfo.bas koji upravo pišemo)
$dirty = git status --porcelain -- . ':(exclude)src-vba/modBuildInfo.bas'
if (-not [string]::IsNullOrWhiteSpace(($dirty -join ''))) { $sha = "$sha+dirty" }

$out = Join-Path $root 'src-vba\modBuildInfo.bas'
$content = @"
Attribute VB_Name = "modBuildInfo"
Option Explicit

' AUTO-GENERISANO (tools/stamp-build). Ne edituj rucno; vidi docs/RELEASE_PROCEDURE.md.
Public Const BUILD_SHA As String = "$sha"
Public Const BUILD_DATE As String = "$date"
"@

Set-Content -Path $out -Value $content -Encoding ASCII
Write-Host "Stamped $out -> $sha ($date)"
