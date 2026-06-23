# tools/release.ps1 — release rutina (git deo): pull -> bump APP_VERSION -> commit -> tag -> push -> stamp.
# Pokreni iz repo klona NA BUILD MASINI (gde Excel modVbaTools.FOLDER pokazuje na isti src-vba):
#   powershell -ExecutionPolicy Bypass -File tools\release.ps1 2.2.2
# Posle ovoga ostaju 3 Excel klika + restore — skripta ih ispiše. Vidi docs/RELEASE_PROCEDURE.md.
param([Parameter(Mandatory = $true)][string]$Version)
$ErrorActionPreference = 'Stop'

$ver = $Version.TrimStart('v')
$tag = "vba-v$ver"
$root = (git rev-parse --show-toplevel).Trim()
$cfg = Join-Path $root 'src-vba\modConfig.bas'
$mbi = 'src-vba/modBuildInfo.bas'

Write-Host "== 1/5 main + pull =="
git -C $root checkout main
git -C $root pull --ff-only origin main

# radni direktorijum mora biti cist (modBuildInfo.bas se ignorise — njega stamp prepisuje)
$dirty = git -C $root status --porcelain -- . ":(exclude)$mbi"
if (-not [string]::IsNullOrWhiteSpace(($dirty -join ''))) {
  Write-Error "Radni direktorijum nije cist (osim $mbi). Commit/oclisti pa ponovi."
}

Write-Host "== 2/5 bump APP_VERSION -> $ver =="
$content = Get-Content -Raw $cfg
if ($content -notmatch 'APP_VERSION As String') { Write-Error "Ne nalazim APP_VERSION u $cfg" }
$content = [regex]::Replace($content, 'APP_VERSION As String = "[^"]*"', 'APP_VERSION As String = "' + $ver + '"')
Set-Content -Path $cfg -Value $content -NoNewline -Encoding ASCII

$cfgDirty = git -C $root status --porcelain -- $cfg
if (-not [string]::IsNullOrWhiteSpace(($cfgDirty -join ''))) {
  git -C $root commit -m "release: vba v$ver" -- $cfg
  git -C $root push origin main
} else {
  Write-Host "   APP_VERSION je vec $ver — nema commita."
}

Write-Host "== 3/5 tag $tag =="
git -C $root rev-parse -q --verify "refs/tags/$tag" *> $null
if ($LASTEXITCODE -eq 0) { Write-Host "   Tag $tag vec postoji lokalno — preskačem." }
else { git -C $root tag $tag }
git -C $root push origin $tag

Write-Host "== 4/5 stamp build otisak =="
powershell -ExecutionPolicy Bypass -File (Join-Path $root 'tools\stamp-build.ps1')

Write-Host "== 5/5 git deo gotov =="
Write-Host ""
Write-Host "------------------------------------------------------------"
Write-Host "Sad u Excelu (master .xlsm) — RUCNO:"
Write-Host "  1) Alt+F8 -> ImportAllVBA"
Write-Host "  2) Debug -> Compile VBAProject   (mora bez greske)"
Write-Host "  3) (opciono) VBE: Tools -> Digital Signature   (ako potpisuješ)"
Write-Host "  4) Snimi (Ctrl+S)"
Write-Host ""
Write-Host "Zatim vrati placeholder i isporuci:"
Write-Host "  git checkout -- $mbi"
Write-Host "  -> pošalji .xlsm klijentima"
Write-Host ""
Write-Host "Ne zaboravi: dopuni docs/RELEASE_NOTES.md (par rečenica o ovom izdanju)."
Write-Host ""
Write-Host "Provera: GAS rebuildMonitoringFleet() / Fleet tab (ili auto-trigger na sat)."
Write-Host "------------------------------------------------------------"
