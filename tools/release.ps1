# tools/release.ps1 -- release rutina sa OBAVEZNOM kapijom pre taga.
#
# Redosled je bitan i menjan je namerno. Ranije je skripta radila bump ->
# commit -> push -> tag -> push, pa TEK ONDA rekla operateru da uradi Import i
# Compile u Excelu. Behavior gate nije bio deo skripte uopste. Tako je nastao
# vba-v2.40.0: tag-ovan i objavljen uz posteno zapisanu napomenu da testovi nisu
# ni pokrenuti. Postena napomena posle taga ne pomaze -- tag vec postoji.
#
# Sada je redosled:
#   1  main + pull + cisto radno stablo
#   2  bump APP_VERSION U RADNO STABLO (bez commita)   <- gate mora da vrti BAS ovo
#   3  release gate (static / fixture / behavior / green / compile / external)
#   4  commit + push                                    <- tek posle zelene kapije
#   5  anotiran tag sa verdiktom kapije + push
#   6  stamp build otisak
#   7  preostali Excel koraci
#
# Ako kapija padne, bump se vraca i radno stablo ostaje cisto -- kao da se nista
# nije desilo.
#
# Pokreni iz repo klona NA BUILD MASINI (Windows + Excel + pywin32):
#   powershell -ExecutionPolicy Bypass -File tools\release.ps1 2.41.0
#   powershell -ExecutionPolicy Bypass -File tools\release.ps1 2.41.0 -DryRun
#   powershell -ExecutionPolicy Bypass -File tools\release.ps1 2.41.0 -Waive external -WaiveReason "sandbox SEF nedostupan 14.08."
#
# Vidi docs/RELEASE_PROCEDURE.md i .claude/rules/testovi.md.
param(
  [Parameter(Mandatory = $true)][string]$Version,
  [string[]]$Waive = @(),
  [string]$WaiveReason = "",
  [switch]$DryRun
)
$ErrorActionPreference = 'Stop'

$ver = $Version.TrimStart('v')
$tag = "vba-v$ver"
$root = (git rev-parse --show-toplevel).Trim()
$cfg = Join-Path $root 'src-vba\modConfig.bas'
$mbi = 'src-vba/modBuildInfo.bas'

function Restore-Bump {
  git -C $root checkout -- 'src-vba/modConfig.bas'
  Write-Host "   Bump vracen -- radno stablo je cisto."
}

Write-Host "== 1/7 main + pull =="
git -C $root checkout main
git -C $root pull --ff-only origin main

# radni direktorijum mora biti cist (modBuildInfo.bas se ignorise -- njega stamp prepisuje)
$dirty = git -C $root status --porcelain -- . ":(exclude)$mbi"
if (-not [string]::IsNullOrWhiteSpace(($dirty -join ''))) {
  Write-Error "Radni direktorijum nije cist (osim $mbi). Commit/oclisti pa ponovi."
}

Write-Host "== 2/7 bump APP_VERSION -> $ver (samo u radno stablo) =="
$content = Get-Content -Raw $cfg
if ($content -notmatch 'APP_VERSION As String') { Write-Error "Ne nalazim APP_VERSION u $cfg" }
$content = [regex]::Replace($content, 'APP_VERSION As String = "[^"]*"', 'APP_VERSION As String = "' + $ver + '"')
Set-Content -Path $cfg -Value $content -NoNewline -Encoding ASCII

Write-Host "== 3/7 RELEASE GATE =="
$gateArgs = @('tools\release_gate.py', '--version', $ver)
foreach ($w in $Waive) { $gateArgs += @('--waive', $w) }
if ($WaiveReason) { $gateArgs += @('--reason', $WaiveReason) }
if ($DryRun) { $gateArgs += '--dry-run' }

python @gateArgs
$gateRc = $LASTEXITCODE

if ($gateRc -ne 0) {
  Restore-Bump
  Write-Error "Release kapija nije prosla -- nema commita, nema taga, nema push-a. Vidi ispis iznad i tests\last_release.json."
}

if ($DryRun) {
  Restore-Bump
  Write-Host ""
  Write-Host "DRY RUN: kapije su samo izlistane, nista nije commit-ovano ni tag-ovano."
  exit 0
}

Write-Host "== 4/7 commit + push =="
$cfgDirty = git -C $root status --porcelain -- $cfg
if (-not [string]::IsNullOrWhiteSpace(($cfgDirty -join ''))) {
  git -C $root commit -m "release: vba v$ver" -- $cfg
  git -C $root push origin main
} else {
  Write-Host "   APP_VERSION je vec $ver -- nema commita."
}

Write-Host "== 5/7 anotiran tag $tag =="
# Tag nosi verdikt kapije u svojoj poruci. To je jedini zapis koji putuje sa
# kodom: `git show vba-vX.Y.Z` kasnije tacno kaze sta je bilo zeleno, sta
# izuzeto i nad kojim hashom src-vba.
$msgFile = Join-Path $env:TEMP "agrix-tag-$ver.txt"
python tools\release_gate.py --version $ver --tag-message | Out-File -FilePath $msgFile -Encoding utf8

git -C $root rev-parse -q --verify "refs/tags/$tag" *> $null
if ($LASTEXITCODE -eq 0) {
  Write-Host "   Tag $tag vec postoji lokalno -- preskacem kreiranje."
} else {
  git -C $root tag -a $tag -F $msgFile
}
git -C $root push origin $tag
Remove-Item $msgFile -ErrorAction SilentlyContinue

Write-Host "== 6/7 stamp build otisak =="
powershell -ExecutionPolicy Bypass -File (Join-Path $root 'tools\stamp-build.ps1')

Write-Host "== 7/7 git deo gotov =="
Write-Host ""
Write-Host "------------------------------------------------------------"
Write-Host "Kapija je prosla PRE taga -- Excel koraci ispod su isporuka, ne provera."
Write-Host ""
Write-Host "Sad u Excelu (prazan build-master .xlsm):"
Write-Host "  1) Alt+F8 -> ImportAllVBA"
Write-Host "  2) Debug -> Compile VBAProject   (mora bez greske)"
Write-Host "  3) Alt+F8 -> AssertBlankBuild    (mora 'BLANKO OK')"
Write-Host "  3b) Alt+F8 -> PublishReleaseToDrive  (kod + version.json u AgriX_Release)"
Write-Host "  4) (opciono) VBE: Tools -> Digital Signature   (ako potpisujes)"
Write-Host "  5) Snimi / Save As builds\AgriX_$ver.xlsm"
Write-Host ""
Write-Host "Zatim vrati placeholder i isporuci:"
Write-Host "  git checkout -- $mbi"
Write-Host "  -> posalji .xlsm klijentima"
Write-Host ""
Write-Host "Verdikt kapije je u poruci taga:  git show $tag"
Write-Host "Ne zaboravi: dopuni docs/RELEASE_NOTES.md (par recenica o ovom izdanju)."
Write-Host ""
Write-Host "Provera: GAS rebuildMonitoringFleet() / Fleet tab (ili auto-trigger na sat)."
Write-Host "------------------------------------------------------------"
