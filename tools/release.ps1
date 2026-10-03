# tools/release.ps1 — tanka ljuska nad tools/release.sh. JEDNA implementacija.
#   powershell -ExecutionPolicy Bypass -File tools/release.ps1 2.2.2
#   powershell -ExecutionPolicy Bypass -File tools/release.ps1 2.2.2 --waive ponasanje --reason "<zasto>"
#
# ZASTO JE OVO SAMO LJUSKA
# ------------------------
# Od 03.10.2026 `release.sh` nosi KAPIJU: tag ne moze da nastane dok marker ne
# dokaze da je bas taj src-vba prosao suite i rucni compile. Prepisana u
# PowerShell-u, ta logika bi bila DRUGA KOPIJA — pa bi se razisla, i to tiho.
# Tacno to se u ovom projektu vec desilo sa HARD/SOFT algoritmom
# (`modSelfUpdate` vs `modVbaTools`, PR #274), zbog cega sada postoji
# `tools/vba_parity_check.py`. Druga kopija release kapije bi bila gora: njena
# divergencija znaci tag bez dokaza, a nijedna kapija to ne bi videla.
#
# Skripta koja je zaobilazila kapiju je opasnija od skripte koje nema, pa ovde
# nema nijedne odluke — samo prosledjivanje.
$ErrorActionPreference = 'Stop'

$root = (git rev-parse --show-toplevel).Trim()
$sh = Join-Path $root 'tools/release.sh'

$bash = (Get-Command bash -ErrorAction SilentlyContinue)
if (-not $bash) {
  Write-Error @"
Nema 'bash' u PATH-u.

tools/release.ps1 je ljuska nad tools/release.sh (jedna implementacija kapije),
pa bez bash-a release ne moze da se pokrene. Git Bash dolazi sa Git for Windows:
otvori Git Bash i pokreni

    bash tools/release.sh $($args -join ' ')
"@
}

& $bash.Source $sh @args
exit $LASTEXITCODE
