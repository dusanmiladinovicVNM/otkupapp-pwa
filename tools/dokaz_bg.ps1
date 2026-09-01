# tools/dokaz_bg.ps1
# Pusti dvosmerni dokaz ODVOJENO i vrati se ODMAH. Sesija ga NE ceka.
#
# Zasto postoji: dok je jedini izlaz dokaza bio stdout, neko je morao da sedi
# nad njim -- sesija je blokirala u desetominutnim blokovima, trosila kontekst i
# kvotu na cekanje, a verdikt je posle svega postojao samo u scrollback-u.
#
#   powershell -File tools\dokaz_bg.ps1                        # ceo katalog
#   powershell -File tools\dokaz_bg.ps1 modOtkupUI.bas         # samo taj fajl
#   powershell -File tools\dokaz_bg.ps1 -Knjiga modScrIzvestaji.bas
#
# Napredak i verdikt:  python tools\dokaz.py --status
param(
    [switch]$Knjiga,
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Filter
)
$ErrorActionPreference = 'Stop'

$root  = (git rev-parse --show-toplevel).Trim()
$tests = Join-Path $root 'tests'
New-Item -ItemType Directory -Force -Path $tests | Out-Null

$log  = Join-Path $tests 'dokaz_last.log'
$err  = Join-Path $tests 'dokaz_err.log'
$json = Join-Path $tests 'dokaz_last.json'

# Stari verdikt se brise PRE pokretanja. Bez ovoga bi `--status` u prvim
# sekundama pokazao juceranje ZELENO kao da je od ovog run-a -- tacno ona vrsta
# tihe laznosti zbog koje verdikt uopste ide u fajl.
Remove-Item -Force -ErrorAction SilentlyContinue $json, $log, $err

$argumenti = @('tools\dokaz.py', '--json', $json)
if ($Knjiga) { $argumenti += '--knjiga' }
if ($Filter) { $argumenti += $Filter }

$p = Start-Process -FilePath 'python' -ArgumentList $argumenti `
        -WorkingDirectory $root -WindowStyle Hidden -PassThru `
        -RedirectStandardOutput $log -RedirectStandardError $err

Write-Output "dokaz pusten odvojeno (PID $($p.Id)). Sesija ga NE ceka."
Write-Output "  napredak/verdikt:  python tools\dokaz.py --status"
Write-Output "  pun ispis:         tests\dokaz_last.log"
Write-Output "  prekid:            Stop-Process -Id $($p.Id)   (sabotaza.py --vrati posle!)"
