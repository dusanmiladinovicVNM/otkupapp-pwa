# tools/install_nocni_dokaz.ps1
# Registruje `tools\dokaz_nocu.ps1` kao dnevni Scheduled Task.
#
#   powershell -ExecutionPolicy Bypass -File tools\install_nocni_dokaz.ps1
#   powershell -ExecutionPolicy Bypass -File tools\install_nocni_dokaz.ps1 -U 02:30
#   powershell -ExecutionPolicy Bypass -File tools\install_nocni_dokaz.ps1 -Ukloni
#
# Zadatak radi SAMO kad je korisnik prijavljen -- to je namerno, ne propust:
# `run_vba.py` trazi interaktivnu sesiju (compile probe gadja vidljiv i aktivan
# VBE prozor), pa "Run whether user is logged on or not" (Session 0) ne bi radio.
param(
    [string]$U = "03:00",
    [string]$Ime = "AgriX nocni recert",
    [switch]$Ukloni
)
$ErrorActionPreference = 'Stop'

if ($Ukloni) {
    Unregister-ScheduledTask -TaskName $Ime -Confirm:$false
    Write-Output "uklonjen zadatak: $Ime"
    exit 0
}

$root    = (git rev-parse --show-toplevel).Trim()
$skripta = Join-Path $root 'tools\dokaz_nocu.ps1'
if (-not (Test-Path $skripta)) { throw "nema $skripta" }

$akcija = New-ScheduledTaskAction -Execute 'powershell.exe' `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$skripta`"" `
    -WorkingDirectory $root

$okidac = New-ScheduledTaskTrigger -Daily -At $U

# StartWhenAvailable: ako je masina bila ugasena u 03:00, run krene cim moze.
# ExecutionTimeLimit 8h: pun katalog traje satima, podrazumevana 3h bi ga sekla.
$postavke = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Hours 8) `
    -MultipleInstances IgnoreNew

Register-ScheduledTask -TaskName $Ime -Action $akcija -Trigger $okidac `
    -Settings $postavke -Description "vba_check + FULL suite + pun dvosmerni dokaz; izvestaj u tests\dokaz_jutro.md" `
    -Force | Out-Null

Write-Output "registrovan: '$Ime', svakog dana u $U"
Write-Output "  ujutru:   tests\dokaz_jutro.md"
Write-Output "  provera:  Get-ScheduledTask -TaskName '$Ime'"
Write-Output "  probni:   Start-ScheduledTask -TaskName '$Ime'"
Write-Output ""
Write-Output "USLOV: masina prijavljena (Excel trazi interaktivnu sesiju)."
