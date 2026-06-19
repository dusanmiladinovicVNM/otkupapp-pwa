<#
.SYNOPSIS
    Konvertuje POSTOJECI, popunjeni klijentski .xlsm u cist .xlsx
    (zadrzava SVE podatke/config/formatiranje, izbacuje samo VBA kod).

.DESCRIPTION
    Za klijente koji su VEC poceli da pune bazu. Otvara .xlsm sa onemogucenim
    makroima (da se aplikacija ne pokrene), pravi backup originala i snima kao
    .xlsx (FileFormat 51) - Excel pri tome automatski odbacuje VBA projekat.
    Original .xlsm ostaje netaknut na disku (rollback).

    Posle konverzije, kod dolazi iz instaliranog OtkupApp.xlam, koji ovaj .xlsx
    prepoznaje po marker tabeli 'tblConfig' i pokrece aplikaciju pri otvaranju.

.PREREQUISITES
    - Windows + Microsoft Excel.
    - Klijent mora da ZATVORI aplikaciju pre konverzije (da fajl bude otkljucan
      i da poslednji unos bude na disku).

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File build\Convert-XlsmToXlsx.ps1 -Source "C:\OtkupApp\OtkupApp.xlsm"
    powershell -ExecutionPolicy Bypass -File build\Convert-XlsmToXlsx.ps1 -Source "C:\OtkupApp\OtkupApp.xlsm" -Output "C:\OtkupApp\Klijent_C001.xlsx"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Source,
    [string]$Output,
    [string]$BackupDir,
    [switch]$NoBackup
)

$ErrorActionPreference = "Stop"
$xlOpenXMLWorkbook = 51    # .xlsx (bez makroa)
$msoForceDisable = 3       # msoAutomationSecurityForceDisable

function Resolve-FullPath([string]$p) {
    return [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $p))
}

$Source = Resolve-FullPath $Source
if (-not (Test-Path $Source)) { throw "Izvor ne postoji: $Source" }
if ([System.IO.Path]::GetExtension($Source).ToLower() -ne ".xlsm") {
    Write-Warning "Izvor nije .xlsm: $Source (nastavljam svejedno)."
}

$srcDir = Split-Path $Source -Parent
$base = [System.IO.Path]::GetFileNameWithoutExtension($Source)
if (-not $Output) { $Output = Join-Path $srcDir ($base + ".xlsx") }
$Output = Resolve-FullPath $Output
if (($Source.ToLower()) -eq ($Output.ToLower())) { throw "Output ne sme biti isti kao Source." }

# --- 1) Backup originalnog .xlsm ---
if (-not $NoBackup) {
    if (-not $BackupDir) { $BackupDir = Join-Path $srcDir "_pre_split_backup" }
    if (-not (Test-Path $BackupDir)) { New-Item -ItemType Directory -Path $BackupDir -Force | Out-Null }
    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $bak = Join-Path $BackupDir ("$base`_$stamp.xlsm")
    Copy-Item -LiteralPath $Source -Destination $bak -Force
    Write-Host "Backup: $bak" -ForegroundColor Cyan
}

# --- 2) Konverzija (makroi onemoguceni - app se NE pokrece) ---
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
try { $excel.AutomationSecurity = $msoForceDisable } catch {}

$wb = $null
try {
    $wb = $excel.Workbooks.Open($Source)

    # marker provera (da .xlam kasnije prepozna fajl)
    $hasConfig = $false
    foreach ($ws in $wb.Worksheets) {
        foreach ($lo in $ws.ListObjects) {
            if ([string]$lo.Name -ieq "tblConfig") { $hasConfig = $true }
        }
    }
    if (-not $hasConfig) {
        Write-Warning "U fajlu NEMA tabele 'tblConfig' - OtkupApp.xlam ga mozda nece prepoznati automatski."
    }

    if (Test-Path $Output) { Remove-Item -LiteralPath $Output -Force }
    # SaveAs .xlsx pravi NOV fajl; original .xlsm na disku ostaje nepromenjen.
    $wb.SaveAs($Output, $xlOpenXMLWorkbook)
    Write-Host "Napravljen: $Output" -ForegroundColor Green
}
finally {
    if ($null -ne $wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}

# --- 3) Verifikacija: ponovo otvori .xlsx i prebroj tabele/redove ---
$excel2 = New-Object -ComObject Excel.Application
$excel2.Visible = $false
$excel2.DisplayAlerts = $false
$excel2.EnableEvents = $false
try { $excel2.AutomationSecurity = $msoForceDisable } catch {}
$v = $null
try {
    $v = $excel2.Workbooks.Open($Output)
    $tblCount = 0; $cfgRows = -1
    foreach ($ws in $v.Worksheets) {
        foreach ($lo in $ws.ListObjects) {
            $tblCount++
            if ([string]$lo.Name -ieq "tblConfig" -and $null -ne $lo.DataBodyRange) {
                $cfgRows = $lo.DataBodyRange.Rows.Count
            }
        }
    }
    Write-Host "Verifikacija: tabela=$tblCount, tblConfig redova=$cfgRows" -ForegroundColor Green
    Write-Host "OK - podaci zadrzani, VBA izbacen. Original .xlsm netaknut (rollback)." -ForegroundColor Green
}
finally {
    if ($null -ne $v) { $v.Close($false) }
    $excel2.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel2) | Out-Null
}
