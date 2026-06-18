<#
.SYNOPSIS
    Pravi novi, CIST klijentski .xlsx fajl (podaci + config, BEZ ijedne linije VBA).

.DESCRIPTION
    Polazi od postojeceg izvora (trenutni OtkupApp.xlsm ili pripremljeni template),
    isprazni sve poslovne tabele (tbl*) osim config/catalog tabela, opciono obrise
    vrednosti u config tabelama, i snimi kao .xlsx. Posto je izlaz .xlsx, Excel
    automatski odbacuje VBA - klijentski fajl ostaje cист data fajl.

    Aplikacija (OtkupApp.xlam) sama prepoznaje ovaj fajl po marker tabeli 'tblConfig'
    i pokrece se kada se fajl otvori (vidi clsAppEvents / modBoot).

.PREREQUISITES
    - Windows + Microsoft Excel.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File build\New-ClientFile.ps1 `
        -Source "C:\OtkupApp\OtkupApp.xlsm" -Output "C:\OtkupApp\Klijent_C00X.xlsx" -ClearConfigValues
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Source,
    [Parameter(Mandatory = $true)][string]$Output,

    # Tabele cije vrednosti se zadrzavaju (config + katalozi). Sve ostale tbl* se prazne.
    [string[]]$KeepTables = @(
        "tblConfig", "tblSEFConfig", "tblLocalConfig",
        "tblKulture", "tblTipPalete", "tblTipAmbalaze", "tblKvalitet"
    ),

    # Ako je zadato, isprazni i VREDNOSTI u config tabelama (kljucevi ostaju).
    [switch]$ClearConfigValues
)

$ErrorActionPreference = "Stop"
$xlOpenXMLWorkbook = 51   # FileFormat za .xlsx (bez makroa)
$msoForceDisable = 3

function Resolve-FullPath([string]$p) {
    return [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $p))
}
$Source = Resolve-FullPath $Source
$Output = Resolve-FullPath $Output
if (-not (Test-Path $Source)) { throw "Izvor ne postoji: $Source" }
$OutDir = Split-Path $Output -Parent
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }

$configTables = @("tblConfig", "tblSEFConfig", "tblLocalConfig")

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try { $excel.AutomationSecurity = $msoForceDisable } catch {}

$wb = $null
try {
    $wb = $excel.Workbooks.Open($Source)

    $cleared = 0; $kept = 0
    foreach ($ws in $wb.Worksheets) {
        foreach ($lo in $ws.ListObjects) {
            $name = [string]$lo.Name
            if ($name -notlike "tbl*") { continue }

            if ($KeepTables -contains $name) {
                $kept++
                if ($ClearConfigValues -and ($configTables -contains $name)) {
                    # Isprazni vrednosnu kolonu, zadrzi kljuceve.
                    foreach ($valCol in @("ConfigValue", "Vrednost")) {
                        try {
                            $col = $lo.ListColumns.Item($valCol)
                            if ($null -ne $col -and $null -ne $col.DataBodyRange) { $col.DataBodyRange.ClearContents() }
                        } catch {}
                    }
                }
                continue
            }

            # Poslovna tabela -> obrisi sve redove podataka.
            if ($null -ne $lo.DataBodyRange) {
                $lo.DataBodyRange.Delete() | Out-Null
                $cleared++
            }
        }
    }

    if (Test-Path $Output) { Remove-Item -LiteralPath $Output -Force }
    $wb.SaveAs($Output, $xlOpenXMLWorkbook)
    Write-Host "OK: $Output" -ForegroundColor Green
    Write-Host "  ispraznjeno tabela: $cleared, zadrzano: $kept" -ForegroundColor Green
    Write-Host "  VBA: nema (cist .xlsx). Pokrece ga OtkupApp.xlam preko marker tabele tblConfig." -ForegroundColor Green
}
finally {
    if ($null -ne $wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
