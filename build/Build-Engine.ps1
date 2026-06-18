<#
.SYNOPSIS
    Gradi OtkupApp.xlam (zajednicki KOD add-in) iz src-vba/ izvora.

.DESCRIPTION
    Pravi prazan workbook, importuje sve VBA module/klase/forme iz src-vba/,
    ubacuje kod ThisWorkbook (engine) modula, postavlja IsAddin=True i snima
    kao .xlam. Time .xlam postaje *artefakt build-a*, a izvor istine ostaje git.

    Posle build-a aplikacija se azurira tako sto se OVAJ .xlam zameni na PC-u.
    Klijentski .xlsx (podaci/config) se NE dira.

.PREREQUISITES
    - Windows + Microsoft Excel (desktop).
    - "Trust access to the VBA project object model" mora biti ukljuceno
      (File > Options > Trust Center > Trust Center Settings > Macro Settings).
      Skripta pokusava da postavi taj registry kljuc automatski.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File build\Build-Engine.ps1
    powershell -ExecutionPolicy Bypass -File build\Build-Engine.ps1 -SrcDir src-vba -OutFile dist\OtkupApp.xlam
#>
[CmdletBinding()]
param(
    [string]$SrcDir = (Join-Path $PSScriptRoot "..\src-vba"),
    [string]$OutFile = (Join-Path $PSScriptRoot "..\dist\OtkupApp.xlam")
)

$ErrorActionPreference = "Stop"
$xlOpenXMLAddIn = 55   # FileFormat za .xlam

function Resolve-FullPath([string]$p) {
    return [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $p))
}

$SrcDir = Resolve-FullPath $SrcDir
$OutFile = Resolve-FullPath $OutFile
$OutDir = Split-Path $OutFile -Parent
if (-not (Test-Path $SrcDir)) { throw "Nije nadjen src-vba folder: $SrcDir" }
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }

Write-Host "Izvor : $SrcDir"
Write-Host "Izlaz : $OutFile"

# --- Pokusaj da ukljucis VBOM pristup (potrebno za Import VBProject-a) ---
try {
    $ver = (New-Object -ComObject Excel.Application).Version
} catch { throw "Excel nije dostupan preko COM-a. Da li je Excel instaliran?" }
$regPath = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security"
try {
    if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
    Set-ItemProperty -Path $regPath -Name "AccessVBOM" -Value 1 -Type DWord -Force
    Write-Host "AccessVBOM=1 postavljen za Excel $ver."
} catch {
    Write-Warning "Ne mogu da postavim AccessVBOM. Ukljucite rucno 'Trust access to the VBA project object model'."
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
# Onemoguci makroe iz importovanih fajlova tokom build-a.
try { $excel.AutomationSecurity = 3 } catch {}   # msoAutomationSecurityForceDisable

$wb = $null
try {
    $wb = $excel.Workbooks.Add()

    $proj = $wb.VBProject
    if ($null -eq $proj) { throw "Nema pristupa VBProject-u. Ukljucite 'Trust access to the VBA project object model'." }

    # 1) ThisWorkbook (engine) kod -> ubaci u postojeci document modul.
    $tw = Join-Path $SrcDir "ThisWorkbook.doccls"
    if (Test-Path $tw) {
        $lines = Get-Content -LiteralPath $tw
        # Preskoci class header (VERSION/BEGIN/END/Attribute...) - document modul
        # ga ne prima preko AddFromString; zadrzavamo kod od prve ne-header linije.
        $body = New-Object System.Collections.Generic.List[string]
        $skip = $true
        foreach ($ln in $lines) {
            if ($skip -and ($ln -match '^\s*(VERSION|BEGIN|END|MultiUse|Attribute)\b' -or $ln -match '^\s*$')) { continue }
            $skip = $false
            $body.Add($ln)
        }
        $twComp = $proj.VBComponents.Item("ThisWorkbook").CodeModule
        if ($twComp.CountOfLines -gt 0) { $twComp.DeleteLines(1, $twComp.CountOfLines) }
        $twComp.AddFromString([string]::Join("`r`n", $body))
        Write-Host "  + ThisWorkbook (engine) kod ubacen"
    }

    # 2) Importuj sve module/klase/forme.
    $files = Get-ChildItem -LiteralPath $SrcDir -File |
        Where-Object { $_.Extension -in @(".bas", ".cls", ".frm") } |
        Sort-Object Name
    foreach ($f in $files) {
        $proj.VBComponents.Import($f.FullName) | Out-Null
        Write-Host "  + $($f.Name)"
    }

    # 3) Add-in mod + snimanje.
    $wb.IsAddin = $true
    if (Test-Path $OutFile) { Remove-Item -LiteralPath $OutFile -Force }
    $wb.SaveAs($OutFile, $xlOpenXMLAddIn)
    Write-Host "`nOK: napravljen $OutFile" -ForegroundColor Green
    Write-Host "Sledeci korak: potpisati .xlam VBA publisher sertifikatom (VBE > Tools > Digital Signature)." -ForegroundColor Yellow
}
finally {
    if ($null -ne $wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
