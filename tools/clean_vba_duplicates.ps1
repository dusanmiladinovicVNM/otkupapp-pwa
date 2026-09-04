<#
.SYNOPSIS
    Brise duplikate VBA komponenti (clsStmBtn1, clsLookupMenuBtn112, frmX2 ...)
    iz .xlsm radne sveske -- SPOLJA, dok Excel drzi svesku zatvorenu.

.DESCRIPTION
    Zasto spolja: VBComponents.Remove pozvan IZ VBA koda te iste sveske ne brise
    klase i forme odmah -- brisanje se stavlja u red i cesto se izgubi. Kad isti
    poziv dodje iz drugog procesa (ovaj skript preko COM-a), projekat nije "u
    izvrsavanju" i brisanje je trenutno.

    Pravilo brisanja je isto kao u modVbaTools.RemoveDuplicateModules i namerno
    je usko -- komponenta X<broj> se brise SAMO ako:
      1) NE postoji fajl X<broj>.bas/.cls/.frm u src-vba (nije praceni modul),
      2) postoji fajl X.bas/.cls/.frm u src-vba,
      3) komponenta X postoji u projektu (original je tu, kopija je visak).
    Document moduli (ThisWorkbook, Sheet1..) se ne diraju.

.PARAMETER Workbook
    Puna putanja do .xlsm fajla. Excel mora biti ZATVOREN (inace se otvara
    read-only i brisanje se ne moze snimiti).

.PARAMETER SrcVba
    Folder sa izvorom. Podrazumevano src-vba iz ovog repoa.

.PARAMETER Apply
    Bez ovog prekidaca skript samo prijavi sta bi obrisao. Sa njim pravi
    rezervnu kopiju sveske, brise i snima.

.EXAMPLE
    powershell -File tools\clean_vba_duplicates.ps1 -Workbook "C:\Users\Dusan\Desktop\AgriX - C002\Venivo\AgriX_2.39.0_testVenivno.xlsm"
    powershell -File tools\clean_vba_duplicates.ps1 -Workbook "...\AgriX_2.39.0_testVenivno.xlsm" -Apply

.NOTES
    PREDUSLOV: File > Options > Trust Center > Trust Center Settings >
               Macro Settings > "Trust access to the VBA project object model".
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    [string]$SrcVba = (Join-Path $PSScriptRoot '..\src-vba'),
    [switch]$Apply
)

$ErrorActionPreference = 'Stop'

function Get-NameBase([string]$s) {
    $i = $s.Length
    while ($i -gt 0 -and [char]::IsDigit($s[$i - 1])) { $i-- }
    return $s.Substring(0, $i)
}

$wbPath = (Resolve-Path -LiteralPath $Workbook).Path
$srcPath = (Resolve-Path -LiteralPath $SrcVba).Path

# 1) imena koja src-vba prati (bez ekstenzije)
$tracked = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
Get-ChildItem -LiteralPath $srcPath -File | Where-Object { $_.Extension -in '.bas', '.cls', '.frm' } |
    ForEach-Object { [void]$tracked.Add($_.BaseName) }
Write-Host "src-vba prati komponenti: $($tracked.Count)"

# 2) otvori svesku u zasebnoj instanci Excela, bez evenata (Workbook_Open ne sme da krene)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$wb = $null
try {
    if ($Apply) {
        $backup = [IO.Path]::Combine(
            [IO.Path]::GetDirectoryName($wbPath),
            "$([IO.Path]::GetFileNameWithoutExtension($wbPath))_predciscenja_$(Get-Date -Format 'yyyy-MM-dd_HHmm')$([IO.Path]::GetExtension($wbPath))")
        Copy-Item -LiteralPath $wbPath -Destination $backup
        Write-Host "Rezervna kopija: $backup"
    }

    $wb = $xl.Workbooks.Open($wbPath, 0, $false)
    if ($wb.ReadOnly) { throw "Sveska je otvorena read-only -- zatvori Excel pa pokreni ponovo." }

    try { $proj = $wb.VBProject } catch {
        throw "Nema programskog pristupa VBA projektu. Ukljuci: File > Options > Trust Center > Trust Center Settings > Macro Settings > 'Trust access to the VBA project object model'."
    }

    # 3) klasifikacija
    $dup = New-Object System.Collections.ArrayList
    $orphan = New-Object System.Collections.ArrayList
    $unknown = New-Object System.Collections.ArrayList
    $present = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($vbc in $proj.VBComponents) { [void]$present.Add($vbc.Name) }

    foreach ($vbc in $proj.VBComponents) {
        $nm = $vbc.Name
        if ($vbc.Type -eq 100) { continue }          # ThisWorkbook / Sheet* -- nikad
        if ($tracked.Contains($nm)) { continue }     # praceni modul -- original
        $base = Get-NameBase $nm
        if ($base.Length -gt 0 -and $base.Length -lt $nm.Length -and $tracked.Contains($base)) {
            if ($present.Contains($base)) { [void]$dup.Add($nm) }
            else { [void]$orphan.Add("$nm -> nema '$base' u projektu") }
        } else {
            [void]$unknown.Add($nm)
        }
    }

    $total = $proj.VBComponents.Count
    Write-Host ""
    Write-Host "Ukupno komponenti:                  $total"
    Write-Host "Duplikati za brisanje:              $($dup.Count)"
    Write-Host "Bez originala (RUCNO proveriti):    $($orphan.Count)"
    Write-Host "Van src-vba, nije duplikat:         $($unknown.Count)"

    $report = [IO.Path]::Combine([IO.Path]::GetDirectoryName($wbPath), 'vba_duplikati.txt')
    @(
        "Duplikati - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        "Sveska: $wbPath"
        ""
        "[ZA BRISANJE] ($($dup.Count))"
        $(if ($dup.Count) { $dup } else { '  (nema)' })
        ""
        "[BEZ ORIGINALA - rucna provera] ($($orphan.Count))"
        $(if ($orphan.Count) { $orphan } else { '  (nema)' })
        ""
        "[VAN src-vba - ne dira se] ($($unknown.Count))"
        $(if ($unknown.Count) { $unknown } else { '  (nema)' })
    ) | Set-Content -LiteralPath $report -Encoding ASCII
    Write-Host "Spisak: $report"

    if (-not $Apply) {
        Write-Host ""
        Write-Host "PREGLED -- nista nije obrisano. Za brisanje dodaj -Apply."
        $wb.Close($false); $wb = $null
        return
    }

    # 4) brisanje
    $removed = 0
    $failed = New-Object System.Collections.ArrayList
    foreach ($nm in $dup) {
        try {
            $proj.VBComponents.Remove($proj.VBComponents.Item($nm))
            $removed++
        } catch {
            [void]$failed.Add("$nm -> $($_.Exception.Message)")
        }
    }

    $wb.Save()
    $after = $proj.VBComponents.Count
    Write-Host ""
    Write-Host "Obrisano:   $removed"
    Write-Host "Neuspesno:  $($failed.Count)"
    if ($failed.Count) { $failed | Select-Object -First 20 | ForEach-Object { Write-Host "  $_" } }
    Write-Host "Komponenti posle ciscenja: $after (pre: $total)"
    if ($after -ge $total) { Write-Warning "Broj komponenti se nije smanjio -- brisanje NIJE proslo. Sveska je snimljena kakva jeste; vrati rezervnu kopiju." }

    $wb.Close($true); $wb = $null
} finally {
    if ($wb) { try { $wb.Close($false) } catch { } }
    try { $xl.Quit() } catch { }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect()
}
