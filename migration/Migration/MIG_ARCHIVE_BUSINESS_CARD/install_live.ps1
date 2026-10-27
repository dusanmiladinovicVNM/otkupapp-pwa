param(
  [Parameter(Mandatory=$true)][string]$Workbook,
  [switch]$DryRun
)
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest

function Require-Once([string]$whole,[string]$needle,[int]$expected=1) {
  $count=[regex]::Matches($whole,[regex]::Escape($needle)).Count
  if($count -ne $expected){throw "Preflight: expected $expected occurrences, found ${count}: $needle"}
}
function Find-LiteralLine($cm,[string]$needle) {
  $lines=$cm.Lines(1,$cm.CountOfLines) -split '\r\n|\n|\r'
  $hits=@()
  for($i=0;$i -lt $lines.Length;$i++){
    if($lines[$i].Trim() -ceq $needle){$hits+=($i+1)}
  }
  if($hits.Count -ne 1){throw "Expected exact one line, found $($hits.Count): $needle"}
  return [int]$hits[0]
}
function Apply-ReplaceLine($cm,[string]$before,[string]$after){
  $n=Find-LiteralLine $cm $before
  $cm.ReplaceLine($n,$after)
}
function Apply-InsertAfter($cm,[string]$anchor,[string]$newText){
  $n=Find-LiteralLine $cm $anchor
  $cm.InsertLines($n+1,$newText)
}

$wanted=[IO.Path]::GetFullPath($Workbook)
if(-not [IO.File]::Exists($wanted)){throw "Workbook missing: $wanted"}
if([IO.Path]::GetFileName($wanted) -cne 'AgriX_2026_ARHIVA_BUSINESS_VIEW.xlsm'){
  throw 'Safety gate: installer only edits AgriX_2026_ARHIVA_BUSINESS_VIEW.xlsm (a separate copy).'
}
if($wanted -notmatch '[\\/]MigrationArchiveXLSM_05[\\/]'){
  throw 'Safety gate: workbook must be in MigrationArchiveXLSM_05.'
}

try { $xl=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') }
catch { throw 'Open the PATCH copy manually in Excel first; installer will NOT launch Excel (modLogo dialog avoidance).' }
$wb=$null
foreach($book in $xl.Workbooks){
  if([string]::Equals([IO.Path]::GetFullPath([string]$book.FullName),$wanted,[StringComparison]::OrdinalIgnoreCase)){
    $wb=$book;break
  }
}
if($null -eq $wb){throw "Exact archive copy is not open in Excel: $wanted"}
if($wb.ReadOnly){throw 'Archive copy opened ReadOnly. Close, then reopen for editing.'}
try { $project=$wb.VBProject }
catch {throw 'VBProject access denied. In Excel enable Trust access to the VBA project object model temporarily.'}
if($project.Protection -ne 0){throw 'VBA project is locked/protected; refusing modification.'}
try {
  $amb=$project.VBComponents.Item('modAmbalaza').CodeModule
  $izv=$project.VBComponents.Item('modIzvestaj').CodeModule
} catch {throw 'Missing required VBA standard module modAmbalaza or modIzvestaj.'}

$oldAmb=[string]$amb.Lines(1,$amb.CountOfLines)
$oldIzv=[string]$izv.Lines(1,$izv.CountOfLines)
$beforeHeader='Public Function AmbKretanjaNaloga(ByVal tip As String, ByVal id As String) As Variant'
$beforeDim='Dim cStorno As Long, cAmbID As Long'
$beforeAmbId='cAmbID = kolIdx(COL_AMB_ID)'
$beforeAppend='If znak <> 0 Then poz.Add Array(i, znak)'
$beforeReport='ambData = modAmbalaza.AmbKretanjaNaloga(AMB_NALOG_KOOPERANT, Trim$(kooperantID))'
Require-Once $oldAmb $beforeHeader
Require-Once $oldAmb $beforeDim
Require-Once $oldAmb $beforeAmbId
Require-Once $oldAmb $beforeAppend
Require-Once $oldIzv $beforeReport 2
if($oldAmb.Contains('samoPoslovnaKartica') -or $oldIzv.Contains('Trim$(kooperantID), True')){
  throw 'Code already appears patched. No second application permitted.'
}
Write-Output "PREFLIGHT_PASS - VBA source anchors found; workbook=$wanted"
if($DryRun){Write-Output 'DRY_RUN_OK - no changes made';exit 0}

# Save original VBA SOURCE text for manual recovery; source XLSM in _04 is never touched.
$backupDir=Join-Path (Split-Path $wanted -Parent) 'MIG_ARCHIVE_CODE_BACKUP'
[IO.Directory]::CreateDirectory($backupDir) | Out-Null
$utf8BOM=New-Object System.Text.UTF8Encoding($true)
[IO.File]::WriteAllText((Join-Path $backupDir 'modAmbalaza_before.txt'),$oldAmb,$utf8BOM)
[IO.File]::WriteAllText((Join-Path $backupDir 'modIzvestaj_before.txt'),$oldIzv,$utf8BOM)

# Minimal scoped change. Module identity and attributes remain in place;
# there is no remove/reimport of large code modules.
$filterExtra=@'
            If samoPoslovnaKartica And _
               StrComp(Trim$(tip), AMB_NALOG_KOOPERANT, vbTextCompare) = 0 And _
               StrComp(AmbText(data(i, cVrstaKretanja)), AMB_VK_ULAZ_TUDJE, vbTextCompare) = 0 Then
                ' Technical deficit coverage: ledger only; not business turnover.
            Else
                poz.Add Array(i, znak)
            End If
        End If
'@

try {
  # Work bottom-up so earlier original line positions remain addressable.
  $targetLine=Find-LiteralLine $amb $beforeAppend
  $amb.ReplaceLine($targetLine, 'If znak <> 0 Then')
  $amb.InsertLines($targetLine+1, $filterExtra)
  Apply-InsertAfter $amb $beforeAmbId '    cVrstaKretanja = kolIdx(COL_AMB_VRSTA_KRETANJA)'
  Apply-ReplaceLine $amb $beforeDim 'Dim cStorno As Long, cAmbID As Long, cVrstaKretanja As Long'
  Apply-ReplaceLine $amb $beforeHeader 'Public Function AmbKretanjaNaloga(ByVal tip As String, ByVal id As String, Optional ByVal samoPoslovnaKartica As Boolean = False) As Variant'

  $modifiedLine=$beforeReport -replace '\)\)$', '), True)'
  # Two same lines, no shifts after ReplaceLine. Always replace BOTH.
  $newCount=0
  for($i=1;$i -le $izv.CountOfLines;$i++){
    if($izv.Lines($i,1).Trim() -ceq $beforeReport){
      $izv.ReplaceLine($i,$modifiedLine)
      $newCount++
    }
  }
  if($newCount -ne 2){throw "Report replacement mismatch: $newCount"}

  $checkAmb=[string]$amb.Lines(1,$amb.CountOfLines)
  $checkIzv=[string]$izv.Lines(1,$izv.CountOfLines)
  Require-Once $checkAmb 'Optional ByVal samoPoslovnaKartica As Boolean = False'
  Require-Once $checkAmb 'cVrstaKretanja = kolIdx(COL_AMB_VRSTA_KRETANJA)'
  Require-Once $checkAmb 'poz.Add Array(i, znak)'
  Require-Once $checkIzv 'AmbKretanjaNaloga(AMB_NALOG_KOOPERANT, Trim$(kooperantID), True)' 2

  $previousEvents=$xl.EnableEvents
  try {
    $xl.EnableEvents=$false
    $wb.Save()
  }
  finally { $xl.EnableEvents=$previousEvents }
  Write-Output 'ARCHIVE_BUSINESS_CARD_PATCH_SAVED'
  Write-Output "OUTPUT=$wanted"
  Write-Output 'NEXT: Debug > Compile VBAProject; then KOOP-00019 checks for 08.07 and 11.07.'
} catch {
  Write-Error "PATCH STOP: $($_.Exception.Message). If not saved, CLOSE ARCHIVE WITHOUT SAVING and recreate from the _04 copy."
  exit 1
}
