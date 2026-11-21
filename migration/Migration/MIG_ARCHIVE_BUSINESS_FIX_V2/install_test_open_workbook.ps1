param(
  [Parameter(Mandatory=$true)][string]$Workbook,
  [switch]$DryRun
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Normalize-Code([string]$s) {
  if ($null -eq $s) { return '' }
  return $s.TrimStart([char]0xFEFF).Replace("`r", '').TrimEnd([char[]]@([char]10))
}

function Find-Line([string[]]$lines, [string]$fn, [string]$needle) {
  $starts = @()
  for ($i=0; $i -lt $lines.Count; $i++) {
    if ($lines[$i].TrimStart() -match ('^Public Function ' + [regex]::Escape($fn) + '\b')) { $starts += $i }
  }
  if ($starts.Count -ne 1) { throw "Preflight: function $fn occurrences: $($starts.Count)" }
  $start = $starts[0]
  $stop = -1
  for ($i=$start+1; $i -lt $lines.Count; $i++) {
    if ($lines[$i].Trim() -ceq 'End Function') { $stop = $i; break }
  }
  if ($stop -lt 0) { throw "Preflight: end of function $fn missing" }
  $matches = @()
  for ($i=$start; $i -le $stop; $i++) {
    if ($lines[$i] -ceq $needle) { $matches += $i }
  }
  if ($matches.Count -ne 1) {
    throw "Preflight: $fn expected exactly one occurrence; found $($matches.Count): $needle"
  }
  return [int]$matches[0]
}

$target = [IO.Path]::GetFullPath($Workbook)
if (-not (Test-Path -LiteralPath $target -PathType Leaf)) { throw "XLSM not found: $target" }
if (-not $target.EndsWith('AgriX_2026_ARHIVA_BUSINESS_VIEW_TEST.xlsm', [StringComparison]::OrdinalIgnoreCase)) {
  throw 'Fail closed: only AgriX_2026_ARHIVA_BUSINESS_VIEW_TEST.xlsm is allowed.'
}
if ($target -notmatch 'MigrationArchiveXLSM_05') {
  throw 'Fail closed: workbook must be in MigrationArchiveXLSM_05.'
}
$plan = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'patch_plan.json') -Raw -Encoding UTF8 | ConvertFrom-Json
if ($plan.version -ne 'MIG_ARCHIVE_KOOP_BUSINESS_V2') { throw 'Wrong patch plan' }

try {
  $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
} catch {
  throw 'Excel already open is required; no new Excel process will be launched.'
}
$wb = $null
foreach ($x in $excel.Workbooks) {
  if ([string]::Equals([IO.Path]::GetFullPath($x.FullName), $target, [StringComparison]::OrdinalIgnoreCase)) {
    $wb = $x
    break
  }
}
if ($null -eq $wb) { throw 'Target workbook is not open in the Excel COM instance. Close other Excel instances, open exactly _05 workbook, retry.' }
if ($wb.ReadOnly) { throw 'Workbook is read-only; installer must not modify it.' }
if ($null -eq $wb.VBProject) { throw 'Trust access to VBA project object model is not enabled.' }
$project = $wb.VBProject

# Before any edits, check that both modules match exactly the uploaded real-world snapshot.
$modules = @{}
$allActions = @()
foreach ($moduleName in @('modAmbalaza','modIzvestaj')) {
  $comp = $project.VBComponents.Item($moduleName)
  $cm = $comp.CodeModule
  $actual = Normalize-Code ([string]$cm.Lines(1, $cm.CountOfLines))
  $snapshotFile = Join-Path (Join-Path $PSScriptRoot 'snapshot') ($moduleName + '.txt')
  $expected = Normalize-Code ([IO.File]::ReadAllText($snapshotFile, [Text.Encoding]::UTF8))
  if (-not [string]::Equals($actual, $expected, [StringComparison]::Ordinal)) {
    throw "SOURCE_MISMATCH: $moduleName differs from submitted ACTUAL_VBA snapshot. No changes made. Export fresh snapshot."
  }
  $lines = [string[]]($actual -split "`n")
  $modules[$moduleName] = @{ cm=$cm; component=$comp; original=$actual; lines=$lines }
  $actions = @($plan.patches | Where-Object { $_.module -ceq $moduleName })
  foreach ($p in $actions) {
    $idx = Find-Line $lines ([string]$p.function) ([string]$p.old)
    $allActions += [PSCustomObject]@{module=$moduleName; function=$p.function; index=($idx+1); new_lines=@($p.new); old=$p.old}
    Write-Output "PREFLIGHT_MATCH $moduleName / $($p.function) / line $($idx+1)"
  }
}
Write-Output 'PREFLIGHT_PASS'
if ($DryRun) { Write-Output 'DRY_RUN_OK'; exit 0 }
if (-not $wb.Saved) { throw 'Workbook has unsaved changes. Close/save them safely first; do not patch on top of unsaved UI state.' }

# Export untouched actual module copies to an external backup folder first.
$backup = Join-Path ([IO.Path]::GetDirectoryName($target)) ('ARCHIVE_VBA_PATCH_BACKUP_' + (Get-Date -Format 'yyyyMMdd_HHmmss_fff'))
[IO.Directory]::CreateDirectory($backup) | Out-Null
foreach ($moduleName in @('modAmbalaza','modIzvestaj')) {
  $mod = $modules[$moduleName]
  $dest = Join-Path $backup ($moduleName + '.bas')
  $mod.component.Export($dest)
  [IO.File]::WriteAllText((Join-Path $backup ($moduleName + '.code.txt')), $mod.original, [Text.Encoding]::UTF8)
}
Write-Output "BACKUP_SAVED $backup"
$eventsBefore = $excel.EnableEvents
$hadEdits = $false
try {
  $excel.EnableEvents = $false
  foreach ($moduleName in @('modAmbalaza','modIzvestaj')) {
    $cm = $modules[$moduleName].cm
    $moduleActions = @($allActions | Where-Object { $_.module -eq $moduleName } | Sort-Object index -Descending)
    foreach ($action in $moduleActions) {
      $actualLine = [string]$cm.Lines($action.index, 1)
      if ($actualLine.TrimEnd([char[]]@([char]13,[char]10)) -cne $action.old) {
        throw "Code changed during apply at $moduleName line $($action.index)"
      }
      $cm.DeleteLines($action.index, 1)
      $hadEdits = $true
      $cm.InsertLines($action.index, ($action.new_lines -join "`r`n"))
    }
    $expectedBas = [IO.File]::ReadAllText((Join-Path (Join-Path $PSScriptRoot 'patched_preview') ($moduleName + '.bas')), [Text.Encoding]::UTF8)
    $expectedCode = Normalize-Code ($expectedBas.Substring($expectedBas.IndexOf("`n")+1))
    $actualCode = Normalize-Code ([string]$cm.Lines(1,$cm.CountOfLines))
    if (-not [string]::Equals($actualCode, $expectedCode, [StringComparison]::Ordinal)) {
      throw "POST_APPLY_DIFF $moduleName; not saved; original module exports in $backup"
    }
    Write-Output "VERIFIED $moduleName"
  }
  $wb.Save()
  Write-Output 'ARCHIVE_BUSINESS_CARD_PATCH_SAVED'
  Write-Output 'NEXT: Debug > Compile VBAProject; UI and PDF smoke test KOOP-00019; compare ending -10.'
} catch {
  if ($hadEdits) {
    Write-Output 'PATCH_FAILED_WITH_UNSAVED_VBA_EDITS: close _05 without saving and recover from copy _04 if necessary.'
  }
  throw
} finally {
  $excel.EnableEvents = $eventsBefore
}
