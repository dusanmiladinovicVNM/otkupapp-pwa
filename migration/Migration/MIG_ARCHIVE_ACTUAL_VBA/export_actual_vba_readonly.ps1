param(
  [Parameter(Mandatory=$true)][string]$Workbook,
  [string]$OutDir = ''
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# EXPORT ONLY: This script intentionally contains NO Excel Save, SaveAs,
# ReplaceLine, DeleteLines, InsertLines, Import, or VB code modification.
$target = [System.IO.Path]::GetFullPath($Workbook)
if (-not [System.IO.File]::Exists($target)) { throw ('Workbook not found: ' + $target) }
if ([System.IO.Path]::GetFileName($target) -ine 'AgriX_2026_ARHIVA_BUSINESS_VIEW.xlsm') {
    throw ('Wrong workbook name. Refusing: ' + $target)
}
if ($target -notmatch '[\\/]MigrationArchiveXLSM_05[\\/]') {
    throw ('Workbook is not in MigrationArchiveXLSM_05: ' + $target)
}

try {
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
} catch {
    throw 'Excel.Application not found in running COM instances. Open the _05 workbook manually first.'
}
$book = $null
foreach ($w in $excel.Workbooks) {
    if ([string]::Equals([System.IO.Path]::GetFullPath([string]$w.FullName), $target, [System.StringComparison]::OrdinalIgnoreCase)) {
        $book = $w
        break
    }
}
if ($null -eq $book) {
    throw ('The exact _05 workbook is not open in the attached Excel instance: ' + $target)
}

try {
    $project = $book.VBProject
    $null = $project.VBComponents.Count
} catch {
    throw 'VBProject read access is unavailable. Enable Trust access to VBA project object model temporarily.'
}
if ($project.Protection -ne 0) { throw 'VBA project protected. Refusing.' }

if ([string]::IsNullOrWhiteSpace($OutDir)) {
    $base = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($target), 'ACTUAL_VBA_EXPORT')
} else {
    $base = [System.IO.Path]::GetFullPath($OutDir)
}
# Never overwrite a previous source snapshot.
$tag = (Get-Date).ToString('yyyyMMdd_HHmmss_fff')
$folder = [System.IO.Path]::Combine($base, ('snapshot_' + $tag))
[System.IO.Directory]::CreateDirectory($folder) | Out-Null
$utf8bom = New-Object System.Text.UTF8Encoding($true)

$names = @('modAmbalaza','modIzvestaj','modScrIzvestaji','modAmbalazaUgovor')
$manifest = New-Object System.Collections.Generic.List[object]
$components = New-Object System.Collections.Generic.List[object]
foreach ($comp in $project.VBComponents) {
    $components.Add([pscustomobject]@{name=[string]$comp.Name; component_type=[int]$comp.Type})
}

foreach ($name in $names) {
    try {
        $component = $project.VBComponents.Item($name)
    } catch {
        throw ('Missing expected module, snapshot incomplete: ' + $name)
    }
    if ([int]$component.Type -ne 1) {
        throw ('Expected standard .bas module, refusing export of unexpected type: ' + $name)
    }
    $cm = $component.CodeModule
    $nlines = [int]$cm.CountOfLines
    $contents = if ($nlines -gt 0) { [string]$cm.Lines(1,$nlines) } else { '' }
    # Text snapshot is Unicode and reflects the live, in-memory VBA code.
    $txtfile = [System.IO.Path]::Combine($folder, ($name + '.txt'))
    [System.IO.File]::WriteAllText($txtfile, $contents, $utf8bom)
    # Native .bas export includes VB_Name and attributes, suitable for code review.
    $basfile = [System.IO.Path]::Combine($folder, ($name + '.bas'))
    $component.Export($basfile)
    $hashTxt = (Get-FileHash -LiteralPath $txtfile -Algorithm SHA256).Hash
    $hashBas = (Get-FileHash -LiteralPath $basfile -Algorithm SHA256).Hash
    $manifest.Add([pscustomobject]@{
        module=$name; lines=$nlines; txt_sha256=$hashTxt; bas_sha256=$hashBas
    })
    Write-Output ('EXPORTED ' + $name + ' (' + $nlines + ' lines)')
}

# Source fingerprint only; SHA is taken from disk, which should be unchanged.
$bookSha = (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash
$details = [ordered]@{
    status = 'ACTUAL_VBA_EXPORTED_READ_ONLY'
    workbook = $target
    workbook_sha256 = $bookSha
    project_name = [string]$project.Name
    export_time = (Get-Date).ToString('o')
    workbook_saved = $false
    excel_com_opened_new_instance = $false
    module_source_modified = $false
    original_business_tables_modified = $false
    modules = @($manifest.ToArray())
    component_inventory = @($components.ToArray())
}
$manifestPath = [System.IO.Path]::Combine($folder, 'SOURCE_MANIFEST.json')
[System.IO.File]::WriteAllText($manifestPath, ($details | ConvertTo-Json -Depth 8), $utf8bom)

# Save an easy-to-share self-contained archive beside the snapshot.
$zip = [System.IO.Path]::Combine($base, ('ACTUAL_VBA_SNAPSHOT_' + $tag + '.zip'))
Compress-Archive -Path (Join-Path $folder '*') -DestinationPath $zip -Force
Write-Output 'ACTUAL_VBA_EXPORTED_READ_ONLY'
Write-Output ('SNAPSHOT_DIR=' + $folder)
Write-Output ('SNAPSHOT_ZIP=' + $zip)
Write-Output ('WORKBOOK_SHA256=' + $bookSha)
Write-Output 'NO_WORKBOOK_SAVE_OR_VBA_EDIT_PERFORMED'
