param([Parameter(Mandatory=$true)][string]$Workbook)
$ErrorActionPreference='Stop'
$excel=New-Object -ComObject Excel.Application
$excel.Visible=$true
$excel.DisplayAlerts=$true
$excel.EnableEvents=$false
$excel.AutomationSecurity=3 # macros NEVER RUN through this launcher
$excel.AskToUpdateLinks=$false
$resolved=(Resolve-Path -LiteralPath $Workbook).Path
$wb=$excel.Workbooks.Open($resolved,0,$true)
$wb.Worksheets.Item('ARHIVA_START').Activate()
Write-Output ('ARHIVA OPEN READONLY: '+$wb.FullName)
Write-Output 'Svi makroi iskljuceni. Pregledaj ARHIVA_* listove.'
# Excel stays open and visible for the user. No Save/SaveAs action.
