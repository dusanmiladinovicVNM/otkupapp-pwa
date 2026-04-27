# Setup-OtkupApp.ps1
# OtkupApp initial workstation setup
#
# Run as Administrator if possible.
# This script:
#   - creates C:\OtkupApp folder structure
#   - copies OtkupApp.xlsm
#   - unblocks the workbook
#   - installs VBA publisher certificate if present
#   - adds Excel Trusted Location
#   - creates Desktop shortcut

param(
    [string]$InstallRoot = "C:\OtkupApp",
    [string]$ExcelVersion = "16.0"
)

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "=== OtkupApp setup started ==="
Write-Host ""

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

$SourceWorkbook = Join-Path $ScriptRoot "OtkupApp.xlsm"
$SourceCert = Join-Path $ScriptRoot "OtkupApp-VBA-Publisher.cer"
$TargetWorkbook = Join-Path $InstallRoot "OtkupApp.xlsm"

$Folders = @(
    $InstallRoot,
    "$InstallRoot\Backups",
    "$InstallRoot\Logs",
    "$InstallRoot\Journal",
    "$InstallRoot\Export",
    "$InstallRoot\Temp",
    "$InstallRoot\Secrets",
    "$InstallRoot\Bank_Izvodi",
    "$InstallRoot\Bank_Izvodi\Inbox",
    "$InstallRoot\Bank_Izvodi\Processed",
    "$InstallRoot\Bank_Izvodi\Error"
)

foreach ($folder in $Folders) {
    if (!(Test-Path $folder)) {
        New-Item -ItemType Directory -Path $folder | Out-Null
        Write-Host "Created folder: $folder"
    } else {
        Write-Host "Folder exists: $folder"
    }
}

if (!(Test-Path $SourceWorkbook)) {
    throw "Missing OtkupApp.xlsm in install folder: $SourceWorkbook"
}

Copy-Item $SourceWorkbook $TargetWorkbook -Force
Write-Host "Copied workbook to: $TargetWorkbook"

try {
    Unblock-File -Path $TargetWorkbook -ErrorAction SilentlyContinue
    Write-Host "Workbook unblocked."
} catch {
    Write-Warning "Could not unblock workbook: $($_.Exception.Message)"
}

if (Test-Path $SourceCert) {
    try {
        Import-Certificate -FilePath $SourceCert -CertStoreLocation "Cert:\CurrentUser\Root" | Out-Null
        Import-Certificate -FilePath $SourceCert -CertStoreLocation "Cert:\CurrentUser\TrustedPublisher" | Out-Null
        Write-Host "Installed OtkupApp certificate for CurrentUser."
    } catch {
        Write-Warning "Could not install certificate: $($_.Exception.Message)"
    }
} else {
    Write-Warning "Certificate not found. Skipping certificate install: $SourceCert"
}

# Excel Trusted Location
$TrustedLocationKey = "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security\Trusted Locations\OtkupApp"

if (!(Test-Path $TrustedLocationKey)) {
    New-Item -Path $TrustedLocationKey -Force | Out-Null
}

New-ItemProperty -Path $TrustedLocationKey -Name "Path" -Value "$InstallRoot\" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $TrustedLocationKey -Name "AllowSubfolders" -Value 1 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path $TrustedLocationKey -Name "Description" -Value "OtkupApp trusted location" -PropertyType String -Force | Out-Null

Write-Host "Added Excel Trusted Location: $InstallRoot"

# Desktop shortcut
$Desktop = [Environment]::GetFolderPath("Desktop")
$ShortcutPath = Join-Path $Desktop "OtkupApp.lnk"

$Shell = New-Object -ComObject WScript.Shell
$Shortcut = $Shell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $TargetWorkbook
$Shortcut.WorkingDirectory = $InstallRoot
$Shortcut.Description = "Otvori OtkupApp"
$Shortcut.Save()

Write-Host "Created desktop shortcut: $ShortcutPath"

Write-Host ""
Write-Host "=== OtkupApp setup completed ==="
Write-Host ""
Write-Host "Next:"
Write-Host "1. Open OtkupApp from Desktop shortcut."
Write-Host "2. Run SetupNewPC inside OtkupApp."
Write-Host "3. Configure bank folders, Sheets OAuth config and SEF config."
Write-Host ""
Pause
