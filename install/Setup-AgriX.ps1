# Setup-AgriX.ps1
# AgriX initial workstation setup
#
# Run as Administrator if possible.
# This script:
#   - creates C:\AgriX folder structure
#   - installs AgriX.xlsm ONLY when it is not already there, and unblocks it
#     (existing workbook carries client data; -ReplaceWorkbook replaces it, but
#      only after a SHA256-verified timestamped backup)
#   - copies Tools\poppler (pdftotext.exe) next to the workbook
#   - copies docs\ (if present)
#   - installs VBA publisher certificate if present
#   - adds Excel Trusted Location
#   - creates Desktop shortcut
#   - writes an install log and prints a PASS/FAIL summary

param(
    [string]$InstallRoot = "C:\AgriX",
    [string]$ExcelVersion = "16.0",
    # The workbook carries client data. Default is: if it is already there,
    # leave it alone. This switch replaces it, after a verified backup.
    [switch]$ReplaceWorkbook,
    # Runs the workbook-install decision over a temp folder and exits. No admin
    # rights, no registry, no certificate -- safe to run anywhere, incl. CI.
    [switch]$SelfTest
)

$ErrorActionPreference = "Stop"

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

# ------------------------------------------------------------------
# Workbook install decision, separated so -SelfTest can exercise it.
#
# Until v2.28.5 this was one line: Copy-Item $SourceWorkbook $TargetWorkbook
# -Force. Re-running setup is normal and expected (Poppler, certificate,
# Trusted Location, shortcut all get refreshed that way), so that one line was
# a direct path to overwriting a live client workbook with the packaged one.
#
# Returns one of: "installed" | "kept-existing" | "replaced:<backup path>"
# ------------------------------------------------------------------
function Install-Workbook {
    param(
        [Parameter(Mandatory=$true)][string]$Source,
        [Parameter(Mandatory=$true)][string]$Target,
        [Parameter(Mandatory=$true)][string]$BackupDir,
        [bool]$Replace = $false
    )

    if (!(Test-Path $Source)) {
        throw "Missing AgriX.xlsm in install folder: $Source"
    }

    if (!(Test-Path $Target)) {
        Copy-Item $Source $Target -Force
        return "installed"
    }

    if (-not $Replace) {
        return "kept-existing"
    }

    if (!(Test-Path $BackupDir)) {
        New-Item -ItemType Directory -Force -Path $BackupDir | Out-Null
    }
    $stamp  = Get-Date -Format "yyyy-MM-dd_HHmmss"
    $backup = Join-Path $BackupDir ("AgriX_pre-replace_" + $stamp + ".xlsm")
    Copy-Item $Target $backup -ErrorAction Stop

    # An unverified backup is a promise, not a copy. Compare content, not size:
    # a truncated copy of an .xlsm can match on neither but a partial flush can.
    $liveHash = (Get-FileHash -Path $Target -Algorithm SHA256).Hash
    $bakHash  = (Get-FileHash -Path $backup -Algorithm SHA256).Hash
    if ($liveHash -ne $bakHash) {
        throw ("Backup verification FAILED: {0} does not match the live workbook. Workbook NOT replaced." -f $backup)
    }

    Copy-Item $Source $Target -Force
    return ("replaced:" + $backup)
}

$script:StPass = 0
$script:StFail = 0
function Assert-SelfTest {
    param([bool]$Ok, [string]$Label)
    if ($Ok) {
        $script:StPass++
        Write-Host ("  ok    " + $Label)
    } else {
        $script:StFail++
        Write-Host ("  PAO   " + $Label) -ForegroundColor Red
    }
}

function Invoke-SelfTest {
    $root = Join-Path ([System.IO.Path]::GetTempPath()) ("AgriXSelfTest_" + [Guid]::NewGuid().ToString("N"))
    $pkg  = Join-Path $root "pkg"
    $dst  = Join-Path $root "dst"
    $bak  = Join-Path $dst "Backups"
    New-Item -ItemType Directory -Force -Path $pkg | Out-Null
    New-Item -ItemType Directory -Force -Path $dst | Out-Null

    $src = Join-Path $pkg "AgriX.xlsm"
    $tgt = Join-Path $dst "AgriX.xlsm"
    Set-Content -Path $src -Value "PAKET" -Encoding utf8

    Write-Host "Setup-AgriX self-test (workbook install decision)"
    try {
        # 1. cist racunar: sveska se instalira
        $r = Install-Workbook -Source $src -Target $tgt -BackupDir $bak -Replace $false
        Assert-SelfTest ($r -eq "installed") "fresh target -> installed"
        Assert-SelfTest ((Get-Content $tgt -Raw).Trim() -eq "PAKET") "fresh target gets package content"

        # 2. ZIVA sveska, default rerun: mora ostati NETAKNUTA.
        #    Ovo je slucaj zbog kojeg kapija postoji -- support ponovo pokrene
        #    setup zbog Poppler-a i pregazi klijentove podatke.
        Set-Content -Path $tgt -Value "SENTINEL-KLIJENT" -Encoding utf8
        $r = Install-Workbook -Source $src -Target $tgt -BackupDir $bak -Replace $false
        Assert-SelfTest ($r -eq "kept-existing") "existing target -> kept-existing"
        Assert-SelfTest ((Get-Content $tgt -Raw).Trim() -eq "SENTINEL-KLIJENT") "DEFAULT RERUN MUST NOT TOUCH LIVE WORKBOOK"

        # 3. -ReplaceWorkbook: zamenjuje, ali backup mora nositi sentinel
        $r = Install-Workbook -Source $src -Target $tgt -BackupDir $bak -Replace $true
        Assert-SelfTest ($r -like "replaced:*") "-ReplaceWorkbook -> replaced"
        Assert-SelfTest ((Get-Content $tgt -Raw).Trim() -eq "PAKET") "replace puts package content in place"
        $bakFile = $r.Substring("replaced:".Length)
        Assert-SelfTest (Test-Path $bakFile) "backup file exists"
        Assert-SelfTest ((Get-Content $bakFile -Raw).Trim() -eq "SENTINEL-KLIJENT") "backup carries the REPLACED workbook, not the package"

        # 4. nema paketne sveske -> jasna greska, ne tiho preskakanje
        Remove-Item $src -Force
        $threw = $false
        try { Install-Workbook -Source $src -Target $tgt -BackupDir $bak -Replace $false | Out-Null }
        catch { $threw = $true }
        Assert-SelfTest $threw "missing package workbook -> throws"
    }
    finally {
        Remove-Item $root -Recurse -Force -ErrorAction SilentlyContinue
    }

    Write-Host ("PROSLO: {0}   PALO: {1}" -f $script:StPass, $script:StFail)
}

# Dispatch PRE nego sto skripta napravi InstallRoot: -SelfTest ne sme da
# dotakne ni C:\AgriX ni log te masine.
if ($SelfTest) {
    Invoke-SelfTest
    if ($script:StFail -gt 0) { exit 1 }
    exit 0
}

$SourceWorkbook = Join-Path $ScriptRoot "AgriX.xlsm"
$SourceCert     = Join-Path $ScriptRoot "AgriX-VBA-Publisher.cer"
$SourceTools    = Join-Path $ScriptRoot "Tools"
$SourceDocs     = Join-Path $ScriptRoot "docs"
$TargetWorkbook = Join-Path $InstallRoot "AgriX.xlsm"

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

# --- install log: create InstallRoot + Logs first so everything is logged ---
foreach ($f in @($InstallRoot, "$InstallRoot\Logs")) {
    if (!(Test-Path $f)) { New-Item -ItemType Directory -Path $f | Out-Null }
}
$LogFile = Join-Path $InstallRoot "Logs\install-log.txt"
$script:Warnings = @()

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Add-Content -Path $LogFile -Value $line
    if ($Level -eq "WARN")       { Write-Warning $Message }
    elseif ($Level -eq "ERROR")  { Write-Host $Message -ForegroundColor Red }
    else                         { Write-Host $line }
}
function Add-WarnLog { param([string]$Message) $script:Warnings += $Message; Write-Log $Message "WARN" }

Write-Log ""
Write-Log "=== AgriX setup started ==="
Write-Log ("Machine: {0}  User: {1}  InstallRoot: {2}" -f $env:COMPUTERNAME, $env:USERNAME, $InstallRoot)

try {
    # --- folders ---
    foreach ($folder in $Folders) {
        if (!(Test-Path $folder)) {
            New-Item -ItemType Directory -Path $folder | Out-Null
            Write-Log "Created folder: $folder"
        } else {
            Write-Log "Folder exists: $folder"
        }
    }

    # --- workbook (critical) ---
    # Decision + verified backup live in Install-Workbook (see its comment).
    $wbOutcome = Install-Workbook -Source $SourceWorkbook -Target $TargetWorkbook `
                                  -BackupDir "$InstallRoot\Backups" `
                                  -Replace ([bool]$ReplaceWorkbook)
    if ($wbOutcome -eq "kept-existing") {
        Add-WarnLog ("EXISTING INSTALLATION: {0} was NOT touched (it carries client data). Folders, Poppler, certificate, Trusted Location and shortcut are still refreshed. To replace the workbook, re-run with -ReplaceWorkbook -- it takes a SHA256-verified backup into {1}\Backups first." -f $TargetWorkbook, $InstallRoot)
    } elseif ($wbOutcome -like "replaced:*") {
        Write-Log ("Replaced workbook: {0}  (verified backup: {1})" -f $TargetWorkbook, $wbOutcome.Substring("replaced:".Length))
    } else {
        Write-Log "Copied workbook to: $TargetWorkbook"
    }

    try {
        Unblock-File -Path $TargetWorkbook -ErrorAction SilentlyContinue
        Write-Log "Workbook unblocked."
    } catch {
        Add-WarnLog "Could not unblock workbook: $($_.Exception.Message)"
    }

    # --- Poppler (pdftotext.exe) for bank statement import ---
    # Tools\ must sit next to the workbook: VBA (ResolvePdfToTextExePath) resolves the
    # default as <workbook>\Tools\poppler\Library\bin\pdftotext.exe. Non-fatal: some
    # deployments use a PATH-installed pdftotext or set PDFTOTEXT_EXE_PATH manually.
    $TargetTools = Join-Path $InstallRoot "Tools"
    $PopplerExe  = Join-Path $InstallRoot "Tools\poppler\Library\bin\pdftotext.exe"

    if (Test-Path $SourceTools) {
        Copy-Item $SourceTools $InstallRoot -Recurse -Force
        if (Test-Path $PopplerExe) {
            Write-Log "Copied Poppler tools to: $TargetTools"
        } else {
            Add-WarnLog "Tools copied but pdftotext.exe not at expected path: $PopplerExe"
        }
    } else {
        Add-WarnLog "Poppler tools not in package (bank import needs pdftotext.exe): $SourceTools"
    }

    # --- docs (non-fatal) ---
    if (Test-Path $SourceDocs) {
        Copy-Item $SourceDocs $InstallRoot -Recurse -Force
        Write-Log ("Copied docs to: {0}" -f (Join-Path $InstallRoot "docs"))
    } else {
        Write-Log "No docs\ folder in package (skipping)."
    }

    # --- VBA publisher certificate ---
    if (Test-Path $SourceCert) {
        try {
            Import-Certificate -FilePath $SourceCert -CertStoreLocation "Cert:\CurrentUser\Root" | Out-Null
            Import-Certificate -FilePath $SourceCert -CertStoreLocation "Cert:\CurrentUser\TrustedPublisher" | Out-Null
            Write-Log "Installed AgriX certificate for CurrentUser."
        } catch {
            Add-WarnLog "Could not install certificate: $($_.Exception.Message)"
        }
    } else {
        Add-WarnLog "Certificate not found (unsigned workbook -> macros may be blocked): $SourceCert"
    }

    # --- Excel Trusted Location ---
    $TrustedLocationKey = "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security\Trusted Locations\AgriX"
    if (!(Test-Path $TrustedLocationKey)) {
        New-Item -Path $TrustedLocationKey -Force | Out-Null
    }
    New-ItemProperty -Path $TrustedLocationKey -Name "Path" -Value "$InstallRoot\" -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $TrustedLocationKey -Name "AllowSubfolders" -Value 1 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $TrustedLocationKey -Name "Description" -Value "AgriX trusted location" -PropertyType String -Force | Out-Null
    Write-Log "Added Excel Trusted Location: $InstallRoot"

    # --- Desktop shortcut ---
    $Desktop = [Environment]::GetFolderPath("Desktop")
    $ShortcutPath = Join-Path $Desktop "AgriX.lnk"
    $Shell = New-Object -ComObject WScript.Shell
    $Shortcut = $Shell.CreateShortcut($ShortcutPath)
    $Shortcut.TargetPath = $TargetWorkbook
    $Shortcut.WorkingDirectory = $InstallRoot
    $Shortcut.Description = "Otvori AgriX"
    $Shortcut.Save()
    Write-Log "Created desktop shortcut: $ShortcutPath"

    # --- final pdftotext verification (non-fatal) ---
    if (!(Test-Path $PopplerExe)) {
        Add-WarnLog "pdftotext.exe not verified at $PopplerExe (set it in the app: SetupPopplerInteractive)."
    }

    # --- PASS/FAIL summary ---
    Write-Log ""
    Write-Log "=== AgriX setup completed ==="
    if ($script:Warnings.Count -eq 0) {
        Write-Log "RESULT: PASS (no warnings)."
    } else {
        Write-Log ("RESULT: PASS WITH WARNINGS ({0}):" -f $script:Warnings.Count)
        foreach ($w in $script:Warnings) { Write-Log "  - $w" "WARN" }
    }

    Write-Host ""
    Write-Host "Next steps (inside AgriX):"
    Write-Host "  1. Open AgriX from the Desktop shortcut."
    Write-Host "  2. Accept the first-run prompt -> SetupNewPC (must end with APP_SETUP_COMPLETED = DA)."
    Write-Host "  3. Poppler: Alt+F8 -> SetupPopplerInteractive (auto if Tools\poppler sits next to the"
    Write-Host "     workbook), or Maticni podaci -> Podesavanja -> 'Izaberi Poppler'."
    Write-Host "  4. Bank import (if used): set up Google Drive for Desktop for the client's 01_Bank"
    Write-Host "     folder, mark 00_Inbox 'Available offline', then in Podesavanja -> 'Banka / lokalno'"
    Write-Host "     use the '...' button to set BANKA_DRIVE_SOURCE_PATH to the local 01_Bank path."
    Write-Host "  5. Verify links: Alt+F8 -> TestServerLink (Google / GAS / bank Drive folder)."
    Write-Host ""
    Write-Host "Bank two-GAS linking + Drive for Desktop details:"
    Write-Host "  docs\production-runbook-banka-import-setup.md  and  docs\DESKTOP_SETUP_REFERENCE.md"
    Write-Host ""
    Write-Host ("Install log: {0}" -f $LogFile)
    Write-Host ""
}
catch {
    Write-Log "=== AgriX setup FAILED ===" "ERROR"
    Write-Log ("RESULT: FAIL - {0}" -f $_.Exception.Message) "ERROR"
    Write-Host ""
    Write-Host ("Install log: {0}" -f $LogFile)
    Pause
    exit 1
}

Pause
