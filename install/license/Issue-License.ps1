<#
  Issue-License.ps1
  --------------------------------------------------------------
  Izdaje license.lic za JEDNU masinu. Pokreces ti, na svojoj masini,
  kada ti klijent posalje svoj otisak (iz aplikacije: makro ShowMyFingerprint).

  Primer:
    .\Issue-License.ps1 -Fingerprint "A1B2C3..." -Customer "Stanica Novi Sad" `
        -ExpiresAt "2027-06-15" -PrivPem .\license_priv.pem -Out .\license.lic

  Potpisuje se TACNO ovaj payload (UTF-8, bez novog reda):
      fingerprint|customer|issuedAt|expiresAt
  sto modLicense.VerifyRsaSignature na klijentu i verifikuje (SHA256, PKCS1).

  Zahteva OpenSSL na PATH-u.
#>

param(
  [Parameter(Mandatory=$true)][string]$Fingerprint,
  [Parameter(Mandatory=$true)][string]$Customer,
  [Parameter(Mandatory=$true)][string]$ExpiresAt,    # YYYY-MM-DD
  [string]$IssuedAt = (Get-Date -Format "yyyy-MM-dd"),
  [string]$PrivPem = ".\license_priv.pem",
  [string]$Out = ".\license.lic"
)

$ErrorActionPreference = "Stop"
if (-not (Test-Path $PrivPem)) { throw "Privatni kljuc nije nadjen: $PrivPem" }
if (-not (Get-Command openssl -ErrorAction SilentlyContinue)) { throw "OpenSSL nije na PATH-u." }

$payload = "$Fingerprint|$Customer|$IssuedAt|$ExpiresAt"

# Payload kao UTF-8 BEZ BOM i BEZ zavrsnog newline-a (mora da odgovara klijentu).
$tmp = [System.IO.Path]::GetTempFileName()
$sigTmp = [System.IO.Path]::GetTempFileName()
[System.IO.File]::WriteAllBytes($tmp, [System.Text.Encoding]::UTF8.GetBytes($payload))

& openssl dgst -sha256 -sign $PrivPem -out $sigTmp $tmp
$sigB64 = [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($sigTmp))

Remove-Item $tmp, $sigTmp -ErrorAction SilentlyContinue

$lines = @(
  "fingerprint=$Fingerprint",
  "customer=$Customer",
  "issuedAt=$IssuedAt",
  "expiresAt=$ExpiresAt",
  "sig=$sigB64"
)
# UTF-8 bez BOM
[System.IO.File]::WriteAllLines($Out, $lines, (New-Object System.Text.UTF8Encoding($false)))

Write-Host "Izdato: $Out" -ForegroundColor Green
Write-Host "  Posalji klijentu da ga ubaci u ...\Secrets pored radne sveske."
