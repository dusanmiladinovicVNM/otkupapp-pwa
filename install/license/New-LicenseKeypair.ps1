<#
  New-LicenseKeypair.ps1
  --------------------------------------------------------------
  POKRENI JEDNOM, na SVOJOJ masini. Pravi RSA-2048 keypair za AgriX licence.

  Izlaz:
    license_priv.pem   -> TAJNO. Koristi ga Issue-License.ps1 i GAS kill-switch.
    license_pub.cer    -> JAVNO. Ide na svaku otkupnu stanicu (u ...\Secrets).
    license_pub.pem    -> info/rezerva.

  Zahteva OpenSSL na PATH-u (https://slproweb.com/products/Win32OpenSSL.html).
  Klijenti (otkupne stanice) NE trebaju OpenSSL — oni verifikuju preko
  ugradjenog Windows PowerShell-a i .cer sertifikata.
#>

param(
  [string]$OutDir = ".",
  [int]$Days = 3650
)

$ErrorActionPreference = "Stop"
if (-not (Get-Command openssl -ErrorAction SilentlyContinue)) {
  throw "OpenSSL nije nadjen na PATH-u. Instaliraj ga pa pokreni ponovo."
}

$priv = Join-Path $OutDir "license_priv.pem"
$pubPem = Join-Path $OutDir "license_pub.pem"
$cer  = Join-Path $OutDir "license_pub.cer"

Write-Host "Generisem privatni kljuc -> $priv"
& openssl genpkey -algorithm RSA -pkeyopt rsa_keygen_bits:2048 -out $priv

Write-Host "Izvozim javni kljuc -> $pubPem"
& openssl rsa -in $priv -pubout -out $pubPem

# Self-signed cert u DER formatu (.cer) koji ucitava .NET X509Certificate2 na klijentu.
Write-Host "Pravim self-signed sertifikat (DER) -> $cer"
& openssl req -new -x509 -key $priv -sha256 -days $Days -subj "/CN=AgriX License" -outform DER -out $cer

Write-Host ""
Write-Host "GOTOVO." -ForegroundColor Green
Write-Host "  TAJNO (cuvaj van otkupnih stanica): $priv"
Write-Host "  JAVNO (kopiraj u ...\Secrets na svakoj stanici): $cer"
