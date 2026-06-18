<#
  New-LicenseKeypair.ps1
  --------------------------------------------------------------
  POKRENI JEDNOM, na SVOJOJ masini. Pravi DVA odvojena RSA-2048 para:

    1) LICENSE par  -> potpisuje offline license.lic
         license_priv.pem  (TAJNO, NIKAD na server/stanicu; samo Issue-License.ps1)
         license_pub.cer   (JAVNO; ide na svaku stanicu, verifikuje license.lic)

    2) VERDICT par  -> potpisuje online kill-switch odgovore (GAS)
         verdict_priv.pem  (TAJNO; ide SAMO u GAS Script property VERDICT_PRIV_PEM)
         verdict_pub.cer   (JAVNO; ide na svaku stanicu, verifikuje online verdikt)

  Zasto dva para: GAS mora da drzi verdict privatni kljuc. Ako se GAS
  kompromituje, napadac moze da lazira online verdikte, ali NE moze da
  falsifikuje offline licence (license privatni kljuc nikad ne napusta tebe).

  Klijenti (stanice) NE trebaju OpenSSL — verifikuju ugradjenim PowerShell-om.
  Zahteva OpenSSL na PATH-u (https://slproweb.com/products/Win32OpenSSL.html).
#>

param(
  [string]$OutDir = ".",
  [int]$Days = 3650
)

$ErrorActionPreference = "Stop"
if (-not (Get-Command openssl -ErrorAction SilentlyContinue)) {
  throw "OpenSSL nije nadjen na PATH-u. Instaliraj ga pa pokreni ponovo."
}

function New-Pair([string]$name, [string]$cn) {
  $priv   = Join-Path $OutDir "$name`_priv.pem"
  $pubPem = Join-Path $OutDir "$name`_pub.pem"
  $cer    = Join-Path $OutDir "$name`_pub.cer"

  Write-Host "[$name] privatni kljuc -> $priv"
  & openssl genpkey -algorithm RSA -pkeyopt rsa_keygen_bits:2048 -out $priv

  Write-Host "[$name] javni kljuc -> $pubPem"
  & openssl rsa -in $priv -pubout -out $pubPem

  # Self-signed cert u DER (.cer) koji ucitava .NET X509Certificate2 na klijentu.
  Write-Host "[$name] sertifikat (DER) -> $cer"
  & openssl req -new -x509 -key $priv -sha256 -days $Days -subj "/CN=$cn" -outform DER -out $cer
}

New-Pair "license" "AgriX License"
New-Pair "verdict" "AgriX Verdict"

Write-Host ""
Write-Host "GOTOVO." -ForegroundColor Green
Write-Host "  TAJNO (kod tebe):              license_priv.pem"
Write-Host "  TAJNO (samo u GAS property):   verdict_priv.pem"
Write-Host "  JAVNO (u ...\Secrets na svakoj stanici): license_pub.cer  i  verdict_pub.cer"
