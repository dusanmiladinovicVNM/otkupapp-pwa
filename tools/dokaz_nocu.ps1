# tools/dokaz_nocu.ps1
# Nocni pun recert: vba_check + FULL suite + PUN dvosmerni dokaz.
# Ujutru stoji `tests\dokaz_jutro.md` -- jedan ekran, bez kopanja po logovima.
#
# Pun katalog NAMERNO ne cita knjigu (`--knjiga` se ovde ne prosledjuje). Bas to
# je ograda zbog koje je knjiga bezbedna: sve sto je preko dana preskoceno,
# nocu se dokazuje iz pocetka, pa je zastarelost ogranicena na jednu noc.
#
# Registracija:  powershell -File tools\install_nocni_dokaz.ps1
# Rucno:         powershell -File tools\dokaz_nocu.ps1
#
# VAZNO -- Excel i Task Scheduler: `run_vba.py` trazi interaktivnu sesiju
# (compile probe gadja VIDLJIV I AKTIVAN VBE prozor). Zato zadatak radi SAMO kad
# je korisnik prijavljen. Na zakljucanoj masini compile ume da zavrsi kao
# COMPILE NEJASNO -- to NE obara run dok suite-ovi idu (v. run_vba.py), pa
# nocni izvestaj i dalje nosi verdikt, samo bez compile kapije.

$ErrorActionPreference = 'Continue'   # jedan pad ne sme da preskoci izvestaj

$root = (git rev-parse --show-toplevel).Trim()
Set-Location $root

$tests  = Join-Path $root 'tests'
$logDir = Join-Path $tests 'nocni'
New-Item -ItemType Directory -Force -Path $tests, $logDir | Out-Null

$stamp  = Get-Date -Format 'yyyy-MM-dd_HHmm'
$grana  = (git rev-parse --abbrev-ref HEAD).Trim()
$commit = (git rev-parse --short HEAD).Trim()
$json   = Join-Path $tests 'dokaz_last.json'

function Pusti([string]$ime, [string[]]$argumenti) {
    $log = Join-Path $logDir ("{0}_{1}.log" -f $stamp, $ime)
    $t0 = Get-Date
    & python @argumenti *>&1 | Tee-Object -FilePath $log | Out-Null
    $rc = $LASTEXITCODE
    [pscustomobject]@{
        Ime    = $ime
        Rc     = $rc
        Log    = (Split-Path -Leaf $log)
        Minuta = [math]::Round(((Get-Date) - $t0).TotalMinutes, 1)
    }
}

Remove-Item -Force -ErrorAction SilentlyContinue $json

$rez = @()
$rez += Pusti 'vba_check'    @('tools\vba_check.py')
$rez += Pusti 'run_vba_full' @('tools\run_vba.py')
$rez += Pusti 'dokaz_pun'    @('tools\dokaz.py', '--json', $json)

# --- izvestaj za jutro ------------------------------------------------------
$d = $null
if (Test-Path $json) {
    try { $d = Get-Content $json -Raw -Encoding UTF8 | ConvertFrom-Json } catch { $d = $null }
}

$L = New-Object System.Collections.Generic.List[string]
$L.Add("# Nocni recert -- $stamp")
$L.Add("")
$L.Add("grana ``$grana`` @ ``$commit``")
$L.Add("")
$L.Add("| korak | rezultat | minuta | log |")
$L.Add("|---|---|---|---|")
foreach ($r in $rez) {
    $st = if ($r.Rc -eq 0) { 'ZELENO' } else { "PALO (rc=$($r.Rc))" }
    $L.Add("| $($r.Ime) | $st | $($r.Minuta) | ``$($r.Log)`` |")
}
$L.Add("")

if ($null -eq $d) {
    # Alat koji ne zna ishod mora to da kaze glasno, ne da izostavi red.
    $L.Add("## Dokaz: VERDIKT NEDOSTAJE")
    $L.Add("")
    $L.Add("``$json`` ne postoji ili se ne cita -- run je pukao pre prvog upisa.")
    $L.Add("Pogledaj ``tests\nocni\${stamp}_dokaz_pun.log``.")
} else {
    $stanje = if ($d.u_toku) { "$($d.verdikt) -- RUN NIJE ZAVRSEN" } else { $d.verdikt }
    $L.Add("## Dokaz: $stanje")
    $L.Add("")
    $L.Add("sabotaza $($d.sabotaza), obradjeno $($d.obradjeno), crvenih $($d.crvenih), preneseno $($d.preneseno)")
    $L.Add("trajanje $([math]::Round($d.sekundi / 60, 1)) min")
    if ($d.problemi -and $d.problemi.Count -gt 0) {
        $L.Add("")
        $L.Add("### Problemi ($($d.problemi.Count))")
        $L.Add("")
        foreach ($p in $d.problemi) { $L.Add("- ``$($p[0])`` -> $($p[1])") }
    }
    if ($d.poznati -and $d.poznati.Count -gt 0) {
        $L.Add("")
        $L.Add("### Priznati nalazi ($($d.poznati.Count))")
        $L.Add("")
        foreach ($p in $d.poznati) { $L.Add("- ``$($p[0])`` -> $($p[1])") }
    }
}

$L.Add("")
$L.Add("---")
$L.Add("Detalji: ``tests\nocni\`` (logovi ovog run-a nose prefiks ``$stamp``).")

$jutro = Join-Path $tests 'dokaz_jutro.md'
$L -join "`r`n" | Set-Content -Path $jutro -Encoding UTF8
Write-Output "izvestaj: $jutro"

# Logovi stariji od 14 dana se brisu -- nocni run inace zatrpa disk.
Get-ChildItem $logDir -Filter '*.log' -ErrorAction SilentlyContinue |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-14) } |
    Remove-Item -Force -ErrorAction SilentlyContinue

if ($rez | Where-Object { $_.Rc -ne 0 }) { exit 1 } else { exit 0 }
