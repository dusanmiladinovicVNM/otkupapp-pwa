Option Explicit

' ============================================================
' modLicense — desktop licenca / zastita od neovlascene upotrebe
'
' MODEL: OFFLINE potpisana licenca (license.lic) + ONLINE kill-switch.
'   - OFFLINE: payload je potpisan tvojim privatnim kljucem; klijent
'     verifikuje javnim X509 sertifikatom (license_pub.cer). Radi bez mreze.
'   - ONLINE (best-effort): GAS endpoint za revokaciju, sa grace prozorom,
'     da kratak nestanak mreze NE zakljuca legitimnog korisnika.
'
' BEZBEDNO PUSTANJE:
'   Enforcement je iza flag-a u tblSEFConfig:
'       LICENSE_ENFORCE = YES   -> proverava i moze da zatvori svesku
'       (prazno / NO)           -> NO-OP (uvek dozvoli)
'   Tako samo DODAVANJE ovog modula ne moze da zakljuca aplikaciju dok ga
'   svesno ne ukljucis (posle sto si ubacio cert + license.lic).
'
' FAJLOVI (trazi se u APP_SECRETS_PATH, pa pored radne sveske):
'   license.lic        — po masini, izdat tvojim alatom Issue-License.ps1
'   license_pub.cer    — tvoj javni X509 sertifikat (isti za sve masine)
'
' ZAVISNOSTI: Windows PowerShell 5.1 + .NET 4.6+ (stock Win10/11).
'
' UPOZORENJE: VBA je otvoren; ovo je DETERRENT, ne kriptografski zid.
' Determinisani napadac moze da ispegla makro. Svrha: spreciti kopiranje/
' deljenje radne sveske na neovlascene masine + daljinski kill-switch.
'
' STATUS: napisano bez Excel/Windows test-prolaza. Pre produkcije OBAVEZNO
' verifikovati na realnoj masini (videti docs/LICENSING.md).
' ============================================================

' --- Config kljucevi (tblSEFConfig) ---
Private Const LIC_KEY_ENFORCE        As String = "LICENSE_ENFORCE"
Private Const LIC_KEY_LAST_ONLINE_OK As String = "LICENSE_LAST_ONLINE_OK_AT"
Private Const LIC_KEY_ONLINE_URL     As String = "LICENSE_CHECK_URL"
Private Const LIC_KEY_GRACE_DAYS     As String = "LICENSE_OFFLINE_GRACE_DAYS"

' --- Imena fajlova ---
Private Const LIC_FILE        As String = "license.lic"
Private Const LIC_PUBCERT     As String = "license_pub.cer"   ' verifikuje license.lic (offline)
Private Const LIC_VERDICTCERT As String = "verdict_pub.cer"  ' verifikuje online verdikt (kill-switch)

' Fail-mode kad provera NEOCEKIVANO pukne (bug): OPEN (default) ili CLOSED.
Private Const LIC_KEY_FAIL_MODE As String = "LICENSE_FAIL_MODE"

Private Const LIC_DEFAULT_GRACE_DAYS  As Long = 30
Private Const LIC_ONLINE_TIMEOUT_MS   As Long = 3000
Private Const LIC_ONLINE_MAX_AGE_DAYS As Long = 2   ' anti-replay: max starost potpisanog verdikta

' ============================================================
' PUBLIC ENTRY — pozvati na vrhu modMain.StartApp
' ============================================================
Public Function LicenseGateOrQuit() As Boolean
    On Error GoTo EH

    If Not LicenseEnforced() Then
        LicenseGateOrQuit = True
        Exit Function
    End If

    Dim fp As String
    fp = GetMachineFingerprint()

    ' #10: bez PowerShell-a otisak je degradiran (RAW-) i ne moze se pouzdano
    ' uporediti sa licencom -> jasna poruka umesto zbunjujuceg "druga masina".
    If Left$(fp, 4) = "RAW-" Then
        DenyAndClose "Verifikacija licence zahteva Windows PowerShell, koji nije " & _
                     "dostupan/dozvoljen na ovoj masini." & vbCrLf & _
                     "Kontaktirajte dobavljaca." & vbCrLf & vbCrLf & _
                     "Otisak (degradiran):" & vbCrLf & fp
        LicenseGateOrQuit = False
        Exit Function
    End If

    ' 1) OFFLINE provera potpisane licence
    Dim offReason As String
    offReason = ValidateOfflineLicense(fp)      ' "" = OK
    If Len(offReason) > 0 Then
        DenyAndClose "Licenca nije validna za ovu masinu." & vbCrLf & _
                     offReason & vbCrLf & vbCrLf & _
                     "Otisak masine (posaljite dobavljacu):" & vbCrLf & fp
        LicenseGateOrQuit = False
        Exit Function
    End If

    ' 2) ONLINE kill-switch (best-effort)
    Dim onResult As String
    onResult = CheckOnlineKillSwitch(fp)        ' "OK" / "REVOKED" / "UNREACHABLE"

    Select Case onResult
        Case "REVOKED"
            DenyAndClose "Licenca je opozvana." & vbCrLf & vbCrLf & _
                         "Otisak masine:" & vbCrLf & fp
            LicenseGateOrQuit = False
        Case "OK"
            On Error Resume Next
            SetConfigValue LIC_KEY_LAST_ONLINE_OK, Format$(Now, "yyyy-mm-dd hh:nn:ss")
            On Error GoTo EH
            LicenseGateOrQuit = True
        Case Else        ' UNREACHABLE -> grace prozor
            If WithinOfflineGrace() Then
                LicenseGateOrQuit = True
            Else
                DenyAndClose "Licenca zahteva periodicnu online proveru, a grace " & _
                             "prozor je istekao." & vbCrLf & _
                             "Povezite se na internet i pokrenite ponovo." & vbCrLf & vbCrLf & _
                             "Otisak masine:" & vbCrLf & fp
                LicenseGateOrQuit = False
            End If
    End Select

    Exit Function

EH:
    ' Neocekivana greska u SAMOJ proveri (bug). Podrazumevano fail-OPEN da bag
    ' u licenciranju ne zakljuca legitimne korisnike. Posle Windows verifikacije
    ' mozes da postavis LICENSE_FAIL_MODE=CLOSED za stroziji rezim.
    LogErr "modLicense.LicenseGateOrQuit"
    If FailClosedConfigured() Then
        DenyAndClose "Provera licence nije uspela (interna greska)." & vbCrLf & _
                     "Kontaktirajte dobavljaca."
        LicenseGateOrQuit = False
    Else
        LicenseGateOrQuit = True
    End If
End Function

' ============================================================
' PUBLIC — operater procita svoj otisak da ga posalje dobavljacu
' ============================================================
Public Sub ShowMyFingerprint()
    Dim fp As String
    fp = GetMachineFingerprint()
    InputBox "Otisak ove masine (kopirajte i posaljite dobavljacu " & _
             "radi izdavanja licence):", APP_NAME, fp
End Sub

' ============================================================
' ENFORCEMENT FLAG
' ============================================================
Private Function LicenseEnforced() As Boolean
    Dim v As String
    v = UCase$(Trim$(GetConfigValue(LIC_KEY_ENFORCE)))
    Select Case v
        Case "YES", "TRUE", "1", "DA", "ON", "ENABLED"
            LicenseEnforced = True
        Case Else
            LicenseEnforced = False
    End Select
End Function

' Fail-closed samo ako je eksplicitno trazeno; inace fail-open (default).
Private Function FailClosedConfigured() As Boolean
    On Error Resume Next
    Dim v As String
    v = UCase$(Trim$(GetConfigValue(LIC_KEY_FAIL_MODE)))
    FailClosedConfigured = (v = "CLOSED" Or v = "ZATVORENO" Or v = "DENY")
End Function

' ============================================================
' MACHINE FINGERPRINT
' Stabilne vrednosti: serijski broj C: + Windows MachineGuid + ime PC-a.
' (Izbegavamo MAC i samo ime PC-a kao jedini kljuc — menjaju se lako.)
' ============================================================
Public Function GetMachineFingerprint() As String
    On Error Resume Next
    Dim raw As String
    raw = "agrix-fp-v1|" & VolSerialC() & "|" & MachineGuid() & "|" & _
          UCase$(Trim$(Environ$("COMPUTERNAME")))

    Dim h As String
    h = Sha256HexOfString(raw)

    If Len(h) = 0 Then
        ' Fallback ako PowerShell SHA nije dostupan — degradirano, ali stabilno.
        GetMachineFingerprint = "RAW-" & UCase$(Hex$(SimpleHash(raw)))
    Else
        GetMachineFingerprint = UCase$(Left$(h, 40))   ' kratak, citljiv otisak
    End If
End Function

Private Function VolSerialC() As String
    On Error Resume Next
    VolSerialC = CStr(CreateObject("Scripting.FileSystemObject").GetDrive("C:").SerialNumber)
End Function

Private Function MachineGuid() As String
    On Error Resume Next
    MachineGuid = CreateObject("WScript.Shell"). _
        RegRead("HKLM\SOFTWARE\Microsoft\Cryptography\MachineGuid")
End Function

' ============================================================
' OFFLINE LICENSE VALIDATION
' license.lic je INI-stil:
'   fingerprint=...
'   customer=...
'   issuedAt=YYYY-MM-DD
'   expiresAt=YYYY-MM-DD
'   sig=<base64 RSA potpis preko "fingerprint|customer|issuedAt|expiresAt">
' ============================================================
Private Function ValidateOfflineLicense(ByVal fp As String) As String
    On Error GoTo EH

    Dim licPath As String
    licPath = ResolveSecretFile(LIC_FILE)
    If Len(licPath) = 0 Then
        ValidateOfflineLicense = "Nedostaje " & LIC_FILE & " na ovoj masini."
        Exit Function
    End If

    Dim certPath As String
    certPath = ResolveSecretFile(LIC_PUBCERT)
    If Len(certPath) = 0 Then
        ValidateOfflineLicense = "Nedostaje javni sertifikat " & LIC_PUBCERT & "."
        Exit Function
    End If

    Dim content As String
    content = ReadTextFile(licPath)

    Dim licFp As String, licCust As String, licIss As String, licExp As String, licSig As String
    licFp = ParseLicValue(content, "fingerprint")
    licCust = ParseLicValue(content, "customer")
    licIss = ParseLicValue(content, "issuedAt")
    licExp = ParseLicValue(content, "expiresAt")
    licSig = ParseLicValue(content, "sig")

    If Len(licFp) = 0 Or Len(licSig) = 0 Then
        ValidateOfflineLicense = "Licenca je nepotpuna ili ostecena."
        Exit Function
    End If

    ' a) vezana za ovu masinu?
    If StrComp(Trim$(licFp), Trim$(fp), vbTextCompare) <> 0 Then
        ValidateOfflineLicense = "Licenca je izdata za drugu masinu."
        Exit Function
    End If

    ' b) istekla? (ISO yyyy-mm-dd, locale-nezavisno parsiranje)
    If Len(licExp) > 0 Then
        Dim expDate As Date
        If ParseIsoDate(licExp, expDate) Then
            If expDate < Date Then
                ValidateOfflineLicense = "Licenca je istekla (" & licExp & ")."
                Exit Function
            End If
        End If
    End If

    ' c) digitalni potpis validan?
    ' Customer NIJE u potpisu: potpisani payload je cisto ASCII
    ' (fingerprint|issuedAt|expiresAt) -> nema encoding neslaganja izmedju
    ' potpisivanja (UTF-8) i citanja kad naziv kupca ima srpska slova.
    Dim payload As String
    payload = licFp & "|" & licIss & "|" & licExp

    If Not VerifyRsaSignature(payload, licSig, certPath) Then
        ValidateOfflineLicense = "Digitalni potpis licence nije validan."
        Exit Function
    End If

    ValidateOfflineLicense = ""     ' OK
    Exit Function

EH:
    ValidateOfflineLicense = "Greska pri citanju licence: " & Err.description
End Function

' ============================================================
' ONLINE KILL-SWITCH (best-effort)
' Salje fingerprint GAS endpointu; ocekuje potpisan verdikt.
' Odgovor (JSON):  {"status":"ok|revoked","payload":"fp|status|YYYY-MM-DD","sig":"base64"}
' Bilo kakav problem (nema URL-a, timeout, nevalidan potpis) -> "UNREACHABLE".
' ============================================================
Private Function CheckOnlineKillSwitch(ByVal fp As String) As String
    On Error GoTo EH

    Dim url As String
    url = Trim$(GetConfigValue(LIC_KEY_ONLINE_URL))
    If Len(url) = 0 Then
        CheckOnlineKillSwitch = "UNREACHABLE"   ' online provera nije konfigurisana -> cisto offline
        Exit Function
    End If
    If LCase$(Left$(url, 8)) <> "https://" Then
        CheckOnlineKillSwitch = "UNREACHABLE"   ' samo HTTPS
        Exit Function
    End If

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts LIC_ONLINE_TIMEOUT_MS, LIC_ONLINE_TIMEOUT_MS, LIC_ONLINE_TIMEOUT_MS, LIC_ONLINE_TIMEOUT_MS
    http.Open "POST", url, False
    http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.Send "{""fingerprint"":""" & JsonEscapeLic(fp) & """}"

    If http.status < 200 Or http.status >= 300 Then
        CheckOnlineKillSwitch = "UNREACHABLE"
        Exit Function
    End If

    Dim resp As String
    resp = CStr(http.responseText)

    Dim status As String, payload As String, sig As String
    status = JsonValueLic(resp, "status")
    payload = JsonValueLic(resp, "payload")
    sig = JsonValueLic(resp, "sig")

    ' #5: verdikt potpisuje ODVOJEN verdict kljuc (ne license kljuc), da
    ' kompromitacija GAS-a (koji drzi verdict priv. kljuc) ne omoguci
    ' falsifikovanje offline licenci.
    Dim certPath As String
    certPath = ResolveSecretFile(LIC_VERDICTCERT)
    If Len(certPath) = 0 Or Len(sig) = 0 Or Len(payload) = 0 Then
        CheckOnlineKillSwitch = "UNREACHABLE"
        Exit Function
    End If
    If Not VerifyRsaSignature(payload, sig, certPath) Then
        CheckOnlineKillSwitch = "UNREACHABLE"   ' nepotpisan/lazan odgovor — ignorisi
        Exit Function
    End If

    ' payload mora da pocinje ovim fingerprint-om (anti-replay na drugu masinu)
    If InStr(1, payload, fp, vbTextCompare) <> 1 Then
        CheckOnlineKillSwitch = "UNREACHABLE"
        Exit Function
    End If

    ' #7 freshness: odbij stare/buduce potpisane verdikte (neko kesira validan
    ' "ok" i replay-uje ga da izbegne revokaciju).
    If Not VerdictDateFresh(payload) Then
        CheckOnlineKillSwitch = "UNREACHABLE"
        Exit Function
    End If

    If LCase$(status) = "revoked" Then
        CheckOnlineKillSwitch = "REVOKED"
    Else
        CheckOnlineKillSwitch = "OK"
    End If
    Exit Function

EH:
    CheckOnlineKillSwitch = "UNREACHABLE"
End Function

' payload = "fingerprint|status|YYYY-MM-DD". Svez ako datum nije stariji od
' LIC_ONLINE_MAX_AGE_DAYS i nije iz buducnosti (> 1 dan, dozvola za clock-skew).
Private Function VerdictDateFresh(ByVal payload As String) As Boolean
    On Error GoTo EH
    Dim parts() As String
    parts = Split(payload, "|")
    If UBound(parts) < 2 Then Exit Function
    Dim d As Date
    If Not ParseIsoDate(parts(2), d) Then Exit Function
    If DateDiff("d", d, Date) > LIC_ONLINE_MAX_AGE_DAYS Then Exit Function
    If DateDiff("d", Date, d) > 1 Then Exit Function
    VerdictDateFresh = True
    Exit Function
EH:
    VerdictDateFresh = False
End Function

Private Function WithinOfflineGrace() As Boolean
    On Error GoTo EH

    ' Ako online provera uopste nije konfigurisana -> cisto offline rezim,
    ' grace je uvek "u okviru" (oslanjamo se na offline potpis + istek).
    If Len(Trim$(GetConfigValue(LIC_KEY_ONLINE_URL))) = 0 Then
        WithinOfflineGrace = True
        Exit Function
    End If

    Dim graceDays As Long
    graceDays = LIC_DEFAULT_GRACE_DAYS
    If IsNumeric(GetConfigValue(LIC_KEY_GRACE_DAYS)) Then
        graceDays = CLng(GetConfigValue(LIC_KEY_GRACE_DAYS))
    End If

    Dim last As String
    last = Trim$(GetConfigValue(LIC_KEY_LAST_ONLINE_OK))

    If Len(last) = 0 Then
        ' Prvi put bez uspesne online provere — seed-uj sada i dozvoli.
        ' Korisnik dobija graceDays da se jednom poveze.
        On Error Resume Next
        SetConfigValue LIC_KEY_LAST_ONLINE_OK, Format$(Now, "yyyy-mm-dd hh:nn:ss")
        On Error GoTo EH
        WithinOfflineGrace = True
        Exit Function
    End If

    If IsDate(last) Then
        WithinOfflineGrace = (DateDiff("d", CDate(last), Now) <= graceDays)
    Else
        WithinOfflineGrace = True
    End If
    Exit Function

EH:
    WithinOfflineGrace = True
End Function

' ============================================================
' RSA VERIFY preko PowerShell-a (X509 cert, stock Win 5.1 / .NET 4.6+)
' Payload se pise kao UTF-8 bez BOM da bajtovi odgovaraju potpisivanju.
' ============================================================
Private Function VerifyRsaSignature(ByVal payload As String, _
                                    ByVal sigBase64 As String, _
                                    ByVal certPath As String) As Boolean
    On Error GoTo EH

    Dim base As String
    base = TempBase("agrixsig")

    Dim payloadFile As String, sigFile As String
    payloadFile = base & ".bin"
    sigFile = base & ".sig"

    WriteUtf8NoBom payloadFile, payload
    WriteTextFile sigFile, sigBase64

    Dim ps As String
    ps = "try {" & vbCrLf & _
         "  $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2('" & PsLit(certPath) & "')" & vbCrLf & _
         "  $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPublicKey($cert)" & vbCrLf & _
         "  $data = [IO.File]::ReadAllBytes('" & PsLit(payloadFile) & "')" & vbCrLf & _
         "  $sig = [Convert]::FromBase64String((Get-Content '" & PsLit(sigFile) & "' -Raw).Trim())" & vbCrLf & _
         "  if ($rsa.VerifyData($data,$sig,[System.Security.Cryptography.HashAlgorithmName]::SHA256,[System.Security.Cryptography.RSASignaturePadding]::Pkcs1)) {$__r='VALID'} else {$__r='INVALID'}" & vbCrLf & _
         "} catch { $__r = 'ERROR' }"

    Dim r As String
    r = RunPowerShellScript(ps)

    On Error Resume Next
    Kill payloadFile
    Kill sigFile
    On Error GoTo EH

    VerifyRsaSignature = (UCase$(Trim$(r)) = "VALID")
    Exit Function

EH:
    VerifyRsaSignature = False
End Function

Private Function Sha256HexOfString(ByVal s As String) As String
    On Error GoTo EH

    Dim base As String, rawFile As String
    base = TempBase("agrixfp")
    rawFile = base & ".bin"

    WriteUtf8NoBom rawFile, s

    Dim ps As String
    ps = "try {" & vbCrLf & _
         "  $b = [IO.File]::ReadAllBytes('" & PsLit(rawFile) & "')" & vbCrLf & _
         "  $h = [System.Security.Cryptography.SHA256]::Create().ComputeHash($b)" & vbCrLf & _
         "  $__r = ($h | ForEach-Object { $_.ToString('x2') }) -join ''" & vbCrLf & _
         "} catch { $__r = '' }"

    Dim r As String
    r = RunPowerShellScript(ps)

    On Error Resume Next
    Kill rawFile
    On Error GoTo EH

    Sha256HexOfString = Trim$(r)
    Exit Function

EH:
    Sha256HexOfString = ""
End Function

' Pokrece PS skriptu koja postavi $__r; vraca sadrzaj $__r kao string.
Private Function RunPowerShellScript(ByVal psBody As String) As String
    On Error GoTo EH

    Dim base As String, ps1 As String, outFile As String
    base = TempBase("agrixps")
    ps1 = base & ".ps1"
    outFile = base & ".out"

    ' .ps1 kao UTF-8 sa BOM -> Windows PowerShell 5.1 ga cita kao UTF-8, pa
    ' ne-ASCII putanje (npr. C:\Korisnici\Zarko\...) ostaju ispravne.
    WriteUtf8WithBom ps1, psBody & vbCrLf & _
        "$__r | Out-File -Encoding ascii -FilePath '" & PsLit(outFile) & "'"

    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    sh.Run "powershell -NoProfile -ExecutionPolicy Bypass -File """ & ps1 & """", 0, True

    RunPowerShellScript = ReadTextFile(outFile)

    On Error Resume Next
    Kill ps1
    Kill outFile
    Exit Function

EH:
    RunPowerShellScript = ""
End Function

' ============================================================
' FILE / STRING HELPERS
' ============================================================
Private Function ResolveSecretFile(ByVal fileName As String) As String
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 1) ...\Secrets pored radne sveske (modSetup ovde pravi Secrets folder)
    Dim secrets As String
    secrets = ThisWorkbook.Path & "\Secrets"
    If fso.FileExists(secrets & "\" & fileName) Then
        ResolveSecretFile = secrets & "\" & fileName
        Exit Function
    End If

    ' 2) pored radne sveske
    If fso.FileExists(ThisWorkbook.Path & "\" & fileName) Then
        ResolveSecretFile = ThisWorkbook.Path & "\" & fileName
        Exit Function
    End If

    ResolveSecretFile = ""
End Function

Private Function TempBase(ByVal prefix As String) As String
    Randomize
    TempBase = Environ$("TEMP") & "\" & prefix & "_" & _
               Format$(Now, "yyyymmddhhnnss") & "_" & CStr(Int(Rnd * 1000000))
End Function

Private Sub WriteTextFile(ByVal path As String, ByVal content As String)
    On Error Resume Next
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(path, True, False)   ' ASCII
    ts.Write content
    ts.Close
End Sub

Private Function ReadTextFile(ByVal path As String) As String
    On Error Resume Next
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then Exit Function
    Set ts = fso.OpenTextFile(path, 1, False)         ' 1 = ForReading
    If Not ts.AtEndOfStream Then ReadTextFile = ts.ReadAll
    ts.Close
End Function

Private Sub WriteUtf8NoBom(ByVal path As String, ByVal s As String)
    On Error Resume Next
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2                  ' text
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText s

    ' Skini BOM tako sto procitamo binarno od pozicije 3.
    stm.Position = 0
    stm.Type = 1                  ' binary
    stm.Position = 3
    Dim bin
    bin = stm.Read
    stm.Close

    Dim out As Object
    Set out = CreateObject("ADODB.Stream")
    out.Type = 1
    out.Open
    out.Write bin
    out.SaveToFile path, 2        ' 2 = overwrite
    out.Close
End Sub

' UTF-8 SA BOM (za .ps1 koji Windows PowerShell 5.1 treba da procita kao UTF-8).
Private Sub WriteUtf8WithBom(ByVal path As String, ByVal s As String)
    On Error Resume Next
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2                  ' text
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText s
    stm.SaveToFile path, 2        ' 2 = overwrite (ADODB upisuje BOM za utf-8 text)
    stm.Close
End Sub

' Sadrzaj za PowerShell single-quoted string: duplira ' (bezbedne putanje).
Private Function PsLit(ByVal s As String) As String
    PsLit = Replace(s, "'", "''")
End Function

Private Function ParseLicValue(ByVal content As String, ByVal key As String) As String
    On Error Resume Next
    Dim lines() As String, i As Long, ln As String, p As Long
    lines = Split(Replace(content, vbCr, ""), vbLf)
    For i = LBound(lines) To UBound(lines)
        ln = Trim$(lines(i))
        p = InStr(ln, "=")
        If p > 1 Then
            If StrComp(Trim$(Left$(ln, p - 1)), key, vbTextCompare) = 0 Then
                ParseLicValue = Trim$(Mid$(ln, p + 1))
                Exit Function
            End If
        End If
    Next i
End Function

' Minimalni JSON string-vrednost izvlakac za potpuno kontrolisan odgovor.
Private Function JsonValueLic(ByVal json As String, ByVal key As String) As String
    On Error Resume Next
    Dim needle As String, p As Long, q As Long, startQ As Long
    needle = """" & key & """"
    p = InStr(1, json, needle, vbTextCompare)
    If p = 0 Then Exit Function
    p = InStr(p + Len(needle), json, ":")
    If p = 0 Then Exit Function
    startQ = InStr(p, json, """")
    If startQ = 0 Then Exit Function
    q = InStr(startQ + 1, json, """")
    If q = 0 Then Exit Function
    JsonValueLic = Mid$(json, startQ + 1, q - startQ - 1)
End Function

Private Function JsonEscapeLic(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    JsonEscapeLic = s
End Function

' Nekripto fallback hash (samo za degradirani fingerprint bez PowerShell-a).
Private Function SimpleHash(ByVal s As String) As Long
    Dim i As Long, h As Long
    h = 5381
    For i = 1 To Len(s)
        h = ((h * 33) Xor Asc(Mid$(s, i, 1))) And &H7FFFFFFF
    Next i
    SimpleHash = h
End Function

' ISO datum (yyyy-mm-dd) -> Date, locale-nezavisno (eksplicitni DateSerial).
' Vraca False ako format nije ispravan (tada se istek ne proverava ovde, ali
' expiresAt je u potpisanom payload-u pa ga falsifikovanje ionako obara).
Private Function ParseIsoDate(ByVal s As String, ByRef outDate As Date) As Boolean
    On Error GoTo EH
    Dim p() As String
    p = Split(Trim$(s), "-")
    If UBound(p) <> 2 Then Exit Function
    If Not (IsNumeric(p(0)) And IsNumeric(p(1)) And IsNumeric(p(2))) Then Exit Function
    outDate = DateSerial(CLng(p(0)), CLng(p(1)), CLng(p(2)))
    ParseIsoDate = True
    Exit Function
EH:
    ParseIsoDate = False
End Function

' ============================================================
' DENY
' ============================================================
Private Sub DenyAndClose(ByVal msg As String)
    On Error Resume Next
    Application.Visible = True
    MsgBox msg, vbCritical, APP_NAME

    ' Zatvori SAMO ovu svesku (bez Application.Quit) da ne pozatvaramo tudje
    ' nesnimljene sveske. Saved=True spreci "Sacuvaj?" prompt.
    ThisWorkbook.Saved = True
    Application.DisplayAlerts = False
    ThisWorkbook.Close SaveChanges:=False
End Sub
