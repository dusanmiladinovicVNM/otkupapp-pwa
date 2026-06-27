Attribute VB_Name = "modLicense"
'Attribute VB_Name = "modLicense"
Option Explicit

' ============================================================
' modLicense — per-uredjaj (node-locked) licenciranje.
'
' Model: licenca se prodaje PO RACUNARU. Vezivanje (bind) ZIVI NA SERVERU
' (GAS, akcija "checkLicense"), NE u ovom fajlu. Klijent samo izracuna
' "otisak" masine (MachineGuid + SMBIOS UUID + volume serial) i posalje ga;
' SERVER odlucuje da li taj otisak sme. Prvi racun koji aktivira kljuc se
' veze ("prva aktivacija veze"); svaki drugi racun sa istim kljucem dobija
' BOUND_OTHER i blokira se. Kopiranje .xlsm na drugi racun ne pomaze: drugi
' racunar daje drugaciji otisak -> server ne izdaje token.
'
' VAZNO — granica zastite: ovo zaustavlja casual deljenje ("kolega, posalji
' mi fajl"), sto je ~99% realnog rizika. Ko otvori VBE moze da izbaci poziv
' i radi offline — to je univerzalni plafon svakog VBA-locka. Prava tvrda
' zastita = kriticni podaci/obracun zive samo na serveru.
'
' OPT-IN + LATCH: gate radi ako je u tblSEFConfig LICENSE_ENABLED = YES, ILI
' ako je masina vec jednom USPESNO aktivirana (postoje LICENSE_KEY i
' LICENSE_BOUND_PARTS). Tako merge ovog koda NE blokira postojece instalacije
' dok ne provizionises kljuceve, a posle prve aktivacije vracanje
' LICENSE_ENABLED=NO vise NE gasi proveru na toj masini (anti-bypass).
' Pravi ON/OFF ostaje na serveru (Licenses sheet u Stammdaten:
' adminSuspendLicense / adminActivateLicense / adminResetLicenseBinding).
'
' Config kljucevi (tblSEFConfig):
'   LICENSE_ENABLED      YES/NO  (default NO — provera iskljucena DOK masina
'                        nije aktivirana; posle prve aktivacije latch tera
'                        proveru i bez YES — vidi LicenseRequired)
'   LICENSE_ENDPOINT     GAS Web App /exec URL (ako prazno -> MONITORING_ENDPOINT)
'   LICENSE_KEY          licencni kljuc dodeljen ovom racunaru
'   LICENSE_TOKEN        (interno) potpisan token sa servera
'   LICENSE_BOUND_PARTS  (interno) komponente otiska pri zadnjoj uspesnoj proveri
'   LICENSE_NEXT_CHECK   (interno) ISO; pre ovog datuma dozvoljen offline start
'   LICENSE_STATUS       (interno) OK
'
' Poziva se na vrhu modMain.StartApp:
'   If Not LicenseGateOrQuit() Then Exit Sub
'
' Prva aktivacija na racunaru (jednokratno): Alt+F8 -> ActivateLicensePrompt
' Dijagnostika otiska:                       Alt+F8 -> LicenseShowDevice
' ============================================================

Private Const CFG_LIC_ENABLED     As String = "LICENSE_ENABLED"
Private Const CFG_LIC_ENDPOINT    As String = "LICENSE_ENDPOINT"
Private Const CFG_LIC_KEY         As String = "LICENSE_KEY"
Private Const CFG_LIC_TOKEN       As String = "LICENSE_TOKEN"
Private Const CFG_LIC_BOUND_PARTS As String = "LICENSE_BOUND_PARTS"
Private Const CFG_LIC_NEXT_CHECK  As String = "LICENSE_NEXT_CHECK"
Private Const CFG_LIC_STATUS      As String = "LICENSE_STATUS"
Private Const CFG_LIC_HWM          As String = "LICENSE_HWM"   ' anti-rollback high-water-mark
Private Const CFG_MON_ENDPOINT    As String = "MONITORING_ENDPOINT"

' Pinovan license endpoint (build-time). Ako je postavljen, on je AUTORITET i
' IGNORISE config override (LICENSE_ENDPOINT/MONITORING_ENDPOINT) za licencni
' poziv -> zatvara "repoint na lazni server koji vraca OK" bypass. Prazno =
' nije pinovan (zadrzi config ponasanje, ne lomi postojece instalacije).
' Da aktiviras tvrdo vezivanje: upisi svoj GAS /exec URL i re-sign-uj projekat.
Private Const LIC_ENDPOINT_PINNED As String = ""

' Koliko od 3 komponente otiska mora da se poklopi (fuzzy match) da bi se
' tolerisala manja promena hardvera (nov disk, reinstall) bez lockout-a.
Private Const LIC_MIN_MATCH As Long = 2

' Fallback offline grace (dana) ako server ne posalje graceDays.
' Kraci grace = manja vrednost odlaganja re-provere (suspend se primeni brze).
Private Const LIC_DEFAULT_GRACE_DAYS As Long = 3

Private Const HTTP_TIMEOUT_MS As Long = 8000
Private Const HTTP_RECV_TIMEOUT_MS As Long = 15000

' Module-level state: postavljeno na True kad gate (licenca ILI trial) odbije
' pristup, da Workbook_Open rano prekine startup (cleanup/monitoring se ne
' izvrsavaju za odbijenu masinu). Vidi AccessWasDenied / ForceCloseDeniedWorkbook.
Private gAccessDenied As Boolean

' ============================================================
' PUBLIC — Glavni gate (poziva se iz modMain.StartApp)
' ============================================================
'
' Kombinuje licencu i trial po pravilu "TRIAL SAMO AKO NIJE LICENCIRAN":
'   - masina IMA licencni kljuc        -> licenca odlucuje (OK propusta, svako
'                                          odbijanje BOUND_OTHER/SUSPENDED/... blokira);
'   - nema kljuca, trial UKLJUCEN       -> trial gate (istek blokira);
'   - nema kljuca, trial off, licenca on-> trazi licencu (blokira "unesite kljuc");
'   - nista nije ukljuceno              -> propusta.
'
' Licencirana masina NIKAD ne vidi trial.
Public Function AccessGateOrQuit() As Boolean
    Const SRC As String = "modLicense.AccessGateOrQuit"
    On Error GoTo EH

    Dim licOn As Boolean: licOn = LicenseRequired()
    Dim hasKey As Boolean: hasKey = (Len(Trim$(GetConfigValue(CFG_LIC_KEY))) > 0)

    If licOn And hasKey Then
        ' Masina ima licencu -> licenca je autoritet.
        AccessGateOrQuit = LicenseGateOrQuit()

    ElseIf modTrial.TrialEnabled() Then
        If modTrial.TrialActive() Then
            ' Nema kljuca, trial jos traje -> trial gate (propusta + HWM update).
            AccessGateOrQuit = modTrial.TrialGateOrQuit()
        ElseIf licOn Then
            ' Trial istekao, ali licenca je put napred -> ponudi kljuc INLINE
            ' (resava N9: inace se masina zatvori pre nego sto stignes
            ' Alt+F8 -> ActivateLicensePrompt).
            AccessGateOrQuit = PromptLicenseOnTrialExpiry()
        Else
            ' Trial-only (licenca off) i istekao -> standardni trial blok.
            AccessGateOrQuit = modTrial.TrialGateOrQuit()
        End If

    ElseIf licOn Then
        ' Nema kljuca, trial off -> trazi licencu (blokira "unesite kljuc").
        AccessGateOrQuit = LicenseGateOrQuit()

    Else
        AccessGateOrQuit = True                         ' ni licenca ni trial nisu ukljuceni
    End If
    Exit Function

EH:
    ' Bug u orkestraciji ne sme da zakljuca korisnika -> fail-OPEN + log.
    LogErr SRC
    AccessGateOrQuit = True
End Function

' ============================================================
' PUBLIC — Licencni gate (interni; zove ga AccessGateOrQuit)
' ============================================================

' Vraca True ako sme da nastavi; inace blokira, prikaze poruku i zatvori svesku.
Public Function LicenseGateOrQuit() As Boolean
    Const SRC As String = "modLicense.LicenseGateOrQuit"
    On Error GoTo EH

    ' Opt-in + latch: ako provera nije obavezna (flag NO i masina nije vec
    ' aktivirana), ne diramo nista.
    If Not LicenseRequired() Then
        LicenseGateOrQuit = True
        Exit Function
    End If

    ' SVESNA ODLUKA: NE konsultujemo IsCloudSyncEnabled() ovde.
    ' Licenca je AUTH provera, ne data-sync. Kada bi gate zavisio od
    ' CLOUD_SYNC_ENABLED, korisnik bi licencu iskljucio prostim gasenjem te
    ' opcije (CLOUD_SYNC_ENABLED=NO) -> bypass. Zato je licencni HTTP poziv
    ' NAMERNI izuzetak od desktop-only "100% lokalno" pravila i ukljucuje se
    ' iskljucivo preko LICENSE_ENABLED. (Vidi modConfig.IsCloudSyncEnabled.)
    Dim endpoint As String: endpoint = LicenseEndpoint()
    If Len(endpoint) = 0 Then
        ' Ukljuceno ali nema endpoint-a = misconfig. Fail-OPEN da ne brick-ujemo
        ' korisnika; operater mora da podesi LICENSE_ENDPOINT.
        LogWarn SRC, "LICENSE_ENABLED=YES ali endpoint nije podesen. Preskacem proveru (fail-open)."
        LicenseGateOrQuit = True
        Exit Function
    End If

    Dim key As String: key = Trim$(GetConfigValue(CFG_LIC_KEY))
    If Len(key) = 0 Then
        LicenseBlock "Licencni kljuc nije unet na ovom racunaru.", _
                     "Pokrenite makro: Alt+F8 -> ActivateLicensePrompt"
        LicenseGateOrQuit = False
        Exit Function
    End If

    Dim parts As String: parts = GetDeviceParts()
    If LicNonEmptyParts(parts) < LIC_MIN_MATCH Then
        ' Otisak preslab (npr. WMI nedostupan). Ne kaznjavamo korisnika.
        LogWarn SRC, "Nedovoljno komponenti otiska uredjaja. Preskacem proveru (fail-open)."
        LicenseGateOrQuit = True
        Exit Function
    End If

    Dim bound As String: bound = Trim$(GetConfigValue(CFG_LIC_BOUND_PARTS))
    ' NEXT_CHECK je upisan kao ISO sa 'T' (Format "...\Thh:nn:ss"). VBA CDate/
    ' IsDate NE razumeju 'T' separator (IsDate vrati False), pa ga zamenjujemo
    ' razmakom pre parse-a — isti obrazac kao modGoogleAuth.IsTokenExpired i
    ' modStanicaLock. Bez ovoga brzi offline put se nikad ne bi izvrsio.
    Dim nextChk As String: nextChk = Replace(Trim$(GetConfigValue(CFG_LIC_NEXT_CHECK)), "T", " ")

    ' Anti-rollback (N1): ako je sistemski sat vracen ISPOD ranije vidjenog
    ' datuma, NE veruj offline grace-u (forsiraj online). HWM tehnika kao u
    ' modTrial. Deterrent sloj (HWM je editabilan); pravi autoritet je server.
    Dim today As Date: today = Date
    Dim rolledBack As Boolean: rolledBack = LicenseClockRolledBack(today)
    Call LicenseAdvanceHwm(today)

    ' --- BRZI OFFLINE PUT: vezana masina + unutar grace + sat NIJE vracen ---
    If Not rolledBack And LicenseIsBoundMachine(bound, parts) Then
        If Len(nextChk) > 0 Then
            If IsDate(nextChk) Then
                If Now < CDate(nextChk) Then
                    LicenseGateOrQuit = True
                    Exit Function
                End If
            End If
        End If
    End If

    ' --- ONLINE PROVERA (istekao grace ili nema kesa) ---
    Dim resp As String
    Dim httpOk As Boolean
    LicenseStatusBar "Proveravam licencu..."
    httpOk = LicenseHttpCheck(endpoint, key, parts, resp)
    LicenseStatusBar ""                      ' reset pre eventualne MsgBox blokade
    If Not httpOk Then
        ' Server nedostupan. Ako je OVO vec vezana masina -> offline grace,
        ' da privremeni nestanak interneta ne zakljuca platisu. Ako masina
        ' NIJE vezana (nema validnog kesa) -> blokiraj (aktivacija mora online).
        If LicenseIsBoundMachine(bound, parts) Then
            LogWarn SRC, Poruka("LIC_MSG_SERVER_NEDOSTUPAN_OFFLINE")
            LicenseGateOrQuit = True
            Exit Function
        End If
        LicenseBlock "Aktivacija licence zahteva internet konekciju.", _
                     "Povezite se na internet i pokrenite ponovo."
        LicenseGateOrQuit = False
        Exit Function
    End If

    Dim status As String: status = UCase$(ExtractJsonStringGoogle(resp, "status"))
    Select Case status
        Case "OK"
            Call PersistLicenseOk(parts, resp)
            LicenseGateOrQuit = True

        Case "BOUND_OTHER"
            LicenseBlock "Licenca je vec aktivirana na drugom racunaru.", _
                         "Za prenos na ovaj racunar kontaktirajte dobavljaca."
            LicenseGateOrQuit = False

        Case "SUSPENDED"
            LicenseBlock "Licenca je suspendovana.", LicenseErr(resp)
            LicenseGateOrQuit = False

        Case "EXPIRED"
            LicenseBlock "Licenca je istekla.", LicenseErr(resp)
            LicenseGateOrQuit = False

        Case "UNKNOWN_KEY"
            LicenseBlock "Licencni kljuc nije prepoznat.", _
                         "Proverite kljuc ili kontaktirajte dobavljaca."
            LicenseGateOrQuit = False

        Case "BAD_DEVICE"
            ' N2: otisak preslab (prakticno nedostizno — klijent pre-proverava).
            ' Vezanu masinu propusti; inace jasna poruka o uredjaju.
            If LicenseIsBoundMachine(bound, parts) Then
                LicenseGateOrQuit = True
            Else
                LicenseBlock "Ne mogu pouzdano da ocitam ovaj uredjaj.", _
                             "WMI/registry nedostupan. Kontaktirajte dobavljaca."
                LicenseGateOrQuit = False
            End If

        Case Else
            ' N3: prolazna/neocekivana greska servera (success:false -> ERROR,
            ' LOCK_TIMEOUT bez status polja, prazan status...). Tretiraj kao
            ' privremeno: vezana masina -> offline grace; inace pozovi na ponovni
            ' pokusaj (NE alarmantno "kontaktirajte dobavljaca").
            If LicenseIsBoundMachine(bound, parts) Then
                LogWarn SRC, "Prolazna greska/status='" & status & Poruka("LIC_MSG_PROPUSTAM_VEZANU_MASINU")
                LicenseGateOrQuit = True
            Else
                LicenseBlock "Licencni server trenutno nije dostupan.", _
                             "Pokusajte ponovo za koji minut."
                LicenseGateOrQuit = False
            End If
    End Select
    Exit Function

EH:
    ' Bug u proveri ne sme da zakljuca korisnika -> fail-OPEN + log
    ' (isti princip kao modTrial; prava zastita je ionako server-side).
    LogErr SRC
    LicenseGateOrQuit = True
End Function

' ============================================================
' PRIVATE — gate helperi (bound check, anti-rollback, inline aktivacija)
' ============================================================

' Da li je ovo VEZANA masina: ima sacuvane BOUND_PARTS i fuzzy match >= prag.
Private Function LicenseIsBoundMachine(ByVal bound As String, ByVal parts As String) As Boolean
    LicenseIsBoundMachine = (Len(bound) > 0 And LicPartsMatch(parts, bound) >= LIC_MIN_MATCH)
End Function

' Anti-rollback: da li je danasnji datum ISPOD ranije vidjenog (LICENSE_HWM).
Private Function LicenseClockRolledBack(ByVal today As Date) As Boolean
    On Error Resume Next
    Dim hwm As String: hwm = Replace(Trim$(GetConfigValue(CFG_LIC_HWM)), "T", " ")
    If Len(hwm) > 0 Then
        If IsDate(hwm) Then LicenseClockRolledBack = (today < CDate(hwm))
    End If
End Function

' Pomeri LICENSE_HWM na najkasniji vidjeni datum (nikad unazad).
Private Sub LicenseAdvanceHwm(ByVal today As Date)
    On Error Resume Next
    Dim hwm As String: hwm = Replace(Trim$(GetConfigValue(CFG_LIC_HWM)), "T", " ")
    If Len(hwm) = 0 Or (IsDate(hwm) And today > CDate(hwm)) Then
        SetConfigValue CFG_LIC_HWM, Format$(today, "yyyy-mm-dd")
    End If
End Sub

' N9: probni period istekao a licenca je ukljucena -> ponudi unos kljuca odmah
' (jedan InputBox, bez prethodnog trial-blok dijaloga). Ako korisnik odustane,
' blokira kao i inace.
Private Function PromptLicenseOnTrialExpiry() As Boolean
    On Error GoTo EH
    Application.Visible = True
    Dim key As String
    key = Trim$(InputBox( _
        "Probni period je istekao." & vbCrLf & vbCrLf & _
        "Unesite licencni kljuc za nastavak (Cancel = izlaz):", APP_NAME))
    If Len(key) = 0 Then
        LicenseBlock "Probni period je istekao.", _
                     "Unesite licencu (Alt+F8 -> ActivateLicensePrompt) ili kontaktirajte dobavljaca."
        PromptLicenseOnTrialExpiry = False
        Exit Function
    End If
    SetConfigValue CFG_LIC_KEY, key
    SetConfigValue CFG_LIC_NEXT_CHECK, ""
    SetConfigValue CFG_LIC_BOUND_PARTS, ""
    PromptLicenseOnTrialExpiry = LicenseGateOrQuit()    ' standardna online aktivacija
    Exit Function
EH:
    LogErr "modLicense.PromptLicenseOnTrialExpiry"
    PromptLicenseOnTrialExpiry = True                   ' fail-open
End Function

' ============================================================
' PUBLIC — Jednokratna aktivacija na novom racunaru
' ============================================================

Public Sub ActivateLicensePrompt()
    Const SRC As String = "modLicense.ActivateLicensePrompt"
    On Error GoTo EH

    Dim key As String
    key = Trim$(InputBox("Unesite licencni kljuc za OVAJ racunar:", APP_NAME, _
                         Trim$(GetConfigValue(CFG_LIC_KEY))))
    If Len(key) = 0 Then Exit Sub

    SetConfigValue CFG_LIC_KEY, key
    ' Forsiraj svezu online proveru
    SetConfigValue CFG_LIC_NEXT_CHECK, ""
    SetConfigValue CFG_LIC_BOUND_PARTS, ""

    Dim endpoint As String: endpoint = LicenseEndpoint()
    If Len(endpoint) = 0 Then
        MsgBox "LICENSE_ENDPOINT (ili MONITORING_ENDPOINT) nije podesen u tblSEFConfig.", _
               vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim parts As String: parts = GetDeviceParts()
    If LicNonEmptyParts(parts) < LIC_MIN_MATCH Then
        MsgBox "Ne mogu da ocitam dovoljno podataka o uredjaju (WMI nedostupan?).", _
               vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim resp As String
    Dim httpOk As Boolean
    LicenseStatusBar "Proveravam licencu..."
    httpOk = LicenseHttpCheck(endpoint, key, parts, resp)
    LicenseStatusBar ""
    If Not httpOk Then
        MsgBox "Server nije dostupan. Proverite internet konekciju.", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim status As String: status = UCase$(ExtractJsonStringGoogle(resp, "status"))
    If status = "OK" Then
        Call PersistLicenseOk(parts, resp)
        MsgBox "Licenca je uspesno aktivirana na ovom racunaru." & vbCrLf & _
               "Korisnik: " & ExtractJsonStringGoogle(resp, "customer"), _
               vbInformation, APP_NAME
    Else
        MsgBox "Aktivacija nije uspela (" & status & ")." & vbCrLf & vbCrLf & _
               LicenseErr(resp), vbExclamation, APP_NAME
    End If
    Exit Sub

EH:
    LogErr SRC
    MsgBox "Greska pri aktivaciji: " & Err.description, vbCritical, APP_NAME
End Sub

' Dijagnostika: prikazi otisak ovog racunara (za support / rucni bind).
Public Sub LicenseShowDevice()
    Dim parts As String: parts = GetDeviceParts()
    Dim p() As String: p = LicSplitParts(parts)
    MsgBox "Otisak ovog racunara:" & vbCrLf & vbCrLf & _
           "MachineGuid : " & p(0) & vbCrLf & _
           "SMBIOS UUID : " & p(1) & vbCrLf & _
           "Volume SN   : " & p(2) & vbCrLf & vbCrLf & _
           "Racunar: " & Environ$("COMPUTERNAME"), vbInformation, APP_NAME
End Sub

' ============================================================
' PRIVATE — perzistencija uspesne provere
' ============================================================

Private Sub PersistLicenseOk(ByVal parts As String, ByVal resp As String)
    Dim grace As Long
    grace = CLng(val(ExtractJsonStringGoogle(resp, "graceDays")))
    If grace <= 0 Then grace = LIC_DEFAULT_GRACE_DAYS

    SetConfigValue CFG_LIC_TOKEN, ExtractJsonStringGoogle(resp, "token")
    SetConfigValue CFG_LIC_BOUND_PARTS, parts
    SetConfigValue CFG_LIC_NEXT_CHECK, Format$(Now + grace, "yyyy-mm-dd\Thh:nn:ss")
    SetConfigValue CFG_LIC_STATUS, "OK"
End Sub

' ============================================================
' PRIVATE — HTTP ka GAS-u (sinhrono; isti obrazac kao modMonitoring)
' ============================================================

Private Function LicenseHttpCheck(ByVal endpoint As String, _
                                  ByVal key As String, _
                                  ByVal parts As String, _
                                  ByRef respOut As String) As Boolean
    Const SRC As String = "modLicense.LicenseHttpCheck"
    On Error GoTo EH

    Dim p() As String: p = LicSplitParts(parts)

    Dim body As String
    body = "{""action"":""checkLicense""" & _
           ",""licenseKey"":""" & JsonEscape(key) & """" & _
           ",""components"":[""" & JsonEscape(p(0)) & """,""" & JsonEscape(p(1)) & """,""" & JsonEscape(p(2)) & """]" & _
           ",""appVersion"":""" & JsonEscape(APP_VERSION) & """" & _
           ",""buildSha"":""" & JsonEscape(BUILD_SHA) & """" & _
           ",""buildVersion"":""" & JsonEscape(BUILD_VERSION) & """" & _
           ",""buildDate"":""" & JsonEscape(BUILD_DATE) & """" & _
           ",""computerName"":""" & JsonEscape(Environ$("COMPUTERNAME")) & """}"

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", endpoint, False
    http.SetTimeouts HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS, HTTP_RECV_TIMEOUT_MS
    http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.SetRequestHeader "Accept", "application/json"
    http.Send body

    If http.status < 200 Or http.status >= 300 Then
        LogWarn SRC, "HTTP " & http.status
        LicenseHttpCheck = False
        Exit Function
    End If

    respOut = CStr(http.responseText)
    LicenseHttpCheck = (Len(respOut) > 0)
    Exit Function

EH:
    LogErr SRC
    LicenseHttpCheck = False
End Function

' ============================================================
' Otisak uredjaja (GetDeviceParts Public za integ-test; Read* ostaju Private)
' ============================================================

' Vraca "MACHINEGUID|UUID|VOLSERIAL" (uppercased, trimovano). Cista funkcija
' bez side-efekata; Public radi modLicenseTests.TestLicense_DeviceFingerprint.
Public Function GetDeviceParts() As String
    Dim c1 As String, c2 As String, c3 As String
    c1 = UCase$(Trim$(ReadMachineGuid()))
    c2 = UCase$(Trim$(ReadSmbiosUuid()))
    c3 = UCase$(Trim$(ReadVolumeSerial()))
    GetDeviceParts = c1 & "|" & c2 & "|" & c3
End Function

' Windows OS-install GUID. Stabilan, jedinstven po Windows instalaciji.
Private Function ReadMachineGuid() As String
    On Error Resume Next
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    ReadMachineGuid = sh.RegRead("HKLM\SOFTWARE\Microsoft\Cryptography\MachineGuid")
End Function

' SMBIOS / maticna ploca UUID. Razlicit po VM-u, vrlo stabilan.
Private Function ReadSmbiosUuid() As String
    On Error Resume Next
    Dim wmi As Object, item As Object
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    For Each item In wmi.ExecQuery("SELECT UUID FROM Win32_ComputerSystemProduct")
        ReadSmbiosUuid = CStr(item.UUID)
        Exit For
    Next item
End Function

' Volume serial sistemskog diska. Menja se na reformat.
Private Function ReadVolumeSerial() As String
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim sysDrive As String
    sysDrive = Environ$("SystemDrive")
    If Len(sysDrive) = 0 Then sysDrive = "C:"
    ReadVolumeSerial = Hex$(fso.GetDrive(sysDrive).SerialNumber)
End Function

' ============================================================
' PUBLIC — helperi za poredjenje komponenti
' (cista logika, Public radi testabilnosti — vidi modLicenseTests)
' ============================================================

' Uvek vrati niz od bar 3 elementa (pad praznima).
Public Function LicSplitParts(ByVal s As String) As String()
    LicSplitParts = Split(s & "||", "|")
End Function

' Broj poklapanja (po poziciji) ne-praznih komponenti, max 3.
Public Function LicPartsMatch(ByVal a As String, ByVal b As String) As Long
    Dim pa() As String, pb() As String
    pa = LicSplitParts(a)
    pb = LicSplitParts(b)
    Dim i As Long, m As Long
    For i = 0 To 2
        If Len(Trim$(pa(i))) > 0 Then
            If StrComp(Trim$(pa(i)), Trim$(pb(i)), vbTextCompare) = 0 Then m = m + 1
        End If
    Next i
    LicPartsMatch = m
End Function

' Broj ne-praznih komponenti (max 3).
Public Function LicNonEmptyParts(ByVal a As String) As Long
    Dim pa() As String: pa = LicSplitParts(a)
    Dim i As Long, n As Long
    For i = 0 To 2
        If Len(Trim$(pa(i))) > 0 Then n = n + 1
    Next i
    LicNonEmptyParts = n
End Function

' ============================================================
' PRIVATE — config / poruke
' ============================================================

' Latch: da li je na OVOM racunaru licenca vec jednom USPESNO aktivirana.
' Signal je par koji upisuje PersistLicenseOk: licencni kljuc (LICENSE_KEY) +
' sacuvane komponente otiska (LICENSE_BOUND_PARTS). Kada oba postoje, gate je
' OBAVEZAN cak i ako neko naknadno vrati LICENSE_ENABLED na NO -> zatvara
' "spusti flag na NO" bypass. Granica (posteno): ko lokalno obrise oba kljuca
' gubi aktivaciju i latch; pravi autoritet ostaje server.
Private Function LicenseActivatedOnThisMachine() As Boolean
    LicenseActivatedOnThisMachine = _
        (Len(Trim$(GetConfigValue(CFG_LIC_KEY))) > 0) And _
        (Len(Trim$(GetConfigValue(CFG_LIC_BOUND_PARTS))) > 0)
End Function

' Da li gate UOPSTE treba da radi: opt-in flag YES ILI latch (vec aktivirana
' masina). Sve "da li uopste proveravati" tacke gledaju ovo umesto golog
' LicenseEnabled() — tako se anti-bypass primenjuje na jednom mestu.
Private Function LicenseRequired() As Boolean
    LicenseRequired = LicenseEnabled() Or LicenseActivatedOnThisMachine()
End Function

Private Function LicenseEnabled() As Boolean
    Dim v As String
    v = UCase$(Trim$(GetConfigValue(CFG_LIC_ENABLED)))
    Select Case v
        Case "YES", "TRUE", "1", "ON", "ENABLED": LicenseEnabled = True
        Case Else: LicenseEnabled = False
    End Select
End Function

Private Function LicenseEndpoint() As String
    ' Pin (ako postoji) je autoritet — config NE moze da ga prebaci na lazni server.
    If Len(Trim$(LIC_ENDPOINT_PINNED)) > 0 Then
        LicenseEndpoint = Trim$(LIC_ENDPOINT_PINNED)
        Exit Function
    End If
    Dim e As String
    e = Trim$(GetConfigValue(CFG_LIC_ENDPOINT))
    If Len(e) = 0 Then e = Trim$(GetConfigValue(CFG_MON_ENDPOINT))
    LicenseEndpoint = e
End Function

Private Function LicenseErr(ByVal resp As String) As String
    Dim e As String
    e = ExtractJsonStringGoogle(resp, "error")
    If Len(Trim$(e)) = 0 Then e = "Kontaktirajte dobavljaca."
    LicenseErr = e
End Function

' Status-bar hint tokom online provere. Prazan msg = reset (oslobodi status
' bar + kursor). Best-effort: ako je DisplayStatusBar iskljucen, nista se ne
' vidi, ali ne smeta. Isti obrazac kao ThisWorkbook.Workbook_BeforeClose.
Private Sub LicenseStatusBar(ByVal msg As String)
    On Error Resume Next
    If Len(msg) > 0 Then
        Application.Cursor = xlWait
        Application.StatusBar = msg
    Else
        Application.StatusBar = False
        Application.Cursor = xlDefault
    End If
End Sub

Private Sub LicenseBlock(ByVal reason As String, ByVal hint As String)
    On Error Resume Next
    Application.Visible = True
    MsgBox reason & vbCrLf & vbCrLf & hint & vbCrLf & vbCrLf & _
           "Kontaktirajte dobavljaca za nastavak rada.", vbCritical, APP_NAME
    DenyAccessAndScheduleClose
End Sub

' ============================================================
' PUBLIC — Deljeni "access denied" + Workbook_Open integracija
' (koriste ga i license i trial gate)
' ============================================================

' Da li je poslednji gate (licenca ILI trial) odbio pristup. Workbook_Open
' ovo koristi za rani prekid (da ne pokrece dalji startup za odbijenu masinu).
Public Function AccessWasDenied() As Boolean
    AccessWasDenied = gAccessDenied
End Function

' Deljena blokada: oznaci pristup odbijenim i zakazi POUZDANO zatvaranje na
' sledeci idle tick. Pozivaju je i LicenseBlock i modTrial.TrialBlock — nikada
' ne zatvarati svesku sinhrono iz Workbook_Open toka (Excel Close odlaze/ignorise).
Public Sub DenyAccessAndScheduleClose()
    On Error Resume Next
    gAccessDenied = True
    Application.OnTime Now + TimeSerial(0, 0, 1), "modLicense.ForceCloseDeniedWorkbook"
End Sub

' OnTime cilj: stvarno zatvaranje sveske posle blokade. Izvrsava se tek kad se
' Workbook_Open zavrsi, pa Excel ovog puta postuje Close.
' Mora biti Public (Application.OnTime ne moze da pozove Private proceduru).
Public Sub ForceCloseDeniedWorkbook()
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Saved = True
    ThisWorkbook.Close SaveChanges:=False
End Sub

