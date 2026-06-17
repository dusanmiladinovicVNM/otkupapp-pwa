Attribute VB_Name = "modLicense"
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
' OPT-IN: gate radi SAMO ako je u tblSEFConfig LICENSE_ENABLED = YES.
' Tako merge ovog koda NE blokira postojece instalacije dok ne provizionises
' kljuceve i ne ukljucis proveru.
'
' Config kljucevi (tblSEFConfig):
'   LICENSE_ENABLED      YES/NO  (default NO — provera iskljucena)
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
Private Const CFG_MON_ENDPOINT    As String = "MONITORING_ENDPOINT"

' Koliko od 3 komponente otiska mora da se poklopi (fuzzy match) da bi se
' tolerisala manja promena hardvera (nov disk, reinstall) bez lockout-a.
Private Const LIC_MIN_MATCH As Long = 2

' Fallback offline grace (dana) ako server ne posalje graceDays.
Private Const LIC_DEFAULT_GRACE_DAYS As Long = 7

Private Const HTTP_TIMEOUT_MS As Long = 8000
Private Const HTTP_RECV_TIMEOUT_MS As Long = 15000

' ============================================================
' PUBLIC — Gate (poziva se iz modMain.StartApp)
' ============================================================

' Vraca True ako sme da nastavi; inace blokira, prikaze poruku i zatvori svesku.
Public Function LicenseGateOrQuit() As Boolean
    Const SRC As String = "modLicense.LicenseGateOrQuit"
    On Error GoTo EH

    ' Opt-in: ako provera nije ukljucena, ne diramo nista.
    If Not LicenseEnabled() Then
        LicenseGateOrQuit = True
        Exit Function
    End If

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
    If NonEmptyParts(parts) < LIC_MIN_MATCH Then
        ' Otisak preslab (npr. WMI nedostupan). Ne kaznjavamo korisnika.
        LogWarn SRC, "Nedovoljno komponenti otiska uredjaja. Preskacem proveru (fail-open)."
        LicenseGateOrQuit = True
        Exit Function
    End If

    Dim bound As String: bound = Trim$(GetConfigValue(CFG_LIC_BOUND_PARTS))
    Dim nextChk As String: nextChk = Trim$(GetConfigValue(CFG_LIC_NEXT_CHECK))

    ' --- BRZI OFFLINE PUT: unutar grace prozora + ovo je vezana masina ---
    If Len(bound) > 0 Then
        If PartsMatch(parts, bound) >= LIC_MIN_MATCH Then
            If Len(nextChk) > 0 Then
                If IsDate(nextChk) Then
                    If Now < CDate(nextChk) Then
                        LicenseGateOrQuit = True
                        Exit Function
                    End If
                End If
            End If
        End If
    End If

    ' --- ONLINE PROVERA (istekao grace ili nema kesa) ---
    Dim resp As String
    If Not LicenseHttpCheck(endpoint, key, parts, resp) Then
        ' Server nedostupan. Ako je OVO vec vezana masina -> offline grace,
        ' da privremeni nestanak interneta ne zakljuca platisu. Ako masina
        ' NIJE vezana (nema validnog kesa) -> blokiraj (aktivacija mora online).
        If Len(bound) > 0 And PartsMatch(parts, bound) >= LIC_MIN_MATCH Then
            LogWarn SRC, "Server nedostupan — offline grace (masina je vezana)."
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

        Case Else
            ' Neocekivan odgovor: ne brick-uj vec vezanu (placenu) masinu.
            If Len(bound) > 0 And PartsMatch(parts, bound) >= LIC_MIN_MATCH Then
                LogWarn SRC, "Neocekivan status='" & status & "' — propustam vezanu masinu."
                LicenseGateOrQuit = True
            Else
                LicenseBlock "Neocekivan odgovor servera za licencu.", Left$(resp, 200)
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
    If NonEmptyParts(parts) < LIC_MIN_MATCH Then
        MsgBox "Ne mogu da ocitam dovoljno podataka o uredjaju (WMI nedostupan?).", _
               vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim resp As String
    If Not LicenseHttpCheck(endpoint, key, parts, resp) Then
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
    Dim p() As String: p = SplitParts(parts)
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

    Dim p() As String: p = SplitParts(parts)

    Dim body As String
    body = "{""action"":""checkLicense""" & _
           ",""licenseKey"":""" & JsonEscape(key) & """" & _
           ",""components"":[""" & JsonEscape(p(0)) & """,""" & JsonEscape(p(1)) & """,""" & JsonEscape(p(2)) & """]" & _
           ",""appVersion"":""" & JsonEscape(APP_VERSION) & """" & _
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
' PRIVATE — otisak uredjaja
' ============================================================

' Vraca "MACHINEGUID|UUID|VOLSERIAL" (uppercased, trimovano).
Private Function GetDeviceParts() As String
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
' PRIVATE — helperi za poredjenje komponenti
' ============================================================

' Uvek vrati niz od bar 3 elementa (pad praznima).
Private Function SplitParts(ByVal s As String) As String()
    SplitParts = Split(s & "||", "|")
End Function

' Broj poklapanja (po poziciji) ne-praznih komponenti, max 3.
Private Function PartsMatch(ByVal a As String, ByVal b As String) As Long
    Dim pa() As String, pb() As String
    pa = SplitParts(a)
    pb = SplitParts(b)
    Dim i As Long, m As Long
    For i = 0 To 2
        If Len(Trim$(pa(i))) > 0 Then
            If StrComp(Trim$(pa(i)), Trim$(pb(i)), vbTextCompare) = 0 Then m = m + 1
        End If
    Next i
    PartsMatch = m
End Function

' Broj ne-praznih komponenti (max 3).
Private Function NonEmptyParts(ByVal a As String) As Long
    Dim pa() As String: pa = SplitParts(a)
    Dim i As Long, n As Long
    For i = 0 To 2
        If Len(Trim$(pa(i))) > 0 Then n = n + 1
    Next i
    NonEmptyParts = n
End Function

' ============================================================
' PRIVATE — config / poruke
' ============================================================

Private Function LicenseEnabled() As Boolean
    Dim v As String
    v = UCase$(Trim$(GetConfigValue(CFG_LIC_ENABLED)))
    Select Case v
        Case "YES", "TRUE", "1", "ON", "ENABLED": LicenseEnabled = True
        Case Else: LicenseEnabled = False
    End Select
End Function

Private Function LicenseEndpoint() As String
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

Private Sub LicenseBlock(ByVal reason As String, ByVal hint As String)
    On Error Resume Next
    Application.Visible = True
    MsgBox reason & vbCrLf & vbCrLf & hint & vbCrLf & vbCrLf & _
           "Kontaktirajte dobavljaca za nastavak rada.", vbCritical, APP_NAME

    ' Zatvori SAMO ovu svesku (bez Application.Quit). Saved=True spreci prompt.
    ThisWorkbook.Saved = True
    Application.DisplayAlerts = False
    ThisWorkbook.Close SaveChanges:=False
End Sub
