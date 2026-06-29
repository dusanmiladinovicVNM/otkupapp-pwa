Attribute VB_Name = "modUpdateGate"
Option Explicit

' ============================================================
' modUpdateGate - minimalni "min-version gate" za flotu.
'
' Cilj: niko ne ostane na zastareloj verziji. Server (GAS, public action
' "checkVersion") javlja minimalnu dozvoljenu verziju; klijent poredi sa
' svojom APP_VERSION i, ako je ispod:
'   - enforce=YES -> BLOKIRA pokretanje (kao license blok),
'   - inace       -> UPOZORI ali pusti da radi (default, bezbedan rollout).
'
' OPT-IN: radi samo ako su MONITORING_ENDPOINT i MONITORING_SECRET podeseni
' (isti uslov kao monitoring = "u floti si"). Inace ne dira nista.
'
' FAIL-OPEN: svaka greska/timeout/nedostupan server -> propusti (nikad ne
' brick-uj korisnika zbog mreze). Pravi autoritet je server.
'
' Poziva se iz modMain.StartApp posle AccessGateOrQuit:
'   If Not UpdateGateOrQuit() Then Exit Sub
'
' Server config (GAS Script Properties; menja se BEZ redeploy-a):
'   VERSION_MIN      npr. "2.2.0"  (minimalna dozvoljena; prazno = gate off)
'   VERSION_LATEST   npr. "2.3.0"  (samo za poruku korisniku)
'   VERSION_ENFORCE  YES/NO        (default NO = samo upozorenje)
'   VERSION_MESSAGE  opciono custom tekst
' ============================================================

Private Const CFG_MON_ENDPOINT As String = "MONITORING_ENDPOINT"
Private Const CFG_MON_SECRET   As String = "MONITORING_SECRET"
Private Const HTTP_TIMEOUT_MS  As Long = 2500

' Vraca True ako sme da nastavi; False samo kad server enforce-uje a verzija
' je ispod minimuma (tada je vec zakazano zatvaranje sveske).
Public Function UpdateGateOrQuit() As Boolean
    Const SRC As String = "modUpdateGate.UpdateGateOrQuit"
    On Error GoTo EH

    UpdateGateOrQuit = True                      ' default: propusti (fail-open)

    Dim endpoint As String: endpoint = Trim$(GetConfigValue(CFG_MON_ENDPOINT))
    Dim secret As String:   secret = Trim$(GetConfigValue(CFG_MON_SECRET))
    If Len(endpoint) = 0 Or Len(secret) = 0 Then Exit Function   ' nije u floti -> ne diramo

    Dim resp As String
    If Not VersionHttpCheck(endpoint, secret, resp) Then Exit Function   ' offline -> fail-open

    Dim minVer As String: minVer = Trim$(ExtractJsonStringGoogle(resp, "minVersion"))
    If Len(minVer) = 0 Then Exit Function         ' server nema policy -> propusti
    If VersionCompare(APP_VERSION, minVer) >= 0 Then Exit Function   ' dovoljno nova

    ' --- verzija je ISPOD minimuma ---
    Dim latest As String: latest = Trim$(ExtractJsonStringGoogle(resp, "latestVersion"))
    Dim msg As String:    msg = Trim$(ExtractJsonStringGoogle(resp, "message"))
    If Len(msg) = 0 Then
        msg = "Ova verzija (" & APP_VERSION & ") je zastarela."
        If Len(latest) > 0 Then msg = msg & " Najnovija verzija: " & latest & "."
    End If

    Dim enforce As String: enforce = UCase$(Trim$(ExtractJsonStringGoogle(resp, "enforce")))
    Select Case enforce
        Case "YES", "TRUE", "1", "ON"
            Application.Visible = True
            MsgBox msg & vbCrLf & vbCrLf & _
                   "Potrebno je azuriranje pre nastavka rada." & vbCrLf & _
                   "Kontaktirajte dobavlja" & ChrW(269) & "a za novu verziju.", vbCritical, APP_NAME
            DenyAccessAndScheduleClose            ' isti mehanizam kao license blok
            UpdateGateOrQuit = False
        Case Else
            ' WARN (default): obavesti, ali pusti da radi.
            Application.Visible = True
            MsgBox msg & vbCrLf & vbCrLf & _
                   "Preporucuje se azuriranje. Kontaktirajte dobavlja" & ChrW(269) & "a.", _
                   vbExclamation, APP_NAME
            UpdateGateOrQuit = True
    End Select
    Exit Function

EH:
    ' Bug u gate-u ne sme da zakljuca korisnika -> fail-OPEN + log.
    LogErr SRC
    UpdateGateOrQuit = True
End Function

' ============================================================
' PRIVATE - sinhroni HTTP ka GAS-u (isti obrazac kao modLicense.LicenseHttpCheck)
' ============================================================

Private Function VersionHttpCheck(ByVal endpoint As String, _
                                  ByVal secret As String, _
                                  ByRef respOut As String) As Boolean
    Const SRC As String = "modUpdateGate.VersionHttpCheck"
    On Error GoTo EH

    Dim body As String
    body = "{""action"":""checkVersion""" & _
           ",""monitoringSecret"":""" & JsonEscape(secret) & """" & _
           ",""appVersion"":""" & JsonEscape(APP_VERSION) & """" & _
           ",""buildVersion"":""" & JsonEscape(BUILD_VERSION) & """" & _
           ",""buildSha"":""" & JsonEscape(BUILD_SHA) & """}"

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", endpoint, False
    http.SetTimeouts HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS
    http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.SetRequestHeader "Accept", "application/json"
    http.Send body

    If http.status < 200 Or http.status >= 300 Then
        LogWarn SRC, "HTTP " & http.status
        VersionHttpCheck = False
        Exit Function
    End If

    respOut = CStr(http.responseText)
    VersionHttpCheck = (Len(respOut) > 0)
    Exit Function

EH:
    ' Offline / WinHttp greska -> "ne znam"; pozivalac fail-open propusti.
    VersionHttpCheck = False
End Function

' ============================================================
' PUBLIC - poredjenje verzija (cista logika; Public radi testabilnosti)
' ============================================================

' Poredi dve verzije po numerickim segmentima (SemVer-ish). Skida prefiks
' "vba-v"/"vba-"/"v" i odseca suffiks od prvog "-" ili "+" (git-describe
' "2.2.1-3-gabc1234" -> "2.2.1"; "0.0.0-dev" -> "0.0.0"). Poredi major.minor.patch.
' Vraca -1 (a<b), 0 (a=b), 1 (a>b).
Public Function VersionCompare(ByVal a As String, ByVal b As String) As Long
    Dim pa() As String: pa = VersionParts(a)
    Dim pb() As String: pb = VersionParts(b)
    Dim i As Long, va As Long, vb As Long
    For i = 0 To 2
        va = 0: vb = 0
        If i <= UBound(pa) Then va = SafeLng(pa(i))
        If i <= UBound(pb) Then vb = SafeLng(pb(i))
        If va < vb Then VersionCompare = -1: Exit Function
        If va > vb Then VersionCompare = 1:  Exit Function
    Next i
    VersionCompare = 0
End Function

' Normalizuje verziju u niz numerickih segmenata (uvek bar 3 elementa).
Private Function VersionParts(ByVal s As String) As String()
    Dim t As String: t = LCase$(Trim$(s))
    If Left$(t, 5) = "vba-v" Then
        t = Mid$(t, 6)
    ElseIf Left$(t, 4) = "vba-" Then
        t = Mid$(t, 5)
    ElseIf Left$(t, 1) = "v" Then
        t = Mid$(t, 2)
    End If
    Dim p As Long
    p = InStr(t, "-"): If p > 0 Then t = Left$(t, p - 1)
    p = InStr(t, "+"): If p > 0 Then t = Left$(t, p - 1)
    VersionParts = Split(t & ".0.0", ".")
End Function

Private Function SafeLng(ByVal s As String) As Long
    On Error Resume Next
    SafeLng = CLng(val(s))
End Function

' ============================================================
' PUBLIC - rucna provera self-update kanala (AgriX_Release/version.json)
' ============================================================

' Read-only: vrati app_version objavljen u AgriX_Release/version.json (isti
' izvor i format kao modSelfUpdate.CheckForUpdateOnOpen), ili "" ako nedostupno
' (opt-out / offline / nije objavljeno). BEZ UI i side-efekata.
' Reuse modDrive (find/download) + modGoogleAuth (JSON) + REL_FOLDER_ID.
' Stoji OVDE (a ne u modSelfUpdate) namerno: modSelfUpdate je SKIP_MODULES (ne
' stize preko self-update-a), pa bi pozivalac (Admin panel) posle inkrementalnog
' update-a dobio "Sub not defined". modUpdateGate se update-uje normalno.
Public Function ReleaseManifestVersion() As String
    Const SRC As String = "modUpdateGate.ReleaseManifestVersion"
    On Error GoTo EH

    If Len(REL_FOLDER_ID) = 0 Then Exit Function            ' opt-in (kao self-update)
    Dim id As String: id = DriveFindInFolder(REL_FOLDER_ID, "version.json")
    If Len(id) = 0 Then Exit Function
    Dim tmp As String: tmp = Environ$("TEMP") & "\AgriX_relcheck.json"
    If Not DriveDownloadToFile(id, tmp) Then Exit Function

    ReleaseManifestVersion = Trim$(ExtractJsonStringGoogle(ReadManifestFile(tmp), "app_version"))
    Exit Function
EH:
    LogErr SRC, Err.description
End Function

Private Function ReadManifestFile(ByVal path As String) As String
    Dim ff As Integer: ff = FreeFile
    Open path For Input As #ff
    If LOF(ff) > 0 Then ReadManifestFile = Input$(LOF(ff), ff)
    Close #ff
End Function
