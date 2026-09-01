Attribute VB_Name = "modSetup"

Option Explicit

' ============================================================
' modSetup - OtkupApp new PC setup / workstation health-check
'
' CONFIG STRATEGY:
'
' 1) tblConfig  (legacy)
'    Istorijski je drzao Google/PWA kljuceve. Danas Google/PWA config zivi u
'    tblSEFConfig -- runtime (modGoogleAuth, modBrojevi, modGoogleSyncOrchestrator)
'    i Podesavanja ga citaju/pisu preko Get/SetConfigValue, pa i setup provera
'    (CheckGoogleOAuthConfig) gleda tblSEFConfig. tblConfig se ovde samo kolonski
'    validira ako uopste postoji (ne pise se u njega).
'
' 2) tblLocalConfig
'    Local workstation setup config:
'       Kljuc | Vrednost | Opis
'
'    Auto-created by this module if missing.
'
' 3) tblSEFConfig
'    Existing SEF config used by GetConfigValue().
'
' ============================================================

Private Const SETUP_LOG_SHEET As String = "SETUP_LOG"

Private Const LOCAL_CONFIG_SHEET As String = "LocalConfig"
Private Const TBL_LOCAL_CONFIG As String = "tblLocalConfig"

Private Const CFG_KEY As String = "Kljuc"
Private Const CFG_VALUE As String = "Vrednost"
Private Const CFG_DESC As String = "Opis"

' ============================================================
' PUBLIC ENTRY POINTS
' ============================================================

Public Sub SetupNewPC()
    On Error GoTo EH

    Dim report As String
    report = vbNullString

    InitSetupLog
    EnsureLocalConfigTable
    EnsurePoruke

    LogSetup "INFO", "SetupNewPC started"
    LogSetup "INFO", "Workbook: " & ThisWorkbook.fullName
    LogSetup "INFO", "Machine: " & Environ$("COMPUTERNAME")
    LogSetup "INFO", "Windows user: " & Environ$("USERNAME")
    LogSetup "INFO", "Excel version: " & Application.Version
    LogSetup "INFO", "Cloud/PWA sync: " & IIf(IsCloudSyncEnabled(), "ON", "OFF (desktop-only)")

    report = report & CheckRuntimeEnvironment()
    report = report & EnsureAppFolders()
    report = report & SetupBankFolders()
    report = report & CheckGoogleOAuthConfig()
    report = report & CheckSEFConfigForSetup()
    report = report & CheckRequiredTablesForSetup()
    report = report & CheckRequiredColumnsForSetup()

    ' Zivi server-link je ADVISORY: prikazuje se, ali NE ulazi u `report` -> ne obara
    ' zeleno (APP_SETUP_COMPLETED). Offline / ne-autentifikovan Google i sl. daju samo
    ' NAPOMENU; setup i dalje moze da prodje ako je lokalni deo ispravan.
    Dim linkWarn As String
    linkWarn = CheckServerLink()

    If Len(report) = 0 Then
        SetLocalConfigValue "APP_SETUP_COMPLETED", "DA", Poruka("SETUP_MSG_OVAJ_RACUNAR_PROSAO")
        SetLocalConfigValue "APP_SETUP_COMPLETED_AT", Format$(Now, "yyyy-mm-dd hh:nn:ss"), Poruka("SETUP_MSG_DATUM_VREME_ZAVRSETKA")
        SetLocalConfigValue "APP_SETUP_MACHINE_NAME", Environ$("COMPUTERNAME"), "Naziv racunara"
        SetLocalConfigValue "APP_SETUP_WINDOWS_USER", Environ$("USERNAME"), "Windows korisnik"
        SetLocalConfigValue "APP_LAST_HEALTHCHECK_AT", Format$(Now, "yyyy-mm-dd hh:nn:ss"), "Poslednji health-check"

        ' Anti-tamper: sakrij tblSEFConfig (VeryHidden) cim je setup zelen
        ' (config je popunjen). Operativna podesavanja se od sada uredjuju kroz
        ' UI (Maticni podaci -> Podesavanja). Izlaz u nuzdi: Alt+F8 -> ShowConfigSheet.
        On Error Resume Next
        modPodesavanja.HideConfigSheet
        On Error GoTo EH

        LogSetup "OK", "Setup completed successfully"
        If Len(linkWarn) > 0 Then LogSetup "WARN", "Server link: " & linkWarn

        Dim okMsg As String
        okMsg = Poruka("SETUP_MSG_SETUP_USPESNO_ZAVRSEN") & vbCrLf & vbCrLf & _
                "Aplikacija je spremna za ovaj racunar." & vbCrLf & _
                Poruka("SETUP_MSG_PODESAVANJA_MATICNI_PODACI")

        If Len(linkWarn) > 0 Then
            okMsg = okMsg & vbCrLf & vbCrLf & _
                    "NAPOMENA - server link (ne blokira setup):" & vbCrLf & linkWarn
            MsgBox okMsg, vbExclamation, APP_NAME
        Else
            MsgBox okMsg, vbInformation, APP_NAME
        End If
    Else
        SetLocalConfigValue "APP_SETUP_COMPLETED", "NE", Poruka("SETUP_MSG_OVAJ_RACUNAR_PROSAO")
        SetLocalConfigValue "APP_LAST_HEALTHCHECK_AT", Format$(Now, "yyyy-mm-dd hh:nn:ss"), "Poslednji health-check"

        Dim failMsg As String
        failMsg = report
        If Len(linkWarn) > 0 Then _
            failMsg = failMsg & vbCrLf & "Server link (ne blokira):" & vbCrLf & linkWarn

        LogSetup "WARN", failMsg

        MsgBox Poruka("SETUP_MSG_SETUP_ZAVRSEN_ALI") & _
               vbCrLf & vbCrLf & failMsg, _
               vbExclamation, APP_NAME
    End If

    Exit Sub

EH:
    LogSetup "ERROR", "SetupNewPC failed: " & Err.Number & " - " & Err.description
    MsgBox Poruka("SETUP_ERR_GRESKA_TOKOM_SETUP") & Err.description, vbCritical, APP_NAME
End Sub

Public Sub RunSetupHealthCheck()
    On Error GoTo EH

    Dim report As String
    report = vbNullString

    InitSetupLog
    EnsureLocalConfigTable
    EnsurePoruke

    LogSetup "INFO", "RunSetupHealthCheck started"

    report = report & CheckRuntimeEnvironment()
    report = report & CheckCoreFoldersExist()
    report = report & CheckPdfToTextExists()
    report = report & CheckGoogleOAuthConfig()
    report = report & CheckSEFConfigForSetup()
    report = report & CheckRequiredTablesForSetup()
    report = report & CheckRequiredColumnsForSetup()
    ' Zivi link desktop<->server (advisory; NIJE u SetupNewPC zelenom gate-u da offline
    ' ne obara setup). Reuse: GetAccessToken/DriveListFolder/Monitor_Test.
    report = report & CheckServerLink()

    SetLocalConfigValue "APP_LAST_HEALTHCHECK_AT", Format$(Now, "yyyy-mm-dd hh:nn:ss"), "Poslednji health-check"

    If Len(report) = 0 Then
        LogSetup "OK", "Health-check passed"
        MsgBox Poruka("SETUP_MSG_HEALTH_CHECK_PROSAO"), vbInformation, APP_NAME
    Else
        LogSetup "WARN", report
        MsgBox Poruka("SETUP_MSG_HEALTH_CHECK_NASAO") & vbCrLf & vbCrLf & report, _
               vbExclamation, APP_NAME
    End If

    Exit Sub

EH:
    LogSetup "ERROR", "RunSetupHealthCheck failed: " & Err.description
    MsgBox Poruka("SETUP_ERR_GRESKA_TOKOM_HEALTH") & Err.description, vbCritical, APP_NAME
End Sub

Public Function IsSetupHealthy() As Boolean
    On Error GoTo EH

    If UCase$(Trim$(GetLocalConfigValue("APP_SETUP_COMPLETED", ""))) <> "DA" Then
        IsSetupHealthy = False
        Exit Function
    End If

    If Len(CheckCoreFoldersExist()) > 0 Then
        IsSetupHealthy = False
        Exit Function
    End If

    If Len(CheckRequiredTablesForSetup()) > 0 Then
        IsSetupHealthy = False
        Exit Function
    End If

    IsSetupHealthy = True
    Exit Function

EH:
    IsSetupHealthy = False
End Function

' ------------------------------------------------------------
' Desktop-only prekidac.
'
' UKLJUCEN (CLOUD_SYNC_ENABLED=NO): aplikacija radi 100% lokalno -- nema
' Google/PWA saobracaja, a SetupNewPC ne trazi Google kredencijale.
' Flag se cuva u tblSEFConfig i cita ga IsCloudSyncEnabled() (modConfig) --
' isti koji vec gasi runtime sync (modBrojevi, modStanicaLock).
'
' Alt+F8 -> EnableDesktopOnlyMode  (pa ponovo SetupNewPC).
' ------------------------------------------------------------
Public Sub EnableDesktopOnlyMode()
    On Error GoTo EH

    SetConfigValue CFG_KEY_CLOUD_SYNC_ENABLED, "NO"

    InitSetupLog
    LogSetup "OK", "Desktop-only mod UKLJUCEN (CLOUD_SYNC_ENABLED=NO)"

    MsgBox "Desktop-only mod je ukljucen." & vbCrLf & vbCrLf & _
           Poruka("SETUP_MSG_APLIKACIJA_RADI_LOKALNO") & vbCrLf & _
           "Pokrenite SetupNewPC ponovo.", _
           vbInformation, APP_NAME
    Exit Sub

EH:
    MsgBox "Ne mogu da ukljucim desktop-only mod: " & Err.description & vbCrLf & _
           "(Proverite da postoji tabela " & TBL_SEF_CONFIG & ".)", _
           vbCritical, APP_NAME
End Sub

Public Sub EnableCloudSyncMode()
    On Error GoTo EH

    SetConfigValue CFG_KEY_CLOUD_SYNC_ENABLED, "YES"

    InitSetupLog
    LogSetup "OK", "Cloud/PWA sync UKLJUCEN (CLOUD_SYNC_ENABLED=YES)"

    MsgBox "Cloud/PWA sync je ukljucen." & vbCrLf & vbCrLf & _
           "SetupNewPC ce ponovo traziti Google kredencijale u tblSEFConfig.", _
           vbInformation, APP_NAME
    Exit Sub

EH:
    MsgBox "Ne mogu da promenim mod: " & Err.description, vbCritical, APP_NAME
End Sub

Public Sub SetupBankFoldersInteractive()
    On Error GoTo EH

    InitSetupLog
    EnsureLocalConfigTable

    Dim inboxPath As String
    Dim processedPath As String
    Dim errorPath As String
    Dim driveSourcePath As String

    ' Drive izvor: folder koji Google Drive for Desktop sinhronizuje lokalno
    ' (npr. H:\My Drive\AgriX_C001_PROD\00_Inbox\01_Bank). Puller (modBankaImport)
    ' cita odavde. NE pravimo ga (pravi ga Drive sync) -- samo pamtimo putanju.
    ' Cancel = preskoci (puller je onda iskljucen dok se rucno ne podesi).
    driveSourcePath = PickFolder("Izaberi Drive izvor (sinhronizovan 00_Inbox\01_Bank); Cancel da preskocis")
    If Len(Trim$(driveSourcePath)) > 0 Then
        SetLocalConfigValue "BANKA_DRIVE_SOURCE_PATH", driveSourcePath, _
                            "Drive izvor (sinhronizovan folder) za povlacenje izvoda"
    End If

    inboxPath = PickFolder("Izaberi folder za nove bankarske izvode / Inbox")
    If Len(Trim$(inboxPath)) = 0 Then Exit Sub

    processedPath = PickFolder("Izaberi folder za obradene izvode / Processed")
    If Len(Trim$(processedPath)) = 0 Then
        processedPath = inboxPath & "\Processed"
    End If

    errorPath = PickFolder("Izaberi folder za neispravne izvode / Error")
    If Len(Trim$(errorPath)) = 0 Then
        errorPath = inboxPath & "\Error"
    End If

    EnsureFolder inboxPath
    EnsureFolder processedPath
    EnsureFolder errorPath

    SetLocalConfigValue "BANKA_INBOX_PATH", inboxPath, "Folder za nove bankarske izvode"
    SetLocalConfigValue "BANKA_PROCESSED_PATH", processedPath, "Folder za obradene bankarske izvode"
    SetLocalConfigValue "BANKA_ERROR_PATH", errorPath, Poruka("SETUP_MSG_FOLDER_BANKARSKE_IZVODE")

    MsgBox Poruka("SETUP_MSG_BANKARSKI_FOLDERI_PODESENI"), vbInformation, APP_NAME
    Exit Sub

EH:
    LogSetup "ERROR", "SetupBankFoldersInteractive failed: " & Err.description
    MsgBox Poruka("SETUP_ERR_GRESKA_PRI_PODESAVANJU") & Err.description, vbCritical, APP_NAME
End Sub

' RUCNO (Alt+F8): podesi pdftotext.exe (Poppler) za banka import.
' Logika (isto sto ResolvePdfToTextExePath koristi):
'   1) Ako poppler VEC stoji pored xlsm-a (<xlsm>\Tools\poppler\Library\bin\
'      pdftotext.exe) -> upisi PRAZAN PDFTOTEXT_EXE_PATH. Prazna vrednost znaci
'      "koristi auto-default relativan na xlsm" (GetLocalConfigValue na prazno vraca
'      default), pa putanja UVEK prati radnu svesku ako se paket premesti.
'   2) Inace -> pitaj operatera za folder sa pdftotext.exe (trazi i u uobicajenim
'      podfolderima) i upisi apsolutnu putanju.
Public Sub SetupPopplerInteractive()
    On Error GoTo EH

    InitSetupLog
    EnsureLocalConfigTable

    ' 1) Auto pored xlsm-a?
    Dim autoExe As String
    autoExe = Trim$(ThisWorkbook.path) & "\" & APP_PDFTOTEXT_RELATIVE_EXE_PATH
    If Dir$(autoExe) <> "" Then
        SetLocalConfigValue "PDFTOTEXT_EXE_PATH", "", _
            "Prazno = auto (Tools\poppler pored xlsm-a; putanja prati radnu svesku)"
        MsgBox "Poppler je pronadjen pored radne sveske:" & vbCrLf & autoExe & vbCrLf & vbCrLf & _
               "Podeseno na AUTOMATSKI rezim (putanja se racuna relativno na xlsm), " & _
               "pa nastavlja da radi i ako premestis ceo paket.", _
               vbInformation, APP_NAME
        Exit Sub
    End If

    ' 2) Nije pored xlsm-a -> picker.
    Dim picked As String
    picked = PickFolder("Izaberi folder sa pdftotext.exe (npr. ...\Tools\poppler\Library\bin ili poppler root)")
    If Len(Trim$(picked)) = 0 Then Exit Sub

    Dim exePath As String
    exePath = FindPdfToTextExe(picked)

    If Len(exePath) = 0 Then
        MsgBox "pdftotext.exe nije pronadjen u izabranom folderu ni u uobicajenim " & _
               "podfolderima (\Library\bin, \bin, \poppler\Library\bin, \poppler\bin)." & vbCrLf & vbCrLf & _
               "Izaberi tacan folder gde je pdftotext.exe (ili raspakovani poppler root).", _
               vbExclamation, APP_NAME
        Exit Sub
    End If

    SetLocalConfigValue "PDFTOTEXT_EXE_PATH", exePath, "pdftotext.exe (rucno izabran)"
    MsgBox "Poppler podesen:" & vbCrLf & exePath & vbCrLf & vbCrLf & _
           "NAPOMENA: ovo je APSOLUTNA putanja i ostaje ista ako premestis xlsm. " & _
           "Za putanju koja prati radnu svesku, drzi Tools\poppler pored xlsm-a pa " & _
           "ponovo pokreni ovu komandu (upisace se automatski rezim).", _
           vbInformation, APP_NAME
    Exit Sub

EH:
    LogSetup "ERROR", "SetupPopplerInteractive failed: " & Err.description
    MsgBox Poruka("SETUP_ERR_GRESKA_PRI_PODESAVANJU") & Err.description, vbCritical, APP_NAME
End Sub

' Trazi pdftotext.exe u zadatom folderu i uobicajenim podfolderima (poppler layout-i).
' Vraca punu putanju do exe-a ili "" ako nije nadjen. Bez wildcard Dir$ petlje (Dir$ je
' stateful) -- fiksni skup kandidata.
Private Function FindPdfToTextExe(ByVal root As String) As String
    root = Trim$(root)
    If Len(root) = 0 Then Exit Function
    If Right$(root, 1) = "\" Then root = Left$(root, Len(root) - 1)

    Dim subs As Variant
    subs = Array("", "\Library\bin", "\bin", "\poppler\Library\bin", "\poppler\bin")

    Dim i As Long, cand As String
    For i = LBound(subs) To UBound(subs)
        cand = root & subs(i) & "\pdftotext.exe"
        If Dir$(cand) <> "" Then
            FindPdfToTextExe = cand
            Exit Function
        End If
    Next i
End Function

' ============================================================
' PUBLIC CONFIG HELPERS
' ============================================================
' These are intentionally Public so other modules can use tblLocalConfig,
' for example modBankaImport.

Public Function GetLocalConfigValue(ByVal keyName As String, _
                                    Optional ByVal defaultValue As String = "") As String
    On Error GoTo EH

    Dim lo As ListObject
    Dim colKey As Long
    Dim colValue As Long
    Dim r As Long
    Dim currentKey As String

    Set lo = FindListObject(TBL_LOCAL_CONFIG)
    If lo Is Nothing Then
        GetLocalConfigValue = defaultValue
        Exit Function
    End If

    colKey = GetLocalTableColumnIndex(lo, CFG_KEY)
    colValue = GetLocalTableColumnIndex(lo, CFG_VALUE)

    If colKey = 0 Or colValue = 0 Then
        GetLocalConfigValue = defaultValue
        Exit Function
    End If

    If lo.DataBodyRange Is Nothing Then
        GetLocalConfigValue = defaultValue
        Exit Function
    End If

    For r = 1 To lo.DataBodyRange.rows.count
        currentKey = Trim$(CStr(lo.DataBodyRange.cells(r, colKey).value))

        If UCase$(currentKey) = UCase$(Trim$(keyName)) Then
            If Len(Trim$(CStr(lo.DataBodyRange.cells(r, colValue).value))) = 0 Then
                GetLocalConfigValue = defaultValue
            Else
                GetLocalConfigValue = Trim$(CStr(lo.DataBodyRange.cells(r, colValue).value))
            End If

            Exit Function
        End If
    Next r

    GetLocalConfigValue = defaultValue
    Exit Function

EH:
    GetLocalConfigValue = defaultValue
End Function

Public Sub SetLocalConfigValue(ByVal keyName As String, _
                               ByVal valueText As String, _
                               Optional ByVal opisText As String = "")
    On Error GoTo EH

    Dim lo As ListObject
    Dim colKey As Long
    Dim colValue As Long
    Dim colOpis As Long
    Dim r As Long
    Dim currentKey As String
    Dim lr As ListRow

    EnsureLocalConfigTable

    Set lo = FindListObject(TBL_LOCAL_CONFIG)

    If lo Is Nothing Then
        Err.Raise vbObjectError + 9200, "SetLocalConfigValue", _
                  "Ne postoji tabela: " & TBL_LOCAL_CONFIG
    End If

    colKey = GetLocalTableColumnIndex(lo, CFG_KEY)
    colValue = GetLocalTableColumnIndex(lo, CFG_VALUE)
    colOpis = GetLocalTableColumnIndex(lo, CFG_DESC)

    If colKey = 0 Or colValue = 0 Then
        Err.Raise vbObjectError + 9201, "SetLocalConfigValue", _
                  TBL_LOCAL_CONFIG & " mora imati kolone Kljuc i Vrednost."
    End If

    If Not lo.DataBodyRange Is Nothing Then
        For r = 1 To lo.DataBodyRange.rows.count
            currentKey = Trim$(CStr(lo.DataBodyRange.cells(r, colKey).value))

            If UCase$(currentKey) = UCase$(Trim$(keyName)) Then
                lo.DataBodyRange.cells(r, colValue).value = valueText

                If colOpis > 0 And Len(Trim$(opisText)) > 0 Then
                    lo.DataBodyRange.cells(r, colOpis).value = opisText
                End If

                Exit Sub
            End If
        Next r
    End If

    Set lr = lo.ListRows.Add

    lr.Range.cells(1, colKey).value = keyName
    lr.Range.cells(1, colValue).value = valueText

    If colOpis > 0 Then
        lr.Range.cells(1, colOpis).value = opisText
    End If

    Exit Sub

EH:
    Err.Raise Err.Number, "SetLocalConfigValue", Err.description
End Sub

' GetGoogleConfigValue (citac tblConfig-a) je uklonjen: Google/PWA config se cita iz
' tblSEFConfig preko GetConfigValue -- isto kao runtime (modGoogleAuth). Dva citaca na
' dve razlicite tabele su bila uzrok laznog "Nedostaje GOOGLE_... u tblConfig".

' ============================================================
' CHECKS
' ============================================================

Private Function CheckRuntimeEnvironment() As String
    Dim msg As String

#If VBA7 Then
    LogSetup "OK", "VBA7 detected"
#Else
    msg = msg & "- Office/VBA nije VBA7. Proveriti compatibility." & vbCrLf
#End If

    If val(Application.Version) < 16 Then
        msg = msg & "- Excel verzija je starija od preporucene." & vbCrLf
    End If

    CheckRuntimeEnvironment = msg
End Function

Private Function EnsureAppFolders() As String
    Dim msg As String
    Dim rootPath As String

    rootPath = GetLocalConfigValue("APP_ROOT_PATH", "")

    If Len(Trim$(rootPath)) = 0 Then
        rootPath = GetDefaultRootPath()
        SetLocalConfigValue "APP_ROOT_PATH", rootPath, "Root folder aplikacije na ovom racunaru"
    End If

    EnsureFolder rootPath
    EnsureFolder GetLocalConfigWithDefault("APP_BACKUP_PATH", rootPath & "\Backups", "Folder za backup fajlove")
    EnsureFolder GetLocalConfigWithDefault("APP_LOG_PATH", rootPath & "\Logs", "Folder za log fajlove")
    EnsureFolder GetLocalConfigWithDefault("APP_JOURNAL_PATH", rootPath & "\Journal", "Folder za journal/recovery fajlove")
    EnsureFolder GetLocalConfigWithDefault("APP_EXPORT_PATH", rootPath & "\Export", "Folder za eksport fajlove")
    EnsureFolder GetLocalConfigWithDefault("APP_TEMP_PATH", rootPath & "\Temp", "Privremeni folder")
    EnsureFolder GetLocalConfigWithDefault("APP_SECRETS_PATH", rootPath & "\Secrets", "Folder za lokalne tajne/token fajlove ako se koriste")

    ' PDF dokument folderi (pored radne sveske) - da postoje i pre prvog PDF-a
    EnsureAllDocFolders

    EnsureAppFolders = msg
End Function

Private Function SetupBankFolders() As String
    Dim msg As String
    Dim rootPath As String
    Dim bankRoot As String
    Dim inboxPath As String
    Dim processedPath As String
    Dim errorPath As String

    rootPath = GetLocalConfigValue("APP_ROOT_PATH", GetDefaultRootPath())
    bankRoot = rootPath & "\Bank_Izvodi"

    inboxPath = GetLocalConfigWithDefault("BANKA_INBOX_PATH", bankRoot & "\Inbox", "Folder za nove bankarske izvode")
    processedPath = GetLocalConfigWithDefault("BANKA_PROCESSED_PATH", bankRoot & "\Processed", "Folder za obradene bankarske izvode")
    errorPath = GetLocalConfigWithDefault("BANKA_ERROR_PATH", bankRoot & "\Error", Poruka("SETUP_MSG_FOLDER_BANKARSKE_IZVODE"))

    EnsureFolder inboxPath
    EnsureFolder processedPath
    EnsureFolder errorPath

    If Dir$(inboxPath, vbDirectory) = "" Then msg = msg & "- Banka Inbox folder nije dostupan." & vbCrLf
    If Dir$(processedPath, vbDirectory) = "" Then msg = msg & "- Banka Processed folder nije dostupan." & vbCrLf
    If Dir$(errorPath, vbDirectory) = "" Then msg = msg & "- Banka Error folder nije dostupan." & vbCrLf

    If Trim$(GetLocalConfigValue("BANKA_AUTO_IMPORT_ON_START", "")) = "" Then
        SetLocalConfigValue "BANKA_AUTO_IMPORT_ON_START", "NE", "Da li se bankarski izvodi automatski uvoze pri startu"
    End If

    If Trim$(GetLocalConfigValue("BANKA_ALLOWED_EXTENSIONS", "")) = "" Then
        SetLocalConfigValue "BANKA_ALLOWED_EXTENSIONS", "pdf", "Dozvoljene ekstenzije za bankarske izvode"
    End If

    SetupBankFolders = msg
End Function

Private Function CheckGoogleOAuthConfig() As String
    Dim msg As String

    ' Desktop-only deploy (CLOUD_SYNC_ENABLED=NO): aplikacija ne dira Google/PWA,
    ' pa setup ne trazi Google kredencijale. Isti flag gasi runtime saobracaj
    ' (modBrojevi.MaxSeqFromGoogleSheet, modStanicaLock).
    If Not IsCloudSyncEnabled() Then
        LogSetup "INFO", "Desktop-only mod: preskacem Google/PWA proveru (CLOUD_SYNC_ENABLED=NO)"
        CheckGoogleOAuthConfig = vbNullString
        Exit Function
    End If

    Dim googleClientID As String
    Dim googleClientSecret As String
    Dim googlePwaFolderID As String

    ' Google/PWA kredencijali zive u tblSEFConfig -- runtime ih cita preko
    ' GetConfigValue (modGoogleAuth, modBrojevi, modGoogleSyncOrchestrator), a
    ' Podesavanja ih tamo i pisu. Zato i setup provera cita tblSEFConfig (ranije je
    ' greskom gledala tblConfig preko GetGoogleConfigValue -> lazan "Nedostaje").
    googleClientID = Trim$(GetConfigValue("GOOGLE_CLIENT_ID"))
    googleClientSecret = Trim$(GetConfigValue("GOOGLE_CLIENT_SECRET"))
    googlePwaFolderID = Trim$(GetConfigValue("GOOGLE_PWA_FOLDER_ID"))

    If Len(googleClientID) > 0 _
       And Len(googleClientSecret) > 0 _
       And Len(googlePwaFolderID) > 0 Then

        LogSetup "OK", "Google OAuth/PWA config found in tblSEFConfig"
        CheckGoogleOAuthConfig = vbNullString
        Exit Function
    End If

    If Len(googleClientID) = 0 Then
        msg = msg & "- Nedostaje GOOGLE_CLIENT_ID u tblSEFConfig." & vbCrLf
    End If

    If Len(googleClientSecret) = 0 Then
        msg = msg & "- Nedostaje GOOGLE_CLIENT_SECRET u tblSEFConfig." & vbCrLf
    End If

    If Len(googlePwaFolderID) = 0 Then
        msg = msg & "- Nedostaje GOOGLE_PWA_FOLDER_ID u tblSEFConfig." & vbCrLf
    End If

    CheckGoogleOAuthConfig = msg
End Function

Private Function CheckSEFConfigForSetup() As String
    Dim msg As String

    If GetTable(TBL_SEF_CONFIG) Is Nothing Then
        CheckSEFConfigForSetup = "- Nedostaje tabela: " & TBL_SEF_CONFIG & vbCrLf
        Exit Function
    End If

    Dim baseUrl As String, apiKey As String, sefEnv As String
    baseUrl = Trim$(GetConfigValue("SEF_BASE_URL"))
    apiKey = Trim$(GetConfigValue("SEF_API_KEY"))
    sefEnv = Trim$(GetConfigValue("SEF_ENV"))

    ' SEF (e-faktura) je opcion. Ako NIJEDNO SEF polje nije popunjeno -> SEF se ne
    ' koristi na ovoj instalaciji -> preskoci proveru (bez laznog upozorenja). Cim je
    ' bar jedno polje popunjeno, SEF je delimicno podesen pa prijavi sta nedostaje.
    ' Runtime i dalje cvrsto zaustavlja slanje ka SEF-u ako fali kljuc
    ' (modSEFClient/modSEFValidator), pa preskakanje ovde nista ne razbija.
    If Len(baseUrl) = 0 And Len(apiKey) = 0 And Len(sefEnv) = 0 Then
        LogSetup "INFO", "SEF se ne koristi (sva SEF polja prazna) -- preskacem SEF proveru"
        CheckSEFConfigForSetup = vbNullString
        Exit Function
    End If

    If Len(baseUrl) = 0 Then msg = msg & Poruka("SETUP_MSG_SEF_BASE_URL") & vbCrLf
    If Len(apiKey) = 0 Then msg = msg & Poruka("SETUP_MSG_SEF_API_KEY") & vbCrLf
    If Len(sefEnv) = 0 Then msg = msg & Poruka("SETUP_MSG_SEF_ENV_NIJE") & vbCrLf

    CheckSEFConfigForSetup = msg
End Function

Private Function CheckRequiredTablesForSetup() As String
    Dim msg As String
    Dim tbls As Variant
    Dim i As Long

    tbls = Array( _
        TBL_CONFIG, TBL_SEF_CONFIG, TBL_BANKA_IMPORT, TBL_PARTNER_MAP, _
        TBL_KOOPERANTI, TBL_STANICE, TBL_VOZACI, TBL_KUPCI, TBL_KULTURE, _
        TBL_OTKUP, TBL_OTPREMNICA, TBL_ZBIRNA, TBL_PRIJEMNICA, _
        TBL_FAKTURE, TBL_FAKTURA_STAVKE, TBL_NOVAC, TBL_AMBALAZA, TBL_PORUKE _
    )

    For i = LBound(tbls) To UBound(tbls)
        If GetTable(CStr(tbls(i))) Is Nothing Then
            msg = msg & "- Nedostaje tabela: " & CStr(tbls(i)) & vbCrLf
        End If
    Next i

    If FindListObject(TBL_LOCAL_CONFIG) Is Nothing Then
        msg = msg & "- Nedostaje tabela: " & TBL_LOCAL_CONFIG & vbCrLf
    End If

    CheckRequiredTablesForSetup = msg
End Function

Private Function CheckRequiredColumnsForSetup() As String
    Dim msg As String
    Dim loLocal As ListObject

    On Error GoTo EH

    ' tblConfig: legacy tabela (validiraj kolone samo ako postoji; Google config je
    ' danas u tblSEFConfig)
    If Not GetTable(TBL_CONFIG) Is Nothing Then
        If GetColumnIndex(TBL_CONFIG, CFG_KEY) = 0 Then
            msg = msg & "- tblConfig nema kolonu Kljuc." & vbCrLf
        End If

        If GetColumnIndex(TBL_CONFIG, CFG_VALUE) = 0 Then
            msg = msg & "- tblConfig nema kolonu Vrednost." & vbCrLf
        End If
    End If

    ' tblLocalConfig: local setup config table
    Set loLocal = FindListObject(TBL_LOCAL_CONFIG)

    If loLocal Is Nothing Then
        msg = msg & "- Nedostaje " & TBL_LOCAL_CONFIG & "." & vbCrLf
    Else
        If GetLocalTableColumnIndex(loLocal, CFG_KEY) = 0 Then
            msg = msg & "- tblLocalConfig nema kolonu Kljuc." & vbCrLf
        End If

        If GetLocalTableColumnIndex(loLocal, CFG_VALUE) = 0 Then
            msg = msg & "- tblLocalConfig nema kolonu Vrednost." & vbCrLf
        End If
    End If

    If Not GetTable(TBL_BANKA_IMPORT) Is Nothing Then
        RequireColumnIndex TBL_BANKA_IMPORT, COL_BIM_ID, "modSetup.CheckRequiredColumnsForSetup"
        RequireColumnIndex TBL_BANKA_IMPORT, COL_BIM_BROJ_DOKUMENTA, "modSetup.CheckRequiredColumnsForSetup"
        RequireColumnIndex TBL_BANKA_IMPORT, COL_BIM_DATUM_TRANSAKCIJE, "modSetup.CheckRequiredColumnsForSetup"
        RequireColumnIndex TBL_BANKA_IMPORT, COL_BIM_PARTNER, "modSetup.CheckRequiredColumnsForSetup"
        RequireColumnIndex TBL_BANKA_IMPORT, COL_BIM_UPLATA, "modSetup.CheckRequiredColumnsForSetup"
        RequireColumnIndex TBL_BANKA_IMPORT, COL_BIM_ISPLATA, "modSetup.CheckRequiredColumnsForSetup"
        RequireColumnIndex TBL_BANKA_IMPORT, COL_BIM_IZVOR_FAJL, "modSetup.CheckRequiredColumnsForSetup"
    End If

    If Not GetTable(TBL_FAKTURE) Is Nothing Then
        RequireColumnIndex TBL_FAKTURE, COL_FAK_ID, "modSetup.CheckRequiredColumnsForSetup"
        RequireColumnIndex TBL_FAKTURE, COL_FAK_BROJ, "modSetup.CheckRequiredColumnsForSetup"
        RequireColumnIndex TBL_FAKTURE, COL_FAK_KUPAC, "modSetup.CheckRequiredColumnsForSetup"
    End If

    CheckRequiredColumnsForSetup = msg
    Exit Function

EH:
    CheckRequiredColumnsForSetup = "- Schema check failed: " & Err.description & vbCrLf
End Function

Private Function CheckCoreFoldersExist() As String
    Dim msg As String

    msg = msg & CheckFolderExists("APP_BACKUP_PATH", "Backup folder")
    msg = msg & CheckFolderExists("APP_LOG_PATH", "Log folder")
    msg = msg & CheckFolderExists("APP_JOURNAL_PATH", "Journal folder")
    msg = msg & CheckFolderExists("BANKA_INBOX_PATH", "Banka Inbox folder")
    msg = msg & CheckFolderExists("BANKA_PROCESSED_PATH", "Banka Processed folder")
    msg = msg & CheckFolderExists("BANKA_ERROR_PATH", "Banka Error folder")

    CheckCoreFoldersExist = msg
End Function

Private Function CheckFolderExists(ByVal configKey As String, ByVal labelText As String) As String
    Dim p As String
    p = Trim$(GetLocalConfigValue(configKey, ""))

    If p = "" Then
        CheckFolderExists = "- " & labelText & Poruka("SETUP_MSG_NIJE_PODESEN") & configKey & vbCrLf
    ElseIf Dir$(p, vbDirectory) = "" Then
        CheckFolderExists = "- " & labelText & " ne postoji: " & p & vbCrLf
    End If
End Function

' Poppler (pdftotext.exe) je neophodan za banka import. Reuse razresavanja putanje
' iz parsera (ResolvePdfToTextExePath) da ne dupliramo logiku; on raise-uje kad
' nije podesen ili fajl ne postoji, pa ovde samo hvatamo poruku za report.
Private Function CheckPdfToTextExists() As String
    On Error GoTo EH
    Dim p As String
    p = ResolvePdfToTextExePath()
    Exit Function
EH:
    CheckPdfToTextExists = "- pdftotext.exe (banka import): " & Err.description & _
                           " (Alt+F8 -> SetupPopplerInteractive da izaberes folder, ili " & _
                           "stavi Tools\poppler pored xlsm-a)" & vbCrLf
End Function

' Zivi link desktop<->server. Advisory (reuse postojecih primitiva; bez novog HTTP-a):
'   - Google Drive/Sheets: GetAccessToken (OAuth zivi token) + DriveListFolder(PWA folder)
'   - GAS monitoring:       Monitor_Test() ako je MONITORING_ENDPOINT podesen
'   - GAS license:          prisustvo LICENSE_ENDPOINT (zivi check radi startup gate)
'   - Banka Drive folder:   lokalni BANKA_DRIVE_SOURCE_PATH (Drive-for-Desktop)
' Desktop-only (CLOUD_SYNC_ENABLED=NO): preskace server deo, ali i dalje proverava
' lokalni banka folder. Sve fail-soft -- nikad ne baca, samo skuplja poruke.
Private Function CheckServerLink() As String
    Dim msg As String

    If Not IsCloudSyncEnabled() Then
        LogSetup "INFO", "Desktop-only: preskacem server link (proveravam samo lokalni banka folder)"
        CheckServerLink = CheckBankaDriveFolderLink()
        Exit Function
    End If

    ' 1) Google Drive/Sheets -- OAuth zivi token je pouzdan signal "spojen na Google".
    On Error Resume Next
    If Not IsGoogleAuthConfigured() Then
        msg = msg & "- Google OAuth nije konfigurisan (GOOGLE_* / refresh token). Pokreni RunGoogleAuthSetup." & vbCrLf
    Else
        Dim token As String
        token = ""
        token = GetAccessToken()
        If Len(Trim$(token)) = 0 Then
            msg = msg & "- Google link: ne mogu da dobijem access token (re-auth: RunGoogleAuthSetup)." & vbCrLf
        Else
            ' Folder read-proba (soft): 0 stavki = prazan ILI nedostupan -> WARN, ne FAIL.
            Dim pwaFolder As String
            pwaFolder = Trim$(GetConfigValue("GOOGLE_PWA_FOLDER_ID"))
            If Len(pwaFolder) > 0 Then
                Dim d As Object
                Set d = DriveListFolder(pwaFolder)
                If d Is Nothing Then
                    msg = msg & "- Google PWA folder: nedostupan (DriveListFolder). Proveri GOOGLE_PWA_FOLDER_ID / deljenje." & vbCrLf
                ElseIf d.count = 0 Then
                    msg = msg & "- Google PWA folder: 0 stavki (prazan ili nedostupan). Detaljnije: RunProductionHealthCheck." & vbCrLf
                End If
            End If
        End If
    End If
    On Error GoTo 0

    ' 2) GAS monitoring endpoint (samo ako je podesen; inace Monitor_Test lazno pada).
    If Len(Trim$(GetConfigValue("MONITORING_ENDPOINT"))) > 0 Then
        Dim monOk As Boolean
        monOk = False
        On Error Resume Next
        monOk = Monitor_Test()
        On Error GoTo 0
        If Not monOk Then
            msg = msg & "- GAS monitoring endpoint ne odgovara (MONITORING_ENDPOINT / MONITORING_SECRET)." & vbCrLf
        End If
    End If

    ' 3) GAS license endpoint -- samo ako se licenciranje koristi. Zivi check radi
    '    startup gate (AccessGateOrQuit); ovde proveravamo da je endpoint uopste zadat.
    If UCase$(Trim$(GetConfigValue("LICENSE_ENABLED"))) = "YES" Then
        If Len(Trim$(GetConfigValue("LICENSE_ENDPOINT"))) = 0 _
           And Len(Trim$(GetConfigValue("MONITORING_ENDPOINT"))) = 0 Then
            msg = msg & "- Licenciranje ukljuceno, a LICENSE_ENDPOINT (ni MONITORING_ENDPOINT fallback) nije podesen." & vbCrLf
        End If
    End If

    ' 4) Banka Drive folder (lokalni Drive-for-Desktop izvor).
    msg = msg & CheckBankaDriveFolderLink()

    CheckServerLink = msg
End Function

' Lokalni banka Drive izvor (BANKA_DRIVE_SOURCE_PATH). Ako nije podesen -> banka
' Drive-pull se ne koristi (prazno, bez poruke). Ako je podesen a folder ne postoji
' -> Drive-for-Desktop ne sinhronizuje ili je putanja pogresna.
Private Function CheckBankaDriveFolderLink() As String
    Dim p As String
    p = Trim$(GetLocalConfigValue("BANKA_DRIVE_SOURCE_PATH", ""))
    If Len(p) = 0 Then Exit Function

    On Error Resume Next
    If Dir$(p, vbDirectory) = "" Then
        CheckBankaDriveFolderLink = "- Banka Drive izvor nedostupan (BANKA_DRIVE_SOURCE_PATH): " & p & vbCrLf
    End If
    On Error GoTo 0
End Function

' RUCNO (Alt+F8): proveri samo zivi link desktop<->server i prikazi rezultat.
Public Sub TestServerLink()
    On Error GoTo EH
    InitSetupLog

    Dim report As String
    report = CheckServerLink()

    If Len(report) = 0 Then
        MsgBox "Server link OK (Google / GAS / banka Drive folder dostupni ili nisu u upotrebi).", _
               vbInformation, APP_NAME
    Else
        MsgBox "Server link -- stavke za proveru:" & vbCrLf & vbCrLf & report, _
               vbExclamation, APP_NAME
    End If
    Exit Sub

EH:
    LogSetup "ERROR", "TestServerLink failed: " & Err.description
    MsgBox "Greska pri proveri server linka: " & Err.description, vbCritical, APP_NAME
End Sub

' ============================================================
' LOCAL CONFIG TABLE CREATION
' ============================================================

Private Sub EnsureLocalConfigTable()
    On Error GoTo EH

    Dim lo As ListObject
    Set lo = FindListObject(TBL_LOCAL_CONFIG)

    If Not lo Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Set ws = GetOrCreateWorksheet(LOCAL_CONFIG_SHEET)

    If ws Is Nothing Then
        Err.Raise vbObjectError + 9300, "EnsureLocalConfigTable", _
                  "Ne mogu da kreiram sheet: " & LOCAL_CONFIG_SHEET
    End If

    ws.Range("A1").value = CFG_KEY
    ws.Range("B1").value = CFG_VALUE
    ws.Range("C1").value = CFG_DESC

    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C1"), , xlYes)
    lo.name = TBL_LOCAL_CONFIG

    ws.columns("A:C").AutoFit

    LogSetup "OK", "Created " & TBL_LOCAL_CONFIG
    Exit Sub

EH:
    LogSetup "ERROR", "EnsureLocalConfigTable failed: " & Err.description
    Err.Raise Err.Number, "EnsureLocalConfigTable", Err.description
End Sub

' ============================================================
' Paletni list (Phase 2) -- jednokratni schema setup.
' Idempotentno: kreira nedostajuce tabele i kolonu na tblKulture.
' Pokrenuti JEDNOM na master workbook-u (Alt+F8 -> EnsurePaletniListSchema).
' Reuse: FindListObject / GetOrCreateWorksheet / LogSetup (gore).
' ============================================================
Public Sub EnsurePaletniListSchema()
    On Error GoTo EH

    EnsureDataTable TBL_TIP_PALETE, "TipPalete", _
        Array(COL_TPAL_TIP, COL_TPAL_TEZINA)

    EnsureDataTable TBL_TIP_AMBALAZE, "TipAmbalaze", _
        Array(COL_TAMB_TIP, COL_TAMB_TEZINA)

    EnsureDataTable TBL_PALETA, "Palete", _
        Array(COL_PAL_ID, COL_PAL_BROJ, COL_PAL_GODINA, COL_PAL_DATUM, _
              COL_PAL_VRSTA, COL_PAL_SORTA, COL_PAL_KLASA, COL_PAL_TIP_AMBALAZE, _
              COL_PAL_TIP_PALETE, COL_PAL_KAPACITET, COL_PAL_BR_GAJBICA, _
              COL_PAL_NETO, COL_PAL_AMBALAZA, COL_PAL_PALETA_KG, COL_PAL_BRUTO, _
              COL_PAL_STATUS, COL_PAL_PRERADJENO, COL_PAL_CREATED, COL_STORNIRANO, _
              COL_PAL_ISTORIJA)

    EnsureDataTable TBL_PALETA_STAVKA, "PaleteStavke", _
        Array(COL_PALS_ID, COL_PALS_PALETA_ID, COL_PALS_PRIJEMNICA_ID, _
              COL_PALS_BROJ_PRIJ, COL_PALS_BROJ_ZBIRNE, COL_PALS_KLASA, _
              COL_PALS_VRSTA, COL_PALS_SORTA, COL_PALS_BR_GAJBICA, _
              COL_PALS_NETO, COL_PALS_AMBALAZA, COL_PALS_CREATED, COL_STORNIRANO)

    EnsureDataTable TBL_PRERADA, "Prerada", _
        Array(COL_PRE_ID, COL_PRE_BROJ, COL_PRE_GODINA, COL_PRE_DATUM, _
              COL_PRE_NETO_ULAZ, COL_PRE_NETO_IZLAZ, COL_PRE_KUTIJE, COL_PRE_KESE, _
              COL_PRE_TEZINA_PALETE, COL_PRE_BRUTO, COL_PRE_AMBALAZA, _
              COL_PRE_TIP_KUTIJE, COL_PRE_TIP_KESE, COL_PRE_TIP_GP, _
              COL_PRE_NAPOMENA, COL_PRE_CREATED, COL_STORNIRANO)

    EnsureDataTable TBL_PRERADA_STAVKA, "PreradaStavke", _
        Array(COL_PRES_ID, COL_PRES_PRERADA_ID, COL_PRES_PALETA_ID, _
              COL_PRES_BROJ_PALETE, COL_PRES_NETO, COL_PRES_CREATED, COL_STORNIRANO)

    EnsureDataTable TBL_KUTIJE, "Kutije", _
        Array(COL_KUT_TIP, COL_KUT_TEZINA, "Aktivan")

    EnsureDataTable TBL_KESE, "Kese", _
        Array(COL_KES_TIP, COL_KES_TEZINA, "Aktivan")

    EnsureDataTable TBL_VRSTA_GP, "VrstaGotProizvoda", _
        Array(COL_VGP_TIP, "Aktivan")

    EnsureColumnOnTable TBL_KULTURE, COL_KUL_GAJBICA_PALETA
    EnsureColumnOnTable TBL_PALETA, COL_PAL_ISTORIJA   ' vidljivi audit trag (relabel/detach/adjust)

    ' GP grana (fakturisanje gotove robe): stavka fakture moze da nosi
    ' preradu, a prerada svoju fakturisanost -- isti obrazac kao
    ' prijemnica. Idempotentno, kao sve ovde.
    EnsureColumnOnTable TBL_FAKTURA_STAVKE, COL_FS_PRERADA_ID
    EnsureColumnOnTable TBL_FAKTURA_STAVKE, COL_FS_BROJ_PRERADE
    EnsureColumnOnTable TBL_PRERADA, COL_PRE_FAKTURISANO
    EnsureColumnOnTable TBL_PRERADA, COL_PRE_FAKTURA_ID

    EnsureCenovnikSchema

    EnsurePaletaSablon
    EnsurePreradaSablon

    LogSetup "OK", "EnsurePaletniListSchema done"
    MsgBox "Paletni list: seme su kreirane/proverene." & vbCrLf & vbCrLf & _
           "Popunite: tblTipAmbalaze (12/1, 6/1 -> kg), tblTipPalete (tip -> kg)," & vbCrLf & _
           "i kolonu GajbicaPoPaleti u tblKulture (malina = 240)." & vbCrLf & _
           ChrW(352) & "ifarnici tblKutije/tblKese (tip -> kg) i tblVrstaGotovihProizvoda" & vbCrLf & _
           "se popunjavaju preko Mati" & ChrW(269) & "nih podataka." & vbCrLf & _
           "Cenovnik (tblCenovnik) se popunjava preko Mati" & ChrW(269) & "nih podataka.", _
           vbInformation, APP_NAME
    Exit Sub

EH:
    LogSetup "ERROR", "EnsurePaletniListSchema failed: " & Err.description
    MsgBox "Gre" & ChrW(353) & "ka u EnsurePaletniListSchema: " & Err.description, vbCritical, APP_NAME
End Sub

' ============================================================
' Cenovnik (cene po proizvodu) -- jednokratni schema setup.
' Idempotentno: kreira tblCenovnik ako ne postoji, inace dopuni kolone.
' Pokrenuti na master workbook-u (deo je EnsurePaletniListSchema, a moze
' i samostalno: Alt+F8 -> EnsureCenovnikSchema).
' ============================================================
Public Sub EnsureCenovnikSchema()
    On Error GoTo EH

    EnsureDataTable TBL_CENOVNIK, "Cenovnik", _
        Array(COL_CEN_ID, COL_CEN_DATUM, COL_CEN_VRSTA, COL_CEN_SORTA, _
              COL_CEN_KLASA, COL_CEN_CENA, COL_CEN_CREATED, COL_STORNIRANO)

    LogSetup "OK", "EnsureCenovnikSchema done"
    Exit Sub

EH:
    LogSetup "ERROR", "EnsureCenovnikSchema failed: " & Err.description
    Err.Raise Err.Number, "EnsureCenovnikSchema", Err.description
End Sub

' ============================================================
' Storno / correction context (tblStornoVeze) -- persistentni zapis svake
' storno/ispravke (staro -> novo trag, mod, status, recovery). Idempotentno.
' Silent core se poziva iz EnsureRuntimeSchema (self-heal na svaki start), pa
' tabela nastane automatski posle self-update-a KODA -- bez rucnog Alt+F8.
' Alt+F8 -> EnsureStornoVezeSchema (sa MsgBox potvrdom).
' ============================================================
Public Sub EnsureStornoVezeSchemaCore()
    EnsureDataTable TBL_STORNO_VEZE, "StornoVeze", _
        Array(COL_SV_ID, COL_SV_MODE, COL_SV_STATUS, _
              COL_SV_OLD_DOCTYPE, COL_SV_OLD_DOCID, COL_SV_OLD_BROJ, _
              COL_SV_NEW_DOCTYPE, COL_SV_NEW_DOCID, COL_SV_NEW_BROJ, _
              COL_SV_PARENT_DOCTYPE, COL_SV_PARENT_DOCID, COL_SV_PARENT_BROJ, _
              COL_SV_CREATED_AT, COL_SV_CREATED_BY, COL_SV_COMPLETED_AT, _
              COL_SV_MESSAGE, COL_SV_NEEDS_RECOVERY, COL_SV_RECOVERY_ACTION)
End Sub

' Storno operacioni zurnal (append-only cell-level trag za lossless undo).
Public Sub EnsureStornoZurnalSchemaCore()
    EnsureDataTable TBL_STORNO_ZURNAL, "StornoZurnal", _
        Array(COL_SZ_ID, COL_SZ_OP_ID, COL_SZ_TS, COL_SZ_DOCTYPE, COL_SZ_BROJ, _
              COL_SZ_TABELA, COL_SZ_ROWID, COL_SZ_KOLONA, COL_SZ_STARA, COL_SZ_NOVA)
    ' Aditivno za postojece instalacije (v2.24 rani build bez NovaVrednost).
    EnsureColumnOnTable TBL_STORNO_ZURNAL, COL_SZ_NOVA
End Sub

Public Sub EnsureStornoVezeSchema()
    On Error GoTo EH

    EnsureStornoVezeSchemaCore

    LogSetup "OK", "EnsureStornoVezeSchema done"
    MsgBox "Storno/ispravka kontekst (tblStornoVeze) je kreiran/proveren." & vbCrLf & vbCrLf & _
           "Ovde se belezi svaka storno/ispravka: staro -> novo, mod, status i " & vbCrLf & _
           "recovery flag. Pregled: Dokumenta -> Osiroceni dokumenti.", _
           vbInformation, APP_NAME
    Exit Sub

EH:
    LogSetup "ERROR", "EnsureStornoVezeSchema failed: " & Err.description
    MsgBox "Greska u EnsureStornoVezeSchema: " & Err.description, vbCritical, APP_NAME
End Sub

' ============================================================
' Audit kolone (timestamp + userstamp) na svim glavnim tabelama.
' Idempotentno. Alt+F8 -> EnsureAuditColumns.
' Posle ovoga modDataAccess.AppendRow/UpdateCell automatski upisuju
' CreatedAt/CreatedBy (na unos) i ModifiedAt/ModifiedBy (na svaku izmenu) -
' vidi se KO je i KADA radio. Userstamp: prijavljeni korisnik (AUTH) ili Windows nalog.
' ============================================================
Public Sub EnsureAuditColumns()
    On Error GoTo EH

    Dim n As Long
    n = EnsureAuditColumnsCore()

    MsgBox "Audit kolone postavljene na " & n & Poruka("AUD_MSG_KOLONE_SUFIX"), _
           vbInformation, APP_NAME
    Exit Sub

EH:
    LogSetup "ERROR", "EnsureAuditColumns failed: " & Err.description
    MsgBox "Greska u EnsureAuditColumns: " & Err.description, vbCritical, APP_NAME
End Sub

' Silent worker (bez MsgBox-a) -- vraca broj obradjenih tabela. Reuse iz
' migracije (modMigracija) gde ne zelimo popup usred toka kopiranja.
Public Function EnsureAuditColumnsCore() As Long
    InitSetupLog

    Dim tbls As Variant
    tbls = AuditableTables()

    Dim i As Long, n As Long
    For i = LBound(tbls) To UBound(tbls)
        If Not GetTable(CStr(tbls(i))) Is Nothing Then
            EnsureColumnOnTable CStr(tbls(i)), COL_AUDIT_CREATED_AT
            EnsureColumnOnTable CStr(tbls(i)), COL_AUDIT_CREATED_BY
            EnsureColumnOnTable CStr(tbls(i)), COL_AUDIT_MODIFIED_AT
            EnsureColumnOnTable CStr(tbls(i)), COL_AUDIT_MODIFIED_BY
            n = n + 1
        End If
    Next i

    LogSetup "OK", "EnsureAuditColumns done (" & n & " tabela)"
    EnsureAuditColumnsCore = n
End Function

' Spisak tabela koje dobijaju audit kolone (master + transakcione + ledger).
' Config/log/report tabele se namerno preskacu.
Private Function AuditableTables() As Variant
    AuditableTables = Array( _
        TBL_KOOPERANTI, TBL_STANICE, TBL_VOZACI, TBL_KUPCI, TBL_KULTURE, _
        TBL_OTKUP, TBL_OTPREMNICA, TBL_ZBIRNA, TBL_PRIJEMNICA, _
        TBL_FAKTURE, TBL_FAKTURA_STAVKE, TBL_NOVAC, TBL_AMBALAZA, _
        TBL_PARCELE, TBL_ARTIKLI, TBL_MAGACIN, TBL_CENOVNIK, TBL_KORISNICI, _
        TBL_PALETA, TBL_PALETA_STAVKA, TBL_PRERADA, TBL_PRERADA_STAVKA, _
        TBL_TIP_AMBALAZE, TBL_TIP_PALETE, TBL_PARTNER_MAP, TBL_BANKA_IMPORT)
End Function

' ============================================================
' Poruke (resource table) -- idempotent.
' Creates tblPoruke on hidden sheet sPoruke and seeds strings.
' Alt+F8 -> EnsurePoruke.
' ============================================================
Public Sub EnsurePoruke()
    On Error GoTo EH

    EnsureDataTable TBL_PORUKE, PORUKE_SHEET, Array(COL_POR_KLJUC, COL_POR_TEKST)

    Dim lo As ListObject
    Set lo = FindListObject(TBL_PORUKE)

    If Not lo Is Nothing Then
        Dim ws As Worksheet
        Set ws = lo.Parent
        If ws.Visible <> xlSheetVeryHidden Then ws.Visible = xlSheetVeryHidden
        modPoruke.UpsertPoruke lo
    End If

    LogSetup "OK", "EnsurePoruke done"
    Exit Sub

EH:
    LogSetup "ERROR", "EnsurePoruke failed: " & Err.description
    Err.Raise Err.Number, "EnsurePoruke", Err.description
End Sub

' ============================================================
' Runtime schema self-heal -- MsgBox-free, idempotentno. Poziva se iz
' modMain.InitApp na SVAKOM startu (kao EnsurePoruke), pa kolone dodate kroz
' self-update KODA nastanu automatski posle restarta -- bez rucnog Alt+F8.
' EnsureColumnOnTable je no-op kad kolona postoji (cena posle prvog heal-a
' zanemarljiva). Ovde idu SAMO sigurne, idempotentne izmene bez backfill-a.
' ============================================================
Public Sub EnsureRuntimeSchema()
    On Error Resume Next
    ' Pragovi proseka neto kg po gajbici (otkup: upozorenje/blokada).
    EnsureColumnOnTable TBL_KULTURE, COL_KUL_PRAG_PROSEK_UPOZ
    EnsureColumnOnTable TBL_KULTURE, COL_KUL_PRAG_PROSEK_BLOK

    ' GP fakturisanje (v6-ui-189, R4 revizije #248): klijent koji dobije
    ' nov kod self-update-om mora dobiti i kolone -- inace CreateFakturaGP
    ' pada na RequireColumnIndex cim operater proba novu funkciju.
    ' Append-only i idempotentno, bas klasa izmena za runtime self-heal.
    EnsureColumnOnTable TBL_PRERADA, COL_PRE_FAKTURISANO
    EnsureColumnOnTable TBL_PRERADA, COL_PRE_FAKTURA_ID
    EnsureColumnOnTable TBL_FAKTURA_STAVKE, COL_FS_PRERADA_ID
    EnsureColumnOnTable TBL_FAKTURA_STAVKE, COL_FS_BROJ_PRERADE
    SetColumnNumberFormat TBL_KULTURE, COL_KUL_PRAG_PROSEK_UPOZ, "0.00"
    SetColumnNumberFormat TBL_KULTURE, COL_KUL_PRAG_PROSEK_BLOK, "0.00"

    ' Format kolona (schema-drift: reinstall/self-update vrati kolone na General ->
    ' E-notacija/tarabe na dokumentima). Idempotentno, tera se na SVAKI start.
    ' BPG je identifikator (dug broj), ne racunska vrednost -> Text ("@").
    SetColumnNumberFormat TBL_KOOPERANTI, COL_KOOP_BPG, "@"
    ' Prerada: tezine su Double (samo prikaz) -> fiksni decimalni format spreci General/E.
    SetColumnNumberFormat TBL_PRERADA, COL_PRE_TEZINA_PALETE, "0.00"
    SetColumnNumberFormat TBL_PRERADA, COL_PRE_BRUTO, "0.00"
    SetColumnNumberFormat TBL_PRERADA, COL_PRE_AMBALAZA, "0.00"

    ' Vidljivi audit trag na paleti (relabel/detach/adjust). Deo je i punog
    ' EnsurePaletniListSchema, ali se dodaje ovde da nastane AUTOMATSKI posle
    ' self-update-a KODA -- bez rucnog Alt+F8 koraka. EnsureColumnOnTable je no-op
    ' kad kolona postoji.
    EnsureColumnOnTable TBL_PALETA, COL_PAL_ISTORIJA

    ' Centralni storno/correction context (tblStornoVeze). Idempotentno; nastane
    ' automatski posle self-update-a KODA. EnsureDataTable je no-op kad postoji.
    EnsureStornoVezeSchemaCore

    ' Storno operacioni zurnal (tblStornoZurnal) za lossless "Vrati storno".
    ' Idempotentno; nastane automatski posle self-update-a KODA.
    EnsureStornoZurnalSchemaCore

    ' Sledljivost ispravki (ADR-0002 / Faza 7). Idempotentno; nastane automatski
    ' posle self-update-a KODA (bez rucnog Alt+F8).
    EnsureSledljivostSchema
End Sub

' ============================================================
' Sledljivost ispravki (ADR-0002, append-only dokument model; Faza 7 korak 1).
' Dodaje DELJENE trace-kolone na dokument-tabele lanca. SAMO schema (dodavanje
' kolona) -- citanje/upis dolaze u kasnijim koracima Faze 7.
' Idempotentno (EnsureColumnOnTable = no-op kad kolona postoji), pa je jeftino na
' svakom startu i nastane automatski posle self-update-a KODA.
' Bez backfill-a (skup na svakom startu): PRAZAN IzdatoStatus se tumaci kao IZDATO
' (konzervativno -> teski korektivni put) u read-logici kasnijih koraka.
' ============================================================
' Blanket "On Error Resume Next" je ovde bio najgori gard koji postoji: prvi pad
' je cutke preskakao OSTATAK posla, bez ijedne linije u logu. Posledica je bila
' sveska bez kolone GeneracijaID -- a upis dokumenta koji na nju racuna padao je
' satima kasnije, sa porukom koja o uzroku ne kaze nista.
'
' Sada svaka kolona ide kroz svoj gard: pad jedne se ZAPISE i ne zaustavlja
' ostale. Otpornost je ista, trag postoji.
Public Sub EnsureSledljivostSchema()
    Dim tbls As Variant
    tbls = Array(TBL_OTKUP, TBL_OTPREMNICA, TBL_ZBIRNA, TBL_PRIJEMNICA, TBL_FAKTURE, TBL_NOVAC)
    Dim i As Long
    For i = LBound(tbls) To UBound(tbls)
        Dim t As String: t = CStr(tbls(i))
        EnsureKolonaSaTragom t, COL_TRACE_ISPRAVKA_OD
        EnsureKolonaSaTragom t, COL_TRACE_ZAMENJEN_SA
        EnsureKolonaSaTragom t, COL_TRACE_CORRECTION_ID
        EnsureKolonaSaTragom t, COL_TRACE_IZDATO_STATUS
        ' Generacija upisa (Klasa I + II iz istog Multi_TX poziva dele vrednost).
        ' Prefill iz stornirane bira poslednju generaciju po ovoj koloni.
        EnsureKolonaSaTragom t, COL_GENERACIJA_ID
    Next i
    ' Faza 7 korak 5: denorm poslovni kljuc bloka -> otpremnica (stabilan kroz re-verziju).
    EnsureKolonaSaTragom TBL_OTKUP, COL_OTK_BROJ_OTPREMNICE
End Sub

' Jedna kolona, sa tragom. Pad se zapise i NE zaustavlja ostale kolone.
Private Sub EnsureKolonaSaTragom(ByVal tbl As String, ByVal col As String)
    On Error GoTo EH
    EnsureColumnOnTable tbl, col
    Exit Sub
EH:
    LogError "modSetup.EnsureSledljivostSchema", _
             "Kolona '" & col & "' na '" & tbl & "' nije obezbedjena: " & _
             Err.description, Err.Number
End Sub

' ============================================================
' Backfill BrojOtpremnice na tblOtkup iz OtpremnicaID (Faza 7 korak 5). Alt+F8,
' jednokratno; idempotentno (samo prazne). Efikasno: dict OtpremnicaID->Broj u
' jednom prolazu. NE ide u EnsureRuntimeSchema (skup po startu).
' ============================================================
Public Sub BackfillOtkupBrojOtpremnice()
    On Error GoTo EH
    EnsureColumnOnTable TBL_OTKUP, COL_OTK_BROJ_OTPREMNICE
    Dim od As Variant: od = GetTableData(TBL_OTPREMNICA)
    Dim map As Object: Set map = CreateObject("Scripting.Dictionary")
    If IsArray(od) Then
        Dim cOId As Long: cOId = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_ID)
        Dim cOBr As Long: cOBr = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ)
        If cOId > 0 And cOBr > 0 Then
            Dim r As Long
            For r = 1 To UBound(od, 1)
                Dim oid As String: oid = Trim$(CStr(od(r, cOId)))
                If Len(oid) > 0 And Not map.Exists(oid) Then map(oid) = Trim$(CStr(od(r, cOBr)))
            Next r
        End If
    End If
    Dim kd As Variant: kd = GetTableData(TBL_OTKUP)
    Dim n As Long: n = 0
    If IsArray(kd) Then
        Dim cKOtp As Long: cKOtp = GetColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID)
        Dim cKBr As Long: cKBr = GetColumnIndex(TBL_OTKUP, COL_OTK_BROJ_OTPREMNICE)
        If cKOtp > 0 And cKBr > 0 Then
            Dim i As Long
            For i = 1 To UBound(kd, 1)
                If Len(Trim$(CStr(kd(i, cKBr)))) = 0 Then          ' samo prazne (idempotentno)
                    Dim oid2 As String: oid2 = Trim$(CStr(kd(i, cKOtp)))
                    If Len(oid2) > 0 And map.Exists(oid2) Then
                        UpdateCell TBL_OTKUP, i, COL_OTK_BROJ_OTPREMNICE, map(oid2)
                        n = n + 1
                    End If
                End If
            Next i
        End If
    End If
    MsgBox "Backfill BrojOtpremnice: popunjeno " & n & " otkupnih redova.", vbInformation, APP_NAME
    Exit Sub
EH:
    LogErr "modSetup.BackfillOtkupBrojOtpremnice"
End Sub

' ============================================================
' Dorade (soft-delete + tip ambalaze po kulturi + hladnjaca + decimalna
' kolicina) -- jednokratni schema setup. Idempotentno.
' Alt+F8 -> EnsureDoradeSchema.
' ============================================================
Public Sub EnsureDoradeSchema()
    On Error GoTo EH

    ' #1 soft-delete: kolona Aktivan na sifarnicima koji je nemaju (+ backfill).
    EnsureAktivanColumn TBL_KULTURE
    EnsureAktivanColumn TBL_TIP_AMBALAZE
    EnsureAktivanColumn TBL_TIP_PALETE
    EnsureAktivanColumn TBL_KUTIJE
    EnsureAktivanColumn TBL_KESE
    EnsureAktivanColumn TBL_VRSTA_GP

    ' #6: podrazumevani tip ambalaze po kulturi.
    EnsureColumnOnTable TBL_KULTURE, COL_KUL_TIP_AMBALAZE

    ' Pragovi proseka neto kg po gajbici (otkup: upozorenje/blokada).
    ' Deljeno sa startup self-heal-om (modMain.InitApp -> EnsureRuntimeSchema),
    ' pa kolone nastanu i automatski posle self-update-a KODA (bez rucnog koraka).
    EnsureRuntimeSchema

    ' #3: flag hladnjaca na stanici (+ backfill "Ne").
    EnsureColumnOnTable TBL_STANICE, COL_STA_JE_HLADNJACA
    BackfillColumn TBL_STANICE, COL_STA_JE_HLADNJACA, "Ne"

    ' #5: decimalni format kolicine (vrednost je vec Double; samo prikaz).
    SetColumnNumberFormat TBL_OTKUP, COL_OTK_KOLICINA, "0.00"
    SetColumnNumberFormat TBL_OTPREMNICA, COL_OTP_KOLICINA, "0.00"
    SetColumnNumberFormat TBL_PRIJEMNICA, COL_PRJ_KOLICINA, "0.00"
    SetColumnNumberFormat TBL_ZBIRNA, COL_ZBR_KOLICINA, "0.00"

    ' Izdata ambalaza (OM->kooperant uz otkup) -> kolona na tblOtkup za otkupni list.
    EnsureColumnOnTable TBL_OTKUP, COL_OTK_KOL_AMB_IZDATA
    BackfillColumn TBL_OTKUP, COL_OTK_KOL_AMB_IZDATA, "0"

    ' Vreme snimanja otkupa (Now() pri upisu) -> za otkupni list.
    EnsureColumnOnTable TBL_OTKUP, COL_OTK_VREME_UNOSA
    SetColumnNumberFormat TBL_OTKUP, COL_OTK_VREME_UNOSA, "dd.mm.yyyy hh:nn"

    ' Bruto tezina (kad kupac unosi bruto -> sistem cuva neto u Kolicina, bruto ovde).
    ' Prazno = unet neto (bruto == neto); ne backfill-uje se (modPrint tretira >0 kao prisutno).
    EnsureColumnOnTable TBL_OTKUP, COL_OTK_BRUTO
    SetColumnNumberFormat TBL_OTKUP, COL_OTK_BRUTO, "0.00"
    EnsureColumnOnTable TBL_PRIJEMNICA, COL_PRJ_BRUTO
    SetColumnNumberFormat TBL_PRIJEMNICA, COL_PRJ_BRUTO, "0.00"
    EnsureColumnOnTable TBL_OTPREMNICA, COL_OTP_BRUTO
    SetColumnNumberFormat TBL_OTPREMNICA, COL_OTP_BRUTO, "0.00"

    LogSetup "OK", "EnsureDoradeSchema done"
    MsgBox "Dorade: seme su proverene/kreirane." & vbCrLf & vbCrLf & _
           "- Aktivan: Kulture/Ambala" & ChrW(382) & "a/Palete (postojeci -> Aktivan)" & vbCrLf & _
           "- TipAmbalaze: tblKulture (podrazumevani po kulturi)" & vbCrLf & _
           "- JeHladnjaca: tblStanice (Da/Ne)" & vbCrLf & _
           "- Decimalni format kolicine (0.00)" & vbCrLf & _
           "- KolAmbIzdata: tblOtkup (izdata ambala" & ChrW(382) & "a OM->kooperant)" & vbCrLf & _
           "- VremeUnosa: tblOtkup (vreme snimanja otkupa)" & vbCrLf & _
           "- BrutoKg: tblOtkup/tblPrijemnica (bruto unos -> cuva neto)" & vbCrLf & _
           "- PragProsekUpoz/PragProsekBlok: tblKulture (prosek neto kg po gajbici)", vbInformation, APP_NAME
    Exit Sub

EH:
    LogSetup "ERROR", "EnsureDoradeSchema failed: " & Err.description
    MsgBox "Gre" & ChrW(353) & "ka u EnsureDoradeSchema: " & Err.description, vbCritical, APP_NAME
End Sub

' Dodaj kolonu Aktivan ako fali i postavi "Aktivan" na sve prazne (backfill).
Private Sub EnsureAktivanColumn(ByVal tblName As String)
    EnsureColumnOnTable tblName, "Aktivan"
    BackfillColumn tblName, "Aktivan", STATUS_AKTIVAN
End Sub

' Postavi vrednost u koloni za sve redove gde je prazno (ne gazi postojece).
Private Sub BackfillColumn(ByVal tblName As String, ByVal colName As String, ByVal val As String)
    On Error Resume Next
    Dim lo As ListObject
    Set lo = FindListObject(tblName)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Dim col As ListColumn
    Set col = lo.ListColumns(colName)
    If col Is Nothing Then Exit Sub

    Dim cell As Range
    For Each cell In col.DataBodyRange.cells
        If Trim$(CStr(cell.value)) = "" Then cell.value = val
    Next cell
End Sub

' Postavi NumberFormat kolone (no-op ako tabela/kolona ne postoji).
Private Sub SetColumnNumberFormat(ByVal tblName As String, ByVal colName As String, ByVal fmt As String)
    On Error Resume Next
    Dim lo As ListObject
    Set lo = FindListObject(tblName)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Dim col As ListColumn
    Set col = lo.ListColumns(colName)
    If col Is Nothing Then Exit Sub
    col.DataBodyRange.NumberFormat = fmt
End Sub

' ============================================================
' Korisnici (Faza 1) - admin + prava po oblasti. Idempotentni schema setup.
' Model A: jedan red = korisnik; kolona po oblasti = "DA"/"NE". Admin = bypass.
' Reuse: EnsureDataTable / EnsureColumnOnTable (schema-drift safe).
' Pokreni JEDNOM: Alt+F8 -> EnsureKorisniciSchema (ili KreirajPrvogAdmina).
' ============================================================
Public Sub EnsureKorisniciSchema()
    On Error GoTo EH

    EnsureDataTable TBL_KORISNICI, "Korisnici", _
        Array(COL_KOR_ID, COL_KOR_USERNAME, COL_KOR_IME, COL_KOR_PIN, _
              COL_KOR_ULOGA, COL_KOR_AKTIVAN, COL_KOR_STANICA, COL_KOR_CREATED, _
              OBL_OTKUP, OBL_DOKUMENTA, OBL_AGROHEMIJA, OBL_IZVESTAJI, _
              OBL_FAKTURISANJE, OBL_BANKA, OBL_MARZA, OBL_SLEDLJIVOST, OBL_MATICNI, _
              OBL_PALETE, OBL_OTVORI_EXCEL, OBL_SYNC_PWA)

    LogSetup "OK", "EnsureKorisniciSchema done"
    Exit Sub

EH:
    LogSetup "ERROR", "EnsureKorisniciSchema failed: " & Err.description
    Err.Raise Err.Number, "EnsureKorisniciSchema", Err.description
End Sub

' Kreira prvog ADMINA (sa svim pravima). Bezbedan bootstrap protiv lockout-a:
' prijava se ukljucuje tek posle ovoga (EnableAuth proverava da admin postoji).
' Alt+F8 -> KreirajPrvogAdmina.
Public Sub KreirajPrvogAdmina()
    On Error GoTo EH

    If Not modAuth.MozeAdministraciju() Then
        MsgBox Poruka("AUTH_MSG_SAMO_ADMIN_KREIRA"), vbExclamation, APP_NAME
        Exit Sub
    End If

    EnsureKorisniciSchema

    Dim u As String, pin As String, ime As String
    u = Trim$(InputBox("Korisnicko ime za ADMINA:", APP_NAME, "admin"))
    If Len(u) = 0 Then Exit Sub
    pin = Trim$(InputBox("PIN za '" & u & "':", APP_NAME))
    If Len(pin) = 0 Then Exit Sub
    ime = Trim$(InputBox("Ime i prezime (opciono):", APP_NAME))

    ' Spreci duplikat username-a
    Dim postoji As Variant
    postoji = LookupValue(TBL_KORISNICI, COL_KOR_USERNAME, u, COL_KOR_ID)
    If Not IsEmpty(postoji) Then
        If Len(Trim$(CStr(postoji))) > 0 Then
            MsgBox "Korisnik '" & u & Poruka("AUTH_MSG_VEC_POSTOJI"), vbExclamation, APP_NAME
            Exit Sub
        End If
    End If

    Dim newId As String
    newId = GetNextID(TBL_KORISNICI, COL_KOR_ID, "KOR-")

    Dim colCount As Long
    colCount = GetTable(TBL_KORISNICI).ListColumns.count

    Dim rowData() As Variant
    ReDim rowData(1 To colCount)

    Dim idx As Long
    idx = AppendRow(TBL_KORISNICI, rowData)
    If idx = 0 Then Err.Raise vbObjectError + 9400, "KreirajPrvogAdmina", "AppendRow nije uspeo."

    ' Upis po imenu (drift-safe, CLAUDE.md)
    UpdateCell TBL_KORISNICI, idx, COL_KOR_ID, newId
    UpdateCell TBL_KORISNICI, idx, COL_KOR_USERNAME, u
    UpdateCell TBL_KORISNICI, idx, COL_KOR_IME, ime
    UpdateCell TBL_KORISNICI, idx, COL_KOR_PIN, modAuth.PreparePin(pin)
    UpdateCell TBL_KORISNICI, idx, COL_KOR_ULOGA, ULOGA_ADMIN
    UpdateCell TBL_KORISNICI, idx, COL_KOR_AKTIVAN, "DA"
    UpdateCell TBL_KORISNICI, idx, COL_KOR_CREATED, Format$(Now, "yyyy-mm-dd hh:nn:ss")

    ' Admin svejedno bypass-uje; setujemo DA radi preglednosti u gridu.
    Dim obl As Variant
    For Each obl In modAuth.OblastiList()
        UpdateCell TBL_KORISNICI, idx, CStr(obl), "DA"
    Next obl

    MsgBox "Admin '" & u & Poruka("AUTH_MSG_ADMIN_KREIRAN"), _
           vbInformation, APP_NAME
    Exit Sub

EH:
    LogSetup "ERROR", "KreirajPrvogAdmina failed: " & Err.description
    MsgBox "Greska: " & Err.description, vbCritical, APP_NAME
End Sub

' Ukljuci prijavu korisnika (AUTH_ENABLED=YES). Mirror EnableDesktopOnlyMode.
' Bezbednost: ne dozvoljava ukljucivanje bez bar jednog AKTIVNOG admina.
' Alt+F8 -> EnableAuth.
Public Sub EnableAuth()
    On Error GoTo EH

    If Not modAuth.MozeAdministraciju() Then
        MsgBox Poruka("AUTH_MSG_SAMO_ADMIN_PRIJAVA"), vbExclamation, APP_NAME
        Exit Sub
    End If

    If Not PostojiAktivanAdmin() Then
        MsgBox Poruka("AUTH_MSG_NEMA_ADMINA"), vbExclamation, APP_NAME
        Exit Sub
    End If

    SetConfigValue CFG_KEY_AUTH_ENABLED, "YES"
    InitSetupLog
    LogSetup "OK", "AUTH_ENABLED = YES"
    MsgBox Poruka("AUTH_MSG_PRIJAVA_UKLJUCENA"), vbInformation, APP_NAME
    Exit Sub

EH:
    MsgBox Poruka("AUTH_ERR_NE_MOGU_UKLJUCITI") & Err.description, vbCritical, APP_NAME
End Sub

' Iskljuci prijavu (AUTH_ENABLED=NO). Alt+F8 -> DisableAuth.
Public Sub DisableAuth()
    On Error GoTo EH

    If Not modAuth.MozeAdministraciju() Then
        MsgBox Poruka("AUTH_MSG_SAMO_ADMIN_ISKLJUCI"), vbExclamation, APP_NAME
        Exit Sub
    End If

    SetConfigValue CFG_KEY_AUTH_ENABLED, "NO"
    InitSetupLog
    LogSetup "OK", "AUTH_ENABLED = NO"
    MsgBox Poruka("AUTH_MSG_PRIJAVA_ISKLJUCENA"), vbInformation, APP_NAME
    Exit Sub

EH:
    MsgBox "Greska: " & Err.description, vbCritical, APP_NAME
End Sub

' Ukljuci PIN hash (opt-in). Self-test SHA pre ukljucenja. Alt+F8 -> EnablePinHash.
Public Sub EnablePinHash()
    On Error GoTo EH

    If Not modAuth.MozeAdministraciju() Then
        MsgBox Poruka("AUTH_MSG_SAMO_ADMIN_PINHASH"), vbExclamation, APP_NAME
        Exit Sub
    End If

    If StrComp(modAuth.Sha256Hex("abc"), _
               "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad", _
               vbTextCompare) <> 0 Then
        MsgBox Poruka("AUTH_ERR_SHA_NE_RADI"), vbCritical, APP_NAME
        Exit Sub
    End If

    SetConfigValue CFG_KEY_PIN_HASH_ENABLED, "YES"
    InitSetupLog
    LogSetup "OK", "PIN_HASH_ENABLED = YES"
    MsgBox Poruka("AUTH_MSG_PINHASH_UKLJUCEN"), vbInformation, APP_NAME
    Exit Sub

EH:
    MsgBox "Greska: " & Err.description, vbCritical, APP_NAME
End Sub

' Iskljuci PIN hash. Vec hesirani PIN-ovi i dalje rade (prijava ih prepoznaje).
Public Sub DisablePinHash()
    On Error GoTo EH

    If Not modAuth.MozeAdministraciju() Then
        MsgBox Poruka("AUTH_MSG_SAMO_ADMIN_PINHASH"), vbExclamation, APP_NAME
        Exit Sub
    End If

    SetConfigValue CFG_KEY_PIN_HASH_ENABLED, "NO"
    InitSetupLog
    LogSetup "OK", "PIN_HASH_ENABLED = NO"
    MsgBox Poruka("AUTH_MSG_PINHASH_ISKLJUCEN"), vbInformation, APP_NAME
    Exit Sub

EH:
    MsgBox "Greska: " & Err.description, vbCritical, APP_NAME
End Sub

Private Function PostojiAktivanAdmin() As Boolean
    On Error GoTo EH

    Dim lo As ListObject
    Set lo = GetTable(TBL_KORISNICI)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim colUloga As Long, colAkt As Long, r As Long
    colUloga = GetColumnIndex(TBL_KORISNICI, COL_KOR_ULOGA)
    colAkt = GetColumnIndex(TBL_KORISNICI, COL_KOR_AKTIVAN)
    If colUloga = 0 Then Exit Function

    For r = 1 To lo.DataBodyRange.rows.count
        If StrComp(Trim$(CStr(lo.DataBodyRange.cells(r, colUloga).value)), ULOGA_ADMIN, vbTextCompare) = 0 Then
            If colAkt = 0 Then
                PostojiAktivanAdmin = True
                Exit Function
            ElseIf UCase$(Trim$(CStr(lo.DataBodyRange.cells(r, colAkt).value))) <> "NE" Then
                PostojiAktivanAdmin = True
                Exit Function
            End If
        End If
    Next r

    Exit Function

EH:
    PostojiAktivanAdmin = False
End Function

' Dijagnostika: prikazi STVARNE nazive kolona neke tabele.
' Pokreni: Alt+F8 -> DebugKoloneTabele -> unesi npr. tblStanice.
' Reuse: GetTableHeaders (modDataAccess).
Public Sub DebugKoloneTabele()
    Dim t As String
    t = InputBox("Naziv tabele (npr. tblStanice):", APP_NAME, "tblStanice")
    If Trim$(t) = "" Then Exit Sub

    Dim h As Variant
    h = GetTableHeaders(t)
    If IsEmpty(h) Then
        MsgBox "Tabela '" & t & "' nije pronadjena.", vbExclamation, APP_NAME
        Exit Sub
    End If

    MsgBox t & " kolone (" & (UBound(h) - LBound(h) + 1) & "):" & vbCrLf & _
           Join(h, " | "), vbInformation, APP_NAME
End Sub

' Kreira ListObject sa zadatim zaglavljima na (novom) sheet-u. No-op ako vec postoji.
Private Sub EnsureDataTable(ByVal tblName As String, _
                            ByVal sheetName As String, _
                            ByVal headers As Variant)
    Dim lo As ListObject
    Set lo = FindListObject(tblName)
    If Not lo Is Nothing Then
        ' tabela postoji -> dopuni nedostajuce kolone (repair schema drift)
        Dim k As Long
        For k = LBound(headers) To UBound(headers)
            EnsureColumnOnTable tblName, CStr(headers(k))
        Next k
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = GetOrCreateWorksheet(sheetName)
    If ws Is Nothing Then
        Err.Raise vbObjectError + 9310, "EnsureDataTable", _
                  "Ne mogu da kreiram sheet: " & sheetName
    End If

    Dim c As Long
    For c = LBound(headers) To UBound(headers)
        ws.cells(1, c - LBound(headers) + 1).value = headers(c)
    Next c

    Dim LASTCOL As Long
    LASTCOL = UBound(headers) - LBound(headers) + 1

    Set lo = ws.ListObjects.Add(xlSrcRange, _
        ws.Range(ws.cells(1, 1), ws.cells(1, LASTCOL)), , xlYes)
    lo.name = tblName

    ws.columns.AutoFit
    LogSetup "OK", "Created " & tblName
End Sub

' Dodaje kolonu na postojecu tabelu ako je nema. No-op ako tabela ne postoji.
Private Sub EnsureColumnOnTable(ByVal tblName As String, ByVal colName As String)
    Dim lo As ListObject
    Set lo = FindListObject(tblName)
    If lo Is Nothing Then
        ' Tabele nema -- kolona se ne moze dodati. Do sada tih izlaz; sada se vidi,
        ' jer se bas na ovome gubila GeneracijaID.
        LogWarn "modSetup.EnsureColumnOnTable", _
                "Tabela '" & tblName & "' nije nadjena -- kolona '" & colName & "' preskocena."
        Exit Sub
    End If

    Dim col As ListColumn
    On Error Resume Next
    Set col = lo.ListColumns(colName)
    On Error GoTo 0

    If col Is Nothing Then
        lo.ListColumns.Add
        lo.ListColumns(lo.ListColumns.count).name = colName
        LogSetup "OK", "Added column " & colName & " to " & tblName
    End If
End Sub

Private Function GetOrCreateWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateWorksheet Is Nothing Then
        Set GetOrCreateWorksheet = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        GetOrCreateWorksheet.name = sheetName
    End If
End Function

Private Function FindListObject(ByVal tableName As String) As ListObject
    On Error GoTo EH

    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If UCase$(lo.name) = UCase$(tableName) Then
                Set FindListObject = lo
                Exit Function
            End If
        Next lo
    Next ws

    Set FindListObject = Nothing
    Exit Function

EH:
    Set FindListObject = Nothing
End Function

Private Function GetLocalTableColumnIndex(ByVal lo As ListObject, ByVal columnName As String) As Long
    On Error GoTo EH

    Dim i As Long

    For i = 1 To lo.ListColumns.count
        If UCase$(Trim$(lo.ListColumns(i).name)) = UCase$(Trim$(columnName)) Then
            GetLocalTableColumnIndex = i
            Exit Function
        End If
    Next i

    GetLocalTableColumnIndex = 0
    Exit Function

EH:
    GetLocalTableColumnIndex = 0
End Function

' ============================================================
' LOCAL CONFIG DEFAULTS
' ============================================================

Private Function GetLocalConfigWithDefault(ByVal keyName As String, _
                                           ByVal defaultValue As String, _
                                           Optional ByVal opisText As String = "") As String
    Dim currentValue As String

    currentValue = Trim$(GetLocalConfigValue(keyName, ""))

    If currentValue = "" Then
        SetLocalConfigValue keyName, defaultValue, opisText
        GetLocalConfigWithDefault = defaultValue
    Else
        GetLocalConfigWithDefault = currentValue
    End If
End Function

Private Function GetDefaultRootPath() As String
    If Len(Trim$(ThisWorkbook.path)) > 0 Then
        GetDefaultRootPath = ThisWorkbook.path
    Else
        GetDefaultRootPath = Environ$("USERPROFILE") & "\Documents\OtkupApp"
    End If
End Function

' ============================================================
' FILE/FOLDER HELPERS
' ============================================================

Private Sub EnsureFolder(ByVal folderPath As String)
    On Error GoTo EH

    If Len(Trim$(folderPath)) = 0 Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(folderPath) Then
        LogSetup "OK", "Folder exists: " & folderPath
        Exit Sub
    End If

    Dim parentPath As String
    parentPath = fso.GetParentFolderName(folderPath)

    If Len(parentPath) > 0 Then
        If Not fso.FolderExists(parentPath) Then
            EnsureFolder parentPath
        End If
    End If

    fso.CreateFolder folderPath
    LogSetup "OK", "Created folder: " & folderPath

    Exit Sub

EH:
    LogSetup "ERROR", "Cannot create folder: " & folderPath & " | " & Err.description
End Sub

' Vraca putanju do podfoldera za generisane PDF dokumente (pored radne sveske) i
' kreira ga ako fali. Npr. EnsureDocFolder(PDF_DIR_OTKUPNI). Fallback na TEMP ako
' sveska nije snimljena; ako kreiranje ne uspe, vraca root (PDF se i dalje snima).
Public Function EnsureDocFolder(ByVal category As String) As String
    On Error GoTo EH
    Dim rootPath As String
    rootPath = ThisWorkbook.path
    If Len(Trim$(rootPath)) = 0 Then rootPath = Environ$("TEMP")

    If Len(Trim$(category)) = 0 Then
        EnsureDocFolder = rootPath
        Exit Function
    End If

    Dim target As String
    target = rootPath & "\" & category
    EnsureFolder target

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(target) Then
        EnsureDocFolder = target
    Else
        EnsureDocFolder = rootPath
    End If
    Exit Function
EH:
    EnsureDocFolder = rootPath
End Function

' Kreira sve PDF dokument foldere pored radne sveske (poziva se iz setup-a da
' folderi postoje i pre prvog generisanog PDF-a).
Public Sub EnsureAllDocFolders()
    EnsureDocFolder PDF_DIR_OTKUPNI
    EnsureDocFolder PDF_DIR_PRIJEMNICE
    EnsureDocFolder PDF_DIR_OTPREMNICE
    EnsureDocFolder PDF_DIR_REVERS
    EnsureDocFolder PDF_DIR_KARTICE
    EnsureDocFolder PDF_DIR_PALETNI
    EnsureDocFolder PDF_DIR_PRERADA
    EnsureDocFolder PDF_DIR_SPECIFIKACIJE
    EnsureDocFolder PDF_DIR_IZVESTAJI
End Sub

' Public: reuse-uje ga i modPodesavanja (inline "..." browse dugmad u config editoru).
Public Function PickFolder(ByVal titleText As String) As String
    On Error GoTo EH

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    With fd
        .title = titleText
        .AllowMultiSelect = False

        If .Show = -1 Then
            PickFolder = .SelectedItems(1)
        End If
    End With

    Exit Function

EH:
    PickFolder = vbNullString
End Function

' ============================================================
' LOGGING
' ============================================================

Private Sub InitSetupLog()
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SETUP_LOG_SHEET)

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = SETUP_LOG_SHEET
        ws.Range("A1:D1").value = Array("Timestamp", "Level", "Message", "User")
        ws.rows(1).Font.Bold = True
    End If
End Sub

Private Sub LogSetup(ByVal levelText As String, ByVal messageText As String)
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SETUP_LOG_SHEET)

    If ws Is Nothing Then Exit Sub

    Dim r As Long
    r = ws.cells(ws.rows.count, 1).End(xlUp).row + 1

    ws.cells(r, 1).value = Now
    ws.cells(r, 2).value = levelText
    ws.cells(r, 3).value = Left$(messageText, 2000)
    ws.cells(r, 4).value = Environ$("Username")
End Sub

