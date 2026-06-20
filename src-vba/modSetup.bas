
Option Explicit

' ============================================================
' modSetup – OtkupApp new PC setup / workstation health-check
'
' CONFIG STRATEGY:
'
' 1) tblConfig
'    Existing Google/PWA config only:
'       Kljuc | Vrednost | Opis
'
'    Expected keys:
'       GOOGLE_CLIENT_ID
'       GOOGLE_CLIENT_SECRET
'       GOOGLE_PWA_FOLDER_ID
'
'    This module READS tblConfig but does not write local setup values to it.
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

    If Len(report) = 0 Then
        SetLocalConfigValue "APP_SETUP_COMPLETED", "DA", "Da li je ovaj racunar prošao SetupNewPC"
        SetLocalConfigValue "APP_SETUP_COMPLETED_AT", Format$(Now, "yyyy-mm-dd hh:nn:ss"), "Datum i vreme završetka setup-a"
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

        MsgBox "Setup je uspešno završen." & vbCrLf & vbCrLf & _
               "Aplikacija je spremna za ovaj racunar." & vbCrLf & _
               "Podešavanja: Matični podaci -> Podešavanja (tblSEFConfig je skriven).", _
               vbInformation, APP_NAME
    Else
        SetLocalConfigValue "APP_SETUP_COMPLETED", "NE", "Da li je ovaj racunar prošao SetupNewPC"
        SetLocalConfigValue "APP_LAST_HEALTHCHECK_AT", Format$(Now, "yyyy-mm-dd hh:nn:ss"), "Poslednji health-check"

        LogSetup "WARN", report

        MsgBox "Setup je završen, ali postoje stavke za proveru:" & _
               vbCrLf & vbCrLf & report, _
               vbExclamation, APP_NAME
    End If

    Exit Sub

EH:
    LogSetup "ERROR", "SetupNewPC failed: " & Err.Number & " - " & Err.description
    MsgBox "Greška tokom setup-a: " & Err.description, vbCritical, APP_NAME
End Sub

Public Sub RunSetupHealthCheck()
    On Error GoTo EH

    Dim report As String
    report = vbNullString

    InitSetupLog
    EnsureLocalConfigTable

    LogSetup "INFO", "RunSetupHealthCheck started"

    report = report & CheckRuntimeEnvironment()
    report = report & CheckCoreFoldersExist()
    report = report & CheckGoogleOAuthConfig()
    report = report & CheckSEFConfigForSetup()
    report = report & CheckRequiredTablesForSetup()
    report = report & CheckRequiredColumnsForSetup()

    SetLocalConfigValue "APP_LAST_HEALTHCHECK_AT", Format$(Now, "yyyy-mm-dd hh:nn:ss"), "Poslednji health-check"

    If Len(report) = 0 Then
        LogSetup "OK", "Health-check passed"
        MsgBox "Health-check je prošao. Racunar je podešen.", vbInformation, APP_NAME
    Else
        LogSetup "WARN", report
        MsgBox "Health-check je našao stavke za proveru:" & vbCrLf & vbCrLf & report, _
               vbExclamation, APP_NAME
    End If

    Exit Sub

EH:
    LogSetup "ERROR", "RunSetupHealthCheck failed: " & Err.description
    MsgBox "Greška tokom health-check-a: " & Err.description, vbCritical, APP_NAME
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
' UKLJUCEN (CLOUD_SYNC_ENABLED=NO): aplikacija radi 100% lokalno — nema
' Google/PWA saobracaja, a SetupNewPC ne trazi Google kredencijale.
' Flag se cuva u tblSEFConfig i cita ga IsCloudSyncEnabled() (modConfig) —
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
           "Aplikacija radi lokalno — Google/PWA se ne koristi niti proverava." & vbCrLf & _
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
           "SetupNewPC ce ponovo traziti Google kredencijale u tblConfig.", _
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
    SetLocalConfigValue "BANKA_ERROR_PATH", errorPath, "Folder za bankarske izvode sa greškom"

    MsgBox "Bankarski folderi su podešeni.", vbInformation, APP_NAME
    Exit Sub

EH:
    LogSetup "ERROR", "SetupBankFoldersInteractive failed: " & Err.description
    MsgBox "Greška pri podešavanju bankarskih foldera: " & Err.description, vbCritical, APP_NAME
End Sub

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

Public Function GetGoogleConfigValue(ByVal keyName As String, _
                                     Optional ByVal defaultValue As String = "") As String
    On Error GoTo EH

    Dim v As Variant

    v = LookupValue(TBL_CONFIG, CFG_KEY, keyName, CFG_VALUE)

    If isError(v) Or IsNull(v) Or IsEmpty(v) Then
        GetGoogleConfigValue = defaultValue
    ElseIf Len(Trim$(CStr(v))) = 0 Then
        GetGoogleConfigValue = defaultValue
    Else
        GetGoogleConfigValue = Trim$(CStr(v))
    End If

    Exit Function

EH:
    GetGoogleConfigValue = defaultValue
End Function

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
    errorPath = GetLocalConfigWithDefault("BANKA_ERROR_PATH", bankRoot & "\Error", "Folder za bankarske izvode sa greškom")

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

    googleClientID = Trim$(GetGoogleConfigValue("GOOGLE_CLIENT_ID", ""))
    googleClientSecret = Trim$(GetGoogleConfigValue("GOOGLE_CLIENT_SECRET", ""))
    googlePwaFolderID = Trim$(GetGoogleConfigValue("GOOGLE_PWA_FOLDER_ID", ""))

    If Len(googleClientID) > 0 _
       And Len(googleClientSecret) > 0 _
       And Len(googlePwaFolderID) > 0 Then

        LogSetup "OK", "Google OAuth/PWA config found in tblConfig"
        CheckGoogleOAuthConfig = vbNullString
        Exit Function
    End If

    If Len(googleClientID) = 0 Then
        msg = msg & "- Nedostaje GOOGLE_CLIENT_ID u tblConfig." & vbCrLf
    End If

    If Len(googleClientSecret) = 0 Then
        msg = msg & "- Nedostaje GOOGLE_CLIENT_SECRET u tblConfig." & vbCrLf
    End If

    If Len(googlePwaFolderID) = 0 Then
        msg = msg & "- Nedostaje GOOGLE_PWA_FOLDER_ID u tblConfig." & vbCrLf
    End If

    CheckGoogleOAuthConfig = msg
End Function

Private Function CheckSEFConfigForSetup() As String
    Dim msg As String

    If GetTable(TBL_SEF_CONFIG) Is Nothing Then
        CheckSEFConfigForSetup = "- Nedostaje tabela: " & TBL_SEF_CONFIG & vbCrLf
        Exit Function
    End If

    If Trim$(GetConfigValue("SEF_BASE_URL")) = "" Then
        msg = msg & "- SEF_BASE_URL nije podešen." & vbCrLf
    End If

    If Trim$(GetConfigValue("SEF_API_KEY")) = "" Then
        msg = msg & "- SEF_API_KEY nije podešen." & vbCrLf
    End If

    If Trim$(GetConfigValue("SEF_ENV")) = "" Then
        msg = msg & "- SEF_ENV nije podešen." & vbCrLf
    End If

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
        TBL_FAKTURE, TBL_FAKTURA_STAVKE, TBL_NOVAC, TBL_AMBALAZA _
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

    ' tblConfig: existing Google config table
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
        CheckFolderExists = "- " & labelText & " nije podešen: " & configKey & vbCrLf
    ElseIf Dir$(p, vbDirectory) = "" Then
        CheckFolderExists = "- " & labelText & " ne postoji: " & p & vbCrLf
    End If
End Function

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
' Paletni list (Phase 2) — jednokratni schema setup.
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
              COL_PAL_STATUS, COL_PAL_PRERADJENO, COL_PAL_CREATED, COL_STORNIRANO)

    EnsureDataTable TBL_PALETA_STAVKA, "PaleteStavke", _
        Array(COL_PALS_ID, COL_PALS_PALETA_ID, COL_PALS_PRIJEMNICA_ID, _
              COL_PALS_BROJ_PRIJ, COL_PALS_BROJ_ZBIRNE, COL_PALS_KLASA, _
              COL_PALS_VRSTA, COL_PALS_SORTA, COL_PALS_BR_GAJBICA, _
              COL_PALS_NETO, COL_PALS_AMBALAZA, COL_PALS_CREATED, COL_STORNIRANO)

    EnsureDataTable TBL_PRERADA, "Prerada", _
        Array(COL_PRE_ID, COL_PRE_BROJ, COL_PRE_GODINA, COL_PRE_DATUM, _
              COL_PRE_NETO_ULAZ, COL_PRE_NETO_IZLAZ, COL_PRE_KUTIJE, COL_PRE_KESE, _
              COL_PRE_NAPOMENA, COL_PRE_CREATED, COL_STORNIRANO)

    EnsureDataTable TBL_PRERADA_STAVKA, "PreradaStavke", _
        Array(COL_PRES_ID, COL_PRES_PRERADA_ID, COL_PRES_PALETA_ID, _
              COL_PRES_BROJ_PALETE, COL_PRES_NETO, COL_PRES_CREATED, COL_STORNIRANO)

    EnsureColumnOnTable TBL_KULTURE, COL_KUL_GAJBICA_PALETA

    EnsureCenovnikSchema

    EnsurePaletaSablon
    EnsurePreradaSablon

    LogSetup "OK", "EnsurePaletniListSchema done"
    MsgBox "Paletni list: seme su kreirane/proverene." & vbCrLf & vbCrLf & _
           "Popunite: tblTipAmbalaze (12/1, 6/1 -> kg), tblTipPalete (tip -> kg)," & vbCrLf & _
           "i kolonu GajbicaPoPaleti u tblKulture (malina = 240)." & vbCrLf & _
           "Cenovnik (tblCenovnik) se popunjava preko Maticnih podataka.", _
           vbInformation, APP_NAME
    Exit Sub

EH:
    LogSetup "ERROR", "EnsurePaletniListSchema failed: " & Err.description
    MsgBox "Greska u EnsurePaletniListSchema: " & Err.description, vbCritical, APP_NAME
End Sub

' ============================================================
' Cenovnik (cene po proizvodu) — jednokratni schema setup.
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
' Dorade (soft-delete + tip ambalaze po kulturi + hladnjaca + decimalna
' kolicina) — jednokratni schema setup. Idempotentno.
' Alt+F8 -> EnsureDoradeSchema.
' ============================================================
Public Sub EnsureDoradeSchema()
    On Error GoTo EH

    ' #1 soft-delete: kolona Aktivan na sifarnicima koji je nemaju (+ backfill).
    EnsureAktivanColumn TBL_KULTURE
    EnsureAktivanColumn TBL_TIP_AMBALAZE
    EnsureAktivanColumn TBL_TIP_PALETE

    ' #6: podrazumevani tip ambalaze po kulturi.
    EnsureColumnOnTable TBL_KULTURE, COL_KUL_TIP_AMBALAZE

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
    BackfillColumn      TBL_OTKUP, COL_OTK_KOL_AMB_IZDATA, "0"

    ' Vreme snimanja otkupa (Now() pri upisu) -> za otkupni list.
    EnsureColumnOnTable   TBL_OTKUP, COL_OTK_VREME_UNOSA
    SetColumnNumberFormat TBL_OTKUP, COL_OTK_VREME_UNOSA, "dd.mm.yyyy hh:nn"

    ' Bruto tezina (kad kupac unosi bruto -> sistem cuva neto u Kolicina, bruto ovde).
    ' Prazno = unet neto (bruto == neto); ne backfill-uje se (modPrint tretira >0 kao prisutno).
    EnsureColumnOnTable   TBL_OTKUP, COL_OTK_BRUTO
    SetColumnNumberFormat TBL_OTKUP, COL_OTK_BRUTO, "0.00"
    EnsureColumnOnTable   TBL_PRIJEMNICA, COL_PRJ_BRUTO
    SetColumnNumberFormat TBL_PRIJEMNICA, COL_PRJ_BRUTO, "0.00"

    LogSetup "OK", "EnsureDoradeSchema done"
    MsgBox "Dorade: seme su proverene/kreirane." & vbCrLf & vbCrLf & _
           "- Aktivan: Kulture/Ambalaza/Palete (postojeci -> Aktivan)" & vbCrLf & _
           "- TipAmbalaze: tblKulture (podrazumevani po kulturi)" & vbCrLf & _
           "- JeHladnjaca: tblStanice (Da/Ne)" & vbCrLf & _
           "- Decimalni format kolicine (0.00)" & vbCrLf & _
           "- KolAmbIzdata: tblOtkup (izdata ambalaza OM->kooperant)" & vbCrLf & _
           "- VremeUnosa: tblOtkup (vreme snimanja otkupa)" & vbCrLf & _
           "- BrutoKg: tblOtkup/tblPrijemnica (bruto unos -> cuva neto)", vbInformation, APP_NAME
    Exit Sub

EH:
    LogSetup "ERROR", "EnsureDoradeSchema failed: " & Err.description
    MsgBox "Greska u EnsureDoradeSchema: " & Err.description, vbCritical, APP_NAME
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
    For Each cell In col.DataBodyRange.Cells
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

    Dim lastCol As Long
    lastCol = UBound(headers) - LBound(headers) + 1

    Set lo = ws.ListObjects.Add(xlSrcRange, _
        ws.Range(ws.cells(1, 1), ws.cells(1, lastCol)), , xlYes)
    lo.name = tblName

    ws.columns.AutoFit
    LogSetup "OK", "Created " & tblName
End Sub

' Dodaje kolonu na postojecu tabelu ako je nema. No-op ako tabela ne postoji.
Private Sub EnsureColumnOnTable(ByVal tblName As String, ByVal colName As String)
    Dim lo As ListObject
    Set lo = FindListObject(tblName)
    If lo Is Nothing Then Exit Sub

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
    If Len(Trim$(ThisWorkbook.Path)) > 0 Then
        GetDefaultRootPath = ThisWorkbook.Path
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

Private Function PickFolder(ByVal titleText As String) As String
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

