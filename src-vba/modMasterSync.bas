Attribute VB_Name = "modMasterSync"
 Option Explicit

' ============================================================
' modMasterSync - Import OTK-Sheets ? tblOtkup
'
' Liest alle Google Sheets "OTK-*" aus dem PWA-Folder,
' importiert neue Zeilen (SyncStatus != "Synced?Master")
' in tblOtkup, und schreibt SyncStatus zurueck.
'
' Flow:
'   1. Liste alle OTK-* Sheets im PWA-Folder
'   2. Pro Sheet: ReadSheetData ? pruefe SyncStatus
'   3. Neue Zeilen ? Validierung ? AppendRow tblOtkup
'   4. SyncStatus ? "Synced?Master" zurueckschreiben
'
' Config-Keys:
'   GOOGLE_PWA_FOLDER_ID (bereits vorhanden)
'
' Aufruf: Button in frmMain "Uvezi otkupe iz terena"
' ============================================================
Private mLastPWAFatalSyncError As Boolean

Private Const SYNC_STATUS_PENDING As String = "Synced"
Private Const SYNC_STATUS_MASTER As String = "Synced>Master"
Private Const SYNC_STATUS_ERROR As String = "SyncError"
Private Const SYNC_STATUS_DUPLICATE As String = "Duplicate"

Private Const ERR_MASTER_SYNC_GUARD_BASE As Long = vbObjectError + 3900
Private Const MASTER_SYNC_CLIENT_RECORD_ID_COL As String = "ClientRecordID"


' Google Sheet Spaltenindizes (0-based, Header in Row 1)
Private Const GS_CLIENT_RECORD_ID As Long = 1    ' A
Private Const GS_SERVER_RECORD_ID As Long = 2    ' B
Private Const GS_CREATED_AT As Long = 3          ' C
Private Const GS_UPDATED_AT_CLIENT As Long = 4   ' D
Private Const GS_UPDATED_AT_SERVER As Long = 5   ' E
Private Const GS_SYNC_STATUS As Long = 6         ' F
Private Const GS_DEVICE_ID As Long = 7           ' G
Private Const GS_OTKUPAC_ID As Long = 8          ' H
Private Const GS_DATUM As Long = 9               ' I
Private Const GS_KOOPERANT_ID As Long = 10       ' J
Private Const GS_KOOPERANT_NAME As Long = 11     ' K
Private Const GS_VRSTA As Long = 12              ' L
Private Const GS_SORTA As Long = 13              ' M
Private Const GS_KLASA As Long = 14              ' N
Private Const GS_KOLICINA As Long = 15           ' O
Private Const GS_CENA As Long = 16               ' P
Private Const GS_TIP_AMB As Long = 17            ' Q
Private Const GS_KOL_AMB As Long = 18            ' R
Private Const GS_PARCELA_ID As Long = 19         ' S
Private Const GS_VOZAC_ID As Long = 20           ' T
Private Const GS_NAPOMENA As Long = 21           ' U
Private Const GS_RECEIVED_AT As Long = 22        ' V
Private Const GS_BROJ_DOKUMENTA As Long = 23     ' W


' VOZ Sheet Spaltenindizes (1-based, Header in Row 1)
Private Const VS_CLIENT_RECORD_ID As Long = 1   ' A
Private Const VS_SERVER_RECORD_ID As Long = 2   ' B
Private Const VS_CREATED_AT As Long = 3         ' C
Private Const VS_UPDATED_AT_CLIENT As Long = 4  ' D
Private Const VS_UPDATED_AT_SERVER As Long = 5  ' E
Private Const VS_SYNC_STATUS As Long = 6        ' F
Private Const VS_VOZAC_ID As Long = 7           ' G
Private Const VS_DATUM As Long = 8              ' H
Private Const VS_KUPAC_ID As Long = 9           ' I
Private Const VS_KUPAC_NAME As Long = 10        ' J
Private Const VS_VRSTA As Long = 11             ' K
Private Const VS_SORTA As Long = 12             ' L
Private Const VS_KOLICINA_KL_I As Long = 13     ' M
Private Const VS_KOLICINA_KL_II As Long = 14    ' N
Private Const VS_TIP_AMB As Long = 15           ' O
Private Const VS_KOL_AMB As Long = 16           ' P
Private Const VS_KLASA As Long = 17             ' Q
Private Const VS_OTKUP_RECORD_IDS As Long = 18  ' R
Private Const VS_RECEIVED_AT As Long = 19       ' S
Private Const VS_BROJ_ZBIRNE As Long = 20   ' T

' ============================================================
' PUBLIC -- Hauptfunktion
' ============================================================

Public Sub ImportOtkupFromPWA()
    Call ImportOtkupFromPWA_Core(True)
End Sub

Public Function ImportOtkupFromPWA_Core(ByVal showMessages As Boolean) As Boolean
    Dim folderID As String
    Dim sheetIDs As Collection
    Dim sheetNames As Collection
    Dim i As Long
    Dim totalImported As Long
    Dim totalSkipped As Long
    Dim totalErrors As Long
    Dim filesCount As Long

    On Error GoTo EH

    ImportOtkupFromPWA_Core = False
    mLastPWAFatalSyncError = False

    If Not IsGoogleAuthConfigured() Then
        MarkPWAFatalSyncError "ImportOtkupFromPWA_Core", _
            "Google OAuth2 nije konfigurisan."

        If showMessages Then
            MsgBox "Google OAuth2 nije konfigurisan!", vbCritical, APP_NAME
        End If

        Exit Function
    End If

    folderID = GetConfigValue("GOOGLE_PWA_FOLDER_ID")

    If Len(Trim$(folderID)) = 0 Then
        MarkPWAFatalSyncError "ImportOtkupFromPWA_Core", _
            "GOOGLE_PWA_FOLDER_ID nije postavljen."

        If showMessages Then
            MsgBox "GOOGLE_PWA_FOLDER_ID nije postavljen!", vbCritical, APP_NAME
        End If

        Exit Function
    End If

    LogInfo "ImportOtkupFromPWA_Core", "Import started."

    Set sheetIDs = New Collection
    Set sheetNames = New Collection

    If Not FindOTKSheets(folderID, sheetIDs, sheetNames) Then
        MarkPWAFatalSyncError "ImportOtkupFromPWA_Core", _
            "FindOTKSheets failed. Drive list could not be loaded."

        If showMessages Then
            MsgBox "Google Drive lista OTK fajlova nije ucitana. Proveri konekciju i log.", _
                   vbCritical, APP_NAME
        End If

        Exit Function
    End If

    If sheetIDs.count = 0 Then
        Monitor_MasterSyncSuccess _
            procedureName:="ImportOtkupFromPWA_Core", _
            importedCount:=0, _
            skippedCount:=0, _
            errorCount:=0, _
            filesCount:=0

        If showMessages Then
            MsgBox "Nema OTK-* fajlova u PWA folderu.", vbInformation, APP_NAME
        End If

        ImportOtkupFromPWA_Core = True
        Exit Function
    End If

    filesCount = sheetIDs.count

    For i = 1 To sheetIDs.count
        Dim imported As Long
        Dim skipped As Long
        Dim errors As Long

        imported = 0
        skipped = 0
        errors = 0

        Call ImportOneOTKSheet( _
            CStr(sheetIDs(i)), _
            CStr(sheetNames(i)), _
            imported, _
            skipped, _
            errors)

        totalImported = totalImported + imported
        totalSkipped = totalSkipped + skipped
        totalErrors = totalErrors + errors
    Next i

    LogInfo "ImportOtkupFromPWA_Core", _
        "Import completed. Files=" & CStr(filesCount) & _
        "; Imported=" & CStr(totalImported) & _
        "; Skipped=" & CStr(totalSkipped) & _
        "; Errors=" & CStr(totalErrors)

    If mLastPWAFatalSyncError Then
        Monitor_MasterSyncFail _
            procedureName:="ImportOtkupFromPWA_Core", _
            errNum:=0, _
            errDesc:="Fatal PWA sync error occurred during OTK import.", _
            errSrc:="modMasterSync.ImportOtkupFromPWA_Core", _
            importedCount:=totalImported, _
            skippedCount:=totalSkipped, _
            errorCount:=totalErrors

        If showMessages Then
            MsgBox Poruka("SYNC_ERR_UVOZ_OTK_NIJE") & vbCrLf & _
                   "Uvezeno: " & CStr(totalImported) & vbCrLf & _
                   "Preskoceno: " & CStr(totalSkipped) & vbCrLf & _
                   Poruka("SYNC_ERR_GRESKE") & CStr(totalErrors) & vbCrLf & vbCrLf & _
                   "Proveri log.", _
                   vbCritical, APP_NAME
        End If

        ImportOtkupFromPWA_Core = False
        Exit Function
    End If

    If totalErrors > 0 Then
        Monitor_MasterSyncFail _
            procedureName:="ImportOtkupFromPWA_Core", _
            errNum:=0, _
            errDesc:="OTK import completed with row-level errors.", _
            errSrc:="modMasterSync.ImportOtkupFromPWA_Core", _
            importedCount:=totalImported, _
            skippedCount:=totalSkipped, _
            errorCount:=totalErrors

        If showMessages Then
            MsgBox Poruka("SYNC_ERR_UVOZ_OTK_ZAVRSEN") & vbCrLf & vbCrLf & _
                   "Fajlova: " & CStr(filesCount) & vbCrLf & _
                   "Uvezeno: " & CStr(totalImported) & vbCrLf & _
                   "Preskoceno: " & CStr(totalSkipped) & vbCrLf & _
                   Poruka("SYNC_ERR_GRESKE") & CStr(totalErrors) & vbCrLf & vbCrLf & _
                   "Proveri log.", _
                   vbExclamation, APP_NAME
        End If

        ImportOtkupFromPWA_Core = False
        Exit Function
    End If

    Monitor_MasterSyncSuccess _
        procedureName:="ImportOtkupFromPWA_Core", _
        importedCount:=totalImported, _
        skippedCount:=totalSkipped, _
        errorCount:=totalErrors, _
        filesCount:=filesCount

    If showMessages Then
        MsgBox Poruka("SYNC_ERR_UVOZ_OTK_ZAVRSEN_2") & vbCrLf & vbCrLf & _
               "Fajlova: " & CStr(filesCount) & vbCrLf & _
               "Uvezeno: " & CStr(totalImported) & vbCrLf & _
               "Preskoceno: " & CStr(totalSkipped) & vbCrLf & _
               Poruka("SYNC_ERR_GRESKE") & CStr(totalErrors), _
               vbInformation, APP_NAME
    End If

    ImportOtkupFromPWA_Core = True
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next

    MarkPWAFatalSyncError "ImportOtkupFromPWA_Core", errDesc
    LogErr "ImportOtkupFromPWA_Core"

    Monitor_MasterSyncFail _
        procedureName:="ImportOtkupFromPWA_Core", _
        errNum:=errNum, _
        errDesc:=errDesc, _
        errSrc:=errSrc, _
        importedCount:=totalImported, _
        skippedCount:=totalSkipped, _
        errorCount:=totalErrors

    If showMessages Then
        MsgBox Poruka("SYNC_MSG_GRESKA_PRI_UVOZU") & errDesc, vbCritical, APP_NAME
    End If

    ImportOtkupFromPWA_Core = False
End Function


Public Sub ImportOtkupFromPWA_TX()
    Dim tx As clsTransaction
    Dim ok As Boolean

    On Error GoTo EH

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_AMBALAZA

    ok = ImportOtkupFromPWA_Core(False)

    If Not ok Then
        tx.RollbackTx

        Monitor_MasterSyncFail _
            procedureName:="ImportOtkupFromPWA_TX", _
            errNum:=0, _
            errDesc:="PWA import was not confirmed. Transaction rolled back because of fatal sync error.", _
            errSrc:="modMasterSync.ImportOtkupFromPWA_TX"

        MsgBox Poruka("SYNC_MSG_PWA_UVOZ_NIJE"), _
            vbCritical, APP_NAME
        Exit Sub
    End If

    tx.CommitTx

    MsgBox Poruka("SYNC_MSG_PWA_UVOZ_ZAVRSEN"), vbInformation, APP_NAME
    Exit Sub

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next

    LogErr "ImportOtkupFromPWA_TX"

    Monitor_MasterSyncFail _
        procedureName:="ImportOtkupFromPWA_TX", _
        errNum:=errNum, _
        errDesc:=errDesc, _
        errSrc:=errSrc

    If Not tx Is Nothing Then tx.RollbackTx

    MsgBox "Greska pri uvozu, promene vracene: " & errDesc, vbCritical, APP_NAME
End Sub

'======================================================================
' CreateOTKSheetsForAllStanice
'
' Manual wrapper za proveru/kreiranje OTK-* operational sheetova.
' Za full sync koristiti CreateOTKSheetsForAllStanice_Core(False, ...),
' da ne iskacu MsgBox-ovi tokom orchestrated sync ciklusa.
'======================================================================
Public Sub CreateOTKSheetsForAllStanice()
    Dim createdCount As Long
    Dim existingCount As Long
    Dim inactiveCount As Long
    Dim failedCount As Long
    Dim ok As Boolean

    ok = CreateOTKSheetsForAllStanice_Core( _
        True, _
        createdCount, _
        existingCount, _
        inactiveCount, _
        failedCount)

    If ok Then
        MsgBox "OTK sheetovi provereni." & vbCrLf & vbCrLf & _
               "Postojece: " & CStr(existingCount) & vbCrLf & _
               "Kreirano: " & CStr(createdCount) & vbCrLf & _
               "Neaktivne stanice preskocene: " & CStr(inactiveCount), _
               vbInformation, APP_NAME
    Else
        MsgBox "Provera/kreiranje OTK sheetova nije potpuno uspelo." & vbCrLf & vbCrLf & _
               "Postojece: " & CStr(existingCount) & vbCrLf & _
               "Kreirano: " & CStr(createdCount) & vbCrLf & _
               "Neaktivne stanice preskocene: " & CStr(inactiveCount) & vbCrLf & _
               "Greske: " & CStr(failedCount) & vbCrLf & vbCrLf & _
               "Proveri log.", _
               vbExclamation, APP_NAME
    End If
End Sub

'======================================================================
' CreateOTKSheetsForAllStanice_Core
'
' Idempotentno osigurava da svaka aktivna stanica ima svoj OTK-* Google
' spreadsheet u GOOGLE_PWA_FOLDER_ID folderu.
'
' Full sync koristi ovu core funkciju sa showMessages=False.
'
' Returns:
'   True  - sve aktivne stanice imaju OTK sheet ili su uspesno kreirane
'   False - fatal config/auth/schema greska ili bar jedan create/write fail
'======================================================================
Public Function CreateOTKSheetsForAllStanice_Core( _
    Optional ByVal showMessages As Boolean = False, _
    Optional ByRef createdCount As Long = 0, _
    Optional ByRef existingCount As Long = 0, _
    Optional ByRef inactiveCount As Long = 0, _
    Optional ByRef failedCount As Long = 0 _
) As Boolean

    Const SRC As String = "CreateOTKSheetsForAllStanice_Core"

    Dim data As Variant
    Dim colID As Long
    Dim colNaziv As Long
    Dim colAktivan As Long
    Dim folderID As String
    Dim i As Long
    Dim stanicaID As String
    Dim stanicaNaziv As String
    Dim sheetName As String
    Dim existingID As String
    Dim newID As String
    Dim headers As Variant

    On Error GoTo EH

    CreateOTKSheetsForAllStanice_Core = False

    createdCount = 0
    existingCount = 0
    inactiveCount = 0
    failedCount = 0

    If Not IsGoogleAuthConfigured() Then
        LogError SRC, "Google OAuth2 nije konfigurisan."
        If showMessages Then _
            MsgBox "Google OAuth2 nije konfigurisan!", vbCritical, APP_NAME
        Exit Function
    End If

    folderID = GetConfigValue("GOOGLE_PWA_FOLDER_ID")
    If Len(Trim$(folderID)) = 0 Then
        LogError SRC, "GOOGLE_PWA_FOLDER_ID nije postavljen."
        If showMessages Then _
            MsgBox "GOOGLE_PWA_FOLDER_ID nije postavljen!", vbCritical, APP_NAME
        Exit Function
    End If

    data = GetTableData(TBL_STANICE)
    If IsEmpty(data) Then
        LogWarn SRC, "tblStanice je prazan. Nema OTK sheetova za proveru/kreiranje."
        CreateOTKSheetsForAllStanice_Core = True
        Exit Function
    End If

    colID = RequireColumnIndex(TBL_STANICE, "StanicaID", SRC)
    colNaziv = RequireColumnIndex(TBL_STANICE, "Naziv", SRC)
    colAktivan = RequireColumnIndex(TBL_STANICE, "Aktivan", SRC)

    headers = BuildOTKOperationalHeaders_()

    For i = 1 To UBound(data, 1)
        stanicaID = Trim$(CStr(nz(data(i, colID), "")))
        stanicaNaziv = Trim$(CStr(nz(data(i, colNaziv), "")))

        If Len(stanicaID) = 0 Then
            failedCount = failedCount + 1
            LogError SRC, "Stanica bez StanicaID. Row=" & CStr(i)
            GoTo NextStanica
        End If

        If Not IsStanicaActiveForOTK_(data(i, colAktivan)) Then
            inactiveCount = inactiveCount + 1
            GoTo NextStanica
        End If

        sheetName = "OTK-" & stanicaID
        existingID = GetSpreadsheetID(sheetName, folderID)

        If Len(Trim$(existingID)) > 0 Then
            existingCount = existingCount + 1
            GoTo NextStanica
        End If

        newID = CreateSpreadsheet(sheetName, folderID)

        If Len(Trim$(newID)) = 0 Then
            failedCount = failedCount + 1
            LogError SRC, _
                "CreateSpreadsheet failed. Sheet=" & sheetName & _
                "; StanicaID=" & stanicaID & _
                "; Naziv=" & stanicaNaziv
            GoTo NextStanica
        End If

        If Not WriteSheetData(newID, "Sheet1", headers) Then
            failedCount = failedCount + 1
            LogError SRC, _
                "WriteSheetData header failed. Sheet=" & sheetName & _
                "; SpreadsheetID=" & newID & _
                "; StanicaID=" & stanicaID
            GoTo NextStanica
        End If

        createdCount = createdCount + 1

        LogInfo SRC, _
            "OTK sheet created. Sheet=" & sheetName & _
            "; SpreadsheetID=" & newID & _
            "; StanicaID=" & stanicaID & _
            "; Naziv=" & stanicaNaziv

NextStanica:
    Next i

    CreateOTKSheetsForAllStanice_Core = (failedCount = 0)

    If CreateOTKSheetsForAllStanice_Core Then
        LogInfo SRC, _
            "OTK sheet ensure completed. Existing=" & CStr(existingCount) & _
            "; Created=" & CStr(createdCount) & _
            "; InactiveSkipped=" & CStr(inactiveCount) & _
            "; Failed=" & CStr(failedCount)
    Else
        LogWarn SRC, _
            "OTK sheet ensure completed with errors. Existing=" & CStr(existingCount) & _
            "; Created=" & CStr(createdCount) & _
            "; InactiveSkipped=" & CStr(inactiveCount) & _
            "; Failed=" & CStr(failedCount)
    End If

    Exit Function

EH:
    failedCount = failedCount + 1
    LogErr SRC

    If showMessages Then
        MsgBox "Greska pri proveri/kreiranju OTK sheetova: " & Err.description, _
               vbCritical, APP_NAME
    End If

    CreateOTKSheetsForAllStanice_Core = False
End Function

'======================================================================
' BuildOTKOperationalHeaders_
'
' Header schema za operational OTK-* sheetove koje PWA puni, a VBA
' ImportOtkupFromPWA_Core cita.
'======================================================================
Private Function BuildOTKOperationalHeaders_() As Variant
    Dim headers(1 To 1, 1 To 23) As Variant

    headers(1, 1) = "ClientRecordID"
    headers(1, 2) = "ServerRecordID"
    headers(1, 3) = "CreatedAtClient"
    headers(1, 4) = "UpdatedAtClient"
    headers(1, 5) = "UpdatedAtServer"
    headers(1, 6) = "SyncStatus"
    headers(1, 7) = "DeviceID"
    headers(1, 8) = "OtkupacID"
    headers(1, 9) = "Datum"
    headers(1, 10) = "KooperantID"
    headers(1, 11) = "KooperantName"
    headers(1, 12) = "VrstaVoca"
    headers(1, 13) = "SortaVoca"
    headers(1, 14) = "Klasa"
    headers(1, 15) = "Kolicina"
    headers(1, 16) = "Cena"
    headers(1, 17) = "TipAmbalaze"
    headers(1, 18) = "KolAmbalaze"
    headers(1, 19) = "ParcelaID"
    headers(1, 20) = "VozacID"
    headers(1, 21) = "Napomena"
    headers(1, 22) = "ReceivedAt"
    headers(1, 23) = "BrojDokumenta"

    BuildOTKOperationalHeaders_ = headers
End Function

'======================================================================
' IsStanicaActiveForOTK_
'
' Postojeca logika je tretirala sve osim "Ne" kao aktivno.
' Ovaj helper zadrzava isti business rule, ali ga izoluje.
'======================================================================
Private Function IsStanicaActiveForOTK_(ByVal activeValue As Variant) As Boolean
    Dim s As String

    s = UCase$(Trim$(CStr(nz(activeValue, ""))))

    IsStanicaActiveForOTK_ = Not (s = "NE")
End Function

Public Function AutoCreateOtpremniceFromPWA_TX() As Long
    Const SRC As String = "AutoCreateOtpremniceFromPWA_TX"

    Dim tx As clsTransaction
    Dim createdCount As Long

    On Error GoTo EH

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_AMBALAZA

    createdCount = AutoCreateOtpremniceFromPWA()

    tx.CommitTx
    Set tx = Nothing

    AutoCreateOtpremniceFromPWA_TX = createdCount

    LogInfo SRC, "Auto-create Otpremnice completed. Created=" & CStr(createdCount)
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr SRC
    On Error GoTo 0

    Err.Raise errNum, SRC, "Source=" & errSrc & " | " & errDesc
End Function

Public Function AutoCreateOtpremniceFromPWA() As Long
    ' Nach ImportOtkupFromPWA: erstellt Otpremnice fuer PWA-Otkupi mit VozacID
    ' Gruppierung: StanicaID + Datum + VozacID + Klasa (= AutoLink Key)
    ' Returns: Anzahl erstellter Otpremnice
    
    Dim data As Variant
    data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function
    data = ExcludeStornirano(data, TBL_OTKUP)
    If IsEmpty(data) Then Exit Function
    
    Dim colID As Long, colSt As Long, colDat As Long, colVoz As Long
    Dim colOtpID As Long, colKlasa As Long, colVrsta As Long, colSorta As Long
    Dim colKol As Long, colCena As Long, colTipAmb As Long, colKolAmb As Long
    
    colID = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, "AutoCreateOtpremniceFromPWA")
    colSt = RequireColumnIndex(TBL_OTKUP, COL_OTK_STANICA, "AutoCreateOtpremniceFromPWA")
    colDat = RequireColumnIndex(TBL_OTKUP, COL_OTK_DATUM, "AutoCreateOtpremniceFromPWA")
    colVoz = RequireColumnIndex(TBL_OTKUP, COL_OTK_VOZAC, "AutoCreateOtpremniceFromPWA")
    colOtpID = RequireColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID, "AutoCreateOtpremniceFromPWA")
    colKlasa = RequireColumnIndex(TBL_OTKUP, COL_OTK_KLASA, "AutoCreateOtpremniceFromPWA")
    colVrsta = RequireColumnIndex(TBL_OTKUP, COL_OTK_VRSTA, "AutoCreateOtpremniceFromPWA")
    colSorta = RequireColumnIndex(TBL_OTKUP, COL_OTK_SORTA, "AutoCreateOtpremniceFromPWA")
    colKol = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA, "AutoCreateOtpremniceFromPWA")
    colCena = RequireColumnIndex(TBL_OTKUP, COL_OTK_CENA, "AutoCreateOtpremniceFromPWA")
    colTipAmb = RequireColumnIndex(TBL_OTKUP, COL_OTK_TIP_AMB, "AutoCreateOtpremniceFromPWA")
    colKolAmb = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOL_AMB, "AutoCreateOtpremniceFromPWA")
    
    ' Sammle unverknuepfte Otkupi MIT VozacID ? gruppiere nach Key
    ' Key = StanicaID|Datum|VozacID|Klasa
    Dim groups As Object
    Set groups = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim vozID As String: vozID = Trim$(CStr(nz(data(i, colVoz), "")))
        Dim otpID As String: otpID = Trim$(CStr(nz(data(i, colOtpID), "")))
        
        ' Nur Otkupi ohne Otpremnica UND mit VozacID
        If otpID = "" And vozID <> "" Then
            Dim gKey As String
            gKey = CStr(data(i, colSt)) & "|" & _
                   Format$(CDate(data(i, colDat)), "YYYY-MM-DD") & "|" & _
                   vozID & "|" & _
                   CStr(nz(data(i, colKlasa), ""))
            
            If Not groups.Exists(gKey) Then
                groups.Add gKey, New Collection
            End If
            groups(gKey).Add i  ' Row index in data array
        End If
    Next i
    
    If groups.count = 0 Then
        AutoCreateOtpremniceFromPWA = 0
        Exit Function
    End If
    
    ' Fuer jede Gruppe: Otpremnica erstellen + Otkupi verknuepfen
    Dim created As Long
    Dim keys As Variant: keys = groups.keys
    Dim k As Long
    
    ' Otpremnica-Zaehler pro Stanica+Datum vorladen
    Dim otpAll As Variant
    otpAll = GetTableData(TBL_OTPREMNICA)
    If Not IsEmpty(otpAll) Then otpAll = ExcludeStornirano(otpAll, TBL_OTPREMNICA)
    
    Dim colOtpSt As Long: colOtpSt = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_STANICA, "AutoCreateOtpremniceFromPWA")
    Dim colOtpDat As Long: colOtpDat = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_DATUM, "AutoCreateOtpremniceFromPWA")
    
    For k = 0 To UBound(keys)
        Dim parts() As String: parts = Split(keys(k), "|")
        ' parts(0)=StanicaID, parts(1)=Datum, parts(2)=VozacID, parts(3)=Klasa
        
        Dim grpRows As Collection: Set grpRows = groups(keys(k))
        
        ' Aggregiere Kolicina, Ambalaza, nehme Vrsta/Sorta/Cena vom ersten
        Dim totalKol As Double: totalKol = 0
        Dim totalAmb As Long: totalAmb = 0
        Dim firstRow As Long: firstRow = grpRows(1)
        
        Dim r As Long
        For r = 1 To grpRows.count
            Dim ri As Long: ri = grpRows(r)

            RequireSingleMasterSyncRow TBL_OTKUP, COL_OTK_ID, CStr(data(ri, colID)), _
                               "AutoCreateOtpremniceFromPWA"

            totalKol = totalKol + CDbl(nz(data(ri, colKol), 0))
            totalAmb = totalAmb + CLng(nz(data(ri, colKolAmb), 0))
        Next r
        
        ' BrojOtpremnice: kanon "x/ddmmyy[-rb]" preko modBrojevi.GenerateBrojOtpremnice.
        ' Helper interno radi MaxSeqFromTable scan; ne treba lokalni seqDict.
        Dim brojOtp As String
        brojOtp = GenerateBrojOtpremnice(parts(0), CDate(parts(1)))
        
        If Len(brojOtp) = 0 Then
            Err.Raise vbObjectError + 8200, "AutoCreateOtpremniceFromPWA", _
                "GenerateBrojOtpremnice nije vratio broj za stanica=" & parts(0)
        End If
        
        ' Otpremnica erstellen (BrojZbirne leer -- Vozac/Operator setzt spaeter)
        Dim newOtpID As String
        newOtpID = SaveOtpremnica_TX( _
            CDate(parts(1)), _
            parts(0), _
            parts(2), _
            brojOtp, _
            "", _
            CStr(nz(data(firstRow, colVrsta), "")), _
            CStr(nz(data(firstRow, colSorta), "")), _
            totalKol, _
            CDbl(nz(data(firstRow, colCena), 0)), _
            CStr(nz(data(firstRow, colTipAmb), "")), _
            totalAmb, _
            parts(3) _
        )
        
        If Len(Trim$(newOtpID)) = 0 Then
            Err.Raise ERR_MASTER_SYNC_GUARD_BASE + 20, "AutoCreateOtpremniceFromPWA", _
                    "SaveOtpremnica_TX nije vratio OtpremnicaID. BrojOtpremnice=" & brojOtp
        End If

        RequireSingleMasterSyncRow TBL_OTPREMNICA, COL_OTP_ID, newOtpID, _
                           "AutoCreateOtpremniceFromPWA"

        For r = 1 To grpRows.count
            ri = grpRows(r)

            Dim otkupID As String
            otkupID = CStr(data(ri, colID))

            LinkOtkupToOtpremnicaStrict otkupID, newOtpID, _
                                        "AutoCreateOtpremniceFromPWA"
        Next r

        created = created + 1
    Next k
    
    AutoCreateOtpremniceFromPWA = created
End Function

' ============================================================
' MALINA MOD -- C: VozacID := StanicaID na tblOtkup.
'
' AutoCreateOtpremniceFromPWA pravi otpremnice samo za otkupe koji IMAJU
' VozacID (grupisanje StanicaID|Datum|VozacID|Klasa, vidi filter gore).
' U malina modu nema vozaca, pa se PRE auto-otpremnice VozacID popunjava
' StanicaID-em -- time se okidac pali, a brojevi ostaju konzistentni
' (otkupac == stanica). Idempotentno: dira samo prazan VozacID.
' Self-gated: u visnji ne radi nista.
' ============================================================
Public Function StampVozacFromStanicaForMalina_TX() As Long
    Const SRC As String = "StampVozacFromStanicaForMalina_TX"

    Dim tx As clsTransaction

    On Error GoTo EH

    If Not IsMalinaMode() Then
        StampVozacFromStanicaForMalina_TX = 0
        Exit Function
    End If

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP

    StampVozacFromStanicaForMalina_TX = StampVozacFromStanicaForMalina()

    tx.CommitTx
    Set tx = Nothing

    LogInfo SRC, "Malina VozacID:=StanicaID stamped=" & CStr(StampVozacFromStanicaForMalina_TX)
    Exit Function

EH:
    Dim errNum As Long, errDesc As String, errSrc As String
    errNum = Err.Number: errDesc = Err.description: errSrc = Err.SOURCE
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr SRC
    On Error GoTo 0
    Err.Raise errNum, SRC, "Source=" & errSrc & " | " & errDesc
End Function

Public Function StampVozacFromStanicaForMalina() As Long
    Const SRC As String = "StampVozacFromStanicaForMalina"

    Dim data As Variant
    data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function

    Dim colVoz As Long, colSt As Long, colStorno As Long
    colVoz = RequireColumnIndex(TBL_OTKUP, COL_OTK_VOZAC, SRC)
    colSt = RequireColumnIndex(TBL_OTKUP, COL_OTK_STANICA, SRC)
    colStorno = GetColumnIndex(TBL_OTKUP, COL_OTK_STORNIRANO)

    Dim r As Long, cnt As Long
    For r = 1 To UBound(data, 1)
        Dim skip As Boolean
        skip = (colStorno > 0) And _
               (UCase$(Trim$(CStr(nz(data(r, colStorno), "")))) = "DA")
        If Not skip Then
            Dim voz As String: voz = Trim$(CStr(nz(data(r, colVoz), "")))
            Dim st As String: st = Trim$(CStr(nz(data(r, colSt), "")))
            If voz = "" And st <> "" Then
                RequireUpdateCell TBL_OTKUP, r, COL_OTK_VOZAC, st, SRC
                cnt = cnt + 1
            End If
        End If
    Next r

    StampVozacFromStanicaForMalina = cnt
End Function

' ============================================================
' MALINA MOD -- D: auto-zbirna iz otpremnice (1:1).
'
' Za svaku aktivnu otpremnicu sa praznim BrojZbirne (grupisano po
' BrojOtpremnice, Klasa I+II istog dokumenta zajedno) pravi zbirnu preko
' postojeceg SaveZbirnaMulti_TX:
'   - BrojZbirne := BrojOtpremnice (broj garantovano identican otpremnici)
'   - kupac := MALINA_DEFAULT_KUPAC (Hladnjaca); Hladnjaca naziv iz tblKupci
'   - backfill BrojZbirne na otpremnicu i na tblOtkup (preko OtpremnicaID),
'     jer ValidateZbirna/prijemnica/faktura vezu drze preko BrojZbirne.
' Idempotentno: prazan-BrojZbirne filter sprecava duplo kreiranje.
' Self-gated: u visnji ne radi nista.
' ============================================================
' samoBrojOtp: opcioni scope -- obradi SAMO otpremnice tog broja. Prazno = sve
' otvorene (produkcioni poziv iz frmDokumenta). Scope koriste testovi, da run ne
' zahvati nepovezane otvorene otpremnice u svesci.
Public Function AutoCreateZbirnaFromOtpremnice_TX(Optional ByVal samoBrojOtp As String = "") As Long
    Const SRC As String = "AutoCreateZbirnaFromOtpremnice_TX"

    Dim tx As clsTransaction

    On Error GoTo EH

    If Not IsMalinaMode() Then
        AutoCreateZbirnaFromOtpremnice_TX = 0
        Exit Function
    End If

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_OTKUP

    AutoCreateZbirnaFromOtpremnice_TX = AutoCreateZbirnaFromOtpremnice(samoBrojOtp)

    tx.CommitTx
    Set tx = Nothing

    LogInfo SRC, "Malina auto-zbirna created=" & CStr(AutoCreateZbirnaFromOtpremnice_TX)
    Exit Function

EH:
    Dim errNum As Long, errDesc As String, errSrc As String
    errNum = Err.Number: errDesc = Err.description: errSrc = Err.SOURCE
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr SRC
    On Error GoTo 0
    Err.Raise errNum, SRC, "Source=" & errSrc & " | " & errDesc
End Function

Public Function AutoCreateZbirnaFromOtpremnice(Optional ByVal samoBrojOtp As String = "") As Long
    Const SRC As String = "AutoCreateZbirnaFromOtpremnice"

    If Not IsMalinaMode() Then Exit Function

    Dim kupacID As String
    kupacID = Trim$(GetConfigValue(CFG_MALINA_DEFAULT_KUPAC))
    If kupacID = "" Then
        Err.Raise vbObjectError + 8300, SRC, _
            "MALINA_DEFAULT_KUPAC nije postavljen (kljuc u tblSEFConfig)."
    End If

    Dim hladnjaca As String
    hladnjaca = CStr(nz(LookupValue(TBL_KUPCI, COL_KUP_ID, kupacID, "Hladnjaca"), ""))

    Dim data As Variant
    data = GetTableData(TBL_OTPREMNICA)
    If IsEmpty(data) Then Exit Function

    Dim cId As Long, cBrZ As Long, cBrO As Long, cDat As Long, cVoz As Long
    Dim cVrsta As Long, cSorta As Long, cKol As Long, cTipAmb As Long
    Dim cKolAmb As Long, cKlasa As Long, cStorno As Long
    cId = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_ID, SRC)
    cBrZ = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, SRC)
    cBrO = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ, SRC)
    cDat = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_DATUM, SRC)
    cVoz = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VOZAC, SRC)
    cVrsta = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VRSTA, SRC)
    cSorta = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_SORTA, SRC)
    cKol = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA, SRC)
    cTipAmb = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_TIP_AMB, SRC)
    cKolAmb = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KOL_AMB, SRC)
    cKlasa = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KLASA, SRC)
    cStorno = GetColumnIndex(TBL_OTPREMNICA, COL_STORNIRANO)

    ' Obrada PO REDU otpremnice (NE grupisano po BrojOtpremnice):
    '   AutoCreateOtpremniceFromPWA pravi otpremnice PO-KLASI (Klasa I i Klasa II
    '   imaju razlicit BrojOtpremnice), dok rucni SaveOtpremnicaMulti stavlja obe
    '   klase pod isti BrojOtpremnice. Per-red sa SaveZbirna_TX (jedna klasa)
    '   korektno pokriva oba slucaja: BrojZbirne := BrojOtpremnice tog reda.
    '   (Grupisanje+SaveZbirnaMulti je padalo na Klasa-II-only grupi: kolI=0.)
    Dim otpMap As Object
    Set otpMap = CreateObject("Scripting.Dictionary")

    Dim created As Long: created = 0
    Dim r As Long
    For r = 1 To UBound(data, 1)
        Dim brz As String: brz = Trim$(CStr(nz(data(r, cBrZ), "")))
        Dim brO As String: brO = Trim$(CStr(nz(data(r, cBrO), "")))
        Dim isStorno As Boolean
        isStorno = (cStorno > 0) And _
                   (UCase$(Trim$(CStr(nz(data(r, cStorno), "")))) = "DA")

        Dim inScope As Boolean
        inScope = (Len(Trim$(samoBrojOtp)) = 0) Or (brO = Trim$(samoBrojOtp))

        If brz = "" And brO <> "" And Not isStorno And inScope Then
            Dim datum As Date: datum = CDate(data(r, cDat))
            Dim vozacID As String: vozacID = Trim$(CStr(nz(data(r, cVoz), "")))
            Dim vrsta As String: vrsta = CStr(nz(data(r, cVrsta), ""))
            Dim sorta As String: sorta = CStr(nz(data(r, cSorta), ""))
            Dim tipAmb As String: tipAmb = CStr(nz(data(r, cTipAmb), ""))
            Dim klasa As String: klasa = CStr(nz(data(r, cKlasa), ""))
            Dim kol As Double: kol = CDbl(nz(data(r, cKol), 0))
            Dim amb As Long: amb = CLng(nz(data(r, cKolAmb), 0))

            ' Mirror-stanica (VozacID==StanicaID) -> zbirna nosi "S" prefiks
            ' (S1/ddmmyy); otpremnica zadrzava svoj broj (BrojOtpremnice = 1/ddmmyy).
            Dim brZbirne As String: brZbirne = ApplyMirrorPrefix(vozacID, brO)

            Dim zbrRes As String
            zbrRes = SaveZbirna_TX(datum, vozacID, brZbirne, kupacID, _
                        hladnjaca, "", vrsta, sorta, kol, tipAmb, amb, klasa)

            If Len(Trim$(zbrRes)) = 0 Then
                Err.Raise vbObjectError + 8301, SRC, _
                    "SaveZbirna_TX nije vratio ZbirnaID za BrojOtpremnice=" & brO & _
                    " Klasa=" & klasa
            End If

            RequireUpdateCell TBL_OTPREMNICA, r, COL_OTP_BROJ_ZBIRNE, brZbirne, SRC
            Dim otpID As String: otpID = Trim$(CStr(nz(data(r, cId), "")))
            If otpID <> "" Then otpMap(otpID) = brZbirne

            created = created + 1
        End If
    Next r

    If created = 0 Then Exit Function

    ' backfill BrojZbirne na tblOtkup (preko OtpremnicaID), jedan prolaz
    BackfillOtkupBrojZbirneByOtpremnica otpMap, SRC

    AutoCreateZbirnaFromOtpremnice = created
End Function

Private Sub BackfillOtkupBrojZbirneByOtpremnica(ByVal otpMap As Object, ByVal callerSrc As String)
    If otpMap Is Nothing Then Exit Sub
    If otpMap.count = 0 Then Exit Sub

    Dim data As Variant
    data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Sub

    Dim cOtpID As Long, cBrZ As Long
    cOtpID = RequireColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID, callerSrc)
    cBrZ = RequireColumnIndex(TBL_OTKUP, COL_OTK_BROJ_ZBIRNE, callerSrc)

    Dim r As Long
    For r = 1 To UBound(data, 1)
        Dim otpID As String: otpID = Trim$(CStr(nz(data(r, cOtpID), "")))
        If otpID <> "" Then
            If otpMap.Exists(otpID) Then
                Dim cur As String: cur = Trim$(CStr(nz(data(r, cBrZ), "")))
                If cur = "" Then
                    RequireUpdateCell TBL_OTKUP, r, COL_OTK_BROJ_ZBIRNE, _
                        CStr(otpMap(otpID)), callerSrc
                End If
            End If
        End If
    Next r
End Sub

' ============================================================
' PRIVATE -- Find OTK-* Sheets in Folder
' ============================================================

Private Function FindOTKSheets(ByVal folderID As String, _
                               ByRef outIDs As Collection, _
                               ByRef outNames As Collection) As Boolean
    Const SOURCE As String = "FindOTKSheets"

    Dim accessToken As String
    Dim url As String
    Dim http As Object
    Dim query As String
    Dim responseText As String
    Dim nextPageToken As String
    Dim tokenPos As Long

    On Error GoTo EH

    If Len(Trim$(folderID)) = 0 Then
        LogError SOURCE, "folderID je prazan."
        FindOTKSheets = False
        Exit Function
    End If

    accessToken = GetAccessToken()
    If Len(accessToken) = 0 Then
        LogError SOURCE, "Kein Access Token"
        FindOTKSheets = False
        Exit Function
    End If

    query = "name contains 'OTK-' and mimeType='application/vnd.google-apps.spreadsheet'" & _
            " and '" & EscapeDriveQueryValueMasterSync(folderID) & "' in parents and trashed=false"

    nextPageToken = ""

    Do
        url = "https://www.googleapis.com/drive/v3/files" & _
              "?q=" & UrlEncode(query) & _
              "&fields=nextPageToken,files(id,name)" & _
              "&pageSize=100"

        If Len(nextPageToken) > 0 Then
            url = url & "&pageToken=" & UrlEncode(nextPageToken)
        End If

        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
        http.SetTimeouts 10000, 10000, 30000, 30000

        http.Open "GET", url, False
        http.SetRequestHeader "Authorization", "Bearer " & accessToken
        http.Send

        responseText = CStr(http.responseText)

        If http.status <> 200 Then
            LogError SOURCE, _
                     "HTTP " & http.status & ": " & Left$(responseText, 1000), _
                     http.status
            FindOTKSheets = False
            Exit Function
        End If

        Call ParseFileList(responseText, outIDs, outNames)

        tokenPos = InStr(1, responseText, """nextPageToken""", vbTextCompare)
        If tokenPos > 0 Then
            nextPageToken = ExtractJsonValueAt(responseText, tokenPos)
        Else
            nextPageToken = ""
        End If
    Loop While Len(nextPageToken) > 0

    LogInfo SOURCE, "Gefunden: " & outIDs.count & " OTK-Sheets"

    FindOTKSheets = True
    Exit Function

EH:
    LogErr SOURCE
    FindOTKSheets = False
End Function

Private Sub ParseFileList(ByVal json As String, _
                          ByRef outIDs As Collection, _
                          ByRef outNames As Collection)
    ' Parst {"files":[{"id":"xxx","name":"OTK-ST-00001"},...]
    Dim pos As Long, endPos As Long
    Dim fileID As String, fileName As String
    
    pos = 1
    Do
        ' Suche naechstes "id"
        pos = InStr(pos, json, """id""", vbTextCompare)
        If pos = 0 Then Exit Do
        
        fileID = ExtractJsonValueAt(json, pos)
        
        ' Suche "name" danach
        Dim namePos As Long
        namePos = InStr(pos, json, """name""", vbTextCompare)
        If namePos = 0 Then Exit Do
        
        fileName = ExtractJsonValueAt(json, namePos)
        
        If Len(fileID) > 0 And Len(fileName) > 0 Then
            ' Nur OTK-Sheets (Sicherheit)
            If Left$(fileName, 4) = "OTK-" Then
                outIDs.Add fileID
                outNames.Add fileName
            End If
        End If
        
        pos = namePos + 1
    Loop
End Sub

Private Function ExtractJsonValueAt(ByVal json As String, ByVal startPos As Long) As String
    ' Extrahiert den String-Wert nach "key":"value" ab startPos
    Dim p As Long, q As Long
    
    p = InStr(startPos, json, ":")
    If p = 0 Then Exit Function
    
    p = InStr(p, json, """")
    If p = 0 Then Exit Function
    
    p = p + 1
    q = InStr(p, json, """")
    If q = 0 Then Exit Function
    
    ExtractJsonValueAt = Mid$(json, p, q - p)
End Function
Private Function EscapeDriveQueryValueMasterSync(ByVal value As String) As String
    Dim result As String

    result = CStr(value)
    result = Replace(result, "\", "\\")
    result = Replace(result, "'", "\'")

    EscapeDriveQueryValueMasterSync = result
End Function

Private Function ValidateOTKSheetHeader(ByVal data As Variant, _
                                        ByVal sheetName As String) As Boolean
    Const SOURCE As String = "ValidateOTKSheetHeader"

    On Error GoTo EH

    If IsEmpty(data) Then
        LogError SOURCE, "Sheet data is Empty: " & sheetName
        ValidateOTKSheetHeader = False
        Exit Function
    End If

    If UBound(data, 1) < 1 Then
        LogError SOURCE, "Sheet nema header row: " & sheetName
        ValidateOTKSheetHeader = False
        Exit Function
    End If

    If UBound(data, 2) < 22 Then
        LogError SOURCE, _
                 "OTK schema drift: premalo kolona u sheetu " & sheetName & _
                 ". Expected=22, Actual=" & CStr(UBound(data, 2))
        ValidateOTKSheetHeader = False
        Exit Function
    End If

    If Not RequireOTKHeaderValue(data, sheetName, GS_CLIENT_RECORD_ID, "ClientRecordID") Then Exit Function
    If Not RequireOTKHeaderValue(data, sheetName, GS_SERVER_RECORD_ID, "ServerRecordID") Then Exit Function
    If Not RequireOTKHeaderValue(data, sheetName, GS_CREATED_AT, "CreatedAtClient") Then Exit Function
    If Not RequireOTKHeaderValue(data, sheetName, GS_UPDATED_AT_CLIENT, "UpdatedAtClient") Then Exit Function
    If Not RequireOTKHeaderValue(data, sheetName, GS_UPDATED_AT_SERVER, "UpdatedAtServer") Then Exit Function
    If Not RequireOTKHeaderValue(data, sheetName, GS_SYNC_STATUS, "SyncStatus") Then Exit Function
    If Not RequireOTKHeaderValue(data, sheetName, GS_DEVICE_ID, "DeviceID") Then Exit Function
    If Not RequireOTKHeaderValue(data, sheetName, GS_OTKUPAC_ID, "OtkupacID") Then Exit Function
    If Not RequireOTKHeaderValue(data, sheetName, GS_DATUM, "Datum") Then Exit Function
    If Not RequireOTKHeaderValue(data, sheetName, GS_KOOPERANT_ID, "KooperantID") Then Exit Function
    If Not RequireOTKHeaderValue(data, sheetName, GS_KOOPERANT_NAME, "KooperantName") Then Exit Function
    If Not RequireOTKHeaderValue(data, sheetName, GS_VRSTA, "VrstaVoca") Then Exit Function
    If Not RequireOTKHeaderValue(data, sheetName, GS_SORTA, "SortaVoca") Then Exit Function
    If Not RequireOTKHeaderValue(data, sheetName, GS_KLASA, "Klasa") Then Exit Function
    If Not RequireOTKHeaderValue(data, sheetName, GS_KOLICINA, "Kolicina") Then Exit Function
    If Not RequireOTKHeaderValue(data, sheetName, GS_CENA, "Cena") Then Exit Function
    If Not RequireOTKHeaderValue(data, sheetName, GS_TIP_AMB, "TipAmbalaze") Then Exit Function
    If Not RequireOTKHeaderValue(data, sheetName, GS_KOL_AMB, "KolAmbalaze") Then Exit Function
    If Not RequireOTKHeaderValue(data, sheetName, GS_PARCELA_ID, "ParcelaID") Then Exit Function
    If Not RequireOTKHeaderValue(data, sheetName, GS_VOZAC_ID, "VozacID") Then Exit Function
    If Not RequireOTKHeaderValue(data, sheetName, GS_NAPOMENA, "Napomena") Then Exit Function
    If Not RequireOTKHeaderValue(data, sheetName, GS_RECEIVED_AT, "ReceivedAt") Then Exit Function
    If Not RequireOTKHeaderValue(data, sheetName, GS_BROJ_DOKUMENTA, "BrojDokumenta") Then Exit Function

    ValidateOTKSheetHeader = True
    Exit Function

EH:
    LogErr SOURCE, "Sheet: " & sheetName
    ValidateOTKSheetHeader = False
End Function

Private Function RequireOTKHeaderValue(ByVal data As Variant, _
                                       ByVal sheetName As String, _
                                       ByVal colIndex As Long, _
                                       ByVal expectedHeader As String) As Boolean
    Dim actualHeader As String

    actualHeader = Trim$(CStr(data(1, colIndex)))

    If StrComp(actualHeader, expectedHeader, vbBinaryCompare) <> 0 Then
        LogError "ValidateOTKSheetHeader", _
                 "OTK schema drift in " & sheetName & _
                 ". Col=" & CStr(colIndex) & _
                 ", Expected='" & expectedHeader & "'" & _
                 ", Actual='" & actualHeader & "'"
        RequireOTKHeaderValue = False
        Exit Function
    End If

    RequireOTKHeaderValue = True
End Function

' ============================================================
' PRIVATE -- Import eines einzelnen OTK-Sheets
' ============================================================
Public Sub TestHook_ImportOneOTKSheet(ByVal spreadsheetID As String, _
                                      ByVal sheetName As String, _
                                      ByRef outImported As Long, _
                                      ByRef outSkipped As Long, _
                                      ByRef outErrors As Long)
    ' DEV/SMOKE TEST HOOK ONLY.
    ' Keeps ImportOneOTKSheet private for production callers,
    ' but allows isolated fixture-based sync tests.

    Call ImportOneOTKSheet(spreadsheetID, sheetName, outImported, outSkipped, outErrors)
End Sub

Private Sub ImportOneOTKSheet(ByVal spreadsheetID As String, _
                              ByVal sheetName As String, _
                              ByRef outImported As Long, _
                              ByRef outSkipped As Long, _
                              ByRef outErrors As Long)
    Dim data As Variant
    Dim i As Long
    Dim syncStatus As String
    Dim statusUpdates As Collection
    
    On Error GoTo EH
    
    ' Daten lesen (erster Tab)
    data = ReadSheetData(spreadsheetID, "Sheet1")

    If Not IsEmpty(data) Then
        Debug.Print "Rows: " & UBound(data, 1) & " Cols: " & UBound(data, 2)
    End If
    
    If IsEmpty(data) Then
        LogWarn "ImportOneOTKSheet", "Leeres Sheet: " & sheetName
        Exit Sub
    End If
    
    If Not ValidateOTKSheetHeader(data, sheetName) Then
        outErrors = outErrors + 1
        MarkPWAFatalSyncError "ImportOneOTKSheet", _
            "Import aborted because OTK header schema is invalid. Sheet=" & sheetName
        Exit Sub
    End If
    
    ' Erste Zeile = Header, ab Zeile 2 = Daten
    If UBound(data, 1) < 2 Then
        LogInfo "ImportOneOTKSheet", "Keine Daten in: " & sheetName
        Exit Sub
    End If
    
    Set statusUpdates = New Collection
    
    For i = 2 To UBound(data, 1)
        ' Pruefe SyncStatus
        syncStatus = Trim$(CStr(data(i, GS_SYNC_STATUS)))
        
        ' Nur "Synced" importieren (= vom Apps Script geschrieben, noch nicht im Master)
        If syncStatus = SYNC_STATUS_PENDING Then
            
            Dim clientRecordID As String
            clientRecordID = Trim$(CStr(data(i, GS_CLIENT_RECORD_ID)))

            If Len(clientRecordID) = 0 Then
                statusUpdates.Add Array(i, SYNC_STATUS_ERROR & ":ClientRecordID missing", "")
                outErrors = outErrors + 1
                LogWarn "ImportOneOTKSheet", _
                        sheetName & " Row " & i & ": ClientRecordID missing. Import skipped."
                GoTo NextImportRow
            End If

            ' Duplikat-Check im Master
            If IsDuplicateInMaster(clientRecordID) Then
                ' Proveri da li je VozacID update (Otprema tab)
                Dim sheetVozac As String
                sheetVozac = Trim$(CStr(nz(data(i, GS_VOZAC_ID), "")))
                If Len(sheetVozac) > 0 Then
                    If TryUpdateVozacID(clientRecordID, sheetVozac) Then
                        statusUpdates.Add Array(i, SYNC_STATUS_MASTER)
                    Else
                        statusUpdates.Add Array(i, SYNC_STATUS_DUPLICATE)
                    End If
                Else
                    statusUpdates.Add Array(i, SYNC_STATUS_DUPLICATE)
                End If
                outSkipped = outSkipped + 1
            Else
                ' Validierung
                Dim validationError As String
                validationError = ValidatePWAOtkup(data, i)
                
                If Len(validationError) > 0 Then
                    statusUpdates.Add Array(i, SYNC_STATUS_ERROR & ":" & validationError)
                    outErrors = outErrors + 1
                    LogWarn "ImportOneOTKSheet", sheetName & " Row " & i & ": " & validationError
                Else
                    ' Import in tblOtkup
                    Dim newOtkupID As String
                    newOtkupID = ImportRowToTblOtkup_RowTX(data, i, clientRecordID)
                    If Len(newOtkupID) > 0 Then
                        statusUpdates.Add Array(i, SYNC_STATUS_MASTER, newOtkupID)
                        outImported = outImported + 1
                    Else
                        statusUpdates.Add Array(i, SYNC_STATUS_ERROR & ":AppendRow failed", "")
                        outErrors = outErrors + 1
                    End If
                End If
            End If
        Else
            ' Bereits importiert oder Error ? ueberspringen
            outSkipped = outSkipped + 1
        End If

NextImportRow:
    Next i
    
    ' SyncStatus zurueckschreiben in Google Sheet
    If statusUpdates.count > 0 Then
        If Not WriteBackSyncStatus(spreadsheetID, statusUpdates) Then
            outErrors = outErrors + 1
            MarkPWAFatalSyncError "ImportOneOTKSheet", _
                "WriteBackSyncStatus failed. Local import may have succeeded, but Google Sheet status was not updated. Sheet=" & sheetName
        End If
    End If
    
    LogInfo "ImportOneOTKSheet", sheetName & ": " & outImported & " importiert, " & _
            outSkipped & " preskoceno, " & outErrors & " greske"
    Exit Sub

EH:
    MarkPWAFatalSyncError "ImportOneOTKSheet", _
        "Unexpected error while importing OTK sheet=" & sheetName & _
        "; Error=" & Err.description

    LogErr "ImportOneOTKSheet", "Sheet: " & sheetName
    outErrors = outErrors + 1
End Sub

' ============================================================
' PRIVATE -- Validierung
' ============================================================

Private Function ValidatePWAOtkup(ByVal data As Variant, ByVal row As Long) As String
    ' Prueft Pflichtfelder und Plausibilitaet
    ' Returns "" wenn OK, sonst Fehlermeldung
    
    Dim koopID As String
    Dim vrsta As String
    Dim kolicina As Double
    Dim cena As Double
    
    koopID = Trim$(CStr(data(row, GS_KOOPERANT_ID)))
    vrsta = Trim$(CStr(data(row, GS_VRSTA)))
    
    If Len(koopID) = 0 Then
        ValidatePWAOtkup = "KooperantID missing"
        Exit Function
    End If
    
    If Len(vrsta) = 0 Then
        ValidatePWAOtkup = "VrstaVoca missing"
        Exit Function
    End If
    
    ' KooperantID existiert?
    Dim koopName As Variant
    koopName = LookupValue(TBL_KOOPERANTI, "KooperantID", koopID, "Ime")
    If IsEmpty(koopName) Then
        ValidatePWAOtkup = "KooperantID not found: " & koopID
        Exit Function
    End If
    
    ' Kolicina
    On Error Resume Next
    kolicina = CDbl(data(row, GS_KOLICINA))
    On Error GoTo 0
    If kolicina <= 0 Then
        ValidatePWAOtkup = "Kolicina <= 0"
        Exit Function
    End If
    
    ' Cena
    On Error Resume Next
    cena = CDbl(data(row, GS_CENA))
    On Error GoTo 0
    If cena <= 0 Then
        ValidatePWAOtkup = "Cena <= 0"
        Exit Function
    End If
    
    Dim kolAmb As Long
    Dim tipAmb As String

    tipAmb = Trim$(CStr(nz(data(row, GS_TIP_AMB), "")))

    On Error Resume Next
    kolAmb = CLng(nz(data(row, GS_KOL_AMB), 0))
    On Error GoTo 0

    If kolAmb < 0 Then
        ValidatePWAOtkup = "KolAmbalaze < 0"
        Exit Function
    End If

    If kolAmb > 0 And Len(tipAmb) = 0 Then
        ValidatePWAOtkup = "TipAmbalaze missing while KolAmbalaze > 0"
        Exit Function
    End If
    
    ValidatePWAOtkup = ""
End Function

Private Function IsDuplicateInMaster(ByVal clientRecordID As String) As Boolean
    If Len(Trim$(clientRecordID)) = 0 Then
        LogError "IsDuplicateInMaster", "ClientRecordID je prazan. Duplicate check nije validan."
        IsDuplicateInMaster = True
        Exit Function
    End If
    
    Dim data As Variant
    data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then
        IsDuplicateInMaster = False
        Exit Function
    End If
    
    Dim colCRID As Long
    colCRID = RequireColumnIndex(TBL_OTKUP, "ClientRecordID", "IsDuplicateInMaster")
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If CStr(nz(data(i, colCRID), "")) = clientRecordID Then
            IsDuplicateInMaster = True
            Exit Function
        End If
    Next i
    
    IsDuplicateInMaster = False
End Function

' ============================================================
' PRIVATE -- Import Row
' ============================================================
Private Function ImportRowToTblOtkup_RowTX(ByVal data As Variant, _
                                           ByVal row As Long, _
                                           ByVal clientRecordID As String) As String
    Dim tx As clsTransaction

    On Error GoTo EH

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_AMBALAZA

    ImportRowToTblOtkup_RowTX = ImportRowToTblOtkup(data, row, clientRecordID)

    If Len(Trim$(ImportRowToTblOtkup_RowTX)) = 0 Then
        Err.Raise vbObjectError + 8301, "ImportRowToTblOtkup_RowTX", _
                  "ImportRowToTblOtkup nije vratio OtkupID. ClientRecordID=" & clientRecordID
    End If

    tx.CommitTx
    Exit Function

EH:
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr "ImportRowToTblOtkup_RowTX", "ClientRecordID=" & clientRecordID
    On Error GoTo 0

    ImportRowToTblOtkup_RowTX = ""
End Function

Private Function ImportRowToTblOtkup(ByVal data As Variant, _
                                     ByVal row As Long, _
                                     ByVal clientRecordID As String) As String
    Dim newID As String
    Dim datum As Date
    Dim kooperantID As String
    Dim stanicaID As String
    Dim vrstaVoca As String
    Dim sortaVoca As String
    Dim kolicina As Double
    Dim cena As Double
    Dim tipAmb As String
    Dim kolAmb As Long
    Dim klasa As String
    Dim parcelaID As String
    Dim kulturaID As String
    Dim otkupacID As String
    Dim vozacID As String
    
    On Error GoTo EH
    
    ' Daten auslesen
    kooperantID = Trim$(CStr(data(row, GS_KOOPERANT_ID)))
    vrstaVoca = Trim$(CStr(data(row, GS_VRSTA)))
    sortaVoca = Trim$(CStr(data(row, GS_SORTA)))
    klasa = Trim$(CStr(data(row, GS_KLASA)))
    tipAmb = Trim$(CStr(data(row, GS_TIP_AMB)))
    parcelaID = Trim$(CStr(data(row, GS_PARCELA_ID)))
    otkupacID = Trim$(CStr(data(row, GS_OTKUPAC_ID)))
    vozacID = Trim$(CStr(data(row, GS_VOZAC_ID)))
    
    If Len(klasa) = 0 Then klasa = "I"
    
    ' Datum parsen
    On Error Resume Next
    datum = CDate(data(row, GS_DATUM))
    If Err.Number <> 0 Then datum = Date
    On Error GoTo EH
    
    ' Numerische Werte
    kolicina = CDbl(data(row, GS_KOLICINA))
    cena = CDbl(data(row, GS_CENA))
    
    On Error Resume Next
    kolAmb = CLng(data(row, GS_KOL_AMB))
    On Error GoTo EH
    
    If kolAmb < 0 Then
        Err.Raise vbObjectError + 8100, "ImportRowToTblOtkup", _
              "KolAmbalaze ne sme biti negativan. ClientRecordID=" & clientRecordID
    End If

    If kolAmb > 0 And Len(Trim$(tipAmb)) = 0 Then
        Err.Raise vbObjectError + 8101, "ImportRowToTblOtkup", _
              "TipAmbalaze je obavezan kada je KolAmbalaze > 0. ClientRecordID=" & clientRecordID
    End If
    
    ' Procitaj BrojDokumenta iz OTK sheet-a (PWA-generated, kolona 23)
    Dim brojDokumenta As String
    brojDokumenta = Trim$(CStr(nz(data(row, GS_BROJ_DOKUMENTA), "")))
    
    ' StanicaID aus Kooperant holen
    stanicaID = CStr(nz(LookupValue(TBL_KOOPERANTI, "KooperantID", kooperantID, COL_KOOP_STANICA), ""))
    
    ' Wenn OtkupacID = StanicaID (wie bei deinem Setup), nutze das
    If Len(stanicaID) = 0 And Left$(otkupacID, 3) = "ST-" Then
        stanicaID = otkupacID
    End If
    
    ' KulturaID Lookup
    kulturaID = CStr(nz(LookupValue(TBL_KULTURE, "VrstaVoca", vrstaVoca, "KulturaID"), ""))
    If Len(kulturaID) = 0 Then kulturaID = vrstaVoca & "-" & sortaVoca
    
    ' Fallback: prazno = legacy / PWA pre-rollout.
    ' Validacija formata za PWA-generated brojeve (regex kanonski).
    If Len(brojDokumenta) = 0 Then
        brojDokumenta = GenerateBrojDokumenta(stanicaID, datum)
        If Len(brojDokumenta) = 0 Then
            Err.Raise vbObjectError + 8103, "ImportRowToTblOtkup", _
                "Nije moguce generisati BrojDokumenta. ClientRecordID=" & clientRecordID
        End If
        LogWarn "ImportRowToTblOtkup", _
                "BrojDokumenta fallback-generated lokalno za " & clientRecordID
    Else
        If Not IsValidBrojFormat(brojDokumenta) Then
            Err.Raise vbObjectError + 8104, "ImportRowToTblOtkup", _
                "Invalid BrojDokumenta format: " & brojDokumenta & _
                " (CRID=" & clientRecordID & ")"
        End If
    End If
    
    ' Neue ID
    newID = GetNextID(TBL_OTKUP, COL_OTK_ID, "OTK-")
    If Len(Trim$(newID)) = 0 Then
        Err.Raise vbObjectError + 8302, "ImportRowToTblOtkup", _
              "GetNextID nije vratio OtkupID. ClientRecordID=" & clientRecordID
    End If
    
    ' VozacID
    ' BrojDokumenta = "PWA:" & clientRecordID (fuer Duplikat-Check)
    ' Novac = 0, PrimalacNovca = ""
    
    Dim rowData As Variant
    rowData = Array(newID, datum, kooperantID, stanicaID, kulturaID, _
                    vrstaVoca, sortaVoca, kolicina, cena, tipAmb, _
                    kolAmb, vozacID, brojDokumenta, 0, "", klasa, _
                    "", "", "", "", "", parcelaID, _
                    clientRecordID, "PWA")
    
    Dim result As Long
    result = AppendRow(TBL_OTKUP, rowData)
    
    If result > 0 Then
        ' Ambalaza tracken
        If kolAmb > 0 Then
            ' Dvojni upis: kooperant IZLAZ (razduzenje) + OM/Stanica ULAZ (zaduzenje OM).
            TrackAmbalaza datum, tipAmb, kolAmb, "Izlaz", kooperantID, "Kooperant", , newID, DOK_TIP_OTKUP
            TrackAmbalaza datum, tipAmb, kolAmb, "Ulaz", stanicaID, "Stanica", , newID, DOK_TIP_OTKUP
        End If
        
        LogInfo "ImportRowToTblOtkup", "Importiert: " & newID & " ? PWA:" & clientRecordID & _
                " | " & kooperantID & " | " & vrstaVoca & " " & kolicina & "kg"
        ImportRowToTblOtkup = newID
    Else
        LogError "ImportRowToTblOtkup", "AppendRow fehlgeschlagen fuer PWA:" & clientRecordID
        ImportRowToTblOtkup = ""
    End If
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogErr "ImportRowToTblOtkup", "ClientRecordID: " & clientRecordID
    On Error GoTo 0

    Err.Raise errNum, "ImportRowToTblOtkup", _
              "Source=" & errSrc & " | " & errDesc
End Function

' ============================================================
' PRIVATE -- SyncStatus zurueckschreiben
' ============================================================

Private Function WriteBackSyncStatus(ByVal spreadsheetID As String, _
                                     ByVal updates As Collection) As Boolean
    Const SOURCE As String = "WriteBackSyncStatus"

    Dim accessToken As String
    Dim url As String
    Dim body As String
    Dim http As Object
    Dim i As Long
    Dim update As Variant
    Dim rowNum As Long
    Dim syncStatus As String
    Dim serverRecordID As String
    Dim isFirst As Boolean

    On Error GoTo EH

    If Len(Trim$(spreadsheetID)) = 0 Then
        LogError SOURCE, "spreadsheetID je prazan."
        WriteBackSyncStatus = False
        Exit Function
    End If

    If updates Is Nothing Then
        LogError SOURCE, "updates je Nothing."
        WriteBackSyncStatus = False
        Exit Function
    End If

    If updates.count = 0 Then
        WriteBackSyncStatus = True
        Exit Function
    End If

    accessToken = GetAccessToken()
    If Len(accessToken) = 0 Then
        LogError SOURCE, "Kein Access Token"
        WriteBackSyncStatus = False
        Exit Function
    End If

    body = "{""valueInputOption"":""RAW"",""data"":["
    isFirst = True

    For i = 1 To updates.count
        update = updates(i)

        rowNum = CLng(update(0))
        syncStatus = Trim$(CStr(update(1)))

        If rowNum < 2 Then
            LogError SOURCE, "Invalid row number: " & CStr(rowNum)
            WriteBackSyncStatus = False
            Exit Function
        End If

        If Len(syncStatus) = 0 Then
            LogError SOURCE, "SyncStatus je prazan za row: " & CStr(rowNum)
            WriteBackSyncStatus = False
            Exit Function
        End If

        If Not isFirst Then body = body & ","
        isFirst = False

        ' F = SyncStatus
        body = body & "{""range"":""Sheet1!F" & CStr(rowNum) & """," & _
               """values"":[[""" & JsonEscape(syncStatus) & """]]}"

        ' B = ServerRecordID
        If UBound(update) >= 2 Then
            serverRecordID = Trim$(CStr(update(2)))

            If Len(serverRecordID) > 0 Then
                body = body & ",{""range"":""Sheet1!B" & CStr(rowNum) & """," & _
                       """values"":[[""" & JsonEscape(serverRecordID) & """]]}"
            End If
        End If
    Next i

    body = body & "]}"

    url = "https://sheets.googleapis.com/v4/spreadsheets/" & spreadsheetID & _
          "/values:batchUpdate"

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 10000, 10000, 30000, 30000

    http.Open "POST", url, False
    http.SetRequestHeader "Authorization", "Bearer " & accessToken
    http.SetRequestHeader "Content-Type", "application/json"
    http.Send body

    If http.status >= 200 And http.status < 300 Then
        LogInfo SOURCE, CStr(updates.count) & " Status-Updates geschrieben"
        WriteBackSyncStatus = True
    Else
        LogError SOURCE, _
                 "HTTP " & http.status & ": " & Left$(CStr(http.responseText), 1000), _
                 http.status
        WriteBackSyncStatus = False
    End If

    Exit Function

EH:
    LogErr SOURCE
    WriteBackSyncStatus = False
End Function

Private Function TryUpdateVozacID(ByVal clientRecordID As String, _
                                   ByVal newVozacID As String) As Boolean
    ' Ako Otkup u masteru nema VozacID a sheet ga ima -- updateuj
    Dim data As Variant
    data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function
    
    Dim colCRID As Long, colVoz As Long
    colCRID = GetColumnIndex(TBL_OTKUP, "ClientRecordID")
    colVoz = GetColumnIndex(TBL_OTKUP, COL_OTK_VOZAC)
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If CStr(nz(data(i, colCRID), "")) = clientRecordID Then
            Dim currentVoz As String
            currentVoz = Trim$(CStr(nz(data(i, colVoz), "")))
            If currentVoz = "" And newVozacID <> "" Then
                UpdateCell TBL_OTKUP, i, COL_OTK_VOZAC, newVozacID
                LogInfo "TryUpdateVozacID", "Updated VozacID=" & newVozacID & _
                        " for ClientRecordID=" & clientRecordID
                TryUpdateVozacID = True
            End If
            Exit Function
        End If
    Next i
End Function

' ============================================================
' PRIVATE -- Helpers
' ============================================================

Private Function nz(ByVal v As Variant, Optional ByVal Fallback As Variant = "") As Variant
    If isError(v) Then
        nz = Fallback
    ElseIf IsNull(v) Then
        nz = Fallback
    ElseIf IsEmpty(v) Then
        nz = Fallback
    ElseIf Trim$(CStr(v)) = "" Then
        nz = Fallback
    Else
        nz = v
    End If
End Function

Private Function RequireSingleMasterSyncRow(ByVal tblName As String, _
                                            ByVal idColumn As String, _
                                            ByVal idValue As String, _
                                            ByVal sourceName As String) As Long
    If Len(Trim$(tblName)) = 0 Then
        Err.Raise ERR_MASTER_SYNC_GUARD_BASE + 1, sourceName, _
                  "TableName je obavezan."
    End If

    If Len(Trim$(idColumn)) = 0 Then
        Err.Raise ERR_MASTER_SYNC_GUARD_BASE + 2, sourceName, _
                  "IdColumn je obavezan."
    End If

    If Len(Trim$(idValue)) = 0 Then
        Err.Raise ERR_MASTER_SYNC_GUARD_BASE + 3, sourceName, _
                  "ID vrednost je obavezna. Table=" & tblName & _
                  " Column=" & idColumn
    End If

    RequireColumnIndex tblName, idColumn, sourceName

    Dim rows As Collection
    Set rows = FindRows(tblName, idColumn, idValue)

    If rows Is Nothing Then
        Err.Raise ERR_MASTER_SYNC_GUARD_BASE + 4, sourceName, _
                  "FindRows je vratio Nothing. Table=" & tblName & _
                  " Column=" & idColumn & _
                  " ID=" & idValue
    End If

    If rows.count = 0 Then
        Err.Raise ERR_MASTER_SYNC_GUARD_BASE + 5, sourceName, _
                  "Missing document link. Table=" & tblName & _
                  " Column=" & idColumn & _
                  " ID=" & idValue
    End If

    If rows.count > 1 Then
        Err.Raise ERR_MASTER_SYNC_GUARD_BASE + 6, sourceName, _
                  "Duplicate document key. Table=" & tblName & _
                  " Column=" & idColumn & _
                  " ID=" & idValue & _
                  " Count=" & CStr(rows.count)
    End If

    RequireSingleMasterSyncRow = CLng(rows(1))
End Function

Private Sub LinkOtkupToOtpremnicaStrict(ByVal otkupID As String, _
                                        ByVal otpremnicaID As String, _
                                        ByVal sourceName As String)
    Dim rowOtkup As Long
    Dim rowOtpremnica As Long

    rowOtkup = RequireSingleMasterSyncRow(TBL_OTKUP, COL_OTK_ID, otkupID, sourceName)
    rowOtpremnica = RequireSingleMasterSyncRow(TBL_OTPREMNICA, COL_OTP_ID, otpremnicaID, sourceName)

    RequireUpdateCell TBL_OTKUP, rowOtkup, COL_OTK_OTPREMNICA_ID, _
                      otpremnicaID, sourceName
    ' Faza 7 korak 5: dual-write denorm poslovni kljuc (stabilan kroz re-verziju).
    SetOtkupBrojOtpremnice rowOtkup, otpremnicaID
End Sub

Private Sub LinkOtpremnicaToBrojZbirneStrict(ByVal otpremnicaID As String, _
                                             ByVal brojZbirne As String, _
                                             ByVal sourceName As String)
    If Len(Trim$(brojZbirne)) = 0 Then
        Err.Raise ERR_MASTER_SYNC_GUARD_BASE + 50, sourceName, _
                  "BrojZbirne je obavezan za link Otpremnica -> Zbirna. OtpremnicaID=" & otpremnicaID
    End If

    Dim rowOtpremnica As Long
    rowOtpremnica = RequireSingleMasterSyncRow(TBL_OTPREMNICA, COL_OTP_ID, otpremnicaID, sourceName)

    RequireUpdateCell TBL_OTPREMNICA, rowOtpremnica, COL_OTP_BROJ_ZBIRNE, _
                      brojZbirne, sourceName
End Sub

Private Function GetBrojZbirneForIDStrict(ByVal zbirnaID As String, _
                                          ByVal sourceName As String) As String
    Dim rowZbirna As Long
    rowZbirna = RequireSingleMasterSyncRow(TBL_ZBIRNA, COL_ZBR_ID, zbirnaID, sourceName)

    Dim data As Variant
    data = GetTableData(TBL_ZBIRNA)

    If IsEmpty(data) Then
        Err.Raise ERR_MASTER_SYNC_GUARD_BASE + 40, sourceName, _
                  "Tabela je prazna: " & TBL_ZBIRNA
    End If

    Dim colBroj As Long
    colBroj = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, sourceName)

    GetBrojZbirneForIDStrict = Trim$(CStr(nz(data(rowZbirna, colBroj), "")))

    If Len(GetBrojZbirneForIDStrict) = 0 Then
        Err.Raise ERR_MASTER_SYNC_GUARD_BASE + 41, sourceName, _
                  "BrojZbirne je prazan za ZbirnaID=" & zbirnaID
    End If
End Function
' ============================================================
' modMasterSync -- ZBIRNA IMPORT (dodati u postojeci modMasterSync)
' ============================================================

' ============================================================
' PUBLIC -- Hauptfunktion Zbirna Import
' ============================================================
Public Sub ImportZbirneFromPWA()
    Call ImportZbirneFromPWA_Core(True)
End Sub

Public Function ImportZbirneFromPWA_Core(ByVal showMessages As Boolean) As Boolean
    Dim folderID As String
    Dim sheetIDs As Collection
    Dim sheetNames As Collection
    Dim i As Long
    Dim totalImported As Long
    Dim totalSkipped As Long
    Dim totalErrors As Long
    Dim filesCount As Long

    On Error GoTo EH

    ImportZbirneFromPWA_Core = False
    mLastPWAFatalSyncError = False

    If Not IsGoogleAuthConfigured() Then
        MarkPWAFatalSyncError "ImportZbirneFromPWA_Core", _
            "Google OAuth2 nije konfigurisan."

        Monitor_MasterSyncFail _
            procedureName:="ImportZbirneFromPWA_Core", _
            errNum:=0, _
            errDesc:="Google OAuth2 nije konfigurisan.", _
            errSrc:="modMasterSync.ImportZbirneFromPWA_Core", _
            importedCount:=0, _
            skippedCount:=0, _
            errorCount:=0

        If showMessages Then
            MsgBox "Google OAuth2 nije konfigurisan!", vbCritical, APP_NAME
        End If

        Exit Function
    End If

    folderID = GetConfigValue("GOOGLE_PWA_FOLDER_ID")

    If Len(Trim$(folderID)) = 0 Then
        MarkPWAFatalSyncError "ImportZbirneFromPWA_Core", _
            "GOOGLE_PWA_FOLDER_ID nije postavljen."

        Monitor_MasterSyncFail _
            procedureName:="ImportZbirneFromPWA_Core", _
            errNum:=0, _
            errDesc:="GOOGLE_PWA_FOLDER_ID nije postavljen.", _
            errSrc:="modMasterSync.ImportZbirneFromPWA_Core", _
            importedCount:=0, _
            skippedCount:=0, _
            errorCount:=0

        If showMessages Then
            MsgBox "GOOGLE_PWA_FOLDER_ID nije postavljen!", vbCritical, APP_NAME
        End If

        Exit Function
    End If

    LogInfo "ImportZbirneFromPWA_Core", "Import started."

    Set sheetIDs = New Collection
    Set sheetNames = New Collection

    If Not FindVOZSheets(folderID, sheetIDs, sheetNames) Then
        LogWarn "ImportZbirneFromPWA_Core", _
            "FindVOZSheets failed. Drive list could not be loaded. Retry later."

        Monitor_MasterSyncFail _
            procedureName:="ImportZbirneFromPWA_Core", _
            errNum:=0, _
            errDesc:="FindVOZSheets failed. Drive list could not be loaded.", _
            errSrc:="modMasterSync.ImportZbirneFromPWA_Core", _
            importedCount:=0, _
            skippedCount:=0, _
            errorCount:=1

        If showMessages Then
            MsgBox "Google Drive lista VOZ fajlova nije ucitana." & vbCrLf & _
                   "Proveri konekciju i probaj ponovo.", _
                   vbExclamation, APP_NAME
        End If

        ImportZbirneFromPWA_Core = False
        Exit Function
    End If

    If sheetIDs.count = 0 Then
        Monitor_MasterSyncSuccess _
            procedureName:="ImportZbirneFromPWA_Core", _
            importedCount:=0, _
            skippedCount:=0, _
            errorCount:=0, _
            filesCount:=0

        If showMessages Then
            MsgBox "Nema VOZ-* fajlova u PWA folderu.", vbInformation, APP_NAME
        End If

        LogInfo "ImportZbirneFromPWA_Core", "No VOZ files found."

        ImportZbirneFromPWA_Core = True
        Exit Function
    End If

    filesCount = sheetIDs.count

    For i = 1 To sheetIDs.count
        Dim imported As Long
        Dim skipped As Long
        Dim errors As Long

        imported = 0
        skipped = 0
        errors = 0

        Call ImportOneVOZSheet( _
            CStr(sheetIDs(i)), _
            CStr(sheetNames(i)), _
            imported, _
            skipped, _
            errors)

        totalImported = totalImported + imported
        totalSkipped = totalSkipped + skipped
        totalErrors = totalErrors + errors
    Next i

    LogInfo "ImportZbirneFromPWA_Core", _
        "Import completed. Files=" & CStr(filesCount) & _
        "; Imported=" & CStr(totalImported) & _
        "; Skipped=" & CStr(totalSkipped) & _
        "; Errors=" & CStr(totalErrors)

    If mLastPWAFatalSyncError Then
        Monitor_MasterSyncFail _
            procedureName:="ImportZbirneFromPWA_Core", _
            errNum:=0, _
            errDesc:="Fatal PWA sync error occurred during VOZ/Zbirne import.", _
            errSrc:="modMasterSync.ImportZbirneFromPWA_Core", _
            importedCount:=totalImported, _
            skippedCount:=totalSkipped, _
            errorCount:=totalErrors

        If showMessages Then
            MsgBox Poruka("SYNC_ERR_UVOZ_ZBIRNIH_NIJE") & vbCrLf & _
                   "Uvezeno: " & CStr(totalImported) & vbCrLf & _
                   "Preskoceno: " & CStr(totalSkipped) & vbCrLf & _
                   Poruka("SYNC_ERR_GRESKE") & CStr(totalErrors) & vbCrLf & vbCrLf & _
                   "Proveri log.", _
                   vbCritical, APP_NAME
        End If

        ImportZbirneFromPWA_Core = False
        Exit Function
    End If

    If totalErrors > 0 Then
        Monitor_MasterSyncFail _
            procedureName:="ImportZbirneFromPWA_Core", _
            errNum:=0, _
            errDesc:="VOZ/Zbirne import completed with row-level errors.", _
            errSrc:="modMasterSync.ImportZbirneFromPWA_Core", _
            importedCount:=totalImported, _
            skippedCount:=totalSkipped, _
            errorCount:=totalErrors

        If showMessages Then
            MsgBox Poruka("SYNC_ERR_UVOZ_ZBIRNIH_ZAVRSEN") & vbCrLf & vbCrLf & _
                   "Fajlova: " & CStr(filesCount) & vbCrLf & _
                   "Uvezeno: " & CStr(totalImported) & vbCrLf & _
                   "Preskoceno: " & CStr(totalSkipped) & vbCrLf & _
                   Poruka("SYNC_ERR_GRESKE") & CStr(totalErrors) & vbCrLf & vbCrLf & _
                   "Proveri log.", _
                   vbExclamation, APP_NAME
        End If

        ImportZbirneFromPWA_Core = False
        Exit Function
    End If

    Monitor_MasterSyncSuccess _
        procedureName:="ImportZbirneFromPWA_Core", _
        importedCount:=totalImported, _
        skippedCount:=totalSkipped, _
        errorCount:=totalErrors, _
        filesCount:=filesCount

    If showMessages Then
        MsgBox Poruka("SYNC_ERR_UVOZ_ZBIRNIH_ZAVRSEN_2") & vbCrLf & vbCrLf & _
               "Fajlova: " & CStr(filesCount) & vbCrLf & _
               "Uvezeno: " & CStr(totalImported) & vbCrLf & _
               "Preskoceno: " & CStr(totalSkipped) & vbCrLf & _
               Poruka("SYNC_ERR_GRESKE") & CStr(totalErrors), _
               vbInformation, APP_NAME
    End If

    ImportZbirneFromPWA_Core = True
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next

    MarkPWAFatalSyncError "ImportZbirneFromPWA_Core", errDesc
    LogErr "ImportZbirneFromPWA_Core"

    Monitor_MasterSyncFail _
        procedureName:="ImportZbirneFromPWA_Core", _
        errNum:=errNum, _
        errDesc:=errDesc, _
        errSrc:=errSrc, _
        importedCount:=totalImported, _
        skippedCount:=totalSkipped, _
        errorCount:=totalErrors

    If showMessages Then
        MsgBox Poruka("SYNC_MSG_GRESKA_PRI_UVOZU_2") & errDesc, vbCritical, APP_NAME
    End If

    ImportZbirneFromPWA_Core = False
End Function

Public Sub ImportZbirneFromPWA_TX()
    Const SRC As String = "ImportZbirneFromPWA_TX"

    Dim ok As Boolean

    On Error GoTo EH

    ' IMPORTANT:
    ' Do NOT wrap the whole VOZ batch in one outer clsTransaction.
    '
    ' Reason:
    ' - ImportOneVOZSheet writes Google status updates after local row processing.
    ' - Google writeback cannot be rolled back by clsTransaction.
    ' - Row-level atomicity is already handled by ImportVOZRow_RowTX.
    '
    ' Safe model:
    ' - each VOZ row commits/rolls back through ImportVOZRow_RowTX
    ' - successful rows may be written back as Synced>Master
    ' - failed rows are written back as SyncError
    ' - the full import can still return False / partial if any errors occurred
    ok = ImportZbirneFromPWA_Core(False)

    If Not ok Then
        MsgBox Poruka("SYNC_MSG_UVOZ_ZBIRNIH_ZAVRSEN") & vbCrLf & _
               Poruka("SYNC_MSG_USPESNI_REDOVI_KOJI") & vbCrLf & _
               Poruka("SYNC_MSG_NEUSPESNI_REDOVI_OZNACENI") & vbCrLf & _
               "Proveri log.", _
               vbExclamation, APP_NAME
        Exit Sub
    End If

    MsgBox Poruka("SYNC_MSG_UVOZ_ZBIRNIH_ZAVRSEN_2"), vbInformation, APP_NAME
    Exit Sub

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogErr SRC
    On Error GoTo 0

    MsgBox Poruka("SYNC_MSG_GRESKA_PRI_UVOZU_2") & errDesc, vbCritical, APP_NAME
End Sub

' ============================================================
' PRIVATE -- Find VOZ-* Sheets in Folder
' ============================================================

Private Function FindVOZSheets(ByVal folderID As String, _
                               ByRef outIDs As Collection, _
                               ByRef outNames As Collection) As Boolean
    Const SOURCE As String = "FindVOZSheets"

    Dim accessToken As String
    Dim url As String
    Dim http As Object
    Dim query As String
    Dim responseText As String

    On Error GoTo EH

    If Len(Trim$(folderID)) = 0 Then
        LogError SOURCE, "folderID je prazan."
        FindVOZSheets = False
        Exit Function
    End If

    accessToken = GetAccessToken()
    If Len(accessToken) = 0 Then
        LogError SOURCE, "Kein Access Token"
        FindVOZSheets = False
        Exit Function
    End If

    query = "name contains 'VOZ-' and mimeType='application/vnd.google-apps.spreadsheet'" & _
            " and '" & EscapeDriveQueryValueMasterSync(folderID) & "' in parents and trashed=false"

    url = "https://www.googleapis.com/drive/v3/files" & _
          "?q=" & UrlEncode(query) & _
          "&fields=files(id,name)" & _
          "&pageSize=100"

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 10000, 10000, 30000, 30000

    http.Open "GET", url, False
    http.SetRequestHeader "Authorization", "Bearer " & accessToken
    http.Send

    responseText = CStr(http.responseText)

    If http.status <> 200 Then
        LogError SOURCE, _
                 "HTTP " & http.status & ": " & Left$(responseText, 1000), _
                 http.status
        FindVOZSheets = False
        Exit Function
    End If

    Call ParseFileListVOZ(responseText, outIDs, outNames)

    LogInfo SOURCE, "Gefunden: " & outIDs.count & " VOZ-Sheets"
    FindVOZSheets = True
    Exit Function

EH:
    LogErr SOURCE
    FindVOZSheets = False
End Function
    

Private Sub ParseFileListVOZ(ByVal json As String, _
                              ByRef outIDs As Collection, _
                              ByRef outNames As Collection)
    Dim pos As Long
    Dim fileID As String, fileName As String
    
    pos = 1
    Do
        pos = InStr(pos, json, """id""", vbTextCompare)
        If pos = 0 Then Exit Do
        
        fileID = ExtractJsonValueAt(json, pos)
        
        Dim namePos As Long
        namePos = InStr(pos, json, """name""", vbTextCompare)
        If namePos = 0 Then Exit Do
        
        fileName = ExtractJsonValueAt(json, namePos)
        
        If Len(fileID) > 0 And Len(fileName) > 0 Then
            If Left$(fileName, 4) = "VOZ-" Then
                outIDs.Add fileID
                outNames.Add fileName
            End If
        End If
        
        pos = namePos + 1
    Loop
End Sub

' ============================================================
' PRIVATE -- Import eines einzelnen VOZ-Sheets
' ============================================================

Private Sub ImportOneVOZSheet(ByVal spreadsheetID As String, _
                              ByVal sheetName As String, _
                              ByRef outImported As Long, _
                              ByRef outSkipped As Long, _
                              ByRef outErrors As Long)
    Dim data As Variant
    Dim i As Long
    Dim syncStatus As String
    Dim statusUpdates As Collection
    
    On Error GoTo EH
    
    data = ReadSheetData(spreadsheetID, "Sheet1")
    
    If IsEmpty(data) Then
        LogWarn "ImportOneVOZSheet", "Leeres Sheet: " & sheetName
        Exit Sub
    End If
    
    If Not ValidateVOZSheetHeader(data, sheetName) Then
        outErrors = outErrors + 1
        MarkPWAFatalSyncError "ImportOneVOZSheet", _
            "Import aborted because VOZ header schema is invalid. Sheet=" & sheetName
        Exit Sub
    End If
    
    If UBound(data, 1) < 2 Then
        LogInfo "ImportOneVOZSheet", "Keine Daten in: " & sheetName
        Exit Sub
    End If
    
    Set statusUpdates = New Collection
    
    For i = 2 To UBound(data, 1)
        syncStatus = Trim$(CStr(data(i, VS_SYNC_STATUS)))
        
        If syncStatus = SYNC_STATUS_PENDING Then
            
            Dim clientRecordID As String
            clientRecordID = Trim$(CStr(data(i, VS_CLIENT_RECORD_ID)))
            
            If Len(clientRecordID) = 0 Then
                statusUpdates.Add Array(i, SYNC_STATUS_ERROR & ":ClientRecordID missing", "")
                outErrors = outErrors + 1
                LogWarn "ImportOneVOZSheet", sheetName & " Row " & i & ": ClientRecordID missing. Import skipped."
                GoTo NextImportRow
            End If
            
            If IsDuplicateZbirnaInMaster(clientRecordID) Then
                statusUpdates.Add Array(i, SYNC_STATUS_DUPLICATE, "")
                outSkipped = outSkipped + 1
            Else
                Dim validationError As String
                validationError = ValidatePWAZbirna(data, i)
                
                If Len(validationError) > 0 Then
                    statusUpdates.Add Array(i, SYNC_STATUS_ERROR & ":" & validationError, "")
                    outErrors = outErrors + 1
                    LogWarn "ImportOneVOZSheet", sheetName & " Row " & i & ": " & validationError
                Else
                    Dim newZbirnaID As String
                    Dim brojZbirne As String

                    If ImportVOZRow_RowTX(data, i, clientRecordID, newZbirnaID, brojZbirne) Then
                        If Len(Trim$(newZbirnaID)) = 0 Or Len(Trim$(brojZbirne)) = 0 Then
                            statusUpdates.Add Array(i, SYNC_STATUS_ERROR & ":Invalid row TX result", "")
                            outErrors = outErrors + 1

                            MarkPWAFatalSyncError "ImportOneVOZSheet", _
                                "ImportVOZRow_RowTX returned success but output is invalid. Sheet=" & _
                                sheetName & "; Row=" & CStr(i)
                        Else
                            statusUpdates.Add Array(i, SYNC_STATUS_MASTER, newZbirnaID, brojZbirne)
                            outImported = outImported + 1
                        End If
                    Else
                        statusUpdates.Add Array(i, SYNC_STATUS_ERROR & ":Import/link failed", "")
                        outErrors = outErrors + 1

                        MarkPWAFatalSyncError "ImportOneVOZSheet", _
                            "Import/link failed for VOZ row. Sheet=" & sheetName & _
                            "; Row=" & CStr(i) & _
                             "; ClientRecordID=" & clientRecordID
                    End If
                End If
            End If
        Else
            outSkipped = outSkipped + 1
        End If
NextImportRow:
    Next i
    
    If statusUpdates.count > 0 Then
        If Not WriteBackVOZSyncStatus(spreadsheetID, statusUpdates) Then
            outErrors = outErrors + 1
            MarkPWAFatalSyncError "ImportOneVOZSheet", _
                "WriteBackVOZSyncStatus failed. Sheet=" & sheetName
        End If
    End If
    
    LogInfo "ImportOneVOZSheet", sheetName & ": " & outImported & " importiert, " & _
            outSkipped & " preskoceno, " & outErrors & " greske"
    Exit Sub

EH:
    LogErr "ImportOneVOZSheet", "Sheet: " & sheetName
    outErrors = outErrors + 1
End Sub

Private Function ImportVOZRow_RowTX(ByRef data As Variant, _
                                    ByVal rowIndex As Long, _
                                    ByVal clientRecordID As String, _
                                    ByRef outZbirnaID As String, _
                                    ByRef outBrojZbirne As String) As Boolean
    Const SRC As String = "ImportVOZRow_RowTX"

    Dim tx As clsTransaction
    Dim otkupRecordIDs As String

    On Error GoTo EH

    ImportVOZRow_RowTX = False
    outZbirnaID = vbNullString
    outBrojZbirne = vbNullString

    If Len(Trim$(clientRecordID)) = 0 Then
        Err.Raise ERR_MASTER_SYNC_GUARD_BASE + 60, SRC, _
                  "ClientRecordID je obavezan."
    End If

    Set tx = New clsTransaction
    tx.BeginTx

    tx.AddTableSnapshot TBL_ZBIRNA
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_OTPREMNICA

    outZbirnaID = ImportRowToTblZbirna(data, rowIndex, clientRecordID)

    If Len(Trim$(outZbirnaID)) = 0 Then
        Err.Raise ERR_MASTER_SYNC_GUARD_BASE + 61, SRC, _
                  "ImportRowToTblZbirna nije vratio ZbirnaID. ClientRecordID=" & clientRecordID
    End If

    RequireSingleMasterSyncRow TBL_ZBIRNA, COL_ZBR_ID, outZbirnaID, SRC

    outBrojZbirne = GetBrojZbirneForIDStrict(outZbirnaID, SRC)

    otkupRecordIDs = Trim$(CStr(nz(data(rowIndex, VS_OTKUP_RECORD_IDS), "")))

    LinkZbirnaToOtkupAndOtpremnica outBrojZbirne, otkupRecordIDs

    tx.CommitTx
    Set tx = Nothing

    ImportVOZRow_RowTX = True
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next

    If Not tx Is Nothing Then tx.RollbackTx

    LogErr SRC

    outZbirnaID = vbNullString
    outBrojZbirne = vbNullString

    Debug.Print SRC & " failed. Source=" & errSrc & _
                " Err=" & CStr(errNum) & _
                " Desc=" & errDesc & _
                " ClientRecordID=" & clientRecordID & _
                " Row=" & CStr(rowIndex)

    On Error GoTo 0

    ImportVOZRow_RowTX = False
End Function
' ============================================================
' PRIVATE -- Validierung
' ============================================================

Private Function ValidatePWAZbirna(ByVal data As Variant, ByVal row As Long) As String
    Dim vozacID As String
    Dim kupacID As String
    Dim kolKlI As Double
    Dim kolKlII As Double
    
    vozacID = Trim$(CStr(data(row, VS_VOZAC_ID)))
    kupacID = Trim$(CStr(data(row, VS_KUPAC_ID)))
    
    If Len(vozacID) = 0 Then
        ValidatePWAZbirna = "VozacID missing"
        Exit Function
    End If
    
    If Len(kupacID) = 0 Then
        ValidatePWAZbirna = "KupacID missing"
        Exit Function
    End If
    
    ' KupacID existiert?
    Dim kupacName As Variant
    kupacName = LookupValue(TBL_KUPCI, "KupacID", kupacID, "Naziv")
    If IsEmpty(kupacName) Then
        ValidatePWAZbirna = "KupacID not found: " & kupacID
        Exit Function
    End If
    
    ' Mindestens eine Klasa muss Kolicina > 0 haben
    On Error Resume Next
    kolKlI = CDbl(data(row, VS_KOLICINA_KL_I))
    kolKlII = CDbl(data(row, VS_KOLICINA_KL_II))
    On Error GoTo 0
    
    If kolKlI <= 0 And kolKlII <= 0 Then
        ValidatePWAZbirna = "Kolicina KlI + KlII <= 0"
        Exit Function
    End If
    
    ValidatePWAZbirna = ""
End Function

Private Function IsDuplicateZbirnaInMaster(ByVal clientRecordID As String) As Boolean
    If Len(Trim$(clientRecordID)) = 0 Then
        IsDuplicateZbirnaInMaster = False
        Exit Function
    End If
    
    Dim data As Variant
    data = GetTableData(TBL_ZBIRNA)
    If IsEmpty(data) Then
        IsDuplicateZbirnaInMaster = False
        Exit Function
    End If
    
    Dim colCRID As Long
    colCRID = GetColumnIndex(TBL_ZBIRNA, "ClientRecordID")
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If CStr(nz(data(i, colCRID), "")) = clientRecordID Then
            IsDuplicateZbirnaInMaster = True
            Exit Function
        End If
    Next i
    
    IsDuplicateZbirnaInMaster = False
End Function

' ============================================================
' PRIVATE -- Import Row to tblZbirna
' ============================================================

Private Function ImportRowToTblZbirna(ByVal data As Variant, _
                                      ByVal row As Long, _
                                      ByVal clientRecordID As String) As String
    Dim newID As String
    Dim datum As Date
    Dim vozacID As String
    Dim brojZbirne As String
    Dim kupacID As String
    Dim vrstaVoca As String
    Dim sortaVoca As String
    Dim ukupnoKol As Double
    Dim tipAmb As String
    Dim kolAmb As Long
    Dim kolKlI As Double
    Dim kolKlII As Double
    
    On Error GoTo EH
    
    vozacID = Trim$(CStr(data(row, VS_VOZAC_ID)))
    kupacID = Trim$(CStr(data(row, VS_KUPAC_ID)))
    vrstaVoca = Trim$(CStr(data(row, VS_VRSTA)))
    sortaVoca = Trim$(CStr(data(row, VS_SORTA)))
    tipAmb = Trim$(CStr(nz(data(row, VS_TIP_AMB), "")))
    
    ' Datum
    On Error Resume Next
    datum = CDate(data(row, VS_DATUM))
    If Err.Number <> 0 Then datum = Date
    On Error GoTo EH
    
    ' Kolicine po klasi
    On Error Resume Next
    kolKlI = CDbl(data(row, VS_KOLICINA_KL_I))
    kolKlII = CDbl(data(row, VS_KOLICINA_KL_II))
    kolAmb = CLng(data(row, VS_KOL_AMB))
    If kolAmb < 0 Then
        LogError "ImportRowToTblZbirna", _
             "KolAmbalaze ne sme biti negativan. CRID=" & clientRecordID
        ImportRowToTblZbirna = ""
        Exit Function
    End If
    On Error GoTo EH
    
    ukupnoKol = kolKlI + kolKlII
    
    If kolAmb > 0 And Len(Trim$(tipAmb)) = 0 Then
        LogError "ImportRowToTblZbirna", _
             "TipAmbalaze je obavezan kada je KolAmbalaze > 0. CRID=" & clientRecordID
        ImportRowToTblZbirna = ""
        Exit Function
    End If
    
    ' Procitaj BrojZbirne iz VOZ sheet-a (PWA-generated, kolona 20)
    brojZbirne = Trim$(CStr(nz(data(row, VS_BROJ_ZBIRNE), "")))

    ' Fallback: prazno znaci legacy zapis ili PWA pre-rollout-a
    If Len(brojZbirne) = 0 Then
        brojZbirne = GenerateBrojZbirne(vozacID, datum)
        If Len(brojZbirne) = 0 Then
            LogError "ImportRowToTblZbirna", "Nije moguce generisati BrojZbirne za VozacID=" & vozacID
            ImportRowToTblZbirna = ""
            Exit Function
        End If
        LogWarn "ImportRowToTblZbirna", "BrojZbirne fallback-generated lokalno za " & clientRecordID
    Else
        ' Validacija formata za PWA-generated broj
        If Not IsValidBrojZbirneFormat(brojZbirne) Then
            LogError "ImportRowToTblZbirna", "Invalid BrojZbirne format: " & brojZbirne & " (CRID=" & clientRecordID & ")"
            ImportRowToTblZbirna = ""
            Exit Function
        End If
    End If
    
    ' Hladnjaca iz KupacID
    Dim hladnjaca As String
    hladnjaca = CStr(nz(LookupValue(TBL_KUPCI, "KupacID", kupacID, "Hladnjaca"), ""))
    
    newID = GetNextID(TBL_ZBIRNA, COL_ZBR_ID, "ZBR-")
    If Len(Trim$(newID)) = 0 Then
        LogError "ImportRowToTblZbirna", _
             "GetNextID nije vratio ZbirnaID. CRID=" & clientRecordID
        ImportRowToTblZbirna = ""
        Exit Function
    End If
    ' tblZbirna Schema:
    ' tblZbirna Schema:
    ' ZbirnaID | Datum | VozacID | BrojZbirne | KupacID | Hladnjaca | Pogon |
    ' VrstaVoca | SortaVoca | UkupnoKolicina | TipAmbalaze | UkupnoAmbalaze | Klasa |
    ' Stornirano | ClientRecordID | SyncSource
    '
    ' Klasa: Ako ima obe klase, pisi "I/II". Ako samo jedna, pisi tu.
    Dim klasa As String
    If kolKlI > 0 And kolKlII > 0 Then
        klasa = "I/II"
    ElseIf kolKlII > 0 Then
        klasa = "II"
    Else
        klasa = "I"
    End If
    
    Dim rowData As Variant
    rowData = Array(newID, datum, vozacID, brojZbirne, kupacID, _
                    hladnjaca, "", vrstaVoca, sortaVoca, _
                    ukupnoKol, tipAmb, kolAmb, klasa, _
                    "", clientRecordID, "PWA")
    
    Dim result As Long
    result = AppendRow(TBL_ZBIRNA, rowData)
    
    If result > 0 Then
        ' Generacija (nasledjuje se od aktivnih redova istog BrojZbirne, inace nova)
        ' -- PWA import ne sme da ostavi red bez generacije, prefill je cita.
        ApplyGeneracijaID TBL_ZBIRNA, result, COL_ZBR_BROJ, brojZbirne

        LogInfo "ImportRowToTblZbirna", "Importiert: " & newID & " BrojZbirne=" & brojZbirne & _
                " | " & vozacID & " | " & kupacID & " | " & ukupnoKol & "kg"
        ImportRowToTblZbirna = newID
    Else
        LogError "ImportRowToTblZbirna", "AppendRow fehlgeschlagen fuer PWA:" & clientRecordID
        ImportRowToTblZbirna = ""
    End If
    Exit Function

EH:
    LogErr "ImportRowToTblZbirna", "ClientRecordID: " & clientRecordID
    ImportRowToTblZbirna = ""
End Function

' ============================================================
' PRIVATE -- Kaskadno povezivanje Zbirna -> Otpremnice -> Otkupi
' ============================================================

Private Sub LinkZbirnaToOtkupAndOtpremnica(ByVal brojZbirne As String, _
                                           ByVal otkupRecordIDs As String)
    Const SRC As String = "LinkZbirnaToOtkupAndOtpremnica"

    On Error GoTo EH

    If Len(Trim$(brojZbirne)) = 0 Then
        Err.Raise ERR_MASTER_SYNC_GUARD_BASE + 30, SRC, _
                  "BrojZbirne je obavezan za kaskadno povezivanje."
    End If

    If Len(Trim$(otkupRecordIDs)) = 0 Then
        LogInfo SRC, "Nema Otkup ClientRecordID vrednosti za BrojZbirne=" & brojZbirne
        Exit Sub
    End If

    Dim colCRID As Long
    Dim colOtkID As Long
    Dim colOtkOtpID As Long

    colCRID = RequireColumnIndex(TBL_OTKUP, MASTER_SYNC_CLIENT_RECORD_ID_COL, SRC)
    colOtkID = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, SRC)
    colOtkOtpID = RequireColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID, SRC)

    RequireColumnIndex TBL_OTKUP, COL_OTK_BROJ_ZBIRNE, SRC
    RequireColumnIndex TBL_OTPREMNICA, COL_OTP_ID, SRC
    RequireColumnIndex TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, SRC

    Dim crIDs() As String
    crIDs = Split(otkupRecordIDs, ",")

    Dim updatedOtp As Object
    Set updatedOtp = CreateObject("Scripting.Dictionary")

    Dim c As Long
    For c = 0 To UBound(crIDs)
        Dim searchCRID As String
        searchCRID = Trim$(CStr(crIDs(c)))

        If Len(searchCRID) > 0 Then
            Dim rowOtkup As Long
            rowOtkup = RequireSingleMasterSyncRow(TBL_OTKUP, MASTER_SYNC_CLIENT_RECORD_ID_COL, searchCRID, SRC)

            Dim otkData As Variant
            otkData = GetTableData(TBL_OTKUP)

            If IsEmpty(otkData) Then
                Err.Raise ERR_MASTER_SYNC_GUARD_BASE + 31, SRC, _
                          "Tabela je prazna: " & TBL_OTKUP
            End If

            Dim otkupID As String
            otkupID = Trim$(CStr(otkData(rowOtkup, colOtkID)))

            rowOtkup = RequireSingleMasterSyncRow(TBL_OTKUP, COL_OTK_ID, otkupID, SRC)

            otkData = GetTableData(TBL_OTKUP)

            RequireUpdateCell TBL_OTKUP, rowOtkup, COL_OTK_BROJ_ZBIRNE, brojZbirne, SRC

            Dim otpID As String
            otpID = Trim$(CStr(nz(otkData(rowOtkup, colOtkOtpID), "")))

            If Len(otpID) > 0 Then
                If Not updatedOtp.Exists(otpID) Then
                    LinkOtpremnicaToBrojZbirneStrict otpID, brojZbirne, SRC
                    updatedOtp.Add otpID, True
                End If
            End If
        End If
    Next c

    LogInfo SRC, "BrojZbirne=" & brojZbirne & _
                 " linked " & CStr(UBound(crIDs) + 1) & _
                 " otkupa, " & CStr(updatedOtp.count) & " otpremnica"

    Exit Sub

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.description
End Sub
' ============================================================
' PRIVATE -- Helper: BrojZbirne aus ZbirnaID
' ============================================================

Private Function GetBrojZbirneForID(ByVal zbirnaID As String) As String
    Dim val As Variant
    val = LookupValue(TBL_ZBIRNA, COL_ZBR_ID, zbirnaID, COL_ZBR_BROJ)
    If Not IsEmpty(val) Then
        GetBrojZbirneForID = CStr(val)
    Else
        GetBrojZbirneForID = ""
    End If
End Function

Private Function IsValidBrojZbirneFormat(ByVal s As String) As Boolean
    ' Format: [S]x/ddmmyy ili [S]x/ddmmyy-N ("S" = mirror-stanica kao vozac).
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = "^S?\d+/\d{6}(-\d+)?$"
    re.Global = False
    IsValidBrojZbirneFormat = re.Test(s)
End Function

' ============================================================
' PRIVATE -- WriteBack VOZ SyncStatus + ServerRecordID
' ============================================================

Private Function WriteBackVOZSyncStatus(ByVal spreadsheetID As String, _
                                        ByVal updates As Collection) As Boolean
    ' Isti pattern kao WriteBackSyncStatus za OTK
    ' Kolona F = SyncStatus, Kolona B = ServerRecordID
    
    Dim accessToken As String
    Dim url As String
    Dim body As String
    Dim http As Object
    Dim i As Long
    Dim update As Variant
    
    On Error GoTo EH
    
    accessToken = GetAccessToken()
    If Len(accessToken) = 0 Then Exit Function
    
    body = "{""valueInputOption"":""RAW"",""data"":["
    
    Dim isFirst As Boolean
    isFirst = True
    
    For i = 1 To updates.count
        update = updates(i)
        
        If Not isFirst Then body = body & ","
        isFirst = False
        
        ' Kolona F -- SyncStatus
        body = body & "{""range"":""Sheet1!F" & CStr(update(0)) & """," & _
               """values"":[[""" & JsonEscape(CStr(update(1))) & """]]}"
        
        ' Kolona B -- ServerRecordID (2. kolona = B)
        If UBound(update) >= 2 Then
            If Len(CStr(update(2))) > 0 Then
                body = body & ",{""range"":""Sheet1!B" & CStr(update(0)) & """," & _
                       """values"":[[""" & JsonEscape(CStr(update(2))) & """]]}"
            End If
        End If
        
        ' T = BrojZbirne
        If UBound(update) >= 3 Then
            If Len(CStr(update(3))) > 0 Then
                body = body & ",{""range"":""Sheet1!T" & CStr(update(0)) & """," & _
                    """values"":[[""" & JsonEscape(CStr(update(3))) & """]]}"
            End If
        End If
    Next i
    
    body = body & "]}"
    
    url = "https://sheets.googleapis.com/v4/spreadsheets/" & spreadsheetID & _
          "/values:batchUpdate"
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 10000, 10000, 30000, 30000
    
    http.Open "POST", url, False
    http.SetRequestHeader "Authorization", "Bearer " & accessToken
    http.SetRequestHeader "Content-Type", "application/json"
    http.Send body
    
    If http.status >= 200 And http.status < 300 Then
        LogInfo "WriteBackVOZSyncStatus", CStr(updates.count) & " Status-Updates geschrieben"
        WriteBackVOZSyncStatus = True
    Else
        LogError "WriteBackVOZSyncStatus", "HTTP " & http.status & ": " & http.responseText, http.status
        WriteBackVOZSyncStatus = False
    End If

    Exit Function

EH:
    LogErr "WriteBackVOZSyncStatus"
    WriteBackVOZSyncStatus = False
End Function

Private Function GenerateBrojZbirne(ByVal vozacID As String, ByVal datum As Date) As String
    Dim vozacBroj As String
    vozacBroj = ExtractNumericVozacBroj(vozacID)
    
    If Len(vozacBroj) = 0 Then
        GenerateBrojZbirne = ""
        Exit Function
    End If
    
    Dim baza As String
    baza = vozacBroj & "/" & Format$(datum, "ddmmyy")
    
    Dim data As Variant
    data = GetTableData(TBL_ZBIRNA)
    
    Dim seq As Long
    seq = 1
    
    If Not IsEmpty(data) Then
        Dim colDat As Long, colVoz As Long
        colDat = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_DATUM)
        colVoz = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_VOZAC)
        
        Dim i As Long
        For i = 1 To UBound(data, 1)
            If CStr(data(i, colVoz)) = vozacID Then
                If Format$(CDate(data(i, colDat)), "ddmmyy") = Format$(datum, "ddmmyy") Then
                    seq = seq + 1
                End If
            End If
        Next i
    End If
    
    If seq = 1 Then
        GenerateBrojZbirne = baza
    Else
        GenerateBrojZbirne = baza & "-" & seq
    End If

    ' Mirror-stanica (VozacID==StanicaID) -> "S" prefiks (razdvaja od realnih vozaca).
    GenerateBrojZbirne = ApplyMirrorPrefix(vozacID, GenerateBrojZbirne)
End Function

Private Function ValidateVOZSheetHeader(ByVal data As Variant, _
                                        ByVal sheetName As String) As Boolean
    Const SOURCE As String = "ValidateVOZSheetHeader"

    On Error GoTo EH

    If IsEmpty(data) Then
        LogError SOURCE, "Sheet data is Empty: " & sheetName
        ValidateVOZSheetHeader = False
        Exit Function
    End If

    If UBound(data, 1) < 1 Then
        LogError SOURCE, "Sheet nema header row: " & sheetName
        ValidateVOZSheetHeader = False
        Exit Function
    End If

    If UBound(data, 2) < VS_BROJ_ZBIRNE Then
        LogError SOURCE, _
             "VOZ schema drift: premalo kolona u sheetu " & sheetName & _
             ". ExpectedAtLeast=" & CStr(VS_BROJ_ZBIRNE) & _
             ", Actual=" & CStr(UBound(data, 2))
        ValidateVOZSheetHeader = False
        Exit Function
    End If

    If Not RequireVOZHeaderValue(data, sheetName, VS_CLIENT_RECORD_ID, "ClientRecordID") Then Exit Function
    If Not RequireVOZHeaderValue(data, sheetName, VS_SERVER_RECORD_ID, "ServerRecordID") Then Exit Function
    If Not RequireVOZHeaderValue(data, sheetName, VS_CREATED_AT, "CreatedAtClient") Then Exit Function
    If Not RequireVOZHeaderValue(data, sheetName, VS_UPDATED_AT_CLIENT, "UpdatedAtClient") Then Exit Function
    If Not RequireVOZHeaderValue(data, sheetName, VS_UPDATED_AT_SERVER, "UpdatedAtServer") Then Exit Function
    If Not RequireVOZHeaderValue(data, sheetName, VS_SYNC_STATUS, "SyncStatus") Then Exit Function
    If Not RequireVOZHeaderValue(data, sheetName, VS_VOZAC_ID, "VozacID") Then Exit Function
    If Not RequireVOZHeaderValue(data, sheetName, VS_DATUM, "Datum") Then Exit Function
    If Not RequireVOZHeaderValue(data, sheetName, VS_KUPAC_ID, "KupacID") Then Exit Function
    If Not RequireVOZHeaderValue(data, sheetName, VS_KUPAC_NAME, "KupacName") Then Exit Function
    If Not RequireVOZHeaderValue(data, sheetName, VS_VRSTA, "VrstaVoca") Then Exit Function
    If Not RequireVOZHeaderValue(data, sheetName, VS_SORTA, "SortaVoca") Then Exit Function
    If Not RequireVOZHeaderValue(data, sheetName, VS_KOLICINA_KL_I, "KolicinaKlI") Then Exit Function
    If Not RequireVOZHeaderValue(data, sheetName, VS_KOLICINA_KL_II, "KolicinaKlII") Then Exit Function
    If Not RequireVOZHeaderValue(data, sheetName, VS_TIP_AMB, "TipAmbalaze") Then Exit Function
    If Not RequireVOZHeaderValue(data, sheetName, VS_KOL_AMB, "KolAmbalaze") Then Exit Function
    If Not RequireVOZHeaderValue(data, sheetName, VS_KLASA, "Klasa") Then Exit Function
    If Not RequireVOZHeaderValue(data, sheetName, VS_OTKUP_RECORD_IDS, "OtkupRecordIDs") Then Exit Function
    If Not RequireVOZHeaderValue(data, sheetName, VS_RECEIVED_AT, "ReceivedAt") Then Exit Function
    If Not RequireVOZHeaderValue(data, sheetName, VS_BROJ_ZBIRNE, "BrojZbirne") Then Exit Function

    ValidateVOZSheetHeader = True
    Exit Function

EH:
    LogErr SOURCE, "Sheet: " & sheetName
    ValidateVOZSheetHeader = False
End Function

Private Function RequireVOZHeaderValue(ByVal data As Variant, _
                                       ByVal sheetName As String, _
                                       ByVal colIndex As Long, _
                                       ByVal expectedHeader As String) As Boolean
    Dim actualHeader As String

    actualHeader = Trim$(CStr(data(1, colIndex)))

    If StrComp(actualHeader, expectedHeader, vbBinaryCompare) <> 0 Then
        LogError "ValidateVOZSheetHeader", _
                 "VOZ schema drift in " & sheetName & _
                 ". Col=" & CStr(colIndex) & _
                 ", Expected='" & expectedHeader & "'" & _
                 ", Actual='" & actualHeader & "'"
        RequireVOZHeaderValue = False
        Exit Function
    End If

    RequireVOZHeaderValue = True
End Function

Private Function ExtractNumericVozacBroj(ByVal vozacID As String) As String
    Dim i As Long, ch As String, digits As String
    
    For i = 1 To Len(vozacID)
        ch = Mid$(vozacID, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
        End If
    Next i
    
    If Len(digits) = 0 Then
        ExtractNumericVozacBroj = ""
    Else
        ExtractNumericVozacBroj = CStr(CLng(digits))
    End If
End Function

Private Sub MarkPWAFatalSyncError(ByVal sourceName As String, ByVal message As String)
    mLastPWAFatalSyncError = True
    LogError sourceName, message
End Sub

Private Sub Monitor_MasterSyncSuccess(ByVal procedureName As String, _
                                      ByVal importedCount As Long, _
                                      ByVal skippedCount As Long, _
                                      ByVal errorCount As Long, _
                                      ByVal filesCount As Long)
    On Error Resume Next

    Monitor_Event _
        eventType:="MASTERDATA_SYNC_SUCCESS", _
        severity:="INFO", _
        message:="Master sync completed. Files=" & CStr(filesCount) & _
                 "; Imported=" & CStr(importedCount) & _
                 "; Skipped=" & CStr(skippedCount) & _
                 "; Errors=" & CStr(errorCount), _
        userId:="Operator", _
        moduleName:="modMasterSync", _
        procedureName:=procedureName, _
        entityType:="MasterData", _
        entityID:="PWA-OTKUP", _
        correlationId:="MASTERDATA-SYNC-PWA"
End Sub

Private Sub Monitor_MasterSyncFail(ByVal procedureName As String, _
                                   ByVal errNum As Long, _
                                   ByVal errDesc As String, _
                                   ByVal errSrc As String, _
                                   Optional ByVal importedCount As Long = 0, _
                                   Optional ByVal skippedCount As Long = 0, _
                                   Optional ByVal errorCount As Long = 0)
    On Error Resume Next

    Monitor_Error _
        moduleName:="modMasterSync", _
        procedureName:=procedureName, _
        entityType:="MasterData", _
        entityID:="PWA-OTKUP", _
        correlationId:="MASTERDATA-SYNC-PWA", _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="MASTERDATA_SYNC_FAIL", _
        severity:="CRITICAL", _
        message:="Master sync failed. Imported=" & CStr(importedCount) & _
                 "; Skipped=" & CStr(skippedCount) & _
                 "; Errors=" & CStr(errorCount) & _
                 "; Error=" & errDesc, _
        userId:="Operator", _
        moduleName:="modMasterSync", _
        procedureName:=procedureName, _
        entityType:="MasterData", _
        entityID:="PWA-OTKUP", _
        correlationId:="MASTERDATA-SYNC-PWA"
End Sub

' ============================================================
' PARCEL GEO PULL -- Google/Stammdaten -> tblParcele
' ============================================================

Public Function ImportParcelGeoFromGoogleToMaster() As Boolean
    Const SRC As String = "ImportParcelGeoFromGoogleToMaster"

    Dim sheetID As String
    Dim folderID As String
    Dim data As Variant
    Dim parcelData As Variant
    Dim tx As clsTransaction

    Dim cParID As Long
    Dim cPolygon As Long
    Dim cLat As Long
    Dim cLongitude As Long
    Dim cGeoStatus As Long
    Dim cGeoSource As Long
    Dim cN As Long
    Dim cEasting As Long
    Dim cMeteo As Long
    Dim cRizik As Long
    Dim cDatumGeo As Long
    Dim cDatumAzur As Long
    Dim cNapomena As Long

    Dim mParID As Long
    Dim mPolygon As Long
    Dim mLat As Long
    Dim mLongitude As Long
    Dim mGeoStatus As Long
    Dim mGeoSource As Long
    Dim mN As Long
    Dim mEasting As Long
    Dim mMeteo As Long
    Dim mRizik As Long
    Dim mDatumGeo As Long
    Dim mDatumAzur As Long
    Dim mNapomena As Long

    Dim i As Long
    Dim parcelaID As String
    Dim rows As Collection
    Dim masterRow As Long
    Dim changedFields As Long
    Dim updatedParcels As Long
    Dim skippedRows As Long
    Dim missingParcels As Long
    Dim seen As Object

    On Error GoTo EH

    ImportParcelGeoFromGoogleToMaster = False

    If Not IsGoogleAuthConfigured() Then
        LogError SRC, "Google OAuth2 nije konfigurisan."
        Exit Function
    End If

    sheetID = Trim$(GetConfigValue("GOOGLE_STAMMDATEN_SHEET_ID"))

    If Len(sheetID) = 0 Then
        folderID = Trim$(GetConfigValue("GOOGLE_PWA_FOLDER_ID"))

        If Len(folderID) = 0 Then
            LogError SRC, "GOOGLE_STAMMDATEN_SHEET_ID i GOOGLE_PWA_FOLDER_ID nisu postavljeni."
            Exit Function
        End If

        sheetID = GetSpreadsheetID("Stammdaten", folderID)

        If Len(sheetID) > 0 Then
            Call SetConfigValue("GOOGLE_STAMMDATEN_SHEET_ID", sheetID)
        End If
    End If

    If Len(sheetID) = 0 Then
        LogError SRC, "Stammdaten Google Sheet nije pronaden."
        Exit Function
    End If

    data = ReadSheetData(sheetID, "Parcele")

    If IsEmpty(data) Then
        LogError SRC, "Google Stammdaten/Parcele tab je prazan ili nije ucitan."
        Exit Function
    End If

    If UBound(data, 1) < 1 Then
        LogError SRC, "Google Stammdaten/Parcele nema header row."
        Exit Function
    End If

    cParID = GeoHeaderIndex(data, COL_PAR_ID)
    cPolygon = GeoHeaderIndex(data, COL_PAR_POLYGON)
    cLat = GeoHeaderIndex(data, COL_PAR_LAT)
    cLongitude = GeoHeaderIndex(data, COL_PAR_LNG)
    cGeoStatus = GeoHeaderIndex(data, COL_PAR_GEO_STATUS)
    cGeoSource = GeoHeaderIndex(data, COL_PAR_GEO_SOURCE)
    cN = GeoHeaderIndex(data, COL_PAR_N)
    cEasting = GeoHeaderIndex(data, COL_PAR_E)
    cMeteo = GeoHeaderIndex(data, COL_PAR_METEO)
    cRizik = GeoHeaderIndex(data, COL_PAR_RIZIK)
    cDatumGeo = GeoHeaderIndex(data, COL_PAR_DATUM_GEO)
    cDatumAzur = GeoHeaderIndex(data, COL_PAR_DATUM_AZUR)
    cNapomena = GeoHeaderIndex(data, COL_PAR_NAPOMENA)

    If cParID = 0 Then
        LogError SRC, "Google Parcele sheet nema header: " & COL_PAR_ID
        Exit Function
    End If

    If cPolygon = 0 And cLat = 0 And cLongitude = 0 Then
        LogError SRC, "Google Parcele sheet nema geo kolone: " & _
                      COL_PAR_POLYGON & "/" & COL_PAR_LAT & "/" & COL_PAR_LNG
        Exit Function
    End If

    parcelData = GetTableData(TBL_PARCELE)
    If IsEmpty(parcelData) Then
        LogError SRC, "tblParcele je prazan. Geo pull nema gde da upise podatke."
        Exit Function
    End If

    mParID = RequireColumnIndex(TBL_PARCELE, COL_PAR_ID, SRC)
    mPolygon = RequireColumnIndex(TBL_PARCELE, COL_PAR_POLYGON, SRC)
    mLat = RequireColumnIndex(TBL_PARCELE, COL_PAR_LAT, SRC)
    mLongitude = RequireColumnIndex(TBL_PARCELE, COL_PAR_LNG, SRC)
    mGeoStatus = RequireColumnIndex(TBL_PARCELE, COL_PAR_GEO_STATUS, SRC)
    mGeoSource = RequireColumnIndex(TBL_PARCELE, COL_PAR_GEO_SOURCE, SRC)
    mN = RequireColumnIndex(TBL_PARCELE, COL_PAR_N, SRC)
    mEasting = RequireColumnIndex(TBL_PARCELE, COL_PAR_E, SRC)
    mMeteo = RequireColumnIndex(TBL_PARCELE, COL_PAR_METEO, SRC)
    mRizik = RequireColumnIndex(TBL_PARCELE, COL_PAR_RIZIK, SRC)
    mDatumGeo = RequireColumnIndex(TBL_PARCELE, COL_PAR_DATUM_GEO, SRC)
    mDatumAzur = RequireColumnIndex(TBL_PARCELE, COL_PAR_DATUM_AZUR, SRC)
    mNapomena = RequireColumnIndex(TBL_PARCELE, COL_PAR_NAPOMENA, SRC)

    Set seen = CreateObject("Scripting.Dictionary")

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_PARCELE

    If UBound(data, 1) < 2 Then
        tx.CommitTx
        LogInfo SRC, "Google Parcele sheet ima samo header. Nema geo redova za import."
        ImportParcelGeoFromGoogleToMaster = True
        Exit Function
    End If

    For i = 2 To UBound(data, 1)
        parcelaID = Trim$(GeoText(data(i, cParID)))

        If Len(parcelaID) = 0 Then
            skippedRows = skippedRows + 1
            GoTo NextGeoRow
        End If

        If seen.Exists(parcelaID) Then
            Err.Raise vbObjectError + 8601, SRC, _
                      "Dupli ParcelaID u Google Parcele sheet-u: " & parcelaID
        End If
        seen.Add parcelaID, True

        If Not GeoRowHasAnyValue(data, i, cPolygon, cLat, cLongitude, cGeoStatus, cGeoSource, _
                                 cN, cEasting, cMeteo, cRizik, cDatumGeo, cDatumAzur, cNapomena) Then
            skippedRows = skippedRows + 1
            GoTo NextGeoRow
        End If

        Set rows = FindRows(TBL_PARCELE, COL_PAR_ID, parcelaID)

        If rows Is Nothing Then
            missingParcels = missingParcels + 1
            LogWarn SRC, "ParcelaID iz Google sheet-a nije pronaden u tblParcele: " & parcelaID
            GoTo NextGeoRow
        End If

        If rows.count = 0 Then
            missingParcels = missingParcels + 1
            LogWarn SRC, "ParcelaID iz Google sheet-a nije pronaden u tblParcele: " & parcelaID
            GoTo NextGeoRow
        End If

        If rows.count <> 1 Then
            Err.Raise vbObjectError + 8602, SRC, _
                      "ParcelaID nije jedinstven u tblParcele: " & parcelaID & _
                      "; Count=" & CStr(rows.count)
        End If

        masterRow = CLng(rows(1))

        changedFields = 0

        If cPolygon > 0 Then GeoUpdateFieldIfNeeded parcelData, masterRow, mPolygon, COL_PAR_POLYGON, data(i, cPolygon), changedFields
        If cLat > 0 Then GeoUpdateFieldIfNeeded parcelData, masterRow, mLat, COL_PAR_LAT, data(i, cLat), changedFields
        If cLongitude > 0 Then GeoUpdateFieldIfNeeded parcelData, masterRow, mLongitude, COL_PAR_LNG, data(i, cLongitude), changedFields
        If cGeoStatus > 0 Then GeoUpdateFieldIfNeeded parcelData, masterRow, mGeoStatus, COL_PAR_GEO_STATUS, data(i, cGeoStatus), changedFields
        If cGeoSource > 0 Then GeoUpdateFieldIfNeeded parcelData, masterRow, mGeoSource, COL_PAR_GEO_SOURCE, data(i, cGeoSource), changedFields
        If cN > 0 Then GeoUpdateFieldIfNeeded parcelData, masterRow, mN, COL_PAR_N, data(i, cN), changedFields
        If cEasting > 0 Then GeoUpdateFieldIfNeeded parcelData, masterRow, mEasting, COL_PAR_E, data(i, cEasting), changedFields
        If cMeteo > 0 Then GeoUpdateFieldIfNeeded parcelData, masterRow, mMeteo, COL_PAR_METEO, data(i, cMeteo), changedFields
        If cRizik > 0 Then GeoUpdateFieldIfNeeded parcelData, masterRow, mRizik, COL_PAR_RIZIK, data(i, cRizik), changedFields
        If cDatumGeo > 0 Then GeoUpdateFieldIfNeeded parcelData, masterRow, mDatumGeo, COL_PAR_DATUM_GEO, data(i, cDatumGeo), changedFields
        If cDatumAzur > 0 Then GeoUpdateFieldIfNeeded parcelData, masterRow, mDatumAzur, COL_PAR_DATUM_AZUR, data(i, cDatumAzur), changedFields
        If cNapomena > 0 Then GeoUpdateFieldIfNeeded parcelData, masterRow, mNapomena, COL_PAR_NAPOMENA, data(i, cNapomena), changedFields

        If changedFields > 0 Then
            updatedParcels = updatedParcels + 1
            LogInfo SRC, "Geo updated ParcelaID=" & parcelaID & _
                         "; ChangedFields=" & CStr(changedFields)
        Else
            skippedRows = skippedRows + 1
        End If

NextGeoRow:
    Next i

    tx.CommitTx

    LogInfo SRC, "Geo pull completed. UpdatedParcels=" & CStr(updatedParcels) & _
                 "; SkippedRows=" & CStr(skippedRows) & _
                 "; MissingParcels=" & CStr(missingParcels)

    ImportParcelGeoFromGoogleToMaster = True
    Exit Function

EH:
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    LogErr SRC
    ImportParcelGeoFromGoogleToMaster = False
End Function

Private Function GeoHeaderIndex(ByVal data As Variant, ByVal headerName As String) As Long
    Dim j As Long
    Dim actual As String

    On Error GoTo EH

    If IsEmpty(data) Then Exit Function
    If UBound(data, 1) < 1 Then Exit Function

    For j = LBound(data, 2) To UBound(data, 2)
        actual = Trim$(GeoText(data(1, j)))

        If StrComp(actual, headerName, vbTextCompare) = 0 Then
            GeoHeaderIndex = j
            Exit Function
        End If
    Next j

    GeoHeaderIndex = 0
    Exit Function

EH:
    GeoHeaderIndex = 0
End Function

Private Function GeoText(ByVal value As Variant) As String
    On Error GoTo EH

    If isError(value) Then
        GeoText = ""
    ElseIf IsNull(value) Then
        GeoText = ""
    ElseIf IsEmpty(value) Then
        GeoText = ""
    Else
        GeoText = CStr(value)
    End If

    Exit Function

EH:
    GeoText = ""
End Function

Private Function GeoHasValue(ByVal data As Variant, _
                             ByVal rowIndex As Long, _
                             ByVal colIndex As Long) As Boolean
    On Error GoTo EH

    If colIndex <= 0 Then
        GeoHasValue = False
        Exit Function
    End If

    GeoHasValue = (Len(Trim$(GeoText(data(rowIndex, colIndex)))) > 0)
    Exit Function

EH:
    GeoHasValue = False
End Function

Private Function GeoRowHasAnyValue(ByVal data As Variant, _
                                   ByVal rowIndex As Long, _
                                   ParamArray cols() As Variant) As Boolean
    Dim i As Long
    Dim colIndex As Long

    On Error GoTo EH

    For i = LBound(cols) To UBound(cols)
        colIndex = CLng(cols(i))

        If GeoHasValue(data, rowIndex, colIndex) Then
            GeoRowHasAnyValue = True
            Exit Function
        End If
    Next i

    GeoRowHasAnyValue = False
    Exit Function

EH:
    GeoRowHasAnyValue = False
End Function

Private Sub GeoUpdateFieldIfNeeded(ByVal parcelData As Variant, _
                                   ByVal masterRow As Long, _
                                   ByVal masterCol As Long, _
                                   ByVal colName As String, _
                                   ByVal newValue As Variant, _
                                   ByRef changedFields As Long)
    Const SRC As String = "GeoUpdateFieldIfNeeded"

    Dim oldText As String
    Dim newText As String

    On Error GoTo EH

    newText = Trim$(GeoText(newValue))

    ' VAZNO:
    ' Prazan Google value NE sme da obrise postojeci lokalni geo podatak.
    If Len(newText) = 0 Then Exit Sub

    oldText = Trim$(GeoText(parcelData(masterRow, masterCol)))

    If StrComp(oldText, newText, vbBinaryCompare) <> 0 Then
        RequireUpdateCell TBL_PARCELE, masterRow, colName, newValue, SRC
        changedFields = changedFields + 1
    End If

    Exit Sub

EH:
    Err.Raise Err.Number, SRC, Err.description
End Sub
' ============================================================
' TEST
' ============================================================

Public Sub Test_ImportOtkupFromPWA()
    Call ImportOtkupFromPWA
End Sub

