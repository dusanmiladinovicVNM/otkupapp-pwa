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

' TEST SEAM (isti obrazac kao modAutoHladnjaca.mTestFailStep): armira se iz smoke
' suite-a, potrosi se pri PRVOJ upotrebi -- vazi samo za taj jedan poziv. U
' produkciji je uvek prazan.
'
' Postoji zato sto se UpdateCell i WriteSheetData ne mogu naterati da padnu
' "prirodno" (kolone/tabela postoje, HTTP radi), a bas ti fail-putevi su ono sto
' AUD-042(a) i AUD-042(c) popravljaju -- bez seam-a bi ostali netestirani.
Private mTestFailSeam As String

Private Const SYNC_STATUS_PENDING As String = "Synced"
Private Const SYNC_STATUS_MASTER As String = "Synced>Master"
Private Const SYNC_STATUS_ERROR As String = "SyncError"
Private Const SYNC_STATUS_DUPLICATE As String = "Duplicate"

Private Const ERR_MASTER_SYNC_GUARD_BASE As Long = vbObjectError + 3900
Private Const MASTER_SYNC_CLIENT_RECORD_ID_COL As String = "ClientRecordID"

' AUD-042(a): ishodi VozacID update-a nad postojecim (duplikat) redom.
' Pozivalac MORA da razlikuje "nema sta da se menja" (bezopasno -> Duplicate) od
' "upis pao" / "red ima DRUGOG vozaca" / "reda nema" -- ta tri su greske i ne
' smeju da dobiju terminalni Duplicate status uz zelen sync.

' AUD-043(b): dozvoljena razlika u danima izmedju otkupa i zbirne. 0 bi odbilo
' legitiman utovar posle ponoci; neogranicena tolerancija bi dozvolila da otkup
' od pre vise meseci ude u novu zbirnu. 1 dan = post-midnight pravilo.
Private Const MASTER_SYNC_MEMBERSHIP_DAY_TOLERANCE As Long = 1


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

' Tab stavki otkupa u OTK-* sheet-u stanice (S1c, REFAKTOR S14.8 t. 13).
Public Const OTK_STAVKE_TAB As String = "OTK_STAVKE"

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

    Call ImportOtkupSheetLoop(sheetIDs, sheetNames, _
                              totalImported, totalSkipped, totalErrors)

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

    LogErr "ImportOtkupFromPWA_Core"
    On Error Resume Next

    MarkPWAFatalSyncError "ImportOtkupFromPWA_Core", errDesc

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


Private Sub ImportOtkupSheetLoop(ByVal sheetIDs As Collection, _
                                 ByVal sheetNames As Collection, _
                                 ByRef totalImported As Long, _
                                 ByRef totalSkipped As Long, _
                                 ByRef totalErrors As Long)
    ' AUD-002: sheetovi se obraduju NEZAVISNO.
    ' Nema batch transakcije oko petlje -- pad kasnijeg sheeta ne sme da
    ' ponisti vec uvezene redove ranijih sheetova, jer njihov Google
    ' writeback (Synced>Master) ne moze da se rollback-uje.
    Dim i As Long
    Dim imported As Long
    Dim skipped As Long
    Dim errors As Long

    For i = 1 To sheetIDs.count
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
End Sub

Public Sub TestHook_ImportOtkupSheetLoop(ByVal sheetIDs As Collection, _
                                         ByVal sheetNames As Collection, _
                                         ByRef totalImported As Long, _
                                         ByRef totalSkipped As Long, _
                                         ByRef totalErrors As Long)
    ' DEV/SMOKE TEST HOOK ONLY.
    ' Isti kod koji vrti ImportOtkupFromPWA_Core, samo nad zadatom listom
    ' fixture sheetova -- da se cross-sheet ponasanje (AUD-002) moze
    ' proveriti bez skeniranja celog PWA foldera.

    Call ImportOtkupSheetLoop(sheetIDs, sheetNames, _
                              totalImported, totalSkipped, totalErrors)
End Sub

Public Sub ImportOtkupFromPWA_TX()
    Dim ok As Boolean

    On Error GoTo EH

    ' IMPORTANT (AUD-002):
    ' Do NOT wrap the whole OTK batch in one outer clsTransaction.
    '
    ' Reason:
    ' - ImportOneOTKSheet writes Google status updates (WriteBackSyncStatus)
    '   after local row processing, per sheet.
    ' - Google writeback cannot be rolled back by clsTransaction.
    ' - An outer rollback triggered by a LATER sheet would delete locally
    '   imported rows of EARLIER sheets, while those Google rows stay
    '   Synced>Master -> next cycle skips them -> permanent data loss.
    ' - Row-level atomicity is already handled by ImportRowToTblOtkup_RowTX.
    '
    ' Safe model (same as ImportZbirneFromPWA_TX):
    ' - each OTK row commits/rolls back through ImportRowToTblOtkup_RowTX
    ' - successful rows may be written back as Synced>Master
    ' - failed rows are written back as SyncError
    ' - the full import can still return False / partial if any errors occurred
    ok = ImportOtkupFromPWA_Core(False)

    If Not ok Then
        Monitor_MasterSyncFail _
            procedureName:="ImportOtkupFromPWA_TX", _
            errNum:=0, _
            errDesc:="PWA import was not confirmed. Partial import kept; failed rows marked SyncError.", _
            errSrc:="modMasterSync.ImportOtkupFromPWA_TX"

        MsgBox Poruka("SYNC_MSG_PWA_UVOZ_NIJE"), _
            vbCritical, APP_NAME
        Exit Sub
    End If

    MsgBox Poruka("SYNC_MSG_PWA_UVOZ_ZAVRSEN"), vbInformation, APP_NAME
    Exit Sub

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    LogErr "ImportOtkupFromPWA_TX"
    On Error Resume Next


    Monitor_MasterSyncFail _
        procedureName:="ImportOtkupFromPWA_TX", _
        errNum:=errNum, _
        errDesc:=errDesc, _
        errSrc:=errSrc

    MsgBox Poruka("SYNC_MSG_GRESKA_PRI_UVOZU") & errDesc, vbCritical, APP_NAME
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

        ' AUD-001: neuspeo Drive lookup NE SME da se procita kao "ne postoji" --
        ' inace bi ova masovna putanja napravila duplikat OTK sheeta za svaku
        ' stanicu u ciklusu.
        If Not TryGetSpreadsheetID(sheetName, folderID, existingID) Then
            failedCount = failedCount + 1
            LogError SRC, _
                "Drive lookup nije uspeo -- OTK sheet se NE kreira (rizik od duplikata). Sheet=" & sheetName & _
                "; StanicaID=" & stanicaID
            GoTo NextStanica
        End If

        If Len(Trim$(existingID)) > 0 Then
            existingCount = existingCount + 1
            GoTo NextStanica
        End If

        newID = CreateOTKSheetWithHeader(sheetName, folderID, headers, SRC)

        If Len(Trim$(newID)) = 0 Then
            failedCount = failedCount + 1
            LogError SRC, _
                "OTK sheet nije kreiran sa headerom. Sheet=" & sheetName & _
                "; StanicaID=" & stanicaID & _
                "; Naziv=" & stanicaNaziv
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
' CreateOTKSheetWithHeader
'
' AUD-042(c): kreiranje sheeta i upis header-a su JEDNA operacija iz perspektive
' sledeceg run-a. Ako header padne, sheet postoji na Drive-u i TryGetSpreadsheetID
' ga po imenu nalazi kao "existing" -> preskoci se, header NIKAD ne bude upisan, a
' PWA pise u sheet bez sheme. Zato pao header znaci: sheet u Drive trash (Drive
' upit filtrira trashed=false, pa sledeci run pravi cist novi).
'
' Vraca SpreadsheetID na uspeh, "" na bilo koji neuspeh (pozivalac broji failed).
'======================================================================
Private Function CreateOTKSheetWithHeader(ByVal sheetName As String, _
                                         ByVal folderID As String, _
                                         ByVal headers As Variant, _
                                         ByVal sourceName As String) As String
    Dim newID As String
    newID = CreateSpreadsheet(sheetName, folderID)

    If Len(Trim$(newID)) = 0 Then
        LogError sourceName, "CreateSpreadsheet failed. Sheet=" & sheetName
        Exit Function
    End If

    Dim headerOk As Boolean

    If ConsumeFailSeam("OTK_HEADER") Then
        headerOk = False
    Else
        headerOk = WriteSheetData(newID, "Sheet1", headers)
    End If

    If headerOk Then
        CreateOTKSheetWithHeader = newID
        Exit Function
    End If

    LogError sourceName, _
        "WriteSheetData header failed. Sheet=" & sheetName & _
        "; SpreadsheetID=" & newID

    If DriveTrashFile(newID) Then
        LogWarn sourceName, _
            "Poison OTK sheet poslat u Drive trash (header nije upisan). Sheet=" & sheetName & _
            "; SpreadsheetID=" & newID
    Else
        LogError sourceName, _
            "Poison OTK sheet NIJE obrisan -- obrisi ga rucno na Drive-u pre sledeceg run-a. Sheet=" & sheetName & _
            "; SpreadsheetID=" & newID
    End If
End Function

'======================================================================
' KOLONE OTK-* SHEET-A -- JEDINO MESTO (REFAKTOR S14.8 t. 13, nalaz E-5)
'
' OtkZaglavljeKolone: tab Sheet1, red po otkupu. Raspored je PWA ugovor do S5:
' PWA ga puni, ImportOneOTKSheet ga cita poziciono (GS_*). Klasa, Kolicina,
' Cena i KolAmbalaze su u njemu jos samo zato sto PWA salje jednu klasu po
' zapisu; VBA push ih ostavlja PRAZNE i pise stavke u OTK_STAVKE.
'
' OtkStavkeKolone: tab OTK_STAVKE, red po stavci; roditelj je OtkupID
' (= ServerRecordID zaglavlja).
'
' Graditelji redova (modStanicaLock, izvoz OtkupiAllStavke) slazu vrednosti PO
' IMENU iz ovih spiskova, ne po poziciji.
'======================================================================
Public Function OtkZaglavljeKolone() As Variant
    OtkZaglavljeKolone = Array( _
        "ClientRecordID", "ServerRecordID", "CreatedAtClient", "UpdatedAtClient", _
        "UpdatedAtServer", "SyncStatus", "DeviceID", "OtkupacID", "Datum", _
        "KooperantID", "KooperantName", "VrstaVoca", "SortaVoca", "Klasa", _
        "Kolicina", "Cena", "TipAmbalaze", "KolAmbalaze", "ParcelaID", "VozacID", _
        "Napomena", "ReceivedAt", "BrojDokumenta")
End Function

Public Function OtkStavkeKolone() As Variant
    OtkStavkeKolone = Array(COL_OKS_ID, COL_OKS_OTKUP_ID, COL_OKS_RB, COL_OKS_KLASA, _
                            COL_OKS_KOLICINA, COL_OKS_CENA, COL_OKS_KOL_AMB, COL_OKS_BRUTO)
End Function

' Red taba OTK_STAVKE za stavku i iz modOtkup.StavkeOtkupaRedovi, po imenu.
Private Function OtkStavkaPolje(ByVal s As Variant, ByVal i As Long, _
                                ByVal kolona As String) As Variant
    Select Case kolona
        Case COL_OKS_ID: OtkStavkaPolje = s(i, 7)
        Case COL_OKS_OTKUP_ID: OtkStavkaPolje = s(i, 1)
        Case COL_OKS_RB: OtkStavkaPolje = s(i, 2)
        Case COL_OKS_KLASA: OtkStavkaPolje = s(i, 3)
        Case COL_OKS_KOLICINA: OtkStavkaPolje = s(i, 4)
        Case COL_OKS_CENA: OtkStavkaPolje = s(i, 5)
        Case COL_OKS_KOL_AMB: OtkStavkaPolje = s(i, 6)
        Case COL_OKS_BRUTO: OtkStavkaPolje = s(i, 8)
        Case Else
            Err.Raise vbObjectError + 8141, "OtkStavkaPolje", _
                      "Kolona OTK_STAVKE bez izvora: " & kolona
    End Select
End Function

' Stavke otkupa kao redovi taba OTK_STAVKE: OtkupID -> Collection 0-based
' nizova u rasporedu OtkStavkeKolone.
'
' Cita KANONSKU granicu (modOtkup.StavkeOtkupaRedovi): zaglavlje bez stavki,
' nevazeca stavka ili dupli OtkupID padaju ovde po imenu -- izvoz ne sme da
' posalje zaglavlje bez stavki kao dokument od nula kilograma.
Public Function OtkStavkeRedoviPoOtkupu() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Set OtkStavkeRedoviPoOtkupu = dict

    Dim s As Variant
    s = modOtkup.StavkeOtkupaRedovi()
    If Not IsArray(s) Then Exit Function

    Dim kol As Variant, i As Long, k As Long, red() As Variant, oid As String
    kol = OtkStavkeKolone()
    For i = 1 To UBound(s, 1)
        ReDim red(0 To UBound(kol) - LBound(kol))
        For k = LBound(kol) To UBound(kol)
            red(k - LBound(kol)) = OtkStavkaPolje(s, i, CStr(kol(k)))
        Next k
        oid = CStr(s(i, 1))
        If Not dict.Exists(oid) Then dict.Add oid, New Collection
        dict(oid).Add red
    Next i
End Function

' Header red za kreiranje OTK-* sheet-a (oblik 1 x N za WriteSheetData).
Private Function BuildOTKOperationalHeaders_() As Variant
    Dim kol As Variant, headers() As Variant, k As Long
    kol = OtkZaglavljeKolone()
    ReDim headers(1 To 1, 1 To UBound(kol) - LBound(kol) + 1)
    For k = LBound(kol) To UBound(kol)
        headers(1, k - LBound(kol) + 1) = kol(k)
    Next k
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
' AUTO-ZBIRNA VISE NIJE DEO PAUZIRANOG LANCA (S4-4).
'
' IzvedeniLanacIzPwaDostupan je JEDNA kapija nad celim izvedenim lancem, i to
' je bilo tacno dok je auto-zbirna pisala Otkup.BrojZbirne nazad na zaglavlje
' (BackfillOtkupBrojZbirneByOtpremnica). Kanonska auto-zbirna taj backlink NE
' pise -- clanstvo je zapis u tblZbirnaIzvori -- pa razlog za zajednicku kapiju
' za NJU vise ne postoji. VOZ/zbirna uvoz ga i dalje ima i ostaje pauziran.
'
' Razdvajanje je namerno i glasno: kapija koja pokriva vise nego sto mora
' zaustavlja i ono sto je popravljeno, a onda se otvara "u paketu" -- tiho
' pustajuci i ono sto nije.
Public Function AutoZbirnaDostupna() As Boolean
    AutoZbirnaDostupna = IsMalinaMode()
End Function

' JEZGRO: auto-zbirna za JEDNU otpremnicu. Vraca ZbirnaID, "" = nista nije
' napravljeno (nije malina, otpremnica nije slobodna/izdata, ili je greska --
' razlog je tada u outGreska).
'
' ZASTO SE ODLUKA "DA LI SME" NE RACUNA OVDE: pita se
' modDokumenta.NevezaneOtpremnice, ISTA lista koju operater vidi u F2 radnom
' stolu. Drugo pravilo na ovom mestu znacilo bi da ekran i automatika mogu da se
' raziju -- automatika bi vezala otpremnicu koju spisak ne nudi, ili obrnuto.
' Ta lista vec drzi sva tri uslova: IZDATA, nestornirana, bez aktivnog clanstva.
'
' IDEMPOTENTNO PO KONSTRUKCIJI: otpremnica koja vec ima zbirnu nije u toj listi,
' pa ponovljen poziv ne pravi drugu. To nije udobnost nego uslov -- jezgro zovu
' DVA pozivaoca (izdavanje i batch prolaz), a batch ume da stigne prvi.
Public Function AutoZbirnaZaOtpremnicu(ByVal otpremnicaID As String, _
                                       Optional ByRef outGreska As String) As String
    Const SRC As String = "AutoZbirnaZaOtpremnicu"

    outGreska = ""
    If Not AutoZbirnaDostupna() Then Exit Function

    otpremnicaID = Trim$(otpremnicaID)
    If Len(otpremnicaID) = 0 Then Exit Function

    ' JEZGRO NE DIZE GRESKU, VEC JE VRACA (review #383, P2).
    '
    ' Pozivalac na izdavanju je vec COMMIT-ovao otpremnicu. Izuzetak koji odavde
    ' izleti stize u njegov EH i tamo postaje "otpremnica nije izdata" -- laz o
    ' poslovnom dogadjaju koji se desio. AutoZbirnaUpis pritom die na vise mesta
    ' PRE pisca (prazan kupac, prazan vozac, nema broja), pa to nije teorijski
    ' put nego najverovatniji.
    '
    ' Ugovor je zato: "" + prazan outGreska = nije bilo posla; "" + neprazan
    ' outGreska = posla je bilo i NIJE uspeo; ZbirnaID = uspelo.
    On Error GoTo EH

    Dim slobodne As Object
    Set slobodne = modDokumenta.NevezaneOtpremnice()
    If Not slobodne.Exists(UCase$(otpremnicaID)) Then Exit Function

    AutoZbirnaZaOtpremnicu = AutoZbirnaUpis(otpremnicaID, outGreska)
    Exit Function
EH:
    ' Opis PRE LogErr-a, isti razlog kao u batch prolazu.
    Dim errDesc As String
    errDesc = Err.description
    LogErr SRC
    AutoZbirnaZaOtpremnicu = ""
    If Len(outGreska) = 0 Then outGreska = errDesc
End Function

' Upis same zbirne. Odvojeno od AutoZbirnaZaOtpremnicu zato sto batch prolaz vec
' ZNA da je otpremnica slobodna (dobio ju je iz iste liste) -- da zove javni
' ulaz, citao bi celu tabelu po otpremnici.
Private Function AutoZbirnaUpis(ByVal otpremnicaID As String, _
                                ByRef outGreska As String) As String
    Const SRC As String = "AutoZbirnaUpis"

    Dim kupacID As String
    kupacID = Trim$(GetConfigValue(CFG_MALINA_DEFAULT_KUPAC))
    If Len(kupacID) = 0 Then
        Err.Raise vbObjectError + 8300, SRC, _
            "MALINA_DEFAULT_KUPAC nije postavljen (kljuc u tblSEFConfig)."
    End If

    Dim datum As Date, vozacID As String
    datum = CDate(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpremnicaID, COL_OTP_DATUM))
    vozacID = Trim$(NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpremnicaID, _
                                         COL_OTP_VOZAC)))
    If Len(vozacID) = 0 Then
        Err.Raise vbObjectError + 8302, SRC, _
            "Otpremnica " & otpremnicaID & " nema VozacID, a zbirna ga trazi."
    End If

    ' ZBIRNA DOBIJA SVOJ BROJ, NE NASLEDJUJE OTPREMNICIN (odluka 23.09.2026).
    '
    ' Stari kod je pisao ApplyMirrorPrefix(vozacID, BrojOtpremnice) i sam
    ' komentar je to vodio kao DUG: numericki deo je tada pripadao STANICI, a
    ' vlasnik niza zbirne je VOZAC. Prolazilo je samo dok je vozac doslovno
    ' mirror-stanica; cim PWA posalje realnog vozaca, kapija konteksta
    ' (modBrojevi) odbija upis kao TUDJ_VLASNIK.
    '
    ' SuggestNextBroj sam primenjuje mirror prefiks "S" kad je vozac ogledalo
    ' stanice, pa se izgled broja u malini ne menja -- menja se to CIJI je niz.
    ' Sidro duga: Test_BKTX_ZbirnaTudjegVlasnikaOdbijena.
    ' checkRemote:=False -- NE PITA GOOGLE.
    '
    ' Podrazumevano SuggestNextBroj gleda i udaljeni list (VOZ-<vozac>), da
    ' predlog operateru ne bi pogodio broj koji je PWA vec potrosio. Ovde to ne
    ' valja iz dva razloga: batch prolaz bi pravio mrezni poziv PO DOKUMENTU, a
    ' automatika koja zavisi od mreze pada kad mreze nema -- tiho, jer
    ' SuggestNextBroj gresku guta i vraca prazno.
    '
    ' Bezbedno je jer je VOZ/zbirna uvoz PAUZIRAN (S5): nijedna PWA zbirna danas
    ' ne stize u tblZbirna, pa lokalni niz jeste ceo niz. DUG ZA S5: kad se uvoz
    ' vrati, udaljena osa se mora vratiti u racun -- ili ovde, ili tako sto uvoz
    ' rezervise svoj opseg.
    Dim brZbirne As String
    brZbirne = modBrojevi.SuggestNextBroj(KIND_ZBR, vozacID, datum, False)
    If Len(brZbirne) = 0 Then
        ' Imenuj OBA uzroka: prazan predlog znaci ili ugasen auto-broj u
        ' Podesavanjima, ili pad generatora. Automatika nema operatera koji bi
        ' broj ukucao, pa oba znace isto -- stani.
        Err.Raise vbObjectError + 8303, SRC, _
            "Nema slobodnog broja zbirne za vozaca " & vozacID & _
            " (provera: auto-broj dokumenta u Podesavanjima)."
    End If

    Dim h As Object
    Set h = CreateObject("Scripting.Dictionary")
    h.Add "Datum", datum
    h.Add "VozacID", vozacID
    h.Add "BrojZbirne", brZbirne
    h.Add "KupacID", kupacID

    Dim hladnjaca As String
    hladnjaca = Trim$(NzToText(LookupValue(TBL_KUPCI, COL_KUP_ID, kupacID, "Hladnjaca")))
    If Len(hladnjaca) > 0 Then h.Add "Hladnjaca", hladnjaca

    ' VRSTA, SORTA I TIP AMBALAZE SE NE SALJU: cinjenica robe se preuzima od
    ' prvog izvora (ZBR-KANON-04). Poslati ih znacilo bi tvrditi ono sto pisac
    ' sam izvodi -- i razici se s njim cim se izvor promeni.
    Dim izvori As Collection
    Set izvori = New Collection
    izvori.Add otpremnicaID

    ' CreateZbirnaIzIzvora_TX, ne CreateZbirna_TX: automatski tok NEMA nezavisno
    ' ocekivanje da unakrsno proveri. To se kaze izborom ulaza, ne izostavljanjem
    ' argumenta (modDokumenta, 686).
    AutoZbirnaUpis = modDokumenta.CreateZbirnaIzIzvora_TX(h, izvori, outGreska)
End Function

' ============================================================
' MALINA MOD -- C: AUTO-OTPREMNICA IZ UVEZENIH PWA OTKUPA (S5-1).
'
' Zamenjuje DVA zatecena koraka ciklusa odjednom:
'
'   2b  "VozacID := StanicaID na tblOtkup" -- pecat koji je otkup pripremao za
'       grupisanje. U novom modelu vozac je cinjenica ZAGLAVLJA OTPREMNICE, a
'       Otkup.VozacID kolona koja umire; pecat zato nema sta da pripremi.
'   3   auto-otpremnica (obrisana u S1c jer je citala linijska polja zaglavlja
'       otkupa i pisala Otkup.OtpremnicaID).
'
' GRUPISANJE VISE NE NOSI KLASU. Staro je bilo StanicaID|Datum|VozacID|Klasa --
' jedna otpremnica PO KLASI, jer je klasa bila polje zaglavlja. Klasa je sada
' stavka, pa dva bloka I i II klase istog dana sa istog otkupnog mesta idu u
' JEDAN dokument sa dve stavke. To je bas ono sto refaktor tvrdi, i zato se meri
' testom, ne komentarom.
'
' Kljuc grupe je ono sto zaglavlje otpremnice NOSI i sto pisac zahteva da bude
' isto za sve izvore (OtpRequireIzvorValjan): StanicaID, Datum, KulturaID i
' TipAmbalaze. Vozac nije deo kljuca jer je u malini POSLEDICA stanice
' (ogledalo), ne nezavisan podatak.
'
' TipAmbalaze je u kljucu iako ga pisac poredi samo za izvore koji stvarno nose
' gajbe. Posledica: otkup sa deklarisanim tipom a bez ijedne gajbe dobija svoju
' grupu umesto da se pridruzi tudjoj. To je namerno STROZE od minimuma --
' proizvodi eventualno jedan dokument vise, nikad dokument sa pogresnim tipom,
' i nikad upis koji pisac odbije.
'
' Self-gated: u visnji ne radi nista.
' ============================================================

' Sopstvena kapija, po uzoru na AutoZbirnaDostupna (S4-4): kapija koja pokriva
' vise nego sto mora zaustavlja i ono sto je popravljeno.
Public Function AutoOtpremnicaDostupna() As Boolean
    AutoOtpremnicaDostupna = IsMalinaMode()
End Function

' BATCH PROLAZ: auto-otpremnica za sve nevezane izdate otkupe.
'
' PAD JEDNE GRUPE NE OBARA OSTALE, ali se ni ne precutkuje (lekcija iz #383:
' delimican uspeh prijavljen kao potpun pad je laz o poslovnom dogadjaju, a
' prijavljen kao potpun uspeh je gora laz). Svaka grupa je svoja transakcija --
' CreateOtpremnicaIzIzvora_TX je otvara i zatvara -- pa ono sto je proslo JESTE
' upisano. Ugovor je zato:
'
'   povratna vrednost  = broj STVARNO napravljenih otpremnica
'   outGreske          = imenovani razlozi za grupe koje nisu prosle, "; " spojeni
'
' Prazan outGreske uz 0 napravljenih znaci "nije bilo posla". Neprazan znaci
' "posla je bilo i deo NIJE uspeo" -- orkestrator taj korak prijavljuje kao pao,
' ali ciklus ne obara, jer su otkupi uvezeni i to je stvaran napredak.
'
' PAD GRUPE I PAD PROLAZA NISU ISTA STVAR, pa se i ne prijavljuju isto.
'
' Razliku NE pogadja ovaj prolaz iz teksta greske -- nju izrice mesto podizanja
' (modSchemaGuard.RaiseSistemski) i prenosi je pisac (outSistemska). Ovde se
' samo postupa po njoj:
'
'   POSLOVNO ODBIJANJE -- zavisi od podataka TE grupe: imenuj razlog i probaj
'     sledecu. Dokument je jedini gubitnik.
'   SISTEMSKI PAD -- sema nije spremna, kolona nedostaje, AppendRow nije upisao,
'     ili je pukao sam VBA. Oborice i svaku sledecu grupu, pa je dalje
'     pokusavanje samo gomilanje istog razloga: prolaz STAJE i greska ide gore.
'
' Bez toga bi sistemski pad izasao kao "deo otkupa je ostao bez otpremnice", a
' ciklus bi posle stvarnog kvara masine nastavio na outbound sync (review #385).
'
' samoOtkupID suzava prolaz na grupu KOJOJ TAJ OTKUP PRIPADA (identitet, ne
' labela). Grupa se ne sece: otpremnica od dela svoje grupe bila bi drugaciji
' dokument od onog koji pun prolaz pravi.
Public Function AutoCreateOtpremniceFromPWA_TX(Optional ByVal samoOtkupID As String = "", _
                                               Optional ByRef outGreske As String) As Long
    Const SRC As String = "AutoCreateOtpremniceFromPWA_TX"

    outGreske = ""
    On Error GoTo EH

    If Not AutoOtpremnicaDostupna() Then Exit Function

    Dim grupe As Object
    Set grupe = GrupeZaAutoOtpremnicu(Trim$(samoOtkupID))
    If grupe.count = 0 Then Exit Function

    Dim k As Variant, otpID As String, g As String, n As Long
    Dim greske As Collection
    Set greske = New Collection

    For Each k In grupe.Keys
        g = ""
        otpID = AutoOtpremnicaUpis(grupe(k), g)
        If Len(otpID) > 0 Then
            n = n + 1
        Else
            greske.Add CStr(k) & ": " & IIf(Len(g) > 0, g, "nepoznat razlog")
        End If
    Next k


    AutoCreateOtpremniceFromPWA_TX = n
    outGreske = SpojiRazloge(greske)

    If n > 0 Then LogInfo SRC, "Malina auto-otpremnica created=" & CStr(n)
    If Len(outGreske) > 0 Then LogWarn SRC, "Grupe bez otpremnice: " & outGreske
    Exit Function

EH:
    ' Broj, opis i izvor se citaju PRE LogErr-a -- LogErr usput brise stanje
    ' greske, pa bi re-raise posle njega bio Err.Raise 0 (#383, P2).
    Dim errNum As Long, errDesc As String, errSrc As String
    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    LogErr SRC

    AutoCreateOtpremniceFromPWA_TX = 0
    Err.Raise errNum, SRC, "Source=" & errSrc & " | " & errDesc
End Function

' Nevezani izdati otkupi -> Dictionary "kljuc grupe" -> Collection OtkupID-eva.
'
' "Nevezan" se NE racuna ovde: modDokumenta.NevezaniOtkupi je jedini citac koji
' na to pitanje odgovara, i isti koji radni sto u F1 koristi. Lokalna kopija bi
' umela da ponudi otkup koji pisac smatra zauzetim.
'
' Izdatost se filtrira, ne relaksira: pisac i dalje odbija neizdat izvor. Filter
' postoji da nacrt ili pokvaren red ne obori grupu kojoj ionako ne pripada.
Private Function GrupeZaAutoOtpremnicu(ByVal samoOtkupID As String) As Object
    Const SRC As String = "GrupeZaAutoOtpremnicu"

    Dim rez As Object
    Set rez = CreateObject("Scripting.Dictionary")
    Set GrupeZaAutoOtpremnicu = rez

    Dim slobodni As Object
    Set slobodni = modDokumenta.NevezaniOtkupi()
    If slobodni.count = 0 Then Exit Function

    Dim otk As Variant
    otk = GetTableData(TBL_OTKUP)
    If Not IsArray(otk) Then Exit Function

    Dim cId As Long, cDat As Long, cSta As Long, cKul As Long, cAmb As Long, cIzd As Long
    cId = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, SRC)
    cDat = RequireColumnIndex(TBL_OTKUP, COL_OTK_DATUM, SRC)
    cSta = RequireColumnIndex(TBL_OTKUP, COL_OTK_STANICA, SRC)
    cKul = RequireColumnIndex(TBL_OTKUP, COL_OTK_KULTURA, SRC)
    cAmb = RequireColumnIndex(TBL_OTKUP, COL_OTK_TIP_AMB, SRC)
    cIzd = RequireColumnIndex(TBL_OTKUP, COL_TRACE_IZDATO_STATUS, SRC)

    Dim filterKljuc As String
    filterKljuc = ""

    Dim i As Long, oid As String, kljuc As String
    For i = 1 To UBound(otk, 1)
        oid = Trim$(NzToText(otk(i, cId)))
        If Len(oid) > 0 Then
            If slobodni.Exists(UCase$(oid)) Then
                If modDokumenta.IzdatoStatusJeIzdato(otk(i, cIzd)) Then
                    If IsDate(otk(i, cDat)) Then
                        kljuc = KljucGrupe(otk(i, cSta), otk(i, cDat), otk(i, cKul), _
                                           otk(i, cAmb))
                        If Not rez.Exists(kljuc) Then
                            rez.Add kljuc, NovaGrupa(otk(i, cSta), otk(i, cDat), _
                                                     otk(i, cKul), otk(i, cAmb))
                        End If
                        rez(kljuc)("izvori").Add oid
                        If Len(samoOtkupID) > 0 Then
                            If StrComp(oid, samoOtkupID, vbTextCompare) = 0 Then
                                filterKljuc = kljuc
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i

    If Len(samoOtkupID) = 0 Then Exit Function

    ' Scope se primenjuje TEK NA KRAJU, nad vec sastavljenim grupama -- da izbor
    ' jednog otkupa vrati CELU njegovu grupu, a ne samo njega.
    Dim suzeno As Object
    Set suzeno = CreateObject("Scripting.Dictionary")
    If Len(filterKljuc) > 0 Then suzeno.Add filterKljuc, rez(filterKljuc)
    Set GrupeZaAutoOtpremnicu = suzeno
End Function

' Grupa nosi DOSLOVNE cinjenice prvog reda koji ju je otvorio, ne razlozen
' kljuc.
'
' Prva verzija je zaglavlje gradila IZ KLJUCA -- a kljuc je normalizovan na
' velika slova, jer sluzi POREDJENJU. Otpremnica je tako dobijala "TEST GAJBA"
' umesto "Test Gajba". Pisac to propusti (RequireIstoPolje poredi
' vbTextCompare), pa se razlika ne vidi ni u jednoj kapiji -- izadje tek na
' stampi i u izvestajima ambalaze, kao tip koji nigde drugde ne postoji.
'
' Pravilo: normalizacija sluzi POREDJENJU, nikad UPISU.
Private Function NovaGrupa(ByVal stanica As Variant, ByVal datum As Variant, _
                           ByVal kultura As Variant, ByVal tipAmb As Variant, _
                           Optional ByVal vozac As Variant = "") As Object
    Dim g As Object
    Set g = CreateObject("Scripting.Dictionary")
    g.Add "stanica", Trim$(NzToText(stanica))
    g.Add "datum", CDate(Int(CDate(datum)))
    g.Add "kultura", Trim$(NzToText(kultura))
    g.Add "tipAmb", Trim$(NzToText(tipAmb))
    g.Add "vozac", Trim$(NzToText(vozac))
    g.Add "predaja", ""
    g.Add "manifest", ""
    g.Add "stigli", CreateObject("Scripting.Dictionary")
    g.Add "izvori", New Collection
    g.Add "redovi", New Collection
    Set NovaGrupa = g
End Function

' Kljuc grupe -- SAMO za poredjenje. Datum ide kao ceo broj, ne kao formatiran
' string: dve celije sa istim danom a razlicitim vremenom su ISTI poslovni dan,
' a Format$ nad Variant datumom zavisi od lokala.
Private Function KljucGrupe(ByVal stanica As Variant, ByVal datum As Variant, _
                            ByVal kultura As Variant, ByVal tipAmb As Variant) As String
    KljucGrupe = UCase$(Trim$(NzToText(stanica))) & "|" & _
                 CStr(CLng(Int(CDate(datum)))) & "|" & _
                 UCase$(Trim$(NzToText(kultura))) & "|" & _
                 UCase$(Trim$(NzToText(tipAmb)))
End Function

' Upis JEDNE auto-otpremnice. Vraca OtpremnicaID; "" = nije napravljena, razlog
' je u outGreska (NIKAD prazan uz prazan ID -- tiho preskakanje je ishod koji
' operater ne moze da razlikuje od "nije bilo posla").
'
' "" SE VRACA SAMO ZA POSLOVNO ODBIJANJE. Sistemski pad izlazi kao GRESKA, i iz
' pisca (outSistemska) i iz svega sto se ovde racuna pre njega -- jer bi inace
' prolaz nastavio da pokusava nad masinom koja ne radi (review #385, P2).
Private Function AutoOtpremnicaUpis(ByVal grupa As Object, _
                                    ByRef outGreska As String) As String
    Const SRC As String = "AutoOtpremnicaUpis"

    outGreska = ""
    On Error GoTo EH

    Dim stanicaID As String, datum As Date, kulturaID As String, tipAmb As String
    stanicaID = CStr(grupa("stanica"))
    datum = CDate(grupa("datum"))
    kulturaID = CStr(grupa("kultura"))
    tipAmb = CStr(grupa("tipAmb"))

    ' VOZAC DOLAZI SA DVE STRANE, I TO JE JEDINA RAZLIKA IZMEDJU DVA POZIVAOCA.
    '
    '   PREDAJA (S5-2) -- otkupac je u PWA cekirao listove i predao ih BAS tom
    '     vozacu. Identitet je cinjenica sa terena i ovde se ne pogadja.
    '   MALINA (S5-1)  -- vozaca nema, jer otkupac == stanica == vozac. Tada je
    '     to OGLEDALO stanice, i pravilo za njega ima jedno telo (modMalina),
    '     isto koje hladnjacki lanac koristi.
    '
    ' Postojanje vozaca u tblVozaci ne proverava se ovde: OtpNapraviDraft to radi
    ' kroz RequireTacnoJedan, pa bi kopija pravila bila druga kapija nad istom
    ' cinjenicom.
    Dim vozacID As String
    vozacID = Trim$(CStr(grupa("vozac")))
    If Len(vozacID) = 0 Then
        vozacID = modMalina.VozacOgledaloZaStanicu(stanicaID)
        If Len(vozacID) = 0 Then
            outGreska = "nema vozaca-ogledala za stanicu " & stanicaID
            Exit Function
        End If
    End If

    ' Broj iz niza STANICE, lokalno. GenerateBrojOtpremnice namerno ne pita
    ' Google: batch bi pravio mrezni poziv PO GRUPI, a auto-otpremnica nije
    ' predlog operateru nego upis (isti razlog kao checkRemote:=False u S4-4).
    Dim broj As String
    broj = modBrojevi.GenerateBrojOtpremnice(stanicaID, datum)
    If Len(broj) = 0 Then
        outGreska = "nije moguce generisati broj otpremnice za stanicu " & stanicaID
        Exit Function
    End If

    Dim h As Object
    Set h = CreateObject("Scripting.Dictionary")
    h("Datum") = datum
    h("StanicaID") = stanicaID
    h("VozacID") = vozacID
    h("KulturaID") = kulturaID
    h("TipAmbalaze") = tipAmb
    h("BrojOtpremnice") = broj

    ' Identitet utovara ide NA DOKUMENT, ne ostaje u prolazu. Malina auto-
    ' otpremnica ga nema -- ona ne nastaje iz predaje -- pa ostaje prazan.
    If Len(Trim$(CStr(grupa("predaja")))) > 0 Then
        h("PredajaID") = Trim$(CStr(grupa("predaja")))
    End If

    ' Jedan potez: nastaje i ODMAH se izdaje. Auto-tok nema nezavisno ocekivanje
    ' koje bi cekalo potvrdu -- ocekivanje se izvodi iz izvora (ZBR-KANON-04).
    Dim sistemska As Boolean
    AutoOtpremnicaUpis = modDokumenta.CreateOtpremnicaIzIzvora_TX(h, grupa("izvori"), _
                                                                  outGreska, sistemska)
    If Len(AutoOtpremnicaUpis) = 0 Then
        If Len(outGreska) = 0 Then outGreska = "pisac nije vratio OtpremnicaID"

        ' Pisac je vec rollback-ovao i vratio razlog; ovde se samo odlucuje da
        ' li prolaz sme dalje. Greska nosi ISTI tekst koji bi isao u outGreske,
        ' da izvestaj ne osiromasi zato sto je pad tezi.
        If sistemska Then
            Err.Raise vbObjectError + modSchemaGuard.ERR_SIS_OD + 26, SRC, _
                      "Sistemski pad pisca otpremnice (stanica " & stanicaID & "): " & _
                      outGreska
        End If
    End If
    Exit Function

EH:
    ' Opis PRE LogErr-a -- LogErr usput brise stanje greske (#383, P2).
    Dim errNum As Long, errDesc As String, errSrc As String
    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    LogErr SRC, "stanica=" & stanicaID & " datum=" & CStr(datum)

    ' Sve PRE pisca (ogledalo vozaca, broj, citanje grupe) moze da padne i
    ' sistemski -- npr. RequireColumnIndex nad tabelom bez kolone. Takav pad se
    ' ne pretvara u "ova grupa nije prosla".
    If modSchemaGuard.JeSistemskiPad(errNum) Then
        Err.Raise errNum, SRC, "Source=" & errSrc & " | " & errDesc
    End If

    AutoOtpremnicaUpis = ""
    If Len(outGreska) = 0 Then outGreska = errDesc
End Function

' ============================================================
' PREDAJA ROBE VOZACU POSTAJE OTPREMNICA (S5-2).
'
' Poslovni dogadjaj: otkupac u PWA cekira otkupne listove i preda ih vozacu.
' TAJ CIN je osnova otpremnice -- a otpremnica je osnova zbirne. Vozac zato
' nikad ne vidi "slobodne" otkupe: dok mu dodju do ruke, vec su u NJEGOVOJ
' otpremnici.
'
' Zateceni tok je isti dogadjaj upisivao kao PECAT: Otkup.VozacID preko
' TryUpdateVozacID. To je kolona koju ciljni model nema -- vozac pripada
' OTPREMNICI (S4.1c) -- i njen jedini citalac je bio pauziran potrosac.
'
' PREDAJA STIZE KAO N REDOVA, NE KAO JEDAN DOGADJAJ. PWA ponovo posalje svaki
' OTK red sa popunjenim VozacID-em, pa desktop vidi N zasebnih duplikata. Da se
' otpremnica pravi po redu, jedan utovar bi dao deset dokumenata. Zato se predaje
' SKUPLJAJU kroz prolaz lista, pa se grupisu -- isti kljuc kao malina
' auto-otpremnica, plus vozac.
'
' DVA UTOVARA ISTOG DANA SU DVA DOKUMENTA. Otpremnica nastaje IZDATA (jedan
' potez), a izdata se ne dopunjuje (A13). To nije ogranicenje nego istina o
' fizickom dogadjaju: vozac je dolazio dvaput.
' ============================================================

' ISO datum (yyyy-mm-dd, sa opcionim vremenom) -> Date, BEZ oslanjanja na lokal.
'
' CDate nad ISO stringom je lokalno zavisan, i to nije teorija nego MERENJE:
' u ovom okruzenju CDate("2091-01-23") vraca 8230-04-15. Tiho, bez greske --
' pa bi otpremnica nosila datum koji nije nicim povezan sa danom utovara.
'
' PWA salje ISO (toISOString / yyyy-mm-dd), pa se datum predaje cita eksplicitno:
' prvih deset znakova, tri broja, DateSerial. Vrednost koja je VEC Date (Excel
' ume da je tako vrati) se prihvata kakva jeste.
'
' Napomena: IsParsableMasterSyncDate stoji na istom CDate-u i koristi ga i uvoz
' otkupa (GS_DATUM). Da li i tamo stize ISO string ili pravi Date -- nije
' mereno; zapisano u planu kao otvorena stavka, ne dira se iz ovog reza.
Private Function IsoUDatum(ByVal v As Variant, ByRef outD As Date) As Boolean
    On Error GoTo EH

    If IsDate(v) And Not VarType(v) = vbString Then
        outD = Int(CDate(v))
        IsoUDatum = (outD >= DateSerial(2000, 1, 1))
        Exit Function
    End If

    Dim s As String
    s = Trim$(CStr(nz(v, "")))
    If Len(s) < 10 Then Exit Function

    If Mid$(s, 5, 1) <> "-" Or Mid$(s, 8, 1) <> "-" Then Exit Function

    Dim g As Long, m As Long, d As Long
    g = CLng(Mid$(s, 1, 4))
    m = CLng(Mid$(s, 6, 2))
    d = CLng(Mid$(s, 9, 2))

    If m < 1 Or m > 12 Or d < 1 Or d > 31 Then Exit Function

    outD = DateSerial(g, m, d)
    IsoUDatum = (outD >= DateSerial(2000, 1, 1))
    Exit Function

EH:
    IsoUDatum = False
End Function

' Otpremnica koja je vec nastala iz OVOG utovara. "" = nijedna.
'
' STROGO 0-ILI-1, ne "vrati prvi aktivni" (review #388, drugi krug). Jedan klik
' otkupca je jedan dokument, pa su DVE aktivne otpremnice pod istim PredajaID-em
' korupcija -- tacno ono stanje koje ova kapija i postoji da spreci. Da vraca
' prvi, sakrila bi sopstveni promasaj: drugi dokument bi ostao nevidljiv, a
' zakasneli blok bi bio odbijen "zbog" prvog.
'
' Trazi se AKTIVNA: stornirana otpremnica znaci da je utovar ponisten, pa
' ponovljena predaja sme da napravi nov dokument. Ispravka NIJE ponistenje --
' nova verzija nosi isti PredajaID (modDokumenta.OtpIspravi), pa je i dalje
' tacno jedna aktivna.
Private Function OtpremnicaPoPredaji(ByVal predajaID As String) As String
    Const SRC As String = "OtpremnicaPoPredaji"

    If Len(Trim$(predajaID)) = 0 Then Exit Function

    Dim d As Variant
    d = GetTableData(TBL_OTPREMNICA)
    If Not IsArray(d) Then Exit Function

    Dim cPred As Long, cID As Long, cSt As Long
    cPred = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_PREDAJA_ID)
    If cPred = 0 Then Exit Function
    cID = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_ID, SRC)
    cSt = RequireColumnIndex(TBL_OTPREMNICA, COL_STORNIRANO, SRC)

    Dim i As Long, nasao As String, koliko As Long
    For i = 1 To UBound(d, 1)
        If StrComp(Trim$(NzToText(d(i, cPred))), Trim$(predajaID), _
                   vbTextCompare) = 0 Then
            If StrComp(Trim$(NzToText(d(i, cSt))), "Da", vbTextCompare) <> 0 Then
                koliko = koliko + 1
                If koliko = 1 Then
                    nasao = Trim$(NzToText(d(i, cID)))
                Else
                    Err.Raise vbObjectError + 8320, SRC, _
                              "Utovar " & predajaID & " ima VISE aktivnih otpremnica (" & _
                              nasao & ", " & Trim$(NzToText(d(i, cID))) & _
                              "). Jedan klik otkupca je jedan dokument."
                End If
            End If
        End If
    Next i

    OtpremnicaPoPredaji = nasao
End Function

' Kolona OTK lista PO IMENU, iz zaglavlja. 0 = nema je.
'
' Po imenu, ne po poziciji: PredajaID i PredatoAt su NOVE kolone, a zatecen
' list ih jos nema. Pozicioni citac bi nad starim listom procitao susednu
' kolonu ili pukao; ovako izostanak ima jasan ishod -- 0, pa predaja staje sa
' imenovanim razlogom umesto da pogadja.
Private Function OtkKolonaPoImenu(ByRef data As Variant, ByVal ime As String) As Long
    If IsEmpty(data) Then Exit Function
    If UBound(data, 1) < 1 Then Exit Function

    Dim c As Long
    For c = LBound(data, 2) To UBound(data, 2)
        If StrComp(Trim$(CStr(nz(data(1, c), ""))), ime, vbTextCompare) = 0 Then
            OtkKolonaPoImenu = c
            Exit Function
        End If
    Next c
End Function

' Jedan red predaje onako kako ga prolaz vidi:
'   Array(redIndex, OtkupID, VozacID, PredajaID, PredatoAt, PredajaClanovi, CRID)
'
' Sve cinjenice o DOGADJAJU citaju se sa REDA -- nijedna se ne izvodi iz robe.
' PredajaClanovi je MANIFEST: CRID-ovi svih blokova koje je otkupac cekirao u
' tom jednom kliku (v. GrupePredaje).
Private Function PredajaKandidat(ByVal redIdx As Long, ByVal otkupID As String, _
                                 ByVal vozacID As String, ByRef data As Variant) As Variant
    Dim cPred As Long, cKad As Long, cClan As Long, cCrid As Long
    cPred = OtkKolonaPoImenu(data, "PredajaID")
    cKad = OtkKolonaPoImenu(data, "PredatoAt")
    cClan = OtkKolonaPoImenu(data, "PredajaClanovi")
    cCrid = OtkKolonaPoImenu(data, "ClientRecordID")

    Dim predajaID As String, predatoAt As String, clanovi As String, crid As String
    If cPred > 0 Then predajaID = Trim$(CStr(nz(data(redIdx, cPred), "")))
    If cKad > 0 Then predatoAt = Trim$(CStr(nz(data(redIdx, cKad), "")))
    If cClan > 0 Then clanovi = Trim$(CStr(nz(data(redIdx, cClan), "")))
    If cCrid > 0 Then crid = Trim$(CStr(nz(data(redIdx, cCrid), "")))

    PredajaKandidat = Array(redIdx, otkupID, vozacID, predajaID, predatoAt, _
                            clanovi, crid)
End Function

' PREDAJE -> GRUPE, PO IDENTITETU DOGADJAJA (review #387, P2).
'
' Kljuc je PredajaID -- jedan klik otkupca u PWA. NIJE (vozac, stanica, dan,
' kultura, ambalaza): ti atributi opisuju ROBU, ne UTOVAR, pa iz njih dogadjaj
' ne moze da se rekonstruise ni u jednom smeru:
'
'   SPAJANJE  -- dve predaje istom vozacu istog dana imaju iste atribute, pa bi
'                zavrsile kao JEDAN dokument iako su bila dva utovara.
'   DELJENJE  -- jedna predaja sme da nosi listove sa VISE DATUMA (sljiva i
'                drugo voce se kupi danima; odluka operatera 23.09.2026), pa bi
'                grupisanje po Otkup.Datum jedan utovar razbilo na vise.
'
' DATUM OTPREMNICE JE DATUM PREDAJE, ne datum otkupnog lista: otpremnica je
' transportni dokument. Zato PredatoAt, a ne COL_OTK_DATUM.
'
' JEDNA PREDAJA JE JEDNA VRSTA VOCA (odluka operatera): vozac jednim dolaskom
' vozi jednu kulturu. Mesana predaja je GRESKA UNOSA, ne slucaj koji se deli --
' odbija se cela, a poruka imenuje sta je naslo.
'
' outVecPredati: red -> OtpremnicaID, za blok koji je VEC predat ISTOM vozacu.
'   Uredan retry, tih no-op.
' outKonflikti: red -> opis, za blok predat DRUGOM vozacu. To nije retry nego
'   protivrecnost: roba je na tudjoj izdatoj otpremnici, pa staje fail-closed.
Private Function GrupePredaje(ByVal predaje As Collection, _
                              ByRef outVecPredati As Object, _
                              ByRef outKonflikti As Object) As Object
    Const SRC As String = "GrupePredaje"

    Dim rez As Object
    Set rez = CreateObject("Scripting.Dictionary")
    Set GrupePredaje = rez

    Set outVecPredati = CreateObject("Scripting.Dictionary")
    Set outKonflikti = CreateObject("Scripting.Dictionary")
    If predaje Is Nothing Then Exit Function
    If predaje.count = 0 Then Exit Function

    Dim clanstvo As Object
    Set clanstvo = modDokumenta.AktivnoClanstvoOtpremnica()

    Dim i As Long, red As Variant
    Dim otkupID As String, vozacID As String, redIdx As Long
    Dim predajaID As String, predatoAt As String
    Dim stanica As String, kultura As String, tipAmb As String
    Dim kljuc As String, g As Object
    Dim danPredaje As Date
    Dim manifest As String, crid As String

    For i = 1 To predaje.count
        red = predaje(i)
        redIdx = CLng(red(0))
        otkupID = Trim$(CStr(red(1)))
        vozacID = Trim$(CStr(red(2)))
        predajaID = Trim$(CStr(red(3)))
        predatoAt = Trim$(CStr(red(4)))
        manifest = Trim$(CStr(red(5)))
        crid = Trim$(CStr(red(6)))

        If clanstvo.Exists(UCase$(otkupID)) Then
            ' VEC PREDAT -- ali kome? (review #387, P2)
            Dim postojecaOtp As String, postojeciVozac As String
            postojecaOtp = CStr(clanstvo(UCase$(otkupID)))
            postojeciVozac = Trim$(NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, _
                                                        postojecaOtp, COL_OTP_VOZAC)))

            If StrComp(postojeciVozac, vozacID, vbTextCompare) = 0 Then
                outVecPredati.Add redIdx, postojecaOtp
            Else
                outKonflikti.Add redIdx, _
                    "blok je vec predat vozacu " & postojeciVozac & _
                    " (otpremnica " & postojecaOtp & "), a red trazi " & vozacID
            End If
        ElseIf Len(predajaID) = 0 Then
            outKonflikti.Add redIdx, _
                "red nema PredajaID -- identitet utovara se NE izvodi iz robe"
        ElseIf Len(OtpremnicaPoPredaji(predajaID)) > 0 Then
            ' UTOVAR JE VEC IZDAT KAO DOKUMENT (review #388, P1).
            '
            ' GAS obradjuje red po red i neuspeo red se vraca u Pending, pa
            ' jedan klik otkupca ume da stigne u DVA ciklusa. Dok identitet
            ' utovara nije imao trajan trag, drugi ciklus je pravio DRUGU izdatu
            ' otpremnicu za JEDAN fizicki utovar -- a izdata se ne dopunjuje
            ' (A13), pa se to posle ne moze ni popraviti bez ispravke.
            '
            ' Zakasneli blok se zato NE dodaje i NE pravi svoj dokument: staje
            ' fail-closed i imenuje otpremnicu, da operater zna gde je ostatak
            ' tog utovara.
            outKonflikti.Add redIdx, _
                "utovar " & predajaID & " je vec izdat kao otpremnica " & _
                OtpremnicaPoPredaji(predajaID) & _
                "; zakasneo blok se ne dodaje u izdat dokument"
        ElseIf Not IsoUDatum(predatoAt, danPredaje) Then
            outKonflikti.Add redIdx, _
                "PredatoAt nije upotrebljiv ISO datum ('" & predatoAt & _
                "'), a otpremnica nosi datum PREDAJE"
        ElseIf Len(manifest) = 0 Then
            outKonflikti.Add redIdx, _
                "red nema PredajaClanovi -- bez manifesta se ne zna kad je utovar CEO"
        ElseIf Len(crid) = 0 Then
            outKonflikti.Add redIdx, _
                "red nema ClientRecordID, pa se ne moze prebrojati u manifestu"
        Else
            stanica = Trim$(NzToText(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, _
                                                 COL_OTK_STANICA)))
            kultura = Trim$(NzToText(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, _
                                                 COL_OTK_KULTURA)))
            tipAmb = Trim$(NzToText(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, _
                                                COL_OTK_TIP_AMB)))

            kljuc = UCase$(predajaID)
            If Not rez.Exists(kljuc) Then
                rez.Add kljuc, NovaGrupa(stanica, danPredaje, kultura, tipAmb, _
                                         vozacID)
                rez(kljuc)("predaja") = predajaID
                rez(kljuc)("manifest") = manifest
            End If

            Set g = rez(kljuc)
            PredajaProveriJednorodnost g, stanica, kultura, tipAmb, vozacID
            PredajaNeslaganje g, "manifest", "spisak clanova utovara", manifest
            g("stigli")(UCase$(crid)) = 1
            g("izvori").Add otkupID
            g("redovi").Add redIdx
        End If
    Next i
End Function

' JEDNA PREDAJA -- JEDNO ZAGLAVLJE. Sto se ne slaze sa prvim redom grupe, tu je
' greska unosa: otkupac je jednim klikom cekirao robu koja ne ide na isti
' dokument. Razlog se PAMTI na grupi, a ne dize odmah -- da poruka moze da
' imenuje SVE sto ne valja, ne samo prvo.
Private Sub PredajaProveriJednorodnost(ByRef g As Object, ByVal stanica As String, _
                                       ByVal kultura As String, ByVal tipAmb As String, _
                                       ByVal vozacID As String)
    PredajaNeslaganje g, "stanica", "otkupno mesto", stanica
    PredajaNeslaganje g, "kultura", "vrsta voca", kultura
    PredajaNeslaganje g, "tipAmb", "tip ambalaze", tipAmb
    PredajaNeslaganje g, "vozac", "vozac", vozacID
End Sub

Private Sub PredajaNeslaganje(ByRef g As Object, ByVal kljuc As String, _
                              ByVal opis As String, ByVal vrednost As String)
    If StrComp(CStr(g(kljuc)), vrednost, vbTextCompare) = 0 Then Exit Sub

    Dim dosad As String
    If g.Exists("greska") Then dosad = CStr(g("greska")) & "; "

    g("greska") = dosad & "jedna predaja nosi razlicit " & opis & " ('" & _
                  CStr(g(kljuc)) & "' i '" & vrednost & "')"
End Sub

' KOJI CLANOVI UTOVARA JOS NISU STIGLI. "" = utovar je CEO.
'
' Manifest je spisak CRID-ova koje je otkupac cekirao u JEDNOM kliku. Poredi se
' sa onim sto je do sada stiglo, pa se zna kad dokument sme da nastane.
Private Function PredajaStaFali(ByVal g As Object) As String
    Dim manifest As String
    manifest = Trim$(CStr(g("manifest")))
    If Len(manifest) = 0 Then Exit Function

    Dim stigli As Object
    Set stigli = g("stigli")

    Dim delovi As Variant, i As Long, fali As String
    delovi = Split(manifest, ",")

    For i = LBound(delovi) To UBound(delovi)
        Dim c As String
        c = Trim$(CStr(delovi(i)))
        If Len(c) > 0 Then
            If Not stigli.Exists(UCase$(c)) Then
                fali = fali & IIf(Len(fali) > 0, ", ", "") & c
            End If
        End If
    Next i

    PredajaStaFali = fali
End Function

' Napravi otpremnice iz skupljenih predaja.
'
' outIshodi: redIdx -> SYNC_STATUS_* za upis nazad u Google list. Red mora da
' dobije ishod SVOJE grupe, a ne zbirni ishod prolaza: pad jedne grupe ne sme
' da oboji redove koji su uredno zavrseni.
'
' Isti ugovor kao malina batch: poslovno odbijanje imenuje razlog i pusta
' ostale grupe, sistemski pad izlazi kao greska (AutoOtpremnicaUpis).
Private Function CreateOtpremniceIzPredaje(ByVal predaje As Collection, _
                                           ByRef outIshodi As Object, _
                                           ByRef outGreske As String) As Long
    Const SRC As String = "CreateOtpremniceIzPredaje"

    ' NAMERNO BEZ On Error: SISTEMSKI pad iz AutoOtpremnicaUpis (sema, kolona,
    ' AppendRow, runtime) mora da izadje do ImportOneOTKSheet, koji ga pretvara
    ' u fatal sync error za CEO list. Lokalni EH bi ga spustio na nivo grupe --
    ' tacno ona granica koju je review #385 zatvorio.

    outGreske = ""
    Set outIshodi = CreateObject("Scripting.Dictionary")

    Dim vecPredati As Object, konflikti As Object, grupe As Object
    Set grupe = GrupePredaje(predaje, vecPredati, konflikti)

    Dim k As Variant
    For Each k In vecPredati.Keys
        outIshodi(CLng(k)) = SYNC_STATUS_DUPLICATE
    Next k

    ' Konflikt NIJE duplikat: roba je na tudjoj izdatoj otpremnici, ili red ne
    ' nosi identitet utovara. Duplicate je terminalan, pa bi red zauvek ostao
    ' neobradjen, a sync zelen -- zato SyncError i imenovan razlog.
    Dim konfGreske As Collection
    Set konfGreske = New Collection
    For Each k In konflikti.Keys
        outIshodi(CLng(k)) = SYNC_STATUS_ERROR & ":predaja -- " & CStr(konflikti(k))
        konfGreske.Add "red " & CStr(k) & ": " & CStr(konflikti(k))
    Next k

    If grupe.count = 0 Then
        outGreske = SpojiRazloge(konfGreske)
        If Len(outGreske) > 0 Then LogWarn SRC, "Predaje bez otpremnice: " & outGreske
        Exit Function
    End If

    ' SEF ZA SMOKE TEST: "predaja koja ne uspe MORA biti fatalna, ne tih
    ' preskok" (AUD-042a). Ime seam-a je ostalo VOZAC_WRITE jer meri ISTU
    ' sposobnost -- promenio se samo upis: umesto pecata na otkupu, otpremnica.
    Dim simPad As Boolean
    simPad = ConsumeFailSeam("VOZAC_WRITE")

    Dim otpID As String, g As String, n As Long, j As Long
    Dim greske As Collection
    Set greske = New Collection

    Dim cekaju As Collection
    Set cekaju = New Collection

    For Each k In grupe.Keys
        ' UTOVAR SE NE FINALIZUJE DOK NIJE CEO (review #388, drugi krug P1).
        '
        ' GAS obradjuje redove pojedinacno, pa deo jednog klika ume da stigne u
        ' ovom ciklusu a ostatak u sledecem. Bez manifesta je master to video kao
        ' zavrsen utovar i IZDAVAO nepotpun dokument -- a izdata otpremnica se ne
        ' dopunjuje (A13), pa je ostatak zauvek ostajao napolju. Prvi krug je
        ' sprecio DRUGI dokument, ali ne i prerano izdavanje PRVOG.
        '
        ' Nepotpun utovar zato ne dobija ni dokument ni status: redovi ostaju
        ' Pending, pa ih sledeci ciklus opet donese. To nije greska nego cekanje,
        ' i tako se i prijavljuje -- u log, ne u outGreske (koji pali fatal flag).
        Dim fali As String
        fali = PredajaStaFali(grupe(k))

        If Len(fali) > 0 Then
            cekaju.Add CStr(k) & ": ceka jos " & fali
            GoTo SledecaGrupa
        End If

        g = ""
        If simPad Then
            otpID = ""
            g = "simuliran pad upisa (fail seam)"
        ElseIf grupe(k).Exists("greska") Then
            ' Mesana predaja se NE DELI na vise dokumenata i ne salje se piscu:
            ' jedan klik je jedan utovar, a jedan utovar je jedna vrsta voca
            ' (odluka operatera 23.09.2026). Ovo je greska unosa, pa poruka mora
            ' da kaze STA ne valja, ne samo da nije proslo.
            otpID = ""
            g = CStr(grupe(k)("greska"))
        Else
            otpID = AutoOtpremnicaUpis(grupe(k), g)
        End If

        If Len(otpID) > 0 Then
            n = n + 1
            For j = 1 To grupe(k)("redovi").count
                outIshodi(CLng(grupe(k)("redovi")(j))) = SYNC_STATUS_MASTER
            Next j
        Else
            Dim razlog As String
            razlog = IIf(Len(g) > 0, g, "nepoznat razlog")
            greske.Add CStr(k) & ": " & razlog
            For j = 1 To grupe(k)("redovi").count
                outIshodi(CLng(grupe(k)("redovi")(j))) = _
                    SYNC_STATUS_ERROR & ":predaja -- " & razlog
            Next j
        End If

SledecaGrupa:
    Next k

    If cekaju.count > 0 Then
        LogInfo SRC, "Utovari koji cekaju ostatak: " & SpojiRazloge(cekaju)
    End If

    ' Konflikti i padovi grupa idu u ISTI izvestaj -- operater ne razlikuje
    ' "nije uspelo zato sto" po tome gde je u kodu presecen.
    Dim sviRazlozi As Collection
    Set sviRazlozi = New Collection
    For Each k In konfGreske
        sviRazlozi.Add k
    Next k
    For Each k In greske
        sviRazlozi.Add k
    Next k

    CreateOtpremniceIzPredaje = n
    outGreske = SpojiRazloge(sviRazlozi)

    If n > 0 Then LogInfo SRC, "Predaja vozacu -> otpremnica, created=" & CStr(n)
    If Len(outGreske) > 0 Then LogWarn SRC, "Predaje bez otpremnice: " & outGreske
End Function

' Razlozi u jedan red, za summary ciklusa i za log.
Private Function SpojiRazloge(ByVal greske As Collection) As String
    If greske Is Nothing Then Exit Function
    If greske.count = 0 Then Exit Function

    Dim i As Long, s As String
    For i = 1 To greske.count
        If i > 1 Then s = s & "; "
        s = s & CStr(greske(i))
    Next i
    SpojiRazloge = s
End Function

' ============================================================
' MALINA MOD -- D: auto-zbirna iz otpremnice (1:1).
'
' Za svaku SLOBODNU IZDATU otpremnicu pravi zbirnu preko kanonskog pisca
' (modDokumenta.CreateZbirnaIzIzvora_TX):
'   - clanstvo je ZAPIS u tblZbirnaIzvori, ne labela na detetu
'   - zbirna dobija SVOJ broj iz niza vozaca; NE nasledjuje otpremnicin
'   - kupac := MALINA_DEFAULT_KUPAC (Hladnjaca)
'   - vrsta, sorta i tip ambalaze se PREUZIMAJU od izvora (ZBR-KANON-04), pa se
'     ne salju piscu
' Idempotentno: slobodna = nije clan nijedne aktivne zbirne (NevezaneOtpremnice).
' Self-gated: u visnji ne radi nista.
'
' Zateceni opis je do S4-4 govorio o SaveZbirnaMulti_TX, o BrojZbirne :=
' BrojOtpremnice i o backfill-u BrojZbirne na otpremnicu i tblOtkup. Nijedno od
' toga vise ne postoji -- komentar je opisivao obrisan model (review #385, P3).
' ============================================================
' BATCH PROLAZ: auto-zbirna za SVE slobodne izdate otpremnice (S4-4).
'
' Drugi od dva pozivaoca istog jezgra. Postoji zato sto otpremnice ne stizu samo
' kroz desktop izdavanje: PWA sync ih donese gotove, i one kuku na izdavanju
' nikad ne prodju. Bez ovog prolaza bi malina operater za njih ostao bez zbirne
' -- tiho, sto je najgori oblik.
'
' samoOtpID suzava prolaz na JEDAN dokument, po IDENTITETU. Stari parametar je
' bio BrojOtpremnice -- labela, koja ume da pripadne dvama dokumentima (A2).
'
' Pad JEDNE zbirne obara ceo prolaz i podize gresku: polovicno odradjen batch
' koji vrati "napravljeno 3" ne kaze koje tri, a operater nema sta da ponovi.
Public Function AutoCreateZbirnaFromOtpremnice_TX(Optional ByVal samoOtpID As String = "") As Long
    Const SRC As String = "AutoCreateZbirnaFromOtpremnice_TX"

    On Error GoTo EH

    If Not AutoZbirnaDostupna() Then Exit Function

    Dim slobodne As Object
    Set slobodne = modDokumenta.NevezaneOtpremnice()
    If slobodne.count = 0 Then Exit Function

    Dim filter As String
    filter = UCase$(Trim$(samoOtpID))

    Dim k As Variant, g As String, zbrID As String, n As Long
    For Each k In slobodne.Keys
        If Len(filter) = 0 Or filter = CStr(k) Then
            zbrID = AutoZbirnaUpis(CStr(k), g)
            If Len(zbrID) = 0 Then
                Err.Raise vbObjectError + 8301, SRC, _
                    "Auto-zbirna nije napravljena za otpremnicu " & CStr(k) & _
                    IIf(Len(g) > 0, ": " & g, "")
            End If
            n = n + 1
        End If
    Next k

    AutoCreateZbirnaFromOtpremnice_TX = n
    If n > 0 Then LogInfo SRC, "Malina auto-zbirna created=" & CStr(n)
    Exit Function
EH:
    ' OPIS SE CITA PRE LogErr-a -- LogErr usput brise stanje greske (review
    ' #383, P2). Isti obrazac koji modAutoHladnjaca vec nosi u komentaru.
    '
    ' Nije kozmetika: orkestrator odlucuje da li je korak pao BAS po Err.Number
    ' posle "On Error Resume Next". Re-raise sa vec obrisanim Err-om mu je
    ' odnosio i broj i razlog -- a sa njima i signal da se nesto desilo.
    Dim errNum As Long, errDesc As String, errSrc As String
    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    LogErr SRC

    Err.Raise errNum, SRC, "Source=" & errSrc & " | " & errDesc
End Function

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

        nextPageToken = ExtractNextPageToken(responseText)
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

Private Function ExtractNextPageToken(ByVal json As String) As String
    ' AUD-018: zajednicko citanje Drive nextPageToken-a za FindOTKSheets i
    ' FindVOZSheets. Prazan rezultat = poslednja strana (petlja staje).
    Dim tokenPos As Long

    tokenPos = InStr(1, json, """nextPageToken""", vbTextCompare)
    If tokenPos = 0 Then Exit Function

    ExtractNextPageToken = ExtractJsonValueAt(json, tokenPos)
End Function

Public Function TestHook_ExtractNextPageToken(ByVal json As String) As String
    ' DEV/SMOKE TEST HOOK ONLY.

    TestHook_ExtractNextPageToken = ExtractNextPageToken(json)
End Function

Public Sub TestHook_ParseFileListVOZ(ByVal json As String, _
                                     ByRef outIDs As Collection, _
                                     ByRef outNames As Collection)
    ' DEV/SMOKE TEST HOOK ONLY.
    ' Dozvoljava mock paginacije (vise strana u istu kolekciju) bez mreze.

    Call ParseFileListVOZ(json, outIDs, outNames)
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

' ByRef: citac po celiji -- ByVal bi kopirao ceo niz po pozivu (v. KOPIJA_NIZA).
Private Function RequireOTKHeaderValue(ByRef data As Variant, _
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
' DEV/SMOKE TEST HOOKS -- RF-28 (AUD-041/042/043)
'
' Drze produkcione rutine Private, ali daju lokalnim regresionim testovima
' (modBusinessFlowProTests) pristup bez Google/HTTP zavisnosti.
' ============================================================

' S5-2: PRAVI upis predaje vozacu, bez Google Sheets-a.
'
' Prolaz kroz list (ImportOneOTKSheet) trazi Drive, pa ga meri smoke suite
' (TestHook_ImportOneOTKSheet). Ovaj hook izlaze telo koje radi POSAO --
' grupisanje predaja i upis otpremnica -- da lokalna regresija moze da meri
' poslovno pravilo bez mreze. Isti obrazac koji modul vec koristi.
'
' predaje: Collection od Array(redIndex, OtkupID, VozacID).
Public Function TestHook_CreateOtpremniceIzPredaje(ByVal predaje As Collection, _
                                                   ByRef outIshodi As Object, _
                                                   ByRef outGreske As String) As Long
    TestHook_CreateOtpremniceIzPredaje = _
        CreateOtpremniceIzPredaje(predaje, outIshodi, outGreske)
End Function

' AUD-042(b): verdikt ValidatePWAOtkup za zadatu Datum vrednost. Ostala polja se
' popunjavaju tako da PROLAZE validaciju, pa je datum jedina promenljiva.
Public Function TestHook_ValidatePWAOtkupDatum(ByVal kooperantID As String, _
                                              ByVal datumValue As Variant) As String
    Dim data As Variant
    ReDim data(1 To 1, 1 To GS_BROJ_DOKUMENTA)

    data(1, GS_KOOPERANT_ID) = kooperantID
    data(1, GS_VRSTA) = "Malina"
    data(1, GS_KOLICINA) = 100
    data(1, GS_CENA) = 200
    data(1, GS_TIP_AMB) = ""
    data(1, GS_KOL_AMB) = 0
    data(1, GS_DATUM) = datumValue

    TestHook_ValidatePWAOtkupDatum = ValidatePWAOtkup(data, 1)
End Function

' AUD-042(b): isto za VOZ putanju (ValidatePWAZbirna).
Public Function TestHook_ValidatePWAZbirnaDatum(ByVal vozacID As String, _
                                               ByVal kupacID As String, _
                                               ByVal datumValue As Variant) As String
    Dim data As Variant
    ReDim data(1 To 1, 1 To VS_BROJ_ZBIRNE)

    data(1, VS_VOZAC_ID) = vozacID
    data(1, VS_KUPAC_ID) = kupacID
    data(1, VS_KOLICINA_KL_I) = 100
    data(1, VS_KOLICINA_KL_II) = 0
    data(1, VS_KOL_AMB) = 0
    data(1, VS_TIP_AMB) = ""
    data(1, VS_DATUM) = datumValue

    ' Od S5-3 validacija trazi IZVORE, jer zbirna bez njih nije dokument.
    ' Ovaj hook meri SAMO datum, pa ostala polja moraju da PROLAZE -- inace bi
    ' tvrdnja o datumu bila zelena iz pogresnog razloga.
    data(1, VS_OTKUP_RECORD_IDS) = "CRID-HOOK-DATUM"

    TestHook_ValidatePWAZbirnaDatum = ValidatePWAZbirna(data, 1)
End Function

' ZBR-IDENT-01 / A21: PRAVI uvoz jednog VOZ reda u tblZbirna, bez Google Sheets-a.
' Fabrikuje se samo ono sto ImportRowToTblZbirna cita; ostale VOZ kolone ne uticu
' na identitet pa se ne popunjavaju.
'
' NE ide kroz ImportOneVOZSheet, dakle ni kroz IsDuplicateZbirnaInMaster -- test
' sam salje razlicite ClientRecordID-ove, bas kao dva uredjaja sa terena.
Public Function TestHook_ImportZbirnaRowPWA(ByVal crid As String, _
                                           ByVal vozacID As String, _
                                           ByVal kupacID As String, _
                                           ByVal datumValue As Variant, _
                                           ByVal vrsta As String, _
                                           ByVal sorta As String, _
                                           ByVal kolKlI As Double, _
                                           ByVal brojZbirne As String, _
                                           Optional ByVal otkupCrids As String = "") As String
    Dim data As Variant
    ReDim data(1 To 1, 1 To VS_BROJ_ZBIRNE)

    ' Od S5-3 zbirna se sastavlja od OTPREMNICA, pa uvoz bez izvora nema sta da
    ' napravi. Testovi koji mere ODBIJANJE (broj tudjeg vlasnika, los format)
    ' izvore i ne salju -- ta kapija puca pre njih.
    data(1, VS_OTKUP_RECORD_IDS) = otkupCrids

    data(1, VS_CLIENT_RECORD_ID) = crid
    data(1, VS_VOZAC_ID) = vozacID
    data(1, VS_KUPAC_ID) = kupacID
    data(1, VS_DATUM) = datumValue
    data(1, VS_VRSTA) = vrsta
    data(1, VS_SORTA) = sorta
    data(1, VS_KOLICINA_KL_I) = kolKlI
    data(1, VS_KOLICINA_KL_II) = 0
    data(1, VS_TIP_AMB) = ""
    data(1, VS_KOL_AMB) = 0
    data(1, VS_BROJ_ZBIRNE) = brojZbirne

    TestHook_ImportZbirnaRowPWA = ImportRowToTblZbirna(data, 1, crid)
End Function

' S5-3 / review #388: razlika izmedju vec uvezene zbirne i reda koji opet stize.
'
' Izlaze telo koje ODLUCUJE (Duplicate vs konflikt), da lokalna regresija moze
' da meri poslovno pravilo bez Google-a. Prolaz kroz list (ImportOneVOZSheet)
' trazi Drive, pa ga meri smoke suite.
Public Function TestHook_PwaZbirnaRazlika(ByVal zbirnaID As String, _
                                          ByVal crid As String, _
                                          ByVal vozacID As String, _
                                          ByVal kupacID As String, _
                                          ByVal datumValue As Variant, _
                                          ByVal brojZbirne As String, _
                                          ByVal otkupCrids As String) As String
    Dim data As Variant
    ReDim data(1 To 1, 1 To VS_BROJ_ZBIRNE)

    data(1, VS_CLIENT_RECORD_ID) = crid
    data(1, VS_VOZAC_ID) = vozacID
    data(1, VS_KUPAC_ID) = kupacID
    data(1, VS_DATUM) = datumValue
    data(1, VS_BROJ_ZBIRNE) = brojZbirne
    data(1, VS_OTKUP_RECORD_IDS) = otkupCrids

    TestHook_PwaZbirnaRazlika = PwaZbirnaRazlika(zbirnaID, data, 1)
End Function

' AUD-041(b): kanonski ZBR fallback generator (MAX sekvence, ne row-count).
Public Function TestHook_GenerateBrojZbirne(ByVal vozacID As String, _
                                           ByVal datum As Date) As String
    TestHook_GenerateBrojZbirne = GenerateBrojZbirne(vozacID, datum)
End Function
' Armira jednokratni fail seam. Kodovi:
'   "VOZAC_WRITE" -- sledeci UpdateCell VozacID-a se ponasa kao neuspeo (AUD-042a)
'   "OTK_HEADER"  -- sledeci header WriteSheetData se ponasa kao neuspeo (AUD-042c)
' Prazan string = razoruzaj (smoke test to radi u svom EH-u).
Public Sub TestHook_ArmFailSeam(ByVal seamKod As String)
    mTestFailSeam = UCase$(Trim$(seamKod))
End Sub

' AUD-042(c): kreiranje OTK sheeta + header, sa trash-om na pad header-a.
Public Function TestHook_CreateOTKSheetWithHeader(ByVal sheetName As String, _
                                                 ByVal folderID As String) As String
    TestHook_CreateOTKSheetWithHeader = _
        CreateOTKSheetWithHeader(sheetName, folderID, BuildOTKOperationalHeaders_(), _
                                 "TestHook_CreateOTKSheetWithHeader")
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

Public Function TestHook_ConsumePWAFatalSyncError() As Boolean
    ' DEV/SMOKE TEST HOOK ONLY.
    ' Liest das Fatal-Sync-Flag und setzt es zurueck, damit Tests
    ' "read/parse failure ist fatal" vs "leeres Sheet ist ok"
    ' unterscheiden koennen.

    TestHook_ConsumePWAFatalSyncError = mLastPWAFatalSyncError
    mLastPWAFatalSyncError = False
End Function

Private Sub ImportOneOTKSheet(ByVal spreadsheetID As String, _
                              ByVal sheetName As String, _
                              ByRef outImported As Long, _
                              ByRef outSkipped As Long, _
                              ByRef outErrors As Long)
    Dim data As Variant
    Dim i As Long
    Dim syncStatus As String
    Dim statusUpdates As Collection

    ' Predaje se SKUPLJAJU pa grupisu posle prolaza (S5-2): jedan utovar
    ' stize kao N redova, pa bi otpremnica po redu dala deset dokumenata
    ' za jedan poslovni dogadjaj.
    Dim predaje As Collection
    Set predaje = New Collection
    
    On Error GoTo EH
    
    ' Daten lesen (erster Tab)
    ' AUD-001: Lese-/Parse-Fehler MUSS fatal sein. Bei defektem oder
    ' abgeschnittenem JSON darf kein einziger Row-Import und kein
    ' WriteBackSyncStatus laufen -- sonst wird die Google-Zeile als
    ' Synced>Master quittiert und nie wieder geliefert.
    If Not TryReadSheetData(spreadsheetID, "Sheet1", data) Then
        outErrors = outErrors + 1
        MarkPWAFatalSyncError "ImportOneOTKSheet", _
            "Sheet read/parse failed (HTTP or malformed JSON). Import aborted before any row import or writeback. Sheet=" & sheetName
        Exit Sub
    End If

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

            ' Duplikat-Check im Master.
            '
            ' ISTI CRID SA DRUGACIJIM SADRZAJEM NIJE DUPLIKAT NEGO KONFLIKT.
            ' Zatecen kod je svaki poznat CRID prosto preskakao, pa je izmenjen
            ' sadrzaj pod istim CRID-om tiho nestajao: PWA misli da je poslala
            ' ispravku, master je nema i niko ne sazna.
            '
            ' Ispravka ide kroz storno i nov dokument (A13), ne kroz ponovni uvoz
            ' istog CRID-a -- pa se ovde staje glasno.
            '
            ' NEIZMERENO, i to je namerno imenovano: ova grana se dostize samo kroz
            ' ImportOneOTKSheet, koji trazi ceo Google Sheet. Sabotaza koja je gasi
            ' NE grize -- testovi zovu ImportRowToTblOtkup_RowTX direktno, gde isti
            ' konflikt hvata kapija u samom ingestu (Test_PWA_IstiCridDrugiSadrzajPada).
            ' Ingest kapija je ta koja garantuje ispravnost; ova daje operateru
            ' status umesto tihog preskoka, i bez nje bi izmenjen sadrzaj opet
            ' nestajao -- zato ostaje.
            If IsDuplicateInMaster(clientRecordID) Then
                Dim posID As String
                posID = modOtkup.OtkupPoClientRecordID(clientRecordID)
                If Len(posID) > 0 Then
                    If Not PwaIstiSadrzaj(posID, data, i) Then
                        statusUpdates.Add Array(i, SYNC_STATUS_ERROR & _
                            ":CRID konflikt -- isti ClientRecordID, drugi sadrzaj (" & _
                            posID & ")")
                        outErrors = outErrors + 1
                        LogError "ImportOneOTKSheet", _
                                 "CRID konflikt: " & clientRecordID & " -> " & posID
                        GoTo NextImportRow
                    End If
                End If
                ' Proveri da li je VozacID update (Otprema tab)
                Dim sheetVozac As String
                sheetVozac = Trim$(CStr(nz(data(i, GS_VOZAC_ID), "")))
                If Len(sheetVozac) > 0 Then
                    ' PREDAJA ROBE VOZACU JE POSLOVNI DOGADJAJ, NE PECAT (S5-2).
                    '
                    ' Otkupac je u PWA cekirao otkupne listove i predao ih BAS tom
                    ' vozacu. Taj cin je osnova OTPREMNICE -- a otpremnica je
                    ' osnova zbirne. Zateceni tok ga je upisivao kao Otkup.VozacID
                    ' (TryUpdateVozacID), u kolonu koju ciljni model nema: vozac
                    ' pripada OTPREMNICI (S4.1c), a jedini citaoci te kolone su
                    ' bili pauzirani.
                    '
                    ' Red se ovde samo ZABELEZI. Ishod mu se ne zna dok se sve
                    ' predaje ovog lista ne grupisu -- jedan utovar stize kao N
                    ' redova -- pa status upisuje CreateOtpremniceIzPredaje, po
                    ' ishodu SVOJE grupe.
                    If Len(posID) > 0 Then
                        predaje.Add PredajaKandidat(i, posID, sheetVozac, data)
                    Else
                        ' AUD-042(a) ostaje: NE SME da prodje kao Duplicate.
                        ' Duplicate je terminalan (import uzima samo Pending), pa
                        ' bi red zauvek ostao bez otpremnice, a sync bio zelen.
                        statusUpdates.Add Array(i, _
                            SYNC_STATUS_ERROR & ":predaja -- otkup nije u masteru", "")
                        outErrors = outErrors + 1

                        MarkPWAFatalSyncError "ImportOneOTKSheet", _
                            "Predaja vozacu: otkup nije nadjen u masteru. Sheet=" & _
                            sheetName & "; Row=" & CStr(i) & _
                            "; ClientRecordID=" & clientRecordID
                    End If
                Else
                    statusUpdates.Add Array(i, SYNC_STATUS_DUPLICATE)
                    outSkipped = outSkipped + 1
                End If
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
                        ' PREDAJA STIZE I NA PRVOM VIDJENJU REDA (review #387, P1).
                        '
                        ' Otkupac sme da preda blok koji jos NIJE sinhronizovan:
                        ' ekran OTPREME spaja lokalne i serverske redove i filtrira
                        ' samo po "nema vozaca". Takav red prvi put stize u master
                        ' VEC SA VOZACEM, i ide OVOM granom -- ne duplikat granom.
                        '
                        ' Dok se predaja ovde nije gledala, ishod je bio najgori
                        ' moguci: otkup nastane, red dobije "Synced>Master" (sto je
                        ' TERMINALNO, import uzima samo Pending), a otpremnice nema
                        ' i nikad je nece biti. Poslovni dogadjaj se gubi u tisini.
                        '
                        ' Zato red sa vozacem NE dobija status ovde: dobija ga
                        ' CreateOtpremniceIzPredaje, po ishodu svoje predaje.
                        outImported = outImported + 1

                        Dim novVozac As String
                        novVozac = Trim$(CStr(nz(data(i, GS_VOZAC_ID), "")))

                        If Len(novVozac) > 0 Then
                            predaje.Add PredajaKandidat(i, newOtkupID, novVozac, data)
                        Else
                            statusUpdates.Add Array(i, SYNC_STATUS_MASTER, newOtkupID)
                        End If
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

    ' PREDAJE -> OTPREMNICE, posle prolaza (S5-2).
    '
    ' Ide PRE WriteBackSyncStatus, jer red sme da dobije "Master" tek kad je
    ' njegova otpremnica stvarno upisana. Obrnut redosled bi Google listu
    ' potvrdio posao koji jos nije uradjen -- a Duplicate je terminalan, pa se
    ' red nikad vise ne bi ponudio.
    If predaje.count > 0 Then
        Dim ishodi As Object, predajaGreske As String, stvorene As Long
        stvorene = CreateOtpremniceIzPredaje(predaje, ishodi, predajaGreske)

        Dim kljucIshoda As Variant
        For Each kljucIshoda In ishodi.Keys
            statusUpdates.Add Array(CLng(kljucIshoda), CStr(ishodi(kljucIshoda)))
            If InStr(1, CStr(ishodi(kljucIshoda)), SYNC_STATUS_ERROR, _
                     vbTextCompare) = 1 Then
                outErrors = outErrors + 1
            Else
                outSkipped = outSkipped + 1
            End If
        Next kljucIshoda

        LogInfo "ImportOneOTKSheet", sheetName & ": predaja -> " & _
                CStr(stvorene) & " otpremnica" & _
                IIf(Len(predajaGreske) > 0, "; bez otpremnice: " & predajaGreske, "")

        ' AUD-042(a) VAZI I DALJE, samo nad drugim upisom: predaja koja nije
        ' postala otpremnica NE SME da prodje kao tih preskok. Red je vec dobio
        ' SyncError iznad; ovde se pali fatal flag, pa top-level zavrsava sa
        ' Monitor_MasterSyncFail umesto zelenim ciklusom.
        If Len(predajaGreske) > 0 Then
            MarkPWAFatalSyncError "ImportOneOTKSheet", _
                "Predaja vozacu nije postala otpremnica. Sheet=" & sheetName & _
                "; Razlozi=" & predajaGreske
        End If
    End If

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

    ' AUD-042(b): datum se MORA validirati pre importa. Import putanja je do sada
    ' na neparsiran datum tiho stavljala Date() (danas) -- dokument je dobijao
    ' pogresan poslovni dan I pogresan ddmmyy u broju, a red je izgledao uspesno
    ' uvezen. Sada red ide u SyncError i operater ga vidi.
    If Not IsParsableMasterSyncDate(data(row, GS_DATUM)) Then
        ValidatePWAOtkup = "Datum invalid: " & Trim$(CStr(nz(data(row, GS_DATUM), "(prazno)")))
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

' Da li vec uvezen dokument nosi ISTI poslovni sadrzaj kao red koji je stigao.
'
' Poredi se ono sto dokument JESTE, ne kako je zapisan.
'
' UCESTVUJU:  KooperantID, KulturaID, Datum, ParcelaID, TipAmbalaze,
'             i jedina stavka -- Klasa, Kolicina, Cena, KolAmbalaze.
'
' NE UCESTVUJU, i za svako postoji razlog:
'   OtkupID, CreatedAt, redosled   ocekuje se da se razlikuju
'   BrojDokumenta                  USLOVNO: ucestvuje samo kad ga PWA izricito
'                                  posalje. Prazan incoming broj znaci da ga je
'                                  master generisao lokalno, pa bi poredjenje
'                                  prijavljivalo konflikt tamo gde ga nema.
'   VrstaVoca / SortaVoca          ulaze posredno: iz njih se razresava KulturaID,
'                                  pa se razlika vidi kroz FK
'
' Funkcija sama parsira red, da bi je mogla zvati OBA mesta: i ingest, i grana
' koja duplikat preskace. Dva poredjenja istog pojma bi se razisla.
Private Function PwaIstiSadrzaj(ByVal otkupID As String, ByVal data As Variant, _
                                ByVal row As Long) As Boolean
    Const SRC As String = "PwaIstiSadrzaj"

    On Error GoTo EH

    Dim kooperantID As String, vrstaVoca As String, sortaVoca As String, klasa As String
    Dim parcelaID As String, tipAmb As String
    kooperantID = Trim$(CStr(nz(data(row, GS_KOOPERANT_ID), "")))
    vrstaVoca = Trim$(CStr(nz(data(row, GS_VRSTA), "")))
    sortaVoca = Trim$(CStr(nz(data(row, GS_SORTA), "")))
    klasa = Trim$(CStr(nz(data(row, GS_KLASA), "")))
    If Len(klasa) = 0 Then klasa = "I"
    parcelaID = Trim$(CStr(nz(data(row, GS_PARCELA_ID), "")))
    tipAmb = Trim$(CStr(nz(data(row, GS_TIP_AMB), "")))

    If Not IsParsableMasterSyncDate(data(row, GS_DATUM)) Then Exit Function

    Dim datum As Date
    datum = CDate(data(row, GS_DATUM))

    Dim kultGreska As String, kulturaID As String
    kulturaID = modOtkup.RazresiKulturuIzVrsteSorte(vrstaVoca, sortaVoca, kultGreska)
    If Len(kulturaID) = 0 Then Exit Function      ' nerazresivo -> ne moze biti isto

    If Not PwaPoljeJednako(otkupID, COL_OTK_KOOPERANT, kooperantID) Then Exit Function
    If Not PwaPoljeJednako(otkupID, COL_OTK_KULTURA, kulturaID) Then Exit Function

    ' PARCELA I TIP AMBALAZE SU DEO SADRZAJA, ne ukras.
    '
    ' Bez njih je isti CRID sa parcele P1 i sa parcele P2 prolazio kao "isti
    ' sadrzaj" -- pa bi ispravljena parcela tiho nestala. Parcela je GlobalGAP
    ' cinjenica (od koje njive je roba), a tip ambalaze odlucuje ceo dvojni upis
    ' gajbi. Oba menjaju STA dokument tvrdi, dakle oba su konflikt.
    If Not PwaPoljeJednako(otkupID, COL_OTK_PARCELA, parcelaID) Then Exit Function
    If Not PwaPoljeJednako(otkupID, COL_OTK_TIP_AMB, tipAmb) Then Exit Function

    ' BROJ DOKUMENTA: uslovno, i to je jedina postena varijanta.
    '
    ' Kad ga PWA NE posalje, master ga generise lokalno -- poredjenje bi tada
    ' prijavljivalo konflikt tamo gde ga nema, jer se dva generisana broja i
    ' ocekuje da se razlikuju.
    '
    ' Kad ga PWA IZRICITO posalje, on je deo payload-a: isti CRID koji prvi put
    ' kaze 123 a drugi put 124 tvrdi dve razlicite stvari. Uslov na neprazan
    ' incoming broj resava to bez ijednog novog polja -- ranije je ovde stajalo da
    ' bi trebalo znati KO je broj dodelio; ne treba, dovoljno je da li ga je PWA
    ' poslala.
    Dim brojIzPwa As String
    brojIzPwa = Trim$(CStr(nz(data(row, GS_BROJ_DOKUMENTA), "")))
    If Len(brojIzPwa) > 0 Then
        If Not PwaPoljeJednako(otkupID, COL_OTK_BR_DOK, brojIzPwa) Then Exit Function
    End If

    Dim dat As Variant
    dat = LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_DATUM)
    If Not IsDate(dat) Then Exit Function
    If Int(CDbl(CDate(dat))) <> Int(CDbl(datum)) Then Exit Function

    ' Stavka: PWA salje tacno jednu (S7).
    Dim d As Variant
    d = GetTableData(TBL_OTKUP_STAVKE)
    If Not IsArray(d) Then Exit Function

    Dim cOtk As Long, cKlasa As Long, cKol As Long, cCena As Long, cAmb As Long
    cOtk = RequireColumnIndex(TBL_OTKUP_STAVKE, COL_OKS_OTKUP_ID, SRC)
    cKlasa = RequireColumnIndex(TBL_OTKUP_STAVKE, COL_OKS_KLASA, SRC)
    cKol = RequireColumnIndex(TBL_OTKUP_STAVKE, COL_OKS_KOLICINA, SRC)
    cCena = RequireColumnIndex(TBL_OTKUP_STAVKE, COL_OKS_CENA, SRC)
    cAmb = RequireColumnIndex(TBL_OTKUP_STAVKE, COL_OKS_KOL_AMB, SRC)

    Dim k As Long, nasao As Long
    For k = 1 To UBound(d, 1)
        If StrComp(Trim$(nz(d(k, cOtk), "")), otkupID, vbTextCompare) = 0 Then
            nasao = nasao + 1
            If nasao > 1 Then Exit Function       ' vise stavki -> nije PWA dokument
            If StrComp(Trim$(nz(d(k, cKlasa), "")), klasa, vbTextCompare) <> 0 Then Exit Function
            If Abs(PwaBroj(d(k, cKol)) - CDbl(nz(data(row, GS_KOLICINA), 0))) > 0.0001 Then Exit Function
            If Abs(PwaBroj(d(k, cCena)) - CDbl(nz(data(row, GS_CENA), 0))) > 0.0001 Then Exit Function
            If Abs(PwaBroj(d(k, cAmb)) - CDbl(nz(data(row, GS_KOL_AMB), 0))) > 0.0001 Then Exit Function
        End If
    Next k

    PwaIstiSadrzaj = (nasao = 1)
    Exit Function

EH:
    ' Greska pri poredjenju NIJE "isto" -- fail-closed.
    LogErr SRC, "OtkupID=" & otkupID
    PwaIstiSadrzaj = False
End Function

' Broj iz celije. NumVal postoji, ali je Private u modOtkupBlok i
' modScrDokumenti -- odavde nevidljiv. vba_check to ne hvata: poziv u
' IZRAZNOJ poziciji je poznata rupa pravila NEDEFINISAN, pa je greska izasla
' tek kao break u VBE-u.
Private Function PwaBroj(ByVal v As Variant) As Double
    If IsNumeric(v) Then PwaBroj = CDbl(v)
End Function

Private Function PwaPoljeJednako(ByVal otkupID As String, ByVal kolona As String, _
                                 ByVal ocekivano As String) As Boolean
    PwaPoljeJednako = (StrComp(Trim$(nz(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, _
                                                    kolona), "")), _
                               Trim$(ocekivano), vbTextCompare) = 0)
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
' Public zbog testa: PWA ingest se inace ne moze izmeriti -- jedini put dovde je
' ImportOneOTKSheet, koji trazi ceo Google Sheet. Kanonski ingest je od Otkup
' cutover-a produkcioni put, pa ne sme da ostane bez zelenog pokrica
' (RunMasterSyncSmokeSuite je zatecena crvena i nije u FULL prolazu).
Public Function ImportRowToTblOtkup_RowTX(ByVal data As Variant, _
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
    LogErr "ImportRowToTblOtkup_RowTX", "ClientRecordID=" & clientRecordID
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
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
    
    ' Datum -- AUD-042(b) STRIKT (nema tihog fallbacka na Date()).
    ' ValidatePWAOtkup ovo hvata jos pre importa (red -> SyncError); ovo je druga
    ' linija za direktne/test pozive, da nijedan ulaz ne prodje kao "danas".
    If Not IsParsableMasterSyncDate(data(row, GS_DATUM)) Then
        Err.Raise vbObjectError + 8105, "ImportRowToTblOtkup", _
              "Datum nije validan u PWA redu (nema fallbacka na danasnji datum). ClientRecordID=" & clientRecordID
    End If
    datum = CDate(data(row, GS_DATUM))

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
    
    ' STANICA JE CINJENICA O TOME GDE JE ROBA PREDATA -- zna je uredjaj
    ' (OtkupacID iz OTK sheet-a), ne kooperant.
    '
    ' Kooperant NIJE zakljucan za stanicu: svaki moze da preda na svakoj, a
    ' otkupni list pripada stanici, ne kooperantu. tblKooperanti.StanicaID je
    ' MATICNA stanica i ima tacno jednog potrosaca -- filter padajuce liste pri
    ' unosu (KOOP_FILTER_BY_OM, modOtkupUI.bas:7034). To je pretpostavka gde ce
    ' kooperant verovatno doci, ne cinjenica gde je dosao.
    '
    ' Do 13.09.2026. je ovde bilo obrnuto: stanica se citala IZ KOOPERANTA, a
    ' uredjaj je bio samo rezerva kad kooperant nema maticnu. Kooperant sa
    ' maticnom ST-A koji preda na ST-B dobijao je dokument knjizen na ST-A --
    ' dok mu je broj, koji PWA pravi po uredjaju, tvrdio ST-B.
    '
    ' Posledica nije labela nego POGRESNO OTKUPNO MESTO: saldo OM-a, izvestaji po
    ' OM-u, modNovac.IsplataBlokProblem (poredi stanicu), station scope u banci i
    ' kapija duplikata broja -- svi rade po stanici.
    '
    ' FAIL-CLOSED: bez uredjaja se NE ZNA gde je roba predata, a pogadjanje po
    ' kooperantu je upravo greska koja se ovde zatvara.
    stanicaID = Trim$(otkupacID)
    If Left$(stanicaID, 3) <> "ST-" Then
        Err.Raise vbObjectError + 8107, "ImportRowToTblOtkup", _
            "OtkupacID ne imenuje stanicu (dobijeno: '" & otkupacID & "'). " & _
            "Stanica dokumenta se ne sme pogadjati iz kooperanta -- maticna " & _
            "stanica je filter pri unosu, ne knjizenje. ClientRecordID=" & clientRecordID
    End If
    
    ' KULTURA SE RAZRESAVA EGZAKTNO, PO (Vrsta, Sorta).
    '
    ' Zatecen kod je trazio samo po VrstaVoca -- sorta se ignorisala -- a kad ne
    ' nadje, sklapao je "vrsta-sorta" string koji IZGLEDA kao FK a ne pokazuje ni
    ' na sta. Takav "ID" je ulazio u dokument i prezivljavao zauvek (S4.1f).
    Dim kultGreska As String
    kulturaID = modOtkup.RazresiKulturuIzVrsteSorte(vrstaVoca, sortaVoca, kultGreska)
    If Len(kulturaID) = 0 Then
        Err.Raise vbObjectError + 8106, "ImportRowToTblOtkup", _
                  "Kultura se ne razresava: " & kultGreska & _
                  " ClientRecordID=" & clientRecordID
    End If
    
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

        ' KONTEKST BROJA, OBE OSE TVRDO -- i to je odluka, ne inercija.
        '
        ' Ovde je razmatrano da dan-osa bude meka (LogWarn) zbog ponocne trke u
        ' PWA: otkup-form.js je cital sat DVAPUT, jednom u generateBrojDokumenta
        ' i jednom u buildOtkupRecord, sa mreznim await-om izmedju, pa je zapis
        ' snimljen oko ponoci nosio juceradnji ddmmyy uz danasnji Datum.
        '
        ' Odustalo se iz dva razloga. Prvi: red odavde ide u CreateOtkup_TX
        ' (:2348), gde kapija stoji fail-closed -- meka grana ovde ne bi nista
        ' propustila, samo bi pomerila poruku sa mesta koje zna ClientRecordID
        ' na mesto koje ga ne zna. Dve kapije nad istim brojem ne smeju da
        ' govore razlicito. Drugi: izvor je popravljen u istom PR-u (dan se
        ' sada cita jednom i deli ga broj i zapis), pa nov klijent tu
        ' neuskladjenost ne moze da proizvede.
        '
        ' Preostali rizik je imenovan, ne sakriven: zapis koji je STARA verzija
        ' PWA snimila unutar tog prozora odbija se pri uvozu. Poruka imenuje i
        ' broj i ClientRecordID, pa je red nadoknadiv rucno.
        Dim brojVerdikt As Long
        brojVerdikt = modBrojevi.BrojOdgovaraKontekstu( _
                          modBrojevi.KIND_OTK, stanicaID, datum, brojDokumenta)

        If modBrojevi.BrojKontekstOdbija(brojVerdikt) Then
            Err.Raise vbObjectError + 8109, "ImportRowToTblOtkup", _
                modBrojevi.BrojKontekstOpis(brojVerdikt, modBrojevi.KIND_OTK, _
                                            stanicaID, datum, brojDokumenta) & _
                " ClientRecordID=" & clientRecordID
        End If
    End If
    
    ' IDEMPOTENCIJA PO ClientRecordID.
    '
    ' Isti CRID sme da stigne vise puta -- retry, ponovljen sync, prekinut prolaz --
    ' ali sme da napravi SAMO JEDAN dokument. Zatecen kod je isti CRID prosto
    ' PRESKAKAO (IsDuplicateInMaster), pa je izmenjen sadrzaj pod istim CRID-om
    ' tiho nestajao: PWA misli da je poslala ispravku, master je nema.
    '
    '   isti CRID + isti sadrzaj    -> NO-OP, vrati postojeci OtkupID
    '   isti CRID + drugi sadrzaj   -> TVRDA GRESKA
    Dim postojeci As String
    postojeci = modOtkup.OtkupPoClientRecordID(clientRecordID)

    If Len(postojeci) > 0 Then
        If PwaIstiSadrzaj(postojeci, data, row) Then
            LogInfo "ImportRowToTblOtkup", "NO-OP (isti CRID i sadrzaj): " & _
                    postojeci & " <- PWA:" & clientRecordID
            ImportRowToTblOtkup = postojeci
            Exit Function
        End If

        Err.Raise vbObjectError + 8107, "ImportRowToTblOtkup", _
                  "ClientRecordID " & clientRecordID & " je vec uvezen kao " & _
                  postojeci & ", ali sa DRUGACIJIM sadrzajem. Ispravka ide kroz " & _
                  "storno i nov dokument (A13), ne kroz ponovni uvoz istog CRID-a."
    End If

    ' KANONSKI PISAC. Sta vise ne ide u header:
    '   vozacID   vozac pripada otpremnici (S4.1c)
    '   novac     kes ne ulazi kroz otkupni list (S4.1b)
    ' Ambalazu knjizi sam pisac, jednom po dokumentu.
    '
    ' Ranije je ovde stajao goli Array(...) od 24 elementa nad tabelom od 39
    ' kolona -- poziciono, pa bi kolona ubacena u sredinu tiho poslala vrednosti
    ' u pogresna polja (CLAUDE.md S3).
    Dim h As Object
    Set h = CreateObject("Scripting.Dictionary")
    h.Add "Datum", datum
    h.Add "KooperantID", kooperantID
    h.Add "StanicaID", stanicaID
    h.Add "KulturaID", kulturaID
    h.Add "VrstaVoca", vrstaVoca
    h.Add "SortaVoca", sortaVoca
    h.Add "TipAmbalaze", tipAmb
    h.Add "BrojDokumenta", brojDokumenta
    h.Add "ParcelaID", parcelaID
    h.Add "ClientRecordID", clientRecordID
    h.Add "SyncSource", "PWA"

    ' KAD JE RED NASTAO NA TERENU, ne kad je stigao u master.
    '
    ' PWA sema ovo polje TRAZI (RequireOTKHeaderValue nad GS_CREATED_AT), a
    ' CreateOtkup_TX ga prima kao opcion header kljuc -- ali adapter ga do sada
    ' nije prosledjivao, pa se bacao na pola puta. Bez njega je jedini vremenski
    ' trag CreatedAt, koji nosi trenutak SINHRONIZACIJE; posle prekida veze to ume
    ' da bude i nekoliko dana kasnije.
    Dim srcCreated As String
    srcCreated = Trim$(CStr(nz(data(row, GS_CREATED_AT), "")))
    If Len(srcCreated) > 0 Then h.Add "SourceCreatedAt", srcCreated

    ' PWA salje JEDAN zapis = JEDNA klasa = ceo dokument (S7).
    Dim stavka As Object
    Set stavka = CreateObject("Scripting.Dictionary")
    stavka.Add "Klasa", klasa
    stavka.Add "Kolicina", kolicina
    stavka.Add "Cena", cena
    stavka.Add "KolAmbalaze", CDbl(kolAmb)

    Dim stavke As Collection
    Set stavke = New Collection
    stavke.Add stavka

    Dim greska As String
    newID = CreateOtkup_TX(h, stavke, greska)

    If Len(newID) = 0 Then
        Err.Raise vbObjectError + 8108, "ImportRowToTblOtkup", _
                  "CreateOtkup_TX nije vratio OtkupID (CRID=" & clientRecordID & "): " & greska
    End If

    LogInfo "ImportRowToTblOtkup", "Uvezeno: " & newID & " <- PWA:" & clientRecordID & _
            " | " & kooperantID & " | " & vrstaVoca & " " & kolicina & "kg"
    ImportRowToTblOtkup = newID
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    LogErr "ImportRowToTblOtkup", "ClientRecordID: " & clientRecordID
    On Error Resume Next
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

' AUD-042(b): jedini test "da li je ovo poslovni datum" za PWA redove.
' Prazno, tekst, samo-vreme ili 1899-baseline vrednost NISU datum. Bez ovoga je
' CDate greska tiho postajala Date() (danas), pa se los red nije mogao ni naci.
Private Function IsParsableMasterSyncDate(ByVal value As Variant) As Boolean
    Dim d As Date

    On Error GoTo EH

    If Len(Trim$(CStr(nz(value, "")))) = 0 Then Exit Function

    d = CDate(value)

    ' Samo-vreme ("12:30") i prazan baseline daju 1899-12-30 -> nije poslovni dan.
    IsParsableMasterSyncDate = (d >= DateSerial(2000, 1, 1))
    Exit Function

EH:
    IsParsableMasterSyncDate = False
End Function

' Poslovni dan (bez vremena) za poredjenje. False ako vrednost nije datum --
' pozivalac tada NE SME da nastavi kao da su dani jednaki.
Private Function TryMasterSyncDay(ByVal value As Variant, ByRef outDay As Date) As Boolean
    On Error GoTo EH

    outDay = Int(CDate(value))
    TryMasterSyncDay = True
    Exit Function

EH:
    TryMasterSyncDay = False
End Function

' AUD-043(b): BrojZbirne se sme upisati SAMO ako je polje prazno ili vec nosi
' isti broj. Bezuslovni RequireUpdateCell je tiho prepisivao postojecu vezu, pa
' je jedna zbirna mogla "preuzeti" otkupe/otpremnice iz druge (dvostruko
' obracunata roba, a prva zbirna ostaje bez stavki). Konflikt = greska.
' ZBR-CHILD-01: kapija mora da gleda ISTO sto pisac pise.
'
' Ranija verzija se zvala RequireBrojZbirneNotConflicting i gledala je SAMO broj.
' Dok je PoveziDeteNaZbirnu pisao samo broj, upis pod istim brojem je bio
' idempotentan -- prepisivanje iste vrednosti preko sebe. Otkad pisac pise i
' generaciju, isti taj put TIHO PREBACUJE dete sa jednog logickog dokumenta na
' drugi, jer dva dokumenta pod istim brojem su tacno ono sto KR-001 dozvoljava.
' Kapija nije oslabila; upis je ojacao ispod nje. To je ZBR-MUT-01 naopako:
' kapija (broj) uza od aktera (broj + generacija).
'
' Matrica:
'   postojeci broj | postojeca gen | novo (broj/gen) | ishod
'   prazan         | prazna        | X / GEN-A       | ALLOW
'   X              | prazna        | X / GEN-A       | ALLOW  (dovrsava vezu)
'   X              | GEN-A         | X / GEN-A       | ALLOW  (idempotentno)
'   X              | GEN-A         | X / GEN-B       | BLOCK
'   X              | GEN-A         | X / prazna      | BLOCK  (ne brise se znanje)
'   X              | bilo sta      | Y / bilo sta    | BLOCK
'   prazan         | GEN-A         | bilo sta        | BLOCK  (integritet)
'
' Prepisivanje roditelja POSTOJI, ali kroz ispravku/prevez, koji su operaterske
' komande. Ingest zatecene cinjenice nije mesto za promenu vlasnistva dokumenta.
Private Sub RequireZbirnaVezaNotConflicting(ByVal tblName As String, _
                                            ByVal rowIndex As Long, _
                                            ByVal columnName As String, _
                                            ByVal brojZbirne As String, _
                                            ByVal genZbirne As String, _
                                            ByVal contextInfo As String, _
                                            ByVal sourceName As String)
    Dim data As Variant
    data = GetTableData(tblName)

    If IsEmpty(data) Then
        Err.Raise ERR_MASTER_SYNC_GUARD_BASE + 32, sourceName, _
                  "Tabela je prazna: " & tblName
    End If

    Dim colIdx As Long
    colIdx = RequireColumnIndex(tblName, columnName, sourceName)

    Dim colGen As Long
    colGen = RequireColumnIndex(tblName, COL_DETE_ZBIRNA_ROD, sourceName)

    Dim current As String
    current = Trim$(CStr(nz(data(rowIndex, colIdx), "")))

    Dim currentGen As String
    currentGen = Trim$(CStr(nz(data(rowIndex, colGen), "")))

    ' Generacija bez broja: dvoje se menjaju u koraku, pa je ovo pokvaren red.
    ' Fail-closed -- ingest ga ne "popravlja" upisom preko.
    If Len(current) = 0 And Len(currentGen) > 0 Then
        Err.Raise ERR_MASTER_SYNC_GUARD_BASE + 45, sourceName, _
                  "Integritet: red nosi ZbirnaID bez BrojZbirne. Table=" & tblName & _
                  "; " & contextInfo & _
                  "; PostojecaGeneracija=" & currentGen
    End If

    If Len(current) > 0 Then
        If Not BrojJednak(current, brojZbirne) Then
            Err.Raise ERR_MASTER_SYNC_GUARD_BASE + 33, sourceName, _
                      "Konflikt BrojZbirne -- red je vec vezan na drugu zbirnu. Table=" & tblName & _
                      "; " & contextInfo & _
                      "; Postojeci=" & current & _
                      "; Novi=" & Trim$(brojZbirne)
        End If
    End If

    ' Isti broj NIJE isti dokument. Poznata generacija se ne menja ingest-om --
    ' ni na drugu, ni na praznu.
    If Len(currentGen) > 0 Then
        If StrComp(currentGen, Trim$(genZbirne), vbTextCompare) <> 0 Then
            Err.Raise ERR_MASTER_SYNC_GUARD_BASE + 46, sourceName, _
                      "Konflikt ZbirnaID -- red je vec dete DRUGOG dokumenta pod istim " & _
                      "brojem. Table=" & tblName & _
                      "; " & contextInfo & _
                      "; Broj=" & Trim$(brojZbirne) & _
                      "; PostojecaGeneracija=" & currentGen & _
                      "; NovaGeneracija=" & Trim$(genZbirne)
        End If
    End If
End Sub

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

' ZBR-CHILD-01: generaciju PRIMA, ne pogadja.
'
' Pozivalac (LinkZbirnaToOtkupAndOtpremnica) ima konkretan ZbirnaID -- membership
' se i razresava preko PK, bas zato sto broj u multi-device koliziji nije
' jedinstven (AUD-043b). Ponovno pitanje ZbirnaIDZaBroj(brojZbirne) bi
' taj identitet BACILO i vratilo prazno u KR-001 slucaju -- dakle bas tamo gde
' je veza najpotrebnija.
Private Sub LinkOtpremnicaToBrojZbirneStrict(ByVal otpremnicaID As String, _
                                             ByVal brojZbirne As String, _
                                             ByVal genZbirne As String, _
                                             ByVal sourceName As String)
    If Len(Trim$(brojZbirne)) = 0 Then
        Err.Raise ERR_MASTER_SYNC_GUARD_BASE + 50, sourceName, _
                  "BrojZbirne je obavezan za link Otpremnica -> Zbirna. OtpremnicaID=" & otpremnicaID
    End If

    Dim rowOtpremnica As Long
    rowOtpremnica = RequireSingleMasterSyncRow(TBL_OTPREMNICA, COL_OTP_ID, otpremnicaID, sourceName)

    ' AUD-043(b) + ZBR-CHILD-01: isti guard kao na otkupu -- ne prepisuj tudju
    ' vezu u tisini, ni kad je broj isti a dokument drugi.
    RequireZbirnaVezaNotConflicting TBL_OTPREMNICA, rowOtpremnica, COL_OTP_BROJ_ZBIRNE, _
                                    brojZbirne, genZbirne, _
                                    "OtpremnicaID=" & otpremnicaID, sourceName

    PoveziDeteNaZbirnu TBL_OTPREMNICA, rowOtpremnica, COL_OTP_BROJ_ZBIRNE, _
                       brojZbirne, genZbirne, sourceName
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
' Javni ulaz (Alt+F8 / dugme u svesci). Census nad src-vba daje NULA pozivalaca,
' ali Public Sub bez argumenata je tacno oblik koji se vezuje na dugme u .xlsm --
' a to se iz izvora ne moze dokazati. Zato handler, ne pretpostavka.
'
' Bez njega bi operater na pauziranom lancu dobio VBA runtime error umesto
' uredne poruke, jer _Core namerno BACA (v. kapiju tamo).
Public Sub ImportZbirneFromPWA()
    On Error GoTo EH

    Call ImportZbirneFromPWA_Core(True)
    Exit Sub

EH:
    Dim opis As String
    opis = Err.description
    LogError "ImportZbirneFromPWA", opis, Err.Number
    MsgBox opis, vbExclamation, APP_NAME
End Sub

Public Function ImportZbirneFromPWA_Core(ByVal showMessages As Boolean) As Boolean
    ' PAUZA JE SKINUTA (S5-3). Uvoz je presao na kanonskog pisca
    ' (CreateZbirnaIzIzvora_TX), pa vise ne postoji razlog zbog kojeg je stajala:
    ' LinkZbirnaToOtkupAndOtpremnica je obrisan i nista se ne pise nazad na
    ' zaglavlje otkupa.

    Dim folderID As String
    Dim sheetIDs As Collection
    Dim sheetNames As Collection
    Dim i As Long
    Dim totalImported As Long
    Dim totalSkipped As Long
    Dim totalErrors As Long
    Dim filesCount As Long

    On Error GoTo EH
    ' Ugovor o formatu + redosled kolona PRE upisa. Ovaj put je danas TVRDO
    ' PAUZIRAN (izvedeni lanac ceka PR7/PR8), pa kapija nista ne kosta -- ali
    ' se odpauziranjem ne sme otvoriti rupa koju ostatak ugovora zatvara.
    modSchema.SchemaReadyOrFail "ImportZbirneFromPWA_Core", TBL_ZBIRNA


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

    LogErr "ImportZbirneFromPWA_Core"
    On Error Resume Next

    MarkPWAFatalSyncError "ImportZbirneFromPWA_Core", errDesc

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

    LogErr SRC
    On Error Resume Next
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
    Dim nextPageToken As String

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

    nextPageToken = ""

    ' AUD-018: bez paginacije se tiho gubi sve preko prvih 100 VOZ sheetova.
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
            FindVOZSheets = False
            Exit Function
        End If

        Call ParseFileListVOZ(responseText, outIDs, outNames)

        nextPageToken = ExtractNextPageToken(responseText)
    Loop While Len(nextPageToken) > 0

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
    
    ' AUD-001: isti fail-closed model kao ImportOneOTKSheet --
    ' defektan/skracen JSON ne sme da proizvede ni jedan uvezen red
    ' ni jedan writeback.
    If Not TryReadSheetData(spreadsheetID, "Sheet1", data) Then
        outErrors = outErrors + 1
        MarkPWAFatalSyncError "ImportOneVOZSheet", _
            "Sheet read/parse failed (HTTP or malformed JSON). Import aborted before any row import or writeback. Sheet=" & sheetName
        Exit Sub
    End If

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
            
            ' ISTI CRID: NO-OP ILI KONFLIKT, NIKAD TIHI DUPLIKAT (review #388, P2).
            '
            ' Zatecen kod je gledao samo POSTOJI LI isti ClientRecordID, pa je
            ' izmenjen sadrzaj pod istim CRID-om tiho nestajao: Duplicate je
            ' terminalan (import uzima samo Pending), pa master ostaje na staroj
            ' verziji dok PWA misli da je poslala ispravku.
            '
            ' Ista klasa problema je vec zatvorena na OTK ingestu
            ' (PwaIstiSadrzaj); ovde je ostala otvorena.
            Dim zbrPostojeci As String
            zbrPostojeci = ZbirnaPoClientRecordID(clientRecordID)

            If Len(zbrPostojeci) > 0 Then
                Dim zbrRazlika As String
                zbrRazlika = PwaZbirnaRazlika(zbrPostojeci, data, i)

                If Len(zbrRazlika) = 0 Then
                    statusUpdates.Add Array(i, SYNC_STATUS_DUPLICATE, "")
                    outSkipped = outSkipped + 1
                Else
                    statusUpdates.Add Array(i, SYNC_STATUS_ERROR & _
                        ":CRID konflikt -- " & zbrRazlika, "")
                    outErrors = outErrors + 1
                    LogError "ImportOneVOZSheet", _
                             "CRID konflikt: " & clientRecordID & " -> " & _
                             zbrPostojeci & "; " & zbrRazlika
                End If
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

    ' Dokument je zaglavlje + STAVKE + IZVORI: rollback koji vrati samo
    ' zaglavlje ostavlja stavku bez dokumenta. tblOtkup i tblOtpremnica vise
    ' nisu ovde -- uvoz zbirne ih ne dira otkad clanstvo nije labela na detetu.
    tx.AddTableSnapshot TBL_ZBIRNA
    tx.AddTableSnapshot TBL_ZBIRNA_STAVKE
    tx.AddTableSnapshot TBL_ZBIRNA_IZVORI

    outZbirnaID = ImportRowToTblZbirna(data, rowIndex, clientRecordID)

    If Len(Trim$(outZbirnaID)) = 0 Then
        Err.Raise ERR_MASTER_SYNC_GUARD_BASE + 61, SRC, _
                  "ImportRowToTblZbirna nije vratio ZbirnaID. ClientRecordID=" & clientRecordID
    End If

    RequireSingleMasterSyncRow TBL_ZBIRNA, COL_ZBR_ID, outZbirnaID, SRC

    outBrojZbirne = GetBrojZbirneForIDStrict(outZbirnaID, SRC)

    ' CLANSTVO JE VEC UPISANO (S5-3).
    '
    ' Do ovog reza je ovde stajao LinkZbirnaToOtkupAndOtpremnica, koji je pisao
    ' Otkup.BrojZbirne i Otpremnica.BrojZbirne -- labelu na detetu, kao vezu.
    ' Kanonski pisac clanstvo upisuje u tblZbirnaIzvori, pa drugog kanala nema
    ' i ovde nema sta da se doda.

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

    LogErr SRC
    On Error Resume Next

    If Not tx Is Nothing Then tx.RollbackTx


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

    ' AUD-042(b): isti strikt datum kao na OTK putanji. Kod zbirne je tihi
    ' fallback na danas bio jos gori -- ddmmyy ulazi u BrojZbirne, pa je red
    ' dobijao broj iz pogresnog dana.
    If Not IsParsableMasterSyncDate(data(row, VS_DATUM)) Then
        ValidatePWAZbirna = "Datum invalid: " & Trim$(CStr(nz(data(row, VS_DATUM), "(prazno)")))
        Exit Function
    End If

    ' SUMMARY POLJA VISE NE ODLUCUJU DA LI IMPORT SME (review #388, P2).
    '
    ' Do S5-3 je ovde stajalo "bar jedna klasa mora imati Kolicina > 0". Od S5-3
    ' sadrzaj zbirne IZVODI kanonski pisac iz izvornih otpremnica (ZBR-KANON-04),
    ' pa se te kolone i ne citaju. Ostavljene u validaciji, pravile su polustanje
    ' u kom summary NIJE izvor istine ali SME da zabrani kanonski dokument: red
    ' cije se otpremnice razresavaju jednoznacno bio bi odbijen zato sto je PWA u
    ' redundantno polje upisala nulu.
    '
    ' Sta uvozu STVARNO treba: identitet zapisa, vozac, kupac, dan i IZVORI.
    ' Bez izvora zbirna nije dokument -- to je jedina nova tvrdnja.
    If Len(Trim$(CStr(nz(data(row, VS_OTKUP_RECORD_IDS), "")))) = 0 Then
        ValidatePWAZbirna = "OtkupRecordIDs missing -- zbirna bez izvora nije dokument"
        Exit Function
    End If
    
    ValidatePWAZbirna = ""
End Function

' ZbirnaID po ClientRecordID-u. "" = nije uvezena.
'
' Ogledalo modOtkup.OtkupPoClientRecordID. Zamenjuje IsDuplicateZbirnaInMaster,
' koji je vracao samo Boolean -- a "postoji" i "isti je" nisu isto pitanje.
Private Function ZbirnaPoClientRecordID(ByVal crid As String) As String
    Const SRC As String = "ZbirnaPoClientRecordID"

    If Len(Trim$(crid)) = 0 Then Exit Function

    Dim d As Variant
    d = GetTableData(TBL_ZBIRNA)
    If Not IsArray(d) Then Exit Function

    Dim cCrid As Long, cID As Long
    cCrid = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_CLIENT_RECORD_ID, SRC)
    cID = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_ID, SRC)

    Dim i As Long
    For i = 1 To UBound(d, 1)
        If StrComp(Trim$(NzToText(d(i, cCrid))), Trim$(crid), vbBinaryCompare) = 0 Then
            ZbirnaPoClientRecordID = Trim$(NzToText(d(i, cID)))
            Exit Function
        End If
    Next i
End Function

' RAZLIKA IZMEDJU VEC UVEZENE ZBIRNE I REDA KOJI OPET STIZE (review #388, P2).
'
' "" = isti dokument, pa je ponovljen red uredan NO-OP. Neprazan tekst = isti
' ClientRecordID nosi DRUGU tvrdnju, sto nije duplikat nego protivrecnost.
'
' POREDE SE SAMO KANONSKE TVRDNJE -- one od kojih dokument zavisi:
'   vozac, kupac, dan, broj (ako ga PWA salje) i SKUP IZVORA.
'
' Summary polja (kolicine, klasa, vrsta, sorta, ambalaza) se NAMERNO ne porede:
' od S5-3 ih kanonski pisac izvodi iz otpremnica, pa razlika u njima ne znaci
' razlicit dokument -- znaci samo da je PWA drugacije sabrala. Poredjenje po
' njima bi proglasavalo konflikt tamo gde ga nema.
'
' Izvori se porede kao SKUP, ne po redosledu: isti utovar poslat dvaput ume da
' navede otkupe drugim redom.
Private Function PwaZbirnaRazlika(ByVal zbirnaID As String, _
                                  ByRef data As Variant, ByVal row As Long) As String
    Const SRC As String = "PwaZbirnaRazlika"

    On Error GoTo EH

    PwaZbirnaRazlika = PoljeRazlika("VozacID", _
        Trim$(NzToText(LookupValue(TBL_ZBIRNA, COL_ZBR_ID, zbirnaID, COL_ZBR_VOZAC))), _
        Trim$(CStr(nz(data(row, VS_VOZAC_ID), ""))))
    If Len(PwaZbirnaRazlika) > 0 Then Exit Function

    PwaZbirnaRazlika = PoljeRazlika("KupacID", _
        Trim$(NzToText(LookupValue(TBL_ZBIRNA, COL_ZBR_ID, zbirnaID, COL_ZBR_KUPAC))), _
        Trim$(CStr(nz(data(row, VS_KUPAC_ID), ""))))
    If Len(PwaZbirnaRazlika) > 0 Then Exit Function

    ' NEUPOREDIV DATUM JE RAZLIKA, NE PRESKOK (review #388, drugi krug P2).
    '
    ' Ranije je poredjenje stajalo pod "If IsoUDatum(...) And IsDate(...)", pa je
    ' nevalidan datum tiho ispadao iz poredjenja. Ako se sve ostalo poklopi, red
    ' bi dobio Duplicate -- a Duplicate je TERMINALAN, pa bi pokvaren red zauvek
    ' nestao. ValidatePWAZbirna tu ne pomaze: ona se zove tek za NOV red, posle
    ' ove grane.
    Dim danNov As Date, danStari As Variant
    danStari = LookupValue(TBL_ZBIRNA, COL_ZBR_ID, zbirnaID, COL_ZBR_DATUM)

    If Not IsoUDatum(data(row, VS_DATUM), danNov) Then
        PwaZbirnaRazlika = "Datum nije upotrebljiv ISO datum ('" & _
                           Trim$(CStr(nz(data(row, VS_DATUM), "(prazno)"))) & "')"
        Exit Function
    End If

    If Not IsDate(danStari) Then
        PwaZbirnaRazlika = "Datum u masteru nije upotrebljiv za poredjenje"
        Exit Function
    End If

    If Int(CDate(danStari)) <> Int(danNov) Then
        PwaZbirnaRazlika = "Datum: master ima " & _
            Format$(CDate(danStari), "dd.mm.yyyy") & ", red nosi " & _
            Format$(danNov, "dd.mm.yyyy")
        Exit Function
    End If

    ' Broj se poredi SAMO ako ga PWA salje: prazan znaci "generisi lokalno", pa
    ' lokalno generisan broj nije razlika u tvrdnji.
    Dim brojNov As String
    brojNov = Trim$(CStr(nz(data(row, VS_BROJ_ZBIRNE), "")))
    If Len(brojNov) > 0 Then
        PwaZbirnaRazlika = PoljeRazlika("BrojZbirne", _
            Trim$(NzToText(LookupValue(TBL_ZBIRNA, COL_ZBR_ID, zbirnaID, COL_ZBR_BROJ))), _
            brojNov)
        If Len(PwaZbirnaRazlika) > 0 Then Exit Function
    End If

    ' SKUP IZVORA. Razresava se ISTIM putem kao pri uvozu, pa se poredi ono sto
    ' bi dokument stvarno dobio -- ne sirovi CRID spisak.
    Dim noviIzvori As Collection
    Set noviIzvori = OtpremniceIzOtkupRecordIDs( _
                         Trim$(CStr(nz(data(row, VS_OTKUP_RECORD_IDS), ""))), _
                         "PwaZbirnaRazlika")
    If noviIzvori Is Nothing Then
        PwaZbirnaRazlika = "izvori se vise ne razresavaju (v. log)"
        Exit Function
    End If

    PwaZbirnaRazlika = SkupIzvoraRazlika(modDokumenta.IzvoriZbirne(zbirnaID), noviIzvori)
    Exit Function

EH:
    PwaZbirnaRazlika = "poredjenje nije uspelo: " & Err.description
End Function

Private Function PoljeRazlika(ByVal ime As String, ByVal stari As String, _
                              ByVal novi As String) As String
    If StrComp(stari, novi, vbTextCompare) = 0 Then Exit Function
    PoljeRazlika = ime & ": master ima '" & stari & "', red nosi '" & novi & "'"
End Function

' "" = isti skup. Poredi se kao SKUP -- redosled nije tvrdnja.
Private Function SkupIzvoraRazlika(ByVal stari As Collection, _
                                   ByVal novi As Collection) As String
    Dim ss As Object, sn As Object
    Set ss = CreateObject("Scripting.Dictionary")
    ss.CompareMode = vbTextCompare
    Set sn = CreateObject("Scripting.Dictionary")
    sn.CompareMode = vbTextCompare

    Dim i As Long
    If Not stari Is Nothing Then
        For i = 1 To stari.count
            ss(Trim$(NzToText(stari(i)))) = 1
        Next i
    End If
    If Not novi Is Nothing Then
        For i = 1 To novi.count
            sn(Trim$(NzToText(novi(i)))) = 1
        Next i
    End If

    Dim k As Variant, visak As String, manjak As String
    For Each k In sn.Keys
        If Not ss.Exists(k) Then visak = visak & IIf(Len(visak) > 0, ", ", "") & CStr(k)
    Next k
    For Each k In ss.Keys
        If Not sn.Exists(k) Then manjak = manjak & IIf(Len(manjak) > 0, ", ", "") & CStr(k)
    Next k

    If Len(visak) = 0 And Len(manjak) = 0 Then Exit Function

    SkupIzvoraRazlika = "izvori:"
    If Len(visak) > 0 Then SkupIzvoraRazlika = SkupIzvoraRazlika & " red dodaje " & visak
    If Len(manjak) > 0 Then SkupIzvoraRazlika = SkupIzvoraRazlika & " red izostavlja " & manjak
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
' ZBIRNA IZ PWA IDE KROZ KANONSKI PISAC (S5-3).
'
' Zatecen uvoz je radio AppendRow(TBL_ZBIRNA) sa golim Array-em od 16 vrednosti,
' i na ZAGLAVLJE pisao kolicinu, klasu, vrstu, sortu i ambalazu -- stari model.
' Bio je drugi put do istog dokumenta, i drugi oblik istog dokumenta.
'
' CLANSTVO SE RAZRESAVA KROZ KANON, NE POGADJA.
'
' PWA salje otkupRecordIDs -- otkupne listove koje je vozac natovario. Kanonska
' zbirna se sastavlja od OTPREMNICA, a svaki otkup pripada tacno jednoj aktivnoj
' otpremnici (AktivnoClanstvoOtpremnica). Prevod otkup -> otpremnica je zato
' TOTALAN I TACAN, a ne heuristika: razlika u odnosu na predaju, gde dogadjaj
' nije imao nikakav zapis pa mu je trebao sopstveni identitet (S5-2).
'
' Poslovni lanac je predaja -> otpremnica -> zbirna (odluka operatera
' 23.09.2026): vozac nikad ne vidi "slobodne" otkupe, jer su mu dok dodju do
' ruke vec u NJEGOVOJ otpremnici. Otkup bez otpremnice ovde zato NIJE redak
' slucaj nego KVAR, i uvoz staje sa imenom tog otkupa.
'
' Sta se vise NE cita sa VOZ reda:
'   kolicine, klasa    -- stavke zbirne izvodi pisac iz izvora (ZBR-KANON-04)
'   vrsta, sorta, amb  -- iste cinjenice, isti razlog; PWA ih usput spaja
'                         zarezom preko svih dnevnih otkupa, sto nije cinjenica
'                         nijednog dokumenta
'   GeneracijaID       -- osa je ZbirnaID (S4-2a); ovo je bio POSLEDNJI pisac
'                         generacije za zbirnu
Private Function ImportRowToTblZbirna(ByVal data As Variant, _
                                      ByVal row As Long, _
                                      ByVal clientRecordID As String) As String
    Const SRC As String = "ImportRowToTblZbirna"

    On Error GoTo EH

    Dim vozacID As String, kupacID As String, brojZbirne As String
    vozacID = Trim$(CStr(nz(data(row, VS_VOZAC_ID), "")))
    kupacID = Trim$(CStr(nz(data(row, VS_KUPAC_ID), "")))

    Dim datum As Date
    If Not IsoUDatum(data(row, VS_DATUM), datum) Then
        LogError SRC, "Datum nije upotrebljiv. CRID=" & clientRecordID
        Exit Function
    End If

    brojZbirne = Trim$(CStr(nz(data(row, VS_BROJ_ZBIRNE), "")))

    ' Broj je LABELA, ali labela mora da bude ispravna i da ne protivreci
    ' sopstvenom redu. Ista odluka kao na uvozu otkupa: kolizija se prijavljuje,
    ' a broj koji ne pripada svom vozacu i danu se odbija.
    If Len(brojZbirne) = 0 Then
        brojZbirne = GenerateBrojZbirne(vozacID, datum)
        If Len(brojZbirne) = 0 Then
            LogError SRC, "Nije moguce generisati BrojZbirne za VozacID=" & vozacID
            Exit Function
        End If
        LogWarn SRC, "BrojZbirne fallback-generated lokalno za " & clientRecordID
    ElseIf Not IsValidBrojZbirneFormat(brojZbirne) Then
        LogError SRC, "Invalid BrojZbirne format: " & brojZbirne & _
                 " (CRID=" & clientRecordID & ")"
        Exit Function
    Else
        Dim zbrVerdikt As Long
        zbrVerdikt = modBrojevi.BrojOdgovaraKontekstu( _
                         modBrojevi.KIND_ZBR, vozacID, datum, brojZbirne)
        If modBrojevi.BrojKontekstOdbija(zbrVerdikt) Then
            LogError SRC, _
                modBrojevi.BrojKontekstOpis(zbrVerdikt, modBrojevi.KIND_ZBR, _
                                            vozacID, datum, brojZbirne) & _
                " (CRID=" & clientRecordID & ")"
            Exit Function
        End If
    End If

    Dim izvori As Collection
    Set izvori = OtpremniceIzOtkupRecordIDs( _
                     Trim$(CStr(nz(data(row, VS_OTKUP_RECORD_IDS), ""))), _
                     clientRecordID)
    If izvori Is Nothing Then Exit Function

    Dim h As Object
    Set h = CreateObject("Scripting.Dictionary")
    h.Add "Datum", datum
    h.Add "VozacID", vozacID
    h.Add "KupacID", kupacID
    h.Add "BrojZbirne", brojZbirne
    h.Add "Hladnjaca", CStr(nz(LookupValue(TBL_KUPCI, "KupacID", kupacID, _
                                           "Hladnjaca"), ""))
    h.Add "ClientRecordID", clientRecordID
    h.Add "SyncSource", "PWA"

    Dim greska As String
    ' brojSaTerena:=True -- ovo je INGEST vec nastale cinjenice, ne KOMANDA
    ' operatera. Dva uredjaja offline umeju da dodele isti broj (KR-001,
    ' prihvacen rizik), pa bi odbijanje znacilo da vozacev dokument nestane iz
    ' kancelarije. Kolizija se PRIJAVLJUJE (PrijaviKolizijuBrojaZbirne nize), a
    ' identitet je i dalje ZbirnaID -- dve cinjenice se ne stapaju u jednu.
    ImportRowToTblZbirna = modDokumenta.CreateZbirnaIzIzvora_TX(h, izvori, _
                                                               greska, True)

    If Len(ImportRowToTblZbirna) = 0 Then
        LogError SRC, "CreateZbirnaIzIzvora_TX nije vratio ZbirnaID (CRID=" & _
                 clientRecordID & "): " & greska
        Exit Function
    End If

    ' INGEST, PA DETEKCIJA -- ne blokada.
    '
    ' Dva uredjaja offline umeju da dodele isti broj istom vozacu i kupcu
    ' (KR-001). Broj je labela (A2), pa dva dokumenta pod istim brojem NISU
    ' greska podatka -- ali jesu nesto sto operater mora da vidi. Identitet je
    ' ZbirnaID, pa se dve terenske cinjenice i ne mogu stopiti u jednu.
    PrijaviKolizijuBrojaZbirne brojZbirne, vozacID, kupacID, clientRecordID

    LogInfo SRC, "Importiert: " & ImportRowToTblZbirna & " BrojZbirne=" & brojZbirne & _
            " | " & vozacID & " | " & kupacID & " | izvora=" & CStr(izvori.count)
    Exit Function

EH:
    LogErr SRC, "ClientRecordID: " & clientRecordID
    ImportRowToTblZbirna = ""
End Function

' otkupRecordIDs (CRID-ovi, zarezom) -> Collection OtpremnicaID-eva, bez
' ponavljanja, redom prvog pojavljivanja. Nothing = uvoz ovog reda NE SME dalje.
'
' Svaki korak je KANONSKI citac: CRID -> OtkupID (modOtkup.OtkupPoClientRecordID),
' OtkupID -> OtpremnicaID (modDokumenta.OtpremnicaZaOtkup). Nista se ne trazi po
' poslovnom broju i nista se ne pogadja.
'
' Fail-closed je ovde jedini ispravan izbor: zbirna od DELA onoga sto je vozac
' natovario je drugi dokument od onog koji je vozac napravio, a ne "malo manji".
Private Function OtpremniceIzOtkupRecordIDs(ByVal crids As String, _
                                            ByVal clientRecordID As String) As Collection
    Const SRC As String = "OtpremniceIzOtkupRecordIDs"

    If Len(Trim$(crids)) = 0 Then
        LogError SRC, "Zbirna nema nijedan otkupRecordID (CRID=" & clientRecordID & _
                 "). Zbirna bez izvora nije dokument."
        Exit Function
    End If

    Dim rez As New Collection
    Dim vidjene As Object
    Set vidjene = CreateObject("Scripting.Dictionary")
    vidjene.CompareMode = vbTextCompare

    Dim delovi As Variant, i As Long
    delovi = Split(crids, ",")

    For i = LBound(delovi) To UBound(delovi)
        Dim crid As String
        crid = Trim$(CStr(delovi(i)))

        If Len(crid) > 0 Then
            Dim otkupID As String
            otkupID = modOtkup.OtkupPoClientRecordID(crid)

            If Len(otkupID) = 0 Then
                LogError SRC, "Otkup nije u masteru: CRID=" & crid & _
                         " (zbirna CRID=" & clientRecordID & ")"
                Exit Function
            End If

            Dim otpID As String
            otpID = modDokumenta.OtpremnicaZaOtkup(otkupID)

            If Len(otpID) = 0 Then
                LogError SRC, "Otkup " & otkupID & " nije ni u jednoj aktivnoj " & _
                         "otpremnici, a zbirna se sastavlja od otpremnica " & _
                         "(zbirna CRID=" & clientRecordID & ")"
                Exit Function
            End If

            If Not vidjene.Exists(otpID) Then
                vidjene.Add otpID, 1
                rez.Add otpID
            End If
        End If
    Next i

    If rez.count = 0 Then
        LogError SRC, "Nijedan upotrebljiv izvor (CRID=" & clientRecordID & ")"
        Exit Function
    End If

    Set OtpremniceIzOtkupRecordIDs = rez
End Function

' Detekcija kolizije broja POSLE upisa. Zove se sa vec ubacenim redom, pa meri
' STANJE KOJE JE IMPORT OSTAVIO, ne ono pre njega.
'
' NIKAD ne dize gresku: ovo radi unutar transakcije uvoza, pa bi pad detekcije
' oborio i sam upis -- tacno ono sto ova funkcija treba da spreci.
Private Sub PrijaviKolizijuBrojaZbirne(ByVal broj As String, ByVal vozacID As String, _
                                       ByVal kupacID As String, ByVal crid As String)
    Const SRC As String = "ImportRowToTblZbirna"
    Dim id As ZbirnaIdent
    On Error GoTo EH

    id = ZbirnaIdentResolve(broj, vozacID, kupacID)

    If id.integrityStatus <> ZBR_INT_OK Then
        LogWarn SRC, "ZBR-IDENT-01: aktivna zbirna bez GeneracijaID pod brojem " & _
                broj, "CRID=" & crid
        Exit Sub
    End If

    If id.activeLogicalCount > 1 Then
        LogWarn SRC, "ZBR-IDENT-01: broj " & broj & " nosi " & _
                CStr(id.activeLogicalCount) & " aktivna dokumenta (" & _
                CStr(id.activeOwnerCount) & " vlasnika) -- uvoz je prihvacen, " & _
                "vezivanje po broju vise nije jednoznacno", "CRID=" & crid
    End If
    Exit Sub

EH:
    ' Detekcija koja padne ne sme da obori uvoz -- samo se zapise da je pala.
    LogErr SRC & ".PrijaviKolizijuBrojaZbirne", "CRID=" & crid
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

' AUD-041(b): NE brojati redove.
'
' Stari generator je bio ROW-COUNT (seq = 1, pa seq = seq + 1 za svaki red istog
' vozaca i dana). Na svaku rupu u nizu daje duplikat: ako u tblZbirna postoje
' "1/ddmmyy" i "1/ddmmyy-3" (rucni unos, storno, brisanje), count = 2 -> predlog
' je opet "1/ddmmyy-2"... a kad ih je 3 -> "-3", broj koji vec postoji.
'
' Kanonska alokacija zivi u modBrojevi.SuggestNextBroj(KIND_ZBR):
'   MaxSeqFromTable (MAX sekvence, ne count) + ApplyMirrorPrefix ("S" za
'   mirror-stanicu) + BrojZbirneExists bump-loop dok broj ne bude slobodan.
'
' checkRemote:=False -- ovo je fallback UNUTAR importa remote reda (row-TX);
' remote skan istog VOZ sheeta bi ovde bio suvisan HTTP poziv.
'
' Prazan rezultat ostaje fatalan za red: SuggestNextBroj vraca "" na gresku i
' kada je auto-broj toggle (IsAutoBrojDokumenta) iskljucen, a
' ImportRowToTblZbirna prazan broj tretira kao SyncError.
Private Function GenerateBrojZbirne(ByVal vozacID As String, ByVal datum As Date) As String
    ' Zadrzan guard iz starog generatora: vozacID bez cifara ne daje broj
    ' (SuggestNextBroj bi kroz FormatBroj vratio "0/ddmmyy").
    If Len(ExtractNumericVozacBroj(vozacID)) = 0 Then
        GenerateBrojZbirne = ""
        Exit Function
    End If

    GenerateBrojZbirne = SuggestNextBroj(KIND_ZBR, vozacID, datum, checkRemote:=False)
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

' ByRef: citac po celiji -- ByVal bi kopirao ceo niz po pozivu (v. KOPIJA_NIZA).
Private Function RequireVOZHeaderValue(ByRef data As Variant, _
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

' True samo ako je seam armiran BAS za ovaj kod; potrosi ga (jednokratno).
Private Function ConsumeFailSeam(ByVal seamKod As String) As Boolean
    If Len(mTestFailSeam) = 0 Then Exit Function
    If mTestFailSeam <> UCase$(Trim$(seamKod)) Then Exit Function

    mTestFailSeam = ""
    ConsumeFailSeam = True
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
    LogErr SRC
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

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

' ByRef: citac po celiji -- ByVal bi kopirao ceo niz po pozivu (v. KOPIJA_NIZA).
Private Function GeoHasValue(ByRef data As Variant, _
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

