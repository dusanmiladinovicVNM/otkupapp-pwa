Attribute VB_Name = "modStammdatenSync"
Option Explicit

' ============================================================
' modStammdatenSync - Export Stammdaten zu Google Sheet
'
' Schreibt tblKooperanti, tblKulture, tblConfig (Cene)
' in ein Google Sheet "Stammdaten" fuer die PWA.
'
' Config-Keys in tblConfig:
'   GOOGLE_STAMMDATEN_SHEET_ID   (wird automatisch erstellt)
'   GOOGLE_PWA_FOLDER_ID         (Drive Folder fuer PWA-Sheets)
'
' Aufruf: Button in frmMain oder manuell via SyncStammdatenToGoogle
' ============================================================

Private Const KARTICE_TAB_NAME As String = "Kartice"

Private Function StammdatenTabs() As Variant
    StammdatenTabs = Array( _
        "Kooperanti", _
        "Kulture", _
        "Parcele", _
        "Config", _
        "Users", _
        "Fakture", _
        "FakturaStavke", _
        "SaldoOMDetail", _
        "Stanice", _
        "Kupci", _
        "Vozaci", _
        "Artikli", _
        "MagacinKoop" _
    )
End Function

Private Sub EnsureStammdatenTabsBestEffort(ByVal sheetID As String)
    Dim tabs As Variant
    Dim i As Long

    If Len(Trim$(sheetID)) = 0 Then Exit Sub

    tabs = StammdatenTabs()

    For i = LBound(tabs) To UBound(tabs)
        On Error Resume Next
        Call AddSheetTab(sheetID, CStr(tabs(i)))
        If Err.Number <> 0 Then
            LogWarn "EnsureStammdatenTabsBestEffort", _
                    "AddSheetTab failed for tab=" & CStr(tabs(i)) & _
                    "; Error=" & Err.description
            Err.Clear
        End If
        On Error GoTo 0
    Next i
End Sub

Private Function MgmtReportTabs() As Variant
    MgmtReportTabs = Array( _
        "SaldoOM", _
        "SaldoKupci", _
        "OtkupPoOM", _
        "PredatoPoKupcu", _
        "OtkupiAll" _
    )
End Function

Private Sub EnsureKarticeTabsBestEffort(ByVal sheetID As String)
    If Len(Trim$(sheetID)) = 0 Then Exit Sub

    On Error Resume Next
    Call AddSheetTab(sheetID, KARTICE_TAB_NAME)

    If Err.Number <> 0 Then
        LogWarn "EnsureKarticeTabsBestEffort", _
                "AddSheetTab failed for tab=" & KARTICE_TAB_NAME & _
                "; Error=" & Err.description
        Err.Clear
    End If

    On Error GoTo 0
End Sub

Private Function GetReportsFolderID() As String
    Dim folderID As String

    folderID = GetConfigValue("GOOGLE_REPORTS_FOLDER_ID")

    If Len(Trim$(folderID)) = 0 Then
        folderID = GetConfigValue("GOOGLE_PWA_FOLDER_ID")
    End If

    GetReportsFolderID = folderID
End Function

Private Sub EnsureMgmtReportTabsBestEffort(ByVal sheetID As String)
    Dim tabs As Variant
    Dim i As Long

    If Len(Trim$(sheetID)) = 0 Then Exit Sub

    tabs = MgmtReportTabs()

    For i = LBound(tabs) To UBound(tabs)
        On Error Resume Next
        Call AddSheetTab(sheetID, CStr(tabs(i)))
        If Err.Number <> 0 Then
            LogWarn "EnsureMgmtReportTabsBestEffort", _
                    "AddSheetTab failed for tab=" & CStr(tabs(i)) & _
                    "; Error=" & Err.description
            Err.Clear
        End If
        On Error GoTo 0
    Next i
End Sub

 

' ============================================================
' PUBLIC -- Hauptfunktion
' ============================================================
Public Sub SyncStammdatenToGoogle()
    Call SyncStammdatenToGoogle_Core(True)
End Sub

Public Function SyncStammdatenToGoogle_Core(ByVal showMessages As Boolean) As Boolean
    Dim folderID As String
    Dim sheetID As String
    Dim successCount As Long
    
    Const TOTAL_STAMMDATEN_TABS As Long = 13
    
    On Error GoTo EH
    
    SyncStammdatenToGoogle_Core = False
    
    If Not IsGoogleAuthConfigured() Then
        Monitor_StammdatenSyncFail _
            errNum:=0, _
            errDesc:="Google OAuth2 nije konfigurisan.", _
            errSrc:="modStammdatenSync.SyncStammdatenToGoogle_Core", _
            successCount:=0, _
            totalTabs:=TOTAL_STAMMDATEN_TABS
        
        If showMessages Then
            MsgBox "Google OAuth2 nije konfigurisan!" & vbCrLf & _
                   "Pokrenite RunGoogleAuthSetup iz modGoogleAuth.", _
                   vbCritical, APP_NAME
        End If
        
        Exit Function
    End If
    
    folderID = GetConfigValue("GOOGLE_PWA_FOLDER_ID")
    If Len(Trim$(folderID)) = 0 Then
        Monitor_StammdatenSyncFail _
            errNum:=0, _
            errDesc:="GOOGLE_PWA_FOLDER_ID nije postavljen.", _
            errSrc:="modStammdatenSync.SyncStammdatenToGoogle_Core", _
            successCount:=0, _
            totalTabs:=TOTAL_STAMMDATEN_TABS
        
        If showMessages Then
            MsgBox "GOOGLE_PWA_FOLDER_ID nije postavljen u tblConfig!" & vbCrLf & _
                   "Unesite ID Google Drive foldera za PWA.", _
                   vbCritical, APP_NAME
        End If
        
        Exit Function
    End If
    
    sheetID = GetConfigValue("GOOGLE_STAMMDATEN_SHEET_ID")
    
    If Len(Trim$(sheetID)) = 0 Then
        sheetID = GetSpreadsheetID("Stammdaten", folderID)
    End If
    
    If Len(Trim$(sheetID)) = 0 Then
        sheetID = CreateSpreadsheet("Stammdaten", folderID)
        
        If Len(sheetID) = 0 Then
            Monitor_StammdatenSyncFail _
                errNum:=0, _
                errDesc:="Google Stammdaten sheet could not be created.", _
                errSrc:="modStammdatenSync.SyncStammdatenToGoogle_Core", _
                successCount:=0, _
                totalTabs:=TOTAL_STAMMDATEN_TABS
            
            If showMessages Then
                MsgBox "Google Sheet konnte nicht erstellt werden!", vbCritical, APP_NAME
            End If
            
            Exit Function
        End If
    End If
    
    Call SetConfigValue("GOOGLE_STAMMDATEN_SHEET_ID", sheetID)
    Call EnsureStammdatenTabsBestEffort(sheetID)
    
    If ExportKooperanti(sheetID) Then successCount = successCount + 1
    If ExportKulture(sheetID) Then successCount = successCount + 1
    If ExportParcele(sheetID) Then successCount = successCount + 1
    If ExportConfig(sheetID) Then successCount = successCount + 1
    If ExportUsers(sheetID) Then successCount = successCount + 1
    If ExportFakture(sheetID) Then successCount = successCount + 1
    If ExportFakturaStavke(sheetID) Then successCount = successCount + 1
    If ExportSaldoOMDetail(sheetID) Then successCount = successCount + 1
    If ExportStanice(sheetID) Then successCount = successCount + 1
    If ExportKupci(sheetID) Then successCount = successCount + 1
    If ExportVozaci(sheetID) Then successCount = successCount + 1
    If ExportArtikli(sheetID) Then successCount = successCount + 1
    If ExportMagacinKoop(sheetID) Then successCount = successCount + 1
    
    LogInfo "SyncStammdatenToGoogle_Core", _
            "Export abgeschlossen: " & successCount & "/" & TOTAL_STAMMDATEN_TABS & " Tabs"
    
    SyncStammdatenToGoogle_Core = (successCount = TOTAL_STAMMDATEN_TABS)
    
    If SyncStammdatenToGoogle_Core Then
        Monitor_StammdatenSyncSuccess _
            successCount:=successCount, _
            totalTabs:=TOTAL_STAMMDATEN_TABS, _
            sheetID:=sheetID
    Else
        Monitor_StammdatenSyncFail _
            errNum:=0, _
            errDesc:="Stammdaten partial export. SuccessTabs=" & CStr(successCount) & "/" & CStr(TOTAL_STAMMDATEN_TABS), _
            errSrc:="modStammdatenSync.SyncStammdatenToGoogle_Core", _
            successCount:=successCount, _
            totalTabs:=TOTAL_STAMMDATEN_TABS
    End If
    
    If showMessages Then
        MsgBox "Stammdaten exportiert: " & successCount & " od " & _
               TOTAL_STAMMDATEN_TABS & " tabova.", _
               IIf(SyncStammdatenToGoogle_Core, vbInformation, vbExclamation), _
               APP_NAME
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
    
    LogErr "SyncStammdatenToGoogle_Core"
    
    Monitor_StammdatenSyncFail _
        errNum:=errNum, _
        errDesc:=errDesc, _
        errSrc:=errSrc, _
        successCount:=successCount, _
        totalTabs:=TOTAL_STAMMDATEN_TABS
    
    If showMessages Then
        MsgBox "Gre" & ChrW(353) & "ka pri eksportu stammdaten: " & errDesc, vbCritical, APP_NAME
    End If
    
    SyncStammdatenToGoogle_Core = False
End Function

Public Sub ExportKarticeToGoogle()
    Call ExportKarticeToGoogle_Core(True)
End Sub

Public Function ExportKarticeToGoogle_Core(ByVal showMessages As Boolean) As Boolean
    Dim folderID As String
    Dim sheetID As String
    Dim koopData As Variant
    Dim colKoopID As Long
    Dim colAktivan As Long
    Dim i As Long
    Dim j As Long
    Dim allRows() As Variant
    Dim outRow As Long
    Dim totalRows As Long
    Dim datumOd As Date
    Dim datumDo As Date
    
    On Error GoTo EH
    
    ExportKarticeToGoogle_Core = False
    
    If Not IsGoogleAuthConfigured() Then
        LogError "ExportKarticeToGoogle_Core", "Google OAuth2 nije konfigurisan."
        If showMessages Then MsgBox "Google OAuth2 nije konfigurisan!", vbCritical, APP_NAME
        Exit Function
    End If
    
    folderID = GetReportsFolderID()
    If Len(Trim$(folderID)) = 0 Then
        LogError "ExportKarticeToGoogle_Core", "GOOGLE_REPORTS_FOLDER_ID / GOOGLE_PWA_FOLDER_ID nije postavljen."
        If showMessages Then MsgBox "GOOGLE_REPORTS_FOLDER_ID nije postavljen!", vbCritical, APP_NAME
        Exit Function
    End If
    
    sheetID = GetConfigValue("GOOGLE_KARTICE_SHEET_ID")
    If Len(Trim$(sheetID)) = 0 Then sheetID = GetSpreadsheetID("Kartice", folderID)
    
    If Len(Trim$(sheetID)) = 0 Then
        sheetID = CreateSpreadsheet("Kartice", folderID)
        If Len(sheetID) = 0 Then
            LogError "ExportKarticeToGoogle_Core", "Kartice Sheet could not be created."
            If showMessages Then MsgBox "Kartice Sheet konnte nicht erstellt werden!", vbCritical, APP_NAME
            Exit Function
        End If
    End If
    
    Call SetConfigValue("GOOGLE_KARTICE_SHEET_ID", sheetID)
    Call EnsureKarticeTabsBestEffort(sheetID)
    
    koopData = GetTableData(TBL_KOOPERANTI)
    
    If IsEmpty(koopData) Then
        ExportKarticeToGoogle_Core = WriteHeaderOnly(sheetID, KARTICE_TAB_NAME, _
            "KooperantID", "Datum", "BrojDok", "BrojParcele", _
            "Opis", "Zadu" & ChrW(382) & "enje", "Razdu" & ChrW(382) & "enje", "Saldo")
        
        If showMessages Then MsgBox "Kartice exportiert: 0 stavki.", vbInformation, APP_NAME
        Exit Function
    End If
    
    koopData = ExcludeStornirano(koopData, TBL_KOOPERANTI)
    
    If IsEmpty(koopData) Then
        ExportKarticeToGoogle_Core = WriteHeaderOnly(sheetID, KARTICE_TAB_NAME, _
            "KooperantID", "Datum", "BrojDok", "BrojParcele", _
            "Opis", "Zadu" & ChrW(382) & "enje", "Razdu" & ChrW(382) & "enje", "Saldo")
        
        If showMessages Then MsgBox "Kartice exportiert: 0 stavki.", vbInformation, APP_NAME
        Exit Function
    End If
    
    colKoopID = GetColumnIndex(TBL_KOOPERANTI, "KooperantID")
    colAktivan = GetColumnIndex(TBL_KOOPERANTI, "Aktivan")
    
    datumOd = DateSerial(Year(Date), 1, 1)
    datumDo = Date
    
    Dim koopList() As String
    Dim koopCount As Long
    
    ReDim koopList(1 To UBound(koopData, 1))
    
    For i = 1 To UBound(koopData, 1)
        If IsPWAActive(koopData(i, colAktivan)) Then
            koopCount = koopCount + 1
            koopList(koopCount) = CStr(koopData(i, colKoopID))
        End If
    Next i
    
    If koopCount = 0 Then
        ExportKarticeToGoogle_Core = WriteHeaderOnly(sheetID, KARTICE_TAB_NAME, _
            "KooperantID", "Datum", "BrojDok", "BrojParcele", _
            "Opis", "Zadu" & ChrW(382) & "enje", "Razdu" & ChrW(382) & "enje", "Saldo")
        
        If showMessages Then
            MsgBox "Kartice exportiert: 0 stavki za 0 aktivnih kooperanata.", _
                   vbInformation, APP_NAME
        End If
        
        Exit Function
    End If

    Dim karticaResults() As Variant
    ReDim karticaResults(1 To koopCount)
    
    totalRows = 1

    For i = 1 To koopCount
        karticaResults(i) = ReportKarticaKooperanta(koopList(i), datumOd, datumDo)
        If Not IsEmpty(karticaResults(i)) Then
            totalRows = totalRows + UBound(karticaResults(i), 1)
        End If
    Next i

    ReDim allRows(1 To totalRows, 1 To 8)
    
    allRows(1, 1) = "KooperantID"
    allRows(1, 2) = "Datum"
    allRows(1, 3) = "BrojDok"
    allRows(1, 4) = "BrojParcele"
    allRows(1, 5) = "Opis"
    allRows(1, 6) = "Zadu" & ChrW(382) & "enje"
    allRows(1, 7) = "Razdu" & ChrW(382) & "enje"
    allRows(1, 8) = "Saldo"
    
    outRow = 1
    
    For i = 1 To koopCount
        If Not IsEmpty(karticaResults(i)) Then
            Dim kData As Variant
            kData = karticaResults(i)
            
            For j = 1 To UBound(kData, 1)
                outRow = outRow + 1
                allRows(outRow, 1) = koopList(i)
                allRows(outRow, 2) = kData(j, 1)
                allRows(outRow, 3) = kData(j, 2)
                allRows(outRow, 4) = kData(j, 3)
                allRows(outRow, 5) = kData(j, 4)
                allRows(outRow, 6) = kData(j, 5)
                allRows(outRow, 7) = kData(j, 6)
                allRows(outRow, 8) = kData(j, 7)
            Next j
        End If
    Next i
    
    If outRow < totalRows Then
        Dim finalRows() As Variant
        Dim r As Long
        Dim c As Long
        
        ReDim finalRows(1 To outRow, 1 To 8)
        
        For r = 1 To outRow
            For c = 1 To 8
                finalRows(r, c) = allRows(r, c)
            Next c
        Next r
        
        ExportKarticeToGoogle_Core = WriteSheetData(sheetID, KARTICE_TAB_NAME, finalRows)
    Else
        ExportKarticeToGoogle_Core = WriteSheetData(sheetID, KARTICE_TAB_NAME, allRows)
    End If
    
    If ExportKarticeToGoogle_Core Then
        LogInfo "ExportKarticeToGoogle_Core", _
                CStr(outRow - 1) & " Zeilen fuer " & CStr(koopCount) & " Kooperanten"
    Else
        LogError "ExportKarticeToGoogle_Core", "WriteSheetData failed."
    End If
    
    If showMessages Then
        If ExportKarticeToGoogle_Core Then
            MsgBox "Kartice exportiert: " & (outRow - 1) & _
                   " stavki za " & koopCount & " kooperanata.", _
                   vbInformation, APP_NAME
        Else
            MsgBox "Kartice export nije uspeo. Proveri log.", vbExclamation, APP_NAME
        End If
    End If
    
    Exit Function

EH:
    LogErr "ExportKarticeToGoogle_Core"
    
    If showMessages Then
        MsgBox "Gre" & ChrW(353) & "ka: " & Err.description, vbCritical, APP_NAME
    End If
    
    ExportKarticeToGoogle_Core = False
End Function

Public Sub ExportMgmtReports()
    Call ExportMgmtReports_Core(True)
End Sub

Public Function ExportMgmtReports_Core(ByVal showMessages As Boolean) As Boolean
    Dim folderID As String
    Dim sheetID As String
    Dim ok As Long
    
    On Error GoTo EH
    
    ExportMgmtReports_Core = False
    
    If Not IsGoogleAuthConfigured() Then
        LogError "ExportMgmtReports_Core", "Google OAuth2 nije konfigurisan."
        If showMessages Then MsgBox "Google OAuth2 nije konfigurisan!", vbCritical, APP_NAME
        Exit Function
    End If
    
    folderID = GetReportsFolderID()
    
    If Len(Trim$(folderID)) = 0 Then
        LogError "ExportMgmtReports_Core", "GOOGLE_PWA_FOLDER_ID nije postavljen."
        If showMessages Then MsgBox "GOOGLE_PWA_FOLDER_ID nije postavljen!", vbCritical, APP_NAME
        Exit Function
    End If
    
    sheetID = GetConfigValue("GOOGLE_MGMT_SHEET_ID")
    If Len(Trim$(sheetID)) = 0 Then sheetID = GetSpreadsheetID("MgmtReports", folderID)
    
    If Len(Trim$(sheetID)) = 0 Then
        sheetID = CreateSpreadsheet("MgmtReports", folderID)
        
        If Len(sheetID) = 0 Then
            LogError "ExportMgmtReports_Core", "MgmtReports Sheet could not be created."
            If showMessages Then MsgBox "MgmtReports Sheet konnte nicht erstellt werden!", vbCritical, APP_NAME
            Exit Function
        End If
    End If
    
    Call SetConfigValue("GOOGLE_MGMT_SHEET_ID", sheetID)
    Call EnsureMgmtReportTabsBestEffort(sheetID)
    
    If ExportSaldoOM(sheetID) Then ok = ok + 1
    If ExportSaldoKupci(sheetID) Then ok = ok + 1
    If ExportOtkupPoOM(sheetID) Then ok = ok + 1
    If ExportPredatoPoKupcu(sheetID) Then ok = ok + 1
    If ExportOtkupiAll(sheetID) Then ok = ok + 1

    ExportMgmtReports_Core = (ok = 5)
    
    If ExportMgmtReports_Core Then
        LogInfo "ExportMgmtReports_Core", "MgmtReports export completed: 5/5"
    Else
        LogWarn "ExportMgmtReports_Core", "MgmtReports partial export: " & CStr(ok) & "/5"
    End If
    
    If showMessages Then
        If ExportMgmtReports_Core Then
            MsgBox "MgmtReports exportiert: 5/5", vbInformation, APP_NAME
        Else
            MsgBox "MgmtReports exportiert: " & CStr(ok) & "/5. Proveri log.", _
                   vbExclamation, APP_NAME
        End If
    End If
    
    Exit Function

EH:
    LogErr "ExportMgmtReports_Core"
    
    If showMessages Then
        MsgBox "Gre" & ChrW(353) & "ka: " & Err.description, vbCritical, APP_NAME
    End If
    
    ExportMgmtReports_Core = False
End Function
Private Function ExportOtkupPoOM(ByVal sheetID As String) As Boolean
    Dim data As Variant
    Dim colStanica As Long, colVrsta As Long, colKlasa As Long
    Dim colKolicina As Long, colAmb As Long, colCena As Long
    Dim i As Long
    
    On Error GoTo EH
    
    data = GetTableData(TBL_OTKUP)
    If Not IsEmpty(data) Then data = ExcludeStornirano(data, TBL_OTKUP)
    If IsEmpty(data) Then
        ExportOtkupPoOM = WriteHeaderOnly(sheetID, "OtkupPoOM", _
            "StanicaID", "VrstaVoca", "Klasa", "Koli" & ChrW(269) & "ina", _
            "Ambala" & ChrW(382) & "a", "Vrednost", "BrojOtkupa")
        Exit Function
    End If
    
    colStanica = GetColumnIndex(TBL_OTKUP, COL_OTK_STANICA)
    colVrsta = GetColumnIndex(TBL_OTKUP, COL_OTK_VRSTA)
    colKlasa = GetColumnIndex(TBL_OTKUP, COL_OTK_KLASA)
    colKolicina = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
    colAmb = GetColumnIndex(TBL_OTKUP, COL_OTK_KOL_AMB)
    colCena = GetColumnIndex(TBL_OTKUP, COL_OTK_CENA)
    
    ' Aggregieren per Stanica + Vrsta + Klasa
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    For i = 1 To UBound(data, 1)
        Dim key As String
        key = CStr(data(i, colStanica)) & "|" & CStr(data(i, colVrsta)) & "|" & CStr(data(i, colKlasa))
        
        If Not dict.Exists(key) Then dict.Add key, Array(0#, 0#, 0#, 0#) ' Kg, Amb, Vrednost, BrojOtkupa
        Dim vals As Variant
        vals = dict(key)
        vals(0) = vals(0) + CDbl(data(i, colKolicina))
        vals(1) = vals(1) + CDbl(nz(data(i, colAmb), 0))
        vals(2) = vals(2) + CDbl(data(i, colKolicina)) * CDbl(data(i, colCena))
        vals(3) = vals(3) + 1
        dict(key) = vals
    Next i
    
    If dict.count = 0 Then
        ExportOtkupPoOM = WriteHeaderOnly(sheetID, "OtkupPoOM", _
            "StanicaID", "VrstaVoca", "Klasa", "Koli" & ChrW(269) & "ina", _
            "Ambala" & ChrW(382) & "a", "Vrednost", "BrojOtkupa")
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To dict.count + 1, 1 To 7)
    result(1, 1) = "StanicaID"
    result(1, 2) = "VrstaVoca"
    result(1, 3) = "Klasa"
    result(1, 4) = "Koli" & ChrW(269) & "ina"
    result(1, 5) = "Ambala" & ChrW(382) & "a"
    result(1, 6) = "Vrednost"
    result(1, 7) = "BrojOtkupa"
    
    Dim keys As Variant
    keys = dict.keys
    Dim r As Long
    For r = 0 To dict.count - 1
        Dim parts() As String
        parts = Split(keys(r), "|")
        vals = dict(keys(r))
        result(r + 2, 1) = parts(0)
        result(r + 2, 2) = parts(1)
        result(r + 2, 3) = parts(2)
        result(r + 2, 4) = CStr(vals(0))
        result(r + 2, 5) = CStr(vals(1))
        result(r + 2, 6) = CStr(vals(2))
        result(r + 2, 7) = CStr(vals(3))
    Next r
    
    ExportOtkupPoOM = WriteSheetData(sheetID, "OtkupPoOM", result)
    Exit Function
EH:
    LogErr "ExportOtkupPoOM"
    ExportOtkupPoOM = False
End Function

Private Function ExportOtkupiAll(ByVal sheetID As String) As Boolean
    Const TAB_NAME As String = "OtkupiAll"

    Dim data As Variant
    Dim result() As Variant
    Dim i As Long
    Dim outRow As Long

    Dim colID As Long
    Dim colDatum As Long
    Dim colKoop As Long
    Dim colStanica As Long
    Dim colVrsta As Long
    Dim colSorta As Long
    Dim colKlasa As Long
    Dim colKolicina As Long
    Dim colCena As Long
    Dim colTipAmb As Long
    Dim colKolAmb As Long
    Dim colVozac As Long
    Dim colBrDok As Long
    Dim colParcela As Long
    Dim colBrojZbirne As Long
    Dim colOtpremnicaID As Long
    Dim prjIndex As Object

    On Error GoTo EH

    data = GetTableData(TBL_OTKUP)

    If Not IsEmpty(data) Then
        data = ExcludeStornirano(data, TBL_OTKUP)
    End If

    If IsEmpty(data) Then
        ExportOtkupiAll = WriteHeaderOnly(sheetID, TAB_NAME, _
            "ClientRecordID", "ServerRecordID", "CreatedAtClient", _
            "UpdatedAtClient", "UpdatedAtServer", "SyncStatus", _
            "DeviceID", "OtkupacID", "Datum", "KooperantID", _
            "KooperantName", "VrstaVoca", "SortaVoca", "Klasa", _
            "Koli" & ChrW(269) & "ina", "Cena", "TipAmbalaze", "KolAmbalaze", _
            "ParcelaID", "VozacID", "Napomena", "ReceivedAt", _
            "BrojZbirne", "OtpremnicaID", "PrijemnicaID", _
            "BrojPrijemnice", "KupacID", "DatumPrijema", _
            "Primljeno", "TransportStatus")
        Exit Function
    End If

    colID = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, "ExportOtkupiAll")
    colDatum = RequireColumnIndex(TBL_OTKUP, COL_OTK_DATUM, "ExportOtkupiAll")
    colKoop = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT, "ExportOtkupiAll")
    colStanica = RequireColumnIndex(TBL_OTKUP, COL_OTK_STANICA, "ExportOtkupiAll")
    colVrsta = RequireColumnIndex(TBL_OTKUP, COL_OTK_VRSTA, "ExportOtkupiAll")
    colSorta = RequireColumnIndex(TBL_OTKUP, COL_OTK_SORTA, "ExportOtkupiAll")
    colKlasa = RequireColumnIndex(TBL_OTKUP, COL_OTK_KLASA, "ExportOtkupiAll")
    colKolicina = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA, "ExportOtkupiAll")
    colCena = RequireColumnIndex(TBL_OTKUP, COL_OTK_CENA, "ExportOtkupiAll")
    colTipAmb = RequireColumnIndex(TBL_OTKUP, COL_OTK_TIP_AMB, "ExportOtkupiAll")
    colKolAmb = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOL_AMB, "ExportOtkupiAll")
    colVozac = RequireColumnIndex(TBL_OTKUP, COL_OTK_VOZAC, "ExportOtkupiAll")
    colBrDok = RequireColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK, "ExportOtkupiAll")
    colParcela = RequireColumnIndex(TBL_OTKUP, COL_OTK_PARCELA, "ExportOtkupiAll")
    colBrojZbirne = RequireColumnIndex(TBL_OTKUP, COL_OTK_BROJ_ZBIRNE, "ExportOtkupiAll")
    colOtpremnicaID = RequireColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID, "ExportOtkupiAll")

    Set prjIndex = BuildPrijemnicaIndexByBrojZbirne()

    ReDim result(1 To UBound(data, 1) + 1, 1 To 30)

    result(1, 1) = "ClientRecordID"
    result(1, 2) = "ServerRecordID"
    result(1, 3) = "CreatedAtClient"
    result(1, 4) = "UpdatedAtClient"
    result(1, 5) = "UpdatedAtServer"
    result(1, 6) = "SyncStatus"
    result(1, 7) = "DeviceID"
    result(1, 8) = "OtkupacID"
    result(1, 9) = "Datum"
    result(1, 10) = "KooperantID"
    result(1, 11) = "KooperantName"
    result(1, 12) = "VrstaVoca"
    result(1, 13) = "SortaVoca"
    result(1, 14) = "Klasa"
    result(1, 15) = "Koli" & ChrW(269) & "ina"
    result(1, 16) = "Cena"
    result(1, 17) = "TipAmbalaze"
    result(1, 18) = "KolAmbalaze"
    result(1, 19) = "ParcelaID"
    result(1, 20) = "VozacID"
    result(1, 21) = "Napomena"
    result(1, 22) = "ReceivedAt"
    result(1, 23) = "BrojZbirne"
    result(1, 24) = "OtpremnicaID"
    result(1, 25) = "PrijemnicaID"
    result(1, 26) = "BrojPrijemnice"
    result(1, 27) = "KupacID"
    result(1, 28) = "DatumPrijema"
    result(1, 29) = "Primljeno"
    result(1, 30) = "TransportStatus"

    outRow = 1

    For i = 1 To UBound(data, 1)
        Dim otkupID As String
        Dim koopID As String
        Dim koopName As String
        Dim brojZbirne As String
        Dim otpremnicaID As String
        Dim prijemnicaID As String
        Dim brojPrijemnice As String
        Dim kupacID As String
        Dim datumPrijema As String
        Dim primljeno As String
        Dim transportStatus As String
        Dim prjInfo As Variant

        otkupID = CStr(nz(data(i, colID), ""))
        koopID = CStr(nz(data(i, colKoop), ""))
        koopName = GetKooperantDisplayNameForExport(koopID)
        
        brojZbirne = CStr(nz(data(i, colBrojZbirne), ""))
        otpremnicaID = CStr(nz(data(i, colOtpremnicaID), ""))

        prijemnicaID = ""
        brojPrijemnice = ""
        kupacID = ""
        datumPrijema = ""
        primljeno = "Ne"

        If Len(Trim$(brojZbirne)) > 0 Then
            If Not prjIndex Is Nothing Then
                If prjIndex.Exists(brojZbirne) Then
                    prjInfo = prjIndex(brojZbirne)

                    prijemnicaID = CStr(prjInfo(0))
                    brojPrijemnice = CStr(prjInfo(1))
                    kupacID = CStr(prjInfo(2))
                    datumPrijema = CStr(prjInfo(3))
                    primljeno = "Da"
                End If
            End If
        End If

        If primljeno = "Da" Then
            transportStatus = "received"
        ElseIf Len(Trim$(otpremnicaID)) > 0 Or Len(Trim$(brojZbirne)) > 0 Then
            transportStatus = "in_transport"
        ElseIf Len(Trim$(CStr(nz(data(i, colVozac), "")))) > 0 Then
            transportStatus = "assigned"
        Else
            transportStatus = "unassigned"
        End If

        outRow = outRow + 1

        result(outRow, 1) = "VBA-" & otkupID
        result(outRow, 2) = otkupID
        result(outRow, 3) = ""
        result(outRow, 4) = ""
        result(outRow, 5) = Now
        result(outRow, 6) = "Synced>Master"
        result(outRow, 7) = "VBA"
        result(outRow, 8) = CStr(nz(data(i, colStanica), ""))
        result(outRow, 9) = data(i, colDatum)
        result(outRow, 10) = koopID
        result(outRow, 11) = koopName
        result(outRow, 12) = CStr(nz(data(i, colVrsta), ""))
        result(outRow, 13) = CStr(nz(data(i, colSorta), ""))
        result(outRow, 14) = CStr(nz(data(i, colKlasa), "I"))
        result(outRow, 15) = CDbl(nz(data(i, colKolicina), 0))
        result(outRow, 16) = CDbl(nz(data(i, colCena), 0))
        result(outRow, 17) = CStr(nz(data(i, colTipAmb), ""))
        result(outRow, 18) = CLng(nz(data(i, colKolAmb), 0))
        result(outRow, 19) = CStr(nz(data(i, colParcela), ""))
        result(outRow, 20) = CStr(nz(data(i, colVozac), ""))
        result(outRow, 21) = CStr(nz(data(i, colBrDok), ""))
        result(outRow, 22) = Now
        result(outRow, 23) = brojZbirne
        result(outRow, 24) = otpremnicaID
        result(outRow, 25) = prijemnicaID
        result(outRow, 26) = brojPrijemnice
        result(outRow, 27) = kupacID
        result(outRow, 28) = datumPrijema
        result(outRow, 29) = primljeno
        result(outRow, 30) = transportStatus
    Next i

    ExportOtkupiAll = WriteSheetData(sheetID, TAB_NAME, result)
    Exit Function

EH:
    LogErr "ExportOtkupiAll"
    ExportOtkupiAll = False
End Function

Private Function BuildPrijemnicaIndexByBrojZbirne() As Object
    Const SOURCE As String = "BuildPrijemnicaIndexByBrojZbirne"

    Dim dict As Object
    Dim data As Variant
    Dim i As Long

    Dim colPrjID As Long
    Dim colBrojPrijemnice As Long
    Dim colBrojZbirne As Long
    Dim colKupac As Long
    Dim colDatum As Long

    On Error GoTo EH

    Set dict = CreateObject("Scripting.Dictionary")
    Set BuildPrijemnicaIndexByBrojZbirne = dict

    data = GetTableData(TBL_PRIJEMNICA)

    If Not IsEmpty(data) Then
        data = ExcludeStornirano(data, TBL_PRIJEMNICA)
    End If

    If IsEmpty(data) Then Exit Function

    colPrjID = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID, SOURCE)
    colBrojPrijemnice = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ, SOURCE)
    colBrojZbirne = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, SOURCE)
    colKupac = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KUPAC, SOURCE)
    colDatum = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_DATUM, SOURCE)

    For i = 1 To UBound(data, 1)
        Dim bz As String
        Dim prjID As String
        Dim brojPrj As String
        Dim kupacID As String
        Dim datumPrj As String

        bz = Trim$(CStr(nz(data(i, colBrojZbirne), "")))

        If Len(bz) > 0 Then
            prjID = CStr(nz(data(i, colPrjID), ""))
            brojPrj = CStr(nz(data(i, colBrojPrijemnice), ""))
            kupacID = CStr(nz(data(i, colKupac), ""))

            If IsDate(data(i, colDatum)) Then
                datumPrj = Format$(CDate(data(i, colDatum)), "yyyy-mm-dd")
            Else
                datumPrj = CStr(nz(data(i, colDatum), ""))
            End If

            ' Ako postoji vise prijemnica za isti BrojZbirne, prva je dovoljna
            ' za status "received". Kasnije mozemo agregirati ako bude trebalo.
            If Not dict.Exists(bz) Then
                dict.Add bz, Array(prjID, brojPrj, kupacID, datumPrj)
            End If
        End If
    Next i

    Exit Function

EH:
    LogErr SOURCE
    Set BuildPrijemnicaIndexByBrojZbirne = dict
End Function
Private Function GetKooperantDisplayNameForExport(ByVal kooperantID As String) As String
    On Error GoTo EH

    Dim ime As Variant
    Dim prezime As Variant

    ime = LookupValue(TBL_KOOPERANTI, "KooperantID", kooperantID, "Ime")
    prezime = LookupValue(TBL_KOOPERANTI, "KooperantID", kooperantID, "Prezime")

    GetKooperantDisplayNameForExport = Trim$(CStr(nz(ime, "")) & " " & CStr(nz(prezime, "")))

    If Len(GetKooperantDisplayNameForExport) = 0 Then
        GetKooperantDisplayNameForExport = kooperantID
    End If

    Exit Function

EH:
    GetKooperantDisplayNameForExport = kooperantID
End Function

Private Function ExportPredatoPoKupcu(ByVal sheetID As String) As Boolean
    Dim data As Variant
    Dim colKupac As Long, colVrsta As Long, colKlasa As Long
    Dim colKolicina As Long, colAmb As Long, colCena As Long
    Dim i As Long
    
    On Error GoTo EH
    
    data = GetTableData(TBL_PRIJEMNICA)
    If Not IsEmpty(data) Then data = ExcludeStornirano(data, TBL_PRIJEMNICA)
    If IsEmpty(data) Then
        ExportPredatoPoKupcu = WriteHeaderOnly(sheetID, "PredatoPoKupcu", _
            "KupacID", "VrstaVoca", "Klasa", "Koli" & ChrW(269) & "ina", _
            "Ambala" & ChrW(382) & "a", "Vrednost", "BrojPrijemnica")
        Exit Function
    End If
    
    colKupac = GetColumnIndex(TBL_PRIJEMNICA, "KupacID")
    colVrsta = GetColumnIndex(TBL_PRIJEMNICA, "VrstaVoca")
    colKlasa = GetColumnIndex(TBL_PRIJEMNICA, "Klasa")
    colKolicina = GetColumnIndex(TBL_PRIJEMNICA, "Kolicina")
    colAmb = GetColumnIndex(TBL_PRIJEMNICA, "KolAmbalaze")
    colCena = GetColumnIndex(TBL_PRIJEMNICA, "Cena")
    
    ' Aggregieren per Kupac + Vrsta + Klasa
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    For i = 1 To UBound(data, 1)
        Dim key As String
        key = CStr(data(i, colKupac)) & "|" & CStr(data(i, colVrsta)) & "|" & CStr(data(i, colKlasa))
        
        If Not dict.Exists(key) Then dict.Add key, Array(0#, 0#, 0#, 0#) ' Kg, Amb, Vrednost, BrojPrijemnica
        Dim vals As Variant
        vals = dict(key)
        vals(0) = vals(0) + CDbl(data(i, colKolicina))
        vals(1) = vals(1) + CDbl(nz(data(i, colAmb), 0))
        vals(2) = vals(2) + CDbl(data(i, colKolicina)) * CDbl(data(i, colCena))
        vals(3) = vals(3) + 1
        dict(key) = vals
    Next i
    
    If dict.count = 0 Then
        ExportPredatoPoKupcu = WriteHeaderOnly(sheetID, "PredatoPoKupcu", _
            "KupacID", "VrstaVoca", "Klasa", "Koli" & ChrW(269) & "ina", _
            "Ambala" & ChrW(382) & "a", "Vrednost", "BrojPrijemnica")
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To dict.count + 1, 1 To 7)
    result(1, 1) = "KupacID"
    result(1, 2) = "VrstaVoca"
    result(1, 3) = "Klasa"
    result(1, 4) = "Koli" & ChrW(269) & "ina"
    result(1, 5) = "Ambala" & ChrW(382) & "a"
    result(1, 6) = "Vrednost"
    result(1, 7) = "BrojPrijemnica"
    
    Dim keys As Variant
    keys = dict.keys
    Dim r As Long
    For r = 0 To dict.count - 1
        Dim parts() As String
        parts = Split(keys(r), "|")
        vals = dict(keys(r))
        
        Dim kupacNaziv As Variant
        kupacNaziv = LookupValue(TBL_KUPCI, "KupacID", parts(0), "Naziv")
        
        result(r + 2, 1) = CStr(nz(kupacNaziv, parts(0)))
        result(r + 2, 2) = parts(1)
        result(r + 2, 3) = parts(2)
        result(r + 2, 4) = CStr(vals(0))
        result(r + 2, 5) = CStr(vals(1))
        result(r + 2, 6) = CStr(vals(2))
        result(r + 2, 7) = CStr(vals(3))
    Next r
    
    ExportPredatoPoKupcu = WriteSheetData(sheetID, "PredatoPoKupcu", result)
    Exit Function
EH:
    LogErr "ExportPredatoPoKupcu"
    ExportPredatoPoKupcu = False
End Function

Private Function ExportSaldoOM(ByVal sheetID As String) As Boolean
    Dim lstSaldo As Object
    
    On Error GoTo EH
    
    ' ReportSaldoOM gibt Daten in ein ListBox -- wir brauchen die Rohdaten
    ' Hier vereinfacht: OM-Saldo aus tblNovac berechnen
    Dim data As Variant
    Dim colOMID As Long, colTip As Long, colIsplata As Long, colUplata As Long
    Dim i As Long
    
    data = GetTableData(TBL_NOVAC)
    If Not IsEmpty(data) Then data = ExcludeStornirano(data, TBL_NOVAC)
    If IsEmpty(data) Then
        ExportSaldoOM = WriteHeaderOnly(sheetID, "SaldoOM", _
            "StanicaID", "Avans", "Isplaceno", "Saldo")
        Exit Function
    End If
    
    colOMID = GetColumnIndex(TBL_NOVAC, COL_NOV_OM_ID)
    colTip = GetColumnIndex(TBL_NOVAC, COL_NOV_TIP)
    colIsplata = GetColumnIndex(TBL_NOVAC, COL_NOV_ISPLATA)
    colUplata = GetColumnIndex(TBL_NOVAC, COL_NOV_UPLATA)
    
    ' Aggregieren per OM
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    For i = 1 To UBound(data, 1)
        Dim omID As String
        omID = Trim$(CStr(data(i, colOMID)))
        If Len(omID) > 0 Then
            Dim tip As String
            tip = CStr(data(i, colTip))
            
            If Not dict.Exists(omID) Then dict.Add omID, Array(0#, 0#) ' (Avans, Isplaceno)
            Dim vals As Variant
            vals = dict(omID)
            
            If tip = NOV_KES_FIRMA_OTKUPAC Then
                vals(0) = vals(0) + CDbl(data(i, colIsplata))
            ElseIf tip = NOV_KES_OTKUPAC_KOOP Then
                vals(1) = vals(1) + CDbl(data(i, colIsplata))
            End If
            
            dict(omID) = vals
        End If
    Next i
    
    If dict.count = 0 Then
        ExportSaldoOM = WriteHeaderOnly(sheetID, "SaldoOM", _
            "StanicaID", "Avans", "Isplaceno", "Saldo")
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To dict.count + 1, 1 To 4)
    result(1, 1) = "StanicaID"
    result(1, 2) = "Avans"
    result(1, 3) = "Isplaceno"
    result(1, 4) = "Saldo"
    
    Dim keys As Variant
    keys = dict.keys
    Dim r As Long
    For r = 0 To dict.count - 1
        vals = dict(keys(r))
        result(r + 2, 1) = keys(r)
        result(r + 2, 2) = CStr(vals(0))
        result(r + 2, 3) = CStr(vals(1))
        result(r + 2, 4) = CStr(vals(0) - vals(1))
    Next r
    
    ExportSaldoOM = WriteSheetData(sheetID, "SaldoOM", result)
    Exit Function
EH:
    LogErr "ExportSaldoOM"
    ExportSaldoOM = False
End Function

Private Function ExportSaldoOMDetail(ByVal sheetID As String) As Boolean
    Dim otkData As Variant, novData As Variant, magData As Variant
    Dim i As Long
    
    On Error GoTo EH
    
    ' --- OTKUP: Kolicina, Vrednost per Kooperant ---
    otkData = GetTableData(TBL_OTKUP)
    If Not IsEmpty(otkData) Then otkData = ExcludeStornirano(otkData, TBL_OTKUP)
    
    Dim colOtkKoop As Long, colOtkSta As Long, colOtkKg As Long
    Dim colOtkCena As Long, colOtkAmb As Long
    colOtkKoop = GetColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT)
    colOtkSta = GetColumnIndex(TBL_OTKUP, COL_OTK_STANICA)
    colOtkKg = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
    colOtkCena = GetColumnIndex(TBL_OTKUP, COL_OTK_CENA)
    colOtkAmb = GetColumnIndex(TBL_OTKUP, COL_OTK_KOL_AMB)
    
    ' Dict: KoopID ? (StanicaID, Kolicina, Vrednost, AmbOtkup)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    If Not IsEmpty(otkData) Then
        For i = 1 To UBound(otkData, 1)
            Dim koopID As String, staID As String
            koopID = CStr(otkData(i, colOtkKoop))
            staID = CStr(otkData(i, colOtkSta))
            If Len(koopID) > 0 Then
                If Not dict.Exists(koopID) Then
                    ' (StanicaID, Kolicina, Vrednost, AmbOtkup, Isplaceno, AgroZaduzenje)
                    dict.Add koopID, Array(staID, 0#, 0#, 0#, 0#, 0#)
                End If
                Dim v As Variant
                v = dict(koopID)
                v(1) = v(1) + CDbl(nz(otkData(i, colOtkKg), 0))
                v(2) = v(2) + CDbl(nz(otkData(i, colOtkKg), 0)) * CDbl(nz(otkData(i, colOtkCena), 0))
                v(3) = v(3) + CDbl(nz(otkData(i, colOtkAmb), 0))
                dict(koopID) = v
            End If
        Next i
    End If
    
    ' --- NOVAC: Isplaceno per Kooperant ---
    novData = GetTableData(TBL_NOVAC)
    If Not IsEmpty(novData) Then novData = ExcludeStornirano(novData, TBL_NOVAC)
    
    If Not IsEmpty(novData) Then
        Dim colNovKoop As Long, colNovTip As Long, colNovIsplata As Long
        colNovKoop = GetColumnIndex(TBL_NOVAC, COL_NOV_KOOP_ID)
        colNovTip = GetColumnIndex(TBL_NOVAC, COL_NOV_TIP)
        colNovIsplata = GetColumnIndex(TBL_NOVAC, COL_NOV_ISPLATA)
        
        For i = 1 To UBound(novData, 1)
            Dim tip As String
            tip = CStr(novData(i, colNovTip))
            If tip = NOV_KES_OTKUPAC_KOOP Or tip = NOV_VIRMAN_FIRMA_KOOP Or tip = NOV_VIRMAN_AVANS_KOOP Then
                Dim nKoop As String
                nKoop = CStr(nz(novData(i, colNovKoop), ""))
                If dict.Exists(nKoop) Then
                    v = dict(nKoop)
                    v(4) = v(4) + CDbl(nz(novData(i, colNovIsplata), 0))
                    dict(nKoop) = v
                End If
            End If
        Next i
    End If
    
    ' --- MAGACIN: Agro Zaduzenje per Kooperant ---
    magData = GetTableData(TBL_MAGACIN)
    If Not IsEmpty(magData) Then magData = ExcludeStornirano(magData, TBL_MAGACIN)
    
    If Not IsEmpty(magData) Then
        Dim colMagKoop As Long, colMagTip As Long, colMagVrednost As Long
        colMagKoop = GetColumnIndex(TBL_MAGACIN, "KooperantID")
        colMagTip = GetColumnIndex(TBL_MAGACIN, "Tip")
        colMagVrednost = GetColumnIndex(TBL_MAGACIN, "Vrednost")
        
        For i = 1 To UBound(magData, 1)
            If CStr(magData(i, colMagTip)) = MAG_IZLAZ Then
                Dim mKoop As String
                mKoop = CStr(nz(magData(i, colMagKoop), ""))
                If dict.Exists(mKoop) Then
                    v = dict(mKoop)
                    v(5) = v(5) + CDbl(nz(magData(i, colMagVrednost), 0))
                    dict(mKoop) = v
                End If
            End If
        Next i
    End If
    
    If dict.count = 0 Then
        ExportSaldoOMDetail = WriteHeaderOnly(sheetID, "SaldoOMDetail", _
            "KooperantID", "Kooperant", "StanicaID", "Koli" & ChrW(269) & "ina", _
            "Vrednost", "Isplaceno", "AgroZaduzenje", "Saldo", "Ambala" & ChrW(382) & "a")
        Exit Function
    End If
    
    ' --- Build result ---
    Dim result() As Variant
    ReDim result(1 To dict.count + 1, 1 To 9)
    result(1, 1) = "KooperantID"
    result(1, 2) = "Kooperant"
    result(1, 3) = "StanicaID"
    result(1, 4) = "Koli" & ChrW(269) & "ina"
    result(1, 5) = "Vrednost"
    result(1, 6) = "Isplaceno"
    result(1, 7) = "AgroZaduzenje"
    result(1, 8) = "Saldo"
    result(1, 9) = "Ambala" & ChrW(382) & "a"
    
    Dim keys As Variant
    keys = dict.keys
    Dim r As Long
    For r = 0 To dict.count - 1
        v = dict(keys(r))
        Dim koopName As Variant
        koopName = LookupValue(TBL_KOOPERANTI, "KooperantID", keys(r), "Ime")
        Dim koopPrezime As Variant
        koopPrezime = LookupValue(TBL_KOOPERANTI, "KooperantID", keys(r), "Prezime")
        
        Dim saldo As Double
        saldo = v(2) - v(4) - v(5) ' Vrednost - Isplaceno - AgroZaduzenje
        
        result(r + 2, 1) = keys(r)
        result(r + 2, 2) = CStr(nz(koopName, "")) & " " & CStr(nz(koopPrezime, ""))
        result(r + 2, 3) = CStr(v(0))
        result(r + 2, 4) = CStr(v(1))
        result(r + 2, 5) = CStr(v(2))
        result(r + 2, 6) = CStr(v(4))
        result(r + 2, 7) = CStr(v(5))
        result(r + 2, 8) = CStr(saldo)
        result(r + 2, 9) = CStr(v(3))
    Next r
    
    ExportSaldoOMDetail = WriteSheetData(sheetID, "SaldoOMDetail", result)
    Exit Function
EH:
    LogErr "ExportSaldoOMDetail"
    ExportSaldoOMDetail = False
End Function

Private Function ExportSaldoKupci(ByVal sheetID As String) As Boolean
    Dim data As Variant
    Dim colKupac As Long, colIznos As Long, colStatus As Long
    Dim novData As Variant
    Dim colNovPartner As Long, colNovTip As Long, colNovUplata As Long
    Dim i As Long
    
    On Error GoTo EH
    
    ' Fakture laden
    data = GetTableData(TBL_FAKTURE)
    If Not IsEmpty(data) Then data = ExcludeStornirano(data, TBL_FAKTURE)
    If IsEmpty(data) Then
        ExportSaldoKupci = WriteHeaderOnly(sheetID, "SaldoKupci", _
            "KupacID", "Kupac", "Fakturisano", "Placeno", "Saldo")
        Exit Function
    End If
    
    colKupac = GetColumnIndex(TBL_FAKTURE, "KupacID")
    colIznos = GetColumnIndex(TBL_FAKTURE, "Iznos")
    
    ' Novac laden (Kupci Uplate)
    novData = GetTableData(TBL_NOVAC)
    If Not IsEmpty(novData) Then novData = ExcludeStornirano(novData, TBL_NOVAC)
    
    colNovPartner = GetColumnIndex(TBL_NOVAC, COL_NOV_PARTNER_ID)
    colNovTip = GetColumnIndex(TBL_NOVAC, COL_NOV_TIP)
    colNovUplata = GetColumnIndex(TBL_NOVAC, COL_NOV_UPLATA)
    
    ' Aggregieren per Kupac
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    For i = 1 To UBound(data, 1)
        Dim kupacID As String
        kupacID = CStr(data(i, colKupac))
        If Not dict.Exists(kupacID) Then dict.Add kupacID, Array(0#, 0#) ' (Fakturisano, Placeno)
        Dim vals As Variant
        vals = dict(kupacID)
        vals(0) = vals(0) + CDbl(data(i, colIznos))
        dict(kupacID) = vals
    Next i
    
    ' Uplate
    If Not IsEmpty(novData) Then
        For i = 1 To UBound(novData, 1)
            Dim tip As String
            tip = CStr(novData(i, colNovTip))
            If tip = NOV_KUPCI_UPLATA Or tip = NOV_KUPCI_AVANS Then
                Dim pid As String
                pid = CStr(novData(i, colNovPartner))
                If dict.Exists(pid) Then
                    vals = dict(pid)
                    vals(1) = vals(1) + CDbl(novData(i, colNovUplata))
                    dict(pid) = vals
                End If
            End If
        Next i
    End If
    
    If dict.count = 0 Then
        ExportSaldoKupci = WriteHeaderOnly(sheetID, "SaldoKupci", _
            "KupacID", "Kupac", "Fakturisano", "Placeno", "Saldo")
        Exit Function
    End If
    
    ' Kupac-Namen holen
    Dim result() As Variant
    ReDim result(1 To dict.count + 1, 1 To 5)
    result(1, 1) = "KupacID"
    result(1, 2) = "Kupac"
    result(1, 3) = "Fakturisano"
    result(1, 4) = "Placeno"
    result(1, 5) = "Saldo"
    
    Dim keys As Variant
    keys = dict.keys
    Dim r As Long
    For r = 0 To dict.count - 1
        vals = dict(keys(r))
        Dim kupacNaziv As Variant
        kupacNaziv = LookupValue(TBL_KUPCI, "KupacID", keys(r), "Naziv")
        
        result(r + 2, 1) = keys(r)
        result(r + 2, 2) = CStr(nz(kupacNaziv, keys(r)))
        result(r + 2, 3) = CStr(vals(0))
        result(r + 2, 4) = CStr(vals(1))
        result(r + 2, 5) = CStr(vals(0) - vals(1))
    Next r
    
    ExportSaldoKupci = WriteSheetData(sheetID, "SaldoKupci", result)
    Exit Function
EH:
    LogErr "ExportSaldoKupci"
    ExportSaldoKupci = False
End Function

' ============================================================
' PRIVATE -- Export einzelner Tabellen
' ============================================================

Private Function ExportKooperanti(ByVal sheetID As String) As Boolean
    Dim data As Variant
    Dim result() As Variant
    Dim colID As Long, colIme As Long, colPrezime As Long
    Dim colStanica As Long, colAktivan As Long, colBPG As Long
    Dim colTelefon As Long, colMesto As Long
    Dim i As Long, outRow As Long
    Dim colAdresa As Long, colJMBG As Long
    
    On Error GoTo EH
    
    data = GetTableData(TBL_KOOPERANTI)
    If IsEmpty(data) Then
        ExportKooperanti = WriteHeaderOnly(sheetID, "Kooperanti", _
            "KooperantID", "Ime", "Prezime", "StanicaID", "Mesto", _
            "Telefon", "BPGBroj", "Adresa", "JMBG")
        Exit Function
    End If
    
    data = ExcludeStornirano(data, TBL_KOOPERANTI)
    If IsEmpty(data) Then
        ExportKooperanti = WriteHeaderOnly(sheetID, "Kooperanti", _
            "KooperantID", "Ime", "Prezime", "StanicaID", "Mesto", _
            "Telefon", "BPGBroj", "Adresa", "JMBG")
        Exit Function
    End If
    
    colID = GetColumnIndex(TBL_KOOPERANTI, "KooperantID")
    colIme = GetColumnIndex(TBL_KOOPERANTI, "Ime")
    colPrezime = GetColumnIndex(TBL_KOOPERANTI, "Prezime")
    colStanica = GetColumnIndex(TBL_KOOPERANTI, COL_KOOP_STANICA)
    colAktivan = GetColumnIndex(TBL_KOOPERANTI, "Aktivan")
    colMesto = GetColumnIndex(TBL_KOOPERANTI, "Mesto")
    colTelefon = GetColumnIndex(TBL_KOOPERANTI, "Telefon")
    colBPG = GetColumnIndex(TBL_KOOPERANTI, COL_KOOP_BPG)
    colAdresa = GetColumnIndex(TBL_KOOPERANTI, "Adresa")
    colJMBG = GetColumnIndex(TBL_KOOPERANTI, "JMBG")
    
    ' Nur aktive Kooperanten
    Dim activeCount As Long
    For i = 1 To UBound(data, 1)
        If IsPWAActive(data(i, colAktivan)) Then activeCount = activeCount + 1
    Next i
    
    ReDim result(1 To activeCount + 1, 1 To 9)
    
    ' Header
    result(1, 1) = "KooperantID"
    result(1, 2) = "Ime"
    result(1, 3) = "Prezime"
    result(1, 4) = "StanicaID"
    result(1, 5) = "Mesto"
    result(1, 6) = "Telefon"
    result(1, 7) = "BPGBroj"
    result(1, 8) = "Adresa"
    result(1, 9) = "JMBG"
    
    outRow = 1
    For i = 1 To UBound(data, 1)
        If IsPWAActive(data(i, colAktivan)) Then
            outRow = outRow + 1
            result(outRow, 1) = CStr(data(i, colID))
            result(outRow, 2) = CStr(data(i, colIme))
            result(outRow, 3) = CStr(data(i, colPrezime))
            result(outRow, 4) = CStr(data(i, colStanica))
            result(outRow, 5) = CStr(nz(data(i, colMesto), ""))
            result(outRow, 6) = CStr(nz(data(i, colTelefon), ""))
            result(outRow, 7) = CStr(nz(data(i, colBPG), ""))
            result(outRow, 8) = CStr(nz(data(i, colAdresa), ""))
            result(outRow, 9) = CStr(nz(data(i, colJMBG), ""))
        End If
    Next i
    
    ExportKooperanti = WriteSheetData(sheetID, "Kooperanti", result)
    Exit Function

EH:
    LogErr "ExportKooperanti"
    ExportKooperanti = False
End Function

Private Function ExportKulture(ByVal sheetID As String) As Boolean
    Dim data As Variant
    Dim result() As Variant
    Dim colID As Long, colVrsta As Long, colSorta As Long
    Dim i As Long
    
    On Error GoTo EH
    
    data = GetTableData(TBL_KULTURE)
    If IsEmpty(data) Then
        ExportKulture = WriteHeaderOnly(sheetID, "Kulture", _
            "KulturaID", "VrstaVoca", "SortaVoca")
        Exit Function
    End If
    
    colID = GetColumnIndex(TBL_KULTURE, "KulturaID")
    colVrsta = GetColumnIndex(TBL_KULTURE, "VrstaVoca")
    colSorta = GetColumnIndex(TBL_KULTURE, "SortaVoca")
    
    ReDim result(1 To UBound(data, 1) + 1, 1 To 3)
    
    ' Header
    result(1, 1) = "KulturaID"
    result(1, 2) = "VrstaVoca"
    result(1, 3) = "SortaVoca"
    
    For i = 1 To UBound(data, 1)
        result(i + 1, 1) = CStr(data(i, colID))
        result(i + 1, 2) = CStr(data(i, colVrsta))
        result(i + 1, 3) = CStr(data(i, colSorta))
    Next i
    
    ExportKulture = WriteSheetData(sheetID, "Kulture", result)
    Exit Function

EH:
    LogErr "ExportKulture"
    ExportKulture = False
End Function

Private Function ExportParcele(ByVal sheetID As String) As Boolean
    Dim data As Variant
    Dim result() As Variant
    
    Dim colID As Long, colKoop As Long, colKatBroj As Long
    Dim colKatOpstina As Long, colKultura As Long, colPovrsina As Long
    Dim colGGAP As Long, colAktivna As Long, colGeoStatus As Long
    Dim colGeoSource As Long, colN As Long, colE As Long
    Dim colLat As Long, colLng As Long, colPolygon As Long
    Dim colMeteo As Long, colRizik As Long
    Dim colDatumGeo As Long, colDatumAzur As Long, colNapomena As Long
    
    Dim i As Long
    
    On Error GoTo EH
    
    data = GetTableData(TBL_PARCELE)
    If IsEmpty(data) Then
        ExportParcele = WriteHeaderOnly(sheetID, "Parcele", _
            COL_PAR_ID, COL_PAR_KOOP, COL_PAR_KAT_BROJ, COL_PAR_KAT_OPSTINA, _
            COL_PAR_KULTURA, COL_PAR_POVRSINA, COL_PAR_GGAP, COL_PAR_AKTIVNA, _
            COL_PAR_GEO_STATUS, COL_PAR_GEO_SOURCE, COL_PAR_N, COL_PAR_E, _
            COL_PAR_LAT, COL_PAR_LNG, COL_PAR_POLYGON, COL_PAR_METEO, _
            COL_PAR_RIZIK, COL_PAR_DATUM_GEO, COL_PAR_DATUM_AZUR, COL_PAR_NAPOMENA)
        Exit Function
    End If
    
    data = ExcludeStornirano(data, TBL_PARCELE)
    If IsEmpty(data) Then
        ExportParcele = WriteHeaderOnly(sheetID, "Parcele", _
            COL_PAR_ID, COL_PAR_KOOP, COL_PAR_KAT_BROJ, COL_PAR_KAT_OPSTINA, _
            COL_PAR_KULTURA, COL_PAR_POVRSINA, COL_PAR_GGAP, COL_PAR_AKTIVNA, _
            COL_PAR_GEO_STATUS, COL_PAR_GEO_SOURCE, COL_PAR_N, COL_PAR_E, _
            COL_PAR_LAT, COL_PAR_LNG, COL_PAR_POLYGON, COL_PAR_METEO, _
            COL_PAR_RIZIK, COL_PAR_DATUM_GEO, COL_PAR_DATUM_AZUR, COL_PAR_NAPOMENA)
        Exit Function
    End If
    
    colID = GetColumnIndex(TBL_PARCELE, COL_PAR_ID)
    colKoop = GetColumnIndex(TBL_PARCELE, COL_PAR_KOOP)
    colKatBroj = GetColumnIndex(TBL_PARCELE, COL_PAR_KAT_BROJ)
    colKatOpstina = GetColumnIndex(TBL_PARCELE, COL_PAR_KAT_OPSTINA)
    colKultura = GetColumnIndex(TBL_PARCELE, COL_PAR_KULTURA)
    colPovrsina = GetColumnIndex(TBL_PARCELE, COL_PAR_POVRSINA)
    colGGAP = GetColumnIndex(TBL_PARCELE, COL_PAR_GGAP)
    colAktivna = GetColumnIndex(TBL_PARCELE, COL_PAR_AKTIVNA)
    colGeoStatus = GetColumnIndex(TBL_PARCELE, COL_PAR_GEO_STATUS)
    colGeoSource = GetColumnIndex(TBL_PARCELE, COL_PAR_GEO_SOURCE)
    colN = GetColumnIndex(TBL_PARCELE, COL_PAR_N)
    colE = GetColumnIndex(TBL_PARCELE, COL_PAR_E)
    colLat = GetColumnIndex(TBL_PARCELE, COL_PAR_LAT)
    colLng = GetColumnIndex(TBL_PARCELE, COL_PAR_LNG)
    colPolygon = GetColumnIndex(TBL_PARCELE, COL_PAR_POLYGON)
    colMeteo = GetColumnIndex(TBL_PARCELE, COL_PAR_METEO)
    colRizik = GetColumnIndex(TBL_PARCELE, COL_PAR_RIZIK)
    colDatumGeo = GetColumnIndex(TBL_PARCELE, COL_PAR_DATUM_GEO)
    colDatumAzur = GetColumnIndex(TBL_PARCELE, COL_PAR_DATUM_AZUR)
    colNapomena = GetColumnIndex(TBL_PARCELE, COL_PAR_NAPOMENA)
    
    Dim activeCount As Long
    For i = 1 To UBound(data, 1)
        If IsPWAActive(data(i, colAktivna)) Then activeCount = activeCount + 1
    Next i

    If activeCount = 0 Then
        ExportParcele = WriteHeaderOnly(sheetID, "Parcele", _
            COL_PAR_ID, COL_PAR_KOOP, COL_PAR_KAT_BROJ, COL_PAR_KAT_OPSTINA, _
            COL_PAR_KULTURA, COL_PAR_POVRSINA, COL_PAR_GGAP, COL_PAR_AKTIVNA, _
            COL_PAR_GEO_STATUS, COL_PAR_GEO_SOURCE, COL_PAR_N, COL_PAR_E, _
            COL_PAR_LAT, COL_PAR_LNG, COL_PAR_POLYGON, COL_PAR_METEO, _
            COL_PAR_RIZIK, COL_PAR_DATUM_GEO, COL_PAR_DATUM_AZUR, COL_PAR_NAPOMENA)
        Exit Function
    End If

    ReDim result(1 To activeCount + 1, 1 To 20)
    
    ' Header
    result(1, 1) = COL_PAR_ID
    result(1, 2) = COL_PAR_KOOP
    result(1, 3) = COL_PAR_KAT_BROJ
    result(1, 4) = COL_PAR_KAT_OPSTINA
    result(1, 5) = COL_PAR_KULTURA
    result(1, 6) = COL_PAR_POVRSINA
    result(1, 7) = COL_PAR_GGAP
    result(1, 8) = COL_PAR_AKTIVNA
    result(1, 9) = COL_PAR_GEO_STATUS
    result(1, 10) = COL_PAR_GEO_SOURCE
    result(1, 11) = COL_PAR_N
    result(1, 12) = COL_PAR_E
    result(1, 13) = COL_PAR_LAT
    result(1, 14) = COL_PAR_LNG
    result(1, 15) = COL_PAR_POLYGON
    result(1, 16) = COL_PAR_METEO
    result(1, 17) = COL_PAR_RIZIK
    result(1, 18) = COL_PAR_DATUM_GEO
    result(1, 19) = COL_PAR_DATUM_AZUR
    result(1, 20) = COL_PAR_NAPOMENA
    
    Dim outRow As Long
    outRow = 1

    For i = 1 To UBound(data, 1)
        If IsPWAActive(data(i, colAktivna)) Then
            outRow = outRow + 1

            result(outRow, 1) = CStr(nz(data(i, colID), ""))
            result(outRow, 2) = CStr(nz(data(i, colKoop), ""))
            result(outRow, 3) = CStr(nz(data(i, colKatBroj), ""))
            result(outRow, 4) = CStr(nz(data(i, colKatOpstina), ""))
            result(outRow, 5) = CStr(nz(data(i, colKultura), ""))
            result(outRow, 6) = CStr(nz(data(i, colPovrsina), ""))
            result(outRow, 7) = CStr(nz(data(i, colGGAP), ""))
            result(outRow, 8) = CStr(nz(data(i, colAktivna), ""))
            result(outRow, 9) = CStr(nz(data(i, colGeoStatus), ""))
            result(outRow, 10) = CStr(nz(data(i, colGeoSource), ""))
            result(outRow, 11) = CStr(nz(data(i, colN), ""))
            result(outRow, 12) = CStr(nz(data(i, colE), ""))
            result(outRow, 13) = CStr(nz(data(i, colLat), ""))
            result(outRow, 14) = CStr(nz(data(i, colLng), ""))
            result(outRow, 15) = CStr(nz(data(i, colPolygon), ""))
            result(outRow, 16) = CStr(nz(data(i, colMeteo), ""))
            result(outRow, 17) = CStr(nz(data(i, colRizik), ""))
            result(outRow, 18) = CStr(nz(data(i, colDatumGeo), ""))
            result(outRow, 19) = CStr(nz(data(i, colDatumAzur), ""))
            result(outRow, 20) = CStr(nz(data(i, colNapomena), ""))
        End If
    Next i
    
    ExportParcele = WriteSheetData(sheetID, "Parcele", result)
    Exit Function

EH:
    LogErr "ExportParcele"
    ExportParcele = False
End Function

Public Function SyncParceleToGoogle_Core(ByVal showMessages As Boolean) As Boolean
    Const SRC As String = "modStammdatenSync.SyncParceleToGoogle_Core"

    On Error GoTo EH

    Dim folderID As String
    Dim sheetID As String

    SyncParceleToGoogle_Core = False

    If Not IsGoogleAuthConfigured() Then
        LogError SRC, "Google OAuth2 nije konfigurisan."
        If showMessages Then MsgBox "Google OAuth2 nije konfigurisan.", vbCritical, APP_NAME
        Exit Function
    End If

    folderID = GetConfigValue("GOOGLE_PWA_FOLDER_ID")
    If Len(Trim$(folderID)) = 0 Then
        LogError SRC, "GOOGLE_PWA_FOLDER_ID nije postavljen."
        If showMessages Then MsgBox "GOOGLE_PWA_FOLDER_ID nije postavljen.", vbCritical, APP_NAME
        Exit Function
    End If

    sheetID = GetConfigValue("GOOGLE_STAMMDATEN_SHEET_ID")

    If Len(Trim$(sheetID)) = 0 Then
        sheetID = GetSpreadsheetID("Stammdaten", folderID)
    End If

    If Len(Trim$(sheetID)) = 0 Then
        sheetID = CreateSpreadsheet("Stammdaten", folderID)
    End If

    If Len(Trim$(sheetID)) = 0 Then
        LogError SRC, "Stammdaten sheet nije mogao biti pronadjen ili kreiran."
        If showMessages Then MsgBox "Stammdaten Google Sheet nije mogao biti pronadjen ili kreiran.", vbCritical, APP_NAME
        Exit Function
    End If

    Call SetConfigValue("GOOGLE_STAMMDATEN_SHEET_ID", sheetID)

    If Not AddSheetTab(sheetID, "Parcele") Then
        LogError SRC, "Parcele tab nije dostupan u Stammdaten sheet-u."
        If showMessages Then MsgBox "Parcele tab nije dostupan u Google Stammdaten sheet-u.", vbCritical, APP_NAME
        Exit Function
    End If

    SyncParceleToGoogle_Core = ExportParcele(sheetID)

    If SyncParceleToGoogle_Core Then
        LogInfo SRC, "Parcele tab exportovan u Google."
    Else
        LogError SRC, "ExportParcele nije uspeo."
    End If

    If showMessages Then
        If SyncParceleToGoogle_Core Then
            MsgBox "Parcele su sinhronizovane u Google.", vbInformation, APP_NAME
        Else
            MsgBox "Export parcela nije uspeo. Proveri log.", vbExclamation, APP_NAME
        End If
    End If

    Exit Function

EH:
    LogErr SRC

    If showMessages Then
        MsgBox Poruka("STM_ERR_GRESKA_PRI_EXPORTU") & Err.description, vbCritical, APP_NAME
    End If

    SyncParceleToGoogle_Core = False
End Function

Private Function ExportStanice(ByVal sheetID As String) As Boolean
    Dim data As Variant
    Dim result() As Variant
    Dim colID As Long, colNaziv As Long, colMesto As Long, colAktivan As Long
    Dim i As Long, outRow As Long
    
    On Error GoTo EH
    
    data = GetTableData(TBL_STANICE)
    If IsEmpty(data) Then
        ExportStanice = WriteHeaderOnly(sheetID, "Stanice", _
            "StanicaID", "Naziv", "Mesto")
        Exit Function
    End If
    data = ExcludeStornirano(data, TBL_STANICE)

    If IsEmpty(data) Then
        ExportStanice = WriteHeaderOnly(sheetID, "Stanice", _
            "StanicaID", "Naziv", "Mesto")
        Exit Function
    End If
    
    colID = GetColumnIndex(TBL_STANICE, "StanicaID")
    colNaziv = GetColumnIndex(TBL_STANICE, "Naziv")
    colMesto = GetColumnIndex(TBL_STANICE, "Mesto")
    colAktivan = GetColumnIndex(TBL_STANICE, "Aktivan")
    
    ' Erst zaehlen wieviele aktiv
    Dim cnt As Long: cnt = 0
    For i = 1 To UBound(data, 1)
        If IsPWAActive(data(i, colAktivan)) Then cnt = cnt + 1
    Next i
    
    If cnt = 0 Then
        ExportStanice = WriteHeaderOnly(sheetID, "Stanice", _
            "StanicaID", "Naziv", "Mesto")
        Exit Function
    End If
    
    ReDim result(1 To cnt + 1, 1 To 3)
    
    ' Header
    result(1, 1) = "StanicaID"
    result(1, 2) = "Naziv"
    result(1, 3) = "Mesto"
    
    outRow = 2
    For i = 1 To UBound(data, 1)
        If IsPWAActive(data(i, colAktivan)) Then
            result(outRow, 1) = CStr(data(i, colID))
            result(outRow, 2) = CStr(nz(data(i, colNaziv), ""))
            result(outRow, 3) = CStr(nz(data(i, colMesto), ""))
            outRow = outRow + 1
        End If
    Next i
    
    ExportStanice = WriteSheetData(sheetID, "Stanice", result)
    Exit Function
EH:
    LogErr "ExportStanice"
    ExportStanice = False
End Function

Private Function ExportKupci(ByVal sheetID As String) As Boolean
    Dim data As Variant
    Dim result() As Variant
    Dim colID As Long, colNaziv As Long, colMesto As Long, colAktivan As Long
    Dim i As Long, outRow As Long
    
    On Error GoTo EH
    
    data = GetTableData(TBL_KUPCI)
    If IsEmpty(data) Then
        ExportKupci = WriteHeaderOnly(sheetID, "Kupci", _
            "KupacID", "Naziv", "Mesto")
        Exit Function
    End If
    data = ExcludeStornirano(data, TBL_KUPCI)
    If IsEmpty(data) Then
        ExportKupci = WriteHeaderOnly(sheetID, "Kupci", _
            "KupacID", "Naziv", "Mesto")
        Exit Function
    End If
    
    colID = GetColumnIndex(TBL_KUPCI, "KupacID")
    colNaziv = GetColumnIndex(TBL_KUPCI, "Naziv")
    colMesto = GetColumnIndex(TBL_KUPCI, "Mesto")
    colAktivan = GetColumnIndex(TBL_KUPCI, "Aktivan")
    
    ' Erst zaehlen wieviele aktiv
    Dim cnt As Long: cnt = 0
    For i = 1 To UBound(data, 1)
        If IsPWAActive(data(i, colAktivan)) Then cnt = cnt + 1
    Next i
    
    If cnt = 0 Then
        ExportKupci = WriteHeaderOnly(sheetID, "Kupci", _
            "KupacID", "Naziv", "Mesto")
        Exit Function
    End If
    
    ReDim result(1 To cnt + 1, 1 To 3)
    
    ' Header
    result(1, 1) = "KupacID"
    result(1, 2) = "Naziv"
    result(1, 3) = "Mesto"
    
    outRow = 2
    For i = 1 To UBound(data, 1)
        If IsPWAActive(data(i, colAktivan)) Then
            result(outRow, 1) = CStr(data(i, colID))
            result(outRow, 2) = CStr(nz(data(i, colNaziv), ""))
            result(outRow, 3) = CStr(nz(data(i, colMesto), ""))
            outRow = outRow + 1
        End If
    Next i
    
    ExportKupci = WriteSheetData(sheetID, "Kupci", result)
    Exit Function
EH:
    LogErr "ExportKupci"
    ExportKupci = False
End Function

Private Function ExportVozaci(ByVal sheetID As String) As Boolean
    Dim data As Variant
    Dim result() As Variant
    Dim colID As Long, colIme As Long, colPrezime As Long, colTelefon As Long, colKapacitetKG As Long, colAktivan As Long
    Dim i As Long, outRow As Long
    
    On Error GoTo EH
    
    data = GetTableData(TBL_VOZACI)
    If IsEmpty(data) Then
        ExportVozaci = WriteHeaderOnly(sheetID, "Vozaci", _
            "VozacID", "Ime", "Prezime", "Telefon", "KapacitetKG")
        Exit Function
    End If
    data = ExcludeStornirano(data, TBL_VOZACI)
    If IsEmpty(data) Then
        ExportVozaci = WriteHeaderOnly(sheetID, "Vozaci", _
            "VozacID", "Ime", "Prezime", "Telefon", "KapacitetKG")
        Exit Function
    End If
    
    colID = GetColumnIndex(TBL_VOZACI, "VozacID")
    colIme = GetColumnIndex(TBL_VOZACI, "Ime")
    colPrezime = GetColumnIndex(TBL_VOZACI, "Prezime")
    colTelefon = GetColumnIndex(TBL_VOZACI, "Telefon")
    colKapacitetKG = GetColumnIndex(TBL_VOZACI, "KapacitetKG")
    colAktivan = GetColumnIndex(TBL_VOZACI, "Aktivan")
    
    ' Erst zaehlen wieviele aktiv
    Dim cnt As Long: cnt = 0
    For i = 1 To UBound(data, 1)
        If IsPWAActive(data(i, colAktivan)) Then cnt = cnt + 1
    Next i
    
    If cnt = 0 Then
        ExportVozaci = WriteHeaderOnly(sheetID, "Vozaci", _
            "VozacID", "Ime", "Prezime", "Telefon", "KapacitetKG")
        Exit Function
    End If
    
    ReDim result(1 To cnt + 1, 1 To 5)
    
    ' Header
    result(1, 1) = "VozacID"
    result(1, 2) = "Ime"
    result(1, 3) = "Prezime"
    result(1, 4) = "Telefon"
    result(1, 5) = "KapacitetKG"
    
    outRow = 2
    For i = 1 To UBound(data, 1)
        If IsPWAActive(data(i, colAktivan)) Then
            result(outRow, 1) = CStr(data(i, colID))
            result(outRow, 2) = CStr(nz(data(i, colIme), ""))
            result(outRow, 3) = CStr(nz(data(i, colPrezime), ""))
            result(outRow, 4) = CStr(nz(data(i, colTelefon), ""))
            result(outRow, 5) = CStr(nz(data(i, colKapacitetKG), ""))
            outRow = outRow + 1
        End If
    Next i
    
    ExportVozaci = WriteSheetData(sheetID, "Vozaci", result)
    Exit Function
EH:
    LogErr "ExporVozaci"
    ExportVozaci = False
End Function

Private Function ExportArtikli(ByVal sheetID As String) As Boolean
    Dim data As Variant
    Dim result() As Variant
    Dim i As Long, outRow As Long
    
    On Error GoTo EH
    
    data = GetTableData(TBL_ARTIKLI)
    If IsEmpty(data) Then
        ExportArtikli = WriteHeaderOnly(sheetID, "Artikli", _
            "ArtikalID", "Naziv", "Tip", "JedinicaMere", "CenaPoJedinici", _
            "DozaPoHa", "Kultura", "Pakovanje", "BarKod", "Karenca", "Aktivan")
        Exit Function
    End If
    data = ExcludeStornirano(data, TBL_ARTIKLI)
    If IsEmpty(data) Then
        ExportArtikli = WriteHeaderOnly(sheetID, "Artikli", _
            "ArtikalID", "Naziv", "Tip", "JedinicaMere", "CenaPoJedinici", _
            "DozaPoHa", "Kultura", "Pakovanje", "BarKod", "Karenca", "Aktivan")
        Exit Function
    End If
    
    Dim colArtID As Long: colArtID = GetColumnIndex(TBL_ARTIKLI, "ArtikalID")
    Dim colNaziv As Long: colNaziv = GetColumnIndex(TBL_ARTIKLI, "Naziv")
    Dim colTip As Long: colTip = GetColumnIndex(TBL_ARTIKLI, "Tip")
    Dim colJM As Long: colJM = GetColumnIndex(TBL_ARTIKLI, "JedinicaMere")
    Dim colCena As Long: colCena = GetColumnIndex(TBL_ARTIKLI, "CenaPoJedinici")
    Dim colDoza As Long: colDoza = GetColumnIndex(TBL_ARTIKLI, "DozaPoHa")
    Dim colKultura As Long: colKultura = GetColumnIndex(TBL_ARTIKLI, "Kultura")
    Dim colPak As Long: colPak = GetColumnIndex(TBL_ARTIKLI, "Pakovanje")
    Dim colBarKod As Long: colBarKod = GetColumnIndex(TBL_ARTIKLI, "BarKod")
    Dim colKarenca As Long: colKarenca = GetColumnIndex(TBL_ARTIKLI, "KarencaDana")
    Dim colAktivan As Long: colAktivan = GetColumnIndex(TBL_ARTIKLI, "Aktivan")
    
    ' Erst zaehlen wieviele aktiv
    Dim cnt As Long: cnt = 0
    For i = 1 To UBound(data, 1)
        If IsPWAActive(data(i, colAktivan)) Then cnt = cnt + 1
    Next i
    
    If cnt = 0 Then
        ExportArtikli = WriteHeaderOnly(sheetID, "Artikli", _
            "ArtikalID", "Naziv", "Tip", "JedinicaMere", "CenaPoJedinici", _
            "DozaPoHa", "Kultura", "Pakovanje", "BarKod", "Karenca", "Aktivan")
        Exit Function
    End If
    
    ReDim result(1 To cnt + 1, 1 To 11)
    
    ' Header
    result(1, 1) = "ArtikalID"
    result(1, 2) = "Naziv"
    result(1, 3) = "Tip"
    result(1, 4) = "JedinicaMere"
    result(1, 5) = "CenaPoJedinici"
    result(1, 6) = "DozaPoHa"
    result(1, 7) = "Kultura"
    result(1, 8) = "Pakovanje"
    result(1, 9) = "BarKod"
    result(1, 10) = "Karenca"
    result(1, 11) = "Aktivan"
    
    outRow = 2
    For i = 1 To UBound(data, 1)
        If IsPWAActive(data(i, colAktivan)) Then
            result(outRow, 1) = CStr(data(i, colArtID))
            result(outRow, 2) = CStr(nz(data(i, colNaziv), ""))
            result(outRow, 3) = CStr(nz(data(i, colTip), ""))
            result(outRow, 4) = CStr(nz(data(i, colJM), ""))
            result(outRow, 5) = CStr(nz(data(i, colCena), ""))
            result(outRow, 6) = CStr(nz(data(i, colDoza), ""))
            result(outRow, 7) = CStr(nz(data(i, colKultura), ""))
            result(outRow, 8) = CStr(nz(data(i, colPak), ""))
            result(outRow, 9) = CStr(nz(data(i, colBarKod), ""))
            result(outRow, 10) = CStr(nz(data(i, colKarenca), ""))
            result(outRow, 11) = CStr(nz(data(i, colAktivan), ""))
            
            outRow = outRow + 1
        End If
    Next i
    
    ExportArtikli = WriteSheetData(sheetID, "Artikli", result)
    Exit Function
EH:
    LogErr "ExportArtikli"
    ExportArtikli = False
End Function

Private Function ExportMagacinKoop(ByVal sheetID As String) As Boolean
    Dim magData As Variant
    Dim artData As Variant
    Dim result() As Variant
    Dim dict As Object
    Dim artDict As Object
    Dim keys As Variant
    Dim vals As Variant
    Dim meta As Variant
    Dim parts() As String
    Dim i As Long, outRow As Long
    Dim cnt As Long
    Dim kolicina As Double
    
    On Error GoTo EH
    
    magData = GetTableData(TBL_MAGACIN)
    If IsEmpty(magData) Then
        ExportMagacinKoop = WriteHeaderOnly(sheetID, "MagacinKoop", _
                "KooperantID", "ArtikalID", "ArtikalNaziv", "Tip", "JedinicaMere", _
                "CenaPoJedinici", "DozaPoHa", "Pakovanje", "Karenca", _
                "Primljeno", "Utroseno", "Stanje")
        Exit Function
    End If
    magData = ExcludeStornirano(magData, TBL_MAGACIN)
    
    If IsEmpty(magData) Then
        ExportMagacinKoop = WriteHeaderOnly(sheetID, "MagacinKoop", _
            "KooperantID", "ArtikalID", "ArtikalNaziv", "Tip", "JedinicaMere", _
            "CenaPoJedinici", "DozaPoHa", "Pakovanje", "Karenca", _
            "Primljeno", "Utroseno", "Stanje")
        Exit Function
    End If
    
    Dim colMKoop As Long: colMKoop = GetColumnIndex(TBL_MAGACIN, "KooperantID")
    Dim colMArt As Long: colMArt = GetColumnIndex(TBL_MAGACIN, "ArtikalID")
    Dim colMTip As Long: colMTip = GetColumnIndex(TBL_MAGACIN, "Tip")
    Dim colMKol As Long: colMKol = GetColumnIndex(TBL_MAGACIN, "Kolicina")
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    For i = 1 To UBound(magData, 1)
        If CStr(nz(magData(i, colMTip), "")) = "Izlaz" Then
            Dim koopID As String
            Dim artID As String
            Dim key As String
            
            koopID = CStr(nz(magData(i, colMKoop), ""))
            artID = CStr(nz(magData(i, colMArt), ""))
            
            If koopID <> "" And artID <> "" Then
                key = koopID & "|" & artID
                
                If Not dict.Exists(key) Then dict.Add key, 0#
                dict(key) = CDbl(dict(key)) + CDbl(nz(magData(i, colMKol), 0))
            End If
        End If
    Next i
    
    If dict.count = 0 Then
        ExportMagacinKoop = WriteHeaderOnly(sheetID, "MagacinKoop", _
            "KooperantID", "ArtikalID", "ArtikalNaziv", "Tip", "JedinicaMere", _
            "CenaPoJedinici", "DozaPoHa", "Pakovanje", "Karenca", _
            "Primljeno", "Utroseno", "Stanje")
        Exit Function
    End If
    
    artData = GetTableData(TBL_ARTIKLI)
    Set artDict = CreateObject("Scripting.Dictionary")
    
    If Not IsEmpty(artData) Then
        Dim colArtID As Long: colArtID = GetColumnIndex(TBL_ARTIKLI, "ArtikalID")
        Dim colNaziv As Long: colNaziv = GetColumnIndex(TBL_ARTIKLI, "Naziv")
        Dim colTip As Long: colTip = GetColumnIndex(TBL_ARTIKLI, "Tip")
        Dim colJM As Long: colJM = GetColumnIndex(TBL_ARTIKLI, "JedinicaMere")
        Dim colCena As Long: colCena = GetColumnIndex(TBL_ARTIKLI, "CenaPoJedinici")
        Dim colDoza As Long: colDoza = GetColumnIndex(TBL_ARTIKLI, "DozaPoHa")
        Dim colPak As Long: colPak = GetColumnIndex(TBL_ARTIKLI, "Pakovanje")
        Dim colKarenca As Long: colKarenca = GetColumnIndex(TBL_ARTIKLI, "KarencaDana")
        
        For i = 1 To UBound(artData, 1)
            artID = CStr(nz(artData(i, colArtID), ""))
            If artID <> "" Then
                If Not artDict.Exists(artID) Then
                    artDict.Add artID, Array( _
                        CStr(nz(artData(i, colNaziv), "")), _
                        CStr(nz(artData(i, colTip), "")), _
                        CStr(nz(artData(i, colJM), "")), _
                        CStr(nz(artData(i, colCena), "")), _
                        CStr(nz(artData(i, colDoza), "")), _
                        CStr(nz(artData(i, colPak), "")), _
                        CStr(nz(artData(i, colKarenca), "")))
                End If
            End If
        Next i
    End If
    
    cnt = dict.count
    If cnt = 0 Then
        ExportMagacinKoop = WriteHeaderOnly(sheetID, "MagacinKoop", _
            "KooperantID", "ArtikalID", "ArtikalNaziv", "Tip", "JedinicaMere", _
            "CenaPoJedinici", "DozaPoHa", "Pakovanje", "Karenca", _
            "Primljeno", "Utroseno", "Stanje")
        Exit Function
    End If
    
    ReDim result(1 To cnt + 1, 1 To 12)
    
    result(1, 1) = "KooperantID"
    result(1, 2) = "ArtikalID"
    result(1, 3) = "ArtikalNaziv"
    result(1, 4) = "Tip"
    result(1, 5) = "JedinicaMere"
    result(1, 6) = "CenaPoJedinici"
    result(1, 7) = "DozaPoHa"
    result(1, 8) = "Pakovanje"
    result(1, 9) = "Karenca"
    result(1, 10) = "Primljeno"
    result(1, 11) = "Utroseno"
    result(1, 12) = "Stanje"
    
    keys = dict.keys
    outRow = 2
    
    For i = 0 To dict.count - 1
        parts = Split(CStr(keys(i)), "|")
        kolicina = CDbl(dict(keys(i)))
        
        result(outRow, 1) = parts(0)
        result(outRow, 2) = parts(1)
        
        If artDict.Exists(parts(1)) Then
            meta = artDict(parts(1))
            result(outRow, 3) = meta(0)
            result(outRow, 4) = meta(1)
            result(outRow, 5) = meta(2)
            result(outRow, 6) = meta(3)
            result(outRow, 7) = meta(4)
            result(outRow, 8) = meta(5)
            result(outRow, 9) = meta(6)
        Else
            result(outRow, 3) = ""
            result(outRow, 4) = ""
            result(outRow, 5) = ""
            result(outRow, 6) = ""
            result(outRow, 7) = ""
            result(outRow, 8) = ""
            result(outRow, 9) = ""
        End If
        
        result(outRow, 10) = kolicina
        result(outRow, 11) = 0
        result(outRow, 12) = kolicina
        
        outRow = outRow + 1
    Next i
    
    ExportMagacinKoop = WriteSheetData(sheetID, "MagacinKoop", result)
    Exit Function
EH:
    LogErr "ExportMagacinKoop"
    ExportMagacinKoop = False
End Function

Private Function ExportConfig(ByVal sheetID As String) As Boolean
    ' Exportiert PWA-relevante Config-Werte aus tblSEFConfig
    ' Filtert: alles was mit "Cena" beginnt + explizite PWA-Keys
    ' Credentials (GOOGLE_*, SEF_API_KEY etc.) werden NICHT exportiert
    
    Dim result() As Variant
    Dim data As Variant
    Dim colKey As Long, colVal As Long
    Dim i As Long, outRow As Long
    Dim keyStr As String
    Dim include As Boolean
    
    On Error GoTo EH
    
    data = GetTableData("tblSEFConfig")
    If IsEmpty(data) Then
        ExportConfig = WriteHeaderOnly(sheetID, "Config", _
            "Parameter", "Vrednost")
        Exit Function
    End If
    
    colKey = GetColumnIndex("tblSEFConfig", "ConfigKey")
    colVal = GetColumnIndex("tblSEFConfig", "ConfigValue")
    
    ' Explizite PWA-Keys
    Dim pwaKeys As Variant
    pwaKeys = Array("OtkupAktivan", "RadnoVremeOd", "RadnoVremeDo", _
                    "SezonaOd", "SezonaDo", "TipAmbalaze", _
                    "DefaultVrsta", "DefaultSorta", "OtkupRokIsplate", "OtkupPDVStopa", _
                    "SELLER_NAME", "SELLER_PIB", "SELLER_MATICNI_BROJ", _
                    "SELLER_STREET", "SELLER_CITY", "SELLER_POSTAL_CODE", _
                    "SELLER_ACCOUNT")
    
    ' Zaehlen
    Dim matchCount As Long
    For i = 1 To UBound(data, 1)
        keyStr = CStr(data(i, colKey))
        If IsPwaConfigKey(keyStr, pwaKeys) Then matchCount = matchCount + 1
    Next i
    
    If matchCount = 0 Then
        ' Leeres Sheet mit Header schreiben
        ReDim result(1 To 1, 1 To 2)
        result(1, 1) = "Parameter"
        result(1, 2) = "Vrednost"
        ExportConfig = WriteSheetData(sheetID, "Config", result)
        Exit Function
    End If
    
    ReDim result(1 To matchCount + 1, 1 To 2)
    
    ' Header
    result(1, 1) = "Parameter"
    result(1, 2) = "Vrednost"
    
    outRow = 1
    For i = 1 To UBound(data, 1)
        keyStr = CStr(data(i, colKey))
        
        If IsPwaConfigKey(keyStr, pwaKeys) Then
            outRow = outRow + 1
            result(outRow, 1) = keyStr
            result(outRow, 2) = CStr(data(i, colVal))
        End If
    Next i
    
    ExportConfig = WriteSheetData(sheetID, "Config", result)
    Exit Function

EH:
    LogErr "ExportConfig"
    ExportConfig = False
End Function

Private Function IsPwaConfigKey(ByVal keyStr As String, ByVal pwaKeys As Variant) As Boolean
    ' Credentials ausschliessen
    If Left$(keyStr, 7) = "GOOGLE_" Then Exit Function
    If Left$(keyStr, 4) = "SEF_" Then Exit Function
    If Left$(keyStr, 5) = "SYNC_" Then Exit Function
    ' SELLER_* je DOZVOLJEN (za otkupni list)
    
    ' Cena-Keys
    If Left$(keyStr, 4) = "Cena" Then
        IsPwaConfigKey = True
        Exit Function
    End If
    
    ' Explizite PWA-Keys
    Dim k As Long
    For k = LBound(pwaKeys) To UBound(pwaKeys)
        If keyStr = CStr(pwaKeys(k)) Then
            IsPwaConfigKey = True
            Exit Function
        End If
    Next k
End Function

Private Function ExportUsers(ByVal sheetID As String) As Boolean
    Dim koopData As Variant, staData As Variant, vozData As Variant
    Dim result() As Variant
    Dim outRow As Long
    Dim totalRows As Long
    Dim i As Long
    
    On Error GoTo EH
    
    koopData = GetTableData(TBL_KOOPERANTI)
    If Not IsEmpty(koopData) Then koopData = ExcludeStornirano(koopData, TBL_KOOPERANTI)
    
    staData = GetTableData(TBL_STANICE)
    If Not IsEmpty(staData) Then staData = ExcludeStornirano(staData, TBL_STANICE)
    
    vozData = GetTableData(TBL_VOZACI)
    If Not IsEmpty(vozData) Then vozData = ExcludeStornirano(vozData, TBL_VOZACI)
    
    Dim koopCount As Long, staCount As Long, vozCount As Long
    If Not IsEmpty(koopData) Then koopCount = UBound(koopData, 1)
    If Not IsEmpty(staData) Then staCount = UBound(staData, 1)
    If Not IsEmpty(vozData) Then vozCount = UBound(vozData, 1)
    
    totalRows = 1 + koopCount + staCount + vozCount
    ReDim result(1 To totalRows, 1 To 5)
    
    result(1, 1) = "Username"
    result(1, 2) = "PIN"
    result(1, 3) = "Role"
    result(1, 4) = "EntityID"
    result(1, 5) = "DisplayName"
    
    outRow = 1
    
    ' --- Kooperanti ---
    If Not IsEmpty(koopData) Then
        Dim colKID As Long, colKIme As Long, colKPrezime As Long
        Dim colKAktivan As Long, colKPIN As Long
        
        colKID = GetColumnIndex(TBL_KOOPERANTI, "KooperantID")
        colKIme = GetColumnIndex(TBL_KOOPERANTI, "Ime")
        colKPrezime = GetColumnIndex(TBL_KOOPERANTI, "Prezime")
        colKAktivan = GetColumnIndex(TBL_KOOPERANTI, "Aktivan")
        colKPIN = GetColumnIndex(TBL_KOOPERANTI, "PIN")
        
        If colKPIN > 0 Then
            For i = 1 To UBound(koopData, 1)
                If IsPWAActive(koopData(i, colKAktivan)) Then
                    Dim kPin As String
                    kPin = Trim$(CStr(nz(koopData(i, colKPIN), "")))
                    If Len(kPin) > 0 Then
                        outRow = outRow + 1
                        Dim kIme As String, kPrezime As String
                        kIme = Trim$(CStr(koopData(i, colKIme)))
                        kPrezime = Trim$(CStr(koopData(i, colKPrezime)))
                        result(outRow, 1) = LCase$(Left$(kIme, 1) & kPrezime)
                        result(outRow, 2) = kPin
                        result(outRow, 3) = "Kooperant"
                        result(outRow, 4) = CStr(koopData(i, colKID))
                        result(outRow, 5) = kIme & " " & kPrezime
                    End If
                End If
            Next i
        End If
    End If
    
    ' --- Stanice (Otkupci) ---
    If Not IsEmpty(staData) Then
        Dim colSID As Long, colSNaziv As Long, colSAktivan As Long
        Dim colSPIN As Long, colSIme As Long, colSPrezime As Long
        
        colSID = GetColumnIndex(TBL_STANICE, "StanicaID")
        colSNaziv = GetColumnIndex(TBL_STANICE, "Naziv")
        colSIme = GetColumnIndex(TBL_STANICE, "Ime")
        colSPrezime = GetColumnIndex(TBL_STANICE, "Prezime")
        colSAktivan = GetColumnIndex(TBL_STANICE, "Aktivan")
        colSPIN = GetColumnIndex(TBL_STANICE, "PIN")
        
        If colSPIN > 0 And colSIme > 0 And colSPrezime > 0 Then
            For i = 1 To UBound(staData, 1)
                If IsPWAActive(staData(i, colSAktivan)) Then
                    Dim sPin As String
                    sPin = Trim$(CStr(nz(staData(i, colSPIN), "")))
                    If Len(sPin) > 0 Then
                        outRow = outRow + 1
                        Dim sIme As String, sPrezime As String
                        sIme = Trim$(CStr(staData(i, colSIme)))
                        sPrezime = Trim$(CStr(staData(i, colSPrezime)))
                        result(outRow, 1) = LCase$(Left$(sIme, 1) & sPrezime)
                        result(outRow, 2) = sPin
                        result(outRow, 3) = "Otkupac"
                        result(outRow, 4) = CStr(staData(i, colSID))
                        result(outRow, 5) = sIme & " " & sPrezime & " - " & CStr(staData(i, colSNaziv))
                    End If
                End If
            Next i
        End If
    End If
    
    ' --- Vozaci ---
    If Not IsEmpty(vozData) Then
        Dim colVID As Long, colVIme As Long, colVPrezime As Long
        Dim colVAktivan As Long, colVPIN As Long
        
        colVID = GetColumnIndex(TBL_VOZACI, "VozacID")
        colVIme = GetColumnIndex(TBL_VOZACI, "Ime")
        colVPrezime = GetColumnIndex(TBL_VOZACI, "Prezime")
        colVAktivan = GetColumnIndex(TBL_VOZACI, "Aktivan")
        colVPIN = GetColumnIndex(TBL_VOZACI, "PIN")
        
        If colVPIN > 0 Then
            For i = 1 To UBound(vozData, 1)
                If IsPWAActive(vozData(i, colVAktivan)) Then
                    Dim vPin As String
                    vPin = Trim$(CStr(nz(vozData(i, colVPIN), "")))
                    If Len(vPin) > 0 Then
                        outRow = outRow + 1
                        Dim vIme As String, vPrezime As String
                        vIme = Trim$(CStr(vozData(i, colVIme)))
                        vPrezime = Trim$(CStr(vozData(i, colVPrezime)))
                        result(outRow, 1) = LCase$(Left$(vIme, 1) & vPrezime)
                        result(outRow, 2) = vPin
                        result(outRow, 3) = "Vozac"
                        result(outRow, 4) = CStr(vozData(i, colVID))
                        result(outRow, 5) = vIme & " " & vPrezime
                    End If
                End If
            Next i
        End If
    End If
    
    ' --- Management ---
    Dim cfgData As Variant
    cfgData = GetTableData(TBL_SEF_CONFIG)
    
    If Not IsEmpty(cfgData) Then
        Dim colCfgKey As Long, colCfgVal As Long
        colCfgKey = GetColumnIndex(TBL_SEF_CONFIG, "ConfigKey")
        colCfgVal = GetColumnIndex(TBL_SEF_CONFIG, "ConfigValue")
        
        ' Suche MGMT_USER_1, MGMT_USER_2, etc.
        ' Format: "Username|PIN|EntityID|DisplayName"
        For i = 1 To UBound(cfgData, 1)
            Dim cfgKey As String
            cfgKey = CStr(cfgData(i, colCfgKey))
            If Left$(cfgKey, 9) = "MGMT_USER" Then
                Dim parts() As String
                parts = Split(CStr(cfgData(i, colCfgVal)), "|")
                If UBound(parts) >= 3 Then
                    outRow = outRow + 1
                    If outRow > UBound(result, 1) Then
                        ' Expand array
                        Dim tmp() As Variant
                        ReDim tmp(1 To outRow + 5, 1 To 5)
                        Dim ri As Long, cI As Long
                        For ri = 1 To outRow - 1
                            For cI = 1 To 5
                                tmp(ri, cI) = result(ri, cI)
                            Next cI
                        Next ri
                        result = tmp
                    End If
                    result(outRow, 1) = Trim$(parts(0))
                    result(outRow, 2) = Trim$(parts(1))
                    result(outRow, 3) = "Management"
                    result(outRow, 4) = Trim$(parts(2))
                    result(outRow, 5) = Trim$(parts(3))
                End If
            End If
        Next i
    End If
    
    ' Auf tatsaechliche Groesse kuerzen
    If outRow < UBound(result, 1) Then
        Dim finalRows() As Variant
        Dim r As Long, c As Long
        ReDim finalRows(1 To outRow, 1 To 5)
        For r = 1 To outRow
            For c = 1 To 5
                finalRows(r, c) = result(r, c)
            Next c
        Next r
        ExportUsers = WriteSheetData(sheetID, "Users", finalRows)
    Else
        ExportUsers = WriteSheetData(sheetID, "Users", result)
    End If
    
    LogInfo "ExportUsers", "Exportiert: " & (outRow - 1) & " Users"
    Exit Function

EH:
    LogErr "ExportUsers"
    ExportUsers = False
End Function

Private Function ExportFakture(ByVal sheetID As String) As Boolean
    Dim data As Variant
    Dim colID As Long, colBroj As Long, colDatum As Long, colKupac As Long
    Dim colIznos As Long, colStatus As Long, colSEFStatus As Long
    Dim i As Long
    
    On Error GoTo EH
    
    data = GetTableData(TBL_FAKTURE)
    If Not IsEmpty(data) Then data = ExcludeStornirano(data, TBL_FAKTURE)
    
    If IsEmpty(data) Then
        ExportFakture = WriteHeaderOnly(sheetID, "Fakture", _
            "FakturaID", "BrojFakture", "Datum", "KupacID", "Kupac", _
            "Iznos", "Placeno", "Saldo", "Status", "SEFStatus")
        Exit Function
    End If
    
    colID = GetColumnIndex(TBL_FAKTURE, "FakturaID")
    colBroj = GetColumnIndex(TBL_FAKTURE, "BrojFakture")
    colDatum = GetColumnIndex(TBL_FAKTURE, "Datum")
    colKupac = GetColumnIndex(TBL_FAKTURE, "KupacID")
    colIznos = GetColumnIndex(TBL_FAKTURE, "Iznos")
    colStatus = GetColumnIndex(TBL_FAKTURE, "Status")
    colSEFStatus = GetColumnIndex(TBL_FAKTURE, "SEFStatus")
    
    ' Uplate per Faktura aus tblNovac
    Dim novData As Variant
    novData = GetTableData(TBL_NOVAC)
    If Not IsEmpty(novData) Then novData = ExcludeStornirano(novData, TBL_NOVAC)
    
    Dim dictPlaceno As Object
    Set dictPlaceno = CreateObject("Scripting.Dictionary")
    
    If Not IsEmpty(novData) Then
        Dim colNovFaktura As Long, colNovUplata As Long, colNovTip As Long
        colNovFaktura = GetColumnIndex(TBL_NOVAC, COL_NOV_FAKTURA_ID)
        colNovUplata = GetColumnIndex(TBL_NOVAC, COL_NOV_UPLATA)
        colNovTip = GetColumnIndex(TBL_NOVAC, COL_NOV_TIP)
        
        Dim n As Long
        For n = 1 To UBound(novData, 1)
            Dim fakID As String
            fakID = Trim$(CStr(nz(novData(n, colNovFaktura), "")))
            If Len(fakID) > 0 Then
                Dim tip As String
                tip = CStr(novData(n, colNovTip))
                If tip = NOV_KUPCI_UPLATA Then
                    If Not dictPlaceno.Exists(fakID) Then dictPlaceno.Add fakID, 0#
                    dictPlaceno(fakID) = dictPlaceno(fakID) + CDbl(nz(novData(n, colNovUplata), 0))
                End If
            End If
        Next n
    End If
    
    Dim result() As Variant
    ReDim result(1 To UBound(data, 1) + 1, 1 To 10)
    
    result(1, 1) = "FakturaID"
    result(1, 2) = "BrojFakture"
    result(1, 3) = "Datum"
    result(1, 4) = "KupacID"
    result(1, 5) = "Kupac"
    result(1, 6) = "Iznos"
    result(1, 7) = "Placeno"
    result(1, 8) = "Saldo"
    result(1, 9) = "Status"
    result(1, 10) = "SEFStatus"
    
    For i = 1 To UBound(data, 1)
        Dim kupacNaziv As Variant
        kupacNaziv = LookupValue(TBL_KUPCI, "KupacID", CStr(data(i, colKupac)), "Naziv")
        Dim iznos As Double
        iznos = CDbl(nz(data(i, colIznos), 0))
        Dim fID As String
        fID = CStr(data(i, colID))
        Dim placeno As Double
        placeno = 0
        If dictPlaceno.Exists(fID) Then placeno = dictPlaceno(fID)
        
        result(i + 1, 1) = fID
        result(i + 1, 2) = CStr(data(i, colBroj))
        result(i + 1, 3) = CStr(data(i, colDatum))
        result(i + 1, 4) = CStr(data(i, colKupac))
        result(i + 1, 5) = CStr(nz(kupacNaziv, data(i, colKupac)))
        result(i + 1, 6) = CStr(iznos)
        result(i + 1, 7) = CStr(placeno)
        result(i + 1, 8) = CStr(iznos - placeno)
        result(i + 1, 9) = CStr(data(i, colStatus))
        result(i + 1, 10) = CStr(nz(data(i, colSEFStatus), ""))
    Next i
    
    ExportFakture = WriteSheetData(sheetID, "Fakture", result)
    Exit Function
EH:
    LogErr "ExportFakture"
    ExportFakture = False
End Function

Private Function ExportFakturaStavke(ByVal sheetID As String) As Boolean
    Dim data As Variant
    Dim colFakID As Long, colPrijID As Long, colBrojPrij As Long
    Dim colKlasa As Long, colKolicina As Long, colCena As Long
    Dim i As Long
    
    On Error GoTo EH
    
    data = GetTableData(TBL_FAKTURA_STAVKE)
    
    If IsEmpty(data) Then
        ExportFakturaStavke = WriteHeaderOnly(sheetID, "FakturaStavke", _
            "FakturaID", "PrijemnicaID", "BrojPrijemnice", "BrojZbirne", _
            "VrstaVoca", "Klasa", "Koli" & ChrW(269) & "ina", "Cena", "Iznos")
        Exit Function
    End If
    
    colFakID = GetColumnIndex(TBL_FAKTURA_STAVKE, "FakturaID")
    colPrijID = GetColumnIndex(TBL_FAKTURA_STAVKE, "PrijemnicaID")
    colBrojPrij = GetColumnIndex(TBL_FAKTURA_STAVKE, "BrojPrijemnice")
    colKlasa = GetColumnIndex(TBL_FAKTURA_STAVKE, "Klasa")
    colKolicina = GetColumnIndex(TBL_FAKTURA_STAVKE, "Kolicina")
    colCena = GetColumnIndex(TBL_FAKTURA_STAVKE, "Cena")
    
    ' BrojZbirne + VrstaVoca aus tblPrijemnica holen
    Dim prijData As Variant
    prijData = GetTableData(TBL_PRIJEMNICA)
    If Not IsEmpty(prijData) Then prijData = ExcludeStornirano(prijData, TBL_PRIJEMNICA)
    
    Dim dictZbirna As Object
    Set dictZbirna = CreateObject("Scripting.Dictionary")
    Dim dictVrsta As Object
    Set dictVrsta = CreateObject("Scripting.Dictionary")
    
    If Not IsEmpty(prijData) Then
        Dim colPrijPrijID As Long, colPrijZbirna As Long, colPrijVrsta As Long
        colPrijPrijID = GetColumnIndex(TBL_PRIJEMNICA, "PrijemnicaID")
        colPrijZbirna = GetColumnIndex(TBL_PRIJEMNICA, "BrojZbirne")
        colPrijVrsta = GetColumnIndex(TBL_PRIJEMNICA, "VrstaVoca")
        Dim p As Long
        For p = 1 To UBound(prijData, 1)
            Dim pid As String
            pid = CStr(prijData(p, colPrijPrijID))
            If Not dictZbirna.Exists(pid) Then
                dictZbirna.Add pid, CStr(nz(prijData(p, colPrijZbirna), ""))
            End If
            If Not dictVrsta.Exists(pid) Then
                dictVrsta.Add pid, CStr(nz(prijData(p, colPrijVrsta), ""))
            End If
        Next p
    End If
    
    Dim result() As Variant
    ReDim result(1 To UBound(data, 1) + 1, 1 To 9)
    
    result(1, 1) = "FakturaID"
    result(1, 2) = "PrijemnicaID"
    result(1, 3) = "BrojPrijemnice"
    result(1, 4) = "BrojZbirne"
    result(1, 5) = "VrstaVoca"
    result(1, 6) = "Klasa"
    result(1, 7) = "Koli" & ChrW(269) & "ina"
    result(1, 8) = "Cena"
    result(1, 9) = "Iznos"
    
    For i = 1 To UBound(data, 1)
        Dim prijemnicaID As String
        prijemnicaID = CStr(nz(data(i, colPrijID), ""))
        Dim kg As Double, cena As Double
        kg = CDbl(nz(data(i, colKolicina), 0))
        cena = CDbl(nz(data(i, colCena), 0))
        
        result(i + 1, 1) = CStr(data(i, colFakID))
        result(i + 1, 2) = prijemnicaID
        result(i + 1, 3) = CStr(nz(data(i, colBrojPrij), ""))
        result(i + 1, 4) = ""
        If dictZbirna.Exists(prijemnicaID) Then result(i + 1, 4) = dictZbirna(prijemnicaID)
        result(i + 1, 5) = ""
        If dictVrsta.Exists(prijemnicaID) Then result(i + 1, 5) = dictVrsta(prijemnicaID)
        result(i + 1, 6) = CStr(nz(data(i, colKlasa), ""))
        result(i + 1, 7) = CStr(kg)
        result(i + 1, 8) = CStr(cena)
        result(i + 1, 9) = CStr(kg * cena)
    Next i
    
    ExportFakturaStavke = WriteSheetData(sheetID, "FakturaStavke", result)
    Exit Function
EH:
    LogErr "ExportFakturaStavke"
    ExportFakturaStavke = False
End Function

Private Function WriteHeaderOnly(ByVal sheetID As String, _
                                 ByVal tabName As String, _
                                 ParamArray headers() As Variant) As Boolean
    On Error GoTo EH

    Dim result() As Variant
    Dim i As Long
    Dim n As Long

    n = UBound(headers) - LBound(headers) + 1
    If n <= 0 Then
        WriteHeaderOnly = False
        Exit Function
    End If

    ReDim result(1 To 1, 1 To n)

    For i = LBound(headers) To UBound(headers)
        result(1, i - LBound(headers) + 1) = CStr(headers(i))
    Next i

    WriteHeaderOnly = WriteSheetData(sheetID, tabName, result)
    Exit Function

EH:
    LogErr "WriteHeaderOnly", "Tab=" & tabName
    WriteHeaderOnly = False
End Function

Private Function IsPWAActive(ByVal value As Variant) As Boolean
    Dim s As String
    s = UCase$(Trim$(CStr(nz(value, ""))))

    Select Case s
        Case "NE", "NO", "FALSE", "0", "NEAKTIVAN", "INACTIVE"
            IsPWAActive = False
        Case Else
            IsPWAActive = True
    End Select
End Function
Private Sub Monitor_StammdatenSyncSuccess(ByVal successCount As Long, _
                                          ByVal totalTabs As Long, _
                                          ByVal sheetID As String)
    On Error Resume Next

    Monitor_Event _
        eventType:="STAMMDATEN_SYNC_SUCCESS", _
        severity:="INFO", _
        message:="Stammdaten sync completed. SuccessTabs=" & CStr(successCount) & "/" & CStr(totalTabs), _
        userId:="Operator", _
        moduleName:="modStammdatenSync", _
        procedureName:="SyncStammdatenToGoogle_Core", _
        entityType:="MasterData", _
        entityID:="Stammdaten", _
        correlationId:="STAMMDATEN-SYNC"
End Sub

Private Sub Monitor_StammdatenSyncFail(ByVal errNum As Long, _
                                       ByVal errDesc As String, _
                                       ByVal errSrc As String, _
                                       Optional ByVal successCount As Long = 0, _
                                       Optional ByVal totalTabs As Long = 13)
    On Error Resume Next

    Monitor_Error _
        moduleName:="modStammdatenSync", _
        procedureName:="SyncStammdatenToGoogle_Core", _
        entityType:="MasterData", _
        entityID:="Stammdaten", _
        correlationId:="STAMMDATEN-SYNC", _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="STAMMDATEN_SYNC_FAIL", _
        severity:="CRITICAL", _
        message:="Stammdaten sync failed. SuccessTabs=" & CStr(successCount) & "/" & CStr(totalTabs) & _
                 "; Error=" & errDesc, _
        userId:="Operator", _
        moduleName:="modStammdatenSync", _
        procedureName:="SyncStammdatenToGoogle_Core", _
        entityType:="MasterData", _
        entityID:="Stammdaten", _
        correlationId:="STAMMDATEN-SYNC"
End Sub

' ============================================================
' PUBLIC -- Test
' ============================================================

Public Sub Test_SyncStammdaten()
    Call SyncStammdatenToGoogle
End Sub


