Attribute VB_Name = "modStammdatenSync"
Option Explicit

' ============================================================
' modStammdatenSync – Export Stammdaten zu Google Sheet
'
' Schreibt tblKooperanti, tblKulture, tblConfig (Cene)
' in ein Google Sheet "Stammdaten" für die PWA.
'
' Config-Keys in tblConfig:
'   GOOGLE_STAMMDATEN_SHEET_ID   (wird automatisch erstellt)
'   GOOGLE_PWA_FOLDER_ID         (Drive Folder für PWA-Sheets)
'
' Aufruf: Button in frmMain oder manuell via SyncStammdatenToGoogle
' ============================================================

' ============================================================
' PUBLIC — Hauptfunktion
' ============================================================

Public Sub SyncStammdatenToGoogle()
    ' Exportiert Stammdaten zu Google Sheet
    ' Erstellt das Sheet automatisch wenn es nicht existiert
    
    Dim folderID As String
    Dim sheetID As String
    
    On Error GoTo EH
    
    ' Auth prüfen
    If Not IsGoogleAuthConfigured() Then
        MsgBox "Google OAuth2 nije konfigurisan!" & vbCrLf & _
               "Pokrenite RunGoogleAuthSetup iz modGoogleAuth.", _
               vbCritical, APP_NAME
        Exit Sub
    End If
    
    folderID = GetConfigValue("GOOGLE_PWA_FOLDER_ID")
    If Len(Trim$(folderID)) = 0 Then
        MsgBox "GOOGLE_PWA_FOLDER_ID nije postavljen u tblConfig!" & vbCrLf & _
               "Unesite ID Google Drive foldera za PWA.", _
               vbCritical, APP_NAME
        Exit Sub
    End If
    
    ' Sheet finden oder erstellen
    sheetID = GetConfigValue("GOOGLE_STAMMDATEN_SHEET_ID")
    
    If Len(Trim$(sheetID)) = 0 Then
        ' Suche existierendes
        sheetID = GetSpreadsheetID("Stammdaten", folderID)
    End If
    
    If Len(Trim$(sheetID)) = 0 Then
        ' Erstelle neues
        sheetID = CreateSpreadsheet("Stammdaten", folderID)
        If Len(sheetID) = 0 Then
            MsgBox "Google Sheet konnte nicht erstellt werden!", vbCritical, APP_NAME
            Exit Sub
        End If
        
        ' Tabs erstellen (Sheet1 umbenennen geht nicht einfach, also neue Tabs)
        Call AddSheetTab(sheetID, "Kooperanti")
        Call AddSheetTab(sheetID, "Kulture")
        Call AddSheetTab(sheetID, "Parcele")
        Call AddSheetTab(sheetID, "Config")
        Call AddSheetTab(sheetID, "Users")
        Call AddSheetTab(sheetID, "Fakture")
        Call AddSheetTab(sheetID, "FakturaStavke")
        Call AddSheetTab(sheetID, "SaldoOMDetail")
        Call AddSheetTab(sheetID, "Stanice")
        Call AddSheetTab(sheetID, "Kupci")
        Call AddSheetTab(sheetID, "Vozaci")
        Call AddSheetTab(sheetID, "Artikli")
        Call AddSheetTab(sheetID, "MagacinKoop")
    End If
    
    ' Sheet-ID speichern
    Call SetConfigValue("GOOGLE_STAMMDATEN_SHEET_ID", sheetID)
    
    ' Daten exportieren
    Dim successCount As Long
    
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
    
    LogInfo "SyncStammdatenToGoogle", "Export abgeschlossen: " & successCount & "/13 Tabs"
    
    MsgBox "Stammdaten exportiert: " & successCount & " od 13 tabova.", _
           vbInformation, APP_NAME
    Exit Sub

EH:
    LogErr "SyncStammdatenToGoogle"
    MsgBox "Greska pri eksportu stammdaten: " & Err.Description, vbCritical, APP_NAME
End Sub

Public Sub ExportKarticeToGoogle()
    Dim folderID As String
    Dim sheetID As String
    Dim koopData As Variant
    Dim colKoopID As Long, colAktivan As Long
    Dim i As Long, j As Long
    Dim allRows() As Variant
    Dim outRow As Long
    Dim totalRows As Long
    Dim datumOd As Date, datumDo As Date
    
    On Error GoTo EH
    
    If Not IsGoogleAuthConfigured() Then
        MsgBox "Google OAuth2 nije konfigurisan!", vbCritical, APP_NAME
        Exit Sub
    End If
    
    folderID = GetConfigValue("GOOGLE_PWA_FOLDER_ID")
    If Len(Trim$(folderID)) = 0 Then
        MsgBox "GOOGLE_PWA_FOLDER_ID nije postavljen!", vbCritical, APP_NAME
        Exit Sub
    End If
    
    ' Sheet finden oder erstellen
    sheetID = GetConfigValue("GOOGLE_KARTICE_SHEET_ID")
    If Len(Trim$(sheetID)) = 0 Then sheetID = GetSpreadsheetID("Kartice", folderID)
    If Len(Trim$(sheetID)) = 0 Then
        sheetID = CreateSpreadsheet("Kartice", folderID)
        If Len(sheetID) = 0 Then
            MsgBox "Kartice Sheet konnte nicht erstellt werden!", vbCritical, APP_NAME
            Exit Sub
        End If
    End If
    Call SetConfigValue("GOOGLE_KARTICE_SHEET_ID", sheetID)
    
    ' Kooperanten laden
    koopData = GetTableData(TBL_KOOPERANTI)
    If IsEmpty(koopData) Then Exit Sub
    koopData = ExcludeStornirano(koopData, TBL_KOOPERANTI)
    If IsEmpty(koopData) Then Exit Sub
    
    colKoopID = GetColumnIndex(TBL_KOOPERANTI, "KooperantID")
    colAktivan = GetColumnIndex(TBL_KOOPERANTI, "Aktivan")
    
    datumOd = DateSerial(Year(Date), 1, 1)
    datumDo = Date
    
    ' Aktive Kooperanten sammeln + Kartice generieren
    Dim koopList() As String
    Dim koopCount As Long
    ReDim koopList(1 To UBound(koopData, 1))
    
    For i = 1 To UBound(koopData, 1)
        If CStr(koopData(i, colAktivan)) <> "Ne" Then
            koopCount = koopCount + 1
            koopList(koopCount) = CStr(koopData(i, colKoopID))
        End If
    Next i
    
    If koopCount = 0 Then Exit Sub

    ' Kartice generieren und Zeilen zählen
    Dim karticaResults() As Variant
    ReDim karticaResults(1 To koopCount)
    totalRows = 1 ' Header

    For i = 1 To koopCount
        karticaResults(i) = ReportKarticaKooperanta(koopList(i), datumOd, datumDo)
        If Not IsEmpty(karticaResults(i)) Then
            totalRows = totalRows + UBound(karticaResults(i), 1)
        End If
    Next i

    ' Ergebnis bauen
    ReDim allRows(1 To totalRows, 1 To 8)
    allRows(1, 1) = "KooperantID"
    allRows(1, 2) = "Datum"
    allRows(1, 3) = "BrojDok"
    allRows(1, 4) = "BrojParcele"
    allRows(1, 5) = "Opis"
    allRows(1, 6) = "Zaduzenje"
    allRows(1, 7) = "Razduzenje"
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
    
    ' Kürzen und schreiben
    If outRow < totalRows Then
        Dim finalRows() As Variant
        Dim r As Long, c As Long
        ReDim finalRows(1 To outRow, 1 To 8)
        For r = 1 To outRow
            For c = 1 To 8
                finalRows(r, c) = allRows(r, c)
            Next c
        Next r
        WriteSheetData sheetID, "Sheet1", finalRows
    Else
        WriteSheetData sheetID, "Sheet1", allRows
    End If
    
    LogInfo "ExportKarticeToGoogle", outRow - 1 & " Zeilen fuer " & koopCount & " Kooperanten"
    MsgBox "Kartice exportiert: " & (outRow - 1) & " stavki za " & koopCount & " kooperanata.", vbInformation, APP_NAME
    Exit Sub

EH:
    LogErr "ExportKarticeToGoogle"
    MsgBox "Greska: " & Err.Description, vbCritical, APP_NAME
End Sub

Public Sub ExportMgmtReports()
    Dim folderID As String
    Dim sheetID As String
    
    On Error GoTo EH
    
    If Not IsGoogleAuthConfigured() Then
        MsgBox "Google OAuth2 nije konfigurisan!", vbCritical, APP_NAME
        Exit Sub
    End If
    
    folderID = GetConfigValue("GOOGLE_PWA_FOLDER_ID")
    If Len(Trim$(folderID)) = 0 Then Exit Sub
    
    sheetID = GetConfigValue("GOOGLE_MGMT_SHEET_ID")
    If Len(Trim$(sheetID)) = 0 Then sheetID = GetSpreadsheetID("MgmtReports", folderID)
    If Len(Trim$(sheetID)) = 0 Then
        sheetID = CreateSpreadsheet("MgmtReports", folderID)
        If Len(sheetID) = 0 Then Exit Sub
        Call AddSheetTab(sheetID, "SaldoOM")
        Call AddSheetTab(sheetID, "SaldoKupci")
        Call AddSheetTab(sheetID, "OtkupPoOM")
        Call AddSheetTab(sheetID, "PredatoPoKupcu")
    End If
    Call SetConfigValue("GOOGLE_MGMT_SHEET_ID", sheetID)
    
    Dim ok As Long
    If ExportSaldoOM(sheetID) Then ok = ok + 1
    If ExportSaldoKupci(sheetID) Then ok = ok + 1
    If ExportOtkupPoOM(sheetID) Then ok = ok + 1
    If ExportPredatoPoKupcu(sheetID) Then ok = ok + 1
    
    MsgBox "MgmtReports exportiert: " & ok & "/4", vbInformation, APP_NAME
    Exit Sub
EH:
    LogErr "ExportMgmtReports"
    MsgBox "Greska: " & Err.Description, vbCritical, APP_NAME
End Sub

Private Function ExportOtkupPoOM(ByVal sheetID As String) As Boolean
    Dim data As Variant
    Dim colStanica As Long, colVrsta As Long, colKlasa As Long
    Dim colKolicina As Long, colAmb As Long, colCena As Long
    Dim i As Long
    
    On Error GoTo EH
    
    data = GetTableData(TBL_OTKUP)
    If Not IsEmpty(data) Then data = ExcludeStornirano(data, TBL_OTKUP)
    If IsEmpty(data) Then
        ExportOtkupPoOM = False
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
        vals(1) = vals(1) + CDbl(Nz(data(i, colAmb), 0))
        vals(2) = vals(2) + CDbl(data(i, colKolicina)) * CDbl(data(i, colCena))
        vals(3) = vals(3) + 1
        dict(key) = vals
    Next i
    
    If dict.count = 0 Then
        ExportOtkupPoOM = False
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To dict.count + 1, 1 To 7)
    result(1, 1) = "StanicaID"
    result(1, 2) = "VrstaVoca"
    result(1, 3) = "Klasa"
    result(1, 4) = "Kolicina"
    result(1, 5) = "Ambalaza"
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

Private Function ExportPredatoPoKupcu(ByVal sheetID As String) As Boolean
    Dim data As Variant
    Dim colKupac As Long, colVrsta As Long, colKlasa As Long
    Dim colKolicina As Long, colAmb As Long, colCena As Long
    Dim i As Long
    
    On Error GoTo EH
    
    data = GetTableData(TBL_PRIJEMNICA)
    If Not IsEmpty(data) Then data = ExcludeStornirano(data, TBL_PRIJEMNICA)
    If IsEmpty(data) Then
        ExportPredatoPoKupcu = False
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
        vals(1) = vals(1) + CDbl(Nz(data(i, colAmb), 0))
        vals(2) = vals(2) + CDbl(data(i, colKolicina)) * CDbl(data(i, colCena))
        vals(3) = vals(3) + 1
        dict(key) = vals
    Next i
    
    If dict.count = 0 Then
        ExportPredatoPoKupcu = False
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To dict.count + 1, 1 To 7)
    result(1, 1) = "KupacID"
    result(1, 2) = "VrstaVoca"
    result(1, 3) = "Klasa"
    result(1, 4) = "Kolicina"
    result(1, 5) = "Ambalaza"
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
        
        result(r + 2, 1) = CStr(Nz(kupacNaziv, parts(0)))
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
    
    ' ReportSaldoOM gibt Daten in ein ListBox — wir brauchen die Rohdaten
    ' Hier vereinfacht: OM-Saldo aus tblNovac berechnen
    Dim data As Variant
    Dim colOMID As Long, colTip As Long, colIsplata As Long, colUplata As Long
    Dim i As Long
    
    data = GetTableData(TBL_NOVAC)
    If Not IsEmpty(data) Then data = ExcludeStornirano(data, TBL_NOVAC)
    If IsEmpty(data) Then
        ExportSaldoOM = False
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
        ExportSaldoOM = False
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
                v(1) = v(1) + CDbl(Nz(otkData(i, colOtkKg), 0))
                v(2) = v(2) + CDbl(Nz(otkData(i, colOtkKg), 0)) * CDbl(Nz(otkData(i, colOtkCena), 0))
                v(3) = v(3) + CDbl(Nz(otkData(i, colOtkAmb), 0))
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
                nKoop = CStr(Nz(novData(i, colNovKoop), ""))
                If dict.Exists(nKoop) Then
                    v = dict(nKoop)
                    v(4) = v(4) + CDbl(Nz(novData(i, colNovIsplata), 0))
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
                mKoop = CStr(Nz(magData(i, colMagKoop), ""))
                If dict.Exists(mKoop) Then
                    v = dict(mKoop)
                    v(5) = v(5) + CDbl(Nz(magData(i, colMagVrednost), 0))
                    dict(mKoop) = v
                End If
            End If
        Next i
    End If
    
    If dict.count = 0 Then ExportSaldoOMDetail = False: Exit Function
    
    ' --- Build result ---
    Dim result() As Variant
    ReDim result(1 To dict.count + 1, 1 To 9)
    result(1, 1) = "KooperantID"
    result(1, 2) = "Kooperant"
    result(1, 3) = "StanicaID"
    result(1, 4) = "Kolicina"
    result(1, 5) = "Vrednost"
    result(1, 6) = "Isplaceno"
    result(1, 7) = "AgroZaduzenje"
    result(1, 8) = "Saldo"
    result(1, 9) = "Ambalaza"
    
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
        result(r + 2, 2) = CStr(Nz(koopName, "")) & " " & CStr(Nz(koopPrezime, ""))
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
        ExportSaldoKupci = False
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
                Dim pID As String
                pID = CStr(novData(i, colNovPartner))
                If dict.Exists(pID) Then
                    vals = dict(pID)
                    vals(1) = vals(1) + CDbl(novData(i, colNovUplata))
                    dict(pID) = vals
                End If
            End If
        Next i
    End If
    
    If dict.count = 0 Then
        ExportSaldoKupci = False
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
        result(r + 2, 2) = CStr(Nz(kupacNaziv, keys(r)))
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
' PRIVATE — Export einzelner Tabellen
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
        ExportKooperanti = False
        Exit Function
    End If
    
    data = ExcludeStornirano(data, TBL_KOOPERANTI)
    If IsEmpty(data) Then
        ExportKooperanti = False
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
        If CStr(data(i, colAktivan)) <> "Ne" Then activeCount = activeCount + 1
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
        If CStr(data(i, colAktivan)) <> "Ne" Then
            outRow = outRow + 1
            result(outRow, 1) = CStr(data(i, colID))
            result(outRow, 2) = CStr(data(i, colIme))
            result(outRow, 3) = CStr(data(i, colPrezime))
            result(outRow, 4) = CStr(data(i, colStanica))
            result(outRow, 5) = CStr(Nz(data(i, colMesto), ""))
            result(outRow, 6) = CStr(Nz(data(i, colTelefon), ""))
            result(outRow, 7) = CStr(Nz(data(i, colBPG), ""))
            result(outRow, 8) = CStr(Nz(data(i, colAdresa), ""))
            result(outRow, 9) = CStr(Nz(data(i, colJMBG), ""))
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
        ExportKulture = False
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
        ExportParcele = False
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
    
    ReDim result(1 To UBound(data, 1) + 1, 1 To 20)
    
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
    
    For i = 1 To UBound(data, 1)
        result(i + 1, 1) = CStr(Nz(data(i, colID), ""))
        result(i + 1, 2) = CStr(Nz(data(i, colKoop), ""))
        result(i + 1, 3) = CStr(Nz(data(i, colKatBroj), ""))
        result(i + 1, 4) = CStr(Nz(data(i, colKatOpstina), ""))
        result(i + 1, 5) = CStr(Nz(data(i, colKultura), ""))
        result(i + 1, 6) = CStr(Nz(data(i, colPovrsina), ""))
        result(i + 1, 7) = CStr(Nz(data(i, colGGAP), ""))
        result(i + 1, 8) = CStr(Nz(data(i, colAktivna), ""))
        result(i + 1, 9) = CStr(Nz(data(i, colGeoStatus), ""))
        result(i + 1, 10) = CStr(Nz(data(i, colGeoSource), ""))
        result(i + 1, 11) = CStr(Nz(data(i, colN), ""))
        result(i + 1, 12) = CStr(Nz(data(i, colE), ""))
        result(i + 1, 13) = CStr(Nz(data(i, colLat), ""))
        result(i + 1, 14) = CStr(Nz(data(i, colLng), ""))
        result(i + 1, 15) = CStr(Nz(data(i, colPolygon), ""))
        result(i + 1, 16) = CStr(Nz(data(i, colMeteo), ""))
        result(i + 1, 17) = CStr(Nz(data(i, colRizik), ""))
        result(i + 1, 18) = CStr(Nz(data(i, colDatumGeo), ""))
        result(i + 1, 19) = CStr(Nz(data(i, colDatumAzur), ""))
        result(i + 1, 20) = CStr(Nz(data(i, colNapomena), ""))
    Next i
    
    ExportParcele = WriteSheetData(sheetID, "Parcele", result)
    Exit Function

EH:
    LogErr "ExportParcele"
    ExportParcele = False
End Function
Private Function ExportStanice(ByVal sheetID As String) As Boolean
    Dim data As Variant
    Dim result() As Variant
    Dim colID As Long, colNaziv As Long, colMesto As Long, colAktivan As Long
    Dim i As Long, outRow As Long
    
    On Error GoTo EH
    
    data = GetTableData(TBL_STANICE)
    If IsEmpty(data) Then
        ExportStanice = False
        Exit Function
    End If
    data = ExcludeStornirano(data, TBL_STANICE)
    
    colID = GetColumnIndex(TBL_STANICE, "StanicaID")
    colNaziv = GetColumnIndex(TBL_STANICE, "Naziv")
    colMesto = GetColumnIndex(TBL_STANICE, "Mesto")
    colAktivan = GetColumnIndex(TBL_STANICE, "Aktivan")
    
    ' Erst zählen wieviele aktiv
    Dim cnt As Long: cnt = 0
    For i = 1 To UBound(data, 1)
        If CStr(Nz(data(i, colAktivan), "")) = "Aktivan" Then cnt = cnt + 1
    Next i
    
    If cnt = 0 Then
        ExportStanice = False
        Exit Function
    End If
    
    ReDim result(1 To cnt + 1, 1 To 3)
    
    ' Header
    result(1, 1) = "StanicaID"
    result(1, 2) = "Naziv"
    result(1, 3) = "Mesto"
    
    outRow = 2
    For i = 1 To UBound(data, 1)
        If CStr(Nz(data(i, colAktivan), "")) = "Aktivan" Then
            result(outRow, 1) = CStr(data(i, colID))
            result(outRow, 2) = CStr(Nz(data(i, colNaziv), ""))
            result(outRow, 3) = CStr(Nz(data(i, colMesto), ""))
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
        ExportKupci = False
        Exit Function
    End If
    data = ExcludeStornirano(data, TBL_KUPCI)
    
    colID = GetColumnIndex(TBL_KUPCI, "KupacID")
    colNaziv = GetColumnIndex(TBL_KUPCI, "Naziv")
    colMesto = GetColumnIndex(TBL_KUPCI, "Mesto")
    colAktivan = GetColumnIndex(TBL_KUPCI, "Aktivan")
    
    ' Erst zählen wieviele aktiv
    Dim cnt As Long: cnt = 0
    For i = 1 To UBound(data, 1)
        If CStr(Nz(data(i, colAktivan), "")) = "Aktivan" Then cnt = cnt + 1
    Next i
    
    If cnt = 0 Then
        ExportKupci = False
        Exit Function
    End If
    
    ReDim result(1 To cnt + 1, 1 To 3)
    
    ' Header
    result(1, 1) = "KupacID"
    result(1, 2) = "Naziv"
    result(1, 3) = "Mesto"
    
    outRow = 2
    For i = 1 To UBound(data, 1)
        If CStr(Nz(data(i, colAktivan), "")) = "Aktivan" Then
            result(outRow, 1) = CStr(data(i, colID))
            result(outRow, 2) = CStr(Nz(data(i, colNaziv), ""))
            result(outRow, 3) = CStr(Nz(data(i, colMesto), ""))
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
        ExportVozaci = False
        Exit Function
    End If
    data = ExcludeStornirano(data, TBL_VOZACI)
    
    colID = GetColumnIndex(TBL_VOZACI, "VozacID")
    colIme = GetColumnIndex(TBL_VOZACI, "Ime")
    colPrezime = GetColumnIndex(TBL_VOZACI, "Prezime")
    colTelefon = GetColumnIndex(TBL_VOZACI, "Telefon")
    colKapacitetKG = GetColumnIndex(TBL_VOZACI, "KapacitetKG")
    colAktivan = GetColumnIndex(TBL_VOZACI, "Aktivan")
    
    ' Erst zählen wieviele aktiv
    Dim cnt As Long: cnt = 0
    For i = 1 To UBound(data, 1)
        If CStr(Nz(data(i, colAktivan), "")) = "Aktivan" Then cnt = cnt + 1
    Next i
    
    If cnt = 0 Then
        ExportVozaci = False
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
        If CStr(Nz(data(i, colAktivan), "")) = "Aktivan" Then
            result(outRow, 1) = CStr(data(i, colID))
            result(outRow, 2) = CStr(Nz(data(i, colIme), ""))
            result(outRow, 3) = CStr(Nz(data(i, colPrezime), ""))
            result(outRow, 4) = CStr(Nz(data(i, colTelefon), ""))
            result(outRow, 5) = CStr(Nz(data(i, colKapacitetKG), ""))
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
        ExportArtikli = False
        Exit Function
    End If
    data = ExcludeStornirano(data, TBL_ARTIKLI)
    
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
    
    ' Erst zählen wieviele aktiv
    Dim cnt As Long: cnt = 0
    For i = 1 To UBound(data, 1)
        If CStr(Nz(data(i, colAktivan), "")) = "Aktivan" Then cnt = cnt + 1
    Next i
    
    If cnt = 0 Then
        ExportArtikli = False
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
        If CStr(Nz(data(i, colAktivan), "")) = "Aktivan" Then
            result(outRow, 1) = CStr(data(i, colArtID))
            result(outRow, 2) = CStr(Nz(data(i, colNaziv), ""))
            result(outRow, 3) = CStr(Nz(data(i, colTip), ""))
            result(outRow, 4) = CStr(Nz(data(i, colJM), ""))
            result(outRow, 5) = CStr(Nz(data(i, colCena), ""))
            result(outRow, 6) = CStr(Nz(data(i, colDoza), ""))
            result(outRow, 7) = CStr(Nz(data(i, colKultura), ""))
            result(outRow, 8) = CStr(Nz(data(i, colPak), ""))
            result(outRow, 9) = CStr(Nz(data(i, colBarKod), ""))
            result(outRow, 10) = CStr(Nz(data(i, colKarenca), ""))
            result(outRow, 11) = CStr(Nz(data(i, colAktivan), ""))
            
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
        ExportMagacinKoop = False
        Exit Function
    End If
    magData = ExcludeStornirano(magData, TBL_MAGACIN)
    
    Dim colMKoop As Long: colMKoop = GetColumnIndex(TBL_MAGACIN, "KooperantID")
    Dim colMArt As Long: colMArt = GetColumnIndex(TBL_MAGACIN, "ArtikalID")
    Dim colMTip As Long: colMTip = GetColumnIndex(TBL_MAGACIN, "Tip")
    Dim colMKol As Long: colMKol = GetColumnIndex(TBL_MAGACIN, "Kolicina")
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    For i = 1 To UBound(magData, 1)
        If CStr(Nz(magData(i, colMTip), "")) = "Izlaz" Then
            Dim koopID As String
            Dim artID As String
            Dim key As String
            
            koopID = CStr(Nz(magData(i, colMKoop), ""))
            artID = CStr(Nz(magData(i, colMArt), ""))
            
            If koopID <> "" And artID <> "" Then
                key = koopID & "|" & artID
                
                If Not dict.Exists(key) Then dict.Add key, 0#
                dict(key) = CDbl(dict(key)) + CDbl(Nz(magData(i, colMKol), 0))
            End If
        End If
    Next i
    
    If dict.count = 0 Then
        ExportMagacinKoop = False
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
            artID = CStr(Nz(artData(i, colArtID), ""))
            If artID <> "" Then
                If Not artDict.Exists(artID) Then
                    artDict.Add artID, Array( _
                        CStr(Nz(artData(i, colNaziv), "")), _
                        CStr(Nz(artData(i, colTip), "")), _
                        CStr(Nz(artData(i, colJM), "")), _
                        CStr(Nz(artData(i, colCena), "")), _
                        CStr(Nz(artData(i, colDoza), "")), _
                        CStr(Nz(artData(i, colPak), "")), _
                        CStr(Nz(artData(i, colKarenca), "")))
                End If
            End If
        Next i
    End If
    
    cnt = dict.count
    If cnt = 0 Then
        ExportMagacinKoop = False
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
        ExportConfig = False
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
    
    ' Zählen
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
    ' Credentials ausschließen
    If Left$(keyStr, 7) = "GOOGLE_" Then Exit Function
    If Left$(keyStr, 4) = "SEF_" Then Exit Function
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
                If CStr(koopData(i, colKAktivan)) <> "Ne" Then
                    Dim kPin As String
                    kPin = Trim$(CStr(Nz(koopData(i, colKPIN), "")))
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
                If CStr(staData(i, colSAktivan)) <> "Ne" Then
                    Dim sPin As String
                    sPin = Trim$(CStr(Nz(staData(i, colSPIN), "")))
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
                If CStr(vozData(i, colVAktivan)) <> "Ne" Then
                    Dim vPin As String
                    vPin = Trim$(CStr(Nz(vozData(i, colVPIN), "")))
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
                        Dim ri As Long, ci As Long
                        For ri = 1 To outRow - 1
                            For ci = 1 To 5
                                tmp(ri, ci) = result(ri, ci)
                            Next ci
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
    
    ' Auf tatsächliche Größe kürzen
    If outRow < totalRows Then
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
    If IsEmpty(data) Then ExportFakture = False: Exit Function
    
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
            fakID = Trim$(CStr(Nz(novData(n, colNovFaktura), "")))
            If Len(fakID) > 0 Then
                Dim tip As String
                tip = CStr(novData(n, colNovTip))
                If tip = NOV_KUPCI_UPLATA Then
                    If Not dictPlaceno.Exists(fakID) Then dictPlaceno.Add fakID, 0#
                    dictPlaceno(fakID) = dictPlaceno(fakID) + CDbl(Nz(novData(n, colNovUplata), 0))
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
        iznos = CDbl(Nz(data(i, colIznos), 0))
        Dim fID As String
        fID = CStr(data(i, colID))
        Dim placeno As Double
        placeno = 0
        If dictPlaceno.Exists(fID) Then placeno = dictPlaceno(fID)
        
        result(i + 1, 1) = fID
        result(i + 1, 2) = CStr(data(i, colBroj))
        result(i + 1, 3) = CStr(data(i, colDatum))
        result(i + 1, 4) = CStr(data(i, colKupac))
        result(i + 1, 5) = CStr(Nz(kupacNaziv, data(i, colKupac)))
        result(i + 1, 6) = CStr(iznos)
        result(i + 1, 7) = CStr(placeno)
        result(i + 1, 8) = CStr(iznos - placeno)
        result(i + 1, 9) = CStr(data(i, colStatus))
        result(i + 1, 10) = CStr(Nz(data(i, colSEFStatus), ""))
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
    If IsEmpty(data) Then ExportFakturaStavke = False: Exit Function
    
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
            Dim pID As String
            pID = CStr(prijData(p, colPrijPrijID))
            If Not dictZbirna.Exists(pID) Then
                dictZbirna.Add pID, CStr(Nz(prijData(p, colPrijZbirna), ""))
            End If
            If Not dictVrsta.Exists(pID) Then
                dictVrsta.Add pID, CStr(Nz(prijData(p, colPrijVrsta), ""))
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
    result(1, 7) = "Kolicina"
    result(1, 8) = "Cena"
    result(1, 9) = "Iznos"
    
    For i = 1 To UBound(data, 1)
        Dim prijemnicaID As String
        prijemnicaID = CStr(Nz(data(i, colPrijID), ""))
        Dim kg As Double, cena As Double
        kg = CDbl(Nz(data(i, colKolicina), 0))
        cena = CDbl(Nz(data(i, colCena), 0))
        
        result(i + 1, 1) = CStr(data(i, colFakID))
        result(i + 1, 2) = prijemnicaID
        result(i + 1, 3) = CStr(Nz(data(i, colBrojPrij), ""))
        result(i + 1, 4) = ""
        If dictZbirna.Exists(prijemnicaID) Then result(i + 1, 4) = dictZbirna(prijemnicaID)
        result(i + 1, 5) = ""
        If dictVrsta.Exists(prijemnicaID) Then result(i + 1, 5) = dictVrsta(prijemnicaID)
        result(i + 1, 6) = CStr(Nz(data(i, colKlasa), ""))
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


' ============================================================
' PUBLIC — Test
' ============================================================

Public Sub Test_SyncStammdaten()
    Call SyncStammdatenToGoogle
End Sub


