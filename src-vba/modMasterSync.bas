Attribute VB_Name = "modMasterSync"
 Option Explicit

' ============================================================
' modMasterSync – Import OTK-Sheets ? tblOtkup
'
' Liest alle Google Sheets "OTK-*" aus dem PWA-Folder,
' importiert neue Zeilen (SyncStatus != "Synced?Master")
' in tblOtkup, und schreibt SyncStatus zurück.
'
' Flow:
'   1. Liste alle OTK-* Sheets im PWA-Folder
'   2. Pro Sheet: ReadSheetData ? prüfe SyncStatus
'   3. Neue Zeilen ? Validierung ? AppendRow tblOtkup
'   4. SyncStatus ? "Synced?Master" zurückschreiben
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
' PUBLIC — Hauptfunktion
' ============================================================
Public Sub ImportOtkupFromPWA()
    Call ImportOtkupFromPWA_Core(True)
End Sub

Private Function ImportOtkupFromPWA_Core(ByVal showMessages As Boolean) As Boolean
    Dim folderID As String
    Dim sheetIDs As Collection
    Dim sheetNames As Collection
    Dim i As Long
    Dim totalImported As Long
    Dim totalSkipped As Long
    Dim totalErrors As Long

    On Error GoTo EH

    mLastPWAFatalSyncError = False

    If Not IsGoogleAuthConfigured() Then
        MarkPWAFatalSyncError "ImportOtkupFromPWA_Core", "Google OAuth2 nije konfigurisan."
        If showMessages Then MsgBox "Google OAuth2 nije konfigurisan!", vbCritical, APP_NAME
        ImportOtkupFromPWA_Core = False
        Exit Function
    End If

    folderID = GetConfigValue("GOOGLE_PWA_FOLDER_ID")
    If Len(Trim$(folderID)) = 0 Then
        MarkPWAFatalSyncError "ImportOtkupFromPWA_Core", "GOOGLE_PWA_FOLDER_ID nije postavljen."
        If showMessages Then MsgBox "GOOGLE_PWA_FOLDER_ID nije postavljen!", vbCritical, APP_NAME
        ImportOtkupFromPWA_Core = False
        Exit Function
    End If

    LogInfo "ImportOtkupFromPWA", "Import gestartet"

    Set sheetIDs = New Collection
    Set sheetNames = New Collection

    If Not FindOTKSheets(folderID, sheetIDs, sheetNames) Then
        MarkPWAFatalSyncError "ImportOtkupFromPWA_Core", _
            "FindOTKSheets failed. Drive list could not be loaded."
        If showMessages Then MsgBox "Google Drive lista OTK fajlova nije ucitana. Proveri log.", vbCritical, APP_NAME
        ImportOtkupFromPWA_Core = False
        Exit Function
    End If

    If sheetIDs.count = 0 Then
        If showMessages Then MsgBox "Nema OTK-* fajlova u PWA folderu.", vbInformation, APP_NAME
        ImportOtkupFromPWA_Core = True
        Exit Function
    End If

    For i = 1 To sheetIDs.count
        Dim imported As Long
        Dim skipped As Long
        Dim errors As Long

        Call ImportOneOTKSheet(CStr(sheetIDs(i)), CStr(sheetNames(i)), imported, skipped, errors)

        totalImported = totalImported + imported
        totalSkipped = totalSkipped + skipped
        totalErrors = totalErrors + errors
    Next i

    LogInfo "ImportOtkupFromPWA", _
            "Import abgeschlossen: " & totalImported & " importiert, " & _
            totalSkipped & " preskoceno, " & totalErrors & " greske aus " & sheetIDs.count & " fajlova"

    If showMessages Then
        MsgBox "Uvoz zavrsen!" & vbCrLf & vbCrLf & _
               "Fajlova: " & sheetIDs.count & vbCrLf & _
               "Uvezeno: " & totalImported & vbCrLf & _
               "Preskoceno: " & totalSkipped & vbCrLf & _
               "Greske: " & totalErrors, _
               vbInformation, APP_NAME
    End If

    ImportOtkupFromPWA_Core = Not mLastPWAFatalSyncError
    Exit Function

EH:
    MarkPWAFatalSyncError "ImportOtkupFromPWA_Core", Err.description
    LogErr "ImportOtkupFromPWA_Core"
    If showMessages Then MsgBox "Greska pri uvozu: " & Err.description, vbCritical, APP_NAME
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
        MsgBox "PWA uvoz nije potvrden. Promene su vracene zbog fatal sync greške. Proveri log.", _
               vbCritical, APP_NAME
        Exit Sub
    End If

    tx.CommitTx

    MsgBox "PWA uvoz završen i potvrden.", vbInformation, APP_NAME
    Exit Sub

EH:
    LogErr "ImportOtkupFromPWA_TX"
    If Not tx Is Nothing Then tx.RollbackTx
    MsgBox "Greska pri uvozu, promene vracene: " & Err.description, vbCritical, APP_NAME
End Sub
'TODO definieren wo dies stehen soll. Logisch bei Stammdaten sync und syncen immer wenn stammdaten gesynct sind.
Public Sub CreateOTKSheetsForAllStanice()
    Dim data As Variant
    Dim colID As Long, colNaziv As Long, colAktivan As Long
    Dim folderID As String
    Dim i As Long
    Dim sheetName As String
    Dim existingID As String
    Dim created As Long
    
    On Error GoTo EH
    
    folderID = GetConfigValue("GOOGLE_PWA_FOLDER_ID")
    If Len(Trim$(folderID)) = 0 Then
        MsgBox "GOOGLE_PWA_FOLDER_ID nije postavljen!", vbCritical, APP_NAME
        Exit Sub
    End If
    
    data = GetTableData(TBL_STANICE)
    If IsEmpty(data) Then Exit Sub
    
    colID = GetColumnIndex(TBL_STANICE, "StanicaID")
    colNaziv = GetColumnIndex(TBL_STANICE, "Naziv")
    colAktivan = GetColumnIndex(TBL_STANICE, "Aktivan")
    
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colAktivan)) <> "Ne" Then
            sheetName = "OTK-" & CStr(data(i, colID))
            
            existingID = GetSpreadsheetID(sheetName, folderID)
            
            If Len(existingID) = 0 Then
                Dim newID As String
                newID = CreateSpreadsheet(sheetName, folderID)
                
                If Len(newID) > 0 Then
                    ' Header setzen
                    Dim headers(1 To 1, 1 To 22) As Variant
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
                    
                    WriteSheetData newID, "Sheet1", headers
                    
                    created = created + 1
                    LogInfo "CreateOTKSheets", "Erstellt: " & sheetName
                End If
            End If
        End If
    Next i
    
    MsgBox "Kreirano " & created & " novih OTK fajlova.", vbInformation, APP_NAME
    Exit Sub

EH:
    LogErr "CreateOTKSheetsForAllStanice"
    MsgBox "Greska: " & Err.description, vbCritical, APP_NAME
End Sub

Public Function AutoCreateOtpremniceFromPWA() As Long
    ' Nach ImportOtkupFromPWA: erstellt Otpremnice für PWA-Otkupi mit VozacID
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
    
    colID = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    colSt = GetColumnIndex(TBL_OTKUP, COL_OTK_STANICA)
    colDat = GetColumnIndex(TBL_OTKUP, COL_OTK_DATUM)
    colVoz = GetColumnIndex(TBL_OTKUP, COL_OTK_VOZAC)
    colOtpID = GetColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID)
    colKlasa = GetColumnIndex(TBL_OTKUP, COL_OTK_KLASA)
    colVrsta = GetColumnIndex(TBL_OTKUP, COL_OTK_VRSTA)
    colSorta = GetColumnIndex(TBL_OTKUP, COL_OTK_SORTA)
    colKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
    colCena = GetColumnIndex(TBL_OTKUP, COL_OTK_CENA)
    colTipAmb = GetColumnIndex(TBL_OTKUP, COL_OTK_TIP_AMB)
    colKolAmb = GetColumnIndex(TBL_OTKUP, COL_OTK_KOL_AMB)
    
    ' Sammle unverknüpfte Otkupi MIT VozacID ? gruppiere nach Key
    ' Key = StanicaID|Datum|VozacID|Klasa
    Dim groups As Object
    Set groups = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim vozID As String: vozID = Trim$(CStr(Nz(data(i, colVoz), "")))
        Dim otpID As String: otpID = Trim$(CStr(Nz(data(i, colOtpID), "")))
        
        ' Nur Otkupi ohne Otpremnica UND mit VozacID
        If otpID = "" And vozID <> "" Then
            Dim gKey As String
            gKey = CStr(data(i, colSt)) & "|" & _
                   Format$(CDate(data(i, colDat)), "YYYY-MM-DD") & "|" & _
                   vozID & "|" & _
                   CStr(Nz(data(i, colKlasa), ""))
            
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
    
    ' Für jede Gruppe: Otpremnica erstellen + Otkupi verknüpfen
    Dim created As Long
    Dim keys As Variant: keys = groups.keys
    Dim k As Long
    
    ' Otpremnica-Zähler pro Stanica+Datum vorladen
    Dim otpAll As Variant
    otpAll = GetTableData(TBL_OTPREMNICA)
    If Not IsEmpty(otpAll) Then otpAll = ExcludeStornirano(otpAll, TBL_OTPREMNICA)
    
    Dim colOtpSt As Long: colOtpSt = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_STANICA)
    Dim colOtpDat As Long: colOtpDat = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_DATUM)
    
    ' Dict: "StanicaID|Datum" ? count
    Dim seqDict As Object
    Set seqDict = CreateObject("Scripting.Dictionary")
    
    If Not IsEmpty(otpAll) Then
        Dim oi As Long
        For oi = 1 To UBound(otpAll, 1)
            Dim seqKey As String
            seqKey = CStr(otpAll(oi, colOtpSt)) & "|" & _
                     Format$(CDate(otpAll(oi, colOtpDat)), "YYYY-MM-DD")
            If seqDict.Exists(seqKey) Then
                seqDict(seqKey) = seqDict(seqKey) + 1
            Else
                seqDict.Add seqKey, 1
            End If
        Next oi
    End If
    
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
            totalKol = totalKol + CDbl(Nz(data(ri, colKol), 0))
            totalAmb = totalAmb + CLng(Nz(data(ri, colKolAmb), 0))
        Next r
        
        ' BrojOtpremnice: {StanicaNum}/{DDMM}-{seq}
        Dim staNum As String
        staNum = CStr(CLng(Mid$(parts(0), 4)))  ' ST-00001 ? 1
        
        Dim datParts() As String
        datParts = Split(parts(1), "-")  ' 2026-06-15 ? (2026, 06, 15)
        Dim ddmm As String
        ddmm = datParts(2) & datParts(1)  ' 1506
        
        Dim sKey As String
        sKey = parts(0) & "|" & parts(1)
        Dim seq As Long
        If seqDict.Exists(sKey) Then
            seq = seqDict(sKey) + 1
            seqDict(sKey) = seq
        Else
            seq = 1
            seqDict.Add sKey, 1
        End If
        
        Dim brojOtp As String
        brojOtp = staNum & "/" & ddmm & "-" & seq
        
        ' Otpremnica erstellen (BrojZbirne leer — Vozac/Operator setzt später)
        Dim newOtpID As String
        newOtpID = SaveOtpremnica_TX( _
            CDate(parts(1)), _
            parts(0), _
            parts(2), _
            brojOtp, _
            "", _
            CStr(Nz(data(firstRow, colVrsta), "")), _
            CStr(Nz(data(firstRow, colSorta), "")), _
            totalKol, _
            CDbl(Nz(data(firstRow, colCena), 0)), _
            CStr(Nz(data(firstRow, colTipAmb), "")), _
            totalAmb, _
            parts(3) _
        )
        
        If newOtpID <> "" Then
            ' Alle Otkupi dieser Gruppe verknüpfen
            For r = 1 To grpRows.count
                ri = grpRows(r)
                Dim otkupID As String: otkupID = CStr(data(ri, colID))
                Dim otkRows As Collection
                Set otkRows = FindRows(TBL_OTKUP, COL_OTK_ID, otkupID)
                If otkRows.count > 0 Then
                    UpdateCell TBL_OTKUP, otkRows(1), COL_OTK_OTPREMNICA_ID, newOtpID
                End If
            Next r
            created = created + 1
        End If
    Next k
    
    AutoCreateOtpremniceFromPWA = created
End Function

' ============================================================
' PRIVATE — Find OTK-* Sheets in Folder
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
        ' Suche nächstes "id"
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
' PRIVATE — Import eines einzelnen OTK-Sheets
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
        ' Prüfe SyncStatus
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
                sheetVozac = Trim$(CStr(Nz(data(i, GS_VOZAC_ID), "")))
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
                    newOtkupID = ImportRowToTblOtkup(data, i, clientRecordID)
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
            ' Bereits importiert oder Error ? überspringen
            outSkipped = outSkipped + 1
        End If

NextImportRow:
    Next i
    
    ' SyncStatus zurückschreiben in Google Sheet
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
    LogErr "ImportOneOTKSheet", "Sheet: " & sheetName
    outErrors = outErrors + 1
End Sub

' ============================================================
' PRIVATE — Validierung
' ============================================================

Private Function ValidatePWAOtkup(ByVal data As Variant, ByVal row As Long) As String
    ' Prüft Pflichtfelder und Plausibilität
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
        If CStr(Nz(data(i, colCRID), "")) = clientRecordID Then
            IsDuplicateInMaster = True
            Exit Function
        End If
    Next i
    
    IsDuplicateInMaster = False
End Function

' ============================================================
' PRIVATE — Import Row
' ============================================================

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
    If Len(tipAmb) = 0 Then tipAmb = "12/1"
    
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
    
    ' StanicaID aus Kooperant holen
    stanicaID = CStr(Nz(LookupValue(TBL_KOOPERANTI, "KooperantID", kooperantID, COL_KOOP_STANICA), ""))
    
    ' Wenn OtkupacID = StanicaID (wie bei deinem Setup), nutze das
    If Len(stanicaID) = 0 And Left$(otkupacID, 3) = "ST-" Then
        stanicaID = otkupacID
    End If
    
    ' KulturaID Lookup
    kulturaID = CStr(Nz(LookupValue(TBL_KULTURE, "VrstaVoca", vrstaVoca, "KulturaID"), ""))
    If Len(kulturaID) = 0 Then kulturaID = vrstaVoca & "-" & sortaVoca
    
    ' Neue ID
    newID = GetNextID(TBL_OTKUP, COL_OTK_ID, "OTK-")
    
    ' VozacID
    ' BrojDokumenta = "PWA:" & clientRecordID (für Duplikat-Check)
    ' Novac = 0, PrimalacNovca = ""
    
    Dim rowData As Variant
    rowData = Array(newID, datum, kooperantID, stanicaID, kulturaID, _
                    vrstaVoca, sortaVoca, kolicina, cena, tipAmb, _
                    kolAmb, vozacID, "", 0, "", klasa, _
                    "", "", "", "", "", parcelaID, _
                    clientRecordID, "PWA")
    
    Dim result As Long
    result = AppendRow(TBL_OTKUP, rowData)
    
    If result > 0 Then
        ' Ambalaza tracken
        If kolAmb > 0 Then
            TrackAmbalaza datum, tipAmb, kolAmb, "Izlaz", kooperantID, "Kooperant", , newID, DOK_TIP_OTKUP
        End If
        
        LogInfo "ImportRowToTblOtkup", "Importiert: " & newID & " ? PWA:" & clientRecordID & _
                " | " & kooperantID & " | " & vrstaVoca & " " & kolicina & "kg"
        ImportRowToTblOtkup = newID
    Else
        LogError "ImportRowToTblOtkup", "AppendRow fehlgeschlagen für PWA:" & clientRecordID
        ImportRowToTblOtkup = ""
    End If
    Exit Function

EH:
    LogErr "ImportRowToTblOtkup", "ClientRecordID: " & clientRecordID
    ImportRowToTblOtkup = ""
End Function

' ============================================================
' PRIVATE — SyncStatus zurückschreiben
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
               """values"":[[""" & JsonEscapeMasterSync(syncStatus) & """]]}"

        ' B = ServerRecordID
        If UBound(update) >= 2 Then
            serverRecordID = Trim$(CStr(update(2)))

            If Len(serverRecordID) > 0 Then
                body = body & ",{""range"":""Sheet1!B" & CStr(rowNum) & """," & _
                       """values"":[[""" & JsonEscapeMasterSync(serverRecordID) & """]]}"
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

Private Function JsonEscapeMasterSync(ByVal value As String) As String
    Dim result As String

    result = CStr(value)
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\n")
    result = Replace(result, vbLf, "\n")

    JsonEscapeMasterSync = result
End Function

Private Function TryUpdateVozacID(ByVal clientRecordID As String, _
                                   ByVal newVozacID As String) As Boolean
    ' Ako Otkup u masteru nema VozacID a sheet ga ima — updateuj
    Dim data As Variant
    data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function
    
    Dim colCRID As Long, colVoz As Long
    colCRID = GetColumnIndex(TBL_OTKUP, "ClientRecordID")
    colVoz = GetColumnIndex(TBL_OTKUP, COL_OTK_VOZAC)
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If CStr(Nz(data(i, colCRID), "")) = clientRecordID Then
            Dim currentVoz As String
            currentVoz = Trim$(CStr(Nz(data(i, colVoz), "")))
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
' PRIVATE — Helpers
' ============================================================

Private Function Nz(ByVal v As Variant, Optional ByVal Fallback As Variant = "") As Variant
    If IsError(v) Then
        Nz = Fallback
    ElseIf IsNull(v) Then
        Nz = Fallback
    ElseIf IsEmpty(v) Then
        Nz = Fallback
    ElseIf Trim$(CStr(v)) = "" Then
        Nz = Fallback
    Else
        Nz = v
    End If
End Function


' ============================================================
' modMasterSync — ZBIRNA IMPORT (dodati u postojeci modMasterSync)
' ============================================================


' ============================================================
' PUBLIC — Hauptfunktion Zbirna Import
' ============================================================

Public Sub ImportZbirneFromPWA()
    Dim folderID As String
    Dim sheetIDs As Collection
    Dim sheetNames As Collection
    Dim i As Long
    Dim totalImported As Long
    Dim totalSkipped As Long
    Dim totalErrors As Long
    
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
    
    LogInfo "ImportZbirneFromPWA", "Import gestartet"
    
    Set sheetIDs = New Collection
    Set sheetNames = New Collection
    Call FindVOZSheets(folderID, sheetIDs, sheetNames)
    
    If sheetIDs.count = 0 Then
        MsgBox "Nema VOZ-* fajlova u PWA folderu.", vbInformation, APP_NAME
        Exit Sub
    End If
    
    For i = 1 To sheetIDs.count
        Dim imported As Long, skipped As Long, errors As Long
        imported = 0: skipped = 0: errors = 0
        
        Call ImportOneVOZSheet(CStr(sheetIDs(i)), CStr(sheetNames(i)), imported, skipped, errors)
        
        totalImported = totalImported + imported
        totalSkipped = totalSkipped + skipped
        totalErrors = totalErrors + errors
    Next i
    
    LogInfo "ImportZbirneFromPWA", "Import abgeschlossen: " & totalImported & " importiert, " & _
            totalSkipped & " preskoceno, " & totalErrors & " greske aus " & sheetIDs.count & " fajlova"
    
    MsgBox "Uvoz zbirnih zavrsen!" & vbCrLf & vbCrLf & _
           "Fajlova: " & sheetIDs.count & vbCrLf & _
           "Uvezeno: " & totalImported & vbCrLf & _
           "Preskoceno: " & totalSkipped & vbCrLf & _
           "Greske: " & totalErrors, _
           vbInformation, APP_NAME
    Exit Sub

EH:
    LogErr "ImportZbirneFromPWA"
    MsgBox "Greska pri uvozu zbirnih: " & Err.description, vbCritical, APP_NAME
End Sub

Public Sub ImportZbirneFromPWA_TX()
    Dim tx As clsTransaction
    
    On Error GoTo EH
    
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_OTKUP
    
    Call ImportZbirneFromPWA
    
    tx.CommitTx

    Exit Sub

EH:
    LogErr "ImportZbirneFromPWA_TX"
    If Not tx Is Nothing Then tx.RollbackTx
    MsgBox "Greska pri uvozu zbirnih, promene vracene: " & Err.description, vbCritical, APP_NAME
End Sub

' ============================================================
' PRIVATE — Find VOZ-* Sheets in Folder
' ============================================================

Private Sub FindVOZSheets(ByVal folderID As String, _
                          ByRef outIDs As Collection, _
                          ByRef outNames As Collection)
    Dim accessToken As String
    Dim url As String
    Dim http As Object
    Dim query As String
    Dim responseText As String
    
    accessToken = GetAccessToken()
    If Len(accessToken) = 0 Then Exit Sub
    
    query = "name contains 'VOZ-' and mimeType='application/vnd.google-apps.spreadsheet'" & _
            " and '" & folderID & "' in parents and trashed=false"
    
    url = "https://www.googleapis.com/drive/v3/files" & _
          "?q=" & UrlEncode(query) & _
          "&fields=files(id,name)" & _
          "&pageSize=100"
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 10000, 10000, 30000, 30000
    
    http.Open "GET", url, False
    http.SetRequestHeader "Authorization", "Bearer " & accessToken
    http.Send
    
    If http.status <> 200 Then
        LogError "FindVOZSheets", "HTTP " & http.status & ": " & http.responseText, http.status
        Exit Sub
    End If
    
    responseText = http.responseText
    
    ' Reuse ParseFileList ali filtrira VOZ- umesto OTK-
    Call ParseFileListVOZ(responseText, outIDs, outNames)
    
    LogInfo "FindVOZSheets", "Gefunden: " & outIDs.count & " VOZ-Sheets"
End Sub
    

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
' PRIVATE — Import eines einzelnen VOZ-Sheets
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
                    newZbirnaID = ImportRowToTblZbirna(data, i, clientRecordID)
                    If Len(newZbirnaID) > 0 Then
                        ' Kaskadno povezivanje: Zbirna -> Otpremnice -> Otkupi
                        Dim brojZbirne As String
                        brojZbirne = GetBrojZbirneForID(newZbirnaID)
                        If Len(brojZbirne) > 0 Then
                            Dim otkupRecordIDs As String
                            otkupRecordIDs = Trim$(CStr(Nz(data(i, VS_OTKUP_RECORD_IDS), "")))
                            Call LinkZbirnaToOtkupAndOtpremnica(brojZbirne, otkupRecordIDs)
                        End If
                        
                        statusUpdates.Add Array(i, SYNC_STATUS_MASTER, newZbirnaID, brojZbirne)
                        outImported = outImported + 1
                    Else
                        statusUpdates.Add Array(i, SYNC_STATUS_ERROR & ":AppendRow failed", "")
                        outErrors = outErrors + 1
                    End If
                End If
            End If
        Else
            outSkipped = outSkipped + 1
        End If
    Next i
    
    If statusUpdates.count > 0 Then
        Call WriteBackVOZSyncStatus(spreadsheetID, statusUpdates)
    End If
    
    LogInfo "ImportOneVOZSheet", sheetName & ": " & outImported & " importiert, " & _
            outSkipped & " preskoceno, " & outErrors & " greske"
    Exit Sub

EH:
    LogErr "ImportOneVOZSheet", "Sheet: " & sheetName
    outErrors = outErrors + 1
End Sub

' ============================================================
' PRIVATE — Validierung
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
        If CStr(Nz(data(i, colCRID), "")) = clientRecordID Then
            IsDuplicateZbirnaInMaster = True
            Exit Function
        End If
    Next i
    
    IsDuplicateZbirnaInMaster = False
End Function

' ============================================================
' PRIVATE — Import Row to tblZbirna
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
    tipAmb = Trim$(CStr(Nz(data(row, VS_TIP_AMB), "")))
    
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
    On Error GoTo EH
    
    ukupnoKol = kolKlI + kolKlII
    
    If Len(tipAmb) = 0 Then tipAmb = "12/1"
    
    brojZbirne = GenerateBrojZbirne(vozacID, datum)
    If Len(brojZbirne) = 0 Then
            LogError "ImportRowToTblZbirna", "Nije moguce generisati BrojZbirne za VozacID=" & vozacID
            ImportRowToTblZbirna = ""
            Exit Function
    End If
    ' Hladnjaca iz KupacID
    Dim hladnjaca As String
    hladnjaca = CStr(Nz(LookupValue(TBL_KUPCI, "KupacID", kupacID, "Hladnjaca"), ""))
    
    newID = GetNextID(TBL_ZBIRNA, COL_ZBR_ID, "ZBR-")
    
    ' tblZbirna Schema:
    ' ZbirnaID | BrojZbirne | Datum | VozacID | KupacID | Hladnjaca | Pogon |
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
        LogInfo "ImportRowToTblZbirna", "Importiert: " & newID & " BrojZbirne=" & brojZbirne & _
                " | " & vozacID & " | " & kupacID & " | " & ukupnoKol & "kg"
        ImportRowToTblZbirna = newID
    Else
        LogError "ImportRowToTblZbirna", "AppendRow fehlgeschlagen für PWA:" & clientRecordID
        ImportRowToTblZbirna = ""
    End If
    Exit Function

EH:
    LogErr "ImportRowToTblZbirna", "ClientRecordID: " & clientRecordID
    ImportRowToTblZbirna = ""
End Function

' ============================================================
' PRIVATE — Kaskadno povezivanje Zbirna -> Otpremnice -> Otkupi
' ============================================================

Private Sub LinkZbirnaToOtkupAndOtpremnica(ByVal brojZbirne As String, _
                                            ByVal otkupRecordIDs As String)
    ' otkupRecordIDs = comma-separated ClientRecordIDs iz VOZ sheeta
    ' Flow:
    '   1. Za svaki ClientRecordID -> nadji OtkupID u tblOtkup
    '   2. Postavi tblOtkup.BrojZbirne = brojZbirne
    '   3. Iz tog otkupa citaj OtpremnicaID
    '   4. Postavi tblOtpremnica.BrojZbirne = brojZbirne (ako vec nije)
    
    If Len(Trim$(otkupRecordIDs)) = 0 Then Exit Sub
    
    Dim crIDs() As String
    crIDs = Split(otkupRecordIDs, ",")
    
    ' Ucitaj tblOtkup jednom
    Dim otkData As Variant
    otkData = GetTableData(TBL_OTKUP)
    If IsEmpty(otkData) Then Exit Sub
    
    Dim colCRID As Long, colOtkID As Long, colOtkBrZbr As Long, colOtkOtpID As Long
    colCRID = GetColumnIndex(TBL_OTKUP, "ClientRecordID")
    colOtkID = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    colOtkBrZbr = GetColumnIndex(TBL_OTKUP, COL_OTK_BROJ_ZBIRNE)
    colOtkOtpID = GetColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID)
    
    ' Dict za Otpremnice koje smo vec updateovali (da ne radimo duple)
    Dim updatedOtp As Object
    Set updatedOtp = CreateObject("Scripting.Dictionary")
    
    Dim c As Long
    For c = 0 To UBound(crIDs)
        Dim searchCRID As String
        searchCRID = Trim$(crIDs(c))
        If Len(searchCRID) = 0 Then GoTo NextCRID
        
        ' Nadji red u tblOtkup po ClientRecordID
        Dim i As Long
        For i = 1 To UBound(otkData, 1)
            If CStr(Nz(otkData(i, colCRID), "")) = searchCRID Then
                
                ' 1. Postavi BrojZbirne na otkupu
                Dim otkupID As String
                otkupID = CStr(otkData(i, colOtkID))
                
                Dim otkRows As Collection
                Set otkRows = FindRows(TBL_OTKUP, COL_OTK_ID, otkupID)
                If otkRows.count > 0 Then
                    UpdateCell TBL_OTKUP, otkRows(1), COL_OTK_BROJ_ZBIRNE, brojZbirne
                End If
                
                ' 2. Postavi BrojZbirne na otpremnici (ako postoji i ako vec nije)
                Dim otpID As String
                otpID = Trim$(CStr(Nz(otkData(i, colOtkOtpID), "")))
                
                If Len(otpID) > 0 Then
                    If Not updatedOtp.Exists(otpID) Then
                        Dim otpRows As Collection
                        Set otpRows = FindRows(TBL_OTPREMNICA, COL_OTP_ID, otpID)
                        If otpRows.count > 0 Then
                            UpdateCell TBL_OTPREMNICA, otpRows(1), COL_OTP_BROJ_ZBIRNE, brojZbirne
                        End If
                        updatedOtp.Add otpID, True
                    End If
                End If
                
                Exit For  ' Nasli smo otkup, idemo na sledeci CRID
            End If
        Next i
NextCRID:
    Next c
    
    LogInfo "LinkZbirnaToOtkupAndOtpremnica", "BrojZbirne=" & brojZbirne & _
            " linked " & (UBound(crIDs) + 1) & " otkupa, " & updatedOtp.count & " otpremnica"
End Sub

' ============================================================
' PRIVATE — Helper: BrojZbirne aus ZbirnaID
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

' ============================================================
' PRIVATE — WriteBack VOZ SyncStatus + ServerRecordID
' ============================================================

Private Sub WriteBackVOZSyncStatus(ByVal spreadsheetID As String, _
                                    ByVal updates As Collection)
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
    If Len(accessToken) = 0 Then Exit Sub
    
    body = "{""valueInputOption"":""RAW"",""data"":["
    
    Dim isFirst As Boolean
    isFirst = True
    
    For i = 1 To updates.count
        update = updates(i)
        
        If Not isFirst Then body = body & ","
        isFirst = False
        
        ' Kolona F — SyncStatus
        body = body & "{""range"":""Sheet1!F" & CStr(update(0)) & """," & _
               """values"":[[""" & JsonEscape(CStr(update(1))) & """]]}"
        
        ' Kolona B — ServerRecordID (2. kolona = B)
        If UBound(update) >= 2 Then
            If Len(CStr(update(2))) > 0 Then
                body = body & ",{""range"":""Sheet1!B" & CStr(update(0)) & """," & _
                       """values"":[[""" & JsonEscape(CStr(update(2))) & """]]}"
            End If
        End If
        
        If UBound(update) >= 2 Then
            If Len(CStr(update(2))) > 0 Then
                body = body & ",{""range"":""Sheet1!T" & CStr(update(0)) & """," & _
                       """values"":[[""" & JsonEscape(CStr(update(2))) & """]]}"
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
        LogInfo "WriteBackVOZSyncStatus", updates.count & " Status-Updates geschrieben"
    Else
        LogError "WriteBackVOZSyncStatus", "HTTP " & http.status & ": " & http.responseText, http.status
    End If
    Exit Sub

EH:
    LogErr "WriteBackVOZSyncStatus"
End Sub

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

' ============================================================
' TEST
' ============================================================

Public Sub Test_ImportOtkupFromPWA()
    Call ImportOtkupFromPWA
End Sub


