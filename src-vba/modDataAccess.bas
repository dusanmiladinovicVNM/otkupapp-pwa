Attribute VB_Name = "modDataAccess"
Option Explicit

' ============================================================
' modDataAccess - EINZIGER Ort fuer Sheet/Table-Zugriffe
' Alles andere ruft NUR diese Funktionen auf.
' Daten rein als Arrays, Daten raus als Arrays.
' ============================================================

' --- Request-scoped cache za GetTableData ---
' Aktivira se SAMO oko read-only blokova (npr. generisanje izvestaja, preko
' BeginTableCache/EndTableCache). Dok je aktivan, svaka tabela se cita JEDNOM
' pa se vraca kesirana kopija -- jedan "Prikazi" je citao 17-18 puta iste
' velike tabele (TBL_OTKUP 5x, TBL_AMBALAZA 4x, TBL_NOVAC 3x...).
' AppendRow/UpdateCell invalidiraju kes za tu tabelu (sigurnosna mreza ako se
' ipak pise dok je prozor aktivan). Depth-brojac podrzava ugnjezdene Begin/End.
Private mTableCache As Object
Private mTableCacheDepth As Long
Private mExclCache As Object        ' kes ExcludeStornirano rezultata (po tblName), isti prozor
Private mColCache As Object         ' kes GetColumnIndex ("tbl|col" -> index), isti prozor

Public Sub BeginTableCache()
    If mTableCacheDepth = 0 Then
        Set mTableCache = CreateObject("Scripting.Dictionary")
        Set mExclCache = CreateObject("Scripting.Dictionary")
        Set mColCache = CreateObject("Scripting.Dictionary")
    End If
    mTableCacheDepth = mTableCacheDepth + 1
End Sub

Public Sub EndTableCache()
    If mTableCacheDepth > 0 Then mTableCacheDepth = mTableCacheDepth - 1
    If mTableCacheDepth <= 0 Then
        mTableCacheDepth = 0
        Set mTableCache = Nothing
        Set mExclCache = Nothing
        Set mColCache = Nothing
    End If
End Sub

Private Sub InvalidateTableCache(ByVal tblName As String)
    If Not mTableCache Is Nothing Then
        If mTableCache.Exists(tblName) Then mTableCache.Remove tblName
    End If
    If Not mExclCache Is Nothing Then
        If mExclCache.Exists(tblName) Then mExclCache.Remove tblName
    End If
End Sub

' --- ExcludeStornirano kes (koristi modHelpers.ExcludeStornirano) ---
' Kesira se SAMO kad je ulaz PUNA tabela (broj redova = GetTableData), pa
' filtrirani podskupovi (manje redova) nikad ne pogode kes pune tabele.
Public Function ExclCacheActive() As Boolean
    ExclCacheActive = Not (mExclCache Is Nothing)
End Function

Public Function ExclCacheTryGet(ByVal tblName As String, ByRef outData As Variant) As Boolean
    ExclCacheTryGet = False
    If mExclCache Is Nothing Then Exit Function
    If mExclCache.Exists(tblName) Then
        outData = mExclCache(tblName)
        ExclCacheTryGet = True
    End If
End Function

Public Sub ExclCachePut(ByVal tblName As String, ByVal data As Variant)
    If mExclCache Is Nothing Then Exit Sub
    mExclCache(tblName) = data
End Sub

' --- Table Access ---

Public Function GetTable(ByVal tblName As String) As ListObject
    ' Findet ein ListObject ueber alle Sheets hinweg
    Dim ws As Worksheet
    Dim lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If lo.name = tblName Then
                Set GetTable = lo
                Exit Function
            End If
        Next lo
    Next ws
    Set GetTable = Nothing
End Function

Public Function GetTableData(ByVal tblName As String) As Variant
    ' Gibt den gesamten Datenbereich einer Tabelle als 2D-Array zurueck
    If Not mTableCache Is Nothing Then
        If mTableCache.Exists(tblName) Then
            GetTableData = mTableCache(tblName)
            Exit Function
        End If
    End If
    Dim lo As ListObject
    Set lo = GetTable(tblName)
    If lo Is Nothing Then
        GetTableData = Empty
        Exit Function
    End If
    If lo.DataBodyRange Is Nothing Then
        GetTableData = Empty
        Exit Function
    End If
    GetTableData = lo.DataBodyRange.value
    If Not mTableCache Is Nothing Then mTableCache(tblName) = GetTableData
End Function

Public Function GetTableHeaders(ByVal tblName As String) As Variant
    ' Gibt die Spaltennamen als 1D-Array zurueck
    Dim lo As ListObject
    Set lo = GetTable(tblName)
    If lo Is Nothing Then
        GetTableHeaders = Empty
        Exit Function
    End If
    Dim headers() As String
    Dim i As Long
    ReDim headers(1 To lo.ListColumns.count)
    For i = 1 To lo.ListColumns.count
        headers(i) = lo.ListColumns(i).name
    Next i
    GetTableHeaders = headers
End Function

' TEST SEAM: podmetni vrednost u kes indeksa kolone. Tvrdo gejtovan.
'
' Postoji da bi se mogao izmeriti BAS onaj simptom zbog kojeg je poruka i
' prosirena: trazenje vrati nulu za kolonu koja u zaglavlju POSTOJI. Kes se u
' pogonu puni sam (BeginTableCache prozor), pa se to stanje bez seam-a ne moze
' izazvati -- a upravo se ono videlo u logu nad radnom sveskom.
'
' Ovo NE tvrdi da je kes uzrok prvog neuspeha. Tvrdi samo da nula, jednom
' zapamcena, prezivi ceo prozor -- i da poruka to sada ume da razlikuje.
Public Sub KesKoloneTestSet(ByVal tblName As String, ByVal colName As String, _
                            ByVal vrednost As Long)
    If Not IsTestMode() Then Exit Sub
    If mColCache Is Nothing Then Exit Sub
    mColCache(tblName & "|" & colName) = vrednost
End Sub

' TEST SEAM: da li je kljuc u kesu indeksa kolone. Isti razlog kao
' modUiData.KesImaKljuc: pravilo "nula se ne pamti" se ne moze izmeriti kroz
' vracenu vrednost, jer GetColumnIndex daje nulu i kad je zapamtio i kad nije.
Public Function KesKoloneImaKljuc(ByVal tblName As String, _
                                  ByVal colName As String) As Boolean
    If mColCache Is Nothing Then Exit Function
    KesKoloneImaKljuc = mColCache.Exists(tblName & "|" & colName)
End Function

Public Function GetColumnIndex(ByVal tblName As String, ByVal colName As String) As Long
    ' Gibt den Spaltenindex innerhalb der Tabelle zurueck (1-basiert)
    ' Request-scoped kes: u jednom "Prikazi" prozoru kolone se ne menjaju, pa
    ' izbegavamo ~80 GetTable skenova (svaki ide kroz sve sheetove/tabele).
    Dim ck As String
    If Not mColCache Is Nothing Then
        ck = tblName & "|" & colName
        If mColCache.Exists(ck) Then
            GetColumnIndex = mColCache(ck)
            Exit Function
        End If
    End If
    Dim lo As ListObject
    Set lo = GetTable(tblName)
    If lo Is Nothing Then
        GetColumnIndex = 0
        Exit Function
    End If
    On Error Resume Next
    GetColumnIndex = lo.ListColumns(colName).index
    On Error GoTo 0
    ' NULA NIJE ZNANJE, pa se ne pamti.
    '
    ' Jedan trenutan neuspeh je inace postajao TRAJAN: zapamcena nula vazi za
    ' ceo BeginTableCache prozor, pa svaki sledeci poziv nad istom kolonom
    ' dobija istu nulu bez ijednog novog pokusaja -- a fail-closed kapije
    ' (RequireColumnIndex) na to staju. Tako je nastalo "Nedostaje kolona
    ' 'VozacID' u tabeli 'tblZbirna'" nad sveskom u kojoj ta kolona POSTOJI.
    '
    ' Isto pravilo vec drzi kes TABELA (modUiData.CachedTable kesira samo kad
    ' je citanje uspelo). Ovde je bila rupa, na istom mestu i iz istog razloga.
    '
    ' Ovo NE popravlja uzrok PRVOG neuspeha -- on nije reprodukovan (postmortem
    ' par 11) -- nego mu skida trajnost. Test 123.
    If Not mColCache Is Nothing Then
        If GetColumnIndex > 0 Then mColCache(ck) = GetColumnIndex
    End If
End Function

Public Function GetColumnData(ByVal tblName As String, ByVal colName As String) As Variant
    ' Gibt eine einzelne Spalte als 1D-Array zurueck
    Dim lo As ListObject
    Set lo = GetTable(tblName)
    If lo Is Nothing Then
        GetColumnData = Empty
        Exit Function
    End If
    If lo.DataBodyRange Is Nothing Then
        GetColumnData = Empty
        Exit Function
    End If
    
    Dim colIdx As Long
    colIdx = GetColumnIndex(tblName, colName)
    If colIdx = 0 Then
        GetColumnData = Empty
        Exit Function
    End If
    
    Dim rng As Range
    Set rng = lo.ListColumns(colIdx).DataBodyRange
    If rng.rows.count = 1 Then
        Dim arr(1 To 1) As Variant
        arr(1) = rng.value
        GetColumnData = arr
    Else
        GetColumnData = Application.Transpose(rng.value)
    End If
End Function

' --- WRITE Operations ---

Public Function AppendRow(ByVal tblName As String, ByVal rowData As Variant) As Long
    ' Fuegt eine neue Zeile an die Tabelle an
    ' rowData = 1D Array mit Werten in Spaltenreihenfolge
    ' Returns: Zeilennummer der neuen Zeile (0 bei Fehler)
    Dim lo As ListObject
    Set lo = GetTable(tblName)
    If lo Is Nothing Then
        AppendRow = 0
        Exit Function
    End If
    
    On Error GoTo ErrHandler
    
    Dim newRow As ListRow
    Set newRow = lo.ListRows.Add
    
    Dim i As Long
    Dim colCount As Long
    colCount = lo.ListColumns.count
    
    For i = LBound(rowData) To Application.Min(UBound(rowData), LBound(rowData) + colCount - 1)
        newRow.Range.cells(1, i - LBound(rowData) + 1).value = rowData(i)
    Next i
    
    AppendRow = newRow.index
    StampRowAudit lo, newRow.index, True
    WriteJournalRow tblName, rowData
    InvalidateTableCache tblName
    ' KPI sidebar (frmOtkupAPP) zavisi samo od ovih tabela -> oznaci za osvezavanje
    ' pri sledecem povratku na dashboard (umesto racunanja pri svakom Activate).
    If tblName = TBL_OTKUP Or tblName = TBL_OTPREMNICA Or tblName = TBL_PRIJEMNICA Then
        gKpiDirty = True
    End If
    Exit Function
    
ErrHandler:
    Debug.Print "AppendRow ERROR: "; Err.Number; Err.description
    AppendRow = 0
End Function

Public Function UpdateCell(ByVal tblName As String, ByVal rowIndex As Long, _
                           ByVal colName As String, ByVal newValue As Variant) As Boolean
    ' Aktualisiert eine einzelne Zelle
    Dim lo As ListObject
    Set lo = GetTable(tblName)
    If lo Is Nothing Then
        UpdateCell = False
        Exit Function
    End If
    
    Dim colIdx As Long
    colIdx = GetColumnIndex(tblName, colName)
    If colIdx = 0 Then
        UpdateCell = False
        Exit Function
    End If
    
    On Error GoTo ErrHandler
    lo.DataBodyRange.cells(rowIndex, colIdx).value = newValue
    StampRowAudit lo, rowIndex, False
    UpdateCell = True
    InvalidateTableCache tblName
    Exit Function

ErrHandler:
    UpdateCell = False
End Function

' ============================================================
' Audit stamp (timestamp + userstamp) - centralno za AppendRow/UpdateCell.
' Upisuje direktno u celije (ne preko UpdateCell) -> nema rekurzije.
' No-op ako tabela nema audit kolone (npr. config/report tabele).
' Userstamp: prijavljeni korisnik (modAuth) -> fallback Windows nalog.
' ============================================================
Private Sub StampRowAudit(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal isInsert As Boolean)
    On Error Resume Next

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If rowIndex < 1 Then Exit Sub

    Dim cCA As Long, cCB As Long, cMA As Long, cMB As Long
    cCA = LoColIdx(lo, COL_AUDIT_CREATED_AT)
    cCB = LoColIdx(lo, COL_AUDIT_CREATED_BY)
    cMA = LoColIdx(lo, COL_AUDIT_MODIFIED_AT)
    cMB = LoColIdx(lo, COL_AUDIT_MODIFIED_BY)
    If cCA = 0 And cCB = 0 And cMA = 0 And cMB = 0 Then Exit Sub

    Dim ts As String, usr As String
    ts = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    usr = AuditUser()

    ' Created* samo na insert i samo ako je prazno (ne gazi eksplicitnu vrednost).
    If isInsert Then
        If cCA > 0 Then
            If Len(Trim$(CStr(lo.DataBodyRange.cells(rowIndex, cCA).value))) = 0 Then _
                lo.DataBodyRange.cells(rowIndex, cCA).value = ts
        End If
        If cCB > 0 Then
            If Len(Trim$(CStr(lo.DataBodyRange.cells(rowIndex, cCB).value))) = 0 Then _
                lo.DataBodyRange.cells(rowIndex, cCB).value = usr
        End If
    End If

    ' Modified* uvek (i na insert i na izmenu).
    If cMA > 0 Then lo.DataBodyRange.cells(rowIndex, cMA).value = ts
    If cMB > 0 Then lo.DataBodyRange.cells(rowIndex, cMB).value = usr
End Sub

' Indeks kolone po imenu UNUTAR datog ListObject-a (bez ponovnog trazenja tabele).
Private Function LoColIdx(ByVal lo As ListObject, ByVal colName As String) As Long
    On Error Resume Next
    Dim i As Long
    For i = 1 To lo.ListColumns.count
        If StrComp(Trim$(lo.ListColumns(i).name), Trim$(colName), vbTextCompare) = 0 Then
            LoColIdx = i
            Exit Function
        End If
    Next i
End Function

' Userstamp: prijavljeni app-korisnik (kad je AUTH ukljucen) -> Windows nalog -> "Operator".
Private Function AuditUser() As String
    On Error Resume Next
    Dim u As String
    u = modAuth.GetCurrentUser()
    If Len(Trim$(u)) = 0 Then u = Environ$("USERNAME")
    If Len(Trim$(u)) = 0 Then u = "Operator"
    AuditUser = u
End Function

Public Function FindRows(ByVal tblName As String, ByVal colName As String, _
                         ByVal searchValue As Variant) As Collection
    ' Findet alle Zeilenindizes wo colName = searchValue
    Dim result As New Collection
    Dim data As Variant
    Dim colIdx As Long
    
    colIdx = GetColumnIndex(tblName, colName)
    If colIdx = 0 Then
        Set FindRows = result
        Exit Function
    End If
    
    data = GetTableData(tblName)
    If IsEmpty(data) Then
        Set FindRows = result
        Exit Function
    End If
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If data(i, colIdx) = searchValue Then
            result.Add i
        End If
    Next i
    
    Set FindRows = result
End Function

Public Function GetNextID(ByVal tblName As String, ByVal idColName As String, _
                          Optional ByVal prefix As String = "") As String
    ' Generiert die naechste ID (z.B. "OTK-00001")
    Dim data As Variant
    Dim colIdx As Long
    colIdx = GetColumnIndex(tblName, idColName)
    data = GetTableData(tblName)

    Dim maxNum As Long
    maxNum = 0

    If Not IsEmpty(data) Then
        Dim i As Long
        Dim numPart As Long
        For i = 1 To UBound(data, 1)
            If Not IsEmpty(data(i, colIdx)) Then
                If prefix <> "" Then
                    If Left$(CStr(data(i, colIdx)), Len(prefix)) = prefix Then
                        On Error Resume Next
                        numPart = CLng(Mid$(CStr(data(i, colIdx)), Len(prefix) + 1))
                        On Error GoTo 0
                        If numPart > maxNum Then maxNum = numPart
                    End If
                Else
                    On Error Resume Next
                    numPart = CLng(data(i, colIdx))
                    On Error GoTo 0
                    If numPart > maxNum Then maxNum = numPart
                End If
            End If
        Next i
    End If

    If prefix <> "" Then
        GetNextID = prefix & Format$(maxNum + 1, "00000")
    Else
        GetNextID = CStr(maxNum + 1)
    End If
End Function

Public Function CheckDuplicate(ByVal tblName As String, ByVal colName As String, _
                               ByVal searchValue As String, _
                               ByVal datumColName As String) As String
    ' Prueft ob searchValue in colName bereits existiert
    ' Returns: "" wenn kein Duplikat, sonst Fehlermeldung mit Datum
    
    If searchValue = "" Then
        CheckDuplicate = ""
        Exit Function
    End If
    
    Dim data As Variant
    data = GetTableData(tblName)
    If IsEmpty(data) Then
        CheckDuplicate = ""
        Exit Function
    End If
    
    Dim colIdx As Long, datumIdx As Long, colStorno As Long
    colIdx = GetColumnIndex(tblName, colName)
    datumIdx = GetColumnIndex(tblName, datumColName)
    colStorno = GetColumnIndex(tblName, COL_STORNIRANO)
    
    If colIdx = 0 Then
        CheckDuplicate = ""
        Exit Function
    End If
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        ' Stornierte ueberspringen
        If colStorno > 0 Then
            If CStr(data(i, colStorno)) = "Da" Then GoTo NextRow
        End If
            
        If CStr(data(i, colIdx)) = searchValue Then
            Dim datumStr As String
            If datumIdx > 0 And IsDate(data(i, datumIdx)) Then
                datumStr = Format$(CDate(data(i, datumIdx)), "d.m.yyyy")
            Else
                datumStr = "(nepoznat datum)"
            End If
            CheckDuplicate = "Dokument '" & searchValue & "' ve" & ChrW(263) & " postoji! " & _
                             "Unet je " & datumStr & "."
            Exit Function
        End If
NextRow:
    Next i
    
    CheckDuplicate = ""
End Function


' --- Lookup Helpers ---

Public Function LookupValue(ByVal tblName As String, ByVal searchCol As String, _
                            ByVal searchVal As Variant, ByVal returnCol As String) As Variant
    ' Einfacher VLOOKUP-Ersatz ueber ListObjects
    Dim data As Variant
    data = GetTableData(tblName)
    If IsEmpty(data) Then
        LookupValue = Empty
        Exit Function
    End If
    
    Dim searchIdx As Long, returnIdx As Long
    searchIdx = GetColumnIndex(tblName, searchCol)
    returnIdx = GetColumnIndex(tblName, returnCol)
    
    If searchIdx = 0 Or returnIdx = 0 Then
        LookupValue = Empty
        Exit Function
    End If
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If CStr(data(i, searchIdx)) = CStr(searchVal) Then
            LookupValue = data(i, returnIdx)
            Exit Function
        End If
    Next i
    
    LookupValue = Empty
End Function

Public Function BuildLookupDict(ByVal tblName As String, ByVal keyCol As String, _
                                ByVal valCol As String, _
                                Optional ByVal valCol2 As String = "") As Object
    ' Jednoprolazna mapa keyCol -> valCol (opc. + valCol2 spojeno razmakom).
    ' Zamena za LookupValue pozvan U PETLJI: gradi se O(n) JEDNOM, pa je svaki
    ' lookup O(1) -- umesto O(n) po redu (sto je ukupno bilo O(n*m)).
    ' Kljucevi su CStr; prvi pojav pobedjuje (isto kao LookupValue -> prvi match).
    ' Vrednost je CStr (pozivaoci ionako rade CStr(LookupValue(...))).
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set BuildLookupDict = d

    Dim data As Variant
    data = GetTableData(tblName)
    If IsEmpty(data) Then Exit Function

    Dim ki As Long, vi As Long, vi2 As Long
    ki = GetColumnIndex(tblName, keyCol)
    vi = GetColumnIndex(tblName, valCol)
    If ki = 0 Or vi = 0 Then Exit Function
    If valCol2 <> "" Then vi2 = GetColumnIndex(tblName, valCol2)

    Dim i As Long, k As String, v As String
    For i = 1 To UBound(data, 1)
        k = CStr(data(i, ki))
        If Len(k) > 0 Then
            If Not d.Exists(k) Then
                If vi2 > 0 Then
                    v = CStr(data(i, vi)) & " " & CStr(data(i, vi2))
                Else
                    v = CStr(data(i, vi))
                End If
                d.Add k, v
            End If
        End If
    Next i
End Function

Public Function GetLookupList(ByVal tblName As String, ByVal colName As String, _
                              Optional ByVal filterCol As String = "", _
                              Optional ByVal filterVal As Variant, _
                              Optional ByVal onlyActive As Boolean = False) As Variant
    ' Gibt eindeutige Werte einer Spalte als Array zurueck (fuer ComboBox-Fuellung).
    ' onlyActive: preskoci redove gde je Aktivan = "Neaktivan" (soft-delete).
    Dim data As Variant
    data = GetTableData(tblName)
    If IsEmpty(data) Then
        GetLookupList = Array()
        Exit Function
    End If

    Dim colIdx As Long, filterIdx As Long, aktIdx As Long
    colIdx = GetColumnIndex(tblName, colName)
    If filterCol <> "" Then filterIdx = GetColumnIndex(tblName, filterCol)
    If onlyActive Then aktIdx = GetColumnIndex(tblName, "Aktivan")

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim val As Variant
    For i = 1 To UBound(data, 1)
        Dim skipRow As Boolean
        skipRow = False
        If aktIdx > 0 Then
            If StrComp(Trim$(CStr(data(i, aktIdx))), STATUS_NEAKTIVAN, vbTextCompare) = 0 Then skipRow = True
        End If

        If Not skipRow Then
            val = data(i, colIdx)
            If Not IsEmpty(val) And CStr(val) <> "" Then
                If filterCol = "" Then
                    If Not dict.Exists(CStr(val)) Then dict.Add CStr(val), val
                Else
                    If data(i, filterIdx) = filterVal Then
                        If Not dict.Exists(CStr(val)) Then dict.Add CStr(val), val
                    End If
                End If
            End If
        End If
    Next i

    GetLookupList = dict.keys
End Function

