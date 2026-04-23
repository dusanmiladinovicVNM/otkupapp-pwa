Attribute VB_Name = "modDataAccess"
Option Explicit

' ============================================================
' modDataAccess – EINZIGER Ort für Sheet/Table-Zugriffe
' Alles andere ruft NUR diese Funktionen auf.
' Daten rein als Arrays, Daten raus als Arrays.
' ============================================================

' --- Table Access ---

Public Function GetTable(ByVal tblName As String) As ListObject
    ' Findet ein ListObject über alle Sheets hinweg
    Dim ws As Worksheet
    Dim lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = tblName Then
                Set GetTable = lo
                Exit Function
            End If
        Next lo
    Next ws
    Set GetTable = Nothing
End Function

Public Function GetTableData(ByVal tblName As String) As Variant
    ' Gibt den gesamten Datenbereich einer Tabelle als 2D-Array zurück
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
    GetTableData = lo.DataBodyRange.Value
End Function

Public Function GetTableHeaders(ByVal tblName As String) As Variant
    ' Gibt die Spaltennamen als 1D-Array zurück
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
        headers(i) = lo.ListColumns(i).Name
    Next i
    GetTableHeaders = headers
End Function

Public Function GetColumnIndex(ByVal tblName As String, ByVal colName As String) As Long
    ' Gibt den Spaltenindex innerhalb der Tabelle zurück (1-basiert)
    Dim lo As ListObject
    Set lo = GetTable(tblName)
    If lo Is Nothing Then
        GetColumnIndex = 0
        Exit Function
    End If
    On Error Resume Next
    GetColumnIndex = lo.ListColumns(colName).Index
    On Error GoTo 0
End Function

Public Function GetColumnData(ByVal tblName As String, ByVal colName As String) As Variant
    ' Gibt eine einzelne Spalte als 1D-Array zurück
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
        arr(1) = rng.Value
        GetColumnData = arr
    Else
        GetColumnData = Application.Transpose(rng.Value)
    End If
End Function

' --- WRITE Operations ---

Public Function AppendRow(ByVal tblName As String, ByVal rowData As Variant) As Long
    ' Fügt eine neue Zeile an die Tabelle an
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
        newRow.Range.cells(1, i - LBound(rowData) + 1).Value = rowData(i)
    Next i
    
    AppendRow = newRow.Index
    WriteJournalRow tblName, rowData
    Exit Function
    
ErrHandler:
    Debug.Print "AppendRow ERROR: "; Err.Number; Err.Description
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
    lo.DataBodyRange.cells(rowIndex, colIdx).Value = newValue
    UpdateCell = True
    Exit Function
    
ErrHandler:
    UpdateCell = False
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
    ' Generiert die nächste ID (z.B. "OTK-00001")
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
    ' Prüft ob searchValue in colName bereits existiert
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
        ' Stornierte überspringen
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
            CheckDuplicate = "Dokument '" & searchValue & "' vec postoji! " & _
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
    ' Einfacher VLOOKUP-Ersatz über ListObjects
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

Public Function GetLookupList(ByVal tblName As String, ByVal colName As String, _
                              Optional ByVal filterCol As String = "", _
                              Optional ByVal filterVal As Variant) As Variant
    ' Gibt eindeutige Werte einer Spalte als Array zurück (für ComboBox-Füllung)
    Dim data As Variant
    data = GetTableData(tblName)
    If IsEmpty(data) Then
        GetLookupList = Array()
        Exit Function
    End If
    
    Dim colIdx As Long, filterIdx As Long
    colIdx = GetColumnIndex(tblName, colName)
    If filterCol <> "" Then filterIdx = GetColumnIndex(tblName, filterCol)
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim val As Variant
    For i = 1 To UBound(data, 1)
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
    Next i
    
    GetLookupList = dict.keys
End Function

