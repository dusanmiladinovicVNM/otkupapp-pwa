Attribute VB_Name = "modArrayUtils"
Option Explicit

' ============================================================
' modArrayUtils – Array-Operationen im Speicher
' Ersetzt ALLE Copy/Paste/Sort Operationen auf Sheets
' ============================================================

Public Function FilterArray(ByVal data As Variant, ByVal filters As Collection) As Variant
    ' Filtert ein 2D-Array nach mehreren Kriterien
    ' filters = Collection of clsFilterParam
    ' Returns: Gefiltertes 2D-Array (oder Empty)
    
    If IsEmpty(data) Then
        FilterArray = Empty
        Exit Function
    End If
    
    Dim rowCount As Long, colCount As Long
    rowCount = UBound(data, 1)
    colCount = UBound(data, 2)
    
    ' Erst zählen wie viele Zeilen matchen
    Dim matchCount As Long
    Dim matches() As Boolean
    ReDim matches(1 To rowCount)
    
    Dim i As Long, j As Long
    Dim fp As clsFilterParam
    
    For i = 1 To rowCount
        matches(i) = True
        For Each fp In filters
            If Not MatchesFilter(data(i, fp.ColIndex), fp) Then
                matches(i) = False
                Exit For
            End If
        Next fp
        If matches(i) Then matchCount = matchCount + 1
    Next i
    
    If matchCount = 0 Then
        FilterArray = Empty
        Exit Function
    End If
    
    ' Ergebnis-Array bauen
    Dim result() As Variant
    ReDim result(1 To matchCount, 1 To colCount)
    Dim outRow As Long
    outRow = 0
    
    For i = 1 To rowCount
        If matches(i) Then
            outRow = outRow + 1
            For j = 1 To colCount
                result(outRow, j) = data(i, j)
            Next j
        End If
    Next i
    
    FilterArray = result
End Function

Private Function MatchesFilter(ByVal cellValue As Variant, ByVal fp As clsFilterParam) As Boolean
    ' Prüft ob ein Wert dem Filter entspricht
    Select Case fp.Operator
        Case "="
            MatchesFilter = (CStr(cellValue) = CStr(fp.Value1))
        Case ">="
            If IsDate(cellValue) And IsDate(fp.Value1) Then
                MatchesFilter = (CDate(cellValue) >= CDate(fp.Value1))
            Else
                MatchesFilter = (CDbl(cellValue) >= CDbl(fp.Value1))
            End If
        Case "<="
            If IsDate(cellValue) And IsDate(fp.Value1) Then
                MatchesFilter = (CDate(cellValue) <= CDate(fp.Value1))
            Else
                MatchesFilter = (CDbl(cellValue) <= CDbl(fp.Value1))
            End If
        Case "BETWEEN"
            If IsDate(cellValue) Then
                MatchesFilter = (CDate(cellValue) >= CDate(fp.Value1) And _
                                 CDate(cellValue) <= CDate(fp.Value2))
            Else
                MatchesFilter = (CDbl(cellValue) >= CDbl(fp.Value1) And _
                                 CDbl(cellValue) <= CDbl(fp.Value2))
            End If
        Case "<>"
            MatchesFilter = (CStr(cellValue) <> CStr(fp.Value1))
        Case "LIKE"
            MatchesFilter = (InStr(1, CStr(cellValue), CStr(fp.Value1), vbTextCompare) > 0)
        Case Else
            MatchesFilter = True
    End Select
End Function

Public Function SortArray(ByVal data As Variant, ByVal sortCol As Long, _
                          Optional ByVal ascending As Boolean = True, _
                          Optional ByVal sortCol2 As Long = 0) As Variant
    ' QuickSort auf 2D-Array (im Speicher, kein Sheet-Zugriff)
    If IsEmpty(data) Then
        SortArray = Empty
        Exit Function
    End If
    
    Dim rowCount As Long, colCount As Long
    rowCount = UBound(data, 1)
    colCount = UBound(data, 2)
    
    ' Index-Array für Sortierung
    Dim idx() As Long
    ReDim idx(1 To rowCount)
    Dim i As Long
    For i = 1 To rowCount
        idx(i) = i
    Next i
    
    ' QuickSort auf Index-Array
    QuickSortIndex idx, data, sortCol, 1, rowCount, ascending, sortCol2
    
    ' Sortiertes Array bauen
    Dim result() As Variant
    ReDim result(1 To rowCount, 1 To colCount)
    Dim j As Long
    For i = 1 To rowCount
        For j = 1 To colCount
            result(i, j) = data(idx(i), j)
        Next j
    Next i
    
    SortArray = result
End Function

Private Sub QuickSortIndex(ByRef idx() As Long, ByRef data As Variant, _
                           ByVal sortCol As Long, ByVal low As Long, _
                           ByVal high As Long, ByVal ascending As Boolean, _
                           ByVal sortCol2 As Long)
    If low >= high Then Exit Sub
    
    Dim pivot As Long, i As Long, j As Long, temp As Long
    pivot = idx((low + high) \ 2)
    i = low: j = high
    
    Do While i <= j
        If ascending Then
            Do While CompareValues(data(idx(i), sortCol), data(pivot, sortCol), sortCol2, idx(i), pivot, data) < 0
                i = i + 1
            Loop
            Do While CompareValues(data(idx(j), sortCol), data(pivot, sortCol), sortCol2, idx(j), pivot, data) > 0
                j = j - 1
            Loop
        Else
            Do While CompareValues(data(idx(i), sortCol), data(pivot, sortCol), sortCol2, idx(i), pivot, data) > 0
                i = i + 1
            Loop
            Do While CompareValues(data(idx(j), sortCol), data(pivot, sortCol), sortCol2, idx(j), pivot, data) < 0
                j = j - 1
            Loop
        End If
        If i <= j Then
            temp = idx(i): idx(i) = idx(j): idx(j) = temp
            i = i + 1: j = j - 1
        End If
    Loop
    
    If low < j Then QuickSortIndex idx, data, sortCol, low, j, ascending, sortCol2
    If i < high Then QuickSortIndex idx, data, sortCol, i, high, ascending, sortCol2
End Sub

Private Function CompareValues(v1 As Variant, v2 As Variant, _
                               sortCol2 As Long, idx1 As Long, idx2 As Long, _
                               data As Variant) As Long
    Dim cmp As Long
    If IsDate(v1) And IsDate(v2) Then
        If CDate(v1) < CDate(v2) Then
            cmp = -1
        ElseIf CDate(v1) > CDate(v2) Then
            cmp = 1
        Else
            cmp = 0
        End If
    ElseIf IsNumeric(v1) And IsNumeric(v2) Then
        If CDbl(v1) < CDbl(v2) Then
            cmp = -1
        ElseIf CDbl(v1) > CDbl(v2) Then
            cmp = 1
        Else
            cmp = 0
        End If
    Else
        cmp = StrComp(CStr(v1), CStr(v2), vbTextCompare)
    End If
    
    ' Sekundäre Sortierung
    If cmp = 0 And sortCol2 > 0 Then
        Dim s1 As Variant, s2 As Variant
        s1 = data(idx1, sortCol2): s2 = data(idx2, sortCol2)
        If IsDate(s1) And IsDate(s2) Then
            If CDate(s1) < CDate(s2) Then
                cmp = -1
            ElseIf CDate(s1) > CDate(s2) Then
                cmp = 1
            End If
        ElseIf IsNumeric(s1) And IsNumeric(s2) Then
            If CDbl(s1) < CDbl(s2) Then
                cmp = -1
            ElseIf CDbl(s1) > CDbl(s2) Then
                cmp = 1
            End If
        Else
            cmp = StrComp(CStr(s1), CStr(s2), vbTextCompare)
        End If
    End If
    
    CompareValues = cmp
End Function

Public Function SumColumn(ByVal data As Variant, ByVal colIdx As Long) As Double
    ' Summiert eine Spalte im Array
    If IsEmpty(data) Then Exit Function
    Dim i As Long, total As Double
    For i = 1 To UBound(data, 1)
        If IsNumeric(data(i, colIdx)) Then total = total + CDbl(data(i, colIdx))
    Next i
    SumColumn = total
End Function

Public Function GroupBySum(ByVal data As Variant, ByVal groupCol As Long, ByRef sumCols() As Long) As Variant
    ' GROUP BY mit SUM – ersetzt die Zbirni-Reports
    If IsEmpty(data) Then
        GroupBySum = Empty
        Exit Function
    End If
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long, j As Long, k As Long
    Dim key As String
    Dim colCount As Long
    colCount = UBound(sumCols) - LBound(sumCols) + 1
    
    For i = 1 To UBound(data, 1)
        key = CStr(data(i, groupCol))
        If Not dict.Exists(key) Then
            Dim sums() As Double
            ReDim sums(LBound(sumCols) To UBound(sumCols))
            dict.Add key, sums
        End If
        Dim existing() As Double
        existing = dict(key)
        For j = LBound(sumCols) To UBound(sumCols)
            If IsNumeric(data(i, sumCols(j))) Then
                existing(j) = existing(j) + CDbl(data(i, sumCols(j)))
            End If
        Next j
        dict(key) = existing
    Next i
    
    ' Ergebnis als 2D-Array
    Dim result() As Variant
    ReDim result(1 To dict.count, 1 To colCount + 1)
    Dim keys As Variant
    keys = dict.keys
    For i = 0 To dict.count - 1
        result(i + 1, 1) = keys(i)
        Dim vals() As Double
        vals = dict(keys(i))
        For j = LBound(vals) To UBound(vals)
            result(i + 1, j - LBound(vals) + 2) = vals(j)
        Next j
    Next i
    
    GroupBySum = result
End Function

