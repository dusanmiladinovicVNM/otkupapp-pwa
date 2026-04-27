Attribute VB_Name = "modSchemaGuard"

Option Explicit

Public Function RequireColumnIndex(ByVal tableName As String, _
                                   ByVal columnName As String, _
                                   ByVal sourceName As String) As Long
    Dim idx As Long

    idx = GetColumnIndex(tableName, columnName)

    If idx = 0 Then
        Err.Raise vbObjectError + 7300, sourceName, _
                  "Nedostaje kolona '" & columnName & "' u tabeli '" & tableName & "'."
    End If

    RequireColumnIndex = idx
End Function

Public Sub RequireColumns(ByVal tableName As String, _
                          ByVal sourceName As String, _
                          ParamArray columnNames() As Variant)
    Dim i As Long

    For i = LBound(columnNames) To UBound(columnNames)
        If RequireColumnIndex(tableName, CStr(columnNames(i)), sourceName) = 0 Then
            ' RequireColumnIndex vec baca grešku.
        End If
    Next i
End Sub

Public Sub RequireUpdateCell(ByVal tableName As String, _
                              ByVal rowIndex As Long, _
                              ByVal columnName As String, _
                              ByVal newValue As Variant, _
                              ByVal sourceName As String)
    If Not UpdateCell(tableName, rowIndex, columnName, newValue) Then
        Err.Raise vbObjectError + 7400, sourceName, _
                  "UpdateCell fehlgeschlagen: " & tableName & "." & columnName
    End If
End Sub

