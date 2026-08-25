Attribute VB_Name = "modSchemaGuard"

Option Explicit

' PRAZNA TABELA I NEPOSTOJECA TABELA NISU ISTI ISHOD.
'
' GetTableData vraca Empty za oba, pa citac koji radi samo
' "If IsEmpty(data) Then Exit Function" tumaci nedostajucu tabelu kao "nema
' redova". Tamo gde prazna lista nosi POSLOVNO ZNACENJE -- prazan izbor fakture
' je avans, prazan izbor bloka je poziv na broj -- to je fail-open: kvar postane
' legitiman drugi ishod.
'
' RequireColumnIndex ovo NE pokriva: kad tabele nema, citac izadje pre nego sto
' do provere kolona uopste dodje.
Public Sub RequireTable(ByVal tableName As String, ByVal sourceName As String)
    If GetTable(tableName) Is Nothing Then
        Err.Raise vbObjectError + 7301, sourceName, _
                  "Tabela '" & tableName & "' nije dostupna."
    End If
End Sub

' Zaglavlje tabele kao kratak spisak, za poruku o gresci. Nikad ne puca: ovo je
' dijagnostika, i ne sme da zameni gresku koju opisuje.
Private Function ZaglavljeZaPoruku(ByVal tableName As String) As String
    Const MAX_IMENA As Long = 12
    Dim lo As ListObject, i As Long, n As Long, s As String

    On Error Resume Next
    Set lo = GetTable(tableName)
    If lo Is Nothing Then
        ZaglavljeZaPoruku = "tabela nije nadjena"
        Err.Clear
        Exit Function
    End If

    n = lo.ListColumns.count
    For i = 1 To n
        If i > MAX_IMENA Then
            s = s & ", ... (+" & (n - MAX_IMENA) & ")"
            Exit For
        End If
        If Len(s) > 0 Then s = s & ", "
        s = s & lo.ListColumns(i).name
    Next i

    If Len(s) = 0 Then s = "prazno"
    ZaglavljeZaPoruku = s
    Err.Clear
End Function

Public Function RequireColumnIndex(ByVal tableName As String, _
                                   ByVal columnName As String, _
                                   ByVal sourceName As String) As Long
    Dim idx As Long

    idx = GetColumnIndex(tableName, columnName)

    If idx = 0 Then
        ' Poruka nosi i ZAGLAVLJE koje je stvarno videla.
        '
        ' Bez toga se "nedostaje kolona" ne moze razlikovati od "tabele nema",
        ' "zaglavlje je drugacije" ili "citanje je puklo" -- a bas to je jednom
        ' kostalo pola dana nad sveskom u kojoj je kolona postojala. Spisak je
        ' ogranicen, jer poruka ide u log i u dijalog.
        Err.Raise vbObjectError + 7300, sourceName, _
                  "Nedostaje kolona '" & columnName & "' u tabeli '" & tableName & _
                  "'. Vidjeno zaglavlje: " & ZaglavljeZaPoruku(tableName) & "."
    End If

    RequireColumnIndex = idx
End Function

Public Sub RequireColumns(ByVal tableName As String, _
                          ByVal sourceName As String, _
                          ParamArray columnNames() As Variant)
    Dim i As Long

    For i = LBound(columnNames) To UBound(columnNames)
        If RequireColumnIndex(tableName, CStr(columnNames(i)), sourceName) = 0 Then
            ' RequireColumnIndex vec baca gresku.
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

