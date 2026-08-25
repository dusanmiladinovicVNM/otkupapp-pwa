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

' Zaglavlje tabele za poruku o gresci, i odgovor na jedino pitanje koje ovde
' vredi: DA LI JE BAS TRAZENA KOLONA VIDJENA u svezem prolazu.
'
' Nikad ne puca -- dijagnostika ne sme da zameni gresku koju opisuje. Ali ne sme
' ni da GRESKU pretvori u lazno stanje: prva verzija je sve drzala pod jednim
' `On Error Resume Next`, pa bi pad citanja tabele prijavila kao "tabela nije
' nadjena", a pad citanja zaglavlja kao "prazno". To je ista bolest zbog koje je
' ovaj posao i nastao, samo jedan nivo nize. Zato se posle svakog rizicnog koraka
' Err CITA i, ako je postavljen, kaze se da citanje nije uspelo.
'
' Spisak imena je ogranicen (poruka ide u log i u dijalog), ali se TRAZENA kolona
' trazi kroz CELO zaglavlje -- inace bi kolona iza granice ostala nevidljiva bas
' u poruci koja treba da kaze da li postoji.
Private Function ZaglavljeZaPoruku(ByVal tableName As String, _
                                   ByVal columnName As String) As String
    Const MAX_IMENA As Long = 12
    Dim lo As ListObject, i As Long, n As Long, s As String
    Dim errNum As Long, errDesc As String
    Dim poz As Long, ime As String

    On Error Resume Next
    Err.Clear
    Set lo = GetTable(tableName)
    errNum = Err.Number: errDesc = Err.description
    Err.Clear
    On Error GoTo 0
    If errNum <> 0 Then
        ZaglavljeZaPoruku = "citanje tabele NIJE uspelo (Err " & errNum & " " & _
                            errDesc & ")"
        Exit Function
    End If
    If lo Is Nothing Then
        ZaglavljeZaPoruku = "tabela nije nadjena"
        Exit Function
    End If

    On Error Resume Next
    Err.Clear
    n = lo.ListColumns.count
    errNum = Err.Number: errDesc = Err.description
    Err.Clear
    On Error GoTo 0
    If errNum <> 0 Then
        ZaglavljeZaPoruku = "citanje zaglavlja NIJE uspelo (Err " & errNum & " " & _
                            errDesc & ")"
        Exit Function
    End If

    For i = 1 To n
        On Error Resume Next
        Err.Clear
        ime = lo.ListColumns(i).name
        errNum = Err.Number
        Err.Clear
        On Error GoTo 0
        If errNum <> 0 Then
            ZaglavljeZaPoruku = "citanje imena kolone " & i & " NIJE uspelo (Err " & _
                                errNum & ")"
            Exit Function
        End If
        If poz = 0 Then
            If StrComp(ime, columnName, vbTextCompare) = 0 Then poz = i
        End If
        If i <= MAX_IMENA Then
            If Len(s) > 0 Then s = s & ", "
            s = s & ime
        End If
    Next i

    If n > MAX_IMENA Then s = s & ", ... (+" & (n - MAX_IMENA) & ")"
    If Len(s) = 0 Then s = "prazno"

    ' Ovo je podatak zbog kojeg poruka i postoji: ako trazenje kaze da kolone
    ' nema, a svez prolaz je vidi, onda uzrok nije sema nego put do nje.
    If poz > 0 Then
        ZaglavljeZaPoruku = s & ". Trazena kolona VIDJENA, pozicija " & poz
    Else
        ZaglavljeZaPoruku = s & ". Trazena kolona NIJE vidjena"
    End If
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
                  "'. Vidjeno zaglavlje: " & _
                  ZaglavljeZaPoruku(tableName, columnName) & "."
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

