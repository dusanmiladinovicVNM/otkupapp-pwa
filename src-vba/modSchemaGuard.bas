Attribute VB_Name = "modSchemaGuard"

Option Explicit

' REGISTAR STORNA: koje tabele MORAJU nositi kolonu Stornirano.
'
' ExcludeStornirano je pitao GetColumnIndex za tu kolonu i na nulu tiho vracao
' NEFILTRIRANE podatke. Nula tamo ima dva znacenja: "ova tabela storno pojam
' nema" (tacno za maticne podatke) i "kolona nije nadjena" (kvar). Bez ove
' razlike je storniran dokument izlazio kao ziv -- iz 183 poziva, ukljucujuci
' read-modele otvorenih faktura i otkupnih blokova. To je gore od pogresne
' klasifikacije novca: otkazan dokument dobija pogresno POSTOJANJE.
'
' Spisak je DEKLARACIJA OCEKIVANJA, ne snimak zatecene sveske: sema je izvor
' istine po instalaciji, pa tabela iz ovog spiska koja nema kolonu znaci DRIFT,
' i tada se pada glasno umesto da se tiho ne filtrira.
'
' Statickim putem se cuva vba_check pravilom STORNO_REGISTAR: svaki
' ExcludeStornirano(..., TBL_X) mora da imenuje tabelu iz jednog od dva spiska.
Private Const STORNO_TABELE As String = "|" & TBL_OTKUP & "|" & TBL_NOVAC & _
    "|" & TBL_OTPREMNICA & "|" & TBL_ZBIRNA & "|" & TBL_PRIJEMNICA & _
    "|" & TBL_FAKTURE & "|" & TBL_FAKTURA_STAVKE & "|" & TBL_MAGACIN & _
    "|" & TBL_BANKA_IMPORT & "|" & TBL_AMBALAZA & "|" & TBL_CENOVNIK & _
    "|" & TBL_PALETA & "|" & TBL_PALETA_STAVKA & "|" & TBL_PRERADA & _
    "|" & TBL_PRERADA_STAVKA & _
    "|" & TBL_UTOVAR & "|" & TBL_UTOVAR_STAVKE & _
    "|" & TBL_LAGER_JEDINICE & "|" & TBL_PROCES_SARZE & "|" & TBL_PROCES_ULAZI & _
    "|" & TBL_PROCES_IZLAZI & "|" & TBL_PROCES_PARAMETRI & "|"

' Tabele koje storno pojam NEMAJU -- maticni podaci. Prolaz kroz filter je za
' njih tacan ishod, ne propust, i navedene su izricito da se "nije u spisku"
' ne bi moglo procitati kao "zaboravljeno".
Private Const BEZ_STORNA As String = "|" & TBL_KOOPERANTI & "|" & TBL_KUPCI & _
    "|" & TBL_VOZACI & "|" & TBL_STANICE & "|" & TBL_PARCELE & _
    "|" & TBL_ARTIKLI & "|" & TBL_PREVOZNICI & _
    "|" & TBL_TIPOVI_PROCESA & "|" & TBL_PROIZVODI & "|" & TBL_OPREMA & "|"

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

' Mora li ova tabela da nosi kolonu Stornirano.
Public Function TabelaNosiStorno(ByVal tableName As String) As Boolean
    TabelaNosiStorno = _
        (InStr(1, STORNO_TABELE, "|" & Trim$(tableName) & "|", vbTextCompare) > 0)
End Function

' Da li registar uopste ZNA za ovu tabelu -- u bilo kom od dva spiska.
'
' Nepoznata tabela nije ni jedno ni drugo, i tu razliku mora da pravi IZVRSAVANJE,
' ne samo staticka provera: ona namerno preskace pozive sa promenljivim imenom
' tabele (modIntegritet.CollectBrojZbirne, modDokumenta.SumByBroj i slicni), pa
' bi bez ove kapije "TabelaNosiStorno = False" opet znacilo dve stvari --
' "eksplicitno BEZ_STORNA" i "niko je nije klasifikovao". To je ista bolest zbog
' koje je ceo ovaj posao i nastao, samo jedan nivo dalje.
Public Function StornoRegistarZna(ByVal tableName As String) As Boolean
    Dim k As String
    k = "|" & Trim$(tableName) & "|"
    StornoRegistarZna = (InStr(1, STORNO_TABELE, k, vbTextCompare) > 0) Or _
                        (InStr(1, BEZ_STORNA, k, vbTextCompare) > 0)
End Function

' Tabela mora da bude KLASIFIKOVANA pre nego sto se nad njom filtrira storno.
' Bez klasifikacije nema tacnog ishoda: ni pad ni prolaz nisu opravdani, jer se
' nenadjena kolona ne moze razlikovati od tabele koja storno pojam nema.
Public Sub RequireStornoKlasifikaciju(ByVal tableName As String, _
                                      ByVal sourceName As String)
    If StornoRegistarZna(tableName) Then Exit Sub
    Err.Raise vbObjectError + 7302, sourceName, _
              "Tabela '" & tableName & "' nije klasifikovana u registru storna " & _
              "(modSchemaGuard: STORNO_TABELE ili BEZ_STORNA). Dok nije, " & _
              "nenadjena kolona '" & COL_STORNIRANO & "' se ne moze razlikovati " & _
              "od tabele koja storno pojam nema."
End Sub

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

