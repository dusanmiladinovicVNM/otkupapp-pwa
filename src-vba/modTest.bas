Attribute VB_Name = "modTest"
Option Explicit

' ============================================================
' modTest
' Test suite koja pada na PONASANJU, ne na sintaksi. Cilj je da razlikuje
' ispravan od pokvarenog koda -- suite koja je zelena nad cistim kodom, a nije
' dokazano crvena nad pokvarenim, ne dokazuje nista.
'
' Pokretanje: tools/run_vba.py zove Run("RunAllTests") nad temp kopijom
' fixture-a (tests/fixtures/otkup_test.xlsm, pravi ga tools/make_fixture.py).
' Rezultat ide u last_run.txt PORED SVESKE (dakle u temp folder), prvi red
' "TESTS=n FAIL=m". Driver taj fajl cita; nema fajla = pad.
'
' Compile signal stize sam: da bi se RunAllTests uopste pokrenuo, VBA mora da
' kompajlira modTest i sve sto on referencira -- a to je bas kod pod testom
' (frmOtkup, modOtkup, modOtkupBlok). Zato ovde nema posebnog compile gate-a.
'
' Greska se hvata PO TESTU: jedan pad ne obara ostale, i u ispisu stoji ime
' bas tog testa.
' ============================================================

' --- Fixture konstante (moraju da prate tools/make_fixture.py) --------------
Private Const FX_DATUM As String = "15.3.2026"      ' FIXTURE_DATE, d.m.yyyy
Private Const FX_ZBIRNA As String = "ZB-TEST-1"     ' zbirna na OTP-TEST-1
Private Const FX_BROJ_OTP As String = "1/TEST"      ' BrojOtpremnice OTP-TEST-1
Private Const FX_KOOPERANT As String = "KOOP-TEST-1"
Private Const FX_KOOPERANT2 As String = "KOOP-TEST-2"

Private Const ERR_ASSERT As Long = vbObjectError + 9500
Private Const ERR_GOLDEN As Long = vbObjectError + 9501

Private m_Total As Long
Private m_Failed As Long
Private m_Report As String

' ============================================================
' Ulazna tacka
' ============================================================
Public Sub RunAllTests()
    Dim prevMode As Boolean
    prevMode = IsTestMode()
    SetTestMode True

    m_Total = 0
    m_Failed = 0
    m_Report = ""

    RunOne 1
    RunOne 2
    RunOne 3
    RunOne 4          ' menja semu fixture-a -> mora posle testova forme

    SetTestMode prevMode
    WriteResultFile
End Sub

' Svaki test se zove kroz ovu omotnicu: broji se, greska mu se hvata i upisuje
' pod NJEGOVIM imenom. Ime se razresava pre poziva da bi bilo poznato i kad
' test pukne.
Private Sub RunOne(ByVal idx As Long)
    Dim nm As String
    nm = TestName(idx)

    On Error GoTo EH
    m_Total = m_Total + 1
    InvokeTest idx
    AppendReport nm, "OK", ""
    Exit Sub

EH:
    m_Failed = m_Failed + 1
    AppendReport nm, "FAIL", Err.description
End Sub

Private Function TestName(ByVal idx As Long) As String
    Select Case idx
        Case 1: TestName = "T_PosleSnimanja_ZadrzavaKontekstOtpremnice"
        Case 2: TestName = "T_PosleSnimanja_ZadrzavaZbirnu"
        Case 3: TestName = "T_ClearForm_BrisePartnera"
        Case 4: TestName = "T_EnsureSchema_JeIdempotentna"
        Case Else: TestName = "T_Nepoznat_" & idx
    End Select
End Function

' Direktan poziv (ne Application.Run) -- tako VBA mora da kompajlira i test i
' sve sto test referencira.
Private Sub InvokeTest(ByVal idx As Long)
    Select Case idx
        Case 1: T_PosleSnimanja_ZadrzavaKontekstOtpremnice
        Case 2: T_PosleSnimanja_ZadrzavaZbirnu
        Case 3: T_ClearForm_BrisePartnera
        Case 4: T_EnsureSchema_JeIdempotentna
    End Select
End Sub

' ============================================================
' Testovi
' ============================================================

' Posle snimanja otkupnog lista kontekst otpremnice mora da ostane: datum se NE
' brise, jer sledeci blok ide u niz istog datuma. Pada ako se u ClearOtkupFields
' vrati brisanje datuma (txtDatum.value = "").
Private Sub T_PosleSnimanja_ZadrzavaKontekstOtpremnice()
    Dim f As frmOtkup
    Set f = NewOtkupForm()

    f.ClearOtkupFields

    AssertEq f.txtDatum.value, FX_DATUM, _
             "datum posle snimanja mora da ostane datum otpremnice"

    AssertSnapshot DumpKontrole(f), "PosleSnimanja_KontekstOtpremnice"

    Unload f
End Sub

' Broj zbirne ostaje popunjen posle snimanja: sledeci blok iste otpremnice mora
' da dobije istu zbirnu, inace operater kuca broj iznova na svaki unos. Pada ako
' se u ClearOtkupFields vrati txtBrojZbirne.value = "".
Private Sub T_PosleSnimanja_ZadrzavaZbirnu()
    Dim f As frmOtkup
    Set f = NewOtkupForm()

    f.ClearOtkupFields
    AssertEq f.txtBrojZbirne.value, FX_ZBIRNA, _
             "broj zbirne mora da ostane popunjen posle snimanja"

    ' Drugi blok nad istom otpremnicom -- posle jos jednog snimanja zbirna je ista.
    f.cmbKooperant.value = FX_KOOPERANT2
    f.ClearOtkupFields
    AssertEq f.txtBrojZbirne.value, FX_ZBIRNA, _
             "drugi blok mora da dobije istu zbirnu"

    Unload f
End Sub

' Kooperant se BRISE posle snimanja -- sledeci unos je nov partner. Suprotno od
' prethodna dva testa: ovde je brisanje trazeno ponasanje. Pada ako se iz
' ClearOtkupFields ukloni cmbKooperant.value = "".
Private Sub T_ClearForm_BrisePartnera()
    Dim f As frmOtkup
    Set f = NewOtkupForm()

    ' Preduslov: bez ovoga bi test bio zelen i kad kontrola uopste ne prima
    ' vrednost, pa ne bi merio nista.
    AssertEq f.cmbKooperant.value, FX_KOOPERANT, _
             "preduslov: kooperant je postavljen pre ciscenja"

    f.ClearOtkupFields

    AssertEq f.cmbKooperant.value, "", _
             "kooperant mora da bude obrisan posle snimanja"

    Unload f
End Sub

' Idempotencija je JEDINA osobina koju sve Ensure* rutine obecavaju, a nijedna je
' do sada nije dokazivala. Test pusti tiha schema jezgra DVAPUT i tvrdi da drugi
' prolaz nije promenio semu -- ni jednu tabelu, ni jednu kolonu.
'
' Pada na dve klase regresije:
'   - bezuslovan ListColumns.Add / ListObjects.Add u nekom Ensure* (drugi prolaz
'     doda duplu kolonu ili tabelu -> otisak se razlikuje)
'   - MsgBox vracen u tiho jezgro (suite bi stao na modalnom dijalogu; staticki
'     isto hvata vba_check pravilo ENSURE)
'
' Ide na kraj RunAllTests-a: menja semu fixture-a, pa ne sme da utice na testove
' forme pre sebe. Fixture je ionako temp kopija (docs/EXCEL_TEST_HARNESS.md).
Private Sub T_EnsureSchema_JeIdempotentna()
    Dim prvi As String
    Dim drugi As String

    ' Prvi prolaz dovodi svesku u zeljeno stanje (fixture donora ga nema ceo).
    EnsureRuntimeSchema
    EnsureDoradeSchemaCore
    EnsurePaletniListSchemaCore
    prvi = SchemaOtisak()

    ' Drugi prolaz sme SAMO da bude no-op.
    EnsureRuntimeSchema
    EnsureDoradeSchemaCore
    EnsurePaletniListSchemaCore
    drugi = SchemaOtisak()

    ' Preduslov: bez ijedne tabele otisak bi bio prazan u oba prolaza, pa bi test
    ' bio zelen a ne bi merio nista.
    If Len(prvi) = 0 Then
        Err.Raise ERR_ASSERT, "modTest.T_EnsureSchema_JeIdempotentna", _
                  "preduslov: sveska nema nijednu tabelu -- otisak je prazan"
    End If

    AssertEq drugi, prvi, _
             "drugi prolaz Ensure* jezgara je promenio semu -- nije idempotentno"
End Sub

' Forma sa kontekstom otpremnice OTP-TEST-1 iz fixture-a, bez .Show.
Private Function NewOtkupForm() As frmOtkup
    Dim f As frmOtkup
    Set f = New frmOtkup

    Dim ctlCount As Long
    ctlCount = f.Controls.count          ' bez ovoga se UserForm_Initialize ne okine

    f.txtDatum.value = FX_DATUM
    f.txtBrojZbirne.value = FX_ZBIRNA
    f.txtBrojDokumenta.value = FX_BROJ_OTP
    f.cmbKooperant.value = FX_KOOPERANT

    Set NewOtkupForm = f
End Function

' ============================================================
' Assert-i
' ============================================================
Public Sub AssertEq(ByVal actual As Variant, ByVal expected As Variant, _
                    ByVal label As String)
    Dim a As String, e As String
    a = SafeStr(actual)
    e = SafeStr(expected)
    If a <> e Then
        Err.Raise ERR_ASSERT, "modTest.AssertEq", _
                  label & " -- ocekivano [" & e & "], dobijeno [" & a & "]"
    End If
End Sub

' Prazna kontrola ume da vrati Null umesto "", a CStr(Null) puca ("Invalid use
' of Null") -- test bi pao na toj gresci umesto na ponasanju koje meri.
Private Function SafeStr(ByVal v As Variant) As String
    If IsNull(v) Then
        SafeStr = ""
    ElseIf IsEmpty(v) Then
        SafeStr = ""
    Else
        SafeStr = CStr(v)
    End If
End Function

' Snapshot hvata i polja koja niko nije trazio da se provere. Kad golden ne
' postoji, upisuje ga i PADA -- nov golden mora da prodje ljudski pregled pre
' nego sto postane merilo.
Public Sub AssertSnapshot(ByVal tekuci As String, ByVal imeGolden As String)
    Dim path As String
    path = GoldenDir() & imeGolden & ".txt"

    If Len(Dir$(path)) = 0 Then
        WriteTextFile path, tekuci
        Err.Raise ERR_GOLDEN, "modTest.AssertSnapshot", _
                  "Golden nije postojao -- upisan je (" & imeGolden & _
                  ".txt). Pregledaj ga i commit-uj, pa pokreni ponovo."
    End If

    Dim golden As String
    golden = ReadTextFile(path)
    If golden <> tekuci Then
        Err.Raise ERR_ASSERT, "modTest.AssertSnapshot", _
                  "snapshot " & imeGolden & " se razlikuje od golden-a -- " & _
                  FirstDiff(golden, tekuci)
    End If
End Sub

' ============================================================
' Pomocno
' ============================================================

' "Otisak" seme: naziv tabele + broj kolona, jedan red po tabeli. Dovoljno da
' uhvati i dodatu kolonu i dodatu tabelu, a ne zavisi od podataka. Redosled je
' redosled listova/tabela u svesci -- stabilan izmedju dva prolaza u istoj
' sesiji, pa razlika u redosledu i JESTE nalaz (nesto je dodato).
Private Function SchemaOtisak() As String
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim parts As String

    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            parts = parts & lo.name & "=" & lo.ListColumns.count & vbLf
        Next lo
    Next ws

    SchemaOtisak = parts
End Function

' Sve kontrole forme kao sortirano "ime=vrednost", jedan par po liniji.
' Sortira se postojecim modArrayUtils.SortArray (nema novog sorta).
Public Function DumpKontrole(ByVal f As Object) As String
    Dim n As Long
    n = f.Controls.count
    If n = 0 Then
        DumpKontrole = ""
        Exit Function
    End If

    Dim arr() As Variant
    ReDim arr(1 To n, 1 To 1)

    Dim ctl As Object
    Dim i As Long
    i = 0
    For Each ctl In f.Controls
        i = i + 1
        arr(i, 1) = AsciiEscape(ctl.name & "=" & ControlValue(ctl))
    Next ctl

    Dim sorted As Variant
    sorted = SortArray(arr, 1, True)

    Dim sb As String
    For i = 1 To n
        sb = sb & CStr(sorted(i, 1)) & vbLf
    Next i

    DumpKontrole = sb
End Function

' Sve van stampanog ASCII-ja ide kao \uXXXX. Bez ovoga je golden neupotrebljiv:
' VBA Print # pise u ANSI kodnu stranu, koja "Vrsta voca" sa ch ne moze da
' predstavi (cp1252) -- snimi se osakaceno, pa svako sledece poredjenje pada, a
' poruka o razlici izgleda kao da su stringovi isti jer se i ona gubi na istom
' mestu. Uz escape je golden cist ASCII, round-trip je tacan, a razlika citljiva.
Private Function AsciiEscape(ByVal s As String) As String
    Dim i As Long
    Dim ch As Long
    Dim out As String

    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1))
        If ch < 0 Then ch = ch + 65536      ' AscW je Integer: > 32767 dolazi negativno
        If ch = 92 Then
            out = out & "\\"            ' inace bi putanja "C:\users" izgledala kao escape
        ElseIf ch >= 32 And ch <= 126 Then
            out = out & Chr$(ch)
        Else
            out = out & "\u" & Right$("000" & Hex$(ch), 4)
        End If
    Next i

    AsciiEscape = out
End Function

' Kontrole nemaju sve .Value (Label/Frame imaju Caption, neke nemaju nista).
Private Function ControlValue(ByVal ctl As Object) As String
    Dim s As String

    On Error Resume Next
    Err.Clear
    s = CStr(ctl.value)
    If Err.Number <> 0 Then
        Err.Clear
        s = CStr(ctl.caption)
        If Err.Number <> 0 Then
            Err.Clear
            s = "<n/a>"
        End If
    End If
    On Error GoTo 0

    ControlValue = s
End Function

Private Function FirstDiff(ByVal a As String, ByVal b As String) As String
    Dim la As Variant, lb As Variant
    la = Split(a, vbLf)
    lb = Split(b, vbLf)

    Dim n As Long
    n = UBound(la)
    If UBound(lb) < n Then n = UBound(lb)

    Dim i As Long
    For i = 0 To n
        If la(i) <> lb(i) Then
            FirstDiff = "prva razlika: golden [" & la(i) & "] vs tekuci [" & lb(i) & "]"
            Exit Function
        End If
    Next i

    FirstDiff = "razlicit broj linija: golden " & (UBound(la) + 1) & _
                ", tekuci " & (UBound(lb) + 1)
End Function

' Golden fajlovi zive pored sveske; run_vba.py ih kopira iz tests/golden pre
' rana i vraca posle, da nov golden zavrsi u repou na pregled.
Private Function GoldenDir() As String
    Dim d As String
    d = ThisWorkbook.path & Application.PathSeparator & "golden"
    If Len(Dir$(d, vbDirectory)) = 0 Then MkDir d
    GoldenDir = d & Application.PathSeparator
End Function

Private Sub WriteTextFile(ByVal path As String, ByVal content As String)
    Dim fnum As Integer
    fnum = FreeFile
    Open path For Output As #fnum
    Print #fnum, content;
    Close #fnum
End Sub

Private Function ReadTextFile(ByVal path As String) As String
    Dim raw As String
    Dim fnum As Integer
    fnum = FreeFile
    Open path For Input As #fnum
    raw = Input$(LOF(fnum), fnum)
    Close #fnum

    ' CR se izbacuje: .gitattributes drzi golden na LF, ali klon sa drugim
    ' podesavanjem (ili rucno editovanje u Notepad-u) vrati CRLF, a tada golden
    ' vise nije jednak dump-u koji se spaja sa vbLf. Pravi CR u sadrzaju ne
    ' postoji -- AsciiEscape ga pretvara u \u000D.
    ReadTextFile = Replace$(raw, vbCr, "")
End Function

Private Sub AppendReport(ByVal testNm As String, ByVal status As String, _
                         ByVal detail As String)
    m_Report = m_Report & status & " " & testNm
    If Len(detail) > 0 Then m_Report = m_Report & " -- " & detail
    m_Report = m_Report & vbLf
End Sub

Private Sub WriteResultFile()
    Dim path As String
    path = ThisWorkbook.path & Application.PathSeparator & "last_run.txt"
    WriteTextFile path, "TESTS=" & m_Total & " FAIL=" & m_Failed & vbLf & m_Report
End Sub
