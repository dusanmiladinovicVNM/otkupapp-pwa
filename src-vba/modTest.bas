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
        Case Else: TestName = "T_Nepoznat_" & idx
    End Select
End Function

' Direktan poziv (ne Application.Run) -- tako VBA mora da kompajlira i test i
' sve sto test referencira.
Private Sub InvokeTest(ByVal idx As Long)
    Select Case idx
        Case 1: T_PosleSnimanja_ZadrzavaKontekstOtpremnice
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
    Set f = New frmOtkup

    Dim ctlCount As Long
    ctlCount = f.Controls.count          ' bez ovoga se UserForm_Initialize ne okine

    ' Kontekst otpremnice OTP-TEST-1 iz fixture-a (bez .Show)
    f.txtDatum.value = FX_DATUM
    f.txtBrojZbirne.value = FX_ZBIRNA
    f.txtBrojDokumenta.value = FX_BROJ_OTP
    f.cmbKooperant.value = FX_KOOPERANT

    f.ClearOtkupFields

    AssertEq f.txtDatum.value, FX_DATUM, _
             "datum posle snimanja mora da ostane datum otpremnice"

    AssertSnapshot DumpKontrole(f), "PosleSnimanja_KontekstOtpremnice"

    Unload f
End Sub

' ============================================================
' Assert-i
' ============================================================
Public Sub AssertEq(ByVal actual As Variant, ByVal expected As Variant, _
                    ByVal label As String)
    If CStr(actual) <> CStr(expected) Then
        Err.Raise ERR_ASSERT, "modTest.AssertEq", _
                  label & " -- ocekivano [" & CStr(expected) & _
                  "], dobijeno [" & CStr(actual) & "]"
    End If
End Sub

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
        arr(i, 1) = ctl.name & "=" & ControlValue(ctl)
    Next ctl

    Dim sorted As Variant
    sorted = SortArray(arr, 1, True)

    Dim sb As String
    For i = 1 To n
        sb = sb & CStr(sorted(i, 1)) & vbLf
    Next i

    DumpKontrole = sb
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
    Dim fnum As Integer
    fnum = FreeFile
    Open path For Input As #fnum
    ReadTextFile = Input$(LOF(fnum), fnum)
    Close #fnum
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
