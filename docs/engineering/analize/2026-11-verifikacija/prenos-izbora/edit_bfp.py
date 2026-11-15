import re
P = "src-vba/modBusinessFlowProTests.bas"
raw = open(P, "rb").read().decode("ascii")
assert "\r\n" in raw
t = raw.replace("\r\n", "\n")

def rep(old, new):
    global t
    n = t.count(old)
    assert n == 1, (n, old[:80])
    t = t.replace(old, new)

# 1) deklaracije -- uz m_Report (deklaraciona sekcija, pre prve procedure)
rep('''Private m_Report As String

Private Const TEST_ST_ID As String = "ST-90001"
''', '''Private m_Report As String

' IZBOR TESTOVA I VREME PO KORAKU -- isti ugovor kao u modTest (v. tamo
' m_IzborZahtev): RunBusinessFlowProIzbor postavlja zahtev, a
' RunBusinessFlowProSuite ga POTROSI kao prvu naredbu, pa je sledeci goli prolaz
' uvek pun. Ovde nema RunOne omotnice, pa svaki poziv testa u suite-u ide kroz
' BfpIzabran: on odlucuje da li se test pusta i zatvara vreme prethodnog koraka.
' Vreme ide u last_run_bfp_vreme.txt ("sekunde<TAB>status<TAB>ime"); status je
' uvek KORAK, jer BFP verdikt nosi tvrdnja, ne test.
Private m_BfpIzborZahtev As String
Private m_BfpIzbor As String
Private m_BfpPusteni As String
Private m_BfpVreme As String
Private m_BfpKorak As String
Private m_BfpKorakT0 As Single

Private Const TEST_ST_ID As String = "ST-90001"
''')

# 2) telo suite-a
i = t.index("Public Sub RunBusinessFlowProSuite()")
j = t.index("\nEnd Sub\n", i) + len("\nEnd Sub\n")
telo = t[i:j]

telo_n = telo.replace('''Public Sub RunBusinessFlowProSuite()
    On Error GoTo EH

    BeginRun "BUSINESS FLOW PROFESSIONAL SUITE"
''', '''Public Sub RunBusinessFlowProSuite()
    ' PRVA naredba, pre ijedne koja moze da pukne: izbor koji se ne potrosi
    ' prezivi u sledeci goli prolaz i pretvori ga u tih deo suite-a.
    m_BfpIzbor = BfpPreuzmiIzbor(m_BfpIzborZahtev)
    m_BfpPusteni = ""
    m_BfpVreme = ""
    m_BfpKorak = ""

    On Error GoTo EH

    BeginRun "BUSINESS FLOW PROFESSIONAL SUITE"
''', 1)
assert telo_n != telo
telo = telo_n

telo_n = telo.replace('''    SeedBusinessFlowProMasterData
''', '''    ' Seed se pusta UVEK -- i uz izbor, jer ga testovi traze. Ima svoj korak u
    ' vremenu, da se ne pripise testu ispred sebe.
    BfpKorakPocni "SeedBusinessFlowProMasterData"
    SeedBusinessFlowProMasterData
''', 1)
assert telo_n != telo
telo = telo_n

linije = telo.split("\n")
promenjeno = 0
for k, ln in enumerate(linije):
    m = re.fullmatch(r"    (Test_\w+)\s*", ln)
    if m:
        ime = m.group(1)
        linije[k] = '    If BfpIzabran("%s") Then %s' % (ime, ime)
        promenjeno += 1
telo = "\n".join(linije)
assert promenjeno == 360, promenjeno

telo_n = telo.replace('''    On Error GoTo 0        ' verdikt podize EndRun -- bez ovoga bi skocio u EH i dvaput brojao
    EndRun
    Exit Sub

EH:
    LogFatal "RunBusinessFlowProSuite", Err.Number, Err.description
    EndRun
End Sub
''', '''    If BfpIzabran("Test_BFP_IzborBiraCeloIme") Then Test_BFP_IzborBiraCeloIme

    ' Trazen a nepusten test je PAD, ne tisina (v. modTest.ProveriIzbor).
    BfpProveriIzbor
    BfpZavrsiVreme

    On Error GoTo 0        ' verdikt podize EndRun -- bez ovoga bi skocio u EH i dvaput brojao
    EndRun
    Exit Sub

EH:
    LogFatal "RunBusinessFlowProSuite", Err.Number, Err.description
    BfpZavrsiVreme
    EndRun
End Sub

' Isti prolaz kao RunBusinessFlowProSuite, ali samo nad imenovanim testovima, u
' redosledu suite-a; seed se pusta uvek. Spisak je "Test_A,Test_B"; prazan
' spisak je greska, ne pun run. Zove ga tools/run_vba.py --test.
Public Sub RunBusinessFlowProIzbor(ByVal spisak As String)
    m_BfpIzborZahtev = BfpNormalizujIzbor(spisak)
    If Len(m_BfpIzborZahtev) = 0 Then
        Err.Raise ERR_BFP_SUITE_FAILED, "modBusinessFlowProTests.RunBusinessFlowProIzbor", _
                  "IZBOR: prazan spisak testova"
    End If
    RunBusinessFlowProSuite
End Sub

' --- izbor testova i vreme po koraku ----------------------------------------
'
' Blizanac pravila iz modTest (NormalizujIzbor / IzborSadrzi / PreuzmiIzbor /
' IzborNepusteni). Kopija je namerna: ovaj modul ne sme da zavisi od modTest.
' Menjas pravilo ovde -- menjaj i tamo; tvrdnje: Test_BFP_IzborBiraCeloIme i
' modTest.T_TestIzbor_TacnoImeIJednokratno.

Private Function BfpNormalizujIzbor(ByVal spisak As String) As String
    Dim delovi As Variant
    Dim i As Long
    Dim ime As String, izlaz As String
    delovi = Split(spisak, ",")
    For i = LBound(delovi) To UBound(delovi)
        ime = Trim$(CStr(delovi(i)))
        If Len(ime) > 0 Then izlaz = izlaz & "," & ime
    Next i
    If Len(izlaz) > 0 Then BfpNormalizujIzbor = izlaz & ","
End Function

Private Function BfpIzborSadrzi(ByVal izbor As String, ByVal ime As String) As Boolean
    BfpIzborSadrzi = (InStr(1, izbor, "," & ime & ",", vbTextCompare) > 0)
End Function

Private Function BfpPreuzmiIzbor(ByRef zahtev As String) As String
    BfpPreuzmiIzbor = zahtev
    zahtev = ""
End Function

Private Function BfpIzborNepusteni(ByVal izbor As String, ByVal pusteni As String) As String
    Dim delovi As Variant
    Dim i As Long
    Dim izlaz As String
    delovi = Split(izbor, ",")
    For i = LBound(delovi) To UBound(delovi)
        If Len(delovi(i)) > 0 Then
            If Not BfpIzborSadrzi(pusteni, CStr(delovi(i))) Then
                If Len(izlaz) > 0 Then izlaz = izlaz & ","
                izlaz = izlaz & CStr(delovi(i))
            End If
        End If
    Next i
    BfpIzborNepusteni = izlaz
End Function

' Da li se test pusta u ovom prolazu. Zatvara vreme prethodnog koraka u SVAKOM
' slucaju, a otvara svoje samo kad se pusta -- preskocen test nema vreme.
Private Function BfpIzabran(ByVal ime As String) As Boolean
    BfpKorakZatvori
    If Len(m_BfpIzbor) > 0 Then
        If Not BfpIzborSadrzi(m_BfpIzbor, ime) Then Exit Function
        m_BfpPusteni = m_BfpPusteni & "," & ime & ","
    End If
    m_BfpKorak = ime
    m_BfpKorakT0 = Timer
    BfpIzabran = True
End Function

Private Sub BfpKorakPocni(ByVal ime As String)
    BfpKorakZatvori
    m_BfpKorak = ime
    m_BfpKorakT0 = Timer
End Sub

' Str$ umesto Format$: tacka bez obzira na regionalna podesavanja. Timer se u
' ponoc vraca na nulu, pa negativna razlika znaci prelaz dana.
Private Sub BfpKorakZatvori()
    Dim d As Double
    If Len(m_BfpKorak) = 0 Then Exit Sub
    d = CDbl(Timer) - CDbl(m_BfpKorakT0)
    If d < 0 Then d = d + 86400#
    m_BfpVreme = m_BfpVreme & Trim$(Str$(Round(d, 3))) & vbTab & "KORAK" & _
                 vbTab & m_BfpKorak & vbLf
    m_BfpKorak = ""
End Sub

Private Sub BfpProveriIzbor()
    Dim nepusteni As Variant
    Dim i As Long
    If Len(m_BfpIzbor) = 0 Then Exit Sub
    nepusteni = Split(BfpIzborNepusteni(m_BfpIzbor, m_BfpPusteni), ",")
    For i = LBound(nepusteni) To UBound(nepusteni)
        LogFail "IZBOR " & CStr(nepusteni(i)), _
                "test se ne pusta u RunBusinessFlowProSuite (nepoznato ime ili nema poziva)"
    Next i
End Sub

' Dijagnostika, ne verdikt: neuspeh upisa ne sme da obori prolaz, ni da ostavi
' gresku u Err -- zove se i iz EH, pred EndRun.
Private Sub BfpZavrsiVreme()
    Dim fnum As Integer
    On Error Resume Next
    BfpKorakZatvori
    fnum = FreeFile
    Open ThisWorkbook.path & Application.PathSeparator & "last_run_bfp_vreme.txt" _
        For Output As #fnum
    Print #fnum, m_BfpVreme;
    Close #fnum
    Err.Clear
End Sub

' IZBOR TESTOVA BIRA CELO IME I TRAJE JEDAN PROLAZ -- isti ugovor kao
' modTest.T_TestIzbor_TacnoImeIJednokratno, nad kopijom pravila u ovom modulu.
' Ciste funkcije: ne cita ni ne pise tabele.
Private Sub Test_BFP_IzborBiraCeloIme()
    Dim zahtev As String

    AssertEquals ",Test_A,Test_B,", BfpNormalizujIzbor(" Test_A, ,Test_B "), _
                 "BFP izbor se normalizuje u ,ime,ime,"
    AssertEquals "", BfpNormalizujIzbor(" , "), "BFP spisak bez imena je prazan izbor"

    AssertTrue BfpIzborSadrzi(",Test_A,Test_B,", "Test_B"), "BFP izbor nalazi svoje ime"
    AssertFalse BfpIzborSadrzi(",Test_AB,", "Test_A"), "BFP izbor Test_AB ne pusta Test_A"
    AssertFalse BfpIzborSadrzi(",Test_A,", "Test_AB"), "BFP izbor Test_A ne pusta Test_AB"

    AssertEquals "Test_B", BfpIzborNepusteni(",Test_A,Test_B,", ",Test_A,"), _
                 "BFP nepusten test iz izbora se imenuje"

    zahtev = ",Test_X,"
    AssertEquals ",Test_X,", BfpPreuzmiIzbor(zahtev), "BFP prolaz preuzima zahtevan izbor"
    AssertEquals "", BfpPreuzmiIzbor(zahtev), _
                 "BFP izbor traje jedan prolaz -- sledeci prolaz je pun"
End Sub
''', 1)
assert telo_n != telo
telo = telo_n

t = t[:i] + telo + t[j:]
open(P, "wb").write(t.replace("\n", "\r\n").encode("ascii"))
print("ok")
