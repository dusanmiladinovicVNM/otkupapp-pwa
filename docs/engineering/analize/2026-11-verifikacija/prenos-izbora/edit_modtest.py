import io, sys
P = "src-vba/modTest.bas"
raw = open(P, "rb").read().decode("ascii")
assert "\r\n" in raw
t = raw.replace("\r\n", "\n")

def rep(old, new, count=1):
    global t
    n = t.count(old)
    assert n == count, (n, old[:80])
    t = t.replace(old, new)

# 1) deklaracije
rep('''Private Const ERR_GOLDEN As Long = vbObjectError + 9501
''', '''Private Const ERR_GOLDEN As Long = vbObjectError + 9501
Private Const ERR_IZBOR As Long = vbObjectError + 9502
''')
rep('''Private m_Total As Long
Private m_Failed As Long
Private m_Report As String
''', '''Private m_Total As Long
Private m_Failed As Long
Private m_Report As String

' IZBOR TESTOVA ZA JEDAN PROLAZ (tools/run_vba.py --test, tools/dokaz.py --izbor).
'
' m_IzborZahtev postavlja SAMO RunAllTestsIzbor, a RunAllTests ga POTROSI na
' ulazu (PreuzmiIzbor) i prenese u m_Izbor. Sledeci goli RunAllTests je zato
' uvek pun: izbor koji bi prezivio prolaz pretvorio bi pun run u tih deo suite-a
' -- zelen, sa manje testova. Oblik je ",T_A,T_B," (zarez i na krajevima), da se
' ime trazi CELO: izbor "T_A" ne sme da pusti "T_AB".
Private m_IzborZahtev As String
Private m_Izbor As String
Private m_IzborPusteni As String
' VREME PO TESTU: redovi "sekunde<TAB>status<TAB>ime" za last_run_vreme.txt.
' Meri se telo testa (InvokeTest), ne ciscenje posle pada. Dijagnostika, ne
' verdikt -- run_vba ispisuje najsporije, a verdikt i dalje nosi last_run.txt.
Private m_Vreme As String
''')

# 2) RunAllTests: preuzimanje izbora na ulazu
rep('''    m_Total = 0
    m_Failed = 0
    m_Report = ""

    RunOne 1
''', '''    m_Total = 0
    m_Failed = 0
    m_Report = ""
    ' Izbor se preuzima PRE prvog RunOne i brise se iz zahteva istim potezom
    ' (v. deklaraciju m_IzborZahtev).
    m_Izbor = PreuzmiIzbor(m_IzborZahtev)
    m_IzborPusteni = ""
    m_Vreme = ""

    RunOne 1
''')

# 3) RunAllTests: kraj
rep('''    ' Nepotpun rollback: radi nad svojim privremenim listom, pa moze svuda.
    RunOne 200

    SetTestMode prevMode
    WriteResultFile
End Sub
''', '''    ' Nepotpun rollback: radi nad svojim privremenim listom, pa moze svuda.
    RunOne 200

    ' Trazen a nepusten test je PAD, ne tisina: pogresno ime u izboru bi inace
    ' dalo zelen prolaz sa nula testova.
    ProveriIzbor
    m_Izbor = ""

    SetTestMode prevMode
    WriteResultFile
    WriteVremeFile
End Sub

' Isti prolaz kao RunAllTests, ali samo nad imenovanim testovima -- i to u ISTOM
' redosledu u kome ih RunAllTests pusta, jer je redosled nosec (v. komentare uz
' 124-126). Spisak je "T_A,T_B"; prazan spisak je greska, ne pun run. Zove ga
' tools/run_vba.py --test, a rezultat ide u isti last_run.txt.
'
' Test koji SAM ne prolazi, a u punom prolazu prolazi, zavisi od stanja koje mu
' ostavlja raniji test. To nije kvar izbora nego nalaz o testu -- dokaz.py ga
' zato meri celom suite-om, a `dokaz.py --izolacija` ih nabraja.
Public Sub RunAllTestsIzbor(ByVal spisak As String)
    m_IzborZahtev = NormalizujIzbor(spisak)
    If Len(m_IzborZahtev) = 0 Then
        Err.Raise ERR_IZBOR, "modTest.RunAllTestsIzbor", _
                  "IZBOR: prazan spisak testova"
    End If
    RunAllTests
End Sub
''')

# 4) RunOne: izbor + vreme
rep('''Private Sub RunOne(ByVal idx As Long)
    Dim nm As String
    Dim errNum As Long, errDesc As String
    nm = TestName(idx)

    On Error GoTo EH
    m_Total = m_Total + 1
    InvokeTest idx
    AppendReport nm, "OK", ""
    Exit Sub

EH:
    ' Err se cita PRE ciscenja: CleanupPosleTesta ide kroz On Error Resume Next
    ' (OtkupUI_Release ga ima), a to brise Err -- bez ovoga bi izvestaj o padu
    ' ostao prazan.
    errNum = Err.Number
    errDesc = Err.description
    m_Failed = m_Failed + 1
''', '''Private Sub RunOne(ByVal idx As Long)
    Dim nm As String
    Dim errNum As Long, errDesc As String
    Dim t0 As Single
    nm = TestName(idx)
    ' Test van izbora se ne broji -- ni u TESTS= ni u vremenu.
    If Not TestIzabran(nm) Then Exit Sub

    On Error GoTo EH
    m_Total = m_Total + 1
    t0 = Timer
    InvokeTest idx
    ZapisiVreme nm, "OK", t0
    AppendReport nm, "OK", ""
    Exit Sub

EH:
    ' Err se cita PRE ciscenja: CleanupPosleTesta ide kroz On Error Resume Next
    ' (OtkupUI_Release ga ima), a to brise Err -- bez ovoga bi izvestaj o padu
    ' ostao prazan.
    errNum = Err.Number
    errDesc = Err.description
    ZapisiVreme nm, "FAIL", t0
    m_Failed = m_Failed + 1
''')

# 5) pomocne procedure izbora i vremena -- odmah posle RunOne
i = t.index("Private Sub RunOne(ByVal idx As Long)")
j = t.index("\nEnd Sub\n", i) + len("\nEnd Sub\n")
pomocne = '''
' --- izbor testova i vreme po testu -----------------------------------------
'
' Blizanac u modBusinessFlowProTests (Bfp*): iste cetiri funkcije izbora, jer
' modul te suite ne sme da zavisi od modTest. Menjas pravilo ovde -- menjaj i
' tamo; tvrdnje o oba stoje u T_TestIzbor_TacnoImeIJednokratno i
' Test_BFP_IzborBiraCeloIme.

' ",T_A,T_B," iz "T_A, T_B" -- prazni clanovi i razmaci otpadaju. Prazno kad u
' spisku nema nijednog imena.
Private Function NormalizujIzbor(ByVal spisak As String) As String
    Dim delovi As Variant
    Dim i As Long
    Dim ime As String, izlaz As String
    delovi = Split(spisak, ",")
    For i = LBound(delovi) To UBound(delovi)
        ime = Trim$(CStr(delovi(i)))
        If Len(ime) > 0 Then izlaz = izlaz & "," & ime
    Next i
    If Len(izlaz) > 0 Then NormalizujIzbor = izlaz & ","
End Function

' Celo ime, ne podniz: zarezi sa obe strane su deo trazenog teksta.
Private Function IzborSadrzi(ByVal izbor As String, ByVal ime As String) As Boolean
    IzborSadrzi = (InStr(1, izbor, "," & ime & ",", vbTextCompare) > 0)
End Function

' Vrati zahtev i ISPRAZNI ga -- izbor traje tacno jedan prolaz.
Private Function PreuzmiIzbor(ByRef zahtev As String) As String
    PreuzmiIzbor = zahtev
    zahtev = ""
End Function

' Imena iz izbora koja se nisu pojavila u `pusteni`, razdvojena zarezom.
Private Function IzborNepusteni(ByVal izbor As String, ByVal pusteni As String) As String
    Dim delovi As Variant
    Dim i As Long
    Dim izlaz As String
    delovi = Split(izbor, ",")
    For i = LBound(delovi) To UBound(delovi)
        If Len(delovi(i)) > 0 Then
            If Not IzborSadrzi(pusteni, CStr(delovi(i))) Then
                If Len(izlaz) > 0 Then izlaz = izlaz & ","
                izlaz = izlaz & CStr(delovi(i))
            End If
        End If
    Next i
    IzborNepusteni = izlaz
End Function

Private Function TestIzabran(ByVal nm As String) As Boolean
    If Len(m_Izbor) = 0 Then
        TestIzabran = True
    ElseIf IzborSadrzi(m_Izbor, nm) Then
        m_IzborPusteni = m_IzborPusteni & "," & nm & ","
        TestIzabran = True
    End If
End Function

Private Sub ProveriIzbor()
    Dim nepusteni As Variant
    Dim i As Long
    If Len(m_Izbor) = 0 Then Exit Sub
    nepusteni = Split(IzborNepusteni(m_Izbor, m_IzborPusteni), ",")
    For i = LBound(nepusteni) To UBound(nepusteni)
        m_Total = m_Total + 1
        m_Failed = m_Failed + 1
        AppendReport CStr(nepusteni(i)), "FAIL", _
            "IZBOR: test se ne pusta u RunAllTests (nepoznato ime ili nema RunOne)"
    Next i
End Sub

Private Sub ZapisiVreme(ByVal nm As String, ByVal status As String, ByVal t0 As Single)
    m_Vreme = m_Vreme & SekundeOd(t0) & vbTab & status & vbTab & nm & vbLf
End Sub

' Proteklo vreme kao tekst sa TACKOM: Str$ ne zavisi od regionalnih
' podesavanja, a Format$ bi na srpskom Windows-u dao zarez. Timer se u ponoc
' vraca na nulu, pa negativna razlika znaci prelaz dana.
Private Function SekundeOd(ByVal t0 As Single) As String
    Dim d As Double
    d = CDbl(Timer) - CDbl(t0)
    If d < 0 Then d = d + 86400#
    SekundeOd = Trim$(Str$(Round(d, 3)))
End Function

' Dijagnostika, ne verdikt: neuspeh upisa ne sme da obori prolaz, pa ni da
' ostavi gresku u Err pozivaocu.
Private Sub WriteVremeFile()
    On Error Resume Next
    WriteTextFile ThisWorkbook.path & Application.PathSeparator & _
                  "last_run_vreme.txt", m_Vreme
    Err.Clear
End Sub
'''
t = t[:j] + pomocne + t[j:]

# 6) registar: test 311
rep('''        Case 310: TestName = "T_S8R1a_AltF8BrojUJedanDokument"
''', '''        Case 310: TestName = "T_S8R1a_AltF8BrojUJedanDokument"
        Case 311: TestName = "T_TestIzbor_TacnoImeIJednokratno"
''')
rep('''        Case 310: T_S8R1a_AltF8BrojUJedanDokument
''', '''        Case 310: T_S8R1a_AltF8BrojUJedanDokument
        Case 311: T_TestIzbor_TacnoImeIJednokratno
''')
rep('''    RunOne 310
''', '''    RunOne 310
    ' Ciste funkcije, bez fixture-a -- redosled nije nosec.
    RunOne 311
''')

# 7) telo testa -- iza T_ParseDatum_Ugovor (isti rod: cist ugovor funkcije)
sidro = "' ID PARCELE JE SKRIVENA DRUGA KOLONA combo-a"
assert t.count(sidro) == 1
test = '''' IZBOR TESTOVA (run_vba.py --test, dokaz.py --izbor) BIRA CELO IME I TRAJE
' JEDAN PROLAZ.
'
' Oba kvara bi bila TIHA i oba bi davala zelen prolaz sa pogresnim testovima:
'   - podniz umesto celog imena: izbor "T_A" pusti i "T_AB", pa dokaz meri tudji
'     test, a nepusten trazeni ostaje neprimecen ako se poklopi sa drugim;
'   - izbor koji prezivi prolaz: sledeci goli RunAllTests tiho pusta samo
'     prethodni izbor.
' Treci -- trazen a nepusten test -- mora da bude imenovan, jer ProveriIzbor od
' njega pravi FAIL.
'
' Sta ovaj test NE meri: da RunAllTests zaista zove PreuzmiIzbor. To se ne moze
' izmeriti iznutra (test radi usred tog prolaza), pa ga tvrdi
' `run_vba.py --self-test` nad izvorom.
Private Sub T_TestIzbor_TacnoImeIJednokratno()
    Dim zahtev As String

    AssertEq NormalizujIzbor(" T_A, ,T_B "), ",T_A,T_B,", _
             "izbor se normalizuje u ,ime,ime,"
    AssertEq NormalizujIzbor(" , "), "", "spisak bez imena je prazan izbor"

    AssertEq IzborSadrzi(",T_A,T_B,", "T_B"), True, "izbor nalazi svoje ime"
    AssertEq IzborSadrzi(",T_AB,", "T_A"), False, "izbor T_AB ne pusta T_A"
    AssertEq IzborSadrzi(",T_A,", "T_AB"), False, "izbor T_A ne pusta T_AB"

    AssertEq IzborNepusteni(",T_A,T_B,", ",T_A,"), "T_B", _
             "nepusten test iz izbora se imenuje"
    AssertEq IzborNepusteni(",T_A,", ",T_A,"), "", "pusten izbor nema nepustenih"

    zahtev = ",T_X,"
    AssertEq PreuzmiIzbor(zahtev), ",T_X,", "prolaz preuzima zahtevan izbor"
    AssertEq PreuzmiIzbor(zahtev), "", _
             "izbor traje jedan prolaz -- sledeci RunAllTests je pun"
End Sub

'''
t = t.replace(sidro, test + sidro)

open(P, "wb").write(t.replace("\n", "\r\n").encode("ascii"))
print("ok")
