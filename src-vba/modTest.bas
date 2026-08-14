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
'
' NOVI UI (frmOtkupUI + modOtkupUI) ima svoja tri testa, 4-6. Legacy forma se
' NE gasi (docs/UI_MIGRACIJA_KATALOG.md), pa oba skupa stoje jedan pored drugog
' -- ugovor je isti, kod je namerno dvostruk.
' ============================================================

' --- Fixture konstante (moraju da prate tools/make_fixture.py) --------------
Private Const FX_DATUM As String = "15.3.2026"      ' FIXTURE_DATE, d.m.yyyy
Private Const FX_ZBIRNA As String = "ZB-TEST-1"     ' zbirna na OTP-TEST-1
Private Const FX_BROJ_OTP As String = "1/TEST"      ' BrojOtpremnice OTP-TEST-1
Private Const FX_KOOPERANT As String = "KOOP-TEST-1"
Private Const FX_KOOPERANT2 As String = "KOOP-TEST-2"
Private Const FX_OTP_ID As String = "OTP-TEST-1"    ' otpremnica koja nosi FX_BROJ_OTP
Private Const FX_PARCELA As String = "PAR-TEST-1"   ' parcela kooperanta KOOP-TEST-1

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
    RunOne 4
    RunOne 5
    RunOne 6

    SetTestMode prevMode
    WriteResultFile
End Sub

' Svaki test se zove kroz ovu omotnicu: broji se, greska mu se hvata i upisuje
' pod NJEGOVIM imenom. Ime se razresava pre poziva da bi bilo poznato i kad
' test pukne.
Private Sub RunOne(ByVal idx As Long)
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
    CleanupPosleTesta
    ' Pad bez opisa je vec jednom kostao dva rana dijagnostike: "FAIL T_X" bez
    ' razloga ne kaze operateru nista. Broj greske je tada jedini trag.
    If Len(errDesc) = 0 Then errDesc = "greska bez opisa (Err.Number=" & errNum & ")"
    AppendReport nm, "FAIL", errDesc
End Sub

' Test koji je pao NIJE stigao do svog ReleaseOtkupUIForm, pa modul novog UI-ja
' (mFrm, Btns, kes tabela) i aktivna otpremnica u modScrDokumenti ostaju
' zaprljani. Sledeci test bi tada gradio ekran nad ostacima prethodnog i pao BEZ
' SVOJE KRIVICE -- jedna sabotaza obarala bi dva testa, pa bi drugi pad bio lazan
' trag. (Dokazano: sabotaza parcela-tekst obarala je i T_ClearForm_Ugovor, sa
' Err.Number=0 i praznim opisom.)
'
' Ciscenje je idempotentno (OtkupUI_Release je ceo pod On Error Resume Next,
' Scr_OtpOtkazi samo prazni tri promenljive), pa je bezbedno i posle testa koji
' formu nikad nije napravio. Samu formu otpusta odmotavanje steka -- ovde ostaje
' ono sto zivi na MODULIMA i sto odmotavanje ne dira.
Private Sub CleanupPosleTesta()
    On Error Resume Next
    modOtkupUI.OtkupUI_Release
    modScrDokumenti.Scr_OtpOtkazi
End Sub

Private Function TestName(ByVal idx As Long) As String
    Select Case idx
        Case 1: TestName = "T_PosleSnimanja_ZadrzavaKontekstOtpremnice"
        Case 2: TestName = "T_PosleSnimanja_ZadrzavaZbirnu"
        Case 3: TestName = "T_ClearForm_BrisePartnera"
        Case 4: TestName = "T_ParseDatum_Ugovor"
        Case 5: TestName = "T_ParcelaID_IzSkriveneKolone"
        Case 6: TestName = "T_ClearForm_Ugovor"
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
        Case 4: T_ParseDatum_Ugovor
        Case 5: T_ParcelaID_IzSkriveneKolone
        Case 6: T_ClearForm_Ugovor
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

' ============================================================
' Novi UI (frmOtkupUI + modOtkupUI)
' ============================================================

' DATUM DOKUMENTA ide u tblOtkup i u kontekst (predlog broja, zakljucavanje
' stanice), pa "necitljivo" mora da bude 0 -- nikad priblizan datum. Parser je
' NAMERNO deterministican (modParse.TryParseDateValue): CDate isti tekst cita po
' Windows locale-u, pa bi "01.02.2026" na MDY masini bio 2. januar a na DMY
' masini 1. februar. Pada ako se ParseDatum vrati na IsDate/CDate ili ako se
' izgubi skidanje trailing tacke.
Private Sub T_ParseDatum_Ugovor()
    AssertEq modOtkupUI.ParseDatum(""), 0, "prazno polje nije datum"
    AssertEq modOtkupUI.ParseDatum("   "), 0, "sami razmaci nisu datum"
    AssertEq modOtkupUI.ParseDatum("besmislica"), 0, "necitljiv tekst nije datum"

    AssertEq modOtkupUI.ParseDatum("11.08.2026"), CDbl(DateSerial(2026, 8, 11)), _
             "d.m.yyyy se cita kao dan.mesec.godina"

    ' Trailing tacka je nacin na koji se datum kod nas pise ("11.08.2026."), pa
    ' se skida umesto da obori unos. Petlja, ne jedno skidanje.
    AssertEq modOtkupUI.ParseDatum("11.08.2026."), CDbl(DateSerial(2026, 8, 11)), _
             "trailing tacka se skida, ne obara unos"
    AssertEq modOtkupUI.ParseDatum("11.08.2026.."), CDbl(DateSerial(2026, 8, 11)), _
             "skidaju se SVE trailing tacke, ne samo poslednja"

    ' AUD-007: DateSerial se na nemogucem datumu PRELIVA (30.02 -> 2.3, mesec 13
    ' -> januar sledece godine) umesto da pukne. Round-trip u parseru to odbija --
    ' inace bi dokument tiho dobio pomeren datum.
    AssertEq modOtkupUI.ParseDatum("30.02.2026"), 0, _
             "nepostojeci dan se odbija, ne preliva u sledeci mesec"
    AssertEq modOtkupUI.ParseDatum("01.13.2026"), 0, _
             "mesec 13 se odbija, ne preliva u sledecu godinu"

    ' Kapija poslovnih godina dolazi iz zajednickog parsera (modParse), ali se
    ' vidi kroz ovo polje -- zato stoji ovde, uz ostatak ugovora.
    AssertEq modOtkupUI.ParseDatum("11.08.1899"), 0, "godina van poslovnog opsega"
End Sub

' ID PARCELE JE SKRIVENA DRUGA KOLONA combo-a, kao kod svih ostalih dropdown-a
' (PartnerID / modComboBinding.GetComboID). Regres koji ovaj test cuva: ID se
' nekad vadio iz prikaznog teksta trazenjem " - ", a FillParcele gradi prikaz sa
' " " & ChrW(183) & " " -- separator se nikad nije nasao, pa je ceo prikazni
' string odlazio u ParcelaID i u tblOtkup. Pada ako se ID opet cita iz teksta,
' ili ako se izgubi provera vidljivosti polja.
Private Sub T_ParcelaID_IzSkriveneKolone()
    Dim f As frmOtkupUI, fr As Object, CB As MSForms.ComboBox
    Set f = NewOtkupUIForm()

    Set fr = f.Controls("zForm").Controls("fgParcela")
    Set CB = fr.Controls("fgParcelaT")
    fr.Visible = True

    ' Isti oblik koji gradi FillParcele: prikaz u koloni 1, ID u koloni 2.
    ' Prikaz NAMERNO nosi separator koji nije " - ".
    CB.Clear
    CB.ColumnCount = 2
    CB.BoundColumn = 1
    CB.TextColumn = 1
    CB.AddItem "1001   " & ChrW(183) & "   Malina   " & ChrW(183) & "   1,20 ha"
    CB.List(0, 1) = FX_PARCELA
    CB.ListIndex = 0

    ' Preduslov: bez ovoga bi test bio zelen i kad combo uopste ne prima stavke.
    AssertEq CB.ListCount, 1, "preduslov: parcela je u listi"

    AssertEq modOtkupUI.ParcelaID(), FX_PARCELA, _
             "ID parcele dolazi iz skrivene kolone, ne iz prikaznog teksta"

    CB.ListIndex = -1
    AssertEq modOtkupUI.ParcelaID(), "", "bez izabrane parcele dokument ne dobija ID"

    ' PRACENJE_PARCELA iskljuceno -> polje je sakriveno. Zatecen izbor tada NE
    ' sme da procuri u dokument.
    CB.ListIndex = 0
    fr.Visible = False
    AssertEq modOtkupUI.ParcelaID(), "", "sakriveno polje ne salje parcelu u dokument"

    ReleaseOtkupUIForm f
End Sub

' UGOVOR ClearForm-a, isti kao frmOtkup.ClearOtkupFields (.claude/rules/
' otkup-i-dokumenta.md odeljak 1 i 5): datum i broj zbirne su KONTEKST
' otpremnice i ostaju, partner se brise. Uz to i nova razlika koju legacy nema:
' bez aktivne otpremnice datum se vraca na danas.
'
' Zasto datum: otpremnica 8/220726 od 22.07 dobijala je blok 8/110826 od 11.08 --
' vracanje na danas je i broj i datum bloka odvlacilo iz niza otpremnice.
Private Sub T_ClearForm_Ugovor()
    Dim f As frmOtkupUI, zf As Object, ctx As Object
    Dim datumBloka As String, danas As String
    Set f = NewOtkupUIForm()
    Set zf = f.Controls("zForm")
    Set ctx = f.Controls("zCtx")

    ' Datum se izvodi iz danasnjeg, da NIKAD ne bude jednak "danas" -- zakucan
    ' datum bi jednog dana u godini prosao test i kad pravilo ne radi.
    datumBloka = Format$(Date - 30, "dd.mm.yyyy")

    ' Blok koji se upravo snimio nad aktivnom otpremnicom.
    '
    ' Datum i zbirna se postavljaju kroz ApplyPrefill, ne pisanjem u kontrolu:
    ' to je put kojim ih i produkcija dobija (izbor otpremnice), i jedini koji
    ' ide pod mLoading. Direktan upis u fgDatum okine OnDatumChanged, a on trazi
    ' stanica-lock i predlog broja SA PITANJEM GOOGLE-U -- mreza u testu.
    ' Kilogrami i ambalaza su TextBox-evi: njihova promena samo preracunava
    ' vrednost, pa idu direktno.
    modScrDokumenti.Scr_OtpTestSet FX_OTP_ID, FX_BROJ_OTP
    modOtkupUI.ApplyPrefill "datum=" & datumBloka & "|brzbirne=" & FX_ZBIRNA
    SetPolje zf, "fgKgI", "123,4"
    SetPolje zf, "fgKolAmb", "10"
    ctx.Controls("cbKupac").value = FX_KOOPERANT

    ' Preduslovi: bez njih bi test bio zelen i kad kontrole uopste ne primaju
    ' vrednost, pa ne bi merio nista.
    AssertEq Polje(zf, "fgDatum"), datumBloka, "preduslov: datum otpremnice je upisan"
    AssertEq Polje(zf, "fgBrZbir"), FX_ZBIRNA, "preduslov: broj zbirne je upisan"
    AssertEq Polje(zf, "fgKgI"), "123,4", "preduslov: kilogrami su upisani"
    AssertEq ctx.Controls("cbKupac").value, FX_KOOPERANT, "preduslov: partner je upisan"

    modOtkupUI.ClearForm

    ' 1) DATUM OSTAJE -- sledeci blok ide u niz istog datuma otpremnice.
    AssertEq Polje(zf, "fgDatum"), datumBloka, _
             "dok je otpremnica aktivna datum se NE vraca na danas"
    ' 2) BROJ ZBIRNE OSTAJE -- svi blokovi jedne otpremnice idu na istu zbirnu.
    AssertEq Polje(zf, "fgBrZbir"), FX_ZBIRNA, _
             "broj zbirne je kontekst -- ne brise se posle snimanja"
    ' 3) PARTNER SE BRISE -- sledeci unos je nov kooperant. Obrnut smer od prva
    '    dva: ovde je brisanje trazeno ponasanje.
    AssertEq ctx.Controls("cbKupac").value, "", _
             "partner mora da bude obrisan posle snimanja"
    ' ... a podaci bloka odlaze sa njim.
    AssertEq Polje(zf, "fgKgI"), "", "kilogrami se brisu posle snimanja"
    AssertEq Polje(zf, "fgKolAmb"), "", "kolicina ambalaze se brise posle snimanja"

    ' BEZ AKTIVNE OTPREMNICE datum se vraca na danas: prazno ili staro polje bi
    ' bila greska koju operater mora da ispravlja pri svakom novom dokumentu.
    modScrDokumenti.Scr_OtpOtkazi
    modOtkupUI.ApplyPrefill "datum=" & datumBloka & "|brzbirne=" & FX_ZBIRNA
    danas = Format$(Date, "dd.mm.yyyy")
    modOtkupUI.ClearForm
    AssertEq Polje(zf, "fgDatum"), danas, _
             "bez aktivne otpremnice datum se vraca na danas"

    ReleaseOtkupUIForm f
End Sub

' Novi UI bez prikaza. Gradnja se okida dodirom Controls.count, isto kao kod
' frmOtkup; .Show se NE zove -- GoFullScreen, raspored i punjenje mreze idu tek
' u UserForm_Activate, a nista od toga ovi testovi ne mere.
Private Function NewOtkupUIForm() As frmOtkupUI
    Dim f As frmOtkupUI
    Set f = New frmOtkupUI

    Dim ctlCount As Long
    ctlCount = f.Controls.count          ' bez ovoga se UserForm_Initialize ne okine

    ' UserForm_Initialize hvata pad gradnje i salje ga u OtkupUI_BuildFailed, pa
    ' greska NE stize ovamo. Bez ove provere bi svaka sledeca tvrdnja padala na
    ' "Could not find the specified object" -- pad na trazenju kontrole, a ne na
    ' ponasanju koje test meri.
    If ctlCount < 2 Then
        Err.Raise ERR_ASSERT, "modTest.NewOtkupUIForm", _
                  "frmOtkupUI nije izgradjen (kontrola: " & ctlCount & ")"
    End If

    Set NewOtkupUIForm = f
End Function

' Unload gasi formu (Terminate -> OtkupUI_FormClosed), a OtkupUI_Release pusta i
' ono sto ostaje na modulu (Btns, kes tabela, num-polja) -- inace sledeci test
' gradi ekran nad ostacima prethodnog. Aktivna otpremnica zivi u TRECEM modulu
' (modScrDokumenti) i nju OtkupUI_Release ne dira, pa se otpusta ovde.
Private Sub ReleaseOtkupUIForm(f As frmOtkupUI)
    Unload f
    modOtkupUI.OtkupUI_Release
    modScrDokumenti.Scr_OtpOtkazi
End Sub

' Polja novog UI-ja su ugnjezdena: zona -> okvir polja -> kontrola (ime + "T").
' Test se kroz to stablo krece SAM, ne kroz modOtkupUI.FldText/SetFld: rutina
' koja se testira ne sme da bude i merni instrument.
Private Function Polje(z As Object, ByVal grp As String) As String
    Polje = z.Controls(grp).Controls(grp & "T").text
End Function

Private Sub SetPolje(z As Object, ByVal grp As String, ByVal v As String)
    z.Controls(grp).Controls(grp & "T").text = v
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
