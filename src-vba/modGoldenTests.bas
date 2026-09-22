Attribute VB_Name = "modGoldenTests"

Option Explicit

'=====================================================================
' modGoldenTests -- GOLDEN POSLOVNI SCENARIJI
'
' Spec i recnik tvrdnji: docs/DOMEN/GOLDEN_SCENARIJI.md
'
' Kriterijum kvaliteta:
'   SCENARIO KOJI BI MORAO DA SE MENJA U REFAKTORU JE NAPISAN POGRESNO.
'
' IZOLACIJA ULAZNOG STANJA -- ne samo izlaznog.
'   Rollback cisti ono sto scenario OSTAVI. Ne cisti ono sto je ZATEKAO.
'   Prva verzija je koristila KOOP-TEST-1 iz fixture-a i svih pet golden-a je
'   javljalo "placeno 1000.00" iako nijedan scenario ne placa nista -- to je bio
'   zatecen avans, koji pisac otkupa automatski primeni. Izmena tog avansa
'   u fixture-u oborila bi golden a da niko nije dirao poslovanje.
'   Zato scenariji koriste SOPSTVENE identitete (KOOP-GLD-1 i dr.) koji nemaju
'   transakcionu istoriju, a GldPreduslov to i PROVERAVA pre svakog scenarija.
'
' DETERMINIZAM
'   Svaki scenario radi u transakciji i na kraju RADI ROLLBACK; snapshot se
'   uzima pre toga. Datumi su fiksni -- golden vezan za Date() bi pao sutra.
'
' GOLDEN QUERY ADAPTER
'   GldSnapshot i njegovi pomocnici su JEDINO mesto koje zna kako su podaci
'   slozeni. Za deo cinjenica produkcioni read-model ne postoji, pa se cita
'   direktno -- to je svesna granica, ne propust. Kad Zbirna postane
'   header+stavke, menja se adapter; scenario i golden fajl ostaju isti.
'=====================================================================

' Svoj kod greske: modTest-ov je Private i iz ovog modula se ne vidi
' ("Variable not defined"). Compile tada pada, a run_vba to ne vidi kao pao
' test nego kao VISENJE -- 449 s do timeout-a i Excel u [break].
Private Const GLD_ERR As Long = vbObjectError + 9501

' Fiksan datum -- golden vezan za Date() bi pao sutra.
Private Const GLD_DATUM As Date = #3/15/2026#

' SOPSTVENI identiteti, bez transakcione istorije. Ne koristiti fixture ID-eve.
Private Const GLD_KOOP As String = "KOOP-GLD-1"
Private Const GLD_STANICA As String = "STA-GLD-1"
Private Const GLD_VOZAC As String = "VOZ-GLD-1"
' Drugi vozac: zbirna se broji PO VOZACU, pa dva dokumenta sa istim brojem
' mogu legalno postojati samo ako im je vlasnik razlicit (v. G2).
Private Const GLD_VOZAC2 As String = "VOZ-GLD-2"
Private Const GLD_KUPAC As String = "KUP-GLD-1"
Private Const GLD_VRSTA As String = "TESTVOCE"
Private Const GLD_SORTA As String = "TESTSORTA"
Private Const GLD_AMB As String = "12/1"
' Kultura je FK: nov pisac je trazi i proverava da se snapshot vrsta/sorta
' slaze sa njom (S4.1f). Stari pisac je nije trazio, pa je golden fixture do
' sada nije ni imao.
Private Const GLD_KULTURA As String = "GLD-KUL-1"

' Sta je scenario napravio. Resetuje se na pocetku svakog scenarija.
Private m_Otk As Collection
Private m_Otp As Collection
Private m_Zbr As Collection
Private m_Prj As Collection
Private m_Fak As Collection

' Kljuc scenarija (BrojZbirne). Adapter po njemu CITA sistem.
Private m_Kljuc As String

Private m_Total As Long
Private m_Failed As Long
Private m_Report As String


'=====================================================================
' RUNNER
'=====================================================================

Public Sub RunGoldenSuite()
    Dim prevMode As Boolean
    prevMode = IsTestMode()
    SetTestMode True

    m_Total = 0
    m_Failed = 0
    m_Report = ""

    ' 6 (B1) i 9 (B4) su UKLONJENI: kes se nikad ne vezuje za otkupni list, pa
    ' su merili putanju koja u domenu ne postoji. Ekran otkupnog lista nema polje
    ' za novac, a modOtkupUnos salje novac:=0 uvek -- v. GOLDEN_SCENARIJI.md S8.
    ' 10 (D1) NIJE registrovan: storno otpremnice ostavlja zbirnu bez izvora i
    ' invarijanta puca. Zamrznuti "PUKLA" kao ocekivano znacilo bi da golden
    ' tvrdi kako je posle legalne operacije dozvoljeno nekonzistentno stanje --
    ' suprotno od svrhe mreze. Ceka poslovnu odluku: zabrani / kaskadiraj /
    ' rekalkulisi (GOLDEN_SCENARIJI.md S10).
    GldOne 12
    GldOne 15

    SetTestMode prevMode

    Debug.Print "GOLDEN: " & CStr(m_Total) & " ukupno, " & CStr(m_Failed) & " palo"
    Debug.Print m_Report

    If m_Failed > 0 Then
        Err.Raise GLD_ERR, "RunGoldenSuite", _
                  CStr(m_Failed) & " od " & CStr(m_Total) & " golden scenarija palo:" & _
                  vbCrLf & m_Report
    End If
End Sub

Private Function GldIme(ByVal idx As Long) As String
    Select Case idx
        Case 12: GldIme = "D3_storno_dvoklasnog_otkupa"
        Case 15: GldIme = "G2_isti_broj_dva_dokumenta"
    End Select
End Function

Private Sub GldPozovi(ByVal idx As Long)
    Select Case idx
        Case 12: Gld_D3_StornoDvoklasnogOtkupa
        Case 15: Gld_G2_IstiBrojDvaDokumenta
    End Select
End Sub

Private Sub GldOne(ByVal idx As Long)
    Dim nm As String
    Dim errDesc As String

    nm = GldIme(idx)
    m_Total = m_Total + 1

    On Error GoTo EH
    GldPozovi idx
    m_Report = m_Report & "        OK " & nm & vbCrLf
    Exit Sub

EH:
    errDesc = Err.description
    m_Failed = m_Failed + 1
    m_Report = m_Report & "        FAIL " & nm & " -- " & errDesc & vbCrLf
End Sub


'=====================================================================
' OKVIR
'=====================================================================

Private Function GldTx() As clsTransaction
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_ZBIRNA
    tx.AddTableSnapshot TBL_ZBIRNA_STAVKE
    tx.AddTableSnapshot TBL_PRIJEMNICA
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot TBL_FAKTURA_STAVKE
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_PALETA
    tx.AddTableSnapshot TBL_PALETA_STAVKA
    tx.AddTableSnapshot TBL_KOOPERANTI
    tx.AddTableSnapshot TBL_STANICE
    tx.AddTableSnapshot TBL_VOZACI
    tx.AddTableSnapshot TBL_KULTURE
    tx.AddTableSnapshot TBL_OTKUP_STAVKE
    tx.AddTableSnapshot TBL_KUPCI

    Set GldTx = tx
End Function

Private Sub GldReset()
    Set m_Otk = New Collection
    Set m_Otp = New Collection
    Set m_Zbr = New Collection
    Set m_Prj = New Collection
    Set m_Fak = New Collection

End Sub

' Maticni podaci scenarija.
'
' NIJE idempotentan i to je namerno: GldMoraDaNePostoji pada ako rezervisani
' identitet vec postoji. Svaki scenario radi u svojoj transakciji i
' rollback-uje se, pa na pocetku uvek ne postoje.
Private Sub GldSeed()
    GldSeedRed TBL_STANICE, "StanicaID", GLD_STANICA, "Naziv", "GOLDEN STANICA"
    GldSeedRed TBL_VOZACI, "VozacID", GLD_VOZAC, "Ime", "GOLDEN VOZAC"
    GldSeedRed TBL_VOZACI, "VozacID", GLD_VOZAC2, "Ime", "GOLDEN VOZAC 2"
    GldSeedRed TBL_KUPCI, "KupacID", GLD_KUPAC, "Naziv", "GOLDEN KUPAC"
    GldSeedKooperant
    GldSeedKultura
End Sub

' Kultura golden scenarija. Vrsta i sorta MORAJU biti iste kao GLD_VRSTA i
' GLD_SORTA -- nov pisac odbija dokument ciji se snapshot ne slaze sa kulturom.
Private Sub GldSeedKultura()
    Dim rowData As Variant

    GldMoraDaNePostoji TBL_KULTURE, "KulturaID", GLD_KULTURA

    rowData = GldPrazanRed(TBL_KULTURE)
    GldPolje rowData, TBL_KULTURE, "KulturaID", GLD_KULTURA
    GldPolje rowData, TBL_KULTURE, "VrstaVoca", GLD_VRSTA
    GldPolje rowData, TBL_KULTURE, "SortaVoca", GLD_SORTA
    GldPolje rowData, TBL_KULTURE, "TipAmbalaze", GLD_AMB
    GldPolje rowData, TBL_KULTURE, "Aktivan", STATUS_AKTIVAN

    If AppendRow(TBL_KULTURE, rowData) <= 0 Then
        Err.Raise GLD_ERR, "GldSeedKultura", "AppendRow nije uspeo za tblKulture"
    End If
End Sub

Private Sub GldSeedRed(ByVal tbl As String, ByVal kljucKol As String, _
                       ByVal kljuc As String, ByVal nazivKol As String, _
                       ByVal naziv As String)
    Dim rowData As Variant

    GldMoraDaNePostoji tbl, kljucKol, kljuc

    rowData = GldPrazanRed(tbl)
    GldPolje rowData, tbl, kljucKol, kljuc
    GldPolje rowData, tbl, nazivKol, naziv
    GldPolje rowData, tbl, "Aktivan", STATUS_AKTIVAN

    If AppendRow(tbl, rowData) <= 0 Then
        Err.Raise GLD_ERR, "GldSeedRed", "AppendRow nije uspeo za " & tbl
    End If
End Sub

' Rezervisani identitet NE SME da postoji pre seed-a.
'
' Ranija verzija je bila fail-open: postojeci red se prihvatao ako ga ima
' tacno jedan. Tada bi zatecen STA-GLD-1 sa drugim nazivom ili Aktivan=Ne
' usao u scenario -- ista bolest kao zatecen avans: scenario vise ne
' poseduje ceo svoj ulaz.
'
' Idempotentnost ovde NE treba: svaki scenario radi u svojoj transakciji i
' rollback-uje se, pa GLD identiteti na pocetku uvek NE postoje. Ako
' postoje, to je nalaz -- ili je prethodni rollback zakazao, ili ime nije
' vise rezervisano.
Private Sub GldMoraDaNePostoji(ByVal tbl As String, ByVal kolona As String, _
                               ByVal vrednost As String)
    Dim n As Long

    n = GldBrojRedova(tbl, kolona, vrednost)
    If n > 0 Then
        Err.Raise GLD_ERR, "GldSeed", _
                  "rezervisani identitet " & vrednost & " vec postoji u " & _
                  tbl & " (" & CStr(n) & ") -- scenario ne poseduje svoj ulaz"
    End If
End Sub

Private Sub GldSeedKooperant()
    Dim rowData As Variant

    GldMoraDaNePostoji TBL_KOOPERANTI, "KooperantID", GLD_KOOP

    rowData = GldPrazanRed(TBL_KOOPERANTI)
    GldPolje rowData, TBL_KOOPERANTI, "KooperantID", GLD_KOOP
    GldPolje rowData, TBL_KOOPERANTI, "Ime", "GOLDEN"
    GldPolje rowData, TBL_KOOPERANTI, "Prezime", "KOOPERANT"
    GldPolje rowData, TBL_KOOPERANTI, COL_KOOP_STANICA, GLD_STANICA
    GldPolje rowData, TBL_KOOPERANTI, "Aktivan", STATUS_AKTIVAN

    If AppendRow(TBL_KOOPERANTI, rowData) <= 0 Then
        Err.Raise GLD_ERR, "GldSeedKooperant", "AppendRow nije uspeo"
    End If
End Sub

' Preduslov: identitet scenarija NEMA transakcionu istoriju.
'
' Bez ove provere scenario tiho nasledjuje tudje stanje -- tacno greska zbog
' koje je prva verzija javljala "placeno 1000.00" a nista nije platila.
' Preduslov: NI identitet NI kljuc scenarija nemaju zatecenu istoriju.
'
' Dedicated master ID resava samo deo: GldZbirna cita otpremnice i prijemnice po
' BrojZbirne, pa bi zatecen red sa istim kljucem usao u rezultat i kad kooperant
' nema nijedan stari otkup.
Private Sub GldPreduslov(ByVal kljuc As String)
    Dim n As Long

    GldNemaZatecenog TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, kljuc
    GldNemaZatecenog TBL_ZBIRNA, COL_ZBR_BROJ, kljuc
    GldNemaZatecenog TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, kljuc

    n = GldBrojRedova(TBL_OTKUP, COL_OTK_KOOPERANT, GLD_KOOP)
    If n > 0 Then
        Err.Raise GLD_ERR, "GldPreduslov", _
                  "kooperant " & GLD_KOOP & " vec ima " & CStr(n) & _
                  " otkupa -- scenario bi merio zatecen state"
    End If

    n = GldBrojRedova(TBL_NOVAC, COL_NOV_KOOP_ID, GLD_KOOP)
    If n > 0 Then
        Err.Raise GLD_ERR, "GldPreduslov", _
                  "kooperant " & GLD_KOOP & " vec ima " & CStr(n) & _
                  " redova u novcu (avans?) -- scenario bi merio zatecen state"
    End If
End Sub

Private Sub GldNemaZatecenog(ByVal tbl As String, ByVal kolona As String, _
                             ByVal vrednost As String)
    Dim n As Long

    n = GldBrojRedova(tbl, kolona, vrednost)
    If n > 0 Then
        Err.Raise GLD_ERR, "GldPreduslov", _
                  tbl & " vec ima " & CStr(n) & " redova sa " & kolona & "=" & _
                  vrednost & " -- scenario bi merio zatecen state"
    End If
End Sub


'=====================================================================
' GOLDEN QUERY ADAPTER
'
' Jedino mesto koje zna kako su podaci slozeni. Menja se u cutover-ima;
' scenariji i
' golden fajlovi ostaju isti.
'=====================================================================

Private Function GldSnapshot(ByVal naslov As String, ByVal brojZbirne As String) As String
    Dim s As String

    s = "== " & naslov & " ==" & vbLf
    s = s & GldDokumenti()
    s = s & GldOtkupi()
    If Len(brojZbirne) > 0 Then s = s & GldZbirna(brojZbirne)
    s = s & GldStatus()
    s = s & GldAmbalaza()
    s = s & GldNovacAlokacija()
    s = s & GldFakture()

    GldSnapshot = s
End Function

' Novac: koliko je isplaceno, koliko je jos VEZANO za otkupe scenarija, koliko
' je ostalo nealokirano.
'
' Storno otkupa NE ponistava novcanu transakciju -- covek je fizicki dobio pare.
' StornoOtkup radi ResetNovacOtkupLink: iznos ostaje u istoriji, ali prestaje da
' bude alociran na taj otkup. Bez ova tri broja golden to ne bi video, pa bi
' ostao zelen i kad bi neko obrisao ResetNovacOtkupLink.
'
' Ovaj oblik prezivi i buduci tblNovacAlokacije model.
' Koja je zbirna stornirana, a koja je ostala -- po REDOSLEDU nastanka.
'
' Dva razloga zasto G2 ne sme da koristi obicnu ZBIRNA sekciju:
'
'   1. Ona racuna preko SumOtpremniceByKlasa(BrojZbirne), a broj sam ne
'      razlikuje dve logicke zbirne -- pa bi brojke dva dokumenta bile
'      ZAKLJUCANE kao ocekivane. Posle Zbirna cutover-a taj adapter prima
'      ZbirnaID, pa bi ispravka arhitekture oborila golden koji je treba da
'      stiti.
'   2. "aktivnih 1 / storniranih 1" ne kaze KOJA je stornirana. Bug koji
'      stornira drugu umesto prve ostavlja iste brojeve i test ostaje zelen.
'
' Zato se izvestava kolicina i status po poziciji. ID se NE ispisuje -- ni
' danasnji ni buduci. Posle Zbirna cutover-a adapter trazi red po ZbirnaID; golden isti.
Private Function GldIdentitetZbirni() As String
    Dim data As Variant
    Dim cID As Long, cKol As Long, cSt As Long
    Dim i As Long, k As Long
    Dim s As String
    Dim oznaka As String
    Dim nadjen As Boolean

    s = "IDENTITET ZBIRNE" & vbLf
    If m_Zbr.count = 0 Then
        GldIdentitetZbirni = s & "  nema" & vbLf
        Exit Function
    End If

    data = GetTableData(TBL_ZBIRNA)
    If Not IsArray(data) Then
        GldIdentitetZbirni = s & "  nema" & vbLf
        Exit Function
    End If

    cID = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_ID, "GldIdentitetZbirni")
    cKol = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KOLICINA, "GldIdentitetZbirni")
    cSt = RequireColumnIndex(TBL_ZBIRNA, COL_STORNIRANO, "GldIdentitetZbirni")

    For k = 1 To m_Zbr.count
        Select Case k
            Case 1: oznaka = "prva "
            Case 2: oznaka = "druga"
            Case Else: oznaka = "br." & CStr(k)
        End Select

        nadjen = False
        For i = 1 To UBound(data, 1)
            If StrComp(Trim$(NzToText(data(i, cID))), Trim$(CStr(m_Zbr(k))), _
                       vbTextCompare) = 0 Then
                s = s & "  " & oznaka & "  kg=" & Fmt2(SafeD(data(i, cKol))) & _
                    "  " & IIf(UCase$(Trim$(NzToText(data(i, cSt)))) = "DA", _
                               "STORNIRANA", "AKTIVNA") & vbLf
                nadjen = True
                Exit For
            End If
        Next i
        If Not nadjen Then s = s & "  " & oznaka & "  NEMA REDA" & vbLf
    Next k

    GldIdentitetZbirni = s
End Function

Private Function GldNovacAlokacija() As String
    Dim data As Variant
    Dim cKoop As Long, cIspl As Long, cOtk As Long
    Dim i As Long
    Dim ukupno As Double, vezano As Double
    Dim id As String

    data = GetTableData(TBL_NOVAC)
    If IsArray(data) Then
        data = ExcludeStornirano(data, TBL_NOVAC)
        If IsArray(data) Then
            cKoop = RequireColumnIndex(TBL_NOVAC, COL_NOV_KOOP_ID, "GldNovacAlokacija")
            cIspl = RequireColumnIndex(TBL_NOVAC, COL_NOV_ISPLATA, "GldNovacAlokacija")
            cOtk = RequireColumnIndex(TBL_NOVAC, COL_NOV_OTKUP_ID, "GldNovacAlokacija")
            For i = 1 To UBound(data, 1)
                If StrComp(Trim$(NzToText(data(i, cKoop))), GLD_KOOP, vbTextCompare) = 0 Then
                    ukupno = ukupno + SafeD(data(i, cIspl))
                    id = Trim$(NzToText(data(i, cOtk)))
                    If Len(id) > 0 Then vezano = vezano + SafeD(data(i, cIspl))
                End If
            Next i
        End If
    End If

    GldNovacAlokacija = "NOVAC" & vbLf & _
        "  ukupno isplaceno " & Fmt2(ukupno) & vbLf & _
        "  vezano za otkup  " & Fmt2(vezano) & vbLf & _
        "  nealocirano      " & Fmt2(ukupno - vezano) & vbLf
End Function

' Koliko je dokumenata AKTIVNO, a koliko STORNIRANO -- u opsegu scenarija.
'
' Storniran dokument ostaje u tabeli i izlazi iz agregata (README S2), pa
' "aktivnih 0 / storniranih 1" jeste poslovna cinjenica, ne broj redova.
' Broji se isto kao DOKUMENTI: distinct logicki identitet.
Private Function GldStatus() As String
    GldStatus = "STATUS" & vbLf & _
        "  otkup           aktivnih " & CStr(GldBrojDok(TBL_OTKUP, COL_OTK_KOOPERANT, _
            GLD_KOOP, COL_OTK_BR_DOK)) & "  storniranih " & _
            CStr(GldBrojStorno(TBL_OTKUP, COL_OTK_KOOPERANT, GLD_KOOP, COL_OTK_BR_DOK)) & vbLf & _
        "  otpremnica      aktivnih " & CStr(GldBrojDok(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, _
            m_Kljuc, COL_OTP_BROJ)) & "  storniranih " & _
            CStr(GldBrojStorno(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, m_Kljuc, COL_OTP_BROJ)) & vbLf & _
        "  zbirna          aktivnih " & CStr(GldBrojDok(TBL_ZBIRNA, COL_ZBR_BROJ, _
            m_Kljuc, COL_ZBR_BROJ)) & "  storniranih " & _
            CStr(GldBrojStorno(TBL_ZBIRNA, COL_ZBR_BROJ, m_Kljuc, COL_ZBR_BROJ)) & vbLf & _
        "  prijemnica      aktivnih " & CStr(GldBrojDok(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, _
            m_Kljuc, COL_PRJ_BROJ)) & "  storniranih " & _
            CStr(GldBrojStorno(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, m_Kljuc, COL_PRJ_BROJ)) & vbLf
End Function

' Ambalazni saldo -- IZVEDEN iz ledgera, nikad iz kolone (README S2).
Private Function GldAmbalaza() As String
    GldAmbalaza = "AMBALAZA" & vbLf & _
        "  kooperant " & GLD_AMB & "  " & CStr(GldAmbSaldo(GLD_KOOP, "Kooperant")) & vbLf
End Function

Private Function GldAmbSaldo(ByVal entID As String, ByVal entTip As String) As Long
    Dim st As Variant
    Dim i As Long

    st = GetAmbalazeStanje(entID, entTip)
    If Not IsArray(st) Then Exit Function

    For i = 1 To UBound(st, 1)
        If StrComp(Trim$(NzToText(st(i, 1))), GLD_AMB, vbTextCompare) = 0 Then
            If IsNumeric(st(i, 2)) Then GldAmbSaldo = CLng(st(i, 2))
            Exit Function
        End If
    Next i
End Function

' Isti obracun kao GldBrojDok, ali nad STORNIRANIM redovima.
Private Function GldBrojStorno(ByVal tbl As String, ByVal scopeKol As String, _
                               ByVal scopeVal As String, ByVal brojKol As String) As Long
    Dim data As Variant
    Dim cScope As Long, cBroj As Long, cGen As Long, cSt As Long
    Dim i As Long
    Dim d As Object
    Dim k As String

    If Len(scopeVal) = 0 Then Exit Function

    data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Function

    cScope = RequireColumnIndex(tbl, scopeKol, "GldBrojStorno")
    cBroj = RequireColumnIndex(tbl, brojKol, "GldBrojStorno")
    cSt = RequireColumnIndex(tbl, COL_STORNIRANO, "GldBrojStorno")
    cGen = GetColumnIndex(tbl, COL_GENERACIJA_ID)

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    For i = 1 To UBound(data, 1)
        If StrComp(Trim$(NzToText(data(i, cScope))), scopeVal, vbTextCompare) = 0 Then
            If UCase$(Trim$(NzToText(data(i, cSt)))) = "DA" Then
                k = ""
                If cGen > 0 Then k = Trim$(NzToText(data(i, cGen)))
                If Len(k) = 0 Then k = "BR:" & Trim$(NzToText(data(i, cBroj)))
                If Not d.Exists(k) Then d.Add k, True
            End If
        End If
    Next i

    GldBrojStorno = d.count
End Function

' Broj LOGICKIH dokumenata koje je scenario napravio.
'
' Ovo je poslovna cinjenica ("dva bloka su otisla na jednu otpremnicu"), ne
' broj fizickih redova -- pa preziva header+stavke nepromenjeno. Bez nje su
' A3 i A4 zeleni i kad bi bug napravio dve otpremnice po 500 kg: agregat je
' isti, a kardinalnost nije.
Private Function GldDokumenti() As String
    GldDokumenti = "DOKUMENTI" & vbLf & _
        "  otkupa          " & CStr(GldBrojDok(TBL_OTKUP, COL_OTK_KOOPERANT, _
                                   GLD_KOOP, COL_OTK_BR_DOK)) & vbLf & _
        "  otpremnica      " & CStr(GldBrojDok(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, _
                                   m_Kljuc, COL_OTP_BROJ)) & vbLf & _
        "  zbirnih         " & CStr(GldBrojDok(TBL_ZBIRNA, COL_ZBR_BROJ, _
                                   m_Kljuc, COL_ZBR_BROJ)) & vbLf & _
        "  prijemnica      " & CStr(GldBrojDok(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, _
                                   m_Kljuc, COL_PRJ_BROJ)) & vbLf & _
        "  faktura         " & CStr(GldBrojFaktura()) & vbLf
End Function

' Koliko LOGICKIH dokumenata sistem STVARNO ima, u opsegu scenarija.
'
' Prva verzija je brojala pozive koje je test sam izvrsio (m_nOtp = m_nOtp + 1).
' To je tautologija: A3 je dokazivao "test je jednom pozvao GldOtpremnica", ne
' "sistem je napravio jednu otpremnicu". Bug koji od jednog poziva napravi dve
' otpremnice po 500 kg ostavio bi agregat isti i test ZELEN -- bas kvar zbog
' kojeg je sekcija i dodata.
'
' Identitet dokumenta danas: GeneracijaID ako ga red nosi, inace poslovni broj.
' modOtkup NE pise GeneracijaID, pa se otkup broji po BrojDokumenta; otpremnica,
' zbirna i prijemnica ga imaju. Posle Otkup cutover-a sve postaje COUNT(DISTINCT <Doc>ID) --
' menja se OVAJ adapter, golden ostaje isti.
Private Function GldBrojDok(ByVal tbl As String, ByVal scopeKol As String, _
                            ByVal scopeVal As String, ByVal brojKol As String) As Long
    Dim data As Variant
    Dim cScope As Long, cBroj As Long, cGen As Long
    Dim i As Long
    Dim d As Object
    Dim k As String

    If Len(scopeVal) = 0 Then Exit Function

    data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Function
    data = ExcludeStornirano(data, tbl)
    If IsEmpty(data) Then Exit Function

    cScope = RequireColumnIndex(tbl, scopeKol, "GldBrojDok")
    cBroj = RequireColumnIndex(tbl, brojKol, "GldBrojDok")
    cGen = GetColumnIndex(tbl, COL_GENERACIJA_ID)

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    For i = 1 To UBound(data, 1)
        If StrComp(Trim$(NzToText(data(i, cScope))), scopeVal, vbTextCompare) = 0 Then
            k = ""
            If cGen > 0 Then k = Trim$(NzToText(data(i, cGen)))
            If Len(k) = 0 Then k = "BR:" & Trim$(NzToText(data(i, cBroj)))
            If Not d.Exists(k) Then d.Add k, True
        End If
    Next i

    GldBrojDok = d.count
End Function

' Faktura nema prirodan kljuc opsega, pa se broje one koje je scenario napravio.
' Faktura se ne deli po klasama, pa ID i dokument jesu isto.
Private Function GldBrojFaktura() As Long
    Dim d As Object
    Dim i As Long
    Dim k As String

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    For i = 1 To m_Fak.count
        k = Trim$(CStr(m_Fak(i)))
        If Len(k) > 0 Then
            If Not d.Exists(k) Then d.Add k, True
        End If
    Next i

    GldBrojFaktura = d.count
End Function

Private Function GldOtkupi() As String
    Dim data As Variant
    Dim cKlasa As Long, cKol As Long, cCena As Long, cIspl As Long, cID As Long
    Dim i As Long, k As Long
    Dim kolI As Double, kolII As Double, vrednost As Double, placeno As Double
    Dim svePlaceno As Boolean
    Dim imaAktivnih As Boolean
    Dim trazeni As Object
    Dim id As String
    Dim s As String

    svePlaceno = True

    Set trazeni = CreateObject("Scripting.Dictionary")
    trazeni.CompareMode = vbTextCompare
    For k = 1 To m_Otk.count
        id = Trim$(CStr(m_Otk(k)))
        If Len(id) > 0 Then
            If Not trazeni.Exists(id) Then trazeni.Add id, True
        End If
    Next k

    s = "OTKUP" & vbLf
    If trazeni.count = 0 Then
        GldOtkupi = s & "  nema" & vbLf
        Exit Function
    End If

    data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then
        GldOtkupi = s & "  nema" & vbLf
        Exit Function
    End If
    data = ExcludeStornirano(data, TBL_OTKUP)
    If IsEmpty(data) Then
        GldOtkupi = s & "  nema" & vbLf
        Exit Function
    End If

    cID = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, "GldOtkupi")

    ' Klasa/Kolicina/Cena se citaju sa STAVKI, ne sa headera. Header ih u ciljnoj
    ' semi nema (S4.1), a nov pisac ih vec ostavlja prazne -- oracle koji bi ih i
    ' dalje citao merio bi nulu i zvao je istinom.
    Dim vrednostReda As Double

    For i = 1 To UBound(data, 1)
        id = Trim$(NzToText(data(i, cID)))
        If trazeni.Exists(id) Then
            vrednostReda = GldStavkeOtkupa(id, kolI, kolII)
            vrednost = vrednost + vrednostReda
            placeno = placeno + GetIsplataForOtkup(id)
            imaAktivnih = True

            ' ISPLACENOST IZ KNJIGE, ne iz kolone. Isplaceno u ciljnom modelu nije
            ' polje nego izvedeno stanje: placeno = SUM(tblNovac po OtkupID)
            ' (S4.1c, S6.1). Kolona i knjiga su mogle da se raziju -- bez kolone
            ' nema sta da se ne slaze.
            If GetIsplataForOtkup(id) + 0.0001 < vrednostReda Then svePlaceno = False
        End If
    Next i

    ' Bez ijednog aktivnog reda "isplaceno svi DA" bi tvrdilo da je sve placeno
    ' kad nema sta da se plati -- v. D3 posle storna. Prazno stanje se kaze kao
    ' prazno.
    If Not imaAktivnih Then svePlaceno = False

    s = s & "  predao          I=" & Fmt2(kolI) & "  II=" & Fmt2(kolII) & vbLf
    s = s & "  vrednost        " & Fmt2(vrednost) & vbLf
    s = s & "  placeno         " & Fmt2(placeno) & vbLf
    s = s & "  od toga avansom " & Fmt2(GldNovacPoTipu(NOV_VIRMAN_AVANS_KOOP)) & vbLf
    s = s & "  od toga kesom   " & Fmt2(GldNovacPoTipu(NOV_KES_OTKUPAC_KOOP)) & vbLf
    s = s & "  isplaceno svi   " & IIf(svePlaceno, "DA", "NE") & vbLf

    GldOtkupi = s
End Function

' Stavke jednog otkupa: dopunjuje kolicine po klasi i vraca vrednost dokumenta.
'
' Vrednost je SUM(Kolicina x Cena) po stavkama -- dokument vise nema jednu cenu,
' pa se ni vrednost ne moze procitati sa headera.
Private Function GldStavkeOtkupa(ByVal otkupID As String, _
                                 ByRef kolI As Double, _
                                 ByRef kolII As Double) As Double
    Dim d As Variant
    d = GetTableData(TBL_OTKUP_STAVKE)
    If Not IsArray(d) Then Exit Function

    Dim cOtk As Long, cKlasa As Long, cKol As Long, cCena As Long
    cOtk = RequireColumnIndex(TBL_OTKUP_STAVKE, COL_OKS_OTKUP_ID, "GldStavkeOtkupa")
    cKlasa = RequireColumnIndex(TBL_OTKUP_STAVKE, COL_OKS_KLASA, "GldStavkeOtkupa")
    cKol = RequireColumnIndex(TBL_OTKUP_STAVKE, COL_OKS_KOLICINA, "GldStavkeOtkupa")
    cCena = RequireColumnIndex(TBL_OTKUP_STAVKE, COL_OKS_CENA, "GldStavkeOtkupa")

    Dim i As Long, kol As Double
    For i = 1 To UBound(d, 1)
        If StrComp(Trim$(NzToText(d(i, cOtk))), otkupID, vbTextCompare) = 0 Then
            kol = SafeD(d(i, cKol))
            If UCase$(Trim$(NzToText(d(i, cKlasa)))) = "II" Then
                kolII = kolII + kol
            Else
                kolI = kolI + kol
            End If
            GldStavkeOtkupa = GldStavkeOtkupa + kol * SafeD(d(i, cCena))
        End If
    Next i
End Function

' Broj se koristi zato sto ga DANAS trazi SumOtpremniceByKlasa -- zatecen API,
' ne izbor scenarija. Posle Zbirna cutover-a prima ZbirnaID, pa se menja OVAJ
' adapter.
'
' RED "invarijanta" JE SKINUT U S4-3a. Merio je zaglavlje tblZbirna naspram
' zbira otpremnica po BrojZbirne -- a kanonski pisac ne upisuje ni jedno ni
' drugo, pa su obe strane bile nule i red je stajao "OK" nad svakom zbirnom.
' Golden koji uvek kaze OK ne meri nista, a izgleda kao da meri.
Private Function GldZbirna(ByVal brojZbirne As String) As String
    Dim sumOtp As Object
    Dim s As String
    Dim prijI As Double, prijII As Double
    Dim data As Variant
    Dim cBr As Long, cKlasa As Long, cKol As Long
    Dim i As Long

    s = "ZBIRNA" & vbLf

    Set sumOtp = SumOtpremniceByKlasa(brojZbirne)
    s = s & "  poslato         I=" & Fmt2(GldDict(sumOtp, "kgI")) & _
            "  II=" & Fmt2(GldDict(sumOtp, "kgII")) & vbLf

    data = GetTableData(TBL_PRIJEMNICA)
    If IsArray(data) Then
        data = ExcludeStornirano(data, TBL_PRIJEMNICA)
        If IsArray(data) Then
            cBr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, "GldZbirna")
            cKlasa = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KLASA, "GldZbirna")
            cKol = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA, "GldZbirna")
            For i = 1 To UBound(data, 1)
                If StrComp(Trim$(NzToText(data(i, cBr))), brojZbirne, vbTextCompare) = 0 Then
                    If UCase$(Trim$(NzToText(data(i, cKlasa)))) = "II" Then
                        prijII = prijII + SafeD(data(i, cKol))
                    Else
                        prijI = prijI + SafeD(data(i, cKol))
                    End If
                End If
            Next i
        End If
    End If

    s = s & "  primljeno       I=" & Fmt2(prijI) & "  II=" & Fmt2(prijII) & vbLf
    s = s & "  kalo            I=" & Fmt2(GldDict(sumOtp, "kgI") - prijI) & _
            "  II=" & Fmt2(GldDict(sumOtp, "kgII") - prijII) & vbLf

    GldZbirna = s
End Function

' Samo fakture koje je scenario napravio.
Private Function GldFakture() As String
    Dim data As Variant
    Dim cID As Long, cIznos As Long
    Dim i As Long, k As Long
    Dim iznos As Double
    Dim trazeni As Object
    Dim id As String

    If m_Fak.count = 0 Then
        GldFakture = "FAKTURA" & vbLf & "  nema" & vbLf
        Exit Function
    End If

    Set trazeni = CreateObject("Scripting.Dictionary")
    trazeni.CompareMode = vbTextCompare
    For k = 1 To m_Fak.count
        id = Trim$(CStr(m_Fak(k)))
        If Len(id) > 0 Then
            If Not trazeni.Exists(id) Then trazeni.Add id, True
        End If
    Next k

    data = GetTableData(TBL_FAKTURE)
    If IsArray(data) Then
        data = ExcludeStornirano(data, TBL_FAKTURE)
        If IsArray(data) Then
            cID = RequireColumnIndex(TBL_FAKTURE, COL_FAK_ID, "GldFakture")
            cIznos = RequireColumnIndex(TBL_FAKTURE, COL_FAK_IZNOS, "GldFakture")
            For i = 1 To UBound(data, 1)
                If trazeni.Exists(Trim$(NzToText(data(i, cID)))) Then
                    iznos = iznos + SafeD(data(i, cIznos))
                End If
            Next i
        End If
    End If

    GldFakture = "FAKTURA" & vbLf & _
        "  iznos           " & Fmt2(iznos) & vbLf & _
        "  osirocena       " & GldFakturaOsirocena() & vbLf & _
        "  osirocenih stavki " & CStr(GldOsirocenihStavki()) & vbLf
End Function

' Izdata faktura se posle storna svoje prijemnice NE brise -- markira se kao
' osirocena (COL_OSIROCENO_OD) i trazi korekciju. Bez ovoga bi D2 ostao zelen i
' kad bi neko obrisao MarkFakturaOrphaned / MarkFakturaStavkeOrphaned.
Private Function GldFakturaOsirocena() As String
    Dim data As Variant
    Dim cID As Long, cOs As Long
    Dim i As Long, k As Long

    GldFakturaOsirocena = "-"
    If m_Fak.count = 0 Then Exit Function

    data = GetTableData(TBL_FAKTURE)
    If Not IsArray(data) Then Exit Function

    cID = RequireColumnIndex(TBL_FAKTURE, COL_FAK_ID, "GldFakturaOsirocena")
    cOs = GetColumnIndex(TBL_FAKTURE, COL_OSIROCENO_OD)
    If cOs = 0 Then Exit Function

    GldFakturaOsirocena = "NE"
    For i = 1 To UBound(data, 1)
        For k = 1 To m_Fak.count
            If StrComp(Trim$(NzToText(data(i, cID))), Trim$(CStr(m_Fak(k))), _
                       vbTextCompare) = 0 Then
                If Len(Trim$(NzToText(data(i, cOs)))) > 0 Then
                    GldFakturaOsirocena = "DA"
                    Exit Function
                End If
            End If
        Next k
    Next i
End Function

Private Function GldOsirocenihStavki() As Long
    Dim data As Variant
    Dim cFak As Long, cOs As Long
    Dim i As Long, k As Long

    If m_Fak.count = 0 Then Exit Function

    data = GetTableData(TBL_FAKTURA_STAVKE)
    If Not IsArray(data) Then Exit Function

    cFak = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_FAKTURA_ID, "GldOsirocenihStavki")
    cOs = GetColumnIndex(TBL_FAKTURA_STAVKE, COL_OSIROCENO_OD)
    If cOs = 0 Then Exit Function

    For i = 1 To UBound(data, 1)
        If Len(Trim$(NzToText(data(i, cOs)))) > 0 Then
            For k = 1 To m_Fak.count
                If StrComp(Trim$(NzToText(data(i, cFak))), Trim$(CStr(m_Fak(k))), _
                           vbTextCompare) = 0 Then
                    GldOsirocenihStavki = GldOsirocenihStavki + 1
                End If
            Next k
        End If
    Next i
End Function


'=====================================================================
' SITNI CITACI
'=====================================================================

' Koliko je isplaceno kooperantu scenarija PO TIPU novca.
'
' Bez ovoga B2 (avans 20000) i B3 (kes 20000) daju identican golden, pa nijedan
' ne dokazuje SVOJ mehanizam -- "koliko avansom, koliko gotovinom" je i u
' recniku dozvoljenih tvrdnji (GOLDEN_SCENARIJI.md S2).
Private Function GldNovacPoTipu(ByVal tip As String) As Double
    Dim data As Variant
    Dim cKoop As Long, cTip As Long, cIspl As Long
    Dim i As Long

    data = GetTableData(TBL_NOVAC)
    If IsEmpty(data) Then Exit Function
    data = ExcludeStornirano(data, TBL_NOVAC)
    If IsEmpty(data) Then Exit Function

    cKoop = RequireColumnIndex(TBL_NOVAC, COL_NOV_KOOP_ID, "GldNovacPoTipu")
    cTip = RequireColumnIndex(TBL_NOVAC, COL_NOV_TIP, "GldNovacPoTipu")
    cIspl = RequireColumnIndex(TBL_NOVAC, COL_NOV_ISPLATA, "GldNovacPoTipu")

    For i = 1 To UBound(data, 1)
        If StrComp(Trim$(NzToText(data(i, cKoop))), GLD_KOOP, vbTextCompare) = 0 Then
            If StrComp(Trim$(NzToText(data(i, cTip))), tip, vbTextCompare) = 0 Then
                GldNovacPoTipu = GldNovacPoTipu + SafeD(data(i, cIspl))
            End If
        End If
    Next i
End Function

Private Function GldDict(ByVal d As Object, ByVal k As String) As Double
    If d Is Nothing Then Exit Function
    If d.Exists(k) Then GldDict = SafeD(d(k))
End Function

Private Function SafeD(ByVal v As Variant) As Double
    If IsNumeric(v) Then SafeD = CDbl(v)
End Function

' Decimalna TACKA nezavisno od lokala. Format$ postuje regionalna podesavanja,
' pa bi golden snimljen na masini sa zarezom pao na masini sa tackom.
Private Function Fmt2(ByVal v As Double) As String
    Fmt2 = Replace(Format$(v, "0.00"), ",", ".")
End Function

Private Function GldPrazanRed(ByVal tbl As String) As Variant
    Dim lo As ListObject
    Dim arr() As Variant

    Set lo = GetTable(tbl)
    If lo Is Nothing Then
        Err.Raise GLD_ERR, "GldPrazanRed", "Tabela nije nadjena: " & tbl
    End If

    ReDim arr(1 To lo.ListColumns.count)
    GldPrazanRed = arr
End Function

Private Sub GldPolje(ByRef rowData As Variant, ByVal tbl As String, _
                     ByVal kolona As String, ByVal vrednost As Variant)
    Dim idx As Long
    idx = GetColumnIndex(tbl, kolona)
    If idx > 0 Then rowData(idx) = vrednost
End Sub

Private Function GldBrojRedova(ByVal tbl As String, ByVal kolona As String, _
                               ByVal vrednost As String) As Long
    Dim data As Variant
    Dim idx As Long
    Dim i As Long

    If GetTable(tbl) Is Nothing Then Exit Function

    idx = GetColumnIndex(tbl, kolona)
    If idx = 0 Then Exit Function

    data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Function

    For i = 1 To UBound(data, 1)
        If StrComp(Trim$(NzToText(data(i, idx))), vrednost, vbTextCompare) = 0 Then
            GldBrojRedova = GldBrojRedova + 1
        End If
    Next i
End Function


'=====================================================================
' POSLOVNI POTEZI -- tanki omotaci nad PRODUKCIONIM writer-ima
'=====================================================================

' Razlaze "OTK-1 + OTK-2". Taj format uklanja Otkup cutover; tada se svodi na
' jedan Add, a golden se NE menja.
Private Sub GldDodaj(ByVal cilj As Collection, ByVal rez As String)
    Dim delovi() As String
    Dim i As Long
    Dim t As String

    If Len(Trim$(rez)) = 0 Then Exit Sub

    delovi = Split(rez, " + ")
    For i = LBound(delovi) To UBound(delovi)
        t = Trim$(delovi(i))
        If Len(t) > 0 Then cilj.Add t
    Next i
End Sub

' Otkup golden scenarija -- KANONSKI pisac.
'
' brojZbirne se vise NE prosledjuje: broj NE SME da bude veza nego labela (A2) --
' pravilo, ne opis zatecenog stanja (BrojZbirne je i danas join kljuc u
' modDokumenta) -- i nov pisac ga ne prima. Golden ga nigde ne cita (mereno: 0 pojava
' COL_OTK_BROJ_ZBIRNE u ovom modulu), pa nijedan scenario time ne gubi tvrdnju.
'
' VozacID takodje odlazi -- vozac pripada otpremnici (S4.1c).
Private Sub GldOtkup(ByVal brojZbirne As String, ByVal brDok As String, _
                     ByVal kolI As Double, ByVal cenaI As Double, _
                     ByVal kolII As Double, ByVal cenaII As Double)
    Dim h As Object
    Set h = CreateObject("Scripting.Dictionary")
    h.Add "Datum", GLD_DATUM
    h.Add "KooperantID", GLD_KOOP
    h.Add "StanicaID", GLD_STANICA
    h.Add "KulturaID", GLD_KULTURA
    h.Add "VrstaVoca", GLD_VRSTA
    h.Add "SortaVoca", GLD_SORTA
    h.Add "TipAmbalaze", GLD_AMB
    h.Add "BrojDokumenta", brDok

    ' Ambalaza: 50 gajbi na Klasu I, NULA na Klasu II -- tacno onako kako je stari
    ' poziv radio (kolAmb=50, kolAmbII neprosledjen pa 0). Prva verzija ovog prelaza
    ' je stavila 50 na SVAKU klasu i golden je to odmah prijavio: saldo kooperanta
    ' -50 postao -100. Podatak scenarija se ne sme menjati usput.
    Dim stavke As Collection
    Set stavke = New Collection
    If kolI > 0 Then stavke.Add GldStavka(KLASA_I, kolI, cenaI, 50#)
    If kolII > 0 Then stavke.Add GldStavka(KLASA_II, kolII, cenaII, 0#)

    Dim greska As String
    Dim res As String
    res = CreateOtkup_TX(h, stavke, greska)
    If Len(res) = 0 Then
        Err.Raise GLD_ERR, "GldOtkup", "CreateOtkup_TX nije vratio ID: " & greska
    End If

    m_Otk.Add res
End Sub

Private Function GldStavka(ByVal klasa As String, ByVal kol As Double, _
                           ByVal cena As Double, ByVal amb As Double) As Object
    Dim s As Object
    Set s = CreateObject("Scripting.Dictionary")
    s.Add "Klasa", klasa
    s.Add "Kolicina", kol
    s.Add "Cena", cena
    s.Add "KolAmbalaze", amb
    Set GldStavka = s
End Function

' Avans kooperantu -- scenario ga pravi SAM, jer je zatecen avans bio prva
' velika rupa u izolaciji (GOLDEN_SCENARIJI.md S1b).
Private Sub GldAvans(ByVal iznos As Double)
    If Len(SaveNovac_TX("GLD-AV", GLD_DATUM, "GOLDEN KOOPERANT", GLD_KOOP, _
            "Kooperant", GLD_STANICA, GLD_KOOP, "", GLD_VRSTA, _
            NOV_VIRMAN_AVANS_KOOP, 0#, iznos)) = 0 Then
        Err.Raise GLD_ERR, "GldAvans", "avans nije snimljen"
    End If
End Sub

' Storno CELOG otkupnog bloka po poslovnom broju -- danas je to jedini ulaz koji
' zahvati obe klase. Otkup cutover ga zamenjuje storno-om po DocumentID.
' Zbirna za zadatog vozaca -- G2 pravi dve sa istim brojem.
'
' Od S4-2c/2a je ovo SEED, ne pisac. Stari pisac je obrisan, a kanonski nacrt
' NAMERNO ostavlja Kolicina i Klasa na zaglavlju prazne -- bas te kolone cita
' GldIdentitetZbirni, pa bi golden izlaz postao "kg=0.00" i sidro bi se pomerilo
' bez ijedne promene u ponasanju koje ovaj scenario meri.
'
' Seed zato pravi ZATECENI oblik reda, isti koji je pisac pravio, i pecati
' generaciju istim scope-om (broj + vozac + kupac) -- GldStornoZbirne je cita iz
' reda. I seed i citaoci zaglavlja nestaju zajedno u S3e-2.
Private Sub GldZbirnaZaVozaca(ByVal broj As String, ByVal vozac As String, _
                              ByVal kolI As Double)
    Dim zbirnaID As String
    zbirnaID = GetNextID(TBL_ZBIRNA, COL_ZBR_ID, "ZBR-")
    If Len(zbirnaID) = 0 Then
        Err.Raise GLD_ERR, "GldZbirnaZaVozaca", "zbirna nije snimljena"
    End If

    Dim lo As ListObject
    Set lo = GetTable(TBL_ZBIRNA)
    If lo Is Nothing Then
        Err.Raise GLD_ERR, "GldZbirnaZaVozaca", "tblZbirna ne postoji"
    End If

    Dim nr As ListRow
    Set nr = lo.ListRows.Add

    GldCelija nr, TBL_ZBIRNA, COL_ZBR_ID, zbirnaID
    GldCelija nr, TBL_ZBIRNA, COL_ZBR_DATUM, GLD_DATUM
    GldCelija nr, TBL_ZBIRNA, COL_ZBR_VOZAC, vozac
    GldCelija nr, TBL_ZBIRNA, COL_ZBR_BROJ, broj
    GldCelija nr, TBL_ZBIRNA, COL_ZBR_KUPAC, GLD_KUPAC
    GldCelija nr, TBL_ZBIRNA, COL_ZBR_VRSTA, GLD_VRSTA
    GldCelija nr, TBL_ZBIRNA, COL_ZBR_SORTA, GLD_SORTA
    GldCelija nr, TBL_ZBIRNA, COL_ZBR_KOLICINA, kolI
    GldCelija nr, TBL_ZBIRNA, COL_ZBR_TIP_AMB, GLD_AMB
    GldCelija nr, TBL_ZBIRNA, COL_ZBR_KOL_AMB, 50
    GldCelija nr, TBL_ZBIRNA, COL_ZBR_KLASA, KLASA_I

    ' Red se trazi PO ID-u, ne preko nr.Index: ApplyGeneracijaID ocekuje indeks u
    ' istoj konvenciji koju vraca AppendRow, a ta dva se ne moraju poklapati.
    Dim redovi As Collection
    Set redovi = FindRows(TBL_ZBIRNA, COL_ZBR_ID, zbirnaID)
    If redovi Is Nothing Then
        Err.Raise GLD_ERR, "GldZbirnaZaVozaca", "zbirna nije nadjena posle upisa"
    End If
    If redovi.count <> 1 Then
        Err.Raise GLD_ERR, "GldZbirnaZaVozaca", "zbirna nije nadjena tacno jednom"
    End If

    ApplyGeneracijaID TBL_ZBIRNA, CLng(redovi(1)), COL_ZBR_BROJ, broj, _
                      COL_ZBR_VOZAC, vozac, COL_ZBR_KUPAC, GLD_KUPAC

    ' Zaglavlje bez stavki obara strog citalac zbirne nad CELIM registrom
    ' (S4-1), a golden se vrti pre ostalih suite-a i ostavlja svoje redove u
    ' svesci. Stavka ide uz zaglavlje iz istog razloga iz kog je SeedZbirna
    ' dobija u storno suite-u.
    GldZbirnaStavka zbirnaID, KLASA_I, kolI, 50
    GldDodaj m_Zbr, zbirnaID
End Sub

' Stavka uz zaglavlje zbirne -- kroz ListRow, kao ostali seed-ovi ovog modula.
Private Sub GldZbirnaStavka(ByVal zbirnaID As String, ByVal klasa As String, _
                            ByVal kol As Double, ByVal amb As Long)
    If kol <= 0 Then Exit Sub

    Dim lo As ListObject
    Set lo = GetTable(TBL_ZBIRNA_STAVKE)
    If lo Is Nothing Then Exit Sub

    Dim nr As ListRow
    Set nr = lo.ListRows.Add

    GldCelija nr, TBL_ZBIRNA_STAVKE, COL_ZBS_ID, zbirnaID & "-S1"
    GldCelija nr, TBL_ZBIRNA_STAVKE, COL_ZBS_ZBIRNA_ID, zbirnaID
    GldCelija nr, TBL_ZBIRNA_STAVKE, COL_ZBS_RB, 1
    GldCelija nr, TBL_ZBIRNA_STAVKE, COL_ZBS_KLASA, klasa
    GldCelija nr, TBL_ZBIRNA_STAVKE, COL_ZBS_KOLICINA, kol
    GldCelija nr, TBL_ZBIRNA_STAVKE, COL_ZBS_KOL_AMB, amb
End Sub

Private Sub GldCelija(ByVal nr As ListRow, ByVal tbl As String, _
                      ByVal kolona As String, ByVal vrednost As Variant)
    Dim ci As Long
    ci = GetColumnIndex(tbl, kolona)
    If ci > 0 Then nr.Range.cells(1, ci).value = vrednost
End Sub

' Storno JEDNE od vise zbirnih pod istim brojem.
'
' Danas StornoZbirna_TX trazi broj + generaciju, jer broj sam nije identitet --
' adapter generaciju cita iz reda koji je writer vratio. Posle Zbirna
' cutover-a ide
' StornoZbirna_TX(ZbirnaID); scenario i golden ostaju isti.
Private Sub GldStornoZbirne(ByVal idx As Long)
    Dim data As Variant
    Dim cID As Long, cGen As Long, cBroj As Long
    Dim i As Long
    Dim ciljID As String
    Dim gen As String, broj As String

    ciljID = Trim$(CStr(m_Zbr(idx)))

    data = GetTableData(TBL_ZBIRNA)
    If Not IsArray(data) Then
        Err.Raise GLD_ERR, "GldStornoZbirne", "tblZbirna je prazna"
    End If

    cID = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_ID, "GldStornoZbirne")
    cBroj = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, "GldStornoZbirne")
    cGen = GetColumnIndex(TBL_ZBIRNA, COL_GENERACIJA_ID)

    For i = 1 To UBound(data, 1)
        If StrComp(Trim$(NzToText(data(i, cID))), ciljID, vbTextCompare) = 0 Then
            broj = Trim$(NzToText(data(i, cBroj)))
            If cGen > 0 Then gen = Trim$(NzToText(data(i, cGen)))
            Exit For
        End If
    Next i

    If Len(broj) = 0 Then
        Err.Raise GLD_ERR, "GldStornoZbirne", "zbirna " & ciljID & " nije nadjena"
    End If

    ' Storno ide po ZbirnaID-u (S4-2) -- a ciljID je bas on, pa se broj i
    ' generacija ovde vise ne razresavaju. Broj ostaje samo za poruku ispod.
    If Not StornoZbirna_TX(ciljID) Then
        Err.Raise GLD_ERR, "GldStornoZbirne", "storno zbirne nije uspeo"
    End If
End Sub

Private Sub GldStornoOtkupa(ByVal brDok As String)
    ' Storno ide po OtkupID-u (S1e). Broj se ovde razresava samo zato sto ga
    ' scenario zna; golden meri posledice storna, ne izbor dokumenta.
    Dim otkupID As String
    otkupID = LookupActiveID(TBL_OTKUP, COL_OTK_BR_DOK, brDok, COL_OTK_ID)
    If Len(otkupID) = 0 Then
        Err.Raise GLD_ERR, "GldStornoOtkupa", "otkup " & brDok & " nije nadjen"
    End If
    If Not StornoOtkup_TX(otkupID) Then
        Err.Raise GLD_ERR, "GldStornoOtkupa", "storno otkupa nije uspeo"
    End If
End Sub

Private Sub GldPocni(ByRef tx As clsTransaction, ByVal kljuc As String)
    Set tx = GldTx()
    GldReset
    m_Kljuc = kljuc
    GldSeed
    GldPreduslov kljuc
End Sub


'=====================================================================
' GRUPA A -- Fresh Fruit Flow
'=====================================================================

'=====================================================================
' GRUPA B -- Novac
'=====================================================================

'=====================================================================
' GRUPA D -- Storno
'=====================================================================

' D3: storno DVOKLASNOG otkupa -- jedan logicki dokument, obe klase.
'
' Ovo je scenario koji Otkup cutover najvise menja: danas se storno radi po
' poslovnom broju bas zato sto dokument nema jedan ID. Posle toga ide po
' DocumentID,
' a golden mora ostati isti.
Private Sub Gld_D3_StornoDvoklasnogOtkupa()
    Dim tx As clsTransaction
    Dim gldDesc As String
    Dim broj As String

    On Error GoTo EH
    broj = "GLD-D3"
    GldPocni tx, broj

    ' Novac ide AVANSOM: kes se ne vezuje za otkupni list (S8). Avans se pri
    ' upisu alocira na otkup, pa storno ima sta da odveze.
    GldAvans 56000#
    GldOtkup broj, broj & "-B", 1000#, 50#, 200#, 30#
    GldStornoOtkupa broj & "-B"

    AssertSnapshot GldSnapshot("D3 storno dvoklasnog otkupa", broj), GldIme(12)

    tx.RollbackTx
    Exit Sub
EH:
    gldDesc = Err.description
    If Not tx Is Nothing Then tx.RollbackTx
    Err.Raise GLD_ERR, "Gld_D3", gldDesc
End Sub


'=====================================================================
' GRUPA F -- Faktura
'=====================================================================

'=====================================================================
' GRUPA G -- Ivicni
'=====================================================================

' G2: DVA dokumenta sa istim poslovnim brojem.
'
' Kapija G2 ugovora: storno jednog ne sme da dirne drugi. Danas to drzi
' GeneracijaID; posle Otkup cutover-a drzi DocumentID, a golden ostaje isti.
Private Sub Gld_G2_IstiBrojDvaDokumenta()
    Dim tx As clsTransaction
    Dim gldDesc As String
    Dim broj As String

    On Error GoTo EH
    broj = "GLD-G2"
    GldPocni tx, broj

    ' DVE zbirne sa ISTIM BrojZbirne, razlicitih vlasnika (dva vozaca) -- stanje
    ' koje generator ne pravi, ali rucni unos i uvoz prave (ZBR_IDENTITET.md S3).
    GldZbirnaZaVozaca broj, GLD_VOZAC, 400#
    GldZbirnaZaVozaca broj, GLD_VOZAC2, 600#

    ' storno PRVE -- druga mora ostati netaknuta
    GldStornoZbirne 1

    ' Prazan kljuc -> preskace se broj-based ZBIRNA sekcija; identitet ide
    ' zasebno, po poziciji.
    AssertSnapshot GldSnapshot("G2 isti broj dva dokumenta", "") & _
                   GldIdentitetZbirni(), GldIme(15)

    tx.RollbackTx
    Exit Sub
EH:
    gldDesc = Err.description
    If Not tx Is Nothing Then tx.RollbackTx
    Err.Raise GLD_ERR, "Gld_G2", gldDesc
End Sub

