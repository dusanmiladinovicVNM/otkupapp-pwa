Attribute VB_Name = "modGoldenTests"

Option Explicit

'=====================================================================
' modGoldenTests -- GOLDEN POSLOVNI SCENARIJI
'
' Spec i recnik tvrdnji: docs/DOMEN/GOLDEN_SCENARIJI.md
'
' Kriterijum kvaliteta:
'   SCENARIO KOJI BI MORAO DA SE MENJA U PR3 JE NAPISAN POGRESNO.
'
' IZOLACIJA ULAZNOG STANJA -- ne samo izlaznog.
'   Rollback cisti ono sto scenario OSTAVI. Ne cisti ono sto je ZATEKAO.
'   Prva verzija je koristila KOOP-TEST-1 iz fixture-a i svih pet golden-a je
'   javljalo "placeno 1000.00" iako nijedan scenario ne placa nista -- to je bio
'   zatecen avans, koji SaveOtkupMulti_TX automatski primeni. Izmena tog avansa
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
Private Const GLD_KUPAC As String = "KUP-GLD-1"
Private Const GLD_VRSTA As String = "TESTVOCE"
Private Const GLD_SORTA As String = "TESTSORTA"
Private Const GLD_AMB As String = "12/1"

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

    GldOne 1
    GldOne 2
    GldOne 3
    GldOne 4
    GldOne 5
    ' 6 (B1) i 9 (B4) su UKLONJENI: kes se nikad ne vezuje za otkupni list, pa
    ' su merili putanju koja u domenu ne postoji. Ekran otkupnog lista nema polje
    ' za novac, a modOtkupUnos salje novac:=0 uvek -- v. GOLDEN_SCENARIJI.md S8.
    GldOne 7
    GldOne 8
    GldOne 10
    GldOne 11
    GldOne 12
    GldOne 13
    GldOne 14
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
        Case 1: GldIme = "A1_pun_lanac_do_fakture"
        Case 2: GldIme = "A2_dvoklasni_lanac"
        Case 3: GldIme = "A3_vise_blokova_jedna_otpremnica"
        Case 4: GldIme = "A4_vise_otpremnica_jedna_zbirna"
        Case 5: GldIme = "A5_kalo"
        Case 7: GldIme = "B2_avans_primenjen"
        Case 8: GldIme = "B3_delimican_avans"
        Case 10: GldIme = "D1_storno_otpremnice"
        Case 11: GldIme = "D2_storno_fakturisane_prijemnice"
        Case 12: GldIme = "D3_storno_dvoklasnog_otkupa"
        Case 13: GldIme = "F2_delimicno_fakturisanje"
        Case 14: GldIme = "G1_samo_klasa_dva"
        Case 15: GldIme = "G2_isti_broj_dva_dokumenta"
    End Select
End Function

Private Sub GldPozovi(ByVal idx As Long)
    Select Case idx
        Case 1: Gld_A1_PunLanacDoFakture
        Case 2: Gld_A2_DvoklasniLanac
        Case 3: Gld_A3_ViseBlokova
        Case 4: Gld_A4_ViseOtpremnica
        Case 5: Gld_A5_Kalo
        Case 7: Gld_B2_Avans
        Case 8: Gld_B3_DelimicanAvans
        Case 10: Gld_D1_StornoOtpremnice
        Case 11: Gld_D2_StornoFakturisanePrijemnice
        Case 12: Gld_D3_StornoDvoklasnogOtkupa
        Case 13: Gld_F2_DelimicnoFakturisanje
        Case 14: Gld_G1_SamoKlasaDva
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

' Maticni podaci scenarija. Idempotentno; sve se rollback-uje.
Private Sub GldSeed()
    GldSeedRed TBL_STANICE, "StanicaID", GLD_STANICA, "Naziv", "GOLDEN STANICA"
    GldSeedRed TBL_VOZACI, "VozacID", GLD_VOZAC, "Ime", "GOLDEN VOZAC"
    GldSeedRed TBL_KUPCI, "KupacID", GLD_KUPAC, "Naziv", "GOLDEN KUPAC"
    GldSeedKooperant
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
' Jedino mesto koje zna kako su podaci slozeni. Menja se u PR3+; scenariji i
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
    s = s & GldFakture()

    GldSnapshot = s
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

' Fakturisanost PO KLASI -- dokaz da je Fakturisano line-level.
Private Function GldFakturisanoPoKlasi() As String
    Dim data As Variant
    Dim cBr As Long, cKlasa As Long, cFak As Long
    Dim i As Long
    Dim fakI As String, fakII As String

    fakI = "-"
    fakII = "-"

    data = GetTableData(TBL_PRIJEMNICA)
    If IsArray(data) Then
        data = ExcludeStornirano(data, TBL_PRIJEMNICA)
        If IsArray(data) Then
            cBr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, "GldFakPoKlasi")
            cKlasa = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KLASA, "GldFakPoKlasi")
            cFak = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_FAKTURISANO, "GldFakPoKlasi")
            For i = 1 To UBound(data, 1)
                If StrComp(Trim$(NzToText(data(i, cBr))), m_Kljuc, vbTextCompare) = 0 Then
                    If UCase$(Trim$(NzToText(data(i, cKlasa)))) = "II" Then
                        fakII = IIf(UCase$(Trim$(NzToText(data(i, cFak)))) = "DA", "DA", "NE")
                    Else
                        fakI = IIf(UCase$(Trim$(NzToText(data(i, cFak)))) = "DA", "DA", "NE")
                    End If
                End If
            Next i
        End If
    End If

    GldFakturisanoPoKlasi = "PRIJEM" & vbLf & _
        "  fakturisano     I=" & fakI & "  II=" & fakII & vbLf
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
' zbirna i prijemnica ga imaju. Posle PR5 sve postaje COUNT(DISTINCT <Doc>ID) --
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

    cKlasa = RequireColumnIndex(TBL_OTKUP, COL_OTK_KLASA, "GldOtkupi")
    cKol = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA, "GldOtkupi")
    cCena = RequireColumnIndex(TBL_OTKUP, COL_OTK_CENA, "GldOtkupi")
    cIspl = RequireColumnIndex(TBL_OTKUP, COL_OTK_ISPLACENO, "GldOtkupi")
    cID = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, "GldOtkupi")

    For i = 1 To UBound(data, 1)
        id = Trim$(NzToText(data(i, cID)))
        If trazeni.Exists(id) Then
            If UCase$(Trim$(NzToText(data(i, cKlasa)))) = "II" Then
                kolII = kolII + SafeD(data(i, cKol))
            Else
                kolI = kolI + SafeD(data(i, cKol))
            End If
            vrednost = vrednost + SafeD(data(i, cKol)) * SafeD(data(i, cCena))
            placeno = placeno + GetIsplataForOtkup(id)
            imaAktivnih = True
            If UCase$(Trim$(NzToText(data(i, cIspl)))) <> "DA" Then svePlaceno = False
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

' Broj se koristi zato sto ga DANAS traze SumOtpremniceByKlasa i
' IsZbirnaConsistent -- zatecen API, ne izbor scenarija. Posle PR4 primaju
' ZbirnaID, pa se menja OVAJ adapter.
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
    s = s & "  invarijanta     " & IIf(IsZbirnaConsistent(brojZbirne), "OK", "PUKLA") & vbLf

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

    GldFakture = "FAKTURA" & vbLf & "  iznos           " & Fmt2(iznos) & vbLf
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

Private Function GldRedPostoji(ByVal tbl As String, ByVal kolona As String, _
                               ByVal vrednost As String) As Boolean
    GldRedPostoji = (GldBrojRedova(tbl, kolona, vrednost) > 0)
End Function

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

' Razlaze "OTK-1 + OTK-2". Taj format PR5 uklanja; tada se ovo svodi na
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

Private Sub GldOtkup(ByVal brojZbirne As String, ByVal brDok As String, _
                     ByVal kolI As Double, ByVal cenaI As Double, _
                     ByVal kolII As Double, ByVal cenaII As Double)
    Dim res As String

    res = SaveOtkupMulti_TX(GLD_DATUM, GLD_KOOP, GLD_STANICA, GLD_VRSTA, GLD_SORTA, _
                            kolI, cenaI, GLD_AMB, 50, GLD_VOZAC, brDok, 0#, "", "", _
                            brojZbirne, (kolII > 0), kolII, cenaII)
    If Len(res) = 0 Then
        Err.Raise GLD_ERR, "GldOtkup", "SaveOtkupMulti_TX nije vratio ID"
    End If

    GldDodaj m_Otk, res
End Sub

Private Sub GldOtpremnica(ByVal broj As String, ByVal brojOtp As String, _
        ByVal kolI As Double, ByVal cenaI As Double, _
        ByVal kolII As Double, ByVal cenaII As Double, ByVal ambII As Long)
    Dim res As String

    ' Ambalaza mora biti ista na otpremnici i na zbirnoj -- inace invarijanta
    ' puca na TEST PODACIMA, a golden bi zabelezio "PUKLA" kao da je sistem kriv.
    res = SaveOtpremnicaMulti_TX(GLD_DATUM, GLD_STANICA, GLD_VOZAC, brojOtp, _
            broj, GLD_VRSTA, GLD_SORTA, kolI, cenaI, GLD_AMB, 50, _
            (kolII > 0), kolII, cenaII, 0#, ambII, 0#)
    If Len(res) = 0 Then
        Err.Raise GLD_ERR, "GldOtpremnica", "otpremnica nije snimljena"
    End If

    GldDodaj m_Otp, res
End Sub

' ambI je UKUPNA ambalaza Klase I na zbirnoj -- mora biti ZBIR svih otpremnica
' tog broja. A4 salje dve otpremnice po 50 gajbi, pa zbirna dobija 100; sa 50
' bi invarijanta pukla na TEST PODACIMA i golden bi zabelezio "PUKLA" kao da je
' sistem kriv.
Private Sub GldZbirnaIPrijemnica(ByVal broj As String, _
        ByVal kolI As Double, ByVal kolII As Double, _
        ByVal ambI As Long, ByVal ambII As Long, _
        ByVal prijI As Double, ByVal prijII As Double)
    Dim res As String

    res = SaveZbirnaMulti_TX(GLD_DATUM, GLD_VOZAC, broj, GLD_KUPAC, "", "", _
            GLD_VRSTA, GLD_SORTA, kolI, GLD_AMB, ambI, (kolII > 0), kolII, ambII)
    If Len(res) = 0 Then
        Err.Raise GLD_ERR, "GldZbirna", "zbirna nije snimljena"
    End If
    GldDodaj m_Zbr, res

    res = SavePrijemnicaMulti_TX(GLD_DATUM, GLD_KUPAC, GLD_VOZAC, broj & "-P", _
            broj, GLD_VRSTA, GLD_SORTA, prijI, 55#, GLD_AMB, 50, 0, _
            (prijII > 0), prijII, 35#)
    If Len(res) = 0 Then
        Err.Raise GLD_ERR, "GldPrijemnica", "prijemnica nije snimljena"
    End If
    GldDodaj m_Prj, res
End Sub

' Faktura nad SVIM prijemnicama koje je scenario napravio.
Private Sub GldFaktura()
    Dim stavke As Collection
    Dim i As Long
    Dim res As String

    Set stavke = New Collection
    For i = 1 To m_Prj.count
        stavke.Add Array(Trim$(CStr(m_Prj(i))))
    Next i

    res = CreateFaktura_TX(GLD_KUPAC, stavke)
    If Len(res) = 0 Then
        Err.Raise GLD_ERR, "GldFaktura", "CreateFaktura_TX nije vratio ID"
    End If

    GldDodaj m_Fak, res
End Sub

' Avans kooperantu -- scenario ga pravi SAM, jer je zatecen avans bio prva
' velika rupa u izolaciji (GOLDEN_SCENARIJI.md S1b).
Private Sub GldAvans(ByVal iznos As Double)
    If Len(SaveNovac_TX("GLD-AV", GLD_DATUM, "GOLDEN KOOPERANT", GLD_KOOP, _
            "Kooperant", GLD_STANICA, GLD_KOOP, "", GLD_VRSTA, _
            NOV_VIRMAN_AVANS_KOOP, 0#, iznos)) = 0 Then
        Err.Raise GLD_ERR, "GldAvans", "avans nije snimljen"
    End If
End Sub

Private Sub GldStornoOtpremnice(ByVal idx As Long)
    If Not StornoOtpremnica_TX(Trim$(CStr(m_Otp(idx)))) Then
        Err.Raise GLD_ERR, "GldStornoOtpremnice", "storno otpremnice nije uspeo"
    End If
End Sub

Private Sub GldStornoPrijemnice(ByVal idx As Long)
    If Not StornoPrijemnica_TX(Trim$(CStr(m_Prj(idx)))) Then
        Err.Raise GLD_ERR, "GldStornoPrijemnice", "storno prijemnice nije uspeo"
    End If
End Sub

' Storno CELOG otkupnog bloka po poslovnom broju -- danas je to jedini ulaz koji
' zahvati obe klase. PR5 ga zamenjuje storno-om po DocumentID.
Private Sub GldStornoOtkupa(ByVal brDok As String)
    If Not StornoOtkupByBrDok_TX(brDok) Then
        Err.Raise GLD_ERR, "GldStornoOtkupa", "storno otkupa nije uspeo"
    End If
End Sub

Private Sub GldLanac(ByVal broj As String, _
        ByVal kolI As Double, ByVal cenaI As Double, _
        ByVal kolII As Double, ByVal cenaII As Double, _
        ByVal prijI As Double, ByVal prijII As Double)
    Dim ambII As Long

    If kolII > 0 Then ambII = 10

    GldOtkup broj, broj & "-B", kolI, cenaI, kolII, cenaII
    GldOtpremnica broj, broj & "-O", kolI, cenaI, kolII, cenaII, ambII
    GldZbirnaIPrijemnica broj, kolI, kolII, 50, ambII, prijI, prijII
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

' A1: baseline CELOG lanca, ukljucujuci Fakturu.
'
' Ide do fakture namerno: bas taj deo menja PR9 (FakturaStavka ->
' PrijemnicaStavkaID), pa baseline koji staje na prijemnici ne bi stitio nista.
Private Sub Gld_A1_PunLanacDoFakture()
    Dim tx As clsTransaction
    Dim gldDesc As String
    Dim broj As String

    On Error GoTo EH
    broj = "GLD-A1"
    GldPocni tx, broj
    GldLanac broj, 1000#, 50#, 0#, 0#, 1000#, 0#
    GldFaktura

    AssertSnapshot GldSnapshot("A1 pun lanac do fakture", broj), GldIme(1)

    tx.RollbackTx
    Exit Sub
EH:
    gldDesc = Err.description
    If Not tx Is Nothing Then tx.RollbackTx
    Err.Raise GLD_ERR, "Gld_A1", gldDesc
End Sub

' A2: dvoklasni lanac. Scenario koji PR5 najvise menja iznutra -- ishod isti.
Private Sub Gld_A2_DvoklasniLanac()
    Dim tx As clsTransaction
    Dim gldDesc As String
    Dim broj As String

    On Error GoTo EH
    broj = "GLD-A2"
    GldPocni tx, broj
    GldLanac broj, 1000#, 50#, 200#, 30#, 1000#, 200#
    GldFaktura

    AssertSnapshot GldSnapshot("A2 dvoklasni lanac", broj), GldIme(2)

    tx.RollbackTx
    Exit Sub
EH:
    gldDesc = Err.description
    If Not tx Is Nothing Then tx.RollbackTx
    Err.Raise GLD_ERR, "Gld_A2", gldDesc
End Sub

' A3: DVA otkupna bloka -> JEDNA otpremnica.
'
' Kardinalnost je u snapshotu (DOKUMENTI: otkupa 2, otpremnica 1). Bez toga bi
' bug koji napravi dve otpremnice po 500 kg ostavio agregat isti i test zelen.
Private Sub Gld_A3_ViseBlokova()
    Dim tx As clsTransaction
    Dim gldDesc As String
    Dim broj As String

    On Error GoTo EH
    broj = "GLD-A3"
    GldPocni tx, broj
    GldOtkup broj, broj & "-B1", 400#, 50#, 0#, 0#
    GldOtkup broj, broj & "-B2", 600#, 50#, 0#, 0#
    GldOtpremnica broj, broj & "-O", 1000#, 50#, 0#, 0#, 0
    GldZbirnaIPrijemnica broj, 1000#, 0#, 50, 0, 1000#, 0#

    AssertSnapshot GldSnapshot("A3 vise blokova jedna otpremnica", broj), GldIme(3)

    tx.RollbackTx
    Exit Sub
EH:
    gldDesc = Err.description
    If Not tx Is Nothing Then tx.RollbackTx
    Err.Raise GLD_ERR, "Gld_A3", gldDesc
End Sub

' A4: DVE otpremnice -> JEDNA zbirna. Kardinalnost opet u snapshotu.
Private Sub Gld_A4_ViseOtpremnica()
    Dim tx As clsTransaction
    Dim gldDesc As String
    Dim broj As String

    On Error GoTo EH
    broj = "GLD-A4"
    GldPocni tx, broj
    GldOtkup broj, broj & "-B", 1000#, 50#, 0#, 0#
    GldOtpremnica broj, broj & "-O1", 400#, 50#, 0#, 0#, 0
    GldOtpremnica broj, broj & "-O2", 600#, 50#, 0#, 0#, 0
    GldZbirnaIPrijemnica broj, 1000#, 0#, 100, 0, 1000#, 0#

    AssertSnapshot GldSnapshot("A4 vise otpremnica jedna zbirna", broj), GldIme(4)

    tx.RollbackTx
    Exit Sub
EH:
    gldDesc = Err.description
    If Not tx Is Nothing Then tx.RollbackTx
    Err.Raise GLD_ERR, "Gld_A4", gldDesc
End Sub

' A5: poslato 1000, primljeno 975. Razlika je POSLOVNA CINJENICA (kalo).
Private Sub Gld_A5_Kalo()
    Dim tx As clsTransaction
    Dim gldDesc As String
    Dim broj As String

    On Error GoTo EH
    broj = "GLD-A5"
    GldPocni tx, broj
    GldLanac broj, 1000#, 50#, 0#, 0#, 975#, 0#

    AssertSnapshot GldSnapshot("A5 kalo", broj), GldIme(5)

    tx.RollbackTx
    Exit Sub
EH:
    gldDesc = Err.description
    If Not tx Is Nothing Then tx.RollbackTx
    Err.Raise GLD_ERR, "Gld_A5", gldDesc
End Sub


'=====================================================================
' GRUPA B -- Novac
'=====================================================================

' B2: avans koji je scenario SAM napravio primenjuje se na otkup.
Private Sub Gld_B2_Avans()
    Dim tx As clsTransaction
    Dim gldDesc As String
    Dim broj As String

    On Error GoTo EH
    broj = "GLD-B2"
    GldPocni tx, broj

    ' avans pokriva PUNU vrednost (1000 x 50)
    GldAvans 50000#
    GldOtkup broj, broj & "-B", 1000#, 50#, 0#, 0#
    GldOtpremnica broj, broj & "-O", 1000#, 50#, 0#, 0#, 0
    GldZbirnaIPrijemnica broj, 1000#, 0#, 50, 0, 1000#, 0#

    AssertSnapshot GldSnapshot("B2 avans primenjen na otkup", broj), GldIme(7)

    tx.RollbackTx
    Exit Sub
EH:
    gldDesc = Err.description
    If Not tx Is Nothing Then tx.RollbackTx
    Err.Raise GLD_ERR, "Gld_B2", gldDesc
End Sub

' B3: avans pokriva SAMO deo vrednosti.
'
' Isplata ide preko avansa, ne kesa: kes se nikad ne vezuje za otkupni list
' (ekran nema to polje, modOtkupUnos salje novac:=0 uvek).
Private Sub Gld_B3_DelimicanAvans()
    Dim tx As clsTransaction
    Dim gldDesc As String
    Dim broj As String

    On Error GoTo EH
    broj = "GLD-B3"
    GldPocni tx, broj

    ' vrednost 50000, avansom pokriveno 20000
    GldAvans 20000#
    GldOtkup broj, broj & "-B", 1000#, 50#, 0#, 0#
    GldOtpremnica broj, broj & "-O", 1000#, 50#, 0#, 0#, 0
    GldZbirnaIPrijemnica broj, 1000#, 0#, 50, 0, 1000#, 0#

    AssertSnapshot GldSnapshot("B3 delimican avans", broj), GldIme(8)

    tx.RollbackTx
    Exit Sub
EH:
    gldDesc = Err.description
    If Not tx Is Nothing Then tx.RollbackTx
    Err.Raise GLD_ERR, "Gld_B3", gldDesc
End Sub

'=====================================================================
' GRUPA D -- Storno
'=====================================================================

' D1: storno otpremnice -- zbirna vise nema svoj izvor.
Private Sub Gld_D1_StornoOtpremnice()
    Dim tx As clsTransaction
    Dim gldDesc As String
    Dim broj As String

    On Error GoTo EH
    broj = "GLD-D1"
    GldPocni tx, broj

    GldOtkup broj, broj & "-B", 1000#, 50#, 0#, 0#
    GldOtpremnica broj, broj & "-O", 1000#, 50#, 0#, 0#, 0
    GldZbirnaIPrijemnica broj, 1000#, 0#, 50, 0, 1000#, 0#
    GldStornoOtpremnice 1

    AssertSnapshot GldSnapshot("D1 storno otpremnice", broj), GldIme(10)

    tx.RollbackTx
    Exit Sub
EH:
    gldDesc = Err.description
    If Not tx Is Nothing Then tx.RollbackTx
    Err.Raise GLD_ERR, "Gld_D1", gldDesc
End Sub

' D2: storno prijemnice koja je vec fakturisana -- kaskada.
Private Sub Gld_D2_StornoFakturisanePrijemnice()
    Dim tx As clsTransaction
    Dim gldDesc As String
    Dim broj As String

    On Error GoTo EH
    broj = "GLD-D2"
    GldPocni tx, broj

    GldLanac broj, 1000#, 50#, 0#, 0#, 1000#, 0#
    GldFaktura
    GldStornoPrijemnice 1

    AssertSnapshot GldSnapshot("D2 storno fakturisane prijemnice", broj), GldIme(11)

    tx.RollbackTx
    Exit Sub
EH:
    gldDesc = Err.description
    If Not tx Is Nothing Then tx.RollbackTx
    Err.Raise GLD_ERR, "Gld_D2", gldDesc
End Sub

' D3: storno DVOKLASNOG otkupa -- jedan logicki dokument, obe klase.
'
' Ovo je scenario koji PR5 najvise menja: danas se storno radi po poslovnom
' broju bas zato sto dokument nema jedan ID. Posle refaktora ide po DocumentID,
' a golden mora ostati isti.
Private Sub Gld_D3_StornoDvoklasnogOtkupa()
    Dim tx As clsTransaction
    Dim gldDesc As String
    Dim broj As String

    On Error GoTo EH
    broj = "GLD-D3"
    GldPocni tx, broj

    GldOtkupSaNovcem broj, broj & "-B", 1000#, 50#, 200#, 30#, 56000#
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

' F2: fakturise se SAMO Klasa I. Dokaz da je Fakturisano line-level --
' da je na headeru, ovo stanje ne bi moglo ni da postoji.
Private Sub Gld_F2_DelimicnoFakturisanje()
    Dim tx As clsTransaction
    Dim gldDesc As String
    Dim broj As String
    Dim stavke As Collection

    On Error GoTo EH
    broj = "GLD-F2"
    GldPocni tx, broj

    GldLanac broj, 1000#, 50#, 200#, 30#, 1000#, 200#

    ' samo prva prijemnicna stavka (Klasa I)
    Set stavke = New Collection
    stavke.Add Array(Trim$(CStr(m_Prj(1))))
    GldDodaj m_Fak, CreateFaktura_TX(GLD_KUPAC, stavke)

    AssertSnapshot GldSnapshot("F2 delimicno fakturisanje", broj) & _
                   GldFakturisanoPoKlasi(), GldIme(13)

    tx.RollbackTx
    Exit Sub
EH:
    gldDesc = Err.description
    If Not tx Is Nothing Then tx.RollbackTx
    Err.Raise GLD_ERR, "Gld_F2", gldDesc
End Sub


'=====================================================================
' GRUPA G -- Ivicni
'=====================================================================

' G1: samo Klasa II, bez Klase I. Grana koju hasKlasaI = False menja.
Private Sub Gld_G1_SamoKlasaDva()
    Dim tx As clsTransaction
    Dim gldDesc As String
    Dim broj As String

    On Error GoTo EH
    broj = "GLD-G1"
    GldPocni tx, broj

    GldOtkup broj, broj & "-B", 0#, 0#, 200#, 30#
    GldOtpremnica broj, broj & "-O", 0#, 0#, 200#, 30#, 10
    GldZbirnaIPrijemnica broj, 0#, 200#, 0, 10, 0#, 200#

    AssertSnapshot GldSnapshot("G1 samo klasa II", broj), GldIme(14)

    tx.RollbackTx
    Exit Sub
EH:
    gldDesc = Err.description
    If Not tx Is Nothing Then tx.RollbackTx
    Err.Raise GLD_ERR, "Gld_G1", gldDesc
End Sub

' G2: DVA dokumenta sa istim poslovnim brojem.
'
' Kapija G2 ugovora: storno jednog ne sme da dirne drugi. Danas to drzi
' GeneracijaID; posle PR5 drzi DocumentID, a golden ostaje isti.
Private Sub Gld_G2_IstiBrojDvaDokumenta()
    Dim tx As clsTransaction
    Dim gldDesc As String
    Dim broj As String

    On Error GoTo EH
    broj = "GLD-G2"
    GldPocni tx, broj

    GldOtkup broj, broj & "-B1", 400#, 50#, 0#, 0#
    GldOtkup broj, broj & "-B2", 600#, 50#, 0#, 0#
    GldStornoOtkupa broj & "-B1"

    AssertSnapshot GldSnapshot("G2 isti broj dva dokumenta", broj), GldIme(15)

    tx.RollbackTx
    Exit Sub
EH:
    gldDesc = Err.description
    If Not tx Is Nothing Then tx.RollbackTx
    Err.Raise GLD_ERR, "Gld_G2", gldDesc
End Sub

' Otkup sa novcem -- KORISTI SE SAMO u D3, da storno ima sta da ponisti.
'
' Redovna putanja NIKAD ne salje novac uz otkupni list: ekran nema to polje, a
' modOtkupUnos salje novac:=0. Parametar postoji jos samo na writer-u i ide u
' brisanje zajedno sa kolonama Novac/PrimalacNovca (v. DOCUMENT_HEADER_LINES S4.1).
Private Sub GldOtkupSaNovcem(ByVal brojZbirne As String, ByVal brDok As String, _
                             ByVal kolI As Double, ByVal cenaI As Double, _
                             ByVal kolII As Double, ByVal cenaII As Double, _
                             ByVal novac As Double)
    Dim res As String

    res = SaveOtkupMulti_TX(GLD_DATUM, GLD_KOOP, GLD_STANICA, GLD_VRSTA, GLD_SORTA, _
                            kolI, cenaI, GLD_AMB, 50, GLD_VOZAC, brDok, novac, _
                            "GOLDEN", "", brojZbirne, (kolII > 0), kolII, cenaII)
    If Len(res) = 0 Then
        Err.Raise GLD_ERR, "GldOtkupSaNovcem", "SaveOtkupMulti_TX nije vratio ID"
    End If

    GldDodaj m_Otk, res
End Sub
