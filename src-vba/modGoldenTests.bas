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

' Broj LOGICKIH dokumenata -- jedan poziv writer-a = jedan dokument.
'
' NE broji se po ID-u: dvoklasni dokument danas vraca "OTK-1 + OTK-2", pa bi
' brojanje ID-eva davalo 2 za JEDAN otkup -- broj fizickih redova, tacno ono
' sto recnik zabranjuje. Posle PR5 writer vraca jedan ID, a ovaj broj OSTAJE
' isti; golden se ne menja.
Private m_nOtk As Long
Private m_nOtp As Long
Private m_nZbr As Long
Private m_nPrj As Long
Private m_nFak As Long

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
    End Select
End Function

Private Sub GldPozovi(ByVal idx As Long)
    Select Case idx
        Case 1: Gld_A1_PunLanacDoFakture
        Case 2: Gld_A2_DvoklasniLanac
        Case 3: Gld_A3_ViseBlokova
        Case 4: Gld_A4_ViseOtpremnica
        Case 5: Gld_A5_Kalo
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

    m_nOtk = 0
    m_nOtp = 0
    m_nZbr = 0
    m_nPrj = 0
    m_nFak = 0
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

    If GldRedPostoji(tbl, kljucKol, kljuc) Then Exit Sub

    rowData = GldPrazanRed(tbl)
    GldPolje rowData, tbl, kljucKol, kljuc
    GldPolje rowData, tbl, nazivKol, naziv
    GldPolje rowData, tbl, "Aktivan", STATUS_AKTIVAN

    If AppendRow(tbl, rowData) <= 0 Then
        Err.Raise GLD_ERR, "GldSeedRed", "AppendRow nije uspeo za " & tbl
    End If
End Sub

Private Sub GldSeedKooperant()
    Dim rowData As Variant

    If GldRedPostoji(TBL_KOOPERANTI, "KooperantID", GLD_KOOP) Then Exit Sub

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
Private Sub GldPreduslov()
    Dim n As Long

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
    s = s & GldFakture()

    GldSnapshot = s
End Function

' Broj LOGICKIH dokumenata koje je scenario napravio.
'
' Ovo je poslovna cinjenica ("dva bloka su otisla na jednu otpremnicu"), ne
' broj fizickih redova -- pa preziva header+stavke nepromenjeno. Bez nje su
' A3 i A4 zeleni i kad bi bug napravio dve otpremnice po 500 kg: agregat je
' isti, a kardinalnost nije.
Private Function GldDokumenti() As String
    GldDokumenti = "DOKUMENTI" & vbLf & _
        "  otkupa          " & CStr(m_nOtk) & vbLf & _
        "  otpremnica      " & CStr(m_nOtp) & vbLf & _
        "  zbirnih         " & CStr(m_nZbr) & vbLf & _
        "  prijemnica      " & CStr(m_nPrj) & vbLf & _
        "  faktura         " & CStr(m_nFak) & vbLf
End Function

Private Function GldOtkupi() As String
    Dim data As Variant
    Dim cKlasa As Long, cKol As Long, cCena As Long, cIspl As Long, cID As Long
    Dim i As Long, k As Long
    Dim kolI As Double, kolII As Double, vrednost As Double, placeno As Double
    Dim svePlaceno As Boolean
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
            If UCase$(Trim$(NzToText(data(i, cIspl)))) <> "DA" Then svePlaceno = False
        End If
    Next i

    s = s & "  predao          I=" & Fmt2(kolI) & "  II=" & Fmt2(kolII) & vbLf
    s = s & "  vrednost        " & Fmt2(vrednost) & vbLf
    s = s & "  placeno         " & Fmt2(placeno) & vbLf
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
    m_nOtk = m_nOtk + 1
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
    m_nOtp = m_nOtp + 1
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
    m_nZbr = m_nZbr + 1

    res = SavePrijemnicaMulti_TX(GLD_DATUM, GLD_KUPAC, GLD_VOZAC, broj & "-P", _
            broj, GLD_VRSTA, GLD_SORTA, prijI, 55#, GLD_AMB, 50, 0, _
            (prijII > 0), prijII, 35#)
    If Len(res) = 0 Then
        Err.Raise GLD_ERR, "GldPrijemnica", "prijemnica nije snimljena"
    End If
    GldDodaj m_Prj, res
    m_nPrj = m_nPrj + 1
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
    m_nFak = m_nFak + 1
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

Private Sub GldPocni(ByRef tx As clsTransaction)
    Set tx = GldTx()
    GldReset
    GldSeed
    GldPreduslov
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
    GldPocni tx

    broj = "GLD-A1"
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
    GldPocni tx

    broj = "GLD-A2"
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
    GldPocni tx

    broj = "GLD-A3"
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
    GldPocni tx

    broj = "GLD-A4"
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
    GldPocni tx

    broj = "GLD-A5"
    GldLanac broj, 1000#, 50#, 0#, 0#, 975#, 0#

    AssertSnapshot GldSnapshot("A5 kalo", broj), GldIme(5)

    tx.RollbackTx
    Exit Sub
EH:
    gldDesc = Err.description
    If Not tx Is Nothing Then tx.RollbackTx
    Err.Raise GLD_ERR, "Gld_A5", gldDesc
End Sub
