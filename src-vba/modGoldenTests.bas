Attribute VB_Name = "modGoldenTests"

Option Explicit

'=====================================================================
' modGoldenTests -- GOLDEN POSLOVNI SCENARIJI
'
' Sigurnosna mreza za PR3-PR12. Spec i recnik tvrdnji:
' docs/DOMEN/GOLDEN_SCENARIJI.md
'
' Kriterijum kvaliteta je jedna recenica:
'   SCENARIO KOJI BI MORAO DA SE MENJA U PR3 JE NAPISAN POGRESNO.
'
' Zato scenario tvrdi POSLOVNE CINJENICE (kolicine po klasi, saldo, kalo,
' iznos fakture), nikad oblik implementacije (broj redova, GeneracijaID,
' format ID-a, "primarni red"). Kad Zbirna postane header+stavke, menja se
' UNUTRASNJOST GldSnapshot-a -- ne scenario i ne golden fajl.
'
' DETERMINIZAM: svaki scenario radi u transakciji i na kraju RADI ROLLBACK.
' Snapshot se uzima PRE rollback-a. Zato:
'   - fixture ostaje netaknut, pa je drugi prolaz identican prvom
'   - GetNextID daje iste ID-eve, jer je stanje tabela vraceno
'   - scenariji ne zavise od redosleda izvrsavanja
' Datumi su FIKSNI (GLD_DATUM), nikad Date/Now -- inace golden pada sutra.
'
' Maticni podaci se NE seju: koriste se oni koje make_fixture.py vec pravi
' (KOOP-TEST-1, STA-TEST-1, VOZ-TEST-1, KUP-TEST-1, TESTVOCE/TESTSORTA).
'=====================================================================

Private Const GLD_LOG As String = "GOLDEN_TEST_LOG"

' Svoj kod greske: modTest-ov je Private i iz ovog modula se ne vidi
' ("Variable not defined"). Compile tada pada, a run_vba to ne vidi kao
' pao test nego kao VISENJE -- 449 s do timeout-a i Excel u [break].
Private Const GLD_ERR As Long = vbObjectError + 9501

' Fiksan datum -- golden vezan za Date() bi pao sutra.
Private Const GLD_DATUM As Date = #3/15/2026#

' Maticni podaci iz fixture-a (tools/make_fixture.py)
Private Const GLD_KOOP As String = "KOOP-TEST-1"
Private Const GLD_STANICA As String = "STA-TEST-1"
Private Const GLD_VOZAC As String = "VOZ-TEST-1"
Private Const GLD_KUPAC As String = "KUP-TEST-1"
Private Const GLD_VRSTA As String = "TESTVOCE"
Private Const GLD_SORTA As String = "TESTSORTA"
Private Const GLD_AMB As String = "12/1"

Private m_Total As Long
Private m_Failed As Long
Private m_Report As String


'=====================================================================
' RUNNER
'=====================================================================

' gate: True u tools/run_vba.py -- pad mora da stigne do runnera kao greska,
' ne samo u Immediate.
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
        Err.Raise vbObjectError + 9500, "RunGoldenSuite", _
                  CStr(m_Failed) & " od " & CStr(m_Total) & " golden scenarija palo:" & _
                  vbCrLf & m_Report
    End If
End Sub

Private Function GldIme(ByVal idx As Long) As String
    Select Case idx
        Case 1: GldIme = "A1_jednoklasni_lanac"
        Case 2: GldIme = "A2_dvoklasni_lanac"
        Case 3: GldIme = "A3_vise_blokova_jedna_otpremnica"
        Case 4: GldIme = "A4_vise_otpremnica_jedna_zbirna"
        Case 5: GldIme = "A5_kalo"
    End Select
End Function

Private Sub GldPozovi(ByVal idx As Long)
    Select Case idx
        Case 1: Gld_A1_JednoklasniLanac
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
' OKVIR -- transakcija oko scenarija
'=====================================================================

' Snapshotuje SVE tabele koje lanac dokumenata moze da dodirne. Namerno sire
' nego sto pojedini scenario treba: rollback koji propusti jednu tabelu ostavlja
' fixture prljav, a sledeci prolaz onda vise nije identican.
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

    Set GldTx = tx
End Function


'=====================================================================
' SNAPSHOT -- iskljucivo poslovne cinjenice
'
' Ovo je JEDINO mesto koje sme da zna kako su podaci slozeni. Kad Zbirna
' postane header+stavke, menja se OVDE -- scenariji i golden fajlovi ostaju.
'=====================================================================

' Snapshot izvestava ISKLJUCIVO o dokumentima koje je scenario stvorio.
'
' Prva verzija je citala celu svesku po kooperantu i javljala "predao 1720" iako
' je scenario snimio 1000 -- ostatak su bili zatecen fixture. Golden je time
' zavisio od svake izmene fixture-a, a scenario nije merio ono sto tvrdi.
'
' otkupIDs je Collection ID-eva koje su vratili writer-i: scenario ZNA sta je
' napravio, pa snapshot ne mora da pretrazuje po poslovnom broju.
' Redovi se spajaju sa vbLf, NE vbCrLf: modTest.ReadTextFile izbacuje CR jer
' .gitattributes drzi tests/golden/*.txt na LF. Sa CRLF bi svaki snapshot bio
' razlicit od procitanog golden-a, a poruka bi pokazivala identicne linije.
Private Function GldSnapshot(ByVal naslov As String, _
                             ByVal otkupIDs As Collection, _
                             ByVal brojZbirne As String) As String
    Dim s As String

    s = "== " & naslov & " ==" & vbLf
    s = s & GldOtkupi(otkupIDs)
    If Len(brojZbirne) > 0 Then s = s & GldZbirna(brojZbirne)

    GldSnapshot = s
End Function

' Sta je kooperant predao kroz OVE otkupe, i da li je placeno.
Private Function GldOtkupi(ByVal otkupIDs As Collection) As String
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
    For k = 1 To otkupIDs.count
        id = Trim$(CStr(otkupIDs(k)))
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

' Zbirna: sta je poslato, sta primljeno, kolika je razlika (kalo).
'
' Broj se ovde koristi zato sto ga DANAS traze i SumOtpremniceByKlasa i
' IsZbirnaConsistent -- to je zatecen API, ne izbor scenarija. Posle PR4 oba
' primaju ZbirnaID, pa se menja OVAJ helper; scenario i golden ostaju isti.
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

Private Function GldDict(ByVal d As Object, ByVal k As String) As Double
    If d Is Nothing Then Exit Function
    If d.Exists(k) Then GldDict = SafeD(d(k))
End Function

Private Function SafeD(ByVal v As Variant) As Double
    If IsNumeric(v) Then SafeD = CDbl(v)
End Function

' Decimalna TACKA nezavisno od lokala. Format$ postuje regionalna podesavanja,
' pa bi golden snimljen na masini sa zarezom pao na masini sa tackom -- test bi
' merio Control Panel, ne poslovanje.
Private Function Fmt2(ByVal v As Double) As String
    Dim t As String
    t = Format$(v, "0.00")
    Fmt2 = Replace(t, ",", ".")
End Function


'=====================================================================
' GRUPA A -- Fresh Fruit Flow
'=====================================================================

' A1: baseline. Sve ostalo je odstupanje od ovoga.
Private Sub Gld_A1_JednoklasniLanac()
    Dim tx As clsTransaction
    Dim gldDesc As String
    Dim broj As String
    Dim ids As Collection

    On Error GoTo EH
    Set tx = GldTx()
    Set ids = New Collection

    broj = "GLD-A1"
    GldLanac ids, broj, 1000#, 50#, 0#, 0#, 1000#, 0#

    AssertSnapshot GldSnapshot("A1 jednoklasni lanac", ids, broj), _
                   GldIme(1)

    tx.RollbackTx
    Exit Sub
EH:
    ' Opis se cita PRE rollback-a: RollbackTx ima svoj On Error i obrise
    ' Err, pa bi poruka stigla PRAZNA (isti obrazac kao MRTAV_LOG).
    gldDesc = Err.description
    If Not tx Is Nothing Then tx.RollbackTx
    Err.Raise GLD_ERR, "Gld_A1", gldDesc
End Sub

' A2: dvoklasni lanac. Scenario koji PR5 najvise menja iznutra -- ishod isti.
Private Sub Gld_A2_DvoklasniLanac()
    Dim tx As clsTransaction
    Dim gldDesc As String
    Dim broj As String
    Dim ids As Collection

    On Error GoTo EH
    Set tx = GldTx()
    Set ids = New Collection

    broj = "GLD-A2"
    GldLanac ids, broj, 1000#, 50#, 200#, 30#, 1000#, 200#

    AssertSnapshot GldSnapshot("A2 dvoklasni lanac", ids, broj), _
                   GldIme(2)

    tx.RollbackTx
    Exit Sub
EH:
    ' Opis se cita PRE rollback-a: RollbackTx ima svoj On Error i obrise
    ' Err, pa bi poruka stigla PRAZNA (isti obrazac kao MRTAV_LOG).
    gldDesc = Err.description
    If Not tx Is Nothing Then tx.RollbackTx
    Err.Raise GLD_ERR, "Gld_A2", gldDesc
End Sub

' A3: dva otkupna bloka -> jedna otpremnica (N:1).
Private Sub Gld_A3_ViseBlokova()
    Dim tx As clsTransaction
    Dim gldDesc As String
    Dim broj As String
    Dim ids As Collection
    Dim r1 As String, r2 As String

    On Error GoTo EH
    Set tx = GldTx()
    Set ids = New Collection

    broj = "GLD-A3"
    r1 = SaveOtkupMulti_TX(GLD_DATUM, GLD_KOOP, GLD_STANICA, GLD_VRSTA, GLD_SORTA, _
                           400#, 50#, GLD_AMB, 20, GLD_VOZAC, "GLD-A3-B1", 0#, "", "", broj)
    r2 = SaveOtkupMulti_TX(GLD_DATUM, GLD_KOOP, GLD_STANICA, GLD_VRSTA, GLD_SORTA, _
                           600#, 50#, GLD_AMB, 30, GLD_VOZAC, "GLD-A3-B2", 0#, "", "", broj)
    If Len(r1) = 0 Or Len(r2) = 0 Then
        Err.Raise GLD_ERR, "Gld_A3", "otkup nije snimljen"
    End If
    GldDodajIDs ids, r1
    GldDodajIDs ids, r2

    GldOtpremnicaZbirnaPrijemnica broj, 1000#, 50#, 0#, 0#, 1000#, 0#

    AssertSnapshot GldSnapshot("A3 vise blokova jedna otpremnica", ids, broj), _
                   GldIme(3)

    tx.RollbackTx
    Exit Sub
EH:
    ' Opis se cita PRE rollback-a: RollbackTx ima svoj On Error i obrise
    ' Err, pa bi poruka stigla PRAZNA (isti obrazac kao MRTAV_LOG).
    gldDesc = Err.description
    If Not tx Is Nothing Then tx.RollbackTx
    Err.Raise GLD_ERR, "Gld_A3", gldDesc
End Sub

' A4: dve otpremnice -> jedna zbirna. Invarijanta S6.2 nad agregatom.
Private Sub Gld_A4_ViseOtpremnica()
    Dim tx As clsTransaction
    Dim gldDesc As String
    Dim broj As String
    Dim ids As Collection

    On Error GoTo EH
    Set tx = GldTx()
    Set ids = New Collection

    broj = "GLD-A4"
    GldOtkup ids, broj, "GLD-A4-B1", 1000#, 50#, 0#, 0#

    If Len(SaveOtpremnicaMulti_TX(GLD_DATUM, GLD_STANICA, GLD_VOZAC, "GLD-A4-O1", _
            broj, GLD_VRSTA, GLD_SORTA, 400#, 50#, GLD_AMB, 20)) = 0 Then
        Err.Raise GLD_ERR, "Gld_A4", "otpremnica 1 nije snimljena"
    End If
    If Len(SaveOtpremnicaMulti_TX(GLD_DATUM, GLD_STANICA, GLD_VOZAC, "GLD-A4-O2", _
            broj, GLD_VRSTA, GLD_SORTA, 600#, 50#, GLD_AMB, 30)) = 0 Then
        Err.Raise GLD_ERR, "Gld_A4", "otpremnica 2 nije snimljena"
    End If

    GldZbirnaIPrijemnica broj, 1000#, 0#, 1000#, 0#

    AssertSnapshot GldSnapshot("A4 vise otpremnica jedna zbirna", ids, broj), _
                   GldIme(4)

    tx.RollbackTx
    Exit Sub
EH:
    ' Opis se cita PRE rollback-a: RollbackTx ima svoj On Error i obrise
    ' Err, pa bi poruka stigla PRAZNA (isti obrazac kao MRTAV_LOG).
    gldDesc = Err.description
    If Not tx Is Nothing Then tx.RollbackTx
    Err.Raise GLD_ERR, "Gld_A4", gldDesc
End Sub

' A5: poslato 1000, primljeno 975. Razlika je POSLOVNA CINJENICA (kalo).
Private Sub Gld_A5_Kalo()
    Dim tx As clsTransaction
    Dim gldDesc As String
    Dim broj As String
    Dim ids As Collection

    On Error GoTo EH
    Set tx = GldTx()
    Set ids = New Collection

    broj = "GLD-A5"
    GldLanac ids, broj, 1000#, 50#, 0#, 0#, 975#, 0#

    AssertSnapshot GldSnapshot("A5 kalo", ids, broj), GldIme(5)

    tx.RollbackTx
    Exit Sub
EH:
    ' Opis se cita PRE rollback-a: RollbackTx ima svoj On Error i obrise
    ' Err, pa bi poruka stigla PRAZNA (isti obrazac kao MRTAV_LOG).
    gldDesc = Err.description
    If Not tx Is Nothing Then tx.RollbackTx
    Err.Raise GLD_ERR, "Gld_A5", gldDesc
End Sub


'=====================================================================
' POSLOVNI POTEZI -- tanki omotaci nad PRODUKCIONIM writer-ima
'
' Namerno se zovu isti ulazi koje zove UI. Scenario koji bi pisao u tabele
' direktno ne bi merio poslovanje nego sopstveni seed.
'
' Svaki vraca ID-eve koje je writer napravio -- snapshot izvestava samo o
' njima, pa zatecen fixture ne moze da zaprlja golden.
'=====================================================================

' Razlaze "OTK-1 + OTK-2" u pojedinacne ID-eve.
'
' Taj format je tacno ono sto PR5 uklanja; dok postoji, scenario mora da ga
' razume da bi znao STA je napravio. Kad Create*_TX pocne da vraca jedan ID,
' ova funkcija se svodi na jedan Add -- golden se NE menja.
Private Sub GldDodajIDs(ByVal cilj As Collection, ByVal rez As String)
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

Private Sub GldOtkup(ByVal cilj As Collection, _
                     ByVal brojZbirne As String, ByVal brDok As String, _
                     ByVal kolI As Double, ByVal cenaI As Double, _
                     ByVal kolII As Double, ByVal cenaII As Double)
    Dim res As String

    res = SaveOtkupMulti_TX(GLD_DATUM, GLD_KOOP, GLD_STANICA, GLD_VRSTA, GLD_SORTA, _
                            kolI, cenaI, GLD_AMB, 50, GLD_VOZAC, brDok, 0#, "", "", _
                            brojZbirne, (kolII > 0), kolII, cenaII)
    If Len(res) = 0 Then
        Err.Raise GLD_ERR, "GldOtkup", "SaveOtkupMulti_TX nije vratio ID"
    End If

    GldDodajIDs cilj, res
End Sub

Private Sub GldOtpremnicaZbirnaPrijemnica(ByVal broj As String, _
        ByVal kolI As Double, ByVal cenaI As Double, _
        ByVal kolII As Double, ByVal cenaII As Double, _
        ByVal prijI As Double, ByVal prijII As Double)

    ' Ambalaza mora biti ista na otpremnici i na zbirnoj -- inace invarijanta
    ' puca na TEST PODACIMA, a golden bi zabelezio "PUKLA" kao da je sistem kriv.
    ' Prva verzija je otpremnici davala amb samo za Klasu I (kolAmbII je ostajao
    ' 0), a zbirnoj za obe -- pa je A2 snimio lazan pad.
    If Len(SaveOtpremnicaMulti_TX(GLD_DATUM, GLD_STANICA, GLD_VOZAC, broj & "-O", _
            broj, GLD_VRSTA, GLD_SORTA, kolI, cenaI, GLD_AMB, 50, _
            (kolII > 0), kolII, cenaII, 0#, IIf(kolII > 0, 10, 0), 0#)) = 0 Then
        Err.Raise GLD_ERR, "GldOtpremnica", "otpremnica nije snimljena"
    End If

    GldZbirnaIPrijemnica broj, kolI, kolII, prijI, prijII
End Sub

Private Sub GldZbirnaIPrijemnica(ByVal broj As String, _
        ByVal kolI As Double, ByVal kolII As Double, _
        ByVal prijI As Double, ByVal prijII As Double)

    If Len(SaveZbirnaMulti_TX(GLD_DATUM, GLD_VOZAC, broj, GLD_KUPAC, "", "", _
            GLD_VRSTA, GLD_SORTA, kolI, GLD_AMB, 50, _
            (kolII > 0), kolII, 10)) = 0 Then
        Err.Raise GLD_ERR, "GldZbirna", "zbirna nije snimljena"
    End If

    If Len(SavePrijemnicaMulti_TX(GLD_DATUM, GLD_KUPAC, GLD_VOZAC, broj & "-P", _
            broj, GLD_VRSTA, GLD_SORTA, prijI, 55#, GLD_AMB, 50, 0, _
            (prijII > 0), prijII, 35#)) = 0 Then
        Err.Raise GLD_ERR, "GldPrijemnica", "prijemnica nije snimljena"
    End If
End Sub

Private Sub GldLanac(ByVal cilj As Collection, ByVal broj As String, _
        ByVal kolI As Double, ByVal cenaI As Double, _
        ByVal kolII As Double, ByVal cenaII As Double, _
        ByVal prijI As Double, ByVal prijII As Double)

    GldOtkup cilj, broj, broj & "-B", kolI, cenaI, kolII, cenaII
    GldOtpremnicaZbirnaPrijemnica broj, kolI, cenaI, kolII, cenaII, prijI, prijII
End Sub
