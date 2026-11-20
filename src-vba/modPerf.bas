Attribute VB_Name = "modPerf"
Option Explicit

' ============================================================
' modPerf - merni tajmer (profiler) za merenje kod operatera.
'
' POSMATRA, NE MENJA: ne pise ni u jednu tabelu sveske (A11), ne otvara
' dijalog u putu upisa i ne dize gresku. Izlaz:
'   <folder sveske>\Log\Perf_YYYY-MM-DD.csv      red po sekciji, po akciji
'   <folder sveske>\Log\OtkupApp_YYYY-MM-DD.log  jedan red sazetka (LogInfo)
'   <folder sveske>\Log\Perf_sazetak_*.txt       Alt+F8 -> Perf_Sazetak
'
' PREKIDAC: gPerfOn / gPerfDetalj. Pozivno mesto ih cita SAMO:
'     If gPerfOn Then PerfPocni "OZNAKA"
' pa iskljucen merac kosta jedno citanje Boolean-a: nema poziva i nema
' spajanja stringa oznake. Ukljucuje se jednom na startu iz tblLocalConfig
' (PERF_MERENJE = DA | DETALJ) ili rucno: Alt+F8 -> Perf_Ukljuci.
'   nivo 1 (gPerfOn):     akcije, sekcije domena, TX, BROJACI primitiva
'   nivo 2 (gPerfDetalj): i VREMENA primitiva modDataAccess (skuplje)
'
' SAT: QueryPerformanceCounter preko Currency. Currency je 64-bitni ceo broj
' skaliran sa 10^4; QPC upise svoja 64 bita u njega, a odnos razlike brojaca i
' frekvencije skalu ponistava. Isto radi u 32- i 64-bitnom Office-u (LongLong
' postoji samo u 64-bitnom). Kad QPC ne radi, rezerva je Timer (oko 1/64 s,
' v. modStornoImpact) uz prelaz ponoci (v. modBusinessFlowProTests.BfpKorak*).
'
' SEKCIJE: stek. "ukupno" po oznaci se dodaje samo kad se zatvori SPOLJNI okvir
' te oznake (rekurzija se ne broji dvaput); "self" = trajanje - deca. Sekcija
' ostavljena otvorenom (izlaz kroz Err.Raise) zatvara se sa prvom spoljnom i
' broji se u META Nezatvorenih -- merena procedura ne treba sopstveni EH.
'
' ERR-NEUTRALNO: javna procedura sa On Error sacuva Err na ulazu i vrati ga na
' izlazu (Number, Source, Description). I pored toga hook NE ide izmedju
' Application.Run i ScrPadPreuzmi/ScrRunIshod, ni izmedju operacije i provere
' Err.Number koja je cita, ni pre LogErr u EH (LogErr cita Err).
'
' POZIVNA MESTA (sva su na linijama koje optimizaciona grana analize brzine
' 2026-11 ne dira, pa isti merni kod radi i pre i posle optimizacija):
'   modMain.StartApp             PerfInitIzKonfiga (posle LogAppStart)
'   modOtkupUI.UiEvent           AKCIJA "UI.<dogadjaj>.<tag>" -- svaka radnja
'                                operatera; programski Change unutar radnje
'                                postaje sekcija AKCIJA.UI.Change.*
'   modOtkupUI.CommitDokument    UI.Commit.ScrSave / .ClearForm / .Refresh
'   modOtkupUI.RefreshFromData   AKCIJA "REFRESH" (ugnjezdena ili samostalna)
'                                delovi REF.* (ResetCache/Kpi/NavBrojaci/ZbirneCombo/
'                                MaticniCombos/OtvoreniBlokovi/Prikaz/OtpTraka), KPI.* i
'                                GRID.Ucitaj/
'                                GRID.Crtaj (ReloadGrid) -- rasclanjavanje 20.11.
'   modOtkupUI.OsveziNavBrojace  NAV|<ekran> po znacki menija; Oporavak
'                                (modStornoRecovery.GetNedovrseno) OPO.* po delu
'   modOtkupUI.Dijalog           UI.CekaOperatera (oduzima se od aktivnog)
'   modOtkupUnos.PotvrdiManjak.. UI.CekaOperatera (pitanje o manjku gajbi)
'   modDataAccess.GetTableData   BROJ "DA.GTD|<tabela>" (celije) i
'                                "DA.GTD.KES|<tabela>"; nivo 2 i vreme
'   clsTransaction.AddTableSnap. TX.Snapshot|<tabela>
'   modPrint.OutputOtkupniList   PRINT.OtkupniList
'   modScrDokumenti.Scr_Save     OTK.AutoLanacHladnjaca
'   modAutoHladnjaca.AutoLanac.. LNC.PostojeciLanac/.Brojevi/.Otpremnica/.Zbirna/
'                                .Prijemnica (rasclanjavanje 20.11.)
'   modMasterSync.AutoZbirnaZa.. DOK.NevezaneOtpremnice
'   modScrIzvestaji.PuniSnimak   IZV.PuniSnimak|<lista>|<P/Z>
'   modJournaling.AutoSaveTick   AKCIJA "ONTIME.AUTOSAVE"
'   modStornoWarm.BuildWarm      AKCIJA "ONTIME.WARM"
'   modIzvestaj.ReportSaldoOM..  IZV.SaldoOM i delovi .Otkup/.ZbirStavki/
'                                .Sifarnik/.Novac/.Doplata/.Agro/.Amb;
'                                zbirni: IZV.Zbirni.Stanice (rascl. 20.11.)
'   modAmbalaza (knjiga)         AMB.SemaKapija, AMB.KnjigaIntegritet,
'                                AMB.DokVlasnikMapa (odluka D1, analiza)
'                                BROJ AMB.KnjigaDokaz.KES: dokaz iz kesa (D1-IZV)
'                                AMB.KnjigaDokaz.Nastavak + BROJ .NASTAVAK (D1-UPIS)
'   modTest.ResetSeamova         PerfIskljuciTiho
' ============================================================

' Declare u JEDNOM redu. Prelom " _" tik pre "(" liste parametara tera VBE
' AddFromString da na kraj modula doda red "()": modul se ne kompajlira, a
' ImportAllVBA javi drift (izmereno 20.11.2026 u cistoj instanci Excela;
' #If nije bitan). Posle prepisa u jedan red ImportAllVBA prolazi (20.11.2026).
' Prelom ne vracati -- cuva ga vba_check PRELOM_POTPISA.
#If VBA7 Then
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
#Else
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
#End If

Public gPerfOn As Boolean           ' nivo 1
Public gPerfDetalj As Boolean       ' nivo 2 (ukljucuje i nivo 1)

Private Const PERF_CFG_KLJUC As String = "PERF_MERENJE"    ' DA | DETALJ | NE
Private Const PERF_CFG_PRAG As String = "PERF_PRAG_MS"     ' kraca akcija ne ide u fajl (podrazumevano 10)
Private Const PERF_FOLDER As String = "Log"
Private Const PERF_PREFIKS As String = "Perf_"
Private Const PERF_SEP As String = ";"
Private Const PERF_CEKANJE As String = "UI.CekaOperatera"
Private Const PERF_UGNJ As String = "AKCIJA."
Private Const PERF_STEK_MAX As Long = 512
Private Const PERF_POC_KAP As Long = 256
Private Const PERF_VRSTA_SEK As Long = 1
Private Const PERF_VRSTA_BROJ As Long = 2
Private Const PERF_KALIB_PAROVA As Long = 2000
Private Const PERF_TOP_N As Long = 25
Private Const PERF_ZASTARELA_MS As Double = 600000#        ' 10 min otvorena = procurela

' --- sat ---
Private mSatSpreman As Boolean
Private mImaQpc As Boolean
Private mQpcBaza As Currency
Private mQpcHz As Currency
Private mMsPoJedinici As Double
Private mTimerBaza As Double
Private mTimerPoslednji As Double
Private mTimerDana As Long

' --- tekuca akcija ---
Private mAkcija As String
Private mAkcijaDub As Long
Private mIstoIme As Long
Private mAkcijaT0 As Double
Private mAkcijaRb As Long
Private mAkcijaRbIstog As Long
Private mAkcijaCalc As Long
Private mNapomena As String
Private mPrag As Double
Private mOverheadSekUs As Double
Private mOverheadBrojUs As Double
Private mTopSuma As Double
Private mNeuparenih As Long
Private mNezatvorenih As Long
Private mPreliv As Long
Private mVanAkcije As Long

' --- statistika po oznaci (tekuca akcija), paralelni nizovi ---
Private mIdx As Object
Private mN As Long
Private mLbl() As String
Private mVrsta() As Long
Private mPoz() As Double
Private mKol() As Double
Private mUk() As Double
Private mSelf() As Double
Private mMax() As Double
Private mAkt() As Long

' --- stek otvorenih sekcija ---
Private mDub As Long
Private mStIdx(1 To PERF_STEK_MAX) As Long
Private mStT0(1 To PERF_STEK_MAX) As Double
Private mStDete(1 To PERF_STEK_MAX) As Double

' --- velicine tabela i sesija ---
Private mVelAkc As Object           ' tabela -> "RxC" (tekuca akcija)
Private mVelSes As Object           ' tabela -> "RxC" (najveca vidjena)
Private mSesAkc As Object           ' akcija -> Array(broj, zbir, min, max, prvi)
Private mSesSek As Object           ' akcija|oznaka -> Array(poz, uk, self, max, kol, vrsta)
Private mBrojIstog As Object        ' akcija -> koliko puta od ukljucenja

Private mFajl As String

'=====================================================================
' UKLJUCIVANJE
'=====================================================================

' Jednom, iz modMain.StartApp (posle LogAppStart). GetLocalConfigValue cita
' tblLocalConfig celiju po celiju, pa se nikad ne zove po merenju.
Public Sub PerfInitIzKonfiga()
    Dim v As String
    On Error Resume Next
    v = UCase$(Trim$(modSetup.GetLocalConfigValue(PERF_CFG_KLJUC, "NE")))
    If v = "DA" Then PerfUkljuci False
    If v = "DETALJ" Then PerfUkljuci True
    Err.Clear
End Sub

' Alt+F8. Vazi do zatvaranja sveske. Trajno, u Immediate:
'     SetLocalConfigValue "PERF_MERENJE", "DA"
Public Sub Perf_Ukljuci()
    PerfUkljuci False
    Debug.Print "modPerf: nivo 1. QPC=" & IIf(mImaQpc, "DA", "NE") & _
                " sekcija us=" & N3(mOverheadSekUs) & " brojac us=" & N3(mOverheadBrojUs)
End Sub

Public Sub Perf_UkljuciDetalj()
    PerfUkljuci True
    Debug.Print "modPerf: nivo 2. QPC=" & IIf(mImaQpc, "DA", "NE") & _
                " sekcija us=" & N3(mOverheadSekUs) & " brojac us=" & N3(mOverheadBrojUs)
End Sub

Public Sub Perf_Iskljuci()
    On Error Resume Next
    If gPerfOn And mAkcijaDub > 0 Then
        mNapomena = Trim$(mNapomena & " PREKINUTO")
        ZatvoriAkciju
    End If
    gPerfOn = False
    gPerfDetalj = False
    Debug.Print "modPerf: iskljuceno."
End Sub

Public Sub Perf_Reset()
    SesijaNova
    mAkcijaRb = 0
    Debug.Print "modPerf: sesija ociscena."
End Sub

' Rucni makro, nije u putu upisa; dijalog samo van test rezima.
Public Sub Perf_Sazetak()
    Dim txt As String, putanja As String
    On Error Resume Next
    txt = SazetakTekst()
    Debug.Print txt
    putanja = UpisiSazetak(txt)
    If Not IsTestMode() Then
        MsgBox "Sazetak merenja: " & IIf(Len(putanja) > 0, putanja, "nije upisan, v. Immediate"), _
               vbInformation, APP_NAME
    End If
End Sub

' Za testove i modTest.ResetSeamova: gasi bez upisa i bez dijaloga.
Public Sub PerfIskljuciTiho()
    gPerfOn = False
    gPerfDetalj = False
    mAkcijaDub = 0
    mIstoIme = 0
    mDub = 0
End Sub

'=====================================================================
' AKCIJA -- koren merenja (UiEvent, OnTime). Ugnjezdena akcija drugog
' imena postaje sekcija "AKCIJA.<naziv>" spoljne.
'=====================================================================

Public Sub PerfAkcijaPocni(ByVal naziv As String)
    Dim errNum As Long, errSrc As String, errDesc As String
    errNum = Err.Number: errSrc = Err.Source: errDesc = Err.Description
    On Error Resume Next
    If gPerfOn Then
        If mAkcijaDub > 0 Then
            If SadMs() - mAkcijaT0 > PERF_ZASTARELA_MS Then
                mNapomena = Trim$(mNapomena & " NEZAVRSENA")
                ZatvoriAkciju
            End If
        End If
        If mAkcijaDub = 0 Then
            ResetAkcije
            mAkcija = naziv
            mAkcijaRb = mAkcijaRb + 1
            mBrojIstog(naziv) = CLng(mBrojIstog(naziv)) + 1
            mAkcijaRbIstog = CLng(mBrojIstog(naziv))
            mAkcijaCalc = Application.Calculation
            mAkcijaDub = 1
            mAkcijaT0 = SadMs()
        Else
            mAkcijaDub = mAkcijaDub + 1
            If naziv = mAkcija Then
                mIstoIme = mIstoIme + 1
            Else
                PerfPocni PERF_UGNJ & naziv
            End If
        End If
    End If
    ErrVrati errNum, errSrc, errDesc
End Sub

Public Sub PerfAkcijaZavrsi(ByVal naziv As String)
    Dim errNum As Long, errSrc As String, errDesc As String
    errNum = Err.Number: errSrc = Err.Source: errDesc = Err.Description
    On Error Resume Next
    If gPerfOn And mAkcijaDub > 0 Then
        If naziv = mAkcija Then
            If mIstoIme > 0 Then
                mIstoIme = mIstoIme - 1
                mAkcijaDub = mAkcijaDub - 1
            Else
                ZatvoriAkciju
            End If
        Else
            PerfKraj PERF_UGNJ & naziv
            If mAkcijaDub > 1 Then mAkcijaDub = mAkcijaDub - 1
        End If
    End If
    ErrVrati errNum, errSrc, errDesc
End Sub

' Dopuna opisa tekuce akcije (npr. "rezim=OTKUP"). Bez On Error: ne dira Err.
Public Sub PerfNapomena(ByVal tekst As String)
    If Not gPerfOn Then Exit Sub
    If mAkcijaDub = 0 Then Exit Sub
    If InStr(1, mNapomena, tekst, vbBinaryCompare) = 0 Then mNapomena = Trim$(mNapomena & " " & tekst)
End Sub

'=====================================================================
' SEKCIJE I BROJACI
'=====================================================================

Public Sub PerfPocni(ByVal oznaka As String)
    Dim errNum As Long, errSrc As String, errDesc As String, k As Long
    errNum = Err.Number: errSrc = Err.Source: errDesc = Err.Description
    On Error Resume Next
    If gPerfOn Then
        If mAkcijaDub = 0 Then
            mVanAkcije = mVanAkcije + 1
        ElseIf mDub >= PERF_STEK_MAX Then
            mPreliv = mPreliv + 1
        Else
            k = IdxOznake(oznaka, PERF_VRSTA_SEK)
            If k > 0 Then
                mDub = mDub + 1
                mStIdx(mDub) = k
                mStDete(mDub) = 0
                mAkt(k) = mAkt(k) + 1
                mStT0(mDub) = SadMs()
            End If
        End If
    End If
    ErrVrati errNum, errSrc, errDesc
End Sub

' Trazi oznaku odozgo niz stek; sve sto je iznad nje ostalo je otvoreno
' (izlaz kroz gresku) i zatvara se sada, uz brojanje u Nezatvorenih.
Public Sub PerfKraj(ByVal oznaka As String)
    Dim errNum As Long, errSrc As String, errDesc As String
    Dim i As Long, nadjen As Long
    errNum = Err.Number: errSrc = Err.Source: errDesc = Err.Description
    On Error Resume Next
    If gPerfOn And mAkcijaDub > 0 Then
        For i = mDub To 1 Step -1
            If mLbl(mStIdx(i)) = oznaka Then
                nadjen = i
                Exit For
            End If
        Next i
        If nadjen = 0 Then
            mNeuparenih = mNeuparenih + 1
        Else
            Do While mDub > nadjen
                mNezatvorenih = mNezatvorenih + 1
                ZatvoriVrh
            Loop
            ZatvoriVrh
        End If
    End If
    ErrVrati errNum, errSrc, errDesc
End Sub

Public Sub PerfBroji(ByVal oznaka As String, Optional ByVal kolicina As Double = 0)
    Dim errNum As Long, errSrc As String, errDesc As String, k As Long
    errNum = Err.Number: errSrc = Err.Source: errDesc = Err.Description
    On Error Resume Next
    If gPerfOn Then
        If mAkcijaDub > 0 Then
            k = IdxOznake(oznaka, PERF_VRSTA_BROJ)
            If k > 0 Then
                mPoz(k) = mPoz(k) + 1
                mKol(k) = mKol(k) + kolicina
            End If
        Else
            mVanAkcije = mVanAkcije + 1
        End If
    End If
    ErrVrati errNum, errSrc, errDesc
End Sub

' Broji poziv (kolicina = broj celija) i pamti dimenzije tabele.
' v je ByRef: ByVal Variant bi kopirao ceo niz tabele.
Public Sub PerfVelicinaNiza(ByVal oznaka As String, ByVal tabela As String, ByRef v As Variant)
    Dim errNum As Long, errSrc As String, errDesc As String
    Dim r As Long, c As Long, k As Long
    errNum = Err.Number: errSrc = Err.Source: errDesc = Err.Description
    On Error Resume Next
    If gPerfOn And mAkcijaDub > 0 Then
        If IsArray(v) Then
            r = UBound(v, 1) - LBound(v, 1) + 1
            c = UBound(v, 2) - LBound(v, 2) + 1
            If Err.Number <> 0 Then
                Err.Clear
                c = 1
            End If
        ElseIf Not IsEmpty(v) Then
            r = 1
            c = 1
        End If
        k = IdxOznake(oznaka, PERF_VRSTA_BROJ)
        If k > 0 Then
            mPoz(k) = mPoz(k) + 1
            mKol(k) = mKol(k) + CDbl(r) * CDbl(c)
        End If
        mVelAkc(tabela) = CStr(r) & "x" & CStr(c)
        If r >= Val(CStr(mVelSes(tabela))) Then mVelSes(tabela) = CStr(r) & "x" & CStr(c)
    End If
    ErrVrati errNum, errSrc, errDesc
End Sub

' Broji poziv (kolicina = redova x kolona) i pamti dimenzije tabele, bez niza:
' pozivno mesto u GetTableData daje dimenzije opsega, ne povratni niz.
Public Sub PerfVelicina(ByVal oznaka As String, ByVal tabela As String, _
                        ByVal redova As Long, ByVal kolona As Long)
    Dim errNum As Long, errSrc As String, errDesc As String, k As Long
    errNum = Err.Number: errSrc = Err.Source: errDesc = Err.Description
    On Error Resume Next
    If gPerfOn And mAkcijaDub > 0 Then
        k = IdxOznake(oznaka, PERF_VRSTA_BROJ)
        If k > 0 Then
            mPoz(k) = mPoz(k) + 1
            mKol(k) = mKol(k) + CDbl(redova) * CDbl(kolona)
        End If
        mVelAkc(tabela) = CStr(redova) & "x" & CStr(kolona)
        If redova >= Val(CStr(mVelSes(tabela))) Then mVelSes(tabela) = CStr(redova) & "x" & CStr(kolona)
    End If
    ErrVrati errNum, errSrc, errDesc
End Sub

' "c7_3" -> "c#_#", "act:ispravblok:12" -> "act:ispravblok:#". Oznaka akcije
' ne sme da nosi redni broj reda ni ID: broj razlicitih oznaka mora ostati mali.
Public Function PerfTagKlasa(ByVal tag As String) As String
    Dim i As Long, ch As String, s As String, cif As Boolean
    For i = 1 To Len(tag)
        ch = Mid$(tag, i, 1)
        If ch Like "#" Then
            If Not cif Then s = s & "#"
            cif = True
        Else
            s = s & ch
            cif = False
        End If
    Next i
    PerfTagKlasa = Left$(s, 40)
End Function

Public Function PerfSadMs() As Double
    SatInit
    PerfSadMs = SadMs()
End Function

' Za test: vrednost iz poslednje zatvorene (ili tekuce) akcije; -1 = nema oznake.
Public Function PerfStatTest(ByVal oznaka As String, ByVal polje As String) As Double
    Dim k As Long
    PerfStatTest = -1
    If Not IsTestMode() Then Exit Function
    If mIdx Is Nothing Then Exit Function
    If Not mIdx.Exists(oznaka) Then Exit Function
    k = CLng(mIdx(oznaka))
    Select Case polje
        Case "poziva":   PerfStatTest = mPoz(k)
        Case "ukupno":   PerfStatTest = mUk(k)
        Case "self":     PerfStatTest = mSelf(k)
        Case "max":      PerfStatTest = mMax(k)
        Case "kolicina": PerfStatTest = mKol(k)
    End Select
End Function

Public Function PerfMetaTest(ByVal ime As String) As Long
    If Not IsTestMode() Then Exit Function
    Select Case ime
        Case "Neuparenih":   PerfMetaTest = mNeuparenih
        Case "Nezatvorenih": PerfMetaTest = mNezatvorenih
        Case "Preliv":       PerfMetaTest = mPreliv
        Case "Dubina":       PerfMetaTest = mDub
        Case "Akcija":       PerfMetaTest = mAkcijaDub
        Case "Qpc":          PerfMetaTest = IIf(mImaQpc, 1, 0)
    End Select
End Function

'=====================================================================
' PRIVATNO -- sat, Err, stek
'=====================================================================

Private Sub PerfUkljuci(ByVal detalj As Boolean)
    On Error Resume Next
    SatInit
    If mSesAkc Is Nothing Then SesijaNova
    mPrag = Val(modSetup.GetLocalConfigValue(PERF_CFG_PRAG, "10"))
    If mPrag < 0 Then mPrag = 0
    mAkcijaDub = 0
    mIstoIme = 0
    Kalibrisi
    gPerfOn = True
    gPerfDetalj = detalj
    Err.Clear
End Sub

Private Sub SatInit()
    Dim hz As Currency, c As Currency
    If mSatSpreman Then Exit Sub
    mTimerBaza = CDbl(Timer)
    mTimerPoslednji = mTimerBaza
    mTimerDana = 0
    mImaQpc = False
    mSatSpreman = True
    If QueryPerformanceFrequency(hz) = 0 Then Exit Sub
    If hz <= 0 Then Exit Sub
    If QueryPerformanceCounter(c) = 0 Then Exit Sub
    mQpcHz = hz
    mQpcBaza = c
    mMsPoJedinici = 1000# / CDbl(hz)
    mImaQpc = True
End Sub

' Milisekunde od ukljucenja. Razlika se racuna u Currency (tacno), pa tek onda
' prelazi u Double.
Private Function SadMs() As Double
    Dim c As Currency, t As Double
    If mImaQpc Then
        QueryPerformanceCounter c
        SadMs = CDbl(c - mQpcBaza) * mMsPoJedinici
    Else
        t = CDbl(Timer)
        If t < mTimerPoslednji Then mTimerDana = mTimerDana + 1
        mTimerPoslednji = t
        SadMs = (t - mTimerBaza + 86400# * mTimerDana) * 1000#
    End If
End Function

' Vraca Err pozivaoca tacno kakav je bio pre poziva merenja.
Private Sub ErrVrati(ByVal errNum As Long, ByVal errSrc As String, ByVal errDesc As String)
    Err.Clear
    If errNum <> 0 Then
        Err.Number = errNum
        Err.Source = errSrc
        Err.Description = errDesc
    End If
End Sub

Private Sub SesijaNova()
    Set mSesAkc = CreateObject("Scripting.Dictionary")
    Set mSesSek = CreateObject("Scripting.Dictionary")
    Set mBrojIstog = CreateObject("Scripting.Dictionary")
    Set mVelSes = CreateObject("Scripting.Dictionary")
End Sub

Private Sub ResetAkcije()
    If Not mSatSpreman Then SatInit
    If mSesAkc Is Nothing Then SesijaNova
    If mIdx Is Nothing Then
        Set mIdx = CreateObject("Scripting.Dictionary")
        ReDim mLbl(1 To PERF_POC_KAP)
        ReDim mVrsta(1 To PERF_POC_KAP)
        ReDim mPoz(1 To PERF_POC_KAP)
        ReDim mKol(1 To PERF_POC_KAP)
        ReDim mUk(1 To PERF_POC_KAP)
        ReDim mSelf(1 To PERF_POC_KAP)
        ReDim mMax(1 To PERF_POC_KAP)
        ReDim mAkt(1 To PERF_POC_KAP)
    Else
        mIdx.RemoveAll
    End If
    If mVelAkc Is Nothing Then
        Set mVelAkc = CreateObject("Scripting.Dictionary")
    Else
        mVelAkc.RemoveAll
    End If
    mN = 0
    mDub = 0
    mTopSuma = 0
    mNeuparenih = 0
    mNezatvorenih = 0
    mPreliv = 0
    mNapomena = ""
    mIstoIme = 0
End Sub

Private Function IdxOznake(ByVal oznaka As String, ByVal vrsta As Long) As Long
    Dim k As Long, kap As Long
    If mIdx.Exists(oznaka) Then
        IdxOznake = CLng(mIdx(oznaka))
        Exit Function
    End If
    k = mN + 1
    kap = UBound(mLbl)
    If k > kap Then
        kap = kap * 2
        ReDim Preserve mLbl(1 To kap)
        ReDim Preserve mVrsta(1 To kap)
        ReDim Preserve mPoz(1 To kap)
        ReDim Preserve mKol(1 To kap)
        ReDim Preserve mUk(1 To kap)
        ReDim Preserve mSelf(1 To kap)
        ReDim Preserve mMax(1 To kap)
        ReDim Preserve mAkt(1 To kap)
    End If
    mN = k
    mLbl(k) = oznaka
    mVrsta(k) = vrsta
    mPoz(k) = 0: mKol(k) = 0: mUk(k) = 0: mSelf(k) = 0: mMax(k) = 0: mAkt(k) = 0
    mIdx.Add oznaka, k
    IdxOznake = k
End Function

Private Sub ZatvoriVrh()
    Dim k As Long, d As Double
    k = mStIdx(mDub)
    d = SadMs() - mStT0(mDub)
    If d < 0 Then d = 0
    mPoz(k) = mPoz(k) + 1
    mAkt(k) = mAkt(k) - 1
    If mAkt(k) <= 0 Then
        mAkt(k) = 0
        mUk(k) = mUk(k) + d
    End If
    mSelf(k) = mSelf(k) + d - mStDete(mDub)
    If d > mMax(k) Then mMax(k) = d
    mDub = mDub - 1
    If mDub > 0 Then
        mStDete(mDub) = mStDete(mDub) + d
    Else
        mTopSuma = mTopSuma + d
    End If
End Sub

Private Sub ZatvoriAkciju()
    Dim ukupno As Double
    ukupno = SadMs() - mAkcijaT0
    Do While mDub > 0
        mNezatvorenih = mNezatvorenih + 1
        ZatvoriVrh
    Loop
    mAkcijaDub = 0
    mIstoIme = 0
    ZapisiAkciju ukupno
End Sub

'=====================================================================
' IZLAZ
'=====================================================================

Private Sub ZapisiAkciju(ByVal ukupno As Double)
    Dim sb As String, pre As String, i As Long, kk As Variant, vr As String
    Dim cekanje As Double, oh As Double, aktivno As Double
    For i = 1 To mN
        If mVrsta(i) = PERF_VRSTA_SEK Then
            oh = oh + mPoz(i) * mOverheadSekUs / 1000#
            If mLbl(i) = PERF_CEKANJE Then cekanje = mUk(i)
        Else
            oh = oh + mPoz(i) * mOverheadBrojUs / 1000#
        End If
    Next i
    aktivno = ukupno - cekanje
    SesijaDodaj aktivno
    If ukupno < mPrag Or IsTestMode() Then
        mVanAkcije = 0
        Exit Sub
    End If

    pre = Format$(Now, "yyyy-mm-dd hh:nn:ss") & PERF_SEP & N3(mAkcijaT0) & PERF_SEP & _
          CStr(mAkcijaRb) & PERF_SEP & Csv(mAkcija) & PERF_SEP & CStr(mAkcijaRbIstog) & PERF_SEP
    sb = Red(pre, "AKCIJA", mAkcija, 1, ukupno, aktivno, ukupno, 0, mNapomena)
    sb = sb & Red(pre, "META", "CekaOperateraMs", 0, 0, 0, 0, cekanje, "")
    sb = sb & Red(pre, "META", "OverheadProcenaMs", 0, 0, 0, 0, oh, "")
    sb = sb & Red(pre, "META", "NemerenoMs", 0, 0, 0, 0, ukupno - mTopSuma, "")
    sb = sb & Red(pre, "META", "Calculation", 0, 0, 0, 0, CDbl(mAkcijaCalc), "")
    sb = sb & Red(pre, "META", "QpcHz", 0, 0, 0, 0, CDbl(mQpcHz) * 10000#, "")
    sb = sb & Red(pre, "META", "Nivo", 0, 0, 0, 0, IIf(gPerfDetalj, 2, 1), "")
    sb = sb & Red(pre, "META", "Neuparenih", 0, 0, 0, 0, CDbl(mNeuparenih), "")
    sb = sb & Red(pre, "META", "Nezatvorenih", 0, 0, 0, 0, CDbl(mNezatvorenih), "")
    sb = sb & Red(pre, "META", "Preliv", 0, 0, 0, 0, CDbl(mPreliv), "")
    sb = sb & Red(pre, "META", "VanAkcije", 0, 0, 0, 0, CDbl(mVanAkcije), "")
    sb = sb & Red(pre, "META", "TableCacheDepth", 0, 0, 0, 0, CDbl(modDataAccess.TableCacheDepth()), "")
    For i = 1 To mN
        If mVrsta(i) = PERF_VRSTA_SEK Then vr = "SEK" Else vr = "BROJ"
        sb = sb & Red(pre, vr, mLbl(i), mPoz(i), mUk(i), mSelf(i), mMax(i), mKol(i), "")
    Next i
    For Each kk In mVelAkc.keys
        sb = sb & Red(pre, "VEL", CStr(kk), 0, 0, 0, 0, 0, CStr(mVelAkc(kk)))
    Next kk
    UpisiUFajl sb
    LogInfo "modPerf", Csv(mAkcija) & " " & N3(ukupno) & " ms, aktivno " & N3(aktivno) & _
                       " ms, overhead~" & N3(oh) & " ms; " & TopSelf(3)
    mVanAkcije = 0
End Sub

Private Sub SesijaDodaj(ByVal aktivno As Double)
    Dim a As Variant, i As Long, k As String
    If mSesAkc Is Nothing Then SesijaNova
    If mSesAkc.Exists(mAkcija) Then
        a = mSesAkc(mAkcija)
        a(0) = a(0) + 1
        a(1) = a(1) + aktivno
        If aktivno < a(2) Then a(2) = aktivno
        If aktivno > a(3) Then a(3) = aktivno
    Else
        a = Array(1#, aktivno, aktivno, aktivno, aktivno)
    End If
    mSesAkc(mAkcija) = a
    For i = 1 To mN
        k = mAkcija & "|" & mLbl(i)
        If mSesSek.Exists(k) Then
            a = mSesSek(k)
            a(0) = a(0) + mPoz(i)
            a(1) = a(1) + mUk(i)
            a(2) = a(2) + mSelf(i)
            If mMax(i) > a(3) Then a(3) = mMax(i)
            a(4) = a(4) + mKol(i)
        Else
            a = Array(mPoz(i), mUk(i), mSelf(i), mMax(i), mKol(i), CDbl(mVrsta(i)))
        End If
        mSesSek(k) = a
    Next i
End Sub

Private Function TopSelf(ByVal n As Long) As String
    Dim uzeto() As Boolean, i As Long, j As Long, best As Long, s As String
    If mN = 0 Then Exit Function
    ReDim uzeto(1 To mN)
    For j = 1 To n
        best = 0
        For i = 1 To mN
            If Not uzeto(i) And mVrsta(i) = PERF_VRSTA_SEK Then
                If best = 0 Then
                    best = i
                ElseIf mSelf(i) > mSelf(best) Then
                    best = i
                End If
            End If
        Next i
        If best = 0 Then Exit For
        uzeto(best) = True
        s = s & mLbl(best) & " " & N3(mSelf(best)) & "; "
    Next j
    TopSelf = s
End Function

Private Function SazetakTekst() As String
    Dim s As String, kk As Variant, a As Variant, prosek As Double
    Dim kljucevi As Variant, uzeto() As Boolean, i As Long, j As Long
    Dim best As Long, bestV As Double, cnt As Long
    s = "modPerf sazetak " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    s = s & "QPC=" & IIf(mImaQpc, "DA", "NE (Timer)") & " Hz=" & N3(CDbl(mQpcHz) * 10000#) & _
        " overhead sekcija/brojac us=" & N3(mOverheadSekUs) & "/" & N3(mOverheadBrojUs) & _
        " nivo=" & IIf(gPerfDetalj, "2", "1") & vbCrLf
    If mSesAkc Is Nothing Then
        SazetakTekst = s
        Exit Function
    End If
    s = s & "AKCIJA;broj;prvi_ms;prosek_ostalih_ms;min_ms;max_ms" & vbCrLf
    For Each kk In mSesAkc.keys
        a = mSesAkc(kk)
        If a(0) > 1 Then prosek = (a(1) - a(4)) / (a(0) - 1) Else prosek = 0
        s = s & CStr(kk) & ";" & N3(a(0)) & ";" & N3(a(4)) & ";" & N3(prosek) & ";" & _
            N3(a(2)) & ";" & N3(a(3)) & vbCrLf
    Next kk
    cnt = mSesSek.count
    If cnt > 0 Then
        kljucevi = mSesSek.keys
        ReDim uzeto(0 To cnt - 1)
        s = s & "TOP self (akcija|oznaka);self_ms;ukupno_ms;poziva;max_ms" & vbCrLf
        For j = 1 To PERF_TOP_N
            best = -1
            bestV = -1
            For i = 0 To cnt - 1
                If Not uzeto(i) Then
                    a = mSesSek(kljucevi(i))
                    If a(5) = PERF_VRSTA_SEK And a(2) > bestV Then
                        best = i
                        bestV = a(2)
                    End If
                End If
            Next i
            If best < 0 Then Exit For
            uzeto(best) = True
            a = mSesSek(kljucevi(best))
            s = s & CStr(kljucevi(best)) & ";" & N3(a(2)) & ";" & N3(a(1)) & ";" & _
                N3(a(0)) & ";" & N3(a(3)) & vbCrLf
        Next j
        s = s & "BROJACI (akcija|oznaka);poziva;kolicina" & vbCrLf
        For i = 0 To cnt - 1
            a = mSesSek(kljucevi(i))
            If a(5) = PERF_VRSTA_BROJ Then
                s = s & CStr(kljucevi(i)) & ";" & N3(a(0)) & ";" & N3(a(4)) & vbCrLf
            End If
        Next i
    End If
    s = s & "VELICINE (tabela;najveca vidjena RxC)" & vbCrLf
    For Each kk In mVelSes.keys
        s = s & CStr(kk) & ";" & CStr(mVelSes(kk)) & vbCrLf
    Next kk
    SazetakTekst = s
End Function

' Dir/MkDir samo kad se promeni ime fajla (prvi upis u danu), ne po akciji.
Private Sub UpisiUFajl(ByVal tekst As String)
    Dim folder As String, fajl As String, ff As Integer, nov As Boolean
    On Error GoTo Kraj
    folder = ThisWorkbook.Path & "\" & PERF_FOLDER
    fajl = folder & "\" & PERF_PREFIKS & Format$(Date, "yyyy-mm-dd") & ".csv"
    If fajl <> mFajl Then
        If Len(Dir(folder, vbDirectory)) = 0 Then MkDir folder
        nov = (Len(Dir(fajl)) = 0)
        mFajl = fajl
    End If
    ff = FreeFile
    Open fajl For Append As #ff
    If nov Then Print #ff, ZaglavljeCsv()
    Print #ff, tekst;
    Close #ff
    Exit Sub
Kraj:
    On Error Resume Next
    If ff <> 0 Then Close #ff
    mFajl = ""
End Sub

Private Function UpisiSazetak(ByVal txt As String) As String
    Dim folder As String, fajl As String, ff As Integer
    On Error GoTo Kraj
    If IsTestMode() Then Exit Function
    folder = ThisWorkbook.Path & "\" & PERF_FOLDER
    If Len(Dir(folder, vbDirectory)) = 0 Then MkDir folder
    fajl = folder & "\" & PERF_PREFIKS & "sazetak_" & Format$(Now, "yyyy-mm-dd_hhnnss") & ".txt"
    ff = FreeFile
    Open fajl For Output As #ff
    Print #ff, txt;
    Close #ff
    UpisiSazetak = fajl
    Exit Function
Kraj:
    On Error Resume Next
    If ff <> 0 Then Close #ff
End Function

Private Function ZaglavljeCsv() As String
    ZaglavljeCsv = "datum;t0_ms;rb_akcije;akcija;rb_istog;vrsta;oznaka;poziva;" & _
                   "ukupno_ms;self_ms;max_ms;kolicina;napomena"
End Function

Private Function Red(ByVal pre As String, ByVal vrsta As String, ByVal oznaka As String, _
                     ByVal poziva As Double, ByVal uk As Double, ByVal slf As Double, _
                     ByVal mx As Double, ByVal kol As Double, ByVal napomena As String) As String
    Red = pre & vrsta & PERF_SEP & Csv(oznaka) & PERF_SEP & N3(poziva) & PERF_SEP & _
          N3(uk) & PERF_SEP & N3(slf) & PERF_SEP & N3(mx) & PERF_SEP & N3(kol) & _
          PERF_SEP & Csv(napomena) & vbCrLf
End Function

' Str$ umesto Format$: tacka bez obzira na regionalna podesavanja (obrazac BFP).
Private Function N3(ByVal x As Double) As String
    N3 = Trim$(Str$(Round(x, 3)))
End Function

Private Function Csv(ByVal s As String) As String
    s = Replace(s, PERF_SEP, ",")
    s = Replace(s, vbCr, " ")
    Csv = Replace(s, vbLf, " ")
End Function

' Cena jednog para Pocni/Kraj i jednog brojaca, sa sve cuvanjem Err.
' Ne meri If na pozivnom mestu ni spajanje oznake -- procena je DONJA granica.
Private Sub Kalibrisi()
    Dim i As Long, t0 As Double
    gPerfOn = True
    ResetAkcije
    mAkcija = "PERF.KALIBRACIJA"
    mAkcijaDub = 1
    mAkcijaT0 = SadMs()
    t0 = SadMs()
    For i = 1 To PERF_KALIB_PAROVA
        PerfPocni "PERF.K"
        PerfKraj "PERF.K"
    Next i
    mOverheadSekUs = (SadMs() - t0) * 1000# / PERF_KALIB_PAROVA
    t0 = SadMs()
    For i = 1 To PERF_KALIB_PAROVA
        PerfBroji "PERF.B"
    Next i
    mOverheadBrojUs = (SadMs() - t0) * 1000# / PERF_KALIB_PAROVA
    mAkcijaDub = 0
    ResetAkcije
    gPerfOn = False
End Sub
