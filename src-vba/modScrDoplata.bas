Attribute VB_Name = "modScrDoplata"
'=====================================================================
' modScrDoplata - ekran "Doplata kooperantima" (knjizno odobrenje, KO).
' Domen i odluke: docs/DOMEN/DOPLATA.md (DOP-ODL-*, DOP-*). Faza F3.
'
' Ljuska ga ne poznaje po imenu: dobija ga preko Application.Run, da klijent
' kome ovaj modul nedostaje i dalje radi (zamka #19).
'
' STA JE OVDE, A STA NIJE: ovde je REDOSLED i PRIKAZ. Racun, kapije, upis,
' storno i papir nisu ovde:
'   - pregled obracuna       -> modDoplata.DoplataPlan (ISTI plan koji pisac izvrsava)
'   - provera stopa          -> modDoplata.DoplataStopeProblem (isto pravilo kao pisac)
'   - ponavljanje (pitanje)  -> modDoplata.DoplataPreklapanja
'   - knjizenje              -> modDoplata.KnjiziDoplatu_TX
'   - storno KO / obracuna   -> modStorno.StornoDoplata_TX / StornoObracunDoplate_TX
'   - stampa KO              -> modPrint.StampajKO
'
' STORNO KO JE OVDE, A NE NA EKRANU STORNO. Prekidac lista tamo nosi 11 lista,
' a to je MAX_SEG (modOtkupUI) -- dvanaesta ne staje. KO je dokument jednog
' kooperanta bez nizvodnog lanca robe, pa mu ne treba ni pregled posledica po
' modovima koji je razlog postojanja ekrana Storno.
'
' TRI LISTE u deljenoj mrezi:
'   STOPE  stope koje cekaju knjizenje (korpa); radnja: ukloni stopu
'   PLAN   pregled obracuna za unet period i stope -- jedan red po kooperantu,
'          stanici i kljucu, tacno ono sto ce "Knjizi" upisati
'   KO     aktivna knjizna odobrenja; radnje: stampaj, storniraj KO, storniraj
'          ceo obracun
'
' PITANJA OPERATERU idu kroz Pitaj, obrazac KI-015: u test rezimu pitanje se
' BELEZI i odgovor je uvek "Ne" -- MsgBox u nevidljivom Excelu visi. Put "Da"
' je KnjiziJezgro, koje test zove direktno (Scr_DopKnjiziTest), a pisac sa
' potvrdom meri test 312.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Private Const EKRAN As String = "DOPLATA"

' Kljucevi lista
Private Const DOP_STOPE As String = "STOPE"
Private Const DOP_PLAN  As String = "PLAN"
Private Const DOP_KO    As String = "KO"

' Zona: dva reda polja, red objasnjenja i red dugmadi.
Private Const DOP_ZONA_H  As Single = 172
Private Const DOP_GAP     As Single = 10
Private Const DOP_FLD_H   As Single = 28
Private Const DOP_Y_LBL_A As Single = 22
Private Const DOP_Y_LBL_B As Single = 72
Private Const DOP_Y_FLD_B As Single = 88
Private Const DOP_Y_HINT  As Single = 122
Private Const DOP_Y_BTN   As Single = 142
Private Const DOP_BTN_H   As Single = 24

' Kolone koje nose IDENTITET reda (nevidljive, prioritet 4). Jedan broj deljen
' izmedju opisa kolona, punjenja redova i radnje -- isti razlog kao
' ST_BLOK_COL_ID na ekranu Storno.
Private Const ST_COL_KLJUC As Long = 5
Private Const KO_COL_ID    As Long = 7
Private Const KO_COL_OBR   As Long = 8

Private mLista As String
Private mStope As Collection        ' Dictionary: VrstaVoca, SortaVoca, Klasa, Cena
Private mPlan As Variant            ' poslednji pregled (DoplataPlan) ili Empty
Private mPlanBez As String          ' stope bez otkupa iz poslednjeg pregleda
Private mFill As Boolean            ' punjenje kontrola okida Change
Private mCombosPunjeni As Boolean
Private mStep As String

' TEST SEAM (KI-015 obrazac): poslednje pitanje koje bi ekran postavio. Seam
' samo POSMATRA -- odgovor u test rezimu je uvek "Ne". Gasi ga Scr_DopTestReset,
' koji je i u modTest.ResetSeamova (seam nosi stanje).
Private mPitanjeTest As String

' TEST SEAM: unos zone bez zone (forma se u testu ne crta). Postavlja ga samo
' Scr_DopUnosTestSet, tvrdo gejtovan; citalac ponovo pita IsTestMode().
Private mUnosTest As Object

'--------------------------------------------------------- UGOVOR EKRANA
Public Function Scr_Meta() As String
    Scr_Meta = "kljuc=" & EKRAN & "|naslov=OTKUI_NAV_DOPLATA|sub=OTKUI_SCRDOP_SUB" & _
               "|lista=OTKUI_SCRDOP_LISTA|oblik=zona+mreza|upis=zona"
End Function

' Stope su prve: to je ono sto operater u tom trenutku menja. Plan je pregled
' pre knjizenja, KO su knjizena odobrenja.
Public Function Scr_Liste() As Variant
    Scr_Liste = Array( _
        DOP_STOPE & "|OTKUI_SEG_DOP_STOPE|OTKUI_GRID_TITLE_DOP_STOPE|58", _
        DOP_PLAN & "|OTKUI_SEG_DOP_PLAN|OTKUI_GRID_TITLE_DOP_PLAN|64", _
        DOP_KO & "|OTKUI_SEG_DOP_KO|OTKUI_GRID_TITLE_DOP_KO|74")
End Function

Public Function Scr_Lista() As String
    If Len(mLista) = 0 Then mLista = DOP_STOPE
    Scr_Lista = mLista
End Function

' Koliko stavki CEKA operatera: stope koje nisu knjizene.
Public Function Scr_Brojac() As Long
    Scr_Brojac = BrojStopa()
End Function

Public Function Scr_Radnje() As String
    Select Case Scr_Lista()
        Case DOP_STOPE
            Scr_Radnje = "dopdel:OTKUI_BTN_DOP_UKLONI:110:danger:1"
        Case DOP_KO
            Scr_Radnje = "dopprint:OTKUI_BTN_DOP_STAMPA:104:ghost:1|" & _
                         "dopstorno:OTKUI_BTN_DOP_STORNO:118:danger:1|" & _
                         "dopstornoobr:OTKUI_BTN_DOP_STORNO_OBR:150:danger:1"
    End Select
End Function

'=====================================================================
' ZONA
'=====================================================================
Public Sub Scr_Build(ByVal z As Object)
    ' Pad ne sme da izadje kroz Application.Run (KI-016, UI-PAD-01).
    On Error GoTo EH
    modUiScreens.ScrPadAkoTest "modScrDoplata.Scr_Build"

    ' Bela podloga ispod oba reda polja -- LABELA, ne Frame (v. modScrAgro:
    ' Frame se crta iznad bezprozorskih kontrola i pokrio bi natpise).
    modUiKit.NewLbl z, "dopBg", "", 0, 0, 100, 10, 8, False, 0, C_WHITE
    modUiKit.NewLbl z, "dopCap", UCase$(Poruka("OTKUI_SCRDOP_CAP")), PAD, 6, 320, 11, _
                    TS_MICRO, True, C_MUTED, -1

    ' Polja pravi ljuska (NewFieldG), ekran im samo kaze gde stoje. Prefiks
    ' "scr" je obavezan: bez njega promena teksta ide ljusci.
    ' --- red A: datum knjizenja i period sa satom (DOP-ODL-03a, 04) ---
    modOtkupUI.NewFieldG z, "scrDopDatum", Poruka("OTKUI_FLD_DOP_DATUM"), "txt", "", 1, False, False, "DOP"
    modOtkupUI.NewFieldG z, "scrDopOd", Poruka("OTKUI_FLD_DOP_OD"), "txt", "", 1, False, False, "DOP"
    modOtkupUI.NewFieldG z, "scrDopOdSat", Poruka("OTKUI_FLD_DOP_SAT"), "txt", "", 1, False, False, "DOP"
    modOtkupUI.NewFieldG z, "scrDopDo", Poruka("OTKUI_FLD_DOP_DO"), "txt", "", 1, False, False, "DOP"
    modOtkupUI.NewFieldG z, "scrDopDoSat", Poruka("OTKUI_FLD_DOP_SAT"), "txt", "", 1, False, False, "DOP"
    modOtkupUI.NewFieldG z, "scrDopNap", Poruka("OTKUI_FLD_DOP_NAPOMENA"), "txt", "", 1, False, False, "DOP"
    ' --- red B: stopa po kljucu cenovnika (DOP-ODL-01, 02) ---
    modOtkupUI.NewFieldG z, "scrDopVrsta", Poruka("OTKUI_FLD_DOP_VRSTA"), "cmb", "", 1, False, False, "DOP"
    modOtkupUI.NewFieldG z, "scrDopSorta", Poruka("OTKUI_FLD_DOP_SORTA"), "cmb", "", 1, False, False, "DOP"
    modOtkupUI.NewFieldG z, "scrDopKlasa", Poruka("OTKUI_FLD_DOP_KLASA"), "cmb", "", 1, False, False, "DOP"
    modOtkupUI.NewFieldG z, "scrDopCena", Poruka("OTKUI_FLD_DOP_STOPA"), "txt", "din/kg", 1, True, False, "DOP"
    modUiKit.BtnV z, "scrDopDodaj", Poruka("OTKUI_BTN_DOP_DODAJ"), 0, DOP_Y_FLD_B, _
                  130, DOP_FLD_H, "soft"

    ' --- objasnjenje: stanje pregleda i stope bez otkupa ---
    modUiKit.NewLbl z, "dopHint", "", PAD, DOP_Y_HINT, 600, 12, TS_META, False, C_MUTED, -1

    ' --- dugmad ---
    modUiKit.BtnV z, "scrDopPregled", Poruka("OTKUI_BTN_DOP_PREGLED"), PAD, DOP_Y_BTN, _
                  150, DOP_BTN_H, "soft"
    modUiKit.BtnV z, "scrDopKnjizi", Poruka("OTKUI_BTN_DOP_KNJIZI"), PAD + 158, DOP_Y_BTN, _
                  150, DOP_BTN_H, "primary"
    modUiKit.BtnV z, "scrDopOcisti", Poruka("OTKUI_BTN_DOP_OCISTI"), PAD + 316, DOP_Y_BTN, _
                  118, DOP_BTN_H, "ghost"

    modUiKit.NewLbl z, "dopLnB", "", 0, DOP_ZONA_H - 1, 100, 1, 8, False, 0, C_BORDER

    ' Podrazumevan unos: knjizenje danas, period od pocetka godine do juce
    ' kraja dana. Operater ga menja; ekran nista ne pretpostavlja o sezoni.
    PostaviPoljeU z, "scrDopDatum", Format$(Date, "dd.mm.yyyy")
    PostaviPoljeU z, "scrDopOd", Format$(DateSerial(Year(Date), 1, 1), "dd.mm.yyyy")
    PostaviPoljeU z, "scrDopOdSat", "00:00"
    PostaviPoljeU z, "scrDopDo", Format$(Date - 1, "dd.mm.yyyy")
    PostaviPoljeU z, "scrDopDoSat", "23:59"
    Exit Sub
EH:
    modUiScreens.ScrPadZabelezi "modScrDoplata.Scr_Build"
End Sub

Public Function Scr_Layout(ByVal z As Object, ByVal w As Single, ByVal h As Single) As Single
    RasporediPolja z, w
    Scr_Layout = DOP_ZONA_H
End Function

Private Sub RasporediPolja(ByVal z As Object, ByVal w As Single)
    Dim fw As Single, fwS As Single, X As Single
    On Error Resume Next
    If z Is Nothing Then Exit Sub
    If w < 200 Then Exit Sub

    z.Controls("dopBg").Left = PAD - 10
    z.Controls("dopBg").top = DOP_Y_LBL_A - 8
    z.Controls("dopBg").width = w - 2 * (PAD - 10)
    z.Controls("dopBg").Height = DOP_Y_BTN - DOP_Y_LBL_A - 2

    ' RED A: datum, od + sat, do + sat, napomena. Sat je uzi -- nosi "hh:mm".
    fwS = 64
    fw = (w - PAD * 2 - DOP_GAP * 5 - fwS * 2) / 4
    If fw < 90 Then fw = 90
    X = PAD
    PoljeX z, "scrDopDatum", X, fw, DOP_Y_LBL_A: X = X + fw + DOP_GAP
    PoljeX z, "scrDopOd", X, fw, DOP_Y_LBL_A: X = X + fw + DOP_GAP
    PoljeX z, "scrDopOdSat", X, fwS, DOP_Y_LBL_A: X = X + fwS + DOP_GAP
    PoljeX z, "scrDopDo", X, fw, DOP_Y_LBL_A: X = X + fw + DOP_GAP
    PoljeX z, "scrDopDoSat", X, fwS, DOP_Y_LBL_A: X = X + fwS + DOP_GAP
    PoljeX z, "scrDopNap", X, fw, DOP_Y_LBL_A

    ' RED B: vrsta, sorta, klasa, stopa, dugme "dodaj stopu".
    fw = (w - PAD * 2 - DOP_GAP * 4 - 130) / 4
    If fw < 90 Then fw = 90
    X = PAD
    PoljeX z, "scrDopVrsta", X, fw, DOP_Y_LBL_B: X = X + fw + DOP_GAP
    PoljeX z, "scrDopSorta", X, fw, DOP_Y_LBL_B: X = X + fw + DOP_GAP
    PoljeX z, "scrDopKlasa", X, fw, DOP_Y_LBL_B: X = X + fw + DOP_GAP
    PoljeX z, "scrDopCena", X, fw, DOP_Y_LBL_B: X = X + fw + DOP_GAP
    modUiKit.MoveBtn z, "scrDopDodaj", X, DOP_Y_FLD_B

    z.Controls("dopHint").width = w - PAD * 2

    ' "Isprazni stope" uz desnu ivicu: razdvaja radnju koja baca rad od onih
    ' koje prave dokument -- isti raspored kao na ekranu Agrohemija.
    modUiKit.MoveBtn z, "scrDopPregled", PAD, DOP_Y_BTN
    modUiKit.MoveBtn z, "scrDopKnjizi", PAD + 158, DOP_Y_BTN
    modUiKit.MoveBtn z, "scrDopOcisti", w - PAD - 118, DOP_Y_BTN

    z.Controls("dopLnB").width = w
End Sub

' Polje je OKVIR: pomera se i siri kao celina, unutrasnjost prerasporedjuje
' ljuska (isto kao modScrAgro.PoljeX).
Private Sub PoljeX(ByVal z As Object, ByVal nm As String, ByVal X As Single, _
                   ByVal w As Single, ByVal yLbl As Single)
    On Error Resume Next
    z.Controls(nm).Left = X
    z.Controls(nm).top = yLbl
    z.Controls(nm).width = w
    modOtkupUI.LayoutFieldInner z.Controls(nm)
End Sub

Private Function Zona() As Object
    On Error Resume Next
    Set Zona = modOtkupUI.ScreenZone(EKRAN)
End Function

' Unutrasnja kontrola polja (nmT) -- isti oblik koji trazi i panel za izbor.
Private Function Kontrola(ByVal nm As String) As Object
    Dim z As Object
    On Error Resume Next
    Set z = Zona()
    If z Is Nothing Then Exit Function
    Set Kontrola = z.Controls(nm).Controls(nm & "T")
End Function

Private Function PoljeTekst(ByVal nm As String) As String
    Dim c As Object
    On Error Resume Next
    Set c = Kontrola(nm)
    If c Is Nothing Then Exit Function
    PoljeTekst = Trim$(CStr(c.text))
End Function

Private Sub PostaviPoljeU(ByVal z As Object, ByVal nm As String, ByVal v As String)
    Dim pre As Boolean
    On Error Resume Next
    pre = mFill
    mFill = True
    z.Controls(nm).Controls(nm & "T").text = v
    mFill = pre
End Sub

'=====================================================================
' STANJE EKRANA
'=====================================================================
Private Function Stope() As Collection
    If mStope Is Nothing Then Set mStope = New Collection
    Set Stope = mStope
End Function

Private Function BrojStopa() As Long
    If mStope Is Nothing Then Exit Function
    BrojStopa = mStope.count
End Function

' Kljuc stope za prikaz i za uklanjanje -- vrsta|sorta|klasa, bez razlike u
' velicini slova, kao kljuc cenovnika (DOP-ODL-02).
Private Function KljucStope(ByVal s As Object) As String
    KljucStope = UCase$(Trim$(CStr(s("VrstaVoca")))) & "|" & _
                 UCase$(Trim$(CStr(s("SortaVoca")))) & "|" & UCase$(Trim$(CStr(s("Klasa"))))
End Function

' Pregled vise ne vazi cim se promene stope ili period: knjizenje bi inace
' operateru pokazalo jedno, a upisalo drugo.
Private Sub PonistiPregled()
    mPlan = Empty
    mPlanBez = ""
End Sub

'=====================================================================
' UNOS ZONE
'=====================================================================
' Vrednost polja zone -- ili iz seam-a, u test rezimu kad je postavljen.
Private Function Unos(ByVal nm As String) As String
    If IsTestMode() And Not mUnosTest Is Nothing Then
        If mUnosTest.Exists(nm) Then Unos = CStr(mUnosTest(nm))
        Exit Function
    End If
    Unos = PoljeTekst(nm)
End Function

' Sat iz teksta "hh:mm" (ili "hh", ili "hh.mm"); prazno = podrazumevano.
' -1 kad tekst nije sat.
Private Function SatIzTeksta(ByVal s As String, ByVal podrazumevano As Double) As Double
    Dim t As String, p As Long, hs As String, ms As String
    t = Replace(Trim$(s), ".", ":")
    If Len(t) = 0 Then
        SatIzTeksta = podrazumevano
        Exit Function
    End If
    p = InStr(t, ":")
    If p = 0 Then
        hs = t
        ms = "00"
    Else
        hs = Left$(t, p - 1)
        ms = Mid$(t, p + 1)
    End If
    If Not (hs Like "#" Or hs Like "##") Or Not (ms Like "##") Then
        SatIzTeksta = -1
        Exit Function
    End If
    If CLng(hs) > 23 Or CLng(ms) > 59 Then
        SatIzTeksta = -1
        Exit Function
    End If
    SatIzTeksta = CDbl(TimeSerial(CLng(hs), CLng(ms), 0))
End Function

' Period i datum knjizenja iz zone. False uz razlog za operatera kad nesto
' nije citljivo -- ekran ne pogadja datum.
Private Function UnosPerioda(ByRef outOd As Date, ByRef outDo As Date, _
                             ByRef outDatum As Date, ByRef outGreska As String) As Boolean
    Dim dOd As Double, dDo As Double, dDat As Double, sOd As Double, sDo As Double
    dDat = modOtkupUI.ParseDatum(Unos("scrDopDatum"))
    dOd = modOtkupUI.ParseDatum(Unos("scrDopOd"))
    dDo = modOtkupUI.ParseDatum(Unos("scrDopDo"))
    sOd = SatIzTeksta(Unos("scrDopOdSat"), 0#)
    sDo = SatIzTeksta(Unos("scrDopDoSat"), CDbl(TimeSerial(23, 59, 0)))

    If dDat <= 0 Then
        outGreska = Poruka("OTKUI_ERR_DOP_DATUM")
    ElseIf dOd <= 0 Or dDo <= 0 Then
        outGreska = Poruka("OTKUI_ERR_DOP_PERIOD")
    ElseIf sOd < 0 Or sDo < 0 Then
        outGreska = Poruka("OTKUI_ERR_DOP_SAT")
    Else
        outDatum = CDate(Int(dDat))
        outOd = CDate(Int(dOd) + sOd)
        outDo = CDate(Int(dDo) + sDo)
        UnosPerioda = True
    End If
End Function

'=====================================================================
' DOGADJAJI
'=====================================================================
Public Function Scr_Event(ByVal tag As String, ByVal ev As String) As Boolean
    Dim errDesc As String
    On Error GoTo EH
    Scr_Event = ObradiKlik(tag)
    Err.Clear
    Exit Function
EH:
    errDesc = Err.description
    LogErr "modScrDoplata.Scr_Event"
    modOtkupUI.ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & errDesc, True
    Err.Clear
End Function

' Vraca True samo kad su se PODACI liste promenili (mreza se tada cita ponovo).
Private Function ObradiKlik(ByVal tag As String) As Boolean
    If Left$(tag, 2) = "ls" Then
        If Mid$(tag, 3) = Scr_Lista() Then Exit Function
        mLista = Mid$(tag, 3)
        ObradiKlik = True
        Exit Function
    End If
    If Left$(tag, 4) = "row:" Then Exit Function
    If Left$(tag, 4) = "dbl:" Then Exit Function
    If Left$(tag, 4) = "chg:" Then
        ObradiPromenu Mid$(tag, 5)
        Exit Function
    End If
    If Left$(tag, 4) = "act:" Then
        ObradiKlik = RadnjaNadRedom(tag)
        Exit Function
    End If

    Select Case tag
        Case "scrDopDodaj":   ObradiKlik = DodajStopu()
        Case "scrDopPregled": ObradiKlik = Pregledaj()
        Case "scrDopKnjizi":  ObradiKlik = KnjiziKlik()
        Case "scrDopOcisti":  ObradiKlik = IsprazniStope()
    End Select
End Function

' Tag je ime KONTROLE (nmT), ne polja -- tako ga ljuska i salje.
Private Sub ObradiPromenu(ByVal tag As String)
    If mFill Then Exit Sub
    Select Case tag
        Case "scrDopVrstaT"
            PuniSorte
        Case "scrDopDatumT", "scrDopOdT", "scrDopOdSatT", "scrDopDoT", "scrDopDoSatT"
            ' period je deo pregleda: promena ga ponistava (v. PonistiPregled)
            PonistiPregled
            OsveziZonu
    End Select
End Sub

' Radnja nad redom mreze: "act:<kljuc>:<red>". Red se bira po IDENTITETU koji
' nosi (nevidljiva kolona), ne po rednom broju -- mreza sortira.
Private Function RadnjaNadRedom(ByVal tag As String) As Boolean
    Dim p() As String, red As Long
    p = Split(Mid$(tag, 5), ":")
    If UBound(p) < 1 Then Exit Function
    red = CLng(val(p(1)))
    If red < 1 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_NEMA_REDA"), True
        Exit Function
    End If
    Select Case p(0)
        Case "dopdel"
            If Scr_Lista() = DOP_STOPE Then _
                RadnjaNadRedom = UkloniStopu(Trim$(CStr(modOtkupUI.GridCell(red, ST_COL_KLJUC))))
        Case "dopprint"
            If Scr_Lista() = DOP_KO Then StampajKO Trim$(CStr(modOtkupUI.GridCell(red, KO_COL_ID)))
        Case "dopstorno"
            If Scr_Lista() = DOP_KO Then _
                RadnjaNadRedom = StornirajKO(Trim$(CStr(modOtkupUI.GridCell(red, KO_COL_ID))))
        Case "dopstornoobr"
            If Scr_Lista() = DOP_KO Then _
                RadnjaNadRedom = StornirajObracun(Trim$(CStr(modOtkupUI.GridCell(red, KO_COL_OBR))))
        Case Else
            modOtkupUI.ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & p(0), True
    End Select
End Function

'=====================================================================
' RADNJE
'=====================================================================
' Pitanje operateru. U test rezimu se BELEZI i odgovor je "Ne" (KI-015).
Private Function Pitaj(ByVal tekst As String) As Boolean
    If IsTestMode() Then
        mPitanjeTest = tekst
        Exit Function
    End If
    Pitaj = (MsgBox(modOtkupUI.PorukaZaDijalog(tekst), vbQuestion + vbYesNo, APP_NAME) = vbYes)
End Function

Private Function DodajStopu() As Boolean
    Dim greska As String
    greska = DodajStopuJezgro(PoljeTekst("scrDopVrsta"), PoljeTekst("scrDopSorta"), _
                              PoljeTekst("scrDopKlasa"), _
                              modOtkupUI.ParseNum(PoljeTekst("scrDopCena")))
    If Len(greska) > 0 Then
        modOtkupUI.ShowToast greska, True
        Exit Function
    End If
    mLista = DOP_STOPE
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_DOP_STOPA_DODATA"), False
    DodajStopu = True
End Function

' Dodaje stopu posle iste provere koju pisac radi nad celim spiskom (duplikat,
' klasa, iznos > 0). "" kad je dodata, inace razlog.
Private Function DodajStopuJezgro(ByVal vrsta As String, ByVal sorta As String, _
                                  ByVal klasa As String, ByVal cena As Double) As String
    Dim s As Object, kandidat As Collection, i As Long, greska As String
    Set s = CreateObject("Scripting.Dictionary")
    s.Add "VrstaVoca", Trim$(vrsta)
    s.Add "SortaVoca", Trim$(sorta)
    s.Add "Klasa", Trim$(klasa)
    s.Add "Cena", cena

    Set kandidat = New Collection
    For i = 1 To BrojStopa()
        kandidat.Add mStope(i)
    Next i
    kandidat.Add s
    greska = modDoplata.DoplataStopeProblem(kandidat)
    If Len(greska) > 0 Then
        DodajStopuJezgro = ChrW(10007) & " " & greska
        Exit Function
    End If
    Stope().Add s
    PonistiPregled
End Function

Private Function UkloniStopu(ByVal kljuc As String) As Boolean
    Dim i As Long
    If Len(kljuc) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_NEMA_REDA"), True
        Exit Function
    End If
    For i = BrojStopa() To 1 Step -1
        If StrComp(KljucStope(mStope(i)), kljuc, vbTextCompare) = 0 Then
            mStope.Remove i
            PonistiPregled
            UkloniStopu = True
            Exit Function
        End If
    Next i
    modOtkupUI.ShowToast Poruka("OTKUI_ERR_NEMA_REDA"), True
End Function

Private Function IsprazniStope() As Boolean
    If BrojStopa() = 0 Then Exit Function
    If Not Pitaj(Poruka("OTKUI_ASK_DOP_OCISTI")) Then Exit Function
    Set mStope = New Collection
    PonistiPregled
    IsprazniStope = (Scr_Lista() <> DOP_KO)
End Function

' Pregled obracuna: isti plan koji ce pisac izvrsiti (DoplataPlan).
Private Function Pregledaj() As Boolean
    Dim greska As String
    greska = PregledJezgro()
    If Len(greska) > 0 Then
        modOtkupUI.ShowToast greska, True
        Exit Function
    End If
    mLista = DOP_PLAN
    Pregledaj = True
End Function

Private Function PregledJezgro() As String
    Dim pOd As Date, pDo As Date, dat As Date, greska As String
    PonistiPregled
    If BrojStopa() = 0 Then
        PregledJezgro = Poruka("OTKUI_ERR_DOP_BEZ_STOPA")
        Exit Function
    End If
    If Not UnosPerioda(pOd, pDo, dat, greska) Then
        PregledJezgro = greska
        Exit Function
    End If
    On Error GoTo EH
    mPlan = modDoplata.DoplataPlan(pOd, pDo, Stope(), mPlanBez)
    Exit Function
EH:
    PregledJezgro = ChrW(10007) & " " & Err.description
    PonistiPregled
End Function

' "Knjizi": pregled, pitanje sa brojkama, pitanje o ponavljanju, pa jezgro.
Private Function KnjiziKlik() As Boolean
    Dim greska As String, pOd As Date, pDo As Date, dat As Date
    Dim preklop As String, potvrda As Boolean

    greska = PregledJezgro()
    If Len(greska) > 0 Then
        modOtkupUI.ShowToast greska, True
        Exit Function
    End If
    If Len(mPlanBez) > 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_DOP_STOPA_BEZ_OTKUPA") & " " & mPlanBez, True
        mLista = DOP_PLAN
        KnjiziKlik = True
        Exit Function
    End If
    If Not IsArray(mPlan) Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_DOP_PRAZAN_PLAN"), True
        Exit Function
    End If

    If Not Pitaj(PitanjeKnjizenja()) Then Exit Function

    If UnosPerioda(pOd, pDo, dat, greska) Then
        preklop = modDoplata.DoplataPreklapanja(pOd, pDo, Stope())
        If Len(preklop) > 0 Then
            If Not Pitaj(preklop) Then Exit Function
            potvrda = True
        End If
    End If
    KnjiziKlik = KnjiziJezgro(potvrda)
End Function

' Tekst pitanja pred knjizenje: koliko dokumenata, kg i dinara -- iz istog plana.
' Dokument je po BLOKU (DOPLATA.md par. 2a), pa se broje blokovi plana sa
' doplatom (kolona 11, status 16); red "bez doplate" se ne knjizi.
Private Function PitanjeKnjizenja() As String
    Dim i As Long, blok As Object, kg As Double, din As Double
    Set blok = CreateObject("Scripting.Dictionary")
    blok.CompareMode = vbTextCompare
    If IsArray(mPlan) Then
        For i = 1 To UBound(mPlan, 1)
            If CLng(mPlan(i, 16)) = modDoplata.DOP_PLAN_KNJIZI Then
                blok(CStr(mPlan(i, 11))) = True
                kg = kg + CDbl(mPlan(i, 6))
                din = din + CDbl(mPlan(i, 8))
            End If
        Next i
    End If
    PitanjeKnjizenja = Poruka("OTKUI_ASK_DOP_KNJIZI") & " " & CStr(blok.count) & " KO, " & _
                       FmtKolicina(kg) & " kg, " & Format$(din, "#,##0.00") & " din?"
End Function

' Upis: jedini poziv pisca sa ovog ekrana. Uspeh prazni stope i pregled i
' prebacuje na listu KO -- sledeci obracun je nov obracun.
Private Function KnjiziJezgro(ByVal potvrdaPonavljanja As Boolean) As Boolean
    Dim pOd As Date, pDo As Date, dat As Date, greska As String
    Dim obr As String, en As Long

    If Not UnosPerioda(pOd, pDo, dat, greska) Then
        modOtkupUI.ShowToast greska, True
        Exit Function
    End If
    obr = modDoplata.KnjiziDoplatu_TX(dat, pOd, pDo, Stope(), Unos("scrDopNap"), _
                                      potvrdaPonavljanja, greska, en)
    If Len(obr) = 0 Then
        modOtkupUI.ShowToast ChrW(10007) & " " & greska, True
        Exit Function
    End If

    Set mStope = New Collection
    PonistiPregled
    mLista = DOP_KO
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_DOP_KNJIZENO") & " " & _
                         CStr(modDoplata.AktivneDoplateObracuna(obr).count) & " KO", False
    KnjiziJezgro = True
End Function

Private Sub StampajKO(ByVal doplataID As String)
    Dim greska As String
    If Len(doplataID) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_NEMA_REDA"), True
        Exit Sub
    End If
    greska = modPrint.StampajKO(doplataID)
    If Len(greska) > 0 Then modOtkupUI.ShowToast greska, True
End Sub

Private Function StornirajKO(ByVal doplataID As String) As Boolean
    Dim broj As String
    If Len(doplataID) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_NEMA_REDA"), True
        Exit Function
    End If
    broj = NzToText(LookupValue(TBL_DOPLATA, COL_DOP_ID, doplataID, COL_DOP_BROJ))
    If Not Pitaj(Poruka("OTKUI_ASK_DOP_STORNO") & " " & broj & "?") Then Exit Function
    If Not modStorno.StornoDoplata_TX(doplataID) Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_DOP_STORNO") & " " & broj, True
        Exit Function
    End If
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_DOP_STORNIRANO") & " " & broj, False
    StornirajKO = True
End Function

Private Function StornirajObracun(ByVal obracunID As String) As Boolean
    Dim n As Long
    If Len(obracunID) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_NEMA_REDA"), True
        Exit Function
    End If
    n = modDoplata.AktivneDoplateObracuna(obracunID).count
    If Not Pitaj(Poruka("OTKUI_ASK_DOP_STORNO_OBR") & " " & CStr(n) & " KO?") Then Exit Function
    If Not modStorno.StornoObracunDoplate_TX(obracunID) Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_DOP_STORNO") & " " & obracunID, True
        Exit Function
    End If
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_DOP_STORNIRANO") & " " & CStr(n) & " KO", False
    StornirajObracun = True
End Function

'=====================================================================
' ZONA: COMBO-I I OBJASNJENJE
'=====================================================================
' Vrste i klase se pune jednom po otvaranju ekrana; sorte po izabranoj vrsti.
' Izvor je sifarnik kultura -- isti par (vrsta, sorta) koji nosi otkup, pa i
' kljuc cenovnika (DOP-ODL-02).
Private Sub PuniCombos()
    Dim z As Object, CB As Object, d As Variant, i As Long, cV As Long
    Dim vidjene As Object, v As String
    On Error GoTo EH
    If mCombosPunjeni Then Exit Sub
    Set z = Zona()
    If z Is Nothing Then Exit Sub

    mFill = True
    mStep = "vrste"
    Set CB = z.Controls("scrDopVrsta").Controls("scrDopVrstaT")
    CB.Clear
    Set vidjene = CreateObject("Scripting.Dictionary")
    vidjene.CompareMode = vbTextCompare
    d = GetTableData(TBL_KULTURE)
    If IsArray(d) Then
        cV = RequireColumnIndex(TBL_KULTURE, COL_KUL_VRSTA, "modScrDoplata.PuniCombos")
        For i = 1 To UBound(d, 1)
            v = Trim$(NzToText(d(i, cV)))
            If Len(v) > 0 And Not vidjene.Exists(v) Then
                vidjene.Add v, True
                CB.AddItem v
            End If
        Next i
    End If

    mStep = "klase"
    Set CB = z.Controls("scrDopKlasa").Controls("scrDopKlasaT")
    CB.Clear
    CB.AddItem KLASA_I
    CB.AddItem KLASA_II

    mCombosPunjeni = True
    mFill = False
    Exit Sub
EH:
    mFill = False
    Debug.Print "modScrDoplata.PuniCombos PAO na koraku [" & mStep & "]: " & _
                Err.Number & " " & Err.description
End Sub

Private Sub PuniSorte()
    Dim z As Object, CB As Object, d As Variant, i As Long, cV As Long, cS As Long
    Dim vrsta As String, vidjene As Object, so As String, pre As Boolean
    On Error GoTo EH
    Set z = Zona()
    If z Is Nothing Then Exit Sub
    vrsta = PoljeTekst("scrDopVrsta")
    pre = mFill
    mFill = True
    Set CB = z.Controls("scrDopSorta").Controls("scrDopSortaT")
    CB.Clear
    Set vidjene = CreateObject("Scripting.Dictionary")
    vidjene.CompareMode = vbTextCompare
    d = GetTableData(TBL_KULTURE)
    If IsArray(d) And Len(vrsta) > 0 Then
        cV = RequireColumnIndex(TBL_KULTURE, COL_KUL_VRSTA, "modScrDoplata.PuniSorte")
        cS = RequireColumnIndex(TBL_KULTURE, COL_KUL_SORTA, "modScrDoplata.PuniSorte")
        For i = 1 To UBound(d, 1)
            If StrComp(Trim$(NzToText(d(i, cV))), vrsta, vbTextCompare) = 0 Then
                so = Trim$(NzToText(d(i, cS)))
                If Not vidjene.Exists(so) Then
                    vidjene.Add so, True
                    CB.AddItem so
                End If
            End If
        Next i
    End If
    mFill = pre
    Exit Sub
EH:
    mFill = pre
    Debug.Print "modScrDoplata.PuniSorte PAO: " & Err.Number & " " & Err.description
End Sub

Private Sub OsveziZonu()
    Dim z As Object, t As String, i As Long, kg As Double, din As Double
    Dim koop As Object
    On Error Resume Next
    Set z = Zona()
    If z Is Nothing Then Exit Sub
    PuniCombos

    t = Poruka("OTKUI_LBL_DOP_STOPA") & " " & CStr(BrojStopa())
    If IsArray(mPlan) Then
        Set koop = CreateObject("Scripting.Dictionary")
        For i = 1 To UBound(mPlan, 1)
            koop(CStr(mPlan(i, 1))) = True
            kg = kg + CDbl(mPlan(i, 6))
            din = din + CDbl(mPlan(i, 8))
        Next i
        t = t & "   " & ChrW(183) & "   " & Poruka("OTKUI_LBL_DOP_PREGLED") & " " & _
            CStr(koop.count) & " KO, " & FmtKolicina(kg) & " kg, " & _
            Format$(din, "#,##0.00") & " din"
    Else
        t = t & "   " & ChrW(183) & "   " & Poruka("OTKUI_LBL_DOP_BEZ_PREGLEDA")
    End If
    If Len(mPlanBez) > 0 Then
        t = t & "   " & ChrW(183) & "   " & Poruka("OTKUI_ERR_DOP_STOPA_BEZ_OTKUPA") & " " & mPlanBez
    End If
    z.Controls("dopHint").caption = t
    z.Controls("dopHint").ForeColor = IIf(Len(mPlanBez) > 0, C_RUST, C_MUTED)
End Sub

'=====================================================================
' REDOVI ZA MREZU
'=====================================================================
Public Function Scr_Rows(ByVal filter As String, ByVal q As String) As Variant
    ' Pad ne sme da izadje kroz Application.Run (KI-016, UI-PAD-01).
    On Error GoTo EH
    modUiScreens.ScrPadAkoTest "modScrDoplata.Scr_Rows"
    OsveziZonu
    Select Case Scr_Lista()
        Case DOP_PLAN: Scr_Rows = RedoviPlan(q)
        Case DOP_KO:   Scr_Rows = RedoviKO(q)
        Case Else:     Scr_Rows = RedoviStope(q)
    End Select
    Exit Function
EH:
    modUiScreens.ScrPadZabelezi "modScrDoplata.Scr_Rows"
End Function

Private Function PrazanRezultat(ByVal kolone As Variant) As Variant
    PrazanRezultat = Array(kolone, Empty, 0, 0#, 0#, Array(0, 0, 0))
End Function

'---------------------------------------------------------- LISTA: STOPE
Private Function StopeKolone() As Variant
    StopeKolone = Array( _
        "OTKUI_HD_VRSTA||txt|0|1", _
        "OTKUI_HD_SORTA||txt|140|1", _
        "OTKUI_HD_KLASA||txt|60|1", _
        "OTKUI_HD_STOPA||rsd|100|1", _
        "OTKUI_HDI_REF||txt|1|4")
End Function

Private Function RedoviStope(ByVal q As String) As Variant
    Dim outA() As Variant, n As Long, i As Long, s As Object, hay As String
    If BrojStopa() = 0 Then
        RedoviStope = PrazanRezultat(StopeKolone())
        Exit Function
    End If
    ReDim outA(1 To mStope.count, 1 To 5)
    For i = 1 To mStope.count
        Set s = mStope(i)
        hay = CStr(s("VrstaVoca")) & "|" & CStr(s("SortaVoca"))
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeci
        End If
        n = n + 1
        outA(n, 1) = CStr(s("VrstaVoca"))
        outA(n, 2) = CStr(s("SortaVoca"))
        outA(n, 3) = CStr(s("Klasa"))
        outA(n, 4) = CDbl(s("Cena"))
        outA(n, ST_COL_KLJUC) = KljucStope(s)
Sledeci:
    Next i
    RedoviStope = Array(StopeKolone(), outA, n, 0#, 0#, Array(0, 0, 0))
End Function

'----------------------------------------------------------- LISTA: PLAN
Private Function PlanKolone() As Variant
    PlanKolone = Array( _
        "OTKUI_HD_KOOP||txt|0|1", _
        "OTKUI_HD_OM||txt|130|2", _
        "OTKUI_HD_VRSTA||txt|110|1", _
        "OTKUI_HD_SORTA||txt|100|2", _
        "OTKUI_HD_KLASA||txt|50|1", _
        "OTKUI_HD_KG||kg|84|1", _
        "OTKUI_HD_STOPA||rsd|80|1", _
        "OTKUI_HD_IZNOS||rsd|96|1", _
        "OTKUI_HD_OTKUPA||num|60|3")
End Function

Private Function RedoviPlan(ByVal q As String) As Variant
    Dim outA() As Variant, n As Long, i As Long, hay As String
    Dim imena As Object, stanice As Object, sumKg As Double, sumIzn As Double
    Dim ime As String, sta As String
    If Not IsArray(mPlan) Then
        RedoviPlan = PrazanRezultat(PlanKolone())
        Exit Function
    End If
    Set imena = BuildLookupDict(TBL_KOOPERANTI, COL_KOOP_ID, "Ime", "Prezime")
    Set stanice = BuildLookupDict(TBL_STANICE, COL_STA_ID, "Naziv")
    ReDim outA(1 To UBound(mPlan, 1), 1 To 9)
    For i = 1 To UBound(mPlan, 1)
        ime = CStr(mPlan(i, 1))
        If imena.Exists(ime) Then ime = Trim$(CStr(imena(ime)))
        sta = CStr(mPlan(i, 2))
        If stanice.Exists(sta) Then sta = Trim$(CStr(stanice(sta)))
        hay = ime & "|" & sta & "|" & CStr(mPlan(i, 3)) & "|" & CStr(mPlan(i, 4))
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeci
        End If
        n = n + 1
        outA(n, 1) = ime
        outA(n, 2) = sta
        outA(n, 3) = CStr(mPlan(i, 3))
        outA(n, 4) = CStr(mPlan(i, 4))
        outA(n, 5) = CStr(mPlan(i, 5))
        outA(n, 6) = CDbl(mPlan(i, 6))
        outA(n, 7) = CDbl(mPlan(i, 7))
        outA(n, 8) = CDbl(mPlan(i, 8))
        outA(n, 9) = CLng(mPlan(i, 9))
        sumKg = sumKg + CDbl(mPlan(i, 6))
        sumIzn = sumIzn + CDbl(mPlan(i, 8))
Sledeci:
    Next i
    RedoviPlan = Array(PlanKolone(), outA, n, sumKg, sumIzn, Array(0, 0, 0))
End Function

'------------------------------------------------------------- LISTA: KO
Private Function KOKolone() As Variant
    KOKolone = Array( _
        "OTKUI_HD_BROJ||txt|100|1", _
        "OTKUI_HD_DATUM||txt|80|1", _
        "OTKUI_HD_KOOP||txt|0|1", _
        "OTKUI_HD_IZNOS||rsd|100|1", _
        "OTKUI_HD_PERIOD||txt|190|2", _
        "OTKUI_HD_STAVKI||num|56|3", _
        "OTKUI_HDI_REF||txt|1|4", _
        "OTKUI_HD_OBRACUN||txt|1|4")
End Function

' Jedan red po AKTIVNOM KO, iz istog citaca koji puni karticu i saldo.
Private Function RedoviKO(ByVal q As String) As Variant
    Dim r As Variant, i As Long, ko As Object, red As Variant, id As String
    Dim outA() As Variant, n As Long, k As Variant, imena As Object, ime As String
    Dim hay As String, sumIzn As Double

    r = modDoplata.DoplataStavkeRedovi()
    If Not IsArray(r) Then
        RedoviKO = PrazanRezultat(KOKolone())
        Exit Function
    End If

    Set ko = CreateObject("Scripting.Dictionary")
    ko.CompareMode = vbTextCompare
    For i = 1 To UBound(r, 1)
        id = CStr(r(i, 1))
        If Not ko.Exists(id) Then
            ko.Add id, Array(r(i, 2), r(i, 3), r(i, 4), 0#, r(i, 14), r(i, 15), 0&, r(i, 5))
        End If
        red = ko(id)
        red(3) = CDbl(red(3)) + CDbl(r(i, 12))
        red(6) = CLng(red(6)) + 1
        ko(id) = red
    Next i

    Set imena = BuildLookupDict(TBL_KOOPERANTI, COL_KOOP_ID, "Ime", "Prezime")
    ReDim outA(1 To ko.count, 1 To 8)
    For Each k In ko.keys
        red = ko(k)
        ime = CStr(red(2))
        If imena.Exists(ime) Then ime = Trim$(CStr(imena(ime)))
        hay = CStr(red(0)) & "|" & ime
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeci
        End If
        n = n + 1
        outA(n, 1) = CStr(red(0))
        outA(n, 2) = Format$(CDate(red(1)), "dd.mm.yyyy")
        outA(n, 3) = ime
        outA(n, 4) = modNovac.ZaokruziNovac(CDbl(red(3)))
        outA(n, 5) = Format$(CDate(red(4)), "dd.mm.yyyy hh:nn") & " " & ChrW(8211) & " " & _
                     Format$(CDate(red(5)), "dd.mm.yyyy hh:nn")
        outA(n, 6) = CLng(red(6))
        outA(n, KO_COL_ID) = CStr(k)
        outA(n, KO_COL_OBR) = CStr(red(7))
        sumIzn = sumIzn + CDbl(outA(n, 4))
Sledeci:
    Next k
    RedoviKO = Array(KOKolone(), outA, n, 0#, sumIzn, Array(0, 0, 0))
End Function

'=====================================================================
' TEST SEAM-ovi -- svi tvrdo gejtovani
'=====================================================================
' Ceo ekran u pocetno stanje: stope, pregled, lista, unos i zabelezeno pitanje.
' Zove ga i modTest.ResetSeamova.
Public Sub Scr_DopTestReset()
    If Not IsTestMode() Then Exit Sub
    Set mStope = Nothing
    PonistiPregled
    mLista = ""
    Set mUnosTest = Nothing
    mPitanjeTest = ""
End Sub

' Unos zone bez zone. Kljucevi su imena polja (scrDopDatum, scrDopOd,
' scrDopOdSat, scrDopDo, scrDopDoSat, scrDopNap).
Public Sub Scr_DopUnosTestSet(ByVal datum As String, ByVal od As String, ByVal odSat As String, _
                              ByVal doDatum As String, ByVal doSat As String, ByVal napomena As String)
    If Not IsTestMode() Then Exit Sub
    Set mUnosTest = CreateObject("Scripting.Dictionary")
    mUnosTest("scrDopDatum") = datum
    mUnosTest("scrDopOd") = od
    mUnosTest("scrDopOdSat") = odSat
    mUnosTest("scrDopDo") = doDatum
    mUnosTest("scrDopDoSat") = doSat
    mUnosTest("scrDopNap") = napomena
End Sub

' Dodavanje stope istim putem kao dugme (ista provera); "" kad je dodata.
Public Function Scr_DopStopaTestDodaj(ByVal vrsta As String, ByVal sorta As String, _
                                      ByVal klasa As String, ByVal cena As Double) As String
    If Not IsTestMode() Then Exit Function
    Scr_DopStopaTestDodaj = DodajStopuJezgro(vrsta, sorta, klasa, cena)
End Function

Public Function Scr_DopBrojStopa() As Long
    Scr_DopBrojStopa = BrojStopa()
End Function

' Uklanjanje stope po KLJUCU reda (vrsta|sorta|klasa) -- isto kao radnja.
Public Function Scr_DopUkloniTest(ByVal kljuc As String) As Boolean
    If Not IsTestMode() Then Exit Function
    Scr_DopUkloniTest = UkloniStopu(kljuc)
End Function

' Pregled istim putem kao dugme; "" kad je prosao, inace razlog.
Public Function Scr_DopPregledTest() As String
    If Not IsTestMode() Then Exit Function
    Scr_DopPregledTest = PregledJezgro()
    If Len(Scr_DopPregledTest) = 0 Then mLista = DOP_PLAN
End Function

' Klik na "Knjizi" -- ceo put sa pitanjima. U test rezimu je odgovor "Ne", pa
' ovaj put meri da se pitanje POSTAVI i da bez potvrde nista ne nastane.
Public Function Scr_DopKnjiziKlikTest() As Boolean
    If Not IsTestMode() Then Exit Function
    Scr_DopKnjiziKlikTest = KnjiziKlik()
End Function

' Upis posle pitanja -- jezgro dugmeta, sa odgovorom na pitanje o ponavljanju.
Public Function Scr_DopKnjiziTest(ByVal potvrdaPonavljanja As Boolean) As Boolean
    If Not IsTestMode() Then Exit Function
    Scr_DopKnjiziTest = KnjiziJezgro(potvrdaPonavljanja)
End Function

' Storno KO istim putem kao radnja nad redom. U test rezimu je odgovor na
' pitanje "Ne", pa ovaj put meri da se pitanje POSTAVI i da KO ostane aktivan.
Public Function Scr_DopStornoTest(ByVal doplataID As String) As Boolean
    If Not IsTestMode() Then Exit Function
    Scr_DopStornoTest = StornirajKO(doplataID)
End Function

Public Function Scr_DopPitanjeTest() As String
    If Not IsTestMode() Then Exit Function
    Scr_DopPitanjeTest = mPitanjeTest
End Function

Public Sub Scr_DopListaTestSet(ByVal kljuc As String)
    If Not IsTestMode() Then Exit Sub
    mLista = kljuc
End Sub
