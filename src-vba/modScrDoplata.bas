Attribute VB_Name = "modScrDoplata"
'=====================================================================
' modScrDoplata - ekran "Doplata kooperantima" (dokument doplate po bloku).
' Domen i odluke: docs/DOMEN/DOPLATA.md (DOP-ODL-*, DOP-*). Faza F3, P2.
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
' PET LISTA u deljenoj mrezi (P2, odluka 05.11.2026; prikazi su ZBIRNI,
' DOP-ODL-18):
'   STOPE     stope koje cekaju knjizenje (korpa); radnja: ukloni stopu
'   PLAN      "Pregled": zbirno po kooperantu, kulturi i klasi, SAMO ukljuceni
'             blokovi sa doplatom -- tacno ono sto ce "Knjizi" upisati
'   BLOKOVI   jedan red po bloku i klasi poslednjeg pregleda, sa kvacicom; klik
'             na red ukljucuje/iskljucuje ceo blok (DOP-ODL-16); radnje:
'             ukljuci sve, iskljuci sve
'   KNJIZENO  zbirno po obracunu i kooperantu; radnja: storniraj obracun
'   KO        "Dokumenti": jedan red po dokumentu doplate (bloku); radnje:
'             stampaj, storniraj dokument, storniraj ceo obracun
'
' NACIN UNOSA (DOP-ODL-15) vazi za ceo obracun i menja se samo dok nema unetih
' stopa: ista vrednost u drugom nacinu znaci nesto drugo (10,80 doplate po kg
' nije 10,80 konacne cene). Ekran pamti ISKLJUCENE blokove (podrazumevano su
' svi ukljuceni); iskljucenje ostaje i kroz nov pregled, do knjizenja ili
' praznjenja stopa. Izbor ide piscu kao OtkupID-evi (KnjiziDoplatu_TX), a pisac
' izabrane ponovo proverava.
'
' OBUHVAT OBRACUNA (DOP-ODL-31): polje "Samo kooperant". Prazno = obracun za
' sve kooperante; izabran kooperant = obracun samo za njega. Ekran ga samo
' PROSLEDJUJE motoru (DoplataPlan / DoplataPreklapanja / KnjiziDoplatu_TX) -- plan
' tada vec nosi samo njegove blokove, pa liste, zbirovi i izbor blokova rade bez
' ijedne izmene. Tekst koji nije izabran iz liste NIJE "svi": pregled i knjizenje
' staju. Isti obuhvat postavlja i kartica kooperanta u Izvestajima
' (Scr_DopOtvoriZaKooperanta). Posle knjizenja obuhvat se prazni, kao stope.
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
Private Const DOP_BLOKOVI  As String = "BLOKOVI"
Private Const DOP_KNJIZENO As String = "KNJIZENO"

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
' Lista BLOKOVI: kvacica i identitet bloka (OtkupID). Lista KNJIZENO: obracun.
Private Const BL_COL_OZN   As Long = 1
Private Const BL_COL_ID    As Long = 12
Private Const KNJ_COL_OBR  As Long = 8
Private Const KNJ_COL_KOOP As Long = 9

Private mLista As String
Private mStope As Collection        ' Dictionary: VrstaVoca, SortaVoca, Klasa, Cena
Private mPlan As Variant            ' poslednji pregled (DoplataPlan) ili Empty
Private mPlanBez As String          ' stope bez otkupa iz poslednjeg pregleda
Private mFill As Boolean            ' punjenje kontrola okida Change
Private mCombosPunjeni As Boolean
Private mStep As String
Private mNacin As Long              ' DOP_NACIN_* (DOP-ODL-15); 0 = nije biran -> nacin 1
Private mIskljuceni As Object       ' Dictionary: OtkupID iskljucenih blokova (DOP-ODL-16)
' OBUHVAT OBRACUNA (DOP-ODL-31): "" = svi kooperanti, inace KooperantID za koga
' se obracun radi. mObuhvatNerazresen: polje nosi tekst koji NIJE izabran
' kooperant -- to nije "svi" nego zastoj (pregled i knjizenje odbijaju).
Private mObuhvat As String
Private mObuhvatNerazresen As Boolean
Private mKoopZastareli As Boolean   ' sifarnik kooperanata se menjao (Scr_ResetCache)

' TEST SEAM (KI-015 obrazac): poslednje pitanje koje bi ekran postavio. Seam
' samo POSMATRA -- odgovor u test rezimu je uvek "Ne". Gasi ga Scr_DopTestReset,
' koji je i u modTest.ResetSeamova (seam nosi stanje).
Private mPitanjeTest As String

' TEST SEAM: unos zone bez zone (forma se u testu ne crta). Postavlja ga samo
' Scr_DopUnosTestSet, tvrdo gejtovan; citalac ponovo pita IsTestMode().
Private mUnosTest As Object

' TEST SEAM: obuhvat postavljen bez polja (Scr_DopObuhvatTestSet). Dok je
' postavljen, stanje se u test rezimu NE cita ponovo iz polja zone.
Private mObuhvatTest As Boolean

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
        DOP_BLOKOVI & "|OTKUI_SEG_DOP_BLOKOVI|OTKUI_GRID_TITLE_DOP_BLOKOVI|64", _
        DOP_KNJIZENO & "|OTKUI_SEG_DOP_KNJIZENO|OTKUI_GRID_TITLE_DOP_KNJIZENO|74", _
        DOP_KO & "|OTKUI_SEG_DOP_KO|OTKUI_GRID_TITLE_DOP_KO|82")
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
        Case DOP_BLOKOVI
            Scr_Radnje = "dopsve:OTKUI_BTN_DOP_UKLJUCI_SVE:110:ghost:0|" & _
                         "dopnista:OTKUI_BTN_DOP_ISKLJUCI_SVE:110:ghost:0|" & _
                         "dopfilt:OTKUI_BTN_DOP_UKLJUCI_FILT:140:ghost:0"
        Case DOP_KNJIZENO
            Scr_Radnje = "dopspec:OTKUI_BTN_DOP_SPEC:170:ghost:1|" & _
                         "dopstornoobr:OTKUI_BTN_DOP_STORNO_OBR:150:danger:1"
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
    ' obuhvat obracuna (DOP-ODL-31): prazno = svi kooperanti
    modOtkupUI.NewFieldG z, "scrDopKoop", Poruka("OTKUI_FLD_DOP_KOOPERANT"), "cmb", "", 1, False, False, "DOP"
    modOtkupUI.NewFieldG z, "scrDopNap", Poruka("OTKUI_FLD_DOP_NAPOMENA"), "txt", "", 1, False, False, "DOP"
    ' --- red B: nacin unosa (DOP-ODL-15) i stopa po kljucu cenovnika (DOP-ODL-02) ---
    modOtkupUI.NewFieldG z, "scrDopNacin", Poruka("OTKUI_FLD_DOP_NACIN"), "cmb", "", 1, False, False, "DOP"
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

    ' RED A: datum, od + sat, do + sat, samo kooperant, napomena. Sat je uzi --
    ' nosi "hh:mm".
    fwS = 64
    fw = (w - PAD * 2 - DOP_GAP * 6 - fwS * 2) / 5
    If fw < 90 Then fw = 90
    X = PAD
    PoljeX z, "scrDopDatum", X, fw, DOP_Y_LBL_A: X = X + fw + DOP_GAP
    PoljeX z, "scrDopOd", X, fw, DOP_Y_LBL_A: X = X + fw + DOP_GAP
    PoljeX z, "scrDopOdSat", X, fwS, DOP_Y_LBL_A: X = X + fwS + DOP_GAP
    PoljeX z, "scrDopDo", X, fw, DOP_Y_LBL_A: X = X + fw + DOP_GAP
    PoljeX z, "scrDopDoSat", X, fwS, DOP_Y_LBL_A: X = X + fwS + DOP_GAP
    PoljeX z, "scrDopKoop", X, fw, DOP_Y_LBL_A: X = X + fw + DOP_GAP
    PoljeX z, "scrDopNap", X, fw, DOP_Y_LBL_A

    ' RED B: nacin, vrsta, sorta, klasa, stopa, dugme "dodaj stopu".
    fw = (w - PAD * 2 - DOP_GAP * 5 - 130) / 5
    If fw < 90 Then fw = 90
    X = PAD
    PoljeX z, "scrDopNacin", X, fw, DOP_Y_LBL_B: X = X + fw + DOP_GAP
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

' Postavi obuhvat obracuna (DOP-ODL-31). Promena ponistava pregled i izbor
' blokova: plan je racunat za drugog kooperanta (ili za sve), pa bi "Knjizi"
' pokazao jedno a upisao drugo. Nerazresen unos nikad ne nosi ID.
Private Sub PostaviObuhvat(ByVal kooperantID As String, ByVal nerazresen As Boolean)
    Dim novi As String
    If Not nerazresen Then novi = Trim$(kooperantID)
    If StrComp(novi, mObuhvat, vbTextCompare) = 0 And nerazresen = mObuhvatNerazresen Then Exit Sub
    mObuhvat = novi
    mObuhvatNerazresen = nerazresen
    Set mIskljuceni = Nothing
    PonistiPregled
End Sub

' Stanje obuhvata prati POLJE kad zona postoji: operater gleda u polje, pa je
' ono izvor istine u trenutku pregleda i knjizenja, ne samo pri promeni.
' Prazno = svi. Tekst bez IZABRANE stavke nije "svi" -- operater je hteo jednog,
' a obracun bi platio sve -- pa je to zastoj (testovi.md par. 5, nerazresen
' izbor). Cita se ListIndex, ne GetComboID: njegov rezervni put vadi ID iz
' teksta, a ovde ID sme da dodje samo iz izabrane stavke.
Private Sub ObuhvatSaPolja()
    Dim c As Object, iD As String
    If IsTestMode() And mObuhvatTest Then Exit Sub
    Set c = Kontrola("scrDopKoop")
    If c Is Nothing Then Exit Sub
    If Len(Trim$(CStr(c.text))) = 0 Then
        PostaviObuhvat "", False
        Exit Sub
    End If
    If c.ListIndex >= 0 Then iD = Trim$(CStr(c.List(c.ListIndex, 1)))
    PostaviObuhvat iD, (Len(iD) = 0)
End Sub

' Polje pokazuje ono sto stanje nosi (dolazak sa kartice, praznjenje posle
' knjizenja). Upis u kontrolu okida Change, pa ide pod mFill.
'
' False kad polje POSTOJI a ne moze da pokaze kooperanta obuhvata. Pozivalac tada
' ne sme da ostavi obuhvat u stanju: pregled cita POLJE (ObuhvatSaPolja), pa bi
' prazno polje tiho pretvorilo "samo ovaj kooperant" u "svi". Kooperant upisan
' posle punjenja liste se zato trazi jos jednom, nad ponovo napunjenom listom.
Private Function ObuhvatUPolje() As Boolean
    Dim c As Object, pre As Boolean, ok As Boolean, muk As Boolean
    On Error Resume Next
    PuniCombos
    Set c = Kontrola("scrDopKoop")
    If c Is Nothing Then
        ObuhvatUPolje = True            ' zone jos nema: stanje je jedini izvor
        Err.Clear
        Exit Function
    End If
    ' Upis nije kucanje: bez garda bi ljuska otvorila panel izbora nad poljem.
    pre = mFill
    mFill = True
    muk = modOtkupUI.PanelIzboraMiruj(True)
    ok = SetComboByID(c, mObuhvat)
    If Not ok Then
        PuniKooperante
        ok = SetComboByID(c, mObuhvat)
    End If
    modOtkupUI.PanelIzboraMiruj muk
    mFill = pre
    ObuhvatUPolje = ok
    ' "Resume Next" prigusuje gresku, ali je ne brise -- ne sme da stigne ljusci.
    Err.Clear
End Function

' Zastoj zbog obuhvata kao TEKST za operatera, "" kad ga nema.
Private Function ObuhvatProblem() As String
    ObuhvatSaPolja
    If mObuhvatNerazresen Then ObuhvatProblem = Poruka("OTKUI_ERR_DOP_KOOP_NERAZRESEN")
End Function

' Ime kooperanta obuhvata za pitanja i poruke, sa ID-em -- dva kooperanta istog
' imena se inace ne razlikuju, a ovo je finansijska odluka.
Private Function ObuhvatNaziv() As String
    Dim ime As String
    ime = Trim$(NzToText(LookupValue(TBL_KOOPERANTI, COL_KOOP_ID, mObuhvat, "Ime")) & " " & _
                NzToText(LookupValue(TBL_KOOPERANTI, COL_KOOP_ID, mObuhvat, "Prezime")))
    ObuhvatNaziv = ComboDisplayWithID(ime, mObuhvat)
End Function

' Nacin unosa za ceo obracun (DOP-ODL-15). Dok operater ne izabere, nacin 1.
Private Function Nacin() As Long
    If mNacin = DOP_NACIN_KONACNA Then
        Nacin = DOP_NACIN_KONACNA
    Else
        Nacin = DOP_NACIN_PO_KG
    End If
End Function

' Promena nacina. Stope su zadate u starom nacinu, pa se nacin menja SAMO dok
' nema unetih stopa (odluka 05.11.2026) -- inace bi 10,80 doplate tiho postalo
' 10,80 konacne cene. "" kad je promenjen (ili je isti), inace razlog.
Private Function PromeniNacinJezgro(ByVal novi As Long) As String
    If novi <> DOP_NACIN_PO_KG And novi <> DOP_NACIN_KONACNA Then
        PromeniNacinJezgro = Poruka("OTKUI_ERR_DOP_NACIN")
        Exit Function
    End If
    If novi = Nacin() Then Exit Function
    If BrojStopa() <> 0 Then
        PromeniNacinJezgro = Poruka("OTKUI_ERR_DOP_NACIN_STOPE")
        Exit Function
    End If
    mNacin = novi
    PonistiPregled
End Function

' ISKLJUCENI blokovi (DOP-ODL-16: podrazumevano su svi ukljuceni). Ostaju i kroz
' nov pregled; prazne ih knjizenje i praznjenje stopa.
Private Function Iskljuceni() As Object
    If mIskljuceni Is Nothing Then
        Set mIskljuceni = CreateObject("Scripting.Dictionary")
        mIskljuceni.CompareMode = vbTextCompare
    End If
    Set Iskljuceni = mIskljuceni
End Function

' Blokovi poslednjeg pregleda sa bar jednom stavkom za knjizenje (plan:
' kolona 11 OtkupID, 16 status -- modDoplata.DoplataPlan).
Private Function BlokSaDoplatom() As Object
    Dim d As Object, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    If IsArray(mPlan) Then
        For i = 1 To UBound(mPlan, 1)
            If CLng(mPlan(i, 16)) = modDoplata.DOP_PLAN_KNJIZI Then d(CStr(mPlan(i, 11))) = True
        Next i
    End If
    Set BlokSaDoplatom = d
End Function

' Izbor za pisca: blokovi pregleda sa doplatom koje operater nije iskljucio.
' Kljuc je OtkupID -- oblik koji primaju KnjiziDoplatu_TX i DoplataPreklapanja.
Private Function Izabrani() As Object
    Dim d As Object, k As Variant
    Set d = BlokSaDoplatom()
    For Each k In Iskljuceni().keys
        If d.Exists(CStr(k)) Then d.Remove CStr(k)
    Next k
    Set Izabrani = d
End Function

' Ukljuci ili iskljuci ceo blok (sve njegove klase). "" kad je prebacen, inace
' razlog. Blok bez doplate nema sta da ukljuci -- pisac ga ionako ne knjizi.
Private Function PrebaciBlok(ByVal otkupID As String) As String
    Dim d As Object
    If Len(otkupID) = 0 Or Not IsArray(mPlan) Then
        PrebaciBlok = Poruka("OTKUI_ERR_NEMA_REDA")
        Exit Function
    End If
    If Not BlokSaDoplatom().Exists(otkupID) Then
        PrebaciBlok = Poruka("OTKUI_ERR_DOP_BLOK_BEZ_DOPLATE")
        Exit Function
    End If
    Set d = Iskljuceni()
    If d.Exists(otkupID) Then
        d.Remove otkupID
    Else
        d(otkupID) = True
    End If
End Function

Private Function PostaviSveBlokove(ByVal ukljuci As Boolean) As Boolean
    Dim d As Object, k As Variant
    If Not IsArray(mPlan) Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_DOP_PRAZAN_PLAN"), True
        Exit Function
    End If
    Set mIskljuceni = Nothing
    If Not ukljuci Then
        Set d = Iskljuceni()
        For Each k In BlokSaDoplatom().keys
            d(CStr(k)) = True
        Next k
    End If
    PostaviSveBlokove = True
End Function

' "UKLJUCI FILTRIRANE" (zahtev operatera 05.11.2026): doplata jednom kooperantu
' kroz pretragu -- Iskljuci sve, pretraga po imenu, Ukljuci filtrirane. Ukljucuje
' SVE blokove koji prolaze trenutnu pretragu, na SVIM stranama: redovi se citaju
' iz celog pogleda mreze (GridBrojRedova / GridCell), a strana je samo crtanje.
' Pravilo pretrage zato postoji jednom (RedoviBlokovi). Ostali blokovi zadrzavaju
' svoje stanje. Poruka kaze koliko blokova filter pokriva, pa zaboravljena
' pretraga ne prolazi neprimecena.
Private Function UkljuciFiltrirane() As Boolean
    Dim d As Object, vidjeni As Object, r As Long, oid As String
    If Not IsArray(mPlan) Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_DOP_PRAZAN_PLAN"), True
        Exit Function
    End If
    If modOtkupUI.GridBrojRedova() = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_DOP_NEMA_FILT"), True
        Exit Function
    End If
    Set d = Iskljuceni()
    Set vidjeni = CreateObject("Scripting.Dictionary")
    vidjeni.CompareMode = vbTextCompare
    For r = 1 To modOtkupUI.GridBrojRedova()
        oid = Trim$(CStr(modOtkupUI.GridCell(r, BL_COL_ID)))
        If Len(oid) > 0 And Not vidjeni.Exists(oid) Then
            vidjeni(oid) = True
            If d.Exists(oid) Then d.Remove oid
        End If
    Next r
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_DOP_UKLJUCENO_FILT") & " " & CStr(vidjeni.count), False
    UkljuciFiltrirane = True
End Function

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
    If Left$(tag, 4) = "row:" Then
        If Scr_Lista() = DOP_BLOKOVI Then ObradiKlik = KlikNaBlok(CLng(val(Mid$(tag, 5))))
        Exit Function
    End If
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

' Klik na red liste Blokovi ukljucuje ili iskljucuje ceo blok -- obrazac liste
' blokova na ekranu Storno (kvacica, identitet u nevidljivoj koloni), samo sto
' ekran ovde pamti ISKLJUCENE. True kad se lista menja (mreza se cita ponovo).
Private Function KlikNaBlok(ByVal red As Long) As Boolean
    Dim greska As String
    If red < 1 Then Exit Function
    greska = PrebaciBlok(Trim$(CStr(modOtkupUI.GridCell(red, BL_COL_ID))))
    If Len(greska) > 0 Then
        modOtkupUI.ShowToast greska, True
        Exit Function
    End If
    KlikNaBlok = True
End Function

' Combo nacina. Odbijena promena vraca combo na stvarni nacin (OsveziZonu).
Private Sub PromeniNacinIzCombo()
    Dim c As Object, greska As String
    Set c = Kontrola("scrDopNacin")
    If c Is Nothing Then Exit Sub
    If c.ListIndex < 0 Then Exit Sub
    greska = PromeniNacinJezgro(c.ListIndex + 1)
    If Len(greska) > 0 Then modOtkupUI.ShowToast greska, True
    OsveziZonu
End Sub

' Tag je ime KONTROLE (nmT), ne polja -- tako ga ljuska i salje.
Private Sub ObradiPromenu(ByVal tag As String)
    If mFill Then Exit Sub
    Select Case tag
        Case "scrDopNacinT"
            PromeniNacinIzCombo
        Case "scrDopVrstaT"
            PuniSorte
        Case "scrDopKoopT"
            ' obuhvat je deo pregleda: promena ga ponistava (v. PostaviObuhvat)
            ObuhvatSaPolja
            OsveziZonu
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
    ' Radnje nad celom listom ne traze red (trebaRed = 0).
    Select Case p(0)
        Case "dopsve", "dopnista"
            If Scr_Lista() = DOP_BLOKOVI Then RadnjaNadRedom = PostaviSveBlokove(p(0) = "dopsve")
            Exit Function
        Case "dopfilt"
            If Scr_Lista() = DOP_BLOKOVI Then RadnjaNadRedom = UkljuciFiltrirane()
            Exit Function
    End Select
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
        Case "dopspec"
            If Scr_Lista() = DOP_KNJIZENO Then _
                StampajSpecifikaciju Trim$(CStr(modOtkupUI.GridCell(red, KNJ_COL_OBR))), _
                                     Trim$(CStr(modOtkupUI.GridCell(red, KNJ_COL_KOOP)))
        Case "dopstorno"
            If Scr_Lista() = DOP_KO Then _
                RadnjaNadRedom = StornirajKO(Trim$(CStr(modOtkupUI.GridCell(red, KO_COL_ID))))
        Case "dopstornoobr"
            If Scr_Lista() = DOP_KO Then
                RadnjaNadRedom = StornirajObracun(Trim$(CStr(modOtkupUI.GridCell(red, KO_COL_OBR))))
            ElseIf Scr_Lista() = DOP_KNJIZENO Then
                RadnjaNadRedom = StornirajObracun(Trim$(CStr(modOtkupUI.GridCell(red, KNJ_COL_OBR))))
            End If
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
    Set mIskljuceni = Nothing
    PonistiPregled
    IsprazniStope = (Scr_Lista() <> DOP_KO And Scr_Lista() <> DOP_KNJIZENO)
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
    greska = ObuhvatProblem()
    If Len(greska) > 0 Then
        PregledJezgro = greska
        Exit Function
    End If
    On Error GoTo EH
    mPlan = modDoplata.DoplataPlan(pOd, pDo, Stope(), mPlanBez, Nacin(), mObuhvat)
    Exit Function
EH:
    PregledJezgro = ChrW(10007) & " " & Err.description
    PonistiPregled
End Function

' "Knjizi": pregled, pitanje sa brojkama, pitanje o ponavljanju, pa jezgro.
Private Function KnjiziKlik() As Boolean
    Dim greska As String, pOd As Date, pDo As Date, dat As Date
    Dim preklop As String, potvrda As Boolean, izbor As Object

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

    Set izbor = Izabrani()
    If izbor.count = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_DOP_NIJEDAN_BLOK"), True
        mLista = DOP_BLOKOVI
        KnjiziKlik = True
        Exit Function
    End If

    If Not Pitaj(PitanjeKnjizenja(izbor)) Then Exit Function

    If UnosPerioda(pOd, pDo, dat, greska) Then
        preklop = modDoplata.DoplataPreklapanja(pOd, pDo, Stope(), Nacin(), izbor, mObuhvat)
        If Len(preklop) > 0 Then
            If Not Pitaj(preklop) Then Exit Function
            potvrda = True
        End If
    End If
    KnjiziKlik = KnjiziJezgro(potvrda)
End Function

' Tekst pitanja pred knjizenje: koliko dokumenata, kg i dinara -- iz istog plana
' i istog izbora koji ide piscu. Dokument je po BLOKU (DOPLATA.md par. 2a): broje
' se UKLJUCENI blokovi sa doplatom (plan: kolona 11 OtkupID, 16 status), a
' iskljuceni se navode posebno.
Private Function PitanjeKnjizenja(ByVal izbor As Object) As String
    Dim i As Long, kg As Double, din As Double, nIskl As Long
    If IsArray(mPlan) Then
        For i = 1 To UBound(mPlan, 1)
            If CLng(mPlan(i, 16)) = modDoplata.DOP_PLAN_KNJIZI Then
                If izbor.Exists(CStr(mPlan(i, 11))) Then
                    kg = kg + CDbl(mPlan(i, 6))
                    din = din + CDbl(mPlan(i, 8))
                End If
            End If
        Next i
    End If
    nIskl = BlokSaDoplatom().count - izbor.count
    ' Obracun za jednog kooperanta se IMENUJE pre brojki (DOP-ODL-31): operater
    ' potvrdjuje kome placa, ne samo koliko.
    If Len(mObuhvat) > 0 Then
        PitanjeKnjizenja = Poruka("OTKUI_LBL_DOP_SAMO_KOOP") & " " & ObuhvatNaziv() & ". "
    End If
    PitanjeKnjizenja = PitanjeKnjizenja & _
                       Poruka("OTKUI_ASK_DOP_KNJIZI") & " " & CStr(izbor.count) & " " & _
                       Poruka("OTKUI_LBL_DOP_DOK") & ", " & FmtKolicina(kg) & " kg, " & _
                       Format$(din, "#,##0.00") & " din"
    If nIskl > 0 Then
        PitanjeKnjizenja = PitanjeKnjizenja & " (" & Poruka("OTKUI_LBL_DOP_ISKLJUCENO") & " " & _
                           CStr(nIskl) & ")"
    End If
    PitanjeKnjizenja = PitanjeKnjizenja & "?"
End Function

' Upis: jedini poziv pisca sa ovog ekrana -- u izabranom nacinu, sa tacno
' ukljucenim blokovima poslednjeg pregleda (pisac ih ponovo proverava i pada po
' imenu ako se nesto promenilo). Uspeh prazni stope, pregled i izbor i prelazi
' na listu Knjizeno -- sledeci obracun je nov obracun.
Private Function KnjiziJezgro(ByVal potvrdaPonavljanja As Boolean) As Boolean
    Dim pOd As Date, pDo As Date, dat As Date, greska As String
    Dim obr As String, en As Long, izbor As Object

    If Not UnosPerioda(pOd, pDo, dat, greska) Then
        modOtkupUI.ShowToast greska, True
        Exit Function
    End If
    ' Obuhvat se cita iz polja i ovde: promena posle pregleda ponistava plan
    ' (PostaviObuhvat), pa ga PregledJezgro racuna ponovo, u obuhvatu koji
    ' operater sada vidi -- i tu staje ako kooperant nije izabran iz liste.
    ObuhvatSaPolja
    If Not IsArray(mPlan) Then
        greska = PregledJezgro()
        If Len(greska) > 0 Then
            modOtkupUI.ShowToast greska, True
            Exit Function
        End If
    End If
    Set izbor = Izabrani()
    If izbor.count = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_DOP_NIJEDAN_BLOK"), True
        Exit Function
    End If
    obr = modDoplata.KnjiziDoplatu_TX(dat, pOd, pDo, Stope(), Unos("scrDopNap"), _
                                      potvrdaPonavljanja, greska, en, Nacin(), _
                                      izbor, mObuhvat)
    If Len(obr) = 0 Then
        modOtkupUI.ShowToast ChrW(10007) & " " & greska, True
        Exit Function
    End If

    Set mStope = New Collection
    Set mIskljuceni = Nothing
    PonistiPregled
    ' Sledeci obracun je nov obracun: obuhvat se prazni kao i stope, pa polje.
    mObuhvatTest = False
    PostaviObuhvat "", False
    ObuhvatUPolje
    mLista = DOP_KNJIZENO
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_DOP_KNJIZENO") & " " & _
                         CStr(modDoplata.AktivneDoplateObracuna(obr).count) & " " & _
                         Poruka("OTKUI_LBL_DOP_DOK"), False
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

' Specifikacija = jedan kooperant u jednom obracunu (DOP-ODL-21) -- red liste
' Knjizeno nosi oba identiteta u nevidljivim kolonama.
Private Sub StampajSpecifikaciju(ByVal obracunID As String, ByVal koopID As String)
    Dim greska As String
    If Len(obracunID) = 0 Or Len(koopID) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_NEMA_REDA"), True
        Exit Sub
    End If
    greska = modPrint.StampajSpecifikacijuDoplate(obracunID, koopID)
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
    If Not Pitaj(Poruka("OTKUI_ASK_DOP_STORNO_OBR") & " " & CStr(n) & " " & _
                 Poruka("OTKUI_LBL_DOP_DOK") & "?") Then Exit Function
    If Not modStorno.StornoObracunDoplate_TX(obracunID) Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_DOP_STORNO") & " " & obracunID, True
        Exit Function
    End If
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_DOP_STORNIRANO") & " " & CStr(n) & " " & _
                         Poruka("OTKUI_LBL_DOP_DOK"), False
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

    mStep = "kooperanti"
    PuniKooperante

    ' Redosled stavki JE nacin: ListIndex + 1 = DOP_NACIN_PO_KG / DOP_NACIN_KONACNA.
    mStep = "nacin"
    Set CB = z.Controls("scrDopNacin").Controls("scrDopNacinT")
    CB.Clear
    CB.AddItem Poruka("OTKUI_DOP_NACIN_1")
    CB.AddItem Poruka("OTKUI_DOP_NACIN_2")

    mCombosPunjeni = True
    mFill = False
    Exit Sub
EH:
    mFill = False
    Debug.Print "modScrDoplata.PuniCombos PAO na koraku [" & mStep & "]: " & _
                Err.Number & " " & Err.description
End Sub

' Lista kooperanata u polju obuhvata (DOP-ODL-31). Prikaz nosi i ID: izbor je
' finansijska odluka, a dva kooperanta istog imena moraju da se razlikuju. Svi
' kooperanti, i neaktivni -- doplata placa robu koja je vec predata. Izabran
' kooperant prezivljava ponovno punjenje, a ukucan a neizabran tekst ostaje u
' polju: ispraznjeno polje bi zastoj ("nije izabran") pretvorilo u "svi".
Private Sub PuniKooperante()
    Dim CB As Object, mapa As Object, k As Variant, pre As Boolean, cur As String
    Dim tekst As String, muk As Boolean, gard As Boolean
    pre = mFill
    On Error GoTo EH
    Set CB = Kontrola("scrDopKoop")
    If CB Is Nothing Then Exit Sub
    mFill = True
    ' Punjenje i vracanje izbora nisu kucanje (v. modOtkupUI.PanelIzboraMiruj).
    muk = modOtkupUI.PanelIzboraMiruj(True)
    gard = True
    If CB.ListIndex >= 0 Then
        cur = Trim$(CStr(CB.List(CB.ListIndex, 1)))
    Else
        tekst = CStr(CB.text)
    End If
    CB.Clear
    CB.ColumnCount = 2
    CB.ColumnWidths = "180 pt;0 pt"
    CB.BoundColumn = 1
    CB.TextColumn = 1
    Set mapa = BuildLookupDict(TBL_KOOPERANTI, COL_KOOP_ID, "Ime", "Prezime")
    For Each k In mapa.keys
        CB.AddItem Trim$(CStr(mapa(k)))
        CB.List(CB.ListCount - 1, 1) = CStr(k)
    Next k
    ShowIDInComboDisplay CB
    If Len(cur) > 0 Then
        SetComboByID CB, cur
    ElseIf Len(tekst) > 0 Then
        CB.text = tekst
    End If
    mKoopZastareli = False
    modOtkupUI.PanelIzboraMiruj muk
    mFill = pre
    Exit Sub
EH:
    If gard Then modOtkupUI.PanelIzboraMiruj muk
    mFill = pre
    Debug.Print "modScrDoplata.PuniKooperante PAO: " & Err.Number & " " & Err.description
    Err.Clear
End Sub

' Ljuska ga zove pri svakom osvezavanju AKTIVNOG ekrana (RefreshFromData), pa i
' pri ulasku na ekran. Sifarnik kooperanata se mogao promeniti dok je operater
' bio drugde, pa se lista u polju obuhvata puni ponovo pri sledecem osvezavanju
' zone -- inace bi nov kooperant bio "neizabran iz liste" do restarta. Isti
' obrazac kao RefillMaticniCombos u ljusci.
Public Sub Scr_ResetCache()
    mKoopZastareli = True
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
    Dim izbor As Object, nSa As Long, nSvih As Long, svi As Object, pre As Boolean
    On Error Resume Next
    Set z = Zona()
    If z Is Nothing Then Exit Sub
    PuniCombos
    If mKoopZastareli Then PuniKooperante

    ' Combo i natpis iznosa prate STVARNI nacin (i posle odbijene promene).
    pre = mFill
    mFill = True
    z.Controls("scrDopNacin").Controls("scrDopNacinT").ListIndex = Nacin() - 1
    mFill = pre
    If Nacin() = DOP_NACIN_KONACNA Then
        z.Controls("scrDopCena").Controls("scrDopCenaL").caption = UCase$(Poruka("OTKUI_FLD_DOP_KONACNA"))
    Else
        z.Controls("scrDopCena").Controls("scrDopCenaL").caption = UCase$(Poruka("OTKUI_FLD_DOP_STOPA"))
    End If

    t = Poruka("OTKUI_LBL_DOP_STOPA") & " " & CStr(BrojStopa())
    If IsArray(mPlan) Then
        ' Zbir je zbir UKLJUCENIH blokova -- isto sto pita i upisuje "Knjizi".
        Set izbor = Izabrani()
        Set svi = CreateObject("Scripting.Dictionary")
        For i = 1 To UBound(mPlan, 1)
            svi(CStr(mPlan(i, 11))) = True
            If CLng(mPlan(i, 16)) = modDoplata.DOP_PLAN_KNJIZI And izbor.Exists(CStr(mPlan(i, 11))) Then
                kg = kg + CDbl(mPlan(i, 6))
                din = din + CDbl(mPlan(i, 8))
            End If
        Next i
        nSa = BlokSaDoplatom().count
        nSvih = svi.count
        t = t & "   " & ChrW(183) & "   " & Poruka("OTKUI_LBL_DOP_PREGLED") & " " & _
            CStr(izbor.count) & " " & Poruka("OTKUI_LBL_DOP_DOK") & ", " & FmtKolicina(kg) & _
            " kg, " & Format$(din, "#,##0.00") & " din"
        If nSa > izbor.count Then
            t = t & " (" & Poruka("OTKUI_LBL_DOP_ISKLJUCENO") & " " & CStr(nSa - izbor.count) & ")"
        End If
        If nSvih > nSa Then
            t = t & "   " & ChrW(183) & "   " & Poruka("OTKUI_DOP_ST_BEZ") & ": " & CStr(nSvih - nSa)
        End If
    Else
        t = t & "   " & ChrW(183) & "   " & Poruka("OTKUI_LBL_DOP_BEZ_PREGLEDA")
    End If
    If Len(mPlanBez) > 0 Then
        t = t & "   " & ChrW(183) & "   " & Poruka("OTKUI_ERR_DOP_STOPA_BEZ_OTKUPA") & " " & mPlanBez
    End If
    If mObuhvatNerazresen Then
        t = t & "   " & ChrW(183) & "   " & Poruka("OTKUI_ERR_DOP_KOOP_NERAZRESEN")
    End If
    z.Controls("dopHint").caption = t
    z.Controls("dopHint").ForeColor = IIf(Len(mPlanBez) > 0 Or mObuhvatNerazresen, C_RUST, C_MUTED)
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
        Case DOP_PLAN:     Scr_Rows = RedoviPlan(q)
        Case DOP_BLOKOVI:  Scr_Rows = RedoviBlokovi(q)
        Case DOP_KNJIZENO: Scr_Rows = RedoviKnjizeno(q)
        Case DOP_KO:       Scr_Rows = RedoviKO(q)
        Case Else:         Scr_Rows = RedoviStope(q)
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

'-------------------------------------------------------- LISTA: PREGLED
' Zbirno po kooperantu, kulturi i klasi (DOP-ODL-18), SAMO ukljuceni blokovi sa
' doplatom -- tacno ono sto ce "Knjizi" upisati. Detalj po bloku je lista Blokovi.
Private Function PlanKolone() As Variant
    PlanKolone = Array( _
        "OTKUI_HD_KOOP||txt|0|1", _
        "OTKUI_HD_VRSTA||txt|110|1", _
        "OTKUI_HD_SORTA||txt|100|2", _
        "OTKUI_HD_KLASA||txt|50|1", _
        "OTKUI_HD_BLOKOVA||num|64|3", _
        "OTKUI_HD_KG||kg|84|1", _
        "OTKUI_HD_IZNOS||rsd|96|1")
End Function

Private Function RedoviPlan(ByVal q As String) As Variant
    Dim outA() As Variant, n As Long, i As Long, hay As String, k As String
    Dim imena As Object, grupe As Object, izbor As Object, blokovi As Object
    Dim a As Variant, kk As Variant, ime As String, sumKg As Double, sumIzn As Double
    If Not IsArray(mPlan) Then
        RedoviPlan = PrazanRezultat(PlanKolone())
        Exit Function
    End If
    ' grupe: "koop|VRSTA|SORTA|KLASA" -> Array(koop, vrsta, sorta, klasa, blokova, kg, din)
    Set izbor = Izabrani()
    Set grupe = CreateObject("Scripting.Dictionary")
    grupe.CompareMode = vbTextCompare
    Set blokovi = CreateObject("Scripting.Dictionary")
    blokovi.CompareMode = vbTextCompare
    For i = 1 To UBound(mPlan, 1)
        If CLng(mPlan(i, 16)) <> modDoplata.DOP_PLAN_KNJIZI Then GoTo SledeciRed
        If Not izbor.Exists(CStr(mPlan(i, 11))) Then GoTo SledeciRed
        k = CStr(mPlan(i, 1)) & "|" & UCase$(CStr(mPlan(i, 3))) & "|" & _
            UCase$(CStr(mPlan(i, 4))) & "|" & UCase$(CStr(mPlan(i, 5)))
        If Not grupe.Exists(k) Then
            grupe.Add k, Array(mPlan(i, 1), mPlan(i, 3), mPlan(i, 4), mPlan(i, 5), 0&, 0#, 0#)
        End If
        a = grupe(k)
        If Not blokovi.Exists(k & "|" & CStr(mPlan(i, 11))) Then
            blokovi.Add k & "|" & CStr(mPlan(i, 11)), True
            a(4) = CLng(a(4)) + 1
        End If
        a(5) = CDbl(a(5)) + CDbl(mPlan(i, 6))
        a(6) = CDbl(a(6)) + CDbl(mPlan(i, 8))
        grupe(k) = a
SledeciRed:
    Next i
    If grupe.count = 0 Then
        RedoviPlan = PrazanRezultat(PlanKolone())
        Exit Function
    End If

    Set imena = BuildLookupDict(TBL_KOOPERANTI, COL_KOOP_ID, "Ime", "Prezime")
    ReDim outA(1 To grupe.count, 1 To 7)
    For Each kk In grupe.keys
        a = grupe(kk)
        ime = CStr(a(0))
        If imena.Exists(ime) Then ime = Trim$(CStr(imena(ime)))
        hay = ime & "|" & CStr(a(1)) & "|" & CStr(a(2))
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo SledecaGrupa
        End If
        n = n + 1
        outA(n, 1) = ime
        outA(n, 2) = CStr(a(1))
        outA(n, 3) = CStr(a(2))
        outA(n, 4) = CStr(a(3))
        outA(n, 5) = CLng(a(4))
        outA(n, 6) = CDbl(a(5))
        outA(n, 7) = modNovac.ZaokruziNovac(CDbl(a(6)))
        sumKg = sumKg + CDbl(a(5))
        sumIzn = sumIzn + CDbl(outA(n, 7))
SledecaGrupa:
    Next kk
    RedoviPlan = Array(PlanKolone(), outA, n, sumKg, sumIzn, Array(0, 0, 0))
End Function

'-------------------------------------------------------- LISTA: BLOKOVI
' Jedan red po bloku i klasi poslednjeg pregleda. Kvacica pripada BLOKU (sve
' njegove klase): klik na red ga ukljucuje ili iskljucuje (KlikNaBlok). Status
' kaze zasto stavka nije u knjizenju, ili da blok vec ima doplatu (DOP-ODL-06).
' Zbir u podnozju je zbir UKLJUCENIH -- isti kao Pregled i pitanje pred upisom.
Private Function BlokoviKolone() As Variant
    BlokoviKolone = Array( _
        "OTKUI_HD_OZN||txt|36|1", _
        "OTKUI_HD_BROJ||txt|110|1", _
        "OTKUI_HD_DATUM||txt|80|2", _
        "OTKUI_HD_KOOP||txt|0|1", _
        "OTKUI_HD_OM||txt|120|2", _
        "OTKUI_HD_KLASA||txt|50|1", _
        "OTKUI_HD_KG||kg|84|1", _
        "OTKUI_HD_CENA_OTKUPA||rsd|90|3", _
        "OTKUI_HD_STOPA||rsd|80|1", _
        "OTKUI_HD_IZNOS||rsd|96|1", _
        "OTKUI_HD_STATUS||txt|120|2", _
        "OTKUI_HDI_REF||txt|1|4")
End Function

Private Function RedoviBlokovi(ByVal q As String) As Variant
    Dim outA() As Variant, n As Long, i As Long, hay As String, oid As String
    Dim imena As Object, stanice As Object, izbor As Object, ime As String, sta As String
    Dim sumKg As Double, sumIzn As Double, knjizi As Boolean, ukljucen As Boolean
    If Not IsArray(mPlan) Then
        RedoviBlokovi = PrazanRezultat(BlokoviKolone())
        Exit Function
    End If
    Set izbor = Izabrani()
    Set imena = BuildLookupDict(TBL_KOOPERANTI, COL_KOOP_ID, "Ime", "Prezime")
    Set stanice = BuildLookupDict(TBL_STANICE, COL_STA_ID, "Naziv")
    ReDim outA(1 To UBound(mPlan, 1), 1 To BL_COL_ID)
    For i = 1 To UBound(mPlan, 1)
        oid = CStr(mPlan(i, 11))
        ime = CStr(mPlan(i, 1))
        If imena.Exists(ime) Then ime = Trim$(CStr(imena(ime)))
        sta = CStr(mPlan(i, 2))
        If stanice.Exists(sta) Then sta = Trim$(CStr(stanice(sta)))
        hay = CStr(mPlan(i, 12)) & "|" & ime & "|" & sta
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeci
        End If
        knjizi = (CLng(mPlan(i, 16)) = modDoplata.DOP_PLAN_KNJIZI)
        ukljucen = (knjizi And izbor.Exists(oid))
        n = n + 1
        If ukljucen Then outA(n, BL_COL_OZN) = ChrW(10003) Else outA(n, BL_COL_OZN) = ""
        outA(n, 2) = CStr(mPlan(i, 12))
        If IsDate(mPlan(i, 17)) Then outA(n, 3) = Format$(CDate(mPlan(i, 17)), "dd.mm.yyyy")
        outA(n, 4) = ime
        outA(n, 5) = sta
        outA(n, 6) = CStr(mPlan(i, 5))
        outA(n, 7) = CDbl(mPlan(i, 6))
        outA(n, 8) = CDbl(mPlan(i, 13))
        outA(n, 9) = CDbl(mPlan(i, 7))
        outA(n, 10) = CDbl(mPlan(i, 8))
        outA(n, 11) = StatusBloka(knjizi, izbor.Exists(oid), CDbl(mPlan(i, 15)))
        outA(n, BL_COL_ID) = oid
        If ukljucen Then
            sumKg = sumKg + CDbl(mPlan(i, 6))
            sumIzn = sumIzn + CDbl(mPlan(i, 8))
        End If
Sledeci:
    Next i
    RedoviBlokovi = Array(BlokoviKolone(), outA, n, sumKg, sumIzn, Array(0, 0, 0))
End Function

' Status stavke na listi Blokovi: bez doplate (nacin 2, cena otkupa na ili
' iznad konacne), iskljucen, ili blok koji vec ima doplatu za tu klasu.
Private Function StatusBloka(ByVal knjizi As Boolean, ByVal izabran As Boolean, _
                             ByVal vecPoKg As Double) As String
    If Not knjizi Then
        StatusBloka = Poruka("OTKUI_DOP_ST_BEZ")
    ElseIf Not izabran Then
        StatusBloka = Poruka("OTKUI_DOP_ST_ISKLJUCEN")
    ElseIf vecPoKg > 0.000001 Then
        StatusBloka = Poruka("OTKUI_DOP_ST_PONOVLJENA")
    End If
End Function

'------------------------------------------------------- LISTA: KNJIZENO
' Zbirno: jedan red po obracunu i kooperantu (DOP-ODL-18; to je i obim buduce
' specifikacije, DOP-ODL-21) -- iz istog citaca koji puni karticu i saldo.
' Detalj po dokumentu (bloku) je lista Dokumenti.
Private Function KnjizenoKolone() As Variant
    KnjizenoKolone = Array( _
        "OTKUI_HD_DATUM||txt|80|1", _
        "OTKUI_HD_KOOP||txt|0|1", _
        "OTKUI_HD_SPEC||txt|130|1", _
        "OTKUI_HD_PERIOD||txt|190|2", _
        "OTKUI_HD_DOKUMENATA||num|64|3", _
        "OTKUI_HD_KG||kg|84|1", _
        "OTKUI_HD_IZNOS||rsd|100|1", _
        "OTKUI_HD_OBRACUN||txt|1|4", _
        "OTKUI_HDI_REF||txt|1|4")
End Function

Private Function RedoviKnjizeno(ByVal q As String) As Variant
    Dim r As Variant, i As Long, g As Object, dok As Object, k As String, a As Variant
    Dim outA() As Variant, n As Long, kk As Variant, imena As Object, ime As String
    Dim hay As String, sumKg As Double, sumIzn As Double

    r = modDoplata.DoplataStavkeRedovi()
    If Not IsArray(r) Then
        RedoviKnjizeno = PrazanRezultat(KnjizenoKolone())
        Exit Function
    End If

    ' g: "ObracunID|KooperantID" -> Array(datum, koop, periodOd, periodDo, dok, kg, din,
    '    obr, broj specifikacije) -- par je specifikacija (DOP-ODL-21)
    Set g = CreateObject("Scripting.Dictionary")
    g.CompareMode = vbTextCompare
    Set dok = CreateObject("Scripting.Dictionary")
    dok.CompareMode = vbTextCompare
    For i = 1 To UBound(r, 1)
        k = CStr(r(i, 5)) & "|" & CStr(r(i, 4))
        If Not g.Exists(k) Then
            g.Add k, Array(r(i, 3), r(i, 4), r(i, 14), r(i, 15), 0&, 0#, 0#, CStr(r(i, 5)), _
                           CStr(r(i, 20)))
        End If
        a = g(k)
        If Not dok.Exists(CStr(r(i, 1))) Then
            dok.Add CStr(r(i, 1)), True
            a(4) = CLng(a(4)) + 1
        End If
        a(5) = CDbl(a(5)) + CDbl(r(i, 10))
        a(6) = CDbl(a(6)) + CDbl(r(i, 12))
        g(k) = a
    Next i

    Set imena = BuildLookupDict(TBL_KOOPERANTI, COL_KOOP_ID, "Ime", "Prezime")
    ReDim outA(1 To g.count, 1 To KNJ_COL_KOOP)
    For Each kk In g.keys
        a = g(kk)
        ime = CStr(a(1))
        If imena.Exists(ime) Then ime = Trim$(CStr(imena(ime)))
        hay = ime & "|" & CStr(a(7)) & "|" & CStr(a(8))
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeci
        End If
        n = n + 1
        outA(n, 1) = Format$(CDate(a(0)), "dd.mm.yyyy")
        outA(n, 2) = ime
        outA(n, 3) = CStr(a(8))
        outA(n, 4) = Format$(CDate(a(2)), "dd.mm.yyyy hh:nn") & " " & ChrW(8211) & " " & _
                     Format$(CDate(a(3)), "dd.mm.yyyy hh:nn")
        outA(n, 5) = CLng(a(4))
        outA(n, 6) = CDbl(a(5))
        outA(n, 7) = modNovac.ZaokruziNovac(CDbl(a(6)))
        outA(n, KNJ_COL_OBR) = CStr(a(7))
        outA(n, KNJ_COL_KOOP) = CStr(a(1))
        sumKg = sumKg + CDbl(a(5))
        sumIzn = sumIzn + CDbl(outA(n, 7))
Sledeci:
    Next kk
    RedoviKnjizeno = Array(KnjizenoKolone(), outA, n, sumKg, sumIzn, Array(0, 0, 0))
End Function

'------------------------------------------------------ LISTA: DOKUMENTI
' Kljuc liste je istorijski "KO"; prikaz je "Dokumenti" -- jedan red po
' dokumentu doplate (bloku), za stampu i storno jednog dokumenta.
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
    mNacin = 0
    Set mIskljuceni = Nothing
    mObuhvatTest = False
    mObuhvat = ""
    mObuhvatNerazresen = False
    ObuhvatUPolje
End Sub

' ULAZ SA KARTICE KOOPERANTA (modScrIzvestaji, DOP-ODL-31): obracun samo za tog
' kooperanta. Zove se POSLE prelaska na ekran (modOtkupUI.IdiNaEkran) -- zona
' tada postoji, pa polje pokazuje isto sto stanje nosi. "" kad je postavljen,
' inace razlog. NIJE test seam: radi i u produkciji. Kooperant mora da postoji
' tacno jednom -- ID stize iz izvestaja, ali ulaz ne veruje pozivaocu. Unete
' stope ostaju: one su kljuc i vrednost, ne zavise od kooperanta.
Public Function Scr_DopOtvoriZaKooperanta(ByVal kooperantID As String) As String
    Dim iD As String
    On Error GoTo EH
    iD = Trim$(kooperantID)
    If Len(iD) = 0 Then
        Scr_DopOtvoriZaKooperanta = Poruka("OTKUI_ERR_DOP_KOOP_NEMA")
        Exit Function
    End If
    RequireTacnoJedan TBL_KOOPERANTI, COL_KOOP_ID, iD, "KooperantID", _
                      "modScrDoplata.Scr_DopOtvoriZaKooperanta"
    mObuhvatTest = False
    PostaviObuhvat iD, False
    If Not ObuhvatUPolje() Then
        PostaviObuhvat "", False
        ObuhvatUPolje
        Scr_DopOtvoriZaKooperanta = Poruka("OTKUI_ERR_DOP_KOOP_NEMA")
        Exit Function
    End If
    mLista = DOP_STOPE
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_DOP_OTVORENO_ZA") & " " & ObuhvatNaziv(), False
    Exit Function
EH:
    Scr_DopOtvoriZaKooperanta = Poruka("OTKUI_ERR_DOP_KOOP_NEMA") & " " & Err.description
    Err.Clear
End Function

' Obuhvat koji ekran trenutno nosi ("" = svi kooperanti). Samo cita.
Public Function Scr_DopObuhvat() As String
    Scr_DopObuhvat = mObuhvat
End Function

' TEST SEAM: obuhvat bez polja (zona se u testu ne crta uvek). nerazresen = True
' glumi tekst ukucan u polje bez izbora iz liste. Tvrdo gejtovan; dok je
' postavljen, stanje se u test rezimu ne cita ponovo iz polja.
Public Sub Scr_DopObuhvatTestSet(ByVal kooperantID As String, _
                                 Optional ByVal nerazresen As Boolean = False)
    If Not IsTestMode() Then Exit Sub
    mObuhvatTest = True
    PostaviObuhvat kooperantID, nerazresen
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

' Nacin unosa istim putem kao combo (ista kapija); "" kad je promenjen.
Public Function Scr_DopNacinTestSet(ByVal novi As Long) As String
    If Not IsTestMode() Then Exit Function
    Scr_DopNacinTestSet = PromeniNacinJezgro(novi)
End Function

Public Function Scr_DopNacin() As Long
    Scr_DopNacin = Nacin()
End Function

' Klik na blok liste Blokovi, po OtkupID-u (mreza se u testu ne crta) -- isto
' jezgro kao klik na red. "" kad je prebacen, inace razlog.
Public Function Scr_DopBlokTestPrebaci(ByVal otkupID As String) As String
    If Not IsTestMode() Then Exit Function
    Scr_DopBlokTestPrebaci = PrebaciBlok(otkupID)
End Function

' "Ukljuci sve" / "Iskljuci sve" -- ista radnja kao dugme liste Blokovi.
Public Function Scr_DopBlokoviTestSvi(ByVal ukljuci As Boolean) As Boolean
    If Not IsTestMode() Then Exit Function
    Scr_DopBlokoviTestSvi = PostaviSveBlokove(ukljuci)
End Function

' "Ukljuci filtrirane" -- ista radnja kao dugme; cita mrezu koju je test napunio
' (modOtkupUI.GridTestLoad sa pretragom).
Public Function Scr_DopBlokoviTestFiltrirani() As Boolean
    If Not IsTestMode() Then Exit Function
    Scr_DopBlokoviTestFiltrirani = UkljuciFiltrirane()
End Function
