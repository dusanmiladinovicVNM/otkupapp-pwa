Attribute VB_Name = "modScrAgro"
'=====================================================================
' modScrAgro - ekran "Agrohemija" (v6-ui-171). Faza E, stavka 15.
'
' Ljuska ga ne poznaje po imenu: dobija ga preko Application.Run, da klijent
' kome ovaj modul nedostaje i dalje radi (zamka #19).
'
' ODAKLE DOLAZI: frmAgrohemija drzi dve sekcije jednu pored druge (izlaz i
' ulaz), svaku sa svojom korpom, svojim brojem dokumenta i svojom listom. U
' ljusci postoji JEDNA mreza, pa su dve sekcije postale PREKIDAC REZIMA u
' zoni: IZDAVANJE i PRIJEM. Polja se dele gde znace isto (artikal, kolicina,
' broj dokumenta), a razlikuju gde ne (kooperant + parcele vs dobavljac +
' cena). Obe korpe zive istovremeno -- prelazak izmedju rezima ne prazni nista.
'
' STA JE OVDE, A STA NIJE: ovde je REDOSLED i PRIKAZ. Nijedno poslovno
' pravilo, nijedna kapija i nijedan upis nisu ovde:
'   - korpa, provere i transakcija    -> modAgroUnos
'   - stanje, promet, dug, smart doza -> modAgrohemija
'   - pocetni dug (migracija)         -> modAgrohemija.BookPocetniDug
'   - odbitak duga iz tblNovac        -> modNovac.GetAgroAbzugMapa
'
' CETIRI LISTE u deljenoj mrezi (prekidac iznad nje):
'   KORPA    stavke koje cekaju upis, za AKTIVAN rezim; radnja: ukloni red
'   STANJE   stanje magacina po artiklu (GetMagacinStanje)
'   PROMET   ceo ledger, ulazi i izlazi (GetMagacinPrometForGrid)
'   DUGOVI   dug po kooperantu (GetAgroDugoviForGrid)
'
' Legacy nema nijednu od te cetiri liste -- imao je samo dve korpe i broj duga
' uz izabranog kooperanta. Mreza je ovde besplatna: ljuska je vec nosi.
'
' POLJA SU LJUSKINA, NE EKRANOVA. Sklop "natpis + shell + kontrola" pravi
' modOtkupUI.NewFieldG, isti onaj koji crta i unosnu formu; raspored unutar
' polja radi modOtkupUI.LayoutFieldInner. Ekran bira SAMO gde polje stoji i
' koliko je siroko. Oboje je otvoreno za ekrane u v6-ui-159 (unos prerade na
' Paletama, Faza C/10) -- ovaj ekran je drugi korisnik, ne prvi, i namerno ne
' pravi svoju fabriku: dve fabrike polja znace dva izgleda istog polja.
'
' Kombo u zoni zato MORA da bude polje (okvir 'nm' + kontrola 'nmT'), a ne gola
' kontrola: panel za izbor (modOtkupUI.FindCombo) trazi bas taj oblik.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const SCRAG_BUILD As String = "v6-ui-171"

' Visina zone. Veca je od KPI_H jer zona nosi ceo unos: prekidac rezima, dva
' reda polja, dva reda objasnjenja i red dugmadi.
Private Const AG_ZONA_H  As Single = 202

' Razmak izmedju polja u redu. Visinu polja i unutrasnji padding drzi ljuska
' (FIELD_GRP_H / FIELD_H / INPUT_PAD u modOtkupUI) -- ekran ih ne ponavlja.
Private Const AG_GAP     As Single = 10
' Visina polja koje pravi NewFieldG: natpis (0..12) + kutija (16..44) + vazduh.
' Ekran je zna samo da bi znao gde pocinje SLEDECI red.
Private Const AG_FLD_H   As Single = 28

' Redovi zone (Y koordinate)
Private Const AG_Y_SEG   As Single = 18
Private Const AG_Y_LBL_A As Single = 46
Private Const AG_Y_FLD_A As Single = 62
Private Const AG_Y_LBL_B As Single = 96
Private Const AG_Y_HINT  As Single = 144
Private Const AG_Y_HINT2 As Single = 157
Private Const AG_Y_BTN   As Single = 172
Private Const AG_BTN_H   As Single = 24
Private Const AG_KPI_W   As Single = 140

' Rezimi. Nisu F-tasteri: ekran ih nema, ovo je prekidac unutar zone.
Private Const AG_IZLAZ As String = "IZLAZ"
Private Const AG_ULAZ  As String = "ULAZ"

' Kljucevi lista
Private Const AG_KORPA  As String = "KORPA"
Private Const AG_STANJE As String = "STANJE"
Private Const AG_PROMET As String = "PROMET"
Private Const AG_DUGOVI As String = "DUGOVI"

Private mMod As String              ' AG_IZLAZ | AG_ULAZ
Private mLista As String            ' AG_KORPA | AG_STANJE | AG_PROMET | AG_DUGOVI
Private mKorpaI As Collection       ' korpa izdavanja
Private mKorpaU As Collection       ' korpa prijema

' Izabrane parcele. Legacy sabira ha SVIH oznacenih parcela iz liste; mreza
' ljuske bira jedan red, a combo jednu stavku, pa se vise parcela skuplja
' dugmetom "+ Parcela". Zbir ha je ono sto smart doza racuna, pa se drzi uz
' spisak i ne racuna ponovo.
Private mParIds As Collection       ' izabrani ParcelaID
Private mParHa As Double            ' zbir ha izabranih parcela
Private mParMapa As Object          ' ParcelaID -> ha (za kooperanta u comboju)

' Punjenje lista pise u combo i time okida Change -- isto kao mPopMute u
' ljusci. Bez ovoga bi punjenje parcela pozvalo obradu promene artikla.
Private mFill As Boolean
Private mCombosPunjeni As Boolean

' Artikal -> AgroArtikalInfo. Vrednost se preracunava na SVAKI otkucaj, a
' info je cetiri citanja tblArtikli; bez kesa je to citanje po znaku.
Private mArtKes As Object

' Kooperant -> dug. OsveziZonu se zove pri SVAKOM citanju mreze i pri svakom
' otkucaju, a dug je dva puna prolaza kroz tblMagacin i tblNovac. Bez kesa se
' to na pravim podacima vidi isto kao "efekat storna po modu" u modScrStorno:
' polje kuca sporije nego sto operater pise. Kes cisti Scr_ResetCache, a
' ljuska ga zove posle svakog upisa (RefreshFromData).
Private mDugKes As Object

' Prikaz -> identitet, za dvoklik. Mreza je SORTIRANA, pa redni broj reda nije
' redni broj u izvoru; jedino sto red pouzdano nosi je ono sto se u njemu vidi.
' KOLIZIJA SE PAMTI KAO PRAZNO: dva kooperanta istog imena (ili dva artikla
' istog naziva) daju dvosmislen prikaz, i tada dvoklik ODBIJA da bira umesto da
' pogodi -- isto pravilo kao "dvosmislen broj -> MANUAL" u storno okviru.
Private mDugIds As Object           ' "Ime Prezime" -> KooperantID ("" = dvosmisleno)
Private mArtIds As Object           ' naziv artikla  -> ArtikalID   ("" = dvosmisleno)

Private mStep As String             ' korak za poruku o gresci

'--------------------------------------------------------- UGOVOR EKRANA
Public Function Scr_Meta() As String
    Scr_Meta = "kljuc=AGRO|naslov=OTKUI_NAV_AGRO|sub=OTKUI_SCRAG_SUB" & _
               "|lista=OTKUI_SCRAG_LISTA|oblik=zona+mreza|upis=zona"
End Function

' Prekidac bira LISTU u mrezi, ne rezim unosa -- rezim je u zoni, jer se unos
' i pregled ne prekidaju istim prekidacem. Korpa je prva: to je jedino sto
' operater u tom trenutku menja.
Public Function Scr_Liste() As Variant
    Scr_Liste = Array( _
        AG_KORPA & "|OTKUI_SEG_AG_KORPA|OTKUI_GRID_TITLE_AG_KORPA|64", _
        AG_STANJE & "|OTKUI_SEG_AG_STANJE|OTKUI_GRID_TITLE_AG_STANJE|58", _
        AG_PROMET & "|OTKUI_SEG_AG_PROMET|OTKUI_GRID_TITLE_AG_PROMET|58", _
        AG_DUGOVI & "|OTKUI_SEG_AG_DUGOVI|OTKUI_GRID_TITLE_AG_DUGOVI|58")
End Function

Public Function Scr_Lista() As String
    If Len(mLista) = 0 Then mLista = AG_KORPA
    Scr_Lista = mLista
End Function

' U listi korpe naslov nosi rezim -- inace se dve korpe ne razlikuju.
Public Function Scr_NaslovDopuna() As String
    If Scr_Lista() = AG_KORPA Then
        If Rezim() = AG_ULAZ Then
            Scr_NaslovDopuna = ChrW(8212) & " " & Poruka("OTKUI_SEG_AG_PRIJEM")
        Else
            Scr_NaslovDopuna = ChrW(8212) & " " & Poruka("OTKUI_SEG_AG_IZDAVANJE")
        End If
    End If
End Function

' Cipovi AKTIVNE liste. Korpa ih nema: to je nekoliko redova koje je operater
' upravo uneo, tu se ne trazi nego se gleda. Ostale tri liste su pregledi preko
' cele istorije, pa im suzavanje treba.
'
' Prvi cip je svuda "sve" -- ljuska na njega pada kad zatecen filter ne pripada
' listi na koju se upravo preslo (RefreshChipsForScreen).
Public Function Scr_Cipovi() As String
    Select Case Scr_Lista()
        Case AG_STANJE
            Scr_Cipovi = "sve:OTKUI_CHIP_SVE:40|" & _
                         "ima:OTKUI_CIPA_IMA:78|" & _
                         "nema:OTKUI_CIPA_NEMA:88"
        Case AG_PROMET
            Scr_Cipovi = "sve:OTKUI_CHIP_SVE:40|" & _
                         "ulaz:OTKUI_CIPA_ULAZ:52|" & _
                         "izlaz:OTKUI_CIPA_IZLAZ:52|" & _
                         "godina:OTKUI_CIPA_GODINA:84"
        Case AG_DUGOVI
            Scr_Cipovi = "sve:OTKUI_CHIP_SVE:40|" & _
                         "duguju:OTKUI_CIPA_DUGUJU:78"
    End Select
End Function

' PRAVILA CIPOVA, odvojena od mreze da bi mogla da se izmere bez nje. Kljuc je
' EKRANOV -- ljuska ga je samo vratila onakvog kakvog ga je dobila iz Scr_Cipovi.
' Nepoznat i prazan kljuc PUSTAJU sve: ekran koji dobije filter koji ne poznaje
' pokazuje punu listu, ne praznu.
Public Function AgCipStanje(ByVal filter As String, ByVal stanje As Double) As Boolean
    Select Case filter
        Case "ima":  AgCipStanje = (stanje > 0)
        Case "nema": AgCipStanje = (stanje <= 0)
        Case Else:   AgCipStanje = True
    End Select
End Function

Public Function AgCipPromet(ByVal filter As String, ByVal tip As String, _
                            ByVal datum As Double) As Boolean
    Select Case filter
        Case "ulaz":   AgCipPromet = (StrComp(Trim$(tip), MAG_ULAZ, vbTextCompare) = 0)
        Case "izlaz":  AgCipPromet = (StrComp(Trim$(tip), MAG_IZLAZ, vbTextCompare) = 0)
        Case "godina"
            ' Red bez citljivog datuma NE prolazi kroz cip godine: propustiti ga
            ' znacilo bi tvrditi da je iz tekuce godine, a ne zna se.
            If datum <= 0 Then Exit Function
            AgCipPromet = (Year(CDate(datum)) = Year(Date))
        Case Else:     AgCipPromet = True
    End Select
End Function

Public Function AgCipDugovi(ByVal filter As String, ByVal dug As Double) As Boolean
    Select Case filter
        Case "duguju": AgCipDugovi = (dug > 0)
        Case Else:     AgCipDugovi = True
    End Select
End Function

' Koliko stavki CEKA operatera. Na ovom ekranu to je korpa: redovi koje je uneo
' a nije proknjizio. Jedino su one prolazne -- sve ostalo na ekranu je vec u
' tabelama. Bez ove brojke operater koji predje na drugi ekran nema nijedan
' znak da mu je korpa ostala puna.
Public Function Scr_Brojac() As Long
    Scr_Brojac = BrojUKorpi(mKorpaI) + BrojUKorpi(mKorpaU)
End Function

' Radnja nad redom postoji samo u korpi: red koji jos nije upisan sme da se
' ukloni. Stanje, promet i dugovi su pregledi -- nad njima se ne radi nista,
' storno stavke magacina je posao ekrana Storno.
Public Function Scr_Radnje() As String
    If Scr_Lista() = AG_KORPA Then _
        Scr_Radnje = "agdel:OTKUI_BTN_AG_UKLONI:110:danger:1"
End Function

'=====================================================================
' ZONA
'=====================================================================
Public Sub Scr_Build(ByVal z As Object)
    Dim i As Long

    ' Bela podloga ispod oba reda polja. Zona je krem, a polja su bela -- bez
    ' podloge se izmedju njih vidi pozadina zone i panel izgleda kao niz
    ' odvojenih ostrva. Ista popravka i isti razlog kao "preBg" na Paletama.
    '
    ' MORA da bude LABELA, ne Frame: Frame je u MSForms prozorska kontrola i
    ' crta se IZNAD bezprozorskih bez obzira na z-order, pa bi kao Frame
    ' pokrila natpise i objasnjenja. Napravljena PRVA, labela ostaje ispod
    ' svega -- a polja su Frame-ovi, pa se probijaju iznad nje kako i treba.
    modUiKit.NewLbl z, "agBg", "", 0, 0, 100, 10, 8, False, 0, C_WHITE

    modUiKit.NewLbl z, "agCap", UCase$(Poruka("OTKUI_SCRAG_CAP")), PAD, 6, 260, 11, _
                    TS_MICRO, True, C_MUTED, -1

    ' PREKIDAC REZIMA je segmentni prekidac, isti kao onaj nad mrezom -- pa se i
    ' pravi istom fabrikom (NewSegBtn, vrsta "seg"), ne kao obicno dugme.
    ' Vrsta nije kozmetika: clsFlatBtn.IsSelected priznaje izabrano stanje
    ' (Font.Bold) samo za "nav", "chip" i "seg". Kao "btn" je izabran rezim bio
    ' obicno dugme kome hover-out vrati zapamcenu belu -- v. OsveziPrekidacRezima.
    modUiKit.NewSegBtn z, "scrAgSegI", Poruka("OTKUI_SEG_AG_IZDAVANJE"), _
                       PAD, AG_Y_SEG, 112, 22, True
    modUiKit.NewSegBtn z, "scrAgSegU", Poruka("OTKUI_SEG_AG_PRIJEM"), _
                       PAD + 116, AG_Y_SEG, 112, 22, False

    ' cetiri brojke desno -- iste one koje legacy drzi u KPI traci iznad forme
    For i = 0 To 3
        modUiKit.NewLbl z, "agKL" & i, "", 0, 6, AG_KPI_W, 11, TS_MICRO, True, C_MUTED, -1
        modUiKit.NewLbl z, "agKV" & i, ChrW(8212), 0, AG_Y_SEG, AG_KPI_W, 20, _
                        TS_KPI, True, C_FOREST, -1, fmTextAlignLeft, F_NUM
    Next i

    ' POLJA. Pravi ih ljuska (NewFieldG), ekran im samo kaze gde stoje -- v.
    ' zaglavlje. Prefiks "scr" je OBAVEZAN: bez njega promena teksta ide ljusci,
    ' koja o ovim poljima ne zna nista. Grupa "AG" je samo oznaka pripadnosti;
    ' raspored radi Scr_Layout, ne LayoutFields.
    ' --- red A ---
    modOtkupUI.NewFieldG z, "scrAgKoop", Poruka("OTKUI_FLD_AG_KOOPERANT"), "cmb", "", 1, False, False, "AG"
    modOtkupUI.NewFieldG z, "scrAgArt", Poruka("OTKUI_FLD_AG_ARTIKAL"), "cmb", "", 1, False, False, "AG"
    modOtkupUI.NewFieldG z, "scrAgPar", Poruka("OTKUI_FLD_AG_PARCELA"), "cmb", "", 1, False, False, "AG"
    modOtkupUI.NewFieldG z, "scrAgDob", Poruka("OTKUI_FLD_AG_DOBAVLJAC"), "txt", "", 1, False, False, "AG"
    ' spisak izabranih parcela stoji na mestu natpisa cetvrtog polja
    modUiKit.NewLbl z, "agParTxt", "", 0, AG_Y_LBL_A, 200, 12, TS_LABEL, True, C_MUTED, -1
    modUiKit.BtnV z, "scrAgParAdd", Poruka("OTKUI_BTN_AG_PAR_DODAJ"), 0, AG_Y_FLD_A, _
                  100, AG_FLD_H, "soft"
    modUiKit.BtnV z, "scrAgParClr", Poruka("OTKUI_BTN_AG_PAR_OCISTI"), 0, AG_Y_FLD_A, _
                  76, AG_FLD_H, "ghost"

    ' --- red B ---
    ' Natpis polja kolicine se menja sa rezimom (pakovanja / kolicina), pa
    ' jedinica ostaje prazna -- u izdavanju bi "kg" bilo netacno.
    modOtkupUI.NewFieldG z, "scrAgKol", Poruka("OTKUI_FLD_AG_PAKOVANJA"), "txt", "", 1, True, False, "AG"
    modOtkupUI.NewFieldG z, "scrAgCena", Poruka("OTKUI_FLD_AG_CENA"), "txt", "RSD", 1, True, False, "AG"
    modOtkupUI.NewFieldG z, "scrAgDok", Poruka("OTKUI_FLD_AG_BRDOK"), "txt", "", 1, False, False, "AG"

    ' --- objasnjenja ---
    modUiKit.NewLbl z, "agHint", "", PAD, AG_Y_HINT, 400, 12, TS_META, False, C_MUTED, -1
    modUiKit.NewLbl z, "agVred", "", PAD, AG_Y_HINT2, 400, 12, TS_META, True, C_FOREST, -1

    ' --- dugmad ---
    modUiKit.BtnV z, "scrAgDodaj", Poruka("OTKUI_BTN_AG_DODAJ"), PAD, AG_Y_BTN, _
                  150, AG_BTN_H, "soft"
    modUiKit.BtnV z, "scrAgZavrsi", Poruka("OTKUI_BTN_AG_ZAVRSI_IZL"), PAD + 158, AG_Y_BTN, _
                  164, AG_BTN_H, "primary"
    modUiKit.BtnV z, "scrAgPocDug", Poruka("OTKUI_BTN_AG_POC_DUG"), PAD + 330, AG_Y_BTN, _
                  118, AG_BTN_H, "ghost"
    modUiKit.BtnV z, "scrAgOcisti", Poruka("OTKUI_BTN_AG_OCISTI"), PAD + 456, AG_Y_BTN, _
                  118, AG_BTN_H, "ghost"

    modUiKit.NewLbl z, "agLnB", "", 0, AG_ZONA_H - 1, 100, 1, 8, False, 0, C_BORDER
End Sub

Public Function Scr_Layout(ByVal z As Object, ByVal w As Single, ByVal h As Single) As Single
    RasporediPolja z, w
    Scr_Layout = AG_ZONA_H
End Function

' Raspored zavisi od REZIMA (koja polja postoje), pa se ne moze racunati samo
' pri promeni velicine forme: prekidac rezima mora da ga pokrene ponovo.
' Ljuska posle "scr" klika zove RefreshFromData, ne LayoutOtkup -- bez ovoga bi
' polja prijema ostala na koordinatama izdavanja, jedno preko drugog.
Private Sub RasporediPolja(ByVal z As Object, ByVal w As Single)
    Dim i As Long, fw As Single, kx As Single, segDesno As Single
    Dim slot3 As Single, izl As Boolean
    On Error Resume Next
    If z Is Nothing Then Exit Sub
    If w < 200 Then Exit Sub
    izl = (Rezim() <> AG_ULAZ)
    OsveziPrekidacRezima z, izl

    ' Podloga ide od ivice do ivice, sa malim uvlacenjem -- bez njega prvi
    ' natpis stoji zalepljen za belu ivicu. Pokriva oba reda polja i dva reda
    ' objasnjenja; prekidac rezima i dugmad ostaju na krem podlozi zone.
    z.Controls("agBg").Left = PAD - 10
    z.Controls("agBg").top = AG_Y_LBL_A - 8
    z.Controls("agBg").width = w - 2 * (PAD - 10)
    z.Controls("agBg").Height = AG_Y_BTN - AG_Y_LBL_A + 2

    ' Cetiri jednaka polja u redu; parcele u redu A dele slot sa dva dugmeta.
    fw = (w - PAD * 2 - AG_GAP * 3) / 4
    If fw < 90 Then fw = 90
    slot3 = PAD + (fw + AG_GAP) * 3

    ' brojke idu uz desnu ivicu; sakriva se ona koja bi nalegla na prekidac
    segDesno = PAD + 232
    For i = 0 To 3
        kx = w - PAD - (4 - i) * AG_KPI_W
        z.Controls("agKL" & i).Left = kx
        z.Controls("agKV" & i).Left = kx
        z.Controls("agKL" & i).Visible = (kx > segDesno)
        z.Controls("agKV" & i).Visible = (kx > segDesno)
    Next i

    ' KOJA POLJA POSTOJE U OVOM REZIMU. Pali se ovde, a ne u OsveziZonu, iz dva
    ' razloga: raspored i vidljivost su ista odluka (grana 'izl' je jedna), pa se
    ' ne mogu razici; i Scr_Layout dobija zonu ARGUMENTOM, pa se u testu moze
    ' pozvati nad golim okvirom -- OsveziZonu zonu trazi od ljuske i u testu je
    ' nema. Isti raspored kao PoljaPrerade na ekranu Palete.
    PoljeVidi z, "scrAgKoop", izl
    PoljeVidi z, "scrAgPar", izl And IsPracenjeParcela()
    PoljeVidi z, "scrAgDob", Not izl
    PoljeVidi z, "scrAgCena", Not izl
    PoljeVidi z, "scrAgKol", True
    PoljeVidi z, "scrAgDok", True
    z.Controls("agParTxt").Visible = izl And IsPracenjeParcela()
    modUiKit.BoxShow z, "scrAgParAdd", izl And IsPracenjeParcela()
    modUiKit.BoxShow z, "scrAgParClr", izl And IsPracenjeParcela()
    modUiKit.BoxShow z, "scrAgPocDug", izl

    ' RED A. Izdavanje: kooperant, artikal, parcela, dugmad parcela.
    ' Prijem: artikal, dobavljac -- artikal se pomera u prvi slot, jer bi inace
    ' red poceo prazninom tamo gde kooperanta nema.
    If izl Then
        PoljeX z, "scrAgKoop", PAD, fw, AG_Y_LBL_A
        PoljeX z, "scrAgArt", PAD + fw + AG_GAP, fw, AG_Y_LBL_A
        PoljeX z, "scrAgPar", PAD + (fw + AG_GAP) * 2, fw, AG_Y_LBL_A
        z.Controls("agParTxt").Left = slot3
        z.Controls("agParTxt").width = fw
        modUiKit.MoveBox z, "scrAgParAdd", slot3, AG_Y_FLD_A, fw * 0.56
        modUiKit.MoveBox z, "scrAgParClr", slot3 + fw * 0.56 + 6, AG_Y_FLD_A, fw * 0.44 - 6
    Else
        PoljeX z, "scrAgArt", PAD, fw, AG_Y_LBL_A
        PoljeX z, "scrAgDob", PAD + fw + AG_GAP, fw, AG_Y_LBL_A
    End If

    ' RED B. Broj dokumenta je uvek POSLEDNJI: kod prijema iza cene, kod
    ' izdavanja odmah iza broja pakovanja.
    PoljeX z, "scrAgKol", PAD, fw, AG_Y_LBL_B
    If izl Then
        PoljeX z, "scrAgDok", PAD + fw + AG_GAP, fw, AG_Y_LBL_B
    Else
        PoljeX z, "scrAgCena", PAD + fw + AG_GAP, fw, AG_Y_LBL_B
        PoljeX z, "scrAgDok", PAD + (fw + AG_GAP) * 2, fw, AG_Y_LBL_B
    End If

    z.Controls("agHint").width = w - PAD * 2
    z.Controls("agVred").width = w - PAD * 2

    ' RED DUGMADI. "Isprazni korpu" ide uz DESNU ivicu, a ne u niz sa ostalima:
    ' razdvaja radnju koja pravi dokument od one koja baca rad. Ostala tri
    ' teku sleva; "Pocetni dug" u prijemu ne postoji, pa iza njega ne ostaje
    ' rupa jer je poslednji u nizu.
    modUiKit.MoveBtn z, "scrAgDodaj", PAD, AG_Y_BTN
    modUiKit.MoveBtn z, "scrAgZavrsi", PAD + 158, AG_Y_BTN
    modUiKit.MoveBtn z, "scrAgPocDug", PAD + 330, AG_Y_BTN
    modUiKit.MoveBtn z, "scrAgOcisti", w - PAD - 118, AG_Y_BTN

    z.Controls("agLnB").width = w
End Sub

' Polje je OKVIR: pomera se i siri kao celina, a unutrasnjost prerasporedjuje
' ljuska. Bez LayoutFieldInner unutrasnje kontrole ostaju na merama iz gradnje
' (180pt), pa se jedinica nadje nasred polja a unos izgleda odsecen.
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
    Set Zona = modOtkupUI.ScreenZone("AGRO")
End Function

'=====================================================================
' STANJE EKRANA
'=====================================================================
Private Function Rezim() As String
    If Len(mMod) = 0 Then mMod = AG_IZLAZ
    Rezim = mMod
End Function

Private Function Korpa() As Collection
    If Rezim() = AG_ULAZ Then
        If mKorpaU Is Nothing Then Set mKorpaU = modAgroUnos.NovaAgroKorpa()
        Set Korpa = mKorpaU
    Else
        If mKorpaI Is Nothing Then Set mKorpaI = modAgroUnos.NovaAgroKorpa()
        Set Korpa = mKorpaI
    End If
End Function

Private Function BrojUKorpi(ByVal k As Collection) As Long
    If Not k Is Nothing Then BrojUKorpi = k.count
End Function

' Kesiran opis artikla. Cisti ga Scr_ResetCache -- posle upisa se cena moze
' promeniti u sifarniku.
Private Function ArtInfo(ByVal artID As String) As Object
    If mArtKes Is Nothing Then
        Set mArtKes = CreateObject("Scripting.Dictionary")
        mArtKes.CompareMode = vbTextCompare
    End If
    If Not mArtKes.Exists(artID) Then _
        Set mArtKes(artID) = modAgroUnos.AgroArtikalInfo(artID)
    Set ArtInfo = mArtKes(artID)
End Function

' Ljuska ovo zove posle SVAKE radnje koja je promenila podatke -- pa ovde sme
' da stoji samo ono sto se iz podataka moze ponovo izvesti. Korpa NE sme: ona
' je nesnimljen rad operatera, a ne izvedena mapa.
Public Sub Scr_ResetCache()
    Set mArtKes = Nothing
    Set mDugKes = Nothing
    Set mArtIds = Nothing
    Set mDugIds = Nothing
End Sub

' Dug kooperanta = zaduzenje iz magacina minus odbitak iz tblNovac. Ista
' formula koju frmAgrohemija pokazuje uz izabranog kooperanta.
Private Function DugZaKoop(ByVal koopID As String) As Double
    If Len(koopID) = 0 Then Exit Function
    If mDugKes Is Nothing Then
        Set mDugKes = CreateObject("Scripting.Dictionary")
        mDugKes.CompareMode = vbTextCompare
    End If
    If Not mDugKes.Exists(koopID) Then _
        mDugKes(koopID) = GetAgrohemijaDug(koopID) - GetAgroAbzug(koopID)
    DugZaKoop = CDbl(mDugKes(koopID))
End Function

'=====================================================================
' COMBO-I
'=====================================================================
' Kooperanti i artikli se pune jednom po otvaranju ekrana. Punjenje pise u
' kontrolu i time okida Change, pa ide pod mFill.
Private Sub PuniCombos()
    Dim z As Object, CB As Object, mapa As Object, k As Variant
    Dim data As Variant, colID As Long, colNaziv As Long, colJM As Long, i As Long
    On Error GoTo EH
    If mCombosPunjeni Then Exit Sub
    Set z = Zona()
    If z Is Nothing Then Exit Sub

    mFill = True
    mStep = "kooperanti"
    Set CB = z.Controls("scrAgKoop").Controls("scrAgKoopT")
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

    mStep = "artikli"
    Set CB = z.Controls("scrAgArt").Controls("scrAgArtT")
    CB.Clear
    CB.ColumnCount = 2
    CB.ColumnWidths = "180 pt;0 pt"
    CB.BoundColumn = 1
    CB.TextColumn = 1
    data = GetTableData(TBL_ARTIKLI)
    If IsArray(data) Then
        colID = GetColumnIndex(TBL_ARTIKLI, COL_ART_ID)
        colNaziv = GetColumnIndex(TBL_ARTIKLI, COL_ART_NAZIV)
        colJM = GetColumnIndex(TBL_ARTIKLI, COL_ART_JM)
        For i = 1 To UBound(data, 1)
            ' Rezervisani virtuelni artikal (pocetni dug) NIJE roba i ne sme u
            ' listu -- isto izuzimanje kao u legacy LoadArtikli i u
            ' GetMagacinStanje. Ne dirati.
            If Len(Trim$(CStr(data(i, colID)))) > 0 Then
                If CStr(data(i, colID)) <> ART_POCETNI_DUG Then
                    CB.AddItem CStr(data(i, colNaziv)) & " [" & CStr(data(i, colJM)) & "]"
                    CB.List(CB.ListCount - 1, 1) = CStr(data(i, colID))
                End If
            End If
        Next i
    End If

    mCombosPunjeni = True
    mFill = False
    Exit Sub
EH:
    mFill = False
    ' Prazan combo bez traga je bio glavni razlog zasto je izgledalo da "nista
    ' nije povezano" -- isto kao u modOtkupUI.FillCombos.
    Debug.Print "modScrAgro.PuniCombos PAO na koraku [" & mStep & "]: " & _
                Err.Number & " " & Err.description
End Sub

' Parcele izabranog kooperanta. Racun je u modAgrohemija.GetParceleByKooperant;
' ovde je samo punjenje liste i mapa ha, koju smart doza sabira.
Private Sub PuniParcele()
    Dim z As Object, CB As Object, p As Variant, i As Long, koopID As String
    Dim pre As Boolean
    On Error Resume Next
    Set z = Zona()
    If z Is Nothing Then Exit Sub

    pre = mFill
    mFill = True
    Set CB = z.Controls("scrAgPar").Controls("scrAgParT")
    CB.Clear
    CB.ColumnCount = 2
    CB.ColumnWidths = "180 pt;0 pt"
    CB.BoundColumn = 1
    CB.TextColumn = 1
    Set mParMapa = CreateObject("Scripting.Dictionary")
    mParMapa.CompareMode = vbTextCompare
    OcistiParcele

    koopID = IzabranKooperant()
    ' Pracenje parcela OFF -> lista se ne puni i polje je ugaseno; unos ide bez
    ' parcele, a smart doza se preskace. Isti flag kao u frmOtkup.
    If Len(koopID) > 0 And IsPracenjeParcela() Then
        p = GetParceleByKooperant(koopID)
        If IsArray(p) Then
            For i = 1 To UBound(p, 1)
                CB.AddItem CStr(p(i, 6))
                CB.List(CB.ListCount - 1, 1) = CStr(p(i, 1))
                mParMapa(CStr(p(i, 1))) = CDbl(p(i, 5))
            Next i
        End If
    End If
    mFill = pre
End Sub

' Kontrola polja. Polje je okvir 'nm', kontrola u njemu je 'nmT' -- isti oblik
' koji trazi i panel za izbor (modOtkupUI.FindCombo). Nothing kad zona jos nije
' izgradjena (test).
Private Function Kontrola(ByVal nm As String) As Object
    Dim z As Object
    On Error Resume Next
    Set z = Zona()
    If z Is Nothing Then Exit Function
    Set Kontrola = z.Controls(nm).Controls(nm & "T")
End Function

Private Function IzabranKooperant() As String
    Dim c As Object
    On Error Resume Next
    Set c = Kontrola("scrAgKoop")
    If c Is Nothing Then Exit Function
    IzabranKooperant = GetComboID(c)
End Function

Private Function IzabranArtikal() As String
    Dim c As Object
    On Error Resume Next
    Set c = Kontrola("scrAgArt")
    If c Is Nothing Then Exit Function
    IzabranArtikal = GetComboID(c)
End Function

Private Function PoljeTekst(ByVal nm As String) As String
    Dim c As Object
    On Error Resume Next
    Set c = Kontrola(nm)
    If c Is Nothing Then Exit Function
    PoljeTekst = Trim$(CStr(c.text))
End Function

' Programski upis u polje. Zastavica se CUVA i VRACA, ne gasi bezuslovno:
' PostaviPolje se zove i iz rutina koje su same vec pod mFill (punjenje
' parcela -> preporuka -> upis broja pakovanja), pa bi golo mFill = False
' roditelju skinulo zastitu na pola posla.
Private Sub PostaviPolje(ByVal nm As String, ByVal v As String)
    Dim c As Object, pre As Boolean
    On Error Resume Next
    Set c = Kontrola(nm)
    If c Is Nothing Then Exit Sub
    pre = mFill
    mFill = True
    c.text = v
    mFill = pre
End Sub

Private Sub OcistiParcele()
    Set mParIds = New Collection
    mParHa = 0#
End Sub

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
    LogErr "modScrAgro.Scr_Event"
    modOtkupUI.ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & errDesc, True
    Err.Clear
End Function

' Vraca True samo kad su se PODACI promenili (mreza se tada cita ponovo).
Private Function ObradiKlik(ByVal tag As String) As Boolean
    Dim nov As String

    If Left$(tag, 2) = "ls" Then
        If Mid$(tag, 3) = Scr_Lista() Then Exit Function
        mLista = Mid$(tag, 3)
        ObradiKlik = True
        Exit Function
    End If

    ' Izbor reda ne menja podatke ni u jednoj listi: korpa se menja radnjom nad
    ' redom, a stanje / promet / dugovi su pregledi.
    If Left$(tag, 4) = "row:" Then Exit Function

    ' Promena u polju zone. Ljuska je salje kao "chg:<tag kontrole>", simetricno
    ' sa "act:" i "row:" -- vrednost se cita iz same kontrole, ne iz taga.
    ' Vraca False: preracun je prikaz, mreza se zbog njega ne cita ponovo.
    If Left$(tag, 4) = "chg:" Then
        ObradiPromenu Mid$(tag, 5)
        Exit Function
    End If

    ' Dvoklik na red PREUZIMA red u unos: iz dugova kooperanta, iz stanja
    ' artikal. To je jedan potez umesto tri (zapamti ime, predji na korpu,
    ' nadji ga u padajucoj listi). Vraca False -- lista se nije promenila,
    ' promenila se zona.
    If Left$(tag, 4) = "dbl:" Then
        DvoklikNaRed CLng(val(Mid$(tag, 5)))
        Exit Function
    End If

    If Left$(tag, 4) = "act:" Then
        ObradiKlik = RadnjaNadRedom(tag)
        Exit Function
    End If

    Select Case tag
        Case "scrAgSegI", "scrAgSegU"
            nov = IIf(tag = "scrAgSegU", AG_ULAZ, AG_IZLAZ)
            If nov = Rezim() Then Exit Function
            mMod = nov
            OsveziZonu
            ' korpa je druga, pa i lista mora da se procita ponovo
            ObradiKlik = (Scr_Lista() = AG_KORPA)

        Case "scrAgParAdd": DodajParcelu
        Case "scrAgParClr"
            OcistiParcele
            OsveziZonu

        Case "scrAgDodaj": ObradiKlik = DodajUKorpu()
        Case "scrAgZavrsi": ObradiKlik = ZavrsiUnos()
        Case "scrAgOcisti": ObradiKlik = IsprazniKorpu()
        Case "scrAgPocDug": ObradiKlik = PocetniDug()
    End Select
End Function

' Tag je ime KONTROLE (nmT), ne polja -- tako ga ljuska i salje.
Private Sub ObradiPromenu(ByVal tag As String)
    If mFill Then Exit Sub
    Select Case tag
        Case "scrAgKoopT"
            PuniParcele
            OsveziZonu
        Case "scrAgArtT"
            ' Prijem predlaze cenu iz sifarnika, izdavanje racuna smart dozu --
            ' oboje na promenu artikla, kao u legacy formi.
            If Rezim() = AG_ULAZ Then PredloziCenu
            PredloziKolicinu
            OsveziZonu
        Case "scrAgParT"
            ' izbor u listi parcela nista ne menja dok se ne potvrdi dugmetom
        Case "scrAgKolT", "scrAgCenaT"
            OsveziZonu
    End Select
End Sub

' Radnja nad redom mreze: "act:<kljuc>:<red>".
Private Function RadnjaNadRedom(ByVal tag As String) As Boolean
    Dim p() As String, red As Long
    p = Split(Mid$(tag, 5), ":")
    If UBound(p) < 1 Then Exit Function
    red = CLng(val(p(1)))
    If p(0) <> "agdel" Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & p(0), True
        Exit Function
    End If
    If Scr_Lista() <> AG_KORPA Then Exit Function
    If red < 1 Or red > BrojUKorpi(Korpa()) Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_NEMA_REDA"), True
        Exit Function
    End If
    ' Mreza je SORTIRANA, pa red u prikazu nije red u korpi. Stavka se zato
    ' trazi po vrednostima iz prikazanog reda, ne po njegovom rednom broju --
    ' bez toga bi klik na "Ukloni" nad sortiranom listom obrisao drugu stavku.
    RadnjaNadRedom = UkloniPoPrikazu(red)
End Function

' Dvoklik: preuzmi red u unos. Bira se po IDENTITETU iz mape koju je napunio
' citac liste, ne po tekstu reda -- a kad je prikaz dvosmislen (dva kooperanta
' istog imena), mapa nosi prazno i dvoklik ODBIJA da bira. Pogadjanje bi ovde
' izdalo robu pogresnom coveku.
Private Sub DvoklikNaRed(ByVal red As Long)
    Dim prikaz As String, iD As String

    If red < 1 Then Exit Sub
    prikaz = Trim$(CStr(modOtkupUI.GridCell(red, 1)))
    If Len(prikaz) = 0 Then Exit Sub

    Select Case Scr_Lista()
        Case AG_DUGOVI
            If mDugIds Is Nothing Then Exit Sub
            If Not mDugIds.Exists(prikaz) Then Exit Sub
            iD = CStr(mDugIds(prikaz))
            If Len(iD) = 0 Then
                modOtkupUI.ShowToast Poruka("OTKUI_ERR_AG_DVOSMISLEN"), True
                Exit Sub
            End If
            ' Dug se izdaje, ne prima -- dvoklik zato i prebacuje u IZDAVANJE.
            mMod = AG_IZLAZ
            IzaberiUComboPoId "scrAgKoop", iD
            PuniParcele
            OsveziZonu
            modOtkupUI.ShowToast Poruka("OTKUI_MSG_AG_UZET_KOOP") & " " & prikaz, False

        Case AG_STANJE
            If mArtIds Is Nothing Then Exit Sub
            If Not mArtIds.Exists(prikaz) Then Exit Sub
            iD = CStr(mArtIds(prikaz))
            If Len(iD) = 0 Then
                modOtkupUI.ShowToast Poruka("OTKUI_ERR_AG_DVOSMISLEN"), True
                Exit Sub
            End If
            IzaberiUComboPoId "scrAgArt", iD
            If Rezim() = AG_ULAZ Then PredloziCenu
            PredloziKolicinu
            OsveziZonu
            modOtkupUI.ShowToast Poruka("OTKUI_MSG_AG_UZET_ART") & " " & prikaz, False
    End Select
End Sub

' Postavi combo na stavku sa datim ID-em. ID je u SKRIVENOJ koloni (kolona 1),
' pa se bira red a ne tekst -- upis teksta bi kod dva ista naziva izabrao prvi.
Private Sub IzaberiUComboPoId(ByVal nm As String, ByVal iD As String)
    Dim c As Object, i As Long, pre As Boolean
    On Error Resume Next
    Set c = Kontrola(nm)
    If c Is Nothing Then Exit Sub
    pre = mFill
    mFill = True
    For i = 0 To c.ListCount - 1
        If Trim$(CStr(c.List(i, 1))) = Trim$(iD) Then
            c.ListIndex = i
            Exit For
        End If
    Next i
    mFill = pre
End Sub

' Izbaci red korpe koji odgovara prikazanom redu mreze (artikal + kolicina).
Private Function UkloniPoPrikazu(ByVal red As Long) As Boolean
    Dim k As Collection, i As Long, artNaziv As String, kol As Double
    Set k = Korpa()
    If k Is Nothing Then Exit Function
    artNaziv = Trim$(CStr(modOtkupUI.GridCell(red, 1)))
    kol = AgD(modOtkupUI.GridCell(red, 4))
    For i = 1 To k.count
        If Trim$(CStr(k(i)("naziv"))) = artNaziv Then
            If Abs(CDbl(k(i)("kolicina")) - kol) < 0.0000001 Then
                k.Remove i
                OsveziZonu
                modOtkupUI.ShowToast Poruka("OTKUI_MSG_AG_UKLONJENO"), False
                UkloniPoPrikazu = True
                Exit Function
            End If
        End If
    Next i
    modOtkupUI.ShowToast Poruka("OTKUI_ERR_NEMA_REDA"), True
End Function

'=====================================================================
' RADNJE ZONE
'=====================================================================
Private Sub DodajParcelu()
    Dim z As Object, parID As String, i As Long
    Set z = Zona()
    If z Is Nothing Then Exit Sub
    If Not IsPracenjeParcela() Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_AG_PARCELE_OFF"), True
        Exit Sub
    End If
    parID = GetComboID(z.Controls("scrAgPar").Controls("scrAgParT"))
    If Len(parID) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_AG_NEMA_PARCELE"), True
        Exit Sub
    End If
    If mParIds Is Nothing Then OcistiParcele
    For i = 1 To mParIds.count
        If CStr(mParIds(i)) = parID Then
            modOtkupUI.ShowToast Poruka("OTKUI_ERR_AG_PARCELA_VEC"), True
            Exit Sub
        End If
    Next i
    mParIds.Add parID
    If Not mParMapa Is Nothing Then
        If mParMapa.Exists(parID) Then mParHa = mParHa + CDbl(mParMapa(parID))
    End If
    ' Nova parcela menja zbir ha, pa i preporuku -- isto kao lstParcele_Click.
    PredloziKolicinu
    OsveziZonu
End Sub

' Smart doza -> broj pakovanja u polje kolicine. Racun je u modAgroUnos;
' ovde se samo upisuje predlog i sastavlja recenica.
Private Sub PredloziKolicinu()
    Dim pre As Object, artID As String
    If Rezim() = AG_ULAZ Then Exit Sub
    artID = IzabranArtikal()
    If Len(artID) = 0 Then Exit Sub
    If Not IsPracenjeParcela() Then Exit Sub
    If mParHa <= 0 Then Exit Sub
    Set pre = modAgroUnos.AgroPreporukaInfo(artID, mParHa)
    If Len(CStr(pre("greska"))) > 0 Then Exit Sub
    If CLng(pre("brojPak")) <= 0 Then Exit Sub
    PostaviPolje "scrAgKol", CStr(pre("brojPak"))
End Sub

Private Sub PredloziCenu()
    Dim artID As String, info As Object
    artID = IzabranArtikal()
    If Len(artID) = 0 Then Exit Sub
    Set info = ArtInfo(artID)
    ' Format mora biti isti onaj koji ParseNum ume da vrati (tacka = hiljade,
    ' zarez = decimale). CStr bi na nekim lokalima dao "1234.5", a ParseNum bi
    ' iz toga procitao 12345 -- prijem bi tiho knjizio desetostruku cenu.
    PostaviPolje "scrAgCena", Format$(CDbl(info("cena")), "#,##0.00")
End Sub

Private Function DodajUKorpu() As Boolean
    Dim greska As String, fokus As String, parID As String, i As Long
    Dim artID As String

    artID = IzabranArtikal()
    If Rezim() = AG_ULAZ Then
        greska = modAgroUnos.AgroDodajUlaz(Korpa(), artID, _
                     modOtkupUI.ParseNum(PoljeTekst("scrAgKol")), _
                     modOtkupUI.ParseNum(PoljeTekst("scrAgCena")), fokus)
    Else
        If Not mParIds Is Nothing Then
            For i = 1 To mParIds.count
                If Len(parID) > 0 Then parID = parID & ";"
                parID = parID & CStr(mParIds(i))
            Next i
        End If
        greska = modAgroUnos.AgroDodajIzlaz(Korpa(), artID, _
                     modOtkupUI.ParseNum(PoljeTekst("scrAgKol")), parID, fokus)
    End If

    If greska = AGRO_ODUSTAO Then Exit Function
    If Len(greska) > 0 Then
        modOtkupUI.ShowToast greska, True
        Exit Function
    End If

    ' Posle dodavanja se prazni SAMO ono sto pripada stavci -- kooperant, broj
    ' dokumenta i izabrane parcele ostaju, jer se sledeca stavka najcesce izdaje
    ' istom kooperantu po istom dokumentu (legacy radi isto).
    ZapamtiIzbor "scrAgArt", ""
    PostaviPolje "scrAgKol", ""
    If Rezim() = AG_ULAZ Then PostaviPolje "scrAgCena", ""
    ' Potvrda mora da postoji i kad korpa NIJE prikazana lista: bez nje jedini
    ' trag da je stavka usla je brojka u zoni, koju operater u tom trenutku ne
    ' gleda -- pa isti unos ode dva puta.
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_AG_U_KORPI") & " " & _
                         BrojUKorpi(Korpa()), False
    OsveziZonu
    ' korpa je lista -- ako je prikazana, mora da se procita ponovo
    DodajUKorpu = (Scr_Lista() = AG_KORPA)
End Function

Private Sub ZapamtiIzbor(ByVal nm As String, ByVal v As String)
    Dim c As Object, pre As Boolean
    On Error Resume Next
    Set c = Kontrola(nm)
    If c Is Nothing Then Exit Sub
    pre = mFill
    mFill = True
    c.ListIndex = -1
    c.text = v
    mFill = pre
End Sub

Private Function PotvrdiPraznjenje() As Boolean
    PotvrdiPraznjenje = (MsgBox(Poruka("OTKUI_ASK_AG_OCISTI"), _
                                vbQuestion + vbYesNo, APP_NAME) = vbYes)
End Function

Private Function IsprazniKorpu() As Boolean
    If BrojUKorpi(Korpa()) = 0 Then Exit Function
    If Not PotvrdiPraznjenje() Then Exit Function
    If Rezim() = AG_ULAZ Then
        Set mKorpaU = modAgroUnos.NovaAgroKorpa()
    Else
        Set mKorpaI = modAgroUnos.NovaAgroKorpa()
    End If
    OsveziZonu
    IsprazniKorpu = (Scr_Lista() = AG_KORPA)
End Function

Private Function IspravnaKorpa() As Boolean
    IspravnaKorpa = (BrojUKorpi(Korpa()) > 0)
End Function

Private Function ZavrsiUnos() As Boolean
    Dim greska As String, upisano As Long, brDok As String

    If Not IspravnaKorpa() Then
        modOtkupUI.ShowToast Poruka("AGROU_ERR_KORPA_PRAZNA"), True
        Exit Function
    End If
    brDok = PoljeTekst("scrAgDok")

    If Rezim() = AG_ULAZ Then
        greska = modAgroUnos.AgroUpisiUlaz(Korpa(), PoljeTekst("scrAgDob"), _
                                           brDok, Date, upisano)
    Else
        greska = modAgroUnos.AgroUpisiIzlaz(Korpa(), IzabranKooperant(), _
                                            brDok, Date, upisano)
    End If

    If Len(greska) > 0 Then
        modOtkupUI.ShowToast greska, True
        Exit Function
    End If

    ' Upisano -> korpa se prazni, broj dokumenta i izbor kooperanta se resetuju.
    ' Sledeci dokument je nov dokument; zadrzavanje broja bi drugu isporuku
    ' knjizilo pod istim brojem.
    If Rezim() = AG_ULAZ Then
        Set mKorpaU = modAgroUnos.NovaAgroKorpa()
        PostaviPolje "scrAgDob", ""
    Else
        Set mKorpaI = modAgroUnos.NovaAgroKorpa()
        ZapamtiIzbor "scrAgKoop", ""
        OcistiParcele
        PuniParcele
    End If
    PostaviPolje "scrAgDok", ""
    Set mDugKes = Nothing            ' izdavanje je promenilo dug
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_AG_UPISANO") & " " & brDok & _
                         " (" & upisano & ")", False
    OsveziZonu
    ZavrsiUnos = True
End Function

' Pocetni dug kooperanta (migracija). Ceo racun je u
' modAgrohemija.BookPocetniDug -- ovde su samo dva pitanja i potvrda, isto kao
' u legacy dugmetu.
Private Function PocetniDug() As Boolean
    Dim koopID As String, odg As String, iznos As Double, brDok As String
    Dim novID As String, trenutni As Double

    koopID = IzabranKooperant()
    If Len(koopID) = 0 Then
        modOtkupUI.ShowToast Poruka("AGROU_ERR_NEMA_KOOPERANTA"), True
        Exit Function
    End If

    trenutni = DugZaKoop(koopID)
    odg = InputBox(Poruka("OTKUI_ASK_AG_POC_DUG") & vbCrLf & vbCrLf & _
                   Poruka("OTKUI_LBL_AG_DUG") & " " & Format$(trenutni, "#,##0"), _
                   APP_NAME)
    If Len(Trim$(odg)) = 0 Then Exit Function
    If Not IsNumeric(odg) Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_AG_IZNOS"), True
        Exit Function
    End If
    iznos = CDbl(odg)
    If iznos <= 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_AG_IZNOS"), True
        Exit Function
    End If

    brDok = Trim$(InputBox(Poruka("OTKUI_ASK_AG_POC_BRDOK"), APP_NAME, "POC-DUG"))
    If MsgBox(Poruka("OTKUI_ASK_AG_POC_POTVRDA") & " " & _
              Format$(iznos, "#,##0") & "?", vbQuestion + vbYesNo, APP_NAME) <> vbYes Then
        Exit Function
    End If

    novID = BookPocetniDug(koopID, iznos, brDok, Date)
    If Len(Trim$(novID)) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_AG_POC_DUG"), True
        Exit Function
    End If
    Set mDugKes = Nothing            ' dug se upravo promenio
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_AG_POC_DUG") & " " & novID, False
    OsveziZonu
    PocetniDug = True
End Function

'=====================================================================
' OSVEZAVANJE ZONE
'=====================================================================
Private Sub OsveziZonu()
    Dim z As Object
    On Error Resume Next
    Set z = Zona()
    If z Is Nothing Then Exit Sub
    PuniCombos
    OsveziRezim z
    RasporediPolja z, z.width
    OsveziObjasnjenje z
    OsveziBrojke z
End Sub

' Koja polja i koja dugmad postoje u ovom rezimu, i kako se zovu.
' IZABRAN REZIM SE MORA I JAVITI SINK-U, ne samo obojiti.
'
' clsFlatBtn pamti osnovnu boju pri Bind-u i na izlazak pokazivaca je VRACA
' (ResetVisual). BoxState menja kontrolu, ali ne i tu zapamcenu osnovu -- pa je
' izabran rezim bio zelen tek dok je pokazivac nad njim: cim predje dalje,
' ispuna se vrati na belu, a natpis ostane krem (labela natpisa je vezana kao
' "chev" i nju reset ne dira), pa aktivno dugme postane skoro necitljivo.
' Tacno to je operater prijavio na prvom smoke-u.
'
' RebaseSink je bas za to: render koji promeni boju javlja novu osnovu. Isti
' kvar i ista popravka kao StilDugmeta u modScrStorno, gde se videlo samo na
' jednom od cetiri dugmeta jer su ostala tri ionako tamna na belom.
'
' Vrsta "seg" (v. Scr_Build) resava drugu polovinu: izabrano dugme hover uopste
' ne prefarbava, isto kao prekidac lista nad mrezom. RebaseSink i dalje treba --
' bez njega bi dugme koje je PRESTALO da bude izabrano vratilo zelenu osnovu
' zapamcenu pri gradnji. Zato oba, tacno kao RefreshListSeg u ljusci.
'
' Stoji uz RASPORED, a ne uz osvezavanje zone, iz istog razloga kao vidljivost
' polja: koji je rezim izabran je JEDNA odluka, pa boja i raspored ne mogu da se
' raziju -- i Scr_Layout dobija zonu argumentom, pa se moze izmeriti u testu.
Private Sub OsveziPrekidacRezima(ByVal z As Object, ByVal izl As Boolean)
    On Error Resume Next
    modUiKit.BoxState z, "scrAgSegI", IIf(izl, C_FOREST, C_WHITE), _
                      IIf(izl, C_CREAM, C_FOREST), izl
    modUiKit.BoxState z, "scrAgSegU", IIf(izl, C_WHITE, C_FOREST), _
                      IIf(izl, C_FOREST, C_CREAM), Not izl
    modOtkupUI.RebaseSink "scrAgSegI"
    modOtkupUI.RebaseSink "scrAgSegU"
End Sub

Private Sub OsveziRezim(ByVal z As Object)
    Dim izl As Boolean
    izl = (Rezim() <> AG_ULAZ)

    z.Controls("scrAgKol").Controls("scrAgKolL").caption = UCase$(Poruka(IIf(izl, _
        "OTKUI_FLD_AG_PAKOVANJA", "OTKUI_FLD_AG_KOLICINA")))
    z.Controls("scrAgDodajC").caption = Poruka(IIf(izl, _
        "OTKUI_BTN_AG_DODAJ", "OTKUI_BTN_AG_PRIJEM"))
    z.Controls("scrAgZavrsiC").caption = Poruka(IIf(izl, _
        "OTKUI_BTN_AG_ZAVRSI_IZL", "OTKUI_BTN_AG_ZAVRSI_ULZ"))

    ' spisak izabranih parcela (vidljivost polja postavlja RasporediPolja)
    If izl And IsPracenjeParcela() Then
        If mParIds Is Nothing Then
            z.Controls("agParTxt").caption = UCase$(Poruka("OTKUI_LBL_AG_PAR_NEMA"))
        ElseIf mParIds.count = 0 Then
            z.Controls("agParTxt").caption = UCase$(Poruka("OTKUI_LBL_AG_PAR_NEMA"))
        Else
            z.Controls("agParTxt").caption = UCase$(mParIds.count & " " & _
                Poruka("OTKUI_LBL_AG_PAR_IZAB") & " " & _
                Format$(mParHa, "0.00") & " ha")
        End If
    End If
End Sub

' Polje je jedan okvir, pa se gasi jednim potezom -- natpis, ivica, kontrola,
' strelica i jedinica su unutra.
Private Sub PoljeVidi(ByVal z As Object, ByVal nm As String, ByVal vis As Boolean)
    On Error Resume Next
    z.Controls(nm).Visible = vis
End Sub

' Dve recenice ispod polja: sta smart doza predlaze i koliko stavka vredi.
' Ista dva podatka koja legacy drzi u lblPreporuka i lblVrednost.
Private Sub OsveziObjasnjenje(ByVal z As Object)
    Dim artID As String, info As Object, pre As Object
    Dim kol As Double, cena As Double, ukupno As Double, jm As String

    z.Controls("agHint").caption = ""
    z.Controls("agVred").caption = ""
    z.Controls("agHint").ForeColor = C_MUTED

    artID = IzabranArtikal()
    If Len(artID) = 0 Then
        z.Controls("agHint").caption = Poruka("OTKUI_HINT_AG_IZABERI")
        Exit Sub
    End If
    Set info = ArtInfo(artID)
    jm = CStr(info("jm"))

    If Rezim() = AG_ULAZ Then
        ' prijem: doza iz sifarnika je informacija, ne racun
        z.Controls("agHint").caption = Poruka("OTKUI_HINT_AG_DOZA") & " " & _
            Format$(CDbl(info("doza")), "0.##") & " " & jm & "/ha"
        kol = modOtkupUI.ParseNum(PoljeTekst("scrAgKol"))
        cena = modOtkupUI.ParseNum(PoljeTekst("scrAgCena"))
        If kol > 0 Then
            z.Controls("agVred").caption = modAgroUnos.AgroFmtKol(kol) & " " & jm & _
                "  " & ChrW(215) & "  " & Format$(cena, "#,##0") & "  =  " & _
                Format$(kol * cena, "#,##0") & " RSD"
        End If
        Exit Sub
    End If

    ' izdavanje: invarijanta nad pakovanjem je kapija, pa se prijavljuje odmah
    If Len(CStr(info("greska"))) > 0 Then
        z.Controls("agHint").caption = CStr(info("greska"))
        z.Controls("agHint").ForeColor = C_RUST
        Exit Sub
    End If

    If IsPracenjeParcela() Then
        If mParHa > 0 Then
            Set pre = modAgroUnos.AgroPreporukaInfo(artID, mParHa)
            z.Controls("agHint").caption = Poruka("OTKUI_HINT_AG_DOZA") & " " & _
                Format$(CDbl(pre("dozaKg")), "0.00") & " " & jm & " " & _
                Poruka("OTKUI_HINT_AG_ZA") & " " & Format$(mParHa, "0.00") & " ha " & _
                ChrW(8212) & " " & Poruka("OTKUI_HINT_AG_IZDAJ") & " " & _
                CStr(pre("brojPak")) & " " & ChrW(215) & " " & _
                modAgroUnos.AgroFmtKol(CDbl(info("pakovanje"))) & " " & jm
        Else
            z.Controls("agHint").caption = Poruka("OTKUI_HINT_AG_PARCELE")
        End If
    Else
        z.Controls("agHint").caption = Poruka("OTKUI_HINT_AG_RUCNO")
    End If

    kol = modOtkupUI.ParseNum(PoljeTekst("scrAgKol"))
    If kol > 0 Then
        ukupno = kol * CDbl(info("pakovanje"))
        z.Controls("agVred").caption = modAgroUnos.AgroFmtKol(kol) & " " & _
            ChrW(215) & " " & modAgroUnos.AgroFmtKol(CDbl(info("pakovanje"))) & " " & jm & _
            "  =  " & modAgroUnos.AgroFmtKol(ukupno) & " " & jm & "  |  " & _
            Format$(ukupno * CDbl(info("cena")), "#,##0") & " RSD"
    End If
End Sub

' Cetiri brojke -- iste one koje legacy drzi u KPI traci: dug kooperanta,
' obe korpe i dug posle izdavanja.
Private Sub OsveziBrojke(ByVal z As Object)
    Dim koopID As String, dug As Double, zbirI As Double, zbirU As Double

    koopID = IzabranKooperant()
    dug = DugZaKoop(koopID)
    zbirI = modAgroUnos.AgroZbirKorpe(mKorpaI)
    zbirU = modAgroUnos.AgroZbirKorpe(mKorpaU)

    z.Controls("agKL0").caption = UCase$(Poruka("OTKUI_KPI_AG_DUG"))
    If Len(koopID) = 0 Then
        z.Controls("agKV0").caption = ChrW(8212)
    Else
        z.Controls("agKV0").caption = Format$(dug, "#,##0")
    End If

    z.Controls("agKL1").caption = UCase$(Poruka("OTKUI_KPI_AG_KORPA_IZL"))
    z.Controls("agKV1").caption = BrojUKorpi(mKorpaI) & " / " & Format$(zbirI, "#,##0")

    z.Controls("agKL2").caption = UCase$(Poruka("OTKUI_KPI_AG_KORPA_ULZ"))
    z.Controls("agKV2").caption = BrojUKorpi(mKorpaU) & " / " & Format$(zbirU, "#,##0")

    z.Controls("agKL3").caption = UCase$(Poruka("OTKUI_KPI_AG_DUG_POSLE"))
    If Len(koopID) = 0 Then
        z.Controls("agKV3").caption = ChrW(8212)
    Else
        z.Controls("agKV3").caption = Format$(dug + zbirI, "#,##0")
    End If
End Sub

'=====================================================================
' REDOVI ZA MREZU
'=====================================================================
Public Function Scr_Rows(ByVal filter As String, ByVal q As String) As Variant
    ' Zona se puni odavde, kao i na ekranu Palete: gradi se jednom, a podaci
    ' za nju postoje tek kad se lista cita.
    OsveziZonu
    Select Case Scr_Lista()
        Case AG_STANJE: Scr_Rows = RedoviStanje(filter, q): Exit Function
        Case AG_PROMET: Scr_Rows = RedoviPromet(filter, q): Exit Function
        Case AG_DUGOVI: Scr_Rows = RedoviDugovi(filter, q): Exit Function
    End Select
    ' Korpa nema cipove, pa filter ne gleda -- v. Scr_Cipovi.
    Scr_Rows = RedoviKorpa(q)
End Function

Private Function PrazanRezultat(ByVal kolone As Variant) As Variant
    PrazanRezultat = Array(kolone, Empty, 0, 0#, 0#, Array(0, 0, 0))
End Function

'--------------------------------------------------------- LISTA: KORPA
Private Function KorpaKolone() As Variant
    KorpaKolone = Array( _
        "OTKUI_HDA_ARTIKAL||part|0|1", _
        "OTKUI_HDA_JM||txt|46|2", _
        "OTKUI_HDA_PAKOVANJA||num|76|2", _
        "OTKUI_HDA_KOLICINA||num|86|1", _
        "OTKUI_HD_CENA||rsd|80|1", _
        "OTKUI_HDA_VREDNOST||rsd|94|1", _
        "OTKUI_HDA_PARCELA||txt|130|3")
End Function

Private Function RedoviKorpa(ByVal q As String) As Variant
    Dim k As Collection, i As Long, n As Long, outA() As Variant
    Dim hay As String, zbir As Double
    On Error GoTo EH
    mStep = "korpa"

    Set k = Korpa()
    If k Is Nothing Then
        RedoviKorpa = PrazanRezultat(KorpaKolone())
        Exit Function
    End If
    If k.count = 0 Then
        RedoviKorpa = PrazanRezultat(KorpaKolone())
        Exit Function
    End If

    ReDim outA(1 To k.count, 1 To 7)
    For i = 1 To k.count
        hay = CStr(k(i)("naziv")) & "|" & CStr(k(i)("parcelaID"))
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeci
        End If
        n = n + 1
        outA(n, 1) = CStr(k(i)("naziv"))
        outA(n, 2) = CStr(k(i)("jm"))
        outA(n, 3) = CDbl(k(i)("brojPak"))
        outA(n, 4) = CDbl(k(i)("kolicina"))
        outA(n, 5) = CDbl(k(i)("cena"))
        outA(n, 6) = CDbl(k(i)("vrednost"))
        outA(n, 7) = CStr(k(i)("parcelaID"))
        zbir = zbir + CDbl(k(i)("vrednost"))
Sledeci:
    Next i

    mStep = "OK"
    RedoviKorpa = Array(KorpaKolone(), outA, n, 0#, zbir, Array(0, 0, 0))
    Exit Function
EH:
    Err.Raise Err.Number, "modScrAgro.RedoviKorpa[" & mStep & "]", Err.description
End Function

'-------------------------------------------------------- LISTA: STANJE
Private Function StanjeKolone() As Variant
    StanjeKolone = Array( _
        "OTKUI_HDA_ARTIKAL||part|0|1", _
        "OTKUI_HDA_TIP||txt|96|3", _
        "OTKUI_HDA_JM||txt|46|2", _
        "OTKUI_HDA_ULAZ||num|86|2", _
        "OTKUI_HDA_IZLAZ||num|86|2", _
        "OTKUI_HDA_STANJE||num|92|1")
End Function

' GetMagacinStanje je 1-bazirano i vec izuzima ART_POCETNI_DUG:
'   1 ArtikalID | 2 Naziv | 3 Tip | 4 JM | 5 Ulaz | 6 Izlaz | 7 Stanje
Private Function RedoviStanje(ByVal filter As String, ByVal q As String) As Variant
    Dim src As Variant, i As Long, n As Long, outA() As Variant, hay As String
    Dim naziv As String, artID As String
    On Error GoTo EH
    mStep = "stanje"

    Set mArtIds = CreateObject("Scripting.Dictionary")
    mArtIds.CompareMode = vbTextCompare
    src = GetMagacinStanje()
    If Not IsArray(src) Then
        RedoviStanje = PrazanRezultat(StanjeKolone())
        Exit Function
    End If

    ReDim outA(1 To UBound(src, 1), 1 To 6)
    For i = 1 To UBound(src, 1)
        artID = Trim$(CStr(src(i, 1)))
        naziv = CStr(src(i, 2))
        ' Prikaz -> ID; isti naziv na dva artikla znaci DVOSMISLENO (prazno).
        If mArtIds.Exists(naziv) Then
            If CStr(mArtIds(naziv)) <> artID Then mArtIds(naziv) = ""
        Else
            mArtIds(naziv) = artID
        End If
        If Not AgCipStanje(filter, AgD(src(i, 7))) Then GoTo Sledeci
        hay = artID & "|" & naziv & "|" & CStr(src(i, 3))
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeci
        End If
        n = n + 1
        outA(n, 1) = naziv
        outA(n, 2) = CStr(src(i, 3))
        outA(n, 3) = CStr(src(i, 4))
        outA(n, 4) = AgD(src(i, 5))
        outA(n, 5) = AgD(src(i, 6))
        outA(n, 6) = AgD(src(i, 7))
Sledeci:
    Next i

    mStep = "OK"
    RedoviStanje = Array(StanjeKolone(), outA, n, 0#, 0#, Array(0, 0, 0))
    Exit Function
EH:
    Err.Raise Err.Number, "modScrAgro.RedoviStanje[" & mStep & "]", Err.description
End Function

'-------------------------------------------------------- LISTA: PROMET
Private Function PrometKolone() As Variant
    PrometKolone = Array( _
        "OTKUI_HD_DATUM||date|74|1", _
        "OTKUI_HDA_SMER||txt|56|1", _
        "OTKUI_HDA_ARTIKAL||part|0|1", _
        "OTKUI_HDA_KOLICINA||num|80|1", _
        "OTKUI_HDA_JM||txt|42|3", _
        "OTKUI_HD_CENA||rsd|76|2", _
        "OTKUI_HDA_VREDNOST||rsd|90|1", _
        "OTKUI_HDA_PARTNER||txt|140|1", _
        "OTKUI_HDA_BRDOK||txt|94|2")
End Function

' GetMagacinPrometForGrid je 0-bazirano, 12 kolona (v. modAgrohemija).
Private Function RedoviPromet(ByVal filter As String, ByVal q As String) As Variant
    Dim src As Variant, r As Long, n As Long, outA() As Variant
    Dim hay As String, zbir As Double
    On Error GoTo EH
    mStep = "promet"

    src = GetMagacinPrometForGrid()
    If Not IsArray(src) Then
        RedoviPromet = PrazanRezultat(PrometKolone())
        Exit Function
    End If

    ReDim outA(1 To UBound(src, 1) + 1, 1 To 9)
    For r = 0 To UBound(src, 1)
        If Not AgCipPromet(filter, CStr(src(r, 2)), AgDatum(src(r, 1))) Then GoTo Sledeci
        hay = CStr(src(r, 2)) & "|" & CStr(src(r, 4)) & "|" & _
              CStr(src(r, 9)) & "|" & CStr(src(r, 10))
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeci
        End If
        n = n + 1
        outA(n, 1) = AgDatum(src(r, 1))
        outA(n, 2) = CStr(src(r, 2))
        outA(n, 3) = CStr(src(r, 4))
        outA(n, 4) = AgD(src(r, 6))
        outA(n, 5) = CStr(src(r, 5))
        outA(n, 6) = AgD(src(r, 7))
        outA(n, 7) = AgD(src(r, 8))
        outA(n, 8) = CStr(src(r, 9))
        outA(n, 9) = CStr(src(r, 10))
        zbir = zbir + AgD(src(r, 8))
Sledeci:
    Next r

    mStep = "OK"
    RedoviPromet = Array(PrometKolone(), outA, n, 0#, zbir, Array(0, 0, 0))
    Exit Function
EH:
    Err.Raise Err.Number, "modScrAgro.RedoviPromet[" & mStep & "]", Err.description
End Function

'-------------------------------------------------------- LISTA: DUGOVI
Private Function DugoviKolone() As Variant
    DugoviKolone = Array( _
        "OTKUI_HDA_KOOPERANT||part|0|1", _
        "OTKUI_HDA_ZADUZENJE||rsd|110|1", _
        "OTKUI_HDA_ODBITAK||rsd|110|2", _
        "OTKUI_HDA_DUG||rsd|110|1")
End Function

' GetAgroDugoviForGrid je 0-bazirano:
'   0 KooperantID | 1 Kooperant | 2 Zaduzenje | 3 Odbitak | 4 Dug
Private Function RedoviDugovi(ByVal filter As String, ByVal q As String) As Variant
    Dim src As Variant, r As Long, n As Long, outA() As Variant
    Dim hay As String, zbir As Double, naziv As String, koopID As String
    On Error GoTo EH
    mStep = "dugovi"

    Set mDugIds = CreateObject("Scripting.Dictionary")
    mDugIds.CompareMode = vbTextCompare
    src = GetAgroDugoviForGrid()
    If Not IsArray(src) Then
        RedoviDugovi = PrazanRezultat(DugoviKolone())
        Exit Function
    End If

    ReDim outA(1 To UBound(src, 1) + 1, 1 To 4)
    For r = 0 To UBound(src, 1)
        koopID = Trim$(CStr(src(r, 0)))
        naziv = CStr(src(r, 1))
        If mDugIds.Exists(naziv) Then
            If CStr(mDugIds(naziv)) <> koopID Then mDugIds(naziv) = ""
        Else
            mDugIds(naziv) = koopID
        End If
        If Not AgCipDugovi(filter, AgD(src(r, 4))) Then GoTo Sledeci
        hay = koopID & "|" & naziv
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeci
        End If
        n = n + 1
        outA(n, 1) = naziv
        outA(n, 2) = AgD(src(r, 2))
        outA(n, 3) = AgD(src(r, 3))
        outA(n, 4) = AgD(src(r, 4))
        zbir = zbir + AgD(src(r, 4))
Sledeci:
    Next r

    mStep = "OK"
    RedoviDugovi = Array(DugoviKolone(), outA, n, 0#, zbir, Array(0, 0, 0))
    Exit Function
EH:
    Err.Raise Err.Number, "modScrAgro.RedoviDugovi[" & mStep & "]", Err.description
End Function

'=====================================================================
' CELIJE
'=====================================================================
Private Function AgD(ByVal v As Variant) As Double
    On Error Resume Next
    If IsNumeric(v) Then
        AgD = CDbl(v)
    Else
        AgD = val(Replace(CStr(v), ",", "."))
    End If
End Function

Private Function AgDatum(ByVal v As Variant) As Double
    On Error Resume Next
    If IsDate(v) Then
        AgDatum = Int(CDbl(CDate(v)))
    ElseIf IsNumeric(v) Then
        AgDatum = Int(CDbl(v))
    End If
End Function

'=====================================================================
' TEST SEAM
' Zona se u testu ne crta (forma se ne prikazuje), pa se stanje ekrana ne
' moze procitati iz kontrola. Isti razlog i isti oblik kao Scr_*TestSet u
' modScrStorno -- ukljucujuci kapiju: seam koji MENJA stanje ekrana van
' test-rezima ne radi nista. Scr_KorpaTestReset bi inace, pozvan iz liste
' makroa, tiho bacio operateru neproknjizenu korpu.
'=====================================================================
Public Sub Scr_ListaTestSet(ByVal kljuc As String)
    If Not IsTestMode() Then Exit Sub
    mLista = kljuc
End Sub

Public Sub Scr_RezimTestSet(ByVal rezim As String)
    If Not IsTestMode() Then Exit Sub
    mMod = rezim
End Sub

Public Function Scr_Rezim() As String
    Scr_Rezim = Rezim()
End Function

Public Function Scr_KorpaBroj() As Long
    Scr_KorpaBroj = BrojUKorpi(Korpa())
End Function

Public Function Scr_KorpaZbir() As Double
    Scr_KorpaZbir = modAgroUnos.AgroZbirKorpe(Korpa())
End Function

' Dodaj u korpu izdavanja bez zone. Zona u testu ne postoji, a korpa jeste
' stanje EKRANA (Scr_Brojac je cita), pa se do nje mora nekako doci.
Public Function Scr_KorpaTestDodaj(ByVal artikalID As String, _
                                   ByVal brojPak As Double, _
                                   ByVal parcelaID As String) As String
    Dim fokus As String
    If Not IsTestMode() Then Exit Function
    mMod = AG_IZLAZ
    Scr_KorpaTestDodaj = modAgroUnos.AgroDodajIzlaz(Korpa(), artikalID, _
                                                    brojPak, parcelaID, fokus)
End Function

' Identitet iza prikazanog reda. Prazno znaci DVOSMISLENO ili nepoznato -- i to
' je bas ono sto se meri: dvoklik na dvosmislen red ne sme da bira.
Public Function Scr_DugIdTest(ByVal prikaz As String) As String
    If mDugIds Is Nothing Then Exit Function
    If mDugIds.Exists(prikaz) Then Scr_DugIdTest = CStr(mDugIds(prikaz))
End Function

Public Function Scr_ArtIdTest(ByVal prikaz As String) As String
    If mArtIds Is Nothing Then Exit Function
    If mArtIds.Exists(prikaz) Then Scr_ArtIdTest = CStr(mArtIds(prikaz))
End Function

Public Sub Scr_KorpaTestReset()
    If Not IsTestMode() Then Exit Sub
    Set mKorpaI = modAgroUnos.NovaAgroKorpa()
    Set mKorpaU = modAgroUnos.NovaAgroKorpa()
    OcistiParcele
End Sub
