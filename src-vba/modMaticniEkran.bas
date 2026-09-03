Attribute VB_Name = "modMaticniEkran"
'=====================================================================
' modMaticniEkran - zajednicko telo tri maticna ekrana (korak M1).
'
' Partneri, Proizvodi i cene, Ambalaza i pakovanje se razlikuju SAMO po tome
' koje sekcije nose. Zona, dogadjaji i citanje redova su im isti, pa stoje
' ovde: tri kopije istog rasporeda bi se razisle prvom doradom.
'
' ZASTO OVO NIJE U modMaticniIzvor: taj modul opisuje PODATKE i ne sme da zna
' za kontrole ni za ljusku. Ovaj zna za zonu i za mrezu, a nista o tabelama.
'
' ZONA IMA TRI STANJA:
'   zatvorena  pregled -- naziv sekcije, koliko zapisa, koliko aktivnih;
'   editor     polja sekcije, "Sacuvaj" i "Odustani" (M2b);
'   geo        koordinate izabrane parcele i sest alatki (M3).
' Editor i geo se ISKLJUCUJU: jedna stvar u zoni u isto vreme. Otvaranje jednog
' zatvara drugo -- dva panela jedan preko drugog su ista klasa kvara kao traka
' zOtp koja je ostajala upaljena na tudjem ekranu.
' Obrazac je iz liste "Nova prerada" na ekranu Palete: polja postoje uvek, a
' Scr_Layout ih pali, gasi i rasporedjuje; zona raste samo dok se uredjuje.
'
' POLJA SU BAZEN, ne po sekciji. NewFieldG pri GRADNJI odlucuje da li pravi
' tekst ili combo, a sekcije se razlikuju -- pa se gradi deset tekstualnih i
' sest combo polja (tacno budzet koji frmStammdaten ima kroz txtField1..10 i
' cmbField1..6), a raspored svakoj sekciji dodeljuje sledece slobodno polje
' njene vrste. Isti razlog zbog kog mreza ima MAX_COLS kolona, a ne po ekranu.
'
' UPIS NE ZIVI OVDE. Sve ide u modMaticniUnos -- isti pisac koga zove i legacy
' forma (v. UI_MIGRACIJA_KATALOG 26.5 i 26.15).
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const MATEKR_BUILD As String = "v6-ui-199"

' Visina zone je ista kao KPI traka, pa naslov ispod nje pada u isti red na
' svim ekranima -- isto pravilo koje vec postuju Palete i Oporavak.
Private Const MAT_ZONA_H As Single = KPI_H

' Bazen polja editora -- isti budzet koji legacy forma ima u .frx.
Private Const MAT_MAX_TXT As Long = 10
Private Const MAT_MAX_CMB As Long = 6
Private Const MAT_KOL As Long = 4          ' polja po redu
Private Const MAT_RED_H As Single = 46     ' = FIELD_GRP_H
Private Const MAT_GAP As Single = 10

' Ekran cija je zona poslednja crtana. Tri ekrana dele ovo telo, a dogadjaji
' editora ne dobijaju kljuc ekrana -- pa se pamti pri citanju redova, sto se
' desava pre svakog dogadjaja.
Private mZonaEkran As String

' Ekran koji je OTVORIO editor. Cetiri maticna ekrana dele ovo telo, pa deljeno
' stanje bez vlasnika znaci da editor otvoren nad Kooperantima prezivi prelazak
' na Robu -- a "Sacuvaj" bi tada pisao u tblKooperanti dok operater gleda robu.
' Zato svaka mutacija pita CIJI je editor, i odbija ako nije njen.
Private mEditEkran As String

' Stanje editora. Prazan mEditKljuc znaci da je zona zatvorena; prazan mEditID
' uz otvoren editor znaci NOV zapis.
Private mEditKljuc As String
Private mEditID As String
Private mIzabranID As String               ' PK reda izabranog u mrezi

' Parcela ciji je GEO panel otvoren. Prazno = panel je zatvoren.
Private mGeoID As String

' Korisnik ciju matricu prava lista PRAVA pokazuje. Pamti se ODVOJENO od
' mIzabranID zato sto prezivljava prelazak na drugu listu: cim se na listi PRAVA
' izabere red, mIzabranID postaje oblast, a korisnik mora da ostane.
Private mKorisnikID As String

' Straza od povratnog udara pri punjenju zavisnih combo-a. FillCmb i
' PostaviPolje menjaju kontrolu, a to opet okida Change -- koji stize nazad kao
' "chg:". Kaskada je danas plitka (sorta nema svoje zavisne) pa bi se sama
' zaustavila, ali straza je uslov da druga zavisnost sutra ne postane petlja.
Private mUZavisnima As Boolean

' Visina zone sa otvorenim GEO panelom: jedan red polja (N i E) pa dugmad.
Private Const MAT_GEO_H As Single = MAT_ZONA_H + 18 + MAT_RED_H + 38

' Visina jednog reda alatki GEO panela. Traka se prelama kad ne staje u sirinu,
' pa zona raste za po ovoliko po prelomljenom redu.
Private Const GEO_RED_H As Single = 32

'--------------------------------------------------------------- ZONA
Public Sub ZonaGradi(ByVal z As Object)
    Dim i As Long
    modUiKit.NewLbl z, "matCap", "", PAD, 6, 300, 11, TS_MICRO, True, C_MUTED, -1
    modUiKit.NewLbl z, "matBroj", ChrW(8212), PAD, 18, 300, 20, TS_KPI, True, C_FOREST, -1
    ' Napomena da se unos jos radi kroz staru formu. Stoji u zoni, ne u dijalogu:
    ' dijalog se zatvori i zaboravi, a operater koji ovde trazi dugme "Dodaj"
    ' mora da vidi gde ono jeste.
    modUiKit.NewLbl z, "matHint", Poruka("OTKUI_MAT_UNOS_LEGACY"), PAD, 40, 520, 13, _
                    TS_META, False, C_MUTED, -1

    For i = 0 To 1
        modUiKit.NewLbl z, "matKL" & i, "", 0, 6, 120, 12, TS_MICRO, True, C_MUTED, -1
        modUiKit.NewLbl z, "matKV" & i, ChrW(8212), 0, 18, 120, 20, TS_KPI, True, _
                        C_FOREST, -1, fmTextAlignLeft, F_NUM
    Next i
    ' --- EDITOR ---------------------------------------------------------
    ' Bela podloga MORA biti labela, ne Frame: Frame je prozorska kontrola i
    ' crta se iznad bezprozorskih bez obzira na z-order (nauceno na panelu
    ' prerade). Napravljena PRVA, ostaje ispod svega.
    modUiKit.NewLbl z, "matEdBg", "", 0, 0, 100, 10, 8, False, 0, C_WHITE
    modUiKit.NewLbl z, "matEdCap", "", PAD, 0, 420, 12, TS_MICRO, True, C_MUTED, -1
    For i = 0 To MAT_MAX_TXT - 1
        modOtkupUI.NewFieldG z, "scrMatT" & i, "", "txt", "", 1, False, False, "MAT"
    Next i
    For i = 0 To MAT_MAX_CMB - 1
        modOtkupUI.NewFieldG z, "scrMatC" & i, "", "cmb", "", 1, False, False, "MAT"
    Next i
    modUiKit.BtnV z, "scrMatSacuvaj", Poruka("OTKUI_BTN_MAT_SACUVAJ"), 0, 0, 120, 26, "primary"
    modUiKit.BtnV z, "scrMatOdustani", Poruka("OTKUI_BTN_MAT_ODUSTANI"), 0, 0, 104, 26, "soft"
    modUiKit.BtnV z, "scrMatNovi", Poruka("OTKUI_BTN_MAT_NOVI"), 0, 0, 130, 26, "primary"

    ' --- GEO PANEL (samo Parcele) ---------------------------------------
    ' Sest alatki koje legacy forma ima kao sest dugmadi uz listu parcela.
    ' Posao im je u modMaticniGeo; ovde su samo kontrole.
    modUiKit.NewLbl z, "matGeoCap", UCase$(Poruka("OTKUI_MATG_CAP")), PAD, 0, 200, 12, _
                    TS_MICRO, True, C_MUTED, -1
    modUiKit.NewLbl z, "matGeoOpis", "", PAD, 0, 520, 13, TS_META, False, C_FOREST, -1
    modOtkupUI.NewFieldG z, "scrGeoN", Poruka("OTKUI_MATG_N"), "txt", "", 1, True, False, "GEO"
    modOtkupUI.NewFieldG z, "scrGeoE", Poruka("OTKUI_MATG_E"), "txt", "", 1, True, False, "GEO"
    modUiKit.BtnV z, "scrGeoSacuvaj", Poruka("OTKUI_BTN_GEO_SACUVAJ"), 0, 0, 108, 26, "primary"
    modUiKit.BtnV z, "scrGeoNalepi", Poruka("OTKUI_BTN_GEO_NALEPI"), 0, 0, 138, 26, "soft"
    modUiKit.BtnV z, "scrGeoPortal", Poruka("OTKUI_BTN_GEO_PORTAL"), 0, 0, 96, 26, "soft"
    modUiKit.BtnV z, "scrGeoMape", Poruka("OTKUI_BTN_GEO_MAPE"), 0, 0, 110, 26, "soft"
    modUiKit.BtnV z, "scrGeoPoligon", Poruka("OTKUI_BTN_GEO_POLIGON"), 0, 0, 88, 26, "soft"
    modUiKit.BtnV z, "scrGeoObrisi", Poruka("OTKUI_BTN_GEO_OBRISI"), 0, 0, 100, 26, "danger"
    modUiKit.BtnV z, "scrGeoZatvori", Poruka("OTKUI_BTN_GEO_ZATVORI"), 0, 0, 88, 26, "ghost"

    modUiKit.NewLbl z, "matLnB", "", 0, MAT_ZONA_H - 1, 100, 1, 8, False, 0, C_BORDER
End Sub

' Koliko redova polja sekcija trazi, i kolika je zona kad je editor otvoren.
Private Function BrojRedovaPolja(ByVal lista As String) As Long
    Dim a As Variant
    a = modMaticniIzvor.MatPolja(lista)
    If Not IsArray(a) Then Exit Function
    BrojRedovaPolja = -Int(-(UBound(a) + 1) / MAT_KOL)
End Function

Private Function VisinaOtvorene(ByVal lista As String) As Single
    VisinaOtvorene = MAT_ZONA_H + 16 + BrojRedovaPolja(lista) * MAT_RED_H + 36
End Function

Public Function ZonaRaspored(ByVal z As Object, ByVal w As Single, _
                             ByVal lista As String) As Single
    Dim i As Long, h As Single
    On Error Resume Next
    For i = 0 To 1
        z.Controls("matKL" & i).Left = w - PAD - (2 - i) * 150
        z.Controls("matKV" & i).Left = w - PAD - (2 - i) * 150
    Next i
    ' Napomena deli red sa dve plocice desno; na uskom prozoru se skloni umesto
    ' da dobije negativnu sirinu.
    z.Controls("matHint").Visible = (w - 2 * PAD - 320 > 120) And _
                                    Len(mEditKljuc) = 0 And Len(mGeoID) = 0
    If w - 2 * PAD - 320 > 120 Then z.Controls("matHint").width = w - 2 * PAD - 320

    If Len(mGeoID) > 0 Then
        SakrijEditor z
        modUiKit.BoxShow z, "scrMatNovi", False
        h = RasporediGeo(z, w)
    ElseIf Len(mEditKljuc) = 0 Then
        SakrijEditor z
        SakrijGeo z
        ' "Nova stavka" postoji i kad je editor zatvoren -- to mu je jedini ulaz.
        modUiKit.MoveBox z, "scrMatNovi", w - PAD - 130, 14, 130
        ' Sekcija bez polja nema ni unos: prava se ne dodaju, nego ukljucuju.
        ' Uslov se cita iz opisa polja, ne iz spiska sekcija -- da nova sekcija
        ' bez editora ne bi dobila dugme koje ne vodi nikuda.
        modUiKit.BoxShow z, "scrMatNovi", _
            (Len(lista) > 0) And IsArray(modMaticniIzvor.MatPolja(lista))
        h = MAT_ZONA_H
    Else
        SakrijGeo z
        modUiKit.BoxShow z, "scrMatNovi", False
        h = RasporediEditor(z, w)
    End If

    z.Controls("matLnB").top = h - 1
    z.Controls("matLnB").width = w
    ZonaRaspored = h
End Function

Private Sub SakrijEditor(ByVal z As Object)
    On Error Resume Next
    PoljaEditora z, "", False
    modUiKit.BoxShow z, "scrMatSacuvaj", False
    modUiKit.BoxShow z, "scrMatOdustani", False
    z.Controls("matEdBg").Visible = False
    z.Controls("matEdCap").Visible = False
End Sub

Private Sub SakrijGeo(ByVal z As Object)
    Dim nm As Variant
    On Error Resume Next
    z.Controls("matGeoCap").Visible = False
    z.Controls("matGeoOpis").Visible = False
    z.Controls("scrGeoN").Visible = False
    z.Controls("scrGeoE").Visible = False
    For Each nm In GeoDugmad()
        modUiKit.BoxShow z, CStr(nm), False
    Next nm
End Sub

Private Function GeoDugmad() As Variant
    GeoDugmad = Array("scrGeoSacuvaj", "scrGeoNalepi", "scrGeoPortal", _
                      "scrGeoMape", "scrGeoPoligon", "scrGeoObrisi", "scrGeoZatvori")
End Function

' Raspored GEO panela: dva polja levo, sest alatki i "Zatvori" u redu ispod.
Private Function RasporediGeo(ByVal z As Object, ByVal w As Single) As Single
    Dim kol As Single, y0 As Single, x As Single, nm As Variant, i As Long

    z.Controls("matEdBg").Visible = True
    z.Controls("matEdBg").Left = PAD - 10
    z.Controls("matEdBg").top = MAT_ZONA_H
    z.Controls("matEdBg").width = w - 2 * (PAD - 10)
    z.Controls("matEdBg").Height = MAT_GEO_H - MAT_ZONA_H - 1

    z.Controls("matGeoCap").Visible = True
    z.Controls("matGeoCap").top = MAT_ZONA_H + 4
    z.Controls("matGeoOpis").Visible = True
    z.Controls("matGeoOpis").top = MAT_ZONA_H + 4
    z.Controls("matGeoOpis").Left = PAD + 110
    z.Controls("matGeoOpis").width = w - PAD - 120
    z.Controls("matGeoOpis").caption = modMaticniGeo.GeoOpis(mGeoID)

    kol = (w - 2 * PAD - MAT_GAP) / 4
    If kol < 140 Then kol = 140
    y0 = MAT_ZONA_H + 20
    z.Controls("scrGeoN").Visible = True
    z.Controls("scrGeoN").Left = PAD
    z.Controls("scrGeoN").top = y0
    z.Controls("scrGeoN").width = kol
    modOtkupUI.LayoutFieldInner z.Controls("scrGeoN")
    z.Controls("scrGeoE").Visible = True
    z.Controls("scrGeoE").Left = PAD + kol + MAT_GAP
    z.Controls("scrGeoE").top = y0
    z.Controls("scrGeoE").width = kol
    modOtkupUI.LayoutFieldInner z.Controls("scrGeoE")

    ' Dugmad idu DESNO od polja kad ima mesta, inace u red ispod njih.
    '
    ' Sedam alatki trazi 764pt. Prag 640 je bio PROCENA, i na prozoru od 900pt
    ' (radna povrsina ~690) je poslednje dugme izlazilo van zone -- bez ijedne
    ' greske, samo odseceno. Sada se meri STVARNA sirina i red se prelama kad
    ' sledece dugme ne staje, a visina zone se racuna iz broja redova umesto da
    ' bude konstanta.
    x = PAD + 2 * (kol + MAT_GAP)
    If x + SirinaGeoTrake() > w - PAD Then
        x = PAD
        y0 = y0 + MAT_RED_H
    Else
        y0 = y0 + 8
    End If

    Dim redova As Long, xPoc As Single, bw As Single
    redova = 1
    xPoc = x
    For Each nm In GeoDugmad()
        bw = SirinaGeoDugmeta(CStr(nm))
        If x > xPoc And x + bw > w - PAD Then
            x = xPoc
            y0 = y0 + GEO_RED_H
            redova = redova + 1
        End If
        modUiKit.MoveBox z, CStr(nm), x, y0, bw
        modUiKit.BoxShow z, CStr(nm), True
        x = x + bw + 6
        i = i + 1
    Next nm

    RasporediGeo = MAT_GEO_H + (redova - 1) * GEO_RED_H
End Function

' Ukupna sirina trake alatki, sa razmacima. Racuna se iz istog spiska iz kog se
' i crta -- broj u konstanti bi se razisao prvim dodatim dugmetom.
Private Function SirinaGeoTrake() As Single
    Dim nm As Variant, uk As Single
    For Each nm In GeoDugmad()
        uk = uk + SirinaGeoDugmeta(CStr(nm)) + 6
    Next nm
    SirinaGeoTrake = uk - 6
End Function

Private Function SirinaGeoDugmeta(ByVal nm As String) As Single
    Select Case nm
        Case "scrGeoSacuvaj":  SirinaGeoDugmeta = 108
        Case "scrGeoNalepi":   SirinaGeoDugmeta = 138
        Case "scrGeoPortal":   SirinaGeoDugmeta = 96
        Case "scrGeoMape":     SirinaGeoDugmeta = 110
        Case "scrGeoPoligon":  SirinaGeoDugmeta = 88
        Case "scrGeoObrisi":   SirinaGeoDugmeta = 100
        Case Else:             SirinaGeoDugmeta = 88
    End Select
End Function

' Raspored otvorenog editora: polja u redovima po cetiri, pa dugmad ispod.
Private Function RasporediEditor(ByVal z As Object, ByVal w As Single) As Single
    Dim a As Variant, i As Long, n As Long, kol As Single, y0 As Single, h As Single
    Dim nmT As Long, nmC As Long, nm As String, sp As String

    a = modMaticniIzvor.MatPolja(mEditKljuc)
    If Not IsArray(a) Then
        RasporediEditor = MAT_ZONA_H
        Exit Function
    End If
    n = UBound(a) + 1
    h = VisinaOtvorene(mEditKljuc)

    z.Controls("matEdBg").Visible = True
    z.Controls("matEdBg").Left = PAD - 10
    z.Controls("matEdBg").top = MAT_ZONA_H
    z.Controls("matEdBg").width = w - 2 * (PAD - 10)
    z.Controls("matEdBg").Height = h - MAT_ZONA_H - 1

    z.Controls("matEdCap").Visible = True
    z.Controls("matEdCap").top = MAT_ZONA_H + 4
    z.Controls("matEdCap").caption = UCase$(NatpisEditora())

    kol = (w - 2 * PAD - (MAT_KOL - 1) * MAT_GAP) / MAT_KOL
    If kol < 120 Then kol = 120
    y0 = MAT_ZONA_H + 18

    PoljaEditora z, mEditKljuc, True
    For i = 0 To n - 1
        sp = CStr(a(LBound(a) + i))
        If modMaticniIzvor.PoljeF(sp, 2) = "cmb" Then
            nm = "scrMatC" & nmC: nmC = nmC + 1
        Else
            nm = "scrMatT" & nmT: nmT = nmT + 1
        End If
        z.Controls(nm).Left = PAD + (i Mod MAT_KOL) * (kol + MAT_GAP)
        z.Controls(nm).top = y0 + (i \ MAT_KOL) * MAT_RED_H
        z.Controls(nm).width = kol
        ' Bez ovoga unutrasnje kontrole ostaju na merama iz gradnje (180pt).
        modOtkupUI.LayoutFieldInner z.Controls(nm)
    Next i

    modUiKit.MoveBox z, "scrMatSacuvaj", PAD, h - 32, 120
    modUiKit.MoveBox z, "scrMatOdustani", PAD + 128, h - 32, 104
    modUiKit.BoxShow z, "scrMatSacuvaj", True
    modUiKit.BoxShow z, "scrMatOdustani", True
    RasporediEditor = h
End Function

' Pali polja koja sekcija koristi, gasi ostatak bazena, i postavlja natpise.
' Polje koje ostane upaljeno iz prethodne sekcije bi trazilo vrednost koju
' pisac ne ocekuje -- otud gasenje CELOG bazena pri svakoj promeni.
Private Sub PoljaEditora(ByVal z As Object, ByVal lista As String, ByVal vis As Boolean)
    Dim a As Variant, i As Long, nmT As Long, nmC As Long, nm As String, sp As String
    On Error Resume Next
    For i = 0 To MAT_MAX_TXT - 1
        z.Controls("scrMatT" & i).Visible = False
    Next i
    For i = 0 To MAT_MAX_CMB - 1
        z.Controls("scrMatC" & i).Visible = False
    Next i
    If Not vis Or Len(lista) = 0 Then Exit Sub

    a = modMaticniIzvor.MatPolja(lista)
    If Not IsArray(a) Then Exit Sub
    For i = LBound(a) To UBound(a)
        sp = CStr(a(i))
        If modMaticniIzvor.PoljeF(sp, 2) = "cmb" Then
            nm = "scrMatC" & nmC: nmC = nmC + 1
        Else
            nm = "scrMatT" & nmT: nmT = nmT + 1
        End If
        z.Controls(nm).Visible = True
        z.Controls(nm).Controls(nm & "L").caption = _
            UCase$(Poruka(modMaticniIzvor.PoljeF(sp, 1)))
    Next i
End Sub

' Ime kontrole bazena koja nosi dato polje AKTIVNE sekcije. Racuna se iz opisa,
' istim redosledom kao raspored -- pa se dodela ne moze razici sa prikazom.
Private Function KontrolaPolja(ByVal lista As String, ByVal poljeKljuc As String) As String
    Dim a As Variant, i As Long, nmT As Long, nmC As Long, sp As String
    a = modMaticniIzvor.MatPolja(lista)
    If Not IsArray(a) Then Exit Function
    For i = LBound(a) To UBound(a)
        sp = CStr(a(i))
        If modMaticniIzvor.PoljeF(sp, 2) = "cmb" Then
            If modMaticniIzvor.PoljeF(sp, 0) = poljeKljuc Then
                KontrolaPolja = "scrMatC" & nmC
                Exit Function
            End If
            nmC = nmC + 1
        Else
            If modMaticniIzvor.PoljeF(sp, 0) = poljeKljuc Then
                KontrolaPolja = "scrMatT" & nmT
                Exit Function
            End If
            nmT = nmT + 1
        End If
    Next i
End Function

Private Function NatpisEditora() As String
    If Len(mEditID) = 0 Then
        NatpisEditora = Poruka("OTKUI_MAT_NOV_ZAPIS")
    Else
        NatpisEditora = Poruka("OTKUI_MAT_IZMENA") & " " & mEditID
    End If
End Function

' Brojke se pisu POSLE citanja redova, iz istog prolaza -- pa se broj u zoni i
' lista u mrezi ne mogu razici. Sekcija bez kolone statusa nema sta da razlozi
' na aktivne i neaktivne, pa te dve plocice ostaju prazne (em-crta), a ne nula:
' nula bi tvrdila da neaktivnih nema, a odgovor je da pojam ne postoji.
Private Sub OsveziZonu(ByVal ekran As String, ByVal lista As String)
    Dim z As Object, imaStatus As Boolean
    On Error Resume Next
    Set z = modOtkupUI.ScreenZone(ekran)
    If z Is Nothing Then Exit Sub
    imaStatus = (Len(modMaticniIzvor.MatStatusKolona(lista)) > 0)

    z.Controls("matCap").caption = UCase$(Poruka(NaslovListe(ekran, lista)))
    z.Controls("matBroj").caption = CStr(modMaticniIzvor.MatUkupno()) & " " & _
                                    Poruka("OTKUI_MAT_ZAPISA")
    z.Controls("matKL0").caption = UCase$(Poruka("OTKUI_MAT_AKTIVNIH"))
    z.Controls("matKL1").caption = UCase$(Poruka("OTKUI_MAT_NEAKTIVNIH"))
    ' Napomena kaze CIJA su prava na ekranu. Bez toga je lista od dvanaest
    ' oblasti bez vlasnika -- a promasen korisnik je tiha greska koja se vidi
    ' tek kad se neko ne prijavi.
    z.Controls("matHint").caption = Napomena(lista)
    If lista = "PRAVA" Then
        ' Prava nemaju "aktivne" i "neaktivne" zapise nego oblasti sa pravom i
        ' bez njega -- ista brojka, drugo ime.
        z.Controls("matKL0").caption = UCase$(Poruka("OTKUI_KOR_IMA"))
        z.Controls("matKL1").caption = UCase$(Poruka("OTKUI_KOR_NEMA"))
        z.Controls("matKV0").caption = CStr(modMaticniIzvor.MatAktivnih())
        z.Controls("matKV1").caption = CStr(modMaticniIzvor.MatNeaktivnih())
        Exit Sub
    End If
    If imaStatus Then
        z.Controls("matKV0").caption = CStr(modMaticniIzvor.MatAktivnih())
        z.Controls("matKV1").caption = CStr(modMaticniIzvor.MatNeaktivnih())
    Else
        z.Controls("matKV0").caption = ChrW(8212)
        z.Controls("matKV1").caption = ChrW(8212)
    End If
End Sub

' Napomena u zoni. Za prava kaze cija su, za ostalo gde je unos.
Private Function Napomena(ByVal lista As String) As String
    If lista <> "PRAVA" Then
        Napomena = Poruka("OTKUI_MAT_UNOS_LEGACY")
    ElseIf Len(mKorisnikID) = 0 Then
        Napomena = Poruka("OTKUI_KOR_BEZ_IZBORA")
    Else
        Napomena = Poruka("OTKUI_KOR_PRAVA_ZA") & " " & _
                   modMaticniKorisnici.KorNaziv(mKorisnikID)
    End If
End Function

' Naslov aktivne liste iz ISTOG spiska koji puni prekidac -- da se natpis u
' zoni i natpis na dugmetu ne mogu razici.
Private Function NaslovListe(ByVal ekran As String, ByVal lista As String) As String
    Dim a As Variant, r As Variant, p() As String
    a = modMaticniIzvor.MatSekcijeEkrana(ekran)
    If Not IsArray(a) Then Exit Function
    For Each r In a
        p = Split(CStr(r), "|")
        If p(0) = lista Then
            NaslovListe = p(2)
            Exit Function
        End If
    Next r
End Function

'-------------------------------------------------------------- REDOVI
Public Function Redovi(ByVal ekran As String, ByVal lista As String, _
                       ByVal filter As String, ByVal q As String) As Variant
    mZonaEkran = ekran
    ' Kontekst nosi izabranog korisnika -- treba samo listi PRAVA, i samo ona
    ' ga cita. Izbor je stanje EKRANA, pa se prosledjuje, a ne pamti u izvoru.
    Redovi = modMaticniIzvor.MatRedovi(lista, filter, q, mKorisnikID)
    ' Zona se osvezava POSLE citanja, iz istog prolaza -- brojke i lista se tako
    ' ne mogu razici. Ekran se prosledjuje jer tri ekrana dele ovo telo, pa se
    ' mora znati CIJU zonu ljuska treba da vrati.
    OsveziZonu ekran, lista
End Function

'------------------------------------------------------------- RADNJE
' Radnje nad redom za AKTIVNU listu.
'
' Cenovnik nema "Izmeni": append-only, nova cena je nov red -- isto pravilo po
' kom legacy forma za Cenovnik krije to dugme. "Deaktiviraj" postoji samo tamo
' gde sekcija STVARNO ima kolonu statusa; inace bi dugme tiho ne radilo nista.
Public Function Radnje(ByVal lista As String) As String
    Dim s As String
    If Len(lista) = 0 Then Exit Function
    ' Prava se ne uredjuju u editoru: red je oblast, ne zapis, i jedina radnja
    ' je da se ukljuci ili iskljuci. Ni "Izmeni" ni "Deaktiviraj" ovde nemaju
    ' sta da urade -- oba bi trazila zapis koji ne postoji.
    If lista = "PRAVA" Then
        Radnje = "pravo:OTKUI_BTN_KOR_PRAVO:150:soft:1"
        Exit Function
    End If
    If lista <> "CENOVNIK" Then s = "izmeni:OTKUI_BTN_MAT_IZMENI:92:soft:1"
    If Len(modMaticniIzvor.MatStatusKolona(lista)) > 0 Then
        If Len(s) > 0 Then s = s & "|"
        s = s & "status:OTKUI_BTN_MAT_STATUS:150:danger:1"
    End If
    ' GEO ima SAMO Parcele -- jedina sekcija koja nosi koordinate. Dugme na
    ' ostalima bi otvaralo panel koji nema sta da pokaze.
    If lista = "PARCELE" Then
        If Len(s) > 0 Then s = s & "|"
        s = s & "geo:OTKUI_BTN_MAT_GEO:70:soft:1"
    End If
    Radnje = s
End Function

'------------------------------------------------------------ DOGADJAJI
' Prekidac lista, izbor reda, radnje nad redom i dugmad editora.
'
' Lista se prima ByRef zato sto je stanje EKRANA, ne ovog modula: tri ekrana
' dele telo, ali svaki pamti svoju aktivnu listu.
'
' Vraca True SAMO kad se lista mora ponovo procitati. Izbor reda vraca False --
' operater bi inace gubio mesto u listi pri svakom kliku.
Public Function Dogadjaj(ByVal tag As String, ByRef lista As String) As Boolean
    Dim preEdit As String
    preEdit = mEditKljuc & "/" & mEditID & "/" & mGeoID
    Dogadjaj = ObradiDogadjaj(tag, lista)
    ' Otvaranje ili zatvaranje editora menja VISINU zone. Ljuska to ne moze da
    ' pogodi -- LayoutScreenZone ide samo iz LayoutOtkup -- pa se trazi izricito.
    If (mEditKljuc & "/" & mEditID & "/" & mGeoID) <> preEdit Then _
        modOtkupUI.OsveziRasporedEkrana
End Function

Private Function ObradiDogadjaj(ByVal tag As String, ByRef lista As String) As Boolean
    If Left$(tag, 2) = "ls" Then
        If Mid$(tag, 3) = lista Then Exit Function
        ' Odlazak sa liste zatvara editor: polja pripadaju sekciji koja se
        ' napusta, a ostavljena bi trazila upis u drugu tabelu. Zato se PITA --
        ' do sada je otvoren unos tiho nestajao, pa se otkucano gubilo bez reci.
        If Len(mEditKljuc) > 0 Then
            If MsgBox(Poruka("MATU_ASK_ODBACI_UNOS"), _
                      vbExclamation + vbYesNo + vbDefaultButton2, APP_NAME) <> vbYes Then Exit Function
        End If
        ZatvoriPanele
        lista = Mid$(tag, 3)
        ObradiDogadjaj = True
        Exit Function
    End If

    If Left$(tag, 4) = "row:" Then
        ZapamtiIzbor lista, CLng(val(Mid$(tag, 5)))
        ' Izbor korisnika menja SADRZAJ liste prava. Dok se gleda lista
        ' korisnika mreza se ne cita ponovo (operater bi gubio mesto), ali cim
        ' se predje na prava, prekidac lista svakako trazi novo citanje.
        Exit Function
    End If

    ' Dvoklik otvara izmenu -- jedan potez umesto dva (izaberi red, pa dugme).
    If Left$(tag, 4) = "dbl:" Then
        ZapamtiIzbor lista, CLng(val(Mid$(tag, 5)))
        If lista = "PRAVA" Then
            ObradiDogadjaj = PromeniPravo()
        Else
            OtvoriIzmenu lista
        End If
        Exit Function
    End If

    If Left$(tag, 4) = "act:" Then
        ObradiDogadjaj = Akcija(tag, lista)
        Exit Function
    End If

    ' Promena teksta u polju editora. Stize kao "chg:<ime kontrole>" iz
    ' modOtkupUI.UiChange -- ozicavanje postoji od pocetka (NewFieldG ->
    ' WireInput), ali ga ekran nije koristio, pa kaskada sorte nije radila.
    ' Vraca False: kucanje ne menja podatke, mreza nema sta da cita ponovo.
    If Left$(tag, 4) = "chg:" Then
        If Len(mEditKljuc) > 0 Then _
            OsveziZavisne mEditKljuc, PoljeIzKontrole(mEditKljuc, Mid$(tag, 5))
        Exit Function
    End If

    ' Otvaranje i zatvaranje editora NE menjaju podatke, pa vracaju False:
    ' ljuska bi inace na svaki klik radila pun RefreshFromData (KPI, kes,
    ' combo-i) za cistu promenu prikaza. Raspored se trazi izricito, gore.
    Select Case tag
        Case "scrMatNovi":     OtvoriNov lista
        Case "scrMatOdustani": ZatvoriEditor
        Case "scrMatSacuvaj":  ObradiDogadjaj = Sacuvaj(lista)
        Case "scrGeoZatvori":  mGeoID = ""
        Case "scrGeoNalepi":   GeoNalepi
        Case "scrGeoPortal":   GeoJavi modMaticniGeo.GeoOtvoriPortal(mGeoID), "MATG_OK_PORTAL"
        Case "scrGeoMape":     GeoJavi modMaticniGeo.GeoOtvoriMape(mGeoID), ""
        Case "scrGeoPoligon":  GeoJavi modMaticniGeo.GeoOtvoriPoligon(mGeoID), ""
        Case "scrGeoSacuvaj":  ObradiDogadjaj = GeoSacuvajKlik()
        Case "scrGeoObrisi":   ObradiDogadjaj = GeoObrisiKlik()
    End Select
End Function

'---------------------------------------------------------------- GEO
' Panel se otvara nad IZABRANIM redom i zatvara editor -- jedna stvar u zoni.
Private Sub OtvoriGeo()
    Dim nTxt As String, eTxt As String
    If Len(mIzabranID) = 0 Then
        modOtkupUI.ShowToast Poruka("MATG_ERR_NEMA_PARCELE"), True
        Exit Sub
    End If
    ZatvoriEditor
    mGeoID = mIzabranID
    modMaticniGeo.GeoKoordinate mGeoID, nTxt, eTxt
    PostaviGeoPolje "scrGeoN", nTxt
    PostaviGeoPolje "scrGeoE", eTxt
End Sub

' Zajednicki izlaz za alatke koje samo otvaraju adresu: poruka o gresci ide u
' crveni toast, uspeh u obican -- i to samo tamo gde uspeh ima sta da kaze
' (portal je kopirao pretragu; mape i poligon se vide same).
Private Sub GeoJavi(ByVal odgovor As String, ByVal kljucOK As String)
    If Len(odgovor) > 0 Then
        modOtkupUI.ShowToast odgovor, True
    ElseIf Len(kljucOK) > 0 Then
        modOtkupUI.ShowToast Poruka(kljucOK), False
    End If
End Sub

Private Sub GeoNalepi()
    Dim txt As String, n As Double, e As Double
    txt = Trim$(GetClipboardText())
    If Not modMaticniGeo.GeoIzTeksta(txt, n, e) Then
        modOtkupUI.ShowToast Poruka("MATG_ERR_NALEPLJENO"), True
        Exit Sub
    End If
    ' Format bez eksponenta i bez hiljadica -- UTM je sedmocifren broj.
    PostaviGeoPolje "scrGeoN", Format$(n, "0.##")
    PostaviGeoPolje "scrGeoE", Format$(e, "0.##")
    modOtkupUI.ShowToast Poruka("MATG_OK_NALEPLJENO"), False
End Sub

Private Function GeoSacuvajKlik() As Boolean
    Dim odgovor As String, fokus As String
    odgovor = modMaticniGeo.GeoSacuvaj(mGeoID, CitajGeoPolje("scrGeoN"), _
                                       CitajGeoPolje("scrGeoE"), fokus)
    If Len(odgovor) > 0 Then
        modOtkupUI.ShowToast odgovor, True
        If fokus = "e" Then FokusGeo "scrGeoE" Else FokusGeo "scrGeoN"
        Exit Function
    End If
    modOtkupUI.ShowToast Poruka("MATG_OK_SACUVANO"), False
    GeoSacuvajKlik = True
End Function

' Brisanje trazi potvrdu: tacka i poligon se gube, a poligon se rucno crta.
' Legacy je to resavao dugmetom u dva koraka; potvrda je ista brana, samo
' citljivija.
Private Function GeoObrisiKlik() As Boolean
    Dim odgovor As String
    If Len(mGeoID) = 0 Then Exit Function
    If MsgBox(Poruka("MATG_ASK_OBRISI") & vbCrLf & vbCrLf & mGeoID, _
              vbExclamation + vbYesNo + vbDefaultButton2, APP_NAME) <> vbYes Then Exit Function
    odgovor = modMaticniGeo.GeoObrisi(mGeoID)
    If Len(odgovor) > 0 Then
        modOtkupUI.ShowToast odgovor, True
        Exit Function
    End If
    PostaviGeoPolje "scrGeoN", ""
    PostaviGeoPolje "scrGeoE", ""
    modOtkupUI.ShowToast Poruka("MATG_OK_OBRISANO"), False
    GeoObrisiKlik = True
End Function

Private Function CitajGeoPolje(ByVal nm As String) As String
    Dim z As Object
    On Error Resume Next
    Set z = Zona()
    If z Is Nothing Then Exit Function
    CitajGeoPolje = Trim$(NzToText(z.Controls(nm).Controls(nm & "T").value))
    Err.Clear
End Function

Private Sub PostaviGeoPolje(ByVal nm As String, ByVal v As String)
    Dim z As Object
    On Error Resume Next
    Set z = Zona()
    If z Is Nothing Then Exit Sub
    z.Controls(nm).Controls(nm & "T").value = v
    Err.Clear
End Sub

Private Sub FokusGeo(ByVal nm As String)
    Dim z As Object
    On Error Resume Next
    Set z = Zona()
    If z Is Nothing Then Exit Sub
    z.Controls(nm).Controls(nm & "T").SetFocus
    Err.Clear
End Sub

Private Function Akcija(ByVal tag As String, ByVal lista As String) As Boolean
    Dim p() As String, red As Long
    p = Split(Mid$(tag, 5), ":")
    If UBound(p) < 1 Then
        LogWarn "modMaticniEkran.Akcija", "Radnja bez rednog broja: '" & tag & "'."
        Exit Function
    End If
    ' Identitet reda dolazi iz KOLONE KOJU SEKCIJA PRIJAVLJUJE, ne iz rednog
    ' broja -- posle sortiranja i pretrage pozicija ne znaci nista.
    ZapamtiIzbor lista, CLng(val(p(1)))
    Select Case p(0)
        Case "izmeni": OtvoriIzmenu lista
        Case "status": Akcija = PromeniStatus(lista)
        Case "geo":    OtvoriGeo
        Case "pravo":  Akcija = PromeniPravo()
        Case Else
            LogWarn "modMaticniEkran.Akcija", "Nepoznata radnja '" & p(0) & "'."
            modOtkupUI.ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & p(0), True
    End Select
End Function

'------------------------------------------------------------- EDITOR
Private Sub OtvoriNov(ByVal lista As String)
    If Len(lista) = 0 Then Exit Sub
    mGeoID = ""                      ' jedna stvar u zoni u isto vreme
    mEditKljuc = lista
    mEditID = ""
    mEditEkran = mZonaEkran
    OcistiPolja lista
End Sub

' Ne vraca nista: otvaranje editora ne menja podatke, pa mreza nema sta da
' procita ponovo. Raspored zone trazi omotnica Dogadjaj, po promeni stanja.
'
' CENOVNIK je poseban: istorija cena se ne menja, pa se otvara NOV unos sa
' vrednostima izabranog reda (proizvod ostaje, cena se unosi nova). To je ono
' sto operater i hoce kad klikne na staru cenu -- a legacy forma je za isto
' krila dugme "Izmeni" i trazila da se sve otkuca ponovo.
Private Sub OtvoriIzmenu(ByVal lista As String)
    Dim red As Long, v As Object, a As Variant, r As Variant, nm As String
    Dim jeCenovnik As Boolean

    If Len(lista) = 0 Then Exit Sub
    If Len(mIzabranID) = 0 Then
        modOtkupUI.ShowToast Poruka("MATU_ERR_NEMA_REDA"), True
        Exit Sub
    End If

    red = modMaticniUnos.MatRedPoID(lista, mIzabranID)
    If red = 0 Then
        modOtkupUI.ShowToast Poruka("MATU_ERR_NEMA_REDA"), True
        Exit Sub
    End If

    jeCenovnik = (lista = "CENOVNIK")
    mGeoID = ""                      ' jedna stvar u zoni u isto vreme
    mEditKljuc = lista
    mEditEkran = mZonaEkran
    ' Prazan mEditID znaci UNOS -- otud cenovnik uvek dodaje nov red.
    mEditID = IIf(jeCenovnik, "", mIzabranID)
    OcistiPolja lista

    Set v = modMaticniUnos.MatVrednostiReda(lista, red)
    a = modMaticniIzvor.MatPolja(lista)
    If IsArray(a) Then
        For Each r In a
            nm = modMaticniIzvor.PoljeF(CStr(r), 0)
            If v.Exists(nm) Then PostaviPolje lista, nm, CStr(v(nm))
        Next r
    End If

    ' Zavisni combo-i se pune TEK SADA: pri otvaranju roditelj jos nije imao
    ' vrednost, pa je spisak bio prazan (sorte zavise od izabrane vrste).
    OsveziZavisne lista, ""

    If jeCenovnik Then
        ' Cena se unosi NOVA; datum ostaje prazan pa pisac uzima danasnji.
        PostaviPolje lista, "cena", ""
        PostaviPolje lista, "datum", ""
        modOtkupUI.ShowToast Poruka("MATU_ERR_CENOVNIK_APPEND"), False
    End If
End Sub

' Identitet izabranog reda. Kolona se PITA (MatKolonaID) jer je kod prava
' skrivena: prva kolona tamo nosi lokalizovan naziv oblasti, a ne kljuc.
'
' Izbor u listi KORISNICI se pamti i posebno: lista prava se cita za njega, pa
' mora da prezivi prelazak na tu listu.
Private Sub ZapamtiIzbor(ByVal lista As String, ByVal red As Long)
    mIzabranID = Trim$(CStr(modOtkupUI.GridCell(red, modMaticniIzvor.MatKolonaID(lista))))
    If lista = "KORISNICI" Then mKorisnikID = mIzabranID
End Sub

' Ukljucuje ili iskljucuje jedno pravo izabranog korisnika. Bez izabranog
' korisnika nema sta da se menja -- i to se kaze, umesto da dugme tiho ne radi.
Private Function PromeniPravo() As Boolean
    Dim odgovor As String, novo As String
    If Len(mKorisnikID) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_KOR_BEZ_IZBORA"), True
        Exit Function
    End If
    If Len(mIzabranID) = 0 Then
        modOtkupUI.ShowToast Poruka("MATU_ERR_NEMA_REDA"), True
        Exit Function
    End If
    odgovor = modMaticniKorisnici.KorPromeniPravo(mKorisnikID, mIzabranID, novo)
    If Len(odgovor) > 0 Then
        modOtkupUI.ShowToast odgovor, True
        Exit Function
    End If
    modOtkupUI.ShowToast modMaticniKorisnici.KorOblastNaziv(mIzabranID) & ": " & _
        IIf(novo = modMaticniKorisnici.KOR_DA, Poruka("OTKUI_KOR_IMA"), _
            Poruka("OTKUI_KOR_NEMA")), False
    PromeniPravo = True
End Function

Private Function PromeniStatus(ByVal lista As String) As Boolean
    Dim red As Long, odgovor As String, noviStatus As String

    If Len(mIzabranID) = 0 Then
        modOtkupUI.ShowToast Poruka("MATU_ERR_NEMA_REDA"), True
        Exit Function
    End If
    red = modMaticniUnos.MatRedPoID(lista, mIzabranID)
    If red = 0 Then
        modOtkupUI.ShowToast Poruka("MATU_ERR_NEMA_REDA"), True
        Exit Function
    End If
    ' Radnja MENJA podatke i vidi se u svakom izvestaju, pa trazi potvrdu --
    ' isti obrazac koji Oporavak koristi za odbacivanje ispravke.
    If MsgBox(Poruka("OTKUI_MAT_STATUS_ASK") & vbCrLf & vbCrLf & mIzabranID, _
              vbExclamation + vbYesNo + vbDefaultButton2, APP_NAME) <> vbYes Then Exit Function

    odgovor = modMaticniUnos.MatPromeniStatus(lista, red, noviStatus)
    If Len(odgovor) > 0 Then
        modOtkupUI.ShowToast odgovor, True
        Exit Function
    End If
    modOtkupUI.ShowToast Poruka("MATU_OK_STATUS") & " " & noviStatus, False
    PromeniStatus = True
End Function

Private Function Sacuvaj(ByVal lista As String) As Boolean
    Dim polja As Object, odgovor As String, noviID As String, red As Long

    If Len(mEditKljuc) = 0 Then Exit Function
    ' KAPIJA: editor otvoren na drugom ekranu ne sme da pise odavde. Bez ovoga
    ' bi "Sacuvaj" upisao u sifarnik koji operater vise ne gleda.
    If Not EditorJeNas(mZonaEkran) Then
        ZatvoriPanele
        modOtkupUI.ShowToast Poruka("MATU_ERR_TUDJI_EDITOR"), True
        modOtkupUI.OsveziRasporedEkrana
        Exit Function
    End If
    Set polja = PokupiPolja(mEditKljuc)

    If Len(mEditID) = 0 Then
        odgovor = modMaticniUnos.MatDodaj(mEditKljuc, polja, noviID)
    Else
        red = modMaticniUnos.MatRedPoID(mEditKljuc, mEditID)
        odgovor = modMaticniUnos.MatIzmeni(mEditKljuc, red, polja)
    End If

    If Len(odgovor) > 0 Then
        modOtkupUI.ShowToast odgovor, True
        FokusNaPolje polja
        Exit Function
    End If

    modOtkupUI.ShowToast IIf(Len(mEditID) = 0, _
        Poruka("MATU_OK_DODATO") & " " & noviID, Poruka("MATU_OK_IZMENJENO")), False
    ' Cenovnik ostaje otvoren sa istim proizvodom -- operater obicno unosi vise
    ' cena zaredom. Isto sto legacy forma radi (brise samo polje cene).
    If mEditKljuc = "CENOVNIK" Then
        PostaviPolje mEditKljuc, "cena", ""
    Else
        ZatvoriEditor
    End If
    Sacuvaj = True
End Function

Public Sub ZatvoriEditor()
    mEditKljuc = ""
    mEditID = ""
    mEditEkran = ""
    mUZavisnima = False
End Sub

' Ljuska javlja da se ekran napusta. Zatvara editor, GEO panel i izbor -- sve
' troje pripada ekranu koji odlazi. Zove se kroz ugovor (Scr_Deaktiviraj), pa
' ljuska i dalje ne zna nijedan ekran po imenu.
' Oslobadjanje uz rusenje ljuske. Ovaj modul NE drzi referencu na formu -- zonu
' trazi kroz modOtkupUI.ScreenZone pri svakoj upotrebi -- ali drzi STANJE koje
' bi preko rusenja preslo na sledecu instancu (otvoren editor nad zonom koje
' vise nema). Zove ga Scr_Deaktiviraj i, preko njega, ljuska.
Public Sub Deaktiviraj()
    ZatvoriPanele
    mIzabranID = ""
    mKorisnikID = ""
    mZonaEkran = ""
End Sub

' Da li je otvoreni editor BAS ovog ekrana. Prazan editor je "nicij" i prolazi.
Private Function EditorJeNas(ByVal ekran As String) As Boolean
    If Len(mEditKljuc) = 0 Then
        EditorJeNas = True
    Else
        EditorJeNas = (mEditEkran = ekran)
    End If
End Function

' Prelazak na drugu listu zatvara i editor i GEO panel: oba pripadaju sekciji
' koja se napusta.
Private Sub ZatvoriPanele()
    ZatvoriEditor
    mGeoID = ""
End Sub

'--------------------------------------------------- POLJA <-> RECNIK
Private Function PokupiPolja(ByVal lista As String) As Object
    Dim d As Object, a As Variant, r As Variant, k As String
    Set d = CreateObject("Scripting.Dictionary")
    Set PokupiPolja = d
    a = modMaticniIzvor.MatPolja(lista)
    If Not IsArray(a) Then Exit Function
    For Each r In a
        k = modMaticniIzvor.PoljeF(CStr(r), 0)
        d(k) = CitajPolje(lista, k)
    Next r
End Function

Private Function CitajPolje(ByVal lista As String, ByVal poljeKljuc As String) As String
    Dim z As Object, nm As String
    On Error Resume Next
    Set z = Zona()
    If z Is Nothing Then Exit Function
    nm = KontrolaPolja(lista, poljeKljuc)
    If Len(nm) = 0 Then Exit Function
    CitajPolje = Trim$(NzToText(z.Controls(nm).Controls(nm & "T").value))
    Err.Clear
End Function

Private Sub PostaviPolje(ByVal lista As String, ByVal poljeKljuc As String, ByVal v As String)
    Dim z As Object, nm As String
    On Error Resume Next
    Set z = Zona()
    If z Is Nothing Then Exit Sub
    nm = KontrolaPolja(lista, poljeKljuc)
    If Len(nm) = 0 Then Exit Sub
    z.Controls(nm).Controls(nm & "T").value = v
    Err.Clear
End Sub

' Prazna polja + sveze napunjeni combo-i. Combo se puni pri SVAKOM otvaranju:
' spisak stanica i kultura se menja iz drugih sekcija istog ekrana.
Private Sub OcistiPolja(ByVal lista As String)
    Dim z As Object, a As Variant, r As Variant, k As String, nm As String
    On Error Resume Next
    Set z = Zona()
    If z Is Nothing Then Exit Sub
    PoljaEditora z, lista, True
    a = modMaticniIzvor.MatPolja(lista)
    If Not IsArray(a) Then Exit Sub
    For Each r In a
        k = modMaticniIzvor.PoljeF(CStr(r), 0)
        nm = KontrolaPolja(lista, k)
        If Len(nm) > 0 Then
            PuniCombo z, lista, CStr(r)
            z.Controls(nm).Controls(nm & "T").value = ""
        End If
    Next r
    Err.Clear
End Sub

' Puni JEDAN combo iz njegovog izvora, sa kontekstom procitanim IZ POLJA od kog
' izvor zavisi. Polje bez izvora se preskace.
Private Sub PuniCombo(ByVal z As Object, ByVal lista As String, ByVal spec As String)
    Dim izvor As String, nm As String, kontekst As String, zavisi As String
    On Error Resume Next
    izvor = modMaticniIzvor.PoljeF(spec, 5)
    If Len(izvor) = 0 Then Exit Sub
    nm = KontrolaPolja(lista, modMaticniIzvor.PoljeF(spec, 0))
    If Len(nm) = 0 Then Exit Sub
    zavisi = modMaticniIzvor.MatComboZavisi(izvor)
    If Len(zavisi) > 0 Then kontekst = CitajPolje(lista, zavisi)
    FillCmb z.Controls(nm).Controls(nm & "T"), _
            modMaticniIzvor.MatComboStavke(izvor, kontekst)
    Err.Clear
End Sub

' Ponovo puni combo-e koji ZAVISE od datog polja, i zadrzava izbor ako i dalje
' postoji u novom spisku -- tacno ono sto je legacy radio u cmbField1_Change.
' Prazno "promenjeno" znaci "sve zavisne", i tako se zove posle ucitavanja
' postojeceg zapisa: spisak se puni pri otvaranju, kad roditelj jos nema
' vrednost, pa bi bez ovoga sorta ostala prazna i na izmeni i na novom unosu.
Private Sub OsveziZavisne(ByVal lista As String, ByVal promenjeno As String)
    Dim z As Object, a As Variant, r As Variant, izvor As String
    Dim k As String, staro As String
    If mUZavisnima Then Exit Sub
    On Error Resume Next
    Set z = Zona()
    If z Is Nothing Then Exit Sub
    a = modMaticniIzvor.MatPolja(lista)
    If Not IsArray(a) Then Exit Sub
    mUZavisnima = True
    For Each r In a
        izvor = modMaticniIzvor.PoljeF(CStr(r), 5)
        If Len(izvor) > 0 Then
            k = modMaticniIzvor.MatComboZavisi(izvor)
            If Len(k) > 0 Then
                If Len(promenjeno) = 0 Or k = promenjeno Then
                    k = modMaticniIzvor.PoljeF(CStr(r), 0)
                    staro = CitajPolje(lista, k)
                    PuniCombo z, lista, CStr(r)
                    ' FillCmb prazni i listu i vrednost; izbor se vraca samo ako
                    ' ga nov spisak i dalje sadrzi.
                    If Len(staro) > 0 Then PostaviPolje lista, k, staro
                End If
            End If
        End If
    Next r
    mUZavisnima = False
    Err.Clear
End Sub

' Kljuc polja koje nosi datu kontrolu bazena -- obrnut smer od KontrolaPolja.
' Racuna se iz istog opisa, pa se dva smera ne mogu razici.
Private Function PoljeIzKontrole(ByVal lista As String, ByVal nm As String) As String
    Dim a As Variant, r As Variant, k As String
    a = modMaticniIzvor.MatPolja(lista)
    If Not IsArray(a) Then Exit Function
    For Each r In a
        k = modMaticniIzvor.PoljeF(CStr(r), 0)
        If KontrolaPolja(lista, k) = nm Then
            PoljeIzKontrole = k
            Exit Function
        End If
    Next r
End Function

Private Sub FokusNaPolje(ByVal polja As Object)
    Dim z As Object, nm As String
    On Error Resume Next
    If polja Is Nothing Then Exit Sub
    If Not polja.Exists(modMaticniUnos.MAT_FOKUS) Then Exit Sub
    Set z = Zona()
    If z Is Nothing Then Exit Sub
    nm = KontrolaPolja(mEditKljuc, CStr(polja(modMaticniUnos.MAT_FOKUS)))
    If Len(nm) > 0 Then z.Controls(nm).Controls(nm & "T").SetFocus
    Err.Clear
End Sub

Private Function Zona() As Object
    Set Zona = modOtkupUI.ScreenZone(mZonaEkran)
End Function

'------------------------------------------------------------ TEST SEAM
Public Function MatEditKljucTest() As String
    MatEditKljucTest = mEditKljuc
End Function

Public Function MatEditIDTest() As String
    MatEditIDTest = mEditID
End Function

' Kontrola bazena koja nosi dato polje -- tvrdnja da se dodela ne razilazi sa
' rasporedom (tekst i combo idu u odvojene brojace).
Public Function MatKontrolaPoljaTest(ByVal lista As String, ByVal poljeKljuc As String) As String
    MatKontrolaPoljaTest = KontrolaPolja(lista, poljeKljuc)
End Function

' Kljuc polja od kog combo datog polja zavisi -- "" ako ne zavisi ni od cega.
' Racuna se ISTIM putem kojim se combo puni (opis polja -> izvor -> zavisnost),
' pa tvrdnja o kaskadi meri ono sto editor stvarno radi, a ne svoju kopiju.
Public Function MatZavisnostPoljaTest(ByVal lista As String, ByVal poljeKljuc As String) As String
    Dim a As Variant, r As Variant
    a = modMaticniIzvor.MatPolja(lista)
    If Not IsArray(a) Then Exit Function
    For Each r In a
        If modMaticniIzvor.PoljeF(CStr(r), 0) = poljeKljuc Then
            MatZavisnostPoljaTest = _
                modMaticniIzvor.MatComboZavisi(modMaticniIzvor.PoljeF(CStr(r), 5))
            Exit Function
        End If
    Next r
End Function

Public Function MatPoljeIzKontroleTest(ByVal lista As String, ByVal nm As String) As String
    MatPoljeIzKontroleTest = PoljeIzKontrole(lista, nm)
End Function

Public Function MatVisinaZoneTest(ByVal lista As String, ByVal otvoren As Boolean) As Single
    If otvoren Then
        MatVisinaZoneTest = VisinaOtvorene(lista)
    Else
        MatVisinaZoneTest = MAT_ZONA_H
    End If
End Function
