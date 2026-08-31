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
' ZONA IMA DVA STANJA (M2b):
'   zatvorena  pregled -- naziv sekcije, koliko zapisa, koliko aktivnih;
'   otvorena   editor -- polja sekcije, "Sacuvaj" i "Odustani".
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
' forma (v. UI_MIGRACIJA_KATALOG 24.5 i 24.15).
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const MATEKR_BUILD As String = "v6-ui-190"

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

' Stanje editora. Prazan mEditKljuc znaci da je zona zatvorena; prazan mEditID
' uz otvoren editor znaci NOV zapis.
Private mEditKljuc As String
Private mEditID As String
Private mIzabranID As String               ' PK reda izabranog u mrezi

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
    z.Controls("matHint").Visible = (w - 2 * PAD - 320 > 120) And Len(mEditKljuc) = 0
    If w - 2 * PAD - 320 > 120 Then z.Controls("matHint").width = w - 2 * PAD - 320

    If Len(mEditKljuc) = 0 Then
        PoljaEditora z, "", False
        modUiKit.BoxShow z, "scrMatSacuvaj", False
        modUiKit.BoxShow z, "scrMatOdustani", False
        z.Controls("matEdBg").Visible = False
        z.Controls("matEdCap").Visible = False
        ' "Nova stavka" postoji i kad je editor zatvoren -- to mu je jedini ulaz.
        modUiKit.MoveBox z, "scrMatNovi", w - PAD - 130, 14, 130
        modUiKit.BoxShow z, "scrMatNovi", (Len(lista) > 0)
        h = MAT_ZONA_H
    Else
        modUiKit.BoxShow z, "scrMatNovi", False
        h = RasporediEditor(z, w)
    End If

    z.Controls("matLnB").top = h - 1
    z.Controls("matLnB").width = w
    ZonaRaspored = h
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
    If imaStatus Then
        z.Controls("matKV0").caption = CStr(modMaticniIzvor.MatAktivnih())
        z.Controls("matKV1").caption = CStr(modMaticniIzvor.MatNeaktivnih())
    Else
        z.Controls("matKV0").caption = ChrW(8212)
        z.Controls("matKV1").caption = ChrW(8212)
    End If
End Sub

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
    Redovi = modMaticniIzvor.MatRedovi(lista, filter, q)
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
    If lista <> "CENOVNIK" Then s = "izmeni:OTKUI_BTN_MAT_IZMENI:92:soft:1"
    If Len(modMaticniIzvor.MatStatusKolona(lista)) > 0 Then
        If Len(s) > 0 Then s = s & "|"
        s = s & "status:OTKUI_BTN_MAT_STATUS:150:danger:1"
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
    preEdit = mEditKljuc & "/" & mEditID
    Dogadjaj = ObradiDogadjaj(tag, lista)
    ' Otvaranje ili zatvaranje editora menja VISINU zone. Ljuska to ne moze da
    ' pogodi -- LayoutScreenZone ide samo iz LayoutOtkup -- pa se trazi izricito.
    If (mEditKljuc & "/" & mEditID) <> preEdit Then modOtkupUI.OsveziRasporedEkrana
End Function

Private Function ObradiDogadjaj(ByVal tag As String, ByRef lista As String) As Boolean
    Dim red As Long

    If Left$(tag, 2) = "ls" Then
        If Mid$(tag, 3) = lista Then Exit Function
        ' Odlazak sa liste zatvara editor: polja pripadaju sekciji koja se
        ' napusta, a ostavljena bi trazila upis u drugu tabelu.
        ZatvoriEditor
        lista = Mid$(tag, 3)
        ObradiDogadjaj = True
        Exit Function
    End If

    If Left$(tag, 4) = "row:" Then
        red = CLng(val(Mid$(tag, 5)))
        mIzabranID = Trim$(CStr(modOtkupUI.GridCell(red, 1)))
        Exit Function
    End If

    ' Dvoklik otvara izmenu -- jedan potez umesto dva (izaberi red, pa dugme).
    If Left$(tag, 4) = "dbl:" Then
        red = CLng(val(Mid$(tag, 5)))
        mIzabranID = Trim$(CStr(modOtkupUI.GridCell(red, 1)))
        OtvoriIzmenu lista
        Exit Function
    End If

    If Left$(tag, 4) = "act:" Then
        ObradiDogadjaj = Akcija(tag, lista)
        Exit Function
    End If

    ' Otvaranje i zatvaranje editora NE menjaju podatke, pa vracaju False:
    ' ljuska bi inace na svaki klik radila pun RefreshFromData (KPI, kes,
    ' combo-i) za cistu promenu prikaza. Raspored se trazi izricito, gore.
    Select Case tag
        Case "scrMatNovi":     OtvoriNov lista
        Case "scrMatOdustani": ZatvoriEditor
        Case "scrMatSacuvaj":  ObradiDogadjaj = Sacuvaj(lista)
    End Select
End Function

Private Function Akcija(ByVal tag As String, ByVal lista As String) As Boolean
    Dim p() As String, red As Long
    p = Split(Mid$(tag, 5), ":")
    If UBound(p) < 1 Then
        LogWarn "modMaticniEkran.Akcija", "Radnja bez rednog broja: '" & tag & "'."
        Exit Function
    End If
    red = CLng(val(p(1)))
    ' Identitet reda dolazi iz KOLONE 1 (PK), ne iz rednog broja -- posle
    ' sortiranja i pretrage pozicija ne znaci nista.
    mIzabranID = Trim$(CStr(modOtkupUI.GridCell(red, 1)))
    Select Case p(0)
        Case "izmeni": OtvoriIzmenu lista
        Case "status": Akcija = PromeniStatus(lista)
        Case Else
            LogWarn "modMaticniEkran.Akcija", "Nepoznata radnja '" & p(0) & "'."
            modOtkupUI.ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & p(0), True
    End Select
End Function

'------------------------------------------------------------- EDITOR
Private Sub OtvoriNov(ByVal lista As String)
    If Len(lista) = 0 Then Exit Sub
    mEditKljuc = lista
    mEditID = ""
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
    mEditKljuc = lista
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

    If jeCenovnik Then
        ' Cena se unosi NOVA; datum ostaje prazan pa pisac uzima danasnji.
        PostaviPolje lista, "cena", ""
        PostaviPolje lista, "datum", ""
        modOtkupUI.ShowToast Poruka("MATU_ERR_CENOVNIK_APPEND"), False
    End If
End Sub

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
    Dim z As Object, a As Variant, r As Variant, k As String, izvor As String
    Dim nm As String, stavke As Variant, kontekst As String
    On Error Resume Next
    Set z = Zona()
    If z Is Nothing Then Exit Sub
    PoljaEditora z, lista, True
    a = modMaticniIzvor.MatPolja(lista)
    If Not IsArray(a) Then Exit Sub
    For Each r In a
        k = modMaticniIzvor.PoljeF(CStr(r), 0)
        izvor = modMaticniIzvor.PoljeF(CStr(r), 5)
        nm = KontrolaPolja(lista, k)
        If Len(nm) > 0 Then
            If Len(izvor) > 0 Then
                kontekst = ""
                ' Sorte zavise od izabrane vrste -- kaskada iz legacy forme.
                If izvor = "@sorte" Then kontekst = CitajPolje(lista, "vrsta")
                stavke = modMaticniIzvor.MatComboStavke(izvor, kontekst)
                FillCmb z.Controls(nm).Controls(nm & "T"), stavke
            End If
            z.Controls(nm).Controls(nm & "T").value = ""
        End If
    Next r
    Err.Clear
End Sub

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

Public Function MatVisinaZoneTest(ByVal lista As String, ByVal otvoren As Boolean) As Single
    If otvoren Then
        MatVisinaZoneTest = VisinaOtvorene(lista)
    Else
        MatVisinaZoneTest = MAT_ZONA_H
    End If
End Function
