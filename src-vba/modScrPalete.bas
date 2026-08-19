Attribute VB_Name = "modScrPalete"
'=====================================================================
' modScrPalete - ekran "Palete".
'
' Ljuska ga ne poznaje po imenu: dobija ga preko Application.Run, da klijent
' kome ovaj modul nedostaje i dalje radi (zamka #19).
'
' PUN EKRAN (od P1). Prenosi ono sto legacy frmPalete radi kroz tri liste
' istog prekidaca koji F1 vec koristi:
'
'   PALETE   glavna lista; radnje: stampa, PDF, zatvaranje, storno,
'            i stampa svih nepotpunih (radnja bez reda)
'   STAVKE   stavke izabrane palete (prijemnice od kojih je slozena)
'   PRERADE  prerade; radnja: stampa i storno prerade
'
' Zona ekrana nosi AKTIVNU PALETU (broj, roba, status) i tri ukupna broja -
' isti raspored kao traka otpremnice u F1.
'
' NISTA se ovde ne racuna niti upisuje: podaci dolaze iz modPaletniList
' (GetPaleteForGrid / GetPaletaStavkeForGrid / GetPreradeForGrid), a radnje iz
' modPaletniList i modStorno - iste rutine koje zove i legacy forma.
'
' NIJE preneto u P1: unos prerade (tezina palete, bruto, kutije, kese, tip
' gotovog proizvoda). Taj panel trazi da ljuska prosledjuje ekranu i dogadjaje
' sopstvenih kontrola, sto je sledeci korak.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const SCRPAL_BUILD As String = "v6-ui-165"

' Visina zone = visina KPI trake na ekranu dokumenata. Zona ugovornog ekrana
' stoji na istom mestu i iste je visine, pa naslov ispod nje pada u isti red
' na oba ekrana.
Private Const PAL_ZONA_H As Single = KPI_H

' Stanje ekrana. Aktivna paleta je ona na koju je poslednji put kliknuto u
' listi paleta; nju gledaju i stavke i radnje nad jednom paletom.
Private mLista As String          ' "PALETE" | "STAVKE" | "PRERADE" | ST_NOVA

' CETVRTA LISTA: unos nove prerade. Nije pregled kao ostale tri nego RADNI
' ekran -- mreza sluzi da se palete oznace, a zona nosi polja unosa. Legacy je
' isti posao radio panelom sa sedam polja i multiselektom liste.
Private Const ST_NOVA As String = "NOVAPRERADA"

' Kolona liste za preradu koja nosi PaletaID. Jedan broj deljen izmedju opisa
' kolona, punjenja reda i radnje -- isti razlog kao NED_COL_CID na Oporavku.
Public Const PAL_NOVA_COL_ID As Long = 14

' Visina zone dok se unosi prerada: KPI traka + dva reda polja + red dugmeta.
' Scr_Layout vraca visinu, pa ekran sme da je promeni po listi.
Private Const PAL_ZONA_NOVA_H As Single = PAL_ZONA_H + 2 * 46 + 40

' Razmak izmedju polja. Ljuskin GAP je Private, a ekran nema pravo da ga
' otvara zbog sopstvenog rasporeda -- isti broj, manji domet.
Private Const PRE_GAP As Single = 10

' Oznacene palete, po PaletaID. Prazno = nijedna, i to je podrazumevano stanje.
Private mPreOznacene As Object
' PaletaID -> broj palete. Izbor se DRZI po identitetu (broj se ponavlja kroz
' godine), ali operater misli u brojevima -- pa zona mora da ume da ih imenuje.
Private mPreBrojevi As Object
' Kombo ponude se pune jednom, ne pri svakom rasporedu.
Private mPreCombosFilled As Boolean
Private mPalID As String
Private mPalBroj As String
Private mPalIds As Object         ' broj palete -> PaletaID
Private mPreIds As Object         ' broj prerade -> PreradaID
Private mStep As String           ' korak za poruku o gresci
' Kratak opis aktivne palete (vrsta, sorta, klasa, status). Pamti se pri izboru
' reda da bi prezivelo prelazak na listu stavki, gde tih kolona nema.
Private mPalOpis As String

' Kolona "PRERADJENO" u mrezi paleta - radnja storna je gleda pre poziva.
Private Const PAL_KOL_PRERADJENO As Long = 12

'--------------------------------------------------------- UGOVOR EKRANA
Public Function Scr_Meta() As String
    Scr_Meta = "kljuc=PALETE|naslov=OTKUI_NAV_PALETE|sub=OTKUI_SCRPAL_SUB" & _
               "|lista=OTKUI_SCRPAL_LISTA|oblik=lista|upis=ne"
End Function

Public Function Scr_Liste() As Variant
    Scr_Liste = Array( _
        "PALETE|OTKUI_SEG_PAL_PALETE|OTKUI_GRID_TITLE_PALETE|96", _
        "STAVKE|OTKUI_SEG_PAL_STAVKE|OTKUI_GRID_TITLE_PALSTAVKE|110", _
        "PRERADE|OTKUI_SEG_PAL_PRERADE|OTKUI_GRID_TITLE_PRERADE|96", _
        ST_NOVA & "|OTKUI_SEG_PAL_NOVA|OTKUI_GRID_TITLE_PALNOVA|104")
End Function

Public Function Scr_Lista() As String
    If Len(mLista) = 0 Then mLista = "PALETE"
    Scr_Lista = mLista
End Function

' U listi stavki naslov nosi broj palete cije su to stavke.
Public Function Scr_NaslovDopuna() As String
    If Scr_Lista() = "STAVKE" Then Scr_NaslovDopuna = mPalBroj
End Function

' Cipovi liste paleta: godina, status, preradjenost. Isti posao koji je legacy
' panel radio kroz tri odvojena kontrolna polja iznad liste.
Public Function Scr_Cipovi() As String
    If Scr_Lista() <> "PALETE" Then Exit Function
    Scr_Cipovi = "sve:OTKUI_CHIP_SVE:40|" & _
                 "godina:OTKUI_CIPP_GODINA:84|" & _
                 "otvorene:OTKUI_CIPP_OTVORENE:76|" & _
                 "zatvorene:OTKUI_CIPP_ZATVORENE:84|" & _
                 "preradjene:OTKUI_CIPP_PRERADJENE:88"
End Function

' Da li paleta prolazi kroz izabrani cip. Kljuc je EKRANOV -- ljuska ga je
' samo vratila onakvog kakvog ga je dobila iz Scr_Cipovi. Javna je da bi
' pravilo moglo da se izmeri bez mreze.
Public Function PalCipProlaz(ByVal filter As String, ByVal status As String, _
                             ByVal godina As String, _
                             ByVal preradjeno As String) As Boolean
    Select Case filter
        Case "godina":     PalCipProlaz = (val(godina) = Year(Date))
        Case "otvorene":   PalCipProlaz = (UCase$(Trim$(status)) <> "ZATVORENA")
        Case "zatvorene":  PalCipProlaz = (UCase$(Trim$(status)) = "ZATVORENA")
        Case "preradjene": PalCipProlaz = (UCase$(Trim$(preradjeno)) = "DA")
        Case Else:         PalCipProlaz = True
    End Select
End Function

' Radnje nad redom za aktivnu listu: kljuc:natpis:sirina:stil:trebaRed
Public Function Scr_Radnje() As String
    Select Case Scr_Lista()
        Case "PALETE"
            Scr_Radnje = "palstavke:OTKUI_BTN_PAL_STAVKE:88:soft:1|" & _
                         "palprint:OTKUI_BTN_PAL_PRINT:112:ghost:1|" & _
                         "palpdf:OTKUI_BTN_PAL_PDF:70:ghost:1|" & _
                         "palzatvori:OTKUI_BTN_PAL_ZATVORI:124:soft:1|" & _
                         "palstorno:OTKUI_BTN_RED_STORNO:88:danger:1|" & _
                         "palnepotpune:OTKUI_BTN_PAL_NEPOTPUNE:148:ghost:0"
        Case "PRERADE"
            Scr_Radnje = "preprint:OTKUI_BTN_PRE_PRINT:132:ghost:1|" & _
                         "prestorno:OTKUI_BTN_PRE_STORNO:150:danger:1"
    End Select
End Function

'--------------------------------------------------------------- ZONA
Public Sub Scr_Build(ByVal z As Object)
    Dim i As Long
    ' levo: aktivna paleta - isti raspored kao traka otpremnice u F1
    modUiKit.NewLbl z, "palCap", UCase$(Poruka("OTKUI_PAL_AKTIVNA")), PAD, 6, 140, 11, _
                    TS_MICRO, True, C_MUTED, -1
    modUiKit.NewLbl z, "palBroj", ChrW(8212), PAD, 18, 260, 20, TS_KPI, True, C_FOREST, -1
    modUiKit.NewLbl z, "palSub", Poruka("OTKUI_PAL_NEMA"), PAD, 40, 420, 13, _
                    TS_META, False, C_MUTED, -1

    ' desno: tri brojke - isti materijal kao KPI traka, samo uze
    For i = 0 To 2
        modUiKit.NewLbl z, "palKL" & i, "", 0, 6, 120, 12, TS_MICRO, True, C_MUTED, -1
        modUiKit.NewLbl z, "palKV" & i, ChrW(8212), 0, 18, 120, 20, TS_KPI, True, _
                        C_FOREST, -1, fmTextAlignLeft, F_NUM
    Next i

    ' NOVA PRERADA: polja unosa. Postoje uvek, a vide se samo dok je ta lista
    ' aktivna -- Scr_Layout ih pali i gasi. Prefiks 'scr' je OBAVEZAN: bez njega
    ' promena teksta ide ljusci, koja o ovim poljima ne zna nista.
    ' Bela podloga ispod celog panela. Bez nje se izmedju polja vidi krem
    ' pozadina zone, pa panel izgleda kao niz odvojenih ostrva umesto kao jedna
    ' celina -- operater je to prijavio kao 'ruzni prekidi izmedju belih polja'.
    ' Pravi se PRE polja, jer u MSForms kasnije dodata kontrola stoji IZNAD.
    modUiKit.NewFrame z, "preBg", 0, 0, 100, 10, C_WHITE
    modUiKit.NewLbl z, "preCap", UCase$(Poruka("OTKUI_PRE_CAP")), PAD, PAL_ZONA_H + 4, 200, 11, _
                    TS_MICRO, True, C_MUTED, -1
    modOtkupUI.NewFieldG z, "scrPreBruto", Poruka("OTKUI_PRE_BRUTO"), "txt", "kg", 1, True, False, "PRE"
    modOtkupUI.NewFieldG z, "scrPreTezPal", Poruka("OTKUI_PRE_TEZPAL"), "txt", "kg", 1, True, False, "PRE"
    modOtkupUI.NewFieldG z, "scrPreGP", Poruka("OTKUI_PRE_GP"), "cmb", "", 1, False, False, "PRE"
    modOtkupUI.NewFieldG z, "scrPreNap", Poruka("OTKUI_PRE_NAP"), "txt", "", 1, False, False, "PRE"
    modOtkupUI.NewFieldG z, "scrPreKut", Poruka("OTKUI_PRE_KUT"), "txt", "kom", 1, True, False, "PRE"
    modOtkupUI.NewFieldG z, "scrPreTipKut", Poruka("OTKUI_PRE_TIPKUT"), "cmb", "", 1, False, False, "PRE"
    modOtkupUI.NewFieldG z, "scrPreKes", Poruka("OTKUI_PRE_KES"), "txt", "kom", 1, True, False, "PRE"
    modOtkupUI.NewFieldG z, "scrPreTipKes", Poruka("OTKUI_PRE_TIPKES"), "cmb", "", 1, False, False, "PRE"

    ' Neto se ne unosi nego RACUNA, pa stoji kao brojka a ne kao polje: bruto
    ' minus tezina palete minus ambalaza. Menja se pri svakom kucanju.
    modUiKit.NewLbl z, "preNetoL", UCase$(Poruka("OTKUI_PRE_NETO")), 0, 0, 130, 11, _
                    TS_MICRO, True, C_MUTED, -1
    modUiKit.NewLbl z, "preNetoV", ChrW(8212), 0, 0, 130, 20, TS_KPI, True, C_FOREST, -1, _
                    fmTextAlignLeft, F_NUM
    modUiKit.NewLbl z, "preIzbor", "", 0, 0, 220, 13, TS_META, False, C_MUTED, -1
    modUiKit.BtnV z, "scrPreradi", Poruka("OTKUI_BTN_PRE_URADI"), 0, 0, 168, 26, "primary"

    modUiKit.NewLbl z, "palLnB", "", 0, PAL_ZONA_H - 1, 100, 1, 8, False, 0, C_BORDER
End Sub

' Ponude kombo polja. Isti izvori koje legacy panel koristi (GetKutijeOptions,
' GetKeseOptions, GetVrstaGPOptions) -- ekran ih ne izmislja.
'
' FillCmb prima ByRef MSForms.ComboBox, pa Object ne sme direktno (Argument type
' mismatch); zato tipizirani lokali. Procedure-level 'As MSForms.' je bezbedno --
' IsHardModuleBody skenira samo modul-level deo, pa modul ostaje mek za
' self-update.
Private Sub PuniPreradaCombo(ByVal z As Object)
    Dim cbGP As MSForms.ComboBox, cbKut As MSForms.ComboBox, cbKes As MSForms.ComboBox
    On Error Resume Next
    If mPreCombosFilled Then Exit Sub
    Set cbGP = z.Controls("scrPreGP").Controls("scrPreGPT")
    Set cbKut = z.Controls("scrPreTipKut").Controls("scrPreTipKutT")
    Set cbKes = z.Controls("scrPreTipKes").Controls("scrPreTipKesT")
    If cbGP Is Nothing Or cbKut Is Nothing Or cbKes Is Nothing Then Exit Sub
    FillCmb cbGP, GetVrstaGPOptions()
    FillCmb cbKut, GetKutijeOptions()
    FillCmb cbKes, GetKeseOptions()
    mPreCombosFilled = True
End Sub

' Polja postoje uvek; vide se samo u listi za unos prerade.
' Pali i gasi panel za unos prerade.
'
' Kontrola koje NEMA se ovde ne moze popraviti, ali sme da bude PRIJAVLJENA:
' bez toga `On Error Resume Next` proguta i ime i broj, pa operater vidi
' prazninu na mestu polja, a log ne kaze nista. Prijavljuje se jednom po
' skupu, ne po kontroli, da ne zatrpa log pri svakom prelasku liste.
Private Sub PoljaPrerade(ByVal z As Object, ByVal vis As Boolean)
    Dim nm As Variant, fale As String
    For Each nm In Array("scrPreBruto", "scrPreTezPal", "scrPreGP", "scrPreNap", _
                         "scrPreKut", "scrPreTipKut", "scrPreKes", "scrPreTipKes", _
                         "preBg", "preCap", "preNetoL", "preNetoV", "preIzbor")
        If Not UpaliKontrolu(z, CStr(nm), vis) Then fale = fale & " " & CStr(nm)
    Next nm
    On Error Resume Next
    modUiKit.BoxShow z, "scrPreradi", vis
    ' Prijava ne sme da obori ono sto prijavljuje -- zato ostaje pod Resume Next.
    If Len(fale) > 0 Then
        LogWarn "modScrPalete.PoljaPrerade", _
                "zona nema kontrole:" & fale & " | zona=" & ZonaOpis(z)
        ' I NA EKRAN, ne samo u log: rupa u panelu se vidi odmah, pa i njen razlog
        ' treba da stigne tu gde operater gleda. Samo pri paljenju -- gasenje panela
        ' nad listom pregleda ne treba nikome da javlja nista.
        If vis Then modOtkupUI.ShowToast _
            "Zona: nema" & fale & " (" & ZonaOpis(z) & ")", True
    End If
    Err.Clear
End Sub

' True ako kontrola postoji i vidljivost je postavljena.
Private Function UpaliKontrolu(ByVal z As Object, ByVal nm As String, _
                               ByVal vis As Boolean) As Boolean
    On Error Resume Next
    z.Controls(nm).Visible = vis
    UpaliKontrolu = (Err.Number = 0)
    Err.Clear
End Function

' Ime i broj kontrola zone -- bez toga se iz loga ne vidi da li je zona
' uopste ona prava, ili je gradnja stala na pola.
Private Function ZonaOpis(ByVal z As Object) As String
    On Error Resume Next
    ZonaOpis = z.name & "/" & z.Controls.count
    Err.Clear
End Function

Public Function Scr_Layout(ByVal z As Object, ByVal w As Single, ByVal h As Single) As Single
    Dim i As Long
    On Error Resume Next
    For i = 0 To 2
        z.Controls("palKL" & i).Left = w - PAD - (3 - i) * 150
        z.Controls("palKV" & i).Left = w - PAD - (3 - i) * 150
    Next i
    z.Controls("palLnB").width = w

    ' Zona raste SAMO za unos prerade. Ostale tri liste su pregledi i njima je
    ' KPI traka dovoljna -- visa zona bi im samo pojela redove mreze.
    If Scr_Lista() <> ST_NOVA Then
        PoljaPrerade z, False
        z.Controls("palLnB").top = PAL_ZONA_H - 1
        Scr_Layout = PAL_ZONA_H
        Exit Function
    End If

    PoljaPrerade z, True
    PuniPreradaCombo z
    ' Podloga ide od ivice do ivice zone, sa malim uvlacenjem levo i desno --
    ' bez njega prva labela stoji zalepljena za belu ivicu.
    z.Controls("preBg").Left = PAD - 10
    z.Controls("preBg").top = PAL_ZONA_H
    z.Controls("preBg").width = w - 2 * (PAD - 10)
    z.Controls("preBg").Height = PAL_ZONA_NOVA_H - PAL_ZONA_H - 1
    Dim kol As Single, x0 As Single, y0 As Single, nm As Variant
    kol = (w - PAD * 2 - 3 * PRE_GAP - 200) / 4
    If kol < 120 Then kol = 120
    y0 = PAL_ZONA_H + 18
    i = 0
    For Each nm In Array("scrPreBruto", "scrPreTezPal", "scrPreGP", "scrPreNap", _
                         "scrPreKut", "scrPreTipKut", "scrPreKes", "scrPreTipKes")
        z.Controls(CStr(nm)).Left = PAD + (i Mod 4) * (kol + PRE_GAP)
        z.Controls(CStr(nm)).top = y0 + (i \ 4) * 46
        z.Controls(CStr(nm)).width = kol
        ' Bez ovoga unutrasnje kontrole ostaju na merama iz gradnje (180pt):
        ' jedinica se nadje nasred polja, a unos izgleda odsecen.
        modOtkupUI.LayoutFieldInner z.Controls(CStr(nm))
        i = i + 1
    Next nm

    x0 = PAD + 4 * (kol + PRE_GAP)
    z.Controls("preNetoL").Left = x0
    z.Controls("preNetoL").top = y0
    z.Controls("preNetoV").Left = x0
    z.Controls("preNetoV").top = y0 + 12
    z.Controls("preIzbor").Left = x0
    z.Controls("preIzbor").top = y0 + 36
    modUiKit.MoveBox z, "scrPreradi", x0, y0 + 56, 168
    z.Controls("palLnB").top = PAL_ZONA_NOVA_H - 1
    Scr_Layout = PAL_ZONA_NOVA_H
End Function

'-------------------------------------------------------------- RADNJE
Public Function Scr_Event(ByVal tag As String, ByVal ev As String) As Boolean
    On Error Resume Next

    If Left$(tag, 2) = "ls" Then
        If Mid$(tag, 3) = Scr_Lista() Then Exit Function
        ' Odlazak sa liste za unos ponistava oznacene palete: one pripadaju
        ' preradi koja se upravo sprema, a ostavljene bi sledeci put usle u
        ' spisak koji operater nije video.
        If Scr_Lista() = ST_NOVA Then OcistiPreradu
        mLista = Mid$(tag, 3)
        Scr_Event = True
        Exit Function
    End If

    ' Klik na red u listi za unos UKLJUCUJE ILI ISKLJUCUJE paletu. Vraca True da
    ' bi se mreza procitala ponovo -- kvacica se crta iz podataka.
    If Left$(tag, 4) = "row:" And Scr_Lista() = ST_NOVA Then
        Scr_Event = OznaciPaletu(CLng(Mid$(tag, 5)))
        Exit Function
    End If

    ' Promena u polju zone. Neto se racuna uzivo, pa operater vidi rezultat pre
    ' nego sto potvrdi -- isti razlog zbog kog uvid o stornu stoji pre odluke.
    If Left$(tag, 4) = "chg:" Then
        OsveziNeto
        Exit Function
    End If

    If tag = "scrPreradi" Then
        Scr_Event = PreradiIzabrane()
        Exit Function
    End If

    ' Izbor reda u listi paleta postavlja AKTIVNU paletu. Ekran vraca False:
    ' izbor ne menja podatke, pa mreza ne sme da se prazni - radnje nad redom
    ' rade bas nad tim izabranim redom.
    If Left$(tag, 4) = "row:" And Scr_Lista() = "PALETE" Then
        PostaviAktivnu CLng(Mid$(tag, 5))
        Exit Function
    End If

    ' Dvoklik na paletu OTVARA njene stavke -- jedan potez umesto dva (izaberi
    ' red, pa prebaci prekidac). Vraca True: lista se promenila, pa ljuska cita
    ' mrezu ponovo i pretvara prekidac.
    If Left$(tag, 4) = "dbl:" And Scr_Lista() = "PALETE" Then
        Scr_Event = OtvoriStavke(CLng(Mid$(tag, 5)))
        Exit Function
    End If

    If Left$(tag, 4) = "act:" Then
        Scr_Event = PalAkcija(tag)
        Exit Function
    End If
End Function

' Vraca True samo ako je radnja PROMENILA podatke (mreza se tada cita ponovo).
Private Function PalAkcija(ByVal tag As String) As Boolean
    Dim p() As String, red As Long, broj As String, iD As String
    On Error GoTo EH
    p = Split(Mid$(tag, 5), ":")
    If UBound(p) < 1 Then Exit Function
    red = CLng(val(p(1)))
    broj = Trim$(CStr(modOtkupUI.GridCell(red, 1)))

    Select Case p(0)
        Case "palstavke"
            ' ne menja podatke, ali menja LISTU -- mreza mora da se procita ponovo
            PalAkcija = OtvoriStavke(red)
            Exit Function
        Case "palnepotpune"
            ' jedina radnja bez reda: stampa SVE nepotpune palete
            modOtkupUI.ShowToast Poruka("OTKUI_MSG_PAL_NEPOTPUNE") & " " & _
                                 CStr(PrintNepotpunePalete()), False
            Exit Function
    End Select

    iD = IdZaRed(p(0), broj)
    If Len(iD) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_NEMA_REDA"), True
        Exit Function
    End If

    Select Case p(0)
        Case "palprint"
            PrintPaletniList iD
            modOtkupUI.ShowToast Poruka("OTKUI_MSG_STAMPA") & " " & broj, False

        Case "palpdf"
            ExportPaletniListPDF iD, True
            modOtkupUI.ShowToast Poruka("OTKUI_MSG_STAMPA") & " " & broj, False

        Case "palzatvori"
            If MsgBox(Poruka("OTKUI_ASK_PAL_ZATVORI") & " " & broj & "?", _
                      vbQuestion + vbYesNo, APP_NAME) = vbNo Then Exit Function
            ClosePaletaManual_TX iD
            modOtkupUI.ShowToast Poruka("OTKUI_MSG_PAL_ZATVORENA") & " " & broj, False
            PalAkcija = True

        Case "palstorno"
            ' Preradjena paleta se ne stornira direktno - modStorno to odbija
            ' tiho. Isto upozorenje kao u legacy formi, samo pre poziva.
            If UCase$(Trim$(CStr(modOtkupUI.GridCell(red, PAL_KOL_PRERADJENO)))) = "DA" Then
                modOtkupUI.ShowToast Poruka("OTKUI_ERR_PAL_PRERADJENA"), True
                Exit Function
            End If
            If MsgBox(Poruka("OTKUI_ASK_PAL_STORNO") & " " & broj & "?", _
                      vbQuestion + vbYesNo, APP_NAME) = vbNo Then Exit Function
            If StornoPaleta_TX(iD) Then
                modOtkupUI.ShowToast Poruka("OTKUI_MSG_PAL_STORNIRANA") & " " & broj, False
                PalAkcija = True
            Else
                modOtkupUI.ShowToast Poruka("OTKUI_ERR_STORNO") & " " & broj, True
            End If

        Case "preprint"
            OutputPreradaList iD
            modOtkupUI.ShowToast Poruka("OTKUI_MSG_STAMPA") & " " & broj, False

        Case "prestorno"
            If MsgBox(Poruka("OTKUI_ASK_PRE_STORNO") & " " & broj & "?" & vbCrLf & _
                      Poruka("OTKUI_ASK_PRE_STORNO2"), vbQuestion + vbYesNo, _
                      APP_NAME) = vbNo Then Exit Function
            If StornoPrerada_TX(iD) Then
                modOtkupUI.ShowToast Poruka("OTKUI_MSG_PRE_STORNIRANA") & " " & broj, False
                PalAkcija = True
            Else
                modOtkupUI.ShowToast Poruka("OTKUI_ERR_STORNO") & " " & broj, True
            End If

        Case Else
            modOtkupUI.ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & p(0), True
    End Select
    Exit Function
EH:
    modOtkupUI.ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & Err.description, True
End Function

' Broj iz prve kolone -> ID, po tome kojoj listi radnja pripada.
Private Function IdZaRed(ByVal akcija As String, ByVal broj As String) As String
    If Len(broj) = 0 Then Exit Function
    If Left$(akcija, 3) = "pre" Then
        If mPreIds Is Nothing Then Exit Function
        If mPreIds.Exists(broj) Then IdZaRed = CStr(mPreIds(broj))
    Else
        If mPalIds Is Nothing Then Exit Function
        If mPalIds.Exists(broj) Then IdZaRed = CStr(mPalIds(broj))
    End If
End Function

'--------------------------------------------------------------- REDOVI
Public Function Scr_Rows(ByVal filter As String, ByVal q As String) As Variant
    Select Case Scr_Lista()
        Case "STAVKE":  Scr_Rows = RowsStavke(q): Exit Function
        Case "PRERADE": Scr_Rows = RowsPrerade(q): Exit Function
    End Select
    If Scr_Lista() = ST_NOVA Then
        Scr_Rows = RowsNovaPrerada(q)
        Exit Function
    End If
    Scr_Rows = RowsPalete(filter, q)
End Function

' Postavi AKTIVNU paletu iz reda mreze. Vraca False ako red ne nosi paletu
' (prazna mreza, red van skupa) -- pozivalac tada ne sme nista da menja.
Private Function PostaviAktivnu(ByVal red As Long) As Boolean
    Dim broj As String
    On Error Resume Next
    broj = Trim$(CStr(modOtkupUI.GridCell(red, 1)))
    If Len(broj) = 0 Then Exit Function
    If mPalIds Is Nothing Then Exit Function
    If Not mPalIds.Exists(broj) Then Exit Function
    mPalID = CStr(mPalIds(broj))
    mPalBroj = broj
    ' vrsta, sorta, klasa i status stoje u redu koji je upravo izabran
    mPalOpis = Trim$(CStr(modOtkupUI.GridCell(red, 3))) & "  " & _
               ChrW(183) & "  " & Trim$(CStr(modOtkupUI.GridCell(red, 4))) & _
               "  " & ChrW(183) & "  " & _
               Trim$(CStr(modOtkupUI.GridCell(red, 11)))
    RefreshAktivna
    PostaviAktivnu = True
End Function

' Otvori stavke izabrane palete: aktivna paleta pa prebacaj liste. Zona i dalje
' pokazuje KOJA je paleta otvorena, pa se sa liste stavki zna gde se stoji.
Private Function OtvoriStavke(ByVal red As Long) As Boolean
    If Not PostaviAktivnu(red) Then Exit Function
    If Scr_Lista() = ST_NOVA Then OcistiPreradu
    mLista = "STAVKE"
    OtvoriStavke = True
End Function

' Kolone deljene mreze: KLJUC_KATALOGA | izvor | vrsta | sirina | prio
'--- UNOS PRERADE ----------------------------------------------------
' Zona ekrana; Nothing dok forma jos nije izgradjena (test) -- tada se ne crta.
Private Function Zona() As Object
    On Error Resume Next
    Set Zona = modOtkupUI.ScreenZone("PALETE")
End Function

Private Function OznaciPaletu(ByVal red As Long) As Boolean
    Dim ident As String
    On Error Resume Next
    ident = Trim$(CStr(modOtkupUI.GridCell(red, PAL_NOVA_COL_ID)))
    If Len(ident) = 0 Then Exit Function
    If mPreOznacene Is Nothing Then Set mPreOznacene = CreateObject("Scripting.Dictionary")
    If mPreOznacene.Exists(ident) Then
        mPreOznacene.Remove ident
    Else
        mPreOznacene(ident) = True
    End If
    OsveziNeto
    OznaciPaletu = True
End Function

' Vrednost polja zone. Ekran cita svoja polja kroz ljusku, kao sto i mrezu cita
' kroz GridCell -- kontrole su njegove, ali ih drzi forma.
Private Function PoljeP(ByVal nm As String) As String
    Dim z As Object
    On Error Resume Next
    Set z = Zona()
    If z Is Nothing Then Exit Function
    PoljeP = Trim$(CStr(z.Controls(nm).Controls(nm & "T").value))
End Function

' NETO = bruto - tezina palete - ambalaza. Isti racun koji legacy radi pred
' upisom; ovde stoji uzivo, pa se greska u unosu vidi odmah a ne posle potvrde.
' Racun je jedan i deli ga i sam upis -- da se prikaz i upisana vrednost ne mogu
' razici.
' RACUN je izdvojen iz PRIKAZA, po pravilu iz par.6 kataloga: polja cita jedna
' rutina, racuna druga. Bez toga se neto ne bi mogao izmeriti -- zona se crta nad
' formom koju harness gradi bez .Show, pa bi jedina poslovna formula na ovom
' ekranu ostala nepokrivena.
Public Function NetoIzracun(ByVal bruto As Double, ByVal tezPal As Double, _
                            ByVal brKut As Double, ByVal tipKut As String, _
                            ByVal brKes As Double, ByVal tipKes As String) As Double
    Dim amb As Double
    On Error Resume Next
    amb = brKut * GetTezinaKutije(tipKut) + brKes * GetTezinaKese(tipKes)
    NetoIzracun = bruto - tezPal - amb
    ' Negativan neto nije podatak nego znak da unos jos nije potpun; nula je
    ' iskrenija od minusa, a validacija ionako ne pusta prazna polja.
    If NetoIzracun < 0 Then NetoIzracun = 0
End Function

Private Function NetoPrerade() As Double
    On Error Resume Next
    NetoPrerade = NetoIzracun( _
        Val(Replace(PoljeP("scrPreBruto"), ",", ".")), _
        Val(Replace(PoljeP("scrPreTezPal"), ",", ".")), _
        Val(Replace(PoljeP("scrPreKut"), ",", ".")), PoljeP("scrPreTipKut"), _
        Val(Replace(PoljeP("scrPreKes"), ",", ".")), PoljeP("scrPreTipKes"))
End Function

Private Sub OsveziNeto()
    Dim z As Object
    On Error Resume Next
    Set z = Zona()
    If z Is Nothing Then Exit Sub
    z.Controls("preNetoV").caption = Format$(NetoPrerade(), "#,##0.##")
    z.Controls("preIzbor").caption = SpisakIzabranih()
End Sub

' Prazan izbor i prazna polja posle upisa. Bez ovoga bi sledeca prerada krenula
' sa vec potrosenim izborom -- ista lekcija kao kod oznacenih otkupnih blokova.
' Koje su palete izabrane, po BROJU. Brojka sama ne kaze nista -- operater je
' prijavio da nigde ne pise koje palete sveze robe ulaze u preradu, a bas to
' je odluka koju donosi. Dugacak spisak se skracuje: zona ima jedan red, a
' poenta je prepoznavanje, ne inventar.
Private Function SpisakIzabranih() As String
    Dim k As Variant, n As Long, spisak As String, broj As String
    On Error Resume Next
    n = PalOznacenihBroj()
    If n = 0 Then SpisakIzabranih = Poruka("OTKUI_PRE_NIJEDNA"): Exit Function
    For Each k In mPreOznacene.keys
        broj = CStr(k)
        If Not mPreBrojevi Is Nothing Then
            If mPreBrojevi.Exists(CStr(k)) Then broj = CStr(mPreBrojevi(CStr(k)))
        End If
        If Len(spisak) > 46 Then
            spisak = spisak & ChrW(8230)
            Exit For
        End If
        If Len(spisak) > 0 Then spisak = spisak & ", "
        spisak = spisak & broj
    Next k
    SpisakIzabranih = n & " " & Poruka("OTKUI_PRE_IZABRANO") & ":  " & spisak
End Function

Private Sub OcistiPreradu()
    Dim z As Object, nm As Variant
    On Error Resume Next
    Set mPreOznacene = Nothing
    Set z = Zona()
    If z Is Nothing Then Exit Sub
    For Each nm In Array("scrPreBruto", "scrPreTezPal", "scrPreGP", "scrPreNap", _
                         "scrPreKut", "scrPreTipKut", "scrPreKes", "scrPreTipKes")
        z.Controls(CStr(nm)).Controls(CStr(nm) & "T").value = ""
    Next nm
    OsveziNeto
End Sub

' Prva provera koja padne vraca svoju poruku; prazan string znaci da je sve u
' redu. Svih sedam su prenete iz legacy panela i i dalje su pod prekidacem
' VALIDACIJA_UNOSA -- nijedna nije nova kapija.
Private Function RazlogNepotpuneP() As String
    If Not IsValidacijaUnosa() Then Exit Function
    If Val(Replace(PoljeP("scrPreBruto"), ",", ".")) <= 0 Then RazlogNepotpuneP = Poruka("OTKUI_PRE_V_BRUTO"): Exit Function
    If Val(Replace(PoljeP("scrPreTezPal"), ",", ".")) <= 0 Then RazlogNepotpuneP = Poruka("OTKUI_PRE_V_TEZPAL"): Exit Function
    If Len(PoljeP("scrPreGP")) = 0 Then RazlogNepotpuneP = Poruka("OTKUI_PRE_V_GP"): Exit Function
    If Val(PoljeP("scrPreKut")) <= 0 Then RazlogNepotpuneP = Poruka("OTKUI_PRE_V_KUT"): Exit Function
    If Len(PoljeP("scrPreTipKut")) = 0 Then RazlogNepotpuneP = Poruka("OTKUI_PRE_V_TIPKUT"): Exit Function
    If Val(PoljeP("scrPreKes")) <= 0 Then RazlogNepotpuneP = Poruka("OTKUI_PRE_V_KES"): Exit Function
    If Len(PoljeP("scrPreTipKes")) = 0 Then RazlogNepotpuneP = Poruka("OTKUI_PRE_V_TIPKES"): Exit Function
End Function

' Ekran ne racuna i ne upisuje sam: spisak i polja idu u SavePrerada_TX, isti
' writer koji zove i legacy panel.
Private Function PreradiIzabrane() As Boolean
    Dim ids As Collection, k As Variant, razlog As String, preID As String
    Dim errDesc As String
    On Error GoTo EH
    If PalOznacenihBroj() = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_PRE_NEMA_IZBORA"), True
        Exit Function
    End If
    razlog = RazlogNepotpuneP()
    If Len(razlog) > 0 Then
        MsgBox razlog, vbExclamation, APP_NAME
        Exit Function
    End If

    If MsgBox(Poruka("OTKUI_PRE_ASK") & vbCrLf & vbCrLf & _
              PalOznacenihBroj() & " " & Poruka("OTKUI_PRE_IZABRANO") & vbCrLf & _
              Poruka("OTKUI_PRE_ASK2"), _
              vbExclamation + vbYesNo + vbDefaultButton2, APP_NAME) <> vbYes Then Exit Function

    Set ids = New Collection
    For Each k In mPreOznacene.keys
        ids.Add CStr(k)
    Next k

    preID = SavePrerada_TX(ids, _
                           CLng(Val(PoljeP("scrPreKut"))), _
                           CLng(Val(PoljeP("scrPreKes"))), _
                           NetoPrerade(), _
                           PoljeP("scrPreNap"), _
                           Val(Replace(PoljeP("scrPreTezPal"), ",", ".")), _
                           Val(Replace(PoljeP("scrPreBruto"), ",", ".")), _
                           0, _
                           PoljeP("scrPreTipKut"), _
                           PoljeP("scrPreTipKes"), _
                           PoljeP("scrPreGP"))
    If Len(preID) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_PRE_ERR"), True
        Exit Function
    End If

    OutputPreradaList preID
    OcistiPreradu
    modOtkupUI.ShowToast Poruka("OTKUI_PRE_SACUVANA") & " " & preID, False
    PreradiIzabrane = True
    Exit Function
EH:
    errDesc = Err.description
    LogErr "modScrPalete.PreradiIzabrane"
    Err.Clear
    modOtkupUI.ShowToast Poruka("OTKUI_PRE_ERR") & " " & errDesc, True
End Function

'--- NOVA PRERADA ----------------------------------------------------
' Isti izvor kao lista Palete, sa dve dodate kolone: kvacica napred i PaletaID
' pozadi (nevidljiv, prioritet 4). Izbor se drzi po ID-u, ne po broju palete:
' spisak zavrsava u SavePrerada_TX, dakle u mutaciji.
'
' Ne filtrira se unapred na 'nepreradjene' -- legacy prikazuje sve i pusta
' operatera da bira, a suzavanje ide kroz pretragu.
Private Function NovaGridCols() As Variant
    NovaGridCols = Array( _
        "OTKUI_HDP_OZN||txt|34|1", _
        "OTKUI_HDP_BROJ||txt|64|1", _
        "OTKUI_HDP_GODINA||txt|50|3", _
        "OTKUI_HD_VRSTA||txt|84|2", _
        "OTKUI_HD_SORTA||part|0|2", _
        "OTKUI_HD_KLASA||txt|44|3", _
        "OTKUI_HD_TIP_AMB||txt|66|3", _
        "OTKUI_HDP_GAJBICA||num|58|1", _
        "OTKUI_HDP_KAPACITET||num|58|2", _
        "OTKUI_HDP_NETO||kg|72|1", _
        "OTKUI_HDP_BRUTO||kg|72|2", _
        "OTKUI_HD_STATUS||txt|78|1", _
        "OTKUI_HDP_PRERADJENO||txt|76|2", _
        "OTKUI_HD_IDENT||txt|0|4")
End Function

Private Function RowsNovaPrerada(ByVal q As String) As Variant
    Dim src As Variant, r As Long, n As Long, outA() As Variant
    Dim hay As String, sumNeto As Double, uk As Long, st As String, ident As String
    On Error GoTo EH
    mStep = "nova prerada"
    If mPreBrojevi Is Nothing Then Set mPreBrojevi = CreateObject("Scripting.Dictionary")
    src = GetPaleteForGrid()
    If Not IsArray(src) Then
        RowsNovaPrerada = Array(NovaGridCols(), Empty, 0, 0#, 0#, Array(0, 0, 0))
        Exit Function
    End If

    uk = UBound(src, 1) + 1
    ReDim outA(1 To uk, 1 To PAL_NOVA_COL_ID)
    For r = 0 To UBound(src, 1)
        st = CStr(src(r, 11))
        ident = CStr(src(r, 0))
        mPreBrojevi(ident) = CStr(src(r, 1))
        hay = CStr(src(r, 1)) & "|" & CStr(src(r, 2)) & "|" & CStr(src(r, 3)) & "|" & _
              CStr(src(r, 4)) & "|" & CStr(src(r, 5)) & "|" & CStr(src(r, 6)) & "|" & _
              st & "|" & CStr(src(r, 12))
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeci
        End If
        n = n + 1
        outA(n, 1) = IIf(PalOznacena(ident), ChrW(10003), "")
        outA(n, 2) = CStr(src(r, 1))
        outA(n, 3) = CStr(src(r, 2))
        outA(n, 4) = CStr(src(r, 3))
        outA(n, 5) = CStr(src(r, 4))
        outA(n, 6) = CStr(src(r, 5))
        outA(n, 7) = CStr(src(r, 6))
        outA(n, 8) = Val(CStr(src(r, 7)))
        outA(n, 9) = Val(CStr(src(r, 8)))
        outA(n, 10) = PalD(src(r, 9))
        outA(n, 11) = PalD(src(r, 10))
        outA(n, 12) = st
        outA(n, 13) = CStr(src(r, 12))
        outA(n, PAL_NOVA_COL_ID) = ident
        sumNeto = sumNeto + PalD(src(r, 9))
Sledeci:
    Next r
    mStep = "OK"
    RowsNovaPrerada = Array(NovaGridCols(), outA, n, sumNeto, 0#, Array(0, 0, 0))
    Exit Function
EH:
    Err.Raise Err.Number, "modScrPalete.RowsNovaPrerada[" & mStep & "]", Err.description
End Function

Private Function PalOznacena(ByVal ident As String) As Boolean
    If mPreOznacene Is Nothing Then Exit Function
    PalOznacena = mPreOznacene.Exists(ident)
End Function

' Koliko ih je oznaceno. Zona i potvrda citaju isti broj -- da se ne razidju.
Public Function PalOznacenihBroj() As Long
    If mPreOznacene Is Nothing Then Exit Function
    PalOznacenihBroj = mPreOznacene.count
End Function

' Test seam: aktivna lista bez prekidaca. Tvrdo gejtovan.
Public Sub Scr_PalTestSet(ByVal kljuc As String)
    If Not IsTestMode() Then Exit Sub
    If kljuc <> mLista Then OcistiPreradu
    mLista = kljuc
End Sub

' Test seam: izbor palete bez mreze. Tvrdo gejtovan, kao Scr_OtpTestSet.
Public Sub Scr_PreTestSet(ByVal ident As String)
    If Not IsTestMode() Then Exit Sub
    If mPreOznacene Is Nothing Then Set mPreOznacene = CreateObject("Scripting.Dictionary")
    If mPreOznacene.Exists(ident) Then mPreOznacene.Remove ident Else mPreOznacene(ident) = True
End Sub

Private Function PalGridCols() As Variant
    PalGridCols = Array( _
        "OTKUI_HDP_BROJ||txt|64|1", _
        "OTKUI_HDP_GODINA||txt|50|3", _
        "OTKUI_HD_VRSTA||txt|84|2", _
        "OTKUI_HD_SORTA||part|0|2", _
        "OTKUI_HD_KLASA||txt|44|3", _
        "OTKUI_HD_TIP_AMB||txt|66|3", _
        "OTKUI_HDP_GAJBICA||num|58|1", _
        "OTKUI_HDP_KAPACITET||num|58|2", _
        "OTKUI_HDP_NETO||kg|72|1", _
        "OTKUI_HDP_BRUTO||kg|72|2", _
        "OTKUI_HD_STATUS||txt|78|1", _
        "OTKUI_HDP_PRERADJENO||txt|76|2")
End Function

' Palete iz modPaletniList.GetPaleteForGrid (0-bazirano, 13 kolona):
'   0 PaletaID | 1 Broj | 2 Godina | 3 Vrsta | 4 Sorta | 5 Klasa | 6 TipAmb
'   7 Gajbice | 8 Kapacitet | 9 Neto | 10 Bruto | 11 Status | 12 Preradjeno
'
' Filteri legacy forme (godina, vrsta, sorta, status, preradjeno) ovde rade
' kroz JEDNU pretragu: sve te vrednosti su kolone, pa se kucanjem "OTVORENA"
' ili "Willamette" dobija isti rez, a kolone se uz to mogu i sortirati.
Private Function RowsPalete(ByVal filter As String, ByVal q As String) As Variant
    Dim src As Variant, r As Long, n As Long, outA() As Variant
    Dim hay As String, sumNeto As Double, gajbi As Double, otvorene As Long
    Dim uk As Long, st As String
    On Error GoTo EH
    mStep = "palete"

    Set mPalIds = CreateObject("Scripting.Dictionary")
    src = GetPaleteForGrid()
    If Not IsArray(src) Then
        RefreshBrojke 0, 0, 0
        RowsPalete = Array(PalGridCols(), Empty, 0, 0#, 0#, Array(0, 0, 0))
        Exit Function
    End If

    uk = UBound(src, 1) + 1
    ReDim outA(1 To uk, 1 To 12)
    For r = 0 To UBound(src, 1)
        st = CStr(src(r, 11))
        ' zbirovi u zoni idu preko SVIH paleta, ne preko filtrirane liste
        gajbi = gajbi + Val(CStr(src(r, 7)))
        If UCase$(st) <> "ZATVORENA" Then otvorene = otvorene + 1

        ' Cip suzava listu PRE pretrage. Zbirovi u zoni ostaju preko svih
        ' paleta -- oni govore o stanju hladnjace, ne o tome sta je na ekranu.
        If Not PalCipProlaz(filter, st, CStr(src(r, 2)), CStr(src(r, 12))) _
            Then GoTo Sledeci

        hay = CStr(src(r, 1)) & "|" & CStr(src(r, 2)) & "|" & CStr(src(r, 3)) & "|" & _
              CStr(src(r, 4)) & "|" & CStr(src(r, 5)) & "|" & CStr(src(r, 6)) & "|" & _
              st & "|" & CStr(src(r, 12))
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeci
        End If

        n = n + 1
        outA(n, 1) = CStr(src(r, 1))
        mPalIds(CStr(outA(n, 1))) = CStr(src(r, 0))
        outA(n, 2) = CStr(src(r, 2))
        outA(n, 3) = CStr(src(r, 3))
        outA(n, 4) = CStr(src(r, 4))
        outA(n, 5) = CStr(src(r, 5))
        outA(n, 6) = CStr(src(r, 6))
        outA(n, 7) = Val(CStr(src(r, 7)))
        outA(n, 8) = Val(CStr(src(r, 8)))
        outA(n, 9) = PalD(src(r, 9))
        outA(n, 10) = PalD(src(r, 10))
        outA(n, 11) = st
        outA(n, 12) = CStr(src(r, 12))
        sumNeto = sumNeto + PalD(src(r, 9))
Sledeci:
    Next r

    RefreshBrojke uk, otvorene, gajbi
    mStep = "OK"
    RowsPalete = Array(PalGridCols(), outA, n, sumNeto, 0#, Array(0, 0, 0))
    Exit Function
EH:
    Err.Raise Err.Number, "modScrPalete.RowsPalete[" & mStep & "]", Err.description
End Function

'--------------------------------------------------------- LISTA: STAVKE
Private Function StavkeGridCols() As Variant
    StavkeGridCols = Array( _
        "OTKUI_HDS_PRIJEMNICA||txt|120|1", _
        "OTKUI_HD_BROJ_ZBIRNE||part|0|2", _
        "OTKUI_HDP_GAJBICA||num|70|1", _
        "OTKUI_HDP_NETO||kg|86|1")
End Function

' Stavke izabrane palete (0-bazirano): 0 PrijemnicaID | 1 BrojPrijemnice
' | 2 BrojZbirne | 3 Gajbice | 4 Neto
Private Function RowsStavke(ByVal q As String) As Variant
    Dim src As Variant, r As Long, n As Long, outA() As Variant
    Dim hay As String, sumNeto As Double, sumGajbi As Double
    On Error GoTo EH
    mStep = "stavke"

    If Len(mPalID) = 0 Then
        RowsStavke = Array(StavkeGridCols(), Empty, 0, 0#, 0#, Array(0, 0, 0))
        Exit Function
    End If
    src = GetPaletaStavkeForGrid(mPalID)
    If Not IsArray(src) Then
        RowsStavke = Array(StavkeGridCols(), Empty, 0, 0#, 0#, Array(0, 0, 0))
        Exit Function
    End If

    ReDim outA(1 To UBound(src, 1) + 1, 1 To 4)
    For r = 0 To UBound(src, 1)
        hay = CStr(src(r, 1)) & "|" & CStr(src(r, 2))
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeci
        End If
        n = n + 1
        outA(n, 1) = CStr(src(r, 1))
        outA(n, 2) = CStr(src(r, 2))
        outA(n, 3) = Val(CStr(src(r, 3)))
        outA(n, 4) = PalD(src(r, 4))
        sumGajbi = sumGajbi + Val(CStr(src(r, 3)))
        sumNeto = sumNeto + PalD(src(r, 4))
Sledeci:
    Next r

    mStep = "OK"
    RowsStavke = Array(StavkeGridCols(), outA, n, sumNeto, 0#, Array(0, 0, 0))
    Exit Function
EH:
    Err.Raise Err.Number, "modScrPalete.RowsStavke[" & mStep & "]", Err.description
End Function

'-------------------------------------------------------- LISTA: PRERADE
Private Function PreGridCols() As Variant
    PreGridCols = Array( _
        "OTKUI_HDP_BROJ||txt|70|1", _
        "OTKUI_HD_DATUM||date|70|1", _
        "OTKUI_HDR_TIPGP||part|0|1", _
        "OTKUI_HDR_KUTIJE||num|66|2", _
        "OTKUI_HDR_KESE||num|66|2", _
        "OTKUI_HDP_NETO||kg|86|1")
End Function

' Prerade (0-bazirano): 0 PreradaID | 1 Broj | 2 Datum | 3 Neto | 4 Kutije
' | 5 Kese | 6 TipGotovogProizvoda
Private Function RowsPrerade(ByVal q As String) As Variant
    Dim src As Variant, r As Long, n As Long, outA() As Variant
    Dim hay As String, sumNeto As Double
    On Error GoTo EH
    mStep = "prerade"

    Set mPreIds = CreateObject("Scripting.Dictionary")
    src = GetPreradeForGrid()
    If Not IsArray(src) Then
        RowsPrerade = Array(PreGridCols(), Empty, 0, 0#, 0#, Array(0, 0, 0))
        Exit Function
    End If

    ReDim outA(1 To UBound(src, 1) + 1, 1 To 6)
    For r = 0 To UBound(src, 1)
        hay = CStr(src(r, 1)) & "|" & CStr(src(r, 2)) & "|" & CStr(src(r, 6))
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeci
        End If
        n = n + 1
        outA(n, 1) = CStr(src(r, 1))
        mPreIds(CStr(outA(n, 1))) = CStr(src(r, 0))
        outA(n, 2) = PalDatum(src(r, 2))
        outA(n, 3) = CStr(src(r, 6))
        outA(n, 4) = Val(CStr(src(r, 4)))
        outA(n, 5) = Val(CStr(src(r, 5)))
        outA(n, 6) = PalD(src(r, 3))
        sumNeto = sumNeto + PalD(src(r, 3))
Sledeci:
    Next r

    mStep = "OK"
    RowsPrerade = Array(PreGridCols(), outA, n, sumNeto, 0#, Array(0, 0, 0))
    Exit Function
EH:
    Err.Raise Err.Number, "modScrPalete.RowsPrerade[" & mStep & "]", Err.description
End Function

'---------------------------------------------------------------- ZONA
' Tri brojke i aktivna paleta. Zona se gradi jednom, pa se puni tek kad se
' citaju podaci - a to je u Scr_Rows.
Private Sub RefreshBrojke(ByVal uk As Long, ByVal otvorene As Long, ByVal gajbi As Double)
    Dim z As Object
    On Error Resume Next
    Set z = modOtkupUI.ScreenZone("PALETE")
    If z Is Nothing Then Exit Sub
    z.Controls("palKL0").caption = UCase$(Poruka("OTKUI_SCRPAL_UKUPNO"))
    z.Controls("palKV0").caption = Format$(uk, "#,##0")
    z.Controls("palKL1").caption = UCase$(Poruka("OTKUI_SCRPAL_OTVORENE"))
    z.Controls("palKV1").caption = Format$(otvorene, "#,##0")
    z.Controls("palKL2").caption = UCase$(Poruka("OTKUI_SCRPAL_GAJBICA"))
    z.Controls("palKV2").caption = Format$(gajbi, "#,##0")
    RefreshAktivna
End Sub

Private Sub RefreshAktivna()
    Dim z As Object, sub_ As String
    On Error Resume Next
    Set z = modOtkupUI.ScreenZone("PALETE")
    If z Is Nothing Then Exit Sub
    If Len(mPalBroj) = 0 Then
        z.Controls("palBroj").caption = ChrW(8212)
        z.Controls("palSub").caption = Poruka("OTKUI_PAL_NEMA")
        Exit Sub
    End If
    z.Controls("palBroj").caption = mPalBroj
    ' opis aktivne palete se cita iz mreze samo ako je lista paleta prikazana;
    ' u ostalim listama ostaje ono sto je zapisano pri izboru
    sub_ = mPalOpis
    If Len(sub_) = 0 Then sub_ = Poruka("OTKUI_PAL_IZABRANA")
    z.Controls("palSub").caption = sub_
End Sub

Public Sub Scr_ResetCache()
    Set mPalIds = Nothing
    Set mPreIds = Nothing
End Sub

'-------------------------------------------------------------- CELIJE
Private Function PalD(ByVal v As Variant) As Double
    On Error Resume Next
    If IsNumeric(v) Then
        PalD = CDbl(v)
    Else
        PalD = Val(Replace(CStr(v), ",", "."))
    End If
End Function

Private Function PalDatum(ByVal v As Variant) As Double
    On Error Resume Next
    If IsDate(v) Then
        PalDatum = Int(CDbl(CDate(v)))
    ElseIf IsNumeric(v) Then
        PalDatum = Int(CDbl(v))
    End If
End Function
