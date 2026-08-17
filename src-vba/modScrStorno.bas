Attribute VB_Name = "modScrStorno"
'=====================================================================
' modScrStorno - ekran "Storno" (v6-ui-142).
'
' Ljuska ga ne poznaje po imenu: dobija ga preko Application.Run, da
' klijent kome ovaj modul nedostaje i dalje radi (zamka #19).
'
' ODAKLE DOLAZI: storno je do v6-ui-141 bio REZIM F8 unosnog ekrana. To je
' imalo dve posledice, obe vidljive operateru:
'
'   1. F8 je crtao unosnu formu koju ne koristi. Scr_Save za "STORNO" je
'      padao u Case Else i vracao "Nije vezano na postojecu rutinu" -- dakle
'      primarno dugme je bilo MRTVO, a forma je pozivala operatera da ukuca
'      podatke dokumenta koji hoce da stornira. #201 je to sakrio grid-maxom;
'      ovde je reseno tako sto forme vise nema.
'   2. Pregled posledica je bio niz MsgBox-ova. Legacy ekran "Storno /
'      potvrda" (frmDokumenta) pokazuje CEO LANAC, palete i pojedinacne
'      blokove PRE nego sto se bilo sta desi. MsgBox to ne moze -- i katalog
'      je tu odluku zabelezio kao "cetiri odgovora ne staju u jedan MsgBox".
'      Ta odluka se ovim menja: cetiri odgovora su cetiri DUGMETA, svako sa
'      objasnjenjem ispod, a iznad njih stoje posledice.
'
' STA JE OVDE, A STA NIJE: ovde je REDOSLED pitanja i ono sto se pokazuje.
' Nijedno poslovno pravilo, nijedna kapija i nijedan upis nisu ovde -- sve
' racuna modStornoFlow kroz modStornoDok, a uvid sklapa modStornoImpact.
' Redovi liste dolaze iz modScrDokumenti.RedoviZaTip, istog citaca koji puni
' i unosni ekran; kopije nema.
'
' IDENTITET. Lista nosi NEVIDLJIVU kolonu kanonskog identiteta (GeneracijaID),
' koju modScrDokumenti.GridCols dodaje samo kad se trazi -- a trazi je samo
' ovaj ekran. Sve nizvodno (uvid, izbor moda, izvrsenje, dodatni storno
' blokova) ide po TOM identitetu, ne po broju. Broj nije jedinstven, i to
' nije teorija: GenerateBrojPrijemnice nema proveru jedinstvenosti. Devet
' rundi pregleda u #198 je tu putanju kalilo -- ne sme da regresira.
'
' STA NIJE PRENETO IZ LEGACY-JA (svesno, Korak 2):
'   - tri tabele odjednom: ljuska ima JEDNU mrezu, pa lanac i blokovi idu u
'     zonu kao labele, a ne kao zasebne liste;
'   - multiselect otkupnih blokova: mreza bira JEDAN red, pa je dodatni
'     storno blokova SVE-ILI-NISTA, isto kao i do sada.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const SCRST_BUILD As String = "v6-ui-142"

' Visina zone. Veca je od KPI_H jer zona nosi ceo uvid: zaglavlje dokumenta,
' tabelu efekta po modu, palete i red odluke.
Private Const ST_ZONA_H As Single = 196
Private Const ST_EF_REDOVA As Long = 5      ' redova tabele "efekat po modu"
Private Const ST_PAL_REDOVA As Long = 3     ' redova detalja paleta

' Kljuc navigacione liste. Nije tip dokumenta nego pogled PREKO tipova, pa ne
' moze da se sudari sa STIP_* kljucevima.
Private Const ST_SVI As String = "SVI"

' Aktivna lista: STIP_* kljuc ili ST_SVI.
Private mLista As String

' IZABRAN DOKUMENT. Broj je ono sto operater vidi; docID je ono po cemu se
' nizvodno bira. Drze se zajedno i cuvaju se ZAJEDNO -- prazan docID uz
' neprazan broj je legitimno stanje (zatecen zapis bez generacije), i tada
' nizvodno vazi fail-closed kapija nad jednoznacnoscu broja.
Private mSelTip As String
Private mSelBroj As String
Private mSelDocID As String
' Ono sto se iz broja ne vidi: smer reversa, odnosno racun izvoda.
Private mSelOpcija As String

' "Ne diraj palete" -- do v6-ui-142 MsgBox (STORNO_ASK_PALETE), sada prekidac
' koji stoji uz posledice, pa se vidi PRE odluke a ne posle nje.
Private mNeDiraj As Boolean

' Uvid za izabran dokument. Gradi ga BuildStornoImpact pri izboru reda, a ne
' pri svakom crtanju: to je sedam citanja tabela.
Private mImpact As Object

Private mStep As String                     ' korak za poruku o gresci

'--------------------------------------------------------- UGOVOR EKRANA
Public Function Scr_Meta() As String
    Scr_Meta = "kljuc=STORNO|naslov=OTKUI_NAV_STORNO|sub=OTKUI_SCRST_SUB" & _
               "|lista=OTKUI_SCRST_LISTA|oblik=lista|upis=ne"
End Function

' Prekidac bira TIP DOKUMENTA -- isti izbor koji legacy trazi kroz
' cmbStornoDokument, samo sto se ovde bira lista pa red u njoj, umesto da se
' broj kuca napamet. Fakture i izvodi su tu iako novi UI nema rezim koji ih
' kreira: stornirati se moraju, a legacy ih ima u istom combo-u.
'
' "Svi" je NAVIGACIONA lista preko tipova, koju legacy ima ("Nadji dokument")
' a novi UI do sada nije. Stoji PRVA jer je to pogled sa kojim se pocinje kad
' se ne zna gde dokument zivi.
Public Function Scr_Liste() As Variant
    Scr_Liste = Array( _
        ST_SVI & "|OTKUI_SEG_ST_SVI|OTKUI_GRID_TITLE_ST_SVI|40", _
        STIP_OTKUP & "|OTKUI_SEG_ST_OTKUP|OTKUI_GRID_TITLE_OTKUP|64", _
        STIP_OTPREMNICA & "|OTKUI_SEG_ST_OTP|OTKUI_GRID_TITLE_OTPREMNICA|72", _
        STIP_ZBIRNA & "|OTKUI_SEG_ST_ZBR|OTKUI_GRID_TITLE_ZBIRNA|56", _
        STIP_PRIJEMNICA & "|OTKUI_SEG_ST_PRJ|OTKUI_GRID_TITLE_PRIJEMNICA|72", _
        STIP_ISPLATE & "|OTKUI_SEG_ST_ISP|OTKUI_GRID_TITLE_AMB_ISPLATE|58", _
        STIP_UPLATE & "|OTKUI_SEG_ST_UPL|OTKUI_GRID_TITLE_AMB_UPLATE|54", _
        STIP_REVERSI & "|OTKUI_SEG_ST_REV|OTKUI_GRID_TITLE_REVERSI|58", _
        STIP_FAKTURA & "|OTKUI_SEG_ST_FAK|OTKUI_GRID_TITLE_FAKTURA|58", _
        STIP_IZVOD & "|OTKUI_SEG_ST_IZV|OTKUI_GRID_TITLE_IZVOD|54")
End Function

' Podrazumevan je otkupni list - najcesci dokument, pa i najcesci storno.
Public Function Scr_Lista() As String
    If Len(mLista) = 0 Then mLista = STIP_OTKUP
    Scr_Lista = mLista
End Function

' Radnji NAD REDOM nema. Odluka se donosi u zoni, gde stoje i posledice --
' dugme u redu mreze bi vodilo pravo u izvrsenje, a ceo smisao ovog ekrana je
' da se posledice vide PRE odluke.
Public Function Scr_Radnje() As String
End Function

' U zaglavlju liste stoji izabran dokument, da se vidi na sta se odnosi zona.
Public Function Scr_NaslovDopuna() As String
    Scr_NaslovDopuna = mSelBroj
End Function

' TEST SEAM: izbor tipa bez klika po prekidacu. Tvrdo je gejtovan -- van
' test-rezima ne radi nista, pa se ne moze upotrebiti kao precica u produkciji.
Public Sub Scr_TipTestSet(ByVal tipKljuc As String)
    If Not IsTestMode() Then Exit Sub
    mLista = tipKljuc
End Sub

' TEST SEAM: izbor dokumenta bez ucitane mreze. Produkcija ga bira klikom na
' red (Scr_Event "row:N"), sto trazi mrezu, filtere i strane -- previse za
' proveru pravila da identitet stize do writer-a.
Public Sub Scr_IzborTestSet(ByVal tip As String, ByVal broj As String, _
                            ByVal docID As String, ByVal opcija As String)
    If Not IsTestMode() Then Exit Sub
    mSelTip = tip
    mSelBroj = broj
    mSelDocID = docID
    mSelOpcija = opcija
End Sub

' TEST SEAM: identitet koji bi ovaj ekran poslao nizvodno. Bez njega se
' "identitet stize do writer-a" ne moze izmeriti -- writer ga primi ili ne
' primi, a oba slucaja iz testa izgledaju isto.
Public Function Scr_IzabranDocID() As String
    Scr_IzabranDocID = mSelDocID
End Function

'--------------------------------------------------------------- REDOVI
' Redovi za deljenu mrezu. Tipizirane liste dolaze iz istog citaca koji puni i
' unosni ekran -- odatle i nevidljiva kolona identiteta (True na kraju).
Public Function Scr_Rows(ByVal filter As String, ByVal q As String) As Variant
    If Scr_Lista() = ST_SVI Then
        Scr_Rows = RowsSvi(q)
        Exit Function
    End If
    Scr_Rows = modScrDokumenti.RedoviZaTip(Scr_Lista(), filter, q, True)
End Function

Private Function SviGridCols() As Variant
    SviGridCols = Array( _
        "OTKUI_HD_TIP||txt|96|1", _
        "OTKUI_HD_BROJ||txt|110|1", _
        "OTKUI_HD_DATUM||date|58|1", _
        "OTKUI_HD_PARTNER||txt|0|1", _
        "OTKUI_HD_KG||kg|60|1")
End Function

' Lista PREKO tipova. Namerno je samo NAVIGACIONA: red joj dolazi iz
' GetActiveDocumentsForStorno, koji pokriva tri framework tipa i NEMA kolonu
' identiteta. Radnja nad takvim redom bi se vratila na biranje po broju --
' tacno ono sto je #198 vadio iz koda. Zato klik na red ovde samo PREBACUJE na
' tipiziranu listu, gde identitet postoji, a odluka se donosi tamo.
Private Function RowsSvi(ByVal q As String) As Variant
    Dim src As Collection, outA() As Variant, n As Long, i As Long
    Dim red As Variant, hay As String, sumKg As Double
    On Error GoTo EH
    mStep = "GetWarmStornoDocs"
    Set src = modStornoWarm.GetWarmStornoDocs()
    If src Is Nothing Then Exit Function
    If src.count = 0 Then
        RowsSvi = Array(SviGridCols(), Empty, 0, 0#, 0#, Array(0, 0, 0))
        Exit Function
    End If

    mStep = "petlja"
    ReDim outA(1 To src.count, 1 To 5)
    n = 0
    For i = 1 To src.count
        ' red = niz(0..7): tip, broj, datum, brojZbirne, kupac, vozac,
        ' otkupnaMesta, kolicina (v. modStornoFlow.GetActiveDocumentsForStorno)
        red = src(i)
        hay = CStr(red(0)) & "|" & CStr(red(1)) & "|" & CStr(red(4)) & "|" & CStr(red(6))
        If Len(q) = 0 Or InStr(1, hay, q, vbTextCompare) > 0 Then
            n = n + 1
            outA(n, 1) = CStr(red(0))
            outA(n, 2) = CStr(red(1))
            outA(n, 3) = SviDatum(red(2))
            outA(n, 4) = SviPartner(red)
            outA(n, 5) = SafeDbl(red(7))
            sumKg = sumKg + SafeDbl(red(7))
        End If
    Next i
    RowsSvi = Array(SviGridCols(), outA, n, sumKg, 0#, Array(0, 0, 0))
    Exit Function
EH:
    Err.Raise Err.Number, "modScrStorno.RowsSvi[" & mStep & "]", Err.description
End Function

' Kupac ako ga dokument ima, inace otkupna mesta (otpremnica nema kupca).
Private Function SviPartner(ByVal red As Variant) As String
    On Error Resume Next
    SviPartner = Trim$(CStr(red(4)))
    If Len(SviPartner) = 0 Then SviPartner = Trim$(CStr(red(6)))
End Function

Private Function SviDatum(ByVal v As Variant) As Double
    On Error Resume Next
    If IsDate(v) Then SviDatum = CDbl(CDate(v))
End Function

Private Function SafeDbl(ByVal v As Variant) As Double
    On Error Resume Next
    If IsNumeric(v) Then SafeDbl = CDbl(v) Else SafeDbl = val(CStr(v))
End Function

'----------------------------------------------------------------- ZONA
' Levo zaglavlje izabranog dokumenta i tabela "efekat po modu"; desno palete i
' prekidac za njih. Dole red odluke: do cetiri dugmeta, svako sa objasnjenjem
' ISPOD sebe -- objasnjenje je ono sto je do sada zivelo u MsgBox tekstu.
Public Sub Scr_Build(ByVal z As Object)
    Dim r As Long, c As Long, i As Long

    modUiKit.NewLbl z, "stCap", UCase$(Poruka("OTKUI_SCRST_IZBOR")), PAD, 6, 220, 11, _
                    TS_MICRO, True, C_MUTED, -1
    modUiKit.NewLbl z, "stDok", Poruka("OTKUI_SCRST_NEMA"), PAD, 18, 460, 20, _
                    TS_KPI, True, C_FOREST, -1
    modUiKit.NewLbl z, "stMeta", "", PAD, 40, 460, 13, TS_META, False, C_MUTED, -1

    modUiKit.NewLbl z, "stEfCap", UCase$(Poruka("OTKUI_SCRST_EFEKAT")), PAD, 58, 260, 11, _
                    TS_MICRO, True, C_MUTED, -1
    For r = 0 To ST_EF_REDOVA - 1
        For c = 0 To 2
            modUiKit.NewLbl z, "stE" & r & "_" & c, "", PAD, 72 + r * 13, 140, 12, _
                            TS_META, (c = 0), IIf(c = 0, C_FOREST, C_MUTED), -1
        Next c
    Next r

    modUiKit.NewLbl z, "stPalCap", UCase$(Poruka("OTKUI_SCRST_PALETE")), 0, 6, 200, 11, _
                    TS_MICRO, True, C_MUTED, -1
    modUiKit.NewLbl z, "stPalZbir", ChrW(8212), 0, 18, 320, 20, TS_KPI, True, C_FOREST, -1
    For i = 0 To ST_PAL_REDOVA - 1
        modUiKit.NewLbl z, "stPal" & i, "", 0, 42 + i * 13, 320, 12, TS_META, False, C_MUTED, -1
    Next i
    ' Prekidac, ne dugme: menja stanje odluke, ne izvrsava je. Boju stanja
    ' postavlja OsveziZonu kroz BoxState, kao cipovi u mrezi.
    modUiKit.BtnV z, "scrStPal", Poruka("OTKUI_SCRST_NEDIRAJ"), 0, 84, 160, 20, "ghost"

    ' Cetiri moda + obican storno + ispravka reversa dele isti red; koji se od
    ' njih vide, odlucuje AkcijeZaTip pri svakom osvezavanju zone.
    For i = 0 To 3
        modUiKit.BtnV z, "scrStA" & i, "", PAD + i * 176, 140, 168, 22, "secondary"
        modUiKit.NewLbl z, "stH" & i, "", PAD + i * 176, 164, 168, 12, _
                        TS_MICRO, False, C_MUTED, -1
    Next i

    modUiKit.NewLbl z, "stLnB", "", 0, ST_ZONA_H - 1, 100, 1, 8, False, 0, C_BORDER
End Sub

Public Function Scr_Layout(ByVal z As Object, ByVal w As Single, ByVal h As Single) As Single
    Dim i As Long, r As Long, c As Long, x2 As Single, wl As Single
    On Error Resume Next
    ' Leva kolona uzima 58% sirine; desna pocinje odatle. Tabela efekta ima tri
    ' kolone: dokument (fiksno), info (fiksno), napomena (uzima ostatak).
    wl = (w - PAD * 2) * 0.58
    x2 = PAD + wl + PAD
    z.Controls("stDok").width = wl
    z.Controls("stMeta").width = wl
    For r = 0 To ST_EF_REDOVA - 1
        For c = 0 To 2
            z.Controls("stE" & r & "_" & c).Left = PAD + c * (wl / 3)
            z.Controls("stE" & r & "_" & c).width = (wl / 3) - 6
        Next c
    Next r

    z.Controls("stPalCap").Left = x2
    z.Controls("stPalZbir").Left = x2
    z.Controls("stPalZbir").width = w - x2 - PAD
    For i = 0 To ST_PAL_REDOVA - 1
        z.Controls("stPal" & i).Left = x2
        z.Controls("stPal" & i).width = w - x2 - PAD
    Next i
    modUiKit.MoveBox z, "scrStPal", x2, 84, 160

    z.Controls("stLnB").width = w
    Scr_Layout = ST_ZONA_H
End Function

' Zona ekrana; Nothing kad forma jos nije izgradjena (test) - tada se ne crta.
Private Function Zona() As Object
    On Error Resume Next
    Set Zona = modOtkupUI.ScreenZone("STORNO")
End Function

'------------------------------------------------------------- OSVEZAVANJE
Private Sub OsveziZonu()
    Dim z As Object
    On Error Resume Next
    Set z = Zona()
    If z Is Nothing Then Exit Sub
    OsveziZaglavlje z
    OsveziEfekat z
    OsveziPalete z
    OsveziAkcije z
End Sub

Private Sub OsveziZaglavlje(ByVal z As Object)
    Dim hd As Object, sm As Object, s As String
    On Error Resume Next
    If mImpact Is Nothing Then
        z.Controls("stDok").caption = Poruka("OTKUI_SCRST_NEMA")
        z.Controls("stMeta").caption = ""
        Exit Sub
    End If
    Set hd = mImpact("header")
    Set sm = mImpact("summary")
    s = modStornoDok.TipNaziv(mSelTip, mSelOpcija) & " " & mSelBroj
    If Len(CStr(hd("partner"))) > 0 Then s = s & "  " & ChrW(183) & "  " & CStr(hd("partner"))
    If Len(CStr(hd("datum"))) > 0 Then s = s & "  " & ChrW(183) & "  " & FmtDat(hd("datum"))
    If Len(CStr(hd("kolicina"))) > 0 Then s = s & "  " & ChrW(183) & "  " & CStr(hd("kolicina")) & " kg"
    z.Controls("stDok").caption = s

    s = ""
    If CLng(sm("blockCount")) > 0 Then _
        s = CStr(sm("blockCount")) & " " & Poruka("OTKUI_SCRST_BLOKOVI")
    If CBool(mImpact("faktura")("hasFaktura")) Then _
        s = s & IIf(Len(s) > 0, "  " & ChrW(183) & "  ", "") & _
            Poruka("OTKUI_SEG_ST_FAK") & " " & CStr(mImpact("faktura")("fakturaID"))
    z.Controls("stMeta").caption = s
End Sub

' Tabela "efekat po modu": redovi lanca iz modStornoFlow.GetStornoChainRows,
' koji su vec sloga [dokument, info, napomena] i vec opisuju sta se desava po
' modu ("UVEK Duplikat pa Ponistenje"). Visak redova preko pet se ne crta --
' ceo lanac koji staje u pet redova pokriva svaki zatecen slucaj, a zona ne
' sme da raste preko mreze.
Private Sub OsveziEfekat(ByVal z As Object)
    Dim ch As Collection, r As Long, c As Long, red As Variant, n As Long
    On Error Resume Next
    For r = 0 To ST_EF_REDOVA - 1
        For c = 0 To 2
            z.Controls("stE" & r & "_" & c).caption = ""
        Next c
    Next r
    If mImpact Is Nothing Then Exit Sub
    Set ch = mImpact("chain")
    If ch Is Nothing Then Exit Sub
    n = ch.count
    If n > ST_EF_REDOVA Then n = ST_EF_REDOVA
    For r = 1 To n
        red = ch(r)
        For c = 0 To 2
            z.Controls("stE" & (r - 1) & "_" & c).caption = ChTxt(red, c)
        Next c
    Next r
End Sub

' Bezbedno citanje polja reda lanca - red moze imati manje od tri polja.
Private Function ChTxt(ByVal red As Variant, ByVal i As Long) As String
    On Error Resume Next
    If Not IsArray(red) Then Exit Function
    If i > UBound(red) Then Exit Function
    ChTxt = Trim$(CStr(red(i)))
End Function

Private Sub OsveziPalete(ByVal z As Object)
    Dim pal As Collection, sm As Object, i As Long, n As Long, p As Object
    On Error Resume Next
    For i = 0 To ST_PAL_REDOVA - 1
        z.Controls("stPal" & i).caption = ""
    Next i
    modUiKit.BoxShow z, "scrStPal", PitanjeOPaletama()
    modUiKit.BoxState z, "scrStPal", IIf(mNeDiraj, C_FOREST, C_WHITE), _
                      IIf(mNeDiraj, C_CREAM, C_FOREST), mNeDiraj

    If mImpact Is Nothing Then
        z.Controls("stPalZbir").caption = ChrW(8212)
        Exit Sub
    End If
    Set pal = mImpact("palete")
    Set sm = mImpact("summary")
    If pal Is Nothing Then Exit Sub
    If pal.count = 0 Then
        z.Controls("stPalZbir").caption = Poruka("OTKUI_SCRST_BEZ_PALETA")
        Exit Sub
    End If
    ' Zbir je ono sto bi DUPLI/PONISTENJE skinulo - gajbe, neto i ambalaza.
    z.Controls("stPalZbir").caption = _
        CStr(pal.count) & " " & ChrW(183) & " " & CStr(sm("detachGajb")) & " gajb " & _
        ChrW(183) & " " & Format$(CDbl(sm("detachNeto")), "#,##0.##") & " kg " & _
        ChrW(183) & " " & Format$(CDbl(sm("detachAmb")), "#,##0.##") & " amb"

    n = pal.count
    If n > ST_PAL_REDOVA Then n = ST_PAL_REDOVA
    For i = 1 To n
        Set p = pal(i)
        z.Controls("stPal" & (i - 1)).caption = _
            CStr(p("label")) & "  " & CStr(p("used")) & "/" & CStr(p("cap")) & " gajb  " & _
            ChrW(183) & " " & Format$(CDbl(p("thisNeto")), "#,##0.##") & " kg  " & _
            ChrW(183) & " " & Format$(CDbl(p("thisAmb")), "#,##0.##") & " amb" & _
            IIf(CBool(p("preradjena")), "  [PRERADJENA]", "")
    Next i
End Sub

' "Ne diraj palete" ima smisla samo tamo gde palete vise za dokument, i samo
' kad dokument nestaje bez naslednika. Uz ISPRAVKU se palete prevezuju na novi
' dokument, pa se pitanje ne postavlja -- isti uslov koji je do v6-ui-142 stajao
' u IzvrsiMod, samo sto se sada vidi kao prekidac umesto da iskoci kao MsgBox.
Private Function PitanjeOPaletama() As Boolean
    PitanjeOPaletama = (mSelTip = STIP_PRIJEMNICA And Len(mSelBroj) > 0)
End Function

'-------------------------------------------------------- RED ODLUKE
' Koja dugmad stoje u redu odluke, za IZABRAN dokument. Oblik reda:
'   "<mod>|<kljuc natpisa>|<kljuc objasnjenja>|<stil>"
' Prazan povratak = nema izabranog dokumenta, pa nema ni odluke.
'
' Tri slucaja, isti koje je do v6-ui-142 razlikovala IspravkaPreuzela:
'   REVERS         nema nizvodni tok, ali IMA poslovnu razliku storno vs
'                  ispravka (isti fizicki dogadjaj, pogresno unet)
'   framework tip  pun izbor moda, i to SAMO kad StornoTraziIzborModa kaze da
'                  ima o cemu da se odlucuje (prijemnica, palete, faktura,
'                  blokovi); inace je obican storno dovoljan
'   ostali         obican storno
Private Function AkcijeZaTip() As Variant
    Dim dt As String
    On Error GoTo EH
    If Len(mSelBroj) = 0 Then Exit Function

    dt = modStornoDok.TipUFlowDoc(mSelTip)
    If Len(dt) = 0 Then
        AkcijeZaTip = Array("|OTKUI_SCRST_B_STORNO|OTKUI_SCRST_H_STORNO|danger")
        Exit Function
    End If

    If mSelTip = STIP_REVERSI Then
        AkcijeZaTip = Array( _
            "|OTKUI_SCRST_B_STORNO|OTKUI_SCRST_H_STORNO|danger", _
            SV_MODE_ISPRAVKA & "|OTKUI_SCRST_B_REV_ISPR|OTKUI_SCRST_H_REV_ISPR|secondary")
        Exit Function
    End If

    If Not modStornoDok.StornoTraziIzborModa(mSelTip, mSelBroj, mSelOpcija, mSelDocID) Then
        AkcijeZaTip = Array("|OTKUI_SCRST_B_STORNO|OTKUI_SCRST_H_STORNO|danger")
        Exit Function
    End If

    AkcijeZaTip = Array( _
        SV_MODE_ISPRAVKA & "|OTKUI_SCRST_B_ISPRAVKA|OTKUI_SCRST_H_ISPRAVKA|soft", _
        SV_MODE_DUPLI & "|OTKUI_SCRST_B_DUPLI|OTKUI_SCRST_H_DUPLI|secondary", _
        SV_MODE_PONISTENJE & "|OTKUI_SCRST_B_PONISTI|OTKUI_SCRST_H_PONISTI|danger", _
        SV_MODE_RESI_KASNIJE & "|OTKUI_SCRST_B_KASNIJE|OTKUI_SCRST_H_KASNIJE|ghost")
    Exit Function
EH:
    ' FAIL-CLOSED, isto kao StornoTraziIzborModa: neizvesnost vodi ka VISE
    ' pitanja, ne ka manje. Prazan red odluke znaci da se nista ne moze
    ' pokrenuti dok se ne izabere ponovo.
    LogErr "modScrStorno.AkcijeZaTip"
End Function

Private Sub OsveziAkcije(ByVal z As Object)
    Dim akc As Variant, i As Long, p As Variant
    On Error Resume Next
    akc = AkcijeZaTip()
    For i = 0 To 3
        If IsArray(akc) Then
            If i <= UBound(akc) Then
                p = Split(CStr(akc(i)), "|")
                z.Controls("scrStA" & i & "C").caption = Poruka(CStr(p(1)))
                z.Controls("stH" & i).caption = Poruka(CStr(p(2)))
                modUiKit.BoxShow z, "scrStA" & i, True
                z.Controls("stH" & i).Visible = True
                StilDugmeta z, "scrStA" & i, CStr(p(3))
                GoTo NextI
            End If
        End If
        modUiKit.BoxShow z, "scrStA" & i, False
        z.Controls("stH" & i).Visible = False
NextI:
    Next i
End Sub

' Boje po stilu - isti recnik koji BtnV koristi pri gradnji, samo primenjen
' naknadno: ista kontrola nosi razlicit mod zavisno od tipa dokumenta.
Private Sub StilDugmeta(ByVal z As Object, ByVal nm As String, ByVal stil As String)
    On Error Resume Next
    Select Case stil
        Case "danger":    modUiKit.BoxState z, nm, C_RUST, C_WHITE, True
        Case "soft":      modUiKit.BoxState z, nm, C_SOFT_BG, C_GREEN, True
        Case "ghost":     modUiKit.BoxState z, nm, C_WHITE, C_MUTED, False
        Case Else:        modUiKit.BoxState z, nm, C_WHITE, C_FOREST, False
    End Select
End Sub

Private Function FmtDat(ByVal v As Variant) As String
    On Error Resume Next
    FmtDat = CStr(v)
    If IsDate(v) Then FmtDat = Format$(CDate(v), "dd.mm.yyyy.")
End Function

'-------------------------------------------------------------- RADNJE
Public Function Scr_Event(ByVal tag As String, ByVal ev As String) As Boolean
    On Error GoTo EH

    If Left$(tag, 2) = "ls" Then
        If Mid$(tag, 3) = Scr_Lista() Then Exit Function
        mLista = Mid$(tag, 3)
        OcistiIzbor
        Scr_Event = True
        Exit Function
    End If

    If Left$(tag, 4) = "row:" Then
        Scr_Event = IzborReda(CLng(val(Mid$(tag, 5))))
        Exit Function
    End If

    ' Prekidac paleta menja odluku, ne podatke -- mreza se ne sme procitati
    ' ponovo, operater bi izgubio mesto u listi.
    If tag = "scrStPal" Then
        mNeDiraj = Not mNeDiraj
        OsveziZonu
        Exit Function
    End If

    If Left$(tag, 6) = "scrStA" Then
        Scr_Event = PokreniAkciju(CLng(val(Mid$(tag, 7))))
        Exit Function
    End If
    Exit Function
EH:
    LogErr "modScrStorno.Scr_Event"
    modOtkupUI.ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & Err.description, True
End Function

Private Sub OcistiIzbor()
    mSelTip = ""
    mSelBroj = ""
    mSelDocID = ""
    mSelOpcija = ""
    mNeDiraj = False
    Set mImpact = Nothing
    OsveziZonu
End Sub

' Klik na red BIRA dokument i gradi uvid. Nista se ne menja u podacima, pa
' vraca False: mreza ostaje kakva jeste.
'
' U navigacionoj listi klik prebacuje na tipiziranu listu tog dokumenta -- tamo
' postoji kolona identiteta, pa tamo i odluka moze da se donese.
Private Function IzborReda(ByVal red As Long) As Boolean
    Dim tip As String, broj As String
    On Error GoTo EH

    If Scr_Lista() = ST_SVI Then
        tip = TipIzNaziva(Trim$(CStr(modOtkupUI.GridCell(red, 1))))
        If Len(tip) = 0 Then Exit Function
        mLista = tip
        OcistiIzbor
        IzborReda = True
        Exit Function
    End If

    tip = Scr_Lista()
    broj = Trim$(CStr(modOtkupUI.GridCell(red, 1)))
    If Len(broj) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_NEMA_REDA"), True
        Exit Function
    End If

    mSelTip = tip
    mSelBroj = broj
    mSelOpcija = ""
    ' IDENTITET IZABRANOG REDA. Operater je kliknuo konkretan dokument; broj je
    ' ono sto vidi, ali sve nizvodno bira po OVOME. Bez toga bi svaki sloj ispod
    ' ponovo trazio po broju i mogao da nadje tudji dokument istog broja -- a kod
    ' moda RESI KASNIJE se guarded writer uopste ne zove, pa bi se napravio
    ' TRAJAN recovery zapis nad pogresnim dokumentom.
    mSelDocID = IdentIzReda(red, tip)

    ' Dva tipa traze jos jedan podatak koji se iz broja ne vidi.
    If tip = STIP_REVERSI Then mSelOpcija = ReversSmerZaBroj(broj)
    ' Izvod: broj sam nije kljuc (isti broj postoji na vise banaka), a treca
    ' kolona liste JESTE broj racuna - odatle "broj/racun" oblik koji
    ' ResolveIzvodZaStorno razume. Cita se iz reda, ne iz mape: mapa bi kod dva
    ' izvoda istog broja zapamtila samo jedan racun.
    If tip = STIP_IZVOD Then _
        mSelBroj = IzvodKljuc(broj, Trim$(CStr(modOtkupUI.GridCell(red, 3))))

    mNeDiraj = False
    Set mImpact = GradiUvid()
    OsveziZonu
    Exit Function
EH:
    LogErr "modScrStorno.IzborReda"
End Function

' Uvid postoji samo za tipove koje modStornoImpact razume (framework tipovi).
' Za ostale zona pokazuje zaglavlje bez lanca -- i to je tacno, jer lanca nema.
Private Function GradiUvid() As Object
    Dim dt As String
    On Error GoTo EH
    dt = modStornoDok.TipUFlowDoc(mSelTip)
    If Len(dt) = 0 Then Exit Function
    If dt = FLOW_DOC_REVERS Then Exit Function
    Set GradiUvid = modStornoImpact.BuildStornoImpact(dt, mSelBroj, mSelOpcija, mSelDocID)
    Exit Function
EH:
    LogErr "modScrStorno.GradiUvid"
End Function

' Naziv tipa iz navigacione liste -> STIP_* kljuc. Nazivi dolaze iz
' FLOW_DOC_* konstanti, pa se poredi sa njima a ne sa prevodima.
Private Function TipIzNaziva(ByVal naziv As String) As String
    Select Case naziv
        Case FLOW_DOC_OTPREMNICA: TipIzNaziva = STIP_OTPREMNICA
        Case FLOW_DOC_ZBIRNA:     TipIzNaziva = STIP_ZBIRNA
        Case FLOW_DOC_PRIJEMNICA: TipIzNaziva = STIP_PRIJEMNICA
    End Select
End Function

' Indeks nevidljive kolone identiteta: uvek POSLEDNJA koju je GridCols dodao.
' Racuna se iz istog niza koji je mreza dobila, pa ne moze da se razidje sa njim.
Private Function IdentKolona(ByVal tip As String) As Long
    If Len(modScrDokumenti.IdKolonaTipa(tip)) = 0 Then Exit Function
    Dim cols As Variant: cols = modScrDokumenti.GridCols(tip, True)
    If Not IsArray(cols) Then Exit Function
    IdentKolona = UBound(cols) + 1
End Function

' Kanonski identitet KLIKNUTOG reda. Prazno = tip ga nema (revers, izvod) ili
' zatecen zapis bez generacije -- tada nizvodno vazi fail-closed kapija nad
' jednoznacnoscu broja, ista koju koristi i prevezivanje.
Private Function IdentIzReda(ByVal red As Long, ByVal tip As String) As String
    Dim k As Long: k = IdentKolona(tip)
    If k = 0 Then Exit Function
    On Error Resume Next
    IdentIzReda = Trim$(CStr(modOtkupUI.GridCell(red, k)))
End Function

' Koji je od cetiri smera reversa aktivan pod ovim brojem. Cetiri smera dele
' jedan brojevni niz (KIND_REV), pa broj sam ne kaze koji je red u tblAmbalaza;
' pita se istom rutinom kojom legacy proverava postojanje.
Private Function ReversSmerZaBroj(ByVal broj As String) As String
    Dim smerovi As Variant, s As Variant
    On Error Resume Next
    smerovi = Array(DOK_TIP_OM_IZLAZ_KOOP, DOK_TIP_OM_ULAZ_KOOP, _
                    DOK_TIP_OM_ULAZ_FIRMA, DOK_TIP_OM_IZLAZ_FIRMA)
    For Each s In smerovi
        If ActiveAmbalazaDokExists(broj, CStr(s)) Then
            ReversSmerZaBroj = CStr(s)
            Exit Function
        End If
    Next s
End Function

' "broj/racun" - jednoznacan kljuc izvoda. Bez racuna se salje goli broj, pa
' ResolveIzvodZaStorno sam trazi racun (i odbije ako ih ima vise) - isto sto se
' desi kad operater u legacy formi ukuca samo broj.
Private Function IzvodKljuc(ByVal broj As String, ByVal racun As String) As String
    If Len(racun) = 0 Then IzvodKljuc = broj Else IzvodKljuc = broj & "/" & racun
End Function

'------------------------------------------------------------ IZVRSENJE
' Klik na dugme u redu odluke. Prvi mod je prazan string za "obican storno";
' ostali su SV_MODE_*.
Private Function PokreniAkciju(ByVal i As Long) As Boolean
    Dim akc As Variant, mode As String, razlog As String
    On Error GoTo EH
    akc = AkcijeZaTip()
    If Not IsArray(akc) Then Exit Function
    If i > UBound(akc) Then Exit Function
    mode = Split(CStr(akc(i)), "|")(0)

    ' Kapija PRE svega: operater vidi razlog, a ne tih neuspeh posle potvrde.
    razlog = modStornoDok.StornoRazlog(mSelTip, mSelBroj, mSelOpcija, mSelDocID)
    If Len(razlog) > 0 Then
        MsgBox razlog, vbExclamation, APP_NAME
        Exit Function
    End If

    If Len(mode) = 0 Then
        PokreniAkciju = ObicanStorno()
    Else
        PokreniAkciju = StornoPoModu(mode)
    End If
    If PokreniAkciju Then
        Scr_ResetCache
        OcistiIzbor
    End If
    Exit Function
EH:
    LogErr "modScrStorno.PokreniAkciju"
    modOtkupUI.ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & Err.description, True
End Function

' Obican storno: bez framework-a ispravke. Posledice su vec u zoni, pa je
' potvrda kratka -- izvod je izuzetak, jer njegov izbor ishoda za staging redove
' JESTE odluka operatera o PDF-u, a ne pravilo (zato je ovde, a ne u modStornoDok).
Private Function ObicanStorno() As Boolean
    Dim poruka As String, odg As VbMsgBoxResult, opcija As String
    On Error GoTo EH
    opcija = mSelOpcija

    ' Pozivi Poruka() su ovde KVALIFIKOVANI: lokalna "poruka As String" bi ih
    ' inace zaklonila (VBA je case-insensitive), pa bi poziv postao indeksiranje
    ' stringa -- "Expected array". Dvanaest takvih poziva je zivelo neotkriveno
    ' od v6-ui-119 do v6-ui-141, jer VBA proceduru kompajlira tek kad se pozove.
    If mSelTip = STIP_IZVOD Then
        If MsgBox(modStornoDok.StornoPregled(mSelTip, mSelBroj, "") & vbCrLf & vbCrLf & _
                  modPoruke.Poruka("STORNO_ASK_IZVOD"), vbQuestion + vbYesNo, APP_NAME) <> vbYes Then Exit Function
        odg = MsgBox(modPoruke.Poruka("STORNO_ASK_IZVOD_PDF"), vbQuestion + vbYesNoCancel, APP_NAME)
        If odg = vbCancel Then Exit Function
        opcija = IIf(odg = vbYes, IZVOD_STORNO_REMAP, IZVOD_STORNO_REIMPORT)
        If Not modStornoDok.StornoIzvrsi(mSelTip, mSelBroj, opcija, poruka, mSelDocID) Then
            MsgBox poruka, vbExclamation, APP_NAME
            Exit Function
        End If
        ' StornoIzvod_TX vraca izvestaj (koliko redova, koji ishod) - to je vise
        ' nego sto toast moze da pokaze.
        MsgBox poruka, vbInformation, APP_NAME
        ObicanStorno = True
        Exit Function
    End If

    If MsgBox(modPoruke.Poruka("STORNO_ASK") & " " & modStornoDok.TipNaziv(mSelTip, opcija) & _
              " " & mSelBroj & "?", vbQuestion + vbYesNo, APP_NAME) = vbNo Then Exit Function
    If Not modStornoDok.StornoIzvrsi(mSelTip, mSelBroj, opcija, poruka, mSelDocID) Then
        modOtkupUI.ShowToast poruka, True
        Exit Function
    End If
    ' Zbirna moze da vrati UPOZORENJE uz uspeh (aktivna prijemnica ostaje vezana
    ' za storniranu zbirnu) - toast bi ga progutao, pa ide u MsgBox.
    If InStr(1, poruka, modPoruke.Poruka("STORNO_MSG_OK"), vbBinaryCompare) = 1 Then
        modOtkupUI.ShowToast poruka & " " & mSelBroj, False
    Else
        MsgBox poruka, vbExclamation, APP_NAME
    End If
    ObicanStorno = True
    Exit Function
EH:
    LogErr "modScrStorno.ObicanStorno"
End Function

' Izvrsenje izabranog moda.
'
' PONISTENJE se NAMERNO izvrsava u dva poziva: prvi vrati blocked=True i pun
' spisak posledica, i tek posle svesne potvrde ide drugi sa forceConfirm. Spisak
' se tako pravi PRE nego sto se ista promeni.
Private Function StornoPoModu(ByVal mode As String) As Boolean
    Dim res As Object, neDiraj As Boolean
    On Error GoTo EH
    ' Prekidac vazi samo tamo gde je i ponudjen; uz ISPRAVKU se palete prevezuju
    ' na novi dokument, pa se ne postuje ni ako je ostao ukljucen.
    neDiraj = (mNeDiraj And mSelTip = STIP_PRIJEMNICA And _
               (mode = SV_MODE_DUPLI Or mode = SV_MODE_PONISTENJE))

    Set res = modStornoDok.StornoIzvrsiMod(mSelTip, mSelBroj, mSelOpcija, mode, _
                                           False, neDiraj, mSelDocID)
    If res Is Nothing Then
        modOtkupUI.ShowToast Poruka("STORNO_ERR_NEPOZNAT_TIP") & " " & mSelTip, True
        Exit Function
    End If
    If CBool(res("blocked")) Then
        If MsgBox(CStr(res("message")) & vbCrLf & vbCrLf & Poruka("STORNO_ASK_PONISTI"), _
                  vbExclamation + vbYesNo, APP_NAME) <> vbYes Then
            modOtkupUI.ShowToast Poruka("STORNO_MSG_ODUSTANAK"), False
            Exit Function
        End If
        Set res = modStornoDok.StornoIzvrsiMod(mSelTip, mSelBroj, mSelOpcija, mode, _
                                               True, neDiraj, mSelDocID)
        If res Is Nothing Then Exit Function
    End If

    If CBool(res("needsForm")) Then
        ' ISPRAVKA: stari dokument je stigao do storna, novi se tek unosi.
        OtvoriIspravku mSelTip, mSelBroj, CStr(res("correctionID"))
        MsgBox CStr(res("message")) & vbCrLf & vbCrLf & _
               Poruka("STORNO_MSG_ISPRAVKA_DALJE"), vbInformation, APP_NAME
        StornoPoModu = True
        Exit Function
    End If

    If CBool(res("success")) Then
        MsgBox CStr(res("message")), vbInformation, APP_NAME
        ' Blokovi se storniraju POSLE dokumenta: kapija gleda da li je roditeljska
        ' otpremnica jos aktivna, a to zavisi od toga sta je upravo odradjeno.
        StornirajBlokoveAko mode, CStr(res("correctionID"))
        StornoPoModu = True
        Exit Function
    End If

    MsgBox CStr(res("message")), vbExclamation, APP_NAME
    Exit Function
EH:
    LogErr "modScrStorno.StornoPoModu"
    modOtkupUI.ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & Err.description, True
End Function

' Dodatni storno otkupnih blokova koji vise o dokumentu.
'
' Dozvoljen je SAMO uz DUPLI/PONISTENJE - roba tada stvarno nestaje. Uz ISPRAVKU
' blokovi ostaju: storno pa ponovni unos dokumenta je celina, a blokovi se
' prevezuju na novi. Uz RESI KASNIJE se ne dira nista.
'
' RAZLIKA U ODNOSU NA LEGACY, i to je jedino sto od panela nije preneto: legacy
' nudi multiselect - operater cekira KOJE blokove stornira, a ostali se
' oslobadjaju. Ovde je sve-ili-nista, ali se pre pitanja ispise pun spisak
' (broj, klasa, kilogrami, kooperant), pa operater vidi tacno nad cim odlucuje.
' Delimican izbor ostaje na ekranu Oporavak, gde izgubljeni blokovi imaju svoju
' listu i radnju po redu. Kolona sa checkbox-om je Korak 2.
'
' Kapija BlockStornoDriftReason ostaje ista (ADR-0001): blok vezan za ZIVU
' otpremnicu se ne stornira, jer bi je ostavio precenjenu.
Private Sub StornirajBlokoveAko(ByVal mode As String, ByVal correctionID As String)
    Dim dt As String, blokovi As Collection, ids As Collection
    Dim spisak As String, i As Long, red As Variant, n As Long, razlog As String
    On Error GoTo EH
    If mode <> SV_MODE_DUPLI And mode <> SV_MODE_PONISTENJE Then Exit Sub
    dt = modStornoDok.TipUFlowDoc(mSelTip)
    If Len(dt) = 0 Then Exit Sub

    ' Identitet izabranog reda ide i ovde: spisak koji operater potvrdjuje
    ' zavrsava u StornoSelectedBlocks_TX, dakle u MUTACIJI.
    Set blokovi = GetStornoBlockRows(dt, mSelBroj, mSelOpcija, mSelDocID)
    If blokovi Is Nothing Then Exit Sub
    If blokovi.count = 0 Then Exit Sub

    For i = 1 To blokovi.count
        red = blokovi(i)
        spisak = spisak & vbCrLf & "  " & CStr(red(1)) & "  " & ChrW(183) & " Kl." & _
                 CStr(red(3)) & "  " & ChrW(183) & " " & CStr(red(2)) & " kg  " & _
                 ChrW(183) & " " & CStr(red(4))
    Next i

    If MsgBox(Poruka("STORNO_ASK_BLOKOVI_1") & " " & blokovi.count & _
              Poruka("STORNO_ASK_BLOKOVI_2") & spisak & vbCrLf & vbCrLf & _
              Poruka("STORNO_ASK_BLOKOVI_3"), _
              vbExclamation + vbYesNo + vbDefaultButton2, APP_NAME) <> vbYes Then Exit Sub

    Set ids = New Collection
    For i = 1 To blokovi.count
        red = blokovi(i)
        ids.Add CStr(red(0))
    Next i

    razlog = BlockStornoDriftReason(dt, mode, ids)
    If Len(razlog) > 0 Then
        MsgBox razlog, vbExclamation, APP_NAME
        Exit Sub
    End If

    n = StornoSelectedBlocks_TX(ids)
    If n > 0 Then
        modOtkupUI.ShowToast Poruka("STORNO_MSG_BLOKOVI_OK") & " " & n, False
        Exit Sub
    End If

    ' Storno dokumenta je vec komitovan; ako blok-storno padne, posao se ne gubi
    ' u poruci koja prodje nego ostaje vidljiv kao MANUAL zadatak.
    MsgBox Poruka("STORNO_ERR_BLOKOVI"), vbExclamation, APP_NAME
    If Len(correctionID) > 0 Then
        modStornoContext.MarkCorrectionManual correctionID, _
            "Storniraj otkupne blokove rucno.", _
            "Posle " & mode & " nad " & dt & " " & mSelBroj & " storno blokova nije uspeo."
    End If
    Exit Sub
EH:
    LogErr "modScrStorno.StornirajBlokoveAko"
End Sub

'--------------------------------------------------- PREDAJA UNOSNOM EKRANU
' Otvori unos zamenskog dokumenta: predji na EKRAN dokumenata, pa u njegov
' REZIM, pa prepisi polja iz stornirane.
'
' REDOSLED JE OBAVEZAN, i to su tri koraka a ne dva otkad je storno svoj ekran:
'   1. IdiNaEkran -- bez toga prefill upisuje u zForm koja je sakrivena, pa
'      operater posle poruke "unesi novi dokument" gleda u listu storna;
'   2. IdiNaRezim -- SelectMode CISTI formu, pa bi prefill pre njega bio obrisan;
'   3. ApplyPrefill.
'
' Prefill polazi od PK-a stornirane (OldDocID iz correction context-a), ne od
' broja: broj nije globalno jedinstven.
Private Sub OtvoriIspravku(ByVal tip As String, ByVal broj As String, _
                           ByVal correctionID As String)
    Dim spec As String
    On Error Resume Next
    modOtkupUI.IdiNaEkran "DOKUMENTI"
    modOtkupUI.IdiNaRezim RezimZaTip(tip)
    spec = modStornoDok.PrefillIzStorniranog(tip, broj, modStornoDok.StorniraniDocID(correctionID))
    If Len(spec) > 0 Then modOtkupUI.ApplyPrefill spec
End Sub

' Rezim u kome se unosi zamenski dokument. Revers ide u F7; smer operater bira
' sam, jer prefill reversa ne postoji ni u legacy.
Private Function RezimZaTip(ByVal tip As String) As String
    Select Case tip
        Case STIP_OTKUP:      RezimZaTip = "F1"
        Case STIP_OTPREMNICA: RezimZaTip = "F2"
        Case STIP_ZBIRNA:     RezimZaTip = "F3"
        Case STIP_PRIJEMNICA: RezimZaTip = "F4"
        Case STIP_REVERSI:    RezimZaTip = "F7"
        Case Else:            RezimZaTip = modOtkupUI.ActiveMode
    End Select
End Function

'--------------------------------------------------------------- KES
Public Sub Scr_ResetCache()
    Set mImpact = Nothing
    modUiData.ResetCache
End Sub
