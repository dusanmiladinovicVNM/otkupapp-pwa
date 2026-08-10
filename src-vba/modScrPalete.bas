Attribute VB_Name = "modScrPalete"
'=====================================================================
' modScrPalete - ekran "Palete" (faza S3b).
'
' PRVI ekran koji ljuska crta ISKLJUCIVO kroz ugovor iz modUiScreens.
' Ljuska ga ne poznaje po imenu: dobija ga preko Application.Run, da
' klijent kome ovaj modul nedostaje i dalje radi (zamka #19).
'
' Ekran dobije praznu zonu i u njoj radi sta hoce. Ovde: zaglavlje sa
' naslovom i tri brojke, pa lista poslednjih paleta. Crtanje ide kroz
' modUiKit, boje kroz modOtkupUI - nijedna nova boja i nijedna nova
' fabrika kontrola.
'
' Podaci se citaju iz tblPaleta preko postojeceg modDataAccess
' (GetTableData / GetColumnIndex). Nista se ne upisuje.
'
' STO OVDE NAMERNO NIJE: deljena mreza. Ona danas zivi u ljusci i puni
' se iz FillGrid, koje je vezano za ekran dokumenata. Dok se to ne
' razdvoji, ovaj ekran crta svoju, jednostavniju listu. Kad mreza dobije
' vlasnika, ovaj ekran je zamenjuje sa Scr_Grid i gubi pola koda.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const SCRPAL_BUILD As String = "v6-ui-74"

Private Const PAL_ROWS As Long = 18          ' koliko redova liste se pravi
Private Const PAL_ROW_H As Single = 21
Private Const PAL_HEAD_H As Single = 21
Private Const PAL_TOP As Single = 86         ' ispod naslova i brojki
' Razmak izmedju kolona. Bez njega se desno poravnat broj lepi uz levo
' poravnat tekst susedne kolone: zaglavlje je pisalo "BRUTO KGSTATUS", a
' red "50,47Otvorena". Kolona dobija svoju sirinu MINUS ovaj razmak.
Private Const PAL_GUT As Single = 10

' Koliko redova je trenutno napunjeno. Sirina zone se zna tek u Scr_Layout,
' pa se i podaci ucitavaju tamo - u Scr_Build zona jos ima pocetnih 400pt i
' stalo bi 11 redova umesto 18.
Private mLoaded As Long

' Kolone liste: naslov | sirina | desno poravnanje
Private Function PalCols() As Variant
    PalCols = Array("BROJ|70|0", "GODINA|56|0", "DATUM|62|0", "VRSTA|90|0", _
                    "SORTA|110|0", "KLASA|48|0", "GAJBICA|64|1", "NETO KG|76|1", _
                    "BRUTO KG|76|1", "STATUS|84|0")
End Function

'--------------------------------------------------------- UGOVOR EKRANA
Public Function Scr_Meta() As String
    Scr_Meta = "kljuc=PALETE|naslov=OTKUI_NAV_PALETE|oblik=lista|upis=ne"
End Function

Public Sub Scr_Build(ByVal z As Object)
    Dim i As Long, c As Long, cols As Variant, p() As String

    modUiKit.NewLbl z, "palTitle", Poruka("OTKUI_NAV_PALETE"), PAD, 22, 300, 26, _
                    TS_H1, True, C_FOREST, -1
    modUiKit.NewLbl z, "palSub", Poruka("OTKUI_SCRPAL_SUB"), PAD, 50, 460, 14, _
                    TS_META, False, C_MUTED, -1

    ' tri brojke desno od naslova - isti materijal kao KPI trake, samo uze
    For i = 0 To 2
        modUiKit.NewLbl z, "palKL" & i, "", 0, 22, 120, 12, TS_MICRO, True, C_MUTED, -1
        modUiKit.NewLbl z, "palKV" & i, ChrW(8212), 0, 36, 120, 20, TS_KPI, True, C_FOREST, _
                        -1, fmTextAlignLeft, F_NUM
    Next i

    ' zaglavlje liste
    modUiKit.NewLbl z, "palHead", "", 0, PAL_TOP, 100, PAL_HEAD_H, 8, False, 0, C_HEAD_BG
    cols = PalCols()
    For c = 0 To UBound(cols)
        p = Split(CStr(cols(c)), "|")
        modUiKit.NewLbl z, "palH" & c, p(0), 0, PAL_TOP + 6, 60, 12, TS_MICRO, True, _
                        C_MUTED, -1, IIf(p(2) = "1", fmTextAlignRight, fmTextAlignLeft)
    Next c
    modUiKit.NewLbl z, "palHLn", "", 0, PAL_TOP + PAL_HEAD_H, 100, 1, 8, False, 0, C_BORDER

    ' redovi - prave se jednom, pune se pri svakom osvezavanju
    For i = 0 To PAL_ROWS - 1
        modUiKit.NewLbl z, "palRB" & i, "", 0, 0, 100, PAL_ROW_H, 8, False, 0, C_WHITE
        For c = 0 To UBound(cols)
            p = Split(CStr(cols(c)), "|")
            modUiKit.NewLbl z, "palC" & i & "_" & c, "", 0, 0, 60, 12, TS_CELL, False, _
                            C_FOREST, -1, IIf(p(2) = "1", fmTextAlignRight, fmTextAlignLeft)
        Next c
        modUiKit.NewLbl z, "palRL" & i, "", 0, 0, 100, 1, 8, False, 0, C_ROW_LN
    Next i

    modUiKit.NewLbl z, "palFoot", "", PAD, 0, 400, 14, TS_META, False, C_MUTED, -1
    mLoaded = -1                     ' puni se u Scr_Layout, kad se zna visina
End Sub

Public Function Scr_Layout(ByVal z As Object, ByVal w As Single, ByVal h As Single) As Single
    Dim i As Long, c As Long, cols As Variant, p() As String
    Dim X As Single, Y As Single, uk As Single, n As Long
    On Error Resume Next

    ' brojke se lepe uz desnu ivicu, istim redom kao u KPI traci
    For i = 0 To 2
        z.Controls("palKL" & i).Left = w - PAD - (3 - i) * 130
        z.Controls("palKV" & i).Left = w - PAD - (3 - i) * 130
    Next i

    z.Controls("palHead").Left = PAD
    z.Controls("palHead").width = w - 2 * PAD
    z.Controls("palHLn").Left = PAD
    z.Controls("palHLn").width = w - 2 * PAD

    cols = PalCols()
    For c = 0 To UBound(cols)
        uk = uk + CSng(Val(Split(CStr(cols(c)), "|")(1)))
    Next c

    n = VidljivihRedova(h)
    Y = PAL_TOP + PAL_HEAD_H + 1
    For i = 0 To PAL_ROWS - 1
        If i < n Then
            z.Controls("palRB" & i).Left = PAD
            z.Controls("palRB" & i).width = w - 2 * PAD
            z.Controls("palRB" & i).top = Y
            X = PAD + 6
            For c = 0 To UBound(cols)
                p = Split(CStr(cols(c)), "|")
                With z.Controls("palC" & i & "_" & c)
                    .Left = X
                    .width = CSng(Val(p(1))) * ((w - 2 * PAD - 12) / uk) - PAL_GUT
                    .top = Y + (PAL_ROW_H - modUiKit.TxtH(TS_CELL)) / 2
                End With
                z.Controls("palC" & i & "_" & c).ZOrder 0
                X = X + CSng(Val(p(1))) * ((w - 2 * PAD - 12) / uk)
            Next c
            z.Controls("palRL" & i).Left = PAD
            z.Controls("palRL" & i).width = w - 2 * PAD
            z.Controls("palRL" & i).top = Y + PAL_ROW_H
            Y = Y + PAL_ROW_H + 1
        End If
        z.Controls("palRB" & i).Visible = (i < n)
        z.Controls("palRL" & i).Visible = (i < n)
        For c = 0 To UBound(cols)
            z.Controls("palC" & i & "_" & c).Visible = (i < n)
        Next c
    Next i

    ' naslovi kolona prate iste sirine
    X = PAD + 6
    For c = 0 To UBound(cols)
        p = Split(CStr(cols(c)), "|")
        z.Controls("palH" & c).Left = X
        z.Controls("palH" & c).width = CSng(Val(p(1))) * ((w - 2 * PAD - 12) / uk) - PAL_GUT
        z.Controls("palH" & c).ZOrder 0
        X = X + CSng(Val(p(1))) * ((w - 2 * PAD - 12) / uk)
    Next c

    z.Controls("palFoot").top = Y + 8
    z.Controls("palFoot").width = w - 2 * PAD

    ' Podaci se citaju tek kad se zna koliko redova staje. Ponovo se citaju
    ' samo ako se taj broj promenio - promena sirine prozora ne dira tabelu.
    If n <> mLoaded Then LoadPalete z, n
    Scr_Layout = Y + 24
End Function

' Ekran jos nema svoje radnje - klik nigde nista ne menja.
Public Function Scr_Event(ByVal tag As String, ByVal ev As String) As Boolean
    Scr_Event = False
End Function

'-------------------------------------------------------------- PODACI
Private Function VidljivihRedova(ByVal h As Single) As Long
    Dim n As Long
    n = Int((h - PAL_TOP - PAL_HEAD_H - 40) / (PAL_ROW_H + 1))
    If n < 1 Then n = 1
    If n > PAL_ROWS Then n = PAL_ROWS
    VidljivihRedova = n
End Function

' Poslednje palete, najnovija prva. Cita se postojecim modDataAccess-om;
' nista se ne upisuje i nijedna paleta se ne zatvara odavde.
Private Sub LoadPalete(ByVal z As Object, ByVal n As Long)
    Dim src As Variant, r As Long, i As Long, cc As Long
    Dim iBroj As Long, iGod As Long, iDat As Long, iVrsta As Long, iSorta As Long
    Dim iKlasa As Long, iGajb As Long, iNeto As Long, iBruto As Long, iStat As Long
    Dim otvorene As Long, gajbi As Double, neto As Double, uk As Long
    On Error GoTo EH

    src = GetTableData(TBL_PALETA)
    If Not IsArray(src) Then Exit Sub

    iBroj = GetColumnIndex(TBL_PALETA, COL_PAL_BROJ)
    iGod = GetColumnIndex(TBL_PALETA, COL_PAL_GODINA)
    iDat = GetColumnIndex(TBL_PALETA, COL_PAL_DATUM)
    iVrsta = GetColumnIndex(TBL_PALETA, COL_PAL_VRSTA)
    iSorta = GetColumnIndex(TBL_PALETA, COL_PAL_SORTA)
    iKlasa = GetColumnIndex(TBL_PALETA, COL_PAL_KLASA)
    iGajb = GetColumnIndex(TBL_PALETA, COL_PAL_BR_GAJBICA)
    iNeto = GetColumnIndex(TBL_PALETA, COL_PAL_NETO)
    iBruto = GetColumnIndex(TBL_PALETA, COL_PAL_BRUTO)
    iStat = GetColumnIndex(TBL_PALETA, COL_PAL_STATUS)

    uk = UBound(src, 1)
    ' zbirovi idu preko SVIH paleta, ne samo preko prikazanih
    For r = 1 To uk
        gajbi = gajbi + PalD(src, r, iGajb)
        neto = neto + PalD(src, r, iNeto)
        If UCase$(PalS(src, r, iStat)) <> "ZATVORENA" Then otvorene = otvorene + 1
    Next r

    i = 0
    For r = uk To 1 Step -1                  ' najnovija prva
        If i >= n Then Exit For
        z.Controls("palC" & i & "_0").caption = PalS(src, r, iBroj)
        z.Controls("palC" & i & "_1").caption = PalS(src, r, iGod)
        z.Controls("palC" & i & "_2").caption = PalDatum(src, r, iDat)
        z.Controls("palC" & i & "_3").caption = PalS(src, r, iVrsta)
        z.Controls("palC" & i & "_4").caption = PalS(src, r, iSorta)
        z.Controls("palC" & i & "_5").caption = PalS(src, r, iKlasa)
        z.Controls("palC" & i & "_6").caption = Format$(PalD(src, r, iGajb), "#,##0")
        z.Controls("palC" & i & "_7").caption = Format$(PalD(src, r, iNeto), "#,##0.00")
        z.Controls("palC" & i & "_8").caption = Format$(PalD(src, r, iBruto), "#,##0.00")
        z.Controls("palC" & i & "_9").caption = PalS(src, r, iStat)
        z.Controls("palRB" & i).BackColor = IIf(i Mod 2 = 1, C_ROW_ALT, C_WHITE)
        i = i + 1
    Next r

    z.Controls("palKL0").caption = UCase$(Poruka("OTKUI_SCRPAL_UKUPNO"))
    z.Controls("palKV0").caption = Format$(uk, "#,##0")
    z.Controls("palKL1").caption = UCase$(Poruka("OTKUI_SCRPAL_OTVORENE"))
    z.Controls("palKV1").caption = Format$(otvorene, "#,##0")
    z.Controls("palKL2").caption = UCase$(Poruka("OTKUI_SCRPAL_GAJBICA"))
    z.Controls("palKV2").caption = Format$(gajbi, "#,##0")
    z.Controls("palFoot").caption = Poruka("OTKUI_SCRPAL_FOOT") & " " & _
                                   Format$(i, "#,##0") & " / " & Format$(uk, "#,##0") & _
                                   "  " & ChrW(183) & "  " & Format$(neto, "#,##0.00") & " kg"
    ' redovi bez podatka ne smeju da ostanu vidljivi sa starim tekstom
    For r = i To PAL_ROWS - 1
        z.Controls("palRB" & r).Visible = False
        z.Controls("palRL" & r).Visible = False
        For cc = 0 To UBound(PalCols())
            z.Controls("palC" & r & "_" & cc).Visible = False
        Next cc
    Next r
    mLoaded = i
    Exit Sub
EH:
    Debug.Print "modScrPalete.LoadPalete: " & Err.Number & " " & Err.description
End Sub

Private Function PalS(ByRef src As Variant, ByVal r As Long, ByVal c As Long) As String
    If c < 1 Then Exit Function
    Dim v As Variant: v = src(r, c)
    If IsEmpty(v) Then Exit Function
    PalS = Trim$(CStr(v))
End Function

Private Function PalD(ByRef src As Variant, ByVal r As Long, ByVal c As Long) As Double
    If c < 1 Then Exit Function
    Dim v As Variant: v = src(r, c)
    If IsNumeric(v) Then PalD = CDbl(v)
End Function

Private Function PalDatum(ByRef src As Variant, ByVal r As Long, ByVal c As Long) As String
    If c < 1 Then Exit Function
    Dim v As Variant: v = src(r, c)
    If IsDate(v) Then
        PalDatum = Format$(CDate(v), "dd.mm.")
    ElseIf IsNumeric(v) Then
        If CDbl(v) > 0 Then PalDatum = Format$(CDate(CDbl(v)), "dd.mm.")
    End If
End Function
