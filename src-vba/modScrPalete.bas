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

Public Const SCRPAL_BUILD As String = "v6-ui-106"

' Visina zone = visina KPI trake na ekranu dokumenata. Zona ugovornog ekrana
' stoji na istom mestu i iste je visine, pa naslov ispod nje pada u isti red
' na oba ekrana.
Private Const PAL_ZONA_H As Single = KPI_H

' Stanje ekrana. Aktivna paleta je ona na koju je poslednji put kliknuto u
' listi paleta; nju gledaju i stavke i radnje nad jednom paletom.
Private mLista As String          ' "PALETE" | "STAVKE" | "PRERADE"
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
        "PRERADE|OTKUI_SEG_PAL_PRERADE|OTKUI_GRID_TITLE_PRERADE|96")
End Function

Public Function Scr_Lista() As String
    If Len(mLista) = 0 Then mLista = "PALETE"
    Scr_Lista = mLista
End Function

' U listi stavki naslov nosi broj palete cije su to stavke.
Public Function Scr_NaslovDopuna() As String
    If Scr_Lista() = "STAVKE" Then Scr_NaslovDopuna = mPalBroj
End Function

' Radnje nad redom za aktivnu listu: kljuc:natpis:sirina:stil:trebaRed
Public Function Scr_Radnje() As String
    Select Case Scr_Lista()
        Case "PALETE"
            Scr_Radnje = "palprint:OTKUI_BTN_PAL_PRINT:112:ghost:1|" & _
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
    modUiKit.NewLbl z, "palLnB", "", 0, PAL_ZONA_H - 1, 100, 1, 8, False, 0, C_BORDER
End Sub

Public Function Scr_Layout(ByVal z As Object, ByVal w As Single, ByVal h As Single) As Single
    Dim i As Long
    On Error Resume Next
    For i = 0 To 2
        z.Controls("palKL" & i).Left = w - PAD - (3 - i) * 150
        z.Controls("palKV" & i).Left = w - PAD - (3 - i) * 150
    Next i
    z.Controls("palLnB").width = w
    Scr_Layout = PAL_ZONA_H
End Function

'-------------------------------------------------------------- RADNJE
Public Function Scr_Event(ByVal tag As String, ByVal ev As String) As Boolean
    Dim broj As String
    On Error Resume Next

    If Left$(tag, 2) = "ls" Then
        If Mid$(tag, 3) = Scr_Lista() Then Exit Function
        mLista = Mid$(tag, 3)
        Scr_Event = True
        Exit Function
    End If

    ' Izbor reda u listi paleta postavlja AKTIVNU paletu. Ekran vraca False:
    ' izbor ne menja podatke, pa mreza ne sme da se prazni - radnje nad redom
    ' rade bas nad tim izabranim redom.
    If Left$(tag, 4) = "row:" And Scr_Lista() = "PALETE" Then
        broj = Trim$(CStr(modOtkupUI.GridCell(CLng(Mid$(tag, 5)), 1)))
        If Len(broj) = 0 Then Exit Function
        If mPalIds Is Nothing Then Exit Function
        If Not mPalIds.Exists(broj) Then Exit Function
        mPalID = CStr(mPalIds(broj))
        mPalBroj = broj
        ' vrsta, sorta, klasa i status stoje u redu koji je upravo izabran
        mPalOpis = Trim$(CStr(modOtkupUI.GridCell(CLng(Mid$(tag, 5)), 3))) & "  " & _
                   ChrW(183) & "  " & Trim$(CStr(modOtkupUI.GridCell(CLng(Mid$(tag, 5)), 4))) & _
                   "  " & ChrW(183) & "  " & _
                   Trim$(CStr(modOtkupUI.GridCell(CLng(Mid$(tag, 5)), 11)))
        RefreshAktivna
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
    Scr_Rows = RowsPalete(q)
End Function

' Kolone deljene mreze: KLJUC_KATALOGA | izvor | vrsta | sirina | prio
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
Private Function RowsPalete(ByVal q As String) As Variant
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
