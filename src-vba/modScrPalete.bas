Attribute VB_Name = "modScrPalete"
'=====================================================================
' modScrPalete - ekran "Palete".
'
' PRVI ekran koji ljuska crta ISKLJUCIVO kroz ugovor iz modUiScreens.
' Ljuska ga ne poznaje po imenu: dobija ga preko Application.Run, da
' klijent kome ovaj modul nedostaje i dalje radi (zamka #19).
'
' Od S4a ekran VISE NE CRTA SVOJU LISTU. Prva verzija je imala sopstvenu
' tabelu od stotinak linija - bez strana, sortiranja, pretrage i izbora
' reda, pa je od 188 paleta pokazivala 18. Druga lista u istoj aplikaciji
' je uz to bila tacno ono dupliranje koje CLAUDE.md zabranjuje.
'
' Sada ekran daje samo PODATKE (Scr_Rows), a mrezu crta ljuska - istu
' onu koju koristi i ekran dokumenata, sa svime sto ona vec ume.
'
' U svojoj zoni ostaju samo tri brojke (ukupno / otvorene / gajbica).
'
' Podaci se citaju iz tblPaleta preko postojeceg modDataAccess
' (GetTableData / GetColumnIndex). Nista se ne upisuje.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const SCRPAL_BUILD As String = "v6-ui-80"

' Visina zone = visina KPI trake na ekranu dokumenata. Zona ugovornog ekrana
' stoji na istom mestu i iste je visine, pa naslov ispod nje pada u isti red
' na oba ekrana.
Private Const PAL_ZONA_H As Single = KPI_H

'--------------------------------------------------------- UGOVOR EKRANA
Public Function Scr_Meta() As String
    Scr_Meta = "kljuc=PALETE|naslov=OTKUI_NAV_PALETE|sub=OTKUI_SCRPAL_SUB" & _
               "|lista=OTKUI_SCRPAL_LISTA|oblik=lista|upis=ne"
End Function

Public Sub Scr_Build(ByVal z As Object)
    Dim i As Long
    ' tri brojke uz desnu ivicu - isti materijal kao KPI traka, samo uze
    For i = 0 To 2
        modUiKit.NewLbl z, "palKL" & i, "", 0, 14, 120, 12, TS_MICRO, True, C_MUTED, -1
        modUiKit.NewLbl z, "palKV" & i, ChrW(8212), 0, 28, 120, 20, TS_KPI, True, _
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

' Ekran jos nema svoje radnje - klik nigde nista ne menja.
Public Function Scr_Event(ByVal tag As String, ByVal ev As String) As Boolean
    Scr_Event = False
End Function

'--------------------------------------------------------------- REDOVI
' Kolone deljene mreze. Isti oblik kao kod ekrana dokumenata:
'   KLJUC_KATALOGA | izvor (ovde prazno - punimo rucno) | vrsta | sirina | prio
' Vrsta odredjuje tipografiju celije u mrezi (num/kg/txt/date/pill).
Private Function PalGridCols() As Variant
    PalGridCols = Array( _
        "OTKUI_HDP_BROJ||txt|70|1", _
        "OTKUI_HDP_GODINA||txt|56|3", _
        "OTKUI_HD_DATUM||date|62|1", _
        "OTKUI_HD_VRSTA||txt|90|2", _
        "OTKUI_HD_SORTA||txt|110|2", _
        "OTKUI_HD_KLASA||txt|46|3", _
        "OTKUI_HDP_GAJBICA||num|64|1", _
        "OTKUI_HDP_NETO||kg|76|1", _
        "OTKUI_HDP_BRUTO||kg|76|2", _
        "OTKUI_HD_STATUS||txt|84|1")
End Function

' Ugovor: Array(kolone, redovi, n, zbirKg, zbirVal).
' Redovi idu 1-bazirano po oba indeksa - tako ih mreza ocekuje.
Public Function Scr_Rows(ByVal filter As String, ByVal q As String) As Variant
    Dim src As Variant, outA() As Variant, r As Long, n As Long, uk As Long
    Dim iBroj As Long, iGod As Long, iDat As Long, iVrsta As Long, iSorta As Long
    Dim iKlasa As Long, iGajb As Long, iNeto As Long, iBruto As Long, iStat As Long
    Dim otvorene As Long, gajbi As Double, sumNeto As Double, hay As String
    Dim st As String, keep As Boolean
    On Error GoTo EH

    src = GetTableData(TBL_PALETA)
    If Not IsArray(src) Then Exit Function

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
    ReDim outA(1 To uk + 1, 1 To 10)
    For r = 1 To uk
        st = PalS(src, r, iStat)
        ' zbirovi u zaglavlju idu preko SVIH paleta, ne preko filtrirane liste
        gajbi = gajbi + PalD(src, r, iGajb)
        If UCase$(st) <> "ZATVORENA" Then otvorene = otvorene + 1

        hay = PalS(src, r, iBroj) & "|" & PalS(src, r, iVrsta) & "|" & _
              PalS(src, r, iSorta) & "|" & st
        keep = True
        If Len(q) > 0 Then keep = (InStr(1, hay, q, vbTextCompare) > 0)
        If Not keep Then GoTo Sledeci

        n = n + 1
        outA(n, 1) = PalS(src, r, iBroj)
        outA(n, 2) = PalS(src, r, iGod)
        outA(n, 3) = PalDatum(src, r, iDat)
        outA(n, 4) = PalS(src, r, iVrsta)
        outA(n, 5) = PalS(src, r, iSorta)
        outA(n, 6) = PalS(src, r, iKlasa)
        outA(n, 7) = Format$(PalD(src, r, iGajb), "#,##0")
        outA(n, 8) = Format$(PalD(src, r, iNeto), "#,##0.00")
        outA(n, 9) = Format$(PalD(src, r, iBruto), "#,##0.00")
        outA(n, 10) = st
        sumNeto = sumNeto + PalD(src, r, iNeto)
Sledeci:
    Next r

    RefreshBrojke uk, otvorene, gajbi
    Scr_Rows = Array(PalGridCols(), outA, n, sumNeto, 0#)
    Exit Function
EH:
    Debug.Print "modScrPalete.Scr_Rows: " & Err.Number & " " & Err.description
End Function

' Tri brojke u zoni ekrana. Zona se gradi jednom, pa se pune tek kad se cita
' tabela - a to je u Scr_Rows.
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
End Sub

'-------------------------------------------------------------- CELIJE
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
