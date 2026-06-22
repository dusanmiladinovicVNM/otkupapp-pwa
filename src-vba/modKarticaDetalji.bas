Attribute VB_Name = "modKarticaDetalji"
'Attribute VB_Name = "modKarticaDetalji"
Option Explicit

' ============================================================
' modKarticaDetalji — read-only panel "Detalji otkupa" uz karticu
' kooperanta (frmIzvestaj, Page 8 / lstKartica).
'
' Klik na red kartice prikazuje SVE bitne stavke izabranog otkupnog
' lista (polja koja se unose u frmOtkup) u listi DESNO — samo pregled
' ("ne za izmenu, samo za pregled"); ListBox po prirodi nije editabilan.
'
' Kontrole su dinamicke (Controls.Add u ISTI container kao lstKartica —
' lstKartica.Parent.Controls, kao frmDokumenta), frmIzvestaj.frx se NE
' dira (CLAUDE.md). Ref-kljuc reda se nosi u skrivenoj koloni lstKartica
' (idx 7, vidi frmIzvestaj.SetupListBoxes): "OTK|<OtkupID>" / "NOV" / "MAG".
' ============================================================

' Indeks skrivene kolone sa ref-kljucem u lstKartica.
' (kol. 6 = Saldo ambalaze, kol. 7 = skriveni ref-kljuc)
Private Const KART_REFKEY_COL As Long = 7

Private mLst As MSForms.ListBox        ' detail lista (2 kolone: Stavka | Vrednost)
Private mLblTitle As MSForms.label
Private mBuilt As Boolean
Private mFormLevel As Boolean          ' True ako su kontrole na formi (fallback), ne na Page-u
Private mCurOtkupID As String          ' OtkupID trenutno prikazan u panelu (za "Stampaj otkupni list")

' Dijagnostika (samo kad build ne uspe): jednom po sesiji prikazi razlog.
Private mDiag As String
Private mDiagShown As Boolean

' Reset modul-stanja (zovi iz frmIzvestaj setup-a): posle Unload-a forme
' dinamicke kontrole vise ne postoje, pa se mBuilt/mLst moraju resetovati.
Public Sub KarticaDetalji_Reset()
    Set mLst = Nothing
    Set mLblTitle = Nothing
    mBuilt = False
    mCurOtkupID = ""
End Sub

' Lazy-build panel desno od lstKartica (idempotentno).
Public Sub KarticaDetalji_Ensure(ByVal frm As Object, ByVal lstKartica As MSForms.ListBox)
    On Error GoTo EH
    If mBuilt Then Exit Sub
    If lstKartica Is Nothing Then Exit Sub

    ' Kontrole se prave NA FORMI (frm.Controls.Add — proveren obrazac, modOtkupBlok)
    ' i pozicioniraju forma-relativnim koordinatama (frm.InsideWidth/Height) tako da
    ' NE mogu da budu van ekrana. ZOrder ih dize u prvi plan (preko MultiPage-a).
    ' SetVisible ih krije van kartice (mFormLevel=True).
    Dim cont As Object
    Set cont = frm
    mFormLevel = True

    Dim leftX As Double, topY As Double, panelH As Double, panelW As Double
    Dim insW As Double, insH As Double
    insW = 700: insH = 460
    On Error Resume Next
    insW = frm.InsideWidth
    insH = frm.InsideHeight
    On Error GoTo EH

    panelW = insW * 0.27
    If panelW < 190 Then panelW = 190
    If panelW > 270 Then panelW = 270
    leftX = insW - panelW - 12
    If leftX < 8 Then leftX = 8
    topY = insH * 0.17
    If topY < 40 Then topY = 60
    panelH = insH * 0.64
    If panelH < 140 Then panelH = 300
    Dim availW As Double: availW = panelW

    On Error Resume Next
    mDiag = "cont=" & TypeName(cont) & "; leftX=" & CStr(CLng(leftX)) & _
            "; panelW=" & CStr(CLng(panelW)) & "; insW=" & CStr(CLng(insW))
    On Error GoTo EH

    Set mLblTitle = cont.Controls.Add("Forms.Label.1", "lblKarticaDetaljiNaslov", True)
    With mLblTitle
        .Left = leftX
        .top = topY - 16
        If .top < 0 Then .top = topY
        .width = availW
        .Height = 14
        .caption = "DETALJI OTKUPA (klik na red)"
    End With
    On Error Resume Next
    StyleListHeaderLabel mLblTitle
    On Error GoTo EH

    Set mLst = cont.Controls.Add("Forms.ListBox.1", "lstKarticaDetalji", True)
    With mLst
        .Left = leftX
        .top = topY
        .width = availW
        .Height = panelH
        .ColumnCount = 2
        .ColumnWidths = CStr(Int(availW * 0.4)) & ";" & CStr(Int(availW * 0.57))
    End With
    On Error Resume Next
    StyleListBox mLst
    mLblTitle.ZOrder 0          ' fmTop -> u prvi plan (preko MultiPage-a)
    mLst.ZOrder 0
    On Error GoTo EH

    mBuilt = True
    Exit Sub
EH:
    mDiag = mDiag & " | ERR " & CStr(Err.Number) & ": " & Err.description
    LogErr "modKarticaDetalji.KarticaDetalji_Ensure"
End Sub

Public Sub KarticaDetalji_Clear()
    On Error Resume Next
    mCurOtkupID = ""
    If Not mLst Is Nothing Then mLst.Clear
    If Not mLblTitle Is Nothing Then mLblTitle.caption = "DETALJI OTKUPA (klik na red)"
End Sub

' OtkupID trenutno prikazanog reda ("" ako trenutni red nije otkup).
Public Function KarticaDetalji_CurrentOtkupID() As String
    KarticaDetalji_CurrentOtkupID = mCurOtkupID
End Function

' Geometrija panela detalja (za pozicioniranje dugmeta ispod njega). 0 ako panel
' jos ne postoji.
Public Sub KarticaDetalji_PanelRect(ByRef l As Double, ByRef t As Double, _
                                    ByRef w As Double, ByRef h As Double)
    l = 0: t = 0: w = 0: h = 0
    On Error Resume Next
    If Not mLst Is Nothing Then
        l = mLst.Left
        t = mLst.Top
        w = mLst.width
        h = mLst.Height
    End If
End Sub

Public Sub KarticaDetalji_SetVisible(ByVal b As Boolean)
    On Error Resume Next
    If b Then
        If Not mLst Is Nothing Then mLst.Visible = True
        If Not mLblTitle Is Nothing Then mLblTitle.Visible = True
    ElseIf mFormLevel Then
        ' Sakrij samo forma-level kontrole; page-deca se kriju zajedno sa stranicom.
        If Not mLst Is Nothing Then mLst.Visible = False
        If Not mLblTitle Is Nothing Then mLblTitle.Visible = False
    End If
End Sub

' Prikazi detalje za trenutno izabran red lstKartica.
Public Sub KarticaDetalji_ShowForRow(ByVal frm As Object, ByVal lstKartica As MSForms.ListBox)
    On Error GoTo EH
    If lstKartica Is Nothing Then Exit Sub

    KarticaDetalji_Ensure frm, lstKartica
    If mLst Is Nothing Then
        ' Panel nije kreiran -> jednom prijavi razlog (da ne ostane "nista se ne desava").
        If Not mDiagShown Then
            mDiagShown = True
            MsgBox "Panel 'Detalji otkupa' nije kreiran." & vbCrLf & mDiag, _
                   vbExclamation, "OtkupApp"
        End If
        Exit Sub
    End If

    ' Panel postoji -> osiguraj da je vidljiv i u prvom planu.
    mLst.Visible = True
    If Not mLblTitle Is Nothing Then mLblTitle.Visible = True
    On Error Resume Next
    mLst.ZOrder 0          ' u prvi plan (preko praznog dela liste)
    mLblTitle.ZOrder 0
    On Error GoTo EH

    Dim idx As Long
    idx = lstKartica.ListIndex
    If idx < 0 Then Exit Sub

    Dim refKey As String
    On Error Resume Next
    refKey = CStr(lstKartica.List(idx, KART_REFKEY_COL))
    On Error GoTo EH

    mLst.Clear

    If Left$(refKey, 4) = "OTK|" Then
        ShowOtkupDetails Mid$(refKey, 5)
    Else
        ShowBasicRow lstKartica, idx
    End If
    Exit Sub
EH:
    LogErr "modKarticaDetalji.KarticaDetalji_ShowForRow"
End Sub

' Ne-otkup red (novac/agrohemija/UKUPNO): prikazi osnovne stavke iz same liste.
Private Sub ShowBasicRow(ByVal lstKartica As MSForms.ListBox, ByVal idx As Long)
    On Error Resume Next
    mCurOtkupID = ""
    mLblTitle.caption = "DETALJI STAVKE"
    AddPair "Datum", CStr(lstKartica.List(idx, 0))
    AddPair "Broj dok.", CStr(lstKartica.List(idx, 1))
    AddPair "Opis", CStr(lstKartica.List(idx, 2))
    If Len(Trim$(CStr(lstKartica.List(idx, 3)))) > 0 Then AddPair "Zaduzenje", CStr(lstKartica.List(idx, 3))
    If Len(Trim$(CStr(lstKartica.List(idx, 4)))) > 0 Then AddPair "Razduzenje", CStr(lstKartica.List(idx, 4))
    AddPair "Saldo", CStr(lstKartica.List(idx, 5))
    If Len(Trim$(CStr(lstKartica.List(idx, 6)))) > 0 Then AddPair "Saldo amb.", CStr(lstKartica.List(idx, 6))
End Sub

' Otkup red: sve bitne stavke otkupnog lista (polja iz frmOtkup), read-only.
Private Sub ShowOtkupDetails(ByVal otkupID As String)
    On Error GoTo EH
    mCurOtkupID = ""
    mLblTitle.caption = "DETALJI OTKUPA"

    Dim data As Variant
    data = GetTableData(TBL_OTKUP)
    If Not IsArray(data) Then Exit Sub

    Dim cId As Long
    cId = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    If cId = 0 Then Exit Sub

    Dim r As Long, found As Long
    found = 0
    For r = 1 To UBound(data, 1)
        If CStr(data(r, cId)) = otkupID Then
            found = r
            Exit For
        End If
    Next r
    If found = 0 Then
        AddPair "Otkup", "Nije pronadjen (" & otkupID & ")"
        Exit Sub
    End If
    mCurOtkupID = otkupID   ' validan otkup red -> dostupan za "Stampaj otkupni list"

    Dim datum As Variant: datum = CellVal(data, found, COL_OTK_DATUM)
    Dim brDok As String: brDok = CStr(CellVal(data, found, COL_OTK_BR_DOK))
    Dim stanicaID As String: stanicaID = CStr(CellVal(data, found, COL_OTK_STANICA))
    Dim koopID As String: koopID = CStr(CellVal(data, found, COL_OTK_KOOPERANT))
    Dim parcelaID As String: parcelaID = CStr(CellVal(data, found, COL_OTK_PARCELA))
    Dim vrsta As String: vrsta = CStr(CellVal(data, found, COL_OTK_VRSTA))
    Dim sorta As String: sorta = CStr(CellVal(data, found, COL_OTK_SORTA))
    Dim klasa As String: klasa = CStr(CellVal(data, found, COL_OTK_KLASA))
    Dim kol As Double: kol = NumOf(CellVal(data, found, COL_OTK_KOLICINA))
    Dim cena As Double: cena = NumOf(CellVal(data, found, COL_OTK_CENA))
    Dim tipAmb As String: tipAmb = CStr(CellVal(data, found, COL_OTK_TIP_AMB))
    Dim kolAmb As Double: kolAmb = NumOf(CellVal(data, found, COL_OTK_KOL_AMB))
    Dim vozacID As String: vozacID = CStr(CellVal(data, found, COL_OTK_VOZAC))
    Dim brZbirne As String: brZbirne = CStr(CellVal(data, found, COL_OTK_BROJ_ZBIRNE))
    Dim novac As Double: novac = NumOf(CellVal(data, found, COL_OTK_NOVAC))
    Dim primalac As String: primalac = CStr(CellVal(data, found, COL_OTK_PRIMALAC))
    Dim bruto As Double: bruto = NumOf(CellVal(data, found, COL_OTK_BRUTO))
    Dim ambIzdata As Double: ambIzdata = NumOf(CellVal(data, found, COL_OTK_KOL_AMB_IZDATA))

    If IsDate(datum) Then AddPair "Datum", Format$(CDate(datum), "d.m.yyyy")
    AddPair "Broj dokumenta", brDok
    AddPair "Otkupno mesto", DisplayStanica(stanicaID)
    AddPair "Kooperant", DisplayKooperant(koopID)
    If Len(Trim$(parcelaID)) > 0 Then AddPair "Parcela", DisplayParcela(parcelaID)
    AddPair "Vrsta voca", vrsta
    If Len(Trim$(sorta)) > 0 Then AddPair "Sorta", sorta
    AddPair "Klasa", klasa
    AddPair "Kolicina", FmtKolicina(kol) & " kg"
    If bruto > 0 Then AddPair "Bruto", FmtKolicina(bruto) & " kg"
    AddPair "Cena", Format$(cena, "#,##0.00")
    AddPair "Vrednost", Format$(kol * cena, "#,##0.00")
    If Len(Trim$(tipAmb)) > 0 Then AddPair "Tip ambalaze", tipAmb
    If kolAmb > 0 Then AddPair "Ambalaza (gajbe)", Format$(kolAmb, "#,##0")
    If ambIzdata > 0 Then AddPair "Izdata ambalaza", Format$(ambIzdata, "#,##0")
    If Len(Trim$(vozacID)) > 0 Then AddPair "Vozac", DisplayVozac(vozacID)
    If Len(Trim$(brZbirne)) > 0 Then AddPair "Broj zbirne", brZbirne
    If novac > 0 Then AddPair "Isplaceno (kes)", Format$(novac, "#,##0.00")
    If Len(Trim$(primalac)) > 0 Then AddPair "Primalac novca", primalac
    Exit Sub
EH:
    LogErr "modKarticaDetalji.ShowOtkupDetails"
End Sub

Private Sub AddPair(ByVal labelText As String, ByVal valueText As String)
    On Error Resume Next
    mLst.AddItem labelText
    mLst.List(mLst.ListCount - 1, 1) = valueText
End Sub

' Citanje celije tblOtkup po imenu kolone (schema-safe; 0 -> prazno).
Private Function CellVal(ByVal data As Variant, ByVal r As Long, ByVal colName As String) As Variant
    Dim c As Long
    c = GetColumnIndex(TBL_OTKUP, colName)
    If c >= 1 Then
        CellVal = data(r, c)
    Else
        CellVal = ""
    End If
End Function

Private Function NumOf(ByVal v As Variant) As Double
    If IsNumeric(v) Then NumOf = CDbl(v)
End Function

Private Function DisplayStanica(ByVal stanicaID As String) As String
    DisplayStanica = stanicaID
    If Len(Trim$(stanicaID)) = 0 Then Exit Function
    On Error Resume Next
    Dim nz As String
    nz = CStr(LookupValue(TBL_STANICE, "StanicaID", stanicaID, "Naziv"))
    If Len(Trim$(nz)) > 0 Then DisplayStanica = nz & " (" & stanicaID & ")"
End Function

Private Function DisplayKooperant(ByVal koopID As String) As String
    DisplayKooperant = koopID
    If Len(Trim$(koopID)) = 0 Then Exit Function
    On Error Resume Next
    Dim ime As String, pr As String
    ime = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", koopID, "Ime"))
    pr = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", koopID, "Prezime"))
    Dim nm As String: nm = Trim$(ime & " " & pr)
    If Len(nm) > 0 Then DisplayKooperant = nm & " (" & koopID & ")"
End Function

Private Function DisplayVozac(ByVal vozacID As String) As String
    DisplayVozac = vozacID
    If Len(Trim$(vozacID)) = 0 Then Exit Function
    On Error Resume Next
    Dim ime As String, pr As String
    ime = CStr(LookupValue(TBL_VOZACI, "VozacID", vozacID, "Ime"))
    pr = CStr(LookupValue(TBL_VOZACI, "VozacID", vozacID, "Prezime"))
    Dim nm As String: nm = Trim$(ime & " " & pr)
    If Len(nm) > 0 Then DisplayVozac = nm & " (" & vozacID & ")"
End Function

Private Function DisplayParcela(ByVal parcelaID As String) As String
    DisplayParcela = parcelaID
    If Len(Trim$(parcelaID)) = 0 Then Exit Function
    On Error Resume Next
    Dim kat As String, kult As String
    kat = CStr(LookupValue(TBL_PARCELE, COL_PAR_ID, parcelaID, COL_PAR_KAT_BROJ))
    kult = CStr(LookupValue(TBL_PARCELE, COL_PAR_ID, parcelaID, COL_PAR_KULTURA))
    Dim s As String: s = Trim$(kat)
    If Len(Trim$(kult)) > 0 Then s = Trim$(s & " | " & kult)
    If Len(s) > 0 Then DisplayParcela = s
End Function

