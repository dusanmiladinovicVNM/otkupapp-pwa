VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmIzvestaj 
   Caption         =   "UserForm1"
   ClientHeight    =   10065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14535
   OleObjectBlob   =   "frmIzvestaj.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmIzvestaj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmIzvestaj v2.1 – Reporting
' GEÄNDERT: tblIsporuka ? tblPrijemnica
' NEU: Manjak-Tab (Page 6)
' Tabs: Saldo | Otkupljena roba | Primljena ambalaza |
'       Isplata | Zbirni po OM | Prosecna cena | Manjak
' ============================================================

Private m_SetupDone As Boolean
Private m_IsInitializing As Boolean
Private m_IsChangingToggle As Boolean
Private m_IsChangingTipToggle As Boolean
Private m_IsRefreshing As Boolean

Private mChromeRemoved As Boolean

Private Sub RemoveTitleBar()
    Dim hwnd As LongPtr
    Dim style As Long

    hwnd = FindWindow("ThunderDFrame", Me.Caption)

    If hwnd <> 0 Then
        style = GetWindowLong(hwnd, GWL_STYLE)
        style = style And Not WS_CAPTION
        SetWindowLong hwnd, GWL_STYLE, style
        DrawMenuBar hwnd
    End If
End Sub

Private Sub UserForm_Activate()
    If Not mChromeRemoved Then
        Me.Caption = ""
        RemoveTitleBar
        mChromeRemoved = True
    End If
    
    ApplyTheme Me, BG_MAIN
    If m_SetupDone Then Exit Sub
    m_SetupDone = True

    m_IsInitializing = True

    'ApplyTheme Me
    Me.Caption = "Izvestaji"

    ' defaults (won't run events because we're initializing)
    SetTipToggle "tglPojedinacni"
    SetEntitetToggle "tglOM"

    ' fill controls
    LoadEntiteti

    cmbVrstaRobe.Clear
    cmbVrstaRobe.AddItem ""
    Dim kulture As Variant, i As Long
    kulture = GetLookupList(TBL_KULTURE, "VrstaVoca")
    If IsArray(kulture) Then
        For i = LBound(kulture) To UBound(kulture)
            cmbVrstaRobe.AddItem CStr(kulture(i))
        Next i
    End If

    txtDatumOd.Value = "1.1." & Year(Date)
    txtDatumDo.Value = Format$(Date, "d.m.yyyy")

    SetupListBoxes
    UpdateReportMode

    ' ? unlock only now
    m_IsInitializing = False

    ' now it’s safe to auto-run once
    AutoRefresh

End Sub
Private Sub AutoRefresh()

    If m_IsInitializing Then Exit Sub
    If m_IsRefreshing Then Exit Sub

    If tglPojedinacni.Value Then
        If cmbEntitet.ListIndex < 0 Or cmbEntitet.Value = "" Then Exit Sub
    End If

    m_IsRefreshing = True
    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    btnUnos_Click

CleanExit:
    Application.ScreenUpdating = True
    m_IsRefreshing = False
    Exit Sub

CleanFail:
    ' optional: Debug.Print Err.Number, Err.Description
    Resume CleanExit
End Sub
' ============================================================
' KASKADIERUNG
' ============================================================

Private Function GetActiveEntitetTip() As String
    If tglOM.Value Then
        GetActiveEntitetTip = "Otkupna mesta"
    ElseIf tglKupci.Value Then
        GetActiveEntitetTip = "Kupci"
    ElseIf tglVozaci.Value Then
        GetActiveEntitetTip = "Vozaci"
    ElseIf tglKooperanti.Value Then
        GetActiveEntitetTip = "Kooperanti"
    Else
        GetActiveEntitetTip = "Otkupna mesta" ' fallback
    End If
End Function

Private Sub SetEntitetToggle(ByVal activeName As String)

    If m_IsChangingToggle Then Exit Sub

    m_IsChangingToggle = True

    tglOM.Value = (activeName = "tglOM")
    tglKupci.Value = (activeName = "tglKupci")
    tglVozaci.Value = (activeName = "tglVozaci")
    tglKooperanti.Value = (activeName = "tglKooperanti")

    m_IsChangingToggle = False

    UpdateEntitetToggleUI

End Sub

Private Sub UpdateEntitetToggleUI()
    ' najjednostavnije: bold na aktivnom
    tglOM.Font.Bold = tglOM.Value
    tglKupci.Font.Bold = tglKupci.Value
    tglVozaci.Font.Bold = tglVozaci.Value
    tglKooperanti.Font.Bold = tglKooperanti.Value
End Sub

Private Sub tglOM_Click()
    If m_IsChangingToggle Or m_IsInitializing Then Exit Sub
    SetEntitetToggle "tglOM"
    LoadEntiteti
    UpdateReportMode
    AutoRefresh
End Sub

Private Sub tglKupci_Click()
    If m_IsChangingToggle Or m_IsInitializing Then Exit Sub
    SetEntitetToggle "tglKupci"
    LoadEntiteti
    UpdateReportMode
    AutoRefresh
End Sub

Private Sub tglVozaci_Click()
    If m_IsChangingToggle Or m_IsInitializing Then Exit Sub
    SetEntitetToggle "tglVozaci"
    LoadEntiteti
    UpdateReportMode
    AutoRefresh
End Sub
Private Sub tglKooperanti_Click()
    If m_IsChangingToggle Then Exit Sub
    SetEntitetToggle "tglKooperanti"
    LoadEntiteti
    UpdateReportMode
    AutoRefresh
End Sub

Private Function IsZbirniMode() As Boolean
    IsZbirniMode = tglZbirni.Value
End Function

Private Sub SetTipToggle(ByVal activeName As String)

    If m_IsChangingTipToggle Then Exit Sub
    m_IsChangingTipToggle = True

    tglPojedinacni.Value = (activeName = "tglPojedinacni")
    tglZbirni.Value = (activeName = "tglZbirni")

    m_IsChangingTipToggle = False

    UpdateTipToggleUI
End Sub

Private Sub UpdateTipToggleUI()
    tglPojedinacni.Font.Bold = tglPojedinacni.Value
    tglZbirni.Font.Bold = tglZbirni.Value
End Sub

Private Sub tglPojedinacni_Click()
    If m_IsChangingTipToggle Or m_IsInitializing Then Exit Sub
    SetTipToggle "tglPojedinacni"
    UpdateReportMode
    AutoRefresh
End Sub

Private Sub tglZbirni_Click()
    If m_IsChangingTipToggle Or m_IsInitializing Then Exit Sub
    SetTipToggle "tglZbirni"
    UpdateReportMode
    AutoRefresh
End Sub

Private Sub LoadEntiteti()
    cmbEntitet.Clear

    Select Case GetActiveEntitetTip()
        Case "Otkupna mesta"
            FillCmb cmbEntitet, GetLookupList(TBL_STANICE, "Naziv")
        Case "Kupci"
            FillCmb cmbEntitet, GetLookupList(TBL_KUPCI, "Naziv")
        Case "Vozaci"
            FillCmb cmbEntitet, GetVozacDisplayList()
        Case "Kooperanti"
            Dim koopData As Variant
            koopData = GetTableData(TBL_KOOPERANTI)
            If IsArray(koopData) Then
                Dim colID As Long, colIme As Long, colPrezime As Long, colAktivan As Long
                colID = GetColumnIndex(TBL_KOOPERANTI, "KooperantID")
                colIme = GetColumnIndex(TBL_KOOPERANTI, "Ime")
                colPrezime = GetColumnIndex(TBL_KOOPERANTI, "Prezime")
                colAktivan = GetColumnIndex(TBL_KOOPERANTI, "Aktivan")
                Dim k As Long
                For k = 1 To UBound(koopData, 1)
                    If CStr(koopData(k, colAktivan)) = STATUS_AKTIVAN Then
                        cmbEntitet.AddItem CStr(koopData(k, colIme)) & " " & _
                                          CStr(koopData(k, colPrezime)) & " (" & _
                                          CStr(koopData(k, colID)) & ")"
                    End If
                Next k
            End If
    End Select

    If cmbEntitet.ListCount > 0 Then cmbEntitet.ListIndex = 0
End Sub
Private Sub cmbEntitet_Change()
    AutoRefresh
End Sub

' ============================================================
' ZBIRNI/POJEDINACNI
' ============================================================

Private Sub UpdateReportMode()
    Dim isPojed As Boolean
    isPojed = tglPojedinacni.Value
    cmbEntitet.Enabled = isPojed
    
    ' Alle erstmal ausblenden
    Dim pg As Long
    For pg = 0 To mpReports.Pages.count - 1
        mpReports.Pages(pg).Visible = False
    Next pg
    
    If isPojed Then
        Select Case GetActiveEntitetTip()
            Case "Otkupna mesta"
                mpReports.Pages(0).Visible = True  ' Saldo OM
                mpReports.Pages(2).Visible = True  ' Otkupljena roba
                mpReports.Pages(3).Visible = True  ' Ambalaza
                mpReports.Pages(4).Visible = True  ' Isplata
                mpReports.Pages(6).Visible = True  ' Prosecna cena
            Case "Kupci"
                mpReports.Pages(1).Visible = True  ' Saldo Kupci
                mpReports.Pages(2).Visible = True  ' Otkupljena roba
                mpReports.Pages(3).Visible = True  ' Ambalaza
                mpReports.Pages(6).Visible = True  ' Prosecna cena
                mpReports.Pages(7).Visible = True  ' Manjak
            Case "Vozaci"
                mpReports.Pages(3).Visible = True  ' Ambalaza
                mpReports.Pages(7).Visible = True  ' Manjak
            Case "Kooperanti"
                mpReports.Pages(8).Visible = True  ' Kartica
        End Select
    Else
        mpReports.Pages(5).Visible = True  ' Zbirni
        mpReports.Pages(6).Visible = True  ' Prosecna cena
        mpReports.Pages(7).Visible = True  ' Manjak
    End If
    
    ' posle Visible=True podesavanja:
    Dim firstVisible As Long
    firstVisible = -1
    For pg = 0 To mpReports.Pages.count - 1
        If mpReports.Pages(pg).Visible Then
            firstVisible = pg
            Exit For
        End If
    Next pg
    If firstVisible >= 0 Then mpReports.Value = firstVisible

End Sub

' ============================================================
' LISTBOX SETUP
' ============================================================

Private Sub SetupListBoxes()
    
    With lstSaldoOM
        .ColumnCount = 7
        .ColumnWidths = "120;65;80;80;80;80;50"
    End With
    
    ' Saldo Kupci: Vrsta | Kolicina | Cena | Vrednost |  Novac | Saldo | Ambalaza
    With lstSaldoKupci
        .ColumnCount = 7
        .ColumnWidths = "80;80;70;80;80;80;60"
    End With
    
    With lstOtkupRoba
        .ColumnCount = 10
        .ColumnWidths = "45;50;45;25;80;50;50;45;45;40"
    End With
    
    With lstAmbalaza
        .ColumnCount = 6
        .ColumnWidths = "55;90;45;80;50;50"
    End With
    
    With lstIsplata
        .ColumnCount = 5
        .ColumnWidths = "120;80;80;80;80"
    End With
    
    With lstZbirni
        .ColumnCount = 5
        .ColumnWidths = "100;80;80;80;80"
    End With
    
    With lstProsecnaCena
        .ColumnCount = 4
        .ColumnWidths = "100;80;80;80"
    End With
    
    With lstManjak
        .ColumnCount = 6
        .ColumnWidths = "70;70;70;60;55;60"
    End With
    
    With lstKartica
        .ColumnCount = 6
        .ColumnWidths = "60;70;140;80;80;80"
    End With
End Sub

' ============================================================
' HAUPTAKTION
' ============================================================

Private Sub btnUnos_Click()
    On Error GoTo EH

    Dim datumOd As Date, datumDo As Date
    datumOd = CDate(txtDatumOd.Value)
    datumDo = CDate(txtDatumDo.Value)

    Dim zbirni As Boolean
    zbirni = IsZbirniMode()

    Dim entitetID As String
    Dim entitetTip As String

    ' Entitet tip je uvek potreban (OM/Kupac/Vozac),
    ' entitetID je potreban samo za POJEDINACNI
    Select Case GetActiveEntitetTip()
        Case "Otkupna mesta": entitetTip = "OM"
        Case "Kupci":         entitetTip = "Kupac"
        Case "Vozaci":        entitetTip = "Vozac"
        Case "Kooperanti":    entitetTip = "Kooperant"
        Case Else:            entitetTip = "OM"
    End Select

    If Not zbirni Then
        If cmbEntitet.Value = "" Then
            MsgBox "Izaberite entitet!", vbExclamation, APP_NAME
            Exit Sub
        End If

        Select Case entitetTip
            Case "OM"
                entitetID = CStr(LookupValue(TBL_STANICE, "Naziv", cmbEntitet.Value, "StanicaID"))
            Case "Kupac"
                entitetID = CStr(LookupValue(TBL_KUPCI, "Naziv", cmbEntitet.Value, "KupacID"))
            Case "Vozac"
                entitetID = ExtractIDFromDisplay(cmbEntitet.Value)
           Case "Kooperant"
                entitetID = ExtractIDFromDisplay(cmbEntitet.Value)
        End Select
    Else
        entitetID = ""
    End If

    Application.ScreenUpdating = False

    If zbirni Then
        GenerateZbirniReport datumOd, datumDo, entitetTip
        GenerateProsecnaCenaReport entitetTip, "", datumOd, datumDo
        GenerateManjakReport entitetTip, "", datumOd, datumDo
    Else
        GenerateSaldoReport entitetTip, entitetID, datumOd, datumDo
        GenerateOtkupRobaReport entitetTip, entitetID, datumOd, datumDo
        GenerateAmbalazeReport entitetTip, entitetID, datumOd, datumDo
        GenerateIsplataReport entitetTip, entitetID, datumOd, datumDo
        GenerateProsecnaCenaReport entitetTip, entitetID, datumOd, datumDo
        GenerateManjakReport entitetTip, entitetID, datumOd, datumDo
        If entitetTip = "Kooperant" Then
            GenerateKarticaReport entitetID, datumOd, datumDo
        End If
    End If

    WriteReportTables entitetTip, entitetID, datumOd, datumDo, zbirni

    Application.ScreenUpdating = True
    Exit Sub

EH:
    LogErr "frmIzvestaj.btnUnos"
    Application.ScreenUpdating = True
    MsgBox "Greska pri ucitavanju izvestaja: " & Err.Description, vbCritical, APP_NAME
End Sub

Private Sub UpdateUnosButtonState()

    ' If Zbirni ? no entitet required
    If tglZbirni.Value Then
        btnUnos.Enabled = True
        Exit Sub
    End If

    ' If Pojedinacni ? entitet must be selected
    btnUnos.Enabled = (cmbEntitet.ListIndex >= 0 And cmbEntitet.Value <> "")

End Sub

' ============================================================
' SALDO
' ============================================================

Private Sub GenerateSaldoReport(ByVal entitetTip As String, ByVal entitetID As String, _
                                ByVal datumOd As Date, ByVal datumDo As Date)
    If entitetTip = "OM" Then
        GenerateSaldoOM entitetID, datumOd, datumDo
    ElseIf entitetTip = "Kupac" Then
        GenerateSaldoKupci entitetID, datumOd, datumDo
    End If
End Sub

Private Sub GenerateSaldoOM(ByVal stanicaID As String, _
                            ByVal datumOd As Date, ByVal datumDo As Date)
    lstSaldoOM.Clear
    
    Dim data As Variant
    data = ReportSaldoOM(stanicaID, datumOd, datumDo)
    If IsEmpty(data) Then Exit Sub
    
    Dim i As Long, j As Long
    For i = 1 To UBound(data, 1)
        lstSaldoOM.AddItem CStr(data(i, 1))
        For j = 2 To UBound(data, 2)
            If IsNumeric(data(i, j)) Then
                If j = 7 Then  ' Ambalaza
                    lstSaldoOM.List(lstSaldoOM.ListCount - 1, j - 1) = Format$(data(i, j), "#,##0")
                Else
                    lstSaldoOM.List(lstSaldoOM.ListCount - 1, j - 1) = Format$(data(i, j), "#,##0.00")
                End If
            Else
                lstSaldoOM.List(lstSaldoOM.ListCount - 1, j - 1) = CStr(data(i, j))
            End If
        Next j
    Next i
End Sub

Private Sub GenerateSaldoKupci(ByVal kupacID As String, _
                               ByVal datumOd As Date, ByVal datumDo As Date)
    lstSaldoKupci.Clear
    
    Dim data As Variant
    data = ReportSaldoKupci(kupacID, datumOd, datumDo)
    If IsEmpty(data) Then Exit Sub
    
    Dim fmts As Variant
    fmts = Array("", "#,##0.00", "#,##0.00", "#,##0.00", "#,##0.00", "#,##0.00", "#,##0")
    
    Dim i As Long, j As Long
    For i = 1 To UBound(data, 1)
        lstSaldoKupci.AddItem CStr(data(i, 1))
        For j = 2 To UBound(data, 2)
            If IsNumeric(data(i, j)) And data(i, j) <> "" Then
                lstSaldoKupci.List(lstSaldoKupci.ListCount - 1, j - 1) = _
                    Format$(CDbl(data(i, j)), CStr(fmts(j - 1)))
            ElseIf CStr(data(i, j)) <> "" Then
                lstSaldoKupci.List(lstSaldoKupci.ListCount - 1, j - 1) = CStr(data(i, j))
            End If
        Next j
    Next i
End Sub

' ============================================================
' OTKUPLJENA ROBA
' ============================================================

Private Sub GenerateOtkupRobaReport(ByVal entitetTip As String, ByVal entitetID As String, _
                                    ByVal datumOd As Date, ByVal datumDo As Date)
    lstOtkupRoba.Clear
    
    Dim data As Variant
    data = ReportOtkupRoba(entitetTip, entitetID, datumOd, datumDo)
    If IsEmpty(data) Then Exit Sub
    
    Dim i As Long, j As Long
    For i = 1 To UBound(data, 1)
        If IsDate(data(i, 1)) Then
            lstOtkupRoba.AddItem Format$(CDate(data(i, 1)), "d.m.yy")
        Else
            lstOtkupRoba.AddItem CStr(IIf(IsEmpty(data(i, 1)), "", data(i, 1)))
        End If
        
        For j = 2 To UBound(data, 2)
            If j = 10 Then
                ' Manjak% formatieren
                If IsNumeric(data(i, j)) And Not IsEmpty(data(i, j)) Then
                    lstOtkupRoba.List(lstOtkupRoba.ListCount - 1, j - 1) = Format$(CDbl(data(i, j)), "0.00") & "%"
                End If
            ElseIf IsNumeric(data(i, j)) And Not IsEmpty(data(i, j)) Then
                lstOtkupRoba.List(lstOtkupRoba.ListCount - 1, j - 1) = Format$(CDbl(data(i, j)), "#,##0")
            Else
                lstOtkupRoba.List(lstOtkupRoba.ListCount - 1, j - 1) = _
                    CStr(IIf(IsEmpty(data(i, j)), "", data(i, j)))
            End If
        Next j
    Next i
End Sub

' ============================================================
' AMBALAZA (unverändert)
' ============================================================

Private Sub GenerateAmbalazeReport(ByVal entitetTip As String, ByVal entitetID As String, _
                                   ByVal datumOd As Date, ByVal datumDo As Date)
    lstAmbalaza.Clear
    
    Dim isZb As Boolean
    isZb = IsZbirniMode()
    
    Dim data As Variant
    data = ReportAmbalaza(entitetTip, entitetID, datumOd, datumDo, isZb)
    If IsEmpty(data) Then Exit Sub
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If IsDate(data(i, 1)) Then
            lstAmbalaza.AddItem Format$(CDate(data(i, 1)), "d.m.yy")
        Else
            lstAmbalaza.AddItem CStr(data(i, 1))
        End If
        lstAmbalaza.List(lstAmbalaza.ListCount - 1, 1) = CStr(data(i, 2))
        lstAmbalaza.List(lstAmbalaza.ListCount - 1, 2) = CStr(data(i, 3))
        lstAmbalaza.List(lstAmbalaza.ListCount - 1, 3) = CStr(data(i, 4))
        If IsNumeric(data(i, 5)) And data(i, 5) <> "" Then
            lstAmbalaza.List(lstAmbalaza.ListCount - 1, 4) = Format$(CLng(data(i, 5)), "#,##0")
        End If
        If IsNumeric(data(i, 6)) And data(i, 6) <> "" Then
            lstAmbalaza.List(lstAmbalaza.ListCount - 1, 5) = Format$(CLng(data(i, 6)), "#,##0")
        End If
    Next i
End Sub

' ============================================================
' ISPLATA (unverändert)
' ============================================================

Private Sub GenerateIsplataReport(ByVal entitetTip As String, ByVal entitetID As String, _
                                  ByVal datumOd As Date, ByVal datumDo As Date)
    lstIsplata.Clear
    
    Dim data As Variant
    data = ReportIsplata(entitetTip, entitetID, datumOd, datumDo)
    If IsEmpty(data) Then Exit Sub
    
    Dim i As Long, j As Long
    For i = 1 To UBound(data, 1)
        lstIsplata.AddItem CStr(IIf(IsEmpty(data(i, 1)), "", data(i, 1)))
        For j = 2 To UBound(data, 2)
            If IsNumeric(data(i, j)) And Not IsEmpty(data(i, j)) Then
                lstIsplata.List(lstIsplata.ListCount - 1, j - 1) = Format$(CDbl(data(i, j)), "#,##0.00")
            Else
                lstIsplata.List(lstIsplata.ListCount - 1, j - 1) = ""
            End If
        Next j
    Next i
End Sub

' ============================================================
' ZBIRNI PO OM (unverändert)
' ============================================================

Private Sub GenerateZbirniReport(ByVal datumOd As Date, ByVal datumDo As Date, _
                                 ByVal entitetTip As String)
    lstZbirni.Clear
    
    Dim data As Variant
    data = ReportZbirni(entitetTip, datumOd, datumDo)
    If IsEmpty(data) Then Exit Sub
    
    Dim isVozac As Boolean
    isVozac = (entitetTip = "Vozac")
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        lstZbirni.AddItem CStr(data(i, 1))
        
        If isVozac Then
            ' Amb Izlaz, Amb Vracena, Manjak kg, Manjak %
            If IsNumeric(data(i, 2)) Then lstZbirni.List(lstZbirni.ListCount - 1, 1) = Format$(CLng(data(i, 2)), "#,##0")
            If IsNumeric(data(i, 3)) Then lstZbirni.List(lstZbirni.ListCount - 1, 2) = Format$(CLng(data(i, 3)), "#,##0")
            If IsNumeric(data(i, 4)) Then lstZbirni.List(lstZbirni.ListCount - 1, 3) = Format$(CDbl(data(i, 4)), "#,##0.0") & " kg"
            If IsNumeric(data(i, 5)) Then lstZbirni.List(lstZbirni.ListCount - 1, 4) = Format$(CDbl(data(i, 5)), "#,##0.00") & "%"
        Else
            ' Vrsta, Kolicina, Vrednost, Prosek
            lstZbirni.List(lstZbirni.ListCount - 1, 1) = CStr(data(i, 2))
            If IsNumeric(data(i, 3)) Then lstZbirni.List(lstZbirni.ListCount - 1, 2) = Format$(CDbl(data(i, 3)), "#,##0.00")
            If IsNumeric(data(i, 4)) Then lstZbirni.List(lstZbirni.ListCount - 1, 3) = Format$(CDbl(data(i, 4)), "#,##0.00")
            If IsNumeric(data(i, 5)) Then
                If CDbl(data(i, 5)) > 0 Then
                    lstZbirni.List(lstZbirni.ListCount - 1, 4) = Format$(CDbl(data(i, 5)), "#,##0.00")
                End If
            End If
        End If
    Next i
End Sub

' ============================================================
' PROSECNA CENA
' ============================================================

Private Sub GenerateProsecnaCenaReport(ByVal entitetTip As String, ByVal entitetID As String, _
                                       ByVal datumOd As Date, ByVal datumDo As Date)
    lstProsecnaCena.Clear
    
    Dim data As Variant
    data = ReportProsecnaCena(entitetTip, entitetID, datumOd, datumDo)
    If IsEmpty(data) Then Exit Sub
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        lstProsecnaCena.AddItem CStr(data(i, 1))
        If IsNumeric(data(i, 2)) Then lstProsecnaCena.List(lstProsecnaCena.ListCount - 1, 1) = Format$(CDbl(data(i, 2)), "#,##0.00")
        If IsNumeric(data(i, 3)) Then lstProsecnaCena.List(lstProsecnaCena.ListCount - 1, 2) = Format$(CDbl(data(i, 3)), "#,##0.00")
        If IsNumeric(data(i, 4)) Then
            If CDbl(data(i, 4)) > 0 Then
                lstProsecnaCena.List(lstProsecnaCena.ListCount - 1, 3) = Format$(CDbl(data(i, 4)), "#,##0.00")
            End If
        End If
    Next i
End Sub


' ============================================================
' MANJAK (NEU)
' Header: BrojZbirne | Zbirna kg | Prijemnica kg | Manjak kg | Manjak % | Prosek gajbe
' ============================================================

Private Sub GenerateManjakReport(ByVal entitetTip As String, ByVal entitetID As String, _
                                 ByVal datumOd As Date, ByVal datumDo As Date)
    lstManjak.Clear
    If lstManjak Is Nothing Then Exit Sub
    
    Dim data As Variant
    data = ReportManjak(entitetTip, entitetID, datumOd, datumDo)
    If IsEmpty(data) Then Exit Sub
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        lstManjak.AddItem CStr(data(i, 1))
        If IsNumeric(data(i, 2)) Then lstManjak.List(lstManjak.ListCount - 1, 1) = Format$(CDbl(data(i, 2)), "#,##0.0")
        If IsNumeric(data(i, 3)) Then lstManjak.List(lstManjak.ListCount - 1, 2) = Format$(CDbl(data(i, 3)), "#,##0.0")
        If IsNumeric(data(i, 4)) Then lstManjak.List(lstManjak.ListCount - 1, 3) = Format$(CDbl(data(i, 4)), "#,##0.0")
        If IsNumeric(data(i, 5)) Then lstManjak.List(lstManjak.ListCount - 1, 4) = Format$(CDbl(data(i, 5)), "#,##0.00") & "%"
        If IsNumeric(data(i, 6)) Then
            If CDbl(data(i, 6)) > 0 Then
                lstManjak.List(lstManjak.ListCount - 1, 5) = Format$(CDbl(data(i, 6)), "#,##0.00")
            End If
        End If
    Next i
End Sub

Private Sub GenerateKarticaReport(ByVal entitetID As String, _
                                  ByVal datumOd As Date, ByVal datumDo As Date)
    lstKartica.Clear
    
    Dim ime As String, prezime As String
    ime = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", entitetID, "Ime"))
    prezime = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", entitetID, "Prezime"))
    lblKarticaKoop.Caption = "KARTICA: " & ime & " " & prezime & " (" & entitetID & ")"
    lblKarticaPeriod.Caption = "Period: " & txtDatumOd.Value & " - " & txtDatumDo.Value
    
    Dim data As Variant
    data = ReportKarticaKooperanta(entitetID, datumOd, datumDo)
    If IsEmpty(data) Then Exit Sub
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        ' Spalte 1: Datum
        If IsDate(data(i, 1)) Then
            lstKartica.AddItem Format$(CDate(data(i, 1)), "d.m.yyyy")
        Else
            lstKartica.AddItem CStr(IIf(IsEmpty(data(i, 1)), "", data(i, 1)))
        End If
        
        ' Spalte 2: BrojDok
        lstKartica.List(lstKartica.ListCount - 1, 1) = _
            CStr(IIf(IsEmpty(data(i, 2)), "", data(i, 2)))
        
        ' Spalte 3: Opis
        lstKartica.List(lstKartica.ListCount - 1, 2) = _
            CStr(IIf(IsEmpty(data(i, 3)), "", data(i, 3)))
        
        ' Spalten 4-6: Zaduzenje, Razduzenje, Saldo
        Dim j As Long
        For j = 4 To 6
            If IsNumeric(data(i, j)) And Not IsEmpty(data(i, j)) Then
                lstKartica.List(lstKartica.ListCount - 1, j - 1) = Format$(CDbl(data(i, j)), "#,##0.00")
            End If
        Next j
    Next i
End Sub

Private Sub btnStampajKarticu_Click()
    If cmbEntitet.ListIndex < 0 Then Exit Sub
    Dim koopID As String
    koopID = ExtractIDFromDisplay(cmbEntitet.Value)
    PrintKarticaPDF koopID, CDate(txtDatumOd.Value), CDate(txtDatumDo.Value)
End Sub

' ============================================================
' REPORT-TABELLEN
' ============================================================

Private Sub WriteReportTables(ByVal entitetTip As String, ByVal entitetID As String, _
                              ByVal datumOd As Date, ByVal datumDo As Date, _
                              ByVal zbirni As Boolean)
    Dim lo As ListObject
    Dim i As Long
    
    If Not zbirni Then
        If entitetTip = "OM" Then
            Set lo = SafeGetTable(TBL_RPT_SALDO_OM)
            If Not lo Is Nothing Then
                If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
                For i = 0 To lstSaldoOM.ListCount - 1
                    AppendRow TBL_RPT_SALDO_OM, Array( _
                        Format$(Date, "yyyy-mm-dd"), entitetID, _
                        lstSaldoOM.List(i, 0), lstSaldoOM.List(i, 1), _
                        lstSaldoOM.List(i, 2), lstSaldoOM.List(i, 3), _
                        lstSaldoOM.List(i, 4), "", lstSaldoOM.List(i, 5))
                Next i
            End If
            
        ElseIf entitetTip = "Kupac" Then
            Set lo = SafeGetTable(TBL_RPT_SALDO_KUPCI)
            If Not lo Is Nothing Then
                If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
                For i = 0 To lstSaldoKupci.ListCount - 1
                    AppendRow TBL_RPT_SALDO_KUPCI, Array( _
                        Format$(Date, "yyyy-mm-dd"), entitetID, _
                        lstSaldoKupci.List(i, 0), lstSaldoKupci.List(i, 1), _
                        lstSaldoKupci.List(i, 2), lstSaldoKupci.List(i, 3), _
                        lstSaldoKupci.List(i, 4), lstSaldoKupci.List(i, 5))
                Next i
            End If
        End If
    End If
    
    Set lo = SafeGetTable(TBL_RPT_MARZA)
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
        For i = 0 To lstProsecnaCena.ListCount - 1
            AppendRow TBL_RPT_MARZA, Array( _
                Format$(Date, "yyyy-mm-dd"), _
                lstProsecnaCena.List(i, 0), _
                lstProsecnaCena.List(i, 1), _
                lstProsecnaCena.List(i, 2), _
                "", "", "", lstProsecnaCena.List(i, 3))
        Next i
    End If
    
    If zbirni Then
        Set lo = SafeGetTable(TBL_RPT_ZBIRNI)
        If Not lo Is Nothing Then
            If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
            For i = 0 To lstZbirni.ListCount - 1
                AppendRow TBL_RPT_ZBIRNI, Array( _
                    Format$(Date, "yyyy-mm-dd"), _
                    lstZbirni.List(i, 0), lstZbirni.List(i, 1), _
                    lstZbirni.List(i, 2), lstZbirni.List(i, 3), _
                    lstZbirni.List(i, 4))
            Next i
        End If
    End If
End Sub

' ============================================================
' DRUCKEN
' ============================================================
Private Sub btnStampaj_Click()
    On Error GoTo EH
    Dim activeTab As Long
    activeTab = mpReports.Value
    
    Dim lst As MSForms.ListBox
    Dim title As String
    Dim headers As Variant
    
    Select Case activeTab
        Case 0: Set lst = lstSaldoOM: title = "Saldo OM"
            headers = Array("Kooperant", "Kolicina", "Vrednost", "Novac", "Ambalaza", "Saldo")
        Case 1: Set lst = lstSaldoKupci: title = "Saldo Kupci"
            headers = Array("Vrsta", "Kolicina", "Cena", "Vrednost", "Ambalaza", "Saldo")
        Case 2: Set lst = lstOtkupRoba: title = "Otkupljena roba"
            headers = Array("Datum", "Dokument", "Kolicina", "Vrednost")
        Case 3: Set lst = lstAmbalaza: title = "Primljena ambalaza"
            headers = Array("Datum", "Mesto", "Tip", "Dokument", "Ulaz", "Izlaz")
        Case 4: Set lst = lstIsplata: title = "Isplata"
            headers = Array("Datum", "Kooperant", "Primalac", "Iznos")
        Case 5: Set lst = lstZbirni: title = "Zbirni izvestaj"
            headers = Array("Entitet", "Vrsta/Info", "Kolicina", "Vrednost", "Prosek")
        Case 6: Set lst = lstProsecnaCena: title = "Prosecna cena"
            headers = Array("Vrsta", "Kolicina", "Vrednost", "Prosecna cena")
        Case 7: Set lst = lstManjak: title = "Manjak"
            headers = Array("Br.Zbirne", "Zbirna kg", "Prijemnica kg", "Manjak kg", "Manjak %", "Prosek gajbe")
    End Select
    
    If lst Is Nothing Then Exit Sub
    If lst.ListCount = 0 Then
        MsgBox "Nema podataka za stampu!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Dim data() As Variant
    ReDim data(1 To lst.ListCount, 1 To lst.ColumnCount)
    Dim i As Long, j As Long
    For i = 0 To lst.ListCount - 1
        For j = 0 To lst.ColumnCount - 1
            data(i + 1, j + 1) = lst.List(i, j)
        Next j
    Next i
    
    Dim entLabel As String
    entLabel = GetActiveEntitetTip()   ' "Otkupna mesta" / "Kupci" / "Vozaci"

    Dim entName As String
    If IsZbirniMode() Then
        entName = "Svi"
    Else
        entName = cmbEntitet.Value
    End If
    Dim fullTitle As String
    fullTitle = title & " - " & entLabel & ": " & entName & _
            " (" & txtDatumOd.Value & " - " & txtDatumDo.Value & ")"
    
    PrintIzvestaj data, fullTitle, headers
    Exit Sub
EH:
    LogErr "frmIzvestaj.btnStampaj"
    MsgBox "Greska pri stampanju: " & Err.Description, vbCritical, APP_NAME
End Sub

' ============================================================
' NAVIGATION & HELPER
' ============================================================

Private Sub btnPovratak_Click()
    Me.Hide
    frmOtkupAPP.Show
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Unload Me
        frmOtkupAPP.Show
    End If
End Sub




