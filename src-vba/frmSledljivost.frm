VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSledljivost 
   Caption         =   "UserForm1"
   ClientHeight    =   13815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20235
   OleObjectBlob   =   "frmSledljivost.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSledljivost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ============================================================
' frmOtkupniBlokovi - Sledljivost & Povezivanje
' ============================================================
Option Explicit

Private m_UnlinkedData As Variant
Private m_CandidateOtpIDs() As String
Private mChromeRemoved As Boolean

Private Sub UserForm_Activate()
    On Error GoTo EH
    
    EnsureUserFormChromeRemoved Me, mChromeRemoved
    
    ApplyTheme Me, BG_MAIN()
    ApplyThemeToControls Me
    
    ' Header zone
    On Error Resume Next
    StyleFrameTitleLabel lblKopf, "Izve" & ChrW(353) & "taj o sledljivosti"
    StyleSubtitle lblSubtitle, "Auto/manuelno povezivanje + sledljivost po zbirnoj"
    On Error GoTo EH
    
    ' Section headers
    On Error Resume Next
    StyleListHeaderLabel lblSidebarNepovezani
    lblSidebarNepovezani.caption = "NEPOVEZANI OTKUPI"
    
    StyleListHeaderLabel lblSidebarOtpremnice
    lblSidebarOtpremnice.caption = "MOGUCE OTPREMNICE"
    
    StyleListHeaderLabel lblSidebarTrace
    lblSidebarTrace.caption = "SLEDLJIVOST PO ZBIRNOJ"
    
    ' Accent linije
    lblAccentNepovezani.BackColor = BTN_ACTIVE()
    lblAccentNepovezani.BackStyle = fmBackStyleOpaque
    lblAccentNepovezani.BorderStyle = fmBorderStyleNone
    
    lblAccentOtpremnice.BackColor = BTN_ACTIVE()
    lblAccentOtpremnice.BackStyle = fmBackStyleOpaque
    lblAccentOtpremnice.BorderStyle = fmBorderStyleNone
    
    lblAccentTrace.BackColor = BTN_ACTIVE()
    lblAccentTrace.BackStyle = fmBackStyleOpaque
    lblAccentTrace.BorderStyle = fmBorderStyleNone
    On Error GoTo EH
    
    ' Action buttons
    StylePrimaryButton btnAutoLink, "Automatsko povezivanje"
    StylePrimaryButton btnPovezi, Poruka("SLED_LBL_POVEZI")
    StylePrimaryButton btnStampaj, Poruka("FAK_LBL_STAMPAJ")
    StyleExitButton btnPovratak, "Povratak"
    
    ' Filter label
    On Error Resume Next
    StyleLabel lblZbirnaLabel, TXT_MUTED(), False
    lblZbirnaLabel.Font.Size = FONT_SIZE_SMALL
    On Error GoTo EH
    
    ' Status label
    On Error Resume Next
    StyleLabel lblStatus, TXT_LIGHT(), True
    On Error GoTo EH
    
    ' Column headers
    SetupAllColumnHeaders
    
    Exit Sub

EH:
    LogErr "frmSledljivost.UserForm_Activate"
End Sub

Private Sub UserForm_Initialize()
    SetupListBoxes
    LoadZbirne
    LoadNepovezani         ' DODATO - operator vidi nepovezane odmah
    UpdateStatus
End Sub

' ============================================================
' SETUP
' ============================================================

Private Sub SetupAllColumnHeaders()
    On Error Resume Next
    
    ' Nepovezani Otkupi (kolone 1-6 vidljive, OtkupID hidden)
    SetColumnHeader lbl_H_NEP1, "Datum"
    SetColumnHeader lbl_H_NEP2, "Stanica"
    SetColumnHeader lbl_H_NEP3, "Vozac"
    SetColumnHeader lbl_H_NEP4, "Kooperant"
    SetColumnHeader lbl_H_NEP5, "Koli" & ChrW(269) & "ina"
    SetColumnHeader lbl_H_NEP6, "Klasa"
    
    ' Otpremnice (kolone 1-4 vidljive, OtpremnicaID hidden)
    SetColumnHeader lbl_H_OTP1, "Broj otp."
    SetColumnHeader lbl_H_OTP2, "Broj zbirne"
    SetColumnHeader lbl_H_OTP3, "Koli" & ChrW(269) & "ina"
    SetColumnHeader lbl_H_OTP4, "Klasa"
    
    ' Trace (kolone 1-5 vidljive, OtkupID/OtpremnicaID hidden)
    SetColumnHeader lbl_H_TRC1, "Kooperant"
    SetColumnHeader lbl_H_TRC2, "Koli" & ChrW(269) & "ina"
    SetColumnHeader lbl_H_TRC3, "Vrsta vo" & ChrW(263) & "a"
    SetColumnHeader lbl_H_TRC4, "Stanica"
    SetColumnHeader lbl_H_TRC5, "Datum"
    
    On Error GoTo 0
End Sub

Private Sub ResetActionButtons()
    StylePrimaryButton btnAutoLink, "Automatsko povezivanje"
    StylePrimaryButton btnPovezi, Poruka("SLED_LBL_POVEZI")
    StylePrimaryButton btnStampaj, Poruka("FAK_LBL_STAMPAJ")
    StyleExitButton btnPovratak, "Povratak"
End Sub

Private Sub btnAutoLink_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
ResetActionButtons:     ButtonHover btnAutoLink
End Sub

Private Sub btnPovezi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
ResetActionButtons:     ButtonHover btnPovezi
End Sub

Private Sub btnStampaj_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
ResetActionButtons:     ButtonHover btnStampaj
End Sub

Private Sub btnPovratak_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
ResetActionButtons:     ButtonHover btnPovratak
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
End Sub

Private Sub SetColumnHeader(ByVal lbl As MSForms.label, ByVal txt As String)
    On Error Resume Next
    StyleListHeaderLabel lbl
    lbl.caption = txt
    On Error GoTo 0
End Sub

Private Sub RelayoutSledljivost()
    On Error Resume Next
    
    ' === POVEZIVANJE ZONE - eksplicitno uskladjenje ===
    lstNepovezani.top = 134
    lstNepovezani.Height = 280
    
    LstOtpremnice.top = 134
    LstOtpremnice.Height = 280
    
    btnPovezi.top = 254
    
    ' === SLEDLJIVOST ZONE - pomereno gore ===
    ' Listbox-i Povezivanje zone se zavrsavaju na 134 + 280 = 414
    ' Sledljivost section pocinje 30 pt nize = 444
    
    lblSidebarTrace.top = 444
    lblSidebarTrace.Left = 16
    lblSidebarTrace.width = 250
    lblSidebarTrace.Height = 14
    
    lblAccentTrace.top = 462
    lblAccentTrace.Left = 16
    lblAccentTrace.width = 990
    lblAccentTrace.Height = 2
    
    lblZbirnaLabel.top = 474
    lblZbirnaLabel.Left = 16
    lblZbirnaLabel.width = 60
    lblZbirnaLabel.Height = 14
    
    cmbZbirna.top = 494
    cmbZbirna.Left = 16
    cmbZbirna.width = 300
    cmbZbirna.Height = 22
    
    btnStampaj.top = 492
    btnStampaj.Left = 330
    btnStampaj.width = 100
    btnStampaj.Height = 26
    
    ' Column headers iznad lstTrace
    lbl_H_TRC1.top = 532
    lbl_H_TRC2.top = 532
    lbl_H_TRC3.top = 532
    lbl_H_TRC4.top = 532
    lbl_H_TRC5.top = 532
    
    lstTrace.top = 552
    lstTrace.Left = 16
    lstTrace.width = 990
    lstTrace.Height = 380
    
    ' btnPovratak - dno desno, dynamic od InsideHeight
    btnPovratak.top = Me.InsideHeight - 50
    btnPovratak.Left = Me.InsideWidth - btnPovratak.width - 16
    btnPovratak.width = 100
    btnPovratak.Height = 30
    
    On Error GoTo 0
End Sub

Private Sub SetupListBoxes()
    ' lstNepovezani: Datum, Stanica, Vozac, Kooperant, Kolicina, Klasa
    With lstNepovezani
        .ColumnCount = 7
        .ColumnWidths = "0;60;80;80;100;60;30"  ' OtkupID hidden
        .ListStyle = fmListStylePlain
    End With
    
    ' lstOtpremnice: OtpremnicaID, BrojOtp, BrojZbirne, Kolicina, Klasa
    With LstOtpremnice
        .ColumnCount = 5
        .ColumnWidths = "0;80;80;60;30"  ' OtpremnicaID hidden
        .ListStyle = fmListStylePlain
    End With
    
    ' lstTrace: Kooperant, Kolicina, VrstaVoca, Stanica, Datum, OtkupID, OtpremnicaID
    With lstTrace
        .ColumnCount = 7
        .ColumnWidths = "120;60;80;80;60;0;0"  ' IDs hidden
        .ListStyle = fmListStylePlain
    End With
End Sub

Private Sub LoadZbirne()
    cmbZbirna.Clear
    Dim data As Variant
    data = GetTableData(TBL_ZBIRNA)
    If IsEmpty(data) Then Exit Sub
    data = ExcludeStornirano(data, TBL_ZBIRNA)
    If IsEmpty(data) Then Exit Sub
    
    Dim colBroj As Long
    colBroj = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ)
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim br As String
        br = CStr(data(i, colBroj))
        If Not dict.Exists(br) Then
            dict.Add br, True
            cmbZbirna.AddItem br
        End If
    Next i
End Sub

Private Sub UpdateStatus()
    Dim unlinked As Variant
    unlinked = GetUnlinkedOtkupi()
    
    Dim totalOtkup As Long
    Dim data As Variant
    data = GetTableData(TBL_OTKUP)
    If Not IsEmpty(data) Then
        data = ExcludeStornirano(data, TBL_OTKUP)
        If IsArray(data) Then totalOtkup = UBound(data, 1)
    End If
    
    Dim unlinkedCount As Long
    If Not IsEmpty(unlinked) Then unlinkedCount = UBound(unlinked, 1)
    
    lblStatus.caption = "Povezano: " & (totalOtkup - unlinkedCount) & " od " & totalOtkup
End Sub

' ============================================================
' AUTO-LINK
' ============================================================

Private Sub btnAutoLink_Click()
    On Error GoTo EH
    
    Dim linked As Long
    linked = AutoLinkOtkupOtpremnica_TX()
    
    MsgBox "Automatski povezano: " & linked & " otkupa", vbInformation, APP_NAME
    
    LoadNepovezani
    UpdateStatus
    Exit Sub
EH:
    LogErr "frmSledljivost.btnAutoLink"
    MsgBox "Gre" & ChrW(353) & "ka pri povezivanju: " & Err.description, vbCritical, APP_NAME
End Sub

' ============================================================
' NEPOVEZANI (Unverknuepfte Otkupi)
' ============================================================

Private Sub LoadNepovezani()
    lstNepovezani.Clear
    LstOtpremnice.Clear
    Erase m_CandidateOtpIDs
    
    m_UnlinkedData = GetUnlinkedOtkupi()
    If IsEmpty(m_UnlinkedData) Then Exit Sub
    
    Dim i As Long
    For i = 1 To UBound(m_UnlinkedData, 1)
        ' Resolve names
        Dim stNaziv As String
        stNaziv = CStr(LookupValue(TBL_STANICE, "StanicaID", CStr(m_UnlinkedData(i, 3)), "Naziv"))
        
        Dim vozNaziv As String
        vozNaziv = CStr(LookupValue(TBL_VOZACI, "VozacID", CStr(m_UnlinkedData(i, 4)), "Ime")) & " " & _
                   CStr(LookupValue(TBL_VOZACI, "VozacID", CStr(m_UnlinkedData(i, 4)), "Prezime"))
        
        Dim koopNaziv As String
        koopNaziv = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", CStr(m_UnlinkedData(i, 5)), "Ime")) & " " & _
                    CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", CStr(m_UnlinkedData(i, 5)), "Prezime"))
        
        lstNepovezani.AddItem CStr(m_UnlinkedData(i, 1))  ' OtkupID
        lstNepovezani.List(lstNepovezani.ListCount - 1, 1) = Format$(CDate(m_UnlinkedData(i, 2)), "DD.MM.YYYY")
        lstNepovezani.List(lstNepovezani.ListCount - 1, 2) = stNaziv
        lstNepovezani.List(lstNepovezani.ListCount - 1, 3) = vozNaziv
        lstNepovezani.List(lstNepovezani.ListCount - 1, 4) = koopNaziv
        lstNepovezani.List(lstNepovezani.ListCount - 1, 5) = Format$(CDbl(m_UnlinkedData(i, 6)), "#,##0")
        lstNepovezani.List(lstNepovezani.ListCount - 1, 6) = CStr(m_UnlinkedData(i, 7))
    Next i
End Sub

Private Sub lstNepovezani_Click()
    ' Zeige moegliche Otpremnice fuer ausgewaehlten Otkup
    LstOtpremnice.Clear
    Erase m_CandidateOtpIDs
    
    If lstNepovezani.ListIndex < 0 Then Exit Sub
    
    Dim idx As Long
    idx = lstNepovezani.ListIndex + 1
    
    Dim stanicaID As String: stanicaID = CStr(m_UnlinkedData(idx, 3))
    Dim datum As Date: datum = CDate(m_UnlinkedData(idx, 2))
    
    ' Alle Otpremnice fuer diese Station + Datum
    Dim otpData As Variant
    otpData = GetTableData(TBL_OTPREMNICA)
    If IsEmpty(otpData) Then Exit Sub
    otpData = ExcludeStornirano(otpData, TBL_OTPREMNICA)
    If IsEmpty(otpData) Then Exit Sub
    
    Dim colID As Long, colSt As Long, colDat As Long
    Dim colBrOtp As Long, colBrZbr As Long, colKol As Long, colKlasa As Long
    colID = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_ID)
    colSt = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_STANICA)
    colDat = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_DATUM)
    colBrOtp = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ)
    colBrZbr = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE)
    colKol = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA)
    colKlasa = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_KLASA)
    
    Dim count As Long
    Dim i As Long
    For i = 1 To UBound(otpData, 1)
        If CStr(otpData(i, colSt)) = stanicaID And _
           IsDate(otpData(i, colDat)) Then
            If CDate(otpData(i, colDat)) = datum Then
                count = count + 1
            End If
        End If
    Next i
    
    If count = 0 Then Exit Sub
    ReDim m_CandidateOtpIDs(0 To count - 1)
    
    Dim cIdx As Long
    For i = 1 To UBound(otpData, 1)
        If CStr(otpData(i, colSt)) = stanicaID And _
           IsDate(otpData(i, colDat)) Then
            If CDate(otpData(i, colDat)) = datum Then
                m_CandidateOtpIDs(cIdx) = CStr(otpData(i, colID))
                
                LstOtpremnice.AddItem CStr(otpData(i, colID))
                LstOtpremnice.List(LstOtpremnice.ListCount - 1, 1) = CStr(otpData(i, colBrOtp))
                LstOtpremnice.List(LstOtpremnice.ListCount - 1, 2) = CStr(otpData(i, colBrZbr))
                LstOtpremnice.List(LstOtpremnice.ListCount - 1, 3) = Format$(CDbl(otpData(i, colKol)), "#,##0")
                LstOtpremnice.List(LstOtpremnice.ListCount - 1, 4) = CStr(otpData(i, colKlasa))
                
                cIdx = cIdx + 1
            End If
        End If
    Next i
End Sub

' ============================================================
' MANUELLES VERKNUePFEN
' ============================================================

Private Sub btnPovezi_Click()
    On Error GoTo EH
    
    If lstNepovezani.ListIndex < 0 Then
        MsgBox "Izaberite otkupni blok!", vbExclamation, APP_NAME
        Exit Sub
    End If
    If LstOtpremnice.ListIndex < 0 Then
        MsgBox "Izaberite otpremnicu!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Dim otkupID As String
    otkupID = lstNepovezani.List(lstNepovezani.ListIndex, 0)
    
    Dim otpremnicaID As String
    otpremnicaID = m_CandidateOtpIDs(LstOtpremnice.ListIndex)
    
    Dim rows As Collection
    Set rows = FindRows(TBL_OTKUP, COL_OTK_ID, otkupID)

    If rows.count = 0 Then
        Err.Raise vbObjectError + 1910, "frmSledljivost.btnPovezi", _
                "Otkup row nije prona" & ChrW(273) & "en: " & otkupID
    End If

RequireUpdateCell TBL_OTKUP, rows(1), COL_OTK_OTPREMNICA_ID, _
                  otpremnicaID, "frmSledljivost.btnPovezi"
    
    LoadNepovezani
    UpdateStatus
    Exit Sub
EH:
    LogErr "frmSledljivost.btnPovezi"
    MsgBox "Gre" & ChrW(353) & "ka pri povezivanju: " & Err.description, vbCritical, APP_NAME
End Sub

' ============================================================
' TRACEABILITY
' ============================================================

Private Sub cmbZbirna_Change()
    lstTrace.Clear
    If cmbZbirna.value = "" Then Exit Sub
    
    Dim traceData As Variant
    traceData = TraceByZbirna(cmbZbirna.value)
    If IsEmpty(traceData) Then Exit Sub
    
    Dim i As Long
    For i = 1 To UBound(traceData, 1)
        Dim stNaziv As String
        stNaziv = CStr(LookupValue(TBL_STANICE, "StanicaID", CStr(traceData(i, 4)), "Naziv"))
        
        lstTrace.AddItem CStr(traceData(i, 1))  ' Kooperant
        lstTrace.List(lstTrace.ListCount - 1, 1) = Format$(CDbl(traceData(i, 2)), "#,##0")
        lstTrace.List(lstTrace.ListCount - 1, 2) = CStr(traceData(i, 3))
        lstTrace.List(lstTrace.ListCount - 1, 3) = stNaziv
        lstTrace.List(lstTrace.ListCount - 1, 4) = Format$(CDate(traceData(i, 5)), "DD.MM.YYYY")
        lstTrace.List(lstTrace.ListCount - 1, 5) = CStr(traceData(i, 6))
        lstTrace.List(lstTrace.ListCount - 1, 6) = CStr(traceData(i, 7))
    Next i
End Sub

Private Sub btnStampaj_Click()
    On Error GoTo EH
    
    If cmbZbirna.value = "" Then
        MsgBox "Izaberite zbirnu!", vbExclamation, APP_NAME
        Exit Sub
    End If
    

    PrintTracePDF cmbZbirna.value
    Exit Sub
EH:
    LogErr "frmSledljivost.btnStampaj"
    MsgBox "Gre" & ChrW(353) & "ka pri " & ChrW(353) & "tampanju: " & Err.description, vbCritical, APP_NAME
End Sub

Public Sub PrintTracePDF(ByVal brojZbirne As String)
    Dim traceData As Variant
    traceData = TraceByZbirna(brojZbirne)
    If IsEmpty(traceData) Then
        MsgBox "Nema podataka za zbirnu: " & brojZbirne, vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim zbrData As Variant
    zbrData = GetTableData(TBL_ZBIRNA)
    If IsEmpty(zbrData) Then Exit Sub

    Dim colZbrBroj As Long, colZbrDatum As Long, colZbrVozac As Long
    Dim colZbrKupac As Long, colZbrVrsta As Long
    colZbrBroj = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ)
    colZbrDatum = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_DATUM)
    colZbrVozac = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_VOZAC)
    colZbrKupac = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_KUPAC)
    colZbrVrsta = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_VRSTA)

    Dim zbrRow As Long, z As Long
    For z = 1 To UBound(zbrData, 1)
        If CStr(zbrData(z, colZbrBroj)) = brojZbirne Then zbrRow = z: Exit For
    Next z
    If zbrRow = 0 Then Exit Sub

    Dim vozacID As String: vozacID = CStr(zbrData(zbrRow, colZbrVozac))
    Dim vozacNaziv As String
    vozacNaziv = CStr(LookupValue(TBL_VOZACI, "VozacID", vozacID, "Ime")) & " " & _
                 CStr(LookupValue(TBL_VOZACI, "VozacID", vozacID, "Prezime"))
    Dim kupacID As String: kupacID = CStr(zbrData(zbrRow, colZbrKupac))
    Dim kupacNaziv As String
    kupacNaziv = CStr(LookupValue(TBL_KUPCI, "KupacID", kupacID, "Naziv"))
    Dim datumOtpreme As String
    datumOtpreme = Format$(CDate(zbrData(zbrRow, colZbrDatum)), "DD.MM.YYYY")
    Dim vrsta As String: vrsta = CStr(zbrData(zbrRow, colZbrVrsta))

    Dim prijKg As Double
    Dim prijData As Variant
    prijData = GetTableData(TBL_PRIJEMNICA)
    If Not IsEmpty(prijData) Then
        prijData = ExcludeStornirano(prijData, TBL_PRIJEMNICA)
        If IsArray(prijData) Then
            Dim colPrjZbr As Long, colPrjKol As Long
            colPrjZbr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
            colPrjKol = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA)
            Dim p As Long
            For p = 1 To UBound(prijData, 1)
                If CStr(prijData(p, colPrjZbr)) = brojZbirne Then
                    If IsNumeric(prijData(p, colPrjKol)) Then prijKg = prijKg + CDbl(prijData(p, colPrjKol))
                End If
            Next p
        End If
    End If

    Dim ws As Worksheet
    Set ws = FillSledljivostSablon(brojZbirne, datumOtpreme, vozacNaziv, kupacNaziv, _
                                   vrsta, traceData, prijKg)
    If ws Is Nothing Then Exit Sub

    Dim pdfPath As String
    pdfPath = ThisWorkbook.path & "\Sledljivost_" & Replace(brojZbirne, "/", "-") & ".pdf"
    Dim mode As String
    mode = DocResolveMode(GetConfigValue(CFG_SLEDLJIVOST_PRINT_MODE), "PDF")
    Select Case mode
        Case "PRINT", "PREVIEW"
            DocPrintWs ws, mode
        Case "PDF"
            DocExportPdf ws, pdfPath, True
        ' OFF -> bez izlaza
    End Select
End Sub

Private Sub btnPovratak_Click()
    On Error GoTo EH

    ButtonActive btnPovratak

    frmOtkupAPP.ReturnToDashboard "Sekcija zatvorena."
    Unload Me

    Exit Sub

EH:
    LogErr "frmOtkup.btnPovratak_Click"
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next

    If CloseMode = vbFormControlMenu Then
        frmOtkupAPP.ReturnToDashboard "Sekcija zatvorena."
    End If

    On Error GoTo 0
End Sub

