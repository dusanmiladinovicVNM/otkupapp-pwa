VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSledljivost 
   Caption         =   "UserForm1"
   ClientHeight    =   10755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19290
   OleObjectBlob   =   "frmSledljivost.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSledljivost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================
' frmOtkupniBlokovi – Sledljivost & Povezivanje
' ============================================================
Option Explicit

Private m_UnlinkedData As Variant
Private m_CandidateOtpIDs() As String
Private mChromeRemoved As Boolean

Private Sub RemoveTitleBar()
    Dim hwnd As LongPtr
    Dim style As Long

    hwnd = FindWindow("ThunderDFrame", Me.caption)

    If hwnd <> 0 Then
        style = GetWindowLong(hwnd, GWL_STYLE)
        style = style And Not WS_CAPTION
        SetWindowLong hwnd, GWL_STYLE, style
        DrawMenuBar hwnd
    End If
End Sub

Private Sub UserForm_Activate()
    If Not mChromeRemoved Then
        Me.caption = ""
        RemoveTitleBar
        mChromeRemoved = True
    End If
End Sub

Private Sub UserForm_Initialize()
    ApplyTheme Me, BG_MAIN
    SetupListBoxes
    LoadZbirne
    UpdateStatus
End Sub

' ============================================================
' SETUP
' ============================================================

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
    MsgBox "Greska pri povezivanju: " & Err.Description, vbCritical, APP_NAME
End Sub

' ============================================================
' NEPOVEZANI (Unverknüpfte Otkupi)
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
    ' Zeige mögliche Otpremnice für ausgewählten Otkup
    LstOtpremnice.Clear
    Erase m_CandidateOtpIDs
    
    If lstNepovezani.ListIndex < 0 Then Exit Sub
    
    Dim idx As Long
    idx = lstNepovezani.ListIndex + 1
    
    Dim stanicaID As String: stanicaID = CStr(m_UnlinkedData(idx, 3))
    Dim datum As Date: datum = CDate(m_UnlinkedData(idx, 2))
    
    ' Alle Otpremnice für diese Station + Datum
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
' MANUELLES VERKNÜPFEN
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
                "Otkup row nije pronaden: " & otkupID
    End If

RequireUpdateCell TBL_OTKUP, rows(1), COL_OTK_OTPREMNICA_ID, _
                  otpremnicaID, "frmSledljivost.btnPovezi"
    
    LoadNepovezani
    UpdateStatus
    Exit Sub
EH:
    LogErr "frmSledljivost.btnPovezi"
    MsgBox "Greska pri povezivanju: " & Err.Description, vbCritical, APP_NAME
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
    MsgBox "Greska pri stampanju: " & Err.Description, vbCritical, APP_NAME
End Sub

Public Sub PrintTracePDF(ByVal brojZbirne As String)
    ' Template Sheet prüfen
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("SledljivostSablon")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "SledljivostSablon sheet ne postoji!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' Trace-Daten holen
    Dim traceData As Variant
    traceData = TraceByZbirna(brojZbirne)
    If IsEmpty(traceData) Then
        MsgBox "Nema podataka za zbirnu: " & brojZbirne, vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' Zbirna-Daten
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
    
    Dim zbrRow As Long
    Dim z As Long
    For z = 1 To UBound(zbrData, 1)
        If CStr(zbrData(z, colZbrBroj)) = brojZbirne Then
            zbrRow = z
            Exit For
        End If
    Next z
    If zbrRow = 0 Then Exit Sub
    
    ' Vozac Name
    Dim vozacID As String
    vozacID = CStr(zbrData(zbrRow, colZbrVozac))
    Dim vozacNaziv As String
    vozacNaziv = CStr(LookupValue(TBL_VOZACI, "VozacID", vozacID, "Ime")) & " " & _
                 CStr(LookupValue(TBL_VOZACI, "VozacID", vozacID, "Prezime"))
    
    ' Kupac Name
    Dim kupacID As String
    kupacID = CStr(zbrData(zbrRow, colZbrKupac))
    Dim kupacNaziv As String
    kupacNaziv = CStr(LookupValue(TBL_KUPCI, "KupacID", kupacID, "Naziv"))
    
    ' Header befüllen
    Application.ScreenUpdating = False
    
    ws.Range("LOTBroj").value = brojZbirne
    ws.Range("DatumOtpreme").value = Format$(CDate(zbrData(zbrRow, colZbrDatum)), "DD.MM.YYYY")
    ws.Range("VozacNaziv").value = vozacNaziv
    ws.Range("KupacNaziv").value = kupacNaziv
    ws.Range("VrstaVoca").value = CStr(zbrData(zbrRow, colZbrVrsta))
    
    Const NUM_COLS As Long = 12      ' <-- war 10, jetzt 12
    
    ' Alte Daten löschen
    Dim startRow As Long
    startRow = ws.Range("TraceStart").row
    Dim lastRow As Long
    lastRow = ws.cells(ws.rows.count, 1).End(xlUp).row
    If lastRow >= startRow Then
        ws.Range(ws.cells(startRow, 1), ws.cells(lastRow, NUM_COLS)).ClearContents
        ws.Range(ws.cells(startRow, 1), ws.cells(lastRow, NUM_COLS)).ClearFormats
    End If
    
    ' Text-Format für BPG + KatParcela + KatBroj(Parcela)
    ws.Range(ws.cells(startRow, 3), ws.cells(startRow + 50, 4)).NumberFormat = "@"
    
    
    
    ' Trace-Zeilen einfügen
    Dim totalOtkupKg As Double
    Dim i As Long
    For i = 1 To UBound(traceData, 1)
    
        Dim outRow As Long
        outRow = startRow + i - 1
        
        Dim stNaziv As String
        stNaziv = CStr(LookupValue(TBL_STANICE, "StanicaID", CStr(traceData(i, 4)), "Naziv"))
        
        ws.cells(outRow, 1).value = i                          ' Rb
        ws.cells(outRow, 2).value = CStr(traceData(i, 1))      ' Kooperant
        ws.cells(outRow, 3).value = CStr(traceData(i, 8))      ' BPG
        ws.cells(outRow, 4).value = CStr(traceData(i, 9))      ' KatBroj (Parcela)
        ws.cells(outRow, 5).value = CStr(traceData(i, 10))     ' GGAP (Parcela)
        ws.cells(outRow, 6).value = CStr(traceData(i, 13))     ' Kultura (Parcela)
        ws.cells(outRow, 7).value = CStr(traceData(i, 14))     ' Povrsina (Parcela)
        ws.cells(outRow, 8).value = stNaziv                     ' Stanica
        ws.cells(outRow, 9).value = Format$(CDate(traceData(i, 5)), "DD.MM.YYYY")
        ws.cells(outRow, 10).value = CDbl(traceData(i, 2))     ' Kg
        ws.cells(outRow, 11).value = CStr(traceData(i, 11))    ' Klasa
        ws.cells(outRow, 12).value = CStr(traceData(i, 6))     ' OtkupID
        
        totalOtkupKg = totalOtkupKg + CDbl(traceData(i, 2))
    Next i
    
    ' Formatierung
    Dim dataRange As Range
    Set dataRange = ws.Range(ws.cells(startRow, 1), _
                    ws.cells(startRow + UBound(traceData, 1) - 1, NUM_COLS))
    
    With dataRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    ' Kg-Spalte = Spalte 10
    ws.Range(ws.cells(startRow, 10), _
             ws.cells(startRow + UBound(traceData, 1) - 1, 10)).NumberFormat = "#,##0"
    
    ' Alternierende Farbe
    Dim r As Long
    For r = 0 To UBound(traceData, 1) - 1
        If r Mod 2 = 1 Then
            ws.Range(ws.cells(startRow + r, 1), _
                     ws.cells(startRow + r, NUM_COLS)).Interior.Color = RGB(217, 225, 242)
        End If
    Next r
    
    ' Summen
    Dim sumRow As Long
    sumRow = startRow + UBound(traceData, 1) + 1
    
    ' Prijemnica Kolicina
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
                    If IsNumeric(prijData(p, colPrjKol)) Then
                        prijKg = prijKg + CDbl(prijData(p, colPrjKol))
                    End If
                End If
            Next p
        End If
    End If
    
    Dim manjak As Double
    manjak = totalOtkupKg - prijKg
    Dim manjakPct As Double
    If totalOtkupKg > 0 Then manjakPct = manjak / totalOtkupKg * 100
    
    ws.cells(sumRow, 1).value = "Ukupno otkup:"
    ws.cells(sumRow, 10).value = totalOtkupKg           ' <-- war 8, jetzt 10
    ws.cells(sumRow + 1, 1).value = "Ukupno prijemnica:"
    ws.cells(sumRow + 1, 10).value = prijKg              ' <-- war 8
    ws.cells(sumRow + 2, 1).value = "Manjak:"
    ws.cells(sumRow + 2, 10).value = manjak              ' <-- war 8
    ws.cells(sumRow + 2, 11).value = Format$(manjakPct, "0.00") & "%"  ' <-- war 9
    
    ws.cells(sumRow + 4, 1).value = "Datum stampe: " & Format$(Date, "DD.MM.YYYY")
    ws.cells(sumRow + 5, 1).value = "Potpis: ___________"
    ws.cells(sumRow + 5, 8).value = "Pecat: ___________"  ' <-- war 6
    
    ' PDF Export
    Dim pdfPath As String
    pdfPath = ThisWorkbook.Path & "\Sledljivost_" & Replace(brojZbirne, "/", "-") & ".pdf"
    
    ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, _
                           Quality:=xlQualityStandard, _
                           IncludeDocProperties:=False, _
                           OpenAfterPublish:=True
    
    Application.ScreenUpdating = True
End Sub

Private Sub btnPovratak_Click()
    Me.Hide
    frmOtkupAPP.Show
End Sub


