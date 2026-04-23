VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMarza 
   Caption         =   "UserForm1"
   ClientHeight    =   9285.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17070
   OleObjectBlob   =   "frmMarza.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMarza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmMarza v2.1 – Margenberechnung
' GEÄNDERT: tblIsporuka ? tblPrijemnica
' Verkaufsseite = Prijemnica.Kolicina × Prijemnica.Cena
' VrstaVoca-Lookup: Prijemnica.BrojZbirne ? Otpremnica.VrstaVoca
' ============================================================

Private m_SetupDone As Boolean
Private m_LastMarzaData As Variant

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
    
    If m_SetupDone Then Exit Sub
    m_SetupDone = True
    
    Me.Caption = "Marza"
    txtDatumOd.Value = "1.1." & Year(Date)
    txtDatumDo.Value = Format$(Date, "d.m.yyyy")
    
    cmbTipMarze.Clear
    cmbTipMarze.AddItem "Po Kupcu"
    cmbTipMarze.AddItem "Po Otkupnom mestu"
    cmbTipMarze.AddItem "Ukupno"
    cmbTipMarze.ListIndex = 0
    
    LoadEntiteti
    
    With lstMarza
        .ColumnCount = 8
        .ColumnWidths = "80;70;70;80;70;80;80;55"
    End With
End Sub

Private Sub cmbTipMarze_Change()
    LoadEntiteti
End Sub

Private Sub LoadEntiteti()
    cmbEntitet.Clear
    cmbEntitet.Enabled = True
    
    Select Case cmbTipMarze.Value
        Case "Po Kupcu"
            Dim kupci As Variant
            kupci = GetLookupList(TBL_KUPCI, "Naziv")
            If IsArray(kupci) Then
                Dim i As Long
                For i = LBound(kupci) To UBound(kupci)
                    cmbEntitet.AddItem CStr(kupci(i))
                Next i
            End If
        Case "Po Otkupnom mestu"
            Dim stanice As Variant
            stanice = GetLookupList(TBL_STANICE, "Naziv")
            If IsArray(stanice) Then
                For i = LBound(stanice) To UBound(stanice)
                    cmbEntitet.AddItem CStr(stanice(i))
                Next i
            End If
        Case "Ukupno"
            cmbEntitet.Enabled = False
    End Select
    
    If cmbEntitet.ListCount > 0 Then cmbEntitet.ListIndex = 0
End Sub

Private Sub btnPrikazi_Click()
    On Error GoTo EH
    Dim datumOd As Date, datumDo As Date
    datumOd = CDate(txtDatumOd.Value)
    datumDo = CDate(txtDatumDo.Value)
    
    lstMarza.Clear
    
    Dim data As Variant
    Select Case cmbTipMarze.Value
        Case "Po Kupcu"
            Dim kupacID As String
            kupacID = CStr(LookupValue(TBL_KUPCI, "Naziv", cmbEntitet.Value, "KupacID"))
            data = ReportMarzaByKupac(kupacID, datumOd, datumDo)
        Case "Po Otkupnom mestu"
            Dim stanicaID As String
            stanicaID = CStr(LookupValue(TBL_STANICE, "Naziv", cmbEntitet.Value, "StanicaID"))
            data = ReportMarzaByOM(stanicaID, datumOd, datumDo)
        Case "Ukupno"
            data = ReportMarzaUkupno(datumOd, datumDo)
    End Select
    
    If IsEmpty(data) Then
        MsgBox "Nema podataka!", vbInformation, APP_NAME
        Exit Sub
    End If
    
    m_LastMarzaData = data
    FillMarzaList data
    WriteMarza datumOd, datumDo
    Exit Sub
EH:
    LogErr "frmMarza.btnPrikazi"
    MsgBox "Greska pri ucitavanju marze: " & Err.Description, vbCritical, APP_NAME
End Sub

Private Sub FillMarzaList(ByVal data As Variant)
    Dim i As Long, j As Long
    For i = 1 To UBound(data, 1)
        lstMarza.AddItem CStr(data(i, 1))
        For j = 2 To UBound(data, 2)
            If IsNumeric(data(i, j)) And data(i, j) <> "" Then
                If j = 8 Then
                    lstMarza.List(lstMarza.ListCount - 1, j - 1) = Format$(CDbl(data(i, j)), "#,##0.0") & "%"
                Else
                    lstMarza.List(lstMarza.ListCount - 1, j - 1) = Format$(CDbl(data(i, j)), "#,##0.00")
                End If
            End If
        Next j
    Next i
End Sub

Private Sub WriteMarza(ByVal datumOd As Date, ByVal datumDo As Date)
    If IsEmpty(m_LastMarzaData) Then Exit Sub
    
    Dim lo As ListObject
    Set lo = GetTable(TBL_RPT_MARZA)
    If lo Is Nothing Then Exit Sub
    
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    
    Dim i As Long
    For i = 1 To UBound(m_LastMarzaData, 1)
        Dim rowData As Variant
        rowData = Array(Format$(Date, "yyyy-mm-dd"), _
                       m_LastMarzaData(i, 1), _
                       m_LastMarzaData(i, 2), _
                       m_LastMarzaData(i, 3), _
                       m_LastMarzaData(i, 4), _
                       m_LastMarzaData(i, 5), _
                       m_LastMarzaData(i, 6), _
                       m_LastMarzaData(i, 7))
        AppendRow TBL_RPT_MARZA, rowData
    Next i
End Sub
Private Sub btnPovratak_Click()
    Me.Hide
    frmOtkupAPP.Show
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Unload Me
        frmMain.Show
    End If
End Sub

