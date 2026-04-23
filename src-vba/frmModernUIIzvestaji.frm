VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmModernUIIzvestaji 
   Caption         =   "UserForm1"
   ClientHeight    =   11820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20565
   OleObjectBlob   =   "frmModernUIIzvestaji.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmModernUIIzvestaji"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' VBA UserForm: frmIzvestaji
' Version without ListView
' Uses only standard MSForms controls so it works on more machines.
'
' Create a UserForm named: frmIzvestaji
' Suggested size:
'   Width = 1040
'   Height = 620
'
' Place these controls on the form:
' -------------------------------------------------
' Labels:
'   lblTitle
'   lblClose
'   lblMenuOtkupljenaMesta
'   lblMenuKupci
'   lblMenuKooperanti
'   lblMenuVozaci
'   lblMenuPojedinoci
'   lblFooterIcon1
'   lblFooterIcon2
'   lblFooterIcon3
'   lblFilterCaption
'   lblDash
'
' MultiPage:
'   mpTabs   (5 pages)
'   Page captions:
'     0 = Sažetak
'     1 = Otkupljeno voca
'     2 = Primljena ambalaža
'     3 = Izpisani
'     4 = Prosecna cena
'
' Frame:
'   fraFilters
'
' ComboBoxes:
'   cboRazina
'
' TextBoxes:
'   txtDateFrom
'   txtDateTo
'
' CommandButtons:
'   btnPrikazi
'   btnStampaj
'   btnPovratak
'
' ListBox on page 0:
'   lstReport
'
' =============================================================
' UserForm code
' =============================================================
Option Explicit

Private Const BG_FORM As Long = &H1A2233
Private Const BG_PANEL As Long = &H202B40
Private Const BG_ACTIVE As Long = &H8A4E1A
Private Const BG_BUTTON As Long = &HCC6A1E
Private Const BG_BUTTON_ALT As Long = &H34445F
Private Const FG_TEXT As Long = &HF4F7FB
Private Const FG_MUTED As Long = &HC9D2E3

Private Sub UserForm_Initialize()
    BuildShell
    BuildSidebar
    BuildTabs
    BuildFilters
    BuildButtons
    BuildReportList
    LoadSampleData
End Sub

Private Sub BuildShell()
    With Me
        .Caption = vbNullString
        .BackColor = BG_FORM
        .Width = 1040
        .Height = 620
        .StartUpPosition = 1
        .BorderStyle = 0
    End With

    With lblTitle
        .Caption = "Izveštaji"
        .Left = 24
        .Top = 18
        .Width = 240
        .Height = 28
        .Font.Name = "Segoe UI Semibold"
        .Font.Size = 18
        .foreColor = FG_TEXT
        .BackStyle = fmBackStyleTransparent
    End With

    With lblClose
        .Caption = "×"
        .Left = Me.Width - 52
        .Top = 14
        .Width = 24
        .Height = 24
        .TextAlign = fmTextAlignCenter
        .Font.Name = "Segoe UI"
        .Font.Size = 18
        .foreColor = FG_MUTED
        .BackStyle = fmBackStyleTransparent
        '.MousePointer = fmMousePointerHand
    End With
End Sub

Private Sub BuildSidebar()
    StyleMenuLabel lblMenuOtkupljenaMesta, "Otkupljena mesta", 18, 78, True
    StyleMenuLabel lblMenuKupci, "Kupci", 18, 126, False
    StyleMenuLabel lblMenuKooperanti, "Kooperanti", 18, 174, False
    StyleMenuLabel lblMenuVozaci, "Vozaci", 18, 222, False
    StyleMenuLabel lblMenuPojedinoci, "Pojedinci", 18, 270, False

    StyleFooterIcon lblFooterIcon1, "?", 36, Me.Height - 78
    StyleFooterIcon lblFooterIcon2, "=", 86, Me.Height - 78
    StyleFooterIcon lblFooterIcon3, "?", 136, Me.Height - 78
End Sub

Private Sub BuildTabs()
    With mpTabs
        .Left = 250
        .Top = 72
        .Width = 750
        .Height = 390
        .style = fmTabStyleTabs
        .Font.Name = "Segoe UI Semibold"
        .Font.Size = 10
    End With

    Dim i As Long
End Sub

Private Sub BuildFilters()
    With fraFilters
        .Caption = vbNullString
        .Left = 250
        .Top = 470
        .Width = 360
        .Height = 58
        .BackColor = BG_FORM
        .BorderStyle = fmBorderStyleNone
    End With

    With lblFilterCaption
        .Caption = "Razina:"
        .Left = 252
        .Top = 487
        .Width = 56
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .foreColor = FG_TEXT
        .BackStyle = fmBackStyleTransparent
    End With

    With cboRazina
        .Left = 320
        .Top = 482
        .Width = 120
        .Height = 28
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .BackColor = BG_PANEL
        .foreColor = FG_TEXT
        .AddItem "Dnevno"
        .AddItem "Nedeljno"
        .AddItem "Mesecno"
        .AddItem "Sezonski"
        .ListIndex = 3
    End With

    With txtDateFrom
        .Left = 452
        .Top = 482
        .Width = 108
        .Height = 28
        .Text = Format(Date, "d.m.yyyy")
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .BackColor = BG_PANEL
        .foreColor = FG_TEXT
        .BorderStyle = fmBorderStyleSingle
    End With

    With lblDash
        .Caption = "-"
        .Left = 568
        .Top = 487
        .Width = 12
        .Height = 18
        .Font.Name = "Segoe UI"
        .Font.Size = 12
        .foreColor = FG_MUTED
        .BackStyle = fmBackStyleTransparent
    End With

    With txtDateTo
        .Left = 586
        .Top = 482
        .Width = 108
        .Height = 28
        .Text = Format(Date, "d.m.yyyy")
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .BackColor = BG_PANEL
        .foreColor = FG_TEXT
        .BorderStyle = fmBorderStyleSingle
    End With
End Sub

Private Sub BuildButtons()
    StyleActionButton btnPovratak, "Povratak", 604, 532, 126, 38, False
    StyleActionButton btnPrikazi, "Prikaži", 740, 532, 126, 38, True
    StyleActionButton btnStampaj, "Štampaj", 876, 532, 126, 38, False
End Sub

Private Sub BuildReportList()
    With lstReport
        .Left = 14
        .Top = 14
        .Width = 716
        .Height = 318
        .Font.Name = "Consolas"
        .Font.Size = 10
        .BackColor = BG_PANEL
        .foreColor = FG_TEXT
        .BorderStyle = fmBorderStyleSingle
        .ColumnCount = 1
        .IntegralHeight = False
    End With
End Sub

Private Sub LoadSampleData()
    Dim s As String

    lstReport.Clear

    s = PadR("Kupac", 18) & PadR("Kooperant", 16) & PadR("Vozac", 14) & _
        PadR("Kolicina", 12) & PadR("Vrednost", 16) & PadR("Ambalaža", 18) & _
        PadR("Zaduženo", 14) & PadR("Saldo", 14)
    lstReport.AddItem s
    lstReport.AddItem String(130, "-")

    s = PadR("Ekoplant D.O.O.", 18) & PadR("Korenski", 16) & PadR("Jovanovic", 14) & _
        PadR("82,560 kg", 12) & PadR("13,568,000", 16) & PadR("Plasticne gajbe", 18) & _
        PadR("58,600", 14) & PadR("6,184,620", 14)
    lstReport.AddItem s

    s = PadR("Agrofarm", 18) & PadR("Mitrovic", 16) & PadR("Petrovic", 14) & _
        PadR("65,300 kg", 12) & PadR("9,825,000", 16) & PadR("Drvene kace", 18) & _
        PadR("79,800", 14) & PadR("1,745,200", 14)
    lstReport.AddItem s

    s = PadR("Vocarstvo", 18) & PadR("Nikolic", 16) & PadR("Stanic", 14) & _
        PadR("74,180 kg", 12) & PadR("12,460,000", 16) & PadR("Metalne bacve", 18) & _
        PadR("162,400", 14) & PadR("10,580,000", 14)
    lstReport.AddItem s
End Sub

Private Function PadR(ByVal txt As String, ByVal totalLen As Long) As String
    If Len(txt) >= totalLen Then
        PadR = Left$(txt, totalLen)
    Else
        PadR = txt & Space$(totalLen - Len(txt))
    End If
End Function

Private Sub StyleMenuLabel(lbl As MSForms.Label, txt As String, X As Double, Y As Double, isActive As Boolean)
    With lbl
        .Caption = "   " & txt
        .Left = X
        .Top = Y
        .Width = 170
        .Height = 40
        .Font.Name = "Segoe UI Semibold"
        .Font.Size = 12
        .foreColor = IIf(isActive, FG_TEXT, FG_MUTED)
        .BackStyle = fmBackStyleOpaque
        .BackColor = IIf(isActive, BG_ACTIVE, BG_PANEL)
        .BorderStyle = fmBorderStyleSingle
        .TextAlign = fmTextAlignLeft
        '.MousePointer = fmMousePointerHand
    End With
End Sub

Private Sub StyleFooterIcon(lbl As MSForms.Label, txt As String, X As Double, Y As Double)
    With lbl
        .Caption = txt
        .Left = X
        .Top = Y
        .Width = 32
        .Height = 32
        .Font.Name = "Segoe UI Symbol"
        .Font.Size = 16
        .foreColor = FG_MUTED
        .BackStyle = fmBackStyleTransparent
        .TextAlign = fmTextAlignCenter
        '.MousePointer = fmMousePointerHand
    End With
End Sub

Private Sub StyleActionButton(btn As MSForms.CommandButton, txt As String, X As Double, Y As Double, w As Double, h As Double, isPrimary As Boolean)
    With btn
        .Caption = txt
        .Left = X
        .Top = Y
        .Width = w
        .Height = h
        .Font.Name = "Segoe UI Semibold"
        .Font.Size = 11
        .BackColor = IIf(isPrimary, BG_BUTTON, BG_BUTTON_ALT)
        .foreColor = FG_TEXT
        .TakeFocusOnClick = False
    End With
End Sub

Private Sub ActivateMenu(ByVal menuName As String)
    StyleMenuLabel lblMenuOtkupljenaMesta, "Otkupljena mesta", 18, 78, (menuName = "mesta")
    StyleMenuLabel lblMenuKupci, "Kupci", 18, 126, (menuName = "kupci")
    StyleMenuLabel lblMenuKooperanti, "Kooperanti", 18, 174, (menuName = "kooperanti")
    StyleMenuLabel lblMenuVozaci, "Vozaci", 18, 222, (menuName = "vozaci")
    StyleMenuLabel lblMenuPojedinoci, "Pojedinci", 18, 270, (menuName = "pojedinci")
End Sub

Private Sub lblMenuOtkupljenaMesta_Click()
    ActivateMenu "mesta"
End Sub

Private Sub lblMenuKupci_Click()
    ActivateMenu "kupci"
End Sub

Private Sub lblMenuKooperanti_Click()
    ActivateMenu "kooperanti"
End Sub

Private Sub lblMenuVozaci_Click()
    ActivateMenu "vozaci"
End Sub

Private Sub lblMenuPojedinoci_Click()
    ActivateMenu "pojedinci"
End Sub

Private Sub btnPrikazi_Click()
    LoadSampleData
End Sub

Private Sub btnStampaj_Click()
    MsgBox "Ovde dodaj logiku za štampu izveštaja.", vbInformation
End Sub

Private Sub btnPovratak_Click()
    Unload Me
End Sub

Private Sub lblClose_Click()
    Unload Me
End Sub

' =============================================================
' Standard module: modShowReportUI
' =============================================================
' Option Explicit
'
' Public Sub ShowModernReportUI()
'     frmIzvestaji.Show
' End Sub


