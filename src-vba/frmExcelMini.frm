VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExcelMini 
   Caption         =   "UserForm1"
   ClientHeight    =   1020
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3000
   OleObjectBlob   =   "frmExcelMini.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExcelMini"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' ============================================================
' frmExcelMini / plutajuca kartica dok je Excel otvoren
' Responsibility:
'   - jedno dugme: sakrij Excel i vrati ljusku (modOtkupUI.ShowOtkupUI)
'   - no business logic
'
' Otvara je modOtkupUI.DoShowExcel ("Otvori Excel"). Izgled prati ljusku:
' krem kartica sa 1pt ivicom, forest traka levo, "AX OtkupApp" u display
' fontu, zeleno primarno dugme. Kontrola iz dizajnera (btnCloseExcel) se samo
' stilizuje (modUiKit.PanelStilDugme -- isti primitiv koji oblaci dugmad
' panela); ostalo je runtime (modUiKit.NewLbl), pa se .frx ne dira. Nema
' module-level MSForms deklaracija (meka forma).
' ============================================================

Private Const MINI_W As Single = 232
Private Const MINI_H As Single = 78

Private mChromeRemoved As Boolean
Private m_IsClosing As Boolean

Private Sub UserForm_Initialize()
    On Error GoTo EH

    mChromeRemoved = False
    m_IsClosing = False

    BuildMini

    ' gore desno u Excelu, kao i do sada
    Me.StartUpPosition = 0
    If Application.Visible Then
        Me.Left = Application.Left + Application.width - Me.width - 20
        Me.top = Application.top + 40
    End If

    Exit Sub

EH:
    LogErr "frmExcelMini.UserForm_Initialize"
End Sub

Private Sub BuildMini()
    Dim fnt As String
    fnt = DisplayFont()

    Me.width = MINI_W
    Me.Height = MINI_H
    Me.BackColor = C_CREAM

    ' ivica + ispuna (isti par kao NewShell) iza svega; ispuna pa ivica, da
    ' ivica zavrsi najdublje
    NewLbl Me, "mnB", "", 0, 0, MINI_W, MINI_H, 8, False, 0, C_BORDER
    NewLbl Me, "mnF", "", 1, 1, MINI_W - 2, MINI_H - 2, 8, False, 0, C_CREAM
    Me.Controls("mnF").ZOrder 1
    Me.Controls("mnB").ZOrder 1
    NewLbl Me, "mnBar", "", 1, 1, 5, MINI_H - 2, 8, False, 0, C_FOREST

    NewLbl Me, "mnAX", "AX", 16, 9, 22, TxtH(TS_H1), TS_H1, True, C_GOLD, -1, fmTextAlignLeft, fnt
    NewLbl Me, "mnName", "OtkupApp", 38, 9, 90, TxtH(TS_H1), TS_H1, True, C_FOREST, -1, fmTextAlignLeft, fnt
    NewLbl Me, "mnSub", Poruka("OTKUI_MINI_EXCEL"), 120, 11, MINI_W - 132, TxtH(TS_META), _
           TS_META, False, C_MUTED, -1, fmTextAlignRight

    With btnCloseExcel
        .caption = Poruka("OTKUI_MINI_NAZAD")
        .Left = 16
        .top = 38
        .width = MINI_W - 32
        .Height = 28
        .ZOrder 0
    End With
    PanelStilDugme btnCloseExcel, "primary"
End Sub

Private Sub UserForm_Activate()
    On Error GoTo EH
    EnsureUserFormChromeRemoved Me, mChromeRemoved
    Exit Sub
EH:
    LogErr "frmExcelMini.UserForm_Activate"
End Sub

Private Sub btnCloseExcel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    PanelStilDugmeHover btnCloseExcel, "primary", True
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    PanelStilDugmeHover btnCloseExcel, "primary", False
End Sub

Private Sub btnCloseExcel_Click()
    ReturnToAppShell
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        ReturnToAppShell
    End If
End Sub

Private Sub ReturnToAppShell()
    On Error GoTo EH

    If m_IsClosing Then Exit Sub
    m_IsClosing = True

    Application.Visible = False

    On Error Resume Next
    modOtkupUI.ShowOtkupUI
    On Error GoTo EH

    Unload Me
    Exit Sub

EH:
    LogErr "frmExcelMini.ReturnToAppShell"

    On Error Resume Next
    Application.Visible = False
    modOtkupUI.ShowOtkupUI
    Unload Me
End Sub
