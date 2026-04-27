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
' frmCloseExcel / Excel close helper
' Responsibility:
'   - hide Excel
'   - return operator to frmOtkupAPP
'   - no business logic
' ============================================================

Private mChromeRemoved As Boolean
Private m_IsClosing As Boolean

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

Private Sub UserForm_Initialize()
    On Error GoTo EH

    mChromeRemoved = False
    m_IsClosing = False

    btnCloseExcel.caption = "Zatvori Excel"

    Me.StartUpPosition = 0

    If Application.Visible Then
        Me.Left = Application.Left + Application.Width - Me.Width - 20
        Me.Top = Application.Top + 40
    End If

    Me.BackColor = BG_MAIN()
    StylePrimaryButton btnCloseExcel, "Zatvori Excel"

    Exit Sub

EH:
    LogErr "frmCloseExcel.UserForm_Initialize"
End Sub

Private Sub UserForm_Activate()
    On Error GoTo EH

    Me.BackColor = BG_MAIN()

    If Not mChromeRemoved Then
        Me.caption = ""
        RemoveTitleBar
        mChromeRemoved = True
    End If

    Exit Sub

EH:
    LogErr "frmCloseExcel.UserForm_Activate"
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
    frmOtkupAPP.Show
    On Error GoTo EH

    Unload Me
    Exit Sub

EH:
    LogErr "frmCloseExcel.ReturnToAppShell"

    On Error Resume Next
    Application.Visible = False
    frmOtkupAPP.Show
    Unload Me
End Sub
