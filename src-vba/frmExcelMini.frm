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
    Me.BackColor = BG_MAIN()
    If Not mChromeRemoved Then
        Me.Caption = ""
        RemoveTitleBar
        mChromeRemoved = True
    End If
    StylePrimaryButton btnCloseExcel, "Zatvori Excel"
End Sub
Private Sub UserForm_Initialize()
    btnCloseExcel.Caption = "Zatvori Excel"
    Me.StartUpPosition = 0
    Me.Left = Application.Left + Application.Width - Me.Width - 20
    Me.Top = Application.Top + 40
End Sub

Private Sub btnCloseExcel_Click()
    Application.Visible = False
    Unload Me
    frmOtkupAPP.Show
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Application.Visible = False
    frmOtkupAPP.Show
End Sub
