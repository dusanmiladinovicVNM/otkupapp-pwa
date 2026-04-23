VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSplash 
   Caption         =   "UserForm1"
   ClientHeight    =   3525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmSplash.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSplash"
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


Private Sub UserForm_Initialize()
    lblApp.Caption = "OtkupApp"
    lblVersion.Caption = "v2.1.0"
    lblBy.Caption = "Powered by AgriX"
End Sub

Private Sub UserForm_Activate()
    Me.BackColor = BG_MAIN()
    If Not mChromeRemoved Then
        Me.Caption = ""
        RemoveTitleBar
        mChromeRemoved = True
    End If
    StyleLabel lblApp, TXT_MUTED(), True
    StyleLabel lblVersion, TXT_MUTED(), True
    StyleLabel lblBy, TXT_MUTED(), True
    Dim t As Single
    t = Timer

    Do While Timer < t + 2
        DoEvents
    Loop

    Unload Me
    frmOtkupAPP.Show
End Sub
