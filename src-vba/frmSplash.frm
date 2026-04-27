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
Option Explicit

' ============================================================
' frmSplash / startup splash
' Responsibility:
'   - show branding briefly
'   - then open frmOtkupAPP
'   - no business logic
' ============================================================

Private mChromeRemoved As Boolean
Private m_Started As Boolean
Private m_IsNavigating As Boolean

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
    m_Started = False
    m_IsNavigating = False

    lblApp.caption = "OtkupApp"
    lblVersion.caption = "v" & APP_VERSION
    lblBy.caption = "Powered by AgriX"

    Me.BackColor = BG_MAIN()

    StyleLabel lblApp, TXT_MUTED(), True
    StyleLabel lblVersion, TXT_MUTED(), True
    StyleLabel lblBy, TXT_MUTED(), True

    Exit Sub

EH:
    LogErr "frmSplash.UserForm_Initialize"
End Sub

Private Sub UserForm_Activate()
    On Error GoTo EH

    Me.BackColor = BG_MAIN()

    If Not mChromeRemoved Then
        Me.caption = ""
        RemoveTitleBar
        mChromeRemoved = True
    End If

    If m_Started Then Exit Sub
    m_Started = True

    WaitSeconds 2
    OpenAppShell

    Exit Sub

EH:
    LogErr "frmSplash.UserForm_Activate"
    OpenAppShell
End Sub

Private Sub WaitSeconds(ByVal secondsToWait As Double)
    On Error GoTo EH

    Dim endTime As Date
    endTime = DateAdd("s", secondsToWait, Now)

    Do While Now < endTime
        DoEvents
    Loop

    Exit Sub

EH:
    LogErr "frmSplash.WaitSeconds"
End Sub

Private Sub OpenAppShell()
    On Error GoTo EH

    If m_IsNavigating Then Exit Sub
    m_IsNavigating = True

    Unload Me
    frmOtkupAPP.Show

    Exit Sub

EH:
    LogErr "frmSplash.OpenAppShell"

    On Error Resume Next
    frmOtkupAPP.Show
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        OpenAppShell
    End If
End Sub

