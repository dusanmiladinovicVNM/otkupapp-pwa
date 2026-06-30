VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "Prijava"
   ClientHeight    =   4215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "frmLogin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginOK As Boolean
Private mAttempts As Long

Private Sub UserForm_Initialize()
    On Error Resume Next
    Me.caption = APP_NAME & " - Prijava"

    ApplyTheme Me, BG_MAIN()            ' pozadina kao u aplikaciji
    ApplyThemeToControls Me

    StyleFrameTitleLabel lblTitle, "PRIJAVA"
    StyleSubtitle lblSubtitle, "Pristup aplikaciji"

    lblUser.caption = "Korisni" & ChrW(269) & "ko ime:"
    lblPin.caption = "PIN:"
    lblErr.caption = ""
    lblErr.ForeColor = RGB(200, 0, 0)

    txtPin.PasswordChar = ChrW(8226)   ' bullet -> maskiran PIN

    StylePrimaryButton btnOK, "Prijava"
    StyleExitButton btnCancel, "Otka" & ChrW(382) & "i"

    Me.LoginOK = False
    mAttempts = 0
End Sub

Private Sub btnOK_Click()
    On Error GoTo EH
    lblErr.caption = ""

    If modAuth.ValidateLogin(txtUser.text, txtPin.text) Then
        Me.LoginOK = True
        Me.Hide
        Exit Sub
    End If

    mAttempts = mAttempts + 1
    txtPin.text = ""
    If mAttempts >= 3 Then
        Me.LoginOK = False
        Me.Hide
        Exit Sub
    End If

    lblErr.caption = "Pogre" & ChrW(353) & "no korisni" & ChrW(269) & "ko ime ili PIN (" & mAttempts & "/3)."
    txtUser.SetFocus
    Exit Sub
EH:
    LogErr "frmLogin.btnOK_Click"
End Sub

Private Sub btnCancel_Click()
    Me.LoginOK = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then Me.LoginOK = False
End Sub
