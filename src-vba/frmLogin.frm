VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin
   Caption         =   "Prijava"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3960
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmLogin – prijava korisnika (Faza 1). UI only; validacija je u modAuth.
'
' VAZNO (CLAUDE.md §4): kontrole se grade U RUNTIME-u (Controls.Add +
' form-level WithEvents), pa .frx NIJE potreban. Ako VBA ne uvuce ovaj .frm,
' napravi praznu UserForm-u imena "frmLogin" (Insert > UserForm) i nalepi ovaj
' kod — rezultat je identican.
' ============================================================

Public LoginOK As Boolean

Private mAttempts As Long
Private WithEvents btnOK As MSForms.CommandButton
Attribute btnOK.VB_VarHelpID = -1
Private WithEvents btnCancel As MSForms.CommandButton
Attribute btnCancel.VB_VarHelpID = -1
Private txtUser As MSForms.TextBox
Private txtPin As MSForms.TextBox
Private lblErr As MSForms.Label

Private Sub UserForm_Initialize()
    On Error GoTo EH

    Me.Caption = APP_NAME & " — Prijava"
    Me.Width = 240
    Me.Height = 152

    AddLabel "Korisničko ime:", 10, 12, 95
    Set txtUser = Me.Controls.Add("Forms.TextBox.1", "txtUser", True)
    txtUser.Left = 110: txtUser.Top = 10: txtUser.Width = 112: txtUser.Height = 18

    AddLabel "PIN:", 10, 42, 95
    Set txtPin = Me.Controls.Add("Forms.TextBox.1", "txtPin", True)
    txtPin.Left = 110: txtPin.Top = 40: txtPin.Width = 112: txtPin.Height = 18
    txtPin.PasswordChar = "*"

    Set lblErr = Me.Controls.Add("Forms.Label.1", "lblErr", True)
    lblErr.Left = 10: lblErr.Top = 66: lblErr.Width = 212: lblErr.Height = 16
    lblErr.ForeColor = RGB(200, 0, 0)
    lblErr.Caption = vbNullString

    Set btnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel", True)
    btnCancel.Caption = "Otkaži": btnCancel.Left = 12: btnCancel.Top = 92
    btnCancel.Width = 100: btnCancel.Height = 24: btnCancel.Cancel = True

    Set btnOK = Me.Controls.Add("Forms.CommandButton.1", "btnOK", True)
    btnOK.Caption = "Prijava": btnOK.Left = 122: btnOK.Top = 92
    btnOK.Width = 100: btnOK.Height = 24: btnOK.Default = True

    mAttempts = 0
    Me.LoginOK = False
    Exit Sub
EH:
    LogErr "frmLogin.UserForm_Initialize"
End Sub

Private Sub AddLabel(ByVal cap As String, ByVal l As Single, ByVal t As Single, ByVal w As Single)
    Dim lb As MSForms.Label
    Set lb = Me.Controls.Add("Forms.Label.1", , True)
    lb.Caption = cap: lb.Left = l: lb.Top = t: lb.Width = w: lb.Height = 16
End Sub

Private Sub btnOK_Click()
    On Error GoTo EH

    lblErr.Caption = vbNullString

    If modAuth.ValidateLogin(txtUser.Text, txtPin.Text) Then
        Me.LoginOK = True
        Me.Hide
        Exit Sub
    End If

    mAttempts = mAttempts + 1
    txtPin.Text = vbNullString

    If mAttempts >= 3 Then
        Me.LoginOK = False
        Me.Hide
        Exit Sub
    End If

    lblErr.Caption = "Pogrešno korisničko ime ili PIN (" & mAttempts & "/3)."
    On Error Resume Next
    txtPin.SetFocus
    Exit Sub
EH:
    LogErr "frmLogin.btnOK_Click"
End Sub

Private Sub btnCancel_Click()
    Me.LoginOK = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Zatvaranje preko "X" = otkaz (ne sme da prodje kao uspeh).
    If CloseMode = vbFormControlMenu Then Me.LoginOK = False
End Sub
