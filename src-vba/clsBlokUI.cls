VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBlokUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' clsBlokUI – Event-omotac za dinamicki kreirane kontrole
' panela "Otkupni blokovi" (vidi modOtkupBlok).
'
' Kontrole se prave u kodu (Controls.Add), pa im se eventi
' hvataju samo preko WithEvents instance u klasi. Svaka instanca
' obmota JEDNU kontrolu i prosledjuje event nazad u modOtkupBlok,
' uz "Action" tag koji kaze sta event znaci.
' ============================================================

Public action As String

Public WithEvents btn As MSForms.CommandButton
Public WithEvents txt As MSForms.TextBox
Public WithEvents lst As MSForms.ListBox
Public WithEvents cmb As MSForms.ComboBox

Private Sub btn_Click()
    OtkupBlok_OnButton action
End Sub

Private Sub cmb_Change()
    OtkupBlok_OnComboChange action
End Sub

Private Sub txt_Change()
    OtkupBlok_OnText action
End Sub

Private Sub txt_AfterUpdate()
    OtkupBlok_OnTextAfter action
End Sub

Private Sub lst_Click()
    OtkupBlok_OnListClick action
End Sub
