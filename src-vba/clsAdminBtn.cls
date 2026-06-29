VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAdminBtn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' ============================================================
' clsAdminBtn -- Event-omotac za dinamicki kreirano dugme u Admin panelu
' (vidi modAdmin).
'
' Dugmad se prave u kodu (Controls.Add) da se NE bi dirao frmStammdaten.frx.
' Klik se hvata samo preko WithEvents instance u ovoj klasi, koja prosledjuje
' akciju nazad u modAdmin.AdminPanel_OnClick.
'
' Isti obrazac kao clsConfigBtn / clsLookupMenuBtn / clsBlokUI.
' ============================================================

Public action As String            ' kljuc akcije (vidi modAdmin.AdminPanel_OnClick)
Public WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1

Private Sub btn_Click()
    AdminPanel_OnClick action
End Sub
