VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConfigBtn1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' ============================================================
' clsConfigBtn — Event-omotac za dinamicki kreirano dugme u editoru
' "Podesavanja" (vidi modPodesavanja).
'
' Dugmad se prave u kodu (Controls.Add) da se NE bi dirao
' frmStammdaten.frx. Klik se hvata samo preko WithEvents instance u ovoj
' klasi, koja prosledjuje akciju nazad u modPodesavanja.
'
' Isti obrazac kao clsLookupMenuBtn / clsBlokUI.
' ============================================================

Public action As String            ' "save" | "toggle" | "back"
Public WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1

Private Sub btn_Click()
    ConfigEditor_OnClick action
End Sub
