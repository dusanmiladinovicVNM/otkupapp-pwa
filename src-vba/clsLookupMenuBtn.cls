VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLookupMenuBtn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' ============================================================
' clsLookupMenuBtn - Event-omotac za dinamicki kreirano dugme
' menija "Maticni podaci" (vidi modMaticniLookups).
'
' Dugmad menija se prave u kodu (Controls.Add) da se NE bi dirao
' frmMaticniPodaci.frx. Klik se hvata samo preko WithEvents instance
' u ovoj klasi, koja prosledjuje event nazad u modMaticniLookups
' uz Tag/Naziv sekcije koju treba otvoriti u frmStammdaten.
'
' Isti obrazac kao clsBlokUI / modOtkupBlok.
' ============================================================

Public sekcijaTag As String       ' Tag za frmStammdaten (npr. "TipAmbalaze")
Public sekcijaCaption As String   ' Naziv u meniju / naslovu (npr. "Ambalaza")

Public WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1

Private Sub btn_Click()
    MaticniMenu_OnClick sekcijaTag, sekcijaCaption
End Sub

Private Sub btn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                          ByVal X As Single, ByVal Y As Single)
    MaticniMenu_OnHover btn
End Sub

