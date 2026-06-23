VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStmBtn2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' ============================================================
' clsStmBtn – Event-omotac za runtime dugme "Deaktiviraj/Aktiviraj"
' u frmStammdaten (soft-delete). Dugme se pravi u kodu (Controls.Add),
' pa se klik hvata preko WithEvents (kao clsLookupMenuBtn/clsBlokUI).
' ============================================================

Public WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1

Private Sub btn_Click()
    frmStammdaten.OnSoftDeleteClick
End Sub
