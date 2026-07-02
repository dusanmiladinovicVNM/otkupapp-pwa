VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWheelList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================
' clsWheelList - tanki WithEvents omotac oko JEDNE MSForms.ListBox.
'
' Jedina svrha: kad mis predje preko liste (MouseMove), zapamti bas
' tu listu kao "aktivnu" (mHot) u modMouseWheel. Tockic misa se hvata
' u modMouseWheel.MouseProc i skroluje iskljucivo tu, mHot, listu.
' Time se izbegava svaka geometrijska/DPI/HWND racunica (najmanja
' moguca povrsina za gresku).
'
' Instance se prave u modMouseWheel.BuildWrappers (isti "WithEvents
' omotac oko dinamicke kontrole" idiom kao clsBlokUI/modOtkupBlok) i
' zive u kolekciji mWrappers dok forma ima fokus.
'
' Bez stanja, bez API poziva -> ne moze da srusi ni uspori app.
' ============================================================
Option Explicit

Public WithEvents lst As MSForms.ListBox
Attribute lst.VB_VarHelpID = -1

Private Sub lst_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error Resume Next
    MouseWheel_SetHot lst
End Sub
