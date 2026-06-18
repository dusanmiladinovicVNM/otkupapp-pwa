VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' clsAppEvents - Application-level event hook (ENGINE / .xlam rezim).
'
' Add-in oslukuje otvaranje/zatvaranje SVIH workbook-ova. Kada se otvori
' klijentski .xlsx (prepoznat po marker tabeli tblConfig), pokrece aplikaciju.
' Tako klijentski fajl moze biti CIST .xlsx bez ijedne linije VBA koda.
'
' Instancira se i kaci u modBoot.InstallAppEvents.
' ============================================================

Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    On Error Resume Next
    If modContext.IsClientDataBook(Wb) Then
        modBoot.BootFromData Wb
    End If
End Sub

Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    On Error Resume Next
    If modContext.IsBoundDataBook(Wb) Then
        modBoot.ShutdownFromData
    End If
End Sub
