Attribute VB_Name = "modPrint"
Option Explicit

' ============================================================
' modPrint – Druckausgabe (ersetzt direkte PrintOut-Aufrufe)
' ============================================================

Public Sub PrintIzvestaj(ByVal data As Variant, ByVal reportTitle As String, _
                         ByVal headers As Variant)
    ' Generischer Report-Druck
    ' Schreibt in ein temporäres Print-Sheet und druckt
    
    Dim wsPrint As Worksheet
    On Error Resume Next
    Set wsPrint = ThisWorkbook.Sheets("_Print")
    On Error GoTo 0
    
    If wsPrint Is Nothing Then
        Set wsPrint = ThisWorkbook.Sheets.Add
        wsPrint.Name = "_Print"
    End If
    
    wsPrint.cells.Clear
    
    ' Titel
    wsPrint.Range("A1").Value = reportTitle
    wsPrint.Range("A1").Font.Size = 14
    wsPrint.Range("A1").Font.Bold = True
    
    ' Daten ausgeben
    OutputToSheet data, wsPrint.Range("A3"), headers
    
    ' Drucken
    wsPrint.PrintOut Copies:=1
    
    ' Aufräumen
    wsPrint.Visible = xlSheetVeryHidden
End Sub

Public Sub PrintOtkupniList(ByVal otkupID As String)
    ' Druckt einen einzelnen Otkupni list
    ' TODO: Template füllen und drucken
End Sub

