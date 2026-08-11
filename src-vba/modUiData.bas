Attribute VB_Name = "modUiData"
'=====================================================================
' modUiData - PRISTUP TABELAMA za novi UI (faza S4b).
'
' Kes pune tabele i cetiri citaca celije. Nista vise: ni jedno ime ekrana,
' ni jedno ime kolone, nijedno poslovno pravilo.
'
' Postoji zato sto od S4b i ljuska i ekranski moduli citaju iste tabele:
' ljuska za KPI i combo-e, ekran za svoje redove. Da je ostalo u ljusci,
' ekran bi morao da zove njeno privatno telo; da je otislo u ekran, ljuska
' bi zavisila od ekrana. Zajednicki sloj resava oba.
'
' Kes se prazni pri gradnji ekrana i posle upisa (modOtkupUI.RefreshFromData).
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const UIDATA_BUILD As String = "v6-ui-95"

Private mCache As Object

Public Sub ResetCache()
    Set mCache = Nothing
End Sub

Public Function CachedTable(ByVal tblName As String) As Variant
    If mCache Is Nothing Then Set mCache = CreateObject("Scripting.Dictionary")
    If Not mCache.Exists(tblName) Then mCache(tblName) = GetTableData(tblName)
    CachedTable = mCache(tblName)
End Function

Public Function ColIdx(ByVal tblName As String, ByVal colName As String) As Long
    If Len(colName) = 0 Then Exit Function
    On Error Resume Next
    ColIdx = GetColumnIndex(tblName, colName)
End Function

Public Function CellS(ByRef src As Variant, ByVal r As Long, ByVal c As Long) As String
    If c < 1 Then Exit Function
    Dim v As Variant: v = src(r, c)
    If IsEmpty(v) Then Exit Function
    CellS = Trim$(CStr(v))
End Function

Public Function CellD(ByRef src As Variant, ByVal r As Long, ByVal c As Long) As Double
    If c < 1 Then Exit Function
    Dim v As Variant: v = src(r, c)
    If IsNumeric(v) Then CellD = CDbl(v)
End Function

Public Function CellDate(ByRef src As Variant, ByVal r As Long, ByVal c As Long) As Double
    If c < 1 Then Exit Function
    Dim v As Variant: v = src(r, c)
    If IsNumeric(v) Then
        CellDate = Int(CDbl(v))
    ElseIf IsDate(v) Then
        CellDate = Int(CDbl(CDate(v)))
    End If
End Function
