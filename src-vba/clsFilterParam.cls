VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFilterParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' clsFilterParam – Filter-Parameter Objekt
' ============================================================

Public ColIndex As Long     ' Spaltenindex im Array
Public Operator As String   ' "=", ">=", "<=", "BETWEEN", "<>", "LIKE"
Public Value1 As Variant    ' Primärwert
Public Value2 As Variant    ' Sekundärwert (nur für BETWEEN)

Public Function Init(ByVal col As Long, ByVal op As String, _
                     ByVal val1 As Variant, _
                     Optional ByVal val2 As Variant = Empty) As clsFilterParam
    ColIndex = col
    Operator = op
    Value1 = val1
    Value2 = val2
    Set Init = Me
End Function

