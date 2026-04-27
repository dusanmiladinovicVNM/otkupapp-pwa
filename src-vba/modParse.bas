Attribute VB_Name = "modParse"

Option Explicit

Public Function TryParseDouble(ByVal rawText As String, ByRef result As Double) As Boolean
    On Error GoTo EH

    Dim s As String
    s = NormalizeNumericText(rawText)

    If s = "" Then Exit Function

    If IsNumeric(s) Then
        result = CDbl(s)
        TryParseDouble = True
    End If

    Exit Function

EH:
    TryParseDouble = False
End Function

Public Function TryParseLong(ByVal rawText As String, ByRef result As Long) As Boolean
    On Error GoTo EH

    Dim d As Double

    If Not TryParseDouble(rawText, d) Then Exit Function
    If d < 0 Then Exit Function
    If Abs(d - Fix(d)) > 0.000001 Then Exit Function

    result = CLng(d)
    TryParseLong = True

    Exit Function

EH:
    TryParseLong = False
End Function

Public Function TryParseDateValue(ByVal rawText As String, ByRef result As Date) As Boolean
    On Error GoTo EH

    Dim s As String
    s = Trim$(rawText)

    If s = "" Then Exit Function

    If IsDate(s) Then
        result = CDate(s)
        TryParseDateValue = True
        Exit Function
    End If

    Dim parts() As String
    parts = Split(Replace(s, "/", "."), ".")

    If UBound(parts) = 2 Then
        Dim d As Long
        Dim m As Long
        Dim Y As Long

        If IsNumeric(parts(0)) And IsNumeric(parts(1)) And IsNumeric(parts(2)) Then
            d = CLng(parts(0))
            m = CLng(parts(1))
            Y = CLng(parts(2))

            If Y < 100 Then Y = 2000 + Y

            result = DateSerial(Y, m, d)
            TryParseDateValue = True
        End If
    End If

    Exit Function

EH:
    TryParseDateValue = False
End Function

Public Function NormalizeNumericText(ByVal rawText As String) As String
    Dim s As String
    s = Trim$(rawText)

    If s = "" Then Exit Function

    s = Replace(s, ChrW(160), "")
    s = Replace(s, " ", "")
    s = Replace(s, "RSD", "", , , vbTextCompare)
    s = Replace(s, "kg", "", , , vbTextCompare)

    Dim decSep As String
    decSep = Application.International(xlDecimalSeparator)

    Dim lastComma As Long
    Dim lastDot As Long

    lastComma = InStrRev(s, ",")
    lastDot = InStrRev(s, ".")

    If lastComma > 0 And lastDot > 0 Then
        ' Ako postoje i "." i ",", poslednji separator tretiramo kao decimalni.
        If lastComma > lastDot Then
            ' Format: 1.234,56
            s = Replace(s, ".", "")
            s = Replace(s, ",", decSep)
        Else
            ' Format: 1,234.56
            s = Replace(s, ",", "")
            s = Replace(s, ".", decSep)
        End If

    ElseIf lastComma > 0 Then
        s = Replace(s, ",", decSep)

    ElseIf lastDot > 0 Then
        s = Replace(s, ".", decSep)
    End If

    NormalizeNumericText = s
End Function
