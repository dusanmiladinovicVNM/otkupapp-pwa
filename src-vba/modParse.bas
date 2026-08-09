Attribute VB_Name = "modParse"

Option Explicit

' Deklarisan opseg poslovnih godina za TryParseDateValue. Van njega datum nije
' poslovni podatak nego artefakt parsiranja (npr. `CDate("12:30")` = 1899-12-30).
Private Const MIN_POSLOVNA_GODINA As Long = 1900
Private Const MAX_POSLOVNA_GODINA As Long = 2199

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

' Da li je datum u deklarisanom poslovnom opsegu. Isti kriterijum koriste
' `TryParseDateValue`/`TryParseBankaDateDMY` i staging writer banke, pa je opseg
' definisan na jednom mestu.
Public Function IsPoslovnaGodina(ByVal d As Date) As Boolean
    IsPoslovnaGodina = (Year(d) >= MIN_POSLOVNA_GODINA And Year(d) <= MAX_POSLOVNA_GODINA)
End Function

' Deterministicko `d.m.yyyy` / `d/m/yyyy` parsiranje - BEZ `CDate`, pa bez uticaja
' regionalnih podesavanja masine. To je format koji daju parseri izvoda banaka
' (`dd.mm.yyyy`), gde je zamena dana i meseca tiho pogresno datiranje transakcije.
' Vraca False za sve sto nije taj oblik (pozivalac sme da proba dalje).
Public Function TryParseBankaDateDMY(ByVal rawText As String, ByRef result As Date) As Boolean
    On Error GoTo EH

    Dim s As String
    Dim parts() As String
    Dim d As Long
    Dim m As Long
    Dim Y As Long
    Dim probe As Date

    s = Trim$(rawText)
    If s = "" Then Exit Function

    parts = Split(Replace(s, "/", "."), ".")
    If UBound(parts) <> 2 Then Exit Function

    If Not (IsNumeric(parts(0)) And IsNumeric(parts(1)) And IsNumeric(parts(2))) Then Exit Function

    d = CLng(parts(0))
    m = CLng(parts(1))
    Y = CLng(parts(2))

    If Y < 100 Then Y = 2000 + Y

    ' AUD-007: DateSerial NE puca na nemogucem datumu nego se "prelije"
    ' (30.02.2026 -> 02.03.2026, 32.01. -> 01.02., mesec 13 -> januar sledece
    ' godine). Zato: prvo opseg, pa round-trip - sto je uslo mora i da izadje.
    If d < 1 Or d > 31 Then Exit Function
    If m < 1 Or m > 12 Then Exit Function
    If Y < MIN_POSLOVNA_GODINA Or Y > MAX_POSLOVNA_GODINA Then Exit Function

    probe = DateSerial(Y, m, d)
    If Day(probe) <> d Then Exit Function
    If Month(probe) <> m Then Exit Function
    If Year(probe) <> Y Then Exit Function

    result = probe
    TryParseBankaDateDMY = True

    Exit Function

EH:
    TryParseBankaDateDMY = False
End Function

Public Function TryParseDateValue(ByVal rawText As String, ByRef result As Date) As Boolean
    On Error GoTo EH

    Dim s As String
    Dim probe As Date

    s = Trim$(rawText)

    If s = "" Then Exit Function

    ' REDOSLED JE BITAN. `d.m.yyyy` (i `d/m/yyyy`) se parsira DETERMINISTICKI, PRE
    ' `IsDate`/`CDate`. `CDate` je locale-zavisan: na masini sa MM/DD podesavanjem
    ' "01.02.2026" postaje 2. januar umesto 1. februara -- a bas taj format daju
    ' parseri izvoda banaka. Ranije je `IsDate` grana isla prva i tiho pomerala
    ' datum transakcije. `IsDate` ostaje kao fallback za ostale zapise (npr.
    ' "2026-02-01" ili vrednost koja je vec Date).
    If TryParseBankaDateDMY(s, probe) Then
        result = probe
        TryParseDateValue = True
        Exit Function
    End If

    If Not IsDate(s) Then Exit Function

    probe = CDate(s)

    ' Ista kapija opsega kao u DMY grani: `CDate("12:30")` daje 1899-12-30, a
    ' 1899. nije poslovni datum -- ranije je ta grana pisala pravo u `result`.
    If Year(probe) < MIN_POSLOVNA_GODINA Or Year(probe) > MAX_POSLOVNA_GODINA Then Exit Function

    result = probe
    TryParseDateValue = True

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
