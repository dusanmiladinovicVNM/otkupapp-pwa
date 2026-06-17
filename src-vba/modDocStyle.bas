Attribute VB_Name = "modDocStyle"
Option Explicit

' ============================================================
' modDocStyle - zajednicki stil za stampane obrasce
' (otkupni / paletni / preradni list): logo, zaglavlje firme,
' naslov u dve linije, "labela + podebljana vrednost".
' Funkcije su Public da bi ih delili modPrint i modPaletniList.
' ============================================================

Public Function DocColHeaderFill() As Long
    DocColHeaderFill = RGB(217, 225, 242)
End Function

Public Function DocColGray() As Long
    DocColGray = RGB(90, 90, 90)
End Function

Public Function DocColRule() As Long
    DocColRule = RGB(110, 110, 110)
End Function

' GetConfigValue sa folbekom na podrazumevanu vrednost.
Public Function DocConfigOr(ByVal key As String, ByVal dflt As String) As String
    Dim v As String
    On Error Resume Next
    v = Trim$(CStr(GetConfigValue(key)))
    On Error GoTo 0
    If v = "" Then DocConfigOr = dflt Else DocConfigOr = v
End Function

' Cita config probajuci vise mogucih kljuceva, neosetljivo na velika/mala slova
' i na visak razmaka. Vraca prvu nepraznu vrednost; inace dflt.
Public Function DocConfigOrKeys(ByVal dflt As String, ParamArray keys() As Variant) As String
    Dim i As Long, v As String
    For i = LBound(keys) To UBound(keys)
        v = DocConfigCI(CStr(keys(i)))
        If v <> "" Then DocConfigOrKeys = v: Exit Function
    Next i
    DocConfigOrKeys = dflt
End Function

' Case-insensitive citanje iz tblSEFConfig (ConfigKey -> ConfigValue).
Public Function DocConfigCI(ByVal key As String) As String
    On Error GoTo done
    Dim d As Variant: d = GetTableData("tblSEFConfig")
    If IsEmpty(d) Then Exit Function
    Dim ki As Long, vi As Long
    ki = GetColumnIndex("tblSEFConfig", "ConfigKey")
    vi = GetColumnIndex("tblSEFConfig", "ConfigValue")
    If ki = 0 Or vi = 0 Then Exit Function
    Dim i As Long
    For i = 1 To UBound(d, 1)
        If StrComp(Trim$(CStr(d(i, ki))), Trim$(key), vbTextCompare) = 0 Then
            DocConfigCI = Trim$(CStr(d(i, vi)))
            Exit Function
        End If
    Next i
done:
End Function

' Podrazumevana klauzula o PDV nadoknadi (cl. 34 ZPDV). Override preko
' config kljuca OTKUP_KLAUZULA (vidi modConfig.CFG_OTKUP_KLAUZULA).
Public Function OtkupKlauzulaDefault() As String
    OtkupKlauzulaDefault = _
        "PDV nadoknada obracunata je u skladu sa clanom 34. Zakona o porezu na " & _
        "dodatu vrednost. Otkupljivac se obavezuje da poljoprivredniku isplati " & _
        "vrednost otkupljenih poljoprivrednih proizvoda uvecanu za iznos PDV " & _
        "nadoknade. Pravo na odbitak PDV nadoknade kao prethodnog poreza " & _
        "otkupljivac ostvaruje pod uslovom da je izvrsio isplatu poljoprivredniku " & _
        "na njegov tekuci racun."
End Function

' Putanja loga: config SELLER_LOGO_PATH, pa <workbook>\logo.png / logo.jpg. "" ako nema.
Public Function DocLogoPath() As String
    On Error Resume Next
    Dim p As String
    p = Trim$(CStr(GetConfigValue("SELLER_LOGO_PATH")))
    If p <> "" Then
        If Dir$(p) <> "" Then DocLogoPath = p: Exit Function
    End If
    Dim cand As String
    cand = ThisWorkbook.Path & "\logo.png"
    If Dir$(cand) <> "" Then DocLogoPath = cand: Exit Function
    cand = ThisWorkbook.Path & "\logo.jpg"
    If Dir$(cand) <> "" Then DocLogoPath = cand
End Function

' Ubacuje logo gore desno (preko zaglavlja). Tiho preskace ako ga nema.
Public Sub DocDrawLogo(ByVal ws As Worksheet, ByVal topRow As Long, ByVal rightCol As Long)
    On Error GoTo done
    Dim p As String: p = DocLogoPath()
    If p = "" Then Exit Sub

    Dim w As Double, hgt As Double
    w = 52: hgt = 40
    Dim rcell As Range: Set rcell = ws.cells(topRow, rightCol)
    Dim L As Double, T As Double
    L = rcell.Left + rcell.Width - w
    If L < ws.cells(topRow, 1).Left Then L = ws.cells(topRow, 1).Left
    T = rcell.Top

    ws.Shapes.AddPicture fileName:=p, LinkToFile:=msoFalse, _
                         SaveWithDocument:=msoTrue, _
                         Left:=L, Top:=T, Width:=w, Height:=hgt
done:
End Sub

' Upisuje "labela vrednost" u jednu celiju, sa podebljanim delom vrednosti.
Public Sub DocLabelVal(ByVal ws As Worksheet, ByVal rowIx As Long, ByVal colIx As Long, _
                       ByVal lbl As String, ByVal val As String)
    Dim s As String
    If val = "" Then s = lbl Else s = lbl & " " & val
    With ws.cells(rowIx, colIx)
        .value = s
        .Font.Bold = False
        If Len(val) > 0 Then
            On Error Resume Next
            .Characters(Start:=Len(lbl) + 2, Length:=Len(val)).Font.Bold = True
            On Error GoTo 0
        End If
    End With
End Sub

' Zaglavlje firme (naziv/adresa/PIB-MB-ziro) u koloni 1, logo gore desno,
' linija ispod. Cita SELLER_* iz configa. Vraca prvi slobodan red.
Public Function DocSellerHeader(ByVal ws As Worksheet, ByVal atRow As Long, _
                                ByVal lastCol As Long, ByVal rightCol As Long) As Long
    With ws.cells(atRow, 1)
        .value = GetConfigValue("SELLER_NAME")
        .Font.Bold = True
        .Font.Size = 12
    End With
    ws.cells(atRow + 1, 1).value = Trim$(GetConfigValue("SELLER_STREET") & ", " & _
        GetConfigValue("SELLER_POSTAL_CODE") & " " & GetConfigValue("SELLER_CITY"))
    With ws.cells(atRow + 2, 1)
        .value = "PIB: " & GetConfigValue("SELLER_PIB") & "    MB: " & _
                 GetConfigValue("SELLER_MATICNI_BROJ") & "    Ziro: " & _
                 GetConfigValue("SELLER_ACCOUNT")
        .Font.Size = 9
        .Font.Color = DocColGray()
    End With
    DocDrawLogo ws, atRow, rightCol
    With ws.Range(ws.cells(atRow + 2, 1), ws.cells(atRow + 2, lastCol)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = DocColRule()
    End With
    DocSellerHeader = atRow + 3
End Function

' Naslov: sitan opis + veliki naslov, centrirano preko 1..lastCol, linija ispod.
' Vraca prvi slobodan red.
Public Function DocTitleBlock(ByVal ws As Worksheet, ByVal atRow As Long, _
                              ByVal lastCol As Long, ByVal descriptor As String, _
                              ByVal title As String) As Long
    ws.Range(ws.cells(atRow, 1), ws.cells(atRow, lastCol)).Merge
    With ws.cells(atRow, 1)
        .value = descriptor
        .Font.Italic = True
        .Font.Size = 9
        .Font.Color = DocColGray()
        .HorizontalAlignment = xlCenter
    End With
    ws.Range(ws.cells(atRow + 1, 1), ws.cells(atRow + 1, lastCol)).Merge
    With ws.cells(atRow + 1, 1)
        .value = title
        .Font.Bold = True
        .Font.Size = 16
        .HorizontalAlignment = xlCenter
    End With
    ws.rows(atRow + 1).RowHeight = 22
    With ws.Range(ws.cells(atRow + 1, 1), ws.cells(atRow + 1, lastCol)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = DocColRule()
    End With
    DocTitleBlock = atRow + 2
End Function
