Option Explicit

' ============================================================
' modSmokeTestIzvestaj
' Brzi smoke test za modIzvestaj posle refaktora/hardening patch-eva
'
' Pokretanje:
'   SmokeTest_modIzvestaj
'
' Output:
'   Immediate Window / Ctrl+G
' ============================================================

Public Sub SmokeTest_modIzvestaj()
    On Error GoTo EH
    
    Dim datumOd As Date
    Dim datumDo As Date
    
    datumOd = DateSerial(Year(Date), 1, 1)
    datumDo = Date
    
    Debug.Print String(70, "=")
    Debug.Print "SmokeTest_modIzvestaj START | Period: " & _
                Format$(datumOd, "yyyy-mm-dd") & " - " & Format$(datumDo, "yyyy-mm-dd")
    Debug.Print String(70, "-")
    
    Dim stanicaID As String
    Dim kupacID As String
    Dim vozacID As String
    Dim kooperantID As String
    
    stanicaID = Smoke_FirstValue(TBL_STANICE, "StanicaID")
    kupacID = Smoke_FirstValue(TBL_KUPCI, COL_KUP_ID)
    vozacID = Smoke_FirstValue(TBL_VOZACI, "VozacID")
    kooperantID = Smoke_FirstValue(TBL_KOOPERANTI, COL_KOOP_ID)
    
    Debug.Print "Sample IDs:"
    Debug.Print "  StanicaID:   " & Smoke_TextOrSkip(stanicaID)
    Debug.Print "  KupacID:     " & Smoke_TextOrSkip(kupacID)
    Debug.Print "  VozacID:     " & Smoke_TextOrSkip(vozacID)
    Debug.Print "  KooperantID: " & Smoke_TextOrSkip(kooperantID)
    Debug.Print String(70, "-")
    
    ' ========================================================
    ' SALDO
    ' ========================================================
    If stanicaID <> "" Then
        Smoke_RunReport "ReportSaldoOM", ReportSaldoOM(stanicaID, datumOd, datumDo)
    Else
        Smoke_Skip "ReportSaldoOM", "Nema StanicaID u " & TBL_STANICE
    End If
    
    If kupacID <> "" Then
        Smoke_RunReport "ReportSaldoKupci", ReportSaldoKupci(kupacID, datumOd, datumDo)
    Else
        Smoke_Skip "ReportSaldoKupci", "Nema KupacID u " & TBL_KUPCI
    End If
    
    ' ========================================================
    ' KARTICA KOOPERANTA
    ' ========================================================
    If kooperantID <> "" Then
        Smoke_RunReport "ReportKarticaKooperanta", ReportKarticaKooperanta(kooperantID, datumOd, datumDo)
    Else
        Smoke_Skip "ReportKarticaKooperanta", "Nema KooperantID u " & TBL_KOOPERANTI
    End If
    
    ' ========================================================
    ' ISPLATA
    ' ========================================================
    If stanicaID <> "" Then
        Smoke_RunReport "ReportIsplata OM", ReportIsplata("OM", stanicaID, datumOd, datumDo)
    Else
        Smoke_Skip "ReportIsplata OM", "Nema StanicaID u " & TBL_STANICE
    End If
    
    If kupacID <> "" Then
        Smoke_RunReport "ReportIsplata Kupac", ReportIsplata("Kupac", kupacID, datumOd, datumDo)
    Else
        Smoke_Skip "ReportIsplata Kupac", "Nema KupacID u " & TBL_KUPCI
    End If
    
    ' ========================================================
    ' OTKUPLJENA ROBA
    ' ========================================================
    If stanicaID <> "" Then
        Smoke_RunReport "ReportOtkupRoba OM", ReportOtkupRoba("OM", stanicaID, datumOd, datumDo)
    Else
        Smoke_Skip "ReportOtkupRoba OM", "Nema StanicaID u " & TBL_STANICE
    End If
    
    If kupacID <> "" Then
        Smoke_RunReport "ReportOtkupRoba Kupac", ReportOtkupRoba("Kupac", kupacID, datumOd, datumDo)
    Else
        Smoke_Skip "ReportOtkupRoba Kupac", "Nema KupacID u " & TBL_KUPCI
    End If
    
    If vozacID <> "" Then
        Smoke_RunReport "ReportOtkupRoba Vozac", ReportOtkupRoba("Vozac", vozacID, datumOd, datumDo)
    Else
        Smoke_Skip "ReportOtkupRoba Vozac", "Nema VozacID u " & TBL_VOZACI
    End If
    
    ' ========================================================
    ' AMBALAZA
    ' ========================================================
    If stanicaID <> "" Then
        Smoke_RunReport "ReportAmbalaza OM pojedinacni", ReportAmbalaza("OM", stanicaID, datumOd, datumDo, False)
        Smoke_RunReport "ReportAmbalaza OM zbirni", ReportAmbalaza("OM", stanicaID, datumOd, datumDo, True)
    Else
        Smoke_Skip "ReportAmbalaza OM", "Nema StanicaID u " & TBL_STANICE
    End If
    
    If kupacID <> "" Then
        Smoke_RunReport "ReportAmbalaza Kupac pojedinacni", ReportAmbalaza("Kupac", kupacID, datumOd, datumDo, False)
        Smoke_RunReport "ReportAmbalaza Kupac zbirni", ReportAmbalaza("Kupac", kupacID, datumOd, datumDo, True)
    Else
        Smoke_Skip "ReportAmbalaza Kupac", "Nema KupacID u " & TBL_KUPCI
    End If
    
    If vozacID <> "" Then
        Smoke_RunReport "ReportAmbalaza Vozac pojedinacni", ReportAmbalaza("Vozac", vozacID, datumOd, datumDo, False)
        Smoke_RunReport "ReportAmbalaza Vozac zbirni", ReportAmbalaza("Vozac", vozacID, datumOd, datumDo, True)
    Else
        Smoke_Skip "ReportAmbalaza Vozac", "Nema VozacID u " & TBL_VOZACI
    End If
    
    ' ========================================================
    ' PROSECNA CENA
    ' ========================================================
    If stanicaID <> "" Then
        Smoke_RunReport "ReportProsecnaCena OM", ReportProsecnaCena("OM", stanicaID, datumOd, datumDo)
    Else
        Smoke_Skip "ReportProsecnaCena OM", "Nema StanicaID u " & TBL_STANICE
    End If
    
    If kupacID <> "" Then
        Smoke_RunReport "ReportProsecnaCena Kupac", ReportProsecnaCena("Kupac", kupacID, datumOd, datumDo)
    Else
        Smoke_Skip "ReportProsecnaCena Kupac", "Nema KupacID u " & TBL_KUPCI
    End If
    
    Smoke_RunReport "ReportProsecnaCena zbirni/all", ReportProsecnaCena("OM", "", datumOd, datumDo)
    
    ' ========================================================
    ' MANJAK
    ' ========================================================
    Smoke_RunReport "ReportManjak zbirni/all", ReportManjak("", "", datumOd, datumDo)
    
    If kupacID <> "" Then
        Smoke_RunReport "ReportManjak Kupac", ReportManjak("Kupac", kupacID, datumOd, datumDo)
    Else
        Smoke_Skip "ReportManjak Kupac", "Nema KupacID u " & TBL_KUPCI
    End If
    
    If vozacID <> "" Then
        Smoke_RunReport "ReportManjak Vozac", ReportManjak("Vozac", vozacID, datumOd, datumDo)
    Else
        Smoke_Skip "ReportManjak Vozac", "Nema VozacID u " & TBL_VOZACI
    End If
    
    ' ========================================================
    ' ZBIRNI
    ' ========================================================
    Smoke_RunReport "ReportZbirni OM", ReportZbirni("OM", datumOd, datumDo)
    Smoke_RunReport "ReportZbirni Kupac", ReportZbirni("Kupac", datumOd, datumDo)
    Smoke_RunReport "ReportZbirni Vozac", ReportZbirni("Vozac", datumOd, datumDo)
    
    Debug.Print String(70, "-")
    Debug.Print "SmokeTest_modIzvestaj OK"
    Debug.Print String(70, "=")
    Exit Sub

EH:
    Debug.Print String(70, "!")
    Debug.Print "SmokeTest_modIzvestaj ERROR"
    Debug.Print "Err.Number:      " & Err.Number
    Debug.Print "Err.Source:      " & Err.SOURCE
    Debug.Print "Err.Description: " & Err.description
    Debug.Print String(70, "!")
End Sub

Private Sub Smoke_RunReport(ByVal reportName As String, ByVal data As Variant)
    On Error GoTo EH
    
    Debug.Print Smoke_Pad(reportName, 38) & " | " & Smoke_ArrayShape(data)
    Exit Sub
    
EH:
    Debug.Print Smoke_Pad(reportName, 38) & " | ERROR | " & _
                Err.Number & " | " & Err.SOURCE & " | " & Err.description
End Sub

Private Sub Smoke_Skip(ByVal reportName As String, ByVal reason As String)
    Debug.Print Smoke_Pad(reportName, 38) & " | SKIP  | " & reason
End Sub

Private Function Smoke_ArrayShape(ByVal data As Variant) As String
    On Error GoTo EH
    
    If IsEmpty(data) Then
        Smoke_ArrayShape = "EMPTY"
    ElseIf Not IsArray(data) Then
        Smoke_ArrayShape = "NOT ARRAY"
    Else
        Smoke_ArrayShape = "OK    | " & _
                           CStr(UBound(data, 1)) & " rows x " & _
                           CStr(UBound(data, 2)) & " cols"
    End If
    
    Exit Function
    
EH:
    Smoke_ArrayShape = "INVALID ARRAY | " & Err.description
End Function

Private Function Smoke_FirstValue(ByVal tableName As String, ByVal columnName As String) As String
    On Error GoTo EH
    
    Dim data As Variant
    data = GetTableData(tableName)
    
    If IsEmpty(data) Or Not IsArray(data) Then
        Smoke_FirstValue = ""
        Exit Function
    End If
    
    Dim colIdx As Long
    colIdx = GetColumnIndex(tableName, columnName)
    
    If colIdx <= 0 Then
        Smoke_FirstValue = ""
        Exit Function
    End If
    
    Smoke_FirstValue = Trim$(CStr(data(1, colIdx)))
    Exit Function
    
EH:
    Smoke_FirstValue = ""
End Function

Private Function Smoke_TextOrSkip(ByVal value As String) As String
    If Trim$(value) = "" Then
        Smoke_TextOrSkip = "(nije pronadjen)"
    Else
        Smoke_TextOrSkip = value
    End If
End Function

Private Function Smoke_Pad(ByVal value As String, ByVal width As Long) As String
    If Len(value) >= width Then
        Smoke_Pad = Left$(value, width)
    Else
        Smoke_Pad = value & Space$(width - Len(value))
    End If
End Function

