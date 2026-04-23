Attribute VB_Name = "modHelpers"
Option Explicit

Public Function ExtractIDFromDisplay(ByVal displayText As String) As String
    ' Unterstützt: "ID - Name" und "(ID) Name"
    Dim dashPos As Long
    dashPos = InStr(displayText, " - ")
    If dashPos > 0 Then
        ExtractIDFromDisplay = Left$(displayText, dashPos - 1)
        Exit Function
    End If
    
    Dim startPos As Long, endPos As Long
    startPos = InStr(displayText, "(")
    endPos = InStr(displayText, ")")
    If startPos > 0 And endPos > startPos Then
        ExtractIDFromDisplay = Mid$(displayText, startPos + 1, endPos - startPos - 1)
        Exit Function
    End If
    
    ExtractIDFromDisplay = displayText
End Function

Public Function GetVozacDisplayList() As Variant
    Dim data As Variant
    data = GetTableData(TBL_VOZACI)
    If IsEmpty(data) Then
        GetVozacDisplayList = Array()
        Exit Function
    End If
    
    Dim colID As Long, colIme As Long, colPrezime As Long, colAktivan As Long
    colID = GetColumnIndex(TBL_VOZACI, "VozacID")
    colIme = GetColumnIndex(TBL_VOZACI, "Ime")
    colPrezime = GetColumnIndex(TBL_VOZACI, "Prezime")
    colAktivan = GetColumnIndex(TBL_VOZACI, "Aktivan")
    
    Dim result() As String
    Dim count As Long
    ReDim result(0 To UBound(data, 1) - 1)
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colAktivan)) = STATUS_AKTIVAN Then
            result(count) = CStr(data(i, colIme)) & " " & _
                           CStr(data(i, colPrezime)) & " (" & _
                           CStr(data(i, colID)) & ")"
            count = count + 1
        End If
    Next i
    
    If count = 0 Then
        GetVozacDisplayList = Array()
    Else
        ReDim Preserve result(0 To count - 1)
        GetVozacDisplayList = result
    End If
End Function

Public Sub FillCmb(ByRef cmb As MSForms.ComboBox, ByVal items As Variant)
    cmb.Clear
    If IsEmpty(items) Then Exit Sub
    If Not IsArray(items) Then Exit Sub
    Dim i As Long
    For i = LBound(items) To UBound(items)
        If CStr(items(i)) <> "" Then cmb.AddItem CStr(items(i))
    Next i
End Sub

Public Sub FillComboKooperantiByStanica(ByRef cmb As MSForms.ComboBox, ByVal stanicaID As String)
    cmb.Clear
    Dim data As Variant
    data = GetTableData(TBL_KOOPERANTI)
    If IsEmpty(data) Then Exit Sub
    
    Dim colID As Long, colIme As Long, colPrezime As Long, colStanica As Long
    colID = GetColumnIndex(TBL_KOOPERANTI, "KooperantID")
    colIme = GetColumnIndex(TBL_KOOPERANTI, "Ime")
    colPrezime = GetColumnIndex(TBL_KOOPERANTI, "Prezime")
    colStanica = GetColumnIndex(TBL_KOOPERANTI, "StanicaID")
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colStanica)) = stanicaID Then
            cmb.AddItem CStr(data(i, colID)) & " - " & _
                CStr(data(i, colIme)) & " " & CStr(data(i, colPrezime))
        End If
    Next i
End Sub

Public Function ExcludeStornirano(ByVal data As Variant, _
                                  ByVal tblName As String) As Variant
    ' Filtert Stornirano="Da" Zeilen raus, gibt bereinigtes Array zurück
    If IsEmpty(data) Then
        ExcludeStornirano = data
        Exit Function
    End If
    
    Dim colStorno As Long
    colStorno = GetColumnIndex(tblName, COL_STORNIRANO)
    If colStorno = 0 Then
        ExcludeStornirano = data
        Exit Function
    End If
    
    Dim filters As New Collection
    Dim fp As clsFilterParam
    Set fp = New clsFilterParam
    fp.Init colStorno, "<>", "Da"
    filters.Add fp
    
    ExcludeStornirano = FilterArray(data, filters)
End Function

Public Function SafeGetTable(ByVal tableName As String) As ListObject
    On Error Resume Next
    Set SafeGetTable = GetTable(tableName)
    On Error GoTo 0
End Function


Public Function Nz(ByVal val As Variant, Optional ByVal default As String = "") As String
    If IsEmpty(val) Or IsNull(val) Then
        Nz = default
    Else
        Nz = CStr(val)
    End If
End Function

Public Function BuildManjakDict(Optional ByVal filterZbirneKeys As Object = Nothing) As Object
    ' Returns: Dictionary BrojZbirne ? Array(ZbirnaKg, PrijemnicaKg)
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Zbirna
    Dim zbrData As Variant
    zbrData = GetTableData(TBL_ZBIRNA)
    If Not IsArray(zbrData) Then
        Set BuildManjakDict = dict
        Exit Function
    End If
    zbrData = ExcludeStornirano(zbrData, TBL_ZBIRNA)
    If Not IsArray(zbrData) Then
        Set BuildManjakDict = dict
        Exit Function
    End If
    
    Dim colBroj As Long, colZbrKol As Long
    colBroj = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ)
    colZbrKol = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_KOLICINA)
    
    Dim z As Long
    For z = 1 To UBound(zbrData, 1)
        Dim brZbr As String
        brZbr = CStr(zbrData(z, colBroj))
        If Not dict.Exists(brZbr) Then dict.Add brZbr, Array(0#, 0#)
        Dim vals As Variant
        vals = dict(brZbr)
        If IsNumeric(zbrData(z, colZbrKol)) Then vals(0) = vals(0) + CDbl(zbrData(z, colZbrKol))
        dict(brZbr) = vals
    Next z
    
    ' Prijemnica
    Dim prijData As Variant
    prijData = GetTableData(TBL_PRIJEMNICA)
    If Not IsArray(prijData) Then
        Set BuildManjakDict = dict
        Exit Function
    End If
    prijData = ExcludeStornirano(prijData, TBL_PRIJEMNICA)
    If Not IsArray(prijData) Then
        Set BuildManjakDict = dict
        Exit Function
    End If
    
    Dim colPBrZbr As Long, colPKol As Long
    colPBrZbr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
    colPKol = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA)
    
    Dim p As Long
    For p = 1 To UBound(prijData, 1)
        Dim pZbr As String
        pZbr = CStr(prijData(p, colPBrZbr))
        If dict.Exists(pZbr) Then
            vals = dict(pZbr)
            If IsNumeric(prijData(p, colPKol)) Then vals(1) = vals(1) + CDbl(prijData(p, colPKol))
            dict(pZbr) = vals
        End If
    Next p
    
    Set BuildManjakDict = dict
End Function

Public Function CheckVerwaisteDokumente() As String
    Dim warnings As String
    
    ' 1. Otkup ohne OtpremnicaID
    Dim otkupData As Variant
    otkupData = GetTableData(TBL_OTKUP)
    If IsArray(otkupData) Then
        otkupData = ExcludeStornirano(otkupData, TBL_OTKUP)
        If IsArray(otkupData) Then
            Dim colOtpID As Long, colOtkID As Long, colOtkKol As Long, colOtkBrDok As Long
            colOtpID = GetColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID)
            colOtkID = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
            colOtkKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
            colOtkBrDok = GetColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK)
            
            Dim cntOtkup As Long, detailOtkup As String, i As Long
            For i = 1 To UBound(otkupData, 1)
                If CStr(otkupData(i, colOtpID)) = "" Then
                    cntOtkup = cntOtkup + 1
                    If cntOtkup <= 40 Then
                        detailOtkup = detailOtkup & "  " & CStr(otkupData(i, colOtkID)) & _
                                      " (" & CStr(otkupData(i, colOtkBrDok)) & ") " & _
                                      Format$(CDbl(otkupData(i, colOtkKol)), "#,##0") & "kg" & vbCrLf
                    End If
                End If
            Next i
            If cntOtkup > 0 Then
                warnings = warnings & cntOtkup & " otkup(a) bez otpremnice:" & vbCrLf & detailOtkup
                If cntOtkup > 40 Then warnings = warnings & "  ..." & vbCrLf
            End If
        End If
    End If
    
    ' 2. Otpremnice ohne BrojZbirne
    Dim otpData As Variant
    otpData = GetTableData(TBL_OTPREMNICA)
    If IsArray(otpData) Then
        otpData = ExcludeStornirano(otpData, TBL_OTPREMNICA)
        If IsArray(otpData) Then
            Dim colOtpZbr As Long, colOtpBroj As Long, colOtpKol As Long, colOtpAmb As Long
            colOtpZbr = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE)
            colOtpBroj = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ)
            colOtpKol = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA)
            colOtpAmb = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_KOL_AMB)
            
            Dim cntOtp As Long, detailOtp As String
            For i = 1 To UBound(otpData, 1)
                If CStr(otpData(i, colOtpZbr)) = "" Then
                    cntOtp = cntOtp + 1
                    If cntOtp <= 40 Then
                        detailOtp = detailOtp & "  " & CStr(otpData(i, colOtpBroj)) & " " & _
                                    Format$(CDbl(otpData(i, colOtpKol)), "#,##0") & "kg " & _
                                    Format$(CLng(otpData(i, colOtpAmb)), "#,##0") & " amb" & vbCrLf
                    End If
                End If
            Next i
            If cntOtp > 0 Then
                warnings = warnings & cntOtp & " otpremnica(e) bez zbirne:" & vbCrLf & detailOtp
                If cntOtp > 40 Then warnings = warnings & "  ..." & vbCrLf
            End If
        End If
    End If
    
    ' 2. Verwaiste Otpremnice (stornierte Zbirna)
    Dim verwOtp As Variant
    verwOtp = GetVerwaisteDokumente("Otpremnica")
    If IsArray(verwOtp) Then
        warnings = warnings & UBound(verwOtp, 1) & " otpremnica(e) sa storniranom zbirnom:" & vbCrLf
        Dim o As Long
        For o = 1 To IIf(UBound(verwOtp, 1) > 5, 5, UBound(verwOtp, 1))
            warnings = warnings & "  " & CStr(verwOtp(o, 2)) & " (Zbr:" & CStr(verwOtp(o, 3)) & ") " & _
                       Format$(CDbl(verwOtp(o, 5)), "#,##0") & "kg" & vbCrLf
        Next o
        If UBound(verwOtp, 1) > 40 Then warnings = warnings & "  ..." & vbCrLf
    End If
    
    ' 3. Verwaiste Prijemnice (stornierte Zbirna)
    Dim verwPrij As Variant
    verwPrij = GetVerwaisteDokumente("Prijemnica")
    If IsArray(verwPrij) Then
        warnings = warnings & UBound(verwPrij, 1) & " prijemnica(e) sa storniranom zbirnom:" & vbCrLf
        Dim pr As Long
        For pr = 1 To IIf(UBound(verwPrij, 1) > 5, 5, UBound(verwPrij, 1))
            warnings = warnings & "  " & CStr(verwPrij(pr, 2)) & " (Zbr:" & CStr(verwPrij(pr, 3)) & ") " & _
                       Format$(CDbl(verwPrij(pr, 5)), "#,##0") & "kg" & vbCrLf
        Next pr
        If UBound(verwPrij, 1) > 40 Then warnings = warnings & "  ..." & vbCrLf
    End If
    
    ' 4. Zbirna ohne Prijemnica
    Dim zbrData As Variant
    zbrData = GetTableData(TBL_ZBIRNA)
    If IsArray(zbrData) Then
        zbrData = ExcludeStornirano(zbrData, TBL_ZBIRNA)
        If IsArray(zbrData) Then
            Dim prijData As Variant
            prijData = GetTableData(TBL_PRIJEMNICA)
            Dim prijDict As Object
            Set prijDict = CreateObject("Scripting.Dictionary")
            If IsArray(prijData) Then
                prijData = ExcludeStornirano(prijData, TBL_PRIJEMNICA)
                If IsArray(prijData) Then
                    Dim colPZbr As Long
                    colPZbr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
                    Dim p As Long
                    For p = 1 To UBound(prijData, 1)
                        Dim pKey As String
                        pKey = CStr(prijData(p, colPZbr))
                        If Not prijDict.Exists(pKey) Then prijDict.Add pKey, True
                    Next p
                End If
            End If
            
            Dim colZBroj As Long, colZKol As Long, colZAmb As Long
            colZBroj = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ)
            colZKol = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_KOLICINA)
            colZAmb = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_KOL_AMB)
            
            Dim cntZbr As Long, detailZbr As String
            Dim z As Long
            For z = 1 To UBound(zbrData, 1)
                Dim zBroj As String
                zBroj = CStr(zbrData(z, colZBroj))
                If Not prijDict.Exists(zBroj) Then
                    cntZbr = cntZbr + 1
                    If cntZbr <= 40 Then
                        Dim zKg As String: zKg = ""
                        If IsNumeric(zbrData(z, colZKol)) Then zKg = Format$(CDbl(zbrData(z, colZKol)), "#,##0") & "kg"
                        Dim zAmb As String: zAmb = ""
                        If IsNumeric(zbrData(z, colZAmb)) Then zAmb = Format$(CLng(zbrData(z, colZAmb)), "#,##0") & " amb"
                        detailZbr = detailZbr & "  " & zBroj & " " & zKg & " " & zAmb & vbCrLf
                    End If
                End If
            Next z
            
            If cntZbr > 0 Then
                warnings = warnings & cntZbr & " zbirna(e) bez prijemnice:" & vbCrLf & detailZbr
                If cntZbr > 40 Then warnings = warnings & "  ..." & vbCrLf
            End If
        End If
    End If
    
    CheckVerwaisteDokumente = warnings
End Function

