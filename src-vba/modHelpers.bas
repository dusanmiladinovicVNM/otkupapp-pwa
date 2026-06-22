Attribute VB_Name = "modHelpers"
Option Explicit

' Format kolicine (kg) za izvestaje: ceo broj BEZ decimalnog zareza
' (1000 -> "1.000"), a sa decimalom prikazi decimale (1234.5 -> "1.234,5").
' Eksplicitan If jer Format$(n,"#,##0.##") u nekim lokalizacijama (DE) ostavi
' prazan decimalni zarez ("500,"). Jedinstveni izvor istine za kg prikaz.
Public Function FmtKolicina(ByVal x As Double) As String
    If x = Int(x) Then
        FmtKolicina = Format$(x, "#,##0")
    Else
        FmtKolicina = Format$(x, "#,##0.##")
    End If
End Function

Public Function ExtractIDFromDisplay(ByVal displayText As String) As String
    ' Unterst�tzt: "ID - Name" und "(ID) Name"
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

    ' 2-kolonski: kol0 = "Ime Prezime" (vidljivo, filter po imenu),
    ' kol1 = KooperantID (skriveno). BoundColumn=1 -> .Value je ime
    ' (omogucava i slobodan unos novog imena). GetComboID cita kol1.
    On Error Resume Next
    cmb.ColumnCount = 2
    cmb.ColumnWidths = ";0"
    cmb.TextColumn = 1
    cmb.BoundColumn = 1
    cmb.MatchEntry = fmMatchEntryComplete
    cmb.MatchRequired = False
    On Error GoTo 0

    Dim data As Variant
    data = GetTableData(TBL_KOOPERANTI)
    If IsEmpty(data) Then Exit Sub

    Dim colID As Long, colIme As Long, colPrezime As Long, colStanica As Long
    colID = GetColumnIndex(TBL_KOOPERANTI, "KooperantID")
    colIme = GetColumnIndex(TBL_KOOPERANTI, "Ime")
    colPrezime = GetColumnIndex(TBL_KOOPERANTI, "Prezime")
    colStanica = GetColumnIndex(TBL_KOOPERANTI, "StanicaID")
    Dim colAkt As Long: colAkt = GetColumnIndex(TBL_KOOPERANTI, "Aktivan")

    Dim names() As String, ids() As String
    ReDim names(1 To UBound(data, 1))
    ReDim ids(1 To UBound(data, 1))
    Dim n As Long: n = 0

    Dim i As Long
    For i = 1 To UBound(data, 1)
        ' stanicaID = "" -> svi kooperanti (toggle KOOP_FILTER_BY_OM = OFF)
        Dim aktOk As Boolean: aktOk = True
        If colAkt > 0 Then
            If StrComp(Trim$(CStr(data(i, colAkt))), STATUS_NEAKTIVAN, vbTextCompare) = 0 Then aktOk = False
        End If
        If (stanicaID = "" Or CStr(data(i, colStanica)) = stanicaID) And aktOk Then
            n = n + 1
            names(n) = Trim$(CStr(data(i, colIme)) & " " & CStr(data(i, colPrezime)))
            ids(n) = CStr(data(i, colID))
        End If
    Next i
    If n = 0 Then Exit Sub

    ' insertion sort po imenu (case-insensitive)
    Dim a As Long, b As Long
    For a = 2 To n
        Dim kn As String: kn = names(a)
        Dim ki As String: ki = ids(a)
        b = a - 1
        Do While b >= 1
            If LCase$(names(b)) <= LCase$(kn) Then Exit Do
            names(b + 1) = names(b): ids(b + 1) = ids(b): b = b - 1
        Loop
        names(b + 1) = kn: ids(b + 1) = ki
    Next a

    For a = 1 To n
        cmb.AddItem names(a)
        cmb.List(cmb.ListCount - 1, 1) = ids(a)
    Next a
End Sub

Public Function ExcludeStornirano(ByVal data As Variant, _
                                  ByVal tblName As String) As Variant
    ' Filtert Stornirano="Da" Zeilen raus, gibt bereinigtes Array zur�ck
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

' Null-safe pretvaranje vrednosti iz tabele/celije u tekst.
' Vraca "" za Null/Empty/Error; inace CStr(v). Pozivaoci sami rade Trim$ gde treba.
Public Function NzToText(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then
        NzToText = ""
    ElseIf IsError(v) Then
        NzToText = ""
    Else
        NzToText = CStr(v)
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

' ============================================================
' GetKontrolaPregled - vidljivi dnevni kontrolni pregled (dashboard).
'
' Vraca formatiran tekst kontrolnih zbirova za lblStatus karticu na
' frmOtkupAPP (poziva se iz UserForm_Activate). Postavlja imaProblema =
' True ako postoji greska (neuskladjene kolicine, nevalidni otkup redovi,
' ili orphani iz CheckVerwaisteDokumente) -> pozivalac boji karticu crveno.
'
' Jedan prolaz po tabeli (Activate se okida cesto -> mora biti brzo).
' Reuse: CheckVerwaisteDokumente (orphani), ExcludeStornirano.
' ============================================================
Public Function GetKontrolaPregled(ByRef imaProblema As Boolean) As String
    On Error GoTo EH

    imaProblema = False

    Dim otk As Variant, otp As Variant, zbr As Variant, prj As Variant, amb As Variant
    otk = GetTableData(TBL_OTKUP):       If IsArray(otk) Then otk = ExcludeStornirano(otk, TBL_OTKUP)
    otp = GetTableData(TBL_OTPREMNICA):  If IsArray(otp) Then otp = ExcludeStornirano(otp, TBL_OTPREMNICA)
    zbr = GetTableData(TBL_ZBIRNA):      If IsArray(zbr) Then zbr = ExcludeStornirano(zbr, TBL_ZBIRNA)
    prj = GetTableData(TBL_PRIJEMNICA):  If IsArray(prj) Then prj = ExcludeStornirano(prj, TBL_PRIJEMNICA)
    amb = GetTableData(TBL_AMBALAZA):    If IsArray(amb) Then amb = ExcludeStornirano(amb, TBL_AMBALAZA)

    Dim otkKgByOtp As Object: Set otkKgByOtp = CreateObject("Scripting.Dictionary")
    Dim otpKgByZbr As Object: Set otpKgByZbr = CreateObject("Scripting.Dictionary")
    Dim zbrKgByBroj As Object: Set zbrKgByBroj = CreateObject("Scripting.Dictionary")
    Dim prjKgByBroj As Object: Set prjKgByBroj = CreateObject("Scripting.Dictionary")
    Dim ambByTip As Object: Set ambByTip = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim cntNeispl As Long, valNeispl As Double, cntInvalid As Long

    ' --- OTKUP: grupisanje po OtpremnicaID, neisplaceno, nevalidni redovi ---
    If IsArray(otk) Then
        Dim cOtkOtp As Long, cOtkKol As Long, cOtkCena As Long, cOtkIspl As Long
        cOtkOtp = GetColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID)
        cOtkKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
        cOtkCena = GetColumnIndex(TBL_OTKUP, COL_OTK_CENA)
        cOtkIspl = GetColumnIndex(TBL_OTKUP, COL_OTK_ISPLACENO)

        Dim kol As Double, cena As Double, otpID As String
        For i = 1 To UBound(otk, 1)
            kol = NumOrZero(otk(i, cOtkKol))
            cena = NumOrZero(otk(i, cOtkCena))
            otpID = Trim$(CStr(otk(i, cOtkOtp)))

            If Len(otpID) > 0 Then DAdd otkKgByOtp, otpID, kol
            If kol <= 0 Or cena <= 0 Then cntInvalid = cntInvalid + 1

            If UCase$(Trim$(CStr(otk(i, cOtkIspl)))) <> UCase$(STATUS_ISPLACENO) Then
                cntNeispl = cntNeispl + 1
                valNeispl = valNeispl + kol * cena
            End If
        Next i
    End If

    ' --- OTPREMNICA: balans otkup<->otpremnica + kg po BrojZbirne ---
    Dim cntMisOtk As Long
    If IsArray(otp) Then
        Dim cOtpID As Long, cOtpKol As Long, cOtpZbr As Long
        cOtpID = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_ID)
        cOtpKol = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA)
        cOtpZbr = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE)

        Dim oID As String, oKol As Double, oZbr As String, sumOtk As Double
        For i = 1 To UBound(otp, 1)
            oID = Trim$(CStr(otp(i, cOtpID)))
            oKol = NumOrZero(otp(i, cOtpKol))
            oZbr = Trim$(CStr(otp(i, cOtpZbr)))

            If Len(oZbr) > 0 Then DAdd otpKgByZbr, oZbr, oKol

            If otkKgByOtp.Exists(oID) Then sumOtk = CDbl(otkKgByOtp(oID)) Else sumOtk = 0#
            If Abs(sumOtk - oKol) >= 0.01 Then cntMisOtk = cntMisOtk + 1
        Next i
    End If

    ' --- ZBIRNA: kg po broju ---
    If IsArray(zbr) Then
        Dim cZbrBroj As Long, cZbrKol As Long
        cZbrBroj = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ)
        cZbrKol = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_KOLICINA)
        Dim zBroj As String
        For i = 1 To UBound(zbr, 1)
            zBroj = Trim$(CStr(zbr(i, cZbrBroj)))
            If Len(zBroj) > 0 Then DAdd zbrKgByBroj, zBroj, NumOrZero(zbr(i, cZbrKol))
        Next i
    End If

    ' --- PRIJEMNICA: kg po broju zbirne ---
    If IsArray(prj) Then
        Dim cPrjZbr As Long, cPrjKol As Long
        cPrjZbr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
        cPrjKol = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA)
        Dim pBroj As String
        For i = 1 To UBound(prj, 1)
            pBroj = Trim$(CStr(prj(i, cPrjZbr)))
            If Len(pBroj) > 0 Then DAdd prjKgByBroj, pBroj, NumOrZero(prj(i, cPrjKol))
        Next i
    End If

    ' --- AMBALAZA: neto saldo po EntitetTip (Smer Ulaz=+, Izlaz=-) ---
    If IsArray(amb) Then
        Dim cAmbSmer As Long, cAmbKol As Long, cAmbTip As Long
        cAmbSmer = GetColumnIndex(TBL_AMBALAZA, COL_AMB_SMER)
        cAmbKol = GetColumnIndex(TBL_AMBALAZA, COL_AMB_KOLICINA)
        cAmbTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET_TIP)

        If cAmbSmer > 0 And cAmbKol > 0 And cAmbTip > 0 Then
            Dim sm As String, et As String
            For i = 1 To UBound(amb, 1)
                If IsNumeric(amb(i, cAmbKol)) Then
                    sm = Trim$(CStr(amb(i, cAmbSmer)))
                    et = Trim$(CStr(amb(i, cAmbTip)))
                    If Len(et) > 0 Then
                        If sm = "Ulaz" Then
                            DAdd ambByTip, et, CDbl(amb(i, cAmbKol))
                        ElseIf sm = "Izlaz" Then
                            DAdd ambByTip, et, -CDbl(amb(i, cAmbKol))
                        End If
                    End If
                End If
            Next i
        End If
    End If

    ' --- IZRACUNAJ ZBIROVE (otvorene zbirne, manjak, otpremnica<->zbirna) ---
    Dim cntOtvorene As Long, kgOtvorene As Double, manjak As Double
    Dim cntMisOtpZbr As Long
    Dim kKey As Variant, d As Double
    For Each kKey In zbrKgByBroj.keys
        If prjKgByBroj.Exists(kKey) Then
            d = CDbl(zbrKgByBroj(kKey)) - CDbl(prjKgByBroj(kKey))
            If d > 0 Then manjak = manjak + d
        Else
            cntOtvorene = cntOtvorene + 1
            kgOtvorene = kgOtvorene + CDbl(zbrKgByBroj(kKey))
        End If

        If otpKgByZbr.Exists(kKey) Then
            If Abs(CDbl(otpKgByZbr(kKey)) - CDbl(zbrKgByBroj(kKey))) >= 0.01 Then
                cntMisOtpZbr = cntMisOtpZbr + 1
            End If
        End If
    Next kKey

    ' --- SASTAVI TEKST: kontrolni zbirovi ---
    Dim s As String
    s = "DNEVNI PREGLED (Kontrola)" & vbCrLf
    s = s & "- Otvorene zbirne (bez prijemnice): " & cntOtvorene & _
            "  |  " & Format$(kgOtvorene, "#,##0") & " kg" & vbCrLf
    s = s & "- Ukupan manjak (primljeno): " & Format$(manjak, "#,##0") & " kg" & vbCrLf
    s = s & "- Ambalaza saldo: " & FormatAmbSaldo(ambByTip) & vbCrLf
    s = s & "- Neisplaceno: " & cntNeispl & " otkup(a)  |  " & Format$(valNeispl, "#,##0")

    ' --- PROBLEMI: balans + nevalidni redovi + orphani (reuse) ---
    Dim prob As String
    If cntMisOtk > 0 Then prob = prob & "- Neuskladjene kolicine otkup<->otpremnica: " & cntMisOtk & vbCrLf
    If cntMisOtpZbr > 0 Then prob = prob & "- Neuskladjene kolicine otpremnica<->zbirna: " & cntMisOtpZbr & vbCrLf
    If cntInvalid > 0 Then prob = prob & "- Nevalidni otkup redovi (kolicina/cena <= 0): " & cntInvalid & vbCrLf

    Dim orphani As String
    orphani = CheckVerwaisteDokumente()

    If Len(prob) > 0 Or Len(orphani) > 0 Then
        imaProblema = True
        s = s & vbCrLf & vbCrLf & "PROBLEMI:" & vbCrLf & prob & orphani
    End If

    GetKontrolaPregled = s
    Exit Function

EH:
    LogErr "modHelpers.GetKontrolaPregled"
    imaProblema = False
    GetKontrolaPregled = "DNEVNI PREGLED (Kontrola)" & vbCrLf & _
                         "- nedostupno (greska pri racunanju)"
End Function

' Akumuliraj v u dict(k) (Scripting.Dictionary).
Private Sub DAdd(ByVal d As Object, ByVal k As String, ByVal v As Double)
    If d.Exists(k) Then
        d(k) = CDbl(d(k)) + v
    Else
        d(k) = v
    End If
End Sub

' Numericka vrednost celije ili 0.
Private Function NumOrZero(ByVal v As Variant) As Double
    If IsNumeric(v) Then NumOrZero = CDbl(v) Else NumOrZero = 0#
End Function

' Saldo po EntitetTip-u: "Kooperant +540  |  Stanica -200  |  Kupac -340".
Private Function FormatAmbSaldo(ByVal d As Object) As String
    Dim s As String, kk As Variant
    For Each kk In d.keys
        If Len(s) > 0 Then s = s & "  |  "
        s = s & CStr(kk) & " " & Format$(CDbl(d(kk)), "+#,##0;-#,##0;0")
    Next kk
    If Len(s) = 0 Then s = "-"
    FormatAmbSaldo = s
End Function

