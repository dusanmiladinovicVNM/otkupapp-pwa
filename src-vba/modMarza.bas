Attribute VB_Name = "modMarza"

Option Explicit


' ============================================================
' modMarza v3.0 - Margenberechnung Business Logic
' Alle Funktionen geben 2D-Arrays zurueck
' Spalten: Vrsta, OtkKg, ProsekOtk, IspKg, IspRSD, OtkKosten, Marza, MarzaPct
' ============================================================

' ============================================================
' SHARED: Dict ? 2D Result (Kupac-Modus)
' ============================================================

Private Function BuildMarzaResult(ByVal dictVerkauf As Object, _
                                  ByVal dictEinkauf As Object) As Variant
    ' dictVerkauf = Prijemnice (VK), dictEinkauf = Otkup (EK)
    ' Basis = VK-Zeilen, EK liefert Durchschnittspreis
    
    If dictVerkauf.count = 0 Then
        BuildMarzaResult = Empty
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To dictVerkauf.count + 1, 1 To 8)
    
    Dim keys As Variant
    keys = dictVerkauf.keys
    Dim totalOtkRSD As Double, totalIspKg As Double
    Dim totalIspRSD As Double, totalMarza As Double
    
    Dim i As Long
    For i = 0 To dictVerkauf.count - 1
        Dim valsV As Variant: valsV = dictVerkauf(keys(i))
        Dim ispKg As Double: ispKg = valsV(0)
        Dim ispRSD As Double: ispRSD = valsV(1)
        
        Dim otkKg As Double: otkKg = 0
        Dim otkRSD As Double: otkRSD = 0
        Dim prosekOtk As Double: prosekOtk = 0
        
        If dictEinkauf.Exists(keys(i)) Then
            Dim valsE As Variant: valsE = dictEinkauf(keys(i))
            otkKg = valsE(0)
            otkRSD = valsE(1)
            If otkKg > 0 Then prosekOtk = otkRSD / otkKg
        End If
        
        Dim otkKosten As Double: otkKosten = ispKg * prosekOtk
        Dim marza As Double: marza = ispRSD - otkKosten
        Dim marzaPct As Double
        If ispRSD > 0 Then marzaPct = marza / ispRSD * 100 Else marzaPct = 0
        
        result(i + 1, 1) = keys(i)
        result(i + 1, 2) = otkKg
        result(i + 1, 3) = prosekOtk
        result(i + 1, 4) = ispKg
        result(i + 1, 5) = ispRSD
        result(i + 1, 6) = otkKosten
        result(i + 1, 7) = marza
        result(i + 1, 8) = marzaPct
        
        totalOtkRSD = totalOtkRSD + otkKosten
        totalIspKg = totalIspKg + ispKg
        totalIspRSD = totalIspRSD + ispRSD
        totalMarza = totalMarza + marza
    Next i
    
    ' UKUPNO
    Dim u As Long: u = dictVerkauf.count + 1
    result(u, 1) = "UKUPNO"
    result(u, 2) = ""
    result(u, 3) = ""
    result(u, 4) = totalIspKg
    result(u, 5) = totalIspRSD
    result(u, 6) = totalOtkRSD
    result(u, 7) = totalMarza
    If totalIspRSD > 0 Then result(u, 8) = totalMarza / totalIspRSD * 100 Else result(u, 8) = 0
    
    BuildMarzaResult = result
End Function

' ============================================================
' SHARED: Dict ? 2D Result (OM-Modus - umgekehrte Logik)
' ============================================================

Private Function BuildMarzaResultOM(ByVal dictEinkauf As Object, _
                                    ByVal dictVerkauf As Object) As Variant
    ' OM: EK ist Basis, VK-Durchschnitt als fiktiver Erloes
    
    If dictEinkauf.count = 0 Then
        BuildMarzaResultOM = Empty
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To dictEinkauf.count + 1, 1 To 8)
    
    Dim keys As Variant
    keys = dictEinkauf.keys
    Dim totalOtkKg As Double, totalOtkRSD As Double
    Dim totalIspRSD As Double, totalMarza As Double
    
    Dim i As Long
    For i = 0 To dictEinkauf.count - 1
        Dim valsO As Variant: valsO = dictEinkauf(keys(i))
        Dim otkKg As Double: otkKg = valsO(0)
        Dim otkRSD As Double: otkRSD = valsO(1)
        Dim prosekOtk As Double
        If otkKg > 0 Then prosekOtk = otkRSD / otkKg Else prosekOtk = 0
        
        Dim prosekVK As Double: prosekVK = 0
        If dictVerkauf.Exists(keys(i)) Then
            Dim valsP As Variant: valsP = dictVerkauf(keys(i))
            If valsP(0) > 0 Then prosekVK = valsP(1) / valsP(0)
        End If
        
        Dim ispErloes As Double: ispErloes = otkKg * prosekVK
        Dim marza As Double: marza = ispErloes - otkRSD
        Dim marzaPct As Double
        If ispErloes > 0 Then marzaPct = marza / ispErloes * 100 Else marzaPct = 0
        
        result(i + 1, 1) = keys(i)
        result(i + 1, 2) = otkKg
        result(i + 1, 3) = prosekOtk
        result(i + 1, 4) = otkKg       ' IspKg = OtkKg bei OM
        result(i + 1, 5) = ispErloes
        result(i + 1, 6) = otkRSD
        result(i + 1, 7) = marza
        result(i + 1, 8) = marzaPct
        
        totalOtkKg = totalOtkKg + otkKg
        totalOtkRSD = totalOtkRSD + otkRSD
        totalIspRSD = totalIspRSD + ispErloes
        totalMarza = totalMarza + marza
    Next i
    
    Dim u As Long: u = dictEinkauf.count + 1
    result(u, 1) = "UKUPNO"
    result(u, 2) = totalOtkKg
    result(u, 3) = ""
    result(u, 4) = ""
    result(u, 5) = totalIspRSD
    result(u, 6) = totalOtkRSD
    result(u, 7) = totalMarza
    If totalIspRSD > 0 Then result(u, 8) = totalMarza / totalIspRSD * 100 Else result(u, 8) = 0
    
    BuildMarzaResultOM = result
End Function

' ============================================================
' SHARED HELPERS
' ============================================================

Private Function AggregatePrijemniceByVrsta(ByVal prijData As Variant, _
                                            ByVal vrstaCache As Object) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    If IsEmpty(prijData) Then
        Set AggregatePrijemniceByVrsta = dict
        Exit Function
    End If
    
    Dim colKol As Long, colCena As Long, colBrZbr As Long
    colKol = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA)
    colCena = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_CENA)
    colBrZbr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
    
    Dim i As Long
    For i = 1 To UBound(prijData, 1)
        Dim vrsta As String
        vrsta = GetVrstaFromCache(vrstaCache, CStr(prijData(i, colBrZbr)))
        If vrsta = "" Then vrsta = "(Nepoznato)"
        
        If Not dict.Exists(vrsta) Then dict.Add vrsta, Array(0#, 0#)
        Dim vals As Variant
        vals = dict(vrsta)
        If IsNumeric(prijData(i, colKol)) Then vals(0) = vals(0) + CDbl(prijData(i, colKol))
        If IsNumeric(prijData(i, colKol)) And IsNumeric(prijData(i, colCena)) Then
            vals(1) = vals(1) + CDbl(prijData(i, colKol)) * CDbl(prijData(i, colCena))
        End If
        dict(vrsta) = vals
    Next i
    
    Set AggregatePrijemniceByVrsta = dict
End Function

Private Function AggregatePrijemniceByVrstaFiltered(ByVal data As Variant, _
                                                     ByVal vrstaCache As Object, _
                                                     ByVal datumOd As Date, _
                                                     ByVal datumDo As Date) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    If IsEmpty(data) Then
        Set AggregatePrijemniceByVrstaFiltered = dict
        Exit Function
    End If
    
    Dim colKol As Long, colCena As Long, colBrZbr As Long, colDat As Long
    colKol = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA)
    colCena = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_CENA)
    colBrZbr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
    colDat = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_DATUM)
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If datumOd > 0 Then
            If IsDate(data(i, colDat)) Then
                If CDate(data(i, colDat)) < datumOd Or CDate(data(i, colDat)) > datumDo Then GoTo NextPrij
            End If
        End If
        
        Dim vrsta As String
        vrsta = GetVrstaFromCache(vrstaCache, CStr(data(i, colBrZbr)))
        If vrsta = "" Then vrsta = "(Nepoznato)"
        
        If Not dict.Exists(vrsta) Then dict.Add vrsta, Array(0#, 0#)
        Dim vals As Variant: vals = dict(vrsta)
        If IsNumeric(data(i, colKol)) Then vals(0) = vals(0) + CDbl(data(i, colKol))
        If IsNumeric(data(i, colKol)) And IsNumeric(data(i, colCena)) Then
            vals(1) = vals(1) + CDbl(data(i, colKol)) * CDbl(data(i, colCena))
        End If
        dict(vrsta) = vals
NextPrij:
    Next i
    
    Set AggregatePrijemniceByVrstaFiltered = dict
End Function

