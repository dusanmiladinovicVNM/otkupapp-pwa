Attribute VB_Name = "modMarza"

Option Explicit


' ============================================================
' modMarza v3.0 – Margenberechnung Business Logic
' Alle Funktionen geben 2D-Arrays zurück
' Spalten: Vrsta, OtkKg, ProsekOtk, IspKg, IspRSD, OtkKosten, Marza, MarzaPct
' ============================================================

Public Function ReportMarzaByKupac(ByVal kupacID As String, _
                                   ByVal datumOd As Date, _
                                   ByVal datumDo As Date) As Variant
    Dim prijData As Variant
    prijData = GetPrijemniceByKupac(kupacID, datumOd, datumDo)
    If IsEmpty(prijData) Then
        ReportMarzaByKupac = Empty
        Exit Function
    End If
    prijData = ExcludeStornirano(prijData, TBL_PRIJEMNICA)
    If IsEmpty(prijData) Then
        ReportMarzaByKupac = Empty
        Exit Function
    End If
    
    Dim otkupData As Variant
    otkupData = GetTableData(TBL_OTKUP)
    If Not IsEmpty(otkupData) Then
        otkupData = ExcludeStornirano(otkupData, TBL_OTKUP)
        Dim filters As New Collection
        Dim fp As clsFilterParam
        Set fp = New clsFilterParam
        fp.Init GetColumnIndex(TBL_OTKUP, COL_OTK_DATUM), "BETWEEN", datumOd, datumDo
        filters.Add fp
        otkupData = FilterArray(otkupData, filters)
    End If
    
    Dim vrstaCache As Object
    Set vrstaCache = BuildZbirnaVrstaCache()
    
    Dim dictPrij As Object
    Set dictPrij = AggregatePrijemniceByVrsta(prijData, vrstaCache)
    
    Dim dictOtk As Object
    Set dictOtk = AggregateOtkupByVrsta(otkupData)
    
    ReportMarzaByKupac = BuildMarzaResult(dictPrij, dictOtk)
End Function

Public Function ReportMarzaByOM(ByVal stanicaID As String, _
                                ByVal datumOd As Date, _
                                ByVal datumDo As Date) As Variant
    Dim otkupData As Variant
    otkupData = GetOtkupByStation(stanicaID, datumOd, datumDo)
    If IsEmpty(otkupData) Then
        ReportMarzaByOM = Empty
        Exit Function
    End If
    otkupData = ExcludeStornirano(otkupData, TBL_OTKUP)
    If IsEmpty(otkupData) Then
        ReportMarzaByOM = Empty
        Exit Function
    End If
    
    Dim prijData As Variant
    prijData = GetTableData(TBL_PRIJEMNICA)
    If Not IsEmpty(prijData) Then
        prijData = ExcludeStornirano(prijData, TBL_PRIJEMNICA)
        Dim filters As New Collection
        Dim fp As clsFilterParam
        Set fp = New clsFilterParam
        fp.Init GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_DATUM), "BETWEEN", datumOd, datumDo
        filters.Add fp
        prijData = FilterArray(prijData, filters)
    End If
    
    Dim vrstaCache As Object
    Set vrstaCache = BuildZbirnaVrstaCache()
    
    Dim dictOtk As Object
    Set dictOtk = AggregateOtkupByVrsta(otkupData)
    
    Dim dictPrij As Object
    Set dictPrij = AggregatePrijemniceByVrsta(prijData, vrstaCache)
    
    ReportMarzaByOM = BuildMarzaResultOM(dictOtk, dictPrij)
End Function

Public Function ReportMarzaUkupno(ByVal datumOd As Date, _
                                  ByVal datumDo As Date) As Variant
    ' Tabellen einmal laden
    Dim otkupData As Variant
    otkupData = GetTableData(TBL_OTKUP)
    If Not IsEmpty(otkupData) Then otkupData = ExcludeStornirano(otkupData, TBL_OTKUP)
    
    Dim prijData As Variant
    prijData = GetTableData(TBL_PRIJEMNICA)
    If Not IsEmpty(prijData) Then prijData = ExcludeStornirano(prijData, TBL_PRIJEMNICA)
    
    Dim vrstaCache As Object
    Set vrstaCache = BuildZbirnaVrstaCache()
    
    ' Einmal aggregieren
    Dim dictOtk As Object
    Set dictOtk = AggregateOtkupByVrstaFiltered(otkupData, datumOd, datumDo)
    
    Dim dictPrij As Object
    Set dictPrij = AggregatePrijemniceByVrstaFiltered(prijData, vrstaCache, datumOd, datumDo)
    
    ' Alle Vrsta sammeln
    Dim allVrsta As Object
    Set allVrsta = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In dictOtk.keys
        If Not allVrsta.Exists(k) Then allVrsta.Add k, True
    Next k
    For Each k In dictPrij.keys
        If Not allVrsta.Exists(k) Then allVrsta.Add k, True
    Next k
    
    If allVrsta.count = 0 Then
        ReportMarzaUkupno = Empty
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To allVrsta.count + 1, 1 To 8)
    
    Dim keys As Variant
    keys = allVrsta.keys
    Dim totalOtkKg As Double, totalOtkRSD As Double
    Dim totalIspKg As Double, totalIspRSD As Double, totalMarza As Double
    Dim idx As Long
    
    Dim i As Long
    For i = 0 To allVrsta.count - 1
        Dim otkKg As Double: otkKg = 0
        Dim otkRSD As Double: otkRSD = 0
        Dim ispKg As Double: ispKg = 0
        Dim ispRSD As Double: ispRSD = 0
        
        If dictOtk.Exists(keys(i)) Then
            Dim vO As Variant: vO = dictOtk(keys(i))
            otkKg = vO(0): otkRSD = vO(1)
        End If
        If dictPrij.Exists(keys(i)) Then
            Dim vP As Variant: vP = dictPrij(keys(i))
            ispKg = vP(0): ispRSD = vP(1)
        End If
        
        If otkKg > 0 Or ispKg > 0 Then
            idx = idx + 1
            Dim prosekOtk As Double
            If otkKg > 0 Then prosekOtk = otkRSD / otkKg Else prosekOtk = 0
            
            Dim marza As Double: marza = ispRSD - otkRSD
            Dim marzaPct As Double
            If ispRSD > 0 Then marzaPct = marza / ispRSD * 100 Else marzaPct = 0
            
            result(idx, 1) = keys(i)
            result(idx, 2) = otkKg
            result(idx, 3) = prosekOtk
            result(idx, 4) = ispKg
            result(idx, 5) = ispRSD
            result(idx, 6) = otkRSD
            result(idx, 7) = marza
            result(idx, 8) = marzaPct
            
            totalOtkKg = totalOtkKg + otkKg
            totalOtkRSD = totalOtkRSD + otkRSD
            totalIspKg = totalIspKg + ispKg
            totalIspRSD = totalIspRSD + ispRSD
            totalMarza = totalMarza + marza
        End If
    Next i
    
    If idx = 0 Then
        ReportMarzaUkupno = Empty
        Exit Function
    End If
    
    ' Array trimmen wenn idx < allVrsta.Count
    idx = idx + 1
    result(idx, 1) = "UKUPNO"
    result(idx, 2) = totalOtkKg
    result(idx, 3) = ""
    result(idx, 4) = totalIspKg
    result(idx, 5) = totalIspRSD
    result(idx, 6) = totalOtkRSD
    result(idx, 7) = totalMarza
    If totalIspRSD > 0 Then result(idx, 8) = totalMarza / totalIspRSD * 100 Else result(idx, 8) = 0
    
    ReportMarzaUkupno = result
End Function
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
' SHARED: Dict ? 2D Result (OM-Modus – umgekehrte Logik)
' ============================================================

Private Function BuildMarzaResultOM(ByVal dictEinkauf As Object, _
                                    ByVal dictVerkauf As Object) As Variant
    ' OM: EK ist Basis, VK-Durchschnitt als fiktiver Erlös
    
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

Private Function AggregateOtkupByVrsta(ByVal otkupData As Variant) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    If IsEmpty(otkupData) Then
        Set AggregateOtkupByVrsta = dict
        Exit Function
    End If
    
    Dim colVrsta As Long, colKol As Long, colCena As Long
    colVrsta = GetColumnIndex(TBL_OTKUP, COL_OTK_VRSTA)
    colKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
    colCena = GetColumnIndex(TBL_OTKUP, COL_OTK_CENA)
    
    Dim i As Long
    For i = 1 To UBound(otkupData, 1)
        Dim key As String
        key = CStr(otkupData(i, colVrsta))
        If Not dict.Exists(key) Then dict.Add key, Array(0#, 0#)
        Dim vals As Variant
        vals = dict(key)
        If IsNumeric(otkupData(i, colKol)) Then vals(0) = vals(0) + CDbl(otkupData(i, colKol))
        If IsNumeric(otkupData(i, colKol)) And IsNumeric(otkupData(i, colCena)) Then
            vals(1) = vals(1) + CDbl(otkupData(i, colKol)) * CDbl(otkupData(i, colCena))
        End If
        dict(key) = vals
    Next i
    
    Set AggregateOtkupByVrsta = dict
End Function

Private Function AggregateOtkupByVrstaFiltered(ByVal data As Variant, _
                                                ByVal datumOd As Date, _
                                                ByVal datumDo As Date) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    If IsEmpty(data) Then
        Set AggregateOtkupByVrstaFiltered = dict
        Exit Function
    End If
    
    Dim colVrsta As Long, colKol As Long, colCena As Long, colDat As Long
    colVrsta = GetColumnIndex(TBL_OTKUP, COL_OTK_VRSTA)
    colKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
    colCena = GetColumnIndex(TBL_OTKUP, COL_OTK_CENA)
    colDat = GetColumnIndex(TBL_OTKUP, COL_OTK_DATUM)
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If datumOd > 0 Then
            If IsDate(data(i, colDat)) Then
                If CDate(data(i, colDat)) < datumOd Or CDate(data(i, colDat)) > datumDo Then GoTo NextOtk
            End If
        End If
        
        Dim key As String
        key = CStr(data(i, colVrsta))
        If Not dict.Exists(key) Then dict.Add key, Array(0#, 0#)
        Dim vals As Variant: vals = dict(key)
        If IsNumeric(data(i, colKol)) Then vals(0) = vals(0) + CDbl(data(i, colKol))
        If IsNumeric(data(i, colKol)) And IsNumeric(data(i, colCena)) Then
            vals(1) = vals(1) + CDbl(data(i, colKol)) * CDbl(data(i, colCena))
        End If
        dict(key) = vals
NextOtk:
    Next i
    
    Set AggregateOtkupByVrstaFiltered = dict
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

