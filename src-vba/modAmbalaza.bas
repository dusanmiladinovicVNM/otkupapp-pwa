Attribute VB_Name = "modAmbalaza"
Option Explicit
' ============================================================
' modAmbalaza v2.1 – Verpackung-Tracking
' NEU: VozacID-Parameter
' ============================================================

Public Sub TrackAmbalaza(ByVal datum As Date, ByVal tipAmb As String, _
                         ByVal kolicina As Long, ByVal smer As String, _
                         ByVal entitetID As String, ByVal entitetTip As String, _
                         Optional ByVal vozacID As String = "", _
                         Optional ByVal dokumentID As String = "", _
                         Optional ByVal dokumentTip As String = "")
    If kolicina = 0 Then Exit Sub
    
    Dim newID As String
    newID = GetNextID(TBL_AMBALAZA, COL_AMB_ID, "AMB-")
    
    Dim rowData As Variant
    rowData = Array(newID, datum, tipAmb, kolicina, smer, entitetID, entitetTip, _
                    vozacID, dokumentID, dokumentTip)
    AppendRow TBL_AMBALAZA, rowData
End Sub

Public Function GetAmbalazeStanje(ByVal entitetID As String, _
                                  ByVal entitetTip As String) As Variant
    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then
        GetAmbalazeStanje = Empty
        Exit Function
    End If
    data = ExcludeStornirano(data, TBL_AMBALAZA)
        If IsEmpty(data) Then
        GetAmbalazeStanje = Empty
        Exit Function
    End If
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim colTip As Long, colKol As Long, colSmer As Long
    Dim colEntID As Long, colEntTip As Long
    colTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_TIP)
    colKol = GetColumnIndex(TBL_AMBALAZA, COL_AMB_KOLICINA)
    colSmer = GetColumnIndex(TBL_AMBALAZA, COL_AMB_SMER)
    colEntID = GetColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET)
    colEntTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET_TIP)
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colEntID)) = entitetID And _
           CStr(data(i, colEntTip)) = entitetTip Then
            Dim key As String
            key = CStr(data(i, colTip))
            If Not dict.Exists(key) Then dict.Add key, 0
            If CStr(data(i, colSmer)) = "Ulaz" Then
                dict(key) = dict(key) + CLng(data(i, colKol))
            Else
                dict(key) = dict(key) - CLng(data(i, colKol))
            End If
        End If
    Next i
    
    If dict.count = 0 Then
        GetAmbalazeStanje = Empty
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To dict.count, 1 To 2)
    Dim keys As Variant
    keys = dict.keys
    For i = 0 To dict.count - 1
        result(i + 1, 1) = keys(i)
        result(i + 1, 2) = dict(keys(i))
    Next i
    
    GetAmbalazeStanje = result
End Function

Public Function GetVozacAmbSaldo(ByVal vozacID As String, _
                                  Optional ByVal datumOd As Date, _
                                  Optional ByVal datumDo As Date) As Variant
    ' Ambalaza-Saldo pro Vozac: Izlaz - Ulaz
    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then
        GetVozacAmbSaldo = Empty
        Exit Function
    End If
    data = ExcludeStornirano(data, TBL_AMBALAZA)  ' ? NEU
    If IsEmpty(data) Then
        GetVozacAmbSaldo = Empty
        Exit Function
    End If
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim colTip As Long, colKol As Long, colSmer As Long
    Dim colVozac As Long, colDatum As Long
    colTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_TIP)
    colKol = GetColumnIndex(TBL_AMBALAZA, COL_AMB_KOLICINA)
    colSmer = GetColumnIndex(TBL_AMBALAZA, COL_AMB_SMER)
    colVozac = GetColumnIndex(TBL_AMBALAZA, COL_AMB_VOZAC)
    colDatum = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DATUM)
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colVozac)) = vozacID Then
            ' Datum-Filter wenn angegeben
            If datumOd <> 0 Then
                Dim D As Date
                D = CDate(data(i, colDatum))
                If D < datumOd Or D > datumDo Then GoTo NextRow
            End If
            
            Dim key As String
            key = CStr(data(i, colTip))
            If Not dict.Exists(key) Then dict.Add key, Array(0, 0) ' Izlaz, Ulaz
            Dim vals As Variant
            vals = dict(key)
            If CStr(data(i, colSmer)) = "Izlaz" Then
                vals(0) = vals(0) + CLng(data(i, colKol))
            Else
                vals(1) = vals(1) + CLng(data(i, colKol))
            End If
            dict(key) = vals
        End If
NextRow:
    Next i
    
    If dict.count = 0 Then
        GetVozacAmbSaldo = Empty
        Exit Function
    End If
    
    ' Array(TipAmb, Izlaz, Ulaz, Saldo)
    Dim result() As Variant
    ReDim result(1 To dict.count, 1 To 4)
    Dim keys As Variant
    keys = dict.keys
    For i = 0 To dict.count - 1
        vals = dict(keys(i))
        result(i + 1, 1) = keys(i)
        result(i + 1, 2) = vals(0)
        result(i + 1, 3) = vals(1)
        result(i + 1, 4) = vals(0) - vals(1)
    Next i
    
    GetVozacAmbSaldo = result
End Function

