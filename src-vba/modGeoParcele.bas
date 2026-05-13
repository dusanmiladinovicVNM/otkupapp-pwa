Attribute VB_Name = "modGeoParcele"
Option Explicit

Public Sub SaveParcelGeoPoint(ByVal rowIndex As Long, ByVal nCoord As Double, ByVal eCoord As Double)
    If Not SaveParcelGeoPoint_TX(rowIndex, nCoord, eCoord) Then
        Err.Raise vbObjectError + 7601, "modGeoParcele.SaveParcelGeoPoint", _
                  "SaveParcelGeoPoint_TX nije uspeo."
    End If
End Sub

Public Function SaveParcelGeoPoint_TX(ByVal rowIndex As Long, _
                                      ByVal nCoord As Double, _
                                      ByVal eCoord As Double) As Boolean
    Const SRC As String = "modGeoParcele.SaveParcelGeoPoint_TX"

    On Error GoTo EH

    If rowIndex <= 0 Then
        Err.Raise vbObjectError + 7602, SRC, "rowIndex nije validan."
    End If

    If nCoord <= 0 Or eCoord <= 0 Then
        Err.Raise vbObjectError + 7603, SRC, "Koordinate moraju biti pozitivne."
    End If

    Dim tx As clsTransaction
    Set tx = New clsTransaction

    Dim lat As Double
    Dim lng As Double

    ConvertUTM34ToLatLng eCoord, nCoord, lat, lng

    tx.BeginTx
    tx.AddTableSnapshot TBL_PARCELE

    RequireUpdateCell TBL_PARCELE, rowIndex, COL_PAR_N, CDbl(nCoord), SRC
    RequireUpdateCell TBL_PARCELE, rowIndex, COL_PAR_E, CDbl(eCoord), SRC
    RequireUpdateCell TBL_PARCELE, rowIndex, COL_PAR_LAT, CDbl(Round(lat, 6)), SRC
    RequireUpdateCell TBL_PARCELE, rowIndex, COL_PAR_LNG, CDbl(Round(lng, 6)), SRC
    RequireUpdateCell TBL_PARCELE, rowIndex, COL_PAR_GEO_STATUS, "point", SRC
    RequireUpdateCell TBL_PARCELE, rowIndex, COL_PAR_GEO_SOURCE, "selenium", SRC
    RequireUpdateCell TBL_PARCELE, rowIndex, COL_PAR_METEO, "Da", SRC
    RequireUpdateCell TBL_PARCELE, rowIndex, COL_PAR_DATUM_GEO, Now, SRC
    RequireUpdateCell TBL_PARCELE, rowIndex, COL_PAR_DATUM_AZUR, Now, SRC

    tx.CommitTx

    SaveParcelGeoPoint_TX = True
    Exit Function

EH:
    LogErr SRC

    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    SaveParcelGeoPoint_TX = False
End Function

Public Sub ClearParcelGeo(ByVal rowIndex As Long)
    If Not ClearParcelGeo_TX(rowIndex) Then
        Err.Raise vbObjectError + 7611, "modGeoParcele.ClearParcelGeo", _
                  "ClearParcelGeo_TX nije uspeo."
    End If
End Sub

Public Function ClearParcelGeo_TX(ByVal rowIndex As Long) As Boolean
    Const SRC As String = "modGeoParcele.ClearParcelGeo_TX"

    On Error GoTo EH

    If rowIndex <= 0 Then
        Err.Raise vbObjectError + 7612, SRC, "rowIndex nije validan."
    End If

    Dim tx As clsTransaction
    Set tx = New clsTransaction

    tx.BeginTx
    tx.AddTableSnapshot TBL_PARCELE

    RequireUpdateCell TBL_PARCELE, rowIndex, COL_PAR_N, "", SRC
    RequireUpdateCell TBL_PARCELE, rowIndex, COL_PAR_E, "", SRC
    RequireUpdateCell TBL_PARCELE, rowIndex, COL_PAR_LAT, "", SRC
    RequireUpdateCell TBL_PARCELE, rowIndex, COL_PAR_LNG, "", SRC
    RequireUpdateCell TBL_PARCELE, rowIndex, COL_PAR_GEO_STATUS, "none", SRC
    RequireUpdateCell TBL_PARCELE, rowIndex, COL_PAR_GEO_SOURCE, "", SRC
    RequireUpdateCell TBL_PARCELE, rowIndex, COL_PAR_METEO, "Ne", SRC
    RequireUpdateCell TBL_PARCELE, rowIndex, COL_PAR_DATUM_AZUR, Now, SRC

    tx.CommitTx

    ClearParcelGeo_TX = True
    Exit Function

EH:
    LogErr SRC

    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    ClearParcelGeo_TX = False
End Function

Public Sub ConvertUTM34ToLatLng(ByVal eCoord As Double, ByVal nCoord As Double, _
                                ByRef lat As Double, ByRef lng As Double)

    Const a As Double = 6378137#
    Const eccSquared As Double = 0.00669437999014
    Const k0 As Double = 0.9996
    
    Dim eccPrimeSquared As Double
    Dim e1 As Double
    
    Dim X As Double, Y As Double
    Dim m As Double, mu As Double
    Dim phi1Rad As Double
    Dim N1 As Double, T1 As Double, C1 As Double, R1 As Double, d As Double
    Dim zoneNumber As Long
    Dim lonOrigin As Double
    Dim latRad As Double, lonRad As Double
    
    zoneNumber = 34
    
    X = eCoord - 500000#
    Y = nCoord
    
    eccPrimeSquared = eccSquared / (1# - eccSquared)
    
    m = Y / k0
    mu = m / (a * (1# - eccSquared / 4# - 3# * eccSquared ^ 2 / 64# - 5# * eccSquared ^ 3 / 256#))
    
    e1 = (1# - Sqr(1# - eccSquared)) / (1# + Sqr(1# - eccSquared))
    
    phi1Rad = mu _
        + (3# * e1 / 2# - 27# * e1 ^ 3 / 32#) * Sin(2# * mu) _
        + (21# * e1 ^ 2 / 16# - 55# * e1 ^ 4 / 32#) * Sin(4# * mu) _
        + (151# * e1 ^ 3 / 96#) * Sin(6# * mu) _
        + (1097# * e1 ^ 4 / 512#) * Sin(8# * mu)
    
    N1 = a / Sqr(1# - eccSquared * Sin(phi1Rad) ^ 2)
    T1 = Tan(phi1Rad) ^ 2
    C1 = eccPrimeSquared * Cos(phi1Rad) ^ 2
    R1 = a * (1# - eccSquared) / (1# - eccSquared * Sin(phi1Rad) ^ 2) ^ 1.5
    d = X / (N1 * k0)
    
    latRad = phi1Rad - (N1 * Tan(phi1Rad) / R1) * _
        (d ^ 2 / 2# _
        - (5# + 3# * T1 + 10# * C1 - 4# * C1 ^ 2 - 9# * eccPrimeSquared) * d ^ 4 / 24# _
        + (61# + 90# * T1 + 298# * C1 + 45# * T1 ^ 2 - 252# * eccPrimeSquared - 3# * C1 ^ 2) * d ^ 6 / 720#)
    
    lonOrigin = (zoneNumber - 1#) * 6# - 180# + 3#   ' 21
    
    lonRad = (d _
        - (1# + 2# * T1 + C1) * d ^ 3 / 6# _
        + (5# - 2# * C1 + 28# * T1 - 3# * C1 ^ 2 + 8# * eccPrimeSquared + 24# * T1 ^ 2) * d ^ 5 / 120#) / Cos(phi1Rad)
    
    lat = latRad * 180# / 3.14159265358979
    lng = lonOrigin + lonRad * 180# / 3.14159265358979

End Sub


Public Function SyncSelectedParcelaToGoogle(ByVal parcelaID As String) As Boolean
    Const SRC As String = "modGeoParcele.SyncSelectedParcelaToGoogle"

    On Error GoTo EH

    Dim cleanID As String
    cleanID = Trim$(parcelaID)

    SyncSelectedParcelaToGoogle = False

    If Len(cleanID) = 0 Then
        LogError SRC, "ParcelaID je prazan."
        Exit Function
    End If

    If Not LocalParcelaExists(cleanID) Then
        LogError SRC, "Parcela ne postoji u lokalnom tblParcele: " & cleanID
        Exit Function
    End If

    ' Reuse postojeci canonical Parcele export.
    ' Ne pravimo paralelni header mapping u modGeoParcele.
    If Not SyncParceleToGoogle_Core(False) Then
        LogError SRC, "SyncParceleToGoogle_Core nije uspeo za parcelu: " & cleanID
        Exit Function
    End If

    ' Safety check: editor cita Google, zato proveravamo da je red zaista tamo.
    If Not GoogleParcelaExists(cleanID) Then
        LogError SRC, "Parcela nije pronadjena u Google Parcele tabu posle sync-a: " & cleanID
        Exit Function
    End If

    LogInfo SRC, "Parcela spremna za editor: " & cleanID
    SyncSelectedParcelaToGoogle = True
    Exit Function

EH:
    LogErr SRC
    SyncSelectedParcelaToGoogle = False
End Function

Private Function LocalParcelaExists(ByVal parcelaID As String) As Boolean
    Const SRC As String = "modGeoParcele.LocalParcelaExists"

    On Error GoTo EH

    Dim data As Variant
    Dim idCol As Long
    Dim i As Long

    data = GetTableData(TBL_PARCELE)
    If IsEmpty(data) Then Exit Function

    idCol = GetColumnIndex(TBL_PARCELE, COL_PAR_ID)
    If idCol = 0 Then
        LogError SRC, "Nedostaje kolona: " & COL_PAR_ID
        Exit Function
    End If

    For i = 1 To UBound(data, 1)
        If Trim$(NzToText(data(i, idCol))) = Trim$(parcelaID) Then
            LocalParcelaExists = True
            Exit Function
        End If
    Next i

    Exit Function

EH:
    LogErr SRC
    LocalParcelaExists = False
End Function

Private Function GoogleParcelaExists(ByVal parcelaID As String) As Boolean
    Const SRC As String = "modGeoParcele.GoogleParcelaExists"

    On Error GoTo EH

    Dim sheetID As String
    Dim data As Variant
    Dim idCol As Long
    Dim r As Long

    sheetID = GetConfigValue("GOOGLE_STAMMDATEN_SHEET_ID")

    If Len(Trim$(sheetID)) = 0 Then
        LogError SRC, "GOOGLE_STAMMDATEN_SHEET_ID nije postavljen posle sync-a."
        Exit Function
    End If

    data = ReadSheetData(sheetID, "Parcele")
    If IsEmpty(data) Then
        LogError SRC, "Google Parcele tab je prazan ili nije citljiv."
        Exit Function
    End If

    idCol = FindHeaderColumnInArray(data, COL_PAR_ID)
    If idCol = 0 Then
        LogError SRC, "Google Parcele tab nema header: " & COL_PAR_ID
        Exit Function
    End If

    For r = 2 To UBound(data, 1)
        If Trim$(SafeArrayText(data, r, idCol)) = Trim$(parcelaID) Then
            GoogleParcelaExists = True
            Exit Function
        End If
    Next r

    Exit Function

EH:
    LogErr SRC
    GoogleParcelaExists = False
End Function

Private Function FindHeaderColumnInArray(ByVal data As Variant, ByVal headerName As String) As Long
    On Error GoTo EH

    Dim c As Long

    If IsEmpty(data) Then Exit Function

    For c = 1 To UBound(data, 2)
        If StrComp(Trim$(SafeArrayText(data, 1, c)), Trim$(headerName), vbTextCompare) = 0 Then
            FindHeaderColumnInArray = c
            Exit Function
        End If
    Next c

    Exit Function

EH:
    FindHeaderColumnInArray = 0
End Function

Private Function SafeArrayText(ByVal data As Variant, ByVal rowIndex As Long, ByVal colIndex As Long) As String
    On Error GoTo EH

    If IsEmpty(data) Then Exit Function
    If rowIndex < 1 Or colIndex < 1 Then Exit Function
    If rowIndex > UBound(data, 1) Then Exit Function
    If colIndex > UBound(data, 2) Then Exit Function

    If isError(data(rowIndex, colIndex)) Then
        SafeArrayText = ""
    Else
        SafeArrayText = Trim$(NzToText(data(rowIndex, colIndex)))
    End If

    Exit Function

EH:
    SafeArrayText = ""
End Function
