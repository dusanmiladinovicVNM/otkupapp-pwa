Attribute VB_Name = "modGeoParcele"
Option Explicit

Public Sub SaveParcelGeoPoint(ByVal rowIndex As Long, ByVal nCoord As Double, ByVal eCoord As Double)
    
    Dim lat As Double
    Dim lng As Double
    
    ConvertUTM34ToLatLng eCoord, nCoord, lat, lng
    
    UpdateCell TBL_PARCELE, rowIndex, "N_Coord", CDbl(nCoord)
    UpdateCell TBL_PARCELE, rowIndex, "E_Coord", CDbl(eCoord)
    UpdateCell TBL_PARCELE, rowIndex, "Lat", CDbl(Round(lat, 6))
    UpdateCell TBL_PARCELE, rowIndex, "Lng", CDbl(Round(lng, 6))
    UpdateCell TBL_PARCELE, rowIndex, "GeoStatus", "point"
    UpdateCell TBL_PARCELE, rowIndex, "GeoSource", "selenium"
    UpdateCell TBL_PARCELE, rowIndex, "MeteoEnabled", "Da"
    UpdateCell TBL_PARCELE, rowIndex, "DatumGeoUnosa", Now
    UpdateCell TBL_PARCELE, rowIndex, "DatumAzuriranja", Now
    
End Sub

Public Sub ClearParcelGeo(ByVal rowIndex As Long)

    UpdateCell TBL_PARCELE, rowIndex, "N_Coord", ""
    UpdateCell TBL_PARCELE, rowIndex, "E_Coord", ""
    UpdateCell TBL_PARCELE, rowIndex, "Lat", ""
    UpdateCell TBL_PARCELE, rowIndex, "Lng", ""
    UpdateCell TBL_PARCELE, rowIndex, "GeoStatus", "none"
    UpdateCell TBL_PARCELE, rowIndex, "GeoSource", ""
    UpdateCell TBL_PARCELE, rowIndex, "MeteoEnabled", "Ne"
    UpdateCell TBL_PARCELE, rowIndex, "DatumAzuriranja", Now

End Sub

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
    Dim N1 As Double, T1 As Double, C1 As Double, R1 As Double, D As Double
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
    D = X / (N1 * k0)
    
    latRad = phi1Rad - (N1 * Tan(phi1Rad) / R1) * _
        (D ^ 2 / 2# _
        - (5# + 3# * T1 + 10# * C1 - 4# * C1 ^ 2 - 9# * eccPrimeSquared) * D ^ 4 / 24# _
        + (61# + 90# * T1 + 298# * C1 + 45# * T1 ^ 2 - 252# * eccPrimeSquared - 3# * C1 ^ 2) * D ^ 6 / 720#)
    
    lonOrigin = (zoneNumber - 1#) * 6# - 180# + 3#   ' 21
    
    lonRad = (D _
        - (1# + 2# * T1 + C1) * D ^ 3 / 6# _
        + (5# - 2# * C1 + 28# * T1 - 3# * C1 ^ 2 + 8# * eccPrimeSquared + 24# * T1 ^ 2) * D ^ 5 / 120#) / Cos(phi1Rad)
    
    lat = latRad * 180# / 3.14159265358979
    lng = lonOrigin + lonRad * 180# / 3.14159265358979

End Sub
