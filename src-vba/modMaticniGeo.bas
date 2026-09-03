Attribute VB_Name = "modMaticniGeo"
'=====================================================================
' modMaticniGeo - GEO radnje nad parcelom, PO IDENTITETU (korak M3).
'
' Sve sto legacy forma radi sa sest GEO dugmadi, ovde stoji kao operacije nad
' ParcelaID-jem, bez ijedne kontrole. Novi ekran ih zove iz svoje zone, forma
' ih zove iz svojih handlera -- jedno mesto po pravilu.
'
' ODAKLE DOLAZI (frmStammdaten, Tag = "Parcele"):
'   btnGeoOpen_Click        GeoSrbija + pretraga u klipbord
'   btnPasteCoords_Click    prepoznavanje dve koordinate iz nalepljenog teksta
'   btnGeoSave_Click        provera + SaveParcelGeoPointByID
'   btnGeoClear_Click       ClearParcelGeoByID (uz potvrdu u dva koraka)
'   btnOpenMap_Click        Lat/Lng -> Google Maps
'   btnOpenPolygonEditor    parcel-draw.html
'
' STA MODUL NE RADI: ne pise sam u tabelu. Upis i brisanje idu kroz
' modGeoParcele (SaveParcelGeoPointByID / ClearParcelGeoByID), koje rade u
' transakciji i sama racunaju Lat/Lng iz UTM34 -- iste rutine koje je forma
' zvala. Ovde je samo provera unosa, citanje po ID-ju i otvaranje adrese.
'
' ZASTO PO ID-ju, A NE PO REDU: legacy je slao m_SelectedRow, redni broj u
' tabeli izveden iz pozicije u listboxu. U mrezi koja se sortira i pretrazuje
' pozicija ne znaci nista, a modGeoParcele vec ima *ByID varijantu svake
' operacije (sa RequireSingleParcelaRow kapijom). Koristi se ona.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const MATGEO_BUILD As String = "v6-ui-191"

Private Const SRC As String = "modMaticniGeo"

' Adresa katastarskog portala. Jedno mesto -- forma i ekran otvaraju istu.
Public Const GEO_URL_SRBIJA As String = "https://a3.geosrbija.rs/"

'------------------------------------------------------------- ADRESE
' Adresa Google Maps za tacku. Decimalni zarez se pretvara u tacku: URL ne
' postuje regionalna podesavanja masine.
Public Function GeoUrlMape(ByVal lat As Double, ByVal lng As Double) As String
    GeoUrlMape = "https://www.google.com/maps?q=" & _
                 Replace(CStr(lat), ",", ".") & "," & _
                 Replace(CStr(lng), ",", ".")
End Function

Public Function GeoUrlPoligon(ByVal parcelaID As String) As String
    If Len(Trim$(parcelaID)) = 0 Then Exit Function
    GeoUrlPoligon = "https://dusanmiladinovicvnm.github.io/otkupapp-pwa/parcel-draw.html?parcelaId=" & _
                    WorksheetFunction.EncodeURL(Trim$(parcelaID))
End Function

'------------------------------------------------------------ RADNJE
' Svaka radnja vraca "" kad je proslo, inace poruku za operatera -- isti ugovor
' kao modMaticniUnos, pa pozivalac ne mora da zna sta se desilo iznutra.

' Otvara GeoSrbija i stavlja pretragu (kat. broj + opstina) u klipbord.
' Parcela bez katastarskih podataka se ODBIJA: portal bi se otvorio na praznu
' pretragu, a operater bi mislio da je nesto uradjeno.
Public Function GeoOtvoriPortal(ByVal parcelaID As String) As String
    Dim katBroj As String, katOpstina As String, pretraga As String
    On Error GoTo EH
    If Len(Trim$(parcelaID)) = 0 Then
        GeoOtvoriPortal = Poruka("MATG_ERR_NEMA_PARCELE")
        Exit Function
    End If
    katBroj = PoljeParcele(parcelaID, COL_PAR_KAT_BROJ)
    katOpstina = PoljeParcele(parcelaID, COL_PAR_KAT_OPSTINA)
    If Len(katBroj) = 0 Or Len(katOpstina) = 0 Then
        GeoOtvoriPortal = Poruka("MATG_ERR_NEMA_KATASTRA")
        Exit Function
    End If
    ' Prefiks "KO " se skida -- portal ga ne prepoznaje u pretrazi.
    pretraga = katBroj & " " & Replace(katOpstina, "KO ", "")
    CopyToClipboard pretraga
    ThisWorkbook.FollowHyperlink GEO_URL_SRBIJA
    Exit Function
EH:
    GeoOtvoriPortal = Greska("GeoOtvoriPortal")
End Function

Public Function GeoOtvoriMape(ByVal parcelaID As String) As String
    Dim lat As Double, lng As Double
    On Error GoTo EH
    If Not GeoTacka(parcelaID, lat, lng) Then
        GeoOtvoriMape = Poruka("MATG_ERR_NEMA_TACKE")
        Exit Function
    End If
    ThisWorkbook.FollowHyperlink GeoUrlMape(lat, lng)
    Exit Function
EH:
    GeoOtvoriMape = Greska("GeoOtvoriMape")
End Function

Public Function GeoOtvoriPoligon(ByVal parcelaID As String) As String
    Dim url As String
    On Error GoTo EH
    url = GeoUrlPoligon(parcelaID)
    If Len(url) = 0 Then
        GeoOtvoriPoligon = Poruka("MATG_ERR_NEMA_PARCELE")
        Exit Function
    End If
    ThisWorkbook.FollowHyperlink url
    Exit Function
EH:
    GeoOtvoriPoligon = Greska("GeoOtvoriPoligon")
End Function

' Provera i upis tacke. Koordinate moraju biti POZITIVNE: UTM34 nad Srbijom
' nema negativnih vrednosti, a nula znaci "nije uneto" -- ista provera koju
' legacy forma radi pre SaveParcelGeoPoint.
Public Function GeoSacuvaj(ByVal parcelaID As String, ByVal nTxt As String, _
                           ByVal eTxt As String, ByRef fokus As String) As String
    Dim nVal As Double, eVal As Double
    On Error GoTo EH
    fokus = ""
    If Len(Trim$(parcelaID)) = 0 Then
        GeoSacuvaj = Poruka("MATG_ERR_NEMA_PARCELE")
        Exit Function
    End If
    If Not TryParseDouble(nTxt, nVal) Then
        fokus = "n"
        GeoSacuvaj = Poruka("MATG_ERR_N")
        Exit Function
    End If
    If Not TryParseDouble(eTxt, eVal) Then
        fokus = "e"
        GeoSacuvaj = Poruka("MATG_ERR_E")
        Exit Function
    End If
    If nVal <= 0 Or eVal <= 0 Then
        fokus = "n"
        GeoSacuvaj = Poruka("MATG_ERR_POZITIVNE")
        Exit Function
    End If
    SaveParcelGeoPointByID parcelaID, nVal, eVal
    Exit Function
EH:
    GeoSacuvaj = Greska("GeoSacuvaj")
End Function

Public Function GeoObrisi(ByVal parcelaID As String) As String
    On Error GoTo EH
    If Len(Trim$(parcelaID)) = 0 Then
        GeoObrisi = Poruka("MATG_ERR_NEMA_PARCELE")
        Exit Function
    End If
    ClearParcelGeoByID parcelaID
    Exit Function
EH:
    GeoObrisi = Greska("GeoObrisi")
End Function

'------------------------------------------------------------ CITANJE
' Lat/Lng izabrane parcele. False znaci da parcela nema upotrebljivu tacku --
' i prazno i neispravno se tretiraju isto, jer se mapa ne moze otvoriti ni na
' jedno od toga.
Public Function GeoTacka(ByVal parcelaID As String, ByRef lat As Double, _
                         ByRef lng As Double) As Boolean
    On Error GoTo EH
    If Not TryParseDouble(PoljeParcele(parcelaID, COL_PAR_LAT), lat) Then Exit Function
    If Not TryParseDouble(PoljeParcele(parcelaID, COL_PAR_LNG), lng) Then Exit Function
    GeoTacka = True
    Exit Function
EH:
    GeoTacka = False
End Function

' Zatecene UTM koordinate parcele, kao TEKST za polja unosa.
Public Sub GeoKoordinate(ByVal parcelaID As String, ByRef nTxt As String, _
                         ByRef eTxt As String)
    nTxt = PoljeParcele(parcelaID, COL_PAR_N)
    eTxt = PoljeParcele(parcelaID, COL_PAR_E)
End Sub

' Jedan red opisa: status geo podatka, izvor i da li ima poligon. Ono sto je
' legacy pokazivao kroz lblGeoStatus posle svake radnje, ovde stoji stalno.
Public Function GeoOpis(ByVal parcelaID As String) As String
    Dim st As String, izv As String, poly As String, lat As Double, lng As Double
    If Len(Trim$(parcelaID)) = 0 Then
        GeoOpis = Poruka("MATG_OPIS_NEMA_IZBORA")
        Exit Function
    End If
    st = PoljeParcele(parcelaID, COL_PAR_GEO_STATUS)
    izv = PoljeParcele(parcelaID, COL_PAR_GEO_SOURCE)
    poly = PoljeParcele(parcelaID, COL_PAR_POLYGON)

    If Not GeoTacka(parcelaID, lat, lng) Then
        GeoOpis = parcelaID & "  " & ChrW(183) & "  " & Poruka("MATG_OPIS_BEZ_TACKE")
    Else
        GeoOpis = parcelaID & "  " & ChrW(183) & "  " & _
                  Format$(lat, "0.00000") & ", " & Format$(lng, "0.00000")
        If Len(st) > 0 Then GeoOpis = GeoOpis & "  " & ChrW(183) & "  " & st
        If Len(izv) > 0 Then GeoOpis = GeoOpis & " / " & izv
    End If
    If Len(poly) > 0 Then GeoOpis = GeoOpis & "  " & ChrW(183) & "  " & Poruka("MATG_OPIS_POLIGON")
End Function

'------------------------------------------------- PREPOZNAVANJE TEKSTA
' Dve koordinate iz nalepljenog teksta. Preneto iz frmStammdaten
' (TryExtractTwoCoordinates + CleanCoordToken) da bi ista pravila vazila i u
' formi i na ekranu -- forma od v6-ui-191 zove ovo.
'
' Prag |d| > 1000 je namerni filter: UTM34 koordinate nad Srbijom su
' sedmocifrene, pa se time odbacuju brojevi parcele, godine i sve ostalo sto se
' zatekne u nalepljenom redu.
Public Function GeoIzTeksta(ByVal rawText As String, ByRef prva As Double, _
                            ByRef druga As Double) As Boolean
    Dim txt As String, tokens() As String, vals(0 To 1) As Double
    Dim n As Long, i As Long, d As Double, kandidat As String
    On Error GoTo EH

    txt = Trim$(rawText)
    If txt = "" Then Exit Function

    txt = Replace(txt, vbCr, " ")
    txt = Replace(txt, vbLf, " ")
    txt = Replace(txt, vbTab, " ")
    txt = Replace(txt, ";", " ")
    Do While InStr(txt, "  ") > 0
        txt = Replace(txt, "  ", " ")
    Loop

    tokens = Split(txt, " ")
    For i = LBound(tokens) To UBound(tokens)
        kandidat = OcistiToken(tokens(i))
        If TryParseDouble(kandidat, d) Then
            If Abs(d) > 1000 Then
                vals(n) = d
                n = n + 1
                If n = 2 Then Exit For
            End If
        End If
    Next i

    If n < 2 Then Exit Function
    prva = vals(0)
    druga = vals(1)
    GeoIzTeksta = True
    Exit Function
EH:
    GeoIzTeksta = False
End Function

' Skida oznake koje portali lepe uz broj (N=, E:, zagrade).
Public Function OcistiToken(ByVal token As String) As String
    Dim s As String
    s = Trim$(token)
    s = Replace(s, "N=", "", , , vbTextCompare)
    s = Replace(s, "E=", "", , , vbTextCompare)
    s = Replace(s, "N:", "", , , vbTextCompare)
    s = Replace(s, "E:", "", , , vbTextCompare)
    s = Replace(s, "(", "")
    s = Replace(s, ")", "")
    s = Replace(s, "[", "")
    s = Replace(s, "]", "")
    s = Replace(s, "{", "")
    s = Replace(s, "}", "")
    OcistiToken = s
End Function

'---------------------------------------------------------- UNUTRASNJE
' Vrednost kolone parcele po ID-ju. Prazno kad parcele nema ili kolone nema --
' schema drift ne sme da obori radnju, isto pravilo kao na citanju liste.
Private Function PoljeParcele(ByVal parcelaID As String, ByVal kolona As String) As String
    Dim red As Long, c As Long, data As Variant
    On Error GoTo EH
    If Len(Trim$(parcelaID)) = 0 Then Exit Function
    red = modMaticniUnos.MatRedPoID("PARCELE", parcelaID)
    If red = 0 Then Exit Function
    c = GetColumnIndex(TBL_PARCELE, kolona)
    If c = 0 Then Exit Function
    data = GetTableData(TBL_PARCELE)
    If IsEmpty(data) Then Exit Function
    PoljeParcele = Trim$(NzToText(data(red, c)))
    Exit Function
EH:
    PoljeParcele = ""
End Function

' Jedan izlaz za greske: trag u logu i poruka operateru. Err se cita PRE
' ciscenja.
Private Function Greska(ByVal gde As String) As String
    Dim errDesc As String
    errDesc = Err.description
    LogError SRC & "." & gde, errDesc
    Err.Clear
    Greska = Poruka("MATG_ERR_RADNJA") & " " & errDesc
End Function
