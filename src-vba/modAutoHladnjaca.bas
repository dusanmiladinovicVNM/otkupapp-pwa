Attribute VB_Name = "modAutoHladnjaca"
Option Explicit

' ============================================================
' modAutoHladnjaca (#3)
'
' Kada je otkupno mesto (stanica) HLADNJACA (tblStanice.JeHladnjaca = "Da")
' i toggle AUTO_PRIJEMNICA_HLADNJACA je ukljucen, iz otkupa se automatski
' kreira ceo lanac: otpremnica + zbirna + prijemnica.
'
' Kupac (hladnjaca) = MALINA_DEFAULT_KUPAC (KupacID iz tblKupci).
' Cena = cena iz otkupa. Svi ostali podaci dolaze iz otkupa.
'
' Broj prijemnice: "1/ddmmyy" za prvu tog dana, pa "1/ddmmyy-2", "-3" ... -n.
' Broj otpremnice = broj zbirne = broj otkupnog dokumenta (malina konvencija);
' ako otkup nema broj, generise se HL-ddmmyy-hhnnss.
'
' Reuse: SaveOtpremnica_TX / SaveZbirna_TX / SavePrijemnica_TX (modDokumenta),
'        LookupValue / GetConfigValue, IsAutoPrijemnicaHladnjaca (modConfig).
' ============================================================

' Da li je stanica oznacena kao hladnjaca (tblStanice.JeHladnjaca = "Da").
Public Function IsHladnjacaStanica(ByVal stanicaID As String) As Boolean
    On Error Resume Next
    Dim v As String
    v = Trim$(Nz(LookupValue(TBL_STANICE, "StanicaID", stanicaID, COL_STA_JE_HLADNJACA), ""))
    IsHladnjacaStanica = (StrComp(v, "Da", vbTextCompare) = 0)
End Function

' Sledeci broj prijemnice po obrascu 1/ddmmyy [-k].
Public Function NextBrojPrijemnice(ByVal datum As Date) As String
    On Error GoTo EH

    Dim baseNum As String
    baseNum = "1/" & Format$(datum, "ddmmyy")

    Dim cnt As Long
    cnt = 0

    Dim data As Variant
    data = GetTableData(TBL_PRIJEMNICA)
    If Not IsEmpty(data) Then
        Dim cBroj As Long
        cBroj = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ)
        If cBroj > 0 Then
            Dim i As Long, b As String
            For i = 1 To UBound(data, 1)
                b = Trim$(Nz(data(i, cBroj), ""))
                If b = baseNum Or Left$(b, Len(baseNum) + 1) = baseNum & "-" Then cnt = cnt + 1
            Next i
        End If
    End If

    If cnt = 0 Then
        NextBrojPrijemnice = baseNum
    Else
        NextBrojPrijemnice = baseNum & "-" & CStr(cnt + 1)
    End If
    Exit Function

EH:
    NextBrojPrijemnice = "1/" & Format$(datum, "ddmmyy")
End Function

' Auto-lanac za hladnjacu. Poziva se posle uspesnog SaveOtkupMulti_TX (frmOtkup).
' Best-effort: greska NE sme da obori potvrdu otkupa.
Public Sub AutoChainHladnjaca(ByVal datum As Date, ByVal stanicaID As String, _
                              ByVal vrsta As String, ByVal sorta As String, _
                              ByVal vozacID As String, ByVal tipAmb As String, _
                              ByVal kolAmb As Long, _
                              ByVal kolicinaI As Double, ByVal cenaI As Double, _
                              ByVal hasKlasaII As Boolean, _
                              ByVal kolicinaII As Double, ByVal cenaII As Double, _
                              ByVal brDok As String)
    On Error GoTo EH

    If Not IsAutoPrijemnicaHladnjaca() Then Exit Sub
    If Not IsHladnjacaStanica(stanicaID) Then Exit Sub

    Dim kupacID As String
    kupacID = Trim$(GetConfigValue(CFG_MALINA_DEFAULT_KUPAC))
    If Len(kupacID) = 0 Then
        ' Toggle ukljucen ali kupac-hladnjaca nije podesen -> ne mozemo dalje.
        LogError "modAutoHladnjaca.AutoChainHladnjaca", _
                 "AUTO_PRIJEMNICA_HLADNJACA ukljucen ali MALINA_DEFAULT_KUPAC prazan."
        Exit Sub
    End If

    Dim hladnjaca As String
    hladnjaca = Trim$(Nz(LookupValue(TBL_KUPCI, COL_KUP_ID, kupacID, "Hladnjaca"), ""))

    ' Broj otpremnice = broj zbirne = broj otkupa (ili generisan).
    Dim brOtp As String
    brOtp = Trim$(brDok)
    If Len(brOtp) = 0 Then brOtp = "HL-" & Format$(datum, "ddmmyy") & "-" & Format$(Now, "hhnnss")
    Dim brZbr As String
    brZbr = brOtp

    ' Klasa I (ambalaza se broji samo ovde, da se ne duplira).
    Dim brPrij As String
    brPrij = NextBrojPrijemnice(datum)
    SaveOtpremnica_TX datum, stanicaID, vozacID, brOtp, brZbr, vrsta, sorta, _
                      kolicinaI, cenaI, tipAmb, kolAmb, KLASA_I
    SaveZbirna_TX datum, vozacID, brZbr, kupacID, hladnjaca, "", vrsta, sorta, _
                  kolicinaI, tipAmb, kolAmb, KLASA_I
    SavePrijemnica_TX datum, kupacID, vozacID, brPrij, brZbr, vrsta, sorta, _
                      kolicinaI, cenaI, tipAmb, kolAmb, 0, KLASA_I

    ' Klasa II (bez ambalaze; vec izlazna na Klasa I).
    If hasKlasaII And kolicinaII > 0 Then
        Dim brPrij2 As String
        brPrij2 = NextBrojPrijemnice(datum)
        SaveOtpremnica_TX datum, stanicaID, vozacID, brOtp, brZbr, vrsta, sorta, _
                          kolicinaII, cenaII, tipAmb, 0, KLASA_II
        SaveZbirna_TX datum, vozacID, brZbr, kupacID, hladnjaca, "", vrsta, sorta, _
                      kolicinaII, tipAmb, 0, KLASA_II
        SavePrijemnica_TX datum, kupacID, vozacID, brPrij2, brZbr, vrsta, sorta, _
                          kolicinaII, cenaII, tipAmb, 0, 0, KLASA_II
    End If
    Exit Sub

EH:
    LogErr "modAutoHladnjaca.AutoChainHladnjaca"
End Sub
