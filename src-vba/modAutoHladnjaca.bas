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
' Posle kreiranja, izvorni tblOtkup red dobija nazad: OtpremnicaID + BrojZbirne +
' VozacID (LinkOtkupRedNaDokument), da otkup ostane povezan sa dokumentima.
'
' Reuse: SaveOtpremnica_TX / SaveZbirna_TX / SavePrijemnica_TX (modDokumenta),
'        FindRows / RequireUpdateCell / clsTransaction (vezivanje nazad u otkup),
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
                              ByVal brDok As String, _
                              ByVal otkupIDs As String, _
                              Optional ByVal brutoKgI As Double = 0, _
                              Optional ByVal kolAmbII As Long = 0, _
                              Optional ByVal brutoKgII As Double = 0)
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

    ' Vozac je obavezan na otpremnici/zbirnoj/prijemnici. U malina/hladnjaca
    ' konvenciji par-vozac ima VozacID == StanicaID. Ako otkup nema vozaca,
    ' koristimo stanicu kao vozaca (i napravimo par-vozaca ako je malina mod).
    If Len(Trim$(vozacID)) = 0 Then
        On Error Resume Next
        EnsureVozacMirrorForStanica stanicaID, _
            Trim$(Nz(LookupValue(TBL_STANICE, "StanicaID", stanicaID, "Naziv"), "")), _
            "(hladnjaca)", ""
        On Error GoTo EH
        vozacID = stanicaID
    End If

    ' Broj otpremnice = broj otkupa (ili generisan). Zbirna se razdvaja: mirror-
    ' stanica (vozac==stanica) nosi "S" prefiks (S1/ddmmyy), otpremnica/otkup
    ' zadrzavaju svoj broj (1/ddmmyy) da se ne sudaraju sa realnim vozacem.
    Dim brOtp As String
    brOtp = Trim$(brDok)
    If Len(brOtp) = 0 Then brOtp = "HL-" & Format$(datum, "ddmmyy") & "-" & Format$(Now, "hhnnss")
    Dim brZbr As String
    brZbr = ApplyMirrorPrefix(vozacID, brOtp)

    ' Klasa I je opciona (kolicinaI = 0 -> kroz lanac ide samo Klasa II).
    Dim hasKlasaI As Boolean: hasKlasaI = (kolicinaI > 0)

    ' OtkupID-jevi za vezivanje nazad u tblOtkup. Format iz SaveOtkupMulti_TX:
    ' "resultI", "resultI + resultII", ili (samo II klasa) "resultII".
    Dim idI As String, idII As String
    If Len(Trim$(otkupIDs)) > 0 Then
        Dim parts() As String
        parts = Split(otkupIDs, " + ")
        If hasKlasaI Then
            idI = Trim$(parts(LBound(parts)))
            If UBound(parts) > LBound(parts) Then idII = Trim$(parts(LBound(parts) + 1))
        Else
            idII = Trim$(parts(LBound(parts)))
        End If
    End If

    ' Klasa I (svoja kolicina ambalaze = kolAmb). Preskace se ako se unosi samo II.
    If hasKlasaI Then
        Dim brPrij As String
        brPrij = NextBrojPrijemnice(datum)
        Dim otpID As String
        otpID = SaveOtpremnica_TX(datum, stanicaID, vozacID, brOtp, brZbr, vrsta, sorta, _
                                  kolicinaI, cenaI, tipAmb, kolAmb, KLASA_I)
        SaveZbirna_TX datum, vozacID, brZbr, kupacID, hladnjaca, "", vrsta, sorta, _
                      kolicinaI, tipAmb, kolAmb, KLASA_I
        SavePrijemnica_TX datum, kupacID, vozacID, brPrij, brZbr, vrsta, sorta, _
                          kolicinaI, cenaI, tipAmb, kolAmb, 0, KLASA_I, brutoKgI
        ' Veza nazad u otkup red: OtpremnicaID + BrojZbirne + VozacID.
        LinkOtkupRedNaDokument idI, otpID, brZbr, vozacID
    End If

    ' Klasa II: zasebne gajbe (kolAmbII) -> ide kroz ceo lanac kao Klasa I
    ' (otpremnica/zbirna/prijemnica + paletizacija). Ranije se ambalaza vodila samo
    ' na Klasi I; sada Klasa II ima svoju kolicinu ambalaze.
    If hasKlasaII And kolicinaII > 0 Then
        Dim brPrij2 As String
        brPrij2 = NextBrojPrijemnice(datum)
        Dim otpID2 As String
        otpID2 = SaveOtpremnica_TX(datum, stanicaID, vozacID, brOtp, brZbr, vrsta, sorta, _
                                   kolicinaII, cenaII, tipAmb, kolAmbII, KLASA_II)
        SaveZbirna_TX datum, vozacID, brZbr, kupacID, hladnjaca, "", vrsta, sorta, _
                      kolicinaII, tipAmb, kolAmbII, KLASA_II
        SavePrijemnica_TX datum, kupacID, vozacID, brPrij2, brZbr, vrsta, sorta, _
                          kolicinaII, cenaII, tipAmb, kolAmbII, 0, KLASA_II, brutoKgII
        LinkOtkupRedNaDokument idII, otpID2, brZbr, vozacID
    End If
    Exit Sub

EH:
    LogErr "modAutoHladnjaca.AutoChainHladnjaca"
End Sub

' Upisi vezu nazad u tblOtkup red(ove) za dati OtkupID:
'   OtpremnicaID + BrojZbirne (tek kreirani -> upisuju se),
'   VozacID samo ako je prazan (da se ne pregazi operaterov izbor).
' Reuse postojecih primitiva: FindRows / RequireUpdateCell / clsTransaction.
Private Sub LinkOtkupRedNaDokument(ByVal otkupID As String, ByVal otpID As String, _
                                   ByVal brZbr As String, ByVal vozacID As String)
    Const SRC As String = "modAutoHladnjaca.LinkOtkupRedNaDokument"
    Dim tx As clsTransaction
    On Error GoTo EH

    otkupID = Trim$(otkupID)
    If Len(otkupID) = 0 Then Exit Sub

    Dim rows As Collection
    Set rows = FindRows(TBL_OTKUP, COL_OTK_ID, otkupID)
    If rows Is Nothing Then Exit Sub
    If rows.count = 0 Then Exit Sub

    Dim curVoz As String
    curVoz = Trim$(Nz(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_VOZAC), ""))

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP

    Dim k As Long, r As Long
    For k = 1 To rows.count
        r = rows(k)
        If Len(otpID) > 0 Then RequireUpdateCell TBL_OTKUP, r, COL_OTK_OTPREMNICA_ID, otpID, SRC
        If Len(brZbr) > 0 Then RequireUpdateCell TBL_OTKUP, r, COL_OTK_BROJ_ZBIRNE, brZbr, SRC
        If Len(vozacID) > 0 And curVoz = "" Then _
            RequireUpdateCell TBL_OTKUP, r, COL_OTK_VOZAC, vozacID, SRC
    Next k

    tx.CommitTx
    Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr SRC
End Sub
