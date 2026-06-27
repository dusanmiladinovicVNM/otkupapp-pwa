Attribute VB_Name = "modAutoHladnjaca"
'Attribute VB_Name = "modAutoHladnjaca"
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
' Broj prijemnice (modBrojevi.GenerateBrojPrijemnice): "1/ddmmyy" za prvu tog
' dana, pa "1/ddmmyy-2", "-3" ... -n. Dvoklasna prijemnica (Kl I + Kl II) nosi
' ISTI broj -- jedna prijemnica = jedan broj (kao i kod rucnog unosa).
' Broj otpremnice = broj zbirne = broj otkupnog dokumenta (malina konvencija);
' ako otkup nema broj, generise se HL-ddmmyy-hhnnss.
'
' Posle kreiranja, izvorni tblOtkup red dobija nazad: OtpremnicaID + BrojZbirne +
' VozacID (LinkOtkupRedNaDokument), da otkup ostane povezan sa dokumentima.
'
' Reuse: SaveOtpremnica_TX / SaveZbirna_TX / SavePrijemnica_TX (modDokumenta),
'        GenerateBrojPrijemnice (modBrojevi), FindRows / RequireUpdateCell /
'        clsTransaction (vezivanje nazad u otkup), LookupValue / GetConfigValue,
'        IsAutoPrijemnicaHladnjaca (modConfig).
' ============================================================

' Da li je stanica oznacena kao hladnjaca (tblStanice.JeHladnjaca = "Da").
Public Function IsHladnjacaStanica(ByVal stanicaID As String) As Boolean
    On Error Resume Next
    Dim v As String
    v = Trim$(nz(LookupValue(TBL_STANICE, "StanicaID", stanicaID, COL_STA_JE_HLADNJACA), ""))
    IsHladnjacaStanica = (StrComp(v, "Da", vbTextCompare) = 0)
End Function

' Auto-lanac za hladnjacu. Poziva se posle uspesnog SaveOtkupMulti_TX (frmOtkup).
' Best-effort: greska NE sme da obori potvrdu otkupa. Vraca "" kad je lanac
' kompletan; inace tekst upozorenja (frmOtkup ga prikaze operateru).
Public Function AutoChainHladnjaca(ByVal datum As Date, ByVal stanicaID As String, _
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
                              Optional ByVal brutoKgII As Double = 0) As String
    On Error GoTo EH

    ' Vraca "" kad je lanac kompletan; inace tekst upozorenja za frmOtkup
    ' (npr. prijemnica nije kreirana -> lanac nepotpun).
    Dim failKlase As String

    If Not IsAutoPrijemnicaHladnjaca() Then Exit Function
    If Not IsHladnjacaStanica(stanicaID) Then Exit Function

    Dim kupacID As String
    kupacID = Trim$(GetConfigValue(CFG_MALINA_DEFAULT_KUPAC))
    If Len(kupacID) = 0 Then
        ' Toggle ukljucen ali kupac-hladnjaca nije podesen -> ne mozemo dalje.
        LogError "modAutoHladnjaca.AutoChainHladnjaca", _
                 "AUTO_PRIJEMNICA_HLADNJACA ukljucen ali MALINA_DEFAULT_KUPAC prazan."
        AutoChainHladnjaca = "Auto-lanac hladnjace nije pokrenut: kupac-hladnjaca " & _
            "(MALINA_DEFAULT_KUPAC) nije pode" & ChrW(353) & "en u Podesavanjima."
        Exit Function
    End If

    Dim hladnjaca As String
    hladnjaca = Trim$(nz(LookupValue(TBL_KUPCI, COL_KUP_ID, kupacID, "Hladnjaca"), ""))

    ' Vozac je obavezan na otpremnici/zbirnoj/prijemnici. U malina/hladnjaca
    ' konvenciji par-vozac ima VozacID == StanicaID. Ako otkup nema vozaca,
    ' koristimo stanicu kao vozaca (i napravimo par-vozaca ako je malina mod).
    If Len(Trim$(vozacID)) = 0 Then
        On Error Resume Next
        EnsureVozacMirrorForStanica stanicaID, _
            Trim$(nz(LookupValue(TBL_STANICE, "StanicaID", stanicaID, "Naziv"), "")), _
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

    ' Broj prijemnice: jedan po dokumentu; Klasa I i II nose ISTI broj
    ' (GenerateBrojPrijemnice, modBrojevi). Generise se i kad se unosi samo Klasa II.
    Dim brPrij As String
    brPrij = GenerateBrojPrijemnice(kupacID, datum)

    ' Klasa I (svoja kolicina ambalaze = kolAmb). Preskace se ako se unosi samo II.
    If hasKlasaI Then
        Dim otpID As String
        otpID = SaveOtpremnica_TX(datum, stanicaID, vozacID, brOtp, brZbr, vrsta, sorta, _
                                  kolicinaI, cenaI, tipAmb, kolAmb, KLASA_I, brutoKgI)
        SaveZbirna_TX datum, vozacID, brZbr, kupacID, hladnjaca, "", vrsta, sorta, _
                      kolicinaI, tipAmb, kolAmb, KLASA_I
        Dim prjI As String
        prjI = SavePrijemnica_TX(datum, kupacID, vozacID, brPrij, brZbr, vrsta, sorta, _
                          kolicinaI, cenaI, tipAmb, kolAmb, 0, KLASA_I, brutoKgI)
        If Len(prjI) = 0 Then failKlase = "I"
        ' Veza nazad u otkup red: OtpremnicaID + BrojZbirne + VozacID.
        LinkOtkupRedNaDokument idI, otpID, brZbr, vozacID
    End If

    ' Klasa II: zasebne gajbe (kolAmbII) -> ceo lanac kao Klasa I; ISTI broj
    ' prijemnice (brPrij) kao Klasa I (jedna prijemnica = jedan broj).
    If hasKlasaII And kolicinaII > 0 Then
        Dim otpID2 As String
        otpID2 = SaveOtpremnica_TX(datum, stanicaID, vozacID, brOtp, brZbr, vrsta, sorta, _
                                   kolicinaII, cenaII, tipAmb, kolAmbII, KLASA_II, brutoKgII)
        SaveZbirna_TX datum, vozacID, brZbr, kupacID, hladnjaca, "", vrsta, sorta, _
                      kolicinaII, tipAmb, kolAmbII, KLASA_II
        Dim prjII As String
        prjII = SavePrijemnica_TX(datum, kupacID, vozacID, brPrij, brZbr, vrsta, sorta, _
                          kolicinaII, cenaII, tipAmb, kolAmbII, 0, KLASA_II, brutoKgII)
        If Len(prjII) = 0 Then
            If Len(failKlase) > 0 Then failKlase = failKlase & " i II" Else failKlase = "II"
        End If
        LinkOtkupRedNaDokument idII, otpID2, brZbr, vozacID
    End If

    ' Vidljivo upozorenje za frmOtkup: otpremnica/zbirna su kreirane, ali prijemnica
    ' nije. Najcesci uzrok: broj prijemnice je vec paletizovan (zaostala stavka u
    ' tblPaletaStavka posle ciscenja tblPrijemnica). Tehnicki razlog je u logu
    ' (Monitor_Event DOKUMENT_SAVE_FAIL / SavePrijemnica_TX).
    If Len(failKlase) > 0 Then
        AutoChainHladnjaca = "Otkup je sa" & ChrW(269) & "uvan, ali AUTO-LANAC hladnjace je NEPOTPUN: " & _
            "otpremnica i zbirna su kreirane, a PRIJEMNICA nije (Klasa " & failKlase & "). " & _
            "Najcesci uzrok: broj prijemnice je ve" & ChrW(263) & " paletizovan (zaostala stavka u " & _
            "tblPaletaStavka). Detalji su u logu."
    End If
    Exit Function

EH:
    Dim eDesc As String: eDesc = Err.description
    LogErr "modAutoHladnjaca.AutoChainHladnjaca"
    AutoChainHladnjaca = "Auto-lanac hladnjace prekinut gre" & ChrW(353) & "kom: " & eDesc & " (vidi log)."
End Function

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
    curVoz = Trim$(nz(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_VOZAC), ""))

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

' ============================================================
' Jednokratni backfill: kreira prijemnice za vec postojece otpremnice
' hladnjaca-lanca kojima prijemnica NEDOSTAJE (npr. posle reseta tblPrijemnica
' kada je auto-lanac pravio otpremnicu+zbirnu, a prijemnica padala).
'
' Anchor = otpremnica (jedina nosi sve sto prijemnici treba: cena + bruto +
' kol. ambalaze po klasi). Idempotentno: preskace (BrojZbirne + Klasa) za koji
' vec postoji ne-stornirana prijemnica. Klasa I i II istog dokumenta (isti
' BrojZbirne) dele JEDAN broj prijemnice (kao i live auto-lanac).
'
' VAZNO (pre pokretanja): ako je tblPrijemnica praznjena, obrisi orphan redove u
' tblPaleta i tblPaletaStavka -> inace paletizacija u SavePrijemnica_TX puca
' ("vec paletizovana") i prijemnica se ne kreira. Macro tada NE rusi nista
' (svaka prijemnica je atomicna): samo prijavi koliko je palo.
'
' Pokretanje: Alt+F8 -> BackfillPrijemniceHladnjaca.
' Reuse: SavePrijemnica_TX, GenerateBrojPrijemnice, IsHladnjacaStanica,
'        GetTableData / RequireColumnIndex / GetColumnIndex.
' ============================================================
Public Sub BackfillPrijemniceHladnjaca()
    Const SRC As String = "modAutoHladnjaca.BackfillPrijemniceHladnjaca"
    On Error GoTo EH

    Dim kupacID As String
    kupacID = Trim$(GetConfigValue(CFG_MALINA_DEFAULT_KUPAC))
    If Len(kupacID) = 0 Then
        MsgBox "MALINA_DEFAULT_KUPAC (kupac-hladnjaca) nije pode" & ChrW(353) & "en. " & _
               "Podesi ga u Podesavanjima pa pokreni ponovo.", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim otp As Variant: otp = GetTableData(TBL_OTPREMNICA)
    If IsEmpty(otp) Then
        MsgBox "Nema otpremnica za backfill.", vbInformation, APP_NAME
        Exit Sub
    End If

    ' Kolone otpremnice (izvor podataka).
    Dim cDat As Long, cSta As Long, cVoz As Long, cZbr As Long, cVrs As Long
    Dim cSor As Long, cKol As Long, cCen As Long, cTip As Long, cAmb As Long
    Dim cKla As Long, cBru As Long, cStorno As Long
    cDat = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_DATUM, SRC)
    cSta = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_STANICA, SRC)
    cVoz = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VOZAC, SRC)
    cZbr = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, SRC)
    cVrs = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VRSTA, SRC)
    cSor = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_SORTA, SRC)
    cKol = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA, SRC)
    cCen = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_CENA, SRC)
    cTip = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_TIP_AMB, SRC)
    cAmb = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KOL_AMB, SRC)
    cKla = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KLASA, SRC)
    cBru = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BRUTO)        ' opciono
    cStorno = GetColumnIndex(TBL_OTPREMNICA, COL_STORNIRANO)    ' opciono

    ' Postojece (ne-stornirane) prijemnice -> set kljuceva (BrojZbirne|Klasa).
    Dim have As Object: Set have = CreateObject("Scripting.Dictionary")
    Dim prj As Variant: prj = GetTableData(TBL_PRIJEMNICA)
    If Not IsEmpty(prj) Then
        Dim pZbr As Long, pKla As Long, pStorno As Long
        pZbr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, SRC)
        pKla = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KLASA, SRC)
        pStorno = GetColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO)
        Dim pr As Long
        For pr = 1 To UBound(prj, 1)
            If pStorno = 0 Or UCase$(Trim$(CStr(prj(pr, pStorno)))) <> "DA" Then
                have(KeyZbrKlasa(CStr(prj(pr, pZbr)), CStr(prj(pr, pKla)))) = True
            End If
        Next pr
    End If

    ' Kandidati: ne-stornirane hladnjaca-otpremnice bez (BrojZbirne|Klasa) prijemnice.
    Dim cand As Collection: Set cand = New Collection
    Dim r As Long
    For r = 1 To UBound(otp, 1)
        If (cStorno = 0) Or (UCase$(Trim$(CStr(otp(r, cStorno)))) <> "DA") Then
            Dim sid As String: sid = Trim$(CStr(otp(r, cSta)))
            Dim zbr As String: zbr = Trim$(CStr(otp(r, cZbr)))
            Dim kla As String: kla = Trim$(CStr(otp(r, cKla)))
            If Len(zbr) > 0 And IsHladnjacaStanica(sid) Then
                If Not have.Exists(KeyZbrKlasa(zbr, kla)) Then cand.Add r
            End If
        End If
    Next r

    If cand.count = 0 Then
        MsgBox "Nema hladnjaca-otpremnica bez prijemnice. Nista za backfill.", _
               vbInformation, APP_NAME
        Exit Sub
    End If

    If MsgBox("Pronadjeno " & cand.count & " otpremnica (hladnjaca) bez prijemnice." & _
              vbCrLf & "Kreirati prijemnice za njih sada?" & vbCrLf & vbCrLf & _
              "NAPOMENA: ako je tblPrijemnica ranije praznjena, prvo ocisti orphan " & _
              "redove u tblPaleta i tblPaletaStavka (inace paletizacija puca).", _
              vbQuestion + vbYesNo, APP_NAME) <> vbYes Then Exit Sub

    ' Jedan broj prijemnice po dokumentu (BrojZbirne) -> Klasa I i II dele broj.
    Dim brByZbr As Object: Set brByZbr = CreateObject("Scripting.Dictionary")
    Dim ok As Long, fail As Long, i As Long
    For i = 1 To cand.count
        r = cand(i)
        If Not IsDate(otp(r, cDat)) Then
            fail = fail + 1
            GoTo ContinueLoop
        End If
        Dim dDat As Date: dDat = CDate(otp(r, cDat))
        Dim zb As String: zb = Trim$(CStr(otp(r, cZbr)))

        Dim brPrij As String
        If brByZbr.Exists(zb) Then
            brPrij = CStr(brByZbr(zb))
        Else
            brPrij = GenerateBrojPrijemnice(kupacID, dDat)
            brByZbr(zb) = brPrij
        End If

        Dim bru As Double: bru = 0
        If cBru > 0 Then bru = AsDbl(otp(r, cBru))

        Dim res As String
        res = SavePrijemnica_TX(dDat, kupacID, Trim$(CStr(otp(r, cVoz))), brPrij, zb, _
                  Trim$(CStr(otp(r, cVrs))), Trim$(CStr(otp(r, cSor))), _
                  AsDbl(otp(r, cKol)), AsDbl(otp(r, cCen)), Trim$(CStr(otp(r, cTip))), _
                  AsLng(otp(r, cAmb)), 0, ClassOrDefault(otp(r, cKla)), bru)
        If Len(res) > 0 Then ok = ok + 1 Else fail = fail + 1
ContinueLoop:
    Next i

    LogInfo SRC, "Backfill prijemnice: ok=" & ok & " fail=" & fail
    MsgBox "Backfill zavr" & ChrW(353) & "en." & vbCrLf & _
           "Kreirano prijemnica: " & ok & vbCrLf & _
           "Neuspesno: " & fail & _
           IIf(fail > 0, vbCrLf & "(vidi log; najcesce orphan paleta -> ocisti " & _
                              "tblPaleta/tblPaletaStavka pa pokreni ponovo)", ""), _
           IIf(fail > 0, vbExclamation, vbInformation), APP_NAME
    Exit Sub
EH:
    LogErr SRC
    MsgBox "Gre" & ChrW(353) & "ka u backfill-u: " & Err.description, vbCritical, APP_NAME
End Sub

Private Function KeyZbrKlasa(ByVal zbr As String, ByVal klasa As String) As String
    KeyZbrKlasa = Trim$(zbr) & "|" & Trim$(klasa)
End Function

' Javni Nz (modHelpers) je String-tipa; ovde trebaju brojevi iz numerickih
' celija (NzD/NzL u modPaletniList su Private -> nedostupni). Minimalne lokalne
' koercije, IsNumeric-bezbedne (prazna/ne-broj celija -> 0).
Private Function AsDbl(ByVal v As Variant) As Double
    If IsNumeric(v) Then AsDbl = CDbl(v)
End Function

Private Function AsLng(ByVal v As Variant) As Long
    If IsNumeric(v) Then AsLng = CLng(v)
End Function

Private Function ClassOrDefault(ByVal v As Variant) As String
    ClassOrDefault = Trim$(CStr(v))
    If Len(ClassOrDefault) = 0 Then ClassOrDefault = KLASA_I
End Function

