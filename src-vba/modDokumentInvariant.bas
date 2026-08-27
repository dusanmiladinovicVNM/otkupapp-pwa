Attribute VB_Name = "modDokumentInvariant"
Option Explicit

' ============================================================
' modDokumentInvariant - centralni invariant engine
'
' Kljucni poslovni invariant:
'   ZBIRNA = tacno zbir svih svojih AKTIVNIH otpremnica (po BrojZbirne).
'   - KG se proverava PO KLASI (I / II) -> hard.
'   - Ambalaza UKUPNO -> hard.
'   - Ambalaza PO KLASI -> soft (postojeci rucni unos zbirne validira amb samo
'     kao ukupno; realni podaci mogu imati amb koncentriranu na jednoj klasi).
'     RecalculateZbirnaFromOtpremnice_TX upisuje amb PO KLASI pa posle korekcije
'     i strogi per-klasa invariant prolazi.
'
' Izvor istine: OTPREMNICE. Zbirna je agregat. Ako se menja otpremnica -> mora se
' validirati i/ili rekalkulisati zbirna.
'
' Stil: reuse modDataAccess (GetTableData/GetColumnIndex/RequireUpdateCell),
' clsTransaction za mutacije, LogErr/Monitor_Event za greske. Bez MsgBox
' (business sloj). Sve mutacije u *_TX funkciji.
' ============================================================

Private Const MOD_NAME As String = "modDokumentInvariant"
Private Const EPS_KG As Double = 0.01

' Test-observability seam: Monitor_Event je HTTP (nema lokalni red) i moze biti
' iskljucen, pa se emisija audita ne moze asertovati direktno. AuditIssuedZbirnaChange
' usput postavlja ovaj marker (delta poslednjeg audita izdate zbirne) da regres-test
' potvrdi da je gate-putanja (izdato + promena) stvarno prosla. Ne utice na ponasanje.
Private mLastIssuedZbirnaAudit As String

' ============================================================
' Per-klasa suma AKTIVNIH otpremnica za dati BrojZbirne.
' Vraca Scripting.Dictionary sa kljucevima:
'   kgI, kgII, kgOther, kgTotal
'   ambI, ambII, ambOther, ambTotal
'   nRows, nRowsI, nRowsII
'   vrstaI, sortaI, tipAmbI (reprezentativna zaglavlja za klasu I)
'   vrstaII, sortaII, tipAmbII
' ============================================================
Public Function SumOtpremniceByKlasa(ByVal brojZbirne As String) As Object
    Const SRC As String = MOD_NAME & ".SumOtpremniceByKlasa"
    Dim d As Object
    Set d = NewSumDict()
    Set SumOtpremniceByKlasa = d

    On Error GoTo EH
    brojZbirne = Trim$(brojZbirne)
    If Len(brojZbirne) = 0 Then Exit Function

    Dim data As Variant
    data = GetTableData(TBL_OTPREMNICA)
    If IsEmpty(data) Then Exit Function

    Dim cZbr As Long, cKol As Long, cAmb As Long, cKlasa As Long, cStorno As Long
    Dim cVrsta As Long, cSorta As Long, cTipAmb As Long
    cZbr = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, SRC)
    cKol = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA, SRC)
    cAmb = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KOL_AMB, SRC)
    cKlasa = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KLASA, SRC)
    cStorno = RequireColumnIndex(TBL_OTPREMNICA, COL_STORNIRANO, SRC)
    cVrsta = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_VRSTA)
    cSorta = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_SORTA)
    cTipAmb = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_TIP_AMB)

    Dim i As Long, klasa As String
    Dim kol As Double, amb As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cZbr))) = brojZbirne Then
            If Not IsDaFlag(data(i, cStorno)) Then
                klasa = Trim$(CStr(data(i, cKlasa)))
                kol = 0#: amb = 0
                If IsNumeric(data(i, cKol)) Then kol = CDbl(data(i, cKol))
                If IsNumeric(data(i, cAmb)) Then amb = CLng(data(i, cAmb))

                d("kgTotal") = CDbl(d("kgTotal")) + kol
                d("ambTotal") = CLng(d("ambTotal")) + amb
                d("nRows") = CLng(d("nRows")) + 1

                If klasa = KLASA_I Then
                    d("kgI") = CDbl(d("kgI")) + kol
                    d("ambI") = CLng(d("ambI")) + amb
                    d("nRowsI") = CLng(d("nRowsI")) + 1
                    CaptureHeader d, "I", data, i, cVrsta, cSorta, cTipAmb
                ElseIf klasa = KLASA_II Then
                    d("kgII") = CDbl(d("kgII")) + kol
                    d("ambII") = CLng(d("ambII")) + amb
                    d("nRowsII") = CLng(d("nRowsII")) + 1
                    CaptureHeader d, "II", data, i, cVrsta, cSorta, cTipAmb
                Else
                    d("kgOther") = CDbl(d("kgOther")) + kol
                    d("ambOther") = CLng(d("ambOther")) + amb
                End If
            End If
        End If
    Next i
    Exit Function
EH:
    LogErr SRC
End Function

' Per-klasa suma AKTIVNIH zbirna redova za dati BrojZbirne (isti oblik kao gore,
' ali samo kg/amb + broj redova; zaglavlja se ne skupljaju).
Private Function SumZbirnaByKlasa(ByVal brojZbirne As String) As Object
    Const SRC As String = MOD_NAME & ".SumZbirnaByKlasa"
    Dim d As Object
    Set d = NewSumDict()
    Set SumZbirnaByKlasa = d

    On Error GoTo EH
    brojZbirne = Trim$(brojZbirne)
    If Len(brojZbirne) = 0 Then Exit Function

    Dim data As Variant
    data = GetTableData(TBL_ZBIRNA)
    If IsEmpty(data) Then Exit Function

    Dim cBroj As Long, cKol As Long, cAmb As Long, cKlasa As Long, cStorno As Long
    cBroj = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, SRC)
    cKol = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KOLICINA, SRC)
    cAmb = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KOL_AMB, SRC)
    cKlasa = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KLASA, SRC)
    cStorno = RequireColumnIndex(TBL_ZBIRNA, COL_STORNIRANO, SRC)

    Dim i As Long, klasa As String
    Dim kol As Double, amb As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cBroj))) = brojZbirne Then
            If Not IsDaFlag(data(i, cStorno)) Then
                klasa = Trim$(CStr(data(i, cKlasa)))
                kol = 0#: amb = 0
                If IsNumeric(data(i, cKol)) Then kol = CDbl(data(i, cKol))
                If IsNumeric(data(i, cAmb)) Then amb = CLng(data(i, cAmb))

                d("kgTotal") = CDbl(d("kgTotal")) + kol
                d("ambTotal") = CLng(d("ambTotal")) + amb
                d("nRows") = CLng(d("nRows")) + 1

                If klasa = KLASA_I Then
                    d("kgI") = CDbl(d("kgI")) + kol
                    d("ambI") = CLng(d("ambI")) + amb
                    d("nRowsI") = CLng(d("nRowsI")) + 1
                ElseIf klasa = KLASA_II Then
                    d("kgII") = CDbl(d("kgII")) + kol
                    d("ambII") = CLng(d("ambII")) + amb
                    d("nRowsII") = CLng(d("nRowsII")) + 1
                Else
                    d("kgOther") = CDbl(d("kgOther")) + kol
                    d("ambOther") = CLng(d("ambOther")) + amb
                End If
            End If
        End If
    Next i
    Exit Function
EH:
    LogErr SRC
End Function

' ============================================================
' ValidateZbirnaInvariant: zbirna vs suma aktivnih otpremnica.
' Vraca Scripting.Dictionary. Kljucne provere:
'   kgOkI, kgOkII   (hard, tolerancija EPS_KG)
'   ambOkTotal      (hard, integer jednakost)
'   ambOkI, ambOkII (soft; granularnost postojeceg unosa)
'   isValid         = kgOkI And kgOkII And ambOkTotal
'   isValidStrict   = isValid And ambOkI And ambOkII
' Plus sirovi brojevi (otp/zbr po klasi) i "message" (kratak opis mismatch-a).
' ============================================================
Public Function ValidateZbirnaInvariant(ByVal brojZbirne As String) As Object
    Const SRC As String = MOD_NAME & ".ValidateZbirnaInvariant"
    Dim r As Object
    Set r = CreateObject("Scripting.Dictionary")
    Set ValidateZbirnaInvariant = r

    On Error GoTo EH
    brojZbirne = Trim$(brojZbirne)

    Dim o As Object, z As Object
    Set o = SumOtpremniceByKlasa(brojZbirne)
    Set z = SumZbirnaByKlasa(brojZbirne)

    r("brojZbirne") = brojZbirne
    r("hasOtpremnice") = (CLng(o("nRows")) > 0)
    r("hasZbirna") = (CLng(z("nRows")) > 0)

    ' KG po klasi (+ ostale klase kao total-doprinos)
    r("kgOtpI") = CDbl(o("kgI")): r("kgZbrI") = CDbl(z("kgI"))
    r("kgOtpII") = CDbl(o("kgII")): r("kgZbrII") = CDbl(z("kgII"))
    r("kgOtpTotal") = CDbl(o("kgTotal")): r("kgZbrTotal") = CDbl(z("kgTotal"))
    r("kgDiffI") = CDbl(o("kgI")) - CDbl(z("kgI"))
    r("kgDiffII") = CDbl(o("kgII")) - CDbl(z("kgII"))
    r("kgDiffTotal") = CDbl(o("kgTotal")) - CDbl(z("kgTotal"))
    r("kgOkI") = (Abs(r("kgDiffI")) < EPS_KG)
    r("kgOkII") = (Abs(r("kgDiffII")) < EPS_KG)
    r("kgOkTotal") = (Abs(r("kgDiffTotal")) < EPS_KG)

    ' Ambalaza
    r("ambOtpI") = CLng(o("ambI")): r("ambZbrI") = CLng(z("ambI"))
    r("ambOtpII") = CLng(o("ambII")): r("ambZbrII") = CLng(z("ambII"))
    r("ambOtpTotal") = CLng(o("ambTotal")): r("ambZbrTotal") = CLng(z("ambTotal"))
    r("ambDiffI") = CLng(o("ambI")) - CLng(z("ambI"))
    r("ambDiffII") = CLng(o("ambII")) - CLng(z("ambII"))
    r("ambDiffTotal") = CLng(o("ambTotal")) - CLng(z("ambTotal"))
    r("ambOkI") = (r("ambDiffI") = 0)
    r("ambOkII") = (r("ambDiffII") = 0)
    r("ambOkTotal") = (r("ambDiffTotal") = 0)

    r("isValid") = (r("kgOkI") And r("kgOkII") And r("kgOkTotal") And r("ambOkTotal"))
    r("isValidStrict") = (r("isValid") And r("ambOkI") And r("ambOkII"))
    r("message") = BuildInvariantMessage(r)
    Exit Function
EH:
    LogErr SRC
    If Not r.Exists("isValid") Then r("isValid") = False
    If Not r.Exists("isValidStrict") Then r("isValidStrict") = False
    If Not r.Exists("message") Then r("message") = "Greska pri proveri invarijante."
End Function

' Kratka bool provera (za guard-ove / testove).
Public Function IsZbirnaConsistent(ByVal brojZbirne As String) As Boolean
    Dim r As Object
    Set r = ValidateZbirnaInvariant(brojZbirne)
    IsZbirnaConsistent = CBool(r("isValid"))
End Function

' ============================================================
' RecalculateZbirnaFromOtpremnice_TX: upisi zbirna redove tako da budu tacno
' jednaki sumi AKTIVNIH otpremnica PO KLASI (kg + amb). Otpremnice = izvor istine.
'
' Ponasanje po klasi K (I/II) za dati BrojZbirne:
'   - postoji aktivan zbirna red -> UpdateCell UkupnoKolicina/UkupnoAmbalaze
'   - nema zbirna reda a ima otpremnica -> AppendRow novi zbirna red (zaglavlje
'     kopirano iz postojeceg zbirna reda istog broja + vrsta/sorta/tipAmb iz
'     otpremnica te klase)
'   - postoji zbirna red a nema otpremnica te klase -> upisi 0/0 (0 = zbir niceg)
'
' Ne dira Vrsta/Sorta/TipAmb na postojecim redovima (minimalna izmena). Vraca
' True na uspeh. Raise ako uopste nema zbirna reda za broj (nema zaglavlja za
' nasledjivanje) -> caller (modStornoFlow) to hvata kao MANUAL_REQUIRED.
' ============================================================
Public Function RecalculateZbirnaFromOtpremnice_TX(ByVal brojZbirne As String, _
        Optional ByVal correctionID As String = "", Optional ByVal reason As String = "") As Boolean
    Const SRC As String = MOD_NAME & ".RecalculateZbirnaFromOtpremnice_TX"
    Dim tx As clsTransaction
    On Error GoTo EH

    brojZbirne = Trim$(brojZbirne)
    If Len(brojZbirne) = 0 Then Exit Function

    ' CENTRALNA KAPIJA -- U primitivu, ne oko njega.
    '
    ' Ova rutina mutira SVE zbirna redove sa datim brojem, a broj nije
    ' identitet: dva vlasnika mogu nositi isti. Do sada je zastita stajala po
    ' call-site-u -- ZbirnaBrojJeDvosmislenIkad na sest mesta u modStornoFlow --
    ' pa je nov pozivalac bio bezbedan samo ako se autor kapije seti. Katalog
    ' je to i trazio: centralna kapija umesto zastite po call-site-u.
    '
    ' Broji se IKAD, ne samo aktivni: storniran vlasnik i dalje ima aktivnu
    ' decu, a posle storna je aktivan jedan pa broj izgleda jednoznacan.
    '
    ' Kapija dize gresku; EH je hvata i vraca False, sto pozivaoci vec
    ' obradjuju kao neuspeh (MANUAL_REQUIRED). Test 124.
    RequireJedanVlasnikIkadPoBroju TBL_ZBIRNA, COL_ZBR_BROJ, brojZbirne, SRC, _
                                   COL_ZBR_VOZAC, COL_ZBR_KUPAC

    Dim o As Object
    Set o = SumOtpremniceByKlasa(brojZbirne)

    ' Mapiraj postojece AKTIVNE zbirna redove po klasi (row index) + zapamti
    ' zaglavlje reda za nasledjivanje kad neka klasa fali.
    Dim data As Variant
    data = GetTableData(TBL_ZBIRNA)
    If IsEmpty(data) Then
        Err.Raise ERR_STORNO_FW_BASE + 1, SRC, _
                  "Nema zbirna reda za broj: " & brojZbirne & " (nista za rekalkulaciju)."
    End If

    Dim cBroj As Long, cKlasa As Long, cStorno As Long, cKol As Long, cAmb As Long
    cBroj = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, SRC)
    cKlasa = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KLASA, SRC)
    cStorno = RequireColumnIndex(TBL_ZBIRNA, COL_STORNIRANO, SRC)
    cKol = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KOLICINA, SRC)
    cAmb = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KOL_AMB, SRC)

    Dim rowI As Long, rowII As Long, templateRow As Long
    rowI = 0: rowII = 0: templateRow = 0
    Dim i As Long, klasa As String
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cBroj))) = brojZbirne Then
            If templateRow = 0 Then templateRow = i     ' prvi red (aktivan ili ne) = izvor zaglavlja
            If Not IsDaFlag(data(i, cStorno)) Then
                klasa = Trim$(CStr(data(i, cKlasa)))
                If klasa = KLASA_I Then
                    rowI = i
                ElseIf klasa = KLASA_II Then
                    rowII = i
                End If
                If templateRow = 0 Then templateRow = i
            End If
        End If
    Next i

    If templateRow = 0 Then
        Err.Raise ERR_STORNO_FW_BASE + 1, SRC, _
                  "Nema zbirna reda za broj: " & brojZbirne & " (nista za rekalkulaciju)."
    End If

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA

    ' Klasa I
    ApplyKlasaRecalc SRC, brojZbirne, KLASA_I, rowI, _
        CDbl(o("kgI")), CLng(o("ambI")), CLng(o("nRowsI")), _
        CStr(o("vrstaI")), CStr(o("sortaI")), CStr(o("tipAmbI")), templateRow, _
        correctionID, reason

    ' Klasa II
    ApplyKlasaRecalc SRC, brojZbirne, KLASA_II, rowII, _
        CDbl(o("kgII")), CLng(o("ambII")), CLng(o("nRowsII")), _
        CStr(o("vrstaII")), CStr(o("sortaII")), CStr(o("tipAmbII")), templateRow, _
        correctionID, reason

    tx.CommitTx
    Set tx = Nothing

    RecalculateZbirnaFromOtpremnice_TX = True
    Monitor_Event eventType:="ZBIRNA_RECALC", severity:="INFO", _
        message:="Zbirna " & brojZbirne & " rekalkulisana iz otpremnica.", _
        moduleName:=MOD_NAME, procedureName:="RecalculateZbirnaFromOtpremnice_TX", _
        entityType:="Zbirna", entityID:=brojZbirne, correlationId:=brojZbirne
    Exit Function
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr SRC
    RecalculateZbirnaFromOtpremnice_TX = False
End Function

' Upisi jednu klasu zbirne na (kg, amb). Kreira red ako fali a ima otpremnica.
Private Sub ApplyKlasaRecalc(ByVal SRC As String, ByVal brojZbirne As String, _
                             ByVal klasa As String, ByVal existingRow As Long, _
                             ByVal kg As Double, ByVal amb As Long, ByVal nOtp As Long, _
                             ByVal vrsta As String, ByVal sorta As String, _
                             ByVal tipAmb As String, ByVal templateRow As Long, _
                             Optional ByVal correctionID As String = "", _
                             Optional ByVal reason As String = "")
    If existingRow > 0 Then
        ' Audit-trag (bez re-verzije): ako se IZDATA zbirna menja in-place, zabelezi
        ' staru->novu vrednost + CorrectionID/razlog. Zbirna je izveden agregat (suma
        ' aktivnih otpremnica) pa se NE re-verzionise, ali izmena izdatog dokumenta
        ' mora ostaviti trag (ADR-0001/0002). Eksplicitna ISPRAVKA vec nosi svoj trace.
        Dim oldKg As Double, oldAmb As Long
        Dim zdata As Variant: zdata = GetTableData(TBL_ZBIRNA)
        If IsArray(zdata) Then
            Dim ck As Long, ca As Long
            ck = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_KOLICINA)
            ca = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_KOL_AMB)
            If ck > 0 Then oldKg = SafeDbl(zdata(existingRow, ck))
            If ca > 0 Then oldAmb = SafeLng(zdata(existingRow, ca))
        End If

        RequireUpdateCell TBL_ZBIRNA, existingRow, COL_ZBR_KOLICINA, kg, SRC
        RequireUpdateCell TBL_ZBIRNA, existingRow, COL_ZBR_KOL_AMB, amb, SRC

        If (Abs(oldKg - kg) > EPS_KG) Or (oldAmb <> amb) Then
            If DocIsIssued(TBL_ZBIRNA, COL_ZBR_BROJ, brojZbirne) Then
                AuditIssuedZbirnaChange brojZbirne, klasa, oldKg, kg, oldAmb, amb, correctionID, reason
            End If
        End If
        Exit Sub
    End If

    ' Nema zbirna reda za ovu klasu. Kreiraj SAMO ako ta klasa ima otpremnica.
    If nOtp = 0 Then Exit Sub

    ' Nasledi zaglavlje iz template reda (isti BrojZbirne).
    Dim data As Variant: data = GetTableData(TBL_ZBIRNA)
    Dim datum As Variant, vozac As String, kupac As String, hladnjaca As String, pogon As String
    datum = data(templateRow, RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_DATUM, SRC))
    vozac = NzTx(data(templateRow, GetColumnIndex(TBL_ZBIRNA, COL_ZBR_VOZAC)))
    kupac = NzTx(data(templateRow, GetColumnIndex(TBL_ZBIRNA, COL_ZBR_KUPAC)))
    hladnjaca = NzTx(data(templateRow, GetColumnIndex(TBL_ZBIRNA, COL_ZBR_HLADNJACA)))
    pogon = NzTx(data(templateRow, GetColumnIndex(TBL_ZBIRNA, COL_ZBR_POGON)))

    Dim newID As String
    newID = GetNextID(TBL_ZBIRNA, COL_ZBR_ID, "ZBR-")

    ' NAMERNO se NE zove modDokumenta.SaveZbirna: njegov ValidateZbirnaInput je
    ' pravilo za OPERATERSKI unos (vozac/kupac obavezni, kg > 0). Rekalkulacija je
    ' SISTEMSKA korekcija koja mora da odrzi invarijantu i sa retkim zaglavljem
    ' (nasledjenim iz template reda) -> koristi se direktan AppendRow.
    ' Redosled kolona je IDENTICAN modDokumenta.SaveZbirna (potvrdjen izvor istine):
    ' ID, Datum, VozacID, BrojZbirne, KupacID, Hladnjaca, Pogon, VrstaVoca,
    ' SortaVoca, UkupnoKolicina, TipAmbalaze, UkupnoAmbalaze, Klasa
    Dim rowData(0 To 12) As Variant
    rowData(0) = newID
    rowData(1) = datum
    rowData(2) = vozac
    rowData(3) = brojZbirne
    rowData(4) = kupac
    rowData(5) = hladnjaca
    rowData(6) = pogon
    rowData(7) = vrsta
    rowData(8) = sorta
    rowData(9) = kg
    rowData(10) = tipAmb
    rowData(11) = amb
    rowData(12) = klasa

    Dim newRow As Long
    newRow = AppendRow(TBL_ZBIRNA, rowData)

    If newRow = 0 Then
        Err.Raise ERR_STORNO_FW_BASE + 2, SRC, _
                  "AppendRow zbirna (klasa " & klasa & ") nije uspeo."
    End If

    ' Generacija: nasledjuje se od aktivnih redova istog broja (druga klasa iste
    ' rekalkulacije), inace nova. Bez ovoga bi sistemski upis ostao bez generacije.
    ApplyGeneracijaID TBL_ZBIRNA, newRow, COL_ZBR_BROJ, brojZbirne, _
                      COL_ZBR_VOZAC, vozac, COL_ZBR_KUPAC, kupac
End Sub

' Audit izmene IZDATE zbirne bez re-verzije (in-place recalc). Durabilan trag u
' Monitoring-u: stara->nova vrednost + CorrectionID/razlog. WARN jer je dirnut izdat
' dokument (operater/kupac mozda imaju stariju verziju).
Private Sub AuditIssuedZbirnaChange(ByVal brojZbirne As String, ByVal klasa As String, _
        ByVal oldKg As Double, ByVal newKg As Double, ByVal oldAmb As Long, ByVal newAmb As Long, _
        ByVal correctionID As String, ByVal reason As String)
    On Error Resume Next
    ' Test-observability marker (pre Monitor_Event-a: belezi da je gate prosla, nezavisno od HTTP-a).
    mLastIssuedZbirnaAudit = brojZbirne & "|K" & klasa & _
        "|kg " & Format$(oldKg, "0.##") & "->" & Format$(newKg, "0.##") & _
        "|amb " & oldAmb & "->" & newAmb
    Dim cidLabel As String: cidLabel = correctionID
    If Len(cidLabel) = 0 Then cidLabel = "(auto-recalc)"
    Dim corr As String: corr = correctionID
    If Len(corr) = 0 Then corr = brojZbirne
    Dim msg As String
    msg = "IZDATA zbirna " & brojZbirne & " [K" & klasa & "] promenjena in-place (bez re-verzije): " & _
          "kg " & Format$(oldKg, "0.##") & "->" & Format$(newKg, "0.##") & ", " & _
          "amb " & oldAmb & "->" & newAmb & ". CorrectionID=" & cidLabel
    If Len(reason) > 0 Then msg = msg & ". Razlog: " & reason
    Monitor_Event eventType:="ZBIRNA_IZDATA_RECALC", severity:="WARN", _
        message:=msg, moduleName:=MOD_NAME, procedureName:="RecalculateZbirnaFromOtpremnice_TX", _
        entityType:="Zbirna", entityID:=brojZbirne, correlationId:=corr
End Sub

' Test-observability: poslednji audit izdate zbirne (delta) + reset. Samo za regres-test.
Public Function LastIssuedZbirnaAudit() As String
    LastIssuedZbirnaAudit = mLastIssuedZbirnaAudit
End Function
Public Sub ResetIssuedZbirnaAudit()
    mLastIssuedZbirnaAudit = ""
End Sub

Private Function SafeDbl(ByVal v As Variant) As Double
    On Error Resume Next
    If Not IsError(v) And Not IsNull(v) And Not IsEmpty(v) Then SafeDbl = CDbl(Val(CStr(v)))
End Function

Private Function SafeLng(ByVal v As Variant) As Long
    On Error Resume Next
    If Not IsError(v) And Not IsNull(v) And Not IsEmpty(v) Then SafeLng = CLng(Val(CStr(v)))
End Function

' ============================================================
' ValidateOtpremnicaZbirnaImpact: uticaj prevezivanja otpremnice sa STARE na
' NOVU zbirnu. Vraca Dictionary sa invarijantom OBE zbirne na trenutnom stanju
' (posle promene) + zbirnim verdiktom. Koristi se u ISPRAVKA/DUPLI flow-u da
' operater vidi da nijedna strana nije ostavljena u mismatch-u.
'   old  -> ValidateZbirnaInvariant(oldBrojZbirne)
'   new  -> ValidateZbirnaInvariant(newBrojZbirne)
'   bothValid = old.isValid And new.isValid (prazan broj se tretira kao valid)
' ============================================================
Public Function ValidateOtpremnicaZbirnaImpact(ByVal oldBrojZbirne As String, _
                                               ByVal newBrojZbirne As String) As Object
    Const SRC As String = MOD_NAME & ".ValidateOtpremnicaZbirnaImpact"
    Dim r As Object
    Set r = CreateObject("Scripting.Dictionary")
    Set ValidateOtpremnicaZbirnaImpact = r

    On Error GoTo EH
    oldBrojZbirne = Trim$(oldBrojZbirne)
    newBrojZbirne = Trim$(newBrojZbirne)

    Dim okOld As Boolean, okNew As Boolean
    okOld = True: okNew = True

    If Len(oldBrojZbirne) > 0 Then
        Dim ro As Object: Set ro = ValidateZbirnaInvariant(oldBrojZbirne)
        Set r("old") = ro
        okOld = CBool(ro("isValid"))
    End If
    If Len(newBrojZbirne) > 0 Then
        Dim rn As Object: Set rn = ValidateZbirnaInvariant(newBrojZbirne)
        Set r("new") = rn
        okNew = CBool(rn("isValid"))
    End If

    r("oldValid") = okOld
    r("newValid") = okNew
    r("bothValid") = (okOld And okNew)
    Exit Function
EH:
    LogErr SRC
    r("bothValid") = False
End Function

' ============================================================
' HELPERS
' ============================================================

Private Function NewSumDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("kgI") = 0#: d("kgII") = 0#: d("kgOther") = 0#: d("kgTotal") = 0#
    d("ambI") = 0&: d("ambII") = 0&: d("ambOther") = 0&: d("ambTotal") = 0&
    d("nRows") = 0&: d("nRowsI") = 0&: d("nRowsII") = 0&
    d("vrstaI") = "": d("sortaI") = "": d("tipAmbI") = ""
    d("vrstaII") = "": d("sortaII") = "": d("tipAmbII") = ""
    Set NewSumDict = d
End Function

' ByRef: citac po celiji -- ByVal bi kopirao ceo niz po pozivu (v. KOPIJA_NIZA).
Private Sub CaptureHeader(ByRef d As Object, ByVal klasa As String, ByRef data As Variant, _
                          ByVal rowIdx As Long, ByVal cVrsta As Long, _
                          ByVal cSorta As Long, ByVal cTipAmb As Long)
    ' Zapamti prvu ne-praznu vrednost zaglavlja po klasi (za kreiranje reda u recalc).
    If Len(CStr(d("vrsta" & klasa))) = 0 And cVrsta > 0 Then d("vrsta" & klasa) = NzTx(data(rowIdx, cVrsta))
    If Len(CStr(d("sorta" & klasa))) = 0 And cSorta > 0 Then d("sorta" & klasa) = NzTx(data(rowIdx, cSorta))
    If Len(CStr(d("tipAmb" & klasa))) = 0 And cTipAmb > 0 Then d("tipAmb" & klasa) = NzTx(data(rowIdx, cTipAmb))
End Sub

Private Function IsDaFlag(ByVal v As Variant) As Boolean
    IsDaFlag = (UCase$(Trim$(CStr(v))) = "DA")
End Function

Private Function NzTx(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        NzTx = ""
    Else
        NzTx = Trim$(CStr(v))
    End If
End Function

Private Function BuildInvariantMessage(ByRef r As Object) As String
    If CBool(r("isValid")) Then
        BuildInvariantMessage = "Zbirna = zbir aktivnih otpremnica (OK)."
        Exit Function
    End If

    Dim m As String
    m = "MISMATCH zbirna " & CStr(r("brojZbirne")) & ": "
    If Not CBool(r("kgOkI")) Then _
        m = m & "KG I (otp " & Fmt(r("kgOtpI")) & " / zbr " & Fmt(r("kgZbrI")) & "); "
    If Not CBool(r("kgOkII")) Then _
        m = m & "KG II (otp " & Fmt(r("kgOtpII")) & " / zbr " & Fmt(r("kgZbrII")) & "); "
    If Not CBool(r("ambOkTotal")) Then _
        m = m & "AMB ukupno (otp " & CStr(r("ambOtpTotal")) & " / zbr " & CStr(r("ambZbrTotal")) & "); "
    BuildInvariantMessage = m
End Function

Private Function Fmt(ByVal v As Variant) As String
    On Error Resume Next
    Fmt = Format$(CDbl(v), "0.##")
End Function

' ============================================================
' Faza 7 (3.0) - KANONSKO ADRESIRANJE append-only modela.
' Linijski model: jedan poslovni broj ima vise redova (po klasi). Identitet reda =
' (broj, klasa) -> AKTIVAN red. Vrati indeks tog reda:
'   0  = nema aktivnog reda za (broj, klasa)
'   -1 = VISE aktivnih (integritet povreda; u append-only sme najvise jedan)
' klasa == "" -> ignorisi klasu (match samo po broju; ambiguo kad ima vise klasa).
' Osnov za: PWA sync migraciju (3.1), append-only re-verziju (3.2), citace (3.3).
' ============================================================
Public Function FindSingleActiveRow(ByVal tbl As String, ByVal brojCol As String, _
        ByVal broj As String, ByVal klasaCol As String, ByVal klasa As String) As Long
    On Error GoTo EH
    Dim data As Variant: data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Function
    Dim cBr As Long: cBr = GetColumnIndex(tbl, brojCol)
    If cBr = 0 Then Exit Function
    Dim cSt As Long: cSt = GetColumnIndex(tbl, COL_STORNIRANO)
    Dim cKl As Long: cKl = 0
    If Len(klasaCol) > 0 Then cKl = GetColumnIndex(tbl, klasaCol)
    broj = Trim$(broj): klasa = Trim$(klasa)
    Dim i As Long, found As Long, cnt As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cBr))) = broj Then
            If cSt = 0 Or Not IsDaFlag(data(i, cSt)) Then
                Dim klMatch As Boolean: klMatch = True
                If Len(klasa) > 0 And cKl > 0 Then klMatch = (Trim$(CStr(data(i, cKl))) = klasa)
                If klMatch Then found = i: cnt = cnt + 1
            End If
        End If
    Next i
    If cnt > 1 Then FindSingleActiveRow = -1 Else FindSingleActiveRow = found
    Exit Function
EH:
    LogErr MOD_NAME & ".FindSingleActiveRow"
End Function

' ============================================================
' Faza 7 - IzdatoStatus gate (ADR-0001 granica: izdat/prosledjen dokument je
' nepromenljiv -> koriguje se storno+reizdaj, ne in-place).
' Ova app NEMA "draft" fazu za chain dokumente -> IzdatoStatus je podrazumevano
' IZDATO; prazno = IZDATO (konzervativno). DRAFT je rezervisan (buduci parkiran/
' held dokument), PROSLEDJENO za buduci sync-push ka PWA/kupcu.
' DocIsIssued: True ako je IZDAT/PROSLEDJEN, False SAMO ako eksplicitno DRAFT.
' ============================================================
' #7: IzdatoStatus se cita sa AKTIVNOG reda (LookupActiveID), ne sa bilo kog reda
' (LookupValue je mogao pokupiti STORNIRAN red istog broja i procitati njegov status).
Public Function DocIsIssued(ByVal tbl As String, ByVal brojCol As String, ByVal broj As String) As Boolean
    On Error Resume Next
    DocIsIssued = True                                  ' default: izdato (konzervativno)
    If GetColumnIndex(tbl, COL_TRACE_IZDATO_STATUS) = 0 Then Exit Function
    Dim v As String
    v = UCase$(Trim$(LookupActiveID(tbl, brojCol, broj, COL_TRACE_IZDATO_STATUS)))
    DocIsIssued = (v <> UCase$(IZDATO_DRAFT))
End Function

' Postavi IzdatoStatus na jedan red (buduci prelazi: PROSLEDJENO pri sync-u ka PWA).
' Guarded na kolonu (schema-drift safe).
Public Sub SetIzdatoStatus(ByVal tbl As String, ByVal rowIndex As Long, ByVal status As String)
    On Error Resume Next
    If GetColumnIndex(tbl, COL_TRACE_IZDATO_STATUS) = 0 Then Exit Sub
    UpdateCell tbl, rowIndex, COL_TRACE_IZDATO_STATUS, status
End Sub

' ============================================================
' TEST (Alt+F8) - rollback-safe (clsTransaction snapshot + rollback; fixture ne
' ostaje). Automatske asertacije -> Debug.Print (Ctrl+G). Faza 7 (3.0).
' ============================================================
Public Sub Test_FindSingleActiveRow()
    Dim tx As clsTransaction
    Dim ok As Boolean: ok = True
    On Error GoTo EH
    Const B As String = "SVT-FSAR-Z"
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA
    ' fixture: KlasaI aktivan, KlasaII aktivan, KlasaI storniran (3 reda, isti broj).
    FsarSeed B, "I", ""
    FsarSeed B, "II", ""
    FsarSeed B, "I", "Da"

    Dim rI As Long: rI = FindSingleActiveRow(TBL_ZBIRNA, COL_ZBR_BROJ, B, COL_ZBR_KLASA, "I")
    Dim rII As Long: rII = FindSingleActiveRow(TBL_ZBIRNA, COL_ZBR_BROJ, B, COL_ZBR_KLASA, "II")
    Dim rAmb As Long: rAmb = FindSingleActiveRow(TBL_ZBIRNA, COL_ZBR_BROJ, B, COL_ZBR_KLASA, "")
    Dim rNone As Long: rNone = FindSingleActiveRow(TBL_ZBIRNA, COL_ZBR_BROJ, B, COL_ZBR_KLASA, "III")

    ok = FsarChk(rI > 0, "KlasaI -> jedan aktivan red (" & rI & ")") And ok
    ok = FsarChk(rII > 0, "KlasaII -> jedan aktivan red (" & rII & ")") And ok
    ok = FsarChk(rI <> rII, "KlasaI != KlasaII (razliciti redovi)") And ok
    Dim dchk As Variant: dchk = GetTableData(TBL_ZBIRNA)
    Dim cSt As Long: cSt = GetColumnIndex(TBL_ZBIRNA, COL_STORNIRANO)
    ok = FsarChk(rI > 0 And Not IsDaFlag(dchk(rI, cSt)), "KlasaI red je AKTIVAN (ne storniran)") And ok
    ok = FsarChk(rAmb = -1, "klasa='' + 2 aktivna -> -1 (ambiguous)") And ok
    ok = FsarChk(rNone = 0, "nepostojeca klasa -> 0") And ok

    tx.RollbackTx: Set tx = Nothing
    Debug.Print "=== Test_FindSingleActiveRow: " & IIf(ok, "PROSAO", "PAO") & " (fixture rollback-ovan) ==="
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "Test_FindSingleActiveRow GRESKA: " & Err.description
End Sub

Private Sub FsarSeed(ByVal broj As String, ByVal klasa As String, ByVal storno As String)
    Dim lo As ListObject: Set lo = GetTable(TBL_ZBIRNA)
    If lo Is Nothing Then Exit Sub
    Dim nr As ListRow: Set nr = lo.ListRows.Add
    Dim ri As Long: ri = nr.Index
    UpdateCell TBL_ZBIRNA, ri, COL_ZBR_ID, "SVT-FSAR-" & klasa & "-" & IIf(Len(storno) > 0, "S", "A")
    UpdateCell TBL_ZBIRNA, ri, COL_ZBR_BROJ, broj
    UpdateCell TBL_ZBIRNA, ri, COL_ZBR_KLASA, klasa
    If Len(storno) > 0 Then UpdateCell TBL_ZBIRNA, ri, COL_STORNIRANO, storno
End Sub

Private Function FsarChk(ByVal cond As Boolean, ByVal nm As String) As Boolean
    Debug.Print IIf(cond, "OK   ", "FAIL ") & nm
    FsarChk = cond
End Function
