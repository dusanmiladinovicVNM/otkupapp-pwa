Attribute VB_Name = "modOtkupBlok"
Option Explicit

' ============================================================
' modOtkupBlok - otkupni blokovi po otpremnici: zbirovi, specifikacija, rang.
'
' Stari panel "Otkupni blokovi" u frmOtkup i njegov event-omotac clsBlokUI su
' obrisani u S1b-2: nijedan ulaz ih vise nije kacio (frmOtkupUI je jedina forma).
' Ostaje ono sto koriste ekrani DOKUMENTI i IZVESTAJI:
'   - SumKolByOtp / SumAmbByOtp / ExistingBlokCena / ExistingBlokZbirna /
'     BuildNapisanoByOtp       -- bilans otpremnice i predlog za nov blok;
'   - PrintSpecifikacija       -- specifikacija blokova izabranih otpremnica;
'   - KoopRangRows             -- rang kooperanata po iznosu;
'   - LinkOtkupIDsToOtpremnica -- veza novog bloka na aktivnu otpremnicu.
'
' Kolicina, vrednost i gajbe citaju se sa STAVKI (modOtkup.ZbirStavkiPoOtkupu);
' veza bloka i otpremnice (Otkup.OtpremnicaID) ostaje do S3.
' ============================================================

' Specifikacija RUCNO izabranih otpremnica (postojeci tok: dugme "Biraj
' otpremnice" -> multiselect -> ChrW(352) & "tampaj specifikaciju"). Tanak omotac oko
' zajednickog renderera RenderSpec (filter po skupu OtpremnicaID).
Public Sub PrintSpecifikacija(ByVal otpIDs As Collection)
    On Error GoTo EH
    Dim selSet As Object: Set selSet = CreateObject("Scripting.Dictionary")
    Dim v As Variant
    For Each v In otpIDs
        Dim oid0 As String: oid0 = CStr(v)
        If Not selSet.Exists(oid0) Then selSet.Add oid0, True
    Next v

    Dim subtitle As String
    subtitle = "Datum stampe: " & Format$(Date, "d.m.yyyy") & "     Otpremnica: " & otpIDs.count
    RenderSpec selSet, False, Date, Date, subtitle
    Exit Sub
EH:
    LogErr "modOtkupBlok.PrintSpecifikacija"
    MsgBox "Gre" & ChrW(353) & "ka pri stampi specifikacije: " & Err.description, vbCritical, APP_NAME
End Sub

' Jezgro: ispisuje specifikaciju otkupnih blokova (tabela sa okvirima, A4
' landscape) i exportuje u PDF. Filter po redu:
'   byDate=True  -> kolona Datum u [datumOd, datumDo]  (selSet sme biti Nothing)
'   byDate=False -> OtpremnicaID u selSet              (rucna selekcija)
' Izlaz je sortiran po (Otkupno mesto, Datum) radi grupisanja.
Private Sub RenderSpec(ByVal selSet As Object, ByVal byDate As Boolean, _
                       ByVal datumOd As Date, ByVal datumDo As Date, _
                       ByVal subtitle As String)
    On Error GoTo EH

    Dim dKo As Object: Set dKo = BuildKoopNames()
    Dim dSt As Object: Set dSt = BuildLookup(TBL_STANICE, "StanicaID", "Naziv")
    Dim dZbr As Object: Set dZbr = BuildLookup(TBL_OTPREMNICA, COL_OTP_ID, COL_OTP_BROJ_ZBIRNE)
    Dim dOtp As Object: Set dOtp = BuildLookup(TBL_OTPREMNICA, COL_OTP_ID, COL_OTP_BROJ)
    ' Kolona "Kupac" (firma kome ide roba): BrojZbirne -> KupacID (zbirna) -> Naziv (kupci).
    Dim dKupId As Object: Set dKupId = BuildLookup(TBL_ZBIRNA, COL_ZBR_BROJ, COL_ZBR_KUPAC)
    Dim dKupNaziv As Object: Set dKupNaziv = BuildLookup(TBL_KUPCI, COL_KUP_ID, COL_KUP_NAZIV)
    Dim stopa As Double: stopa = PdvStopa()

    Dim data As Variant: data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then MsgBox "Nema podataka u otkupu.", vbInformation, APP_NAME: Exit Sub
    data = ExcludeStornirano(data, TBL_OTKUP)
    If IsEmpty(data) Then MsgBox "Nema blokova.", vbInformation, APP_NAME: Exit Sub

    Dim cOtp As Long, cKoop As Long, cId As Long, cBr As Long, cDat As Long, cSt As Long
    Dim cVrsta As Long, cSorta As Long
    cOtp = GetColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID)
    cKoop = GetColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT)
    cId = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    cBr = GetColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK)
    cDat = GetColumnIndex(TBL_OTKUP, COL_OTK_DATUM)
    cSt = GetColumnIndex(TBL_OTKUP, COL_OTK_STANICA)
    cVrsta = GetColumnIndex(TBL_OTKUP, COL_OTK_VRSTA)
    cSorta = GetColumnIndex(TBL_OTKUP, COL_OTK_SORTA)

    ' --- 1) skupi indekse redova koji prolaze filter + sort-kljuc (OM | datum) ---
    Dim n As Long: n = UBound(data, 1)
    Dim idx() As Long: ReDim idx(1 To n)
    Dim keys() As String: ReDim keys(1 To n)
    Dim m As Long: m = 0
    Dim i As Long
    For i = 1 To n
        Dim oid As String: oid = Trim$(CStr(data(i, cOtp)))
        Dim pass As Boolean: pass = False
        If byDate Then
            If IsDate(data(i, cDat)) Then
                Dim dd As Double: dd = Int(CDbl(CDate(data(i, cDat))))
                pass = (dd >= Int(CDbl(datumOd)) And dd <= Int(CDbl(datumDo)))
            End If
        ElseIf Not selSet Is Nothing Then
            pass = selSet.Exists(oid)
        End If
        If pass Then
            m = m + 1
            idx(m) = i
            Dim dkey As String: dkey = "00000000"
            If IsDate(data(i, cDat)) Then dkey = Format$(CDate(data(i, cDat)), "yyyymmdd")
            keys(m) = DictVal(dSt, CStr(data(i, cSt))) & "|" & dkey
        End If
    Next i

    If m = 0 Then
        If byDate Then
            MsgBox "Nema otkupnih blokova u izabranom periodu.", vbInformation, APP_NAME
        Else
            MsgBox "Izabrane otpremnice nemaju blokova.", vbInformation, APP_NAME
        End If
        Exit Sub
    End If

    ' --- 2) insertion sort po (Otkupno mesto, Datum) ASC ---
    Dim a As Long, b As Long, ti As Long, tk As String
    For a = 2 To m
        ti = idx(a): tk = keys(a): b = a - 1
        Do While b >= 1
            If keys(b) <= tk Then Exit Do
            idx(b + 1) = idx(b): keys(b + 1) = keys(b): b = b - 1
        Loop
        idx(b + 1) = ti: keys(b + 1) = tk
    Next a

    ' --- 3) prikupi u niz + sume; render i izlaz u modPrint ---
    Dim stavkeZbir As Object: Set stavkeZbir = ZbirStavkiPoOtkupu()
    Dim spec() As Variant: ReDim spec(1 To m, 1 To 13)
    Dim sumKol As Double, sumVred As Double, sumPdv As Double, sumUk As Double, cnt As Long
    Dim j As Long
    For j = 1 To m
        i = idx(j)
        Dim oid2 As String: oid2 = Trim$(CStr(data(i, cOtp)))
        ' Kolicina i vrednost su na STAVKAMA (S1b-2). Cena je bruto po stavci, pa
        ' dokument sa dve klase nema jednu cenu: neto cena reda je prosek.
        Dim z As Variant: z = ZbirStavkiZaOtkup(stavkeZbir, CStr(data(i, cId)), "modOtkupBlok.RenderSpec")
        Dim kol As Double: kol = CDbl(z(0))
        Dim uk As Double: uk = CDbl(z(1))
        Dim vred As Double: vred = uk / (1 + stopa / 100)
        Dim pdv As Double: pdv = uk - vred
        Dim neto As Double: neto = 0
        If kol > 0 Then neto = vred / kol
        spec(j, 1) = DictVal(dZbr, oid2)
        spec(j, 2) = KupacNazivZaZbirnu(dKupId, dKupNaziv, CStr(spec(j, 1)))
        spec(j, 3) = DictVal(dOtp, oid2)
        spec(j, 4) = DictVal(dSt, CStr(data(i, cSt)))
        spec(j, 5) = CStr(data(i, cBr))
        spec(j, 6) = DictVal(dKo, Trim$(CStr(data(i, cKoop))))
        spec(j, 7) = FmtDate(data(i, cDat))
        spec(j, 8) = Trim$(CStr(data(i, cVrsta)) & " " & CStr(data(i, cSorta)))
        spec(j, 9) = kol
        spec(j, 10) = neto
        spec(j, 11) = vred
        spec(j, 12) = pdv
        spec(j, 13) = uk
        sumKol = sumKol + kol: sumVred = sumVred + vred
        sumPdv = sumPdv + pdv: sumUk = sumUk + uk
        cnt = cnt + 1
    Next j

    Dim ws As Worksheet
    Set ws = FillSpecifikacijaSablon(spec, m, subtitle, cnt, sumKol, sumVred, sumPdv, sumUk)
    If ws Is Nothing Then Exit Sub

    Dim pdfPath As String
    pdfPath = EnsureDocFolder(PDF_DIR_SPECIFIKACIJE) & "\Specifikacija_" & Format$(Now, "yyyymmdd_hhnnss") & ".pdf"
    Dim mode As String
    mode = DocResolveMode(GetConfigValue(CFG_SPECIFIKACIJA_PRINT_MODE), "PDF")
    Select Case mode
        Case "PRINT", "PREVIEW"
            DocPrintWs ws, mode
        Case "PDF"
            DocExportPdf ws, pdfPath, True
        ' OFF -> bez izlaza
    End Select
    Exit Sub
EH:
    Application.ScreenUpdating = True
    LogErr "modOtkupBlok.RenderSpec"
    MsgBox "Gre" & ChrW(353) & "ka pri stampi specifikacije: " & Err.description, vbCritical, APP_NAME
End Sub

Public Sub LinkOtkupIDsToOtpremnica(ByVal otkupIDs As String, ByVal otpID As String)
    Dim tx As clsTransaction
    On Error GoTo EH
    If Len(otpID) = 0 Or Len(Trim$(otkupIDs)) = 0 Then Exit Sub

    ' INTEGRITET (guard 1): nikad ne vezuj otkup na storniranu ili nepostojecu
    ' otpremnicu. Ista invarijanta kao modDokumenta.ReassignOtkupToOtpremnica_TX.
    ' Bez ove provere stale mActiveOtpID (npr. posle storna bloka u panelu, dok
    ' panel ostaje otvoren) je vezivao svez otkup na mrtvu otpremnicu -> tiha
    ' korupcija baze (OtkupID pokazuje na storniranu OtpremnicaID).
    If IsEmpty(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_BROJ)) Then Exit Sub
    If UCase$(Trim$(NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, _
                                         COL_STORNIRANO)))) = "DA" Then Exit Sub

    Dim ids() As String: ids = Split(otkupIDs, " + ")

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP

    Dim j As Long
    For j = LBound(ids) To UBound(ids)
        Dim id As String: id = Trim$(ids(j))
        If Len(id) > 0 Then
            ' INTEGRITET (guard 2): NE pregazi vec uspostavljenu vezu. Hladnjaca
            ' auto-lanac (modAutoHladnjaca) je mozda upravo upisao svoju namensku
            ' otpremnicu za ovaj otkup. Panel vezuje SAMO redove bez otpremnice
            ' (prazan OtpremnicaID) -> spreci da panelski mActiveOtpID pregazi
            ' ispravnu vezu auto-lanca.
            If Len(Trim$(NzToText(LookupValue(TBL_OTKUP, COL_OTK_ID, id, _
                                              COL_OTK_OTPREMNICA_ID)))) = 0 Then
                Dim rows As Collection: Set rows = FindRows(TBL_OTKUP, COL_OTK_ID, id)
                Dim k As Long
                For k = 1 To rows.count
                    RequireUpdateCell TBL_OTKUP, rows(k), COL_OTK_OTPREMNICA_ID, otpID, _
                                      "modOtkupBlok.LinkOtkupIDsToOtpremnica"
                    SetOtkupBrojOtpremnice rows(k), otpID
                Next k
            End If
        End If
    Next j

    tx.CommitTx
    Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr "modOtkupBlok.LinkOtkupIDsToOtpremnica"
End Sub

' OtkupID-evi ne-storniranih blokova vezanih za otpremnicu. Veza Otkup.OtpremnicaID
' ostaje do S3; S1b-2 menja samo izvor kolicine -- stavke.
Private Function BlokoviOtpremnice(ByVal otpID As String) As Collection
    Dim res As New Collection
    Set BlokoviOtpremnice = res
    If Len(otpID) = 0 Then Exit Function
    Dim data As Variant: data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function
    data = ExcludeStornirano(data, TBL_OTKUP)
    If IsEmpty(data) Then Exit Function

    Dim cOtp As Long, cId As Long
    cOtp = RequireColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID, "modOtkupBlok.BlokoviOtpremnice")
    cId = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, "modOtkupBlok.BlokoviOtpremnice")
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cOtp))) = otpID Then res.Add Trim$(CStr(data(i, cId)))
    Next i
End Function

' Zbir jednog polja zbira stavki (0 = kg, 2 = gajbe) preko blokova otpremnice.
' Blok bez stavki pada po imenu u ZbirStavkiZaOtkup -- nikad nula.
Private Function SumStavkeByOtp(ByVal otpID As String, ByVal polje As Long, _
                                ByVal sourceName As String) As Double
    Dim ids As Collection: Set ids = BlokoviOtpremnice(otpID)
    If ids.count = 0 Then Exit Function
    Dim zbir As Object: Set zbir = ZbirStavkiPoOtkupu()
    Dim v As Variant, z As Variant, s As Double
    For Each v In ids
        z = ZbirStavkiZaOtkup(zbir, CStr(v), sourceName)
        s = s + CDbl(z(polje))
    Next v
    SumStavkeByOtp = s
End Function

Public Function SumKolByOtp(ByVal otpID As String) As Double
    SumKolByOtp = SumStavkeByOtp(otpID, 0, "modOtkupBlok.SumKolByOtp")
End Function

Public Function SumAmbByOtp(ByVal otpID As String) As Double
    SumAmbByOtp = SumStavkeByOtp(otpID, 2, "modOtkupBlok.SumAmbByOtp")
End Function

' Cena po kojoj su vec pisani blokovi otpremnice: cena PRVE stavke prvog bloka.
' Predlog za nov blok, ne obracun -- dokument sa dve klase ima dve cene.
Public Function ExistingBlokCena(ByVal otpID As String) As Double
    Dim oid As String: oid = Trim$(NzToText(FirstBlokVal(otpID, COL_OTK_ID)))
    If Len(oid) = 0 Then Exit Function
    Dim s As Variant: s = StavkeOtkupaRedovi()
    If Not IsArray(s) Then Exit Function
    Dim i As Long
    For i = 1 To UBound(s, 1)
        If CStr(s(i, 1)) = oid Then ExistingBlokCena = CDbl(s(i, 5)): Exit Function
    Next i
End Function

' Broj zbirne sa vec napisanih blokova otpremnice. Otpremnica svoju zbirnu ne
' mora da zna (veza se pravi kasnije), a blokovi je nose - pa je ovo drugi
' izvor za "zbirna je poznata u ovom trenutku".
Public Function ExistingBlokZbirna(ByVal otpID As String) As String
    ExistingBlokZbirna = Trim$(NzToText(FirstBlokVal(otpID, COL_OTK_BROJ_ZBIRNE)))
End Function

' Vrednost trazene kolone iz PRVOG bloka otpremnice. Blokovi jedne otpremnice
' dele i cenu i zbirnu, pa je prvi red dovoljan; jedno citanje za oba pozivaoca.
Private Function FirstBlokVal(ByVal otpID As String, ByVal col As String) As Variant
    If Len(otpID) = 0 Then Exit Function
    Dim rows As Collection
    Set rows = FindRows(TBL_OTKUP, COL_OTK_OTPREMNICA_ID, otpID)
    If rows.count = 0 Then Exit Function

    Dim data As Variant: data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function
    Dim c As Long: c = GetColumnIndex(TBL_OTKUP, col)
    If c < 1 Then Exit Function
    FirstBlokVal = data(rows(1), c)
End Function

' OtpremnicaID -> ukupna kolicina svih (ne-storniranih) blokova, sa stavki.
' Isti bilans otpremnice koji ekran DOKUMENTI prikazuje u listi otpremnica.
Public Function BuildNapisanoByOtp() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set BuildNapisanoByOtp = d

    Dim data As Variant: data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function
    data = ExcludeStornirano(data, TBL_OTKUP)
    If IsEmpty(data) Then Exit Function
    Dim cOtp As Long, cId As Long
    cOtp = RequireColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID, "modOtkupBlok.BuildNapisanoByOtp")
    cId = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, "modOtkupBlok.BuildNapisanoByOtp")

    Dim zbir As Object: Set zbir = ZbirStavkiPoOtkupu()
    Dim i As Long, z As Variant
    For i = 1 To UBound(data, 1)
        Dim k As String: k = Trim$(CStr(data(i, cOtp)))
        If Len(k) > 0 Then
            z = ZbirStavkiZaOtkup(zbir, Trim$(CStr(data(i, cId))), "modOtkupBlok.BuildNapisanoByOtp")
            d(k) = CDbl(d(k)) + CDbl(z(0))
        End If
    Next i
End Function

Private Function BuildLookup(ByVal tbl As String, ByVal keyName As String, _
                             ByVal valName As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set BuildLookup = d

    Dim data As Variant: data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Function

    Dim ck As Long, cv As Long
    ck = GetColumnIndex(tbl, keyName)
    cv = GetColumnIndex(tbl, valName)
    If ck = 0 Or cv = 0 Then Exit Function

    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim k As String: k = Trim$(CStr(data(i, ck)))
        If Len(k) > 0 And Not d.Exists(k) Then d.Add k, CStr(data(i, cv))
    Next i
End Function

Private Function BuildKoopNames() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set BuildKoopNames = d

    Dim data As Variant: data = GetTableData(TBL_KOOPERANTI)
    If IsEmpty(data) Then Exit Function

    Dim cId As Long, cIme As Long, cPr As Long
    cId = GetColumnIndex(TBL_KOOPERANTI, "KooperantID")
    cIme = GetColumnIndex(TBL_KOOPERANTI, "Ime")
    cPr = GetColumnIndex(TBL_KOOPERANTI, "Prezime")
    If cId = 0 Then Exit Function

    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim k As String: k = Trim$(CStr(data(i, cId)))
        If Len(k) > 0 And Not d.Exists(k) Then
            d.Add k, Trim$(CStr(data(i, cIme)) & " " & CStr(data(i, cPr)))
        End If
    Next i
End Function

Private Function DictVal(ByVal d As Object, ByVal k As String) As String
    If d Is Nothing Then Exit Function
    k = Trim$(k)
    If d.Exists(k) Then DictVal = CStr(d(k))
End Function

' KooperantID -> naziv maticne stanice (OM), za listu kooperanata. Prazno ako
' kooperant nema stanicu ili stanica nema naziv (tada padne na StanicaID).
Private Function BuildKoopOM() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set BuildKoopOM = d

    Dim data As Variant: data = GetTableData(TBL_KOOPERANTI)
    If IsEmpty(data) Then Exit Function

    Dim cId As Long, cSt As Long
    cId = GetColumnIndex(TBL_KOOPERANTI, COL_KOOP_ID)
    cSt = GetColumnIndex(TBL_KOOPERANTI, COL_KOOP_STANICA)
    If cId = 0 Then Exit Function

    Dim dSt As Object: Set dSt = BuildLookup(TBL_STANICE, "StanicaID", "Naziv")

    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim k As String: k = Trim$(CStr(data(i, cId)))
        If Len(k) > 0 And Not d.Exists(k) Then
            Dim stId As String: stId = ""
            If cSt > 0 Then stId = Trim$(CStr(data(i, cSt)))
            Dim nm As String: nm = DictVal(dSt, stId)
            If Len(nm) = 0 Then nm = stId
            d.Add k, nm
        End If
    Next i
End Function

' BrojZbirne -> Kupac (firma) naziv; fallback na KupacID ako naziv fali.
Private Function KupacNazivZaZbirnu(ByVal dKupId As Object, ByVal dKupNaziv As Object, _
                                    ByVal brojZbirne As String) As String
    Dim kid As String: kid = DictVal(dKupId, brojZbirne)
    Dim nm As String: nm = DictVal(dKupNaziv, kid)
    If Len(nm) > 0 Then KupacNazivZaZbirnu = nm Else KupacNazivZaZbirnu = kid
End Function

Private Function PdvStopa() As Double
    Dim s As Double
    If Not TryParseDouble(GetConfigValue(CFG_PDV_NADOKNADA_STOPA), s) Then s = 0
    If s <= 0 Then s = PDV_NADOKNADA_DEFAULT
    PdvStopa = s
End Function

Private Function NumVal(ByVal v As Variant) As Double
    If IsNumeric(v) Then NumVal = CDbl(v)
End Function

Private Function FmtDate(ByVal v As Variant) As String
    If IsDate(v) Then FmtDate = Format$(CDate(v), "d.m.yyyy")
End Function

Private Function RowYear(ByVal v As Variant) As Integer
    On Error Resume Next
    If IsDate(v) Then RowYear = Year(CDate(v))
End Function

' Napuni rang: svi kooperanti firme sortirani opadajuce po ukupnom iznosu
' otkupnih listova (Sum Kolicina*Cena) u tekucoj godini; bez storniranih.
' RACUN ranga, bez ijedne kontrole. Vraca 1-bazirani 2D niz (n x 4):
'   1 KooperantID | 2 ime | 3 otkupno mesto | 4 iznos
' sortiran opadajuce po iznosu, plus kontrolne sume kroz izlazne parametre.
'
' Izdvojeno iz LoadKoopRang da isti racun mogu da koriste i legacy panel i novi
' ekran (modScrDokumenti), umesto da se agregacija prepisuje na dva mesta.
' Opsezne granice (odN/doN, serijski dani) su Optional: bez njih vazi staro
' pravilo "tekuca godina" (legacy panel i Dokumenti nepromenjeni); ekran
' Izvestaji salje svoj Od-Do, pa rang postuje isti period kao ostale liste.
Public Function KoopRangRows(ByRef rawKg As Double, ByRef rawVal As Double, _
                             ByRef emptyKg As Double, ByRef emptyVal As Double, _
                             Optional ByVal odN As Double = 0, _
                             Optional ByVal doN As Double = 0) As Variant
    On Error GoTo EH
    rawKg = 0: rawVal = 0: emptyKg = 0: emptyVal = 0
    Dim yr As Integer: yr = Year(Date)

    Dim data As Variant: data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function
    data = ExcludeStornirano(data, TBL_OTKUP)
    If IsEmpty(data) Then Exit Function

    Dim cKoop As Long, cId As Long, cDat As Long
    cKoop = GetColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT)
    cId = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    cDat = GetColumnIndex(TBL_OTKUP, COL_OTK_DATUM)
    If cKoop = 0 Or cId = 0 Then Exit Function

    ' Iznos i kilaza dokumenta su na STAVKAMA: CreateOtkup_TX ih na zaglavlju
    ' ostavlja prazne, pa je rang za nov dokument sabirao nulu (REFAKTOR S14.7).
    Dim stavkeZbir As Object: Set stavkeZbir = modOtkup.ZbirStavkiPoOtkupu()

    Dim agg As Object: Set agg = CreateObject("Scripting.Dictionary")
    Dim i As Long, uKrug As Boolean, dSer As Double
    For i = 1 To UBound(data, 1)
        If odN > 0 Or doN > 0 Then
            uKrug = False
            If cDat > 0 Then
                If IsDate(data(i, cDat)) Then
                    dSer = Int(CDbl(CDate(data(i, cDat))))
                    uKrug = (odN = 0 Or dSer >= odN) And (doN = 0 Or dSer <= doN)
                End If
            End If
        Else
            uKrug = (cDat = 0 Or RowYear(data(i, cDat)) = yr)
        End If
        If uKrug Then
            Dim kg As Double, vred As Double, zRang As Variant
            zRang = modOtkup.ZbirStavkiZaOtkup(stavkeZbir, CStr(data(i, cId)), _
                        "modOtkupBlok.KoopRangRows")
            kg = CDbl(zRang(0))
            vred = CDbl(zRang(1))
            rawKg = rawKg + kg
            rawVal = rawVal + vred
            Dim k As String: k = Trim$(CStr(data(i, cKoop)))
            If Len(k) > 0 Then
                If agg.Exists(k) Then
                    agg(k) = CDbl(agg(k)) + vred
                Else
                    agg.Add k, vred
                End If
            Else
                emptyKg = emptyKg + kg
                emptyVal = emptyVal + vred
            End If
        End If
    Next i
    If agg.count = 0 Then Exit Function

    Dim n As Long: n = agg.count
    Dim koopIDs() As String: ReDim koopIDs(1 To n)
    Dim iznosi() As Double: ReDim iznosi(1 To n)
    Dim kk As Variant, j As Long: j = 0
    For Each kk In agg.keys
        j = j + 1
        koopIDs(j) = CStr(kk)
        iznosi(j) = CDbl(agg(kk))
    Next kk
    SortDescByVal koopIDs, iznosi

    Dim dKo As Object: Set dKo = BuildKoopNames()
    Dim dOM As Object: Set dOM = BuildKoopOM()
    Dim outA() As Variant: ReDim outA(1 To n, 1 To 4)
    For i = 1 To n
        Dim nm2 As String: nm2 = DictVal(dKo, koopIDs(i))
        If Len(nm2) = 0 Then nm2 = koopIDs(i)
        outA(i, 1) = koopIDs(i)
        outA(i, 2) = nm2
        outA(i, 3) = DictVal(dOM, koopIDs(i))
        outA(i, 4) = iznosi(i)
    Next i
    KoopRangRows = outA
    Exit Function
EH:
    LogErr "modOtkupBlok.KoopRangRows"
End Function

' Opadajuci sort paralelnih nizova po iznosu (umeren broj kooperanata -> bubble).
Private Sub SortDescByVal(ByRef ids() As String, ByRef vals() As Double)
    Dim i As Long, j As Long, n As Long
    n = UBound(vals)
    For i = 1 To n - 1
        For j = 1 To n - i
            If vals(j) < vals(j + 1) Then
                Dim tv As Double: tv = vals(j): vals(j) = vals(j + 1): vals(j + 1) = tv
                Dim ts As String: ts = ids(j): ids(j) = ids(j + 1): ids(j + 1) = ts
            End If
        Next j
    Next i
End Sub
