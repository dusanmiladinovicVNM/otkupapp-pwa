Attribute VB_Name = "modAmbalaza"

'Attribute VB_Name = "modAmbalaza"
Option Explicit

' ============================================================
' modAmbalaza v2.2 - Verpackung-Tracking / Ambalaza ledger
'
' v2.2 hardening:
' - TrackAmbalaza fail-fast validation
' - AppendRow result check
' - RequireColumnIndex schema guards
' - Strict Smer = Ulaz / Izlaz
' - Safe date filtering in GetVozacAmbSaldo
' - Numeric guards before CLng
'
' NOTE:
' Ne praviti poseban TrackAmbalaza_TX ovde.
' TrackAmbalaza se poziva iz vecih poslovnih TX wrapper-a
' koji vec snapshot-uju tblAmbalaza.
' ============================================================

Private Const AMB_SMER_ULAZ As String = "Ulaz"
Private Const AMB_SMER_IZLAZ As String = "Izlaz"

' --- KNJIGA (AMB-10b): greske koje POZIVALAC mora da razlikuje ---
'
' Potvrda deficita nije kvar nego PITANJE operateru (AMB-10, 6.5). Ekran je mora
' razlikovati od svake druge greske PO BROJU, ne po tekstu poruke: tekst je
' prevodiv i menja se, broj je ugovor.
Public Const AMB_ERR_POTVRDA_DEFICITA As Long = vbObjectError + 4470
Public Const AMB_ERR_DEFICIT_NEPOKRIV As Long = vbObjectError + 4471
Public Const AMB_ERR_IDENTITET As Long = vbObjectError + 4472
Public Const AMB_ERR_JEDAN_PARTNER As Long = vbObjectError + 4475
Public Const AMB_ERR_STORNO As Long = vbObjectError + 4476
Private Const AMB_ERR_KNJIGA_KVAR As Long = vbObjectError + 4473
Private Const AMB_ERR_KNJIGA_ULAZ As Long = vbObjectError + 4474

' ============================================================
' Helpers
' ============================================================

Private Function AmbText(ByVal v As Variant) As String
    If isError(v) Or IsNull(v) Or IsEmpty(v) Then
        AmbText = ""
    Else
        AmbText = Trim$(CStr(v))
    End If
End Function

Private Function IsValidAmbSmer(ByVal smer As String) As Boolean
    Select Case Trim$(smer)
        Case AMB_SMER_ULAZ, AMB_SMER_IZLAZ
            IsValidAmbSmer = True
        Case Else
            IsValidAmbSmer = False
    End Select
End Function

' ============================================================
' Vozac-perspektiva smera ambalaze (jednostran ledger).
'
' Ambalaza se upisuje JEDNOM, entitetski-relativno. Vozac je transporter
' = INVERZNI protivpartner entiteta: sta entitetu UDE (Ulaz), iz vozaca
' IZLAZI, i obrnuto. Definicija "za koga vazi suprotno":
'   - Stanica (OM):    otpremnica / OM-ulaz   -> vozac = inverzno.
'   - Kupac (hladnj.): prijemnica / izlaz-kupci -> vozac = inverzno.
'   - Kooperant (otkup): NEMA vozaca -> izuzet uzvodno (filter / skip).
' Primeri: otpremnica (Stanica Izlaz) -> vozac ULAZ (puni se);
'          prijemnica-pune (Kupac Ulaz) -> vozac IZLAZ (prazni se).
' Tako kompletna ruta daje saldo 0; otvorena otpremnica = pozitivan
' saldo (gajbice jos kod vozaca).
' ============================================================
Public Function VozacAmbEffectiveSmer(ByVal smer As String, _
                                      ByVal entitetTip As String) As String
    Select Case Trim$(entitetTip)
        Case "Stanica", "Kupac"
            ' Transport: vozac je inverzni protivpartner entiteta.
            Select Case Trim$(smer)
                Case AMB_SMER_IZLAZ: VozacAmbEffectiveSmer = AMB_SMER_ULAZ
                Case AMB_SMER_ULAZ:  VozacAmbEffectiveSmer = AMB_SMER_IZLAZ
                Case Else:           VozacAmbEffectiveSmer = smer
            End Select
        Case Else
            ' Kooperant (otkup) nema vozaca -> izuzet uzvodno; ostavi sirovo.
            VozacAmbEffectiveSmer = smer
    End Select
End Function

Private Sub ValidateAmbalazaInput(ByVal tipAmb As String, _
                                  ByVal kolicina As Long, _
                                  ByVal smer As String, _
                                  ByVal entitetID As String, _
                                  ByVal entitetTip As String, _
                                  ByVal sourceName As String)

    If kolicina < 0 Then
        Err.Raise vbObjectError + 4401, sourceName, _
                  "Koli" & ChrW(269) & "ina ambala" & ChrW(382) & "e ne sme biti negativna."
    End If

    ' Nula je legalan no-op.
    If kolicina = 0 Then Exit Sub

    If Len(Trim$(tipAmb)) = 0 Then
        Err.Raise vbObjectError + 4402, sourceName, _
                  "Tip ambala" & ChrW(382) & "e je obavezan kada postoji koli" & ChrW(269) & "ina."
    End If

    If Not IsValidAmbSmer(smer) Then
        Err.Raise vbObjectError + 4403, sourceName, _
                  "Neispravan smer ambala" & ChrW(382) & "e: " & smer
    End If

    If Len(Trim$(entitetID)) = 0 Then
        Err.Raise vbObjectError + 4404, sourceName, _
                  "EntitetID je obavezan za ambala" & ChrW(382) & "u."
    End If

    If Len(Trim$(entitetTip)) = 0 Then
        Err.Raise vbObjectError + 4405, sourceName, _
                  "EntitetTip je obavezan za ambala" & ChrW(382) & "u."
    End If
End Sub

Private Sub RequireAmbalazaSchema(ByVal sourceName As String)
    ' Fail-fast schema guard. Ne koristimo indekse ovde za rowData,
    ' ali eksplicitno proveravamo da tabela ima ocekivane kolone.
    modSchema.SchemaReadyOrFail sourceName, TBL_AMBALAZA
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_ID, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DATUM, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_TIP, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_KOLICINA, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_SMER, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET_TIP, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_VOZAC, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_ID, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_TIP, sourceName)
End Sub

' ============================================================
' WRITE
' ============================================================

' Nov identitet logickog reversa: opaque "RID-<32 hex>" iz centralne fabrike
' NewEntityID. Transakcioni identitet, ne max+1 -- GetNextID ostaje za maticne
' podatke (DOCUMENT_HEADER_LINES.md par. 2). JEDAN po dokumentu: pisac ga kuje
' jednom i daje svim nogama (REV-IDENT-01, ARCHITECTURE_CONTRACT.md).
Public Function NoviReversID() As String
    Const SRC As String = "modAmbalaza.NoviReversID"

    NoviReversID = NewEntityID("RID-")

    If Len(Trim$(NoviReversID)) = 0 Then
        Err.Raise vbObjectError + 4408, SRC, _
                  "NewEntityID nije vratio ReversID."
    End If
End Function

Public Sub TrackAmbalaza(ByVal datum As Date, ByVal tipAmb As String, _
                         ByVal kolicina As Long, ByVal smer As String, _
                         ByVal entitetID As String, ByVal entitetTip As String, _
                         Optional ByVal vozacID As String = "", _
                         Optional ByVal dokumentID As String = "", _
                         Optional ByVal dokumentTip As String = "", _
                         Optional ByVal reversID As String = "")

    Const SRC As String = "modAmbalaza.TrackAmbalaza"

    On Error GoTo EH

    Call ValidateAmbalazaInput(tipAmb, kolicina, smer, entitetID, entitetTip, SRC)

    If kolicina = 0 Then Exit Sub

    Call RequireAmbalazaSchema(SRC)

    Dim newID As String
    newID = GetNextID(TBL_AMBALAZA, COL_AMB_ID, "AMB-")

    If Len(Trim$(newID)) = 0 Then
        Err.Raise vbObjectError + 4406, SRC, _
                  "GetNextID nije vratio AmbID."
    End If

    Dim rowData As Variant
    rowData = Array( _
        newID, _
        datum, _
        Trim$(tipAmb), _
        kolicina, _
        Trim$(smer), _
        Trim$(entitetID), _
        Trim$(entitetTip), _
        Trim$(vozacID), _
        Trim$(dokumentID), _
        Trim$(dokumentTip))

    Dim rowIdx As Long
    rowIdx = AppendRow(TBL_AMBALAZA, rowData)
    If rowIdx <= 0 Then
        Err.Raise vbObjectError + 4407, SRC, _
                  "AppendRow nije uspeo za tblAmbalaza."
    End If

    ' ReversID (REV-IDENT-01) ide PO IMENU: kolona stoji iza Stornirano i audit
    ' kolona, pa je pozicioni niz ne doseze. Pecat je u istom modulu, pa vlasnik
    ' reda ostaje jedan (A11).
    If Len(Trim$(reversID)) > 0 Then
        RequireUpdateCell TBL_AMBALAZA, rowIdx, COL_AMB_REVERS_ID, Trim$(reversID), SRC
    End If

    Exit Sub

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    LogErr SRC
    On Error Resume Next
    On Error GoTo 0

    Err.Raise errNum, SRC, "Source=" & errSrc & " | " & errDesc
End Sub

' ============================================================
' READ: saldo ambalaze po entitetu
'
' Returns:
'   result(row, 1) = TipAmbalaze
'   result(row, 2) = Saldo
'
' Semantika:
'   Ulaz  => +Kolicina
'   Izlaz => -Kolicina
' ============================================================

Public Function GetAmbalazeStanje(ByVal entitetID As String, _
                                  ByVal entitetTip As String) As Variant
    Const SRC As String = "modAmbalaza.GetAmbalazeStanje"

    On Error GoTo EH

    If Len(Trim$(entitetID)) = 0 Or Len(Trim$(entitetTip)) = 0 Then
        GetAmbalazeStanje = Empty
        Exit Function
    End If

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

    Dim colTip As Long
    Dim colKol As Long
    Dim colSmer As Long
    Dim colEntID As Long
    Dim colEntTip As Long

    colTip = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_TIP, SRC)
    colKol = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_KOLICINA, SRC)
    colSmer = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_SMER, SRC)
    colEntID = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET, SRC)
    colEntTip = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET_TIP, SRC)

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If AmbText(data(i, colEntID)) = Trim$(entitetID) And _
           AmbText(data(i, colEntTip)) = Trim$(entitetTip) Then

            Dim key As String
            key = AmbText(data(i, colTip))

            If Len(key) = 0 Then
                Err.Raise vbObjectError + 4411, SRC, _
                          "Prazan TipAmbalaze u redu " & CStr(i)
            End If

            If Not IsNumeric(data(i, colKol)) Then
                Err.Raise vbObjectError + 4410, SRC, _
                          "Neispravna koli" & ChrW(269) & "ina ambala" & ChrW(382) & "e u redu " & CStr(i)
            End If

            If Not dict.Exists(key) Then dict.Add key, 0&

            Select Case AmbText(data(i, colSmer))
                Case AMB_SMER_ULAZ
                    dict(key) = CLng(dict(key)) + CLng(data(i, colKol))

                Case AMB_SMER_IZLAZ
                    dict(key) = CLng(dict(key)) - CLng(data(i, colKol))

                Case Else
                    Err.Raise vbObjectError + 4412, SRC, _
                              "Neispravan smer ambala" & ChrW(382) & "e u redu " & CStr(i)
            End Select
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
    Exit Function

EH:
    LogErr SRC
    GetAmbalazeStanje = Empty
End Function

' ============================================================
' READ: pocetno stanje ambalaze kooperanta PRE datog bloka.
'
' "Pre bloka" = svi redovi tog kooperanta (Ulaz +, Izlaz -) za dati
' tipAmb upisani PRE prvog reda ovog bloka (po redosledu upisa u
' append-only tabeli). Blok se identifikuje preko DokumentID-a iz
' blockOtkupIDs (obe noge - Otkup i OM-Izlaz-Koop - dele DokumentID =
' otkupID). Tako je ispravno i kod ponovne stampe starijeg bloka:
' kasniji blokovi se NE uracunavaju u pocetno stanje.
'
' Fallback: ako se nijedan red bloka ne nadje (npr. legacy red pre
' dvojnog upisa), uzima SVE kooperantove redove tog tipa kao pocetno.
'
' blockOtkupIDs: niz (npr. Split rezultat) ili string "OTK-1 + OTK-2".
' Vraca Long (entitetski saldo: Ulaz = +Kolicina, Izlaz = -Kolicina).
' ============================================================
Public Function GetKooperantAmbOpening(ByVal koopID As String, _
                                       ByVal tipAmb As String, _
                                       ByVal blockOtkupIDs As Variant) As Long
    Const SRC As String = "modAmbalaza.GetKooperantAmbOpening"

    On Error GoTo EH

    GetKooperantAmbOpening = 0

    If Len(Trim$(koopID)) = 0 Or Len(Trim$(tipAmb)) = 0 Then Exit Function

    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then Exit Function

    data = ExcludeStornirano(data, TBL_AMBALAZA)
    If IsEmpty(data) Then Exit Function

    Dim colTip As Long, colKol As Long, colSmer As Long
    Dim colEntID As Long, colEntTip As Long, colDokID As Long
    colTip = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_TIP, SRC)
    colKol = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_KOLICINA, SRC)
    colSmer = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_SMER, SRC)
    colEntID = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET, SRC)
    colEntTip = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET_TIP, SRC)
    colDokID = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_ID, SRC)

    ' Skup otkup-ID-jeva ovog bloka (niz ili string "OTK-1 + OTK-2").
    Dim blk As Object
    Set blk = CreateObject("Scripting.Dictionary")
    blk.CompareMode = vbTextCompare

    If IsArray(blockOtkupIDs) Then
        Dim el As Variant
        For Each el In blockOtkupIDs
            If Len(Trim$(CStr(el))) > 0 Then blk(Trim$(CStr(el))) = True
        Next el
    Else
        Dim parts() As String
        parts = Split(CStr(blockOtkupIDs), " + ")
        Dim p As Long
        For p = LBound(parts) To UBound(parts)
            If Len(Trim$(parts(p))) > 0 Then blk(Trim$(parts(p))) = True
        Next p
    End If

    ' 1) Prvi (najraniji) red ovog bloka za ovog kooperanta = granica "pre bloka".
    Dim i As Long
    Dim minIdx As Long: minIdx = 0
    If blk.count > 0 Then
        For i = 1 To UBound(data, 1)
            If AmbText(data(i, colEntTip)) = "Kooperant" And _
               AmbText(data(i, colEntID)) = Trim$(koopID) Then
                If blk.Exists(AmbText(data(i, colDokID))) Then
                    minIdx = i
                    Exit For
                End If
            End If
        Next i
    End If

    ' 2) Saldo svih redova PRE granice (ili svih, ako blok nije nadjen -> fallback).
    Dim cutoff As Long
    If minIdx > 0 Then
        cutoff = minIdx - 1
    Else
        cutoff = UBound(data, 1)
    End If

    Dim saldo As Long: saldo = 0
    For i = 1 To cutoff
        If AmbText(data(i, colEntTip)) = "Kooperant" And _
           AmbText(data(i, colEntID)) = Trim$(koopID) And _
           AmbText(data(i, colTip)) = Trim$(tipAmb) Then

            If IsNumeric(data(i, colKol)) Then
                Select Case AmbText(data(i, colSmer))
                    Case AMB_SMER_ULAZ:  saldo = saldo + CLng(data(i, colKol))
                    Case AMB_SMER_IZLAZ: saldo = saldo - CLng(data(i, colKol))
                End Select
            End If
        End If
    Next i

    GetKooperantAmbOpening = saldo
    Exit Function

EH:
    LogErr SRC
    GetKooperantAmbOpening = 0
End Function

' ============================================================
' READ: entitetski saldo gajbica STANICE (OM) za dati tip ambalaze.
'
' Pandan GetKooperantAmbOpening, ali za entitet "Stanica" (puni se iz
' otpremnica / OM-ulaz). Vraca pun trenutni saldo (Ulaz +, Izlaz -). Za
' grupni otkupni list to je "pocetno stanje" jer prijemnica pise "Kupac"
' stranu i ne pomera "Stanica" entitet. 0 ako nema pokreta / pri gresci.
' ============================================================
Public Function GetStanicaAmbSaldo(ByVal stanicaID As String, _
                                   ByVal tipAmb As String) As Long
    Const SRC As String = "modAmbalaza.GetStanicaAmbSaldo"

    On Error GoTo EH

    GetStanicaAmbSaldo = 0
    If Len(Trim$(stanicaID)) = 0 Or Len(Trim$(tipAmb)) = 0 Then Exit Function

    Dim st As Variant
    st = GetAmbalazeStanje(stanicaID, "Stanica")
    If Not IsArray(st) Then Exit Function

    Dim i As Long
    For i = LBound(st, 1) To UBound(st, 1)
        If AmbText(st(i, 1)) = Trim$(tipAmb) Then
            If IsNumeric(st(i, 2)) Then GetStanicaAmbSaldo = CLng(st(i, 2))
            Exit Function
        End If
    Next i
    Exit Function

EH:
    LogErr SRC
    GetStanicaAmbSaldo = 0
End Function

' ============================================================
' READ: saldo ambalaze po vozacu
'
' Returns:
'   result(row, 1) = TipAmbalaze
'   result(row, 2) = Izlaz
'   result(row, 3) = Ulaz
'   result(row, 4) = Saldo = Izlaz - Ulaz
'
' Semantika:
'   Vozac saldo se racuna iz svih aktivnih pokreta gde je VozacID isti.
'   Ne filtrira se po DokumentTip-u.
' ============================================================

Public Function GetVozacAmbSaldo(ByVal vozacID As String, _
                                  Optional ByVal datumOd As Date = 0, _
                                  Optional ByVal datumDo As Date = 0) As Variant
    Const SRC As String = "modAmbalaza.GetVozacAmbSaldo"

    On Error GoTo EH

    If Len(Trim$(vozacID)) = 0 Then
        GetVozacAmbSaldo = Empty
        Exit Function
    End If

    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA)

    If IsEmpty(data) Then
        GetVozacAmbSaldo = Empty
        Exit Function
    End If

    data = ExcludeStornirano(data, TBL_AMBALAZA)

    If IsEmpty(data) Then
        GetVozacAmbSaldo = Empty
        Exit Function
    End If

    Dim colTip As Long
    Dim colKol As Long
    Dim colSmer As Long
    Dim colVozac As Long
    Dim colDatum As Long
    Dim colEntTip As Long
    Dim colDokTip As Long

    colTip = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_TIP, SRC)
    colKol = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_KOLICINA, SRC)
    colSmer = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_SMER, SRC)
    colVozac = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_VOZAC, SRC)
    colDatum = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DATUM, SRC)
    colEntTip = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET_TIP, SRC)
    colDokTip = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_TIP, SRC)

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If AmbText(data(i, colVozac)) = Trim$(vozacID) Then

            ' Otkup (Kooperant-nabavka) nije vozaceva transportna noga -> izuzmi
            ' (inace dupli teret: otkup + otpremnica iste gajbice). Saldo vozaca =
            ' otpremnica (utovar) - prijemnica (predaja).
            If AmbText(data(i, colDokTip)) = DOK_TIP_OTKUP Then GoTo NextRow

            If datumOd > 0 Or datumDo > 0 Then
                If Not IsDate(data(i, colDatum)) Then GoTo NextRow

                Dim d As Date
                d = CDate(data(i, colDatum))

                If datumOd > 0 And d < datumOd Then GoTo NextRow
                If datumDo > 0 And d > datumDo Then GoTo NextRow
            End If

            Dim key As String
            key = AmbText(data(i, colTip))

            If Len(key) = 0 Then
                Err.Raise vbObjectError + 4421, SRC, _
                          "Prazan TipAmbalaze u redu " & CStr(i)
            End If

            If Not IsNumeric(data(i, colKol)) Then
                Err.Raise vbObjectError + 4420, SRC, _
                          "Neispravna koli" & ChrW(269) & "ina ambala" & ChrW(382) & "e u redu " & CStr(i)
            End If

            If Not dict.Exists(key) Then dict.Add key, Array(0&, 0&) ' Izlaz, Ulaz

            Dim vals As Variant
            vals = dict(key)

            ' Vozac = inverzni protivpartner entiteta (Stanica / Kupac); ruta
            ' otpremnica -> prijemnica se netira na 0. Otkup nema vozaca (skip gore).
            Dim effSmer As String
            effSmer = VozacAmbEffectiveSmer(AmbText(data(i, colSmer)), AmbText(data(i, colEntTip)))

            Select Case effSmer
                Case AMB_SMER_IZLAZ
                    vals(0) = CLng(vals(0)) + CLng(data(i, colKol))

                Case AMB_SMER_ULAZ
                    vals(1) = CLng(vals(1)) + CLng(data(i, colKol))

                Case Else
                    Err.Raise vbObjectError + 4422, SRC, _
                              "Neispravan smer ambala" & ChrW(382) & "e u redu " & CStr(i)
            End Select

            dict(key) = vals
        End If

NextRow:
    Next i

    If dict.count = 0 Then
        GetVozacAmbSaldo = Empty
        Exit Function
    End If

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
    Exit Function

EH:
    LogErr SRC
    GetVozacAmbSaldo = Empty
End Function

' ============================================================
' READ: opcije za "Tip ambalaze" combo (frmOtkup / frmDokumenta)
'
' Vraca tipove koje je kupac uneo u tblTipAmbalaze (maticni podaci).
' Fallback na ugradjene konstante (12/1, 6/1) ako je sifarnik prazan
' ili tabela ne postoji -- da postojece instalacije ne ostanu bez opcija.
' ============================================================
Public Function GetTipAmbalazeOptions() As Variant
    On Error GoTo Fallback

    Dim arr As Variant
    arr = GetLookupList(TBL_TIP_AMBALAZE, COL_TAMB_TIP, , , True)

    If IsArray(arr) Then
        If (UBound(arr) - LBound(arr) + 1) > 0 Then
            GetTipAmbalazeOptions = arr
            Exit Function
        End If
    End If

Fallback:
    GetTipAmbalazeOptions = Array(AMB_12_1, AMB_6_1)
End Function

' Podrazumevani tip ambalaze za kulturu (tblKulture.TipAmbalaze).
' Match po VrstaVoca (+ SortaVoca ako je data). Vraca "" ako nema.
' Koristi se za auto-popunjavanje tipa ambalaze u frmOtkup/frmDokumenta.
Public Function GetKulturaTipAmbalaze(ByVal vrsta As String, ByVal sorta As String) As String
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_KULTURE)
    If IsEmpty(data) Then Exit Function

    Dim cv As Long, cS As Long, cT As Long
    cv = GetColumnIndex(TBL_KULTURE, "VrstaVoca")
    cS = GetColumnIndex(TBL_KULTURE, "SortaVoca")
    cT = GetColumnIndex(TBL_KULTURE, COL_KUL_TIP_AMBALAZE)
    If cv = 0 Or cT = 0 Then Exit Function

    Dim i As Long, hit As String
    For i = 1 To UBound(data, 1)
        If StrComp(Trim$(nz(data(i, cv))), Trim$(vrsta), vbTextCompare) = 0 Then
            If Len(Trim$(sorta)) = 0 Or cS = 0 _
               Or StrComp(Trim$(nz(data(i, cS))), Trim$(sorta), vbTextCompare) = 0 Then
                hit = Trim$(nz(data(i, cT)))
                If Len(hit) > 0 Then
                    GetKulturaTipAmbalaze = hit
                    Exit Function
                End If
            End If
        End If
    Next i
    Exit Function

EH:
    GetKulturaTipAmbalaze = ""
End Function



' ============================================================
' KNJIGA AMBALAZE -- PISAC (AMB-10b)
' ============================================================
'
' Dogadjaj je PRENOS: jedan red imenuje OBE strane (OdNalog -> NaNalog), kolicina
' je UVEK pozitivna, smera kao podatka nema. Knjiga je APPEND-ONLY -- pogresan
' unos se ne popravlja UPDATE-om nego stornom i novim dogadjajem. Pun model:
' docs/DOMEN/AMBALAZA.md 6.1-6.9.
'
' Ugovor (zatvorene liste, matrica strana, razresavanje naloga, doprinos obavezi,
' veza dokument <-> kretanje) je modAmbalazaUgovor. Ovde je UPIS i ono sto se MORA
' procitati da bi upis bio zakonit -- nista vise.
'
' OVAJ REZ NE DIRA NIJEDNO POZIVNO MESTO. TrackAmbalaza i dalje pise stari oblik;
' devet mesta knjizenja, razlaganje SaveOMUlaz_TX / SaveKupciIzlaz_TX i citaoci su
' 10b-2. Zato tblAmbalaza tokom prelaza nosi DVA oblika reda i svaki citalac mora
' da kaze koji cita -- to radi KnjigaZaCitanje, i radi to FAIL-CLOSED.
'
' ZASTO PISAC CITA STANJE: AMB-INV-07 (nijedan realan nalog ispod nule) i
' AMB-INV-09 (obaveza >= 0) su granice koje se bez stanja ne mogu proveriti. UI
' sme da ih prikaze unapred, ali racun koji vazi je ovaj, u trenutku upisa --
' izmedju pitanja i odgovora stanje se moglo promeniti drugim unosom (6.5). Isti
' obrazac koji repo vec drzi kod ApplyAvansToOtkup / IsplataBlokProblem.
'
' ZASTO OVDE NEMA "On Error GoTo EH / LogErr" BLOKA, kao u TrackAmbalaza:
' odbijanje pisca je najcesce POSLOVNI ishod, ne kvar. Potvrda deficita je
' pitanje operateru, a ne greska -- da svaki upit zavrsi u logu gresaka, log bi
' prestao da bude signal. Greske se zato ne gutaju i ne prepakuju nego dizu dalje,
' sa svojim brojem.
'
' STA PISAC NE RADI:
'   STORNO       -- kontra-stav sa StornoOd. PrenesiAmbalazu upisuje samo
'                   ORIGINALE, pa StornoOd ostaje prazan; ulaz za storno je
'                   StornirajAmbalazuDokumenta (AMB-10-ODL-16).
'   AMB-INV-08   -- "upis u istoj transakciji sa izvornim dokumentom" se ne moze
'                   dokazati iznutra: clsTransaction nema globalan registar aktivne
'                   transakcije. To je STATICKA kapija nad pozivnim mestima, a njih
'                   u ovom rezu nema -- ide uz njih, u 10b-2.
'   BROJ          -- numericki niz ambalaznog dokumenta (koji modBrojevi kind, koji
'                   kontekst, i prozor jedinstvenosti) vezan je za pozivna mesta:
'                   stari revers broji po (stanica, dan), a revers kupca stanicu
'                   nema. Odluka ide uz cutover; ovde se broj samo zahteva kao
'                   neprazan, uz vlasnika niza (AMB-10-DOK).

Private Sub RequireKnjigaSchema(ByVal sourceName As String)
    modSchema.SchemaReadyOrFail sourceName, TBL_AMBALAZA

    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_ID, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DATUM, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_TIP, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_KOLICINA, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_OD_TIP, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_OD_ID, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_NA_TIP, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_NA_ID, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_ID, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_TIP, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_VRSTA_KRETANJA, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_STORNO_OD, sourceName)
End Sub

Private Sub RequireAmbDokSchema(ByVal sourceName As String)
    modSchema.SchemaReadyOrFail sourceName, TBL_AMBALAZA_DOKUMENT

    Call RequireColumnIndex(TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA_DOKUMENT, COL_AMBD_VRSTA, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA_DOKUMENT, COL_AMBD_BROJ, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA_DOKUMENT, COL_AMBD_DATUM, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA_DOKUMENT, COL_AMBD_BROJ_OWNER_TIP, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA_DOKUMENT, COL_AMBD_BROJ_OWNER_ID, sourceName)
End Sub

' Indeksi kolona knjige, JEDNOM po citaocu: ime -> indeks. Svaki citalac je ranije
' nosio svoj blok RequireColumnIndex poziva, a ugovor zapisanog reda trazi SVE
' kolone -- jedanaest parametara po pozivu bilo bi necitljivo i lako se razilazi.
Private Function KnjigaIndeksi(ByVal sourceName As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    Dim imena As Variant, j As Long
    ' Kolone STAROG modela su ovde jer citalac mora da proveri i NJEGOV ugovor:
    ' "nije nov red" nije isto sto i "valjan star red". Odlaze zajedno sa njima u
    ' 10e, kad stari model nestane iz kanona.
    imena = Array(COL_AMB_ID, COL_AMB_DATUM, COL_AMB_TIP, COL_AMB_KOLICINA, _
                  COL_AMB_OD_TIP, COL_AMB_OD_ID, COL_AMB_NA_TIP, COL_AMB_NA_ID, _
                  COL_AMB_DOK_TIP, COL_AMB_DOK_ID, COL_AMB_VRSTA_KRETANJA, _
                  COL_AMB_STORNO_OD, _
                  COL_AMB_SMER, COL_AMB_ENTITET, COL_AMB_ENTITET_TIP)

    For j = LBound(imena) To UBound(imena)
        d.Add CStr(imena(j)), RequireColumnIndex(TBL_AMBALAZA, CStr(imena(j)), sourceName)
    Next j

    Set KnjigaIndeksi = d
End Function

' TRI STANJA REDA, NE DVA (review #400, P2).
'
' Red DOTICE knjigu kad je BILO KOJA nova kolona popunjena. Prva verzija je gledala
' samo OdNalogTip, pa je red sa praznim izvorom a popunjenim odredistem prolazio kao
' "stari oblik" -- tiho, i van svakog salda.
Private Function RedDoticeKnjigu(ByRef data As Variant, ByVal i As Long, _
                                 ByRef kol As Object) As Boolean
    Dim imena As Variant, j As Long
    imena = Array(COL_AMB_OD_TIP, COL_AMB_OD_ID, COL_AMB_NA_TIP, COL_AMB_NA_ID, _
                  COL_AMB_VRSTA_KRETANJA, COL_AMB_STORNO_OD)

    For j = LBound(imena) To UBound(imena)
        If Len(AmbText(data(i, kol(CStr(imena(j)))))) > 0 Then
            RedDoticeKnjigu = True
            Exit Function
        End If
    Next j
End Function

' PRAZAN RED NIJE RED. Prazna Excel tabela ima jedan prazan red u DataBodyRange,
' pa bi ga svaka kapija inace prijavila kao kvar na praznoj knjizi. Meri se po SVIM
' kolonama koje knjiga uopste cita -- ni jedno polje, ni staro ni novo.
Private Function RedPrazan(ByRef data As Variant, ByVal i As Long, _
                           ByRef kol As Object) As Boolean
    Dim k As Variant
    For Each k In kol.Keys
        If Len(AmbText(data(i, kol(CStr(k))))) > 0 Then Exit Function
    Next k

    RedPrazan = True
End Function

' UGOVOR STAROG REDA -- jer "nije nov" nije isto sto i "valjan star".
'
' Red koji ne dotice knjigu preskace se kao legacy. Ako mu pritom ne vazi ni stari
' ugovor, on ne pripada NIJEDNOM modelu: stari citalac ga ne vidi (entitet se ne
' poklapa), novi ga preskace -- a kolicina koju nosi nestaje iz svakog salda bez
' ijedne poruke (review #400, treci krug).
'
' Ugovor je onaj koji stari pisac vec drzi (ValidateAmbalazaInput): smer je Ulaz ili
' Izlaz, entitet ima tip i ID, tip ambalaze postoji, kolicina je broj veci od nule.
' Odlazi zajedno sa starim modelom u 10e.
Private Function LegacyRedProblem(ByRef data As Variant, ByVal i As Long, _
                                  ByRef kol As Object) As String
    If Len(AmbText(data(i, kol(COL_AMB_ID)))) = 0 Then
        LegacyRedProblem = "nema AmbID"
        Exit Function
    End If

    If Not IsValidAmbSmer(AmbText(data(i, kol(COL_AMB_SMER)))) Then
        LegacyRedProblem = "nije ni red knjige ni valjan stari red -- nema Smer"
        Exit Function
    End If

    If Len(AmbText(data(i, kol(COL_AMB_ENTITET_TIP)))) = 0 Or _
       Len(AmbText(data(i, kol(COL_AMB_ENTITET)))) = 0 Then
        LegacyRedProblem = "stari red bez entiteta"
        Exit Function
    End If

    If Len(AmbText(data(i, kol(COL_AMB_TIP)))) = 0 Then
        LegacyRedProblem = "stari red bez tipa ambalaze"
        Exit Function
    End If

    If Not IsNumeric(data(i, kol(COL_AMB_KOLICINA))) Then
        LegacyRedProblem = "stari red sa nenumerickom kolicinom"
        Exit Function
    End If

    If CDbl(data(i, kol(COL_AMB_KOLICINA))) <= 0 Then
        LegacyRedProblem = "stari red sa kolicinom <= 0"
    End If
End Function

' UGOVOR ZAPISANOG REDA -- jedno mesto za SVE citaoce (saldo, obaveza,
' idempotencija, identitet, i storno u 10d).
'
' Zasto nad svakim citanjem a ne samo pri upisu: saldo ulazi u kapiju deficita, pa
' pokvaren ZAPISAN red menja odluku SLEDECEG upisa. Red uracunat pola-pola (izvor
' bez ID-a) razbija ocuvanje kolicine bez ijedne poruke -- partner dobije +20, a
' nijedan stvarni nalog ne dobije -20.
'
' Postojanje naloga u maticnoj tabeli se OVDE NE proverava: to je kapija UPISA
' (AmbNalogProblem u RequireAmbPrenos). Da citalac to radi, obrisan maticni red bi
' retroaktivno oborio svako citanje, a saldo bi postao kvadratan nad tabelom.
'
' STORNO SE MERI INVERZNO. Kontra-stav ima zamenjene strane, pa bi matrica za
' njegovu vrstu pala na ISPRAVNOM redu (inverz ULAZ_TUDJE je PARTNER -> GRANICA).
' Zamena argumenata JE provera inverza -- bez druge matrice i bez izuzetka.
Private Function KnjigaRedProblem(ByRef data As Variant, ByVal i As Long, _
                                  ByRef kol As Object) As String
    If Len(AmbText(data(i, kol(COL_AMB_ID)))) = 0 Then
        KnjigaRedProblem = "nema AmbID"
        Exit Function
    End If

    If Not IsDate(data(i, kol(COL_AMB_DATUM))) Then
        KnjigaRedProblem = "nema datum"
        Exit Function
    End If

    If Len(AmbText(data(i, kol(COL_AMB_DOK_TIP)))) = 0 Or _
       Len(AmbText(data(i, kol(COL_AMB_DOK_ID)))) = 0 Then
        KnjigaRedProblem = "nema identitet dokumenta"
        Exit Function
    End If

    If Not IsNumeric(data(i, kol(COL_AMB_KOLICINA))) Then
        KnjigaRedProblem = "kolicina nije broj"
        Exit Function
    End If

    Dim odTip As String, odID As String, naTip As String, naID As String
    Dim kolicina As Double, tipAmb As String, vrsta As String, p As String

    odTip = AmbText(data(i, kol(COL_AMB_OD_TIP)))
    odID = AmbText(data(i, kol(COL_AMB_OD_ID)))
    naTip = AmbText(data(i, kol(COL_AMB_NA_TIP)))
    naID = AmbText(data(i, kol(COL_AMB_NA_ID)))
    kolicina = CDbl(data(i, kol(COL_AMB_KOLICINA)))
    tipAmb = AmbText(data(i, kol(COL_AMB_TIP)))
    vrsta = AmbText(data(i, kol(COL_AMB_VRSTA_KRETANJA)))

    If Len(AmbText(data(i, kol(COL_AMB_STORNO_OD)))) > 0 Then
        p = modAmbalazaUgovor.AmbPrenosStrukturaProblem(naTip, naID, odTip, odID, _
                                                        kolicina, tipAmb, vrsta)
        If Len(p) > 0 Then KnjigaRedProblem = "storno: " & p
        Exit Function
    End If

    KnjigaRedProblem = modAmbalazaUgovor.AmbPrenosStrukturaProblem(odTip, odID, naTip, naID, _
                                                                  kolicina, tipAmb, vrsta)
End Function

' INTEGRITET KNJIGE -- jedan prolaz, pa SVAKI citalac dobija isti odgovor.
'
' Vraca mapu AmbID -> VrstaKretanja; ta mapa je i sama dokaz jedinstvenosti
' identiteta, a citaocu obaveze treba za vrstu ORIGINALA kad naidje na kontra-stav.
'
' JEDINSTVENOST AmbID-a NIJE SVOJSTVO REDA NEGO KNJIGE, pa je KnjigaRedProblem ne
' moze proveriti. Ranije je stajala samo u citaocu obaveze (review #400, drugi
' krug), pa je AmbSaldoNaloga sabirao dva reda sa istim stabilnim identitetom:
'
'   AMB-X  SpoljniSvet -> Stanica  100  NABAVKA
'   AMB-X  SpoljniSvet -> Stanica  100  NABAVKA       -> Stanica = 200
'
' Taj saldo ulazi u kapiju deficita, pa korumpiran identitet otvara izlaz BEZ
' pokrica. Za 10d je gore: StornoOd = AMB-X vise ne pokazuje na jedan original.
' Pravilo je isto kao za nalog (AmbNalogProblem): 0 pada, 1 prolazi, 2+ pada.
Private Function KnjigaIntegritet(ByRef data As Variant, ByRef kol As Object, _
                                  ByVal sourceName As String) As Object
    Dim vrste As Object
    Set vrste = CreateObject("Scripting.Dictionary")
    Set KnjigaIntegritet = vrste

    If IsEmpty(data) Then Exit Function

    Dim i As Long, p As String, ambID As String
    For i = 1 To UBound(data, 1)
        If RedPrazan(data, i, kol) Then
            ' artefakt prazne tabele -- nije red
        ElseIf Not RedDoticeKnjigu(data, i, kol) Then
            ' Ne dotice knjigu: sme da bude samo VALJAN stari red. Sve ostalo ne
            ' pripada nijednom modelu i tiho bi nestalo iz svakog salda.
            p = LegacyRedProblem(data, i, kol)
            If Len(p) > 0 Then
                Err.Raise AMB_ERR_KNJIGA_KVAR, sourceName, _
                          "Red " & CStr(i) & ": " & p
            End If
        Else
            p = KnjigaRedProblem(data, i, kol)
            If Len(p) > 0 Then
                Err.Raise AMB_ERR_KNJIGA_KVAR, sourceName, _
                          "Red knjige " & CStr(i) & ": " & p
            End If

            ambID = AmbText(data(i, kol(COL_AMB_ID)))
            If vrste.Exists(ambID) Then
                Err.Raise AMB_ERR_KNJIGA_KVAR, sourceName, _
                          "Dva reda knjige nose AmbID '" & ambID & "'."
            End If
            ' DODELA, ne .Add: Add bi na postojecem kljucu pukao sam od sebe i time
            ' bio SLUCAJNA druga brana ispred imenovane provere iznad. Dvosmerni
            ' dokaz je to i pokazao -- sa ugasenom imenovanom kapijom test je i
            ' dalje padao, samo sa tudjom porukom. Kapija sme biti samo jedna.
            vrste(ambID) = AmbText(data(i, kol(COL_AMB_VRSTA_KRETANJA)))
        End If
    Next i

    ' StornoOd mora da pokazuje na postojeci red -- DRUGI prolaz, jer original sme
    ' da stoji posle svog kontra-stava (knjiga je append-only, ne sortirana).
    Dim st As String
    For i = 1 To UBound(data, 1)
        If RedDoticeKnjigu(data, i, kol) Then
            st = AmbText(data(i, kol(COL_AMB_STORNO_OD)))
            If Len(st) > 0 Then
                If Not vrste.Exists(st) Then
                    Err.Raise AMB_ERR_KNJIGA_KVAR, sourceName, _
                              "StornoOd '" & st & "' ne pokazuje na red knjige."
                End If
            End If
        End If
    Next i
End Function

' JEDAN ULAZ ZA CITANJE KNJIGE: indeksi I dokazan integritet, nikad odvojeno.
'
' Spojeni su namerno. Da `KnjigaIndeksi` ostane javno dostupan, nov citalac bi uzeo
' indekse i preskocio proveru -- a tacno taj oblik propusta je i bio P2: citalac
' obaveze je proveravao identitet, citalac salda nije.
'
' CENA JE JEDAN PROLAZ PO POZIVU CITAOCA, i to je svesno. Nad fixture-om se ne meri;
' nad pravom knjigom ce 10c morati da izmeri, a ako zaboli, odgovor je kes po
' transakciji -- ne slabije citanje.
Private Function KnjigaZaCitanje(ByRef data As Variant, ByVal sourceName As String, _
                                 ByRef outVrste As Object) As Object
    Dim kol As Object
    Set kol = KnjigaIndeksi(sourceName)
    Set outVrste = KnjigaIntegritet(data, kol, sourceName)
    Set KnjigaZaCitanje = kol
End Function

' Nalog kao kljuc, i neuredjen PAR naloga kao kljuc.
'
' Par je NEUREDJEN jer jedan revers sme da nosi i izdavanje i povrat prema istom
' partneru: Stanica -> K1 i K1 -> Stanica su isti poslovni par.
Private Function NalogKljuc(ByVal tip As String, ByVal id As String) As String
    NalogKljuc = Trim$(tip) & ":" & Trim$(id)
End Function

Private Function ParKljuc(ByVal tipA As String, ByVal idA As String, _
                          ByVal tipB As String, ByVal idB As String) As String
    Dim a As String, b As String
    a = NalogKljuc(tipA, idA)
    b = NalogKljuc(tipB, idB)

    If StrComp(a, b, vbTextCompare) <= 0 Then
        ParKljuc = a & " <-> " & b
    Else
        ParKljuc = b & " <-> " & a
    End If
End Function

Private Function IstiNalog(ByVal tipA As String, ByVal idA As String, _
                           ByVal tipB As String, ByVal idB As String) As Boolean
    IstiNalog = (StrComp(Trim$(tipA), Trim$(tipB), vbTextCompare) = 0) And _
                (StrComp(Trim$(idA), Trim$(idB), vbTextCompare) = 0)
End Function

' DAN, ne trenutak: knjiga ambalaze je dnevna (broj reversa je niz po danu), a
' celija moze da nosi i vreme. Poredjenje po trenutku bi isti dogadjaj unet dva
' puta istog dana prijavilo kao dva.
Private Function IstiDan(ByVal a As Variant, ByVal b As Date) As Boolean
    If Not IsDate(a) Then Exit Function
    IstiDan = (Int(CDate(a)) = Int(b))
End Function

' SALDO NALOGA IZ KNJIGE: +kolicina kad je nalog ODREDISTE, -kolicina kad je
' IZVOR. Storno je kontra-stav sa zamenjenim stranama, pa se gasi istim pravilom
' -- bez posebnog slucaja i bez kolone Stornirano.
'
' NEMA "On Error -> vrati 0". Zatecen GetStanicaAmbSaldo tako radi i to je
' fail-open koji ovde ne sme da postoji: na ovaj broj se oslanja kapija deficita,
' pa bi progutana greska proizvela upis BEZ pokrica -- tj. negativan saldo.
Public Function AmbSaldoNaloga(ByVal tip As String, ByVal id As String, _
                               ByVal tipAmb As String) As Double
    Const SRC As String = "modAmbalaza.AmbSaldoNaloga"

    Dim p As String
    p = modAmbalazaUgovor.AmbNalogProblem(tip, id)
    If Len(p) > 0 Then Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, p

    If Len(Trim$(tipAmb)) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                  "Saldo se vodi PO TIPU ambalaze -- tip je obavezan."
    End If

    RequireKnjigaSchema SRC

    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then Exit Function

    Dim cOdTip As Long, cOdID As Long, cNaTip As Long, cNaID As Long
    Dim cVK As Long, cTipA As Long, cKol As Long
    Dim kol As Object, vrste As Object
    Set kol = KnjigaZaCitanje(data, SRC, vrste)

    cOdTip = kol(COL_AMB_OD_TIP)
    cOdID = kol(COL_AMB_OD_ID)
    cNaTip = kol(COL_AMB_NA_TIP)
    cNaID = kol(COL_AMB_NA_ID)
    cVK = kol(COL_AMB_VRSTA_KRETANJA)
    cTipA = kol(COL_AMB_TIP)
    cKol = kol(COL_AMB_KOLICINA)

    Dim i As Long, saldo As Double
    For i = 1 To UBound(data, 1)
        If RedDoticeKnjigu(data, i, kol) Then
            If StrComp(AmbText(data(i, cTipA)), Trim$(tipAmb), vbTextCompare) = 0 Then
                If IstiNalog(AmbText(data(i, cNaTip)), AmbText(data(i, cNaID)), tip, id) Then
                    saldo = saldo + CDbl(data(i, cKol))
                ElseIf IstiNalog(AmbText(data(i, cOdTip)), AmbText(data(i, cOdID)), tip, id) Then
                    saldo = saldo - CDbl(data(i, cKol))
                End If
            End If
        End If
    Next i

    AmbSaldoNaloga = saldo
End Function

' OBAVEZA FIRME PREMA PARTNERU -- izvedena iz iste knjige, bez ijedne mutabilne
' kolone, preko DOPRINOSA po dogadjaju (AMB-INV-09).
'
' Prosta razlika dve sume NIJE dovoljna: storno ULAZA upisuje kontra-stav koji
' anulira fizicko stanje, a razlika dve sume ostavila bi obavezu da visi.
'
' Vrsta ORIGINALA se cita iz reda na koji StornoOd pokazuje, a ne iz samog
' kontra-stava. Razlika je bitna: inace bi 10d mogao da ugasi obavezu upisujuci
' storno sa drugom vrstom, i nijedna provera to ne bi videla. StornoOd koji ne
' pokazuje nigde je kvar, ne nula.
Public Function AmbObavezaPartneru(ByVal tip As String, ByVal id As String, _
                                   ByVal tipAmb As String) As Double
    Const SRC As String = "modAmbalaza.AmbObavezaPartneru"

    Dim p As String
    p = modAmbalazaUgovor.AmbNalogProblem(tip, id)
    If Len(p) > 0 Then Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, p

    If Len(Trim$(tipAmb)) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                  "Obaveza se vodi PO TIPU ambalaze -- tip je obavezan."
    End If

    ' OBAVEZA POSTOJI SAMO PREMA NALOGU KOJI MOZE DA PRIMI TUDJU AMBALAZU.
    '
    ' Bez ove kapije funkcija vraca broj i za stanicu: njeni VRACANJE_TUDJE redovi
    ' daju -N, pa bi "firma duguje stanici -70" izgledalo kao podatak. Klasa se
    ' CITA iz iste matrice koja definise pokrice (ULAZ_TUDJE: GRANICA -> PARTNER)
    ' -- dug nastaje tacno tim dogadjajem, pa mu je klasa ista po konstrukciji.
    Dim klasaDuga As String
    klasaDuga = modAmbalazaUgovor.AmbPokriceKlasa()
    If Len(klasaDuga) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_KVAR, SRC, "Klasa duga nije definisana u matrici."
    End If

    If Not modAmbalazaUgovor.AmbNalogUKlasi(klasaDuga, tip) Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                  "Obaveza se ne vodi prema nalogu " & Trim$(tip) & ": dug nastaje " & _
                  "ulazom tudje ambalaze, a on ide samo na " & klasaDuga & "."
    End If

    RequireKnjigaSchema SRC

    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then Exit Function

    Dim cID As Long, cOdTip As Long, cOdID As Long, cNaTip As Long, cNaID As Long
    Dim cVK As Long, cTipA As Long, cKol As Long, cSt As Long
    Dim kol As Object, vrste As Object
    Set kol = KnjigaZaCitanje(data, SRC, vrste)

    cID = kol(COL_AMB_ID)
    cOdTip = kol(COL_AMB_OD_TIP)
    cOdID = kol(COL_AMB_OD_ID)
    cNaTip = kol(COL_AMB_NA_TIP)
    cNaID = kol(COL_AMB_NA_ID)
    cVK = kol(COL_AMB_VRSTA_KRETANJA)
    cTipA = kol(COL_AMB_TIP)
    cKol = kol(COL_AMB_KOLICINA)
    cSt = kol(COL_AMB_STORNO_OD)

    ' Mapa AmbID -> vrsta dolazi iz KAPIJE INTEGRITETA, ne iz prolaza ovog citaoca:
    ' jedinstvenost identiteta je svojstvo KNJIGE, pa mora vaziti i za citaoca salda
    ' (review #400, drugi krug).
    Dim i As Long
    Dim obaveza As Double, stornoOd As String, vrstaOrig As String
    For i = 1 To UBound(data, 1)
        If RedDoticeKnjigu(data, i, kol) Then
            If StrComp(AmbText(data(i, cTipA)), Trim$(tipAmb), vbTextCompare) = 0 Then
                If IstiNalog(AmbText(data(i, cNaTip)), AmbText(data(i, cNaID)), tip, id) Or _
                   IstiNalog(AmbText(data(i, cOdTip)), AmbText(data(i, cOdID)), tip, id) Then

                    stornoOd = AmbText(data(i, cSt))
                    vrstaOrig = ""
                    If Len(stornoOd) > 0 Then
                        If Not vrste.Exists(stornoOd) Then
                            Err.Raise AMB_ERR_KNJIGA_KVAR, SRC, _
                                      "StornoOd '" & stornoOd & "' ne pokazuje na red knjige."
                        End If
                        vrstaOrig = CStr(vrste(stornoOd))
                    End If

                    obaveza = obaveza + modAmbalazaUgovor.AmbDoprinosObavezi( _
                                            AmbText(data(i, cVK)), _
                                            CDbl(data(i, cKol)), _
                                            vrstaOrig)
                End If
            End If
        End If
    Next i

    AmbObavezaPartneru = obaveza
End Function

' DEFICIT: koliko bi IZVORNOM nalogu falilo da se prenos upise bez pokrica.
'
' GRANICA (SpoljniSvet) nema fizicko stanje -- ona je izvor i ponor opticaja
' (6.3) -- pa prenos IZ nje nikad nije u manjku. Za svaki realan nalog deficit je
' obican racun, i javan je zato sto UI sme da ga PRIKAZE pre upisa.
Public Function AmbDeficitZaPrenos(ByVal odTip As String, ByVal odID As String, _
                                   ByVal tipAmb As String, _
                                   ByVal kolicina As Double) As Double
    If Not modAmbalazaUgovor.AmbNalogUKlasi(AMB_KLASA_REALAN, odTip) Then Exit Function

    Dim saldo As Double
    saldo = AmbSaldoNaloga(odTip, odID, tipAmb)
    If kolicina > saldo Then AmbDeficitZaPrenos = kolicina - saldo
End Function

' ============================================================
' AMBALAZNI DOKUMENT -- zaglavlje (AMB-10-DOK)
' ============================================================

Public Function NoviAmbDokID() As String
    Const SRC As String = "modAmbalaza.NoviAmbDokID"

    NoviAmbDokID = NewEntityID("ADK-")

    If Len(Trim$(NoviAmbDokID)) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_KVAR, SRC, "NewEntityID nije vratio AmbDokID."
    End If
End Function

' Zaglavlje dokumenta koji nosi dogadjaje bez sopstvenog poslovnog dokumenta
' (revers, nabavka, otpis). Vraca AmbDokID -- on ide u tblAmbalaza.DokumentID,
' uz DokumentTIP = DOK_TIP_AMBALAZA_DOKUMENT.
'
' Red se gradi PO IMENU KOLONE (SetRowValueByColumn), ne golim Array(...):
' AppendRow pise poziciono, pa bi kolona ubacena u sredinu tiho poslala sve iza
' sebe u pogresna polja.
' Zaglavlje trazi SVOJU tabelu u snapshotu, ne knjigu: pozivalac koji kreira
' dokument I redove mora da snapshotuje OBE, inace rollback vraca pola
' dokumenta -- zaglavlje bez redova ili redove bez zaglavlja.
Public Function UpisiAmbDokument(ByVal tx As clsTransaction, _
                                 ByVal vrsta As String, ByVal broj As String, _
                                 ByVal datum As Date, _
                                 ByVal brojOwnerTip As String, _
                                 ByVal brojOwnerID As String, _
                                 Optional ByVal napomena As String = "") As String
    Const SRC As String = "modAmbalaza.UpisiAmbDokument"

    RequireAmbTxVlasnistvo tx, TBL_AMBALAZA_DOKUMENT, SRC
    RequireAmbDokSchema SRC
    modAmbalazaUgovor.RequireAmbDok vrsta, broj, datum, brojOwnerTip, brojOwnerID, SRC

    ' ZAUZETOST BROJA JE KAPIJA PISCA, NE UI-ja. RequireAmbDok sudi OBLIK
    ' (vrsta, neprazan broj, datum, klasa vlasnika); bez ovoga su dva poziva sa
    ' istim rucno prosledjenim brojem davala dva AmbDokID-a i jedan poslovni broj
    ' u istom nizu (review 03.10.2026, P2 #2). Opseg je kanonski:
    ' (BrojOwnerTip, BrojOwnerID, dan) -- isti koji generator koristi.
    Dim zauzeo As String
    zauzeo = modBrojevi.AmbDokBrojZauzet(brojOwnerTip, brojOwnerID, datum, broj)
    If Len(zauzeo) > 0 Then
        Err.Raise AMB_ERR_IDENTITET, SRC, _
                  "Broj '" & Trim$(broj) & "' je u nizu " & Trim$(brojOwnerTip) & _
                  " '" & Trim$(brojOwnerID) & "' tog dana vec zauzet (dokument " & _
                  zauzeo & "). Storno ne oslobadja broj."
    End If

    Dim vrstaK As String, ownerK As String
    vrstaK = modAmbalazaUgovor.AmbDokVrstaKanon(vrsta)
    ownerK = modAmbalazaUgovor.AmbNalogTipKanon(brojOwnerTip)
    If Len(vrstaK) = 0 Or Len(ownerK) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_KVAR, SRC, _
                  "Kanonski zapis nije nadjen posle prosle kapije -- kvar ugovora."
    End If

    Dim colCount As Long
    colCount = TabelaBrojKolona(TBL_AMBALAZA_DOKUMENT)
    If colCount <= 0 Then
        Err.Raise AMB_ERR_KNJIGA_KVAR, SRC, "Ne mogu da odredim broj kolona za tblAmbalazaDokument."
    End If

    Dim novID As String
    novID = NoviAmbDokID()

    Dim rowData() As Variant
    ReDim rowData(0 To colCount - 1)

    SetRowValueByColumn rowData, TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, novID, SRC
    SetRowValueByColumn rowData, TBL_AMBALAZA_DOKUMENT, COL_AMBD_VRSTA, vrstaK, SRC
    SetRowValueByColumn rowData, TBL_AMBALAZA_DOKUMENT, COL_AMBD_BROJ, Trim$(broj), SRC
    SetRowValueByColumn rowData, TBL_AMBALAZA_DOKUMENT, COL_AMBD_DATUM, datum, SRC
    SetRowValueByColumn rowData, TBL_AMBALAZA_DOKUMENT, COL_AMBD_BROJ_OWNER_TIP, ownerK, SRC
    SetRowValueByColumn rowData, TBL_AMBALAZA_DOKUMENT, COL_AMBD_BROJ_OWNER_ID, Trim$(brojOwnerID), SRC
    SetRowValueByColumn rowData, TBL_AMBALAZA_DOKUMENT, COL_AMBD_NAPOMENA, Trim$(napomena), SRC
    SetRowValueByColumn rowData, TBL_AMBALAZA_DOKUMENT, COL_STORNIRANO, "", SRC

    Dim rowIdx As Long
    rowIdx = AppendRow(TBL_AMBALAZA_DOKUMENT, rowData)
    If rowIdx <= 0 Then
        Err.Raise AMB_ERR_KNJIGA_KVAR, SRC, "AppendRow nije uspeo za tblAmbalazaDokument."
    End If

    ' Dokument koji smo upravo napravili pripada OVOJ transakciji, pa ga ona i
    ' vezuje: sledeci PrenesiAmbalazu nad njim ne trazi nista od pozivaoca.
    tx.BindSourceDocument DOK_TIP_AMBALAZA_DOKUMENT, novID

    UpisiAmbDokument = novID
End Function

' Zaglavlje u celini: vrsta + vlasnik numerickog niza.
'
' Postoji jer AMB-10-ODL-10 trazi da se BrojOwner uporedi sa stranom kretanja,
' a to se ne moze iz vrste same. AmbDokVrstaZaID ostaje javan (testovi ga
' koriste) i poziva ovo, da provera storniranog i postojanja stoji na JEDNOM
' mestu.
' Vlasnik broja ROBNOG dokumenta, iz zatvorene mape (AMB-10-ODL-22).
'
' Prazno za tip koji ga ne objavljuje -- i to je fail-closed, jer obrnuta
' kapija ODL-10 bez vlasnika broja odbija povrat od kupca.
'
' Cita se IZ TABELE dokumenta, ne iz argumenta pozivaoca: tabela je cinjenica,
' a argument bi bio tvrdnja pisca o sebi (isti razlog kao u modStornoZurnal).
Public Sub AmbRobniZaglavlje(ByVal dokTip As String, ByVal dokID As String, _
                             ByRef outOwnerTip As String, _
                             ByRef outOwnerID As String)
    Const SRC As String = "modAmbalaza.AmbRobniZaglavlje"

    outOwnerTip = ""
    outOwnerID = ""
    If Len(Trim$(dokID)) = 0 Then Exit Sub

    Dim mapa As Variant, i As Long, red As Variant
    mapa = modAmbalazaUgovor.AmbRobniVlasniciBroja()
    For i = LBound(mapa) To UBound(mapa)
        red = mapa(i)
        If StrComp(Trim$(dokTip), CStr(red(0)), vbTextCompare) = 0 Then
            outOwnerTip = CStr(red(3))
            outOwnerID = Trim$(NzToText(LookupValue(CStr(red(1)), CStr(red(2)), _
                                                    dokID, CStr(red(4)))))
            If Len(outOwnerID) = 0 Then
                Err.Raise AMB_ERR_IDENTITET, SRC, _
                          "Dokument " & Trim$(dokTip) & " " & Trim$(dokID) & _
                          " ne nosi vlasnika svog broja (" & CStr(red(4)) & _
                          "). AMB-10-ODL-22 trazi da ga robni dokument objavi."
            End If
            Exit Sub
        End If
    Next i
End Sub

Public Sub AmbDokZaglavlje(ByVal ambDokID As String, _
                           ByRef vrsta As String, _
                           ByRef brojOwnerTip As String, _
                           ByRef brojOwnerID As String)
    Const SRC As String = "modAmbalaza.AmbDokZaglavlje"

    RequireAmbDokSchema SRC
    RequireTacnoJedan TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, Trim$(ambDokID), _
                      "Ambalazni dokument", SRC

    Dim st As Variant
    st = LookupValue(TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, Trim$(ambDokID), COL_STORNIRANO)
    If Len(AmbText(st)) > 0 Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                  "Ambalazni dokument '" & Trim$(ambDokID) & "' je storniran."
    End If

    vrsta = AmbText(LookupValue(TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, _
                                Trim$(ambDokID), COL_AMBD_VRSTA))
    brojOwnerTip = AmbText(LookupValue(TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, _
                                       Trim$(ambDokID), COL_AMBD_BROJ_OWNER_TIP))
    brojOwnerID = AmbText(LookupValue(TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, _
                                      Trim$(ambDokID), COL_AMBD_BROJ_OWNER_ID))

    If Len(vrsta) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_KVAR, SRC, _
                  "Ambalazni dokument '" & Trim$(ambDokID) & "' nema vrstu."
    End If
End Sub

' Vrsta posla sa zaglavlja, za kapiju AmbDokDozvoljavaKretanje.
'
' STORNIRAN DOKUMENT NE PRIMA NOVA KRETANJA: zaglavlje sme da nosi Stornirano
' (ono nije knjiga), a dopisivanje na ponisten dokument bilo bi kretanje bez
' ziveg povoda.
Public Function AmbDokVrstaZaID(ByVal ambDokID As String) As String
    Dim ot As String, oi As String
    AmbDokZaglavlje ambDokID, AmbDokVrstaZaID, ot, oi
End Function

' ============================================================
' UPIS JEDNOG REDA KNJIGE -- jedini put do tblAmbalaza u novom modelu
' ============================================================
'
' AMB-INV-04 se proverava OVDE, nad svakim upisanim redom, a ne samo u
' PrenesiAmbalazu: pokrice deficita i ostatak podele pisac generise SAM, pa bi ih
' provera na ulazu promasila.
'
' Kljuc je (DokumentTIP, DokumentID, VrstaKretanja, TipAmbalaze) i vazi za
' ORIGINALE. Kontra-stav (StornoOd <> "") ima isti kljuc kao original po
' konstrukciji, pa se na njega ne primenjuje -- njegovu jedinstvenost drzi
' AMB-INV-06, u 10d.
' AMB-INV-08 ZIVI OVDE, jer kroz ovu funkciju prolazi SVAKI red knjige -- i
' pokrice deficita i ostatak podele, ne samo trazeni prenos. Kopija kapije na
' ulazu u PrenesiAmbalazu bila bi placebo: jezgro bi odbilo isti upis i bez nje,
' pa je nijedna sabotaza ne bi mogla oboriti. Isti razlog je tamo vec zapisan za
' identitet dokumenta.
Private Function UpisiRedKnjige(ByVal tx As clsTransaction, _
                                ByVal datum As Date, ByVal tipAmb As String, _
                                ByVal kolicina As Double, _
                                ByVal odTip As String, ByVal odID As String, _
                                ByVal naTip As String, ByVal naID As String, _
                                ByVal dokTip As String, ByVal dokID As String, _
                                ByVal vrsta As String, ByVal stornoOd As String, _
                                ByVal sourceName As String) As String
    RequireAmbTxVlasnistvo tx, TBL_AMBALAZA, sourceName
    RequireAmbTxIzvorniDokument tx, dokTip, dokID, sourceName
    RequireKnjigaSchema sourceName

    ' KONTRA-STAV SE PROVERAVA U OBRNUTOM SMERU -- isto pravilo koje citalac
    ' vec ima (KnjigaRedProblem). Kontra-stav nosi zamenjene Od i Na, pa bi ga
    ' prava provera odbila na matrici klasa: IZDATA_PRAZNA trazi SOPSTVENI kao
    ' izvor, a kontra-stav tu ima partnera. Pravilo stoji na JEDNOM mestu u
    ' pisacu, da ne moze da se razidje sa citaocem.
    If Len(Trim$(stornoOd)) = 0 Then
        modAmbalazaUgovor.RequireAmbPrenos odTip, odID, naTip, naID, kolicina, tipAmb, vrsta, sourceName
    Else
        modAmbalazaUgovor.RequireAmbPrenos naTip, naID, odTip, odID, kolicina, tipAmb, vrsta, sourceName
    End If

    Dim vrstaK As String, odTipK As String, naTipK As String
    vrstaK = modAmbalazaUgovor.AmbVrstaKanon(vrsta)
    odTipK = modAmbalazaUgovor.AmbNalogTipKanon(odTip)
    naTipK = modAmbalazaUgovor.AmbNalogTipKanon(naTip)
    If Len(vrstaK) = 0 Or Len(odTipK) = 0 Or Len(naTipK) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_KVAR, sourceName, _
                  "Kanonski zapis nije nadjen posle prosle kapije -- kvar ugovora."
    End If

    If Len(Trim$(dokTip)) = 0 Or Len(Trim$(dokID)) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, sourceName, _
                  "Red knjige nema identitet dokumenta (AMB-INV-04 i -08)."
    End If

    If Len(Trim$(stornoOd)) = 0 Then
        ' AMB-INV-10 (sprovodjenje AMB-10-ODL-3): dokument ZAKLJUCAVA svoj poslovni
        ' par naloga. AMB-INV-04 to ne pokriva -- njegov kljuc nosi vrstu i tip, pa
        ' dva protivpartnera prolaze cim se razlikuje bilo koje od toga dvoga:
        '
        '   ADK-1  Stanica -> K1  IZDATA_PRAZNA  GAJBA_A
        '   ADK-1  K2 -> Stanica  POVRAT_PRAZNE  GAJBA_A    <- druga vrsta
        '   ADK-1  Stanica -> K2  IZDATA_PRAZNA  GAJBA_B    <- drugi tip
        Dim par As String
        par = DokumentParProblem(dokTip, dokID, odTipK, odID, naTipK, naID, vrstaK, sourceName)
        If Len(par) > 0 Then
            Err.Raise AMB_ERR_JEDAN_PARTNER, sourceName, par
        End If

        Dim sudar As String
        sudar = IdentitetZauzeo(dokTip, dokID, vrstaK, tipAmb, sourceName)
        If Len(sudar) > 0 Then
            Err.Raise AMB_ERR_IDENTITET, sourceName, _
                      "AMB-INV-04: " & Trim$(dokTip) & " '" & Trim$(dokID) & "' je vec knjizio " & _
                      vrstaK & " za '" & Trim$(tipAmb) & "' (red " & sudar & ")."
        End If
    End If

    Dim colCount As Long
    colCount = TabelaBrojKolona(TBL_AMBALAZA)
    If colCount <= 0 Then
        Err.Raise AMB_ERR_KNJIGA_KVAR, sourceName, "Ne mogu da odredim broj kolona za tblAmbalaza."
    End If

    ' Transakcioni identitet, ne max+1: StornoOd pokazuje na AmbID, pa on mora biti
    ' stabilan i neprotumaciv (ista odluka kao ReversID, DOCUMENT_HEADER_LINES par. 2).
    Dim novID As String
    novID = NewEntityID("AMB-")
    If Len(Trim$(novID)) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_KVAR, sourceName, "NewEntityID nije vratio AmbID."
    End If

    Dim rowData() As Variant
    ReDim rowData(0 To colCount - 1)

    SetRowValueByColumn rowData, TBL_AMBALAZA, COL_AMB_ID, novID, sourceName
    SetRowValueByColumn rowData, TBL_AMBALAZA, COL_AMB_DATUM, datum, sourceName
    SetRowValueByColumn rowData, TBL_AMBALAZA, COL_AMB_TIP, Trim$(tipAmb), sourceName
    SetRowValueByColumn rowData, TBL_AMBALAZA, COL_AMB_KOLICINA, kolicina, sourceName
    SetRowValueByColumn rowData, TBL_AMBALAZA, COL_AMB_OD_TIP, odTipK, sourceName
    SetRowValueByColumn rowData, TBL_AMBALAZA, COL_AMB_OD_ID, Trim$(odID), sourceName
    SetRowValueByColumn rowData, TBL_AMBALAZA, COL_AMB_NA_TIP, naTipK, sourceName
    SetRowValueByColumn rowData, TBL_AMBALAZA, COL_AMB_NA_ID, Trim$(naID), sourceName
    SetRowValueByColumn rowData, TBL_AMBALAZA, COL_AMB_DOK_TIP, Trim$(dokTip), sourceName
    SetRowValueByColumn rowData, TBL_AMBALAZA, COL_AMB_DOK_ID, Trim$(dokID), sourceName
    SetRowValueByColumn rowData, TBL_AMBALAZA, COL_AMB_VRSTA_KRETANJA, vrstaK, sourceName
    SetRowValueByColumn rowData, TBL_AMBALAZA, COL_AMB_STORNO_OD, Trim$(stornoOd), sourceName

    Dim rowIdx As Long
    rowIdx = AppendRow(TBL_AMBALAZA, rowData)
    If rowIdx <= 0 Then
        Err.Raise AMB_ERR_KNJIGA_KVAR, sourceName, "AppendRow nije uspeo za tblAmbalaza."
    End If

    UpisiRedKnjige = novID
End Function

' AMB-INV-10: dokument ZAKLJUCAVA svoj poslovni par naloga.
'
' Prva verzija je BROJALA naloge van granice i trazila "najvise dva". To pada tacno
' tamo gde je jedna strana granica (review #400, drugi krug):
'
'   ADK-NAB-1  SpoljniSvet -> Stanica1  NABAVKA  GAJBA_A
'   ADK-NAB-1  SpoljniSvet -> Stanica2  NABAVKA  GAJBA_B
'
' dva realna naloga, broj = 2, PROSLO BI -- a to je jedan nabavni dokument preko DVE
' stanice, tacno ono zbog cega AMB-10-ODL-3 postoji. Ogledalno vazi za OTPIS.
'
' Zato se ne broji nego POREDI: svi ORIGINALNI redovi dokumenta, izuzev generisanog
' pokrica, imaju JEDAN I ISTI neuredjen par {Od, Na}.
'
' POKRICE JE IZUZETO jer ga pisac generise sam (ULAZ_TUDJE nikad nije zahtev), pa
' njegov par ({SpoljniSvet, izvor}) nije poslovni par dokumenta. Ali ne sme da uvede
' trecu stranu, pa njegovo ODREDISTE mora biti clan zakljucanog para.
'
' Vazi za SVE dokumente, ne samo ambalazne: merenje svih devet mesta knjizenja daje
' po dokumentu tacno jedan par (otkup K1<->Stanica, otpremnica Stanica<->Vozac,
' prijemnica Kupac<->Vozac, revers dve strane). Ogranicenje je time na KNJIZI, ne na
' vrsti dokumenta -- isti razlog zbog kog AMB-INV-04 nosi DokumentTIP.
Private Function DokumentParProblem(ByVal dokTip As String, ByVal dokID As String, _
                                    ByVal odTipK As String, ByVal odID As String, _
                                    ByVal naTipK As String, ByVal naID As String, _
                                    ByVal vrstaK As String, _
                                    ByVal sourceName As String) As String
    Dim parovi As Object, clanovi As Object, pokrica As Object
    Set parovi = CreateObject("Scripting.Dictionary")
    Set clanovi = CreateObject("Scripting.Dictionary")
    Set pokrica = CreateObject("Scripting.Dictionary")

    DodajStranu parovi, clanovi, pokrica, odTipK, odID, naTipK, naID, vrstaK

    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA)

    If Not IsEmpty(data) Then
        Dim kol As Object, vrste As Object
        Set kol = KnjigaZaCitanje(data, sourceName, vrste)

        Dim i As Long
        For i = 1 To UBound(data, 1)
            If RedDoticeKnjigu(data, i, kol) Then
                If Len(AmbText(data(i, kol(COL_AMB_STORNO_OD)))) = 0 Then
                    If StrComp(AmbText(data(i, kol(COL_AMB_DOK_TIP))), Trim$(dokTip), vbTextCompare) = 0 And _
                       StrComp(AmbText(data(i, kol(COL_AMB_DOK_ID))), Trim$(dokID), vbTextCompare) = 0 Then
                        DodajStranu parovi, clanovi, pokrica, _
                                    AmbText(data(i, kol(COL_AMB_OD_TIP))), AmbText(data(i, kol(COL_AMB_OD_ID))), _
                                    AmbText(data(i, kol(COL_AMB_NA_TIP))), AmbText(data(i, kol(COL_AMB_NA_ID))), _
                                    AmbText(data(i, kol(COL_AMB_VRSTA_KRETANJA)))
                    End If
                End If
            End If
        Next i
    End If

    If parovi.count > 1 Then
        DokumentParProblem = "AMB-INV-10: " & Trim$(dokTip) & " '" & Trim$(dokID) & _
                             "' bi nosio " & CStr(parovi.count) & " razlicita para naloga (" & _
                             Join(parovi.Keys, " | ") & ") -- jedan dokument, jedan poslovni par."
        Exit Function
    End If

    ' Pokrice bez zakljucanog para nema sa cim da se poredi. Pisac ga ne proizvodi
    ' sam: trazeni prenos se upisuje PRE pokrica, pa par uvek postoji. Ostaje kao
    ' izgovorena granica, ne kao tiha.
    If parovi.count = 0 Then Exit Function

    Dim k As Variant
    For Each k In pokrica.Keys
        If Not clanovi.Exists(CStr(k)) Then
            DokumentParProblem = "AMB-INV-10: pokrice deficita ide na '" & CStr(k) & _
                                 "', a par dokumenta je " & Join(parovi.Keys, "") & "."
            Exit Function
        End If
    Next k
End Function

' Jedan red doprinosi ILI paru dokumenta ILI spisku pokrica -- nikad oboma.
Private Sub DodajStranu(ByRef parovi As Object, ByRef clanovi As Object, _
                        ByRef pokrica As Object, _
                        ByVal odTip As String, ByVal odID As String, _
                        ByVal naTip As String, ByVal naID As String, _
                        ByVal vrsta As String)
    If StrComp(Trim$(vrsta), AMB_VK_ULAZ_TUDJE, vbTextCompare) = 0 Then
        Dim p As String
        p = NalogKljuc(naTip, naID)
        If Not pokrica.Exists(p) Then pokrica.Add p, True
        Exit Sub
    End If

    Dim par As String
    par = ParKljuc(odTip, odID, naTip, naID)
    If Not parovi.Exists(par) Then parovi.Add par, True

    Dim a As String, b As String
    a = NalogKljuc(odTip, odID)
    b = NalogKljuc(naTip, naID)
    If Not clanovi.Exists(a) Then clanovi.Add a, True
    If Not clanovi.Exists(b) Then clanovi.Add b, True
End Sub

' Vraca AmbID reda koji je vec zauzeo isti identitet efekta, inace "".
Private Function IdentitetZauzeo(ByVal dokTip As String, ByVal dokID As String, _
                                 ByVal vrstaK As String, ByVal tipAmb As String, _
                                 ByVal sourceName As String) As String
    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then Exit Function

    Dim cID As Long, cOdTip As Long, cNaTip As Long, cVK As Long
    Dim cTipA As Long, cDokT As Long, cDokI As Long, cSt As Long
    Dim kol As Object, vrste As Object
    Set kol = KnjigaZaCitanje(data, sourceName, vrste)

    cID = kol(COL_AMB_ID)
    cOdTip = kol(COL_AMB_OD_TIP)
    cNaTip = kol(COL_AMB_NA_TIP)
    cVK = kol(COL_AMB_VRSTA_KRETANJA)
    cTipA = kol(COL_AMB_TIP)
    cDokT = kol(COL_AMB_DOK_TIP)
    cDokI = kol(COL_AMB_DOK_ID)
    cSt = kol(COL_AMB_STORNO_OD)

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If RedDoticeKnjigu(data, i, kol) Then
            If Len(AmbText(data(i, cSt))) = 0 Then
                If StrComp(AmbText(data(i, cDokT)), Trim$(dokTip), vbTextCompare) = 0 And _
                   StrComp(AmbText(data(i, cDokI)), Trim$(dokID), vbTextCompare) = 0 And _
                   StrComp(AmbText(data(i, cVK)), Trim$(vrstaK), vbTextCompare) = 0 And _
                   StrComp(AmbText(data(i, cTipA)), Trim$(tipAmb), vbTextCompare) = 0 Then
                    IdentitetZauzeo = AmbText(data(i, cID))
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

' ZBIR VEC UPISANOG ZA ISTI ZAHTEV -- osnova idempotencije.
'
' Meri se po PARU NALOGA i po SKUPU vrsta koje taj zahtev sme da proizvede
' (AmbVrsteZahteva), a ne po trazenoj vrsti i kolicini jednog reda: podela
' obaveze razbija zahtev od 20 na 12 + 8, pa bi poredjenje sa jednim redom
' prijavilo sudar nad ispravnim ponavljanjem.
Private Function ZbirZahteva(ByVal dokTip As String, ByVal dokID As String, _
                             ByVal datum As Date, ByVal tipAmb As String, _
                             ByVal odTip As String, ByVal odID As String, _
                             ByVal naTip As String, ByVal naID As String, _
                             ByVal vrstaK As String, _
                             ByRef outAmbID As String, _
                             ByVal sourceName As String) As Double
    outAmbID = ""

    Dim dozvoljene As Variant
    dozvoljene = modAmbalazaUgovor.AmbVrsteZahteva(vrstaK)
    If UBound(dozvoljene) < LBound(dozvoljene) Then Exit Function

    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then Exit Function

    Dim cID As Long, cOdTip As Long, cOdID As Long, cNaTip As Long, cNaID As Long
    Dim cVK As Long, cTipA As Long, cKol As Long, cDokT As Long, cDokI As Long
    Dim cSt As Long, cDat As Long
    Dim kol As Object, vrste As Object
    Set kol = KnjigaZaCitanje(data, sourceName, vrste)

    cID = kol(COL_AMB_ID)
    cOdTip = kol(COL_AMB_OD_TIP)
    cOdID = kol(COL_AMB_OD_ID)
    cNaTip = kol(COL_AMB_NA_TIP)
    cNaID = kol(COL_AMB_NA_ID)
    cVK = kol(COL_AMB_VRSTA_KRETANJA)
    cTipA = kol(COL_AMB_TIP)
    cKol = kol(COL_AMB_KOLICINA)
    cDokT = kol(COL_AMB_DOK_TIP)
    cDokI = kol(COL_AMB_DOK_ID)
    cSt = kol(COL_AMB_STORNO_OD)
    cDat = kol(COL_AMB_DATUM)

    Dim i As Long, j As Long, vr As String, zbir As Double
    For i = 1 To UBound(data, 1)
        If RedDoticeKnjigu(data, i, kol) Then
            If Len(AmbText(data(i, cSt))) = 0 Then
                If StrComp(AmbText(data(i, cDokT)), Trim$(dokTip), vbTextCompare) = 0 And _
                   StrComp(AmbText(data(i, cDokI)), Trim$(dokID), vbTextCompare) = 0 And _
                   StrComp(AmbText(data(i, cTipA)), Trim$(tipAmb), vbTextCompare) = 0 And _
                   IstiDan(data(i, cDat), datum) Then
                    If IstiNalog(AmbText(data(i, cOdTip)), AmbText(data(i, cOdID)), odTip, odID) And _
                       IstiNalog(AmbText(data(i, cNaTip)), AmbText(data(i, cNaID)), naTip, naID) Then

                        vr = AmbText(data(i, cVK))
                        For j = LBound(dozvoljene) To UBound(dozvoljene)
                            If StrComp(vr, CStr(dozvoljene(j)), vbTextCompare) = 0 Then
                                zbir = zbir + CDbl(data(i, cKol))
                                If Len(outAmbID) = 0 Or _
                                   StrComp(vr, Trim$(vrstaK), vbTextCompare) = 0 Then
                                    outAmbID = AmbText(data(i, cID))
                                End If
                                Exit For
                            End If
                        Next j
                    End If
                End If
            End If
        End If
    Next i

    ZbirZahteva = zbir
End Function

' ============================================================
' STORNO U KNJIZI -- KONTRA-STAV, NE ZASTAVICA (AMB-10-ODL-16)
' ============================================================
'
' Zastavica Stornirano je mehanizam STAROG modela i nov citalac je ne gleda:
' RedDoticeKnjigu trazi Od_*/Na_*/VrstaKretanja/StornoOd, a AmbSaldoNaloga
' sumira po tome. Dokument presecen na nov model a storniran zastavicom
' ostavio bi gajbe na saldu TIHO, pa ovaj ulaz ide PRED cutover mesta
' knjizenja, ne posle njega.
'
' DATUM KONTRA-STAVA JE DATUM ORIGINALA, ne danasnji. Stara zastavica je red
' uklanjala iz SVIH perioda; isti datum je jedini oblik koji ne menja nijedan
' periodski saldo. Danasnji datum bi ostavio fantom u starom periodu i visak
' u novom.
'
' IDEMPOTENTNO: original koji VEC ima kontra-stav se preskace, pa drugi poziv
' vraca 0 i ne duplira. Vraca broj upisanih kontra-stavova.
Public Function StornirajAmbalazuDokumenta(ByVal tx As clsTransaction, _
                                           ByVal dokTip As String, _
                                           ByVal dokID As String) As Long
    Const SRC As String = "modAmbalaza.StornirajAmbalazuDokumenta"

    If tx Is Nothing Then
        Err.Raise AMB_ERR_STORNO, SRC, _
                  "Storno knjige trazi aktivnu transakciju (AMB-INV-08)."
    End If
    If Len(Trim$(dokTip)) = 0 Or Len(Trim$(dokID)) = 0 Then
        Err.Raise AMB_ERR_STORNO, SRC, "Storno knjige trazi identitet dokumenta."
    End If

    RequireKnjigaSchema SRC

    ' OVAJ PRIMITIV NE VEZUJE DOKUMENT, I TO JE SUSTINA KAPIJE.
    '
    ' Ranije je zvao BindSourceDocument sam, uz obrazlozenje da to nije
    ' samopotvrda jer ista kapija trazi i izvornu tabelu u snapshotu. Ta odbrana
    ' je falsifikovana (review 03.10.2026, P1 #1): snapshot je jeftin i ne
    ' dokazuje da je dokument PROMENJEN. Pozivalac je mogao da snapshotuje
    ' tblOtkup, anulira efekat AKTIVNOG otkupa i prodje sve kapije.
    '
    ' Ledger-storno NIJE pisac izvornog dokumenta. Vezivanje zato pripada
    ' kanonskom piscu koji je zaglavlje stvarno promenio (modStorno.StornoOtkup:
    ' MarkRowStornirano, pa bind, pa ovaj poziv). Ako nije vezao, ovde pada
    ' fail-closed kroz RequireAmbTxIzvorniDokument -- sto je i jedini nacin da
    ' AMB-INV-08 znaci "izvorni dokument je promenjen u istoj TX", a ne
    ' "transakcija je rekla da ga poseduje".
    '
    ' Posledica: tblAmbalazaDokument (nabavka, revers) jos NEMA kanonskog
    ' storno pisca, pa njegov ledger storno ovde pada -- namerno, dok taj pisac
    ' ne nastane (10d).

    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then Exit Function

    Dim kol As Object, vrste As Object
    Set kol = KnjigaZaCitanje(data, SRC, vrste)

    Dim cID As Long, cDat As Long, cTipA As Long, cKol As Long
    Dim cOdTip As Long, cOdID As Long, cNaTip As Long, cNaID As Long
    Dim cVK As Long, cDokT As Long, cDokI As Long, cSt As Long
    cID = kol(COL_AMB_ID)
    cDat = kol(COL_AMB_DATUM)
    cTipA = kol(COL_AMB_TIP)
    cKol = kol(COL_AMB_KOLICINA)
    cOdTip = kol(COL_AMB_OD_TIP)
    cOdID = kol(COL_AMB_OD_ID)
    cNaTip = kol(COL_AMB_NA_TIP)
    cNaID = kol(COL_AMB_NA_ID)
    cVK = kol(COL_AMB_VRSTA_KRETANJA)
    cDokT = kol(COL_AMB_DOK_TIP)
    cDokI = kol(COL_AMB_DOK_ID)
    cSt = kol(COL_AMB_STORNO_OD)

    ' PRVI PROLAZ: na sta kontra-stavovi vec pokazuju. Bez ovoga drugi poziv
    ' udvaja storno, a saldo prelazi na drugu stranu umesto da stane na nuli.
    '
    ' NEDOSTIZNA ODBRANA -- namerno, i tako imenovana. Primitiv ne vezuje dokument
    ' (P1 #1), pa ga test ne moze pozvati direktno; jedini pozivalac je
    ' modStorno.StornoOtkup, a RequireStornoAllowed odbija DRUGI storno pre njega.
    ' Grana se zato NE MOZE dosegnuti i njena sabotaza je obrisana iz kataloga:
    ' sabotaza koja ne moze da obori nijednu tvrdnju je placebo. Ostaje jer bi
    ' drugi pozivalac (otpremnica, prijemnica) bez nje udvajao kontra-stav.
    Dim vecStornirani As Object
    Set vecStornirani = CreateObject("Scripting.Dictionary")
    Dim i As Long, st As String
    For i = 1 To UBound(data, 1)
        If RedDoticeKnjigu(data, i, kol) Then
            st = AmbText(data(i, cSt))
            If Len(st) > 0 Then
                If Not vecStornirani.Exists(st) Then vecStornirani.Add st, True
            End If
        End If
    Next i

    ' DRUGI PROLAZ: aktivni originali ovog dokumenta.
    Dim originali As Collection
    Set originali = New Collection
    ' SVI nalozi koje kontra-stavovi dotice -- iz njih se posle mere DVE
    ' invarijante: AMB-INV-07 nad realnim, AMB-INV-09 nad partnerskim.
    Dim nalozi As Object
    Set nalozi = CreateObject("Scripting.Dictionary")

    For i = 1 To UBound(data, 1)
        If RedDoticeKnjigu(data, i, kol) Then
            If Len(AmbText(data(i, cSt))) = 0 Then
                If StrComp(AmbText(data(i, cDokT)), Trim$(dokTip), vbTextCompare) = 0 And _
                   StrComp(AmbText(data(i, cDokI)), Trim$(dokID), vbTextCompare) = 0 Then
                    If Not vecStornirani.Exists(AmbText(data(i, cID))) Then
                        If Not IsDate(data(i, cDat)) Then
                            Err.Raise AMB_ERR_STORNO, SRC, _
                                      "Red knjige " & AmbText(data(i, cID)) & _
                                      " nema datum -- kontra-stav ga nasledjuje, pa bez " & _
                                      "njega storno ne moze da nastane."
                        End If
                        originali.Add Array(AmbText(data(i, cID)), _
                                            CDate(data(i, cDat)), _
                                            AmbText(data(i, cTipA)), _
                                            CDbl(data(i, cKol)), _
                                            AmbText(data(i, cOdTip)), AmbText(data(i, cOdID)), _
                                            AmbText(data(i, cNaTip)), AmbText(data(i, cNaID)), _
                                            AmbText(data(i, cVK)))
                        ZabeleziNalog nalozi, _
                                      AmbText(data(i, cOdTip)), AmbText(data(i, cOdID)), _
                                      AmbText(data(i, cTipA))
                        ZabeleziNalog nalozi, _
                                      AmbText(data(i, cNaTip)), AmbText(data(i, cNaID)), _
                                      AmbText(data(i, cTipA))
                    End If
                End If
            End If
        End If
    Next i

    If originali.count = 0 Then Exit Function

    Dim r As Variant
    For Each r In originali
        ' Od i Na ZAMENJENI, StornoOd pokazuje na original.
        UpisiRedKnjige tx, CDate(r(1)), CStr(r(2)), CDbl(r(3)), _
                       CStr(r(6)), CStr(r(7)), CStr(r(4)), CStr(r(5)), _
                       Trim$(dokTip), Trim$(dokID), CStr(r(8)), CStr(r(0)), SRC
    Next r

    ' DVE INVARIJANTE NAD POSLE-STANJEM.
    '
    ' Kontra-stav ide direktno kroz UpisiRedKnjige, pa ZAOBILAZI sve sto stoji u
    ' PrenesiAmbalazu -- a tamo zivi i AMB-INV-07. Bez ovih provera storno pise
    ' stanje koje normalan pisac eksplicitno zabranjuje (review 03.10.2026, P1 #2):
    '
    '   otkup donese 20 na stanicu -> ta 20 odu dalje (saldo 0) -> storno otkupa
    '   upise kontra-stav -20  =>  saldo -20, a INV-09 je uredan
    '
    ' Mere se POSTOJECIM citaocima nad upisanim stanjem, ne drugom kopijom pravila
    ' o znaku -- druga kopija bi se razisla sa prvom.
    Dim k As Variant, nl As Variant
    Dim saldo As Double, ob As Double

    ' AMB-INV-07: realan nalog ne sme da ostane u minusu. GRANICA (SpoljniSvet)
    ' je izuzeta po konstrukciji -- ona nema fizicko stanje (6.3), pa nije u
    ' klasi REALAN i petlja je i ne vidi.
    For Each k In nalozi.Keys
        nl = nalozi(k)
        If modAmbalazaUgovor.AmbNalogUKlasi(AMB_KLASA_REALAN, CStr(nl(0))) Then
            saldo = AmbSaldoNaloga(CStr(nl(0)), CStr(nl(1)), CStr(nl(2)))
            If saldo < 0 Then
                Err.Raise AMB_ERR_STORNO, SRC, _
                          "AMB-INV-07: storno bi ostavio saldo " & CStr(saldo) & " na " & _
                          CStr(nl(0)) & " '" & CStr(nl(1)) & "' za '" & CStr(nl(2)) & _
                          "' -- ta ambalaza je posle ovog dokumenta otisla dalje, pa " & _
                          "se dokument ne moze stornirati bez njenog vracanja."
            End If
        End If
    Next k

    ' AMB-INV-09: zahtev koji je ovaj ulaz NASLEDIO od pisca (PrenesiAmbalazu,
    ' VRACANJE_TUDJE): storno ULAZA tudje ambalaze cija je obaveza vec zatvorena
    ' vracanjem daje NEGATIVNU obavezu. Klasa duga se CITA iz matrice pokrica, jer
    ' dug nastaje tacno tim dogadjajem -- AmbObavezaPartneru dize gresku za nalog
    ' van te klase, pa se van nje i ne sme pitati.
    Dim klasaDuga As String
    klasaDuga = modAmbalazaUgovor.AmbPokriceKlasa()
    If Len(klasaDuga) > 0 Then
        For Each k In nalozi.Keys
            nl = nalozi(k)
            If modAmbalazaUgovor.AmbNalogUKlasi(klasaDuga, CStr(nl(0))) Then
                ob = AmbObavezaPartneru(CStr(nl(0)), CStr(nl(1)), CStr(nl(2)))
                If ob < 0 Then
                    Err.Raise AMB_ERR_STORNO, SRC, _
                              "Storno bi ostavio obavezu " & CStr(ob) & " prema " & _
                              CStr(nl(0)) & " '" & CStr(nl(1)) & "' za '" & CStr(nl(2)) & _
                              "' -- tudja ambalaza je vec vracena, pa se njen ulaz " & _
                              "ne moze stornirati (AMB-INV-09)."
                End If
            End If
        Next k
    End If

    StornirajAmbalazuDokumenta = originali.count
End Function

' Nalog koji je kontra-stav dotakao. BEZ klasne kapije: jedan popis, a klasu
' bira CITALAC -- INV-07 gleda REALAN, INV-09 klasu pokrica. Dva popisa bi se
' razisla, a prazan popis bi tiho ugasio onu proveru koja ga nema.
Private Sub ZabeleziNalog(ByRef nalozi As Object, _
                          ByVal tip As String, ByVal id As String, _
                          ByVal tipAmb As String)
    If Len(Trim$(tip)) = 0 Then Exit Sub

    Dim kljuc As String
    kljuc = NalogKljuc(tip, id) & "|" & UCase$(Trim$(tipAmb))
    If Not nalozi.Exists(kljuc) Then
        nalozi.Add kljuc, Array(Trim$(tip), Trim$(id), Trim$(tipAmb))
    End If
End Sub

' IMA LI KNJIGA KONTRA-STAV ZA OVAJ DOKUMENT?
'
' Javno jer ga undo garda (modStornoZurnal.UndoGuardReasonZaOp) mora pitati, a
' knjigu ne sme da cita sama -- oblik reda knjige nije njen posao.
'
' KLJUC JE KOMPOZITAN: (DokumentTIP, DokumentID). Prva verzija je trazila samo
' ID, uz obrazlozenje da je ID globalno jedinstven pa tip ne dodaje razlucivost
' -- a AMB-INV-04 nosi DokumentTIP tacno zato sto se jedan globalni namespace
' DokumentID-eva NE SME pretpostaviti (review 03.10.2026, P2 #1). Garda koja
' sluzi svim presecenim dokumentima ne sme da ima slabiji identitet od knjige.
Public Function AmbImaKontraStav(ByVal dokTip As String, _
                                 ByVal dokID As String) As Boolean
    Const SRC As String = "modAmbalaza.AmbImaKontraStav"

    If Len(Trim$(dokTip)) = 0 Or Len(Trim$(dokID)) = 0 Then Exit Function

    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then Exit Function

    Dim cDokT As Long, cDokI As Long, cSt As Long
    cDokT = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_TIP)
    cDokI = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_ID)
    cSt = GetColumnIndex(TBL_AMBALAZA, COL_AMB_STORNO_OD)
    ' Zatecena sveska bez nove kolone nema ni kontra-stavova -- ali odgovor
    ' "nema" bi tada bio pretpostavka. Sema je kanon, pa nedostatak kolone je
    ' kvar, i tako se i dize.
    If cDokT = 0 Then Err.Raise AMB_ERR_KNJIGA_KVAR, SRC, _
                                "tblAmbalaza nema " & COL_AMB_DOK_TIP & "."
    If cDokI = 0 Then Err.Raise AMB_ERR_KNJIGA_KVAR, SRC, _
                                "tblAmbalaza nema " & COL_AMB_DOK_ID & "."
    If cSt = 0 Then Err.Raise AMB_ERR_KNJIGA_KVAR, SRC, _
                              "tblAmbalaza nema " & COL_AMB_STORNO_OD & "."

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Len(AmbText(data(i, cSt))) > 0 Then
            If StrComp(AmbText(data(i, cDokT)), Trim$(dokTip), vbTextCompare) = 0 Then
                If StrComp(AmbText(data(i, cDokI)), Trim$(dokID), vbTextCompare) = 0 Then
                    AmbImaKontraStav = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

' ============================================================
' NABAVKA -- jedini put kojim NASE gajbe ulaze u opticaj
' ============================================================
'
' Postoji jer AMB-INV-07 i AMB-10-ODL-8 ne dozvoljavaju stanici da izda gajbe
' koje nema: njen manjak nije "tudja ambalaza" nego SKRIVEN MANJAK, pa je put
' NABAVKA, sa svojim dokumentom i svojom cenom. Na svezoj instalaciji svaka
' stanica pocinje od nule, pa bi bez ovog ulaza prvo izdavanje praznih bilo
' odbijeno -- cutover devet mesta bi blokirao rad. Odluka operatera 03.10.2026.
'
' VLASNIK BROJA JE STANICA, izvedeno a ne izmisljeno: AMB-10-ODL-3 trazi jedan
' protivpartner po dokumentu, a 6.9 izricito kaze da bi "jedan NABAVKA dokument
' nad dve stanice" tu invarijantu odmah prekrsio.
'
' NEMA PROTOKOLA POTVRDE: AmbDeficitZaPrenos vraca 0 za sve sto nije REALAN, a
' SpoljniSvet je iz te klase iskljucen (6.3 -- granica nema fizicko stanje), pa
' prenos IZ nje nikad nije u manjku.
'
' Broj se moze proslediti (operater prepisuje sa racuna) ili izracunati. Oba
' puta idu kroz istu kapiju zaglavlja -- a ona od 03.10.2026 sudi i ZAUZETOST
' u nizu (BrojOwnerTip, BrojOwnerID, dan), ne samo oblik. Zato prosledjen broj
' nije povlasten: ako je zauzet, upis pada.
' STORNO AMBALAZNOG DOKUMENTA -- revers, nabavka, otpis.
'
' Do ovog reza ga NIJE BILO: tblAmbalazaDokument je imao pisca i kolonu
' Stornirano, ali nijedan put da je okrene. Dok je revers bio skup nogu u
' tblAmbalaza, storno je isao po REDU (ekran Storno, AmbID noge). Cim revers
' postane dokument, storno mora da bude PO DOKUMENTU -- inace bi zaglavlje i
' knjiga mogli da se razidju.
'
' Redosled je isti kao kod otkupa, otpremnice i prijemnice: oznaci zaglavlje,
' pa VEZI (AMB-10-ODL-15 trazi bind POSLE stvarne izmene), pa kontra-stav.
' Kontra-stav sam proverava AMB-INV-07 i -09 nad POSLE-stanjem.
Public Function StornirajAmbDokument_TX(ByVal ambDokID As String) As Boolean
    Const SRC As String = "modAmbalaza.StornirajAmbDokument_TX"

    Dim tx As clsTransaction
    Dim red As Long, errNum As Long, errDesc As String

    On Error GoTo EH

    red = RedAmbDokumenta(ambDokID, SRC)

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_AMBALAZA_DOKUMENT
    tx.AddTableSnapshot TBL_AMBALAZA

    ' MarkRowStornirano je PRIVATE u modStorno, pa se iz ovog modula ne vidi --
    ' a ono je ionako samo omotac oko RequireUpdateCell. Zove se primitiv
    ' direktno, umesto da se tudja privatnost otvara zbog jednog poziva.
    '
    ' Kvar se nije video ni u jednoj jeftinoj kapiji: VBA kompajlira NA ZAHTEV,
    ' pa je 'Sub or Function not defined' pukao tek kad je prvi test pozvao ovu
    ' proceduru -- posle 600s i ubijenog Excela.
    ' "Da" je literal jer je STORNO_DA takodje Private u modStorno. Vrednost je
    ' ista koju citaju IsStorniranoValue i svi ostali pisci.
    RequireUpdateCell TBL_AMBALAZA_DOKUMENT, red, COL_STORNIRANO, "Da", SRC
    tx.BindSourceDocument DOK_TIP_AMBALAZA_DOKUMENT, ambDokID
    StornirajAmbalazuDokumenta tx, DOK_TIP_AMBALAZA_DOKUMENT, ambDokID

    tx.CommitTx
    Set tx = Nothing
    StornirajAmbDokument_TX = True
    Exit Function

EH:
    errNum = Err.Number
    errDesc = Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    Set tx = Nothing
    On Error GoTo 0
    LogError SRC, errDesc, errNum
    StornirajAmbDokument_TX = False
End Function

' Red zaglavlja, fail-closed. Nepostojeci ili vec storniran dokument se NE
' stornira drugi put -- druga storno operacija bi upisala drugi kontra-stav nad
' istim redovima, a AmbImaKontraStav bi ga tek posle prijavio.
Private Function RedAmbDokumenta(ByVal ambDokID As String, _
                                 ByVal SRC As String) As Long
    Dim redovi As Collection
    RequireTacnoJedan TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, Trim$(ambDokID), _
                      "AmbDokID", SRC
    Set redovi = FindRows(TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, Trim$(ambDokID))
    RedAmbDokumenta = CLng(redovi(1))

    If UCase$(Trim$(NzToText(LookupValue(TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, _
                                         Trim$(ambDokID), COL_STORNIRANO)))) = "DA" Then
        Err.Raise AMB_ERR_IDENTITET, SRC, _
                  "Ambalazni dokument '" & Trim$(ambDokID) & "' je vec storniran."
    End If
End Function

' REVERS KAO AMBALAZNI DOKUMENT (10b-2, cetvrto mesto knjizenja).
'
' Stari pisac (modDokumenta.SaveOMUlaz_TX) je knjizio SEST nogu u cetiri smera,
' a vozaca nosio kao ZIG -- pa se njegov saldo dobijao INVERZIJOM smera, sto je
' fail-open: citalac koji inverziju zaboravi dobija pogresan ZNAK, ne gresku.
' Ovde je po smeru JEDAN red koji imenuje obe strane, iz zatvorene mape
' AmbReversSmerovi.
'
' Revers je AMBALAZNI dokument (AMB-10-ODL-5), pa dobija pravo zaglavlje u
' tblAmbalazaDokument umesto brojDok rasutog po nogama. Broj je NAS
' (AmbDokBrojOwnerKlasa = SOPSTVENI, vlasnik stanica), a kapija zauzetosti iz
' AMB-10-ODL-20 radi isti posao kao zateceni RequireBrojSlobodanUNizu --
' nad istim kljucem (vlasnik, dan).
Public Function UpisiReversAmbalaze_TX(ByVal datum As Date, ByVal broj As String, _
                                       ByVal stanicaID As String, _
                                       ByVal tipAmb As String, _
                                       ByVal kolicina As Double, _
                                       ByVal smer As String, _
                                       ByVal kooperantID As String, _
                                       ByVal vozacID As String, _
                                       Optional ByVal napomena As String = "") As String
    Const SRC As String = "modAmbalaza.UpisiReversAmbalaze_TX"

    Dim tx As clsTransaction
    Dim dokID As String, brojK As String
    Dim red As Variant, odTip As String, naTip As String, vrstaK As String
    Dim errNum As Long, errDesc As String

    On Error GoTo EH

    If kolicina <= 0 Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                  "Revers trazi pozitivnu kolicinu."
    End If

    red = ReversSmerRed(smer, SRC)
    odTip = CStr(red(1))
    naTip = CStr(red(2))
    vrstaK = CStr(red(3))

    brojK = Trim$(broj)
    If Len(brojK) > 0 Then
        ' DVE KAPIJE BROJA, A NOV MODEL POKRIVA SAMO JEDNU.
        '
        ' Zauzetost (AMB-10-ODL-20) radi UpisiAmbDokument -- to je zamena za
        ' zateceni RequireBrojSlobodanUNizu. Ali OBLIK I KONTEKST broja (pripada li
        ' bas nizu te stanice i tog dana) nov model NE proverava, pa bi prelazak
        ' tiho izgubio tu kapiju.
        '
        ' Zato ostaje, i to samo za broj KOJI JE POZIVALAC ZADAO: generisan broj
        ' dolazi iz ambalaznog niza i nema REV oblik, pa bi ga ova kapija odbila.
        modBrojevi.RequireBrojUKontekstu modBrojevi.KIND_REV, stanicaID, datum, _
                                         brojK, SRC
        ' I ZAUZETOST NAD STARIM NIZOM -- ovo sam prvo ispustio.
        '
        ' Mislio sam da je AMB-10-ODL-20 (kapija u UpisiAmbDokument) zamena za
        ' RequireBrojSlobodanUNizu. Nije: ODL-20 sudi nad tblAmbalazaDokument, a
        ' stari niz zivi nad tblAmbalaza. To su DVA RAZLICITA NIZA, pa bi broj
        ' zauzet starim reversom bio slobodan za nov -- i obrnuto.
        '
        ' Dok oba oblika postoje (do 10e), oba niza moraju da vaze. Kad stari
        ' redovi nestanu, ova provera postaje mrtva i brise se sa njima.
        modBrojevi.RequireBrojSlobodanUNizu modBrojevi.KIND_REV, stanicaID, datum, _
                                            brojK, SRC
    Else
        brojK = modBrojevi.GenerateBrojAmbDokumenta(AMB_NALOG_STANICA, stanicaID, datum)
        If Len(brojK) = 0 Then
            Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                      "Broj reversa nije izracunat -- vidi Log."
        End If
    End If

    Set tx = New clsTransaction
    tx.BeginTx
    ' Obe tabele: zaglavlje i knjiga nastaju u ISTOM rollback-u (AMB-INV-08).
    tx.AddTableSnapshot TBL_AMBALAZA_DOKUMENT
    tx.AddTableSnapshot TBL_AMBALAZA

    dokID = UpisiAmbDokument(tx, AMB_DOK_REVERS, brojK, datum, _
                             AMB_NALOG_STANICA, stanicaID, napomena)

    PrenesiAmbalazu tx, datum, tipAmb, kolicina, _
                    odTip, ReversNalogID(odTip, stanicaID, kooperantID, vozacID, smer, SRC), _
                    naTip, ReversNalogID(naTip, stanicaID, kooperantID, vozacID, smer, SRC), _
                    vrstaK, DOK_TIP_AMBALAZA_DOKUMENT, dokID

    tx.CommitTx
    Set tx = Nothing

    UpisiReversAmbalaze_TX = dokID
    Exit Function

EH:
    errNum = Err.Number
    errDesc = Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    Set tx = Nothing
    On Error GoTo 0
    LogError SRC, errDesc, errNum
    Err.Raise errNum, SRC, errDesc
End Function

' POVRAT PRAZNIH OD KUPCA BEZ PRIJEMNICE -- KUPCEV DOKUMENT (AMB-10-ODL-23).
'
' Presuda operatera 06.10.2026: kupac vraca prazne gajbe i bez prijemnice, i to
' redovno. Takav povrat je NJEGOV dokument -- nosi njegov broj, vrsta je
' REVERS_PARTNERA, par je Kupac -> Vozac (lanac kupac -> vozac -> stanica,
' AMB-10-ODL-9), kretanje POVRAT_PRAZNE.
'
' ZASTO NIJE PETI SMER U AmbReversSmerovi. Ta mapa je mapa NASEG reversa: njen
' vlasnik broja je stanica, pa bi peti red tiho uveo dokument sa TUDJIM brojem i
' drugom vrstom u mapu koja o njima ne zna nista. Dva pisca se ovde razlikuju po
' vlasniku niza, ne po stilu.
'
' SVE INVARIJANTE SU U JEZGRU, ne ovde: AmbDokKretanjeProblem sudi par i
' vlasnika broja (grana jePartnerov), AmbDokDozvoljavaKretanje pusta samo
' POVRAT_PRAZNE uz ovu vrstu, UpisiAmbDokument meri zauzetost broja u opsegu
' (Kupac, KupacID, dan) po ODL-20, a PrenesiAmbalazu drzi AMB-INV-04 i -07.
' Ovaj pisac zato nosi tacno jedno pravilo koje nigde drugde ne postoji: BROJ JE
' OBAVEZAN I NE PREDLAZE SE.
'
' STORNO JE VEC POKRIVEN: StornirajAmbDokument_TX radi nad svakim ambalaznim
' dokumentom, pa ova vrsta ne trazi svoju putanju.
Public Function UpisiReversPartnera_TX(ByVal datum As Date, ByVal broj As String, _
                                       ByVal kupacID As String, _
                                       ByVal vozacID As String, _
                                       ByVal tipAmb As String, _
                                       ByVal kolicina As Double, _
                                       Optional ByVal napomena As String = "") As String
    Const SRC As String = "modAmbalaza.UpisiReversPartnera_TX"

    Dim tx As clsTransaction
    Dim dokID As String, brojK As String
    Dim errNum As Long, errDesc As String

    On Error GoTo EH

    If kolicina <= 0 Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                  "Povrat praznih od kupca trazi pozitivnu kolicinu."
    End If

    ' BROJ JE OBAVEZAN I NE PREDLAZE SE -- to je jedino pravilo koje je ovde, i
    ' jedina razlika prema UpisiReversAmbalaze_TX, koji na prazan broj zove
    ' generator. Predlog iz NASEG niza bio bi izmisljen broj TUDJE serije: papir
    ' koji operater drzi u ruci nosi kupcev broj, pa bi nas predlog bio drugi broj
    ' za isti dokument -- i ni jedan ne bi bio onaj po kome se dokument trazi.
    brojK = Trim$(broj)
    If Len(brojK) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                  "Broj je obavezan i ne predlaze se: dokument je kupcev."
    End If

    ' Oba naloga su obavezna, i razlog je razlicit za svaki: bez kupca nema
    ' vlasnika broja, bez vozaca nema odredista lanca.
    If Len(Trim$(kupacID)) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                  "Povrat od kupca trazi nalog Kupac -- on je i vlasnik broja."
    End If
    If Len(Trim$(vozacID)) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                  "Povrat od kupca trazi nalog Vozac -- lanac ide kupac -> vozac."
    End If

    Set tx = New clsTransaction
    tx.BeginTx
    ' Obe tabele: zaglavlje i knjiga nastaju u ISTOM rollback-u (AMB-INV-08).
    tx.AddTableSnapshot TBL_AMBALAZA_DOKUMENT
    tx.AddTableSnapshot TBL_AMBALAZA

    dokID = UpisiAmbDokument(tx, AMB_DOK_REVERS_PARTNERA, brojK, datum, _
                             AMB_NALOG_KUPAC, Trim$(kupacID), napomena)

    PrenesiAmbalazu tx, datum, tipAmb, kolicina, _
                    AMB_NALOG_KUPAC, Trim$(kupacID), _
                    AMB_NALOG_VOZAC, Trim$(vozacID), _
                    AMB_VK_POVRAT_PRAZNE, DOK_TIP_AMBALAZA_DOKUMENT, dokID

    tx.CommitTx
    Set tx = Nothing

    UpisiReversPartnera_TX = dokID
    Exit Function

EH:
    errNum = Err.Number
    errDesc = Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    Set tx = Nothing
    On Error GoTo 0
    LogError SRC, errDesc, errNum
    Err.Raise errNum, SRC, errDesc
End Function

' Red zatvorene mape za trazeni smer. Nepoznat smer je FAIL-CLOSED, i poruka
' nabraja sta je dozvoljeno -- isto kao zateceni Case Else u SaveOMUlaz_TX, samo
' sto spisak sada dolazi IZ MAPE, pa ne moze da se razidje sa njom.
Private Function ReversSmerRed(ByVal smer As String, ByVal SRC As String) As Variant
    Dim mapa As Variant, i As Long, spisak As String
    mapa = modAmbalazaUgovor.AmbReversSmerovi()
    For i = LBound(mapa) To UBound(mapa)
        If StrComp(Trim$(smer), CStr(mapa(i)(0)), vbTextCompare) = 0 Then
            ReversSmerRed = mapa(i)
            Exit Function
        End If
        If Len(spisak) > 0 Then spisak = spisak & ", "
        spisak = spisak & CStr(mapa(i)(0))
    Next i
    Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
              "Nepoznat smer reversa '" & Trim$(smer) & "'. Dozvoljeni: " & spisak & "."
End Function

' Koji ID popunjava stranu datog tipa. Prazan ID je FAIL-CLOSED: stari pisac je
' imao po dve takve provere u svakom od cetiri Case bloka -- osam kopija jednog
' pravila. Ovde je jedno telo, pa ne moze da se razidje po granama.
Private Function ReversNalogID(ByVal nalogTip As String, ByVal stanicaID As String, _
                               ByVal kooperantID As String, ByVal vozacID As String, _
                               ByVal smer As String, ByVal SRC As String) As String
    Select Case Trim$(nalogTip)
        Case AMB_NALOG_STANICA: ReversNalogID = Trim$(stanicaID)
        Case AMB_NALOG_KOOPERANT: ReversNalogID = Trim$(kooperantID)
        Case AMB_NALOG_VOZAC: ReversNalogID = Trim$(vozacID)
    End Select

    If Len(ReversNalogID) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                  "Revers (" & Trim$(smer) & ") trazi nalog " & Trim$(nalogTip) & _
                  ", a nije zadat."
    End If
End Function

Public Function NabaviAmbalazu_TX(ByVal datum As Date, _
                                  ByVal stanicaID As String, _
                                  ByVal tipAmb As String, _
                                  ByVal kolicina As Double, _
                                  Optional ByVal broj As String = "", _
                                  Optional ByVal napomena As String = "") As String
    Const SRC As String = "modAmbalaza.NabaviAmbalazu_TX"

    Dim tx As clsTransaction
    Dim dokID As String, brojK As String
    Dim errNum As Long, errDesc As String

    On Error GoTo EH

    If kolicina <= 0 Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                  "Nabavka trazi pozitivnu kolicinu."
    End If

    brojK = Trim$(broj)
    If Len(brojK) = 0 Then
        brojK = modBrojevi.GenerateBrojAmbDokumenta(AMB_NALOG_STANICA, stanicaID, datum)
        If Len(brojK) = 0 Then
            Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                      "Broj nabavke nije izracunat -- vidi Log."
        End If
    End If

    Set tx = New clsTransaction
    tx.BeginTx
    ' Obe tabele: zaglavlje i knjiga nastaju u ISTOM rollback-u (AMB-INV-08).
    tx.AddTableSnapshot TBL_AMBALAZA_DOKUMENT
    tx.AddTableSnapshot TBL_AMBALAZA

    ' UpisiAmbDokument sam vezuje dokument za ovu transakciju.
    dokID = UpisiAmbDokument(tx, AMB_DOK_NABAVKA, brojK, datum, _
                             AMB_NALOG_STANICA, stanicaID, napomena)

    PrenesiAmbalazu tx, datum, tipAmb, kolicina, _
                    AMB_NALOG_SPOLJNI, "", AMB_NALOG_STANICA, stanicaID, _
                    AMB_VK_NABAVKA, DOK_TIP_AMBALAZA_DOKUMENT, dokID

    tx.CommitTx
    Set tx = Nothing

    NabaviAmbalazu_TX = dokID
    Exit Function

EH:
    ' Err se cuva PRE rollback-a, i PONOVO se dize SA ISTIM BROJEM: pozivalac
    ' (UI) razlikuje kapije po kodu greske, pa bi zamena broja unistila tu
    ' informaciju -- i protokol potvrde deficita na drugim putanjama.
    errNum = Err.Number
    errDesc = Err.description
    LogErr SRC

    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    Set tx = Nothing
    On Error GoTo 0

    Err.Raise errNum, SRC, errDesc
End Function

' AMB-INV-08: nijedan upis u knjigu ne nastaje van vlasnistva transakcije
' izvornog dokumenta.
'
' Ovo je bila STATICKA kapija u planu, i napisana je -- pa pala na svom drugom
' nivou dokaza. Pravilo "neki predak u pozivnom lancu poseduje tx sa snapshotom"
' je nad 4120 procedura i 17 vlasnika uvek istinito ako se ide dovoljno visoko:
' skinuta su oba AddTableSnapshot iz modOtkup, a kapija je ostala zelena. Alat
' koji se moze prevariti dubinom ne sme da cuva invarijantu.
'
' Zato pisac TRAZI transakciju i sam pita. Razlika nije stilska:
'   - poziv bez tx je COMPILE ERROR, ne nalaz koji se moze ignorisati;
'   - nema grafa, nema dubine, nema laznog zelenog ni laznog nalaza;
'   - sabotaza je prava: skini snapshot, pisac padne po imenu.
Private Sub RequireAmbTxVlasnistvo(ByVal tx As clsTransaction, _
                                   ByVal tabela As String, _
                                   ByVal sourceName As String)
    If tx Is Nothing Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, sourceName, _
                  "AMB-INV-08: upis u " & tabela & " bez transakcije. Pisac " & _
                  "trazi clsTransaction izvornog dokumenta."
    End If
    If Not tx.ImaSnapshot(tabela) Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, sourceName, _
                  "AMB-INV-08: transakcija ne snapshotuje " & tabela & _
                  " -- rollback izvornog dokumenta ne bi vratio ambalazu."
    End If
End Sub

' Tacka 3 iste invarijante: izvorni dokument mora biti vezan BAS za ovu
' transakciju. Snapshot tabele dokazuje da se knjiga MOZE vratiti; ovo dokazuje
' da se vraca ZAJEDNO sa dokumentom koji ju je izazvao.
Private Sub RequireAmbTxIzvorniDokument(ByVal tx As clsTransaction, _
                                        ByVal dokTip As String, _
                                        ByVal dokID As String, _
                                        ByVal sourceName As String)
    If Not tx.OwnsSourceDocument(dokTip, dokID) Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, sourceName, _
                  "AMB-INV-08: izvorni dokument " & Trim$(dokTip) & " '" & _
                  Trim$(dokID) & "' nije vezan za ovu transakciju " & _
                  "(BindSourceDocument) -- knjiga i dokument ne dele vlasnika, " & _
                  "pa rollback dokumenta ne bi vratio ambalazu."
    End If

    ' Vezivanje dokazuje IDENTITET transakcije, ne i da je tabela tog dokumenta
    ' u njenom snapshotu. Bez ovoga prolazi bind nad Otkup-om uz snapshot samo
    ' knjige -- rollback tada nije zajednicki. Nepoznat tip je fail-closed.
    Dim izvornaTbl As String
    izvornaTbl = modAmbalazaUgovor.AmbIzvornaTabela(dokTip)
    If Len(izvornaTbl) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, sourceName, _
                  "AMB-INV-08: tip izvornog dokumenta '" & Trim$(dokTip) & _
                  "' nema poznatu izvornu tabelu -- ne moze se dokazati da " & _
                  "knjiga i dokument dele rollback."
    End If
    If Not tx.ImaSnapshot(izvornaTbl) Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, sourceName, _
                  "AMB-INV-08: transakcija ne snapshotuje " & izvornaTbl & _
                  " -- izvorni dokument je vezan, ali se ne bi vratio " & _
                  "rollback-om zajedno sa knjigom."
    End If
End Sub

' ============================================================
' PRENESI AMBALAZU -- javni ulaz u knjigu
' ============================================================
'
' Vraca AmbID glavnog reda. Jedan poziv moze da upise do TRI reda, i svaki je
' posledica imenovane invarijante:
'
'   pokrice deficita   SpoljniSvet -> izvor   ULAZ_TUDJE_AMBALAZE   (AMB-INV-07)
'   trazeni prenos     od -> na               trazena vrsta
'   ostatak podele     od -> na               IZDATA_PRAZNA         (AMB-INV-09)
'
' potvrdaDeficita: -1 znaci "nije data". Nula NIJE sentinela -- deficit nula ne
' trazi potvrdu, pa bi 0 bila dvosmislena vrednost.
Public Function PrenesiAmbalazu(ByVal tx As clsTransaction, _
                                ByVal datum As Date, ByVal tipAmb As String, _
                                ByVal kolicina As Double, _
                                ByVal odTip As String, ByVal odID As String, _
                                ByVal naTip As String, ByVal naID As String, _
                                ByVal vrsta As String, _
                                ByVal dokTip As String, ByVal dokID As String, _
                                Optional ByVal potvrdaDeficita As Double = -1) As String
    Const SRC As String = "modAmbalaza.PrenesiAmbalazu"

    RequireKnjigaSchema SRC
    modAmbalazaUgovor.RequireAmbPrenos odTip, odID, naTip, naID, kolicina, tipAmb, vrsta, SRC

    ' ULAZ_TUDJE_AMBALAZE nije zahtev nego POSLEDICA -- generise je ovaj pisac kao
    ' pokrice deficita. Da je i zahtev, pokrice i eksplicitan zahtev nad istim
    ' dokumentom delili bi par i vrstu, pa bi jedan progutao drugi.
    If Not modAmbalazaUgovor.AmbVrstaJeZahtev(vrsta) Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                  "'" & Trim$(vrsta) & "' nije zahtev nego posledica -- pisac je generise sam."
    End If

    ' IDENTITET DOKUMENTA SE NE PROVERAVA OVDE. Ista provera stoji u
    ' UpisiRedKnjige, kroz koji prolazi SVAKI red -- i pokrice deficita i ostatak
    ' podele. Kopija na ulazu bila bi placebo: nijedna sabotaza je ne bi mogla
    ' oboriti, jer bi jezgro odgovorilo isto. Invarijanta zivi u jezgru.
    Dim vrstaK As String
    vrstaK = modAmbalazaUgovor.AmbVrstaKanon(vrsta)

    ' VEZA DOKUMENT <-> KRETANJE. Vazi samo za tblAmbalazaDokument: za robne
    ' dokumente (otkup, otpremnica, prijemnica) tipovi nisu zatvoren skup, pa
    ' ugovor o njima namerno ne tvrdi nista.
    Dim dokVrsta As String, dokOwnerTip As String, dokOwnerID As String
    If StrComp(Trim$(dokTip), DOK_TIP_AMBALAZA_DOKUMENT, vbTextCompare) = 0 Then
        AmbDokZaglavlje dokID, dokVrsta, dokOwnerTip, dokOwnerID
        If Not modAmbalazaUgovor.AmbDokDozvoljavaKretanje(dokVrsta, vrstaK) Then
            Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                      "Dokument vrste " & dokVrsta & " ne nosi kretanje " & vrstaK & "."
        End If
    End If

    ' AMB-10-ODL-22: robni dokument koji je I SAM partnerov objavljuje vlasnika
    ' svog broja. Prijemnica je takav: eksterna je, njen broj je kupcev, i
    ' povrat praznih se knjizi POD TIM brojem (operater, 05.10.2026). Bez ovoga
    ' bi obrnuta kapija odbila prijemnicu -- merila bi VRSTU dokumenta kao
    ' zamenu za vlasnika broja, a to je zamena koju je ODL-22 pobio.
    If Len(dokOwnerTip) = 0 Then
        AmbRobniZaglavlje dokTip, dokID, dokOwnerTip, dokOwnerID
    End If

    ' AMB-10-ODL-9/-10: vrsta, vlasnik broja i par naloga se gledaju ZAJEDNO.
    ' Obrnuta kapija vazi i kad dokument NIJE ambalazni (dokVrsta ostaje prazna):
    ' povrat praznih od kupca mora da nosi kupcev broj, pa ga dokument koji
    ' vlasnika broja ne objavi ne sme nositi.
    Dim parProblem As String
    parProblem = modAmbalazaUgovor.AmbDokKretanjeProblem( _
        dokVrsta, dokOwnerTip, dokOwnerID, odTip, odID, naTip, vrstaK)
    If Len(parProblem) > 0 Then
        Err.Raise AMB_ERR_IDENTITET, SRC, parProblem
    End If

    ' AMB-INV-04, idempotencija: isti zahtev nad istim dokumentom i parem naloga.
    Dim vecUpisano As Double, vecID As String
    vecUpisano = ZbirZahteva(dokTip, dokID, datum, tipAmb, odTip, odID, naTip, naID, _
                             vrstaK, vecID, SRC)
    If Len(vecID) > 0 Then
        If vecUpisano = kolicina Then
            PrenesiAmbalazu = vecID
            Exit Function
        End If

        Err.Raise AMB_ERR_IDENTITET, SRC, _
                  "AMB-INV-04: isti dogadjaj je vec knjizen sa " & CStr(vecUpisano) & _
                  ", a trazi se " & CStr(kolicina) & " (" & Trim$(dokTip) & " '" & _
                  Trim$(dokID) & "', " & vrstaK & ")."
    End If

    ' AMB-INV-07: nijedan realan nalog ne sme posle commit-a biti ispod nule.
    Dim deficit As Double, pokriceZaUpis As Double
    deficit = AmbDeficitZaPrenos(odTip, odID, tipAmb, kolicina)

    If deficit > 0 Then
        Dim pokrice As String
        pokrice = modAmbalazaUgovor.AmbPokriceProblem(odTip, odID)
        If Len(pokrice) > 0 Then
            Err.Raise AMB_ERR_DEFICIT_NEPOKRIV, SRC, _
                      pokrice & " Manjak: " & CStr(deficit) & " '" & Trim$(tipAmb) & "'."
        End If

        If potvrdaDeficita < 0 Then
            Err.Raise AMB_ERR_POTVRDA_DEFICITA, SRC, _
                      "Nalog " & Trim$(odTip) & " '" & Trim$(odID) & "' nema " & CStr(kolicina) & _
                      " '" & Trim$(tipAmb) & "'. Manjak " & CStr(deficit) & " ulazi u opticaj kao " & _
                      "tudja ambalaza -- potvrdi tacno taj broj."
        End If

        ' Potvrda se meri prema SVEZE izracunatom deficitu: izmedju pitanja i
        ' odgovora stanje se moglo promeniti drugim unosom, pa stara potvrda ne vazi.
        If potvrdaDeficita <> deficit Then
            Err.Raise AMB_ERR_POTVRDA_DEFICITA, SRC, _
                      "Potvrda " & CStr(potvrdaDeficita) & " se ne slaze sa manjkom " & _
                      CStr(deficit) & " -- stanje se promenilo, potvrdi ponovo."
        End If

        ' POKRICE SE UPISUJE POSLE TRAZENOG PRENOSA, i to je sada deo ugovora:
        ' AMB-INV-10 poredi odredisten pokrica sa ZAKLJUCANIM parom dokumenta, pa par
        ' mora postojati pre njega. Redosled unutar transakcije je slobodan --
        ' AMB-INV-07 govori o stanju POSLE commit-a, ne o putu do njega.
        pokriceZaUpis = deficit
    End If

    ' AMB-INV-09: ne moze se vratiti vise tudje ambalaze nego sto je uzeto. Visak
    ' nije vracanje nego NOVO zaduzenje partnera -- podelu radi pisac, jer samo on
    ' cita obavezu pouzdano u trenutku upisa.
    Dim glavna As Double, ostatak As Double
    glavna = kolicina
    ostatak = 0

    If StrComp(vrstaK, AMB_VK_VRACANJE_TUDJE, vbTextCompare) = 0 Then
        Dim obaveza As Double
        obaveza = AmbObavezaPartneru(naTip, naID, tipAmb)

        ' NEGATIVNA OBAVEZA SE NE SPUSTA NA NULU, NEGO PADA. Ona znaci da je
        ' AMB-INV-09 prekrsena PRE ovog poziva, a tisina bi je pretvorila u
        ' normalno stanje. Do njega 10d MOZE da dodje: storno ULAZA tudje
        ' ambalaze cija je obaveza vec zatvorena vracanjem daje -N. Odgovor na to
        ' je da se takav storno odbije, i to je zahtev koji 10d nasledjuje -- ne
        ' da ga ovaj pisac prekrije.
        If obaveza < 0 Then
            Err.Raise AMB_ERR_KNJIGA_KVAR, SRC, _
                      "Obaveza prema " & Trim$(naTip) & " '" & Trim$(naID) & "' je " & _
                      CStr(obaveza) & " -- AMB-INV-09 je prekrsena pre ovog poziva."
        End If

        If kolicina > obaveza Then
            glavna = obaveza
            ostatak = kolicina - obaveza
        End If
    End If

    Dim glavniID As String
    If glavna > 0 Then
        glavniID = UpisiRedKnjige(tx, datum, tipAmb, glavna, odTip, odID, naTip, naID, _
                                  dokTip, dokID, vrstaK, "", SRC)
    End If

    If ostatak > 0 Then
        Dim ostatakID As String
        ostatakID = UpisiRedKnjige(tx, datum, tipAmb, ostatak, odTip, odID, naTip, naID, _
                                   dokTip, dokID, AMB_VK_IZDATA_PRAZNA, "", SRC)
        If Len(glavniID) = 0 Then glavniID = ostatakID
    End If

    If Len(glavniID) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_KVAR, SRC, "Prenos nije upisao ni jedan red."
    End If

    If pokriceZaUpis > 0 Then
        UpisiRedKnjige tx, datum, tipAmb, pokriceZaUpis, _
                       AMB_NALOG_SPOLJNI, "", odTip, odID, _
                       dokTip, dokID, AMB_VK_ULAZ_TUDJE, "", SRC
    End If

    PrenesiAmbalazu = glavniID
End Function
