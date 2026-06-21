Attribute VB_Name = "modAmbalaza"
'Attribute VB_Name = "modAmbalaza"
Option Explicit

' ============================================================
' modAmbalaza v2.2 – Verpackung-Tracking / Ambalaza ledger
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

' ============================================================
' Helpers
' ============================================================

Private Function AmbText(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
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
                  "Kolicina ambalaze ne sme biti negativna."
    End If

    ' Nula je legalan no-op.
    If kolicina = 0 Then Exit Sub

    If Len(Trim$(tipAmb)) = 0 Then
        Err.Raise vbObjectError + 4402, sourceName, _
                  "Tip ambalaze je obavezan kada postoji kolicina."
    End If

    If Not IsValidAmbSmer(smer) Then
        Err.Raise vbObjectError + 4403, sourceName, _
                  "Neispravan smer ambalaze: " & smer
    End If

    If Len(Trim$(entitetID)) = 0 Then
        Err.Raise vbObjectError + 4404, sourceName, _
                  "EntitetID je obavezan za ambalazu."
    End If

    If Len(Trim$(entitetTip)) = 0 Then
        Err.Raise vbObjectError + 4405, sourceName, _
                  "EntitetTip je obavezan za ambalazu."
    End If
End Sub

Private Sub RequireAmbalazaSchema(ByVal sourceName As String)
    ' Fail-fast schema guard. Ne koristimo indekse ovde za rowData,
    ' ali eksplicitno proveravamo da tabela ima ocekivane kolone.
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

Public Sub TrackAmbalaza(ByVal datum As Date, ByVal tipAmb As String, _
                         ByVal kolicina As Long, ByVal smer As String, _
                         ByVal entitetID As String, ByVal entitetTip As String, _
                         Optional ByVal vozacID As String = "", _
                         Optional ByVal dokumentID As String = "", _
                         Optional ByVal dokumentTip As String = "")

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

    If AppendRow(TBL_AMBALAZA, rowData) <= 0 Then
        Err.Raise vbObjectError + 4407, SRC, _
                  "AppendRow nije uspeo za tblAmbalaza."
    End If

    Exit Sub

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogErr SRC
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
                          "Neispravna kolicina ambalaze u redu " & CStr(i)
            End If

            If Not dict.Exists(key) Then dict.Add key, 0&

            Select Case AmbText(data(i, colSmer))
                Case AMB_SMER_ULAZ
                    dict(key) = CLng(dict(key)) + CLng(data(i, colKol))

                Case AMB_SMER_IZLAZ
                    dict(key) = CLng(dict(key)) - CLng(data(i, colKol))

                Case Else
                    Err.Raise vbObjectError + 4412, SRC, _
                              "Neispravan smer ambalaze u redu " & CStr(i)
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
                          "Neispravna kolicina ambalaze u redu " & CStr(i)
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
                              "Neispravan smer ambalaze u redu " & CStr(i)
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
' ili tabela ne postoji — da postojece instalacije ne ostanu bez opcija.
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

    Dim cV As Long, cS As Long, cT As Long
    cV = GetColumnIndex(TBL_KULTURE, "VrstaVoca")
    cS = GetColumnIndex(TBL_KULTURE, "SortaVoca")
    cT = GetColumnIndex(TBL_KULTURE, COL_KUL_TIP_AMBALAZE)
    If cV = 0 Or cT = 0 Then Exit Function

    Dim i As Long, hit As String
    For i = 1 To UBound(data, 1)
        If StrComp(Trim$(Nz(data(i, cV))), Trim$(vrsta), vbTextCompare) = 0 Then
            If Len(Trim$(sorta)) = 0 Or cS = 0 _
               Or StrComp(Trim$(Nz(data(i, cS))), Trim$(sorta), vbTextCompare) = 0 Then
                hit = Trim$(Nz(data(i, cT)))
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

