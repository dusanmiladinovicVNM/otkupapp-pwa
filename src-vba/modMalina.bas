Attribute VB_Name = "modMalina"
Option Explicit

' ============================================================
' modMalina – malina-mod master-data glue
'
' U malina modu (IsMalinaMode, vidi modConfig) vazi: otkupac == stanica
' == vozac. Da izvestaji/ambalaza koji joinuju na tblVozaci imaju naziv,
' svaka stanica dobija "shadow" vozaca sa VozacID := StanicaID (isti string).
'
' Backend toka (VozacID:=StanicaID na otkupu, auto-zbirna iz otpremnice) je
' u modMasterSync; ovde je samo master-data mirror koji nema prirodan dom
' (modStammdatenSync je export-only, frmStammdaten je UI shell).
'
' Append po nazivu kolone (ReDim na ListColumns.Count + RequireColumnIndex)
' je isti obrazac kao modConfig.SetConfigValue — schema-robustno, bez
' pozicijskog Array-a i bez test-only BlankRow/SetField helpera.
' ============================================================

' Idempotentno pravi par-vozaca za datu stanicu. Vraca True samo ako je
' kreiran NOV red (False ako vec postoji, ako nije malina mod, ili ako je
' stanicaID prazan).
Public Function EnsureVozacMirrorForStanica(ByVal stanicaID As String, _
                                            ByVal ime As String, _
                                            ByVal prezime As String, _
                                            ByVal telefon As String) As Boolean
    Const SRC As String = "modMalina.EnsureVozacMirrorForStanica"

    On Error GoTo EH

    EnsureVozacMirrorForStanica = False
    If Not IsMalinaMode() Then Exit Function

    Dim sid As String: sid = Trim$(stanicaID)
    If sid = "" Then Exit Function

    ' Idempotencija: vozac sa VozacID == StanicaID vec postoji?
    If Len(Trim$(Nz(LookupValue(TBL_VOZACI, "VozacID", sid, "VozacID"), ""))) > 0 Then
        Exit Function
    End If

    Dim lo As ListObject
    Set lo = GetTable(TBL_VOZACI)
    If lo Is Nothing Then
        Err.Raise vbObjectError + 8400, SRC, "Tabela ne postoji: " & TBL_VOZACI
    End If

    Dim rowData() As Variant
    ReDim rowData(1 To lo.ListColumns.count)
    rowData(RequireColumnIndex(TBL_VOZACI, "VozacID", SRC)) = sid
    rowData(RequireColumnIndex(TBL_VOZACI, "Ime", SRC)) = ime
    rowData(RequireColumnIndex(TBL_VOZACI, "Prezime", SRC)) = prezime

    Dim ci As Long
    ci = GetColumnIndex(TBL_VOZACI, "Telefon")
    If ci > 0 Then rowData(ci) = telefon
    ci = GetColumnIndex(TBL_VOZACI, "Aktivan")
    If ci > 0 Then rowData(ci) = STATUS_AKTIVAN

    If AppendRow(TBL_VOZACI, rowData) > 0 Then
        EnsureVozacMirrorForStanica = True
        LogInfo SRC, "Vozac mirror kreiran: VozacID=" & sid
    End If
    Exit Function

EH:
    LogErr SRC
End Function

' Jednokratni backfill: za svaku stanicu napravi par-vozaca ako ga nema.
' Vraca broj novo-kreiranih. Operater ga pokrene jednom pri ukljucivanju
' malina moda (postojece stanice; nove ide kroz frmStammdaten hook).
Public Function BackfillVozacMirrorsForMalina() As Long
    Const SRC As String = "modMalina.BackfillVozacMirrorsForMalina"

    On Error GoTo EH

    BackfillVozacMirrorsForMalina = 0
    If Not IsMalinaMode() Then Exit Function

    Dim data As Variant
    data = GetTableData(TBL_STANICE)
    If IsEmpty(data) Then Exit Function

    Dim cId As Long, cNaziv As Long, cMesto As Long, cTel As Long
    cId = RequireColumnIndex(TBL_STANICE, "StanicaID", SRC)
    cNaziv = RequireColumnIndex(TBL_STANICE, "Naziv", SRC)
    cMesto = GetColumnIndex(TBL_STANICE, "Mesto")
    cTel = GetColumnIndex(TBL_STANICE, "Telefon")

    Dim r As Long, cnt As Long
    For r = 1 To UBound(data, 1)
        Dim sid As String: sid = Trim$(Nz(data(r, cId), ""))
        If sid <> "" Then
            Dim naziv As String: naziv = Nz(data(r, cNaziv), "")
            Dim mesto As String: mesto = ""
            If cMesto > 0 Then mesto = Nz(data(r, cMesto), "")
            Dim tel As String: tel = ""
            If cTel > 0 Then tel = Nz(data(r, cTel), "")
            If EnsureVozacMirrorForStanica(sid, naziv, mesto, tel) Then cnt = cnt + 1
        End If
    Next r

    BackfillVozacMirrorsForMalina = cnt
    LogInfo SRC, "Vozac mirror backfill created=" & CStr(cnt)
    Exit Function

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.description
End Function
