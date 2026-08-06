Attribute VB_Name = "modMalina"
'Attribute VB_Name = "modMalina"
Option Explicit

' ============================================================
' modMalina - malina-mod master-data glue
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
' je isti obrazac kao modConfig.SetConfigValue -- schema-robustno, bez
' pozicijskog Array-a i bez test-only BlankRow/SetField helpera.
' ============================================================

' AUD-046: canonical provera "postoji li stanica-mirror par".
'
' Mirror je par: red u tblStanice I red u tblVozaci sa ISTIM ID-em. Pozivaoci
' koji stampaju VozacID := StanicaID (modMasterSync.StampVozacFromStanicaForMalina,
' modAutoHladnjaca.AutoChainHladnjaca) MORAJU ovo pitati pre upisa -- inace
' dokument dobije FK bez pokrica, a svaki join na tblVozaci (izvestaji, ambalaza,
' banka nalozi) vraca prazno ime.
'
' NAPOMENA: modBrojevi.IsStanicaMirrorVozac je namerno DRUGO pitanje ("da li je
' ovaj vozacID zapravo stanica" -> "S" prefiks u broju) i gleda samo tblStanice.
' Ovde se proverava par, i to je jedini validator pred upis FK-a.
Public Function IsManagedStationMirror(ByVal entityID As String) As Boolean
    On Error GoTo EH

    Dim sid As String: sid = Trim$(entityID)
    If sid = "" Then Exit Function

    If Len(Trim$(nz(LookupValue(TBL_STANICE, "StanicaID", sid, "StanicaID"), ""))) = 0 Then
        Exit Function
    End If

    If Len(Trim$(nz(LookupValue(TBL_VOZACI, "VozacID", sid, "VozacID"), ""))) = 0 Then
        Exit Function
    End If

    IsManagedStationMirror = True
    Exit Function

EH:
    LogErr "modMalina.IsManagedStationMirror", "entityID=" & entityID
    IsManagedStationMirror = False
End Function

' Idempotentno pravi par-vozaca za datu stanicu. Vraca True samo ako je
' kreiran NOV red (False ako vec postoji, ako nije malina mod, ili ako je
' stanicaID prazan).
'
' AUD-046: greska se NE guta. Do sada je EH samo logovao (za razliku od
' BackfillVozacMirrorsForMalina koji re-raise-uje), a AppendRow = 0 je prolazio
' kao uspeh -- pozivaoci su posle toga stampali VozacID := StanicaID verujuci da
' mirror postoji. Zato: stanica mora postojati i biti jedinstvena, AppendRow = 0
' je greska, i EH re-raise-uje. Pozivaoci koji smeju da nastave bez mirror-a
' (frmStammdaten unos, best-effort auto-lanac) sami drze On Error Resume Next i
' posle poziva pitaju IsManagedStationMirror.
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
    If Len(Trim$(nz(LookupValue(TBL_VOZACI, "VozacID", sid, "VozacID"), ""))) > 0 Then
        Exit Function
    End If

    ' AUD-046: mirror za nepostojecu stanicu je otrovan podatak -- pravio bi
    ' vozaca kojem nista ne odgovara i legitimizovao FK koji nema pokrice.
    Dim stRows As Collection
    Set stRows = FindRows(TBL_STANICE, "StanicaID", sid)

    If stRows Is Nothing Then
        Err.Raise vbObjectError + 8410, SRC, _
                  "FindRows je vratio Nothing za " & TBL_STANICE & "; StanicaID=" & sid
    End If

    If stRows.count = 0 Then
        Err.Raise vbObjectError + 8411, SRC, _
                  "Stanica ne postoji -- vozac mirror se ne kreira. StanicaID=" & sid
    End If

    If stRows.count > 1 Then
        Err.Raise vbObjectError + 8412, SRC, _
                  "Duplikat StanicaID u " & TBL_STANICE & " -- mirror nije jednoznacan. StanicaID=" & sid & _
                  "; Count=" & CStr(stRows.count)
    End If

    ' Neaktivna stanica NIJE blokada (backfill prolazi kroz sve stanice, a mirror
    ' je master-podatak), ali se loguje da se zna odakle je red dosao.
    Dim cAkt As Long
    cAkt = GetColumnIndex(TBL_STANICE, "Aktivan")
    If cAkt > 0 Then
        Dim stData As Variant
        stData = GetTableData(TBL_STANICE)
        If Not IsEmpty(stData) Then
            If UCase$(Trim$(CStr(nz(stData(stRows(1), cAkt), "")))) <> UCase$(STATUS_AKTIVAN) Then
                LogWarn SRC, "Mirror se pravi za NEAKTIVNU stanicu. StanicaID=" & sid
            End If
        End If
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

    Dim cI As Long
    cI = GetColumnIndex(TBL_VOZACI, "Telefon")
    If cI > 0 Then rowData(cI) = telefon
    cI = GetColumnIndex(TBL_VOZACI, "Aktivan")
    If cI > 0 Then rowData(cI) = STATUS_AKTIVAN

    ' AUD-046: AppendRow = 0 znaci da red NIJE upisan. Tihi False je pozivaoce
    ' ostavljao u uverenju "mirror postoji (samo ga nisam ja kreirao)".
    If AppendRow(TBL_VOZACI, rowData) = 0 Then
        Err.Raise vbObjectError + 8413, SRC, _
                  "AppendRow nije upisao vozac mirror u " & TBL_VOZACI & "; VozacID=" & sid
    End If

    EnsureVozacMirrorForStanica = True
    LogInfo SRC, "Vozac mirror kreiran: VozacID=" & sid
    Exit Function

EH:
    ' AUD-046: re-raise (isti model kao BackfillVozacMirrorsForMalina).
    ' Pozivalac odlucuje da li nastavlja -- ali NE sme da nastavi u tisini.
    EnsureVozacMirrorForStanica = False
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.description
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
        Dim sid As String: sid = Trim$(nz(data(r, cId), ""))
        If sid <> "" Then
            Dim naziv As String: naziv = nz(data(r, cNaziv), "")
            Dim mesto As String: mesto = ""
            If cMesto > 0 Then mesto = nz(data(r, cMesto), "")
            Dim tel As String: tel = ""
            If cTel > 0 Then tel = nz(data(r, cTel), "")
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

