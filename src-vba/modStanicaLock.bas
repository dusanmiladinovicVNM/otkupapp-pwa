Attribute VB_Name = "modStanicaLock"
'Attribute VB_Name = "modStanicaLock"
Option Explicit

' ============================================================
' modStanicaLock -- per-stanica lock za PWA + bulk push pri unlock-u.
'
' Lock store: SyncControl tab u Stammdaten spreadsheet-u (postojeca
' infrastruktura, dosad koriscena samo za MASTER_SYNC_*). Per-stanica
' kljucevi:
'   STANICA_LOCK_{stanicaID}_LOCK       = YES/NO
'   STANICA_LOCK_{stanicaID}_UPDATED_AT = ISO timestamp (TTL 10 min)
'   STANICA_LOCK_{stanicaID}_MESSAGE    = poruka za PWA overlay
'   STANICA_LOCK_{stanicaID}_OWNER      = "VBA"
'
' Heartbeat: Application.OnTime tick svakih 90s dok je forma aktivna.
' TTL stale-after: 10 min (postojeci MASTER_SYNC_LOCK_MAX_AGE_MIN u GAS-u).
'
' Bulk push: pri release-u (stanica change ili form close), iteriraj
' tblOtkup za (StanicaID = X, Datum = Y, ClientRecordID is empty),
' push svaki red u OTK-{stanicaID} sheet via AppendRowToSheet, posle
' uspeha upisi ClientRecordID = "VBA:" & OtkupID kao "uspesno pushed"
' marker (idempotent retry pri sledecem release-u).
'
' Public API:
'   AcquireStanicaLock(stanicaID, datum) -> Boolean
'   ReleaseStanicaLock(stanicaID) + BulkPushPendingForStanica(stanicaID, datum)
'   ChangeStanica(oldStanica, oldDatum, newStanica, newDatum) -- atomic switch
'   HeartbeatStanicaLock                                       -- interna, Application.OnTime
'   StartHeartbeatTimer / StopHeartbeatTimer
'   CleanupOrphanedLocks                                       -- Workbook_Open
'   GetActiveStanica                                            -- diagnostic
' ============================================================

Private Const SYNC_CONTROL_TAB As String = "SyncControl"
Private Const HEARTBEAT_INTERVAL_SEC As Long = 90
Private Const SYNC_STATUS_MASTER_GOOGLE As String = "Master>Google"
Private Const DEVICE_ID_VBA As String = "vba"

' Module-level state
Private gActiveStanica As String
Private gActiveDatum As Date
Private gHeartbeatScheduled As Boolean
Private gNextHeartbeatTime As Date

' TestHook_OtkStavkeSimulacija: kad je postavljen, append u OTK_STAVKE ide u ovu
' kolekciju umesto u Google, a poziv broj mSimPadNa se ponasa kao neuspeo.
Private mSimRedovi As Collection
Private mSimPadNa As Long
Private mSimPoziva As Long

' ============================================================
' PUBLIC -- Acquire
' ============================================================

' Acquire lock za novu stanicu. Vraca True ako je acquire uspeo.
' Ako je vec neka stanica aktivna, prvo je release-uje + bulk push.
Public Function AcquireStanicaLock(ByVal stanicaID As String, _
                                    ByVal datum As Date) As Boolean
    Const SRC As String = "AcquireStanicaLock"
    
    On Error GoTo EH
    
    If Len(Trim$(stanicaID)) = 0 Then
        LogError SRC, "stanicaID je prazan."
        AcquireStanicaLock = False
        Exit Function
    End If
    
    ' DODATO: desktop-only deploy nema cloud -- lock nije relevantan.
    ' Vracamo True da pozivalac (frmOtkup) misli da je lock acquired
    ' i nastavi normalno; nikakav SyncControl write se ne desava.
    If Not IsCloudSyncEnabled() Then
        gActiveStanica = stanicaID
        gActiveDatum = datum
        AcquireStanicaLock = True
        Exit Function
    End If
    
    ' Ako je vec aktivna ova stanica, refresh-uj UPDATED_AT (heartbeat efekt)
    If gActiveStanica = stanicaID And gActiveDatum = datum Then
        AcquireStanicaLock = UpdateStanicaLockTimestamp(stanicaID)
        Exit Function
    End If
    
    ' Ako je aktivna druga stanica, prvo release + bulk push
    If Len(gActiveStanica) > 0 Then
        Call ReleaseStanicaLockInternal(gActiveStanica, gActiveDatum, True)
    End If
    
    ' Postavi lock keys
    Dim updates As Object
    Set updates = CreateObject("Scripting.Dictionary")
    
    Dim prefix As String: prefix = "STANICA_LOCK_" & stanicaID & "_"
    updates(prefix & "LOCK") = "YES"
    updates(prefix & "UPDATED_AT") = Format$(Now, "yyyy-mm-dd\Thh:nn:ss")
    updates(prefix & "MESSAGE") = "Stanica je trenutno u kancelarijskoj obradi. Sacekajte zavrsetak unosa."
    updates(prefix & "OWNER") = DEVICE_ID_VBA
    
    If Not ApplySyncControlUpdates(updates) Then
        LogError SRC, "ApplySyncControlUpdates failed za stanicu=" & stanicaID
        AcquireStanicaLock = False
        Exit Function
    End If
    
    gActiveStanica = stanicaID
    gActiveDatum = datum
    
    StartHeartbeatTimer
    
    LogInfo SRC, "Lock acquired: stanica=" & stanicaID & " datum=" & Format$(datum, "yyyy-mm-dd")
    AcquireStanicaLock = True
    Exit Function

EH:
    LogErr SRC, "stanica=" & stanicaID
    AcquireStanicaLock = False
End Function

' ============================================================
' PUBLIC -- Release
' ============================================================

' Release lock + bulk push pending redova za tu stanicu/datum.
' Pozivaj iz frmOtkup form_close i pri stanica change.
Public Sub ReleaseStanicaLock(ByVal stanicaID As String)
    Call ReleaseStanicaLockInternal(stanicaID, gActiveDatum, True)
End Sub

' Otpusti lock stanice koju OVA sesija drzi, BEZ argumenta. Postoji da bi
' self-update mogao da ga pozove KASNO VEZANO (CallOptional) iz teardown-a --
' modSelfUpdate je frozen i ne sme da early-bind-uje updatable module (zamka #24),
' a ReleaseStanicaLock trazi stanicaID koji pozivalac tamo nema.
'
' MORA se pozvati PRE prve izmene koda: gActiveStanica je module-level, a izmena
' bilo kog modula brise module-level stanje u svim modulima (mereno, zamka #29).
' Posle toga se vise ne zna koji lock drzimo, pa se ne moze ni otpustiti.
' doBulkPush = FALSE, i to NIJE detalj. ReleaseStanicaLockInternal sa True prvo
' pozove BulkPushPendingForStanica, koji APPENDUJE pending otkupne redove na
' Google sheet pa TEK ONDA lokalno upise ClientRecordID = "VBA:" & OtkupID.
' Taj lokalni marker je JEDINI idempotency guard za ponovni push
' ("If Len(crid) > 0 Then GoTo NextRow").
'
' Ciscenje pred self-update ne sme da nosi poslovni side effect: push bi se desio
' PRE nego sto se zna da ce update uspeti, a AbortSelfUpdate zatvara svesku sa
' SaveChanges:=False -- cloud append bi ostao, lokalni marker bi se odbacio, i
' sledeca sesija bi iste redove poslala PONOVO. Duplikati u cloud-u iz rutine
' koja je trebalo samo da otpusti lock.
'
' Sa False radi se tacno ono sto lifecycle cleanup treba: LOCK=NO, cist
' OWNER/MESSAGE, reset gActiveStanica/gActiveDatum, StopHeartbeatTimer. Pending
' redovi ostaju pending i idu kroz normalan poslovni tok.
Public Sub ReleaseActiveStanicaLock()
    If Len(gActiveStanica) = 0 Then Exit Sub
    Call ReleaseStanicaLockInternal(gActiveStanica, gActiveDatum, False)
End Sub

' Otpustanje na IZLASKU iz aplikacije (ThisWorkbook.Workbook_BeforeClose).
'
' Bulk push se radi samo kad je izlaz NORMALAN. Pri prekinutom VBA importu
' sveska se zatvara sa SaveChanges:=False, a to je tacno kombinacija opisana
' iznad: BulkPushPendingForStanica APPENDUJE red u cloud pa TEK ONDA lokalno
' upise ClientRecordID. Odbacene promene odnose taj marker, cloud red ostaje, i
' sledeca sesija salje iste redove PONOVO -- duplikati koje rollback ne vraca.
'
' Odluka stoji OVDE, a ne u ThisWorkbook-u: BeforeClose je poslednja brana bez
' obzira odakle je Close dosao, pa mesto odluke mora biti jedno i testabilno.
' Efekat (stvaran cloud append) trazi mrezu i meri se rucno; test meri ODLUKU,
' preko BulkPushNaIzlasku.
Public Sub ReleaseStanicaLockOnExit()
    If Len(gActiveStanica) = 0 Then Exit Sub
    Call ReleaseStanicaLockInternal(gActiveStanica, gActiveDatum, BulkPushNaIzlasku())
End Sub

' Sme li izlazak da nosi poslovni side effect (push u cloud)? Javno zbog testa:
' pogresan smer ove odluke ne pravi crven test nego duplikate kod klijenta.
Public Function BulkPushNaIzlasku() As Boolean
    BulkPushNaIzlasku = Not modImportState.ImportNijeDovrsen()
End Function

' Atomic stanica switch: release stari sa bulk push, acquire novi.
' Pozivaj iz cmbOtkupnoMesto_Change u frmOtkup kad korisnik menja
' stanicu unutar iste form sesije.
Public Function ChangeStanica(ByVal newStanica As String, _
                               ByVal newDatum As Date) As Boolean
    Const SRC As String = "ChangeStanica"
    
    On Error GoTo EH
    
    ' Acquire automatski release-uje prethodnu sa bulk push-om.
    ChangeStanica = AcquireStanicaLock(newStanica, newDatum)
    Exit Function

EH:
    LogErr SRC, "newStanica=" & newStanica
    ChangeStanica = False
End Function

Private Sub ReleaseStanicaLockInternal(ByVal stanicaID As String, _
                                        ByVal datum As Date, _
                                        ByVal doBulkPush As Boolean)
    Const SRC As String = "ReleaseStanicaLockInternal"
    
    On Error GoTo EH
    
    If Len(Trim$(stanicaID)) = 0 Then Exit Sub
    
    ' DODATO: desktop-only -- preskacemo i bulk push i sheet write.
    ' Resetujemo samo module state.
    If Not IsCloudSyncEnabled() Then
        If gActiveStanica = stanicaID Then
            gActiveStanica = ""
            gActiveDatum = 0
        End If
        Exit Sub
    End If
    
    ' Bulk push pre release-a -- jos uvek je locked, druge strane cekaju.
    If doBulkPush Then
        Call BulkPushPendingForStanica(stanicaID, datum)
    End If
    
    ' Brisanje lock-a (LOCK=NO + clear MESSAGE)
    Dim updates As Object
    Set updates = CreateObject("Scripting.Dictionary")
    
    Dim prefix As String: prefix = "STANICA_LOCK_" & stanicaID & "_"
    updates(prefix & "LOCK") = "NO"
    updates(prefix & "UPDATED_AT") = Format$(Now, "yyyy-mm-dd\Thh:nn:ss")
    updates(prefix & "MESSAGE") = ""
    updates(prefix & "OWNER") = ""
    
    Call ApplySyncControlUpdates(updates)
    
    ' Resetuj module state ako je ova stanica bila aktivna
    If gActiveStanica = stanicaID Then
        gActiveStanica = ""
        gActiveDatum = 0
        StopHeartbeatTimer
    End If
    
    LogInfo SRC, "Lock released: stanica=" & stanicaID
    Exit Sub

EH:
    LogErr SRC, "stanica=" & stanicaID
End Sub

' ============================================================
' PUBLIC -- Heartbeat (Application.OnTime callback)
' ============================================================

' Refresh UPDATED_AT za aktivnu stanicu. Pozvati iz Application.OnTime
' samo. Ako nema aktivne stanice, zaustavlja timer.
Public Sub HeartbeatStanicaLock()
    Const SRC As String = "HeartbeatStanicaLock"
    
    On Error GoTo EH
    
    gHeartbeatScheduled = False
    
    ' DODATO: desktop-only -- heartbeat nema smisla.
    If Not IsCloudSyncEnabled() Then Exit Sub
    
    If Len(gActiveStanica) = 0 Then Exit Sub
    
    If Len(gActiveStanica) = 0 Then
        ' Nema aktivnog lock-a, ne treba dalje
        Exit Sub
    End If
    
    Call UpdateStanicaLockTimestamp(gActiveStanica)
    
    ' Reschedule
    StartHeartbeatTimer
    Exit Sub

EH:
    LogErr SRC
    ' Pokusaj sledeci heartbeat ipak -- recoverable error
    StartHeartbeatTimer
End Sub

Public Sub StartHeartbeatTimer()
    If gHeartbeatScheduled Then Exit Sub
    If Len(gActiveStanica) = 0 Then Exit Sub
    
    ' DODATO: desktop-only -- ne sched-uj timer (heartbeat nema smisla)
    If Not IsCloudSyncEnabled() Then Exit Sub
    
    gNextHeartbeatTime = Now + TimeSerial(0, 0, HEARTBEAT_INTERVAL_SEC)
    
    On Error Resume Next
    Application.OnTime gNextHeartbeatTime, "modStanicaLock.HeartbeatStanicaLock"
    On Error GoTo 0
    
    gHeartbeatScheduled = True
End Sub

Public Sub StopHeartbeatTimer()
    If Not gHeartbeatScheduled Then Exit Sub
    
    On Error Resume Next
    Application.OnTime gNextHeartbeatTime, "modStanicaLock.HeartbeatStanicaLock", , False
    On Error GoTo 0
    
    gHeartbeatScheduled = False
End Sub

' Update samo UPDATED_AT za jednu stanicu (laksi update od full acquire).
Private Function UpdateStanicaLockTimestamp(ByVal stanicaID As String) As Boolean
    Dim updates As Object
    Set updates = CreateObject("Scripting.Dictionary")
    
    Dim prefix As String: prefix = "STANICA_LOCK_" & stanicaID & "_"
    updates(prefix & "UPDATED_AT") = Format$(Now, "yyyy-mm-dd\Thh:nn:ss")
    
    UpdateStanicaLockTimestamp = ApplySyncControlUpdates(updates)
End Function

' ============================================================
' PUBLIC -- diagnostic
' ============================================================

Public Function GetActiveStanica() As String
    GetActiveStanica = gActiveStanica
End Function

Public Function GetActiveDatum() As Date
    GetActiveDatum = gActiveDatum
End Function

' ============================================================
' PUBLIC -- Workbook_Open cleanup
' ============================================================

' Brisanje VBA-owned stanica lockova starijih od TTL. Pozvati u Workbook_Open
' jer ako je prethodna sesija crash-ovala, lock je ostao i blokira PWA.
' GAS TTL takode smatra lock stale-om, ali eksplicitno ciscenje je sigurnije.
Public Sub CleanupOrphanedLocks()
    Const SRC As String = "CleanupOrphanedLocks"
    
    On Error GoTo EH
    
    ' DODATO: desktop-only -- nema lockova za ciscenje
    If Not IsCloudSyncEnabled() Then Exit Sub
    
    Dim sheetID As String
    sheetID = GetSyncControlSpreadsheetID()
    If Len(sheetID) = 0 Then Exit Sub
    
    Dim kv As Object
    If Not TryReadSyncControlAsDict(sheetID, kv) Then
        LogWarn SRC, "SyncControl nije procitan -- ciscenje orphan lockova se preskace."
        Exit Sub
    End If

    Dim updates As Object
    Set updates = CreateObject("Scripting.Dictionary")
    
    Dim staleThreshold As Date
    staleThreshold = Now - TimeSerial(0, 10, 0)   ' 10 minuta = TTL
    
    Dim k As Variant
    For Each k In kv.keys
        Dim keyStr As String: keyStr = CStr(k)
        ' Nas zanima samo STANICA_LOCK_*_LOCK = YES
        If Left$(keyStr, 13) = "STANICA_LOCK_" And Right$(keyStr, 5) = "_LOCK" Then
            If UCase$(CStr(kv(keyStr))) = "YES" Then
                ' Proveri UPDATED_AT
                Dim updatedKey As String
                updatedKey = Left$(keyStr, Len(keyStr) - 5) & "_UPDATED_AT"
                
                Dim updatedAtStr As String
                updatedAtStr = CStr(kv(updatedKey))
                
                If Len(updatedAtStr) > 0 Then
                    On Error Resume Next
                    Dim ts As Date
                    ts = CDate(Replace(updatedAtStr, "T", " "))
                    On Error GoTo EH
                    
                    If ts < staleThreshold Then
                        ' Stale ? release
                        Dim ownerKey As String: ownerKey = Left$(keyStr, Len(keyStr) - 5) & "_OWNER"
                        Dim msgKey As String: msgKey = Left$(keyStr, Len(keyStr) - 5) & "_MESSAGE"
                        
                        updates(keyStr) = "NO"
                        updates(updatedKey) = Format$(Now, "yyyy-mm-dd\Thh:nn:ss")
                        updates(msgKey) = ""
                        updates(ownerKey) = ""
                        
                        LogInfo SRC, "Cleared stale lock: " & keyStr & " (last=" & updatedAtStr & ")"
                    End If
                End If
            End If
        End If
    Next k
    
    If updates.count > 0 Then
        Call ApplySyncControlUpdates(updates)
    End If
    Exit Sub

EH:
    LogErr SRC
End Sub

' ============================================================
' PUBLIC -- Bulk push pending za stanicu
' ============================================================

' Iterira tblOtkup za (StanicaID = X, Datum = Y, ClientRecordID is empty).
' Za svaki red (S1c: zaglavlje + stavke, REFAKTOR S14.8 t. 13):
'   1. Stavke otkupa -> tab OTK_STAVKE, red po stavci
'   2. Zaglavlje -> Sheet1, TEK kad su sve stavke upisane. Zaglavlje je
'      oznaka zavrsenog push-a: stavka bez zaglavlja je nedovrsen pokusaj.
'      Ponovljen pokusaj NE dupla stavku: OTK_STAVKE je jedinstven po
'      OtkupStavkaID (v. PosaljiStavkeOtkupa).
'   3. Na uspeh: upisi ClientRecordID = "VBA:" & OtkupID + SyncSource = "VBA"
'      (sledeci bulk push nece ga ponovo pokusati)
'   4. Na fail: ostavi ClientRecordID empty ? retry pri sledecem unlock-u
'
' Vraca broj uspesno push-ovanih redova.
Public Function BulkPushPendingForStanica(ByVal stanicaID As String, _
                                           ByVal datum As Date) As Long
    Const SRC As String = "BulkPushPendingForStanica"
    
    On Error GoTo EH
    
    ' DODATO: desktop-only -- bulk push nije potreban
    If Not IsCloudSyncEnabled() Then
        BulkPushPendingForStanica = 0
        Exit Function
    End If
    
    Dim spreadsheetID As String
    spreadsheetID = ResolveOTKSheetID(stanicaID)
    If Len(spreadsheetID) = 0 Then
        LogWarn SRC, "OTK-" & stanicaID & " sheet ne postoji ili offline. Skip."
        BulkPushPendingForStanica = 0
        Exit Function
    End If
    
    Dim lo As ListObject
    Set lo = GetTable(TBL_OTKUP)
    If lo Is Nothing Then
        BulkPushPendingForStanica = 0
        Exit Function
    End If
    If lo.DataBodyRange Is Nothing Then
        BulkPushPendingForStanica = 0
        Exit Function
    End If
    
    Dim iID As Long: iID = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    Dim iStanica As Long: iStanica = GetColumnIndex(TBL_OTKUP, COL_OTK_STANICA)
    Dim iDatum As Long: iDatum = GetColumnIndex(TBL_OTKUP, COL_OTK_DATUM)
    Dim iCRID As Long: iCRID = GetColumnIndex(TBL_OTKUP, "ClientRecordID")
    Dim iSource As Long: iSource = GetColumnIndex(TBL_OTKUP, "SyncSource")
    
    If iID = 0 Or iStanica = 0 Or iDatum = 0 Or iCRID = 0 Then
        LogError SRC, "Required tblOtkup kolone nedostaju."
        BulkPushPendingForStanica = 0
        Exit Function
    End If
    
    Dim datumStr As String: datumStr = Format$(datum, "yyyy-mm-dd")
    Dim pushed As Long: pushed = 0

    ' Stavke kroz kanonsku granicu, jednom za ceo prolaz. Pokvaren dokument
    ' obara ceo push po imenu (EH) -- zaglavlje bez stavki se ne salje.
    Dim stavkePoOtkupu As Object
    Set stavkePoOtkupu = modMasterSync.OtkStavkeRedoviPoOtkupu()
    Dim tabStavkiSpreman As Boolean
    Dim indeksStavki As Object
    
    Dim r As Long
    For r = 1 To lo.DataBodyRange.rows.count
        Dim rowStanica As String
        rowStanica = CStr(lo.DataBodyRange.cells(r, iStanica).value)
        If rowStanica <> stanicaID Then GoTo NextRow
        
        Dim rowDatumStr As String
        rowDatumStr = ""
        On Error Resume Next
        rowDatumStr = Format$(CDate(lo.DataBodyRange.cells(r, iDatum).value), "yyyy-mm-dd")
        On Error GoTo EH
        If rowDatumStr <> datumStr Then GoTo NextRow
        
        Dim crid As String
        crid = CStr(nz(lo.DataBodyRange.cells(r, iCRID).value, ""))
        If Len(crid) > 0 Then GoTo NextRow   ' Vec push-ovan
        
        ' Push ovaj red
        Dim otkupID As String
        otkupID = CStr(lo.DataBodyRange.cells(r, iID).value)
        
        Dim rowData As Variant
        rowData = BuildOTKSheetRowForOtkup(otkupID, stanicaID, lo, r, iID)
        If IsEmpty(rowData) Then GoTo NextRow

        If Not stavkePoOtkupu.Exists(otkupID) Then
            LogError SRC, "Otkup bez stavki se ne salje: OtkupID=" & otkupID
            GoTo NextRow
        End If
        If Not tabStavkiSpreman Then
            If Not PripremiOtkStavkeTab(spreadsheetID, indeksStavki) Then
                LogWarn SRC, "Tab " & OTK_STAVKE_TAB & " nije spreman; push odlozen."
                Exit For
            End If
            tabStavkiSpreman = True
        End If
        Dim greskaStavki As String
        If Not PosaljiStavkeOtkupa(spreadsheetID, indeksStavki, stavkePoOtkupu(otkupID), greskaStavki) Then
            LogWarn SRC, "Push stavki nije zavrsen za OtkupID=" & otkupID & _
                         " (" & greskaStavki & "), zaglavlje se ne salje"
            GoTo NextRow
        End If

        If AppendRowToSheet(spreadsheetID, "Sheet1", rowData) Then
            ' Mark as pushed
            lo.DataBodyRange.cells(r, iCRID).value = "VBA:" & otkupID
            If iSource > 0 Then
                lo.DataBodyRange.cells(r, iSource).value = "VBA"
            End If
            pushed = pushed + 1
        Else
            LogWarn SRC, "Push failed za OtkupID=" & otkupID & _
                         ", ostavlja ClientRecordID prazan za retry"
        End If
        
NextRow:
    Next r
    
    LogInfo SRC, "Bulk push complete: stanica=" & stanicaID & _
                 " datum=" & datumStr & " pushed=" & pushed
    BulkPushPendingForStanica = pushed
    Exit Function

EH:
    LogErr SRC, "stanica=" & stanicaID
    BulkPushPendingForStanica = 0
End Function

' Red zaglavlja otkupa za Sheet1 OTK-* sheet-a, po imenu iz
' modMasterSync.OtkZaglavljeKolone (jedino mesto rasporeda).
'
' Klasa, Kolicina, Cena i KolAmbalaze su PRAZNI: to su polja stavke i idu u
' OTK_STAVKE. Kolona koju ovaj graditelj ne poznaje pada -- nova kolona u
' spisku ne sme tiho da ode prazna.
Public Function BuildOTKSheetRowForOtkup(ByVal otkupID As String, _
                                           ByVal stanicaID As String, _
                                           ByVal lo As ListObject, _
                                           ByVal rowIdx As Long, _
                                           ByVal iID As Long) As Variant
    On Error GoTo EH

    Dim kooperantID As String
    kooperantID = CStr(OtkCelija(lo, rowIdx, COL_OTK_KOOPERANT))

    Dim kooperantName As String
    kooperantName = CStr(nz(LookupValue(TBL_KOOPERANTI, "KooperantID", kooperantID, "Ime"), ""))
    Dim koopPrezime As String
    koopPrezime = CStr(nz(LookupValue(TBL_KOOPERANTI, "KooperantID", kooperantID, "Prezime"), ""))
    If Len(koopPrezime) > 0 Then
        kooperantName = Trim$(kooperantName & " " & koopPrezime)
    End If

    Dim nowIso As String: nowIso = Format$(Now, "yyyy-mm-dd\Thh:nn:ss")

    Dim kol As Variant
    kol = modMasterSync.OtkZaglavljeKolone()
    Dim rowOut() As Variant
    ReDim rowOut(0 To UBound(kol) - LBound(kol))

    Dim k As Long, v As Variant
    For k = LBound(kol) To UBound(kol)
        Select Case CStr(kol(k))
            Case "ClientRecordID": v = "VBA:" & otkupID
            Case "ServerRecordID": v = otkupID
            Case "CreatedAtClient", "UpdatedAtClient", "ReceivedAt": v = nowIso
            Case "UpdatedAtServer", "Napomena": v = ""
            Case "SyncStatus": v = SYNC_STATUS_MASTER_GOOGLE
            Case "DeviceID": v = DEVICE_ID_VBA
            Case "OtkupacID": v = stanicaID
            Case "Datum": v = Format$(CDate(OtkCelija(lo, rowIdx, COL_OTK_DATUM)), "yyyy-mm-dd")
            Case "KooperantID": v = kooperantID
            Case "KooperantName": v = kooperantName
            Case "VrstaVoca": v = CStr(OtkCelija(lo, rowIdx, COL_OTK_VRSTA))
            Case "SortaVoca": v = CStr(nz(OtkCelija(lo, rowIdx, COL_OTK_SORTA), ""))
            Case "TipAmbalaze": v = CStr(nz(OtkCelija(lo, rowIdx, COL_OTK_TIP_AMB), ""))
            Case "ParcelaID": v = CStr(nz(OtkCelija(lo, rowIdx, COL_OTK_PARCELA), ""))
            Case "VozacID": v = CStr(nz(OtkCelija(lo, rowIdx, COL_OTK_VOZAC), ""))
            Case "BrojDokumenta": v = CStr(nz(OtkCelija(lo, rowIdx, COL_OTK_BR_DOK), ""))
            Case "Klasa", "Kolicina", "Cena", "KolAmbalaze": v = ""   ' stavka -> OTK_STAVKE
            Case Else
                Err.Raise vbObjectError + 8142, "BuildOTKSheetRowForOtkup", _
                          "Kolona OTK zaglavlja bez izvora: " & CStr(kol(k))
        End Select
        rowOut(k - LBound(kol)) = v
    Next k

    BuildOTKSheetRowForOtkup = rowOut
    Exit Function

EH:
    LogErr "BuildOTKSheetRowForOtkup", "otkupID=" & otkupID
    BuildOTKSheetRowForOtkup = Empty
End Function

' Celija reda tblOtkup po imenu kolone; kolona koja fali pada po imenu.
Private Function OtkCelija(ByVal lo As ListObject, ByVal rowIdx As Long, _
                           ByVal kolona As String) As Variant
    OtkCelija = lo.DataBodyRange.cells(rowIdx, _
        RequireColumnIndex(TBL_OTKUP, kolona, "BuildOTKSheetRowForOtkup")).value
End Function

' ============================================================
' OTK_STAVKE -- UGOVOR: JEDAN RED PO OtkupStavkaID (review #357, P1)
'
' Pisac je IDEMPOTENTAN. Pre slanja se tab procita jednom; stavka ciji ID vec
' postoji sa ISTIM sadrzajem se ne salje ponovo, a isti ID sa DRUGACIJIM
' sadrzajem je konflikt: otkup se ne salje (ni stavke ni zaglavlje) i ostaje
' za operatera. Zato ponovljen push posle mreznog pada ne moze da promeni
' kolicinu: tab nikad ne dobije drugi red istog ID-a iz ovog pisca.
'
' Citalac (S5) sme da tretira dupli OtkupStavkaID kao kvar, ne kao zbir.
' Indeks vazi za jedan prolaz: push radi pod lock-om stanice, a PWA do S5 ne
' pise u OTK_STAVKE.
' ============================================================

' Tab postoji, naslov je TACNO modMasterSync.OtkStavkeKolone (istim redom), i
' outIndeks nosi postojece stavke. Neuspelo citanje nije "prazan tab", a tudji
' naslov nije "dovoljno blizu": False, push se odlaze (fail-closed, review #357 P2).
Private Function PripremiOtkStavkeTab(ByVal spreadsheetID As String, _
                                      ByRef outIndeks As Object) As Boolean
    Const SRC As String = "PripremiOtkStavkeTab"
    On Error GoTo EH

    Set outIndeks = Nothing
    If Not AddSheetTab(spreadsheetID, OTK_STAVKE_TAB, True) Then Exit Function

    Dim postojece As Variant
    If Not TryReadSheetData(spreadsheetID, OTK_STAVKE_TAB, postojece) Then Exit Function
    If IsEmpty(postojece) Then
        If Not AppendRowToSheet(spreadsheetID, OTK_STAVKE_TAB, _
                                modMasterSync.OtkStavkeKolone()) Then Exit Function
    End If

    Set outIndeks = OtkStavkeIndeksIzTaba(postojece)
    PripremiOtkStavkeTab = True
    Exit Function

EH:
    LogErr SRC
    PripremiOtkStavkeTab = False
End Function

' Sadrzaj OTK_STAVKE (2D, red 1 = naslov; Empty = prazan tab) -> indeks
' OtkupStavkaID -> kljuc sadrzaja. Pada po imenu na: naslov koji nije tacno
' OtkStavkeKolone, red bez OtkupStavkaID, isti ID sa razlicitim sadrzajem.
Public Function OtkStavkeIndeksIzTaba(ByVal data As Variant) As Object
    Const SRC As String = "OtkStavkeIndeksIzTaba"

    Dim indeks As Object
    Set indeks = CreateObject("Scripting.Dictionary")
    Set OtkStavkeIndeksIzTaba = indeks
    If IsEmpty(data) Then Exit Function

    Dim kol As Variant, nk As Long, k As Long
    kol = modMasterSync.OtkStavkeKolone()
    nk = UBound(kol) - LBound(kol) + 1

    Dim lb2 As Long, ub2 As Long
    lb2 = LBound(data, 2)
    ub2 = UBound(data, 2)
    If ub2 - lb2 + 1 < nk Then
        Err.Raise vbObjectError + 8144, SRC, _
                  "Naslov taba " & OTK_STAVKE_TAB & " ima manje kolona od ugovora."
    End If
    For k = 0 To ub2 - lb2
        If k < nk Then
            If CStr(data(LBound(data, 1), lb2 + k)) <> CStr(kol(LBound(kol) + k)) Then
                Err.Raise vbObjectError + 8144, SRC, _
                          "Naslov taba " & OTK_STAVKE_TAB & " kolona " & (k + 1) & " je '" & _
                          CStr(data(LBound(data, 1), lb2 + k)) & "', ugovor trazi '" & _
                          CStr(kol(LBound(kol) + k)) & "'."
            End If
        ElseIf Len(Trim$(CStr(data(LBound(data, 1), lb2 + k)))) > 0 Then
            Err.Raise vbObjectError + 8144, SRC, _
                      "Naslov taba " & OTK_STAVKE_TAB & " ima kolonu van ugovora: " & _
                      CStr(data(LBound(data, 1), lb2 + k))
        End If
    Next k

    Dim r As Long, red() As Variant, id As String, kljuc As String
    ReDim red(0 To nk - 1)
    For r = LBound(data, 1) + 1 To UBound(data, 1)
        For k = 0 To nk - 1
            red(k) = data(r, lb2 + k)
        Next k
        id = OtkStavkaIdReda(red)
        If Len(id) = 0 Then
            Err.Raise vbObjectError + 8145, SRC, _
                      "Red " & r & " taba " & OTK_STAVKE_TAB & " nema OtkupStavkaID."
        End If
        kljuc = OtkStavkaKljuc(red)
        If indeks.Exists(id) Then
            If indeks(id) <> kljuc Then
                Err.Raise vbObjectError + 8146, SRC, _
                          "Konflikt u " & OTK_STAVKE_TAB & ": OtkupStavkaID " & id & _
                          " ima dva razlicita sadrzaja."
            End If
        Else
            indeks.Add id, kljuc
        End If
    Next r
End Function

' Salje stavke jednog otkupa idempotentno po OtkupStavkaID i azurira indeks.
' True = sve stavke su u tabu (sada ili od ranije). False + outGreska: pad
' appenda (ostatak se salje pri sledecem pokusaju) ili konflikt (nista se ne
' salje -- konflikt se proverava za SVE stavke pre prvog upisa).
Public Function PosaljiStavkeOtkupa(ByVal spreadsheetID As String, ByVal indeks As Object, _
                                    ByVal stavke As Collection, ByRef outGreska As String) As Boolean
    outGreska = ""
    If indeks Is Nothing Then
        outGreska = "indeks " & OTK_STAVKE_TAB & " nije procitan"
        Exit Function
    End If

    Dim red As Variant, id As String
    For Each red In stavke
        id = OtkStavkaIdReda(red)
        If Len(id) = 0 Then
            outGreska = "stavka bez OtkupStavkaID"
            Exit Function
        End If
        If indeks.Exists(id) Then
            If indeks(id) <> OtkStavkaKljuc(red) Then
                outGreska = "konflikt: OtkupStavkaID " & id & " u " & OTK_STAVKE_TAB & _
                            " ima drugaciji sadrzaj"
                Exit Function
            End If
        End If
    Next red

    For Each red In stavke
        id = OtkStavkaIdReda(red)
        If Not indeks.Exists(id) Then
            If Not OtkStavkaAppend(spreadsheetID, red) Then
                outGreska = "append nije uspeo za OtkupStavkaID " & id
                Exit Function
            End If
            indeks.Add id, OtkStavkaKljuc(red)
        End If
    Next red

    PosaljiStavkeOtkupa = True
End Function

Private Function OtkStavkaIdReda(ByVal red As Variant) As String
    Dim kol As Variant, k As Long
    kol = modMasterSync.OtkStavkeKolone()
    For k = LBound(kol) To UBound(kol)
        If CStr(kol(k)) = COL_OKS_ID Then
            OtkStavkaIdReda = Trim$(CStr(red(LBound(red) + k - LBound(kol))))
            Exit Function
        End If
    Next k
    Err.Raise vbObjectError + 8147, "OtkStavkaIdReda", "OtkStavkeKolone nema " & COL_OKS_ID
End Function

' Kljuc sadrzaja stavke. Broj se normalizuje (Sheets vraca 100 ili "100"),
' tekst se trimuje; prazno ostaje prazno. Normalizacija koja pogresi daje
' LAZAN konflikt (push staje), nikad tihi duplikat.
Private Function OtkStavkaKljuc(ByVal red As Variant) As String
    Dim k As Long, v As Variant, s As String
    For k = LBound(red) To UBound(red)
        v = red(k)
        If IsNumeric(v) And Len(Trim$(CStr(v))) > 0 Then
            s = s & "|" & CStr(CDbl(v))
        Else
            s = s & "|" & Trim$(CStr(v))
        End If
    Next k
    OtkStavkaKljuc = s
End Function

Private Function OtkStavkaAppend(ByVal spreadsheetID As String, ByVal red As Variant) As Boolean
    If mSimRedovi Is Nothing Then
        OtkStavkaAppend = AppendRowToSheet(spreadsheetID, OTK_STAVKE_TAB, red)
        Exit Function
    End If
    mSimPoziva = mSimPoziva + 1
    If mSimPoziva = mSimPadNa Then Exit Function
    mSimRedovi.Add red
    OtkStavkaAppend = True
End Function

' Test: append u OTK_STAVKE ide u redovi; poziv broj padNa (1-based, od
' postavljanja) vraca neuspeh. Nothing = razoruzaj.
Public Sub TestHook_OtkStavkeSimulacija(ByVal redovi As Collection, ByVal padNa As Long)
    Set mSimRedovi = redovi
    mSimPadNa = padNa
    mSimPoziva = 0
End Sub

' ============================================================
' PRIVATE -- SyncControl read/write helperi
' ============================================================

' Cita ceo SyncControl tab kao Scripting.Dictionary key -> value.
'
' AUD-001: FAIL-CLOSED. Vraca False na svaku gresku citanja/parsiranja.
' Prazan dictionary na neuspelo citanje je ranije vodio u
' ApplySyncControlUpdates -> WriteSheetData sa NEPOTPUNIM skupom, sto brise
' lockove drugih stanica i ostale SyncControl parametre. Prolazna mrezna ili
' JSON greska ne sme da prepise ceo tab.
'
'   True  + outDict -> tab procitan (moze biti i prazan)
'   False           -> citanje/parsiranje nije uspelo, NE upisivati nista
Private Function TryReadSyncControlAsDict(ByVal sheetID As String, _
                                          ByRef outDict As Object) As Boolean
    Dim result As Object
    Dim data As Variant
    Dim r As Long
    Dim k As String, v As String

    Set outDict = Nothing
    Set result = CreateObject("Scripting.Dictionary")

    If Not TryReadSheetData(sheetID, SYNC_CONTROL_TAB, data) Then
        LogError "TryReadSyncControlAsDict", _
                 "SyncControl citanje nije uspelo (HTTP/JSON). Upis se preskace da se tab ne prepise nepotpuno."
        Exit Function
    End If

    If IsEmpty(data) Then
        ' Validan, ali prazan tab.
        Set outDict = result
        TryReadSyncControlAsDict = True
        Exit Function
    End If

    ' Header je red 1, podaci od reda 2
    For r = LBound(data, 1) + 1 To UBound(data, 1)
        k = Trim$(CStr(nz(data(r, 1), "")))
        v = Trim$(CStr(nz(data(r, 2), "")))
        If Len(k) > 0 Then
            result(k) = v
        End If
    Next r

    Set outDict = result
    TryReadSyncControlAsDict = True
End Function

' Read-modify-write SyncControl tab sa specific updates. Cuva sve ostale kljuceve.
Private Function ApplySyncControlUpdates(ByVal updates As Object) As Boolean
    Const SRC As String = "ApplySyncControlUpdates"
    
    On Error GoTo EH
    
    Dim sheetID As String
    sheetID = GetSyncControlSpreadsheetID()
    If Len(sheetID) = 0 Then
        ApplySyncControlUpdates = False
        Exit Function
    End If
    
    ' Read existing.
    ' AUD-001: read-modify-write -- ako citanje padne, NE SME se upisivati,
    ' inace bi WriteSheetData prepisao ceo tab sa samo trenutnim kljucevima.
    Dim existing As Object
    If Not TryReadSyncControlAsDict(sheetID, existing) Then
        LogError SRC, "SyncControl nije procitan -- update je prekinut bez upisa."
        ApplySyncControlUpdates = False
        Exit Function
    End If

    ' Merge updates
    Dim k As Variant
    For Each k In updates.keys
        existing(CStr(k)) = updates(k)
    Next k
    
    ' Build 2D array (header + sorted keys za stabilnost)
    Dim rowCount As Long: rowCount = existing.count + 1
    Dim arr() As Variant
    ReDim arr(1 To rowCount, 1 To 2)
    arr(1, 1) = "Parameter"
    arr(1, 2) = "Vrednost"
    
    Dim r As Long: r = 2
    For Each k In existing.keys
        arr(r, 1) = CStr(k)
        arr(r, 2) = CStr(existing(k))
        r = r + 1
    Next k
    
    ApplySyncControlUpdates = WriteSheetData(sheetID, SYNC_CONTROL_TAB, arr)
    Exit Function

EH:
    LogErr SRC
    ApplySyncControlUpdates = False
End Function

' Resolve Stammdaten spreadsheet ID za SyncControl tab.
' Reuse postojeci pattern iz modGoogleSyncOrchestrator (Private),
' ali kako je tamo Private, ponavljamo logiku ovde.
Private Function GetSyncControlSpreadsheetID() As String
    Dim sheetID As String
    
    sheetID = GetConfigValue("GOOGLE_STAMMDATEN_SHEET_ID")
    If Len(sheetID) > 0 Then
        GetSyncControlSpreadsheetID = sheetID
        Exit Function
    End If
    
    ' Fallback: lookup po imenu
    Dim folderID As String
    folderID = GetConfigValue("GOOGLE_PWA_FOLDER_ID")
    If Len(folderID) = 0 Then
        GetSyncControlSpreadsheetID = ""
        Exit Function
    End If
    
    GetSyncControlSpreadsheetID = GetSpreadsheetID("Stammdaten", folderID)
End Function

' Resolve OTK-{stanicaID} sheet ID.
Private Function ResolveOTKSheetID(ByVal stanicaID As String) As String
    Dim folderID As String
    folderID = GetConfigValue("GOOGLE_PWA_FOLDER_ID")
    If Len(folderID) = 0 Then
        ResolveOTKSheetID = ""
        Exit Function
    End If
    
    ResolveOTKSheetID = GetSpreadsheetID("OTK-" & stanicaID, folderID)
End Function

' ============================================================
' PRIVATE -- local Nz (postojeci je Private u modMasterSync)
' ============================================================
Private Function nz(ByVal v As Variant, _
                     Optional ByVal Fallback As Variant = "") As Variant
    If IsNull(v) Then
        nz = Fallback
    ElseIf IsEmpty(v) Then
        nz = Fallback
    ElseIf VarType(v) = vbString And Len(v) = 0 Then
        nz = Fallback
    Else
        nz = v
    End If
End Function

