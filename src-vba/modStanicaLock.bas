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
    Set kv = ReadSyncControlAsDict(sheetID)
    
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
' Za svaki red:
'   1. Sastavi 23-col GAS row (preuzimajuci podatke iz tblOtkup + lookup-ovi)
'   2. AppendRowToSheet ? OTK-{stanicaID}
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

' Konstruise 1D Array (23 elementa) za OTK sheet append iz tblOtkup reda.
Private Function BuildOTKSheetRowForOtkup(ByVal otkupID As String, _
                                           ByVal stanicaID As String, _
                                           ByVal lo As ListObject, _
                                           ByVal rowIdx As Long, _
                                           ByVal iID As Long) As Variant
    On Error GoTo EH
    
    Dim iDatum As Long: iDatum = GetColumnIndex(TBL_OTKUP, COL_OTK_DATUM)
    Dim iKoop As Long: iKoop = GetColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT)
    Dim iVrsta As Long: iVrsta = GetColumnIndex(TBL_OTKUP, COL_OTK_VRSTA)
    Dim iSorta As Long: iSorta = GetColumnIndex(TBL_OTKUP, COL_OTK_SORTA)
    Dim iKol As Long: iKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
    Dim iCena As Long: iCena = GetColumnIndex(TBL_OTKUP, COL_OTK_CENA)
    Dim iTipAmb As Long: iTipAmb = GetColumnIndex(TBL_OTKUP, COL_OTK_TIP_AMB)
    Dim iKolAmb As Long: iKolAmb = GetColumnIndex(TBL_OTKUP, COL_OTK_KOL_AMB)
    Dim iVozac As Long: iVozac = GetColumnIndex(TBL_OTKUP, COL_OTK_VOZAC)
    Dim iBrDok As Long: iBrDok = GetColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK)
    Dim iKlasa As Long: iKlasa = GetColumnIndex(TBL_OTKUP, COL_OTK_KLASA)
    Dim iParcela As Long: iParcela = GetColumnIndex(TBL_OTKUP, COL_OTK_PARCELA)
    
    Dim kooperantID As String
    kooperantID = CStr(lo.DataBodyRange.cells(rowIdx, iKoop).value)
    
    Dim kooperantName As String
    kooperantName = CStr(nz(LookupValue(TBL_KOOPERANTI, "KooperantID", kooperantID, "Ime"), ""))
    Dim koopPrezime As String
    koopPrezime = CStr(nz(LookupValue(TBL_KOOPERANTI, "KooperantID", kooperantID, "Prezime"), ""))
    If Len(koopPrezime) > 0 Then
        kooperantName = Trim$(kooperantName & " " & koopPrezime)
    End If
    
    Dim nowIso As String: nowIso = Format$(Now, "yyyy-mm-dd\Thh:nn:ss")
    Dim datumIso As String: datumIso = Format$(CDate(lo.DataBodyRange.cells(rowIdx, iDatum).value), "yyyy-mm-dd")
    
    Dim rowOut(0 To 22) As Variant
    rowOut(0) = "VBA:" & otkupID                                              ' 1  ClientRecordID
    rowOut(1) = otkupID                                                       ' 2  ServerRecordID
    rowOut(2) = nowIso                                                        ' 3  CreatedAtClient
    rowOut(3) = nowIso                                                        ' 4  UpdatedAtClient
    rowOut(4) = ""                                                            ' 5  UpdatedAtServer
    rowOut(5) = SYNC_STATUS_MASTER_GOOGLE                                     ' 6  SyncStatus
    rowOut(6) = DEVICE_ID_VBA                                                 ' 7  DeviceID
    rowOut(7) = stanicaID                                                     ' 8  OtkupacID
    rowOut(8) = datumIso                                                      ' 9  Datum
    rowOut(9) = kooperantID                                                   ' 10 KooperantID
    rowOut(10) = kooperantName                                                ' 11 KooperantName
    rowOut(11) = CStr(lo.DataBodyRange.cells(rowIdx, iVrsta).value)           ' 12 VrstaVoca
    rowOut(12) = CStr(nz(lo.DataBodyRange.cells(rowIdx, iSorta).value, ""))   ' 13 SortaVoca
    rowOut(13) = CStr(lo.DataBodyRange.cells(rowIdx, iKlasa).value)           ' 14 Klasa
    rowOut(14) = CDbl(lo.DataBodyRange.cells(rowIdx, iKol).value)             ' 15 Kolicina
    rowOut(15) = CDbl(lo.DataBodyRange.cells(rowIdx, iCena).value)            ' 16 Cena
    rowOut(16) = CStr(nz(lo.DataBodyRange.cells(rowIdx, iTipAmb).value, ""))  ' 17 TipAmbalaze
    rowOut(17) = CLng(nz(lo.DataBodyRange.cells(rowIdx, iKolAmb).value, 0))   ' 18 KolAmbalaze
    rowOut(18) = CStr(nz(lo.DataBodyRange.cells(rowIdx, iParcela).value, "")) ' 19 ParcelaID
    rowOut(19) = CStr(nz(lo.DataBodyRange.cells(rowIdx, iVozac).value, ""))   ' 20 VozacID
    rowOut(20) = ""                                                            ' 21 Napomena
    rowOut(21) = nowIso                                                        ' 22 ReceivedAt
    rowOut(22) = CStr(nz(lo.DataBodyRange.cells(rowIdx, iBrDok).value, ""))   ' 23 BrojDokumenta
    
    BuildOTKSheetRowForOtkup = rowOut
    Exit Function

EH:
    LogErr "BuildOTKSheetRowForOtkup", "otkupID=" & otkupID
    BuildOTKSheetRowForOtkup = Empty
End Function

' ============================================================
' PRIVATE -- SyncControl read/write helperi
' ============================================================

' Cita ceo SyncControl tab kao Scripting.Dictionary key -> value.
Private Function ReadSyncControlAsDict(ByVal sheetID As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    Dim data As Variant
    data = ReadSheetData(sheetID, SYNC_CONTROL_TAB)
    If IsEmpty(data) Then
        Set ReadSyncControlAsDict = result
        Exit Function
    End If
    
    ' Header je red 1, podaci od reda 2
    Dim r As Long
    For r = LBound(data, 1) + 1 To UBound(data, 1)
        Dim k As String, v As String
        k = Trim$(CStr(nz(data(r, 1), "")))
        v = Trim$(CStr(nz(data(r, 2), "")))
        If Len(k) > 0 Then
            result(k) = v
        End If
    Next r
    
    Set ReadSyncControlAsDict = result
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
    
    ' Read existing
    Dim existing As Object
    Set existing = ReadSyncControlAsDict(sheetID)
    
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

