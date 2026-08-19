Attribute VB_Name = "modStornoZurnal"
Option Explicit

' ============================================================
' modStornoZurnal - append-only cell-level zurnal storno operacija za LOSSLESS
' "Vrati storno" (pravi inverz storna).
'
' Ambient op-kontekst: storno primitiva (StornoOtkup / StornoOMKoopByBrDok) pozove
' BeginStornoOp na ulazu; mutacione tacke JournalCell-uju (StaraVrednost, NovaVrednost)
' PRE mutacije; EndStornoOp zatvara. UndoOperation_TX vrati svaku celiju na
' StaraVrednost i cilja SAMO tu operaciju.
'
' Opseg (faza 1): Otkup + Revers (OM-koop). Chain kasnije.
'
' SIGURNOST (posle review-a):
'  - BeginStornoOp je FAIL-CLOSED: bez OperationID -> raise (storno se prekida), nema
'    "SOP-1" fallback; nested poziv drugog dokumenta -> raise (mesanje operacija).
'  - JournalCell je FAIL-CLOSED: neuspeo upis -> raise -> storno rollback.
'  - UndoOperation_TX pre mutacije: konzistentnost op-a (isti DocType+Broj), podrzana
'    tabela + kolona, TACNO JEDAN ciljni red (dup-PK -> odbij), i DRIFT guard
'    (trenutna vrednost == NovaVrednost; inace stanje promenjeno posle storna -> odbij).
'  - Otkup dup-guard je PO (broj, klasa) reda (ne broj-level) -> parcijalni storno
'    jedne klase se moze vratiti iako je druga klasa aktivna.
' ============================================================

Private Const MOD_NAME As String = "modStornoZurnal"
Private Const ERR_SZ_BASE As Long = vbObjectError + 3100

Private mOpID As String
Private mActive As Boolean
Private mDocType As String
Private mBroj As String
Private mNextZur As Long          ' kesiran sledeci ZurnalID broj (perf: bez GetNextID po celiji)

' ============================================================
' AMBIENT OP KONTEKST
' ============================================================
' Vraca True ako je OVAJ poziv otvorio operaciju (pa je on i zatvara). Nested poziv
' ISTOG (docType, broj) se pridruzuje (vraca False); nested poziv DRUGOG dokumenta je
' greska (mesanje operacija). FAIL-CLOSED: bez OperationID dize gresku.
Public Function BeginStornoOp(ByVal docType As String, ByVal broj As String) As Boolean
    Const SRC As String = MOD_NAME & ".BeginStornoOp"
    If Len(Trim$(docType)) = 0 Then Err.Raise ERR_SZ_BASE + 20, SRC, "DocType je obavezan za storno operaciju."
    If mActive Then
        If StrComp(docType, mDocType, vbTextCompare) <> 0 Or StrComp(broj, mBroj, vbTextCompare) <> 0 Then
            Err.Raise ERR_SZ_BASE + 21, SRC, "Pokusaj mesanja storno operacija (" & _
                mDocType & " " & mBroj & " vs " & docType & " " & broj & ")."
        End If
        Exit Function                       ' pridruzi se aktivnoj (owns=False)
    End If
    mOpID = GetNextID(TBL_STORNO_ZURNAL, COL_SZ_OP_ID, "SOP-")
    If Len(mOpID) = 0 Then Err.Raise ERR_SZ_BASE + 22, SRC, "OperationID nije generisan (zurnal sema?)."
    Dim zbase As String: zbase = GetNextID(TBL_STORNO_ZURNAL, COL_SZ_ID, "ZUR-")
    If Len(zbase) = 0 Then Err.Raise ERR_SZ_BASE + 23, SRC, "ZurnalID nije generisan (zurnal sema?)."
    mNextZur = OpNum4(zbase, "ZUR-")
    mDocType = docType: mBroj = broj
    mActive = True
    BeginStornoOp = True
End Function

Public Sub EndStornoOp(ByVal owns As Boolean)
    If Not owns Then Exit Sub
    AbortStornoOp
End Sub

' Force-reset op-konteksta (za EH putanje entry _TX-ova) -> nikad ne ostavi op otvoren.
Public Sub AbortStornoOp()
    mActive = False: mOpID = "": mDocType = "": mBroj = "": mNextZur = 0
End Sub

Public Function StornoOpActive() As Boolean
    StornoOpActive = mActive
End Function

' JournalCell: (tabela, RowID(PK), kolona, STARA, NOVA) za tekucu op. No-op ako op
' nije aktivan. FAIL-CLOSED: neuspeo upis dize gresku -> storno primitiva reraise ->
' _TX rollback (lossless zavisi od KOMPLETNOG zurnala). Perf: ZurnalID iz kesa.
Public Sub JournalCell(ByVal tbl As String, ByVal rowID As String, ByVal col As String, _
                       ByVal oldVal As Variant, ByVal newVal As Variant)
    Const SRC As String = MOD_NAME & ".JournalCell"
    If Not mActive Then Exit Sub
    Dim zid As String: zid = "ZUR-" & CStr(mNextZur)
    ' Redosled MORA pratiti EnsureStornoZurnalSchemaCore:
    ' ZurnalID, OperationID, Timestamp, DocType, Broj, Tabela, RowID, Kolona, StaraVrednost, NovaVrednost
    If AppendRow(TBL_STORNO_ZURNAL, Array(zid, mOpID, Format$(Now, "yyyy-mm-dd hh:nn:ss"), _
        mDocType, mBroj, tbl, CStr(rowID), col, CStr(oldVal), CStr(newVal))) = 0 Then
        Err.Raise ERR_SZ_BASE + 10, SRC, "Zurnal upis nije uspeo (" & tbl & "." & col & _
            ") -> storno se prekida (lossless garancija)."
    End If
    mNextZur = mNextZur + 1
End Sub

' ============================================================
' UNDO - pravi inverz jedne operacije.
' ============================================================
Public Function UndoOperation_TX(ByVal opID As String) As Boolean
    Const SRC As String = MOD_NAME & ".UndoOperation_TX"
    Dim tx As clsTransaction
    On Error GoTo EH
    opID = Trim$(opID)
    If Len(opID) = 0 Then Err.Raise ERR_SZ_BASE + 1, SRC, "OperationID je obavezan."

    Dim data As Variant: data = GetTableData(TBL_STORNO_ZURNAL)
    If IsEmpty(data) Then Err.Raise ERR_SZ_BASE + 2, SRC, "Storno zurnal je prazan."

    Dim cOp As Long, cTbl As Long, cRow As Long, cCol As Long, cOld As Long, cNew As Long, cDoc As Long, cBroj As Long
    cOp = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_OP_ID)
    cTbl = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_TABELA)
    cRow = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_ROWID)
    cCol = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_KOLONA)
    cOld = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_STARA)
    cNew = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_NOVA)
    cDoc = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_DOCTYPE)
    cBroj = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_BROJ)
    If cOp = 0 Or cNew = 0 Then Err.Raise ERR_SZ_BASE + 6, SRC, "Zurnal sema nije kompletna (NovaVrednost?)."

    Dim rows As Collection: Set rows = New Collection
    Dim docType As String, broj As String, gotHdr As Boolean
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cOp))) = opID Then
            ' Konzistentnost operacije: svi redovi istog opID moraju deliti DocType+Broj.
            If Not gotHdr Then
                docType = CStr(data(i, cDoc)): broj = CStr(data(i, cBroj)): gotHdr = True
            ElseIf StrComp(CStr(data(i, cDoc)), docType, vbTextCompare) <> 0 _
                 Or StrComp(CStr(data(i, cBroj)), broj, vbTextCompare) <> 0 Then
                Err.Raise ERR_SZ_BASE + 11, SRC, "Nekonzistentna operacija " & opID & _
                    " (pomesani DocType/Broj) -> odbijeno."
            End If
            rows.Add Array(CStr(data(i, cTbl)), CStr(data(i, cRow)), CStr(data(i, cCol)), _
                           CStr(data(i, cOld)), CStr(data(i, cNew)))
        End If
    Next i
    If rows.count = 0 Then Err.Raise ERR_SZ_BASE + 3, SRC, "Operacija nije nadjena: " & opID

    ' Garde (broj-level): dead-parent (otkup) + active-dup ambalaze (revers). Otkup
    ' active-dup se proverava PO REDU nize (parcijalna klasa).
    Dim gr As String: gr = UndoGuardReason(docType, broj)
    If Len(gr) > 0 Then Err.Raise ERR_SZ_BASE + 4, SRC, gr

    ' Pre-validacija (SVE-ILI-NISTA) PRE ijedne mutacije:
    For i = 1 To rows.count
        Dim vt As String: vt = CStr(rows(i)(0))
        Dim vpk As String: vpk = PkColForTable(vt)
        If Len(vpk) = 0 Then Err.Raise ERR_SZ_BASE + 7, SRC, "Nepodrzana tabela u zurnalu: " & vt
        Dim vcol As String: vcol = CStr(rows(i)(2))
        If GetColumnIndex(vt, vcol) = 0 Then Err.Raise ERR_SZ_BASE + 8, SRC, "Kolona ne postoji: " & vt & "." & vcol
        Dim cnt As Long: cnt = CountRowsByKey(vt, vpk, CStr(rows(i)(1)))
        If cnt = 0 Then Err.Raise ERR_SZ_BASE + 9, SRC, "Ciljni red ne postoji: " & vt & " " & CStr(rows(i)(1))
        If cnt > 1 Then Err.Raise ERR_SZ_BASE + 12, SRC, "Vise redova sa istim PK (" & vt & " " & CStr(rows(i)(1)) & ") -> odbijeno."
        ' DRIFT: trenutna vrednost mora biti tacno ono sto je storno OSTAVIO (NovaVrednost).
        Dim curV As String: curV = Trim$(CStr(LookupValue(vt, vpk, CStr(rows(i)(1)), vcol)))
        If StrComp(curV, Trim$(CStr(rows(i)(4))), vbBinaryCompare) <> 0 Then
            Err.Raise ERR_SZ_BASE + 13, SRC, "Stanje se promenilo posle storna (" & vt & "." & vcol & _
                " = '" & curV & "', ocekivano '" & CStr(rows(i)(4)) & "') -> undo odbijen (ne gazi noviju izmenu)."
        End If
        ' Otkup PO REDU: mrtav-roditelj (bas tog reda) + active-dup iste (broj,klasa) slot.
        If StrComp(vt, TBL_OTKUP, vbTextCompare) = 0 And StrComp(vcol, COL_STORNIRANO, vbTextCompare) = 0 Then
            Dim dpr As String: dpr = OtkupBlockDeadParentByID(CStr(rows(i)(1)))
            If Len(dpr) > 0 Then Err.Raise ERR_SZ_BASE + 15, SRC, "Roditelj/provera bloka " & _
                CStr(rows(i)(1)) & ": " & dpr & " -> undo bi ostavio siroce. Odbijeno."
            If OtkupReissueDupExists(CStr(rows(i)(1))) Then _
                Err.Raise ERR_SZ_BASE + 14, SRC, "Vec postoji AKTIVAN otkup iste (broj,klasa) kao red " & _
                    CStr(rows(i)(1)) & " -> undo bi duplirao. Odbijeno."
        End If
    Next i

    Set tx = New clsTransaction
    tx.BeginTx
    Dim snapped As Object: Set snapped = CreateObject("Scripting.Dictionary")
    snapped.CompareMode = vbTextCompare
    For i = 1 To rows.count
        Dim t As String: t = CStr(rows(i)(0))
        If Len(t) > 0 And Not snapped.Exists(t) Then tx.AddTableSnapshot t: snapped(t) = True
    Next i
    For i = 1 To rows.count
        RestoreCell CStr(rows(i)(0)), CStr(rows(i)(1)), CStr(rows(i)(2)), CStr(rows(i)(3)), SRC
    Next i
    tx.CommitTx: Set tx = Nothing

    UndoOperation_TX = True
    Monitor_Event eventType:="STORNO_UNDO_OP", severity:="INFO", _
        message:="Vrati storno (op " & opID & "): " & docType & " " & broj & " -> " & _
                 rows.count & " celija vraceno.", _
        moduleName:=MOD_NAME, procedureName:="UndoOperation_TX", _
        entityType:=docType, entityID:=broj, correlationId:=opID
    Exit Function
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr SRC
    UndoOperation_TX = False
End Function

Private Sub RestoreCell(ByVal tbl As String, ByVal rowID As String, _
                        ByVal col As String, ByVal oldVal As String, ByVal SRC As String)
    Dim pkCol As String: pkCol = PkColForTable(tbl)
    If Len(pkCol) = 0 Then Exit Sub
    Dim ri As Long: ri = FindRowIndexByKey(tbl, pkCol, rowID)
    If ri > 0 Then RequireUpdateCell tbl, ri, col, oldVal, SRC
End Sub

' Postoji li AKTIVAN otkup red iste (broj, klasa) kao dati stornirani red, a DRUGI PK?
' -> reizdaje isti slot -> undo bi duplirao. (Parcijalno-klasni undo je bezbedan jer
'  druga klasa ima drugaciju Klasu pa nije dup.)
' FAIL-CLOSED: greska/nedostajuca sema -> RAISE (UndoOperation_TX to hvata i odbija
' undo). Ne sme tiho vratiti False ("nema duplikata") kad provera nije izvedena.
Private Function OtkupReissueDupExists(ByVal otkupID As String) As Boolean
    Const SRC As String = MOD_NAME & ".OtkupReissueDupExists"
    Dim data As Variant: data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function
    Dim cId As Long, cBr As Long, cKl As Long, cSt As Long
    cId = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, SRC)
    cBr = RequireColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK, SRC)
    cKl = GetColumnIndex(TBL_OTKUP, COL_OTK_KLASA)
    cSt = GetColumnIndex(TBL_OTKUP, COL_STORNIRANO)
    Dim tBroj As String, tKlasa As String, i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cId))) = Trim$(otkupID) Then
            tBroj = Trim$(CStr(data(i, cBr)))
            If cKl > 0 Then tKlasa = Trim$(CStr(data(i, cKl)))
            Exit For
        End If
    Next i
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cId))) <> Trim$(otkupID) Then
            If Trim$(CStr(data(i, cBr))) = tBroj Then
                Dim kMatch As Boolean: kMatch = True
                If cKl > 0 Then kMatch = (Trim$(CStr(data(i, cKl))) = tKlasa)
                If kMatch Then
                    Dim isStor As Boolean: isStor = False
                    If cSt > 0 Then isStor = (UCase$(Trim$(CStr(data(i, cSt)))) = "DA")
                    If Not isStor Then OtkupReissueDupExists = True: Exit Function
                End If
            End If
        End If
    Next i
End Function

' PK kolona po tabeli (opseg faze 1: Otkup + Revers -> tblOtkup/Ambalaza/Novac).
Private Function PkColForTable(ByVal tbl As String) As String
    Select Case tbl
        Case TBL_OTKUP: PkColForTable = COL_OTK_ID
        Case TBL_AMBALAZA: PkColForTable = COL_AMB_ID
        Case TBL_NOVAC: PkColForTable = COL_NOV_ID
    End Select
End Function

Private Function FindRowIndexByKey(ByVal tbl As String, ByVal keyCol As String, _
                                   ByVal keyVal As String) As Long
    Dim data As Variant: data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Function
    Dim c As Long: c = GetColumnIndex(tbl, keyCol)
    If c = 0 Then Exit Function
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, c))) = Trim$(keyVal) Then FindRowIndexByKey = i: Exit Function
    Next i
End Function

Private Function CountRowsByKey(ByVal tbl As String, ByVal keyCol As String, _
                                ByVal keyVal As String) As Long
    Dim data As Variant: data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Function
    Dim c As Long: c = GetColumnIndex(tbl, keyCol)
    If c = 0 Then Exit Function
    Dim i As Long, n As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, c))) = Trim$(keyVal) Then n = n + 1
    Next i
    CountRowsByKey = n
End Function

' Najskoriji OperationID za (docType, broj) - za UI gate + fallback.
Public Function LatestOpFor(ByVal docType As String, ByVal broj As String) As String
    On Error GoTo EH
    Dim data As Variant: data = GetTableData(TBL_STORNO_ZURNAL)
    If IsEmpty(data) Then Exit Function
    Dim cOp As Long, cDoc As Long, cBroj As Long
    cOp = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_OP_ID)
    cDoc = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_DOCTYPE)
    cBroj = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_BROJ)
    If cOp = 0 Then Exit Function
    Dim i As Long, best As String, bestN As Long
    For i = 1 To UBound(data, 1)
        If StrComp(Trim$(CStr(data(i, cDoc))), docType, vbTextCompare) = 0 _
           And StrComp(Trim$(CStr(data(i, cBroj))), broj, vbTextCompare) = 0 Then
            Dim opv As String: opv = Trim$(CStr(data(i, cOp)))
            Dim numv As Long: numv = OpNum4(opv, "SOP-")
            If numv >= bestN Then bestN = numv: best = opv
        End If
    Next i
    LatestOpFor = best
    Exit Function
EH:
    LogErr MOD_NAME & ".LatestOpFor"
End Function

' Numericki deo ID-a sa datim prefiksom (SOP-/ZUR-). 0 ako ne pasuje.
Private Function OpNum4(ByVal id As String, ByVal prefix As String) As Long
    On Error Resume Next
    If Left$(id, Len(prefix)) = prefix Then OpNum4 = CLng(Mid$(id, Len(prefix) + 1))
End Function

' ============================================================
' READ-MODEL za operation-centric "Vrati storno" UI. Lista undoable operacija
' (NAJNOVIJE prvo): svaki dict {opID, ts, docType, broj, count, status}. status je
' informativan (prikaz): "moguce" | "vec vraceno" | "izmenjeno". UI zove
' UndoOperation_TX(opID) direktno -> cilja KONKRETNU operaciju (ne LatestOpFor po
' broju) -> resava reused-broj (razne generacije = razliciti OperationID).
' ============================================================
Public Function GetUndoableStornoOperations() As Collection
    Dim result As New Collection
    Set GetUndoableStornoOperations = result
    On Error GoTo EH
    Dim data As Variant: data = GetTableData(TBL_STORNO_ZURNAL)
    If IsEmpty(data) Then Exit Function
    Dim cOp As Long, cTs As Long, cDoc As Long, cBroj As Long
    Dim cTab As Long, cRow As Long, cCol As Long, cOld As Long, cNew As Long
    cOp = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_OP_ID)
    cTs = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_TS)
    cDoc = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_DOCTYPE)
    cBroj = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_BROJ)
    cTab = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_TABELA)
    cRow = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_ROWID)
    cCol = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_KOLONA)
    cOld = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_STARA)
    cNew = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_NOVA)
    If cOp = 0 Then Exit Function

    Dim ops As Object: Set ops = CreateObject("Scripting.Dictionary")
    ops.CompareMode = vbTextCompare
    Dim order As Collection: Set order = New Collection
    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim op As String: op = Trim$(CStr(data(i, cOp)))
        If Len(op) > 0 Then
            If Not ops.Exists(op) Then
                Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
                d("opID") = op
                d("ts") = TxZ(data, i, cTs)
                d("docType") = TxZ(data, i, cDoc)
                d("broj") = TxZ(data, i, cBroj)
                d("count") = 0&
                d("firstTab") = TxZ(data, i, cTab)
                d("firstRow") = TxZ(data, i, cRow)
                d("firstCol") = TxZ(data, i, cCol)
                d("firstOld") = TxZ(data, i, cOld)
                d("firstNew") = TxZ(data, i, cNew)
                Set ops(op) = d
                order.Add op
            End If
            ops(op)("count") = CLng(ops(op)("count")) + 1
        End If
    Next i

    Dim k As Long
    For k = order.count To 1 Step -1        ' najnovije (poslednji upisan) na vrh
        Dim o As Object: Set o = ops(CStr(order(k)))
        o("status") = OpStatusFor(o)
        result.Add o
    Next k
    Exit Function
EH:
    LogErr MOD_NAME & ".GetUndoableStornoOperations"
End Function

' Informativni status operacije po PRVOM redu (trenutna vrednost vs NovaVrednost/StaraVrednost).
Private Function OpStatusFor(ByVal o As Object) As String
    On Error GoTo unknown
    Dim tbl As String: tbl = CStr(o("firstTab"))
    Dim pk As String: pk = PkColForTable(tbl)
    If Len(pk) = 0 Then OpStatusFor = "?": Exit Function
    Dim cur As String: cur = Trim$(CStr(LookupValue(tbl, pk, CStr(o("firstRow")), CStr(o("firstCol")))))
    If StrComp(cur, Trim$(CStr(o("firstNew"))), vbBinaryCompare) = 0 Then OpStatusFor = "moguce": Exit Function
    If StrComp(cur, Trim$(CStr(o("firstOld"))), vbBinaryCompare) = 0 Then OpStatusFor = "vec vraceno": Exit Function
    OpStatusFor = "izmenjeno"
    Exit Function
unknown:
    OpStatusFor = "?"
End Function

' `data` je ByRef namerno. ByVal na Variantu koji SADRZI niz kopira ceo niz pri
' svakom pozivu, a ovo je citac PO CELIJI -- u petlji se zove vise puta po redu.
' Mereno na istom obrascu u modPaletniList.SafeCell: 1063 stavke, 1883 ms, to jest
' 1.8 ms po redu za citanje dva polja iz niza koji je vec u memoriji.
'
' Funkcija niz samo CITA, nikad ne pise, pa je razlika iskljucivo u tome sto se
' niz ne umnozava.
Private Function TxZ(ByRef data As Variant, ByVal r As Long, ByVal c As Long) As String
    If c > 0 Then TxZ = Trim$(CStr(data(r, c)))
End Function
