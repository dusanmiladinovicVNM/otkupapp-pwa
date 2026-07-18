Attribute VB_Name = "modStornoRecovery"
Option Explicit

' ============================================================
' modStornoRecovery - "Nedovrseno" (jedan pregled) + "Vrati storno" (reverzija).
' Faza 5.
'
' GetNedovrseno (READ-ONLY): ujedinjen pregled onoga sto ceka -> PENDING/MANUAL
'   contexti (tblStornoVeze, strukturirano) + brojevi osirocenih (prijemnice /
'   palete / izgubljeni blokovi; detalj i akcije su u recovery panelu).
'
' UndoStorno_TX (MUTIRA): reverzija soft-delete-a SAMO za SAMOSTALNE tipove ->
'   Otkup (blok) i Revers (OM-koop ambalaza). Chain dokumenti (Otpremnica/Zbirna/
'   Prijemnica/Faktura) se NAMERNO odbijaju -> koristi ISPRAVKA / ponovni unos.
'   Za storna napravljena POSLE storno-zurnala (modStornoZurnal) prvo pokusava
'   LOSSLESS put (UndoOperation_TX preko LatestOpFor) -> vraca i tblNovac.OtkupID i
'   cilja bas tu operaciju. Stara storna (bez zurnala) padaju na legacy best-effort
'   (ne vraca novac vezu) -> produkciono dugme to ODBIJA (LatestOpFor=""); legacy je
'   dostupan jedino kroz Test_UndoStorno macro.
'
' UI dugme "Vrati storno" je uvezano (frmDokumenta, UNDO_UI_ENABLED) i cilja
' konkretan OperationID; garde su fail-closed. Guard: UndoGuardReason (deljena).
' ============================================================

Private Const MOD_NAME As String = "modStornoRecovery"
Private Const ERR_REC_BASE As Long = vbObjectError + 2900
Private Const STORNO_DA As String = "Da"

' ============================================================
' NEDOVRSENO - ujedinjen read-only pregled.
' ============================================================
' Svaki red nosi PUN strukturirani zapis (ne samo tekst) -> panel moze da deluje po
' redu. Kljucevi: kind, ref (poslovni broj), status, opis, akcija (tekst), correctionID,
' docType, newBroj, mode, actionCode (CONTEXT | PRIJ | PAL | BLOK -> UI dispecuje).
' Dedup: osirocene se izostave ako isti poslovni broj vec nosi PENDING/MANUAL context
' (isti problem se ne prikazuje dvaput).
Public Function GetNedovrseno() As Collection
    Dim result As New Collection
    Set GetNedovrseno = result
    On Error GoTo EH

    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    ' 1) Persistentni contexti (tblStornoVeze) - strukturirano, sa CorrectionID.
    Dim ctx As Collection: Set ctx = GetPendingCorrections()
    If Not ctx Is Nothing Then
        Dim i As Long
        For i = 1 To ctx.count
            Dim c As Object: Set c = ctx(i)
            Dim oB As String: oB = CStr(c("oldBroj"))
            AddNedRowFull result, "CONTEXT/" & CStr(c("mode")), oB, CStr(c("status")), _
                CStr(c("message")), CStr(c("recoveryAction")), _
                CStr(c("id")), CStr(c("oldDocType")), CStr(c("newBroj")), CStr(c("mode")), "CONTEXT"
            ' brojevi koje context vec pokriva -> osirocene za njih preskoci (dedup).
            If Len(oB) > 0 Then seen(oB) = True
            If Len(CStr(c("newBroj"))) > 0 Then seen(CStr(c("newBroj"))) = True
            If Len(CStr(c("parentBroj"))) > 0 Then seen(CStr(c("parentBroj"))) = True
        Next i
    End If

    ' 2) Osirocene - PER-STAVKA red (poslovni broj + akcija), deduplikovano protiv contexta.
    '    Akcija svih: "Osiroceni dokumenti" panel (postojeci re-point / skidanje).
    AddOsiroceneRows result, GetOsirocenePrijemnice(), "OSIROCENA_PRIJEMNICA", 1, _
        "Prijemnica", "PRIJ", "Osiroceni dokumenti (prevezi prijemnicu)", seen
    AddOsiroceneRows result, GetPrijemniceSaOsirocenimPaletama(), "OSIROCENE_PALETE", 1, _
        "Prijemnica (palete)", "PAL", "Osiroceni dokumenti (Mod: Palete)", seen
    AddOsiroceneRows result, GetLostOtkupBlokovi(), "IZGUBLJEN_BLOK", 2, _
        "Otkupni blok", "BLOK", "Otkupni blokovi (Preuzmi / prevezi)", seen
    Exit Function
EH:
    LogErr MOD_NAME & ".GetNedovrseno"
End Function

Private Sub AddNedRowFull(ByRef col As Collection, ByVal kind As String, ByVal ref As String, _
                          ByVal status As String, ByVal opis As String, ByVal akcija As String, _
                          ByVal correctionID As String, ByVal docType As String, _
                          ByVal newBroj As String, ByVal mode As String, ByVal actionCode As String)
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("kind") = kind: d("ref") = ref: d("status") = status
    d("opis") = opis: d("akcija") = akcija
    d("correctionID") = correctionID: d("docType") = docType
    d("newBroj") = newBroj: d("mode") = mode: d("actionCode") = actionCode
    col.Add d
End Sub

' Prosiri osirocene (2D niz iz Get*-funkcija) u pojedinacne redove. brojCol = 1-based
' kolona sa poslovnim brojem. Preskace broj koji je vec u `seen` (context dedup + interni).
Private Sub AddOsiroceneRows(ByRef result As Collection, ByVal v As Variant, ByVal kind As String, _
                             ByVal brojCol As Long, ByVal docLabel As String, ByVal actionCode As String, _
                             ByVal akcija As String, ByRef seen As Object)
    On Error GoTo done
    If Not Is2DArray(v) Then Exit Sub
    Dim r As Long
    For r = 1 To UBound(v, 1)
        Dim b As String: b = Trim$(CStr(v(r, brojCol)))
        If Len(b) > 0 Then
            If Not seen.Exists(b) Then
                seen(b) = True
                AddNedRowFull result, kind, b, "OSIROCENO", docLabel & " " & b & " -- ceka doradu", _
                    akcija, "", docLabel, "", "", actionCode
            End If
        End If
    Next r
done:
End Sub

Private Function Is2DArray(ByVal v As Variant) As Boolean
    On Error GoTo no
    If IsEmpty(v) Then Exit Function
    If Not IsArray(v) Then Exit Function
    Dim probe As Long: probe = UBound(v, 2)   ' baci gresku ako nije 2D
    Is2DArray = (UBound(v, 1) >= 1)
    Exit Function
no:
End Function

' ============================================================
' VRATI STORNO - konzervativna reverzija (Otkup / Revers). MUTIRA.
' ============================================================
Public Function UndoStorno_TX(ByVal docType As String, ByVal broj As String, _
                              Optional ByVal dokumentTip As String = "") As Boolean
    Const SRC As String = MOD_NAME & ".UndoStorno_TX"
    Dim tx As clsTransaction
    On Error GoTo EH
    broj = Trim$(broj)
    If Len(broj) = 0 Then Err.Raise ERR_REC_BASE + 1, SRC, "Broj dokumenta je obavezan."

    ' LOSSLESS put: ako postoji storno-zurnal operacija za (docType, broj) -> pravi
    ' inverz preko zurnala (vraca i tblNovac.OtkupID + cilja bas tu generaciju).
    ' Stara storna (pre zurnala) padaju na legacy best-effort ispod.
    Dim opID As String: opID = LatestOpFor(docType, broj)
    If Len(opID) > 0 Then
        UndoStorno_TX = UndoOperation_TX(opID)
        Exit Function
    End If

    Select Case docType
        Case DOK_TIP_OTKUP
            ' Guard: ako vec postoji AKTIVAN otkup istog broja -> reverzija bi duplirala.
            If Len(LookupActiveID(TBL_OTKUP, COL_OTK_BR_DOK, broj, COL_OTK_ID)) > 0 Then
                Err.Raise ERR_REC_BASE + 2, SRC, "Vec postoji AKTIVAN otkup " & broj & _
                          " -> reverzija bi duplirala. Odbijeno."
            End If
            ' Guard (B3 - izgubljeni blok): ako je roditeljski lanac (otpremnica ILI
            ' zbirna na koje je blok vezan) stornirano, reaktivacija bi napravila siroce
            ' (aktivan blok bez ziveg roditelja). Kaskadni storno oslobodi SAMO aktivne
            ' blokove -> vec-storniran blok zadrzi vezu ka mrtvom roditelju. Dup-guard to
            ' ne hvata. Odbij; koristi ponovni unos / prevezivanje (Osiroceni dokumenti).
            Dim deadParent As String: deadParent = OtkupBlockDeadParent(broj)
            If Len(deadParent) > 0 Then
                Err.Raise ERR_REC_BASE + 5, SRC, "Ne mogu da vratim storno otkupa " & broj & _
                    ": roditeljski " & deadParent & " je storniran -> blok bi ostao siroce. " & _
                    "Unesi blok ponovo ili ga prevezi (Osiroceni dokumenti)."
            End If
            Set tx = New clsTransaction
            tx.BeginTx
            tx.AddTableSnapshot TBL_OTKUP
            tx.AddTableSnapshot TBL_AMBALAZA
            Dim ids As Collection: Set ids = New Collection
            Dim n As Long
            n = UnmarkStorniranoCollect(TBL_OTKUP, COL_OTK_BR_DOK, broj, COL_OTK_ID, ids, SRC)
            If n = 0 Then Err.Raise ERR_REC_BASE + 3, SRC, "Nema storniranog otkupa " & broj & "."
            Dim k As Long
            For k = 1 To ids.count
                UnmarkAmbalazaByDokument CStr(ids(k)), DOK_TIP_OTKUP, SRC
                UnmarkAmbalazaByDokument CStr(ids(k)), DOK_TIP_OM_IZLAZ_KOOP, SRC
            Next k
            tx.CommitTx: Set tx = Nothing
            UndoStorno_TX = True
            MonUndo SRC, "Otkup", broj, "Vraceno " & n & " redova + ambalaza."

        Case DOK_TIP_OM_IZLAZ_KOOP, DOK_TIP_OM_ULAZ_KOOP, DOK_TIP_OM_IZLAZ_FIRMA, DOK_TIP_OM_ULAZ_FIRMA
            ' Revers je list (samo tblAmbalaza redovi po broj+dokTip). dokTip = docType.
            ' Guard (paralela otkup dup-guardu): ako vec postoji AKTIVAN revers istog
            ' broj+tip -> reverzija bi duplirala. Odbij. (Ranije je fantomski unmark-ovao
            ' sve stornirane redove i bez ove provere.)
            If ActiveAmbalazaDokExists(broj, docType) Then
                Err.Raise ERR_REC_BASE + 6, SRC, "Vec postoji AKTIVAN revers " & broj & _
                    " [" & docType & "] -> reverzija bi duplirala. Odbijeno."
            End If
            Set tx = New clsTransaction
            tx.BeginTx
            tx.AddTableSnapshot TBL_AMBALAZA
            Dim m As Long: m = UnmarkAmbalazaByDokument(broj, docType, SRC)
            If m = 0 Then Err.Raise ERR_REC_BASE + 4, SRC, _
                          "Nema storniranog reversa " & broj & " [" & docType & "]."
            tx.CommitTx: Set tx = Nothing
            UndoStorno_TX = True
            MonUndo SRC, docType, broj, "Vracen revers (" & m & " redova)."

        Case Else
            Err.Raise ERR_REC_BASE + 9, SRC, "Vrati storno je podrzan SAMO za Otkup i Revers. " & _
                      "Chain dokument (" & docType & ") -> koristi ISPRAVKA / ponovni unos."
    End Select
    Exit Function
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr SRC
    UndoStorno_TX = False
End Function

' Skini "Stornirano" sa svih redova gde keyCol=keyVal (a bilo je stornirano);
' sakupi njihove idCol vrednosti u outIds. Vraca broj vracenih redova.
Private Function UnmarkStorniranoCollect(ByVal tbl As String, ByVal keyCol As String, _
        ByVal keyVal As String, ByVal idCol As String, ByRef outIds As Collection, _
        ByVal SRC As String) As Long
    Dim data As Variant: data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Function
    Dim cKey As Long, cId As Long, cSt As Long
    cKey = GetColumnIndex(tbl, keyCol)
    cId = GetColumnIndex(tbl, idCol)
    cSt = GetColumnIndex(tbl, COL_STORNIRANO)
    If cKey = 0 Or cSt = 0 Then Exit Function
    Dim i As Long, n As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cKey))) = Trim$(keyVal) _
           And UCase$(Trim$(CStr(data(i, cSt)))) = UCase$(STORNO_DA) Then
            RequireUpdateCell tbl, i, COL_STORNIRANO, "", SRC
            If cId > 0 Then outIds.Add Trim$(CStr(data(i, cId)))
            n = n + 1
        End If
    Next i
    UnmarkStorniranoCollect = n
End Function

' Da li bi reaktivacija storniranih blokova broja "broj" napravila siroce?
' Za svaki storniran red proveri vezu: OtpremnicaID -> aktivna otpremnica?
' BrojZbirne -> aktivna zbirna? Ako je bilo koja veza ka MRTVOM (stornirano/nema)
' roditelju -> vrati opis (prvog) mrtvog roditelja; inace "". Unbound blok (bez
' veze) i blok sa zivim roditeljem su bezbedni za undo.
' FAIL-CLOSED: sada je deo produkcijske undo kapije -> na gresku/nedostajucu semu
' vraca BLOKIRAJUCI marker (ne prazan string). "" znaci samo "roditelj je ziv/bezbedno".
Public Function OtkupBlockDeadParent(ByVal broj As String) As String
    On Error GoTo EH
    Dim data As Variant: data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function
    Dim cBr As Long, cSt As Long, cOtp As Long, cZbr As Long
    cBr = GetColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK)
    cSt = GetColumnIndex(TBL_OTKUP, COL_STORNIRANO)
    cOtp = GetColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID)
    cZbr = GetColumnIndex(TBL_OTKUP, COL_OTK_BROJ_ZBIRNE)
    If cBr = 0 Or cSt = 0 Then OtkupBlockDeadParent = "(provera roditelja nije moguca - sema)": Exit Function
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cBr))) = Trim$(broj) _
           And UCase$(Trim$(CStr(data(i, cSt)))) = UCase$(STORNO_DA) Then
            If cOtp > 0 Then
                Dim otpID As String: otpID = Trim$(CStr(data(i, cOtp)))
                If Len(otpID) > 0 Then
                    If Len(LookupActiveID(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_ID)) = 0 Then
                        OtkupBlockDeadParent = "otpremnica (ID " & otpID & ")": Exit Function
                    End If
                End If
            End If
            If cZbr > 0 Then
                Dim brZbr As String: brZbr = Trim$(CStr(data(i, cZbr)))
                If Len(brZbr) > 0 Then
                    If Len(LookupActiveID(TBL_ZBIRNA, COL_ZBR_BROJ, brZbr, COL_ZBR_BROJ)) = 0 Then
                        OtkupBlockDeadParent = "zbirna " & brZbr: Exit Function
                    End If
                End If
            End If
        End If
    Next i
    Exit Function
EH:
    LogErr MOD_NAME & ".OtkupBlockDeadParent"
    OtkupBlockDeadParent = "(provera roditelja nije uspela)"     ' fail-closed -> blokira undo
End Function

' Reaktiviraj tblAmbalaza redove dokumenta (DokID + DokTip) koji su stornirani.
Private Function UnmarkAmbalazaByDokument(ByVal dokID As String, ByVal dokTip As String, _
                                          ByVal SRC As String) As Long
    Dim data As Variant: data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then Exit Function
    Dim cDok As Long, cTip As Long, cSt As Long
    cDok = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_ID)
    cTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_TIP)
    cSt = GetColumnIndex(TBL_AMBALAZA, COL_STORNIRANO)
    If cDok = 0 Or cTip = 0 Or cSt = 0 Then Exit Function
    Dim i As Long, n As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cDok))) = Trim$(dokID) _
           And Trim$(CStr(data(i, cTip))) = Trim$(dokTip) _
           And UCase$(Trim$(CStr(data(i, cSt)))) = UCase$(STORNO_DA) Then
            RequireUpdateCell TBL_AMBALAZA, i, COL_STORNIRANO, "", SRC
            n = n + 1
        End If
    Next i
    UnmarkAmbalazaByDokument = n
End Function

' Postoji li AKTIVAN (ne-storniran) ambalaza red za dati dokument (broj) + tip?
' Guard za reverse undo -> spreci duplikat ako revers vec ima zivu verziju.
' FAIL-CLOSED: nedostajuca sema/greska -> RAISE (UndoGuardReason to hvata i BLOKIRA).
' Ne sme tiho vratiti False ("nema duplikata") kad provera nije izvedena.
Private Function ActiveAmbalazaDokExists(ByVal dokID As String, ByVal dokTip As String) As Boolean
    Const SRC As String = MOD_NAME & ".ActiveAmbalazaDokExists"
    Dim data As Variant: data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then Exit Function
    Dim cDok As Long, cTip As Long, cSt As Long
    cDok = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_ID, SRC)
    cTip = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_TIP, SRC)
    cSt = RequireColumnIndex(TBL_AMBALAZA, COL_STORNIRANO, SRC)
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cDok))) = Trim$(dokID) _
           And Trim$(CStr(data(i, cTip))) = Trim$(dokTip) Then
            Dim isStor As Boolean: isStor = False
            If cSt > 0 Then isStor = (UCase$(Trim$(CStr(data(i, cSt)))) = UCase$(STORNO_DA))
            If Not isStor Then ActiveAmbalazaDokExists = True: Exit Function
        End If
    Next i
End Function

' ZAJEDNICKA broj-level undo garda (i legacy put i zurnal-put UndoOperation_TX).
' Vraca "" ako je bezbedno; inace razlog. Otkup: mrtav-roditelj (fail-closed).
' Revers (OM): aktivan-dup ambalaze istog broj+tip (#134). Otkup active-dup se NE
' proverava ovde (bio je broj-level pa je preblokirao parcijalni storno jedne klase)
' -> zurnal-put to radi PO REDU (OtkupReissueDupExists po (broj,klasa)).
' FAIL-CLOSED: greska u proveri -> blokirajuci razlog (ne dozvoli tih prolaz).
Public Function UndoGuardReason(ByVal docType As String, ByVal broj As String) As String
    On Error GoTo EH
    Select Case docType
        Case DOK_TIP_OTKUP
            Dim dead As String: dead = OtkupBlockDeadParent(broj)
            If Len(dead) > 0 Then _
                UndoGuardReason = "Roditelj/provera: " & dead & " -> blok bi ostao siroce / nesigurno. Odbijeno."
        Case DOK_TIP_OM_IZLAZ_KOOP, DOK_TIP_OM_ULAZ_KOOP, DOK_TIP_OM_IZLAZ_FIRMA, DOK_TIP_OM_ULAZ_FIRMA
            If ActiveAmbalazaDokExists(broj, docType) Then _
                UndoGuardReason = "Vec postoji AKTIVAN revers " & broj & " [" & docType & _
                    "] -> undo bi duplirao. Odbijeno."
    End Select
    Exit Function
EH:
    LogErr MOD_NAME & ".UndoGuardReason"
    UndoGuardReason = "Greska pri proveri undo garde -> odbijeno (fail-closed)."
End Function

Private Sub MonUndo(ByVal procName As String, ByVal entityType As String, _
                   ByVal entityID As String, ByVal msg As String)
    On Error Resume Next
    Monitor_Event _
        eventType:="UNDO_STORNO_" & UCase$(entityType), severity:="INFO", _
        message:="Vrati storno: " & entityType & " " & entityID & " -> " & msg, _
        userId:="Operator", moduleName:=MOD_NAME, procedureName:=procName, _
        entityType:=entityType, entityID:=entityID, correlationId:=entityID
End Sub

' ============================================================
' TEST HOOK-ovi (Alt+F8) - verifikacija PRE vezivanja UI-a.
' ============================================================
Public Sub Test_GetNedovrseno()
    Dim c As Collection: Set c = GetNedovrseno()
    Debug.Print "=== NEDOVRSENO (" & IIf(c Is Nothing, 0, c.count) & ") ==="
    Dim i As Long
    If Not c Is Nothing Then
        For i = 1 To c.count
            Dim d As Object: Set d = c(i)
            Debug.Print d("kind") & " | ref=" & d("ref") & " | " & d("status") & " | " & d("opis")
        Next i
    End If
    Debug.Print "=== kraj ==="
End Sub

Public Sub Test_UndoStorno()
    Dim tip As String
    tip = Trim$(InputBox("Tip za reverziju:" & vbCrLf & _
        "- 'Otkup'" & vbCrLf & "- revers dokTip: OM-Izlaz-Koop / OM-Ulaz-Koop / OM-Izlaz-Firma / OM-Ulaz-Firma", _
        "Vrati storno - TEST"))
    If Len(tip) = 0 Then Exit Sub
    Dim broj As String
    broj = Trim$(InputBox("Broj STORNIRANOG dokumenta:", "Vrati storno - TEST"))
    If Len(broj) = 0 Then Exit Sub

    Dim ok As Boolean: ok = UndoStorno_TX(tip, broj)
    If ok Then
        MsgBox "Vrati storno USPEO: " & tip & " " & broj, vbInformation, "Vrati storno"
    Else
        MsgBox "Vrati storno NIJE uspeo za " & tip & " " & broj & "." & vbCrLf & _
               "(Guard/greska -> vidi Immediate / Monitor. Chain dokumenti se odbijaju.)", _
               vbExclamation, "Vrati storno"
    End If
End Sub
