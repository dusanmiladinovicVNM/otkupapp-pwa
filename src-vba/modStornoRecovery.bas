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
' UndoStorno_TX (MUTIRA): konzervativna reverzija soft-delete-a SAMO za SAMOSTALNE
'   tipove -> Otkup (blok) i Revers (OM-koop ambalaza). Chain dokumenti
'   (Otpremnica/Zbirna/Prijemnica/Faktura) se NAMERNO odbijaju: njihova reverzija
'   trazi grupni operativni zurnal koji ne postoji -> koristi ISPRAVKA / ponovni
'   unos. Guard: ne dozvoli reverziju ako vec postoji AKTIVAN dokument istog broja
'   (duplo). Reaktivira i ambalaza-ledger dokumenta. Sve u jednoj transakciji.
'
' UI se NAMERNO NE vezuje u prvom prolazu: reverzija se prvo verifikuje kroz
' Test_UndoStorno (Alt+F8) na realnim podacima, pa se dugme vezuje kasnije.
' ============================================================

Private Const MOD_NAME As String = "modStornoRecovery"
Private Const ERR_REC_BASE As Long = vbObjectError + 2900
Private Const STORNO_DA As String = "Da"

' ============================================================
' NEDOVRSENO - ujedinjen read-only pregled.
' ============================================================
Public Function GetNedovrseno() As Collection
    Dim result As New Collection
    Set GetNedovrseno = result
    On Error GoTo EH

    ' 1) Persistentni contexti (tblStornoVeze) - strukturirano.
    Dim ctx As Collection: Set ctx = GetPendingCorrections()
    If Not ctx Is Nothing Then
        Dim i As Long
        For i = 1 To ctx.count
            Dim c As Object: Set c = ctx(i)
            AddNedRow result, "CONTEXT/" & CStr(c("mode")), CStr(c("oldBroj")), _
                      CStr(c("status")), CStr(c("message")), CStr(c("recoveryAction"))
        Next i
    End If

    ' 2) Osirocene - samo brojevi (detalj/akcije: Osiroceni dokumenti panel).
    AddCountRow result, "OSIROCENE_PRIJEMNICE", CountVar(GetOsirocenePrijemnice())
    AddCountRow result, "OSIROCENE_PALETE", CountVar(GetPrijemniceSaOsirocenimPaletama())
    AddCountRow result, "IZGUBLJENI_BLOKOVI", CountVar(GetLostOtkupBlokovi())
    Exit Function
EH:
    LogErr MOD_NAME & ".GetNedovrseno"
End Function

Private Sub AddNedRow(ByRef col As Collection, ByVal kind As String, ByVal ref As String, _
                      ByVal status As String, ByVal opis As String, ByVal akcija As String)
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("kind") = kind: d("ref") = ref: d("status") = status
    d("opis") = opis: d("akcija") = akcija
    col.Add d
End Sub

Private Sub AddCountRow(ByRef col As Collection, ByVal kind As String, ByVal n As Long)
    If n <= 0 Then Exit Sub
    AddNedRow col, kind, CStr(n), "", "ukupno: " & n, "Osiroceni dokumenti"
End Sub

Private Function CountVar(ByVal v As Variant) As Long
    On Error Resume Next
    If Not IsEmpty(v) Then
        If IsArray(v) Then CountVar = UBound(v, 1)
    End If
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
Public Function OtkupBlockDeadParent(ByVal broj As String) As String
    On Error Resume Next
    Dim data As Variant: data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function
    Dim cBr As Long, cSt As Long, cOtp As Long, cZbr As Long
    cBr = GetColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK)
    cSt = GetColumnIndex(TBL_OTKUP, COL_STORNIRANO)
    cOtp = GetColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID)
    cZbr = GetColumnIndex(TBL_OTKUP, COL_OTK_BROJ_ZBIRNE)
    If cBr = 0 Or cSt = 0 Then Exit Function
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
