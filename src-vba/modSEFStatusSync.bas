Attribute VB_Name = "modSEFStatusSync"
 Option Explicit

' =========================================================
' OUTBOUND STATUS MODEL
'
' SEFWorkflowState = internal/local process control state
' SEFStatus        = exact latest external status returned by SEF API
'
' These two fields are related but do NOT have to be identical.
'
' Examples:
'   SEFWorkflowState = SEF_SENT,     SEFStatus = SENT
'   SEFWorkflowState = SEF_SENT,     SEFStatus = DRAFT
'   SEFWorkflowState = SEF_SENT,     SEFStatus = STORNO
'   SEFWorkflowState = SEF_ACCEPTED, SEFStatus = ACCEPTED
'   SEFWorkflowState = SEF_REJECTED, SEFStatus = REJECTED
'
' WorkflowState changes only when the LOCAL state machine changes.
' SEFStatus is updated on every successful refresh from SEF.
' =========================================================

Public Function RefreshSubmissionStatus(ByVal fakturaID As String) As Boolean
    RefreshSubmissionStatus = RefreshSEFStatus_TX(fakturaID)
End Function

' =========================================================
' ADAPTER: zvanicni SEF status -> interna klasa (AUD-032b)
'
' Zvanicni enum SalesInvoiceStatus (SEF tehnicko uputstvo):
'   New, Draft, Sending, Sent, Seen, Approved, Rejected, Cancelled, Storno,
'   Paid, OverDue, Archived, Mistake, Deleted, Unknown
'
' KLJUCNO: prihvatanje se na SEF-u zove **Approved**, ne "Accepted". Stari kod je
' prepoznavao samo "ACCEPTED", pa je stvarno odobrena faktura padala u Case Else.
' Dok je Case Else bio "SENT", to se nije videlo (AUD-032b); da je ostalo posle
' popravke, odobrena faktura bi zavrsavala kao "nepoznat status - rucna provera".
'
' Nepoznat status (ukljucujuci zvanicni "Unknown" i prazan) NIKAD ne sme da se
' protumaci kao poslato/prihvaceno - ide na rucnu proveru.
' =========================================================
Public Function ClassifySEFExternalStatus(ByVal apiStatus As String) As String

    Select Case UCase$(Trim$(apiStatus))

        ' Kupac je odobrio fakturu. "ACCEPTED" se zadrzava jer ga
        ' ParseSubmitResponse postavlja iz submit odgovora (accepted=true) i jer
        ' postoje zatecene fakture sa tom vrednoscu u tblFakture.
        Case "APPROVED", "ACCEPTED"
            ClassifySEFExternalStatus = SEF_CLS_ACCEPTED

        Case "REJECTED"
            ClassifySEFExternalStatus = SEF_CLS_REJECTED

        ' Jos nije doneta odluka kupca -- faktura je na putu ili ceka.
        Case "NEW", "DRAFT", "DRAFTED", "SENDING", "SENT", "SEEN"
            ClassifySEFExternalStatus = SEF_CLS_PENDING

        ' Storno ima LOKALNI parnjak (WF_SEF_STORNO) -- zato zasebna klasa.
        Case "STORNO"
            ClassifySEFExternalStatus = SEF_CLS_STORNO

        ' Otkazan/obrisan dokument nema lokalno stanje: lokalni state machine
        ' NEMA WF_SEF_CANCELLED. Ovo je namerno "external-terminal-only" --
        ' belezi se u SEFStatus, workflow se samo izvlaci iz "salje se".
        Case "CANCELLED", "CANCELED", "DELETED"
            ClassifySEFExternalStatus = SEF_CLS_TERMINAL

        ' "Mistake" = greska prilikom slanja dokumenta. Namerno NIJE u
        ' SEF_CLS_TERMINAL: kad je bio tamo, planer ga je vodio u WF_SEF_SENT, pa
        ' je NEUSPELO slanje lokalno postajalo "poslato" -- batch ga je preskakao
        ' kao terminalan, Cancel je bio zakljucan, a Retry nedostupan jer
        ' workflow nije SEF_TECH_FAILED. Faktura bi ostala bez ijedne putanje.
        Case "MISTAKE"
            ClassifySEFExternalStatus = SEF_CLS_SEND_FAILED

        ' Poznati statusi koji NE govore nista o odluci kupca (izdavalac moze da
        ' obelezi placeno, rok moze da istekne, dokument moze biti arhiviran).
        ' ALI: svi dokazuju da je dokument u SEF lifecycle-u, pa izvlace fakturu
        ' iz "salje se" -- samo je ne proglasavaju prihvacenom.
        Case "PAID", "OVERDUE", "OVER_DUE", "ARCHIVED"
            ClassifySEFExternalStatus = SEF_CLS_INFO

        Case "ERROR"
            ClassifySEFExternalStatus = SEF_CLS_ERROR

        ' Ovde padaju i zvanicni "Unknown" i prazan status i svaki nov status
        ' koji SEF uvede posle ove verzije.
        Case Else
            ClassifySEFExternalStatus = SEF_CLS_UNKNOWN

    End Select

End Function

' Tanak omotac nad klasifikatorom -- "da li SEF sloj uopste razume ovaj status".
Public Function IsKnownSEFRefreshStatus(ByVal apiStatus As String) As Boolean
    IsKnownSEFRefreshStatus = _
        (ClassifySEFExternalStatus(apiStatus) <> SEF_CLS_UNKNOWN)
End Function

' Da li je refresh dao upotrebljiv odgovor (za povratnu vrednost i brojace).
Public Function IsUsableSEFRefreshClass(ByVal classification As String) As Boolean
    Select Case classification
        Case SEF_CLS_ACCEPTED, SEF_CLS_REJECTED, SEF_CLS_PENDING, _
             SEF_CLS_STORNO, SEF_CLS_TERMINAL, SEF_CLS_INFO, SEF_CLS_SEND_FAILED
            ' SEND_FAILED je upotrebljiv odgovor: SEF nam je jasno rekao da
            ' slanje nije uspelo. Refresh je odradio svoje, faktura ide u retry.
            IsUsableSEFRefreshClass = True
        Case Else
            IsUsableSEFRefreshClass = False
    End Select
End Function

' =========================================================
' PLANER TRANZICIJE (AUD-032b/c)
'
' Jedina odluka "gde workflow ide posle refresh-a". Vraca ciljno stanje ili
' PRAZAN string = "ne diraj workflow, upisi samo refresh/error polja".
'
' Funkcija je cista i **sama proverava dozvoljenost** kroz
' IsSEFTransitionAllowed, pa po konstrukciji ne moze da predlozi tranziciju
' koju modSEFValidator odbija. Bez toga su nastajale bas ove kontradikcije:
'   SEF_SYNC_ERROR + pad API-ja -> SEF_SYNC_ERROR (self-transition, zabranjena)
'   SEF_UNKNOWN    + pad API-ja -> SEF_SYNC_ERROR (zabranjena)
' obe su zavrsavale izuzetkom i rollback-om bas kad SEF ponovo ne odgovara.
'
' Dvokorak (npr. SEF_SYNC_ERROR -> SEF_SENT -> SEF_ACCEPTED) se dobija time sto
' pozivalac planer zove ponovo sa novim stanjem; vidi ApplySEFExternalOutcome_Row.
' =========================================================
Public Function SEFRefreshTargetState(ByVal currentState As String, _
                                      ByVal classification As String) As String

    Dim curState As String
    Dim desired As String

    curState = UCase$(Trim$(currentState))

    Select Case classification

        Case SEF_CLS_ACCEPTED
            desired = WF_SEF_ACCEPTED

        Case SEF_CLS_REJECTED
            desired = WF_SEF_REJECTED

        Case SEF_CLS_SEND_FAILED
            ' Greska pri slanju (SEF "Mistake"): faktura mora iz "poslato" u
            ' SEF_TECH_FAILED, jer NIJE poslata. To NE znaci da se automatski
            ' nudi retry -- dokument na SEF-u postoji, pa `CanSendSEFInvoice`
            ' drzi slanje zatvorenim dok se ne otkaze (AUD-032b, runda 4).
            desired = WF_SEF_TECH_FAILED

        Case SEF_CLS_STORNO
            ' AUD-032b: uspesan/registrovan storno mora da pomeri i LOKALNI
            ' workflow, inace ostaje trajna kontradikcija
            ' (SEFWorkflowState = SEF_SENT/SEF_ACCEPTED, SEFStatus = STORNO),
            ' a batch takvu fakturu preskace kao terminalnu pa je niko vise ne
            ' ispravlja. Iz SEF_SENDING/SEF_UNKNOWN/SEF_SYNC_ERROR state machine
            ' nema direktan put, pa most preko SEF_SENT resava drugi korak.
            desired = WF_SEF_STORNO

        Case SEF_CLS_PENDING, SEF_CLS_TERMINAL, SEF_CLS_INFO
            ' Sve tri klase dokazuju da je dokument stigao u SEF lifecycle, pa
            ' izvlace zaglavljeni SEF_SENDING / SEF_UNKNOWN / SEF_SYNC_ERROR na
            ' SEF_SENT. Iz SEF_SENT je to self-transition, pa planer ispod vrati
            ' prazno i upisu se samo refresh polja.
            '
            ' INFO (Paid/OverDue/Archived) je namerno OVDE, a ne u "ne diraj":
            ' takav status ne govori da li je kupac odobrio fakturu, ali dokazuje
            ' da dokument sigurno nije vise "u slanju". Dok je vracao prazno,
            ' faktura je ostajala SEF_SENDING i startup recovery ju je nalazio
            ' pri svakom pokretanju. Prihvacenom je NE proglasavamo -- za to
            ' postoji zasebna klasa SEF_CLS_ACCEPTED.
            desired = WF_SEF_SENT

        Case Else
            ' SEF_CLS_ERROR i SEF_CLS_UNKNOWN: nemamo upotrebljiv status.
            ' SEF_SENDING mora da izadje iz "salje se" (inace ga startup recovery
            ' nalazi zauvek), SEF_SENT ide u SEF_SYNC_ERROR. Sva ostala stanja
            ' ostaju netaknuta -- ponovni pad ne sme da obori refresh.
            If curState = UCase$(WF_SEF_SENDING) Then
                desired = WF_SEF_UNKNOWN
            ElseIf curState = UCase$(WF_SEF_SENT) Then
                desired = WF_SEF_SYNC_ERROR
            Else
                SEFRefreshTargetState = ""
                Exit Function
            End If

    End Select

    ' Faktura bez zabelezenog stanja: UpdateFakturaSEFState_Row u tom slucaju
    ' preskace validaciju tranzicije, pa je upis bezbedan.
    If Len(curState) = 0 Then
        SEFRefreshTargetState = desired
        Exit Function
    End If

    ' Vec smo u ciljnom stanju -> nema transition-a u samog sebe.
    If curState = UCase$(desired) Then
        SEFRefreshTargetState = ""
        Exit Function
    End If

    ' Sve ostalo odlucuje state machine.
    If IsSEFTransitionAllowed(curState, desired) Then
        SEFRefreshTargetState = desired
        Exit Function
    End If

    ' Direktan skok nije dozvoljen, ali state machine mozda dozvoljava put preko
    ' SEF_SENT. Konkretan slucaj: prethodni refresh je pao (SEF_SYNC_ERROR), a
    ' sada stize finalni status -- SEF_SYNC_ERROR -> SEF_SENT -> SEF_ACCEPTED.
    ' Vracamo PRVI korak; pozivalac (ApplySEFExternalOutcome_Row) zove planer ponovo
    ' sa novim stanjem i dobija drugi. Zato je most ogranicen na jedno stanje --
    ' dva koraka su i gornja granica petlje.
    If IsSEFTransitionAllowed(curState, WF_SEF_SENT) Then
        If IsSEFTransitionAllowed(WF_SEF_SENT, desired) Then
            SEFRefreshTargetState = WF_SEF_SENT
            Exit Function
        End If
    End If

    ' Nema legalnog puta -> ne diramo workflow (npr. SEF_ACCEPTED se ne vraca u
    ' SEF_SENT samo zato sto je spoljni status jos "pending").
    SEFRefreshTargetState = ""

End Function

' Vraca True samo ako je SEF vratio upotrebljiv status.
' AUD-032c: ranije je bezuslovno vracala True, pa je i pad API-ja (SEF_SYNC_ERROR)
' i nepoznat status izgledao kao uspesan refresh - i pozivaocu (frmSEF, recovery)
' i batch brojacima.
Public Function RefreshSEFStatus_TX(ByVal fakturaID As String) As Boolean

    Dim tx As clsTransaction
    Dim sefDocumentId As String
    Dim submissionID As String
    Dim response As clsSEFResponse
    Dim apiStatus As String
    Dim currentState As String
    Dim refreshOk As Boolean
    Dim classification As String
    Dim statusText As String

    On Error GoTo EH
    
    sefDocumentId = GetFakturaSEFDocumentId(fakturaID)
    If Len(Trim$(sefDocumentId)) = 0 Then
        Err.Raise ERR_SEF_STATE, "RefreshSEFStatus_TX", _
            "No SEFDocumentId found for faktura " & fakturaID
    End If
    
    submissionID = GetLastSEFSubmissionID(fakturaID)
    currentState = GetFakturaSEFWorkflowState(fakturaID)
    
    Set response = GetInvoiceStatus(sefDocumentId)
    apiStatus = UCase$(Trim$(response.apiStatus))
    
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot "tblSEFSubmission"
    tx.AddTableSnapshot "tblSEFEventLog"
    
    ' Optional:
    ' keep if you intentionally want latest API snapshot on submission row too
    If Len(Trim$(submissionID)) > 0 Then
        'Call SaveSEFSubmissionResult_Row(submissionID, response)
    End If
    
    ' =========================================================
    ' AUD-032b/c: jedan adapter + jedan planer za sve ishode
    ' =========================================================
    ' Klasa se izvodi iz response flagova (koje klijent postavlja preko istog
    ' klasifikatora) i iz samog apiStatus-a. Pad HTTP/API poziva je uvek
    ' SEF_CLS_ERROR bez obzira sta pise u apiStatus polju.
    If Not response.Success Then
        classification = SEF_CLS_ERROR
    ElseIf response.Accepted Then
        classification = SEF_CLS_ACCEPTED
    ElseIf response.Rejected Then
        classification = SEF_CLS_REJECTED
    Else
        classification = ClassifySEFExternalStatus(apiStatus)
    End If

    refreshOk = IsUsableSEFRefreshClass(classification)

    ' SEFStatus se cuva DOSLOVNO onako kako ga je SEF vratio (kolona je po
    ' definiciji "poslednji spoljni status"); samo prazan status dobija marker.
    statusText = apiStatus
    If Len(Trim$(statusText)) = 0 Then statusText = SEF_STATUS_UNKNOWN

    Select Case classification

        Case SEF_CLS_ACCEPTED

            ApplySEFExternalOutcome_Row fakturaID, classification, statusText, _
                                   response.sefDocumentId, "", ""

            Call AppendSEFEvent_Row( _
                fakturaID:=fakturaID, _
                submissionID:=submissionID, _
                eventType:=SEF_EVT_SYNC_OK, _
                message:="SEF status refreshed: invoice approved by buyer.", _
                details:="ApiStatus=" & statusText & _
                         "; SEFDocumentId=" & response.sefDocumentId)

        Case SEF_CLS_REJECTED

            ApplySEFExternalOutcome_Row fakturaID, classification, statusText, _
                                   response.sefDocumentId, _
                                   response.errorCode, response.errorMessage

            Call AppendSEFEvent_Row( _
                fakturaID:=fakturaID, _
                submissionID:=submissionID, _
                eventType:=SEF_EVT_VALIDATION_FAILED, _
                message:="SEF status refreshed: REJECTED.", _
                details:=response.errorCode & " | " & response.errorMessage)

        Case SEF_CLS_PENDING

            ApplySEFExternalOutcome_Row fakturaID, classification, statusText, _
                                   response.sefDocumentId, "", ""

            Call AppendSEFEvent_Row( _
                fakturaID:=fakturaID, _
                submissionID:=submissionID, _
                eventType:=SEF_EVT_SYNC_OK, _
                message:="SEF status unchanged (pending).", _
                details:=statusText)

        Case SEF_CLS_STORNO, SEF_CLS_TERMINAL

            ApplySEFExternalOutcome_Row fakturaID, classification, statusText, _
                                   response.sefDocumentId, "", ""

            Call AppendSEFEvent_Row( _
                fakturaID:=fakturaID, _
                submissionID:=submissionID, _
                eventType:=SEF_EVT_SYNC_OK, _
                message:="SEF status refreshed: " & statusText & ".", _
                details:=statusText)

        Case SEF_CLS_SEND_FAILED

            ' SEF javlja gresku pri slanju dokumenta -> lokalno SEF_TECH_FAILED,
            ' odakle UI nudi retry (i cancel je dozvoljen na SEF strani).
            ApplySEFExternalOutcome_Row fakturaID, classification, statusText, _
                                   response.sefDocumentId, statusText, _
                                   "SEF reported a document send error; retry or cancel required."

            Call AppendSEFEvent_Row( _
                fakturaID:=fakturaID, _
                submissionID:=submissionID, _
                eventType:=SEF_EVT_SYNC_FAILED, _
                message:="SEF reported a document send error (Mistake).", _
                details:="ApiStatus=" & statusText & _
                         "; LocalState=" & currentState)

        Case SEF_CLS_INFO

            ' Poznat status koji ne nosi odluku kupca (PAID/OVERDUE/ARCHIVED):
            ' belezi se i izvlaci fakturu iz "salje se", ali je NE proglasava
            ' prihvacenom.
            ApplySEFExternalOutcome_Row fakturaID, classification, statusText, _
                                   response.sefDocumentId, "", ""

            Call AppendSEFEvent_Row( _
                fakturaID:=fakturaID, _
                submissionID:=submissionID, _
                eventType:=SEF_EVT_SYNC_OK, _
                message:="SEF status refreshed (informational): " & statusText & ".", _
                details:=statusText)

        Case SEF_CLS_ERROR

            ' SEFStatus nosi ono sto je SEF stvarno vratio (FAILED / HTTP_ERROR /
            ' ERROR), a ne lokalno ime stanja: pri padu nad zaglavljenim
            ' SEF_SENDING workflow ide u SEF_UNKNOWN, pa bi upis "SEF_SYNC_ERROR"
            ' u kolonu spoljnog statusa bio i netacan i zbunjujuci.
            ApplySEFExternalOutcome_Row fakturaID, classification, statusText, _
                                   "", response.errorCode, response.errorMessage

            Call AppendSEFEvent_Row( _
                fakturaID:=fakturaID, _
                submissionID:=submissionID, _
                eventType:=SEF_EVT_SYNC_FAILED, _
                message:="SEF status refresh failed.", _
                details:=response.errorCode & " | " & response.errorMessage & _
                         "; LocalState=" & currentState)

        Case Else

            ' AUD-032b: prazan / nepoznat status NIJE dokaz da je faktura
            ' poslata. Stari kod ga je mapirao na WF_SEF_SENT.
            ApplySEFExternalOutcome_Row fakturaID, classification, statusText, _
                                   response.sefDocumentId, SEF_STATUS_UNKNOWN, _
                                   "SEF returned unknown status; manual review required."

            Call AppendSEFEvent_Row( _
                fakturaID:=fakturaID, _
                submissionID:=submissionID, _
                eventType:=SEF_EVT_SYNC_FAILED, _
                message:="SEF returned unknown status; manual review required.", _
                details:="ApiStatus=" & statusText & _
                         "; LocalState=" & currentState)

    End Select

    Call UpdateSEFLastSyncAt_Row(fakturaID)
    
    tx.CommitTx

    On Error Resume Next

    If response Is Nothing Then

        Monitor_SEF _
            eventType:="SEF_STATUS_REFRESH_FAIL", _
            severity:="ERROR", _
            invoiceLocalId:=fakturaID, _
            businessInvoiceNo:=fakturaID, _
            sefStatus:="UNKNOWN", _
            localStatus:=GetFakturaSEFWorkflowState(fakturaID), _
            sefRequestId:=submissionID, _
            sefInvoiceId:=sefDocumentId, _
            attemptCount:=0, _
            lastHttpCode:="0", _
            lastError:="SEF status response object is Nothing.", _
            nextAction:="RETRY", _
            needsManualReview:=False

    Else

        ' Monitoring ide po ISTOJ klasi po kojoj je i workflow odlucen -- da se
        ' dve liste statusa ne bi razisle (AUD-032b).
        Select Case classification

            Case SEF_CLS_ACCEPTED
                Monitor_SEF _
                    eventType:="SEF_STATUS_ACCEPTED", _
                    severity:="INFO", _
                    invoiceLocalId:=fakturaID, _
                    businessInvoiceNo:=fakturaID, _
                    sefStatus:=statusText, _
                    localStatus:=GetFakturaSEFWorkflowState(fakturaID), _
                    sefRequestId:=submissionID, _
                    sefInvoiceId:=response.sefDocumentId, _
                    attemptCount:=0, _
                    lastHttpCode:=CStr(response.httpStatus), _
                    lastError:="", _
                    nextAction:="WAIT", _
                    needsManualReview:=False

            Case SEF_CLS_REJECTED
                Monitor_SEF _
                    eventType:="SEF_STATUS_REJECTED", _
                    severity:="WARN", _
                    invoiceLocalId:=fakturaID, _
                    businessInvoiceNo:=fakturaID, _
                    sefStatus:=statusText, _
                    localStatus:=GetFakturaSEFWorkflowState(fakturaID), _
                    sefRequestId:=submissionID, _
                    sefInvoiceId:=response.sefDocumentId, _
                    attemptCount:=0, _
                    lastHttpCode:=CStr(response.httpStatus), _
                    lastError:=response.errorCode & " | " & response.errorMessage, _
                    nextAction:="MANUAL_REVIEW", _
                    needsManualReview:=True

            Case SEF_CLS_PENDING
                Monitor_SEF _
                    eventType:="SEF_STATUS_PENDING", _
                    severity:="INFO", _
                    invoiceLocalId:=fakturaID, _
                    businessInvoiceNo:=fakturaID, _
                    sefStatus:=statusText, _
                    localStatus:=GetFakturaSEFWorkflowState(fakturaID), _
                    sefRequestId:=submissionID, _
                    sefInvoiceId:=response.sefDocumentId, _
                    attemptCount:=0, _
                    lastHttpCode:=CStr(response.httpStatus), _
                    lastError:="", _
                    nextAction:="WAIT", _
                    needsManualReview:=False

            Case SEF_CLS_SEND_FAILED
                ' Nije terminalno i nije obicna greska osvezavanja: dokument je
                ' na SEF-u obelezen kao neuspelo slanje -> retry/cancel.
                Monitor_SEF _
                    eventType:="SEF_STATUS_SEND_FAILED", _
                    severity:="WARN", _
                    invoiceLocalId:=fakturaID, _
                    businessInvoiceNo:=fakturaID, _
                    sefStatus:=statusText, _
                    localStatus:=GetFakturaSEFWorkflowState(fakturaID), _
                    sefRequestId:=submissionID, _
                    sefInvoiceId:=response.sefDocumentId, _
                    attemptCount:=0, _
                    lastHttpCode:=CStr(response.httpStatus), _
                    lastError:="SEF reported a document send error (Mistake).", _
                    nextAction:="RETRY", _
                    needsManualReview:=False

            Case SEF_CLS_INFO
                ' Odvojeno od TERMINAL: PAID/OVERDUE/ARCHIVED nisu kraj zivota
                ' dokumenta, pa ih ni telemetrija ne sme tako prijavljivati.
                Monitor_SEF _
                    eventType:="SEF_STATUS_INFO", _
                    severity:="INFO", _
                    invoiceLocalId:=fakturaID, _
                    businessInvoiceNo:=fakturaID, _
                    sefStatus:=statusText, _
                    localStatus:=GetFakturaSEFWorkflowState(fakturaID), _
                    sefRequestId:=submissionID, _
                    sefInvoiceId:=response.sefDocumentId, _
                    attemptCount:=0, _
                    lastHttpCode:=CStr(response.httpStatus), _
                    lastError:="", _
                    nextAction:="WAIT", _
                    needsManualReview:=False

            Case SEF_CLS_STORNO, SEF_CLS_TERMINAL
                Monitor_SEF _
                    eventType:="SEF_STATUS_TERMINAL", _
                    severity:="INFO", _
                    invoiceLocalId:=fakturaID, _
                    businessInvoiceNo:=fakturaID, _
                    sefStatus:=statusText, _
                    localStatus:=GetFakturaSEFWorkflowState(fakturaID), _
                    sefRequestId:=submissionID, _
                    sefInvoiceId:=response.sefDocumentId, _
                    attemptCount:=0, _
                    lastHttpCode:=CStr(response.httpStatus), _
                    lastError:="", _
                    nextAction:="WAIT", _
                    needsManualReview:=False

            Case SEF_CLS_ERROR
                Monitor_SEF _
                    eventType:="SEF_STATUS_REFRESH_FAIL", _
                    severity:="ERROR", _
                    invoiceLocalId:=fakturaID, _
                    businessInvoiceNo:=fakturaID, _
                    sefStatus:=statusText, _
                    localStatus:=GetFakturaSEFWorkflowState(fakturaID), _
                    sefRequestId:=submissionID, _
                    sefInvoiceId:=response.sefDocumentId, _
                    attemptCount:=0, _
                    lastHttpCode:=CStr(response.httpStatus), _
                    lastError:=response.errorCode & " | " & response.errorMessage, _
                    nextAction:="RETRY", _
                    needsManualReview:=False

            Case Else
                ' AUD-032b: nepoznat status ide kao WARN + rucna provera,
                ' ne kao obican INFO "status update".
                Monitor_SEF _
                    eventType:="SEF_STATUS_UNKNOWN", _
                    severity:="WARN", _
                    invoiceLocalId:=fakturaID, _
                    businessInvoiceNo:=fakturaID, _
                    sefStatus:=statusText, _
                    localStatus:=GetFakturaSEFWorkflowState(fakturaID), _
                    sefRequestId:=submissionID, _
                    sefInvoiceId:=response.sefDocumentId, _
                    attemptCount:=0, _
                    lastHttpCode:=CStr(response.httpStatus), _
                    lastError:="SEF returned unknown status; manual review required.", _
                    nextAction:="MANUAL_REVIEW", _
                    needsManualReview:=True

        End Select

    End If

    On Error GoTo 0

    RefreshSEFStatus_TX = refreshOk
    Exit Function

EH:
    Dim errNo As Long
    Dim errDesc As String
    Dim errSrc As String

    errNo = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    LogErr "RefreshSEFStatus_TX"
    On Error Resume Next


    Monitor_Error _
        moduleName:="modSEFStatusSync", _
        procedureName:="RefreshSEFStatus_TX", _
        entityType:="Faktura", _
        entityID:=fakturaID, _
        correlationId:=fakturaID, _
        errorNumber:=errNo, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_SEF _
        eventType:="SEF_STATUS_REFRESH_EXCEPTION", _
        severity:="ERROR", _
        invoiceLocalId:=fakturaID, _
        businessInvoiceNo:=fakturaID, _
        sefStatus:="UNKNOWN", _
        localStatus:=currentState, _
        sefRequestId:=submissionID, _
        sefInvoiceId:=sefDocumentId, _
        attemptCount:=0, _
        lastHttpCode:="vba-exception", _
        lastError:=errDesc, _
        nextAction:="RETRY", _
        needsManualReview:=False

    If Not tx Is Nothing Then tx.RollbackTx

    On Error GoTo 0
    Err.Raise errNo, "RefreshSEFStatus_TX", errDesc
End Function

' AUD-032b/c: JEDINI upisivac lokalnog stanja iz SPOLJNOG SEF ishoda.
'
' Koriste ga i refresh (RefreshSEFStatus_TX) i akcije koje spoljni ishod dobiju
' direktno od SEF-a (storno/cancel u modSEFService) -- da bi "sta spoljni status
' znaci za lokalni workflow" bilo odluceno na JEDNOM mestu. Pre toga je storno
' menjao samo SEFStatus, pa je nastajala trajna kontradikcija
' (SEFWorkflowState = SEF_SENT/SEF_ACCEPTED uz SEFStatus = STORNO).
'
' Ciljno stanje odredjuje cist planer SEFRefreshTargetState (koji sam proverava
' dozvoljenost tranzicije), pa ovaj Sub nema sopstvenu logiku o tome "sme li" --
' samo primenjuje plan. Prazan plan = ne diraj workflow, upisi samo refresh polja.
'
' Petlja od najvise dva koraka pokriva dvokorak koji state machine zahteva:
' SEF_SYNC_ERROR -> SEF_SENT -> SEF_ACCEPTED/SEF_REJECTED (direktan skok iz
' SEF_SYNC_ERROR u finalno stanje nije dozvoljen). Posle drugog koraka planer
' vrati prazno, pa petlja staje sama.
Public Sub ApplySEFExternalOutcome_Row(ByVal fakturaID As String, _
                                   ByVal classification As String, _
                                   ByVal sefStatus As String, _
                                   Optional ByVal sefDocumentId As String = "", _
                                   Optional ByVal errorCode As String = "", _
                                   Optional ByVal errorMessage As String = "")
    On Error GoTo EH

    Const MAX_HOPS As Long = 2

    Dim currentState As String
    Dim targetState As String
    Dim hop As Long
    Dim wroteState As Boolean

    currentState = UCase$(Trim$(GetFakturaSEFWorkflowState(fakturaID)))

    For hop = 1 To MAX_HOPS

        targetState = SEFRefreshTargetState(currentState, classification)
        If Len(targetState) = 0 Then Exit For

        UpdateFakturaSEFState_Row _
            fakturaID:=fakturaID, _
            newState:=targetState, _
            sefStatus:=sefStatus, _
            sefDocumentId:=sefDocumentId, _
            errorCode:=errorCode, _
            errorMessage:=errorMessage

        wroteState = True
        currentState = UCase$(Trim$(targetState))

    Next hop

    If Not wroteState Then
        UpdateFakturaSEFRefreshFields_Row _
            fakturaID:=fakturaID, _
            sefStatus:=sefStatus, _
            sefDocumentId:=sefDocumentId, _
            errorCode:=errorCode, _
            errorMessage:=errorMessage
    End If

    Exit Sub

EH:
    ' AUD-054: greska se hvata PRE LogErr-a; LogErr moze da resetuje Err objekat,
    ' pa bi "Err.Raise Err.Number" izbacilo Err 0 i sakrilo pravi uzrok.
    Dim errNum As Long
    Dim errDesc As String

    errNum = Err.Number
    errDesc = Err.description

    LogErr "modSEFStatusSync.ApplySEFExternalOutcome_Row"
    On Error Resume Next
    On Error GoTo 0

    If errNum = 0 Then errNum = ERR_SEF_STATE

    Err.Raise errNum, "modSEFStatusSync.ApplySEFExternalOutcome_Row", errDesc
End Sub

' Batch prolaz preskace fakture ciji je spoljni status terminalan (storno,
' otkazano, obrisano, greska dokumenta) -- za njih nema sta da se osvezava.
' Ide preko istog klasifikatora, pa se spisak ne moze razici sa adapterom.
Private Function IsTerminalExternalRefreshStatus(ByVal sefStatus As String) As Boolean
    Select Case ClassifySEFExternalStatus(sefStatus)
        Case SEF_CLS_STORNO, SEF_CLS_TERMINAL
            IsTerminalExternalRefreshStatus = True
        Case Else
            IsTerminalExternalRefreshStatus = False
    End Select
End Function

' AUD-032f: vraca sazetak prolaza (Scanned/Refreshed/Unresolved/SkippedTerminal/Failed)
' da bi frmSEF mogao da ga pokaze operateru, umesto tihog "Pending fakture osvezene".
' Statement-pozivi (`Call RefreshPendingOutboundInvoices_TX`) rade i dalje.
Public Function RefreshPendingOutboundInvoices_TX() As String

    On Error GoTo EH

    Const SRC As String = "modSEFStatusSync.RefreshPendingOutboundInvoices_TX"
    
    On Error Resume Next
    Monitor_Event _
        eventType:="SEF_REFRESH_PENDING_START", _
        severity:="INFO", _
        message:="Started pending outbound SEF refresh", _
        userId:="Operator", _
        moduleName:="modSEFStatusSync", _
        procedureName:="RefreshPendingOutboundInvoices_TX", _
        entityType:="SEF", _
        entityID:="PendingOutbound", _
        correlationId:="SEF-PENDING-REFRESH"
    On Error GoTo EH

    Dim i As Long
    Dim fakturaID As String
    Dim workflowState As String
    Dim sefStatus As String
    Dim scannedCount As Long
    Dim refreshedCount As Long
    Dim unresolvedCount As Long
    Dim skippedTerminalCount As Long
    Dim failedCount As Long
    Dim itemOk As Boolean
    Dim summaryText As String

    Dim data As Variant
    data = GetTableData(TBL_FAKTURE)

    If IsEmpty(data) Then GoTo SummaryExit

    Dim colFakturaID As Long
    Dim colWorkflow As Long
    Dim colSEFStatus As Long

    colFakturaID = RequireColumnIndex(TBL_FAKTURE, "FakturaID", SRC)
    colWorkflow = RequireColumnIndex(TBL_FAKTURE, "SEFWorkflowState", SRC)
    colSEFStatus = RequireColumnIndex(TBL_FAKTURE, "SEFStatus", SRC)

    For i = 1 To UBound(data, 1)

        fakturaID = Trim$(CStr(data(i, colFakturaID)))
        workflowState = UCase$(Trim$(CStr(data(i, colWorkflow))))
        sefStatus = UCase$(Trim$(CStr(data(i, colSEFStatus))))

        Select Case workflowState

            Case UCase$(WF_SEF_SENT), UCase$(WF_SEF_SYNC_ERROR)
                scannedCount = scannedCount + 1
                
                If IsTerminalExternalRefreshStatus(sefStatus) Then
                    skippedTerminalCount = skippedTerminalCount + 1
                    GoTo NextInvoice
                End If

                On Error Resume Next
                itemOk = False
                itemOk = RefreshSEFStatus_TX(fakturaID)

                If Err.Number <> 0 Then
                    failedCount = failedCount + 1

                    Dim itemErrNo As Long
                    Dim itemErrDesc As String
                    Dim itemErrSrc As String

                    itemErrNo = Err.Number
                    itemErrDesc = Err.description
                    itemErrSrc = Err.SOURCE

                    LogErr SRC & ".Invoice." & fakturaID

                    Monitor_Error _
                        moduleName:="modSEFStatusSync", _
                        procedureName:="RefreshPendingOutboundInvoices_TX.Invoice", _
                        entityType:="Faktura", _
                        entityID:=fakturaID, _
                        correlationId:=fakturaID, _
                        errorNumber:=itemErrNo, _
                        errorDescription:=itemErrDesc, _
                        errorSource:=itemErrSrc

                    Monitor_SEF _
                        eventType:="SEF_PENDING_REFRESH_INVOICE_FAIL", _
                        severity:="ERROR", _
                        invoiceLocalId:=fakturaID, _
                        businessInvoiceNo:=fakturaID, _
                        sefStatus:="UNKNOWN", _
                        localStatus:=workflowState, _
                        sefRequestId:=GetLastSEFSubmissionID(fakturaID), _
                        sefInvoiceId:=GetFakturaSEFDocumentId(fakturaID), _
                        attemptCount:=0, _
                        lastHttpCode:="vba-exception", _
                        lastError:=itemErrDesc, _
                        nextAction:="RETRY", _
                        needsManualReview:=False

                    Err.Clear
                ElseIf itemOk Then
                    refreshedCount = refreshedCount + 1
                Else
                    ' AUD-032c: poziv nije pukao, ali SEF nije vratio upotrebljiv
                    ' status (SYNC_ERROR ili nepoznat status). To NIJE osvezena
                    ' faktura i ne sme da se broji kao takva.
                    unresolvedCount = unresolvedCount + 1
                End If

                On Error GoTo EH

                Application.Wait Now + TimeSerial(0, 0, 2)

        End Select

NextInvoice:
    Next i

SummaryExit:
    summaryText = _
        "Scanned=" & scannedCount & _
        "; Refreshed=" & refreshedCount & _
        "; Unresolved=" & unresolvedCount & _
        "; SkippedTerminal=" & skippedTerminalCount & _
        "; Failed=" & failedCount

    RefreshPendingOutboundInvoices_TX = summaryText

    On Error Resume Next
    Monitor_Event _
        eventType:="SEF_REFRESH_PENDING_SUMMARY", _
        severity:="INFO", _
        message:="Pending SEF refresh completed. " & summaryText, _
        userId:="Operator", _
        moduleName:="modSEFStatusSync", _
        procedureName:="RefreshPendingOutboundInvoices_TX", _
        entityType:="SEF", _
        entityID:="PendingOutbound", _
        correlationId:="SEF-PENDING-REFRESH"
    On Error GoTo 0

    Exit Function

EH:
    Dim errNo As Long
    Dim errDesc As String
    Dim errSrc As String

    errNo = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    LogErr SRC
    On Error Resume Next


    Monitor_Error _
        moduleName:="modSEFStatusSync", _
        procedureName:="RefreshPendingOutboundInvoices_TX", _
        entityType:="SEF", _
        entityID:="PendingOutbound", _
        correlationId:="SEF-PENDING-REFRESH", _
        errorNumber:=errNo, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="SEF_REFRESH_PENDING_FAIL", _
        severity:="CRITICAL", _
        message:=errDesc, _
        userId:="Operator", _
        moduleName:="modSEFStatusSync", _
        procedureName:="RefreshPendingOutboundInvoices_TX", _
        entityType:="SEF", _
        entityID:="PendingOutbound", _
        correlationId:="SEF-PENDING-REFRESH"

    On Error GoTo 0
    Err.Raise errNo, SRC, errDesc
End Function


' AUD-032e: dev makroi ispod su Private -- vise se ne vide u Alt+F8 listi.
' Pokrecu se iz VBE-a (kursor u proceduri -> F5). Gadjaju ZIVI SEF status API.
Private Sub Test2_RefreshSEFStatus_TX()

    On Error GoTo EH
    
    Dim ok As Boolean
    Dim fakturaID As String
    
    fakturaID = "FAK-00008"
    
    ok = RefreshSEFStatus_TX(fakturaID)
    
    Debug.Print "Refresh OK: "; ok
    Debug.Print "WorkflowState: "; GetFakturaSEFWorkflowState(fakturaID)
    Debug.Print "SEFDocumentId: "; GetFakturaSEFDocumentId(fakturaID)
    Debug.Print "LastSubmissionID: "; GetLastSEFSubmissionID(fakturaID)
    Debug.Print "SEFStatus: "; LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFStatus")
    Debug.Print "SEFLastErrorCode: "; LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFLastErrorCode")
    Debug.Print "SEFLastErrorMessage: "; LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFLastErrorMessage")
    
    Exit Sub

EH:
    Debug.Print "ERR " & Err.Number & " - " & Err.description
End Sub

Private Sub Test1_RefreshSEFStatus_TX()

    On Error GoTo EH
    
    Dim ok As Boolean
    
    ok = RefreshSEFStatus_TX("FAK-00008")
    
    Debug.Print "Refresh OK: "; ok
    Debug.Print "WorkflowState: "; GetFakturaSEFWorkflowState("FAK-00008")
    Debug.Print "SEFDocumentId: "; GetFakturaSEFDocumentId("FAK-00008")
    Debug.Print "LastSubmissionID: "; GetLastSEFSubmissionID("FAK-00008")
    Debug.Print "SEFStatus: "; LookupValue(TBL_FAKTURE, "FakturaID", "FAK-00008", "SEFStatus")
    
    Exit Sub

EH:
    Debug.Print "ERR " & Err.Number & " - " & Err.description
End Sub
