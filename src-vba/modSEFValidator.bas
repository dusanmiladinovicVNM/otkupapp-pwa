Attribute VB_Name = "modSEFValidator"
Option Explicit

Public Sub ValidateAllowedTransition(ByVal oldState As String, ByVal newState As String)
    
    Select Case oldState
        
        Case WF_LOCAL_DRAFT
            If newState <> WF_LOCAL_FINALIZED Then GoTo InvalidTransition
        
        Case WF_LOCAL_FINALIZED
            If newState <> WF_SEF_READY Then GoTo InvalidTransition
        
        Case WF_SEF_READY
            If newState <> WF_SEF_SENDING Then GoTo InvalidTransition
            
        Case WF_SEF_SENDING
            Select Case newState
                Case WF_SEF_SENT, WF_SEF_ACCEPTED, WF_SEF_REJECTED, WF_SEF_TECH_FAILED, WF_SEF_UNKNOWN
            Case Else
                GoTo InvalidTransition
        End Select
        
        ' AUD-032b: SEF_SENT -> SEF_TECH_FAILED je dodat zbog zvanicnog statusa
        ' "Mistake" (greska prilikom slanja). NORMALNA sekvenca je: uspesan
        ' submit -> lokalno SEF_SENT -> refresh vrati MISTAKE. Bez ove tranzicije
        ' faktura je ostajala SEF_SENT ("uspesno poslata") iako SEF tvrdi da
        ' slanje nije uspelo, a batch ju je osvezavao u nedogled.
        Case WF_SEF_SENT
            Select Case newState
                Case WF_SEF_ACCEPTED, WF_SEF_REJECTED, WF_SEF_SYNC_ERROR, _
                     WF_SEF_STORNO, WF_SEF_TECH_FAILED
                Case Else
                    GoTo InvalidTransition
            End Select

        Case WF_SEF_TECH_FAILED
            If newState <> WF_SEF_READY Then GoTo InvalidTransition

        ' Isti razlog: refresh iz SEF_SYNC_ERROR takodje moze da vrati MISTAKE.
        Case WF_SEF_SYNC_ERROR
            Select Case newState
                Case WF_SEF_SENT, WF_SEF_TECH_FAILED
                Case Else
                    GoTo InvalidTransition
            End Select
        
        Case WF_SEF_ACCEPTED
            If newState <> WF_SEF_STORNO Then GoTo InvalidTransition

        Case WF_SEF_REJECTED
            If newState <> WF_SEF_READY Then GoTo InvalidTransition
            
        Case WF_SEF_UNKNOWN
            ' AUD-032b: SEF_UNKNOWN je stanje "SEF je vratio nepoznat status,
            ' potrebna rucna provera". Do RF-22 je bilo slepo crevo (ulazak iz
            ' SEF_SENDING je dozvoljen, a izlaz nije postojao), pa je faktura
            ' ostajala zauvek zaglavljena. Operater sada moze da je osvezi.
            Select Case newState
                Case WF_SEF_SENT, WF_SEF_ACCEPTED, WF_SEF_REJECTED, WF_SEF_TECH_FAILED
                Case Else
                    GoTo InvalidTransition
            End Select

        Case WF_SEF_STORNO
            GoTo InvalidTransition

        Case Else
            Err.Raise ERR_SEF_STATE, "ValidateAllowedTransition", _
                "Unknown current workflow state: " & oldState
    End Select
    
    Exit Sub

InvalidTransition:
    Err.Raise ERR_SEF_STATE, "ValidateAllowedTransition", _
        "Illegal SEF state transition: " & oldState & " -> " & newState
End Sub

' AUD-032c: ne-bacajuci oblik iste odluke. Sluzi da pozivalac (planer tranzicije
' u modSEFStatusSync) moze da PITA state machine umesto da pretpostavlja -- pa ne
' moze da predlozi tranziciju koja ce dole puknuti i oboriti refresh. Jedini
' izvor istine ostaje ValidateAllowedTransition.
Public Function IsSEFTransitionAllowed(ByVal oldState As String, _
                                       ByVal newState As String) As Boolean
    On Error GoTo NotAllowed

    ValidateAllowedTransition oldState, newState

    IsSEFTransitionAllowed = True
    Exit Function

NotAllowed:
    IsSEFTransitionAllowed = False
End Function

Public Sub ValidateFakturaForSEF(ByVal fakturaID As String)
    On Error GoTo EH

    Const SRC As String = "modSEFValidator.ValidateFakturaForSEF"

    Dim fakture As Variant
    Dim i As Long

    Dim colFakturaID As Long
    Dim colKupacID As Long
    Dim colWorkflow As Long
    Dim colBrojFakture As Long
    Dim colIznos As Long
    Dim colStornirano As Long
    Dim colOsiroceno As Long

    Dim found As Boolean
    Dim kupacID As String
    Dim workflowState As String
    Dim brojFakture As String
    Dim iznosRaw As String
    Dim iznosValue As Double
    Dim storniranoRaw As String
    Dim osirocenoRaw As String

    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "FakturaID is required."
    End If

    fakture = GetTableData(TBL_FAKTURE)

    If IsEmpty(fakture) Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "TBL_FAKTURE is empty."
    End If

    colFakturaID = RequireColumnIndex(TBL_FAKTURE, "FakturaID", SRC)
    colKupacID = RequireColumnIndex(TBL_FAKTURE, "KupacID", SRC)
    colWorkflow = RequireColumnIndex(TBL_FAKTURE, "SEFWorkflowState", SRC)
    colBrojFakture = RequireColumnIndex(TBL_FAKTURE, "BrojFakture", SRC)
    colIznos = RequireColumnIndex(TBL_FAKTURE, "Iznos", SRC)

    ' AUD-031a: cancellation / orphan markers gate the tax send path, so they
    ' are REQUIRED (fail-closed). A missing column must raise here rather than
    ' silently letting a stornirana/orphaned faktura become sendable again.
    colStornirano = RequireColumnIndex(TBL_FAKTURE, COL_STORNIRANO, SRC)
    colOsiroceno = RequireColumnIndex(TBL_FAKTURE, COL_OSIROCENO_OD, SRC)

    For i = 1 To UBound(fakture, 1)
        If CStr(fakture(i, colFakturaID)) = fakturaID Then
            found = True
            kupacID = Trim$(CStr(fakture(i, colKupacID)))
            workflowState = Trim$(CStr(fakture(i, colWorkflow)))
            brojFakture = Trim$(CStr(fakture(i, colBrojFakture)))
            iznosRaw = Trim$(CStr(fakture(i, colIznos)))
            storniranoRaw = Trim$(CStr(fakture(i, colStornirano)))
            osirocenoRaw = Trim$(CStr(fakture(i, colOsiroceno)))
            Exit For
        End If
    Next i

    If Not found Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "Faktura not found: " & fakturaID
    End If

    ' AUD-031a: a stornirana faktura must never be sent to SEF.
    If UCase$(storniranoRaw) = "DA" Then
        Err.Raise ERR_SEF_STATE, SRC, _
                  "Faktura is stornirana (cancelled) and cannot be sent to SEF: " & fakturaID
    End If

    ' AUD-031: an orphaned faktura (its invoiced prijemnica was storno-ed, so
    ' some/all lines carry OsirocenoOd) is arithmetically inconsistent and must
    ' not go to the tax authority. Block here so the operator gets a clear
    ' message instead of a downstream total-mismatch during UBL build.
    If Len(osirocenoRaw) > 0 Then
        Err.Raise ERR_SEF_STATE, SRC, _
                  "Faktura has orphaned (osiroceno) items and cannot be sent to SEF: " & fakturaID
    End If

    If Len(kupacID) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, _
                  "KupacID is missing for faktura " & fakturaID
    End If

    If Len(brojFakture) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, _
                  "BrojFakture is missing for faktura " & fakturaID
    End If

    If Not TryParseDouble(iznosRaw, iznosValue) Then
        Err.Raise ERR_SEF_VALIDATION, SRC, _
                  "UkupanIznos is not numeric for faktura " & fakturaID
    End If

    If iznosValue <= 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, _
                  "UkupanIznos must be > 0 for faktura " & fakturaID
    End If

    Select Case workflowState

        Case WF_LOCAL_FINALIZED, WF_SEF_READY, WF_SEF_TECH_FAILED
            ' allowed

        Case WF_SEF_ACCEPTED
            Err.Raise ERR_SEF_STATE, SRC, _
                      "Faktura already accepted on SEF."

        Case WF_SEF_SENDING
            Err.Raise ERR_SEF_STATE, SRC, _
                      "Faktura is already in SEF_SENDING state."

        Case WF_SEF_SENT
            Err.Raise ERR_SEF_STATE, SRC, _
                      "Faktura already sent. Refresh status first."

        Case WF_SEF_REJECTED
            Err.Raise ERR_SEF_STATE, SRC, _
                      "Faktura was rejected. Correction flow required."

        Case Else
            Err.Raise ERR_SEF_STATE, SRC, _
                      "Faktura is not in a sendable state: " & workflowState
    End Select

    ' AUD-032b: dokument koji vec postoji na SEF-u (npr. status MISTAKE, ili
    ' CANCELLED posle otkazivanja) ne sme da se posalje ponovo. Provera ide PRE
    ' generickog duplicate guard-a da bi operater dobio poruku sa razlogom i
    ' sledecim korakom, a ne suvo "already has a successful SEF submission".
    ' Isti spisak koristi frmSEF za paljenje dugmeta, pa forma ne nudi akciju
    ' koju kapija odbija.
    Dim sefStatusText As String
    Dim sefDocIdText As String

    sefStatusText = GetFakturaSEFStatusText(fakturaID, SRC)
    sefDocIdText = GetFakturaSEFDocumentId(fakturaID)

    If Not CanSendSEFInvoice(workflowState, sefStatusText, sefDocIdText) Then
        Err.Raise ERR_SEF_STATE, SRC, _
                  "Faktura cannot be sent while a SEF document exists (SEFStatus=" & _
                  sefStatusText & "; SEFDocumentId=" & sefDocIdText & "). " & _
                  SEFSendBlockedNextStep(sefStatusText)
    End If

    If HasSuccessfulSEFSubmission(fakturaID) Then
        Err.Raise ERR_SEF_DUPLICATE, SRC, _
                  "Faktura already has a successful SEF submission."
    End If

    ValidateFakturaHasStavke fakturaID
    ValidateKupacForSEF kupacID
    ValidateSEFConfig

    Exit Sub

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.description
End Sub


Private Sub ValidateFakturaHasStavke(ByVal fakturaID As String)
    On Error GoTo EH

    Const SRC As String = "modSEFValidator.ValidateFakturaHasStavke"

    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "FakturaID is required."
    End If

    RequireColumnIndex TBL_FAKTURA_STAVKE, "FakturaID", SRC

    Dim rowsFound As Collection
    Set rowsFound = FindRows(TBL_FAKTURA_STAVKE, "FakturaID", fakturaID)

    If rowsFound Is Nothing Or rowsFound.count = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, _
                  "Faktura has no stavke: " & fakturaID
    End If

    Exit Sub

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.description
End Sub

Public Sub ValidateSEFPayload(ByVal payload As String)
    On Error GoTo EH

    Const SRC As String = "modSEFValidator.ValidateSEFPayload"

    If Len(Trim$(payload)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "SEF payload is empty."
    End If

    If InStr(1, payload, "InvoiceNumber", vbTextCompare) = 0 _
       And InStr(1, payload, "<cbc:ID>", vbTextCompare) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, _
                  "SEF payload does not contain an invoice identifier."
    End If

    Exit Sub

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.description
End Sub
Private Sub ValidateKupacForSEF(ByVal kupacID As String)
    On Error GoTo EH

    Const SRC As String = "modSEFValidator.ValidateKupacForSEF"

    If Len(Trim$(kupacID)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "KupacID is required."
    End If

    RequireColumnIndex TBL_KUPCI, "KupacID", SRC
    RequireColumnIndex TBL_KUPCI, "Naziv", SRC
    RequireColumnIndex TBL_KUPCI, "PIB", SRC

    Dim naziv As Variant
    Dim pib As Variant

    naziv = LookupValue(TBL_KUPCI, "KupacID", kupacID, "Naziv")
    pib = LookupValue(TBL_KUPCI, "KupacID", kupacID, "PIB")

    If IsEmpty(naziv) Or IsNull(naziv) Or Len(Trim$(CStr(naziv))) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, _
                  "Kupac naziv is missing for kupac " & kupacID
    End If

    If IsEmpty(pib) Or IsNull(pib) Or Len(Trim$(CStr(pib))) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, _
                  "Kupac PIB is missing for kupac " & kupacID
    End If

    Exit Sub

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.description
End Sub

Private Sub ValidateSEFConfig()
    On Error GoTo EH

    Const SRC As String = "modSEFValidator.ValidateSEFConfig"

    Dim baseUrl As String
    Dim apiKey As String

    baseUrl = Trim$(GetConfigValue("SEF_BASE_URL"))
    apiKey = Trim$(GetConfigValue("SEF_API_KEY"))

    If Len(baseUrl) = 0 Then
        Err.Raise ERR_SEF_CONFIG, SRC, _
                  "SEF_BASE_URL missing in tblSEFConfig."
    End If

    If Len(apiKey) = 0 Then
        Err.Raise ERR_SEF_CONFIG, SRC, _
                  "SEF_API_KEY missing in tblSEFConfig."
    End If

    If LCase$(Left$(baseUrl, 8)) <> "https://" Then
        Err.Raise ERR_SEF_CONFIG, SRC, _
              "SEF_BASE_URL must start with https://. Plain HTTP is not allowed for SEF."
    End If

    Exit Sub

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.description
End Sub

Private Function GetFakturaSEFStatusText(ByVal fakturaID As String, _
                                         ByVal sourceName As String) As String
    On Error GoTo EH

    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_STATE, sourceName, "FakturaID is required."
    End If

    RequireColumnIndex TBL_FAKTURE, "FakturaID", sourceName
    RequireColumnIndex TBL_FAKTURE, "SEFStatus", sourceName

    Dim v As Variant
    v = LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFStatus")

    If IsEmpty(v) Or IsNull(v) Then
        GetFakturaSEFStatusText = ""
    Else
        GetFakturaSEFStatusText = UCase$(Trim$(CStr(v)))
    End If

    Exit Function

EH:
    LogErr sourceName
    Err.Raise Err.Number, sourceName, Err.description
End Function

' =========================================================
' CAPABILITY UGOVOR (AUD-032b)
'
' Jedan spisak po akciji, koji koriste I validator (fail-closed kapija) I forma
' (enable dugmeta). Ranije su to bile dve rucne liste na dva mesta -- forma je
' nudila dugme za statuse koje validator odbija, i obrnuto.
'
' Cancel: zvanicno uputstvo dozvoljava otkazivanje dokumenta u statusima
' Draft, New i Mistake ("greska prilikom slanja"). "ERROR" je nas interni
' marker iz ranijih verzija i zadrzan je zbog zatecenih redova.
' =========================================================
Public Function CanCancelSEFStatus(ByVal sefStatus As String) As Boolean
    Select Case UCase$(Trim$(sefStatus))
        Case "DRAFT", "NEW", "MISTAKE", "ERROR"
            CanCancelSEFStatus = True
        Case Else
            CanCancelSEFStatus = False
    End Select
End Function

' Sme li se faktura (ponovo) poslati na SEF -- JEDAN spisak za formu (enable
' dugmeta) i za fail-closed kapiju u ValidateFakturaForSEF.
'
' Uslov je dvodelan, jer sam workflow nije dovoljan:
'   1) lokalno stanje mora biti sendable (LOCAL_FINALIZED / SEF_READY /
'      SEF_TECH_FAILED), i
'   2) na SEF-u NE sme postojati dokument koji bi novo slanje pretvorilo u
'      duplikat.
'
' AUD-032b, tacka 2 -- poslovna odluka za "Mistake":
' MISTAKE fakturu vodimo u SEF_TECH_FAILED (jer NIJE poslata), ali retry NE
' nudimo. Razlog je proverljiv u kodu: prvi uspesan submit upisuje
' SubmissionStatus = SENT, status refresh namerno ne dira submission red, pa
' HasSuccessfulSEFSubmission (fail-closed duplicate guard, AUD-031d) svako
' sledece slanje odbija kao duplikat -- a ShouldReuseLastSubmission trazi
' FAILED/CREATED, pa ne bi ni reuse-ovao isti requestId. Ranije je forma palila
' "Retry slanje na SEF" koji je kapija svakako odbijala.
' Da li SEF uopste prihvata ponovni POST istog requestId za dokument u statusu
' Mistake NIJE proverivo staticki (isti zakljucak kao FM-0037 #2) -- dok se ne
' potvrdi na demo SEF-u, putanja za MISTAKE je Cancel + rucna provera.
' Isto vazi posle uspesnog cancel-a: dokument na SEF-u postoji (CANCELLED), pa
' se ista faktura ne nudi za ponovno slanje.
'
' REJECTED je namerno DOZVOLJEN: to je postojeci resubmit tok
' (PrepareRejectedInvoiceForResubmit vraca workflow u SEF_READY, a SEFStatus
' ostaje REJECTED).
Public Function CanSendSEFInvoice(ByVal workflowState As String, _
                                  ByVal sefStatus As String, _
                                  ByVal sefDocumentId As String) As Boolean

    Select Case UCase$(Trim$(workflowState))
        Case UCase$(WF_LOCAL_FINALIZED), UCase$(WF_SEF_READY), UCase$(WF_SEF_TECH_FAILED)
            ' sendable lokalno stanje
        Case Else
            CanSendSEFInvoice = False
            Exit Function
    End Select

    ' PRVO trajna cinjenica, tek onda status. `SEFStatus` je PROMENLJIV: svaki
    ' neuspeo refresh ga prepise u FAILED/HTTP_ERROR (klasa UNKNOWN), pa bi
    ' provera samo po statusu ponovo upalila "Retry" nad fakturom koja ima ziv
    ' dokument na SEF-u -- i klik bi pao na duplicate guard. `SEFDocumentId`
    ' postoji samo ako je SEF stvarno primio dokument i ne brise se pri padu
    ' refresh-a (jedino ga `ClearFakturaLastSubmission_Row` cisti, u resubmit
    ' toku odbijene fakture).
    If Len(Trim$(sefDocumentId)) > 0 Then
        CanSendSEFInvoice = False
        Exit Function
    End If

    ' Rezervna odbrana za slucaj da je dokument nastao a docId se izgubio:
    ' sam status i dalje dokazuje da dokument zivi na SEF-u.
    Select Case ClassifySEFExternalStatus(sefStatus)

        Case SEF_CLS_ACCEPTED, SEF_CLS_PENDING, SEF_CLS_INFO, _
             SEF_CLS_TERMINAL, SEF_CLS_SEND_FAILED
            CanSendSEFInvoice = False

        Case SEF_CLS_REJECTED
            ' Odbijena faktura se salje ponovo SAMO kroz pripremljen tok
            ' (`PrepareRejectedInvoiceForResubmit` -> SEF_READY, obrisan
            ' SEFDocumentId i submission link). Bilo koje drugo stanje sa
            ' statusom REJECTED je nesredjeno -> rucna provera.
            CanSendSEFInvoice = (UCase$(Trim$(workflowState)) = UCase$(WF_SEF_READY))

        Case Else
            ' ERROR / UNKNOWN / prazno + nema SEFDocumentId = slanje nije ni
            ' stiglo do SEF-a (obican tehnicki pad). Retry je ispravan.
            CanSendSEFInvoice = True

    End Select

End Function

' Sledeci korak za operatera kad je slanje blokirano. Poruka se izvodi iz ISTIH
' capability funkcija koje odlucuju sta je dozvoljeno, pa ne moze da uputi na
' akciju koja nije moguca (raniji tekst je za svaki blokiran status savetovao
' Cancel, koji nije dozvoljen za Approved/Sent/Paid/Archived/Cancelled).
Public Function SEFSendBlockedNextStep(ByVal sefStatus As String) As String

    If CanCancelSEFStatus(sefStatus) Then
        SEFSendBlockedNextStep = "Cancel the SEF document first, then handle it manually."
    ElseIf CanStornoSEFStatus(sefStatus) Then
        SEFSendBlockedNextStep = "Storno the SEF document first, then handle it manually."
    Else
        SEFSendBlockedNextStep = "Refresh the SEF status and check the SEF portal (manual review)."
    End If

End Function

' Storno se radi nad dokumentom koji je stvarno predat kupcu.
' "APPROVED" je zvanicni SEF naziv (AUD-032b); "ACCEPTED" ostaje zbog
' zatecenih redova i submit odgovora.
Public Function CanStornoSEFStatus(ByVal sefStatus As String) As Boolean
    Select Case UCase$(Trim$(sefStatus))
        Case "SENT", "ACCEPTED", "APPROVED", "REJECTED"
            CanStornoSEFStatus = True
        Case Else
            CanStornoSEFStatus = False
    End Select
End Function

Public Sub ValidateFakturaCanBeCancelledOnSEF(ByVal fakturaID As String)
    On Error GoTo EH

    Const SRC As String = "modSEFValidator.ValidateFakturaCanBeCancelledOnSEF"

    Dim sefDocumentId As String
    Dim sefStatus As String

    sefDocumentId = GetFakturaSEFDocumentId(fakturaID)

    If Len(Trim$(sefDocumentId)) = 0 Then
        Err.Raise ERR_SEF_STATE, SRC, _
                  "No SEFDocumentId found for faktura " & fakturaID
    End If

    sefStatus = GetFakturaSEFStatusText(fakturaID, SRC)

    If Not CanCancelSEFStatus(sefStatus) Then
        Err.Raise ERR_SEF_STATE, SRC, _
                  "Invoice cannot be cancelled on SEF in status: " & sefStatus
    End If

    Exit Sub

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.description
End Sub

Public Sub ValidateFakturaCanBeStorniranoOnSEF(ByVal fakturaID As String)
    On Error GoTo EH

    Const SRC As String = "modSEFValidator.ValidateFakturaCanBeStorniranoOnSEF"

    Dim sefDocumentId As String
    Dim sefStatus As String

    sefDocumentId = GetFakturaSEFDocumentId(fakturaID)

    If Len(Trim$(sefDocumentId)) = 0 Then
        Err.Raise ERR_SEF_STATE, SRC, _
                  "No SEFDocumentId found for faktura " & fakturaID
    End If

    sefStatus = GetFakturaSEFStatusText(fakturaID, SRC)

    If Not CanStornoSEFStatus(sefStatus) Then
        Err.Raise ERR_SEF_STATE, SRC, _
                  "Invoice cannot be storno on SEF in status: " & sefStatus
    End If

    Exit Sub

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.description
End Sub

' AUD-032b: telo je izdvojeno u `_Row` (pozivalac obezbedjuje TX) po obrascu
' koji projekat vec koristi (CreateFaktura/_TX, SaveMagacinCore/SaveMagacin).
' Razlog nije stil: clsTransaction.BeginTx PUCA na ugnezdjenu transakciju, pa se
' ceo tok resubmit-a nije mogao pokriti testom dok je logika zivela unutar TX-a.
Public Sub PrepareRejectedInvoiceForResubmit(ByVal fakturaID As String)

    Dim tx As clsTransaction

    On Error GoTo EH

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot TBL_SEF_SUBMISSION
    tx.AddTableSnapshot "tblSEFEventLog"

    PrepareRejectedInvoiceForResubmit_Row fakturaID

    tx.CommitTx

    Exit Sub

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogErr "PrepareRejectedInvoiceForResubmit"

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    If errNum <> 0 Then
        Err.Raise errNum, "PrepareRejectedInvoiceForResubmit", _
                  "Source=" & errSrc & " | " & errDesc
    Else
        Err.Raise ERR_SEF_STATE, "PrepareRejectedInvoiceForResubmit", _
                  "Unexpected error preparing rejected invoice; original Err was lost before EH capture."
    End If
End Sub

Public Sub PrepareRejectedInvoiceForResubmit_Row(ByVal fakturaID As String)

    Const SRC As String = "modSEFValidator.PrepareRejectedInvoiceForResubmit_Row"

    Dim currentState As String
    Dim lastSubmissionID As String
    Dim discharged As Boolean

    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_STATE, SRC, "FakturaID is required."
    End If

    currentState = GetFakturaSEFWorkflowState(fakturaID)

    ' Procitaj link PRE nego sto ga ClearFakturaLastSubmission_Row obrise --
    ' razduzuje se tacno ta submisija, ne "sve SENT za ovu fakturu".
    lastSubmissionID = GetLastSEFSubmissionID(fakturaID)

    If currentState <> WF_SEF_REJECTED Then
        Err.Raise ERR_SEF_STATE, SRC, _
            "Invoice is not in SEF_REJECTED state: " & currentState
    End If

    Call UpdateFakturaSEFState_Row( _
        fakturaID:=fakturaID, _
        newState:=WF_SEF_READY, _
        sefStatus:=WF_SEF_READY, _
        errorCode:="", _
        errorMessage:="", _
        submissionID:="")

    Call ClearFakturaLastSubmission_Row(fakturaID)

    ' AUD-032b: bez ovoga je dokumentovani tok bio mrtav -- faktura odbijena TEK
    ' NA REFRESH-u zadrzava submission red u statusu SENT (refresh ga namerno ne
    ' dira), pa bi `HasSuccessfulSEFSubmission` oborio bas ovaj pripremljeni
    ' resubmit kao duplikat.
    discharged = DischargeSEFSubmission_Row(lastSubmissionID, fakturaID)

    ' FAIL-CLOSED: priprema sme da uspe SAMO ako je faktura posle nje stvarno
    ' posiljiva. Ako je i dalje blokira uspesna submisija (prethodna je ACCEPTED,
    ' ili postoji stariji SENT red -- oba su neuskladjen podatak), bolje je da
    ' priprema padne glasno nego da operater dobije "pripremljeno" pa tek klik na
    ' slanje odbijanje zbog duplikata. TX se vraca, faktura ostaje SEF_REJECTED.
    If HasSuccessfulSEFSubmission(fakturaID) Then
        Err.Raise ERR_SEF_DUPLICATE, SRC, _
                  "Faktura still has a successful SEF submission after discharge " & _
                  "(LastSubmissionID=" & lastSubmissionID & _
                  "; Discharged=" & CStr(discharged) & "). Manual review required."
    End If

    Call AppendSEFEvent_Row( _
        fakturaID:=fakturaID, _
        submissionID:="", _
        eventType:=SEF_EVT_STATE_CHANGED, _
        message:="Rejected invoice prepared for corrected resubmission.", _
        details:="PreviousState=" & currentState & _
                 "; DischargedSubmissionID=" & lastSubmissionID & _
                 "; Discharged=" & CStr(discharged))

End Sub

Public Function IsFinalSEFStatus(ByVal sefStatus As String) As Boolean
    
    ' AUD-032b: prati zvanicni SEF enum kroz zajednicki klasifikator
    ' (APPROVED/ACCEPTED, REJECTED, STORNO/CANCELLED/DELETED).
    ' MISTAKE NIJE finalan -- to je greska pri slanju, ima Cancel/rucnu putanju.
    Select Case ClassifySEFExternalStatus(sefStatus)
        Case SEF_CLS_ACCEPTED, SEF_CLS_REJECTED, SEF_CLS_TERMINAL
            IsFinalSEFStatus = True
        Case Else
            IsFinalSEFStatus = False
    End Select
    
End Function

Public Function IsPendingSEFStatus(ByVal sefStatus As String) As Boolean
    
    ' AUD-032b: isti klasifikator kao svuda (dodaje SENDING i SEEN iz zvanicnog
    ' SEF enum-a, koje je stara lista propustala).
    Select Case ClassifySEFExternalStatus(sefStatus)
        Case SEF_CLS_PENDING
            IsPendingSEFStatus = True
        Case Else
            IsPendingSEFStatus = False
    End Select
    
End Function

Public Function GetSEFDisplayStatus(ByVal workflowState As String, ByVal sefStatus As String) As String
    
    If Len(Trim$(sefStatus)) > 0 Then
        GetSEFDisplayStatus = Trim$(sefStatus)
    Else
        GetSEFDisplayStatus = Trim$(workflowState)
    End If
    
End Function
