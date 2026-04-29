Option Explicit

' ============================================================
' modProductionHealthCheck
'
' Read-only production pre-flight / integrity audit.
'
' Entry point:
'   RunProductionHealthCheck
'
' This module DOES NOT modify business data.
' It only checks schema and high-risk integrity invariants.
' ============================================================

Private Const HEALTH_LOG_SHEET As String = "PRODUCTION_HEALTH_LOG"

Private m_RunID As String
Private m_Total As Long
Private m_Ok As Long
Private m_Warn As Long
Private m_Fail As Long

' ============================================================
' PUBLIC ENTRY POINT
' ============================================================

Public Sub RunProductionHealthCheck()
    On Error GoTo EH

    BeginHealthRun "PRODUCTION HEALTH CHECK"

    Check_CoreTablesAndColumns
    Check_NovacRowsAreFinanciallyValid
    Check_FakturaStavkeReferences
    Check_PrijemnicaFakturaFlags
    Check_FakturaPaymentConsistency
    Check_OtkupPaymentConsistency
    Check_OtkupOtpremnicaCrossZbirnaLinks
    Check_SEFOutboundConsistency
    Check_DocumentSoftDeleteReferences
    Check_GoogleSyncHealth

    EndHealthRun
    Exit Sub

EH:
    HealthFail "RunProductionHealthCheck fatal error", _
               "Err.Number=" & CStr(Err.Number) & _
               " Source=" & Err.SOURCE & _
               " Description=" & Err.description
    EndHealthRun
End Sub

' ============================================================
' CHECK 1: CORE SCHEMA
' ============================================================

Private Sub Check_CoreTablesAndColumns()
    On Error GoTo EH

    HealthRequireTable TBL_OTKUP
    HealthRequireTable TBL_OTPREMNICA
    HealthRequireTable TBL_ZBIRNA
    HealthRequireTable TBL_PRIJEMNICA
    HealthRequireTable TBL_FAKTURE
    HealthRequireTable TBL_FAKTURA_STAVKE
    HealthRequireTable TBL_NOVAC
    HealthRequireTable TBL_AMBALAZA

    HealthRequireColumns TBL_OTKUP, Array( _
        "OtkupID", "Datum", "KooperantID", "StanicaID", "Kolicina", "Cena", _
        "VozacID", "BrojDokumenta", "Klasa", "Stornirano", "BrojZbirne", _
        "Isplaceno", "DatumIsplate", "OtpremnicaID")

    HealthRequireColumns TBL_OTPREMNICA, Array( _
        "OtpremnicaID", "Datum", "StanicaID", "VozacID", "BrojOtpremnice", _
        "BrojZbirne", "Kolicina", "Cena", "KolAmbalaze", "Klasa", "Stornirano")

    HealthRequireColumns TBL_ZBIRNA, Array( _
        "ZbirnaID", "Datum", "VozacID", "BrojZbirne", "KupacID", _
        "UkupnoKolicina", "UkupnoAmbalaze", "Klasa", "Stornirano")

    HealthRequireColumns TBL_PRIJEMNICA, Array( _
        "PrijemnicaID", "Datum", "KupacID", "VozacID", "BrojPrijemnice", _
        "BrojZbirne", "Kolicina", "Cena", "KolAmbalaze", "kolAmbVracena", _
        "Klasa", "Fakturisano", "FakturaID", "Stornirano")

    HealthRequireColumns TBL_FAKTURE, Array( _
        "FakturaID", "BrojFakture", "Datum", "KupacID", "Iznos", _
        "Status", "DatumPlacanja", "Stornirano")

    HealthRequireColumns TBL_FAKTURA_STAVKE, Array( _
        "StavkaID", "FakturaID", "PrijemnicaID", "Kolicina", "Cena", _
        "Klasa", "BrojPrijemnice")

    HealthRequireColumns TBL_NOVAC, Array( _
        "NovacID", "Datum", "Partner", "PartnerID", "EntitetTip", _
        "FakturaID", "OtkupID", "Tip", "Uplata", "Isplata", "Stornirano")

    HealthOk "Core tables and required columns exist", ""
    Exit Sub

EH:
    HealthFail "Core tables and required columns exist", FormatHealthErr()
End Sub

' ============================================================
' CHECK 2: NOVAC INTEGRITY
' ============================================================

Private Sub Check_NovacRowsAreFinanciallyValid()
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_NOVAC)

    If IsEmpty(data) Then
        HealthWarn "Novac integrity", "tblNovac is empty."
        Exit Sub
    End If

    Dim colID As Long
    Dim colUplata As Long
    Dim colIsplata As Long
    Dim colTip As Long
    Dim colStorno As Long

    colID = RequireColumnIndex(TBL_NOVAC, "NovacID", "Check_NovacRowsAreFinanciallyValid")
    colUplata = RequireColumnIndex(TBL_NOVAC, "Uplata", "Check_NovacRowsAreFinanciallyValid")
    colIsplata = RequireColumnIndex(TBL_NOVAC, "Isplata", "Check_NovacRowsAreFinanciallyValid")
    colTip = RequireColumnIndex(TBL_NOVAC, "Tip", "Check_NovacRowsAreFinanciallyValid")
    colStorno = RequireColumnIndex(TBL_NOVAC, "Stornirano", "Check_NovacRowsAreFinanciallyValid")

    Dim i As Long
    Dim id As String
    Dim uplata As Double
    Dim isplata As Double
    Dim badCount As Long

    For i = 1 To UBound(data, 1)

        If IsStorniranoValue(data(i, colStorno)) Then GoTo NextRow

        id = Trim$(CStr(data(i, colID)))
        uplata = HealthNumeric(data(i, colUplata))
        isplata = HealthNumeric(data(i, colIsplata))

        If uplata < 0 Or isplata < 0 Then
            badCount = badCount + 1
            HealthFail "Novac negative amount", _
                       "NovacID=" & id & " Uplata=" & CStr(uplata) & " Isplata=" & CStr(isplata)
        End If

        If uplata > 0 And isplata > 0 Then
            badCount = badCount + 1
            HealthFail "Novac has both Uplata and Isplata", _
                       "NovacID=" & id & " Uplata=" & CStr(uplata) & " Isplata=" & CStr(isplata)
        End If

        If uplata = 0 And isplata = 0 Then
            badCount = badCount + 1
            HealthWarn "Novac has zero movement", _
                       "NovacID=" & id & " Tip=" & CStr(data(i, colTip))
        End If

        If Len(Trim$(CStr(data(i, colTip)))) = 0 Then
            badCount = badCount + 1
            HealthFail "Novac missing Tip", "NovacID=" & id
        End If

NextRow:
    Next i

    If badCount = 0 Then
        HealthOk "Novac rows are financially valid", ""
    End If

    Exit Sub

EH:
    HealthFail "Novac integrity check failed", FormatHealthErr()
End Sub

' ============================================================
' CHECK 3: FAKTURA STAVKE REFERENCES
' ============================================================

Private Sub Check_FakturaStavkeReferences()
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_FAKTURA_STAVKE)

    If IsEmpty(data) Then
        HealthWarn "Faktura stavke references", "tblFakturaStavke is empty."
        Exit Sub
    End If

    Dim colStavkaID As Long
    Dim colFakturaID As Long
    Dim colPrijemnicaID As Long

    colStavkaID = RequireColumnIndex(TBL_FAKTURA_STAVKE, "StavkaID", "Check_FakturaStavkeReferences")
    colFakturaID = RequireColumnIndex(TBL_FAKTURA_STAVKE, "FakturaID", "Check_FakturaStavkeReferences")
    colPrijemnicaID = RequireColumnIndex(TBL_FAKTURA_STAVKE, "PrijemnicaID", "Check_FakturaStavkeReferences")

    Dim i As Long
    Dim badCount As Long
    Dim stavkaID As String
    Dim fakturaID As String
    Dim prijemnicaID As String

    For i = 1 To UBound(data, 1)
        stavkaID = Trim$(CStr(data(i, colStavkaID)))
        fakturaID = Trim$(CStr(data(i, colFakturaID)))
        prijemnicaID = Trim$(CStr(data(i, colPrijemnicaID)))

        If Len(fakturaID) = 0 Or Not ActiveRowExists(TBL_FAKTURE, "FakturaID", fakturaID) Then
            badCount = badCount + 1
            HealthFail "FakturaStavka references missing faktura", _
                       "StavkaID=" & stavkaID & " FakturaID=" & fakturaID
        End If

        If Len(prijemnicaID) = 0 Or Not ActiveRowExists(TBL_PRIJEMNICA, "PrijemnicaID", prijemnicaID) Then
            badCount = badCount + 1
            HealthFail "FakturaStavka references missing prijemnica", _
                       "StavkaID=" & stavkaID & " PrijemnicaID=" & prijemnicaID
        End If
    Next i

    If badCount = 0 Then
        HealthOk "Faktura stavke references are valid", ""
    End If

    Exit Sub

EH:
    HealthFail "Faktura stavke reference check failed", FormatHealthErr()
End Sub

' ============================================================
' CHECK 4: PRIJEMNICA <-> FAKTURA FLAGS
' ============================================================

Private Sub Check_PrijemnicaFakturaFlags()
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_PRIJEMNICA)

    If IsEmpty(data) Then
        HealthWarn "Prijemnica faktura flags", "tblPrijemnica is empty."
        Exit Sub
    End If

    Dim colID As Long
    Dim colFakturisano As Long
    Dim colFakturaID As Long
    Dim colStorno As Long

    colID = RequireColumnIndex(TBL_PRIJEMNICA, "PrijemnicaID", "Check_PrijemnicaFakturaFlags")
    colFakturisano = RequireColumnIndex(TBL_PRIJEMNICA, "Fakturisano", "Check_PrijemnicaFakturaFlags")
    colFakturaID = RequireColumnIndex(TBL_PRIJEMNICA, "FakturaID", "Check_PrijemnicaFakturaFlags")
    colStorno = RequireColumnIndex(TBL_PRIJEMNICA, "Stornirano", "Check_PrijemnicaFakturaFlags")

    Dim i As Long
    Dim badCount As Long
    Dim prijID As String
    Dim fakturisano As String
    Dim fakturaID As String

    For i = 1 To UBound(data, 1)

        If IsStorniranoValue(data(i, colStorno)) Then GoTo NextRow

        prijID = Trim$(CStr(data(i, colID)))
        fakturisano = UCase$(Trim$(CStr(data(i, colFakturisano))))
        fakturaID = Trim$(CStr(data(i, colFakturaID)))

        If fakturisano = "DA" And Len(fakturaID) = 0 Then
            badCount = badCount + 1
            HealthFail "Prijemnica Fakturisano=Da without FakturaID", _
                       "PrijemnicaID=" & prijID
        End If

        If Len(fakturaID) > 0 And Not ActiveRowExists(TBL_FAKTURE, "FakturaID", fakturaID) Then
            badCount = badCount + 1
            HealthFail "Prijemnica points to missing/stornirana faktura", _
                       "PrijemnicaID=" & prijID & " FakturaID=" & fakturaID
        End If

NextRow:
    Next i

    If badCount = 0 Then
        HealthOk "Prijemnica faktura flags are valid", ""
    End If

    Exit Sub

EH:
    HealthFail "Prijemnica faktura flag check failed", FormatHealthErr()
End Sub

' ============================================================
' CHECK 5: FAKTURA PAYMENT CONSISTENCY
' ============================================================

Private Sub Check_FakturaPaymentConsistency()
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_FAKTURE)

    If IsEmpty(data) Then
        HealthWarn "Faktura payment consistency", "tblFakture is empty."
        Exit Sub
    End If

    Dim colID As Long
    Dim colIznos As Long
    Dim colStatus As Long
    Dim colDatumPlacanja As Long
    Dim colStorno As Long

    colID = RequireColumnIndex(TBL_FAKTURE, "FakturaID", "Check_FakturaPaymentConsistency")
    colIznos = RequireColumnIndex(TBL_FAKTURE, "Iznos", "Check_FakturaPaymentConsistency")
    colStatus = RequireColumnIndex(TBL_FAKTURE, "Status", "Check_FakturaPaymentConsistency")
    colDatumPlacanja = RequireColumnIndex(TBL_FAKTURE, "DatumPlacanja", "Check_FakturaPaymentConsistency")
    colStorno = RequireColumnIndex(TBL_FAKTURE, "Stornirano", "Check_FakturaPaymentConsistency")

    Dim i As Long
    Dim badCount As Long
    Dim fakturaID As String
    Dim iznos As Double
    Dim uplaceno As Double
    Dim statusText As String
    Dim datumPlacanja As String

    For i = 1 To UBound(data, 1)

        If IsStorniranoValue(data(i, colStorno)) Then GoTo NextRow

        fakturaID = Trim$(CStr(data(i, colID)))
        iznos = HealthNumeric(data(i, colIznos))
        uplaceno = HealthGetUplataForFaktura(fakturaID)

        statusText = UCase$(Trim$(CStr(data(i, colStatus))))
        datumPlacanja = Trim$(CStr(data(i, colDatumPlacanja)))

        If statusText = UCase$("Placeno") Or statusText = UCase$(STATUS_PLACENO) Then
            If iznos <= 0 Or uplaceno + 0.0001 < iznos Then
                badCount = badCount + 1
                HealthFail "Faktura marked paid without enough uplata", _
                           "FakturaID=" & fakturaID & _
                           " Iznos=" & CStr(iznos) & _
                           " Uplaceno=" & CStr(uplaceno)
            End If

            If Len(datumPlacanja) = 0 Then
                badCount = badCount + 1
                HealthWarn "Faktura paid without DatumPlacanja", _
                           "FakturaID=" & fakturaID
            End If
        End If

NextRow:
    Next i

    If badCount = 0 Then
        HealthOk "Faktura payment consistency is valid", ""
    End If

    Exit Sub

EH:
    HealthFail "Faktura payment consistency check failed", FormatHealthErr()
End Sub

' ============================================================
' CHECK 6: OTKUP PAYMENT CONSISTENCY
' ============================================================

Private Sub Check_OtkupPaymentConsistency()
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_OTKUP)

    If IsEmpty(data) Then
        HealthWarn "Otkup payment consistency", "tblOtkup is empty."
        Exit Sub
    End If

    Dim colID As Long
    Dim colKol As Long
    Dim colCena As Long
    Dim colIsplaceno As Long
    Dim colDatumIsplate As Long
    Dim colStorno As Long

    colID = RequireColumnIndex(TBL_OTKUP, "OtkupID", "Check_OtkupPaymentConsistency")
    colKol = RequireColumnIndex(TBL_OTKUP, "Kolicina", "Check_OtkupPaymentConsistency")
    colCena = RequireColumnIndex(TBL_OTKUP, "Cena", "Check_OtkupPaymentConsistency")
    colIsplaceno = RequireColumnIndex(TBL_OTKUP, "Isplaceno", "Check_OtkupPaymentConsistency")
    colDatumIsplate = RequireColumnIndex(TBL_OTKUP, "DatumIsplate", "Check_OtkupPaymentConsistency")
    colStorno = RequireColumnIndex(TBL_OTKUP, "Stornirano", "Check_OtkupPaymentConsistency")

    Dim i As Long
    Dim badCount As Long
    Dim otkupID As String
    Dim vrednost As Double
    Dim isplacenoAmount As Double
    Dim statusText As String
    Dim datumIsplate As String

    For i = 1 To UBound(data, 1)

        If IsStorniranoValue(data(i, colStorno)) Then GoTo NextRow

        otkupID = Trim$(CStr(data(i, colID)))
        vrednost = HealthNumeric(data(i, colKol)) * HealthNumeric(data(i, colCena))
        isplacenoAmount = HealthGetIsplataForOtkup(otkupID)

        statusText = UCase$(Trim$(CStr(data(i, colIsplaceno))))
        datumIsplate = Trim$(CStr(data(i, colDatumIsplate)))

        If statusText = UCase$("Da") Or statusText = UCase$(STATUS_ISPLACENO) Then
            If vrednost <= 0 Or isplacenoAmount + 0.0001 < vrednost Then
                badCount = badCount + 1
                HealthFail "Otkup marked paid without enough isplata", _
                           "OtkupID=" & otkupID & _
                           " Vrednost=" & CStr(vrednost) & _
                           " Isplaceno=" & CStr(isplacenoAmount)
            End If

            If Len(datumIsplate) = 0 Then
                badCount = badCount + 1
                HealthWarn "Otkup paid without DatumIsplate", _
                           "OtkupID=" & otkupID
            End If
        End If

NextRow:
    Next i

    If badCount = 0 Then
        HealthOk "Otkup payment consistency is valid", ""
    End If

    Exit Sub

EH:
    HealthFail "Otkup payment consistency check failed", FormatHealthErr()
End Sub

' ============================================================
' CHECK 7: OTKUP -> OTPREMNICA BROJZBIRNE LINK
' ============================================================

Private Sub Check_OtkupOtpremnicaCrossZbirnaLinks()
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_OTKUP)

    If IsEmpty(data) Then
        HealthWarn "Otkup/Otpremnica cross-zbirna links", "tblOtkup is empty."
        Exit Sub
    End If

    Dim colOtkID As Long
    Dim colOtkBrojZbirne As Long
    Dim colOtpremnicaID As Long
    Dim colStorno As Long

    colOtkID = RequireColumnIndex(TBL_OTKUP, "OtkupID", "Check_OtkupOtpremnicaCrossZbirnaLinks")
    colOtkBrojZbirne = RequireColumnIndex(TBL_OTKUP, "BrojZbirne", "Check_OtkupOtpremnicaCrossZbirnaLinks")
    colOtpremnicaID = RequireColumnIndex(TBL_OTKUP, "OtpremnicaID", "Check_OtkupOtpremnicaCrossZbirnaLinks")
    colStorno = RequireColumnIndex(TBL_OTKUP, "Stornirano", "Check_OtkupOtpremnicaCrossZbirnaLinks")

    Dim i As Long
    Dim badCount As Long
    Dim otkupID As String
    Dim otkZbr As String
    Dim otpID As String
    Dim otpZbr As String

    For i = 1 To UBound(data, 1)

        If IsStorniranoValue(data(i, colStorno)) Then GoTo NextRow

        otkupID = Trim$(CStr(data(i, colOtkID)))
        otkZbr = Trim$(CStr(data(i, colOtkBrojZbirne)))
        otpID = Trim$(CStr(data(i, colOtpremnicaID)))

        If Len(otpID) > 0 Then
            If Not ActiveRowExists(TBL_OTPREMNICA, "OtpremnicaID", otpID) Then
                badCount = badCount + 1
                HealthFail "Otkup points to missing/stornirana otpremnica", _
                           "OtkupID=" & otkupID & " OtpremnicaID=" & otpID
            Else
                otpZbr = Trim$(CStr(GetValueByKeySafe(TBL_OTPREMNICA, "OtpremnicaID", otpID, "BrojZbirne")))

                If Len(otkZbr) > 0 And Len(otpZbr) > 0 Then
                    If StrComp(otkZbr, otpZbr, vbTextCompare) <> 0 Then
                        badCount = badCount + 1
                        HealthFail "Cross-zbirna otkup/otpremnica link", _
                                   "OtkupID=" & otkupID & _
                                   " Otkup.BrojZbirne=" & otkZbr & _
                                   " OtpremnicaID=" & otpID & _
                                   " Otpremnica.BrojZbirne=" & otpZbr
                    End If
                End If
            End If
        End If

NextRow:
    Next i

    If badCount = 0 Then
        HealthOk "Otkup/Otpremnica BrojZbirne links are valid", ""
    End If

    Exit Sub

EH:
    HealthFail "Otkup/Otpremnica link check failed", FormatHealthErr()
End Sub

' ============================================================
' CHECK 8: SEF OUTBOUND CONSISTENCY
' ============================================================

Private Sub Check_SEFOutboundConsistency()
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_FAKTURE)

    If IsEmpty(data) Then
        HealthWarn "SEF outbound consistency", "tblFakture is empty."
        Exit Sub
    End If

    Dim colID As Long
    Dim colWorkflow As Long
    Dim colSEFDocID As Long
    Dim colSubmissionLast As Long
    Dim colSentAt As Long
    Dim colStorno As Long

    colID = RequireColumnIndex(TBL_FAKTURE, "FakturaID", "Check_SEFOutboundConsistency")
    colStorno = RequireColumnIndex(TBL_FAKTURE, "Stornirano", "Check_SEFOutboundConsistency")

    colWorkflow = GetColumnIndex(TBL_FAKTURE, "SEFWorkflowState")
    colSEFDocID = GetColumnIndex(TBL_FAKTURE, "SEFDocumentId")
    colSubmissionLast = GetColumnIndex(TBL_FAKTURE, "SEFSubmissionIDLast")
    colSentAt = GetColumnIndex(TBL_FAKTURE, "SEFSentAt")

    If colWorkflow = 0 Then
        HealthWarn "SEF outbound consistency skipped", "Column SEFWorkflowState not found."
        Exit Sub
    End If

    If colSEFDocID = 0 Then
        HealthWarn "SEF outbound consistency skipped", "Column SEFDocumentId not found."
        Exit Sub
    End If

    Dim i As Long
    Dim badCount As Long
    Dim fakturaID As String
    Dim wf As String
    Dim docID As String
    Dim submissionID As String
    Dim sentAt As String

    For i = 1 To UBound(data, 1)

        If IsStorniranoValue(data(i, colStorno)) Then GoTo NextRow

        fakturaID = Trim$(CStr(data(i, colID)))
        wf = UCase$(Trim$(CStr(data(i, colWorkflow))))
        docID = Trim$(CStr(data(i, colSEFDocID)))

        If colSubmissionLast > 0 Then submissionID = Trim$(CStr(data(i, colSubmissionLast))) Else submissionID = ""
        If colSentAt > 0 Then sentAt = Trim$(CStr(data(i, colSentAt))) Else sentAt = ""

        Select Case wf
            Case "SEF_SENT", "SEF_ACCEPTED", "SEF_CANCELLED", "SEF_STORNO"
                If Len(docID) = 0 Then
                    badCount = badCount + 1
                    HealthFail "SEF workflow state without SEFDocumentId", _
                               "FakturaID=" & fakturaID & " SEFWorkflowState=" & wf
                End If

                If Len(submissionID) = 0 And colSubmissionLast > 0 Then
                    badCount = badCount + 1
                    HealthWarn "SEF workflow state without last submission", _
                               "FakturaID=" & fakturaID & " SEFWorkflowState=" & wf
                End If

                If Len(sentAt) = 0 And colSentAt > 0 Then
                    badCount = badCount + 1
                    HealthWarn "SEF workflow state without SEFSentAt", _
                               "FakturaID=" & fakturaID & " SEFWorkflowState=" & wf
                End If
        End Select

NextRow:
    Next i

    If badCount = 0 Then
        HealthOk "SEF outbound consistency is valid", ""
    End If

    Exit Sub

EH:
    HealthFail "SEF outbound consistency check failed", FormatHealthErr()
End Sub

' ============================================================
' CHECK 9: SOFT DELETE REFERENCES
' ============================================================

Private Sub Check_DocumentSoftDeleteReferences()
    On Error GoTo EH

    Dim badCount As Long

    badCount = badCount + CountActiveReferencesToStornirano( _
        TBL_OTKUP, "OtkupID", "OtpremnicaID", _
        TBL_OTPREMNICA, "OtpremnicaID", _
        "Active otkup references stornirana otpremnica")

    badCount = badCount + CountActiveReferencesToStornirano( _
        TBL_PRIJEMNICA, "PrijemnicaID", "FakturaID", _
        TBL_FAKTURE, "FakturaID", _
        "Active prijemnica references stornirana faktura")

    badCount = badCount + CountActiveReferencesToStornirano( _
        TBL_FAKTURA_STAVKE, "StavkaID", "PrijemnicaID", _
        TBL_PRIJEMNICA, "PrijemnicaID", _
        "Faktura stavka references stornirana prijemnica")

    If badCount = 0 Then
        HealthOk "No active references to stornirano documents found", ""
    End If

    Exit Sub

EH:
    HealthFail "Soft-delete reference check failed", FormatHealthErr()
End Sub

' ============================================================
' CHECK 10: GOOGLE / PWA SYNC HEALTH
' ============================================================

Private Sub Check_GoogleSyncHealth()
    On Error GoTo EH

    Check_GoogleSyncConfig
    Check_GoogleSyncAuth
    Check_GoogleSyncFolderReachable
    Check_GoogleSyncMasterSchema
    Check_GoogleSyncFeatureFlags

    HealthOk "Google/PWA sync health is valid", ""
    Exit Sub

EH:
    HealthFail "Google/PWA sync health check failed", FormatHealthErr()
End Sub

Private Sub Check_GoogleSyncConfig()
    On Error GoTo EH

    HealthRequireTable TBL_SEF_CONFIG
    HealthRequireColumns TBL_SEF_CONFIG, Array("ConfigKey", "ConfigValue")

    HealthRequireConfigValue "GOOGLE_CLIENT_ID"
    HealthRequireConfigValue "GOOGLE_CLIENT_SECRET"
    HealthRequireConfigValue "GOOGLE_PWA_FOLDER_ID"
    HealthRequireConfigValue "GOOGLE_REFRESH_TOKEN"

    ' Optional but useful; may be empty before first refresh.
    HealthCheckOptionalConfig "GOOGLE_ACCESS_TOKEN"
    HealthCheckOptionalConfig "GOOGLE_TOKEN_EXPIRES_AT"
    HealthCheckOptionalConfig "SHEETS_SYNC_ENABLED"
    HealthCheckOptionalConfig "APP_LAST_HEALTHCHECK_AT"

    HealthOk "Google sync config keys are valid", ""
    Exit Sub

EH:
    HealthFail "Google sync config keys are valid", FormatHealthErr()
End Sub

Private Sub Check_GoogleSyncAuth()
    On Error GoTo EH

    Dim token As String

    If Not IsGoogleAuthConfigured() Then
        HealthFail "Google OAuth configured", "IsGoogleAuthConfigured=False"
        Exit Sub
    End If

    token = GetAccessToken()

    If Len(Trim$(token)) = 0 Then
        HealthFail "Google access token available", _
                   "GetAccessToken returned empty. Re-run RunGoogleAuthSetup."
        Exit Sub
    End If

    HealthOk "Google OAuth token is available", ""
    Exit Sub

EH:
    HealthFail "Google OAuth token is available", FormatHealthErr()
End Sub

Private Sub Check_GoogleSyncFolderReachable()
    On Error GoTo EH

    Dim folderID As String
    Dim probeName As String
    Dim foundID As String

    folderID = Trim$(GetConfigValue("GOOGLE_PWA_FOLDER_ID"))

    If Len(folderID) = 0 Then
        HealthFail "Google PWA folder configured", "GOOGLE_PWA_FOLDER_ID is empty."
        Exit Sub
    End If

    ' Read-only probe. Expected result is empty string.
    ' If auth/folder/API is broken, GetSpreadsheetID logs internally and should fail safely.
    probeName = "__SYNC_HEALTH_PROBE_DOES_NOT_EXIST__" & Format$(Now, "yyyymmddhhnnss")
    foundID = GetSpreadsheetID(probeName, folderID)

    If Len(Trim$(foundID)) > 0 Then
        HealthWarn "Google PWA folder read probe", _
                   "Unexpected spreadsheet found for probe name: " & probeName
    Else
        HealthOk "Google PWA folder read path is reachable", ""
    End If

    Exit Sub

EH:
    HealthFail "Google PWA folder read path is reachable", FormatHealthErr()
End Sub

Private Sub Check_GoogleSyncMasterSchema()
    On Error GoTo EH

    HealthRequireTable TBL_OTKUP
    HealthRequireTable TBL_AMBALAZA
    HealthRequireTable TBL_KOOPERANTI
    HealthRequireTable TBL_STANICE
    HealthRequireTable TBL_KULTURE

    HealthRequireColumns TBL_OTKUP, Array( _
        "OtkupID", _
        "ClientRecordID", _
        "Datum", _
        "KooperantID", _
        "StanicaID", _
        "Kolicina", _
        "Cena", _
        "Klasa", _
        "Stornirano", _
        "OtpremnicaID")

    HealthRequireColumns TBL_KOOPERANTI, Array( _
        "KooperantID", _
        COL_KOOP_STANICA)

    HealthRequireColumns TBL_STANICE, Array( _
        "StanicaID")

    HealthOk "Google sync master tables/columns are valid", ""
    Exit Sub

EH:
    HealthFail "Google sync master tables/columns are valid", FormatHealthErr()
End Sub

Private Sub Check_GoogleSyncFeatureFlags()
    On Error GoTo EH

    Dim syncEnabled As String
    Dim setupCompleted As String

    syncEnabled = Trim$(GetConfigValue("SHEETS_SYNC_ENABLED"))
    setupCompleted = Trim$(GetConfigValue("APP_SETUP_COMPLETED"))

    If Len(syncEnabled) = 0 Then
        HealthWarn "SHEETS_SYNC_ENABLED", _
                   "Config key missing/empty. Sync is not explicitly enabled."
    ElseIf Not HealthIsTruthy(syncEnabled) Then
        HealthWarn "SHEETS_SYNC_ENABLED", _
                   "Sync flag is not enabled: " & syncEnabled
    Else
        HealthOk "SHEETS_SYNC_ENABLED is enabled", ""
    End If

    If Len(setupCompleted) = 0 Then
        HealthWarn "APP_SETUP_COMPLETED", "Config key missing/empty."
    Else
        HealthOk "APP_SETUP_COMPLETED is present", ""
    End If

    Exit Sub

EH:
    HealthFail "Google sync feature flags are valid", FormatHealthErr()
End Sub

' ============================================================
' SUPPORT: REFERENCE CHECK
' ============================================================

Private Function CountActiveReferencesToStornirano(ByVal sourceTable As String, _
                                                   ByVal sourceIDColumn As String, _
                                                   ByVal sourceRefColumn As String, _
                                                   ByVal targetTable As String, _
                                                   ByVal targetIDColumn As String, _
                                                   ByVal checkName As String) As Long
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(sourceTable)

    If IsEmpty(data) Then Exit Function

    Dim colSourceID As Long
    Dim colRef As Long
    Dim colSourceStorno As Long

    colSourceID = RequireColumnIndex(sourceTable, sourceIDColumn, "CountActiveReferencesToStornirano")
    colRef = RequireColumnIndex(sourceTable, sourceRefColumn, "CountActiveReferencesToStornirano")
    colSourceStorno = GetColumnIndex(sourceTable, "Stornirano")

    Dim i As Long
    Dim sourceID As String
    Dim refID As String

    For i = 1 To UBound(data, 1)

        If colSourceStorno > 0 Then
            If IsStorniranoValue(data(i, colSourceStorno)) Then GoTo NextRow
        End If

        sourceID = Trim$(CStr(data(i, colSourceID)))
        refID = Trim$(CStr(data(i, colRef)))

        If Len(refID) > 0 Then
            If IsStorniranoRow(targetTable, targetIDColumn, refID) Then
                CountActiveReferencesToStornirano = CountActiveReferencesToStornirano + 1
                HealthFail checkName, _
                           sourceTable & "." & sourceIDColumn & "=" & sourceID & _
                           " -> " & targetTable & "." & targetIDColumn & "=" & refID
            End If
        End If

NextRow:
    Next i

    Exit Function

EH:
    HealthFail checkName & " failed", FormatHealthErr()
End Function

' ============================================================
' HEALTH RUN LOGGING
' ============================================================

Private Sub BeginHealthRun(ByVal suiteName As String)
    m_RunID = Format$(Now, "yyyymmddhhnnss") & "-" & CStr(Int((9000 * Rnd) + 1000))

    m_Total = 0
    m_Ok = 0
    m_Warn = 0
    m_Fail = 0

    InitHealthLog

    Debug.Print String$(70, "=")
    Debug.Print suiteName & " started at " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Debug.Print "RunID=" & m_RunID
    Debug.Print String$(70, "=")

    AppendHealthLog "SUITE", suiteName, "START", ""
End Sub

Private Sub EndHealthRun()
    Dim summary As String

    summary = "RunID=" & m_RunID & _
              " | Total=" & m_Total & _
              " | OK=" & m_Ok & _
              " | Warn=" & m_Warn & _
              " | Fail=" & m_Fail

    Debug.Print String$(70, "-")
    Debug.Print "PRODUCTION HEALTH SUMMARY: " & summary
    Debug.Print String$(70, "-")

    AppendHealthLog "SUITE", "SUMMARY", "INFO", summary
    
    On Error Resume Next
        SetConfigValue "APP_LAST_HEALTHCHECK_AT", Format$(Now, "yyyy-mm-dd hh:nn:ss")
    On Error GoTo 0
    
    If m_Fail > 0 Then
        MsgBox "Production health check finished with FAILURES." & vbCrLf & summary, _
               vbCritical, APP_NAME
    ElseIf m_Warn > 0 Then
        MsgBox "Production health check finished with warnings." & vbCrLf & summary, _
               vbExclamation, APP_NAME
    Else
        MsgBox "Production health check passed." & vbCrLf & summary, _
               vbInformation, APP_NAME
    End If
End Sub

Private Sub HealthOk(ByVal checkName As String, ByVal details As String)
    m_Total = m_Total + 1
    m_Ok = m_Ok + 1

    Debug.Print "[OK] " & checkName & IIf(Len(details) > 0, " :: " & details, "")
    AppendHealthLog "CHECK", checkName, "OK", details
End Sub

Private Sub HealthWarn(ByVal checkName As String, ByVal details As String)
    m_Total = m_Total + 1
    m_Warn = m_Warn + 1

    Debug.Print "[WARN] " & checkName & IIf(Len(details) > 0, " :: " & details, "")
    AppendHealthLog "CHECK", checkName, "WARN", details
End Sub

Private Sub HealthFail(ByVal checkName As String, ByVal details As String)
    m_Total = m_Total + 1
    m_Fail = m_Fail + 1

    Debug.Print "[FAIL] " & checkName & IIf(Len(details) > 0, " :: " & details, "")
    AppendHealthLog "CHECK", checkName, "FAIL", details
End Sub

Private Sub InitHealthLog()
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(HEALTH_LOG_SHEET)

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = HEALTH_LOG_SHEET
        ws.Range("A1:G1").value = Array("Timestamp", "RunID", "Kind", "Name", "Status", "Details", "Operator")
        ws.rows(1).Font.Bold = True
    End If
End Sub

Private Sub AppendHealthLog(ByVal kindText As String, _
                            ByVal nameText As String, _
                            ByVal statusText As String, _
                            ByVal detailsText As String)
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(HEALTH_LOG_SHEET)

    If ws Is Nothing Then Exit Sub

    Dim r As Long
    r = ws.cells(ws.rows.count, 1).End(xlUp).row + 1

    ws.cells(r, 1).value = Now
    ws.cells(r, 2).value = m_RunID
    ws.cells(r, 3).value = kindText
    ws.cells(r, 4).value = nameText
    ws.cells(r, 5).value = statusText
    ws.cells(r, 6).value = Left$(detailsText, 2000)
    ws.cells(r, 7).value = Environ$("Username")
End Sub

' ============================================================
' GENERAL HELPERS
' ============================================================

Private Sub HealthRequireTable(ByVal tableName As String)
    If Not HealthTableExists(tableName) Then
        Err.Raise vbObjectError + 9601, "HealthRequireTable", _
                  "Missing table: " & tableName
    End If
End Sub

Private Sub HealthRequireColumns(ByVal tableName As String, ByVal columns As Variant)
    Dim i As Long

    For i = LBound(columns) To UBound(columns)
        If GetColumnIndex(tableName, CStr(columns(i))) = 0 Then
            Err.Raise vbObjectError + 9602, "HealthRequireColumns", _
                      "Missing column: " & tableName & "." & CStr(columns(i))
        End If
    Next i
End Sub

Private Sub HealthRequireConfigValue(ByVal configKey As String)
    Dim value As String

    value = Trim$(GetConfigValue(configKey))

    If Len(value) = 0 Then
        Err.Raise vbObjectError + 9701, "HealthRequireConfigValue", _
                  "Missing required config value: " & configKey
    End If
End Sub

Private Sub HealthCheckOptionalConfig(ByVal configKey As String)
    Dim value As String

    value = Trim$(GetConfigValue(configKey))

    If Len(value) = 0 Then
        HealthWarn "Optional config is empty", configKey
    Else
        HealthOk "Optional config present", configKey
    End If
End Sub

Private Function HealthIsTruthy(ByVal value As String) As Boolean
    Dim s As String

    s = UCase$(Trim$(value))

    HealthIsTruthy = (s = "TRUE" Or s = "DA" Or s = "YES" Or s = "1")
End Function


Private Function HealthTableExists(ByVal tableName As String) As Boolean
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.name, tableName, vbTextCompare) = 0 Then
                HealthTableExists = True
                Exit Function
            End If
        Next lo
    Next ws
End Function

Private Function HealthNumeric(ByVal value As Variant) As Double
    If IsNumeric(value) Then
        HealthNumeric = CDbl(value)
    Else
        HealthNumeric = 0#
    End If
End Function

Private Function IsStorniranoValue(ByVal value As Variant) As Boolean
    IsStorniranoValue = (UCase$(Trim$(CStr(value))) = "DA")
End Function

Private Function ActiveRowExists(ByVal tableName As String, _
                                 ByVal idColumn As String, _
                                 ByVal idValue As String) As Boolean
    If Len(Trim$(idValue)) = 0 Then Exit Function

    Dim data As Variant
    data = GetTableData(tableName)

    If IsEmpty(data) Then Exit Function

    Dim colID As Long
    Dim colStorno As Long

    colID = RequireColumnIndex(tableName, idColumn, "ActiveRowExists")
    colStorno = GetColumnIndex(tableName, "Stornirano")

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, colID))) = Trim$(idValue) Then
            If colStorno > 0 Then
                If IsStorniranoValue(data(i, colStorno)) Then
                    ActiveRowExists = False
                    Exit Function
                End If
            End If

            ActiveRowExists = True
            Exit Function
        End If
    Next i
End Function

Private Function IsStorniranoRow(ByVal tableName As String, _
                                 ByVal idColumn As String, _
                                 ByVal idValue As String) As Boolean
    If Len(Trim$(idValue)) = 0 Then Exit Function

    Dim data As Variant
    data = GetTableData(tableName)

    If IsEmpty(data) Then Exit Function

    Dim colID As Long
    Dim colStorno As Long

    colID = RequireColumnIndex(tableName, idColumn, "IsStorniranoRow")
    colStorno = GetColumnIndex(tableName, "Stornirano")

    If colStorno = 0 Then Exit Function

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, colID))) = Trim$(idValue) Then
            IsStorniranoRow = IsStorniranoValue(data(i, colStorno))
            Exit Function
        End If
    Next i
End Function

Private Function GetValueByKeySafe(ByVal tableName As String, _
                                   ByVal keyColumn As String, _
                                   ByVal keyValue As String, _
                                   ByVal returnColumn As String) As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(tableName)

    If IsEmpty(data) Then
        GetValueByKeySafe = Empty
        Exit Function
    End If

    Dim colKey As Long
    Dim colReturn As Long

    colKey = RequireColumnIndex(tableName, keyColumn, "GetValueByKeySafe")
    colReturn = RequireColumnIndex(tableName, returnColumn, "GetValueByKeySafe")

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, colKey))) = Trim$(keyValue) Then
            GetValueByKeySafe = data(i, colReturn)
            Exit Function
        End If
    Next i

    GetValueByKeySafe = Empty
    Exit Function

EH:
    GetValueByKeySafe = Empty
End Function

Private Function FormatHealthErr() As String
    FormatHealthErr = "Err.Number=" & CStr(Err.Number) & _
                      " Source=" & Err.SOURCE & _
                      " Description=" & Err.description
End Function

' ============================================================
' LOCAL PAYMENT SUMS
' ============================================================

Private Function HealthGetUplataForFaktura(ByVal fakturaID As String) As Double
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_NOVAC)

    If IsEmpty(data) Then Exit Function

    data = ExcludeStornirano(data, TBL_NOVAC)

    If IsEmpty(data) Then Exit Function

    Dim colFakturaID As Long
    Dim colUplata As Long

    colFakturaID = RequireColumnIndex(TBL_NOVAC, "FakturaID", "HealthGetUplataForFaktura")
    colUplata = RequireColumnIndex(TBL_NOVAC, "Uplata", "HealthGetUplataForFaktura")

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, colFakturaID))) = Trim$(fakturaID) Then
            If IsNumeric(data(i, colUplata)) Then
                HealthGetUplataForFaktura = HealthGetUplataForFaktura + CDbl(data(i, colUplata))
            End If
        End If
    Next i

    Exit Function

EH:
    HealthGetUplataForFaktura = 0#
End Function

Private Function HealthGetIsplataForOtkup(ByVal otkupID As String) As Double
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_NOVAC)

    If IsEmpty(data) Then Exit Function

    data = ExcludeStornirano(data, TBL_NOVAC)

    If IsEmpty(data) Then Exit Function

    Dim colOtkupID As Long
    Dim colIsplata As Long

    colOtkupID = RequireColumnIndex(TBL_NOVAC, "OtkupID", "HealthGetIsplataForOtkup")
    colIsplata = RequireColumnIndex(TBL_NOVAC, "Isplata", "HealthGetIsplataForOtkup")

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, colOtkupID))) = Trim$(otkupID) Then
            If IsNumeric(data(i, colIsplata)) Then
                HealthGetIsplataForOtkup = HealthGetIsplataForOtkup + CDbl(data(i, colIsplata))
            End If
        End If
    Next i

    Exit Function

EH:
    HealthGetIsplataForOtkup = 0#
End Function

