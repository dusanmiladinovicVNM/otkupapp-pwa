Attribute VB_Name = "modFakturaTests"
Option Explicit

' ============================================================
' modFakturaTests
' Dev/test-only smoke/regression suite for modFaktura.
' ============================================================

Private m_Total As Long
Private m_Passed As Long
Private m_Failed As Long
Private m_Skipped As Long

Private Const FAKTURA_TEST_LOG_SHEET As String = "FAKTURA_TEST_LOG"

Public Sub RunFakturaSmokeSuite()
    On Error GoTo EH

    ResetFakturaTestCounters
    InitFakturaTestLog

    StartFakturaSuite "FAKTURA SMOKE SUITE"

    Test_CreateFakturaUsesCanonicalPrijemnicaValues
    Test_CreateFakturaBlocksDuplicatePrijemnica
    Test_CreateFakturaBlocksStorniranaPrijemnica
    Test_CreateFakturaBlocksAlreadyFakturisanaPrijemnica

    ' RF-08 / AUD-011 + AUD-027
    Test_CreateFakturaBlocksPrijemnicaDrugogKupca
    Test_CreateFakturaBlocksDuplicatePrijemnicaID

    Test_UpdateFakturaStatusPreservesExistingDatumPlacanja
    Test_UpdateFakturaStatusReopensWhenPaymentMissing
    Test_UpdateFakturaStatusSkipsStorniranaFaktura

    Test_PrintFakturaBlocksStorniranaFaktura
    Test_ReprintBlocksStorniraniOtkup
    Test_FakturaSablonCleanupPratiBrojStavki

    FinishFakturaSuite

    ' HARD GATE -- summary je vec ispisan iznad.
    ' On Error GoTo 0: raise ne sme da udje u EH ispod (dupli FAIL brojac,
    ' dupli FinishFakturaSuite i dupli MsgBox). Isti obrazac kao
    ' modGoogleSyncSmokeTests.RunMasterSyncSmokeSuite.
    On Error GoTo 0
    RequireFakturaSuiteGreen
    Exit Sub

EH:
    LogFakturaFatal "RunFakturaSmokeSuite", Err.Number, Err.description
    FinishFakturaSuite
    RequireFakturaSuiteGreen
End Sub

' Tvrd gate: suite koja je pala mora da PADNE i za pozivaoca (isti obrazac kao
' modIzvestajTests.RunIzvestajTests) -- inace zeleno/crveno vidi samo onaj ko
' cita Immediate prozor.
Private Sub RequireFakturaSuiteGreen()
    If m_Failed <= 0 Then Exit Sub

    Err.Raise vbObjectError + 9210, "modFakturaTests.RunFakturaSmokeSuite", _
              "RunFakturaSmokeSuite: " & m_Failed & " od " & m_Total & " provera palo."
End Sub

Private Sub Test_CreateFakturaUsesCanonicalPrijemnicaValues()
    On Error GoTo EH

    Dim code As String
    code = NewFakturaTestCode("CANON")

    Dim kupacID As String
    Dim prijemnicaID As String
    Dim realBrojPrij As String

    kupacID = "KUP-TST-" & code
    prijemnicaID = "PRJ-TST-" & code
    realBrojPrij = "TST-PRJ-" & code

    AppendTestPrijemnicaRow _
        prijemnicaID:=prijemnicaID, _
        kupacID:=kupacID, _
        brojPrijemnice:=realBrojPrij, _
        kolicina:=100#, _
        cena:=10#, _
        klasa:="I", _
        stornirano:="", _
        fakturisano:="", _
        fakturaID:=""

    Dim stavke As New Collection

    ' Namerno pogreL^ni caller payload.
    ' CreateFaktura sme da koristi samo stavka(0) = PrijemnicaID.
    stavke.Add Array(prijemnicaID, 9999#, 9999#, "BAD-KLASA", "BAD-BROJ")

    Dim fakturaID As String
    fakturaID = CreateFaktura_TX(kupacID, stavke)

    AssertFakturaTrue Len(Trim$(fakturaID)) > 0, _
                      "CreateFaktura returns FakturaID"

    If Len(Trim$(fakturaID)) = 0 Then Exit Sub

    Dim fakturaIznos As Double
    fakturaIznos = CDbl(LookupValue(TBL_FAKTURE, COL_FAK_ID, fakturaID, COL_FAK_IZNOS))

    AssertFakturaDoubleEquals 1000#, fakturaIznos, _
                              "CreateFaktura uses canonical Prijemnica total"

    AssertFakturaDoubleEquals 100#, _
        CDbl(GetFirstFakturaStavkaValue(fakturaID, COL_FS_KOLICINA)), _
        "FakturaStavka Koli" & ChrW(269) & "ina comes from Prijemnica"

    AssertFakturaDoubleEquals 10#, _
        CDbl(GetFirstFakturaStavkaValue(fakturaID, COL_FS_CENA)), _
        "FakturaStavka Cena comes from Prijemnica"

    AssertFakturaTextEquals "I", _
        CStr(GetFirstFakturaStavkaValue(fakturaID, COL_FS_KLASA)), _
        "FakturaStavka Klasa comes from Prijemnica"

    AssertFakturaTextEquals realBrojPrij, _
        CStr(GetFirstFakturaStavkaValue(fakturaID, COL_FS_BROJ_PRIJEMNICE)), _
        "FakturaStavka BrojPrijemnice comes from Prijemnica"

    AssertFakturaTextEquals "Da", _
        CStr(LookupValue(TBL_PRIJEMNICA, COL_PRJ_ID, prijemnicaID, COL_PRJ_FAKTURISANO)), _
        "CreateFaktura marks Prijemnica as fakturisano"

    AssertFakturaTextEquals fakturaID, _
        CStr(LookupValue(TBL_PRIJEMNICA, COL_PRJ_ID, prijemnicaID, COL_PRJ_FAKTURA_ID)), _
        "CreateFaktura writes FakturaID back to Prijemnica"

    Exit Sub

EH:
    LogFakturaFail "CreateFaktura uses canonical Prijemnica values", _
                   FormatErrDetails()
End Sub

Private Sub Test_CreateFakturaBlocksDuplicatePrijemnica()
    On Error GoTo EH

    Dim code As String
    code = NewFakturaTestCode("DUPPRJ")

    Dim kupacID As String
    Dim prijemnicaID As String

    kupacID = "KUP-TST-" & code
    prijemnicaID = "PRJ-TST-" & code

    AppendTestPrijemnicaRow prijemnicaID, kupacID, "TST-PRJ-" & code, 50#, 20#, "I", "", "", ""

    Dim stavke As New Collection
    stavke.Add Array(prijemnicaID)
    stavke.Add Array(prijemnicaID)

    Dim fakturaID As String
    fakturaID = CreateFaktura_TX(kupacID, stavke)

    If Len(Trim$(fakturaID)) = 0 Then
        LogFakturaPass "CreateFaktura blocks duplicate Prijemnica"
    Else
        LogFakturaFail "CreateFaktura blocks duplicate Prijemnica", _
                       "Expected empty result, got FakturaID=" & fakturaID
    End If

    AssertFakturaTextEquals "", _
        CStr(LookupValue(TBL_PRIJEMNICA, COL_PRJ_ID, prijemnicaID, COL_PRJ_FAKTURISANO)), _
        "Duplicate block leaves Prijemnica not fakturisano"

    Exit Sub

EH:
    LogFakturaFail "CreateFaktura blocks duplicate Prijemnica", _
                   FormatErrDetails()
End Sub

Private Sub Test_CreateFakturaBlocksStorniranaPrijemnica()
    On Error GoTo EH

    Dim code As String
    code = NewFakturaTestCode("STOPRJ")

    Dim kupacID As String
    Dim prijemnicaID As String

    kupacID = "KUP-TST-" & code
    prijemnicaID = "PRJ-TST-" & code

    AppendTestPrijemnicaRow prijemnicaID, kupacID, "TST-PRJ-" & code, 50#, 20#, "I", "Da", "", ""

    Dim stavke As New Collection
    stavke.Add Array(prijemnicaID)

    Dim fakturaID As String
    fakturaID = CreateFaktura_TX(kupacID, stavke)

    If Len(Trim$(fakturaID)) = 0 Then
        LogFakturaPass "CreateFaktura blocks stornirana Prijemnica"
    Else
        LogFakturaFail "CreateFaktura blocks stornirana Prijemnica", _
                       "Expected empty result, got FakturaID=" & fakturaID
    End If

    Exit Sub

EH:
    LogFakturaFail "CreateFaktura blocks stornirana Prijemnica", _
                   FormatErrDetails()
End Sub

Private Sub Test_CreateFakturaBlocksAlreadyFakturisanaPrijemnica()
    On Error GoTo EH

    Dim code As String
    code = NewFakturaTestCode("ALRPRJ")

    Dim kupacID As String
    Dim prijemnicaID As String

    kupacID = "KUP-TST-" & code
    prijemnicaID = "PRJ-TST-" & code

    AppendTestPrijemnicaRow prijemnicaID, kupacID, "TST-PRJ-" & code, 50#, 20#, "I", "", "Da", ""

    Dim stavke As New Collection
    stavke.Add Array(prijemnicaID)

    Dim fakturaID As String
    fakturaID = CreateFaktura_TX(kupacID, stavke)

    If Len(Trim$(fakturaID)) = 0 Then
        LogFakturaPass "CreateFaktura blocks already fakturisana Prijemnica"
    Else
        LogFakturaFail "CreateFaktura blocks already fakturisana Prijemnica", _
                       "Expected empty result, got FakturaID=" & fakturaID
    End If

CleanExit:
    On Error Resume Next
    MarkTestPrijemnicaStornirano prijemnicaID
    On Error GoTo 0
    Exit Sub

EH:
    LogFakturaFail "CreateFaktura blocks already fakturisana Prijemnica", _
                   FormatErrDetails()
    Resume CleanExit
End Sub

' Markira SVE redove sa tim PrijemnicaID (duplikat-test namerno ostavlja dva),
' da test podaci ne ostanu aktivni u tabeli.
Private Sub MarkTestPrijemnicaStornirano(ByVal prijemnicaID As String)
    Dim rows As Collection
    Dim i As Long

    If Len(Trim$(prijemnicaID)) = 0 Then Exit Sub

    Set rows = FindRows(TBL_PRIJEMNICA, COL_PRJ_ID, prijemnicaID)

    If rows Is Nothing Then Exit Sub
    If rows.count = 0 Then Exit Sub

    For i = 1 To rows.count
        Call UpdateCell(TBL_PRIJEMNICA, CLng(rows(i)), COL_STORNIRANO, "Da")
    Next i
End Sub

' AUD-011 / FM-0034 #1: prijemnica tudjeg kupca ne sme da udje u fakturu.
Private Sub Test_CreateFakturaBlocksPrijemnicaDrugogKupca()
    On Error GoTo EH

    Dim code As String
    code = NewFakturaTestCode("XKUPAC")

    Dim kupacID As String
    Dim drugiKupacID As String
    Dim prijemnicaID As String

    kupacID = "KUP-TST-A-" & code
    drugiKupacID = "KUP-TST-B-" & code
    prijemnicaID = "PRJ-TST-" & code

    ' Prijemnica pripada kupcu B, faktura se pravi za kupca A.
    AppendTestPrijemnicaRow prijemnicaID, drugiKupacID, "TST-PRJ-" & code, _
                            50#, 20#, "I", "", "", ""

    Dim stavke As New Collection
    stavke.Add Array(prijemnicaID)

    Dim fakturaID As String
    fakturaID = CreateFaktura_TX(kupacID, stavke)

    If Len(Trim$(fakturaID)) = 0 Then
        LogFakturaPass "CreateFaktura blocks Prijemnica of another Kupac"
    Else
        LogFakturaFail "CreateFaktura blocks Prijemnica of another Kupac", _
                       "Expected empty result, got FakturaID=" & fakturaID
    End If

    AssertFakturaTextEquals "", _
        CStr(LookupValue(TBL_PRIJEMNICA, COL_PRJ_ID, prijemnicaID, COL_PRJ_FAKTURISANO)), _
        "Cross-kupac block leaves Prijemnica not fakturisano"

CleanExit:
    On Error Resume Next
    MarkTestPrijemnicaStornirano prijemnicaID
    On Error GoTo 0
    Exit Sub

EH:
    LogFakturaFail "CreateFaktura blocks Prijemnica of another Kupac", _
                   FormatErrDetails()
    Resume CleanExit
End Sub

' AUD-011 / FM-0034 #2: dva aktivna reda sa istim PrijemnicaID -> Count=1 guard
' mora da pukne umesto da tiho uzme prvi pogodak.
Private Sub Test_CreateFakturaBlocksDuplicatePrijemnicaID()
    On Error GoTo EH

    Dim code As String
    code = NewFakturaTestCode("DUPID")

    Dim kupacID As String
    Dim prijemnicaID As String

    kupacID = "KUP-TST-" & code
    prijemnicaID = "PRJ-TST-" & code

    ' Isti PrijemnicaID, razlicite kolicine/cene -> tihi izbor prvog reda bi
    ' fakturisao 1000, a ne 2100.
    AppendTestPrijemnicaRow prijemnicaID, kupacID, "TST-PRJ-A-" & code, _
                            50#, 20#, "I", "", "", ""
    AppendTestPrijemnicaRow prijemnicaID, kupacID, "TST-PRJ-B-" & code, _
                            70#, 30#, "I", "", "", ""

    Dim stavke As New Collection
    stavke.Add Array(prijemnicaID)

    Dim fakturaID As String
    fakturaID = CreateFaktura_TX(kupacID, stavke)

    If Len(Trim$(fakturaID)) = 0 Then
        LogFakturaPass "CreateFaktura blocks duplicate PrijemnicaID rows"
    Else
        LogFakturaFail "CreateFaktura blocks duplicate PrijemnicaID rows", _
                       "Expected empty result, got FakturaID=" & fakturaID
    End If

CleanExit:
    On Error Resume Next
    MarkTestPrijemnicaStornirano prijemnicaID
    On Error GoTo 0
    Exit Sub

EH:
    LogFakturaFail "CreateFaktura blocks duplicate PrijemnicaID rows", _
                   FormatErrDetails()
    Resume CleanExit
End Sub

Private Sub Test_UpdateFakturaStatusPreservesExistingDatumPlacanja()
    On Error GoTo EH

    Dim code As String
    code = NewFakturaTestCode("PAYDATE")

    Dim fakturaID As String
    Dim kupacID As String
    Dim oldDate As Date

    fakturaID = "FAK-TST-" & code
    kupacID = "KUP-TST-" & code
    oldDate = DateSerial(2020, 1, 15)

    AppendTestFakturaRow fakturaID, "TST-FAK-" & code, kupacID, 100#, STATUS_PLACENO, oldDate, ""

    Dim novacID As String
    novacID = SaveNovac( _
        "TST-UPL-" & code, Date, _
        "TEST KUPAC " & code, kupacID, "Kupac", _
        "", "", fakturaID, "", _
        NOV_KUPCI_UPLATA, _
        100#, 0#, _
        "Test payment for faktura status")

    If Len(Trim$(novacID)) = 0 Then
        LogFakturaFail "UpdateFakturaStatus preserves existing DatumPlacanja", _
                       "Failed to create payment row."
        Exit Sub
    End If

    UpdateFakturaStatus fakturaID

    Dim actualDate As Variant
    actualDate = LookupValue(TBL_FAKTURE, COL_FAK_ID, fakturaID, COL_FAK_DATUM_PLACANJA)

    If IsDate(actualDate) And CDate(actualDate) = oldDate Then
        LogFakturaPass "UpdateFakturaStatus preserves existing DatumPlacanja"
    Else
        LogFakturaFail "UpdateFakturaStatus preserves existing DatumPlacanja", _
                       "Expected=" & Format$(oldDate, "yyyy-mm-dd") & _
                       " Actual=" & CStr(actualDate)
    End If

    Exit Sub

EH:
    LogFakturaFail "UpdateFakturaStatus preserves existing DatumPlacanja", _
                   FormatErrDetails()
End Sub

Private Sub Test_UpdateFakturaStatusReopensWhenPaymentMissing()
    On Error GoTo EH

    Dim code As String
    code = NewFakturaTestCode("REOPEN")

    Dim fakturaID As String
    Dim kupacID As String

    fakturaID = "FAK-TST-" & code
    kupacID = "KUP-TST-" & code

    AppendTestFakturaRow fakturaID, "TST-FAK-" & code, kupacID, 100#, STATUS_PLACENO, Date, ""

    ' Nema aktivne uplate za ovu fakturu.
    UpdateFakturaStatus fakturaID

    AssertFakturaTextEquals STATUS_NEPLACENO, _
        CStr(LookupValue(TBL_FAKTURE, COL_FAK_ID, fakturaID, COL_FAK_STATUS)), _
        "UpdateFakturaStatus reopens unpaid faktura"

    AssertFakturaTextEquals "", _
        CStr(LookupValue(TBL_FAKTURE, COL_FAK_ID, fakturaID, COL_FAK_DATUM_PLACANJA)), _
        "UpdateFakturaStatus clears DatumPlacanja when unpaid"

    Exit Sub

EH:
    LogFakturaFail "UpdateFakturaStatus reopens when payment missing", _
                   FormatErrDetails()
End Sub

Private Sub Test_UpdateFakturaStatusSkipsStorniranaFaktura()
    On Error GoTo EH

    Dim code As String
    code = NewFakturaTestCode("SKIPSTO")

    Dim fakturaID As String
    Dim kupacID As String
    Dim oldDate As Date

    fakturaID = "FAK-TST-" & code
    kupacID = "KUP-TST-" & code
    oldDate = DateSerial(2020, 2, 20)

    AppendTestFakturaRow fakturaID, "TST-FAK-" & code, kupacID, 100#, STATUS_PLACENO, oldDate, "Da"

    UpdateFakturaStatus fakturaID

    AssertFakturaTextEquals STATUS_PLACENO, _
        CStr(LookupValue(TBL_FAKTURE, COL_FAK_ID, fakturaID, COL_FAK_STATUS)), _
        "UpdateFakturaStatus skips stornirana faktura status"

    Dim actualDate As Variant
    actualDate = LookupValue(TBL_FAKTURE, COL_FAK_ID, fakturaID, COL_FAK_DATUM_PLACANJA)

    If IsDate(actualDate) And CDate(actualDate) = oldDate Then
        LogFakturaPass "UpdateFakturaStatus skips stornirana faktura date"
    Else
        LogFakturaFail "UpdateFakturaStatus skips stornirana faktura date", _
                       "Expected=" & Format$(oldDate, "yyyy-mm-dd") & _
                       " Actual=" & CStr(actualDate)
    End If

    Exit Sub

EH:
    LogFakturaFail "UpdateFakturaStatus skips stornirana faktura", _
                   FormatErrDetails()
End Sub

Private Sub Test_PrintFakturaBlocksStorniranaFaktura()
    On Error GoTo EH

    Dim code As String
    code = NewFakturaTestCode("PRNSTO")

    Dim fakturaID As String
    Dim kupacID As String

    fakturaID = "FAK-TST-" & code
    kupacID = "KUP-TST-" & code

    AppendTestFakturaRow fakturaID, "TST-FAK-" & code, kupacID, 100#, STATUS_NEPLACENO, Empty, "Da"

    On Error GoTo ExpectedError
    PrintFaktura fakturaID

    LogFakturaFail "PrintFaktura blocks stornirana faktura", _
                   "Expected error, but PrintFaktura completed."
    Exit Sub

ExpectedError:
    LogFakturaPass "PrintFaktura blocks stornirana faktura"
    Exit Sub

EH:
    LogFakturaFail "PrintFaktura blocks stornirana faktura", _
                   FormatErrDetails()
End Sub

' AUD-027 / FM-0031 #3: reprint otkupnog lista mora da odbije storniran otkup,
' i mora da bude FAIL-CLOSED (nepoznat / dupliran ID takodje ne prolazi).
' Testira se seam `RequireOtkupAktivanZaStampu`, jer sam reprint na blokadi
' otvara MsgBox (test ne sme da ceka operatera).
' Redovi se upisuju u tblOtkup pod transakcijom i UVEK se rollback-uju.
Private Sub Test_ReprintBlocksStorniraniOtkup()
    Dim tx As clsTransaction
    On Error GoTo EH

    If GetTable(TBL_OTKUP) Is Nothing Then
        LogFakturaSkip "Reprint blocks storniran otkup", "nema tabele " & TBL_OTKUP
        Exit Sub
    End If

    Dim code As String
    code = NewFakturaTestCode("RPRSTO")

    Dim otkupStorniran As String
    Dim otkupAktivan As String
    Dim otkupDupliran As String

    otkupStorniran = "OTK-TST-S-" & code
    otkupAktivan = "OTK-TST-A-" & code
    otkupDupliran = "OTK-TST-D-" & code

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP

    AppendTestOtkupRow otkupStorniran, "TST-BRD-S-" & code, "Da"
    AppendTestOtkupRow otkupAktivan, "TST-BRD-A-" & code, ""
    AppendTestOtkupRow otkupDupliran, "TST-BRD-D1-" & code, ""
    AppendTestOtkupRow otkupDupliran, "TST-BRD-D2-" & code, ""

    AssertFakturaTrue ReprintGuardRaises(otkupStorniran), _
                      "Reprint blocks storniran otkup"

    AssertFakturaTrue Not ReprintGuardRaises(otkupAktivan), _
                      "Reprint allows aktivan otkup"

    AssertFakturaTrue ReprintGuardRaises(otkupDupliran), _
                      "Reprint blocks duplicate OtkupID (fail-closed)"

    AssertFakturaTrue ReprintGuardRaises("OTK-TST-NEMA-" & code), _
                      "Reprint blocks unknown OtkupID (fail-closed)"

    tx.RollbackTx
    Set tx = Nothing
    Exit Sub

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    Set tx = Nothing
    On Error GoTo 0

    LogFakturaFail "Reprint blocks storniran otkup", _
                   "Err.Number=" & CStr(errNum) & _
                   " Source=" & errSrc & _
                   " Description=" & errDesc
End Sub

' True = kapija je odbila stampu (podigla gresku). Zamenjuje raniji Boolean
' helper: guard je sada fail-closed, pa se ishod meri time DA LI je pukao.
Private Function ReprintGuardRaises(ByVal otkupID As String) As Boolean
    On Error Resume Next
    Err.Clear
    RequireOtkupAktivanZaStampu otkupID, "modFakturaTests"
    ReprintGuardRaises = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
End Function

' AUD-027 / FM-0031 #19, review nalaz 1: cleanup opseg mora da prati STVARAN broj
' stavki, ne fiksnih 80 redova. Renderuje se 81 pa 82 stavke -- druga faktura
' mora da prodje bez zaostalog merge-a sa offseta 81 (ranije: pisanje preko
' merged celije -> EH -> `Nothing` -> tiho izostala stampa).
' Review nalaz 2: granica NE sme da dodje iz `UsedRange` -- sablon je formatiran
' preko `ws.cells.Font...`, pa bi prazna FORMATIRANA celija duboko ispod
' dokumenta razvukla cleanup na milione celija. Zato se granica trazi po
' sadrzaju (`SablonLastContentRow`), sto se ovde i tvrdi.
' Ne dira tabele: FillFakturaSablon prima stavke kao niz.
Private Sub Test_FakturaSablonCleanupPratiBrojStavki()
    ' Konvencija: Const na vrhu procedure, ne izmedju izvrsnih linija.
    Const FAR_ROW As Long = 5000
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = FillFakturaSablon("TST-81", Date, "Test kupac", _
                               BuildFakturaTestStavke(81), 81, 81#)

    AssertFakturaTrue Not ws Is Nothing, _
                      "FillFakturaSablon renders 81 stavki"

    Set ws = FillFakturaSablon("TST-82", Date, "Test kupac", _
                               BuildFakturaTestStavke(82), 82, 82#)

    AssertFakturaTrue Not ws Is Nothing, _
                      "FillFakturaSablon renders 82 stavki after 81"

    If ws Is Nothing Then GoTo CleanExit

    Dim startCell As Range
    Set startCell = ws.Range("FakStavkaStart")

    AssertFakturaTrue Not startCell.Offset(81, 1).MergeCells, _
                      "Stavka 82 nije u zaostalom merge-u"

    AssertFakturaTextEquals "PRJ-82", _
        CStr(startCell.Offset(81, 1).value), _
        "Stavka 82 je stvarno upisana"

    AssertFakturaTextEquals "@", _
        CStr(startCell.Offset(81, 1).NumberFormat), _
        "Broj prijemnice ostaje tekst i posle 80. reda"

    ' Zaglavlje sablona ima NAMERNE merge-ove -- cleanup ne sme da ih pokida.
    AssertFakturaTrue ws.Range("FakKupac").MergeCells, _
                      "Merge zaglavlja (FakKupac) je netaknut"

    ' --- Granica ciscenja mora da dolazi iz SADRZAJA, ne iz formatiranja ---
    Dim sadrzajPre As Long
    sadrzajPre = SablonLastContentRow(ws, 1, 6)

    ' Prazna, ali FORMATIRANA celija duboko ispod dokumenta.
    ws.cells(FAR_ROW, 2).Interior.Color = RGB(255, 0, 0)

    AssertFakturaTrue SablonLastContentRow(ws, 1, 6) = sadrzajPre, _
                      "Formatirana prazna celija ne siri granicu ciscenja"

    AssertFakturaTrue sadrzajPre < FAR_ROW, _
                      "Granica ciscenja ostaje kod stvarnog sadrzaja"

    ' Render posle toga ne sme da dodirne taj red (dokaz da cleanup nije
    ' razvucen do FAR_ROW: boja bi bila obrisana `Interior.ColorIndex = xlNone`).
    Set ws = FillFakturaSablon("TST-3", Date, "Test kupac", _
                               BuildFakturaTestStavke(3), 3, 3#)

    AssertFakturaTrue Not ws Is Nothing, _
                      "FillFakturaSablon renders 3 stavke after 82"

    If ws Is Nothing Then GoTo CleanExit

    AssertFakturaTrue ws.cells(FAR_ROW, 2).Interior.ColorIndex <> xlNone, _
                      "Cleanup nije razvucen do formatirane celije na " & CStr(FAR_ROW)

CleanExit:
    ' Vrati sablon u zateceno stanje: obrisan list `EnsureFakturaSablon`
    ' regenerise pri sledecoj stampi, pa test ne ostavlja ni test stavke ni
    ' potpise ni PrintArea ni onu formatiranu celiju na 5000.
    On Error Resume Next
    Dim cleanupWs As Worksheet
    Set cleanupWs = ThisWorkbook.Sheets(WS_FAKTURA_SABLON)
    If Not cleanupWs Is Nothing Then
        Application.DisplayAlerts = False
        cleanupWs.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0
    Exit Sub

EH:
    LogFakturaFail "FillFakturaSablon cleanup prati broj stavki", _
                   FormatErrDetails()
    Resume CleanExit
End Sub

Private Function BuildFakturaTestStavke(ByVal n As Long) As Variant
    Dim arr() As Variant
    ReDim arr(1 To n, 1 To 5)

    Dim i As Long
    For i = 1 To n
        arr(i, 1) = "PRJ-" & CStr(i)
        arr(i, 2) = "I"
        arr(i, 3) = 1#
        arr(i, 4) = 1#
        arr(i, 5) = 1#
    Next i

    BuildFakturaTestStavke = arr
End Function

Private Sub AppendTestOtkupRow(ByVal otkupID As String, _
                               ByVal brojDokumenta As String, _
                               ByVal stornirano As String)
    Const SRC As String = "AppendTestOtkupRow"

    Dim values As Object
    Set values = CreateObject("Scripting.Dictionary")

    values.Add COL_OTK_ID, otkupID
    values.Add COL_OTK_BR_DOK, brojDokumenta
    values.Add COL_OTK_DATUM, Date
    values.Add COL_OTK_KLASA, "I"
    values.Add COL_OTK_KOLICINA, 100#
    values.Add COL_OTK_CENA, 10#
    values.Add COL_STORNIRANO, stornirano

    AppendFakturaTestRowByColumnMap TBL_OTKUP, values, SRC
End Sub

Private Function NewFakturaTestCode(ByVal prefixText As String) As String
    Randomize
    NewFakturaTestCode = prefixText & "-" & _
                         Format$(Now, "yyyymmddhhnnss") & "-" & _
                         CStr(Int((9000 * Rnd) + 1000))
End Function

' kupacID je obavezan: CreateFaktura od RF-08 proverava da prijemnica pripada
' kupcu fakture (AUD-011 / FM-0034 #1), pa test red bez KupacID-a vise ne prolazi.
Private Sub AppendTestPrijemnicaRow(ByVal prijemnicaID As String, _
                                    ByVal kupacID As String, _
                                    ByVal brojPrijemnice As String, _
                                    ByVal kolicina As Double, _
                                    ByVal cena As Double, _
                                    ByVal klasa As String, _
                                    ByVal stornirano As String, _
                                    ByVal fakturisano As String, _
                                    ByVal fakturaID As String)
    Const SRC As String = "AppendTestPrijemnicaRow"

    Dim values As Object
    Set values = CreateObject("Scripting.Dictionary")

    values.Add COL_PRJ_ID, prijemnicaID
    values.Add COL_PRJ_KUPAC, kupacID
    values.Add COL_PRJ_BROJ, brojPrijemnice
    values.Add COL_PRJ_BROJ_ZBIRNE, "TST-ZBR-" & prijemnicaID
    values.Add COL_PRJ_DATUM, Date
    values.Add COL_PRJ_KOLICINA, kolicina
    values.Add COL_PRJ_CENA, cena
    values.Add COL_PRJ_KLASA, klasa
    values.Add COL_PRJ_FAKTURISANO, fakturisano
    values.Add COL_PRJ_FAKTURA_ID, fakturaID
    values.Add COL_STORNIRANO, stornirano

    AppendFakturaTestRowByColumnMap TBL_PRIJEMNICA, values, SRC
End Sub

Private Sub AppendTestFakturaRow(ByVal fakturaID As String, _
                                 ByVal brojFakture As String, _
                                 ByVal kupacID As String, _
                                 ByVal iznos As Double, _
                                 ByVal statusText As String, _
                                 ByVal datumPlacanja As Variant, _
                                 ByVal stornirano As String)
    Const SRC As String = "AppendTestFakturaRow"

    Dim values As Object
    Set values = CreateObject("Scripting.Dictionary")

    values.Add COL_FAK_ID, fakturaID
    values.Add COL_FAK_BROJ, brojFakture
    values.Add COL_FAK_DATUM, Date
    values.Add COL_FAK_KUPAC, kupacID
    values.Add COL_FAK_IZNOS, iznos
    values.Add COL_FAK_STATUS, statusText
    values.Add COL_FAK_DATUM_PLACANJA, datumPlacanja
    values.Add COL_STORNIRANO, stornirano

    AppendFakturaTestRowByColumnMap TBL_FAKTURE, values, SRC
End Sub

Private Sub AppendFakturaTestRowByColumnMap(ByVal tableName As String, _
                                            ByVal values As Object, _
                                            ByVal sourceName As String)
    Dim colCount As Long
    colCount = GetFakturaTestTableColumnCount(tableName)

    If colCount <= 0 Then
        Err.Raise vbObjectError + 9201, sourceName, _
                  "Could not resolve table column count. Table=" & tableName
    End If

    Dim rowData() As Variant
    ReDim rowData(0 To colCount - 1)

    Dim key As Variant
    Dim colIndex As Long

    For Each key In values.keys
        colIndex = GetColumnIndex(tableName, CStr(key))

        If colIndex > 0 Then
            rowData(colIndex - 1) = values(key)
        End If
    Next key

    If AppendRow(tableName, rowData) <= 0 Then
        Err.Raise vbObjectError + 9202, sourceName, _
                  "Failed to append test row. Table=" & tableName
    End If
End Sub

Private Function GetFakturaTestTableColumnCount(ByVal tableName As String) As Long
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.name, tableName, vbTextCompare) = 0 Then
                GetFakturaTestTableColumnCount = lo.ListColumns.count
                Exit Function
            End If
        Next lo
    Next ws
End Function

Private Function GetFirstFakturaStavkaValue(ByVal fakturaID As String, _
                                            ByVal columnName As String) As Variant
    Const SRC As String = "GetFirstFakturaStavkaValue"

    Dim data As Variant
    data = GetTableData(TBL_FAKTURA_STAVKE)

    If IsEmpty(data) Then
        Err.Raise vbObjectError + 9203, SRC, _
                  "tblFakturaStavke is empty."
    End If

    Dim colFakID As Long
    Dim colWanted As Long

    colFakID = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_FAKTURA_ID, SRC)
    colWanted = RequireColumnIndex(TBL_FAKTURA_STAVKE, columnName, SRC)

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, colFakID))) = Trim$(fakturaID) Then
            GetFirstFakturaStavkaValue = data(i, colWanted)
            Exit Function
        End If
    Next i

    Err.Raise vbObjectError + 9204, SRC, _
              "No faktura stavka found for FakturaID=" & fakturaID
End Function

Private Sub AssertFakturaTrue(ByVal condition As Boolean, _
                              ByVal testName As String)
    If condition Then
        LogFakturaPass testName
    Else
        LogFakturaFail testName, "Assertion failed."
    End If
End Sub

Private Sub AssertFakturaDoubleEquals(ByVal expectedValue As Double, _
                                      ByVal actualValue As Double, _
                                      ByVal testName As String)
    If Abs(expectedValue - actualValue) <= 0.0001 Then
        LogFakturaPass testName
    Else
        LogFakturaFail testName, _
                       "Expected=" & CStr(expectedValue) & _
                       " Actual=" & CStr(actualValue)
    End If
End Sub

Private Sub AssertFakturaTextEquals(ByVal expectedValue As String, _
                                    ByVal actualValue As String, _
                                    ByVal testName As String)
    If Trim$(CStr(expectedValue)) = Trim$(CStr(actualValue)) Then
        LogFakturaPass testName
    Else
        LogFakturaFail testName, _
                       "Expected=[" & expectedValue & "] Actual=[" & actualValue & "]"
    End If
End Sub

Private Function FormatErrDetails() As String
    FormatErrDetails = "Err.Number=" & CStr(Err.Number) & _
                       " Source=" & Err.SOURCE & _
                       " Description=" & Err.description
End Function

Private Sub ResetFakturaTestCounters()
    m_Total = 0
    m_Passed = 0
    m_Failed = 0
    m_Skipped = 0
End Sub

Private Sub StartFakturaSuite(ByVal suiteName As String)
    Debug.Print String$(70, "=")
    Debug.Print suiteName & " started at " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Debug.Print String$(70, "=")

    AppendFakturaTestLog "SUITE", suiteName, "START", ""
End Sub

Private Sub FinishFakturaSuite()
    Dim summary As String

    summary = "Total=" & m_Total & _
              " | Passed=" & m_Passed & _
              " | Failed=" & m_Failed & _
              " | Skipped=" & m_Skipped

    Debug.Print String$(70, "-")
    Debug.Print "FAKTURA TEST SUMMARY: " & summary
    Debug.Print String$(70, "-")

    AppendFakturaTestLog "SUITE", "SUMMARY", "INFO", summary

    If m_Failed > 0 Then
        MsgBox "Faktura tests finished with failures." & vbCrLf & summary, _
               vbExclamation, APP_NAME
    Else
        MsgBox "Faktura tests finished." & vbCrLf & summary, _
               vbInformation, APP_NAME
    End If
End Sub

Private Sub LogFakturaPass(ByVal testName As String)
    m_Total = m_Total + 1
    m_Passed = m_Passed + 1

    Debug.Print "[PASS] " & testName
    AppendFakturaTestLog "TEST", testName, "PASS", ""
End Sub

Private Sub LogFakturaFail(ByVal testName As String, _
                           ByVal details As String)
    m_Total = m_Total + 1
    m_Failed = m_Failed + 1

    Debug.Print "[FAIL] " & testName & " :: " & details
    AppendFakturaTestLog "TEST", testName, "FAIL", details
End Sub

Private Sub LogFakturaSkip(ByVal testName As String, _
                           ByVal details As String)
    m_Total = m_Total + 1
    m_Skipped = m_Skipped + 1

    Debug.Print "[SKIP] " & testName & " :: " & details
    AppendFakturaTestLog "TEST", testName, "SKIP", details
End Sub

Private Sub LogFakturaFatal(ByVal sourceName As String, _
                            ByVal errNum As Long, _
                            ByVal errDesc As String)
    m_Total = m_Total + 1
    m_Failed = m_Failed + 1

    Debug.Print "[FATAL] " & sourceName & " :: " & CStr(errNum) & " - " & errDesc
    AppendFakturaTestLog "FATAL", sourceName, "FAIL", CStr(errNum) & " - " & errDesc
End Sub

Private Sub InitFakturaTestLog()
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(FAKTURA_TEST_LOG_SHEET)

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = FAKTURA_TEST_LOG_SHEET
        ws.Range("A1:F1").value = Array("Timestamp", "Kind", "Name", "Status", "Details", "Operator")
        ws.rows(1).Font.Bold = True
    End If
End Sub

Private Sub AppendFakturaTestLog(ByVal kindText As String, _
                                 ByVal nameText As String, _
                                 ByVal statusText As String, _
                                 ByVal detailsText As String)
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(FAKTURA_TEST_LOG_SHEET)

    If ws Is Nothing Then Exit Sub

    Dim r As Long
    r = ws.cells(ws.rows.count, 1).End(xlUp).row + 1

    ws.cells(r, 1).value = Now
    ws.cells(r, 2).value = kindText
    ws.cells(r, 3).value = nameText
    ws.cells(r, 4).value = statusText
    ws.cells(r, 5).value = Left$(detailsText, 2000)
    ws.cells(r, 6).value = Environ$("Username")
End Sub
