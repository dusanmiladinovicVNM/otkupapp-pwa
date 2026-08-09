Attribute VB_Name = "modTestBanka"
Option Explicit

' ============================================================
' modTestBanka - tvrde kapije za RF-09 (banka import + mapiranje)
'
' Pokretanje: Alt+F8 -> RunBankaImportTestSuite
'
' Sopstveni harness (mPass/mFail/mReport + Chk/ChkEq), po uzoru na
' modTestStorno/modTestPalete. Testovi koji diraju tabele rade u JEDNOJ
' transakciji koja se UVEK rollback-uje -> nula ostavljenih podataka. Inner *_TX
' commituju svoje TX; spoljni snapshot-restore ih ponisti.
'
' Prefiks svih test podataka: "BIT-" (Banka-Import-Test) -> izolacija.
'
' Napomena: T03 pusta pravi AutoMapAll batch, pa se salju monitoring dogadjaji
' (BANKA_AUTOMAP_ALL_*) ako je monitoring ukljucen -- to je telemetrija, ne podatak
' u tabelama. Journal se za vreme suite-a stisava (SetTestModeQuiet).
'
' Pokriva cetiri nalaza koje RF-09 zatvara:
'  T01 AUD-007  nemoguc datum (30.02., dan 32, mesec 13) se ODBIJA, ne pomera
'  T02 AUD-025  dedupe kljuc ukljucuje broj racuna (multi-account transakcija)
'  T03 AUD-025  3+ kandidata bloka obara SAMO taj red, ne ceo AutoMapAll batch
'  T04 AUD-025  rucno mapiranje pogresnog smera je odbijeno
' ============================================================

Private mPass As Long
Private mFail As Long
Private mFails As String
Private mReport As String

Private Const P As String = "BIT-"

Public Sub RunBankaImportTestSuite()
    Dim tx As clsTransaction
    Dim wasQuiet As Boolean
    Dim quietSet As Boolean

    On Error GoTo EH

    If GetTable(TBL_BANKA_IMPORT) Is Nothing Or GetTable(TBL_OTKUP) Is Nothing Then
        MsgBox "Tabele tblBankaImport/tblOtkup ne postoje. Prekid.", vbExclamation, APP_NAME
        Exit Sub
    End If

    If MsgBox("Pokrenuti banka import/mapiranje test suite (RF-09)?" & vbCrLf & vbCrLf & _
              "Svi podaci (BIT-*) se prave u transakciji i UVEK se ponistavaju " & _
              "(rollback). Nista ne ostaje u tabelama.", _
              vbQuestion + vbYesNo, APP_NAME) <> vbYes Then Exit Sub

    mPass = 0: mFail = 0: mFails = "": mReport = ""

    ' T01 je cist parser (bez tabela) -> moze i van transakcije.
    T01_NemoguciDatumOdbijen

    ' AppendRow/UpdateCell pisu CSV crash-recovery journal koji tx.RollbackTx NE
    ' povlaci -- test redovi bi ostali u Journal folderu i sledeci start bi javio
    ' lazno upozorenje o gubitku podataka. Isti obrazac kao modSEFTests.
    wasQuiet = modJournaling.IsTestModeQuiet()
    modJournaling.SetTestModeQuiet True
    quietSet = True

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_BANKA_IMPORT
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_KOOPERANTI
    tx.AddTableSnapshot TBL_STANICE
    tx.AddTableSnapshot TBL_PARTNER_MAP
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot TBL_KUPCI

    T02_DedupeUkljucujeBrojRacuna
    T03_TriKandidataNeObarajuBatch
    T04_SmerGuardOdbijaPogresanTip

    tx.RollbackTx
    Set tx = Nothing

    RestoreJournalQuiet quietSet, wasQuiet

    ReportResults
    Exit Sub

EH:
    Dim errDesc As String
    errDesc = Err.description

    On Error Resume Next
    LogErr "modTestBanka.RunBankaImportTestSuite"
    If Not tx Is Nothing Then tx.RollbackTx
    gBankaSilentBatch = False
    RestoreJournalQuiet quietSet, wasQuiet
    On Error GoTo 0

    Fail "SUITE prekinut greskom: " & errDesc
    ReportResults
End Sub

' Vraca journal test-mode na ZATECENO stanje (ne bezuslovno False) i otkazuje
' zakazan AutoSave tick -- posle rollback-a nema sta da se snima.
Private Sub RestoreJournalQuiet(ByVal wasSet As Boolean, ByVal previousValue As Boolean)
    On Error Resume Next
    If wasSet Then
        modJournaling.SetTestModeQuiet previousValue
        modJournaling.StopAutoSaveTimer
    End If
    On Error GoTo 0
End Sub

' ============================================================
' T01 - AUD-007 (P0): TryParseDateValue i nemoguc datum.
'
' DateSerial(2026, 2, 30) ne puca nego vraca 02.03.2026. Bez round-trip provere
' bi datum transakcije iz izvoda tiho zavrsio u sledecem mesecu.
' ============================================================
Private Sub T01_NemoguciDatumOdbijen()
    Const S As String = "T01 nemoguc datum: "

    Dim d As Date

    Chk Not TryParseDateValue("30.02.2026", d), S & "30.02.2026 odbijen (ne prelije se u mart)"
    Chk Not TryParseDateValue("31.04.2026", d), S & "31.04.2026 odbijen (april ima 30 dana)"
    Chk Not TryParseDateValue("32.01.2026", d), S & "dan 32 odbijen"
    Chk Not TryParseDateValue("01.13.2026", d), S & "mesec 13 odbijen"
    Chk Not TryParseDateValue("29.02.2026", d), S & "29.02. u neprestupnoj godini odbijen"
    Chk Not TryParseDateValue("00.01.2026", d), S & "dan 0 odbijen"

    ' Validni datumi moraju i dalje da prolaze, i to TACNO.
    d = 0
    Chk TryParseDateValue("29.02.2024", d), S & "29.02.2024 (prestupna) prolazi"
    ChkEq Format$(d, "yyyy-mm-dd"), "2024-02-29", S & "29.02.2024 tacno parsiran"

    d = 0
    Chk TryParseDateValue("31.12.2026", d), S & "31.12.2026 prolazi"
    ChkEq Format$(d, "yyyy-mm-dd"), "2026-12-31", S & "31.12.2026 tacno parsiran"

    d = 0
    Chk TryParseDateValue("1.2.26", d), S & "dvocifrena godina prolazi"
    ChkEq Format$(d, "yyyy-mm-dd"), "2026-02-01", S & "dvocifrena godina -> 2026"

    d = 0
    Chk TryParseDateValue("15/03/2026", d), S & "kosa crta kao separator prolazi"
    ChkEq Format$(d, "yyyy-mm-dd"), "2026-03-15", S & "15/03/2026 tacno parsiran"
End Sub

' ============================================================
' T02 - AUD-025: dedupe kljuc mora da sadrzi broj racuna.
'
' Broj izvoda je jedinstven PO RACUNU. Bez racuna u kljucu je ista transakcija na
' drugom racunu firme tiho odbacena kao duplikat i nikad ne stigne u staging.
' ============================================================
Private Sub T02_DedupeUkljucujeBrojRacuna()
    Const S As String = "T02 dedupe + broj racuna: "

    Dim dTx As Date
    dTx = Date

    ' Ista transakcija (broj izvoda, datum, iznos, partner) na racunu 1.
    SeedBim P & "BIM-D1", P & "IZV-7", P & "RAC-1", P & "PARTNER-D", 5000, 0, "", "", ""

    Chk IsDuplicateBankaImport(P & "IZV-7", dTx, 5000, 0, P & "PARTNER-D", "", P & "RAC-1"), _
        S & "isti racun -> duplikat"

    Chk Not IsDuplicateBankaImport(P & "IZV-7", dTx, 5000, 0, P & "PARTNER-D", "", P & "RAC-2"), _
        S & "drugi racun -> NIJE duplikat (transakcija se uvozi)"

    ' Ista provera i na jakoj grani kljuca (BankaReferenz).
    SeedBim P & "BIM-D2", P & "IZV-8", P & "RAC-1", P & "PARTNER-D", 0, 900, "", "", P & "REF-1"

    Chk IsDuplicateBankaImport(P & "IZV-8", dTx, 0, 900, P & "PARTNER-D", P & "REF-1", P & "RAC-1"), _
        S & "isti racun + ista referenca -> duplikat"

    Chk Not IsDuplicateBankaImport(P & "IZV-8", dTx, 0, 900, P & "PARTNER-D", P & "REF-1", P & "RAC-2"), _
        S & "drugi racun + ista referenca -> NIJE duplikat"
End Sub

' ============================================================
' T03 - AUD-025: blok sa 3+ otvorenih stavki.
'
' Ranije: ReDim(1 To 2) + count=3 -> "Subscript out of range" iz AutoMapAll ->
' rollback CELOG batch-a (i vec mapirani redovi se ponistavaju).
' Sada: jasna greska ERR_BMAP_MANUAL_REQUIRED koja obara SAMO taj red.
' ============================================================
Private Sub T03_TriKandidataNeObarajuBatch()
    Const S As String = "T03 3+ kandidata: "

    Dim errNum As Long
    Dim mapped As Long
    Dim manualRequired As Long
    Dim dummy As Variant

    SeedStanica P & "OM-1", P & "Stanica 1"
    SeedKooperant P & "K-1", "Test", "Kooperant", P & "OM-1"

    ' Blok sa TRI otvorene stavke -> automatska raspodela se ne pogadja.
    SeedOtkup P & "OTK-A", P & "K-1", P & "BLOK-3K", 100, 10, "Malina"
    SeedOtkup P & "OTK-B", P & "K-1", P & "BLOK-3K", 100, 12, "Kupina"
    SeedOtkup P & "OTK-C", P & "K-1", P & "BLOK-3K", 100, 14, "Visnja"

    ' Blok sa JEDNOM otvorenom stavkom -> mora proci u istom batch-u.
    SeedOtkup P & "OTK-OK", P & "K-1", P & "BLOK-1K", 100, 20, "Malina"

    ' 1) Resolver kandidata dize jasnu, prepoznatljivu gresku.
    On Error Resume Next
    dummy = GetOtkupCandidatesForKooperantBlock(P & "K-1", P & "BLOK-3K")
    errNum = Err.Number
    Err.Clear
    On Error GoTo 0

    ChkEq errNum, ERR_BMAP_MANUAL_REQUIRED, S & "3 kandidata -> ERR_BMAP_MANUAL_REQUIRED"

    ' 2) Blok sa 2 kandidata i dalje mora da radi (granica se ne pomera).
    On Error Resume Next
    dummy = GetOtkupCandidatesForKooperantBlock(P & "K-1", P & "BLOK-1K")
    errNum = Err.Number
    Err.Clear
    On Error GoTo 0

    ChkEq errNum, 0, S & "1 kandidat -> bez greske"

    ' 3) Batch: jedan red trazi rucno, drugi mora biti mapiran (bez rollback-a svega).
    SeedBim P & "BIM-3K", P & "IZV-9", P & "RAC-1", P & "PARTNER-K", 0, 3000, P & "BLOK-3K", "", ""
    SeedBim P & "BIM-OK", P & "IZV-9", P & "RAC-1", P & "PARTNER-K", 0, 2000, P & "BLOK-1K", "", ""

    ' Postojeci backlog se sklanja sa puta da batch bude deterministicki
    ' (rollback suite-a vraca originalne statuse).
    SkipPostojeceOtvorene

    gBankaSilentBatch = True
    mapped = AutoMapAllBankaImport_TX(manualRequired)
    gBankaSilentBatch = False

    ChkEq BimObradjeno(P & "BIM-OK"), "Da", S & "zdrav red je mapiran (batch nije rollback-ovan)"
    ChkEq BimObradjeno(P & "BIM-3K"), "Error", S & "anomalan red je oznacen za rucno"
    Chk mapped >= 1, S & "batch prijavio bar jedno mapiranje [mapped=" & CStr(mapped) & "]"
    Chk manualRequired >= 1, S & "batch prijavio 'za rucno' [manualRequired=" & CStr(manualRequired) & "]"
    Chk NovacZaBim(P & "BIM-OK") > 0, S & "zdrav red ima red(ove) u tblNovac"
    Chk NovacZaBim(P & "BIM-3K") = 0, S & "anomalan red NEMA parcijalno knjizenje"
End Sub

' ============================================================
' T04 - AUD-025: rucno mapiranje ne sme da ignorise smer.
'
' Poziva se NE-TX varijanta (isti kod, bez MsgBox-a iz TX omotaca) da bi test bio
' neinteraktivan; TX omotac istu gresku pokazuje operateru i vraca promene.
' ============================================================
Private Sub T04_SmerGuardOdbijaPogresanTip()
    Const S As String = "T04 smer guard: "

    Dim errNum As Long
    Dim errDesc As String
    Dim res As String
    Dim n As Long

    SeedStanica P & "OM-2", P & "Stanica 2"
    SeedKooperant P & "K-2", "Test", "Smer", P & "OM-2"
    SeedKupac P & "KUP-1", P & "Kupac 1"

    SeedBim P & "BIM-UPL", P & "IZV-10", P & "RAC-1", P & "PARTNER-S", 4000, 0, "", "", ""
    SeedBim P & "BIM-ISP", P & "IZV-10", P & "RAC-1", P & "PARTNER-S", 0, 4000, "", "", ""

    ' Uplata knjizena kao kooperant (isplata) -> odbijeno.
    errDesc = ""
    On Error Resume Next
    n = MapBankaImportAsKooperantBlock(P & "BIM-UPL", P & "K-2", False)
    errNum = Err.Number
    errDesc = Err.description
    Err.Clear
    On Error GoTo 0

    Chk errNum <> 0, S & "uplata + tip Kooperant -> odbijeno"
    Chk InStr(1, errDesc, "Smer ne odgovara", vbTextCompare) > 0, _
        S & "razlog odbijanja je smer [" & errDesc & "]"
    ChkEq BimObradjeno(P & "BIM-UPL"), "", S & "odbijena uplata ostaje otvorena"
    ChkEq NovacZaBim(P & "BIM-UPL"), 0, S & "odbijena uplata nije knjizena"

    ' Isplata knjizena kao kupac (uplata) -> odbijeno.
    errDesc = ""
    On Error Resume Next
    res = MapBankaImportAsKupac(P & "BIM-ISP", P & "KUP-1", "", False)
    errNum = Err.Number
    errDesc = Err.description
    Err.Clear
    On Error GoTo 0

    Chk errNum <> 0, S & "isplata + tip Kupac -> odbijeno"
    Chk InStr(1, errDesc, "Smer ne odgovara", vbTextCompare) > 0, _
        S & "razlog odbijanja je smer [" & errDesc & "]"
    ChkEq BimObradjeno(P & "BIM-ISP"), "", S & "odbijena isplata ostaje otvorena"
    ChkEq NovacZaBim(P & "BIM-ISP"), 0, S & "odbijena isplata nije knjizena"
    ChkEq res, "", S & "odbijeno mapiranje nije vratilo NovacID"
End Sub

' ============================================================
' SEED / READ HELPERS
' ============================================================

Private Sub SeedBim(ByVal bimID As String, ByVal brojIzvoda As String, _
                    ByVal racun As String, ByVal partner As String, _
                    ByVal uplata As Double, ByVal isplata As Double, _
                    ByVal poziv As String, ByVal konto As String, _
                    ByVal referenz As String)
    BitAppend TBL_BANKA_IMPORT, _
        Array(COL_BIM_ID, COL_BIM_BROJ_DOKUMENTA, COL_BIM_BROJ_RACUNA, COL_BIM_DATUM_TRANSAKCIJE, _
              COL_BIM_DATUM_IZVODA, COL_BIM_PARTNER, COL_BIM_PARTNER_KONTO, COL_BIM_UPLATA, _
              COL_BIM_ISPLATA, COL_BIM_POZIV_NA_BROJ, COL_BIM_BANKA_REFERENZ, COL_BIM_OBRADJENO), _
        Array(bimID, brojIzvoda, racun, Date, _
              Date, partner, konto, uplata, _
              isplata, poziv, referenz, "")
End Sub

Private Sub SeedOtkup(ByVal otkID As String, ByVal koopID As String, _
                      ByVal brDok As String, ByVal kolicina As Double, _
                      ByVal cena As Double, ByVal vrsta As String)
    BitAppend TBL_OTKUP, _
        Array(COL_OTK_ID, COL_OTK_BR_DOK, COL_OTK_KOOPERANT, COL_OTK_KOLICINA, _
              COL_OTK_CENA, COL_OTK_VRSTA, COL_OTK_DATUM), _
        Array(otkID, brDok, koopID, kolicina, cena, vrsta, Date)
End Sub

Private Sub SeedKooperant(ByVal koopID As String, ByVal ime As String, _
                          ByVal prezime As String, ByVal stanicaID As String)
    BitAppend TBL_KOOPERANTI, _
        Array("KooperantID", "Ime", "Prezime", COL_KOOP_STANICA), _
        Array(koopID, ime, prezime, stanicaID)
End Sub

Private Sub SeedKupac(ByVal kupacID As String, ByVal naziv As String)
    BitAppend TBL_KUPCI, _
        Array("KupacID", "Naziv"), _
        Array(kupacID, naziv)
End Sub

Private Sub SeedStanica(ByVal stanicaID As String, ByVal naziv As String)
    BitAppend TBL_STANICE, _
        Array("StanicaID", "Naziv"), _
        Array(stanicaID, naziv)
End Sub

Private Sub BitAppend(ByVal tblName As String, ByVal cols As Variant, ByVal vals As Variant)
    Dim lo As ListObject
    Set lo = GetTable(tblName)
    If lo Is Nothing Then
        Err.Raise vbObjectError + 2960, "modTestBanka.BitAppend", "Nema tabele: " & tblName
    End If

    Dim nr As ListRow
    Set nr = lo.ListRows.Add

    Dim i As Long
    Dim ci As Long

    For i = LBound(cols) To UBound(cols)
        ci = GetColumnIndex(tblName, CStr(cols(i)))
        If ci > 0 Then nr.Range.cells(1, ci).value = vals(i)
    Next i
End Sub

' Sve zatecene OTVORENE staging redove (koji nisu BIT-*) privremeno oznaci kao
' "Skip" da batch test radi samo nad test podacima. Rollback suite-a ih vraca.
Private Sub SkipPostojeceOtvorene()
    Dim data As Variant
    Dim colID As Long
    Dim colObr As Long
    Dim i As Long
    Dim bimID As String

    data = GetTableData(TBL_BANKA_IMPORT)
    If IsEmpty(data) Then Exit Sub

    colID = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_ID)
    colObr = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_OBRADJENO)

    For i = 1 To UBound(data, 1)
        bimID = Trim$(CStr(data(i, colID)))

        If Left$(bimID, Len(P)) <> P Then
            If Trim$(CStr(NzTb(data(i, colObr)))) = "" Then
                UpdateCell TBL_BANKA_IMPORT, i, COL_BIM_OBRADJENO, "Skip"
            End If
        End If
    Next i
End Sub

Private Function BimObradjeno(ByVal bimID As String) As String
    BimObradjeno = Trim$(CStr(NzTb(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bimID, COL_BIM_OBRADJENO))))
End Function

' Broj nestorniranih redova u tblNovac koji nose BIM marker ovog staging reda.
Private Function NovacZaBim(ByVal bimID As String) As Long
    Dim data As Variant
    Dim colNap As Long
    Dim i As Long

    data = GetTableData(TBL_NOVAC)
    If IsEmpty(data) Then Exit Function

    data = ExcludeStornirano(data, TBL_NOVAC)
    If IsEmpty(data) Then Exit Function

    colNap = GetColumnIndex(TBL_NOVAC, COL_NOV_NAPOMENA)

    For i = 1 To UBound(data, 1)
        If BimIdFromNapomena(CStr(NzTb(data(i, colNap)))) = bimID Then
            NovacZaBim = NovacZaBim + 1
        End If
    Next i
End Function

Private Function NzTb(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then NzTb = "" Else NzTb = CStr(v)
End Function

' ============================================================
' ASSERT + REPORT (isti stil kao modTestStorno)
' ============================================================

Private Sub Chk(ByVal cond As Boolean, ByVal nm As String)
    If cond Then
        mPass = mPass + 1
        mReport = mReport & "OK    " & nm & vbCrLf
    Else
        Fail nm
    End If
End Sub

Private Sub ChkEq(ByVal act As Variant, ByVal exp As Variant, ByVal nm As String)
    If CStr(act) = CStr(exp) Then
        mPass = mPass + 1
        mReport = mReport & "OK    " & nm & vbCrLf
    Else
        Fail nm & " [dobijeno=" & CStr(act) & " ocekivano=" & CStr(exp) & "]"
    End If
End Sub

Private Sub Fail(ByVal nm As String)
    mFail = mFail + 1
    mFails = mFails & " - " & nm & vbCrLf
    mReport = mReport & "PAO   " & nm & vbCrLf
End Sub

Private Sub ReportResults()
    Dim hdr As String
    hdr = "BANKA IMPORT TEST SUITE (RF-09)  ->  PASS=" & mPass & "  FAIL=" & mFail

    Debug.Print String(60, "=")
    Debug.Print hdr
    Debug.Print String(60, "=")
    Debug.Print mReport

    Dim msg As String
    msg = hdr & vbCrLf & vbCrLf

    If mFail > 0 Then
        msg = msg & "PALI TESTOVI:" & vbCrLf & mFails & vbCrLf
    End If

    msg = msg & "Detalji: Immediate prozor (Ctrl+G)."

    If mFail > 0 Then
        MsgBox msg, vbCritical, APP_NAME
    Else
        MsgBox msg, vbInformation, APP_NAME
    End If
End Sub
