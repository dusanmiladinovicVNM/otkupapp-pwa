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

' Tvrd gate: posle izvestaja suite PODIZE gresku ako je ijedna provera pala, da
' automatizovan pozivalac (ne samo operater koji gleda MsgBox) vidi neuspeh.
Private Const ERR_BIT_SUITE_FAILED As Long = vbObjectError + 2961

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
    T05_StagingCuvaTypedDatum
    T06_UplataPrekoOtvorenogSeDeli

    tx.RollbackTx
    Set tx = Nothing

    RestoreJournalQuiet quietSet, wasQuiet

    On Error GoTo 0
    ReportResults

    If mFail > 0 Then
        Err.Raise ERR_BIT_SUITE_FAILED, "modTestBanka.RunBankaImportTestSuite", _
            "RunBankaImportTestSuite: " & CStr(mFail) & " provera palo (PASS=" & _
            CStr(mPass) & "). Detalji u Immediate prozoru."
    End If

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

    Err.Raise ERR_BIT_SUITE_FAILED, "modTestBanka.RunBankaImportTestSuite", _
        "RunBankaImportTestSuite prekinut: " & errDesc
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

    ' Deklarisan opseg godina vazi za OBE grane parsera -- i za onu koju VBA sam
    ' prihvati (IsDate/CDate), inace bi 1899. i 2200. prosle na mala vrata.
    Chk Not TryParseDateValue("01.01.1899", d), S & "godina 1899 odbijena (van opsega)"
    Chk Not TryParseDateValue("01.01.2200", d), S & "godina 2200 odbijena (van opsega)"
    Chk Not TryParseDateValue("1899-01-01", d), S & "1899 odbijena i kroz IsDate granu"
    Chk Not TryParseDateValue("12:30", d), S & "samo-vreme odbijeno (daje 1899-12-30)"

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

    ' Locale kapija: dd.mm.yyyy iz izvoda banke mora da se parsira DETERMINISTICKI
    ' (dan pa mesec), bez oslanjanja na CDate. Na MM/DD masini bi "01.02.2026"
    ' inace postalo 2. januar umesto 1. februara.
    d = 0
    Chk TryParseBankaDateDMY("01.02.2026", d), S & "DMY parser prihvata 01.02.2026"
    ChkEq Format$(d, "yyyy-mm-dd"), "2026-02-01", S & "01.02.2026 = 1. februar (ne 2. januar)"

    d = 0
    Chk TryParseDateValue("01.02.2026", d), S & "TryParseDateValue prihvata 01.02.2026"
    ChkEq Format$(d, "yyyy-mm-dd"), "2026-02-01", S & "TryParseDateValue ide DMY granom prva"

    d = 0
    Chk TryParseDateValue("13.01.2026", d), S & "13.01.2026 prolazi (dan > 12)"
    ChkEq Format$(d, "yyyy-mm-dd"), "2026-01-13", S & "13.01.2026 tacno parsiran"

    Chk Not TryParseBankaDateDMY("2026-02-01", d), S & "ISO oblik nije DMY (ide na fallback)"
    Chk Not TryParseBankaDateDMY("30.02.2026", d), S & "DMY parser odbija nemoguc datum"
End Sub

' ============================================================
' T05 - staging cuva TYPED datum (a ne originalni tekst).
'
' Validiran Date je ranije bio odbacen: rezultat parsera je vracao sirovi tekst,
' staging ga upisivao kroz CStr, a mapiranje ga opet tumacilo locale-zavisnim
' CDate. Sada `SaveBankaImportRows` upisuje Date serial.
' ============================================================
Private Sub T05_StagingCuvaTypedDatum()
    Const S As String = "T05 typed datum u stagingu: "

    Dim red(1 To 1, 1 To 17) As Variant
    Dim saved As Long
    Dim v As Variant

    red(1, 1) = P & "IZV-11"          ' BrojDokumenta
    red(1, 2) = DateSerial(2026, 2, 1) ' DatumIzvoda (typed)
    red(1, 3) = P & "RAC-9"           ' BrojRacuna
    red(1, 4) = DateSerial(2026, 2, 1) ' DatumTransakcije (typed)
    red(1, 5) = P & "PARTNER-T"       ' Partner
    red(1, 6) = ""                    ' PartnerKonto
    red(1, 7) = 1234.56               ' Uplata
    red(1, 8) = 0                     ' Isplata
    red(1, 9) = ""                    ' Sifra
    red(1, 10) = "test"               ' Opis / Svrha
    red(1, 11) = ""                   ' PozivNaBroj
    red(1, 12) = P & "REF-T"          ' BankaReferenz
    red(1, 13) = "test.pdf"           ' IzvorFajl
    red(1, 14) = 0
    red(1, 15) = 0
    red(1, 16) = 0
    red(1, 17) = 0

    saved = SaveBankaImportRows(red)
    ChkEq saved, 1, S & "red je staged"

    v = LookupValue(TBL_BANKA_IMPORT, COL_BIM_BROJ_DOKUMENTA, P & "IZV-11", COL_BIM_DATUM_TRANSAKCIJE)

    ' Ako ovo padne uz VarType=8 (String), kolona DatumTransakcije je formatirana
    ' kao Tekst pa Excel i typed Date upisuje kao string -- promeni format kolone.
    Chk VarType(v) = vbDate, S & "DatumTransakcije je Date, ne String [VarType=" & CStr(VarType(v)) & "]"

    If VarType(v) = vbDate Then
        ChkEq Format$(CDate(v), "yyyy-mm-dd"), "2026-02-01", S & "datum je 1. februar (bez pomeranja)"
    End If

    ' Dedupe mora da radi i kad je jedna strana String (zatecen legacy staging).
    Chk IsDuplicateBankaImport(P & "IZV-11", DateSerial(2026, 2, 1), 1234.56, 0, _
                               P & "PARTNER-T", "", P & "RAC-9"), _
        S & "dedupe prepoznaje duplikat po typed datumu"
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

    ' 1) Resolver kandidata dize jasnu, prepoznatljivu gresku (automatski put).
    On Error Resume Next
    dummy = GetOtkupCandidatesForKooperantBlock(P & "K-1", P & "BLOK-3K")
    errNum = Err.Number
    Err.Clear
    On Error GoTo 0

    ChkEq errNum, ERR_BMAP_MANUAL_REQUIRED, S & "3 kandidata -> ERR_BMAP_MANUAL_REQUIRED"

    ' ...ali uz izricitu potvrdu (rucni put) vraca punu listu, sortiranu opadajuce.
    Dim sviKandidati As Variant
    On Error Resume Next
    sviKandidati = GetOtkupCandidatesForKooperantBlock(P & "K-1", P & "BLOK-3K", True)
    errNum = Err.Number
    Err.Clear
    On Error GoTo 0

    ChkEq errNum, 0, S & "allowOverMax:=True ne dize gresku"
    Chk Not IsEmpty(sviKandidati), S & "allowOverMax:=True vraca kandidate"

    If Not IsEmpty(sviKandidati) Then
        ChkEq UBound(sviKandidati, 1), 3, S & "vracena sva tri kandidata"
        ChkEqD CDbl(sviKandidati(1, 2)), 1400, S & "sortirano opadajuce (najveci otvoreni prvi)"

        ' Planer koji koristi i preview i pisac: 3000 = 1400 + 1200 + 400.
        Dim plan As Variant
        plan = PlanBlokRaspodela(sviKandidati, 3000)

        ChkEq UBound(plan, 1), 3, S & "plan raspodele ima tri reda"
        ChkEqD CDbl(plan(3, 2)), 400, S & "poslednji red dobija samo ostatak (bez preplate)"
    End If

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

    ' 4) Red oznacen "za rucno" mora stvarno da se ZAVRSI rucnom putanjom (istom
    ' koju zove dugme "Rucno mapiraj red" posle potvrde podele). Bez ovoga bi red
    ' bio trajno nezavrsiv: rucni wrapper zove isti core i isti resolver.
    Dim n As Long
    n = MapBankaImportAsKooperantBlockManual_TX(P & "BIM-3K", P & "K-1", P & "BLOK-3K", False, True)

    Chk n >= 3, S & "rucna putanja (potvrdjena podela) knjizi sve stavke [n=" & CStr(n) & "]"
    ChkEq BimObradjeno(P & "BIM-3K"), "Da", S & "posle rucnog mapiranja red je zatvoren"
    ChkEq NovacZaBim(P & "BIM-3K"), n, S & "broj redova u tblNovac odgovara vracenom broju"
    ChkEqD IsplataZaBim(P & "BIM-3K"), 3000, S & "ukupno knjizeno = iznos stavke izvoda (bez preplate)"
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
' T06 - uplata veca od otvorenog iznosa fakture.
'
' Ranije je CEO iznos izvoda isao na fakturu kao NOV_KUPCI_UPLATA -> faktura
' preplacena, a stvaran avans neevidentiran. Sada: na fakturu ide najvise otvoren
' iznos, visak je NOV_KUPCI_AVANS, oba reda u istoj transakciji.
' ============================================================
Private Sub T06_UplataPrekoOtvorenogSeDeli()
    Const S As String = "T06 preplata fakture: "

    Dim res As String
    Dim errNum As Long

    SeedKupac P & "KUP-2", P & "Kupac 2"
    SeedFaktura P & "FAK-1", P & "KUP-2", P & "F-001", 1000

    ' (a) Uplata 1500 na fakturu sa 1000 otvorenog -> 1000 + 500 avans.
    SeedBim P & "BIM-VISAK", P & "IZV-12", P & "RAC-1", P & "PARTNER-F", 1500, 0, "", "", ""

    res = MapBankaImportAsKupac(P & "BIM-VISAK", P & "KUP-2", P & "FAK-1", False)

    Chk res <> "", S & "mapiranje vratilo NovacID"
    ChkEq NovacZaBim(P & "BIM-VISAK"), 2, S & "nastala DVA reda (faktura + avans)"
    ChkEqD UplataZaFakturu(P & "FAK-1"), 1000, S & "na fakturu tacno otvoreni iznos (bez preplate)"
    ChkEqD UplataZaBimPoTipu(P & "BIM-VISAK", NOV_KUPCI_AVANS), 500, S & "visak knjizen kao avans"
    ChkEqD UplataZaBim(P & "BIM-VISAK"), 1500, S & "zbir knjizenog = iznos iz izvoda"
    ChkEq BimObradjeno(P & "BIM-VISAK"), "Da", S & "stavka zatvorena tek kad su oba reda knjizena"

    ' (b) Delimicna uplata (manja od otvorenog) ide cela na fakturu, bez avansa.
    SeedFaktura P & "FAK-2", P & "KUP-2", P & "F-002", 2000
    SeedBim P & "BIM-DEO", P & "IZV-12", P & "RAC-1", P & "PARTNER-F", 800, 0, "", "", ""

    res = MapBankaImportAsKupac(P & "BIM-DEO", P & "KUP-2", P & "FAK-2", False)

    ChkEq NovacZaBim(P & "BIM-DEO"), 1, S & "delimicna uplata = jedan red"
    ChkEqD UplataZaFakturu(P & "FAK-2"), 800, S & "cela delimicna uplata na fakturu"
    ChkEqD UplataZaBimPoTipu(P & "BIM-DEO", NOV_KUPCI_AVANS), 0, S & "nema avans reda"

    ' (c) Vec placena faktura se odbija (saldo se promenio izmedju prikaza i klika).
    SeedBim P & "BIM-PLAC", P & "IZV-12", P & "RAC-1", P & "PARTNER-F", 300, 0, "", "", ""

    On Error Resume Next
    res = MapBankaImportAsKupac(P & "BIM-PLAC", P & "KUP-2", P & "FAK-1", False)
    errNum = Err.Number
    Err.Clear
    On Error GoTo 0

    Chk errNum <> 0, S & "uplata na vec placenu fakturu je odbijena"
    ChkEq NovacZaBim(P & "BIM-PLAC"), 0, S & "odbijena uplata nije knjizena"
    ChkEq BimObradjeno(P & "BIM-PLAC"), "", S & "odbijena uplata ostaje otvorena"
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

Private Sub SeedFaktura(ByVal fakturaID As String, ByVal kupacID As String, _
                        ByVal brojFakture As String, ByVal iznos As Double)
    BitAppend TBL_FAKTURE, _
        Array(COL_FAK_ID, COL_FAK_KUPAC, COL_FAK_BROJ, COL_FAK_IZNOS, COL_FAK_DATUM), _
        Array(fakturaID, kupacID, brojFakture, iznos, Date)
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

' Zbir isplata u tblNovac koje nose BIM marker ovog staging reda.
Private Function IsplataZaBim(ByVal bimID As String) As Double
    Dim data As Variant
    Dim colNap As Long
    Dim colIsplata As Long
    Dim i As Long

    data = GetTableData(TBL_NOVAC)
    If IsEmpty(data) Then Exit Function

    data = ExcludeStornirano(data, TBL_NOVAC)
    If IsEmpty(data) Then Exit Function

    colNap = GetColumnIndex(TBL_NOVAC, COL_NOV_NAPOMENA)
    colIsplata = GetColumnIndex(TBL_NOVAC, COL_NOV_ISPLATA)

    For i = 1 To UBound(data, 1)
        If BimIdFromNapomena(CStr(NzTb(data(i, colNap)))) = bimID Then
            IsplataZaBim = IsplataZaBim + CDbl(nz(data(i, colIsplata), "0"))
        End If
    Next i
End Function

' Zbir uplata u tblNovac koje nose BIM marker ovog staging reda.
Private Function UplataZaBim(ByVal bimID As String) As Double
    UplataZaBim = SumNovacZaBim(bimID, "")
End Function

' Zbir uplata za BIM marker, filtriran po tipu knjizenja.
Private Function UplataZaBimPoTipu(ByVal bimID As String, ByVal tipNovca As String) As Double
    UplataZaBimPoTipu = SumNovacZaBim(bimID, tipNovca)
End Function

Private Function SumNovacZaBim(ByVal bimID As String, ByVal tipNovca As String) As Double
    Dim data As Variant
    Dim colNap As Long, colUplata As Long, colTip As Long
    Dim i As Long

    data = GetTableData(TBL_NOVAC)
    If IsEmpty(data) Then Exit Function

    data = ExcludeStornirano(data, TBL_NOVAC)
    If IsEmpty(data) Then Exit Function

    colNap = GetColumnIndex(TBL_NOVAC, COL_NOV_NAPOMENA)
    colUplata = GetColumnIndex(TBL_NOVAC, COL_NOV_UPLATA)
    colTip = GetColumnIndex(TBL_NOVAC, COL_NOV_TIP)

    For i = 1 To UBound(data, 1)
        If BimIdFromNapomena(CStr(NzTb(data(i, colNap)))) = bimID Then
            If tipNovca = "" Or Trim$(CStr(NzTb(data(i, colTip)))) = tipNovca Then
                SumNovacZaBim = SumNovacZaBim + CDbl(nz(data(i, colUplata), "0"))
            End If
        End If
    Next i
End Function

' Zbir aktivnih uplata vezanih za fakturu (isti obracun koji koristi writer).
Private Function UplataZaFakturu(ByVal fakturaID As String) As Double
    UplataZaFakturu = GetUplataForFaktura(fakturaID)
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

Private Sub ChkEqD(ByVal act As Double, ByVal exp As Double, ByVal nm As String)
    If Abs(act - exp) <= 0.001 Then
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
