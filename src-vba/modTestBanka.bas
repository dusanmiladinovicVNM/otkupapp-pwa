Attribute VB_Name = "modTestBanka"
Option Explicit

' ============================================================
' modTestBanka - tvrde kapije za banku: RF-09 (import + mapiranje)
'                i RF-10 (export naloga / AUD-026)
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
' Pokriva nalaze koje RF-09 zatvara:
'  T01 AUD-007  nemoguc datum se ODBIJA (ne pomera); dd.mm.yyyy je locale-nezavisan
'  T02 AUD-025  dedupe kljuc ukljucuje broj racuna (multi-account transakcija)
'  T03 AUD-025  3+ kandidata bloka obara SAMO taj red, ne ceo AutoMapAll batch;
'               red oznacen "za rucno" se zavrsava javnom rucnom putanjom
'  T04 AUD-025  rucno mapiranje pogresnog smera je odbijeno
'  T05 AUD-007  staging cuva TYPED datum, a javni writer odbija nevalidan
'  T06 FM-0023  uplata preko otvorenog iznosa fakture se deli (faktura + avans)
'  T07 FM-0024  istoimeni partneri se razlikuju po ID-u (prikaz i knjizenje)
'  T08 AUD-014  batch koji padne propagira gresku (ne prijavljuje "0 mapirano")
'  T09 FM-0024  AUTO mapiranje ide po pozivu na broj (ne po rucnom izboru bloka)
'  T10 AUD-025  vec placena faktura ne obara ceo "Auto sve" batch
'  T11 AUD-025  rucni kooperant sa PRAZNIM izborom bloka ide kroz potvrdjenu podelu
'  T12 AUD-026  zaostao "Isplatiti" override se clamp-uje na trenutno otvoreno
'  T13 AUD-026  CSV nalozi se ODBIJAJU kada nalog trazi vise nego sto je otvoreno
'               (granice u cent-domenu: 600.01 na otvoreno 600.00 je preplata)
'  T14 AUD-026  kapija je UNUTAR putanje koja gradi CSV payload (ne samo validator)
'  T15 AUD-026  dupli/prazan OtkupID obara saldo-mapu (fail-closed kanonski kljuc)
'  T16 AUD-026  "Primeni avans" odbija dupli OtkupID (core guard, tabelarni test)
'  T17 AUD-026  SAKRIVEN duplikat (samo jedan red otvoren) ne isplacuje pogresnog
'               kooperanta -- vlasnik se ne resava "prvim pogotkom" nad sirovom tabelom
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

    If MsgBox("Pokrenuti banka test suite (RF-09 uvoz/mapiranje + RF-10 nalozi)?" & vbCrLf & vbCrLf & _
              "Svi podaci (BIT-*) se prave u transakciji i UVEK se ponistavaju " & _
              "(rollback). Nista ne ostaje u tabelama.", _
              vbQuestion + vbYesNo, APP_NAME) <> vbYes Then Exit Sub

    mPass = 0: mFail = 0: mFails = "": mReport = ""

    ' T01 je cist parser (bez tabela) -> moze i van transakcije.
    T01_NemoguciDatumOdbijen

    ' RF-10 seam-ovi rade nad kolekcijom/Dictionary-jem u memoriji (bez tabela).
    T12_ClampOverrideNaOtvoreno
    T13_NalogPrekoOtvorenogSeOdbija
    T14_CsvPayloadNosiKapiju
    T15_DupliOtkupIDObaraSaldoMapu

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
    T07_IstoimeniPartneriSeRazlikujuPoID
    T08_PaliBatchNePrijavljujeUspeh
    T09_AutoBlokIdePoPozivuNaBroj
    T10_PlacenaFakturaNeObaraBatch
    T11_RucniKooperantBezIzboraBloka
    ' T17 ide PRE T16: T16 namerno ostavlja dva OTVORENA reda sa istim OtkupID,
    ' sto bi oborilo kontrolni slucaj u T17 (rollback je tek na kraju suite-a).
    T17_SakrivenDuplikatNeIsplacujePogresnog
    T16_DupliOtkupIDNeDozvoljavaAvans

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

    ' PAZNJA (nalaz iz prvog pravog pokretanja suite-a): ove provere ne cuvaju samo
    ' DMY parser nego i to da `CDate` fallback NE "spasava" d.m.y vrednost koju je
    ' DMY parser odbio. Na en-US masini `IsDate("01.13.2026")` je True jer VBA
    ' ZAMENI dan i mesec (13. januar) -- pa je bas ta provera pala dok fallback nije
    ' zatvoren za d.m.y oblik.
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

    Dim saved As Long
    Dim v As Variant
    Dim errNum As Long

    saved = SaveBankaImportRows_TX(BuildStagingRow(P & "IZV-11", P & "RAC-9", _
                DateSerial(2026, 2, 1), DateSerial(2026, 2, 1), P & "REF-T"))
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

    Chk IsDuplicateBankaImport(P & "IZV-11", "01.02.2026", 1234.56, 0, _
                               P & "PARTNER-T", "", P & "RAC-9"), _
        S & "dedupe poredi typed datum i string zapis istog dana"

    ' Javni writer je fail-closed za datum: tekst koji nije datum se ODBIJA
    ' (ranije se upisivao kakav jeste), kao i datum van poslovnog opsega.
    On Error Resume Next
    saved = SaveBankaImportRows_TX(BuildStagingRow(P & "IZV-14", P & "RAC-9", _
                DateSerial(2026, 2, 1), "nije datum", P & "REF-X1"))
    errNum = Err.Number
    Err.Clear
    On Error GoTo 0

    Chk errNum <> 0, S & "writer odbija tekst koji nije datum"

    On Error Resume Next
    saved = SaveBankaImportRows_TX(BuildStagingRow(P & "IZV-15", P & "RAC-9", _
                DateSerial(2026, 2, 1), DateSerial(1899, 12, 30), P & "REF-X2"))
    errNum = Err.Number
    Err.Clear
    On Error GoTo 0

    Chk errNum <> 0, S & "writer odbija datum van poslovnog opsega (1899)"

    ChkEq StagingRedova(P & "IZV-14") + StagingRedova(P & "IZV-15"), 0, _
        S & "odbijeni redovi nisu upisani"

    ' Batch je atomican: prvi red validan, drugi nevalidan -> NIJEDAN ne ostaje.
    ' Bez sopstvene TX u javnom ulazu prvi red bi ostao upisan (polovicno uvezen
    ' izvod koji vise nije ni duplikat ni ceo).
    Dim mesovit As Variant
    mesovit = BuildStagingRowPair(P & "IZV-16", P & "RAC-9", _
                DateSerial(2026, 2, 1), "nije datum")

    On Error Resume Next
    saved = SaveBankaImportRows_TX(mesovit)
    errNum = Err.Number
    Err.Clear
    On Error GoTo 0

    Chk errNum <> 0, S & "mesovit batch (validan + nevalidan) je odbijen"
    ChkEq StagingRedova(P & "IZV-16"), 0, _
        S & "ni PRVI (validan) red nije ostao -- batch je atomican"
End Sub

' ============================================================
' T07 - istoimeni partneri (FM-0024 #7).
'
' Rezolucija po nazivu vraca PRVI pogodak, a lista je istoimene partnere jos i
' spajala u jednu stavku -- operater nije mogao da ih razlikuje, a uplata je mogla
' da ode na pogresnog kupca. Identitet ide preko ID-a, a prikaz ga pokazuje.
' ============================================================
Private Sub T07_IstoimeniPartneriSeRazlikujuPoID()
    Const S As String = "T07 istoimeni partneri: "

    Dim novID As String

    SeedKupac P & "KUP-A", P & "Agro Trade"
    SeedKupac P & "KUP-B", P & "Agro Trade"
    SeedStanica P & "OM-A", P & "Stanica Ista"
    SeedStanica P & "OM-B", P & "Stanica Ista"

    ' Zasto identitet NE sme da ide po nazivu: lookup vraca prvi pogodak.
    ChkEq CStr(LookupValue(TBL_KUPCI, "Naziv", P & "Agro Trade", "KupacID")), P & "KUP-A", _
        S & "lookup po nazivu vraca PRVOG (zato se ne koristi za identitet)"
    ChkEq CStr(LookupValue(TBL_STANICE, "Naziv", P & "Stanica Ista", "StanicaID")), P & "OM-A", _
        S & "isto vazi i za stanice"

    ' Prikaz u listi mora da razlikuje duplikate (ID uz naziv).
    ChkEq ComboDisplayWithID(P & "Agro Trade", P & "KUP-B"), _
          P & "Agro Trade" & "  [" & P & "KUP-B" & "]", S & "prikaz nosi ID"
    ChkEq ComboDisplayWithID(P & "Agro Trade", ""), P & "Agro Trade", _
        S & "bez ID-a prikaz ostaje nepromenjen"

    ' Knjizenje ide na IZABRAN ID, ne na prvi po nazivu.
    SeedBim P & "BIM-DUP", P & "IZV-13", P & "RAC-1", P & "Agro Trade", 700, 0, "", "", ""

    gBankaSilentBatch = True
    novID = MapBankaImportAsKupac_TX(P & "BIM-DUP", P & "KUP-B", "", False)
    gBankaSilentBatch = False

    Chk novID <> "", S & "mapiranje na drugi ID je proslo"
    ChkEq NovacPartnerID(novID), P & "KUP-B", _
        S & "novac pripada IZABRANOM kupcu, ne prvom po nazivu"
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
    OstaviOtvorenimSamo Array(P & "BIM-3K", P & "BIM-OK")

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
' Ide kroz javni `_TX` API (ne-TX mutatori su Private): greska se loguje i vraca
' pozivaocu, a `gBankaSilentBatch` sprecava blokirajuci dijalog u test rezimu.
' ============================================================
Private Sub T04_SmerGuardOdbijaPogresanTip()
    Const S As String = "T04 smer guard: "

    Dim res As String
    Dim n As Long

    SeedStanica P & "OM-2", P & "Stanica 2"
    SeedKooperant P & "K-2", "Test", "Smer", P & "OM-2"
    SeedKupac P & "KUP-1", P & "Kupac 1"

    SeedBim P & "BIM-UPL", P & "IZV-10", P & "RAC-1", P & "PARTNER-S", 4000, 0, "", "", ""
    SeedBim P & "BIM-ISP", P & "IZV-10", P & "RAC-1", P & "PARTNER-S", 0, 4000, "", "", ""

    ' Javni API su samo `_TX` ulazi (ne-TX mutatori su Private), pa i test ide
    ' kroz njih; gBankaSilentBatch drzi greske bez blokirajuceg dijaloga.
    gBankaSilentBatch = True

    ' Uplata knjizena kao kooperant (isplata) -> odbijeno.
    n = MapBankaImportAsKooperantBlock_TX(P & "BIM-UPL", P & "K-2", False)

    ChkEq n, 0, S & "uplata + tip Kooperant -> odbijeno (0 redova)"
    ChkEq BimObradjeno(P & "BIM-UPL"), "", S & "odbijena uplata ostaje otvorena"
    ChkEq NovacZaBim(P & "BIM-UPL"), 0, S & "odbijena uplata nije knjizena"

    ' Isplata knjizena kao kupac (uplata) -> odbijeno.
    res = MapBankaImportAsKupac_TX(P & "BIM-ISP", P & "KUP-1", "", False)

    ChkEq res, "", S & "isplata + tip Kupac -> odbijeno (nema NovacID)"
    ChkEq BimObradjeno(P & "BIM-ISP"), "", S & "odbijena isplata ostaje otvorena"
    ChkEq NovacZaBim(P & "BIM-ISP"), 0, S & "odbijena isplata nije knjizena"

    ' Razlog odbijanja mora biti smer -- proverava se na ne-TX putanji kroz
    ' javni klasifikator (ClassifyBimSmer), koji writer koristi za istu odluku.
    ChkEq ClassifyBimSmer(4000, 0), BIM_SMER_UPLATA, S & "klasifikator: 4000/0 = UPLATA"
    ChkEq ClassifyBimSmer(0, 4000), BIM_SMER_ISPLATA, S & "klasifikator: 0/4000 = ISPLATA"
    ChkEq ClassifyBimSmer(100, 100), BIM_SMER_NEJASAN, S & "klasifikator: oba iznosa = NEJASAN"
    ChkEq ClassifyBimSmer(0, 0), BIM_SMER_NEJASAN, S & "klasifikator: bez iznosa = NEJASAN"

    gBankaSilentBatch = False
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

    SeedKupac P & "KUP-2", P & "Kupac 2"
    SeedFaktura P & "FAK-1", P & "KUP-2", P & "F-001", 1000

    ' (a) Uplata 1500 na fakturu sa 1000 otvorenog -> 1000 + 500 avans.
    SeedBim P & "BIM-VISAK", P & "IZV-12", P & "RAC-1", P & "PARTNER-F", 1500, 0, "", "", ""

    res = MapBankaImportAsKupac_TX(P & "BIM-VISAK", P & "KUP-2", P & "FAK-1", False)

    Chk res <> "", S & "mapiranje vratilo NovacID"
    ChkEq NovacZaBim(P & "BIM-VISAK"), 2, S & "nastala DVA reda (faktura + avans)"
    ChkEqD UplataZaFakturu(P & "FAK-1"), 1000, S & "na fakturu tacno otvoreni iznos (bez preplate)"
    ChkEqD UplataZaBimPoTipu(P & "BIM-VISAK", NOV_KUPCI_AVANS), 500, S & "visak knjizen kao avans"
    ChkEqD UplataZaBim(P & "BIM-VISAK"), 1500, S & "zbir knjizenog = iznos iz izvoda"
    ChkEq BimObradjeno(P & "BIM-VISAK"), "Da", S & "stavka zatvorena tek kad su oba reda knjizena"

    ' (b) Delimicna uplata (manja od otvorenog) ide cela na fakturu, bez avansa.
    SeedFaktura P & "FAK-2", P & "KUP-2", P & "F-002", 2000
    SeedBim P & "BIM-DEO", P & "IZV-12", P & "RAC-1", P & "PARTNER-F", 800, 0, "", "", ""

    res = MapBankaImportAsKupac_TX(P & "BIM-DEO", P & "KUP-2", P & "FAK-2", False)

    ChkEq NovacZaBim(P & "BIM-DEO"), 1, S & "delimicna uplata = jedan red"
    ChkEqD UplataZaFakturu(P & "FAK-2"), 800, S & "cela delimicna uplata na fakturu"
    ChkEqD UplataZaBimPoTipu(P & "BIM-DEO", NOV_KUPCI_AVANS), 0, S & "nema avans reda"

    ' (c) Vec placena faktura se odbija (saldo se promenio izmedju prikaza i klika).
    '
    ' `_TX` omotac gresku HVATA, rollback-uje i vraca "" -- ne propagira je dalje,
    ' pa `Err.Number` posle poziva nije deo ugovora. Zato se proverava ugovor koji
    ' omotac stvarno daje: prazan rezultat + nista knjizeno + stavka otvorena.
    SeedBim P & "BIM-PLAC", P & "IZV-12", P & "RAC-1", P & "PARTNER-F", 300, 0, "", "", ""

    gBankaSilentBatch = True
    res = MapBankaImportAsKupac_TX(P & "BIM-PLAC", P & "KUP-2", P & "FAK-1", False)
    gBankaSilentBatch = False

    ChkEq res, "", S & "uplata na vec placenu fakturu je odbijena (nema NovacID)"
    ChkEq NovacZaBim(P & "BIM-PLAC"), 0, S & "odbijena uplata nije knjizena"
    ChkEq BimObradjeno(P & "BIM-PLAC"), "", S & "odbijena uplata ostaje otvorena"
    ChkEqD UplataZaFakturu(P & "FAK-1"), 1000, S & "saldo fakture nepromenjen posle odbijanja"

    ' Razlog odbijanja se proverava na cistom validatoru (writer koristi isti):
    ' otvoreno = iznos - aktivne uplate, pa je vec placena faktura 0.
    ChkEqD GetOtvorenoNaFakturi(P & "FAK-1"), 0, S & "GetOtvorenoNaFakturi = 0 za placenu fakturu"
    ChkEqD GetOtvorenoNaFakturi(P & "FAK-2"), 1200, S & "GetOtvorenoNaFakturi = 2000 - 800"
End Sub

' ============================================================
' T08 - batch koji padne NE sme da izgleda kao "0 mapirano".
'
' Ranije je AutoMapAllBankaImport_TX posle rollback-a vracao 0 bez greske, pa je
' forma odmah prikazivala "Automatski mapirano: 0" -- uredno zavrsen batch bez
' pogodaka izgleda isto kao batch koji je ponisten. Greska se sada propagira.
'
' Pad se izaziva PODACIMA (bez test-hook-a u produkciji): stavka izvoda bez broja
' dokumenta obara RequireIzvodBroj, sto NIJE ERR_BMAP_MANUAL_REQUIRED, pa batch
' omotac tu gresku ne guta.
' ============================================================
Private Sub T08_PaliBatchNePrijavljujeUspeh()
    Const S As String = "T08 pali batch: "

    Dim mapped As Long
    Dim manualRequired As Long
    Dim errNum As Long

    SeedStanica P & "OM-3", P & "Stanica 3"
    SeedKooperant P & "K-3", "Test", "Batch", P & "OM-3"

    SeedOtkup P & "OTK-G1", P & "K-3", P & "BLOK-G1", 100, 10, "Malina"   ' 1000
    SeedOtkup P & "OTK-G2", P & "K-3", P & "BLOK-G2", 100, 20, "Malina"   ' 2000

    ' Prvi red je ispravan i bio bi mapiran; drugi nema broj izvoda -> pad.
    SeedBim P & "BIM-G1", P & "IZV-17", P & "RAC-1", P & "PARTNER-B", 0, 1000, P & "BLOK-G1", "", ""
    SeedBim P & "BIM-G2", "", P & "RAC-1", P & "PARTNER-B", 0, 2000, P & "BLOK-G2", "", ""

    OstaviOtvorenimSamo Array(P & "BIM-G1", P & "BIM-G2")

    gBankaSilentBatch = True
    On Error Resume Next
    mapped = AutoMapAllBankaImport_TX(manualRequired)
    errNum = Err.Number
    Err.Clear
    On Error GoTo 0
    gBankaSilentBatch = False

    Chk errNum <> 0, S & "pali batch propagira gresku (ne vraca tiho 0)"
    ChkEq mapped, 0, S & "rezultat je 0 (nista nije ostalo mapirano)"

    ' Rollback mora da vrati i red koji je USPEO pre pada.
    ChkEq BimObradjeno(P & "BIM-G1"), "", S & "uspeli red je vracen u otvoreno stanje"
    ChkEq NovacZaBim(P & "BIM-G1"), 0, S & "uspeli red nema zaostao novac posle rollback-a"
    ChkEq NovacZaBim(P & "BIM-G2"), 0, S & "pali red nije ostavio parcijalno knjizenje"
End Sub

' ============================================================
' T09 - AUTO mapiranje ide ISKLJUCIVO po pozivu na broj iz izvoda.
'
' Preview u formi je jedno vreme za AUTO granu racunao blok preko istog helpera
' koji koristi RUCNI put (koji gleda cmbOtkupBlok), pa je operater mogao da vidi
' blok B a da "Automatski mapiraj red" proknjizi blok A. Writer i preview sada
' zovu ISTU funkciju (AutoBlockNoForBim); ovaj test cuva njen ugovor i dokazuje
' da auto knjizenje pogadja blok iz poziva na broj, a ne neki drugi blok istog
' kooperanta.
' ============================================================
Private Sub T09_AutoBlokIdePoPozivuNaBroj()
    Const S As String = "T09 auto blok = poziv na broj: "

    Dim res As String

    SeedStanica P & "OM-4", P & "Stanica 4"
    SeedKooperant P & "K-4", "Test", "Auto", P & "OM-4"

    ' Dva bloka ISTOG kooperanta; izvod pokazuje na prvi.
    SeedOtkup P & "OTK-PA", P & "K-4", P & "BLOK-PA", 100, 10, "Malina"   ' 1000
    SeedOtkup P & "OTK-PB", P & "K-4", P & "BLOK-PB", 100, 10, "Malina"   ' 1000

    SeedBim P & "BIM-AUTO", P & "IZV-18", P & "RAC-1", P & "PARTNER-A", 0, 1000, P & "BLOK-PA", "", ""

    ChkEq AutoBlockNoForBim(P & "BIM-AUTO"), P & "BLOK-PA", _
        S & "AutoBlockNoForBim vraca poziv na broj iz izvoda"

    gBankaSilentBatch = True
    res = AutoMapBankaImportRow_TX(P & "BIM-AUTO")
    gBankaSilentBatch = False

    Chk res <> "", S & "auto mapiranje je proslo"
    ChkEq BimObradjeno(P & "BIM-AUTO"), "Da", S & "stavka je zatvorena"

    ' Knjizenje mora da pogodi blok iz poziva na broj, ne drugi blok kooperanta.
    ChkEqD GetUplataForOtkup(P & "OTK-PA"), 1000, S & "placen otkup iz poziva na broj (BLOK-PA)"
    ChkEqD GetUplataForOtkup(P & "OTK-PB"), 0, S & "drugi blok kooperanta NIJE dirnut (BLOK-PB)"
End Sub

' ============================================================
' T10 - vec placena faktura ne sme da obori CEO "Auto sve" batch.
'
' AUD-025 cilj je "jedan problematican red ne obara batch". Prva verzija je u
' batch-u gutala samo ERR_BMAP_MANUAL_REQUIRED, pa je red koji se AUTO razresi na
' vec placenu fakturu dizao ERR_BMAP_BASE + 53 i rollback-ovao sve -- ista rupa,
' samo kroz drugu gresku. Obe znace "operater odlucuje" i obe se dizu PRE upisa.
' ============================================================
Private Sub T10_PlacenaFakturaNeObaraBatch()
    Const S As String = "T10 placena faktura u batch-u: "

    Dim mapped As Long
    Dim manualRequired As Long
    Dim errNum As Long

    SeedKupac P & "KUP-3", P & "Kupac 3"
    SeedFaktura P & "FAK-3", P & "KUP-3", P & "F-003", 500

    ' Zdrav red: kooperant sa jednim otvorenim blokom (mapira se).
    SeedStanica P & "OM-5", P & "Stanica 5"
    SeedKooperant P & "K-5", "Test", "Batch2", P & "OM-5"
    SeedOtkup P & "OTK-H1", P & "K-5", P & "BLOK-H1", 100, 10, "Malina"
    SeedBim P & "BIM-H1", P & "IZV-19", P & "RAC-1", P & "PARTNER-H", 0, 1000, P & "BLOK-H1", "", ""

    ' Problematican red: poziv na broj pogadja fakturu koja je VEC placena.
    SeedBim P & "BIM-PLAC2", P & "IZV-19", P & "RAC-1", P & "PARTNER-H2", 300, 0, P & "F-003", "", ""
    SeedNovacUplataFakture P & "NOV-PLAC", P & "KUP-3", P & "FAK-3", 500   ' faktura zatvorena

    ChkEqD GetOtvorenoNaFakturi(P & "FAK-3"), 0, S & "faktura je stvarno placena (otvoreno = 0)"

    OstaviOtvorenimSamo Array(P & "BIM-H1", P & "BIM-PLAC2")

    gBankaSilentBatch = True
    On Error Resume Next
    mapped = AutoMapAllBankaImport_TX(manualRequired)
    errNum = Err.Number
    Err.Clear
    On Error GoTo 0
    gBankaSilentBatch = False

    ChkEq errNum, 0, S & "batch NIJE pao zbog placene fakture"
    ChkEq BimObradjeno(P & "BIM-H1"), "Da", S & "zdrav red je mapiran (batch nije rollback-ovan)"
    ChkEq BimObradjeno(P & "BIM-PLAC2"), "Error", S & "red sa placenom fakturom je oznacen za rucno"
    Chk manualRequired >= 1, S & "batch prijavio 'za rucno' [" & CStr(manualRequired) & "]"
    ChkEq NovacZaBim(P & "BIM-PLAC2"), 0, S & "problematican red nije parcijalno knjizen"
    ChkEqD UplataZaFakturu(P & "FAK-3"), 500, S & "placena faktura nije preplacena"
End Sub

' ============================================================
' T11 - rucni kooperant sa PRAZNIM izborom bloka.
'
' Lista blokova se puni ali se NE auto-selektuje, pa je prazan combo default
' slucaj. Ranije je ta grana isla na auto-put (bez potvrde podele), pa je blok sa
' 3+ otvorenih stavki tu zavrsavao generickom greskom -- "3+ blok ima izlaz" je
' radilo samo ako operater rucno klikne blok. Sada obe grane koriste ISTI
' efektivni blok (isti koji preview prikazuje) i isti potvrdjeni put.
' ============================================================
Private Sub T11_RucniKooperantBezIzboraBloka()
    Const S As String = "T11 rucno bez izbora bloka: "

    Dim n As Long

    SeedStanica P & "OM-6", P & "Stanica 6"
    SeedKooperant P & "K-6", "Test", "Prazno", P & "OM-6"

    ' Blok sa TRI otvorene stavke; izvod pokazuje na njega preko poziva na broj.
    SeedOtkup P & "OTK-Q1", P & "K-6", P & "BLOK-Q", 100, 10, "Malina"
    SeedOtkup P & "OTK-Q2", P & "K-6", P & "BLOK-Q", 100, 12, "Kupina"
    SeedOtkup P & "OTK-Q3", P & "K-6", P & "BLOK-Q", 100, 14, "Visnja"

    SeedBim P & "BIM-Q", P & "IZV-20", P & "RAC-1", P & "PARTNER-Q", 0, 3000, P & "BLOK-Q", "", ""

    ' Efektivni blok kad combo NIJE popunjen = poziv na broj iz izvoda; forma taj
    ' isti blok salje u potvrdjeni rucni put.
    ChkEq AutoBlockNoForBim(P & "BIM-Q"), P & "BLOK-Q", S & "efektivni blok = poziv na broj"

    gBankaSilentBatch = True
    n = MapBankaImportAsKooperantBlockManual_TX(P & "BIM-Q", P & "K-6", _
                                                AutoBlockNoForBim(P & "BIM-Q"), False, True)
    gBankaSilentBatch = False

    Chk n >= 3, S & "potvrdjena podela knjizi sve stavke [n=" & CStr(n) & "]"
    ChkEq BimObradjeno(P & "BIM-Q"), "Da", S & "stavka je zatvorena"
    ChkEqD IsplataZaBim(P & "BIM-Q"), 3000, S & "ukupno knjizeno = iznos iz izvoda"
End Sub

' ============================================================
' T12 - AUD-026 / FM-0020 #1: clamp zaostalog "Isplatiti" override-a.
'
' Override po bloku prezivljava reload liste. Ako se posle unosa deo bloka
' isplati (ili se blok zatvori/stornira), zaostao override moze da naruci
' vise nego sto je otvoreno. ClampOverridesToOpen ga mora spustiti/obrisati.
' Cist seam: Dictionary + kolekcija blokova, bez tabela.
' ============================================================
Private Sub T12_ClampOverrideNaOtvoreno()
    Const S As String = "T12 clamp override: "

    Dim blokovi As Collection
    Set blokovi = New Collection
    blokovi.Add MakeBlok(P & "OTK-C1", P & "BLOK-C1", 600)    ' otvoreno palo sa 1000 na 600
    blokovi.Add MakeBlok(P & "OTK-C2", P & "BLOK-C2", 900)    ' delimicna isplata, jos vazi
    blokovi.Add MakeBlok(P & "OTK-C4", P & "BLOK-C4", 0)      ' vise nema otvorenog iznosa
    blokovi.Add MakeBlok(P & "OTK-C5", P & "BLOK-C5", 600)    ' granica: override je tacno cent veci
    blokovi.Add MakeBlok(P & "OTK-C6", P & "BLOK-C6", 600)    ' granica: sub-cent ispod polovine

    Dim ovr As Object
    Set ovr = CreateObject("Scripting.Dictionary")
    ovr(P & "OTK-C1") = 1000#      ' > otvoreno -> clamp na 600
    ovr(P & "OTK-C2") = 300#       ' < otvoreno -> namerna delimicna isplata, ne dirati
    ovr(P & "OTK-C3") = 500#       ' bloka vise nema u listi -> brisanje
    ovr(P & "OTK-C4") = 200#       ' blok bez otvorenog iznosa -> brisanje
    ovr(P & "OTK-C5") = 600.01     ' PUN cent preko otvorenog -> mora se spustiti
    ovr(P & "OTK-C6") = 600.004    ' zaokruzuje se na 600.00 -> nije preplata

    Dim changed As Long
    changed = ClampOverridesToOpen(ovr, blokovi)

    ChkEq changed, 4, S & "4 promenjena unosa (2 clamp + 2 brisanja)"
    ChkEqD CDbl(ovr(P & "OTK-C1")), 600, S & "override 1000 spusten na otvoreno 600"
    ChkEqD CDbl(ovr(P & "OTK-C2")), 300, S & "override manji od otvorenog ostaje netaknut"
    Chk Not ovr.Exists(P & "OTK-C3"), S & "override za nestao blok je obrisan"
    Chk Not ovr.Exists(P & "OTK-C4"), S & "override za blok bez otvorenog je obrisan"
    ChkEqD CDbl(ovr(P & "OTK-C5")), 600, S & "override 600.01 na otvoreno 600.00 je spusten (pun cent)"
    ChkEqD CDbl(ovr(P & "OTK-C6")), 600.004, S & "sub-cent ispod polovine se ne dira"
    ChkEq ovr.count, 4, S & "u Dictionary-ju su ostala samo 4 vazeca override-a"

    ' Ponovljen poziv nad vec uskladjenim stanjem ne sme nista da menja.
    ChkEq ClampOverridesToOpen(ovr, blokovi), 0, S & "drugi prolaz je no-op"

    ' Bez liste blokova nema dokaza da je ijedan override vazeci -> sve pada.
    ChkEq ClampOverridesToOpen(ovr, Nothing), 4, S & "bez liste blokova se svi override-i brisu"
    ChkEq ovr.count, 0, S & "Dictionary je prazan posle brisanja"
End Sub

' ============================================================
' T13 - AUD-026 / FM-0020 #2 + FM-0021 #1: finalna saldo kapija CSV naloga.
'
' Izmedju prikaza i klika na "Generisi" saldo se moze promeniti (uvoz izvoda,
' vezan avans, storno). ValidateNalogSaldo mora da odbije ceo CSV kada ijedan
' nalog trazi vise nego sto je TADA otvoreno -- fail-closed, sa imenom bloka.
' ============================================================
Private Sub T13_NalogPrekoOtvorenogSeOdbija()
    Const S As String = "T13 saldo revalidacija: "

    Dim otvoreno As Object
    Set otvoreno = CreateObject("Scripting.Dictionary")
    otvoreno(P & "OTK-V1") = 600#
    otvoreno(P & "OTK-V2") = 900#

    ' 1) Nalozi u granicama otvorenog -> prolaz.
    Dim ok As Collection
    Set ok = New Collection
    ok.Add MakeNalog(P & "OTK-V1", P & "BLOK-V1", 600, True)
    ok.Add MakeNalog(P & "OTK-V2", P & "BLOK-V2", 250, True)
    ChkEq ValidateNalogSaldo(ok, otvoreno), "", S & "nalozi u granicama otvorenog prolaze"

    ' 2) Jedan nalog preko otvorenog -> odbijanje, sa brojem bloka u poruci.
    Dim preplata As Collection
    Set preplata = New Collection
    preplata.Add MakeNalog(P & "OTK-V1", P & "BLOK-V1", 600, True)
    preplata.Add MakeNalog(P & "OTK-V2", P & "BLOK-V2", 1000, True)

    Dim res As String
    res = ValidateNalogSaldo(preplata, otvoreno)
    Chk LenB(res) > 0, S & "nalog veci od otvorenog je odbijen"
    Chk InStr(res, P & "BLOK-V2") > 0, S & "poruka imenuje problematican blok"
    Chk InStr(res, "1.000,00") > 0 Or InStr(res, "1,000.00") > 0, _
        S & "poruka sadrzi trazeni iznos"
    Chk InStr(res, P & "BLOK-V1") = 0, S & "ispravan blok se ne prijavljuje"

    ' 3) Blok kog vise NEMA medju otvorenima = otvoreno 0 -> svaki iznos je preplata.
    Dim nestao As Collection
    Set nestao = New Collection
    nestao.Add MakeNalog(P & "OTK-V9", P & "BLOK-V9", 10, True)
    Chk LenB(ValidateNalogSaldo(nestao, otvoreno)) > 0, _
        S & "blok koji vise nije otvoren je odbijen (fail-closed)"

    ' 4) Blok BEZ tekuceg racuna ne ide u CSV, pa ne sme ni da obori generisanje.
    Dim bezTR As Collection
    Set bezTR = New Collection
    bezTR.Add MakeNalog(P & "OTK-V1", P & "BLOK-V1", 5000, False)
    ChkEq ValidateNalogSaldo(bezTR, otvoreno), "", S & "blok bez TR ne obara naloge"

    ' 5) GRANICE u cent-domenu (bez epsilon tolerancije). CSV nosi dve decimale,
    '    pa se poredi tacno ono sto ce u fajlu i biti.
    Chk LenB(ValidateNalogSaldo(JedanNalog(P & "OTK-V1", P & "BLOK-V1", 600#), otvoreno)) = 0, _
        S & "600.00 na otvoreno 600.00 prolazi"
    Chk LenB(ValidateNalogSaldo(JedanNalog(P & "OTK-V1", P & "BLOK-V1", 600.01), otvoreno)) > 0, _
        S & "600.01 na otvoreno 600.00 je ODBIJEN (pun cent preplate)"
    Chk LenB(ValidateNalogSaldo(JedanNalog(P & "OTK-V1", P & "BLOK-V1", 600.005), otvoreno)) > 0, _
        S & "600.005 se zaokruzuje na 600.01 -> odbijen"
    Chk LenB(ValidateNalogSaldo(JedanNalog(P & "OTK-V1", P & "BLOK-V1", 600.004), otvoreno)) = 0, _
        S & "600.004 se zaokruzuje na 600.00 -> prolazi"

    ' 6) Nepostojeca mapa otvorenih iznosa -> odbijanje, ne tiho propustanje.
    Chk LenB(ValidateNalogSaldo(ok, Nothing)) > 0, S & "bez mape otvorenih se odbija"
End Sub

' ============================================================
' T14 - AUD-026: kapija je UNUTAR putanje koja gradi CSV.
'
' T13 gadja samo validator. Da uklanjanje ili preskakanje tog poziva ne bi
' prosla nezapazeno, ovde se testira BuildNalogCsvPayload -- ista funkcija
' koju GenerisiNalogeCSV koristi da napravi sadrzaj fajla, i jedina putanja
' do WriteAllTextUtf8. Preplata mora da da PRAZAN payload (nema sta da se
' upise), a ispravan izbor payload sa tacno onim iznosima koje kapija odobri.
' Bez diranja diska.
' ============================================================
Private Sub T14_CsvPayloadNosiKapiju()
    Const S As String = "T14 CSV payload: "
    Const RACUN As String = "160-1111111111-11"

    Dim otvoreno As Object
    Set otvoreno = CreateObject("Scripting.Dictionary")
    otvoreno(P & "OTK-P1") = 600#
    otvoreno(P & "OTK-P2") = 900#

    ' 1) Preplata -> payload je PRAZAN i razlog je popunjen (fajl se ne pise).
    Dim preplata As Collection
    Set preplata = New Collection
    preplata.Add MakeNalog(P & "OTK-P1", P & "BLOK-P1", 600#, True)
    preplata.Add MakeNalog(P & "OTK-P2", P & "BLOK-P2", 900.01, True)

    Dim odbijeno As String
    Dim payload As String
    payload = BuildNalogCsvPayload(preplata, RACUN, otvoreno, odbijeno)

    ChkEq payload, "", S & "preplata daje prazan payload (nema sta da se upise)"
    Chk LenB(odbijeno) > 0, S & "razlog odbijanja je vracen"
    Chk InStr(odbijeno, P & "BLOK-P2") > 0, S & "razlog imenuje problematican blok"

    ' 2) Ispravan izbor -> zaglavlje + tacno 2 reda, iznosi sa decimalnom TACKOM.
    Dim ok As Collection
    Set ok = New Collection
    ok.Add MakeNalog(P & "OTK-P1", P & "BLOK-P1", 600#, True)
    ok.Add MakeNalog(P & "OTK-P2", P & "BLOK-P2", 250.5, True)

    odbijeno = "x"
    payload = BuildNalogCsvPayload(ok, RACUN, otvoreno, odbijeno)

    Chk LenB(payload) > 0, S & "ispravan izbor daje payload"
    ChkEq odbijeno, "", S & "nema razloga odbijanja kad je sve u granicama"
    ChkEq CsvRedova(payload), 2, S & "payload ima tacno 2 naloga"
    Chk InStr(payload, "RacunPlatioca;") = 1, S & "payload pocinje zaglavljem"
    Chk InStr(payload, ";600.00;") > 0, S & "iznos ide sa decimalnom tackom"
    Chk InStr(payload, ";250.50;") > 0, S & "druga stavka je u payload-u"
    Chk InStr(payload, P & "BLOK-P1") > 0, S & "poziv na broj = broj bloka"

    ' 3) Sub-cent iznos se u fajl upisuje NORMALIZOVAN (ista vrednost koju je
    '    kapija odobrila) -- writer i validator ne smeju da se raziidju.
    Dim subCent As Collection
    Set subCent = New Collection
    subCent.Add MakeNalog(P & "OTK-P1", P & "BLOK-P1", 600.004, True)
    payload = BuildNalogCsvPayload(subCent, RACUN, otvoreno, odbijeno)
    Chk InStr(payload, ";600.00;") > 0, S & "600.004 je u fajlu normalizovan na 600.00"

    ' 4) Blok bez tekuceg racuna ne ulazi u fajl (i ne obara ga).
    Dim bezTR As Collection
    Set bezTR = New Collection
    bezTR.Add MakeNalog(P & "OTK-P1", P & "BLOK-P1", 600#, True)
    bezTR.Add MakeNalog(P & "OTK-P2", P & "BLOK-P2", 100#, False)
    payload = BuildNalogCsvPayload(bezTR, RACUN, otvoreno, odbijeno)
    ChkEq CsvRedova(payload), 1, S & "blok bez TR nije u fajlu"
    ChkEq odbijeno, "", S & "blok bez TR ne obara generisanje"

    ' 5) Sub-cent ostatak NE sme da napravi nalog na nula dinara. Otvoreno iz
    '    GetOpenOtkupi je sirovo (kolicina * cena - isplaceno), pa je ostatak
    '    tipa 0.004 RSD legitiman produkcioni rezultat; sirov prag "> 0" bi ga
    '    pustio, a CsvIznos bi ga upisao kao "0.00" -- nevalidan nalog koji
    '    banci moze da obori uvoz celog paketa.
    Dim ostatak As Object
    Set ostatak = CreateObject("Scripting.Dictionary")
    ostatak(P & "OTK-P1") = 600#
    ostatak(P & "OTK-P3") = 0.004

    Dim saOstatkom As Collection
    Set saOstatkom = New Collection
    saOstatkom.Add MakeNalog(P & "OTK-P1", P & "BLOK-P1", 600#, True)
    saOstatkom.Add MakeNalog(P & "OTK-P3", P & "BLOK-P3", 0.004, True)

    payload = BuildNalogCsvPayload(saOstatkom, RACUN, ostatak, odbijeno)
    ChkEq odbijeno, "", S & "sub-cent ostatak nije preplata (ne obara generisanje)"
    ChkEq CsvRedova(payload), 1, S & "sub-cent ostatak ne pravi svoj nalog"
    Chk InStr(payload, ";0.00;") = 0, S & "u fajlu nema naloga na 0.00"
    Chk InStr(payload, P & "BLOK-P3") = 0, S & "blok sa ostatkom nije u fajlu"
    Chk InStr(payload, ";600.00;") > 0, S & "regularan nalog je i dalje tu"
End Sub

' ============================================================
' T15 - AUD-026: dupli kanonski kljuc obara saldo-mapu (fail-closed).
'
' GetOpenOtkupi ne agregira po OtkupID nego vraca red po red. Kod korumpiranog
' PK-a (dva reda, isti OtkupID, RAZLICIT broj dokumenta i otvoreni iznos) mapa
' bi -- da se gradi assignment-om -- pustila da poslednji red odluci saldo, pa
' bi otvoreno bloka B odobrilo nalog za blok A. Zato BuildOpenAmountDict na
' duplikat PADA; posledica je da ni payload ne nastane (mapa se gradi PRE
' njega), dakle nema fajla za banku.
'
' Namerno se testiraju KONFLIKTNE vrednosti -- to je slucaj koji tiho menja
' finansijski ishod.
' ============================================================
Private Sub T15_DupliOtkupIDObaraSaldoMapu()
    Const S As String = "T15 dupli OtkupID: "

    ' Zdrava lista i dalje mora da prodje (da kapija ne bude preosetljiva).
    Dim zdrava As Collection
    Set zdrava = New Collection
    zdrava.Add MakeBlok(P & "OTK-D1", P & "BLOK-D1", 500)
    zdrava.Add MakeBlok(P & "OTK-D2", P & "BLOK-D2", 2000)

    Dim mapa As Object
    Dim errNum As Long

    On Error Resume Next
    Err.Clear
    Set mapa = BuildOpenAmountDict(zdrava)
    errNum = Err.Number
    On Error GoTo 0
    ChkEq errNum, 0, S & "zdrava lista gradi mapu bez greske"
    If errNum = 0 Then ChkEq mapa.count, 2, S & "mapa ima oba bloka"

    ' Dva reda, ISTI OtkupID, razlicit dokument i razlicit otvoreni iznos.
    Dim korumpirana As Collection
    Set korumpirana = New Collection
    korumpirana.Add MakeBlok(P & "OTK-D9", P & "BLOK-A", 500)
    korumpirana.Add MakeBlok(P & "OTK-D9", P & "BLOK-B", 2000)

    On Error Resume Next
    Err.Clear
    Set mapa = BuildOpenAmountDict(korumpirana)
    errNum = Err.Number
    Dim errDesc As String
    errDesc = Err.description
    On Error GoTo 0

    Chk errNum <> 0, S & "dupli OtkupID PODIZE gresku (ne bira poslednji red)"
    ChkEq errNum, ERR_ISPLATA_DUPLI_OTKUPID, S & "greska je tipizirana"
    Chk InStr(errDesc, P & "OTK-D9") > 0, S & "poruka imenuje sporan OtkupID"

    ' Klamp koristi istu mapu -> ne sme da proguta pad (inace bi pregled tiho
    ' odlucio koji duplikat vazi).
    Dim ovr As Object
    Set ovr = CreateObject("Scripting.Dictionary")
    ovr(P & "OTK-D9") = 1000#

    On Error Resume Next
    Err.Clear
    ClampOverridesToOpen ovr, korumpirana
    errNum = Err.Number
    On Error GoTo 0
    Chk errNum <> 0, S & "ClampOverridesToOpen propagira pad (ne guta ga)"

    ' Prazan OtkupID je ista klasa problema -- saldo se ne moze vezati za blok.
    Dim praznaID As Collection
    Set praznaID = New Collection
    praznaID.Add MakeBlok("", P & "BLOK-D0", 500)

    On Error Resume Next
    Err.Clear
    Set mapa = BuildOpenAmountDict(praznaID)
    errNum = Err.Number
    On Error GoTo 0
    ChkEq errNum, ERR_ISPLATA_PRAZAN_OTKUPID, S & "prazan OtkupID PODIZE tipiziranu gresku"
End Sub

' Broj DATA redova u CSV payload-u (bez zaglavlja i bez zavrsnog praznog reda).
Private Function CsvRedova(ByVal payload As String) As Long
    If LenB(payload) = 0 Then Exit Function
    Dim parts As Variant
    parts = Split(payload, vbCrLf)
    Dim i As Long, n As Long
    For i = LBound(parts) To UBound(parts)
        If LenB(Trim$(CStr(parts(i)))) > 0 Then n = n + 1
    Next i
    If n > 0 Then n = n - 1          ' zaglavlje
    CsvRedova = n
End Function

' Kolekcija sa tacno jednim nalogom -- za granicne provere u T13.
Private Function JedanNalog(ByVal otkupID As String, ByVal brDok As String, _
                            ByVal iznos As Double) As Collection
    Dim c As Collection
    Set c = New Collection
    c.Add MakeNalog(otkupID, brDok, iznos, True)
    Set JedanNalog = c
End Function

' Blok za RF-10 seam testove (otvoreni iznos je jedino sto clamp gleda).
Private Function MakeBlok(ByVal otkupID As String, ByVal brDok As String, _
                          ByVal otvoreno As Double) As clsBlokIsplata
    Dim blk As clsBlokIsplata
    Set blk = New clsBlokIsplata
    blk.otkupID = otkupID
    blk.brojDokumenta = brDok
    blk.kooperantNaziv = "Test Kooperant"
    blk.OtvorenIznos = otvoreno
    Set MakeBlok = blk
End Function

' Blok pripremljen za nalog: IsplatitiIznos + TR (sto GenerisiNalogeCSV gleda).
Private Function MakeNalog(ByVal otkupID As String, ByVal brDok As String, _
                           ByVal iznos As Double, ByVal imaTR As Boolean) As clsBlokIsplata
    Dim blk As clsBlokIsplata
    Set blk = MakeBlok(otkupID, brDok, iznos)
    blk.IsplatitiIznos = iznos
    blk.HasTekuciRacun = imaTR
    If imaTR Then blk.TekuciRacun = "160-0000000000-11"
    Set MakeNalog = blk
End Function

' ============================================================
' T16 - AUD-026: "Primeni avans" odbija dupli kanonski OtkupID.
'
' Tabelarni test (ne dictionary): ApplyAvansToOtkup je citao target red kao
' FindRows(...)(1), dakle "prvi red pobedjuje". Kod dupliranog OtkupID-a to
' znaci da se vrednost otkupa uzima sa JEDNOG reda, a avans se u ledgeru vezuje
' samo za dvosmislen OtkupID -- akcija nad prikazanim redom moze da se izvrsi
' nad drugim redom. Ekran isplata nudi bas tu akciju ("Primeni avans na blok" /
' "(sel.)"), pa kapija mora da bude u CORE-u, ne u formi.
'
' Testira se ISTI kooperant na oba reda: to je slucaj koji je ranije TIHO
' prolazio. Razlicit kooperant je vec obarao zatecen owner guard (AUD-010), pa
' se i on proverava -- da ta odbrana ne oslabi.
' ============================================================
Private Sub T16_DupliOtkupIDNeDozvoljavaAvans()
    Const S As String = "T16 avans nad duplim OtkupID: "

    SeedStanica P & "OM-16", P & "Stanica 16"
    SeedKooperant P & "K-16A", "Test", "Duplikat", P & "OM-16"
    SeedKooperant P & "K-16B", "Test", "Drugi", P & "OM-16"

    ' --- (1) ISTI kooperant, dva reda pod istim OtkupID, razlicit blok/iznos.
    SeedOtkup P & "OTK-DUP", P & "K-16A", P & "BLOK-16A", 100, 5, "Malina"   ' 500
    SeedOtkup P & "OTK-DUP", P & "K-16A", P & "BLOK-16B", 100, 20, "Kupina"  ' 2000
    SeedAvansKooperanta P & "NOV-16", P & "K-16A", 300

    Dim primenjeno As Double
    Dim ok As Boolean
    primenjeno = -1
    ok = ApplyAvansToOtkup_TX(P & "K-16A", P & "OTK-DUP", primenjeno)

    Chk Not ok, S & "avans nad duplim OtkupID je ODBIJEN"
    ChkEqD primenjeno, 0, S & "nista nije proknjizeno (appliedAmount = 0)"
    ChkEq NovacRedovaSaOtkupID(P & "OTK-DUP"), 0, S & "nijedan Novac red nije vezan za sporan OtkupID"
    ChkEqD SlobodanAvansKooperanta(P & "K-16A"), 300, S & "avans kooperanta je ostao netaknut"
    ChkEq OtkupRedova(P & "OTK-DUP"), 2, S & "oba Otkup reda su ostala netaknuta"

    ' --- (2) Razlicit kooperant na duplikatu: i dalje odbijeno (owner guard
    '         AUD-010 je gadjao bas to; nova kapija ga ne sme zameniti tise).
    SeedOtkup P & "OTK-DUP2", P & "K-16A", P & "BLOK-16C", 100, 5, "Malina"
    SeedOtkup P & "OTK-DUP2", P & "K-16B", P & "BLOK-16D", 100, 20, "Kupina"
    SeedAvansKooperanta P & "NOV-16B", P & "K-16B", 300

    primenjeno = -1
    ok = ApplyAvansToOtkup_TX(P & "K-16B", P & "OTK-DUP2", primenjeno)

    Chk Not ok, S & "duplikat sa drugim kooperantom je takodje odbijen"
    ChkEqD primenjeno, 0, S & "nista nije proknjizeno ni u tom slucaju"
    ChkEq NovacRedovaSaOtkupID(P & "OTK-DUP2"), 0, S & "ledger ostaje cist"

    ' --- (3) KONTROLA: jedinstven OtkupID mora i dalje da radi (kapija ne sme
    '         da blokira zdrav slucaj). Zaseban kooperant, da avans iz slucaja
    '         (1) -- koji je ostao NEVEZAN jer je akcija odbijena -- ne udje u
    '         ovaj zbir i ne pomeri ocekivan iznos.
    SeedKooperant P & "K-16C", "Test", "Zdrav", P & "OM-16"
    SeedOtkup P & "OTK-OK16", P & "K-16C", P & "BLOK-16E", 100, 5, "Malina"   ' 500
    SeedAvansKooperanta P & "NOV-16C", P & "K-16C", 200

    primenjeno = -1
    ok = ApplyAvansToOtkup_TX(P & "K-16C", P & "OTK-OK16", primenjeno)

    Chk ok, S & "jedinstven OtkupID i dalje prolazi"
    ChkEqD primenjeno, 200, S & "primenjen je ceo raspolozivi avans (200)"
    Chk NovacRedovaSaOtkupID(P & "OTK-OK16") > 0, S & "avans je vezan za otkup"
    ChkEqD SlobodanAvansKooperanta(P & "K-16C"), 0, S & "avans vise nije nevezan"
End Sub

' ============================================================
' T17 - AUD-026: SAKRIVEN duplikat ne sme da isplati pogresnog kooperanta.
'
' Najopasnija varijanta, koju stroga saldo-mapa NE vidi: duplikat postoji u
' sirovoj tblOtkup, ali samo JEDAN od dva reda je otvoren.
'
'   red A: OTK-HID, kooperant K-17A (racun 160-AAA), blok BLOK-17A, Isplaceno
'   red B: OTK-HID, kooperant K-17B (racun 160-BBB), blok BLOK-17B, otvoren
'
' GetOpenOtkupi vrati samo B, pa u otvorenoj listi nema duplikata i
' BuildOpenAmountDict nema sta da prijavi. Ali vlasnik se ranije citao
' first-wins iz SIROVE tabele -> blok B bi dobio naziv i TEKUCI RACUN
' kooperanta A, i nalog od ~2.000 bi otisao POGRESNOM PRIMAOCU.
'
' Test zato daje kooperantima ocigledno razlicite racune i tvrdi da nikakav
' payload nije nastao -- ne samo da je iznos tacan.
' ============================================================
Private Sub T17_SakrivenDuplikatNeIsplacujePogresnog()
    Const S As String = "T17 sakriven duplikat: "
    Const RACUN_A As String = "160-AAAAAAAAAA-11"
    Const RACUN_B As String = "160-BBBBBBBBBB-22"

    SeedStanica P & "OM-17", P & "Stanica 17"
    SeedKooperantSaRacunom P & "K-17A", "Prvi", "Vlasnik", P & "OM-17", RACUN_A
    SeedKooperantSaRacunom P & "K-17B", "Drugi", "Vlasnik", P & "OM-17", RACUN_B

    Dim lista As Collection
    Dim errNum As Long
    Dim errDesc As String

    ' --- (1) KONTROLA PRVA: duplikat kome NIJEDAN red nije otvoren ne sme da
    '         obori pregled (istorijski duplikat ne proizvodi nalog).
    SeedOtkupIsplacen P & "OTK-HIST", P & "K-17A", P & "BLOK-17C", 100, 5, "Malina"
    SeedOtkupIsplacen P & "OTK-HIST", P & "K-17B", P & "BLOK-17D", 100, 20, "Kupina"

    On Error Resume Next
    Err.Clear
    Set lista = BuildBlokIsplataList()
    errNum = Err.Number
    On Error GoTo 0
    ChkEq errNum, 0, S & "zatvoren (istorijski) duplikat NE obara pregled"

    ' --- (2) Sada pravi slucaj: red A isplacen (duplikat je skriven od
    '         GetOpenOtkupi), red B otvoren i na DRUGOG kooperanta.
    SeedOtkupIsplacen P & "OTK-HID", P & "K-17A", P & "BLOK-17A", 100, 5, "Malina"
    SeedOtkup P & "OTK-HID", P & "K-17B", P & "BLOK-17B", 100, 20, "Kupina"

    On Error Resume Next
    Err.Clear
    Set lista = BuildBlokIsplataList()
    errNum = Err.Number
    errDesc = Err.description
    On Error GoTo 0

    ChkEq errNum, ERR_ISPLATA_DUPLI_OTKUPID, S & "otvoren blok sa dupliranim OtkupID obara listu"
    Chk InStr(errDesc, P & "OTK-HID") > 0, S & "poruka imenuje sporan OtkupID"

    ' --- (3) I cela CSV putanja mora ostati bez fajla. Nalog se pravi za
    '         BLOK-17B (otvoren red), a GenerisiNalogeCSV interno gradi svezu
    '         listu -> pad pre nego sto payload uopste nastane.
    Dim nalozi As Collection
    Set nalozi = New Collection
    nalozi.Add MakeNalog(P & "OTK-HID", P & "BLOK-17B", 1500, True)

    Dim odbijeno As String
    Dim putanja As String
    putanja = GenerisiNalogeCSV(nalozi, "160-1111111111-11", odbijeno)

    ChkEq putanja, "", S & "nijedan CSV fajl nije napravljen"
    Chk LenB(odbijeno) > 0, S & "razlog je vracen operateru"
    Chk InStr(odbijeno, P & "OTK-HID") > 0, S & "razlog imenuje sporan OtkupID"

End Sub

' ============================================================
' SEED / READ HELPERS
' ============================================================

' Kooperant sa tekucim racunom (T17: primalac naloga mora biti razlucljiv).
Private Sub SeedKooperantSaRacunom(ByVal koopID As String, ByVal ime As String, _
                                   ByVal prezime As String, ByVal stanicaID As String, _
                                   ByVal racun As String)
    BitAppend TBL_KOOPERANTI, _
        Array("KooperantID", "Ime", "Prezime", COL_KOOP_STANICA, COL_KOOP_TEKUCI_RACUN), _
        Array(koopID, ime, prezime, stanicaID, racun)
End Sub

' Otkup red koji je VEC ISPLACEN -> GetOpenOtkupi ga preskace.
Private Sub SeedOtkupIsplacen(ByVal otkID As String, ByVal koopID As String, _
                              ByVal brDok As String, ByVal kolicina As Double, _
                              ByVal cena As Double, ByVal vrsta As String)
    BitAppend TBL_OTKUP, _
        Array(COL_OTK_ID, COL_OTK_BR_DOK, COL_OTK_KOOPERANT, COL_OTK_KOLICINA, _
              COL_OTK_CENA, COL_OTK_VRSTA, COL_OTK_DATUM, COL_OTK_ISPLACENO), _
        Array(otkID, brDok, koopID, kolicina, cena, vrsta, Date, STATUS_ISPLACENO)
End Sub

' Slobodan (nevezan) virman avans kooperanta -- ulaz za ApplyAvansToOtkup.
Private Sub SeedAvansKooperanta(ByVal novacID As String, ByVal koopID As String, _
                                ByVal iznos As Double)
    BitAppend TBL_NOVAC, _
        Array(COL_NOV_ID, COL_NOV_DATUM, COL_NOV_KOOP_ID, COL_NOV_PARTNER_ID, _
              COL_NOV_ENTITET_TIP, COL_NOV_TIP, COL_NOV_ISPLATA, COL_NOV_OTKUP_ID), _
        Array(novacID, Date, koopID, koopID, "Kooperant", NOV_VIRMAN_AVANS_KOOP, iznos, "")
End Sub

' Koliko Novac redova nosi dati OtkupID (0 = nista nije vezano).
Private Function NovacRedovaSaOtkupID(ByVal otkupID As String) As Long
    Dim data As Variant
    data = GetTableData(TBL_NOVAC)
    If IsEmpty(data) Then Exit Function

    Dim c As Long
    c = GetColumnIndex(TBL_NOVAC, COL_NOV_OTKUP_ID)
    If c = 0 Then Exit Function

    Dim i As Long, n As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, c))) = otkupID Then n = n + 1
    Next i
    NovacRedovaSaOtkupID = n
End Function

' Zbir NEVEZANOG avansa kooperanta (OtkupID prazan).
Private Function SlobodanAvansKooperanta(ByVal koopID As String) As Double
    Dim data As Variant
    data = GetTableData(TBL_NOVAC)
    If IsEmpty(data) Then Exit Function

    Dim cK As Long, cT As Long, cI As Long, cO As Long
    cK = GetColumnIndex(TBL_NOVAC, COL_NOV_KOOP_ID)
    cT = GetColumnIndex(TBL_NOVAC, COL_NOV_TIP)
    cI = GetColumnIndex(TBL_NOVAC, COL_NOV_ISPLATA)
    cO = GetColumnIndex(TBL_NOVAC, COL_NOV_OTKUP_ID)
    If cK = 0 Or cT = 0 Or cI = 0 Or cO = 0 Then Exit Function

    Dim i As Long, sum As Double
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cK))) = koopID Then
            If CStr(data(i, cT)) = NOV_VIRMAN_AVANS_KOOP Then
                If LenB(Trim$(CStr(data(i, cO)))) = 0 Then
                    If IsNumeric(data(i, cI)) Then sum = sum + CDbl(data(i, cI))
                End If
            End If
        End If
    Next i
    SlobodanAvansKooperanta = sum
End Function

' Koliko redova u tblOtkup nosi dati OtkupID (za proveru da nista nije diralo).
Private Function OtkupRedova(ByVal otkupID As String) As Long
    Dim rowsFound As Collection
    Set rowsFound = FindRows(TBL_OTKUP, COL_OTK_ID, otkupID)
    If Not rowsFound Is Nothing Then OtkupRedova = rowsFound.count
End Function

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

' Novac red koji ZATVARA fakturu (koristi ga T10: auto put nadje placenu fakturu).
Private Sub SeedNovacUplataFakture(ByVal novacID As String, ByVal kupacID As String, _
                                   ByVal fakturaID As String, ByVal iznos As Double)
    BitAppend TBL_NOVAC, _
        Array(COL_NOV_ID, COL_NOV_DATUM, COL_NOV_PARTNER_ID, COL_NOV_ENTITET_TIP, _
              COL_NOV_FAKTURA_ID, COL_NOV_TIP, COL_NOV_UPLATA), _
        Array(novacID, Date, kupacID, "Kupac", fakturaID, NOV_KUPCI_UPLATA, iznos)
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

' 17-kolonski staging red u obliku koji vraca ParseBankaIzvodForImport.
Private Function BuildStagingRow(ByVal brojIzvoda As String, ByVal racun As String, _
                                 ByVal datumIzvoda As Variant, ByVal datumTx As Variant, _
                                 ByVal referenz As String) As Variant
    Dim red(1 To 1, 1 To 17) As Variant

    red(1, 1) = brojIzvoda        ' BrojDokumenta
    red(1, 2) = datumIzvoda       ' DatumIzvoda
    red(1, 3) = racun             ' BrojRacuna
    red(1, 4) = datumTx           ' DatumTransakcije
    red(1, 5) = P & "PARTNER-T"   ' Partner
    red(1, 6) = ""                ' PartnerKonto
    red(1, 7) = 1234.56           ' Uplata
    red(1, 8) = 0                 ' Isplata
    red(1, 9) = ""                ' Sifra
    red(1, 10) = "test"           ' Opis / SvrhaPlacanja
    red(1, 11) = ""               ' PozivNaBroj
    red(1, 12) = referenz         ' BankaReferenz
    red(1, 13) = "test.pdf"       ' IzvorFajl
    red(1, 14) = 0
    red(1, 15) = 0
    red(1, 16) = 0
    red(1, 17) = 0

    BuildStagingRow = red
End Function

' Dva staging reda istog izvoda: prvi validan, drugi sa zadatim (lose) datumom.
Private Function BuildStagingRowPair(ByVal brojIzvoda As String, ByVal racun As String, _
                                     ByVal datumOk As Variant, _
                                     ByVal datumLos As Variant) As Variant
    Dim par(1 To 2, 1 To 17) As Variant
    Dim prvi As Variant
    Dim drugi As Variant
    Dim c As Long

    prvi = BuildStagingRow(brojIzvoda, racun, datumOk, datumOk, P & "REF-P1")
    drugi = BuildStagingRow(brojIzvoda, racun, datumOk, datumLos, P & "REF-P2")

    For c = 1 To 17
        par(1, c) = prvi(1, c)
        par(2, c) = drugi(1, c)
    Next c

    BuildStagingRowPair = par
End Function

' Koliko staging redova nosi dati broj izvoda (za proveru da odbijen red NIJE upisan).
Private Function StagingRedova(ByVal brojIzvoda As String) As Long
    Dim data As Variant
    Dim colBrojDok As Long
    Dim i As Long

    data = GetTableData(TBL_BANKA_IMPORT)
    If IsEmpty(data) Then Exit Function

    colBrojDok = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_BROJ_DOKUMENTA)

    For i = 1 To UBound(data, 1)
        If Trim$(CStr(NzTb(data(i, colBrojDok)))) = brojIzvoda Then
            StagingRedova = StagingRedova + 1
        End If
    Next i
End Function

Private Function NovacPartnerID(ByVal novacID As String) As String
    NovacPartnerID = Trim$(CStr(NzTb(LookupValue(TBL_NOVAC, COL_NOV_ID, novacID, COL_NOV_PARTNER_ID))))
End Function

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

' Ostavi otvorenim SAMO navedene staging redove; sve ostale otvorene -- i zatecene
' iz baze i one koje su RANIJE test grupe seed-ovale -- oznaci "Skip".
'
' Bez ovoga jedna grupa curi u sledecu: T08 namerno ostavlja red BEZ broja izvoda
' (da obori batch), pa ga je T10 pokupio u svoj AutoMapAll i pao na tudjem redu
' (ERR_BMAP_BASE + 45), sto je izgledalo kao da fiks ne radi. Rollback suite-a
' vraca originalne statuse.
Private Sub OstaviOtvorenimSamo(ByVal bimIDs As Variant)
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
        bimID = Trim$(CStr(NzTb(data(i, colID))))

        If Trim$(CStr(NzTb(data(i, colObr)))) = "" Then
            If Not SadrziID(bimIDs, bimID) Then
                UpdateCell TBL_BANKA_IMPORT, i, COL_BIM_OBRADJENO, "Skip"
            End If
        End If
    Next i
End Sub

Private Function SadrziID(ByVal bimIDs As Variant, ByVal traziID As String) As Boolean
    Dim i As Long

    For i = LBound(bimIDs) To UBound(bimIDs)
        If CStr(bimIDs(i)) = traziID Then
            SadrziID = True
            Exit Function
        End If
    Next i
End Function

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
    hdr = "BANKA TEST SUITE (RF-09 + RF-10)  ->  PASS=" & mPass & "  FAIL=" & mFail

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
