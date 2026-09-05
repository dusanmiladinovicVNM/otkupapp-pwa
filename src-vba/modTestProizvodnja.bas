Attribute VB_Name = "modTestProizvodnja"
Option Explicit

' ============================================================
' modTestProizvodnja -- suite za Preradu 2.0 (proizvodno jezgro).
' Spec: docs/PRERADA_2_MODEL_I_PLAN.md par. 14.2.
'
' POKRETANJE: Alt+F8 -> RunProizvodnjaTestSuite
'
' BEZBEDNOST: ceo suite radi UNUTAR JEDNE clsTransaction (obrazac
' modTestPalete): snapshot svih tabela koje testovi diraju, a na kraju UVEK
' RollbackTx -- i kad svi produ i kad neki padne. Test podaci nose
' "TST-PRZ-" identitet. Sema (tabele) se ne vraca rollback-om: suite je
' pokrece kroz EnsureProizvodnjaSchemaCore, sto je idempotentno i isto sto
' radi self-heal na svakom startu -- na fixture-u tabele vec postoje.
'
' STA POKRIVA (Faza A):
'  T01 sema idempotentna (seed tipova/proizvoda; drugi poziv = 0 novih)
'  T02 materijalizacija legacy prerada (aktivna sa tipom, aktivna bez tipa,
'      stornirana; obrnuti pokazivac; backfill LJ na utovarnoj stavci;
'      dupli PreradaID se preskace; MaterijalizujPreradu idempotentan i
'      fail-closed na dupli)
'  T03 raspolozivo = jedna funkcija (fizicko - utovareno; storno utovara
'      vraca stanje; PALETA poreklo cita zivu tblPaleta.NetoKg; dupli LJ = -1)
'  T17 oznaka jedinice + rok kao SNAPSHOT (promena pravila ga ne menja)
'
' Rezultat: MsgBox rezime + izvestaj na listu "_TestProizvodnja" + Immediate.
' ============================================================

Private Const TST_PRE1 As String = "TST-PRZ-PRE-1"   ' aktivna, tip poznat, snapshot roka
Private Const TST_PRE2 As String = "TST-PRZ-PRE-2"   ' aktivna, BEZ tipa GP
Private Const TST_PRE3 As String = "TST-PRZ-PRE-3"   ' stornirana, bez snapshota roka
Private Const TST_DUP As String = "TST-PRZ-DUP"      ' dupli PreradaID (korupcija)
Private Const TST_TIP As String = "TST-PRZ-GP"       ' vrsta gotovog proizvoda (RokMeseci=6)
Private Const TST_UT As String = "TST-PRZ-UT-1"
Private Const TST_UTS As String = "TST-PRZ-UTS-1"
Private Const TST_PAL As String = "TST-PRZ-PAL-1"
Private Const TST_LJ_PAL As String = "TST-PRZ-LJ-PAL"
Private Const TST_LJ_DUP As String = "TST-PRZ-LJ-DUP"

' Gate: suite mora da PODIGNE gresku i kad provera padne i kad se ne pokrene
' (konvencija modTestBanka.ERR_BIT_SUITE_FAILED; zauzeti offseti u
' docs/EXCEL_TEST_HARNESS.md).
Private Const ERR_PRZ_SUITE_FAILED As Long = vbObjectError + 2970

Private mPass As Long
Private mFail As Long
Private mFails As String
Private mReport As String

Public Sub RunProizvodnjaTestSuite()
    Dim tx As clsTransaction
    On Error GoTo EH

    If FindRows(TBL_PRERADA, COL_PRE_ID, TST_PRE1).count > 0 Then
        MsgBox "Nadjeni su TST-PRZ- test podaci od ranije (verovatno prekinut test-run)." & vbCrLf & _
               "Zatvori svesku BEZ snimanja i otvori ponovo, pa pokreni test iz cistog stanja.", _
               vbExclamation, APP_NAME
        Err.Raise ERR_PRZ_SUITE_FAILED, "modTestProizvodnja.RunProizvodnjaTestSuite", _
            "suite NIJE pokrenut: zatecen TST-PRZ- ostatak od prekinutog run-a"
    End If

    If MsgBox("Pokrece PROIZVODNJA (Prerada 2.0) test-suite na OVOJ radnoj svesci." & vbCrLf & vbCrLf & _
              "Sve izmene se rade u jednoj transakciji i UVEK se ponistavaju" & vbCrLf & _
              "(i na uspehu i na gresci) - podaci ostaju netaknuti." & vbCrLf & vbCrLf & _
              "Nastaviti?", vbYesNo + vbQuestion, APP_NAME) <> vbYes Then
        Err.Raise ERR_PRZ_SUITE_FAILED, "modTestProizvodnja.RunProizvodnjaTestSuite", _
            "suite NIJE pokrenut: operater odustao na potvrdi"
    End If

    mPass = 0: mFail = 0: mFails = "": mReport = ""

    ' Sema pre snapshota: tabele koje fale nastaju sada (idempotentno, kao
    ' self-heal), pa snapshot obuhvata i njih.
    EnsureProizvodnjaSchemaCore

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_PRERADA
    tx.AddTableSnapshot TBL_LAGER_JEDINICE
    tx.AddTableSnapshot TBL_PROIZVODI
    tx.AddTableSnapshot TBL_TIPOVI_PROCESA
    tx.AddTableSnapshot TBL_VRSTA_GP
    tx.AddTableSnapshot TBL_PALETA
    If Not GetTable(TBL_UTOVAR) Is Nothing Then tx.AddTableSnapshot TBL_UTOVAR
    If Not GetTable(TBL_UTOVAR_STAVKE) Is Nothing Then tx.AddTableSnapshot TBL_UTOVAR_STAVKE
    If Not GetTable(TBL_FAKTURA_STAVKE) Is Nothing Then tx.AddTableSnapshot TBL_FAKTURA_STAVKE

    T01_SemaIdempotentna
    T02_MaterijalizacijaLegacyPrerade
    T03_RaspolozivoJednaFunkcija
    T17_LjOznakaIRok

    tx.RollbackTx
    Set tx = Nothing
    On Error GoTo 0
    ReportResults True

    If mFail > 0 Then
        Err.Raise ERR_PRZ_SUITE_FAILED, "modTestProizvodnja.RunProizvodnjaTestSuite", _
            "RunProizvodnjaTestSuite: " & CStr(mFail) & " provera palo (PASS=" & _
            CStr(mPass) & "). Detalji u Immediate prozoru."
    End If
    Exit Sub

EH:
    Dim eDesc As String: eDesc = Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    Set tx = Nothing
    Fail "NEOCEKIVANA GRESKA (suite prekinut): " & eDesc
    ReportResults False
    On Error GoTo 0
    Err.Raise ERR_PRZ_SUITE_FAILED, "modTestProizvodnja.RunProizvodnjaTestSuite", _
        "RunProizvodnjaTestSuite prekinut: " & eDesc
End Sub

' ============================================================
' TESTOVI
' ============================================================

Private Sub T01_SemaIdempotentna()
    Dim s As String, nT As Long, nP As Long, nL As Long
    s = EnsureProizvodnjaSchemaSve()
    nT = RowCount(TBL_TIPOVI_PROCESA): nP = RowCount(TBL_PROIZVODI): nL = RowCount(TBL_LAGER_JEDINICE)
    Chk nT >= 12, "T01 seed tipova procesa daje bar 12 redova (ima " & nT & ")"
    ChkEq FindRows(TBL_TIPOVI_PROCESA, COL_TPR_SIFRA, "ZAMRZAVANJE").count, 1, "T01 ZAMRZAVANJE postoji tacno jednom"
    ChkEq FindRows(TBL_TIPOVI_PROCESA, COL_TPR_SIFRA, "PRERADA_LEGACY").count, 1, "T01 PRERADA_LEGACY postoji tacno jednom"
    Chk GetColumnIndex(TBL_PRERADA, COL_PRE_LJ_ID) > 0, "T01 tblPrerada.LagerJedinicaID postoji"
    Chk GetColumnIndex(TBL_LAGER_JEDINICE, COL_STORNIRANO) > 0, "T01 tblLagerJedinice.Stornirano postoji"
    Chk TabelaNosiStorno(TBL_LAGER_JEDINICE), "T01 tblLagerJedinice je u registru storna"
    Chk TabelaNosiStorno(TBL_PROCES_SARZE), "T01 tblProcesSarze je u registru storna"
    Chk Not TabelaNosiStorno(TBL_PROIZVODI), "T01 tblProizvodi je maticna (bez storna)"
    Chk StornoRegistarZna(TBL_PROIZVODI), "T01 tblProizvodi je klasifikovana u registru"

    s = EnsureProizvodnjaSchemaSve()
    ChkEq RowCount(TBL_TIPOVI_PROCESA), nT, "T01 drugi poziv: tipovi procesa nepromenjeni"
    ChkEq RowCount(TBL_PROIZVODI), nP, "T01 drugi poziv: proizvodi nepromenjeni"
    ChkEq RowCount(TBL_LAGER_JEDINICE), nL, "T01 drugi poziv: lager jedinice nepromenjene"
End Sub

Private Sub T02_MaterijalizacijaLegacyPrerade()
    Const SRC As String = "modTestProizvodnja.T02"
    ' Vrsta GP sa rokom 6 meseci -> proizvod nastaje seed-om.
    TstAppend TBL_VRSTA_GP, Array(COL_VGP_TIP, "Aktivan", COL_VGP_ROK), Array(TST_TIP, STATUS_AKTIVAN, 6)
    SeedProizvodi
    Dim prz As String: prz = ProizvodIDZaTipGP(TST_TIP)
    Chk Len(prz) > 0, "T02 seed proizvoda iz nove vrste GP (" & prz & ")"

    Dim rokSnap As Date: rokSnap = DateAdd("d", 100, Date)
    SeedPrerada TST_PRE1, 9001, 100, TST_TIP, "", rokSnap
    SeedPrerada TST_PRE2, 9002, 50, "", "", Empty
    SeedPrerada TST_PRE3, 9003, 40, TST_TIP, "Da", Empty
    SeedPrerada TST_DUP, 9004, 10, TST_TIP, "", Empty
    SeedPrerada TST_DUP, 9005, 10, TST_TIP, "", Empty

    If Not GetTable(TBL_UTOVAR_STAVKE) Is Nothing And Not GetTable(TBL_UTOVAR) Is Nothing Then
        TstAppend TBL_UTOVAR, Array(COL_UT_ID, COL_UT_BROJ, COL_UT_GODINA, COL_UT_DATUM, COL_UT_KUPAC, COL_STORNIRANO), _
                  Array(TST_UT, 9001, Year(Date), Date, "TST-PRZ-KUP", "")
        TstAppend TBL_UTOVAR_STAVKE, Array(COL_UTS_ID, COL_UTS_UTOVAR_ID, COL_UTS_PRERADA_ID, COL_UTS_BROJ_PRERADE, COL_UTS_KOLICINA, COL_STORNIRANO), _
                  Array(TST_UTS, TST_UT, TST_PRE1, "9001/" & Year(Date), 30, "")
    End If

    Dim n As Long
    n = MaterijalizujLegacyPrerade()
    ChkEq n, 3, "T02 materijalizovane tacno 3 nove jedinice (dupli PreradaID preskocen)"

    Dim mapa As Object: Set mapa = LjPoIzvoru(LJ_IZVOR_PRERADA)
    Chk mapa.Exists(TST_PRE1) And mapa.Exists(TST_PRE2) And mapa.Exists(TST_PRE3), "T02 sve tri prerade imaju LJ"
    Chk Not mapa.Exists(TST_DUP), "T02 dupli PreradaID nema LJ (fail-closed, P4)"

    Dim lj1 As String: lj1 = CStr(mapa(TST_PRE1))
    ChkEqD NzD(LjVrednost(lj1, COL_LJ_KG_POCETNO)), 100, "T02 KgPocetno = NetoIzlazKg"
    ChkEq LjPolje(lj1, COL_LJ_PROIZVOD), prz, "T02 ProizvodID iz tipa GP"
    ChkEq LjPolje(lj1, COL_LJ_IZVOR_TIP), LJ_IZVOR_PRERADA, "T02 IzvorTip=PRERADA"
    ChkEq LjPolje(lj1, COL_LJ_IZVOR_ID), TST_PRE1, "T02 IzvorID=PreradaID"
    ChkEq LjPolje(lj1, COL_LJ_TIP), LJ_TIP_PALETA, "T02 TipJedinice=PALETA"
    ChkEq UCase$(LjPolje(lj1, COL_STORNIRANO)), "", "T02 aktivna prerada -> aktivna LJ"
    ChkEq CStr(LjVrednost(lj1, COL_LJ_ROK)), CStr(rokSnap), "T02 DatumIsteka KOPIRAN sa prerade (snapshot)"
    ChkEq CStr(LookupValue(TBL_PRERADA, COL_PRE_ID, TST_PRE1, COL_PRE_LJ_ID)), lj1, "T02 obrnuti pokazivac tblPrerada.LagerJedinicaID"

    Dim lj2 As String: lj2 = CStr(mapa(TST_PRE2))
    ChkEq LjPolje(lj2, COL_LJ_PROIZVOD), "", "T02 prerada bez tipa -> LJ bez ProizvodID (nije prodajna)"

    Dim lj3 As String: lj3 = CStr(mapa(TST_PRE3))
    ChkEq UCase$(LjPolje(lj3, COL_STORNIRANO)), "DA", "T02 stornirana prerada -> stornirana LJ"
    ChkEq CStr(LjVrednost(lj3, COL_LJ_ROK)), CStr(DateAdd("m", 6, Date)), "T02 bez snapshota -> rok po pravilu vrste (6 m)"

    ChkEq MaterijalizujLegacyPrerade(), 0, "T02 drugi prolaz ne pravi nista"
    ChkEq MaterijalizujPreradu(TST_PRE1, SRC), lj1, "T02 MaterijalizujPreradu vraca postojecu LJ"

    Dim puklo As Boolean
    On Error Resume Next
    MaterijalizujPreradu TST_DUP, SRC
    puklo = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    Chk puklo, "T02 MaterijalizujPreradu nad duplim PreradaID pada (fail-closed)"

    If Not GetTable(TBL_UTOVAR_STAVKE) Is Nothing Then
        Dim nb As Long: nb = BackfillLjNaStavkama()
        Chk nb >= 1, "T02 backfill upisao LJ bar na test stavku (" & nb & ")"
        ChkEq CStr(LookupValue(TBL_UTOVAR_STAVKE, COL_UTS_ID, TST_UTS, COL_UTS_LJ_ID)), lj1, "T02 utovarna stavka nosi LJ prerade"
        ChkEq BackfillLjNaStavkama(), 0, "T02 drugi backfill = 0"
    End If
End Sub

Private Sub T03_RaspolozivoJednaFunkcija()
    Const SRC As String = "modTestProizvodnja.T03"
    Dim mapa As Object: Set mapa = LjPoIzvoru(LJ_IZVOR_PRERADA)
    Dim lj1 As String: lj1 = CStr(mapa(TST_PRE1))
    Dim lj3 As String: lj3 = CStr(mapa(TST_PRE3))
    Dim m As Object: Set m = RaspolozivoPoJedinici()

    If Not GetTable(TBL_UTOVAR_STAVKE) Is Nothing Then
        ChkEqD CDbl(m(lj1)), 70, "T03 raspolozivo = 100 - 30 utovareno"
    Else
        ChkEqD CDbl(m(lj1)), 100, "T03 raspolozivo = fizicko (sveska bez utovara)"
    End If
    ChkEqD CDbl(m(CStr(mapa(TST_PRE2)))), 50, "T03 jedinica bez utovara = fizicko"
    Chk Not m.Exists(lj3), "T03 stornirana jedinica nije u mapi"
    ChkEqD RaspolozivoKg("LJ-NEPOSTOJECA"), 0, "T03 nepoznata jedinica = 0"

    If Not GetTable(TBL_UTOVAR) Is Nothing Then
        Dim r As Collection: Set r = FindRows(TBL_UTOVAR, COL_UT_ID, TST_UT)
        If r.count = 1 Then RequireUpdateCell TBL_UTOVAR, CLng(r(1)), COL_STORNIRANO, "Da", SRC
        Set m = RaspolozivoPoJedinici()
        ChkEqD CDbl(m(lj1)), 100, "T03 storno utovara vraca stanje (izvedeno, nista se ne pise na LJ)"
    End If

    ' PALETA poreklo: fizicko je ZIVA tblPaleta.NetoKg, ne snimak KgPocetno.
    TstAppend TBL_PALETA, Array(COL_PAL_ID, COL_PAL_BROJ, COL_PAL_GODINA, COL_PAL_DATUM, COL_PAL_VRSTA, _
                                COL_PAL_NETO, COL_PAL_STATUS, COL_STORNIRANO), _
              Array(TST_PAL, 31, Year(Date), Date, "TST-PRZ-VOCE", 250, PAL_STATUS_ZATVORENA, "")
    TstAppend TBL_LAGER_JEDINICE, Array(COL_LJ_ID, COL_LJ_BROJ, COL_LJ_GODINA, COL_LJ_TIP, COL_LJ_KG_POCETNO, _
                                        COL_LJ_DATUM, COL_LJ_IZVOR_TIP, COL_LJ_IZVOR_ID, COL_STORNIRANO), _
              Array(TST_LJ_PAL, 31, Year(Date), LJ_TIP_PALETA, 200, Date, LJ_IZVOR_PALETA, TST_PAL, "")
    Set m = RaspolozivoPoJedinici()
    ChkEqD CDbl(m(TST_LJ_PAL)), 250, "T03 PALETA poreklo cita zivu tblPaleta.NetoKg (250, ne 200)"

    ' Dupli LagerJedinicaID nikad nije raspoloziv.
    TstAppend TBL_LAGER_JEDINICE, Array(COL_LJ_ID, COL_LJ_KG_POCETNO, COL_LJ_IZVOR_TIP, COL_LJ_IZVOR_ID, COL_STORNIRANO), _
              Array(TST_LJ_DUP, 10, LJ_IZVOR_SARZA, "SRZ-TST", "")
    TstAppend TBL_LAGER_JEDINICE, Array(COL_LJ_ID, COL_LJ_KG_POCETNO, COL_LJ_IZVOR_TIP, COL_LJ_IZVOR_ID, COL_STORNIRANO), _
              Array(TST_LJ_DUP, 10, LJ_IZVOR_SARZA, "SRZ-TST", "")
    Set m = RaspolozivoPoJedinici()
    ChkEqD CDbl(m(TST_LJ_DUP)), -1, "T03 dupli LagerJedinicaID = -1 (nikad raspoloziv)"
End Sub

Private Sub T17_LjOznakaIRok()
    Const SRC As String = "modTestProizvodnja.T17"
    Dim mapa As Object: Set mapa = LjPoIzvoru(LJ_IZVOR_PRERADA)
    Dim lj1 As String: lj1 = CStr(mapa(TST_PRE1))
    ChkEq LjOznaka(lj1), "PRE 9001/" & Year(Date), "T17 oznaka legacy lota"
    ChkEq LjOznaka(TST_LJ_PAL), "PAL 31/" & Year(Date), "T17 oznaka jedinice sa palete"
    ChkEq LjOznaka("LJ-NEMA"), "LJ-NEMA", "T17 nepoznata jedinica vraca ID"

    Dim rokSnap As Date: rokSnap = DateAdd("d", 100, Date)
    ChkEq CStr(LjRokTrajanja(lj1)), CStr(rokSnap), "T17 rok = snapshot DatumIsteka"

    ' Promena pravila (RokMeseci 6 -> 12) NE menja snapshot.
    Dim r As Collection: Set r = FindRows(TBL_VRSTA_GP, COL_VGP_TIP, TST_TIP)
    If r.count = 1 Then RequireUpdateCell TBL_VRSTA_GP, CLng(r(1)), COL_VGP_ROK, 12, SRC
    ChkEq CStr(LjRokTrajanja(lj1)), CStr(rokSnap), "T17 promena RokMeseci ne menja snapshot"

    ' Bez snapshota: fallback po TEKUCEM pravilu (isti kao stampa), sad 12 m.
    Dim rl As Collection: Set rl = FindRows(TBL_LAGER_JEDINICE, COL_LJ_ID, lj1)
    If rl.count = 1 Then RequireUpdateCell TBL_LAGER_JEDINICE, CLng(rl(1)), COL_LJ_ROK, "", SRC
    ChkEq CStr(LjRokTrajanja(lj1)), CStr(DateAdd("m", 12, Date)), "T17 bez snapshota -> fallback po tekucem pravilu"
    Chk IsEmpty(LjRokTrajanja(TST_LJ_PAL)), "T17 jedinica sa palete bez snapshota -> Empty (ne izmislja se)"
End Sub

' ============================================================
' SEED / POMOCNE
' ============================================================

Private Sub SeedPrerada(ByVal id As String, ByVal broj As Long, ByVal netoIzlaz As Double, _
                        ByVal tipGP As String, ByVal storno As String, ByVal rok As Variant)
    TstAppend TBL_PRERADA, _
        Array(COL_PRE_ID, COL_PRE_BROJ, COL_PRE_GODINA, COL_PRE_DATUM, COL_PRE_NETO_ULAZ, _
              COL_PRE_NETO_IZLAZ, COL_PRE_KUTIJE, COL_PRE_KESE, COL_PRE_TIP_KUTIJE, _
              COL_PRE_TIP_GP, COL_STORNIRANO, COL_PRE_ROK), _
        Array(id, broj, Year(Date), Date, netoIzlaz, _
              netoIzlaz, 10, 20, "TST-KUT", _
              tipGP, storno, rok)
End Sub

Private Function RowCount(ByVal tbl As String) As Long
    Dim d As Variant
    d = GetTableData(tbl)
    If IsArray(d) Then RowCount = UBound(d, 1)
End Function

' Tekst polja jedinice po imenu kolone ("" kad je nema ili je prazno).
Private Function LjPolje(ByVal ljID As String, ByVal colName As String) As String
    LjPolje = Trim$(CStr(nz(LjVrednost(ljID, colName))))
End Function

Private Function LjVrednost(ByVal ljID As String, ByVal colName As String) As Variant
    LjVrednost = Empty
    Dim r As Collection: Set r = FindRows(TBL_LAGER_JEDINICE, COL_LJ_ID, ljID)
    If r.count <> 1 Then Exit Function
    Dim c As Long: c = GetColumnIndex(TBL_LAGER_JEDINICE, colName)
    If c = 0 Then Exit Function
    Dim d As Variant: d = GetTableData(TBL_LAGER_JEDINICE)
    LjVrednost = d(CLng(r(1)), c)
End Function

' Append po IMENU kolone (mirror PalAppendRow; kolone koje ne postoje se
' preskacu -> bezbedno pod schema drift-om). Pada glasno ako append ne uspe.
Private Sub TstAppend(ByVal tblName As String, ByVal cols As Variant, ByVal vals As Variant)
    Dim lo As ListObject
    Set lo = GetTable(tblName)
    If lo Is Nothing Then Err.Raise vbObjectError + 9930, "modTestProizvodnja.TstAppend", _
                                    "Nema tabele: " & tblName
    Dim n As Long: n = lo.ListColumns.count
    Dim rowData() As Variant
    ReDim rowData(0 To n - 1)
    Dim i As Long, idx As Long
    For i = LBound(cols) To UBound(cols)
        idx = GetColumnIndex(tblName, CStr(cols(i)))
        If idx >= 1 And idx <= n Then rowData(idx - 1) = vals(i)
    Next i
    If AppendRow(tblName, rowData) = 0 Then
        Err.Raise vbObjectError + 9931, "modTestProizvodnja.TstAppend", _
                  "AppendRow nije uspeo: " & tblName
    End If
End Sub

' ============================================================
' ASSERT + IZVESTAJ (obrazac modTestPalete)
' ============================================================

Private Sub Chk(ByVal cond As Boolean, ByVal nm As String)
    If cond Then
        mPass = mPass + 1
        mReport = mReport & "OK    " & nm & vbLf
    Else
        Fail nm
    End If
End Sub

Private Sub ChkEq(ByVal act As Variant, ByVal exp As Variant, ByVal nm As String)
    If CStr(act) = CStr(exp) Then
        mPass = mPass + 1
        mReport = mReport & "OK    " & nm & vbLf
    Else
        Fail nm & "  [dobijeno=" & CStr(act) & ", ocekivano=" & CStr(exp) & "]"
    End If
End Sub

Private Sub ChkEqD(ByVal act As Double, ByVal exp As Double, ByVal nm As String)
    If Abs(act - exp) <= 0.001 Then
        mPass = mPass + 1
        mReport = mReport & "OK    " & nm & vbLf
    Else
        Fail nm & "  [dobijeno=" & Format$(act, "0.###") & ", ocekivano=" & Format$(exp, "0.###") & "]"
    End If
End Sub

Private Sub Fail(ByVal nm As String)
    mFail = mFail + 1
    mReport = mReport & "PAO   " & nm & vbLf
    If Len(mFails) > 0 Then mFails = mFails & vbCrLf
    mFails = mFails & "- " & nm
End Sub

Private Sub ReportResults(ByVal clean As Boolean)
    On Error Resume Next
    Debug.Print "===== PROIZVODNJA TEST SUITE ====="
    Debug.Print mReport
    Debug.Print "PROSLO: " & mPass & "   PALO: " & mFail

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("_TestProizvodnja")
    If ws Is Nothing Then
        Err.Clear
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = "_TestProizvodnja"
    End If
    If Not ws Is Nothing Then
        ws.Cells.Clear
        ws.Cells(1, 1).value = "PROIZVODNJA TEST SUITE - " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
        ws.Cells(2, 1).value = "PROSLO: " & mPass & "   PALO: " & mFail & _
                               IIf(clean, "", "   (PREKINUTO GRESKOM - vidi poslednju liniju)")
        Dim lines() As String: lines = Split(mReport, vbLf)
        Dim i As Long
        For i = 0 To UBound(lines)
            ws.Cells(4 + i, 1).value = lines(i)
        Next i
        ws.Columns(1).ColumnWidth = 110
        ws.Activate
    End If
    On Error GoTo 0

    If mFail = 0 And clean Then
        MsgBox "PROIZVODNJA TEST: svi testovi prosli (" & mPass & " provera)." & vbCrLf & _
               "Svi test podaci su vraceni (rollback). Detalji: list _TestProizvodnja.", _
               vbInformation, APP_NAME
    Else
        MsgBox "PROIZVODNJA TEST: " & mPass & " proslo, " & mFail & " PALO." & vbCrLf & vbCrLf & _
               mFails & vbCrLf & vbCrLf & _
               "Svi test podaci su vraceni (rollback). Detalji: list _TestProizvodnja.", _
               vbExclamation, APP_NAME
    End If
End Sub
