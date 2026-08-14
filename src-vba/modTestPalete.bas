Attribute VB_Name = "modTestPalete"
'Attribute VB_Name = "modTestPalete"
Option Explicit

' ============================================================
' modTestPalete -- integracioni test-suite za paleta-engine
' (storno/ispravka prijemnice, re-point, relabel, korekcija kolicina,
' detach, storno prazne palete, zbirna re-point).
'
' POKRETANJE: Alt+F8 -> RunPaleteTestSuite
'
' BEZBEDNOST (zasto sme na zivoj svesci):
'  - Ceo suite radi UNUTAR JEDNE clsTransaction: na startu se snimi snapshot
'    svih tabela koje testovi diraju, a na kraju se UVEK radi RollbackTx --
'    i kad svi testovi produ i kad neki padne. RestoreTable vraca i broj
'    redova i vrednosti, pa test ne ostavlja NIKAKVE podatke.
'  - Unutrasnje _TX funkcije engine-a normalno komituju svoje transakcije;
'    spoljni rollback ih sve ponisti (snapshot-restore semantika).
'  - Excel OnTime autosave ne moze da se okine DOK makro radi, pa se na disk
'    nikad ne snima medju-stanje; posle rollback-a snima se original.
'  - Test podaci koriste izmisljen identitet (vrsta "TST-VOCE", tip ambalaze
'    "TST-AMB", brojevi "TST-*"), pa se ne mesaju sa stvarnim otvorenim
'    paletama (GetOrCreateOpenPaleta poklapa vrsta+sorta+klasa+tipAmb).
'  - Ako od ranije postoje TST- podaci (npr. prekinut run), suite odbija start.
'  - Monitor eventi nastali u testu (PALETA_ADJUST i sl.) ostaju u logu --
'    bezopasno; prepoznaju se po "TST-" broju prijemnice.
'  - Preporuka za PRVI put: pokreni na kopiji fajla.
'
' STA POKRIVA (engine; UI dijalozi se testiraju rucno po checklisti):
'  T01 paletizacija + spill preko kapaciteta (delegacija na SpillGajbice)
'      + idempotency guard (ista prijemnica ne moze dvaput)
'  T02 re-point CLEAN + KG-sync (iste gajbice, promenjen kg)
'  T03 RELABEL gejt (bez allowRelabel odbija; sa njim menja etiketu)
'  T04 per-klasa verdikt GAJBICA (agregat isti, klase razlicite) + Adjust obe klase
'  T05 Adjust: visak ne staje -> ADJ_NEEDS_CHOICE bez izmena, pa PRELIJ
'  T06 Adjust: PREKO kapaciteta na istoj paleti
'  T07 Adjust: smanjenje kroz vise paleta (storno stavke na 0, reopen)
'  T08 detach: su-stanar netaknut; prazna fantom paleta stornirana
'  T09 Adjust blokiran na preradjenoj paleti
'  T10 zbirna re-point: i paletne stavke dobijaju novu zbirnu
'
' Rezultat: MsgBox rezime + pun izvestaj na listu "_TestPalete" + Immediate.
' ============================================================

Private Const TST_VRSTA As String = "TST-VOCE"
Private Const TST_AMB As String = "TST-AMB"
Private Const TST_CAP As Long = 10          ' kapacitet test-palete (tblKulture seed)
Private Const TST_CRATE_KG As Double = 2#   ' tezina test-gajbice (tblTipAmbalaze seed)

Private mPass As Long
Private mFail As Long
Private mFails As String

' Gate: suite mora da PODIGNE gresku i kad provera padne i kad se uopste NE
' pokrene (paletiranje iskljuceno, zatecen TST- ostatak, operater odustao).
' Tih Exit Sub bi runner video kao OK -- "nije pokrenuto" nije "proslo".
' Konvencija: modTestBanka.ERR_BIT_SUITE_FAILED.
Private Const ERR_PALETE_SUITE_FAILED As Long = vbObjectError + 2964
Private mReport As String

Public Sub RunPaleteTestSuite()
    Dim tx As clsTransaction
    On Error GoTo EH

    If Not IsPaletiranjeEnabled() Then
        MsgBox "Paletiranje je iskljuceno u Podesavanjima - test-suite radi na paleta-engine-u" & vbCrLf & _
               "i zahteva ukljuceno paletiranje. Ukljuci pa pokreni ponovo.", vbExclamation, APP_NAME
        Err.Raise ERR_PALETE_SUITE_FAILED, "modTestPalete.RunPaleteTestSuite", _
            "suite NIJE pokrenut: paletiranje je iskljuceno u Podesavanjima"
    End If
    SetPaletizeSkip False

    ' Ostaci prethodnog (prekinutog) run-a? Ne diraj nista, uputi operatera.
    If FindRows(TBL_KULTURE, "VrstaVoca", TST_VRSTA).count > 0 _
       Or FindRows(TBL_PRIJEMNICA, COL_PRJ_BROJ, "TST-P1").count > 0 Then
        MsgBox "Nadjeni su TST- test podaci od ranije (verovatno prekinut test-run)." & vbCrLf & _
               "Zatvori svesku BEZ snimanja i otvori ponovo, pa pokreni test iz cistog stanja.", _
               vbExclamation, APP_NAME
        Err.Raise ERR_PALETE_SUITE_FAILED, "modTestPalete.RunPaleteTestSuite", _
            "suite NIJE pokrenut: zatecen TST- ostatak od prekinutog run-a"
    End If

    If MsgBox("Pokrece PALETE test-suite na OVOJ radnoj svesci." & vbCrLf & vbCrLf & _
              "Sve izmene se rade u jednoj transakciji i UVEK se ponistavaju" & vbCrLf & _
              "(i na uspehu i na gresci) - podaci ostaju netaknuti." & vbCrLf & _
              "Preporuka za prvi put: pokreni na KOPIJI fajla." & vbCrLf & vbCrLf & _
              "Nastaviti?", vbYesNo + vbQuestion, APP_NAME) <> vbYes Then
        Err.Raise ERR_PALETE_SUITE_FAILED, "modTestPalete.RunPaleteTestSuite", _
            "suite NIJE pokrenut: operater odustao na potvrdi"
    End If

    mPass = 0: mFail = 0: mFails = "": mReport = ""

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_PRIJEMNICA
    tx.AddTableSnapshot TBL_PALETA
    tx.AddTableSnapshot TBL_PALETA_STAVKA
    tx.AddTableSnapshot TBL_KULTURE
    tx.AddTableSnapshot TBL_TIP_AMBALAZE
    tx.AddTableSnapshot TBL_ZBIRNA

    SeedMasterData

    T01_PaletizacijaISpill
    T02_ReassignCleanKgSync
    T03_RelabelGate
    T04_PerKlasaVerdiktIAdjust
    T05_NeedsChoicePaPrelij
    T06_PrekoKapaciteta
    T07_SmanjenjeKrozVisePaleta
    T08_DetachSuStanarFantom
    T09_AdjustBlokPreradjena
    T10_ZbirnaRepointStavke
    T11_RelabelBlokNaDeljenojPaleti

    tx.RollbackTx                ' UVEK vrati sve - test ne ostavlja podatke
    Set tx = Nothing
    On Error GoTo 0
    ReportResults True

    If mFail > 0 Then
        Err.Raise ERR_PALETE_SUITE_FAILED, "modTestPalete.RunPaleteTestSuite", _
            "RunPaleteTestSuite: " & CStr(mFail) & " provera palo (PASS=" & _
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

    Err.Raise ERR_PALETE_SUITE_FAILED, "modTestPalete.RunPaleteTestSuite", _
        "RunPaleteTestSuite prekinut: " & eDesc
End Sub

' ============================================================
' TESTOVI
' ============================================================

' Paletizacija 25 gajbica na kapacitet 10 -> 3 palete (10+10+5), prve dve
' zatvorene, treca otvorena; neto/amb proporcionalno; header = suma; bruto
' konzistentan; druga paletizacija iste prijemnice odbijena (guard).
Private Sub T01_PaletizacijaISpill()
    Const S As String = "T01 paletizacija+spill: "
    MakePrij "TSTPRJ-01", "TST-P1", "TST-Z1", "I", 250, 25, "TST-S01"
    Paletize "TSTPRJ-01", "TST-P1", "TST-Z1", "I", 250, 25, "TST-S01"

    Dim st As Variant: st = StInfo("TST-P1")
    ChkEq StCnt("TST-P1"), 3, S & "3 stavke"
    If StCnt("TST-P1") = 3 Then
        ChkEq st(1, 3), 10, S & "stavka1 = 10 gajb"
        ChkEq st(2, 3), 10, S & "stavka2 = 10 gajb"
        ChkEq st(3, 3), 5, S & "stavka3 = 5 gajb"
        ChkEqD CDbl(st(1, 4)), 100, S & "stavka1 neto 100"
        ChkEqD CDbl(st(3, 4)), 50, S & "stavka3 neto 50"
        ChkEqD CDbl(st(1, 5)), 20, S & "stavka1 amb 20 (10 x 2kg)"
        Chk CStr(st(1, 2)) <> CStr(st(2, 2)) And CStr(st(2, 2)) <> CStr(st(3, 2)), S & "3 razlicite palete"

        Dim p1 As Variant: p1 = PalInfo(CStr(st(1, 2)))
        Dim p3 As Variant: p3 = PalInfo(CStr(st(3, 2)))
        ChkEq p1(1), 10, S & "paleta1 header 10 gajb"
        ChkEq CStr(p1(5)), PAL_STATUS_ZATVORENA, S & "paleta1 zatvorena"
        ChkEq p3(1), 5, S & "paleta3 header 5 gajb"
        ChkEq CStr(p3(5)), PAL_STATUS_OTVORENA, S & "paleta3 otvorena"
        ChkEqD CDbl(p1(4)), 100 + 20 + CDbl(p1(9)), S & "paleta1 bruto = neto+amb+palKg"
    End If

    ' Idempotency guard: ponovna paletizacija iste prijemnice mora da padne.
    Dim hadErr As Boolean
    On Error Resume Next
    Err.Clear
    Paletize "TSTPRJ-01", "TST-P1", "TST-Z1", "I", 250, 25, "TST-S01"
    hadErr = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    Chk hadErr, S & "idempotency guard (druga paletizacija odbijena)"
End Sub

' Ispravka bez promene gajbica, promenjen kg: verdikt CLEAN, re-point + KG-sync
' (stavka i header palete dobijaju nov neto), stara prijemnica ostaje bez stavki.
Private Sub T02_ReassignCleanKgSync()
    Const S As String = "T02 clean re-point + kg-sync: "
    MakePrij "TSTPRJ-02", "TST-P2A", "TST-Z2", "I", 80, 8, "TST-S02"
    Paletize "TSTPRJ-02", "TST-P2A", "TST-Z2", "I", 80, 8, "TST-S02"
    MarkStornoPrij "TST-P2A"
    MakePrij "TSTPRJ-03", "TST-P2B", "TST-Z2", "I", 96, 8, "TST-S02"   ' ispravka: kg 80 -> 96

    Dim vv As Variant: vv = EvaluatePaletaReassign("TST-P2A", "TST-P2B")
    ChkEq CStr(vv(0)), "CLEAN", S & "verdikt CLEAN"

    Dim w As String, gd As Boolean
    Chk ReassignPaleteToPrijemnica_TX("TST-P2A", "TST-P2B", w, False, gd), S & "reassign uspeo"
    Chk Not gd, S & "gajbDiff = False"
    ChkEq StCnt("TST-P2A"), 0, S & "stara bez aktivnih stavki"
    ChkEq StCnt("TST-P2B"), 1, S & "nova ima 1 stavku"
    Dim st As Variant: st = StInfo("TST-P2B")
    If StCnt("TST-P2B") = 1 Then
        ChkEq st(1, 3), 8, S & "gajbice ostale 8"
        ChkEqD CDbl(st(1, 4)), 96, S & "KG-sync: stavka neto 96"
        ChkEq CStr(st(1, 7)), "TST-Z2", S & "zbirna preneta"
        Dim pi As Variant: pi = PalInfo(CStr(st(1, 2)))
        ChkEqD CDbl(pi(2)), 96, S & "header neto 96"
    End If
End Sub

' Ispravka sa drugom sortom: gejt odbija bez allowRelabel (nista ne menja);
' sa allowRelabel prevezuje + prepravlja etiketu stavke i palete + Istorija.
Private Sub T03_RelabelGate()
    Const S As String = "T03 relabel gejt: "
    MakePrij "TSTPRJ-04", "TST-P3A", "TST-Z3", "I", 50, 5, "TST-S03"
    Paletize "TSTPRJ-04", "TST-P3A", "TST-Z3", "I", 50, 5, "TST-S03"
    MarkStornoPrij "TST-P3A"
    MakePrij "TSTPRJ-05", "TST-P3B", "TST-Z3", "I", 50, 5, "TST-S03B"  ' ispravka: sorta

    Dim vv As Variant: vv = EvaluatePaletaReassign("TST-P3A", "TST-P3B")
    ChkEq CStr(vv(0)), "RELABEL", S & "verdikt RELABEL"

    Dim w As String, gd As Boolean
    Chk Not ReassignPaleteToPrijemnica_TX("TST-P3A", "TST-P3B", w, False, gd), _
        S & "bez allowRelabel odbijeno"
    Chk InStr(w, "identitet") > 0, S & "warn objasnjava identitet"
    ChkEq StCnt("TST-P3A"), 1, S & "nista nije promenjeno posle odbijanja"

    Chk ReassignPaleteToPrijemnica_TX("TST-P3A", "TST-P3B", w, True, gd), _
        S & "sa allowRelabel uspeo"
    Dim st As Variant: st = StInfo("TST-P3B")
    If StCnt("TST-P3B") = 1 Then
        ChkEq CStr(st(1, 8)), "TST-S03B", S & "stavka nosi novu sortu"
        Dim pi As Variant: pi = PalInfo(CStr(st(1, 2)))
        ChkEq CStr(pi(7)), "TST-S03B", S & "paleta nosi novu sortu"
        Chk InStr(CStr(pi(8)), "RELABEL") > 0, S & "Istorija sadrzi RELABEL"
    End If
End Sub

' Dual-klasa regresija: stara I=6/II=4, nova I=4/II=6 (agregat 10=10!) mora
' da bude GAJBICA (per klasa), a Adjust koriguje OBE klase u mestu + Istorija
' nosi deltu po paleti (-2 na I, +2 na II).
Private Sub T04_PerKlasaVerdiktIAdjust()
    Const S As String = "T04 per-klasa verdikt+adjust: "
    MakePrij "TSTPRJ-06", "TST-P4", "TST-Z4", "I", 60, 6, "TST-S04"
    Paletize "TSTPRJ-06", "TST-P4", "TST-Z4", "I", 60, 6, "TST-S04"
    MakePrij "TSTPRJ-07", "TST-P4", "TST-Z4", "II", 40, 4, "TST-S04"
    Paletize "TSTPRJ-07", "TST-P4", "TST-Z4", "II", 40, 4, "TST-S04"
    MarkStornoPrij "TST-P4"
    MakePrij "TSTPRJ-08", "TST-P5", "TST-Z4", "I", 40, 4, "TST-S04"
    MakePrij "TSTPRJ-09", "TST-P5", "TST-Z4", "II", 60, 6, "TST-S04"

    Dim vv As Variant: vv = EvaluatePaletaReassign("TST-P4", "TST-P5")
    ChkEq CStr(vv(0)), "GAJBICA", S & "verdikt GAJBICA (ne CLEAN na isti agregat)"

    Dim w As String, gd As Boolean
    Chk ReassignPaleteToPrijemnica_TX("TST-P4", "TST-P5", w, False, gd), S & "reassign uspeo"
    Chk gd, S & "gajbDiff = True"

    Dim info As String
    ChkEq AdjustPaletaGajbiceZaPrijemnicu_TX("TST-P5", "", info), 2, S & "korigovane 2 klase"

    Dim st As Variant: st = StInfo("TST-P5")
    Dim gI As Long, gII As Long, palI As String, palII As String
    Dim i As Long
    For i = 1 To StCnt("TST-P5")
        If CStr(st(i, 6)) = "II" Then
            gII = gII + CLng(st(i, 3)): palII = CStr(st(i, 2))
        Else
            gI = gI + CLng(st(i, 3)): palI = CStr(st(i, 2))
        End If
    Next i
    ChkEq gI, 4, S & "klasa I = 4 gajbice"
    ChkEq gII, 6, S & "klasa II = 6 gajbica"
    Chk InStr(CStr(PalInfo(palI)(8)), "gajb=-2") > 0, S & "Istorija I: gajb=-2"
    Chk InStr(CStr(PalInfo(palII)(8)), "gajb=+2") > 0, S & "Istorija II: gajb=+2"
End Sub

' Visak koji ne staje: Adjust("") vraca ADJ_NEEDS_CHOICE i NE menja nista;
' PRELIJ dopuni paletu do kapaciteta (zatvori je) + visak na novu paletu,
' neto se re-sinhronizuje proporcionalno.
Private Sub T05_NeedsChoicePaPrelij()
    Const S As String = "T05 needs-choice + prelij: "
    MakePrij "TSTPRJ-10", "TST-P6A", "TST-Z5", "I", 80, 8, "TST-S05"
    Paletize "TSTPRJ-10", "TST-P6A", "TST-Z5", "I", 80, 8, "TST-S05"
    MarkStornoPrij "TST-P6A"
    MakePrij "TSTPRJ-11", "TST-P6B", "TST-Z5", "I", 150, 15, "TST-S05"  ' 8 -> 15

    Dim w As String, gd As Boolean, info As String
    Chk ReassignPaleteToPrijemnica_TX("TST-P6A", "TST-P6B", w, False, gd), S & "reassign uspeo"
    Chk gd, S & "gajbDiff = True"

    ChkEq AdjustPaletaGajbiceZaPrijemnicu_TX("TST-P6B", "", info), ADJ_NEEDS_CHOICE, _
          S & "bez moda -> ADJ_NEEDS_CHOICE"
    Dim st As Variant: st = StInfo("TST-P6B")
    ChkEq st(1, 3), 8, S & "needs-choice NIJE nista promenio (gajb 8)"
    ChkEqD CDbl(st(1, 4)), 80, S & "needs-choice NIJE nista promenio (neto 80)"

    Chk AdjustPaletaGajbiceZaPrijemnicu_TX("TST-P6B", "PRELIJ", info) >= 1, S & "PRELIJ uspeo"
    st = StInfo("TST-P6B")
    ChkEq StCnt("TST-P6B"), 2, S & "2 stavke posle preliva"
    If StCnt("TST-P6B") = 2 Then
        ChkEq st(1, 3), 10, S & "prva dopunjena na 10"
        ChkEq st(2, 3), 5, S & "prelivena stavka 5"
        ChkEqD CDbl(st(1, 4)), 100, S & "neto resync prve (100)"
        ChkEqD CDbl(st(2, 4)), 50, S & "neto prelivene (50)"
        ChkEq CStr(PalInfo(CStr(st(1, 2)))(5)), PAL_STATUS_ZATVORENA, S & "puna paleta zatvorena"
        ChkEq CStr(PalInfo(CStr(st(2, 2)))(5)), PAL_STATUS_OTVORENA, S & "nova paleta otvorena"
    End If
End Sub

' PREKO kapaciteta: sve ostaje na istoj paleti (15/10), paleta se zatvara,
' Istorija nosi +7.
Private Sub T06_PrekoKapaciteta()
    Const S As String = "T06 preko kapaciteta: "
    MakePrij "TSTPRJ-12", "TST-P8A", "TST-Z6", "I", 80, 8, "TST-S06"
    Paletize "TSTPRJ-12", "TST-P8A", "TST-Z6", "I", 80, 8, "TST-S06"
    MarkStornoPrij "TST-P8A"
    MakePrij "TSTPRJ-13", "TST-P8B", "TST-Z6", "I", 150, 15, "TST-S06"

    Dim w As String, gd As Boolean, info As String
    Chk ReassignPaleteToPrijemnica_TX("TST-P8A", "TST-P8B", w, False, gd), S & "reassign uspeo"
    Chk AdjustPaletaGajbiceZaPrijemnicu_TX("TST-P8B", "PREKO", info) >= 1, S & "PREKO uspeo"

    ChkEq StCnt("TST-P8B"), 1, S & "jedna stavka (bez preliva)"
    Dim st As Variant: st = StInfo("TST-P8B")
    If StCnt("TST-P8B") = 1 Then
        ChkEq st(1, 3), 15, S & "stavka 15 gajbica"
        Dim pi As Variant: pi = PalInfo(CStr(st(1, 2)))
        ChkEq pi(1), 15, S & "header 15 (preko kapaciteta 10)"
        ChkEq CStr(pi(5)), PAL_STATUS_ZATVORENA, S & "paleta zatvorena"
        ChkEqD CDbl(pi(2)), 150, S & "header neto 150"
        Chk InStr(CStr(pi(8)), "gajb=+7") > 0, S & "Istorija: gajb=+7"
    End If
End Sub

' Smanjenje kroz vise paleta: 25 (10/10/5) -> 12; skida se od poslednje stavke,
' stavka na 0 se stornira, zatvorena paleta ispod kapaciteta se ponovo otvara.
Private Sub T07_SmanjenjeKrozVisePaleta()
    Const S As String = "T07 smanjenje kroz palete: "
    MakePrij "TSTPRJ-14", "TST-P9A", "TST-Z7", "I", 250, 25, "TST-S07"
    Paletize "TSTPRJ-14", "TST-P9A", "TST-Z7", "I", 250, 25, "TST-S07"
    MarkStornoPrij "TST-P9A"
    MakePrij "TSTPRJ-15", "TST-P9B", "TST-Z7", "I", 120, 12, "TST-S07"  ' 25 -> 12

    Dim w As String, gd As Boolean, info As String
    Chk ReassignPaleteToPrijemnica_TX("TST-P9A", "TST-P9B", w, False, gd), S & "reassign uspeo"
    Chk AdjustPaletaGajbiceZaPrijemnicu_TX("TST-P9B", "", info) >= 1, S & "smanjenje bez pitanja"

    ChkEq StCnt("TST-P9B"), 2, S & "ostale 2 aktivne stavke (treca stornirana)"
    ChkEq StSumG("TST-P9B"), 12, S & "ukupno 12 gajbica"
    Dim st As Variant: st = StInfo("TST-P9B")
    If StCnt("TST-P9B") = 2 Then
        ChkEq st(1, 3), 10, S & "prva stavka 10"
        ChkEq st(2, 3), 2, S & "druga stavka 2"
        ChkEqD CDbl(st(1, 4)), 100, S & "neto resync prve (100)"
        ChkEqD CDbl(st(2, 4)), 20, S & "neto druge (20)"
        Dim p2 As Variant: p2 = PalInfo(CStr(st(2, 2)))
        ChkEq CStr(p2(5)), PAL_STATUS_OTVORENA, S & "druga paleta reopen (2/10)"
        ChkEq CStr(PalInfo(CStr(st(1, 2)))(5)), PAL_STATUS_ZATVORENA, S & "prva ostala zatvorena"
    End If
End Sub

' Detach: dve prijemnice na ISTOJ paleti (su-stanari). Skidanje prve ne dira
' drugu ni paletu; skidanje i druge prazni paletu -> paleta se stornira
' (kanonski StornoPaleta) + Istorija STORNO_PRAZNA.
Private Sub T08_DetachSuStanarFantom()
    Const S As String = "T08 detach su-stanar+fantom: "
    MakePrij "TSTPRJ-16", "TST-P10", "TST-Z8", "I", 40, 4, "TST-S08"
    Paletize "TSTPRJ-16", "TST-P10", "TST-Z8", "I", 40, 4, "TST-S08"
    MakePrij "TSTPRJ-17", "TST-P11", "TST-Z8B", "I", 30, 3, "TST-S08"
    Paletize "TSTPRJ-17", "TST-P11", "TST-Z8B", "I", 30, 3, "TST-S08"

    Dim stA As Variant: stA = StInfo("TST-P10")
    Dim stB As Variant: stB = StInfo("TST-P11")
    Dim palID As String: palID = CStr(stA(1, 2))
    ChkEq CStr(stB(1, 2)), palID, S & "obe prijemnice na istoj paleti"

    MarkStornoPrij "TST-P10"
    Dim info As String
    ChkEq DetachOsirocenePaletaStavke_TX("TST-P10", info), 1, S & "skinuta 1 stavka"
    ChkEq StCnt("TST-P10"), 0, S & "P10 bez aktivnih stavki"
    ChkEq StCnt("TST-P11"), 1, S & "su-stanar P11 netaknut"
    Dim pi As Variant: pi = PalInfo(palID)
    ChkEq CStr(pi(6)), "", S & "paleta jos aktivna (nosi P11)"
    ChkEq pi(1), 3, S & "header 3 gajbice (samo P11)"

    MarkStornoPrij "TST-P11"
    ChkEq DetachOsirocenePaletaStavke_TX("TST-P11", info), 1, S & "skinuta i P11 stavka"
    pi = PalInfo(palID)
    ChkEq UCase$(CStr(pi(6))), "DA", S & "prazna fantom paleta stornirana"
    Chk InStr(CStr(pi(8)), "STORNO_PRAZNA") > 0, S & "Istorija sadrzi STORNO_PRAZNA"
    Chk InStr(info, "stornirane") > 0, S & "outInfo javlja storniranu paletu"
End Sub

' Preradjena paleta: korekcija kolicina je blokirana (ADJ_BLOCKED), bez izmena.
Private Sub T09_AdjustBlokPreradjena()
    Const S As String = "T09 preradjena blokada: "
    MakePrij "TSTPRJ-18", "TST-P12", "TST-Z9", "I", 50, 5, "TST-S09"
    Paletize "TSTPRJ-18", "TST-P12", "TST-Z9", "I", 50, 5, "TST-S09"

    Dim st As Variant: st = StInfo("TST-P12")
    Dim palRow As Long: palRow = FindRows(TBL_PALETA, COL_PAL_ID, CStr(st(1, 2))).item(1)
    RequireUpdateCell TBL_PALETA, palRow, COL_PAL_PRERADJENO, "Da", "modTestPalete.T09"

    ' dokument menja gajbice 5 -> 7 (direktan upis, kao ispravka)
    Dim prjRow As Long: prjRow = FindRows(TBL_PRIJEMNICA, COL_PRJ_BROJ, "TST-P12").item(1)
    RequireUpdateCell TBL_PRIJEMNICA, prjRow, COL_PRJ_KOL_AMB, 7, "modTestPalete.T09"

    Dim info As String
    ChkEq AdjustPaletaGajbiceZaPrijemnicu_TX("TST-P12", "", info), ADJ_BLOCKED, _
          S & "ADJ_BLOCKED na preradjenoj"
    Chk InStr(info, "preradjena") > 0, S & "poruka objasnjava razlog"
    st = StInfo("TST-P12")
    ChkEq st(1, 3), 5, S & "stavka netaknuta (5 gajbica)"
End Sub

' Zbirna re-point: prijemnica + NJENE PALETNE STAVKE dobijaju novu zbirnu
' (sledljivost gap-fix).
Private Sub T10_ZbirnaRepointStavke()
    Const S As String = "T10 zbirna re-point: "
    TstAppend TBL_ZBIRNA, Array(COL_ZBR_ID, COL_ZBR_BROJ, COL_STORNIRANO), _
              Array("TSTZBR-A", "TST-Z10A", "")
    TstAppend TBL_ZBIRNA, Array(COL_ZBR_ID, COL_ZBR_BROJ, COL_STORNIRANO), _
              Array("TSTZBR-B", "TST-Z10B", "")

    MakePrij "TSTPRJ-19", "TST-P13", "TST-Z10A", "I", 50, 5, "TST-S10"
    Paletize "TSTPRJ-19", "TST-P13", "TST-Z10A", "I", 50, 5, "TST-S10"

    Chk ReassignPrijemnicaToZbirna_TX("TST-P13", "TST-Z10B"), S & "re-point uspeo"

    Dim prjRow As Long: prjRow = FindRows(TBL_PRIJEMNICA, COL_PRJ_BROJ, "TST-P13").item(1)
    Dim d As Variant: d = GetTableData(TBL_PRIJEMNICA)
    ChkEq Trim$(CStr(d(prjRow, GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)))), _
          "TST-Z10B", S & "prijemnica nosi novu zbirnu"

    Dim st As Variant: st = StInfo("TST-P13")
    ChkEq CStr(st(1, 7)), "TST-Z10B", S & "paletna stavka nosi novu zbirnu"
End Sub

' HARD GUARD: RELABEL na paleti koju dele DVE prijemnice mora biti BLOKIRAN
' (ne warning) -- promena identiteta headera bi iskvarila robu druge prijemnice.
' Provera: reassign vraca False, outWarn sadrzi "BLOKIRANO", i NISTA se ne menja
' (stavke stare prijemnice ostaju, header palete ostaje stara sorta, su-stanar ceo).
Private Sub T11_RelabelBlokNaDeljenojPaleti()
    Const S As String = "T11 relabel blok na deljenoj paleti: "
    ' Dve prijemnice, isti identitet (TST-S11) -> dele istu paletu (4+3 = 7/10).
    MakePrij "TSTPRJ-20", "TST-P14", "TST-Z11", "I", 40, 4, "TST-S11"
    Paletize "TSTPRJ-20", "TST-P14", "TST-Z11", "I", 40, 4, "TST-S11"
    MakePrij "TSTPRJ-21", "TST-P15", "TST-Z11B", "I", 30, 3, "TST-S11"
    Paletize "TSTPRJ-21", "TST-P15", "TST-Z11B", "I", 30, 3, "TST-S11"

    Dim stA As Variant: stA = StInfo("TST-P14")
    Dim palID As String: palID = CStr(stA(1, 2))
    ChkEq CStr(StInfo("TST-P15")(1, 2)), palID, S & "obe prijemnice na istoj paleti"

    ' Storno P14, ispravka sa DRUGOM sortom (RELABEL) -> mora blokada zbog su-stanara.
    MarkStornoPrij "TST-P14"
    MakePrij "TSTPRJ-22", "TST-P16", "TST-Z11", "I", 40, 4, "TST-S11X"  ' druga sorta

    Dim vv As Variant: vv = EvaluatePaletaReassign("TST-P14", "TST-P16")
    ChkEq CStr(vv(0)), "RELABEL", S & "verdikt RELABEL"

    Dim w As String, gd As Boolean
    Chk Not ReassignPaleteToPrijemnica_TX("TST-P14", "TST-P16", w, True, gd), _
        S & "BLOKIRANO cak i uz allowRelabel=True"
    Chk InStr(w, "BLOKIRANO") > 0, S & "outWarn objasnjava blokadu"

    ' Nista se nije promenilo:
    ChkEq StCnt("TST-P14"), 1, S & "stavka stare prijemnice ostala (nije prevezana)"
    ChkEq StCnt("TST-P16"), 0, S & "nova prijemnica bez stavki"
    ChkEq StCnt("TST-P15"), 1, S & "su-stanar netaknut"
    Dim pi As Variant: pi = PalInfo(palID)
    ChkEq CStr(pi(7)), "TST-S11", S & "header palete ostao stara sorta (nije iskvaren)"
    ChkEq pi(1), 7, S & "header gajbice ostao 7 (4+3)"

    ' Bezbedan izlaz postoji: skini stavke stare pa nov unos (detach ne dira su-stanara).
    Dim info As String
    ChkEq DetachOsirocenePaletaStavke_TX("TST-P14", info), 1, S & "detach kao bezbedan izlaz radi"
    ChkEq StCnt("TST-P15"), 1, S & "su-stanar i posle detach-a netaknut"
End Sub

' ============================================================
' SEED + DATA HELPERI (privatni za test; TST- identitet, snapshot-ovane tabele)
' ============================================================

Private Sub SeedMasterData()
    TstAppend TBL_KULTURE, _
        Array("KulturaID", "VrstaVoca", "SortaVoca", COL_KUL_GAJBICA_PALETA), _
        Array("TST-KUL", TST_VRSTA, "TST", TST_CAP)
    TstAppend TBL_TIP_AMBALAZE, _
        Array(COL_TAMB_TIP, COL_TAMB_TEZINA), _
        Array(TST_AMB, TST_CRATE_KG)
End Sub

' Append po IMENU kolone (mirror PalAppendRow obrasca; kolone koje ne postoje
' se preskacu -> bezbedno pod schema drift-om). Pada glasno ako append ne uspe.
Private Sub TstAppend(ByVal tblName As String, ByVal cols As Variant, ByVal vals As Variant)
    Dim lo As ListObject
    Set lo = GetTable(tblName)
    If lo Is Nothing Then Err.Raise vbObjectError + 9910, "modTestPalete.TstAppend", _
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
        Err.Raise vbObjectError + 9911, "modTestPalete.TstAppend", _
                  "AppendRow nije uspeo: " & tblName
    End If
End Sub

Private Function MakePrij(ByVal id As String, ByVal broj As String, ByVal zbirna As String, _
                          ByVal klasa As String, ByVal kolKg As Double, ByVal gajb As Long, _
                          ByVal sorta As String) As String
    TstAppend TBL_PRIJEMNICA, _
        Array(COL_PRJ_ID, COL_PRJ_BROJ, COL_PRJ_DATUM, COL_PRJ_KLASA, COL_PRJ_KOLICINA, _
              COL_PRJ_KOL_AMB, COL_PRJ_BROJ_ZBIRNE, COL_PRJ_VRSTA, COL_PRJ_SORTA, _
              COL_PRJ_TIP_AMB, COL_PRJ_CENA, COL_STORNIRANO), _
        Array(id, broj, Date, klasa, kolKg, _
              gajb, zbirna, TST_VRSTA, sorta, _
              TST_AMB, 0, "")
    MakePrij = id
End Function

Private Sub Paletize(ByVal id As String, ByVal broj As String, ByVal zbirna As String, _
                     ByVal klasa As String, ByVal kolKg As Double, ByVal gajb As Long, _
                     ByVal sorta As String)
    PaletizePrijemnica prijemnicaID:=id, brojPrij:=broj, brojZbirne:=zbirna, _
                       vrstaVoca:=TST_VRSTA, sortaVoca:=sorta, klasa:=klasa, _
                       netoKg:=kolKg, brGajbica:=gajb, tipAmb:=TST_AMB
End Sub

' Storniraj SVE redove prijemnice datog broja (test-ekvivalent storna dokumenta;
' engine cita samo Stornirano flag).
Private Sub MarkStornoPrij(ByVal broj As String)
    Dim rows As Collection: Set rows = FindRows(TBL_PRIJEMNICA, COL_PRJ_BROJ, broj)
    Dim k As Long
    For k = 1 To rows.count
        RequireUpdateCell TBL_PRIJEMNICA, rows(k), COL_STORNIRANO, "Da", "modTestPalete.MarkStornoPrij"
    Next k
End Sub

' Aktivne paleta-stavke broja, hronoloski. 1..n x kolone:
' 1=rowIdx, 2=PaletaID, 3=gajb, 4=neto, 5=amb, 6=klasa, 7=zbirna, 8=sorta.
' Empty ako nema.
Private Function StInfo(ByVal broj As String) As Variant
    Dim s As Variant: s = GetTableData(TBL_PALETA_STAVKA)
    If IsEmpty(s) Then Exit Function
    Dim cBr As Long, cPal As Long, cGa As Long, cNe As Long, cAm As Long
    Dim cKl As Long, cZb As Long, cSo As Long, cSt As Long
    cBr = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_PRIJ)
    cPal = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PALETA_ID)
    cGa = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BR_GAJBICA)
    cNe = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_NETO)
    cAm = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_AMBALAZA)
    cKl = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_KLASA)
    cZb = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_ZBIRNE)
    cSo = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_SORTA)
    cSt = GetColumnIndex(TBL_PALETA_STAVKA, COL_STORNIRANO)

    Dim hits As Collection: Set hits = New Collection
    Dim r As Long
    For r = 1 To UBound(s, 1)
        If Trim$(CStr(SafeCell(s, r, cBr))) = broj _
           And UCase$(Trim$(CStr(SafeCell(s, r, cSt)))) <> "DA" Then hits.Add r
    Next r
    If hits.count = 0 Then Exit Function

    Dim res As Variant: ReDim res(1 To hits.count, 1 To 8)
    Dim k As Long
    For k = 1 To hits.count
        r = hits(k)
        res(k, 1) = r
        res(k, 2) = CStr(SafeCell(s, r, cPal))
        res(k, 3) = NzL(SafeCell(s, r, cGa))
        res(k, 4) = NzD(SafeCell(s, r, cNe))
        res(k, 5) = NzD(SafeCell(s, r, cAm))
        res(k, 6) = Trim$(CStr(SafeCell(s, r, cKl)))
        res(k, 7) = Trim$(CStr(SafeCell(s, r, cZb)))
        res(k, 8) = Trim$(CStr(SafeCell(s, r, cSo)))
    Next k
    StInfo = res
End Function

Private Function StCnt(ByVal broj As String) As Long
    Dim st As Variant: st = StInfo(broj)
    If IsEmpty(st) Then StCnt = 0 Else StCnt = UBound(st, 1)
End Function

Private Function StSumG(ByVal broj As String) As Long
    Dim st As Variant: st = StInfo(broj)
    If IsEmpty(st) Then Exit Function
    Dim k As Long, tot As Long
    For k = 1 To UBound(st, 1)
        tot = tot + CLng(st(k, 3))
    Next k
    StSumG = tot
End Function

' Header palete: 1=gajb, 2=neto, 3=amb, 4=bruto, 5=status, 6=stornirano,
' 7=sorta, 8=istorija, 9=palKg.
Private Function PalInfo(ByVal palID As String) As Variant
    Dim rows As Collection: Set rows = FindRows(TBL_PALETA, COL_PAL_ID, palID)
    Dim res(1 To 9) As Variant
    If rows.count = 0 Then PalInfo = res: Exit Function
    Dim r As Long: r = rows.item(1)
    Dim d As Variant: d = GetTableData(TBL_PALETA)
    res(1) = NzL(SafeCell(d, r, GetColumnIndex(TBL_PALETA, COL_PAL_BR_GAJBICA)))
    res(2) = NzD(SafeCell(d, r, GetColumnIndex(TBL_PALETA, COL_PAL_NETO)))
    res(3) = NzD(SafeCell(d, r, GetColumnIndex(TBL_PALETA, COL_PAL_AMBALAZA)))
    res(4) = NzD(SafeCell(d, r, GetColumnIndex(TBL_PALETA, COL_PAL_BRUTO)))
    res(5) = Trim$(CStr(SafeCell(d, r, GetColumnIndex(TBL_PALETA, COL_PAL_STATUS))))
    res(6) = Trim$(CStr(SafeCell(d, r, GetColumnIndex(TBL_PALETA, COL_STORNIRANO))))
    res(7) = Trim$(CStr(SafeCell(d, r, GetColumnIndex(TBL_PALETA, COL_PAL_SORTA))))
    res(8) = CStr(SafeCell(d, r, GetColumnIndex(TBL_PALETA, COL_PAL_ISTORIJA)))
    res(9) = NzD(SafeCell(d, r, GetColumnIndex(TBL_PALETA, COL_PAL_PALETA_KG)))
    PalInfo = res
End Function

' ============================================================
' ASSERT + IZVESTAJ
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
    ' Broj provera runneru (min_asserts u tests/suite_manifest.json).
    TR_Report "RunPaleteTestSuite", mPass, mFail

    On Error Resume Next
    Debug.Print "===== PALETE TEST SUITE ====="
    Debug.Print mReport
    Debug.Print "PROSLO: " & mPass & "   PALO: " & mFail

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("_TestPalete")
    If ws Is Nothing Then
        Err.Clear
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = "_TestPalete"
    End If
    If Not ws Is Nothing Then
        ws.Cells.Clear
        ws.Cells(1, 1).value = "PALETE TEST SUITE - " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
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
        MsgBox "PALETE TEST: svi testovi prosli (" & mPass & " provera)." & vbCrLf & _
               "Svi test podaci su vraceni (rollback). Detalji: list _TestPalete.", _
               vbInformation, APP_NAME
    Else
        MsgBox "PALETE TEST: " & mPass & " proslo, " & mFail & " PALO." & vbCrLf & vbCrLf & _
               mFails & vbCrLf & vbCrLf & _
               "Svi test podaci su vraceni (rollback). Detalji: list _TestPalete.", _
               vbExclamation, APP_NAME
    End If
End Sub
