Attribute VB_Name = "modDoplata"
Option Explicit

' ============================================================
' modDoplata - DOPLATA KOOPERANTU (knjizno odobrenje, KO)
'
' Domen i odluke: docs/DOMEN/DOPLATA.md (DOP-ODL-*, DOP-*).
'
' Globalni unos: stopa u din/kg po KLJUCU CENOVNIKA (VrstaVoca + SortaVoca +
' Klasa) za period otkupa. Obracun sabere kg AKTIVNIH otkupa svakog kooperanta
' po tom kljucu i knjizi JEDAN KO PO KOOPERANTU, svi pod istim ObracunID-em, u
' JEDNOJ transakciji. Iznos je ono sto se isplacuje -- bez PDV nadoknade.
'
' PRIKAZ I PISAC DELE JEDAN PLAN (obrazac BimAutoPlan): DoplataPlan racuna
' stavke bez ijednog upisa; ekran ga prikazuje, KnjiziDoplatu_TX ga izvrsava.
' Druga kopija racuna bi se razisla prvom izmenom.
'
' Sta doplata NE dira (DOP-01): robu -- kg doplate ne ulazi ni u jedan robni
' zbir --, cenu na stavkama otkupa, otkupni list, hladnjacki lanac. Knjizen KO
' je SNIMAK (DOP-ODL-07): kasniji storno ili ispravka otkupa ga ne menja.
'
' Storno zaglavlja je u modStorno (StornoDoplata_TX), kao za ostale dokumente.
' Stavka nema svoj storno (BEZ_STORNA): aktivna je dok je aktivno zaglavlje.
' ============================================================

' Greske pisca (opseg vbObjectError + 6300..6399; potvrda je poslednji broj
' opsega, da je nijedna kapija ne bi slucajno ponela).
'
' POTVRDA PONAVLJANJA NIJE KVAR NEGO PITANJE POZIVAOCU (DOP-ODL-06, obrazac
' AMB_ERR_POTVRDA_DEFICITA): pisac vraca "" bez loga i monitora, uz tekst koji
' imenuje raniju doplatu; pozivalac pita operatera i ponavlja poziv sa potvrdom.
Public Const DOP_ERR_POTVRDA_PONAVLJANJA As Long = vbObjectError + 6399
Private Const DOP_ERR_BASE As Long = vbObjectError + 6300

' Stopa je Scripting.Dictionary sa TACNO ovim kljucevima (zatvoren spisak, kao
' stavka otkupa): tipfeler u imenu kljuca ne sme da prodje kao "nije uneto".
Private Const DOP_STOPA_KLJUCEVI As String = "|VRSTAVOCA|SORTAVOCA|KLASA|CENA|"

' Minuta u danu: granica perioda se poredi na minut, isto koliko otkupni list
' stampa (hh:nn).
Private Const MIN_U_DANU As Long = 1440

' Kolone plana (DoplataPlan). Nove kolone idu NA KRAJ: citaoci indeksiraju
' poziciono.
Private Const PL_KOOP As Long = 1
Private Const PL_STANICA As Long = 2
Private Const PL_VRSTA As Long = 3
Private Const PL_SORTA As Long = 4
Private Const PL_KLASA As Long = 5
Private Const PL_KOLICINA As Long = 6
Private Const PL_CENA As Long = 7
Private Const PL_IZNOS As Long = 8
Private Const PL_BROJ_OTKUPA As Long = 9
Private Const PL_KLJUC As Long = 10
Private Const PL_KOLONA As Long = 10

' ============================================================
' PLAN -- bez ijednog upisa
' ============================================================

' Stavke obracuna za period i stope. 2D niz (1..n, 1..10):
'   1 KooperantID | 2 StanicaID | 3 VrstaVoca | 4 SortaVoca | 5 Klasa
'   6 Kolicina (kg, 2 dec.) | 7 Cena (din/kg) | 8 Iznos (din, 2 dec.)
'   9 broj otkupa u stavci | 10 kljuc poretka (interni)
' poredano po kooperantu, stanici, vrsti, sorti, klasi. Empty kad nema nijedne.
'
' outBezOtkupa: stope (vrsta|sorta|klasa) za koje u periodu nema nijednog
' aktivnog otkupa. Nije greska plana -- ekran ih pokazuje uz pregled -- ali pisac
' ih odbija: tipfeler u sorti bi inace tiho dao doplatu nula za ceo kljuc.
'
' Obuhvat (DOP-ODL-03): svaki AKTIVAN otkup, na svakom otkupnom mestu,
' ukljucujuci hladnjacke blokove. Kilogrami su sa STAVKI (StavkeOtkupaRedovi,
' strog citac kanona) -- isti izvor koji citaju kartica i saldo.
Public Function DoplataPlan(ByVal periodOd As Date, ByVal periodDo As Date, _
                            ByVal stope As Collection, _
                            Optional ByRef outBezOtkupa As String) As Variant
    Const SRC As String = "modDoplata.DoplataPlan"

    DoplataPlan = Empty
    outBezOtkupa = ""

    RequirePeriod periodOd, periodDo, SRC

    Dim cene As Object, vrsteSorte As Object, prikaz As Object
    UcitajStope stope, cene, vrsteSorte, prikaz, SRC

    ' --- 1. Aktivna zaglavlja u periodu, za vrstu i sortu iz stopa -----------
    Dim d As Variant
    d = GetTableData(TBL_OTKUP)
    If Not IsArray(d) Then
        outBezOtkupa = SpisakKljuceva(prikaz, Nothing)
        Exit Function
    End If
    d = ExcludeStornirano(d, TBL_OTKUP)
    If Not IsArray(d) Then
        outBezOtkupa = SpisakKljuceva(prikaz, Nothing)
        Exit Function
    End If

    Dim cId As Long, cDat As Long, cKoop As Long, cSta As Long
    Dim cVr As Long, cSo As Long, cBr As Long, cSrc As Long, cCre As Long
    cId = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, SRC)
    cDat = RequireColumnIndex(TBL_OTKUP, COL_OTK_DATUM, SRC)
    cKoop = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT, SRC)
    cSta = RequireColumnIndex(TBL_OTKUP, COL_OTK_STANICA, SRC)
    cVr = RequireColumnIndex(TBL_OTKUP, COL_OTK_VRSTA, SRC)
    cSo = RequireColumnIndex(TBL_OTKUP, COL_OTK_SORTA, SRC)
    cBr = RequireColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK, SRC)
    cSrc = RequireColumnIndex(TBL_OTKUP, COL_OTK_SOURCE_CREATED_AT, SRC)
    cCre = RequireColumnIndex(TBL_OTKUP, COL_AUDIT_CREATED_AT, SRC)

    Dim uPeriodu As Object
    Set uPeriodu = CreateObject("Scripting.Dictionary")
    uPeriodu.CompareMode = vbTextCompare

    Dim i As Long, oid As String, vs As String
    For i = 1 To UBound(d, 1)
        vs = KljucVrstaSorta(d(i, cVr), d(i, cSo))
        If vrsteSorte.Exists(vs) Then
            oid = Trim$(NzToText(d(i, cId)))
            If OtkupUPeriodu(d(i, cDat), d(i, cSrc), d(i, cCre), periodOd, periodDo, _
                             oid, NzToText(d(i, cBr)), SRC) Then
                ' Dupli OtkupID obara StavkeOtkupaRedovi pre agregacije; ovde
                ' bi tiho pobedio poslednji red.
                uPeriodu(oid) = Array(Trim$(NzToText(d(i, cKoop))), _
                                      Trim$(NzToText(d(i, cSta))), vs)
            End If
        End If
    Next i

    ' --- 2. Stavke tih otkupa, po kljucu stope ---------------------------------
    Dim agg As Object, pokriveno As Object
    Set agg = CreateObject("Scripting.Dictionary")
    Set pokriveno = CreateObject("Scripting.Dictionary")
    pokriveno.CompareMode = vbTextCompare

    If uPeriodu.count > 0 Then
        Dim s As Variant
        s = modOtkup.StavkeOtkupaRedovi()
        If IsArray(s) Then
            Dim h As Variant, kStope As String, kAgg As String, v As Variant
            For i = 1 To UBound(s, 1)
                oid = CStr(s(i, 1))
                If uPeriodu.Exists(oid) Then
                    h = uPeriodu(oid)
                    kStope = CStr(h(2)) & "|" & UCase$(Trim$(CStr(s(i, 3))))
                    If cene.Exists(kStope) Then
                        pokriveno(kStope) = True
                        kAgg = KljucPoretka(CStr(h(0)), CStr(h(1)), kStope, CStr(s(i, 3)))
                        If agg.Exists(kAgg) Then
                            v = agg(kAgg)
                        Else
                            v = Array(CStr(h(0)), CStr(h(1)), kStope, 0#, 0&)
                        End If
                        v(3) = CDbl(v(3)) + CDbl(s(i, 4))
                        v(4) = CLng(v(4)) + 1
                        agg(kAgg) = v
                    End If
                End If
            Next i
        End If
    End If

    outBezOtkupa = SpisakKljuceva(prikaz, pokriveno)
    If agg.count = 0 Then Exit Function

    ' --- 3. Niz plana, poredan ------------------------------------------------
    Dim res() As Variant, n As Long, k As Variant, deo As Variant, kg As Double
    ReDim res(1 To agg.count, 1 To PL_KOLONA)
    For Each k In agg.keys
        v = agg(k)
        deo = prikaz(CStr(v(2)))
        n = n + 1
        kg = modNovac.ZaokruziNovac(CDbl(v(3)))   ' kg na dve decimale, kao stavka otkupa
        res(n, PL_KOOP) = v(0)
        res(n, PL_STANICA) = v(1)
        res(n, PL_VRSTA) = deo(0)
        res(n, PL_SORTA) = deo(1)
        res(n, PL_KLASA) = deo(2)
        res(n, PL_KOLICINA) = kg
        res(n, PL_CENA) = CDbl(cene(CStr(v(2))))
        res(n, PL_IZNOS) = modNovac.ZaokruziNovac(kg * CDbl(cene(CStr(v(2)))))
        res(n, PL_BROJ_OTKUPA) = v(4)
        res(n, PL_KLJUC) = CStr(k)
    Next k

    DoplataPlan = SortArrayPoKljucu(res, PL_KLJUC)
End Function

' Ranije AKTIVNE doplate nad istim kljucem stope i preklopljenim periodom
' (DOP-ODL-06). "" kad ih nema; inace tekst za operatera koji ih imenuje.
' Preklop je zatvoren interval: [Od, Do] obe strane ukljucene, na minut.
Public Function DoplataPreklapanja(ByVal periodOd As Date, ByVal periodDo As Date, _
                                   ByVal stope As Collection) As String
    Const SRC As String = "modDoplata.DoplataPreklapanja"

    Dim cene As Object, vrsteSorte As Object, prikaz As Object
    UcitajStope stope, cene, vrsteSorte, prikaz, SRC

    Dim r As Variant
    r = DoplataStavkeRedovi()
    If Not IsArray(r) Then Exit Function

    Dim brojevi As Object
    Set brojevi = CreateObject("Scripting.Dictionary")
    brojevi.CompareMode = vbTextCompare

    Dim i As Long, kStope As String
    For i = 1 To UBound(r, 1)
        kStope = KljucVrstaSorta(r(i, 7), r(i, 8)) & "|" & UCase$(Trim$(CStr(r(i, 9))))
        If cene.Exists(kStope) Then
            If MinutOd(CDate(r(i, 14))) <= MinutOd(periodDo) And _
               MinutOd(periodOd) <= MinutOd(CDate(r(i, 15))) Then
                brojevi(CStr(r(i, 2))) = True
            End If
        End If
    Next i

    If brojevi.count = 0 Then Exit Function

    Dim t As String, k As Variant, n As Long
    For Each k In brojevi.keys
        n = n + 1
        If n <= 5 Then
            If Len(t) > 0 Then t = t & ", "
            t = t & CStr(k)
        End If
    Next k
    If brojevi.count > 5 Then t = t & " (+" & CStr(brojevi.count - 5) & ")"

    DoplataPreklapanja = "Za iste stope i preklopljen period vec postoji doplata: " & t & _
                         ". Nova doplata se placa NA TO. Potvrdi ponavljanje."
End Function

' Problem sa stopama kao TEKST za operatera, "" kad su ispravne. Isto pravilo
' koje pisac dize kao gresku (UcitajStope) -- jedna implementacija, dva
' pozivaoca: ekran Doplata ga zove pri dodavanju stope, da operater dobije
' razlog uz polje umesto tek pri knjizenju.
Public Function DoplataStopeProblem(ByVal stope As Collection) As String
    Dim cene As Object, vrsteSorte As Object, prikaz As Object
    On Error GoTo EH
    UcitajStope stope, cene, vrsteSorte, prikaz, "modDoplata.DoplataStopeProblem"
    Exit Function
EH:
    DoplataStopeProblem = Err.description
End Function

' ============================================================
' KNJIZENJE -- jedini ulaz upisa
' ============================================================

' Knjizi obracun: jedan KO po kooperantu iz plana, svi pod jednim ObracunID-em,
' u jednoj transakciji. Vraca ObracunID ili "" (outGreska kaze zasto).
'
' datum         datum knjizenja = datum zaduzenja na kartici (DOP-ODL-04)
' periodOd/Do   period otkupa, sa satom (DOP-ODL-03a)
' stope         Collection stopa (v. UcitajStope)
' potvrdaPonavljanja  odgovor operatera na DoplataPreklapanja (DOP-ODL-06)
'
' Kapije pre ijednog upisa (DOP-ODL-10): period je zatvoren i vec prosao, a
' datum knjizenja nije pre kraja perioda ni posle danas -- zaduzenje ne sme da
' stoji na kartici pre robe koju placa. Svaka stopa mora da pogodi bar jedan
' otkup. Nema praznog obracuna.
Public Function KnjiziDoplatu_TX(ByVal datum As Date, _
                                 ByVal periodOd As Date, _
                                 ByVal periodDo As Date, _
                                 ByVal stope As Collection, _
                                 Optional ByVal napomena As String = "", _
                                 Optional ByVal potvrdaPonavljanja As Boolean = False, _
                                 Optional ByRef outGreska As String, _
                                 Optional ByRef outErrNum As Long) As String
    Const SRC As String = "modDoplata.KnjiziDoplatu_TX"

    Dim tx As clsTransaction
    Set tx = New clsTransaction

    outGreska = ""
    outErrNum = 0

    On Error GoTo EH

    ' Sema pre upisa: AppendRow pise POZICIONO. Ide PRE BeginTx, kao kod otkupa.
    modSchema.SchemaReadyOrFail SRC, TBL_DOPLATA & "|" & TBL_DOPLATA_STAVKE

    RequireDatumKnjizenja datum, periodDo, SRC

    Dim plan As Variant, bezOtkupa As String
    plan = DoplataPlan(periodOd, periodDo, stope, bezOtkupa)

    If Len(bezOtkupa) > 0 Then
        Err.Raise DOP_ERR_BASE + 30, SRC, _
                  "Stope bez ijednog aktivnog otkupa u periodu: " & bezOtkupa & _
                  ". Ukloni ih ili ispravi kljuc (kljuc cenovnika: vrsta, sorta, klasa)."
    End If
    If Not IsArray(plan) Then
        Err.Raise DOP_ERR_BASE + 31, SRC, _
                  "U periodu nema nijednog aktivnog otkupa za zadate stope."
    End If

    Dim preklop As String
    preklop = DoplataPreklapanja(periodOd, periodDo, stope)
    If Len(preklop) > 0 And Not potvrdaPonavljanja Then
        Err.Raise DOP_ERR_POTVRDA_PONAVLJANJA, SRC, preklop
    End If

    tx.BeginTx
    tx.AddTableSnapshot TBL_DOPLATA
    tx.AddTableSnapshot TBL_DOPLATA_STAVKE

    KnjiziDoplatu_TX = KnjiziPlan(datum, periodOd, periodDo, plan, napomena, SRC)

    tx.CommitTx
    Set tx = Nothing

    On Error Resume Next
    Monitor_Event _
        eventType:="DOPLATA_KNJIZENA", _
        severity:="INFO", _
        message:="Doplata knjizena. ObracunID=" & KnjiziDoplatu_TX, _
        userId:="Operator", _
        moduleName:="modDoplata", _
        procedureName:=SRC, _
        entityType:="Doplata", _
        entityID:=KnjiziDoplatu_TX, _
        correlationId:=KnjiziDoplatu_TX
    On Error GoTo 0
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx

    outErrNum = errNum
    outGreska = errDesc
    KnjiziDoplatu_TX = ""

    ' Pitanje nije kvar: bez loga i monitora, da log ostane signal.
    If errNum <> DOP_ERR_POTVRDA_PONAVLJANJA Then
        LogError SRC, errDesc, errNum
        Monitor_Error _
            moduleName:="modDoplata", _
            procedureName:=SRC, _
            entityType:="Doplata", _
            entityID:="", _
            correlationId:="", _
            errorNumber:=errNum, _
            errorDescription:=errDesc, _
            errorSource:=errSrc
    End If
    On Error GoTo 0
End Function

' Core -- NE zovi spolja. Jedini ulaz je KnjiziDoplatu_TX, koji drzi snapshot
' obe tabele. Plan je vec poredan po kooperantu, pa jedan prolaz pravi KO po
' kooperantu. Brojevi KO su uzastopni u nizu godine datuma knjizenja.
Private Function KnjiziPlan(ByVal datum As Date, _
                            ByVal periodOd As Date, _
                            ByVal periodDo As Date, _
                            ByVal plan As Variant, _
                            ByVal napomena As String, _
                            ByVal sourceName As String) As String
    Dim obracunID As String
    obracunID = NewEntityID("DOB-")
    If Len(obracunID) = 0 Then
        Err.Raise DOP_ERR_BASE + 40, sourceName, "NewEntityID nije vratio ObracunID."
    End If

    Dim godina As Long, sledeci As Long
    godina = Year(datum)
    sledeci = PoslednjiBrojKO(godina) + 1

    Dim i As Long, tekuciKoop As String, doplataID As String, rb As Long
    Dim rowData() As Variant, stavkaID As String

    For i = 1 To UBound(plan, 1)
        If StrComp(CStr(plan(i, PL_KOOP)), tekuciKoop, vbTextCompare) <> 0 Then
            tekuciKoop = CStr(plan(i, PL_KOOP))
            RequireTacnoJedan TBL_KOOPERANTI, COL_KOOP_ID, tekuciKoop, "KooperantID", sourceName

            doplataID = NewEntityID("DOP-")
            If Len(doplataID) = 0 Then
                Err.Raise DOP_ERR_BASE + 41, sourceName, "NewEntityID nije vratio DoplataID."
            End If

            rowData = PrazanRed(TBL_DOPLATA, sourceName)
            SetRowValueByColumn rowData, TBL_DOPLATA, COL_DOP_ID, doplataID, sourceName
            SetRowValueByColumn rowData, TBL_DOPLATA, COL_DOP_BROJ, BrojKO(sledeci, godina), sourceName
            SetRowValueByColumn rowData, TBL_DOPLATA, COL_DOP_DATUM, datum, sourceName
            SetRowValueByColumn rowData, TBL_DOPLATA, COL_DOP_KOOPERANT, tekuciKoop, sourceName
            SetRowValueByColumn rowData, TBL_DOPLATA, COL_DOP_OBRACUN, obracunID, sourceName
            SetRowValueByColumn rowData, TBL_DOPLATA, COL_DOP_PERIOD_OD, periodOd, sourceName
            SetRowValueByColumn rowData, TBL_DOPLATA, COL_DOP_PERIOD_DO, periodDo, sourceName
            SetRowValueByColumn rowData, TBL_DOPLATA, COL_DOP_NAPOMENA, Trim$(napomena), sourceName
            SetRowValueByColumn rowData, TBL_DOPLATA, COL_STORNIRANO, "", sourceName

            If AppendRow(TBL_DOPLATA, rowData) <= 0 Then
                Err.Raise DOP_ERR_BASE + 42, sourceName, _
                          "AppendRow nije upisao zaglavlje KO za kooperanta " & tekuciKoop & "."
            End If

            sledeci = sledeci + 1
            rb = 0
        End If

        rb = rb + 1
        stavkaID = NewEntityID("DPS-")
        If Len(stavkaID) = 0 Then
            Err.Raise DOP_ERR_BASE + 43, sourceName, "NewEntityID nije vratio DoplataStavkaID."
        End If

        rowData = PrazanRed(TBL_DOPLATA_STAVKE, sourceName)
        SetRowValueByColumn rowData, TBL_DOPLATA_STAVKE, COL_DOPS_ID, stavkaID, sourceName
        SetRowValueByColumn rowData, TBL_DOPLATA_STAVKE, COL_DOPS_DOPLATA_ID, doplataID, sourceName
        SetRowValueByColumn rowData, TBL_DOPLATA_STAVKE, COL_DOPS_RB, rb, sourceName
        SetRowValueByColumn rowData, TBL_DOPLATA_STAVKE, COL_DOPS_STANICA, plan(i, PL_STANICA), sourceName
        SetRowValueByColumn rowData, TBL_DOPLATA_STAVKE, COL_DOPS_VRSTA, plan(i, PL_VRSTA), sourceName
        SetRowValueByColumn rowData, TBL_DOPLATA_STAVKE, COL_DOPS_SORTA, plan(i, PL_SORTA), sourceName
        SetRowValueByColumn rowData, TBL_DOPLATA_STAVKE, COL_DOPS_KLASA, plan(i, PL_KLASA), sourceName
        SetRowValueByColumn rowData, TBL_DOPLATA_STAVKE, COL_DOPS_KOLICINA, plan(i, PL_KOLICINA), sourceName
        SetRowValueByColumn rowData, TBL_DOPLATA_STAVKE, COL_DOPS_CENA, plan(i, PL_CENA), sourceName
        SetRowValueByColumn rowData, TBL_DOPLATA_STAVKE, COL_DOPS_IZNOS, plan(i, PL_IZNOS), sourceName

        If AppendRow(TBL_DOPLATA_STAVKE, rowData) <= 0 Then
            Err.Raise DOP_ERR_BASE + 44, sourceName, _
                      "AppendRow nije upisao stavku KO za kooperanta " & tekuciKoop & "."
        End If
    Next i

    KnjiziPlan = obracunID
End Function

' ============================================================
' CITACI
' ============================================================

' Stavke AKTIVNIH KO, sa zaglavljem. 2D niz (1..n, 1..15):
'   1 DoplataID | 2 BrojDokumenta | 3 Datum | 4 KooperantID | 5 ObracunID
'   6 StanicaID | 7 VrstaVoca | 8 SortaVoca | 9 Klasa | 10 Kolicina | 11 Cena
'   12 Iznos | 13 DoplataStavkaID | 14 PeriodOd | 15 PeriodDo
' redosledom tabele stavki. Empty kad nema nijedne aktivne stavke.
'
' STROG, kao StavkeOtkupaRedovi: dupli DoplataID, stavka bez zaglavlja i
' aktivno zaglavlje bez stavki padaju PO IMENU -- tiha nula bi sakrila dug firme
' prema kooperantu. Stavka storniranog zaglavlja nije aktivna (DOP-03).
Public Function DoplataStavkeRedovi() As Variant
    Const SRC As String = "modDoplata.DoplataStavkeRedovi"

    DoplataStavkeRedovi = Empty

    ' Sveska PRE NADOGRADNJE: obe tabele jos ne postoje (pravi ih EnsureAllTables
    ' na startu), pa KO nema -- prazno, isti obrazac kao modUtovar.UtovarenoPoPreradi.
    ' Bez ovoga bi kartica i Saldo OM pali SVIM kooperantima do prvog restarta.
    ' Samo jedna od dve je drift i pada po imenu nize (RequireColumnIndex).
    If GetTable(TBL_DOPLATA) Is Nothing And GetTable(TBL_DOPLATA_STAVKE) Is Nothing Then
        Exit Function
    End If

    Dim hd As Variant
    hd = GetTableData(TBL_DOPLATA)

    Dim sve As Object, aktivna As Object
    Set sve = CreateObject("Scripting.Dictionary")
    sve.CompareMode = vbTextCompare
    Set aktivna = CreateObject("Scripting.Dictionary")
    aktivna.CompareMode = vbTextCompare

    Dim cId As Long, cBr As Long, cDat As Long, cKoop As Long, cObr As Long
    Dim cOd As Long, cDo As Long
    cId = RequireColumnIndex(TBL_DOPLATA, COL_DOP_ID, SRC)
    cBr = RequireColumnIndex(TBL_DOPLATA, COL_DOP_BROJ, SRC)
    cDat = RequireColumnIndex(TBL_DOPLATA, COL_DOP_DATUM, SRC)
    cKoop = RequireColumnIndex(TBL_DOPLATA, COL_DOP_KOOPERANT, SRC)
    cObr = RequireColumnIndex(TBL_DOPLATA, COL_DOP_OBRACUN, SRC)
    cOd = RequireColumnIndex(TBL_DOPLATA, COL_DOP_PERIOD_OD, SRC)
    cDo = RequireColumnIndex(TBL_DOPLATA, COL_DOP_PERIOD_DO, SRC)

    Dim i As Long, id As String
    If IsArray(hd) Then
        For i = 1 To UBound(hd, 1)
            id = Trim$(NzToText(hd(i, cId)))
            If Len(id) = 0 Then
                Err.Raise DOP_ERR_BASE + 50, SRC, _
                          "Zaglavlje KO bez DoplataID-a (red " & CStr(i) & ")."
            End If
            If sve.Exists(id) Then
                Err.Raise DOP_ERR_BASE + 51, SRC, "DoplataID nije jedinstven: " & id & "."
            End If
            sve.Add id, True
        Next i

        Dim akt As Variant
        akt = ExcludeStornirano(hd, TBL_DOPLATA)
        If IsArray(akt) Then
            For i = 1 To UBound(akt, 1)
                aktivna.Add Trim$(NzToText(akt(i, cId))), _
                    Array(NzToText(akt(i, cBr)), akt(i, cDat), Trim$(NzToText(akt(i, cKoop))), _
                          Trim$(NzToText(akt(i, cObr))), akt(i, cOd), akt(i, cDo))
            Next i
        End If
    End If

    Dim sd As Variant
    sd = GetTableData(TBL_DOPLATA_STAVKE)

    Dim sId As Long, sDop As Long, sSta As Long, sVr As Long, sSo As Long
    Dim sKl As Long, sKol As Long, sCena As Long, sIzn As Long
    sId = RequireColumnIndex(TBL_DOPLATA_STAVKE, COL_DOPS_ID, SRC)
    sDop = RequireColumnIndex(TBL_DOPLATA_STAVKE, COL_DOPS_DOPLATA_ID, SRC)
    sSta = RequireColumnIndex(TBL_DOPLATA_STAVKE, COL_DOPS_STANICA, SRC)
    sVr = RequireColumnIndex(TBL_DOPLATA_STAVKE, COL_DOPS_VRSTA, SRC)
    sSo = RequireColumnIndex(TBL_DOPLATA_STAVKE, COL_DOPS_SORTA, SRC)
    sKl = RequireColumnIndex(TBL_DOPLATA_STAVKE, COL_DOPS_KLASA, SRC)
    sKol = RequireColumnIndex(TBL_DOPLATA_STAVKE, COL_DOPS_KOLICINA, SRC)
    sCena = RequireColumnIndex(TBL_DOPLATA_STAVKE, COL_DOPS_CENA, SRC)
    sIzn = RequireColumnIndex(TBL_DOPLATA_STAVKE, COL_DOPS_IZNOS, SRC)

    Dim imaStavku As Object
    Set imaStavku = CreateObject("Scripting.Dictionary")
    imaStavku.CompareMode = vbTextCompare

    Dim n As Long
    If IsArray(sd) Then
        For i = 1 To UBound(sd, 1)
            id = Trim$(NzToText(sd(i, sDop)))
            If Not sve.Exists(id) Then
                Err.Raise DOP_ERR_BASE + 52, SRC, _
                          "Stavka KO pokazuje na zaglavlje koje ne postoji: DoplataID=" & id & _
                          ", stavka " & NzToText(sd(i, sId)) & "."
            End If
            If aktivna.Exists(id) Then
                If Not IsNumeric(sd(i, sKol)) Or Not IsNumeric(sd(i, sCena)) _
                   Or Not IsNumeric(sd(i, sIzn)) Then
                    Err.Raise DOP_ERR_BASE + 53, SRC, _
                              "Stavka KO nema brojcanu kolicinu, cenu ili iznos: " & _
                              NzToText(sd(i, sId)) & "."
                End If
                imaStavku(id) = True
                n = n + 1
            End If
        Next i
    End If

    Dim k As Variant, h As Variant
    For Each k In aktivna.keys
        If Not imaStavku.Exists(CStr(k)) Then
            h = aktivna(k)
            Err.Raise DOP_ERR_BASE + 54, SRC, _
                      "Aktivan KO nema nijednu stavku: " & CStr(h(0)) & _
                      " (" & CStr(k) & ")."
        End If
    Next k

    If n = 0 Then Exit Function

    Dim res() As Variant
    ReDim res(1 To n, 1 To 15)
    n = 0
    For i = 1 To UBound(sd, 1)
        id = Trim$(NzToText(sd(i, sDop)))
        If aktivna.Exists(id) Then
            h = aktivna(id)
            n = n + 1
            res(n, 1) = id
            res(n, 2) = h(0)
            res(n, 3) = h(1)
            res(n, 4) = h(2)
            res(n, 5) = h(3)
            res(n, 6) = Trim$(NzToText(sd(i, sSta)))
            res(n, 7) = Trim$(NzToText(sd(i, sVr)))
            res(n, 8) = Trim$(NzToText(sd(i, sSo)))
            res(n, 9) = Trim$(NzToText(sd(i, sKl)))
            res(n, 10) = CDbl(sd(i, sKol))
            res(n, 11) = CDbl(sd(i, sCena))
            res(n, 12) = CDbl(sd(i, sIzn))
            res(n, 13) = Trim$(NzToText(sd(i, sId)))
            res(n, 14) = h(4)
            res(n, 15) = h(5)
        End If
    Next i

    DoplataStavkeRedovi = res
End Function

' Iznos AKTIVNIH KO po kooperantu: KooperantID -> din (zbir stavki).
'
' stanicaID  samo stavke te stanice -- stavka pripada stanici na kojoj je roba
'            predata (DOPLATA.md par. 3), pa Saldo OM svake stanice vidi svoj
'            deo, a zbir po stanicama je ceo KO. Prazno = sve stanice.
' datumOd/Do Datum KO = datum knjizenja = datum zaduzenja (DOP-ODL-04), po
'            danu, obe granice ukljucene. Izostavljeno = bez granice.
'
' Jedan citac za Saldo OM i PWA izvoz salda -- isti filter storna i ista
' pripadnost stanici, bez druge kopije.
Public Function DoplataIznosPoKooperantu(Optional ByVal stanicaID As String = "", _
                                         Optional ByVal datumOd As Variant, _
                                         Optional ByVal datumDo As Variant) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Set DoplataIznosPoKooperantu = dict

    Dim r As Variant
    r = DoplataStavkeRedovi()
    If Not IsArray(r) Then Exit Function

    Dim imaOd As Boolean, imaDo As Boolean, danOd As Long, danDo As Long
    If Not IsMissing(datumOd) Then
        If IsDate(datumOd) Then imaOd = True: danOd = Int(CDbl(CDate(datumOd)))
    End If
    If Not IsMissing(datumDo) Then
        If IsDate(datumDo) Then imaDo = True: danDo = Int(CDbl(CDate(datumDo)))
    End If

    Dim i As Long, dan As Long, koop As String, ok As Boolean
    For i = 1 To UBound(r, 1)
        ok = True
        If Len(Trim$(stanicaID)) > 0 Then
            ok = (StrComp(CStr(r(i, 6)), Trim$(stanicaID), vbTextCompare) = 0)
        End If
        If ok And (imaOd Or imaDo) Then
            dan = Int(CDbl(CDate(r(i, 3))))
            If imaOd And dan < danOd Then ok = False
            If imaDo And dan > danDo Then ok = False
        End If
        If ok Then
            koop = CStr(r(i, 4))
            dict(koop) = CDbl(dict(koop)) + CDbl(r(i, 12))
        End If
    Next i
End Function

' AKTIVNI KO jednog kooperanta, jedan red po dokumentu -- za karticu
' (modIzvestaj.ReportKarticaKooperanta). 2D niz (1..n, 1..5):
'   1 DoplataID | 2 BrojDokumenta | 3 Datum | 4 Iznos (zbir stavki)
'   5 delovi opisa: Collection od Array(vrsta, sorta, klasa, kg, cena)
' Stavke istog kljuca i iste stope sa razlicitih stanica se sabiraju: kartica
' je kooperantova, ne stanicina. Redosled: redosled prve stavke u tabeli; poredak
' na kartici ionako daje kljuc hronologije (KART-HRON-01). Empty kad nema KO.
Public Function DoplataKOKooperanta(ByVal kooperantID As String) As Variant
    DoplataKOKooperanta = Empty

    Dim r As Variant
    r = DoplataStavkeRedovi()
    If Not IsArray(r) Then Exit Function

    Dim ko As Object, red As Variant, delovi As Collection, kDeo As Object
    Dim i As Long, id As String, kljuc As String, j As Long, deo As Variant
    Set ko = CreateObject("Scripting.Dictionary")
    ko.CompareMode = vbTextCompare
    Set kDeo = CreateObject("Scripting.Dictionary")
    kDeo.CompareMode = vbTextCompare

    For i = 1 To UBound(r, 1)
        If StrComp(CStr(r(i, 4)), Trim$(kooperantID), vbTextCompare) = 0 Then
            id = CStr(r(i, 1))
            If Not ko.Exists(id) Then
                Set delovi = New Collection
                ko.Add id, Array(r(i, 2), r(i, 3), 0#, delovi)
            End If
            red = ko(id)
            red(2) = CDbl(red(2)) + CDbl(r(i, 12))
            Set delovi = red(3)
            kljuc = id & "|" & UCase$(CStr(r(i, 7))) & "|" & UCase$(CStr(r(i, 8))) & "|" & _
                    UCase$(CStr(r(i, 9))) & "|" & CStr(r(i, 11))
            If kDeo.Exists(kljuc) Then
                j = CLng(kDeo(kljuc))
                deo = delovi(j)
                deo(3) = CDbl(deo(3)) + CDbl(r(i, 10))
                delovi.Remove j
                If j > delovi.count Then
                    delovi.Add deo
                Else
                    delovi.Add deo, Before:=j
                End If
            Else
                delovi.Add Array(CStr(r(i, 7)), CStr(r(i, 8)), CStr(r(i, 9)), CDbl(r(i, 10)), _
                                 CDbl(r(i, 11)))
                kDeo.Add kljuc, delovi.count
            End If
            ko(id) = red
        End If
    Next i

    If ko.count = 0 Then Exit Function

    Dim res() As Variant, n As Long, k As Variant
    ReDim res(1 To ko.count, 1 To 5)
    For Each k In ko.keys
        red = ko(k)
        n = n + 1
        res(n, 1) = CStr(k)
        res(n, 2) = red(0)
        res(n, 3) = red(1)
        res(n, 4) = modNovac.ZaokruziNovac(CDbl(red(2)))
        Set res(n, 5) = red(3)
    Next k
    DoplataKOKooperanta = res
End Function

' ============================================================
' ISPLATA KO (F4) -- otvorena obaveza po dokumentu
' ============================================================
'
' KO je obaveza firme prema kooperantu kao i otkupni blok: isplacuje se kroz
' naloge za banku i zatvara uvozom izvoda (DOP-ODL-05). Isplaceno je zbir
' AKTIVNOG novca vezanog DoplataID-em (modNovac.BuildIsplataDictByDoplata) --
' isto pravilo kao za otkup (BuildIsplataDictByOtkup), drugi kljuc.

' OTVORENA KO. Blizanac modNovac.GetOpenOtkupi; cita ga
' modBankaExportPregled.BuildBlokIsplataList. 2D niz (1..n, 1..8):
'   1 DoplataID | 2 BrojDokumenta | 3 Datum | 4 KooperantID
'   5 StanicaID -- MATICNA stanica kooperanta (DOP-ODL-05), ne stanica stavke:
'     KO je jedan dokument i kad ima stavke na dve stanice
'   6 Iznos | 7 Isplaceno | 8 Otvoreno
' U listi je KO ciji je ostatak, u cent-domenu, veci od nule. kooperantID suzava
' na jednog kooperanta. Empty kad nema nijednog.
Public Function OtvoreniKORedovi(Optional ByVal kooperantID As String = "") As Variant
    OtvoreniKORedovi = Empty

    Dim glava As Object, iznosi As Object, isplate As Object
    Set iznosi = IznosPoKO(glava)
    If iznosi.count = 0 Then Exit Function
    Set isplate = modNovac.BuildIsplataDictByDoplata()

    Dim maticna As Object
    Set maticna = BuildLookupDict(TBL_KOOPERANTI, COL_KOOP_ID, COL_KOOP_STANICA)

    Dim k As Variant, h As Variant, n As Long, otv As Double, ispl As Double
    Dim res() As Variant
    ReDim res(1 To iznosi.count, 1 To 8)
    For Each k In iznosi.keys
        h = glava(k)
        If Len(Trim$(kooperantID)) > 0 Then
            If StrComp(CStr(h(2)), Trim$(kooperantID), vbTextCompare) <> 0 Then GoTo Sledeci
        End If
        ispl = 0
        If isplate.Exists(CStr(k)) Then ispl = CDbl(isplate(CStr(k)))
        otv = modNovac.ZaokruziNovac(CDbl(iznosi(k)) - ispl)
        If otv <= 0 Then GoTo Sledeci
        n = n + 1
        res(n, 1) = CStr(k)
        res(n, 2) = h(0)
        res(n, 3) = h(1)
        res(n, 4) = CStr(h(2))
        res(n, 5) = ""
        If maticna.Exists(CStr(h(2))) Then res(n, 5) = Trim$(CStr(maticna(CStr(h(2)))))
        res(n, 6) = modNovac.ZaokruziNovac(CDbl(iznosi(k)))
        res(n, 7) = modNovac.ZaokruziNovac(ispl)
        res(n, 8) = otv
Sledeci:
    Next k
    If n = 0 Then Exit Function

    Dim outA() As Variant, i As Long, j As Long
    ReDim outA(1 To n, 1 To 8)
    For i = 1 To n
        For j = 1 To 8
            outA(i, j) = res(i, j)
        Next j
    Next i
    OtvoreniKORedovi = outA
End Function

' Otvoreno na JEDNOM KO, u cent-domenu: iznos minus isplaceno. 0 za storniran ili
' nepostojeci KO -- kapiju vlasnistva i zivota drzi pisac pre ovoga
' (modBankaMapiranje.PotvrdiKOZaKooperanta).
Public Function OtvorenoNaKO(ByVal doplataID As String) As Double
    Dim glava As Object, iznosi As Object, isplate As Object, id As String
    id = Trim$(doplataID)
    If Len(id) = 0 Then Exit Function
    Set iznosi = IznosPoKO(glava)
    If Not iznosi.Exists(id) Then Exit Function
    Set isplate = modNovac.BuildIsplataDictByDoplata()
    Dim ispl As Double
    If isplate.Exists(id) Then ispl = CDbl(isplate(id))
    OtvorenoNaKO = modNovac.ZaokruziNovac(CDbl(iznosi(id)) - ispl)
End Function

' POZIV NA BROJ -> AKTIVAN KO. Array(DoplataID, KooperantID) kad TACNO jedan
' aktivan KO nosi taj broj; Empty inace. Blizanac
' modBankaMapiranje.TryResolveKooperantByOtkupPoziv: isto poredjenje
' (NormalizePozivKey, slova i cifre), isti odgovor na storno (storniran KO se ne
' razresava -- uplata tada ide putem kooperanta, kao za storniran otkup).
'
' Broj KO nosi slovni prefiks (DOP_BROJ_PREFIKS), a broj otkupa je od cifara
' (x/ddmmyy), pa se dva niza ne mogu sudariti. Poziv bez prefiksa zato ni ne
' cita tabelu: mreza uvoza zove ovo za SVAKI red izvoda.
Public Function KOPoPozivu(ByVal pozivNaBroj As String) As Variant
    Const SRC As String = "modDoplata.KOPoPozivu"
    KOPoPozivu = Empty

    Dim target As String, pref As String
    target = modBankaMapiranje.NormalizePozivKey(pozivNaBroj)
    pref = modBankaMapiranje.NormalizePozivKey(DOP_BROJ_PREFIKS)
    If Len(target) <= Len(pref) Then Exit Function
    If Left$(target, Len(pref)) <> pref Then Exit Function
    If GetTable(TBL_DOPLATA) Is Nothing Then Exit Function

    Dim d As Variant
    d = GetTableData(TBL_DOPLATA)
    If Not IsArray(d) Then Exit Function
    d = ExcludeStornirano(d, TBL_DOPLATA)
    If Not IsArray(d) Then Exit Function

    Dim cId As Long, cBr As Long, cKoop As Long
    cId = RequireColumnIndex(TBL_DOPLATA, COL_DOP_ID, SRC)
    cBr = RequireColumnIndex(TBL_DOPLATA, COL_DOP_BROJ, SRC)
    cKoop = RequireColumnIndex(TBL_DOPLATA, COL_DOP_KOOPERANT, SRC)

    Dim i As Long, hits As Long, hitID As String, hitKoop As String
    For i = 1 To UBound(d, 1)
        If modBankaMapiranje.NormalizePozivKey(NzToText(d(i, cBr))) = target Then
            hits = hits + 1
            hitID = Trim$(NzToText(d(i, cId)))
            hitKoop = Trim$(NzToText(d(i, cKoop)))
        End If
    Next i

    If hits = 1 And Len(hitID) > 0 And Len(hitKoop) > 0 Then KOPoPozivu = Array(hitID, hitKoop)
End Function

' Da li je ID knjizno odobrenje (bilo kog statusa). Lista naloga nosi KO pod
' istim kljucem bloka kao otkup (clsBlokIsplata.otkupID), pa radnja koja vazi
' samo za otkup (primena avansa) ovim pita pre nego sto ponudi potvrdu.
Public Function JeKO(ByVal id As String) As Boolean
    If Len(Trim$(id)) = 0 Then Exit Function
    If GetTable(TBL_DOPLATA) Is Nothing Then Exit Function
    JeKO = (Len(NzToText(LookupValue(TBL_DOPLATA, COL_DOP_ID, Trim$(id), COL_DOP_ID))) > 0)
End Function

' Iznos svakog AKTIVNOG KO: DoplataID -> din (zbir stavki, DoplataStavkeRedovi --
' strog citac, pa storniran KO ne ulazi). outGlava: DoplataID ->
' Array(BrojDokumenta, Datum, KooperantID).
Private Function IznosPoKO(ByRef outGlava As Object) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set outGlava = CreateObject("Scripting.Dictionary")
    outGlava.CompareMode = vbTextCompare
    Set IznosPoKO = d

    Dim r As Variant
    r = DoplataStavkeRedovi()
    If Not IsArray(r) Then Exit Function

    Dim i As Long, id As String
    For i = 1 To UBound(r, 1)
        id = CStr(r(i, 1))
        If Not d.Exists(id) Then
            d.Add id, 0#
            outGlava.Add id, Array(r(i, 2), r(i, 3), CStr(r(i, 4)))
        End If
        d(id) = CDbl(d(id)) + CDbl(r(i, 12))
    Next i
End Function

' OSNOVICA KNJIZENE DOPLATE SE PROMENILA (DOP-ODL-07). Knjizen KO je SNIMAK:
' kasniji storno, ispravka ili retro unos otkupa u periodu ga ne menja -- ovde
' se to VIDI. Plan svakog aktivnog obracuna racuna se ponovo nad danasnjim
' otkupima (isti DoplataPlan, iste stope i isti period) i poredi po kooperantu,
' stanici i kljucu. Signal, ne popravka: ispravka je storno KO pa nov obracun.
'
' 2D niz (1..n, 1..6): 1 BrojKO ("" kad kooperant tada nije bio obuhvacen) |
' 2 KooperantID | 3 kljuc ("vrsta sorta klasa @ stanica") | 4 kg knjizeno |
' 5 kg sada | 6 ObracunID. Obracun ciji se plan danas ne moze izracunati (npr.
' granicni dan bez poznatog sata) daje JEDAN red sa razlogom u koloni 3.
' Empty kad odstupanja nema. Read-only.
Public Function DoplataOdstupanjaOsnovice() As Variant
    DoplataOdstupanjaOsnovice = Empty

    Dim r As Variant
    r = DoplataStavkeRedovi()
    If Not IsArray(r) Then Exit Function

    ' Obracun -> Array(periodOd, periodDo, stope, knjizeno). Recnici se drze
    ' kao objekti u nizu i pune preko svoje promenljive -- bez lanaca tipa
    ' d(k)(i), koje ovaj modul namerno izbegava.
    '   stope:    kljuc stope -> Array(vrsta, sorta, klasa, cena)
    '   knjizeno: kooperant|stanica|kljuc stope -> Array(kg, brojKO, koop, opis)
    Dim obr As Object, o As Variant, k As String, kStope As String, i As Long
    Dim st As Object, kn As Object
    Set obr = CreateObject("Scripting.Dictionary")
    obr.CompareMode = vbTextCompare
    For i = 1 To UBound(r, 1)
        k = CStr(r(i, 5))
        If Not obr.Exists(k) Then
            Set st = CreateObject("Scripting.Dictionary")
            st.CompareMode = vbTextCompare
            Set kn = CreateObject("Scripting.Dictionary")
            kn.CompareMode = vbTextCompare
            obr.Add k, Array(r(i, 14), r(i, 15), st, kn)
        End If
        o = obr(k)
        Set st = o(2)
        Set kn = o(3)
        kStope = KljucVrstaSorta(r(i, 7), r(i, 8)) & "|" & UCase$(Trim$(CStr(r(i, 9))))
        If Not st.Exists(kStope) Then _
            st.Add kStope, Array(CStr(r(i, 7)), CStr(r(i, 8)), CStr(r(i, 9)), CDbl(r(i, 11)))
        kn.Add UCase$(CStr(r(i, 4))) & "|" & UCase$(CStr(r(i, 6))) & "|" & kStope, _
               Array(CDbl(r(i, 10)), CStr(r(i, 2)), CStr(r(i, 4)), _
                     OpisKljuca(r(i, 7), r(i, 8), r(i, 9), r(i, 6)))
    Next i

    Dim out As New Collection, ok As Variant, stope As Collection, s As Object
    Dim plan As Variant, bez As String, sada As Object, kk As Variant, j As Long
    Dim kgKnj As Double, kgSada As Double, deo As Variant, greska As String
    Dim vKnj As Variant, vSada As Variant
    For Each ok In obr.keys
        o = obr(ok)
        Set st = o(2)
        Set kn = o(3)
        Set stope = New Collection
        For Each kk In st.keys
            deo = st(kk)
            Set s = CreateObject("Scripting.Dictionary")
            s.Add "VrstaVoca", deo(0)
            s.Add "SortaVoca", deo(1)
            s.Add "Klasa", deo(2)
            s.Add "Cena", deo(3)
            stope.Add s
        Next kk

        greska = ""
        plan = Empty
        On Error Resume Next
        plan = DoplataPlan(CDate(o(0)), CDate(o(1)), stope, bez)
        If Err.Number <> 0 Then greska = Err.description
        Err.Clear
        On Error GoTo 0

        If Len(greska) > 0 Then
            out.Add Array("", "", "osnovica se ne moze izracunati: " & greska, 0#, 0#, CStr(ok))
        Else
            Set sada = CreateObject("Scripting.Dictionary")
            sada.CompareMode = vbTextCompare
            If IsArray(plan) Then
                For j = 1 To UBound(plan, 1)
                    k = UCase$(CStr(plan(j, PL_KOOP))) & "|" & UCase$(CStr(plan(j, PL_STANICA))) & "|" & _
                        KljucVrstaSorta(plan(j, PL_VRSTA), plan(j, PL_SORTA)) & "|" & _
                        UCase$(Trim$(CStr(plan(j, PL_KLASA))))
                    sada.Add k, Array(CDbl(plan(j, PL_KOLICINA)), CStr(plan(j, PL_KOOP)), _
                                      OpisKljuca(plan(j, PL_VRSTA), plan(j, PL_SORTA), _
                                                 plan(j, PL_KLASA), plan(j, PL_STANICA)))
                Next j
            End If
            For Each kk In kn.keys
                vKnj = kn(kk)
                kgKnj = CDbl(vKnj(0))
                kgSada = 0
                If sada.Exists(kk) Then
                    vSada = sada(kk)
                    kgSada = CDbl(vSada(0))
                End If
                If Abs(kgSada - kgKnj) >= 0.005 Then
                    out.Add Array(vKnj(1), vKnj(2), vKnj(3), kgKnj, kgSada, CStr(ok))
                End If
            Next kk
            For Each kk In sada.keys
                If Not kn.Exists(kk) Then
                    vSada = sada(kk)
                    out.Add Array("", vSada(1), vSada(2), 0#, CDbl(vSada(0)), CStr(ok))
                End If
            Next kk
        End If
    Next ok

    If out.count = 0 Then Exit Function
    Dim res() As Variant, red As Variant
    ReDim res(1 To out.count, 1 To 6)
    For i = 1 To out.count
        red = out(i)
        For j = 1 To 6
            res(i, j) = red(j - 1)
        Next j
    Next i
    DoplataOdstupanjaOsnovice = res
End Function

' "vrsta sorta klasa @ stanica" -- citljiv kljuc reda za nalaz integriteta.
Private Function OpisKljuca(ByVal vrsta As Variant, ByVal sorta As Variant, _
                            ByVal klasa As Variant, ByVal stanica As Variant) As String
    OpisKljuca = Trim$(Trim$(NzToText(vrsta) & " " & NzToText(sorta)) & " " & NzToText(klasa)) & _
                 " @ " & NzToText(stanica)
End Function

' DoplataID-evi AKTIVNIH KO jednog obracuna -- za storno celog obracuna
' (modStorno.StornoObracunDoplate_TX). Prazna kolekcija kad ih nema.
Public Function AktivneDoplateObracuna(ByVal obracunID As String) As Collection
    Const SRC As String = "modDoplata.AktivneDoplateObracuna"

    Dim c As New Collection
    Set AktivneDoplateObracuna = c

    obracunID = Trim$(obracunID)
    If Len(obracunID) = 0 Then Exit Function

    Dim d As Variant
    d = GetTableData(TBL_DOPLATA)
    If Not IsArray(d) Then Exit Function
    d = ExcludeStornirano(d, TBL_DOPLATA)
    If Not IsArray(d) Then Exit Function

    Dim cId As Long, cObr As Long, i As Long
    cId = RequireColumnIndex(TBL_DOPLATA, COL_DOP_ID, SRC)
    cObr = RequireColumnIndex(TBL_DOPLATA, COL_DOP_OBRACUN, SRC)
    For i = 1 To UBound(d, 1)
        If StrComp(Trim$(NzToText(d(i, cObr))), obracunID, vbTextCompare) = 0 Then
            c.Add Trim$(NzToText(d(i, cId)))
        End If
    Next i
End Function

' ============================================================
' BROJ KO -- "KO-<n>/<godina>" (DOP-ODL-09)
' ============================================================

Public Function BrojKO(ByVal n As Long, ByVal godina As Long) As String
    BrojKO = DOP_BROJ_PREFIKS & CStr(n) & "/" & CStr(godina)
End Function

' Najveci n u nizu godine, preko SVIH KO -- i storniranih: storno ne oslobadja
' broj (A9). 0 kad niz te godine jos nema broj.
Public Function PoslednjiBrojKO(ByVal godina As Long) As Long
    Const SRC As String = "modDoplata.PoslednjiBrojKO"

    Dim d As Variant
    d = GetTableData(TBL_DOPLATA)
    If Not IsArray(d) Then Exit Function

    Dim cBr As Long, i As Long, n As Long
    cBr = RequireColumnIndex(TBL_DOPLATA, COL_DOP_BROJ, SRC)
    For i = 1 To UBound(d, 1)
        n = BrojKOIzTeksta(NzToText(d(i, cBr)), godina)
        If n > PoslednjiBrojKO Then PoslednjiBrojKO = n
    Next i
End Function

' n iz "KO-<n>/<godina>" kad je godina ta; 0 inace.
Private Function BrojKOIzTeksta(ByVal broj As String, ByVal godina As Long) As Long
    Dim t As String, p As Long, deoN As String, deoG As String
    t = UCase$(Trim$(broj))
    If Left$(t, Len(DOP_BROJ_PREFIKS)) <> UCase$(DOP_BROJ_PREFIKS) Then Exit Function
    t = Mid$(t, Len(DOP_BROJ_PREFIKS) + 1)
    p = InStr(t, "/")
    If p <= 1 Then Exit Function
    deoN = Left$(t, p - 1)
    deoG = Mid$(t, p + 1)
    If Not (deoN Like String$(Len(deoN), "#")) Then Exit Function
    If Not (deoG Like String$(Len(deoG), "#")) Or Len(deoG) = 0 Then Exit Function
    If CLng(deoG) <> godina Then Exit Function
    BrojKOIzTeksta = CLng(deoN)
End Function

' ============================================================
' PERIOD SA SATOM (DOP-ODL-03a)
' ============================================================

' Da li otkup ulazi u period [Od, Do], obe granice ukljucene, na minut.
'
' DAN je Datum dokumenta. SAT je vreme nastanka po pravilu otkupnog lista
' (modOtkup.OtkupVremeNastanka: SourceCreatedAt, pa CreatedAt), i trazi se SAMO
' kad granica sece dan: Od posle 00:00 na prvom danu, Do pre 23:59 na poslednjem.
' Unutrasnji dani i cele granicne dane odlucuje datum.
'
' STROZE OD STAMPE, namerno: kad SourceCreatedAt postoji a nije vreme (PWA salje
' ISO tekst), sat se NE uzima iz CreatedAt -- za PWA red to je trenutak
' SINHRONIZACIJE, ne otkupa. Tada je sat nepoznat, a otkup na granicnom danu
' obara obracun po imenu: pogresan dinar na granici je gori od zastoja.
Private Function OtkupUPeriodu(ByVal datum As Variant, ByVal srcCreated As Variant, _
                               ByVal created As Variant, ByVal periodOd As Date, _
                               ByVal periodDo As Date, ByVal oid As String, _
                               ByVal broj As String, ByVal sourceName As String) As Boolean
    If Not IsDate(datum) Then
        Err.Raise DOP_ERR_BASE + 20, sourceName, _
                  "Otkup " & broj & " (" & oid & ") nema ispravan Datum."
    End If

    Dim dan As Long, danOd As Long, danDo As Long
    dan = Int(CDbl(CDate(datum)))
    danOd = Int(CDbl(periodOd))
    danDo = Int(CDbl(periodDo))

    If dan < danOd Or dan > danDo Then Exit Function

    Dim trebaSat As Boolean
    If dan = danOd And (MinutOd(periodOd) Mod MIN_U_DANU) > 0 Then trebaSat = True
    If dan = danDo And (MinutOd(periodDo) Mod MIN_U_DANU) < MIN_U_DANU - 1 Then trebaSat = True

    If Not trebaSat Then
        OtkupUPeriodu = True
        Exit Function
    End If

    Dim v As Variant
    If Len(Trim$(NzToText(srcCreated))) > 0 And Not IsDate(srcCreated) Then
        v = Empty
    Else
        v = modOtkup.OtkupVremeNastanka(srcCreated, created)
    End If

    If Not IsDate(v) Then
        Err.Raise DOP_ERR_BASE + 21, sourceName, _
                  "Otkup " & broj & " (" & oid & ") je na granicnom danu perioda, a vreme " & _
                  "nastanka mu nije poznato (SourceCreatedAt: '" & NzToText(srcCreated) & _
                  "'). Pomeri granicu na pocetak ili kraj dana."
    End If

    Dim m As Double
    m = CDbl(dan) * MIN_U_DANU + Hour(CDate(v)) * 60 + Minute(CDate(v))
    OtkupUPeriodu = (m >= MinutOd(periodOd) And m <= MinutOd(periodDo))
End Function

' Trenutak na minut: dan * 1440 + sat * 60 + minut. Sekunde ne ulaze.
Private Function MinutOd(ByVal t As Date) As Double
    MinutOd = Int(CDbl(t)) * CDbl(MIN_U_DANU) + Hour(t) * 60 + Minute(t)
End Function

' ============================================================
' KAPIJE ULAZA
' ============================================================

Private Sub RequirePeriod(ByVal periodOd As Date, ByVal periodDo As Date, _
                          ByVal sourceName As String)
    If MinutOd(periodOd) > MinutOd(periodDo) Then
        Err.Raise DOP_ERR_BASE + 10, sourceName, _
                  "Pocetak perioda je posle kraja: " & Format$(periodOd, "dd.mm.yyyy hh:nn") & _
                  " > " & Format$(periodDo, "dd.mm.yyyy hh:nn") & "."
    End If
    If MinutOd(periodDo) > MinutOd(Now) Then
        Err.Raise DOP_ERR_BASE + 11, sourceName, _
                  "Kraj perioda je u buducnosti (" & Format$(periodDo, "dd.mm.yyyy hh:nn") & _
                  "). Doplata se obracunava za otkupe koji su vec nastali."
    End If
End Sub

Private Sub RequireDatumKnjizenja(ByVal datum As Date, ByVal periodDo As Date, _
                                  ByVal sourceName As String)
    If Int(CDbl(datum)) < Int(CDbl(periodDo)) Then
        Err.Raise DOP_ERR_BASE + 12, sourceName, _
                  "Datum knjizenja (" & Format$(datum, "dd.mm.yyyy") & ") je pre kraja " & _
                  "perioda (" & Format$(periodDo, "dd.mm.yyyy") & "). Zaduzenje ne sme da " & _
                  "stoji na kartici pre robe koju placa."
    End If
    If Int(CDbl(datum)) > Int(CDbl(Date)) Then
        Err.Raise DOP_ERR_BASE + 13, sourceName, _
                  "Datum knjizenja je u buducnosti: " & Format$(datum, "dd.mm.yyyy") & "."
    End If
End Sub

' Stope -> tri recnika:
'   cene        "VRSTA|SORTA|KLASA" -> din/kg
'   vrsteSorte  "VRSTA|SORTA"       -> True   (prefilter zaglavlja)
'   prikaz      "VRSTA|SORTA|KLASA" -> Array(vrsta, sorta, klasa) kako je uneto
' Kljuc je kljuc cenovnika (DOP-ODL-02): trim, bez razlike u velicini slova,
' prazna sorta = kultura bez sorte (ne "sve sorte"). Klasa I ili II, cena > 0,
' bez dupliranog kljuca.
Private Sub UcitajStope(ByVal stope As Collection, ByRef cene As Object, _
                        ByRef vrsteSorte As Object, ByRef prikaz As Object, _
                        ByVal sourceName As String)
    If stope Is Nothing Then
        Err.Raise DOP_ERR_BASE + 1, sourceName, "Nije zadata nijedna stopa doplate."
    End If
    If stope.count = 0 Then
        Err.Raise DOP_ERR_BASE + 1, sourceName, "Nije zadata nijedna stopa doplate."
    End If

    Set cene = CreateObject("Scripting.Dictionary")
    cene.CompareMode = vbTextCompare
    Set vrsteSorte = CreateObject("Scripting.Dictionary")
    vrsteSorte.CompareMode = vbTextCompare
    Set prikaz = CreateObject("Scripting.Dictionary")
    prikaz.CompareMode = vbTextCompare

    Dim i As Long, s As Object, k As Variant
    Dim vrsta As String, sorta As String, klasa As String, cena As Variant, kljuc As String

    For i = 1 To stope.count
        If Not IsObject(stope(i)) Then
            Err.Raise DOP_ERR_BASE + 2, sourceName, "Stopa " & CStr(i) & " nije Dictionary."
        End If
        Set s = stope(i)

        For Each k In s.keys
            If InStr(1, DOP_STOPA_KLJUCEVI, "|" & UCase$(CStr(k)) & "|") = 0 Then
                Err.Raise DOP_ERR_BASE + 3, sourceName, _
                          "Stopa " & CStr(i) & " ima nepoznat kljuc: " & CStr(k) & "."
            End If
        Next k
        If Not (s.Exists("VrstaVoca") And s.Exists("SortaVoca") And s.Exists("Klasa") _
                And s.Exists("Cena")) Then
            Err.Raise DOP_ERR_BASE + 4, sourceName, _
                      "Stopa " & CStr(i) & " mora imati VrstaVoca, SortaVoca, Klasa i Cena."
        End If

        vrsta = Trim$(NzToText(s("VrstaVoca")))
        sorta = Trim$(NzToText(s("SortaVoca")))
        klasa = Trim$(NzToText(s("Klasa")))
        cena = s("Cena")

        If Len(vrsta) = 0 Then
            Err.Raise DOP_ERR_BASE + 5, sourceName, "Stopa " & CStr(i) & ": vrsta je obavezna."
        End If
        modOtkup.RequireValidOtkupClass klasa, sourceName
        If Not IsNumeric(cena) Then
            Err.Raise DOP_ERR_BASE + 6, sourceName, _
                      "Stopa " & CStr(i) & ": iznos doplate nije broj."
        End If
        If CDbl(cena) <= 0 Then
            Err.Raise DOP_ERR_BASE + 7, sourceName, _
                      "Stopa " & CStr(i) & ": iznos doplate mora biti veci od nule (din/kg)."
        End If

        kljuc = KljucVrstaSorta(vrsta, sorta) & "|" & UCase$(klasa)
        If cene.Exists(kljuc) Then
            Err.Raise DOP_ERR_BASE + 8, sourceName, _
                      "Ista stopa je zadata dvaput: " & vrsta & " " & sorta & " " & klasa & "."
        End If
        cene.Add kljuc, CDbl(cena)
        prikaz.Add kljuc, Array(vrsta, sorta, klasa)
        vrsteSorte(KljucVrstaSorta(vrsta, sorta)) = True
    Next i
End Sub

' ============================================================
' POMOCNICI
' ============================================================

Private Function KljucVrstaSorta(ByVal vrsta As Variant, ByVal sorta As Variant) As String
    KljucVrstaSorta = UCase$(Trim$(NzToText(vrsta))) & "|" & UCase$(Trim$(NzToText(sorta)))
End Function

' Kljuc poretka plana: kooperant, stanica, vrsta|sorta|klasa, pa klasa I pre II.
' Chr$(1) razdvaja delove (binarno poredjenje, SortArrayPoKljucu).
Private Function KljucPoretka(ByVal koop As String, ByVal stanica As String, _
                              ByVal kStope As String, ByVal klasa As String) As String
    Dim rang As String
    If UCase$(Trim$(klasa)) = UCase$(KLASA_I) Then rang = "1" Else rang = "2"
    KljucPoretka = UCase$(koop) & Chr$(1) & UCase$(stanica) & Chr$(1) & _
                   UCase$(kStope) & Chr$(1) & rang
End Function

' Spisak kljuceva stopa koje NISU pokrivene (pokriveno Nothing = nijedna).
Private Function SpisakKljuceva(ByVal prikaz As Object, ByVal pokriveno As Object) As String
    Dim k As Variant, t As String, deo As Variant
    For Each k In prikaz.keys
        Dim pogodjen As Boolean
        pogodjen = False
        If Not pokriveno Is Nothing Then pogodjen = pokriveno.Exists(CStr(k))
        If Not pogodjen Then
            deo = prikaz(k)
            If Len(t) > 0 Then t = t & "; "
            t = t & Trim$(CStr(deo(0)) & " " & CStr(deo(1))) & " " & CStr(deo(2))
        End If
    Next k
    SpisakKljuceva = t
End Function

Private Function PrazanRed(ByVal tblName As String, ByVal sourceName As String) As Variant()
    Dim n As Long
    n = TabelaBrojKolona(tblName)
    If n <= 0 Then
        Err.Raise DOP_ERR_BASE + 45, sourceName, _
                  "Ne mogu da odredim broj kolona za " & tblName & "."
    End If
    Dim r() As Variant
    ReDim r(0 To n - 1)
    PrazanRed = r
End Function
