Attribute VB_Name = "modCenovnik"
'Attribute VB_Name = "modCenovnik"
Option Explicit

' ============================================================
' modCenovnik - otkupna cena voca (docs/DOMEN/CENOVNIK.md, CEN-01..13)
'
' Model je preuzet iz AgriXSystems (ADR 0019 sa dopunom 11. 10., ADR 0027,
' packages/domain/src/maticni/cena.ts) i preveden na tekstualni kljuc ove
' sveske (VrstaVoca + SortaVoca + Klasa, DOP-ODL-02):
'
'   - Red je ODLUKA FIRME o ceni (din/kg, sa PDV-om) za robu i klasu, od dana
'     Datum, za celu firmu (StanicaID prazno = GLOBALNA) ili za jednu stanicu
'     (POSEBNA).
'   - Red samo-dodavanje: posle unosa se ne menja (CEN-02). Pogresan red se
'     stornira -- razlog, autor i vreme zajedno, jednom i konacno (CEN-11,
'     CEN-12). Ispravka je storno starog i nov red u JEDNOJ transakciji, sa
'     vezom u oba smera (CEN-13).
'   - Najvise jedan nestorniran red po (vrsta, sorta, klasa, stanica, dan)
'     (CEN-03). Drzi ga pisac (CenaUnosProblem); citalac prekrsen preduslov
'     prijavljuje greskom, ne bira (CenPobednik). Zatecene duplikate starog
'     modela resava MigrirajDupleCene (CEN-ODL-04), iz Ensure-a.
'   - VAZECA cena za (stanica, roba, klasa, dan): najkasniji nestorniran
'     Datum <= dan; istog dana posebna pobedjuje globalnu (CEN-05). Sorta bez
'     reda koristi REZERVU: istu vrstu bez sorte, po istom pravilu; sorta
'     pobedjuje i noviju rezervu (CEN-06). Bez reda NEMA cene -- nikad 0
'     (CEN-07). Buduca cena ne vazi pre svog dana (stari bag ne prelazi).
'
' Pravilo ima JEDNOG vlasnika: CenPobednik / CenVazeciRed, ciste funkcije nad
' normalizovanim nizom (CEN_R_*). Citalac (TryVazecaCena) i pregled pre upisa
' (PrekinutePosebneCene) zovu isto pravilo, pa se ne mogu razici.
'
' Odnos prema drugom modelu cene (NE duplirati, .claude/rules/agrohemija-i-cene.md):
' tblArtikli.CenaPoJedinici je jedna tekuca cena artikla agrohemije; ovaj modul
' pokriva samo otkup voca. Cena na stavci dokumenta je SNIMAK (predlog odavde,
' DOCUMENT_HEADER_LINES S4): storno ili ispravka cenovnika ne menja izdat otkup.
'
' PWA jos NE cita tblCenovnik (otkup-form.js uzima Config kljuc Cena<Vrsta>);
' to je zaseban izvor cene na terenu, van ovog modula.
' ============================================================

' Normalizovan red cenovnika za pravilo (CenovnikRedovi i testovi grade isti
' oblik): 2D niz (1..n, 1..CEN_R_KOLONA).
Public Const CEN_R_ID As Long = 1
Public Const CEN_R_VRSTA As Long = 2
Public Const CEN_R_SORTA As Long = 3
Public Const CEN_R_KLASA As Long = 4
Public Const CEN_R_STANICA As Long = 5      ' "" = globalna
Public Const CEN_R_DAN As Long = 6          ' Long serijski dan; 0 = red bez ispravnog datuma (ne ucestvuje)
Public Const CEN_R_CENA As Long = 7         ' Double
Public Const CEN_R_STORNO As Long = 8       ' Boolean
Public Const CEN_R_KOLONA As Long = 8

' Prekrsen preduslov CEN-03 u citanju: dva nestornirana reda istog kljuca i dana.
Public Const ERR_CEN_DUPLI_KLJUC As Long = vbObjectError + 7710

Private Const ERR_CEN_UNOS As Long = vbObjectError + 7706
Private Const ERR_CEN_STORNIRANA As Long = vbObjectError + 7707
Private Const ERR_CEN_ID As Long = vbObjectError + 7708
Private Const ERR_CEN_UPIS As Long = vbObjectError + 7709

Private Const STORNO_DA As String = "Da"
' Gornja granica cene: numeric(12,2) u AgriXSystems.
Private Const CEN_MAX As Double = 9999999999.99

' Razlog storna u migraciji (CEN-ODL-04). ASCII namerno: ide u podatke, a izvor
' mora ostati ASCII.
Private Const MIG_RAZLOG_PRE As String = "Zamenjen kasnijim redom "
Private Const MIG_RAZLOG_POSLE As String = " istog kljuca i dana (migracija CEN-ODL-04)."

' ============================================================
' PRAVILO (ciste funkcije, bez sveske)
' ============================================================

' Pobednik za JEDNU robu (vrsta + sorta) i klasu: indeks reda u `redovi` ili 0.
' Kandidati su nestornirani redovi sa danom 1..dan, globalni ili bas te stanice
' (prazna stanica upita vidi samo globalne). Pobedjuje najkasniji dan; istog
' dana posebna pobedjuje globalnu. Dva reda istog nivoa istog dana su
' prekrsen preduslov (CEN-03), ne izbor -- greska ERR_CEN_DUPLI_KLJUC.
Public Function CenPobednik(ByRef redovi As Variant, ByVal vrsta As String, _
                            ByVal sorta As String, ByVal klasa As String, _
                            ByVal stanicaID As String, ByVal dan As Date) As Long
    Dim i As Long, best As Long, d As Long, danUpita As Long
    Dim st As String, bestSt As String

    CenPobednik = 0
    If Not IsArray(redovi) Then Exit Function
    danUpita = CLng(Int(CDbl(dan)))

    For i = LBound(redovi, 1) To UBound(redovi, 1)
        If Not CBool(redovi(i, CEN_R_STORNO)) Then
            d = CLng(redovi(i, CEN_R_DAN))
            If d > 0 And d <= danUpita Then
                If IstiTekst(redovi(i, CEN_R_VRSTA), vrsta) Then
                If IstiTekst(redovi(i, CEN_R_SORTA), sorta) Then
                If IstiTekst(redovi(i, CEN_R_KLASA), klasa) Then
                    st = Trim$(CStr(redovi(i, CEN_R_STANICA)))
                    If Len(st) = 0 Or IstiTekst(st, stanicaID) Then
                        If best = 0 Then
                            best = i
                        ElseIf d > CLng(redovi(best, CEN_R_DAN)) Then
                            best = i
                        ElseIf d = CLng(redovi(best, CEN_R_DAN)) Then
                            bestSt = Trim$(CStr(redovi(best, CEN_R_STANICA)))
                            If (Len(st) = 0) = (Len(bestSt) = 0) Then
                                Err.Raise ERR_CEN_DUPLI_KLJUC, "modCenovnik.CenPobednik", _
                                          "Dva nestornirana reda istog kljuca i dana (CEN-03): " & _
                                          CStr(redovi(best, CEN_R_ID)) & ", " & CStr(redovi(i, CEN_R_ID))
                            End If
                            If Len(st) > 0 Then best = i
                        End If
                    End If
                End If
                End If
                End If
            End If
        End If
    Next i

    CenPobednik = best
End Function

' Vazeca cena (CEN-05..07): indeks reda ili 0 = bez cene. Kad sorta nema red,
' isto pravilo vazi za istu vrstu bez sorte (rezerva = True); sorta pobedjuje i
' noviju rezervu, jer se rezerva trazi tek kad sorta nema NIJEDAN vazeci red.
Public Function CenVazeciRed(ByRef redovi As Variant, ByVal vrsta As String, _
                             ByVal sorta As String, ByVal klasa As String, _
                             ByVal stanicaID As String, ByVal dan As Date, _
                             ByRef rezerva As Boolean) As Long
    rezerva = False
    CenVazeciRed = CenPobednik(redovi, vrsta, sorta, klasa, stanicaID, dan)
    If CenVazeciRed > 0 Then Exit Function
    If Len(Trim$(sorta)) = 0 Then Exit Function
    CenVazeciRed = CenPobednik(redovi, vrsta, "", klasa, stanicaID, dan)
    rezerva = (CenVazeciRed > 0)
End Function

' ============================================================
' CITANJE
' ============================================================

' Redovi tblCenovnik u obliku pravila (CEN_R_*), redom tabele: indeks reda je
' indeks u tabeli. Empty kad tabela nema redova; nepostojeca tabela ili kolona
' je greska (prazno i nepostojece nisu isti ishod, modSchemaGuard.RequireTable).
Public Function CenovnikRedovi() As Variant
    Const SRC As String = "modCenovnik.CenovnikRedovi"
    Dim data As Variant, res() As Variant, i As Long
    Dim cId As Long, cVr As Long, cSo As Long, cKl As Long
    Dim cSt As Long, cDa As Long, cCe As Long, cSto As Long

    RequireTable TBL_CENOVNIK, SRC
    cId = RequireColumnIndex(TBL_CENOVNIK, COL_CEN_ID, SRC)
    cVr = RequireColumnIndex(TBL_CENOVNIK, COL_CEN_VRSTA, SRC)
    cSo = RequireColumnIndex(TBL_CENOVNIK, COL_CEN_SORTA, SRC)
    cKl = RequireColumnIndex(TBL_CENOVNIK, COL_CEN_KLASA, SRC)
    cSt = RequireColumnIndex(TBL_CENOVNIK, COL_CEN_STANICA, SRC)
    cDa = RequireColumnIndex(TBL_CENOVNIK, COL_CEN_DATUM, SRC)
    cCe = RequireColumnIndex(TBL_CENOVNIK, COL_CEN_CENA, SRC)
    cSto = RequireColumnIndex(TBL_CENOVNIK, COL_STORNIRANO, SRC)

    data = GetTableData(TBL_CENOVNIK)
    If Not IsArray(data) Then Exit Function

    ReDim res(1 To UBound(data, 1), 1 To CEN_R_KOLONA)
    For i = 1 To UBound(data, 1)
        res(i, CEN_R_ID) = Trim$(NzToText(data(i, cId)))
        res(i, CEN_R_VRSTA) = Trim$(NzToText(data(i, cVr)))
        res(i, CEN_R_SORTA) = Trim$(NzToText(data(i, cSo)))
        res(i, CEN_R_KLASA) = Trim$(NzToText(data(i, cKl)))
        res(i, CEN_R_STANICA) = Trim$(NzToText(data(i, cSt)))
        res(i, CEN_R_DAN) = DanVrednosti(data(i, cDa))
        res(i, CEN_R_CENA) = BrojVrednosti(data(i, cCe))
        res(i, CEN_R_STORNO) = JeStorno(data(i, cSto))
    Next i

    CenovnikRedovi = res
End Function

' Vazeca cena za (stanica, roba, klasa, dan). True + cena kad postoji; False
' kad cene NEMA (CEN-07) -- pozivalac tada ne dira polje. Prazna stanica vidi
' samo globalne cene. Kvar citanja (i prekrsen CEN-03) se loguje i znaci "nema
' cene": predlog cene nikad ne sme da obori unos dokumenta.
Public Function TryVazecaCena(ByVal vrsta As String, ByVal sorta As String, _
                              ByVal klasa As String, ByVal stanicaID As String, _
                              ByVal dan As Date, ByRef cena As Double, _
                              Optional ByRef cenaID As String = "", _
                              Optional ByRef rezerva As Boolean = False) As Boolean
    Dim redovi As Variant, r As Long

    On Error GoTo EH
    cena = 0
    cenaID = ""
    rezerva = False
    If Len(Trim$(vrsta)) = 0 Then Exit Function

    redovi = CenovnikRedovi()
    If Not IsArray(redovi) Then Exit Function

    r = CenVazeciRed(redovi, vrsta, sorta, klasa, stanicaID, dan, rezerva)
    If r = 0 Then
        rezerva = False
        Exit Function
    End If
    If CDbl(redovi(r, CEN_R_CENA)) <= 0 Then
        rezerva = False
        Exit Function
    End If

    cena = CDbl(redovi(r, CEN_R_CENA))
    cenaID = CStr(redovi(r, CEN_R_ID))
    TryVazecaCena = True
    Exit Function

EH:
    LogErr "modCenovnik.TryVazecaCena", Err.description
    cena = 0
    cenaID = ""
    rezerva = False
    TryVazecaCena = False
End Function

' ============================================================
' PROVERE (jedna implementacija, dva pozivaoca: editor i pisac)
' ============================================================

' "" kad red sme da se upise, inace KLJUC poruke (CEN_ERR_*). Pozivalac bira
' polje za fokus po kljucu; pisac isti kljuc dize kao gresku.
'   proveriGranicu = False samo za ispravku koja zadrzava stari Datum (CEN-09).
'   izuzmiID       = red koji ispravka stornira u istoj transakciji: njegov
'                    kljuc ne blokira nov red (storno oslobadja kljuc, CEN-03).
Public Function CenaUnosProblem(ByVal datum As Date, ByVal vrsta As String, _
                                ByVal sorta As String, ByVal klasa As String, _
                                ByVal cena As Double, ByVal stanicaID As String, _
                                Optional ByVal proveriGranicu As Boolean = True, _
                                Optional ByVal izuzmiID As String = "") As String
    Dim dan As Long, danas As Long, redovi As Variant, i As Long

    If Len(Trim$(vrsta)) = 0 Then
        CenaUnosProblem = "CEN_ERR_VRSTA_OBAVEZNA"
        Exit Function
    End If
    If Not KulturaPostoji(vrsta, sorta) Then
        CenaUnosProblem = "CEN_ERR_KULTURA_NEPOZNATA"
        Exit Function
    End If
    If Len(KlasaKanon(klasa)) = 0 Then
        CenaUnosProblem = "CEN_ERR_KLASA"
        Exit Function
    End If
    If cena <= 0 Then
        CenaUnosProblem = "CEN_ERR_CENA_POZITIVNA"
        Exit Function
    End If
    If cena > CEN_MAX Then
        CenaUnosProblem = "CEN_ERR_CENA_OPSEG"
        Exit Function
    End If
    If Not JeUParama(cena) Then
        CenaUnosProblem = "CEN_ERR_CENA_DECIMALE"
        Exit Function
    End If
    If Len(Trim$(stanicaID)) > 0 Then
        If Not StanicaPostoji(stanicaID) Then
            CenaUnosProblem = "CEN_ERR_STANICA_NEPOZNATA"
            Exit Function
        End If
    End If
    If Not IsPoslovnaGodina(datum) Then
        CenaUnosProblem = "CEN_ERR_DATUM"
        Exit Function
    End If

    dan = CLng(Int(CDbl(datum)))
    If proveriGranicu Then
        danas = CLng(Date)
        If dan < danas - CEN_DANA_UNAZAD Or dan > danas + CEN_DANA_UNAPRED Then
            CenaUnosProblem = "CEN_ERR_DATUM_GRANICA"
            Exit Function
        End If
    End If

    redovi = CenovnikRedovi()
    If IsArray(redovi) Then
        For i = 1 To UBound(redovi, 1)
            If Not CBool(redovi(i, CEN_R_STORNO)) Then
            If CLng(redovi(i, CEN_R_DAN)) = dan Then
            If Not IstiTekst(redovi(i, CEN_R_ID), izuzmiID) Or Len(Trim$(izuzmiID)) = 0 Then
                If IstiTekst(redovi(i, CEN_R_VRSTA), vrsta) And _
                   IstiTekst(redovi(i, CEN_R_SORTA), sorta) And _
                   IstiTekst(redovi(i, CEN_R_KLASA), klasa) And _
                   IstiTekst(redovi(i, CEN_R_STANICA), stanicaID) Then
                    CenaUnosProblem = "CEN_ERR_DUPLIKAT"
                    Exit Function
                End If
            End If
            End If
            End If
        Next i
    End If

    CenaUnosProblem = ""
End Function

' "" kad red sme da se stornira, inace kljuc poruke (CEN-11, CEN-12).
Public Function CenaStornoProblem(ByVal cenaID As String, ByVal razlog As String) As String
    Dim redovi As Variant, r As Long

    redovi = CenovnikRedovi()
    r = IndeksPoID(redovi, cenaID)
    If r = 0 Then
        CenaStornoProblem = "CEN_ERR_NEMA_REDA"
        Exit Function
    End If
    If CBool(redovi(r, CEN_R_STORNO)) Then
        CenaStornoProblem = "CEN_ERR_STORNIRANA"
        Exit Function
    End If
    CenaStornoProblem = RazlogProblem(razlog)
End Function

' "" kad ispravka sme da se upise, inace kljuc poruke (CEN-13). Ispravka nosi
' pun nov red. Bez ijedne promene se odbija PRE storna; poredi se po
' vrednosti (cena u parama, tekst bez razlike u velicini slova). Granica datuma
' ne vazi kad ispravka zadrzava stari Datum (CEN-09).
Public Function CenaIspravkaProblem(ByVal staraID As String, ByVal datum As Date, _
                                    ByVal vrsta As String, ByVal sorta As String, _
                                    ByVal klasa As String, ByVal cena As Double, _
                                    ByVal stanicaID As String) As String
    Dim redovi As Variant, r As Long, dan As Long, istiDan As Boolean

    redovi = CenovnikRedovi()
    r = IndeksPoID(redovi, staraID)
    If r = 0 Then
        CenaIspravkaProblem = "CEN_ERR_NEMA_REDA"
        Exit Function
    End If
    If CBool(redovi(r, CEN_R_STORNO)) Then
        CenaIspravkaProblem = "CEN_ERR_STORNIRANA"
        Exit Function
    End If

    dan = CLng(Int(CDbl(datum)))
    istiDan = (dan = CLng(redovi(r, CEN_R_DAN)))

    If istiDan And _
       IstiTekst(redovi(r, CEN_R_VRSTA), vrsta) And _
       IstiTekst(redovi(r, CEN_R_SORTA), sorta) And _
       StrComp(KlasaKanon(CStr(redovi(r, CEN_R_KLASA))), KlasaKanon(klasa), vbBinaryCompare) = 0 And _
       IstiTekst(redovi(r, CEN_R_STANICA), stanicaID) And _
       UPare(CDbl(redovi(r, CEN_R_CENA))) = UPare(cena) Then
        CenaIspravkaProblem = "CEN_ERR_ISPRAVKA_BEZ_PROMENE"
        Exit Function
    End If

    CenaIspravkaProblem = CenaUnosProblem(datum, vrsta, sorta, klasa, cena, stanicaID, _
                                          Not istiDan, staraID)
End Function

' ============================================================
' UPIS (jedini pisac tblCenovnik, WRITE_OWNERSHIP.json)
' ============================================================

' Nov red cene. Vraca CenaID ili "" pri gresci (uzrok je u logu; editor pre
' poziva pusta CenaUnosProblem da operateru kaze sta ne valja). stanicaID
' prazno = globalna cena.
Public Function AddCena(ByVal datum As Date, _
                        ByVal vrstaVoca As String, _
                        ByVal sortaVoca As String, _
                        ByVal klasa As String, _
                        ByVal cena As Double, _
                        Optional ByVal stanicaID As String = "") As String
    On Error GoTo EH
    AddCena = UpisiCenu(datum, vrstaVoca, sortaVoca, klasa, cena, stanicaID, True, "", "")
    Exit Function

EH:
    LogErr "modCenovnik.AddCena", Err.description
    AddCena = ""
End Function

' Storno reda (CEN-11, CEN-12): Stornirano + vreme + autor + razlog, u jednoj
' transakciji. Jednokratan i konacan: storniran red se ne stornira ponovo i
' storno se ne ponistava (nema ga u zurnalu storna). "" = uspeh, inace poruka
' za operatera.
Public Function StornirajCenu(ByVal cenaID As String, ByVal razlog As String) As String
    Const SRC As String = "modCenovnik.StornirajCenu"
    Dim tx As clsTransaction, problem As String, errDesc As String

    On Error GoTo EH
    problem = CenaStornoProblem(cenaID, razlog)
    If Len(problem) > 0 Then
        StornirajCenu = Poruka(problem)
        Exit Function
    End If

    modSchema.SchemaReadyOrFail SRC, TBL_CENOVNIK
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_CENOVNIK

    StornirajRedNaIndeksu IndeksUTabeli(cenaID, SRC), razlog, SRC

    tx.CommitTx
    Set tx = Nothing
    StornirajCenu = ""
    Exit Function

EH:
    errDesc = Err.description
    LogErr SRC, errDesc
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    StornirajCenu = Poruka("CEN_ERR_STORNO_NIJE_UPISAN") & " " & errDesc
End Function

' Ispravka (CEN-13): u JEDNOJ transakciji storno starog reda (oslobadja kljuc),
' nov red sa IspravkaOdID = stari, pa ZamenjenSaID = nov na starom. noviID
' dobija ID novog reda. "" = uspeh, inace poruka za operatera; pri gresci se
' tabela vraca i noviID je prazan.
Public Function IspraviCenu(ByVal staraID As String, ByVal datum As Date, _
                            ByVal vrsta As String, ByVal sorta As String, _
                            ByVal klasa As String, ByVal cena As Double, _
                            ByVal stanicaID As String, ByVal razlog As String, _
                            ByRef noviID As String) As String
    Const SRC As String = "modCenovnik.IspraviCenu"
    Dim tx As clsTransaction, problem As String, errDesc As String
    Dim rStari As Long, redovi As Variant, staroIdx As Long, granica As Boolean

    On Error GoTo EH
    noviID = ""

    problem = CenaStornoProblem(staraID, razlog)
    If Len(problem) = 0 Then
        problem = CenaIspravkaProblem(staraID, datum, vrsta, sorta, klasa, cena, stanicaID)
    End If
    If Len(problem) > 0 Then
        IspraviCenu = Poruka(problem)
        Exit Function
    End If

    redovi = CenovnikRedovi()
    staroIdx = IndeksPoID(redovi, staraID)
    granica = (CLng(Int(CDbl(datum))) <> CLng(redovi(staroIdx, CEN_R_DAN)))

    modSchema.SchemaReadyOrFail SRC, TBL_CENOVNIK
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_CENOVNIK

    rStari = IndeksUTabeli(staraID, SRC)
    StornirajRedNaIndeksu rStari, razlog, SRC
    noviID = UpisiCenu(datum, vrsta, sorta, klasa, cena, stanicaID, granica, staraID, staraID)
    ' AppendRow dodaje na kraj, pa indeks starog reda ostaje isti.
    RequireUpdateCell TBL_CENOVNIK, rStari, COL_TRACE_ZAMENJEN_SA_ID, noviID, SRC

    tx.CommitTx
    Set tx = Nothing
    IspraviCenu = ""
    Exit Function

EH:
    errDesc = Err.description
    LogErr SRC, errDesc
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    noviID = ""
    IspraviCenu = Poruka("CEN_ERR_ISPRAVKA_NIJE_UPISANA") & " " & errDesc
End Function

' MIGRACIJA ZATECENIH DUPLIKATA (CEN-ODL-04). Stari model je dozvoljavao vise
' nestorniranih redova istog kljuca istog dana i birao KASNIJI red tabele.
' Ovde taj red ostaje vazeci, a raniji se storniraju sa razlogom koji imenuje
' pobednika -- isto pravilo koje AgriXSystems F6 uvoz primenjuje. Kljuc je
' (vrsta, sorta, klasa, stanica, dan), bez razlike u velicini slova, kao u
' pravilu. Red bez ispravnog datuma se ne dira (ne ucestvuje u pravilu).
'
' Jednokratno po prirodi i idempotentno: posle prvog prolaza duplikata nema,
' pa drugi prolaz ne upisuje nista. Jedna transakcija; vraca broj storniranih
' redova, -1 pri gresci (tabela vracena).
Public Function MigrirajDupleCene() As Long
    Const SRC As String = "modCenovnik.MigrirajDupleCene"
    Dim redovi As Variant, zadnji As Object, i As Long, k As String, n As Long
    Dim tx As clsTransaction, errDesc As String

    On Error GoTo EH
    MigrirajDupleCene = 0

    redovi = CenovnikRedovi()
    If Not IsArray(redovi) Then Exit Function

    Set zadnji = CreateObject("Scripting.Dictionary")
    zadnji.CompareMode = vbTextCompare
    For i = 1 To UBound(redovi, 1)
        If JeAktivanSaDanom(redovi, i) Then zadnji(KljucDana(redovi, i)) = i
    Next i

    For i = 1 To UBound(redovi, 1)
        If JeAktivanSaDanom(redovi, i) Then
            If CLng(zadnji(KljucDana(redovi, i))) <> i Then n = n + 1
        End If
    Next i
    If n = 0 Then Exit Function

    modSchema.SchemaReadyOrFail SRC, TBL_CENOVNIK
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_CENOVNIK

    For i = 1 To UBound(redovi, 1)
        If JeAktivanSaDanom(redovi, i) Then
            k = KljucDana(redovi, i)
            If CLng(zadnji(k)) <> i Then
                StornirajRedNaIndeksu i, MIG_RAZLOG_PRE & CStr(redovi(CLng(zadnji(k)), CEN_R_ID)) & _
                                         MIG_RAZLOG_POSLE, SRC
            End If
        End If
    Next i

    tx.CommitTx
    Set tx = Nothing
    MigrirajDupleCene = n
    Exit Function

EH:
    errDesc = Err.description
    LogErr SRC, errDesc
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    MigrirajDupleCene = -1
End Function

' ============================================================
' PREGLED PRE UPISA (ADR 0019: "ekran to pokazuje pre upisa")
' ============================================================

' Nova GLOBALNA cena od `datum` prekida raniju POSEBNU cenu. Vraca spisak
' (jedan red po stanici i robi) na kojima posebna cena tog dana prestaje da
' vazi, ili "" kad takvih nema. Racun je suvo izvrsenje ISTOG pravila: pre =
' redovi kakvi jesu, posle = isti redovi + nov globalni red (i bez reda
' izuzmiID, koji ispravka stornira). Bez sorte se gleda i svaka sorta te vrste
' koja bi presla na rezervu.
Public Function PrekinutePosebneCene(ByVal datum As Date, ByVal vrsta As String, _
                                     ByVal sorta As String, ByVal klasa As String, _
                                     Optional ByVal izuzmiID As String = "") As String
    Dim pre As Variant, posle() As Variant, n As Long, i As Long, j As Long
    Dim stanice As Collection, kulture As Collection, st As Variant, ku As Variant
    Dim rPre As Long, rPosle As Long, rez As Boolean, s As String, linija As String

    pre = CenovnikRedovi()
    If Not IsArray(pre) Then Exit Function      ' bez ijednog reda nema ni posebne cene
    n = UBound(pre, 1)

    ReDim posle(1 To n + 1, 1 To CEN_R_KOLONA)
    For i = 1 To n
        For j = 1 To CEN_R_KOLONA
            posle(i, j) = pre(i, j)
        Next j
        If Len(Trim$(izuzmiID)) > 0 Then
            If IstiTekst(posle(i, CEN_R_ID), izuzmiID) Then posle(i, CEN_R_STORNO) = True
        End If
    Next i
    posle(n + 1, CEN_R_ID) = ""
    posle(n + 1, CEN_R_VRSTA) = Trim$(vrsta)
    posle(n + 1, CEN_R_SORTA) = Trim$(sorta)
    posle(n + 1, CEN_R_KLASA) = KlasaKanon(klasa)
    posle(n + 1, CEN_R_STANICA) = ""
    posle(n + 1, CEN_R_DAN) = CLng(Int(CDbl(datum)))
    posle(n + 1, CEN_R_CENA) = 1#
    posle(n + 1, CEN_R_STORNO) = False

    Set stanice = StaniceIDs()
    Set kulture = New Collection
    kulture.Add Array(Trim$(vrsta), Trim$(sorta))
    If Len(Trim$(sorta)) = 0 Then
        For Each ku In SorteVrste(vrsta)
            kulture.Add Array(Trim$(vrsta), CStr(ku))
        Next ku
    End If

    For Each st In stanice
        For Each ku In kulture
            rPre = CenVazeciRed(pre, CStr(ku(0)), CStr(ku(1)), klasa, CStr(st), datum, rez)
            If rPre > 0 Then
                If Len(Trim$(CStr(pre(rPre, CEN_R_STANICA)))) > 0 Then
                    rPosle = CenVazeciRed(posle, CStr(ku(0)), CStr(ku(1)), klasa, CStr(st), datum, rez)
                    If rPosle = n + 1 Then
                        linija = NazivStanice(CStr(st)) & " (" & CStr(st) & "): " & CStr(ku(0))
                        If Len(CStr(ku(1))) > 0 Then linija = linija & " " & CStr(ku(1))
                        linija = linija & " / " & KlasaKanon(klasa) & " -- " & _
                                 Format$(CDbl(pre(rPre, CEN_R_CENA)), "0.00") & " (od " & _
                                 Format$(CDate(CLng(pre(rPre, CEN_R_DAN))), "dd.mm.yyyy") & ")"
                        If Len(s) > 0 Then s = s & vbCrLf
                        s = s & linija
                    End If
                End If
            End If
        Next ku
    Next st

    PrekinutePosebneCene = s
End Function

' ============================================================
' POMOCNE (privatne)
' ============================================================

Private Function UpisiCenu(ByVal datum As Date, ByVal vrsta As String, ByVal sorta As String, _
                           ByVal klasa As String, ByVal cena As Double, ByVal stanicaID As String, _
                           ByVal proveriGranicu As Boolean, ByVal izuzmiID As String, _
                           ByVal ispravkaOdID As String) As String
    Const SRC As String = "modCenovnik.UpisiCenu"
    Dim problem As String, newID As String, rowData As Variant

    problem = CenaUnosProblem(datum, vrsta, sorta, klasa, cena, stanicaID, proveriGranicu, izuzmiID)
    If Len(problem) > 0 Then Err.Raise ERR_CEN_UNOS, SRC, Poruka(problem)

    ' Pozicioni AppendRow: kapija redosleda pre upisa (podaci-i-config.md).
    modSchema.SchemaReadyOrFail SRC, TBL_CENOVNIK
    RequireColumns TBL_CENOVNIK, SRC, _
                   COL_CEN_ID, COL_CEN_DATUM, COL_CEN_VRSTA, COL_CEN_SORTA, _
                   COL_CEN_KLASA, COL_CEN_CENA, COL_CEN_CREATED, COL_STORNIRANO, _
                   COL_AUDIT_CREATED_BY, COL_AUDIT_MODIFIED_AT, COL_AUDIT_MODIFIED_BY, _
                   COL_CEN_STANICA, COL_CEN_STORNIRANO_AT, COL_CEN_STORNIRANO_BY, _
                   COL_CEN_STORNO_RAZLOG, COL_TRACE_ISPRAVKA_OD_ID, COL_TRACE_ZAMENJEN_SA_ID

    newID = GetNextID(TBL_CENOVNIK, COL_CEN_ID, "CEN-")
    If Len(Trim$(newID)) = 0 Then Err.Raise ERR_CEN_ID, SRC, "GetNextID nije vratio CenaID."

    ' Redosled mora pratiti kanon tblCenovnik (schema/schema.json):
    ' CenaID, Datum, VrstaVoca, SortaVoca, Klasa, Cena, CreatedAt, Stornirano,
    ' CreatedBy, ModifiedAt, ModifiedBy (pecat AppendRow-a), StanicaID,
    ' StorniranoAt, StorniranoBy, StornoRazlog, IspravkaOdID, ZamenjenSaID
    rowData = Array( _
        newID, _
        CDate(CLng(Int(CDbl(datum)))), _
        Trim$(vrsta), _
        Trim$(sorta), _
        KlasaKanon(klasa), _
        Round(cena, 2), _
        Now, _
        "", _
        "", "", "", _
        Trim$(stanicaID), _
        "", "", "", _
        Trim$(ispravkaOdID), _
        "")

    If AppendRow(TBL_CENOVNIK, rowData) <= 0 Then
        Err.Raise ERR_CEN_UPIS, SRC, "AppendRow nije uspeo za tblCenovnik."
    End If

    UpisiCenu = newID
End Function

' Drugi zid jednokratnosti (CEN-12): stanje storna se cita iz tabele u istoj
' transakciji, ne iz provere pre nje.
Private Sub StornirajRedNaIndeksu(ByVal r As Long, ByVal razlog As String, ByVal src As String)
    Dim cSto As Long, data As Variant

    cSto = RequireColumnIndex(TBL_CENOVNIK, COL_STORNIRANO, src)
    data = GetTableData(TBL_CENOVNIK)
    If Not IsArray(data) Then Err.Raise ERR_CEN_UPIS, src, "tblCenovnik nema redova."
    If r < 1 Or r > UBound(data, 1) Then Err.Raise ERR_CEN_UPIS, src, "Red cenovnika van opsega: " & r
    If JeStorno(data(r, cSto)) Then Err.Raise ERR_CEN_STORNIRANA, src, Poruka("CEN_ERR_STORNIRANA")

    RequireUpdateCell TBL_CENOVNIK, r, COL_STORNIRANO, STORNO_DA, src
    RequireUpdateCell TBL_CENOVNIK, r, COL_CEN_STORNIRANO_AT, Now, src
    RequireUpdateCell TBL_CENOVNIK, r, COL_CEN_STORNIRANO_BY, AuditUser(), src
    RequireUpdateCell TBL_CENOVNIK, r, COL_CEN_STORNO_RAZLOG, TrimBeline(razlog), src
End Sub

' Indeks reda u tabeli po CenaID-u; tacno jedan pogodak ili greska.
Private Function IndeksUTabeli(ByVal cenaID As String, ByVal src As String) As Long
    RequireTacnoJedan TBL_CENOVNIK, COL_CEN_ID, Trim$(cenaID), "Red cenovnika", src
    IndeksUTabeli = FindRows(TBL_CENOVNIK, COL_CEN_ID, Trim$(cenaID))(1)
End Function

' Indeks u normalizovanim redovima; 0 kad ID nema ili nije jednoznacan.
Private Function IndeksPoID(ByRef redovi As Variant, ByVal cenaID As String) As Long
    Dim i As Long, nadjeno As Long, pogodaka As Long
    If Not IsArray(redovi) Then Exit Function
    If Len(Trim$(cenaID)) = 0 Then Exit Function
    For i = 1 To UBound(redovi, 1)
        If StrComp(CStr(redovi(i, CEN_R_ID)), Trim$(cenaID), vbTextCompare) = 0 Then
            pogodaka = pogodaka + 1
            nadjeno = i
        End If
    Next i
    If pogodaka = 1 Then IndeksPoID = nadjeno
End Function

Private Function RazlogProblem(ByVal razlog As String) As String
    Dim t As String
    t = TrimBeline(razlog)
    If Len(t) = 0 Then
        RazlogProblem = "CEN_ERR_RAZLOG_OBAVEZAN"
    ElseIf Len(t) > CEN_RAZLOG_MAX Then
        RazlogProblem = "CEN_ERR_RAZLOG_DUZINA"
    End If
End Function

' Beline na krajevima: razmak, tab, CR, LF, NBSP (razlog iz InputBox-a ili
' nalepljen tekst). Trim$ skida samo razmak.
Private Function TrimBeline(ByVal s As String) As String
    Dim c As String
    Do While Len(s) > 0
        c = Left$(s, 1)
        If c = " " Or c = vbTab Or c = vbCr Or c = vbLf Or c = ChrW(160) Then
            s = Mid$(s, 2)
        Else
            Exit Do
        End If
    Loop
    Do While Len(s) > 0
        c = Right$(s, 1)
        If c = " " Or c = vbTab Or c = vbCr Or c = vbLf Or c = ChrW(160) Then
            s = Left$(s, Len(s) - 1)
        Else
            Exit Do
        End If
    Loop
    TrimBeline = s
End Function

' Kljuc teksta (DOP-ODL-02): trim, bez razlike u velicini slova.
Private Function IstiTekst(ByVal a As Variant, ByVal b As Variant) As Boolean
    IstiTekst = (StrComp(Trim$(NzToText(a)), Trim$(NzToText(b)), vbTextCompare) = 0)
End Function

' "I" / "II" u kanonskom obliku, "" za sve ostalo.
Private Function KlasaKanon(ByVal klasa As String) As String
    If StrComp(Trim$(klasa), KLASA_I, vbTextCompare) = 0 Then
        KlasaKanon = KLASA_I
    ElseIf StrComp(Trim$(klasa), KLASA_II, vbTextCompare) = 0 Then
        KlasaKanon = KLASA_II
    Else
        KlasaKanon = ""
    End If
End Function

Private Function JeStorno(ByVal v As Variant) As Boolean
    JeStorno = (StrComp(Trim$(NzToText(v)), STORNO_DA, vbTextCompare) = 0)
End Function

' Cena ima najvise dve decimale (pare).
Private Function JeUParama(ByVal c As Double) As Boolean
    JeUParama = (Abs(c * 100# - Int(c * 100# + 0.5)) < 0.000001)
End Function

Private Function UPare(ByVal c As Double) As Double
    UPare = Int(c * 100# + 0.5)
End Function

' Serijski dan iz celije; 0 kad datum nije citljiv ili nije poslovna godina.
Private Function DanVrednosti(ByVal v As Variant) As Long
    Dim d As Date
    On Error GoTo Nema
    Select Case VarType(v)
        Case vbDate
            d = v
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal, vbByte
            If CDbl(v) < 1 Then Exit Function
            d = CDate(CDbl(v))
        Case vbString
            If Not TryParseDateValue(CStr(v), d) Then Exit Function
        Case Else
            Exit Function
    End Select
    If Not IsPoslovnaGodina(d) Then Exit Function
    DanVrednosti = CLng(Int(CDbl(d)))
    Exit Function
Nema:
    DanVrednosti = 0
End Function

Private Function BrojVrednosti(ByVal v As Variant) As Double
    Dim d As Double
    On Error GoTo Nema
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal, vbByte
            BrojVrednosti = CDbl(v)
        Case vbString
            If TryParseDouble(CStr(v), d) Then BrojVrednosti = d
    End Select
    Exit Function
Nema:
    BrojVrednosti = 0
End Function

Private Function JeAktivanSaDanom(ByRef redovi As Variant, ByVal i As Long) As Boolean
    If CBool(redovi(i, CEN_R_STORNO)) Then Exit Function
    JeAktivanSaDanom = (CLng(redovi(i, CEN_R_DAN)) > 0)
End Function

' Kljuc CEN-03: (vrsta, sorta, klasa, stanica, dan). Dictionary je vbTextCompare.
Private Function KljucDana(ByRef redovi As Variant, ByVal i As Long) As String
    KljucDana = Trim$(CStr(redovi(i, CEN_R_VRSTA))) & "|" & _
                Trim$(CStr(redovi(i, CEN_R_SORTA))) & "|" & _
                Trim$(CStr(redovi(i, CEN_R_KLASA))) & "|" & _
                Trim$(CStr(redovi(i, CEN_R_STANICA))) & "|" & _
                CStr(redovi(i, CEN_R_DAN))
End Function

' Vrsta (i sorta, kad je zadata) postoji u sifarniku kultura. Bez sorte je
' dovoljno da vrsta postoji: prazna sorta znaci "kultura bez sorte"
' (DOP-ODL-02), a vrsta ne mora imati poseban red bez sorte.
Private Function KulturaPostoji(ByVal vrsta As String, ByVal sorta As String) As Boolean
    Const SRC As String = "modCenovnik.KulturaPostoji"
    Dim data As Variant, i As Long, cVr As Long, cSo As Long

    RequireTable TBL_KULTURE, SRC
    cVr = RequireColumnIndex(TBL_KULTURE, COL_KUL_VRSTA, SRC)
    cSo = RequireColumnIndex(TBL_KULTURE, COL_KUL_SORTA, SRC)
    data = GetTableData(TBL_KULTURE)
    If Not IsArray(data) Then Exit Function

    For i = 1 To UBound(data, 1)
        If IstiTekst(data(i, cVr), vrsta) Then
            If Len(Trim$(sorta)) = 0 Then
                KulturaPostoji = True
                Exit Function
            End If
            If IstiTekst(data(i, cSo), sorta) Then
                KulturaPostoji = True
                Exit Function
            End If
        End If
    Next i
End Function

' Sorte date vrste iz sifarnika kultura (bez prazne).
Private Function SorteVrste(ByVal vrsta As String) As Collection
    Const SRC As String = "modCenovnik.SorteVrste"
    Dim data As Variant, i As Long, cVr As Long, cSo As Long, s As String
    Dim res As Collection, vidjeno As Object

    Set res = New Collection
    Set SorteVrste = res
    RequireTable TBL_KULTURE, SRC
    cVr = RequireColumnIndex(TBL_KULTURE, COL_KUL_VRSTA, SRC)
    cSo = RequireColumnIndex(TBL_KULTURE, COL_KUL_SORTA, SRC)
    data = GetTableData(TBL_KULTURE)
    If Not IsArray(data) Then Exit Function

    Set vidjeno = CreateObject("Scripting.Dictionary")
    vidjeno.CompareMode = vbTextCompare
    For i = 1 To UBound(data, 1)
        If IstiTekst(data(i, cVr), vrsta) Then
            s = Trim$(NzToText(data(i, cSo)))
            If Len(s) > 0 Then
                If Not vidjeno.Exists(s) Then
                    vidjeno.Add s, True
                    res.Add s
                End If
            End If
        End If
    Next i
End Function

Private Function StanicaPostoji(ByVal stanicaID As String) As Boolean
    RequireTable TBL_STANICE, "modCenovnik.StanicaPostoji"
    StanicaPostoji = (FindRows(TBL_STANICE, COL_STA_ID, Trim$(stanicaID)).count = 1)
End Function

Private Function StaniceIDs() As Collection
    Const SRC As String = "modCenovnik.StaniceIDs"
    Dim data As Variant, i As Long, c As Long, s As String, res As Collection

    Set res = New Collection
    Set StaniceIDs = res
    RequireTable TBL_STANICE, SRC
    c = RequireColumnIndex(TBL_STANICE, COL_STA_ID, SRC)
    data = GetTableData(TBL_STANICE)
    If Not IsArray(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        s = Trim$(NzToText(data(i, c)))
        If Len(s) > 0 Then res.Add s
    Next i
End Function

Private Function NazivStanice(ByVal stanicaID As String) As String
    On Error Resume Next
    NazivStanice = Trim$(NzToText(LookupValue(TBL_STANICE, COL_STA_ID, stanicaID, "Naziv")))
    If Len(NazivStanice) = 0 Then NazivStanice = stanicaID
End Function
