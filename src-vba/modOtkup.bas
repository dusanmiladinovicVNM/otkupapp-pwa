Attribute VB_Name = "modOtkup"
Option Explicit

' ============================================================
' modOtkup - Aufkauf-Geschaeftslogik
' Kernmodul: Erfassung Lieferant zu Station
' ============================================================

' ============================================================
' OTKUP -- header + stavke (skela)
' ============================================================
'
' Jedan javni ulaz koji vraca JEDAN OtkupID. Obrazac je CreateZbirna_TX: _TX
' drzi transakciju i monitoring, Private core radi posao, kompletna
' prevalidacija PRE ijednog upisa.
'
' RAZLIKA U ODNOSU NA ZBIRNU: otkup je PRIMARNA CINJENICA.
'
' Zbirna i otpremnica su izvedeni dokumenti, pa njihovi writeri primaju IZVORE i
' racunaju stavke. Otkup nema izvor ispod sebe -- kolicine nastaju neposrednim
' unosom. Zato prima stavke, i zato nema ni "ocekivano" ni drugi ulaz.
'
' Sta se menja u odnosu na SaveOtkupMulti_TX:
'
'   staro:  dva reda u tblOtkup (Klasa I i Klasa II), dva ID-a, pa string
'           "OTK-1 + OTK-2" koji pozivalac posle parsira na devet mesta
'   novo:   JEDAN header + N stavki, jedan ID
'
' KulturaID SE PRIMA, NE RAZRESAVA.
'
' Zatecen kod ga fabrikuje na dva mesta -- modOtkup.bas:556 i
' modMasterSync.bas:1959 -- tako sto trazi samo po VrstaVoca, a kad ne nadje
' sklopi "vrsta-sorta" string koji izgleda kao FK a ne pokazuje ni na sta.
' Razresavanje (Vrsta, Sorta) -> KulturaID je posao ADAPTERA; writer proverava
' da FK postoji i da se snapshot vrsta/sorta slaze sa tom kulturom.
'
' VISE NIJE SKELA, I VISE NEMA TAKMACA. Od Otkup cutover-a ovo je JEDINI put
' kojim otkup nastaje: ekran (modOtkupUnos), PWA uvoz (modMasterSync) i golden
' mreza zovu bas njega. Stari pisac po klasi je obrisan (korak 3).
'
' Izuzetak je SaveOtkup_TX: nije u pogonu, nego sluzi testovima da naprave
' zaglavlje BEZ STAVKI -- oblik koji nove kapije moraju da odbiju. Obrazlozenje
' stoji uz njega.
'
' Header i dalje ostavlja Kolicina / Cena / Klasa / KolAmbalaze / BrutoKg /
' VozacID / Isplaceno / DatumIsplate / VremeUnosa prazne -- to su kolone koje u
' ciljnoj semi ne postoje (DOCUMENT_HEADER_LINES S4.1) i brisu se u koraku 4.
'
' Transakcija obuhvata tblOtkup, tblOtkupStavke, tblAmbalaza i tblNovac.
' Ambalaza se knjizi jednom po dokumentu; tblNovac dira SAMO primena zatecenog
' avansa -- kes i dalje ne ulazi kroz otkupni list (S4.1b).
'
' Header (h) -- Scripting.Dictionary, obavezni kljucevi:
'   Datum, KooperantID, StanicaID, KulturaID, VrstaVoca, BrojDokumenta
' obavezan KLJUC, vrednost sme biti prazna:
'   SortaVoca     prazna samo ako je i sama kultura bez sorte
'   TipAmbalaze   prazan samo ako nema ni primljene ni izdate ambalaze
' opcioni:
'   ParcelaID, KolAmbIzdata, ClientRecordID, SyncSource, SourceCreatedAt
'
' Stavke -- Collection diktova; spisak kljuceva je ZATVOREN kao i na headeru:
'   Klasa (I ili II), Kolicina (> 0, NETO), Cena (> 0), KolAmbalaze (>= 0),
'   BrutoKg (opciono; > 0 samo kad je unos bio bruto)
'
' Redosled stavki u Collection-u NE odredjuje RedniBroj -- klase se pisu u
' kanonskom redu, isto kao kod zbirne.
Public Function CreateOtkup_TX(ByVal h As Object, _
                               ByVal stavke As Collection, _
                               Optional ByRef outGreska As String) As String
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    outGreska = ""

    On Error GoTo EH

    ' Sema pre upisa: AppendRow pise POZICIONO. Ide PRE BeginTx -- kapija sme da
    ' digne gresku, a nema smisla otvarati transakciju koja se odmah rollback-uje.
    modSchema.SchemaReadyOrFail "CreateOtkup_TX", _
        TBL_OTKUP & "|" & TBL_OTKUP_STAVKE & "|" & TBL_AMBALAZA & "|" & TBL_NOVAC

    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_OTKUP_STAVKE
    ' Ambalaza je u snapshotu zbog pada IZMEDJU dva TrackAmbalaza poziva. Taj pad
    ' se iz javnog API-ja ne moze izazvati (svi ulazi su vec provereni), pa je ovo
    ' NEIZMERENA odbrana -- namerno, i tako imenovana.
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_NOVAC

    CreateOtkup_TX = CreateOtkup(h, stavke)

    If CreateOtkup_TX = "" Then
        Err.Raise vbObjectError + 1860, "CreateOtkup_TX", _
                  "CreateOtkup nije vratio OtkupID."
    End If

    tx.CommitTx

    Set tx = Nothing
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogError "CreateOtkup_TX", errDesc, errNum
    Monitor_Error _
        moduleName:="modOtkup", _
        procedureName:="CreateOtkup_TX", _
        entityType:="Otkup", _
        entityID:=CreateOtkup_TX, _
        correlationId:=CreateOtkup_TX, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="DOKUMENT_SAVE_FAIL", _
        severity:="ERROR", _
        message:="CreateOtkup_TX failed. Error=" & errDesc, _
        userId:="Operator", _
        moduleName:="modOtkup", _
        procedureName:="CreateOtkup_TX", _
        entityType:="Otkup", _
        entityID:=CreateOtkup_TX, _
        correlationId:=CreateOtkup_TX

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    CreateOtkup_TX = ""
    outGreska = errDesc

    PrintOtkupTxFailure "CreateOtkup_TX", errSrc, errNum, errDesc
End Function

' Core -- NE zovi spolja. Jedini ulaz je CreateOtkup_TX, koji drzi snapshot
' transakciju; direktan poziv bi kod greske ostavio header bez stavki.
' ISPRAVKA OTKUPA -- NOV DOKUMENT, NOV BROJ, VEZA PO ID-u (A9).
'
' Ispravka NIKAD ne menja snimljen otkup (A13, S4.1e). Zatecen dokument se
' stornira, nastaje nov, a veza izmedju njih zivi u tabeli:
'
'     OTK-101 / broj 17   Stornirano=Da,  ZamenjenSaID = OTK-202
'           v
'     OTK-202 / broj 18   IspravkaOdID = OTK-101
'     oba nose isti CorrectionID
'
' ZASTO NOV BROJ, a ne isti: broj je labela koju operater vidi NA PAPIRU (A2).
' Dva papira sa istim brojem i razlicitim sadrzajem nisu razlucivi izvan sistema,
' pa je ista-broj varijanta zakljucana kao greska, ne kao opcija.
'
' ZASTO PO ID-u, a ne po broju: modStornoFlow.StampIspravkaTrace to i danas radi
' po broju za ostale tri tabele -- i zato ne moze da razlikuje dve verzije istog
' dokumenta. Otkup je prvi koji prelazi; ostali idu u PR7/PR8.
'
' NOVAC SE NE PRENOSI SAM -- MERENO, i to je NALAZ, ne osobina ovog pisca.
'
' StornoOtkup radi ResetNovacOtkupLink: knjizene isplate se OSLOBADJAJU (OtkupID
' se prazni), ne stornirju. Ocekivano je bilo da ih nov dokument pokupi kroz
' ApplyAvansToOtkup -- ne pokupi ih:
'
'   modNovac:1624   avans-petlja uzima SAMO Tip = NOV_VIRMAN_AVANS_KOOP
'   modNovac:1982   GetKooperantUnallocatedAvans isto
'   modNovac:2031   BuildKooperantUnallocatedAvansDict isto
'
' Odvezana isplata tipa VirmanFirmaKoop zato ostaje NEVIDLJIVA i za dug (nema
' OtkupID) i za avans (pogresan tip). Vidi se jos samo u kartici kooperanta
' (modIzvestaj:2507), gde ulazi u saldo -- pa novac nije izgubljen, ali jeste
' ispao iz svake masinerije koja odlucuje sta se placa.
'
' Ovo NIJE uvedeno ovde: isto radi obican StornoOtkup_TX i radio je oduvek.
' Ispravka ga samo cini lakse dostizivim. Test ga tvrdi kao ZATECENO stanje, da
' se ne bi tumacilo kao osobina; odluka o prenosu je poslovna i ceka operatera.
'
' Opcije, cena svake i preporuka: docs/DOMEN/ODLUKA_NOVAC_PRI_STORNU.md
'
' JEDNA TRANSAKCIJA obuhvata sve: nov dokument, storno starog, obe veze i
' correction context. Delimicna ispravka -- nov dokument bez storna starog, ili
' storno bez naslednika -- gora je od nijedne.
'
' NEMA PRODUKCIONOG POZIVAOCA U OVOM PR-u, i to je merena odluka. Da bi ekran
' zvao ovaj put, stari OtkupID mora da otputuje od prefill-a do pisca; danas
' prefill ide kroz spec string ciji se kljucevi mapiraju NA KONTROLE
' (modOtkupUI.ApplyPrefill), pa ID nema gde. Resenje bi bilo modul-stanje u
' ljusci -- isti oblik in-memory veze koji ovaj korak treba da ukine. Wiring zato
' ide sa PR7, kad se PrefillIzStorniranog ionako prepisuje sa po-klasnih redova.
Public Function IspravkaOtkupa_TX(ByVal stariOtkupID As String, _
                                  ByVal h As Object, _
                                  ByVal stavke As Collection, _
                                  Optional ByRef outGreska As String) As String
    Const SRC As String = "IspravkaOtkupa_TX"

    Dim tx As clsTransaction
    Set tx = New clsTransaction

    outGreska = ""

    On Error GoTo EH

    modSchema.SchemaReadyOrFail SRC, _
        TBL_OTKUP & "|" & TBL_OTKUP_STAVKE & "|" & TBL_AMBALAZA & "|" & TBL_NOVAC

    stariOtkupID = Trim$(stariOtkupID)
    If Len(stariOtkupID) = 0 Then
        Err.Raise vbObjectError + 1910, SRC, "Prazan OtkupID dokumenta koji se ispravlja."
    End If

    ' Izvor mora postojati TACNO JEDNOM i biti AKTIVAN. Ispravka storniranog
    ' dokumenta nije ispravka nego nov unos -- za to postoji CreateOtkup_TX.
    RequireTacnoJedan TBL_OTKUP, COL_OTK_ID, stariOtkupID, "Otkup", SRC

    ' REDOSLED KAPIJA JE MEREN, ne stilski. Vec ispravljen dokument je UVEK i
    ' storniran, pa bi storno-kapija prva uhvatila oba slucaja i operateru rekla
    ' manje korisnu istinu ("storniran") umesto korisnije ("zamenjen sa OTK-x --
    ' ispravi POSLEDNJU verziju"). Specificnija kapija zato ide prva.
    Dim vecZamenjen As String
    vecZamenjen = Trim$(NzToText(LookupValue(TBL_OTKUP, COL_OTK_ID, stariOtkupID, _
                                             COL_TRACE_ZAMENJEN_SA_ID)))
    If Len(vecZamenjen) > 0 Then
        Err.Raise vbObjectError + 1912, SRC, _
                  "Dokument je vec zamenjen dokumentom " & vecZamenjen & _
                  ". Ispravlja se POSLEDNJA verzija, ne istorijska."
    End If

    If UCase$(Trim$(NzToText(LookupValue(TBL_OTKUP, COL_OTK_ID, stariOtkupID, _
                                         COL_STORNIRANO)))) = "DA" Then
        Err.Raise vbObjectError + 1911, SRC, _
                  "Dokument je vec storniran, nema sta da se ispravi: " & stariOtkupID
    End If

    ' A13: OTKUP KOJI IMA RODITELJA SE NE ISPRAVLJA LOKALNO -- NIJEDAN.
    '
    ' Ispravka menja koji dokument postoji. Ako otkup ucestvuje u otpremnici, ta
    ' promena nije lokalna ni u jednom od dva stanja:
    '
    '   IZDATO   otpremnica je izdata sa tim sastavom, pa mora dobiti NOVU
    '            VERZIJU sa novim brojem (H1, GOLDEN_SCENARIJI S12), a za njom i
    '            zbirna. Tiha alternativa je najgora: nov otkup, stara otpremnica
    '            netaknuta, i izdat papir koji vise ne opisuje robu koju nosi.
    '
    '   DRAFT    clanstvo JESTE mutabilno, ali ovaj pisac ga NE dira. Rezultat bi
    '            bio draft ciji izvor pokazuje na STORNIRAN otkup, dok naslednik
    '            stoji van njega. To je isto medjustanje koje je PR5 vec odbio kod
    '            UpdateOtpremnicaDraft_TX: invarijanta mora da vazi IZMEDJU dva
    '            klika, ne tek pri izdavanju. Revalidacija u IzdajOtpremnicu_TX
    '            hvata posledicu prekasno -- posao je do tada vec izgubljen.
    '
    ' Ispravno resenje za DRAFT je ATOMSKA zamena clanstva (ukloni stari izvor,
    ' dodaj naslednika, revalidiraj stanicu/kulturu/ambalazu) u istoj transakciji.
    ' To trazi pisca otpremnice, dakle PR7 -- pa se do tada staje GLASNO za oba.
    Dim roditelj As String
    roditelj = modDokumenta.OtpremnicaZaOtkup(stariOtkupID)
    If Len(roditelj) > 0 Then
        ' Poruka NE nudi "storniraj otpremnicu pa ponovi". To bi bilo uputstvo za
        ' obilazak same kapije: OtpremnicaZaOtkup gleda samo AKTIVNE otpremnice,
        ' pa bi posle storna roditelja lokalna ispravka prosla -- i dala otkup bez
        ' naslednika otpremnice, sto NIJE propagacija nego gubitak lanca.
        Err.Raise vbObjectError + 1918, SRC, _
                  "Otkup ucestvuje u otpremnici " & roditelj & _
                  IIf(modDokumenta.OtpremnicaJeIzdata(roditelj), " (IZDATA)", " (DRAFT)") & _
                  ". Ispravka bi promenila sastav izvedenog dokumenta (A13), a " & _
                  "propagacija na otpremnicu i zbirnu jos ne postoji -- " & _
                  "operacija nije dostupna do PR7."
    End If

    ' A9: nov poslovni broj. Poredi se pre pisca, da poruka imenuje PRAVILO, a ne
    ' jedinstvenost broja -- ista greska sa dva razlicita uzroka zbunjuje operatera.
    Dim stariBroj As String, noviBroj As String
    stariBroj = Trim$(NzToText(LookupValue(TBL_OTKUP, COL_OTK_ID, stariOtkupID, COL_OTK_BR_DOK)))
    noviBroj = Trim$(OtkHdrOpcion(h, "BrojDokumenta"))

    If StrComp(noviBroj, stariBroj, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 1913, SRC, _
                  "Ispravka mora dobiti NOV broj dokumenta (A9). Stari: " & stariBroj & "."
    End If

    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_OTKUP_STAVKE
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_STORNO_VEZE

    ' ZURNAL JE DEO OVE TRANSAKCIJE, ne tudja briga.
    '
    ' StornoOtkup otvara storno operaciju (BeginStornoOp) i pise JournalCell
    ' redove -- zato ih StornoOtkup_TX izricito snapshotuje (modStorno:54). Bez
    ' istog snapshota ovde, pad IZMEDJU storna i uspesnog naslednika ostavlja
    ' zurnal sa zapisom storna koji se posle rollback-a nije desio: undo bi
    ' nudio ponistenje operacije nad dokumentom koji je i dalje aktivan.
    tx.AddTableSnapshot TBL_STORNO_ZURNAL

    ' Redosled je bitan: STORNO PRVI. Nov dokument ide kroz istu kapiju
    ' jedinstvenosti broja kao svaki drugi, a stari jos drzi svoj -- ali kapija
    ' gleda (Stanica, Datum, Broj) i broj je nov, pa sudara nema. Storno prvi ide
    ' zbog novca: ResetNovacOtkupLink oslobodi avans PRE nego sto ga
    ' ApplyAvansToOtkup u novom dokumentu potrazi.
    If Not modStorno.StornoOtkup(stariOtkupID) Then
        Err.Raise vbObjectError + 1914, SRC, _
                  "Storno dokumenta koji se ispravlja nije uspeo: " & stariOtkupID
    End If

    Dim noviID As String
    noviID = CreateOtkup(h, stavke)

    If Len(noviID) = 0 Then
        Err.Raise vbObjectError + 1915, SRC, "CreateOtkup nije vratio OtkupID."
    End If

    ' Correction context: CorrectionID mora biti PRAV ID iz tblStornoVeze, ne broj
    ' koji ovaj modul izmisli. Kolona je deklarisana kao veza na tu tabelu
    ' (modConfig:1031); izmisljen ID bi bio visece pokazivanje.
    Dim cid As String
    cid = modStornoContext.CreateCorrectionContext( _
        SV_MODE_ISPRAVKA, DOK_TIP_OTKUP, stariOtkupID, stariBroj, _
        DOK_TIP_OTKUP, noviID, noviBroj, , , , _
        "Ispravka otkupa: " & stariBroj & " -> " & noviBroj & ".")

    If Len(cid) = 0 Then
        Err.Raise vbObjectError + 1916, SRC, "Correction context nije kreiran."
    End If

    ' Veza na OBA kraja. Redovi se traze ponovo: storno i upis su pomerili tabelu.
    RequireTacnoJedan TBL_OTKUP, COL_OTK_ID, noviID, "Nov otkup", SRC
    RequireTacnoJedan TBL_OTKUP, COL_OTK_ID, stariOtkupID, "Stari otkup", SRC

    Dim rNovi As Long, rStari As Long
    rNovi = FindRows(TBL_OTKUP, COL_OTK_ID, noviID)(1)
    rStari = FindRows(TBL_OTKUP, COL_OTK_ID, stariOtkupID)(1)

    RequireUpdateCell TBL_OTKUP, rNovi, COL_TRACE_ISPRAVKA_OD_ID, stariOtkupID, SRC
    RequireUpdateCell TBL_OTKUP, rNovi, COL_TRACE_CORRECTION_ID, cid, SRC
    RequireUpdateCell TBL_OTKUP, rStari, COL_TRACE_ZAMENJEN_SA_ID, noviID, SRC
    RequireUpdateCell TBL_OTKUP, rStari, COL_TRACE_CORRECTION_ID, cid, SRC

    modStornoContext.CompleteCorrectionContext cid, noviID, noviBroj, _
        "Ispravka otkupa zavrsena: nov dokument " & noviID & "."

    tx.CommitTx
    Set tx = Nothing

    IspravkaOtkupa_TX = noviID
    Exit Function

EH:
    Dim errNum As Long, errDesc As String, errSrc As String
    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogError SRC, errDesc, errNum
    Monitor_Event _
        eventType:="DOKUMENT_SAVE_FAIL", _
        severity:="ERROR", _
        message:="IspravkaOtkupa_TX failed. Stari=" & stariOtkupID & "; Error=" & errDesc, _
        userId:="Operator", _
        moduleName:="modOtkup", _
        procedureName:=SRC, _
        entityType:="Otkup", _
        entityID:=stariOtkupID, _
        correlationId:=stariOtkupID

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    outGreska = errDesc
    IspravkaOtkupa_TX = ""
End Function

' Poslednja ziva verzija dokumenta: prati ZamenjenSaID dok ima kuda.
'
' Bez ovoga bi svaki citalac sam sledio lanac -- a lanac ume da bude dublji od
' jedne karike (ispravka ispravke). Petlja je ogranicena brojem redova: ciklus u
' podacima ne sme da zavrti citaoca.
Public Function PoslednjaVerzijaOtkupa(ByVal otkupID As String) As String
    Const SRC As String = "PoslednjaVerzijaOtkupa"

    otkupID = Trim$(otkupID)
    If Len(otkupID) = 0 Then Exit Function

    Dim maxKoraka As Long
    maxKoraka = OtkBrojRedovaTabele(TBL_OTKUP) + 1

    Dim tekuci As String, sledeci As String, korak As Long
    tekuci = otkupID

    Do
        sledeci = Trim$(NzToText(LookupValue(TBL_OTKUP, COL_OTK_ID, tekuci, _
                                             COL_TRACE_ZAMENJEN_SA_ID)))
        If Len(sledeci) = 0 Then Exit Do
        tekuci = sledeci
        korak = korak + 1
        If korak > maxKoraka Then
            Err.Raise vbObjectError + 1917, SRC, _
                      "Lanac ispravki nema kraj (ciklus?): " & otkupID
        End If
    Loop

    PoslednjaVerzijaOtkupa = tekuci
End Function

Private Function OtkBrojRedovaTabele(ByVal tblName As String) As Long
    Dim d As Variant
    d = GetTableData(tblName)
    If IsArray(d) Then OtkBrojRedovaTabele = UBound(d, 1)
End Function

Private Function CreateOtkup(ByVal h As Object, _
                             ByVal stavke As Collection) As String
    Const SRC As String = "CreateOtkup"

    On Error GoTo EH

    If h Is Nothing Then
        Err.Raise vbObjectError + 1861, SRC, "Header nije prosledjen."
    End If

    If stavke Is Nothing Then
        Err.Raise vbObjectError + 1862, SRC, "Stavke nisu prosledjene."
    End If

    If stavke.count = 0 Then
        Err.Raise vbObjectError + 1863, SRC, _
                  "Otkup mora imati bar jednu stavku."
    End If

    ' Fail-fast nad semom pre ijednog upisa.
    RequireColumnIndex TBL_OTKUP, COL_OTK_ID, SRC
    RequireColumnIndex TBL_OTKUP, COL_OTK_DATUM, SRC
    RequireColumnIndex TBL_OTKUP, COL_OTK_KOOPERANT, SRC
    RequireColumnIndex TBL_OTKUP, COL_OTK_STANICA, SRC
    RequireColumnIndex TBL_OTKUP, COL_OTK_KULTURA, SRC
    RequireColumnIndex TBL_OTKUP, COL_OTK_BR_DOK, SRC

    RequireColumnIndex TBL_OTKUP_STAVKE, COL_OKS_ID, SRC
    RequireColumnIndex TBL_OTKUP_STAVKE, COL_OKS_OTKUP_ID, SRC
    RequireColumnIndex TBL_OTKUP_STAVKE, COL_OKS_RB, SRC
    RequireColumnIndex TBL_OTKUP_STAVKE, COL_OKS_KLASA, SRC
    RequireColumnIndex TBL_OTKUP_STAVKE, COL_OKS_KOLICINA, SRC
    RequireColumnIndex TBL_OTKUP_STAVKE, COL_OKS_CENA, SRC
    RequireColumnIndex TBL_OTKUP_STAVKE, COL_OKS_KOL_AMB, SRC

    OtkHdrProveriKljuceve h, SRC

    Dim datum As Date
    Dim kooperantID As String, stanicaID As String, kulturaID As String
    Dim vrstaVoca As String, sortaVoca As String, tipAmb As String
    Dim brDok As String, parcelaID As String

    datum = OtkHdrDatum(h, "Datum", SRC)
    kooperantID = OtkHdrObavezan(h, "KooperantID", SRC)
    stanicaID = OtkHdrObavezan(h, "StanicaID", SRC)
    kulturaID = OtkHdrObavezan(h, "KulturaID", SRC)
    vrstaVoca = OtkHdrObavezan(h, "VrstaVoca", SRC)
    ' SortaVoca i TipAmbalaze: KLJUC je obavezan (tipfeler ne sme da prodje kao
    ' "nije uneto"), ali vrednost sme da bude prazna. Kultura bez sorte postoji, i
    ' otkup bez ijedne gajbe postoji. Kada sme prazno kaze DOMEN, ne writer:
    ' sortu razresava RequireKulturaSeSlaze (prazno prolazi tacno kad je i kultura
    ' bez sorte), tip ambalaze provera nize (obavezan kad ambalaze ima).
    sortaVoca = OtkHdrObavezanKljuc(h, "SortaVoca", SRC)
    tipAmb = OtkHdrObavezanKljuc(h, "TipAmbalaze", SRC)
    brDok = OtkHdrObavezan(h, "BrojDokumenta", SRC)
    parcelaID = OtkHdrOpcion(h, "ParcelaID")

    ' FK-ovi ka maticnim podacima. "Neprazan string" nije FK: red koji pokazuje
    ' na kooperanta koga nema slomljen je isto koliko i fabrikovan KulturaID,
    ' samo se ne vidi dok neko ne pokusa da ga spoji (DOCUMENT_HEADER_LINES S5).
    '
    ' StanicaID se NE izvodi iz kooperanta. To su dve razlicite cinjenice:
    ' tblKooperanti.StanicaID je maticno otkupno mesto (banka po njemu razvrstava
    ' uplate), a Otkup.StanicaID je mesto GDE JE OTKUP OBAVLJEN -- kod desktopa
    ' iz zakljucane sesije (modStanicaLock.gActiveStanica), pa isti kooperant sme
    ' da preda robu na drugoj stanici (S4.1f).
    RequireTacnoJedan TBL_KOOPERANTI, COL_KOOP_ID, kooperantID, "KooperantID", SRC
    RequireTacnoJedan TBL_STANICE, COL_STA_ID, stanicaID, "StanicaID", SRC
    RequireBrojJedinstven stanicaID, datum, brDok, SRC
    RequireKulturaSeSlaze kulturaID, vrstaVoca, sortaVoca, SRC
    RequireParcelaKooperanta parcelaID, kooperantID, SRC

    ' Prevalidacija SVIH stavki pre bilo kog upisa: dokument sa dve stavke od
    ' kojih druga ne valja ne sme da ostavi prvu u tabeli.
    Dim vidjeneKlase As Object
    Set vidjeneKlase = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim s As Object
    Dim klasa As String
    Dim kolicina As Double, cena As Double, kolAmb As Double, bruto As Double
    Dim imaAmbalaze As Boolean

    For i = 1 To stavke.count
        If Not IsObject(stavke(i)) Then
            Err.Raise vbObjectError + 1864, SRC, _
                      "Stavka " & CStr(i) & " nije Dictionary."
        End If

        Set s = stavke(i)
        OtkStavkaProveriKljuceve s, i, SRC

        klasa = Trim$(NzToText(OtkStavkaVrednost(s, "Klasa", i, SRC)))
        RequireValidOtkupClass klasa, SRC

        ' Dokument ima najvise jednu stavku po klasi -- dve iste klase su bas
        ' bug koji header+stavke uklanja: jedan logicki dokument rasut po redovima.
        If vidjeneKlase.Exists(UCase$(klasa)) Then
            Err.Raise vbObjectError + 1865, SRC, _
                      "Dve stavke iste klase: " & klasa
        End If
        ' Vrednost je INDEKS stavke -- upis nize ide kanonskim redom klasa, pa
        ' mora da zna gde je koja stavka u ulaznom Collection-u.
        vidjeneKlase.Add UCase$(klasa), i

        kolicina = OtkStavkaBroj(s, "Kolicina", i, SRC)
        If kolicina <= 0 Then
            Err.Raise vbObjectError + 1866, SRC, _
                      "Kolicina mora biti veca od nule. Stavka " & CStr(i) & _
                      ", klasa " & klasa & "."
        End If

        ' Cenovnik je PREDLOG; writer trazi samo da cena postoji. Override je
        ' legitiman -- sacuvana cena je istorijska cinjenica dokumenta.
        cena = OtkStavkaBroj(s, "Cena", i, SRC)
        If cena <= 0 Then
            Err.Raise vbObjectError + 1867, SRC, _
                      "Cena mora biti veca od nule. Stavka " & CStr(i) & _
                      ", klasa " & klasa & "."
        End If

        kolAmb = OtkStavkaBroj(s, "KolAmbalaze", i, SRC)
        If kolAmb < 0 Then
            Err.Raise vbObjectError + 1868, SRC, _
                      "Kolicina ambalaze ne sme biti negativna. Stavka " & _
                      CStr(i) & ", klasa " & klasa & "."
        End If
        RequireCeoBrojOtk kolAmb, "Ambalaza na stavci " & CStr(i), SRC
        If kolAmb > 0 Then imaAmbalaze = True

        ' Bruto se cuva SAMO kad je unos bio bruto. Kolicina je uvek neto, pa
        ' bruto koji je manji od nje znaci zamenjene vrednosti, ne rubni slucaj.
        bruto = OtkStavkaBrojOpcion(s, "BrutoKg", i, SRC)
        If bruto < 0 Then
            Err.Raise vbObjectError + 1869, SRC, _
                      "BrutoKg ne sme biti negativan. Stavka " & CStr(i) & "."
        End If
        If bruto > 0 And bruto < kolicina Then
            Err.Raise vbObjectError + 1870, SRC, _
                      "BrutoKg (" & Fmt2Otk(bruto) & ") je manji od neto kolicine (" & _
                      Fmt2Otk(kolicina) & "). Stavka " & CStr(i) & "."
        End If
    Next i

    Dim kolAmbIzdata As Double
    kolAmbIzdata = OtkHdrBrojOpcion(h, "KolAmbIzdata", SRC)
    If kolAmbIzdata < 0 Then
        Err.Raise vbObjectError + 1872, SRC, _
                  "Izdata ambalaza ne sme biti negativna."
    End If
    RequireCeoBrojOtk kolAmbIzdata, "Izdata ambalaza", SRC

    ' Tip ambalaze vezuje SVAKA ambalaza -- i primljena (stavke) i izdata
    ' (header). Isto pravilo drzi zatecen ekran (modOtkupUnos:158), pa nov writer
    ' ne menja poslovno pravilo usput. Zato provera stoji tek ovde: pre nje se ne
    ' zna da li izdate ambalaze ima.
    If (imaAmbalaze Or kolAmbIzdata > 0) And Len(tipAmb) = 0 Then
        Err.Raise vbObjectError + 1871, SRC, _
                  "Tip ambalaze je obavezan kada postoji ambalaza " & _
                  "(primljena na stavkama ili izdata na headeru)."
    End If

    ' --- upis ----------------------------------------------------------------
    Dim otkupID As String
    otkupID = NewEntityID("OTK-")

    If otkupID = "" Then
        Err.Raise vbObjectError + 1873, SRC, _
                  "NewEntityID nije vratio OtkupID."
    End If

    Dim rowData As Variant
    rowData = BuildOtkupHeaderRowData(otkupID, datum, kooperantID, stanicaID, _
                                      kulturaID, vrstaVoca, sortaVoca, tipAmb, _
                                      brDok, parcelaID, kolAmbIzdata, _
                                      OtkHdrOpcion(h, "ClientRecordID"), _
                                      OtkHdrOpcion(h, "SyncSource"), _
                                      OtkHdrOpcion(h, "SourceCreatedAt"))

    If AppendRow(TBL_OTKUP, rowData) <= 0 Then
        Err.Raise vbObjectError + 1874, SRC, _
                  "AppendRow nije upisao header u tblOtkup."
    End If

    ' RedniBroj ide po KANONSKOM redu klasa, ne po redosledu u Collection-u.
    ' Ista poslovna cinjenica (I=400, II=600) mora dati isti dokument bez obzira
    ' na to kojim ih je redom adapter sklopio -- inace RedniBroj nosi trag poziva,
    ' a ne dokumenta. Zbirna to pravilo vec drzi, pa se koristi ista funkcija.
    Dim redKlasa As Collection
    Set redKlasa = modDokumenta.KlaseUKanonskomRedu(vidjeneKlase)

    Dim stavkaID As String
    Dim rb As Long, idx As Long
    For rb = 1 To redKlasa.count
        ' idx je pozicija u ULAZU -- poruke o gresci moraju da imenuju stavku
        ' onako kako ju je pozivalac poslao.
        idx = CLng(vidjeneKlase(CStr(redKlasa(rb))))
        Set s = stavke(idx)

        ' Fail-closed: NewEntityID vraca "" kad CoCreateGuid ne uspe. Red bez
        ' identiteta je gori od pada -- niko ga posle ne moze ni naci ni vezati.
        stavkaID = NewEntityID("OKS-")
        If stavkaID = "" Then
            Err.Raise vbObjectError + 1875, SRC, _
                      "NewEntityID nije vratio OtkupStavkaID za stavku " & CStr(idx) & "."
        End If

        rowData = BuildOtkupStavkaRowData(stavkaID, otkupID, rb, _
                    Trim$(NzToText(s("Klasa"))), _
                    OtkStavkaBroj(s, "Kolicina", idx, SRC), _
                    OtkStavkaBroj(s, "Cena", idx, SRC), _
                    OtkStavkaBroj(s, "KolAmbalaze", idx, SRC), _
                    OtkStavkaBrojOpcion(s, "BrutoKg", idx, SRC))

        If AppendRow(TBL_OTKUP_STAVKE, rowData) <= 0 Then
            Err.Raise vbObjectError + 1876, SRC, _
                      "AppendRow nije upisao stavku " & CStr(idx) & "."
        End If
    Next rb

    KnjiziOtkupAmbalazu otkupID, datum, tipAmb, kooperantID, stanicaID, _
                        ZbirAmbalazeStavki(stavke), kolAmbIzdata, SRC

    ' ZATECEN AVANS SE PRIMENJUJE -- jednom po dokumentu.
    '
    ' Kes NE ulazi kroz otkupni list (S4.1b) i te parametre nov pisac ni nema,
    ' ali avans je druga stvar: to je jedini put koji stvarno postavlja placenost
    ' (S6.1), i zatecen pisac ga primenjuje PO KLASNOM REDU (modOtkup:1043-1044).
    ' Sa jednim headerom se primenjuje jednom, nad vrednoscu celog dokumenta.
    '
    ' Nalaz: prvi prelaz golden mreze na ovog pisca je pao bas ovde --
    ' B2_avans_primenjen i B3_delimican_avans su prijavili placeno 50000 -> 0.
    ApplyAvansToOtkup kooperantID, otkupID

    CreateOtkup = otkupID
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogError SRC, errDesc, errNum
    On Error GoTo 0

    Err.Raise errNum, SRC, "Source=" & errSrc & " | " & errDesc
End Function

' (Vrsta, Sorta) -> TACNO jedan KulturaID. Za ADAPTERE, ne za pisca.
'
' Razresavanje je posao adaptera (S4.1f) -- writer prima gotov FK i samo ga
' proverava. Ali adaptera ima dva, desktop ekran i PWA ingest, pa pravilo zivi
' na jednom mestu; dve implementacije istog razresavanja su vec jednom dale dva
' razlicita ponasanja (modOtkup:556 i modMasterSync:1959 su OBA fabrikovala
' "vrsta-sorta" string kad lookup ne uspe).
'
' Nula i vise od jedan su ISTA greska: izbor se ne prevodi u jedan maticni
' podatak. Tiho uzimanje prvog je bas ono sto je stari kod radio.
Public Function RazresiKulturuIzVrsteSorte(ByVal vrsta As String, _
                                           ByVal sorta As String, _
                                           ByRef outGreska As String) As String
    Const SRC As String = "RazresiKulturuIzVrsteSorte"
    outGreska = ""

    Dim kult As Variant
    kult = GetTableData(TBL_KULTURE)
    If Not IsArray(kult) Then
        outGreska = "Sifarnik kultura je prazan."
        Exit Function
    End If

    Dim cID As Long, cVr As Long, cSo As Long
    cID = RequireColumnIndex(TBL_KULTURE, COL_KUL_ID, SRC)
    cVr = RequireColumnIndex(TBL_KULTURE, COL_KUL_VRSTA, SRC)
    cSo = RequireColumnIndex(TBL_KULTURE, COL_KUL_SORTA, SRC)

    Dim i As Long, nadjen As String, koliko As Long
    For i = 1 To UBound(kult, 1)
        If StrComp(Trim$(NzToText(kult(i, cVr))), Trim$(vrsta), vbTextCompare) = 0 Then
            If StrComp(Trim$(NzToText(kult(i, cSo))), Trim$(sorta), vbTextCompare) = 0 Then
                nadjen = Trim$(NzToText(kult(i, cID)))
                koliko = koliko + 1
            End If
        End If
    Next i

    If koliko <> 1 Then
        outGreska = "(" & vrsta & ", " & sorta & ") se ne prevodi u tacno jednu " & _
                    "kulturu; pogodaka: " & CStr(koliko) & "."
        Exit Function
    End If

    RazresiKulturuIzVrsteSorte = nadjen
End Function

' Otkup koji je vec uvezen pod tim ClientRecordID-em, ili "".
'
' Idempotencija PWA uvoza pociva na ovome: isti CRID sme da stigne vise puta
' (retry, ponovljen sync), ali sme da napravi SAMO JEDAN dokument.
Public Function OtkupPoClientRecordID(ByVal crid As String) As String
    Const SRC As String = "OtkupPoClientRecordID"

    If Len(Trim$(crid)) = 0 Then Exit Function

    Dim d As Variant
    d = GetTableData(TBL_OTKUP)
    If Not IsArray(d) Then Exit Function

    Dim cCrid As Long, cID As Long
    cCrid = RequireColumnIndex(TBL_OTKUP, COL_OTK_CLIENT_RECORD_ID, SRC)
    cID = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, SRC)

    Dim i As Long
    For i = 1 To UBound(d, 1)
        If StrComp(Trim$(NzToText(d(i, cCrid))), Trim$(crid), vbTextCompare) = 0 Then
            OtkupPoClientRecordID = Trim$(NzToText(d(i, cID)))
            Exit Function
        End If
    Next i
End Function

' Broj otkupnog lista je jedinstven po STANICI I DANU -- KROZ CELU ISTORIJU.
'
' Opseg nije izabran nego procitan iz generatora: GenerateBrojDokumenta racuna
' sledeci broj kao MaxSeqFromTable(tblOtkup, BrojDokumenta, Datum, StanicaID)
' (modBrojevi:137) -- dakle StanicaID + Datum + broj.
'
' STORNO NE OSLOBADJA BROJ. Prva verzija ove kapije je radila ExcludeStornirano
' uz obrazlozenje "inace ispravka ne bi mogla da zadrzi isti broj" -- a to je bas
' ono sto A9 zabranjuje: za lanac Otkup -> Otpremnica -> Zbirna ispravka dobija
' NOV BrojDokumenta, da dva papira razlicitog sadrzaja ne bi delila broj.
'
'   OTK120  storniran/zamenjen
'   OTK121  ispravka OTK120        <- nov broj, ne recikliran 120
'
' Kapija zato gleda SVE istorijske redove, ukljucujuci stornirane.
'
' Broj i dalje NIJE identitet (A2) -- ovo je jedinstvenost labele, ne veza.
'
' JEDNA IMPLEMENTACIJA: ekran zove BrojDokumentaZauzet da bi operater dobio
' povratnu informaciju rano, ali pravilo i opseg zive samo ovde. Dve
' implementacije istog invarijanta su se vec razisle -- UI je gledao broj+datum
' bez stanice, pa bi odbio dokument koji je writer smatrao legalnim.
Public Function BrojDokumentaZauzet(ByVal stanicaID As String, ByVal datum As Date, _
                                    ByVal brDok As String) As String
    Const SRC As String = "BrojDokumentaZauzet"

    If Len(Trim$(brDok)) = 0 Then Exit Function

    Dim d As Variant
    d = GetTableData(TBL_OTKUP)
    If Not IsArray(d) Then Exit Function

    Dim cBr As Long, cDat As Long, cSt As Long, cID As Long
    cBr = RequireColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK, SRC)
    cDat = RequireColumnIndex(TBL_OTKUP, COL_OTK_DATUM, SRC)
    cSt = RequireColumnIndex(TBL_OTKUP, COL_OTK_STANICA, SRC)
    cID = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, SRC)

    Dim i As Long
    For i = 1 To UBound(d, 1)
        If StrComp(Trim$(NzToText(d(i, cBr))), Trim$(brDok), vbTextCompare) = 0 Then
            If StrComp(Trim$(NzToText(d(i, cSt))), stanicaID, vbTextCompare) = 0 Then
                If IsDate(d(i, cDat)) Then
                    If Int(CDbl(CDate(d(i, cDat)))) = Int(CDbl(datum)) Then
                        BrojDokumentaZauzet = Trim$(NzToText(d(i, cID)))
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
End Function

Private Sub RequireBrojJedinstven(ByVal stanicaID As String, ByVal datum As Date, _
                                  ByVal brDok As String, ByVal src As String)
    Dim zauzeo As String
    zauzeo = BrojDokumentaZauzet(stanicaID, datum, brDok)

    If Len(zauzeo) > 0 Then
        Err.Raise vbObjectError + 1898, src, _
                  "Broj otkupnog lista " & brDok & " je vec izdat na stanici " & _
                  stanicaID & " tog dana: " & zauzeo & ". Storno ne oslobadja broj " & _
                  "-- ispravka dobija NOV broj (A9)."
    End If
End Sub

' Vrednost dokumenta = SUM(stavke.Kolicina x stavke.Cena).
'
' Cita se SAMO tblOtkupStavke. Nema fallback-a na header -- to bi bio
' compatibility sloj za podatke koje ne cuvamo, a tiho bi vracao 0 tamo gde
' stavki nema umesto da se vidi da dokument nije po novom modelu.
'
' Dokument vise nema JEDNU cenu, pa se vrednost ni ne moze procitati sa headera:
' dve klase legitimno nose dve razlicite cene (S4.1d).
'
' UGOVOR JE PUN (korak 4). Cetiri kapije, sve fail-closed:
'
'   1) prazan OtkupID
'   2) zaglavlje mora postojati TACNO JEDNOM -- ni nula ni dva
'   3) svaka stavka ima numericku Kolicinu i Cenu, obe VECE OD NULE
'      (isto pravilo koje pisac vec trazi na upisu, modOtkup:252/261)
'   4) bar jedna stavka -- dokument BEZ stavki nije dokument vrednosti nula
'
' Zasto (4) nije kozmetika: nula je legitiman odgovor samo kad stavke postoje a
' zbir im je nula. Bez te razlike ApplyAvansToOtkup cita 0 kao "nema sta da se
' plati" i TIHO preskoci primenu avansa -- kvar koji je golden vec jednom
' prijavio (B2/B3).
'
' Ranije je ovaj ugovor bio nepotpun i to je bilo IMENOVANO: kapije su obarale
' 10 do 33 tvrdnje jer su tada jos postojala dva pisca zaglavlja bez stavki
' (stari multi-pisac i PWA uvoz). Oba su zatvorena -- pisac je obrisan (korak 3),
' PWA ide kroz CreateOtkup_TX (korak 2) -- pa kapije vise nemaju sta da obore.
Public Function VrednostOtkupa(ByVal otkupID As String) As Double
    Const SRC As String = "VrednostOtkupa"

    If Len(Trim$(otkupID)) = 0 Then
        Err.Raise vbObjectError + 1901, SRC, "Prazan OtkupID."
    End If

    ' Zaglavlje TACNO JEDNOM. Nula znaci da se racuna vrednost necega sto ne
    ' postoji; dva znace da je OtkupID prestao da bude identitet -- oba su tisi
    ' oblik iste greske od pogresnog zbira.
    Dim hdr As Collection
    Set hdr = FindRows(TBL_OTKUP, COL_OTK_ID, otkupID)

    Dim koliko As Long
    If Not hdr Is Nothing Then koliko = hdr.count

    If koliko <> 1 Then
        Err.Raise vbObjectError + 1902, SRC, _
                  "Zaglavlje otkupa se ne nalazi tacno jednom: " & otkupID & _
                  " (pogodaka: " & CStr(koliko) & ")."
    End If

    Dim d As Variant
    d = GetTableData(TBL_OTKUP_STAVKE)

    Dim cOtk As Long, cKol As Long, cCena As Long
    Dim i As Long, nasao As Long

    If IsArray(d) Then
        cOtk = RequireColumnIndex(TBL_OTKUP_STAVKE, COL_OKS_OTKUP_ID, SRC)
        cKol = RequireColumnIndex(TBL_OTKUP_STAVKE, COL_OKS_KOLICINA, SRC)
        cCena = RequireColumnIndex(TBL_OTKUP_STAVKE, COL_OKS_CENA, SRC)

        For i = 1 To UBound(d, 1)
            If StrComp(Trim$(NzToText(d(i, cOtk))), otkupID, vbTextCompare) = 0 Then
                nasao = nasao + 1

                If Not IsNumeric(d(i, cKol)) Or Not IsNumeric(d(i, cCena)) Then
                    Err.Raise vbObjectError + 1903, SRC, _
                              "Stavka nije brojcana: " & otkupID & _
                              ", stavka " & CStr(nasao) & "."
                End If

                Dim kol As Double, cena As Double
                kol = CDbl(d(i, cKol))
                cena = CDbl(d(i, cCena))

                If kol <= 0 Or cena <= 0 Then
                    Err.Raise vbObjectError + 1904, SRC, _
                              "Kolicina i cena stavke moraju biti vece od nule: " & _
                              otkupID & ", stavka " & CStr(nasao) & "."
                End If

                VrednostOtkupa = VrednostOtkupa + kol * cena
            End If
        Next i
    End If

    If nasao = 0 Then
        Err.Raise vbObjectError + 1899, SRC, _
                  "Otkup nema nijednu stavku: " & otkupID & _
                  ". Vrednost dokumenta se racuna iz tblOtkupStavke."
    End If
End Function

' Ambalaza otkupa -- dvojni upis, isti obrazac kao zatecen pisac.
'
' RAZLIKA: knjizi se JEDNOM po dokumentu, nad ZBIROM stavki, a ne po klasi.
' tblAmbalaza nema kolonu Klasa, pa bi dva reda po klasi bila dva reda koja se
' razlikuju samo u kolicini -- a zbir je isti. Zatecen pisac ih pravi dva samo
' zato sto ima dva OtkupID-a; sa jednim headerom taj razlog nestaje. Storno time
' dobija jedan DokumentID umesto dva.
'
' VOZAC SE NE ZIGOSE. Zatecen kod salje vozacID na kooperantovu nogu iako mu
' sopstveni komentar kaze "otkup nema vozaca na OM-strani" (modOtkup:1294).
' U ciljnom modelu otkup vozaca ni nema -- gajbe idu kooperant -> OM, a vozac
' dolazi tek sa otpremnicom (S4.1c). Posledica je merena: saldo ambalaze po
' vozacu gubi otkupnu nogu (modAmbalaza:499, modIzvestaj:1975, 3011, 4153).
Private Sub KnjiziOtkupAmbalazu(ByVal otkupID As String, ByVal datum As Date, _
                                ByVal tipAmb As String, _
                                ByVal kooperantID As String, _
                                ByVal stanicaID As String, _
                                ByVal primljeno As Double, _
                                ByVal izdato As Double, _
                                ByVal src As String)
    If primljeno > 0 Then
        ' Kooperant predaje pune gajbe na OM:
        '   kooperant IZLAZ (razduzuje se), OM ULAZ (zaduzuje se).
        TrackAmbalaza datum, tipAmb, CLng(primljeno), "Izlaz", _
                      kooperantID, "Kooperant", "", _
                      otkupID, DOK_TIP_OTKUP
        TrackAmbalaza datum, tipAmb, CLng(primljeno), "Ulaz", _
                      stanicaID, "Stanica", "", _
                      otkupID, DOK_TIP_OTKUP
    End If

    If izdato > 0 Then
        ' OM izdaje prazne gajbe kooperantu uz otkup:
        '   kooperant ULAZ (dobija prazne), OM IZLAZ (razduzuje se).
        ' Isti DokumentID -> storno otkupa hvata i ovu nogu (modStorno).
        TrackAmbalaza datum, tipAmb, CLng(izdato), "Ulaz", _
                      kooperantID, "Kooperant", "", _
                      otkupID, DOK_TIP_OM_IZLAZ_KOOP
        TrackAmbalaza datum, tipAmb, CLng(izdato), "Izlaz", _
                      stanicaID, "Stanica", "", _
                      otkupID, DOK_TIP_OM_IZLAZ_KOOP
    End If
End Sub

' Zbir ambalaze svih stavki -- primljene gajbe dokumenta.
Private Function ZbirAmbalazeStavki(ByVal stavke As Collection) As Double
    Dim i As Long
    For i = 1 To stavke.count
        ZbirAmbalazeStavki = ZbirAmbalazeStavki + _
            OtkStavkaBroj(stavke(i), "KolAmbalaze", i, "ZbirAmbalazeStavki")
    Next i
End Function

' KulturaID se NE razresava ovde -- proverava se.
'
' Mora postojati tacno jednom, i vrsta/sorta koje dokument nosi kao snapshot
' moraju odgovarati toj kulturi. Time fabrikovan "vrsta-sorta" string pada odmah:
' takvog reda u tblKulture nema.
Private Sub RequireKulturaSeSlaze(ByVal kulturaID As String, _
                                  ByVal vrstaVoca As String, _
                                  ByVal sortaVoca As String, _
                                  ByVal src As String)
    RequireTacnoJedan TBL_KULTURE, COL_KUL_ID, kulturaID, "KulturaID", src

    Dim kVrsta As String, kSorta As String
    kVrsta = Trim$(NzToText(LookupValue(TBL_KULTURE, COL_KUL_ID, kulturaID, COL_KUL_VRSTA)))
    kSorta = Trim$(NzToText(LookupValue(TBL_KULTURE, COL_KUL_ID, kulturaID, COL_KUL_SORTA)))

    If StrComp(kVrsta, vrstaVoca, vbTextCompare) <> 0 Or _
       StrComp(kSorta, sortaVoca, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 1879, src, _
                  "Vrsta/sorta se ne slazu sa kulturom " & kulturaID & _
                  ": dokument nosi '" & vrstaVoca & "/" & sortaVoca & _
                  "', kultura je '" & kVrsta & "/" & kSorta & "'."
    End If
End Sub

' Parcela mora pripadati kooperantu dokumenta. Tudja parcela ne prolazi
' kanonski writer (DOCUMENT_HEADER_LINES S4.1f).
'
' Neslaganje KULTURE parcele ostaje stvar ekrana (warning uz override) -- za
' tvrdo pravilo tu nema dovoljno osnova.
Private Sub RequireParcelaKooperanta(ByVal parcelaID As String, _
                                     ByVal kooperantID As String, _
                                     ByVal src As String)
    If Len(parcelaID) = 0 Then Exit Sub

    RequireTacnoJedan TBL_PARCELE, COL_PAR_ID, parcelaID, "ParcelaID", src

    Dim vlasnik As String
    vlasnik = Trim$(NzToText(LookupValue(TBL_PARCELE, COL_PAR_ID, parcelaID, COL_PAR_KOOP)))

    If StrComp(vlasnik, kooperantID, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 1882, src, _
                  "Parcela " & parcelaID & " pripada kooperantu " & vlasnik & _
                  ", a otkup je za " & kooperantID & "."
    End If
End Sub

' Header ciljne seme. Kolone koje u ciljnom modelu ne postoje ostaju PRAZNE --
' Kolicina, Cena, Klasa, KolAmbalaze, BrutoKg (stavka), VozacID (otpremnica),
' Isplaceno / DatumIsplate (read-model), VremeUnosa (CreatedAt/SourceCreatedAt),
' Novac / PrimalacNovca, veze po broju i GeneracijaID.
Private Function BuildOtkupHeaderRowData(ByVal otkupID As String, _
                                         ByVal datum As Date, _
                                         ByVal kooperantID As String, _
                                         ByVal stanicaID As String, _
                                         ByVal kulturaID As String, _
                                         ByVal vrstaVoca As String, _
                                         ByVal sortaVoca As String, _
                                         ByVal tipAmb As String, _
                                         ByVal brDok As String, _
                                         ByVal parcelaID As String, _
                                         ByVal kolAmbIzdata As Double, _
                                         ByVal clientRecordID As String, _
                                         ByVal syncSource As String, _
                                         ByVal sourceCreatedAt As String) As Variant
    Const SRC As String = "BuildOtkupHeaderRowData"

    Dim colCount As Long
    colCount = TabelaBrojKolona(TBL_OTKUP)

    If colCount <= 0 Then
        Err.Raise vbObjectError + 1883, SRC, _
                  "Ne mogu da odredim broj kolona za tblOtkup."
    End If

    Dim rowData() As Variant
    ReDim rowData(0 To colCount - 1)

    SetRowValueByColumn rowData, TBL_OTKUP, COL_OTK_ID, otkupID, SRC
    SetRowValueByColumn rowData, TBL_OTKUP, COL_OTK_DATUM, datum, SRC
    SetRowValueByColumn rowData, TBL_OTKUP, COL_OTK_KOOPERANT, kooperantID, SRC
    SetRowValueByColumn rowData, TBL_OTKUP, COL_OTK_STANICA, stanicaID, SRC
    SetRowValueByColumn rowData, TBL_OTKUP, COL_OTK_KULTURA, kulturaID, SRC
    SetRowValueByColumn rowData, TBL_OTKUP, COL_OTK_VRSTA, vrstaVoca, SRC
    SetRowValueByColumn rowData, TBL_OTKUP, COL_OTK_SORTA, sortaVoca, SRC
    SetRowValueByColumn rowData, TBL_OTKUP, COL_OTK_TIP_AMB, tipAmb, SRC
    SetRowValueByColumn rowData, TBL_OTKUP, COL_OTK_BR_DOK, brDok, SRC
    SetRowValueByColumn rowData, TBL_OTKUP, COL_OTK_PARCELA, parcelaID, SRC
    SetRowValueByColumn rowData, TBL_OTKUP, COL_OTK_KOL_AMB_IZDATA, kolAmbIzdata, SRC
    SetRowValueByColumn rowData, TBL_OTKUP, COL_OTK_STORNIRANO, "", SRC

    ' Bez "ako kolona postoji": sve cetiri su KANONSKE, a SchemaReadyOrFail je
    ' vec potvrdio semu pre transakcije. Tiho preskakanje bi od nedostajuce
    ' kanonske kolone napravilo prazno polje umesto pada -- i sakrilo bas drift
    ' zbog kojeg kapija postoji. Nema produkcionih svesaka koje bi to stitilo.
    SetRowValueByColumn rowData, TBL_OTKUP, COL_OTK_CLIENT_RECORD_ID, clientRecordID, SRC
    SetRowValueByColumn rowData, TBL_OTKUP, COL_OTK_SYNC_SOURCE, syncSource, SRC
    SetRowValueByColumn rowData, TBL_OTKUP, COL_OTK_SOURCE_CREATED_AT, sourceCreatedAt, SRC

    ' IZDATO se pise EKSPLICITNO -- nov model se ne oslanja na legacy konvenciju
    ' "prazno = IZDATO". Otkup nema persistentan DRAFT: forma je njegov draft, a
    ' dokument nastaje vec izdat (DOCUMENT_HEADER_LINES S4.1e).
    SetRowValueByColumn rowData, TBL_OTKUP, COL_TRACE_IZDATO_STATUS, IZDATO_IZDATO, SRC

    BuildOtkupHeaderRowData = rowData
End Function

Private Function BuildOtkupStavkaRowData(ByVal stavkaID As String, _
                                         ByVal otkupID As String, _
                                         ByVal redniBroj As Long, _
                                         ByVal klasa As String, _
                                         ByVal kolicina As Double, _
                                         ByVal cena As Double, _
                                         ByVal kolAmb As Double, _
                                         ByVal bruto As Double) As Variant
    Const SRC As String = "BuildOtkupStavkaRowData"

    Dim colCount As Long
    colCount = TabelaBrojKolona(TBL_OTKUP_STAVKE)

    If colCount <= 0 Then
        Err.Raise vbObjectError + 1884, SRC, _
                  "Ne mogu da odredim broj kolona za tblOtkupStavke."
    End If

    Dim rowData() As Variant
    ReDim rowData(0 To colCount - 1)

    SetRowValueByColumn rowData, TBL_OTKUP_STAVKE, COL_OKS_ID, stavkaID, SRC
    SetRowValueByColumn rowData, TBL_OTKUP_STAVKE, COL_OKS_OTKUP_ID, otkupID, SRC
    SetRowValueByColumn rowData, TBL_OTKUP_STAVKE, COL_OKS_RB, redniBroj, SRC
    SetRowValueByColumn rowData, TBL_OTKUP_STAVKE, COL_OKS_KLASA, klasa, SRC
    SetRowValueByColumn rowData, TBL_OTKUP_STAVKE, COL_OKS_KOLICINA, kolicina, SRC
    SetRowValueByColumn rowData, TBL_OTKUP_STAVKE, COL_OKS_CENA, cena, SRC
    SetRowValueByColumn rowData, TBL_OTKUP_STAVKE, COL_OKS_KOL_AMB, kolAmb, SRC

    ' BrutoKg ostaje PRAZAN kad je unos bio neto -- prazno je podatak, ne nula.
    If bruto > 0 Then
        SetRowValueByColumn rowData, TBL_OTKUP_STAVKE, COL_OKS_BRUTO, bruto, SRC
    End If

    BuildOtkupStavkaRowData = rowData
End Function

Private Sub RequireCeoBrojOtk(ByVal v As Double, ByVal opis As String, _
                              ByVal src As String)
    If Abs(v - Fix(v)) > 0.0000001 Then
        Err.Raise vbObjectError + 1886, src, _
                  opis & " mora biti ceo broj, a nije: " & Fmt2Otk(v)
    End If
End Sub

' Broj u poruku, nezavisno od Windows locale-a (decimalna tacka uvek).
Private Function Fmt2Otk(ByVal v As Double) As String
    Fmt2Otk = Replace(Format$(v, "0.00"), ",", ".")
End Function

' --- citanje DTO-a ----------------------------------------------------------
'
' Nedostajuci kljuc je GRESKA, ne prazna vrednost: Dictionary(k) nad nepostojecim
' kljucem tiho vraca Empty i doda kljuc, pa bi tipfeler prosao kao "nije uneto".
'
' Sve sto nije na spisku je greska -- tipfeler u OPCIONOM polju je inace
' nevidljiv. VozacID / Isplaceno / Kolicina i drustvo NISU na spisku namerno:
' pozivalac koji ih salje radi po starom modelu i mora to da cuje.
Private Function OtkHdrKljucPoznat(ByVal kljuc As String) As Boolean
    Select Case LCase$(Trim$(kljuc))
        Case "datum", "kooperantid", "stanicaid", "kulturaid", _
             "vrstavoca", "sortavoca", "tipambalaze", "brojdokumenta", _
             "parcelaid", "kolambizdata", _
             "clientrecordid", "syncsource", "sourcecreatedat"
            OtkHdrKljucPoznat = True
    End Select
End Function

Private Sub OtkHdrProveriKljuceve(ByVal h As Object, ByVal src As String)
    Dim kljuc As Variant

    For Each kljuc In h.Keys
        If Not OtkHdrKljucPoznat(CStr(kljuc)) Then
            Err.Raise vbObjectError + 1887, src, _
                      "Header ima nepoznat kljuc: " & CStr(kljuc) & _
                      ". Kolicina/Cena/Klasa/KolAmbalaze/BrutoKg idu na STAVKU, " & _
                      "VozacID na otpremnicu, Isplaceno je izvedeno."
        End If
    Next kljuc
End Sub

Private Function OtkHdrObavezan(ByVal h As Object, ByVal kljuc As String, _
                                ByVal src As String) As String
    If Not h.Exists(kljuc) Then
        Err.Raise vbObjectError + 1888, src, _
                  "Header nema obavezan kljuc: " & kljuc
    End If

    OtkHdrObavezan = Trim$(NzToText(h(kljuc)))

    If Len(OtkHdrObavezan) = 0 Then
        Err.Raise vbObjectError + 1889, src, _
                  "Header polje je prazno: " & kljuc
    End If
End Function

' Kljuc mora postojati, vrednost sme biti prazna. Razlika prema OtkHdrObavezan
' je namerna: nedostajuci kljuc je uvek greska pozivaoca (tipfeler), a prazna
' vrednost je legitiman podatak za polja koja domen ne trazi uvek.
Private Function OtkHdrObavezanKljuc(ByVal h As Object, ByVal kljuc As String, _
                                     ByVal src As String) As String
    If Not h.Exists(kljuc) Then
        Err.Raise vbObjectError + 1896, src, _
                  "Header nema obavezan kljuc: " & kljuc
    End If

    OtkHdrObavezanKljuc = Trim$(NzToText(h(kljuc)))
End Function

Private Function OtkHdrOpcion(ByVal h As Object, ByVal kljuc As String) As String
    If h.Exists(kljuc) Then OtkHdrOpcion = Trim$(NzToText(h(kljuc)))
End Function

Private Function OtkHdrBrojOpcion(ByVal h As Object, ByVal kljuc As String, _
                                  ByVal src As String) As Double
    If Not h.Exists(kljuc) Then Exit Function

    Dim v As Variant
    v = h(kljuc)
    If IsEmpty(v) Then Exit Function
    If Len(Trim$(NzToText(v))) = 0 Then Exit Function

    If Not IsNumeric(v) Then
        Err.Raise vbObjectError + 1890, src, _
                  "Header polje " & kljuc & " nije broj: " & NzToText(v)
    End If

    OtkHdrBrojOpcion = CDbl(v)
End Function

Private Function OtkHdrDatum(ByVal h As Object, ByVal kljuc As String, _
                             ByVal src As String) As Date
    If Not h.Exists(kljuc) Then
        Err.Raise vbObjectError + 1891, src, _
                  "Header nema obavezan kljuc: " & kljuc
    End If

    If Not IsDate(h(kljuc)) Then
        Err.Raise vbObjectError + 1892, src, _
                  "Header polje nije datum: " & kljuc
    End If

    OtkHdrDatum = CDate(h(kljuc))
End Function

' Stavka ima ZATVOREN spisak kljuceva, iz istog razloga kao header.
'
' Bez ovoga tipfeler u OPCIONOM polju prolazi kao da polja nema: "BruttoKg" se
' ne procita, BrutoKg ostane prazan, i bruto unos se tiho upise kao neto. Ostala
' polja tipfeler prijave sama (nedostajuci OBAVEZAN kljuc pada), pa je bas
' opciono polje jedino mesto gde greska nema svoj glas.
Private Function OtkStavkaKljucPoznat(ByVal kljuc As String) As Boolean
    Select Case LCase$(Trim$(kljuc))
        Case "klasa", "kolicina", "cena", "kolambalaze", "brutokg"
            OtkStavkaKljucPoznat = True
    End Select
End Function

Private Sub OtkStavkaProveriKljuceve(ByVal s As Object, ByVal idx As Long, _
                                     ByVal src As String)
    Dim kljuc As Variant

    For Each kljuc In s.Keys
        If Not OtkStavkaKljucPoznat(CStr(kljuc)) Then
            Err.Raise vbObjectError + 1897, src, _
                      "Stavka " & CStr(idx) & " ima nepoznat kljuc: " & CStr(kljuc) & _
                      ". Dozvoljeni su Klasa / Kolicina / Cena / KolAmbalaze / BrutoKg."
        End If
    Next kljuc
End Sub

Private Function OtkStavkaVrednost(ByVal s As Object, ByVal kljuc As String, _
                                   ByVal idx As Long, ByVal src As String) As Variant
    If Not s.Exists(kljuc) Then
        Err.Raise vbObjectError + 1893, src, _
                  "Stavka " & CStr(idx) & " nema kljuc: " & kljuc
    End If

    OtkStavkaVrednost = s(kljuc)
End Function

Private Function OtkStavkaBroj(ByVal s As Object, ByVal kljuc As String, _
                               ByVal idx As Long, ByVal src As String) As Double
    Dim v As Variant
    v = OtkStavkaVrednost(s, kljuc, idx, src)

    If Not IsNumeric(v) Then
        Err.Raise vbObjectError + 1894, src, _
                  "Stavka " & CStr(idx) & ", polje " & kljuc & _
                  " nije broj: " & NzToText(v)
    End If

    OtkStavkaBroj = CDbl(v)
End Function

' BrutoKg je jedino polje stavke koje sme da izostane -- neto unos ga nema.
Private Function OtkStavkaBrojOpcion(ByVal s As Object, ByVal kljuc As String, _
                                     ByVal idx As Long, ByVal src As String) As Double
    If Not s.Exists(kljuc) Then Exit Function

    Dim v As Variant
    v = s(kljuc)
    If IsEmpty(v) Then Exit Function
    If Len(Trim$(NzToText(v))) = 0 Then Exit Function

    If Not IsNumeric(v) Then
        Err.Raise vbObjectError + 1895, src, _
                  "Stavka " & CStr(idx) & ", polje " & kljuc & _
                  " nije broj: " & NzToText(v)
    End If

    OtkStavkaBrojOpcion = CDbl(v)
End Function

Public Function SaveOtkup_TX(ByVal datum As Date, ByVal kooperantID As String, _
                              ByVal stanicaID As String, ByVal vrstaVoca As String, _
                              ByVal sortaVoca As String, ByVal kolicina As Double, _
                              ByVal cena As Double, ByVal tipAmb As String, _
                              ByVal kolAmb As Long, ByVal vozacID As String, _
                              ByVal brDok As String, ByVal novac As Double, _
                              ByVal primalac As String, _
                              Optional ByVal klasa As String = "I", _
                              Optional ByVal parcelaID As String = "", _
                              Optional ByVal brojZbirne As String = "") As String

    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

        ' Sema pre upisa: AppendRow pise POZICIONO (v. CreateOtkup_TX).
    modSchema.SchemaReadyOrFail "SaveOtkup_TX", _
        TBL_OTKUP & "|" & TBL_AMBALAZA & "|" & TBL_NOVAC

tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_AMBALAZA

    SaveOtkup_TX = SaveOtkup(datum, kooperantID, stanicaID, vrstaVoca, _
                              sortaVoca, kolicina, cena, tipAmb, kolAmb, _
                              vozacID, brDok, novac, primalac, klasa, _
                              parcelaID, brojZbirne)

    If SaveOtkup_TX = "" Then
        Err.Raise vbObjectError + 1801, "SaveOtkup_TX", _
                  "SaveOtkup fehlgeschlagen"
    End If

    tx.CommitTx

    On Error Resume Next
    Monitor_Event _
        eventType:="OTKUP_SAVE_SUCCESS", _
        severity:="INFO", _
        message:="Otkup saved. KooperantID=" & kooperantID & _
                 "; StanicaID=" & stanicaID & _
                 "; Vrsta=" & vrstaVoca & _
                 "; Koli" & ChrW(269) & "ina=" & CStr(kolicina), _
        userId:="Operator", _
        moduleName:="modOtkup", _
        procedureName:="SaveOtkup_TX", _
        entityType:="Otkup", _
        entityID:=SaveOtkup_TX, _
        correlationId:=SaveOtkup_TX
    On Error GoTo 0

    Set tx = Nothing
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    LogErr "SaveOtkup_TX"
    On Error Resume Next
    Monitor_Error _
        moduleName:="modOtkup", _
        procedureName:="SaveOtkup_TX", _
        entityType:="Otkup", _
        entityID:=SaveOtkup_TX, _
        correlationId:=brDok, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="OTKUP_SAVE_FAIL", _
        severity:="ERROR", _
        message:="Otkup save failed. KooperantID=" & kooperantID & _
                 "; StanicaID=" & stanicaID & _
                 "; BrDok=" & brDok & _
                 "; Error=" & errDesc, _
        userId:="Operator", _
        moduleName:="modOtkup", _
        procedureName:="SaveOtkup_TX", _
        entityType:="Otkup", _
        entityID:=SaveOtkup_TX, _
        correlationId:=brDok

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    SaveOtkup_TX = ""

    PrintOtkupTxFailure "SaveOtkup_TX", errSrc, errNum, errDesc
End Function

' SaveOtkupMulti_TX JE OBRISAN (Otkup cutover, korak 3).
'
' Bio je pisac po klasi: jedan unos je davao DVA reda u tblOtkup, vezana samo
' zajednickim BrojDokumenta, i vracao "ID1 + ID2" da bi pozivalac mogao da ih
' razdvoji. To je bila kompenzacija za nedostatak stavki -- tacno ono sto
' CreateOtkup_TX uklanja. Poslednji pozivi su bili fixture-i testova.
'
' SaveOtkup_TX (iznad) NAMERNO ostaje. Nije pisac u pogonu -- niko ga iz UI-ja
' ne zove -- nego JEDINI posten nacin da test napravi zaglavlje BEZ STAVKI, oblik
' koji nove kapije moraju da odbiju (Test_OTK_VrednostBezStavkiPada,
' Test_OTP_StariOtkupNeUlazi). Alternativa bi bila AppendRow iz testa, sto duplira
' znanje o semi i cini test pisacem tabele. Odlazi zajedno sa kolonama u koraku 7.

' ============================================================
' Kontrola proseka neto kg po gajbici (Kolicina / KolAmbalaze).
' Pragovi se citaju iz tblKulture po VrstaVoca (PragProsekUpoz/PragProsekBlok).
' Prazno / 0 -> provera se preskace (opt-in po kulturi; fail-safe i kad kolone
' jos ne postoje na klijentu -- LookupValue vrati Empty -> 0).
'   prosek > PragProsekBlok -> tvrda blokada (False, bez override-a)
'   prosek > PragProsekUpoz -> upozorenje (vbYesNo; False samo ako operater odustane)
' Poziva se iz frmOtkup.btnUnos_Click POSLE bruto->neto konverzije (Kolicina = neto).
' Vraca True kad je unos dozvoljen (ili potvrdjen), False kad treba prekinuti.
' ============================================================
Public Function OtkupProsekGajbiceOK(ByVal vrstaVoca As String, _
        ByVal kolicinaI As Double, ByVal kolAmbI As Long, _
        ByVal kolicinaII As Double, ByVal kolAmbII As Long) As Boolean

    OtkupProsekGajbiceOK = True
    On Error GoTo EH

    Dim pragUpoz As Double, pragBlok As Double
    pragUpoz = KulturaProsekPrag(vrstaVoca, COL_KUL_PRAG_PROSEK_UPOZ)
    pragBlok = KulturaProsekPrag(vrstaVoca, COL_KUL_PRAG_PROSEK_BLOK)

    ' Nijedan prag nije podesen za ovu kulturu -> nema provere.
    If pragUpoz <= 0 And pragBlok <= 0 Then Exit Function

    ' Najveci prosek po klasama (svaka klasa ima svoje gajbe).
    Dim maxProsek As Double, klasaLbl As String
    ProsekKlase kolicinaI, kolAmbI, "I", maxProsek, klasaLbl
    ProsekKlase kolicinaII, kolAmbII, "II", maxProsek, klasaLbl

    If maxProsek <= 0 Then Exit Function

    Dim poruka As String
    poruka = "Prosek po gajbici" & IIf(Len(klasaLbl) > 0, " (klasa " & klasaLbl & ")", "") & _
             " je " & Format$(maxProsek, "0.00") & " kg."

    ' Tvrda blokada.
    If pragBlok > 0 And maxProsek > pragBlok Then
        MsgBox poruka & vbCrLf & _
               "Dozvoljeni maksimum je " & Format$(pragBlok, "0.00") & " kg po gajbici." & vbCrLf & _
               "Unos je blokiran -- proverite neto kila" & ChrW(382) & "u i broj gajbi.", _
               vbCritical, APP_NAME
        OtkupProsekGajbiceOK = False
        Exit Function
    End If

    ' Upozorenje uz mogucnost nastavka.
    If pragUpoz > 0 And maxProsek > pragUpoz Then
        If MsgBox(poruka & vbCrLf & _
                  "Preporu" & ChrW(269) & "eni maksimum je " & Format$(pragUpoz, "0.00") & " kg po gajbici." & vbCrLf & _
                  "Da li ipak " & ChrW(382) & "elite da nastavite?", _
                  vbExclamation + vbYesNo, APP_NAME) = vbNo Then
            OtkupProsekGajbiceOK = False
        End If
    End If
    Exit Function

EH:
    ' Fail-safe: greska u proveri ne sme da obori normalan unos.
    LogErr "modOtkup.OtkupProsekGajbiceOK"
    OtkupProsekGajbiceOK = True
End Function

' Prag proseka za kulturu (po VrstaVoca) iz tblKulture; Empty/ne-broj -> 0.
Private Function KulturaProsekPrag(ByVal vrstaVoca As String, ByVal colName As String) As Double
    Dim v As Variant
    v = LookupValue(TBL_KULTURE, "VrstaVoca", vrstaVoca, colName)
    If IsNumeric(v) Then KulturaProsekPrag = CDbl(v)
End Function

' Prosek jedne klase (neto/gajbe); azurira maxProsek + labelu ako je veci od dosad.
Private Sub ProsekKlase(ByVal kolicina As Double, ByVal kolAmb As Long, _
        ByVal klasa As String, ByRef maxProsek As Double, ByRef klasaLbl As String)
    If kolicina <= 0 Or kolAmb <= 0 Then Exit Sub
    Dim p As Double: p = kolicina / kolAmb
    If p > maxProsek Then
        maxProsek = p
        klasaLbl = klasa
    End If
End Sub

Public Function SaveOtkup(ByVal datum As Date, ByVal kooperantID As String, _
                          ByVal stanicaID As String, ByVal vrstaVoca As String, _
                          ByVal sortaVoca As String, ByVal kolicina As Double, _
                          ByVal cena As Double, ByVal tipAmb As String, _
                          ByVal kolAmb As Long, ByVal vozacID As String, _
                          ByVal brDok As String, ByVal novac As Double, _
                          ByVal primalac As String, _
                          Optional ByVal klasa As String = "I", _
                          Optional ByVal parcelaID As String = "", _
                          Optional ByVal brojZbirne As String = "", _
                          Optional ByVal kolAmbIzdata As Long = 0, _
                          Optional ByVal brutoKg As Double = 0) As String
    On Error GoTo EH

    If Trim$(kooperantID) = "" Then
        Err.Raise vbObjectError + 1820, "SaveOtkup", _
                  "Kooperant mora biti izabran."
    End If

    If Trim$(stanicaID) = "" Then
        Err.Raise vbObjectError + 1821, "SaveOtkup", _
                  "Stanica mora biti izabrana."
    End If

    If Trim$(vrstaVoca) = "" Then
        Err.Raise vbObjectError + 1822, "SaveOtkup", _
                  "Vrsta vo" & ChrW(263) & "a je obavezna."
    End If

    If kolicina <= 0 Then
        Err.Raise vbObjectError + 1823, "SaveOtkup", _
                  "Koli" & ChrW(269) & "ina mora biti veca od nule."
    End If

    If cena <= 0 Then
        Err.Raise vbObjectError + 1824, "SaveOtkup", _
                  "Cena mora biti veca od nule."
    End If

    If kolAmb < 0 Then
        Err.Raise vbObjectError + 1825, "SaveOtkup", _
                  "Koli" & ChrW(269) & "ina ambala" & ChrW(382) & "e ne sme biti negativna."
    End If

    If novac < 0 Then
        Err.Raise vbObjectError + 1826, "SaveOtkup", _
                  "Novac ne sme biti negativan."
    End If

    If kolAmb > 0 And Trim$(tipAmb) = "" Then
        Err.Raise vbObjectError + 1827, "SaveOtkup", _
                  "Tip ambala" & ChrW(382) & "e je obavezan kada postoji ambala" & ChrW(382) & "a."
    End If

    If kolAmbIzdata < 0 Then
        Err.Raise vbObjectError + 1831, "SaveOtkup", _
                  "Koli" & ChrW(269) & "ina izdate ambala" & ChrW(382) & "e ne sme biti negativna."
    End If

    If kolAmbIzdata > 0 And Trim$(tipAmb) = "" Then
        Err.Raise vbObjectError + 1832, "SaveOtkup", _
                  "Tip ambala" & ChrW(382) & "e je obavezan kada postoji izdata ambala" & ChrW(382) & "a."
    End If
    
    Call RequireValidOtkupClass(klasa, "SaveOtkup")

    RequireColumns TBL_OTKUP, "SaveOtkup", _
                   COL_OTK_ID, _
                   COL_OTK_DATUM, _
                   COL_OTK_KOOPERANT, _
                   COL_OTK_STANICA, _
                   COL_OTK_KULTURA, _
                   COL_OTK_VRSTA, _
                   COL_OTK_SORTA, _
                   COL_OTK_KOLICINA, _
                   COL_OTK_CENA, _
                   COL_OTK_TIP_AMB, _
                   COL_OTK_KOL_AMB, _
                   COL_OTK_VOZAC, _
                   COL_OTK_BR_DOK, _
                   COL_OTK_NOVAC, _
                   COL_OTK_PRIMALAC, _
                   COL_OTK_KLASA, _
                   COL_OTK_STORNIRANO, _
                   COL_OTK_BROJ_ZBIRNE, _
                   COL_OTK_OTPREMNICA_ID, _
                   COL_OTK_PARCELA

    Dim newID As String
    newID = GetNextID(TBL_OTKUP, COL_OTK_ID, "OTK-")

    If newID = "" Then
        Err.Raise vbObjectError + 1828, "SaveOtkup", _
                  "GetNextID nije vratio OtkupID."
    End If

    Dim kulturaID As String
    kulturaID = CStr(LookupValue(TBL_KULTURE, "VrstaVoca", vrstaVoca, "KulturaID"))

    If kulturaID = "" Then
        kulturaID = vrstaVoca & "-" & sortaVoca
    End If

    Dim rowData As Variant
    rowData = Array( _
        newID, _
        datum, _
        kooperantID, _
        stanicaID, _
        kulturaID, _
        vrstaVoca, _
        sortaVoca, _
        kolicina, _
        cena, _
        tipAmb, _
        kolAmb, _
        vozacID, _
        brDok, _
        novac, _
        primalac, _
        klasa, _
        "", _
        brojZbirne, _
        "", _
        parcelaID _
    )

    Dim newRow As Long
    newRow = AppendRow(TBL_OTKUP, rowData)
    If newRow <= 0 Then
        Err.Raise vbObjectError + 1829, "SaveOtkup", _
                  "AppendRow fehlgeschlagen fuer tblOtkup."
    End If

    ' Izdata ambalaza (OM->kooperant) -> upis u kolonu PO IMENU (kolona je na kraju
    ' tblOtkup; pozicijski rowData se ne dira). Kolona postoji posle EnsureDoradeSchema.
    If kolAmbIzdata > 0 Then
        UpdateCell TBL_OTKUP, newRow, COL_OTK_KOL_AMB_IZDATA, kolAmbIzdata
    End If

    ' Vreme snimanja otkupa (Now()) -> upis po imenu (kolona na kraju tblOtkup).
    UpdateCell TBL_OTKUP, newRow, COL_OTK_VREME_UNOSA, Now

    ' Bruto tezina (kad je unet bruto pa oduzeta ambalaza) -> upis po imenu; prazno = neto.
    If brutoKg > 0 Then UpdateCell TBL_OTKUP, newRow, COL_OTK_BRUTO, brutoKg

    If kolAmb > 0 Then
        ' Kooperant predaje pune gajbe na OM -> DVOJNI upis (otkup nema vozaca na OM-strani):
        '   1) Kooperant IZLAZ (kooperant se razduzuje),
        '   2) OM/Stanica ULAZ (OM se zaduzuje za isti iznos).
        TrackAmbalaza datum, tipAmb, kolAmb, "Izlaz", _
                      kooperantID, "Kooperant", vozacID, _
                      newID, DOK_TIP_OTKUP
        TrackAmbalaza datum, tipAmb, kolAmb, "Ulaz", _
                      stanicaID, "Stanica", "", _
                      newID, DOK_TIP_OTKUP
    End If

    If kolAmbIzdata > 0 Then
        ' OM izdaje prazne gajbe kooperantu (uz otkup) -> DVOJNI upis (bez vozaca):
        '   1) Kooperant ULAZ (dobija prazne),
        '   2) OM/Stanica IZLAZ (OM se razduzuje).
        ' Isti DokumentID (otkupID) -> storno otkupa hvata i ovu nogu (modStorno).
        TrackAmbalaza datum, tipAmb, kolAmbIzdata, "Ulaz", _
                      kooperantID, "Kooperant", "", _
                      newID, DOK_TIP_OM_IZLAZ_KOOP
        TrackAmbalaza datum, tipAmb, kolAmbIzdata, "Izlaz", _
                      stanicaID, "Stanica", "", _
                      newID, DOK_TIP_OM_IZLAZ_KOOP
    End If

    SaveOtkup = newID
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    LogErr "SaveOtkup"
    On Error Resume Next
    On Error GoTo 0

    Err.Raise errNum, "SaveOtkup", _
              "Source=" & errSrc & " | " & errDesc
End Function

Private Function GetKooperantNazivForNovac(ByVal kooperantID As String) As String
    On Error GoTo EH

    Dim ime As String
    Dim prezime As String

    ime = Trim$(CStr(LookupValue(TBL_KOOPERANTI, COL_KOOP_ID, kooperantID, "Ime")))
    prezime = Trim$(CStr(LookupValue(TBL_KOOPERANTI, COL_KOOP_ID, kooperantID, "Prezime")))

    GetKooperantNazivForNovac = Trim$(ime & " " & prezime)

    If GetKooperantNazivForNovac = "" Then
        GetKooperantNazivForNovac = kooperantID
    End If

    Exit Function

EH:
    LogErr "GetKooperantNazivForNovac"
    GetKooperantNazivForNovac = kooperantID
End Function

Public Function GetOtkupByStation(ByVal stanicaID As String, _
                                  Optional ByVal datumOd As Date = 0, _
                                  Optional ByVal datumDo As Date = 0) As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_OTKUP)

    If IsEmpty(data) Then
        GetOtkupByStation = Empty
        Exit Function
    End If
    
    data = ExcludeStornirano(data, TBL_OTKUP)

    If IsEmpty(data) Then
        GetOtkupByStation = Empty
        Exit Function
    End If

    Dim filters As New Collection
    Dim fp As clsFilterParam

    Set fp = New clsFilterParam
    fp.Init RequireColumnIndex(TBL_OTKUP, COL_OTK_STANICA, _
            "modOtkup.GetOtkupByStation"), "=", stanicaID
    filters.Add fp

    If datumOd > 0 And datumDo > 0 Then
        Set fp = New clsFilterParam
        fp.Init RequireColumnIndex(TBL_OTKUP, COL_OTK_DATUM, _
                "modOtkup.GetOtkupByStation"), "BETWEEN", datumOd, datumDo
        filters.Add fp
    End If

    GetOtkupByStation = FilterArray(data, filters)
    Exit Function

EH:
    LogErr "modOtkup.GetOtkupByStation"
    GetOtkupByStation = Empty
End Function

Public Function GetOtkupByKooperant(ByVal kooperantID As String, _
                                    Optional ByVal datumOd As Date = 0, _
                                    Optional ByVal datumDo As Date = 0) As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_OTKUP)

    If IsEmpty(data) Then
        GetOtkupByKooperant = Empty
        Exit Function
    End If
    
    data = ExcludeStornirano(data, TBL_OTKUP)

    If IsEmpty(data) Then
        GetOtkupByKooperant = Empty
        Exit Function
    End If

    Dim filters As New Collection
    Dim fp As clsFilterParam

    Set fp = New clsFilterParam
    fp.Init RequireColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT, _
            "modOtkup.GetOtkupByKooperant"), "=", kooperantID
    filters.Add fp

    If datumOd > 0 And datumDo > 0 Then
        Set fp = New clsFilterParam
        fp.Init RequireColumnIndex(TBL_OTKUP, COL_OTK_DATUM, _
                "modOtkup.GetOtkupByKooperant"), "BETWEEN", datumOd, datumDo
        filters.Add fp
    End If

    GetOtkupByKooperant = FilterArray(data, filters)
    Exit Function

EH:
    LogErr "modOtkup.GetOtkupByKooperant"
    GetOtkupByKooperant = Empty
End Function
' GetSaldoByStation JE OBRISAN (korak 7).
'
' Sabirao je Kolicina, Novac i KolAmbalaze SA ZAGLAVLJA po kooperantu -- tri
' kolone koje nov pisac ne pise. Da ga je iko zvao, vracao bi nule.
'
' Nije ga zvao niko: grep po celom src-vba daje samo redove unutar same funkcije.
' Mrtav citac mrtve kolone -- brise se, ne prepisuje. Saldo po stanici, kad
' zatreba, ide iz stavki i tblNovac, ne iz zaglavlja.

Private Sub PrintOtkupTxFailure(ByVal sourceName As String, _
                                ByVal errSrc As String, _
                                ByVal errNum As Long, _
                                ByVal errDesc As String)
    Debug.Print sourceName & " failed. Source=" & errSrc & _
                " Err=" & CStr(errNum) & _
                " Desc=" & errDesc
End Sub

Private Sub RequireValidOtkupClass(ByVal klasa As String, _
                                   ByVal sourceName As String)

    Select Case Trim$(CStr(klasa))
        Case KLASA_I, KLASA_II
            Exit Sub
    End Select

    Err.Raise vbObjectError + 1830, sourceName, _
              "Neispravna klasa otkupa: " & klasa
End Sub

