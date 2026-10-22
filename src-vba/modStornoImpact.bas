Attribute VB_Name = "modStornoImpact"
Option Explicit

' ============================================================
' modStornoImpact - agregator UVIDA za Storno centar (Faza 1).
'
' READ-ONLY: sklapa pun model uticaja storna iz postojecih citaca; NE mutira
' nista. Jedan poziv (BuildStornoImpact) daje UI-u sve za "Uvid" ekran:
'   header  - {tip, broj, partnerID, datum, kolicina}
'   chain   - Collection redova [dok, info, napomena]  (modStornoFlow.GetStornoChainRows)
'   blocks  - Collection redova [otkupID, brDok, kg, klasa, koop]  (GetStornoBlockRows)
'   flags   - {hasDependents, canPonistenjeClean, dependentsText}  (GetChainFlags)
'   palete  - Collection dicta po paleti (agregat + detach delta)  (GetPaleteImpactByField)
'   faktura - {hasFaktura, fakturaID}
'   summary - {blockCount, paleteCount, detachGajb, detachNeto, detachAmb}  (traka uticaja)
'
' Reuse: modStornoFlow (Public citaci), modPaletniList.GetPaleteImpactByField,
' modDataAccess.LookupValue. Bez MsgBox, bez TX.
' ============================================================

Private Const MOD_NAME As String = "modStornoImpact"

' Prag ispod koga se ne javlja nista. 400 ms je granica na kojoj operater pocinje
' da primeti cekanje; ispod nje bi linija po kliku bila cist sum u logu.
Private Const UVID_PRAG_MS As Long = 400

' docID = KANONSKI IDENTITET izabranog dokumenta (GeneracijaID za robna
' dokumenta). Do v6-ui-143 ga ovaj sloj nije imao, pa je ceo uvid isao po
' BROJU -- a broj nije jedinstven (GenerateBrojPrijemnice nema proveru
' jedinstvenosti). Posledica nije bila teorijska: pod kolizijom broja je
' pregled pokazivao lanac i blokove TUDJEG dokumenta, a operater bi na osnovu
' toga doneo odluku o stornu. Tri citaca ispod su vec primala docID od
' v6-ui-136/140 -- nedostajao je samo ovaj sloj koji ih spaja.
'
' Prazan docID i dalje prolazi: zatecen zapis bez generacije nema identitet, i
' tada nizvodno vazi fail-closed kapija nad jednoznacnoscu broja.
' "valid" JE DEO UGOVORA. Model se do v6-ui-143 vracao pozivaocu PRE nego sto je
' izgradjen (Set BuildStornoImpact = d na pocetku), pa je pad na pola davao
' PARCIJALAN recnik koji spolja izgleda kao ispravan uvid. Ekran na osnovu njega
' crta posledice i nudi dugmad za mutaciju -- dakle tacno suprotno od svrhe ovog
' ekrana ("prvo vidi posledice, pa odluci"). Sada svaka sekcija mora da se
' izgradi, i tek onda se "valid" postavlja na True; pozivalac koji ga ne proveri
' dobija recnik ciji je valid False.
' strict = "valid znaci NESTO". Kad je True, svih sedam sekcija cita se bez
' gutanja: schema drift, necitljiva tabela i greska u prolazu obaraju ceo uvid
' umesto da vrate prazno. Trazi ga ekran Storno, jer na osnovu ovog modela nudi
' MUTACIJU -- "ne znam da li ima paleta" ne sme da prodje kao "nema paleta".
'
' Podrazumevano False: legacy frmDokumenta zove ovaj isti model za svoj panel,
' ne cita "valid", i njegovo ponasanje se ovom granom ne dira.
Public Function BuildStornoImpact(ByVal docType As String, ByVal broj As String, _
                                  Optional ByVal dokumentTip As String = "", _
                                  Optional ByVal docID As String = "", _
                                  Optional ByVal strict As Boolean = False) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim t0 As Double, tHdr As Double, tChn As Double, tBlk As Double
    Dim tFlg As Double, tPal As Double, tFak As Double, tSum As Double
    d("valid") = False
    d("greska") = ""
    Set BuildStornoImpact = d
    On Error GoTo EH
    broj = Trim$(broj)

    ' MERENJE PO SEKCIJI. Operater je prijavio 2-3 sekunde po kliku na red, a
    ' sedam sekcija cita iste tabele -- bez podatka o tome KOJA trosi vreme svaka
    ' optimizacija je pogadjanje. Zato se meri svaka, a zapisuje se SAMO kad
    ' ukupno predje prag: brz put ostaje tih, spor sam sebe prijavi.
    t0 = Timer
    Set d("header") = ImpactHeader(docType, broj, docID, strict)
    tHdr = Timer
    Set d("chain") = GetStornoChainRows(docType, broj, dokumentTip, docID, strict)
    tChn = Timer
    ' Sekcija otkupnih blokova (GetStornoBlockRows nad Otkup.Kolicina/Klasa i
    ' vezom preko broja) je obrisana u S1b-1 -- stari model; vraca S3.
    Set d("blocks") = New Collection
    tBlk = Timer
    Set d("flags") = GetChainFlags(docType, broj, dokumentTip, docID, strict)
    tFlg = Timer
    Set d("palete") = ImpactPalete(docType, broj, docID, strict)
    tPal = Timer
    Set d("faktura") = ImpactFaktura(docType, broj, docID, strict)
    tFak = Timer
    Set d("summary") = ImpactSummary(d)
    tSum = Timer
    d("valid") = True
    PrijaviTrajanje docType, broj, t0, tHdr, tChn, tBlk, tFlg, tPal, tFak, tSum
    Exit Function
EH:
    d("greska") = Err.description
    LogErr MOD_NAME & ".BuildStornoImpact"
End Function

' Trajanje po sekciji, u JEDNOJ liniji. Sedam linija po kliku bi log pretvorilo u
' buku, a poredjenje sekcija je bas ono sto se trazi -- zato jedan red.
'
' Timer vraca sekunde od ponoci kao Single, sa rezolucijom oko 1/64 s na Windows-u.
' To je grubo za jednu sekciju od 5 ms, ali sasvim dovoljno da se od tri sekunde
' vidi ko ih trosi -- a to je jedino pitanje.
Private Sub PrijaviTrajanje(ByVal docType As String, ByVal broj As String, _
                            ByVal t0 As Double, ByVal tHdr As Double, ByVal tChn As Double, _
                            ByVal tBlk As Double, ByVal tFlg As Double, ByVal tPal As Double, _
                            ByVal tFak As Double, ByVal tSum As Double)
    Dim ukupno As Double
    On Error Resume Next
    ukupno = (tSum - t0) * 1000
    If ukupno < UVID_PRAG_MS Then Exit Sub
    LogWarn MOD_NAME & ".BuildStornoImpact", _
            "[" & modOtkupUI.OTKUI_BUILD & "] Uvid " & docType & " " & broj & _
            " trajao " & Format$(ukupno, "0") & " ms: " & _
            "zaglavlje " & Format$((tHdr - t0) * 1000, "0") & _
            ", lanac " & Format$((tChn - tHdr) * 1000, "0") & _
            ", blokovi " & Format$((tBlk - tChn) * 1000, "0") & _
            ", zastavice " & Format$((tFlg - tBlk) * 1000, "0") & _
            ", palete " & Format$((tPal - tFlg) * 1000, "0") & _
            ", faktura " & Format$((tFak - tPal) * 1000, "0") & _
            ", zbir " & Format$((tSum - tFak) * 1000, "0") & " ms."
End Sub

' ------------------------------------------------------------
' Header po tipu (samo potvrdjene kolone; partner ostaje ID -> UI resolvira ime).
' ------------------------------------------------------------
Private Function ImpactHeader(ByVal docType As String, ByVal broj As String, _
                              Optional ByVal docID As String = "", _
                              Optional ByVal strict As Boolean = False) As Object
    Dim h As Object: Set h = CreateObject("Scripting.Dictionary")
    Set ImpactHeader = h
    On Error GoTo EH
    h("tip") = docType: h("broj") = broj
    h("partnerID") = "": h("partner") = "": h("datum") = "": h("kolicina") = ""
    h("ispravkaOd") = "": h("zamenjenSa") = ""
    Dim tTbl As String, tCol As String: tTbl = "": tCol = ""
    Select Case docType
        ' kolicina = SUMA aktivnih redova broja (Klasa I + II dele broj; HL bi vratio
        ' samo prvu klasu -> potceni dvoklasni dokument u uvidu). partner/datum su
        ' isti po klasama pa im je dovoljan jedan red -- ali BAS njegov, pa idu
        ' kroz HLI, koji uz broj postuje i identitet.
        Case FLOW_DOC_ZBIRNA
            tTbl = TBL_ZBIRNA: tCol = COL_ZBR_BROJ
        Case FLOW_DOC_PRIJEMNICA
            tTbl = TBL_PRIJEMNICA: tCol = COL_PRJ_BROJ
    End Select
    Select Case docType
        Case FLOW_DOC_ZBIRNA
            ' Identitet zbirne je ZbirnaID (S4-2), pa se red bira njime -- ne
            ' parom (broj, generacija). Kanonski dokument generaciju nema.
            '
            ' ID se razresava JEDNOM, na ulazu: zadat se proverava (postoji i
            ' nosi BAS taj broj), a kad ga nema -- zatecen poziv po broju --
            ' izvodi se, i to samo ako je broj jednoznacan.
            Dim zbrUvid As String
            zbrUvid = ZbrIdUvid(broj, docID, strict)
            h("partnerID") = HLZ(COL_ZBR_KUPAC, zbrUvid, strict)
            h("datum") = HLZ(COL_ZBR_DATUM, zbrUvid, strict)
            h("kolicina") = ZbirnaKgZaUvid(zbrUvid, strict)
        Case FLOW_DOC_PRIJEMNICA
            ' Identitet je PrijemnicaID (S6 korak 4), razresen JEDNOM istim
            ' razresavacem koji koristi okvir; kg dokumenta je zbir STAVKI (R5) --
            ' zaglavlje novog oblika kolicinu ne nosi.
            Dim prjUvid As String
            prjUvid = modStornoFlow.PrijemnicaIDZaStorno(broj, docID, strict)
            h("partnerID") = HLP(COL_PRJ_KUPAC, prjUvid)
            h("datum") = HLP(COL_PRJ_DATUM, prjUvid)
            h("kolicina") = PrijemnicaKgZaUvid(prjUvid, strict)
    End Select
    ' Razresi ID -> naziv (otpremnica = stanica; zbirna/prijemnica = kupac). Fallback ID.
    h("partner") = ResolvePartnerName(docType, CStr(h("partnerID")))
    ' Sledljivost (Faza 7): da li je ovaj dokument ispravka drugog / zamenjen drugim.
    If Len(tTbl) > 0 Then
        If docType = FLOW_DOC_ZBIRNA Then
            h("ispravkaOd") = HLZ(COL_TRACE_ISPRAVKA_OD, zbrUvid, strict)
            h("zamenjenSa") = HLZ(COL_TRACE_ZAMENJEN_SA, zbrUvid, strict)
        ElseIf docType = FLOW_DOC_PRIJEMNICA Then
            h("ispravkaOd") = HLP(COL_TRACE_ISPRAVKA_OD, prjUvid)
            h("zamenjenSa") = HLP(COL_TRACE_ZAMENJEN_SA, prjUvid)
        Else
            h("ispravkaOd") = HLI(tTbl, tCol, broj, COL_TRACE_ISPRAVKA_OD, docID, strict)
            h("zamenjenSa") = HLI(tTbl, tCol, broj, COL_TRACE_ZAMENJEN_SA, docID, strict)
        End If
    End If
    Exit Function
EH:
    ' Greska se PROPUSTA dalje, ne guta. Ovaj sloj je deo modela koji se posle
    ' oznacava kao valid -- progutana greska bi dala prazno polje koje spolja
    ' izgleda kao "podatka nema", a znaci "ne znam".
    Dim errNum As Long, errDesc As String
    errNum = Err.Number: errDesc = Err.description
    LogErr MOD_NAME & ".ImpactHeader"
    If strict Then Err.Raise errNum, MOD_NAME & ".ImpactHeader", errDesc
End Function

' ID partnera -> citljiv naziv. Otpremnica: tblStanice (StanicaID -> Naziv);
' zbirna/prijemnica: tblKupci (KupacID -> Naziv). Ako naziv nema -> vrati ID.
Private Function ResolvePartnerName(ByVal docType As String, ByVal partnerID As String) As String
    On Error Resume Next
    ResolvePartnerName = partnerID
    If Len(Trim$(partnerID)) = 0 Then Exit Function
    Dim nm As String
    Select Case docType
        Case FLOW_DOC_ZBIRNA, FLOW_DOC_PRIJEMNICA
            nm = Trim$(CStr(LookupValue(TBL_KUPCI, COL_KUP_ID, partnerID, COL_KUP_NAZIV)))
    End Select
    If Len(nm) > 0 Then ResolvePartnerName = nm
End Function

' ------------------------------------------------------------
' Palete po tipu, SUZENO NA IZABRAN DOKUMENT gde god sema to dozvoljava.
'
' Do v6-ui-143 je ovde stajao goli broj, i to je bila prava rupa: zaglavlje,
' lanac i blokovi su isli po identitetu, a palete i faktura po BROJU -- pa je
' pod kolizijom broja uvid pokazivao palete OBA dokumenta, dok writer nizvodno
' mutira samo jedan. Ekran je time tvrdio posledice koje ne odgovaraju radnji.
'
' Dokle sema dozvoljava suzavanje:
'   PRIJEMNICA  tblPaletaStavka nosi PrijemnicaID -> puno suzavanje. Jedan
'               logicki dokument ima VISE redova (Klasa I i II dele generaciju,
'               a imaju razlicit PrijemnicaID), pa se salje SKUP ID-jeva.
'   OTPREMNICA  stavke ne nose otpremnicu nego BrojZbirne; identitet se ipak
'               koristi da se nadje zbirna BAS te otpremnice (HLI, ne HL).
'   ZBIRNA      stavke nose BrojZbirne, ne ZbirnaID -- suzavanje po identitetu
'               nije moguce. Ista granica sheme kao kod FLOW_DOC_ZBIRNA u
'               ActiveBlocksForFlow (tblOtkup nosi denormalizovan BrojZbirne).
'               Prijavljuje se kao granica, ne krpi se pogadjanjem.
' ------------------------------------------------------------
Private Function ImpactPalete(ByVal docType As String, ByVal broj As String, _
                              Optional ByVal docID As String = "", _
                              Optional ByVal strict As Boolean = False) As Collection
    Dim ids As Object, bz As String
    On Error GoTo EH
    Select Case docType
        Case FLOW_DOC_PRIJEMNICA
            Set ids = PrijemniceIDPoIdentitetu(broj, docID)
            If ids Is Nothing Then
                ' IDENTITET SE NE ODUSTAJE. Kad je docID zadat a ne moze da se
                ' razresi (schema drift, nema reda pod tom generacijom), povratak
                ' na broj bi vratio tacno ono sto je #198 vadio -- i to unutar
                ' modela koji se posle oznacava kao valid. Prazan docID je druga
                ' prica: zatecen zapis nema identitet, pa je broj sve sto postoji.
                If strict And Len(Trim$(docID)) > 0 Then
                    Err.Raise ERR_UI_BASE + 26, MOD_NAME & ".ImpactPalete", _
                              "Identitet prijemnice se ne moze razresiti -- palete se ne mogu suziti."
                End If
                Set ImpactPalete = GetPaleteImpactByField(COL_PALS_BROJ_PRIJ, broj, Nothing, strict)
            Else
                Set ImpactPalete = GetPaleteImpactByField(COL_PALS_PRIJEMNICA_ID, "", ids, strict)
            End If
        Case FLOW_DOC_ZBIRNA
            Set ImpactPalete = GetPaleteImpactByField(COL_PALS_BROJ_ZBIRNE, broj, Nothing, strict)
        Case Else
            Set ImpactPalete = New Collection
    End Select
    Exit Function
EH:
    ' Greska se PODIZE, ne guta. Prazna kolekcija je legitiman odgovor (dokument
    ' nema palete) i ne sme da znaci isto sto i neuspelo citanje -- inace bi uvid
    ' tvrdio da posledica nema, a operater bi na osnovu toga stornirao.
    LogErr MOD_NAME & ".ImpactPalete"
    If strict Then
        Err.Raise ERR_UI_BASE + 23, MOD_NAME & ".ImpactPalete", _
                  "Palete izabranog dokumenta se ne mogu procitati."
    End If
    Set ImpactPalete = New Collection
End Function

' PrijemnicaID IZABRANOG dokumenta kao skup od jednog clana (S6 korak 4: identitet
' je PrijemnicaID, ne generacija). Nothing = identitet nije zadat ili se ne
' razresava -> pozivalac se vraca na broj (samo van strict-a, v. ImpactPalete).
' Prazan skup NIJE "nema paleta" nego "identitet nije nadjen".
Private Function PrijemniceIDPoIdentitetu(ByVal broj As String, _
                                          ByVal docID As String) As Object
    Dim pid As String
    pid = modStornoFlow.PrijemnicaIDZaStorno(broj, docID, False)
    If Len(pid) = 0 Then Exit Function
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.Add pid, 1
    Set PrijemniceIDPoIdentitetu = d
End Function

' Polje zaglavlja prijemnice po PrijemnicaID-u (razresen gore). Prazan ID daje
' prazno -- uvid tada ne opisuje tudji dokument istog broja.
Private Function HLP(ByVal valCol As String, ByVal prijemnicaID As String) As String
    If Len(Trim$(prijemnicaID)) = 0 Then Exit Function
    HLP = Trim$(NzToText(LookupValue(TBL_PRIJEMNICA, COL_PRJ_ID, Trim$(prijemnicaID), valCol)))
End Function

' Kg prijemnice u uvidu je zbir STAVKI (R5). Strict: strogi citalac (dokument bez
' stavki obara uvid po imenu); van strict-a meko, jer uvid mora nesto da pokaze.
Private Function PrijemnicaKgZaUvid(ByVal prijemnicaID As String, _
                                    ByVal strict As Boolean) As String
    Const SRC As String = "modStornoImpact.PrijemnicaKgZaUvid"
    If Len(Trim$(prijemnicaID)) = 0 Then Exit Function
    Dim kg As Variant
    If strict Then
        kg = modDokumenta.ZbirStavkiZaPrijemnicu(modDokumenta.ZbirStavkiPoPrijemnici(), _
                                                 prijemnicaID, SRC)(0)
    Else
        kg = modDokumenta.KgPrijemniceMeko(modDokumenta.ZbirStavkiPoPrijemniciMeko(), prijemnicaID)
        If Not IsNumeric(kg) Then Exit Function
    End If
    ' Isti formater kao lista za storno (KgTekst), da uvid i lista ne pokazu
    ' istu kolicinu na dva nacina.
    PrijemnicaKgZaUvid = modStornoDok.KgTekst(CDbl(kg))
End Function

' Faktura (samo prijemnica ima direktan link u ovom sloju), po IDENTITETU.
' HL vraca PRVI red po broju -- pod kolizijom je uvid umeo da prijavi fakturu
' tudje prijemnice, pa i da tvrdi "fakturisano" za dokument koji to nije.
Private Function ImpactFaktura(ByVal docType As String, ByVal broj As String, _
                               Optional ByVal docID As String = "", _
                               Optional ByVal strict As Boolean = False) As Object
    Dim f As Object: Set f = CreateObject("Scripting.Dictionary")
    Set ImpactFaktura = f
    f("hasFaktura") = False: f("fakturaID") = ""
    On Error GoTo EH
    If docType = FLOW_DOC_PRIJEMNICA Then
        ' ISTI predikat kao kapija storna (S6 korak 4, odluka 1): izveden iz
        ' aktivnih linija fakture (S6 korak 5, FAK-LIN-03). Uvid koji kaze "nije
        ' fakturisana" nad dokumentom koji kapija odbija bio bi gori od uvida koga nema.
        Dim prjF As String, fakF As String
        prjF = modStornoFlow.PrijemnicaIDZaStorno(broj, docID, strict)
        If Len(prjF) > 0 Then
            f("hasFaktura") = modFaktura.PrijemnicaJeFakturisana(prjF, fakF)
            f("fakturaID") = fakF
        End If
    End If
    Exit Function
EH:
    ' Greska se PROPUSTA dalje, ne guta. Ovaj sloj je deo modela koji se posle
    ' oznacava kao valid -- progutana greska bi dala prazno polje koje spolja
    ' izgleda kao "podatka nema", a znaci "ne znam".
    Dim errNum As Long, errDesc As String
    errNum = Err.Number: errDesc = Err.description
    LogErr MOD_NAME & ".ImpactFaktura"
    If strict Then Err.Raise errNum, MOD_NAME & ".ImpactFaktura", errDesc
End Function

' ------------------------------------------------------------
' Summary (traka uticaja): brojevi + detach delta zbir preko svih paleta.
' detach* = koliko bi DUPLI/PONISTENJE skinulo (thisGajb/thisNeto/thisAmb).
' ------------------------------------------------------------
Private Function ImpactSummary(ByVal d As Object) As Object
    Dim sm As Object: Set sm = CreateObject("Scripting.Dictionary")
    Set ImpactSummary = sm
    sm("blockCount") = 0: sm("paleteCount") = 0
    sm("detachGajb") = 0&: sm("detachNeto") = 0#: sm("detachAmb") = 0#
    On Error Resume Next

    Dim blocks As Collection: Set blocks = d("blocks")
    If Not blocks Is Nothing Then sm("blockCount") = blocks.count

    Dim palete As Collection: Set palete = d("palete")
    If Not palete Is Nothing Then
        sm("paleteCount") = palete.count
        Dim pg As Long, pk As Double, pa As Double, i As Long
        For i = 1 To palete.count
            pg = pg + CLng(palete(i)("thisGajb"))
            pk = pk + CDbl(palete(i)("thisNeto"))
            pa = pa + CDbl(palete(i)("thisAmb"))
        Next i
        sm("detachGajb") = pg: sm("detachNeto") = pk: sm("detachAmb") = pa
    End If
End Function

' Safe lookup -> Trim$ string ("" na gresku/prazno).
Private Function HL(ByVal tbl As String, ByVal keyCol As String, _
                    ByVal keyVal As String, ByVal valCol As String) As String
    On Error Resume Next
    HL = Trim$(CStr(LookupValue(tbl, keyCol, keyVal, valCol)))
End Function

' Isto, ali uz IDENTITET. Prazan docID -> obican lookup po broju (zatecen zapis
' bez generacije). Neprazan docID -> red koji ima i taj broj i tu generaciju;
' ako ga nema, vraca prazno umesto tudje vrednosti.
'
' LookupValue vraca PRVI red po broju, pa je pod kolizijom umeo da prikaze
' partnera i datum drugog dokumenta -- zaglavlje uvida bi opisivalo dokument
' koji se ne stornira.
Private Function HLI(ByVal tbl As String, ByVal keyCol As String, _
                     ByVal keyVal As String, ByVal valCol As String, _
                     ByVal docID As String, _
                     Optional ByVal strict As Boolean = False) As String
    If Len(Trim$(docID)) = 0 Then
        ' Bez identiteta se cita po broju -- ali u strict rezimu i to mora da
        ' bude POUZDANO citanje. HL ima svoj "On Error Resume Next", pa bi
        ' schema drift prosao kao "vrednost je prazna" i uvid bi ostao valid.
        If strict Then
            If GetColumnIndex(tbl, keyCol) = 0 Or GetColumnIndex(tbl, valCol) = 0 Then
                Err.Raise ERR_UI_BASE + 40, MOD_NAME & ".HLI", _
                          "Kolona " & keyCol & " ili " & valCol & " ne postoji u " & tbl & "."
            End If
        End If
        HLI = HL(tbl, keyCol, keyVal, valCol)
        Exit Function
    End If
    Dim data As Variant: data = GetTableData(tbl)
    If IsEmpty(data) Then
        If strict Then
            Err.Raise ERR_UI_BASE + 29, MOD_NAME & ".HLI", _
                      "Tabela " & tbl & " nije citljiva."
        End If
        Exit Function
    End If
    Dim cKey As Long, cVal As Long, cGen As Long
    cKey = GetColumnIndex(tbl, keyCol)
    cVal = GetColumnIndex(tbl, valCol)
    cGen = GetColumnIndex(tbl, COL_GENERACIJA_ID)
    If cKey = 0 Or cVal = 0 Then
        If strict Then
            Err.Raise ERR_UI_BASE + 28, MOD_NAME & ".HLI", _
                      "Kolona " & keyCol & " ili " & valCol & " ne postoji u " & tbl & "."
        End If
        Exit Function
    End If
    ' Tabela bez kolone generacije, a identitet JE zadat -> ne moze da se sazna
    ' o kom je dokumentu rec. Povratak na broj bi ovde bio tiha degradacija
    ' unutar modela koji se posle oznacava kao valid, pa se umesto toga dize
    ' greska i ceo uvid pada.
    If cGen = 0 Then
        If strict Then
            Err.Raise ERR_UI_BASE + 27, MOD_NAME & ".HLI", _
                      "Tabela " & tbl & " nema kolonu " & COL_GENERACIJA_ID & "."
        End If
        HLI = HL(tbl, keyCol, keyVal, valCol)
        Exit Function
    End If
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cKey))) = Trim$(keyVal) Then
            If Trim$(CStr(data(i, cGen))) = Trim$(docID) Then
                HLI = Trim$(CStr(data(i, cVal)))
                Exit Function
            End If
        End If
    Next i
End Function

' Kilaza zbirne u uvidu dolazi sa STAVKI (S4-1).
'
' Zaglavlje je od PR3 prazno, pa bi SumActiveNum nad UkupnoKolicina za svaki
' kanonski dokument vratio prazno -- a ovo je ekran koji operateru kaze STA
' stornira. Prazno polje bi tamo znacilo "dokument je prazan", ne "citam
' pogresno mesto".
'
' Izbor redova je ISTI kao u SumActiveNum: aktivni redovi datog broja, suzeni
' generacijom kad je docID poznat. Kanonski pisac generaciju ne pise, pa za nov
' dokument suzavanja nema -- tada broj bira dokument, kao i pre.
' ZbirnaID za uvid: zadat se PROVERAVA, odsutan se IZVODI iz broja.
'
' Dve tvrdnje u jednoj funkciji, jer su jedna odluka:
'   1) zadat ID mora da postoji tacno jednom I da nosi bas taj broj -- inace bi
'      uvid tvrdio posledice jednog dokumenta pod brojem drugog (review #371);
'   2) bez ID-a se ide po broju, ali samo dok je broj jednoznacan. To NIJE
'      ustupak zatecenim podacima (njih nema) nego pozivaocima koji jos ne nose
'      identitet -- okvir ispravke i njegovi testovi, do S4-3.
'
' U strict rezimu oba slucaja DIZU gresku: uvid se posle oznacava kao valid, pa
' tiho prazno ovde znaci model koji tvrdi posledice a ne zna nad cim.
Private Function ZbrIdUvid(ByVal broj As String, ByVal docID As String, _
                           ByVal strict As Boolean) As String
    Const SRC As String = MOD_NAME & ".ZbrIdUvid"

    Dim data As Variant: data = GetTableData(TBL_ZBIRNA)
    If IsEmpty(data) Then
        If strict Then Err.Raise ERR_UI_BASE + 29, SRC, _
                                 "Tabela " & TBL_ZBIRNA & " nije citljiva."
        Exit Function
    End If

    Dim cId As Long, cBroj As Long, cSt As Long
    cId = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_ID)
    cBroj = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ)
    cSt = GetColumnIndex(TBL_ZBIRNA, COL_STORNIRANO)
    If cId = 0 Or cBroj = 0 Then
        If strict Then Err.Raise ERR_UI_BASE + 28, SRC, _
                                 "Kolona identiteta ili broja ne postoji u " & TBL_ZBIRNA & "."
        Exit Function
    End If

    Dim i As Long, nadjen As String, n As Long

    If Len(Trim$(docID)) > 0 Then
        For i = 1 To UBound(data, 1)
            If StrComp(Trim$(NzToText(data(i, cId))), Trim$(docID), vbTextCompare) = 0 Then
                n = n + 1
                nadjen = Trim$(NzToText(data(i, cBroj)))
            End If
        Next i

        If n <> 1 Then
            If strict Then Err.Raise ERR_UI_BASE + 34, SRC, _
                                     "ZbirnaID " & docID & " se nalazi " & CStr(n) & " puta."
            Exit Function
        End If

        If Len(Trim$(broj)) > 0 Then
            If StrComp(nadjen, Trim$(broj), vbTextCompare) <> 0 Then
                If strict Then Err.Raise ERR_UI_BASE + 35, SRC, _
                                         "Broj i identitet ne pripadaju istom dokumentu: " & _
                                         "broj " & broj & ", a ZbirnaID " & docID & _
                                         " nosi broj " & nadjen & "."
                Exit Function
            End If
        End If

        ZbrIdUvid = Trim$(docID)
        Exit Function
    End If

    ' Bez identiteta: po broju, i samo dok je jednoznacan.
    For i = 1 To UBound(data, 1)
        If StrComp(Trim$(NzToText(data(i, cBroj))), Trim$(broj), vbTextCompare) = 0 Then
            Dim jeStorn As Boolean: jeStorn = False
            If cSt > 0 Then jeStorn = (UCase$(Trim$(NzToText(data(i, cSt)))) = "DA")
            If Not jeStorn Then
                n = n + 1
                nadjen = Trim$(NzToText(data(i, cId)))
            End If
        End If
    Next i

    If n <> 1 Then
        If strict Then Err.Raise ERR_UI_BASE + 36, SRC, _
                                 "Broj " & broj & " nosi " & CStr(n) & " aktivnih zbirnih."
        Exit Function
    End If

    ZbrIdUvid = nadjen
End Function

' Vrednost kolone zbirne po ZbirnaID-u (S4-2).
'
' HLI bira red po paru (broj, generacija) i tako mora da ostane za prijemnicu,
' koja je do S6 na starom modelu. Zbirna od S4-2 ima identitet, pa joj se red
' bira njime -- kanonski dokument generaciju uopste nema, a broj nije identitet.
Private Function HLZ(ByVal valCol As String, ByVal zbirnaID As String, _
                     ByVal strict As Boolean) As String
    Dim data As Variant: data = GetTableData(TBL_ZBIRNA)
    If IsEmpty(data) Then
        If strict Then
            Err.Raise ERR_UI_BASE + 29, MOD_NAME & ".HLZ", _
                      "Tabela " & TBL_ZBIRNA & " nije citljiva."
        End If
        Exit Function
    End If

    Dim cId As Long, cVal As Long
    cId = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_ID)
    cVal = GetColumnIndex(TBL_ZBIRNA, valCol)
    If cId = 0 Or cVal = 0 Then
        If strict Then
            Err.Raise ERR_UI_BASE + 28, MOD_NAME & ".HLZ", _
                      "Kolona " & COL_ZBR_ID & " ili " & valCol & _
                      " ne postoji u " & TBL_ZBIRNA & "."
        End If
        Exit Function
    End If

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If StrComp(Trim$(NzToText(data(i, cId))), Trim$(zbirnaID), vbTextCompare) = 0 Then
            HLZ = Trim$(NzToText(data(i, cVal)))
            Exit Function
        End If
    Next i
End Function

Private Function ZbirnaKgZaUvid(ByVal zbirnaID As String, _
                                ByVal strict As Boolean) As String
    Dim zid As String
    zid = Trim$(zbirnaID)
    If Len(zid) = 0 Then
        If strict Then
            Err.Raise ERR_UI_BASE + 33, MOD_NAME & ".ZbirnaKgZaUvid", _
                      "ZbirnaID je obavezan: uvid ne bira dokument po broju."
        End If
        Exit Function
    End If

    Dim zbirStavki As Object
    Set zbirStavki = modDokumenta.ZbirStavkiPoZbirni()

    If Not zbirStavki.Exists(zid) Then
        If strict Then
            Err.Raise ERR_UI_BASE + 32, MOD_NAME & ".ZbirnaKgZaUvid", _
                      "Zbirna " & zid & " nema nijednu stavku."
        End If
        Exit Function
    End If

    ZbirnaKgZaUvid = CStr(CDbl(zbirStavki(zid)(0)))
End Function

' ============================================================
' TEST HOOK (Alt+F8): ispisi model uticaja u Immediate (Ctrl+G). Read-only.
' ============================================================
Public Sub Test_BuildStornoImpact()
    Dim broj As String
    broj = Trim$(InputBox("Broj dokumenta:", "Impact test"))
    If Len(broj) = 0 Then Exit Sub
    Dim tip As String
    tip = Trim$(InputBox("Tip (Prijemnica / Otpremnica / Zbirna):", "Impact test", "Prijemnica"))
    If Len(tip) = 0 Then Exit Sub

    Dim d As Object: Set d = BuildStornoImpact(tip, broj)
    Dim h As Object: Set h = d("header")
    Dim sm As Object: Set sm = d("summary")
    Debug.Print "=== IMPACT " & tip & " " & broj & " ==="
    Debug.Print "Header: partnerID=" & h("partnerID") & " datum=" & h("datum") & " kol=" & h("kolicina")
    Debug.Print "Summary: blocks=" & sm("blockCount") & " palete=" & sm("paleteCount") & _
                " | detach gajb=" & sm("detachGajb") & " neto=" & sm("detachNeto") & " amb=" & sm("detachAmb")

    Dim pal As Collection: Set pal = d("palete")
    Dim i As Long
    For i = 1 To pal.count
        Dim p As Object: Set p = pal(i)
        Debug.Print "  Paleta " & p("label") & ": " & p("used") & "/" & p("cap") & " gajb" & _
                    " | this=" & p("thisGajb") & " (neto " & p("thisNeto") & ", amb " & p("thisAmb") & ")" & _
                    IIf(CBool(p("preradjena")), " [PRERADJENA]", "")
    Next i

    Dim fk As Object: Set fk = d("faktura")
    Debug.Print "Faktura: has=" & fk("hasFaktura") & " id=" & fk("fakturaID")
    Debug.Print "=== kraj ==="
End Sub
