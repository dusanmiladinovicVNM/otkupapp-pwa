Attribute VB_Name = "modDokumenta"
'Attribute VB_Name = "modDokumenta"

Option Explicit

' ============================================================
' modDokumenta - Otpremnica, Zbirna, Prijemnica
' Dokumentenfluss: Otkup zu Otpremnica zu Zbirna zu Prijemnica zu Faktura
' ============================================================

' ============================================================
' ZBR-IDENT-01 -- IDENTITET ZBIRNE
'
' Ugovor i acceptance testovi: docs/DOMEN/ZBR_IDENTITET.md
'
' Broj zbirne je LABELA, ne identitet. Identitet logickog dokumenta je
' GeneracijaID; broj + VozacID + KupacID je SCOPE u kome se generacija nalazi
' ili kuje. U STAROM modelu je dvoklasna zbirna bila JEDAN dokument na DVA
' reda: pisac je isti broj, vozaca i kupca upisivao dvaput, pa su oba reda
' nosila istu generaciju. Kanon to vise ne radi -- klase su stavke jednog
' zaglavlja -- ali resolver i dalje mora da procita zateceni dvoredni oblik.
'
' Resolver odgovara na DVA NEZAVISNA pitanja, i to je cela poenta razdvajanja:
'   resolutionStatus     -- da li je broj SADA jednoznacno razresiv (bez storniranih)
'   historicalOwnerCount -- da li je broj IKAD pripadao vise od jednog vlasnika
' UNIQUE uz historicalOwnerCount = 2 je VALIDNO stanje, ne kontradikcija: broj je
' danas jednoznacan, a u proslosti ga je drzao i drugi vlasnik.
'
' activeLogicalCount i activeOwnerCount se NE izvode jedan iz drugog: dvoklasna
' zbirna daje 1/1 iz dva reda, dva zasebna unosa istog vozaca daju 2/1.
' ============================================================
Public Const ZBR_INT_OK As String = "OK"
Public Const ZBR_INT_ERROR As String = "INTEGRITY_ERROR"

Public Const ZBR_RES_NONE As String = "NONE"
Public Const ZBR_RES_UNIQUE As String = "UNIQUE"
Public Const ZBR_RES_OWNER_MISMATCH As String = "OWNER_MISMATCH"
Public Const ZBR_RES_AMBIGUOUS As String = "CURRENT_AMBIGUOUS"

' Razlog odbijanja novog unosa. Kod, ne poruka: modDokumenta je sloj podataka i
' ne nosi korisnicki tekst (prevod je u modDokUnos.ZbirnaGatePoruka).
Public Const ZBR_GATE_INTEGRITET As String = "INTEGRITET"
Public Const ZBR_GATE_AKTIVNA As String = "AKTIVNA"
Public Const ZBR_GATE_TUDJ As String = "TUDJ"
Public Const ZBR_GATE_SIROCE As String = "SIROCE"
' Broj stornirane zbirne ISTOG vlasnika: storno ne oslobadja broj (A9, odluka
' 14.09.2026). Nije TUDJ -- ta poruka bi operateru rekla da je broj pripadao
' drugom vozacu ili kupcu, sto nije tacno.
Public Const ZBR_GATE_STORNIRAN As String = "STORNIRAN"

' Razlog odbijanja RODITELJA (F4). Zaseban skup od ZBR_GATE_*: tamo je pitanje
' "sme li NOV broj", ovde "sme li se prijemnica vezati na POSTOJECI dokument".
Public Const ZBR_PARENT_NEMA As String = "NEMA"
Public Const ZBR_PARENT_DVOSMISLEN As String = "DVOSMISLEN"
Public Const ZBR_PARENT_TUDJ As String = "TUDJ_VLASNIK"
Public Const ZBR_PARENT_ISTORIJA As String = "ISTORIJA"

' ZBR-MUT-01: razlozi zbog kojih se po BROJU ne sme mutirati.
'
' Deca zbirne (otpremnica, prijemnica, paletna stavka, denormalizovan otkup) od
' faze 1 nose i ZbirnaGeneracijaID -- ali on sme biti PRAZAN (roditelj jos nije
' razresen), pa se na njega ne moze racunati bez provere.
'
' Rutina koja decu bira po broju zato i dalje zahvata SVE dokumente tog broja --
' osim kad je izbor scoped po generaciji (faza 3), a to zna samo pozivalac.
' Otud `scopedPoGeneraciji` parametar nize: kapija popusta tek kad akter dokaze
' da bira po generaciji, i to za CELU svoju operaciju.
'
' Dva razloga su RAZLICITA i ne smeju se stopiti:
'   VLASNIK   -- broj je IKAD pripadao vise od jednog (VozacID, KupacID).
'                Storniran vlasnik se broji: deca mu ostaju aktivna.
'   DOKUMENTI -- broj SADA nosi vise od jednog aktivnog logickog dokumenta,
'                makar bili istog vlasnika (A17). Do v6-ui-224 to stanje uvoz
'                nije ni pravio, jer je stapao generacije; sada ga pravi
'                ispravno, pa mora i da se vidi.
'
' historicalLogicalCount se MERI i prijavljuje, ali NE blokira. Obrazlozenje je
' bilo da ga redovan re-entry pod istim brojem (ALLOW grana ZBR-ACTIVE-NUMBER-01)
' zakuje na 2. Ta grana je ukinuta 14.09.2026 -- ispravka dobija nov broj -- pa
' re-entry pod istim brojem sada nastaje samo mimo F3 (PWA uvoz, malina
' auto-zbirna -- oba pauzirana). Postrozavanje ove kapije je zaseban korak
' (ZBR_IDENTITET.md, faza 4), ne posledica te odluke.
Public Const ZBR_MUT_INTEGRITET As String = "INTEGRITET"
Public Const ZBR_MUT_VISE_VLASNIKA As String = "VISE_VLASNIKA"
Public Const ZBR_MUT_VISE_DOKUMENATA As String = "VISE_DOKUMENATA"

' ZBR-CHILD-01 faza 2: ZASTO backfill preskace broj. Tri razloga, ne jedan --
' prvo merenje nad pravim podacima je dalo 85 popunjenih i 4572 preskocena, a
' poruka je sva tri prijavljivala kao "broj je IKAD nosio vise dokumenata".
' Migracija koja ne ume da kaze STA je zatekla ne moze da vodi sledeci korak.
Public Const ZBR_BF_INTEGRITET As String = "INTEGRITET"
Public Const ZBR_BF_NEMA_GENERACIJE As String = "NEMA_GENERACIJE"
Public Const ZBR_BF_VISE_GENERACIJA As String = "VISE_GENERACIJA"

Public Type ZbirnaIdent
    normalizedBroj As String
    integrityStatus As String
    activeLogicalCount As Long
    activeOwnerCount As Long
    historicalOwnerCount As Long
    historicalLogicalCount As Long
    historicalOnlyZbirnaID As String       ' popunjeno samo kad je historicalLogicalCount = 1
    scopeProvided As Boolean
    historicalOwnerIsScope As Boolean
    matchingScopeActiveLogicalCount As Long
    resolutionStatus As String
    selectedZbirnaID As String
    selectedVozacID As String
    selectedKupacID As String
    brojUAktivnojPrijemnici As Boolean
End Type

' ============================================================
' OTPREMNICA
'
' S3b-1: stari pisac "red po klasi" (SaveOtpremnicaMulti_TX, SaveOtpremnica_TX,
' SaveOtpremnica, ValidateOtpremnicaInput) je OBRISAN. Od S3a nije imao
' produkcionog pozivaoca, a pravio je dokument koji novi citaoci odbijaju:
' zaglavlje bez stavki, vezano za zbirnu kroz BrojZbirne. Pisci otpremnice su
' CreateOtpremnicaDraft_TX / CreateOtpremnicaIzIzvora_TX i ostali *_TX nize.
' ============================================================

Public Function GetOtpremniceByZbirna(ByVal brojZbirne As String) As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_OTPREMNICA)

    If IsEmpty(data) Then
        GetOtpremniceByZbirna = Empty
        Exit Function
    End If
    data = ExcludeStornirano(data, TBL_OTPREMNICA)

    If IsEmpty(data) Then
        GetOtpremniceByZbirna = Empty
        Exit Function
    End If

    Dim filters As New Collection
    Dim fp As clsFilterParam

    Set fp = New clsFilterParam
    fp.Init RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, _
            "modDokumenta.GetOtpremniceByZbirna"), "=", brojZbirne
    filters.Add fp

    GetOtpremniceByZbirna = FilterArray(data, filters)
    Exit Function

EH:
    LogErr "modDokumenta.GetOtpremniceByZbirna"
    GetOtpremniceByZbirna = Empty
End Function

' True ako zbirna sa datim brojem postoji u tblZbirna (iskljucujuci stornirane).
' Storno-aware namerno: prijemnica ne sme da referencira storniranu zbirnu.
' (modBrojevi.BrojZbirneExists je Private i broji i stornirane -> drugi scenario.)
Public Function ZbirnaPostoji(ByVal brojZbirne As String) As Boolean
    On Error GoTo EH

    Dim b As String: b = Trim$(brojZbirne)
    If Len(b) = 0 Then Exit Function

    Dim data As Variant
    data = GetTableData(TBL_ZBIRNA)
    If IsEmpty(data) Then Exit Function
    data = ExcludeStornirano(data, TBL_ZBIRNA)
    If IsEmpty(data) Then Exit Function

    Dim iBroj As Long
    iBroj = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, "modDokumenta.ZbirnaPostoji")

    Dim r As Long
    For r = 1 To UBound(data, 1)
        If StrComp(Trim$(CStr(data(r, iBroj))), b, vbTextCompare) = 0 Then
            ZbirnaPostoji = True
            Exit Function
        End If
    Next r
    Exit Function

EH:
    LogErr "modDokumenta.ZbirnaPostoji", "broj=" & brojZbirne
End Function

' Normalizacija broja zbirne -- JEDNO mesto za ceo ZBR-IDENT lanac.
'
' CheckDuplicate poredi SIROVO (CStr(...) = searchValue, bez Trim, case-sensitive),
' pa " 5/070926 " prolazi pored "5/070926". Zbirna ga od 14.09.2026 vise ne zove
' (D2 je povucena): zauzetost drzi modBrojevi.BrojZauzetUNizu, a normalizaciju
' ZBR-IDENT lanca i dalje OVAJ omotac.
Public Function ZbirnaBrojNorm(ByVal broj As String) As String
    ZbirnaBrojNorm = Trim$(NzToText(broj))
End Function

' Vlasnik zbirne je KOMPOZITAN: VozacID + KupacID. Isti spisak koji koriste
' ApplyGeneracijaID i modStorno.VlasniciPoBroju.
Private Function ZbirnaVlasnikKljuc(ByVal vozacID As Variant, _
                                    ByVal kupacID As Variant) As String
    ZbirnaVlasnikKljuc = UCase$(Trim$(NzToText(vozacID))) & "|" & _
                         UCase$(Trim$(NzToText(kupacID)))
End Function

' Razresava BROJ zbirne u LOGICKI DOKUMENT. Puni ceo DTO iz jednog citanja
' tabele: sirov niz nosi IKAD, ExcludeStornirano nad njim nosi SADA.
'
' Zasto ExcludeStornirano a ne sopstveni test na kolonu Stornirano: to je ISTI
' filtar koji zovu ZbirnaPostoji i GeneracijaIDZaBrojArr. Resolver, picker i
' writer time gledaju istu definiciju "aktivne" i ne mogu da se raziidju.
'
' Zasto NE modStorno.VlasniciPoBroju, koja je istog oblika: ona poredi broj
' case-sensitive (ZBR-NORM-02) i broji samo vlasnike, ne generacije. Ovde su
' potrebna oba, pa i normalizacija mora da bude sopstvena.
'
' FAIL-CLOSED: kad integrityStatus nije OK, resolutionStatus NIKAD nije NONE.
' NONE je jedina vrednost koja negde znaci "sme se" (kapija za nov unos), pa se
' ne sme dobiti iz greske. Aktivan red bez generacije zato daje CURRENT_AMBIGUOUS
' uz INTEGRITY_ERROR -- oba znace tvrdu blokadu, a integrityStatus kaze zasto.
' Identitet se NE pogadja iz broja i vlasnika (D4: nema legacy fallbacka).
Public Function ZbirnaIdentResolve(ByVal broj As String, _
                                   Optional ByVal vozacID As String = "", _
                                   Optional ByVal kupacID As String = "") As ZbirnaIdent
    Const SRC As String = "modDokumenta.ZbirnaIdentResolve"

    Dim res As ZbirnaIdent
    res.integrityStatus = ZBR_INT_OK
    res.resolutionStatus = ZBR_RES_NONE
    res.normalizedBroj = ZbirnaBrojNorm(broj)
    res.scopeProvided = (Len(Trim$(NzToText(vozacID))) > 0) And _
                        (Len(Trim$(NzToText(kupacID))) > 0)

    On Error GoTo EH

    If Len(res.normalizedBroj) = 0 Then GoTo XIT

    ' I1: da li broj vec drzi AKTIVNA prijemnica. Cita se iz kanonskog
    ' read-modela, da postoji JEDNA definicija "zauzetog broja".
    Dim unija As Object
    Set unija = AktivniBrojeviZbirne()
    If unija.Exists(res.normalizedBroj) Then
        res.brojUAktivnojPrijemnici = _
            (InStr(1, CStr(unija(res.normalizedBroj)), "P", vbBinaryCompare) > 0)
    End If

    Dim sirovo As Variant
    sirovo = GetTableData(TBL_ZBIRNA)
    If Not IsArray(sirovo) Then GoTo XIT

    Dim cBr As Long, cVoz As Long, cKup As Long, cIdent As Long
    cBr = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, SRC)
    cVoz = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_VOZAC, SRC)
    cKup = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KUPAC, SRC)
    ' IDENTITET REDA ZBIRNE JE ZbirnaID (S4-3b).
    '
    ' Do ovog reza je ovde stajao COL_GENERACIJA_ID, i to je bilo tacno dok su
    ' generaciju PISALI modMasterSync i modDokumentInvariant. Prvi je pauziran,
    ' drugi je obrisan u S4-3a -- pa nijedan ziv pisac vise ne pecati generaciju,
    ' a kanonski je nikad nije ni pisao.
    '
    ' Posledica koju je ovaj rez nasao kao ZIV KVAR: svaka kanonski napravljena
    ' zbirna padala je na "prazna generacija = integritetska greska", pa su DUPLI
    ' i PONISTENJE bili nedostupni za SVAKI dokument koji aplikacija danas pravi
    ' -- uz poruku koja operatera salje na B9 proveru koja to ne moze popraviti.
    cIdent = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_ID, SRC)

    Dim scopeKljuc As String
    If res.scopeProvided Then scopeKljuc = ZbirnaVlasnikKljuc(vozacID, kupacID)

    ' --- IKAD: sirov niz, stornirani se BROJE ---
    Dim ikadVl As Object: Set ikadVl = CreateObject("Scripting.Dictionary")
    Dim ikadIds As Object: Set ikadIds = CreateObject("Scripting.Dictionary")
    Dim r As Long, vl As String, ikadId As String
    For r = 1 To UBound(sirovo, 1)
        If StrComp(Trim$(NzToText(sirovo(r, cBr))), res.normalizedBroj, vbTextCompare) = 0 Then
            vl = ZbirnaVlasnikKljuc(sirovo(r, cVoz), sirovo(r, cKup))
            If Not ikadVl.Exists(vl) Then ikadVl.Add vl, 1
            ikadId = UCase$(Trim$(NzToText(sirovo(r, cIdent))))
            If Len(ikadId) > 0 Then
                If Not ikadIds.Exists(ikadId) Then ikadIds.Add ikadId, 1
            End If
        End If
    Next r
    res.historicalOwnerCount = ikadVl.Count
    ' MERI SE, NE BLOKIRA -- v. komentar uz ZBR_MUT_* konstante.
    res.historicalLogicalCount = ikadIds.Count
    If res.historicalLogicalCount = 1 Then res.historicalOnlyZbirnaID = ikadIds.Keys()(0)

    If res.scopeProvided And res.historicalOwnerCount = 1 Then
        res.historicalOwnerIsScope = (StrComp(ikadVl.Keys()(0), scopeKljuc, vbTextCompare) = 0)
    End If

    ' --- SADA: isti niz kroz ExcludeStornirano ---
    Dim akt As Variant
    akt = ExcludeStornirano(sirovo, TBL_ZBIRNA)
    If Not IsArray(akt) Then GoTo XIT

    Dim aktIds As Object: Set aktIds = CreateObject("Scripting.Dictionary")
    Dim aktVl As Object: Set aktVl = CreateObject("Scripting.Dictionary")
    Dim idVoz As Object: Set idVoz = CreateObject("Scripting.Dictionary")
    Dim idKup As Object: Set idKup = CreateObject("Scripting.Dictionary")
    Dim scopeIds As Object: Set scopeIds = CreateObject("Scripting.Dictionary")

    Dim zid As String, prazneAktivne As Long
    For r = 1 To UBound(akt, 1)
        If StrComp(Trim$(NzToText(akt(r, cBr))), res.normalizedBroj, vbTextCompare) = 0 Then
            vl = ZbirnaVlasnikKljuc(akt(r, cVoz), akt(r, cKup))
            If Not aktVl.Exists(vl) Then aktVl.Add vl, 1

            zid = UCase$(Trim$(NzToText(akt(r, cIdent))))
            If Len(zid) = 0 Then
                prazneAktivne = prazneAktivne + 1
            Else
                If Not aktIds.Exists(zid) Then
                    aktIds.Add zid, 1
                    idVoz.Add zid, Trim$(NzToText(akt(r, cVoz)))
                    idKup.Add zid, Trim$(NzToText(akt(r, cKup)))
                End If
                If res.scopeProvided Then
                    If StrComp(vl, scopeKljuc, vbTextCompare) = 0 Then
                        If Not scopeIds.Exists(zid) Then scopeIds.Add zid, 1
                    End If
                End If
            End If
        End If
    Next r

    res.activeLogicalCount = aktIds.Count
    res.activeOwnerCount = aktVl.Count
    res.matchingScopeActiveLogicalCount = scopeIds.Count

    ' ZBR-IDENT-01, sada nad ZbirnaID-em: aktivan red MORA da nosi identitet.
    ' Prazan i dalje jeste integritetska greska -- ali to je greska koju pisac ne
    ' moze da napravi (NewEntityID), pa je ovo kapija nad zateceno pokvarenim
    ' redom, a ne nad normalnim dokumentom. Isto meri Chk_B9.
    If prazneAktivne > 0 Then
        res.integrityStatus = ZBR_INT_ERROR
        res.resolutionStatus = ZBR_RES_AMBIGUOUS
        GoTo XIT
    End If

    If res.activeLogicalCount = 0 Then
        res.resolutionStatus = ZBR_RES_NONE
    ElseIf res.activeLogicalCount > 1 Then
        res.resolutionStatus = ZBR_RES_AMBIGUOUS
    ElseIf res.scopeProvided And res.matchingScopeActiveLogicalCount = 0 Then
        res.resolutionStatus = ZBR_RES_OWNER_MISMATCH
    Else
        res.resolutionStatus = ZBR_RES_UNIQUE
        zid = aktIds.Keys()(0)
        res.selectedZbirnaID = zid
        res.selectedVozacID = idVoz(zid)
        res.selectedKupacID = idKup(zid)
    End If

XIT:
    ZbirnaIdentResolve = res
    Exit Function

EH:
    LogErr SRC, "broj=" & broj
    res.integrityStatus = ZBR_INT_ERROR
    res.resolutionStatus = ZBR_RES_AMBIGUOUS
    res.selectedZbirnaID = ""
    res.selectedVozacID = ""
    res.selectedKupacID = ""
    ZbirnaIdentResolve = res
End Function

' Kapija ZBR-ACTIVE-NUMBER-01 (ugovor par.5) kao TABELA, na jednom mestu.
'
' Strogo, BEZ izuzetka za istog vlasnika: broj koji je IKAD nosila zbirna znaci
' NE, ma ciji bio i bio aktivan ili storniran. Ispravka je storno pa nov unos pod
' NOVIM brojem (A9, odluka 14.09.2026). To nista ne lomi: ZbirnaValidiraj se zove
' TACNO jednom (modScrDokumenti Scr_Save), PRE upisa, pa validator nikad ne
' vidi red koji je sam upravo napisao.
Public Function ZbirnaNovUnosRazlog(ByRef id As ZbirnaIdent) As String
    If id.integrityStatus <> ZBR_INT_OK Then
        ZbirnaNovUnosRazlog = ZBR_GATE_INTEGRITET
        Exit Function
    End If

    If id.activeLogicalCount > 0 Then
        ZbirnaNovUnosRazlog = ZBR_GATE_AKTIVNA
        Exit Function
    End If

    If id.historicalOwnerCount = 0 Then
        ' Nijedna zbirna IKAD pod tim brojem -- slobodno, OSIM ako ga vec drzi
        ' aktivna prijemnica (I1). Tada bi nova zbirna tiho postala njen
        ' roditelj: prijemnica se vezuje samo brojem i ne bi ni primetila.
        If id.brojUAktivnojPrijemnici Then ZbirnaNovUnosRazlog = ZBR_GATE_SIROCE
        Exit Function
    End If

    ' Aktivnih nema, a istorija postoji: broj je zakljucan za SVE. Drugi vlasnik
    ' dobija TUDJ, isti STORNIRAN -- ispravka zbirne dobija nov broj (A9, odluka
    ' 14.09.2026), a ZbirnaValidiraj joj izvor trazi po starom broju iz konteksta
    ' ispravke. Do tada je isti vlasnik ovde smeo ponovo.
    If id.historicalOwnerIsScope Then
        ZbirnaNovUnosRazlog = ZBR_GATE_STORNIRAN
    Else
        ZbirnaNovUnosRazlog = ZBR_GATE_TUDJ
    End If
End Function

' Tanak omotac nad JEDNOM tabelom iznad -- da druga kopija pravila ne odluta.
Public Function ZbirnaSmeNovUnos(ByRef id As ZbirnaIdent) As Boolean
    ZbirnaSmeNovUnos = (Len(ZbirnaNovUnosRazlog(id)) = 0)
End Function

' Ugovor par.6: prijemnica se vezuje SAMO na jednoznacno razresen dokument.
' Kod CURRENT_AMBIGUOUS / OWNER_MISMATCH / INTEGRITY_ERROR se kanonski roditelj
' NE trazi -- biranje "najverovatnijeg" iz dvosmislenog skupa je tiho pogadjanje.
'
' ISTORIJA JE DEO BEZBEDNOSTI, ne samo sadasnje stanje. UNIQUE danas uz broj koji
' je IKAD drzalo vise vlasnika i dalje nije bezbedan roditelj: nizvodna operacija
' koja ide po broju moze da zahvati i tudje. Prijemnica od faze 1 ima FK na
' generaciju, ali on sme biti prazan, pa sam po sebi ne ukida ovu kapiju.
' Zato postoji i modStorno.RequireJedanVlasnikIkadPoBroju -- ista kapija za
' mutaciju po broju.
' ZBR-MUT-01 -- sme li se po BROJU mutirati ono sto visi o zbirni.
'
' Trece pitanje, uz kapiju za kreiranje (F3) i za roditelja (F4). Ovde se ne pita
' "sme li nov unos" ni "koji je roditelj", nego: SME LI RUTINA KOJA DECU BIRA PO
' BROJU da radi. Izbor po broju zahvata sve dokumente tog broja; generacija na
' detetu postoji od faze 1, ali sme biti prazna, pa je popustanje uslovljeno --
' v. scopedPoGeneraciji.
'
' Vraca prazno kad sme. Razlog se ne stapa u jednu poruku: vise vlasnika i vise
' dokumenata istog vlasnika su dva razlicita poteza za operatera.
'
' Namerno NE gleda historicalLogicalCount -- v. komentar uz ZBR_MUT_* konstante.
' `scopedPoGeneraciji` je faza 4: kapija SME da pusti dva aktivna dokumenta pod
' istim brojem, ali samo tamo gde akter vise ne bira decu po broju.
'
' Default je False i to nije opreznost nego nuznost: kapiju zove vise mesta, a
' samo deo njih bira decu po generaciji. Ostali i dalje biraju PO BROJU, pa bi
' bezuslovna grana pustila aktera na tudje redove -- tiho i u kilogramima. To je
' ZBR-MUT-01 naopako: ne sirenjem aktera nego suzavanjem kapije.
'
' S4-3a: primer koji je ovde stajao (rekalkulacija zbirne iz otpremnica po
' BrojZbirne) je obrisan sa okvirom ispravke. Pravilo nije -- DUPLI i
' PONISTENJE decu i dalje biraju brojem.
'
' Pozivalac NE sme da salje "postoji generacija" nego BAS onu odluku koju vec
' racuna za svoju selekciju (`genEff <> ""`). Kapija i akter tako gledaju isti
' izraz, ne dva slicna.
Public Function ZbirnaMutacijaPoBrojuRazlog(ByRef id As ZbirnaIdent, _
                                            Optional ByVal scopedPoGeneraciji As Boolean = False) As String
    If id.integrityStatus <> ZBR_INT_OK Then
        ZbirnaMutacijaPoBrojuRazlog = ZBR_MUT_INTEGRITET
        Exit Function
    End If

    ' IKAD, ne samo sada: storniran vlasnik i dalje moze imati AKTIVNU decu, a
    ' ona nose isti broj.
    If id.historicalOwnerCount > 1 Then
        ZbirnaMutacijaPoBrojuRazlog = ZBR_MUT_VISE_VLASNIKA
        Exit Function
    End If

    ' SADA: samo aktivni dokumenti konkurisu za decu. Dva aktivna istog vlasnika
    ' (A17) owner-brojac ne vidi -- zbog toga ova grana i postoji.
    '
    ' Faza 4: kad akter bira decu po generaciji, "dva aktivna dokumenta" mu vise
    ' nije opasnost -- dira samo svoje. Grana tada nema sta da brani.
    If scopedPoGeneraciji Then Exit Function

    If id.activeLogicalCount > 1 Then
        ZbirnaMutacijaPoBrojuRazlog = ZBR_MUT_VISE_DOKUMENATA
    End If
End Function

' Isto, ali od samog broja -- za pozivaoce koji nemaju razresen DTO.
'
' FAIL-CLOSED na sopstvenu gresku: "ne mogu da dokazem jednoznacnost" je za
' kapiju isto sto i "ne mutiraj". Prazan broj nije nerazresen nego "nema
' roditelja" -- nema sta da se mutira.
Public Function ZbirnaMutacijaPoBrojuRazlogZaBroj(ByVal broj As String, _
                                                 Optional ByVal scopedPoGeneraciji As Boolean = False) As String
    Dim id As ZbirnaIdent
    On Error GoTo EH
    If Len(Trim$(NzToText(broj))) = 0 Then Exit Function
    id = ZbirnaIdentResolve(broj)
    ZbirnaMutacijaPoBrojuRazlogZaBroj = ZbirnaMutacijaPoBrojuRazlog(id, scopedPoGeneraciji)
    Exit Function
EH:
    LogErr "modDokumenta.ZbirnaMutacijaPoBrojuRazlogZaBroj"
    ZbirnaMutacijaPoBrojuRazlogZaBroj = ZBR_MUT_INTEGRITET
End Function

Public Function ZbirnaRoditeljRazlog(ByRef id As ZbirnaIdent) As String
    If id.integrityStatus <> ZBR_INT_OK Then
        ZbirnaRoditeljRazlog = ZBR_GATE_INTEGRITET
        Exit Function
    End If

    Select Case id.resolutionStatus
        Case ZBR_RES_UNIQUE
            ' Jednoznacan DANAS jos nije dovoljno -- v. komentar iznad.
            If id.historicalOwnerCount > 1 Then
                ZbirnaRoditeljRazlog = ZBR_PARENT_ISTORIJA
            End If
        Case ZBR_RES_AMBIGUOUS
            ZbirnaRoditeljRazlog = ZBR_PARENT_DVOSMISLEN
        Case ZBR_RES_OWNER_MISMATCH
            ZbirnaRoditeljRazlog = ZBR_PARENT_TUDJ
        Case Else
            ' NONE: pod tim brojem nema aktivnog dokumenta, pa nema ni roditelja.
            ' Politiku za taj slucaj (BLOK ili UPOZORENJE, po
            ' PRIJEMNICA_ZBIRNA_PROVERA) drzi ZbirnaPostoji u validatoru IZNAD
            ' ove kapije -- ovde stoji samo da funkcija ne laze kad se pozove sama.
            ZbirnaRoditeljRazlog = ZBR_PARENT_NEMA
    End Select
End Function

' Tanak omotac nad JEDNOM tabelom iznad -- da druga kopija pravila ne odluta.
Public Function ZbirnaRoditeljOK(ByRef id As ZbirnaIdent) As Boolean
    ZbirnaRoditeljOK = (Len(ZbirnaRoditeljRazlog(id)) = 0)
End Function

' I1 -- KANONSKI READ-MODEL AKTIVNIH BROJEVA ZBIRNE (ugovor par.8).
'
' Skup = normalizovani brojevi AKTIVNIH zbirnih UNIJA normalizovani BrojZbirne
' iz AKTIVNIH prijemnica.
'
' Unija nije opreznost nego ispravka rupe: broj koji referencira aktivna
' prijemnica je ZAUZET i onda kad mu je zbirna-red storniran ili nikad nije
' upisan. Bez tog drugog izvora takav broj izgleda slobodan.
'
' Vrednost kaze ODAKLE broj dolazi -- "Z", "P" ili "ZP" -- da pozivalac moze da
' razlikuje siroce (samo prijemnica) od normalnog para.
Public Function AktivniBrojeviZbirne() As Object
    Const SRC As String = "modDokumenta.AktivniBrojeviZbirne"

    Dim res As Object
    Set res = CreateObject("Scripting.Dictionary")
    res.CompareMode = vbTextCompare
    Set AktivniBrojeviZbirne = res

    On Error GoTo EH

    DodajBrojeve res, TBL_ZBIRNA, COL_ZBR_BROJ, "Z", SRC
    DodajBrojeve res, TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, "P", SRC
    Exit Function

EH:
    LogErr SRC
End Function

' Jedan izvor u recnik. Oznaka se DODAJE na postojecu, pa broj koji nose obe
' tabele dobije "ZP" -- razlika siroce/par se ne gubi.
Private Sub DodajBrojeve(ByRef res As Object, ByVal tblName As String, _
                         ByVal brojCol As String, ByVal oznaka As String, _
                         ByVal sourceName As String)
    Dim data As Variant
    data = GetTableData(tblName)
    If Not IsArray(data) Then Exit Sub
    data = ExcludeStornirano(data, tblName)
    If Not IsArray(data) Then Exit Sub

    Dim cBr As Long
    cBr = RequireColumnIndex(tblName, brojCol, sourceName)

    Dim r As Long, b As String
    For r = 1 To UBound(data, 1)
        b = ZbirnaBrojNorm(NzToText(data(r, cBr)))
        If Len(b) > 0 Then
            If res.Exists(b) Then
                If InStr(1, CStr(res(b)), oznaka, vbBinaryCompare) = 0 Then
                    res(b) = CStr(res(b)) & oznaka
                End If
            Else
                res.Add b, oznaka
            End If
        End If
    Next r
End Sub

Public Function GetOtpremniceByStation(ByVal stanicaID As String, _
                                       Optional ByVal datumOd As Date = 0, _
                                       Optional ByVal datumDo As Date = 0) As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_OTPREMNICA)

    If IsEmpty(data) Then
        GetOtpremniceByStation = Empty
        Exit Function
    End If
    
    data = ExcludeStornirano(data, TBL_OTPREMNICA)

    If IsEmpty(data) Then
        GetOtpremniceByStation = Empty
        Exit Function
    End If

    Dim filters As New Collection
    Dim fp As clsFilterParam

    Set fp = New clsFilterParam
    fp.Init RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_STANICA, _
            "modDokumenta.GetOtpremniceByStation"), "=", stanicaID
    filters.Add fp

    If datumOd > 0 And datumDo > 0 Then
        Set fp = New clsFilterParam
        fp.Init RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_DATUM, _
                "modDokumenta.GetOtpremniceByStation"), "BETWEEN", datumOd, datumDo
        filters.Add fp
    End If

    GetOtpremniceByStation = FilterArray(data, filters)
    Exit Function

EH:
    LogErr "modDokumenta.GetOtpremniceByStation"
    GetOtpremniceByStation = Empty
End Function

' ============================================================
' ZBIRNA - Gesamtdokument Fahrer
' ============================================================
' ============================================================
' ZBIRNA -- header + stavke (PR3)
' ============================================================
'
' Jedan javni ulaz koji vraca JEDAN ID, i koji obuhvata CELU poslovnu operaciju.
' Obrazac je CreateFaktura_TX (modFaktura.bas:11): _TX drzi transakciju i
' monitoring, Private core radi posao, kompletna prevalidacija PRE ijednog upisa.
'
' Sta se promenilo u odnosu na stari pisac (obrisan u S4-2c/2a):
'
'   staro:  dva reda u tblZbirna (Klasa I i Klasa II), dva ID-a, pa string
'           "ZBR-1 + ZBR-2" koji pozivalac posle parsira
'   novo:   JEDAN header u tblZbirna + N redova u tblZbirnaStavke, jedan ID
'
' KOLICINE SE NE PRIMAJU -- IZVODE SE.
'
' Zbirna je agregat svojih otpremnica (DOCUMENT_HEADER_LINES.md S4.3: stavke su
' kes, otpremnice su izvor istine). Zato writer prima IZVORNE OTPREMNICE, a
' stavke racuna iz njih. Prva verzija je primala gotove stavke, pa je bilo
' legalno napraviti zbirnu od izmisljenih 400+600 kg bez ijedne otpremnice --
' kes koji se ne slaze sa izvorom, i to kroz kanonski writer.
'
' Iz istog razloga VrstaVoca / SortaVoca / TipAmbalaze NISU u headeru: dolaze iz
' otpremnica, koje moraju biti saglasne. Dokument ima jednu vrstu, jednu sortu i
' jedan tip ambalaze (REFAKTOR_DOKUMENT_HEADER_STAVKE.md S2).
'
' CLANSTVO JE JEDINA VEZA, I UPISUJE SE U ISTOJ TRANSAKCIJI.
'
' tblZbirnaIzvori je JEDINI zapis clanstva -- "od kojih je otpremnica ova VERZIJA
' sastavljena" (A15). Nema pratioca -- ni jedne kolone na otpremnici.
'
' Ranija verzija je uz clanstvo drzala i Otpremnica.ZbirnaID kao kes. To je bilo
' jedno jeftinije citanje po ceni cele nove klase problema: drift izmedju kanona
' i kesa, provera tog drifta, snapshot jos jedne tabele, jos jedan upis i jos
' dva testa. Bez produkcionih podataka nema nikoga kome to placamo.
'
' Pitanje "na kojoj je aktivnoj zbirnoj otpremnica sada" racuna se iz clanstva --
' AktivnaZbirnaZaOtpremnicu.
'
' ZbirnaID je OPAQUE (NewEntityID), ne GetNextID: broj vise nije identitet, pa
' ni ID ne sme da bude brojac po kome se pogadja "sledeci".
'
' GeneracijaID se NE pise. Ta masinerija je kompenzacija za nepostojeci header i
' brise se u Zbirna cutover-u; nov pisac je ne sme ozivljavati.
'
' PR3 je ADITIVAN. Zatecen tok (modDokUnos.bas:533, modAutoHladnjaca,
' modMasterSync) i dalje idu starim putem; cutover citalaca, invarijante i storna
' je Zbirna cutover. Zato header NAMERNO ostavlja UkupnoKolicina /
' UkupnoAmbalaze / Klasa
' prazne -- to su kolone koje u ciljnoj semi ne postoje jer kolicina zivi na
' stavci. Ko ih procita dobija prazno, i to je tacan odgovor: nije "nula
' kilograma", nego "ne pitaj header za kolicinu".
'
' BrojZbirne se pise na HEADER zbirne -- to je poslovna labela dokumenta i
' operater je vidi na papiru. Ono sto se ne radi: broj se ne koristi kao VEZA.
' Pripadnost otpremnice zbirnoj zna iskljucivo tblZbirnaIzvori; broj nije ni
' relacija ni rezervni put.
'
' OVAJ WRITER PRAVI I ODMAH FINALIZUJE DOKUMENT (IzdatoStatus = IZDATO).
'
' Zato trazi bar jednu izvornu otpremnicu: zbirna bez izvora nije izdat dokument.
' Draft-first tok -- gde operater prvo otvori zbirnu pa vezuje otpremnice -- je
' zasebna funkcija koja jos ne postoji (CreateZbirnaDraft_TX). Dok je nema,
' kanonski tok je JEDAN: otpremnice postoje, pa se zbirna napravi i izda.
'
' DVA JAVNA ULAZA, JEDNO JEZGRO -- namera se vidi na callsite-u.
'
'   CreateZbirna_TX           rucni poslovni unos. "ocekivano" je OBAVEZNO:
'                             ono sto je operater otkucao mora da se poredi sa
'                             izvedenim iz otpremnica.
'   CreateZbirnaIzIzvora_TX   automatski tok (auto-hladnjaca, malina). Nema
'                             nezavisnog ocekivanja, i to se KAZE.
'
' Ranije je bio jedan ulaz sa Optional ocekivano, pa se kontrola mogla iskljuciti
' time sto se argument prosto ne prosledi -- tiho, bez traga na pozivu. Sada se
' ne moze iskljuciti; moze se samo izabrati drugi ulaz, i to se vidi.
'
' Argumenti:
'   h                 Scripting.Dictionary. Obavezno: Datum, VozacID,
'                     BrojZbirne, KupacID. Opciono: Hladnjaca, Pogon.
'                     Nepoznat kljuc je GRESKA (v. HdrProveriKljuceve).
'   izvorOtpremnice   Collection OtpremnicaID-jeva. Bar jedan.
'   ocekivano         Collection diktova {Klasa, Kolicina, KolAmbalaze} -- ono
'                     sto je operater OTKUCAO. Neslaganje sa izvedenim obara upis.
'   outGreska         RAZLOG odbijanja, ne samo cinjenica.
'
' outGreska postoji jer se bez njega ne moze razlikovati "kapija je odbila upis"
' od "upis je pukao iz drugog razloga pa je ispalo isto". Nije teorijska
' razlika: sabotaza koja je iskljucila proveru duple klase ostavila je suite
' ZELEN, jer je posao preuzeo Dictionary.Add svojom greskom o duplom kljucu.
Public Function CreateZbirna_TX(ByVal h As Object, _
                                ByVal izvorOtpremnice As Collection, _
                                ByVal ocekivano As Collection, _
                                Optional ByRef outGreska As String) As String
    CreateZbirna_TX = ZbirnaUpis(h, izvorOtpremnice, ocekivano, True, outGreska)
End Function

' Izvedeno bez nezavisne kontrole -- za automatske tokove koji nemaju sta da
' unakrsno provere. Izricito, ne prece prosledjivanjem Nothing.
Public Function CreateZbirnaIzIzvora_TX(ByVal h As Object, _
                                        ByVal izvorOtpremnice As Collection, _
                                        Optional ByRef outGreska As String) As String
    CreateZbirnaIzIzvora_TX = ZbirnaUpis(h, izvorOtpremnice, Nothing, False, outGreska)
End Function

Private Function ZbirnaUpis(ByVal h As Object, _
                            ByVal izvorOtpremnice As Collection, _
                            ByVal ocekivano As Collection, _
                            ByVal ocekivanoObavezno As Boolean, _
                            ByRef outGreska As String) As String
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    outGreska = ""

    On Error GoTo EH

    ' Sema pre upisa: AppendRow pise POZICIONO, pa tabela sa kolonom manje ili u
    ' pogresnom rasporedu tiho salje vrednosti u pogresna polja. Ide PRE BeginTx:
    ' kapija sme da digne gresku, a nema smisla otvarati transakciju koja se
    ' odmah rollback-uje.
    ' tblOtpremnica se CITA, ne menja -- zato nije u snapshotu.
    modSchema.SchemaReadyOrFail "CreateZbirna_TX", _
        TBL_ZBIRNA & "|" & TBL_ZBIRNA_STAVKE & "|" & TBL_ZBIRNA_IZVORI & _
        "|" & TBL_OTPREMNICA

    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA
    tx.AddTableSnapshot TBL_ZBIRNA_STAVKE
    tx.AddTableSnapshot TBL_ZBIRNA_IZVORI

    ZbirnaUpis = CreateZbirna(h, izvorOtpremnice, ocekivano, ocekivanoObavezno)

    If ZbirnaUpis = "" Then
        Err.Raise vbObjectError + 1220, "CreateZbirna_TX", _
                  "CreateZbirna nije vratio ZbirnaID."
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
    LogError "CreateZbirna_TX", errDesc, errNum
    Monitor_Error _
        moduleName:="modDokumenta", _
        procedureName:="CreateZbirna_TX", _
        entityType:="Zbirna", _
        entityID:=ZbirnaUpis, _
        correlationId:=ZbirnaUpis, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="DOKUMENT_SAVE_FAIL", _
        severity:="ERROR", _
        message:="CreateZbirna_TX failed. Error=" & errDesc, _
        userId:="Operator", _
        moduleName:="modDokumenta", _
        procedureName:="CreateZbirna_TX", _
        entityType:="Zbirna", _
        entityID:=ZbirnaUpis, _
        correlationId:=ZbirnaUpis

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    ZbirnaUpis = ""
    outGreska = errDesc

    PrintTxFailure "CreateZbirna_TX", errSrc, errNum, errDesc
End Function

' Core -- NE zovi spolja. Ulazi su CreateZbirna_TX i CreateZbirnaIzIzvora_TX,
' oba preko ZbirnaUpis koji drzi snapshot transakciju; direktan poziv bi kod
' greske ostavio pola dokumenta.
Private Function CreateZbirna(ByVal h As Object, _
                              ByVal izvorOtpremnice As Collection, _
                              ByVal ocekivano As Collection, _
                              ByVal ocekivanoObavezno As Boolean) As String
    Const SRC As String = "CreateZbirna"

    On Error GoTo EH

    If h Is Nothing Then
        Err.Raise vbObjectError + 1221, SRC, "Header nije prosledjen."
    End If

    If izvorOtpremnice Is Nothing Then
        Err.Raise vbObjectError + 1222, SRC, _
                  "Izvorne otpremnice nisu prosledjene."
    End If

    If izvorOtpremnice.count = 0 Then
        Err.Raise vbObjectError + 1223, SRC, _
                  "Zbirna mora imati bar jednu izvornu otpremnicu."
    End If

    ' Rucni unos MORA da donese ono sto je operater otkucao. Prazna kolekcija je
    ' isto sto i nijedna -- inace bi se kontrola gasila praznim argumentom.
    If ocekivanoObavezno Then
        If ocekivano Is Nothing Then
            Err.Raise vbObjectError + 1261, SRC, _
                      "Ocekivane vrednosti su obavezne za rucni unos. Za " & _
                      "automatski tok koristi CreateZbirnaIzIzvora_TX."
        End If
        If ocekivano.count = 0 Then
            Err.Raise vbObjectError + 1262, SRC, _
                      "Ocekivane vrednosti su prazne. Za automatski tok " & _
                      "koristi CreateZbirnaIzIzvora_TX."
        End If
    End If

    ' Fail-fast nad semom pre ijednog upisa.
    RequireColumnIndex TBL_ZBIRNA, COL_ZBR_ID, SRC
    RequireColumnIndex TBL_ZBIRNA, COL_ZBR_DATUM, SRC
    RequireColumnIndex TBL_ZBIRNA, COL_ZBR_VOZAC, SRC
    RequireColumnIndex TBL_ZBIRNA, COL_ZBR_BROJ, SRC
    RequireColumnIndex TBL_ZBIRNA, COL_ZBR_KUPAC, SRC
    RequireColumnIndex TBL_ZBIRNA, COL_ZBR_TIP_AMB, SRC

    RequireColumnIndex TBL_ZBIRNA_STAVKE, COL_ZBS_ID, SRC
    RequireColumnIndex TBL_ZBIRNA_STAVKE, COL_ZBS_ZBIRNA_ID, SRC
    RequireColumnIndex TBL_ZBIRNA_STAVKE, COL_ZBS_RB, SRC
    RequireColumnIndex TBL_ZBIRNA_STAVKE, COL_ZBS_KLASA, SRC
    RequireColumnIndex TBL_ZBIRNA_STAVKE, COL_ZBS_KOLICINA, SRC
    RequireColumnIndex TBL_ZBIRNA_STAVKE, COL_ZBS_KOL_AMB, SRC

    RequireColumnIndex TBL_ZBIRNA_IZVORI, COL_ZBI_ID, SRC
    RequireColumnIndex TBL_ZBIRNA_IZVORI, COL_ZBI_ZBIRNA_ID, SRC
    RequireColumnIndex TBL_ZBIRNA_IZVORI, COL_ZBI_OTPREMNICA_ID, SRC

    RequireColumnIndex TBL_OTPREMNICA, COL_OTP_ID, SRC

    HdrProveriKljuceve h, SRC

    Dim datum As Date
    Dim vozacID As String
    Dim brojZbirne As String
    Dim kupacID As String

    ZbrHdrCitajIProveri h, SRC, datum, vozacID, brojZbirne, kupacID

    modBrojevi.RequireBrojUKontekstu modBrojevi.KIND_ZBR, vozacID, datum, _
                                     brojZbirne, SRC
    modBrojevi.RequireBrojSlobodanUNizu modBrojevi.KIND_ZBR, vozacID, datum, _
                                        brojZbirne, SRC

    ' --- izvor: procitaj, proveri, izvedi ------------------------------------
    Dim data As Variant
    data = GetTableData(TBL_OTPREMNICA)
    If Not IsArray(data) Then
        Err.Raise vbObjectError + 1224, SRC, "Tabela otpremnica je prazna."
    End If

    Dim cID As Long
    Dim cVrsta As Long, cSorta As Long, cTip As Long
    Dim cStorno As Long, cVozac As Long

    cID = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_ID, SRC)
    cVrsta = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VRSTA, SRC)
    cSorta = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_SORTA, SRC)
    cTip = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_TIP_AMB, SRC)
    cVozac = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VOZAC, SRC)
    cStorno = RequireColumnIndex(TBL_OTPREMNICA, COL_STORNIRANO, SRC)

    ' Kanonsko clanstvo se cita JEDNOM, pre petlje.
    Dim clanstvo As Object
    Set clanstvo = AktivnoClanstvoPoKanonu(SRC)

    ' Klasa, kolicina i gajbe otpremnice su na STAVKAMA (S3b-1) -- zaglavlje ih
    ' od S3a ne nosi. Citaju se JEDNOM, kroz kanonski citalac koji zaglavlje bez
    ' stavki odbija po imenu. Da li zbirna sme da primi samo IZDATU otpremnicu
    ' odlucuje S4; ovde se menja samo odakle dolaze kilogrami.
    Dim stavkeOtp As Object
    Set stavkeOtp = StavkeOtpremnicePoDokumentu()
    Dim stavkeJedne As Collection, sOtp As Variant, sK As Long

    Dim redPoID As Object
    Dim kolPoKlasi As Object
    Dim ambPoKlasi As Object
    Set redPoID = CreateObject("Scripting.Dictionary")
    Set kolPoKlasi = CreateObject("Scripting.Dictionary")
    Set ambPoKlasi = CreateObject("Scripting.Dictionary")

    Dim vrsta As String, sorta As String, tipAmb As String
    Dim prvi As Boolean
    Dim i As Long, r As Long
    Dim otpID As String, klasa As String
    Dim kolRed As Double, ambRed As Double

    prvi = True

    For i = 1 To izvorOtpremnice.count
        otpID = Trim$(NzToText(izvorOtpremnice(i)))
        If Len(otpID) = 0 Then
            Err.Raise vbObjectError + 1225, SRC, _
                      "Prazan OtpremnicaID na poziciji " & CStr(i) & "."
        End If

        If redPoID.Exists(UCase$(otpID)) Then
            Err.Raise vbObjectError + 1226, SRC, _
                      "Ista otpremnica navedena dvaput: " & otpID
        End If

        r = NadjiJedanRedOtpremnice(data, cID, otpID, SRC)
        redPoID.Add UCase$(otpID), r

        ' Ista definicija valjanog izvora koju koristi i put nacrta: nije
        ' stornirana I jeste izdata (review #372, P1). Do tada je ovaj ulaz
        ' primao i DRAFT otpremnicu -- dakle robu koja nije otisla.
        RequireOtpValidanIzvorZbirne data, r, otpID, SRC

        ' Fail-closed na vec vezanu otpremnicu. Jedini izvor je clanstvo.
        '
        ' BrojZbirne se NE gleda: on nikad nije bio veza nego labela, a stari
        ' writer koji ga puni nema sta da stiti -- nema produkcionih podataka.
        ' Nov kanonski writer ne sme da poznaje stari model odnosa.
        If clanstvo.Exists(UCase$(otpID)) Then
            Err.Raise vbObjectError + 1228, SRC, _
                      "Otpremnica je vec u sastavu aktivne zbirne: " & otpID & _
                      " -> " & CStr(clanstvo(UCase$(otpID)))
        End If

        ' Zbirna je JEDAN transport JEDNOG vozaca. Otpremnica drugog vozaca u
        ' istoj zbirnoj je korupcija domena, ne rubni slucaj: BrojZbirne je
        ' scoped po vozacu, pa bi takav dokument bio i nepretraziv.
        RequireIstoPolje vozacID, Trim$(NzToText(data(r, cVozac))), _
                         "VozacID", otpID, SRC

        ' Dokument ima JEDNU vrstu, sortu i tip ambalaze. Otpremnica koja se ne
        ' slaze ne pripada ovoj zbirnoj -- i to je strukturno, ne stvar ukusa.
        If prvi Then
            vrsta = Trim$(NzToText(data(r, cVrsta)))
            sorta = Trim$(NzToText(data(r, cSorta)))
            tipAmb = Trim$(NzToText(data(r, cTip)))
            prvi = False
        Else
            RequireIstoPolje vrsta, Trim$(NzToText(data(r, cVrsta))), _
                             "VrstaVoca", otpID, SRC
            RequireIstoPolje sorta, Trim$(NzToText(data(r, cSorta))), _
                             "SortaVoca", otpID, SRC
            RequireIstoPolje tipAmb, Trim$(NzToText(data(r, cTip))), _
                             "TipAmbalaze", otpID, SRC
        End If

        Set stavkeJedne = StavkeZaOtpremnicu(stavkeOtp, otpID, SRC)
        For sK = 1 To stavkeJedne.count
            sOtp = stavkeJedne(sK)
            klasa = Trim$(NzToText(sOtp(3)))
            RequireValidDocumentClass klasa, SRC
            klasa = UCase$(klasa)

            kolRed = CDbl(sOtp(4))
            ambRed = CDbl(sOtp(6))

            If Not kolPoKlasi.Exists(klasa) Then
                kolPoKlasi.Add klasa, 0#
                ambPoKlasi.Add klasa, 0#
            End If
            kolPoKlasi(klasa) = CDbl(kolPoKlasi(klasa)) + kolRed
            ambPoKlasi(klasa) = CDbl(ambPoKlasi(klasa)) + ambRed
        Next sK
    Next i

    ' --- izvedene stavke moraju biti smislene --------------------------------
    Dim klase As Collection
    Set klase = KlaseUKanonskomRedu(kolPoKlasi)

    Dim k As Long
    For k = 1 To klase.count
        klasa = CStr(klase(k))
        If CDbl(kolPoKlasi(klasa)) <= 0 Then
            Err.Raise vbObjectError + 1231, SRC, _
                      "Zbir kolicine za klasu " & klasa & " nije veci od nule."
        End If
        If CDbl(ambPoKlasi(klasa)) < 0 Then
            Err.Raise vbObjectError + 1232, SRC, _
                      "Zbir ambalaze za klasu " & klasa & " je negativan."
        End If
        ' Ambalaza je BROJ KOMADA, ne tezina. Legacy validator unosa je to
        ' drzao tipom (ukupnoAmb As Long); rec/nista drugo to vise ne cuva, pa
        ' se trazi ovde. Odbija se, ne zaokruzuje: "20.5 gajbica" je kvar u
        ' izvoru, a tiha ispravka ga sakriva.
        RequireCeoBroj CDbl(ambPoKlasi(klasa)), _
                       "Ambalaza za klasu " & klasa, SRC
    Next k

    RequireOcekivanoSeSlaze ocekivano, kolPoKlasi, ambPoKlasi, SRC

    ' --- upis ----------------------------------------------------------------
    Dim zbirnaID As String
    zbirnaID = NewEntityID("ZBR-")

    If zbirnaID = "" Then
        Err.Raise vbObjectError + 1233, SRC, _
                  "NewEntityID nije vratio ZbirnaID."
    End If

    Dim rowData As Variant
    rowData = BuildZbirnaHeaderRowData(zbirnaID, datum, vozacID, brojZbirne, _
                                       kupacID, HdrOpcion(h, "Hladnjaca"), _
                                       HdrOpcion(h, "Pogon"), vrsta, sorta, _
                                       tipAmb)

    If AppendRow(TBL_ZBIRNA, rowData) <= 0 Then
        Err.Raise vbObjectError + 1234, SRC, _
                  "AppendRow nije upisao header u tblZbirna."
    End If

    Dim stavkaID As String
    For k = 1 To klase.count
        klasa = CStr(klase(k))

        ' Fail-closed: NewEntityID vraca "" kad CoCreateGuid ne uspe. Red bez
        ' identiteta je gori od pada -- niko ga posle ne moze ni naci ni vezati.
        stavkaID = NewEntityID("ZBS-")
        If stavkaID = "" Then
            Err.Raise vbObjectError + 1235, SRC, _
                      "NewEntityID nije vratio ZbirnaStavkaID za klasu " & klasa & "."
        End If

        rowData = BuildZbirnaStavkaRowData(stavkaID, zbirnaID, k, klasa, _
                                           CDbl(kolPoKlasi(klasa)), _
                                           CDbl(ambPoKlasi(klasa)))

        If AppendRow(TBL_ZBIRNA_STAVKE, rowData) <= 0 Then
            Err.Raise vbObjectError + 1236, SRC, _
                      "AppendRow nije upisao stavku za klasu " & klasa & "."
        End If
    Next k

    ' --- clanstvo, u ISTOJ transakciji kao header i stavke -------------------
    '
    ' Od kojih je otpremnica ova VERZIJA sastavljena. Posle ispravke jedne
    ' otpremnice nastaje nova verzija zbirne, a sestre koje se nisu menjale
    ' pripadaju i staroj i novoj -- jedan FK to ne bi mogao da pokaze.
    Dim kljuc As Variant
    Dim izvorID As String

    For Each kljuc In redPoID.Keys
        izvorID = NewEntityID("ZBI-")
        If izvorID = "" Then
            Err.Raise vbObjectError + 1256, SRC, _
                      "NewEntityID nije vratio ZbirnaIzvorID."
        End If

        rowData = BuildZbirnaIzvorRowData(izvorID, zbirnaID, _
                                          IDIzRedaOtpremnice(data, cID, _
                                                             CLng(redPoID(kljuc))))
        If AppendRow(TBL_ZBIRNA_IZVORI, rowData) <= 0 Then
            Err.Raise vbObjectError + 1257, SRC, _
                      "AppendRow nije upisao clanstvo otpremnice."
        End If
    Next kljuc

    CreateZbirna = zbirnaID
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

' Tacno jedan red za dati OtpremnicaID. Nula i vise od jednog su oba greska:
' tihi "uzmi prvi pogodak" je klasa buga zbog koje CreateFaktura ima
' RequireSingleFakturaRow.
Private Function NadjiJedanRedOtpremnice(ByRef data As Variant, _
                                         ByVal colID As Long, _
                                         ByVal otpID As String, _
                                         ByVal src As String) As Long
    Dim i As Long, nadjen As Long, koliko As Long

    For i = 1 To UBound(data, 1)
        If StrComp(Trim$(NzToText(data(i, colID))), otpID, vbTextCompare) = 0 Then
            nadjen = i
            koliko = koliko + 1
        End If
    Next i

    If koliko = 0 Then
        Err.Raise vbObjectError + 1237, src, _
                  "Otpremnica nije pronadjena: " & otpID
    End If
    If koliko > 1 Then
        Err.Raise vbObjectError + 1238, src, _
                  "Duplikat OtpremnicaID=" & otpID & "; Count=" & CStr(koliko)
    End If

    NadjiJedanRedOtpremnice = nadjen
End Function

Private Sub RequireIstoPolje(ByVal ocekivano As String, ByVal stvarno As String, _
                             ByVal polje As String, ByVal otpID As String, _
                             ByVal src As String)
    If StrComp(ocekivano, stvarno, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 1239, src, _
                  "Izvor " & otpID & " ima drugo polje " & polje & _
                  ": ocekivano '" & ocekivano & "', naslo '" & stvarno & "'."
    End If
End Sub

Private Sub RequireCeoBroj(ByVal v As Double, ByVal opis As String, _
                           ByVal src As String)
    If Abs(v - Fix(v)) > 0.0000001 Then
        Err.Raise vbObjectError + 1240, src, _
                  opis & " mora biti ceo broj, a nije: " & Fmt2Zbr(v)
    End If
End Sub

' Klase u kanonskom redu: I, pa II, pa sve ostalo azbucno. RedniBroj mora biti
' determinisan -- Dictionary.Keys cuva redosled ubacivanja, a on zavisi od
' redosleda otpremnica u pozivu.
'
' Public zbog modOtkup.CreateOtkup: tamo redosled ubacivanja zavisi od adaptera,
' a ista poslovna cinjenica mora dati isti dokument. Ulaz je Dictionary ciji su
' KLJUCEVI UCase-ovane klase; vrednosti se ne citaju, pa isti poziv radi i nad
' mapom klasa->kolicina (zbirna) i nad mapom klasa->indeks stavke (otkup).
Public Function KlaseUKanonskomRedu(ByVal kolPoKlasi As Object) As Collection
    Dim c As Collection
    Set c = New Collection

    If kolPoKlasi.Exists(UCase$(KLASA_I)) Then c.Add UCase$(KLASA_I)
    If kolPoKlasi.Exists(UCase$(KLASA_II)) Then c.Add UCase$(KLASA_II)

    Dim kljuc As Variant
    Dim ostatak As Collection
    Set ostatak = New Collection

    For Each kljuc In kolPoKlasi.Keys
        If StrComp(CStr(kljuc), UCase$(KLASA_I), vbTextCompare) <> 0 And _
           StrComp(CStr(kljuc), UCase$(KLASA_II), vbTextCompare) <> 0 Then
            ostatak.Add CStr(kljuc)
        End If
    Next kljuc

    Dim i As Long, j As Long, tmp As String
    Dim arr() As String
    If ostatak.count > 0 Then
        ReDim arr(1 To ostatak.count)
        For i = 1 To ostatak.count
            arr(i) = CStr(ostatak(i))
        Next i
        For i = 1 To UBound(arr) - 1
            For j = i + 1 To UBound(arr)
                If StrComp(arr(i), arr(j), vbTextCompare) > 0 Then
                    tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
                End If
            Next j
        Next i
        For i = 1 To UBound(arr)
            c.Add arr(i)
        Next i
    End If

    Set KlaseUKanonskomRedu = c
End Function

' Unakrsna provera: ono sto je operater otkucao mora da se slaze sa onim sto je
' izvedeno iz otpremnica. Neslaganje je danas tiha greska -- ekran prikaze svoje
' brojeve, tabela nosi druge.
Private Sub RequireOcekivanoSeSlaze(ByVal ocekivano As Collection, _
                                    ByVal kolPoKlasi As Object, _
                                    ByVal ambPoKlasi As Object, _
                                    ByVal src As String)
    If ocekivano Is Nothing Then Exit Sub
    If ocekivano.count = 0 Then Exit Sub

    Dim vidjene As Object
    Set vidjene = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim s As Object
    Dim klasa As String
    Dim kol As Double, amb As Double

    For i = 1 To ocekivano.count
        If Not IsObject(ocekivano(i)) Then
            Err.Raise vbObjectError + 1241, src, _
                      "Ocekivana stavka " & CStr(i) & " nije Dictionary."
        End If
        Set s = ocekivano(i)

        klasa = UCase$(Trim$(NzToText(StavkaVrednost(s, "Klasa", i, src))))
        If vidjene.Exists(klasa) Then
            Err.Raise vbObjectError + 1242, src, _
                      "Dve ocekivane stavke iste klase: " & klasa
        End If
        vidjene.Add klasa, True

        If Not kolPoKlasi.Exists(klasa) Then
            Err.Raise vbObjectError + 1243, src, _
                      "Ocekivana klasa " & klasa & " ne postoji na izvornim otpremnicama."
        End If

        kol = StavkaBroj(s, "Kolicina", i, src)
        amb = StavkaBroj(s, "KolAmbalaze", i, src)

        If Abs(kol - CDbl(kolPoKlasi(klasa))) > 0.001 Then
            Err.Raise vbObjectError + 1244, src, _
                      "Kolicina za klasu " & klasa & " se ne slaze sa otpremnicama: " & _
                      "uneto " & Fmt2Zbr(kol) & ", izvedeno " & Fmt2Zbr(CDbl(kolPoKlasi(klasa))) & "."
        End If
        If Abs(amb - CDbl(ambPoKlasi(klasa))) > 0.001 Then
            Err.Raise vbObjectError + 1245, src, _
                      "Ambalaza za klasu " & klasa & " se ne slaze sa otpremnicama: " & _
                      "uneto " & Fmt2Zbr(amb) & ", izvedeno " & Fmt2Zbr(CDbl(ambPoKlasi(klasa))) & "."
        End If
    Next i

    If vidjene.count <> kolPoKlasi.count Then
        Err.Raise vbObjectError + 1246, src, _
                  "Ocekivano ima " & CStr(vidjene.count) & " klasa, a otpremnice daju " & _
                  CStr(kolPoKlasi.count) & "."
    End If
End Sub

' Broj u poruku, nezavisno od Windows locale-a (decimalna tacka uvek).
Private Function Fmt2Zbr(ByVal v As Double) As String
    Fmt2Zbr = Replace(Format$(v, "0.00"), ",", ".")
End Function

' Header ciljne seme: BEZ UkupnoKolicina / UkupnoAmbalaze / Klasa. Te kolone jos
' postoje u tabeli (brisu se u cutover-u) i ostaju prazne namerno -- v. gore
' CreateZbirna_TX. GeneracijaID se ne pise.
Private Function BuildZbirnaHeaderRowData(ByVal zbirnaID As String, _
                                          ByVal datum As Date, _
                                          ByVal vozacID As String, _
                                          ByVal brojZbirne As String, _
                                          ByVal kupacID As String, _
                                          ByVal hladnjaca As String, _
                                          ByVal pogon As String, _
                                          ByVal vrstaVoca As String, _
                                          ByVal sortaVoca As String, _
                                          ByVal tipAmb As String, _
                                          Optional ByVal izdatoStatus As String = "") As Variant
    Const SRC As String = "BuildZbirnaHeaderRowData"

    Dim colCount As Long
    colCount = TabelaBrojKolona(TBL_ZBIRNA)

    If colCount <= 0 Then
        Err.Raise vbObjectError + 1247, SRC, _
                  "Ne mogu da odredim broj kolona za tblZbirna."
    End If

    Dim rowData() As Variant
    ReDim rowData(0 To colCount - 1)

    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_ID, zbirnaID, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_DATUM, datum, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_VOZAC, vozacID, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_BROJ, brojZbirne, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_KUPAC, kupacID, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_HLADNJACA, hladnjaca, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_POGON, pogon, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_VRSTA, vrstaVoca, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_SORTA, sortaVoca, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_TIP_AMB, tipAmb, SRC

    If GetColumnIndex(TBL_ZBIRNA, COL_STORNIRANO) > 0 Then
        SetRowValueByColumn rowData, TBL_ZBIRNA, COL_STORNIRANO, "", SRC
    End If

    ' IZDATO se pise EKSPLICITNO.
    '
    ' Legacy konvencija je "prazno = IZDATO" (modDokumentInvariant.DocIsIssued,
    ' konzervativno). Nov model se na nju ne oslanja: ovaj writer pravi i odmah
    ' FINALIZUJE dokument, pa to i kaze. Kad se pojavi draft-first tok, on ce
    ' imati svoj ulaz (CreateZbirnaDraft_TX) i pisati DRAFT -- a razlika izmedju
    ' "niko nije upisao" i "izdato" tada vise nije pretpostavka.
    If GetColumnIndex(TBL_ZBIRNA, COL_TRACE_IZDATO_STATUS) > 0 Then
        Dim st As String
        st = Trim$(izdatoStatus)
        If Len(st) = 0 Then st = IZDATO_IZDATO
        SetRowValueByColumn rowData, TBL_ZBIRNA, COL_TRACE_IZDATO_STATUS, st, SRC
    End If

    BuildZbirnaHeaderRowData = rowData
End Function

' Na kojoj je AKTIVNOJ zbirnoj otpremnica sada. "" = ni na jednoj.
'
' Ovo je zamena za obrisanu kolonu Otpremnica.ZbirnaID: isto pitanje, ali
' racunato iz jedinog zapisa clanstva umesto cuvano na drugom mestu.
Public Function AktivnaZbirnaZaOtpremnicu(ByVal otpremnicaID As String) As String
    Const SRC As String = "AktivnaZbirnaZaOtpremnicu"

    Dim mapa As Object
    Set mapa = AktivnoClanstvoPoKanonu(SRC)

    Dim kljuc As String
    kljuc = UCase$(Trim$(otpremnicaID))
    If mapa.Exists(kljuc) Then AktivnaZbirnaZaOtpremnicu = CStr(mapa(kljuc))
End Function

' Kanonsko clanstvo: UCase(OtpremnicaID) -> ZbirnaID, samo za AKTIVNE zbirne.
'
' Posle A13 ista otpremnica sme da ima VISE zapisa clanstva -- po jedan za svaku
' verziju zbirne kroz koju je prosla. Zauzeta je samo ako je clan zbirne koja
' NIJE stornirana; clanstvo u superseded verziji je istorija, ne prepreka.
'
' DVA AKTIVNA ZAPISA ZA ISTU OTPREMNICU SU TVRDA GRESKA.
'
' Pravilo je prosto: jedna otpremnica ima TACNO 0 ili 1 aktivan zapis clanstva.
' Drugi zapis je greska bez obzira da li pokazuje na DRUGU ili na ISTU zbirnu --
' dupli red iste veze bi kasnije obican join sabrao dvaput.
'
' Prva verzija je radila prosto mapa(otpID) = zbrID, pa bi drugi red tiho
' pregazio prvi. Druga je hvatala samo razlicit ZbirnaID. Loader koji nelegalno
' stanje normalizuje u legalno radi protiv kardinaliteta koji A15 cuva.
Private Function AktivnoClanstvoPoKanonu(ByVal src As String) As Object
    Dim mapa As Object
    Set mapa = CreateObject("Scripting.Dictionary")
    Set AktivnoClanstvoPoKanonu = mapa

    Dim izv As Variant
    izv = GetTableData(TBL_ZBIRNA_IZVORI)
    If Not IsArray(izv) Then Exit Function

    Dim cIzvZbr As Long, cIzvOtp As Long
    cIzvZbr = RequireColumnIndex(TBL_ZBIRNA_IZVORI, COL_ZBI_ZBIRNA_ID, src)
    cIzvOtp = RequireColumnIndex(TBL_ZBIRNA_IZVORI, COL_ZBI_OTPREMNICA_ID, src)

    ' Skup storniranih zbirnih -- jedan prolaz, pa provera po kljucu.
    Dim stornirane As Object
    Set stornirane = CreateObject("Scripting.Dictionary")

    Dim zbr As Variant
    zbr = GetTableData(TBL_ZBIRNA)
    If IsArray(zbr) Then
        Dim cZbrPk As Long, cZbrSt As Long, k As Long
        cZbrPk = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_ID, src)
        cZbrSt = GetColumnIndex(TBL_ZBIRNA, COL_STORNIRANO)
        If cZbrSt > 0 Then
            For k = 1 To UBound(zbr, 1)
                If StrComp(Trim$(NzToText(zbr(k, cZbrSt))), "Da", vbTextCompare) = 0 Then
                    stornirane(UCase$(Trim$(NzToText(zbr(k, cZbrPk))))) = True
                End If
            Next k
        End If
    End If

    Dim i As Long
    Dim zbrID As String, otpID As String

    For i = 1 To UBound(izv, 1)
        zbrID = Trim$(NzToText(izv(i, cIzvZbr)))
        otpID = UCase$(Trim$(NzToText(izv(i, cIzvOtp))))
        If Len(zbrID) > 0 And Len(otpID) > 0 Then
            If Not stornirane.Exists(UCase$(zbrID)) Then
                If mapa.Exists(otpID) Then
                    Err.Raise vbObjectError + 1260, src, _
                              "Kanonsko clanstvo je nekonzistentno: otpremnica " & _
                              otpID & " ima dva aktivna zapisa clanstva (" & _
                              CStr(mapa(otpID)) & " i " & zbrID & ")."
                End If
                mapa.Add otpID, zbrID
            End If
        End If
    Next i
End Function

' =====================================================================
' NACRT ZBIRNE (S4-2b)
' =====================================================================
'
' Odluka operatera (21.09.2026): zbirna mora da radi na OBA nacina, isto kao
' otpremnica:
'
'   1) NACRT PA POKRIVANJE -- operater prvo najavi sta zbirna nosi (kilaza i
'      gajbe po klasi), pa dodaje izdate otpremnice dok najava ne bude pokrivena.
'      Izdavanje trazi TACNU jednakost.
'   2) JEDAN POTEZ -- operater izabere otpremnice, a ocekivanje se izvodi iz njih
'      (CreateZbirnaIzIzvora_TX, postoji od PR3).
'
' Razlika prema otpremnici je samo u tome STA je izvor: otpremnici su izvor
' otkupni blokovi, zbirnoj IZDATE OTPREMNICE. Nacrt otpremnice nije roba koja je
' otisla, pa ne moze da bude deo zbirne -- odluka koju je S14.14 ostavila S4.
' PROSLEDJENO se racuna kao izdato (review #362): sync ne menja cinjenicu da je
' roba otpremljena.
'
' ULAZI -- svi osim izdavanja rade ISKLJUCIVO nad nacrtom:
'   CreateZbirnaDraft_TX(h, ocekivano)         -> ZBR-...  DRAFT + stavke
'   UpdateZbirnaDraft_TX(zbrID, h, ocekivano)  izmena najave, clanstvo netaknuto
'   DodajZbirnaIzvor_TX(zbrID, otpID)          clanstvo +1
'   UkloniZbirnaIzvor_TX(zbrID, otpID)         clanstvo -1
'   IzdajZbirnu_TX(zbrID)                      revalidacija + jednakost + IZDATO

Public Function CreateZbirnaDraft_TX(ByVal h As Object, _
                                     ByVal ocekivano As Collection, _
                                     Optional ByRef outGreska As String) As String
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    outGreska = ""

    On Error GoTo EH

    modSchema.SchemaReadyOrFail "CreateZbirnaDraft_TX", _
        TBL_ZBIRNA & "|" & TBL_ZBIRNA_STAVKE

    ' Isti razlog kao kod otpremnice: Nothing je privatan signal jednopoteznog
    ' puta i ne sme da procuri u rucni API -- inace nastaje DRAFT bez ocekivanja,
    ' dakle draft koji nema sta da meri a izgleda ispravno.
    If ocekivano Is Nothing Then
        Err.Raise vbObjectError + 1340, "CreateZbirnaDraft_TX", _
                  "Ocekivanje nije prosledjeno. Rucni nacrt mora da prijavi sta " & _
                  "zbirna nosi; za jedan potez koristi CreateZbirnaIzIzvora_TX."
    End If

    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA
    tx.AddTableSnapshot TBL_ZBIRNA_STAVKE

    CreateZbirnaDraft_TX = ZbrNapraviDraft(h, ocekivano)

    If CreateZbirnaDraft_TX = "" Then
        Err.Raise vbObjectError + 1341, "CreateZbirnaDraft_TX", _
                  "ZbrNapraviDraft nije vratio ZbirnaID."
    End If

    tx.CommitTx
    Set tx = Nothing
    Exit Function

EH:
    outGreska = ZbrPadTransakcije(tx, "CreateZbirnaDraft_TX", CreateZbirnaDraft_TX)
    CreateZbirnaDraft_TX = ""
End Function

Public Function UpdateZbirnaDraft_TX(ByVal zbirnaID As String, _
                                     ByVal h As Object, _
                                     ByVal ocekivano As Collection, _
                                     Optional ByRef outGreska As String) As Boolean
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    outGreska = ""

    On Error GoTo EH

    ' Otpremnice ulaze u spisak jer izmena zaglavlja REVALIDIRA postojece
    ' clanstvo (v. ZbrIzmeniDraft), pa cita i njih.
    modSchema.SchemaReadyOrFail "UpdateZbirnaDraft_TX", _
        TBL_ZBIRNA & "|" & TBL_ZBIRNA_STAVKE & "|" & TBL_ZBIRNA_IZVORI & _
        "|" & TBL_OTPREMNICA & "|" & TBL_OTPREMNICA_STAVKE

    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA
    tx.AddTableSnapshot TBL_ZBIRNA_STAVKE

    ZbrIzmeniDraft zbirnaID, h, ocekivano

    tx.CommitTx
    Set tx = Nothing
    UpdateZbirnaDraft_TX = True
    Exit Function

EH:
    outGreska = ZbrPadTransakcije(tx, "UpdateZbirnaDraft_TX", zbirnaID)
    UpdateZbirnaDraft_TX = False
End Function

Public Function DodajZbirnaIzvor_TX(ByVal zbirnaID As String, _
                                    ByVal otpremnicaID As String, _
                                    Optional ByRef outGreska As String) As Boolean
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    outGreska = ""

    On Error GoTo EH

    modSchema.SchemaReadyOrFail "DodajZbirnaIzvor_TX", _
        TBL_ZBIRNA & "|" & TBL_ZBIRNA_IZVORI & "|" & TBL_OTPREMNICA & _
        "|" & TBL_OTPREMNICA_STAVKE

    tx.BeginTx
    ' Zaglavlje ULAZI u snapshot: prvi izvor definise vrstu, sortu i tip
    ' ambalaze, pa dodavanje menja i njega. Bez ovoga bi pad posle preuzimanja
    ' cinjenica ostavio zaglavlje popunjeno, a clanstvo prazno.
    tx.AddTableSnapshot TBL_ZBIRNA
    tx.AddTableSnapshot TBL_ZBIRNA_IZVORI

    ZbrRequireIzvorValjan zbirnaID, Trim$(otpremnicaID), "ZbrDodajIzvor", True
    ZbrPreuzmiCinjeniceZaIzvor zbirnaID, Trim$(otpremnicaID), "ZbrDodajIzvor"
    ZbrUpisiClanstvo zbirnaID, Trim$(otpremnicaID), "ZbrDodajIzvor"

    tx.CommitTx
    Set tx = Nothing
    DodajZbirnaIzvor_TX = True
    Exit Function

EH:
    outGreska = ZbrPadTransakcije(tx, "DodajZbirnaIzvor_TX", zbirnaID)
    DodajZbirnaIzvor_TX = False
End Function

Public Function UkloniZbirnaIzvor_TX(ByVal zbirnaID As String, _
                                     ByVal otpremnicaID As String, _
                                     Optional ByRef outGreska As String) As Boolean
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    outGreska = ""

    On Error GoTo EH

    modSchema.SchemaReadyOrFail "UkloniZbirnaIzvor_TX", _
        TBL_ZBIRNA & "|" & TBL_ZBIRNA_IZVORI

    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA_IZVORI
    ' Zaglavlje ULAZI u snapshot: uklanjanje POSLEDNJEG izvora brise preuzetu
    ' vrstu, sortu i tip ambalaze (ZBR-KANON-04), pa i ono ucestvuje u izmeni.
    tx.AddTableSnapshot TBL_ZBIRNA

    Dim rZbr As Long
    rZbr = ZbrRedHeadera(zbirnaID, "ZbrUkloniIzvor")
    RequireZbrDraft zbirnaID, rZbr, "ZbrUkloniIzvor"

    ZbrUkloniIzvor zbirnaID, Trim$(otpremnicaID), "ZbrUkloniIzvor"
    ZbrOcistiCinjeniceBezClanstva zbirnaID, "ZbrUkloniIzvor"

    tx.CommitTx
    Set tx = Nothing
    UkloniZbirnaIzvor_TX = True
    Exit Function

EH:
    outGreska = ZbrPadTransakcije(tx, "UkloniZbirnaIzvor_TX", zbirnaID)
    UkloniZbirnaIzvor_TX = False
End Function

Public Function IzdajZbirnu_TX(ByVal zbirnaID As String, _
                               Optional ByRef outGreska As String) As Boolean
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    outGreska = ""

    On Error GoTo EH

    modSchema.SchemaReadyOrFail "IzdajZbirnu_TX", _
        TBL_ZBIRNA & "|" & TBL_ZBIRNA_STAVKE & "|" & TBL_ZBIRNA_IZVORI & _
        "|" & TBL_OTPREMNICA & "|" & TBL_OTPREMNICA_STAVKE

    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA

    ' Zbirna NE knjizi ambalazu: gajbe su vec knjizene pri izdavanju otpremnice,
    ' a knjizice se ponovo pri prijemu (S6). Zbirna je prevozni spisak, ne
    ' promena stanja gajbi -- zato ovde nema snapshota tblAmbalaza.
    ZbrIzdaj zbirnaID

    tx.CommitTx
    Set tx = Nothing
    IzdajZbirnu_TX = True
    Exit Function

EH:
    outGreska = ZbrPadTransakcije(tx, "IzdajZbirnu_TX", zbirnaID)
    IzdajZbirnu_TX = False
End Function

Public Function ZbirnaJeIzdata(ByVal zbirnaID As String) As Boolean
    zbirnaID = Trim$(zbirnaID)
    If Len(zbirnaID) = 0 Then Exit Function

    ZbirnaJeIzdata = IzdatoStatusJeIzdato(LookupValue( _
                         TBL_ZBIRNA, COL_ZBR_ID, zbirnaID, _
                         COL_TRACE_IZDATO_STATUS))
End Function

' Clanovi JEDNOG nacrta -- prazna kolekcija je uredno stanje (nacrt jos nije
' pokriven). Za IZDATU zbirnu prazno bi bio kvar, i to hvata IzvoriZbirne.
Public Function ZbrClanovi(ByVal zbirnaID As String) As Collection
    Const SRC As String = "ZbrClanovi"

    Dim c As Collection
    Set c = New Collection
    Set ZbrClanovi = c

    Dim izv As Variant
    izv = GetTableData(TBL_ZBIRNA_IZVORI)
    If Not IsArray(izv) Then Exit Function

    Dim cZbr As Long, cOtp As Long
    cZbr = RequireColumnIndex(TBL_ZBIRNA_IZVORI, COL_ZBI_ZBIRNA_ID, SRC)
    cOtp = RequireColumnIndex(TBL_ZBIRNA_IZVORI, COL_ZBI_OTPREMNICA_ID, SRC)

    Dim i As Long, zid As String
    zid = Trim$(zbirnaID)
    For i = 1 To UBound(izv, 1)
        If StrComp(Trim$(NzToText(izv(i, cZbr))), zid, vbTextCompare) = 0 Then
            c.Add Trim$(NzToText(izv(i, cOtp)))
        End If
    Next i
End Function

' NAPREDAK POKRIVANJA NACRTA -- ogledalo GetOtpremnicaProgress.
'
' Klase su UNIJA najavljenih i povezanih. Klasa koju izvori nose a nacrt je nije
' najavio MORA da se vidi kao visak, a ne da nestane iz prikaza: bez unije bi
' operater gledao uredan spisak i ne bi razumeo zasto izdavanje pada. Isto vazi
' obrnuto -- najavljena klasa koju nijedan izvor ne nosi ostaje sa preostalim.
'
' Cita se ISTIM parom koji izdavanje koristi za jednakost (ZbrUcitajOcekivano /
' ZbrUcitajPovezano), pa prikaz i kapija ne mogu da se raziju.
Public Function GetZbirnaProgress(ByVal zbirnaID As String) As Object
    Const SRC As String = "GetZbirnaProgress"

    Dim rez As Object
    Set rez = CreateObject("Scripting.Dictionary")
    Set GetZbirnaProgress = rez

    Dim ocek As Object, ocekAmb As Object
    Dim pov As Object, povAmb As Object
    ZbrUcitajOcekivano zbirnaID, ocek, ocekAmb, SRC
    ZbrUcitajPovezano zbirnaID, ZbrClanovi(zbirnaID), pov, povAmb, SRC

    Dim sve As Object
    Set sve = CreateObject("Scripting.Dictionary")

    Dim k As Variant
    For Each k In ocek.Keys
        sve(CStr(k)) = True
    Next k
    For Each k In pov.Keys
        sve(CStr(k)) = True
    Next k

    Dim red As Object
    For Each k In KlaseUKanonskomRedu(sve)
        Set red = CreateObject("Scripting.Dictionary")
        red("ocekivano") = OtpBroj(ocek, CStr(k))
        red("povezano") = OtpBroj(pov, CStr(k))
        red("preostalo") = OtpBroj(ocek, CStr(k)) - OtpBroj(pov, CStr(k))
        red("ocekivanoAmb") = OtpBroj(ocekAmb, CStr(k))
        red("povezanoAmb") = OtpBroj(povAmb, CStr(k))
        red("preostaloAmb") = OtpBroj(ocekAmb, CStr(k)) - OtpBroj(povAmb, CStr(k))
        Set rez(CStr(k)) = red
    Next k
End Function

' --- core: nacrt ------------------------------------------------------------
Private Function ZbrNapraviDraft(ByVal h As Object, _
                                 ByVal ocekivano As Collection) As String
    Const SRC As String = "ZbrNapraviDraft"

    HdrProveriKljuceve h, SRC

    Dim datum As Date
    Dim vozacID As String, brojZbirne As String, kupacID As String

    ZbrHdrCitajIProveri h, SRC, datum, vozacID, brojZbirne, kupacID

    modBrojevi.RequireBrojUKontekstu modBrojevi.KIND_ZBR, vozacID, datum, _
                                     brojZbirne, SRC
    modBrojevi.RequireBrojSlobodanUNizu modBrojevi.KIND_ZBR, vozacID, datum, _
                                        brojZbirne, SRC

    Dim zbirnaID As String
    zbirnaID = NewEntityID("ZBR-")
    If Len(zbirnaID) = 0 Then
        Err.Raise vbObjectError + 1363, SRC, "NewEntityID nije vratio ZbirnaID."
    End If

    ' VRSTA, SORTA I TIP AMBALAZE NE DOLAZE IZ ZAGLAVLJA (review #372, P1).
    '
    ' HdrProveriKljuceve ih izricito ne dozvoljava -- one su cinjenica ROBE, a
    ' robu donosi izvor. Nacrt zato krece prazan, a PRVI izvor ih definise
    ' (ZbrPreuzmiCinjenice); svaki sledeci mora da im se poklopi.
    '
    ' Prva verzija ovog koda ih je citala iz h i time uvek dobijala prazno, pa je
    ' zbirna ostajala bez vrste i sorte i posle izdavanja.
    Dim rowData As Variant
    rowData = BuildZbirnaHeaderRowData(zbirnaID, datum, vozacID, brojZbirne, _
                                       kupacID, HdrOpcion(h, "Hladnjaca"), _
                                       HdrOpcion(h, "Pogon"), _
                                       "", "", "", IZDATO_DRAFT)

    If AppendRow(TBL_ZBIRNA, rowData) <= 0 Then
        Err.Raise vbObjectError + 1343, SRC, "AppendRow za tblZbirna nije uspeo."
    End If

    ZbrUpisiOcekivano zbirnaID, ocekivano, SRC

    ZbrNapraviDraft = zbirnaID
End Function

' Ocekivanje -> stavke. Jedna stavka po klasi, u kanonskom redu, isti ugovor
' koji citalac (StavkeZbirneRedovi) posle trazi.
'
' STA JE VALJANO OCEKIVANJE -- jedna definicija za OBA ulaza (review #373, P1).
'
' Tvrdnja "nacrt ima bar jednu klasu" je do ovog reza stajala u ZbrNapraviDraft,
' dakle samo na putu nastanka. Izmena je proveravala jedino da kolekcija nije
' Nothing, pa je PRAZNA kolekcija prolazila: ZbrObrisiOcekivano obrise sve
' stavke, ovde nema nijedne iteracije, i commit ostavi zaglavlje BEZ IJEDNE
' STAVKE -- dokument koji RequireZaglavljaZbirneSaStavkama proglasava
' korumpiranim. Pisac ne sme da napravi stanje koje njegov citalac zabranjuje.
'
' Isti kvar koji je review #372 nasao kod izvora: dva ulaza, dve definicije istog
' pojma. Zato tvrdnja sada zivi tacno ovde -- na jedinom mestu kroz koje prolaze
' i nastanak i izmena -- a ne u svakom wrapper-u posebno.
Private Sub ZbrUpisiOcekivano(ByVal zbirnaID As String, _
                              ByVal ocekivano As Collection, _
                              ByVal src As String)
    If ocekivano Is Nothing Then
        Err.Raise vbObjectError + 1365, src, _
                  "Ocekivanje nije prosledjeno. Nacrt bez ocekivanja nema sta da meri."
    End If

    If ocekivano.count = 0 Then
        Err.Raise vbObjectError + 1342, src, _
                  "Ocekivanje je prazno. Nacrt mora da prijavi bar jednu klasu."
    End If

    Dim kolPoKlasi As Object, ambPoKlasi As Object
    Set kolPoKlasi = CreateObject("Scripting.Dictionary")
    Set ambPoKlasi = CreateObject("Scripting.Dictionary")

    Dim i As Long, s As Object, klasa As String
    For i = 1 To ocekivano.count
        Set s = ocekivano(i)
        klasa = UCase$(Trim$(NzToText(s("Klasa"))))
        RequireValidDocumentClass klasa, src

        If kolPoKlasi.Exists(klasa) Then
            Err.Raise vbObjectError + 1344, src, _
                      "Ocekivanje ima dve stavke iste klase: " & klasa
        End If

        If CDbl(s("Kolicina")) <= 0 Then
            Err.Raise vbObjectError + 1345, src, _
                      "Ocekivana kolicina za klasu " & klasa & " nije veca od nule."
        End If
        ' RequireCeoBroj meri SAMO celobrojnost: -10 je savrseno ceo broj, pa je
        ' bez ove provere pisac primao negativne gajbe koje strog citalac posle
        ' odbija (RequireZbrStavkaUgovor). Pisac i citalac moraju da traze isto.
        If CDbl(s("KolAmbalaze")) < 0 Then
            Err.Raise vbObjectError + 1360, src, _
                      "Ocekivana ambalaza za klasu " & klasa & " je negativna."
        End If
        RequireCeoBroj CDbl(s("KolAmbalaze")), _
                       "Ocekivana ambalaza za klasu " & klasa, src

        kolPoKlasi.Add klasa, CDbl(s("Kolicina"))
        ambPoKlasi.Add klasa, CDbl(s("KolAmbalaze"))
    Next i

    Dim klase As Collection
    Set klase = KlaseUKanonskomRedu(kolPoKlasi)

    Dim k As Long, rb As Long, stavkaID As String
    For k = 1 To klase.count
        klasa = CStr(klase(k))
        rb = rb + 1

        ' Red bez identiteta je gori od pada: prolazi kroz commit, a nijedna
        ' kasnija radnja ne moze da ga pogodi. Isti standard vec drzi CreateZbirna.
        stavkaID = NewEntityID("ZBS-")
        If Len(stavkaID) = 0 Then
            Err.Raise vbObjectError + 1361, src, _
                      "NewEntityID nije vratio ZbirnaStavkaID za klasu " & klasa & "."
        End If

        If AppendRow(TBL_ZBIRNA_STAVKE, _
                     BuildZbirnaStavkaRowData(stavkaID, zbirnaID, rb, _
                                              klasa, CDbl(kolPoKlasi(klasa)), _
                                              CDbl(ambPoKlasi(klasa)))) <= 0 Then
            Err.Raise vbObjectError + 1346, src, _
                      "AppendRow za " & TBL_ZBIRNA_STAVKE & " nije uspeo."
        End If
    Next k
End Sub

' --- core: izmena nacrta ----------------------------------------------------
'
' Nacrt je radna povrsina dok nije izdat; posle izdavanja je tvrdnja o robi koja
' je otisla i ispravlja se novom verzijom (ZBR-KANON-03). Ista granica koju
' otpremnica drzi od S3b-1, i isti obrazac -- OtpIzmeniDraft.
'
' STA OVDE NE MOZE DA SE PROMENI: VrstaVoca, SortaVoca i TipAmbalaze. One nisu
' polja zaglavlja nego cinjenica ROBE koju donosi prvi izvor (review #372), pa
' ih HdrProveriKljuceve i ne prima. Menjaju se jedino sastavom.
Private Sub ZbrIzmeniDraft(ByVal zbirnaID As String, ByVal h As Object, _
                           ByVal ocekivano As Collection)
    Const SRC As String = "ZbrIzmeniDraft"

    If h Is Nothing Then
        Err.Raise vbObjectError + 1364, SRC, "Header nije prosledjen."
    End If

    Dim rZbr As Long
    rZbr = ZbrRedHeadera(zbirnaID, SRC)
    RequireZbrDraft zbirnaID, rZbr, SRC

    HdrProveriKljuceve h, SRC

    Dim datum As Date
    Dim vozacID As String, brojZbirne As String, kupacID As String

    ZbrHdrCitajIProveri h, SRC, datum, vozacID, brojZbirne, kupacID

    ' Nacrt sme da promeni vozaca, datum i broj u ISTOM potezu, pa broj koji je
    ' bio tacan postane tudj bez ijedne druge provere.
    modBrojevi.RequireBrojUKontekstu modBrojevi.KIND_ZBR, vozacID, datum, _
                                     brojZbirne, SRC

    ' Nacrt sme da ZADRZI svoj broj i dan, a ne sme da preuzme tudji: sopstveni
    ' red se izuzima po ZbirnaID-u, ne po datumu, jer zaglavlje sme da ostane na
    ' istom danu.
    modBrojevi.RequireBrojSlobodanUNizu modBrojevi.KIND_ZBR, vozacID, datum, _
                                        brojZbirne, SRC, zbirnaID

    RequireUpdateCell TBL_ZBIRNA, rZbr, COL_ZBR_DATUM, datum, SRC
    RequireUpdateCell TBL_ZBIRNA, rZbr, COL_ZBR_VOZAC, vozacID, SRC
    RequireUpdateCell TBL_ZBIRNA, rZbr, COL_ZBR_BROJ, brojZbirne, SRC
    RequireUpdateCell TBL_ZBIRNA, rZbr, COL_ZBR_KUPAC, kupacID, SRC
    RequireUpdateCell TBL_ZBIRNA, rZbr, COL_ZBR_HLADNJACA, _
                      HdrOpcion(h, "Hladnjaca"), SRC
    RequireUpdateCell TBL_ZBIRNA, rZbr, COL_ZBR_POGON, HdrOpcion(h, "Pogon"), SRC

    ' Ocekivanje se pise IZNOVA, ne krpi: klasa koja je nestala iz najave mora da
    ' nestane i iz stavki, inace nacrt meri prema klasi koju vise ne tvrdi.
    ZbrObrisiOcekivano zbirnaID, SRC
    ZbrUpisiOcekivano zbirnaID, ocekivano, SRC

    ' IZMENA ZAGLAVLJA SME DA POKVARI VEC VALJANO CLANSTVO: nacrt vozaca V1 sa
    ' clanom vozaca V1, prebacen na V2, nosi clana koga Dodaj nikad ne bi primio.
    ' Izdavanje bi to na kraju uhvatilo, ali invarijanta ne sme da bude prekrsena
    ' IZMEDJU dva klika -- ekran u medjuvremenu uredno prikazuje nevalidan sastav.
    '
    ' Provera ide POSLE upisa, nad NOVIM vrednostima; pad ovde rollback-uje ceo
    ' update, pa staro zaglavlje i staro ocekivanje ostaju netaknuti.
    Dim clanovi As Collection
    Set clanovi = ZbrClanovi(zbirnaID)

    Dim k As Long
    For k = 1 To clanovi.count
        ZbrRequireIzvorValjan zbirnaID, CStr(clanovi(k)), SRC, False
    Next k
End Sub

Private Sub ZbrObrisiOcekivano(ByVal zbirnaID As String, ByVal src As String)
    Dim d As Variant
    d = GetTableData(TBL_ZBIRNA_STAVKE)
    If Not IsArray(d) Then Exit Sub

    Dim cZbr As Long
    cZbr = RequireColumnIndex(TBL_ZBIRNA_STAVKE, COL_ZBS_ZBIRNA_ID, src)

    ' Odozdo nagore: brisanje reda pomera indekse iznad njega.
    Dim i As Long
    For i = UBound(d, 1) To 1 Step -1
        If StrComp(Trim$(NzToText(d(i, cZbr))), Trim$(zbirnaID), vbTextCompare) = 0 Then
            RequireDeleteRow TBL_ZBIRNA_STAVKE, i, src
        End If
    Next i
End Sub

' STA JE VALJAN IZVOR ZBIRNE -- jedna definicija za oba kanonska ulaza
' (review #372, P1).
'
' Do ovog reza je pravilo "izvor mora biti IZDATA otpremnica" stajalo samo na
' putu nacrta (ZbrRequireIzvorValjan), dok ga je jednopotezni CreateZbirna
' preskakao -- pa je ista DRAFT otpremnica bila odbijena na jednom ulazu i
' primljena na drugom. Dva kanonska command-a ne smeju da imaju dve definicije
' istog pojma.
'
' Ovde su SAMO tvrdnje o samom izvoru. Tvrdnje o odnosu prema konkretnoj zbirnoj
' (slobodan, isti vozac, ista vrsta/sorta/ambalaza) ostaju kod pozivaoca, jer
' zavise od cilja.
Private Sub RequireOtpValidanIzvorZbirne(ByRef otp As Variant, ByVal rOtp As Long, _
                                         ByVal otpremnicaID As String, _
                                         ByVal src As String)
    If StrComp(Trim$(NzToText(otp(rOtp, _
               RequireColumnIndex(TBL_OTPREMNICA, COL_STORNIRANO, src)))), _
               "Da", vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 1352, src, _
                  "Otpremnica je stornirana: " & otpremnicaID
    End If

    If Not OtpremnicaJeIzdata(otpremnicaID) Then
        Err.Raise vbObjectError + 1353, src, _
                  "Otpremnica nije izdata: " & otpremnicaID & _
                  ". Zbirna nosi robu koja je otisla, a nacrt nije otisao."
    End If
End Sub

' --- core: nacrt zbirne, clanstvo i izdavanje -------------------------------

Private Function ZbrRedHeadera(ByVal zbirnaID As String, ByVal src As String) As Long
    Dim redovi As Collection
    Set redovi = FindRows(TBL_ZBIRNA, COL_ZBR_ID, Trim$(zbirnaID))

    Dim n As Long
    If Not redovi Is Nothing Then n = redovi.count

    If n <> 1 Then
        Err.Raise vbObjectError + 1347, src, _
                  "Zaglavlje zbirne se ne nalazi tacno jednom: " & zbirnaID & _
                  " (nadjeno " & CStr(n) & ")."
    End If

    ZbrRedHeadera = CLng(redovi(1))
End Function

' Nacrt se menja, izdata zbirna ne. Ista granica kao kod otpremnice: posle
' izdavanja dokument je tvrdnja o robi koja je otisla, a ne radna povrsina.
Private Sub RequireZbrDraft(ByVal zbirnaID As String, ByVal rZbr As Long, _
                            ByVal src As String)
    Dim data As Variant
    data = GetTableData(TBL_ZBIRNA)

    If StrComp(Trim$(NzToText(data(rZbr, _
               RequireColumnIndex(TBL_ZBIRNA, COL_STORNIRANO, src)))), _
               "Da", vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 1348, src, _
                  "Zbirna je stornirana: " & zbirnaID
    End If

    Dim status As String
    status = UCase$(Trim$(NzToText(data(rZbr, _
             RequireColumnIndex(TBL_ZBIRNA, COL_TRACE_IZDATO_STATUS, src)))))

    If status <> UCase$(IZDATO_DRAFT) Then
        Err.Raise vbObjectError + 1349, src, _
                  "Zbirna nije nacrt (status " & status & "): " & zbirnaID & _
                  ". Izdata zbirna se ne menja -- ispravlja se novom verzijom."
    End If
End Sub

' IZVOR ZBIRNE JE IZDATA OTPREMNICA.
'
' Odluka koju je S14.14 ostavila S4: nacrt otpremnice je najava, ne roba koja je
' otisla, pa ne moze da bude deo prevoznog spiska. PROSLEDJENO se racuna kao
' izdato (review #362).
'
' Ostale kapije drze da zbirna bude JEDAN transport JEDNOG vozaca: isti vozac, i
' ista vrsta/sorta/tip ambalaze koje je nacrt najavio. Zaglavlje koje neku od tih
' tvrdnji nije dalo se ne poredi -- tada je izvor taj koji je definise.
Private Sub ZbrRequireIzvorValjan(ByVal zbirnaID As String, _
                                  ByVal otpremnicaID As String, _
                                  ByVal src As String, _
                                  ByVal traziSlobodan As Boolean)
    If Len(otpremnicaID) = 0 Then
        Err.Raise vbObjectError + 1350, src, "Prazan OtpremnicaID."
    End If

    Dim rZbr As Long
    rZbr = ZbrRedHeadera(zbirnaID, src)
    RequireZbrDraft zbirnaID, rZbr, src

    Dim zbr As Variant
    zbr = GetTableData(TBL_ZBIRNA)

    Dim otp As Variant
    otp = GetTableData(TBL_OTPREMNICA)
    If Not IsArray(otp) Then
        Err.Raise vbObjectError + 1351, src, "Tabela otpremnica je prazna."
    End If

    Dim cOtpID As Long
    cOtpID = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_ID, src)

    Dim rOtp As Long
    rOtp = NadjiJedanRedOtpremnice(otp, cOtpID, otpremnicaID, src)

    RequireOtpValidanIzvorZbirne otp, rOtp, otpremnicaID, src

    RequireIstoPolje Trim$(NzToText(zbr(rZbr, _
                         RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_VOZAC, src)))), _
                     Trim$(NzToText(otp(rOtp, _
                         RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VOZAC, src)))), _
                     "VozacID", otpremnicaID, src

    ZbrRequireIstiAko zbr, rZbr, COL_ZBR_VRSTA, otp, rOtp, COL_OTP_VRSTA, _
                      "VrstaVoca", otpremnicaID, src
    ZbrRequireIstiAko zbr, rZbr, COL_ZBR_SORTA, otp, rOtp, COL_OTP_SORTA, _
                      "SortaVoca", otpremnicaID, src
    ZbrRequireIstiAko zbr, rZbr, COL_ZBR_TIP_AMB, otp, rOtp, COL_OTP_TIP_AMB, _
                      "TipAmbalaze", otpremnicaID, src

    If traziSlobodan Then
        Dim clanstvo As Object
        Set clanstvo = AktivnoClanstvoPoKanonu(src)
        If clanstvo.Exists(UCase$(otpremnicaID)) Then
            Err.Raise vbObjectError + 1354, src, _
                      "Otpremnica je vec u sastavu aktivne zbirne: " & otpremnicaID & _
                      " -> " & CStr(clanstvo(UCase$(otpremnicaID)))
        End If
    End If
End Sub

' Poredi samo kad je zaglavlje tu tvrdnju DALO. Prazno polje na nacrtu znaci
' "nisam rekao", ne "mora biti prazno kod izvora".
Private Sub ZbrRequireIstiAko(ByRef zbr As Variant, ByVal rZbr As Long, _
                              ByVal kolZbr As String, _
                              ByRef otp As Variant, ByVal rOtp As Long, _
                              ByVal kolOtp As String, _
                              ByVal opis As String, ByVal otpremnicaID As String, _
                              ByVal src As String)
    Dim ocekivano As String
    ocekivano = Trim$(NzToText(zbr(rZbr, RequireColumnIndex(TBL_ZBIRNA, kolZbr, src))))
    If Len(ocekivano) = 0 Then Exit Sub

    RequireIstoPolje ocekivano, _
                     Trim$(NzToText(otp(rOtp, RequireColumnIndex(TBL_OTPREMNICA, kolOtp, src)))), _
                     opis, otpremnicaID, src
End Sub

' Preuzimanje cinjenica nad (zbirnaID, otpremnicaID) -- redove nalazi sam, da
' pozivaoci ne bi nosili indekse kroz slojeve.
Private Sub ZbrPreuzmiCinjeniceZaIzvor(ByVal zbirnaID As String, _
                                       ByVal otpremnicaID As String, _
                                       ByVal src As String)
    Dim otp As Variant
    otp = GetTableData(TBL_OTPREMNICA)
    If Not IsArray(otp) Then Exit Sub

    Dim rOtp As Long
    rOtp = NadjiJedanRedOtpremnice(otp, _
               RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_ID, src), otpremnicaID, src)

    ZbrPreuzmiCinjenice zbirnaID, ZbrRedHeadera(zbirnaID, src), otp, rOtp, src
End Sub

' Prvi izvor DEFINISE cinjenice robe koje nacrt nije imao odakle da zna.
'
' Poziva se POSLE provere valjanosti, pa se upisuje samo ono sto je vec proslo
' kapiju poklapanja: polje koje zaglavlje ima mora da se slaze (ZbrRequireIstiAko),
' a polje koje nema se preuzima. Sledeci izvor tada vise nema sta da definise --
' meri se prema upisanom.
Private Sub ZbrPreuzmiCinjenice(ByVal zbirnaID As String, ByVal rZbr As Long, _
                                ByRef otp As Variant, ByVal rOtp As Long, _
                                ByVal src As String)
    ZbrPreuzmiPolje zbirnaID, rZbr, COL_ZBR_VRSTA, otp, rOtp, COL_OTP_VRSTA, src
    ZbrPreuzmiPolje zbirnaID, rZbr, COL_ZBR_SORTA, otp, rOtp, COL_OTP_SORTA, src
    ZbrPreuzmiPolje zbirnaID, rZbr, COL_ZBR_TIP_AMB, otp, rOtp, COL_OTP_TIP_AMB, src
End Sub

Private Sub ZbrPreuzmiPolje(ByVal zbirnaID As String, ByVal rZbr As Long, _
                            ByVal kolZbr As String, _
                            ByRef otp As Variant, ByVal rOtp As Long, _
                            ByVal kolOtp As String, ByVal src As String)
    Dim zbr As Variant
    zbr = GetTableData(TBL_ZBIRNA)

    Dim sada As String
    sada = Trim$(NzToText(zbr(rZbr, RequireColumnIndex(TBL_ZBIRNA, kolZbr, src))))
    If Len(sada) > 0 Then Exit Sub

    Dim izIzvora As String
    izIzvora = Trim$(NzToText(otp(rOtp, RequireColumnIndex(TBL_OTPREMNICA, kolOtp, src))))
    If Len(izIzvora) = 0 Then Exit Sub

    RequireUpdateCell TBL_ZBIRNA, rZbr, kolZbr, izIzvora, src
End Sub

Private Sub ZbrUpisiClanstvo(ByVal zbirnaID As String, _
                             ByVal otpremnicaID As String, ByVal src As String)
    Dim izvorID As String
    izvorID = NewEntityID("ZBI-")
    If Len(izvorID) = 0 Then
        Err.Raise vbObjectError + 1362, src, _
                  "NewEntityID nije vratio ZbirnaIzvorID."
    End If

    If AppendRow(TBL_ZBIRNA_IZVORI, _
                 BuildZbirnaIzvorRowData(izvorID, zbirnaID, otpremnicaID)) <= 0 Then
        Err.Raise vbObjectError + 1355, src, _
                  "AppendRow za " & TBL_ZBIRNA_IZVORI & " nije uspeo."
    End If
End Sub

Private Sub ZbrUkloniIzvor(ByVal zbirnaID As String, _
                           ByVal otpremnicaID As String, ByVal src As String)
    Dim izv As Variant
    izv = GetTableData(TBL_ZBIRNA_IZVORI)
    If Not IsArray(izv) Then
        Err.Raise vbObjectError + 1356, src, _
                  "Otpremnica nije u sastavu ove zbirne: " & otpremnicaID
    End If

    Dim cZbr As Long, cOtp As Long
    cZbr = RequireColumnIndex(TBL_ZBIRNA_IZVORI, COL_ZBI_ZBIRNA_ID, src)
    cOtp = RequireColumnIndex(TBL_ZBIRNA_IZVORI, COL_ZBI_OTPREMNICA_ID, src)

    Dim i As Long, nadjen As Long, koliko As Long
    For i = 1 To UBound(izv, 1)
        If StrComp(Trim$(NzToText(izv(i, cZbr))), Trim$(zbirnaID), vbTextCompare) = 0 Then
            If StrComp(Trim$(NzToText(izv(i, cOtp))), otpremnicaID, vbTextCompare) = 0 Then
                nadjen = i
                koliko = koliko + 1
            End If
        End If
    Next i

    If koliko = 0 Then
        Err.Raise vbObjectError + 1356, src, _
                  "Otpremnica nije u sastavu ove zbirne: " & otpremnicaID
    End If
    If koliko > 1 Then
        Err.Raise vbObjectError + 1357, src, _
                  "Isti par (zbirna, otpremnica) postoji " & CStr(koliko) & _
                  " puta: " & zbirnaID & " / " & otpremnicaID
    End If

    RequireDeleteRow TBL_ZBIRNA_IZVORI, nadjen, src
End Sub

' ZBR-KANON-04: IZVEDENA CINJENICA ZIVI TACNO KOLIKO I NJEN IZVOR.
'
' Vrstu, sortu i tip ambalaze nacrta ne bira operater nego ih donosi PRVI izvor
' (ZbrPreuzmiCinjenice). Kad se ukloni i poslednji izvor, iza njih ne stoji
' nijedna otpremnica: ostavljene, one su tvrdnja o robi bez robe -- i tiho suzuju
' prazan nacrt na vrstu koju operater nikad nije izabrao niti je na ekranu vidi.
'
' Dok ima BAR JEDNOG clana se NE diraju: tada ih izvor i dalje pokriva, a svaki
' sledeci izvor se meri prema njima (ZbrRequireIstiAko).
Private Sub ZbrOcistiCinjeniceBezClanstva(ByVal zbirnaID As String, _
                                          ByVal src As String)
    If ZbrClanovi(zbirnaID).count > 0 Then Exit Sub

    Dim rZbr As Long
    rZbr = ZbrRedHeadera(zbirnaID, src)

    RequireUpdateCell TBL_ZBIRNA, rZbr, COL_ZBR_VRSTA, "", src
    RequireUpdateCell TBL_ZBIRNA, rZbr, COL_ZBR_SORTA, "", src
    RequireUpdateCell TBL_ZBIRNA, rZbr, COL_ZBR_TIP_AMB, "", src
End Sub

' RAZLOG ZBOG KOG IZDAVANJE NE BI PROSLO nad trenutnim clanstvom.
' "" = proslo bi. Ne dize gresku -- pozivaoci se razlikuju po tome sta rade sa
' odgovorom: ZbrIzdaj ga dize, traka ga prikazuje.
'
' Meri SAMO valjanost izvora (storniran, vise nije izdat, tudji vozac, druga
' vrsta/sorta/tip ambalaze). Jednakost najavljenog i povezanog je zaseban sud i
' ostaje u ZbrIzdaj -- traka je vec pokazuje kroz ostatak i semafor.
Public Function ZbrIzvoriNevaljaniRazlog(ByVal zbirnaID As String) As String
    Const SRC As String = "ZbrIzvoriNevaljaniRazlog"

    Dim clanovi As Collection, k As Long
    On Error GoTo EH

    Set clanovi = ZbrClanovi(zbirnaID)
    For k = 1 To clanovi.count
        ZbrRequireIzvorValjan zbirnaID, CStr(clanovi(k)), SRC, False
    Next k
    Exit Function
EH:
    ZbrIzvoriNevaljaniRazlog = Err.description
End Function

' --- core: izdavanje zbirne -------------------------------------------------
Private Sub ZbrIzdaj(ByVal zbirnaID As String)
    Const SRC As String = "ZbrIzdaj"

    Dim rZbr As Long
    rZbr = ZbrRedHeadera(zbirnaID, SRC)
    RequireZbrDraft zbirnaID, rZbr, SRC

    Dim clanovi As Collection
    Set clanovi = ZbrClanovi(zbirnaID)

    If clanovi.count = 0 Then
        Err.Raise vbObjectError + 1358, SRC, _
                  "Zbirna nema nijedan izvor: " & zbirnaID & _
                  ". Zbirna bez otpremnica nije prevoz."
    End If

    ' REVALIDACIJA, isti razlog kao kod otpremnice: izmedju dodavanja i izdavanja
    ' prolazi vreme, pa izvor moze da bude storniran ili ispravljen. Provera i
    ' upotreba moraju biti u istom trenutku.
    '
    ' JEDNA IMPLEMENTACIJA, DVA POZIVAOCA (review #381, P2): isti sud koristi i
    ' traka napretka, koja ga pokazuje kao STANJE umesto da ga digne kao gresku.
    ' Da svaki racuna svoje, ekran bi mogao da kaze "spremna" za zbirnu koju
    ' izdavanje odbija.
    Dim razlogIzvora As String
    razlogIzvora = ZbrIzvoriNevaljaniRazlog(zbirnaID)
    If Len(razlogIzvora) > 0 Then
        Err.Raise vbObjectError + 1366, SRC, razlogIzvora
    End If

    Dim ocek As Object, ocekAmb As Object
    Dim pov As Object, povAmb As Object
    ZbrUcitajOcekivano zbirnaID, ocek, ocekAmb, SRC
    ZbrUcitajPovezano zbirnaID, clanovi, pov, povAmb, SRC

    ZbrRequireJednakost zbirnaID, ocek, pov, "Kolicina", SRC
    ZbrRequireJednakost zbirnaID, ocekAmb, povAmb, "KolAmbalaze", SRC

    RequireUpdateCell TBL_ZBIRNA, rZbr, COL_TRACE_IZDATO_STATUS, IZDATO_IZDATO, SRC
End Sub

Private Sub ZbrUcitajOcekivano(ByVal zbirnaID As String, ByRef ocek As Object, _
                               ByRef ocekAmb As Object, ByVal src As String)
    Set ocek = CreateObject("Scripting.Dictionary")
    Set ocekAmb = CreateObject("Scripting.Dictionary")

    Dim stavke As Collection
    Set stavke = StavkeZaZbirnu(StavkeZbirnePoDokumentu(), zbirnaID, src)

    Dim i As Long, red As Variant, klasa As String
    For i = 1 To stavke.count
        red = stavke(i)
        klasa = UCase$(Trim$(CStr(red(3))))
        ocek(klasa) = CDbl(red(4))
        ocekAmb(klasa) = CDbl(red(5))
    Next i
End Sub

' Povezano = zbir stavki IZVORNIH otpremnica, po klasi.
Private Sub ZbrUcitajPovezano(ByVal zbirnaID As String, ByVal clanovi As Collection, _
                              ByRef pov As Object, ByRef povAmb As Object, _
                              ByVal src As String)
    Set pov = CreateObject("Scripting.Dictionary")
    Set povAmb = CreateObject("Scripting.Dictionary")

    Dim poDok As Object
    Set poDok = StavkeOtpremnicePoDokumentu()

    Dim k As Long, i As Long, stavke As Collection, red As Variant, klasa As String
    For k = 1 To clanovi.count
        Set stavke = StavkeZaOtpremnicu(poDok, CStr(clanovi(k)), src)
        For i = 1 To stavke.count
            red = stavke(i)
            klasa = UCase$(Trim$(CStr(red(3))))
            If Not pov.Exists(klasa) Then
                pov.Add klasa, 0#
                povAmb.Add klasa, 0#
            End If
            pov(klasa) = CDbl(pov(klasa)) + CDbl(red(4))
            povAmb(klasa) = CDbl(povAmb(klasa)) + CDbl(red(6))
        Next i
    Next k
End Sub

' Jednakost PO KLASI, u oba smera: klasa koju je nacrt najavio a izvori je nemaju
' je isto neslaganje kao klasa koju izvori nose a nacrt je nije najavio.
Private Sub ZbrRequireJednakost(ByVal zbirnaID As String, ByVal ocek As Object, _
                                ByVal pov As Object, ByVal opis As String, _
                                ByVal src As String)
    Dim sve As Object
    Set sve = CreateObject("Scripting.Dictionary")

    Dim k As Variant
    For Each k In ocek.keys
        If Not sve.Exists(CStr(k)) Then sve.Add CStr(k), True
    Next k
    For Each k In pov.keys
        If Not sve.Exists(CStr(k)) Then sve.Add CStr(k), True
    Next k

    Dim a As Double, b As Double
    For Each k In sve.keys
        a = 0#: b = 0#
        If ocek.Exists(CStr(k)) Then a = CDbl(ocek(CStr(k)))
        If pov.Exists(CStr(k)) Then b = CDbl(pov(CStr(k)))

        If Abs(a - b) > 0.0001 Then
            Err.Raise vbObjectError + 1359, src, _
                      "Zbirna " & zbirnaID & ": " & opis & " za klasu " & CStr(k) & _
                      " -- najavljeno " & Fmt2Zbr(a) & ", povezano " & Fmt2Zbr(b) & _
                      ". Zbirna se izdaje tek kad je najava pokrivena."
        End If
    Next k
End Sub

Private Function ZbrPadTransakcije(ByRef tx As clsTransaction, _
                                   ByVal ulaz As String, _
                                   ByVal entitetID As String) As String
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogError ulaz, errDesc, errNum
    Monitor_Error _
        moduleName:="modDokumenta", _
        procedureName:=ulaz, _
        entityType:="Zbirna", _
        entityID:=entitetID, _
        correlationId:=entitetID, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    PrintTxFailure ulaz, errSrc, errNum, errDesc
    ZbrPadTransakcije = errDesc
End Function

' --- kanonski citaoci stavki zbirne (S4-1) ----------------------------------
'
' Zaglavlje zbirne od PR3 NE nosi UkupnoKolicina, UkupnoAmbalaze ni Klasu --
' CreateZbirna_TX ih namerno ostavlja prazne (v. BuildZbirnaHeaderRowData), a
' kolicinu po klasi pise u tblZbirnaStavke. Svaki citalac koji je te kolone i
' dalje citao za kanonsku zbirnu dobija PRAZNO: nula kilograma, prazna klasa,
' prosek gajbe 0. Ovde je jedno mesto sa kojeg svi citaju, po istom obrascu koji
' otkup ima od S14.7 (StavkeOtkupaRedovi) i otpremnica od S3b-1.
'
' UGOVOR JE DOKUMENTSKI, NE STAVKARSKI (isti razlog kao kod otkupa i
' otpremnice, review #334 P1):
'   1) svaka stavka ima neprazan ZbirnaID i zaglavlje TACNO JEDNOM;
'   2) svako zaglavlje sa ZbirnaID-em ima BAR JEDNU stavku -- pisac to vec trazi
'      (CreateZbirna odbija klasu sa zbirom <= 0), pa citalac koji bi nulu vratio
'      kao podatak laze o dokumentu koji pisac ne moze da napravi;
'   3) Kolicina > 0, Klasa kroz istu kapiju koju pisac zove, KolAmbalaze ceo
'      broj >= 0 (prazno = 0 gajbi);
'   4) jedna stavka po KLASI po dokumentu -- pisac deli po klasi
'      (KlaseUKanonskomRedu), pa dve stavke klase I nisu dokument nego kvar.
'
' ZBIRNA NEMA CENU. tblZbirna nema kolonu Cena i nijedan pisac je ne prima
' (v. modDokUnos, F3). Mesto vrednosti u zbiru zato ostaje Null, ne nula --
' slucajna upotreba treba da padne, a ne da tiho sabere nulu.
Public Function StavkeZbirneRedovi() As Variant
    Const SRC As String = "StavkeZbirneRedovi"

    StavkeZbirneRedovi = Empty

    Dim zagl As Object
    Set zagl = ZaglavljaZbirnePoID(SRC)
    RequireJedinstvenoZaglavljeZbirne zagl, SRC

    Dim imaStavku As Object
    Set imaStavku = CreateObject("Scripting.Dictionary")

    Dim d As Variant
    d = GetTableData(TBL_ZBIRNA_STAVKE)

    Dim cZbr As Long, cRb As Long, cKl As Long, cKol As Long
    Dim cAmb As Long, cId As Long
    Dim i As Long, n As Long, zid As String

    If IsArray(d) Then
        cZbr = RequireColumnIndex(TBL_ZBIRNA_STAVKE, COL_ZBS_ZBIRNA_ID, SRC)
        cRb = RequireColumnIndex(TBL_ZBIRNA_STAVKE, COL_ZBS_RB, SRC)
        cKl = RequireColumnIndex(TBL_ZBIRNA_STAVKE, COL_ZBS_KLASA, SRC)
        cKol = RequireColumnIndex(TBL_ZBIRNA_STAVKE, COL_ZBS_KOLICINA, SRC)
        cAmb = RequireColumnIndex(TBL_ZBIRNA_STAVKE, COL_ZBS_KOL_AMB, SRC)
        cId = RequireColumnIndex(TBL_ZBIRNA_STAVKE, COL_ZBS_ID, SRC)

        Dim parKlasa As Object, kParKl As String
        Set parKlasa = CreateObject("Scripting.Dictionary")

        For i = 1 To UBound(d, 1)
            zid = Trim$(NzToText(d(i, cZbr)))
            RequireZbrStavkaUgovor zagl, zid, d(i, cKl), d(i, cKol), d(i, cAmb), _
                                   i, SRC
            kParKl = UCase$(zid) & "|" & UCase$(Trim$(NzToText(d(i, cKl))))
            If parKlasa.Exists(kParKl) Then
                Err.Raise vbObjectError + 1952, SRC, _
                          "Dve stavke iste klase na zbirnoj: ZbirnaID=" & zid & _
                          ", klasa " & Trim$(NzToText(d(i, cKl))) & "."
            End If
            parKlasa.Add kParKl, True
            imaStavku(zid) = True
            n = n + 1
        Next i
    End If

    RequireZaglavljaZbirneSaStavkama zagl, imaStavku, SRC

    If n = 0 Then Exit Function

    ' Kolone: 1 ZbirnaID, 2 RedniBroj, 3 Klasa, 4 Kolicina, 5 KolAmbalaze,
    ' 6 ZbirnaStavkaID. Nove kolone idu NA KRAJ: citaoci indeksiraju poziciono.
    Dim res() As Variant
    ReDim res(1 To n, 1 To 6)
    n = 0
    For i = 1 To UBound(d, 1)
        n = n + 1
        res(n, 1) = Trim$(NzToText(d(i, cZbr)))
        res(n, 2) = d(i, cRb)
        res(n, 3) = Trim$(NzToText(d(i, cKl)))
        res(n, 4) = CDbl(d(i, cKol))
        res(n, 5) = OtpDbl(d(i, cAmb))
        res(n, 6) = Trim$(NzToText(d(i, cId)))
    Next i

    StavkeZbirneRedovi = res
End Function

Private Function ZaglavljaZbirnePoID(ByVal sourceName As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Set ZaglavljaZbirnePoID = dict

    Dim d As Variant
    d = GetTableData(TBL_ZBIRNA)
    If Not IsArray(d) Then Exit Function

    Dim cId As Long
    cId = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_ID, sourceName)

    Dim i As Long, zid As String
    For i = 1 To UBound(d, 1)
        zid = Trim$(NzToText(d(i, cId)))
        If Len(zid) > 0 Then
            If dict.Exists(zid) Then
                dict(zid) = CLng(dict(zid)) + 1
            Else
                dict.Add zid, 1&
            End If
        End If
    Next i
End Function

' JEDAN LOGICKI DOKUMENT = JEDAN ID.
'
' Stari model je zbirnu sa dve klase drzao kao DVA REDA tblZbirna pod istim
' BrojZbirne -- zato je i trebao GeneracijaID da razlikuje dokumente. U kanonu
' je to jedno zaglavlje sa dve stavke, pa dva reda sa istim ZbirnaID nisu
' dvoklasna zbirna nego kvar: dokument-level citaoci iteriraju zaglavlja, pa bi
' isti teret bio izbrojan dvaput.
Private Sub RequireJedinstvenoZaglavljeZbirne(ByVal zagl As Object, _
                                              ByVal sourceName As String)
    Dim k As Variant
    For Each k In zagl.keys
        If CLng(zagl(k)) <> 1 Then
            Err.Raise vbObjectError + 1951, sourceName, _
                      "Zaglavlje zbirne se ne nalazi tacno jednom: " & CStr(k) & _
                      " (headera: " & CStr(zagl(k)) & "). ZbirnaID je identitet " & _
                      "dokumenta -- dva reda sa istim ID-em nisu dokument."
        End If
    Next k
End Sub

Private Sub RequireZbrStavkaUgovor(ByVal zagl As Object, ByVal zid As String, _
                                   ByVal klasa As Variant, ByVal kol As Variant, _
                                   ByVal amb As Variant, ByVal red As Long, _
                                   ByVal sourceName As String)
    If Len(zid) = 0 Then
        Err.Raise vbObjectError + 1953, sourceName, _
                  "Stavka zbirne bez ZbirnaID-a: " & TBL_ZBIRNA_STAVKE & _
                  ", red " & CStr(red) & ". Stavka bez dokumenta se ne moze citati."
    End If

    If Not zagl.Exists(zid) Then
        Err.Raise vbObjectError + 1954, sourceName, _
                  "Zaglavlje zbirne ne postoji: " & zid & ", stavka u redu " & _
                  CStr(red) & ". Stavka bez dokumenta se ne moze citati."
    End If

    If Not IsNumeric(kol) Then
        Err.Raise vbObjectError + 1955, sourceName, _
                  "Kolicina stavke nije brojcana: ZbirnaID=" & zid & "."
    End If

    If CDbl(kol) <= 0 Then
        Err.Raise vbObjectError + 1956, sourceName, _
                  "Kolicina stavke mora biti veca od nule: ZbirnaID=" & zid & "."
    End If

    ' Gajbe su KOMADI: prazno je 0, a nebrojcano, negativno i decimalno je kvar --
    ' isto sto pisac odbija (CreateZbirna -> RequireCeoBroj).
    If Len(Trim$(NzToText(amb))) > 0 Then
        If Not IsNumeric(amb) Then
            Err.Raise vbObjectError + 1955, sourceName, _
                      "KolAmbalaze stavke nije brojcana: ZbirnaID=" & zid & "."
        End If
        If CDbl(amb) < 0 Then
            Err.Raise vbObjectError + 1956, sourceName, _
                      "KolAmbalaze stavke ne sme biti negativna: ZbirnaID=" & zid & "."
        End If
        RequireCeoBroj CDbl(amb), "KolAmbalaze stavke (ZbirnaID=" & zid & ")", _
                       sourceName
    End If

    RequireValidDocumentClass Trim$(NzToText(klasa)), _
                              sourceName & " (ZbirnaID=" & zid & ")"
End Sub

Private Sub RequireZaglavljaZbirneSaStavkama(ByVal zagl As Object, _
                                             ByVal imaStavku As Object, _
                                             ByVal sourceName As String)
    Dim k As Variant
    For Each k In zagl.keys
        If Not imaStavku.Exists(CStr(k)) Then
            Err.Raise vbObjectError + 1957, sourceName, _
                      "Zbirna nema nijednu stavku: " & CStr(k) & _
                      ". Kolicina dokumenta se racuna iz " & TBL_ZBIRNA_STAVKE & "."
        End If
    Next k
End Sub

' Zbir stavki po dokumentu. Kljuc je ZbirnaID, vrednost Array(kg, Null, gajbe,
' klase) -- isti oblik koji vec citaju mreza F8 i izvestaji za otkup i
' otpremnicu, da treci tip ne donese treci raspored pozicija.
'
' SVAKI kanonski dokument JE u recniku -- zaglavlje bez stavki obara citaoca
' gore, pa "nema kljuca" ne moze da znaci 0 kg.
Public Function ZbirStavkiPoZbirni() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Set ZbirStavkiPoZbirni = dict

    Dim s As Variant
    s = StavkeZbirneRedovi()
    If Not IsArray(s) Then Exit Function

    Dim i As Long, zid As String, kl As String, rec As Variant
    For i = 1 To UBound(s, 1)
        zid = CStr(s(i, 1))
        If dict.Exists(zid) Then
            rec = dict(zid)
        Else
            rec = Array(0#, Null, 0#, "")
        End If
        rec(0) = CDbl(rec(0)) + CDbl(s(i, 4))
        rec(2) = CDbl(rec(2)) + CDbl(s(i, 5))
        kl = Trim$(CStr(s(i, 3)))
        If Len(kl) > 0 Then
            If Len(CStr(rec(3))) = 0 Then
                rec(3) = kl
            ElseIf InStr(1, ", " & CStr(rec(3)) & ", ", ", " & kl & ", ", _
                         vbTextCompare) = 0 Then
                rec(3) = CStr(rec(3)) & ", " & kl
            End If
        End If
        dict(zid) = rec
    Next i
End Function

' Stavke grupisane po dokumentu -- za citaoce koji razvijaju jedan dokument u
' vise redova (stampa, izvestaj po klasi).
Public Function StavkeZbirnePoDokumentu() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Set StavkeZbirnePoDokumentu = dict

    Dim s As Variant
    s = StavkeZbirneRedovi()
    If Not IsArray(s) Then Exit Function

    Dim i As Long, j As Long, zid As String, c As Collection, red As Variant
    For i = 1 To UBound(s, 1)
        zid = CStr(s(i, 1))
        If dict.Exists(zid) Then
            Set c = dict(zid)
        Else
            Set c = New Collection
            dict.Add zid, c
        End If
        ReDim red(1 To 6)
        For j = 1 To 6
            red(j) = s(i, j)
        Next j
        c.Add red
    Next i
End Function

' Stavke JEDNE zbirne -- nedostajuci kljuc je GRESKA, isti razlog kao gore.
Public Function StavkeZaZbirnu(ByVal poDok As Object, _
                               ByVal zbirnaID As String, _
                               ByVal sourceName As String) As Collection
    If poDok Is Nothing Then
        Err.Raise vbObjectError + 1957, sourceName, _
                  "Stavke zbirne nisu ucitane."
    End If

    Dim zid As String
    zid = Trim$(zbirnaID)

    If Not poDok.Exists(zid) Then
        Err.Raise vbObjectError + 1957, sourceName, _
                  "Zbirna nema nijednu stavku: " & zid & _
                  ". Stavke se citaju iz " & TBL_ZBIRNA_STAVKE & "."
    End If

    Set StavkeZaZbirnu = poDok(zid)
End Function

' Kolicina i gajbe JEDNE zbirne po klasi: UCase(klasa) -> Array(kg, gajbe).
' Za citaoce koji porede po klasi (invarijanta, integritet, manjak).
Public Function ZbirnaPoKlasi(ByVal zbirnaID As String) As Object
    Const SRC As String = "ZbirnaPoKlasi"

    Dim rez As Object
    Set rez = CreateObject("Scripting.Dictionary")
    Set ZbirnaPoKlasi = rez

    Dim stavke As Collection
    Set stavke = StavkeZaZbirnu(StavkeZbirnePoDokumentu(), zbirnaID, SRC)

    Dim i As Long, red As Variant, kl As String
    For i = 1 To stavke.count
        red = stavke(i)
        kl = UCase$(Trim$(CStr(red(3))))
        If rez.Exists(kl) Then
            rez(kl) = Array(CDbl(rez(kl)(0)) + CDbl(red(4)), _
                            CDbl(rez(kl)(1)) + CDbl(red(5)))
        Else
            rez.Add kl, Array(CDbl(red(4)), CDbl(red(5)))
        End If
    Next i
End Function

' Izvori JEDNE zbirne: kolekcija OtpremnicaID-jeva iz tblZbirnaIzvori.
'
' Zbirna BEZ izvora je greska, ne prazan dokument: pisac trazi bar jedan izvor
' (CreateZbirna, 1223), pa bi prazna kolekcija znacila da je clanstvo izgubljeno,
' a citalac koji to vrati kao "nema izvora" tu stetu sakriva.
Public Function IzvoriZbirne(ByVal zbirnaID As String) As Collection
    Const SRC As String = "IzvoriZbirne"

    Dim c As Collection
    Set c = New Collection
    Set IzvoriZbirne = c

    Dim zid As String
    zid = UCase$(Trim$(zbirnaID))
    If Len(zid) = 0 Then
        Err.Raise vbObjectError + 1958, SRC, "ZbirnaID je obavezan."
    End If

    Dim izv As Variant
    izv = GetTableData(TBL_ZBIRNA_IZVORI)

    If IsArray(izv) Then
        Dim cZbr As Long, cOtp As Long, i As Long, otpID As String
        Dim vidjen As Object
        Set vidjen = CreateObject("Scripting.Dictionary")

        cZbr = RequireColumnIndex(TBL_ZBIRNA_IZVORI, COL_ZBI_ZBIRNA_ID, SRC)
        cOtp = RequireColumnIndex(TBL_ZBIRNA_IZVORI, COL_ZBI_OTPREMNICA_ID, SRC)

        For i = 1 To UBound(izv, 1)
            If StrComp(Trim$(NzToText(izv(i, cZbr))), zid, vbTextCompare) = 0 Then
                otpID = Trim$(NzToText(izv(i, cOtp)))
                If Len(otpID) = 0 Then
                    Err.Raise vbObjectError + 1959, SRC, _
                              "Clanstvo bez OtpremnicaID-a: ZbirnaID=" & zbirnaID & _
                              ", red " & CStr(i) & "."
                End If
                If vidjen.Exists(UCase$(otpID)) Then
                    Err.Raise vbObjectError + 1959, SRC, _
                              "Ista otpremnica je dvaput clan zbirne " & zbirnaID & _
                              ": " & otpID & "."
                End If
                vidjen.Add UCase$(otpID), True
                c.Add otpID
            End If
        Next i
    End If

    If c.count = 0 Then
        Err.Raise vbObjectError + 1960, SRC, _
                  "Zbirna nema nijedan izvor: " & zbirnaID & _
                  ". Clanstvo se cita iz " & TBL_ZBIRNA_IZVORI & "."
    End If
End Function

' Javni ulaz u kanonsko clanstvo: UCase(OtpremnicaID) -> ZbirnaID, samo za
' AKTIVNE zbirne. Jezgro je privatno jer ga pisac zove sa svojim SRC-om.
Public Function AktivnoZbrClanstvoPoKanonu(Optional ByVal sourceName As String = _
                                           "AktivnoZbrClanstvoPoKanonu") As Object
    Set AktivnoZbrClanstvoPoKanonu = AktivnoClanstvoPoKanonu(sourceName)
End Function

' OTPREMNICE KOJE CEKAJU ZBIRNU -- ogledalo NevezaniOtkupi, sa jednom razlikom
' koja je cela poenta S4: izvor zbirne mora da bude IZDATA otpremnica.
'
' Nacrt otpremnice je NAJAVA, ne roba koja je otisla, pa se ovde ne nudi. Da se
' nudi, operater bi ga vezao i tek bi ga izdavanje zbirne odbilo -- nekoliko
' klikova kasnije, porukom o dokumentu koji je sam izabrao sa ponudjenog spiska.
' Kapija koja odbija tek na kraju je losija od spiska koji ne laze.
'
' Vrednost je ISTORIJA, ne veza: brojevi STORNIRANIH zbirnih u cijem je sastavu
' otpremnica bila. Aktivno clanstvo cita samo AktivnoClanstvoPoKanonu.
Public Function NevezaneOtpremnice() As Object
    Const SRC As String = "NevezaneOtpremnice"

    Dim res As Object
    Set res = CreateObject("Scripting.Dictionary")
    Set NevezaneOtpremnice = res

    Dim otp As Variant
    otp = GetTableData(TBL_OTPREMNICA)
    If Not IsArray(otp) Then Exit Function

    Dim aktivno As Object, bila As Object
    Set aktivno = AktivnoClanstvoPoKanonu(SRC)
    Set bila = BivseZbirneIzvora(SRC)

    Dim cId As Long, cSto As Long, i As Long
    Dim oidRaw As String, oid As String
    cId = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_ID, SRC)
    cSto = RequireColumnIndex(TBL_OTPREMNICA, COL_STORNIRANO, SRC)

    For i = 1 To UBound(otp, 1)
        oidRaw = Trim$(NzToText(otp(i, cId)))
        oid = UCase$(oidRaw)
        If Len(oid) > 0 Then
            If StrComp(Trim$(NzToText(otp(i, cSto))), "Da", vbTextCompare) <> 0 Then
                If OtpremnicaJeIzdata(oidRaw) Then
                    If Not aktivno.Exists(oid) Then
                        If bila.Exists(oid) Then
                            res(oid) = CStr(bila(oid))
                        Else
                            res(oid) = ""
                        End If
                    End If
                End If
            End If
        End If
    Next i
End Function

' UCase(OtpremnicaID) -> brojevi STORNIRANIH zbirnih u cijem je sastavu bila,
' redom clanstva, spojeni ", ". Isti obrazac kao BivseOtpremniceIzvora.
Private Function BivseZbirneIzvora(ByVal src As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set BivseZbirneIzvora = d

    Dim izv As Variant, zbr As Variant
    izv = GetTableData(TBL_ZBIRNA_IZVORI)
    If Not IsArray(izv) Then Exit Function
    zbr = GetTableData(TBL_ZBIRNA)
    If Not IsArray(zbr) Then Exit Function

    ' Broj STORNIRANE zbirne po ID-u.
    Dim brStor As Object
    Set brStor = CreateObject("Scripting.Dictionary")

    ' Imena indeksa NE smeju da lice na 'zbr': VBA je case-insensitive, pa bi
    ' 'zBr' bila druga deklaracija istog imena i modul se ne bi kompajlirao.
    Dim cZbrId As Long, cZbrBroj As Long, cZbrSto As Long, j As Long
    cZbrId = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_ID, src)
    cZbrBroj = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, src)
    cZbrSto = RequireColumnIndex(TBL_ZBIRNA, COL_STORNIRANO, src)
    For j = 1 To UBound(zbr, 1)
        If StrComp(Trim$(NzToText(zbr(j, cZbrSto))), "Da", vbTextCompare) = 0 Then
            brStor(UCase$(Trim$(NzToText(zbr(j, cZbrId))))) = _
                Trim$(NzToText(zbr(j, cZbrBroj)))
        End If
    Next j
    If brStor.count = 0 Then Exit Function

    ' Par (otpremnica, zbirna) se broji jednom -- isti broj dve razlicite
    ' stornirane zbirne ostaje dva puta, jer su to dva dokumenta.
    Dim vidjen As Object
    Set vidjen = CreateObject("Scripting.Dictionary")

    Dim cZbr As Long, cOtp As Long, i As Long, zbrU As String, otpU As String
    cZbr = RequireColumnIndex(TBL_ZBIRNA_IZVORI, COL_ZBI_ZBIRNA_ID, src)
    cOtp = RequireColumnIndex(TBL_ZBIRNA_IZVORI, COL_ZBI_OTPREMNICA_ID, src)
    For i = 1 To UBound(izv, 1)
        zbrU = UCase$(Trim$(NzToText(izv(i, cZbr))))
        otpU = UCase$(Trim$(NzToText(izv(i, cOtp))))
        If Len(otpU) > 0 And brStor.Exists(zbrU) Then
            If Not vidjen.Exists(otpU & "|" & zbrU) Then
                vidjen.Add otpU & "|" & zbrU, True
                If d.Exists(otpU) Then
                    d(otpU) = CStr(d(otpU)) & ", " & CStr(brStor(zbrU))
                Else
                    d(otpU) = CStr(brStor(zbrU))
                End If
            End If
        End If
    Next i
End Function

' Clanstvo se pise ORIGINALNIM OtpremnicaID-em iz tabele, ne kljucem recnika:
' kljuc je UCase$ normalizovan da bi duplikat bio uhvatljiv, a u tabelu mora da
' ode ono sto tamo stvarno stoji.
Private Function IDIzRedaOtpremnice(ByRef data As Variant, ByVal colID As Long, _
                                    ByVal red As Long) As String
    IDIzRedaOtpremnice = Trim$(NzToText(data(red, colID)))
End Function

Private Function BuildZbirnaIzvorRowData(ByVal izvorID As String, _
                                         ByVal zbirnaID As String, _
                                         ByVal otpremnicaID As String) As Variant
    Const SRC As String = "BuildZbirnaIzvorRowData"

    Dim colCount As Long
    colCount = TabelaBrojKolona(TBL_ZBIRNA_IZVORI)

    If colCount <= 0 Then
        Err.Raise vbObjectError + 1258, SRC, _
                  "Ne mogu da odredim broj kolona za tblZbirnaIzvori."
    End If

    Dim rowData() As Variant
    ReDim rowData(0 To colCount - 1)

    SetRowValueByColumn rowData, TBL_ZBIRNA_IZVORI, COL_ZBI_ID, izvorID, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA_IZVORI, COL_ZBI_ZBIRNA_ID, zbirnaID, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA_IZVORI, COL_ZBI_OTPREMNICA_ID, _
                        otpremnicaID, SRC

    BuildZbirnaIzvorRowData = rowData
End Function

Private Function BuildZbirnaStavkaRowData(ByVal stavkaID As String, _
                                          ByVal zbirnaID As String, _
                                          ByVal redniBroj As Long, _
                                          ByVal klasa As String, _
                                          ByVal kolicina As Double, _
                                          ByVal kolAmb As Double) As Variant
    Const SRC As String = "BuildZbirnaStavkaRowData"

    Dim colCount As Long
    colCount = TabelaBrojKolona(TBL_ZBIRNA_STAVKE)

    If colCount <= 0 Then
        Err.Raise vbObjectError + 1248, SRC, _
                  "Ne mogu da odredim broj kolona za tblZbirnaStavke."
    End If

    Dim rowData() As Variant
    ReDim rowData(0 To colCount - 1)

    SetRowValueByColumn rowData, TBL_ZBIRNA_STAVKE, COL_ZBS_ID, stavkaID, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA_STAVKE, COL_ZBS_ZBIRNA_ID, zbirnaID, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA_STAVKE, COL_ZBS_RB, redniBroj, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA_STAVKE, COL_ZBS_KLASA, klasa, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA_STAVKE, COL_ZBS_KOLICINA, kolicina, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA_STAVKE, COL_ZBS_KOL_AMB, kolAmb, SRC

    BuildZbirnaStavkaRowData = rowData
End Function

' --- citanje DTO-a ----------------------------------------------------------
'
' Nedostajuci kljuc je GRESKA, ne prazna vrednost: Dictionary(k) nad nepostojecim
' kljucem tiho vraca Empty i doda kljuc, pa bi tipfeler u imenu polja prosao kao
' "korisnik nije uneo".

' Sve sto nije na spisku je greska.
'
' Provera obaveznih kljuceva hvata samo pola problema: tipfeler u OPCIONOM polju
' ("Hladnjca") ne obara nista, samo tiho ostavi prazno. Zato spisak.
'
' VrstaVoca / SortaVoca / TipAmbalaze NISU na spisku namerno -- izvode se iz
' izvornih otpremnica. Pozivalac koji ih salje pravi drugi izvor istine.
Private Function HdrKljucPoznat(ByVal kljuc As String) As Boolean
    Select Case LCase$(Trim$(kljuc))
        Case "datum", "vozacid", "brojzbirne", "kupacid", "hladnjaca", "pogon"
            HdrKljucPoznat = True
    End Select
End Function

Private Sub HdrProveriKljuceve(ByVal h As Object, ByVal src As String)
    Dim kljuc As Variant

    For Each kljuc In h.Keys
        If Not HdrKljucPoznat(CStr(kljuc)) Then
            Err.Raise vbObjectError + 1249, src, _
                      "Header ima nepoznat kljuc: " & CStr(kljuc) & _
                      ". Dozvoljeni: Datum, VozacID, BrojZbirne, KupacID, " & _
                      "Hladnjaca, Pogon. Vrsta/sorta/tip ambalaze dolaze iz otpremnica."
        End If
    Next kljuc
End Sub

' ZAGLAVLJE ZBIRNE: PROCITAJ I PROVERI DA VEZE POSTOJE (review #383, P2).
'
' Do sada je zbirna jedina od tri dokumenta proveravala samo da su polja
' NEPRAZNA. Otpremnica i otkup odavno traze da red STVARNO postoji, uz isti
' razlog koji vazi i ovde: neprazan string nije dokaz da red postoji, a slomljena
' veza se vidi tek kad je neko spoji -- tada je dokument vec IZDATO i finalan.
'
' Konkretan put kojim je to moglo da udje: MALINA_DEFAULT_KUPAC sa typo-om ili
' zastarelim ID-em. Auto-zbirna je proveravala samo Len() > 0, LookupValue za
' hladnjacu je za nepostojeceg kupca vracao prazno bez greske, i nastajala bi
' finalna zbirna sa KupacID-em koji nema pokrice.
'
' Kapija stoji OVDE, a ne u auto-putu: svaki pozivalac zbirnog pisca dobija isti
' razlog, pa se F3, uvoz i automatika ne mogu raziici.
Private Sub ZbrHdrCitajIProveri(ByVal h As Object, ByVal SRC As String, _
                                ByRef datum As Date, ByRef vozacID As String, _
                                ByRef brojZbirne As String, ByRef kupacID As String)
    datum = HdrDatum(h, "Datum", SRC)
    vozacID = HdrObavezan(h, "VozacID", SRC)
    brojZbirne = HdrObavezan(h, "BrojZbirne", SRC)
    kupacID = HdrObavezan(h, "KupacID", SRC)

    RequireTacnoJedan TBL_VOZACI, COL_VOZ_ID, vozacID, "VozacID", SRC
    RequireTacnoJedan TBL_KUPCI, COL_KUP_ID, kupacID, "KupacID", SRC
End Sub

Private Function HdrObavezan(ByVal h As Object, ByVal kljuc As String, _
                             ByVal src As String) As String
    If Not h.Exists(kljuc) Then
        Err.Raise vbObjectError + 1250, src, _
                  "Header nema obavezan kljuc: " & kljuc
    End If

    HdrObavezan = Trim$(NzToText(h(kljuc)))

    If Len(HdrObavezan) = 0 Then
        Err.Raise vbObjectError + 1251, src, _
                  "Header polje je prazno: " & kljuc
    End If
End Function

Private Function HdrOpcion(ByVal h As Object, ByVal kljuc As String) As String
    If h.Exists(kljuc) Then HdrOpcion = Trim$(NzToText(h(kljuc)))
End Function

Private Function HdrDatum(ByVal h As Object, ByVal kljuc As String, _
                          ByVal src As String) As Date
    If Not h.Exists(kljuc) Then
        Err.Raise vbObjectError + 1252, src, _
                  "Header nema obavezan kljuc: " & kljuc
    End If

    If Not IsDate(h(kljuc)) Then
        Err.Raise vbObjectError + 1253, src, _
                  "Header polje nije datum: " & kljuc
    End If

    HdrDatum = CDate(h(kljuc))
End Function

Private Function StavkaVrednost(ByVal s As Object, ByVal kljuc As String, _
                                ByVal idx As Long, ByVal src As String) As Variant
    If Not s.Exists(kljuc) Then
        Err.Raise vbObjectError + 1254, src, _
                  "Stavka " & CStr(idx) & " nema kljuc: " & kljuc
    End If

    StavkaVrednost = s(kljuc)
End Function

Private Function StavkaBroj(ByVal s As Object, ByVal kljuc As String, _
                            ByVal idx As Long, ByVal src As String) As Double
    Dim v As Variant
    v = StavkaVrednost(s, kljuc, idx, src)

    If Not IsNumeric(v) Then
        Err.Raise vbObjectError + 1255, src, _
                  "Stavka " & CStr(idx) & ", polje " & kljuc & _
                  " nije broj: " & NzToText(v)
    End If

    StavkaBroj = CDbl(v)
End Function

Public Function GetZbirnaByKupac(ByVal kupacID As String, _
                                  Optional ByVal datumOd As Date = 0, _
                                  Optional ByVal datumDo As Date = 0) As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_ZBIRNA)

    If IsEmpty(data) Then
        GetZbirnaByKupac = Empty
        Exit Function
    End If
    
    data = ExcludeStornirano(data, TBL_ZBIRNA)

    If IsEmpty(data) Then
        GetZbirnaByKupac = Empty
        Exit Function
    End If

    Dim filters As New Collection
    Dim fp As clsFilterParam

    Set fp = New clsFilterParam
    fp.Init RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KUPAC, _
            "modDokumenta.GetZbirnaByKupac"), "=", kupacID
    filters.Add fp

    If datumOd > 0 And datumDo > 0 Then
        Set fp = New clsFilterParam
        fp.Init RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_DATUM, _
                "modDokumenta.GetZbirnaByKupac"), "BETWEEN", datumOd, datumDo
        filters.Add fp
    End If

    GetZbirnaByKupac = FilterArray(data, filters)
    Exit Function

EH:
    LogErr "modDokumenta.GetZbirnaByKupac"
    GetZbirnaByKupac = Empty
End Function

' ============================================================
' OTPREMNICA -- header + stavke + clanstvo (skela, PR5)
' ============================================================
'
' RAZLIKA U ODNOSU NA ZBIRNU I OTKUP: otpremnica ima PERSISTENTAN DRAFT, i
' njene stavke drafta su OCEKIVANJE -- ne izveden kes.
'
' Realan tok: operater otvori otpremnicu i prijavi sta ona nosi, pa pod njom
' unosi otkupne listove gledajuci ocekivano / povezano / preostalo.
'
'   stavke drafta   OCEKIVANO   sta je vozac/operater prijavio
'   clanstvo        POVEZANO    SUM nad otkupnim stavkama clanova
'   preostalo       = ocekivano - povezano, po klasi
'
' To NISU dva izvora istine nego dve razlicite cinjenice, i njihov mismatch je
' bas ono zbog cega panel postoji. Danas ocekivanje zivi na Otpremnica.Kolicina
' i cita ga cetiri sposobnosti panela (modOtkupBlok:223, 262, 500, 1384); posto
' ta kolona u ciljnom modelu odlazi na stavku, ocekivanje bez stavki drafta
' nema gde da zivi (DOCUMENT_HEADER_LINES S4.2a, REFAKTOR S13b).
'
' Pri izdavanju stavke PRESTAJU da budu ocekivanje i postaju sadrzaj verzije.
'
'   CreateOtpremnicaDraft_TX(h, ocekivano)   -> OTP-...  DRAFT + stavke
'   UpdateOtpremnicaDraft_TX(otpID, h, ocekivano)   samo DRAFT
'   DodajOtpremnicaIzvor_TX(otpID, otkupID)         samo DRAFT
'   UkloniOtpremnicaIzvor_TX(otpID, otkupID)        samo DRAFT
'   GetOtpremnicaProgress(otpID)             -> ocekivano/povezano/preostalo
'   IzdajOtpremnicu_TX(otpID)                -> revalidacija + jednakost + IZDATO
'
'   CreateOtpremnicaIzIzvora_TX(h, izvori)   jedan potez, JEDNA transakcija
'
' Jednopotezni ulaz sme da IZVEDE ocekivanje iz izvora jer tu nezavisnog
' operaterskog ocekivanja nema -- auto-lanac i PWA ne prijavljuju sta nose, oni
' to znaju. Rucni tok bez ocekivanja bi ostao bez svoje jedine kontrole.
'
' KULTURA SE PRIMA NA DRAFTU, ne izvodi pri izdavanju. Smer podataka je
' otpremnica -> otkup: panel prefiluje formu otkupa vrstom, sortom, stanicom,
' vozacem i cenom (modOtkupBlok.PrefillLeftForm:706-730). Otpremnica koja
' kulturu saznaje tek pri izdavanju ne bi imala cime da prefiluje prvi otkup.
' VrstaVoca/SortaVoca su snapshot te kulture; TipAmbalaze se izvodi iz izvora
' (operater bira gajbu po otkupu, ne po kulturi).
'
' IZVOR MORA BITI PO NOVOM MODELU. Kolicine se citaju iz tblOtkupStavke; otkup
' koga je napisao stari writer nema stavke i bice odbijen.
'
' Otkup.OtpremnicaID se NE dira -- 39 ne-test citalaca, kolona ide u PR7.
'
' Header (h) -- Scripting.Dictionary, obavezni kljucevi:
'   Datum, StanicaID, VozacID, KulturaID, BrojOtpremnice, TipAmbalaze
'
' Ocekivano -- Collection diktova; spisak kljuceva je ZATVOREN:
'   Klasa (I ili II), Kolicina (> 0), KolAmbalaze (>= 0, ceo broj)
'   PredlogCena (opciona, >= 0) -- izricito NE-FINANSIJSKO polje: predlog za
'   prefill otkupnih blokova TE KLASE. Do S3a je stajao na zaglavlju kao jedan
'   broj, pa je otpremnica sa dve klase obe prefilovala istom cenom.
Public Function CreateOtpremnicaDraft_TX(ByVal h As Object, _
                                         ByVal ocekivano As Collection, _
                                         Optional ByRef outGreska As String) As String
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    outGreska = ""

    On Error GoTo EH

    modSchema.SchemaReadyOrFail "CreateOtpremnicaDraft_TX", _
        TBL_OTPREMNICA & "|" & TBL_OTPREMNICA_STAVKE & "|" & TBL_KULTURE

    ' Nothing je PRIVATAN signal jednopoteznog puta ("ocekivanje izvodim iz
    ' izvora kasnije"), i ne sme da procuri u javni rucni API. Bez ove kapije
    ' je CreateOtpremnicaDraft_TX(h, Nothing) pravio validan DRAFT BEZ
    ' ocekivanja -- dakle draft koji nema sta da meri, a izgleda ispravno.
    If ocekivano Is Nothing Then
        Err.Raise vbObjectError + 1329, "CreateOtpremnicaDraft_TX", _
                  "Ocekivanje nije prosledjeno. Rucni draft mora da prijavi sta " & _
                  "otpremnica nosi; za automatski tok koristi CreateOtpremnicaIzIzvora_TX."
    End If

    tx.BeginTx
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_OTPREMNICA_STAVKE

    CreateOtpremnicaDraft_TX = OtpNapraviDraft(h, ocekivano)

    If CreateOtpremnicaDraft_TX = "" Then
        Err.Raise vbObjectError + 1280, "CreateOtpremnicaDraft_TX", _
                  "OtpNapraviDraft nije vratio OtpremnicaID."
    End If

    tx.CommitTx
    Set tx = Nothing
    Exit Function

EH:
    outGreska = OtpPadTransakcije(tx, "CreateOtpremnicaDraft_TX", _
                                  CreateOtpremnicaDraft_TX)
    CreateOtpremnicaDraft_TX = ""
End Function

' Ispravka drafta pre izdavanja: operater je pogresio broj, ili se teret
' promenio. Menja se i zaglavlje i ocekivanje -- clanstvo ostaje.
Public Function UpdateOtpremnicaDraft_TX(ByVal otpremnicaID As String, _
                                         ByVal h As Object, _
                                         ByVal ocekivano As Collection, _
                                         Optional ByRef outGreska As String) As Boolean
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    outGreska = ""

    On Error GoTo EH

    modSchema.SchemaReadyOrFail "UpdateOtpremnicaDraft_TX", _
        TBL_OTPREMNICA & "|" & TBL_OTPREMNICA_STAVKE & "|" & TBL_KULTURE

    tx.BeginTx
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_OTPREMNICA_STAVKE

    OtpIzmeniDraft otpremnicaID, h, ocekivano

    tx.CommitTx
    Set tx = Nothing
    UpdateOtpremnicaDraft_TX = True
    Exit Function

EH:
    outGreska = OtpPadTransakcije(tx, "UpdateOtpremnicaDraft_TX", otpremnicaID)
    UpdateOtpremnicaDraft_TX = False
End Function

Public Function DodajOtpremnicaIzvor_TX(ByVal otpremnicaID As String, _
                                        ByVal otkupID As String, _
                                        Optional ByRef outGreska As String) As Boolean
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    outGreska = ""

    On Error GoTo EH

    modSchema.SchemaReadyOrFail "DodajOtpremnicaIzvor_TX", _
        TBL_OTPREMNICA & "|" & TBL_OTPREMNICA_IZVORI & "|" & TBL_OTKUP & _
        "|" & TBL_OTKUP_STAVKE

    tx.BeginTx
    tx.AddTableSnapshot TBL_OTPREMNICA_IZVORI

    OtpRequireIzvorValjan otpremnicaID, otkupID, "OtpDodajIzvor", True
    OtpUpisiClanstvo otpremnicaID, otkupID, "OtpDodajIzvor"

    tx.CommitTx
    Set tx = Nothing
    DodajOtpremnicaIzvor_TX = True
    Exit Function

EH:
    outGreska = OtpPadTransakcije(tx, "DodajOtpremnicaIzvor_TX", otpremnicaID)
    DodajOtpremnicaIzvor_TX = False
End Function

Public Function UkloniOtpremnicaIzvor_TX(ByVal otpremnicaID As String, _
                                         ByVal otkupID As String, _
                                         Optional ByRef outGreska As String) As Boolean
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    outGreska = ""

    On Error GoTo EH

    modSchema.SchemaReadyOrFail "UkloniOtpremnicaIzvor_TX", _
        TBL_OTPREMNICA & "|" & TBL_OTPREMNICA_IZVORI

    tx.BeginTx
    tx.AddTableSnapshot TBL_OTPREMNICA_IZVORI

    OtpUkloniIzvor otpremnicaID, otkupID

    tx.CommitTx
    Set tx = Nothing
    UkloniOtpremnicaIzvor_TX = True
    Exit Function

EH:
    outGreska = OtpPadTransakcije(tx, "UkloniOtpremnicaIzvor_TX", otpremnicaID)
    UkloniOtpremnicaIzvor_TX = False
End Function

Public Function IzdajOtpremnicu_TX(ByVal otpremnicaID As String, _
                                   Optional ByRef outGreska As String) As Boolean
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    outGreska = ""

    On Error GoTo EH

    modSchema.SchemaReadyOrFail "IzdajOtpremnicu_TX", _
        TBL_OTPREMNICA & "|" & TBL_OTPREMNICA_STAVKE & "|" & _
        TBL_OTPREMNICA_IZVORI & "|" & TBL_OTKUP & "|" & TBL_OTKUP_STAVKE & _
        "|" & TBL_AMBALAZA

    tx.BeginTx
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_OTPREMNICA_STAVKE
    ' Izdavanje knjizi izlaz gajbi (OtpKnjiziAmbalazu), pa tabela mora da ume da
    ' se vrati: pad na promeni statusa POSLE knjizenja bi ostavio gajbe skinute
    ' sa stanice, a otpremnicu neizdatu.
    tx.AddTableSnapshot TBL_AMBALAZA

    OtpIzdaj otpremnicaID

    tx.CommitTx
    Set tx = Nothing
    IzdajOtpremnicu_TX = True
    Exit Function

EH:
    outGreska = OtpPadTransakcije(tx, "IzdajOtpremnicu_TX", otpremnicaID)
    IzdajOtpremnicu_TX = False
End Function

' Jedan potez: draft + izvori + izdavanje, u JEDNOJ transakciji.
'
' Ocekivanje se IZVODI iz izvora -- ovde nezavisnog operaterskog ocekivanja
' nema, pa bi trazenje jednakosti sa samim sobom bilo prazan ritual.
'
' Pad na trecem izvoru ne sme da ostavi otpremnicu sa dva: pola sastava je gore
' od nijednog, jer izgleda kao zavrsen dokument.
Public Function CreateOtpremnicaIzIzvora_TX(ByVal h As Object, _
                                            ByVal izvori As Collection, _
                                            Optional ByRef outGreska As String) As String
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    outGreska = ""

    On Error GoTo EH

    modSchema.SchemaReadyOrFail "CreateOtpremnicaIzIzvora_TX", _
        TBL_OTPREMNICA & "|" & TBL_OTPREMNICA_STAVKE & "|" & _
        TBL_OTPREMNICA_IZVORI & "|" & TBL_OTKUP & "|" & TBL_OTKUP_STAVKE & _
        "|" & TBL_KULTURE & "|" & TBL_AMBALAZA

    tx.BeginTx
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_OTPREMNICA_STAVKE
    tx.AddTableSnapshot TBL_OTPREMNICA_IZVORI
    ' Jedan potez se zavrsava izdavanjem, pa i on knjizi gajbe (v. OtpIzdaj).
    tx.AddTableSnapshot TBL_AMBALAZA

    If izvori Is Nothing Then
        Err.Raise vbObjectError + 1281, "CreateOtpremnicaIzIzvora_TX", _
                  "Izvori nisu prosledjeni."
    End If
    If izvori.count = 0 Then
        Err.Raise vbObjectError + 1312, "CreateOtpremnicaIzIzvora_TX", _
                  "Otpremnica nema nijedan izvor. Otpremnica bez otkupa nije isporuka."
    End If

    ' Draft bez ocekivanja, pa se ocekivanje izvede iz izvora tek kad su svi
    ' provereni i upisani -- inace bi prvi lose izvor ostavio pogresne stavke.
    CreateOtpremnicaIzIzvora_TX = OtpNapraviDraft(h, Nothing)

    Dim i As Long
    For i = 1 To izvori.count
        OtpRequireIzvorValjan CreateOtpremnicaIzIzvora_TX, _
                              Trim$(NzToText(izvori(i))), "OtpDodajIzvor", True
        OtpUpisiClanstvo CreateOtpremnicaIzIzvora_TX, _
                         Trim$(NzToText(izvori(i))), "OtpDodajIzvor"
    Next i

    OtpUpisiOcekivanjeIzIzvora CreateOtpremnicaIzIzvora_TX
    OtpIzdaj CreateOtpremnicaIzIzvora_TX

    tx.CommitTx
    Set tx = Nothing
    Exit Function

EH:
    outGreska = OtpPadTransakcije(tx, "CreateOtpremnicaIzIzvora_TX", _
                                  CreateOtpremnicaIzIzvora_TX)
    CreateOtpremnicaIzIzvora_TX = ""
End Function

' ISPRAVKA IZDATE OTPREMNICE -- JEDAN POTEZ (S3c).
'
' Stara se stornira, a nova nastaje kao NACRT koji nasledjuje zaglavlje,
' ocekivanje i sve izvore stare. Sve u jednoj transakciji.
'
' Zasto jedan potez, a ne stari dvokorak ("storno odmah, zamenu snimi
' kasnije"): izmedju ta dva koraka je postojao prozor u kome je stara vec
' oborena a nove jos nema. Okvir je zato morao da pamti pending kontekst i da
' pravi MANUAL zadatke kad se u medjuvremenu nesto pomeri (blok ode na drugu
' otpremnicu, broj postane dvosmislen). Tog prozora ovde nema: padne li bilo
' koja kapija, transakcija vraca sve i nista nije stornirano.
'
' NACRT SE NE ISPRAVLJA -- njega menja UpdateOtpremnicaDraft_TX, bez storna i
' bez novog broja.
'
' Broj se NE nasledjuje: storno ne oslobadja broj (A9, modBrojevi), pa nova
' dobija sledeci slobodan iz istog niza (stanica, dan).
'
' Ocekivanje se prepisuje DOSLOVNO, ne izvodi iz izvora: izjednacavanje
' ocekivanog sa povezanim odbijeno je u S3b-2a (ocekivanje bi postalo
' formalnost). Ako je bas ocekivanje bilo pogresno, operater ga doradi u F2.
'
' Vraca OtpremnicaID NOVOG nacrta; "" = nista nije promenjeno, razlog je u
' outGreska.
Public Function IspravkaOtpremnice_TX(ByVal otpremnicaID As String, _
                                      Optional ByRef outGreska As String) As String
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    outGreska = ""

    On Error GoTo EH

    modSchema.SchemaReadyOrFail "IspravkaOtpremnice_TX", _
        TBL_OTPREMNICA & "|" & TBL_OTPREMNICA_STAVKE & "|" & _
        TBL_OTPREMNICA_IZVORI & "|" & TBL_OTKUP & "|" & TBL_OTKUP_STAVKE & _
        "|" & TBL_KULTURE & "|" & TBL_AMBALAZA

    tx.BeginTx
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_OTPREMNICA_STAVKE
    tx.AddTableSnapshot TBL_OTPREMNICA_IZVORI
    ' Storno stare vraca gajbe koje je njeno izdavanje knjizilo (OtpIzdaj).
    tx.AddTableSnapshot TBL_AMBALAZA

    IspravkaOtpremnice_TX = OtpIspravi(otpremnicaID)

    tx.CommitTx
    Set tx = Nothing
    Exit Function

EH:
    outGreska = OtpPadTransakcije(tx, "IspravkaOtpremnice_TX", otpremnicaID)
    IspravkaOtpremnice_TX = ""
End Function

' Read-model panela: ocekivano / povezano / preostalo po klasi.
'
' Vraca Dictionary klasa -> Dictionary sa kljucevima "ocekivano", "povezano",
' "preostalo", "ocekivanoAmb", "povezanoAmb", "preostaloAmb". Klase su UNIJA
' ocekivanih i povezanih -- klasa koja je povezana a nije ocekivana mora da se
' VIDI, inace bi visak bio nevidljiv do izdavanja.
Public Function GetOtpremnicaProgress(ByVal otpremnicaID As String) As Object
    Const SRC As String = "GetOtpremnicaProgress"

    Dim rez As Object
    Set rez = CreateObject("Scripting.Dictionary")
    Set GetOtpremnicaProgress = rez

    Dim ocek As Object, ocekAmb As Object
    Dim pov As Object, povAmb As Object, povBruto As Object, povBrutoPun As Object
    Set ocek = CreateObject("Scripting.Dictionary")
    Set ocekAmb = CreateObject("Scripting.Dictionary")

    OtpUcitajOcekivano otpremnicaID, ocek, ocekAmb, SRC
    OtpUcitajPovezano otpremnicaID, pov, povAmb, povBruto, povBrutoPun, SRC

    Dim sve As Object
    Set sve = CreateObject("Scripting.Dictionary")

    Dim k As Variant
    For Each k In ocek.Keys
        sve(CStr(k)) = True
    Next k
    For Each k In pov.Keys
        sve(CStr(k)) = True
    Next k

    Dim red As Object
    For Each k In KlaseUKanonskomRedu(sve)
        Set red = CreateObject("Scripting.Dictionary")
        red("ocekivano") = OtpBroj(ocek, CStr(k))
        red("povezano") = OtpBroj(pov, CStr(k))
        red("preostalo") = OtpBroj(ocek, CStr(k)) - OtpBroj(pov, CStr(k))
        red("ocekivanoAmb") = OtpBroj(ocekAmb, CStr(k))
        red("povezanoAmb") = OtpBroj(povAmb, CStr(k))
        red("preostaloAmb") = OtpBroj(ocekAmb, CStr(k)) - OtpBroj(povAmb, CStr(k))
        Set rez(CStr(k)) = red
    Next k
End Function

' --- kanonski citaoci stavki otpremnice (S3b) --------------------------------
'
' Zaglavlje otpremnice od S3a vise NE nosi Klasu, Kolicinu, KolAmbalaze ni
' Cenu -- CreateOtpremnicaDraft_TX ih ostavlja prazne, a ocekivanje po klasi
' pise u tblOtpremnicaStavke. Citaoci koji su te kolone i dalje citali (mreza
' F2, stampa, izvestaji, invarijanta) od S3a dobijaju PRAZNO. Ovde je jedno
' mesto sa kojeg svi citaju stavke, po istom obrascu koji otkup ima od S14.7
' (modOtkup.StavkeOtkupaRedovi / ZbirStavkiPoOtkupu).
'
' UGOVOR JE DOKUMENTSKI, NE STAVKARSKI (isti razlog kao kod otkupa, review
' #334 P1):
'   1) svaka stavka ima neprazan OtpremnicaID i zaglavlje TACNO JEDNOM;
'   2) svako zaglavlje sa OtpremnicaID-em ima BAR JEDNU stavku -- pisac to vec
'      trazi (OtpUpisiOcekivano odbija prazno ocekivanje), pa citalac koji bi
'      nulu vratio kao podatak laze o dokumentu koji pisac ne moze da napravi;
'   3) Kolicina > 0, Klasa kroz istu kapiju koju pisac zove (RequireValidKlasa),
'      KolAmbalaze ceo broj >= 0 (prazno = 0 gajbi).
'
' PredlogCena je JEDINI izuzetak od "vece od nule": ona je PREDLOG, ne knjizena
' cena (odluka 14.8 t. 2), pa prazna kolona znaci "nije predlozena" i to nije
' greska. Kad postoji, mora biti broj > 0 -- nula bi bila predlog da se roba da
' besplatno, a to niko nije uneo.
Public Function StavkeOtpremniceRedovi() As Variant
    Const SRC As String = "StavkeOtpremniceRedovi"

    StavkeOtpremniceRedovi = Empty

    Dim zagl As Object
    Set zagl = ZaglavljaOtpremnicePoID(SRC)
    RequireJedinstvenoZaglavljeOtpremnice zagl, SRC

    Dim imaStavku As Object
    Set imaStavku = CreateObject("Scripting.Dictionary")

    Dim d As Variant
    d = GetTableData(TBL_OTPREMNICA_STAVKE)

    Dim cOtp As Long, cRb As Long, cKl As Long, cKol As Long
    Dim cCena As Long, cAmb As Long, cId As Long, cBruto As Long
    Dim i As Long, n As Long, oid As String

    If IsArray(d) Then
        cOtp = RequireColumnIndex(TBL_OTPREMNICA_STAVKE, COL_OPS_OTPREMNICA_ID, SRC)
        cRb = RequireColumnIndex(TBL_OTPREMNICA_STAVKE, COL_OPS_RB, SRC)
        cKl = RequireColumnIndex(TBL_OTPREMNICA_STAVKE, COL_OPS_KLASA, SRC)
        cKol = RequireColumnIndex(TBL_OTPREMNICA_STAVKE, COL_OPS_KOLICINA, SRC)
        cCena = RequireColumnIndex(TBL_OTPREMNICA_STAVKE, COL_OPS_PREDLOG_CENA, SRC)
        cAmb = RequireColumnIndex(TBL_OTPREMNICA_STAVKE, COL_OPS_KOL_AMB, SRC)
        cId = RequireColumnIndex(TBL_OTPREMNICA_STAVKE, COL_OPS_ID, SRC)
        cBruto = RequireColumnIndex(TBL_OTPREMNICA_STAVKE, COL_OPS_BRUTO, SRC)

        ' Jedna stavka po KLASI po dokumentu -- isto sto pisac trazi
        ' (OtpUpisiOcekivano odbija dve stavke iste klase, review #362 P2).
        ' Bez ove kapije bi pokvaren dokument sa dve stavke klase I bio u
        ' izvestaju sabran kao 2 x I, a u stampi dao dva reda iste klase.
        Dim parKlasa As Object, kParKl As String
        Set parKlasa = CreateObject("Scripting.Dictionary")

        For i = 1 To UBound(d, 1)
            oid = Trim$(NzToText(d(i, cOtp)))
            RequireOtpStavkaUgovor zagl, oid, d(i, cKl), d(i, cKol), d(i, cCena), _
                                   d(i, cAmb), i, SRC
            kParKl = UCase$(oid) & "|" & UCase$(Trim$(NzToText(d(i, cKl))))
            If parKlasa.Exists(kParKl) Then
                Err.Raise vbObjectError + 1942, SRC, _
                          "Dve stavke iste klase na otpremnici: OtpremnicaID=" & oid & _
                          ", klasa " & Trim$(NzToText(d(i, cKl))) & "."
            End If
            parKlasa.Add kParKl, True
            imaStavku(oid) = True
            n = n + 1
        Next i
    End If

    RequireZaglavljaOtpremniceSaStavkama zagl, imaStavku, SRC

    If n = 0 Then Exit Function

    ' Kolone: 1 OtpremnicaID, 2 RedniBroj, 3 Klasa, 4 Kolicina, 5 PredlogCena,
    ' 6 KolAmbalaze, 7 OtpremnicaStavkaID, 8 BrutoKg. Nove kolone idu NA KRAJ:
    ' citaoci indeksiraju poziciono.
    Dim res() As Variant
    ReDim res(1 To n, 1 To 8)
    n = 0
    For i = 1 To UBound(d, 1)
        n = n + 1
        res(n, 1) = Trim$(NzToText(d(i, cOtp)))
        res(n, 2) = d(i, cRb)
        res(n, 3) = Trim$(NzToText(d(i, cKl)))
        res(n, 4) = CDbl(d(i, cKol))
        res(n, 5) = OtpDbl(d(i, cCena))
        res(n, 6) = OtpDbl(d(i, cAmb))
        res(n, 7) = Trim$(NzToText(d(i, cId)))
        res(n, 8) = OtpDbl(d(i, cBruto))
    Next i

    StavkeOtpremniceRedovi = res
End Function

' Zbir stavki po dokumentu. Kljuc je OtpremnicaID, vrednost
' Array(kg, Null, gajbe, klase):
'   kg       = SUM(Kolicina)
'   (1)      = Null -- mesto VREDNOSTI je namerno prazno (review #362, P1).
'              PredlogCena je predlog za prefill, izricito NE-finansijsko polje:
'              "nigde se vrednost otpremnice ne racuna kao Kolicina x Cena".
'              Vrednost otpremnice je vrednost njenih IZVORA --
'              VrednostIzvoraPoOtpremnici. Null (ne 0) da slucajna upotreba
'              padne, a ne da tiho sabere nulu.
'   gajbe    = SUM(KolAmbalaze)
'   klase    = klase stavki redom upisa, bez ponavljanja ("I, II")
'
' Oblik je isti kao kod ZbirStavkiPoOtkupu, jer mreza F2 za oba tipa cita iste
' pozicije (kg, gajbe, klase).
'
' SVAKI kanonski dokument JE u recniku -- zaglavlje bez stavki obara citaoca
' gore, pa "nema kljuca" ne moze da znaci 0 kg.
Public Function ZbirStavkiPoOtpremnici() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Set ZbirStavkiPoOtpremnici = dict

    Dim s As Variant
    s = StavkeOtpremniceRedovi()
    If Not IsArray(s) Then Exit Function

    Dim i As Long, oid As String, kl As String, rec As Variant
    For i = 1 To UBound(s, 1)
        oid = CStr(s(i, 1))
        If dict.Exists(oid) Then
            rec = dict(oid)
        Else
            rec = Array(0#, Null, 0#, "")
        End If
        rec(0) = CDbl(rec(0)) + CDbl(s(i, 4))
        rec(2) = CDbl(rec(2)) + CDbl(s(i, 6))
        kl = CStr(s(i, 3))
        If Len(kl) > 0 Then
            If InStr(1, ", " & CStr(rec(3)) & ", ", ", " & kl & ", ", vbBinaryCompare) = 0 Then
                If Len(CStr(rec(3))) > 0 Then rec(3) = CStr(rec(3)) & ", "
                rec(3) = CStr(rec(3)) & kl
            End If
        End If
        dict(oid) = rec
    Next i
End Function

' Zbir stavki JEDNE zbirne iz recnika -- nedostajuci kljuc je GRESKA.
'
' Postoji zato sto PORUKA pripada dokumentu: do review-a #370 je lista F8 za
' zbirnu isla kroz ZbirStavkiZaOtpremnicu, pa bi pad nad zbirnom rekao
' "Otpremnica nema nijednu stavku" i uputio na tblOtpremnicaStavke. Racun je bio
' tacan (strog lookup po kljucu), ali bi dijagnostika slala operatera i
' programera u pogresnu tabelu.
Public Function ZbirStavkiZaZbirnu(ByVal zbir As Object, _
                                   ByVal zbirnaID As String, _
                                   ByVal sourceName As String) As Variant
    If zbir Is Nothing Then
        Err.Raise vbObjectError + 1957, sourceName, _
                  "Zbir stavki zbirne nije izgradjen."
    End If

    Dim zid As String
    zid = Trim$(zbirnaID)

    If Len(zid) = 0 Then
        Err.Raise vbObjectError + 1953, sourceName, _
                  "Zaglavlje zbirne bez ZbirnaID-a se ne moze citati."
    End If

    If Not zbir.Exists(zid) Then
        Err.Raise vbObjectError + 1957, sourceName, _
                  "Zbirna nema nijednu stavku: " & zid & _
                  ". Kolicina dokumenta se racuna iz " & TBL_ZBIRNA_STAVKE & "."
    End If

    ZbirStavkiZaZbirnu = zbir(zid)
End Function

' Zbir stavki JEDNOG dokumenta iz recnika -- nedostajuci kljuc je GRESKA.
' Druga brana, kao ZbirStavkiZaOtkup: nijedan citalac ne sme da ima granu koja
' nulu vraca kao podatak.
Public Function ZbirStavkiZaOtpremnicu(ByVal zbir As Object, _
                                       ByVal otpremnicaID As String, _
                                       ByVal sourceName As String) As Variant
    If zbir Is Nothing Then
        Err.Raise vbObjectError + 1929, sourceName, _
                  "Zbir stavki otpremnice nije izgradjen."
    End If

    Dim oid As String
    oid = Trim$(otpremnicaID)

    If Len(oid) = 0 Then
        Err.Raise vbObjectError + 1927, sourceName, _
                  "Zaglavlje otpremnice bez OtpremnicaID-a se ne moze citati."
    End If

    If Not zbir.Exists(oid) Then
        Err.Raise vbObjectError + 1929, sourceName, _
                  "Otpremnica nema nijednu stavku: " & oid & _
                  ". Kolicina dokumenta se racuna iz " & TBL_OTPREMNICA_STAVKE & "."
    End If

    ZbirStavkiZaOtpremnicu = zbir(oid)
End Function

' Stavke grupisane po dokumentu -- za citaoce kojima jedan red izlaza pripada
' JEDNOJ klasi (stampa, izvestaj po otkupnom mestu). Vrednost je Collection
' redova iz StavkeOtpremniceRedovi, redom upisa.
Public Function StavkeOtpremnicePoDokumentu() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Set StavkeOtpremnicePoDokumentu = dict

    Dim s As Variant
    s = StavkeOtpremniceRedovi()
    If Not IsArray(s) Then Exit Function

    Dim i As Long, j As Long, oid As String, c As Collection, red As Variant
    For i = 1 To UBound(s, 1)
        oid = CStr(s(i, 1))
        If dict.Exists(oid) Then
            Set c = dict(oid)
        Else
            Set c = New Collection
            dict.Add oid, c
        End If
        ReDim red(1 To 8)
        For j = 1 To 8
            red(j) = s(i, j)
        Next j
        c.Add red
    Next i
End Function

' Stavke JEDNOG dokumenta -- nedostajuci kljuc je GRESKA, isti razlog kao gore.
Public Function StavkeZaOtpremnicu(ByVal poDok As Object, _
                                   ByVal otpremnicaID As String, _
                                   ByVal sourceName As String) As Collection
    If poDok Is Nothing Then
        Err.Raise vbObjectError + 1929, sourceName, _
                  "Stavke otpremnice nisu ucitane."
    End If

    Dim oid As String
    oid = Trim$(otpremnicaID)

    If Not poDok.Exists(oid) Then
        Err.Raise vbObjectError + 1929, sourceName, _
                  "Otpremnica nema nijednu stavku: " & oid & _
                  ". Stavke se citaju iz " & TBL_OTPREMNICA_STAVKE & "."
    End If

    Set StavkeZaOtpremnicu = poDok(oid)
End Function

' VREDNOST OTPREMNICE DOLAZI IZ IZVORA (review #362, P1).
'
' Vrednost robe na otpremnici je vrednost njenih IZVORNIH otkupa: zbir
' Kolicina x Cena njihovih stavki, dakle cene koje su stvarno placene. Dva bloka
' sa razlicitom cenom ostaju dve cene -- jedna PredlogCena ih ne sme zameniti.
'
' Kljucevi:
'   "OtpremnicaID"        -> ukupna vrednost izvora (Double)
'   "OtpremnicaID|KLASA"  -> Array(kg, vrednost) izvora te klase
'
' Clanstvo se cita iz kanona (tblOtpremnicaIzvori, bez storniranih otpremnica),
' stavke kroz strog citac otkupa. Otpremnica BEZ clanova nema kljuc: pitanje
' "koliko vredi" za nju nema odgovor -- strog pristupnik nize ga odbija.
Public Function VrednostIzvoraPoOtpremnici() As Object
    Const SRC As String = "VrednostIzvoraPoOtpremnici"

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set VrednostIzvoraPoOtpremnici = d

    Dim clanstvo As Object
    Set clanstvo = AktivnoOtpClanstvoPoKanonu(SRC)
    If clanstvo.count = 0 Then Exit Function

    Dim s As Variant
    s = modOtkup.StavkeOtkupaRedovi()
    If Not IsArray(s) Then Exit Function

    Dim i As Long, otkID As String, otp As String, kl As String
    Dim kol As Double, vr As Double, kKlasa As String, rec As Variant
    For i = 1 To UBound(s, 1)
        otkID = UCase$(Trim$(CStr(s(i, 1))))
        If clanstvo.Exists(otkID) Then
            otp = CStr(clanstvo(otkID))
            kl = UCase$(Trim$(CStr(s(i, 3))))
            kol = CDbl(s(i, 4))
            vr = kol * CDbl(s(i, 5))

            If d.Exists(otp) Then
                d(otp) = CDbl(d(otp)) + vr
            Else
                d.Add otp, vr
            End If

            kKlasa = otp & "|" & kl
            If d.Exists(kKlasa) Then
                rec = d(kKlasa)
            Else
                rec = Array(0#, 0#)
            End If
            rec(0) = CDbl(rec(0)) + kol
            rec(1) = CDbl(rec(1)) + vr
            d(kKlasa) = rec
        End If
    Next i
End Function

' Vrednost izvora JEDNE otpremnice -- nedostajuci kljuc je GRESKA. Zove se samo
' za IZDATE otpremnice, a IzdajOtpremnicu_TX ne izdaje otpremnicu bez izvora:
' izdata bez izvora je kvar, i nula bi ga sakrila.
Public Function VrednostIzvoraZaOtpremnicu(ByVal vrednosti As Object, _
                                           ByVal otpremnicaID As String, _
                                           ByVal sourceName As String) As Double
    Dim oid As String
    oid = Trim$(otpremnicaID)
    If vrednosti Is Nothing Then
        Err.Raise vbObjectError + 1946, sourceName, "Vrednosti izvora nisu izgradjene."
    End If
    If Not vrednosti.Exists(oid) Then
        Err.Raise vbObjectError + 1946, sourceName, _
                  "Otpremnica nema nijedan izvor: " & oid & ". Vrednost otpremnice je " & _
                  "vrednost njenih izvornih otkupa -- izdata otpremnica bez izvora je kvar."
    End If
    VrednostIzvoraZaOtpremnicu = CDbl(vrednosti(oid))
End Function

' Isto, po klasi: Array(kg, vrednost) izvora te klase. Za stampu.
Public Function VrednostIzvoraKlase(ByVal vrednosti As Object, _
                                    ByVal otpremnicaID As String, _
                                    ByVal klasa As String, _
                                    ByVal sourceName As String) As Variant
    Dim k As String
    k = Trim$(otpremnicaID) & "|" & UCase$(Trim$(klasa))
    If vrednosti Is Nothing Then
        Err.Raise vbObjectError + 1946, sourceName, "Vrednosti izvora nisu izgradjene."
    End If
    If Not vrednosti.Exists(k) Then
        Err.Raise vbObjectError + 1946, sourceName, _
                  "Otpremnica " & Trim$(otpremnicaID) & " nema izvor klase " & klasa & "."
    End If
    VrednostIzvoraKlase = vrednosti(k)
End Function

Private Function ZaglavljaOtpremnicePoID(ByVal sourceName As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Set ZaglavljaOtpremnicePoID = dict

    Dim d As Variant
    d = GetTableData(TBL_OTPREMNICA)
    If Not IsArray(d) Then Exit Function

    Dim cId As Long
    cId = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_ID, sourceName)

    Dim i As Long, oid As String
    For i = 1 To UBound(d, 1)
        oid = Trim$(NzToText(d(i, cId)))
        If Len(oid) > 0 Then
            If dict.Exists(oid) Then
                dict(oid) = CLng(dict(oid)) + 1
            Else
                dict.Add oid, 1&
            End If
        End If
    Next i
End Function

' JEDAN LOGICKI DOKUMENT = JEDAN ID. Dva zaglavlja sa istim OtpremnicaID
' obaraju citaoca PRE ijednog zbira -- dokument-level citaoci iteriraju
' zaglavlja, pa bi isti teret bio izbrojan dvaput.
Private Sub RequireJedinstvenoZaglavljeOtpremnice(ByVal zagl As Object, _
                                                  ByVal sourceName As String)
    Dim k As Variant
    For Each k In zagl.keys
        If CLng(zagl(k)) <> 1 Then
            Err.Raise vbObjectError + 1941, sourceName, _
                      "Zaglavlje otpremnice se ne nalazi tacno jednom: " & CStr(k) & _
                      " (headera: " & CStr(zagl(k)) & "). OtpremnicaID je identitet " & _
                      "dokumenta -- dva reda sa istim ID-em nisu dokument."
        End If
    Next k
End Sub

Private Sub RequireOtpStavkaUgovor(ByVal zagl As Object, ByVal oid As String, _
                                   ByVal klasa As Variant, ByVal kol As Variant, _
                                   ByVal cena As Variant, ByVal amb As Variant, _
                                   ByVal red As Long, ByVal sourceName As String)
    If Len(oid) = 0 Then
        Err.Raise vbObjectError + 1927, sourceName, _
                  "Stavka otpremnice bez OtpremnicaID-a: " & TBL_OTPREMNICA_STAVKE & _
                  ", red " & CStr(red) & ". Stavka bez dokumenta se ne moze citati."
    End If

    If Not zagl.Exists(oid) Then
        Err.Raise vbObjectError + 1928, sourceName, _
                  "Zaglavlje otpremnice ne postoji: " & oid & ", stavka u redu " & _
                  CStr(red) & ". Stavka bez dokumenta se ne moze citati."
    End If

    If Not IsNumeric(kol) Then
        Err.Raise vbObjectError + 1925, sourceName, _
                  "Kolicina stavke nije brojcana: OtpremnicaID=" & oid & "."
    End If

    If CDbl(kol) <= 0 Then
        Err.Raise vbObjectError + 1926, sourceName, _
                  "Kolicina stavke mora biti veca od nule: OtpremnicaID=" & oid & "."
    End If

    ' Prazna PredlogCena je uredno stanje (predlog nije dat). Upisana nula ili
    ' tekst nisu -- to je pokvaren podatak, ne odsustvo predloga.
    If Len(Trim$(NzToText(cena))) > 0 Then
        If Not IsNumeric(cena) Then
            Err.Raise vbObjectError + 1925, sourceName, _
                      "PredlogCena stavke nije brojcana: OtpremnicaID=" & oid & "."
        End If
        If CDbl(cena) <= 0 Then
            Err.Raise vbObjectError + 1926, sourceName, _
                      "PredlogCena stavke mora biti veca od nule kad je upisana: " & _
                      "OtpremnicaID=" & oid & "."
        End If
    End If

    ' Gajbe su KOMADI: prazno je 0, a nebrojcano, negativno i decimalno je kvar --
    ' isto sto pisac odbija (OtpUpisiOcekivano -> RequireCeoBroj). Citalac koji bi
    ' 1,5 gajbu sabrao razisao bi se sa piscem, a tiho zaokruzivanje sakriva izvor.
    If Len(Trim$(NzToText(amb))) > 0 Then
        If Not IsNumeric(amb) Then
            Err.Raise vbObjectError + 1925, sourceName, _
                      "KolAmbalaze stavke nije brojcana: OtpremnicaID=" & oid & "."
        End If
        If CDbl(amb) < 0 Then
            Err.Raise vbObjectError + 1926, sourceName, _
                      "KolAmbalaze stavke ne sme biti negativna: OtpremnicaID=" & oid & "."
        End If
        RequireCeoBroj CDbl(amb), "KolAmbalaze stavke (OtpremnicaID=" & oid & ")", _
                       sourceName
    End If

    RequireValidKlasa Trim$(NzToText(klasa)), _
                      sourceName & " (OtpremnicaID=" & oid & ")"
End Sub

Private Sub RequireZaglavljaOtpremniceSaStavkama(ByVal zagl As Object, _
                                                 ByVal imaStavku As Object, _
                                                 ByVal sourceName As String)
    Dim k As Variant
    For Each k In zagl.keys
        If Not imaStavku.Exists(CStr(k)) Then
            Err.Raise vbObjectError + 1929, sourceName, _
                      "Otpremnica nema nijednu stavku: " & CStr(k) & _
                      ". Kolicina dokumenta se racuna iz " & TBL_OTPREMNICA_STAVKE & "."
        End If
    Next k
End Sub

' Jedan EH za sve ulaze: monitoring, rollback i poruka su im isti, a sest
' kopija bi bilo sest mesta na kojima se rollback moze zaboraviti.
Private Function OtpPadTransakcije(ByRef tx As clsTransaction, _
                                   ByVal ulaz As String, _
                                   ByVal entitetID As String) As String
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogError ulaz, errDesc, errNum
    Monitor_Error _
        moduleName:="modDokumenta", _
        procedureName:=ulaz, _
        entityType:="Otpremnica", _
        entityID:=entitetID, _
        correlationId:=entitetID, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="DOKUMENT_SAVE_FAIL", _
        severity:="ERROR", _
        message:=ulaz & " failed. Error=" & errDesc, _
        userId:="Operator", _
        moduleName:="modDokumenta", _
        procedureName:=ulaz, _
        entityType:="Otpremnica", _
        entityID:=entitetID, _
        correlationId:=entitetID

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    OtpPadTransakcije = errDesc
    PrintTxFailure ulaz, errSrc, errNum, errDesc
End Function

' --- core: draft ------------------------------------------------------------
Private Function OtpNapraviDraft(ByVal h As Object, _
                                 ByVal ocekivano As Collection) As String
    Const SRC As String = "OtpNapraviDraft"

    If h Is Nothing Then
        Err.Raise vbObjectError + 1282, SRC, "Header nije prosledjen."
    End If

    RequireColumnIndex TBL_OTPREMNICA, COL_OTP_ID, SRC
    RequireColumnIndex TBL_OTPREMNICA, COL_OTP_DATUM, SRC
    RequireColumnIndex TBL_OTPREMNICA, COL_OTP_STANICA, SRC
    RequireColumnIndex TBL_OTPREMNICA, COL_OTP_VOZAC, SRC
    RequireColumnIndex TBL_OTPREMNICA, COL_OTP_KULTURA, SRC
    RequireColumnIndex TBL_OTPREMNICA, COL_OTP_BROJ, SRC
    RequireColumnIndex TBL_OTPREMNICA, COL_TRACE_IZDATO_STATUS, SRC

    OtpHdrProveriKljuceve h, SRC

    Dim datum As Date
    Dim stanicaID As String, vozacID As String, brojOtp As String, kulturaID As String
    Dim tipAmb As String

    datum = HdrDatum(h, "Datum", SRC)
    stanicaID = HdrObavezan(h, "StanicaID", SRC)
    vozacID = HdrObavezan(h, "VozacID", SRC)
    kulturaID = HdrObavezan(h, "KulturaID", SRC)
    brojOtp = HdrObavezan(h, "BrojOtpremnice", SRC)

    ' TipAmbalaze je HEADER cinjenica, primljena pri otvaranju. Kljuc je obavezan,
    ' vrednost sme prazna -- obaveznost zavisi od ocekivane ambalaze i proverava se
    ' u OtpUpisiOcekivano. Bez ovoga bi ocekivanje glasilo "50 gajbi" a tek bi prvi
    ' izvor rekao KOJIH 50; zatecen posao to vec resava drugacije -- legacy
    ' SaveOtpremnicaMulti_TX prima tipAmb JEDNOM, kao header podatak, i bas njime
    ' knjizi ambalazu pri nastanku otpremnice (modDokumenta:382 TrackAmbalaza).
    tipAmb = OtpHdrObavezanKljuc(h, "TipAmbalaze", SRC)

    ' FK-ovi otpremnice. Isti razlog kao kod otkupa (S4.1f): neprazan string nije
    ' dokaz da red postoji, a slomljena veza se vidi tek kad je neko spoji.
    RequireTacnoJedan TBL_STANICE, COL_STA_ID, stanicaID, "StanicaID", SRC
    RequireTacnoJedan TBL_VOZACI, COL_VOZ_ID, vozacID, "VozacID", SRC
    RequireTacnoJedan TBL_KULTURE, COL_KUL_ID, kulturaID, "KulturaID", SRC

    ' Draft sme da promeni stanicu, datum i broj u istom potezu, pa broj
    ' koji je bio tacan postane tudj bez ijedne druge provere.
    modBrojevi.RequireBrojUKontekstu modBrojevi.KIND_OTP, stanicaID, datum, _
                                     brojOtp, SRC
    modBrojevi.RequireBrojSlobodanUNizu modBrojevi.KIND_OTP, stanicaID, datum, _
                                        brojOtp, SRC

    Dim otpID As String
    otpID = NewEntityID("OTP-")

    If otpID = "" Then
        Err.Raise vbObjectError + 1284, SRC, "NewEntityID nije vratio OtpremnicaID."
    End If

    Dim rowData As Variant
    rowData = BuildOtpremnicaHeaderRowData(otpID, datum, stanicaID, vozacID, _
                                           kulturaID, tipAmb, brojOtp)

    If AppendRow(TBL_OTPREMNICA, rowData) <= 0 Then
        Err.Raise vbObjectError + 1285, SRC, _
                  "AppendRow nije upisao header u tblOtpremnica."
    End If

    ' Nothing = ocekivanje se izvodi kasnije (jednopotezni ulaz). Prazna
    ' Collection je NESTO DRUGO: rucni draft bez ijedne stavke, i to je greska.
    If Not ocekivano Is Nothing Then
        OtpUpisiOcekivano otpID, ocekivano, SRC
    End If

    OtpNapraviDraft = otpID
End Function

Private Sub OtpIzmeniDraft(ByVal otpremnicaID As String, ByVal h As Object, _
                           ByVal ocekivano As Collection)
    Const SRC As String = "OtpIzmeniDraft"

    If h Is Nothing Then
        Err.Raise vbObjectError + 1313, SRC, "Header nije prosledjen."
    End If

    Dim rOtp As Long
    rOtp = OtpRedHeadera(otpremnicaID, SRC)
    RequireOtpDraft otpremnicaID, rOtp, SRC

    OtpHdrProveriKljuceve h, SRC

    Dim datum As Date
    Dim stanicaID As String, vozacID As String, brojOtp As String, kulturaID As String

    datum = HdrDatum(h, "Datum", SRC)
    stanicaID = HdrObavezan(h, "StanicaID", SRC)
    vozacID = HdrObavezan(h, "VozacID", SRC)
    kulturaID = HdrObavezan(h, "KulturaID", SRC)
    brojOtp = HdrObavezan(h, "BrojOtpremnice", SRC)

    RequireTacnoJedan TBL_STANICE, COL_STA_ID, stanicaID, "StanicaID", SRC
    RequireTacnoJedan TBL_VOZACI, COL_VOZ_ID, vozacID, "VozacID", SRC
    RequireTacnoJedan TBL_KULTURE, COL_KUL_ID, kulturaID, "KulturaID", SRC

    ' Draft sme da promeni stanicu, datum i broj u istom potezu, pa broj
    ' koji je bio tacan postane tudj bez ijedne druge provere.
    modBrojevi.RequireBrojUKontekstu modBrojevi.KIND_OTP, stanicaID, datum, _
                                     brojOtp, SRC

    ' Draft sme da ZADRZI svoj broj i dan, a ne sme da preuzme tudj: sopstveni
    ' red se izuzima po ID-u, ne po datumu, jer header moze ostati na istom danu.
    modBrojevi.RequireBrojSlobodanUNizu modBrojevi.KIND_OTP, stanicaID, datum, _
                                        brojOtp, SRC, otpremnicaID

    ' Zaglavlje se menja tek posto se zna da je novo ocekivanje ispravno --
    ' inace bi lose ocekivanje ostavilo pola izmenjen header.
    Dim staroOcek As Collection
    Set staroOcek = ocekivano
    If staroOcek Is Nothing Then
        Err.Raise vbObjectError + 1314, SRC, _
                  "Ocekivanje nije prosledjeno. Draft bez ocekivanja nema sta da meri."
    End If

    ' Zaglavlje ide PRE ocekivanja: obaveznost tipa ambalaze se sudi prema NOVOM
    ' tipu, ne prema starom. Sve je u istoj transakciji, pa pad bilo gde vraca sve.
    RequireUpdateCell TBL_OTPREMNICA, rOtp, COL_OTP_TIP_AMB, _
                      OtpHdrObavezanKljuc(h, "TipAmbalaze", SRC), SRC
    RequireUpdateCell TBL_OTPREMNICA, rOtp, COL_OTP_DATUM, datum, SRC
    RequireUpdateCell TBL_OTPREMNICA, rOtp, COL_OTP_STANICA, stanicaID, SRC
    RequireUpdateCell TBL_OTPREMNICA, rOtp, COL_OTP_VOZAC, vozacID, SRC
    RequireUpdateCell TBL_OTPREMNICA, rOtp, COL_OTP_KULTURA, kulturaID, SRC
    RequireUpdateCell TBL_OTPREMNICA, rOtp, COL_OTP_BROJ, brojOtp, SRC
    RequireUpdateCell TBL_OTPREMNICA, rOtp, COL_OTP_VRSTA, _
        Trim$(NzToText(LookupValue(TBL_KULTURE, COL_KUL_ID, kulturaID, COL_KUL_VRSTA))), SRC
    RequireUpdateCell TBL_OTPREMNICA, rOtp, COL_OTP_SORTA, _
        Trim$(NzToText(LookupValue(TBL_KULTURE, COL_KUL_ID, kulturaID, COL_KUL_SORTA))), SRC

    OtpObrisiOcekivano otpremnicaID, SRC
    OtpUpisiOcekivano otpremnicaID, staroOcek, SRC

    ' Izmena zaglavlja sme da pokvari VEC VALIDNO clanstvo: draft sa stanicom ST1
    ' i clanom sa ST1 posle prebacivanja na ST2 nosi clana koga Dodaj nikad ne bi
    ' primio. Izdavanje bi to kasnije uhvatilo, ali invarijanta ne sme da bude
    ' prekrsena IZMEDJU dva klika -- GetOtpremnicaProgress u medjuvremenu uredno
    ' racuna nevalidno clanstvo.
    '
    ' Provera ide POSLE upisa zaglavlja, nad NOVIM vrednostima; pad ovde rollback-uje
    ' ceo update, pa staro zaglavlje i staro ocekivanje ostaju netaknuti.
    Dim clanovi As Collection
    Set clanovi = OtpClanovi(otpremnicaID, SRC)

    Dim k As Long
    For k = 1 To clanovi.count
        OtpRequireIzvorValjan otpremnicaID, CStr(clanovi(k)), SRC, False
    Next k
End Sub

' Header DRAFT-a. VrstaVoca/SortaVoca su SNAPSHOT kulture -- upisuju se odmah,
' da panel ima cime da prefiluje formu otkupa. TipAmbalaze je primljena header
' cinjenica: ocekivanje "50 gajbi" mora da zna KOJIH 50 vec pri otvaranju.
'
' Kolicina / KolAmbalaze / Klasa / BrutoKg su polja stavke i u ciljnoj semi ih
' na headeru nema (S4.2).
Private Function BuildOtpremnicaHeaderRowData(ByVal otpID As String, _
                                              ByVal datum As Date, _
                                              ByVal stanicaID As String, _
                                              ByVal vozacID As String, _
                                              ByVal kulturaID As String, _
                                              ByVal tipAmb As String, _
                                              ByVal brojOtp As String) As Variant
    Const SRC As String = "BuildOtpremnicaHeaderRowData"

    Dim colCount As Long
    colCount = TabelaBrojKolona(TBL_OTPREMNICA)

    If colCount <= 0 Then
        Err.Raise vbObjectError + 1286, SRC, _
                  "Ne mogu da odredim broj kolona za tblOtpremnica."
    End If

    Dim rowData() As Variant
    ReDim rowData(0 To colCount - 1)

    SetRowValueByColumn rowData, TBL_OTPREMNICA, COL_OTP_ID, otpID, SRC
    SetRowValueByColumn rowData, TBL_OTPREMNICA, COL_OTP_DATUM, datum, SRC
    SetRowValueByColumn rowData, TBL_OTPREMNICA, COL_OTP_STANICA, stanicaID, SRC
    SetRowValueByColumn rowData, TBL_OTPREMNICA, COL_OTP_VOZAC, vozacID, SRC
    SetRowValueByColumn rowData, TBL_OTPREMNICA, COL_OTP_KULTURA, kulturaID, SRC
    SetRowValueByColumn rowData, TBL_OTPREMNICA, COL_OTP_TIP_AMB, tipAmb, SRC
    SetRowValueByColumn rowData, TBL_OTPREMNICA, COL_OTP_BROJ, brojOtp, SRC
    SetRowValueByColumn rowData, TBL_OTPREMNICA, COL_STORNIRANO, "", SRC

    SetRowValueByColumn rowData, TBL_OTPREMNICA, COL_OTP_VRSTA, _
        Trim$(NzToText(LookupValue(TBL_KULTURE, COL_KUL_ID, kulturaID, COL_KUL_VRSTA))), SRC
    SetRowValueByColumn rowData, TBL_OTPREMNICA, COL_OTP_SORTA, _
        Trim$(NzToText(LookupValue(TBL_KULTURE, COL_KUL_ID, kulturaID, COL_KUL_SORTA))), SRC

    ' COL_OTP_CENA se NE pise od S3a: predlog cene je po klasi, na stavci.
    ' Kolona ostaje u kanonu do S3e, kad odu i linijska polja zaglavlja.
    SetRowValueByColumn rowData, TBL_OTPREMNICA, COL_TRACE_IZDATO_STATUS, _
                        IZDATO_DRAFT, SRC

    BuildOtpremnicaHeaderRowData = rowData
End Function

' --- core: ocekivanje (stavke drafta) ---------------------------------------
Private Sub OtpUpisiOcekivano(ByVal otpremnicaID As String, _
                              ByVal ocekivano As Collection, ByVal src As String)
    RequireColumnIndex TBL_OTPREMNICA_STAVKE, COL_OPS_ID, src
    RequireColumnIndex TBL_OTPREMNICA_STAVKE, COL_OPS_OTPREMNICA_ID, src
    RequireColumnIndex TBL_OTPREMNICA_STAVKE, COL_OPS_RB, src
    RequireColumnIndex TBL_OTPREMNICA_STAVKE, COL_OPS_KLASA, src
    RequireColumnIndex TBL_OTPREMNICA_STAVKE, COL_OPS_KOLICINA, src
    RequireColumnIndex TBL_OTPREMNICA_STAVKE, COL_OPS_KOL_AMB, src

    If ocekivano.count = 0 Then
        Err.Raise vbObjectError + 1316, src, _
                  "Ocekivanje je prazno. Draft bez ijedne stavke nema sta da meri."
    End If

    Dim kol As Object, amb As Object, cena As Object
    Set kol = CreateObject("Scripting.Dictionary")
    Set amb = CreateObject("Scripting.Dictionary")
    Set cena = CreateObject("Scripting.Dictionary")

    Dim i As Long, s As Object, klasa As String
    Dim k As Double, a As Double, c As Double

    For i = 1 To ocekivano.count
        If Not IsObject(ocekivano(i)) Then
            Err.Raise vbObjectError + 1317, src, _
                      "Ocekivana stavka " & CStr(i) & " nije Dictionary."
        End If
        Set s = ocekivano(i)
        OtpOcekProveriKljuceve s, i, src

        klasa = Trim$(NzToText(OtpStavkaVrednost(s, "Klasa", i, src)))
        RequireValidKlasa klasa, src

        If kol.Exists(UCase$(klasa)) Then
            Err.Raise vbObjectError + 1318, src, _
                      "Dve ocekivane stavke iste klase: " & klasa
        End If

        k = OtpStavkaBroj(s, "Kolicina", i, src)
        If k <= 0 Then
            Err.Raise vbObjectError + 1319, src, _
                      "Ocekivana kolicina mora biti veca od nule. Klasa " & klasa & "."
        End If

        a = OtpStavkaBroj(s, "KolAmbalaze", i, src)
        If a < 0 Then
            Err.Raise vbObjectError + 1320, src, _
                      "Ocekivana ambalaza ne sme biti negativna. Klasa " & klasa & "."
        End If
        RequireCeoBroj a, "Ocekivana ambalaza, klasa " & klasa, src

        ' Predlog cene sme da izostane (jednopotezni put ga nema -- tamo otkupi
        ' vec postoje, pa nema sta da se prefiluje), ali ne sme da bude minus.
        c = OtpStavkaBrojOpcion(s, "PredlogCena", i, src)
        If c < 0 Then
            Err.Raise vbObjectError + 1335, src, _
                      "Predlog cene ne sme biti negativan. Klasa " & klasa & "."
        End If

        kol(UCase$(klasa)) = k
        amb(UCase$(klasa)) = a
        cena(UCase$(klasa)) = c
    Next i

    ' Tip ambalaze je obavezan tacno kad se ambalaza i ocekuje. Provera je OVDE
    ' jer se tek sada zna zbir -- i vazi za oba puta, i rucni draft i izvedeno
    ' ocekivanje jednopoteznog ulaza.
    Dim ukupnoAmb As Double
    Dim kljuc As Variant
    For Each kljuc In amb.Keys
        ukupnoAmb = ukupnoAmb + OtpBroj(amb, CStr(kljuc))
    Next kljuc

    If ukupnoAmb > 0 Then
        If Len(Trim$(NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, _
                                          otpremnicaID, COL_OTP_TIP_AMB)))) = 0 Then
            Err.Raise vbObjectError + 1331, src, _
                      "Tip ambalaze je obavezan kada se ambalaza ocekuje (" & _
                      Fmt2Zbr(ukupnoAmb) & "). Ocekivanje mora da kaze KOJIH gajbi."
        End If
    End If

    Dim rowData As Variant
    Dim stavkaID As String
    Dim rb As Long
    Dim kl As Variant

    For Each kl In KlaseUKanonskomRedu(kol)
        rb = rb + 1
        stavkaID = NewEntityID("OPS-")
        If stavkaID = "" Then
            Err.Raise vbObjectError + 1302, src, _
                      "NewEntityID nije vratio OtpremnicaStavkaID za klasu " & CStr(kl) & "."
        End If

        rowData = BuildOtpremnicaStavkaRowData(stavkaID, otpremnicaID, rb, CStr(kl), _
                                               OtpBroj(kol, CStr(kl)), _
                                               OtpBroj(amb, CStr(kl)), 0#, _
                                               OtpBroj(cena, CStr(kl)))

        If AppendRow(TBL_OTPREMNICA_STAVKE, rowData) <= 0 Then
            Err.Raise vbObjectError + 1303, src, _
                      "AppendRow nije upisao stavku klase " & CStr(kl) & "."
        End If
    Next kl
End Sub

Private Sub OtpObrisiOcekivano(ByVal otpremnicaID As String, ByVal src As String)
    Dim d As Variant
    d = GetTableData(TBL_OTPREMNICA_STAVKE)
    If Not IsArray(d) Then Exit Sub

    Dim cOtp As Long
    cOtp = RequireColumnIndex(TBL_OTPREMNICA_STAVKE, COL_OPS_OTPREMNICA_ID, src)

    ' Odozdo nagore: brisanje reda pomera indekse iznad njega.
    Dim i As Long
    For i = UBound(d, 1) To 1 Step -1
        If StrComp(Trim$(NzToText(d(i, cOtp))), otpremnicaID, vbTextCompare) = 0 Then
            RequireDeleteRow TBL_OTPREMNICA_STAVKE, i, src
        End If
    Next i
End Sub

' Ocekivanje IZVEDENO iz izvora -- samo za jednopotezni ulaz.
Private Sub OtpUpisiOcekivanjeIzIzvora(ByVal otpremnicaID As String)
    Const SRC As String = "OtpUpisiOcekivanjeIzIzvora"

    Dim pov As Object, povAmb As Object, povBruto As Object, povBrutoPun As Object
    OtpUcitajPovezano otpremnicaID, pov, povAmb, povBruto, povBrutoPun, SRC

    If pov.count = 0 Then
        Err.Raise vbObjectError + 1301, SRC, _
                  "Izvori nemaju nijednu stavku: " & otpremnicaID
    End If

    Dim c As Collection
    Set c = New Collection

    Dim kl As Variant, s As Object
    For Each kl In KlaseUKanonskomRedu(pov)
        Set s = CreateObject("Scripting.Dictionary")
        s.Add "Klasa", CStr(kl)
        s.Add "Kolicina", OtpBroj(pov, CStr(kl))
        s.Add "KolAmbalaze", OtpBroj(povAmb, CStr(kl))
        c.Add s
    Next kl

    OtpUpisiOcekivano otpremnicaID, c, SRC
End Sub

' --- core: clanstvo ---------------------------------------------------------

' Sve sto izvor mora da zadovolji. Zove se i iz Dodaj i iz Izdaj -- izdavanje NE
' veruje onome sto je Dodaj proverio, jer izmedju njih prolazi vreme.
Private Sub OtpRequireIzvorValjan(ByVal otpremnicaID As String, _
                                  ByVal otkupID As String, _
                                  ByVal src As String, _
                                  ByVal traziSlobodan As Boolean)
    If Len(otkupID) = 0 Then
        Err.Raise vbObjectError + 1287, src, "Prazan OtkupID."
    End If

    Dim rOtp As Long
    rOtp = OtpRedHeadera(otpremnicaID, src)
    RequireOtpDraft otpremnicaID, rOtp, src

    RequireTacnoJedan TBL_OTKUP, COL_OTK_ID, otkupID, "OtkupID", src

    If StrComp(Trim$(NzToText(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, _
                                          COL_STORNIRANO))), "Da", _
               vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 1288, src, "Otkup je storniran: " & otkupID
    End If

    Dim stavke As Collection
    Set stavke = FindRows(TBL_OTKUP_STAVKE, COL_OKS_OTKUP_ID, otkupID)
    If stavke.count = 0 Then
        Err.Raise vbObjectError + 1289, src, _
                  "Otkup nema stavke: " & otkupID & _
                  ". Kanonski writer cita kolicine iz tblOtkupStavke, pa otkup " & _
                  "po starom modelu ne moze da udje u otpremnicu."
    End If

    ' Otpremnica je isporuka SA JEDNOG otkupnog mesta -- header nosi jedan
    ' StanicaID, pa dve stanice ne mogu ni da se predstave (S4.2a).
    RequireIstoPolje Trim$(NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, _
                                                otpremnicaID, COL_OTP_STANICA))), _
                     Trim$(NzToText(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, _
                                                COL_OTK_STANICA))), _
                     "StanicaID", otkupID, src

    ' Kultura je relaciona istina; vrsta/sorta se poklapaju posledicno.
    RequireIstoPolje Trim$(NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, _
                                                otpremnicaID, COL_OTP_KULTURA))), _
                     Trim$(NzToText(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, _
                                                COL_OTK_KULTURA))), _
                     "KulturaID", otkupID, src

    ' Izvor mora biti IZDAT dokument, ne bilo koji red koji slucajno ima stavke.
    ' Danas svaki otkup iz CreateOtkup_TX jeste IZDATO, pa ovo nije ziv bug --
    ' ali kanonska veza treba da kaze sta trazi, a ne da se oslanja na to sto
    ' drugi pisac trenutno ne pravi drugacije redove.
    Dim izdato As String
    izdato = UCase$(Trim$(NzToText(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, _
                                               COL_TRACE_IZDATO_STATUS))))
    If izdato <> UCase$(IZDATO_IZDATO) Then
        Err.Raise vbObjectError + 1330, src, _
                  "Otkup nije izdat nego '" & izdato & "': " & otkupID & _
                  ". Otpremnica se sastavlja od IZDATIH otkupnih listova."
    End If

    ' Tip ambalaze se poredi sa HEADEROM, i to SAMO ako izvor stvarno nosi gajbe.
    '
    ' Otkup sme da ima TipAmbalaze zbog KolAmbIzdata -- gajbi koje su OTISLE
    ' kooperantu -- a da njegove stavke ne nose nijednu gajbu u otpremnicu. Takav
    ' izvor ne odredjuje transportnu ambalazu, pa poredjenje golih header stringova
    ' svih otkupa nije precizno.
    If OtpGajbeIzvora(otkupID, src) > 0 Then
        RequireIstoPolje Trim$(NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, _
                                                    otpremnicaID, COL_OTP_TIP_AMB))), _
                         Trim$(NzToText(LookupValue(TBL_OTKUP, COL_OTK_ID, _
                                                    otkupID, COL_OTK_TIP_AMB))), _
                         "TipAmbalaze", otkupID, src
    End If

    ' Kanonsko clanstvo je jedini izvor. Otkup.OtpremnicaID se NE gleda: to je
    ' stari model, koji skela ne dira.
    Dim clanstvo As Object
    Set clanstvo = AktivnoOtpClanstvoPoKanonu(src)

    If clanstvo.Exists(UCase$(otkupID)) Then
        If traziSlobodan Then
            Err.Raise vbObjectError + 1291, src, _
                      "Otkup je vec u sastavu aktivne otpremnice: " & otkupID & _
                      " -> " & CStr(clanstvo(UCase$(otkupID)))
        End If
        ' Pri izdavanju izvor SME da bude vezan -- ali bas za OVU otpremnicu.
        If StrComp(CStr(clanstvo(UCase$(otkupID))), otpremnicaID, vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + 1321, src, _
                      "Otkup " & otkupID & " je u sastavu druge aktivne otpremnice: " & _
                      CStr(clanstvo(UCase$(otkupID)))
        End If
    ElseIf Not traziSlobodan Then
        Err.Raise vbObjectError + 1322, src, _
                  "Otkup " & otkupID & " vise nije u kanonskom clanstvu ove otpremnice."
    End If
End Sub

Private Sub OtpUpisiClanstvo(ByVal otpremnicaID As String, ByVal otkupID As String, _
                             ByVal src As String)
    RequireColumnIndex TBL_OTPREMNICA_IZVORI, COL_OPI_ID, src
    RequireColumnIndex TBL_OTPREMNICA_IZVORI, COL_OPI_OTPREMNICA_ID, src
    RequireColumnIndex TBL_OTPREMNICA_IZVORI, COL_OPI_OTKUP_ID, src

    Dim clanovi As Collection
    Set clanovi = OtpClanovi(otpremnicaID, SRC)

    Dim k As Long
    For k = 1 To clanovi.count
        If StrComp(CStr(clanovi(k)), otkupID, vbTextCompare) = 0 Then
            Err.Raise vbObjectError + 1290, src, _
                      "Otkup je vec u sastavu ove otpremnice: " & otkupID
        End If
    Next k

    Dim izvorID As String
    izvorID = NewEntityID("OPI-")

    If izvorID = "" Then
        Err.Raise vbObjectError + 1292, src, _
                  "NewEntityID nije vratio OtpremnicaIzvorID za " & otkupID & "."
    End If

    Dim rowData As Variant
    rowData = BuildOtpremnicaIzvorRowData(izvorID, otpremnicaID, otkupID)

    If AppendRow(TBL_OTPREMNICA_IZVORI, rowData) <= 0 Then
        Err.Raise vbObjectError + 1293, src, _
                  "AppendRow nije upisao clanstvo za " & otkupID & "."
    End If
End Sub

Private Function BuildOtpremnicaIzvorRowData(ByVal izvorID As String, _
                                             ByVal otpremnicaID As String, _
                                             ByVal otkupID As String) As Variant
    Const SRC As String = "BuildOtpremnicaIzvorRowData"

    Dim colCount As Long
    colCount = TabelaBrojKolona(TBL_OTPREMNICA_IZVORI)

    If colCount <= 0 Then
        Err.Raise vbObjectError + 1294, SRC, _
                  "Ne mogu da odredim broj kolona za tblOtpremnicaIzvori."
    End If

    Dim rowData() As Variant
    ReDim rowData(0 To colCount - 1)

    SetRowValueByColumn rowData, TBL_OTPREMNICA_IZVORI, COL_OPI_ID, izvorID, SRC
    SetRowValueByColumn rowData, TBL_OTPREMNICA_IZVORI, COL_OPI_OTPREMNICA_ID, _
                        otpremnicaID, SRC
    SetRowValueByColumn rowData, TBL_OTPREMNICA_IZVORI, COL_OPI_OTKUP_ID, otkupID, SRC

    BuildOtpremnicaIzvorRowData = rowData
End Function

' Uklanjanje je FIZICKO brisanje reda, i samo dok je otpremnica DRAFT.
'
' Izvor uklonjen pre izdavanja nikad nije bio deo dokumenta -- tombstone bi
' znacio da svaki citalac sastava filtrira redove koji nikad nisu vazili
' (v. modDataAccess.DeleteRow).
Private Sub OtpUkloniIzvor(ByVal otpremnicaID As String, ByVal otkupID As String)
    Const SRC As String = "OtpUkloniIzvor"

    Dim rOtp As Long
    rOtp = OtpRedHeadera(otpremnicaID, SRC)
    RequireOtpDraft otpremnicaID, rOtp, SRC

    Dim izv As Variant
    izv = GetTableData(TBL_OTPREMNICA_IZVORI)
    If Not IsArray(izv) Then
        Err.Raise vbObjectError + 1295, SRC, _
                  "Otkup nije u sastavu ove otpremnice: " & otkupID
    End If

    Dim cOtp As Long, cOtk As Long
    cOtp = RequireColumnIndex(TBL_OTPREMNICA_IZVORI, COL_OPI_OTPREMNICA_ID, SRC)
    cOtk = RequireColumnIndex(TBL_OTPREMNICA_IZVORI, COL_OPI_OTKUP_ID, SRC)

    Dim i As Long, nadjen As Long, koliko As Long
    For i = 1 To UBound(izv, 1)
        If StrComp(Trim$(NzToText(izv(i, cOtp))), otpremnicaID, vbTextCompare) = 0 Then
            If StrComp(Trim$(NzToText(izv(i, cOtk))), otkupID, vbTextCompare) = 0 Then
                nadjen = i
                koliko = koliko + 1
            End If
        End If
    Next i

    If koliko = 0 Then
        Err.Raise vbObjectError + 1296, SRC, _
                  "Otkup nije u sastavu ove otpremnice: " & otkupID
    End If
    If koliko > 1 Then
        Err.Raise vbObjectError + 1297, SRC, _
                  "Isti par (otpremnica, otkup) postoji " & CStr(koliko) & _
                  " puta: " & otpremnicaID & " / " & otkupID
    End If

    RequireDeleteRow TBL_OTPREMNICA_IZVORI, nadjen, SRC
End Sub

' --- core: izdavanje --------------------------------------------------------
Private Sub OtpIzdaj(ByVal otpremnicaID As String)
    Const SRC As String = "OtpIzdaj"

    Dim rOtp As Long
    rOtp = OtpRedHeadera(otpremnicaID, SRC)
    RequireOtpDraft otpremnicaID, rOtp, SRC

    Dim clanovi As Collection
    Set clanovi = OtpClanovi(otpremnicaID, SRC)

    If clanovi.count = 0 Then
        Err.Raise vbObjectError + 1298, SRC, _
                  "Otpremnica nema nijedan izvor: " & otpremnicaID & _
                  ". Otpremnica bez otkupa nije isporuka."
    End If

    ' REVALIDACIJA. Izmedju Dodaj i Izdaj prolazi vreme -- u panelu i po nekoliko
    ' sati -- pa se izvor u medjuvremenu moze stornirati ili ispraviti. Provera i
    ' upotreba moraju biti u istom trenutku, inace je ovo TOCTOU: DRAFT -> dodaj
    ' OTK1 -> storno OTK1 -> Izdaj je izdavao dokument iz storniranog izvora.
    Dim k As Long
    For k = 1 To clanovi.count
        OtpRequireIzvorValjan otpremnicaID, CStr(clanovi(k)), SRC, False
    Next k

    ' --- ocekivano = povezano, po klasi --------------------------------------
    Dim ocek As Object, ocekAmb As Object
    Dim pov As Object, povAmb As Object, povBruto As Object, povBrutoPun As Object
    Set ocek = CreateObject("Scripting.Dictionary")
    Set ocekAmb = CreateObject("Scripting.Dictionary")

    OtpUcitajOcekivano otpremnicaID, ocek, ocekAmb, SRC
    OtpUcitajPovezano otpremnicaID, pov, povAmb, povBruto, povBrutoPun, SRC

    OtpRequireJednakost otpremnicaID, ocek, pov, "Kolicina", SRC
    OtpRequireJednakost otpremnicaID, ocekAmb, povAmb, "KolAmbalaze", SRC

    ' --- zamrzavanje: stavke prestaju da budu ocekivanje ---------------------
    '
    ' Brojevi se ne prepisuju -- jednakost je upravo dokazana. Dopisuje se samo
    ' BrutoKg, koji operater ne prijavljuje nego dolazi iz izvora.
    Dim d As Variant
    d = GetTableData(TBL_OTPREMNICA_STAVKE)
    If Not IsArray(d) Then
        Err.Raise vbObjectError + 1323, SRC, _
                  "Otpremnica nema stavke: " & otpremnicaID
    End If

    Dim cOtp As Long, cKlasa As Long
    cOtp = RequireColumnIndex(TBL_OTPREMNICA_STAVKE, COL_OPS_OTPREMNICA_ID, SRC)
    cKlasa = RequireColumnIndex(TBL_OTPREMNICA_STAVKE, COL_OPS_KLASA, SRC)

    Dim i As Long, klasa As String
    For i = 1 To UBound(d, 1)
        If StrComp(Trim$(NzToText(d(i, cOtp))), otpremnicaID, vbTextCompare) = 0 Then
            klasa = UCase$(Trim$(NzToText(d(i, cKlasa))))
            ' BrutoKg SAMO ako ga nosi svaka izvorna stavka te klase. Parcijalan
            ' zbir bi dao fizicki nemoguc red (bruto manji od neta) -- v. S4.2.
            If OtpBroj(povBrutoPun, klasa) > 0 Then
                RequireUpdateCell TBL_OTPREMNICA_STAVKE, i, COL_OPS_BRUTO, _
                                  OtpBroj(povBruto, klasa), SRC
            End If
        End If
    Next i

    OtpKnjiziAmbalazu otpremnicaID, rOtp, ocekAmb, SRC

    RequireUpdateCell TBL_OTPREMNICA, rOtp, COL_TRACE_IZDATO_STATUS, IZDATO_IZDATO, SRC
End Sub

' --- core: ispravka izdate -------------------------------------------------
Private Function OtpIspravi(ByVal staraID As String) As String
    Const SRC As String = "OtpIspravi"

    Dim rStara As Long
    rStara = OtpRedHeadera(staraID, SRC)

    ' Samo IZDATA. Nacrt jos nije dokument: on se menja, a ispravka bi ga
    ' stornirala i potrosila jos jedan broj niza ni za sta.
    If Not IzdatoStatusJeIzdato(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, staraID, _
                                            COL_TRACE_IZDATO_STATUS)) Then
        Err.Raise vbObjectError + 1338, SRC, _
                  "Ispravlja se samo IZDATA otpremnica: " & staraID & _
                  ". Nacrt se menja (UpdateOtpremnicaDraft_TX), bez storna i novog broja."
    End If

    ' Sastav i ocekivanje se citaju PRE storna, kroz iste stroge kanonske
    ' citace koje koristi izdavanje: otpremnicu koju izdavanje ne bi primilo
    ' ispravka ne sme da prepise u nov nacrt.
    Dim clanovi As Collection
    Set clanovi = OtpClanovi(staraID, SRC)
    If clanovi.count = 0 Then
        Err.Raise vbObjectError + 1339, SRC, _
                  "Izdata otpremnica nema nijedan izvor: " & staraID & _
                  ". Bez izvora nema sta da se prenese na novu."
    End If

    Dim ocekivano As Collection
    Set ocekivano = OtpOcekivanjeKolekcija(staraID, SRC)

    Dim datum As Date
    Dim stanicaID As String, vozacID As String, kulturaID As String, tipAmb As String
    datum = CDate(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, staraID, COL_OTP_DATUM))
    stanicaID = Trim$(NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, staraID, COL_OTP_STANICA)))
    vozacID = Trim$(NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, staraID, COL_OTP_VOZAC)))
    kulturaID = Trim$(NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, staraID, COL_OTP_KULTURA)))
    tipAmb = Trim$(NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, staraID, COL_OTP_TIP_AMB)))

    ' Storno stare IDE PRE nego sto nova primi izvore: izvor sme da bude u
    ' tacno jednoj aktivnoj otpremnici (OtpRequireIzvorValjan), pa bi obrnut
    ' redosled sam sebe odbio. Jezgro nosi i kapiju izvora aktivne zbirne.
    If Not modStorno.StornoOtpremnica(staraID) Then
        Err.Raise vbObjectError + 1340, SRC, _
                  "Storno stare otpremnice nije uspeo: " & staraID
    End If

    Dim h As Object
    Set h = CreateObject("Scripting.Dictionary")
    h("Datum") = datum
    h("StanicaID") = stanicaID
    h("VozacID") = vozacID
    h("KulturaID") = kulturaID
    h("TipAmbalaze") = tipAmb

    Dim noviBroj As String
    noviBroj = modBrojevi.GenerateBrojOtpremnice(stanicaID, datum)
    If Len(noviBroj) = 0 Then
        Err.Raise vbObjectError + 1341, SRC, _
                  "Nov broj otpremnice nije generisan (stanica " & stanicaID & ")."
    End If
    h("BrojOtpremnice") = noviBroj

    Dim novaID As String
    novaID = OtpNapraviDraft(h, ocekivano)

    Dim i As Long
    For i = 1 To clanovi.count
        OtpRequireIzvorValjan novaID, CStr(clanovi(i)), SRC, True
        OtpUpisiClanstvo novaID, CStr(clanovi(i)), SRC
    Next i

    ' Trag ide po IDENTITETU (kao otkup od S1e), ne po broju -- broj je
    ' jedinstven tek po (stanica, dan), pa bi veza po broju umela da pokaze na
    ' tudji dokument.
    RequireUpdateCell TBL_OTPREMNICA, rStara, COL_TRACE_ZAMENJEN_SA_ID, novaID, SRC
    RequireUpdateCell TBL_OTPREMNICA, OtpRedHeadera(novaID, SRC), _
                      COL_TRACE_ISPRAVKA_OD_ID, staraID, SRC

    OtpIspravi = novaID
End Function

' Ocekivanje jednog dokumenta kao kolekcija stavki za OtpNapraviDraft.
'
' Cita se kroz StavkeOtpremniceRedovi -- strog citac koji drzi ugovor pisca
' (jedna stavka po klasi, kolicina > 0, gajbe ceo broj), pa se prepisuje samo
' ocekivanje koje bi i sam pisac prihvatio.
' Kolone: 1 OtpremnicaID, 3 Klasa, 4 Kolicina, 5 PredlogCena, 6 KolAmbalaze.
Private Function OtpOcekivanjeKolekcija(ByVal otpremnicaID As String, _
                                        ByVal src As String) As Collection
    Dim c As Collection
    Set c = New Collection
    Set OtpOcekivanjeKolekcija = c

    Dim redovi As Variant
    redovi = StavkeOtpremniceRedovi()

    Dim i As Long, s As Object
    If IsArray(redovi) Then
        For i = 1 To UBound(redovi, 1)
            If StrComp(CStr(redovi(i, 1)), otpremnicaID, vbTextCompare) = 0 Then
                Set s = CreateObject("Scripting.Dictionary")
                s("Klasa") = CStr(redovi(i, 3))
                s("Kolicina") = CDbl(redovi(i, 4))
                s("PredlogCena") = CDbl(redovi(i, 5))
                s("KolAmbalaze") = CDbl(redovi(i, 6))
                c.Add s
            End If
        Next i
    End If

    If c.count = 0 Then
        Err.Raise vbObjectError + 1342, src, _
                  "Otpremnica nema ocekivanje: " & otpremnicaID
    End If
End Function

' AMBALAZA SE KNJIZI PRI IZDAVANJU, ne pri otvaranju nacrta (odluka S14.8 t. 1).
'
' Gajbe fizicki odlaze sa stanice kad ih vozac preuzme, a to je izdavanje. Nacrt
' je najava: menja se, ispravlja i sme da ostane neizdat -- da je ambalaza isla
' sa njim, svaka izmena ocekivanja bi trazila i storniranje knjizenja, a
' napusten nacrt bi trajno umanjio stanje gajbi na stanici.
'
' Zato je i kolicina ZBIR STAVKI izdate otpremnice, a ne broj sa zaglavlja: posle
' S1/S3a zaglavlje kolicinu ambalaze vise i ne nosi.
Private Sub OtpKnjiziAmbalazu(ByVal otpremnicaID As String, ByVal rOtp As Long, _
                              ByVal ocekAmb As Object, ByVal src As String)
    Dim ukupno As Double
    Dim kljuc As Variant

    If Not ocekAmb Is Nothing Then
        For Each kljuc In ocekAmb.Keys
            ukupno = ukupno + OtpBroj(ocekAmb, CStr(kljuc))
        Next kljuc
    End If

    If ukupno <= 0 Then Exit Sub

    Dim d As Variant
    d = GetTableData(TBL_OTPREMNICA)
    If Not IsArray(d) Then
        Err.Raise vbObjectError + 1336, src, _
                  "Nema zaglavlja otpremnice: " & otpremnicaID
    End If

    Dim datum As Date
    Dim tipAmb As String, stanicaID As String, vozacID As String

    datum = CDate(d(rOtp, RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_DATUM, src)))
    tipAmb = Trim$(NzToText(d(rOtp, RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_TIP_AMB, src))))
    stanicaID = Trim$(NzToText(d(rOtp, RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_STANICA, src))))
    vozacID = Trim$(NzToText(d(rOtp, RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VOZAC, src))))

    ' Tip ambalaze je vec bio uslov da ocekivanje uopste nastane
    ' (OtpUpisiOcekivano), pa je ovde prazan samo ako je neko posle toga
    ' ispraznio zaglavlje. Knjizenje bez tipa nije "manje precizno" nego
    ' neupotrebljivo -- stanje gajbi se vodi PO TIPU.
    If Len(tipAmb) = 0 Then
        Err.Raise vbObjectError + 1337, src, _
                  "Otpremnica " & otpremnicaID & " nema tip ambalaze, a izdaje " & _
                  Fmt2Zbr(ukupno) & " gajbi."
    End If

    TrackAmbalaza datum, tipAmb, CLng(ukupno), "Izlaz", stanicaID, "Stanica", _
                  vozacID, otpremnicaID, DOK_TIP_OTPREMNICA
End Sub

Private Sub OtpRequireJednakost(ByVal otpremnicaID As String, _
                                ByVal ocek As Object, ByVal pov As Object, _
                                ByVal polje As String, ByVal src As String)
    Dim sve As Object
    Set sve = CreateObject("Scripting.Dictionary")

    Dim k As Variant
    For Each k In ocek.Keys
        sve(CStr(k)) = True
    Next k
    For Each k In pov.Keys
        sve(CStr(k)) = True
    Next k

    Dim o As Double, p As Double
    For Each k In sve.Keys
        o = OtpBroj(ocek, CStr(k))
        p = OtpBroj(pov, CStr(k))
        If Abs(o - p) > 0.0001 Then
            Err.Raise vbObjectError + 1324, src, _
                      "Otpremnica " & otpremnicaID & " nije spremna za izdavanje: " & _
                      polje & ", klasa " & CStr(k) & " -- ocekivano " & Fmt2Zbr(o) & _
                      ", povezano " & Fmt2Zbr(p) & ", preostalo " & Fmt2Zbr(o - p) & "."
        End If
    Next k
End Sub

Private Function BuildOtpremnicaStavkaRowData(ByVal stavkaID As String, _
                                              ByVal otpremnicaID As String, _
                                              ByVal redniBroj As Long, _
                                              ByVal klasa As String, _
                                              ByVal kolicina As Double, _
                                              ByVal kolAmb As Double, _
                                              ByVal bruto As Double, _
                                              ByVal predlogCena As Double) As Variant
    Const SRC As String = "BuildOtpremnicaStavkaRowData"

    Dim colCount As Long
    colCount = TabelaBrojKolona(TBL_OTPREMNICA_STAVKE)

    If colCount <= 0 Then
        Err.Raise vbObjectError + 1304, SRC, _
                  "Ne mogu da odredim broj kolona za tblOtpremnicaStavke."
    End If

    Dim rowData() As Variant
    ReDim rowData(0 To colCount - 1)

    SetRowValueByColumn rowData, TBL_OTPREMNICA_STAVKE, COL_OPS_ID, stavkaID, SRC
    SetRowValueByColumn rowData, TBL_OTPREMNICA_STAVKE, COL_OPS_OTPREMNICA_ID, _
                        otpremnicaID, SRC
    SetRowValueByColumn rowData, TBL_OTPREMNICA_STAVKE, COL_OPS_RB, redniBroj, SRC
    SetRowValueByColumn rowData, TBL_OTPREMNICA_STAVKE, COL_OPS_KLASA, klasa, SRC
    SetRowValueByColumn rowData, TBL_OTPREMNICA_STAVKE, COL_OPS_KOLICINA, kolicina, SRC
    SetRowValueByColumn rowData, TBL_OTPREMNICA_STAVKE, COL_OPS_KOL_AMB, kolAmb, SRC

    ' BrutoKg se na draftu NE upisuje -- operater ga ne prijavljuje. Dopisuje ga
    ' izdavanje, i samo kad ga nosi svaki izvor te klase.
    If bruto > 0 Then
        SetRowValueByColumn rowData, TBL_OTPREMNICA_STAVKE, COL_OPS_BRUTO, bruto, SRC
    End If

    ' Predlog cene: prazno kad ga nema, da prazna celija ostane prazna a ne nula.
    ' Nula bi u prefillu otkupa izgledala kao "cena je dogovorena i iznosi 0".
    If predlogCena > 0 Then
        SetRowValueByColumn rowData, TBL_OTPREMNICA_STAVKE, COL_OPS_PREDLOG_CENA, _
                            predlogCena, SRC
    End If

    BuildOtpremnicaStavkaRowData = rowData
End Function

' --- citanje ----------------------------------------------------------------
'
' OBA UCITAVANJA IDU KROZ KANONSKE STROGE CITAOCE (review #363, P1), ne kroz
' sirove tabele. Ranije su ovde tblOtpremnicaStavke i tblOtkupStavke citani
' direktno, a nebrojcana vrednost je postajala nula (OtpDbl). Dve stavke iste
' klase su se SABIRALE: nacrt sa 400 + 100 kg klase I i izvorom od 500 kg
' prolazio je jednakost i postajao IZDATO -- dokument koji kanonski citalac
' odbija kao korumpiran. Sada read-model (GetOtpremnicaProgress) i izdavanje
' drze ISTI ugovor stavki kao mreza, stampa i izvestaji: korupcija pada po
' imenu, a nijedno pravilo nije napisano drugi put.
Private Sub OtpUcitajOcekivano(ByVal otpremnicaID As String, _
                               ByRef ocek As Object, ByRef ocekAmb As Object, _
                               ByVal src As String)
    If ocek Is Nothing Then Set ocek = CreateObject("Scripting.Dictionary")
    If ocekAmb Is Nothing Then Set ocekAmb = CreateObject("Scripting.Dictionary")

    ' 1 OtpremnicaID | 3 Klasa | 4 Kolicina | 6 KolAmbalaze -- klasa je vec
    ' proverena (jedna stavka po klasi), kolicina > 0, gajbe ceo broj >= 0.
    Dim ocekRedovi As Variant
    ocekRedovi = StavkeOtpremniceRedovi()
    If Not IsArray(ocekRedovi) Then Exit Sub

    Dim i As Long, klasa As String
    For i = 1 To UBound(ocekRedovi, 1)
        If StrComp(CStr(ocekRedovi(i, 1)), otpremnicaID, vbTextCompare) = 0 Then
            klasa = UCase$(CStr(ocekRedovi(i, 3)))
            ocek(klasa) = OtpBroj(ocek, klasa) + CDbl(ocekRedovi(i, 4))
            ocekAmb(klasa) = OtpBroj(ocekAmb, klasa) + CDbl(ocekRedovi(i, 6))
        End If
    Next i
End Sub

' Povezano: zbir otkupnih stavki svih clanova, po klasi.
'
' brutoPun je 1 samo ako SVAKA izvorna stavka te klase nosi bruto -- cim je
' jedna prazna, ostaje 0 i bruto se ne upisuje (S4.2).
Private Sub OtpUcitajPovezano(ByVal otpremnicaID As String, _
                              ByRef pov As Object, ByRef povAmb As Object, _
                              ByRef povBruto As Object, ByRef povBrutoPun As Object, _
                              ByVal src As String)
    Set pov = CreateObject("Scripting.Dictionary")
    Set povAmb = CreateObject("Scripting.Dictionary")
    Set povBruto = CreateObject("Scripting.Dictionary")
    Set povBrutoPun = CreateObject("Scripting.Dictionary")

    Dim clanovi As Collection
    Set clanovi = OtpClanovi(otpremnicaID, src)
    If clanovi.count = 0 Then Exit Sub

    Dim clanSet As Object
    Set clanSet = CreateObject("Scripting.Dictionary")

    Dim k As Long
    For k = 1 To clanovi.count
        clanSet(UCase$(CStr(clanovi(k)))) = True
    Next k

    ' Kanonski strog citac otkupa (modOtkup.StavkeOtkupaRedovi): klasa I/II,
    ' kolicina i cena > 0, gajbe ceo broj >= 0, svako zaglavlje sa stavkom.
    ' 1 OtkupID | 3 Klasa | 4 Kolicina | 6 KolAmbalaze | 8 BrutoKg ("" = neto)
    Dim sve As Variant
    sve = modOtkup.StavkeOtkupaRedovi()
    If Not IsArray(sve) Then Exit Sub

    Dim brojStavki As Object, brojSaBrutom As Object
    Set brojStavki = CreateObject("Scripting.Dictionary")
    Set brojSaBrutom = CreateObject("Scripting.Dictionary")

    Dim i As Long, klasa As String, bruto As Double
    For i = 1 To UBound(sve, 1)
        If clanSet.Exists(UCase$(CStr(sve(i, 1)))) Then
            klasa = UCase$(CStr(sve(i, 3)))

            pov(klasa) = OtpBroj(pov, klasa) + CDbl(sve(i, 4))
            povAmb(klasa) = OtpBroj(povAmb, klasa) + CDbl(sve(i, 6))

            bruto = 0
            If IsNumeric(sve(i, 8)) Then bruto = CDbl(sve(i, 8))
            povBruto(klasa) = OtpBroj(povBruto, klasa) + bruto
            brojStavki(klasa) = OtpBroj(brojStavki, klasa) + 1
            If bruto > 0 Then brojSaBrutom(klasa) = OtpBroj(brojSaBrutom, klasa) + 1
        End If
    Next i

    Dim kl As Variant
    For Each kl In pov.Keys
        If OtpBroj(brojSaBrutom, CStr(kl)) = OtpBroj(brojStavki, CStr(kl)) Then
            povBrutoPun(CStr(kl)) = 1
        End If
    Next kl
End Sub

' Sastav otpremnice po KANONU (tblOtpremnicaIzvori), redom upisa.
'
' STRIKT LOADER, i to je jedini put do clanstva JEDNOG dokumenta -- koriste ga i
' read-model (GetOtpremnicaProgress) i sva tri pisca. Isti ugovor nad CELOM
' tabelom drzi AktivnoOtpClanstvoPoKanonu, koju ova funkcija zove na kraju (i
' koju citaju bulk citaoci); pravilo je jedno, mesta izvrsavanja dva.
' Razlog: citalac koji tiho normalizuje
' korupciju (dupli par -> jedan clan, veza na nepostojeci otkup -> manji zbir)
' pokazuje operateru brojeve koji izgledaju ispravno, a finalizacija istu tu
' korupciju prijavi tek sat kasnije. Ugovor mora biti isti na oba mesta.
Private Function OtpClanovi(ByVal otpremnicaID As String, _
                            ByVal src As String) As Collection
    Dim c As Collection
    Set c = New Collection
    Set OtpClanovi = c

    ' Roditelj mora da postoji tacno jednom -- clanstvo bez headera nije sastav.
    RequireTacnoJedan TBL_OTPREMNICA, COL_OTP_ID, otpremnicaID, "OtpremnicaID", src

    Dim izv As Variant
    izv = GetTableData(TBL_OTPREMNICA_IZVORI)
    If Not IsArray(izv) Then Exit Function

    Dim cOtp As Long, cOtk As Long
    cOtp = RequireColumnIndex(TBL_OTPREMNICA_IZVORI, COL_OPI_OTPREMNICA_ID, src)
    cOtk = RequireColumnIndex(TBL_OTPREMNICA_IZVORI, COL_OPI_OTKUP_ID, src)

    Dim vidjeni As Object
    Set vidjeni = CreateObject("Scripting.Dictionary")

    Dim i As Long, otkupID As String
    For i = 1 To UBound(izv, 1)
        If StrComp(Trim$(NzToText(izv(i, cOtp))), otpremnicaID, vbTextCompare) = 0 Then
            otkupID = Trim$(NzToText(izv(i, cOtk)))

            If Len(otkupID) = 0 Then
                Err.Raise vbObjectError + 1332, src, _
                          "Clanstvo bez OtkupID-a, otpremnica " & otpremnicaID & "."
            End If

            If vidjeni.Exists(UCase$(otkupID)) Then
                Err.Raise vbObjectError + 1333, src, _
                          "Kanonsko clanstvo je nekonzistentno: par (" & _
                          otpremnicaID & ", " & otkupID & ") postoji vise puta."
            End If
            vidjeni.Add UCase$(otkupID), True

            ' Dete mora da postoji. Bez ovoga bi veza na nepostojeci otkup dala
            ' samo MANJI zbir -- tisi ishod od pada, i zato gori.
            RequireTacnoJedan TBL_OTKUP, COL_OTK_ID, otkupID, "OtkupID (clanstvo)", src

            c.Add otkupID
        End If
    Next i

    ' Globalno: isti otkup ne sme da bude u dve aktivne otpremnice.
    AktivnoOtpClanstvoPoKanonu src
End Function

' Koliko gajbi izvor stvarno nosi u otpremnicu (zbir po njegovim stavkama).
' KolAmbIzdata NIJE tu -- to su gajbe koje su otisle kooperantu.
Private Function OtpGajbeIzvora(ByVal otkupID As String, ByVal src As String) As Double
    Dim d As Variant
    d = GetTableData(TBL_OTKUP_STAVKE)
    If Not IsArray(d) Then Exit Function

    Dim cOtk As Long, cAmb As Long
    cOtk = RequireColumnIndex(TBL_OTKUP_STAVKE, COL_OKS_OTKUP_ID, src)
    cAmb = RequireColumnIndex(TBL_OTKUP_STAVKE, COL_OKS_KOL_AMB, src)

    Dim i As Long
    For i = 1 To UBound(d, 1)
        If StrComp(Trim$(NzToText(d(i, cOtk))), otkupID, vbTextCompare) = 0 Then
            OtpGajbeIzvora = OtpGajbeIzvora + OtpDbl(d(i, cAmb))
        End If
    Next i
End Function

' ZAMENA IZVORA U NACRTU -- DVA KORAKA, oba UNUTAR TUDJE transakcije (S3d-2).
'
' Zove ih modOtkup.IspravkaOtkupa_TX, i to razdvojeno, jer izmedju njih mora da
' stane storno starog bloka:
'
'     IzvadiIzvorIzNacrta(nacrt, stari)   <- PRE storna
'     StornoOtkup(stari)
'     CreateOtkup(novi)
'     UvediIzvorUNacrt(nacrt, novi)       <- POSLE upisa
'
' ZASTO BAS TIM REDOM. Jezgro modStorno.StornoOtkup ODBIJA storno bloka koji je
' u sastavu aktivne otpremnice -- i to je tacno pravilo, ne prepreka koju treba
' zaobici. Zato stari izvor prvo IZADJE iz nacrta; kapija tada prolazi jer je
' istina da blok vise nije ni u jednom sastavu. Alternativa ("dodaj izuzetak u
' StornoOtkup") bi otvorila zadnja vrata svakom pozivaocu.
'
' Storno ide PRE upisa naslednika i to se ne menja: pisac ispravke tako
' oslobadja novac i avanse koje naslednik posle preuzima.
'
' NE ZOVI SPOLJA BEZ TRANSAKCIJE. Pozivalac drzi snapshot tblOtpremnicaIzvori;
' bez njega bi pad izmedju dva koraka ostavio nacrt BEZ ijednog izvora -- a
' upravo to je jedini prozor koji ova podela pravi.
'
' IZDATA SE NE DIRA (RequireOtpDraft): njen sastav je istorijska cinjenica, a
' izmena bi bila tiha promena izdatog papira (A13). Za nju je put storno +
' reizdavanje, sto radi ispravka otpremnice (S3c).
Public Sub IzvadiIzvorIzNacrta(ByVal otpremnicaID As String, _
                               ByVal otkupID As String)
    Const SRC As String = "IzvadiIzvorIzNacrta"

    Dim rOtp As Long
    rOtp = OtpRedHeadera(otpremnicaID, SRC)
    RequireOtpDraft otpremnicaID, rOtp, SRC

    OtpUkloniIzvor otpremnicaID, Trim$(otkupID)
End Sub

' Naslednik ulazi u nacrt tek kad postoji, i prolazi ISTE kapije kao svaki izvor
' (aktivan, sa stavkama, iste stanice i kulture, slobodan). Ispravka ne sme da
' bude zadnja vrata: blok koji nacrt ne bi primio kroz "Vezi" ne prima se ni
' ovuda.
Public Sub UvediIzvorUNacrt(ByVal otpremnicaID As String, _
                            ByVal otkupID As String)
    Const SRC As String = "UvediIzvorUNacrt"

    Dim rOtp As Long
    rOtp = OtpRedHeadera(otpremnicaID, SRC)
    RequireOtpDraft otpremnicaID, rOtp, SRC

    OtpRequireIzvorValjan otpremnicaID, Trim$(otkupID), SRC, True
    OtpUpisiClanstvo otpremnicaID, Trim$(otkupID), SRC
End Sub

' Denormalizovan BrojOtpremnice na bloku -- OSTAJE dok traje kaskada zbirne.
'
' S3e-1 je ovo prvo obrisalo kao mrtvo, pa vratilo: `vba_check` je nasao zivog
' pozivaoca (modStornoFlow.FreeOtkupBloksInline, kaskada PONISTENJA zbirne),
' koga je moj grep propustio jer poziv nosi komentar na kraju reda. Zapisano
' namerno -- brisanje "mrtvog" koda po grepu bez kapije je tacno ta klasa greske.
'
' Umire zajedno sa kolonom Otkup.BrojOtpremnice u S3e-2, kad S4 prevede kaskadu
' zbirne na kanon. Do tada: otpID prazan -> ocisti (unbind), guarded na kolonu.
Public Sub SetOtkupBrojOtpremnice(ByVal rowIndex As Long, ByVal otpID As String)
    On Error Resume Next
    If GetColumnIndex(TBL_OTKUP, COL_OTK_BROJ_OTPREMNICE) = 0 Then Exit Sub
    Dim broj As String: broj = ""
    If Len(Trim$(otpID)) > 0 Then _
        broj = Trim$(CStr(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_BROJ)))
    UpdateCell TBL_OTKUP, rowIndex, COL_OTK_BROJ_OTPREMNICE, broj
End Sub

' Kojoj AKTIVNOJ otpremnici otkup pripada, ili "" kad nijednoj.
'
' Pripadnost zivi iskljucivo u tblOtpremnicaIzvori (S4.1e) -- nema kolone na
' otkupu koja bi na to odgovorila, i to je namerno. Javni ulaz postoji zato sto
' pisac otkupa mora da pita pre ispravke (A13), a duplirana implementacija istog
' pitanja je upravo ono sto je kanonsko clanstvo trebalo da ukine.
'
' Dva aktivna clanstva podizu gresku, ne vracaju jedno od njih -- v. citac ispod.
Public Function OtpremnicaZaOtkup(ByVal otkupID As String) As String
    Const SRC As String = "OtpremnicaZaOtkup"

    otkupID = Trim$(otkupID)
    If Len(otkupID) = 0 Then Exit Function

    Dim mapa As Object
    Set mapa = AktivnoOtpClanstvoPoKanonu(SRC)
    If mapa.Exists(UCase$(otkupID)) Then OtpremnicaZaOtkup = CStr(mapa(UCase$(otkupID)))
End Function

' Sastav otpremnice: OtkupID-evi njenih izvora, redom upisa.
'
' Isti strog citac koji koriste read-model (GetOtpremnicaProgress) i sva tri
' pisca -- dupli par, clanstvo bez otkupa ili otkup u dve aktivne otpremnice
' PADAJU, ne normalizuju se. Javni ulaz postoji za radni sto u F1 (lista
' BLOKOVI, S3b-2): ekran koji bi tblOtpremnicaIzvori citao sam pokazao bi
' operateru sastav koji izdavanje kasnije odbije.
Public Function IzvoriOtpremnice(ByVal otpremnicaID As String) As Collection
    Set IzvoriOtpremnice = OtpClanovi(Trim$(otpremnicaID), "IzvoriOtpremnice")
End Function

' Kanonsko clanstvo SVIH aktivnih otpremnica: UCase(OtkupID) -> OtpremnicaID.
'
' Javni ulaz za citaoce koji pitaju za MNOGO izvora odjednom (specifikacija
' blokova, S3b-2b): OtpremnicaZaOtkup po izvoru bi za svaki prosao celu tabelu
' clanstva. Isti UGOVOR kao OtpClanovi nad jednim dokumentom -- v. kapije u
' AktivnoOtpClanstvoPoKanonu.
Public Function AktivnoClanstvoOtpremnica() As Object
    Set AktivnoClanstvoOtpremnica = AktivnoOtpClanstvoPoKanonu("AktivnoClanstvoOtpremnica")
End Function

' Otkupi BEZ AKTIVNE OTPREMNICE (A-025, S3b-2b; odluka operatera 19.09.2026:
' lista "nevezanih", ne samo "izgubljenih").
'
' Kljuc: UCase(OtkupID) -> brojevi STORNIRANIH otpremnica ciji je bio izvor,
' spojeni ", " -- ili "" kad ga storno nije oslobodio. Samo nestornirani otkupi.
'
' "Vezan" znaci ISTO sto i za kapiju storna otkupa i za pisca izvora: clan
' aktivne otpremnice po AktivnoOtpClanstvoPoKanonu. Jedan strog citac, pa lista
' ne moze da ponudi za vezivanje otkup koji pisac smatra zauzetim. Blok ostaje
' bez otpremnice na tri nacina: upisan bez izabrane otpremnice, uklonjen iz
' nacrta, oslobodjen stornom otpremnice -- storno ne brise clanstvo, pa ono
' ostaje kao istorija ("bila u").
Public Function NevezaniOtkupi() As Object
    Const SRC As String = "NevezaniOtkupi"

    Dim res As Object
    Set res = CreateObject("Scripting.Dictionary")
    Set NevezaniOtkupi = res

    Dim otk As Variant
    otk = GetTableData(TBL_OTKUP)
    If Not IsArray(otk) Then Exit Function

    Dim aktivno As Object, bila As Object
    Set aktivno = AktivnoOtpClanstvoPoKanonu(SRC)
    Set bila = BivseOtpremniceIzvora(SRC)

    Dim cId As Long, cSto As Long, i As Long, oid As String
    cId = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, SRC)
    cSto = RequireColumnIndex(TBL_OTKUP, COL_STORNIRANO, SRC)
    For i = 1 To UBound(otk, 1)
        oid = UCase$(Trim$(NzToText(otk(i, cId))))
        If Len(oid) > 0 Then
            If StrComp(Trim$(NzToText(otk(i, cSto))), "Da", vbTextCompare) <> 0 Then
                If Not aktivno.Exists(oid) Then
                    If bila.Exists(oid) Then
                        res(oid) = CStr(bila(oid))
                    Else
                        res(oid) = ""
                    End If
                End If
            End If
        End If
    Next i
End Function

' UCase(OtkupID) -> brojevi STORNIRANIH otpremnica u cijem je sastavu bio,
' redom clanstva, spojeni ", ". Istorija za kolonu "bila u", ne veza: aktivno
' clanstvo cita samo AktivnoOtpClanstvoPoKanonu.
Private Function BivseOtpremniceIzvora(ByVal src As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set BivseOtpremniceIzvora = d

    Dim izv As Variant, otp As Variant
    izv = GetTableData(TBL_OTPREMNICA_IZVORI)
    If Not IsArray(izv) Then Exit Function
    otp = GetTableData(TBL_OTPREMNICA)
    If Not IsArray(otp) Then Exit Function

    ' Broj STORNIRANE otpremnice po ID-u.
    Dim brStor As Object
    Set brStor = CreateObject("Scripting.Dictionary")
    Dim oId As Long, oBr As Long, oSto As Long, j As Long
    oId = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_ID, src)
    oBr = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ, src)
    oSto = RequireColumnIndex(TBL_OTPREMNICA, COL_STORNIRANO, src)
    For j = 1 To UBound(otp, 1)
        If StrComp(Trim$(NzToText(otp(j, oSto))), "Da", vbTextCompare) = 0 Then
            brStor(UCase$(Trim$(NzToText(otp(j, oId))))) = Trim$(NzToText(otp(j, oBr)))
        End If
    Next j
    If brStor.count = 0 Then Exit Function

    ' Par (otkup, otpremnica) se broji jednom -- isti broj dve razlicite
    ' stornirane otpremnice ostaje dva puta, jer su to dva dokumenta.
    Dim vidjen As Object
    Set vidjen = CreateObject("Scripting.Dictionary")
    Dim cOtp As Long, cOtk As Long, i As Long, otpU As String, otkU As String
    cOtp = RequireColumnIndex(TBL_OTPREMNICA_IZVORI, COL_OPI_OTPREMNICA_ID, src)
    cOtk = RequireColumnIndex(TBL_OTPREMNICA_IZVORI, COL_OPI_OTKUP_ID, src)
    For i = 1 To UBound(izv, 1)
        otpU = UCase$(Trim$(NzToText(izv(i, cOtp))))
        otkU = UCase$(Trim$(NzToText(izv(i, cOtk))))
        If Len(otkU) > 0 And brStor.Exists(otpU) Then
            If Not vidjen.Exists(otkU & "|" & otpU) Then
                vidjen.Add otkU & "|" & otpU, True
                If d.Exists(otkU) Then
                    d(otkU) = CStr(d(otkU)) & ", " & CStr(brStor(otpU))
                Else
                    d(otkU) = CStr(brStor(otpU))
                End If
            End If
        End If
    Next i
End Function

' Da li je otpremnica IZDATA (nije DRAFT).
'
' Prazan status NIJE "izdato": nov pisac ga upisuje eksplicitno (S4.1e), pa je
' prazno polje zatecen red ili drift -- a tumaciti drift kao izdat dokument bi
' otvorilo bas kapiju koju A13 zatvara.
Public Function OtpremnicaJeIzdata(ByVal otpremnicaID As String) As Boolean
    otpremnicaID = Trim$(otpremnicaID)
    If Len(otpremnicaID) = 0 Then Exit Function

    OtpremnicaJeIzdata = IzdatoStatusJeIzdato(LookupValue( _
                             TBL_OTPREMNICA, COL_OTP_ID, otpremnicaID, _
                             COL_TRACE_IZDATO_STATUS))
End Function

' JEDNO pravilo za "otpremnica je izdata", nad vrednoscu celije -- da citaoci
' koji vec drze red ne rade LookupValue po redu (review #362, P1).
'
'   IZDATO       -> izdata: gajbe knjizene, izvori revalidirani, ocekivano = povezano
'   PROSLEDJENO  -> izdata, i dalje prosledjena (buduci sync ka PWA/kupcu, ADR-0001)
'   DRAFT        -> nije: najava, bez izvora i bez knjizenih gajbi
'   prazno       -> nije
'   nepoznato    -> nije
'
' PROSLEDJENO MORA biti izdato: prelaz IZDATO -> PROSLEDJENO kaze samo da je
' dokument otisao dalje. Da ga pravilo ne broji, sync bi retroaktivno izbrisao
' vec otpremljenu robu iz izvestaja i zabranio stampu istog fizickog dokumenta
' (review #362, drugi krug).
'
' NAMERNO nije stari ugovor modDokumentInvariant.DocIsIssued ("sve osim DRAFT"):
' tamo prazan status znaci izdato, jer stari lanac nije imao nacrt. U novom
' modelu otpremnica nastaje kao nacrt, pa prazan ili nepoznat status nije
' dokaz izdavanja -- samo imenovani statusi to jesu.
'
' Operativni citaoci (roba po vozacu, roba po OM, stampa) broje samo izdate.
' Mreza F2 vidi i nacrte, jer bas tu operater radi sa njima.
Public Function IzdatoStatusJeIzdato(ByVal v As Variant) As Boolean
    Select Case UCase$(Trim$(NzToText(v)))
        Case UCase$(IZDATO_IZDATO), UCase$(IZDATO_PROSLEDJENO)
            IzdatoStatusJeIzdato = True
    End Select
End Function

' OtkupID -> OtpremnicaID, za sve NEstornirane otpremnice.
'
' STROG LOADER CELOG CLANSTVA, sa ISTIM ugovorom koji OtpClanovi drzi nad jednim
' dokumentom (review #364, P1):
'   - clanstvo bez OtpremnicaID-a ili bez OtkupID-a,
'   - roditelj koji ne postoji tacno jednom,
'   - dete koje ne postoji tacno jednom,
'   - isti par (otpremnica, otkup) dvaput,
'   - otkup u dve AKTIVNE otpremnice
' su tvrde greske, ne stanja koja se normalizuju.
'
' Do S3b-2b je ovaj prolaz bio SLABIJI od citaca jednog dokumenta: prazne ID-eve
' je preskakao, a postojanje dokumenata nije proveravao. Korupcija je tako
' postajala UREDAN poslovni odgovor -- specifikacija bi nad clanstvom na
' nepostojeci otkup odstampala PDF bez tog izvora (validan sibling zadovoljava
' kapiju "izdata bez izvora"), a lista nevezanih bi slobodan blok sa clanstvom
' na nepostojecu otpremnicu proglasila zauzetim i sklonila ga. Citalac koji
' korupciju pretvori u naizgled ispravan odgovor je tacno ono sto ovaj refaktor
' uklanja, pa ugovor stoji na JEDNOM mestu -- ovde, a ne u svakom pozivaocu.
Private Function AktivnoOtpClanstvoPoKanonu(ByVal src As String) As Object
    Dim mapa As Object
    Set mapa = CreateObject("Scripting.Dictionary")
    Set AktivnoOtpClanstvoPoKanonu = mapa

    Dim izv As Variant
    izv = GetTableData(TBL_OTPREMNICA_IZVORI)
    If Not IsArray(izv) Then Exit Function

    Dim cOtp As Long, cOtk As Long
    cOtp = RequireColumnIndex(TBL_OTPREMNICA_IZVORI, COL_OPI_OTPREMNICA_ID, src)
    cOtk = RequireColumnIndex(TBL_OTPREMNICA_IZVORI, COL_OPI_OTKUP_ID, src)

    ' Brojevi redova po ID-u: jedan prolaz po tabeli, umesto RequireTacnoJedan
    ' po clanu (to bi bio pun sken tabele za svaki red clanstva).
    Dim brojOtp As Object, brojOtk As Object, stornirane As Object
    Set brojOtp = BrojRedovaPoID(TBL_OTPREMNICA, COL_OTP_ID, src)
    Set brojOtk = BrojRedovaPoID(TBL_OTKUP, COL_OTK_ID, src)
    Set stornirane = StorniraneOtpremnice()

    Dim vidjenPar As Object
    Set vidjenPar = CreateObject("Scripting.Dictionary")

    Dim i As Long, otpID As String, otkupID As String, par As String
    For i = 1 To UBound(izv, 1)
        otpID = Trim$(NzToText(izv(i, cOtp)))
        otkupID = Trim$(NzToText(izv(i, cOtk)))

        If Len(otpID) = 0 Then
            Err.Raise vbObjectError + 1331, src, _
                      "Clanstvo bez OtpremnicaID-a u " & TBL_OTPREMNICA_IZVORI & _
                      " (red " & i & ")."
        End If
        If Len(otkupID) = 0 Then
            Err.Raise vbObjectError + 1332, src, _
                      "Clanstvo bez OtkupID-a, otpremnica " & otpID & "."
        End If

        ' Roditelj i dete moraju da postoje TACNO jednom: clanstvo na dokument
        ' koga nema nije "manji zbir" nego kvar.
        If ClanstvoBrojPoID(brojOtp, otpID) <> 1 Then
            Err.Raise vbObjectError + 1334, src, _
                      "Clanstvo pokazuje na otpremnicu koja ne postoji tacno jednom: " & _
                      otpID & " (nadjeno " & ClanstvoBrojPoID(brojOtp, otpID) & _
                      ", otkup " & otkupID & ")."
        End If
        If ClanstvoBrojPoID(brojOtk, otkupID) <> 1 Then
            Err.Raise vbObjectError + 1335, src, _
                      "Clanstvo pokazuje na otkup koji ne postoji tacno jednom: " & _
                      otkupID & " (nadjeno " & ClanstvoBrojPoID(brojOtk, otkupID) & _
                      ", otpremnica " & otpID & ")."
        End If

        par = UCase$(otpID) & "|" & UCase$(otkupID)
        If vidjenPar.Exists(par) Then
            Err.Raise vbObjectError + 1333, src, _
                      "Kanonsko clanstvo je nekonzistentno: par (" & otpID & ", " & _
                      otkupID & ") postoji vise puta."
        End If
        vidjenPar.Add par, True

        If Not stornirane.Exists(UCase$(otpID)) Then
            If mapa.Exists(UCase$(otkupID)) Then
                Err.Raise vbObjectError + 1305, src, _
                          "Kanonsko clanstvo je nekonzistentno: otkup " & _
                          otkupID & " ima dva aktivna zapisa clanstva (" & _
                          CStr(mapa(UCase$(otkupID))) & " i " & otpID & ")."
            End If
            mapa.Add UCase$(otkupID), otpID
        End If
    Next i
End Function

' UCase(ID) -> broj redova sa tim ID-em, jednim prolazom kroz tabelu.
Private Function BrojRedovaPoID(ByVal tblName As String, ByVal colName As String, _
                                ByVal src As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set BrojRedovaPoID = d

    Dim data As Variant
    data = GetTableData(tblName)
    If Not IsArray(data) Then Exit Function

    Dim c As Long, i As Long, k As String
    c = RequireColumnIndex(tblName, colName, src)
    For i = 1 To UBound(data, 1)
        k = UCase$(Trim$(NzToText(data(i, c))))
        If Len(k) > 0 Then
            If d.Exists(k) Then
                d(k) = CLng(d(k)) + 1
            Else
                d.Add k, 1&
            End If
        End If
    Next i
End Function

Private Function ClanstvoBrojPoID(ByVal brojevi As Object, ByVal id As String) As Long
    Dim k As String
    k = UCase$(Trim$(id))
    If brojevi.Exists(k) Then ClanstvoBrojPoID = CLng(brojevi(k))
End Function

Private Function StorniraneOtpremnice() As Object
    Dim s As Object
    Set s = CreateObject("Scripting.Dictionary")
    Set StorniraneOtpremnice = s

    Dim data As Variant
    data = GetTableData(TBL_OTPREMNICA)
    If Not IsArray(data) Then Exit Function

    Dim cID As Long, cStorno As Long
    cID = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_ID)
    cStorno = GetColumnIndex(TBL_OTPREMNICA, COL_STORNIRANO)
    If cID = 0 Or cStorno = 0 Then Exit Function

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If StrComp(Trim$(NzToText(data(i, cStorno))), "Da", vbTextCompare) = 0 Then
            s(UCase$(Trim$(NzToText(data(i, cID))))) = True
        End If
    Next i
End Function

Private Function OtpRedHeadera(ByVal otpremnicaID As String, _
                               ByVal src As String) As Long
    If Len(Trim$(otpremnicaID)) = 0 Then
        Err.Raise vbObjectError + 1306, src, "Prazan OtpremnicaID."
    End If

    Dim data As Variant
    data = GetTableData(TBL_OTPREMNICA)
    If Not IsArray(data) Then
        Err.Raise vbObjectError + 1307, src, "Tabela otpremnica je prazna."
    End If

    OtpRedHeadera = NadjiJedanRedOtpremnice(data, _
                        RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_ID, src), _
                        otpremnicaID, src)
End Function

' DRAFT je JEDINO stanje u kom se clanstvo i ocekivanje menjaju i u kom dokument
' sme da se izda. Posle izdavanja je sastav istorijska cinjenica, a izmena je
' NOVA VERZIJA (A13/A14/A15) -- ne izmena na mestu.
Private Sub RequireOtpDraft(ByVal otpremnicaID As String, ByVal rOtp As Long, _
                            ByVal src As String)
    Dim data As Variant
    data = GetTableData(TBL_OTPREMNICA)

    If StrComp(Trim$(NzToText(data(rOtp, _
               RequireColumnIndex(TBL_OTPREMNICA, COL_STORNIRANO, src)))), _
               "Da", vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 1308, src, _
                  "Otpremnica je stornirana: " & otpremnicaID
    End If

    Dim status As String
    status = UCase$(Trim$(NzToText(data(rOtp, _
             RequireColumnIndex(TBL_OTPREMNICA, COL_TRACE_IZDATO_STATUS, src)))))

    If status <> UCase$(IZDATO_DRAFT) Then
        Err.Raise vbObjectError + 1309, src, _
                  "Otpremnica nije DRAFT nego '" & status & "': " & otpremnicaID & _
                  ". Sastav izdatog dokumenta je istorijska cinjenica; izmena je " & _
                  "nova verzija (A13)."
    End If
End Sub

' --- sitni helperi ----------------------------------------------------------
Private Function OtpBroj(ByVal d As Object, ByVal kljuc As String) As Double
    If d.Exists(kljuc) Then OtpBroj = CDbl(d(kljuc))
End Function

Private Function OtpDbl(ByVal v As Variant) As Double
    If IsNumeric(v) Then OtpDbl = CDbl(v)
End Function

Private Sub RequireValidKlasa(ByVal klasa As String, ByVal src As String)
    Select Case Trim$(CStr(klasa))
        Case KLASA_I, KLASA_II
            Exit Sub
    End Select

    Err.Raise vbObjectError + 1325, src, "Neispravna klasa: " & klasa
End Sub

Private Function OtpHdrKljucPoznat(ByVal kljuc As String) As Boolean
    Select Case LCase$(Trim$(kljuc))
        Case "datum", "stanicaid", "vozacid", "kulturaid", "brojotpremnice", _
             "tipambalaze"
            OtpHdrKljucPoznat = True
    End Select
End Function

' Zatvoren spisak kljuceva, kao na otkupu. VrstaVoca / SortaVoca NISU na spisku:
' oni su snapshot KulturaID-a i writer ih sam prepisuje.
'
' "Cena" je SKINUTA sa spiska u S3a (odluka S14.8 t. 2). Predlog cene je od tada
' po klasi, na stavci. Kljuc se ODBIJA umesto da se ignorise: pozivalac koji i
' dalje salje cenu na zaglavlju misli da je poslao cenu, a ona bi nestala bez
' reci -- a bas zbog tihog gubitka jedan broj za dve klase i menjamo.
Private Sub OtpHdrProveriKljuceve(ByVal h As Object, ByVal src As String)
    Dim kljuc As Variant

    For Each kljuc In h.Keys
        If Not OtpHdrKljucPoznat(CStr(kljuc)) Then
            Err.Raise vbObjectError + 1310, src, _
                      "Header ima nepoznat kljuc: " & CStr(kljuc) & _
                      ". VrstaVoca/SortaVoca su snapshot KulturaID-a, " & _
                      "Kolicina/KolAmbalaze/Klasa idu na STAVKU."
        End If
    Next kljuc
End Sub

Private Function OtpOcekKljucPoznat(ByVal kljuc As String) As Boolean
    Select Case LCase$(Trim$(kljuc))
        Case "klasa", "kolicina", "kolambalaze", "predlogcena"
            OtpOcekKljucPoznat = True
    End Select
End Function

' Ocekivanje je ono sto operater PRIJAVLJUJE: klasa, kolicina, ambalaza i
' predlog cene TE KLASE. BrutoKg tu ne postoji -- njega izdavanje izvodi iz
' izvora, jer ga operater na otpremnici i ne meri.
'
' PredlogCena NIJE novac ovog dokumenta (v. modConfig, tblOtpremnicaStavke):
' ona samo prefiluje formu otkupa. Do S3a je stajala na zaglavlju kao jedan
' broj za obe klase.
Private Sub OtpOcekProveriKljuceve(ByVal s As Object, ByVal idx As Long, _
                                   ByVal src As String)
    Dim kljuc As Variant

    For Each kljuc In s.Keys
        If Not OtpOcekKljucPoznat(CStr(kljuc)) Then
            Err.Raise vbObjectError + 1326, src, _
                      "Ocekivana stavka " & CStr(idx) & " ima nepoznat kljuc: " & _
                      CStr(kljuc) & ". Dozvoljeni su Klasa / Kolicina / " & _
                      "KolAmbalaze / PredlogCena."
        End If
    Next kljuc
End Sub

' Opcion broj ocekivane stavke: kljuc sme da nedostaje (stariji pozivalac,
' jednopotezni put), ali ako postoji mora biti broj.
Private Function OtpStavkaBrojOpcion(ByVal s As Object, ByVal kljuc As String, _
                                     ByVal idx As Long, ByVal src As String) As Double
    If Not s.Exists(kljuc) Then Exit Function

    Dim v As Variant
    v = s(kljuc)
    If IsEmpty(v) Then Exit Function
    If Len(Trim$(NzToText(v))) = 0 Then Exit Function

    OtpStavkaBrojOpcion = OtpStavkaBroj(s, kljuc, idx, src)
End Function

Private Function OtpStavkaVrednost(ByVal s As Object, ByVal kljuc As String, _
                                   ByVal idx As Long, ByVal src As String) As Variant
    If Not s.Exists(kljuc) Then
        Err.Raise vbObjectError + 1327, src, _
                  "Ocekivana stavka " & CStr(idx) & " nema kljuc: " & kljuc
    End If

    OtpStavkaVrednost = s(kljuc)
End Function

Private Function OtpStavkaBroj(ByVal s As Object, ByVal kljuc As String, _
                               ByVal idx As Long, ByVal src As String) As Double
    Dim v As Variant
    v = OtpStavkaVrednost(s, kljuc, idx, src)

    If Not IsNumeric(v) Then
        Err.Raise vbObjectError + 1328, src, _
                  "Ocekivana stavka " & CStr(idx) & ", polje " & kljuc & _
                  " nije broj: " & NzToText(v)
    End If

    OtpStavkaBroj = CDbl(v)
End Function

' Kljuc mora postojati, vrednost sme biti prazna. Razlika prema HdrObavezan je
' namerna: nedostajuci kljuc je uvek greska pozivaoca (tipfeler), a prazna
' vrednost je legitiman podatak za polja koja domen ne trazi uvek.
Private Function OtpHdrObavezanKljuc(ByVal h As Object, ByVal kljuc As String, _
                                     ByVal src As String) As String
    If Not h.Exists(kljuc) Then
        Err.Raise vbObjectError + 1334, src, _
                  "Header nema obavezan kljuc: " & kljuc
    End If

    OtpHdrObavezanKljuc = Trim$(NzToText(h(kljuc)))
End Function

Private Function OtpHdrBrojOpcion(ByVal h As Object, ByVal kljuc As String, _
                                  ByVal src As String) As Double
    If Not h.Exists(kljuc) Then Exit Function

    Dim v As Variant
    v = h(kljuc)
    If IsEmpty(v) Then Exit Function
    If Len(Trim$(NzToText(v))) = 0 Then Exit Function

    If Not IsNumeric(v) Then
        Err.Raise vbObjectError + 1311, src, _
                  "Header polje " & kljuc & " nije broj: " & NzToText(v)
    End If

    OtpHdrBrojOpcion = CDbl(v)
End Function

' ============================================================
' ZBIRNA VALIDIERUNG
' ============================================================
Public Function ValidateZbirna(ByVal brojZbirne As String) As Variant
    ' Prueft Summe Otpremnice vs Zbirna
    ' Returns: Array(SumaOtpKg, ZbirnaKg, RazlikaKg, ValidKg,
    '                SumaOtpAmb, ZbirnaAmb, RazlikaAmb)
    On Error GoTo EH

    Dim otpData As Variant
    otpData = GetOtpremniceByZbirna(brojZbirne)

    If IsArray(otpData) Then otpData = ExcludeStornirano(otpData, TBL_OTPREMNICA)

    Dim sumaOtpKg As Double
    Dim sumaOtpAmb As Long

    ' Kilaza i gajbe otpremnice su na STAVKAMA (S3b) -- sa zaglavlja bi ova
    ' provera svaku zbirnu poredila sa nulom i proglasavala je neispravnom.
    If Not IsEmpty(otpData) Then
        Dim colOtpID As Long
        colOtpID = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_ID, _
                                      "modDokumenta.ValidateZbirna")

        Dim zbirStav As Object
        Set zbirStav = ZbirStavkiPoOtpremnici()

        Dim i As Long, z As Variant
        For i = 1 To UBound(otpData, 1)
            z = ZbirStavkiZaOtpremnicu(zbirStav, Trim$(NzToText(otpData(i, colOtpID))), _
                                       "modDokumenta.ValidateZbirna")
            sumaOtpKg = sumaOtpKg + CDbl(z(0))
            sumaOtpAmb = sumaOtpAmb + CLng(z(2))
        Next i
    End If

    Dim zbirnaKg As Double
    Dim zbirnaAmb As Long

    Dim zbrData As Variant
    zbrData = GetTableData(TBL_ZBIRNA)

    If IsArray(zbrData) Then zbrData = ExcludeStornirano(zbrData, TBL_ZBIRNA)

    If Not IsEmpty(zbrData) Then
        Dim colZbrBroj As Long
        Dim colZbrKol As Long
        Dim colZbrAmb As Long

        colZbrBroj = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, _
                                        "modDokumenta.ValidateZbirna")
        colZbrKol = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KOLICINA, _
                                       "modDokumenta.ValidateZbirna")
        colZbrAmb = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KOL_AMB, _
                                       "modDokumenta.ValidateZbirna")

        For i = 1 To UBound(zbrData, 1)
            If CStr(zbrData(i, colZbrBroj)) = brojZbirne Then
                If IsNumeric(zbrData(i, colZbrKol)) Then zbirnaKg = zbirnaKg + CDbl(zbrData(i, colZbrKol))
                If IsNumeric(zbrData(i, colZbrAmb)) Then zbirnaAmb = zbirnaAmb + CLng(zbrData(i, colZbrAmb))
            End If
        Next i
    End If

    ValidateZbirna = Array(sumaOtpKg, zbirnaKg, sumaOtpKg - zbirnaKg, _
                           (Abs(sumaOtpKg - zbirnaKg) < 0.01), _
                           sumaOtpAmb, zbirnaAmb, sumaOtpAmb - zbirnaAmb)

    Exit Function

EH:
    LogErr "modDokumenta.ValidateZbirna"
    ValidateZbirna = Array(0#, 0#, 0#, False, 0&, 0&, 0&)
End Function

Public Function ValidateZbirnaPreUnosa(ByVal brojZbirne As String, _
                                      ByVal inputKgKlI As Double, _
                                      ByVal inputKgKlII As Double, _
                                      ByVal inputAmb As Long) As Variant
    On Error GoTo EH

    Dim otpData As Variant
    otpData = GetOtpremniceByZbirna(brojZbirne)

    If IsArray(otpData) Then otpData = ExcludeStornirano(otpData, TBL_OTPREMNICA)

    Dim sumaKgKlI As Double
    Dim sumaKgKlII As Double
    Dim sumaAmb As Long

    ' Klasa, kilaza i gajbe su na STAVKAMA (S3b). Jedna otpremnica moze da nosi
    ' obe klase, pa red zaglavlja vise ne pripada jednoj klasi.
    If IsArray(otpData) Then
        Dim colOtpID As Long
        Dim i As Long, s As Long

        colOtpID = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_ID, _
                                      "modDokumenta.ValidateZbirnaPreUnosa")

        Dim stavkeDok As Object
        Set stavkeDok = StavkeOtpremnicePoDokumentu()

        Dim stavke As Collection, stavka As Variant, rowKlasa As String

        For i = 1 To UBound(otpData, 1)
            Set stavke = StavkeZaOtpremnicu(stavkeDok, _
                             Trim$(NzToText(otpData(i, colOtpID))), _
                             "modDokumenta.ValidateZbirnaPreUnosa")
            For s = 1 To stavke.count
                stavka = stavke(s)
                rowKlasa = Trim$(CStr(stavka(3)))

                If rowKlasa = KLASA_I Then
                    sumaKgKlI = sumaKgKlI + CDbl(stavka(4))
                ElseIf rowKlasa = KLASA_II Then
                    sumaKgKlII = sumaKgKlII + CDbl(stavka(4))
                End If

                sumaAmb = sumaAmb + CLng(stavka(6))
            Next s
        Next i
    End If

    ValidateZbirnaPreUnosa = Array( _
        sumaKgKlI, inputKgKlI, sumaKgKlI - inputKgKlI, (Abs(sumaKgKlI - inputKgKlI) < 0.01), _
        sumaKgKlII, inputKgKlII, sumaKgKlII - inputKgKlII, (Abs(sumaKgKlII - inputKgKlII) < 0.01), _
        sumaAmb, inputAmb, sumaAmb - inputAmb _
    )

    Exit Function

EH:
    LogErr "modDokumenta.ValidateZbirnaPreUnosa"
    ValidateZbirnaPreUnosa = Array(0#, inputKgKlI, -inputKgKlI, False, _
                                   0#, inputKgKlII, -inputKgKlII, False, _
                                   0&, inputAmb, -inputAmb)
End Function

' Da li izvorne (nestornirane) otpremnice date zbirne imaju Klasu II.
' Koristi frmDokumenta (hard-blokada unosa zbirne kad je "Dve klase" iskljuceno --
' inace bi pisac dobio hasKlasaII:=False i Kl.II bi se tiho izgubila).
' Isti izvor kao validacija u formi: ValidateZbirnaPreUnosa -> val(4) = sumaKgKlII.
Public Function ZbirnaIzvorImaKlasuII(ByVal brojZbirne As String) As Boolean
    On Error GoTo EH

    If Trim$(brojZbirne) = "" Then Exit Function

    Dim v As Variant
    v = ValidateZbirnaPreUnosa(brojZbirne, 0, 0, 0)

    If IsArray(v) Then
        If UBound(v) >= 4 Then
            If IsNumeric(v(4)) Then ZbirnaIzvorImaKlasuII = (CDbl(v(4)) > 0)
        End If
    End If

    Exit Function

EH:
    LogErr "modDokumenta.ZbirnaIzvorImaKlasuII"
    ZbirnaIzvorImaKlasuII = False
End Function

' ============================================================
' Generacija dokumenta (COL_GENERACIJA_ID)
'
' PRAVILO (jedno mesto, svi writer-i): red nasledjuje generaciju od AKTIVNIH
' (nestorniranih) redova ISTOG broja dokumenta; ako takvih nema -> nova generacija.
'
' Time bez ijednog dodatnog parametra vazi:
'   - Klasa I i Klasa II istog upisa dele generaciju (druga nasledjuje od prve),
'     bez obzira da li ih pise Multi_TX, AutoChainHladnjaca (dva zasebna _TX
'     poziva), malina auto-zbirna (red po red) ili backfill prijemnice;
'   - ISPRAVKA posle storna dobija NOVU generaciju -- svi stariji redovi tog broja
'     su tada stornirani, pa se nema sta naslediti;
'   - dopuna aktivnog dokumenta ostaje u njegovoj generaciji.
'
' Kolona je obavezna (RequireColumnIndex): stvara je modSetup.EnsureSledljivostSchema
' na svakom startu. Bez nje upis pada umesto da tiho nastane red bez generacije.
' ============================================================

' Generacija za red koji se upisuje pod datim brojem dokumenta.
' scopePairs: parovi (kolona, vrednost) koji uz broj cine IDENTITET dokumenta --
' broj sam po sebi NIJE globalno jedinstven (npr. GenerateBrojPrijemnice racuna
' sekvencu po kupcu, pa dva kupca istog dana oba dobiju "1/050826").
Public Function GeneracijaIDZaBroj(ByVal tableName As String, _
                                   ByVal brojCol As String, _
                                   ByVal broj As String, _
                                   ParamArray scopePairs() As Variant) As String
    GeneracijaIDZaBroj = GeneracijaIDZaBrojArr(tableName, brojCol, broj, _
                                               ScopePairsToArray(scopePairs, "modDokumenta.GeneracijaIDZaBroj"))
End Function

' Radna varijanta (scope kao obican niz) -- ParamArray se ne prosledjuje dalje.
Private Function GeneracijaIDZaBrojArr(ByVal tableName As String, _
                                       ByVal brojCol As String, _
                                       ByVal broj As String, _
                                       ByVal sc As Variant) As String
    Const SRC As String = "modDokumenta.GeneracijaIDZaBroj"

    Dim cGen As Long
    cGen = RequireColumnIndex(tableName, COL_GENERACIJA_ID, SRC)

    Dim target As String
    target = Trim$(broj)

    If Len(target) > 0 Then
        Dim data As Variant
        data = GetTableData(tableName)

        If IsArray(data) Then data = ExcludeStornirano(data, tableName)

        If IsArray(data) Then
            Dim cBr As Long
            cBr = RequireColumnIndex(tableName, brojCol, SRC)

            Dim scCols() As Long
            Dim i As Long
            If IsArray(sc) Then
                ReDim scCols(0 To UBound(sc))
                For i = 0 To UBound(sc) Step 2
                    scCols(i) = RequireColumnIndex(tableName, CStr(sc(i)), SRC)
                Next i
            End If

            Dim best As Double, bestGen As String
            Dim r As Long, g As String, rank As Double
            Dim inScope As Boolean
            best = -1

            For r = 1 To UBound(data, 1)
                If Trim$(NzToText(data(r, cBr))) = target Then
                    inScope = True
                    If IsArray(sc) Then
                        For i = 0 To UBound(sc) Step 2
                            If Trim$(NzToText(data(r, scCols(i)))) <> Trim$(NzToText(sc(i + 1))) Then
                                inScope = False
                                Exit For
                            End If
                        Next i
                    End If

                    If inScope Then
                        g = Trim$(NzToText(data(r, cGen)))
                        If Len(g) > 0 Then
                            rank = IdRank(g)
                            If rank > best Then
                                best = rank
                                bestGen = g
                            End If
                        End If
                    End If
                End If
            Next r

            If Len(bestGen) > 0 Then
                GeneracijaIDZaBrojArr = bestGen
                Exit Function
            End If
        End If
    End If

    GeneracijaIDZaBrojArr = NewGeneracijaID(tableName)
End Function

' ParamArray -> obican niz (parni: kolona, neparni: vrednost). Prazan -> Empty.
Private Function ScopePairsToArray(ByVal scopePairs As Variant, _
                                   ByVal sourceName As String) As Variant
    If Not IsArray(scopePairs) Then Exit Function
    If UBound(scopePairs) < LBound(scopePairs) Then Exit Function

    Dim n As Long
    n = UBound(scopePairs) - LBound(scopePairs) + 1

    If (n Mod 2) <> 0 Then
        Err.Raise vbObjectError + 1017, sourceName, _
                  "scopePairs mora imati parove (kolona, vrednost)."
    End If

    Dim out() As Variant
    ReDim out(0 To n - 1)

    Dim i As Long
    For i = 0 To n - 1
        out(i) = scopePairs(LBound(scopePairs) + i)
    Next i

    ScopePairsToArray = out
End Function

' Nova generacija za dokument-tabelu ("GEN-00042").
Public Function NewGeneracijaID(ByVal tableName As String) As String
    Const SRC As String = "modDokumenta.NewGeneracijaID"

    RequireColumnIndex tableName, COL_GENERACIJA_ID, SRC

    NewGeneracijaID = GetNextID(tableName, COL_GENERACIJA_ID, "GEN-")

    If Len(Trim$(NewGeneracijaID)) = 0 Then
        Err.Raise vbObjectError + 1015, SRC, _
                  "GetNextID nije vratio GeneracijaID za " & tableName & "."
    End If
End Function

' Izracunaj i upisi generaciju na upravo dodat red. Koriste je i writer-i koji
' ne idu kroz Save* (PWA import, invariant rekalkulacija).
Public Sub ApplyGeneracijaID(ByVal tableName As String, ByVal rowIndex As Long, _
                             ByVal brojCol As String, ByVal broj As String, _
                             ParamArray vlasnikPairs() As Variant)
    Const SRC As String = "modDokumenta.ApplyGeneracijaID"

    If rowIndex <= 0 Then
        Err.Raise vbObjectError + 1016, SRC, _
                  "Neispravan red za upis generacije (" & tableName & ")."
    End If

    ' Identitet dokumenta = broj + vlasnik: broj sam nije jedinstven. Vlasnik je
    ' otpremnica -> StanicaID, prijemnica -> KupacID, zbirna -> VozacID + KupacID
    ' (sekvenca se broji po vozacu, a dokument pripada kupcu; sam BROJ je
    ' jedinstven -- SuggestNextBroj za ZBR vrti petlju dok ne nadje slobodan).
    RequireUpdateCell tableName, rowIndex, COL_GENERACIJA_ID, _
                      GeneracijaIDZaBrojArr(tableName, brojCol, broj, _
                                            ScopePairsToArray(vlasnikPairs, SRC)), SRC
End Sub

' Pecati NOVU generaciju na upravo dodat red -- bez gledanja na to sta jos stoji
' pod istim brojem.
'
' ApplyGeneracijaID resava DRUGI problem: tamo dva fizicka reda cine JEDAN
' dokument (Kl.I + Kl.II), pa drugi red mora da nasledi generaciju prvog. Nasledje
' je tacno samo za pisca koji jedan dokument deli na vise redova.
'
' Pisac koji svaki red pise kao ZASEBAN dokument mora ovo. Za njega je nasledje
' aktivno stetno: dva odvojena dokumenta bi dobila isti GeneracijaID, pa bi
'   - ZbirnaIdentResolve izbrojao activeLogicalCount = 1 (broji GENERACIJE),
'     dakle UNIQUE -- F4 pusta, B8 cuti, kolizija se ne prijavljuje;
'   - StornoZbirna, koji redove bira po generaciji (RedJeIzabranogDokumenta),
'     stornirao OBA dokumenta na jedan storno;
'   - modScrStorno vise ne bi mogao ni da ih razlikuje: skrivena kolona
'     identiteta u gridu je bas COL_GENERACIJA_ID (modScrDokumenti.IdKolonaTipa).
'
' Prvi korisnik je modMasterSync: PWA import pise JEDAN red po ClientRecordID-u
' (obe klase sabrane u "I/II"), a IsDuplicateZbirnaInMaster odbija ponovljen CRID
' pre upisa -- svaki red koji stigne do AppendRow je dokument koji nikad nije
' vidjen.
' ZBR-CHILD-01: generacija roditeljske zbirne za dati BROJ, ili prazno.
'
' FAIL-CLOSED: prazno se vraca za sve sto nije jednoznacno -- nema aktivne zbirne,
' dvosmislen broj, pokvaren identitet. Dete tada ostaje bez generacije, sto je
' legitimno stanje i znaci "citaj po broju, kao i do sada". Pogadjati identitet
' iz broja je tacno ono protiv cega cela ZBR-IDENT celina i postoji.
'
' Zove se JEDNOM PO BROJU, ne po redu: ZbirnaIdentResolve cita celu tblZbirna, pa
' bi poziv u petlji nad decom bio O(n*m). Petlje zato uzimaju gen jednom i salju
' ga u PoveziDeteNaZbirnu.
Public Function ZbirnaIDZaBroj(ByVal broj As String) As String
    Dim id As ZbirnaIdent
    On Error GoTo EH
    If Len(Trim$(NzToText(broj))) = 0 Then Exit Function
    id = ZbirnaIdentResolve(broj)
    If id.integrityStatus <> ZBR_INT_OK Then Exit Function
    If id.resolutionStatus <> ZBR_RES_UNIQUE Then Exit Function
    ZbirnaIDZaBroj = id.selectedZbirnaID
    Exit Function
EH:
    LogErr "modDokumenta.ZbirnaIDZaBroj", "broj=" & broj
End Function

' ZBR-CHILD-01: JEDINI put kojim dete dobija zbirnu u DVA upisa.
'
' Broj i generacija se upisuju ZAJEDNO. Dva odvojena upisa bi se pre ili kasnije
' razisla: neko doda putanju koja postavlja broj a zaboravi generaciju, i dete
' ostane sa TUDJOM generacijom -- gore od prazne, jer prazna bar znaci "ne znam".
'
' Zovu ga i SaveOtpremnica/SavePrijemnica posle AppendRow, iako je broj vec u
' rowData: prepis iste vrednosti je jeftin, a druga putanja bi znacila da se par
' moze raziciti. Jedini upis koji NE ide ovuda je PalAppendRow u modPaletniList,
' i tamo rizika nema -- oba polja idu u ISTOM append pozivu, pa ih nema sta da
' razdvoji.
'
' gen se prosledjuje, ne racuna ovde: pozivaoci su cesto petlje nad decom istog
' broja (v. ZbirnaIDZaBroj).
Public Sub PoveziDeteNaZbirnu(ByVal tableName As String, ByVal rowIndex As Long, _
                              ByVal brojCol As String, ByVal brojZbirne As String, _
                              ByVal gen As String, ByVal sourceName As String)
    RequireUpdateCell tableName, rowIndex, brojCol, brojZbirne, sourceName
    RequireUpdateCell tableName, rowIndex, COL_DETE_ZBIRNA_GEN, gen, sourceName
End Sub

' ZBR-CHILD-01: dete se odvezuje od zbirne -- oba polja, u istom potezu.
'
' Bez ovoga bi odvezano dete zadrzalo generaciju stornirane zbirne i izgledalo kao
' da joj i dalje pripada, dok mu je broj prazan.
Public Sub OdveziDeteOdZbirne(ByVal tableName As String, ByVal rowIndex As Long, _
                              ByVal brojCol As String, ByVal sourceName As String)
    PoveziDeteNaZbirnu tableName, rowIndex, brojCol, "", "", sourceName
End Sub

' ZBR-CHILD-01 faza 3: da li se CELA operacija sme suzavati.
'
' `SuziDecuNaGeneraciju` odlucuje po JEDNOM skupu, a poslovna mutacija dira vise
' tabela. Kad svaka odlucuje sama, jedna kaskada zna da bude pola scoped a pola
' po broju: otpremnice suzene na GEN-B, a prijemnice -- jer je jedna legacy --
' vracene na broj, pa se stornira i prijemnica GEN-A. Sve-ili-nista mora da vazi
' nad CELOM operacijom, ne nad tabelom.
'
' Odluka se zato racuna JEDNOM, nad svim tabelama koje ta operacija bira po broju,
' pa se svim selektorima prosledi ista: generacija (suzavaj) ili prazno (ne suzavaj).
'
' Poredi kroz BrojJednak namerno, iako cetiri pozivaoca porede tacno. BrojJednak
' je siri, pa je ovde skup kandidata NADSKUP stvarnog: ako svi u nadskupu nose
' generaciju, nosi je i svaki podskup. Greska ide samo u stranu "ne suzavaj",
' nikad u "suzi pogresno".
Public Function SvaAktivnaDecaNoseGeneraciju(ByVal tableName As String, _
                                             ByVal brojCol As String, _
                                             ByVal broj As String) As Boolean
    SvaAktivnaDecaNoseGeneraciju = True

    If Len(Trim$(NzToText(broj))) = 0 Then Exit Function

    Dim data As Variant: data = GetTableData(tableName)
    If IsEmpty(data) Then Exit Function
    If Not IsArray(data) Then Exit Function

    Dim cBroj As Long: cBroj = GetColumnIndex(tableName, brojCol)
    If cBroj = 0 Then Exit Function
    Dim cSt As Long: cSt = GetColumnIndex(tableName, COL_STORNIRANO)
    Dim cGen As Long: cGen = GetColumnIndex(tableName, COL_DETE_ZBIRNA_GEN)

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If BrojJednak(data(i, cBroj), broj) Then
            If cSt = 0 Or UCase$(Trim$(NzToText(data(i, cSt)))) <> "DA" Then
                ' Kandidat postoji, a tabela nema kolonu -> ne moze se scope-ovati.
                If cGen = 0 Then
                    SvaAktivnaDecaNoseGeneraciju = False
                    Exit Function
                End If
                If Len(Trim$(NzToText(data(i, cGen)))) = 0 Then
                    SvaAktivnaDecaNoseGeneraciju = False
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

' ZBR-CHILD-01 faza 3: JEDAN put kojim se skup dece suzava na JEDAN dokument.
'
' Pandan write choke point-u iznad. Namerno NE preuzima i izbor po broju: cetiri
' zatecena odlucivaca porede broj TACNO (Trim$(CStr(..)) = broj), a peti
' (DistinctActiveValues) kanonski preko BrojJednak. Da je ovaj put preuzeo i to
' poredjenje, ona cetiri bi se tiho PROSIRILA -- akter siri od zatecenog je bas
' smer koji ZBR-NORM-02 zove opasnim. Zato svaki pozivalac zadrzava svoje
' poredjenje, a ovde se centralizuje samo ono sto je faza 3: generacija.
'
' PRAVILO (sve-ili-nista, ne hibrid):
'   1) trazena generacija prazna       -> vrati kandidate nepromenjeno
'   2) bilo koji kandidat bez generacije -> vrati kandidate nepromenjeno
'   3) inace                            -> vrati samo one koje se poklapaju
'
' Zasto NE hibrid "poklapa se ILI je prazno": pod jednim brojem mogu stajati dva
' dokumenta, pa bi isti prazan red upao u skup OBA -- dupli detach, pogresan
' racun. Fallback je bit-identican zatecenom ponasanju, sto je i uslov da faza 3
' ne pomeri nijednu zatecenu brojku.
'
' `data` ide ByRef i mora biti BAS onaj snimak nad kojim pozivalac vrti petlju:
' drugo citanje unutar iste transakcije moglo bi da vidi drugo stanje, pa bi
' indeksi redova pokazivali na tudje redove. ByRef i zbog KOPIJA_NIZA.
Public Function SuziDecuNaGeneraciju(ByVal tableName As String, ByRef data As Variant, _
                                     ByVal kandidati As Collection, _
                                     ByVal gen As String) As Collection
    Set SuziDecuNaGeneraciju = kandidati

    If kandidati Is Nothing Then Exit Function
    If kandidati.count = 0 Then Exit Function
    If Len(Trim$(NzToText(gen))) = 0 Then Exit Function
    If IsEmpty(data) Then Exit Function
    If Not IsArray(data) Then Exit Function

    Dim cGen As Long: cGen = GetColumnIndex(tableName, COL_DETE_ZBIRNA_GEN)
    If cGen = 0 Then Exit Function

    Dim k As Long
    For k = 1 To kandidati.count
        If Len(Trim$(NzToText(data(CLng(kandidati(k)), cGen)))) = 0 Then Exit Function
    Next k

    Dim suzeno As New Collection
    For k = 1 To kandidati.count
        If StrComp(Trim$(NzToText(data(CLng(kandidati(k)), cGen))), _
                   Trim$(NzToText(gen)), vbTextCompare) = 0 Then
            suzeno.Add CLng(kandidati(k))
        End If
    Next k

    Set SuziDecuNaGeneraciju = suzeno
End Function

Public Sub ApplyNovaGeneracijaID(ByVal tableName As String, ByVal rowIndex As Long)
    Const SRC As String = "modDokumenta.ApplyNovaGeneracijaID"

    If rowIndex <= 0 Then
        Err.Raise vbObjectError + 1018, SRC, _
                  "Neispravan red za upis generacije (" & tableName & ")."
    End If

    RequireUpdateCell tableName, rowIndex, COL_GENERACIJA_ID, _
                      NewGeneracijaID(tableName), SRC
End Sub

' Bira redove dokumenta za prefill ispravke (frmDokumenta.Prefill*FromStornirana).
'
' Polazi od ANCHOR reda -- konkretnog PK-a stornirane (OldDocID iz correction
' context-a). Broj dokumenta se NE koristi kao identitet: nije globalno jedinstven
' (GenerateBrojPrijemnice racuna sekvencu po kupcu, pa dva kupca istog dana dobiju
' isti broj), pa bi pretraga po broju mogla da prefiluje tudji dokument.
'
' Kad anchor PK nije poznat (stariji context bez OldDocID), fallback je poslednje
' upisan red datog broja -- i tada se ostaje unutar generacije tog reda, pa se
' redovi dva vlasnika ne mesaju.
'
' Iz anchor-a se cita generacija (COL_GENERACIJA_ID) i uzimaju Kl.I i Kl.II SAMO
' iz nje. Bez generacije (red stariji od uvodjenja kolone) prefiluje se samo sam
' anchor -- konzervativno, jer pripadnost druge klase nije dokaziva.
' ============================================================
' IDENTITET DOKUMENTA NA GRANICI PREVEZIVANJA
'
' Broj dokumenta je LABELA, ne identitet: BrojPrijemnice se racuna po kupcu
' (GenerateBrojPrijemnice), broj zbirne po vozacu. Dva dokumenta lako dele broj.
'
' Identitet logickog dokumenta (Klasa I + II zajedno) je GeneracijaID -- kolonu
' pravi modSetup.EnsureSledljivostSchema na SVAKOM startu, a pecate je svi
' writer-i kroz ApplyGeneracijaID, sa vec tacnim kompozitnim vlasnistvom po
' tipu (otpremnica StanicaID, prijemnica KupacID, zbirna VozacID + KupacID).
'
' Rutine prevezivanja su do sada primale samo broj i skenirale tabelu po njemu.
' Ove dve funkcije im daju identitet.
' ============================================================

' Da li red tabele pripada izvornom dokumentu. Sa poznatom generacijom odlucuje
' iskljucivo PK; bez nje broj (pozivalac je pre toga dokazao jednoznacnost).
' ByRef: citac po celiji -- ByVal bi kopirao ceo niz po pozivu (v. KOPIJA_NIZA).
Public Function PripadaIzvoru(ByRef data As Variant, ByVal rowIdx As Long, _
                              ByVal cBroj As Long, ByVal cId As Long, _
                              ByVal broj As String, ByVal srcIds As Object) As Boolean
    If Not srcIds Is Nothing Then
        If srcIds.count > 0 Then
            If cId > 0 Then
                PripadaIzvoru = srcIds.Exists(Trim$(NzToText(data(rowIdx, cId))))
                Exit Function
            End If
        End If
    End If
    If cBroj > 0 Then PripadaIzvoru = (Trim$(NzToText(data(rowIdx, cBroj))) = Trim$(broj))
End Function

' GeneracijaID dokumenta kome pripada dati red (po PK-u). Prazno = red ne
' postoji ili nema generaciju (stari zapis, pre uvodjenja kolone).
Public Function GeneracijaPoID(ByVal tableName As String, ByVal idCol As String, _
                               ByVal docID As String) As String
    On Error Resume Next
    docID = Trim$(docID)
    If Len(docID) = 0 Then Exit Function
    GeneracijaPoID = Trim$(NzToText(LookupValue(tableName, idCol, docID, COL_GENERACIJA_ID)))
End Function

' Svi PK-evi koji pripadaju datoj generaciji. To je "koji su redovi OVAJ
' dokument" - i Klasa I i Klasa II, i nijedan tudji.
'
' Prazna generacija vraca prazan recnik, NE sve redove: pozivalac tada mora da
' padne na kapiju nad brojem, ne na tiho zahvatanje svega.
Public Function IdoviGeneracije(ByVal tableName As String, ByVal idCol As String, _
                                ByVal gen As String) As Object
    Dim res As Object
    Set res = CreateObject("Scripting.Dictionary")
    res.CompareMode = vbTextCompare
    Set IdoviGeneracije = res

    On Error Resume Next
    gen = Trim$(gen)
    If Len(gen) = 0 Then Exit Function

    Dim data As Variant: data = GetTableData(tableName)
    If IsEmpty(data) Then Exit Function

    Dim cId As Long, cGen As Long
    cId = GetColumnIndex(tableName, idCol)
    cGen = GetColumnIndex(tableName, COL_GENERACIJA_ID)
    If cId = 0 Or cGen = 0 Then Exit Function

    Dim i As Long, k As String
    For i = 1 To UBound(data, 1)
        If Trim$(NzToText(data(i, cGen))) = gen Then
            k = Trim$(NzToText(data(i, cId)))
            If Len(k) > 0 Then res(k) = True
        End If
    Next i
End Function

Public Sub PickPrefillRows(ByVal data As Variant, _
                           ByVal cBroj As Long, ByVal cKlasa As Long, _
                           ByVal cId As Long, ByVal cGen As Long, _
                           ByVal broj As String, ByVal oldDocID As String, _
                           ByRef outRowI As Long, ByRef outRowII As Long)
    outRowI = 0
    outRowII = 0

    If Not IsArray(data) Then Exit Sub

    Dim anchor As Long
    anchor = FindAnchorRow(data, cBroj, cId, broj, oldDocID)
    If anchor = 0 Then Exit Sub

    Dim genTop As String
    If cGen > 0 Then genTop = Trim$(NzToText(data(anchor, cGen)))

    If Len(genTop) = 0 Then
        If RowKlasaII(data, anchor, cKlasa) Then
            outRowII = anchor
        Else
            outRowI = anchor
        End If
        Exit Sub
    End If

    Dim bestI As Double, bestII As Double
    Dim r As Long, rank As Double
    bestI = -1
    bestII = -1

    For r = 1 To UBound(data, 1)
        If Trim$(NzToText(data(r, cGen))) = genTop Then
            rank = RowRank(data, r, cId)

            If RowKlasaII(data, r, cKlasa) Then
                If rank > bestII Then
                    bestII = rank
                    outRowII = r
                End If
            Else
                If rank > bestI Then
                    bestI = rank
                    outRowI = r
                End If
            End If
        End If
    Next r
End Sub

' Anchor: red sa datim PK-om; ako PK nije poznat -> poslednje upisan red datog broja.
Private Function FindAnchorRow(ByVal data As Variant, ByVal cBroj As Long, _
                               ByVal cId As Long, ByVal broj As String, _
                               ByVal oldDocID As String) As Long
    Dim r As Long

    If Len(Trim$(oldDocID)) > 0 And cId > 0 Then
        For r = 1 To UBound(data, 1)
            If Trim$(NzToText(data(r, cId))) = Trim$(oldDocID) Then
                FindAnchorRow = r
                Exit Function
            End If
        Next r
        ' PK je ZADAT ali ga u tabeli NEMA -> vrati prazno, bez fallback-a.
        '
        ' Fallback po broju postoji samo za STARE kontekste koji OldDocID
        ' uopste nemaju. Kad kontekst tvrdi konkretan dokument a njega nema,
        ' "uzmi poslednji red istog broja" znaci: prefiluj TUDJI dokument.
        ' Broj nije jedinstven (GenerateBrojPrijemnice broji po kupcu), pa je
        ' to realan scenario, a ne teorijski.
        Exit Function
    End If

    If cBroj <= 0 Then Exit Function

    Dim target As String
    target = Trim$(broj)
    If Len(target) = 0 Then Exit Function

    Dim best As Double, rank As Double
    best = -1

    For r = 1 To UBound(data, 1)
        If Trim$(NzToText(data(r, cBroj))) = target Then
            rank = RowRank(data, r, cId)
            If rank >= best Then
                best = rank
                FindAnchorRow = r
            End If
        End If
    Next r
End Function

' Rang reda: numericki sufiks ID-a (GetNextID je monoton), inace indeks reda.
' ByRef: citac po celiji -- ByVal bi kopirao ceo niz po pozivu (v. KOPIJA_NIZA).
Private Function RowRank(ByRef data As Variant, ByVal rowIndex As Long, _
                         ByVal cId As Long) As Double
    RowRank = -1

    If cId > 0 Then RowRank = IdRank(NzToText(data(rowIndex, cId)))
    If RowRank < 0 Then RowRank = CDbl(rowIndex)
End Function

' Klasa reda je II (prazna/nepoznata klasa se tretira kao I, kao u ostatku koda).
' ByRef: citac po celiji -- ByVal bi kopirao ceo niz po pozivu (v. KOPIJA_NIZA).
Private Function RowKlasaII(ByRef data As Variant, ByVal rowIndex As Long, _
                            ByVal cKlasa As Long) As Boolean
    If cKlasa <= 0 Then Exit Function
    RowKlasaII = (UCase$(Trim$(NzToText(data(rowIndex, cKlasa)))) = "II")
End Function

' Numericki sufiks ID-a ("OTP-00042" -> 42); -1 ako ga nema. GetNextID je monoton
' po tabeli, pa je veci sufiks = kasnije upisan red.
Private Function IdRank(ByVal idText As String) As Double
    Dim s As String
    s = Trim$(idText)

    Dim i As Long, ch As String, digits As String
    For i = Len(s) To 1 Step -1
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = ch & digits
        Else
            Exit For
        End If
    Next i

    If Len(digits) = 0 Then
        IdRank = -1
    Else
        IdRank = Val(digits)
    End If
End Function

' ============================================================
' PRIJEMNICA - Kunde wiegt bei Annahme
' ============================================================

Public Function SavePrijemnicaMulti_TX(ByVal datum As Date, _
                                       ByVal kupacID As String, _
                                       ByVal vozacID As String, _
                                       ByVal brojPrij As String, _
                                       ByVal brojZbirne As String, _
                                       ByVal vrstaVoca As String, _
                                       ByVal sortaVoca As String, _
                                       ByVal kolicinaI As Double, _
                                       ByVal cenaI As Double, _
                                       ByVal tipAmb As String, _
                                       ByVal kolAmb As Long, _
                                       ByVal kolAmbVracena As Long, _
                                       Optional ByVal hasKlasaII As Boolean = False, _
                                       Optional ByVal kolicinaII As Double = 0, _
                                       Optional ByVal cenaII As Double = 0, _
                                       Optional ByVal kolAmbII As Long = 0, _
                                       Optional ByVal brutoKgI As Double = 0, _
                                       Optional ByVal brutoKgII As Double = 0) As String
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    ' Sema pre upisa: AppendRow pise POZICIONO, pa tabela sa kolonom manje ili
    ' u pogresnom rasporedu tiho salje vrednosti u pogresna polja. To je gore od
    ' pada upisa -- greska nastaje u podacima, ne u logu.
    '
    ' Ide PRE BeginTx: kapija sme da digne gresku, a nema smisla otvarati
    ' transakciju koja se odmah rollback-uje.
    modSchema.SchemaReadyOrFail "SavePrijemnicaMulti_TX", _
        TBL_PRIJEMNICA & "|" & TBL_AMBALAZA & "|" & TBL_FAKTURA_STAVKE & _
        "|" & TBL_FAKTURE & "|" & TBL_PALETA & "|" & TBL_PALETA_STAVKA

    tx.BeginTx
    tx.AddTableSnapshot TBL_PRIJEMNICA
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_FAKTURA_STAVKE
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot TBL_PALETA
    tx.AddTableSnapshot TBL_PALETA_STAVKA

    ' Klasa I je opciona (kolicinaI = 0 -> snima se samo Klasa II). Bar jedna klasa.
    Dim hasKlasaI As Boolean: hasKlasaI = (kolicinaI > 0)
    If Not hasKlasaI And Not hasKlasaII Then
        Err.Raise vbObjectError + 1303, "SavePrijemnicaMulti_TX", _
                  "Mora postojati bar jedna klasa (I ili II)."
    End If

    Dim resultI As String
    If hasKlasaI Then
        resultI = SavePrijemnica( _
            datum, _
            kupacID, _
            vozacID, _
            brojPrij, _
            brojZbirne, _
            vrstaVoca, _
            sortaVoca, _
            kolicinaI, _
            cenaI, _
            tipAmb, _
            kolAmb, _
            kolAmbVracena, _
            KLASA_I, _
            brutoKgI)

        If resultI = "" Then
            Err.Raise vbObjectError + 1301, "SavePrijemnicaMulti_TX", _
                      "SavePrijemnica Klasa I fehlgeschlagen"
        End If
    End If

    Dim resultII As String
    If hasKlasaII Then
        resultII = SavePrijemnica( _
            datum, _
            kupacID, _
            vozacID, _
            brojPrij, _
            brojZbirne, _
            vrstaVoca, _
            sortaVoca, _
            kolicinaII, _
            cenaII, _
            tipAmb, _
            kolAmbII, _
            0, _
            KLASA_II, _
            brutoKgII)

        If resultII = "" Then
            Err.Raise vbObjectError + 1302, "SavePrijemnicaMulti_TX", _
                      "SavePrijemnica Klasa II fehlgeschlagen"
        End If
    End If

    If hasKlasaI And hasKlasaII Then
        SavePrijemnicaMulti_TX = resultI & " + " & resultII
    ElseIf hasKlasaI Then
        SavePrijemnicaMulti_TX = resultI
    Else
        SavePrijemnicaMulti_TX = resultII
    End If

    ' Paletizacija Klase I UNUTAR transakcije -> atomicno sa prijemnicom.
    Dim closedPal As Collection
    Set closedPal = New Collection
    If hasKlasaI And kolAmb > 0 Then
        PaletizePrijemnica prijemnicaID:=resultI, brojPrij:=brojPrij, _
            brojZbirne:=brojZbirne, vrstaVoca:=vrstaVoca, sortaVoca:=sortaVoca, _
            klasa:=KLASA_I, netoKg:=kolicinaI, brGajbica:=kolAmb, tipAmb:=tipAmb, _
            closedPalIDs:=closedPal
    End If

    ' Paletizacija Klase II (zasebne gajbe) -> u istu kolekciju zatvorenih paleta.
    If hasKlasaII And kolAmbII > 0 Then
        PaletizePrijemnica prijemnicaID:=resultII, brojPrij:=brojPrij, _
            brojZbirne:=brojZbirne, vrstaVoca:=vrstaVoca, sortaVoca:=sortaVoca, _
            klasa:=KLASA_II, netoKg:=kolicinaII, brGajbica:=kolAmbII, tipAmb:=tipAmb, _
            closedPalIDs:=closedPal
    End If

    tx.CommitTx

    ' Print/PDF zatvorenih paleta = post-commit side effect (bez rollback rizika).
    PaletniListOutputClosed closedPal

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
    LogError "SavePrijemnicaMulti_TX", errDesc, errNum
    Monitor_Error _
        moduleName:="modDokumenta", _
        procedureName:="SavePrijemnicaMulti_TX", _
        entityType:="Prijemnica", _
        entityID:=SavePrijemnicaMulti_TX, _
        correlationId:=brojPrij, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="DOKUMENT_SAVE_FAIL", _
        severity:="ERROR", _
        message:="SavePrijemnicaMulti_TX failed. BrojPrij=" & brojPrij & _
             "; BrojZbirne=" & brojZbirne & _
             "; KupacID=" & kupacID & _
             "; VozacID=" & vozacID & _
             "; HasKlasaII=" & CStr(hasKlasaII) & _
             "; Error=" & errDesc, _
        userId:="Operator", _
        moduleName:="modDokumenta", _
        procedureName:="SavePrijemnicaMulti_TX", _
        entityType:="Prijemnica", _
        entityID:=SavePrijemnicaMulti_TX, _
        correlationId:=brojPrij

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    SavePrijemnicaMulti_TX = ""

    PrintTxFailure "SavePrijemnicaMulti_TX", errSrc, errNum, errDesc
End Function

Public Function SavePrijemnica_TX(ByVal datum As Date, ByVal kupacID As String, _
                                   ByVal vozacID As String, ByVal brojPrij As String, _
                                   ByVal brojZbirne As String, ByVal vrstaVoca As String, _
                                   ByVal sortaVoca As String, ByVal kolicina As Double, _
                                   ByVal cena As Double, ByVal tipAmb As String, _
                                   ByVal kolAmb As Long, ByVal kolAmbVracena As Long, _
                                   Optional ByVal klasa As String = "I", _
                                   Optional ByVal brutoKg As Double = 0) As String
    Dim tx As New clsTransaction

    On Error GoTo EH

        ' Sema pre upisa: AppendRow pise POZICIONO (v. CreateOtkup_TX).
    modSchema.SchemaReadyOrFail "SavePrijemnica_TX", _
        TBL_PRIJEMNICA & "|" & TBL_AMBALAZA

tx.BeginTx
    tx.AddTableSnapshot TBL_PRIJEMNICA
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_FAKTURA_STAVKE
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot TBL_PALETA
    tx.AddTableSnapshot TBL_PALETA_STAVKA

    SavePrijemnica_TX = SavePrijemnica(datum, kupacID, vozacID, brojPrij, _
                                        brojZbirne, vrstaVoca, sortaVoca, _
                                        kolicina, cena, tipAmb, kolAmb, _
                                        kolAmbVracena, klasa, brutoKg)

    If SavePrijemnica_TX = "" Then
        Err.Raise vbObjectError + 1011, "SavePrijemnica_TX", _
                  "SavePrijemnica fehlgeschlagen"
    End If

    ' Paletizacija UNUTAR transakcije -> atomicno sa prijemnicom (gajbice bilo koje klase).
    Dim closedPal As Collection
    Set closedPal = New Collection
    If kolAmb > 0 Then
        PaletizePrijemnica prijemnicaID:=SavePrijemnica_TX, brojPrij:=brojPrij, _
            brojZbirne:=brojZbirne, vrstaVoca:=vrstaVoca, sortaVoca:=sortaVoca, _
            klasa:=klasa, netoKg:=kolicina, brGajbica:=kolAmb, tipAmb:=tipAmb, _
            closedPalIDs:=closedPal
    End If

    tx.CommitTx

    ' Print/PDF zatvorenih paleta = post-commit side effect (bez rollback rizika).
    PaletniListOutputClosed closedPal

    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogError "SavePrijemnica_TX", errDesc, errNum
    Monitor_Error _
        moduleName:="modDokumenta", _
        procedureName:="SavePrijemnica_TX", _
        entityType:="Prijemnica", _
        entityID:=SavePrijemnica_TX, _
        correlationId:=brojPrij, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="DOKUMENT_SAVE_FAIL", _
        severity:="ERROR", _
        message:="SavePrijemnica_TX failed. BrojPrij=" & brojPrij & _
             "; BrojZbirne=" & brojZbirne & _
             "; KupacID=" & kupacID & _
             "; VozacID=" & vozacID & _
             "; Error=" & errDesc, _
        userId:="Operator", _
        moduleName:="modDokumenta", _
        procedureName:="SavePrijemnica_TX", _
        entityType:="Prijemnica", _
        entityID:=SavePrijemnica_TX, _
        correlationId:=brojPrij

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    SavePrijemnica_TX = ""

    PrintTxFailure "SavePrijemnica_TX", errSrc, errNum, errDesc
End Function
    
Public Function SavePrijemnica(ByVal datum As Date, ByVal kupacID As String, _
                               ByVal vozacID As String, ByVal brojPrij As String, _
                               ByVal brojZbirne As String, ByVal vrstaVoca As String, _
                               ByVal sortaVoca As String, ByVal kolicina As Double, _
                               ByVal cena As Double, ByVal tipAmb As String, _
                               ByVal kolAmb As Long, ByVal kolAmbVracena As Long, _
                               Optional ByVal klasa As String = "I", _
                               Optional ByVal brutoKg As Double = 0) As String
    On Error GoTo EH

    Call ValidatePrijemnicaInput(kupacID, vozacID, brojPrij, brojZbirne, _
                             kolicina, cena, tipAmb, kolAmb, _
                             kolAmbVracena, klasa)

    Dim newID As String
    newID = GetNextID(TBL_PRIJEMNICA, COL_PRJ_ID, "PRJ-")

    If newID = "" Then
        Err.Raise vbObjectError + 1013, "SavePrijemnica", _
                  "GetNextID nije vratio PrijemnicaID."
    End If

    Dim rowData As Variant
    rowData = BuildPrijemnicaRowData( _
                newID, datum, kupacID, vozacID, brojPrij, brojZbirne, _
                vrstaVoca, sortaVoca, kolicina, cena, tipAmb, kolAmb, _
                kolAmbVracena, klasa)


    Dim appendedRow As Long
    appendedRow = AppendRow(TBL_PRIJEMNICA, rowData)

    If appendedRow <= 0 Then
        Err.Raise vbObjectError + 1014, "SavePrijemnica", _
                "AppendRow fehlgeschlagen fuer tblPrijemnica."
    End If

    ' ZBR-CHILD-01: generacija roditeljske zbirne (v. isti komentar u
    ' SaveOtpremnica). Prijemnica roditelja obicno IMA, pa je ovde retko prazna.
    PoveziDeteNaZbirnu TBL_PRIJEMNICA, appendedRow, COL_PRJ_BROJ_ZBIRNE, brojZbirne, _
                       ZbirnaIDZaBroj(brojZbirne), "modDokumenta.SavePrijemnica"
    ApplyGeneracijaID TBL_PRIJEMNICA, appendedRow, COL_PRJ_BROJ, brojPrij, _
                      COL_PRJ_KUPAC, kupacID

    ' Bruto tezina (preneto iz otkupa kad je OTKUP_BRUTO_UNOS) -> upis po imenu;
    ' prazno = neto. Kolona postoji posle EnsureDoradeSchema (na kraju tblPrijemnica).
    If brutoKg > 0 Then UpdateCell TBL_PRIJEMNICA, appendedRow, COL_PRJ_BRUTO, brutoKg

    ' Ambalaza je ENTITETSKI-relativna (smer iz ugla hladnjace / Kupca):
    ' 1. txt = pune gajbe koje hladnjaca PRIMA od zbirne -> Kupac ULAZ.
    If kolAmb > 0 Then
        TrackAmbalaza datum, tipAmb, kolAmb, "Ulaz", kupacID, "Kupac", _
                      vozacID, newID, DOK_TIP_PRIJEMNICA
    End If

    ' 2. txt = zamena: prazne gajbe koje hladnjaca VRACA (daje vozacu) -> Kupac IZLAZ.
    If kolAmbVracena > 0 Then
        TrackAmbalaza datum, tipAmb, kolAmbVracena, "Izlaz", kupacID, "Kupac", _
                      vozacID, newID, DOK_TIP_PRIJEMNICA
    End If

    RelinkFakturaStavke newID, brojPrij, klasa

    SavePrijemnica = newID
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogError "SavePrijemnica", errDesc, errNum
    On Error GoTo 0

    Err.Raise errNum, "SavePrijemnica", _
              "Source=" & errSrc & " | " & errDesc
End Function

' Gradi red tblZbirna PO IMENU kolone (AUD-003). Otporno na promenu redosleda
' kolona i na kolonu umetnutu u sredinu -- za razliku od pozicijskog Array(...).
Private Function BuildPrijemnicaRowData(ByVal prijemnicaID As String, _
                                        ByVal datum As Date, _
                                        ByVal kupacID As String, _
                                        ByVal vozacID As String, _
                                        ByVal brojPrij As String, _
                                        ByVal brojZbirne As String, _
                                        ByVal vrstaVoca As String, _
                                        ByVal sortaVoca As String, _
                                        ByVal kolicina As Double, _
                                        ByVal cena As Double, _
                                        ByVal tipAmb As String, _
                                        ByVal kolAmb As Long, _
                                        ByVal kolAmbVracena As Long, _
                                        ByVal klasa As String) As Variant
    Const SRC As String = "BuildPrijemnicaRowData"

    Dim colCount As Long
    colCount = TabelaBrojKolona(TBL_PRIJEMNICA)

    If colCount <= 0 Then
        Err.Raise vbObjectError + 1430, SRC, _
                  "Could not resolve tblPrijemnica column count."
    End If

    Dim rowData() As Variant
    ReDim rowData(0 To colCount - 1)

    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_ID, prijemnicaID, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_DATUM, datum, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_KUPAC, kupacID, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_VOZAC, vozacID, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_BROJ, brojPrij, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, brojZbirne, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_VRSTA, vrstaVoca, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_SORTA, sortaVoca, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_KOLICINA, kolicina, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_CENA, cena, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_TIP_AMB, tipAmb, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_KOL_AMB, kolAmb, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_KOL_AMB_VRACENA, kolAmbVracena, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_KLASA, klasa, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_FAKTURISANO, "Ne", SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_FAKTURA_ID, "", SRC

    If GetColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO) > 0 Then
        SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_STORNIRANO, "", SRC
    End If

    BuildPrijemnicaRowData = rowData
End Function

Public Function GetPrijemniceByKupac(ByVal kupacID As String, _
                                      Optional ByVal datumOd As Date = 0, _
                                      Optional ByVal datumDo As Date = 0, _
                                      Optional ByVal samoNefakturisano As Boolean = False) As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_PRIJEMNICA)

    If IsEmpty(data) Then
        GetPrijemniceByKupac = Empty
        Exit Function
    End If
    
    data = ExcludeStornirano(data, TBL_PRIJEMNICA)

    If IsEmpty(data) Then
        GetPrijemniceByKupac = Empty
        Exit Function
    End If

    Dim filters As New Collection
    Dim fp As clsFilterParam

    Set fp = New clsFilterParam
    fp.Init RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KUPAC, _
            "modDokumenta.GetPrijemniceByKupac"), "=", kupacID
    filters.Add fp

    If datumOd > 0 And datumDo > 0 Then
        Set fp = New clsFilterParam
        fp.Init RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_DATUM, _
                "modDokumenta.GetPrijemniceByKupac"), "BETWEEN", datumOd, datumDo
        filters.Add fp
    End If

    If samoNefakturisano Then
        Set fp = New clsFilterParam
        fp.Init RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_FAKTURISANO, _
                "modDokumenta.GetPrijemniceByKupac"), "<>", "Da"
        filters.Add fp
    End If

    GetPrijemniceByKupac = FilterArray(data, filters)
    Exit Function

EH:
    LogErr "modDokumenta.GetPrijemniceByKupac"
    GetPrijemniceByKupac = Empty
End Function

Public Function SaveKupciIzlaz_TX(ByVal datum As Date, _
                                  ByVal brojDok As String, _
                                  ByVal kupacNaziv As String, _
                                  ByVal kupacID As String, _
                                  ByVal vozacID As String, _
                                  ByVal tipAmb As String, _
                                  ByVal kolAmb As Long, _
                                  ByVal vrstaVoca As String, _
                                  ByVal novac As Double, _
                                  ByVal fakturaID As String, _
                                  ByVal napomena As String, _
                                  ByVal tipNovca As String) As Boolean
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    If kupacID = "" Then
        Err.Raise vbObjectError + 1601, "SaveKupciIzlaz_TX", _
                  "KupacID je obavezan."
    End If

    If kolAmb <= 0 And novac <= 0 Then
        Err.Raise vbObjectError + 1602, "SaveKupciIzlaz_TX", _
                  Poruka("DOK_ERR_NEMA_AMBALAZE_NOVCA")
    End If

    tx.BeginTx
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_FAKTURE

    If kolAmb > 0 Then
        TrackAmbalaza datum, tipAmb, kolAmb, _
                      "Izlaz", kupacID, "Kupac", _
                      vozacID, brojDok, DOK_TIP_IZLAZ_KUPCI
    End If

    If novac > 0 Then
        Dim novacID As String

        ' Ista kapija kao u SaveOMUlaz_TX, samo nad fakturom: vlasnistvo, storno
        ' stanje i TRENUTNO preostalo se citaju ovde, ne uzimaju iz parametara.
        Dim fakErr As String
        fakErr = UplataFakturaProblem(fakturaID, kupacID, novac)
        If Len(fakErr) > 0 Then
            Err.Raise vbObjectError + 1604, "SaveKupciIzlaz_TX", fakErr
        End If

        novacID = SaveNovac( _
            brojDok:=brojDok, _
            datum:=datum, _
            partner:=kupacNaziv, _
            partnerId:=kupacID, _
            entitetTip:="Kupac", _
            omID:="", _
            kooperantID:="", _
            fakturaID:=fakturaID, _
            vrstaVoca:=vrstaVoca, _
            tip:=tipNovca, _
            uplata:=novac, _
            isplata:=0, _
            napomena:=napomena, _
            otkupID:="")

        If novacID = "" Then
            Err.Raise vbObjectError + 1603, "SaveKupciIzlaz_TX", _
                      "SaveNovac fehlgeschlagen"
        End If

        If fakturaID <> "" Then
            UpdateFakturaStatus fakturaID
        End If
    End If

    tx.CommitTx
    Set tx = Nothing

    SaveKupciIzlaz_TX = True
    Exit Function

EH:
    LogErr "SaveKupciIzlaz_TX"

    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    SaveKupciIzlaz_TX = False
End Function

' ============================================================
' MANJAK - Schwundberechnung
' ============================================================

Public Function CalculateManjakPreview(ByVal brojZbirne As String, _
                                      ByVal pendingKgKlI As Double, _
                                      ByVal pendingKgKlII As Double) As Variant
    On Error GoTo EH

    Dim zbirnaKg As Double

    Dim zbrData As Variant
    zbrData = GetTableData(TBL_ZBIRNA)

    If IsArray(zbrData) Then zbrData = ExcludeStornirano(zbrData, TBL_ZBIRNA)

    If IsArray(zbrData) Then
        Dim colBroj As Long
        Dim colKol As Long
        Dim i As Long

        colBroj = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, _
                                     "modDokumenta.CalculateManjakPreview")
        colKol = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KOLICINA, _
                                    "modDokumenta.CalculateManjakPreview")

        For i = 1 To UBound(zbrData, 1)
            If CStr(zbrData(i, colBroj)) = brojZbirne Then
                If IsNumeric(zbrData(i, colKol)) Then zbirnaKg = zbirnaKg + CDbl(zbrData(i, colKol))
            End If
        Next i
    End If

    Dim prijKg As Double
    Dim prijData As Variant

    prijData = GetTableData(TBL_PRIJEMNICA)

    If IsArray(prijData) Then prijData = ExcludeStornirano(prijData, TBL_PRIJEMNICA)

    If IsArray(prijData) Then
        Dim colBrZbr As Long
        Dim colPrijKol As Long

        colBrZbr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, _
                                      "modDokumenta.CalculateManjakPreview")
        colPrijKol = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA, _
                                        "modDokumenta.CalculateManjakPreview")

        For i = 1 To UBound(prijData, 1)
            If CStr(prijData(i, colBrZbr)) = brojZbirne Then
                If IsNumeric(prijData(i, colPrijKol)) Then prijKg = prijKg + CDbl(prijData(i, colPrijKol))
            End If
        Next i
    End If

    prijKg = prijKg + pendingKgKlI + pendingKgKlII

    Dim manjakKg As Double
    Dim manjakPct As Double

    manjakKg = zbirnaKg - prijKg

    If zbirnaKg > 0 Then manjakPct = manjakKg / zbirnaKg * 100

    CalculateManjakPreview = Array(zbirnaKg, prijKg, manjakKg, manjakPct)
    Exit Function

EH:
    LogErr "modDokumenta.CalculateManjakPreview"
    CalculateManjakPreview = Array(0#, pendingKgKlI + pendingKgKlII, 0#, 0#)
End Function

' ============================================================
' PROSEK GAJBE - Durchschnittsgewicht pro Kaestchen
' ============================================================

Public Function CalculateProsekGajbeByZbirna(ByVal brojZbirne As String) As Double
    On Error GoTo EH

    If Trim$(brojZbirne) = "" Then Exit Function

    RequireColumnIndex TBL_ZBIRNA, COL_ZBR_BROJ, _
                       "modDokumenta.CalculateProsekGajbeByZbirna"
    RequireColumnIndex TBL_ZBIRNA, COL_ZBR_KOLICINA, _
                       "modDokumenta.CalculateProsekGajbeByZbirna"
    RequireColumnIndex TBL_ZBIRNA, COL_ZBR_KOL_AMB, _
                       "modDokumenta.CalculateProsekGajbeByZbirna"

    ' Dvoklasni doc: sumiraj kol i amb preko SVIH redova broja (ne samo prvi red).
    Dim kol As Double, amb As Double
    kol = SumByBroj(TBL_ZBIRNA, COL_ZBR_BROJ, brojZbirne, COL_ZBR_KOLICINA)
    amb = SumByBroj(TBL_ZBIRNA, COL_ZBR_BROJ, brojZbirne, COL_ZBR_KOL_AMB)

    If amb > 0 Then
        CalculateProsekGajbeByZbirna = kol / amb
    Else
        CalculateProsekGajbeByZbirna = 0
    End If

    Exit Function

EH:
    LogErr "modDokumenta.CalculateProsekGajbeByZbirna"
    CalculateProsekGajbeByZbirna = 0
End Function

' Sumira valCol preko SVIH redova gde brojCol = broj (obe klase dvorednog doc-a).
Private Function SumByBroj(ByVal tbl As String, ByVal brojCol As String, _
                           ByVal broj As String, ByVal valCol As String) As Double
    Dim data As Variant: data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Function
    ' Stornirani redovi ne ulaze u sumu (prosek gajbe je racunao i njih).
    If IsArray(data) Then data = ExcludeStornirano(data, tbl)
    If Not IsArray(data) Then Exit Function
    Dim cB As Long, cV As Long
    cB = GetColumnIndex(tbl, brojCol)
    cV = GetColumnIndex(tbl, valCol)
    If cB = 0 Or cV = 0 Then Exit Function
    Dim i As Long, s As Double
    For i = 1 To UBound(data, 1)
        If Trim$(NzToText(data(i, cB))) = Trim$(broj) Then
            If IsNumeric(data(i, cV)) Then s = s + CDbl(data(i, cV))
        End If
    Next i
    SumByBroj = s
End Function

' ============================================================
' Storno Verweiste
' ============================================================

Public Function GetVerwaisteDokumente(ByVal dokumentTip As String) As Variant
    ' Returns: 2D Array der Dokumente deren BrojZbirne auf eine
    '          stornierte Zbirna zeigt, die selbst aber NICHT storniert sind.
    '
    ' dokumentTip: "Otpremnica" oder "Prijemnica"
    '
    ' Otpremnica Returns: (OtpremnicaID, BrojOtp, BrojZbirne, VrstaVoca, Kolicina)
    ' Prijemnica Returns: (PrijemnicaID, BrojPrij, BrojZbirne, KupacNaziv, Kolicina)
    ' oder Empty
    On Error GoTo EH

    Dim zbrData As Variant
    zbrData = GetTableData(TBL_ZBIRNA)

    If IsEmpty(zbrData) Then
        GetVerwaisteDokumente = Empty
        Exit Function
    End If

    Dim colZbrBroj As Long
    Dim colZbrStorno As Long

    colZbrBroj = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, _
                                    "modDokumenta.GetVerwaisteDokumente")
    colZbrStorno = RequireColumnIndex(TBL_ZBIRNA, COL_STORNIRANO, _
                                      "modDokumenta.GetVerwaisteDokumente")

    Dim storniraneBrojevi As Object
    Set storniraneBrojevi = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim brz As String

    For i = 1 To UBound(zbrData, 1)
        If Trim$(NzToText(zbrData(i, colZbrStorno))) = "Da" Then
            brz = Trim$(NzToText(zbrData(i, colZbrBroj)))

            If brz <> "" Then
                If Not storniraneBrojevi.Exists(brz) Then
                    storniraneBrojevi.Add brz, True
                End If
            End If
        End If
    Next i

    For i = 1 To UBound(zbrData, 1)
        If Trim$(NzToText(zbrData(i, colZbrStorno))) <> "Da" Then
            brz = Trim$(NzToText(zbrData(i, colZbrBroj)))

            If storniraneBrojevi.Exists(brz) Then
                storniraneBrojevi.Remove brz
            End If
        End If
    Next i

    If storniraneBrojevi.count = 0 Then
        GetVerwaisteDokumente = Empty
        Exit Function
    End If

    Select Case dokumentTip
        Case "Otpremnica"
            GetVerwaisteDokumente = GetVerwaisteOtpremnice(storniraneBrojevi)

        Case "Prijemnica"
            GetVerwaisteDokumente = GetVerwaistePrijemnice(storniraneBrojevi)

        Case Else
            GetVerwaisteDokumente = Empty
    End Select

    Exit Function

EH:
    LogErr "modDokumenta.GetVerwaisteDokumente"
    GetVerwaisteDokumente = Empty
End Function

Private Function GetVerwaisteOtpremnice(ByVal storniraneBrojevi As Object) As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_OTPREMNICA)

    If IsEmpty(data) Then
        GetVerwaisteOtpremnice = Empty
        Exit Function
    End If

    Dim colID As Long
    Dim colBrOtp As Long
    Dim colBrZbr As Long
    Dim colVrsta As Long
    Dim colStorno As Long
    Dim zVerw As Variant

    colID = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_ID, _
                               "modDokumenta.GetVerwaisteOtpremnice")
    colBrOtp = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ, _
                                  "modDokumenta.GetVerwaisteOtpremnice")
    colBrZbr = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, _
                                  "modDokumenta.GetVerwaisteOtpremnice")
    colVrsta = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VRSTA, _
                                  "modDokumenta.GetVerwaisteOtpremnice")
    colStorno = RequireColumnIndex(TBL_OTPREMNICA, COL_STORNIRANO, _
                                   "modDokumenta.GetVerwaisteOtpremnice")

    ' Kilaza dolazi sa STAVKI (S3b) -- lista siroceta je lista za oporavak, pa
    ' bi nula pored svakog dokumenta operateru rekla da nema sta da se vrati.
    Dim zbirStav As Object
    Set zbirStav = ZbirStavkiPoOtpremnici()

    Dim count As Long
    Dim i As Long

    For i = 1 To UBound(data, 1)
        If Trim$(NzToText(data(i, colStorno))) = "Da" Then GoTo NextCount

        If storniraneBrojevi.Exists(Trim$(NzToText(data(i, colBrZbr)))) Then
            count = count + 1
        End If

NextCount:
    Next i

    If count = 0 Then
        GetVerwaisteOtpremnice = Empty
        Exit Function
    End If

    Dim result() As Variant
    ReDim result(1 To count, 1 To 5)

    Dim idx As Long
    Dim kol As Double

    For i = 1 To UBound(data, 1)
        If Trim$(NzToText(data(i, colStorno))) = "Da" Then GoTo NextRow

        If storniraneBrojevi.Exists(Trim$(NzToText(data(i, colBrZbr)))) Then
            idx = idx + 1

            zVerw = ZbirStavkiZaOtpremnicu(zbirStav, Trim$(NzToText(data(i, colID))), _
                                           "modDokumenta.GetVerwaisteOtpremnice")
            kol = CDbl(zVerw(0))

            result(idx, 1) = NzToText(data(i, colID))
            result(idx, 2) = NzToText(data(i, colBrOtp))
            result(idx, 3) = NzToText(data(i, colBrZbr))
            result(idx, 4) = NzToText(data(i, colVrsta))
            result(idx, 5) = kol
        End If

NextRow:
    Next i

    GetVerwaisteOtpremnice = result
    Exit Function

EH:
    LogErr "modDokumenta.GetVerwaisteOtpremnice"
    GetVerwaisteOtpremnice = Empty
End Function

Private Function GetVerwaistePrijemnice(ByVal storniraneBrojevi As Object) As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_PRIJEMNICA)

    If IsEmpty(data) Then
        GetVerwaistePrijemnice = Empty
        Exit Function
    End If

    Dim colID As Long
    Dim colBrPrij As Long
    Dim colBrZbr As Long
    Dim colKupac As Long
    Dim colKol As Long
    Dim colStorno As Long

    colID = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID, _
                               "modDokumenta.GetVerwaistePrijemnice")
    colBrPrij = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ, _
                                   "modDokumenta.GetVerwaistePrijemnice")
    colBrZbr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, _
                                  "modDokumenta.GetVerwaistePrijemnice")
    colKupac = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KUPAC, _
                                  "modDokumenta.GetVerwaistePrijemnice")
    colKol = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA, _
                                "modDokumenta.GetVerwaistePrijemnice")
    colStorno = RequireColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO, _
                                   "modDokumenta.GetVerwaistePrijemnice")

    Dim count As Long
    Dim i As Long

    For i = 1 To UBound(data, 1)
        If Trim$(NzToText(data(i, colStorno))) = "Da" Then GoTo NextCount

        If storniraneBrojevi.Exists(Trim$(NzToText(data(i, colBrZbr)))) Then
            count = count + 1
        End If

NextCount:
    Next i

    If count = 0 Then
        GetVerwaistePrijemnice = Empty
        Exit Function
    End If

    Dim result() As Variant
    ReDim result(1 To count, 1 To 5)

    Dim idx As Long
    Dim kupacNaziv As String
    Dim kol As Double

    For i = 1 To UBound(data, 1)
        If Trim$(NzToText(data(i, colStorno))) = "Da" Then GoTo NextRow

        If storniraneBrojevi.Exists(Trim$(NzToText(data(i, colBrZbr)))) Then
            idx = idx + 1

            kupacNaziv = CStr(LookupValue(TBL_KUPCI, COL_KUP_ID, _
                                          NzToText(data(i, colKupac)), COL_KUP_NAZIV))

            kol = 0
            If IsNumeric(data(i, colKol)) Then kol = CDbl(data(i, colKol))

            result(idx, 1) = NzToText(data(i, colID))
            result(idx, 2) = NzToText(data(i, colBrPrij))
            result(idx, 3) = NzToText(data(i, colBrZbr))
            result(idx, 4) = kupacNaziv
            result(idx, 5) = kol
        End If

NextRow:
    Next i

    GetVerwaistePrijemnice = result
    Exit Function

EH:
    LogErr "modDokumenta.GetVerwaistePrijemnice"
    GetVerwaistePrijemnice = Empty
End Function

Public Sub RelinkFakturaStavke(ByVal newPrijemnicaID As String, _
                               ByVal brojPrijemnice As String, _
                               Optional ByVal klasaFilter As String = "")
    ' Sucht verwaiste FakturaStavke die auf eine stornierte Prijemnica
    ' mit gleichem BrojPrijemnice zeigen, und verlinkt sie auf die neue.
    On Error GoTo EH

    If Trim$(newPrijemnicaID) = "" Or Trim$(brojPrijemnice) = "" Then Exit Sub

    Dim stavkeData As Variant
    stavkeData = GetTableData(TBL_FAKTURA_STAVKE)

    If IsEmpty(stavkeData) Then Exit Sub

    Dim colPrijID As Long
    Dim colOsir As Long
    Dim colFakID As Long
    Dim colFsKlasa As Long

    colPrijID = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_PRIJEMNICA_ID, _
                                   "modDokumenta.RelinkFakturaStavke")
    colOsir = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_OSIROCENO_OD, _
                                 "modDokumenta.RelinkFakturaStavke")
    colFakID = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_FAKTURA_ID, _
                                  "modDokumenta.RelinkFakturaStavke")
    colFsKlasa = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_KLASA, _
                                "modDokumenta.RelinkFakturaStavke")

    Dim i As Long
    Dim oldPrijID As String
    Dim oldBroj As String
    Dim fakID As String

    For i = 1 To UBound(stavkeData, 1)
    
        Dim stavkaKlasa As String
        stavkaKlasa = Trim$(NzToText(stavkeData(i, colFsKlasa)))
        If Len(Trim$(klasaFilter)) > 0 Then
            If UCase$(stavkaKlasa) <> UCase$(Trim$(klasaFilter)) Then GoTo NextStavka
        End If

        If Trim$(NzToText(stavkeData(i, colOsir))) = "" Then GoTo NextStavka

        oldPrijID = Trim$(NzToText(stavkeData(i, colPrijID)))

        If oldPrijID = "" Then GoTo NextStavka

        oldBroj = CStr(LookupValue(TBL_PRIJEMNICA, COL_PRJ_ID, oldPrijID, COL_PRJ_BROJ))

        If oldBroj = brojPrijemnice Then
            fakID = Trim$(NzToText(stavkeData(i, colFakID)))
            
            If Len(Trim$(fakID)) = 0 Then
            Err.Raise vbObjectError + 7413, "modDokumenta.RelinkFakturaStavke", _
                    "FakturaStavka nema FakturaID za relink. BrojPrijemnice=" & _
                    brojPrijemnice & "; Klasa=" & stavkaKlasa
            End If

            RequireUpdateCell TBL_FAKTURA_STAVKE, i, COL_FS_PRIJEMNICA_ID, _
                              newPrijemnicaID, "modDokumenta.RelinkFakturaStavke"

            RequireUpdateCell TBL_FAKTURA_STAVKE, i, COL_OSIROCENO_OD, _
                              "", "modDokumenta.RelinkFakturaStavke"

            Dim newPrijRow As Long
            newPrijRow = FindPrijemnicaRowByIDAndKlasa(newPrijemnicaID, stavkaKlasa, _
                                           "modDokumenta.RelinkFakturaStavke")

            If newPrijRow <= 0 Then
                Err.Raise vbObjectError + 7410, "modDokumenta.RelinkFakturaStavke", _
                        "Nova prijemnica nije pronadena za relink. PrijemnicaID=" & _
                        newPrijemnicaID & "; Klasa=" & stavkaKlasa
            End If

            RequireUpdateCell TBL_PRIJEMNICA, newPrijRow, COL_PRJ_FAKTURISANO, _
                            "Da", "modDokumenta.RelinkFakturaStavke"

            RequireUpdateCell TBL_PRIJEMNICA, newPrijRow, COL_PRJ_FAKTURA_ID, _
                            fakID, "modDokumenta.RelinkFakturaStavke"


            If Len(Trim$(fakID)) > 0 Then
                UpdateFakturaStatus fakID
            End If

        End If

NextStavka:
    Next i

    Exit Sub

EH:
    LogErr "modDokumenta.RelinkFakturaStavke"
    Err.Raise Err.Number, "modDokumenta.RelinkFakturaStavke", Err.description
End Sub

' ============================================================
' HELPER - Vozac-Report (ersetzt alten modTransport)
' ============================================================

Public Function GetVozacDokumenta(ByVal vozacID As String, _
                                   Optional ByVal datumOd As Date = 0, _
                                   Optional ByVal datumDo As Date = 0) As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_OTPREMNICA)

    If IsEmpty(data) Then
        GetVozacDokumenta = Empty
        Exit Function
    End If

    data = ExcludeStornirano(data, TBL_OTPREMNICA)

    If IsEmpty(data) Then
        GetVozacDokumenta = Empty
        Exit Function
    End If

    Dim filters As New Collection
    Dim fp As clsFilterParam

    Set fp = New clsFilterParam
    fp.Init RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VOZAC, _
            "modDokumenta.GetVozacDokumenta"), "=", vozacID
    filters.Add fp

    If datumOd > 0 And datumDo > 0 Then
        Set fp = New clsFilterParam
        fp.Init RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_DATUM, _
                "modDokumenta.GetVozacDokumenta"), "BETWEEN", datumOd, datumDo
        filters.Add fp
    End If

    GetVozacDokumenta = FilterArray(data, filters)
    Exit Function

EH:
    LogErr "modDokumenta.GetVozacDokumenta"
    GetVozacDokumenta = Empty
End Function

Public Function BuildZbirnaVrstaCache() As Object
    On Error GoTo EH

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim otpData As Variant
    otpData = GetTableData(TBL_OTPREMNICA)

    If IsEmpty(otpData) Then
        Set BuildZbirnaVrstaCache = dict
        Exit Function
    End If
    
    otpData = ExcludeStornirano(otpData, TBL_OTPREMNICA)

    If IsEmpty(otpData) Then
        Set BuildZbirnaVrstaCache = dict
        Exit Function
    End If

    Dim colBrZbr As Long
    Dim colVrsta As Long

    colBrZbr = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, _
                                  "modDokumenta.BuildZbirnaVrstaCache")
    colVrsta = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VRSTA, _
                                  "modDokumenta.BuildZbirnaVrstaCache")

    Dim i As Long
    Dim brz As String
    Dim vrsta As String

    For i = 1 To UBound(otpData, 1)
        brz = Trim$(NzToText(otpData(i, colBrZbr)))
        vrsta = Trim$(NzToText(otpData(i, colVrsta)))

        If brz <> "" Then
            If Not dict.Exists(brz) Then
                dict.Add brz, vrsta
            End If
        End If
    Next i

    Set BuildZbirnaVrstaCache = dict
    Exit Function

EH:
    LogErr "modDokumenta.BuildZbirnaVrstaCache"

    Dim emptyDict As Object
    Set emptyDict = CreateObject("Scripting.Dictionary")
    Set BuildZbirnaVrstaCache = emptyDict
End Function

Public Function GetVrstaFromCache(ByVal dict As Object, _
                                  ByVal brojZbirne As String) As String
    On Error GoTo EH

    If dict Is Nothing Then
        GetVrstaFromCache = ""
    ElseIf dict.Exists(brojZbirne) Then
        GetVrstaFromCache = CStr(dict(brojZbirne))
    Else
        GetVrstaFromCache = ""
    End If

    Exit Function

EH:
    LogErr "modDokumenta.GetVrstaFromCache"
    GetVrstaFromCache = ""
End Function

Private Sub PrintTxFailure(ByVal sourceName As String, _
                           ByVal errSrc As String, _
                           ByVal errNum As Long, _
                           ByVal errDesc As String)
    Debug.Print sourceName & " failed. Source=" & errSrc & _
                " Err=" & CStr(errNum) & _
                " Desc=" & errDesc
End Sub

Private Sub RequireValidDocumentClass(ByVal klasa As String, _
                                      ByVal sourceName As String)

    Select Case Trim$(CStr(klasa))
        Case KLASA_I, KLASA_II
            Exit Sub
    End Select

    Err.Raise vbObjectError + 1400, sourceName, _
              "Neispravna klasa dokumenta: " & klasa
End Sub

Private Sub ValidatePrijemnicaInput(ByVal kupacID As String, _
                                    ByVal vozacID As String, _
                                    ByVal brojPrij As String, _
                                    ByVal brojZbirne As String, _
                                    ByVal kolicina As Double, _
                                    ByVal cena As Double, _
                                    ByVal tipAmb As String, _
                                    ByVal kolAmb As Long, _
                                    ByVal kolAmbVracena As Long, _
                                    ByVal klasa As String)

    Const SRC As String = "ValidatePrijemnicaInput"

    If Len(Trim$(kupacID)) = 0 Then
        Err.Raise vbObjectError + 1420, SRC, "KupacID je obavezan."
    End If

    If Len(Trim$(vozacID)) = 0 Then
        Err.Raise vbObjectError + 1421, SRC, "VozacID je obavezan."
    End If

    If Len(Trim$(brojPrij)) = 0 Then
        Err.Raise vbObjectError + 1422, SRC, "Broj prijemnice je obavezan."
    End If

    If Len(Trim$(brojZbirne)) = 0 Then
        Err.Raise vbObjectError + 1423, SRC, "Broj zbirne je obavezan."
    End If

    ' Referencijalni integritet: broj zbirne mora da postoji u tblZbirna.
    ' Samo u BLOK modu -- u UPOZORENJE modu forma trazi potvrdu, pa backend ne
    ' sme tvrdo da padne. Ujedno je bezbedna mreza i za ne-form pozivaoce.
    If PrijemnicaZbirnaBlokira() Then
        If Not ZbirnaPostoji(brojZbirne) Then
            Err.Raise vbObjectError + 1428, SRC, _
                "Zbirna '" & brojZbirne & "' ne postoji u sistemu."
        End If
    End If

    If kolicina <= 0 Then
        Err.Raise vbObjectError + 1424, SRC, "Koli" & ChrW(269) & "ina mora biti veca od nule."
    End If

    If cena < 0 Then
        Err.Raise vbObjectError + 1425, SRC, "Cena ne sme biti negativna."
    End If

    If kolAmb < 0 Or kolAmbVracena < 0 Then
        Err.Raise vbObjectError + 1426, SRC, "Koli" & ChrW(269) & "ina ambala" & ChrW(382) & "e ne sme biti negativna."
    End If

    If (kolAmb > 0 Or kolAmbVracena > 0) And Len(Trim$(tipAmb)) = 0 Then
        Err.Raise vbObjectError + 1427, SRC, "Tip ambala" & ChrW(382) & "e je obavezan kada postoji ambala" & ChrW(382) & "a."
    End If

    RequireValidDocumentClass klasa, SRC
End Sub

Private Function FindPrijemnicaRowByIDAndKlasa(ByVal prijemnicaID As String, _
                                               ByVal klasa As String, _
                                               ByVal sourceName As String) As Long
    Dim data As Variant
    data = GetTableData(TBL_PRIJEMNICA)

    If IsEmpty(data) Then Exit Function

    Dim colID As Long
    Dim colKlasa As Long

    colID = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID, sourceName)
    colKlasa = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KLASA, sourceName)

    Dim i As Long
    Dim foundRow As Long
    Dim foundCount As Long

    For i = 1 To UBound(data, 1)
        If Trim$(NzToText(data(i, colID))) = Trim$(prijemnicaID) And _
           Trim$(NzToText(data(i, colKlasa))) = Trim$(klasa) Then

            foundRow = i
            foundCount = foundCount + 1
        End If
    Next i

    If foundCount > 1 Then
        Err.Raise vbObjectError + 7412, sourceName, _
                  "Dupli redovi za PrijemnicaID + Klasa. PrijemnicaID=" & _
                  prijemnicaID & "; Klasa=" & klasa
    End If

    FindPrijemnicaRowByIDAndKlasa = foundRow
End Function

' ============================================================
' STORNO PREGLED (read-only) -- agregira stornirane dokumente po tipu za
' prikaz u panelu unutar frmDokumenta (dugme "Pregled storniranih").
' Soft-delete: red je storniran kad je COL_STORNIRANO = "Da" (modStorno).
' Jedinstven (unifikovan) skup korisnih kolona za sve tipove:
'   Broj | Datum | Partner | Vrsta | Sorta | Klasa | Kolicina | Cena | Iznos
'   | Zbirna | Otpremnica | Faktura  (poslednje 3 = lanac zavisnih dokumenata)
' Partner se razresava na naziv/ime (best-effort; fallback = ID).
' ============================================================

' Tipovi storno dokumenata u redosledu prikaza (isti nazivi kao cmbStornoDokument).
Public Function StorniraniTipovi() As Variant
    StorniraniTipovi = Array("Otkup", "Otpremnica", "Zbirna", "Prijemnica", "Faktura", "Novac", _
                         "Revers izdavanje koop.", "Revers povrat koop.", _
                         "Revers izdato OM (firma).", "Revers prijem od OM (firma).")
End Function

' Zaglavlja unifikovanih kolona (0-bazni niz, 12 kolona).
Public Function StorniraniHeaders() As Variant
    StorniraniHeaders = Array("Broj", "Datum", "Partner", "Vrsta", "Sorta", _
                              "Klasa", "Koli" & ChrW(269) & "ina", "Cena", "Iznos (RSD)", _
                              "Zbirna", "Otpremnica", "Faktura")
End Function

' Stornirani OM<->koop revers (dokTip) iz tblAmbalaza -> unifikovane 12 kolona.
' Jedan red po dokumentu: Kooperant noga (nosi partnera); Stanica noga se preskace.
' Mapiranje: Vrsta=TipAmbalaze, Kolicina=broj gajbica; ostalo prazno.
' Preskace OM-Izlaz-Koop noge knjizene UZ otkup (DokumentID = otkupID "OTK-...";
' vec se vide pod grupom 'Otkup') -> ostaju samo standalone reversi (broj x/ddmmyy).
Private Function GetStorniraniRevers(ByVal dokTip As String, _
                                     Optional ByVal entType As String = "Kooperant") As Variant
    On Error GoTo EH
    Dim data As Variant: data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then Exit Function

    Dim iBroj As Long, iDat As Long, iEntID As Long, iEntTip As Long
    Dim iDokTip As Long, iTipAmb As Long, iKol As Long, iStorno As Long
    iBroj = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_ID)
    iDat = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DATUM)
    iEntID = GetColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET)
    iEntTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET_TIP)
    iDokTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_TIP)
    iTipAmb = GetColumnIndex(TBL_AMBALAZA, COL_AMB_TIP)
    iKol = GetColumnIndex(TBL_AMBALAZA, COL_AMB_KOLICINA)
    iStorno = GetColumnIndex(TBL_AMBALAZA, COL_STORNIRANO)
    If iBroj = 0 Or iDokTip = 0 Or iStorno = 0 Then Exit Function

    Dim pdict As Object
    If entType = "Stanica" Then
        Set pdict = BuildIdNameDict(TBL_STANICE, "StanicaID", "Naziv")
    Else
        Set pdict = BuildIdNameDict(TBL_KOOPERANTI, COL_KOOP_ID, "Ime", "Prezime")
    End If
    Dim rows As Collection: Set rows = New Collection
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If UCase$(Trim$(NzToText(data(i, iStorno)))) = "DA" Then
            If Trim$(CStr(data(i, iDokTip))) = dokTip And _
               Trim$(CStr(data(i, iEntTip))) = entType And _
               UCase$(Left$(Trim$(CStr(data(i, iBroj))), 3)) <> "OTK" Then
                rows.Add Array( _
                    StornoCellText(data, i, iBroj), _
                    StornoDateText(StornoCellRaw(data, i, iDat)), _
                    ResolveNameFromDict(pdict, StornoCellRaw(data, i, iEntID)), _
                    StornoCellText(data, i, iTipAmb), _
                    "", "", _
                    StornoNumText(StornoCellRaw(data, i, iKol), "#,##0"), _
                    "", "", "", "", "")
            End If
        End If
    Next i
    GetStorniraniRevers = StornoRowsTo2D(rows, 12)
    Exit Function
EH:
    LogErr "modDokumenta.GetStorniraniRevers(" & dokTip & ")"
    GetStorniraniRevers = Empty
End Function

' id -> "Naziv" (ili "Ime Prezime") recnik; prazan recnik ako tabela/kolone fale.
' Public: reuse i iz modIntegritet (OtkupnoMestoByZbirna).
Public Function BuildIdNameDict(ByVal tbl As String, ByVal idCol As String, _
                                ByVal nameCol1 As String, _
                                Optional ByVal nameCol2 As String = "") As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set BuildIdNameDict = d

    Dim data As Variant
    data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Function

    Dim ci As Long, c1 As Long, c2 As Long
    ci = GetColumnIndex(tbl, idCol)
    c1 = GetColumnIndex(tbl, nameCol1)
    If Len(nameCol2) > 0 Then c2 = GetColumnIndex(tbl, nameCol2)
    If ci = 0 Or c1 = 0 Then Exit Function

    Dim i As Long, idv As String, nm As String
    For i = 1 To UBound(data, 1)
        idv = Trim$(NzToText(data(i, ci)))
        If Len(idv) > 0 Then
            nm = Trim$(NzToText(data(i, c1)))
            If c2 > 0 Then nm = Trim$(nm & " " & NzToText(data(i, c2)))
            If Not d.Exists(idv) Then d.Add idv, nm
        End If
    Next i
End Function

Private Function ResolveNameFromDict(ByVal d As Object, ByVal idRaw As Variant) As String
    Dim s As String
    s = Trim$(NzToText(idRaw))
    If Len(s) = 0 Then Exit Function
    If Not d Is Nothing Then
        If d.Exists(s) Then
            If Len(Trim$(NzToText(d(s)))) > 0 Then
                ResolveNameFromDict = d(s)
                Exit Function
            End If
        End If
    End If
    ResolveNameFromDict = s          ' fallback: prikazi ID ako nema imena
End Function

' Sirova vrednost celije (Empty ako kolona ne postoji -> idx=0).
' `data` je ByRef namerno. ByVal na Variantu koji SADRZI niz kopira ceo niz pri
' svakom pozivu, a ovo je citac PO CELIJI -- u petlji se zove vise puta po redu.
' Mereno na istom obrascu u modPaletniList.SafeCell: 1063 stavke, 1883 ms, to jest
' 1.8 ms po redu za citanje dva polja iz niza koji je vec u memoriji.
'
' Funkcija niz samo CITA, nikad ne pise, pa je razlika iskljucivo u tome sto se
' niz ne umnozava.
Private Function StornoCellRaw(ByRef data As Variant, ByVal r As Long, ByVal idx As Long) As Variant
    If idx = 0 Then Exit Function
    StornoCellRaw = data(r, idx)
End Function

' ByRef iz istog razloga kao StornoCellRaw: ByVal bi kopirao ceo niz po pozivu.
Private Function StornoCellText(ByRef data As Variant, ByVal r As Long, ByVal idx As Long) As String
    If idx = 0 Then Exit Function
    StornoCellText = Trim$(NzToText(data(r, idx)))
End Function

Private Function StornoDateText(ByVal v As Variant) As String
    If IsDate(v) Then
        StornoDateText = Format$(CDate(v), "d.m.yyyy")
    Else
        StornoDateText = Trim$(NzToText(v))
    End If
End Function

Private Function StornoNumText(ByVal v As Variant, ByVal fmt As String) As String
    Dim d As Double
    If TryParseDouble(Trim$(NzToText(v)), d) Then
        If d <> 0 Then StornoNumText = Format$(d, fmt)
    End If
End Function

' Iznos = v1 - v2 (Novac: Uplata - Isplata; ostali: v2 prazno -> samo v1).
Private Function StornoIznosText(ByVal v1 As Variant, ByVal v2 As Variant) As String
    Dim u As Double, s As Double
    If Not TryParseDouble(Trim$(NzToText(v1)), u) Then u = 0
    If Not TryParseDouble(Trim$(NzToText(v2)), s) Then s = 0
    Dim net As Double: net = u - s
    If net <> 0 Then StornoIznosText = Format$(net, "#,##0.00")
End Function

' Iznos = Kolicina x Cena (prazno ako je proizvod 0).
Private Function StornoMnozi(ByVal vKol As Variant, ByVal vCena As Variant) As String
    Dim kol As Double, cena As Double
    If Not TryParseDouble(Trim$(NzToText(vKol)), kol) Then kol = 0
    If Not TryParseDouble(Trim$(NzToText(vCena)), cena) Then cena = 0
    Dim p As Double: p = kol * cena
    If p <> 0 Then StornoMnozi = Format$(p, "#,##0.00")
End Function

' --- Indeks lanca dokumenata (reverzni lookup-i preko BrojZbirne / FakturaID) ---

' Prazan recnik sa case-insensitive poredjenjem kljuceva.
Private Function NewDict() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set NewDict = d
End Function

Private Function DictGet(ByVal d As Object, ByVal key As String) As String
    If d Is Nothing Then Exit Function
    If Len(key) = 0 Then Exit Function
    If d.Exists(key) Then DictGet = CStr(d(key))
End Function

' Dodaj val u listu pod key (", "-spojeno, bez duplikata).
Private Sub DictAppend(ByVal d As Object, ByVal key As String, ByVal val As String)
    If d Is Nothing Then Exit Sub
    If Len(key) = 0 Or Len(val) = 0 Then Exit Sub
    If Not d.Exists(key) Then
        d.Add key, val
    Else
        Dim cur As String: cur = CStr(d(key))
        If InStr(1, ", " & cur & ", ", ", " & val & ", ", vbTextCompare) = 0 Then
            d(key) = cur & ", " & val
        End If
    End If
End Sub

' Izgradi indeks lanca: otpByZbr, fakByZbr, fakById, zbrByFak (sve case-insensitive).
Private Function BuildChainIndex() As Object
    Dim idx As Object: Set idx = CreateObject("Scripting.Dictionary")
    Set BuildChainIndex = idx

    Dim fakById As Object:  Set fakById = NewDict()
    Dim otpByZbr As Object: Set otpByZbr = NewDict()
    Dim fakByZbr As Object: Set fakByZbr = NewDict()
    Dim zbrByFak As Object: Set zbrByFak = NewDict()
    idx.Add "fakById", fakById
    idx.Add "otpByZbr", otpByZbr
    idx.Add "fakByZbr", fakByZbr
    idx.Add "zbrByFak", zbrByFak

    Dim d As Variant, i As Long, k As String, v As String

    ' Faktura: FakturaID -> BrojFakture
    d = GetTableData(TBL_FAKTURE)
    If Not IsEmpty(d) Then
        Dim cfi As Long, cfb As Long
        cfi = GetColumnIndex(TBL_FAKTURE, COL_FAK_ID)
        cfb = GetColumnIndex(TBL_FAKTURE, COL_FAK_BROJ)
        If cfi > 0 And cfb > 0 Then
            For i = 1 To UBound(d, 1)
                k = Trim$(NzToText(d(i, cfi))): v = Trim$(NzToText(d(i, cfb)))
                If Len(k) > 0 And Not fakById.Exists(k) Then fakById.Add k, v
            Next i
        End If
    End If

    ' Otpremnica: BrojZbirne -> lista BrojOtpremnice
    d = GetTableData(TBL_OTPREMNICA)
    If Not IsEmpty(d) Then
        Dim coz As Long, cob As Long
        coz = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE)
        cob = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ)
        If coz > 0 And cob > 0 Then
            For i = 1 To UBound(d, 1)
                DictAppend otpByZbr, Trim$(NzToText(d(i, coz))), Trim$(NzToText(d(i, cob)))
            Next i
        End If
    End If

    ' Prijemnica: BrojZbirne <-> FakturaID  (i BrojZbirne -> BrojFakture preko fakById)
    d = GetTableData(TBL_PRIJEMNICA)
    If Not IsEmpty(d) Then
        Dim cpz As Long, cpf As Long
        cpz = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
        cpf = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_FAKTURA_ID)
        If cpz > 0 And cpf > 0 Then
            For i = 1 To UBound(d, 1)
                Dim z As String, fid As String
                z = Trim$(NzToText(d(i, cpz))): fid = Trim$(NzToText(d(i, cpf)))
                If Len(fid) > 0 Then
                    DictAppend zbrByFak, fid, z
                    If fakById.Exists(fid) Then DictAppend fakByZbr, z, CStr(fakById(fid))
                End If
            Next i
        End If
    End If
End Function

' Kolekcija 0-baznih 1D nizova -> 2D (1..n, 1..ncol). Empty ako je prazna.
Private Function StornoRowsTo2D(ByVal rows As Collection, ByVal ncol As Long) As Variant
    If rows Is Nothing Then Exit Function
    If rows.count = 0 Then Exit Function

    Dim arr() As Variant
    ReDim arr(1 To rows.count, 1 To ncol)
    Dim r As Long, c As Long, one As Variant
    For r = 1 To rows.count
        one = rows(r)
        For c = 1 To ncol
            arr(r, c) = one(c - 1)
        Next c
    Next r
    StornoRowsTo2D = arr
End Function

' Osirocene prijemnice: AKTIVNE prijemnice cija zbirna (BrojZbirne) vise nije
' aktivna (stornirana ili ne postoji). Jedan red po BrojPrijemnice (Klasa I+II
' dele broj). Za re-point UI (frmDokumenta recovery panel).
' Kolone 1..7: BrojPrijemnice|Datum|Vrsta|Sorta|Kolicina|StaraZbirna|Status
Public Function GetOsirocenePrijemnice() As Variant
    On Error GoTo EH

    Dim prj As Variant: prj = GetTableData(TBL_PRIJEMNICA)
    If IsEmpty(prj) Then Exit Function

    ' BrojZbirne -> ima li bar jednu AKTIVNU zbirnu; i postoji li uopste.
    Dim aktZbr As Object: Set aktZbr = CreateObject("Scripting.Dictionary")
    Dim allZbr As Object: Set allZbr = CreateObject("Scripting.Dictionary")
    aktZbr.CompareMode = vbTextCompare: allZbr.CompareMode = vbTextCompare
    Dim zd As Variant: zd = GetTableData(TBL_ZBIRNA)
    If Not IsEmpty(zd) Then
        Dim zBr As Long, zSt As Long, zr As Long, zk As String
        zBr = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ)
        zSt = GetColumnIndex(TBL_ZBIRNA, COL_STORNIRANO)
        For zr = 1 To UBound(zd, 1)
            zk = Trim$(NzToText(zd(zr, zBr)))
            If Len(zk) > 0 Then
                allZbr(zk) = True
                If UCase$(Trim$(NzToText(zd(zr, zSt)))) <> "DA" Then aktZbr(zk) = True
            End If
        Next zr
    End If

    Dim cBr As Long, cDat As Long, cVr As Long, cSo As Long
    Dim cKol As Long, cZbr As Long, cSt As Long, cGen As Long
    cGen = GetColumnIndex(TBL_PRIJEMNICA, COL_GENERACIJA_ID)
    cBr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ)
    cDat = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_DATUM)
    cVr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_VRSTA)
    cSo = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_SORTA)
    cKol = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA)
    cZbr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
    cSt = GetColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO)
    If cBr = 0 Or cZbr = 0 Then Exit Function

    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    Dim rows As Collection: Set rows = New Collection
    Dim i As Long
    For i = 1 To UBound(prj, 1)
        Dim isStor As Boolean: isStor = False
        If cSt > 0 Then isStor = (UCase$(Trim$(NzToText(prj(i, cSt)))) = "DA")
        If Not isStor Then
            Dim bz As String: bz = Trim$(NzToText(prj(i, cZbr)))
            Dim brp As String: brp = Trim$(NzToText(prj(i, cBr)))
            If Len(bz) > 0 And Len(brp) > 0 Then
                If Not aktZbr.Exists(bz) Then
                    ' Grupise se po BROJ + GENERACIJA, ne po samom broju: dva
                    ' kupca istog dana dele "1/ddmmyy", pa bi dedup po broju
                    ' spojio DVA dokumenta u jedan red - i recovery bi posle
                    ' prevezao onaj koji zatekne.
                    Dim gen As String: gen = Trim$(NzToText(prj(i, cGen)))
                    Dim kljuc As String: kljuc = brp & Chr$(1) & gen
                    If Not seen.Exists(kljuc) Then
                        seen(kljuc) = True
                        Dim st As String
                        If allZbr.Exists(bz) Then st = "zbirna stornirana" Else st = "zbirna ne postoji"
                        ' 8. kolona je GENERACIJA - dodata na KRAJ, pa legacy
                        ' panel (cita fiksno 1..7) ostaje netaknut.
                        rows.Add Array(brp, prj(i, cDat), Trim$(NzToText(prj(i, cVr))), _
                            Trim$(NzToText(prj(i, cSo))), StornoNumText(prj(i, cKol), "#,##0.00"), bz, st, gen)
                    End If
                End If
            End If
        End If
    Next i

    GetOsirocenePrijemnice = StornoRowsTo2D(rows, 8)
    Exit Function
EH:
    LogErr "modDokumenta.GetOsirocenePrijemnice"
    GetOsirocenePrijemnice = Empty
End Function

' Stornirane prijemnice koje imaju AKTIVNE (osirocene) paleta-stavke. Za P1 pallet
' re-point. Kolone 1..6: BrojPrijemnice|Datum|Vrsta|Sorta|Gajbica|StavkiAktivnih
Public Function GetPrijemniceSaOsirocenimPaletama() As Variant
    On Error GoTo EH
    Dim ps As Variant: ps = GetTableData(TBL_PALETA_STAVKA)
    If IsEmpty(ps) Then Exit Function

    Dim stPrij As Object: Set stPrij = CreateObject("Scripting.Dictionary"): stPrij.CompareMode = vbTextCompare
    Dim prj As Variant: prj = GetTableData(TBL_PRIJEMNICA)
    If Not IsEmpty(prj) Then
        Dim qId As Long, qSt As Long, q As Long
        qId = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID)
        qSt = GetColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO)
        For q = 1 To UBound(prj, 1)
            If qSt > 0 Then
                If UCase$(Trim$(NzToText(prj(q, qSt)))) = "DA" Then stPrij(Trim$(NzToText(prj(q, qId)))) = True
            End If
        Next q
    End If

    Dim cBr As Long, cPid As Long, cVr As Long, cSo As Long, cGajb As Long, cSt As Long
    cBr = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_PRIJ)
    cPid = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PRIJEMNICA_ID)
    cVr = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_VRSTA)
    cSo = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_SORTA)
    cGajb = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BR_GAJBICA)
    cSt = GetColumnIndex(TBL_PALETA_STAVKA, COL_STORNIRANO)
    If cBr = 0 Or cPid = 0 Then Exit Function

    ' Grupise se po GENERACIJI storniranog dokumenta, ne po njegovom broju.
    ' Dve stornirane prijemnice razlicitih kupaca lako dele broj (racuna se po
    ' kupcu), pa bi grupisanje po broju spojilo DVA dokumenta u jedan recovery
    ' red - a prevezivanje bi posle pomerilo palete oba.
    ' Generacija se cita sa PRIJEMNICE, preko PrijemnicaID koji stavka nosi.
    Dim gSum As Object: Set gSum = CreateObject("Scripting.Dictionary"): gSum.CompareMode = vbTextCompare
    Dim sCnt As Object: Set sCnt = CreateObject("Scripting.Dictionary"): sCnt.CompareMode = vbTextCompare
    Dim vrD As Object: Set vrD = CreateObject("Scripting.Dictionary"): vrD.CompareMode = vbTextCompare
    Dim soD As Object: Set soD = CreateObject("Scripting.Dictionary"): soD.CompareMode = vbTextCompare
    Dim brD As Object: Set brD = CreateObject("Scripting.Dictionary"): brD.CompareMode = vbTextCompare
    Dim genD As Object: Set genD = CreateObject("Scripting.Dictionary"): genD.CompareMode = vbTextCompare
    Dim order As Collection: Set order = New Collection
    Dim i As Long
    For i = 1 To UBound(ps, 1)
        Dim stx As Boolean: stx = False
        If cSt > 0 Then stx = (UCase$(Trim$(NzToText(ps(i, cSt)))) = "DA")
        If Not stx Then
            Dim pid As String: pid = Trim$(NzToText(ps(i, cPid)))
            Dim br As String: br = Trim$(NzToText(ps(i, cBr)))
            If stPrij.Exists(pid) And Len(br) > 0 Then
                Dim gen As String: gen = GeneracijaPoID(TBL_PRIJEMNICA, COL_PRJ_ID, pid)
                Dim kl As String: kl = br & Chr$(1) & gen
                If Not gSum.Exists(kl) Then
                    gSum(kl) = 0&: sCnt(kl) = 0&: order.Add kl
                    vrD(kl) = Trim$(NzToText(ps(i, cVr))): soD(kl) = Trim$(NzToText(ps(i, cSo)))
                    brD(kl) = br: genD(kl) = gen
                End If
                If IsNumeric(ps(i, cGajb)) Then gSum(kl) = CLng(gSum(kl)) + CLng(ps(i, cGajb))
                sCnt(kl) = CLng(sCnt(kl)) + 1
            End If
        End If
    Next i

    Dim rows As Collection: Set rows = New Collection
    Dim v As Variant
    For Each v In order
        Dim k2 As String: k2 = CStr(v)
        ' 7. kolona je GENERACIJA - na KRAJ, pa legacy panel (cita 1..6) ostaje.
        rows.Add Array(CStr(brD(k2)), _
                       LookupValue(TBL_PRIJEMNICA, COL_PRJ_BROJ, CStr(brD(k2)), COL_PRJ_DATUM), _
                       CStr(vrD(k2)), CStr(soD(k2)), CLng(gSum(k2)), CLng(sCnt(k2)), CStr(genD(k2)))
    Next v
    GetPrijemniceSaOsirocenimPaletama = StornoRowsTo2D(rows, 7)
    Exit Function
EH:
    LogErr "modDokumenta.GetPrijemniceSaOsirocenimPaletama"
    GetPrijemniceSaOsirocenimPaletama = Empty
End Function

' Aktivne (ne-stornirane) prijemnice = ciljevi za pallet re-point. Ukljucuje i
' paletizovane i NEpaletizovane (motor ReassignPaleteToPrijemnica_TX podrzava cilj
' bez svezih stavki: STEP1 tada ne undo-uje nista). Kolone 1..7:
'   BrojPrijemnice|Datum|Vrsta|Sorta|Gajbica(KolAmb)|Zbirna|Paletizovana(Da/Ne)
Public Function GetAktivnePrijemnice() As Variant
    On Error GoTo EH
    Dim prj As Variant: prj = GetTableData(TBL_PRIJEMNICA)
    If IsEmpty(prj) Then Exit Function

    ' BrojPrijemnice koje imaju bar jednu AKTIVNU paleta-stavku (-> Paletizovana=Da).
    Dim palBr As Object: Set palBr = CreateObject("Scripting.Dictionary"): palBr.CompareMode = vbTextCompare
    Dim ps As Variant: ps = GetTableData(TBL_PALETA_STAVKA)
    If Not IsEmpty(ps) Then
        Dim xBr As Long, xSt As Long, x As Long
        xBr = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_PRIJ)
        xSt = GetColumnIndex(TBL_PALETA_STAVKA, COL_STORNIRANO)
        For x = 1 To UBound(ps, 1)
            Dim sx As Boolean: sx = False
            If xSt > 0 Then sx = (UCase$(Trim$(NzToText(ps(x, xSt)))) = "DA")
            If Not sx Then palBr(Trim$(NzToText(ps(x, xBr)))) = True
        Next x
    End If

    Dim cBr As Long, cDat As Long, cVr As Long, cSo As Long, cAmb As Long, cZbr As Long, cSt As Long
    cBr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ)
    cDat = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_DATUM)
    cVr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_VRSTA)
    cSo = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_SORTA)
    cAmb = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOL_AMB)
    cZbr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
    cSt = GetColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO)
    If cBr = 0 Then Exit Function

    Dim gSum As Object: Set gSum = CreateObject("Scripting.Dictionary"): gSum.CompareMode = vbTextCompare
    Dim datD As Object: Set datD = CreateObject("Scripting.Dictionary"): datD.CompareMode = vbTextCompare
    Dim vrD As Object: Set vrD = CreateObject("Scripting.Dictionary"): vrD.CompareMode = vbTextCompare
    Dim soD As Object: Set soD = CreateObject("Scripting.Dictionary"): soD.CompareMode = vbTextCompare
    Dim zbD As Object: Set zbD = CreateObject("Scripting.Dictionary"): zbD.CompareMode = vbTextCompare
    Dim order As Collection: Set order = New Collection
    Dim i As Long
    For i = 1 To UBound(prj, 1)
        Dim stp As Boolean: stp = False
        If cSt > 0 Then stp = (UCase$(Trim$(NzToText(prj(i, cSt)))) = "DA")
        If Not stp Then
            Dim br As String: br = Trim$(NzToText(prj(i, cBr)))
            If Len(br) > 0 Then
                If Not gSum.Exists(br) Then
                    gSum(br) = 0&: order.Add br
                    datD(br) = prj(i, cDat): vrD(br) = Trim$(NzToText(prj(i, cVr)))
                    soD(br) = Trim$(NzToText(prj(i, cSo))): zbD(br) = Trim$(NzToText(prj(i, cZbr)))
                End If
                If cAmb > 0 Then If IsNumeric(prj(i, cAmb)) Then gSum(br) = CLng(gSum(br)) + CLng(prj(i, cAmb))
            End If
        End If
    Next i

    Dim rows As Collection: Set rows = New Collection
    Dim v As Variant
    For Each v In order
        Dim br2 As String: br2 = CStr(v)
        rows.Add Array(br2, datD(br2), CStr(vrD(br2)), CStr(soD(br2)), CLng(gSum(br2)), _
                       CStr(zbD(br2)), IIf(palBr.Exists(br2), "Da", "Ne"))
    Next v
    GetAktivnePrijemnice = StornoRowsTo2D(rows, 7)
    Exit Function
EH:
    LogErr "modDokumenta.GetAktivnePrijemnice"
    GetAktivnePrijemnice = Empty
End Function

' Re-point prijemnice na drugu (aktivnu) zbirnu po BrojZbirne. Koristi se posle
' storna otpremnice+zbirne (van autohladnjace): operater napravi novu otpremnicu+
' zbirnu pa "osirocenu" prijemnicu prevezuje na nju. Handle = BrojPrijemnice
' (citljiv, eksteran kad nije autohladnjaca); stvarna veza lanca = BrojZbirne.
' PrijemnicaID se NE dira -> faktura-stavke i paleta-stavke ostaju ispravne.
' prijemnicaGeneracijaID / zbirnaGeneracijaID: identitet prijemnice koja se
' prevezuje i zbirne na koju ide. Kad su poznati, bira se BAS taj dokument, a
' broj ostaje labela. Oba su opciona zbog zatecenih pozivalaca - bez njih vazi
' kapija nad jednoznacnoscu broja, na svakoj strani zasebno.
Public Function ReassignPrijemnicaToZbirna_TX(ByVal brPrijemnice As String, _
                                              ByVal targetBrZbirne As String, _
                                              Optional ByVal prijemnicaGeneracijaID As String = "", _
                                              Optional ByVal zbirnaGeneracijaID As String = "") As Boolean
    Const SRC As String = "modDokumenta.ReassignPrijemnicaToZbirna_TX"
    Dim tx As clsTransaction
    On Error GoTo EH

    brPrijemnice = Trim$(brPrijemnice)
    targetBrZbirne = Trim$(targetBrZbirne)
    If Len(brPrijemnice) = 0 Or Len(targetBrZbirne) = 0 Then Exit Function

    ' Cilj mora biti AKTIVNA zbirna (postoji + nije stornirana). ZbirnaID je uvek
    ' popunjen -> dokaz postojanja; Stornirano je blank za aktivnu pa se NE sme
    ' koristiti kao dokaz postojanja.
    ' Sa generacijom cilj se bira po IDENTITETU: LookupValue po broju uzima prvi
    ' pogodak. Broj zbirne JESTE jedinstven kad ga da generator (SuggestNextBroj
    ' za ZBR bumpuje sekvencu dok BrojZbirneExists ne kaze da je slobodan), ali
    ' ne i kad je unet rucno -- auto-broj se moze iskljuciti u Podesavanjima.
    Dim tgtIds As Object
    Set tgtIds = IdoviGeneracije(TBL_ZBIRNA, COL_ZBR_ID, zbirnaGeneracijaID)

    If Len(Trim$(zbirnaGeneracijaID)) > 0 And tgtIds.count = 0 Then Exit Function

    Dim tId As Variant
    If tgtIds.count > 0 Then
        tId = tgtIds.Keys()(0)
        If UCase$(Trim$(NzToText(LookupValue(TBL_ZBIRNA, COL_ZBR_ID, CStr(tId), _
                                             COL_STORNIRANO)))) = "DA" Then Exit Function
        ' Labela koja se upisuje mora da opisuje BAS izabrani dokument. Kad
        ' generacija odlucuje, broj se cita iz nje - ne veruje se onome sto je
        ' pozivalac poslao, jer bi neuskladjen par tiho upisao tudji broj.
        targetBrZbirne = Trim$(NzToText(LookupValue(TBL_ZBIRNA, COL_ZBR_ID, CStr(tId), _
                                                    COL_ZBR_BROJ)))
        If Len(targetBrZbirne) = 0 Then Exit Function
    Else
        ' CILJ SE NE RAZRESAVA PO REDU KOJI JE SLUCAJNO PRVI.
        '
        ' LookupValue po broju vraca PRVI pogodak i ne gleda storno. Posle
        ' storna jednog vlasnika prvi red sa tim brojem moze biti storniran
        ' dok pod istim brojem stoji AKTIVNA zbirna -- a tada je legitimno
        ' prevezivanje TIHO stajalo: bez poruke, bez loga, samo False.
        '
        ' ZbirnaPostoji gleda samo AKTIVNE redove, pa jedno tacno pitanje
        ' ("ima li aktivnog cilja pod ovim brojem") zamenjuje dva pogresna
        ' ("postoji li ijedan red" + "da li je PRVI storniran"). Test 125.
        '
        ' RAZRESAVA SE ISTIM POREDJENJEM KOJIM RADI KAPIJA ISPOD.
        '
        ' ZbirnaPostoji ovde NE valja iako pita pravu stvar: on poredi bez
        ' obzira na velicinu slova (StrComp vbTextCompare), a VlasniciPoBroju
        ' -- koji stoji iza kapije -- poredi TACNO. Sa "zb-test-kask" bi
        ' postojanje reklo DA, kapija bi videla NULA vlasnika (a ona hvata samo
        ' n > 1, pa bi propustila), i u tblPrijemnica i tblPaletaStavka bi se
        ' upisala labela POZIVAOCA umesto one iz tabele. Postojanje i
        ' vlasnistvo bi govorili o dve razlicite stvari.
        '
        ' Jedan poziv daje oba odgovora: broj AKTIVNIH vlasnika pod tim brojem.
        ' Nula znaci "nema aktivnog cilja", a vise od jedan hvata kapija ispod
        ' -- koja ostaje zbog svoje poruke. Test 125.
        Dim aktivniVlasnici As Long
        aktivniVlasnici = VlasniciPoBroju(TBL_ZBIRNA, COL_ZBR_BROJ, _
                                          targetBrZbirne, SRC, False, _
                                          Array(COL_ZBR_VOZAC, COL_ZBR_KUPAC)).count
        If aktivniVlasnici = 0 Then Exit Function

        ' Bez generacije CILJ MORA BITI JEDNOZNACAN, i to ISTORIJSKI.
        '
        ' Vlasnistvo zbirne je vozac + kupac, isti par koji koriste
        ' StornoZbirna i ApplyGeneracijaID. Broji se IKAD, ne samo aktivni:
        ' storniran vlasnik i dalje ima AKTIVNU decu, pa posle njegovog storna
        ' ostane jedan aktivan i broj IZGLEDA jednoznacan -- a nije.
        '
        ' Ovde je to obavezno, ne opciono: recovery panel u frmDokumenta zove
        ' ovu funkciju sa DVA argumenta (bez generacije) i BEZ ijedne spoljne
        ' kapije, pa se funkcija mora braniti sama. Putanje iz modStornoFlow
        ' imaju ZbirnaBrojJeDvosmislenIkad iznad sebe; ta nema nista.
        '
        ' Veza koja se upisuje je gola labela (COL_PRJ_BROJ_ZBIRNE,
        ' COL_PALS_BROJ_ZBIRNE), pa bi dete zavrsilo vezano za broj koji
        ' pripada dvama vlasnickim tokovima. Test 125.
        RequireJedanVlasnikIkadPoBroju TBL_ZBIRNA, COL_ZBR_BROJ, targetBrZbirne, _
                                       SRC, COL_ZBR_VOZAC, COL_ZBR_KUPAC
    End If

    Dim cBrPrij As Long, cBrZbr As Long, cStorno As Long
    cBrPrij = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ)
    cBrZbr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
    cStorno = GetColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO)
    If cBrPrij = 0 Or cBrZbr = 0 Then Exit Function

    Dim data As Variant: data = GetTableData(TBL_PRIJEMNICA)
    If IsEmpty(data) Then Exit Function

    ' IZVOR SE BIRA PO IDENTITETU. Sa poznatom generacijom menjaju se samo redovi
    ' TE generacije (Klasa I i II zajedno, i nijedan tudji). Bez nje ostaje izbor
    ' po broju, ali tek posto se dokaze da je broj jednoznacan - inace bi se
    ' prevezala i prijemnica drugog kupca koja deli broj.
    Dim srcIds As Object
    Set srcIds = IdoviGeneracije(TBL_PRIJEMNICA, COL_PRJ_ID, prijemnicaGeneracijaID)
    If Len(Trim$(prijemnicaGeneracijaID)) > 0 Then
        ' ZADATA generacija koja se ne razresava je GRESKA, ne poziv na fallback.
        ' Pad na broj bi znacio: pozivalac je rekao "bas ovaj dokument", nije
        ' nadjen, pa se dira nesto drugo. Prazan argument je nesto sasvim drugo
        ' (legacy zapis) i za njega fallback ostaje.
        If srcIds.count = 0 Then Exit Function
    Else
        ' AKTIVNI vlasnik mora biti jedan. NAMERNO nije IKAD -- probano pa
        ' povuceno, jer je bila sira zabrana bez dokaza pogresne mutacije.
        '
        ' Zastita ovde nije jedna kapija nego SLOJEVI, i svaki radi svoj posao:
        '   zaglavlje prijemnice  -> u targetRows ulaze samo AKTIVNI redovi,
        '                            pa storniran dokument drugog kupca ne
        '                            moze da se pomeri (test 126);
        '   paletna stavka sa ID  -> odlucuje IDENTITET (docIds.Exists);
        '   legacy stavka bez ID  -> brojDvosmislen racuna IKAD i puca, uz
        '                            rollback cele transakcije.
        '
        ' IKAD kapija ovde bi zabranila i potpuno resiv legacy oporavak: broj
        ' prijemnice je numerisan PO KUPCU, pa je kolizija ocekivana, a writer
        ' vec ume da razdvoji aktivan dokument. Sira zabrana bi operatera
        ' terala na rucni rad bez ijedne izmerene koristi.
        RequireJedanVlasnikPoBroju TBL_PRIJEMNICA, COL_PRJ_BROJ, brPrijemnice, SRC, _
                                   COL_PRJ_KUPAC
    End If
    Dim cPrjId As Long: cPrjId = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID)

    ' Aktivni redovi prijemnice (Klasa I i II dele broj I generaciju).
    Dim targetRows As Collection: Set targetRows = New Collection
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If PripadaIzvoru(data, i, cBrPrij, cPrjId, brPrijemnice, srcIds) Then
            If cStorno = 0 Or UCase$(Trim$(CStr(data(i, cStorno)))) <> "DA" Then
                targetRows.Add i
            End If
        End If
    Next i
    If targetRows.count = 0 Then Exit Function                  ' nema aktivne prijemnice

    ' ZBR-CHILD-01: generacija CILJA, jednom za sve redove.
    '
    ' Kad je pozivalac zadao generaciju, ona JE identitet cilja -- iznad je vec
    ' dokazano da se razresava (tgtIds). Kad nije, izvodi se iz broja i prazna je
    ' ako broj nije jednoznacan; tada dete ostaje bez generacije, sto je isto
    ' stanje kao pre ove kolone.
    Dim genCilja As String
    genCilja = Trim$(NzToText(zbirnaGeneracijaID))
    If Len(genCilja) = 0 Then genCilja = ZbirnaIDZaBroj(targetBrZbirne)

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_PRIJEMNICA
    tx.AddTableSnapshot TBL_PALETA_STAVKA

    Dim k As Long
    For k = 1 To targetRows.count
        PoveziDeteNaZbirnu TBL_PRIJEMNICA, targetRows(k), COL_PRJ_BROJ_ZBIRNE, _
                           targetBrZbirne, genCilja, SRC
    Next k

    ' Sledljivost: paletne stavke te prijemnice moraju dobiti NOVU BrojZbirne, inace
    ' ostaju sa mrtvom zbirnom (paleta -> zbirna -> kooperanti pukne). Menja se samo
    ' veza (BrojZbirne); roba/pripadnost prijemnici ostaje.
    '
    ' PO PrijemnicaID, NE PO BROJU. Ovaj upis je ranije isao po BrojPrijemnice i
    ' time je ponistavao ceo izbor iznad: tblPrijemnica se menjala samo dokumentu
    ' A, a paletne stavke i dokumenta A i dokumenta B. Dokument B je tako
    ' postajao SAM SEBI PROTIVRECAN -- prijemnica na staroj zbirni, njena paleta
    ' na novoj. Redovi su vec izabrani u targetRows; odatle se citaju ID-evi.
    Dim docIds As Object: Set docIds = CreateObject("Scripting.Dictionary")
    docIds.CompareMode = vbTextCompare
    Dim q As Long, pidT As String
    For q = 1 To targetRows.count
        If cPrjId > 0 Then
            pidT = Trim$(NzToText(data(targetRows(q), cPrjId)))
            If Len(pidT) > 0 Then docIds(pidT) = True
        End If
    Next q

    ' Stavka bez PrijemnicaID (zatecen zapis) sme po broju SAMO ako taj broj nosi
    ' jedan dokument. Kad ga nose dva, ne moze se utvrditi cija je, pa se staje
    ' umesto da se pogodi -- transakcija se povlaci u celini.
    Dim brojDvosmislen As Boolean
    brojDvosmislen = (VlasniciPoBroju(TBL_PRIJEMNICA, COL_PRJ_BROJ, brPrijemnice, SRC, _
                                      True, Array(COL_PRJ_KUPAC)).count > 1)

    Dim ps As Variant: ps = GetTableData(TBL_PALETA_STAVKA)
    If Not IsEmpty(ps) Then
        Dim pBr As Long, pZb As Long, pSt As Long, pPid As Long
        pBr = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_PRIJ)
        pZb = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_ZBIRNE)
        pSt = GetColumnIndex(TBL_PALETA_STAVKA, COL_STORNIRANO)
        pPid = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PRIJEMNICA_ID)
        If pBr > 0 And pZb > 0 Then
            Dim r2 As Long, pidS As String, pripada As Boolean
            For r2 = 1 To UBound(ps, 1)
                If pSt = 0 Or UCase$(Trim$(CStr(ps(r2, pSt)))) <> "DA" Then
                    pidS = ""
                    If pPid > 0 Then pidS = Trim$(NzToText(ps(r2, pPid)))
                    If Len(pidS) > 0 Then
                        pripada = docIds.Exists(pidS)
                    Else
                        If Trim$(CStr(ps(r2, pBr))) = brPrijemnice And brojDvosmislen Then
                            Err.Raise vbObjectError + 7311, SRC, _
                                      "Paletna stavka broja '" & brPrijemnice & "' nema " & _
                                      "PrijemnicaID, a taj broj nose dokumenta vise kupaca. " & _
                                      "Ne moze se utvrditi cija je -- prevezivanje je " & _
                                      "prekinuto. Popuni PrijemnicaID pa ponovi."
                        End If
                        pripada = (Trim$(CStr(ps(r2, pBr))) = brPrijemnice)
                    End If
                    If pripada Then
                        PoveziDeteNaZbirnu TBL_PALETA_STAVKA, r2, COL_PALS_BROJ_ZBIRNE, _
                                           targetBrZbirne, genCilja, SRC
                    End If
                End If
            Next r2
        End If
    End If

    tx.CommitTx
    Set tx = Nothing
    ReassignPrijemnicaToZbirna_TX = True
    Exit Function
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr SRC
    ReassignPrijemnicaToZbirna_TX = False
End Function

' ============================================================
' OM ULAZ (ambalaza + kes) -- servis bez UI-ja
' Premesteno iz frmDokumenta (RF-05): nema referenci na kontrole, a smesteno u
' modul moze da se testira bez instanciranja forme (core guard za smer ambalaze).
' ============================================================

Public Function SaveOMUlaz_TX(ByVal datum As Date, _
                              ByVal brojDok As String, _
                              ByVal stanicaNaziv As String, _
                              ByVal stanicaID As String, _
                              ByVal vozacID As String, _
                              ByVal tipAmb As String, _
                              ByVal kolAmb As Long, _
                              ByVal vrstaVoca As String, _
                              ByVal novac As Double, _
                              ByVal kooperantID As String, _
                              ByVal primalacDisplay As String, _
                              ByVal otkupID As String, _
                              ByVal tipNovca As String, _
                              ByVal koopSmer As String) As Boolean
    Dim tx As clsTransaction
    Dim reversID As String
    Set tx = New clsTransaction

    On Error GoTo EH

    If kolAmb <= 0 And novac <= 0 Then
        Err.Raise vbObjectError + 1501, "SaveOMUlaz_TX", _
                  Poruka("DOK_ERR_NEMA_AMBALAZE_NOVCA")
    End If

    tx.BeginTx
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_OTKUP

    If kolAmb > 0 Then
        ' Kapija broja stoji SAMO ovde, u revers grani. Cist gotovinski
        ' promet (F5 isplata / F6 uplata) prolazi kroz istu proceduru sa
        ' kolAmb = 0, nema svoj brojevni niz (broj je slobodan unos) i
        ' tamo stanicaID postaje partner-OM -- kapija nad celom procedurom
        ' odbijala bi legitimnu isplatu.
        modBrojevi.RequireBrojUKontekstu modBrojevi.KIND_REV, stanicaID, datum, _
                                         brojDok, "SaveOMUlaz_TX"

        ' Zauzetost broja u nizu (stanica, dan), sa storniranima (A9) -- ista
        ' provera koju ekran zove u ReversValidiraj. Jednom po dokumentu, pre
        ' nogu: provera po nozi odbila bi sopstvenu nogu Kooperant. Vazi za SVA
        ' CETIRI smera, i jedina je provera broja: isti KOOP broj, smer i dan na
        ' drugoj stanici je legalan (REV-IDENT-01 Faza 2b -- noge povezuje ReversID).
        modBrojevi.RequireBrojSlobodanUNizu modBrojevi.KIND_REV, stanicaID, datum, _
                                            brojDok, "SaveOMUlaz_TX"

        ' REV-IDENT-01: JEDAN identitet po dokumentu, zajednicki svim nogama
        ' (Kooperant + Stanica za KOOP, sama Stanica za FIRMA). Kuje se jednom i
        ' NASLEDJUJE u svakoj nozi -- nikad po nozi.
        reversID = modAmbalaza.NoviReversID()

        Select Case koopSmer
        Case "IZDAVANJE"
            ' OM IZDAJE prazne kooperantu -> DVOJNI upis (bez vozaca):
            '   1) Kooperant ULAZ (dobija prazne), 2) OM/Stanica IZLAZ (razduzenje OM).
            If Trim$(kooperantID) = "" Then
                Err.Raise vbObjectError + 1503, "SaveOMUlaz_TX", _
                          "Izdavanje kooperantu: kooperant je obavezan."
            End If
            If Trim$(stanicaID) = "" Then
                Err.Raise vbObjectError + 1504, "SaveOMUlaz_TX", _
                          "Izdavanje kooperantu: OM (otkupno mesto) je obavezan za razdu" & ChrW(382) & "enje."
            End If
            TrackAmbalaza datum, tipAmb, kolAmb, _
                          "Ulaz", kooperantID, "Kooperant", _
                          "", brojDok, DOK_TIP_OM_IZLAZ_KOOP, reversID
            TrackAmbalaza datum, tipAmb, kolAmb, _
                          "Izlaz", stanicaID, "Stanica", _
                          "", brojDok, DOK_TIP_OM_IZLAZ_KOOP, reversID
        Case "PRIJEM"
            ' KOOPERANT VRACA prazne na OM (povrat) -> DVOJNI upis, mirror izdavanja:
            '   1) Kooperant IZLAZ (predaje prazne), 2) OM/Stanica ULAZ (zaduzenje OM).
            If Trim$(kooperantID) = "" Then
                Err.Raise vbObjectError + 1505, "SaveOMUlaz_TX", _
                          "Prijem od kooperanta: kooperant je obavezan."
            End If
            If Trim$(stanicaID) = "" Then
                Err.Raise vbObjectError + 1506, "SaveOMUlaz_TX", _
                          "Prijem od kooperanta: OM (otkupno mesto) je obavezan za zadu" & ChrW(382) & "enje."
            End If
            TrackAmbalaza datum, tipAmb, kolAmb, _
                          "Izlaz", kooperantID, "Kooperant", _
                          "", brojDok, DOK_TIP_OM_ULAZ_KOOP, reversID
            TrackAmbalaza datum, tipAmb, kolAmb, _
                          "Ulaz", stanicaID, "Stanica", _
                          "", brojDok, DOK_TIP_OM_ULAZ_KOOP, reversID
        Case "IZDATO_OM"
            ' Vozac raspodeljuje prazne na OM (revers ide na OM): OM (Stanica) ULAZ +
            ' vozac (inverzno Izlaz = vozac se razduzuje). Vozac je prethodno zaduzen
            ' kod kupca (prijemnica-povrat / kupci-izlaz) -> hladnjaca se NE knjizi ovde.
            If Trim$(stanicaID) = "" Then
                Err.Raise vbObjectError + 1507, "SaveOMUlaz_TX", _
                          "Izdato OM: OM (otkupno mesto) je obavezan."
            End If
            If Trim$(vozacID) = "" Then
                Err.Raise vbObjectError + 1509, "SaveOMUlaz_TX", _
                          "Izdato OM: vozac je obavezan (firma<->OM ide preko vozaca)."
            End If
            TrackAmbalaza datum, tipAmb, kolAmb, _
                          "Ulaz", stanicaID, "Stanica", _
                          vozacID, brojDok, DOK_TIP_OM_ULAZ_FIRMA, reversID
        Case "PRIJEM_OD_OM"
            ' OM vraca prazne vozacu (revers ide na OM): OM (Stanica) IZLAZ + vozac
            ' (inverzno Ulaz = vozac se zaduzuje). Vozac kasnije razduzuje firmi
            ' (hladnjaci) kroz postojece kupac tokove -> hladnjaca se NE knjizi ovde.
            If Trim$(stanicaID) = "" Then
                Err.Raise vbObjectError + 1508, "SaveOMUlaz_TX", _
                          "Prijem od OM: OM (otkupno mesto) je obavezan."
            End If
            If Trim$(vozacID) = "" Then
                Err.Raise vbObjectError + 1510, "SaveOMUlaz_TX", _
                          "Prijem od OM: vozac je obavezan (firma<->OM ide preko vozaca)."
            End If
            TrackAmbalaza datum, tipAmb, kolAmb, _
                          "Izlaz", stanicaID, "Stanica", _
                          vozacID, brojDok, DOK_TIP_OM_IZLAZ_FIRMA, reversID
        Case Else
            ' Smer je OBAVEZAN uz kolicinu ambalaze. Ranije je ovde tiho knjizen
            ' legacy "OM prima od vozaca" (Stanica ULAZ, DOK_TIP_OM_ULAZ), pa je
            ' prazan/nepoznat smer davao pogresan ledger red bez ijedne poruke.
            ' UI blokira prazan smer, ovo je core guard za sve ostale pozivaoce.
            Err.Raise vbObjectError + 1511, "SaveOMUlaz_TX", _
                      "Nepoznat smer ambalaze '" & koopSmer & "'. Dozvoljeni: " & _
                      "IZDAVANJE, PRIJEM, IZDATO_OM, PRIJEM_OD_OM."
        End Select
    End If

    If novac > 0 Then
        Dim novacID As String

        ' KAPIJA VLASNISTVA I TRENUTNOG OSTATKA (AUD-026 obrazac). UI je ovo vec
        ' proverio, ali nad snimkom iz trenutka kad je lista punjena -- a izmedju
        ' punjenja i potvrde stanje se moze promeniti. Writer zato ne veruje
        ' parametrima nego cita stanje SADA. Vazi za SVAKOG pozivaoca, pa i za
        ' legacy frmDokumenta.
        Dim blokErr As String
        blokErr = IsplataBlokProblem(otkupID, kooperantID, stanicaID, novac)
        If Len(blokErr) > 0 Then
            Err.Raise vbObjectError + 1512, "SaveOMUlaz_TX", blokErr
        End If

        ' ISTA KAPIJA ZA AVANS OTKUPNOG MESTA. Blok i faktura su vec bili
        ' zasticeni i u writer-u; avans je ostajao samo na UI sloju
        ' (modNovacUnos.IsplataValidiraj), pa je writer bio poslednja linija
        ' za dve od tri stvari koje isti dokument moze da prekoraci.
        '
        ' Kes isplata kooperantu TROSI avans koji je firma dala otkupnom mestu -
        ' GetOMAvansSaldo ga i racuna kao "avansi firma->OM minus vec izdate
        ' NOV_KES_OTKUPAC_KOOP isplate". Isplata preko tog salda pravi minus
        ' koji se nigde ne vidi do sledeceg obracuna.
        If tipNovca = NOV_KES_OTKUPAC_KOOP Then
            Dim avansSaldo As Double
            avansSaldo = ZaokruziNovac(GetOMAvansSaldo(stanicaID))
            If ZaokruziNovac(novac) > avansSaldo Then
                Err.Raise vbObjectError + 1513, "SaveOMUlaz_TX", _
                          Poruka("NOVAC_ERR_AVANS_PREKO") & " " & _
                          Format$(avansSaldo, "#,##0.00")
            End If
        End If

        novacID = SaveNovac( _
            brojDok:=brojDok, _
            datum:=datum, _
            partner:=stanicaNaziv, _
            partnerId:=stanicaID, _
            entitetTip:="OM", _
            omID:=stanicaID, _
            kooperantID:=kooperantID, _
            fakturaID:="", _
            vrstaVoca:=vrstaVoca, _
            tip:=tipNovca, _
            uplata:=0, _
            isplata:=novac, _
            napomena:=primalacDisplay, _
            otkupID:=otkupID)

        If novacID = "" Then
            Err.Raise vbObjectError + 1502, "SaveOMUlaz_TX", _
                      "SaveNovac fehlgeschlagen"
        End If

    End If

    tx.CommitTx

    Set tx = Nothing

    SaveOMUlaz_TX = True
    Exit Function

EH:
    LogErr "SaveOMUlaz_TX"

    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    SaveOMUlaz_TX = False
End Function

