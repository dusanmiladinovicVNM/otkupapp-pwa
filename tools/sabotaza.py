#!/usr/bin/env python3
"""Namerno kvarenje koda -- druga polovina dokaza iz CLAUDE.md paragraf 5.

Suite koja je zelena nad ispravnim kodom, a nije POKAZANA crvena nad pokvarenim,
ne dokazuje da isla sta meri (PR #181: cetiri puta zeleno-ali-nedokazano-crveno).
Za svaku proveru zato postoji sabotaza koja bas nju obara, po imenu.

    python tools/sabotaza.py --lista
    python tools/sabotaza.py clear-datum
    python tools/run_vba.py --suite RunAllTests      # ocekuj FAIL po imenu
    python tools/sabotaza.py --vrati

TRI ZAMKE koje su ovde vec pokupljene, da ih ne pokupi operater:

1. KRAJ REDA. `src-vba` se na Windows-u checkout-uje kao CRLF, a na Linuxu kao
   LF. Sidro sa zakucanim `\\n` ne pogodi nista, skripta tiho ne uradi nista, run
   prodje nad NEIZMENJENIM fajlom -- i izgleda kao da sabotaza "nije oborila"
   suite. Zato se kraj reda detektuje, a pogodak se TVRDI (tacno jednom).

2. UVLACENJE. Sidro se poredi od POCETKA REDA. Bez toga je
   `    mFrm...cbKupac.value = ""` (4 razmaka) podniz istog reda uvucenog za 8,
   pa je isto sidro pogadjalo dva razlicita mesta.

3. VRACANJE. `git checkout --` vraca fajl na HEAD, pa BRISE i nesnimljene izmene
   koje sa sabotazom nemaju veze (jednom vec pojelo test seam-ove). Zato se
   vraca obrnutom zamenom -- dira se tacno ono sto je i pokvareno.

4. KOMENTAR POSLE `_`. U VBA line-continuation `_` mora biti POSLEDNJI znak u
   redu; `..., _   ' SABOTAZA` je syntax error. Sabotaza tada ne obara test nego
   COMPILE: run visi do timeout-a, Excel ostaje u [break], a izlaz je
   "Exception occurred" umesto imena tvrdnje. Ako sabotaza pada tako, greska je
   u sabotazi. Oznaku pisi u red IZNAD ili je izostavi -- ime u katalogu je
   dovoljna dokumentacija.

5. POGADJAJ BAS SVOJU TVRDNJU. Sabotaza koja obori PRVU tvrdnju u testu (npr.
   tako sto rutina digne gresku pa vrati False) dokazuje samo da se kod izvrsava,
   ne i da ta konkretna tvrdnja meri. Ako izlaz prijavi drugu tvrdnju od
   ocekivane, suzi sabotazu dok ne pogodi svoju.

6. `AssertEq` DIZE GRESKU, pa se test PREKIDA na prvom padu. Tvrdnje posle njega
   se ne izvrsavaju -- a to znaci da sabotaza koja obori uzgrednu tvrdnju
   ("operacija je vratila success=False") ostavlja poslovnu tvrdnju ispod nje
   NEMERENOM, i to izgleda kao uspesan dvosmerni dokaz. Redosled tvrdnji u testu
   je zato deo dokaza: NAJVAZNIJA tvrdnja ide PRVA. Simptom: izlaz prijavi drugu
   tvrdnju od one koja je u katalogu (v. zamka 5).

7. ZAMENA NE SME BITI PRAZNA. `--vrati` radi obrnutu zamenu -- trazi ZAMENU u
   fajlu i vraca sidro. Prazan string se "nalazi" na svakoj poziciji, pa tvrdnja
   o tacno jednom pogotku ne prolazi: skripta tiho ne uradi nista i prijavi
   "nema sta da se vrati", dok je fajl i dalje pokvaren. Kod fajla koji jos nije
   komitovan ni `git checkout` nije mreza. Ako sabotaza treba da UKLONI red,
   zameni ga necim ravnopravnim (duplikat susednog reda) umesto praznim.
   Placeno jednom, na `storno-cip-svi-nestao`.

8. ZAMENA NE SME BITI PODNIZ SIDRA. `--vrati` trazi ZAMENU u fajlu; ako je ona
   sadrzana u sidru, nadje je i u ZDRAVOM kodu i "vrati" ga -- to jest doda jos
   jedan primerak razlike. Simptom: posle svakog `--vrati` fajl dobije red vise
   (kod nas tri uzastopna `Err.Clear`), a git diff raste bez ijedne namerne
   izmene. Ako se sabotaza svodi na UKLANJANJE reda, dodaj joj oznaku
   (`   ' SABOTAZA: ...`) da zamena postane jedinstvena. Placeno jednom, na
   `ekran-curi-greska`.
"""

import argparse
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_VBA = os.path.join(ROOT, "src-vba")

# ime -> (fajl, sidro, zamena, test koji MORA da padne, sta tvrdnja kaze)
# Sidro i zamena se porede od POCETKA REDA (v. zamka 2) -- ne pisati vodece \n.
SABOTAZE = {
    # --- ParseDatum ---------------------------------------------------------
    "parse-tacka": (
        "modOtkupUI.bas",
        '    Do While Right$(t, 1) = "."\n'
        "        t = Left$(t, Len(t) - 1)\n"
        "    Loop\n",
        "    ' SABOTAZA: trailing tacka se vise ne skida\n",
        "T_ParseDatum_Ugovor",
        "trailing tacka se skida, ne obara unos",
    ),
    "parse-cdate": (
        "modOtkupUI.bas",
        "    If TryParseDateValue(t, d) Then ParseDatum = CDbl(d)\n",
        "    If IsDate(t) Then ParseDatum = CDbl(CDate(t))   ' SABOTAZA\n",
        "T_ParseDatum_Ugovor",
        "godina van poslovnog opsega",
    ),
    # --- ParcelaID ----------------------------------------------------------
    "parcela-tekst": (
        "modOtkupUI.bas",
        "    If CB.ListIndex >= 0 Then ParcelaID = Trim$(CStr(CB.List(CB.ListIndex, 1)))\n",
        "    If CB.ListIndex >= 0 Then ParcelaID = Trim$(CStr(CB.text))   ' SABOTAZA\n",
        "T_ParcelaID_IzSkriveneKolone",
        "ID parcele dolazi iz skrivene kolone, ne iz prikaznog teksta",
    ),
    "parcela-vidljivost": (
        "modOtkupUI.bas",
        '    If Not mFrm.Controls("zForm").Controls("fgParcela").Visible Then Exit Function\n',
        "    ' SABOTAZA: provera vidljivosti polja uklonjena\n",
        "T_ParcelaID_IzSkriveneKolone",
        "sakriveno polje ne salje parcelu u dokument",
    ),
    # --- ClearForm ----------------------------------------------------------
    "clear-datum": (
        "modOtkupUI.bas",
        '    If Not imaOtp Then SetDatumDanas mFrm.Controls("zForm")\n',
        '    SetDatumDanas mFrm.Controls("zForm")   \' SABOTAZA\n',
        "T_ClearForm_Ugovor",
        "dok je otpremnica aktivna datum se NE vraca na danas",
    ),
    "clear-zbirna": (
        "modOtkupUI.bas",
        '    nmv = Array("fgBrOtpr", "fgKgI", "fgKgII", "fgKolAmb", "fgAmbPr", "fgNovac")\n',
        '    nmv = Array("fgBrOtpr", "fgKgI", "fgKgII", "fgKolAmb", "fgAmbPr", "fgNovac", "fgBrZbir")\n',
        "T_ClearForm_Ugovor",
        "broj zbirne je kontekst -- ne brise se posle snimanja",
    ),
    "clear-partner": (
        "modOtkupUI.bas",
        '    mFrm.Controls("zCtx").Controls("cbKupac").value = ""\n',
        "    ' SABOTAZA: partner se vise ne brise\n",
        "T_ClearForm_Ugovor",
        "partner mora da bude obrisan posle snimanja",
    ),
    # --- upis zbirne (F3) ---------------------------------------------------
    # Sidra su namerno vise-linijska: OtpremnicaValidiraj u istom fajlu ima
    # doslovno iste redove, pa jednolinijsko sidro pogadja dva mesta i skripta
    # odbija da radi (v. zamka 2).
    "zbirna-vozac": (
        "modDokUnos.bas",
        '    If Len(S(p, "vozacID")) = 0 Then\n'
        '        fokus = "vozacID": ZbirnaValidiraj = Poruka("DOKUNOS_ERR_VOZAC"): Exit Function\n'
        "    End If\n",
        "    ' SABOTAZA: zbirna vise ne trazi vozaca\n",
        "T_ZbirnaValidiraj_TraziVozaca",
        "zbirna bez vozaca se odbija",
    ),
    "zbirna-kapija": (
        "modDokUnos.bas",
        '    If Not ZbirnaSeSlazeSaIzvorom(S(p, "brDok"), kolI, kolII, kolAmb + kolAmbII, dveKl) Then\n'
        '        fokus = "kolicinaI"\n'
        '        ZbirnaValidiraj = Poruka("DOK_MSG_VALIDACIJA_NIJE_PROSLA")\n'
        "        Exit Function\n"
        "    End If\n",
        "    ' SABOTAZA: zbir se vise ne poredi sa otpremnicama\n",
        "T_ZbirnaValidiraj_MoraDaSeSlazeSaOtpremnicama",
        "zbirna koja ne prijavljuje sve kilograme otpremnica se odbija",
    ),
    # Podmukliji oblik iste greske: kapija ostaje, ali se gejtuje podesavanjem.
    # Sa ukljucenom validacijom (default) sve i dalje radi -- pada tek tvrdnja
    # da kapija vazi i kad je VALIDACIJA_UNOSA iskljucena.
    "zbirna-kapija-strogo": (
        "modDokUnos.bas",
        '    If Not ZbirnaSeSlazeSaIzvorom(S(p, "brDok"), kolI, kolII, kolAmb + kolAmbII, dveKl) Then\n',
        '    If strogo And Not ZbirnaSeSlazeSaIzvorom(S(p, "brDok"), kolI, kolII, kolAmb + kolAmbII, dveKl) Then   \' SABOTAZA\n',
        "T_ZbirnaValidiraj_MoraDaSeSlazeSaOtpremnicama",
        "kapija vazi i kad je VALIDACIJA_UNOSA iskljucena",
    ),
    # --- upis prijemnice (F4) -----------------------------------------------
    "prijemnica-kupac": (
        "modDokUnos.bas",
        '    If Len(S(p, "kupacID")) = 0 Then\n'
        '        fokus = "kupacID": PrijemnicaValidiraj = Poruka("DOKUNOS_ERR_KUPAC"): Exit Function\n'
        "    End If\n",
        "    ' SABOTAZA: prijemnica vise ne trazi kupca\n",
        "T_PrijemnicaValidiraj_TraziKupca",
        "prijemnica bez kupca se odbija",
    ),
    "bruto-prijemnica": (
        "modDokUnos.bas",
        '            PrijemnicaValidiraj = Poruka("DOK_MSG_TEZINA_AMBALAZE") & Format$(tara, "#,##0.00") & _\n'
        '                                  " kg) " & Poruka("OTKUNOS_ERR_TARA_VECA")\n'
        "            Exit Function\n"
        "        End If\n"
        '        p("brutoKgI") = kolI\n'
        "        kolI = kolI - tara\n"
        '        p("kolicinaI") = kolI\n',
        '            PrijemnicaValidiraj = Poruka("DOK_MSG_TEZINA_AMBALAZE") & Format$(tara, "#,##0.00") & _\n'
        '                                  " kg) " & Poruka("OTKUNOS_ERR_TARA_VECA")\n'
        "            Exit Function\n"
        "        End If\n"
        "        ' SABOTAZA: uneti bruto ostaje u Kolicini, tara se ne oduzima\n",
        "T_BrutoNeto_PoRezimu",
        "u Kolicinu Kl.I ide neto (bruto - tara)",
    ),
    # Obrnut smer istog pravila: zbirna DOBIJA bruto->neto koji ne sme da ima.
    "bruto-zbirna": (
        "modDokUnos.bas",
        "    ' Hard-blokada: izvorne otpremnice imaju Klasu II a prekidac je iskljucen ->\n",
        '    If OtkupBrutoUnos() And kolAmb > 0 Then kolI = kolI - kolAmb * GetTezinaGajbice(S(p, "tipAmb")): p("kolicinaI") = kolI   \' SABOTAZA\n'
        "    ' Hard-blokada: izvorne otpremnice imaju Klasu II a prekidac je iskljucen ->\n",
        "T_BrutoNeto_PoRezimu",
        "zbirna se NE preracunava iz bruta",
    ),
    # --- upis isplate (F5) --------------------------------------------------
    # Tip novca je jedino sto ovaj rezim odlucuje, pa su sve tri sabotaze o
    # njemu: pogresan tip se ne vidi u formi, nego tek u saldu.
    "isplata-tip-blok": (
        "modNovacUnos.bas",
        '                p("tipNovca") = NOV_VIRMAN_FIRMA_KOOP\n',
        '                p("tipNovca") = NOV_VIRMAN_AVANS_KOOP   \' SABOTAZA\n',
        "T_IsplataValidiraj_TipNovcaPoIzboru",
        "uz blok bez prekidaca isplata je virman firme, ne avans",
    ),
    "isplata-avans-saldo": (
        "modNovacUnos.bas",
        '                omSaldo = GetOMAvansSaldo(S(p, "stanicaID"))\n'
        "                If novac > omSaldo Then\n"
        '                    fokus = "novac"\n'
        '                    IsplataValidiraj = Poruka("DOK_MSG_NEDOVOLJNO_AVANSA_RASPOLOZIVO") & " " & _\n'
        '                                       Format$(omSaldo, "#,##0.00") & " RSD"\n'
        "                    Exit Function\n"
        "                End If\n",
        "                ' SABOTAZA: saldo OM avansa se vise ne proverava\n",
        "T_IsplataValidiraj_TipNovcaPoIzboru",
        "iz OM avansa se ne isplacuje vise nego sto ga ima",
    ),
    "isplata-om-entitet": (
        "modNovacUnos.bas",
        '        If partTip = "OM" And Len(partID) > 0 Then\n'
        '            p("stanicaID") = partID\n'
        '            If Len(S(p, "partnerTekst")) > 0 Then p("stanicaTekst") = S(p, "partnerTekst")\n'
        "        End If\n",
        "        ' SABOTAZA: izabrano otkupno mesto vise nije entitet novca\n",
        "T_IsplataValidiraj_TipNovcaPoIzboru",
        "primalac otkupno mesto JESTE entitet novca, ne kontekst forme",
    ),
    # --- upis uplate (F6) ---------------------------------------------------
    "uplata-tip-faktura": (
        "modNovacUnos.bas",
        '        p("tipNovca") = NOV_KUPCI_UPLATA\n',
        '        p("tipNovca") = NOV_KUPCI_AVANS   \' SABOTAZA\n',
        "T_UplataValidiraj_FakturaOdlucujeTip",
        "uz izabranu fakturu uplata zatvara fakturu, nije avans",
    ),
    "uplata-preko-fakture": (
        "modNovacUnos.bas",
        '        ostatak = D(p, "fakturaOstatak")\n'
        "        If ostatak > 0 And novac > ostatak Then\n"
        '            fokus = "novac"\n'
        '            UplataValidiraj = Poruka("NOVUNOS_ERR_VECI_OD_FAKTURE") & " " & _\n'
        '                              Format$(ostatak, "#,##0.00")\n'
        "            Exit Function\n"
        "        End If\n",
        "        ' SABOTAZA: uplata preko preostalog iznosa fakture vise ne staje\n",
        "T_UplataValidiraj_FakturaOdlucujeTip",
        "preko preostalog iznosa fakture se ne uplacuje",
    ),
    # --- upis reversa (F7) --------------------------------------------------
    "revers-smer": (
        "modNovacUnos.bas",
        '    smer = L(p, "smerRev")\n'
        "    If smer < SMER_REV_IZD_KOOP Or smer > SMER_REV_PRI_OM Then\n"
        '        fokus = "smerRev": ReversValidiraj = Poruka("NOVUNOS_ERR_SMER"): Exit Function\n'
        "    End If\n",
        '    smer = L(p, "smerRev")   \' SABOTAZA: smer vise nije obavezan\n',
        "T_ReversValidiraj_SmerJeObavezan",
        "revers bez izabranog smera se ne knjizi",
    ),
    "revers-kupac": (
        "modNovacUnos.bas",
        '        If Len(S(p, "partnerID")) = 0 Or partTip <> "KOOP" Then\n',
        '        If Len(S(p, "partnerID")) = 0 Then   \' SABOTAZA: tip partnera se ne gleda\n',
        "T_ReversValidiraj_SmerJeObavezan",
        "kooperantski smer reversa ne prima kupca",
    ),
    # --- ukucan a nerazresen izbor (F5/F6/F7) -------------------------------
    # Najopasnija klasa greske u ovom rezimu: dokument se knjizi kao ISPRAVAN,
    # samo na pogresnog partnera / kao avans. Bez ovih kapija testovi su zeleni
    # a novac ide na pogresno mesto.
    "nerazresen-partner": (
        "modNovacUnos.bas",
        '    If NerazresenIzbor(S(p, "partnerTekst"), S(p, "partnerID")) Then\n'
        '        fokus = "partnerID": IsplataValidiraj = Poruka("NOVUNOS_ERR_PARTNER_NEIZABRAN"): Exit Function\n'
        "    End If\n",
        "    ' SABOTAZA: ukucan partner bez izbora opet prolazi\n",
        "T_NerazresenIzbor_NeProlaziKaoPrazno",
        "ukucano ime partnera bez izbora iz liste se ne knjizi",
    ),
    "nerazresen-faktura": (
        "modNovacUnos.bas",
        '    If NerazresenIzbor(S(p, "fakturaTekst"), S(p, "fakturaID")) Then\n'
        '        fokus = "fakturaID": UplataValidiraj = Poruka("NOVUNOS_ERR_FAKTURA_NEIZABRANA"): Exit Function\n'
        "    End If\n",
        "    ' SABOTAZA: ukucana faktura bez izbora opet prolazi kao avans\n",
        "T_NerazresenIzbor_NeProlaziKaoPrazno",
        "ukucana faktura bez izbora iz liste ne postaje avans kupca",
    ),
    # --- kapija vlasnistva i trenutnog ostatka ------------------------------
    "blok-tudj-koop": (
        "modNovac.bas",
        '    colKoop = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT, SRC)\n'
        '    If StrComp(Trim$(CStr(data(r, colKoop))), Trim$(kooperantID), vbTextCompare) <> 0 Then\n'
        '        IsplataBlokProblem = Poruka("NOVAC_ERR_BLOK_TUDJ_KOOP") & " " & otkupID\n'
        "        Exit Function\n"
        "    End If\n",
        "    ' SABOTAZA: vlasnik bloka se vise ne proverava\n",
        "T_IsplataBlokGuard_VlasnistvoITrenutniOstatak",
        "isplata se ne vezuje za blok drugog kooperanta",
    ),
    "blok-tudj-om": (
        "modNovac.bas",
        '            If StrComp(Trim$(CStr(data(r, colSt))), Trim$(stanicaID), vbTextCompare) <> 0 Then\n'
        '                IsplataBlokProblem = Poruka("NOVAC_ERR_BLOK_TUDJ_OM") & " " & otkupID\n'
        "                Exit Function\n"
        "            End If\n",
        "            ' SABOTAZA: otkupno mesto bloka se vise ne proverava\n",
        "T_IsplataBlokGuard_VlasnistvoITrenutniOstatak",
        "blok sa drugog otkupnog mesta se ne razduzuje na aktivnom",
    ),
    # Podmukliji oblik: kapija ostaje, ali umesto trenutnog stanja veruje
    # vrednosti koju je poslao ekran. Obara TRI testa i to je tacan nalaz --
    # isto pravilo je namerno provereno na tri nivoa (kapija, put unosa, ruta).
    "blok-ostatak-snapshot": (
        "modNovac.bas",
        "    preostalo = vrednost - GetUplataForOtkup(otkupID)\n",
        "    preostalo = vrednost + 1000000   ' SABOTAZA: ostatak se ne cita iz podataka\n",
        "T_IsplataBlokGuard_VlasnistvoITrenutniOstatak",
        "neisplaceni ostatak se cita iz podataka, ne iz snimka ekrana",
    ),
    "faktura-tudj-kupac": (
        "modNovac.bas",
        '    colKupac = RequireColumnIndex(TBL_FAKTURE, COL_FAK_KUPAC, SRC)\n'
        '    If StrComp(Trim$(CStr(data(r, colKupac))), Trim$(kupacID), vbTextCompare) <> 0 Then\n'
        '        UplataFakturaProblem = Poruka("NOVAC_ERR_FAK_TUDJ_KUPAC") & " " & fakturaID\n'
        "        Exit Function\n"
        "    End If\n",
        "    ' SABOTAZA: vlasnik fakture se vise ne proverava\n",
        "T_UplataValidiraj_FakturaOdlucujeTip",
        "uplata se ne vezuje za fakturu drugog kupca",
    ),
    # --- writer se brani sam ------------------------------------------------
    # Jedina sabotaza koja ne dira UI put: dokazuje da kapija postoji i kad
    # pozivalac nije nas ekran (legacy forma, uvoz, bilo ko).
    "writer-bez-kapije": (
        "modDokumenta.bas",
        "        Dim blokErr As String\n"
        "        blokErr = IsplataBlokProblem(otkupID, kooperantID, stanicaID, novac)\n"
        "        If Len(blokErr) > 0 Then\n"
        '            Err.Raise vbObjectError + 1512, "SaveOMUlaz_TX", blokErr\n'
        "        End If\n",
        "        ' SABOTAZA: writer opet veruje parametrima\n",
        "T_WriterGuard_OdbijaTudjBlok",
        "SaveOMUlaz_TX sam odbija nemogucu kombinaciju bloka i otkupnog mesta",
    ),
    # --- ruta ekrana --------------------------------------------------------
    # --- kapije nad novcem (hotfix posle pregleda #190) ---------------------
    # Vracanje na tacno onu formulaciju koja je propustala uplatu na vec
    # zatvorenu fakturu: kad je preostalo 0 ili manje, "preostalo > 0" je False
    # pa cela kapija cuti.
    "faktura-preostalo-nula": (
        "modNovac.bas",
        "    If preostalo <= 0 Then\n"
        '        UplataFakturaProblem = Poruka("NOVAC_ERR_FAK_VEC_PLACENA") & " " & fakturaID\n'
        "    ElseIf ZaokruziNovac(iznos) > preostalo Then\n",
        "    If False Then   ' SABOTAZA: vec placena faktura opet prolazi\n"
        '        UplataFakturaProblem = Poruka("NOVAC_ERR_FAK_VEC_PLACENA") & " " & fakturaID\n'
        "    ElseIf preostalo > 0 And ZaokruziNovac(iznos) > preostalo Then\n",
        "T_UplataGuard_VecPlacenaFaktura",
        "potpuno placena faktura ne prima jos jednu uplatu",
    ),
    # Obrnut smer istog pravila: kapija se prosiruje i na fakturu BEZ iznosa,
    # koju nikad nije ni smela da blokira.
    "faktura-bez-iznosa": (
        "modNovac.bas",
        "    If iznosFak <= 0 Then Exit Function          ' faktura bez iznosa - bez kapije\n",
        "    ' SABOTAZA: i faktura bez iznosa se sada blokira\n",
        "T_UplataGuard_VecPlacenaFaktura",
        "faktura bez evidentiranog iznosa ne blokira uplatu",
    ),
    "avans-bez-writer-kapije": (
        "modDokumenta.bas",
        "        If tipNovca = NOV_KES_OTKUPAC_KOOP Then\n"
        "            Dim avansSaldo As Double\n"
        "            avansSaldo = ZaokruziNovac(GetOMAvansSaldo(stanicaID))\n"
        "            If ZaokruziNovac(novac) > avansSaldo Then\n",
        "        If False Then   ' SABOTAZA: writer vise ne cuva avans saldo OM\n"
        "            Dim avansSaldo As Double\n"
        "            avansSaldo = ZaokruziNovac(GetOMAvansSaldo(stanicaID))\n"
        "            If ZaokruziNovac(novac) > avansSaldo Then\n",
        "T_WriterGuard_AvansSaldoOM",
        "isplata iz OM avansa preko salda se odbija u WRITER-u, ne samo u UI",
    ),
    "ruta-zbirna": (
        "modScrDokumenti.bas",
        '        Case "ZBIRNA"\n'
        "            Scr_Save = SaveZbirna(polja)\n"
        "            Exit Function\n",
        "        ' SABOTAZA: zbirna vise nije vezana na svoj upis\n",
        "T_ScrSave_RutaPoRezimu",
        "zbirna ide u modDokUnos.ZbirnaValidiraj",
    ),
    "ruta-isplata": (
        "modScrDokumenti.bas",
        '        Case "AMB_ISPLATE"\n'
        "            Scr_Save = SaveIsplata(polja)\n"
        "            Exit Function\n",
        "        ' SABOTAZA: isplata vise nije vezana na svoj upis\n",
        "T_ScrSave_RutaPoRezimu",
        "isplata ide u modNovacUnos.IsplataValidiraj",
    ),
    # --- F8 storno centar ---------------------------------------------------
    "f8-jedna-tabela": (
        "modScrDokumenti.bas",
        "    If mk = \"STORNO\" Then EffKey = StornoTipKey() Else EffKey = mk\n",
        "    EffKey = mk   ' SABOTAZA: F8 opet svira po jednoj tabeli\n",
        "T_F8_TipBiraTabeluIKolone",
        "F8 cita tabelu IZABRANOG tipa, ne uvek tblOtpremnica",
    ),
    "f8-tabela-tipa": (
        "modScrDokumenti.bas",
        "        Case \"FAKTURA\":     TabelaTipa = TBL_FAKTURE\n",
        "        ' SABOTAZA: tip fakture ispao iz mape tabela\n",
        "T_F8_TipBiraTabeluIKolone",
        "svaki od devet tipova F8 ima svoju tabelu",
    ),
    # --- kapije storna ------------------------------------------------------
    "storno-nema-dok": (
        "modStornoDok.bas",
        "            If Len(LookupActiveID(TBL_OTKUP, COL_OTK_BR_DOK, broj, COL_OTK_ID)) = 0 Then _\n"
        "                StornoRazlog = NijePronadjen(broj)\n",
        "            ' SABOTAZA: nepostojeci otkup prolazi kapiju\n",
        "T_StornoDok_KapijePreUpisa",
        "kapija zaustavlja nepostojeci dokument PRE poziva Storno*_TX",
    ),
    # --- prefill posle storna (Z10) -----------------------------------------
    "prefill-zbirna-kolona": (
        "modStornoDok.bas",
        "        Case STIP_ZBIRNA:     ColKolicinaZaPrefill = COL_ZBR_KOLICINA\n",
        '        Case STIP_ZBIRNA:     ColKolicinaZaPrefill = "Kolicina"   \' SABOTAZA\n',
        "T_PrefillIzStorniranog_CitaSvojuTabelu",
        "zbirna cita UkupnoKolicina -- literal 'Kolicina' tiho vraca nulu",
    ),
    "prefill-tabela": (
        "modStornoDok.bas",
        "        Case STIP_OTKUP:      TabelaZaPrefill = TBL_OTKUP\n",
        "        Case STIP_OTKUP:      TabelaZaPrefill = TBL_OTPREMNICA   ' SABOTAZA\n",
        "T_PrefillIzStorniranog_CitaSvojuTabelu",
        "prefill cita tabelu SVOG tipa (otkup i otpremnica dele broj 1/TEST)",
    ),
    "prefill-broj": (
        "modStornoDok.bas",
        '    res = Spoji(res, "fokus", "kolicina")\n',
        '    res = Spoji(res, "brdok", NzToText(d(base, cBroj)))   \' SABOTAZA\n',
        "T_PrefillIzStorniranog_CitaSvojuTabelu",
        "broj dokumenta se NE preuzima -- ispravka je nov dokument, nov broj",
    ),
    "framework-otkup": (
        "modStornoDok.bas",
        "        Case STIP_OTPREMNICA: TipUFlowDoc = FLOW_DOC_OTPREMNICA\n",
        "        Case STIP_OTPREMNICA, STIP_OTKUP: TipUFlowDoc = FLOW_DOC_OTPREMNICA   ' SABOTAZA\n",
        "T_FrameworkIspravke_SamoCetiriTipa",
        "framework ispravke vazi SAMO za cetiri tipa sa nizvodnim tokom",
    ),
    # --- identitet dokumenta i fail-closed grane (hardening posle review-a) ---
    "prefill-fallback-po-broju": (
        "modDokumenta.bas",
        "        ' to realan scenario, a ne teorijski.\n"
        "        Exit Function\n",
        "        ' SABOTAZA: nepostojeci PK opet pada nazad na broj\n",
        "T_Prefill_PoIdentitetuNePoBroju",
        "zadat a nepostojeci PK NE sme da prefiluje tudji dokument istog broja",
    ),
    "prefill-anchor-broj": (
        "modDokumenta.bas",
        "    If Len(Trim$(oldDocID)) > 0 And cId > 0 Then\n"
        "        For r = 1 To UBound(data, 1)\n"
        "            If Trim$(NzToText(data(r, cId))) = Trim$(oldDocID) Then\n",
        "    If False Then   ' SABOTAZA: PK se ignorise, ide se po broju\n"
        "        For r = 1 To UBound(data, 1)\n"
        "            If Trim$(NzToText(data(r, cId))) = Trim$(oldDocID) Then\n",
        "T_Prefill_PoIdentitetuNePoBroju",
        "prefill bira dokument po PK-u, ne po broju (dva kupca dele broj)",
    ),
    # --- identitet dokumenta na granici prevezivanja (zavrsnica Faze D) -----
    "relink-izvor-po-broju": (
        "modPaletniList.bas",
        "            ElseIf JeIzvornaStavka(bp, oldBroj, Trim$(CStr(ps(i, sPid))), srcIds) Then\n",
        "            ElseIf bp = oldBroj Then   ' SABOTAZA: izvor se opet bira po broju\n",
        "T_RelinkPoGeneraciji_NeDiraTudjDokument",
        "prevezivanje dira SAMO svoj dokument, i kad dva dele broj",
    ),
    "relink-ignorise-generaciju": (
        "modPaletniList.bas",
        "    Dim srcIds As Object: Set srcIds = IdoviGeneracije(TBL_PRIJEMNICA, COL_PRJ_ID, oldGeneracijaID)\n",
        '    Dim srcIds As Object: Set srcIds = IdoviGeneracije(TBL_PRIJEMNICA, COL_PRJ_ID, "")   \' SABOTAZA\n',
        "T_RelinkPoGeneraciji_NeDiraTudjDokument",
        "generacija izvora se stvarno koristi, ne samo prosledjuje",
    ),
    # Cilj prevezivanja. Izvor po identitetu a cilj po labeli i dalje moze da
    # odnese palete pogresnom kupcu -- samo na drugom kraju.
    "relink-cilj-po-broju": (
        "modPaletniList.bas",
        "        If Len(tgtGen) > 0 Then\n"
        "            ciljni = (Trim$(NzToText(prj(r, pcGen))) = tgtGen)\n"
        "        Else\n"
        "            ciljni = (Trim$(CStr(prj(r, pcBr))) = newBroj)\n"
        "        End If\n",
        "        ciljni = (Trim$(CStr(prj(r, pcBr))) = newBroj)   ' SABOTAZA\n",
        "T_RelinkPoGeneraciji_NeDiraTudjDokument",
        "cilj se bira po identitetu, ne po broju koji dele dva kupca",
    ),
    "relink-cilj-bez-kapije": (
        "modPaletniList.bas",
        "        If VlasniciPoBroju(TBL_PRIJEMNICA, COL_PRJ_BROJ, newBroj, SRC, False, _\n"
        "                           Array(COL_PRJ_KUPAC)).count > 1 Then\n",
        "        If False Then   ' SABOTAZA: dvosmislen cilj vise ne zaustavlja\n",
        "T_RelinkPoGeneraciji_NeDiraTudjDokument",
        "bez generacije cilja dvosmislen broj se odbija (fail-closed)",
    ),
    # Propagacija BrojZbirne u paletne stavke. Izbor redova prijemnice je bio
    # tacan, pa je ovaj drugi upis po BROJU ponistavao ceo taj izbor.
    "zbirna-paleta-po-broju": (
        "modDokumenta.bas",
        "                        pripada = docIds.Exists(pidS)\n",
        "                        pripada = (Trim$(CStr(ps(r2, pBr))) = brPrijemnice)   ' SABOTAZA\n",
        "T_PrevezivanjeNaZbirnu_PaletaIdePoIdentitetu",
        "paletna stavka tudjeg dokumenta istog broja se NE pomera",
    ),
    # Zadata generacija koje nema nije poziv na fallback po broju.
    "generacija-nema-pa-po-broju": (
        "modDokumenta.bas",
        "        If srcIds.count = 0 Then Exit Function\n",
        "        If False Then Exit Function   ' SABOTAZA: pada na broj\n",
        "T_ZadataGeneracijaKojeNema_Staje",
        "zadata generacija koje nema zaustavlja upis, ne prelazi na broj",
    ),
    # Presuda o relabelu. Writer bira dokument po generaciji; ako presuda opet
    # trazi dokument po broju, opisuje tudji -- i relabel se tiho preskoci.
    "verdikt-po-broju": (
        "modPaletniList.bas",
        "    verdict = PresudiPaletaReassign(oVrS, oSoS, oTaS, nVr, nSo, nTa, oldGajbByKl, newGajb)\n",
        "    verdict = EvaluatePaletaReassign(oldBroj, newBroj)   ' SABOTAZA\n",
        "T_VerdiktPoIdentitetu_RelabelSeNePreskace",
        "presuda opisuje izabran dokument, ne prvi sa tim brojem",
    ),
    # Otpremnica flow mutira roditeljsku zbirnu po golom broju.
    "otpremnica-bez-kapije-nad-zbirnom": (
        "modStornoFlow.bas",
        "    If mode <> SV_MODE_RESI_KASNIJE Then\n"
        "        If ZbirnaBrojJeDvosmislenIkad(parentZbirna) Then\n",
        "    If False Then   ' SABOTAZA: dvosmislena roditeljska zbirna se ignorise\n"
        "        If ZbirnaBrojJeDvosmislenIkad(parentZbirna) Then\n",
        "T_OtpremnicaNadDvosmislenomZbirnom_Staje",
        "DUPLI staje kad je broj roditeljske zbirne dvosmislen",
    ),
    # Zatecen PENDING context iz starije verzije zaobilazi kapiju na startu.
    "zatecen-context-bez-kapije": (
        "modStornoFlow.bas",
        "    If ZbirnaBrojJeDvosmislenIkad(oldZbirna) Then\n",
        "    If False Then   ' SABOTAZA: zatecen context prolazi bez provere\n",
        "T_ZatecenContext_NePrevezujeTudjePrijemnice",
        "tudja prijemnica NIJE prevezana na novu zbirnu",
    ),
    # Ista kapija, ali ono STO proverava: roditelj po poslovnom broju umesto iz
    # context-a. Vraca tacno diverganciju iz pregleda -- kapija proveri
    # jednoznacnu zbirnu SIBLINGA, a mutacije nize idu nad oldZbirna izabranog
    # dokumenta. Guard prolazi, tudja prijemnica se preveze.
    "stale-parent-po-broju": (
        "modStornoFlow.bas",
        "    If ZbirnaBrojJeDvosmislenIkad(oldZbirna) Then\n",
        "    Dim sabZbirna As String   ' SABOTAZA: roditelj po broju, ne iz context-a\n"
        "    sabZbirna = NzTx(LookupValue(TBL_OTPREMNICA, COL_OTP_BROJ, oldBroj, COL_OTP_BROJ_ZBIRNE))\n"
        "    If ZbirnaBrojJeDvosmislenIkad(sabZbirna) Then\n",
        "T_ZatecenContext_NePrevezujeTudjePrijemnice",
        "tudja prijemnica NIJE prevezana na novu zbirnu",
    ),
    # Nesimetricna zastita: izvor cuvan, CILJ nije. Nizvodne operacije nad ciljem
    # idu po golom broju, a zatecena kapija u writeru broji samo AKTIVNE vlasnike
    # -- pa storniran vlasnik sa aktivnom decom prolazi.
    "cilj-bez-istorijske-kapije": (
        "modStornoFlow.bas",
        "    If ZbirnaBrojJeDvosmislenIkad(newZbirna) Then\n",
        "    If False Then   ' SABOTAZA: ciljna zbirna se ne proverava\n",
        "T_CiljnaZbirnaDvosmislena_Staje",
        "aktivno ciljno zaglavlje NIJE rekalkulisano preko tudje dece",
    ),
    # Kes tabela memoise NEUSPEH -- zatecen incident sa prave instalacije:
    # prazne liste za svaki tip dokumenta, bez ijedne greske, dok je tabela puna.
    "kes-memoise-neuspeh": (
        "modUiData.bas",
        "    If IsArray(src) Then mCache(tblName) = src\n",
        "    mCache(tblName) = src   ' SABOTAZA: kesira se i Empty\n",
        "T_KesTabela_NeMemoiseNeuspeh",
        "neuspeh se NE kesira -- inace tabela ostaje prazna do kraja sesije",
    ),
    # Prazna mapa imena se kesira -> svako ime pada na goli ID (KOOP-00022).
    "mapa-imena-kesira-prazno": (
        "modOtkupUI.bas",
        "    If d.count > 0 Then Set mPartMap(ck) = d\n",
        "    Set mPartMap(ck) = d   ' SABOTAZA: kesira se i prazna mapa\n",
        "T_MapaImena_KljucNosiKolone",
        "prazna mapa se NE kesira -- inace svako ime pada na goli ID",
    ),
    # Kljuc kesa samo po imenu tabele: prvi pozivalac odlucuje za sve ostale.
    "mapa-imena-kljuc-bez-kolona": (
        "modOtkupUI.bas",
        '    ck = tblName & "|" & idCol & "|" & nameCol & "|" & nameCol2\n',
        "    ck = tblName   ' SABOTAZA: kljuc ne nosi kolone\n",
        "T_MapaImena_KljucNosiKolone",
        "kljuc kesa nosi KOLONE -- ime+prezime nije isto sto i samo ime",
    ),
    # --- ekran Storno (v6-ui-143) --------------------------------------------
    # NAJSKUPLJA tvrdnja migracije: nevidljiva kolona identiteta. Do v6-ui-141
    # se dodavala pod uslovom ActiveMode = "F8"; ekran nema rezim, pa bi taj
    # uslov cutke bio False i ceo lanac iz #198 bi pao na biranje po BROJU --
    # bez ijedne greske i bez ijedne crvene suite, jer testovi identiteta
    # (35, 45, 46, 48-52) mere sloj ISPOD mreze.
    "storno-bez-kolone-identiteta": (
        "modScrDokumenti.bas",
        "    If saIdentitetom Then\n",
        "    If False Then   ' SABOTAZA: kolona identiteta se ne dodaje\n",
        "T_StornoEkran_KolonaIdentiteta",
        "opis kolona za Storno nosi kolonu identiteta, i to POSLEDNJU",
    ),
    # Suprotan smer: kapija koja je uvek otvorena nije kapija. GridCols je
    # zajednicki za rezim unosa i za Storno nad istim tipom, pa bi bezuslovno
    # dodavanje promenilo i mrezu unosa.
    "storno-identitet-uvek": (
        "modScrDokumenti.bas",
        "    If saIdentitetom Then\n",
        "    If True Then   ' SABOTAZA: kolona identiteta ide i unosnom rezimu\n",
        "T_StornoEkran_KolonaIdentiteta",
        "unosni rezim NE dobija kolonu identiteta",
    ),
    # Storno je ekran BEZ upisa. Ako mu se vrati "upis=da", ljuska mu crta
    # primarno dugme koje nema sta da pozove -- tacno stanje od pre v6-ui-143.
    "storno-ekran-ima-upis": (
        "modScrStorno.bas",
        '               "|lista=OTKUI_SCRST_LISTA|oblik=lista|upis=ne"\n',
        '               "|lista=OTKUI_SCRST_LISTA|oblik=lista|upis=da"\n',
        "T_StornoJeEkranNeRezim",
        "Storno nema upis -- forma i primarno dugme mu ne pripadaju",
    ),
    # 'valid = True' mora da znaci 'svih sedam sekcija je pouzdano procitano'. Bez
    # strict rezima citac na nedostajucu kolonu vrati PRAZNU kolekciju, uvid stigne
    # do kraja i oznaci se kao valjan -- pa ekran kaze 'nema paleta' i ponudi
    # mutaciju, iako je tacan odgovor 'ne znam da li ih ima'.
    "uvid-guta-necitljivo": (
        "modStornoImpact.bas",
        "                Set ImpactPalete = GetPaleteImpactByField(COL_PALS_PRIJEMNICA_ID, \"\", ids, strict)\n",
        "                Set ImpactPalete = GetPaleteImpactByField(COL_PALS_PRIJEMNICA_ID, \"\", ids)   ' SABOTAZA\n",
        "T_StornoImpact_SchemaDriftJeInvalidan",
        "necitljiva paletna sekcija cini CEO uvid nevalidnim",
    ),
    # Zadat docID koji se ne moze razresiti mora da OBORI uvid. Tihi povratak na
    # poslovni broj vraca tacno ono sto je #198 vadio -- i to unutar modela koji se
    # posle oznacava kao valid, pa nizvodno izgleda kao pouzdan pregled posledica.
    "identitet-degradira-na-broj": (
        "modStornoImpact.bas",
        "                If strict And Len(Trim$(docID)) > 0 Then\n",
        "                If False Then   ' SABOTAZA: identitet pada na broj\n",
        "T_StornoImpact_IdentitetNeDegradira",
        "zadat identitet koji se ne moze razresiti obara uvid, ne pada na broj",
    ),
    # Block sekcija dolazi iz modStornoFlow i tamo je fail-open ziveo jos jednu
    # rundu duze: bez kolone OtkupID spisak blokova ispadne prazan, sto operateru
    # znaci 'nema pogodjenih blokova' -- nad odlukom koja blokove STORNIRA.
    # Block sekcija dolazi iz modStornoFlow i tamo je fail-open ziveo jos jednu
    # rundu duze: bez strict-a GetBlokOtkupIDs proguta drift, vrati prazan skup,
    # GetStornoBlockRows izadje na 'ids.count = 0' PRE svoje kapije -- i uvid
    # zavrsi kao valid sa praznim spiskom. Operateru to znaci 'nema pogodjenih
    # blokova', nad odlukom koja blokove STORNIRA.
    #
    # Sabotaza gadja bas propagaciju, ne kapiju ispod nje: kapija u
    # GetStornoBlockRows se na ovom putu i ne dostigne, pa bi njeno gasenje bilo
    # zeleno-bez-crvenog (zamka 5).
    "uvid-blok-sekcija-guta": (
        "modStornoFlow.bas",
        "            Set ActiveBlocksForFlow = GetBlokOtkupIDs(GetOtpremnicaIDsByBroj(broj, docID), strict)\n",
        "            Set ActiveBlocksForFlow = GetBlokOtkupIDs(GetOtpremnicaIDsByBroj(broj, docID))\n",
        "T_StornoImpact_BlokSekcijaDriftJeInvalidna",
        "necitljiva block sekcija obara CEO uvid",
    ),
    # Err ziv posle uspesne radnje. "On Error Resume Next" prigusuje gresku ali je
    # NE brise, pa prigusena greska iz OtvoriIspravku prezivi povratak i stigne do
    # modUiScreens.ScrEvent, koji je onda prijavi kao 'Radnja nije uspela' -- preko
    # uredno otvorene ispravke. Err.Clear u EH handlerima to ne resava: EH se na
    # uspesnom putu i ne izvrsava.
    "ekran-curi-greska": (
        "modScrStorno.bas",
        "    Scr_Event = ObradiDogadjaj(tag)\n    Err.Clear\n",
        "    Scr_Event = ObradiDogadjaj(tag)   ' SABOTAZA: Err ostaje ziv\n",
        "T_StornoEkran_NeCuriGreska",
        "Scr_Event vraca cist Err -- inace ljuska javi neuspeh za radnju koja je prosla",
    ),
    # Druga i treca grana istog dispecera (zbirna, prijemnica) idu kroz
    # ActiveOtkupIDsByZbirna, gde se strict gubio jos jednu rundu duze nego kod
    # otpremnice. Bez njega drift nad tblOtkup vrati prazan skup, GetStornoBlockRows
    # izadje na 'ids.count = 0' PRE svoje kapije, i uvid zavrsi kao valid.
    "uvid-blok-zbirna-guta": (
        "modStornoFlow.bas",
        "            Set ActiveBlocksForFlow = ActiveOtkupIDsByZbirna(broj, strict)\n",
        "            Set ActiveBlocksForFlow = ActiveOtkupIDsByZbirna(broj)   ' SABOTAZA\n",
        "T_StornoImpact_PrijemnicaBlokDriftJeInvalidan",
        "necitljiva blok sekcija ZBIRNE obara CEO uvid",
    ),
    # Ista rupa, grana prijemnice (preko njene zbirne).
    "uvid-blok-prijemnica-guta": (
        "modStornoFlow.bas",
        "            If Len(bz) > 0 Then Set ActiveBlocksForFlow = ActiveOtkupIDsByZbirna(bz, strict)\n",
        "            If Len(bz) > 0 Then Set ActiveBlocksForFlow = ActiveOtkupIDsByZbirna(bz)\n",
        "T_StornoImpact_PrijemnicaBlokDriftJeInvalidan",
        "necitljiva blok sekcija PRIJEMNICE obara CEO uvid",
    ),
    # Upozorenje uz USPESAN upis mora da nosi oznaku ChrW(10007) -- po njoj
    # CommitDokument odlucuje da li ide i u MsgBox. Bez oznake se tiho gubi: toast
    # sece rep, a uspesan toast se jos i sam sakrije posle cetiri sekunde, pa
    # operater propusti da mu je ostao posao (npr. 'vise ispravki na cekanju').
    "upozorenje-bez-oznake": (
        "modPoruke.bas",
        '    UpsertRow lo, existing, "DOKUNOS_MSG_VISE_ISPRAVKI", ChrW(10007) & " Vi"',
        '    UpsertRow lo, existing, "DOKUNOS_MSG_VISE_ISPRAVKI", " Vi"',
        "T_PorukeUnosa_UpozorenjeNosiOznaku",
        "DOKUNOS_MSG_VISE_ISPRAVKI nosi oznaku upozorenja -- inace se ne vidi",
    ),
    # PkPoIdentitetu je dobio parametar strict, ali ga NIJE koristio: zadata
    # generacija koje nema vracala je prazno, pa je nizvodno izgledala kao
    # 'dokument ne postoji' umesto 'ne mogu da ga razresim' -- a model se posle
    # svega oznacavao kao valid. Komentar iznad koda je tvrdio suprotno od koda.
    "identitet-nestao-prolazi": (
        "modStornoFlow.bas",
        "        If ids.count = 0 Then\n            If strict Then\n",
        "        If ids.count = 0 Then\n            If False Then   ' SABOTAZA\n",
        "T_StornoImpact_NestaoIdentitetJeInvalidan",
        "nestao identitet OTPREMNICE obara uvid",
    ),
    # Ista tvrdnja, grana zbirne: ScanZbirna je prekidao propagaciju strict-a bas
    # na PK resolveru, pa je zbirna prolazila i kad otpremnica nije.
    "zbirna-ne-prosledjuje-strict": (
        "modStornoFlow.bas",
        "                                Array(COL_ZBR_VOZAC, COL_ZBR_KUPAC), strict)\n",
        "                                Array(COL_ZBR_VOZAC, COL_ZBR_KUPAC))   ' SABOTAZA\n",
        "T_StornoImpact_NestaoIdentitetJeInvalidan",
        "nestao identitet ZBIRNE obara uvid",
    ),
    # --- uvid kao kapija (recenzija PR #202) ------------------------------------
    # Uvid je isao po identitetu u zaglavlju, lancu i blokovima, a PALETE po broju.
    # Pod kolizijom broja je ekran tvrdio posledice OBA dokumenta, dok writer
    # nizvodno mutira samo izabrani -- dakle 'ovo su posledice' nije bilo tacno.
    "uvid-palete-po-broju": (
        "modStornoImpact.bas",
        "            Set ids = PrijemniceIDPoIdentitetu(broj, docID)\n",
        "            Set ids = Nothing   ' SABOTAZA: palete se traze po broju\n",
        "T_StornoImpact_PoIdentitetu",
        "sa identitetom uvid nosi SAMO palete izabranog dokumenta",
    ),
    # Red odluke se kesira po DOKUMENTU, ne po stanju podataka. Ako promena podataka
    # ne ponisti kes, vazi odluka izracunata PRE sync-a -- pa dokument koji je u
    # medjuvremenu dobio nizvodni tok i dalje nudi samo 'obican storno'.
    "odluka-prezivi-refresh": (
        "modScrStorno.bas",
        "Public Sub Scr_ResetCache()\n    OcistiIzbor\n",
        "Public Sub Scr_ResetCache()\n    Set mImpact = Nothing   ' SABOTAZA\n",
        "T_StornoAkcije_RefreshInvalidiraOdluku",
        "promena podataka ponistava kes odluke -- inace vazi odluka od pre sync-a",
    ),
    # Ceo smisao ekrana je 'prvo vidi posledice, pa odluci'. Bez ove kapije se
    # mutaciona dugmad nude i kad uvid nije uspeo -- to jest ekran pita isto sto je
    # i MsgBox pitao, samo bez posledica pred sobom.
    "odluka-bez-uvida": (
        "modScrStorno.bas",
        "    If dt <> FLOW_DOC_REVERS Then\n",
        "    If False Then   ' SABOTAZA: odluka se nudi i bez uvida\n",
        "T_StornoBezUvida_NemaAkcije",
        "framework dokument bez uvida ne nudi nijednu radnju",
    ),
    # Posle ispravke je forma bila popunjena a BROJ DOKUMENTA prazan: prefill ga
    # namerno ne donosi (stari broj pripada storniranom), a predlog se nije ni
    # racunao -- RefreshBrojPredlog visi o promeni stanice/datuma, a prefill oba
    # postavlja pod mLoading, pa se nijedan event ne okine.
    "prefill-bez-predloga-broja": (
        "modOtkupUI.bas",
        "    If Not imaBroj Then RefreshBrojPredlog (Not IsTestMode())\n",
        "    ' SABOTAZA: predlog broja se posle prefilla ne racuna\n",
        "T_PrefillBezBroja_PredlaziBroj",
        "prefill bez broja predlaze broj dokumenta za svoj kontekst",
    ),
    # Ljuska ima DVE kapije nad cipovima: MAX_SEG odlucuje da li se cip CRTA, a
    # dispecer klika da li klik na njega ima kome da stigne. Ova sabotaza meri
    # prvu.
    #
    # Druga NEMA sabotazu, i to namerno: test moze da tvrdi da SegIndeksIzTaga
    # razresava poslednji cip, ali ne i da ga dispecer zaista zove -- klik kroz
    # formu se u harnessu ne moze odigrati. Sabotaza nad dispecerom bi zato
    # ostavila suite zelen i lazno tvrdila da je tvrdnja pokrivena (zamka 5).
    # Ta kapija ostaje na smoke-u.
    # Ljuska crta samo prvih MAX_SEG dugmadi prekidaca. Ekran Storno ih ima
    # deset; na devet je "Izvodi" TIHO nestajao -- bio je u Scr_Liste, ali se
    # nije mogao izabrati ni na koji nacin, bez greske i bez traga. Operater je
    # to prijavio kao nedostajuci cip.
    "ljuska-odseca-liste": (
        "modOtkupUI.bas",
        "Private Const MAX_SEG     As Long = 11\n",
        "Private Const MAX_SEG     As Long = 9   ' SABOTAZA\n",
        "T_Storno_UgovorIRadnje",
        "ljuska crta sve liste ekrana -- nijedna se ne odseca tiho",
    ),
    # Navigacioni cip "Svi" je jedino mesto sa kog se dokument trazi kad se ne
    # zna kog je tipa. Legacy ga ima ("Nadji dokument"); bez njega se ekran vraca
    # na "znaj tip pre nego sto pocnes".
    #
    # ZAMENA NIJE PRAZNA, i to je peta zamka ovog kataloga: sabotaza koja BRISE
    # red nema sta da vrati -- `--vrati` trazi zamenu u fajlu, a prazan string se
    # nalazi svuda i nigde, pa tiho ne uradi nista i prijavi "nema sta da se
    # vrati". Kod novog, jos nekomitovanog fajla ni `git checkout` nije mreza.
    # Zato se cip ne brise nego DUPLIRA sa sledecim: broj lista ostaje deset, a
    # pada tvrdnja o kljucevima -- ista poruka, povratan potez.
    #
    # Bez oznake "SABOTAZA" u redu, i to je zamka 4: oznaka bi dosla POSLE `_`,
    # a tamo mora biti kraj reda. Placeno i to jednom -- run je visio do
    # timeout-a, a izlaz je bio "PALO" bez imena tvrdnje. Sirina 40 (umesto 64)
    # cini red jedinstvenim, da `--vrati` ima tacno jedan pogodak.
    "storno-cip-lanac-nestao": (
        "modScrStorno.bas",
        '        ST_LANAC & "|OTKUI_SEG_ST_LANAC|OTKUI_GRID_TITLE_ST_LANAC|76", _\n',
        '        STIP_OTKUP & "|OTKUI_SEG_ST_OTKUP|OTKUI_GRID_TITLE_OTKUP|40", _\n',
        "T_Storno_UgovorIRadnje",
        "redosled i kljucevi lista -- navigaciona je prva",
    ),
    # Prost storno zbirne ne kaskadira, pa prijemnica ostaje vezana za storniranu
    # zbirnu. Bez te poruke operateru sledljivost visi bez upozorenja.
    #
    # NAPOMENA: compile gresku iz istog reda (nekvalifikovan poziv koji zaklanja
    # parametar "poruka") ovaj katalog NE moze da dokaze imenovanom tvrdnjom --
    # takva sabotaza obara COMPILE, pa izlaz bude "Exception occurred" (v. zamka
    # 4). Ono sto test 52 dodaje je da tu proceduru IZVRSAVA: dok je nijedna
    # suite nije zvala, VBA je nije ni kompajlirao.
    "zbirna-poruka-bez-prijemnice": (
        "modStornoDok.bas",
        "                If Len(vezPrij) > 0 Then _\n"
        '                    poruka = modPoruke.Poruka("STORNO_MSG_ZBIRNA_PRIJ") & " " & vezPrij\n',
        "                ' SABOTAZA: poruka ne imenuje vezanu prijemnicu\n",
        "T_StornoIzvrsi_ZbirnaImenujeVezanuPrijemnicu",
        "poruka imenuje prijemnicu koja je ostala vezana",
    ),
    # Spisak blokova za F8 po golom broju: u korpu ulazi i blok drugog dokumenta,
    # a odatle ide pravo u StornoSelectedBlocks_TX.
    "blokovi-po-broju": (
        "modStornoFlow.bas",
        "            Set ActiveBlocksForFlow = GetBlokOtkupIDs(GetOtpremnicaIDsByBroj(broj, docID))\n",
        "            Set ActiveBlocksForFlow = GetBlokOtkupIDs(GetOtpremnicaIDsByBroj(broj))   ' SABOTAZA\n",
        "T_StorniranSibling_ZadrzavaSvojBlok",
        "blok storniranog siblinga je ostao AKTIVAN",
    ),
    # Ista rupa u pregledu: blockCount po broju, pa dijalog nudi tudje blokove.
    "blockcount-po-broju": (
        "modStornoFlow.bas",
        "    Dim allIDs As Collection: Set allIDs = GetOtpremnicaIDsByBroj(broj, gen)\n",
        "    Dim allIDs As Collection: Set allIDs = GetOtpremnicaIDsByBroj(broj)   ' SABOTAZA\n",
        "T_BlokoviF8_PoIdentitetu",
        "pregled broji blokove IZABRANOG dokumenta, ne svih tog broja",
    ),
    # Ispravka ZBIRNE: cilj bez kapije -- zaglavlje dobija zbir tudje dece.
    "zbirna-ispravka-cilj-bez-kapije": (
        "modStornoFlow.bas",
        "    If ZbirnaBrojJeDvosmislenIkad(newBroj) Then\n"
        "        dvosmislen = newBroj: kojaStrana = \"ciljne\"\n",
        "    If False Then   ' SABOTAZA: ciljna strana se ne proverava\n"
        "        dvosmislen = newBroj: kojaStrana = \"ciljne\"\n",
        "T_IspravkaZbirne_KapijaNaObeStrane",
        "dvosmislen CILJ: aktivno zaglavlje nije dobilo zbir tudje dece",
    ),
    # Ispravka ZBIRNE: izvor bez kapije -- sele se deca oba vlasnika broja.
    "zbirna-ispravka-izvor-bez-kapije": (
        "modStornoFlow.bas",
        "    ElseIf ZbirnaBrojJeDvosmislenIkad(oldBroj) Then\n",
        "    ElseIf False Then   ' SABOTAZA: izvorna strana se ne proverava\n",
        "T_IspravkaZbirne_KapijaNaObeStrane",
        "dvosmislen IZVOR: otpremnica nije odseljena sa dvosmislenog broja",
    ),
    # Kapija fail-open na sopstvenu gresku: schema drift -> "jednoznacno je".
    "kapija-fail-open": (
        "modStornoFlow.bas",
        "EH:\n"
        "    LogErr MOD_NAME & \".ZbirnaBrojJeDvosmislenIkad\"\n"
        "    ZbirnaBrojJeDvosmislenIkad = True\n",
        "EH:\n"
        "    LogErr MOD_NAME & \".ZbirnaBrojJeDvosmislenIkad\"\n"
        "    ZbirnaBrojJeDvosmislenIkad = False   ' SABOTAZA: fail-open kapija\n",
        "T_KapijaZbirne_FailClosedNaSvojuGresku",
        "nerazresena jednoznacnost se tretira kao dvosmislena",
    ),
    # Guard koji broji samo AKTIVNE vlasnike. Storniran vlasnik nestaje iz
    # racuna, a njegova aktivna deca ostaju -- pa ih mutacija po broju odvezuje.
    "guard-samo-aktivni-vlasnici": (
        "modStornoFlow.bas",
        '    d("brojDvosmislenIkad") = (VlasniciPoBroju(TBL_ZBIRNA, COL_ZBR_BROJ, broj, _\n'
        "                              MOD_NAME, True, Array(COL_ZBR_VOZAC, COL_ZBR_KUPAC)).count > 1)\n",
        '    d("brojDvosmislenIkad") = (VlasniciPoBroju(TBL_ZBIRNA, COL_ZBR_BROJ, broj, _\n'
        "                              MOD_NAME, False, Array(COL_ZBR_VOZAC, COL_ZBR_KUPAC)).count > 1)\n",
        "T_StorniranVlasnik_JosImaAktivnuDecu",
        "DUPLI staje jer broj je IKAD pripadao dvama vlasnicima",
    ),
    # Zavrsetak ispravke koji ne preveze nijedan blok -- prolazio bi tvrdnju
    # "tudji blok nije pomeren" bez pozitivne kontrole.
    "completion-ne-prevezuje": (
        "modStornoFlow.bas",
        "    Set oldIDs = GetOtpremnicaIDsByBroj(oldBroj, srcGen, srcStanica)\n",
        "    Set oldIDs = New Collection   ' SABOTAZA: nijedan blok se ne prevezuje\n",
        "T_ZavrsetakIspravke_NeDegradiraOldDocID",
        "MOJ blok JESTE prevezan na zamensku otpremnicu",
    ),
    # Zamena zbirne bez kapije: zaglavlje se stornira tacno, a completion posle
    # snimanja zamene odnese decu TUDJE zbirne.
    "zbirna-zamena-bez-kapije": (
        "modStornoFlow.bas",
        "    If mode <> SV_MODE_RESI_KASNIJE Then\n"
        '        If CBool(s("brojDvosmislenIkad")) Then\n',
        "    If False Then   ' SABOTAZA: zamena ide i nad dvosmislenim brojem\n"
        '        If CBool(s("brojDvosmislenIkad")) Then\n',
        "T_ZamenaZbirne_NeDiraDecuTudje",
        "ISPRAVKA staje dok je broj pripadao vise vlasnika",
    ),
    # Zavrsetak ispravke: tacan OldDocID degradiran u prazan opseg -> broj.
    "completion-degradira-olddocid": (
        "modStornoFlow.bas",
        "        If Len(srcGen) = 0 And Len(oldDocID) > 0 Then _\n"
        "            srcStanica = Trim$(NzTx(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, oldDocID, _\n"
        "                                                COL_OTP_STANICA)))\n",
        "        ' SABOTAZA: bez generacije se pada na goli broj\n",
        "T_ZavrsetakIspravke_NeDegradiraOldDocID",
        "blok dokumenta sa druge stanice OSTAJE na svojoj otpremnici",
    ),
    # Otkup bez generacije bez kapije nad brojem. BrojDokumenta je scoped po
    # otkupnom mestu, pa storno po broju hvata i tudje OM.
    "otkup-bez-kapije": (
        "modStorno.bas",
        "    If Len(Trim$(generacijaID)) = 0 Then _\n"
        "        RequireJedanVlasnikPoBroju TBL_OTKUP, COL_OTK_BR_DOK, brDok, SRC, COL_OTK_STANICA\n",
        "    ' SABOTAZA: dvosmislen broj otkupa vise ne zaustavlja storno\n",
        "T_OtkupBezGeneracije_NeStorniraTudjeOM",
        "bez generacije dvosmislen broj otkupa se odbija",
    ),
    # "Jedini vlasnik" po distinct BROJU umesto po dokumentima.
    "sole-owner-po-broju": (
        "modStornoFlow.bas",
        "    If svi.count <> 1 Then Exit Function\n",
        "    If False Then Exit Function   ' SABOTAZA: broji se broj, ne dokument\n",
        "T_SoleOwner_MeriDokumenteNeBrojeve",
        "dve otpremnice istog broja u istoj zbirni NISU jedini vlasnik",
    ),
    # Kaskada zbirne bez fail-closed provere nad dvosmislenim brojem.
    "zbirna-kaskada-bez-kapije": (
        "modStornoFlow.bas",
        "    If VlasniciPoBroju(TBL_ZBIRNA, COL_ZBR_BROJ, brojZbirne, SRC, True, _\n"
        "                       Array(COL_ZBR_VOZAC, COL_ZBR_KUPAC)).count > 1 Then\n"
        '        res("message") = "Broj zbirne \'" & brojZbirne & "\' je pripadao VISE " & _\n',
        "    If False Then   ' SABOTAZA: kaskada ide i nad dvosmislenim brojem\n"
        '        res("message") = "Broj zbirne \'" & brojZbirne & "\' je pripadao VISE " & _\n',
        "T_ZbirnaKaskada_StajeNaDvosmislenom",
        "ponistenje lanca staje dok je broj pripadao vise vlasnika",
    ),
    # Preflight koji primi identitet pa ga ignorise. StornoIzvrsi nize je bio
    # ispravan, ali se do njega nije stizalo -- kapija iznad je odbijala.
    "preflight-ignorise-id": (
        "modStornoDok.bas",
        "            If Len(Trim$(docID)) > 0 Then\n"
        "                If UCase$(Trim$(NzToText(LookupValue(TBL_NOVAC, COL_NOV_ID, docID, _\n",
        "            If False Then   ' SABOTAZA: NovacID se ignorise\n"
        "                If UCase$(Trim$(NzToText(LookupValue(TBL_NOVAC, COL_NOV_ID, docID, _\n",
        "T_Preflight_KoristiIdentitet",
        "sa NovacID-em preflight propusta izabran red",
    ),
    # Kapija nad brojem koja se primenjuje i kad je identitet poznat. Storno je
    # tada bezbedan, ali legitimna ispravka pada -- feature ne radi.
    "kapija-i-uz-identitet": (
        "modStorno.bas",
        "    If Len(Trim$(generacijaID)) = 0 Then _\n"
        "        RequireJedanVlasnikPoBroju TBL_PRIJEMNICA, COL_PRJ_BROJ, brBroj, SRC, COL_PRJ_KUPAC\n",
        "    RequireJedanVlasnikPoBroju TBL_PRIJEMNICA, COL_PRJ_BROJ, brBroj, SRC, COL_PRJ_KUPAC\n",
        "T_IspravkaPrijemnice_PodKolizijomBroja",
        "ISPRAVKA pod kolizijom broja prolazi kad je identitet poznat",
    ),
    # Zaglavlje zbirne po broju umesto po generaciji.
    "zbirna-zaglavlje-po-broju": (
        "modStorno.bas",
        "        If RedJeIzabranogDokumenta(data, i, colBroj, colGenZ, brojZbirne, _\n"
        "                                   generacijaID, SRC) Then\n",
        "        If Trim$(CStr(data(i, colBroj))) = Trim$(brojZbirne) Then   ' SABOTAZA\n",
        "T_Zbirna_ZaglavljePoGeneracijiKaskadaStaje",
        "stornira se SAMO zbirna izabrane generacije",
    ),
    # F8: identitet kliknutog reda. Bez njega correction context pokazuje na
    # prvi dokument tog broja -- a kod RESI KASNIJE se guarded writer uopste ne
    # zove, pa gresku nista ne prijavljuje.
    "f8-identitet-po-broju": (
        "modStornoFlow.bas",
        "    If Len(Trim$(gen)) > 0 Then\n"
        "        Dim ids As Object: Set ids = IdoviGeneracije(tblName, idCol, gen)\n"
        "        ' ZADATA generacija koja se ne razresava je greska, ne poziv na fallback.\n"
        "        If ids.count = 0 Then Exit Function\n"
        "        PkPoIdentitetu = CStr(ids.Keys()(0))\n"
        "        Exit Function\n"
        "    End If\n"
        "\n"
        "    If VlasniciPoBroju(tblName, brojCol, broj, SRC, False, Array(vlasnikCol)).count > 1 Then\n"
        "        Exit Function\n"
        "    End If\n",
        "    ' SABOTAZA: identitet se ignorise -- prvi aktivan red tog broja\n",
        "T_F8_IzabranRedOstajeIzabran",
        "recovery zapis pokazuje na IZABRAN dokument, ne na prvi tog broja",
    ),
    # Kljuc grupisanja u ciljnoj listi kad generacije NEMA (zatecen zapis).
    # Komplementarno sa zbirna-vlasnik-samo-kupac: ta sabotaza dira KOJE kolone
    # cine vlasnika, ova sam kljuc.
    "oporavak-cilj-po-broju": (
        "modScrOporavak.bas",
        "            kljuc = broj & Chr$(1) & vlasnik\n",
        "            kljuc = broj   ' SABOTAZA: dva vlasnika istog broja postaju jedan cilj\n",
        "T_Oporavak_CiljneListe",
        "cilj je DOKUMENT (broj + vlasnik), ne sam broj",
    ),
    # Su-stanar na deljenoj paleti. Dva kupca istog broja i iste robe smeju da
    # dele paletu; poredjenje po broju ih vidi kao istu prijemnicu, pa kapija ne
    # okine i relabel prepravi header cele palete.
    "cotenant-po-broju": (
        "modPaletniList.bas",
        "                    jeIzvor = PripadaDokumentu(bpg, oldBroj, pidG, srcIds, srcDvosmislen)\n"
        "                    jeCilj = PripadaDokumentu(bpg, newBroj, pidG, tgtIds, tgtDvosmislen)\n",
        "                    jeIzvor = (bpg = oldBroj)   ' SABOTAZA\n"
        "                    jeCilj = (bpg = newBroj)\n",
        "T_DeljenaPaleta_SuStanarPoIdentitetu",
        "su-stanar deljene palete je drugi DOKUMENT, ne drugi broj",
    ),
    # Kapija "isti broj" na ulazu u writer, pre razresavanja generacija.
    "writer-isti-broj-odbija": (
        "modPaletniList.bas",
        "    If Len(Trim$(oldGeneracijaID)) > 0 And Len(Trim$(newGeneracijaID)) > 0 Then\n"
        "        If StrComp(Trim$(oldGeneracijaID), Trim$(newGeneracijaID), vbTextCompare) = 0 Then\n",
        "    If False Then   ' SABOTAZA: opet se gleda samo broj\n"
        "        If StrComp(Trim$(oldGeneracijaID), Trim$(newGeneracijaID), vbTextCompare) = 0 Then\n",
        "T_IstiBrojRazliciteGeneracije_NijeIstiDokument",
        "isti broj a razlicite generacije su dva dokumenta i prolaze",
    ),
    # Ciljna lista zbirnih: vlasnistvo je vozac + kupac, ne samo kupac.
    "zbirna-vlasnik-samo-kupac": (
        "modScrOporavak.bas",
        "                                    Array(COL_ZBR_VOZAC, COL_ZBR_KUPAC), _\n",
        "                                    Array(COL_ZBR_KUPAC), _\n",
        "T_Oporavak_CiljneListe",
        "dve zbirne istog broja a razlicitih vozaca ostaju DVA reda",
    ),
    "vlasnik-broji-stornirane": (
        "modStorno.bas",
        "            If ukljuciStornirane Or Not IsStorniranoValue(data(i, cSt)) Then\n",
        "            If True Then   ' SABOTAZA: stornirani se uvek broje kao vlasnici\n",
        "T_Prefill_PoIdentitetuNePoBroju",
        "storniran dokument se ne broji medju AKTIVNIM vlasnicima",
    ),
    # --- ispravka prijemnice od kraja do kraja -------------------------------
    "ispravka-bez-skipa": (
        "modDokUnos.bas",
        "    ispravka = (Len(S(p, \"ispravkaID\")) > 0)\n"
        "    If ispravka Then SetPaletizeSkip True\n",
        "    ispravka = (Len(S(p, \"ispravkaID\")) > 0)\n"
        "    ' SABOTAZA: sveza paletizacija se vise ne preskace\n",
        "T_IspravkaPrijemnice_SkipIRelink",
        "ispravka preskace svezu paletizaciju (inace ista roba ide na dve palete)",
    ),
    "ispravka-bez-relinka": (
        "modDokUnos.bas",
        "    If ispravka Then PreveziPaleteIspravke p, res, poruke\n",
        "    ' SABOTAZA: palete stare prijemnice se vise ne prevezuju\n",
        "T_IspravkaPrijemnice_SkipIRelink",
        "palete stare prijemnice prelaze na novu",
    ),
    "ispravka-context-ostaje": (
        "modDokUnos.bas",
        "            modStornoContext.CompleteCorrectionContext cid, \"\", noviBroj, _\n"
        "                \"Ispravka prijemnice: palete prevezane na \" & noviBroj & \".\"\n",
        "            ' SABOTAZA: correction ostaje PENDING posle uspesnog prevezivanja\n",
        "T_IspravkaPrijemnice_SkipIRelink",
        "correction se zatvara -- inace sledeci unos opet bude ponudjen kao zamena",
    ),
    "ispravka-fail-open": (
        "modDokUnos.bas",
        '        razlog = Poruka("DOKUNOS_MSG_VISE_ISPRAVKI_PRIJ")\n'
        "        NadjiIspravku = -1\n",
        '        razlog = ""   \' SABOTAZA: vise ispravki vise ne zaustavlja upis\n'
        "        NadjiIspravku = 0\n",
        "T_IspravkaDetekcija_FailClosed",
        "dve ispravke na cekanju zaustavljaju upis (safe-stop)",
    ),
    # --- ekran Oporavak -----------------------------------------------------
    "oporavak-registar": (
        "modUiScreens.bas",
        '    c.Add "OPORAVAK|modScrOporavak|OTKUI_NAV_OPORAVAK|" & IC_OPORAVAK & _\n',
        '    c.Add "OPORAVAK|modScrOporavakX|OTKUI_NAV_OPORAVAK|" & IC_OPORAVAK & _\n',
        "T_Oporavak_UgovorIRadnje",
        "ime modula u registru mora da pogadja stvaran modul (kasno vezivanje)",
    ),
    "oporavak-cilj-radnja": (
        "modScrOporavak.bas",
        "        Case \"PRIJEMNICE\"\n"
        "            Scr_Radnje = \"prevezipri:OTKUI_BTN_OPO_PREVEZI:96:soft:1\"\n",
        "        Case \"PRIJEMNICE\", \"ZBIRNE\"   ' SABOTAZA: i ciljna lista dobija dugme\n"
        "            Scr_Radnje = \"prevezipri:OTKUI_BTN_OPO_PREVEZI:96:soft:1\"\n",
        "T_Oporavak_UgovorIRadnje",
        "ciljna lista nema radnju -- dugme bi prevezivalo cilj na samog sebe",
    ),
    "oporavak-stornirani-cilj": (
        "modScrOporavak.bas",
        "        If iSt > 0 Then\n"
        "            If UCase$(modUiData.CellS(src, r, iSt)) = \"DA\" Then GoTo Sledeci\n"
        "        End If\n"
        "        broj = modUiData.CellS(src, r, iBr)\n",
        "        ' SABOTAZA: stornirani dokumenti ulaze u listu ciljeva\n"
        "        broj = modUiData.CellS(src, r, iBr)\n",
        "T_Oporavak_CiljneListe",
        "lista ciljeva nudi SAMO aktivne dokumente",
    ),
    "storno-revers-smer": (
        "modStornoDok.bas",
        "            If Len(Trim$(opcija)) = 0 Then\n"
        "                StornoRazlog = Poruka(\"STORNO_ERR_NEMA_SMERA\")\n"
        "            ElseIf Not ActiveAmbalazaDokExists(broj, opcija) Then\n",
        "            If False Then\n"
        "                StornoRazlog = \"\"   ' SABOTAZA: smer reversa vise nije obavezan\n"
        "            ElseIf Not ActiveAmbalazaDokExists(broj, opcija) Then\n",
        "T_StornoDok_KapijePreUpisa",
        "revers bez smera se odbija -- cetiri smera dele isti brojevni niz",
    ),
    # --- "Odbaci zaostalu ispravku" na ekranu Oporavak ---------------------------
    # Lista Nedovrseno je bila cist pregled: operater vidi da ga safe-stop blokira,
    # a nema cime da to razresi -- jedini izlaz je bila legacy forma.
    "oporavak-nema-odbaci": (
        "modScrOporavak.bas",
        "            Scr_Radnje = \"odbaci:OTKUI_BTN_OPO_ODBACI:150:danger:1\"\n",
        "            Scr_Radnje = \"nista:OTKUI_BTN_OPO_ODBACI:150:danger:1\"   ' SABOTAZA\n",
        "T_Oporavak_UgovorIRadnje",
        "Nedovrseno nudi Odbaci ispravku -- inace je pregled bez izlaza",
    ),
    # Nad istim poslovnim brojem moze da stoji vise contexta (storno, pa opet storno
    # istog dokumenta). Bez CorrectionID-ja u redu, radnja gadja onaj koji zatekne
    # prvi -- a operater je gledao drugi red. Isti razlog zbog kog ekran Storno nosi
    # GeneracijaID u nevidljivoj koloni.
    "oporavak-cid-ne-stize-u-red": (
        "modScrOporavak.bas",
        "        outA(n, 6) = CStr(d(\"correctionID\"))\n",
        "        outA(n, 6) = \"\"   ' SABOTAZA: red nosi samo poslovni broj\n",
        "T_Oporavak_OdbaciIspravku_PoIdentitetu",
        "svaki context red nosi svoj CorrectionID u koloni 6",
    ),
    # Kolona identiteta mora biti prioriteta 4: petlja vidljivosti ide 3 -> 1, pa je
    # 4 jedina vrednost koja je NIKAD ne pokaze. Na prioritetu 1 bi operater dobio
    # kolonu sa internim ID-jem preko pola mreze.
    "oporavak-cid-kolona-vidljiva": (
        "modScrOporavak.bas",
        "        \"OTKUI_HDO_CID||txt|0|4\")\n",
        "        \"OTKUI_HDO_CID||txt|90|1\")   ' SABOTAZA\n",
        "T_Oporavak_OdbaciIspravku_PoIdentitetu",
        "kolona CID je prioriteta 4 -- nikad vidljiva",
    ),
    # Test 71 dokazuje da identitet STIGNE do reda mreze. Ovo je druga tvrdnja:
    # radnja gadja BAS izabrani context. Hard-kodovan ID prolazi 71 netaknut, jer
    # 71 meri transport a ne posledicu.
    "oporavak-odbacuje-prvi-a-ne-izabrani": (
        "modScrOporavak.bas",
        "    OdbaciIspravkuCore = modStornoContext.CancelCorrectionContext(cid, _\n",
        "    OdbaciIspravkuCore = modStornoContext.CancelCorrectionContext(\"SV-TEST-1\", _\n",
        "T_Oporavak_OdbaciIspravku_GasiSamoSvoj",
        "gasi se IZABRANA ispravka -- sused ostaje netaknut",
    ),
    # NED_COL_CID vezuje opis kolona, punjenje reda i radnju u JEDAN broj. Da je
    # radnja imala svoj indeks, drift bi bio nevidljiv: mreza bi izgledala
    # ispravno, a radnja bi citala tudju kolonu.
    "oporavak-cid-kolona-drift": (
        "modScrOporavak.bas",
        "Public Const NED_COL_CID As Long = 6\n",
        "Public Const NED_COL_CID As Long = 5   ' SABOTAZA\n",
        "T_Oporavak_OdbaciIspravku_PoIdentitetu",
        "radnja cita BAS kolonu na kojoj se opis kolona zavrsava",
    ),
    # Zaglavlje palete se od v6-ui-146 cita iz JEDNOG snimka tblPaleta, kroz recnik
    # ID -> red. Ako se red uzme mimo tog recnika, uvid prijavljuje tudju
    # popunjenost i tudju oznaku -- a operater na osnovu toga odlucuje o stornu.
    # Postojeci testovi to ne vide: oni broje palete i sabiraju stavke.
    "palete-zaglavlje-prvi-red": (
        "modPaletniList.bas",
        "        If pIdx.Exists(pid) Then palRow = CLng(pIdx(pid))\n",
        "        If pIdx.Exists(pid) Then palRow = 1   ' SABOTAZA: uvek prvi red\n",
        "T_ImpactPalete_ZaglavljeIzPraveVrste",
        "zaglavlje palete dolazi iz reda BAS te palete",
    ),
    # Efekat posledice se od v6-ui-148 sklapa iz kataloga, sa prefiksom koji se zove
    # ISTO kao dugme odluke. NAPOMENA: slucaj 'kljuc ne postoji u katalogu' ovde
    # NEMA sabotazu -- hvata ga vba_check (provera PORUKA) jos pre nego sto suite
    # krene, pa bi sabotaza pala na tudjoj kapiji i lazno tvrdila da je meri test.
    # Ako se prefiksi spoje i kad se osnovi razlikuju, operater cita da su posledice
    # iste -- a nisu, i bira na osnovu toga.
    "efekat-uvek-spojen-prefiks": (
        "modStornoFlow.bas",
        "    If StrComp(Trim$(dup), Trim$(pon), vbTextCompare) = 0 Then\n",
        "    If True Then   ' SABOTAZA: posledice uvek izgledaju isto\n",
        "T_StornoEfekat_TekstIzKataloga",
        "razlicit efekat nosi OBA prefiksa, ne jedan spojen",
    ),
    # Lista otkupnih blokova radi kao legacy panel: podrazumevano NIJEDAN nije
    # oznacen, oznacen znaci DODATNO storniran. Do v6-ui-149 je nov ekran na
    # potvrdu stornirao SVE -- destruktivnije od legacy-ja, i to slucajno.
    "blokovi-svi-oznaceni": (
        "modScrStorno.bas",
        "        outA(n, 1) = IIf(BlokOznacen(ident), ChrW(10003), \"\")\n",
        "        outA(n, 1) = ChrW(10003)   ' SABOTAZA: sve izgleda oznaceno\n",
        "T_StornoBlokovi_PodrazumevanoNijedan",
        "podrazumevano nijedan blok nije oznacen",
    ),
    # Oznake pripadaju dokumentu nad kojim su napravljene. Ako prezive promenu
    # izbora, sledeci storno gadja blokove koje operater nikad nije video.
    #
    # Zamena nosi oznaku ' SABOTAZA namerno (zamka 8): prva verzija je uklanjala
    # red i ostavljala `mSelTip = ""`, koji postoji i u ZDRAVOM kodu -- pa ga je
    # --vrati nasao tamo i dodao jos jedan `Set mBlokOznaceni = Nothing`.
    "blokovi-oznake-prezive-izbor": (
        "modScrStorno.bas",
        "    Set mBlokOznaceni = Nothing\n",
        "    ' SABOTAZA: oznake prezive promenu izabranog dokumenta\n",
        "T_StornoBlokovi_PodrazumevanoNijedan",
        "promena izabranog dokumenta ponistava oznacene blokove",
    ),
    # Oznaka upozorenja je SIGNAL ZA RUTIRANJE, ne deo recenice: kaze sloju iznad
    # da poruku treba pokazati u dijalogu. MsgBox crta kroz ANSI kodnu stranu u
    # kojoj ChrW(10007) ne postoji, pa ju je operater video kao vodece '?' ispred
    # teksta. U traci poruka, koja je Unicode, ista oznaka OSTAJE.
    "dijalog-nosi-oznaku": (
        "modOtkupUI.bas",
        "    PorukaZaDijalog = Trim$(Replace(txt, ChrW(10007), \"\"))\n",
        "    PorukaZaDijalog = txt   ' SABOTAZA: oznaka ostaje u dijalogu\n",
        "T_PorukeUnosa_UpozorenjeNosiOznaku",
        "poruka u dijalogu ide BEZ oznake za rutiranje",
    ),
    # Red o blokovima u zoni je jedini koji trazi odluku, a odluka se donosi u
    # drugoj listi. Ako ne prati izbor, operater i posle stikliranja cita isti
    # poziv na izbor -- pa ne zna da li je odluka uopste zabelezena.
    "blok-status-ne-prati-izbor": (
        "modScrStorno.bas",
        "    iz = BlokOznacenihBroj()\n",
        "    iz = 0   ' SABOTAZA: izbor se ne vidi u zoni\n",
        "T_StornoBlokovi_PodrazumevanoNijedan",
        "red o blokovima prijavljuje KOLIKO ih je izabrano",
    ),
    # Brojac uz stavku menija ide kroz ugovor, kasno vezano -- ljuska ne sme da
    # sazna nijedan ekran po imenu. Poziv GetNedovrseno direktno bi radio, i to je
    # bas ono sto ceo ugovor izbegava: sledeci ekran koji ima zaostatak morao bi
    # da se doda u ljusku, a ne u svoj modul.
    "brojac-ekran-po-imenu": (
        "modUiScreens.bas",
        "    ScrBrojac = CLng(Application.Run(m & \".Scr_Brojac\"))\n",
        "    If kljuc = \"OPORAVAK\" Then ScrBrojac = 0   ' SABOTAZA: ljuska zna ekran po imenu\n",
        "T_NavBrojac_SamoEkranKojiBroji",
        "ljuska dobija BAS ono sto ekran broji, bez posrednika",
    ),
    # Druga strana: Scr_Brojac je OPCION. Ekran koji ga nema mora da prodje mirno,
    # jer Application.Run na nepostojecu proceduru DIZE gresku -- bez gutanja te
    # greske sidebar se ne bi ni iscrtao.
    "brojac-nije-opcion": (
        "modUiScreens.bas",
        "    If Err.Number <> 0 Then\n        ScrBrojac = 0\n        Err.Clear\n    End If\n",
        "    ' SABOTAZA: greska ekrana bez brojaca se ne guta\n",
        "T_NavBrojac_SamoEkranKojiBroji",
        "ekran bez brojaca daje nulu, ne gresku",
    ),
    # Lista za unos prerade radi kao legacy panel: podrazumevano nijedna paleta
    # nije oznacena, a oznacena znaci DA ULAZI u preradu. Spisak zavrsava u
    # SavePrerada_TX, dakle u mutaciji.
    "prerada-sve-palete": (
        "modScrPalete.bas",
        "        outA(n, 1) = IIf(PalOznacena(ident), ChrW(10003), \"\")\n",
        "        outA(n, 1) = ChrW(10003)   ' SABOTAZA: sve izgleda oznaceno\n",
        "T_NovaPrerada_IzborINeto",
        "podrazumevano nijedna paleta nije oznacena",
    ),
    # Neto je racun, ne unos. Ako se ambalaza ne oduzme, prerada se knjizi sa
    # tezinom kutija i kesa u netu -- greska koja se vidi tek na lageru.
    "prerada-neto-bez-ambalaze": (
        "modScrPalete.bas",
        "    NetoIzracun = bruto - tezPal - amb\n",
        "    NetoIzracun = bruto   ' SABOTAZA: neto je goli bruto\n",
        "T_NovaPrerada_IzborINeto",
        "neto je bruto minus tezina palete",
    ),
    # Dvoklik OTVARA stavke izabrane palete. Bez prebacaja liste operater vidi
    # isti spisak paleta i misli da dvoklik ne radi nista.
    "paleta-dvoklik-ne-otvara": (
        "modScrPalete.bas",
        "    mLista = \"STAVKE\"\n",
        "    mLista = \"PALETE\"   ' SABOTAZA: ostaje na istoj listi\n",
        "T_PaletaDvoklik_OtvaraStavke",
        "dvoklik na paletu otvara njene stavke",
    ),
    # Obican klik samo BIRA. Da i on prebacuje listu, radnje nad redom (zatvori
    # paletu, storniraj, stampaj) postale bi nedostupne -- operater ne bi stigao
    # da ih pritisne.
    "paleta-klik-otvara": (
        "modScrPalete.bas",
        "        PostaviAktivnu CLng(Mid$(tag, 5))\n",
        "        Scr_Event = OtvoriStavke(CLng(Mid$(tag, 5)))   ' SABOTAZA: klik navigira\n",
        "T_PaletaDvoklik_OtvaraStavke",
        "izbor reda ne trazi ponovno citanje mreze",
    ),
    # Zona koje nema je odgovor, ne greska. Ako ScreenZone ostavi Err postavljen,
    # ScrGridData ga procita kao pad ekrana i mreza se isprazni iako su podaci
    # procitani ispravno.
    "zona-curi-gresku": (
        "modOtkupUI.bas",
        "    If Err.Number <> 0 Then Err.Clear\nEnd Function\n",
        "    ' SABOTAZA: Err ostaje postavljen posle zone koje nema\nEnd Function\n",
        "T_PaletaDvoklik_OtvaraStavke",
        "procitana lista se ne prijavljuje kao pad ekrana",
    ),
    # Cip mora STVARNO da suzava listu. Ako filter ne stigne do reda, svi cipovi
    # daju istu punu listu -- operater bira 'Zatvorene' i gleda sve palete.
    "cip-ne-suzava": (
        "modScrPalete.bas",
        "        If Not PalCipProlaz(filter, st, CStr(src(r, 2)), CStr(src(r, 12))) _\n            Then GoTo Sledeci\n",
        "        If False Then GoTo Sledeci   ' SABOTAZA: cip ne suzava\n",
        "T_CipoviEkrana_UgovorIFilter",
        "Otvorene i Zatvorene zajedno daju sve palete",
    ),
    # Bazen cipova je konacan. Ekran koji prijavi vise nego sto bazen ima izgubio
    # bi visak bez ijedne poruke -- cip koji se nikad ne nacrta.
    "bazen-cipova-manji": (
        "modOtkupUI.bas",
        "Public Const MAX_CHIP   As Long = 7\n",
        "Public Const MAX_CHIP   As Long = 4   ' SABOTAZA: bazen manji od opisa\n",
        "T_CipoviEkrana_UgovorIFilter",
        "ekran ne trazi vise cipova nego sto bazen ljuske ima",
    ),
    # Natpis cipa dolazi iz kataloga. Nepostojeci kljuc ne pada nego se ispise
    # kao [KLJUC] -- cip koji operateru nista ne znaci.
    "cip-bez-natpisa": (
        "modScrPalete.bas",
        "                 \"godina:OTKUI_CIPP_GODINA:84|\" & _\n",
        "                 \"godina:OTKUI_CIP_SABOTAZA:84|\" & _\n",
        "T_CipoviEkrana_UgovorIFilter",
        "natpis cipa godina postoji u katalogu",
    ),
    # Panel za unos prerade mora da se UPALI kad je ta lista aktivna. Ako
    # ostane ugasen, operater vidi praznu zonu i nema gde da unese preradu.
    "zona-se-ne-pali": (
        "modScrPalete.bas",
        "    z.Controls(nm).Visible = vis\n",
        "    z.Controls(nm).Visible = False   ' SABOTAZA: panel ostaje ugasen\n",
        "T_ZonaPrerade_SvaPoljaVidljiva",
        "na listi za unos je upaljen ceo panel",
    ),
    # Polje koje se ne NAPRAVI ne moze ni da se upali. Scr_Layout ga tada tiho
    # preskoci (On Error Resume Next), pa je na ekranu rupa bez ijedne poruke.
    "zona-polje-se-ne-pravi": (
        "modScrPalete.bas",
        "    modOtkupUI.NewFieldG z, \"scrPreTezPal\", Poruka(\"OTKUI_PRE_TEZPAL\"), \"txt\", \"kg\", 1, True, False, \"PRE\"\n",
        "    ' SABOTAZA: polje tezine palete se ne pravi\n",
        "T_ZonaPrerade_SvaPoljaVidljiva",
        "panel za unos prerade nema nijednu kontrolu manje",
    ),
    # Neto ulaz je ZBIR neto kilaze izabranih paleta. Bez njega operater ne vidi
    # sa koliko sveze robe ulazi, pa ni da li mu izlaz ima smisla.
    "ulaz-bez-kilaze": (
        "modScrPalete.bas",
        "        mPreNeto(ident) = PalD(src(r, 9))\n",
        "        mPreNeto(ident) = 0   ' SABOTAZA: ulaz je uvek nula\n",
        "T_NovaPrerada_IzborINeto",
        "neto ulaz je zbir neto izabranih paleta",
    ),
    # Bazen dugmadi za radnje nad redom je MAX_ACT. Lista koja prijavi vise njih
    # izgubi visak bez ijedne poruke -- tako je sesta radnja tiho izbacila
    # 'Nepotpune palete'.
    "radnji-vise-nego-dugmadi": (
        "modScrPalete.bas",
        "            Scr_Radnje = \"palprint:OTKUI_BTN_PAL_PRINT:112:ghost:1|\" & _\n",
        "            Scr_Radnje = \"palstavke:OTKUI_BTN_PAL_STAVKE:88:soft:1|\" & _\n                         \"palprint:OTKUI_BTN_PAL_PRINT:112:ghost:1|\" & _\n",
        "T_PaletaDvoklik_OtvaraStavke",
        "lista ne trazi vise radnji nego sto ljuska ima dugmadi",
    ),
    # Cuvar bazena sme da odseca, ali NE sme da cuti. Tiho odsecanje je vec
    # dvaput proslo kroz smoke: jedanaesti cip i sesta radnja nad redom.
    # --- Faza C hardening: identitet, prostor trake, Err ugovor ---------
    # Broj palete i broj prerade se RESETUJU po godini, pa je recnik
    # broj->ID za 12/2025 i 12/2026 imao tacno jedan unos. Ova sabotaza
    # vraca bas to: identitet se trazi po broju, prvi pogodak pobedjuje.
    # Obara OBA testa identiteta -- i palete i prerade.
    "palete-id-po-broju": (
        "modScrPalete.bas",
        "    IdZaRed = Trim$(CStr(modOtkupUI.GridCell(red, kol)))\n",
        "    Dim i As Long, br As String   ' SABOTAZA: lookup samo po broju\n    br = Trim$(CStr(modOtkupUI.GridCell(red, 1)))\n    For i = 1 To modOtkupUI.GridBrojRedova()\n        If Trim$(CStr(modOtkupUI.GridCell(i, 1))) = br Then\n            IdZaRed = Trim$(CStr(modOtkupUI.GridCell(i, kol)))\n            Exit Function\n        End If\n    Next i\n",
        "T_PaleteIdentitet_PoIDNePoBroju",
        "radnja gadja izabrani RED, ne prvi red istog broja",
    ),
    # Telo mreze bez rezerve za traku poruka: poslednji red ulazi 24pt u
    # traku, pa poruka stoji PREKO njega.
    "grid-telo-preko-toasta": (
        "modOtkupUI.bas",
        "    bodyH = zh - (GRID_TOP + GRID_HEAD_H) - GRID_FOOT_H - TOAST_H - 4\n",
        "    bodyH = zh - (GRID_TOP + GRID_HEAD_H) - GRID_FOOT_H - 6   ' SABOTAZA: bez rezerve\n",
        "T_GridTelo_NePokrivaToast",
        "telo mreze staje pre trake poruka",
    ),
    # Vraca BAS zatecen mehanizam: cela funkcija pod Resume Next, bez
    # ciscenja. Greska iznutra se proguta, izvrsavanje se nastavi i Err
    # ostane postavljen -- ljuska tada javlja neuspeh za radnju koja je
    # prosla. (Skidanje samo zavrsnog Err.Clear NE meri nista: ShowToast
    # ima svoj On Error Resume Next, a svaka On Error naredba cisti Err.)
    "palete-event-curi-err": (
        "modScrPalete.bas",
        "    On Error GoTo EH\n    Scr_Event = ObradiDogadjaj(tag)\n    Err.Clear\n",
        "    On Error Resume Next   ' SABOTAZA: greska se guta, Err ostaje\n    Scr_Event = ObradiDogadjaj(tag)\n",
        "T_PaleteScrEvent_NeCuriGreska",
        "Scr_Event ostavlja cist Err i kad je greska obradjena",
    ),
    "bazen-cuti-visak": (
        "modOtkupUI.bas",
        "    If mBazenPrijave.Exists(k) Then Exit Function\n",
        "    Exit Function   ' SABOTAZA: prekoracenje se ne prijavljuje\n",
        "T_BazenLjuske_ViseNegoStoStaje",
        "prekoracenje se prijavljuje",
    ),
    # Odsecanje ide NA velicinu bazena. Odsecanje na nulu bi obrisalo ceo
    # prekidac lista ili celu traku radnji, umesto da izgubi samo visak.
    "bazen-odseca-na-nulu": (
        "modOtkupUI.bas",
        "    BazenStaje = bazen\n",
        "    BazenStaje = 0   ' SABOTAZA: odseca na nulu\n",
        "T_BazenLjuske_ViseNegoStoStaje",
        "visak se odseca na velicinu bazena",
    ),
    # --- agrohemija na novom UI (v6-ui-171) ---------------------------------
    # Ime modula u registru ekrana. Greska u njemu NE PADA: sidebar ekran samo
    # prikaze prigusenog, pa agrohemija nestane iz aplikacije bez ijedne poruke.
    # Isti oblik kao oporavak-modul-ime.
    "agro-modul-ime": (
        "modUiScreens.bas",
        '    c.Add "AGRO|modScrAgro|OTKUI_NAV_AGRO|" & IC_AGRO & _\n',
        '    c.Add "AGRO|modScrAgroX|OTKUI_NAV_AGRO|" & IC_AGRO & _\n',
        "T_Agro_UgovorEkrana",
        "ekran odgovara na Scr_Meta -- kasno vezivanje stvarno razresava modul",
    ),
    # Kapija stanja pri dodavanju u korpu mora da broji i ono sto je VEC u
    # korpi. Bez toga se ista roba doda dva puta preko stanja, a upis pukne tek
    # u petlji i vrati se rollback-om -- operater dobije 4301 umesto recenice.
    "agro-korpa-se-ne-broji": (
        "modAgroUnos.bas",
        '    uKorpi = AgroKorpaKolicina(korpa, artikalID)\n',
        '    uKorpi = 0#   \' SABOTAZA: kapija ne broji ono sto je vec u korpi\n',
        "T_Agro_KapijaStanjaBrojiKorpu",
        "kapija stanja sabira korpu sa novom stavkom",
    ),
    # Druga kapija, pred upis, mora da agregira PO ARTIKLU preko cele korpe.
    # Poredjenje red-po-red propusta korpu koja u zbiru premasuje stanje --
    # tacno scenario "stanje se promenilo izmedju dodavanja i upisa".
    "agro-agregat-po-redu": (
        "modAgroUnos.bas",
        '        treba(artID) = CDbl(treba(artID)) + AD(korpa(i), "kolicina")\n',
        '        treba(artID) = AD(korpa(i), "kolicina")   \' SABOTAZA: bez sabiranja\n',
        "T_Agro_KapijaStanjaBrojiKorpu",
        "kapija pre upisa sabira SVE stavke istog artikla, ne gleda red po red",
    ),
    # Smart doza se zaokruzuje NAGORE: pola pakovanja se ne izdaje. Nanize daje
    # nula pakovanja za 3 l uz pakovanje od 5 l -- predlog bi bio "ne izdaji
    # nista" za robu koja je potrebna.
    # Vidljivost i raspored polja su JEDNA odluka (grana 'izl' u RasporediPolja).
    # Ako se raziju, polje prijema ostane upaljeno u izdavanju -- i sedne tacno
    # preko polja izdavanja, jer oba traze isti slot u redu.
    "agro-rezim-ne-gasi-polja": (
        "modScrAgro.bas",
        '    PoljeVidi z, "scrAgDob", Not izl\n',
        '    PoljeVidi z, "scrAgDob", True   \' SABOTAZA: polje prijema ostaje\n',
        "T_ZonaAgro_PoljaPostojeIPrateRezim",
        "polja prijema su ugasena u izdavanju (i obrnuto)",
    ),
    # Cip koji ne suzava izgleda kao da radi: lista je ista i pre i posle klika.
    # Bojenje bez javljanja nove osnove. clsFlatBtn pamti boju pri Bind-u i na
    # izlazak pokazivaca je vraca; izabran rezim tada pobeli, a natpis ostane
    # krem -- dugme postane necitljivo. Operater je to prijavio na prvom smoke-u.
    # Isti kvar je vec jednom placen u modScrStorno (StilDugmeta).
    # Rez fonta se ne pise unutar "With .Font" -- tamo upis ne hvata i kontrola
    # izlazi bold bez obzira na trazenu vrednost. Kvar je UNIFORMAN, pa se ne
    # vidi golim okom: sve je bold, i izgleda kao odluka dizajna.
    "ljuska-rez-bez-potvrde": (
        "modUiKit.bas",
        "    For i = 1 To 3\n"
        "        If (ctl.Font.Weight >= 700) = bold Then Exit Sub\n"
        "        ctl.Font.bold = bold\n"
        "    Next i\n",
        "    ctl.Font.bold = bold   ' SABOTAZA: rez se upisuje jednom, bez potvrde\n",
        "T_ZonaAgro_PrekidacRezimaZadrzavaBoju",
        "kontrola gradjena bez bold-a stvarno nije bold (Font.Weight 400)",
    ),
    "agro-prekidac-bez-rebase": (
        "modScrAgro.bas",
        '    modOtkupUI.RebaseSink "scrAgSegI"\n'
        '    modOtkupUI.RebaseSink "scrAgSegU"\n',
        "    ' SABOTAZA: nova osnova se ne javlja sink-u\n",
        "T_ZonaAgro_PrekidacRezimaZadrzavaBoju",
        "izabran rezim zadrzava boju i kad pokazivac ode",
    ),
    # Traka korpe pokazuje NAJNOVIJE prvo: operater upravo nesto doda, pa mu je
    # potvrda ono sto trazi. Obrnut redosled izgleda ispravno dok se korpa ne
    # napuni preko cetiri reda.
    "agro-traka-najstarije-prvo": (
        "modScrAgro.bas",
        "        If i > n - 1 Then Exit Function\n"
        "        TrakaRed = KorpaRedPrikaz(k(n - i))\n",
        "        If i > n - 1 Then Exit Function\n"
        "        TrakaRed = KorpaRedPrikaz(k(i + 1))   ' SABOTAZA: najstarije prvo\n",
        "T_Agro_TrakaKorpe_NajnovijePrvoIPreliv",
        "traka pokazuje poslednju dodatu stavku prvu",
    ),
    # Lista koja se tiho odseca izgleda kao cela. Isto pravilo koje ljuska nad
    # sobom vec ima (BazenStaje) -- samo je ovde traka ta koja ne staje.
    "agro-traka-bez-preliva": (
        "modScrAgro.bas",
        '    sakriveno = n - (AG_KORPA_N - 1)\n'
        '    TrakaRed = ChrW(8230) & " " & Poruka("OTKUI_LBL_AG_KORPA_JOS") & " " & sakriveno\n',
        "    sakriveno = n - (AG_KORPA_N - 1)\n"
        "    TrakaRed = KorpaRedPrikaz(k(n - i))   ' SABOTAZA: preliv se ne prijavljuje\n",
        "T_Agro_TrakaKorpe_NajnovijePrvoIPreliv",
        "traka PRIJAVLJUJE koliko stavki nije stalo",
    ),
    # Stavka korpe jos nije u tabeli, pa nema ID iz baze. Ako se trazi po onome
    # sto se u redu VIDI, dve iste stavke ("dva pakovanja sada, dva kasnije")
    # postaju nerazlucive i "Ukloni" izbaci prvu koju nadje -- tiho, jer red koji
    # nestane izgleda isto kao onaj koji je trebalo da nestane.
    "agro-korpa-bez-identiteta": (
        "modAgroUnos.bas",
        '    red("stavkaID") = NovaStavkaId()\n',
        '    red("stavkaID") = "K"   \' SABOTAZA: sve stavke isti identitet\n',
        "T_Agro_KorpaUklanjaPoIdentitetu",
        "svaka stavka korpe nosi SVOJ identitet",
    ),
    # Identitet koji ne stigne do mreze je isto sto i identitet kog nema: ekran
    # ga u trenutku klika nema odakle da procita.
    "agro-identitet-ne-stize-do-mreze": (
        "modScrAgro.bas",
        '        outA(n, 8) = CStr(k(i)("stavkaID"))\n',
        '        outA(n, 8) = ""   \' SABOTAZA: red mreze ne nosi identitet\n',
        "T_Agro_KorpaUklanjaPoIdentitetu",
        "red mreze PRENOSI identitet stavke",
    ),
    # Kolona identiteta je interna. Prioritet 3 je crta, pa operater u korpi
    # gleda sifru koja mu ne znaci nista.
    "agro-identitet-vidljiv": (
        "modScrAgro.bas",
        '        "OTKUI_HDA_STAVKA||txt|1|4")\n',
        '        "OTKUI_HDA_STAVKA||txt|60|3")   \' SABOTAZA: identitet se crta\n',
        "T_Agro_KorpaUklanjaPoIdentitetu",
        "kolona identiteta ostaje van prikaza",
    ),
    # Ljuska brojace pita samo kroz RefreshFromData, a nju zove tek na "podaci su
    # promenjeni". Korpa nije podatak u tabeli, pa bez sopstvenog kanala znacka
    # stoji na nuli dok operater gleda stanje ili dugove i puni korpu.
    "agro-znacka-ne-prati-korpu": (
        "modScrAgro.bas",
        "    mZnacka = Scr_Brojac()\n"
        "    OsveziZonu\n"
        "    modOtkupUI.OsveziNavBrojace\n",
        "    OsveziZonu   ' SABOTAZA: znacka ostaje na staroj vrednosti\n",
        "T_Agro_ZnackaPratiKorpuVanKorpeListe",
        "znacka menija prati korpu i van liste korpe",
    ),
    "agro-cip-ne-suzava": (
        "modScrAgro.bas",
        '        Case "ima":  AgCipStanje = (stanje > 0)\n',
        '        Case "ima":  AgCipStanje = True   \' SABOTAZA: cip ne suzava\n',
        "T_Agro_CipoviSuzavajuListu",
        "cip Ima na stanju stvarno izbacuje artikle bez zaliha",
    ),
    # Lista dugova pokazuje IME, a dvoklik bira KOOPERANTA. Ako mapa na koliziji
    # zapamti prvog pogodjenog umesto praznog, dvoklik izda robu pogresnom
    # coveku -- i izgleda ispravno u svakoj drugoj tvrdnji.
    "agro-dvosmislen-prvi-pobedjuje": (
        "modScrAgro.bas",
        '            If CStr(mDugIds(naziv)) <> koopID Then mDugIds(naziv) = ""\n',
        '            If False Then mDugIds(naziv) = ""   \' SABOTAZA: prvi pobedjuje\n',
        "T_Agro_BrojacIDvoklikPoIdentitetu",
        "dvosmislen prikaz nosi PRAZAN identitet, ne prvog pogodjenog",
    ),
    # Korpa je jedino sto na ovom ekranu ceka operatera. Brojac koji je ne vidi
    # znaci da neproknjizena korpa nestane bez ijednog traga cim se predje na
    # drugi ekran.
    "agro-brojac-ne-vidi-korpu": (
        "modScrAgro.bas",
        "    Scr_Brojac = BrojUKorpi(mKorpaI) + BrojUKorpi(mKorpaU)\n",
        "    Scr_Brojac = 0   ' SABOTAZA: korpa koja ceka se ne vidi\n",
        "T_Agro_BrojacIDvoklikPoIdentitetu",
        "brojac prijavljuje stavke koje cekaju upis",
    ),
    # Bazen ljuske je konacan: visak se TIHO odseca (LayoutChips nacrta prvih
    # MAX_CHIP i stane). Ekran koji trazi vise izgleda ispravno u kodu, a
    # operateru fali dugme.
    "agro-cipova-preko-bazena": (
        "modScrAgro.bas",
        '            Scr_Cipovi = "sve:OTKUI_CHIP_SVE:40|" & _\n'
        '                         "ulaz:OTKUI_CIPA_ULAZ:52|" & _\n'
        '                         "izlaz:OTKUI_CIPA_IZLAZ:52|" & _\n'
        '                         "godina:OTKUI_CIPA_GODINA:84"\n',
        '            Scr_Cipovi = "sve:OTKUI_CHIP_SVE:40|" & _\n'
        '                         "u1:OTKUI_CIPA_ULAZ:52|u2:OTKUI_CIPA_ULAZ:52|" & _\n'
        '                         "u3:OTKUI_CIPA_ULAZ:52|u4:OTKUI_CIPA_ULAZ:52|" & _\n'
        '                         "u5:OTKUI_CIPA_ULAZ:52|izlaz:OTKUI_CIPA_IZLAZ:52|" & _\n'
        '                         "godina:OTKUI_CIPA_GODINA:84"   \' SABOTAZA\n',
        "T_Agro_UgovorEkrana",
        "ekran ne trazi vise cipova nego sto bazen ljuske ima",
    ),
    # Mapa odbitaka sme da bude BRZA kopija GetAgroAbzug, ali ne i DRUGACIJA.
    # Ovde prestaje da SABIRA i pamti poslednji red -- tako se dve kopije
    # istog pravila i razilaze u praksi: lista dugova i kes ekrana pocnu da
    # pokazuju razlicit dug istom coveku.
    "agro-abzug-mapa-ne-sabira": (
        "modNovac.bas",
        "                    d(koopID) = CDbl(d(koopID)) + CDbl(data(i, colUplata))\n",
        "                    d(koopID) = CDbl(data(i, colUplata))   ' SABOTAZA: poslednji pobedjuje\n",
        "T_Agro_AbzugMapaPratiPojedinacni",
        "mapa odbitaka SABIRA redove, isto kao pojedinacni racun",
    ),
    "agro-doza-nanize": (
        "modAgroUnos.bas",
        "    r(\"brojPak\") = CLng(-Int(-dozaKg / pak))\n",
        "    r(\"brojPak\") = CLng(Int(dozaKg / pak))   ' SABOTAZA: nanize\n",
        "T_Agro_SmartDozaZaokruzujeNagore",
        "doza se zaokruzuje NAGORE na cela pakovanja",
    ),
    # --- EKRAN FAKTURISANJE (Faza E/16) -------------------------------------
    # Ljuska Change salje ekranu na SVAKI znak, a GetComboID dok se kuca vraca "".
    # Bez razlike "nerazresen unos" vs "drugi kupac", prvo otkucano slovo baca
    # celu neproknjizenu korpu -- i to a da drugi kupac nije ni izabran.
    "fakture-nerazresen-kupac-brise-korpu": (
        "modScrFakture.bas",
        "    If Len(Trim$(nov)) = 0 Then Exit Function\n",
        "    ' SABOTAZA: prazan ID prolazi kao drugi kupac\n",
        "T_Fak_NerazresenKupacNeDiraKorpu",
        "nerazresen unos u polju kupca ne dira korpu",
    ),
    # LogError pocinje sa `On Error Resume Next`, a svaka On Error naredba brise
    # Err. Citanje Err-a POSLE LogErr-a daje nulu i prazan opis: pozivalac dobije
    # gresku bez razloga, a citac mreze pad seme pretvori u "nema redova".
    # Gadja PrintFaktura, jedan od cetiri popravljena EH bloka -- bas onaj koji
    # zove radnja ekrana "Stampaj".
    "fakture-citac-gubi-gresku": (
        "modFaktura.bas",
        "    errNum = Err.Number\n"
        "    errDesc = Err.description\n"
        "    LogErr \"PrintFaktura\"\n",
        "    LogErr \"PrintFaktura\"   ' SABOTAZA: Err se cita POSLE loga\n"
        "    errNum = Err.Number\n"
        "    errDesc = Err.description\n",
        "T_Fak_GreskaNePreziviLogErr",
        "greska se cita PRE LogErr-a, pa opis prezivi do pozivaoca",
    ),
    # SEF lista se do prvog smoke-a KRILA kad SEF nije podesen. Citanje stanja su
    # kolone tblFakture i ne trazi nikakvu vezu -- kapiju trazi samo RADNJA, i ona
    # je vec ima. Uslovna lista je novi UI cinila UZIM od legacy-ja, koji frmSEF
    # otvara bezuslovno. Ovo vraca tu gresku.
    "fakture-sef-lista-uslovna": (
        "modScrFakture.bas",
        "    Scr_Liste = Array( _\n",
        "    If Not modFaktura.SEFKonfigurisan() Then   ' SABOTAZA: lista se krije\n"
        "        Scr_Liste = Array(FK_ZAFAKT & \"|OTKUI_SEG_FK_ZAFAKT|OTKUI_GRID_TITLE_FK_ZAFAKT|108\")\n"
        "        Exit Function\n"
        "    End If\n"
        "    Scr_Liste = Array( _\n",
        "T_Fak_UgovorEkrana",
        "SEF lista postoji i kad SEF nije podesen",
    ),
    # Lista SEF-a stoji TACNO na MAX_ACT. Sesta radnja se ne prijavljuje kao
    # greska nego se TIHO odseca (RefreshRowActions radi Exit For) -- operater
    # dobije ekran kome fali dugme, bez ijedne poruke.
    "fakture-sef-sesta-radnja": (
        "modScrFakture.bas",
        '                         "sfrecov:OTKUI_BTN_FK_SEF_OPORAVI:88:ghost:1"\n',
        '                         "sfrecov:OTKUI_BTN_FK_SEF_OPORAVI:88:ghost:1|" & _\n'
        '                         "sfvisak:OTKUI_BTN_FK_SEF_STATUS:88:ghost:1"\n',
        "T_Fak_UgovorEkrana",
        "SEF lista ne trazi vise radnji nego sto ljuska ima dugmadi",
    ),
    # Prvi cip je onaj na koji ljuska PADA kad zatecen filter ne pripada listi
    # (RefreshChipsForScreen). Ako nije najsiri, povratak na njega tiho sakrije
    # redove -- operater vidi kracu listu i ne zna zasto.
    "fakture-cip-sve-nije-prvi": (
        "modScrFakture.bas",
        '            FkCipoviZaListu = "sve:OTKUI_CHIP_SVE:40|" & _\n'
        '                                "nepl:OTKUI_CIPF_NEPLACENE:88|" & _\n',
        '            FkCipoviZaListu = "nepl:OTKUI_CIPF_NEPLACENE:88|" & _\n'
        '                                "sve:OTKUI_CHIP_SVE:40|" & _\n',
        "T_Fak_UgovorEkrana",
        "prvi cip svake liste je najsiri ('sve')",
    ),
    # Kolona identiteta je interna. Prioritet 3 je crta, pa operater u listi
    # faktura gleda internu sifru umesto broja.
    "fakture-identitet-vidljiv": (
        "modScrFakture.bas",
        '        "OTKUI_HD_STATUS||paypill|92|1", _\n'
        '        "OTKUI_HDF_FAKID||txt|1|4")\n',
        '        "OTKUI_HD_STATUS||paypill|92|1", _\n'
        '        "OTKUI_HDF_FAKID||txt|90|3")   \' SABOTAZA: identitet se crta\n',
        "T_Fak_IdentitetURedu_NeCrtaSe",
        "kolona identiteta ostaje van prikaza (prioritet 4)",
    ),
    # Dvosmislen ID je ID koji u tabeli postoji dvaput. Ako prvi pobedi, radnja
    # se izvrsi nad redom koji operater NIJE pokazao -- tiho.
    "fakture-dvosmislen-prvi-pobedjuje": (
        "modFaktura.bas",
        "    If CLng(brojac(iD)) <> 1 Then Exit Function\n",
        "    ' SABOTAZA: duplikat prolazi kao identitet\n",
        "T_Fak_IdentitetURedu_NeCrtaSe",
        "ID koji postoji dvaput NIJE identitet",
    ),
    # Dostupnost se cita iz onoga sto RED NOSI, ne iz onoga sto se u njemu VIDI.
    # Prijemnica obelezena kao fakturisana a bez FakturaID ima praznu kolonu
    # fakture -- iz prikaza izgleda slobodna, a CreateFaktura je odbija.
    "fakture-dostupnost-iz-prikaza": (
        "modScrFakture.bas",
        '        outA(n, 11) = IIf(dostupna, "1", "")\n',
        '        outA(n, 11) = IIf(Len(CStr(src(i, 10))) = 0, "1", "")   \' SABOTAZA\n',
        "T_Fak_DostupnostSePrenosiURedu",
        "red mreze prenosi PRAVILO dostupnosti, ne prikaz",
    ),
    # Pravilo zivi na jednom mestu i deli ga kapija IsPrijemnicaAvailableForFaktura
    # sa citacem mreze. Ovde gubi jedan uslov -- i dve strane pocnu da se razilaze.
    "fakture-dostupnost-bez-oznake": (
        "modFaktura.bas",
        '    If Trim$(fakturisano) = "Da" Then Exit Function\n',
        "    ' SABOTAZA: oznaka 'fakturisano' se vise ne gleda\n",
        "T_Fak_DostupnostSePrenosiURedu",
        "obelezena kao fakturisana ne sme u fakturu ni kad FakturaID nedostaje",
    ),
    # Korpa NIJE podatak u tabeli. Ljuska brojace pita samo kroz RefreshFromData,
    # a nju zove tek na "podaci su promenjeni" -- pa bez svog kanala znacka pise
    # nulu dok operater gleda listu faktura i puni korpu.
    "fakture-znacka-ne-prati-korpu": (
        "modScrFakture.bas",
        "    mZnacka = Scr_Brojac()\n"
        "    OsveziZonu\n"
        "    modOtkupUI.OsveziNavBrojace\n",
        "    OsveziZonu   ' SABOTAZA: promena korpe ne stize do znacke\n",
        "T_Fak_KorpaZnackaITraka",
        "promena korpe sama zove OsveziNavBrojace",
    ),
    # "Ukloni" bira po IDENTITETU. Dve stavke istog prikaza (isti broj, ista
    # kolicina, ista cena) su inace nerazlucive, pa nestane pogresna -- tiho,
    # jer red koji nestane izgleda isto kao onaj koji je trebalo da nestane.
    "fakture-korpa-uklanja-prvu": (
        "modScrFakture.bas",
        "    i = UKorpi(prijemnicaID)\n",
        "    i = IIf(Korpa().count > 0, 1, 0)   ' SABOTAZA: uklanja prvu\n",
        "T_Fak_KorpaZnackaITraka",
        "iz korpe se uklanja stavka koju je operater pokazao",
    ),
    # Operater upravo nesto doda, pa mu je potvrda ono sto trazi. Obrnut
    # redosled izgleda ispravno dok se korpa ne napuni preko cetiri reda.
    "fakture-traka-najstarije-prvo": (
        "modScrFakture.bas",
        "    If i < FK_KORPA_N - 1 Then\n"
        "        TrakaRed = KorpaRedPrikaz(n - i)\n",
        "    If i < FK_KORPA_N - 1 Then\n"
        "        TrakaRed = KorpaRedPrikaz(i + 1)   ' SABOTAZA: najstarije prvo\n",
        "T_Fak_KorpaZnackaITraka",
        "traka korpe pokazuje NAJNOVIJE prvo",
    ),
    # Lista koja se tiho odseca izgleda kao cela -- isto pravilo koje ljuska nad
    # sobom vec ima (BazenStaje).
    "fakture-traka-bez-preliva": (
        "modScrFakture.bas",
        '    TrakaRed = ChrW(8230) & " " & Poruka("OTKUI_LBL_AG_KORPA_JOS") & " " & sakriveno\n',
        '    TrakaRed = KorpaRedPrikaz(n - i)   \' SABOTAZA: preliv se precutkuje\n',
        "T_Fak_KorpaZnackaITraka",
        "traka korpe PRIJAVLJUJE koliko stavki ne staje",
    ),
    # Faktura iznosa nula nije placena nego prazna. Da je "placena", cip i znak
    # u istom redu bi tvrdili suprotno jedno od drugog.
    "fakture-prazna-je-placena": (
        "modScrFakture.bas",
        "    If iznos > 0 And uplaceno >= iznos Then\n",
        "    If uplaceno >= iznos Then   ' SABOTAZA: i faktura bez iznosa je placena\n",
        "T_Fak_CipoviPrateStatusFakture",
        "faktura bez iznosa NIJE placena",
    ),
    # Cip "neplacene" mora da primeni ISTA dva uslova kao GetOpenFakture.
    # Bez zapisanog statusa cip pokupi i fakture koje read-model ne vidi, pa se
    # dve implementacije istog pravila raziju.
    "fakture-nepl-ignorise-status": (
        "modScrFakture.bas",
        "            FkCipFaktura = (StrComp(Trim$(status), STATUS_NEPLACENO, vbTextCompare) = 0) _\n"
        "                           And (iznos - uplaceno > 0)\n",
        "            FkCipFaktura = (iznos - uplaceno > 0)   ' SABOTAZA: status se ne gleda\n",
        "T_Fak_CipoviPrateStatusFakture",
        "cip neplacenih se slaze sa modNovac.GetOpenFakture",
    ),
    # Stornirana faktura u listi znaci da joj operater nudi stampu i slanje na SEF.
    "fakture-stornirana-u-listi": (
        "modFaktura.bas",
        "    data = ExcludeStornirano(data, TBL_FAKTURE)\n"
        "    If IsEmpty(data) Then\n"
        "        GetFaktureForGrid = Empty\n",
        "    If IsEmpty(data) Then   ' SABOTAZA: stornirane ostaju u listi\n"
        "        GetFaktureForGrid = Empty\n",
        "T_Fak_CipoviPrateStatusFakture",
        "stornirana faktura ne ulazi u listu",
    ),
    # ---------------------------------------------------------------- BANKA UVOZ
    # Ljuskin FmtDatumKratko odbija sve sto nije IsNumeric, a IsNumeric je nad
    # Date-om FALSE. Datum predat kao Date daje PRAZNU celiju -- bez ijedne
    # greske, bez traga u logu. Nasao ga je tek smoke nad pravim podacima.
    # Ljuskin FmtDatumKratko odbija sve sto nije IsNumeric, a IsNumeric je nad
    # Date-om FALSE. Datum predat kao Date daje PRAZNU celiju -- bez ijedne
    # greske i bez traga u logu. Nasao ga je tek smoke nad pravim podacima.
    "banka-uvoz-datum-nije-broj": (
        "modScrBankaUvoz.bas",
        "        outA(n, 2) = modUiData.CellDate(src, i, 4)\n",
        "        outA(n, 2) = src(i, 4)\n",
        "T_BankaUvoz_UgovorEkrana",
        "datum stize mrezi kao serijski broj, ne kao Date",
    ),
    # Lista stavki stoji TACNO na MAX_ACT. Sesta radnja se ne prijavljuje kao
    # greska nego se TIHO odseca (RefreshRowActions radi Exit For) -- operater
    # dobije ekran kome fali dugme, bez ijedne poruke.
    "banka-uvoz-sesta-radnja": (
        "modScrBankaUvoz.bas",
        '                              "bmsve:OTKUI_BTN_BU_SVE:116:ghost:0"\n',
        '                              "bmsve:OTKUI_BTN_BU_SVE:116:ghost:0|" & _\n'
        '                              "bmvisak:OTKUI_BTN_BU_SKIP:80:ghost:1"\n',
        "T_BankaUvoz_UgovorEkrana",
        "lista stavki ne trazi vise radnji nego sto ljuska ima dugmadi",
    ),
    # Prvi cip je onaj na koji ljuska PADA kad zatecen filter ne pripada listi
    # (RefreshChipsForScreen). Ako nije najsiri, povratak na njega tiho sakrije
    # redove -- operater vidi kracu listu i ne zna zasto.
    "banka-uvoz-cip-sve-nije-prvi": (
        "modScrBankaUvoz.bas",
        '            BuCipoviZaListu = "sve:OTKUI_CHIP_SVE:40|" & _\n'
        '                              "zaobradu:OTKUI_CIPB_ZAOBRADU:80|" & _\n',
        '            BuCipoviZaListu = "zaobradu:OTKUI_CIPB_ZAOBRADU:80|" & _\n'
        '                              "sve:OTKUI_CHIP_SVE:40|" & _\n',
        "T_BankaUvoz_UgovorEkrana",
        "prvi cip svake liste je najsiri ('sve')",
    ),
    # IZVODI su pregled: nijedna operacija se ne radi nad izvodom kao celinom.
    # Radnja koja se tu pojavi trazila bi identitet grupe kao cilj upisa.
    "banka-uvoz-izvodi-imaju-radnju": (
        "modScrBankaUvoz.bas",
        '                              "bmsve:OTKUI_BTN_BU_SVE:116:ghost:0"\n'
        "    End Select\n",
        '                              "bmsve:OTKUI_BTN_BU_SVE:116:ghost:0"\n'
        "        Case BU_IZVODI\n"
        '            BuRadnjeZaListu = "bmauto:OTKUI_BTN_BU_AUTO:112:primary:1"\n'
        "    End Select\n",
        "T_BankaUvoz_UgovorEkrana",
        "lista izvoda nema radnji nad redom",
    ),
    # Kolona identiteta je interna. Prioritet 3 je crta, pa operater u listi
    # stavki gleda internu sifru u dve kolone.
    "banka-uvoz-identitet-vidljiv": (
        "modScrBankaUvoz.bas",
        '        "OTKUI_HDB_BIMKEY||txt|1|4", _\n',
        '        "OTKUI_HDB_BIMKEY||txt|90|3", _\n',
        "T_BankaUvoz_IdentitetURedu_NeCrtaSe",
        "kolona identiteta ostaje van prikaza (prioritet 4)",
    ),
    # Dvosmislen ID je ID koji u tabeli postoji dvaput. Ako prvi pobedi, radnja
    # se izvrsi nad redom koji operater NIJE pokazao -- tiho.
    "banka-uvoz-dvosmislen-prvi-pobedjuje": (
        "modBankaMapiranje.bas",
        "        outA(n, 1) = modFaktura.IdIliPrazno(brojac, Trim$(CStr(data(i, cID))))\n",
        "        outA(n, 1) = Trim$(CStr(data(i, cID)))   ' SABOTAZA: duplikat prolazi\n",
        "T_BankaUvoz_IdentitetURedu_NeCrtaSe",
        "ID koji postoji dvaput NIJE identitet",
    ),
    # Otvorenost se cita iz onoga sto RED NOSI. Nov red ima PRAZAN status, pa se
    # iz prikaza ne razlikuje od reda kome status nije upisan.
    "banka-uvoz-red-ne-nosi-otvorenost": (
        "modScrBankaUvoz.bas",
        '        outA(n, 11) = IIf(CBool(src(i, 10)), "1", "")\n',
        '        outA(n, 11) = "1"   \' SABOTAZA: svaki red izgleda otvoren\n',
        "T_BankaUvoz_RedNosiSmerIOtvorenost",
        "red prenosi otvorenost, radnja je ne izvodi iz prikaza",
    ),
    # Smer se ne izvodi iz toga koja je kolona iznosa popunjena: red sa I
    # uplatom I isplatom izgleda kao uplata, a writer ga odbija.
    "banka-uvoz-red-ne-nosi-smer": (
        "modScrBankaUvoz.bas",
        "        outA(n, 12) = CStr(src(i, 11))\n",
        '        outA(n, 12) = ""   \' SABOTAZA: smer se gubi iz reda\n',
        "T_BankaUvoz_RedNosiSmerIOtvorenost",
        "red prenosi smer stavke",
    ),
    # Zatvorena stavka nema sta da predlozi -- predlog nad njom navodi operatera
    # da pokusa radnju koja ce biti odbijena.
    "banka-uvoz-predlog-i-za-zatvorene": (
        "modScrBankaUvoz.bas",
        "    If Not otvoren Then Exit Function\n",
        "    ' SABOTAZA: predlog se racuna i za zatvorene stavke\n",
        "T_BankaUvoz_RedNosiSmerIOtvorenost",
        "predlog postoji samo za stavke nad kojima jos ima sta da se uradi",
    ),
    # Cip 'jaki kljucevi' i CountStrongKeyReadyBankaImport (koji stoji u natpisu
    # dugmeta) moraju da vide ISTI skup. Pravilo zivi na dva mesta.
    "banka-uvoz-cip-jaki-prolazi-sve": (
        "modScrBankaUvoz.bas",
        '        Case "jaki":       BuCipStavka = modBankaMapiranje.BimOtvoren(s) And jaki\n',
        '        Case "jaki":       BuCipStavka = modBankaMapiranje.BimOtvoren(s)\n',
        "T_BankaUvoz_CipJakihPratiBrojac",
        "cip jakih kljuceva se slaze sa brojacem",
    ),
    # Znacka uz stavku menija broji ono sto CEKA operatera. Bilo koji drugi broj
    # tu izgleda kao posao kog nema (ili sakrije posao koji ima).
    "banka-uvoz-znacka-broji-mapirane": (
        "modScrBankaUvoz.bas",
        "    Scr_Brojac = CLng(k(0))\n",
        "    Scr_Brojac = CLng(k(1))   ' SABOTAZA: znacka broji mapirane\n",
        "T_BankaUvoz_CipJakihPratiBrojac",
        "znacka broji isti skup kao cip 'za obradu'",
    ),
    # 'Obradjeno' i 'preskoceno' su dva razlicita ishoda: prvo je proknjizeno,
    # drugo je svesno ostavljeno. Spojeni cip krije da posao nije zavrsen.
    "banka-uvoz-obradjeno-guta-preskoceno": (
        "modScrBankaUvoz.bas",
        '        Case "obradjeno":  BuCipStavka = (s = BIM_OBR_DA)\n',
        '        Case "obradjeno":  BuCipStavka = (s = BIM_OBR_DA) Or (s = BIM_OBR_SKIP)\n',
        "T_BankaUvoz_CipJakihPratiBrojac",
        "cipovi stanja su razdvojeni i zajedno pokrivaju sve redove",
    ),
    # BROJ IZVODA NIJE IDENTITET: dedupe kljuc pocinje od BROJA RACUNA, pa dva
    # racuna firme legitimno nose izvod istog broja. Grupa bez racuna ih spaja u
    # jedan red i saldo dva razlicita racuna izgleda kao jedan.
    "banka-uvoz-izvod-kljuc-bez-racuna": (
        "modBankaImport.bas",
        '    BimIzvodKljuc = Trim$(brojDokumenta) & "|" & Trim$(brojRacuna)\n',
        "    BimIzvodKljuc = Trim$(brojDokumenta)   ' SABOTAZA: racun ispada iz kljuca\n",
        "T_BankaUvoz_IzvodiSuAgregatPoRacunu",
        "izvod se grupise po (broj + racun), ne samo po broju",
    ),
    # Zbirovi izvoda su isti na SVAKOM redu grupe (parser ih tako upisuje), pa
    # sabiranje daje iznos pomnozen brojem stavki -- i svaki izvod odjednom
    # "ne slaze se".
    "banka-uvoz-saldo-se-sabira": (
        "modBankaImport.bas",
        "        buf(r, 11) = CLng(buf(r, 11)) + 1\n",
        "        buf(r, 5) = CDbl(NzBIM(buf(r, 5), 0#)) + CDbl(NzBIM(data(i, cPoc), 0#))\n"
        "        buf(r, 11) = CLng(buf(r, 11)) + 1\n",
        "T_BankaUvoz_IzvodiSuAgregatPoRacunu",
        "zbirovi izvoda se uzimaju sa reda, ne sabiraju",
    ),
    # Legacy red (uvoz pre v6.18) nema saldo metapodatke -- sva cetiri broja su
    # nula. To NIJE neslaganje nego odsustvo podatka; prikazano kao greska,
    # posalje operatera da trazi kvar kog nema.
    "banka-uvoz-legacy-red-je-razlika": (
        "modBankaImport.bas",
        "        BimSaldoStatus = BIM_SALDO_NEMA\n"
        "        Exit Function\n",
        "        BimSaldoStatus = BIM_SALDO_RAZLIKA   ' SABOTAZA\n"
        "        Exit Function\n",
        "T_BankaUvoz_IzvodiSuAgregatPoRacunu",
        "izvod bez saldo metapodataka nije neslaganje",
    ),
    # Smer-kapija je ista koju RequireBimSmer sprovodi u writeru. OM prima i
    # uplatu i isplatu, ali NE i nejasan smer -- red sa oba iznosa writer odbija.
    "banka-uvoz-om-prima-nejasan-smer": (
        "modBankaMapiranje.bas",
        "            BimSmerOdgovaraTipu = (smer <> BIM_SMER_NEJASAN)\n",
        "            BimSmerOdgovaraTipu = True   ' SABOTAZA: OM prima sve\n",
        "T_BankaUvoz_RucnoMapiranjePravila",
        "smer-kapija ekrana se slaze sa kapijom writera",
    ),
    # Prazan izbor bloka NIJE "nema bloka" nego "uzmi poziv na broj iz izvoda".
    # U formi je prazan combo bio DEFAULT slucaj, pa je blok sa 3+ stavki bez
    # ovog pravila zavrsavao generickom greskom umesto ponudjenom podelom.
    "banka-uvoz-prazan-blok-ostaje-prazan": (
        "modBankaMapiranje.bas",
        "        BimEfektivniBlok = AutoBlockNoForBim(bankaImportID)\n",
        '        BimEfektivniBlok = ""   \' SABOTAZA: poziv na broj se ne koristi\n',
        "T_BankaUvoz_RucnoMapiranjePravila",
        "prazan izbor bloka uzima poziv na broj iz izvoda",
    ),
    # FAIL-CLOSED. Prazna lista faktura i PAD ucitavanja izgledaju isto, a znace
    # suprotno: prazan izbor fakture knjizi AVANS umesto zatvaranja duga.
    "banka-uvoz-fakture-fail-open": (
        "modScrBankaUvoz.bas",
        "    BuSmeMapiranjeKupca = mFaktureOK\n",
        "    BuSmeMapiranjeKupca = True   ' SABOTAZA: pad citanja prolazi\n",
        "T_BankaUvoz_RucnoMapiranjePravila",
        "pad ucitavanja faktura zaustavlja rucno mapiranje kupca",
    ),
    # Lista za rucno mapiranje nudi samo fakture sa OTVORENIM saldom. Zatvorena
    # faktura u toj listi vodi u preplatu.
    "banka-uvoz-fakture-i-zatvorene": (
        "modBankaMapiranje.bas",
        "            otvoreno = GetOtvorenoNaFakturi(fid)\n"
        "            If otvoreno > 0.009 Then\n",
        "            otvoreno = GetOtvorenoNaFakturi(fid)\n"
        "            If otvoreno > -1 Then   ' SABOTAZA: i zatvorene ulaze\n",
        "T_BankaUvoz_RucnoMapiranjePravila",
        "za rucno mapiranje se nude samo fakture sa otvorenim saldom",
    ),
}


def _procitaj(path: str) -> tuple[str, str]:
    """Sadrzaj sa LF krajevima + kraj reda kakav je zatecen na disku."""
    with open(path, "r", encoding="ascii", errors="strict", newline="") as fh:
        raw = fh.read()
    nl = "\r\n" if "\r\n" in raw else "\n"
    return raw.replace("\r\n", "\n"), nl


def _upisi(path: str, tekst: str, nl: str) -> None:
    with open(path, "w", encoding="ascii", newline="") as fh:
        fh.write(tekst.replace("\n", nl))


def _zameni(path: str, staro: str, novo: str) -> tuple[bool, int]:
    """Zameni sidro vezano za pocetak reda. Vraca (uspeh, broj pogodaka)."""
    tekst, nl = _procitaj(path)
    staro, novo = "\n" + staro, "\n" + novo      # zamka 2: sidro od pocetka reda
    pogodaka = tekst.count(staro)
    if pogodaka != 1:
        return False, pogodaka
    _upisi(path, tekst.replace(staro, novo), nl)
    return True, 1


def primeni(ime: str) -> int:
    fajl, staro, novo, test, tvrdnja = SABOTAZE[ime]
    path = os.path.join(SRC_VBA, fajl)

    ok, pogodaka = _zameni(path, staro, novo)
    if not ok:
        razlog = ("sabotaza je vec primenjena" if pogodaka == 0
                  else "sidro nije jednoznacno")
        print(f"sabotaza '{ime}': sidro nadjeno {pogodaka} puta u {fajl}, a mora "
              f"tacno jednom ({razlog}) -- proveri src-vba/{fajl} i sidro u "
              f"tools/sabotaza.py", file=sys.stderr)
        return 2

    print(f"sabotaza '{ime}' primenjena u src-vba/{fajl}")
    print(f"  ocekuj:  FAIL {test}")
    print(f"  tvrdnja: {tvrdnja}")
    print("  pokreni: python tools/run_vba.py --suite RunAllTests")
    print("  vrati:   python tools/sabotaza.py --vrati")
    return 0


def vrati() -> int:
    """Obrnuta zamena, ne git checkout (v. zamka 3)."""
    vraceno = []
    for ime, (fajl, staro, novo, _, _) in SABOTAZE.items():
        ok, _ = _zameni(os.path.join(SRC_VBA, fajl), novo, staro)
        if ok:
            vraceno.append(ime)

    if not vraceno:
        print("nema sta da se vrati -- nijedna sabotaza nije zatecena u src-vba/")
        return 0
    print("vraceno: " + ", ".join(vraceno))
    return 0


def lista() -> int:
    print("Sabotaze (svaka obara TACNO jedan test, po imenu):\n")
    sirina = max(len(k) for k in SABOTAZE)
    for ime, (_, _, _, test, tvrdnja) in SABOTAZE.items():
        print(f"  {ime.ljust(sirina)}  ->  FAIL {test}")
        print(f"  {' ' * sirina}      {tvrdnja}")
    return 0


def main(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(description="Namerno kvarenje koda za dokaz u crvenom smeru")
    ap.add_argument("ime", nargs="?", help="koju sabotazu primeniti")
    ap.add_argument("--lista", action="store_true", help="ispisi sve sabotaze")
    ap.add_argument("--vrati", action="store_true", help="vrati sve zatecene sabotaze")
    args = ap.parse_args(argv)

    if args.lista:
        return lista()
    if args.vrati:
        return vrati()
    if not args.ime:
        ap.print_help()
        return 2
    if args.ime not in SABOTAZE:
        print(f"nepoznata sabotaza '{args.ime}'. Poznate: {', '.join(SABOTAZE)}",
              file=sys.stderr)
        return 2
    return primeni(args.ime)


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
