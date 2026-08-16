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
