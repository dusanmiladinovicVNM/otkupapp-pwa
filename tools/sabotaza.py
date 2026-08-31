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

9. SIDRO ZASTAREVA KAD SE POPRAVI KOD KOJI GADJA. Ispravka po review-u je
   promenila bas onaj uslov na koji je sabotaza bila zakacena, pa je sidro
   prestalo da se nalazi -- a sa njim i dokaz. Sabotaza tada NE javlja
   "test je prosao": javlja da sidro nije nadjeno. To je jedini signal, i
   vidi se samo ako ga gledas: u petlji koja vrti ceo katalog izlaz lako
   prodje kao "sve zeleno". Posle svake izmene koda koji sabotaze gadjaju
   pusti ceo dvosmerni dokaz i tvrdi da je broj CRVENIH jednak broju
   sabotaza -- ne samo da nema neocekivanih padova. Placeno jednom, na
   `banka-uvoz-blok-bez-om-scope`.
"""

import argparse
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ESCN = "\n"
DQ34 = chr(34)
SRC_VBA = os.path.join(ROOT, "src-vba")

# ime -> (fajl, sidro, zamena, test koji MORA da padne, sta tvrdnja kaze)
# Sidro i zamena se porede od POCETKA REDA (v. zamka 2) -- ne pisati vodece \n.
SABOTAZE = {
    # Izbor izvora bez filtera po storniranom. Broj prijemnice je numerisan PO
    # KUPCU, pa isti broj nose dokumenta dva kupca; ono sto storniran tudji
    # dokument drzi na mestu jeste bas taj filter, a ne kapija na ulazu.
    #
    # Ranija verzija ovog unosa je merila SIRU ZABRANU (IKAD kapija na ulazu),
    # a ne pogresnu mutaciju -- i to nad ciljem na kome dokument vec stoji, pa
    # je dokaz bio kruzan. Sada se obara tvrdnja o tome sta se STVARNO pomeri.
    "prijemnica-izvor-i-stornirani": (
        "modDokumenta.bas",
        "        If PripadaIzvoru(data, i, cBrPrij, cPrjId, brPrijemnice, srcIds) Then\n"
        "            If cStorno = 0 Or UCase$(Trim$(CStr(data(i, cStorno)))) <> \"DA\" Then\n"
        "                targetRows.Add i\n"
        "            End If\n"
        "        End If\n",
        "        If PripadaIzvoru(data, i, cBrPrij, cPrjId, brPrijemnice, srcIds) Then\n"
        "            targetRows.Add i   ' SABOTAZA: i storniran tudji dokument ulazi\n"
        "        End If\n",
        "T_Prijemnica_PomeraSamoAktivan",
        "storniran dokument DRUGOG kupca nije pomeren",
    ),
    # Number-only cilj bez IKAD kapije. Kapija po AKTIVNIMA vidi jednog
    # vlasnika i pusta, a storniran vlasnik i dalje ima aktivnu decu -- pa
    # prijemnica zavrsi vezana GOLOM LABELOM za broj koji nose dva vlasnicka
    # toka. Recovery panel u frmDokumenta zove bas ovu putanju, bez ijedne
    # spoljne kapije.
    "cilj-zbirna-kapija-samo-aktivni": (
        "modDokumenta.bas",
        "        RequireJedanVlasnikIkadPoBroju TBL_ZBIRNA, COL_ZBR_BROJ, targetBrZbirne, _\n"
        "                                       SRC, COL_ZBR_VOZAC, COL_ZBR_KUPAC\n",
        "        RequireJedanVlasnikPoBroju TBL_ZBIRNA, COL_ZBR_BROJ, targetBrZbirne, _\n"
        "                                   SRC, COL_ZBR_VOZAC, COL_ZBR_KUPAC   ' SABOTAZA\n",
        "T_CiljZbirna_NePoPrvomRedu",
        "istorijski dvosmislen broj ne prolazi bez generacije",
    ),
    # Bez provere da AKTIVAN cilj postoji, prevezivanje ide i na broj pod kojim
    # su svi redovi stornirani -- IKAD kapija ga ne zaustavlja jer je vlasnik
    # jedan. Ovo je ono sto je zatecena provera prvog reda pokusavala da radi,
    # samo je gledala red koji je SLUCAJNO prvi.
    "cilj-zbirna-bez-provere-postojanja": (
        "modDokumenta.bas",
        "        If aktivniVlasnici = 0 Then Exit Function\n",
        "        ' SABOTAZA: nema provere da aktivan cilj uopste postoji\n",
        "T_CiljZbirna_NePoPrvomRedu",
        "prevezivanje na broj bez ijedne aktivne zbirne ne prolazi",
    ),
    # Postojanje cilja pitano DRUGACIJIM poredjenjem nego kapija ispod:
    # ZbirnaPostoji ide kroz StrComp vbTextCompare, a VlasniciPoBroju poredi
    # tacno. Mala slova bi tada prosla kao "postoji", kapija bi videla NULA
    # vlasnika (hvata samo n > 1), i u prijemnicu bi se upisala labela
    # pozivaoca. To je bila prva verzija ove popravke -- nasla je recenzija.
    "cilj-zbirna-case-mesano": (
        "modDokumenta.bas",
        "        Dim aktivniVlasnici As Long\n"
        "        aktivniVlasnici = VlasniciPoBroju(TBL_ZBIRNA, COL_ZBR_BROJ, _\n"
        "                                          targetBrZbirne, SRC, False, _\n"
        "                                          Array(COL_ZBR_VOZAC, COL_ZBR_KUPAC)).count\n"
        "        If aktivniVlasnici = 0 Then Exit Function\n",
        "        ' SABOTAZA: postojanje poredi drugacije nego kapija ispod\n"
        "        If Not ZbirnaPostoji(targetBrZbirne) Then Exit Function\n",
        "T_CiljZbirna_NePoPrvomRedu",
        "broj sa drugom velicinom slova nije isti broj",
    ),
    # Primitiv koji mutira SVE zbirna redove sa datim brojem, bez kapije u sebi.
    # Zastita je stajala samo po call-site-u (ZbirnaBrojJeDvosmislenIkad, sest
    # mesta u modStornoFlow), pa je nov pozivalac bio bezbedan tek ako se autor
    # kapije seti.
    #
    # Zastavica IKAD nema svoju sabotazu i to je namerno, ne propust: fixture na
    # mestu testa 124 ima IKAD=2 a AKTIVNIH=1, pa bi i okretanje True->False u
    # BrojVlasnikaPoBroju oborilo BAS ovu istu tvrdnju -- dakle zamka 5.
    "zbirna-primitiv-bez-kapije": (
        "modDokumentInvariant.bas",
        "    RequireJedanVlasnikIkadPoBroju TBL_ZBIRNA, COL_ZBR_BROJ, brojZbirne, SRC, _\n"
        "                                   COL_ZBR_VOZAC, COL_ZBR_KUPAC\n",
        "    ' SABOTAZA: primitiv opet mutira po dvosmislenom broju\n",
        "T_RekalkZbirne_KapijaJeUPrimitivu",
        "rekalkulacija po dvosmislenom broju ne prolazi kroz sam primitiv",
    ),
    # Kes indeksa kolone je pamtio i NULU, pa je jedan trenutan neuspeh vazio
    # za ceo BeginTableCache prozor -- svaki sledeci poziv nad istom kolonom
    # dobijao je istu nulu bez novog pokusaja, a RequireColumnIndex na to staje.
    # Odatle "Nedostaje kolona 'VozacID'" nad sveskom u kojoj ta kolona postoji
    # (postmortem par 11). Isto pravilo vec drzi kes TABELA.
    "kes-kolone-pamti-nulu": (
        "modDataAccess.bas",
        "    If Not mColCache Is Nothing Then\n"
        "        If GetColumnIndex > 0 Then mColCache(ck) = GetColumnIndex\n"
        "    End If\n",
        "    If Not mColCache Is Nothing Then\n"
        "        mColCache(ck) = GetColumnIndex   ' SABOTAZA: nula se opet pamti\n"
        "    End If\n",
        "T_KesKolone_NeMemoiseNulu",
        "nula se NE pamti -- trenutan neuspeh ne postaje trajan",
    ),
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
    # DVA PRAVILA, DVE SABOTAZE.
    #
    # Ranija zamena je gadjala CEO poziv TryParseDateValue u ParseDatum, pa je
    # rusila i locale-pravilo i opseg godine odjednom. Padao je prvi po redu
    # (mesec 13), a katalog je deklarisao drugi -- otud PALA DRUGA TVRDNJA.
    #
    # Sada svaka sabotaza gadja svoju kapiju u modParse:
    #   LooksLikeDmyTriple  -- CDate ne sme da "spasava" d.m.y zapis;
    #   opseg godine u DMY  -- 1899 nije poslovni datum.
    "parse-cdate": (
        "modParse.bas",
        "    If LooksLikeDmyTriple(s) Then Exit Function\n",
        "    ' SABOTAZA: CDate opet sme da spasava d.m.y zapis\n",
        "T_ParseDatum_Ugovor",
        "mesec 13 se odbija, ne preliva u sledecu godinu",
    ),
    "parse-godina-opseg": (
        "modParse.bas",
        "    If Y < MIN_POSLOVNA_GODINA Or Y > MAX_POSLOVNA_GODINA Then Exit Function\n",
        "    ' SABOTAZA: godina van poslovnog opsega prolazi\n",
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
    # DVA PRAVILA, DVE ZAMENE nad istim sidrom.
    #
    # Ranija zamena je brisala SVA TRI reda, pa je rusila i zamrzavanje bruta
    # i racun neta. Padalo je zamrzavanje -- prvo po redu -- a katalog je
    # deklarisao racun; otud PALA DRUGA TVRDNJA. Sada svaka gasi jedan red.
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
        "        ' SABOTAZA: bruto se ne zamrzava\n"
        "        kolI = kolI - tara\n"
        '        p("kolicinaI") = kolI\n',
        "T_BrutoNeto_PoRezimu",
        "uneti bruto Kl.I se zamrzava u BrutoKg",
    ),
    "bruto-prijemnica-neto": (
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
        '        p("brutoKgI") = kolI\n'
        "        ' SABOTAZA: tara se ne oduzima\n"
        '        p("kolicinaI") = kolI\n',
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
        "uz blok bez prekidaca isplata je virman firme",
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
        "izabrano otkupno mesto JESTE entitet novca",
    ),
    # --- upis uplate (F6) ---------------------------------------------------
    "uplata-tip-faktura": (
        "modNovacUnos.bas",
        '        p("tipNovca") = NOV_KUPCI_UPLATA\n',
        '        p("tipNovca") = NOV_KUPCI_AVANS   \' SABOTAZA\n',
        "T_UplataValidiraj_FakturaOdlucujeTip",
        "uz fakturu uplata zatvara fakturu",
    ),
    "uplata-preko-fakture": (
        "modNovac.bas",
        "    ElseIf ZaokruziNovac(iznos) > preostalo Then\n"
        "        UplataFakturaProblem = Poruka(\"NOVUNOS_ERR_VECI_OD_FAKTURE\") & \" \" & _\n"
        "                               Format$(preostalo, \"#,##0.00\")\n",
        "    ElseIf False Then   ' SABOTAZA: uplata preko preostalog iznosa fakture ne staje\n"
        "        UplataFakturaProblem = Poruka(\"NOVUNOS_ERR_VECI_OD_FAKTURE\") & \" \" & _\n"
        "                               Format$(preostalo, \"#,##0.00\")\n",
        "T_UplataValidiraj_FakturaOdlucujeTip",
        "preostalog iznosa fakture se ne uplacuje",
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
        "kooperantski smer ne prima kupca",
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
        "ukucan partner bez izbora ne prolazi kao isplata otkupnom mestu",
    ),
    "nerazresen-faktura": (
        "modNovacUnos.bas",
        '    If NerazresenIzbor(S(p, "fakturaTekst"), S(p, "fakturaID")) Then\n'
        '        fokus = "fakturaID": UplataValidiraj = Poruka("NOVUNOS_ERR_FAKTURA_NEIZABRANA"): Exit Function\n'
        "    End If\n",
        "    ' SABOTAZA: ukucana faktura bez izbora opet prolazi kao avans\n",
        "T_NerazresenIzbor_NeProlaziKaoPrazno",
        "ukucana faktura bez izbora ne prolazi kao avans kupca",
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
        "blok drugog kooperanta se odbija",
    ),
    "blok-tudj-om": (
        "modNovac.bas",
        '            If StrComp(Trim$(CStr(data(r, colSt))), Trim$(stanicaID), vbTextCompare) <> 0 Then\n'
        '                IsplataBlokProblem = Poruka("NOVAC_ERR_BLOK_TUDJ_OM") & " " & otkupID\n'
        "                Exit Function\n"
        "            End If\n",
        "            ' SABOTAZA: otkupno mesto bloka se vise ne proverava\n",
        "T_IsplataBlokGuard_VlasnistvoITrenutniOstatak",
        "blok sa drugog otkupnog mesta se odbija",
    ),
    # Podmukliji oblik: kapija ostaje, ali umesto trenutnog stanja veruje
    # vrednosti koju je poslao ekran. Obara TRI testa i to je tacan nalaz --
    # isto pravilo je namerno provereno na tri nivoa (kapija, put unosa, ruta).
    "blok-ostatak-snapshot": (
        "modNovac.bas",
        "    preostalo = vrednost - GetUplataForOtkup(otkupID)\n",
        "    preostalo = vrednost + 1000000   ' SABOTAZA: ostatak se ne cita iz podataka\n",
        "T_IsplataBlokGuard_VlasnistvoITrenutniOstatak",
        "iznos preko trenutnog ostatka se odbija",
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
        "writer odbija blok sa drugog otkupnog mesta i bez UI provere",
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
        "vec placena faktura se odbija (preostalo = 0)",
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
        "writer odbija kes isplatu preko avans salda OM",
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
        "Public Function TabelaTipa(ByVal tk As String) As String\n"
        "    Select Case tk\n",
        "Public Function TabelaTipa(ByVal tk As String) As String\n"
        "    Select Case \"OTPREMNICA\"   ' SABOTAZA: F8 opet svira po jednoj tabeli\n",
        "T_Storno_TipBiraTabeluIKolone",
        "Storno / OTKUP cita svoju tabelu",
    ),
    "f8-tabela-tipa": (
        "modScrDokumenti.bas",
        "        Case \"FAKTURA\":     TabelaTipa = TBL_FAKTURE\n",
        "        ' SABOTAZA: tip fakture ispao iz mape tabela\n",
        "T_Storno_TipBiraTabeluIKolone",
        "Storno / FAKTURA cita svoju tabelu",
    ),
    # --- kapije storna ------------------------------------------------------
    "storno-nema-dok": (
        "modStornoDok.bas",
        "        Case STIP_OTKUP\n"
        "            If Not AktivanPoIdentitetu(TBL_OTKUP, COL_OTK_BR_DOK, COL_OTK_ID, broj, docID) Then _\n"
        "                StornoRazlog = NijePronadjen(broj)\n",
        "        Case STIP_OTKUP\n"
        "            ' SABOTAZA: nepostojeci otkup prolazi kapiju\n"
        "            If False Then StornoRazlog = NijePronadjen(broj)\n",
        "T_StornoDok_KapijePreUpisa",
        "kapija zaustavlja nepostojeci dokument",
    ),
    # --- prefill posle storna (Z10) -----------------------------------------
    "prefill-zbirna-kolona": (
        "modStornoDok.bas",
        "        Case STIP_ZBIRNA:     ColKolicinaZaPrefill = COL_ZBR_KOLICINA\n",
        '        Case STIP_ZBIRNA:     ColKolicinaZaPrefill = "Kolicina"   \' SABOTAZA\n',
        "T_PrefillIzStorniranog_CitaSvojuTabelu",
        "zbirna cita UkupnoKolicina, ne Kolicina",
    ),
    "prefill-tabela": (
        "modStornoDok.bas",
        "        Case STIP_OTKUP:      TabelaZaPrefill = TBL_OTKUP\n",
        "        Case STIP_OTKUP:      TabelaZaPrefill = TBL_OTPREMNICA   ' SABOTAZA\n",
        "T_PrefillIzStorniranog_CitaSvojuTabelu",
        "prefill otkupa nije prazan",
    ),
    "prefill-broj": (
        "modStornoDok.bas",
        '    res = Spoji(res, "fokus", "kolicina")\n',
        '    res = Spoji(res, "brdok", NzToText(d(base, cBroj)))   \' SABOTAZA\n',
        "T_PrefillIzStorniranog_CitaSvojuTabelu",
        "broj dokumenta se NE preuzima",
    ),
    "framework-otkup": (
        "modStornoDok.bas",
        "        Case STIP_OTPREMNICA: TipUFlowDoc = FLOW_DOC_OTPREMNICA\n",
        "        Case STIP_OTPREMNICA, STIP_OTKUP: TipUFlowDoc = FLOW_DOC_OTPREMNICA   ' SABOTAZA\n",
        "T_FrameworkIspravke_SamoCetiriTipa",
        "obican storno, bez framework-a: OTKUP",
    ),
    # --- identitet dokumenta i fail-closed grane (hardening posle review-a) ---
    "prefill-fallback-po-broju": (
        "modDokumenta.bas",
        "        ' to realan scenario, a ne teorijski.\n"
        "        Exit Function\n",
        "        ' SABOTAZA: nepostojeci PK opet pada nazad na broj\n",
        "T_Prefill_PoIdentitetuNePoBroju",
        "nepoznat PK ne pogadja tudji dokument istog broja",
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
        "PK A daje SVOJU kolicinu",
    ),
    # --- identitet dokumenta na granici prevezivanja (zavrsnica Faze D) -----
    "relink-izvor-po-broju": (
        "modPaletniList.bas",
        "            ElseIf PripadaDokumentu(bp, oldBroj, Trim$(CStr(ps(i, sPid))), srcIds, srcDvosmislen) Then\n",
        "            ElseIf bp = oldBroj Then   ' SABOTAZA: izvor se opet bira po broju\n",
        "T_RelinkPoGeneraciji_NeDiraTudjDokument",
        "tudji dokument istog broja OSTAJE na svom mestu",
    ),
    "relink-ignorise-generaciju": (
        "modPaletniList.bas",
        "    Dim srcIds As Object: Set srcIds = IdoviGeneracije(TBL_PRIJEMNICA, COL_PRJ_ID, oldGeneracijaID)\n"
        "    Dim srcDvosmislen As Boolean\n",
        "    Dim srcIds As Object: Set srcIds = IdoviGeneracije(TBL_PRIJEMNICA, COL_PRJ_ID, \"\")      ' SABOTAZA\n"
        "    Dim srcDvosmislen As Boolean\n",
        "T_RelinkPoGeneraciji_NeDiraTudjDokument",
        "prevezivanje po generaciji je proslo",
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
        "roba je stigla na dokument kupca 1 (40 gajbica)",
    ),
    "relink-cilj-bez-kapije": (
        "modPaletniList.bas",
        "    tgtDvosmislen = (VlasniciPoBroju(TBL_PRIJEMNICA, COL_PRJ_BROJ, newBroj, SRC, False, _\n"
        "                                     Array(COL_PRJ_KUPAC)).count > 1)\n",
        "    tgtDvosmislen = False   ' SABOTAZA: dvosmislen cilj vise ne zaustavlja\n",
        "T_RelinkPoGeneraciji_NeDiraTudjDokument",
        "bez generacije CILJA dvosmislen broj se odbija",
    ),
    # Propagacija BrojZbirne u paletne stavke. Izbor redova prijemnice je bio
    # tacan, pa je ovaj drugi upis po BROJU ponistavao ceo taj izbor.
    "zbirna-paleta-po-broju": (
        "modDokumenta.bas",
        "                        pripada = docIds.Exists(pidS)\n",
        "                        pripada = (Trim$(CStr(ps(r2, pBr))) = brPrijemnice)   ' SABOTAZA\n",
        "T_PrevezivanjeNaZbirnu_PaletaIdePoIdentitetu",
        "tudja paleta OSTAJE na staroj zbirni",
    ),
    # Zadata generacija koje nema nije poziv na fallback po broju.
    "generacija-nema-pa-po-broju": (
        "modDokumenta.bas",
        "        If srcIds.count = 0 Then Exit Function\n",
        "        If False Then Exit Function   ' SABOTAZA: pada na broj\n",
        "T_ZadataGeneracijaKojeNema_Staje",
        "zadata generacija prijemnice koje nema zaustavlja upis",
    ),
    # Presuda o relabelu. Writer bira dokument po generaciji; ako presuda opet
    # trazi dokument po broju, opisuje tudji -- i relabel se tiho preskoci.
    "verdikt-po-broju": (
        "modPaletniList.bas",
        "    verdict = PresudiPaletaReassign(oVrS, oSoS, oTaS, nVr, nSo, nTa, oldGajbByKl, newGajb)\n",
        "    verdict = EvaluatePaletaReassign(oldBroj, newBroj)   ' SABOTAZA\n",
        "T_VerdiktPoIdentitetu_RelabelSeNePreskace",
        "stavka je prelabelirana na vrstu ciljnog dokumenta",
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
        # NE deklarise se poslovni ishod, nego PORUKA -- i to je posledica
        # v2.84.0. Otkad RecalculateZbirnaFromOtpremnice_TX nosi kapiju U SEBI,
        # gasenje kapije po call-site-u vise ne menja kolicinu ciljnog
        # zaglavlja: centralna je zaustavi. Ishod time cuvaju DVE kapije, pa ga
        # jedna sabotaza po konstrukciji ne moze oboriti.
        #
        # Razlika se vidi samo u poruci: kapija po call-site-u imenuje CILJNU
        # zbirnu, dok centralna staje iznutra i daje samo neuspeh. Isti oblik
        # kao guard-samo-aktivni-vlasnici.
        "razlog imenuje CILJNU zbirnu, ne staru",
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
        # Prethodna tvrdnja (drift PrijemnicaID) meri modStornoFlow.CountActive,
        # koji ide ranije i pukne prvi -- zato je ova sabotaza bila MRTVA. Drift
        # PaletaID cita jedino paletna sekcija, pa se tek tu meri njena strogost.
        "...i kad nedostaje kolona koju cita SAMO paletna sekcija",
    ),
    # Zadat docID koji se ne moze razresiti mora da OBORI uvid. Tihi povratak na
    # poslovni broj vraca tacno ono sto je #198 vadio -- i to unutar modela koji se
    # posle oznacava kao valid, pa nizvodno izgleda kao pouzdan pregled posledica.
    # DVE KAPIJE, DVE TVRDNJE. Ova sabotaza gadja kapiju u ImpactPalete, a ne onu
    # u modStornoFlow.PkPoIdentitetu -- tu vec cuva 'identitet-nestao-prolazi'.
    #
    # Do sada je bila deklarisana nad PRVOM tvrdnjom testa (generacija koje nema
    # nigde), koju obara BAS PkPoIdentitetu, i to ranije u BuildStornoImpact --
    # pa uklanjanje kapije paleta nije menjalo nista i sabotaza je bila MRTVA.
    #
    # Razlika izmedju dve kapije je stvarna: IdoviGeneracije trazi generaciju
    # kroz celu tabelu, a PrijemniceIDPoIdentitetu trazi broj I generaciju. Zato
    # postoji stanje u kome prva prodje a druga ne -- generacija koja pripada
    # DRUGOM broju -- i tu se meri bas ova kapija.
    "identitet-degradira-na-broj": (
        "modStornoImpact.bas",
        "                If strict And Len(Trim$(docID)) > 0 Then\n",
        "                If False Then   ' SABOTAZA: identitet pada na broj\n",
        "T_StornoImpact_IdentitetNeDegradira",
        "generacija koja pripada DRUGOM broju ne tumaci ovaj dokument",
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
        "Public Const MAX_SEG      As Long = 11\n",
        "Public Const MAX_SEG      As Long = 9   ' SABOTAZA\n",
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
        "            Set ActiveBlocksForFlow = GetBlokOtkupIDs(GetOtpremnicaIDsByBroj(broj, docID), strict)\n",
        "            Set ActiveBlocksForFlow = GetBlokOtkupIDs(GetOtpremnicaIDsByBroj(broj), strict)   ' SABOTAZA\n",
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
        # Deklaracija pomerena na PREVEZIVANJE, ne na rekalkulaciju -- takodje
        # posledica v2.84.0, ali iz drugog razloga nego kod
        # cilj-bez-istorijske-kapije.
        #
        # Centralna kapija stiti REKALKULACIJU (ona je u
        # RecalculateZbirnaFromOtpremnice_TX), pa kolicina ciljnog zaglavlja
        # vise ne mrda. Ali PREVEZIVANJE OTPREMNICE nema svoju centralnu
        # kapiju, pa bez kapije po call-site-u otpremnica izvora STVARNO
        # zavrsi na dvosmislenom cilju -- mereno: ocekivano ZB-TEST-OLDU,
        # dobijeno ZB-TEST-TGT.
        #
        # Nova tvrdnja je zato JACA od stare: opisuje pogresnu mutaciju, ne
        # izostanak jedne. Asimetrija (rekalkulacija ima centralnu kapiju,
        # prevezivanje otpremnice nema) upisana je kao otvoren nalaz.
        "dvosmislen CILJ: otpremnica izvora nije prevezana",
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
    #
    # Ne deklarise se success ("DUPLI staje..."), nego KOJA kapija je stala.
    # Ishod cuvaju DVE nezavisne kapije -- na nivou moda i u detach-u -- pa
    # success ostaje False i kad ova otkaze; jedna sabotaza ga po konstrukciji
    # ne moze oboriti. Razlika se vidi samo u poruci: kapija na nivou moda
    # staje PRE transakcije i kaze razlog, dok bi detach pukao iznutra.
    "guard-samo-aktivni-vlasnici": (
        "modStornoFlow.bas",
        '    d("brojDvosmislenIkad") = (VlasniciPoBroju(TBL_ZBIRNA, COL_ZBR_BROJ, broj, _\n'
        "                              MOD_NAME, True, Array(COL_ZBR_VOZAC, COL_ZBR_KUPAC)).count > 1)\n",
        '    d("brojDvosmislenIkad") = (VlasniciPoBroju(TBL_ZBIRNA, COL_ZBR_BROJ, broj, _\n'
        "                              MOD_NAME, False, Array(COL_ZBR_VOZAC, COL_ZBR_KUPAC)).count > 1)\n",
        "T_StorniranVlasnik_JosImaAktivnuDecu",
        "staje kapija na nivou moda, pre transakcije, sa razlogom",
    ),
    # Zavrsetak ispravke koji ne preveze nijedan blok -- prolazio bi tvrdnju
    # "tudji blok nije pomeren" bez pozitivne kontrole.
    #
    # DVE RAZLICITE TACKE, DVE SABOTAZE. Prazan izvor ID-eva ne stigne do
    # prevezivanja: kapija ga digne kao NERAZRESEN IZVOR, pa completion ne
    # uspe -- otud pada preduslov, a ne poslovna tvrdnja (zamka 6). Zato
    # izvor i prevezivanje imaju svaki svoju sabotazu i svoju tvrdnju.
    "completion-izvor-nerazresen": (
        "modStornoFlow.bas",
        "    Set oldIDs = GetOtpremnicaIDsByBroj(oldBroj, srcGen, srcStanica)\n",
        "    Set oldIDs = New Collection   ' SABOTAZA: izvor se ne razresi\n",
        "T_ZavrsetakIspravke_NeDegradiraOldDocID",
        "zavrsetak ispravke je uspeo",
    ),
    "completion-ne-prevezuje": (
        "modStornoFlow.bas",
        "    Dim blokovi As Collection: Set blokovi = GetBlokOtkupIDs(oldIDs)\n",
        "    Dim blokovi As Collection: Set blokovi = New Collection   ' SABOTAZA: nijedan blok se ne prevezuje\n",
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
        "ISPRAVKA staje dok broj nose dva aktivna dokumenta",
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
        "odbijanje imenuje dvosmislen broj, ne samo neuspeh",
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
        "zbirna drugog vozaca istog broja OSTAJE aktivna",
    ),
    # F8: identitet kliknutog reda. Bez njega correction context pokazuje na
    # prvi dokument tog broja -- a kod RESI KASNIJE se guarded writer uopste ne
    # zove, pa gresku nista ne prijavljuje.
    "f8-identitet-po-broju": (
        "modStornoFlow.bas",
        "    If Len(Trim$(gen)) > 0 Then\n"
        "        Dim ids As Object: Set ids = IdoviGeneracije(tblName, idCol, gen)\n",
        "    If False Then   ' SABOTAZA: identitet se ignorise -- ide se po broju\n"
        "        Dim ids As Object: Set ids = IdoviGeneracije(tblName, idCol, gen)\n",
        "T_F8_IzabranRedOstajeIzabran",
        "recovery zapis pokazuje na IZABRAN dokument",
    ),
    # Kljuc grupisanja u ciljnoj listi kad generacije NEMA (zatecen zapis).
    # Komplementarno sa zbirna-vlasnik-samo-kupac: ta sabotaza dira KOJE kolone
    # cine vlasnika, ova sam kljuc.
    "oporavak-cilj-po-broju": (
        "modScrOporavak.bas",
        "            kljuc = broj & Chr$(1) & vlasnik\n",
        "            kljuc = broj   ' SABOTAZA: dva vlasnika istog broja postaju jedan cilj\n",
        "T_Oporavak_CiljneListe",
        "isti broj zbirne kod dva vozaca daje DVA ciljna dokumenta",
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
        "relabel deljene palete se odbija i uz potvrdu",
    ),
    # Kapija "isti broj" na ulazu u writer, pre razresavanja generacija.
    "writer-isti-broj-odbija": (
        "modPaletniList.bas",
        "    If Len(Trim$(oldGeneracijaID)) > 0 And Len(Trim$(newGeneracijaID)) > 0 Then\n"
        "        If StrComp(Trim$(oldGeneracijaID), Trim$(newGeneracijaID), vbTextCompare) = 0 Then\n",
        "    If False Then   ' SABOTAZA: opet se gleda samo broj\n"
        "        If StrComp(Trim$(oldGeneracijaID), Trim$(newGeneracijaID), vbTextCompare) = 0 Then\n",
        "T_IstiBrojRazliciteGeneracije_NijeIstiDokument",
        "isti broj a razlicite generacije PROLAZI",
    ),
    # Ciljna lista zbirnih: vlasnistvo je vozac + kupac, ne samo kupac.
    "zbirna-vlasnik-samo-kupac": (
        "modScrOporavak.bas",
        "                                    Array(COL_ZBR_VOZAC, COL_ZBR_KUPAC), _\n",
        "                                    Array(COL_ZBR_KUPAC), _\n",
        "T_Oporavak_CiljneListe",
        "isti broj zbirne kod dva vozaca daje DVA ciljna dokumenta",
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
        "nema paletizacije-pa-storna: nijedna stavka nije nastala uzalud",
    ),
    "ispravka-bez-relinka": (
        "modDokUnos.bas",
        "    If ispravka Then PreveziPaleteIspravke p, res, poruke\n",
        "    ' SABOTAZA: palete stare prijemnice se vise ne prevezuju\n",
        "T_IspravkaPrijemnice_SkipIRelink",
        "ispravka nosi prevezenu paletnu stavku",
    ),
    "ispravka-context-ostaje": (
        "modDokUnos.bas",
        "            modStornoContext.CompleteCorrectionContext cid, \"\", noviBroj, _\n"
        "                \"Ispravka prijemnice: palete prevezane na \" & noviBroj & \".\"\n",
        "            ' SABOTAZA: correction ostaje PENDING posle uspesnog prevezivanja\n",
        "T_IspravkaPrijemnice_SkipIRelink",
        "correction context je zatvoren posle uspesnog prevezivanja",
    ),
    "ispravka-fail-open": (
        "modDokUnos.bas",
        '        razlog = Poruka("DOKUNOS_MSG_VISE_ISPRAVKI_PRIJ")\n'
        "        NadjiIspravku = -1\n",
        '        razlog = ""   \' SABOTAZA: vise ispravki vise ne zaustavlja upis\n'
        "        NadjiIspravku = 0\n",
        "T_IspravkaDetekcija_FailClosed",
        "dve ispravke na cekanju zaustavljaju upis",
    ),
    # --- ekran Oporavak -----------------------------------------------------
    "oporavak-registar": (
        "modUiScreens.bas",
        '    c.Add "OPORAVAK|modScrOporavak|OTKUI_NAV_OPORAVAK|" & IC_OPORAVAK & _\n',
        '    c.Add "OPORAVAK|modScrOporavakX|OTKUI_NAV_OPORAVAK|" & IC_OPORAVAK & _\n',
        "T_Oporavak_UgovorIRadnje",
        "modul ekrana odgovara na Scr_Meta (kasno vezivanje radi)",
    ),
    "oporavak-cilj-radnja": (
        "modScrOporavak.bas",
        "        Case \"PRIJEMNICE\"\n"
        "            Scr_Radnje = \"prevezipri:OTKUI_BTN_OPO_PREVEZI:96:soft:1\"\n",
        "        Case \"PRIJEMNICE\", \"ZBIRNE\"   ' SABOTAZA: i ciljna lista dobija dugme\n"
        "            Scr_Radnje = \"prevezipri:OTKUI_BTN_OPO_PREVEZI:96:soft:1\"\n",
        "T_Oporavak_UgovorIRadnje",
        "ciljna lista zbirnih nema radnju",
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
        "stornirana zbirna se NE nudi kao cilj",
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
        "revers bez smera se odbija PRE trazenja dokumenta",
    ),
    # --- "Odbaci zaostalu ispravku" na ekranu Oporavak ---------------------------
    # Lista Nedovrseno je bila cist pregled: operater vidi da ga safe-stop blokira,
    # a nema cime da to razresi -- jedini izlaz je bila legacy forma.
    "oporavak-nema-odbaci": (
        "modScrOporavak.bas",
        "            Scr_Radnje = \"odbaci:OTKUI_BTN_OPO_ODBACI:150:danger:1\"\n",
        "            Scr_Radnje = \"nista:OTKUI_BTN_OPO_ODBACI:150:danger:1\"   ' SABOTAZA\n",
        "T_Oporavak_UgovorIRadnje",
        "Nedovrseno ima Odbaci ispravku",
    ),
    # Nad istim poslovnim brojem moze da stoji vise contexta (storno, pa opet storno
    # istog dokumenta). Bez CorrectionID-ja u redu, radnja gadja onaj koji zatekne
    # prvi -- a operater je gledao drugi red. Isti razlog zbog kog ekran Storno nosi
    # GeneracijaID u nevidljivoj koloni.
    "oporavak-cid-ne-stize-u-red": (
        "modScrOporavak.bas",
        "        outA(n, NED_COL_CID) = CStr(d(\"correctionID\"))\n",
        "        outA(n, NED_COL_CID) = \"\"      ' SABOTAZA: red nosi samo poslovni broj\n",
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
        "SV-TEST-1 ostaje netaknut",
    ),
    # NED_COL_CID vezuje opis kolona, punjenje reda i radnju u JEDAN broj. Da je
    # radnja imala svoj indeks, drift bi bio nevidljiv: mreza bi izgledala
    # ispravno, a radnja bi citala tudju kolonu.
    "oporavak-cid-kolona-drift": (
        "modScrOporavak.bas",
        "Public Const NED_COL_CID As Long = 6\n",
        "Public Const NED_COL_CID As Long = 5   ' SABOTAZA\n",
        "T_Oporavak_OdbaciIspravku_PoIdentitetu",
        "opis kolona se zavrsava BAS na koloni koju radnja cita",
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
        "popunjenost je iz reda BAS te palete",
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
        "razlicit efekat nosi OBA prefiksa u istom redu",
    ),
    # Lista otkupnih blokova radi kao legacy panel: podrazumevano NIJEDAN nije
    # oznacen, oznacen znaci DODATNO storniran. Do v6-ui-149 je nov ekran na
    # potvrdu stornirao SVE -- destruktivnije od legacy-ja, i to slucajno.
    "blokovi-svi-oznaceni": (
        "modScrStorno.bas",
        "        outA(n, 1) = IIf(BlokOznacen(ident), ChrW(10003), \"\")\n",
        "        outA(n, 1) = ChrW(10003)   ' SABOTAZA: sve izgleda oznaceno\n",
        "T_StornoBlokovi_PodrazumevanoNijedan",
        "red 1 nije oznacen bez izricitog izbora",
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
        "promena izbora dokumenta ponistava oznacene blokove",
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
        "DOKUNOS_MSG_VISE_ISPRAVKI u dijalogu ide BEZ oznake",
    ),
    # Red o blokovima u zoni je jedini koji trazi odluku, a odluka se donosi u
    # drugoj listi. Ako ne prati izbor, operater i posle stikliranja cita isti
    # poziv na izbor -- pa ne zna da li je odluka uopste zabelezena.
    "blok-status-ne-prati-izbor": (
        "modScrStorno.bas",
        "    iz = BlokOznacenihBroj()\n",
        "    iz = 0   ' SABOTAZA: izbor se ne vidi u zoni\n",
        "T_StornoBlokovi_PodrazumevanoNijedan",
        "sa izborom red prijavljuje KOLIKO ih je izabrano",
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
    #
    # DVA PRAVILA U ISTOM BLOKU, DVE ZAMENE nad istim sidrom. Ranija zamena
    # je brisala ceo blok, pa je gasila i normalizaciju na nulu i Err.Clear;
    # padalo je curenje Err-a (poslednja tvrdnja), a katalog je deklarisao
    # povratnu vrednost -- otud PALA DRUGA TVRDNJA. Sada svaka gasi jedno.
    "brojac-nije-opcion": (
        "modUiScreens.bas",
        "    If Err.Number <> 0 Then\n        ScrBrojac = 0\n        Err.Clear\n    End If\n",
        "    If Err.Number <> 0 Then\n        ScrBrojac = 0\n"
        "        ' SABOTAZA: greska se guta, ali ostaje POSTAVLJENA\n    End If\n",
        "T_NavBrojac_SamoEkranKojiBroji",
        "poziv ekrana bez brojaca ne ostavlja Err postavljen",
    ),
    # Nula nije slucajna posledica nego deklarisan ugovor: sentinel bi ljuska
    # prikazala kao zaostatak koji ne postoji.
    "brojac-sentinel-umesto-nule": (
        "modUiScreens.bas",
        "    If Err.Number <> 0 Then\n        ScrBrojac = 0\n        Err.Clear\n    End If\n",
        "    If Err.Number <> 0 Then\n"
        "        ScrBrojac = -1   ' SABOTAZA: sentinel umesto nule\n"
        "        Err.Clear\n    End If\n",
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
        "red 1 nije oznacen bez izricitog izbora",
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
    # DODELA IDE U OBRADIDOGADJAJ, ne u Scr_Event.
    #
    # Zamena je do sada pisala `Scr_Event = ...` iz tela ObradiDogadjaj -- a to je
    # dodela imenu TUDJE procedure, dakle compile error. Excel bi stao u [break],
    # suite se ne bi ni pokrenuo, i dokaz.py bi to prijavio kao "NE OBARA NISTA":
    # isto sto i mrtva sabotaza. Sabotaza koja ne kompajlira i sabotaza koja nista
    # ne meri izgledaju identicno, a razlika je velika.
    "paleta-klik-otvara": (
        "modScrPalete.bas",
        "        PostaviAktivnu CLng(Mid$(tag, 5))\n",
        "        ObradiDogadjaj = OtvoriStavke(CLng(Mid$(tag, 5)))   ' SABOTAZA: klik navigira\n",
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
        "stariji red daje SVOJ PaletaID",
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
        "i kad dogadjaj iznutra pukne, Err ne curi kroz Scr_Event",
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
        "modul ekrana Agrohemija odgovara na Scr_Meta (kasno vezivanje radi)",
    ),
    # Kapija stanja pri dodavanju u korpu mora da broji i ono sto je VEC u
    # korpi. Bez toga se ista roba doda dva puta preko stanja, a upis pukne tek
    # u petlji i vrati se rollback-om -- operater dobije 4301 umesto recenice.
    "agro-korpa-se-ne-broji": (
        "modAgroUnos.bas",
        '    uKorpi = AgroKorpaKolicina(korpa, artikalID)\n',
        '    uKorpi = 0#   \' SABOTAZA: kapija ne broji ono sto je vec u korpi\n',
        "T_Agro_KapijaStanjaBrojiKorpu",
        "kapija stanja broji i ono sto je vec u korpi",
    ),
    # Druga kapija, pred upis, mora da agregira PO ARTIKLU preko cele korpe.
    # Poredjenje red-po-red propusta korpu koja u zbiru premasuje stanje --
    # tacno scenario "stanje se promenilo izmedju dodavanja i upisa".
    "agro-agregat-po-redu": (
        "modAgroUnos.bas",
        '        treba(artID) = CDbl(treba(artID)) + AD(korpa(i), "kolicina")\n',
        '        treba(artID) = AD(korpa(i), "kolicina")   \' SABOTAZA: bez sabiranja\n',
        "T_Agro_KapijaStanjaBrojiKorpu",
        "korpa vise ne staje u stanje -- kapija pre upisa to hvata",
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
        "u izdavanju su ugasena polja prijema",
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
        "izdavanje ostaje belo i kad pokazivac ode",
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
        "prvi red trake je POSLEDNJA dodata stavka",
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
        "poslednji red je prelivni, ne cetvrta stavka",
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
        "dve stavke istog prikaza imaju RAZLICIT identitet",
    ),
    # Identitet koji ne stigne do mreze je isto sto i identitet kog nema: ekran
    # ga u trenutku klika nema odakle da procita.
    "agro-identitet-ne-stize-do-mreze": (
        "modScrAgro.bas",
        '        outA(n, 8) = CStr(k(i)("stavkaID"))\n',
        '        outA(n, 8) = ""   \' SABOTAZA: red mreze ne nosi identitet\n',
        "T_Agro_KorpaUklanjaPoIdentitetu",
        "redovi mreze nose razlicite identitete",
    ),
    # Kolona identiteta je interna. Prioritet 3 je crta, pa operater u korpi
    # gleda sifru koja mu ne znaci nista.
    "agro-identitet-vidljiv": (
        "modScrAgro.bas",
        '        "OTKUI_HDA_STAVKA||txt|1|4")\n',
        '        "OTKUI_HDA_STAVKA||txt|60|3")   \' SABOTAZA: identitet se crta\n',
        "T_Agro_KorpaUklanjaPoIdentitetu",
        "kolona identiteta je prioriteta 4 -- mreza je ne crta",
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
        "znacka prati korpu i kad korpa NIJE prikazana lista",
    ),
    "agro-cip-ne-suzava": (
        "modScrAgro.bas",
        '        Case "ima":  AgCipStanje = (stanje > 0)\n',
        '        Case "ima":  AgCipStanje = True   \' SABOTAZA: cip ne suzava\n',
        "T_Agro_CipoviSuzavajuListu",
        "nula nije na stanju",
    ),
    # Lista dugova pokazuje IME, a dvoklik bira KOOPERANTA. Ako mapa na koliziji
    # zapamti prvog pogodjenog umesto praznog, dvoklik izda robu pogresnom
    # coveku -- i izgleda ispravno u svakoj drugoj tvrdnji.
    "agro-dvosmislen-prvi-pobedjuje": (
        "modScrAgro.bas",
        '            If CStr(mDugIds(naziv)) <> koopID Then mDugIds(naziv) = ""\n',
        '            If False Then mDugIds(naziv) = ""   \' SABOTAZA: prvi pobedjuje\n',
        "T_Agro_BrojacIDvoklikPoIdentitetu",
        "dva kooperanta istog imena daju DVOSMISLEN prikaz, ne prvog",
    ),
    # Korpa je jedino sto na ovom ekranu ceka operatera. Brojac koji je ne vidi
    # znaci da neproknjizena korpa nestane bez ijednog traga cim se predje na
    # drugi ekran.
    "agro-brojac-ne-vidi-korpu": (
        "modScrAgro.bas",
        "    Scr_Brojac = BrojUKorpi(mKorpaI) + BrojUKorpi(mKorpaU)\n",
        "    Scr_Brojac = 0   ' SABOTAZA: korpa koja ceka se ne vidi\n",
        "T_Agro_BrojacIDvoklikPoIdentitetu",
        "brojac vidi stavku koja ceka upis",
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
        "lista PROMET ne trazi vise cipova nego sto bazen ima",
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
        "mapa SABIRA odbitke i izuzima stornirane",
    ),
    "agro-doza-nanize": (
        "modAgroUnos.bas",
        "    r(\"brojPak\") = CLng(-Int(-dozaKg / pak))\n",
        "    r(\"brojPak\") = CLng(Int(dozaKg / pak))   ' SABOTAZA: nanize\n",
        "T_Agro_SmartDozaZaokruzujeNagore",
        "3 l trazi JEDNO pakovanje od 5 l",
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
        "prazan ID nije promena kupca -- to je nerazresen unos",
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
        "opis imenuje fakturu koje nema, ne neku drugu gresku",
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
        "pet lista (krug 5: + Utovari), i kad SEF nije podesen",
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
        "SEF ima TACNO MAX_ACT radnji -- sesta bi se tiho odsekla",
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
        "prvi cip liste FAKTURE je najsiri ('sve')",
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
        "lista FAKTURE: kolona identiteta je prioriteta 4",
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
        "red NOSI 'nije dostupna' u koloni 11",
    ),
    # Pravilo zivi na jednom mestu i deli ga kapija IsPrijemnicaAvailableForFaktura
    # sa citacem mreze. Ovde gubi jedan uslov -- i dve strane pocnu da se razilaze.
    "fakture-dostupnost-bez-oznake": (
        "modFaktura.bas",
        '    If Trim$(fakturisano) = "Da" Then Exit Function\n',
        "    ' SABOTAZA: oznaka 'fakturisano' se vise ne gleda\n",
        "T_Fak_DostupnostSePrenosiURedu",
        "obelezena kao fakturisana ne sme -- i kad FakturaID nedostaje",
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
        "znacka prati korpu i kad korpa NIJE prikazana lista",
    ),
    # "Ukloni" bira po IDENTITETU. Dve stavke istog prikaza (isti broj, ista
    # kolicina, ista cena) su inace nerazlucive, pa nestane pogresna -- tiho,
    # jer red koji nestane izgleda isto kao onaj koji je trebalo da nestane.
    "fakture-korpa-uklanja-prvu": (
        "modScrFakture.bas",
        "    i = UKorpi(prijemnicaID)\n",
        "    i = IIf(Korpa().count > 0, 1, 0)   ' SABOTAZA: uklanja prvu\n",
        "T_Fak_KorpaZnackaITraka",
        "ostala je bas ona koja NIJE pokazana",
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
        "prvi red trake je POSLEDNJA dodata stavka",
    ),
    # Lista koja se tiho odseca izgleda kao cela -- isto pravilo koje ljuska nad
    # sobom vec ima (BazenStaje).
    "fakture-traka-bez-preliva": (
        "modScrFakture.bas",
        '    TrakaRed = ChrW(8230) & " " & Poruka("OTKUI_LBL_AG_KORPA_JOS") & " " & sakriveno\n',
        '    TrakaRed = KorpaRedPrikaz(n - i)   \' SABOTAZA: preliv se precutkuje\n',
        "T_Fak_KorpaZnackaITraka",
        "poslednji red trake PRIJAVLJUJE preliv, ne cuti o njemu",
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
        "sa statusom Placeno ne prolazi ni sa ostatkom",
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
        "stornirana faktura NIJE u listi",
    ),
    # ------------------------------------------------------------ MREZA: CELIJA
    # RenderGrid radi pod 'On Error Resume Next'. Dok se tekst racunao U SAMOM
    # UPISU, pad konverzije je preskakao UPIS -- pa je u celiji ostajao natpis od
    # ranijeg crtanja, vrednost sa PRETHODNOG EKRANA. Ova sabotaza vraca tacno
    # taj oblik: prazan rezultat ne prepisuje staru vrednost.
    "mreza-celija-prazno-ne-prepisuje": (
        "modOtkupUI.bas",
        "                                .caption = txt\n",
        "                                If Len(txt) > 0 Then .caption = txt   ' SABOTAZA: prazno ne prepisuje\n",
        "T_MrezaCelija_NeostavljaTudjiTekst",
        "celija koja se ne moze prikazati OSTAJE PRAZNA -- ne zadrzava tudji tekst",
    ),
    # 'IsNumeric' nad Date-om je FALSE. Ekran koji vrednost preda onakvu kakva u
    # tabeli jeste dobijao je praznu celiju -- lista FAKTURA je tako imala prazan
    # datum u SVAKOM redu, i to se nije videlo jer nijedan test nije citao
    # NACRTAN datum.
    # STIL KOLONE JE LAYOUT-OV POSAO. Prva verzija ovog PR-a je pred svaki upis
    # 'vracala celiju u neutralno' (levo poravnanje, bez bold-a) -- i time na SVAKOM
    # ekranu obarala ono sto je LayoutGrid upravo postavio: brojevi bi presli levo,
    # a prva kolona i kolone novca izgubile bold. Nijedna tvrdnja o natpisu to ne bi
    # primetila, pa suite ostaje zelena. Sabotaza vraca tacno taj oblik.
    "mreza-crtanje-kvari-stil": (
        "modOtkupUI.bas",
        "                                txt = CelijaTekst(ColKind(k), mView(r, k + 1), celijaOK)\n",
        "                                txt = CelijaTekst(ColKind(k), mView(r, k + 1), celijaOK): .TextAlign = fmTextAlignLeft   ' SABOTAZA: crtanje kvari stil\n",
        "T_MrezaCelija_NeostavljaTudjiTekst",
        "preduslov: brojcana kolona je poravnata DESNO",
    ),
    # Pilula koja se ne moze naslikati mora da NESTANE. PaintPill menja i pozadinu,
    # boju, sirinu i BackStyle, a PaintRow pill kolone pri vracanju pozadine reda
    # NAMERNO preskace -- pa bi celija kojoj natpis nije obrisan ostala kao stara
    # obojena oznaka nad novim podatkom.
    # DVE VRSTE PILULE, DVA UGOVORA. Pravoj ("pill") sirinu racuna PaintPill i
    # LayoutGrid je preskace; "paypill" sirinu drzi LayoutGrid (mColW - 16), a
    # PaintPayPill je ne vraca. Ko "paypill" ocisti kao pravu pilulu, postavi joj
    # PUNU sirinu kolone -- i ona takva ostane, jer se LayoutGrid ponovo pusta tek
    # kad se promeni opis kolona. Tacno ta greska je jednom vec napravljena ovde.
    "mreza-paypill-kao-pill": (
        "modOtkupUI.bas",
        "                                    .caption = vbNullString\n",
        "                                    OcistiPilulu body.Controls(\"c\" & i & \"_\" & k), mColW(k)   ' SABOTAZA: paypill se cisti kao pill\n",
        "T_MrezaCelija_NeostavljaTudjiTekst",
        "ciscenje statusne oznake ne menja sirinu celije",
    ),
    "mreza-pilula-ostaje": (
        "modOtkupUI.bas",
        "                                    .caption = vbNullString\n",
        "                                    n = n   ' SABOTAZA: stara oznaka ostaje\n",
        "T_MrezaCelija_NeostavljaTudjiTekst",
        "statusna oznaka koja se ne moze naslikati nestaje",
    ),
    "mreza-datum-nije-date": (
        "modOtkupUI.bas",
        "    If IsDate(v) Then\n",
        "    If IsDate(v) And False Then   ' SABOTAZA: Date se vise ne prima\n",
        "T_MrezaCelija_NeostavljaTudjiTekst",
        "prava Date vrednost se crta -- IsNumeric je nad njom False",
    ),
    # Prazna celija je istina, ali TIHA prazna celija je bila pola problema:
    # prvi nalaz ove vrste trazio je zasebnu dijagnostiku da bi se uopste video.
    "mreza-kvar-celije-se-ne-broji": (
        "modOtkupUI.bas",
        "                            mKvarCelija = mKvarCelija + 1\n",
        "                            mKvarCelija = mKvarCelija + 0   ' SABOTAZA: kvar se ne broji\n",
        "T_MrezaCelija_NeostavljaTudjiTekst",
        "kvar prikaza se broji, pa ostaje trag u logu",
    ),
    # ------------------------------------------- BANKA: LEGACY FORMA
    # Prva sabotaza nad .frm fajlom. Pravila ove forme odlucuju hoce li uplata
    # postati avans, a do sada su bila proverljiva samo rukom -- pa je ista greska
    # tri puta prosla kroz review umesto kroz suite.
    "banka-legacy-pad-liste-prolazi": (
        "frmBankaImport.frm",
        "    If Not m_BlokoviLoadOk Then\n",
        "    If Not m_BlokoviLoadOk And False Then   ' SABOTAZA: pad ucitavanja prolazi\n",
        "T_LegacyBanka_PadUcitavanjaNijePraznaLista",
        "pad ucitavanja liste blokova zaustavlja rucno mapiranje",
    ),
    # Prazan combo NIJE izbor: tada blok dolazi iz poziva na broj, gde je avans
    # legitiman. Ko to izjednaci, ili ugasi legitimnu granu ili prijavi writeru
    # izbor kog nije bilo.
    "banka-legacy-prazan-combo-je-izbor": (
        "frmBankaImport.frm",
        '    ManualBlokIzabran = (Trim$(nz(cmbOtkupBlok.value, "")) <> "")\n',
        '    ManualBlokIzabran = True   \' SABOTAZA: prazan combo je izbor\n',
        "T_LegacyBanka_PadUcitavanjaNijePraznaLista",
        "prazan combo NIJE izbor -- tada blok dolazi iz poziva na broj",
    ),
    # ------------------------------------- MATICNI: SEKCIJA I POKRETAC
    # Sidebar nema skrol. Kad stavke predju slobodnu visinu, ne skroluju se nego
    # TIHO nestanu ispod profila -- zato sekcije uopste postoje. Cetiri sabotaze:
    # da mera stvarno meri, da prekidac stvarno gasi drugu sekciju, da radnja
    # bira alatku po identitetu, i da ljuska pita ekran za njegovu branu.
    "maticni-sifarnici-u-radnoj-sekciji": (
        "modUiScreens.bas",
        "                      \"SIFARNICI|OTKUI_NAVG_SIFARNICI|\" & SEK_MATICNI, _\n",
        "                      \"SIFARNICI|OTKUI_NAVG_SIFARNICI|\" & SEK_RAD, _\n",
        "T_Sekcija_SidebarNeStajeZajedno",
        "radna sekcija staje u sidebar",
    ),
    "maticni-sidebar-ne-gasi-drugu-sekciju": (
        "modOtkupUI.bas",
        "    z.Controls(nm).top = Y\n    z.Controls(nm).Visible = vis\n",
        "    z.Controls(nm).top = Y\n    z.Controls(nm).Visible = True   ' SABOTAZA: stavka druge sekcije ostaje\n",
        "T_Sekcija_SidebarNeStajeZajedno",
        "u maticnoj sekciji je radna stavka ugasena",
    ),
    "maticni-prigusena-stavka-se-preboji": (
        "modOtkupUI.bas",
        "                IIf(on_, C_CREAM, IIf(off_, C_DISABLED_FG, RGB(52, 68, 44)))\n",
        "                IIf(on_, C_CREAM, RGB(52, 68, 44))   ' SABOTAZA: prigusenost nestaje\n",
        "T_Sekcija_SidebarNeStajeZajedno",
        "prigusena stavka ostaje prigusena i posle prebojavanja sidebara",
    ),
    "maticni-alatka-po-rednom-broju": (
        "modScrMatSistem.bas",
        "        outA(n, MS_COL_TAG) = CStr(src(i)(2))\n",
        "        outA(n, MS_COL_TAG) = CStr(src(0)(2))   ' SABOTAZA: uvek prva alatka\n",
        "T_MatSistem_UgovorIIdentitet",
        "red 1 posle pretrage nosi Tag SVOJE alatke",
    ),
    "maticni-ljuska-ne-pita-branu-ekrana": (
        "modUiScreens.bas",
        "    If ScrDozvoljen Then ScrDozvoljen = ScrSopstvenaBrana(kljuc)\n",
        "    ' SABOTAZA: ljuska ne pita ekran za njegovu branu\n",
        "T_MatSistem_UgovorIIdentitet",
        "ljuska postuje branu ekrana",
    ),
    # ------------------------------- MATICNI SIFARNICI: OPIS I CITANJE
    # Opis 13 sekcija je ono sto je preneto iz frmStammdaten. Cetiri sabotaze
    # ciljaju cetiri tvrdnje koje bi inace pukle tiho: identitet u koloni 1,
    # podnozje bez laznih zbirova, sema kao izvor istine za kolonu statusa, i
    # cip koji stvarno deli skup.
    "maticni-kolona-1-nije-pk": (
        "modMaticniIzvor.bas",
        "                \"OTKUI_HDM_ID|KooperantID|txt|84|1\", _\n",
        "                \"OTKUI_HDM_ID|Telefon|txt|84|1\", _\n",
        "T_MatIzvor_OpisSekcijaJePotpun",
        "kolona 1 svake sekcije NOSI PK",
    ),
    "maticni-tezina-kao-kilogrami": (
        "modMaticniIzvor.bas",
        "                \"OTKUI_HDM_TEZINA_GAJ|\" & COL_TAMB_TEZINA & \"|num|130|1\")\n",
        "                \"OTKUI_HDM_TEZINA_GAJ|\" & COL_TAMB_TEZINA & \"|kg|130|1\")\n",
        "T_MatIzvor_OpisSekcijaJePotpun",
        "nijedna maticna kolona nije kilogramska ni novcana",
    ),
    "maticni-status-pogadja-umesto-da-trazi": (
        "modMaticniIzvor.bas",
        "    If GetColumnIndex(tbl, \"Aktivan\") > 0 Then\n        MatStatusKolona = \"Aktivan\"\n    ElseIf GetColumnIndex(tbl, \"Aktivna\") > 0 Then\n        MatStatusKolona = \"Aktivna\"\n    End If\n",
        "    MatStatusKolona = \"Aktivan\"   ' SABOTAZA: pogadja umesto da trazi u semi\n",
        "T_MatIzvor_OpisSekcijaJePotpun",
        "kolona statusa se trazi u semi, ne pogadja",
    ),
    "maticni-cip-ne-deli-skup": (
        "modMaticniIzvor.bas",
        "            If filter = MAT_CIP_NEAKT And aktivan Then GoTo Sledeci\n",
        "            ' SABOTAZA: cip 'neaktivni' ne filtrira nista\n",
        "T_MatIzvor_CipIdentitetIPretraga",
        "aktivni + neaktivni = svi",
    ),
    # ------------------------------------- MATICNI: JEDAN PISAC (M2)
    # Provere su iz forme presle u modul; ako modul prestane da odbija, forma
    # to vise nema gde da uhvati -- ona od v6-ui-189 samo prikazuje odgovor.
    "maticni-unos-ne-trazi-obavezno": (
        "modMaticniUnos.bas",
        "        If modMaticniIzvor.PoljeF(spec, 3) = \"1\" And Len(v) = 0 Then\n",
        "        If False Then   ' SABOTAZA: obavezno polje se ne trazi\n",
        "T_MatUnos_ProveraOdbija",
        "kooperant bez imena se odbija",
    ),
    "maticni-unos-nula-hektara-prolazi": (
        "modMaticniUnos.bas",
        "    If kljuc = \"PARCELE\" And poljeKljuc = \"povrsina\" Then TraziPozitivan = True\n",
        "    ' SABOTAZA: nula hektara postaje dozvoljena\n",
        "T_MatUnos_ProveraOdbija",
        "parcela od nula hektara se odbija",
    ),
    "maticni-unos-prag-blok-bez-provere": (
        "modMaticniUnos.bas",
        "                If upoz > 0 And blok > 0 And blok < upoz Then\n",
        "                If False Then   ' SABOTAZA: pragovi se ne uporedjuju\n",
        "T_MatUnos_ProveraOdbija",
        "prag blokade ispod praga upozorenja se odbija",
    ),
    "maticni-alias-se-ne-razresava": (
        "modMaticniIzvor.bas",
        "    If Left$(kol, 7) <> \"@alias:\" Then\n",
        "    If True Then   ' SABOTAZA: alias ostaje nerazresen\n",
        "T_MatUnos_OpisPoljaISema",
        "svako polje pise u kolonu koja POSTOJI u semi",
    ),
    # ------------------------------------------------- MREZA: PODNOZJE
    # Zbir se racuna uvek, ali ljuska odlucuje hoce li ga NACRTATI. Kad novcane
    # kolone nisu na spisku, podnozje se sakrije uz savrseno tacan zbir i zelenu
    # suite -- tacno to se desilo listi izvoda.
    "mreza-rest-nije-novcana-kolona": (
        "modOtkupUI.bas",
        "            Case \"rsd\", \"mult\", \"sum0\", \"rest\": OpisImaValKolonu = True: Exit Function\n",
        "            Case \"rsd\", \"mult\", \"sum0\": OpisImaValKolonu = True: Exit Function   ' SABOTAZA: rest nije novac\n",
        "T_BankaUvoz_IzvodiSuAgregatPoRacunu",
        "ljuska za listu izvoda crta zbir vrednosti u podnozju",
    ),
    # ---------------------------------------- MREZA: JEDINICA PODNOZJA
    # Jedinica i broj decimala u podnozju zavise od EKRANA, ne od globalnog
    # ActiveMode (rezim unosa dokumenata). Cetiri svojstva -> cetiri sabotaze:
    # da ljuska uopste pita, da ekran bez ugovora dobije DINARE (fail-closed),
    # da pitanje ne ode opet globalnom rezimu, i da dinari nose pare.
    "mreza-podnozje-ljuska-ne-pita-ekran": (
        "modUiScreens.bas",
        "    ScrBrojiKomade = CBool(Application.Run(m & \".Scr_BrojiKomade\"))\n",
        "    ScrBrojiKomade = False   ' SABOTAZA: ljuska ne pita ekran\n",
        "T_Mreza_PodnozjeJedinicaIdeIzUgovoraEkrana",
        "Dokumenta na reversima i dalje broje komade",
    ),
    "mreza-podnozje-ugovor-fail-open": (
        "modUiScreens.bas",
        "        ScrBrojiKomade = False\n",
        "        ScrBrojiKomade = True   ' SABOTAZA: ekran bez ugovora nasledjuje komade\n",
        "T_Mreza_PodnozjeJedinicaIdeIzUgovoraEkrana",
        "ugovorni ekran ne nasledjuje rezim unosa dokumenata",
    ),
    "mreza-podnozje-jedinica-iz-globalnog-rezima": (
        "modOtkupUI.bas",
        "    If modUiScreens.ScrBrojiKomade(mScreen) Then\n",
        "    If ModeBrojiKomade(ActiveMode) Then   ' SABOTAZA: jedinica iz tudjeg rezima\n",
        "T_Mreza_PodnozjeJedinicaIdeIzUgovoraEkrana",
        "podnozje ugovornog ekrana ne pominje komade",
    ),
    "mreza-podnozje-novac-bez-para": (
        "modOtkupUI.bas",
        "                           FmtBroj(iznos, 2) & \" \" & Poruka(\"OTKUI_UNIT_RSD\")\n",
        "                           FmtBroj(iznos, 0) & \" \" & Poruka(\"OTKUI_UNIT_RSD\")\n",
        "T_Mreza_PodnozjeJedinicaIdeIzUgovoraEkrana",
        "novac u podnozju ide sa parama",
    ),
    # Storno prikazuje osam tipova, medju njima i REVERSE, i to preko ISTOG
    # citaca kao ekran dokumenata -- pa mu u podnozje stize zbir komada. Dve
    # sabotaze, jer su i dve greske moguce: da ne prijavi komade (fail-closed ga
    # onda proglasi dinarima), i da ih prijavi UVEK (fakture postanu komadi).
    "mreza-podnozje-storno-ne-prijavljuje-komade": (
        "modScrStorno.bas",
        "    Scr_BrojiKomade = modScrDokumenti.TipBrojiKomade(Scr_Lista())\n",
        "    Scr_BrojiKomade = False   ' SABOTAZA: Storno cuti o komadima\n",
        "T_Mreza_PodnozjeJedinicaIdeIzUgovoraEkrana",
        "Storno lista Reversi broji komade",
    ),
    "mreza-podnozje-storno-uvek-komadi": (
        "modScrStorno.bas",
        "    Scr_BrojiKomade = modScrDokumenti.TipBrojiKomade(Scr_Lista())\n",
        "    Scr_BrojiKomade = True   ' SABOTAZA: Storno uvek broji komade\n",
        "T_Mreza_PodnozjeJedinicaIdeIzUgovoraEkrana",
        "ostali tipovi na Stornu ne broje komade",
    ),
    # ---------------------------------------- LEGACY: frmDokumenta, blok/avans
    # Prazna lista posle PADA ucitavanja opet postaje 'nema bloka', pa novac
    # tiho ode kao avans kooperanta. Ista klasa koju je frmBankaImport imao
    # tri puta (PR #220).
    # Sidro je dvoredno jer je naredba dvoredna: marker NE sme iza "_"
    # (zamka 4 -- komentar posle nastavka reda visi kompajler, ne test).
    "dok-pad-liste-blokova-prolazi": (
        "frmDokumenta.frm",
        "    BlokIzborSme = ListaSme(m_BlokoviOk, m_BlokoviErr, _\n",
        "    ' SABOTAZA: pad ucitavanja prolazi\n"
        "    BlokIzborSme = ListaSme(True, m_BlokoviErr, _\n",
        "T_LegacyDok_PadListeBlokovaNijeAvans",
        "pad ucitavanja liste blokova ZAUSTAVLJA knjizenje avansa",
    ),
    # Kapija sira od kvara je isto greska: prazna lista posle USPESNOG
    # citanja stvarno znaci 'nema otvorenih blokova', i avans je tada tacan.
    "dok-kapija-blokova-presiroka": (
        "frmDokumenta.frm",
        "                            \"otkupnih blokova\", \"da bloka nema\", outPoruka)\n",
        "                            \"otkupnih blokova\", \"da bloka nema\", outPoruka)\n"
        "    BlokIzborSme = False   ' SABOTAZA: kapija ne pusta ni urednu listu\n",
        "T_LegacyDok_PadListeBlokovaNijeAvans",
        "uredno ucitana lista pusta avans",
    ),
    # Kapija koja zavisi od izbora vazi samo za pola slucajeva: delimicno
    # napunjen kombo posle pada ucitavanja ima izbor, pa bi prosao.
    # ---- filter storniranih: nula je imala dva znacenja ----
    # Registar mora da PREPOZNA dokument tabelu; prazan registar vraca stari
    # fail-open na sva 183 poziva.
    "storno-registar-prazan": (
        "modSchemaGuard.bas",
        "    TabelaNosiStorno = _\n"
        '        (InStr(1, STORNO_TABELE, "|" & Trim$(tableName) & "|", vbTextCompare) > 0)\n',
        "    TabelaNosiStorno = False   ' SABOTAZA: registar je prazan\n",
        "T_StornoFilter_NedostajucaKolonaNijeTisina",
        "dokument tabela je u registru storna",
    ),
    # ...ali ne sme da hvata maticne podatke, koji storno pojam stvarno nemaju.
    "storno-registar-hvata-i-maticne": (
        "modSchemaGuard.bas",
        "    TabelaNosiStorno = _\n"
        '        (InStr(1, STORNO_TABELE, "|" & Trim$(tableName) & "|", vbTextCompare) > 0)\n',
        "    TabelaNosiStorno = True   ' SABOTAZA: registar hvata i maticne podatke\n",
        "T_StornoFilter_NedostajucaKolonaNijeTisina",
        "maticni podaci nisu u registru storna",
    ),
    # Sama kapija u ExcludeStornirano: kad SVE prodje kao "nema storno pojam",
    # nedostajuca kolona opet cuti.
    "storno-filter-nedostajuca-kolona-prolazi": (
        "modHelpers.bas",
        "    If Not modSchemaGuard.TabelaNosiStorno(tblName) Then\n",
        "    If True Then   ' SABOTAZA: sve prolazi kao tabela bez storna\n",
        "T_StornoFilter_NedostajucaKolonaNijeTisina",
        "nedostajuca kolona storna PADA i imenuje kolonu, ne propusta tiho",
    ),
    # Uzina kapije je tvrdnja: tabela bez storno pojma mora da PRODJE.
    "storno-filter-hvata-i-tabele-bez-storna": (
        "modHelpers.bas",
        "    If Not modSchemaGuard.TabelaNosiStorno(tblName) Then\n",
        "    If False Then   ' SABOTAZA: kapija hvata i tabele bez storna\n",
        "T_StornoFilter_NedostajucaKolonaNijeTisina",
        "tabela bez storno pojma prolazi bez greske",
    ),
    # Prazna tabela iz registra bez kolone je I DALJE drift. Dok je IsEmpty
    # izlazio prvi, kapija se nad njom nikad nije ni pitala.
    "storno-filter-prazna-tabela-preskace-kapiju": (
        "modHelpers.bas",
        '    RequireStornoKlasifikaciju tblName, "modHelpers.ExcludeStornirano"\n',
        "    If IsEmpty(data) Then Exit Function   ' SABOTAZA: prazno preskace kapiju\n"
        '    RequireStornoKlasifikaciju tblName, "modHelpers.ExcludeStornirano"\n',
        "T_StornoFilter_NedostajucaKolonaNijeTisina",
        "...i kad je tabela PRAZNA -- drift ne ceka da bude redova",
    ),
    # Neklasifikovana tabela nije isto sto i BEZ_STORNA. Bez ove kapije se dva
    # stanja opet spajaju -- ista bolest zbog koje je posao i nastao.
    "storno-nepoznata-tabela-prolazi": (
        "modSchemaGuard.bas",
        "    If StornoRegistarZna(tableName) Then Exit Sub\n",
        "    Exit Sub   ' SABOTAZA: neklasifikovana tabela tiho prolazi\n",
        "T_StornoFilter_NedostajucaKolonaNijeTisina",
        "tabela koju registar ne poznaje PADA, ne prolazi kao da nema storno",
    ),
    # ---- ista klasa na strani KUPCA (legacy F6) i u LJUSCI (F5/F6) ----
    # Pad ucitavanja liste faktura mora da zaustavi uplatu -- inace tiho postaje
    # avans kupca, isto kao sto je pad liste blokova postajao avans kooperanta.
    "dok-pad-liste-faktura-prolazi": (
        "frmDokumenta.frm",
        "    FakturaIzborSme = ListaSme(m_FaktureOk, m_FaktureErr, _\n",
        "    ' SABOTAZA: pad ucitavanja liste faktura prolazi\n"
        "    FakturaIzborSme = ListaSme(True, m_FaktureErr, _\n",
        "T_LegacyDok_PadListeFakturaNijeAvans",
        "pad ucitavanja liste faktura ZAUSTAVLJA knjizenje avansa kupca",
    ),
    # Kapija koja zavisi od izbora vazi samo za pola slucajeva: delimicno
    # napunjen kombo posle pada ima izbor, pa bi prosao.
    "dok-faktura-izbor-zaobilazi": (
        "frmDokumenta.frm",
        "    UplataSme = FakturaIzborSme(outPoruka)\n",
        "    If fakturaIzabrana Then UplataSme = True: Exit Function   ' SABOTAZA\n",
        "T_LegacyDok_PadListeFakturaNijeAvans",
        "ni IZABRANA faktura ne prolazi kad je ucitavanje palo",
    ),
    # Uzina kapije je TVRDNJA, ne izuzetak: bez novca nema odluke faktura/avans,
    # pa unos same ambalaze ne sme da stane zbog liste koja ga se ne tice.
    "dok-uplata-kapija-siri-se-na-ambalazu": (
        "frmDokumenta.frm",
        "    If novac <= 0 Then\n        UplataSme = True\n        Exit Function\n    End If\n",
        "    If False Then   ' SABOTAZA: kapija hvata i unos bez novca\n"
        "        UplataSme = True\n        Exit Function\n    End If\n",
        "T_LegacyDok_PadListeFakturaNijeAvans",
        "unos bez novca NE staje zbog liste faktura",
    ),
    # Kapija sira od kvara je isto greska: uredno ucitana prazna lista stvarno
    # znaci "kupac nema otvorenih faktura", i avans je tada tacan.
    "dok-kapija-faktura-presiroka": (
        "frmDokumenta.frm",
        "                               \"otvorenih faktura\", \"da fakture nema\", outPoruka)\n",
        "                               \"otvorenih faktura\", \"da fakture nema\", outPoruka)\n"
        "    FakturaIzborSme = False   ' SABOTAZA: kapija ne pusta ni urednu listu\n",
        "T_LegacyDok_PadListeFakturaNijeAvans",
        "uredno ucitana lista pusta uplatu",
    ),
    # LJUSKA: ista greska, drugi domacin. Pad je ovde isao samo u Debug.Print.
    "ljuska-pad-liste-blokova-prolazi": (
        "modOtkupUI.bas",
        "            NovacListaSme = LjuskaListaSme(mBlokListaOk, mBlokListaErr, _\n",
        "            ' SABOTAZA: pad liste blokova prolazi\n"
        "            NovacListaSme = LjuskaListaSme(True, mBlokListaErr, _\n",
        "T_Ljuska_PadListeNovcaNijeAvans",
        "pad liste blokova ZAUSTAVLJA isplatu kooperantu u ljusci",
    ),
    "ljuska-pad-liste-faktura-prolazi": (
        "modOtkupUI.bas",
        "            NovacListaSme = LjuskaListaSme(mFakListaOk, mFakListaErr, _\n",
        "            ' SABOTAZA: pad liste faktura prolazi\n"
        "            NovacListaSme = LjuskaListaSme(True, mFakListaErr, _\n",
        "T_Ljuska_PadListeNovcaNijeAvans",
        "pad liste faktura ZAUSTAVLJA uplatu kupca u ljusci",
    ),
    # Isplata otkupnom mestu ne dodiruje blokove; kapija koja i nju hvata bi
    # zaustavila rad bez ijednog pogresnog knjizenja.
    "ljuska-kapija-hvata-i-otkupno-mesto": (
        "modOtkupUI.bas",
        '            If UCase$(Trim$(CStr(p("partnerTip")))) <> "KOOP" Then Exit Function\n',
        "            ' SABOTAZA: kapija hvata i isplatu otkupnom mestu\n",
        "T_Ljuska_PadListeNovcaNijeAvans",
        "isplata otkupnom mestu ne zavisi od liste blokova",
    ),
    # Bez novca nema odluke blok/faktura, pa unos same ambalaze ne sme da stane
    # zbog liste koja ga se ne tice. Legacy kopija ovog pravila ima svoju
    # sabotazu (dok-uplata-kapija-siri-se-na-ambalazu); posto PR namerno drzi
    # dve odvojene kopije po domacinu, legacy sabotaza NE dokazuje ovu.
    "ljuska-kapija-hvata-i-bez-novca": (
        "modOtkupUI.bas",
        '    If CDbl(p("novac")) <= 0 Then Exit Function\n',
        "    ' SABOTAZA: kapija hvata i unos bez novca\n",
        "T_Ljuska_PadListeNovcaNijeAvans",
        "unos bez novca ne staje zbog liste",
    ),
    # Rezimi bez tih listi (F1-F4, F7) kapiju ne smeju da osete.
    "ljuska-kapija-hvata-sve-rezime": (
        "modOtkupUI.bas",
        '    Select Case CStr(p("rezim"))\n        Case "AMB_ISPLATE"\n',
        '    Select Case "AMB_ISPLATE"   \' SABOTAZA: kapija van gotovinskih rezima\n'
        '        Case "AMB_ISPLATE"\n',
        "T_Ljuska_PadListeNovcaNijeAvans",
        "rezim bez tih listi kapiju ne oseca",
    ),
    "dok-izbor-zaobilazi-kapiju": (
        "frmDokumenta.frm",
        "    KnjizenjeSme = BlokIzborSme(outPoruka)\n",
        "    If blokIzabran Then KnjizenjeSme = True: Exit Function   ' SABOTAZA\n",
        "T_LegacyDok_PadListeBlokovaNijeAvans",
        "ni IZABRAN blok ne prolazi kad je ucitavanje palo",
    ),
    # ---------------------------------------- MREZA: POZADINA PILULE
    # Natpis se brisao i pre; POZADINA je ostajala, pa je celija i dalje bila
    # obojen pravougaonik koji tvrdi stanje -- samo bez slova. Ovo je bila
    # rupa zapisana u katalogu 10.6 kao neizmerena.
    "mreza-pilula-pozadina-ostaje": (
        "modOtkupUI.bas",
        "    lbl.BackStyle = fmBackStyleTransparent\n",
        "    ' SABOTAZA: pozadina pilule ostaje\n",
        "T_MrezaPilula_PozadinaSeCisti",
        "pilula koja se ne moze prikazati gubi i POZADINU, ne samo natpis",
    ),
    # Ciscenje koje se ne moze ponistiti je druga polovina istog ugovora:
    # ---------------------------------------- SEMA: TRAZENJE KOLONE
    # Vraca zatecen oblik: ListColumns(ime) DIZE gresku 9 za nepostojecu
    # Zatecено ponasanje je poredjenje BEZ obzira na velicinu slova.
    # Poruka bez zaglavlja opisuje tri razlicita stanja istim tekstom:
    # kolone nema, tabele nema, zaglavlje je drugacije.
    "kolona-poruka-bez-zaglavlja": (
        "modSchemaGuard.bas",
        "    If n > MAX_IMENA Then s = s & \", ... (+\" & (n - MAX_IMENA) & \")\"\n"
        "    If Len(s) = 0 Then s = \"prazno\"\n",
        "    s = \"prazno\"   ' SABOTAZA: poruka bez zaglavlja\n",
        "T_Kolona_TrazenjeNeGutaGresku",
        "zaglavlje koje je stvarno videla",
    ),
    # Bez ovoga poruka daje spisak imena, a ne odgovor na jedino pitanje
    # koje se iz nje trazi: da li BAS ta kolona postoji u zaglavlju.
    "kolona-poruka-ne-kaze-da-je-vidjena": (
        "modSchemaGuard.bas",
        "    If poz > 0 Then\n"
        "        ZaglavljeZaPoruku = s & \". Trazena kolona VIDJENA, pozicija \" & poz\n",
        "    If False Then   ' SABOTAZA: poruka cuti o trazenoj koloni\n"
        "        ZaglavljeZaPoruku = s & \". Trazena kolona VIDJENA, pozicija \" & poz\n",
        "T_Kolona_TrazenjeNeGutaGresku",
        "poruka kaze da je trazena kolona VIDJENA u svezem prolazu",
    ),
    # Tabele nema i zaglavlje je prazno su dva razlicita stanja; jedan tekst
    # za oba vraca upravo onu neodredjenost zbog koje je poruka i prosirena.
    "kolona-nema-tabele-kao-prazno": (
        "modSchemaGuard.bas",
        "    If lo Is Nothing Then\n"
        "        ZaglavljeZaPoruku = \"tabela nije nadjena\"\n",
        "    If lo Is Nothing Then\n"
        "        ZaglavljeZaPoruku = \"prazno\"   ' SABOTAZA: nema tabele = prazno\n",
        "T_Kolona_TrazenjeNeGutaGresku",
        "za nepostojecu tabelu poruka kaze da TABELE nema",
    ),
    # ---------------------------------------- MREZA: DVA NOVCANA SLOTA
    # Ekran opet salje samo zbir vrednosti: sedmog clana nema, pa dva broja
    # koja operater trazi nemaju kuda da stignu.
    "mreza-podnozje-slot-nema-ugovora": (
        "modScrBankaUvoz.bas",
        "    RedoviIzvodi = Array(IzvodiKolone(), outA, n, 0#, zbirU + zbirI, Array(0, 0, 0), _\n"
        "                         Array(Array(\"OTKUI_FT_UPLATE\", zbirU), _\n"
        "                               Array(\"OTKUI_FT_ISPLATE\", zbirI)))\n",
        "    RedoviIzvodi = Array(IzvodiKolone(), outA, n, 0#, zbirU + zbirI, Array(0, 0, 0))   ' SABOTAZA: nema sedmog clana\n",
        "T_Mreza_PodnozjeDvaNovcanaSlota",
        "ugovor nosi sedmi clan",
    ),
    # Slot koji ne prati filtere je gori od jednog zbira koji ih prati: dva
    # broja izgledaju preciznije, a opisuju drugu listu.
    "mreza-podnozje-slot-mimo-filtera": (
        "modScrBankaUvoz.bas",
        "    RedoviIzvodi = Array(IzvodiKolone(), outA, n, 0#, zbirU + zbirI, Array(0, 0, 0), _\n"
        "                         Array(Array(\"OTKUI_FT_UPLATE\", zbirU), _\n",
        "    ' SABOTAZA: slot broji mimo liste\n"
        "    RedoviIzvodi = Array(IzvodiKolone(), outA, n, 0#, zbirU + zbirI, Array(0, 0, 0), _\n"
        "                         Array(Array(\"OTKUI_FT_UPLATE\", zbirU + zbirI), _\n",
        "T_Mreza_PodnozjeDvaNovcanaSlota",
        "zbir slotova je promet",
    ),
    # Ekran salje slotove, ljuska ih ne cita -- podnozje ostaje na jednom
    # zbiru, a niko nista ne prijavi.
    "mreza-podnozje-ljuska-ne-uzima-slotove": (
        "modOtkupUI.bas",
        "    If UBound(d) >= 6 Then\n"
        "        If IsArray(d(6)) Then\n",
        "    If False Then   ' SABOTAZA: ljuska ne cita slotove\n"
        "        If IsArray(d(6)) Then\n",
        "T_Mreza_PodnozjeDvaNovcanaSlota",
        "ljuska je preuzela oba slota",
    ),
    # Goli broj bez natpisa: dva iznosa jedan pored drugog, a ne pise koji je
    # koji -- operater bira napamet.
    "mreza-podnozje-slot-bez-natpisa": (
        "modOtkupUI.bas",
        "    PodnozjeSlotTekst = Poruka(kljuc) & \" \" & FmtBroj(iznos, 2) & \" \" & _\n"
        "                        Poruka(unitKljuc)\n",
        "    PodnozjeSlotTekst = FmtBroj(iznos, 2) & \" \" & _\n"
        "                        Poruka(\"OTKUI_UNIT_RSD\")   ' SABOTAZA: slot bez natpisa\n",
        "T_Mreza_PodnozjeDvaNovcanaSlota",
        "prvi slot nosi natpis Uplate",
    ),
    # Oba slota crtaju PRVI iznos: podnozje pokazuje dva broja, a to je jedan
    # isti -- najtise moguce, jer izgleda tacno onako kako treba.
    "mreza-podnozje-oba-slota-isti": (
        "modOtkupUI.bas",
        "        SlotTekstIz = PodnozjeSlotTekst(CStr(mFtSlot(i)(0)), CDbl(mFtSlot(i)(1)))\n"
        "    End If\n",
        "        SlotTekstIz = PodnozjeSlotTekst(CStr(mFtSlot(i)(0)), CDbl(mFtSlot(0)(1)))   ' SABOTAZA: oba slota nose PRVI iznos\n"
        "    End If\n",
        "T_Mreza_PodnozjeDvaNovcanaSlota",
        "slotovi ne nose isti IZNOS",
    ),
    # Ekran ima DVA citaca, i oba su dobila sedmi clan. Bez ove sabotaze bi
    # lista stavki bila samo tvrdnja u opisu PR-a, bez ijednog dokaza.
    "mreza-podnozje-stavke-nema-slotova": (
        "modScrBankaUvoz.bas",
        "    RedoviStavke = Array(StavkeKolone(), outA, n, 0#, zbirU + zbirI, Array(0, 0, 0), _\n"
        "                         Array(Array(\"OTKUI_FT_UPLATE\", zbirU), _\n"
        "                               Array(\"OTKUI_FT_ISPLATE\", zbirI)))\n",
        "    ' SABOTAZA: stavke opet salju samo zbir\n"
        "    RedoviStavke = Array(StavkeKolone(), outA, n, 0#, zbirU + zbirI, Array(0, 0, 0))\n",
        "T_Mreza_PodnozjeDvaNovcanaSlota",
        "i lista stavki nosi sedmi clan",
    ),
    # Slot koji je TACAN nad punom listom a pogresan nad suzenom -- najtezi
    # oblik, jer prva provera prolazi. Promet se namerno ne dira: da je
    # diran, pao bi test prometa iznad i ova sabotaza ne bi merila svoje.
    "mreza-podnozje-slot-ignorise-pretragu": (
        "modScrBankaUvoz.bas",
        "    RedoviIzvodi = Array(IzvodiKolone(), outA, n, 0#, zbirU + zbirI, Array(0, 0, 0), _\n"
        "                         Array(Array(\"OTKUI_FT_UPLATE\", zbirU), _\n",
        "    ' SABOTAZA: slot ne prati pretragu\n"
        "    RedoviIzvodi = Array(IzvodiKolone(), outA, n, 0#, zbirU + zbirI, Array(0, 0, 0), _\n"
        "                         Array(Array(\"OTKUI_FT_UPLATE\", zbirU + IIf(Len(q) > 0, 1000000, 0)), _\n",
        "T_Mreza_PodnozjeDvaNovcanaSlota",
        "pretraga smanjuje i slotove, ne samo redove",
    ),
    # ---------------------------------------- BANKA: PODNOZJE IZVODA
    # Status kaze "ne zna se koji zbirovi vaze", pa brojke ne smeju ni da se
    # prikazu ni da udju u promet. Prikazati vrednost PRVOG reda pored tog
    # natpisa znaci ponuditi tudji podatak kao saldo.
    "banka-izvod-nesaglasan-prikazuje-brojke": (
        "modScrBankaUvoz.bas",
        "        If nesagl Then\n"
        "            outA(n, 4) = 0#\n",
        "        If False Then   ' SABOTAZA: nesaglasan izvod prikazuje brojke\n"
        "            outA(n, 4) = 0#\n",
        "T_BankaUvoz_IzvodiSuAgregatPoRacunu",
        "nesaglasan izvod nema pocetno stanje",
    ),
    "banka-izvod-nesaglasan-ulazi-u-promet": (
        "modScrBankaUvoz.bas",
        "        If Not nesagl Then\n"
        "            zbirU = zbirU + CDbl(src(i, 6))\n",
        "        If True Then   ' SABOTAZA: nesaglasan ulazi u promet\n"
        "            zbirU = zbirU + CDbl(src(i, 6))\n",
        "T_BankaUvoz_IzvodiSuAgregatPoRacunu",
        "nesaglasan izvod ne donosi NISTA u promet -- ne zna se koji zbirovi vaze",
    ),
    # Zbir je stajao IZMEDJU cipa i pretrage, pa je izvod koji pretraga sakrije
    # i dalje ulazio u promet -- traka je tvrdila promet redova kojih nema.
    "banka-izvod-promet-ne-postuje-pretragu": (
        "modScrBankaUvoz.bas",
        "        hay = CStr(src(i, 2)) & \"|\" & CStr(src(i, 3))\n"
        "        If Len(q) > 0 Then\n"
        "            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeci\n"
        "        End If\n"
        "\n"
        "        ' ZBIR IDE POSLE OBA FILTERA. Ranije je stajao izmedju cipa i pretrage,\n"
        "        ' pa je izvod koji pretraga sakrije i dalje ulazio u podnozje -- traka je\n"
        "        ' tvrdila promet redova kojih na ekranu nema.\n"
        "        '\n"
        "        ' Nesaglasan izvod NE ULAZI uopste: njegovi zbirovi su vrednost PRVOG\n"
        "        ' reda, a upravo se ne zna koji red vazi. Sabrati ih znacilo bi tvrditi\n"
        "        ' promet koji nikad nije izracunat.\n"
        "        nesagl = (CLng(src(i, 10)) = BIM_SALDO_NEKONZISTENTAN)\n"
        "        If Not nesagl Then\n"
        "            zbirU = zbirU + CDbl(src(i, 6))\n"
        "            zbirI = zbirI + CDbl(src(i, 7))\n"
        "        End If\n",
        "        ' SABOTAZA: zbir se racuna PRE pretrage\n"
        "        hay = CStr(src(i, 2)) & \"|\" & CStr(src(i, 3))\n"
        "        nesagl = (CLng(src(i, 10)) = BIM_SALDO_NEKONZISTENTAN)\n"
        "        If Not nesagl Then\n"
        "            zbirU = zbirU + CDbl(src(i, 6))\n"
        "            zbirI = zbirI + CDbl(src(i, 7))\n"
        "        End If\n"
        "        If Len(q) > 0 Then\n"
        "            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeci\n"
        "        End If\n",
        "T_BankaUvoz_IzvodiSuAgregatPoRacunu",
        "pretraga smanjuje i promet, ne samo broj redova",
    ),
    # ---------------------------------------- BANKA: SALDO IZVODA
    # Parser upisuje saldo na SVAKI red grupe, pa ga agregat UZIMA sa prvog reda
    # umesto da sabira. To vazi dok su svi redovi saglasni -- a to niko nije
    # proveravao. Rucno editovan red bi prosao kao istina o celom izvodu.
    "banka-izvod-saldo-prvi-red-pobedjuje": (
        "modBankaImport.bas",
        "        If nesaglasan(r) Then outA(r, 10) = BIM_SALDO_NEKONZISTENTAN\n",
        "        If False Then outA(r, 10) = BIM_SALDO_NEKONZISTENTAN   ' SABOTAZA: prvi red pobedjuje\n",
        "T_BankaUvoz_IzvodiSuAgregatPoRacunu",
        "izvod cija se dva reda razlikuju je nesaglasan",
    ),
    # Prag mora da bude isti kao kod slaganja (0.01). Sire poredjenje bi
    # zaokruzenja proglasavalo nesaglasnoscu i brojka bi postala neupotrebljiva.
    "banka-izvod-saldo-prag-preuzak": (
        "modBankaImport.bas",
        "    If Abs(potrazujeA - potrazujeB) > 0.01 Then Exit Function\n",
        "    If Abs(potrazujeA - potrazujeB) > 0.001 Then Exit Function\n",
        "T_BankaUvoz_IzvodiSuAgregatPoRacunu",
        "polovina centa nije -- prag je isti kao kod slaganja",
    ),
    # Nesaglasnost mora da NADJACA i "slaze se": prvi red ovog izvoda sam za sebe
    # daje slaganje, pa bi bez toga stajalo "slaze se" -- tvrdnja o brojkama
    # kojih nema.
    "banka-izvod-nesaglasno-je-razlika": (
        "modScrBankaUvoz.bas",
        "        Case BIM_SALDO_NEKONZISTENTAN\n"
        "            BuSlaganjeTekst = Poruka(\"OTKUI_LBL_BU_SALDO_NESAGLASAN\")\n",
        "        Case BIM_SALDO_NEKONZISTENTAN\n"
        "            BuSlaganjeTekst = Poruka(\"OTKUI_LBL_BU_SALDO_RAZLIKA\")\n",
        "T_BankaUvoz_IzvodiSuAgregatPoRacunu",
        "nesaglasan izvod dobija SVOJU poruku",
    ),
    # ---------------------------------------------------- BANKA: WRITER
    # Prazan skup kandidata writer knjizi kao avans kooperanta i stavku oznaci
    # obradjenom. Za AUTOMATSKO mapiranje je to namerno; za IZABRAN blok je
    # protivrecnost -- operater je rekao KOJI dug placa.
    #
    # PROVERA IDE NAD DRUGOM SUITOM: writer pise, pa tvrdnja zivi u
    # RunBankaImportTestSuite (transakciona, rollback). Dokaz:
    #   python tools/sabotaza.py banka-writer-placen-blok-je-avans
    #   python tools/run_vba.py --suite RunBankaImportTestSuite   # ocekuj FAIL
    "banka-writer-placen-blok-je-avans": (
        "modBankaMapiranje.bas",
        "    If IsEmpty(kandidati) And blokIzabran Then\n",
        "    If IsEmpty(kandidati) And False Then   ' SABOTAZA: izabran placen blok postaje avans\n",
        "T21_IzabranPlacenBlokNijeAvans",
        "izabran placen blok NE knjizi nista",
    ),
    # ---------------------------------------------------------------- BANKA UVOZ
    # Najtisi moguci kvar scope-a: zadat je, ali kolona nije dokaziva, pa filtar
    # otpadne i pozivalac dobije kandidate sa SVIH otkupnih mesta -- u listi koja
    # izgleda savrseno ispravno. Ime kolone je zato argument BimScopeKolona, da
    # bi se ta grana mogla izmeriti bez razbijanja seme fixture-a.
    "banka-uvoz-scope-bez-kolone-stanice": (
        "modBankaMapiranje.bas",
        '        BimScopeKolona = RequireColumnIndex(TBL_OTKUP, kolona, "BimScopeKolona")\n',
        "        BimScopeKolona = GetColumnIndex(TBL_OTKUP, kolona)   ' SABOTAZA: scope tiho otpada\n",
        "T_BankaUvoz_RucnoMapiranjePravila",
        "zadat scope nad nedokazivom kolonom PUCA -- ne vraca nescope-ovane kandidate",
    ),
    # Prazan scope izgleda isto kao "scope nije ni trazen", a znaci nesto sasvim
    # drugo: operater JESTE birao blok, samo taj red nema upisano otkupno mesto.
    # Propusten prazan scope raspodeli novac preko svih mesta sa istim brojem.
    # Lista blokova nudi SVAKI nestorniran broj otkupa i ne proverava dug, a
    # kandidati se biraju samo ako je "otvoreno > 0.009". Placen blok zato stoji
    # u listi a daje NULA kandidata -- i writer to ne prijavljuje kao gresku nego
    # knjizi AVANS i stavku oznaci obradjenom. Rucni izbor takvog bloka mora da
    # stane: operater je rekao KOJI dug placa.
    "banka-uvoz-placen-blok-postaje-avans": (
        "modBankaMapiranje.bas",
        "    BimBlokBezOtvorenih = IsEmpty(kandidati)\n",
        "    BimBlokBezOtvorenih = Not IsEmpty(kandidati)   ' SABOTAZA: placen blok prolazi\n",
        "T_BankaUvoz_RucnoMapiranjePravila",
        "potpuno placen blok NEMA otvorenih stavki",
    ),
    # Kapija sme da vazi SAMO za rucni izbor. Kad blok dolazi iz poziva na broj,
    # avans je namerno i dokumentovano ponasanje -- bezbedan izlaz dok je poreklo
    # dvosmisleno. Sabotaza uklanja tu razliku i gasi legitimnu granu.
    "banka-uvoz-kapija-bloka-i-za-poziv": (
        "modScrBankaUvoz.bas",
        "    If Len(Trim$(izabranBlok)) = 0 Then Exit Function\n",
        "    If Len(Trim$(efektivniBlok)) < 0 Then Exit Function   ' SABOTAZA: kapija i za poziv na broj\n",
        "T_BankaUvoz_RucnoMapiranjePravila",
        "isti blok iz POZIVA NA BROJ ne prolazi kroz kapiju -- avans ostaje namerno ponasanje",
    ),
    "banka-uvoz-blok-bez-om-prolazi": (
        "modScrBankaUvoz.bas",
        "    BuScopeNedostaje = (Len(Trim$(ciljID)) > 0 And Len(Trim$(stanica)) = 0)\n",
        "    BuScopeNedostaje = (Len(Trim$(ciljID)) > 0 And Len(Trim$(stanica)) < 0)\n",
        "T_BankaUvoz_RucnoMapiranjePravila",
        "izabran blok bez otkupnog mesta zaustavlja rucno mapiranje",
    ),
    # Prvi pad citanja u sesiji. Nula bi kroz BrojacTekst dala PRAZNU znacku, a
    # prazna znacka u ovom UI-ju znaci "nema sta da ceka" -- fail-open, samo tisi.
    "banka-uvoz-znacka-nepoznato-je-nula": (
        "modScrBankaUvoz.bas",
        "    BuKpiNepoznato = Array(-1, -1, -1, 0#, 0#)\n",
        "    BuKpiNepoznato = Array(0, 0, 0, 0#, 0#)   ' SABOTAZA: ne znam postaje nula\n",
        "T_BankaUvoz_CipJakihPratiBrojac",
        "bez ijedne poznate brojke stanje je NEPOZNATO",
    ),
    # Broj otkupa je jedinstven PO STANICI, pa isti broj bloka pripada dvama
    # razlicitim blokovima. Bez scope-a u jednu raspodelu ulaze kandidati sa OBA
    # otkupna mesta -- novac na dva razlicita poslovna lanca.
    "banka-uvoz-blok-bez-om-scope": (
        "modBankaMapiranje.bas",
        "        If Len(Trim$(stanicaID)) > 0 Then\n",
        "        If Trim$(stanicaID) = Chr$(0) Then   ' SABOTAZA: scope se ne primenjuje\n",
        "T_BankaUvoz_RucnoMapiranjePravila",
        "sa scope-om ulazi samo jedno otkupno mesto",
    ),
    # Znacka odgovara na pitanje "ima li posla". Kvar citanja pretvoren u nulu
    # kaze "nema posla" -- fail-open koji je Storno vec jednom platio.
    "banka-uvoz-kpi-greska-je-nula": (
        "modScrBankaUvoz.bas",
        "    If IsArray(poslednja) Then\n",
        "    If False Then   ' SABOTAZA: kvar citanja postaje nula\n",
        "T_BankaUvoz_CipJakihPratiBrojac",
        "posle greske se zadrzava POSLEDNJA POZNATA brojka",
    ),
    # Broj van opsega datuma obara CDate u mrezi, a RenderGrid to proguta
    # (On Error Resume Next) -- celija ostane sa natpisom od ranijeg crtanja.
    # Pravilo je LJUSKINO (modUiData), pa ga i sabotaza gadja tamo. Fixture
    # nosi red sa DatumTransakcije = 26062026, ddmmyyyy kao broj, posejan
    # bez datumskog formata -- tacno kako ga zatecene sveske nose.
    "mreza-datum-van-opsega": (
        "modUiData.bas",
        "    DatumSerijskiValidan = (serijski >= 1) And (serijski <= DATUM_SERIJSKI_MAX)\n",
        "    DatumSerijskiValidan = (serijski >= 1)\n",
        "T_MrezaDatum_BrojKojiNijeDatum",
        "ddmmyyyy kao broj NIJE datum",
    ),
    # Geometrija kolona mora da prati OPIS kolona. Bez zastavice RenderGrid
    # crta sa sirinama prethodne liste, pa kolona koja je tamo bila skrivena
    # ostaje nevidljiva i kad joj je vrednost tacna.
    "mreza-geometrija-ne-prati-kolone": (
        "modOtkupUI.bas",
        "        mGeomStara = True\n",
        "        mGeomStara = mGeomStara   \' SABOTAZA: zastavica se ne dize\n",
        "T_MrezaGeometrija_PratiOpisKolona",
        "promena opisa kolona proglasava geometriju zastarelom",
    ),
    # Za OM se ne bira ni faktura ni blok, pa se polje cilja GASI: polje koje
    # ne radi nista poziva operatera da u njega nesto upise. Ovo je uz to
    # jedina sabotaza koja PROLAZI kroz gradnju i raspored zone -- put na kom
    # se compile greske u RasporediPolja uopste i vide.
    "banka-uvoz-om-polje-cilja-radi": (
        "modScrBankaUvoz.bas",
        "    z.Controls(\"scrBuCilj\").Visible = (IzabraniTip() <> BIM_TIP_OM)\n",
        "    z.Controls(\"scrBuCilj\").Visible = True\n",
        "T_ZonaBankaUvoz_PoljaIRaspored",
        "za OM je polje cilja UGASENO",
    ),
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
        "lista STAVKE, red 1, kolona 2: datum je BROJ -- inace ga mreza ne crta",
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
        "stavke nose TACNO MAX_ACT radnji -- sesta bi se tiho odsekla",
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
        "prvi cip liste STAVKE je najsiri ('sve')",
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
        "izvodi su pregled -- nijedna radnja nad redom",
    ),
    # Kolona identiteta je interna. Prioritet 3 je crta, pa operater u listi
    # stavki gleda internu sifru u dve kolone.
    "banka-uvoz-identitet-vidljiv": (
        "modScrBankaUvoz.bas",
        '        "OTKUI_HDB_BIMKEY||txt|1|4", _\n',
        '        "OTKUI_HDB_BIMKEY||txt|90|3", _\n',
        "T_BankaUvoz_IdentitetURedu_NeCrtaSe",
        "identitet stavke je prioriteta 4 -- ne crta se",
    ),
    # Dvosmislen ID je ID koji u tabeli postoji dvaput. Ako prvi pobedi, radnja
    # se izvrsi nad redom koji operater NIJE pokazao -- tiho.
    "banka-uvoz-dvosmislen-prvi-pobedjuje": (
        "modBankaMapiranje.bas",
        "        outA(n, 1) = modFaktura.IdIliPrazno(brojac, Trim$(CStr(data(i, cID))))\n",
        "        outA(n, 1) = Trim$(CStr(data(i, cID)))   ' SABOTAZA: duplikat prolazi\n",
        "T_BankaUvoz_IdentitetURedu_NeCrtaSe",
        "dvosmislen ID nosi PRAZAN identitet -- radnja odbija da bira",
    ),
    # Otvorenost se cita iz onoga sto RED NOSI. Nov red ima PRAZAN status, pa se
    # iz prikaza ne razlikuje od reda kome status nije upisan.
    "banka-uvoz-red-ne-nosi-otvorenost": (
        "modScrBankaUvoz.bas",
        '        outA(n, 11) = IIf(CBool(src(i, 10)), "1", "")\n',
        '        outA(n, 11) = "1"   \' SABOTAZA: svaki red izgleda otvoren\n',
        "T_BankaUvoz_RedNosiSmerIOtvorenost",
        "obradjen red NE nosi otvorenost -- radnja ga odbija",
    ),
    # Smer se ne izvodi iz toga koja je kolona iznosa popunjena: red sa I
    # uplatom I isplatom izgleda kao uplata, a writer ga odbija.
    "banka-uvoz-red-ne-nosi-smer": (
        "modScrBankaUvoz.bas",
        "        outA(n, 12) = CStr(src(i, 11))\n",
        '        outA(n, 12) = ""   \' SABOTAZA: smer se gubi iz reda\n',
        "T_BankaUvoz_RedNosiSmerIOtvorenost",
        "red NOSI svoj smer",
    ),
    # Zatvorena stavka nema sta da predlozi -- predlog nad njom navodi operatera
    # da pokusa radnju koja ce biti odbijena.
    "banka-uvoz-predlog-i-za-zatvorene": (
        "modScrBankaUvoz.bas",
        "    If Not otvoren Then Exit Function\n",
        "    ' SABOTAZA: predlog se racuna i za zatvorene stavke\n",
        "T_BankaUvoz_RedNosiSmerIOtvorenost",
        "zatvorena stavka nema predlog",
    ),
    # Cip 'jaki kljucevi' i CountStrongKeyReadyBankaImport (koji stoji u natpisu
    # dugmeta) moraju da vide ISTI skup. Pravilo zivi na dva mesta.
    "banka-uvoz-cip-jaki-prolazi-sve": (
        "modScrBankaUvoz.bas",
        '        Case "jaki":       BuCipStavka = modBankaMapiranje.BimOtvoren(s) And jaki\n',
        '        Case "jaki":       BuCipStavka = modBankaMapiranje.BimOtvoren(s)\n',
        "T_BankaUvoz_CipJakihPratiBrojac",
        "cip 'jaki kljucevi' i BROJAC vide ISTI skup",
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
        "'sve' je tacno unija tri stanja -- nijedan red ne ispada iz svih cipova",
    ),
    # BROJ IZVODA NIJE IDENTITET: dedupe kljuc pocinje od BROJA RACUNA, pa dva
    # racuna firme legitimno nose izvod istog broja. Grupa bez racuna ih spaja u
    # jedan red i saldo dva razlicita racuna izgleda kao jedan.
    # Kljuc grupe izvoda ima DVE polovine (racun i ciklus), pa i dve sabotaze.
    # Test ih meri odvojenim tvrdnjama nad samim BimIzvodKljuc -- da bi svaka
    # sabotaza oborila BAS SVOJU (zamka 5); preko broja redova bi obe obarale
    # istu tvrdnju "isti broj daje tri reda".
    "banka-uvoz-izvod-kljuc-bez-racuna": (
        "modBankaImport.bas",
        '    BimIzvodKljuc = Trim$(brojDokumenta) & "|" & Trim$(brojRacuna) & _\n'
        '                    "|" & IzvodDatumKljuc(datumIzvoda)\n',
        '    BimIzvodKljuc = Trim$(brojDokumenta) & "|" & IzvodDatumKljuc(datumIzvoda)\n',
        "T_BankaUvoz_IzvodiSuAgregatPoRacunu",
        "isti broj izvoda na DVA RACUNA daje dva kljuca",
    ),
    # Banke numeraciju izvoda ponavljaju po ciklusu: izvod 15 na istom racunu
    # postoji i 2025. i 2026. Bez datuma u kljucu se spajaju u jedan sinteticki
    # red -- saldo i datum sa prvog, stavke sabrane preko oba.
    # Zamena je NAMERNO obrnutog redosleda: da nije, bila bi podniz sidra i
    # `--vrati` bi je nasao i u zdravom kodu (zamka 8).
    "banka-uvoz-izvod-kljuc-bez-datuma": (
        "modBankaImport.bas",
        '    BimIzvodKljuc = Trim$(brojDokumenta) & "|" & Trim$(brojRacuna) & _\n'
        '                    "|" & IzvodDatumKljuc(datumIzvoda)\n',
        '    BimIzvodKljuc = Trim$(brojRacuna) & "|" & Trim$(brojDokumenta)\n',
        "T_BankaUvoz_IzvodiSuAgregatPoRacunu",
        "isti broj i isti racun iz DVA CIKLUSA daju dva kljuca",
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
        "saldo se ne uzima sa tudjeg reda",
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
        "legacy red bez saldo metapodataka NIJE neslaganje nego odsustvo podatka",
    ),
    # Smer-kapija je ista koju RequireBimSmer sprovodi u writeru. OM prima i
    # uplatu i isplatu, ali NE i nejasan smer -- red sa oba iznosa writer odbija.
    "banka-uvoz-om-prima-nejasan-smer": (
        "modBankaMapiranje.bas",
        "            BimSmerOdgovaraTipu = (smer <> BIM_SMER_NEJASAN)\n",
        "            BimSmerOdgovaraTipu = True   ' SABOTAZA: OM prima sve\n",
        "T_BankaUvoz_RucnoMapiranjePravila",
        "nejasan smer ne prolazi ni za OM",
    ),
    # Prazan izbor bloka NIJE "nema bloka" nego "uzmi poziv na broj iz izvoda".
    # U formi je prazan combo bio DEFAULT slucaj, pa je blok sa 3+ stavki bez
    # ovog pravila zavrsavao generickom greskom umesto ponudjenom podelom.
    "banka-uvoz-prazan-blok-ostaje-prazan": (
        "modBankaMapiranje.bas",
        "        BimEfektivniBlok = AutoBlockNoForBim(bankaImportID)\n",
        '        BimEfektivniBlok = ""   \' SABOTAZA: poziv na broj se ne koristi\n',
        "T_BankaUvoz_RucnoMapiranjePravila",
        "prazan izbor uzima poziv na broj iz izvoda",
    ),
    # FAIL-CLOSED. Prazna lista faktura i PAD ucitavanja izgledaju isto, a znace
    # suprotno: prazan izbor fakture knjizi AVANS umesto zatvaranja duga.
    # Kapija je ZAJEDNICKA za kupca i kooperanta: prazna lista na obe rute nosi
    # poslovno znacenje (avans, odnosno poziv na broj), pa pad punjenja ne sme da
    # se pretvori ni u jedno od to dvoje.
    "banka-uvoz-cilj-fail-open": (
        "modScrBankaUvoz.bas",
        "    BuSmeMapiranjeCilja = mCiljOK\n",
        "    BuSmeMapiranjeCilja = True   ' SABOTAZA: pad citanja prolazi\n",
        "T_BankaUvoz_RucnoMapiranjePravila",
        "pad ucitavanja ZAUSTAVLJA rucno mapiranje -- prazan izbor bi bio avans ili poziv na broj",
    ),
    # Zajednicka kapija mora STVARNO da puni listu pre nego sto presudi.
    # Bez punjenja bi zastavica opisivala PRETHODNI izbor -- "ucitano" za tudji
    # combo, a odluka se donosi nad ovim.
    # Fakture gresku vracaju kroz ZASTAVICU, ne dizu je -- pa punjenje mirno
    # stigne do kesiranja. Zapamcen neuspeh drzi radnju blokiranom (fail-closed
    # radi), ali sledeci klik ne pokusava ponovo: izbor ostaje zakljucan.
    "banka-uvoz-kes-pamti-neuspeh": (
        "modScrBankaUvoz.bas",
        "    If ok Then CiljKesKljuc = kljuc\n",
        "    CiljKesKljuc = kljuc   ' SABOTAZA: neuspeh se pamti\n",
        "T_BankaUvoz_RucnoMapiranjePravila",
        "neuspelo punjenje se NE pamti -- sledeci klik pokusava ponovo",
    ),
    "banka-uvoz-cilj-kapija-ne-puni": (
        "modScrBankaUvoz.bas",
        "Private Function CiljUcitan(ByRef outPoruka As String) As Boolean\n"
        "    PuniCiljCombo\n",
        "Private Function CiljUcitan(ByRef outPoruka As String) As Boolean\n"
        "    outPoruka = outPoruka   ' SABOTAZA: lista se vise ne puni\n",
        "T_BankaUvoz_RucnoMapiranjePravila",
        "kapija puni listu cilja pre nego sto presudi",
    ),
    # Prazna tabela i NEPOSTOJECA tabela nisu isti ishod. GetTableData vraca
    # Empty za oba, pa bi citac koji gleda samo IsEmpty tumacio kvar kao "nema
    # redova" -- a prazan izbor fakture je AVANS.
    "banka-uvoz-nema-tabele-je-prazna": (
        "modSchemaGuard.bas",
        "    If GetTable(tableName) Is Nothing Then\n",
        "    If Len(tableName) < 0 Then   ' SABOTAZA: nema tabele prolazi\n",
        "T_BankaUvoz_RucnoMapiranjePravila",
        "nedostajuca tabela PUCA -- ne prolazi kao prazna lista",
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
        "placena faktura nije u listi za mapiranje",
    ),
    # ----- ekran Platni nalozi (v6-ui-185) -----------------------------------
    # Prvi cip je onaj na koji ljuska PADA kad zatecen filter ne pripada listi
    # (RefreshChipsForScreen). Ako nije najsiri, povratak na njega tiho sakrije
    # redove.
    "banka-nalozi-cip-sve-nije-prvi": (
        "modScrBankaNalozi.bas",
        '            BnCipoviZaListu = "sve:OTKUI_CHIP_SVE:40|" & _\n'
        '                              "imarac:OTKUI_CIPN_IMARAC:88|" & _\n',
        '            BnCipoviZaListu = "imarac:OTKUI_CIPN_IMARAC:88|" & _\n'
        '                              "sve:OTKUI_CHIP_SVE:40|" & _\n',
        "T_BankaNalozi_UgovorEkrana",
        "prvi cip liste NALOZI je najsiri ('sve')",
    ),
    # Kolona identiteta je interna. Prioritet 3 je crta, pa bi operater u
    # listi gledao interni OtkupID -- a radnja bi i dalje radila, sto prikaz
    # cini laznim dokazom da je "sve u redu".
    "banka-nalozi-identitet-vidljiv": (
        "modScrBankaNalozi.bas",
        '        "OTKUI_HDN_OTKID||txt|1|4", _\n',
        '        "OTKUI_HDN_OTKID||txt|90|3", _\n',
        "T_BankaNalozi_IdentitetURedu_NeCrtaSe",
        "identitet bloka je prioriteta 4 -- ne crta se",
    ),
    # Radnja avansa knjizi na (KooperantID, OtkupID). Red koji ne prenosi
    # vlasnika tera radnju da ga izvodi iz prikaza -- prikaz sme da se menja,
    # podatak ne.
    "banka-nalozi-red-ne-nosi-koopid": (
        "modScrBankaNalozi.bas",
        "        outA(n, 12) = CStr(src(i, 5))\n",
        '        outA(n, 12) = ""   \' SABOTAZA: red ne nosi vlasnika\n',
        "T_BankaNalozi_IdentitetURedu_NeCrtaSe",
        "red prenosi KooperantID -- radnja avansa ga ne izvodi iz prikaza",
    ),
    # Cip koji pusta sve pretvara "ima racun" u "sve": operater misli da
    # gleda blokove spremne za CSV, a gleda i one bez primaoca.
    "banka-nalozi-cip-imarac-pusta-sve": (
        "modScrBankaNalozi.bas",
        '        Case "imarac": BnCipNalog = imaTR\n',
        '        Case "imarac": BnCipNalog = True   \' SABOTAZA: cip pusta sve\n',
        "T_BankaNalozi_CipoviIKpiPratePravila",
        "cip 'ima racun' se slaze sa HasTekuciRacun iz citaca",
    ),
    # Avans je svojstvo KOOPERANTA: dva otvorena bloka istog coveka ne smeju
    # da mu udvostruce avans u KPI-ju. Legacy pravilo (koopAvansSet u
    # RefreshTopKpis) preneto u NalogeKpi.
    "banka-nalozi-kpi-avans-po-bloku": (
        "modBankaExportPregled.bas",
        "        If Not koopVidjen.Exists(blk.kooperantID) Then\n"
        "            koopVidjen.Add blk.kooperantID, True\n"
        "            avansPool = avansPool + blk.KooperantAvansSaldo\n"
        "        End If\n",
        "        koopVidjen(blk.kooperantID) = True\n"
        "        avansPool = avansPool + blk.KooperantAvansSaldo   ' SABOTAZA: po bloku\n",
        "T_BankaNalozi_CipoviIKpiPratePravila",
        "avans pool se sabira po KOOPERANTU, ne po bloku",
    ),
    # Pad citanja koji se prijavi kao nula znaci "nema posla" umesto "ne
    # znam" -- fail-open vec placen u Stornu i na Uvozu izvoda.
    "banka-nalozi-kpi-greska-je-nula": (
        "modScrBankaNalozi.bas",
        "    If IsArray(poslednja) Then\n"
        "        BnKpiPosleGreske = poslednja\n"
        "    Else\n"
        "        BnKpiPosleGreske = BnKpiNepoznato()\n"
        "    End If\n",
        "    BnKpiPosleGreske = Array(0, 0, 0#, 0#)   ' SABOTAZA: pad = nula\n",
        "T_BankaNalozi_CipoviIKpiPratePravila",
        "posle greske se zadrzava poslednja poznata brojka",
    ),
    # Blok bez tekuceg racuna nema primaoca: u CSV ne moze (writer ga
    # preskace), pa ne sme ni u izbor -- inace potvrda broji naloge koji
    # nikad ne nastanu.
    "banka-nalozi-bez-racuna-u-naloge": (
        "modScrBankaNalozi.bas",
        "    If Not imaTR Then\n"
        '        BnDodaj = Poruka("OTKUI_ERR_BN_BEZ_RACUNA")\n'
        "        Exit Function\n"
        "    End If\n",
        "    ' SABOTAZA: kapija racuna uklonjena\n",
        "T_BankaNalozi_KorpaIIzvoz",
        "blok bez tekuceg racuna ne ulazi u naloge -- nema primaoca",
    ),
    # Prazan identitet znaci da se ne zna KOJI blok -- dodavanje bi kasnije
    # radilo nad pogodjenim.
    "banka-nalozi-prazan-id-ulazi": (
        "modScrBankaNalozi.bas",
        "    If Len(Trim$(otkupID)) = 0 Then\n"
        '        BnDodaj = Poruka("OTKUI_ERR_BN_DVOSMISLEN")\n'
        "        Exit Function\n"
        "    End If\n",
        "    ' SABOTAZA: prazan identitet prolazi\n",
        "T_BankaNalozi_KorpaIIzvoz",
        "prazan identitet ne ulazi u naloge",
    ),
    # Izbor operatera mora da SUZI izvoz. Ignorisan izbor = nalozi za SVE
    # otvorene blokove, a operater je potvrdio brojku za svoj podskup.
    # Obara i T_BankaNalozi_IznosPoBloku (deljena osobina, ne sirina
    # sidra): taj test meri IZNOS kroz izbor od jednog bloka, pa izvoz koji
    # izbor ignorise vrati vise blokova i njegova tvrdnja o count-u padne.
    "banka-nalozi-izvoz-ignorise-izbor": (
        "modBankaExportPregled.bas",
        "    If Not samoOtkupIDs Is Nothing Then imaIzbor = (samoOtkupIDs.count > 0)\n",
        "    imaIzbor = False   ' SABOTAZA: izbor se ignorise, izvoze se svi\n",
        "T_BankaNalozi_KorpaIIzvoz",
        "izbor od jednog bloka daje tacno jedan nalog",
    ),
    # Normalizacija u cent-domen ide PRE praga "> 0" (AUD-026): sirov ostatak
    # od 0.004 prolazi sirov prag, a u fajlu zavrsi kao nalog na "0.00".
    "banka-nalozi-izvoz-sirov-iznos": (
        "modBankaExportPregled.bas",
        "        blk.IsplatitiIznos = ZaokruziNovac(osnovica)\n",
        "        blk.IsplatitiIznos = osnovica   ' SABOTAZA: sirov iznos\n",
        "T_BankaNalozi_KorpaIIzvoz",
        "0.004 se zaokruzi na 0.00 i NE postaje nalog",
    ),
    # Stavka ciji blok vise nije otvoren mora da IZADJE iz izbora pri
    # uskladjivanju -- traka bi inace pokazivala broj i zbir koji ne postoje.
    "banka-nalozi-usklad-ne-cisti": (
        "modScrBankaNalozi.bas",
        '        If ziviOtvoreno.Exists(CStr(mKorpa(i)("otkupID"))) Then\n'
        '            mKorpa(i)("otvoreno") = CDbl(ziviOtvoreno(CStr(mKorpa(i)("otkupID"))))\n'
        "        Else\n"
        "            mKorpa.Remove i\n"
        "            BnUskladiKorpu = BnUskladiKorpu + 1\n"
        "        End If\n",
        '        If ziviOtvoreno.Exists(CStr(mKorpa(i)("otkupID"))) Then\n'
        '            mKorpa(i)("otvoreno") = CDbl(ziviOtvoreno(CStr(mKorpa(i)("otkupID"))))\n'
        "        End If   ' SABOTAZA: mrtva stavka ostaje u izboru\n",
        "T_BankaNalozi_KorpaIIzvoz",
        "stavka koje nema medju otvorenima izlazi iz izbora",
    ),
    # Zadati iznos preko otvorenog narucuje preplatu. Legacy pravilo iz
    # txtIsplatiti_Exit: sve u cent-domenu, nikad preko otvorenog.
    "banka-nalozi-iznos-preko-otvorenog": (
        "modScrBankaNalozi.bas",
        "    If iznosC > otvorenoC Then\n",
        "    If iznosC > otvorenoC * 1000 Then   ' SABOTAZA: granica pomerena\n",
        "T_BankaNalozi_IznosPoBloku",
        "iznos veci od otvorenog se odbija",
    ),
    # Zaostali zadati iznos (otvoreno se u medjuvremenu smanjilo) mora da se
    # spusti pri SVAKOM citanju -- legacy PruneStaleOverrides pravilo. Bez
    # klampa bi prikaz i potvrda nosili iznos koji vise ne postoji.
    "banka-nalozi-citanje-ne-klampuje": (
        "modScrBankaNalozi.bas",
        "    usklIznosa = BnUskladiIznose(zivi)\n"
        "    If usklIznosa > 0 Then PrijaviUskladjivanjeIznosa usklIznosa\n"
        "\n"
        "    ' Upit se normalizuje JEDNOM, haystack po redu -- v. TekstZaPretragu:\n",
        "    usklIznosa = 0   ' SABOTAZA: zaostali iznosi se ne klampuju\n"
        "\n"
        "    ' Upit se normalizuje JEDNOM, haystack po redu -- v. TekstZaPretragu:\n",
        "T_BankaNalozi_IznosPoBloku",
        "zaostali iznos se pri citanju spusta na otvoreno",
    ),
    # Operater je zadao KOLIKO se placa; izvoz koji to ignorise pravi nalog
    # na pun iznos koji operater nije potvrdio.
    "banka-nalozi-izvoz-ignorise-iznos": (
        "modBankaExportPregled.bas",
        "        osnovica = blk.OtvorenIznos\n"
        "        If Not overrideIznosi Is Nothing Then\n"
        "            If overrideIznosi.Exists(blk.otkupID) Then osnovica = CDbl(overrideIznosi(blk.otkupID))\n"
        "        End If\n",
        "        osnovica = blk.OtvorenIznos   ' SABOTAZA: zadati iznos se ignorise\n",
        "T_BankaNalozi_IznosPoBloku",
        "izvoz nosi zadati iznos, ne pun otvoren",
    ),
    # Goli 18-cifreni racun u CSV koloni Excel cita kao broj (2,059E+17) i
    # drzi samo 15 znacajnih cifara -- snimanje iz Excela racun UNISTI pre
    # uvoza u e-banking. Nalaz sa smoke-a 28.08.2026.
    "banka-csv-racun-goli-broj": (
        "modBankaExportPregled.bas",
        "    If Len(r) <> 18 Then Exit Function\n",
        "    Exit Function   ' SABOTAZA: goli broj ostaje u fajlu\n",
        "T22_RacunUCsvJeExcelSafe",
        "18 golih cifara se kanonizuje u NBS oblik 3-13-2",
    ),
    # Radnji je tacno MAX_ACT (5): sesta se tiho odseca (RefreshRowActions
    # radi Exit For) -- operater dobija ekran kome fali dugme, bez poruke.
    "banka-nalozi-sesta-radnja": (
        "modScrBankaNalozi.bas",
        '                              "bnsve:OTKUI_BTN_BN_SVE:116:ghost:0"\n',
        '                              "bnsve:OTKUI_BTN_BN_SVE:116:ghost:0|" & _\n'
        '                              "bnvisak:OTKUI_BTN_BN_IZNALOG:80:ghost:1"\n',
        "T_BankaNalozi_UgovorEkrana",
        "radnji je TACNO MAX_ACT -- sesta bi se tiho odsekla (peta je izricito 'svi')",
    ),
    # CSV ne knjizi isplatu: blokovi su otvoreni i POSLE fajla, a izbor se
    # posle izvoza prazni. Prazan izbor koji znaci "svi" bi zato na drugi
    # klik tiho izvezao SVE otvorene -- ukljucujuci pun iznos bloka ciji je
    # zadati deo upravo izvezen. Recenzija PR-a, tacka 1 (merge blocker).
    "banka-nalozi-prazan-izbor-izvozi-sve": (
        "modScrBankaNalozi.bas",
        "    outBezTR = 0\n"
        "    outIzbaceno = 0\n"
        "    If BnKorpaBroj() = 0 Then\n"
        "        Set BlokoviZaIzvoz = New Collection\n"
        "        Exit Function\n"
        "    End If\n",
        "    outBezTR = 0\n"
        "    outIzbaceno = 0   ' SABOTAZA: prazan izbor prolazi kao 'svi'\n",
        "T_BankaNalozi_KorpaIIzvoz",
        "prazan izbor ne izvozi nista",
    ),
    # Ekranska putanja (ista koju zovu CSV i PDF) MORA da prosledi zadate
    # iznose -- bez ", Iznosi()" UI pokazuje 250, a fajl nosi 600. Domenska
    # polovina (OdaberiBlokoveZaNaloge) to ne vidi. Recenzija PR-a, tacka 3.
    "banka-nalozi-ekran-ne-salje-iznose": (
        "modScrBankaNalozi.bas",
        "    Set BlokoviZaIzvoz = modBankaExportPregled.OdaberiBlokoveZaNaloge( _\n"
        "                             sveze, BnKorpaIDs(), outBezTR, outIzbaceno, Iznosi())\n",
        "    Set BlokoviZaIzvoz = modBankaExportPregled.OdaberiBlokoveZaNaloge( _\n"
        "                             sveze, BnKorpaIDs(), outBezTR, outIzbaceno)   ' SABOTAZA: bez iznosa\n",
        "T_BankaNalozi_IznosPoBloku",
        "ekran salje zadate iznose izvozu",
    ),
    # Blok kome je racun obrisan POSLE dodavanja u izbor: izvoz ga preskace,
    # ali traka, zbir i potvrda ne smeju da ga pokazuju kao spreman.
    # Recenzija PR-a, tacka 4.
    "banka-nalozi-korpa-drzi-bez-racuna": (
        "modScrBankaNalozi.bas",
        "        If CBool(src(i, 11)) Then zivi(Trim$(CStr(src(i, 1)))) = CDbl(src(i, 9))\n",
        "        zivi(Trim$(CStr(src(i, 1)))) = CDbl(src(i, 9))   ' SABOTAZA: i bez racuna\n",
        "T_BankaNalozi_KorpaIIzvoz",
        "blok bez racuna izlazi iz izbora pri citanju",
    ),
    # Snimak liste se cita JEDNOM pa se filtrira -- pun prolaz kroz tabele po
    # otkucaju je na 1.500+ blokova smrzavao ekran ~10 s po slovu i pretraga
    # je delovala mrtvo (smoke 3, izmereno Diag_BnRedovi).
    "banka-nalozi-pretraga-puni-iznova": (
        "modScrBankaNalozi.bas",
        "    If Not mSnimakOK Then\n"
        "        mSnimakPunjenja = mSnimakPunjenja + 1\n"
        "        mSnimak = modBankaExportPregled.GetBlokIsplataForGrid()\n"
        "        mSnimakOK = True\n"
        "    End If\n",
        "    mSnimakPunjenja = mSnimakPunjenja + 1   ' SABOTAZA: pun prolaz svaki put\n"
        "    mSnimak = modBankaExportPregled.GetBlokIsplataForGrid()\n",
        "T_BankaNalozi_UgovorEkrana",
        "pretraga i cipovi ne placaju pun prolaz -- snimak se cita jednom",
    ),
    # Imena u podacima nose kvake; operater na DE/EN tastaturi kuca bez njih.
    # Haystack bez normalizacije = pretraga koja "ne radi" (smoke 28.08).
    "banka-nalozi-pretraga-sa-kvakama": (
        "modScrBankaNalozi.bas",
        '        hay = modUiData.TekstZaPretragu(CStr(src(i, 2)) & "|" & CStr(src(i, 4)) & "|" & _\n'
        '              CStr(src(i, 6)) & "|" & CStr(src(i, 10)) & "|" & iD)\n',
        '        hay = CStr(src(i, 2)) & "|" & CStr(src(i, 4)) & "|" & _\n'
        "              CStr(src(i, 6)) & \"|\" & CStr(src(i, 10)) & \"|\" & iD   ' SABOTAZA: kvake ostaju\n",
        "T_BankaNalozi_UgovorEkrana",
        "ASCII upit nalazi red sa dijakriticnim imenom",
    ),
    # Red trake i zbir ispod njega moraju da nose ISTI iznos -- onaj koji bi
    # se izvezao. Smoke 28.08: red je pokazivao otvoreno (21.798) uz zbir
    # zadatih (10.000), dva broja jedan ispod drugog koja se ne slazu.
    "banka-nalozi-traka-nosi-otvoreno": (
        "modScrBankaNalozi.bas",
        '    KorpaRedPrikaz = CStr(red("broj")) & "   " & ChrW(183) & "   " & _\n'
        '                     Format$(BnIznosZa(CStr(red("otkupID")), CDbl(red("otvoreno"))), "#,##0")\n',
        '    KorpaRedPrikaz = CStr(red("broj")) & "   " & ChrW(183) & "   " & _\n'
        '                     Format$(CDbl(red("otvoreno")), "#,##0")   \' SABOTAZA: red nosi otvoreno\n',
        "T_BankaNalozi_IznosPoBloku",
        "red trake nosi zadati iznos -- isti broj kao zbir ispod njega",
    ),
    # Zona koja se ne gradi cela: dugme koje fali se ne vidi ni u jednom
    # testu nad citacima -- zato test zonu STVARNO gradi.
    "banka-nalozi-zona-bez-dugmeta": (
        "modScrBankaNalozi.bas",
        '    modUiKit.BtnV z, "scrBnCsv", Poruka("OTKUI_BTN_BN_CSV"), PAD, BN_Y_BTN, _\n'
        '                  164, BN_BTN_H, "primary"\n',
        "    ' SABOTAZA: dugme za naloge se ne gradi\n",
        "T_ZonaBankaNalozi_PoljaIRaspored",
        "zona platnih naloga nema nijednu kontrolu manje",
    ),
    # ------------------------------------------------------------------
    # EKRAN IZVESTAJI (modScrIzvestaji, v6-ui-186). Sabotaze gadjaju EKRANSKU
    # polovinu (izdvajanje, kes, matrica-vodi-liste, prikaz istine); tvrdnje
    # slaganja NAD Report* funkcijama nemaju zasebnu sabotazu -- mutacija
    # modIzvestaj bi obarala i RunIzvestajTests (tudju, postojecu suite), isto
    # pravilo kao "storniran nije u listi" u par. 22.8.
    # ------------------------------------------------------------------
    # Prvi cip je onaj na koji ljuska PADA kad zatecen filter ne pripada
    # listi. Ako nije najsiri, povratak na njega tiho sakrije redove.
    "izvestaji-cip-sve-nije-prvi": (
        "modScrIzvestaji.bas",
        '        Case IZ_MANJAK\n'
        '            IzCipoviZaListu = "sve:OTKUI_CHIP_SVE:40|" & _\n'
        '                              "bezprij:OTKUI_CIPIZ_BEZPRIJ:88"\n',
        '        Case IZ_MANJAK\n'
        '            IzCipoviZaListu = "bezprij:OTKUI_CIPIZ_BEZPRIJ:88|" & _\n'
        '                              "sve:OTKUI_CHIP_SVE:40"\n',
        "T_Izv_UgovorEkrana",
        "prvi cip MANJKA je najsiri ('sve')",
    ),
    # Stampa dokumenta iz reda je razlog postojanja radnje na 4 liste; lista
    # kartice bez nje bi operatera vratila u legacy formu za svaki dokument.
    "izvestaji-kartica-bez-stampe": (
        "modScrIzvestaji.bas",
        "        Case IZ_OTKL, IZ_ROBA, IZ_AMB, IZ_KART\n",
        "        Case IZ_OTKL, IZ_ROBA, IZ_AMB   ' SABOTAZA: kartica bez stampe\n",
        "T_Izv_UgovorEkrana",
        "'Stampaj dokument' nose tacno cetiri liste sa dokument-identitetom",
    ),
    # Kolona identiteta je interna. Prioritet 3 je crta, pa bi operater
    # gledao interni OtkupID -- prikaz kao lazni dokaz da je "sve u redu".
    "izvestaji-identitet-vidljiv": (
        "modScrIzvestaji.bas",
        '                "OTKUI_HD_VREDNOST||rsd|96|1", _\n'
        '                "OTKUI_HDI_REF||txt|1|4")\n',
        '                "OTKUI_HD_VREDNOST||rsd|96|1", _\n'
        '                "OTKUI_HDI_REF||txt|90|3")\n',
        "T_Izv_IdentitetURedu_NeCrtaSe",
        "identitet OTKLISTE se ne crta (prio 4)",
    ),
    # UKUPNO red pod filterom tvrdi zbir koji ne odgovara vidljivim
    # redovima -- filtriran skup sa legacy UKUPNO redom je pogresna brojka
    # na najvidljivijem ekranu.
    # UKUPNO red NIKAD ne ide u mrezu: mreza sortira po koloni, pa je legacy
    # poslednji red PLUTAO usred liste (prvi smoke, lista Isplata); zbir
    # prikazanih daje podnozje, a stampa svoj izracunat UKUPNO.
    "izvestaji-ukupno-prezivi-filter": (
        "modScrIzvestaji.bas",
        "        If vrsta = 1 Then GoTo Sledeci\n",
        "        ' SABOTAZA: UKUPNO ulazi u mrezu\n",
        "T_Izv_IdentitetURedu_NeCrtaSe",
        "UKUPNO red se nikad ne crta u mrezi",
    ),
    # Snimak se cita JEDNOM po kontekstu -- pun Report* prolaz po otkucaju
    # pretrage je placen kvar (par. 22.9/N7: ~10 s po slovu na 1.595 blokova).
    "izvestaji-pretraga-puni-iznova": (
        "modScrIzvestaji.bas",
        "    If Not mSnimci.Exists(k) Then\n"
        "        ' Kapa drzi memoriju: preko granice se krece ispocetka (najprostije\n"
        "        ' ispravno; ResetCache ionako prazni sve posle svakog upisa).\n"
        "        If mSnimci.count >= IZ_SNIMAK_KAPA Then mSnimci.RemoveAll\n"
        "        mSnimakPunjenja = mSnimakPunjenja + 1\n"
        "        mSnimci(k) = PuniSnimak(kljuc, tip, zbirni, iD, odN, doN)\n"
        "    End If\n",
        "    mSnimakPunjenja = mSnimakPunjenja + 1   ' SABOTAZA: pun prolaz svaki put\n"
        "    mSnimci(k) = PuniSnimak(kljuc, tip, zbirni, iD, odN, doN)\n",
        "T_Izv_KesPretragaIHint",
        "tri citanja mreze = JEDNO punjenje snimka (pretraga ne placa pun prolaz)",
    ),
    # Imena u podacima nose kvake; operater na DE/EN tastaturi kuca bez njih
    # (par. 22.9/N3). Haystack bez normalizacije = pretraga koja "ne radi".
    "izvestaji-haystack-sirov": (
        "modScrIzvestaji.bas",
        "            hay = modUiData.TekstZaPretragu(HaystackReda(kljuc, tip, zbirni, src, i))\n",
        "            hay = HaystackReda(kljuc, tip, zbirni, src, i)   ' SABOTAZA: kvake ostaju\n",
        "T_Izv_KesPretragaIHint",
        "ASCII upit nalazi dijakriticno ime (TekstZaPretragu, N3)",
    ),
    # Specijalni red u tipiziranim kolonama mreze: prazne celije postaju
    # "0,00" -- ista klasa lazi kao FM-0028 #5. Red ide u zonu, ne u mrezu.
    "izvestaji-omavans-u-mrezi": (
        "modScrIzvestaji.bas",
        "                ElseIf lbl = IZ_LBL_OM_AVANS Then\n"
        "                    mZonaOmAvans = NzD(src(i, 4))\n"
        "                    VrstaReda = 3\n",
        "                ElseIf lbl = IZ_LBL_OM_AVANS Then\n"
        "                    mZonaOmAvans = NzD(src(i, 4))   ' SABOTAZA: red ostaje u mrezi\n",
        "T_Izv_SlaganjeOtkupOM",
        "OM AVANS red nije u mrezi -- izdvojen je u zonu",
    ),
    # Kontrolna brojka isplate koja se izdvoji BEZ vrednosti: zona pokazuje
    # crtu/nulu dok Report* nosi iznos -- podatak nestane bez traga.
    "izvestaji-zona-isplate-prazna": (
        "modScrIzvestaji.bas",
        "                Case IZ_LBL_ISPL_PRIMLJENO:  mZonaIsplPrimljeno = NzD(src(i, 5)): VrstaReda = 3\n",
        "                Case IZ_LBL_ISPL_PRIMLJENO:  VrstaReda = 3   ' SABOTAZA: zona bez brojke\n",
        "T_Izv_SlaganjeIsplataManjakAmb",
        "zona 'primljeno' = rucni zbir Firma->Otkupac avansa",
    ),
    # Dostupnost lista vodi MATRICA (IzvestajTabDostupan) + legacy uslov za
    # runtime liste. Siri uslov = pun naslov nad izvestajem koji ne postoji
    # za taj tip (FM-0029 #3 klasa).
    "izvestaji-matrica-zaobidjena": (
        "modScrIzvestaji.bas",
        '        Case IZ_OTKL\n'
        '            IzListaDostupna = (tip = "OM" And Not zbirni)\n',
        "        Case IZ_OTKL\n"
        "            IzListaDostupna = (Not zbirni)   ' SABOTAZA: svi tipovi\n",
        "T_Izv_MatricaVodiListe",
        "otk. listovi samo za OM",
    ),
    # Prazna lista bez objasnjenja izgleda kao "nema podataka" -- operater
    # ne sme da dobije pun naslov nad trajno praznom listom bez razloga.
    "izvestaji-hint-bez-razloga": (
        "modScrIzvestaji.bas",
        '        mHintKljuc = "OTKUI_IZ_HINT_NEDOSTUPNO"\n',
        '        mHintKljuc = ""   \' SABOTAZA: prazno bez objasnjenja\n',
        "T_Izv_MatricaVodiListe",
        "prazna lista NOSI objasnjenje zasto je prazna",
    ),
    # Kolona tipa "date" trazi serijski broj; tekst bi RenderGrid prebrojao
    # kao kvar celije i ostavio je praznu (par. 9.9).
    "izvestaji-datum-kao-tekst": (
        "modScrIzvestaji.bas",
        "        Case IZ_OTKL\n"
        "            outA(n, 1) = IzDatCell(src(i, 1))\n",
        "        Case IZ_OTKL\n"
        "            outA(n, 1) = NzS(src(i, 1))   ' SABOTAZA: datum kao tekst\n",
        "T_Izv_UgovorEkrana",
        "datum stize kao serijski broj",
    ),
    # Posle upisa snimak MORA da zastari -- inace ekran pokazuje stanje od
    # pre upisa dok ljuska misli da je osvezila.
    "izvestaji-kes-ne-stari": (
        "modScrIzvestaji.bas",
        "    ' Snimci zastarevaju na svaki upis -- sledece citanje ide u Report*.\n"
        "    Set mSnimci = Nothing\n",
        "    ' SABOTAZA: snimci prezive upis\n",
        "T_Izv_KesPretragaIHint",
        "posle upisa (ResetCache) snimak se puni ponovo",
    ),
    # Prazno kad nema prijema JE poruka (FM-0028 #5) -- nula umesto praznog
    # je bio ceo bug koji je RF-06 zatvorio; ekran ga ne sme vratiti.
    "izvestaji-prazno-postaje-nula": (
        "modScrIzvestaji.bas",
        '                \' Prazno kad nema prijema JE poruka (RF-06) -- ne "0,00".\n'
        "                outA(n, 9) = FmtIliPrazno(src(i, 9))\n",
        '                \' Prazno kad nema prijema JE poruka (RF-06) -- ne "0,00".\n'
        "                outA(n, 9) = FmtKolicina(NzD(src(i, 9)))   ' SABOTAZA: nula umesto praznog\n",
        "T_Izv_IdentitetURedu_NeCrtaSe",
        "red bez prijema ima PRAZNU celiju prijema, ne nulu",
    ),
    # U zbirnom rezimu konkretan entitet ne postoji -- polje koje ostane
    # sugerise da izbor nesto znaci, a ekran ga ignorise.
    "izvestaji-zbirni-drzi-entitet": (
        "modScrIzvestaji.bas",
        '    z.Controls("scrIzEnt").Visible = IzTrebaEntitet(Scr_Lista(), mZbirni)\n',
        '    z.Controls("scrIzEnt").Visible = True   \' SABOTAZA: entitet i u zbirnom\n',
        "T_ZonaIzv_PoljaIRaspored",
        "zbirni rezim gasi polje entiteta",
    ),
    # Izabran tip mora da NOSI rez (Font.Weight) i posle rasporeda --
    # par. 7.7/7.10: bez toga se izabrano i neizabrano ne razlikuju.
    "izvestaji-tip-ne-boji": (
        "modScrIzvestaji.bas",
        "    modUiKit.BoxState z, nm, IIf(sel, C_FOREST, C_WHITE), _\n"
        "                      IIf(sel, C_CREAM, C_FOREST), sel\n",
        "    modUiKit.BoxState z, nm, IIf(sel, C_FOREST, C_WHITE), _\n"
        "                      IIf(sel, C_CREAM, C_FOREST), False   ' SABOTAZA: bez reza\n",
        "T_ZonaIzv_PoljaIRaspored",
        "izabran tip je bold i posle rasporeda",
    ),
    # Dugme kartice na listi bez kartice stampa POGRESAN sablon za pogresan
    # kontekst -- vidljivost prati aktivnu listu.
    "izvestaji-kart-dugme-svuda": (
        "modScrIzvestaji.bas",
        '    naKartici = (Scr_Lista() = IZ_KART Or Scr_Lista() = IZ_AMBK)\n',
        "    naKartici = True   ' SABOTAZA: dugme kartice svuda\n",
        "T_ZonaIzv_PoljaIRaspored",
        "dugme kartice se ne nudi na saldo listi",
    ),
    # ------------------------------------------------------------------
    # Recenzija PR #245 (krug 3): deljeni ugovor invalidacije + politika
    # sabirljivosti stampe + kontekstna radnja.
    # ------------------------------------------------------------------
    # Upis sa DRUGOG ekrana ne zove nas Scr_ResetCache -- bez generacijske
    # provere bi povratak na Izvestaje pokazivao STARE brojke (blocker 1).
    "izvestaji-kes-ignorise-generaciju": (
        "modScrIzvestaji.bas",
        "    ' Upis sa drugog ekrana ne zove nas Scr_ResetCache -- generacija podataka\n"
        "    ' je deljeni signal da je snimljeno stanje mozda staro (blocker 1).\n"
        "    If mSnimakGen <> modUiData.DataGeneracija() Then\n"
        "        Set mSnimci = Nothing\n"
        "        mSnimakGen = modUiData.DataGeneracija()\n"
        "    End If\n",
        "    ' SABOTAZA: kes ignorise generaciju podataka\n",
        "T_KesGeneracija_UpisInvalidira",
        "izvestaji: posle tudjeg upisa snimak se puni ponovo (generacija)",
    ),
    # Isti ugovor na Platnim nalozima -- snimak liste je prezivljavao upis
    # sa drugog ekrana.
    "banka-nalozi-kes-ignorise-generaciju": (
        "modScrBankaNalozi.bas",
        "    ' Upis sa drugog ekrana ne zove nas Scr_ResetCache -- generacija\n"
        "    ' podataka je deljeni signal da je snimak mozda star (PR #245).\n"
        "    If mSnimakGen <> modUiData.DataGeneracija() Then\n"
        "        mSnimakOK = False\n"
        "        mSnimakGen = modUiData.DataGeneracija()\n"
        "    End If\n",
        "    ' SABOTAZA: snimak ignorise generaciju podataka\n",
        "T_KesGeneracija_UpisInvalidira",
        "nalozi: posle tudjeg upisa snimak se puni ponovo (generacija)",
    ),
    # Tip kolone opisuje PRIKAZ, ne aditivnost: zbir running salda kartice
    # nije poslovna vrednost (blocker 2 -- "UKUPNO SALDO = zbir medjustanja").
    "izvestaji-ukupno-sabira-saldo": (
        "modScrIzvestaji.bas",
        "        Case IZ_KART:   IzSabirljive = Array(4, 5)     ' promet; NIKAD saldo (6, 7)\n",
        "        Case IZ_KART:   IzSabirljive = Array(4, 5, 6, 7)   ' SABOTAZA: sabira i salda\n",
        "T_Izv_UgovorEkrana",
        "stampani UKUPNO kartice ne sabira saldo",
    ),
    # Radnja je kontekstna (nalaz 3). Od kruga 5 ROBA za kupca je lista
    # prijemnica sa PRJ| identitetom pa radnju IMA -- vracanje starog gate-a
    # (tip <> "OM") bi je ponovo ugasilo. Vozacki gate ostaje u kodu kao
    # odbrana (matrica vozacku robu ionako ne daje).
    "izvestaji-radnja-na-agregatu": (
        "modScrIzvestaji.bas",
        '    If kljuc = IZ_ROBA And (tip = "Vozac" Or zbirni) Then Exit Function\n',
        '    If kljuc = IZ_ROBA And tip <> "OM" Then Exit Function   \' SABOTAZA: kupac opet bez radnje\n',
        "T_Izv_UgovorEkrana",
        "roba za kupca (prijemnice) ima radnju stampe dokumenta",
    ),
    # ------------------------------------------------------------------
    # Smoke krug 4 (krug 5 ispravki): kontekstni tabovi, prijemnice za
    # kupca, zavrsni saldo kartica.
    # ------------------------------------------------------------------
    # Tab liste koja za tip ne postoji ni u jednom rezimu je mrtvo dugme --
    # skup tabova MORA da prati matricu po tipu.
    "izvestaji-tabovi-ne-slusaju-tip": (
        "modScrIzvestaji.bas",
        "Public Function IzListaZaTipPostoji(ByVal kljuc As String, ByVal tip As String) As Boolean\n"
        "    IzListaZaTipPostoji = IzListaDostupna(kljuc, tip, False) Or _\n"
        "                          IzListaDostupna(kljuc, tip, True)\n",
        "Public Function IzListaZaTipPostoji(ByVal kljuc As String, ByVal tip As String) As Boolean\n"
        "    IzListaZaTipPostoji = True   ' SABOTAZA: svi tabovi za svaki tip\n",
        "T_Izv_TabKontekstRobaKupacSaldo",
        "OM ne nudi tabove kartica",
    ),
    # ROBA za kupca su DOKUMENTA (prijemnice), ne agregat po vrsti. Sabotira
    # se GRANA OBLIKOVANJA (UpisiRed): kupac tretiran kao agregat cita prve
    # cetiri kolone snimka, pa kg/vrednost mreze gube vezu sa tblPrijemnica
    # -- slaganje sa rucnim prolazom pada po imenu. (Sabotaza na samom
    # PuniSnimak pozivu bi pukla kao Subscript pre tvrdnje -- vidljiv kvar,
    # ali ne imenovan; zato se meri ovde.)
    "izvestaji-roba-kupac-agregat": (
        "modScrIzvestaji.bas",
        '            ElseIf tip = "Kupac" Then\n'
        "                ' Prijemnice kupca (ReportPrijemniceKupca fiksne kolone).\n",
        "            ElseIf False Then   ' SABOTAZA: kupac tretiran kao agregat\n"
        "                ' Prijemnice kupca (ReportPrijemniceKupca fiksne kolone).\n",
        "T_Izv_TabKontekstRobaKupacSaldo",
        "kg robe kupca = rucni zbir prijemnica",
    ),
    # Zavrsni saldo kartice u zoni MORA doci iz kolone salda UKUPNO reda --
    # promet perioda (kol. 5) je druga brojka i tiho bi lagao operatera.
    "izvestaji-kart-saldo-pogresna-kolona": (
        "modScrIzvestaji.bas",
        "                mZonaKartSaldo = NzD(src(i, 7))\n",
        "                mZonaKartSaldo = NzD(src(i, 5))   ' SABOTAZA: promet umesto salda\n",
        "T_Izv_TabKontekstRobaKupacSaldo",
        "zona saldo = zavrsni running saldo kartice",
    ),
    # Matrica i za KUPCE zbirno (krug 11): vracanje stare grane (bez
    # SALDO_KUPCI/OTKUP_ROBA) mora da padne po imenu.
    "izvestaji-kupci-zbirno-van-matrice": (
        "modIzvestaj.bas",
        "                    Case IZV_TAB_ZBIRNI, IZV_TAB_MANJAK, IZV_TAB_AMBALAZA, _\n"
        "                         IZV_TAB_SALDO_KUPCI, IZV_TAB_OTKUP_ROBA\n",
        "                    Case IZV_TAB_ZBIRNI, IZV_TAB_MANJAK, IZV_TAB_AMBALAZA   ' SABOTAZA: kupci bez zbirnog salda/robe\n",
        "T_Izv_MatricaVodiListe",
        "saldo kupaca zbirno po kupcima (krug 11)",
    ),
    # Roba po vozacu zbirno (krug 12): vracanje stare vozacke grane
    # (bez OTKUP_ROBA) mora da padne po imenu.
    "izvestaji-vozaci-roba-van-matrice": (
        "modIzvestaj.bas",
        "                    Case IZV_TAB_ZBIRNI, IZV_TAB_MANJAK, IZV_TAB_AMBALAZA, _\n"
        "                         IZV_TAB_OTKUP_ROBA\n",
        "                    Case IZV_TAB_ZBIRNI, IZV_TAB_MANJAK, IZV_TAB_AMBALAZA   ' SABOTAZA: vozaci bez zbirne robe\n",
        "T_Izv_MatricaVodiListe",
        "roba po vozacu zbirno (krug 12)",
    ),
    # Roba po vozacu MERI otpremnice bez storniranih -- storno filter
    # koji tiho nestane duplira prevoz.
    "izvestaji-roba-vozaci-storno": (
        "modIzvestaj.bas",
        "        If cStorno = 0 Or CStr(d(i, cStorno)) <> \"Da\" Then\n"
        "            If IsDate(d(i, cDat)) Then\n"
        "                dv = CDate(d(i, cDat))\n"
        "                If dv >= datumOd And dv <= datumDo Then\n"
        "                    k = Trim$(CStr(d(i, cVoz)))\n",
        "        If True Then   ' SABOTAZA: i stornirane otpremnice\n"
        "            If IsDate(d(i, cDat)) Then\n"
        "                dv = CDate(d(i, cDat))\n"
        "                If dv >= datumOd And dv <= datumDo Then\n"
        "                    k = Trim$(CStr(d(i, cVoz)))\n",
        "T_Izv_ZbirniSadrzaj",
        "roba po vozacu: kg = rucni zbir otpremnica",
    ),
    # Rang se OTVARA po rangu rastuce -- shell sort ugovor (recenzija
    # #245 blocker: izvor sortiran, a ekran presortira po imenu).
    "izvestaji-rang-sort-ime": (
        "modOtkupUI.bas",
        "    If kljuc = \"KOOPERANTI\" Or kljuc = \"RANG\" Then\n",
        "    If kljuc = \"KOOPERANTI\" Then   ' SABOTAZA: rang pada u datum-desc granu\n",
        "T_Izv_RangSortIKontekst",
        "rang se otvara po koloni ranga",
    ),
    # Kontekst "Svi" prati LISTU (IzTrebaEntitet), ne rezim -- rang u
    # pojedinacnom ne sme da ispise prazan entitet "()".
    "izvestaji-rang-kontekst-prazan": (
        "modScrIzvestaji.bas",
        "    mCtxEntNaziv = EntitetNaziv(tip, iD, Not IzTrebaEntitet(kljuc, zbirni))\n",
        "    mCtxEntNaziv = EntitetNaziv(tip, iD, zbirni)   ' SABOTAZA: Svi po rezimu\n",
        "T_Izv_RangSortIKontekst",
        "rang kontekst nije prazan entitet '()'",
    ),
    # Univerzum "Svi OM" dolazi IZ PODATAKA -- povratak na sifarnik
    # tiho gubi orphan stanicu (silent omission u finansijskom zbiru).
    "izvestaji-om-univerzum-sifarnik": (
        "modIzvestaj.bas",
        "    IzvStaniceUnion dict, TBL_OTKUP, COL_OTK_STANICA, \"\", \"\"\n",
        "    ' SABOTAZA: otkupi ne sire univerzum stanica\n",
        "T_Izv_ZbirniOrphanStanica",
        "orphan stanica POSTOJI u zbirnom saldu -- silent omission je kvar",
    ),
    # Aktivacija ekrana primenjuje PODRAZUMEVANI sort aktivne liste --
    # tvrdi reset na 2/desc vraca rang-po-imenu posle povratka na ekran
    # (recenzija #245, krug 17 lifecycle blocker).
    "izvestaji-aktivacija-gazi-sort": (
        "modOtkupUI.bas",
        "Public Sub GridSortAktivacijaTest()\n"
        "    If Not IsTestMode() Then Exit Sub\n"
        "    PrimeniSortZaListu ActiveLista()\n"
        "End Sub\n",
        "Public Sub GridSortAktivacijaTest()\n"
        "    If Not IsTestMode() Then Exit Sub\n"
        "    mSortCol = 2: mSortAsc = False   ' SABOTAZA: tvrdi reset kao pre\n"
        "End Sub\n",
        "T_Izv_RangSortIKontekst",
        "povratak na ekran vraca rang na kolonu ranga",
    ),
    # Cip vrste MORA da filtrira po vrednosti reda -- cip koji sve
    # propusta je laz na ekranu (krug 18).
    "izvestaji-cip-vrste-ne-filtrira": (
        "modScrIzvestaji.bas",
        "        CipPropusta = (StrComp(IzVrstaIzReda(kljuc, tip, src, i), _\n"
        "                               Mid$(filter, 3), vbTextCompare) = 0)\n",
        "        CipPropusta = True   ' SABOTAZA: cip vrste propusta sve\n",
        "T_Izv_CipoviVrstaSorta",
        "nepostojeca vrsta = nula redova",
    ),
    # Rang u Izvestajima POSTUJE period zone (nova Optional grana u
    # KoopRangRows) -- bez filtera bi hint tvrdio period koji rang ne
    # primenjuje. Legacy pozivaoci (bez granica) sabotiranu granu ne
    # dodiruju, pa pada samo test ranga u Izvestajima.
    "izvestaji-rang-mimo-perioda": (
        "modOtkupBlok.bas",
        "                    uKrug = (odN = 0 Or dSer >= odN) And (doN = 0 Or dSer <= doN)\n",
        "                    uKrug = True   ' SABOTAZA: rang ignorise period\n",
        "T_Izv_RangKooperanata",
        "rang postuje period -- prazan opseg nema redove",
    ),
    # ------------------------------------------------------------------
    # Krug 9: zbirni sadrzaj ("fali sadrzaj za zbirne izvestaje").
    # ------------------------------------------------------------------
    # Matrica je izvor istine i za NOVE zbirne kombinacije -- vracanje
    # starih grana (bez SALDO/AMBALAZA/ISPLATA zbirno za OM) mora da padne
    # po imenu, ne da tiho suzi ekran.
    "izvestaji-zbirno-van-matrice": (
        "modIzvestaj.bas",
        "                    Case IZV_TAB_ZBIRNI, IZV_TAB_PROSECNA_CENA, IZV_TAB_MANJAK, _\n"
        "                         IZV_TAB_SALDO_OM, IZV_TAB_AMBALAZA, IZV_TAB_ISPLATA, _\n"
        "                         IZV_TAB_OTKUP_ROBA\n",
        "                    Case IZV_TAB_ZBIRNI, IZV_TAB_PROSECNA_CENA, IZV_TAB_MANJAK   ' SABOTAZA: stara matrica\n",
        "T_Izv_MatricaVodiListe",
        "saldo zbirno po stanicama (OM)",
    ),
    # Red zbirnog salda je UKUPNO red pojedinacnog izvestaja te stanice --
    # prvi red (prvi kooperant) umesto UKUPNO bi tiho lagao po stanici.
    "izvestaji-zbirni-saldo-tudji-red": (
        "modIzvestaj.bas",
        "            r = ReportSaldoOM(stID, datumOd, datumDo)\n"
        "            uk = IzvUkupnoRed(r, 1)\n",
        "            r = ReportSaldoOM(stID, datumOd, datumDo)\n"
        "            uk = 1   ' SABOTAZA: prvi red umesto UKUPNO\n",
        "T_Izv_ZbirniSadrzaj",
        "zbirni saldo: kg stanice = rucni prolaz tblOtkup",
    ),
    # Promena rezima MORA da prebaci listu koje u novom rezimu nema --
    # inace je prvi utisak zbirnog rezima prazan ekran sa hintom (krug 9).
    "izvestaji-rezim-bez-prelaza": (
        "modScrIzvestaji.bas",
        "    If Not IzListaDostupna(Scr_Lista(), TrenutniTip(), zbirni) Then\n"
        "        mLista = PrvaListaZaKontekst(TrenutniTip(), zbirni)\n"
        "    End If\n",
        "    ' SABOTAZA: rezim ne prebacuje listu\n",
        "T_Izv_ZbirniSadrzaj",
        "prelaz na zbirno sa otk. listova ide na prvu dostupnu (saldo)",
    ),
    # ------------------------------------------------------------------
    # Smoke krug 3 (Izvestaji): kontekstni cipovi, detalj reda, poslovni
    # broj dokumenta u pregledu ambalaze.
    # ------------------------------------------------------------------
    # Cip nad listom koja za kombinaciju NE POSTOJI je filter necega cega
    # nema -- ne sme ni da se vidi (isti princip kao kontekstna radnja).
    "izvestaji-cip-na-nedostupnoj": (
        "modScrIzvestaji.bas",
        "Public Function IzCipoviZaKontekst(ByVal kljuc As String, ByVal tip As String, _\n"
        "                                   ByVal zbirni As Boolean) As String\n"
        "    If Not IzListaDostupna(kljuc, tip, zbirni) Then Exit Function\n",
        "Public Function IzCipoviZaKontekst(ByVal kljuc As String, ByVal tip As String, _\n"
        "                                   ByVal zbirni As Boolean) As String\n"
        "    ' SABOTAZA: cipovi i na nedostupnoj kombinaciji\n",
        "T_Izv_DetaljICipKontekst",
        "nedostupna kombinacija nema cipove",
    ),
    # Detalj reda je legacy "Detalji otkupa": SVE stavke dokumenta (broj +
    # stanica), ne samo izabrana linija -- Klasa I i II dele dokument.
    "izvestaji-detalj-bez-stavki": (
        "modScrIzvestaji.bas",
        "        If NzS(d(i, cBr)) = brDok And NzS(d(i, cSt)) = stanica Then\n",
        "        If Trim$(CStr(d(i, cId))) = Trim$(otkupID) Then   ' SABOTAZA: samo izabrana linija\n",
        "T_Izv_DetaljICipKontekst",
        "detalj nosi SVE stavke bloka",
    ),
    # Pregled ambalaze pokazuje POSLOVNI broj dokumenta; bez mape prijemnica
    # red nosi interni ID -- operater njime ne moze nista (par. 9.5 princip).
    "izvestaji-amb-broj-ostaje-id": (
        "modIzvestaj.bas",
        "        Case DOK_TIP_PRIJEMNICA\n"
        "            If mapaPrj.Exists(dokID) Then sOut = CStr(mapaPrj(dokID))\n",
        "        Case DOK_TIP_PRIJEMNICA\n"
        "            ' SABOTAZA: prijemnica ostaje interni ID\n",
        "T_Izv_DetaljICipKontekst",
        "ambalaza pokazuje poslovni broj dokumenta, ne interni ID",
    ),
    # ------------------------------------------------------------------
    # Ekran SLEDLJIVOST (v6-ui-187). Lanac koji se ne izmislja: zbirna se
    # cita ISKLJUCIVO iz otpremnice, dvosmislen broj i odsutan prijem su
    # fail-closed oznake, kg curenje je vidljivo, kes snimka i pretraga po
    # N7/R1 pravilima. Report* polovina NEMA pokrice u tudjim suite-ovima
    # (nove funkcije), pa sabotaze gadjaju i nju, ne samo ekran.
    # ------------------------------------------------------------------
    # Otpremnica bez zbirne NE sme da se premosti kroz blokov denorm
    # BrojZbirne -- tacno to premoscenje je klasa laznog lanca koju merilo
    # zadatka zabranjuje. (OTK-TEST-2 tvrdi ZB-TEST-3, otpremnica nema.)
    "sledljivost-lanac-premoscuje-zbirnu": (
        "modIzvestaj.bas",
        "            If Len(blokZbr) > 0 And UCase$(blokZbr) <> UCase$(brZbr) Then\n"
        "                oznaka = SLED_OZN_VEZA\n",
        "            If Len(brZbr) = 0 And Len(blokZbr) > 0 Then\n"
        "                brZbr = blokZbr   ' SABOTAZA: premoscenje kroz blokov denorm\n",
        "T_Sled_FailClosed",
        "raskorak blok/otpremnica zbirne se prijavljuje",
    ),
    # Broj koji dele dva vlasnika tretiran kao jednoznacan: fail-closed
    # kapija vlasnistva pada, red dobija pogresnu oznaku umesto
    # IZV_VLASNIK_NEJASAN.
    "sledljivost-dvosmislen-broj-sabira": (
        "modIzvestaj.bas",
        "    ElseIf nVlasnika > 1 Then\n"
        "        If manjakDict.Exists(\"#O|\" & brZbr & \"|\" & vozID) Then\n",
        "    ElseIf nVlasnika > 1 Then\n"
        "        nVlasnika = 1   ' SABOTAZA: dvosmislen broj tretiran kao jedan vlasnik\n"
        "        If manjakDict.Exists(\"#O|\" & brZbr & \"|\" & vozID) Then\n",
        "T_Sled_FailClosed",
        "dvosmislen broj daje oznaku",
    ),
    # Prijem kg u redu mora biti bas zbir prijemnica razresenog scope-a --
    # nula umesto njega je izmisljen lanac bez robe.
    "sledljivost-prijem-kg-nula": (
        "modIzvestaj.bas",
        "                        result(r, 11) = CDbl(pz(1))\n",
        "                        result(r, 11) = 0#   ' SABOTAZA: prijem kg se gubi\n",
        "T_Sled_LanacSlaganje",
        "prijem kg u redu = rucni zbir prijemnica",
    ),
    # Kg razlika na karici blok<->otpremnica mora biti VIDLJIVA oznaka --
    # prag koji je proguta je precutano curenje (merilo #2 zadatka).
    "sledljivost-kg-razlika-tiha": (
        "modIzvestaj.bas",
        "            If blokSum.Exists(otpID) Then\n"
        "                If Abs(CDbl(blokSum(otpID)) - otpKg) > SLED_EPS_KG Then kg1 = True\n"
        "            End If\n",
        "            If blokSum.Exists(otpID) Then\n"
        "                If Abs(CDbl(blokSum(otpID)) - otpKg) > 1000000# Then kg1 = True   ' SABOTAZA: prag guta razliku\n"
        "            End If\n",
        "T_Sled_FailClosed",
        "kg curenje na karici nosi oznaku",
    ),
    # Storniran otkup ne sme u lanac -- filter, ne odsustvo reda.
    "sledljivost-storniran-ulazi": (
        "modIzvestaj.bas",
        "    otkupData = GetTableData(TBL_OTKUP)\n"
        "    If Not IsArray(otkupData) Then Exit Function\n"
        "    otkupData = ExcludeStornirano(otkupData, TBL_OTKUP)\n"
        "    If Not IsArray(otkupData) Then Exit Function\n"
        "\n"
        "    Dim cOtkId As Long, cOtkDat As Long, cOtkKoop As Long, cOtkSt As Long\n",
        "    otkupData = GetTableData(TBL_OTKUP)\n"
        "    If Not IsArray(otkupData) Then Exit Function\n"
        "    ' SABOTAZA: stornirani otkupi ulaze u lanac\n"
        "    If Not IsArray(otkupData) Then Exit Function\n"
        "\n"
        "    Dim cOtkId As Long, cOtkDat As Long, cOtkKoop As Long, cOtkSt As Long\n",
        "T_Sled_FailClosed",
        "storniran otkup nije u lancu",
    ),
    # Klasa problema koja ispadne iz liste je nevidljiv posao koji ceka --
    # lista NEPOTPUNI je pregled tog posla. Gasi se CEO prolaz prijemnica
    # (sabotaza samo prve grane ne bi oborila nista: ElseIf "Da bez
    # FakturaID" grana bi nefakturisane svejedno uhvatila, pod istom
    # klasom).
    "sledljivost-problemi-gube-klasu": (
        "modIzvestaj.bas",
        "    If IsArray(prijData) Then prijData = ExcludeStornirano(prijData, TBL_PRIJEMNICA)\n"
        "    If IsArray(prijData) Then\n"
        "        Dim cPId As Long, cPBr As Long, cPKup As Long, cPKol As Long\n",
        "    If IsArray(prijData) Then prijData = ExcludeStornirano(prijData, TBL_PRIJEMNICA)\n"
        "    If False Then   ' SABOTAZA: prijemnice ispadaju iz liste problema\n"
        "        Dim cPId As Long, cPBr As Long, cPKup As Long, cPKol As Long\n",
        "T_Sled_FailClosed",
        "problem: Fakturisano=Da bez FakturaID",
    ),
    # Kg razlika u listi problema ima svoj prag nezavisno od lanca.
    "sledljivost-problem-kg-prag": (
        "modIzvestaj.bas",
        "                sumB = CDbl(blokSum(oid))\n"
        "                If Abs(sumB - otpKg) > SLED_EPS_KG Then\n",
        "                sumB = CDbl(blokSum(oid))\n"
        "                If Abs(sumB - otpKg) > 1000000# Then   ' SABOTAZA: prag guta razliku karike\n",
        "T_Sled_FailClosed",
        "problem: kg razlika na otpremnici",
    ),
    # Identitet reda se NE crta (prio 4) -- vidljiv interni kljuc je
    # par. 8.5 klasa kvara.
    "sledljivost-identitet-vidljiv": (
        "modScrSledljivost.bas",
        "                \"OTKUI_HDS_STANJE||txt|78|2\", _\n"
        "                \"OTKUI_HDI_REF||txt|1|4\")\n",
        "                \"OTKUI_HDS_STANJE||txt|78|2\", _\n"
        "                \"OTKUI_HDI_REF||txt|1|1\")   ' SABOTAZA: identitet se crta\n",
        "T_Sled_IdentitetURedu_NeCrtaSe",
        "identitet LANAC je prio 4",
    ),
    # Vrsta karike vodi rutu stampe: zbirna pod tudjom vrstom bi stampala
    # TUDJI dokument umesto da odbije.
    "sledljivost-radnja-tudja-vrsta": (
        "modIzvestaj.bas",
        "                    rows.Add Array(SLEDP_BEZ_PRIJEMA, zbrData(i, cZDat), zBr, naziv, _\n"
        "                                   zKg, \"nijedna prijemnica za broj \" & zBr & _\n"
        "                                   \" (klasa \" & zKla & \")\", _\n"
        "                                   SLED_DOK_ZBIRNA, SledTxt(zbrData(i, cZId)), \"\")\n",
        "                    rows.Add Array(SLEDP_BEZ_PRIJEMA, zbrData(i, cZDat), zBr, naziv, _\n"
        "                                   zKg, \"nijedna prijemnica za broj \" & zBr & _\n"
        "                                   \" (klasa \" & zKla & \")\", _\n"
        "                                   DOK_TIP_PRIJEMNICA, SledTxt(zbrData(i, cZId)), \"\")   ' SABOTAZA: tudja vrsta karike\n",
        "T_Sled_IdentitetURedu_NeCrtaSe",
        "karika zbirne nosi vrstu koja odbija",
    ),
    # Snimak se puni JEDNOM po kontekstu -- pun prolaz po otkucaju je
    # placen kvar (par. 22.9/N7).
    "sledljivost-kes-puni-iznova": (
        "modScrSledljivost.bas",
        "    If Not mSnimci.Exists(k) Then\n"
        "        If mSnimci.count >= SL_SNIMAK_KAPA Then mSnimci.RemoveAll\n",
        "    If True Then   ' SABOTAZA: svaki poziv puni iznova\n"
        "        If mSnimci.count >= SL_SNIMAK_KAPA Then mSnimci.RemoveAll\n",
        "T_Sled_KesPretragaIHint",
        "JEDNO punjenje snimka",
    ),
    # Upis sa DRUGOG ekrana ne prolazi kroz nas Scr_ResetCache -- generacija
    # podataka je jedini signal (par. 23.10/R1).
    "sledljivost-kes-ignorise-generaciju": (
        "modScrSledljivost.bas",
        "    If mSnimakGen <> modUiData.DataGeneracija() Then\n"
        "        Set mSnimci = Nothing\n"
        "        mSnimakGen = modUiData.DataGeneracija()\n"
        "    End If\n",
        "    If False Then   ' SABOTAZA: tudji upis ne invalidira snimak\n"
        "        Set mSnimci = Nothing\n"
        "        mSnimakGen = modUiData.DataGeneracija()\n"
        "    End If\n",
        "T_Sled_KesPretragaIHint",
        "generacija podataka invalidira snimak",
    ),
    # Kvake u podacima, ASCII na tastaturi operatera (par. 22.9/N3): obe
    # strane poredjenja idu kroz TekstZaPretragu.
    "sledljivost-pretraga-sirovi-haystack": (
        "modScrSledljivost.bas",
        "            hay = modUiData.TekstZaPretragu(HaystackReda(kljuc, src, i))\n",
        "            hay = HaystackReda(kljuc, src, i)   ' SABOTAZA: sirov haystack, kvake ostaju\n",
        "T_Sled_KesPretragaIHint",
        "ASCII upit nalazi kooperanta sa kvakama",
    ),
    # Cip "nepotpun" je filter oznake, ne ukras -- pusta li sve, operater
    # misli da gleda nalaze a gleda ceo spisak.
    "sledljivost-cip-nepotpun-pusta-sve": (
        "modScrSledljivost.bas",
        "        Case \"nepotpun\": SlCipLanac = (Len(Trim$(oznaka)) > 0)\n",
        "        Case \"nepotpun\": SlCipLanac = True   ' SABOTAZA: cip pusta sve\n",
        "T_Sled_FailClosed",
        "cip nepotpun NE pusta potpun lanac",
    ),
    # Cip "bez parcele" mora da iskljuci blokove SA parcelom -- inace
    # sertifikaciona rupa izgleda pokrivena.
    "sledljivost-cip-bezpar-pusta-sve": (
        "modScrSledljivost.bas",
        "        Case \"bezpar\": SlCipParcele = (Len(Trim$(parcelaID)) = 0)\n",
        "        Case \"bezpar\": SlCipParcele = True   ' SABOTAZA: cip pusta i sa parcelom\n",
        "T_Sled_LanacSlaganje",
        "cip bez parcele NE propusta blok sa parcelom",
    ),
    # Prvi cip je svuda najsiri -- ljuska na njega pada kad zatecen filter
    # ne pripada listi (RefreshChipsForScreen).
    "sledljivost-cip-sve-nije-prvi": (
        "modScrSledljivost.bas",
        "            SlCipoviZaListu = \"sve:OTKUI_CHIP_SVE:40|\" & _\n"
        "                              \"potpun:OTKUI_CIPSL_POTPUN:86|\" & _\n",
        "            SlCipoviZaListu = \"potpun:OTKUI_CIPSL_POTPUN:86|\" & _\n"
        "                              \"sve:OTKUI_CHIP_SVE:40|\" & _\n",
        "T_Sled_UgovorEkrana",
        "je najsiri ('sve')",
    ),
    # Zona koja izgubi kontrolu tiho -- dugme "ne postoji" bez ijedne
    # greske.
    "sledljivost-zona-bez-dugmeta": (
        "modScrSledljivost.bas",
        "    modUiKit.BtnV z, \"scrSlLanac\", Poruka(\"OTKUI_BTN_SL_LANACPDF\"), PAD + 164, SL_Y_BTN, _\n"
        "                  120, SL_BTN_H, \"soft\"\n",
        "    ' SABOTAZA: dugme lanca se ne gradi\n",
        "T_ZonaSled_PoljaIRaspored",
        "zona sledljivosti nema nijednu kontrolu manje",
    ),
    # PDF lanca bez oznake kompletnosti izgleda kao potpun lanac na papiru
    # -- tacno ono sto kontekst-linija i kolona STATUS sprecavaju.
    "sledljivost-pdf-bez-oznake": (
        "modScrSledljivost.bas",
        "    Dim ozn As String\n"
        "    ozn = NzS(lanac(r, 14))\n",
        "    Dim ozn As String\n"
        "    ozn = \"\"   ' SABOTAZA: PDF lanca gubi oznaku kompletnosti\n",
        "T_Sled_IdentitetURedu_NeCrtaSe",
        "PDF lanac koji curi nosi oznaku",
    ),
    # KPI problema mora iz LISTE PROBLEMA -- brojka iz pogresnog izvora
    # laze operatera o velicini posla.
    "sledljivost-kpi-problemi-iz-lanca": (
        "modScrSledljivost.bas",
        "    If IsArray(problemi) Then\n"
        "        mKpiProblemi = UBound(problemi, 1)\n",
        "    If IsArray(problemi) Then\n"
        "        mKpiProblemi = potpunih   ' SABOTAZA: KPI problema iz pogresnog izvora\n",
        "T_Sled_LanacSlaganje",
        "KPI problema = broj redova liste problema",
    ),
    # Detalj trake nosi karike lanca -- bez otpremnice je "pun lanac" od
    # jedne linije.
    "sledljivost-detalj-bez-karika": (
        "modScrSledljivost.bas",
        "    If Len(NzS(lanac(r, 8))) > 0 Then\n"
        "        linije.Add Poruka(\"OTKUI_IZ_DET_OTPREMNICA\") & \" \" & NzS(lanac(r, 8)) & _\n",
        "    If False Then   ' SABOTAZA: detalj gubi kariku otpremnice\n"
        "        linije.Add Poruka(\"OTKUI_IZ_DET_OTPREMNICA\") & \" \" & NzS(lanac(r, 8)) & _\n",
        "T_Sled_IdentitetURedu_NeCrtaSe",
        "detalj nosi otpremnicu",
    ),
    # Kandidati za rucno povezivanje bez filtera stanice -- tudja
    # otpremnica (OTP-LEG-B, druga stanica, isti datum) usla bi u izbor
    # i pogresno povezivanje bilo bi na klik. Mutiraju se OBA prolaza
    # (brojanje + punjenje) jednim sidrom: mutacija samo jednog bi dala
    # subscript crash umesto imenovanog pada.
    "sledljivost-kandidati-bez-stanice": (
        "modSledljivost.bas",
        "    Dim count As Long\n"
        "    For i = 1 To UBound(otpData, 1)\n"
        "        If Trim$(CStr(otpData(i, cSt))) = stanicaID And IsDate(otpData(i, cDat)) Then\n"
        "            If CDate(otpData(i, cDat)) = datum Then count = count + 1\n"
        "        End If\n"
        "    Next i\n"
        "    If count = 0 Then Exit Function\n"
        "\n"
        "    Dim result() As Variant, idx As Long\n"
        "    ReDim result(1 To count, 1 To 5)\n"
        "    For i = 1 To UBound(otpData, 1)\n"
        "        If Trim$(CStr(otpData(i, cSt))) = stanicaID And IsDate(otpData(i, cDat)) Then\n",
        "    Dim count As Long\n"
        "    For i = 1 To UBound(otpData, 1)\n"
        "        If Len(stanicaID) >= 0 And IsDate(otpData(i, cDat)) Then   ' SABOTAZA: bez stanice\n"
        "            If CDate(otpData(i, cDat)) = datum Then count = count + 1\n"
        "        End If\n"
        "    Next i\n"
        "    If count = 0 Then Exit Function\n"
        "\n"
        "    Dim result() As Variant, idx As Long\n"
        "    ReDim result(1 To count, 1 To 5)\n"
        "    For i = 1 To UBound(otpData, 1)\n"
        "        If Len(stanicaID) >= 0 And IsDate(otpData(i, cDat)) Then   ' SABOTAZA: bez stanice\n",
        "T_Sled_PovezivanjeKandidati",
        "kandidati su samo sa stanice otkupa",
    ),
    # Mete sledljivosti bez storno filtera na paleti -- stornirana paleta
    # (PAL-SLED-X, cija stavka NIJE stornirana) postala bi dokument
    # sledljivosti nepostojece robe.
    "sledljivost-mete-storniranu-paletu": (
        "modIzvestaj.bas",
        "        palData = GetTableData(TBL_PALETA)\n"
        "        If IsArray(palData) Then palData = ExcludeStornirano(palData, TBL_PALETA)\n",
        "        palData = GetTableData(TBL_PALETA)\n"
        "        ' SABOTAZA: stornirane palete ulaze u mete\n",
        "T_Sled_MeteSledljivosti",
        "stornirana paleta nije meta sledljivosti",
    ),
    # Preradjena paleta ponudjena kao "sveza roba" -- operater bi dobio
    # paletni list za robu koje u magacinu sveze robe vise nema.
    "sledljivost-mete-preradjena-kao-sveza": (
        "modIzvestaj.bas",
        "                    If UCase$(Trim$(SledTxt(palData(i, cPalPre)))) <> \"DA\" Then\n",
        "                    If True Then   ' SABOTAZA: preradjena kao sveza\n",
        "T_Sled_MeteSledljivosti",
        "preradjena paleta nije meta 'sveze robe'",
    ),
    # Ponuda polja izbora (krug 3b) bez storno filtera na paleti --
    # stornirana paleta bi usla u dropdown i klik bi stampao list
    # nepostojece robe. Sidro je 4-space varijanta (ReportSledljivost-
    # Dokumenti); 8-space zivi u ReportSledljivostMete i ima svoju
    # sabotazu.
    "sledljivost-dokumenti-storniranu-paletu": (
        "modIzvestaj.bas",
        "    palData = GetTableData(TBL_PALETA)\n"
        "    If IsArray(palData) Then palData = ExcludeStornirano(palData, TBL_PALETA)\n",
        "    palData = GetTableData(TBL_PALETA)\n"
        "    ' SABOTAZA: stornirane palete ulaze u ponudu\n",
        "T_Sled_DokumentiPonuda",
        "stornirana paleta nije u ponudi",
    ),
    # Krug 8 R1 / krug 9: prijemnica koja TVRDI "Da" bez validne aktivne
    # fakture mora da obori kariku -- kad ta grana cuti, kontradikcija
    # izgleda kao potpun lanac.
    "sledljivost-fakture-any-umesto-all": (
        "modIzvestaj.bas",
        "                                    Else\n"
        "                                        prijLose = True\n"
        "                                        If Len(p(4)) > 0 Then refs = refs & \"|\" & p(4)\n"
        "                                    End If\n",
        "                                    Else\n"
        "                                        ' SABOTAZA: neispravna tvrdnja cuti\n"
        "                                        If Len(p(4)) > 0 Then refs = refs & \"|\" & p(4)\n"
        "                                    End If\n",
        "T_Sled_FailClosed",
        "tvrdnja Da bez validne fakture obara kariku (ALL nad tvrdnjama)",
    ),
    # Krug 8 R2: LANAC haystack bez SearchRefs kolone -- broj progutane
    # fakture ("2 fakt.") vise ne nalazi red, obecanje pretrage pada.
    "sledljivost-lanac-pretraga-bez-refs": (
        "modScrSledljivost.bas",
        "                           NzS(src(i, 14)) & \"|\" & NzS(src(i, 27)) & \"|\" & _\n"
        "                           NzS(src(i, 30))\n",
        "                           NzS(src(i, 14)) & \"|\" & _\n"
        "                           NzS(src(i, 30))   ' SABOTAZA: bez SearchRefs\n",
        "T_Sled_KesPretragaIHint",
        "broj progutane fakture nalazi LANAC red",
    ),
    # Prikaz "N pal."/"N pre." guta brojeve -- bez refs-a smer NAZAD od
    # broja palete/prerade/GP fakture ne nalazi nista (princip R2).
    "sledljivost-gp-refs-progutani": (
        "modIzvestaj.bas",
        "                        If Len(gpRefs) > 0 Then\n",
        "                        If False And Len(gpRefs) > 0 Then   ' SABOTAZA: GP brojevi progutani\n",
        "T_Sled_GpLanacIStanja",
        "refs nose broj palete",
    ),
    # =============== krug 5 (utovarna lista -- novi prodajni grain)
    # Kapija stanja je JEDINA brana prekomerne prodaje u writeru.
    "utovar-gp-stanje-kapija": (
        "modUtovar.bas",
        "        If kolicina > raspolozivo + 0.0001 Then\n",
        "        If False Then   ' SABOTAZA: prodaja preko stanja\n",
        "T_FakturaGP_WriterKapijeIStorno",
        "prerada bez izlaza (0 kg na stanju) se ne prodaje",
    ),
    "storno-prerada-sa-utovarom-prolazi": (
        "modStorno.bas",
        "    If UtovarenoKgPrerade(preradaID) > 0 Then\n",
        "    If False Then   ' SABOTAZA: storno ispod isporucene robe\n",
        "T_FakturaGP_WriterKapijeIStorno",
        "prerada sa aktivnim utovarom se ne stornira",
    ),
    "utovar-gp-storno-ne-vraca-stanje": (
        "modStorno.bas",
        "                MarkRowStornirano TBL_UTOVAR_STAVKE, r, SRC\n",
        "                ' SABOTAZA: stavke ostaju aktivne\n",
        "T_FakturaGP_WriterKapijeIStorno",
        "stavke storniranog utovara su stornirane",
    ),
    "storno-gp-rollback-bez-utovara": (
        "modStorno.bas",
        "    If Not GetTable(TBL_UTOVAR) Is Nothing Then\n"
        "        tx.AddTableSnapshot TBL_UTOVAR\n"
        "    End If\n",
        "    If Not GetTable(TBL_UTOVAR) Is Nothing Then\n"
        "        ' SABOTAZA: snapshot utovara uklonjen\n"
        "    End If\n",
        "T_FakturaGP_WriterKapijeIStorno",
        "rollback VRACA marker utovara (bez snapshota = dvostruka prodaja)",
    ),
    "storno-gp-faktura-ne-oslobadja-utovar": (
        "modStorno.bas",
        "                        ReleaseUtovarFromFaktura utID, fakturaID\n",
        "                        ' SABOTAZA: utovar ostaje zarobljen\n",
        "T_FakturaGP_WriterKapijeIStorno",
        "storno fakture oslobadja utovar",
    ),
    # Grid: NA STANJU mora IsNumeric/CDbl putem (Val lokal mina) i
    # dupli id se prazni; vise faktura po preradi se NE krije.
    "faktura-gp-val-lokal-mina": (
        "modFaktura.bas",
        "        If IsNumeric(pd(i, cNeto)) Then naStanju = CDbl(pd(i, cNeto))\n",
        "        naStanju = CDbl(Val(CStr(nz(pd(i, cNeto), \"0\"))))   ' SABOTAZA: Val lokal mina\n",
        "T_Fak_GpListaIKorpa",
        "decimalan izlaz kg prezivi read-model (Val mina)",
    ),
    "faktura-gp-dupli-id-crta-se": (
        "modFaktura.bas",
        "        outA(n, 1) = IdIliPrazno(brojac, preID)\n",
        "        outA(n, 1) = preID   ' SABOTAZA: dupli id se crta\n",
        "T_Fak_GpListaIKorpa",
        "dupli PreradaID prazni identitet",
    ),
    "faktura-gp-vise-faktura-skriveno": (
        "modFaktura.bas",
        "            If preFakture(preID).count = 1 Then\n",
        "            If True Then   ' SABOTAZA: druga faktura progutana\n",
        "T_FakturaGP_WriterKapijeIStorno",
        "dve fakture po preradi su LEGALNE (parcijalna prodaja)",
    ),
    # Lanac/problemi: neusaglasene prodajne veze i stanja.
    "sledljivost-gp-kontradikcija-cuti": (
        "modIzvestaj.bas",
        "                                If CBool(utV(2)) Then gpLose = True\n",
        "                                ' SABOTAZA: lose veza cuti\n",
        "T_Sled_GpLanacIStanja",
        "kontradiktorna prodajna veza obara kariku",
    ),
    "sledljivost-gp-stanje-placebo": (
        "modIzvestaj.bas",
        "                        If anySold And allSold Then\n"
        "                            stanje = SLED_ST_PRODATO_GP\n",
        "                        If anySold And allSold Then\n"
        "                            stanje = SLED_ST_PRERADJENO   ' SABOTAZA: GP prodaja nevidljiva\n",
        "T_Sled_GpLanacIStanja",
        "stanje G = prodato GP",
    ),
    "sledljivost-gp-delimicno-placebo": (
        "modIzvestaj.bas",
        "                        ElseIf anySold Then\n"
        "                            stanje = SLED_ST_DELIMICNO\n",
        "                        ElseIf anySold Then\n"
        "                            stanje = SLED_ST_PRODATO_GP   ' SABOTAZA: pola = celo\n",
        "T_Sled_GpLanacIStanja",
        "stanje P = delimicno prodato (50 od 120 kg)",
    ),
    "sledljivost-gp-prekomerno-cuti": (
        "modIzvestaj.bas",
        "                If CDbl(utVp(0)) > SledDbl(preData(i, cGNeto)) + SLED_EPS_KG Then\n",
        "                If False Then   ' SABOTAZA: prodaja robe koje nema cuti\n",
        "T_Sled_GpLanacIStanja",
        "utovareno preko proizvedenog je problem",
    ),
    "sledljivost-gp-lose-veze-cute": (
        "modIzvestaj.bas",
        "                If CBool(utVp(2)) Then\n",
        "                If False Then   ' SABOTAZA: neusaglasene veze cute\n",
        "T_Sled_GpLanacIStanja",
        "utovar-faktura bez FST stavke = neusaglasena prerada",
    ),
    "sledljivost-gp-siroce-cuti": (
        "modIzvestaj.bas",
        "    For Each sk In fstSiroce.keys\n",
        "    For Each sk In fstSiroce.keys: Exit For   ' SABOTAZA: siroce cuti\n",
        "T_Sled_GpLanacIStanja",
        "prodajna stavka bez utovara (siroce) je problem",
    ),
    # Krug 5b: UI vidljivost utovara -- lista bez broja fakture krije
    # da je isporuka vec naplacena (operater bi je fakturisao ponovo).
    "utovar-gp-lista-bez-fakture": (
        "modUtovar.bas",
        "        If fakBroj.Exists(fid) Then\n"
        "            outA(n, 7) = CStr(fakBroj(fid))\n",
        "        If False Then   ' SABOTAZA: faktura utovara skrivena\n"
        "            outA(n, 7) = CStr(fakBroj(fid))\n",
        "T_Fak_GpListaIKorpa",
        "fakturisan utovar nosi broj fakture",
    ),
    # Krug 5d: stampana utovarna lista bez LOTA (broja prerade) je
    # dokument bez sledljivosti -- roba u kamionu ne moze da se
    # upari sa evidencijom.
    "utovar-gp-stampa-bez-lota": (
        "modPrint.bas",
        "        startCell.Offset(i - 1, 1).value = stavke(i, 1)\n"
        "        startCell.Offset(i - 1, 2).value = stavke(i, 2)\n"
        "        startCell.Offset(i - 1, 3).value = stavke(i, 3)\n"
        "        startCell.Offset(i - 1, 4).value = stavke(i, 4)\n"
        "        startCell.Offset(i - 1, 5).value = stavke(i, 5)\n"
        "        startCell.Offset(i - 1, 6).value = stavke(i, 6)\n",
        "        ' SABOTAZA: lot progutan\n"
        "        startCell.Offset(i - 1, 2).value = stavke(i, 2)\n"
        "        startCell.Offset(i - 1, 3).value = stavke(i, 3)\n"
        "        startCell.Offset(i - 1, 4).value = stavke(i, 4)\n"
        "        startCell.Offset(i - 1, 5).value = stavke(i, 5)\n"
        "        startCell.Offset(i - 1, 6).value = stavke(i, 6)\n",
        "T_FakturaGP_WriterKapijeIStorno",
        "stavka liste nosi broj prerade",
    ),
    # Krug 5d: prazno prevoz polje NE sme da obrise postojecu
    # vrednost (operater dopunjava samo plombu, ostalo ostaje).
    "utovar-gp-prevoz-prazno-brise": (
        "modUtovar.bas",
        "    v = Trim$(vrednost)\n"
        "    If Len(v) = 0 Then Exit Sub\n",
        "    v = Trim$(vrednost)\n"
        "    ' SABOTAZA: prazno pregazi vrednost\n",
        "T_FakturaGP_WriterKapijeIStorno",
        "prazno polje ne dira postojecu vrednost",
    ),
    # Krug 5d: rok trajanja na obrascu = proizvodnja + N meseci iz
    # Podesavanja -- bez toga papir tvrdi pogresan rok.
    "utovar-gp-rok-placebo": (
        "modUtovar.bas",
        "                        stavke(nSt, 4) = RokIstekaZaTip(CStr(pv(0)), CDate(pv(2)))\n",
        "                        stavke(nSt, 4) = CDate(pv(2))   ' SABOTAZA: rok = proizvodnja\n",
        "T_FakturaGP_WriterKapijeIStorno",
        "rok trajanja = proizvodnja + 24 meseca",
    ),
    # Revizija #7 B2: nefakturisan utovar sa aktivnom FST je
    # kontradikcija -- bez brojanja bi re-fakturisanje duplo prodalo.
    "utovar-gp-refakt-aktivna-fst": (
        "modUtovar.bas",
        "                AktivnihFstZaUtovar = AktivnihFstZaUtovar + 1\n",
        "                AktivnihFstZaUtovar = AktivnihFstZaUtovar + 0   ' SABOTAZA\n",
        "T_FakturaGP_WriterKapijeIStorno",
        "nefakturisan utovar sa aktivnom FST se ne fakturise ponovo",
    ),
    # Revizija #7 B3: SEF kolicinski dokaz 1:1 -- bez poredjenja bi
    # poreska faktura od 400 kg prosla nad utovarom od 500 kg.
    "sef-gp-kolicina-placebo": (
        "modSEFMapper.bas",
        "        If Abs(CDbl(gpKg(CStr(k))) - CDbl(utKg(CStr(k)))) > 0.0001 Then\n",
        "        If False Then   ' SABOTAZA: kolicina se ne poredi\n",
        "T_FakturaGP_WriterKapijeIStorno",
        "SEF blokira GP fakturu cija kolicina ne odgovara utovaru",
    ),
    # Revizija #7 B1: datum utovara je poreski podatak -- lock posle
    # SEF slanja; bez njega se lokalni datum razilazi od poslatog.
    "utovar-gp-datum-lock-placebo": (
        "modUtovar.bas",
        "            If Len(wfState) > 0 _\n"
        "               And wfState <> WF_LOCAL_FINALIZED _\n"
        "               And wfState <> WF_SEF_READY _\n"
        "               And wfState <> WF_SEF_TECH_FAILED Then\n",
        "            If False Then   ' SABOTAZA: lock ugasen\n",
        "T_FakturaGP_WriterKapijeIStorno",
        "datum utovara je zakljucan posle SEF slanja",
    ),
    # Jedna faktura nosi JEDNU vrstu robe -- GP u PRJ korpu bi writer
    # kasnije odbio, ali tek posle potvrde operatera (kasna greska).
    "fakture-gp-korpa-mesa-u-prj": (
        "modScrFakture.bas",
        "    If KorpaTip() = \"PRJ\" Then\n"
        "        FkDodajGP = Poruka(\"OTKUI_ERR_FK_MESANJE\")\n",
        "    If False Then   ' SABOTAZA: GP ulazi u PRJ korpu\n"
        "        FkDodajGP = Poruka(\"OTKUI_ERR_FK_MESANJE\")\n",
        "T_Fak_GpListaIKorpa",
        "GP stavka ne ulazi u PRJ korpu",
    ),
    "fakture-gp-korpa-mesa-u-gp": (
        "modScrFakture.bas",
        "    If KorpaTip() = \"GP\" Then\n"
        "        FkDodaj = Poruka(\"OTKUI_ERR_FK_MESANJE\")\n",
        "    If False Then   ' SABOTAZA: PRJ ulazi u GP korpu\n"
        "        FkDodaj = Poruka(\"OTKUI_ERR_FK_MESANJE\")\n",
        "T_Fak_GpListaIKorpa",
        "prijemnica ne ulazi u GP korpu",
    ),
    # R1: GP faktura bez print grane stampa prazna prijemnicka polja
    # umesto proizvoda i broja prerade.
    "faktura-gp-print-prazna-polja": (
        "modFaktura.bas",
        "            If Len(preID) > 0 Then\n"
        "                ' GP: dokument = broj prerade, proizvod = TipGotovogProizvoda.\n",
        "            If False Then   ' SABOTAZA: GP stampa kao sveza\n"
        "                ' GP: dokument = broj prerade, proizvod = TipGotovogProizvoda.\n",
        "T_FakturaGP_WriterKapijeIStorno",
        "GP heder: dokument je prerada",
    ),
    # R2: SEF naziv GP linije mora da nosi broj prerade (lot) -- bez
    # njega UBL gubi vezu na izvor.
    "sef-gp-naziv-bez-prerade": (
        "modSEFMapper.bas",
        "                opis = opis & \" po preradi \" & brojPrerade\n",
        "                ' SABOTAZA: broj prerade progutan\n",
        "T_FakturaGP_WriterKapijeIStorno",
        "SEF naziv nosi broj prerade",
    ),
    # B1: zavrsna GP faktura mora da zauzme kolone Faktura/Kupac --
    # inace red uz prodatu robu pokazuje odrediste prijema.
    "sledljivost-gp-faktura-ne-preuzima": (
        "modIzvestaj.bas",
        "                        ' se NE pogadja: \"N fakt.\"/\"N kup.\" + refs.\n"
        "                        If gpFakD.count > 0 Then\n",
        "                        ' se NE pogadja: \"N fakt.\"/\"N kup.\" + refs.\n"
        "                        If False Then   ' SABOTAZA: GP faktura ne preuzima kolone\n",
        "T_Sled_GpLanacIStanja",
        "zavrsna GP faktura zauzima kolonu Faktura",
    ),
    # Krug 8 R3: sablon bez vlasnicke kapije -- dvosmislen broj bi mesao
    # tudje generacije u jedan dokument sledljivosti.
    "sledljivost-sablon-dvosmislen-broj": (
        "modIzvestaj.bas",
        "    If SledVlasnikaBroja(brojZbirne) > 1 Then\n"
        "        StampajSledljivostZbirne = \"DVOSMISLEN\"\n",
        "    If False Then   ' SABOTAZA: dvosmislen broj prolazi na sablon\n"
        "        StampajSledljivostZbirne = \"DVOSMISLEN\"\n",
        "T_Sled_MeteSledljivosti",
        "sablon odbija dvosmislen broj zbirne",
    ),
    # Krug 8 R4: nevalidan datum tiho sakriven iz ponude -- ugovor kaze
    # da anomalija ostaje VIDLJIVA.
    "sledljivost-dokumenti-nevalidan-datum-skriven": (
        "modIzvestaj.bas",
        "    If Not IsDate(v) Then\n"
        "        SledDatumUPeriodu = True\n",
        "    If Not IsDate(v) Then\n"
        "        SledDatumUPeriodu = False   ' SABOTAZA: nevalidan datum se krije\n",
        "T_Sled_DokumentiPonuda",
        "dokument sa nevalidnim datumom ostaje vidljiv",
    ),
    # Krug 8 R5 (ljuska, v6-ui-188): tekst panela nazad u label PUNE
    # visine -- GDI baseline faza opet varira po redu ("8. red veceg
    # fonta").
    "ljuska-popup-tekst-pun-red": (
        "modOtkupUI.bas",
        "        NewLbl z, \"popT\" & i, \"\", 1, _\n"
        "               CenterY(1 + i * POP_ITEM_H, POP_ITEM_H, TS_BODY), _\n"
        "               178, TxtH(TS_BODY), TS_BODY, False, C_FOREST, -1\n",
        "        NewLbl z, \"popT\" & i, \"\", 1, _\n"
        "               CenterY(1 + i * POP_ITEM_H, POP_ITEM_H, TS_BODY), _\n"
        "               178, POP_ITEM_H, TS_BODY, False, C_FOREST, -1   ' SABOTAZA: pun red\n",
        "T_Ljuska_PopupTekstTraka",
        "tekst labeli su nizi od reda i istog fonta",
    ),
    # PDF karika sa dvotackom (krug 5 S11) -- deljeni detalj-kljucevi
    # ("Zbirna:") bi u PDF koloni stajali pored karika bez dvotacke.
    "sledljivost-pdf-dvotacka": (
        "modScrSledljivost.bas",
        "Private Function BezDvotacke(ByVal s As String) As String\n"
        "    BezDvotacke = Trim$(s)\n"
        "    If Right$(BezDvotacke, 1) = \":\" Then _\n"
        "        BezDvotacke = Trim$(Left$(BezDvotacke, Len(BezDvotacke) - 1))\n"
        "End Function\n",
        "Private Function BezDvotacke(ByVal s As String) As String\n"
        "    BezDvotacke = Trim$(s)   ' SABOTAZA: dvotacka ostaje\n"
        "End Function\n",
        "T_Sled_IdentitetURedu_NeCrtaSe",
        "karika u PDF-u je bez dvotacke",
    ),
    # NEP pretraga bez lanac-brojeva (krug 4 S8) -- broj zbirne vise ne
    # nalazi njene nefakturisane prijemnice, a ekran bas to obecava
    # ("pretraga nalazi svaki broj u lancu").
    "sledljivost-nep-pretraga-bez-lanca": (
        "modScrSledljivost.bas",
        "            HaystackReda = NzS(src(i, 3)) & \"|\" & NzS(src(i, 4)) & \"|\" & _\n"
        "                           NzS(src(i, 6)) & \"|\" & SlProblemNaziv(NzS(src(i, 1))) & _\n"
        "                           \"|\" & NzS(src(i, 9))\n",
        "            HaystackReda = NzS(src(i, 3)) & \"|\" & NzS(src(i, 4)) & \"|\" & _\n"
        "                           NzS(src(i, 6)) & \"|\" & SlProblemNaziv(NzS(src(i, 1)))   ' SABOTAZA: bez lanac-brojeva\n",
        "T_Sled_KesPretragaIHint",
        "pretraga po broju zbirne nalazi neispravnu vezu fakture",
    ),
    # Ponuda polja izbora nudi preradjenu paletu kao "svezu robu".
    "sledljivost-dokumenti-preradjena-kao-sveza": (
        "modIzvestaj.bas",
        "            If UCase$(Trim$(SledTxt(palData(i, cPalPre)))) <> \"DA\" Then\n"
        "                If SledDatumUPeriodu(palData(i, cPalDat), datumOd, datumDo) Then\n",
        "            If True Then   ' SABOTAZA: preradjena u ponudi kao sveza\n"
        "                If SledDatumUPeriodu(palData(i, cPalDat), datumOd, datumDo) Then\n",
        "T_Sled_DokumentiPonuda",
        "preradjena paleta nije u ponudi kao sveza",
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


def _pogodaka(tekst: str, blok: str) -> int:
    """Broj pogodaka VEZANIH ZA POCETAK REDA (zamka 2).

    Jedino pravilo po kome se blok trazi u izvoru, i dele ga primena
    (`_zameni`, dakle i `--vrati`) i staticka provera (`_nalazi`).

    Izdvojeno posto su se ta dva vec razisla: provera je proglasavala izvor
    "ZATECEN SABOTIRAN" cim se zamena negde pojavi, a `--vrati` trazi TACNO
    jedan pogodak -- pa je savet vodio u komandu koja nema sta da uradi.
    """
    return tekst.count("\n" + blok)


def _zameni(path: str, staro: str, novo: str) -> tuple[bool, int]:
    """Zameni sidro vezano za pocetak reda. Vraca (uspeh, broj pogodaka)."""
    tekst, nl = _procitaj(path)
    pogodaka = _pogodaka(tekst, staro)
    if pogodaka != 1:
        return False, pogodaka
    _upisi(path, tekst.replace("\n" + staro, "\n" + novo), nl)
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


# --- staticka provera kataloga --------------------------------------------
#
# Zamka 9 kaze: posle izmene koda pusti CEO dvosmerni dokaz i tvrdi da je broj
# crvenih jednak broju sabotaza. Nad 220 sabotaza to traje oko dva i po sata, pa
# se u praksi vrteo podskup -- i zastarela sidra su prolazila neprimeceno. Kad je
# dokaz prvi put pusten ceo (PR #226), deset sabotaza se vise nije moglo ni
# primeniti: kod ispod njih je odavno popravljen, a sa njim je nestao i dokaz.
#
# Ovo je jeftina polovina istog pravila: sve sto se o katalogu moze utvrditi BEZ
# Excela. Traje sekundu i vrti se uz vba_check, pa sidro ne moze da zastari
# neprimeceno ni izmedju dva puna dokaza.
#
# Sta se NE proverava ovde: da sabotaza stvarno obara svoju tvrdnju. To zna samo
# pun dokaz -- v. tools/dokaz.py.

_TEST_SUB = re.compile(
    r"^(?:Public |Private )?Sub (T\w+)\(\s*\)\s*$", re.M)

# Komentar posle line-continuation `_` je syntax error (zamka 4): sabotaza tada
# ne obara test nego COMPILE, run visi do timeout-a, a izlaz je "Exception
# occurred" umesto imena tvrdnje.
_KOMENTAR_POSLE_PODVLAKE = re.compile(r"\s_\s+'")


# Nalazi koji su PRIZNATI, zapisani i imaju svog vlasnika, a ne mogu se zatvoriti
# bez izmene testa. Ispisuju se kao UPOZORENJE i ne obaraju gejt.
#
# Zasto uopste postoji spisak: crvena provera koju svi nauce da preskoce ne cuva
# nista -- a upravo tako je i nastao ovaj dug (dvosmerni dokaz se pustao nad
# podskupom, pa se "36 od 39" godinu dana citalo kao zeleno). Ime u spisku je
# obaveza, ne izuzetak: brise se cim nalaz nestane, a provera odmah javi ako ga
# neko obrise a nalaz je jos tu.
#
# Vrednost je POCETAK poruke (ili vise njih), ne ime -- nov, drugaciji nalaz nad
# istom sabotazom i dalje obara gejt. Spisak citaju i --proveri-sidra i
# tools/dokaz.py, jer je isti pojam: nalaz koji je priznat, zapisan i ima vlasnika.
POZNATI_NALAZI = {
    "stale-parent-po-broju":
        "deli tvrdnju",   # razdvajanje trazi novu tvrdnju u
                          # T_ZatecenContext_NePrevezujeTudjePrijemnice: obe
                          # sabotaze proizvedu isti vidljiv ishod, pa ih test
                          # bez seam-a nad kapijom ne moze razlikovati.

    # Ista klasa, nadjena zetvom tvrdnji: OBE sabotaze obore BAS istu poruku
    # ("isti broj zbirne kod dva vozaca daje DVA ciljna dokumenta"), pa test
    # ne moze da kaze koja je od njih pala. Razdvajanje trazi novu tvrdnju u
    # T_Oporavak_CiljneListe -- jedna nad ciljem (broj + vlasnik), druga nad
    # brojem redova. Nije deo ovog posla: ovde se popravlja KATALOG, ne testovi.
    "zbirna-vlasnik-samo-kupac":
        "deli tvrdnju sa 'oporavak-cilj-po-broju'",

    # Mrtva sabotaza, ne zastareo tekst: PostaviRez pise pa CITA NAZAD do tri
    # puta, a sabotaza svodi na jedan upis -- sto se u testu ne vidi, jer tamo
    # prvi upis uvek uspe. Invarijanta je otporna na FLAKY upis, pa je merljiva
    # samo nad laznom kontrolom koja prvi upis odbija. Zato ovde nema "tacnog"
    # teksta koji bi se upisao -- sabotazu treba ili opremiti takvim testom ili
    # obrisati uz obrazlozenje.
    "ljuska-rez-bez-potvrde":
        "tvrdnja ZASTARELA -- 'T_ZonaAgro_PrekidacRezimaZadrzavaBoju'",
}


# Isto, ali za PUN DOKAZ (tools/dokaz.py). Dva recnika, jer svaki alat vidi svoje
# nalaze: staticka provera ne moze da zna koja je tvrdnja pala, a dokaz ne moze da
# zna da li je sidro dvosmisleno. Jedan zajednicki bi svakom alatu prijavljivao
# tudje upise kao mrtve.
#
# PREFIKS MORA DA IMENUJE BAS TAJ PAD, ne njegovu vrstu. Goli "PALA DRUGA TVRDNJA"
# je cela KATEGORIJA greske: svaka buduca, sasvim druga tvrdnja u istom testu bila
# bi tiho progutana kao poznata, a dokaz bi zavrsio zeleno. Zato prefiks nosi i ime
# tvrdnje koja stvarno pada -- prepisano iz izmerenog izlaza, ne formulisano.
POZNATI_NALAZI_DOKAZ = {
    # Ista mrtva sabotaza koju vidi i staticka provera (v. POZNATI_NALAZI):
    # tamo kao zastareo tekst, ovde kao pad koji ne obara nista. Dva alata, dva
    # lica istog nalaza.
    "ljuska-rez-bez-potvrde": "NE OBARA NISTA",

    # Obara PREDUSLOV ("sa identitetom se recovery zapis pravi"): gasi celu
    # identitetsku granu, pa zapis ne nastane i ciljana tvrdnja ne dodje na
    # red. Uza varijanta (razresi po broju umesto po generaciji) ne obara
    # NISTA -- mereno: u fixture-u je izabran dokument bas prvi aktivan tog
    # broja, pa LookupActiveID vrati isti PK. Razdvajanje trazi ili drugi
    # fixture red ili novu tvrdnju u T_F8_IzabranRedOstajeIzabran.
    "f8-identitet-po-broju":
        "PALA DRUGA TVRDNJA: sa identitetom se recovery zapis pravi",

}


def _imena_testova() -> set:
    imena = set()
    for f in ("modTest.bas", "modTestBanka.bas"):
        put = os.path.join(SRC_VBA, f)
        if not os.path.exists(put):
            continue
        tekst, _ = _procitaj(put)
        imena.update(_TEST_SUB.findall(tekst))
    return imena


# Tvrdnja iz kataloga mora da bude tvrdnja BAS TOG testa.
#
# dokaz.py trazi da se deklarisan tekst nadje u poruci koja je pala; ako se ne
# nadje, javlja "PALA DRUGA TVRDNJA". Kad se tekst tvrdnje u testu promeni -- a
# menja se pri svakoj doradi -- katalog zastari TIHO, i dokaz.py pocne da laze u
# oba smera: javlja gresku nad sabotazom koja radi savrseno, a sabotazu koja
# stvarno obara tudju tvrdnju niko vise ne cita, jer je alat naucio da laje.
#
# Nadjeno merenjem: 119 od 251 unosa je nosilo zastareo tekst, pa dokaz.py nije
# mogao da vrati DOKAZANO ni nad jednim sirim prefiksom. Isto truljenje dokaza
# zbog kog je nastao --proveri-sidra, samo u drugom polju istog unosa.
_STR_SPOJ = re.compile(r'"\s*&\s*_?\s*\n?\s*"')
_NASTAVAK = re.compile(r"_\s*\n\s*")


def _podeli_amp(izraz: str) -> list:
    """Podeli izraz po '&' koji nisu u zagradama ni u navodnicima."""
    delovi, dubina, u_str, tek = [], 0, False, ""
    i, n = 0, len(izraz)
    while i < n:
        c = izraz[i]
        if c == '"':
            if u_str and i + 1 < n and izraz[i + 1] == '"':
                tek += '""'
                i += 2
                continue
            u_str = not u_str
        elif not u_str and c in "([":
            dubina += 1
        elif not u_str and c in ")]":
            dubina -= 1
        elif not u_str and c == "&" and dubina == 0:
            delovi.append(tek.strip())
            tek = ""
            i += 1
            continue
        tek += c
        i += 1
    if tek.strip():
        delovi.append(tek.strip())
    return delovi


def _kao_literal(op: str):
    """Tekst -- ako je operand CEO string literal. Inace None.

    "0.00" u Format$(x, "0.00") i "|" u Split(x, "|") jesu literali, ali NISU
    deo ispisane poruke: oni su argumenti ugnjezdenog poziva. Zato se ne gleda
    "ima li literala unutra" nego "da li je ceo operand jedan literal".
    """
    op = op.strip()
    if len(op) < 2 or not op.startswith('"'):
        return None
    buf, i, n = [], 1, len(op)
    while i < n:
        c = op[i]
        if c == '"':
            if i + 1 < n and op[i + 1] == '"':      # "" = navodnik u tekstu
                buf.append('"')
                i += 2
                continue
            return "".join(buf) if i == n - 1 else None
        buf.append(c)
        i += 1
    return None


def _poruka_delovi(izraz: str):
    """(pun tekst poruke ili None, staticki fragmenti).

    Operand koji nije sam literal je RUPA -- vrednost poznata tek u radu. Time
    "Storno / " & tip & " cita svoju tabelu" daje fragmente
    ["Storno / ", " cita svoju tabelu"], a "lista " & Split(x, "|")(0) & " ..."
    NE uvlaci "|" medju njih.
    """
    frag, rupa = [], False
    for op in _podeli_amp(izraz):
        t = _kao_literal(op)
        if t is None:
            rupa = True
        else:
            frag.append(t)
    if not frag:
        return None, []
    if rupa:
        return None, frag
    return "".join(frag), frag


# Poruka tvrdnje je POSLEDNJI argument assertion poziva. Cetiri primitive
# pokrivaju oba harness-a:
#
#   AssertEq actual, expected, poruka        (modTest, 1176 poziva)
#   ChkEq    act, exp, nm                    (modTestBanka, 88)
#   ChkEqD   act, exp, nm                    (modTestBanka, 26)
#   Chk      cond, nm                        (modTestBanka, 82)
#
# Zasto ne "bilo koji literal u testu": literal moze da bude i OCEKIVANA VREDNOST
# ili obicna dodela --
#
#     status = "blok drugog kooperanta se odbija"
#     AssertEq rezultat, "Placeno", "status fakture je ispravan"
#
# -- a dokaz.py vidi samo PORUKU koja je pala.
_ASSERT_IMENA = ("asserteq", "chkeqd", "chkeq", "chk")


def _podeli_vrh(tekst: str) -> list:
    """Podeli po zarezima koji NISU u zagradama ni u navodnicima."""
    delovi, dubina, u_str, tek = [], 0, False, ""
    for c in tekst:
        if c == '"':
            u_str = not u_str
        elif not u_str and c in "([":
            dubina += 1
        elif not u_str and c in ")]":
            dubina -= 1
        elif not u_str and c == "," and dubina == 0:
            delovi.append(tek.strip())
            tek = ""
            continue
        tek += c
    if tek.strip():
        delovi.append(tek.strip())
    return delovi


def _unutar_zagrada(arg: str):
    """Sadrzaj zagrada -- samo ako se PRVA '(' zatvara bas na kraju."""
    if not arg.startswith("("):
        return None
    dubina, u_str = 0, False
    for i, c in enumerate(arg):
        if c == '"':
            u_str = not u_str
        elif not u_str and c == "(":
            dubina += 1
        elif not u_str and c == ")":
            dubina -= 1
            if dubina == 0:
                return arg[1:i] if i == len(arg) - 1 else None
    return None


def _poruke_tvrdnji(telo: str) -> list:
    """Izrazi koji su PORUKA assertion poziva, po jedan po pozivu."""
    out = []
    for red in _NASTAVAK.sub(" ", telo).split("\n"):
        r = red.strip()
        # `Call AssertEq(a, b, "x")` -- jedini oblik u kome su spoljne zagrade
        # deo poziva. Bez `Call` je ovo VBA Sub-call: `AssertEq a, b, "x"`, i
        # tada zagrade na krajevima pripadaju ARGUMENTIMA, ne pozivu.
        #
        # Ranija verzija ih je skidala cim ostatak pocinje '(' i zavrsava ')',
        # sto nad stvarnim
        #     AssertEq (X(...) Is Nothing), True, "..." & CStr(i)
        # razbija dubinu, pa se argumenti vise ne razdvajaju i ceo poziv prodje
        # kao "poruka" -- literal iz PRVOG argumenta postane lazna tvrdnja.
        sa_zagradama = False
        if r.lower().startswith("call "):
            r = r[5:].strip()
            sa_zagradama = True
        low = r.lower()
        for ime in _ASSERT_IMENA:
            if not low.startswith(ime):
                continue
            ostatak = r[len(ime):]
            # granica imena: Chk ne sme da pojede ChkEq, ni AssertEq AssertEqX
            if ostatak[:1].isalnum() or ostatak[:1] == "_":
                continue
            arg = ostatak.strip()
            if sa_zagradama:
                arg = _unutar_zagrada(arg)
                if arg is None:
                    break
            delovi = _podeli_vrh(arg)
            if delovi:
                out.append(delovi[-1])
            break
    return out


def _telo_podaci(telo: str):
    """(pune poruke u malim slovima, njihovi staticki fragmenti).

    Izdvojeno da bi self-test isao KROZ ovu funkciju, a ne pored nje: dok je
    self-test sam sklapao svoje "telo", dokazivao je samo da _tvrdnja_pripada
    pretrazuje ono sto DOBIJE -- a rupa je bila u tome STA dobija. Sabotaza ove
    funkcije sada obara sve slucajeve koji to mere.
    """
    literali, sabloni = [], []
    for izraz in _poruke_tvrdnji(telo):
        pun, frag = _poruka_delovi(izraz)
        if pun is not None:
            literali.append(pun.lower())
        if frag:
            sabloni.append(frag)
            # I POJEDINACAN FRAGMENT je deo ispisane poruke, pa katalog sme da
            # nosi njegov deo: "kapija zaustavlja nepostojeci dokument" je
            # prefiks fragmenta "...dokument, tip " & CStr(tip). dokaz.py to
            # nalazi kao podniz poruke; bez ovoga bi staticka provera lazno
            # prijavila zastarelost.
            literali.extend(f.lower() for f in frag)
    return literali, sabloni


def _tela_testova() -> dict:
    """ime testa -> (string-literali u malim slovima, sabloni tvrdnji).

    CUVAJU SE LITERALI, NE CELO TELO. Telo sadrzi i kod i komentare, pa bi
    "AssertEq nosiDok, True" ili recenica iz komentara prosli kao tvrdnja -- a
    dokaz.py poredi sa PORUKOM koja je pala, koja moze da bude samo literal.
    Provera bi tako davala jace obecanje nego sto meri: zeleno ovde, PALA DRUGA
    TVRDNJA u prolazu.
    """
    tela = {}
    for f in ("modTest.bas", "modTestBanka.bas"):
        put = os.path.join(SRC_VBA, f)
        if not os.path.exists(put):
            continue
        tekst = _procitaj(put)[0].replace("\r\n", "\n")
        for m in re.finditer(r"^(?:Public |Private )?(?:Sub|Function) (\w+)",
                             tekst, re.M):
            k = re.search(r"^End (?:Sub|Function)\b", tekst[m.start():], re.M)
            telo = tekst[m.start(): m.start() + (k.end() if k else len(tekst))]
            tela[m.group(1)] = _telo_podaci(telo)
    return tela


def _tvrdnja_pripada(tvrdnja: str, podaci) -> bool:
    """Da li je tvrdnja tvrdnja BAS ovog testa.

    Dva oblika prolaze:
      1) doslovno je u nekom STRING-LITERALU tog testa;
      2) tvrdnja je sklopljena U RADU ("Storno / " & tip & " cita svoju tabelu"):
         literali tog izraza se u njoj nalaze REDOM, a ono sto ostane izmedju
         njih mora da lici na VREDNOST -- najvise tri reci po rupi.

    Rupa se meri brojem reci, ne procentom teksta: procenat je propustao kratke
    tvrdnje ("... u dijalogu ide BEZ oznake" je 51% svoje tvrdnje) a primao bi
    slucajno poklapanje kratkog literala u dugackoj tudjoj tvrdnji.

    Pun tekst se cuva namerno -- skracivanje na zajednicki literal bi dve
    razlicite tvrdnje spojilo u jednu (zamka 5).
    """
    literali, sabloni = podaci
    t = tvrdnja.lower()
    # Podniz je dovoljan, jer dokaz.py isto radi proveru podniza: katalog sme da
    # nosi prepoznatljiv deo duge tvrdnje. Ali samo unutar LITERALA.
    if any(t in l for l in literali):
        return True
    for izraz in sabloni:
        poz, pokriveno, rupe_ok = 0, 0, True
        for f in izraz:
            k = t.find(f.lower(), poz)
            if k < 0:
                rupe_ok = False
                break
            if len(t[poz:k].split()) > 3:      # rupa mora da lici na vrednost
                rupe_ok = False
                break
            poz = k + len(f)
            pokriveno += len(f)
        if rupe_ok and len(t[poz:].split()) > 3:
            rupe_ok = False
        if rupe_ok and pokriveno >= 8:
            return True
    return False


# ZAMENA KOJA DODELJUJE IMENU TUDJE PROCEDURE NE KOMPAJLIRA.
#
# `ObradiDogadjaj` sme da dodeli sebi; `Scr_Event = ...` iz njenog tela je dodela
# imenu druge procedure -- compile error. Posledica nije pao test nego Excel u
# [break]: suite se ne pokrene, dokaz.py ne vidi nijednu palu tvrdnju i prijavi
# "NE OBARA NISTA". Sabotaza koja ne kompajlira i sabotaza koja nista ne meri
# izgledaju IDENTICNO, a razlika je velika -- prva se popravlja u jednom redu,
# druga trazi rad nad testom.
#
# Mereno na paleta-klik-otvara: 68 s po prolazu, Excel ostaje otvoren u break-u.
_ZAMENA_DODELA = re.compile(r"^\s*(?:Set\s+)?([A-Za-z_]\w*)\s*=(?!=)")
_PROC_OTVARA = re.compile(
    r"^(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?"
    r"(Sub|Function|Property)\s+(?:(?:Get|Let|Set)\s+)?(\w+)", re.IGNORECASE)
_PROC_ZATVARA = re.compile(r"^End\s+(?:Sub|Function|Property)\b", re.IGNORECASE)


def _procedure_po_redu(tekst: str) -> list:
    """[(ime, vrsta, prvi_red, poslednji_red)] -- vrsta je sub/function/property."""
    out, ime, vrsta, poc = [], None, None, 0
    for i, red in enumerate(tekst.split("\n")):
        t = red.strip()
        m = _PROC_OTVARA.match(t)
        if m:
            vrsta, ime, poc = m.group(1).lower(), m.group(2), i
        elif ime and _PROC_ZATVARA.match(t):
            out.append((ime, vrsta, poc, i))
            ime, vrsta = None, None
    return out


# `Const` NIJE u ovom skupu: lokalni `Const Foo` jeste zaklonio proceduru, ali
# `Foo = 2` je dodela konstanti -- i dalje compile error. Skup opisuje mesta na
# koja se SME dodeliti, ne sva lokalna imena.
_LOKAL_DODELJIV = re.compile(r"^\s*(?:Dim|Static)\s+", re.IGNORECASE)
_PARAM_UKRAS = re.compile(r"^(?:ByVal|ByRef|Optional|ParamArray)\s+", re.IGNORECASE)


def _po_zarezu_vrh(tekst: str) -> list:
    """Podeli po zarezima koji nisu u zagradama ni u navodnicima."""
    delovi, dubina, u_str, tek = [], 0, False, ""
    for c in tekst:
        if c == '"':
            u_str = not u_str
        elif not u_str and c in "([":
            dubina += 1
        elif not u_str and c in ")]":
            dubina -= 1
        elif not u_str and c == "," and dubina == 0:
            delovi.append(tek.strip())
            tek = ""
            continue
        tek += c
    if tek.strip():
        delovi.append(tek.strip())
    return delovi


def _prva_imena(tekst: str) -> list:
    """Imena iz `a As X, b(1 To 3) As Y, ByVal c As Z`."""
    out = []
    for deo in _po_zarezu_vrh(tekst):
        deo = deo.strip()
        while _PARAM_UKRAS.match(deo):
            deo = _PARAM_UKRAS.sub("", deo, count=1).strip()
        m = re.match(r"([A-Za-z_]\w*)", deo)
        if m:
            out.append(m.group(1).lower())
    return out


def _spoji_nastavke(redovi: list) -> list:
    """Fizicki redovi -> logicki, po VBA nastavku ` _`.

    Potrebno je jer su prelomljeni potpisi u ovom repou uobicajeni:

        Sub P(ByVal x As Long, _
              ByVal Foo As Boolean)

    Citanje po fizickim redovima vidi samo `x`, pa bi `Foo = True` -- dodela
    PARAMETRU -- bilo prijavljeno kao dodela istoimenoj funkciji. Isto vazi za
    prelomljen `Dim`.
    """
    out, tek = [], ""
    for red in redovi:
        t = red.rstrip()
        tek = (tek + " " + t.strip()) if tek else t
        # VBA trazi RAZMAK pa `_`. Golo endswith("_") bi identifikator koji se
        # zavrsava podvlakom progutalo kao nastavak reda.
        if re.search(r"\s_$", tek.rstrip() + ""):
            tek = tek.rstrip()[:-1].rstrip()
            continue
        out.append(tek)
        tek = ""
    if tek:
        out.append(tek)
    return out


def _lokalna_imena(redovi: list, a: int, b: int) -> set:
    """Parametri + lokalni Dim/Static unutar procedure [a..b].

    `Const` NIJE ovde: on zaklanja proceduru, ali dodela konstanti je i dalje
    compile error -- v. _LOKAL_DODELJIV.

    VBA DOZVOLJAVA da lokalno ime ZAKLONI proceduru: `Dim Foo As Boolean` u P()
    znaci da je `Foo = True` dodela promenljivoj, ne pokusaj dodele Sub-u. Repo
    to vec zna -- vba_check ima pravilo ZAKLONJENO nad istom pojavom. Bez ovoga
    bi pravilo prijavljivalo legalan VBA, a lazan nalaz u hook-u je gori od
    propustenog.
    """
    imena = set()
    segment = _spoji_nastavke(redovi[a:min(b, len(redovi) - 1) + 1])
    for i, red in enumerate(segment):
        t = red.strip()
        if i == 0 and _PROC_OTVARA.match(t):
            if "(" in t:
                zagrada = t[t.index("(") + 1:]
                k = zagrada.rfind(")")
                imena.update(_prva_imena(zagrada[:k] if k >= 0 else zagrada))
            continue
        m = _LOKAL_DODELJIV.match(t)
        if m:
            imena.update(_prva_imena(t[m.end():]))
    return imena


def _dodela_tudjoj_proceduri(tekst: str, staro: str, novo: str):
    """Ime procedure kojoj zamena dodeljuje a ne sme, ili None.

    VRSTA ODLUCUJE, i pravilo je namerno UZE nego sto ime sugerise:

      Function  dodela SVOM imenu je povratna vrednost -- dozvoljena.
                Dodela TUDJEM imenu funkcije je compile error.
      Sub       nema povratnu vrednost, pa je dodela njenom imenu greska i
                iznutra i spolja.
      Property  IZUZETA. `X = v` je tamo poziv Property Let, dakle legalan VBA;
                bez uparivanja Get/Let/Set se ne moze reci da je greska, a lazan
                nalaz u hook-u je gori od propustenog.

    Poredi se NEOSETLJIVO NA VELICINU SLOVA, jer je VBA takav: `scr_event = ...`
    je isti compile error kao `Scr_Event = ...`, a poredjenje po tacnom zapisu
    bi ga pustilo -- trivijalan zaobilazak bas ovog pravila.
    """
    idx = tekst.find("\n" + staro)
    if idx < 0:
        return None
    red = tekst[:idx + 1].count("\n")

    # SIMBOLI SE CITAJU IZ MUTIRANOG KODA, ne iz zdravog.
    #
    # Kompajlira se ono sto sabotaza NAPRAVI, pa se i pita o njemu. Racun nad
    # zdravim tekstom gresi u oba smera:
    #   - zamena koja UKLONI `Dim Foo` ostavlja dodelu tudjoj funkciji, a checker
    #     bi jos video stari Dim i pustio je (propusten compile error);
    #   - zamena koja UVEDE `Dim Foo` daje legalan VBA, a checker bi ga prijavio
    #     (lazna uzbuna nad ispravnim kodom).
    #
    # Kes nad originalnim tekstom je time otpao. 251 unos je premalo da bi
    # ustedu vredelo platiti netacnoscu.
    mutirani = tekst.replace("\n" + staro, "\n" + novo, 1)
    procs = _procedure_po_redu(mutirani)
    redovi = mutirani.split("\n")
    unutar, lokalna = None, set()
    for pime, _v, a, b in procs:
        if a <= red <= b:
            unutar = pime.lower()
            lokalna = _lokalna_imena(redovi, a, b)
            break

    # Property se NE skuplja -- time je i izuzeta. Zasebna `if n in props` provera
    # je bila mrtva grana: property ime ionako nije ni u subovi ni u funkcije.
    subovi = {p.lower() for p, v, _a, _b in procs if v == "sub"}
    funkcije = {p.lower() for p, v, _a, _b in procs if v == "function"}

    for linija in novo.split("\n"):
        m = _ZAMENA_DODELA.match(linija)
        if not m:
            continue
        n = m.group(1).lower()
        if n in lokalna:
            continue                       # lokalno ime ZAKLANJA proceduru
        if n in subovi:
            return m.group(1)              # Sub nema povratnu vrednost
        if n in funkcije and n != unutar:
            return m.group(1)              # tudja funkcija
    return None


def _nalazi(katalog: dict, imena: set, tela: dict = None) -> list:
    """Nalazi nad DATIM katalogom. Izdvojeno da bi --self-test mogao da mu
    podmetne izmisljene unose, umesto da alat prepisuje sopstveni fajl."""
    nalazi = []
    videne_tvrdnje = {}
    kes = {}                 # fajl se cita jednom, ne 222 puta (ovo ide u hook)
    if tela is None:
        tela = _tela_testova()

    for ime, (fajl, staro, novo, test, tvrdnja) in katalog.items():
        put = os.path.join(SRC_VBA, fajl)
        if not os.path.exists(put):
            nalazi.append((ime, f"nema fajla src-vba/{fajl}"))
            continue

        if fajl not in kes:
            kes[fajl] = _procitaj(put)[0]
        tekst = kes[fajl]
        pogodaka = _pogodaka(tekst, staro)
        zamene = _pogodaka(tekst, novo)
        if pogodaka != 1:
            # Sidra nema, ali je ZAMENA tu -- izvor nije popravljen nego je
            # jos SABOTIRAN. dokaz.py ciscenje radi kroz `finally`, sto ne
            # stigne kad se proces ubije spolja (taskkill, zatvoren terminal).
            # Bez ovog razdvajanja poruka glasi 'kod ispod sidra je popravljen',
            # a odgovor na nju je da se sidro uskladi sa zatecenim kodom --
            # sto sabotazu zacementira kao novu istinu.
            #
            # TACNO JEDNA zamena, ne 'bar jedna': `--vrati` ide kroz _zameni,
            # koji nad vise pogodaka odbija posao. Prva verzija ovog pravila
            # je tvrdila sabotazu na puko prisustvo, pa je savet 'pokreni
            # --vrati' vodio u komandu koja nema sta da uradi -- ista klasa
            # greske koju ovo pravilo treba da hvata.
            if pogodaka == 0 and zamene == 1:
                razlog = ("izvor je ZATECEN SABOTIRAN -- pokreni "
                          "`python tools/sabotaza.py --vrati`")
            elif pogodaka == 0 and zamene > 1:
                razlog = ("sidra nema, a zamena se nalazi %d puta -- --vrati "
                          "trazi TACNO jedan pogodak, pa ne bi vratio nista; "
                          "ne tvrdi se ni sabotaza ni popravljen kod" % zamene)
            else:
                razlog = ("sidro ZASTARELO -- kod ispod njega je popravljen"
                          if pogodaka == 0 else "sidro nije jednoznacno")
            nalazi.append((ime, f"{razlog} ({pogodaka} pogodaka u {fajl})"))
        elif zamene:
            # Sidro je zdravo, ali zamena vec postoji u izvoru. Posle sabotaze
            # bi je bilo zamene+1, pa `--vrati` (koji trazi tacno jednu) ne bi
            # umeo da je vrati -- sabotaza bi ostala u radnom stablu. Mereno
            # nad zatecenim katalogom: nijedan unos ovo ne krsi, pa pravilo
            # ne zatvara nista postojece nego drzi buduce unose.
            nalazi.append((ime, "zamena vec postoji u ZDRAVOM izvoru (%d puta) "
                                "-- posle sabotaze bi je bilo %d, a --vrati "
                                "trazi tacno jednu" % (zamene, zamene + 1)))

        if test not in imena:
            nalazi.append((ime, f"test '{test}' ne postoji u modTest/modTestBanka"))

        # zamka 7: prazna zamena se "nalazi" svuda, pa --vrati tiho ne uradi nista
        if not novo.strip():
            nalazi.append((ime, "zamena je prazna -- --vrati je nikad nece naci"))
        elif novo == staro:
            nalazi.append((ime, "zamena je jednaka sidru -- sabotaza ne menja nista"))
        # zamka 8: zamena sadrzana u sidru se nalazi i u ZDRAVOM kodu.
        #
        # Poredi se ISTIM pravilom kojim radi _zameni -- od pocetka reda. Golo
        # `novo in staro` daje lazne uzbune: zamena koja uklanja `If ... Then _`
        # jeste podniz sidra kao tekst, ali joj se uvlacenje razlikuje, pa je
        # --vrati u zdravom kodu ne nalazi (primer: kapija-i-uz-identitet).
        elif ("\n" + novo) in ("\n" + staro):
            nalazi.append((ime, "zamena je podniz sidra -- --vrati bi dirao zdrav kod"))

        # zamka 10: dodela imenu tudje procedure -- compile error, ne pad testa
        tudja = _dodela_tudjoj_proceduri(tekst, staro, novo)
        if tudja:
            nalazi.append((ime, "zamena dodeljuje imenu tudje procedure '%s' -- "
                                "to je compile error, pa Excel stane u [break] a "
                                "dokaz.py prijavi 'NE OBARA NISTA'" % tudja))

        if _KOMENTAR_POSLE_PODVLAKE.search(novo):
            nalazi.append((ime, "komentar posle line-continuation '_' -- syntax error"))

        # zamka 9: tvrdnja koja nije tvrdnja SVOG testa
        if not tvrdnja.strip():
            nalazi.append((ime, "katalog nema tvrdnju -- dokaz.py nema sta da poredi"))
        elif test in tela and not _tvrdnja_pripada(tvrdnja, tela[test]):
            drugde = sorted(t for t, p in tela.items()
                            if t != test and _tvrdnja_pripada(tvrdnja, p))
            if drugde:
                nalazi.append((ime, "tvrdnja pripada testu '%s', a deklarisan je "
                                    "'%s'" % (drugde[0], test)))
            else:
                nalazi.append((ime, "tvrdnja ZASTARELA -- '%s' nema takvu tvrdnju "
                                    "(dokaz.py bi javio PALA DRUGA TVRDNJA)" % test))

        # zamka 5: dve sabotaze koje test ne razlikuje
        kljuc = (test, tvrdnja)
        if tvrdnja and kljuc in videne_tvrdnje:
            nalazi.append((ime, "deli tvrdnju sa '%s' -- test ih ne razlikuje"
                                % videne_tvrdnje[kljuc]))
        else:
            videne_tvrdnje[kljuc] = ime

    return nalazi


def poznat_nalaz(ime: str, poruka: str, spisak: dict = None) -> bool:
    """Da li je bas ovaj nalaz priznat i zapisan (v. POZNATI_NALAZI*)."""
    p = (POZNATI_NALAZI if spisak is None else spisak).get(ime)
    if not p:
        return False
    if isinstance(p, str):
        p = (p,)
    return any(poruka.startswith(x) for x in p)


def proveri_sidra(tiho: bool = False) -> int:
    """Sve o katalogu sto se vidi bez pokretanja Excela.

    `tiho` je za hook: cist katalog ne pise nista, jer se hook vrti posle svake
    izmene i njegov izlaz se cita samo kad nesto pukne.
    """
    nalazi = _nalazi(SABOTAZE, _imena_testova())

    tvrdi, poznati, mrtvi_upisi = [], [], set(POZNATI_NALAZI)
    for ime, sta in nalazi:
        if poznat_nalaz(ime, sta):
            poznati.append((ime, sta))
            mrtvi_upisi.discard(ime)
        else:
            tvrdi.append((ime, sta))

    for ime, sta in poznati:
        if not tiho:
            print(f"KATALOG-POZNATO: {ime}: {sta}")
    for ime, sta in tvrdi:
        print(f"KATALOG: {ime}: {sta}", file=sys.stderr)

    # Upis koji vise nista ne pokriva je isto nalaz: ili je nalaz zatvoren pa
    # spisak treba skratiti, ili se sabotaza zove drugacije pa upis ne stiti nista.
    for ime in sorted(mrtvi_upisi):
        print(f"KATALOG: POZNATI_NALAZI['{ime}'] ne pokriva nijedan nalaz -- "
              f"obrisi ga ili ispravi ime", file=sys.stderr)
        tvrdi.append((ime, "mrtav upis"))

    if not tiho or tvrdi:
        print(f"provereno {len(SABOTAZE)} sabotaza, nalaza {len(tvrdi)}"
              f" (+{len(poznati)} poznatih)")
    return 1 if tvrdi else 0


# --- dokaz nad samom proverom ---------------------------------------------
#
# Provera koja nikad nije pokazana crvena ne dokazuje da ista meri -- a bas je
# takva provera i trebalo da spreci dug koji je zatekao katalog (deset mrtvih
# sidara). Zato se za SVAKO pravilo podmetne izmisljen unos i tvrdi se da nalaz
# stigne BAS po tom pravilu.
#
# Katalog je izmisljen, ali fajlovi su pravi: sidro se i dalje trazi u src-vba,
# jer se bas to poredjenje proverava.
_ST_PRAVI = "    On Error Resume Next" + ESCN
# Fixture stoji na redu koji se NE menja sa izdanjem. Prva verzija je stajala
# na OTKUI_BUILD, pa je prvo podizanje builda (v6-ui-183) prijavilo ZDRAV unos
# kao nalaz -- self-test je pao na sopstvenom fixture-u.
_ST_ZDRAVO = "Option Explicit"
# Za viseredni slucaj trebaju STVARNO susedni redovi: izmedju Option Explicit
# i palete stoji prazan red, pa je fixture pored ciljanog nalaza davao i visak
# ("sidro ZASTARELO"). Ova dva su susedna.
_ST_PAR1 = "'--- paleta ----------------------------------------------------------"
_ST_PAR2 = "' PAZI: VBA Long boja je &HBBGGRR, obrnuto od CSS-a. Ovde je stajalo &H140D1E"

_SELF_TEST = [
    ("sidro zastarelo",
     ("modOtkupUI.bas", "    OvogaRedaNemaNigdeUProjektu = 1" + ESCN,
      "    Nesto = 2" + ESCN, None, "tvrdnja A"),
     "sidro ZASTARELO"),
    # Sidra nema, ali ZAMENA jeste u fajlu -- zatecena sabotaza, ne zastarelo
    # sidro. Merena razlika: prekinut dokaz.py ostavi pokvaren src-vba, a stara
    # poruka je tvrdila da je "kod ispod sidra popravljen".
    #
    # Fixture mora imati TACNO jedan pogodak. Prva verzija je koristila
    # _ST_PRAVI (`On Error Resume Next`, 151 pogodak u modOtkupUI.bas), pa je
    # self-test bio zelen nad stanjem koje `--vrati` odbija -- provera je
    # tvrdila vise nego sto meri, bas ono protiv cega postoji.
    ("izvor zatecen sabotiran",
     ("modOtkupUI.bas", "    OvogaRedaNemaNigdeUProjektu = 1" + ESCN,
      _ST_PAR2 + ESCN, None, "tvrdnja A"),
     "ZATECEN SABOTIRAN"),
    # Druga strana istog pravila: zamene ima VISE, pa --vrati ne bi vratio
    # nista. Ne sme se tvrditi ni sabotaza ni popravljen kod.
    ("zamena visestruka -- sabotaza se NE tvrdi",
     ("modOtkupUI.bas", "    OvogaRedaNemaNigdeUProjektu = 1" + ESCN,
      _ST_PRAVI, None, "tvrdnja A"),
     "trazi TACNO jedan pogodak"),
    # Zdravo sidro, ali zamena vec stoji u izvoru: posle sabotaze bi je bilo
    # dve, pa --vrati vise ne bi umeo da je vrati.
    ("zamena vec postoji u zdravom izvoru",
     (None, None, _ST_PAR2 + ESCN, None, None),
     "zamena vec postoji u ZDRAVOM izvoru"),
    ("sidro dvosmisleno",
     ("modOtkupUI.bas", _ST_PRAVI, "    Nesto = 2" + ESCN, None, "tvrdnja B"),
     "sidro nije jednoznacno"),
    ("nema fajla",
     ("modOvogaNema.bas", _ST_PRAVI, "x" + ESCN, None, "tvrdnja C"),
     "nema fajla"),
    ("test ne postoji",
     (None, None, None, "T_OvakavTestNePostoji", "tvrdnja D"),
     "ne postoji u modTest"),
    ("zamena prazna",
     (None, None, "" , None, "tvrdnja E"),
     "zamena je prazna"),
    ("zamena ista kao sidro",
     (None, None, "SIDRO", None, "tvrdnja F"),
     "zamena je jednaka sidru"),
    # Sidro self-testa je Option Explicit -- deklaraciona sekcija, dakle NIJEDNA
    # procedura. Dodela bilo kom imenu procedure tog fajla je tu compile error.
    ("dodela imenu tudje procedure",
     (None, None, "    ParcelaID = 1" + ESCN, None, "tvrdnja H"),
     "dodeljuje imenu tudje procedure"),
    ("komentar posle podvlake",
     (None, None, "    a = Sastavi(b, _   ' SABOTAZA" + ESCN, None, "tvrdnja G"),
     "komentar posle line-continuation"),
    # zamka 9: tekst tvrdnje se menja pri doradi testa, a katalog zastari TIHO
    ("tvrdnja zastarela",
     (None, None, None, None, "ovakve tvrdnje nema u testu"),
     "tvrdnja ZASTARELA"),
    # ostrija varijanta: tekst POSTOJI, ali pripada drugom testu
    ("tvrdnja pripada drugom testu",
     (None, None, None, None, "tudja tvrdnja"),
     "tvrdnja pripada testu"),
    # bez tvrdnje dokaz.py nema sta da poredi -- rupa, ne izuzetak
    ("prazna tvrdnja",
     (None, None, None, None, "   "),
     "katalog nema tvrdnju"),
    # Tekst koji u telu POSTOJI, ali kao KOD -- dokaz.py ga nikad nece videti u
    # poruci, pa bi zelena staticka provera obecavala vise nego sto meri.
    ("tvrdnja postoji samo kao VBA kod",
     (None, None, None, None, "AssertEq nosiDok, True"),
     "tvrdnja ZASTARELA"),
    # Isto za komentar: recenica iz njega nikad ne stigne u poruku testa.
    ("tvrdnja postoji samo u komentaru",
     (None, None, None, None, "recenica iz komentara koja nije tvrdnja"),
     "tvrdnja ZASTARELA"),
    # Literal JESTE u testu, ali kao obicna dodela -- nikad ne stigne u poruku.
    ("tvrdnja postoji samo u string promenljivoj",
     (None, None, None, None, "tekst iz obicne promenljive"),
     "tvrdnja ZASTARELA"),
    # Literal je OCEKIVANA VREDNOST AssertEq-a, ne poruka. dokaz.py bi ga video
    # samo u "ocekivano [...]" delu, a poredi se sa porukom.
    ("tvrdnja je ocekivana vrednost, ne poruka",
     (None, None, None, None, "ocekivana vrednost"),
     "tvrdnja ZASTARELA"),
    # Literal iz PRVOG argumenta, kad poziv izgleda kao da je ceo u zagradama:
    #     AssertEq (Len("LAZNA_TVRDNJA") > 0), True, "prava poruka " & CStr(i)
    # Naivno skidanje spoljnih zagrada razbija dubinu, argumenti se ne razdvoje,
    # i ceo poziv prodje kao "poruka". Reprodukovano na pravom modTest.bas.
    ("tvrdnja je literal iz prvog argumenta u zagradi",
     (None, None, None, None, "LAZNA_TVRDNJA"),
     "tvrdnja ZASTARELA"),
    # Literal koji je argument ugnjezdene funkcije -- nikad se ne ispisuje.
    ("tvrdnja je separator unutar Split-a",
     (None, None, None, None, "|"),
     "tvrdnja ZASTARELA"),
    # Isto za format-spec: Format$(x, "0.00") oblikuje broj, ne pise "0.00".
    ("tvrdnja je format-spec unutar Format$",
     (None, None, None, None, "0.00"),
     "tvrdnja ZASTARELA"),
]


# Granica pravila o dodeli imenu procedure. Polovina su NULE, i ta polovina je
# vaznija: pravilo sme da bude usko, ali ne sme da prijavi legalan VBA.
# (naziv, ocekivano ime ili None, izvor, zamena)
_DODELA_CASES = [
    # VBA je case-insensitive: `foo = ` je isti compile error kao `Foo = `.
    ("casing ne spasava", "foo",
     "Option Explicit\nFunction Foo() As Boolean\nEnd Function\n"
     "Sub P()\n    SIDRO\nEnd Sub\n",
     "    SIDRO\n", "    foo = True\n"),
    # Dodela SVOM imenu je povratna vrednost funkcije -- legalno.
    ("Function dodeljuje svom imenu", None,
     "Option Explicit\nFunction Foo() As Boolean\n    SIDRO\nEnd Function\n",
     "    SIDRO\n", "    Foo = True\n"),
    # `Foo = 1` uz Property Let je POZIV, ne greska.
    ("Property Let se ne prijavljuje", None,
     "Option Explicit\nProperty Let Foo(ByVal v As Long)\nEnd Property\n"
     "Sub P()\n    SIDRO\nEnd Sub\n",
     "    SIDRO\n", "    Foo = 1\n"),
    # Sub nema povratnu vrednost, pa je dodela njenom imenu greska i IZNUTRA.
    ("Sub nema povratnu vrednost", "Foo",
     "Option Explicit\nSub Foo()\n    SIDRO\nEnd Sub\n",
     "    SIDRO\n", "    Foo = 1\n"),
    # VBA dozvoljava da LOKALNO ime zakloni proceduru.
    ("lokalni Dim zaklanja Sub", None,
     "Option Explicit\nSub Foo()\nEnd Sub\n"
     "Sub P()\n    Dim Foo As Boolean\n    SIDRO\nEnd Sub\n",
     "    SIDRO\n", "    Foo = True\n"),
    ("parametar zaklanja Function", None,
     "Option Explicit\nFunction Foo() As Boolean\nEnd Function\n"
     "Sub P(ByVal Foo As Boolean)\n    SIDRO\nEnd Sub\n",
     "    SIDRO\n", "    Foo = True\n"),
    # Const ZAKLANJA, ali dodela konstanti je i dalje compile error.
    ("Const nije dodeljiv", "Foo",
     "Option Explicit\nSub Foo()\nEnd Sub\n"
     "Sub P()\n    Const Foo As Long = 1\n    SIDRO\nEnd Sub\n",
     "    SIDRO\n", "    Foo = 2\n"),
    # Prelomljen potpis je u ovom repou uobicajen.
    ("multiline parametar zaklanja", None,
     "Option Explicit\nFunction Foo() As Boolean\nEnd Function\n"
     "Sub P(ByVal x As Long, _\n      ByVal Foo As Boolean)\n    SIDRO\nEnd Sub\n",
     "    SIDRO\n", "    Foo = True\n"),
    ("multiline Dim zaklanja", None,
     "Option Explicit\nSub Foo()\nEnd Sub\n"
     "Sub P()\n    Dim x As Long, _\n        Foo As Boolean\n    SIDRO\nEnd Sub\n",
     "    SIDRO\n", "    Foo = True\n"),
    # SIMBOLI SE CITAJU IZ MUTIRANOG KODA. Ova dva to i dokazuju: prvi brise
    # deklaraciju koju zdrav kod ima, drugi je uvodi tamo gde je nema.
    ("zamena UKLANJA deklaraciju", "Foo",
     "Option Explicit\nFunction Foo() As Boolean\nEnd Function\n"
     "Sub P()\n    Dim Foo As Boolean\n    SIDRO\nEnd Sub\n",
     "    Dim Foo As Boolean\n    SIDRO\n", "    Foo = True\n"),
    ("zamena UVODI deklaraciju", None,
     "Option Explicit\nFunction Foo() As Boolean\nEnd Function\n"
     "Sub P()\n    SIDRO\nEnd Sub\n",
     "    SIDRO\n", "    Dim Foo As Boolean\n    Foo = True\n"),
]

def _self_test() -> int:
    """Svako pravilo mora da pukne nad izmisljenim unosom -- i samo ono."""
    imena = {"T_Postoji"}
    # Telo se podmece kao i imena -- inace bi se pravilo o tvrdnji merilo nad
    # PRAVIM modTest-om, pa bi self-test zavisio od tudjeg fajla.
    #
    # I PARSIRA se, ne pise rucno: tako se dokazuje da _literali stvarno izbacuje
    # kod i komentare. Recenica iz komentara i komad koda ispod moraju da PADNU
    # kao zastarela tvrdnja, i za to postoje dva slucaja.
    _VBA = ("Private Sub T_Postoji()\n"
            "    ' recenica iz komentara koja nije tvrdnja\n"
            "    Dim status As String\n"
            "    status = \"tekst iz obicne promenljive\"\n"
            "    AssertEq rezultat, \"ocekivana vrednost\", \"tvrdnja uz vrednost\"\n"
            # prvi argument u zagradi + poruka koja se zavrsava pozivom:
            # spoljne zagrade NISU par, pa naivno skidanje razbija argumente
            "    AssertEq (Len(\"LAZNA_TVRDNJA\") > 0), True, _\n"
            "             \"prava poruka \" & CStr(i)\n"
            # literali unutar ugnjezdenih poziva nikad ne stignu u ispis
            "    AssertEq a, b, \"lista \" & Split(CStr(x), \"|\")(0) & \" ima tabelu\"\n"
            "    AssertEq a, b, \"iznos \" & Format$(x, \"0.00\")\n"
            "    AssertEq nosiDok, True, \"zdrava tvrdnja\"\n"
            "    AssertEq a, b, \"tvrdnja A\"\n"
            "    AssertEq a, b, \"tvrdnja B\"\n"
            "    AssertEq a, b, \"tvrdnja C\"\n"
            "    AssertEq a, b, \"tvrdnja D\"\n"
            "    AssertEq a, b, \"tvrdnja E\"\n"
            "    AssertEq a, b, \"tvrdnja F\"\n"
            "    AssertEq a, b, \"tvrdnja G\"\n"
            "End Sub\n")
    tela = {"T_Postoji": _telo_podaci(_VBA),
            "T_Drugi": (["tudja tvrdnja"], [])}
    # zdrav unos: sidro koje postoji tacno jednom, zamena koja nije podniz
    zdravo = ("modOtkupUI.bas",
              _ST_ZDRAVO + ESCN,
              "    Nesto = 1" + ESCN,
              "T_Postoji", "zdrava tvrdnja")

    lose = 0
    n = 0
    # 0) zdrav katalog ne sme da da nijedan nalaz
    if _nalazi({"zdrav": zdravo}, imena, tela):
        print("SELF-TEST: zdrav unos je prijavljen kao nalaz", file=sys.stderr)
        lose += 1
    n += 1

    for opis, polja, ocekivano in _SELF_TEST:
        unos = tuple(z if p is None else p for p, z in zip(polja, zdravo))
        if unos[2] == "SIDRO":
            unos = (unos[0], unos[1], unos[1], unos[3], unos[4])
        nalazi = _nalazi({"probni": unos}, imena, tela)
        n += 1
        if not any(ocekivano in sta for _, sta in nalazi):
            print("SELF-TEST: '%s' nije prijavljeno (%s)" % (opis, nalazi),
                  file=sys.stderr)
            lose += 1

    # PORUKA IZA LAZNIH SPOLJNIH ZAGRADA mora da se prepozna.
    #
    #     AssertEq (Len("X") > 0), True, "prava poruka " & CStr(i)
    #
    # Ostatak pocinje '(' i zavrsava ')', ali to NISU iste zagrade. Ko ih skine
    # bezuslovno, razbije dubinu, argumenti se ne razdvoje i poruka se izgubi --
    # pa ispravna tvrdnja bude prijavljena kao zastarela. Lazna uzbuna u hook-u
    # je skuplja od propusta, jer uci da se checker preskace.
    n += 1
    iza_zagrada = tuple(zdravo[:4]) + ("prava poruka ",)
    if _nalazi({"iza-zagrada": iza_zagrada}, imena, tela):
        print("SELF-TEST: poruka iza laznih spoljnih zagrada nije prepoznata",
              file=sys.stderr)
        lose += 1

    # Granica pravila o dodeli imenu procedure -- ide kroz _dodela_tudjoj_proceduri,
    # istu funkciju koju zove _nalazi. Da je unos u katalogu stvarno provuce kroz
    # nju, dokazuje zaseban slucaj "dodela imenu tudje procedure" iznad.
    for naziv, ocekivano, izvor, sidro, zamena in _DODELA_CASES:
        n += 1
        dobijeno = _dodela_tudjoj_proceduri(izvor, sidro, zamena)
        if dobijeno != ocekivano:
            print("SELF-TEST: dodela/%s: ocekivano %r, dobijeno %r"
                  % (naziv, ocekivano, dobijeno), file=sys.stderr)
            lose += 1

    # ISTO STANJE MORA DA ODBIJE I `--vrati`, ne samo provera.
    #
    # Bez ovoga se dva pravila mogu opet razici: provera bi cutala o stanju
    # koje `--vrati` ne ume da vrati. Zove se BAS _zameni, ista funkcija koju
    # zove --vrati, i to nad KOPIJOM pravog fajla -- da self-test ni u jednom
    # ishodu ne moze da upise u src-vba.
    n += 1
    izvor = os.path.join(SRC_VBA, "modOtkupUI.bas")
    tekst_kopije, nl_kopije = _procitaj(izvor)
    kopija = izvor + ".selftest.tmp"
    _upisi(kopija, tekst_kopije, nl_kopije)
    try:
        ok_vise, k_vise = _zameni(kopija, _ST_PRAVI, "    Nesto = 2" + ESCN)
        ok_nema, k_nema = _zameni(kopija,
                                  "    OvogaRedaNemaNigdeUProjektu = 1" + ESCN,
                                  "    Nesto = 2" + ESCN)
        dirnuto = _procitaj(kopija)[0] != tekst_kopije
    finally:
        os.remove(kopija)
    if ok_vise or k_vise < 2 or ok_nema or k_nema != 0 or dirnuto:
        print("SELF-TEST: _zameni prihvata stanje koje provera ne priznaje "
              "(visestruko=%r, nema=%r, upisano=%r)"
              % ((ok_vise, k_vise), (ok_nema, k_nema), dirnuto), file=sys.stderr)
        lose += 1

    # deljena tvrdnja: dva unosa sa istim (test, tvrdnja)
    par = {"prvi": zdravo, "drugi": zdravo}
    n += 1
    if not any("deli tvrdnju" in sta for _, sta in _nalazi(par, imena, tela)):
        print("SELF-TEST: deljena tvrdnja nije prijavljena", file=sys.stderr)
        lose += 1

    # WHITELIST mora da identifikuje konkretan pad, ne njegovu klasu.
    spisak = {"poznato-ime": "PALA DRUGA TVRDNJA: bas ova tvrdnja"}
    for opis, ime2, poruka, ocekivano in (
            ("poznato ime + poznata poruka", "poznato-ime",
             "PALA DRUGA TVRDNJA: bas ova tvrdnja -- ocekivano [True]", True),
            ("poznato ime + DRUGA poruka", "poznato-ime",
             "PALA DRUGA TVRDNJA: neka sasvim druga tvrdnja", False),
            ("nepoznato ime + poznata poruka", "drugo-ime",
             "PALA DRUGA TVRDNJA: bas ova tvrdnja", False)):
        n += 1
        if poznat_nalaz(ime2, poruka, spisak) != ocekivano:
            print("SELF-TEST: whitelist -- %s" % opis, file=sys.stderr)
            lose += 1

    # zamena koja je TACNO jedan red viseredog sidra
    dvored = ("modOtkupUI.bas",
              _ST_PAR1 + ESCN + _ST_PAR2 + ESCN,
              _ST_PAR1 + ESCN,
              "T_Postoji", "tvrdnja H")
    n += 1
    if not any("podniz sidra" in sta for _, sta in _nalazi({"probni": dvored}, imena)):
        print("SELF-TEST: zamena kao podniz sidra nije prijavljena", file=sys.stderr)
        lose += 1

    if lose:
        print("self-test: %d od %d slucajeva NE MERI" % (lose, n), file=sys.stderr)
        return 1
    print("self-test: cisto (%d slucajeva)" % n)
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
    ap.add_argument("--proveri-sidra", action="store_true",
                    help="staticka provera kataloga (sidra, imena testova, zamke)")
    ap.add_argument("--self-test", action="store_true",
                    help="dokazi da svako pravilo provere stvarno meri")
    args = ap.parse_args(argv)

    if args.self_test:
        return _self_test()
    if args.proveri_sidra:
        return proveri_sidra()
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
