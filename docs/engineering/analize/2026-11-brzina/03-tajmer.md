# Tajmer (modPerf) -- predlog instrumentacije za merenje na Windows-u

Predlog, ne izmena koda. U `src-vba/` nista nije dodato. Skica modula je u `modPerf-skica.bas.txt` (ekstenzija `.txt` da je ne bi pokupili VBA alati i hook).

**Provereno u ovoj sesiji:** skica ubacena u KOPIJU repoa kao `src-vba/modPerf.bas` (CRLF): `python tools/vba_check.py` -- cisto nad celim `src-vba/`; `python tools/who_writes.py --check-ownership` -- exit 0, modPerf nije pisac. **Nije provereno:** Compile u Excelu, ijedan test, ponasanje QPC/Currency na 32/64-bit, nijedan hook u postojecim modulima (oni su samo opisani, nisu ni ubaceni).


## Sta vec postoji i sta se koristi ponovo

U repou nema profilera. Postoje rasuta merenja sa Timer-om, jedan mrezni kanal i jedan kanal u fajl. Sve sam procitao. Sta se od toga koristi ponovo:

1) modStornoImpact.bas:25 (UVID_PRAG_MS = 400), 64-89, PrijaviTrajanje 96-110: Timer po sekciji (7 sekcija), jedan red LogWarn samo kad ukupno predje prag. Komentar na 96-99 kaze da je rezolucija oko 1/64 s. PREUZETO: pravilo 'brz put ostaje tih, spor sam sebe prijavi' postaje PERF_PRAG_MS, uz jedan red po akciji u log.
2) modPaletniList.bas:321/323/389/449 (tA..tD) i PrijaviPalete: Timer, a uz vreme ide broj redova i broj paleta. PREUZETO: velicina ulaza uvek ide uz vreme (PerfVelicinaNiza belezi RxC tabele).
3) modOtkupUI.bas:498 (t0 = Timer) i 559 (mBuildMs): jedno ukupno vreme gradnje ljuske, bez sekcija. Ne dira se.
4) modBusinessFlowProTests.bas:694-722 (BfpIzabran/BfpKorakPocni/BfpKorakZatvori): Timer, prelaz ponoci (+86400), Str$ umesto Format$ da decimala bude tacka bez obzira na regionalna podesavanja. PREUZETO: oba trika (N3 i Timer rezerva u SadMs).
5) modJournaling.bas:275-390 (BackupFileOnStart): durationMs:=CLng((Timer - t0)*1000) ide u Monitor_Backup, dakle na mrezu kroz modMonitoring (HTTP_TIMEOUT_MS = 1200, modMonitoring.bas:30). NE koristi se za profiler: mreza bi merila samu sebe, a podatke bi slala napolje.
6) modLogError.bas:17-90, LogInfo 112, PurgeOldLogs 135-170: <sveska>\Log\OtkupApp_YYYY-MM-DD.log. Svaka linija radi Dir + MkDir + Open/Append/Close (55-85). Format je 'datum | LEVEL | SOURCE | br | poruka | detalji'. Purge brise samo *.log (150). PREUZETO: folder Log i LogInfo za jedan red sazetka po akciji. CSV ide u isti folder; purge ga ne brise, pa se brise rucno.
7) modSetup.bas:376-424 GetLocalConfigValue: cita tblLocalConfig celiju po celiju (petlja 405-416), a na praznu vrednost vraca default. Zato se prekidac cita JEDNOM na startu (modMain.bas:208, posle LogAppStart), nikad pri merenju. SetLocalConfigValue (modSetup.bas:426) je postojeci API vlasnika i sluzi za trajno ukljucivanje iz Immediate.
8) modDataAccess.bas:53-57: obrazac '#If VBA7 Then Private Declare PtrSafe ... #Else ... #End If' (CoCreateGuid). PREUZET za QueryPerformanceCounter/Frequency.
9) modDataAccess.bas:72 TableCacheDepth() se koristi kao META provera: prozor kesa koji ostane otvoren posle akcije bio bi greska.
10) Vec postojeci test brojac mSnimakPunjenja / Scr_IzSnimakPunjenjaTest (modScrIzvestaji.bas:1049). Profiler ga ne zamenjuje, samo stavlja svoj brojac na isto mesto.
11) Grep potvrdjuje da u src-vba nema QueryPerformanceCounter, GetTickCount ni timeGetTime. gKpiDirty (modDataAccess.bas:297-300) se upisuje, a niko ga ne cita; to nije merenje i ne koristi se.


## Odluka: nov modul ili prosirenje

NOV modul src-vba/modPerf.bas (jedan .bas). Razlog: ekvivalent ne postoji. Postojeca merenja (tacke 1-5) su lokalne Timer razlike, rezolucije oko 15,6 ms, bez ugnjezdavanja i bez brojaca poziva, i svako ima svoj format izlaza. Njihovo prosirenje bi dalo pet kopija istog mehanizma.

Odbijene alternative:
- Prosirenje modLogError: to je logger po ugovoru 'nikad ne blokira', sa Dir+Open po liniji. Stek, brojaci i QPC bi mu promenili namenu i cenu.
- modMonitoring: salje na mrezu.
- Tabela u svesci: bio bi nov pisac, a to obara A11 (who_writes).

Sta se koristi ponovo: modLogError.LogInfo, folder Log, modSetup.GetLocalConfigValue i SetLocalConfigValue, Declare obrazac iz modDataAccess, Str$ i prelaz ponoci iz BFP, pravilo praga iz modStornoImpact i modDataAccess.TableCacheDepth.

Postojeca Timer merenja se NE diraju. U postojecim procedurama se dodaju samo jednolinijski hook-ovi 'If gPerfOn Then ...', bez nove procedure i bez promene ijednog potpisa. Jedini izuzetak je jedan Dim perfAkc u modOtkupUI.UiEvent. Nijedan hook nije u generisanom modSchema.bas.

Preporuka (odluka je na operateru): modul i hook-ovi idu na zasebnu granu merenja. Da li modul i korenski hook-ovi (UiEvent, RefreshFromData, ScrGridData, clsTransaction, OnTime) ostaju trajno je zasebna odluka i zaseban PR. Kad je merac iskljucen, cena je jedno citanje Boolean-a po hook-u.


## Dizajn

API (sve Public u modPerf):
- PerfAkcijaPocni naziv / PerfAkcijaZavrsi naziv: koren merenja. Ugnjezdena akcija drugog imena (npr. REFRESH unutar Sacuvaj, ili programski Change unutar Sacuvaj) postaje sekcija 'AKCIJA.<naziv>' spoljne akcije, pa se iz CSV-a direktno vidi kaskada iz UI-7. Ugnjezdena akcija istog imena se samo broji.
- PerfPocni oznaka / PerfKraj oznaka: sekcija na steku. Za svaku oznaku se vodi: poziva, ukupno (samo spoljni okvir iste oznake, pa se rekurzija ne broji dvaput), self (trajanje minus deca), max.
- PerfBroji oznaka[, kolicina]: brojac bez vremena.
- PerfVelicinaNiza oznaka, tabela, niz: brojac sa kolicinom = broj celija, plus pamti RxC tabele. Niz se prima ByRef, da se ne kopira.
- PerfNapomena tekst: dopuna opisa akcije (npr. 'rezim=OTKUP').
- PerfTagKlasa(tag): skida cifre iz taga ('c7_3' -> 'c#_#'), pa se ne pravi po jedna oznaka za svaki red.
- Za test: PerfStatTest, PerfMetaTest, PerfIskljuciTiho, PerfSadMs.
- Alt+F8 makroi: Perf_Ukljuci (nivo 1), Perf_UkljuciDetalj (nivo 2), Perf_Iskljuci, Perf_Reset, Perf_Sazetak.

PREKIDAC: Public gPerfOn (nivo 1: akcije, sekcije domena, TX, brojaci primitiva) i gPerfDetalj (nivo 2: i vremena primitiva modDataAccess). Postavlja ih PerfInitIzKonfiga jednom na startu (modMain.bas:208), iz tblLocalConfig PERF_MERENJE = DA | DETALJ | NE (podrazumevano NE), uz PERF_PRAG_MS. Mogu se ukljuciti i za jednu sesiju kroz Alt+F8. Prekidac je po masini, jer je tblLocalConfig per-masina tabela.

TROSAK KAD JE ISKLJUCEN: pozivno mesto glasi 'If gPerfOn Then PerfPocni ...', pa se iskljucen merac svodi na jedno citanje globalnog Boolean-a i skok. Nema poziva procedure i nema spajanja stringa oznake. PROCENA (nije mereno): desetine ns po hook-u, ispod 0,1 ms po akciji za oko 1.000 hook-ova.

TROSAK KAD JE UKLJUCEN: meri ga Kalibrisi pri ukljucivanju (2000 parova Pocni/Kraj i 2000 brojaca). Rezultat ide u Immediate, u sazetak i u META OverheadProcenaMs svake akcije. PROCENA: 2-10 us po paru, 1-5 us po brojacu. Na nivou 2, sa oko 500 GetColumnIndex i oko 100 GetTable po snimanju, to je nekoliko ms; zato je nivo 2 odvojen.

SAT: QueryPerformanceCounter/Frequency kroz Currency. Currency je 64-bitni ceo broj skaliran sa 10^4; QPC upise svojih 64 bita u njega, a odnos (razlika brojaca)/frekvencija skalu ponistava. Razlika se racuna u Currency (tacno), pa tek onda u Double. Isto radi u 32- i 64-bitnom Office-u. LongLong se namerno ne koristi, jer postoji samo u 64-bitnom. Declare je pod #If VBA7 PtrSafe / #Else. Kad QPC ne radi, rezerva je Timer uz prelaz ponoci. NEPROVERENO na masini operatera; test T_Perf_QpcMonoton.

ROBUSNOST:
(a) Err-neutralno: svaka javna procedura sa On Error na ulazu sacuva Err.Number/Source/Description, a na izlazu ga vrati (ErrVrati). Poziv iz On Error Resume Next bloka zato ne menja ishod pozivaoca.
(b) Tolerantan stek: Kraj trazi oznaku odozgo; sve sto je iznad nje ostalo je otvoreno (izlaz kroz Err.Raise) i zatvara se uz brojanje META Nezatvorenih. Merena procedura zato ne treba sopstveni EH. Kraj bez para se broji u Neuparenih. Stek ima 512 mesta, a preliv se broji.
(c) Akcija otvorena duze od 10 min smatra se procurelom i zatvara se sa napomenom NEZAVRSENA.
(d) U test rezimu nema upisa fajla ni LogInfo-a; statistika ostaje u memoriji za PerfStatTest.

IZLAZ (nista u tabele sveske):
- <folder sveske>\Log\Perf_YYYY-MM-DD.csv, separator ';', decimalna tacka (Str$). Kolone: datum;t0_ms;rb_akcije;akcija;rb_istog;vrsta;oznaka;poziva;ukupno_ms;self_ms;max_ms;kolicina;napomena.
- vrsta = AKCIJA (ukupno, self = aktivno = ukupno - UI.CekaOperatera), META, SEK, BROJ ili VEL.
- META redovi: CekaOperateraMs, OverheadProcenaMs, NemerenoMs, Calculation, QpcHz, Nivo, Neuparenih, Nezatvorenih, Preliv, VanAkcije, TableCacheDepth.
- rb_istog = 1 znaci prvu akciju tog imena od ukljucenja; ako je ukljuceno iz konfiga, to je HLADNO.
- t0_ms je vreme od ukljucenja, pa se vidi i razmak izmedju Sacuvaj i ONTIME.AUTOSAVE/ONTIME.WARM.
- Fajl se otvara JEDNOM po akciji (ceo tekst jednim Print-om). Dir i MkDir idu samo kad se promeni datum.
- Uz to jedan red LogInfo 'modPerf' po akciji iznad praga (top 3 po self) u postojeci OtkupApp_*.log.
- Perf_Sazetak pravi Log\Perf_sazetak_*.txt i ispis u Immediate: po akciji broj, prvi (hladno), prosek ostalih, min i max; TOP 25 akcija|oznaka po self; svi brojaci; najvece vidjene dimenzije tabela. MsgBox ima samo u rucnom makrou i samo van test rezima.

GDE SE AKCIJA OTVARA I DAMPUJE: koren je modOtkupUI.UiEvent (4274), jedini dispecer svih dogadjaja kontrola. CommitDokument (Sacuvaj), RunRowAction/ScrAct (izmene i radnje), prekidaci lista, pretraga i hover su tako svaki svoja akcija, sa dumpom na kraju. RefreshFromData je ugnjezdena akcija REFRESH (samostalna kad je pozove nesto van UiEvent-a). modScrIzvestaji.Scr_Rows je ugnjezdena akcija IZV.ROWS. OnTime poslovi su sopstvene akcije: ONTIME.AUTOSAVE (AutoSaveTick) i ONTIME.WARM (BuildWarm).


## Protokol merenja (radi operater na Windows-u)

Sve ispod radi operater na Windows-u. U ovoj sesiji nista nije pokrenuto.

0) GRANA MERENJA (pravi je neko drugi; ova analiza ne menja repo). U Git Bash-u, iz root-a repoa:
   cd ~/Documents/GitHub/otkupapp-pwa
   git fetch origin
   git checkout <grana-merenja>
   python tools/vba_check.py
   python tools/sabotaza.py --proveri-sidra
   python tools/who_writes.py --check-ownership
   python tools/gen_schema_module.py --check
   Sve cetiri moraju biti zelene pre uvoza.

1) KOPIJA PRAVE SVESKE. Excel mora biti potpuno zatvoren. Putanje su primer, u bash obliku sa /:
   mkdir -p /c/Merenje/pre /c/Merenje/post
   cp "/c/<folder>/<sveska>.xlsm" /c/Merenje/baza.xlsm
   Nikad se ne meri na produkcionoj svesci.

2) PRIPREMA baza_pre.xlsm: otvori kopiju, uvezi kod grane merenja uobicajenim ImportAllVBA putem (folder ~/Documents/GitHub/otkupapp-pwa), pa Alt+F11 -> Debug -> Compile VBAProject. U Immediate (Ctrl+G):
   SetLocalConfigValue "PERF_MERENJE", "DA"
   SetLocalConfigValue "PERF_PRAG_MS", "0"
   Zatim u kopiji iskljuci cloud sync i monitoring, da kopija ne gura podatke u produkcioni GAS (BulkPushPendingForStanica, MasterSync, Monitor_*). Koji je to kljuc NIJE provereno, pa ga treba naci u Podesavanjima. Ako se ne moze iskljuciti, radi bez mreze i zabelezi to, jer HTTP timeout-i iskrivljuju UI-7 i DOD-2.
   Snimi, zatvori i preimenuj u baza_pre.xlsm. baza_post.xlsm se pravi isto, sa granom koja nosi izmenu.

3) PROLAZ = svez fajl: cp /c/Merenje/baza_pre.xlsm /c/Merenje/pre/run01.xlsm, pa otvori run01. Excel pre toga mora biti zatvoren. Prva akcija svakog imena posle otvaranja je HLADNA (rb_istog = 1): ukljucuje kompajliranje modula na zahtev i prvo citanje opsega. Log ide u /c/Merenje/pre/Log/. Posle starta u Immediate pogledaj: QPC=DA i overhead u us (red iz Perf_Ukljuci ili prvi LogInfo).

4) SCENARIJI. Ulaz (kooperant, OM, kolicine, gajbe) se unapred zapise i isti je u pre i post. Za svaki scenario: 1 hladan (prvi u svom prolazu; za hladno merenje glavnih tokova po jedan prolaz koji pocinje bas tim tokom) i 10 toplih. Izmedju dva snimanja 5-10 s pauze.
   S1 otkup bez nacrta, 1 klasa i 2 klase, podrazumevana stampa (OTK-1/2/3/6-10/14/16).
   S2 otkup uz aktivan nacrt; svako snimanje je tacka krive k = 1..20 (OTK-4/5, OTP-3, DOD-2).
   S3 otkup sa manjkom gajbi, odgovor Da (OTK-11, ZAJ-02).
   S4 otkup na hladnjackoj stanici (RUTA_LANAC, nepokriveno).
   S5 izmena nacrta otpremnice, k = 1/5/20 (IZM-1, OTP-2, IZM-12).
   S6 izdavanje otpremnice, k = 1/5/20 (OTP-1, IZM-3).
   S7 zbirna: nacrt, izmena i izdavanje, m = 1/5/10 (ZBR-1, IZM-10).
   S8 prijemnica: samo pune, sa povratom, bez gajbi (PRJ-1/2, PAL-1, AMB-1).
   S9 isplata obicna i spec sa d = 3 dela (NOV-1).
   S10 ispravka slobodnog bloka, bloka u nacrtu i lista u izdatoj; otvaranje i snimanje odvojeno (OTK-13, IZM-2/4/5/6/7/8).
   S11 Izvestaji: entitet A->B->A x3, tip x5, rezim x5, opseg x5; zbirni saldo OM, zbirna ambalaza, zbirni kupci, RANG za period sve; 10 otkucaja pretrage; stampa kartice (IZV-1..11, UI-4).
   S12 UI: hover preko 22 reda, 5 slova pretrage na Dokumentima, 5 promena rezima, 5 prekidaca liste, 10 klikova Dodaj u korpi Agrohemije (UI-3/5/6/9/10/11, DOD-1).
   S13 posle jednog Sacuvaj 70 s mirovanja, pa niz od 5 otkupa bez pauze (ONTIME.AUTOSAVE i ONTIME.WARM: OTK-12/15, ZAJ-03).

5) NIVOI. Ceo set prvo na nivou 1 (PERF_MERENJE=DA); to su brojevi za poredjenje. Zatim jedan prolaz S1, S6, S11 na nivou 2 (DETALJ) za raspodelu po primitivima. Razlika aktivnog vremena nivo 2 minus nivo 1 je stvarni overhead; uporedi je sa META OverheadProcenaMs.

6) KRAJ PROLAZA: Alt+F8 -> modPerf.Perf_Sazetak (daje Log/Perf_sazetak_*.txt), pa zatvori Excel.

7) A/A PROVERA SUMA: ponovi isti pre prolaz u run02. Razlika medijana run01/run02 je sum merenja. Ako je veca od 10%, povecaj broj toplih na 20.

8) VELICINA PODATAKA: VEL redovi daju stvarne N_otkup, N_stavke, N_amb... Za rast po N (OTK-1 kvadratno, OTP-3 po k), ako schema_diff ne blokira (python tools/schema_diff.py "<backup>.xlsm"), ponovi S1/S2 i na starijoj kopiji iz <sveska>/Backup, koja ima manje redova. Dve tacke N su dovoljne za razliku linearno/kvadratno.

9) CITANJE (Git Bash, po jednom fajlu po prolazu, jer rb krece od 1 u svakom prolazu):
   cd /c/Merenje/pre/Log
   Aktivno po akciji, toplo:
   awk -F';' '$6=="AKCIJA" && $5>1 {n[$4]++; s[$4]+=$10} END {for (a in n) printf "%10.1f %4d %s\n", s[a]/n[a], n[a], a}' Perf_*.csv | sort -rn
   Hladno:
   awk -F';' '$6=="AKCIJA" && $5==1 {print $10, $4}' Perf_*.csv
   Top sekcije za Sacuvaj otkupa:
   awk -F';' 'NR==FNR {if ($6=="AKCIJA" && $13 ~ /rezim=OTKUP/ && $5>1) ok[$3]=1; next} ($3 in ok) && $6=="SEK" {s[$7]+=$10} END {for (k in s) printf "%10.1f %s\n", s[k]/length(ok), k}' Perf_X.csv Perf_X.csv | sort -rn | head -30
   Brojaci po akciji: isto, sa $6=="BROJ" i $8 (poziva) ili $12 (celije).
   U Excel-u: Podaci -> Iz teksta/CSV, razdvajac ';', decimalni separator '.'.

10) KRITERIJUM (predlog):
   - Nalaz je POTVRDJEN kad njegova sekcija nosi bar 10% aktivnog vremena akcije ili bar 100 ms (prosek toplih, kopija prave sveske).
   - OBOREN je kad nosi ispod 2% i ispod 20 ms.
   - Izmedju je nizak prioritet.
   - Brojaci (broj GetTableData po tabeli, broj KnjigaZaCitanje, OtpremnicaZaOtkup...) su deterministicni i najjaci su dokaz: tvrdnja '10 citanja' se potvrdjuje ili obara tacnim brojem, nezavisno od suma.
   - PRE/POSLE: isti scenario nad baza_pre i baza_post, medijana toplih. Poboljsanje se prihvata kad je razlika veca od dvostruke A/A razlike I kad je ciljani brojac pao koliko predlog tvrdi (npr. DA.GTD|tblAmbalaza 10 -> 2).
   - META provere: TableCacheDepth mora biti 0, a Nezatvorenih/Neuparenih se objasnjavaju pre zakljucka.
   PROCENA trajanja celog protokola za operatera: 2-4 h (nije mereno).


## Rizici i sta se ne sme

Nista od ovoga nije izvrseno ni kompajlirano (Linux, bez Excela). Sve sto tvrdi ponasanje u runtime-u je NEVERIFIKOVANO.

STA JE PROVERENO: skica (modPerf.bas) je ubacena u KOPIJU src-vba u scratchpad-u.
- python3 tools/vba_check.py nad tom kopijom: cisto (195 fajlova). Prosle su provere ASCII, DEKLARACIJA, REZERVISANO, DUPLIKAT, DUPLIKAT_LOKALNI, DUPLI_LOKAL, ZAKLONJENO, KRAJ_REDA (CRLF) i ARNOST.
- who_writes.py --check-ownership: exit 0, modPerf se ne pojavljuje kao pisac.
- Hook-ovi u postojecim modulima NISU provereni nijednim alatom.

A11 / WHO_WRITES:
- modPerf ne zove AppendRow, UpdateCell ni DeleteRow, a izlaz je fajl.
- SetLocalConfigValue se zove samo rucno iz Immediate (postojeci API vlasnika modSetup; tblLocalConfig nije u WRITE_OWNERSHIP.json).
- ZAMKA: MUTATE_DYN_RE (tools/who_writes.py:93) broji svako 'AppendRow ime,' ili 'UpdateCell ime,' i u tekstu i komentaru. Zato su oznake 'DA.Append|' i 'DA.Update|' i nigde se ne pise 'AppendRow x,'.
- Posle umetanja obavezno: python tools/who_writes.py --check i --check-ownership.

ERR I UGOVOR EKRANA:
- Profiler je Err-neutralan (ErrVrati). To da postavljen Err.Number prezivi izlaz iz procedure sa On Error Resume Next NIJE provereno u Excelu. Oslanja se na zapisano ponasanje 'zaostatka' (KI-014). Mora ga potvrditi test T_Perf_ErrNeutralan.
- Svejedno, hook NE ide: izmedju Application.Run i ScrPadPreuzmi/ScrRunIshod (modUiScreens.bas:622-629 i 698-699, UI-ISHOD-01; ScrRunIshod namerno nema On Error); izmedju RestoreTable i If Err.Number (clsTransaction.cls:149-150); unutar On Error Resume Next bloka GetColumnIndex (modDataAccess.bas:215-217); pre LogErr u EH (LogErr cita Err).
- KI-016 / UI-PAD-01: u Scr_* clanovima hook ide POSLE linije ScrPadAkoTest, nikad izmedju 'On Error GoTo EH' i nje. Taj par je sidro sabotaza ki016-*-pusta-pad.

SIDRA SABOTAZA: posle umetanja obavezno python tools/sabotaza.py --proveri-sidra. Pogodjeni delovi koda: modDataAccess (kes-kolone-pamti-nulu 231-233, mapa-reda-pomerena-za-jedan), clsTransaction (rollback-petlja-staje-na-padu, rollback-cleanup-samo-kad-prodje, odl40-commit-bez-kapije na CommitTx), ki015 u PotvrdiManjakAmbalaze (modOtkupUnos.bas:516-536), maticni-prozor-kesa-ostaje-otvoren. Hook se ne umece unutar teksta sidra.

GENERISAN KOD: modSchema.bas se ne instrumentuje (gen_schema_module.py --check bi pao). Meri se samo na pozivnim mestima SchemaReadyOrFail.

FORME I KLASE: nema WithEvents ni Application dogadjaja (zabrana iz forme-i-kontrole.md; vba_hard_census). Zato se preracun ne broji dogadjajem, nego se meri sekcijom TX.CleanUp.Calc i META Calculation.

MSGBOX: nema ga u putu upisa. Perf_Sazetak ga ima samo kao rucni makro i samo uz 'If Not IsTestMode()'.

TESTOVI:
- Iskljucen merac ne menja ponasanje, pa nijedan postojeci test ne bi trebalo da se promeni. To NIJE provereno; posle umetanja obavezno TARGETED/FULL na Windows-u.
- U test rezimu modPerf ne pise ni fajl ni log.
- Predlozeni novi testovi u modTest: T_Perf_IskljucenNistaNeBelezi, T_Perf_ErrNeutralan, T_Perf_UgnjezdenoSelfIUkupno, T_Perf_TolerantStek (Nezatvorenih i Neuparenih), T_Perf_UgnjezdenaAkcija, T_Perf_QpcMonoton. Svaki mora proci i sam (S-4a).
- Posto su to novi testovi, trazi se dokaz u oba smera: python tools/dokaz.py --izbor --radnika 4 <prefiks>. Kandidati za sabotazu: ukloniti ErrVrati, obrnuti oduzimanje dece u self, ukloniti uslov mAkt(k) <= 0.
- Test koji ukljuci merac gasi ga sam, a modTest.ResetSeamova (904) dobija modPerf.PerfIskljuciTiho. Merac nije seam, ali nosi stanje.

IZOBLICENJE MERENJA:
- Nivo 2 dodaje stotine poziva po akciji. META OverheadProcenaMs je DONJA granica, jer ne meri If na pozivnom mestu ni spajanje oznake. Zato se poredi nivo 1 sa nivoom 2.
- PerfVelicinaNiza i GetTableData su ByRef, da se ne kopira niz.
- Oznake ne smeju da nose ID, broj dokumenta ni ime partnera (PerfTagKlasa skida samo cifre). Da nijedan tag kontrole ne nosi ime partnera NIJE provereno; zato je oznaka skracena na 40 znakova.

FAJL:
- Dir i MkDir idu samo pri promeni datuma. Akcija se nikad ne zatvara unutar Dir petlje (PurgeOld* se ne instrumentuje).
- Perf_*.csv ne brise PurgeOldLogs (brise samo *.log), a sa pragom 0 i hover-om fajl brzo raste. Za duzi rad preporuka je PERF_PRAG_MS = 20.

KOPIJA SVESKE: ako cloud sync i monitoring ostanu ukljuceni, kopija moze da posalje podatke u produkciju. Ovo je najveci operativni rizik protokola.

STA NE SME:
- merni kod u main bez odluke;
- hook koji menja potpis, uvodi Private WithEvents ili pise u tabelu;
- izmena .claude/ zajedno sa ovim;
- tumacenje vremena bez brojaca uz njega;
- prijava 'zeleno' bez run_vba i rucnog Compile-a.

32/64-bit: trik sa Currency i PtrSafe radi u oba po dokumentovanom ponasanju VBA, ali NIJE proveren na masini operatera. Ako postoji, test T_Perf_QpcMonoton se pusta i u 32-bitnom Office-u.


## Tacke merenja (vezane za ID nalaza)

| # | Tok | Oznaka | Mesto | Sta meri | Nalazi |
|---|---|---|---|---|---|
| 1 | zajednicko | `UI.<ev>.<tag#> (akcija, koren)` | `modOtkupUI.bas:4284` | Pocni posle fz-grane (4283), Zavrsi posle End Select (4296), Dim perfAkc na vrhu UiEvent (4274). Ukupno i aktivno vreme svake radnje operatera. Programski Change unutar Sacuvaj ulazi kao sekcija AKCIJA.UI.Change.* (broj i trajanje kaskade). | okvir za sve; direktno UI-7, UI-5, UI-9, UI-4, IZV-1, IZV-2, IZV-11, DOD-1(ui) |
| 2 | snimanje | `UI.CekaOperatera` | `modOtkupUI.bas:8603` | Oko MsgBox u Dijalog (posle test grane 8599-8602). Vreme citanja poruke se oduzima od aktivnog vremena. | sve akcije snimanja (OTK-*, OTP-*, PRJ-*, NOV-1) |
| 3 | snimanje | `napomena rezim=<rezim>` | `modScrDokumenti.bas:2046` | PerfNapomena pre Select Case 2047. Razvrstava Sacuvaj po rezimu: OTKUP, OTPREMNICA, ZBIRNA, PRIJEMNICA, AMB_ISPLATE, AMB_UPLATE, REVERSI. | razvrstavanje za OTK-*, OTP-*, ZBR-1, PRJ-*, NOV-1 |
| 4 | izmena | `SAVE.OtpremnicaZaOtkup` | `modScrDokumenti.bas:2125` | Cena kanonskog clanstva pre upisa ispravke, uz brojac OTP.OtpremnicaZaOtkup. | OTK-13, IZM-4 |
| 5 | snimanje | `SAVE.Prekoracenje / SCR.PrekoracenjeOpis / UI.CekaOperatera` | `modScrDokumenti.bas:2137` | Ceo poziv PotvrdiPrekoracenje. Unutra posebno PrekoracenjeOpis (1967, GetOtpremnicaProgress) i MsgBox (1969-1971) kao cekanje. | nepokriveno OTK (PrekoracenjeOpis), UI-1 (isti citaci) |
| 6 | izmena | `SAVE.PotvrdaLanca + UI.CekaOperatera` | `modScrDokumenti.bas:2147` | Racun plana lanca u potvrdi. MsgBox 1932-1936 kao cekanje. Za slobodan blok pokazuje bacen rad. | IZM-5, IZM-8 |
| 7 | snimanje | `SAVE.OtkupUpisi` | `modScrDokumenti.bas:2153` | Ceo pisac otkupa sa stampom (granica ekran/pisac). | OTK-* ukupno |
| 8 | snimanje | `SAVE.RutaPosleUpisa + brojac HLAD.LanacVaziZaBlok` | `modScrDokumenti.bas:2168` | Trajanje rute. Brojac na pocetku tela LanacVaziZaBlok (modAutoHladnjaca.bas:128) treba da pokaze 1 bez nacrta i 2 uz nacrt. | OTK-4, DOD-2(snimanje) |
| 9 | snimanje | `SAVE.AutoLanac` | `modScrDokumenti.bas:2175` | RUTA_LANAC na hladnjackoj stanici (otpremnica, zbirna i prijemnica u lancu). Danas niko nije izbrojao. | nepokriveno OTK (AutoLanacHladnjaca) |
| 10 | snimanje | `SAVE.VeziZaAktivnu` | `modScrDokumenti.bas:2178` | Vezivanje bloka za aktivan nacrt, po snimanju, za k = 1..20 vec vezanih blokova (kriva po k). | OTK-5, OTP-3, IZM-9, DOD-2(izmena) |
| 11 | snimanje | `OTK.CreateOtkup_TX` | `modOtkupUnos.bas:379` | Prvi poziv pisca. Self pokazuje udeo VBA logike naspram citanja tabela (deca DA.* na nivou 2). | OTK-2, OTK-6, OTK-7, OTK-8, OTK-9, OTK-10, ZAJ-01 |
| 12 | snimanje | `OTK.PutManjka + UI.CekaOperatera` | `modOtkupUnos.bas:399` | 399-412: OtkupDeficitKooperanta i drugi CreateOtkup_TX. MsgBox u PotvrdiManjakAmbalaze (536) je cekanje. Rollback se vidi kroz TX.Restore. | OTK-11, ZAJ-02 |
| 13 | snimanje | `PRINT.OutputOtkupniList` | `modOtkupUnos.bas:422` | Stampa i PDF. Pocni posle On Error Resume Next (421), Kraj pre Err.Clear (423). | OTK-3, OTK-1, OTK-14, OTK-16 |
| 14 | snimanje | `PRINT.Stavke.Petlja + VEL tblOtkupStavke` | `modPrint.bas:705` | Kvadratna petlja 705-741. Uz nju PerfVelicinaNiza ds (N_stavke). Kvadratni rast se proverava na dve velicine sveske. | OTK-1 |
| 15 | snimanje | `PRINT.AmbOpening` | `modPrint.bas:783` | GetKooperantAmbOpening posle commit-a. Treba da bude jos jedan pogodak brojaca knjige. | OTK-14 |
| 16 | snimanje | `PRINT.WriteOtkupCopy + brojac PRINT.Calc=<n>` | `modPrint.bas:839` | 839-842, oba primerka. Brojac nosi Application.Calculation u tom trenutku, da se vidi da li pisanje sablona okida preracun. | OTK-16 |
| 17 | snimanje | `PRINT.DocExportPdf` | `modPrint.bas:508` | ExportAsFixedFormat i otvaranje citaca, odvojeno od punjenja sablona. | OTK-3 |
| 18 | snimanje | `OTK.Schema` | `modOtkup.bas:93` | SchemaReadyOrFail nad 4 tabele (93-94). Meri se na pozivnom mestu, jer je modSchema generisan. | OTK-8, ZAJ-10, SCH-1 |
| 19 | snimanje | `OTK.ApplyAvans + brojac NOV.GetUplataForOtkup` | `modOtkup.bas:898` | Trajanje ApplyAvansToOtkup. Brojac je na pocetku GetUplataForOtkup (modNovac.bas:976). | OTK-6 |
| 20 | izmena | `ISP.Kapije / ISP.TX` | `modOtkup.bas:343` | IspravkaOtkupa_TX. Kapije 343-391 (Pocni posle 341), transakcija 392-415, plus snapshot-i kroz TX.Snapshot. | OTK-13, IZM-4, IZM-11 |
| 21 | zajednicko | `OTK.StavkeOtkupaRedovi / OTK.ZbirStavkiPoOtkupu` | `modOtkup.bas:1203` | Brojac i trajanje gradnje (1203, posle Const; 1468). Ocekivano do 4 puta po RefreshFromData i S puta po zbirnom saldu OM. | UI-2, IZV-3, DOD-1(izvestaji) |
| 22 | zajednicko | `AMB.KnjigaZaCitanje\|<src> + AMB.KnjigaIntegritet + VEL tblAmbalaza` | `modAmbalaza.bas:829` | Brojac po sourceName pre 829, sekcija oko KnjigaIntegritet (829), PerfVelicinaNiza data (N_amb). Ocekivano 10 citanja po otkupu i 11 sa stampom. | OTK-2, OTK-14, AMB-1, DOD-3(snimanje), ZAJ-01, ZAJ-D1, IZV-4 |
| 23 | zajednicko | `SCHEMA.Knjiga` | `modAmbalaza.bas:543` | SchemaReadyOrFail(tblAmbalaza) iz RequireKnjigaSchema: broj i trajanje po akciji (ocekivano oko 8 po otkupu). | OTK-8, ZAJ-10, SCH-1, DOD-3(snimanje), IZV-4 |
| 24 | zajednicko | `brojac AMB.PrenesiAmbalazu` | `modAmbalaza.bas:4126` | Broj prenosa po akciji, pre RequireKnjigaSchema (4127). Sluzi kao imenilac za citanja knjige po prenosu. | AMB-1, ZAJ-D1, OTK-2 |
| 25 | snimanje | `OTP.Dodaj.Kapija` | `modDokumenta.bas:3541` | 3541-3542: OtpRequireIzvorValjan i OtpUpisiClanstvo u DodajOtpremnicaIzvor_TX. | OTK-5, OTP-3, IZM-9, DOD-2(izmena) |
| 26 | zajednicko | `OTP.RequireIzvorValjan` | `modDokumenta.bas:6289` | Pocni pre 6289, Kraj pre 6298. Broj poziva (= broj clanova) i trajanje po pozivu. Deca DA.GTD\|tblOtkup pokazuju oko 7 citanja po pozivu. | OTP-1, OTP-2, OTP-3, IZM-1, IZM-2, IZM-3, IZM-9 |
| 27 | zajednicko | `OTP.ClanstvoKanon` | `modDokumenta.bas:7710` | AktivnoOtpClanstvoPoKanonu: Pocni posle 7709, Kraj pre 7777 (rani izlaz 7713 ostaje kao nezatvoren). Broj gradnji recnika po akciji. | DOD-1(izmena), OTP-1, OTK-13 |
| 28 | izmena | `brojac OTP.OtpremnicaZaOtkup` | `modDokumenta.bas:7519` | Broj poziva po akciji. Ocekivano 6 pri snimanju ispravke bloka, 9 za list u izdatoj, 2-3 pri otvaranju. | OTK-13, IZM-4, IZM-6 |
| 29 | izmena | `OTP.IzmeniDraft.Revalidacija` | `modDokumenta.bas:6054` | 6054-6061: OtpClanovi i petlja revalidacije, za k = 1, 5, 20. | IZM-1, OTP-2 |
| 30 | snimanje | `OTP.Izdaj.Revalidacija` | `modDokumenta.bas:6563` | OtpClanovi (6550) i petlja 6563-6565 u OtpIzdaj, za k = 1, 5, 20. | OTP-1, IZM-3 |
| 31 | izmena | `OTP.Ispravi.Petlja` | `modDokumenta.bas:6801` | 6801-6804, za k = 2, 5, 10, 20: odvaja linearni od kvadratnog dela. | IZM-2 |
| 32 | izmena | `LNC.Sastavi.Petlja / LNC.Sastavi.Izdaj` | `modDokumenta.bas:4988` | Petlja 4988-4991 i OtpIzdaj na 4993: dvostruka revalidacija nove verzije. | IZM-3 |
| 33 | izmena | `LNC.KapijePreTx / LNC.Snimi / LNC.PlanUTx` | `modDokumenta.bas:4788` | Kapije 4788-4791, LncSnimi na 4795 (13 snapshot-a), plan 4798-4807. | IZM-8, IZM-11 |
| 34 | izmena | `ZBR.IzmeniDraft.Revalidacija` | `modDokumenta.bas:2117` | 2117-2122, za m = 1, 5, 10 otpremnica. | IZM-10, ZBR-1 |
| 35 | snimanje | `brojac ZBR.IzvoriNevaljaniRazlog` | `modDokumenta.bas:2421` | Broj poziva po izdavanju zbirne i po osvezavanju trake. Vreme dolazi iz roditeljske sekcije i iz DA.* dece. | ZBR-1 |
| 36 | snimanje | `PRJ.Validiraj` | `modScrDokumenti.bas:2466` | PrijemnicaValidiraj naspram pisca. Brojaci DA.GTD\|tblZbirna i \|tblPrijemnica u oba dela. | PRJ-1 |
| 37 | snimanje | `PRJ.Snapshot / PRJ.Paletize` | `modDokumenta.bas:8711` | Blok snapshot-a 8711-8716 i PaletizePrijemnica (8760). Na rollback-u se vidi TX.Restore po tabeli. | PRJ-2, PAL-1 |
| 38 | zajednicko | `TX.Snapshot\|<tbl> + VEL` | `clsTransaction.cls:77` | Pocni posle 77, Kraj posle 84, uz broj celija (Value2) po tabeli. Broj snapshot-a po akciji. tblOtkup u isplati. | ZAJ-04, IZM-11, NOV-1, PRJ-2, OTK-11 |
| 39 | zajednicko | `TX.CleanUp.Calc` | `clsTransaction.cls:252` | Vracanje Application.Calculation posle svakog commit-a i rollback-a. Pokazuje da li sveska tada preracunava (formule nisu u repou). | OTK-16, nepokriveno zajednicki sloj (preracun po commit-u) |
| 40 | zajednicko | `TX.Commit` | `clsTransaction.cls:93` | Od posle 93 (If Not mActive) do kraja CommitTx: CleanUp, MarkDirtyAndSchedule, InvalidateStornoWarm. | OTK-12, OTK-15 |
| 41 | zajednicko | `TX.Restore\|<tbl>` | `clsTransaction.cls:264` | U telu RestoreTable (posle 264; Kraj i pre Exit Sub 278). NE izmedju 149 i 150 u RollbackTx. | ZAJ-02, OTK-11, PRJ-2 |
| 42 | zajednicko | `DA.GetTable\|<tbl> (brojac) / DA.GetTable.ms\|<tbl> (nivo 2)` | `modDataAccess.bas:118` | Broj skenova listova po akciji i cena po tabeli (redosled listova). Kraj i pre Exit Function u petlji. | OTK-7, ZAJ-1, ZAJ-07 |
| 43 | zajednicko | `DA.GTD\|<tbl> (kolicina = celije) / DA.GTD.KES\|<tbl> / .ms (nivo 2)` | `modDataAccess.bas:137` | Pogodak kesa (137) i citanje sa lista (151) po tabeli, uz broj kopiranih celija. Na nivou 2 i cena kopije iz kesa. | OTK-9, ZAJ-01, DOD-3(izmena), UI-1, UI-2, IZV-6, OTP-1, IZM-1..IZM-10 |
| 44 | zajednicko | `DA.ColIdx.KES / DA.ColIdx\|<tbl>` | `modDataAccess.bas:205` | Pogodak mColCache (205) naspram promasaja (posle 208, svaki promasaj je GetTable sken). Ocekivano oko 430 promasaja po snimanju otkupa. | OTK-7, ZAJ-1 |
| 45 | zajednicko | `DA.Append\|<tbl> + DA.Append.ListRowsAdd / .Celije / .Audit / .Journal` | `modDataAccess.bas:283` | Brojac posle 279. Na nivou 2 pod-sekcije 283, 289-291, 294 i 295 (WriteJournalRow), plus ErrHandler. | ZAJ-05, ZAJ-09 |
| 46 | izmena | `DA.Update\|<tbl> (+ .ms nivo 2)` | `modDataAccess.bas:327` | Brojac posle 326, sekcija 327-330. Ocekivano 8 po izmeni nacrta otpremnice. | ZAJ-06, IZM-12 |
| 47 | zajednicko | `DA.FindRows\|<tbl> / DA.LookupValue\|<tbl> / DA.BuildLookupDict\|<tbl>` | `modDataAccess.bas:631` | Brojaci u prvom redu tela FindRows (432), LookupValue (631) i BuildLookupDict (661). | OTK-9, ZAJ-D2, UI-10, IZV-2 |
| 48 | snimanje | `brojac CFG.GetConfigValue` | `modConfig.bas:1283` | Broj poziva po akciji, pre LookupValue (1283). Ocekivano 15-18 po otkupu. | OTK-10 |
| 49 | zajednicko | `HLP.ExclStorno.KopijaZaPoredjenje` | `modHelpers.bas:180` | Kopija pune tabele iz kesa samo radi poredjenja broja redova (180-184). | IZV-6 |
| 50 | ui-osvezavanje | `REFRESH (ugnjezdena akcija) + RF.ResetCache/Kpi/Nav/ZbirneCombo/MaticniCombos/ReloadGrid/OtpTraka` | `modOtkupUI.bas:6266` | Pocni posle 6265, Zavrsi pre 6294. Sekcije oko 6266-6270, 6274, 6275, 6276, 6277, 6291 i 6293. Broj REFRESH po akciji Izvestaja. | UI-1, ZAJ-08, IZV-1, UI-4, IZV-8, UI-7, UI-8, UI-6, DOD-1(izvestaji) |
| 51 | ui-osvezavanje | `KPI.KgDana / KPI.SaldoOM / KPI.Dokumenata` | `modOtkupUI.bas:4246` | KgOtkupaZaDan (4240), SaldoOMUkupno (4246) i CountForDate (4259-4260). Broj RefreshKpi po promeni rezima i liste. | UI-11, IZV-8, DOD-1(ui) |
| 52 | ui-osvezavanje | `SCR.BROJAC\|<modul>` | `modUiScreens.bas:698` | Pocni izmedju 697 i 698, Kraj posle 700. Cena znacke po ekranu (Oporavak, Banka nalozi...) pri svakom OsveziNavBrojace. | UI-3, UI-12, DOD-1(izvestaji), ZAJ-D2 |
| 53 | ui-osvezavanje | `SCR.ROWS\|<modul>` | `modUiScreens.bas:622` | Pocni izmedju 621 i 622, Kraj posle 626 i posle 634. Cena liste po ekranu pri kucanju, kliku cipa i osvezavanju. Nikad izmedju 622 i 629. | UI-9, UI-10, IZV-2, IZV-11, UI-1 |
| 54 | ui-osvezavanje | `UI.LayoutGrid / UI.LayoutGrid.Celije` | `modOtkupUI.bas:2752` | Petlja celija 2752-2765 naspram cele LayoutGrid (od 2605). Broj poziva po Sacuvaj i po prelazu na ekran. | UI-6 |
| 55 | ui-osvezavanje | `brojac UI.RenderGrid` | `modOtkupUI.bas:2889` | Broj crtanja mreze po akciji (posle provere mFrm na 2889). U akciji UI.Hover.* ukupno vreme je cena prelaza preko reda. | UI-5 |
| 56 | ui-osvezavanje | `UI.ZbirneSort + brojac UI.ZbirneJ (kolicina = j)` | `modOtkupUI.bas:7007` | Kvadratni sort 7007-7016 u FillZbirneCombo, uz j. | UI-8 |
| 57 | izvestaji | `IZV.ROWS (ugnjezdena akcija)` | `modScrIzvestaji.bas:686` | Pocni POSLE 686, Zavrsi pre 701 i posle 703. | IZV-1, UI-4 |
| 58 | izvestaji | `brojac IZV.SnimakPunjenje + IZV.EntitetNaziv` | `modScrIzvestaji.bas:1064` | Brojac na 1049 (isto mesto kao mSnimakPunjenja). Sekcija oko EntitetNaziv (1064) pri svakom otkucaju. | IZV-1, IZV-2 |
| 59 | izvestaji | `IZV.Puni\|<kljuc>` | `modScrIzvestaji.bas:1082` | Pocni posle 1081, Kraj izmedju 1150 i 1151 i posle 1153 (EH). Cena Report* po listi i rezimu. | IZV-3, IZV-4, IZV-5, IZV-7 |
| 60 | izvestaji | `IZV.CipVrste` | `modScrIzvestaji.bas:303` | OsveziCipVrste posle ResetCache. | IZV-9 |
| 61 | izvestaji | `IZV.Oblikuj.Petlja` | `modScrIzvestaji.bas:1198` | Petlja 1198-1218 (normalizacija i haystack) po otkucaju, za najvecu listu. | IZV-11 |
| 62 | izvestaji | `IZV.SaldoOMZbirni.Petlja + IZV.SaldoOM.ZbirStavki + IZV.SaldoOM.AmbSaldo` | `modIzvestaj.bas:2107` | Petlja po stanicama 2107-2129 (brojac S). U ReportSaldoOM posebno 520 (ZbirStavkiPoOtkupu) i 702 (AmbSaldoPoNalogu). | IZV-3, IZV-8, UI-11 |
| 63 | izvestaji | `IZV.SaldoKupciZbirni.Petlja + brojac DOK.StavkePrijemniceRedovi` | `modIzvestaj.bas:2215` | Petlja po kupcima 2215-2237. Brojac na pocetku StavkePrijemniceRedovi (modDokumenta.bas:9263); ocekivano K. | IZV-5 |
| 64 | izvestaji | `IZV.AmbZbirno.Petlja` | `modIzvestaj.bas:2429` | Petlja po entitetima 2429-2442, uz brojace AMB.KnjigaZaCitanje i SCHEMA.Knjiga. | IZV-4 |
| 65 | izvestaji | `IZV.RangSort + brojac K` | `modOtkupBlok.bas:179` | SortDescByVal: bubble sort uz broj kooperanata, za period sve. | IZV-7 |
| 66 | zajednicko | `ONTIME.AUTOSAVE (akcija) + AUTOSAVE.Save` | `modJournaling.bas:834` | Pocni izmedju 833 i 834, Zavrsi posle 834. Sekcija oko ThisWorkbook.Save (733). t0_ms prema kraju Sacuvaj pokazuje da li save pada usred kucanja. | OTK-12, ZAJ-03, OTK-15 |
| 67 | zajednicko | `ONTIME.WARM (akcija)` | `modStornoWarm.bas:116` | Pocni posle 115 (On Error GoTo EH), Zavrsi pre 124 i posle 126. Trajanje BuildWarm i razmak u odnosu na ONTIME.AUTOSAVE. | OTK-15, ZAJ-03 |
| 68 | ui-osvezavanje | `LOCK.BulkPush.Petlja` | `modStanicaLock.bas:480` | Citanje tblOtkup celiju po celiju u BulkPushPendingForStanica (sa cloud-om), i da li se desava unutar Sacuvaj (kaskada UI-7). | DOD-2(ui), UI-7 |
| 69 | ui-osvezavanje | `brojac DOK.CiljAktivan` | `modScrDokumenti.bas:378` | Broj poziva po crtanju mreze dok je izabran cilj (posle On Error 378). | UI-10 |
| 70 | zajednicko | `PerfInitIzKonfiga` | `modMain.bas:208` | Ukljucenje iz tblLocalConfig jednom na startu, posle LogAppStart. Prva akcija posle toga je HLADNA (rb_istog = 1). | prekidac, hladno/toplo |
