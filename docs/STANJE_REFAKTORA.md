# Stanje refaktora „dokument = header + stavke“

> **Ulaz za svaku novu sesiju.** Kratko i ažurno — čita se umesto celog plana. Pun plan i istorijat odluka:
> `docs/REFAKTOR_DOKUMENT_HEADER_STAVKE.md` (odluke po datumu u §14.x; važeće: §14.7 „Odluke operatera 16.09“).
> Ažurira se na kraju svakog koraka, u istom commit-u.

**Ažurirano:** 13.10.2026 (red AMB-10: `10e-3` merge-ovan `#418`, stavka 130; `KI-011` merge-ovan `#419`, stavka 132; `KI-012`/`KI-013` merge-ovani `#420`, stavka 134; čitaoci salda na papiru od stavke 135).

## Pravila koja važe (16.09.2026)

- **Nema produkcije ni podataka koje treba štititi — radi se iznova.** Bez migracija, backfill-a, fallback-a.
- **Legacy kod ne mora da radi između faza.** Ne gradi se ništa što ga čuva živim: pauze, podela kapija, test-only pisci,
  kolone ostavljene za pauzirane čitaoce, mostovi.
- **Apsolutno se čuva samo mapa sposobnosti** — sve što operater danas može da uradi ili dobije.
- **Jedino tehničko ograničenje:** posle svakog PR-a projekat se kompajlira (inače pada ceo `run_vba`). Legacy se briše,
  ne ostavlja polomljen.
- **VBA arhitektura vodi, PWA prati** (23.09.2026). Cilj je idealan VBA model; PWA/GAS se prilagođava njemu
  kasnije. Ne prave se adapteri ni kolone koje postoje samo da bi zatečen PWA payload radio — šta PWA mora da
  šalje zapisuje se kao nizvodni zahtev, a sposobnost koja zbog toga privremeno ne radi se kaže **glasno**.
- **Jedna sesija = jedan korak.** Masovne mehaničke provere se rade spolja, po promptu.

## Gde smo

| Stavka | Status |
|---|---|
| PR0–PR6 (ugovor, šema iz koda, skele Zbirna/Otkup/Otpremnica, Otkup cutover) | ✅ |
| Kapija odluke §14.1 (nastavak u istom repou) | ✅ |
| #333 pre-flight PR7 · #334 čitaoci vrednosti otkupa na stavke | ✅ |
| #335 alat `tools/popis_citalaca.py` + odluke 16.09 | ✅ |
| **Mapa sposobnosti** | ✅ spojena u `docs/DOMEN/MAPA_SPOSOBNOSTI.md` (387 sposobnosti; ulazi A–F ostaju u `docs/DOMEN/mapa_sposobnosti_ulazi/`) |
| Odluke domena | ✅ §14.8 (17.09.2026) |
| Nova tabela slajsova | ✅ §14.9 (17.09.2026) |
| S1–S4 (otkup, banka, otpremnica, zbirna) | ✅ |
| **S3-ostatak** (mrtve linijske kolone zaglavlja otpremnice) + putanja rename-a kolone + KI-008 | ✅ #395 (`8eeca04c`) |
| **S5 (PWA i sync na novom modelu)** | ✅ zatvoren kroz #385–#394; ostatak je jedno mesto `DEGRADIRANO` grane (v. „Sledeće“) |
| **AMB-10 ambalaza kao knjiga prenosa** | ⏳ **ide PRED S6** -- model ✅ · `10a` ugovor ✅ (#398) · `10-DOK` zaglavlje ✅ (#399) · `10b-1` pisac ✅ (#400) · `10b-2` cutover pisaca ✅ (#408) · `10c` čitaoci ✅ (#408, #409, #411; `KI-010` čeka ručni smoke papira) · `10d` ✅ (#414 ispravka reversa jednim potezom, stavke 117–120; #415 undo i istorija `ODL-25/-26/-27`, stavke 121–124) · **`10e` brisanje starog modela ⏳** (stavka 125; `10e-1` čitaoci ✅ merge-ovano `#416` (ručna provera ✅ 13.10.2026, stavka 126) · `10e-2` stari storno, undo i ispravka reversa ✅ merge-ovano `#417` (stavke 127–128) · `10e-3` pisac, dvomodelni čitalac, šema i fixture ✅ merge-ovano `#418` (stavke 129–130); `KI-011` ✅ `#419`; `KI-012`/`KI-013` ✅ `#420` · **čitaoci salda na papiru ⏳** (stavka 135); `docs/DOMEN/AMBALAZA.md` |
| **S6 prijemnica** | ⏸ **parkiran na koraku 1/8** (grana `claude/s6-prijemnica-stavke`) -- nastavlja se posle AMB-10 |
| S7 faktura · S8 palete · S9 sledljivost kao graf | ⏳ |
| **Vraćanje `otk_linija` na nulu** (18 živih čitalaca) | ⏳ — to je ono što još drži linijska polja `tblOtkup` na životu |

## Hronologija rezova — i gde je sledeći (v. stavku 50)

1. Mapa: `docs/DOMEN/MAPA_SPOSOBNOSTI.md`. Odluke: plan §14.8. Slajsovi: §14.9. Pre-flight, S1a, S1b: §14.10.
2. **S1b-1 spojen** (#354). **S1b-2 urađen** (§14.10 „S1b-2 — urađeno“): desktop čitaoci otkupa na stavkama, stari panel
   `modOtkupBlok` i `clsBlokUI` obrisani, AUD-056 zatvoren; `x_otk_stavka` 63 → 22. Spojen (#355).
3. **S1b-3 urađen** (review #355, §14.10 „S1b-3“): bez mosta preko `Otkup.OtpremnicaID` — F1 radni sto otpremnice,
   bilans po otpremnici i ekran SLEDLJIVOST obrisani (vraćaju S3/S9); prefill ispravke fail-closed. Spojen (#356).
4. **Pravilo:** novi model se ne čita kroz staru vezu — takva sposobnost se briše, ne prevodi.
5. **S1c urađen** (§14.10 „S1c“): izvozi `OtkupPoOM`/`OtkupiAll`(+`OtkupiAllStavke`)/`SaldoOMDetail` i push ka stanici
   (`OTK_STAVKE`) čitaju kanonske stavke; raspored OTK kolona na jednom mestu; `AutoCreateOtpremniceFromPWA` obrisan;
   health iz kanona (AUD-055); `x_otk_stavka` = 0; push `OTK_STAVKE` idempotentan po `OtkupStavkaID`. Spojen (#357).
6. **S1d urađen** (§14.10 „S1d“): 8 kolona obrisano iz `tblOtkup` u kanonu i konstante `COL_OTK_*` iz `modConfig`;
   zatečena sveska ih gubi kroz self-heal (`ObrisiKolonuAko`). Spojen (#358).
7. **S1e urađen** (review #358, §14.10 „S1e“): identitet otkupa UI → `OtkupID` → mutacija — F1 red nosi `OtkupID`,
   štampa/storno/F8/hladnjača po ID-u; `StornoOtkupByBrDok_TX`, `OtkupIdsByBrDok`, `StornoSelectedBlocks_TX` obrisani;
   testovi sa dva zaglavlja po klasi obrisani. Spojen (#359). **S1 završen.**
8. **S2 urađen** (§14.11): banka radi po `OtkupID`-u — lista blokova nosi ID, poziv na broj se razreši JEDNOM
   (dvosmislen = ručno, ne raspodela preko dva otkupna mesta), pisač prima ID i dobija kapiju vlasništva i storna;
   `GetOtkupCandidatesForKooperantBlock`, `PlanBlokRaspodela` i ceo scope otkupnog mesta obrisani. Pre merge-a:
   pun `run_vba` + Compile.
9. **S2 spojen** (#360, `5610bfc9`) posle dva review kruga: vezan red novca nosi otkupno mesto **dokumenta**, a višak
   preko duga sme u avans samo uz **saglasnost pozivaoca** (`dozvoliVisakKaoAvans`, default False).
10. **S3 rez na pet koraka** (§14.12): živa ulazna tačka starog pisca otpremnice bila je **jedna**
    (`modDokUnos.OtpremnicaUpisi`), sve ostalo su čitaoci i pauzirani putevi.
11. **S3a urađen** (§14.12): F2 otvara **nacrt** (`CreateOtpremnicaDraft_TX`); `PredlogCena` je po klasi na stavci, a
    zaglavlje je više ne prima (ključ se odbija); **ambalaža se knjiži pri izdavanju**, ne na nacrtu; malina
    auto-zbirna pauzirana uz poruku (vraća S4). Stari pisac (`SaveOtpremnica*`) ostao bez živog pozivaoca —
    briše se u S3e. Pre merge-a: pun `run_vba` + Compile.
    **Review #361 P1:** uklonjen i `ZavrsiIspravkuAko FLOW_DOC_OTPREMNICA` — ispravka otpremnice (B-038) je
    **pauzirana**, ne prevedena: stari tok je preko `ReassignOtkupToOtpremnica_TX` pisao `Otkup.OtpremnicaID` i
    proglašavao nacrt bez izvora zamenom izdate otpremnice. Vraća S3c, po ID-u.
    **Pravilo od S3a:** nijedan nov kod ni test ne sme da zove `SaveOtpremnica*`.
12. **S3b rez na dva** (§14.13): „čitaoci“ i „nov ekran sa tri pisca“ su dve vrste rizika, a panel uz to traži
    `OtpremnicaID` u nevidljivoj koloni reda — kolonu koju `modScrStorno` čita očekujući `GeneracijaID`.
13. **S3b-1 urađen** (§14.13): kanonski čitaoci stavki otpremnice (`StavkeOtpremniceRedovi`,
    `ZbirStavkiPoOtpremnici`, `StavkeOtpremnicePoDokumentu`) i svi živi čitaoci na njih — F2 mreža, štampa
    (jedan red = jedna stavka), izveštaji (po otkupnom mestu **red po klasi**), invarijanta zbirne i provera pre
    unosa, audit B4a/B5b, F8 lista i uvid, prefill ispravke. Zaglavlje bez stavki **pada po imenu**; audit i
    lista za storno čitaju **meko** (postoje da nabroje pokvarene redove). Backfill prijemnica hladnjače
    PAUZIRAN (vraća S3d). Merenje: `otp_cena` živih 5 → **0**, `otp_linija` živih 59 → **24** (činjenice
    zaglavlja + `modSetup`).
    **Kapija:** `python tools/popis_citalaca.py --check` — prag živih mesta po grupi; pravilo „nijedan nov kod ne
    zove `SaveOtpremnica*`“ više nije rečenica nego exit kod (review #361 P2). Dokazano u oba smera.
    **Fixture se MORA regenerisati:** zatečeni ima 28 otpremnica i nijednu stavku; `make_fixture.py` sada izvodi
    `tblOtpremnicaStavke` iz `tblOtpremnica` (jedna stavka po zaglavlju), kao što od S1 izvodi `tblOtkupStavke`.
14. **S3b-1 proširen** (§14.14, odluka 19.09.2026): prvi pun prolaz je pokazao da stari pisac i dalje pravi
    otpremnice bez stavki (u testovima), i jedan propust iz S3a: **od #361 nijedan živi put ne vezuje otpremnicu
    za zbirnu**, pa F3 ne može da sačuva nijednu zbirnu, a F4 nijednu prijemnicu. Urađeno: `SaveOtpremnica*`
    **obrisan**, auto-lanac hladnjače obrisan, **F3 glasno pauziran do S4, F4 do S6**, `SumOtpremniceByKlasa` više ne
    guta grešku (upisivala je 0 kg u zbirnu). Testovi starog lanca zbirne su **obrisani, ne prepravljeni ručnom
    vezom**, a njihov spisak je u §14.14 kao lista koju S4 mora da vrati. `otp_stari_pisac` = **0**.
    `RunAllTests` 200 (bilo 203), golden 2 scenarija (bilo 12).
15. **Review #362** (§14.15): `PredlogCena` više nije vrednost — vrednost otpremnice je vrednost njenih
    **izvora**; operativni izveštaji i štampa broje **samo IZDATO**; F4 je **S6**, ne S4. Fixture izvodi
    `tblOtpremnicaIzvori` iz `Otkup.OtpremnicaID` (17 izdatih) — **regeneracija fixture-a**. Drugi krug:
    `PROSLEDJENO` je izdato (sync ne sme da briše otpremljenu robu), prazan/nepoznat status nije; čitač
    stavki odbija dve stavke iste klase.
    Treći krug: F8 stornira otpremnicu **po `OtpremnicaID`-u** (isti rez kao S1e za otkup); otpremnica nije
    više framework tip, a modovi ISPRAVKA/DUPLI/PONIŠTENJE za nju su **PAUZIRANI do S3c/S4** (B-022..B-025).
    Četvrti krug: pisac (`StornoOtpremnica`) **odbija storno izvora aktivne kanonske zbirne** (A13/A15),
    F8 kaže razlog pre potvrde; kaskada se ne pravi -- odlučuje S4.
    **S3b-1 spojen** (#362, `a8d3bd1a`).
16. **S3b-2a urađen** (§14.16): kapija storna otkupa u sastavu aktivne otpremnice (P1 review #362, jezgro
    `StornoOtkup`); radni sto u F1 nad kanonom — liste OTPREMNICE/BLOKOVI sa ID-em u redu, aktivna otpremnica je
    uvek NACRT, traka sa semaforom po klasi, prekoračenje po klasi, vezivanje posle unosa, veži/ukloni/izdaj;
    izmena nacrta u F2 (odluka: povezano ≠ očekivano se rešava izmenom, ne izjednačavanjem). Prag
    `otp_linija` 24 → 36 (sve činjenice zaglavlja). Pre merge-a: pun `run_vba` + Compile.
    Review #363, prvi krug: read-model i izdavanje čitaju stavke kroz **stroge kanonske čitače** (dve
    stavke iste klase više ne postaju IZDATO); izmena nacrta se ne otvara nad delimičnom formom. Pre S5:
    otkup u `PROSLEDJENO` mora da bude prihvaćen kao izvor (backlog §15).
    Review #363, drugi krug: `StavkeOtkupaRedovi` drži ugovor pisca otkupa (jedna stavka po klasi, bruto
    ≥ neto), pa ni pokvaren izvor ne postaje IZDATO. P2 granica agregata komandi → backlog §15.
17. **S3b-2a spojen** (#363, `0ceccea1`) posle dva review kruga.
18. **S3b-2b urađen** (§14.17): specifikacija otkupnih blokova nad kanonom — red je **stavka izvora** (blok × klasa,
    svaka sa svojom cenom), štampa se **samo izdata** otpremnica, a oznake više ne idu preko broja nego preko
    `OtpremnicaID`-a u nevidljivoj koloni (stari `OtpIdZaBroj` se ne vraća). Vraćen opseg datuma OD/DO iznad liste
    otpremnica (ostatak A-022) i „Po datumu“ koja štampa tačno ono što je u listi.
    **Odluka operatera:** A-025 je lista **NEVEZANIH** blokova (upisan bez otpremnice, uklonjen iz nacrta,
    oslobođen stornom), a ne samo „izgubljenih“; kolona „bila u“ nosi broj stornirane otpremnice. Iz te liste ide
    `veži` za aktivni nacrt. Zbirna i kupac na specifikaciji su prazni dok S4 ne vrati F3.
    Vrsta „izgubljen blok“ na ekranu OPORAVAK (B-041) ide u **S3c**, uz storno otpremnice.
    **Review #364, prvi krug:** bulk čitač članstva (`AktivnoOtpClanstvoPoKanonu`) drži **isti ugovor** kao čitač
    jednog dokumenta — članstvo bez ID-a, roditelj ili dete koje ne postoji tačno jednom, dupli par i dva aktivna
    članstva su tvrde greške. Bez toga je članstvo na nepostojeći otkup davalo uredan PDF bez tog izvora (validan
    sibling je zadovoljavao kapiju „izdata bez izvora“), a članstvo na nepostojeću otpremnicu sklanjalo slobodan
    blok sa liste NEVEZANI. P2 (strog čitač zaglavlja otkupa za štampu) → backlog §15.
19. **S3c urađen** (§14.18): **ispravka izdate otpremnice je jedan potez i jedna transakcija** — storno stare +
    nov NACRT koji nasleđuje zaglavlje, očekivanje i sve izvore; nacrt se i dalje samo menja (F2), a broj se ne
    nasleđuje (A9). Trag ide po identitetu: `tblOtpremnica` dobija `IspravkaOdID`/`ZamenjenSaID`.
    **Odluka operatera:** DUPLI, PONIŠTENJE i REŠI KASNIJE za otpremnicu se **brišu** — DUPLI je u kanonu isto što
    i običan storno, a druga dva nemaju o čemu da odluče dok F3 (S4) i F4 (S6) ne postoje.
    Obrisan ceo okvir modova za otpremnicu (~800 linija, uključujući `StornoOtpremnicaByBroj_TX` i grane uvida).
    **Nalaz:** kapija `BlockStornoDriftReason` je čitala mrtvu vezu `Otkup.OtpremnicaID`, pa od S3a nikad nije
    odbijala; sada čita kanon i fail-closed je. B-039 (MANUAL zadaci) je `INTENTIONALLY REMOVED` — prozora između
    storna i zamene više nema. `RunAllTests` 201 → **197** (obrisana četiri testa starog okvira), sabotaža **511**.
20. **S3c-2 urađen** (§14.19): vrsta „izgubljen blok“ u listi `Nedovršeno` (B-041) i brojač uz meni (B-042)
    vraćeni su **nad kanonskim članstvom**. „Izgubljen“ je samo blok koji je **bio** u otpremnici pa ga je njen
    storno oslobodio — blok upisan bez otpremnice je normalno stanje i čeka na radnom stolu (F1). Pad strogog
    čitača daje **vidljiv red sa greškom**, ne tiho kraću listu. Radnja nad redom je pokazivač na F1; vezivanje
    ostaje kod kanonskog pisca. **S3c zatvoren.**
21. **S3d-1 urađen** (§14.20), posle **NO-GO review-a #367 i prepravke**: kanonski korak otpremnice postoji, ali
    je **aktivacija vezana za S6**. Lanac radi samo za OM podešeno kao **hladnjača**, gde je vozač-ogledalo
    **uvek obavezan** (kooperant sam dovozi robu), i pravi otpremnicu **1:1 po klasi** iz jednog bloka.
    **Automatika je obavezna:** ekran grana **pre** ručnog vezivanja — hladnjački blok ne ide na radni sto, jer bi
    postao član tuđeg nacrta sa drugom kilažom. Provera članstva ostaje kao zaštita od dupliranja.
    **Jedan autoritet nad aktivacijom:** `AUTO_PRIJEMNICA_HLADNJACA` (do sada prekidač koji niko nije čitao) sada
    odlučuje DA LI lanac radi; default **OFF do S6**, jer se zbirna i prijemnica danas ne mogu doraditi ni ručno.
22. **S3d-2 urađen** (§14.21): **A13 kapija je otvorena za NACRT** — ispravka bloka koji je u nacrtu radi
    **atomsku zamenu članstva** u istoj transakciji (`modDokumenta.ZameniOtpremnicaIzvor`, core unutar tuđe
    transakcije); blok **izdate** otpremnice ostaje odbijen, ali poruka sada imenuje put koji od S3c postoji
    (ispravi otpremnicu → nastaje nacrt → ispravi blok). Kapije su u **jednom izvoru**
    (`modOtkup.IspravkaOtkupaRazlog`): ekran ih pita pre forme, pisac ih diže kao grešku.
    **B-040 dobija ekran** — radnja „Ispravi“ nad redom u listama SVI i BLOKOVI; pisac je od S1 bio bez ijednog
    živog pozivaoca. **S3d zatvoren.**
23. **S3e-1 urađen** (§14.22): **merenje je oborilo pretpostavku koraka** — poslednji živi čitaoci kolona koje
    S3e treba da obriše su kaskada **zbirne** (S4) i **OTK list za PWA** (S5). Odluka operatera: S3e se deli.
    Sada je obrisan samo kod bez ijednog živog pozivaoca (`ReassignOtkupToOtpremnica_TX`,
    `CalculateManjakByOtpremnica`, `BackfillOtkupBrojOtpremnice`, ceo `modSledljivost.bas`), a grupe u popisu su
    **podeljene** da prag meri nešto što sme na nulu: `otk_veza_otp` (13) i `otp_linija` (3) idu na nulu u S3e-2,
    `otk_brojzbirne` umire sa S4, a `otp_zaglavlje` namerno **nema prag**. `otp_cena` je dostigla **0**.
    Usput očišćen `WRITE_OWNERSHIP.json` (tri modula koja `tblOtkup` više ne pišu).
24. **S4 rez na četiri koraka** (§14.23): merenje je pokazalo da kanonski pisac zbirne postoji od PR3, ali
    **nijedan kanonski čitalac** — a pisac linijska polja zaglavlja namerno ostavlja prazna. Zato bi „F3 prvo"
    dalo zbirnu koja u bazi postoji a na ekranu je prazna. Redosled: **S4-1 čitaoci → S4-2 F3 → S4-3 okvir
    storna/ispravke → S4-4 malina auto-zbirna**.
25. **S4-1 urađen** (§14.23): sadržaj zbirne (kilaža, gajbe, klasa) čita se iz `tblZbirnaStavke` kroz strog
    kanonski čitalac koji ceo registar odbija po imenu; preneti su **svi živi čitaoci sadržaja** — liste F8,
    ciljna lista Oporavka, uvid i prefill pred storno, izveštaj po vozaču, integritet B7. Fixture izvodi
    `tblZbirnaStavke` iz `tblZbirna` — **regeneracija fixture-a**. Ekranski adapter F3 preimenovan u
    `SnimiZbirnu`. Nove grupe popisa: `zbr_linija` 30, `zbr_stari_pisac` 29.
    **Odluka operatera (ZBR-KANON-03):** izmena izvora ne prepravlja zbirnu u mestu nego pravi **novu verziju**
    (A13); `docs/DOMEN/README.md` je tvrdio suprotno i ispravljen je. Okvir rekalkulacije briše S4-3.
26. **KAPIJA ZA S4-2 (review #370, P1):** identitet zbirne u ljusci je još `GeneracijaID`, koju kanonski pisac
    **ne upisuje** — a `Chk_B9` istovremeno tvrdi da je prazna generacija integritetska greška. Dok je F3
    pauziran to ništa ne laže; počinje da laže **u trenutku kad F3 proradi**, jer bi ljuska pala nazad na
    `BrojZbirne` kao identitet. **S4-2 počinje** prelaskom `IdKolonaTipa("ZBIRNA")` na `COL_ZBR_ID` i
    usklađivanjem B9, pa tek onda skida pauzu. Rešenje NIJE dodati `GeneracijaID` kanonskom piscu.
27. **S4-2a urađen** (§14.24): **identitet zbirne je `ZbirnaID`** — nevidljiva kolona u F8, jezgro storna
    (`StornoZbirna(zbirnaID)`), uvid pred storno i `Chk_B9` (sada „bez ID-a ili sa duplim"). Stari okvir
    (`modStornoFlow`) sam prevodi (broj, generacija) → ID kroz `ZbrIdIliGreska`, **fail-closed**, pa jezgro ne
    poznaje stari model. Kapija `RequireJedanVlasnikPoBroju` je obrisana jer je štitila izbor po broju, kog
    više nema. Rešenje NIJE bilo dodati generaciju kanonskom piscu (dva identiteta).
    **Rez u dva PR-a:** aparatura generacije ima 22 reference samo u `modDokumenta`, a `ZBR-CHILD-01` je veže
    za decu (prijemnica, paleta) koja ostaju do S6 — pa se uklanja identitet zbirne, ne ceo mehanizam.
28. **S4-2b urađen** (§14.25): **zbirna dobija NACRT.** Odluka operatera: mora da radi na oba načina kao
    otpremnica — najava pa pokrivanje izdatim otpremnicama, ali i direktan izbor izvora — a pregled svih
    zbirnih u F3 je uslov bez kog se ne može. Ovaj PR je **pisac**: `CreateZbirnaDraft_TX`,
    `DodajZbirnaIzvor_TX`/`UkloniZbirnaIzvor_TX`, `IzdajZbirnu_TX` (najava = povezano po klasi, uz
    revalidaciju izvora), `ZbirnaJeIzdata`, `ZbrClanovi`. **Izvor zbirne je IZDATA otpremnica** (odluka koju
    je §14.14 ostavila S4); `PROSLEDJENO` se računa kao izdato. Zbirna **ne knjiži ambalažu** — gajbe su
    knjižene na otpremnici i knjiže se ponovo na prijemu (S6).
29. **S4-2c/1 — ulazna kapija nacrta (PR #___).** Kapija koju je review #372 postavio PRED ekrane:
    `UpdateZbirnaDraft_TX` (izmena najave, članstvo netaknuto; nacrt zadržava svoj broj a ne preuzima
    tuđi; izmena zaglavlja **revalidira** postojeće članstvo) + odluka operatera **ZBR-KANON-04** —
    kad članstvo padne na nulu, izvedeni `Vrsta/Sorta/TipAmbalaze` se **brišu**, a dok ima bar jednog
    člana **ostaju**. Review #373 (P1): tvrdnja „nacrt ima bar jednu klasu" preseljena iz wrapper-a u
    jezgro `ZbrUpisiOcekivano` — prazna kolekcija je kroz `Update` pravila zaglavlje bez stavki, koje
    strog čitalac odbija. Nijedna linija ekrana. Detalji: plan §14.26.
30. **S4-2c/2a — stari pisac zbirne obrisan (PR #___).** `SaveZbirna*`, `BuildZbirnaRowData`,
    `ValidateZbirnaInput`, `modDokUnos.ZbirnaUpisi`, `AutoCreateZbirnaFromOtpremnice` +
    `BackfillOtkupBrojZbirneByOtpremnica`. Oba produkciona pozivaoca bila su iza pauze (izmereno).
    Ulaz malina auto-zbirne ostaje i pada **glasno**; telo se vraća u S4-4. Testovi: 4 tvrdnje o
    piscu prešle na kanonski nacrt, 7 + golden na fixture `ZbrZateceniRed` (zatečeni oblik, umire sa
    S3e-2); golden fajl **nepromenjen**. Popis: `zbr_stari_pisac` 29 → **0**, `zbr_linija` 30 → 27,
    `otk_brojzbirne` 27 → 25. Detalji: plan §14.27.
31. **S4-2c/2b-1 — jezgro i unos za ekrane (PR #___).** Merenje je oborilo premisu: radni sto nacrta
    je u **F1**, ne u F2 — pa radni sto zbirne pripada **F2**. Zato je 2b isečen po sloju: 2b-1 =
    `GetZbirnaProgress` + `NevezaneOtpremnice` (**samo IZDATE i slobodne**) + `ZbirnaUpisi` /
    `ZbirnaIzmeniNacrt` + **prepisan** `ZbirnaValidiraj` (bez poređenja po `BrojZbirne`, bez vrste i
    sorte, bez `GeneracijaID` kapija; broj sudi istim alatom kao pisac). Pauza je sada na tačno
    jednom mestu (`SnimiZbirnu`). Nijedna linija ekrana. Detalji: plan §14.28.
32. **S4-2c/2b-2a — F3 piše (PR #___).** Pauza skinuta; mreža F3 nosi nevidljivi `ZbirnaID`, klik na
    red otvara izmenu nacrta po ID-u, snimanje pravi **ili** menja nacrt. Izdata se ne otvara.
    Ekran ništa ne sudi. Ostaje za 2b-2b: polja vrste/sorte/tipa ambalaže u formi primaju unos koji
    se nigde ne upisuje. Detalji: plan §14.29.
33. **S4-2c/2b-2b — radni sto zbirne u F2 (PR #377).** Liste `ZBIRNE`/`IZVORI`/`NEVEZANE`, izbor
    aktivnog nacrta klikom (po `ZbirnaID`), veži/ukloni/izdaj kroz kanonske pisce. Nema novih mreža —
    `RedoviZaSkup` sužava `RedoviZaTip`. Zbirna je prvi put **ceo tok**: najava → pokrivanje →
    izdavanje. Detalji: plan §14.30.
34. **S4-2c/2b-2c-1 — F3 ne traži ono što ne nosi (PR #379).** Cena, tip ambalaže i vrednost skinuti
    sa F3 kroz `FldShow`; lista `SVI` više ne nudi `Veži` (P3 #377). Drugi P3 (`mZbrID`) se **ne**
    zatvara brisanjem stanja — kontekst mora da preživi F2↔F3, rešava ga traka. Detalji: plan §14.31.
35. **S4-2c/2b-2c-2 — traka napretka zbirne (PR #381).** Traka se crta i u F2, nad aktivnim nacrtom;
    čita `GetZbirnaProgress` — isti par čitača koji koristi izdavanje. **Četvrta mera je BROJ IZVORA**
    (odluka operatera): zbirna nema cenu, a pokrivenost je pitanje članstva. Natpise šalje ekran
    (14. polje ugovora), pa ljuska ne pogađa šta je u kom režimu predmet rada. Detalji: plan §14.32.
36. **S4-3a — ispravka zbirne umire, ne seli se (PR #382).** Okvir rekalkulacije u mestu i
    dvokoraka „storno sada, zamena kasnije" je obrisan (ZBR-KANON-03). **Zamena se ne gradi sada:**
    jedan potez traži prenos prijemnica, a one vise na `BrojZbirne` do S6 — pa bi nova zbirna
    (nov broj, A9) ostavila siročad. Odluka operatera: **ispravka zbirne čeka S6**; do tada F8 nudi
    `DUPLI` i `PONIŠTENJE`. Usput nađeno: `ValidateZbirnaInvariant` je pod kanonom bila **vakuumska**
    (obe strane nule → uvek `OK`), i u uvidu i u golden snimku. Detalji: plan §14.33.
37. **S4-4 — malina auto-zbirna nad kanonom (PR #383).** Sposobnost vraćena: jedno jezgro
    (`AutoZbirnaZaOtpremnicu`), dva pozivaoca (izdavanje otpremnice + batch iz sync-a), okidač
    pomeren sa **nacrta** na **izdavanje**, članstvo kanonsko, zbirna dobija **svoj** broj — čime je
    zatvoren dug koji je stari komentar ostavio baš ovom rezu. Kapija lanca razdvojena: VOZ/zbirna
    uvoz ostaje pauziran. **Hladnjački lanac namerno nije dirnut** — njegov ZBR korak okida A/B
    odluku o nastavljivosti, koja je izlazni uslov S6. Detalji: plan §14.34.
38. **S4-3b — identitet zbirne je ZbirnaID (PR #___).** Ušlo kao čišćenje, ispalo **živ kvar**:
    `ZbirnaIdentResolve` je tražio `GeneracijaID` koji više ne piše nijedan pisac (S4-3a je obrisao
    poslednjeg), pa su `DUPLI` i `PONIŠTENJE` bili nedostupni za **svaku** kanonsku zbirnu — baš za
    ono što je S4-3a ponudio kao zamenu. Nijedan test to nije video jer svi mere fixture redove koji
    generaciju nose. Detalji: plan §14.35.
39. **S4-3c — osa dovršena (PR #384).** Trag na detetu nosi **identitet** roditelja, kolona se zove
    `ZbirnaRoditeljID`, scoping ide po ID-u, test-pečati generacije i backfill obrisani. Usput:
    storno suite je dobila izveštaj **po imenu** (bez toga se 9 padova nije moglo trijazirati).
    Detalji: plan §14.36.
40. **S5-1 — auto-otpremnica iz PWA otkupa nad kanonom.** Korak 3 ciklusa vraćen, korak 2b
    (`VozacID := StanicaID` na otkupu) **obrisan**: vozač je činjenica zaglavlja OTPREMNICE.
    **Klasa više nije ključ grupisanja** — dva bloka I i II klase istog dana sa istog otkupnog
    mesta daju JEDAN dokument sa dve stavke. Pad jedne grupe ne obara ostale, ali se imenuje.
    Detalji: plan §14.37.
41. **S5-2 — predaja robe vozaču postaje otpremnica.** Redosled u planu ispravljen po rečenici
    operatera: lanac je **predaja → otpremnica → zbirna**, pa VOZ/zbirna ide POSLE predaje.
    `TryUpdateVozacID` obrisan — bio je poslednji pisac `Otkup.VozacID`, pa `modMasterSync`
    više ne piše `tblOtkup` (4 → 3 pisca). Jedan utovar = jedan dokument, iako stiže kao N
    redova. Detalji: plan §14.38.
42. **S5-3 — VOZ/zbirna uvoz nad kanonskim piscem.** Goli `Array(...)` od 16 vrednosti sa
    količinom i klasom na zaglavlju je otišao; članstvo se razrešava `otkupRecordIDs → OtkupID
    → OtpremnicaZaOtkup`. `LinkZbirnaToOtkupAndOtpremnica` (224) i `IzvedeniLanacIzPwaDostupan`
    (38) obrisani — popis: `pauza` **6 → 0**, „kapije: nema“. `modMasterSync` više ne piše
    `tblZbirna` (3 → 2 pisca). `GeneracijaID` za **zbirnu** nema pisca; za **prijemnicu** ostaje
    do S6. Detalji: plan §14.39.
43. **S5-3b — storno tok zbirne nad kanonskim članstvom (#389).** Tok je bio ostao bez hrane:
    mrtav most preko starog backlinka. Detalji: plan §14.40–14.41.
44. **S5-4a — predaja je SOPSTVEN događaj (#390).** Store `predaje` → `PRED-*` list (append-only)
    → `ImportOnePREDSheet`. Retry se prepoznaje po **utovaru**, ne po vozaču. Predaja više ne živi
    kao tri kolone na OTK redu: cim red dobije `Synced>Master`, uvoz ga ne čita, pa je predaja koja
    stigne kasnije tiho nestajala. Detalji: plan §14.42–14.43.
45. **S5-4b-1 — žica vozača (#391).** Master izvozi otpremnice (zaglavlje + stavke) u `MgmtReports`;
    GAS servira vozaču otpremnice po `Otpremnica.VozacID`. Presedan **mrtvog slota**:
    `VS_OTKUP_RECORD_IDS` ostaje na žici i piše se prazan, jer `ensureSheetColumns` dozidjuje samo na
    kraju a svaka druga razlika je `SCHEMA_DRIFT`.
46. **S5-4b-2 — ekran vozača (#392).** `zbirna.js` / `transport.js` rade nad otpremnicama; zbirna
    šalje `ZbirnaID` + spisak `OtpremnicaID`. `OtpremniceIzOtkupRecordIDs` **obrisan** (mereno: 0).
47. **S5-5a — JS kapija (#393).** `tests/js/` harness izvršava produkcijske fajlove kroz `vm.Script`,
    sabotaža se primenjuje na **tekst u memoriji** (ne na radno stablo). **Četiri kruga review-a, i
    nijedan nalaz u produkcionom kodu** — svi su bili u instrumentu. Zapisano kao pravilo: u rezu koji
    uvodi merenje greška je u sredstvu merenja, i odbrana napisana pre merenja postaje nalaz.
48. **S5-5b — žica otkupa: zaglavlje + stavke (#394, `561fcad9`).** Sva tri sloja: VBA uvoz čita
    `OTK_STAVKE`, GAS piše **stavke pa zaglavlje**, PWA nosi `stavke[]` (odluka: **samo žica, N=1** —
    forma i dalje unosi jednu klasu). Zaglavlje nosi `StavkeCount` kao **manifest**, pa se
    „nedostaje stavka“ razlikuje od „dokument ima jednu stavku“. Četiri linijske kolone su **mrtvi
    slotovi**. Detalji: plan §14.44–14.50.
    **Nalaz koji plan nije predvideo:** tab `OTK_STAVKE` od tada ima **dva pisca**, a PWA red ne zna
    `OtkupStavkaID` — indeks idempotencije push-a ga je čitao kao „red bez identiteta“ i trajno
    blokirao push te stanice, fail-closed nad ispravnim podatkom.
    **Tri kruga review-a, svi na istoj granici — identitetu:** hibridni dokument od partial upisa ·
    completion marker (zaglavlje postoji ⇒ skup je nepromenljiv) · identitet obuhvata **ceo** payload,
    ne samo skup stavki.
    **CI je dao četiri nalaza i sva četiri u instrumentu:** `deepStrictEqual` preko granice `vm`
    realm-a, placebo tvrdnja koja je imenovala jedno a merila drugo, i moja statička provera sidara
    koja je tiho preskakala jedan unos.
49. **Tada predloženo kao sledeće — S1 „Otkup do kraja“; POBIJENO merenjem u #395, v. stavku 50.** Mereno posle #394: `otk_linija` = **18 živih PROD čitalaca**
    (`modMasterSync` 3, `modOtkup` 2, `modPrint` 2, `modSetup` 2, `modStammdatenSync` 2, `modStornoDok` 2,
    `modAutoHladnjaca` 1, ostatak raspoređen). Dok oni ne padnu, otkup je „header + stavke“ **na žici**
    a dvojan **u jezgru**. Rez dira štampu, izveštaje, izvoze, storno prefill i KPI — a štampa i PDF su
    ono što testovi ne mere, pa ide **po čitaocu**, ne po fajlu.
    **Posle S1:** ostatak S3/S4 (`otp_brojzbirne` 15, `otk_veza_otp` 14, `otk_brojzbirne` 8,
    `zbr_linija` 22, `x_trace` 14, `x_literal` 7, `split_plus` 5), pa S6 (prijemnica, F4), S7, S8, S9.
    **Ostatak S4-2c:** vrsta/sorta iz kontekstne zone F3 (traži raspored ljuske) i sužavanje
    liste `NEVEZANE` na aktivan nacrt — oba u backlogu §15.

50. **#395 — tri reza u jednom PR-u, i pre-flight koji je odbio rez** (`8eeca04c`, 27.09.2026).
    **S1 ne postoji kao rez.** Pre-flight je izmerio `x_otk_stavka` = **0**: nijedan čitalac ne
    čita linijska polja `tblOtkup`, jer tih kolona u kanonu **nema od S1d**. `otk_linija` = 18
    **nije** mera S1 nego nosi `TIP_AMB`/`KOL_AMB_IZDATA`, činjenice **zaglavlja**
    ([popis_citalaca.py:102](../tools/popis_citalaca.py)) — moja preporuka u stavci 49 stajala je
    na pogrešno pročitanom brojaču, i to je ispravka premise, ne promena plana.
    **S3-ostatak urađen:** `tblOtpremnica` 29 → 24 kolone (`Kolicina`, `Cena`, `KolAmbalaze`,
    `Klasa`, `BrutoKg` iz **sredine**, sa `Stornirano` između njih), `otp_linija` **3 → 0**.
    **Putanja rename-a kolone:** `ZbirnaGeneracijaID → ZbirnaRoditeljID` self-heal nije imao
    putanju, pa se zatečena sveska **nije mogla izlečiti** (513 padova nad pravim donorom).
    **KI-008 zatvoren:** `OTKUP_BRUTO_UNOS` nije bio pinovan u `make_fixture`, pa se nasleđivao iz
    donora — jedan uzrok, pet simptoma, četvrti put ista klasa.
    **Post-merge review P2 (isti PR, follow-up na `main`):** oporavak dvostrukog imena je čuvao
    **podatke** ali ne i **poziciju** — obriši staro iz sredine, ostavi novo na kraju, i
    `VerifySchema` (kanon = **prefiks po indeksu**) i dalje odbija svesku. Mereno: pogodi **četiri**
    od šest poziva helpera, a dva najizloženija (`IspravkaOdID` 22/29, `ZamenjenSaID` 23/29)
    review nije imenovao. Zatvoreno: preživljava **levlja** pozicija (višak je uvek dopisan),
    sadržaj se preseli u nju, postcondition meri i poziciju. Plan §14.53.

51. **AMB-10 se ubacuje PRED S6; S6 je parkiran na koraku 1/8** (28.09.2026).
    Nastalo iz pitanja operatera o `KolAmbVracena`: povrat ambalaze je efektivno
    revers i njegovo mesto je `tblAmbalaza`, a kolona stoji i na `tblOtkup`
    (`KolAmbIzdata`) i na `tblPrijemnica`. Merenje je poteralo holisticki pregled --
    **devet mesta knjizenja preko pet dokumenata** -- i model je ispao veci od S6.
    Pun zapis: `docs/DOMEN/AMBALAZA.md` (posle **sedam** krugova review-a; sta je
    koji krug promenio stoji u tabeli 6.12 istog dokumenta).
    **Model:** dogadjaj je jedan red koji imenuje obe strane (`Od -> Na`), knjiga je
    append-only, storno je kontra-stav, vozac je obican nalog, a tudja ambalaza
    ulazi eksplicitnim dogadjajem iz `SpoljniSvet`. Nijedan realni nalog ne sme
    zavrsiti sa negativnim saldom.
    **Odluke operatera:** stampa storniranog dokumenta prikazuje stanje pre storna ·
    `KupciIzlaz` je **revers od kupca + uplata**, ne nov dokument · `Firma` je **jedan**
    nalog · enum kretanja je **zatvoren** (osam vrsta) · nijedan dokument nije
    istovremeno ambalazni i novcani (`AMB-10-ODL-5`) · **`POCETNO_STANJE` je UKINUTO**
    u krugu 6: ono je `NABAVKA` + `IZDATA_PRAZNA`, dakle suvisan pojam.
    **Redosled:** `10a` ugovor · `10-DOK` ambalazni dokument · **`10b-1` pisac** ·
    `10b-2` devet mesta + razlaganje dva slozena pisca · `10c` citaoci · `10d` storno ·
    `10e` brisanje starog -- stare strukture se brisu POSLEDNJE.

52. **AMB-10b-1: knjiga ima pisca, cutover je sledeci rez** (01.10.2026).
    `tblAmbalaza` + sest kolona na kraju (`Od`/`Na` nalog, `VrstaKretanja`,
    `StornoOd`; otisak `E23576DB` -> `A90C9F51`), `PrenesiAmbalazu` i
    `UpisiAmbDokument` u `modAmbalaza` (vlasnik reda za obe tabele),
    `AmbSaldoNaloga` / `AmbObavezaPartneru` / `AmbDeficitZaPrenos` kao citaoci koje
    **pisac mora** da ima -- `AMB-INV-07` i `-09` se bez stanja ne mogu proveriti.
    **Nijedno pozivno mesto nije dirnuto**, pa zatecen tok radi nepromenjen.
    **Dve stvari su usput odlucene, obe izvedene iz postojece matrice, ne nove:**
    pokriva se deficit **partnera** a ne sopstvenog naloga (`AMB-10-ODL-8`), i
    ponavljanje zahteva se meri **zbirom nad parem naloga** jer pisac zahtev deli
    (`AMB-INV-04`, 6.9).
    **Ostaje `10b-2`:** devet mesta knjizenja, razlaganje `SaveOMUlaz_TX` i
    `SaveKupciIzlaz_TX`, numeracija ambalaznog dokumenta, staticka kapija za
    `AMB-INV-08` -- i **odluka pre koda**: da li citaoci idu u istom rezu (6.13).
    BFP 2141 -> **2219**, sabotaza 629 -> **649**, `dokaz.py amb-pisac` **20/20
    DOKAZANO** (potpis izvora `2109b9a57ce615f1` pre i posle, identican).
    Drugi krug review-a je dao jos dva P2, oba uska i oba zatvorena: `AMB-INV-10` je
    BROJAO naloge van granice, pa je dokument sa granicom kao jednom stranom
    (`NABAVKA`, `OTPIS`) prolazio preko dve stanice -- kapija sada zakljucava PAR;
    i jedinstvenost `AmbID`-a je nosio samo citalac obaveze, pa je saldo sabirao dva
    reda sa istim identitetom (`AMB-INV-11`, u zajednickoj kapiji integriteta).
    Treci krug je dao jos jedan P2 iste porodice: **„sve nove kolone prazne"
    dokazuje samo da red NIJE nov, ne i da je VALJAN star** -- red sa kolicinom a bez
    ijedne kolone oba modela tiho je nestajao iz svakog salda. Citalac zato ima
    **cetiri** stanja reda (prazan / legacy / knjiga / kvar), a provera starog
    ugovora odlazi zajedno sa starim modelom u `10e`.
    **Review #400 (NO-GO) je nasao dva prava P2, oba zatvorena u istom rezu:**
    `AMB-10-ODL-3` nije bio sproveden -- `AMB-INV-04` hvata dva protivpartnera samo
    kad se poklope vrsta i tip ambalaze, pa dobija svoju kapiju (`AMB-INV-10`); i
    citalac knjige nije bio fail-closed -- polupisan nov red prolazio je kao legacy
    ili ulazio u saldo pola-pola, a taj saldo odlucuje o sledecem upisu. Oba su
    resena JEDNIM ugovorom zapisanog reda, koji koriste svi citaoci.
    Prvi dvosmerni prolaz je vratio **NIJE DOKAZANO** i oba nalaza su bila u testu:
    jedan je padao fatalno pre imenovane tvrdnje, drugi je merio zbrkano -- stanica
    sa saldom nula je padala zbog **deficita**, pa je tvrdnja o storniranom dokumentu
    prolazila iz pogresnog razloga.
    **Zasto pred S6:** `AMB-10e` i zavrsni korak S6 diraju iste citaoce
    (`modPrint`, `modIzvestaj`, `modScrIzvestaji`, `modStornoDok`); ovim redom se
    diraju jednom.

53. **Rollback koji padne na pola prestaje da bude tih** (02.10.2026).
    Review `main`-a je dao dva P1 o gutanju rollback greške (`modStorno`,
    `modBankaMapiranje`). Mereno: obrazac `On Error Resume Next` + `RollbackTx`
    stoji na **25 produkcionih mesta u 14 modula**, a `.RollbackTx` se zove
    **276 puta** (68 produkcionih u 24 modula, 208 u testovima) — pa zakrpa na dva
    imenovana modula ne bi bila ispravka nego još jedan prekršaj `CLAUDE.md` §2.
    **Uzrok je u primitivu, ne u omotačima:** `clsTransaction.RollbackTx` nije
    imao handler i nije bio atomičan — prva greška iz `RestoreTable` (koja se
    diže eksplicitno, `clsTransaction.cls:117`) prekidala je petlju, pa su ostale
    tabele ostajale nevraćene, `CleanUp` se nije izvršavao (`EnableEvents` ostaje
    `False` do kraja sesije), a `mActive` je ostajao `True`.
    **Primitiv svesno NE diže grešku:** 182 od 276 poziva nisu pod
    `On Error Resume Next` i većina je u aktivnom `EH` bloku, gde bi nov raise
    zamenio originalnu poslovnu grešku i u testovima pretvorio cleanup u pad.
    Zato činjenica ide kao **stanje** (`RollbackNepotpun` / `NevraceneTabele`) i
    van trake (lokalni log + `Monitor_Critical`, koji je no-op kad monitoring nije
    podešen — zato dva kanala). Pravilo je zapisano u
    `ARCHITECTURE_CONTRACT.md` uz „Snapshot nije vlasništvo“.
    Sabotaža je **preimenovanje** tabele (`GetTable` vraća `Nothing` → 91), ne
    menjanje šeme; tri privremene tabele imaju **po dve kolone** jer `Value2` nad
    jednom ćelijom vraća skalar, a `RestoreTable` radi `UBound`. Tri sabotaže, po
    jedna na svaku posledicu — ne grupišu se (isti test, ista procedura za dve).
    `WHO_WRITES.md` dobija tri reda `TST_RB_*` sa **0 produkcionih pisaca**:
    generisani artefakt verno prijavljuje `AddTableSnapshot` iz testa. Alternativa
    (naučiti `who_writes` da prećuti prefiks `TST_`) je izmena **kapije** i traži
    svoj dvosmerni dokaz — ostaje kao moguć naredni process rez, ne u ovom.
    **Review je na to dao P1 i bio je u pravu:** prijaviti nije isto sto i
    **zatvoriti**. Property na `tx` objektu nestane kad pozivalac izađe, pa se
    sistem posle nepotpunog rollback-a vraćao u **puno operativan** režim — nova
    transakcija dozvoljena, AutoSave zakazan, a `ZatvoriAplikaciju` radi
    `Close SaveChanges:=True`. Operater koji samo zatvori program zabetonirao bi
    parcijalno vraćen podatak. Prethodna (loša) verzija je sistem ostavljala
    očigledno polomljenim; prva moja verzija ga je vraćala u ispravno stanje dok
    zna da podaci to nisu — gore.
    Brana je sad **globalna za sesiju** (`modTxState`, obrazac `modImportState`) i
    zatvara pet puteva: `BeginTx`, `MarkDirtyAndSchedule`, `AutoSaveAfterCommit`,
    `Workbook_BeforeSave` (jedina tačka kroz koju prolaze Ctrl+S / File > Save /
    Save As / `.Save` iz VBA) i `ZatvoriAplikaciju`. **Bez registra** i
    **fail-closed**, obrnuto od `modImportState`, i oba namerno: recovery *je*
    reload (Save je zatvoren, pa na disku stoji stanje pre transakcije), a nema
    čitanja koje može da pukne. Perzistiran marker bi svesku učinio trajno
    nesnimljivom bez izlaza iz aplikacije.
    Test je prepisan: tvrdio je da `BeginTx` posle nepotpunog rollback-a
    **prolazi** — to je bila acceptance odluka za ponašanje koje ne želimo. Sada
    meri **ishod kroz prave seam-ove**: `BeginTx` diže `UPIS ZATVOREN`, a
    `ThisWorkbook.Save` ostavlja `Saved = False`. Oba smera: i da PRE kompromisa
    Save prolazi, i da POSLE reset-a (= reload) opet prolazi.
    **Četvrti krug review-a je našao rupu u DOKAZNOM modelu, ne u kodu:** poruka se
    u dva EH bloka (`modAgroUnos`) računala **pre** `tx.RollbackTx`, pa je marker
    tada još bio prazan i parcijalan rollback je i dalje vraćao „promene vraćene" —
    a kapija je bila **zelena**, jer je merila *prisustvo* wrappera, ne *redosled*.
    Zato je dodato `ROLLBACK_TVRDNJA_RED` (u proceduri koja sama poseduje `tx`,
    wrapper mora stajati posle zadnjeg `.RollbackTx`), a `Err` se čuva pre
    rollback-a. Mereno po proceduri nad celim izvorom: **tačno 2** takva mesta;
    7 u `modBankaMapiranje` je bilo ispravno, a 11 (`modScrDokumenti`,
    `modScrBankaUvoz`) ne poseduje `tx` pa im je rollback završen unutra.
    Usput: dokaz nivoa „da li se self-test uopšte vrti" otkrio je da je zbir
    slučajeva **ručno** održavan i nije uključio dve nove liste — `--self-test` je
    javljao 130 i posle dodavanja 12 slučajeva. Ispravljen na **142**, ali broj se
    više ne uzima na reč nego se dokazuje padom (v. [[broj-tvrdnji-je-merenje]]).
    **Dokaz je IZMEREN**, na exact head-u `783b7946`:
    `RunAllTests` **200/0 ZELENO** · `dokaz.py rollback` **6/6 DOKAZANO** (potpis
    izvora `935f070a003272e0` identičan pre i posle) ·
    `vba_gate --require-green --suite RunAllTests` **rc=0**
    (`izvor bc84a7d4168e, ugovor 9c3f30001038`) · 19/19 jeftinih kapija `rc=0`
    · 142 self-test slučaja · katalog 649 → **655**.
    `--require-green` **bez** `--suite` je `rc=2` i to je tačno: samo je
    `RunAllTests` puštena nad ovim izvorom, pun prolaz ide pred release.
    **Compile ostaje ručna kapija operatera** (`--mark-compile`).
54. **Ulaz za storno ambalaže pomeren PRED cutover** (03.10.2026, `AMB-10-ODL-16/-17`).
    Plan ga je držao kao `10d`, **posle** devet mesta knjiženja. Merenje pred prvi
    rez je oborilo taj red: nov čitalac salda (`RedDoticeKnjigu`,
    `AmbSaldoNaloga`) **ne čita `Stornirano` nigde**, a `modStorno` otkazuje gajbe
    **zastavicom** — otkup (`:170`), otpremnica (`:245`) i prijemnica (`:458`).
    Dokument presečen na nov model a storniran zastavicom ostavio bi gajbe na
    saldu **tiho**, a tvrdnja koja čita zastavicu ostala bi **zelena**: lažno
    zeleno, ne pad. `NABAVKA` je mogla da legne sama jer storno put **nema**.
    Ulaz je `modAmbalaza.StornirajAmbalazuDokumenta(tx, dokTip, dokID)` — **po
    dokumentu**, jer životni ciklus ima dokument a ne red; datum kontra-stava je
    datum **originala** (zastavica je red uklanjala iz **svih** perioda);
    idempotentan; `AMB-INV-09` se meri nad **posle-stanjem** postojećim čitaocem.
    Kontra-stav nosi zamenjene `Od`/`Na`, pa se proverava **u obrnutom smeru** —
    pravilo stoji na **jednom** mestu u pisaču, uz čitaoca koji ga je već imao.
55. **Otkup — prvo presečeno mesto knjiženja** (03.10.2026, `10b-2`).
    Četiri noge → dva događaja (`UZ_ROBU` Kooperant→Stanica, `IZDATA_PRAZNA`
    Stanica→Kooperant), **oba pod `Otkup`** — čime je napetost **T3** (pozajmljen
    `OM-Izlaz-Koop`) rešena, i to ne iz estetike: `AmbIzvornaTabela` je zatvorena
    mapa, pa bi pozajmljen tip pao fail-closed na `AMB-INV-08`.
    **Dva nalaza koja kapija nije uhvatila, a čitanje je:** promena potpisa
    `CreateOtkup` nije oborila `vba_check` jer je drugi pozivalac
    (`IspravkaOtkupa_TX`) zove u **izraznoj poziciji** (`ARNOST` to ne vidi) —
    projekat se ne bi kompajlirao; i parametar `src` je ostao bez upotrebe kad je
    `TrackAmbalaza` nestao.
    **Posledica na fixture je poslovna, ne tehnička:** pisac sada **traži** da
    kooperantove gajbe postoje, a mereno je **139** poziva `CreateOtkup_TX` u BFP
    suite-u. Opticaj zaseva **jedno** mesto (`SeedAmbalazaOpticaj`) — nabavka po
    stanici i tipu, pa izdavanje praznih kooperantima; u 90–95% slučajeva
    kooperant i u stvarnosti vraća **naše** gajbe (`AMB-10-ODL-18`).
    Protokol potvrde ide **dvema putanjama**: ekran pita i **zadržava podatke**
    (poziv se ponavlja iz istog poziva), sync **auto-potvrđuje** (operatera nema).
    Slučaj se prepoznaje po **broju greške** — zato `outErrNum`, i zato potvrda
    deficita izlazi **pre** `LogError`/`DOKUMENT_SAVE_FAIL`.
    **Neizmereno i tako prijavljeno:** sync auto-potvrda i `MsgBox` grana nemaju
    test.
56. **Tri P1 u storno sloju** (03.10.2026, review `cac06c3c`).
    Sva tri su u **storno sloju**, a glavni write put otkupa je prošao — pa je
    nalaz da se cutover **ne nastavlja** na preostalih osam mesta dok storno
    lifecycle ne legne jednom kako treba.
    **P1 #1 — moja odbrana je falsifikovana.** `StornirajAmbalazuDokumenta` je sam
    zvao `BindSourceDocument`, uz obrazloženje „nije samopotvrda jer kapija traži i
    izvornu tabelu u snapshotu". Snapshot je **jeftin** i ne dokazuje da je dokument
    promenjen: pozivalac je mogao da anulira ambalažni efekat **aktivnog** otkupa i
    prođe sve kapije. Vezivanje je prešlo na kanonskog pisca zaglavlja
    (`modStorno.StornoOtkup`: `MarkRowStornirano` → `bind` → primitiv), pa
    `AMB_BIND_DOZVOLJENI` ponovo znači „pisci izvornog dokumenta". `tblAmbalazaDokument`
    još nema kanonskog storno pisca, pa njegov ledger storno **namerno** pada
    fail-closed.
    **P1 #2 — kontra-stav je zaobilazio `AMB-INV-07`.** Ide direktno kroz
    `UpisiRedKnjige`, a `AMB-INV-07` živi u `PrenesiAmbalazu`; proveravan je bio samo
    `AMB-INV-09`. Storno dokumenta čija je ambalaža kasnije otišla dalje mogao je da
    **commituje negativan fizički saldo** — stanje koje normalan pisac eksplicitno
    zabranjuje. Provera sada ide nad **posle-stanjem**, nad svakim pogođenim realnim
    nalogom; dva popisa naloga su spojena u **jedan** (`ZabeleziNalog`), a klasu bira
    čitalac.
    **P1 #3 — postojeći undo je postao kontradiktoran.** Žurnal je **ćelijski**, a
    kontra-stav je **nov red** koji u njemu ne postoji — `UndoOperation_TX` je vraćao
    zaglavlje u aktivno, a ambalažni efekat je ostajao anuliran (dokument aktivan sa
    nula ambalaže). Odgovor je **fail-closed odbijanje** (`AMB-10-ODL-19`) u
    `UndoGuardReasonZaOp`, koju gledaju i komanda i ekran oporavka. Brisanje
    kontra-stavova i storno storna su **odbijeni jer krše važeći ugovor**; semantika
    „vraćanja storna" nad append-only knjigom je **poslovna odluka** koja stoji
    otvorena, vidljivo i sa razlogom.
    Dokaz: `Test_Amb_StornoNePraviMinus` 6 tvrdnji (kroz **produkcioni**
    `StornoOtkup_TX`, nad **nezasejanim** tipom — nad zasejanom stanicom se minus ne
    može proizvesti) · `Test_Amb_UndoStornaOdbijenNadKnjigom` 6 · tri nove sabotaže.
    Katalog 675 → 678.
57. **Dva P2 — jedna greška: skraćen kanonski identitet** (03.10.2026, review
    `465790f0`). Oba nalaza su na mestu gde se sistem **generizuje** za preostalih
    osam write-site-ova, i oba su isto: uzeo sam uži ključ od onog koji domen već
    nosi.
    **P2 #1.** `AmbImaKontraStav` je tražio samo `DokumentID`, uz komentar da „tip
    ne dodaje razlučivost". `AMB-INV-04` nosi `DokumentTIP` **tačno zato** što se
    jedan globalni namespace `DokumentID`-eva ne sme pretpostaviti — dakle ponovo
    ista pretpostavka koju je domen eksplicitno odbacio. Ključ je sada kompozitan, a
    tip se **izvodi iz tabele žurnalnog reda**, ne iz oznake operacije: za otkup bi
    danas bile iste, ali za revers je oznaka `OM-Izlaz-Koop` dok će u knjizi stajati
    `AmbalazaDokument`. Oba smera (`tip → tabela`, `tabela → tip`) čitaju **jedan
    popis** (`AmbIzvorniParovi`).
    **P2 #2.** `tblAmbalazaDokument` nije imao kapiju **zauzetosti** broja
    (`RequireAmbDok` sudi oblik, ne zauzetost), pa su dva poziva sa istim ručno
    prosleđenim brojem davala dva `AmbDokID`-a i **jedan poslovni broj u istom
    nizu**; a generator je skenirao samo `BrojOwnerID`, dok je kanonski vlasnik
    `BrojOwnerTip + BrojOwnerID` (u AgriX-u `VozacID` može biti jednak `StanicaID`).
    Oba sada čitaju **jedan sken** (`AmbDokNizSken`) sa istim opsegom; storniran
    dokument **drži** svoj broj, kao i otkupni list. Kapija pokriva sve putanje jer
    je `UpisiAmbDokument` jedini pisac te tabele.
    Usput je oboren i moj komentar koji je tvrdio da „prosleđen i izračunat broj
    prolaze istu kapiju" — kapija zauzetost nije sudila.
    Dokaz: `Test_Amb_DokBrojZauzetPoVlasniku` 5 tvrdnji · tri tvrdnje dopune u
    `Test_Amb_UndoStornaOdbijenNadKnjigom` (isti ID pod drugim tipom **nije**
    pogodak) · tri nove sabotaže. Katalog 678 → 681.
58. **Dokaz reza: DOKAZANO (grupno), posle šest nalaza iste klase** (04.10.2026).
    `dokaz.py --grupe 6` nad **25** sabotaža koje je ovaj rez dodao ili dirao:
    **25/25 crvenih, svih 25 obara SVOJU tvrdnju**, izvor identičan pre i posle
    (`71bf98185db62752`). Trebalo je **četiri** prolaza da se tamo dođe, i svaki je
    našao nalaze iste klase: **tvrdnja koja meri ISHOD, a ne RAZLOG, ostaje istinita
    kad se ugasi jedan sloj kapije.**
    Redom: `ODL-9` pokriven `ODL-10` · `ImaSnapshot` pokriven time što `CleanUp`
    briše i snapshote · manjak stanice pokriven protokolom potvrde · `ODL-9`
    validator pokriven protokolom potvrde · seed zavisio od generatora · i
    posledica duplog broja pokrivena **novom kapijom iz istog reza**
    (`AMB-10-ODL-20`): sabotaza generatora više ne pravi dva ista broja nego
    **odbijen upis**. Kapija se nije slabila — tvrdnja je ojačana.
    Uz to je nađeno da **69 BFP tvrdnji nije bilo dokazivo** jer je kapija merila
    podniz (v. red u „Dug sa imenom“).
    `grupno izmereno 19/25` — pun pojedinačni dokaz je isti poziv bez `--grupe` i ide
    pred release.
59. **Otpremnica — drugo presečeno mesto knjiženja** (04.10.2026, 6.12f).
    Jedan događaj: `Stanica → Vozac`, `AMBALAZA_UZ_ROBU`. **Vrsta je pročitana iz
    6.7**, ne izvedena iz para — oba naloga su `SOPSTVENI`, pa bi matrica pustila i
    `PRENOS_INTERNO`; razlika je poslovna.
    **Vozač prestaje da bude žig:** stari red je imao jedan entitet i `VozacID` kao
    oznaku, a vozačev saldo je nastajao **inverzijom smera** — fail-open po 6.8.
    Sada je nalog, pa test može da tvrdi da gajbe **idu na njega**; to se u starom
    modelu nije moglo napisati.
    **Nema protokola potvrde** (razlika od otkupa): izvor je `SOPSTVENI`, pa je po
    `AMB-10-ODL-8` manjak stanice **tvrdo odbijen**, ne pitanje.
    **Redosled se morao promeniti:** knjiženje je stajalo PRE izmene zaglavlja, a
    `ODL-15` traži bind POSLE nje — sada je označi izdato → veži → knjiži.
    `tx` je obavezan i na `OtpIzdaj` i na `StornoOtpremnica` (3 pozivna mesta, sva
    u tx sa obe tabele — izmereno pre koda). Dva zatečena sidra je kapija
    prijavila po imenu i pomerena su bez menjanja tvrdnji. Katalog 680 → 682.
60. **Nalaz dokaza: tvrdnja je stajala IZA rane izlazne tačke** (05.10.2026).
    `dokaz.py` nad otpremničkim rezom: `crvenih 4 / sabotaza 4`, izvor pre/posle
    identičan — ali `ispravka-ne-stornira-staru` **NE OBARA SVOJ TEST**.
    **Uzrok nije u produkcionom kodu nego u položaju tvrdnje.** Kad se ugasi
    `StornoOtpremnica` u `OtpIspravi`, posledica **nije** „dve aktivne
    otpremnice": `OtpRequireIzvorValjan` odbije novu jer je izvor još u sastavu
    aktivne stare — ista veza koju `OtpIspravi` imenuje u svom komentaru
    („storno stare IDE PRE nego sto nova primi izvore"). Ceo poziv padne i vrati
    `""`, pa je u testu pucala samo tvrdnja da je ispravka **uspela** — posledica,
    ne razlog — a ciljana `Ispravka: stara je stornirana` stajala je iza
    `If Len(nova) = 0 Then GoTo Kraj` i **nikad nije bila izvršena**.
    Ispravka je **premeštanje** te tvrdnje iznad kapije: vrednost koju čita
    postoji i kad poziv padne (rollback vraća staru u aktivno stanje), pa tada
    čita `""` umesto `"Da"` i puca **po imenu**. Katalog i tekst tvrdnje ostaju
    netaknuti — pogrešan je bio **redosled u testu**.
    **Treći mehanizam istog oblika.** `NE OBARA SVOJ TEST, nego: …` je do sada
    značio ili pogrešno imenovanu tvrdnju ili `LogFatal`; ovo je **rana izlazna
    tačka** posle tvrdnje o ishodu. Zapisano u memoriju.
    Provereno i za ostale četiri `ispravka-*` sabotaže nad istim testom
    (`nacrt-prolazi`, `bez-clanstva`, `trag-po-broju`, `pad-ostavlja-storniranu`):
    kod njih ispravka **uspeva**, pa su im tvrdnje dostižne — zaklonjena je bila
    tačno jedna.
    **Broj tvrdnji se ne menja** (2318): ista tvrdnja, drugo mesto. Promena broja
    bi značila da se nešto prećutalo (v. „broj tvrdnji je merenje").
    Usput popravljeno: **rep stavke 53** (`modTxState`, `TST_RB_*`, četvrti krug
    review-a, katalog 649 → 655) stajao je od `ef60fc91` na **kraju** liste, pa ga
    je svaka nova stavka odvlačila dalje — od 59 se čitao kao deo otpremničkog
    reza. Vraćen je pod 53, kojoj po sadržaju i datumu (02–03.10) pripada.
    **DOKAZ JE IZMEREN — za stavke 59 i 60 zajedno, nad jednim izvorom.**
    `run_vba.py` pun prolaz **ZELENO**: 12/12 suita, `RunBusinessFlowProSuite`
    **0/2318**, `RunAllTests` 0/200, banka 0/241, storno 0/163, palete 97,
    faktura 35, agrohemija 25, Sheets 72. Marker nad izvorom `c7a34a14b436`
    (ugovor `849cba105e9c`, sveska `otkup_test.xlsm/d24883a3`) ·
    **compile potvrđen** nad istim izvorom — a compile je **jedina** kapija koja
    bi sama uhvatila P1 iz `OtpIspravi` (`Variable not defined: tx`) ·
    `dokaz.py` **DOKAZANO (grupno)**, `crvenih 4/4`, potpis izvora
    `fcde30a28b7da85c` identičan pre i posle, `grupno izmereno 2/4` — pun
    pojedinačni dokaz je isti poziv bez `--grupe` i ide pred release · jeftine
    kapije `rc=0` (`vba_check` 191/682/0+10, arnost 300, scope 5522/0, schema,
    ownership, čitaoci).
    **2318 je držalo** — isti broj pre i posle premeštanja tvrdnje, što je i bila
    tvrdnja o samoj zakrpi: ista tvrdnja, drugo mesto.
    Usput izmereno o **grupisanju**: `amb-otp-storno-bez-kontrastava` u grupi
    *nije* oborila svoju tvrdnju, jer je kaskada iz `ispravka-ne-stornira-staru`
    oborila ceo poziv ispravke pre nje; sama je **OK**. Drugi put da grupisanje
    traži solo ponavljanje — protokol `--grupe` to radi sam, i to je razlog
    zbog kog postoji.
61. **Prijemnica — treće presečeno mesto, i dva nova pravila** (05.10.2026, 6.12g).
    Dva događaja nad **jednim neuređenim parom**: `Vozac → Kupac` uz robu i
    `Kupac → Vozac` povrat praznih. `AMB-INV-10` prolazi jer je par neuređen —
    što je ujedno provera da je tako i mišljen.
    **Rez je prvo BLOKIRAN, i to zapisanim pravilom.** Obrnuta kapija `ODL-13`
    je povrat od kupca dozvoljavala **samo** na `REVERS_PARTNERA`, a
    `modAmbalaza.bas` je to i izričito branio u komentaru. Istovremeno §3 red 7
    i `AMB-04` kažu da prijemnica taj povrat **knjiži**. Dva zapisana pravila,
    jedan događaj — `DOMAIN GAP`, pa nije pisan kod nego je pitanje išlo
    operateru.
    **Odgovor je pobio premisu, ne kapiju:** prijemnica je **i sama partnerov
    dokument** (eksterna je), i povrat se knjiži **pod njenim brojem, bez
    dodatnog**. Dakle `ODL-10` nije zaobiđen nego **ispunjen** — i uslov je
    **vlasnik broja**, ne vrsta dokumenta (`AMB-10-ODL-22`). Običan `REVERS` nad
    `Kupac → Vozac` i dalje pada, jer je njegov broj naš.
    **`AMB-10-ODL-21`: lanac se odmotava obrnuto od fizičkog reda.** Izmereno u
    `modStornoFlow`: kaskada je stornirala **otpremnice pre prijemnica**, a od
    `10b-2` je knjiga stvaran saldo — pa bi vozač otišao u minus i `AMB-INV-07`
    bi oborio celu kaskadu. Red je obrnut; kad lanac nije naš (`ownsChain =
    False`) storno otpremnice **pada**, i to je tačno — stari model je tu
    prijavljivao „delimičan uspeh kao pun".
    **Dva komentara koje je merenje pobilo pre review-a.** (1) Napisao sam da je
    kapija povrata `AMB-INV-09`; `AmbDoprinosObavezi` kaže da obavezi doprinose
    samo `ULAZ_TUDJE`/`VRACANJE_TUDJE`, a ove vrste doprinose **nulu** — kapija
    je `AMB-INV-07`, jer je i `Kupac` REALAN. (2) Hteo sam da spojim dve grane
    kapije koje izgledaju kao duplikat; `jePartnerov` radi `Exit Function` pre
    druge, pa su im sabotaže razlučive — spajanje bi dve svelo na jedno sidro.
    `SeedAmbalazaOpticaj` je morao da dobije **vozače** (`PRENOS_INTERNO` po
    `ODL-7`): prva noga polazi od vozača, a seed je punio samo stanice i
    kooperante — 14 zatečenih pozivnih mesta bi palo na `AMB-INV-07`.
    Katalog 682 → 687.
62. **P1: zaštita eksternog lanca je bila izvedena iz SALDA** (05.10.2026,
    review `cb6c93bb`). Napisao sam da `ownsChain = False` rešava `AMB-INV-07`
    sam — *„prijemnica ostaje, vozač nema gajbe, storno otpremnice padne"*. Važi
    **samo** kad kupac nije vratio dovoljno praznih. **Puna zamena** (20 punih,
    20 praznih) vraća vozačev saldo, kontra-stav otpremnice **prolazi**, i
    kaskada javi `ok = True` nad lancem u kom eksterna prijemnica ostaje aktivna
    i vezana na stornirane dokumente — **lažno uspešno poslovno poništenje**.
    **Uzrok klase:** `AMB-INV-07` sudi **posle-stanje salda**, a ne **lifecycle
    zavisnost**. Saldo ne zna da aktivan nizvodni dokument još zavisi od onog
    koji se stornira. Lek nije jači saldo nego **eksplicitna kapija pred svakom
    mutacijom**; dve politike (eksterni dokument netaknut / poništenje celog
    toka) se iskljucuju, pa se bira **odbijanje**, ne orphaning.
    **Moj test je bio deo problema:** `Test_PRJ_LanacSeOdmotavaObrnuto` koristi
    `kolAmbVracena = 0` — jedini slučaj u kom odbrana iz salda slučajno važi. Nov
    test uzima **punu zamenu** i tvrdnja br. 4 imenuje bas to: *„saldo je vraćen,
    dakle saldo ne bi zaustavio storno"*. Bez te tvrdnje bi i ugasena kapija
    prolazila.
    Kaskada je `Private`, pa je dodat **test seam** `PonistiZbirnaChain_Test` —
    po zatečenom obrascu `DistinctActiveValues_Test`, jer javni put traži
    correction context i `forceConfirm`, pa bi pad mogao da dođe sa tri sloja.
    Usput izmereno pre koda: prazan `scopeID` **širi** skup dece
    (`SuziDecuNaZbirnu`, pravilo 1), pa je kapija nad njim fail-closed — i skup
    se broji **bez obzira na `ownsChain`**, jer je `prijIDs` u toj grani namerno
    prazan i kapija nad njim bi bila placebo.
    Katalog 687 → 688.
63. **P2 zatvoren: vlasnik broja prijemnice NAMERNO ostaje njen `KupacID`**
    (05.10.2026, odluka operatera). Review je postavku imao tačnu — `KupacID`
    odgovara na „ko je kupac u poslu", `BrojOwner` na „čijem nizu pripada broj",
    i `AMB-10` ih svuda drugde razdvaja.
    **Odgovor je potvrđen merenjem koje je pobilo obrazloženje koje sam
    nameravao da napišem.** Hteo sam da napišem „brojevi prijemnice se ne
    generišu kod nas, pa nema ništa čiji bi niz bio" — generator **postoji**:
    `GenerateBrojPrijemnice(kupacID, datum)` scope-uje niz baš po
    `(KupacID, dan)` (`MaxSeqFromTable(... COL_PRJ_KUPAC, kupacID, datum)`), uz
    svoj komentar da auto-numeracija važi **samo** za hladnjača-kupca a ostali
    nose eksterni broj. Dakle vlasnik broja **jeste** kupac u oba režima, i
    poklapa se sa `ODL-20`. To je jače obrazloženje od onog koje sam imao — i
    peti put u ovom rezu da je merenje pobilo odbranu **pre** review-a.
    Kod se **nije menjao** (ponašanje je već bilo takvo); dodat je test koji meri
    **odsustvo grane** (dva različita kupca — jedan ne bi razlikovao pravilo od
    hardkodirane vrednosti) i fail-closed default za tip van mape, plus sabotaža
    `amb-odl22-vlasnik-broja-iz-pogresne-kolone` (tip vlasnika ostaje `Kupac`, pa
    klasa izgleda dobro, a ID je vozačev). Upisan je i uslov za reviziju: ako broj
    ikada počne da se generiše iz **našeg** niza nezavisnog od kupca, mapa traži
    granu. Katalog 688 → 689.
64. **Baza `RunAllTests` je pala, i `dokaz.py` je stao pre merenja**
    (05.10.2026). `STOP: baza nije zelena. Dokaz bi merio crveno koje sabotaza
    nije izazvala.` — kapija je uradila tačno ono zašto postoji.
    **Uzrok izveden iz izvora, bez Immediate prozora:** test 26
    (`T_IspravkaPrijemnice_SkipIRelink`) **dva puta** piše prijemnicu sa 40
    gajbi — i to su **jedina dva** takva poziva u celom `modTest` (mereno). Od
    `10b-2` prijemnica knjiži `Vozac → Kupac`, a `make_fixture` nosi samo
    redove **starog** oblika (`Smer`/`EntitetID`), koje nov čitalac ne vidi —
    saldo vozača u novom modelu je **0**. `PrenesiAmbalazu` sprovodi
    `AMB-INV-07` i na **običnom** upisu (`AmbDeficitZaPrenos`), a manjak
    **sopstvenog** naloga se po `AMB-10-ODL-8` ne pokriva tuđom ambalažom nego
    je **tvrdo odbijen** — vozač nema šta da pokrije manjak.
    **Blast radius je izmeren, ne pretpostavljen:** prijemnicu kroz te ulaze
    pišu samo `modTest` (2 poziva) i `modBusinessFlowProTests` (20, već
    zasejan u 61). `RunPaleteTestSuite` i `RunStornoTestSuite` koriste zatečene
    redove fixture-a, pa ih ovo ne dira.
    Optičaj se zato zasejava **u testu**, kao preduslov sa svojom tvrdnjom — ne
    u `make_fixture` (traži ponovnu izgradnju sveske) i ne u `RunAllTests` (traži
    ga tačno jedan test). Seed je **idempotentan**, i to nije kozmetika: suite se
    vrti nad istom sveskom više puta, a kapija zauzetosti broja (`ODL-20`) bi
    odbila ponovljen broj istog dana.
    **Usput: propuštena kapija cele sesije.** `who_writes --check` (generisani
    `WHO_WRITES.md`) nije bio puštan — puštan je samo `--check-ownership`.
    Dokument je bio zastareo **samo zbog ovog seeda** (regenerisan: `modTest`
    ulazi kao test-pisac `tblAmbalaza` i `tblAmbalazaDokument`, verno, jer seed
    ide kroz produkcione pisce i `AddTableSnapshot`). A11 prolazi. Pravilo
    „CI kapije se vrte sve" je imalo tri clana u mojoj glavi, a ima četiri.
65. **Sirotan u knjizi: `AMB-INV-04` je uhvatio sudar identiteta** (05.10.2026).
    BFP je ostao na 2359/2, i obe pale tvrdnje su bile isti upis. Razlog se
    **nije video iz izveštaja** — `SavePrijemnica_TX` grešku ne propagira nego
    je štampa u Immediate prozor, koji runner ne hvata. To je kvar **instrumenta**:
    tvrdnja o upisu bez razloga me je dvaput poslala u nov prolaz. Dodat je seam
    `PrjUpisiSaRazlogom` koji zove **jezgro** (ono grešku diže) i nosi `Err` u
    tvrdnju — i tek tada se razlog video:

    ```
    AMB-INV-04: Prijemnica 'PRJ-00011' je vec knjizio AMBALAZA_UZ_ROBU
    za 'Test Gajba' (red AMB-4653CCE7...)
    ```

    **Isti ID u oba pada** — broj se ponovo dodeljuje, a u knjizi je ostao red
    starog nosioca. Mehanizam: `SavePrijemnica_TX` commituje **svoju** tx
    (dokument + knjiga), pa spoljna tx testa vrati `tblPrijemnica` ali ne i
    `tblAmbalaza` — koje nema u njenom snimku. Red ostaje **sirotan**, `GetNextID`
    ponovo izda isti broj, i invarijanta ga obori u **tuđem** testu dva testa
    kasnije. Tačno obrazac „rollback vraća CEO dokument".
    **Kapija nije pogrešila — uradila je svoj posao.** Nalaz je u **opsegu
    transakcije jednog testa**, i lek je jedan red: `tblAmbalaza` u njen snimak.
    Izmereno nad svim test modulima: **tačno jedna** takva procedura
    (`Test_ZBR_PaletaNasledjujeGeneracijuPrijemnice`). Izmereno nad produkcijom:
    jedine procedure koje imaju tx i pišu prijemnicu su sama dva `_TX` omotača, i
    **oba** snimaju `tblAmbalaza` — produkcija ovu rupu nema.
66. **Dokaz prijemničkog reza: DOKAZANO (grupno), 7/7** (06.10.2026).
    `RunAllTests` 200/0 · `RunBusinessFlowProSuite` **2376/0** — tačno predviđen
    broj (2359 + 17 tvrdnji koje su do tada preskakane posle palog upisa), pa
    ništa nije prećutano. Potpis izvora `fdae701158aeb01b` identičan pre i posle;
    `grupno izmereno 4/7`, pun pojedinačni dokaz ide pred release.
    **Poslednji nalaz je bio dvoslojna kapija.** `amb-odl22-vlasnik-broja-se-ne-gleda`
    nije obarala ništa: blok `povratOdKupca` ima **dva** uslova i oba nose isti ID
    odluke (`AMB-10-ODL-10`) — prvi traži da je vlasnik broja **Kupac**, drugi da je
    **baš taj** kupac. Nad dokumentom bez vlasnika oba su netačna, pa je gašenje
    prvog ostajalo nevidljivo. Tvrdnja je merila ID odluke; sada meri tekst koji
    proizvodi **samo prvi sloj**, provereno da u izvoru stoji tačno jednom.
    **Šta je ovaj rez koštao, i zašto.** Četiri zastoja do zelene baze, i nijedan
    nije bio u modelu reza: preduslov vozača u `RunAllTests`, dva moja testa koja
    su merila premisu koju nisu postavila, i siroče u knjizi od pretesnog snimka.
    Ali **dva prolaza su otišla samo na to što tvrdnja o upisu nije govorila
    zašto** — pisac vraća `""` i štampa razlog u Immediate prozor koji runner ne
    hvata. Čim je seam počeo da nosi `Err`, uzrok se video iz prvog pokušaja.
    Pouka je instrumentalna, ne domenska: **tvrdnja o upisu bez razloga je slepa
    kapija**, i košta više od samog kvara.
67. **Pun prolaz ZELENO: 12/12 suita** (06.10.2026, izvor `9b149e39aeea`).
    `RunAllTests` 200 · `RunBusinessFlowProSuite` 2376 · `RunStornoTestSuite` **164**
    (bilo 163 — T18 je dobio tvrdnju više) · `Test_StornoCentar_All` · banka 241 ·
    palete 97 · faktura 35 · agrohemija 25 · Sheets 72 · golden · licenca.
    **Pun prolaz je našao dve stvari koje ciljane suite nisu mogle**, i nisu iste
    vrste. **(1) Moja greška u kapiji:** poništenje **prijemnice** nad eksternim
    lancem ukida baš tu prijemnicu, ali ju je grana stornirala **tek posle**
    kaskade — pa je kapija blokirala operaciju zbog dokumenta koji pozivalac
    upravo gasi, a redosled je bio obrnut od fizičkog. Kaskada sada dobija
    **subjekat**: izuzima ga iz blokirajućeg skupa (druga aktivna prijemnica i
    dalje blokira) i stornira ga **prvog**. Subjekat je **skup**, ne jedan ID —
    broj prijemnice pokriva i Klasu I i II. **(2) Zastareo test:** `T18` je tvrdio
    da taj tok **uspeva**, što je tačno ono što je review nazvao lažno uspešnim
    poništenjem; preveden je na nov ugovor. Stara tvrdnja „prijemnica netaknuta"
    je **zadržana uz napomenu da sama ne razlikuje stari i nov ugovor** — prolazi u
    oba, pa stoji uz tvrdnje koje ga razlikuju.
    Ostaje još samo ručna kapija: `Alt+F11 → Debug → Compile VBAProject` +
    `--mark-compile`.
68. **Revers — četvrto presečeno mesto, i `ODL-5` zatvoren** (06.10.2026, 6.12h).
    Šest nogu u četiri smera postalo je **četiri reda**, po jedan na smer, iz
    **zatvorene mape u ugovoru**. Vozač je iz **žiga** postao **nalog** — u starom
    modelu se njegov saldo dobijao inverzijom smera.
    **Merenje je promenilo opseg pre koda.** Dokument opisuje `SaveOMUlaz_TX` kao
    prekršaj `ODL-5` (ambalažni dokument nosi novac pod istim brojem). Pozivna
    mesta kažu da **nijedan živ poziv ne meša klase** — F5 šalje `kolAmb:=0`, F7
    `novac:=0`. Prekršaj je bio u **potpisu**, ne u ponašanju: razlaganje je
    mehaničko, bez promene poslovnog toka i bez pitanja za operatera.
    Zadržano svesno: **`RequireBrojUKontekstu`** — nov model pokriva zauzetost
    (`ODL-20`) ali ne i oblik/kontekst broja, pa bi prelazak tu kapiju tiho
    izgubio. Osam kopija pravila „nalog je obavezan" svedeno na **jedno telo**.
    **Lekcija iz 6.12g primenjena PRE prvog prolaza:** provera naloga je
    dvoslojna, pa tvrdnja o odbijanju meri **tekst prvog sloja**, ne ishod — helper
    zato vraća `Err.Description`. Bez toga bi sabotaža nad tom proverom obarala
    ništa.
    11 zatečenih `REV` testova numeracije **preseljeno parserom**, ne prepisivanjem:
    11 blokova od po šest redova je 11 prilika za tihu grešku u jednom polju. Parser
    je usput našao i **jedan poziv koji nije revers nego isplata** (`novac:=5000#`) —
    on ostaje na starom piscu. Katalog 689 → 692.
69. **Rez reversa izmeren: DOKAZANO + pun prolaz ZELENO** (06.10.2026,
    izvor `ea584f8335fe`). `RunAllTests` 200 · `RunBusinessFlowProSuite` **2392**
    · `RunStornoTestSuite` 164 · banka 241 · i ostalih osam — **12/12, nula**
    **padova**. `dokaz.py` **DOKAZANO (grupno)**, `crvenih 3/3`, potpis izvora
    `c1d1fae401498fa2` identičan pre i posle.
    **Put do zelenog je dao tri nalaza, i sva tri su bila u MOM kodu.**
    **(1) Rezu je falio lifecycle** — presekao sam pisca pre nego što je storno
    postojao, tačno ono što `AMB-10-ODL-16` zabranjuje. `tblAmbalazaDokument` je
    imao kolonu `Stornirano` i nijedan put da je okrene; dobio je
    `StornirajAmbDokument_TX` i time je zatvoren zapisan dug `10d`.
    **(2) Dva niza broja, a ja sam računao na jedan.** Ispustio sam
    `RequireBrojSlobodanUNizu` misleći da ga `ODL-20` zamenjuje — ali `ODL-20`
    sudi nad `tblAmbalazaDokument`, a stari niz nad `tblAmbalaza`. Broj zauzet
    starim reversom bio bi slobodan za nov. Dok oba oblika postoje, oba niza
    važe; kad stari redovi nestanu (`10e`), provera postaje mrtva i briše se s
    njima.
    **(3) Sabotaža vrste je UBIJALA test pre tvrdnje** — `NE OBARA SVOJ TEST,
    nego: <ImeTesta>` je `LogFatal`. Pisac diže grešku, test ga je zvao direktno.
    Oba leka iz memorije primenjena zajedno: poziv od kog se očekuje uspeh ide
    kroz `On Error Resume Next`, a ciljane tvrdnje su se popele **iznad** rane
    izlazne tačke.
    **Četiri zatečene procedure nisu bile zastarele.** Dve mere **B10**
    (`modIntegritet`) — **produkcioni** čekar koji je po konstrukciji stari oblik —
    a dve mere **stari storno ekran**, koji radi po nogama. Oboje živi do
    `10c`/`10e` i mora da ostane merljivo, inače bi čekar koji još radi u
    produkciji ostao bez ijednog testa. Zato seju stari oblik **same**
    (`SejRevStariOblik` → `TrackAmbalaza`), a tvrdnje su im **netaknute**.
    Ostaje ručna kapija: compile + `--mark-compile`. Ona je ovde teža nego obično —
    rez je menjao **potpis** `SaveOMUlaz_TX` i preselio 19 pozivnih mesta, a to je
    tačno klasa koju samo compile hvata sam.
70. **P1: auto broj reversa je imao DVA izvora** (06.10.2026, review `1cc8a186`).
    UI prefill je čitao **stari** oblik (`MaxSeqReversAmbalaza` nad
    `tblAmbalaza`), a pisac je upisivao **zaglavlje**. Nov revers u
    `tblAmbalaza` više nema poslovni broj — tamo stoji `AmbDokID` — pa drugi F7
    istog dana dobije **opet prvi broj** i padne tek na upisu.
    **Reprodukuje se na PRAZNOJ instalaciji, na drugom reversu.** To je bio
    funkcionalni blocker, ne rupa u dokazu.
    **Lek nije spajanje dva niza, i tu sam grešio u prethodnom commit-u.** Vratio
    sam bio `RequireBrojSlobodanUNizu` uz obrazloženje „dok oba oblika postoje, oba
    niza moraju da važe". Program **kreće od nule**: nema podataka koje treba
    pomiriti, pa finalni proizvod ne treba da plaća cenu dvomodelne numeracije.
    `KIND_REV` je zato preveden na **kanonski** niz (`AmbDokNizSken` /
    `AmbDokBrojZauzet`), a `MaxSeqReversAmbalaza` i `BrojZauzetRevers` su
    **obrisani** — sa njima i dva komentara koja su ih pominjala. Duplikat kapije u
    piscu je otpao: `UpisiAmbDokument` sudi istu činjenicu (`ODL-20`).
    Format broja se **ne menja** — `GenerateBrojAmbDokumenta` koristi isti
    `FormatBroj(stanica, datum, seq)` — pa operater vidi isto što i pre.
    **Acceptance koji je falio** je dodat (`Test_REV_AutoBrojJedanNiz`): predlog →
    upis → predlog → upis, istog dana i iste stanice; drugi predlog **mora** da se
    pomeri, oba upisa prolaze, i storno **ne** vraća predlog unazad (A9). Sabotaža
    `amb-rev-broj-iz-pogresnog-niza` gasi baš čitanje kanonskog niza; tvrdnja je
    **razlika dva predloga**, ne uspeh upisa — upis bi svejedno pao na kapiji
    zauzetosti, pa bi tvrdnja o ishodu bila zelena i sa kvarom.
    Usput zatvoren **P3**: `6.12h` je i dalje tvrdio da revers nema storno pisca, a
    zaglavlje `modNovacUnos` da F7 ide na `SaveOMUlaz_TX`. Oba ispravljena.
    **P2 (ekran Storno još ne vidi nov revers) ostaje otvoren i ide u `10c`** — ne
    kao kompatibilnost sa starim modelom nego kao prelazak čitaoca: `STIP_REVERSI`
    na `tblAmbalazaDokument` + `AmbDokID`, a stari put nestaje. Grana se **ne
    mergeuje** bez toga. Katalog 692 → 693.
71. **Rez reversa zatvoren: DOKAZANO 4/4 + pun prolaz ZELENO** (06.10.2026,
    izvor `6e2767e18f87`). `RunAllTests` 200 · `RunBusinessFlowProSuite` **2394** ·
    `RunStornoTestSuite` 164 · banka 241 — **12/12, nula padova**. `dokaz.py`
    `crvenih 4/4`, potpis `4c9d94988cbb1c63` identičan pre i posle.
    **Put do zelenog je dao još dva nalaza, oba u mom kodu.**
    **(1) `StornirajAmbDokument_TX` je zvao TUĐ `Private` simbol** —
    `MarkRowStornirano` iz `modStorno`. VBA kompajlira **na zahtev**, pa je
    `Sub or Function not defined` puklo tek kad je prvi test pozvao baš tu
    proceduru: posle 585 s i ubijenog Excela, uz poruku **bez fajla i linije**.
    Tuđa privatnost nije otvarana — primitiv (`RequireUpdateCell`) se zove direktno,
    jer je `MarkRowStornirano` ionako samo njegov omotač. Napisan je merač za celu
    klasu; **dvosmeran dokaz: 1 nalaz sa fajlom i linijom, 0 posle**.
    **(2) Tri tvrdnje su posle prelaska na kanonski niz ostale da mere stari
    izvor** — u testovima koji seju stari oblik za B10 i stari storno ekran.
    Uklonjene su **odatle**, ne oslabljene: `A9` nad nizom mere dva testa kroz
    pravog pisca. Dve tvrdnje o istom pravilu nad dva izvora bi se razišle — a to
    je i bio ceo P1.
    **Tri lažna nalaza merača su i sama merenje:** repni komentar, labela kao cilj
    skoka, i **LF kopija iz git-a** — `git show` vraća blob sa LF, a merač je delio
    po `
` i dobio ceo fajl kao jedan red, pa je prvi prolaz dvosmernog dokaza
    bio **lažno čist**. Nalaz u **merenju**, ne u meraču — i razlog zbog kog se
    dvosmeran dokaz uopšte radi.
    Ostaje ručna kapija (compile + `--mark-compile`) i **P2 iz review-a**: ekran
    Storno još ne vidi nov revers — ide u `10c`, i grana se **ne mergeuje** bez
    toga.
72. **Uplata kupca — poslednje presečeno mesto knjiženja (red 8)** (06.10.2026, 6.12i).
    `SaveKupciIzlaz_TX` je ostao **samo kasa**: nestali su `vozacID`, `tipAmb`,
    `kolAmb`, snapshot `TBL_AMBALAZA` i noga u knjizi; kapija
    `kolAmb <= 0 And novac <= 0` postala je `novac <= 0`, a poruka koja je
    imenovala ambalažu zamenjena je novčanom u **oba** pisca — `SaveOMUlaz_TX`
    je istu zastarelu rečenicu nosio od svog reza. `DOK_TIP_IZLAZ_KUPCI` je
    **obrisan**: jedan pisac, nula čitalaca.
    **Ovde se knjiženje nije preselilo nego je prestalo** — povrat praznih od
    kupca ima svoje mesto u redu 7 (prijemnica, pod **njenim** brojem, `ODL-9/-10`
    uz `ODL-22`). Dva reda za isti događaj bila bi dva traga.
    **Tri merenja pre koda:** ambalažna noga nije imala **nijednog** produkcionog
    pozivaoca (F6 šalje `kolAmb:=0` tvrdo upisano) · `DOK_TIP_IZLAZ_KUPCI` nije
    imao **nijednog** čitaoca · posle reza `TrackAmbalaza` nema **nijednog**
    produkcionog pozivaoca — čime je **write-side deo `10b-2` zatvoren**: svih
    devet mesta knjiženja piše nov oblik, stari pisac je još samo test-alat.
    **⚠ CAPABILITY, zapisano a ne zatvoreno:** „kupac vraća prazne bez dostave
    robe" — vrsta `AMB_DOK_REVERS_PARTNERA` i kapija postoje, **pisca nema**, a
    ulaza nema ni u legacy-ju od §27.18. Poslovno pitanje za operatera; pisac se
    ne izmišlja pre odgovora.
    Usput popravljeno: **rep stavke 60** (pun prolaz za 59+60) stajao je od
    `18dfc330` na **kraju** sekcije — isti kvar kao rep stavke 53, i to iz **istog
    commit-a koji taj rep popravlja**. Vraćen pod 60. Hronologija je izmerena cela:
    to je bio **jedini** blok posle dva ili više praznih redova (857 redova).
    Jeftine kapije `rc=0`: `vba_check` (191 fajl, 695 sabotaža, 0+10 poznatih),
    `gen_schema_module --check`, `who_writes --check` / `--check-ownership` /
    `--self-test`, `popis_citalaca --check`, arnost 321, scope 5545, nastavak
    195908 redova, privatno 191 fajl. **Skupe kapije čekaju reviewer GO.**
73. **Presuda operatera: povrat praznih od kupca postoji i BEZ prijemnice**
    (06.10.2026, `AMB-10-ODL-23`, 6.12i). Pitanje otvoreno u stavci 72
    odgovoreno je isti dan, i odgovor je **da — redovno**. Dva pod-pitanja su
    namerno postavljena kao „potvrdi ili ispravi kapiju", jer je kod već nosio
    pretpostavku: oba odgovora su je **potvrdila** — dokument nosi **kupčev**
    broj (`REVERS_PARTNERA`, `BrojOwnerTip = Kupac`), a gajbe idu **na vozača**
    (`Kupac → Vozac`, `POVRAT_PRAZNE`). Grana `jePartnerov` u
    `AmbDokKretanjeProblem` traži baš taj oblik, pa **ugovor se ne menja**.
    Dve posledice se čitaju iz presude, ne biraju: broj se **ne predlaže** (iz
    našeg niza bio bi izmišljen broj tuđe serije; zauzetost u opsegu
    `(Kupac, KupacID, dan)`), i peti smer **ne ide** u `AmbReversSmerovi` — ta
    mapa je mapa **našeg** reversa sa staničinim brojem.
    Rez koji sledi nosi zato **samo pisca i ulaz**: `REVERS_PARTNERA` prestaje da
    bude vrsta bez pisca, storno je već pokriven `StornirajAmbDokument_TX`, a F7
    prestaje da važi u delu „ne prima kupca kao partnera".
    **Zabeleženo kao merenje, ne kao pohvala:** ugovor napisan u `10a` izdržao je
    domaće pitanje koje mu je postavljeno **pet dana kasnije**, dok je isti ugovor
    u `ODL-22` pukao na premisi. Razlika je u tome što je ovde kapija merila
    **vlasnika broja** (činjenicu), a tamo **vrstu dokumenta** (zamenu za pravilo).
74. **Rez reda 8 zatvoren: DOKAZANO 2/2 + pun prolaz ZELENO** (06.10.2026).
    `dokaz.py --grupe 6 amb-kup-`: **crvenih 2 / sabotaža 2**, potpis izvora
    `ab359375960e4353` **identičan pre i posle**. Grupisanje je dalo **2 prolaza
    nad 2 sabotaže**, pa je ovo **pun pojedinačni dokaz**, ne grupni — verdikt je
    `DOKAZANO`, bez „(grupno)".
    Baza: `RunAllTests` 200/0, `RunBusinessFlowProSuite` **2399**/0.
    Pun prolaz `run_vba.py`: **ZELENO**, 12/12 suita, nula padova
    (`RunAllTests` 200, BFP 2399, `RunStornoTestSuite` 164, banka 241, palete,
    faktura, agrohemija, Sheets, golden, licenca, StornoCentar, izveštaji).
    GREEN marker: izvor `6baf27d0ae8a`, ugovor `a10b23d003d7`, sveska
    `otkup_test.xlsm/d24883a3`. Marker **pokriva i tekući HEAD**, jer je stavka 73
    (`d190041c`) bila **samo docs** — potpis izvora se nije promenio.
    **2394 → 2399 je tačno pet novih tvrdnji** — koliko ih `Test_KUP_UplataJeSamoNovac`
    i ima. Neobjašnjena razlika bi značila da je test prećutao deo sebe
    (v. „broj tvrdnji je merenje"); ovde se poklapa po stavci.
    Compile je kao i uvek `NEJASNO` (nema dijaloga) — ručna kapija stoji, i sad
    pokriva **dva reza**: revers (stavka 71) i ovaj. Oba su menjala **potpis**, a
    ovaj je i obrisao konstantu `DOK_TIP_IZLAZ_KUPCI`.
    Ostaje nepokriveno i zapisano: **P2 iz review-a** (ekran Storno ne vidi nov
    revers — `10c`, pred merge) i **pisac + ulaz za `AMB-10-ODL-23`** (stavka 73).
75. **Pisac kupčevog reversa** (06.10.2026, `AMB-10-ODL-23`, 6.12j).
    `modAmbalaza.UpisiReversPartnera_TX` — povrat praznih od kupca **bez
    prijemnice**: vrsta `REVERS_PARTNERA`, vlasnik broja `Kupac`/`KupacID`, par
    `Kupac → Vozac`, kretanje `POVRAT_PRAZNE`.
    **Pisac nosi tačno jedno pravilo koje nigde drugde ne postoji: broj je
    obavezan i NE predlaže se.** Sve ostalo je već bilo u jezgru — par i vlasnik
    broja sudi `AmbDokKretanjeProblem` (grana `jePartnerov`), vrstu kretanja
    `AmbDokDozvoljavaKretanje`, zauzetost `UpisiAmbDokument` u opsegu
    `(Kupac, KupacID, dan)`, identitet i saldo `PrenesiAmbalazu`
    (`AMB-INV-04`, `-07`), a **storno je postojao od reza reversa**
    (`StornirajAmbDokument_TX` radi nad svakim ambalažnim dokumentom). Zato je rez
    mali: ugovor je bio napisan pet dana pre pitanja koje ga je potvrdilo.
    **Peti smer nije dodat u `AmbReversSmerovi`** — ta mapa je mapa **našeg**
    reversa (vlasnik broja je stanica), pa bi peti red tiho uveo dokument sa
    tuđim brojem i drugom vrstom u mapu koja o njima ne zna ništa.
    **Tvrdnje o zaglavlju stoje IZNAD rane izlazne tačke** — `LookupValue` nad
    praznim `dokID`-em vraća `""`, pa pucaju **po imenu** i kad sabotaža obori ceo
    upis. Da stoje ispod, sabotaža vlasnika broja obarala bi samo tvrdnju da je
    dokument upisan — **posledicu, ne pravilo** (klasa zatvorena 05.10.2026).
    **⚠ Ulaza još nema:** F7 odbija kupca kao partnera, pa `popis_citalaca` pisca
    vidi kao `SAMO_TEST` — tačan opis stanja, ne propust zapisa. Ulaz je sledeći
    korak. Uz njega ide i **zajednički dug**: prekomeran povrat traži protokol
    potvrde manjka, koji danas prosleđuje **samo otkup** — isto važi za prijemnicu
    od 6.12g, pa parametar ne uvodi ovaj pisac sam.
    Testovi `Test_RVP_KupcevDokumentJedanRed`, `Test_RVP_BrojJeKupcevINePredlazeSe`;
    sabotaže `amb-rvp-broj-se-predlaze`, `amb-rvp-vlasnik-broja-nije-kupac`.
    Katalog 695 → 697. Jeftine kapije `rc=0`: `vba_check` (191 fajl, 697 sabotaža,
    0+10), schema, `who_writes` ×3, čitaoci, arnost 323, scope 5550, nastavak
    196149, privatno 191. **Skupe kapije čekaju reviewer GO.**
76. **Pisac `ODL-23` dokazan: DOKAZANO 2/2, ciljana suite ZELENA** (07.10.2026).
    Redosled je bio namenski: **ciljana BFP suite PRVA**, pa `dokaz.py` samo ako
    je zelena — pisac i njegova dva testa dotad nisu bili izvršeni ni jednom, a
    dokaz nad crvenom bazom meri crveno koje sabotaža nije izazvala.
    `RunBusinessFlowProSuite` **2417**/0 (210 s). `dokaz.py --grupe 6 amb-rvp-`:
    **crvenih 2 / sabotaža 2**, potpis izvora `4bf0a36b54e84179` **identičan pre i
    posle**, 2 prolaza nad 2 sabotaže — dakle **pun pojedinačni dokaz**, verdikt
    `DOKAZANO` bez „(grupno)".
    **2399 → 2417 je tačno osamnaest novih tvrdnji**, koliko ih dva testa i nose
    (10 + 8). Poklapanje po stavci je jedini način da se vidi da nijedan test nije
    prećutao deo sebe (v. „broj tvrdnji je merenje").
    `amb-rvp-vlasnik-broja-nije-kupac` oborila je uz svoju tvrdnju i **pet drugih**
    — kaskada iz palóg upisa, i to je očekivano: zato je tvrdnja o vlasniku broja
    postavljena **iznad** rane izlazne tačke, da crveno ne bude samo posledica.
    Ostaje: ručna kapija (compile + `--mark-compile`, **tri reza**), **ulaz** za
    ovog pisca (F7 odbija kupca kao partnera), i **P2** — ekran Storno ne vidi nov
    revers, `10c`, pred merge.
77. **Review `024995de`: CODE GO, P0 = 0, P1 = 0** (07.10.2026).
    Reviewer je prešao oba reza (`SaveKupciIzlaz_TX` → samo novac;
    `REVERS_PARTNERA` pisac) i **nije našao nijedan P1 u implementiranom kodu**.
    Potvrđeno kao ispravno: razlaganje F6, brisanje `DOK_TIP_IZLAZ_KUPCI`,
    odluka da RVP **nije peti smer** u `AmbReversSmerovi`, hard-fail na prazan
    broj, TX granica bez self-bind rupe, i acceptance koji meri **obe** strane
    (novac upisan **i** knjiga nedirnuta) — „test ne može lažno da pozeleni nad
    potpuno mrtvim writerom".
    Nezavisno merenje koje je reviewer dodao: `WHO_WRITES` sada pokazuje **samo
    `modAmbalaza` i `modStornoRecovery`** kao produkcione mutatore `tblAmbalaza`,
    a `modStornoRecovery` je legacy recovery koji **produkciono dugme odbija**.
    **Tri otvorene stavke iz review-a:**

    | | Šta | Status |
    |---|---|---|
    | P2 #1 | ekran Storno ne vidi `AmbDok` revers — ni naš ni kupčev | **merge blocker**, `10c` |
    | P2 #2 | protokol potvrde deficita: `UpisiReversPartnera_TX` ne može da **primi** potvrđen manjak, pa nema načina da pozivalac ponovi upis | mora **pre** produkcionog ulaza; **zajednički** sa prijemnicom, jedan mehanizam |
    | P3 | RVP acceptance ne meri direktno količinu i salda | širi se zajedno sa deficit scenarijem |

    Evidence dug koji reviewer imenuje: **punih 12/12 nije ponovljeno posle RVP
    commit-a** (ima ciljanu BFP 2417/0 + `dokaz` 2/2), i compile je još
    `NEJASNO`. Ne zaustavlja razvoj, ali stoji pred merge.
    **NALAZ U SUSEDNOM KODU, nađen pri čitanju za P2 #2** (nije moj, nije iz
    review-a): u `SavePrijemnicaMulti_TX` `kolAmbVracena` ide **samo** pozivu za
    Klasu I, a Klasa II dobija tvrdo upisanu `0`. Klasa I je **opciona**
    (`kolicinaI = 0` → snima se samo Klasa II), pa prijemnica sa samo Klasom II i
    vraćenim praznim gajbama **tiho gubi nogu povrata** — bez ijedne poruke.
    Dostupno iz F4: `kolicinaI` i `kolAmbVracena` dolaze **nezavisno**
    (`modDokUnos.PrijemnicaUpisi`). Utvrđeno **čitanjem**, ne pretpostavkom:
    argument je doslovna nula. Ide u isti rez kao P2 #2, jer se dira ista noga.
78. **P2 #2 zatvoren: zajednički protokol potvrde deficita** (07.10.2026, 6.12j).
    Reviewer je pobio moju odbranu iz stavke 75 — i bio je u pravu. Napisao sam
    da parametar `potvrdaDeficita` „ne uvodi ovaj pisac sam", jer bi bio bez
    pozivaoca. Ali `ODL-23` je sposobnost definisao kao **redovnu**, a pisac koji
    potvrdu ne može ni da **primi** nema samo strožu kapiju — on ima **nedostižnu
    poslovnu putanju**: pozivalac nema čime da ponovi upis.
    Protokol je proširen na **oba** pisca, i to **isti** protokol, po obrascu koji
    `CreateOtkup_TX` nosi od 6.5: `UpisiReversPartnera_TX`, `SavePrijemnica`,
    `SavePrijemnica_TX` i `SavePrijemnicaMulti_TX` dobijaju `potvrdaDeficita`, a
    `Multi_TX` i **`outErrNum`** — bez njega F4 dobija tekst greške ali ne i broj,
    a broj je ugovor (tekst je prevodiv). Potvrda **ne ide u log**: kupac koji
    vrati više nego što knjiga kaže je redovan slučaj, a log koji ga beleži kao
    kvar prestaje da bude signal.
    Potvrda ide **samo na nogu povrata**: puna noga polazi od vozača, a on je
    `SOPSTVENI` — njegov manjak je po `ODL-8` **tvrdo** odbijen i potvrda tamo ne
    postoji.
    **NALAZ U SUSEDNOM KODU (nije iz review-a):** u `SavePrijemnicaMulti_TX` je
    `kolAmbVracena` išla **samo** pozivu za Klasu I, a Klasa II je dobijala tvrdo
    upisanu `0`. Klasa I je **opciona**, pa je prijemnica sa samo Klasom II i
    vraćenim gajbama **tiho gubila nogu povrata** — bez ijedne poruke, a iz F4
    dostupno (`kolicinaI` i `kolAmbVracena` dolaze nezavisno). Povrat je **jedan
    događaj**, pa sada ide uz **dokument koji postoji**.
    **Usput naučeno o samim tvrdnjama:** prva verzija je ciljala tvrdnju
    `"... PROLAZI (" & razlog & ")"` — a `dokaz.py` se poklapa po **doslovnom**
    literalu, pa bi sabotaža javila `NE OBARA SVOJ TEST`. Razlog je zato dobio
    **svoju** tvrdnju (`AssertEquals "", razlog`), koja ga ispisuje kad padne.
    `P3` iz review-a je zatvoren istim testom: acceptance sada meri **količine i
    salda** (povrat 20, pokriće 15, kupac 0, vozač +20), nad **svežim** tipom
    ambalaže — nad zajedničkim bi saldo nosili i drugi testovi, pa manjak ne bi
    bio ponovljivo 15.
    Testovi `Test_RVP_DeficitSePotvrdjuje`, `Test_PRJ_PovratIdeSaKlasomKojaPostoji`;
    sabotaže `amb-rvp-potvrda-se-ne-prosledjuje`, `amb-prj-povrat-samo-sa-klasom-i`.
    Katalog 697 → 699. Jeftine kapije `rc=0` (`vba_check` 191/699/0+10, schema,
    `who_writes` ×3, čitaoci, četiri merača).
    Ostaje za ulaz: F7 i F4 moraju da **pitaju** operatera i ponove poziv —
    produkcioni uzorak je `modOtkupUnos` (prepoznaje slučaj po **broju** greške,
    zadržava podatke na ekranu, ponavlja poziv).
79. **P2 #2 dokazan, i evidence dug iz review-a zatvoren** (07.10.2026).
    Četiri koraka u jednom lancu, svaki sa branom na prethodni:

    | | Korak | Ishod |
    |---|---|---|
    | 1 | ciljana `RunBusinessFlowProSuite` | **2436**/0 |
    | 2 | `dokaz.py --grupe 6 amb-rvp-potvrda` | crvenih **1/1**, `DOKAZANO` |
    | 3 | `dokaz.py --grupe 6 amb-prj-povrat-samo` | crvenih **1/1**, `DOKAZANO` |
    | 4 | pun prolaz `run_vba.py` | **ZELENO**, 12/12 suita, nula padova |

    Dva `dokaz` poziva jer su prefiksi različiti, a alat prima **jedan** filter;
    potpis izvora `aec7caf100818d2b` **identičan** pre i posle oba.
    Pun prolaz: `RunAllTests` 200/0, BFP 2436/0, `RunStornoTestSuite` 164/0,
    banka 241/0, i ostalih osam. GREEN marker nad izvorom `c410e67318a5`
    (ugovor `a58367524387`). **Time je zatvoren evidence dug koji je reviewer
    imenovao** — „punih 12/12 nije ponovljeno posle RVP commit-a" — i to nad
    izvorom koji nosi **i** RVP pisca **i** protokol potvrde.
    **2417 → 2436 je tačno devetnaest novih tvrdnji** (14 + 5), koliko ih dva
    testa i nose. Četvrti put u ovom rezu da se broj poklopi po stavci; da nije,
    značilo bi da je test prećutao deo sebe.
    Compile je i dalje `NEJASNO` (nema dijaloga) — ručna kapija sada pokriva
    **četiri reza**: revers, red 8, RVP pisac i protokol potvrde.
    Od review-a `024995de` ostaje **jedna** stavka: `P2 #1`, ekran Storno ne vidi
    `AmbDok` revers (`10c`, merge blocker). `P2 #2` i `P3` su zatvoreni.
80. **Ulaz za `AMB-10-ODL-23`: peti segment na F7** (07.10.2026, 6.12j).
    Odluka gde ulaz živi bila je otvorena — **nov ekran** ili **peti smer na
    F7** — i izabran je peti smer, ne zbog štednje nego zbog **oblika koji ekran
    već ima**: F7 od početka prebacuje politiku po smeru (1–2 traže kooperanta,
    3–4 vozača i nikakvog partnera). „Smer 5 traži kupca i ručno upisan broj" je
    nastavak istog oblika, bez novog ekrana, F-tastera i reda u registru.
    **Partnerska lista se nije menjala:** `PartnerSrcOrder("F7")` je već nosila
    `KUP` — kupci su na F7 postojali, samo ih je validator odbijao. Izmereno pre
    koda; prvo sam planirao refill liste po smeru, i to je bilo nepotrebno.
    Geometrija: pet segmenata po **73pt** umesto četiri po 91 (`1 + 5*73 + 4*1 =
    370`), isti okvir. Najduži natpis ostaje 12 znakova.
    **Šta ulaz nosi:** partner mora biti kupac · vozač obavezan i bez
    `VALIDACIJA_UNOSA` · **broj obavezan i bez predloga** (grana izlazi **pre**
    auto-broja) · zauzetost u opsegu `(Kupac, KupacID, dan)` · upis ide
    `UpisiReversPartnera_TX` · štampa imenuje **kupca**, ne vozača.
    **Dva mesta namerno ostavljena:** `SmerRevKljuc(5)` vraća `""` (prevod bi
    značio da `AmbReversSmerovi` peti smer ipak poznaje), a `ZavrsiIspravkuAko` se
    ne zove — tok ispravke ključa po `(broj, stanica, dan)`, pa bi mogao da
    zatvori **tuđu** ispravku sa slučajno istim brojem.
    **Protokol potvrde deficita dobio je UI na oba mesta** (F7 i F4): hvata se
    **broj** greške, manjak se čita **svež**, operater se pita, poziv se
    **ponavlja** — ekran zadržava podatke. Obrazac prepisan iz `modOtkupUnos`.
    Na F4 je `SetPaletizeSkip False` **pomeren ispod** ponovnog poziva: između dva
    pokušaja mora da ostane uključen, jer je ispravka ista roba.
    **⚠ Nijedan test ne sme da uđe u tu granu** — `MsgBox` u `run_vba` prolazu
    visi do timeout-a i ostavlja Excel u `[break]`. Testovi zato mere **pisca**,
    a dijalog ide u operatersku ček-listu. Isti rizik nosi otkup od 03.10.2026;
    ovo ga ne uvodi, ali ga sada nosi **tri** mesta — upisano kao dug.
    **Zatečena sabotaža je oborila kapiju, i to je dobro:** `revers-smer` je
    sidrila opseg `smer > SMER_REV_PRI_OM`, koji je ovaj rez promenio — `KATALOG`
    je javio „sidro ZASTARELO" (0 pogodaka). Osveženo bez menjanja tvrdnje.
    `RunAllTests` sada ima **201** test (nov: `T_ReversValidiraj_PovratKupcaJeSvojSmer`,
    koji **sam uključuje** `AUTO_BROJ_DOKUMENTA` — bez toga bi tvrdnja „broj je
    ostao prazan" bila zelena i kad je predlog iskqučen u Podešavanjima, pa ne bi
    merila granu nego konfiguraciju).
    Sabotaže `amb-ulaz-kupcev-broj-se-predlaze`, `amb-ulaz-kupcev-smer-prima-kooperanta`.
    Katalog 699 → 701. Jeftine kapije `rc=0` (`vba_check`, schema, `who_writes`
    ×3, čitaoci, četiri merača, popis suita).
81. **`REVERT-FAIL`: dve sabotaze sa ISTIM pokvarenim tekstom** (07.10.2026).
    Dokaz ulaza je stao posle prve sabotaže: `amb-ulaz-kupcev-smer-prima-kooperanta`
    → `REVERT-FAIL`, `izvor pre/posle RAZLIKA`, i radno stablo je ostalo
    **pokvareno** — u KUP grani je pisalo `partTip <> "KOOP"`.
    **Mehanizam:** revert traži **svoju** zamenu i na njeno mesto vraća **svoje**
    sidro. Moja sabotaža je imala **isti** pokvaren tekst kao zatečena
    `revers-kupac` — dve grane istog validatora, razlika samo u `"KUP"`/`"KOOP"`,
    a komentar sabotaže prepisan — pa je revert u KUP granu upisao **tuđe** sidro.
    Izvor time ostaje **zdrav po obliku a pogrešan po sadržaju**: to je najgora
    vrsta ostatka, jer ga nijedna sintaksna provera ne vidi.
    **Nijedna zatečena provera to nije mogla da vidi:** sidro je bilo jednoznacno,
    zamena odsutna u zdravom izvoru, tvrdnje različite. Katalog je imao zamke za
    prazan tekst, zamenu jednaku sidru, zamenu kao podniz sidra, deljenu tvrdnju,
    dodelu tuđoj proceduri — ali ne za **deljenu zamenu**.
    **Zamka 11** je zato napisana: dva unosa nad istim fajlom sa istim
    pokvarenim tekstom → nalaz po imenu oba. Dvosmeran dokaz nad **pravim**
    katalogom (`CLAUDE.md` §5 ga za izmenu checkera i zahteva):

    | | Ishod |
    |---|---|
    | pre razdvajanja | **crvenih 3**, svaki imenuje svoj par |
    | posle razdvajanja | `nalaza 0`, `rc=0` |
    | `--self-test` | 42 → **43** slučaja, čisto |

    **Dva od tri para bila su ZATEČENA** — i to je ono što pravilo opravdava:
    `izmena-nacrta-pravi-nov` / `zbirna-ekran-izmena-pravi-nov` (ista zamena,
    sidra `mIzmenaOtpID` vs `mIzmenaZbrID`) i `otp-kapija-mreza-tiha-nula` /
    `otk-kapija-mreza-tiha-nula` (ista zamena, sidra `ZbirStavkiZaOtpremnicu` vs
    `ZbirStavkiZaOtkup`). Oba bi pri revertu upisala **tuđu** granu, u istom
    obliku kao moj slučaj; nisu pukla samo zato što ih nijedan rez nije pustio
    zajedno. Razdvojeni su **komentarom**, koji je inertan — šta sabotaža meri
    nije dirnuto.
    **Self-test nove zamke tvrdi i šta se NE sme upaliti:** par deli zamenu a
    tvrdnje su različite, pa pravilo o deljenoj tvrdnji mora da ostane tiho —
    inace bi self-test prolazio i bez zamke 11.
    Usput izmereno: `tools/sabotaza.py` je **CRLF** fajl, a moji ranije ubacivani
    blokovi su išli sa `LF` (119 samotnih LF-ova). Python to ne vidi, ali anchor
    građen sa `\n` **ne pogađa** — prva dva pokusaja patch-a su zato javila
    „0 pogodaka". Patch skripte za taj fajl grade redove iz `N = "\r\n"`.
    `git checkout` za vraćanje ostatka je bio **odbijen** (destruktivna radnja),
    pa je red vraćen običnom izmenom izvora — ista vrednost, vidljiv trag.
82. **Ulaz dokazan: DOKAZANO 2/2 + pun prolaz ZELENO 12/12** (07.10.2026).
    Posle razdvajanja zamena (stavka 81) dokaz je prošao iz prvog puta:
    `crvenih 2 / sabotaža 2`, potpis izvora `b031cbdeab622872` **identičan pre i
    posle** — isti potpis kao u palóm prolazu, što i potvrđuje da je ostatak bio
    u **revertu**, ne u mojoj izmeni.
    Pun prolaz: **ZELENO**, 12/12 suita. `RunAllTests` **201**/0, BFP 2436/0,
    `RunStornoTestSuite` 164/0, banka 241/0. GREEN marker nad izvorom
    `1fe99d7bc980` (ugovor `57412040ddf9`).
    `RunAllTests` 200 → 201 je **jedan nov test**, a BFP je ostao 2436 — nov test
    je otišao u `modTest`, ne u BFP, pa se oba broja poklapaju sa onim što je
    dodato.
    **ČEK-LISTA ZA OPERATERA** — ovo se ne meri automatski (`CLAUDE.md` §5):

    | Šta proveriti | Gde |
    |---|---|
    | pet segmenata smera stoji u jednom redu, bez preklapanja i bez odrezanog natpisa | F7, polje „Smer reversa" |
    | izbor „Povrat kupca" nudi **kupce** u listi partnera i prima ih | F7 |
    | broj se **ne** popuni sam kad je izabran „Povrat kupca" | F7 |
    | dijalog potvrde manjka ponovi upis **sa zadržanim podacima** | F7 i F4 |
    | štampani kupčev revers imenuje **kupca**, ne vozača | F7 → PDF |

    Ručna kapija compile sada pokriva **pet rezova**: revers, red 8, RVP pisac,
    protokol potvrde i ulaz. Od review-a `024995de` ostaje jedino `P2 #1` —
    ekran Storno ne vidi `AmbDok` revers (`10c`, merge blocker).
83. **Operaterska provera ulaza: geometrija prolazi, nov ključ poruke traži
    RESTART** (07.10.2026).
    Prva stavka ček-liste je **potvrđena na ekranu**: pet segmenata smera stoji u
    jednom redu, jednake širine, bez preklapanja i bez isečenog natpisa
    („Prijem od OM" i „Povrat kupca" se vide celi). Geometrija 73pt radi.
    **Nalaz usput, i koštao je operatera vremena:** peti segment je prvo pisao
    `[OTKUI_SEG_REV_POVRAT_KUP]` — fallback `Poruka()` za ključ kog **nema u
    `tblPoruke`**. Ostala četiri su bila ispravna, što je odmah isključilo
    geometriju i gradnju forme kao uzrok i pokazalo na tabelu poruka.
    **`EnsurePoruke` iz Immediate prozora nije pomogao** — i to je mehanizam koji
    vredi zapamtiti: on **puni tabelu**, ali natpisi runtime kontrola se **peku u
    trenutku gradnje ljuske** (`NewSegBtn ... Poruka("KLJUC")`). Već izgrađeno
    dugme zadrži stari tekst. Rešenje je **zatvoriti i otvoriti fajl**:
    `modMain.InitApp` zove `EnsurePoruke` na svakom startu **i** gradi ljusku
    iznova — oboje, a potrebno je oboje.
    **Produkciju ne pogađa:** self-update koda ionako restartuje aplikaciju. Ovo
    je zamka razvojne petlje (uvoz koda u otvoren Excel), ne isporuke.
    Pravilo po sadržaju pripada `.claude/rules/forme-i-kontrole.md`, ali `.claude/`
    ide **isključivo kroz zaseban process PR** (`CLAUDE.md` §6), pa ovde stoji
    nalaz, a preseljenje je zasebna stavka.
84. **Operaterska provera našla P1 u ulazu: ljuska je kupčevom reversu
    predlagala NAŠ broj** (07.10.2026).
    Na ekranu je, uz izabran smer „Povrat kupca", u polju **BROJ REVERSA**
    stajalo `1/071026` — broj iz **našeg** niza `(Stanica, dan)`. Da je upis
    prošao, kupčev dokument bi nosio broj koji smo mi izmislili — tačno ono što
    `AMB-10-ODL-23` zabranjuje. Pisac bi ga primio: on traži da broj **postoji**,
    ne da je kupčev.
    **Uzrok je sloj, ne pravilo.** Kapiju sam stavio u `ReversValidiraj`, a broj
    stiže iz **ljuske**: `RefreshBrojPredlog` ga upiše čim se izabere otkupno
    mesto — dakle **pre** nego što se smer uopšte bira. Komentar u
    `modScrDokumenti.SaveRevers` je to i govorio („Broj reversa se predlaže u
    ljusci"), a ja sam ga pročitao tek kad je slika pokazala broj.
    **Isti oblik već postoji u istom fajlu:** `PredlogPrijemnice` — „Ostali kupci
    nose svoj eksterni, nezavisni broj — polje se tada **NE dira**". Repo je
    pravilo znao; moj ulaz ga nije sledio.
    **Ispravka je selidba odluke, ne druga kopija kapije:**
    `modNovacUnos.RevSmerPredlazeBroj(smer)` je sada **jedini izvor**, a zovu je
    **oba** sloja — ljuska pre nego što dodirne polje, validator pre auto-broja.
    Dve kopije istog uslova bi se razišle; prvi put su se i razišle.
    Uz to `SetSmerRev` **prazni** polje na prelasku na kupčev smer i **vraća**
    predlog na povratku — broj prati smer kao što već prati stanicu.
    **Zasto ga nijedan test nije uhvatio:** svi su merili `ReversValidiraj`, a
    kvar je bio u ljusci. Sada tvrdnja meri **funkciju koju oba sloja zovu**, pa
    jedna sabotaža (`amb-ulaz-predlog-ne-gleda-smer`) obara oba puta. Katalog
    701 → 702.
    **Cena je izmerena i vredi je zapisati:** ulaz je prošao `DOKAZANO 2/2` i pun
    prolaz 12/12 **sa ovim kvarom u sebi**. Zelena suite ne pokriva sloj koji
    nijedan test ne dodiruje — operaterska provera je ovde bila **jedina** kapija,
    i zato stoji u ček-listi, ne kao formalnost.
85. **Ispravka predloga dokazana: DOKAZANO 3/3 + pun prolaz ZELENO** (07.10.2026).
    `dokaz.py --grupe 6 amb-ulaz-`: **crvenih 3 / sabotaža 3** (dve zatečene plus
    nova `amb-ulaz-predlog-ne-gleda-smer`), potpis izvora `8676e5910540691d`
    **identičan pre i posle**.
    Pun prolaz: **ZELENO**, 12/12. `RunAllTests` 201/0, BFP 2436/0, Storno 164/0,
    banka 241/0. GREEN marker nad izvorom `366c3083ee2e` (ugovor `f357cc70ea15`).
    Broj testova se **nije** menjao (201) ni broj BFP tvrdnji (2436) — tvrdnje su
    dodate **postojećem** testu, pa se poklapa i to.
    Ostaje operaterska potvrda baš te putanje: izaberi otkupno mesto (broj se
    popuni), pa „Povrat kupca" — **polje mora da se isprazni**.
86. **`AMB-10-ODL-24` otvoren: pozajmica ambalaže od kupca** (08.10.2026, 6.12k).
    Operater je, posle potvrde da ulaz radi, imenovao **recipročan** smer: kupci
    **često pre sezone predaju SVOJE prazne gajbe** — pozajmica nama, ne povrat
    naših — i to **vozaču**, istim lancem. „Time se zaokružuje celina."
    **Nisam krenuo u kod, i to je nalaz:** merenje je pokazalo da sposobnost
    **već postoji** — potvrda manjka daje `SpoljniSvet → Kupac` (`ULAZ_TUDJE`,
    obaveza +N) i `Kupac → Vozac` (`POVRAT_PRAZNE`), saldo kupca 0, vozač +N,
    obaveza po tipu ambalaže tačna. Da sam odmah dodao vrstu u zatvoren enum,
    dodao bih je **pored** mehanizma koji već radi.
    **Rupa je u značenju:** planirana pozajmica i neobjašnjeno odstupanje
    ostavljaju **isti trag**. Dve stvari koje se ne razlikuju su tačno ono što
    ovaj refaktor uklanja.
    **Asimetrija otkrivena usput:** model ume da **vrati** tuđu ambalažu kao
    događaj (`VRACANJE_TUDJE` **jeste** zahtev), a da je **primi** samo kao
    posledicu (`ULAZ_TUDJE` **nije** zahtev, par uvek `SpoljniSvet → nalog`, i ne
    ulazi u par dokumenta).
    **Redosled je operaterov:** prvo `10c` (merge blocker), pozajmica posle
    merge-a. Zapisano na **tri mesta** da ne ispari: kanon (6.12k), tabela „Dug sa
    imenom", i memorija sesije — na operaterov izričit zahtev („ekstremno bitno
    za dalji rad, da se ne zaboravi").
87. **`10c` prvi rez: lista i storno reversa na ambalaznom dokumentu**
    (08.10.2026, 6.12l) — ovo je **`P2 #1`**, jedini preostali merge blocker.
    Ekran Storno i donja lista na F7 dele **jednu** mapu tipa, pa su obe gledale
    u `tblAmbalaza`. Zato je operater video „Prikazano 0" iako je revers upisan
    — isti kvar, dva ekrana, jedan presek.
    **Obim je pao posle merenja:** ekran ne čita stare kolone direktno nego kroz
    mapu tipa, pa su tabela, identitet, broj i datum **četiri reda** u
    `modScrDokumenti`. Preflight storna je postao **jedna linija**
    (`AktivanPoIdentitetu`), ista koju prijemnica već koristi — umesto tri
    razrešavanja (`ActiveAmbalazaDokExists` + `ReversIDRazresi` +
    `ReversStanicaDan`).
    **„Smer" je nestao iz izbora**, i to je suština: dokument ga nema — ima
    vrstu. Smer je bio deo **plutaćeg identiteta** koji je `ODL-16/-17` ukinuo.
    **Sabotaža `storno-revers-smer` je zato OBRISANA, ne preusmerena**, a tvrdnja
    koju je držala **preseljena** na pravilo koje ju je zamenilo. Kapija nad
    obrisanim pravilom zacementira staro stanje — isti obrazac kao „brisanje
    pisca veze: proveri kapije".
    **Odluka koju sam u toku rada preokrenuo, i razlog:** prvo sam planirao
    **dedikovan builder** (kao izvod). Merenje je pokazalo da izvod ga ima jer mu
    je red **grupa redova**; revers je sada **jedan red zaglavlja = jedan red
    mreže**, pa generički builder odgovara — dedikovan bi prepisao filtere,
    pretragu, čipove i status. Četiri ćelije se čitaju iz
    `modAmbalaza.AmbDokRedMapa` (jedan prolaz; po redu bi bio sken po redu).
    **Kolona PARTNER je izbor prikaza, ne podatak:** red imenuje obe strane, pa
    se bira ona koja nije naša — a „ko je naš" čita se iz **ugovora**
    (`AmbNalogUKlasi`), ne iz spiska imena u ekranu. Vozač je dobio mapu imena:
    od `10b-2` je **nalog**, pa se pojavljuje kao protivpartner.
    **Dve mine koje `vba_check` ne vidi, obe uhvaćene pre suite-a:**
    `Optional mVoz As Object = Nothing` (podrazumevana vrednost objekta nije
    validan izraz u VBA) i **heredoc koji je opet pojeo `\n`** u katalogu
    sabotaža — isto pravilo koje memorija izričito zabranjuje; popravljeno
    skriptom iz fajla.
    Katalog 702 → 704 (jedna obrisana, tri nove). Jeftine kapije `rc=0`.
    **Stari klaster u `modStorno` nije obrisan** — još ga zovu testovi,
    `modStornoFlow` i tok ispravke reversa. Briše se u `10e`, po 6.13.
88. **`10c` prvi rez: pet krugova, i četiri nalaza u MOM kodu** (08.10.2026).
    Rez iz stavke 87 je prošao jeftine kapije iz prve, a onda ga je suite četiri
    puta vratio. Nijedan nalaz nije bio „zastareo test".

    | Krug | Nalaz | Gde je bio |
    |---|---|---|
    | 1 | tri tvrdnje mere stari oblik | dve **preseljene**, jedna **obrisana** sa pravilom |
    | 2 | preduslov testa: stanica izdaje gajbe koje ne drži | **moj test** — `ODL-8` ga je odbio po imenu |
    | 3 | `Dim dan As Long`, pa `dan` kao **tekst** | **moj kod** — „8.10.2026." `IsDate` odbija, opis tiho bez stanice |
    | 4 | izgubljena provera para `(identitet, broj)` | **moj kod** — zastareo izbor bi stornirao tuđ dokument |
    | 5 | `AktivanPoIdentitetu` traži `GeneracijaID` | **moj kod** — preflight bi odbijao **svaki** revers |

    **Peti je bio najteži i najopasniji:** `IdoviGeneracije` traži kolonu
    `GeneracijaID`, koju `tblAmbalazaDokument` nema — pa bi storno reversa sa
    ekrana **ne prolazio nikad**, a merge blocker ostao zatvoren sa drugom
    porukom. Prva tvrdnja ga je **prikrivala**: prolazila bi i za validan
    dokument. Otkrila ga je tek tvrdnja o paru, koja do svoje grane nije ni
    stizala — zato uz svaku kapiju odbijanja sada stoji i **protiv-slučaj**
    („aktivan revers sa svojim brojem PROLAZI").
    **Dijagnostička tvrdnja je platila odmah.** Posle dva kruga nagađanja o
    uzroku, jedna tvrdnja koja NAMERNO pada ispisala je stvarnu vrednost
    (`dat=8.10.2026.`) i uzrok je bio vidljiv iz prvog pokušaja. To je jeftinije
    od svake teorije — jedan prolaz suite-a.
    **`dokaz.py` me je tri puta ispravio u MERENJU, ne u kodu:**
    `NE OBARA NISTA` (dvoslojna kapija — par i sam odbija nepostojeći dokument),
    `PALA DRUGA TVRDNJA` (sabotaža opisa je skidala i datum, jer `ReversOpis` sa
    praznom stanicom vraća prazno), i **vakuumska tvrdnja** (prazan broj meri
    generičku kapiju iznad `Select Case`, ne postojanje dokumenta). Dve brane
    razdvaja **jedino storniran dokument**: broj mu odgovara, a postojanje ne —
    pa je tvrdnja postala „već storniran revers se ne stornira ponovo", što je uz
    to i poslovno pravilo za sebe.
    **Završno:** `RunAllTests` **201**/0, `dokaz.py --grupe 6 amb-10c-` crvenih
    **5/5**, pun prolaz **ZELENO 12/12** (BFP 2436/0, Storno 164/0, banka 241/0),
    GREEN marker nad izvorom `74c07d9ce2c0`. Katalog 702 → 706.
    U `Test_StornoCentar_All` su obrisane **dve** provere `DokumentOpis` nad
    nogama — naslednice postoje nad pravim dokumentom. Ostatak tog testa
    (undo, ispravka, Nedovršeno) **namerno** i dalje meri stari model: te
    putanje nisu presečene i žive do `10e`.
89. **`P2 #1` zatvoren: dokaz nad KONACNIM izvorom** (08.10.2026).
    Prethodni dokaz je bio nad starijim potpisom — posle njega je još menjan
    `modTestStornoCentar`, pa je pušten ponovo: `crvenih 5 / sabotaža 5`, potpis
    izvora `ed649ee2f0354214` **identičan pre i posle**, `DOKAZANO`.
    Time je **jedina preostala stavka iz review-a `024995de` zatvorena**: ekran
    Storno vidi i **naš** i **kupčev** ambalazni dokument, bira ga po `AmbDokID` i
    stornira kroz `StornirajAmbDokument_TX`.
    **Šta `10c` još nosi** (nije merge blocker): osam produkcionih modula i dalje
    čita stare kolone — `modIzvestaj` (27), `modAmbalaza` (18), `modStorno` (14),
    `modIntegritet` (4), `modStornoFlow` (2), `modDokumenta` (2), `modStornoZurnal`
    (1). To su izveštaji, integritet i žurnal; oni čitaju **stari** model, koji još
    nosi stare redove, pa im je presek deo „staro i novo jedno protiv drugog" iz
    6.13 — ne ulazni ekran.
90. **Review `7299b03f`: dva P1 zatvorena — i jedan je bio NEVIDLJIV suite-u**
    (08.10.2026).
    **P1 #1, najvažniji nalaz celog reza:** `10c` je presekao **mapu tipa** na
    `tblAmbalazaDokument`, ali je **filter REDA ostao legacy**. `RevRowVisible`
    sudi po `DOK_TIP_OM_*` i entitetu noge, pa vrsta sa zaglavlja (`"REVERS"`)
    pada u `Case Else` → `False` → **svaki nov dokument ispada iz liste**. Ekran bi
    opet pisao „Prikazano 0" — tačno ono što je rez trebalo da popravi.
    **Zašto su jeftine kapije i 12/12 ZELENO bili saglasni sa praznom listom:**
    nijedan test nije zvao **čitaoca**. Svi su merili mapu (tabela, identitet,
    broj), `StornoRazlog`, `StornoIzvrsi` i `DokumentOpis` — sve **oko** njega.
    Zamena je `AmbDokUListiReversa`: odluka na nivou **dokumenta**, zatvoren spisak
    (`REVERS`, `REVERS_PARTNERA`). Namerno **ne** `AmbDokVrstaPoznata` — nabavka i
    otpis su poznate vrste, ali nisu reversi.
    **P1 #2:** broj je pratio **smer**, a mora da prati **vlasnika niza**. Moja
    prva verzija je čistila polje samo na prelasku `1..4 → 5`; obrnuto je zvala
    `RefreshBrojPredlog`, a on sa isključenim `AUTO_BROJ`-em vraća prazno i **izlazi
    bez diranja polja** — pa bi kupčev broj tiho postao broj **našeg** dokumenta.
    Sada: `1 ↔ 2` i `3 ↔ 4` ne diraju polje, a svaka promena **vlasnika** ga prazni.
    **P2:** `ISPRAVKA` se više ne nudi za revers — još ide kroz
    `RunReversCorrection` → `ReversIDRazresi`, koji očekuje `AmbID` noge i stari
    `DOK_TIP_OM_*`. Ponuđena radnja koja nad izabranim dokumentom **ne može** da
    radi gora je od radnje koje nema. Vraća se uz `10d/10e`.
    **P3:** `StornirajAmbalazuDokumenta` je **fail-closed** za ambalazni dokument
    (nula aktivnih originala → greška). Robni sme da nema red — prijemnica bez
    ambalaže je legitimna — pa za njega ostaje tih izlaz.
    **Nov test je klasa koja je falila:** `T_ReversiLista_CitaAmbalazniDokument`
    zove **`RedoviZaTip("REVERSI")`**, isti poziv koji radi mreža. Dve tvrdnje: nov
    dokument **jeste** u listi, a nabavka **nije** — druga meri baš zatvoren spisak.
    Završno: `RunAllTests` **202**/0, `dokaz.py --grupe 6 amb-10c-` crvenih
    **7/7** `DOKAZANO`, pun prolaz **ZELENO 12/12**, marker nad izvorom
    `bd976a311551`. Katalog 706 → 708.
91. **Review `cf14dc92`: broj je preživljavao promenu REZIMA, ne samo smera**
    (08.10.2026).
    Prethodna ispravka je rešila put **unutar** F7 (`SetSmerRev`), ali je polje
    broja **zajedničko** za sve režime (`fgBrOtpr`). `SelectModeCore` je računao
    da će `RefreshBrojPredlog` pregaziti stari broj — a on upisuje **samo kad ima
    šta da predloži**:

    ```
    AUTO_BROJ_DOKUMENTA = NE   -> SuggestNextBroj vraca "" -> Exit Sub
    rezim bez niza (F5, F6)    -> KindZaRezim vraca ""    -> Exit Sub
    ```

    **Nalaz je širi nego što je review opisao:** u režimima bez niza stari broj
    ostaje **i sa uključenim** auto-brojem. Komentar u kodu je tvrdio da je to
    pokriveno („inače bi u polju ostao broj iz prethodnog niza") — implementacija
    se oslanjala na pregazenje koje se ne dešava. **Komentar nije kapija.**
    Posledica nije kozmetička: `SkupiPolja` taj broj šalje piscu, a `OtkupValidiraj`
    ga ne odbija — ručno unet broj van naše šeme je legitiman
    (`BrojOdgovaraKontekstu` → `NEPRIMENLJIVO`). Kupčev broj reversa je tako mogao
    da postane **broj otkupnog lista**.
    Ispravka: `stariRezim` se čita **pre** dodele, polje se prazni **pre** predloga.
    `T_RezimBroja_PrelazakNeNasledjuje` je **životni ciklus**, ne helper: ide kroz
    `SelectMode`, meri sva tri slučaja (AUTO off, režim bez niza, AUTO on), i
    **skuplja nalaze pa tvrdi posle** vraćanja podešavanja i `Unload`-a — pad usred
    testa bi inače ostavio `AUTO_BROJ` isključen za sve naredne testove.
    **P3 zatvoren, uz nalaz u mom testu:** kupčev revers se nije pojavio u listi —
    ne zbog koda, nego zbog **UI keša**. Produkcija posle upisa zove
    `Scr_ResetCache`; test koji zove pisca **direktno** mora sam da ispuni isti
    preduslov, inače meri keš a ne listu.
    Završno: `RunAllTests` **203**/0, `dokaz amb-10c-` **7/7**, `dokaz amb-rezim-`
    **1/1**, oba `DOKAZANO`, pun prolaz **ZELENO 12/12**, marker nad izvorom
    `96c6757e2ded`. Katalog 708 → 709.

92. **Compile potvrdjen nad TACNIM izvorom; `P2` evidencije zatvoren** (08.10.2026).
    Operater je pustio `Debug -> Compile VBAProject` nad `a4e8729e`, pa je marker
    upisan: `python tools/vba_gate.py --mark-compile` -> izvor **`96c6757e2ded`**,
    isti otisak nad kojim stoji i zeleni marker. `--require-green --require-compile`
    izlazi **0**. Compile je vezan **samo za izvor**, pa ovaj zapis (docs) potvrdu ne
    obara.
    **Zasto uopste ide u commit:** marker zivi u `tests/last_green.json`, koji je
    **gitignored** -- review koji cita repo ga ne vidi. Nalaz "compile nije dokazan"
    je zato bio tacan o **evidenciji**, ne o izvoru, i resava se zapisom a ne
    recenicom u chatu.
    **Preostali `P3` je proveren u kodu i siri je nego u opisu.** `SelectModeCore`
    radi `MarkClean` -> `mPopMute = False` -> `mLoading = False`, pa **tek onda** pise
    u polje broja. `MarkDirty` prestaje da propusta programski upis cim `mLoading`
    padne, a `SetFld` menja `.text` -> `UiChange("fgBrOtprT")` -> `MarkDirty`, pa
    zaglavlje nove, prazne forme kaze "nesacuvano".
    **Nije regresija ovog reza:** na `cf14dc92` je `RefreshBrojPredlog` vec stajao
    posle `mLoading = False`, pa je svaki rezim sa auto-brojem badge prljao i pre P1
    ispravke; ispravka je u isti prozor dodala samo jos jedan upis (praznjenje), koji
    to cini i sa **iskljucenim** auto-brojem. Druga dva pozivaoca (prefill, upis
    dokumenta) zovu `MarkClean` **posle** broja -- obrazac postoji, `SelectModeCore`
    je jedino mesto koje ga ne postuje.
    **Ne ulazi u ovaj rez, i to je merenje a ne odlaganje:** svaka izmena `src-vba`
    obara i compile i zeleni marker, a promena ponasanja nosi test -- badge nije
    `Private` tvrdnja nego natpis u zaglavlju, pa nov test po `testovi.md` §6 trazi
    dvosmeran dokaz. Cena jednog premestenog reda je dakle pun prolaz + `dokaz` +
    **drugi** operaterov compile. Ide kao prvi mali rez posle merge-a.

93. **NALAZ: cutover je presekao PISCA, a čitaoci izveštaja su ostali na starom
    modelu** (08.10.2026).
    Tvrdio sam — i u prvo izdanje opisa PR-a #408 upisao — da se stari model „i dalje
    piše", pa merge ne ostavlja `main` polomljen. **Netačno.** Merenje:

    ```
    TrackAmbalaza (stari pisac), produkciona pozivna mesta
      na grani:   0        zivi pozivi su SAMO zasejavanje u modBusinessFlowProTests
      na main-u:  modOtkup 1445-1460, modDokumenta 4969, 6857, 6863, 7026, ...
    modIzvestaj -> nove kolone (Od_Tip / Na_Tip / VrstaKretanja):  0 referenci
    ```

    Stare kolone su, dakle, **prazne na svakom redu koji produkcija upiše**, a šest
    funkcija izveštaja čita **isključivo** njih: `ReportAmbalaza`,
    `ReportAmbalazaZbirnoSvi`, `ReportKarticaAmbalaze`, `ReportKarticaKooperanta`,
    `ReportSaldoOM`, `StampajReversAmbalaze` (+ `IzvStaniceIzPodataka`). Sve vise na
    **živom** ekranu `IZVESTAJI` (`modScrIzvestaji`), a `ReportSaldoOM` zove i sama
    ljuska (`modOtkupUI.RefreshSaldoOM` — saldo koji operater vidi **pri unosu**) i
    **PWA sync** (`modStammdatenSync`).
    **Zašto 12/12 ZELENO ovo ne protivreči:** `modIzvestajTests` (6 mesta) i
    `modBusinessFlowProTests` (11) **same seju stari oblik** kroz `TrackAmbalaza` — pa
    mere čitaoca nad redovima koje produkcija više ne pravi. Zeleno je tačno, ali o
    **premisi koju je test sam postavio**; ista klasa kao „test meri pomoćnik, ne
    poziv", samo na nivou modela podatka.
    Posledica za plan: **`10c` čitaoci nisu čišćenje posle cutover-a nego njegov deo.**
    PR #408 je zato prebačen u **draft**, a opis ispravljen na mestu.

    **Redosled rezova odavde (08.10.2026):**

    1. **`10c-2` `modIzvestaj`** — 24 mesta, 6 funkcija, operaterski vidljivo + PWA.
       Test zasejavanje **ide sa njim** (`modIzvestajTests` 6, BFP 11). Rizik koji se
       imenuje unaprijed: tvrdnja koja seje stari oblik pa tvrdi „nema redova" posle
       prelaska postaje **vakuum** — svaka takva mora da dobije pozitivnu
       protivtvrdnju, inače zeleno ne meri ništa.
    2. **`10c-3` storno i integritet** — `modStorno` 9 (`ActiveAmbalazaDokExists`,
       `ReversIDGranica`, `ReversIDStanice`, `ReversStanicaDan`), `modStornoFlow` 2,
       `modDokumenta` 2, `modIntegritet` 3 (`Chk_B10_ReversBezID`). **Prvo merenje, pa
       rez:** posle prvog reza `10c` deo njih može biti **mrtav** — tada je to
       brisanje, ne prelazak, i ne nosi istu cenu.
    3. **merge #408.**
    4. **`P3` badge** — svoj mali rez; može i uz `10c-2`, pa je jedan operaterov
       compile manje (cena je mešan rez).
    5. **`10d`** — istorijski upit, `UndoOperation_TX`, ispravka reversa na `AmbDokID`;
       time se vraća i privremeno **skinuta** akcija ISPRAVKA.
    6. **`10e`** — brisanje starog modela: kolone, `TrackAmbalaza`, `KnjigaIndeksi`
       (namerno dvomodelni čitalac) i konstante u `modConfig`.
    7. **`AMB-10-ODL-24`** — **odluka ide pred kod**: `VrstaKretanja` je zatvoren enum,
       pa se specifikacija pozajmice piše pre `10e` da se enum širi **jednom**, a
       implementacija ide posle `10e`, kad postoji samo jedan model. Dok traje `10c`,
       nov događaj bez starog ekvivalenta bio bi **nevidljiv** svakom nepresečenom
       čitaocu — to je i razlog zašto ne ide prvi, uz operaterov redosled.

94. **`10c-2` prvi rez: sloj API-ja, pa izveštaj** (08.10.2026).
    Pre koda je izmerena jedna stvar koja je promenila obim: **`modPrint` je bio
    slep bez ijedne reference na staru kolonu.** Zove stare čitaoce u
    `modAmbalaza` (`GetStanicaAmbSaldo` 343 i 1243, `GetKooperantAmbOpening` 741,
    `GetAmbalazeStanje` 2053), pa ga grep po kolonama ne vidi. Prava mera nije
    „imenuje staru kolonu" nego **„dohvata stari model"** — spisak iz stavke 93 je
    zato bio uži od stvarnosti.
    Zbog toga rez ide **slojem**, ne modulom: presečeni su `GetAmbalazeStanje` i
    `GetKooperantAmbOpening`, a `GetStanicaAmbSaldo` je tanak omotač nad prvim —
    pa je **`modPrint` izlečen bez ijedne izmene u `modPrint`**.
    **Znak je izmeren, ne pretpostavljen:** staro `Ulaz(+) / Izlaz(−)` po entitetu
    je **identično** novom `Na(+) / Od(−)` po nalogu, a tipovi naloga
    (`modAmbalazaUgovor.bas:33-38`) su **isti stringovi** kao stari `EntitetTip` —
    pa pozivaoci ostaju nedirnuti. Potvrđeno je i da novi redovi stari čitalac ne
    **ruše** nego ga tiho preskaču: filter je `EntitetID = ...`, a on je prazan, pa
    se `Case Else → Err.Raise "Neispravan smer"` nikad ne dosegne.
    **Lifecycle se menja u istom rezu:** `ExcludeStornirano` (mutabilna zastavica,
    koju nov model **ne piše**) → `RedDoticeKnjigu` (kontra-stav). Bez toga bi
    storniran revers i dalje stajao u saldu.
    **Nov primitiv `AmbSaldoPoNalogu(tip)`** — saldo **svih** naloga jednog tipa u
    **jednom** prolazu. Postoji zato što je `ReportSaldoOM` imao **svoj** prolaz,
    svoj `Select Case` po `Smer`-u i svoj `ExcludeStornirano`: tri stvari koje su
    vlasništvo knjige. Poziv u petlji bi bio N prolaza, pa oblik ide u jezgro, a ne
    kopija pravila u izveštaj.
    **Zašto `2a` i `2b` nisu razdvojivi:** `modTest.IzvAmbSaldo` i
    `modGoldenTests.GldAmbSaldo` čitaju kroz `GetAmbalazeStanje` i porede sa
    `ReportSaldoOM`. Dok je jedan sloj na novom a drugi na starom modelu, ta tvrdnja
    poredi **nov broj sa starim** — a do sada je poredila **0 sa 0**, dakle
    vakuumski. To je isti vakuum koji je stavka 93 imenovala unaprijed.
    **Test #204 `T_AmbSaldo_CitaociSuNaNovomModelu`** seje **isključivo** kroz
    produkcione pisce (`NabaviAmbalazu_TX`, `UpisiReversAmbalaze_TX`,
    `StornirajAmbDokument_TX`) i meri **deltu**, ne apsolutni saldo — fixture već
    nosi ambalažne redove, pa bi apsolutna tvrdnja merila fixture a ne rez. Meri
    četiri stvari: priliv, **obe strane istog reda** (stanica −4 i kooperant +4),
    poklapanje dva oblika istog pravila, i vraćanje na nulu posle storna.
    Katalog sabotaža 709 → **712** (jedna strana, kontra-stav, mapa naloga).
    `GetVozacAmbSaldo` je izmeren kao **mrtav** (0 pozivaoca u `src-vba`) — ne briše
    se ovde nego u `10e`, sa ostatkom starog modela.
    Stanje posle reza: `modIzvestaj` 24 → **21** mesto (ostaje pet funkcija:
    `ReportKarticaKooperanta`, `ReportKarticaAmbalaze`, `ReportAmbalazaZbirnoSvi`,
    `IzvStaniceIzPodataka`, `ReportAmbalaza`, `StampajReversAmbalaze`),
    `modAmbalaza` 17 → **11** (ostatak je namerno dvomodelan: `KnjigaIndeksi`,
    `LegacyRedProblem`, `RequireAmbalazaSchema`, i mrtav `GetVozacAmbSaldo`).
    Jeftine kapije: `vba_check` čisto (712 sabotaža, 0 nalaza), `who_writes`
    `--check` i `--check-ownership`, `gen_schema_module --check`,
    `popis_citalaca --check` — sve zeleno. **Skupe čekaju reviewer GO**, a compile i
    zeleni marker su **oboreni** izmenom izvora (`96c6757e2ded` → `93ac16cfe092`).

95. **`10c-2`, drugi deo: kartice — i jedna odluka pročitana, ne doneta**
    (08.10.2026).
    Kartice su **redni** izveštaji: treba im datum, dokument i doprinos **tog**
    reda, a ne saldo — pa im primitiv iz stavke 94 ne pomaže. Zato nov javni
    `AmbKretanjaNaloga(tip, id)`: 2D niz `Datum · DokID · DokTip · TipAmbalaze ·
    Kolicina **sa znakom**`, u jednom prolazu. Znak ostaje u knjizi (`+` kad nalog
    prima, `−` kad daje, 6.8); do sada ga je **svaka kartica vadila sama** iz
    `Smer`-a, i svaka je nosila svoj `ExcludeStornirano`.
    `ReportKarticaAmbalaze` i `ReportKarticaKooperanta` time gube i filter po
    entitetu (redovi su već naloga) i svoj `Select Case`. Kontra-stav se vraća kao
    **običan red sa svojim znakom** — kartica mora da **pokaže** storno, a ne da ga
    sakrije; saldo se time sam vraća na početno.
    **Odluka o vozaču je pročitana iz 6.8, nije doneta ovde:** „Nema grananja po
    tipu, nema inverzije, vozac ispada sam jer je **nalog**." Time staro isključenje
    `DokumentTip <> Otkup` u vozačevoj grani **otpada po odluci**, kao i
    `VozacAmbEffectiveSmer`. To je **promena ponašanja** vozačevog izveštaja i mora
    tako da se prijavi — nije prevod. Isti pasus daje i `AMB-INV-02`
    (`OdNalog <> NaNalog`), pa je komentar u `AmbSaldoPoNalogu` ispravljen: isti
    **nalog** na obe strane ne postoji, isti **tip** (dve stanice) postoji.
    **`ReportAmbalaza` i `ReportAmbalazaZbirnoSvi` idu ZAJEDNO** i zato nisu u ovom
    rezu: druga samo skuplja spisak naloga pa **delegira** prvoj. Da je presečena
    samo jedna, zbirni izveštaj bi nabrajao naloge iz novog modela a redove računao
    iz starog — dakle nule. Ista spojenost kao `2a`/`2b` u stavci 94.
    Tvrdnja je dodata u **#204** (jedan test, jedno zasejavanje): kartica mora da
    pokaže primljene gajbe **tog** reversa. Katalog 712 → **713**.
    Stanje: `modIzvestaj` 21 → **15** mesta, ostaju četiri funkcije
    (`ReportAmbalaza`, `ReportAmbalazaZbirnoSvi`, `IzvStaniceIzPodataka`,
    `StampajReversAmbalaze`). Jeftine kapije zelene; skupe čekaju reviewer GO.

96. **Review `034a63e3`: dva `P2`, oba moja — jedan ih je i cementirao**
    (08.10.2026).
    **`P2 #1`: kartica je prikazivala tehnički `AmbDokID` kao poslovni broj.**
    Knjiga nosi `(DokumentTip, DokumentID)` i to je **ispravno** — ledger ne sme da
    zna kako se dokument prikazuje. Ali za ambalažni dokument je `DokumentID`
    opaque `ADK-<hex>`, a poslovni broj i vrsta stoje na **zaglavlju**. Kartica je
    zato operateru pokazivala `ADK-8f...` umesto `1/081026`, i generički
    `AmbalazaDokument` umesto `Revers` — jedan koren, dva simptoma.
    **Gore od samog defekta:** moja tvrdnja u `#204` ga je **cementirala** —
    merila je baš `AmbDokID`, a komentar uz nju je tehnički detalj izgovorio kao
    poslovno pravilo („pa mu kartica kao broj prikazuje sam DokumentID"). Test koji
    opisuje zatečeno ponašanje umesto pravila ne čuva ništa; on ga **brani**.
    Ispravka je tamo gde je reviewer pokazao: `AmbKretanjaNaloga` se **ne dira**,
    nego dolazi prevod. Nov `modAmbalaza.AmbDokPrikazMapa()` daje
    `AmbDokID -> "broj|vrsta"` u **jednom** prolazu kroz zaglavlja, a
    `modIzvestaj.KarticaDokPrikaz` bira izvor broja po tipu dokumenta (otkup sa
    svog zaglavlja, ambalažni sa svog). Obe kartice idu kroz njega.
    Uz to je `AmbVrstaDokNaziv` **iseljen** iz `modScrDokumenti` u
    `modAmbalazaUgovor`: spisak vrsta je domenski (`AMB_DOK_*` je tamo), a trebao je
    i izveštaju — zavisnost bi inače išla **report → ekran**, što je obrnuto. Tri
    spiska naziva za isti enum bi se razišla prvom izmenom.
    **`P2 #2`: sabotaža `amb-10c-saldo-ne-vidi-kontrastav` je bila placebo.**
    Menjala je `RedDoticeKnjigu` u `If True` i tvrdila da time kontra-stav prestaje
    da se preskače — **a `RedDoticeKnjigu` kontra-stav ne preskače**: on samo bira
    **kanonske** redove, i original i kontra-stav su to. Storno se poništava
    **algebarski** (`+N` i `−N` daju 0). Za sve kanonske redove su `If
    RedDoticeKnjigu` i `If True` identični, pa sabotaža ne bi obarala ništa.
    Zamenjena je sa `amb-10c-saldo-ne-ponistava-storno`, koja **stvarno** izbacuje
    kontra-stav iz zbira (`Len(StornoOd) = 0`). Netačan komentar u
    `GetAmbalazeStanje` je ispravljen — tvrdio je isto što i placebo tvrdnja.
    `vba_check` ovo **nije mogao** da uhvati: on potvrđuje **oblik** sidra (pogađa
    tačno jedno mesto, tvrdnja je ceo statičan literal), ne njegovu semantiku.
    Zato placebo sabotažu hvata reviewer ili `dokaz.py`, a ne jeftina kapija.
    **Tvrdnje:** `#204` sada meri **poslovni** broj, uz **kontra-tvrdnju** da se
    tehnički `AmbDokID` ne vidi i tvrdnju da je imenovana **vrsta** sa zaglavlja.
    Dodata je i `ReportKarticaKooperanta` — njena samostalna ambalažna putanja je
    presečena istim rezom, a bila je nedokazana.
    Usput je kapija kataloga uhvatila **zastarelo sidro**: preimenovana tvrdnja je
    ostavila `amb-10c-kartica-ne-vidi-prijem` bez mete, i ona je preusmerena na
    tvrdnju koju stvarno obara (kartica kooperanta), da ne bi **delila** tvrdnju sa
    drugom sabotažom. Katalog 713 → **714**.
    **Za skupi prolaz, unaprijed:** golden fajlovi nose `kooperant 12/1  0`. Ta
    nula je bila **zagarantovana** dok je čitalac bio slep, pa danas ne nosi
    informaciju. Posle reza postaje stvarno merenje: ako ostane 0, treba potvrditi
    da je to **neto-nula** (otkup knjiži `UZ_ROBU` i `IZDATA_PRAZNA` u suprotnim
    smerovima), a ne i dalje slepo. Prva crvena u `RunGoldenSuite` zato nije
    automatski regresija.
    **Reviewer evidencija koju priznajem:** `modIzvestajTests` još seje stari oblik
    (`T_E2E_AmbPregledRazdvajaTipDokumenta`, `T_E2E_ReversIstiBrojDveStanice`), ali
    oba mere **`ReportAmbalaza`**, koja još nije presečena — legacy seme uz legacy
    čitaoca je konzistentno. Kad ta funkcija pređe, **moraju** i oni, inače se
    vraća isti vakuum koji je stavka 93 imenovala.

97. **`10c-2` peti korak: `ReportAmbalaza` presečena — i dva testa su time
    CRVENA** (08.10.2026).
    `ReportAmbalaza` + `ReportAmbalazaZbirnoSvi` sada čitaju knjigu kroz
    `AmbKretanjaNaloga`. Iz funkcije su **otišle četiri stvari**, i ni jedna zbog
    estetike:

    | Otišlo | Zašto |
    |---|---|
    | filter po entitetu kroz `clsFilterParam` | nov red imenuje **obe** strane, pa je uslov „nalog je na bilo kojoj" **ILI** preko dve kolone — `FilterArray` to ne ume |
    | `VozacAmbEffectiveSmer` (inverzija) | **po odluci** 6.8: „nema inverzije, vozac ispada sam jer je NALOG" |
    | `DokumentTip <> Otkup` za vozača | **po odluci** 6.8: „nema grananja po tipu" — gajba više nije žigosana na dva dokumenta, pa duplog terećenja nema |
    | `ReversID` u ključu grupisanja | `AmbDokID` **jeste** identitet dokumenta (ODL-16), pa je `(DokTip, DokID, TipAmbalaze)` već jednoznačno |

    **`COL_STORNIRANO` je zamenjen, ne obrisan:** knjiga je nepromenljiva i tu
    kolonu ne piše. Primitiv zato dobija kolonu **`Otkazano`** (kontra-stav **i**
    original na koji pokazuje) — i to je **kolona, a ne filter**, jer se dva
    čitaoca iste knjige legitimno razlikuju (6.8): **pregled kretanja** storniran
    dokument **skriva**, a **kartica** ga **prikazuje**, pošto je storno i sam
    događaj koji operater mora da vidi. Ta asimetrija je sada **tvrdnja**, ne
    navika.
    **Promena ponašanja koju prijavljujem, a nije tražena:** kolona „Mesto" u
    pregledu kretanja sada nosi **protivpartnera**, a ne naslovni entitet. Star red
    je imenovao jednu stranu, pa je u izveštaju po **vozaču** to i bila druga
    strana — ali u izveštaju po OM i po Kupcu ista vrednost u svakom redu, dakle
    nula informacije. Za vozača je ponašanje **identično** kao pre. `AMB-INV-10`
    garantuje jedan neuređen par po dokumentu, pa dva reda istog dokumenta ne mogu
    dati različitog protivpartnera.
    Prevod broja iz stavke 96 je iskorišćen i ovde (`ResolveDokBrojMape` dobija
    `AmbDokPrikazMapa`), kako je reviewer i predvideo. `ResolveEntitetName` je
    dobio **`Vozac`** — bez toga bi protivpartner stajao kao go ID.

    **MERENJE KOJE JE ISPRAVILO MOJE SOPSTVENO MERENJE:** spisak `10c` u stavkama
    93–96 je bio **uži od stvarnog**, jer `COL_AMB_REVERS_ID` nije bio u mom grep
    obrascu. Tačno stanje posle ovog reza:

    ```
    modIzvestaj   5   (IzvStaniceIzPodataka je presecen; ostaju ReversStampaNoge
                       i StampajReversAmbalaze -- STAMPA, ne pregled)
    modAmbalaza  12   (namerno dvomodelno: KnjigaIndeksi, LegacyRedProblem,
                       RequireAmbalazaSchema, mrtav GetVozacAmbSaldo)
    modStorno    14   modIntegritet 4   modStornoFlow 2   modDokumenta 2
    modStornoZurnal 1
    ```

    **DVA TESTA SU SADA CRVENA I TO SE NE PREĆUTKUJE.** Oba zovu `ReportAmbalaza`
    nad **legacy zasejanim** redovima, koje nov čitalac ne vidi:

    - `T_E2E_AmbPregledRazdvajaTipDokumenta` — premisa mu je **nestala**, ne samo
      zasejavanje: tražio je jedan `otkupID` pod **dva** tipa dokumenta
      (`Otkup` + `OM-Izlaz-Koop`), a danas otkup knjiži **oba reda pod `Otkup`**
      (`AMBALAZA.md` tabela, red 1–8). Pravilo koje je čuvao (`DokumentTip` ostaje u
      ključu grupisanja) i dalje postoji u kodu, ali ovaj proizvođač premise ne.
      Nova forma tvrdnje: jedan otkup → **jedan** red sa **oba** stupca, jer su
      `UZ_ROBU` i `IZDATA_PRAZNA` dve vrste istog dokumenta.
    - `T_E2E_ReversIstiBrojDveStanice` — identitet mu je `ReversID` (REV-IDENT-01),
      koga je zamenio `AmbDokID` (ODL-16). Deo koji meri **štampu**
      (`ReversStampaNoge`) je i dalje konzistentan sa starim modelom, jer štampa
      **nije** presečena — pa test straddle-uje dva modela i mora da se podeli kad
      štampa pređe.

    **Zato skupi prolaz NE ide pre tog reza** — bio bi zagarantovano crven, i to
    na mestu koje već znam. To je i reviewer-ova otvorena stavka „report fixture
    cutover još nije završen", sada sa imenima i razlozima.
    Nove tvrdnje u **#204**: zbirni pregled vidi izdate gajbe (po **delti**),
    pregled kretanja pokazuje revers pod **poslovnim** brojem i **ne** pokazuje
    `AmbDokID`, a posle storna pregled ga **skriva** dok ga kartica **i dalje
    prikazuje**. Katalog 714 → **716**.

98. **`10c-2` zatvoren: štampa ambalažnog dokumenta i dva testa** (08.10.2026).
    Nov dokument ima **svoje zaglavlje** i **jedan** poslovni red koji imenuje obe
    strane, pa rekonstrukcija „nogu" (dve strane, `ReversID` kao identitet) nije
    ni potrebna ni moguća. `StampajReversAmbalaze` zato dobija granu
    `StampajAmbDokument`, a stari put **ostaje** za zatečene redove i odlazi sa
    njima u `10e`.
    **Nalaz koji je ovaj rez sprečio:** na **papir za potpis** bi išao `ADK-<hex>`
    umesto poslovnog broja. Legacy samostalan revers je imao
    `DokumentID = brojDok`, pa je stari kod broj dobijao **besplatno**; nov
    `DokumentID` je `AmbDokID`, pa se broj mora razrešiti. Ista klasa greške kao
    `P2 #1` sa kartica — samo na dokumentu koji operater potpisuje. `AmbDokPrikazMapa`
    je zato proširena datumom (papir nosi datum **zaglavlja**, ne reda).
    **Fail-closed na dva mesta, oba namerno:** dokument bez aktivnog poslovnog reda
    (storniran ili nepostojeći) **nema papira**, i izabrani tip gajbe koji nije tip
    poslovnog reda **odbija** štampu — papir na pogrešan tip je greška koju operater
    ne vidi.
    **Dva testa: jedan je dobio novu formu, drugi je izgubio polovinu.**
    `T_E2E_AmbPregledRazdvajaTipDokumenta` → **`T_E2E_AmbPregledKanonskiDokumenti`**.
    Stara premisa (jedan `otkupID` pod **dva** tipa dokumenta) je nestala, ali
    pravilo koje je čuvala nije: pregled ne sme da **spoji dva dokumenta** u jedan
    red, jer ref-ključ tada vodi štampu na pogrešan papir. Usput je dodata tvrdnja
    koju stari **nije imao**, a nov model je traži: **dve vrste kretanja istog
    dokumenta** (`POVRAT_PRAZNE` + `IZDATA_PRAZNA`) daju **jedan** red sa oba
    stupca. Tvrdi se i da kolona dokumenta nosi **poslovni** broj, i da „Mesto"
    nosi **protivpartnera**.
    `T_E2E_ReversIstiBrojDveStanice` je **izgubio deo o pregledu** — seje stari
    oblik, pa bi tvrdnja merila prazan rezultat, ne pravilo. **Deo o štampi ostaje**
    i dalje meri stari put, jer on još postoji; oba odlaze u `10e`.

    **MERENJE KOJE JE PREMESTILO TVRDNJU:** kapija kataloga sabotaža priznaje samo
    `modTest`, `modTestBanka` i `modBusinessFlowProTests` — `dokaz.py` tvrdnju iz
    `modIzvestajTests` **ne može da obori**. Zato je tvrdnja o identitetu dokumenta
    („dva reversa ostaju dva reda") preseljena u **#204**, gde se seje kroz
    produkcione pisce i gde sabotaža stiže. Tvrdnja o **spajanju dve vrste** ostaje
    bez sabotaže i to je **imenovana rupa**, ne propust.
    Katalog 716 → **717** (jedna sabotaža je skinuta jer joj tvrdnja nije dosegljiva).
    Stanje: `modIzvestaj` **5** mesta — sva u **legacy putu štampe**
    (`ReversStampaNoge`, `StampajReversAmbalaze`), koji po planu odlazi u `10e`.
    **`10c-2` je time zatvoren.** Sledi `10c-3`: `modStorno` 14, `modIntegritet` 4,
    `modStornoFlow` 2, `modDokumenta` 2, `modStornoZurnal` 1.

99. **Review `2f2f3eaf`: `P2` delimitera — pretpostavka napisana kao činjenica**
    (08.10.2026).
    `AmbDokPrikazMapa` je pakovala `"broj|vrsta|datum"`, a ja sam uz nju napisao
    komentar: *„Broj je poslovni i ne sadrži `|`"*. To nije bila provera nego
    **pretpostavka u obliku tvrdnje** — i pogrešna: kupčev broj reversa (ODL-23)
    dolazi sa **kupčevog** dokumenta, a izmereno je da **nijedna kapija ne filtrira
    karaktere** (nema `InStr`/`Replace` nad `"|"` ni u `modNovacUnos` ni u
    `modBrojevi`). Dakle `KUP|R-17` je legalan broj, a parser bi pročitao
    `broj="KUP"`, `vrsta="R-17"`, `datum=vrsta` — i to **ne bi puklo** nego tiho
    promenilo i karticu i **papir za potpis**.
    **Ista rupa je bila i sloj niže, pa je i ona zatvorena:** `AmbDokRedMapa` je
    pakovala `"vrstaKretanja|tipAmb|kolicina|..."`, a `TipAmbalaze` je **operaterski**
    podatak iz lookup tabele (`TBL_TIP_AMBALAZE`) — isti rod izloženosti. Reviewer
    je prijavio samo prvu; popravka jedne a ostavljanje druge je tačno obrazac koji
    `CLAUDE.md` §2 zabranjuje.
    Oba protokola sada nose **niz**, ne spojen string: `Array(broj, vrsta, datum)`
    i `Array(vrsta, tipAmb, kolicina, odTip, odID, naTip, naID)`. Nizovi u
    `Scripting.Dictionary` rade i već se koriste u ovom repou — ograničenje iz
    memorije važi za **objekte** (`Set d(k) = obj`), ne za nizove. Datum se usput
    čuva kao **sirova** vrednost ćelije: tekstualni datum je već jednom pojeo
    stanicu iz opisa, jer je `IsDate("8.10.2026.")` **False**.
    Čitaoci presečeni na niz: `KarticaDokPrikaz`, `ResolveDokBrojMape`,
    `StampajAmbDokument`, `modScrDokumenti` (lista reversa). Posle reza **nijedan**
    `Split` nad tim mapama ne postoji.
    **Tvrdnja koja bi ovo uhvatila** je dodata u `#204`, i ide nad **mapom**, ne nad
    karticom: kupčev revers je par `Kupac ↔ Vozac`, pa ga na kartici kooperanta nema
    — a mapa je mesto gde je greška i živela. Upisuje se revers sa brojem
    `KUP|R-17` i tvrdi se da broj **preživljava** mapu i da vrsta **nije pomerena**.
    Sabotaža `amb-10c-broj-se-reze-na-delimiteru` reže broj na `"|"` — isto što je
    stari protokol i činio. Katalog 717 → **718**.
    Usput je kapija kataloga uhvatila **moju** grešku u samoj sabotaži: komentar
    posle nastavka reda `_` je sintaksna greška. To je ona ista zamka koja je dvaput
    naplatila 585 s i ubijen Excel — sada je kapija, pa je naplatila **nula**.

100. **Ispravka okvira: kupčev revers se NE štampa, i to nije rupa**
     (08.10.2026, operaterov nalaz).
     Upisao sam u stavku 98 i u tabelu duga da „kupčev revers nema papir", kao
     rupu u sposobnosti koju šablon ne pokriva. Operater je pitao jednu stvar koja
     je to obesmislila: **kako firma izdaje kupcu revers, i šta će kupcu naše
     prazne gajbe?**
     Merenje potvrđuje da je pitanje bilo pogrešno postavljeno:
     `UpisiReversPartnera_TX` knjiži **samo jedan smer** — `Kupac → Vozac` sa
     `POVRAT_PRAZNE`. Kupac dobija naše gajbe **pune**, uz robu
     (`AMBALAZA_UZ_ROBU`, `Vozac → Kupac`), i vraća ih prazne. Smera „izdavanje
     praznih kupcu" **kao poslovnog događaja nema** — pa nema ni papira koji bismo
     mi izdali.
     Pisač je to i **sam govorio**, u komentaru koji sam napisao a nisam pročitao
     kao odgovor: *„papir koji operater drži u ruci nosi kupčev broj"*. Dokument je
     **kupčev**; zato je broj obavezan i ne predlaže se (ODL-23). Štampati ga
     značilo bi izdati **drugi** dokument za isti čin.
     Ispravka u kodu je mala: odbijanje štampe za `REVERS_PARTNERA` sada imenuje
     **tačan** razlog (dokument je kupčev), a generičko „nema otkupnog mesta"
     ostaje za ostale parove bez stanice, koje šablon stvarno ne pokriva.
     **Red „kupčev revers nema papir" je izbrisan iz tabele duga** — nije dug.
     Za reviewer-ov uslovni `P1` („svaki F7 revers mora biti štampiv") odgovor je
     time **negativan po domenu**, ne po obimu reza.

101. **Klasifikacija `10c-3`: klasa A je PRAZNA — `#408` ga ne mora nositi**
     (08.10.2026).
     Reviewer je tražio da se 23 mesta **klasifikuju**, ne mehanički prepišu.
     Mesta žive u **deset** funkcija, sve `ReversID`-centrične, a `ReversID` je
     stara identitetska kolona koju nov pisac **ne piše** (zamenio ju je `AmbDokID`,
     ODL-16). Doseg je izmeren po pozivaocima, uz izbacivanje komentara:

     | Funkcija | Mesta | Klasa | Dokaz dosega |
     |---|---|---|---|
     | `modStorno.ActiveAmbalazaDokExists` | 2 | **B** | jedini cross-module pozivalac je `modStornoRecovery.UndoGuardReason` → **undo**, a to je `10d` |
     | `modStorno.ReversIDRazresi` | 1 | **B** | nijedan živ cross-module poziv — tri „poziva" u `modScrDokumenti`, `modScrStorno` i `modStornoDok` su **komentari**; put je `RunReversCorrection` = ISPRAVKA |
     | `modStorno.ReversRedoviRID` | 1 | **B** | `modStornoFlow` 322 / 536 / 2532 — ispravka reversa |
     | `modStorno.ReversIDGranica` | 4 | **C** | cross-module samo `modIzvestaj.ReversStampaNoge`, a njega kanonski dokument **više ne dosegne** (grana iz stavke 98) |
     | `modStorno.ReversStanicaDan` | 3 | **B** | `modDokUnos.ZavrsiIspravkuPitanje` (ISPRAVKA) + recovery |
     | `modStorno.ReversIDStanice` | 3 | **B** | `modStornoFlow:632` |
     | `modIntegritet.Chk_B10_ReversBezID` | 4 | **C** | gejtovan `modStorno.ReversTipJe(dokTip)` = **stari** `OM-*` tipovi; kanonski red pada u `ElseIf` sa praznim `ReversID`-om, pa **ne prijavljuje nalaz** |
     | `modStornoFlow.ScanRevers` | 2 | **B** | ispravka reversa |
     | `modDokumenta.GetStorniraniRevers` | 2 | **D** | `Private`, **nula** pozivaoca — mrtva |
     | `modStornoZurnal.ReversIDOperacije` | 1 | **B** | žurnal undo / ispravka |

     **Ključni dokaz za sve B stavke:** `modScrStorno.AkcijeRacun` za `REVERSI`
     vraća **samo** `STORNO` — ISPRAVKA je skinuta sa kanonskog UI-ja u `10c`, i
     vraća se u `10d`. Komentar u tom istom modulu to i kaže: *„ISPRAVKA i dalje ide
     kroz `RunReversCorrection` → `ReversIDRazresi`"*.
     **Ishod: klasa A ima NULA mesta.** Nijedno od 23 ne stoji na putu koji kanonski
     UI danas može da dosegne. `#408` zato **ne mora** da nosi `10c-3`: B ide sa
     `10d` (kad se ISPRAVKA i undo vrate, na `AmbDokID`), C i D sa `10e` (kad stari
     model nestane). Tako `#408` ostaje `10b-2 + 10c`, a ne postaje `10d`.

     **ALI klasifikacija je izbacila nalaz van tih 23 mesta, i on je klase A:**
     `modIntegritet` ima **nula** referenci na kanonske kolone
     (`Od_Tip` / `Na_Tip` / `VrstaKretanja` / `tblAmbalazaDokument`). Posle cutovera
     to znači da integritetni izveštaj proverava **samo legacy redove** — a
     produkcija ih ne piše. Operateru bi, dakle, rekao „nema nalaza" **ne pogledavši
     kanonsku knjigu**. To je fail-open izveštaj, ista klasa kao čitalac koji
     prećuti grešku.
     Nije **regresija** (pre cutovera kanonskih redova nije ni bilo) i ne kvari
     podatke, pa ga ne predlažem kao merge-blocker — ali mora da ima ime i red.
     **I treći put ista moja greška:** u `KarticaDokPrikaz` sam napisao da nogu bez
     zaglavlja „meri integritet". Ne meri. Komentar je ispravljen da kaže šta je
     izmereno, a ne šta bih voleo da važi.

102. **Prvi prolaz posle `10c-2`: alat je rekao „ne kompajlira se", a nije bilo
     tako** (08.10.2026).
     `run_vba.py --suite RunAllTests` je pao sa:

     ```
     SCHEMA FAIL  Cannot run the macro 'EnsureRuntimeSchema'
     SUITE  FAIL  Cannot run the macro 'RunAllTests'   (0.0s)
     ```

     a to je po `docs` i po memoriji **potpis modula koji se ne kompajlira** — jer
     takav modul obara **ceo** projekat, pa greška stiže kao „Cannot run the macro"
     na bilo kom makrou. Druga suite (`RunBankaImportTestSuite`, koju ovaj rez ne
     dira) je pala **isto**, što je potvrdilo „projekat, ne test".
     **Četiri hipoteze su odbačene MERENJEM, ne pretpostavkom:**

     | Hipoteza | Merenje |
     |---|---|
     | operaterov Excel drži instancu | `run_vba` koristi **`DispatchEx`** — svoju instancu, uz `AutomationSecurity = 1`; ne vezuje se na postojeću |
     | zaostala referenca na iseljen `AmbVrstaDokNaziv` | sve tri reference su kvalifikovane na `modAmbalazaUgovor`; nijedna nije ostala |
     | identifikator `do_` se završava podvlakom | kod **već** koristi `AS_`, `on_`, `from_`, `to_`, `IsStanicaActiveForOTK_` i kompajlira se → trailing underscore je legalan. **Da sam „popravio" ovo, bila bi četvrta pretpostavka napisana kao činjenica** |
     | uvoz je pao za neki modul | uvoz prijavljuje samo očekivanih 19 `.doccls` preskoka; 198 komponenti u projektu |

     **Pravo merenje:** zadržana temp kopija (`--keep`) je otvorena svojom
     instancom i `EnsureRuntimeSchema` je pozvan **i kvalifikovano i**
     **nekvalifikovano** — **oba prolaze**. Zatim je nad istom kopijom pozvana
     `RunAllTests`, **bez** runner-ovog koraka compile-a, i prošla je do kraja:

     ```
     TESTS=204 FAIL=5
     ```

     Dakle **projekat se kompajlira i makroi rade.** Ono što ih obara je runner-ov
     korak detekcije compile-a: on otvara VBE prozor i šalje tastaturne komande
     (`vbe.MainWindow.Visible = True`, `AppActivate`), verdikt vrati kao
     **`NEJASNO`** — i posle toga `Application.Run` više ne prolazi. Probe koji
     runner zove je **kvalifikovan** (`'sveska'!Func`) i zato prođe, dok se suite
     zove **nekvalifikovano**. Alat time „nejasan compile" pretvara u **lažno
     „projekat ne radi"**, i to na poruku koja čitaoca šalje da traži compile grešku
     koje nema.
     **Pet padova, po poreklu:**

     1. `T_AmbSaldo_CitaociSuNaNovomModelu` — **moja greška u testu.** Tvrdnja je
        merila deltu prema snimku s početka, a između sam (zbog tvrdnje 3e) upisao
        **drugi** revers (+2). Test je merio premisu koju je sam u međuvremenu
        promenio. Popravljeno: snimak ide **neposredno pred storno**, a tvrdnja je
        sada „storno gasi **tačno svoja četiri**" — meri storno sam po sebi.
     2. i 5. `T_Izv_SlaganjeKartica`, `T_Izv_TabKontekstRobaKupacSaldo` — kartica
        **−68**, kanonski saldo **4**. Oba su sada na kanonskom modelu, ali tvrdnja
        im je **jednakost** koja je važila samo dok su otkupne noge bile jednake:
        kartica kooperanta **izuzima** otkup-vezane redove, a `GetAmbalazeStanje`
        ih **uračunava**. U starom modelu su dve otkupne noge nosile istu količinu
        pa su se anulirale; kanonski `AMBALAZA_UZ_ROBU` i `IZDATA_PRAZNA` su
        **različite** količine. Premisa je nestala — traži odluku šta ambalažna
        kolona kartice **znači**, pa onda tvrdnju.
     3. `T_Izv_SlaganjeIsplataManjakAmb` — „ključ reversa ima nad čim da padne"
        vraća `False`: to je **preduslov** koji fixture više ne proizvodi.
     4. `T_Izv_DetaljICipKontekst` — pregled kupca ne pokazuje poslovni broj
        prijemnice `PRJ-FAK-3`. Jedini od četiri koji **može** biti nalaz u mom
        kodu, i zato ide prvi u sledećem krugu.

     **Stanje: 204 testa, 4 pada posle popravke prvog.** Skupe kapije se ne
     nastavljaju dok ta četiri ne budu zatvorena ili objašnjena — `dokaz.py` nad
     crvenom suitom ne meri ništa.

103. **Ambalažni fixture je kanonizovan u IZVORU — ali donor lanac je ustajao**
     (08.10.2026).
     Reviewer je sveo put do merge-a na dva posla; ovo je prvi.
     `tools/make_fixture.py` je sejao `tblAmbalaza` u **starom** obliku: po **dve
     noge** na događaj (`Smer` + `EntitetID` + `EntitetTip`), jer je star red
     imenovao samo jednu stranu. To je i bio izvor sva četiri preostala pada —
     **ne moj kod.**
     Blok je prepisan: svaki par nogu je postao **jedan** kanonski red, sa
     nedirnutim količinama, pa su i saldi isti:

     ```
     KOOP-TEST-1:  +30 +5 (izdate) -10 (povrat) +-99 (storniran par) -72 (uz-otkup) = -47
     ```

     To je **isti broj** koji je kartica davala i pre cutovera, pa slaganje kartice
     i kanonskog salda ponovo meri **jedan** skup, a ne dva modela.
     Usput su dve stvari prevedene po odluci, ne mehanički: storniran red je
     postao **par original + kontra-stav** (`StornoOd` → `AmbID` originala, strane
     zamenjene, **ista** vrsta kretanja — oblik koji
     `StornirajAmbalazuDokumenta` zaista piše), a stari `OM-Ulaz-Firma` red je
     postao **`NABAVKA`** (`SpoljniSvet → Stanica`), jer naše gajbe ulaze **samo**
     kroz nabavku (`AMB-10-ODL-8`). Dodata su i **zaglavlja** u
     `tblAmbalazaDokument` — bez njih bi kartica, pregled i papir pokazivali
     `ADK-...` umesto poslovnog broja.

     **ALI FIXTURE SE NE MOŽE REGENERISATI, i to je pravi blokator posla #1:**

     ```
     --donor tests/fixtures/otkup_test.xlsm
       SEMA: tblAmbalaza: donor nema kolone
             [OdNalogTip, OdNalogID, NaNalogTip, NaNalogID, VrstaKretanja]
     ```

     Kanonske kolone u svesku dodaje **`modSetup.EnsureRuntimeSchema`**, u runtime-u.
     Commit-ovani fixture ih **nema**, pa bi generator moje redove upisao **bez**
     tih kolona — a takav red ne pripada **nijednom** modelu i `KnjigaIntegritet`
     na njega **podiže grešku**. Taj izlaz je zato odbačen, ne instaliran.
     Donor koji kolone **ima** je temp kopija posle `EnsureRuntimeSchema`. Iz nje
     generisana sveska je **9,8 MB** prema **1,08 MB** sadašnje, i **bez `.sig`**
     fajla. Deset puta veći fixture nepoznatog porekla se ne instalira kao osnova
     projekta — to je tačno zamka koju vodimo kao „donor nosi formate" (jedan
     pogrešan donor je već oborio 10 testova bez ikakve veze sa kodom). Artefakt je
     obrisan.
     **Dakle posao #1 nije „prepiši fixture" nego „obnovi donor lanac":** treba
     donor koji nosi **tekuću** šemu a nije naduvan — najverovatnije operaterova
     radna sveska, koja kroz normalan rad već ima kanonske kolone. To je **njegov**
     fajl i njegova mašina, pa je odluka njegova.
     Izmena u `make_fixture.py` je **commit-ovana** i sintaksno proverena; čeka
     samo donora da bi se izvršila. Dok se fixture ne regeneriše, `RunAllTests`
     ostaje **204 / 4**, i to su ta četiri pada — ne nova.

104. **Fixture je regenerisan: `4 → 21 → 5`, i blokator je imenovan**
     (08.10.2026).
     Stavka 103 je tvrdila da je donor lanac ustajao. **Pola je bilo tačno, pola
     moja greška.** Ispravke po redu:

     1. **„Generator seje mrtve kolone otpremnice" — netačno.** Komentar u
        generatoru izričito kaže: *„Kolone zaglavlja (Klasa/Kolicina/KolAmbalaze/
        Cena) se **NE brišu** — odlaze u **S3e**"*, i iz njih se **izvode stavke**
        (`red["Kolicina"]`). Moja prva izmena ih je obrisala i oborila generator.
        **Vraćena.**
     2. **Pravi uzrok je drugde:** `modSetup.bas:1322-1326` te kolone **aktivno
        briše** (`ObrisiKolonuAko`, rez `S3`). Zato sveska koja je prošla
        `EnsureRuntimeSchema` **ne može biti donor** — i zato je jedini valjan
        donor stari fixture, koji pak nema kanonske ambalažne kolone.
     3. **Rešenje je bilo u samom generatoru**, idiomom koji on već ima:
        `ENSURE_COLS` dodaje kolone koje donor nema (tako već ulazi `ReversID`).
        Dodate su `OdNalogTip`, `OdNalogID`, `NaNalogTip`, `NaNalogID`,
        `VrstaKretanja`, `StornoOd` — i generacija je prošla, sa potpisom.
     4. **„Sveska od 9,8 MB je sumnjiva" — netačno**, i operater me je ispravio:
        prave sveske su 3–12 MB, a stari fixture od 1 MB je bio izuzetak. Ali
        pravi fixture iz **pravog** donora je **1,08 MB** — ista veličina kao
        postojeći, jer donor i jeste stari fixture.

     **Merenje posle regeneracije:**

     ```
     pre izmene fixture-a        204 / 4
     posle, sa mojom greskom     204 / 21
     posle ispravke greske       204 / 5
     ```

     Grešku je dalo **jedno** polje: sistemski nalog **nema ID**, a ja sam u
     `SpoljniSvet` upisao `"SpoljniSvet"` kao ID (`AmbNalogProblem` to odbija po
     imenu). Taj jedan znak je držao **14 od 21** pada — i to je dobar primer
     zašto se fixture meri, a ne procenjuje.

     **Preostalih pet, po poreklu:**

     | Pad | Poreklo |
     |---|---|
     | `T_AmbSaldo_CitaociSuNaNovomModelu` — storno dao `0` umesto `−4` | **moje područje**, traži merenje: ili `StornirajAmbDokument_TX` nije prošao, ili čitalac ne vidi kontra-stav |
     | `T_StornoBezUvida_NemaAkcije` — „revers je storniran" `False` | premisa je bila **`Stornirano="Da"`** red; kanonski storno je kontra-stav, pa zastavice nema |
     | `T_Izv_SlaganjeIsplataManjakAmb` — ručni prolaz `612` vs `0` | oracle još sabira ledger po **starom** pravilu znaka |
     | `T_Novac_BrojNijeJedinstven` — „broj reversa postoji u `tblAmbalaza`" `False` | star `DokumentID` je **bio broj** (`REV-IZV-1`); kanonski je `AmbDokID`, a broj živi na **zaglavlju** |
     | `T_BrojZauzetUNizu_Revers` — noga reversa `REV-IZV-2` ne postoji | isti uzrok kao iznad |

     **BLOKATOR, imenovan:** poslednja dva pada se **ne mogu** zatvoriti bez
     zaglavlja u `tblAmbalazaDokument` — tamo poslovni broj i živi. A generator
     **ne ume da napravi tabelu**: `ENSURE_COLS` dodaje kolonu, a nedostajuća
     tabela mu je `SchemaError` (`tblAmbalazaDokument ne postoji u donoru`; u
     aplikaciji je pravi `modSchema.EnsureAllTables` na startu). Dok to ne nauči,
     fixture nosi kanonske redove **bez** zaglavlja, pa im kartica, pregled i papir
     prikazuju `ADK-IZV-*` umesto broja. Ograničenje je zapisano u samom
     generatoru, na mestu gde je seme skinuto.
     **Nov fixture NIJE instaliran** kao osnova: stoji kao
     `tests/fixtures/otkup_test_kanon.xlsm` (potpis `e2d5d9487b30a3d7`). Zamena
     postojećeg je odluka operatera, ne moja.

105. **ISPRAVKA stavke 104: `make_fixture` UME da napravi tabelu** (09.10.2026).
     Stavka 104 je blokator imenovala kao „generator ne ume da napravi tabelu" i
     predložila nov `ENSURE_TABLES`. **Netačno** — `ENSURE_TABLES` **već postoji**:

     ```
     tools/make_fixture.py:2361   ENSURE_TABLES = { ... }        (na main: 2344)
     tools/make_fixture.py:2828   pravi Worksheet + ListObject   (na main: 2811)
     ```

     Komentar na mestu primene to i kaže: *„Nove TABELE koje donor nema (krug 5) —
     isto što radi `modSetup.EnsureDataTable`: sheet + ListObject sa kolonama"*. Kroz
     njega već ulaze `tblUtovar`, `tblUtovarStavke`, `tblPrevoznici`,
     `tblOtkupStavke`. Mehanizam se izvršava **pre** `DROP_COLS` / `RENAME_COLS` /
     `ENSURE_COLS`, dakle pre sejanja.
     **Pravi blokator je bio uži za jedan red:** `tblAmbalazaDokument` prosto **nije
     upisan** u `ENSURE_TABLES`. Oblik unosa je `"tblX": ("ImeLista", [kolone])`, a
     kanon već nosi i jedno i drugo (`sheet: "AmbalazaDokument"`, dvanaest kolona).
     **Kako je greška nastala:** izmerio sam da nedostajuća tabela diže
     `SchemaError` (`find_table` → `tblAmbalazaDokument ne postoji u donoru`) i iz
     toga **zaključio** da generator tabelu ne ume da napravi — a nisam pretražio
     fajl za postojeći mehanizam. Poruka o grešci je opisivala **simptom**, ne
     odsustvo sposobnosti. To je **četvrti** put u ovom nizu da sam merenje jednog
     mesta proširio u tvrdnju o celini: prethodna tri su `TrackAmbalaza` („stari
     model se i dalje piše"), `AmbDokPrikazMapa` („broj ne sadrži `|`") i
     `KarticaDokPrikaz` („to meri integritet").
     **Pouka koja ide u pravilo, ne u ovu stavku:** pre tvrdnje „alat to ne ume",
     pretraži alat za postojeći mehanizam. Poruka o grešci nije popis sposobnosti.
     **Red „`make_fixture` ne ume da napravi tabelu" je izbrisan iz tabele duga** —
     nije dug. Ostaje posao od jednog unosa u `ENSURE_TABLES` + vraćanje semena
     zaglavlja, i on **ne traži** zaseban rez nad alatom.
     Edge koji pri tome treba izmeriti, a ne pretpostaviti: postojeća grana pravi
     **nov** list (`wb.Worksheets.Add` → `ws_new.Name = sheet_name`), pa se treba
     uveriti šta se dešava kad **list** već postoji a `ListObject` ne — aplikacija
     taj slučaj rešava u `modSetup.EnsureDataTable`.

106. **`RunAllTests` 204 / 0 nad kanonskim fixture-om** (09.10.2026).
     Posao #1 iz reviewer-ovog plana je zatvoren za `RunAllTests`. Put je išao kroz
     **pet merenja**, i svako je ispravilo prethodnu pretpostavku:

     ```
     4   pocetno stanje (legacy fixture, kanonski citaoci)
     21  posle kanonizacije -- 14 od njih jedno polje: sistemski nalog NEMA ID
     5   posle ispravke tog polja
     2   posle tri zastarela oracle-a / premise
     1   posle druge otkupne noge i KolAmbIzdata
     0   posle jednosmernog OTK-NAL-DELIM
     ```

     **`ENSURE_TABLES` je bio jedan unos**, ne nov mehanizam (v. stavku 105):
     `tblAmbalazaDokument` + njegove četiri glave u SEED-u, i zaglavlja su počela da
     se seju.
     **`AMB-INV-07` je uhvatio nelegalan fixture.** Moja prva kanonizacija je otkup
     knjižila **jednom** nogom (`UZ_ROBU`), pa je kooperant pao na **−47** i svaki
     storno je padao sa *„storno bi ostavio saldo −48 na Kooperant"*. Invarijanta je
     izričita i potvrđena 28.09.2026: **nijedan realni nalog nema saldo < 0**, bez
     izuzetka. Otkup je **razmena** — `AMBALAZA.md` kaže da knjiži **dve** vrste
     (`UZ_ROBU` i `IZDATA_PRAZNA`) — pa sada i fixture tako radi. Stari model to nije
     prijavljivao jer ta invarijanta nad njegovim redovima nije ni postojala.
     **Kartica nije bila u krivu, sveska je bila nekonzistentna.** Razlika
     „kartica −43 / knjiga 29" je bila **točno 72** = ukupna otkupna ambalaža:
     `tblOtkup` nosi **dve** kolone (`KolAmbalaze`, `KolAmbIzdata`), čitalac čita
     **obe** (`modIzvestaj:896`), a fixture je sejao samo prvu. Pet otkupa koji u
     knjizi imaju `IZDATA_PRAZNA` nogu zato i na zaglavlju nose istu količinu.
     **Jedan otkup je NAMERNO ostao jednosmeran.** Sejanje `KolAmbIzdata` je oborilo
     `T_PrefillIzStorniranog_CitaSvojuTabelu`, čiji je komentar glasio *„nula se ne
     šalje: fixture nema izdatu ambalažu na otkupu"* — uzeo sam mu premisu.
     Računica dopušta da `OTK-NAL-DELIM` (2 gajbe) ostane jednosmeran: kooperant
     ostaje na **+23**, pa `AMB-INV-07` miruje, a pravilo „nula se ne šalje" **zadržava
     predmet merenja**. Tvrdnja nad `OTK-TEST-1` je prešla na ono što sada važi
     („izdata ambalaža se preuzima"), uz zapis gde je staro pravilo ostalo merljivo.

     **Šta OVO NE tvrdi:** izmereno je samo `RunAllTests`. BFP, Storno, Banka i
     golden **nisu** vrtjeni nad kanonskim fixture-om. Golden posebno nosi
     `kooperant 12/1  0`, a ta nula je bila zagarantovana dok je čitalac bio slep
     (stavka 96) — prva crvena tamo nije automatski regresija.
     **Nov fixture i dalje NIJE instaliran** kao osnova: `otkup_test_kanon.xlsm`,
     potpis `f5b6d4c488fdb068`. Zamena `otkup_test.xlsm` je odluka operatera.
     Posao #2 (`RunIntegritetProvere` slep na kanonski model) je **nedirnut**.

107. **`run_vba` zove makro KVALIFIKOVANO — lažni „kvar projekta" je zatvoren**
     (09.10.2026).
     Stavka 102 je izmerila da alat posle svog koraka detekcije compile-a prijavi
     *„Cannot run the macro 'EnsureRuntimeSchema'"* u `0.0s`, na **svakoj** suite —
     potpis koji `docs` i memorija vode kao **modul koji se ne kompajlira**. Ta
     poruka je dva puta u ovoj sesiji poslala dijagnozu u pogrešan smer.
     Popravka je **jedan red po pozivu**: `xl.Run(suite)` →
     `xl.Run(f"'{wb.Name}'!{suite}")`, isto i za `EnsureRuntimeSchema`. Probe
     (`_run_probe`) je **od početka** bio kvalifikovan i zato je prolazio — razlika
     je bila u **pozivu**, ne u projektu.
     Dokaz u oba smera je sama istorija merenja: nekvalifikovan poziv je davao
     `SCHEMA FAIL` u `0.0s`, a kvalifikovan nad **istom** temp kopijom daje

     ```
     SCHEMA  OK
     SUITE   OK     RunAllTests (73.5s)
     TESTS   RunAllTests: 204 ukupno, 0 palo
     REZULTAT: ZELENO
     ```

     **Red iz tabele duga je time zatvoren.** `COMPILE NEJASNO` i dalje stoji i
     dalje je poznato ograničenje — ali više ne obara poziv posle sebe, pa se
     „nejasan compile" ne pretvara u lažan „projekat ne radi".
     Izmena dira **alat**, pa je prijavljujem kao takvu: ako reviewer traži da ide
     zasebnim PR-om, lako se izdvaja — ali bez nje grana nije **merljiva**, jer
     suite sa dijalozima (BFP, Storno, Banka) traže runner-ov watchdog koji moja
     pomoćna skripta ne nosi.

108. **Pun prolaz 12/12 ZELENO nad kanonskim fixture-om** (09.10.2026).

     ```
     RunAllTests              204 / 0        RunStornoTestSuite       164 / 0
     RunBusinessFlowProSuite 2436 / 0        RunBankaImportTestSuite  241 / 0
     RunIzvestajTests          203 provere   + 7 ostalih suita        OK
     GREEN marker  izvor 214110fe23fb, ugovor 80213dfa2fd1,
                   sveska otkup_test_kanon.xlsm / 45ac0c56
     ```

     Put od prvog prolaza: **3 pala suite → 1 → 0**, i sva tri pada su bila u
     **testovima**, nijedan u produkciji.
     **`RunGoldenSuite` je prošao iz prvog puta.** To je zatvorilo otvorenu sumnju
     iz stavke 96: `kooperant 12/1  0` je **neto-nula**, a ne slepa nula. Otkup
     knjiži dve vrste koje se potiru, pa je taj red tačan i pre i posle cutovera.

     **Tri pada, tri različita uzroka:**

     1. **Kanonski čitalac validira nalog prema matičnoj tabeli** (`AMB-INV-03` —
        obe strane kroz **jednu** kapiju), a stari nije. `IZVT-OM` je živeo kao
        test-only stanica koje u `tblStanice` **nema** i prolazio godinama. To je
        **pooštravanje modela** koje je ovaj rez otkrio, ne regresija — test sada
        seje i stanicu i kooperanta.
     2. **Seme i orakl moraju biti u ISTOM modelu.** Četiri pada `RunStornoTestSuite`
        držao je **jedan** pomoćnik: `AmbSaldo` je zvao `GetAmbalazeStanje`, presečen
        u `10c` na kanonsku knjigu, dok taj modul seje legacy oblik i vozi **legacy**
        tok ispravke (`RunReversCorrection` sa `DOK_TIP_OM_*`, nedostupan sa
        kanonskog UI-ja, klasa B → `10d`). Orakl je vraćen u legacy model, kao i
        seme; oboje odlazi u `10e`.
     3. **`B10` je legacy provera i traži legacy podatak.** BFP premisa
        („`REV-IZV-1` ima 4 noge dva tipa pod jednim `ReversID`-om") izgubila je
        predmet kad je fixture kanonizovan. `Chk_B10_ReversBezID` je gejtovan na
        **stare** tipove dokumenta, pa su četiri legacy noge vraćene u fixture,
        izričito označene da odlaze u `10e`. Kanonski čitaoci ih preskaču
        (`RedDoticeKnjigu` je `False`), pa ne ulaze ni u jedan saldo.

     **Zadnji pad je bio o PRIKAZU, ne o modelu:** „kolona Mesto nosi
     protivpartnera" je padala jer `ResolveEntitetName` za **nepoznatog** kooperanta
     spaja dva prazna lookup-a i vraća `" "` — ne ID. Fixture je imenovao stanicu a
     ne partnera. Popravka je **pojačala** tvrdnju: meri se **ime** (`"IZVT Koop"`),
     pa tvrdnja sada stvarno razdvaja „protivpartner" od „naslovni entitet" — da
     kolona nosi naslovni entitet, tamo bi stajao naziv stanice.

     **ŠTA DOKAZ POKRIVA, A ŠTA NE:** marker je upisan nad
     **`otkup_test_kanon.xlsm`**, ne nad instaliranim fixture-om. Kapija to i kaže:

     ```
     --require-green --sveska tests/fixtures/otkup_test_kanon.xlsm   ->  dokazano
     --require-green  (podrazumevani fixture)                       ->  ODBIJA
         "sadrzaj sveske je drugi -- dokaz nije napravljen nad svescom koja se trazi"
     ```

     To je kapija koja radi kako je projektovana (kontekst sveske je **identitet**,
     ne ime). Da dokaz postane projektov, kanonski fixture mora da **zameni**
     `otkup_test.xlsm` — i to je odluka operatera, ne moja.
     Ostaje: posao #2 (`RunIntegritetProvere` slep na kanonski model), `dokaz.py`
     nad prefiksom `amb-10c-`, i ručni `Debug → Compile` nad konačnim izvorom.

109. **Dokaz je odbio dugačak test: `#204` rasturen na šest** (09.10.2026).
     Prvi `dokaz.py amb-10c-` je dao **NIJE DOKAZANO**, uz četiri problema:

     ```
     PALA DRUGA TVRDNJA   pregled-spaja-dokumente, saldo-ne-ponistava-storno,
                          kartica-ne-vidi-prijem
     NE OBARA NISTA       saldo-samo-jedna-strana
     ```

     **Izvor je bio netaknut** (`e3e87b52f0191b90` pre i posle) — sva četiri su
     bila u **sabotažama i u strukturi testa**, nijedan u produkcionom kodu.
     **Koren je izmeren u `dokaz.py`, ne pogođen:** ocena poredi samo ono što je
     palo **u njenom testu**, a komentar uz to kaže zašto — *„ciljana tvrdnja možda
     nije ni izvršena (`AssertEq` puca na prvom padu)"*. Dakle imenovana tvrdnja
     mora da bude **PRVA koja pukne** u svom testu.
     `#204` je u međuvremenu narastao u niz od devet pravila nad jednim
     zasejavanjem: `nabavka → revers → mapa → kartica → pregled → drugi dokument →
     storno → kartica po stornu → delimiter`. Sabotaža bilo kog čitaoca obarala je
     **prvu** tvrdnju, a ne onu koju imenuje. To je bila moja greška u **strukturi
     testa**, ne u pojedinačnim sabotažama.

     **Rasturen na šest, svaki sa svojim zasejavanjem i JEDNIM pravilom:**

     ```
     204  T_AmbSaldo_ObeStraneJednogReda            obe strane + mapa
     205  T_AmbKartica_PoslovniBrojIVrsta           broj, kontra-tvrdnja, vrsta
     206  T_AmbKarticaKooperanta_PokazujeRevers     svoja ambalazna putanja
     207  T_AmbPregled_DvaDokumentaDvaReda          identitet dokumenta + zbirni
     208  T_AmbStorno_GasiSvojeIKarticaGaPrikazuje  asimetrija dva citaoca
     209  T_AmbBroj_DelimiterPrezivljavaMapu        | u poslovnom broju
     ```

     Tvrdnje **pre** ciljane u svakom testu birane su tako da ih njena sabotaža
     **ne obara** — inače bi opet pucale prve. Zajedničko zasejavanje je izvučeno u
     `AmbSejRevers`, koji ide kroz **produkcione** pisce (nabavka pa revers, jer
     stanica mora da drži gajbe po `ODL-8`).
     Deveta sabotaža je dobila i **ispravno sidro**: prva verzija je zamenila samo
     **prvi** red nastavljene naredbe, pa je drugi (`entitetTip, entitetID) Then`)
     ostao siroče → sintaksna greška → projekat se ne kompajlira i sabotaža ne
     obori ništa. Sidro sada pokriva **ceo** iskaz (i uvlačenje nastavka je bilo
     pogrešno: 29 razmaka, ne 25).

     **Verdikt posle rasturanja:**

     ```
     crvenih 16 / 16        izvor pre/posle 14ccad6e88a3bc6b -> IDENTICAN
     grupno izmereno 14/16  === DOKAZANO (grupno) ===
     RunAllTests            209 / 0
     ```

     **Tri moje greške usput, i sve tri su uhvatile kapije:** dijakritika u VBA
     komentaru (`ASCII`), zamena komentara koja je presekla rečenicu i ostavila red
     bez `#` (`sabotaza.py` se nije učitavao), i **heredoc koji je opet pojeo
     backslash** — treći put u istoj sesiji, uprkos pravilu napisanom baš za to.

110. **Evidencija nad KONAČNIM izvorom, na jednom mestu** (09.10.2026).
     Review je tražio da se brojevi usklade i prijavio da pun prolaz prethodi
     rasturanju `#204`. **Prolaz je zapravo napravljen POSLE rasturanja** — ali to
     nigde nije bilo zapisano, pa je prigovor o evidenciji bio tačan i bez obzira
     na to. Otisak to i dokazuje:

     ```
     vba_gate --hash                  3b19a4defd38
     GREEN marker  izvor              3b19a4defd38      <- isti izvor
                   ugovor             5b73e8d90fc3
                   sveska             otkup_test.xlsm / 45ac0c56
     --require-green                  RC=0
     ```

     **Pun prolaz nad `3b19a4defd38`** (posle rasturanja `#204`):

     ```
     RunAllTests              209 / 0      RunStornoTestSuite       164 / 0
     RunBusinessFlowProSuite 2436 / 0      RunBankaImportTestSuite  241 / 0
     palih suita                0 / 12     REZULTAT                 ZELENO
     ```

     **Dokaz, bez dvosmislenosti:** `crvenih 16 / 16` znači da je **svih šesnaest**
     sabotaža oborilo **svoju** tvrdnju po imenu. `grupno izmereno 14 / 16` znači da
     je četrnaest od njih mereno u **grupi** (više mutacija u jednom prolazu suite),
     a dve pojedinačno. To nisu dva različita broja o istoj stvari nego **ishod** i
     **način merenja** — verdikt je zato `DOKAZANO (grupno)`.
     Po `CLAUDE.md` §5 grupni dokaz je **razvojni**: *„u rezu se pušta grupno, pred
     release pojedinačno"*. Zato ide još jedan prolaz, **bez `--grupe`**, nad istim
     izvorom — to je i reviewer-ova tačka 3.
     Opis PR-a je usklađen: tvrdnja „skupe kapije nisu puštene" je bila tačna kad je
     napisana, a od tada više nije.

111. **Pojedinačni dokaz: `DOKAZANO`, bez kvalifikatora** (09.10.2026).
     Grupni prolaz je dao `DOKAZANO (grupno)`, što je po `CLAUDE.md` §5 **razvojni**
     verdikt. Pojedinačni — svaka sabotaža sama, 16 prolaza suite — dao je:

     ```
     crvenih 16 / 16
     izvor pre/posle  14ccad6e88a3bc6b / 14ccad6e88a3bc6b -> IDENTICAN
     === DOKAZANO ===
     ```

     Posle prolaza: `git status` prazan, `vba_check` čisto (191 fajlova),
     `vba_gate --hash` = `3b19a4defd382ba7`. To je **tačno** izvor pod kojim stoji
     zeleni marker svih 12 suita, pa dokaz i pun prolaz **dele izvor** — nije reč o
     dva merenja nad dva drveta.

     **Devet od šesnaest sabotaža oborilo je i druge testove** pored svog
     (`uz jos N testa`, najviše 7). To je merenje, ne šum: pokvaren čitalac koji
     dele kartica, pregled i saldo obara **svakog** pozivaoca. Ocena broji samo da li
     je **imenovana** tvrdnja pala — i pala je u svih šesnaest. Preostalih sedam
     obara tačno svoj test.

     **Greška usput:** pozvao sam `vba_gate --hash` **dok dokaz radi** i dobio
     `7d31b606a569` — hash **sabotiranog** drveta, ne izvora; `git status` je u tom
     trenutku pokazivao ` M src-vba/modIzvestaj.bas`, što je ubrizgana mutacija.
     Mid-run merenje potpisa je besmisleno: merodavan je potpis koji driver uzme na
     startu. Iz istog razloga operater **ne sme** da pusti compile dok dokaz radi —
     projekat je tada namerno pokvaren, a driver vozi svoju instancu Excela.

     **Compile je jedina otvorena kapija.** Marker stoji nad `izvor 96c6757e2ded /
     git a4e8729e` (08.10.2026 14:39), dakle **pre** rasturanja `#204`. Opis `#408`
     je usklađen i proveren grep-om: tvrdnja „skupe kapije nisu puštene" je
     obrisana, stanje dokaza prepisano, hronologija 93–110.

112. **Compile zatvoren — sve kapije nad JEDNIM izvorom** (09.10.2026).
     Operater je pustio `ImportAllVBA`, pa `Alt+F11 → Debug → Compile VBAProject`.
     Prošlo je, i `--mark-compile` je vezao potvrdu za otisak:

     ```
     izvor    3b19a4defd382ba7
     ugovor   5b73e8d90fc359ff  (fixture b641807d, kapija 34c739f6,
                                 runner 4a9a70bf, golden 9a01af15, verzija 4)
     sveska   otkup_test.xlsm   45ac0c565c336f51
     compile  potvrdjen 2026-10-09T14:34:34 nad 3b19a4defd38
     suites   12 / 12 OK        vba_gate --require-green   RC=0
     ```

     **Nijedna suite ne stoji kao `DRUGI IZVOR`.** Compile, zeleni marker i oba
     sabotažna dokaza (grupni i pojedinačni) dele **jedan** izvor. To je ono što
     marker i treba da tvrdi, a kroz ceo ovaj rez nije mogao: compile je od 08.10.
     stajao nad `a4e8729e`, dakle **pre** rasturanja `#204`, i marker ga je uredno
     prijavljivao kao stariji izvor umesto da ga prećuti.

     Opis `#408` je usaglašen istog trena, jer je do tada nosio tvrdnju
     „otvoren je samo compile" — tačnu kad je napisana, netačnu čim je marker legao.
     Ostaje jedino **finalni reviewer prolaz**; posle njega se izvor ne dira, jer bi
     svaka izmena oborila i compile i oba dokaza.

113. **Kanonska knjiga ulazi u integritetni skener — i usput je nađena rupa u
     štampi** (09–10.10.2026).
     Reviewer je posle `#408` imenovao ovo kao prvu stavku: *„Chk_AMB_KnjigaIntegritet
     koji koristi isti canonical validator, umesto da se AMB invarijante ponovo
     implementiraju u modIntegritet"*. Grana `claude/amb-integritet-knjige`, kod
     `7d43ca0c`.

     **Nalaz koji je odredio oblik rešenja:** čitalac knjige je fail-closed —
     `KnjigaIntegritet` diže `AMB_ERR_KNJIGA_KVAR`, a ta konstanta je **`Private`**.
     Provera koja bi samo pozvala čitaoca ne bi prijavila nalaz nego **oborila ceo
     `RunIntegritetProvere`**, i to na PRVOM pokvarenom redu — operater ne bi video
     ni ostale. Zato jedna implementacija, dva režima:

     ```
     KnjigaIntegritet(data, kol, sourceName, Optional nalazi As Collection)
         nalazi Is Nothing  -> Err.Raise   (zatecen ugovor citaoca, NEPROMENJEN)
         nalazi prosledjen  -> nalazi.Add  (audit)
     Public  modAmbalaza.AmbKnjigaNalazi() As Collection
     Private modIntegritet.Chk_AMB_KnjigaIntegritet   -- samo prikazuje
     ```

     Vraća `Collection`, ne niz: `modIntegritet` već ima `CollToArray`, pa bi treća
     kopija konvertera bila treća stvar koja može da divergira.

     **Kapija je oborila prvi pokušaj testa, i bila je u pravu.** Napisao sam dve
     sabotaže uz **jedan** test; `vba_check` je odbio:

     ```
     KATALOG: amb-int-sakupljac-ne-obilazi-knjigu: deli tvrdnju sa
              amb-int-skener-ne-gleda-knjigu -- test ih ne razlikuje
     ```

     Tačno: i neregistrovana provera i sakupljač koji ne obilazi knjigu daju **isto**
     — prazan ekran. Rešenje nije bilo obrisati sabotažu nego **razdvojiti slojeve**:
     `#210` meri da nalaz STIGNE do ekrana, `#211` da ga jezgro NAPRAVI. `#211` nosi i
     tvrdnju koja se lako gubi: **audit režim ne sme da oslabi čitaoca** — saldo nad
     istim pokvarenim redom i dalje mora da padne fail-closed.

     **Dokaz nad `bce9b4b93b590498`** (ugovor `1efeb85b947ec88c`):

     ```
     RunAllTests              211 / 0      RunStornoTestSuite       164 / 0
     RunBusinessFlowProSuite 2436 / 0      RunBankaImportTestSuite  241 / 0
     palih suita                0 / 12     --require-green            RC=0

     amb-int-skener-ne-gleda-knjigu        OK                    -> #210
     amb-int-sakupljac-ne-obilazi-knjigu   OK (uz jos 1 testa)   -> #211
     amb-pisac-dupli-ambid-u-saldu         OK (uz jos 1 testa)   -> If->ElseIf
     === DOKAZANO ===   izvor pre/posle 54ef3ea8c19ca240 IDENTICAN
     ```

     Druga sabotaža obara i `#210` — tačno i očekivano (prazan sakupljač = prazan
     ekran). Razdvojenost je ipak dokazana, jer **prva obara SAMO `#210`**.

     **Tri `P3` koja reviewer imenuje, sva tri potvrđena i nijedno ne blokira:**
     audit ne prijavljuje duplikat kad je prethodni red sa istim `AmbID`-em
     strukturno neispravan (taj red **jeste** prijavljen, pa duplikat izlazi u
     sledećem prolazu — otkrivanje u dva koraka, ne prećutan kvar); agregacija više
     nalaza nije zasebno dokazana; cena jednog dodatnog prolaza kroz `tblAmbalaza`
     nad **pravom** knjigom nije merena.

     **Usput nađeno nešto teže — `KI-009`.** Reviewer je prijavio da se storniran
     ambalazni dokument može ponovo odštampati. Potvrđeno u izvoru, i komentar koji
     sam napisao u `10c-2` (*„storniran dokument nema aktivan red, štampa se odbija"*)
     je **netačan**: storno dodaje kontra-stav a original ne dira, pa mu `StornoOd`
     ostaje prazan — a `AmbDokRedMapa` zadržava baš takve redove.
     Pri proveri tog nalaza izmereno je nešto šire: **`DOK_TIP_AMBALAZA_DOKUMENT` ima
     NULA pojava u svim `modScr*.bas`.** `Select Case` u `modScrIzvestaji:2775` ga ne
     poznaje, pa kanonski dokument pada u `Case Else` — `„štampa nedostupna"`. Grana
     `StampajAmbDokument`, napisana u `10c-2`, **ne doseže se iz UI-ja**. To je
     **regresija sposobnosti** koju je uveo `10b-2` i promakla je meni i dvama review
     krugovima. Posledica za plan: (a) i (b) se **ne smeju razdvajati** — dodavanje
     `Case`-a bez storno kapije otvara rupu istog trena. Parkirano kao `KI-009`.

     **Compile zatvoren nad istim izvorom** (10.10.2026 15:57:33 nad `bce9b4b93b59`).
     Ništa iz dokaza `#408` nije važilo za ovu granu — i izvor i ugovor su se
     promenili (`3b19a4defd38` → `bce9b4b93b590498`, `5b73e8d90fc3` →
     `1efeb85b947ec88c`), pa je sve mereno iznova. Posle compile-a `--status` daje
     `DRUGI IZVOR = 0`, a `--require-green` `RC=0`.

     > Prva verzija ove stavke je završavala rečenicom *„otvoreno je samo compile"*.
     > Bila je tačna kad je napisana i netačna čim je marker legao — review `#409` ju
     > je s pravom prijavio kao neusaglašenost evidencije. Ispravljeno pre merge-a;
     > docs ne menjaju otisak izvora, pa nijedan marker nije pao.

     **Verdikt review-a `#409`: `GO ZA MERGE`** — `P0 0`, `P1 0`, `P2 0 novih`,
     četiri neblokirajuća `P3`, CI `GREEN`. Potvrđeno je i ono što je trebalo:
     *„produkcioni `AmbSaldoPoNalogu` zaista poziva `KnjigaZaCitanje`, pa novi audit
     režim ne zaobilazi postojeću integritetnu kapiju"*.

     **Četvrti `P3`, njihov nalaz i moja greška:** `#211` meri `Err.Number <> 0`, a
     ne **baš** `AMB_ERR_KNJIGA_KVAR` — pa bi i nepovezana VBA greška zadovoljila
     tvrdnju *„audit režim NE slabi čitaoca"*. Orakl je širi od pravila koje meri.
     Konstanta je `Private` u `modAmbalaza`, pa tvrdnja traži ili javni ulaz ili
     merenje teksta poruke. **Ne ispravlja se u `#409`** (review: *„ne bih sada
     dodavao ništa drugo"*) — ide uz rez `KI-009`, koji ionako dira testove.

114. **`KI-009`: kanonska štampa vraćena, i dokaz koji je uhvatio ORAKL** (10.10.2026).
     Grana `claude/amb-stampa-ki009`, glava `d2772eaf`. Tri kvara u jednom PR-u, jer
     popravka jednog otvara drugi:

     ```
     (a) DOK_TIP_AMBALAZA_DOKUMENT: 0 pojava u svim modScr*.bas -> svaki NOV
         revers je isao u Case Else, "stampa nedostupna". Grana StampajAmbDokument
         iz 10c-2 NIJE bila dosegnuta ni jednom. Regresija sposobnosti iz 10b-2.
     (b) storno DODAJE kontra-stav a original NE dira -> original ostaje u
         AmbDokRedMapa, pa mapaRed.Exists vraca True i posle storna.
     (c) broj = ambDokID, datum = Date PRE citanja zaglavlja -> papir sa
         tehnickim ADK-<hex> i danasnjim datumom.
     ```

     Nalaze (b) i (c) prijavio je nezavisan review; **(a) je izmereno pri proveri
     tog nalaza** i ispalo teže od prijave. Komentar koji sam napisao u `10c-2`
     (*„storniran dokument nema aktivan red"*) bio je **netačan** — komentar ne može
     da bude kapija.

     **Četvrti kvar je napravila sama popravka.** Otvaranje UI rute poslalo je na
     reversov šablon **svaku** vrstu kanonskog dokumenta, a ima ih četiri. `NABAVKA`
     ima par `SpoljniSvet → Stanica`, pa joj je **jedna** strana Stanica i par-kapija
     je pušta — dobila bi papir sa oznakama za kooperanta i praznim partnerovim
     ID-em. Rešeno **belom** listom: nova vrsta mora biti **svesno** puštena.
     Usput je nađeno da je odbijanje kupčevog reversa bilo **uže od pravila** —
     stajalo je u `Else` grani, a razlog je poslovni i ne zavisi od para naloga.

     **NAJVREDNIJI NALAZ REZA: `dokaz.py` je uhvatio grešku u ORAKLU, ne u kodu.**

     ```
     amb-stampa-ne-pita-za-storno   PALA DRUGA TVRDNJA:
                                    "odbijanje imenuje STORNO, a ne drugi razlog"
     === NIJE DOKAZANO ===
     ```

     `#212` nije gasio izlaz. Sa ugasenom storno kapijom dokument je **stvarno**
     krenuo da pravi PDF, to je puklo u headless Excelu, `Err.Number` je postao
     `<> 0` — i tvrdnja *„štampa storniranog je odbijena"* se zadovoljila **slučajnim
     padom** umesto kapijom. Suite je bio `213 / 0` i pre i posle; razliku je
     napravilo jedino to što ocena traži da **imenovana** tvrdnja padne **PRVA**.
     Popravka: izlaz na `"OFF"` (prolazi kroz `DocResolveMode` netaknut, `Select
     Case` nema granu za njega), pa se ceo produkcioni lanac vrti a papir se ne
     pravi.

     **Dva `P3` o izolaciji testa, oba iz review-a i oba moja:** cleanup nije bio
     bezuslovan (greska u piscu preskace vracanje `OM_IZDAVANJE_PRINT_MODE`), i
     posle `On Error GoTo 0` je čitalac kartice išao **bez** handlera. Oba
     zatvorena; red je sada `pročitaj modePre → naoružaj handler → menjaj config`,
     jer bi naoružavanje **pre čitanja** upisalo **prazan** režim da `GetConfigValue`
     pukne.

     **SVESNO PRIHVAĆEN DUG (odluka operatera, 10.10.2026).** U `#213` je
     `SetConfigValue "OFF"` ostao **pre** `On Error GoTo CleanUp`, dok `#212` to već
     radi obrnuto. Asimetrija je stvarna i review ju je imenovao kao `P3`. **Ne
     zatvara se.** Razlog: ujela bi samo ako `SetConfigValue` pukne **pošto je
     delimično promenio** config, a on upisuje **jednu ćeliju u jednoj tabeli** —
     delimična izmena nije moguća. Cena zatvaranja je pun ponovni prolaz (~37 min)
     **i još jedan ručni compile**, za rizik koji je praktično nula. Ako se
     `SetConfigValue` ikad proširi na više ćelija ili tabela, ovaj dug se **mora**
     zatvoriti.

     **Dokaz nad `b696c49e88e6ec1f`** (ugovor `231f920c6423ac94`):

     ```
     RunAllTests              213 / 0      RunStornoTestSuite       164 / 0
     RunBusinessFlowProSuite 2436 / 0      RunBankaImportTestSuite  241 / 0
     palih suita                0 / 12     --require-green            RC=0
     compile  potvrdjen 2026-10-10T19:31:59 nad b696c49e88e6   DRUGI IZVOR 0

     dokaz amb-stampa-  4 / 4  \  oba nad potpisom e2fc6c4bb86e4495,
     dokaz amb-int-     2 / 2  /  izvor pre/posle IDENTICAN   === DOKAZANO ===
     ```

     **Šta dokaz NE tvrdi:** da je papir napravljen ni šta na njemu piše. Dva
     razloga: izlaz je u testu ugašen, i `OutputIzdavanjeAmbalaze` ima **sopstveni**
     `On Error GoTo EH` koji grešku samo loguje — pa *„nema greške"* ne bi bio orakl
     uspeha ni sa uključenim PDF-om. Sadržaj papira ostaje na **ručnom smoke-u**, i
     to je uslov za merge koji je review izričito tražio.

115. **`#411` merge-ovan BEZ ručnog smoke-a — svesna odluka** (10.10.2026).
     Review je ručni print smoke postavio kao uslov za merge. **Nije izvršen**, i
     razlog nije propust: u radnoj svesci još **nema unetih podataka** nad kojima bi
     se štampao kanonski revers. Operater je odlučio da se merge-uje svejedno.

     **Šta to znači, bez ulepšavanja.** Automatski dokaz pokriva **odluke**:
     kapije (zaglavlje, storno, bela lista vrsta), rutiranje iz UI-ja, i razrešenje
     partnera — sve kroz produkcione pisce, uz četiri sabotaze `DOKAZANO` i compile
     nad istim izvorom. Ne pokriva **sam ispis**: da li papir nosi poslovni broj
     umesto `ADK-<hex>`, datum dokumenta umesto današnjeg, tačnog protivpartnera i
     količinu. To je upisano kao **`KI-010`**, otvoreno.

     **Nije regresija:** pre `#411` se kanonski dokument **uopšte nije mogao
     odštampati** — padao je u `Case Else`. Rizik je zato „papir možda nije tačan",
     a ne „papir je bio tačan pa više nije".

     Stanje `main`-a posle merge-a (`ede1ff1b`): potpis `b696c49e88e6ec1f` je
     **bajt-identičan** dokazanom, pa `--require-green` daje `RC=0`, `DRUGI IZVOR`
     je `0`, a compile marker važi i za `main`.

116. **`10d` — `pre-flight` verdikt, PRE koda** (11.10.2026). Predaja čistoj sesiji:
     sve ispod je **izmereno**, da se ne meri ponovo.

     **Zašto `10d` ide prvi:** `modScrStorno.AkcijeRacun` za `STIP_REVERSI` vraća
     **samo `STORNO`**. Operater od `10c` **ne može da ispravi revers** — može samo
     da ga stornira i unese iznova. Skinuto svesno i zapisano u kodu, uz obećanje
     `10d`. To je jedina preostala stavka gde operater ima **manje** nego pre
     refaktora.

     **Verdikt po osama:**

     ```
     DOMAIN              PROVEN   A9 + B-036 (nize)
     IDENTITY            GAP      = predmet reza
     CARDINALITY         PROVEN   AMB-INV-10, AMB-INV-11
     INVARIANTS/OWNER    PROVEN   StornirajAmbDokument_TX postoji i dokazan (#212)
     WRITERS             RELEVANTNO  menja se korekcioni trag I storno pisac
     EVENTS              TRAZI ODLUKU (nize)
     DOWNSTREAM          PROVEN   StornoVeze, StornoZurnal, Prefill, ekran
     CAPABILITY          GAP      = predmet reza
     PLATFORM            N/A
     LANDING             SERIJSKI 10d i 10e dele sedam ReversID* funkcija
     ```

     **DOMAIN, doslovno iz ugovora.** `ARCHITECTURE_CONTRACT.md` **A9**: ispravka
     dobija **nov `DocumentID` I nov poslovni broj**, veza je `IspravkaOdID` /
     `ZamenjenSaID` **po ID-u, nikad po broju**; izričito prošireno na **REV**
     odlukom 14.09.2026. `MAPA_SPOSOBNOSTI` **B-036**: prefill uzima **datum**, ne
     uzima **broj**, i polazi od **PK-a**, ne od broja.

     **Šta tačno treba preseći** (izmereno, sa mestima):

     ```
     modStornoDok:568        Case FLOW_DOC_REVERS -> RunReversCorrection(broj, opcija, ...)
     modStornoFlow:520       RunReversCorrection(brDok, dokumentTip, mode, ambID)
                             -> ReversIDRazresi -> ReversIDGranica -> ReversRedoviRID
                             -> CreateCorrectionContext(..., revID, brDok, ...)
                             -> StornoOMKoopByBrDok_TX     <- LEGACY pisac
     modStornoFlow:604       CompleteReversIspravka(correctionID, newBrDok, ...)
     modScrStorno            AkcijeRacun: vratiti ISPRAVKA segment
     ```

     Kanonska zamena je spremna: **`StornirajAmbDokument_TX`** je dokazan u `#212`,
     a identitet je **`AmbDokID`** umesto `ReversID`.

     **EVENTS — najvažniji nalaz ovog verdikta.** Ispravka reversa **ne pomera robu**:

     ```
     fizicki       NISTA -- gajbe su gde su bile
     poslovni      stari dokument povucen, nov izdat sa NOVIM brojem (A9)
     "finansijski" saldo se vraca kontra-stavom pa ponovo upisuje
     ```

     Kroz knjigu zato prolaze **ČETIRI reda** (kontra-stav + nov upis, obe strane) za
     događaj u kome se fizički **ništa** nije desilo. **Glavna tvrdnja `10d`:** saldo
     posle ispravke mora biti **identičan** ako je promenjen samo broj ili datum, a
     različit **tačno za razliku** ako je promenjena količina. **Nijedan postojeći
     test to ne pokriva.**

     **Druga zamka, imenovana unapred:** A9 traži **nov** poslovni broj. Proveriti da
     kapija jedinstvenosti (`BrojZauzetUNizu`, `ReversValidiraj`) nov broj ne odbije
     i da stari **ne oslobodi** — isti obrazac je već ujeo kod otpremnice (`KI-008`).

     **ACCEPTANCE CONTRACT** (plan dokaza, kod još ne postoji): ispravka reversa daje
     nov `AmbDokID` **i** nov broj; veza ide po ID-u; saldo se ponaša po pravilu
     iznad; stari broj ostaje zauzet; `AMB-INV-07` (nijedan realan nalog ispod nule)
     važi **u svakom međukoraku**, ne samo na kraju; ekran ponovo nudi `ISPRAVKA`.
     Negativan slučaj koji mora pasti: ispravka nad **već storniranim** dokumentom.
     Dokaz: testovi u `modTest` kroz **produkcione** pisce + sabotaze po kapiji.

     **Stanje `main`-a na predaji:** `09d9c467`, potpis `b696c49e88e6ec1f`,
     `--require-green RC=0`, `DRUGI IZVOR 0`. Otvoreno i dalje: `KI-010` (smoke
     štampe, čeka podatke) i `B5b` otpremnica `1/220926`.

117. **`10d` — ispravka našeg reversa u JEDNOM potezu** (11.10.2026). Grana
     `claude/amb-10d-ispravka`, iz stavke 116.

     **Odluke operatera (osa EVENTS iz 116):**

     ```
     oblik             JEDAN POTEZ (kao B-040 / IspravkaOtkupa_TX); dvokorak obrisan
     veza              IspravkaOdID + ZamenjenSaID na kraju tblAmbalazaDokument
                       I red u tblStornoVeze (A9 doslovno)
     RESI KASNIJE      zaseban rez -- v. nalaz 2 ispod: premisa pitanja je bila netacna
     ```

     **Šta je urađeno** (domen: `docs/DOMEN/AMBALAZA.md` 6.13a):

     - `modAmbalaza.IspravkaReversa_TX`: u **jednoj** transakciji storno starog, nov revers
       sa novim brojem i `AmbDokID`-em, trag u `tblStornoVeze` (tip `AmbalazaDokument`,
       ne `FLOW_DOC_REVERS` — stari tip nosi `ReversID` u `OldDocID`), veza po ID-u na
       oba kraja, zatvaranje traga čija se povratna vrednost **čita**.
     - Jedna kapija za ekran i pisca: `IspravkaReversaRazlog` (prazan ID, nejednoznačan
       red, storniran, vrsta ≠ `REVERS`).
     - Jezgra u transakciji pozivaoca: `StornirajAmbDokument` (bind je sada ovde —
       `AMB_BIND_DOZVOLJENI` ažuriran) i `UpisiReversUTx`; kapije broja i smera izdvojene
       (`ReversBrojZaUpis`, `ReversSmerZaUpis`) da ih ispravka prolazi **istim** putem.
     - **Storno prvi, kao otkup:** `AMB-INV-07` važi i u međukoraku.
     - Ekran: Storno nudi „Zamena reversa" (raspored i natpis od pre `10c`), klik otvara
       F7 popunjen starim i ništa ne menja; snimanje F7 ide kroz ispravku.
       `PrefillReversa` (B-036 za revers): datum, smer kroz `AmbReversSmerovi`, strane,
       tip, količina — **bez broja**.
     - Obrisano: `CompleteReversIspravka`, grana ISPRAVKA u `RunReversCorrection`,
       `ZavrsiIspravkuAko`, `ZavrsiIspravkuPitanje`, šest testova dvokoraka u
       `modTestStorno` (`TRev_TragIspravkeNosiReversID` prešao na `DUPLI` — isti trag,
       isti `OldDocID`), varijanta potvrde `StornoPotvrdaTekst(mode)`.

     **Nalazi tokom rada:**

     1. **`IdiNaRezim` ne čisti formu kad je F7 već aktivan režim** — izlazi rano
        (`key = ActiveMode`). Ispravka bi tada nasledila broj prethodnog unosa (sa
        isključenim auto-brojem ga ni predlog ne bi pregazio). Zato `ClearForm` između
        prelaska i prefill-a. **Isti obrazac postoji u `OtvoriIspravku` za prijemnicu** —
        nije diran, van reza.
     2. **Tvrdnja koju sam dao pre odluke bila je netačna:** rekao sam da je `10c` skinuo i
        REŠI KASNIJE za revers. Izmereno posle: ekran Storno reversu taj mod **nije nudio
        od prve verzije** (`dba78545`) — samo storno i zamenu. Grana u
        `RunReversCorrection` postoji, ali je sa ekrana nedostižna od ranije. Odluka
        „zaseban rez" je doneta nad pogrešnom premisom i vraća se operateru (MAPA B-025).
     3. `ARNOST` ne vidi poziv u izraznoj poziciji: `StornoPotvrdaTekst(SV_MODE_ISPRAVKA)`
        u `modTest` bi oborio compile modula. Nađen ručnim grep-om po imenu.
     4. Sabotaža `otp-nacrt-zavrsava-ispravku` je **obrisana**: vraćala je poziv
        `ZavrsiIspravkuAko`, koga više nema — oborila bi compile, ne tvrdnju.
     5. Sidro `amb-inv08-zaglavlje-bez-kapije` je postalo dvoznačno (isti poziv i u jezgru
        storna) — suženo na sledeći red pisca zaglavlja.

     **Testovi:** `modTest` 214–219 (`T_AmbIspravkaReversa_*`), svi kroz produkcione
     pisce; 219 ide kroz `Scr_BrojAkcija` → `IspravkaReversaPripremi` →
     `Scr_IspravkaReversaOtvori` → `Scr_Save` F7. Katalog sabotaža **724 → 733**
     (+10 `amb-isprrev-*`, −1). `RunStornoTestSuite` gubi šest testova dvokoraka, pa mu
     broj tvrdnji **mora** pasti — to je merenje, ne regresija.

     **Verifikacija:** zelene SAMO jeftine kapije — `vba_check` (191 fajl, katalog 733),
     `vba_check --self-test`, `gen_schema_module --check`, `who_writes --check`
     (regenerisan: `modAmbalaza` u transakciji `tblStornoVeze`, kao `modOtkup`),
     `--check-ownership`, `popis_citalaca --check`. **`run_vba` i `dokaz.py` NISU
     pušteni — čeka se reviewer GO.** Šema je promenjena (`0B3506AD`), pa fixture mora
     da se regeneriše pre prvog prolaza. Compile ostaje ručna kapija.

     **Ostaje:** odluka o REŠI KASNIJE za revers (nalaz 2); ispravka kupčevog reversa ne
     postoji (glasno, kapija je odbija); `KI-010`; `B5b`; `10e`.

118. **Review `#414` (`c53c7394`): `P2` jednoznačnosti poslovnog reda — i jedan `P3`
     koji je bio nemeren i kod otkupa** (11.10.2026).

     Verdikt: `NO-GO` za merge, `GO` za nastavak i ciljane testove posle `P2`.
     `P0 0`, `P1 0`, `P2 1`, `P3 2`; static CI `GREEN`.

     **`P2` je stvaran, i širi nego što ga review opisuje.** Kapija je proveravala
     jedno zaglavlje, ali ne i jedan poslovni red u knjizi. Mapa prikaza
     (`AmbDokRedMapa`) uzima **prvi** red, storno povlači **sve**. Review piše da je to
     „rupa nad nejednoznačnim podacima". Merenje kaže da je to **ugovorom dozvoljen
     oblik**: `AMB-INV-10` izričito pušta dokument sa više originalnih redova (dva tipa
     ambalaže, oba smera prema istom partneru). Današnji pisac reversa piše jedan red,
     ali kapija to ne sme da pretpostavi:

     ```
     ADK-100  REVERS
       red A  Stanica -> K1  IZDATA_PRAZNA  3     <- forma vidi samo ovaj
       red B  K1 -> Stanica  POVRAT_PRAZNE  1
     ispravka -> storno A i B -> upis JEDNOG reda iz forme   (dva kretanja tiho postaju jedno)
     ```

     **Popravka bez drugog validatora.** Pravilo ide u **postojeću** kapiju
     (`IspravkaReversaRazlog`), koju već pitaju i ekran i pisac. Definicija „poslovnog
     reda" izdvojena je u `PoslovniRedDokumenta` i deli je mapa prikaza. Kapija broji
     kroz **strogu** `AmbDokBrojPoslovnihRedova`, koja grešku čitanja diže, a ne guta,
     jer kapija ne sme da pročita „jedan" iz knjige koju nije pročitala do kraja. Uslov
     je tačno jedan: nula je zaglavlje bez knjige, dakle kvar, i takođe pada.

     **`P3` o ugnježdenim transakcijama.** Izmereno: unutrašnji `CommitTx` samo zakazuje
     snimanje (`OnTime`, 60 s neaktivnosti), a `OnTime` ne prekida kod koji se izvršava,
     pa snimanje vidi stanje posle spoljnog commit-a ili rollback-a. Rollback reda koji
     je ugnježdeni `CreateCorrectionContext` već commit-ovao **nije bio meren nigde**: ni
     ovde, ni kod ispravke otkupa, čiji rollback test pada **pre** traga (na stavci).
     Zato postoji uzak seam `IspravkaReversaPadniTest`, tvrdo gejtovan test režimom.
     On obara ispravku tik pre commit-a, kad su storno, nov revers, trag i veza već
     napisani, a `#221` meri da se vrati **sve**.

     **`P3` o testu `#219`** (forma se ne gradi) je prihvaćen. Raspored polja na pravoj
     formi ostaje na ručnoj proveri, i ne gradi se nova test platforma.

     **Testovi i sabotaže:** `#220` (dva poslovna reda, napravljena produkcionim
     primitivima, odbijena u kapiji, pisac odbija **pre** storna, ekran ne nudi zamenu)
     i `#221` (pad posle traga). Sabotaže `amb-isprrev-vise-poslovnih-redova` i
     `amb-isprrev-trag-van-snapshota`. Katalog **733 → 735**.

     **Verifikacija:** jeftine kapije zelene. `run_vba` i `dokaz.py` **nisu pušteni**.
     Review je dao `GO` za ciljane testove posle `P2`, ali sama ispravka `P2` još nije
     pregledana. Pre prvog prolaza fixture se regeneriše (šema je promenjena u `10d`).

119. **Dokaz `#414` nad konačnim izvorom — i jedna placebo tvrdnja koju je uhvatio**
     (11.10.2026). Reviewer je posle stavke 118 dao `GO za execution proof`.

     **Fixture:** regeneracija iz starog donora **nije** dodala dve nove kolone
     `tblAmbalazaDokument`, pa je tabela ostala sa 12 kolona umesto 14. Uzrok:
     `make_fixture.py` nosi prepisane spiskove kolona (`ENSURE_TABLES` /
     `ENSURE_COLS`). Pisac bi tada pao na `SchemaReadyOrFail` u svakom testu sa
     reversom, pa bi se merila sveska, a ne kod. Dopunjeno istim obrascem kao kolone
     knjige u `10c` (`1e367c15`). Nov potpis fixture-a je `54298ba411e550be`; stari
     fixture je sačuvan kao `otkup_test_pre_10d.*`.

     **Prvi dokaz (izvor `14c213baaddc`): 12/13, `NIJE DOKAZANO`.**
     `amb-isprrev-isti-broj-prolazi` → `NE OBARA NISTA`. Uzrok je u **mom** kodu:
     poređenje sa starim brojem stajalo je **posle** `ReversBrojZaUpis`, pa je stari
     broj na istoj stanici i danu prvo odbijala kapija zauzetosti. Njena poruka
     **takođe** pominje `NOV broj (A9)`, pa je tvrdnja `#217` bila zadovoljena i bez
     A9 kapije. Komentar u kodu je pri tom tvrdio suprotan redosled.
     Popravka (`4254b2d1`): **zadat** broj se poredi pre kapija broja, **izračunat**
     posle (jedan pomoćnik, `RequireNovBrojIspravke`). Tvrdnja sada traži tekst koji
     nosi **samo** poruka pravila A9.

     **Konačno (izvor `44cb49d0ea21`, ugovor `8694e2348800`, fixture `580ff34c`):**

     ```
     RunAllTests              221 / 0      (#214-221 svi OK)
     RunBusinessFlowProSuite 2436 / 0
     RunBankaImportTestSuite  241 / 0
     RunStornoTestSuite       131 / 0      (bilo 164: obrisano sest testova dvokoraka)
     ostalih 8 suita             OK        12 / 12 ZELENO, marker upisan
     --require-green          RC=0

     dokaz.py amb-isprrev + amb-inv08-zaglavlje-bez-kapije
       crvenih 13 / 13, izvor pre/posle 778a7f7b9411cde3 IDENTICAN
       === DOKAZANO ===
     ```

     Pojedinačno: `amb-inv08-zaglavlje-bez-kapije` OK (uz još 1 test, sidro suženo u
     117) · `bez-storna` OK (uz još 4) · `kolicina-pomerena` OK (uz još 1) ·
     `veza-po-broju` OK (uz još 1) · `trag-ostaje-otvoren`, `isti-broj-prolazi`,
     `storniran-prolazi`, `svaka-vrsta`, `upis-pre-storna`, `ekran-ne-nudi`,
     `ekran-ne-prosledjuje`, `vise-poslovnih-redova`, `trag-van-snapshota` — OK.

     **Compile zatvoren** (11.10.2026): operater je potvrdio `Debug → Compile` posle
     uvoza; marker je upisan nad `44cb49d0ea21`. `--require-green --require-compile`
     daje `RC=0`, a `DRUGI IZVOR` je `0`. Uvoz ne beleži otisak, pa marker počiva na toj
     potvrdi. Review `481e4e91`: `CODE REVIEW GO`, merge čeka samo ručnu proveru F7.

     **Ručna provera ide nad kopijom fixture-a** (`tests/fixtures/rucna_provera_10d.xlsm`,
     trenutni kod, netaknuti podaci). Operater nema svoje podatke, a fixture već nosi sva
     tri slučaja: `2/IZV` (jedan red → zamena), `1/IZV` (**dva** reda, 12/1 + Letvarica →
     samo storno, P2 nad stvarnim podatkom) i `4/IZV` (nabavka → samo storno).

     **Nalaz usput, van `10d`:** `NabaviAmbalazu_TX` nema **nijednog produkcionog
     pozivaoca**, nego samo testove. Na novoj instalaciji stanica zato počinje od nule, a
     izdavanje praznih se odbija (`AMB-INV-07`). Komentar pisca to izričito predviđa,
     a ulaz na ekranu ne postoji niti je zapisan kao odložen. Izdvojeno kao zaseban zadatak.

120. **Ručna provera `#414` nad kopijom fixture-a — ispravka radi, a usput su izmerena dva
     stara kvara štampe** (11.10.2026). Review `481e4e91`: `CODE REVIEW GO`; merge čeka
     ručnu proveru.

     **Potvrđeno (operater, `tests/fixtures/rucna_provera_10d.xlsm`):**

     ```
     2/IZV (prijem 10)  -> "Zamena reversa" -> F7 popunjen: datum 15.03.2026, stanica,
                           smer "Prijem koop.", kooperant, 12/1, 10; broj NOV (1/150326-2)
     kolicina 8, snimi  -> 2/IZV Stornirano=Da, ZamenjenSaID -> nov ADK
                           1/150326-2 IspravkaOdID = ADK-IZV-2; papir nosi nov broj
     obican revers      -> 1/150326-3 snimljen, forma ociscena
     ```

     Vrsta voća se u prefill-u ne popunjava, jer je revers ne čuva, a validacija je traži pa
     je operater bira. Da li revers treba da čuva vrstu, koju papir nosi, poslovna je odluka
     za kasnije.

     **Završetak provere (isti dan):** `1/IZV` (dva poslovna reda) nudi **samo** „Storno bez
     zamene" — P2 nad stvarnim podatkom ✓. Odustanak (prelazak iz zamene u F1) ne upisuje
     ništa, a otvorena zamena se gasi (`SelectModeCore` → `OtkaziIzmenuNacrta` →
     `Scr_IspravkaOtkazi`) ✓. `4/IZV` (nabavka) se na listi „Reversi" uopšte ne vidi, i to je
     **ispravno**: lista namerno prikazuje samo reverse (test 202). Korak iz mog uputstva je
     bio pogrešan. Kupčev revers fixture nema, pa ručno nije proveren; kapija po vrsti je
     ista kao za nabavku (`#217`, sabotaža `amb-isprrev-svaka-vrsta`).

     **Uzgredno, nije iz `10d`:** prelazak režima **nikad ne čisti** uneta polja. Datum 15.03 i
     99 kom iz napuštene zamene ostali su u F1. Komentari „`SelectMode` čisti formu" (u
     `modScrStorno.OtvoriIspravku`, `modOtkupUI.IdiNaEkran` i moj korak 2 u
     `OtvoriIspravkuReversa`) su zato netačni; baš zbog toga postoji `ClearForm` pri otvaranju
     zamene. Komentar se ispravlja u zasebnom PR-u (`KI-011`), da se ne obori dokaz `#414`.

     **`KI-011` — lažan neuspeh posle snimanja reversa.** Posle uspešne ispravke ljuska je
     javila `Scr_Save -> 9 Subscript out of range`, a forma je ostala popunjena. U logu za
     taj trenutak nema `ERROR`-a, samo tri uredna commit-a. Uzrok je izmeren u harness-u
     nad kopijom sveske (bez test režima, sa PDF štampom): `modSetup.LogSetup` ostavlja
     prigušenu grešku 9 kad lista `SETUP_LOG` nema. Greška putuje kroz `EnsureDocFolder` i
     štampu do `ScrSave`, koja je posle `Application.Run` čita kao neuspeh. Pogađa i
     **običan** prijem (3/3 u harness-u), a postoji od `17f44ffa` (27.06.2026). Testovi je ne
     vide jer gase štampu. **Neobjašnjeno:** drugi, običan revers operatera snimljen je bez
     greške, iako `SETUP_LOG` ni tada nije postojao.

     **`KI-012` — `1.dec` umesto `12/1` na papiru** (`modPrint.bas:1972`): tip se upisuje u
     ćeliju opšteg formata, pa ga Excel pretvara u datum. Pogađa svaki revers iz F7.

     **Odluka operatera:** oba kvara idu u **zaseban PR posle `#414`**, a `#414` ostaje sa
     zelenim dokazom nad `44cb49d0ea21`. Rizik do tada: posle štampe operater vidi lažan
     neuspeh i nepočišćenu formu, pa obican revers može da snimi dvaput.

121. **`10d` ostatak: undo i istorija su dve ODLUKE, ne kod** (11.10.2026). Grana
     `claude/amb-10d-ostatak`, iz `main`-a posle merge-a `#414` (`dc5a32c4`).

     **Pre-flight verdikt:** `DOMAIN` GAP (ODL-19 je izričito čekao poslovnu semantiku),
     `CARDINALITY` PROVEN (`AMB-INV-06`: storno se ne stornira), `CAPABILITY` GAP. B-046
     „Vrati storno" je od `10b-2` odbijen za svaki otkup, otpremnicu i prijemnicu sa
     gajbama, a revers i nabavku od `10c` lista uopšte ne nudi, jer
     `StornirajAmbDokument_TX` nije u žurnalu. Istorija: kontra-stav **nasleđuje datum
     originala** (kartica vidi storno unazad), a D-057 seče po **redosledu upisa** (papir
     se ne menja). To su bile dve semantike, nijedna zapisana i nijedna pokrivena testom
     sa stornom.

     **Odluke operatera:**

     ```
     AMB-10-ODL-25  storno dokumenta sa ambalazom je KONACAN; put je ponovni unos
                    (nov broj, A9). B-046 se svesno suzava na dokumente bez ambalaze.
     AMB-10-ODL-26  dve semantike istorije, zapisane i pinovane testom:
                    kartica po periodu -> storno vazi unazad
                    ponovna stampa (D-057) -> papir kakav je bio
     ```

     **Šta je urađeno:** kod se ne menja osim **poruke** kapije undo-a. Ona sada imenuje
     ODL-25 i put, jer je ekran Oporavak prikazuje. Dodato: tvrdnja nad putem u
     `Test_Amb_UndoStornaOdbijenNadKnjigom`, testovi `#222` (kartica, storno unazad, i
     kontra-stav nosi datum originala) i `#223` (ponovna štampa ista posle storna ranijeg
     dokumenta, dok se tekući saldo menja), i tri sabotaže. Katalog **735 → 738**.

     **Imenovan dug:** lista „Vrati storno" pokazuje „moguce" za operaciju koju kapija uvek
     odbija (`OpStatusFor` gleda samo ćeliju). Ispravka traži sken knjige po operaciji, a
     cena nije merena.

     **Verifikacija:** jeftine kapije zelene. `run_vba` i `dokaz.py` nisu pušteni, čeka se
     reviewer GO.

122. **Review `#415` (`f0f16b62`): `P2` ponovna štampa bez granice — reprodukovano, pa
     odlučeno** (11.10.2026). Verdikt: `NO-GO` za zatvaranje ODL-26. Nalazi su `P0 0`,
     `P1 0`, `P2 1`, `P3 1`, a static CI je `GREEN`.

     **`P2`:** `GetKooperantAmbOpening` je, kad ne nađe nijedan red bloka, uzimao **celu
     današnju knjigu** („legacy" fallback). Otkup bez gajbi nema red u knjizi
     (`KnjiziOtkupAmbalazu` piše samo količinu veću od nule), a validacija ga pušta kad je
     stroga validacija isključena i unos neto. Štampa otkupnog lista **uvek** ispisuje
     početno stanje i saldo. Prva štampa je zato bila tačna, a ponovna nije.
     **Reprodukovano pre ispravke** (`Test_Amb_PonovnaStampaBezRedaKnjige`, BFP 2441/1):
     očekivano 17740, dobijeno 17743, tačno kasniji revers od 3.

     **Odluka operatera `AMB-10-ODL-27`:** bez granice nema stanja. Čitalac javlja
     `outPoznato = False`, papir početno i saldo štampa kao „-", a fallback se briše. Nije
     uveden u `#415`; ODL-26 ga je samo razotkrio.

     **`P3` (status „moguce" u listi „Vrati storno")** reviewer prihvata kao zaseban zadatak,
     uz uslov da kapija ostane fail-closed, što i jeste.

     **Ciljano posle ispravke:** `RunAllTests` 223/0, BFP 2442/0 (2436 + 1 tvrdnja nad putem
     undo-a + 5 u reprodukciji). Katalog **738 → 740**. Pun prolaz, dokaz za pet novih
     sabotaža i compile čekaju reviewer GO.

123. **Review `#415` (`5e95876c`): dokazni `P2`, pa dokaz nad konačnim izvorom** (11.10.2026).
     Verdikt je bio `MERGE PENDING`, a `GO` za finalno testiranje posle male ispravke.

     **Dokazni `P2`:** pozitivna kontrola `poznato = True` u `#223` stajala je **poslednja**.
     Pod sabotažom `amb-istorija-granica-nikad` čitalac vraća 0, pa je prva padala računska
     tvrdnja `o1 - preR1 = 3` i zaklanjala imenovanu. Dokaz bi javio „PALA DRUGA TVRDNJA",
     isti obrazac kao A9 i `#212`. Tvrdnja je premeštena odmah iza prvog čitanja
     (`a77502e4`); produkcioni kod nije diran. Ista provera redosleda urađena je i za ostale
     četiri nove sabotaže.

     **Konačno (izvor `87c8d3542702`, ugovor `e288c5255f1a`, fixture `580ff34c`):**

     ```
     RunAllTests              223 / 0      (#222, #223 OK)
     RunBusinessFlowProSuite 2442 / 0
     RunBankaImportTestSuite  241 / 0      RunStornoTestSuite 131 / 0
     12 / 12 ZELENO, marker upisan, --require-green RC=0

     dokaz.py amb-istorija + amb-undo-odbijenica-bez-puta
       crvenih 5 / 5, izvor pre/posle f3a7fa2c78ec8932 IDENTICAN -> DOKAZANO
     ```

     **Ručna provera papira ✅ (12.10.2026)**, sveska `rucna_provera_10d_ostatak.xlsm`,
     validacija unosa isključena, neto unos. Otkup `1/121026` (100 kg, 0 gajbi) nema red u
     knjizi (22 reda pre i posle), pa je to baš slučaj iz ODL-27:

     ```
     10:25:19  prva stampa      Pocetno stanje 12/1 x -   Primljeno/Izdato 12/1 x 0   Saldo 12/1 x -
     10:58:21  revers 8080C     izdato 30, pocetno 18 -> novo stanje 48
     10:58:53  ponovna stampa   Pocetno stanje 12/1 x -   Saldo 12/1 x -   (nepromenjeno)
     ```

     Usput su nađene dve stvari, nijedna iz `#415`:
     - **Oštećeno zaglavlje (nije reprodukovano).** Prvi pokušaj reversa odbijen je sa
       „Nedostaje kolona `AmbDokID`": ćelija `AmbalazaDokument!A1` sadržala je razmak. Nijedna
       procedura u `src-vba` ne piše u zaglavlje tabele. Sveska se otvara na tom listu sa
       izabranom A1, pa je verovatan uzrok taster otkucan u Excel tabelu. Kapija je odbila
       upis i tačno imenovala kolonu. Posle ručne popravke A1 revers je prošao.
     - **Pogrešno ime dokumenta u porukama.** `DOK_MSG_GRESKA_PRI_CUVANJU` („…čuvanju
       otpremnice") dele i zbirna, prijemnica, isplata, uplata i revers, a prefiks
       `OTKUP_ERR_GRESKA_PRI_UNOSU` („…unosu otkupa:") deli 25 validatora. Ide kao zaseban
       zadatak. Odbijenica deficita pri izdavanju preko stanja je ispravna poslovna kapija.

     **Compile ✅** (12.10.2026 11:37, isti izvor `87c8d3542702`):
     `vba_gate --require-green --require-compile` RC=0.

     **Otvoreno:** finalni merge review.

124. **Review `#415` (`87f7dbf8`): kvar knjige nije nepoznato stanje** (12.10.2026).
     Verdikt `NO-GO`, jedan `P2`; ostale kapije (papir, compile, CI #935) reviewer je
     prihvatio.

     **`P2`:** `GetKooperantAmbOpening` je imao handler koji je i **kvar knjige**
     (`KnjigaKvar` → `AMB_ERR_KNJIGA_KVAR`, pad šeme) vraćao kao `outPoznato = False`.
     Papir pokvarene knjige je zato bio isti kao papir otkupa bez gajbi (ODL-27, „-").
     **Reprodukovano pre ispravke:** `#224 T_AmbIstorija_KvarKnjigeNijeNepoznato` nad
     starim kodom daje `RunAllTests` 224/1, i pada baš „kvar knjige: citac pocetnog
     stanja dize gresku, ne vraca nepoznato". Oba preduslova su prošla (list se nad čistom
     knjigom popunjava, pokvaren red je pre kvara nosio odredište).

     **Ispravka:** handler je uklonjen. Kvar ide pozivaocu, a handler
     `FillOtkupSablon` list ne popunjava: nema papira, razlog je u logu. Legitimno
     nepoznato ostaje kakvo je bilo, jer `Test_Amb_PonovnaStampaBezRedaKnjige` (BFP)
     prolazi bez greške.

     **Ciljano posle ispravke** (izvor `1d4a086844e7`): `RunAllTests` **224/0**, BFP
     **2442/0**. Katalog **740 → 742**: `amb-istorija-kvar-kao-nepoznato` (čitač guta
     kvar) i `amb-istorija-stampa-guta-kvar` (štampa guta kvar).

     **Isto gutanje, starije od 10d:** `GetAmbalazeStanje` (vraća `Empty`) i
     `GetStanicaAmbSaldo` (vraća 0) daju na papiru 0 nad pokvarenom knjigom. Ide kao
     zaseban zadatak, jer reviewer traži da `#415` ostane uzak.

     Review `448e505b`: P2 **RESOLVED**, test `#224` GOOD, GO za završni dokaz.

     **Konačno (izvor `1d4a086844e7`, ugovor `e38b761cc34d`, fixture `580ff34c`):**

     ```
     RunAllTests              224 / 0      (#224 OK)
     RunBusinessFlowProSuite 2442 / 0
     RunBankaImportTestSuite  241 / 0      RunStornoTestSuite 131 / 0
     12 / 12 ZELENO, marker upisan, --require-green RC=0

     dokaz.py amb-istorija-kvar-kao-nepoznato amb-istorija-stampa-guta-kvar
       crvenih 2 / 2, izvor pre/posle 218adcf5ff5dc64d IDENTICAN -> DOKAZANO
     ```

     **Compile ✅** (12.10.2026 12:57, isti izvor `1d4a086844e7`):
     `vba_gate --require-green --require-compile` RC=0.

     **Otvoreno:** finalni merge review.

125. **`10e` — `pre-flight` i mapa merenja, PRE koda** (12.10.2026). Grana
     `claude/amb-10e-brisanje-starog-modela`, sa `main`-a `47e590de`. Mapa simbol →
     pozivaoci → dostupnost → klasa je u `AMBALAZA.md` 6.13c i ovde se ne prepisuje.

     **Verdikt.** `DOMAIN`, `IDENTITY`, `CARDINALITY`, `INVARIANTS/OWNER`, `WRITERS`,
     `DOWNSTREAM` i `PLATFORM` su PROVEN, a `EVENTS` je N/A, jer se nijedan poslovni
     događaj ne menja i briše se samo druga predstava istog događaja. `CAPABILITY` je
     PROVEN osim B-025 za revers, gde je pitanje otvoreno. `LANDING` je PROVEN: drugih
     aktivnih grana nad istim fajlovima nema, a šema ide u svom rezu.
     `ACCEPTANCE CONTRACT` je plan po rezu.

     **Merenje je promenilo dve tvrdnje iz plana.**
     - `DOK_TIP_OM_*` je u produkciji **u celini** stari model (12 procedura).
     - Redosled `ODL-24` pre `10e` više nema razlog. `10e` ne dira `VrstaKretanja`
       (`AmbVrsteSve` nema nijednog legacy člana), a razlog „nov događaj je nevidljiv
       nepresečenom čitaocu" otpao je kad se `10c` zatvorio.

     **Predlog rezova** (svaki je svoj PR, serijski):
     1. `10e-1`: čitaoci, štampa i integritet. Uklanja potrošače `modStorno` helpera iz
        drugih modula.
     2. `10e-2`: stari storno, undo i ispravka reversa, plus zastavica `Stornirano`.
     3. `10e-3`: pisac, dvomodelni čitalac, šema i fixture.

     **Odluke operatera (12.10.2026, AMBALAZA 6.13c):** `10e` ide pre specifikacije
     `ODL-24`; REŠI KASNIJE za revers se briše (`AMB-10-ODL-28`); tri reza, tri PR-a;
     čip „print readers swallowing ledger corruption" ide posle `10e`, kao zaseban PR.

126. **`10e-1` — čitaoci, štampa i integritet starog modela obrisani** (12.10.2026).
     Grana `claude/amb-10e1-citaoci`, prvi od tri reza (`AMBALAZA.md` 6.13c). Sve obrisano
     je bilo nedostižno ili bez pozivaoca, a mereno je pre reza.

     **Obrisano.**
     - **Štampa:** rekonstrukcija starog reversa iz nogu (`ReversStampaNoge`,
       `ReversRedPripada`, `IzvAmbRedStorniran`) i grana `OM-*` u `StampajDokumentReda`.
     - **Labele:** `OM-*` u kartici i prevodu broja.
     - **Kolona „osnov":** `OsnovNaziv` i sama kolona; nijedan ekran je ne deklariše od
       `b1458258`.
     - **Integritet i saldo vozača:** `Chk_B10_ReversBezID` i `GetVozacAmbSaldo` sa
       `VozacAmbEffectiveSmer`.
     - **Bez pozivaoca:** `ReversNaziv`, `GetStorniraniRevers`, 10 pomoćnika (5 produkcionih, 5 testnih) i 11 ključeva
       poruka koji su ostali bez korisnika.

     **Merenje siročadi.** Merač nad `47e590de` i nad rezom poredi procedure bez
     pozivaoca, pa se brisanje ponavlja dok nova siročad ne nestanu. Kod u izraznoj
     poziciji je proveren grep-om, jer ga `ARNOST` ne vidi.

     **Promena ponašanja.** `StampajReversAmbalaze` sada drugi tip dokumenta **odbija po
     imenu**. Ranije bi stigao do čitaoca zaglavlja i pao kao „nema zaglavlja", što je
     istinito, ali pogrešan razlog. Ovo meri test `#225 T_AmbStampa_TudjiTipOdbijen`.
     Nova sabotaža je `amb-10e1-stampa-pusta-tudji-tip`, pa katalog ide **742 → 743**.

     **Testovi.**
     - Obrisani su BFP `Test_BKTX_ReversIDNaSvimNogama` i `Test_BKTX_ReversIDJedanDokument`.
       Merili su test-only pisca `TrackAmbalaza` i B10.
     - Obrisan je i `T_E2E_ReversIstiBrojDveStanice`.
     - `T_ReversKljucPoTipu` je postao `T_AmbTipKljuc`, a ključ tipa ostaje meren.

     **Očekivani brojevi** (izvedeni iz diff-a, proverava ih prolaz):

     | Suita | Pre | Posle | Razlog |
     |---|---|---|---|
     | `RunAllTests` | 224 | **225** | `#225` |
     | `RunBusinessFlowProSuite` | 2442 | **2412** | −18 i −12, dva obrisana testa |
     | `RunIzvestajTests` | — | **−14** | −7 ključ, −7 E2E |
     | `RunStornoTestSuite`, `RunBankaImportTestSuite` | 131, 241 | bez promene | — |

     **Mapa.** C-029, C-061, D-056 i `Chk_B10` su zatvoreni u `MAPA_SPOSOBNOSTI.md`.

     **Platforma.** Nijedan modul nije obrisan, pa posle `ImportAllVBA` nema ručnog
     `Remove`. Fixture se ne menja (u `make_fixture` je izmenjen samo komentar; potpis
     podataka je isti).

     **Verifikacija.** Jeftine kapije su zelene (`vba_check`, `--self-test`, `who_writes`
     `--check` i `--check-ownership`, `gen_schema_module --check`, `popis_citalaca --check`,
     `vba_gate --popis`). Reviewer GO je stigao 12.10.2026.

     **Prvi pun prolaz je pao na jednoj tvrdnji, i to je rupa u MOM merenju.**
     `Test_StornoCentar_All` je dao 130/1, sa porukom „B10: ReversID na redu koji nije revers
     prijavljen". Tvrdnja u `Test_StornoReversGranicaRID_Auto` merila je **tekst nalaza**
     integriteta (`TcIntegritetSa`), pa je grep po imenu `Chk_B10` nije video. Uklonjena je,
     zajedno sa pomoćnikom koji je ostao bez korisnika (`5b6b1498`). Ostatak testa pripada
     starom storno klasteru i odlazi u `10e-2`. Pouka za brisanje integritetske provere je
     da se grep radi i po ID-u nalaza (`B10`), ne samo po imenu funkcije.

     **Konačno** (izvor `a569d067a17b`, ugovor `402b3467071c`, fixture `580ff34c`):

     ```
     RunAllTests              225 / 0      (#225 OK)
     RunBusinessFlowProSuite 2412 / 0      (2442 - 18 - 12, kako je predvidjeno)
     RunBankaImportTestSuite  241 / 0      RunStornoTestSuite 131 / 0
     Test_StornoCentar_All, RunIzvestajTests  OK -- broj se u zelenom prolazu ne ispisuje,
                                             pa -1 i -14 NISU izmereni
     12 / 12 ZELENO, marker upisan

     dokaz.py amb-10e1-stampa-pusta-tudji-tip amb-int-skener-ne-gleda-knjigu
              izvestaji-amb-broj-ostaje-id --grupe
       crvenih 3 / 3, izvor pre/posle 61439cd3b0e48b3c IDENTICAN -> DOKAZANO (grupno)
     ```

     **Compile ✅** (12.10.2026 14:43, isti izvor `a569d067a17b`):
     `vba_gate --require-green --require-compile` RC=0.

     **Review `0f2e1e83`:** `CODE REVIEW GO`, `MERGE PENDING`. Nalazi su `P0 0`, `P1 0`, `P2 0`
     i `P3 2`, a static CI #941 je `SUCCESS`.
     - `P3`: broj tvrdnji u `RunIzvestajTests` nije izmeren, jer runner u zelenom prolazu
       ne ispisuje broj za tu suitu.
     - `P3`: štampa iz pravog ekrana nije proverena posle reza.
     - Grupni dokaz je prihvaćen, uz napomenu da nije isto što i pojedinačni.

     **`#416` merge-ovan BEZ ručne provere — svesna odluka operatera (12.10.2026)**, po
     obrascu iz stavke 115. Compile nad konačnim izvorom je potvrđen.
     **Otvoreno, ručno:** štampa aktivnog reversa iz liste `AMBALAZA` i ekran Integritet bez
     nalaza B10. Nalaz se upisuje ovde. Ako provera padne, ispravka ide kao zaseban PR, pre `10e-2`.

     **Ručna provera ✅ (operater, 13.10.2026)**, sveska `tests/fixtures/rucna_provera_10e1_kod.xlsm`
     (kopija fixture-a, `main` `3e0c655d`, kod uvezen kroz `run_vba --compile-only --keep`,
     `modVbaTools` ručno). Isti kod kao `10e-1`, jer `10e-2` ne dira štampu ni integritet.

     ```
     Integritet               nema nalaza B10 ni teksta "ReversID"                        OK
     AMBALAZA, 1/IZV  12/1    papir IzdavanjeAmbalaze_1-IZV_20261013_083347.pdf nastao    OK
     AMBALAZA, 1/IZV  Letv.   "Izabrani tip ambalaze nije tip poslovnog reda dokumenta"   KI-013
     ```

     Tri stvari su nađene usput, i nijedna nije iz `10e-1`:
     - **`KI-013`: višetipni ambalažni dokument štampa samo prvi tip.** `AmbDokRedMapa` drži
       **jedan** poslovni red po dokumentu (prvi u knjizi), pa red Letvarica istog dokumenta
       `StampajAmbDokument` odbija po imenu (`7506`, `modIzvestaj.bas:4347`). Odbijanje je
       fail-closed (pogrešan papir ne nastaje), ali sposobnost C-061 „dva tipa → dva papira"
       **nije** zamenjena, a MAPA ju je u `10e-1` upisala kao REPLACED. Isti put postoji od
       `10c-2`/`KI-009a`. Pisac reversa piše jedan tip po dokumentu (`UpisiReversUTx`: jedan
       `tipAmb`), pa je oblik danas dostižan samo kroz fixture (`ADK-IZV-1`).
     - **Početno stanje na ponovljenom papiru je današnje** (`-12` za `1/IZV`): papir računa
       „početno = saldo posle upisa − ovaj dokument" (`modPrint.bas:2056-2057`), pa ponovna
       štampa ne reprodukuje papir kakav je bio. Za otkupni list to rešava ODL-26 (D-057), a za
       revers pravilo nije zapisano. Pripada PR-u „print readers" posle `10e`
       (`GetAmbalazeStanje`).
     - `ERROR modMain.InitApp.EnsurePoruke | 9 | Subscript out of range` pri startu sveske.
       **Nije izmereno**; verovatno je iste klase kao `KI-011` (greška 9 iz `LogSetup` bez
       lista `SETUP_LOG`).
     - `1.dec` umesto `12/1` na papiru je već poznat `KI-012`.

     Priprema sveske je dvaput bila pogrešna: kopija fixture-a nema kod, a `run_vba` preskače
     `modVbaTools`, pa `Compile` pada na `UveziVBA`. Postupak je zapisan u memoriji.

127. **`10e-2` — stari storno, undo i ispravka reversa obrisani, sa zastavicom na knjizi** (12.10.2026).
     Grana je `claude/amb-10e2-storno`, sa `main`-a `eaf130da`. Ovo je drugi od tri reza
     (`AMBALAZA.md` 6.13c). Šta važi posle reza i zašto stoji u 6.13d i ovde se ne
     prepisuje.

     **`pre-flight`, pre koda.** Ose `DOMAIN`, `IDENTITY`, `INVARIANTS/OWNER`, `WRITERS`,
     `DOWNSTREAM`, `CAPABILITY`, `PLATFORM` i `LANDING` su PROVEN, a `EVENTS` i
     `CARDINALITY` su N/A. Pitanja za operatera nije bilo, jer je merenje zatvorilo minu 1.

     **Merenje je promenilo tri tvrdnje iz plana.**
     - **Pisaca zastavice je četiri, ne tri.** Uz `StornoAmbalazaByDokument` i
       `UnmarkAmbalazaByDokument` stoje i `StornoOMKoopByBrDok` i
       `modStornoZurnal.RestoreCell`, preko `PkColForTable(TBL_AMBALAZA)`.
     - **Mina 1:** legacy granu `UndoStorno_TX` za otkup ne doseže nijedan test, a ni
       produkcija, jer `StornoOtkup` uvek otvara žurnal operaciju. Granu za revers dosezala
       su samo dva testa starog oblika. Zato je obrisana cela legacy grana, a ostao je samo
       žurnal-put.
     - **Mina 3:** zastavicu na `tblAmbalaza` tvrdio je samo `Test_StornoJournalUndo_Auto`,
       sa dve tvrdnje. Ostali testovi undo-a tabelu samo snimaju.

     **Obrisano.**
     - **`modStorno`:** `StornoOMKoopByBrDok(_TX)`, `ReversIDRazresi`, `ReversRedoviRID`,
       `ReversIDGranica`, `ReversStanicaDan` i `ReversIDStanice` (bio je siroče još pre
       reza), `ActiveAmbalazaDokExists`, `ReversTipJe`, `ReversNogaStaniceUKljucu`,
       `StornoAmbalazaByDokument` i `JournalAmbStorno`. Pozivi zastavice su izbačeni iz
       `StornoOtkup`, `StornoOtpremnica` i `StornoPrijemnica`.
     - **`modStornoFlow`:** `RunReversCorrection`, `RunSimpleStornoRevers`, `ScanRevers` i
       `PreviewRevers`, kao i `FLOW_DOC_REVERS` grane u `BuildStornoPreview`,
       `GetChainFlags`, `CorrectionNeedsDialog` i `GetStornoChainRows`.
     - **`modStornoDok`:** revers grana u `StornoIzvrsiMod` i `IspravkaReversOpis`, sa oba
       njegova poziva (`GetNedovrseno`, `OdbaciIspravku`).
     - **`modStornoRecovery`:** legacy grana `UndoStorno_TX` sa `UnmarkStorniranoCollect`,
       `UnmarkAmbalazaByDokument`, `OtkupBlockDeadParent`, `UndoGuardReason` i `MonUndo`.
     - **`modStornoZurnal`:** revers grane u `UndoGuardReasonZaOp` i `UndoOpisOperacije`,
       `ReversIDOperacije`, `LatestOpForRevers` i `tblAmbalaza` u `PkColForTable`.
     - **`modAmbalaza.NoviReversID`:** postao je siroče kad su otišli testovi koji su ga zvali.
     - **Ključ poruke:** `STEF_REVERS`.
     - `FLOW_DOC_REVERS` i `TipUFlowDoc` **ostaju**.

     **Promena ponašanja.**
     - `UndoStorno_TX` bez žurnal operacije sada vraća `False` uz razlog „nema žurnal
       operacije". Ranije je išao u legacy granu.
     - Žurnal red nad `tblAmbalaza` je „nepodržana tabela" (fail-closed).
     - `OdbaciIspravku` i lista Nedovršeno više ne dodaju opis stanice i dana. Takav kontekst
       (`OldDocType = Revers`) je pravio samo obrisani `RunReversCorrection`.

     **Testovi.**
     - Obrisano je šest testova starog oblika u `modTestStornoCentar` (`Test_StornoRevers*`,
       `Test_StornoJournalReversGuard_Auto`, `Test_UndoReverseGuard_Auto`,
       `Test_OtkupBlockDeadParent_Auto`), sa `TcSeedRevNoga`, `TcSeedKoopRevers` i
       `TcAmbStorno`.
     - U `modTestStorno` su obrisani `T07`, `TRev_TragIspravkeNosiReversID` i
       `TRev_PreviewKliknutogReversa`, revers tvrdnja u `T10` i pomoćnici `SeedRevers`,
       `SeedAmb` i `AmbSaldo`.
     - U BFP je obrisan `Test_BKTX_ReversKoopIstiBrojDveStanice`, cela druga polovina
       `Test_BKTX_ReversPisacOdbijaZauzet` i pomoćnici `SejRevStariOblik`, `NogaPoReversID`,
       `ReversIDReda`, `AmbIDNogeStanice`, `AmbIDNoge`, `RedJeStorniran` i `StornoOznaka`.
       Prva polovina meri kanonskog pisca i ostaje. Tri tvrdnje A9 iz druge polovine su od
       10b-2 bile placebo (v. 6.13d).
     - `Test_StornoJournalUndo_Auto` više ne seje red starog oblika i ne tvrdi zastavicu.
       Meri otkup bez ambalaže, a otkup sa ambalažom meri BFP
       `Test_Amb_UndoStornaOdbijenNadKnjigom`.
     - **Nove tvrdnje kapije (mina 7):**
       - `T_AmbIspravkaReversa_EkranNudiIOtvara`: ključevi reda odluke reversa su storno i
         ISPRAVKA;
       - `T_FrameworkIspravke_SamoTriTipa`: okvir nad reversom ne izvršava nijedan mod.

       Nove sabotaže su `amb-10e2-revers-nudi-drugi-mod` i `amb-10e2-okvir-izvrsava-revers`,
       pa katalog ide **743 → 745**.

     **Očekivani brojevi**, izvedeni iz diff-a. Proverava ih prolaz.

     | Suita | Pre | Posle | Razlog |
     |---|---|---|---|
     | `RunAllTests` | 225 | **225** | nove su samo tvrdnje u postojeća dva testa |
     | `RunBusinessFlowProSuite` | 2412 | **2381** | −19 (`KoopIstiBroj`) i −12 (druga polovina `PisacOdbijaZauzet`) |
     | `RunStornoTestSuite` | 131 | **111** | −3 `T07`, −12 i −4 `TRev`, −1 `T10` |
     | `Test_StornoCentar_All` | — | **−74** tvrdnje | broj se u zelenom prolazu ne ispisuje |
     | `RunBankaImportTestSuite`, `RunIzvestajTests` | 241, — | bez promene | — |

     **Mapa.**
     - B-019 je REPLACED (10c), a B-031 REPLACED (10d).
     - B-025 za revers je INTENTIONALLY REMOVED (ODL-28).
     - B-046 i B-047, revers deo, prelaze na ODL-25.
     - Uz B-005 i B-012 je dodata napomena.
     - Red `UndoStorno_TX` je ažuriran, a „B-031 zavisnost reversa" zatvoren.
     - `WRITE_OWNERSHIP.json`: `tblAmbalaza` row_owner je samo `modAmbalaza`.

     **Dug sa imenom „nevaljani ambalažni fixture redovi u `modTestStornoCentar`" je
     zatvoren**, jer su `TcSeedRevNoga` i `colsV` otišli sa testovima koji su ih koristili.

     **Platforma.** Nijedan modul nije obrisan, pa posle `ImportAllVBA` nema ručnog `Remove`.
     Šema i fixture se ne menjaju.

     **Verifikacija.**
     - Jeftine kapije su zelene: `vba_check` (745 sabotaža, 0 nalaza), `--self-test`,
       `who_writes --check`, `--check-ownership` i `--self-test`, `gen_schema_module --check`,
       `popis_citalaca --check` i `vba_gate --popis`.
     - Merač siročadi je pušten pre reza i posle njega i daje 0 novih siročadi. Prva dva
       kruga su našla `NoviReversID`, `AmbIDNoge` i `StornoOznaka`.
     - Svako obrisano ime je proveren grep-om, pa ostaju samo istorijski komentari.
     - **Posle reviewer GO** su pušteni pun prolaz i jedan grupni dokaz nad 15 sabotaža u 7
       prolaza. To su dve nove, devet sa sidrom u presečenim procedurama i četiri nad dva
       testa kojima su dodate tvrdnje.

       ```
       python tools/dokaz.py amb-10e2- amb-otkup-storno-bez-kontrastava amb-otp-storno-bez-kontrastava amb-prj-storno-bez-kontrastava amb-storno-primitiv-vezuje otk-storno-izvora-aktivne-otpremnice otp-storno-izvora-aktivne-zbirne blok-izvor-sme-storno amb-undo-preko-kontrastava amb-undo-odbijenica-bez-puta framework-otkup framework-otpremnica-vracen amb-isprrev-ekran-ne-nudi amb-isprrev-ekran-ne-prosledjuje --grupe
       ```

     **Compile ✅** (12.10.2026 17:13, izvor `4e68278b0c8a`): operater posle `ImportAllVBA`,
     zabeleženo kroz `vba_gate --mark-compile`.

     **Konačno** (izvor `4e68278b0c8a`, ugovor `f236a8675976`, fixture `580ff34c`). Svi
     brojevi su tačno oni izvedeni iz diff-a:

     ```
     RunAllTests              225 / 0
     RunBusinessFlowProSuite 2381 / 0      (2412 - 19 - 12)
     RunStornoTestSuite       111 / 0      (131 - 20)
     RunBankaImportTestSuite  241 / 0
     12 / 12 ZELENO, marker upisan

     dokaz.py <15 sabotaza> --grupe
       crvenih 15 / 15, izvor pre/posle 22ab8d49f5582471 IDENTICAN
       grupno izmereno 12 / 15; amb-otkup-storno-bez-kontrastava i
         amb-storno-primitiv-vezuje su bile same u svojoj grupi (OK)
       PROBLEM: otk-storno-izvora-aktivne-otpremnice NE OBARA SVOJ TEST
       === NIJE DOKAZANO ===  (14 / 15 obara svoju tvrdnju)
     ```

     **Jedan član nije priznat, a to nije bila posledica reza.** `otk-storno-izvora-aktivne-otpremnice`
     je stajala u `POZNATI_NALAZI` (`tools/sabotaza.py`) kao „tvrdnja nije CEO STATIČAN
     literal". Katalog je za nju čuvao podniz „pisac odbija storno izvora", a `dokaz.py`
     traži ceo literal, pa ta sabotaža po konstrukciji nije mogla da dobije priznanje. Test
     je pao baš na dve tvrdnje koje taj podniz nose u punom obliku: „OTK izvor otpremnice
     DRAFT: pisac odbija storno izvora" i ista za IZDATO. Sabotaža je dakle grizla. U filter
     je ušla jer joj je sidro u `StornoOtkup`, a kvar je trebalo prijaviti pre prolaza.

     **Popravljeno ovde, po odluci operatera (12.10.2026).** Tvrdnja u katalogu je sada pun
     literal iz testa, a test se nije menjao, jer ga je već imao. Unos je izašao iz
     `POZNATI_NALAZI` (osam → sedam). Napomena „ovih osam traži izmenu tuđeg testa" za njega
     nije bila tačna. `vba_check` daje 745 sabotaža, 0 nalaza i 9 poznatih (bilo je 10).

     ```
     dokaz.py otk-storno-izvora-aktivne-otpremnice      (pojedinacno)
       crvenih 1 / 1, izvor pre/posle 22ab8d49f5582471 IDENTICAN -> DOKAZANO
     ```

     Svih 15 sabotaža ovog reza sada obara svoju tvrdnju: 14 iz grupnog prolaza, a ova
     pojedinačno. `src-vba` se nije menjao, pa marker zelenog i compile i dalje važe.

     **Otvoreno:**
     - ručna provera iz stavke 126 (štampa reversa iz liste `AMBALAZA`, Integritet bez B10)
       još nije javljena (✅ 13.10.2026, upisano u 126);
     - `10e-3` (pisac, dvomodelni čitalac, šema, fixture) ide posle merge-a.

128. **Review `#417` — `GO` na tačnom head-u, merge-ovano bez dodatnog commita** (12.10.2026).
     Zapis ide ovde, a ne u `#417`: reviewer je dao GO za head `23c206d8a01a`, pa bi svaki
     novi commit na tom PR-u poništio taj GO.

     - Verdikt: `P0 0`, `P1 0`, `P2 0`, `P3 1` (prenesena ručna provera iz stavke 126).
       Static CI #948 `SUCCESS`, ocena 9.4/10.
     - Merge-ovano 12.10.2026 kao `3e0c655d`.
     - `P3` je zatvoren 13.10.2026: provera je prošla, a usput su nađene tri stvari van
       `10e-1` (stavka 126, `KI-013`).
     - Reviewer je za `10e-3` najavio strožiji pregled. Kriterijumi stoje u stavci 129.

129. **`10e-3` — pisac, dvomodelni čitalac, šema i fixture starog modela obrisani** (13.10.2026).
     Grana `claude/amb-10e3-sema`, sa `main`-a `3e0c655d`. Treći, završni rez iz 6.13c. Šta
     posle reza važi i zašto stoji u `AMBALAZA.md` 6.13e i ovde se ne prepisuje.

     **`pre-flight`, pre koda.** `DOMAIN`, `IDENTITY`, `INVARIANTS/OWNER`, `WRITERS`,
     `DOWNSTREAM`, `CAPABILITY`, `PLATFORM` i `LANDING` su PROVEN; `EVENTS` i `CARDINALITY`
     su N/A. Jedno pitanje je išlo operateru (mina 1, registar storna), jer merenje kaže da je
     promena bezbedna, ali ne i gde tabela treba da stoji.

     **Merenje je promenilo tri tvrdnje iz plana.**
     - **Fixture seme je i dalje nosilo četiri reda starog oblika** (`AMB-LEG-*`). Bez
       brisanja bi posle reza svako čitanje knjige palo na fixture-u.
     - **`ENSURE_COLS` je vraćao `ReversID`.** `make_fixture` briše kolone pre dopune, pa bi
       `ReversID` iz sredine otišao i vratio se na kraj, gde ga provera kanonskog prefiksa ne
       vidi. Izlazi iz `ENSURE_COLS`.
     - **Pozicioni pisac posle reza ne postoji.** Drugi `AppendRow(TBL_AMBALAZA)` je
       `UpisiRedKnjige`: niz po `TabelaBrojKolona`, 12× `SetRowValueByColumn`. Sve seme testova
       idu po imenu (`KvarUKnjizi`, `IzvSeed`, `PostaviPoljePoPK`).

     **Odluka operatera (13.10.2026), `AMB-10-ODL-29`:** `tblAmbalaza` je **van oba spiska**
     registra storna. Storno knjige je red, ne kolona; neklasifikovana tabela je fail-closed.
     Odbačeno: `BEZ_STORNA` (pogrešno bi rekao „storno ne postoji", a pravilo `DeleteRow` ga
     čita kao „redovi se smeju brisati") i treći spisak (nov kod bez novog ponašanja).

     **Obrisano.**
     - **Šema:** šest kolona iz kanona (`Smer`, `EntitetID`, `EntitetTip`, `VozacID`,
       `Stornirano`, `ReversID`), regenerisan `modSchema` (otisak `24456ABF`).
     - **`modAmbalaza`:** `TrackAmbalaza`, `ValidateAmbalazaInput`, `RequireAmbalazaSchema`,
       `IsValidAmbSmer`, `LegacyRedProblem`, `AMB_SMER_ULAZ/IZLAZ`, stare kolone iz
       `KnjigaIndeksi`.
     - **`modConfig`:** `COL_AMB_SMER`, `_ENTITET`, `_ENTITET_TIP`, `_VOZAC`, `_REVERS_ID` i
       pet `DOK_TIP_OM_*` (nijedan produkcioni korisnik posle `10e-2`).
     - **Testovi:** sekcija C u `Test_Amb_PisacKnjige` (merila je da se red starog oblika
       preskače) i siročad BFP `AmbPolje`, `AmbKolicina`, `AmbEntitet`, `AmbVozac`.
     - **Fixture:** `AMB-LEG-*` redovi; `DROP_COLS` za šest kolona.

     **Dodato.**
     - `modSetup.ObrisiStareKoloneAmbalaze`, iz `EnsureRuntimeSchema` pre provere otiska.
     - **Promena ponašanja:** red koji nije prazan a ne dotiče knjigu je KVAR po imenu („ne
       dotice knjigu"), i u čitanju i u integritetu (`KnjigaIntegritet`).
     - `modSchemaGuard`: `TBL_AMBALAZA` izlazi iz `STORNO_TABELE`, uz komentar ODL-29.

     **Testovi.**
     - `Test_Amb_PisacKnjige` F2: „red starog oblika nije knjiga nego KVAR" i „…imenuje razlog".
     - Nov `Test_Amb_KanonskiRedPoImenu`: tabela nosi tačno kanonske kolone (i kolona dopisana na
       kraj pada), a svako polje reda kanonskog pisca je u svojoj koloni po imenu.
     - Nov `Test_Amb_StareKoloneSelfHeal`: fizički vraća stari raspored, meri da prefiks pada, pa
       da posle self-heal-a prolazi i da prvi red nosi iste vrednosti po imenu.
     - `T_StornoFilter_NedostajucaKolonaNijeTisina` (+3 tvrdnje): pin ODL-29.
     - Kontra-tvrdnje bez obrisanog tipa: BFP OTK („nijedan red otkupa ne stoji pod drugim tipom
       dokumenta"), `modTest` 2512 („revers ima ime po VRSTI", sada prema katalogu, a ne
       `Len > 0`) i 3950.

     **Sabotaže (745 → 748).**
     - Obrisana je `amb-pisac-nevalidan-legacy-prolazi` (sidro u `LegacyRedProblem`).
     - Nove su `amb-10e3-stari-red-preskocen`, `-heal-dopise-staru-kolonu`,
       `-heal-zaboravi-kolonu` i `-knjiga-u-registru-storna`.
     - Presidrene zamene (posle reza bi oborile COMPILE, ne test):
       `amb-kup-uplata-knjizi-ambalazu` (dopis praznog reda umesto `TrackAmbalaza`) i
       `amb-10c-lista-filtrira-po-nozi` (literal umesto `DOK_TIP_OM_*`).
     - Sabotaža za T1 namerno ne kvari pozicioni upis pisca: pokvarila bi svaki red knjige iz
       ranijih testova, pa bi T1 pao na čitanju pre svoje tvrdnje.

     **Očekivani brojevi**, izvedeni iz diff-a. Proverava ih prolaz.

     | Suita | Pre | Posle | Razlog |
     |---|---|---|---|
     | `RunAllTests` | 225 | **225** | nove su samo tvrdnje u postojećim testovima |
     | `RunBusinessFlowProSuite` | 2381 | **2404** | −1 (sekcija C), +18 (T1), +6 (T2) |
     | `RunStornoTestSuite`, `RunBankaImportTestSuite` | 111, 241 | bez promene | — |

     **Platforma.** Nijedan modul nije obrisan. Fixture je regenerisan: potpis
     `54298ba4` → `d3d3ff6d`, `tblAmbalaza` 22 → 18 redova, a stari je sačuvan kao
     `otkup_test_pre_10e3`. `schema_diff` nad novim fixture-om ne prijavljuje `tblAmbalaza`.
     Nad starim („zatečena sveska") prijavljuje tačno šest viška kolona i pozicioni hazard
     `tblAmbalaza`. Ostala četiri hazarda su zatečeni drift fixture-a koji `EnsureRuntimeSchema`
     leči pri svakom prolazu (38 razlika i pre i posle).

     **Verifikacija (pre GO).** Zelene su `vba_check` (748 sabotaža, 0 nalaza, 9 poznatih) i
     `--self-test`, `who_writes --check`, `--check-ownership` i `--self-test`,
     `gen_schema_module --check`, `popis_citalaca --check` i `vba_gate --popis`. Merač siročadi
     daje 0 novih. Svako obrisano ime je provereno grep-om (0 referenci van komentara). Pun
     prolaz, dokaz i compile čekaju reviewer GO.

     **Review `75eb499e` (13.10.2026):** `CODE REVIEW GO` za izvršne kapije, `MERGE PENDING`.
     Nalazi: `P0 0`, `P1 0`, `P2 0`, `P3 1` (acceptance). Static CI #950 `SUCCESS`, ocena 9.2/10.

     **Konačno** (head `75eb499e`, izvor `a199cdfa0a35`, ugovor `9bc65bff2826`, fixture
     `d3d3ff6d`). Brojevi su tačno oni izvedeni iz diff-a:

     ```
     RunAllTests              225 / 0
     RunBusinessFlowProSuite 2404 / 0      (2381 - 1 + 18 + 6)
     RunStornoTestSuite       111 / 0      RunBankaImportTestSuite 241 / 0
     12 / 12 ZELENO, marker upisan

     dokaz.py amb-10e3- amb-kup-uplata-knjizi-ambalazu amb-10c-lista-filtrira-po-nozi
              amb-pisac-polupisan-red-kao-legacy --grupe
       crvenih 7 / 7, izvor pre/posle 4b72c6a0331452a4 IDENTICAN -> DOKAZANO (grupno)
     ```

     **Compile ✅** (13.10.2026, isti izvor `a199cdfa0a35`): operater ga je pustio nad
     `rucna_provera_10e3_kod.xlsm`, a `vba_gate --require-green --require-compile` daje RC=0.

     **Ručni smoke nad konačnom šemom ✅** (operater, 13.10.2026). Sveska je
     `rucna_provera_10e3_kod.xlsm`: nov fixture (16 kolona `tblAmbalaza`), kod sa tačnog izvora.

     ```
     F7 revers 1/131026, izdato koop. 12/1 x5   papir: pocetno 18 -> novo 23             OK
     storno 1/131026-2 (ekran Storno)           "Otkazane 1"                              OK
     F1 otkup 1/131026, 10 primljeno/10 izdato   snimljen, odmah na listi                  OK
     kartica Prvi Testni, 12/1                   +5, +5/-5 (kontra-stav), +10/-10 -> 23    OK
     Integritet                                  nema bloka AMB1 (knjiga cista)            OK
     ```

     Korak „zatečena sveska operatera" (`ImportAllVBA` u svesku po starom kanonu) nije
     urađen. Put meri `Test_Amb_StareKoloneSelfHeal`, a `schema_diff` nad
     `otkup_test_pre_10e3` pokazuje tačno šest kolona za brisanje.

     **Usput je izmerena stvarna posledica `KI-011`, van `10e-3`: dupli dokument.** Revers
     `1/131026` je snimljen, ali je ljuska javila „Scr_Save → 9", pa nije pozvala
     `RefreshFromData`. Keš ekrana (`modUiData.ResetCache`) je ostao od pre upisa, lista nije
     prikazala revers, operater je snimio ponovo i nastao je `1/131026-2` sa istim sadržajem.
     „Otvori Excel → Nazad u aplikaciju" ne gradi formu iznova, pa ni to ne prazni keš. Tek
     zatvaranje i otvaranje sveske je pokazalo oba reversa. Duplikat je storniran u koraku
     „storno". `KI-011` je dopunjen; predlog je da njegova ispravka bude **prvi PR posle
     `#418`**.

     **Nedoslednost fixture-a (nije iz koda):** `ADK-IZV-X` (`3/IZV`) je storniran
     kontra-stavom u knjizi, ali njegovo zaglavlje u `make_fixture` nema `Stornirano=Da`, pa
     ga lista F7 prikazuje kao „Sačuvana". Pravi storno postavlja oboje. Ide uz sledeću
     izmenu fixture-a.

     **Otvoreno:** finalni merge review.

130. **Review `#418` — `GO` na tačnom head-u, merge-ovano bez dodatnog commita** (13.10.2026).
     Zapis ide ovde, a ne u `#418`: reviewer je dao GO za head `0f99b0e6`, pa bi svaki novi
     commit na tom PR-u poništio taj GO. Obrazac je stavka 128.

     - Verdikt: `P0 0`, `P1 0`, `P2 0`, `P3 1`. `P3` je `ImportAllVBA` nad zatečenom sveskom sa
       starim kolonama, koji nije ručno izvršen; put meri `Test_Amb_StareKoloneSelfHeal`.
       Static CI #951 `SUCCESS`, ocena 9,5/10.
     - Merge-ovano 13.10.2026 kao `6318dba3`. **`10e` je time završen**; red AMB-10 čeka još
       `ODL-24` i dopunu integritetnog audita pre S6.
     - Redosled koji je reviewer naveo u istom verdiktu: `KI-011` (P1 za operativnu upotrebu,
       svoj PR) → `KI-012`, `KI-013` i fail-open čitaoci štampe (`GetAmbalazeStanje`,
       `GetStanicaAmbSaldo`) → `ODL-24` i dopuna integritetnog audita → S6.

131. **`KI-011` — lažna greška posle uspešnog upisa: izvor, kanal ishoda i dupli dokument**
     (13.10.2026). Grana `claude/ki011-lazna-greska`, sa `main`-a `6318dba3`. P1 za operativnu
     upotrebu (review `#418`), ide sam.

     **`pre-flight`, pre koda.** `DOMAIN`, `WRITERS` (bez promene: posle commit-a `SaveRevers`
     ne piše nijednu domensku tabelu), `DOWNSTREAM`, `PLATFORM` i `LANDING` su PROVEN.
     `IDENTITY`, `CARDINALITY`, `EVENTS` i `CAPABILITY` su N/A. `ACCEPTANCE CONTRACT` je
     zatvoren tek posle odluka operatera ispod. Dve nove invarijante:
     - **`LOG-01`**: pomoćni logger ne menja `Err` pozivaoca. Vlasnik je `modSetup.LogSetup`.
     - **`UI-ISHOD-01`**: „nije snimljeno" se javlja samo kad ekran vrati poruku ili kad
       `Scr_Save` ne vrati rezultat. Vlasnik je `modUiScreens.ScrSave` + `CommitDokument`.

     **Merenje je promenilo tri tvrdnje.**
     - **Platforma (sonda u sopstvenom Excelu, 13.10.2026):**
       - Izlaz van handlera ne briše `Err` (`Exit Sub` i `End Sub`).
       - `On Error` u callee-u **gazi** grešku pozivaoca u handleru (777 → 9).
       - Neutralan callee je čuva.
       - `Err.Raise 0` diže grešku 5.
       - **Makro koji digne grešku kroz `Application.Run` je ne vraća pozivaocu**, nego otvara
         runtime dijalog.
       - Makro koji vrati rezultat a ostavi `Err` vraća i jedno i drugo; nepostojeći makro daje
         `Empty` + 1004.

       Posledica: posle `Run`-a je `Err <> 0` uz vraćen rezultat **uvek** zaostatak. `ScrSave`
       je baš to čitala kao neuspeh, a pravi pad `Scr_Save` nikad i nije mogla da vidi.
     - **Zašto otkup ne puca:** `OtkupUpisi` štampa pod `On Error Resume Next` + `Err.Clear`
       (`modOtkupUnos.bas:424-427`), a `StampajRevers` je bez toga.
     - **`ERROR modMain.InitApp.EnsurePoruke | 9`** je ista klasa: `EnsurePoruke` se završava sa
       `LogSetup "OK"`. Log operaterove sveske ga ima na **6 od 6** startova (11–13.10).

     **Popis izvora (pitanje 1).** Statički merač (`curenje_err.py`) je svestan blokova i ima
     propagaciju kroz pozive i nivo „proba postojanja" koja pada u normalnom toku.
     - `LogSetup` je koren 32 procedure u `modSetup`-u i **svih 14** izvoza štampe
       (`EnsureDocFolder`).
     - Posle ispravke `Scr_Save`, `EnsurePoruke` i izvozi više ne curi.
     - Ostatak ide u **`KI-014`** (neizmereno, svoj rez):
       - čitaoci `ScrEvent` / `ScrGridData`;
       - UI probe (`HasCtl`, `ShowToast`, `OsveziNavBrojace`);
       - `EnsurePregledDugmad`;
       - `BackfillColumn`, `SetColumnNumberFormat` i `InitSetupLog`.
     - Usput, **hipoteza, nije izmereno:** „pad PDF-a u headless Excelu" iz komentara T212 je
       verovatno bio ovo curenje.

     **Odluke operatera (13.10.2026)** — sve četiri preporuke:
     1. `KI-012` ide **posle**, u rez čitalaca štampe sa `KI-013` (upisano u `KI-012`).
     2. `LogSetup` je **neutralan prema `Err` pozivaoca**: čuva tuđu grešku i ne ostavlja svoju.
        Time se popravljaju i handleri `modSetup`-a koji posle njega rade
        `Err.Raise Err.Number` ili MsgBox sa `Err.description`. Handler-granu meri samo sonda,
        ne test.
     3. `ScrSave` **razlikuje** nevraćen rezultat (neuspeh) od zaostale greške (`ScrZaostalaGreska`
       + WARN u log, nije neuspeh). Ovo odstupa od doslovne rečenice review-a („`Err.Clear` u
       `ScrSave` nije prihvatljiv kao ispravka"). Ispravka je i dalje na izvoru, a čitalac
       samo više ne laže operatera; izvor ostaje meren posebnom tvrdnjom (226).
     4. Pad štampe posle upisa je **upozorenje uz uspeh** (postojeći kanal, oznaka `ChrW(10007)`,
        test 69). Klasifikuje se u `StampajRevers`, jedinom mestu koje zna da je upis prošao.
        `modPrint` u tom toku (`tiho`) ne otvara svoj dijalog; ponovna štampa iz liste ostaje
        kakva je.

     **Reprodukcija pre ispravke** (`CLAUDE.md` §2), izvor commit-a `5355bc0f`:

     ```
     RunAllTests: 226 ukupno, 1 palo
       FAIL T_KI011_UspesanUpisJeUspeh -- uspesan upis se ne prijavljuje kao neuspeh (KI-011)
            -- ocekivano [], dobijeno [modScrDokumenti.Scr_Save -> 9 Subscript out of range]
     sonda, isti izvor: 1 zaglavlje upisano, toast "Radnja nije uspela: ... -> 9", forma
       ostala popunjena (kol 2, partner, broj), SETUP_LOG ne postoji; drugi klik odbija
       provera broja ("Taj broj je vec izdat..."), pa je duplikat iz smoke-a nastao tek
       kad je operater dobio ili uneo NOV broj
     ```

     Prvo puštanje 226 je palo na drugom mestu: forma bez `.Show` nema punjene liste, pa
     prefill nije izabrao stanicu i potez je stao na validaciji. Test je dobio `FillCombos` i
     preduslov „prefill je izabrao otkupno mesto", pa takav promašaj više ne prolazi tiho.

     **Rez.**
     - `modSetup.LogSetup`: neutralan (stanje pozivaoca se čita pre `On Error` i vraća).
     - `modUiScreens.ScrSave`: sentinel `IsEmpty(v)`, nova `ScrZaostalaGreska`, WARN u log.
     - `modPrint`:
       - `OutputIzdavanjeAmbalaze` je Function koja vraća razlog pada, uz opcioni `tiho`;
       - `ExportIzdavanjeAmbalazePDF` sa `tiho` vraća grešku pozivaocu umesto MsgBox-a;
       - seam `StampaTestSet` (`pada` / `curi`) je tvrdo gejtovan.
     - `modNovacUnos.StampajRevers` je Function: pad vraća kao `NOVUNOS_MSG_STAMPA_NIJE` u
       `poruke` (tri pozivna mesta u `ReversUpisi`).
     - `modOtkupUI`: tri MsgBox-a u `CommitDokument` idu kroz `Dijalog`, koji u test režimu
       beleži. Novi seam-ovi `CommitDokumentTest` i `DijalogTest`.
     - `modPoruke`: `NOVUNOS_MSG_STAMPA_NIJE`.

     **Acceptance contract.**

     | Ishod | Operater vidi | Forma | Lista | Test |
     |---|---|---|---|---|
     | snimljeno | toast „Revers sačuvan: broj" | očišćena | osvežena | 226 |
     | snimljeno, papir nije napravljen | toast + dijalog „Revers je sačuvan, ali papir nije napravljen — NE unosi ga ponovo. Razlog: …" | očišćena | osvežena | 228 |
     | snimljeno uz zaostalu grešku | isto kao „snimljeno"; greška ide u log (WARN) | očišćena | osvežena | 227 |
     | nije snimljeno (pisac odbije) | toast + dijalog sa razlogom | ostaje | — | 227 |
     | nije snimljeno (`Scr_Save` se ne pokrene) | toast „Radnja nije uspela … 1004" | ostaje | — | 227 |

     - **Edge:** sveska bez `SETUP_LOG` je preduslov 226. Drugi klik posle uspeha i posle pada
       štampe ne pravi drugi dokument (226, 228).
     - **Netaknuto:** otkup (već je čistio `Err`), ponovna štampa reversa iz liste (zadržava
       dijalog), `ScrEvent`/`ScrGridData` (`KI-014`).

     **Sabotaže (748 → 754** posle review-a `#419`), dokaz se filtrira na `ki011-` (3 prolaza, `--plan`):
     - `ki011-logsetup-ostavlja-gresku` → 226 „upis ne ostavlja zaostalu gresku u Err (izvor
       KI-011: LogSetup)";
     - `ki011-scrsave-zaostala-je-neuspeh` → 227 „zaostala greska iz pomocnog koraka nije
       neuspeh upisa";
     - `ki011-scrsave-ne-javlja-neuspeh` → 227 „ekran koji ne vrati rezultat se javlja kao
       neuspeh" (pozitivna kontrola);
     - `ki011-stampa-pad-precutan` → 228 „operater vidi da papir nije napravljen".

     **Očekivani brojevi**, izvedeni iz diff-a:

     | Suita | Pre | Posle | Razlog |
     |---|---|---|---|
     | `RunAllTests` | 225 | **229** | +226…+229 (229 iz review-a `#419`); 69 dobija jedan ključ u spisku |
     | `RunBusinessFlowProSuite` | 2404 | **2404** | nijedna BFP tvrdnja se ne menja |
     | `RunStornoTestSuite`, `RunBankaImportTestSuite` | 111, 241 | bez promene | — |

     **Verifikacija pre GO.** Zelene su:
     - `vba_check` (752 sabotaže, 0 nalaza, 9 poznatih) i `--self-test`;
     - `who_writes --check`, `--check-ownership` i `--self-test`;
     - `gen_schema_module --check` i `--self-test`;
     - `popis_citalaca --check`, `vba_gate --popis` i `--self-test`, `run_vba --self-test`;
     - tri kapije self-update motora i `vba_import_marker_gate`.

     **`run_vba` posle ispravke, dokaz i ručna provera NISU pušteni: čeka se reviewer GO.**
     Testovi 227/228 i zelenilo 226 su zato neverifikovani.

     **Otvoreno.**
     - Review.
     - Posle GO: pun prolaz, `dokaz.py ki011- --grupe`, compile i ručna provera (F7 sa štampom;
       „Sačuvaj i štampaj" ide istim putem, jer `stampajUvek` niko ne čita; izazvan pad štampe;
       F1 kontrola).
     - Novi seam-ovi treba da uđu u `.claude/rules/testovi.md` §4. To ide **zasebnim process
       PR-om**.

     **Review `#419` (`2ebdd147`), krug 1: `NO-GO`, `P0 0`, `P1 0`, `P2 1`, `P3 1`, 9,0/10.**
     Static CI #953 `SUCCESS`. Reviewer ne traži nov obim (ni `KI-012`, ni `KI-014`), nego dve
     uske dopune, pa izvršne kapije.
     - **`P2` — nepoznat ekran je bio uspeh.** `ScrSave` za ključ kog nema u registru izlazi
       pre `Application.Run`: `Split("")` daje prazan niz, pa je `m = ""`, a rezultat i
       `ScrLastErr` su prazni. `CommitDokument` to čita kao upis (čisti formu, toast
       „sačuvan:" bez broja). Potvrđeno čitanjem, pa **reprodukovano pre ispravke**: 227 crven
       na „nepoznat ekran ne sme biti uspeh", sve ostale tvrdnje 227 zelene (izvor
       `c310123b`). Ispravka: izlaz bez poziva pisca postavlja `ScrLastErr`. Sabotaža
       `ki011-scrsave-nepoznat-ekran-je-uspeh`.
     - **`P3` — `LOG-01` u handleru nije bio meren testom.** Nov test **229** zove privatni
       `LogSetup` kroz `Application.Run`, u oba smera nad sveskom bez `SETUP_LOG`. Sonda nad
       STARIM kodom je pokazala da put doseže proceduru i da tvrdnje grizu (u handleru
       `4242` → `9 / VBAProject / Subscript out of range`, čist pozivalac → `9`). Sabotaža
       `ki011-logsetup-brise-gresku-pozivaoca`.
     - Isti prolaz (`c310123b`, `RunAllTests` 229/1) je i prvo izvršenje posle ispravke:
       226, 228 i 229 su zeleni.

     **Izvršne kapije posle kruga 1** (head `e840b729`):

     ```
     RunAllTests              229 / 0      (225 + 226..229)
     RunBusinessFlowProSuite 2404 / 0      RunStornoTestSuite 111 / 0
     RunBankaImportTestSuite  241 / 0      12 / 12 ZELENO, marker upisan (izvor fbe6332aca81)
     dokaz.py ki011- --grupe: crvenih 6 / 6, izvor pre/posle c100346b7742ccde IDENTICAN
       -> DOKAZANO (grupno); ki011-logsetup-ostavlja-gresku je u grupi oborio drugu
       tvrdnju (uz sabotazu citaoca), pa je ponovljen sam i oborio svoju
     compile: NEJASNO (automatski) -- rucna kapija operatera
     ```

     **Compile ✅** (operater, 13.10.2026, `rucna_provera_ki011_kod.xlsm`, kod sa `e840b729`):
     `vba_gate --mark-compile`, pa `--require-green --require-compile` daje RC=0 nad izvorom
     `fbe6332aca81`.

     **Ručna provera ✅** (operater, 13.10.2026, ista sveska, bez lista `SETUP_LOG`; snimci ekrana i
     log `tests/fixtures/Log/OtkupApp_2026-10-13.log`):

     ```
     start sveske (2x, 13:02 i 13:05)   bez "ERROR InitApp.EnsurePoruke | 9"   (pre: 6 od 6)   OK
     F7 tip 12 x10 (3 pokusaja)         "nije snimljeno": dijalog sa razlogom (manjak '12'),
                                         forma ostaje, lista bez reversa                       OK
     F7 12/1 x2 -> 1/131026             PDF ..._130808.pdf, forma ociscena, odmah na listi
                                         (pocetno 18 -> novo 20)                                OK
     pad stampe (A: zabranjen upis u
       "Revers ambalaze") -> 1/131026-2  toast "Revers sacuvan: 1/131026-2 . Revers je sacuvan,
                                         ali papir nije napravljen -- NE unosi ga ponovo.
                                         Razlog: Document not saved..." + dijalog; forma
                                         ociscena; lista 1/131026 i -2; PDF za -2 nije nastao;
                                         log: ERROR 1004 iz izvoza, bez WARN "zaostala greska"  OK
     ```

     - **F1 kontrola ✅** (operater, 13:15): otkup sa štampom je snimljen bez greške; log ima
       transakciju `tblOtkup`, `tblOtkupStavke`, `tblAmbalaza`, `tblNovac`, nastao je
       `OtkupniList_OTK-A865…pdf`, a posle toga nema nijednog `ERROR`-a.
     - Drugi klik posle pada štampe nije ručno prikazan; meri ga 228.
     - Papir i dalje nosi `1.dec` umesto `12/1`: to je `KI-012`, ide posle (odluka operatera).
     - U MsgBox-u stoji „sacuvan" bez č, a u traci poruka je ispravno. To nije iz `#419`:
       MsgBox crta kroz ANSI kodnu stranu mašine (ovde DEU, 1252), u kojoj č ne postoji.

     **Usput, nije iz `#419`: `KI-015`.** Poruka o odbijenom upisu reversa glasi „Greška pri čuvanju
     **otpremnice**… Greška pri unosu **otkupa**". Zatečeno na `main`-u: `SaveRevers` koristi
     `DOK_MSG_GRESKA_PRI_CUVANJU`, a `modNovacUnos` koristi `OTKUP_ERR_GRESKA_PRI_UNOSU` na sedam
     mesta. Ide u svoj rez; reviewer je tražio da se `#419` ne širi.

     **Preostalo:** finalni exact-head review.

132. **Review `#419` — `GO` na tačnom head-u, merge-ovano bez dodatnog commita** (13.10.2026).
     Zapis ide ovde, a ne u `#419`: reviewer je dao GO za head `9d6b22d3`, pa bi svaki novi
     commit na tom PR-u poništio taj GO. Obrazac su stavke 128 i 130.

     - Verdikt: `P0 0`, `P1 0`, `P2 0`, `P3 1`. Static CI #957 `SUCCESS`, ocena 9,6/10.
     - `P3` je prihvaćen bez izmene. Automatski „drugi klik" (226, 228) meri klik nad
       **očišćenom** formom. Istorijski scenario, u kome je operater dobio nov broj i ponovo
       uneo isti sadržaj, nije automatski meren.
     - Merge-ovano 13.10.2026 kao `e430d41f`. `KI-011` je `CLOSED`.
     - Redosled po reviewer-u: `KI-012` + `KI-013` zajedno (ovaj rez, stavka 133+), pa
       `KI-014` odvojeno. `KI-015` je zaseban mali rez.

133. **`KI-012` + `KI-013` — tip na papiru je tekst, a papir je po (dokument, tip)** (13.10.2026).
     Grana `claude/amb-stampa-ki012-ki013`, sa `main`-a `e430d41f`.

     **Odluke operatera (13.10.2026), obe po preporuci:**
     1. `KI-013` ide kao **A2, papir po tipu** (`AMB-10-ODL-30`, `AMBALAZA.md` 6.13f).
        `AMB-INV-10` ostaje, a papir je po (dokument, tip), kao legacy C-061.
     2. Fail-open čitaoci štampe (`GetAmbalazeStanje`, `GetStanicaAmbSaldo`) i istorijsko
        početno stanje na ponovljenom papiru reversa idu u **sledeći korak**. To je politika
        greške, a za nju postoji red u „Dug sa imenom".

     **`pre-flight`, pre koda.**

     | Osa | Verdikt | Dokaz |
     |---|---|---|
     | `DOMAIN` | PROVEN | `AMB-INV-10` (`AMBALAZA.md` 6.9) dozvoljava više tipova i oba smera; `ODL-30` |
     | `EVENTS` | N/A | štampa čita dokument i knjigu, ne pravi događaj |
     | `IDENTITY` | PROVEN | papir = (`AmbDokID`, `AmbTipKljuc(tip)`); oba putuju sa reda liste (`modScrIzvestaji.bas:2778`), bez lookup-a po broju. Isti (dokument, tip) sa dva reda daje `7511` |
     | `CARDINALITY` | PROVEN | dokument : poslovni red je 1:N po ugovoru, a 1:1 kod svih produkcionih pisaca (creation path ispod); papir : red je 1:1 |
     | `INVARIANTS/OWNER` | PROVEN | `AMB-INV-10` drži `PrenesiAmbalazu` (bez promene); pravila papira `7506`/`7511` drži `StampajAmbDokument`, na granici komande |
     | `WRITERS` | N/A | read-only izmena: nijedna tabela, `who_writes --check-ownership` bez promene |
     | `DOWNSTREAM` | PROVEN | čitaoci `AmbDokRedMapa` su lista F7 `REVERSI` (`modScrDokumenti.bas:3036`), prefill ispravke (`modStornoDok.bas:1144`, iza kapije 220) i štampa, koja prelazi na `AmbDokPoslovniRedovi`. Ispravka broji kroz `AmbDokBrojPoslovnihRedova`: semantika je ista, samo ide kroz isti čitalac |
     | `CAPABILITY` | PROVEN | C-061 je u legacy-ju bio papir po tipu. Tip je ušao u ključ u `71355a4b` (AUD-012, odbrana od sabiranja preko tipova), a legacy pisci su primali jedan `tipAmb` po pozivu. Operater iz liste `AMBALAZA` doseže red Letvarica (ručna provera `10e-1`) |
     | `ACCEPTANCE CONTRACT` | zatvoren posle odluka | tabela ispod |
     | `PLATFORM` | PROVEN | sonda ispod |
     | `LANDING` | PROVEN | grana sa `main`-a posle `#419`; `.claude/` se ne dira; testovi 230/231; katalog 754 → 758 |

     **Creation path višetipnog dokumenta (pitanje 2).** Nijedan produkcioni pisac ga ne pravi:
     - `UpisiReversUTx` (`modAmbalaza.bas:2216`, i za ispravku), `UpisiReversPartnera_TX`
       (`:2613`) i `NabaviAmbalazu_TX` (`:2717`) zovu `PrenesiAmbalazu` jednom, sa jednim tipom.
     - Podela u `PrenesiAmbalazu` (glavni red + ostatak) radi samo za `VRACANJE_TUDJE`. Smerovi
       reversa (`modAmbalazaUgovor.bas:732`) su `IZDATA_PRAZNA`, `POVRAT_PRAZNE` i
       `PRENOS_INTERNO`, pa je ne dosežu.
     - PWA ne piše `tblAmbalazaDokument`. Pokriće (`ULAZ_TUDJE`) i kontra-stav nisu poslovni red.
     - Oblik postoji samo u fixture-u (`ADK-IZV-1`) i u BFP testovima koji jezgro zovu direktno
       (merač `visered.py`: dvotipna nabavka u `Test_Amb_JedanProtivpartnerPoDokumentu`, test
       220 i još četiri).

     **Popis ćelija šablona (pitanje 1).** Statički merač (`popis_celija.py`, scratchpad) je
     prošao kroz `modPrint`, `modPaletniList`, `modIzvestaj`, `modFaktura`, `modUtovar` i
     `modDocStyle`: 344 dodele `.Value`, od kojih 184 izraza ne počinju labelom. Pregledano ručno:

     | Papir | Ćelija | `@` pre dodele | Nalaz |
     |---|---|---|---|
     | revers (`WriteIzdavanjeCopy`) | tip, kolona 2 | ne | **`KI-012`**, izmereno |
     | grupna prijemnica (`FillPrijemnicaSablon`) | „Tip amb.", kolona 7 | ne | **isti kvar** (`PRJ-TEST-I1` nosi `12/1`) |
     | otkupni, grupni otkupni, otpremnica (`WriteOtkupCopy`) | tip samo u tekstu „12/1 x 18" | — | nije pogođen (sonda) |
     | isti | proizvod (vrsta + sorta), klasa | ne | latentno, samo ako naziv u matičnim liči na datum ili broj |
     | kartica kooperanta, kartica ambalaže | „Broj dok.", „Opis" | ne | naši brojevi `x/ddmmyy` ostaju tekst (sonda); opis je „tip (oznaka)" |
     | faktura, utovar | broj, faktura, PO, plomba, vreme | da | već ispravljeno ranije (`modPrint.bas:2201`, `:2462-2497`) |
     | paleta, prerada | broj, kolona 4 | da | — |
     | specifikacija, sledljivost, izveštaj | ceo opseg podataka | da | — |

     Ispravljene su dve ćelije sa izmerenim oblikom. Latentne ostaju zapisane ovde, bez izmene:
     nijedna vrednost iz domena ih danas ne pogađa.

     **Platforma — sonda u sopstvenoj Excel instanci** (13.10.2026, `sonda_tekst.py`, čuvar gasi
     samo svoj PID). Excel je bio 381 (Srbija), redosled DMY, separator `.`:

     ```
     "12/1"      opsti -> DATUM 2026-12-01, ispis "1.dec"   @ pre -> "12/1"   @ posle -> 46357
     "3/2026", "1/2026", "1-2" -> datum;  "001" -> 1;  "123456789012" -> 1,23E+11
     "1/131026", "1/131026-2", "1/IZV", "12/1 x 18", "13.10.2026" -> ostaju tekst
     ```

     VBA `.Value` tumači string po en-US pravilima (mesec/dan), mimo DMY lokala; lokal menja
     samo ispis. `@` POSLE dodele je gore nego bez formata. Tvrdnja testa (`VarType = vbString`)
     zato ne zavisi od lokala. Izmeren je jedan lokal.

     **Reprodukcija pre ispravke** (`CLAUDE.md` §2, izvor `4d54883e`, fixture `d3d3ff6d`):

     ```
     RunAllTests: 231 ukupno, 2 palo
       FAIL T_KI012_TipNaPapiruJeTekst -- revers: tip ambalaze na papiru je tekst, ne datum
            (KI-012) -- ocekivano [True], dobijeno [False]
       FAIL T_KI013_VisetipniDokumentPapirPoTipu -- visetipni dokument: drugi tip dobija svoj
            papir (KI-013) -- dobijeno [... Izabrani tip ambalaze nije tip poslovnog reda dokumenta.]
     ```

     Preduslovi 230 su prošli, pa je **hipoteza testabilnosti izmerena**: list šablona posle
     PDF izvoza ostaje popunjen (oba primerka), a list prijemnice se puni kroz `Application.Run`.
     Test staje na prvoj paloj tvrdnji, pa prijemnica i dva nova odbijanja pre ispravke nisu
     ispisani. Njih meri dokaz: `ki012-prijemnica-tip-opsti-format` vraća tačno stari kod.

     **Rez.**
     - `modPrint`: format `@` PRE dodele u dve ćelije (revers, grupna prijemnica).
     - `modAmbalaza`:
       - `AmbDokPoslovniRedovi` je nov strog čitalac svih poslovnih redova dokumenta;
       - `PoslovniRedNiz` je jedan oblik niza za mapu i za papir;
       - `AmbDokBrojPoslovnihRedova` broji kroz `AmbDokPoslovniRedovi`.
     - `modIzvestaj.StampajAmbDokument` bira red po tipu; `7506` ostaje, a `7511` je nov.
     - Pisci, knjiga i fixture se ne diraju (potpis `d3d3ff6d`).

     **Acceptance contract.**

     | Slučaj | Papir | Test |
     |---|---|---|
     | jednotipni revers `2/IZV` (12/1 ×10) | tip `12/1` kao tekst, količina 10, isti papir kao pre | 230 |
     | grupna prijemnica `PRJ-TEST-I1` | tip `12/1` kao tekst | 230 |
     | višetipni `1/IZV`, red Letvarica | Letvarica ×5 (pre: `7506`) | 231 |
     | višetipni `1/IZV`, red 12/1 | 12/1 ×30, kao pre | 231 |
     | tip koji dokument ne nosi | odbijen, `7506` | 231 |
     | dva reda istog tipa (oba smera) | odbijen, `7511` (pre: tiho prvi red) | 231 |
     | bez izabranog tipa, višetipni | odbijen, `7511` (pre: tiho prvi red) | 231 |

     - **Netaknuto:** pisci i knjiga, otkupni list i otpremnica, lista F7 `REVERSI` i prefill
       ispravke (prvi red, imenovano u 6.13f), politika greške čitalaca salda (sledeći korak),
       dijalog ponovne štampe.
     - **Edge:** tvrdnja je nad tipom vrednosti, ne nad ispisom, pa ne zavisi od lokala. Tip
       koji liči na datum (`12/1`) je kanonski tip fixture-a.
     - Dokument sa dva reda istog tipa test pravi produkcionim primitivima u svojoj
       transakciji i **vraća** ga posle štampe, pa knjiga za ostale testove ostaje ista.
     - 230 i 231 štampaju u PDF (kao 226), pa `RunAllTests` otvara još tri PDF-a.

     **Sabotaže (754 → 758)**, dokaz se filtrira na `ki012-` i `ki013-`. Filter ne hvata nijedno
     ime iz `POZNATI_NALAZI`.
     - `ki012-revers-tip-opsti-format` → 230 „revers: tip ambalaze na papiru je tekst, ne datum";
     - `ki012-prijemnica-tip-opsti-format` → 230 „prijemnica: tip ambalaze na papiru je tekst, ne
       datum";
     - `ki013-stampa-prvi-red` → 231 „visetipni dokument: papir drugog tipa nosi njegov tip i
       kolicinu";
     - `ki013-dvosmislen-tip-pusta` → 231 „dva reda istog tipa se ne spajaju na jedan papir".

     **Očekivani brojevi**, izvedeni iz diff-a:

     | Suita | Pre | Posle | Razlog |
     |---|---|---|---|
     | `RunAllTests` | 229 | **231** | +230, +231 |
     | `RunBusinessFlowProSuite` | 2404 | **2404** | nijedna BFP tvrdnja se ne menja; broj poslovnih redova je isti |
     | `RunStornoTestSuite`, `RunBankaImportTestSuite` | 111, 241 | bez promene | — |

     **Verifikacija pre GO.** Zelene su:
     - `vba_check` (758 sabotaža, 0 nalaza, 9 poznatih) i `--self-test`;
     - `who_writes --check`, `--check-ownership` i `--self-test`;
     - `gen_schema_module --check` i `--self-test`;
     - `popis_citalaca --check`, `vba_gate --popis` i `--self-test`, `run_vba --self-test`;
     - tri kapije self-update motora i `vba_import_marker_gate`, sa `--self-test`.

     **Review `#420` (`26c01632`): `CODE REVIEW GO` za izvršne kapije, `P0 0`, `P1 0`, `P2 0`,
     `P3 2`, 9,3/10.** Static CI #959 `SUCCESS`. Merge još nije odobren.
     - `P3`: mreža F7 `REVERSI` prikazuje prvi red. Već je imenovano u 6.13f i ostaje za dan kad
       neki pisac počne da pravi višetipni dokument.
     - `P3`: `KNOWN_ISSUES.md` je pokazivao 381/381 reda. Uzrok je moj commit `2cba28cd`: uz
       `core.autocrlf=true` blob je sačuvan kao LF, a na `main`-u je CRLF. Vraćeno u `0d31a43f`,
       pa je diff sada tri reda.
     - Reviewer je tražio izvršnu potvrdu za `Application.Run` nad privatnim
       `FillPrijemnicaSablon`. Prolaz ispod je potvrđuje: preduslov „list prijemnice je
       popunjen" je zelen, a sabotaža prijemnice obara svoju tvrdnju.

     **Posle GO, dve izmene van produkcije.**
     - `27bd0bd4` (`tools/sabotaza.py`): `--plan` je pokazao da `ki013-dvosmislen-tip-pusta`
       gasi obe grane `7511`. Test 231 bi tada stao na ranijoj tvrdnji (bez izabranog tipa), a
       imenovana tvrdnja ne bi bila merena. Sabotaža sada gasi samo granu sa izabranim tipom.
     - `82f2ab6c` (`modTest`): prvi izvršni prolaz je dao `RunAllTests` **231/2**, i oba pada su
       bila na KOLIČINI („Letvarica x", a za 10 prazno). Tip i papir su bili tačni.
       - Sonda: `Offset(0, 3)` iz spojene ćelije B:D računa od spojenog opsega i vraća G5. G5 je
         unutar spojenog E:H, ali nije gornja leva ćelija, pa je prazna. `Cells(red, 5)` vraća 10.
       - Kvar je bio u instrumentu; produkcioni kod nije diran. Količina se sada čita
         apsolutnom kolonom.

     **Izvršne kapije** (head `82f2ab6c`, izvor `a2ce8b2f5fba`, fixture `d3d3ff6d`):

     ```
     RunAllTests              231 / 0      (229 + 230, 231)
     RunBusinessFlowProSuite 2404 / 0      RunStornoTestSuite 111 / 0
     RunBankaImportTestSuite  241 / 0      12 / 12 ZELENO, marker upisan
     dokaz.py ki012- ki013- --grupe: crvenih 4 / 4, izvor pre/posle ca3b080bbd02d6c2
       IDENTICAN -> DOKAZANO (grupno)
     compile: NEJASNO (automatski) -- rucna kapija operatera
     ```

     Brojevi su tačno oni izvedeni iz diff-a. Sveska za ručnu proveru je
     `tests/fixtures/rucna_provera_ki012_kod.xlsm` (`run_vba --compile-only --keep` nad
     `82f2ab6c`).

     **Compile ✅** (operater, 13.10.2026, `rucna_provera_ki012_kod.xlsm`, kod sa `82f2ab6c`):
     `vba_gate --mark-compile`, pa `--require-green --require-compile` daje RC=0 nad izvorom
     `a2ce8b2f5fba`.

     **Ručna provera ✅** (operater, 13.10.2026, ista sveska; snimci PDF-ova u sesiji):

     ```
     F7 revers 12/1 x2 -> 1/131026     "12/1" u tabeli (ne "1.dec"), pocetno 18 -> novo 20    OK
     AMBALAZA 1/IZV, red Letvarica     papir: Letvarica x5, UKUPNO 5 (pre: odbijenica 7506)  OK
     AMBALAZA 1/IZV, red 12/1          papir: 12/1 x30                                       OK
     AMBALAZA 2/IZV (jednotipni)       "Prijem (povrat)", 12/1 x10, kao pre                  OK
     Roba / Test kupac 1, 3/150326     prijemnica (Adobe PDF stampac): "Tip amb." = 12/1,
                                        levo poravnato (tekst)                               OK
     ```

     Usput, ništa od ovoga nije iz ovog reza (diff `modPrint`-a je samo format dve ćelije):
     - **Početno stanje na ponovljenom papiru je izvedeno iz DANAŠNJEG salda.** Na `1/IZV` za
       12/1 piše „-10 kom", jer je novo stanje danas 20, a dokument nosi 30. To je poznata
       stavka za sledeći korak (odluka operatera, red u „Dug sa imenom").
     - **„Otkupno mesto: Test Otkupr"** je odsečeno na svim papirima reversa, jer je kolona
       uska za naziv. Zatečeni raspored.
     - Papir iz F7 nosi „Kooperant: Prvi Testni · Test Otkupno Mesto" (natpis iz liste forme),
       a ponovna štampa „Kooperant: Prvi Testni". Dva puta do istog papira daju različit tekst
       partnera. Nije izmereno odakle F7 uzima naziv; odluka da li je to kvar ostaje operateru.

     **Otvoreno.**
     - Finalni exact-head review i merge.
     - Sledeći korak: fail-open čitaoci salda i početno stanje na ponovljenom papiru reversa.

134. **Review `#420` — `GO` za izvršne kapije, merge-ovano na izričit zahtev operatera** (13.10.2026).
     Zapis ide ovde, a ne u `#420`. Obrazac su stavke 128, 130 i 132.

     - Review `26c01632`: `CODE REVIEW GO` za izvršne kapije, `P0 0`, `P1 0`, `P2 0`, `P3 2`, ocena
       9,3/10, static CI #959. Oba `P3` su obrađena u PR-u:
       - mreža F7 `REVERSI` je imenovano ograničenje (`AMBALAZA.md` 6.13f);
       - `KNOWN_ISSUES.md` je vraćen na CRLF (`0d31a43f`).
     - Posle GO su na PR-u bili test-commit (`82f2ab6c`, Offset iz spojene ćelije), sabotaža
       (`27bd0bd4`) i docs. Izvršne kapije, dokaz 4/4, compile i ručna provera su nad tim izvorom
       (stavka 133).
     - Merge-ovano 13.10.2026 kao `14b9744d` (head `14a5b264`, static CI pass). `KI-012` i `KI-013`
       su `CLOSED`.
     - Sledeći korak (ovaj rez, stavke 135+): čitaoci salda na papiru. `KI-014`, `KI-015` i `ODL-24`
       su svaki svoj rez.

135. **Čitaoci salda na papiru — stanje oko svog reda knjige, kvar knjige ne daje papir**
     (13.10.2026). Grana `claude/amb-stampa-saldo`, sa `main`-a `14b9744d`. Domen:
     `AMBALAZA.md` 6.13g (`AMB-10-ODL-31`, `-32`).

     **Odluke operatera (13.10.2026), sve po preporuci** (upisane i u „Dug sa imenom"):
     1. Kad knjiga ne može da se pročita, papir **ne nastaje** (fail-closed, kao otkupni list po
        ODL-27). `GetAmbalazeStanje` i `GetStanicaAmbSaldo` gube handler.
     2. Istorijsko stanje dobijaju **revers i otpremnica**. Grupni otkupni list ide u S6.
     3. Pokriće deficita: važi **pravilo reversa** (novo = stanje posle dokumenta, početno =
        novo − kretanje). Prikaz pokrića ide uz `ODL-24`.
     4. **Kupčev revers se iz F7 ne štampa**, kao iz liste (7507).
     5. Usput iz ručne provere `#420` (odsečen naziv otkupnog mesta, natpis partnera iz F7): F7
        nosi tekst iz liste forme (`partnerTekst`), a ponovna štampa `ResolveEntitetName`. To
        nije čitalac salda, pa ide u **zaseban mali rez**.

     **`pre-flight`, pre koda.**

     | Osa | Verdikt | Dokaz |
     |---|---|---|
     | `DOMAIN` | PROVEN posle odluka | „početno" na papiru je stanje nosioca salda oko reda tog dokumenta (tabela ispod). Revers i otkupni list su se razlikovali samo kod pokrića deficita; to je bio `DOMAIN GAP` dok ga odluka 3 nije zatvorila |
     | `EVENTS` | N/A | štampa čita knjigu, ne pravi događaj |
     | `IDENTITY` | PROVEN | tačka u knjizi je prvi red dokumenta `(DokumentTIP, DokumentID)` koji dotiče nalog; granica je indeks reda, ne datum. `AmbDokID` putuje od pisca (F7: `revDokID` iz `UpisiReversAmbalaze_TX` / `IspravkaReversa_TX`) i od reda liste (`StampajAmbDokument`) do papira, bez lookup-a po broju. Višetipni `1/IZV`: papir po (dokument, tip), ODL-30 |
     | `CARDINALITY` | PROVEN | papir : (dokument, tip) je 1:1 (ODL-30); oba smera istog tipa odbija 7511 pre papira |
     | `INVARIANTS/OWNER` | PROVEN | pravilo stanja drži `modAmbalaza.AmbStanjeOkoDokumenta`, jedno za otkup, revers i otpremnicu; ODL-27 („bez granice nema stanja") ostaje u njemu |
     | `WRITERS` | N/A | read-only; `who_writes --check` i `--check-ownership` zeleni, bez promene |
     | `DOWNSTREAM` | PROVEN | popis ispod |
     | `CAPABILITY` | PROVEN | D-057 (otkup, referenca), D-058 (otpremnica i grupni: legacy je štampao „trenutno stanje"), C-061 (revers). Da li je legacy otpremnica imala isti kvar **nije mereno** (kada je legacy knjižio izlaz stanice nije pročitano) |
     | `ACCEPTANCE CONTRACT` | zatvoren posle odluka | tabela ispod |
     | `PLATFORM` | N/A | nema novog ponašanja platforme; čitanje ćelije posle izvoza je izmereno u 230/231 |
     | `LANDING` | PROVEN | grana sa `main`-a posle `#420`; `.claude/` se ne dira; testovi 232–234 i jedan BFP; katalog 758 → 765 |

     **Popis čitalaca (pitanje 1).** Proizvodni pozivaoci su samo papiri:

     | Čitalac | Pozivalac | Šta je operater video kad čitanje padne | Dostižno |
     |---|---|---|---|
     | `GetAmbalazeStanje` | `modPrint.IzdAmbSaldoVal` (revers) | papir bez tri reda stanja | zasenčeno: ponovna štampa prvo prolazi kroz strogi `AmbDokPoslovniRedovi`, a F7 čita knjigu posle strogog pisca. Tiho izostavljanje je dostizao samo **kupčev revers** (nalog traže kao `Kooperant`) |
     | `GetStanicaAmbSaldo` | `FillOtpremnicaSablon` | papir sa „x 0" i saldom −gajbe | **da**, izmereno (BFP ispod) |
     | `GetStanicaAmbSaldo` | `FillGrupniOtkupSablon` | papir sa nulom | ne: unos prijemnice je pauziran do S6 |
     | `GetAmbalazeStanje` | `modGoldenTests`, `modTest` (dva pomoćnika) | — | test |

     Izveštaji, ekrani i kartica ova dva čitaoca ne zovu: koriste stroge `AmbSaldoPoNalogu` i
     `AmbKretanjaNaloga`. `modIzvestaj` ih pominje samo u komentarima.

     **Pravilo „početnog" po papiru (pitanje 2).**

     | Papir | Nosilac | Pre ovog reza | Posle |
     |---|---|---|---|
     | otkupni list (referenca, D-057) | kooperant | stanje pre bloka po redosledu upisa; saldo = početno + izdato − primljeno | isto (omotač jezgra) |
     | revers KOOP | kooperant | novo = **današnji** saldo, početno = novo − kretanje | novo = stanje posle dokumenta, početno = novo − kretanje |
     | revers FIRMA | stanica | isto, nad stanicom | isto, nad stanicom (vozač nije prikazan, kao i pre) |
     | otpremnica | stanica | početno = današnji saldo (već posle otpremnice), saldo = početno − gajbe | saldo = stanje posle otpremnice, početno = saldo + gajbe; bez gajbi „-" |
     | grupni otkupni list | `Stanica:<VozacID>` | današnji saldo | isto; samo politika greške (S6) |

     **Reprodukcija pre ispravke** (`CLAUDE.md` §2), izvor `fa7bed98` (samo testovi), fixture
     `d3d3ff6d`:

     ```
     RunAllTests: 234 ukupno, 3 palo  (231 zatecena zelena; kontrole F7 u 232 zelene)
       FAIL T_SaldoPapir_ReversKakoJeBio -- ponovna stampa: papir reversa nosi stanje oko svog
            reda, ne danasnje -- ocekivano [|18 kom|21 kom], dobijeno [|20 kom|23 kom]
       FAIL T_SaldoPapir_KupcevReversF7BezPapira -- kupcev revers: F7 ne pravi papir i ne javlja
            kvar (kao lista, 7507) -- ocekivano [0|tiho], dobijeno [1|tiho]
       FAIL T_SaldoPapir_KvarKnjigeNijeNula -- kvar knjige: GetAmbalazeStanje dize gresku, ne
            vraca prazno -- ocekivano [True], dobijeno [False]
     RunBusinessFlowProSuite: 2414 ukupno, 4 palo  (2404 + 10 novih)
       FAIL OTP papir: pocetno stanje je stanje stanice pre otpremnice
            Expected [Test Gajba x 21193], got [Test Gajba x 21173]
       FAIL OTP papir: saldo ambalaze je stanje stanice posle otpremnice
            Expected [Test Gajba x 21173], got [Test Gajba x 21153]
       FAIL OTP papir: ponovna stampa posle kasnijeg dokumenta je ista
            Expected [..x 21173|..x 21153], got [..x 21180|..x 21160]
       FAIL OTP papir: kvar knjige -- otpremnica se ne popunjava
            Expected [], got [Test Gajba x 0|Test Gajba x -20]
     ```

     Otpremnica je pogrešna i pri prvoj štampi: „početno" 21173 je stanje POSLE nje, a saldo
     21153 oduzima njenih 20 gajbi još jednom.

     **Rez.**
     - `modAmbalaza`:
       - `AmbStanjeOkoDokumenta` je jedan čitalac stanja naloga oko dokumenta. Vraća „pre" i
         `outPosle` (pre + redovi dokumenta za nalog i tip, bez kontra-stava);
       - `GetKooperantAmbOpening` je njegov omotač, sa zatečenim ugovorom bez tipa dokumenta
         (test 223). Sidra `amb-istorija-*` su prenesena doslovno, pa te sabotaže sada pogađaju sva
         tri papira;
       - `GetAmbalazeStanje` i `GetStanicaAmbSaldo` su bez handlera.
     - `modPrint`:
       - revers: `ambDokID` u potpisu `Output`/`Export`/`Fill`; bez svog reda papir ne nastaje
         (7692); `IzdAmbSaldoVal` obrisan, a redovi stanja su uvek na papiru;
       - otpremnica čita jezgro.
     - `modNovacUnos`: `StampajRevers` dobija `revDokID`, a kupčeva grana ne štampa.
     - `modIzvestaj.StampajAmbDokument` prosleđuje `ambDokID`.

     **Acceptance contract.**

     | Slučaj | Papir | Test |
     |---|---|---|
     | F7 revers (prvi papir) | početno = stanje pre, novo = posle; **isto kao pre reza** | 232 (kontrola) |
     | drugi revers istog dana, istog partnera i tipa | nosi stanje posle prvog | 232 (kontrola) |
     | ponovna štampa prvog posle drugog | ista kao prvi papir (pre: +2 od drugog) | 232 |
     | `1/IZV` 12/1 i Letvarica (prvi redovi knjige) | `0 → 30` i `0 → 5` (pre: `−10 → 20`) | 232 |
     | kupčev revers iz F7 | upis uspe, papira nema, upozorenja nema | 233 |
     | kvar knjige, čitaoci salda | dižu grešku (pre: `Empty` / `0`) | 234 |
     | otpremnica, prva štampa | početno = stanje stanice pre, saldo = posle | BFP |
     | otpremnica, ponovna posle kasnije nabavke | ista | BFP |
     | otpremnica nad pokvarenom knjigom | ne nastaje (pre: „x 0", „x −20") | BFP |

     - **Netaknuto:** otkupni list (omotač; 222–224 i BFP `Test_Amb_PonovnaStampaBezRedaKnjige`),
       pravilo papira po tipu (230, 231), kanal ishoda F7 (226–228), pisci i knjiga, fixture
       (potpis `d3d3ff6d`).
     - **Edge:**
       - dva dokumenta istog dana su različita tačka, jer je granica red (232);
       - višetipni dokument: granica je prvi red dokumenta za nalog, a sumira se samo tip papira
         (232, Letvarica);
       - oba smera istog tipa odbija 7511 pre papira (231);
       - pokriće deficita ulazi u „novo" (odluka 3), automatski nije mereno.
     - **Imenovano, nije mereno:**
       - filter tipa dokumenta u `RedJeDokumenta`: ID-evi nose prefiks tipa, pa slučaj bez
         podmetnutog reda ne nastaje;
       - grana 7692 (revers bez svog reda): posle ODL-32 je ne doseže nijedan produkcioni put.

     **Sabotaže (758 → 765).** Dokaz se filtrira na `saldo-` (`--plan`: 5 prolaza), a filter ne
     hvata nijedno ime iz `POZNATI_NALAZI`:
     - `saldo-revers-danasnji` → 232 „ponovna stampa: papir reversa nosi stanje oko svog reda, ne
       danasnje";
     - `saldo-kupcev-revers-stampa` → 233 „kupcev revers: F7 ne pravi papir i ne javlja kvar (kao
       lista, 7507)";
     - `saldo-citac-guta-kvar` → 234 „kvar knjige: GetAmbalazeStanje dize gresku, ne vraca prazno";
     - `saldo-stanica-guta-kvar` → 234 „kvar knjige: GetStanicaAmbSaldo dize gresku, ne vraca nulu";
     - `saldo-otp-dvostruko`, `saldo-otp-danasnji`, `saldo-otp-guta-kvar` → BFP „OTP papir: …"
       (početno, ponovna štampa, kvar).

     U testu 232 tvrdnja sabotaže stoji posle kontrola F7, koje pod njom ostaju zelene: dokument je
     tada poslednji, pa je današnji saldo jednak stanju posle njega.

     **Očekivani brojevi**, izvedeni iz diff-a:

     | Suita | Pre | Posle | Razlog |
     |---|---|---|---|
     | `RunAllTests` | 231 | **234** | +232, +233, +234 |
     | `RunBusinessFlowProSuite` | 2404 | **2414** | +10 u `Test_OTP_PapirStanjeOkoSvogReda` (izmereno u reprodukciji) |
     | `RunStornoTestSuite`, `RunBankaImportTestSuite` | 111, 241 | bez promene | — |

     **Verifikacija pre GO.** Zelene su:
     - `vba_check` (765 sabotaža, 0 nalaza, 9 poznatih) i `--self-test`;
     - `who_writes --check`, `--check-ownership` i `--self-test`;
     - `gen_schema_module --check` i `--self-test`;
     - `popis_citalaca --check`, `vba_gate --popis` i `--self-test`, `run_vba --self-test`;
     - tri kapije self-update motora i `vba_import_marker_gate`, sa `--self-test`.

     **`run_vba` posle ispravke, dokaz i ručna provera NISU pušteni: čeka se reviewer GO.**
     Zelenilo 232–234 i BFP testa posle ispravke je zato neverifikovano.

     **Otvoreno.**
     - Review.
     - Posle GO:
       - pun prolaz i `dokaz.py saldo- --grupe`;
       - compile;
       - ručna provera: ponovljen `1/IZV` (12/1 i Letvarica), F7 nov revers (kontrola), štampa
         otpremnice iz liste Roba, kupčev revers iz F7 bez papira.

136. **Review `#421` (`ca437f68`), krug 1: `NO-GO`, `P0 0`, `P1 0`, `P2 2`, `P3 0`, 9,0/10**
     (13.10.2026). Static CI #963 `SUCCESS`. Reviewer ne menja koncept `AmbStanjeOkoDokumenta`,
     nego traži dva uska zatvaranja sa testom i sabotažom, pa izvršne kapije.
     - **`P2 #1` — otpremnica sa gajbama bez svog reda knjige štampa „-".** Grana
       `otpPoznato = False` nije razlikovala otpremnicu bez gajbi (ODL-27, legitimno) od izdate
       otpremnice sa gajbama kojoj je red knjige otkačen (narušen integritet).
     - **`P2 #2` — istorijski čitalac poredi tip binarno.** `AmbSaldoNaloga` poredi
       `vbTextCompare`, a izbor reda za papir ide kroz `AmbTipKljuc` (`UCase$`). Pisac traži samo
       neprazan tip (`AmbPrenosStrukturaProblem`), a F7 to proverava samo dužinom, pa je mešovit
       zapis dostižan. Knjiga i papir su tada davali dva odgovora.

     **Reprodukcija pre ispravke** (izvor `e9d21598`, samo novi testovi). Isti prolaz je i prvo
     izvršenje ispravke iz stavke 135: 232–234 i ceo `Test_OTP_PapirStanjeOkoSvogReda` su zeleni.

     ```
     RunAllTests: 235 ukupno, 1 palo
       FAIL T_SaldoPapir_TipBezObziraNaVelikaSlova -- tip bez obzira na velika slova: papir nosi
            stanje istog tipa kao saldo -- ocekivano [|3 kom|2 kom], dobijeno [|0 kom|-1 kom]
     RunBusinessFlowProSuite: 2425 ukupno, 2 palo
       FAIL OTP papir bez veze: izgubljen DokumentID -- otpremnica sa gajbama se ne popunjava
            Expected [], got [Test Gajba x -|Test Gajba x -]
       FAIL OTP papir bez veze: red tudjeg tipa dokumenta -- otpremnica sa gajbama se ne popunjava
            Expected [], got [Test Gajba x -|Test Gajba x -]
       (kontrola "OTP papir bez gajbi: stampa nepoznato stanje (ODL-27)" zelena)
     ```

     **Rez.**
     - `modPrint.FillOtpremnicaSablon`: bez granice i sa gajbama papir ne nastaje (7693). „-"
       ostaje samo otpremnici bez gajbi.
     - `modAmbalaza.AmbStanjeOkoDokumenta`: tip se u oba prolaza poredi kroz `IstiTipAmb`
       (`StrComp` `vbTextCompare`, isto kao `AmbSaldoNaloga`).

     **Testovi.**
     - BFP `Test_OTP_PapirBezVezeKnjige` kida vezu na dva načina, na redu te otpremnice, i vraća je
       pre tvrdnji:
       - `DokumentID` više ne pokazuje na otpremnicu;
       - red nosi tuđi `DokumentTIP`. Time je **izmeren i filter tipa dokumenta**, koji je u
         stavci 135 bio imenovan kao nemeren.

       Kontrola je otpremnica bez gajbi, koja i dalje štampa „-".
     - `modTest` 235: dva povrata kooperanta, `letvarica` pa `LETVARICA`, produkcionim primitivima
       u transakciji testa, koja se vraća posle štampe (obrazac 231).

     **Sabotaže (765 → 768):**
     - `saldo-otp-bez-veze-nepoznato` → BFP „OTP papir bez veze: izgubljen DokumentID -- otpremnica
       sa gajbama se ne popunjava";
     - `saldo-dokument-bez-tipa` → BFP „OTP papir bez veze: red tudjeg tipa dokumenta -- otpremnica
       sa gajbama se ne popunjava";
     - `saldo-tip-osetljiv-na-slova` → 235 „tip bez obzira na velika slova: papir nosi stanje istog
       tipa kao saldo".

     `--plan` za `saldo-` daje 10 sabotaža u 6 prolaza; u istoj grupi nema dve sabotaže koje bi
     jedna drugoj zaklonile tvrdnju.

     **Očekivani brojevi** (posle ispravke): `RunAllTests` **235/0**, BFP **2425/0** (2414 + 11),
     Storno 111 i Banka 241 bez promene.

     **Ostaje imenovano, van reza:**
     - `GetAmbalazeStanje` i dalje grupiše tip u rečniku sa binarnim poređenjem. Posle stavke 135
       ga zovu samo grupni otkupni list (S6), golden i dva test pomoćnika, pa ide uz S6 (red u
       „Dug sa imenom");
     - grana 7692 (revers bez svog reda) i dalje nema produkcioni put.

     **Verifikacija pre GO.** Zelene su jeftine kapije (ceo CI skup, `vba_check` 768 / 0 / 9
     poznatih). `run_vba` posle ispravke, dokaz, compile i ručna provera čekaju reviewer GO.

137. **Review `#421` (`80628068`), krug 2: `CODE GO` za izvršne kapije, `P0 0`, `P1 0`, `P2 0`,
     `P3 2`, 9,4/10** (14.10.2026). Static CI #964 `SUCCESS`. Merge je `NO-GO` dok se izvršne kapije
     ne završe.
     - `P3`: `GetAmbalazeStanje` i dalje grupiše tip binarno. To je već imenovano uz S6 (stavka 136),
       pa se `#421` zbog toga ne širi.
     - `P3`: prenos pravila na papire treba potvrditi izvršno, i za otkupni list kao regresionu
       kontrolu. To je ručna provera ispod.

     **Prvi izvršni prolaz** (head `80628068`, izvor `245e6e302d0c`): `run_vba` je dao **12/12 ZELENO**
     (`RunAllTests` 235/0, BFP 2425/0, Banka 241/0, Storno 111/0). Dokaz `saldo- --grupe` je dao
     **9/10**.
     - `saldo-otp-guta-kvar` → **NE OBARA NIŠTA**. Placebo je napravila MOJA ispravka P2 #1:
       otpremnicu sa gajbama sada čuvaju dve brane (strog čitalac, pa 7693 „gajbe bez reda"), pa njen
       papir ne nastaje i kad papir proguta kvar knjige — samo sa drugim razlogom. `--plan` to ne
       vidi; video ga je tek dokaz (memorija `dvoslojna-kapija-sabotaza-ne-grize`).
     - `saldo-tip-osetljiv-na-slova` u grupi nije oborio svoj test, jer ga je zaklonio
       `saldo-revers-danasnji`: kad se „novo" uzme iz kanonskog salda, istorijski čitalac se ne pita.
       Moja procena u stavci 136, da su nezavisni, bila je pogrešna. Ponovljen sam je oborio svoj
       test.

     **Ispravka, samo test i katalog** (`6460d641`; produkcioni kod nije diran). Prva brana se meri
     na slučaju koji druga pušta: otpremnica BEZ gajbi, gde je „-" dozvoljen. U
     `Test_OTP_PapirBezVezeKnjige` kvar knjige ni nad njom ne sme da da papir (+2 tvrdnje), a
     `saldo-otp-guta-kvar` je presidren na tu tvrdnju. Tvrdnja o kvaru u
     `Test_OTP_PapirStanjeOkoSvogReda` ostaje kao ishod.

     **Izvršne kapije** (head `6460d641`, izvor `cc5e27d9cca0`, ugovor `74ffdf8304dd`, fixture
     `d3d3ff6d`):

     ```
     RunAllTests              235 / 0      (231 + 232..235)
     RunBusinessFlowProSuite 2427 / 0      (2404 + 10 + 11 + 2)
     RunBankaImportTestSuite  241 / 0      RunStornoTestSuite 111 / 0
     12 / 12 ZELENO, marker upisan
     dokaz.py saldo- --grupe: crvenih 10 / 10, izvor pre/posle f8e31a078e73bc97 IDENTICAN
       -> DOKAZANO (grupno); saldo-tip-osetljiv-na-slova ponovljen sam (zaklonjen u grupi)
     compile: NEJASNO (automatski) -- rucna kapija operatera
     ```

     Sveska za ručnu proveru je `tests/fixtures/rucna_provera_saldo_kod.xlsm`
     (`run_vba --compile-only --keep` nad `6460d641`).

     > **Fixture otpremnice bez svog reda knjige se posle ovog reza ne štampaju.** Nose gajbe, a
     > red nemaju (sintetičko stanje fixture-a, `fixture-redovi-zive-u-xlsm`), pa ih P2 #1 tačno
     > odbija kao pokidanu vezu. Izuzetak je `OTP-TEST-1` (`1/TEST`): ona svoj red ima (−40), pa
     > se štampa. Prvo izdanje ove napomene je tvrdilo da se ne štampa nijedna; ispravljeno
     > merenjem knjige sveske za ručnu proveru. Pozitivna ručna provera otpremnice traži novu
     > otpremnicu: F1 → F2 → izdavanje.

     **Compile ✅** (operater, 14.10.2026, `rucna_provera_saldo_kod.xlsm`, kod sa `6460d641`):
     `vba_gate --mark-compile`, pa `--require-green --require-compile` daje RC=0 nad izvorom
     `cc5e27d9cca0`.

     **Ručna provera ✅** (operater, 14.10.2026, ista sveska; snimci PDF-ova u sesiji). Brojevi
     otpremnice su nezavisno provereni čitanjem knjige same sveske (openpyxl, kopija, red po red):

     ```
     AMBALAZA 1/IZV, 12/1               0 -> 30 kom   (pre ispravke: -10 -> 20)          OK
     AMBALAZA 1/IZV, Letvarica          0 -> 5 kom                                         OK
     F7 1/141026 (12/1 x2), prvi papir  18 -> 20      (isto kao pre ispravke, kontrola)  OK
     F7 1/141026-2 (x1)                 20 -> 21                                           OK
     1/141026 ponovo iz liste, posle -2 18 -> 20      (ne 19 -> 21)                        OK
     F7 povrat od kupca (123, pokrice 73) upisan, zaglavlje REVERS_PARTNERA, PDF-a nema    OK
     F1 otkupni list 1/141026           pocetno 21, primljeno 10, izdato 10, saldo 21     OK
     F2 otpremnica 1/141026             pocetno 39, saldo 29  (knjiga: 39 -> 29, red 25)  OK
     F7 FIRMA 1/141026-3 (prijem x1)    stanje na OM-u 29 -> 30                           OK
     otpremnica 1/141026 ponovo posle   39 -> 29, nepromenjeno                             OK
       FIRMA reversa
     ```

     - **Otpremnica bez gajbi nije ručno prikazana.** Otkupni list bez gajbi se iz F1 ne može
       uneti, a operater je procenio da sada nije važno. Taj slučaj meri BFP
       `Test_OTP_PapirBezVezeKnjige` (kontrola „-" i kvar knjige).
     - Upozorenje na kupčevom reversu posle „Yes" nije posebno ručno potvrđeno; meri ga 233.
     - Usput, nije iz `#421`: pitanje o manjku na kupčevom reversu glasi „Kooperant nema toliko
       naših gajbi… firma ga duguje kooperantu". Isti par ključeva (`AMB_ASK_DEFICIT_1/_2`) koriste
       otkup, F7 kupca i prijemnica. Dopisano uz `KI-015`.
     - Već poznato, van reza: odsečen naziv otkupnog mesta i natpis partnera iz F7 („Prvi Testni ·
       Test Otkupno Mesto") prema ponovnoj štampi („Prvi Testni").

     **Otvoreno.** Finalni exact-head review i merge.

138. **Merge `#421` — na izričit zahtev operatera, bez finalnog exact-head review-a** (14.10.2026).
     Zapis ide ovde, a ne u `#421`. Obrazac su stavke 128, 130, 132 i 134.

     - Review krug 1 (`ca437f68`): `NO-GO`, `P0 0`, `P1 0`, `P2 2`, `P3 0`. Oba `P2` su zatvorena u
       stavci 136, sa reprodukcijom pre ispravke, testom i sabotažom.
     - Review krug 2 (`80628068`): `CODE GO` za izvršne kapije, `P0 0`, `P1 0`, `P2 0`, `P3 2`,
       9,4/10. Merge je tada bio `NO-GO` dok se izvršne kapije ne završe (stavka 137).
     - **Posle tog GO je ušao test-commit `6460d641`.** Sabotaža `saldo-otp-guta-kvar` je bila
       placebo: otpremnicu sa gajbama čuvaju dve brane, pa je presidrena na otpremnicu bez gajbi.
       Produkcioni kod nije diran. Izvršne kapije (12/12), dokaz `saldo-` 10/10, compile i ručna
       provera su nad tim izvorom (`cc5e27d9cca0`).
     - Merge-ovano 14.10.2026 kao `fbc16ec6`, na **izričit zahtev operatera**. Head `2a69be53`;
       posle `6460d641` su na PR-u bili samo docs commit-i, static CI pass.
     - **Finalni exact-head review nije tražen.** Reviewer nije video ni test-commit `6460d641` ni
       docs posle njega; GO iz kruga 2 važi za `80628068`.
     - Sledeći korak (ovaj rez, stavke 139+): `KI-014`. `KI-015` i `ODL-24` su svaki svoj rez.

139. **`KI-014` — zaostala greška posle `Application.Run` u ugovoru ekrana** (14.10.2026). Grana
     `claude/ki014-zaostala-greska`, sa `main`-a `fbc16ec6`. `KI-010` je `CLOSED` (odluka operatera, ispod).

     **Merenja, pre koda** (skripte u scratchpad-u sesije `0e8f51fe`):
     - **Statički merač** (`curenje_err.py` iz stavke 131, ponovljen nad `fbc16ec6`): 5560 procedura, 739 curi
       probom, 1478 samo kroz Resume Next. U ugovoru ekrana PROBA izvor imaju samo:
       - `Scr_Event` Dokumenti i četiri matična, kroz probu u `ShowToast` (`tstScr`);
       - `Scr_Rows` Agro, BankaUvoz i Fakture, kroz `HasCtl` u `MoveBtn`.

       `Scr_Meta`, `Scr_Dozvoljen`, `Scr_ImaNesacuvano`, `Scr_Sort` i `Scr_BrojiKomade` su čisti, a
       `Scr_Build` i `Scr_Brojac` (Oporavak) curi samo kroz RN. Direktni `Run`-ovi u `modOtkupUI` (932 … 9089)
       odmah rade `Err.Clear`, pa nisu čitaoci. Ne-UI čitaoci su isti kao u stavci 131.
     - **Platformska sonda** (`sonda_ki014.py`, sopstvena instanca):

       ```
       F_Bool/F_Long/F_Str/F_Niz _Curi : IsEmpty=False  Err=9     (vrednost + zaostatak)
       F_Empty_Curi, S_Curi (Sub)      : IsEmpty=True   Err=9     (isto kao nepokrenut clan)
       NE_POSTOJI / nepostojeci modul  : IsEmpty=True   Err=1004
       pogresan broj argumenata        : IsEmpty=True   Err=449
       clan DIGNE gresku               : poziv se NE VRACA; dijalog "Run-time error '9' ...
                                         [Continue] [End] [Debug] [Help]"; posle End:
                                         gStanje=[] forme=0 (projekat resetovan)
       ```
     - **Dinamička sonda** (`sonda_ekrani.py`, prava gradnja ljuske, 14 ekrana × sve liste, prekidač liste
       kroz `ScrEvent`, neobavezni članovi „goli"): **nijedan čitalac nije dobio zaostatak**. Statički izvori
       se u normalnom toku ne dosežu. Usput je izmeren **pravi lažan toast drugog oblika**:
       ANALIZA → IZVEŠTAJI javlja „Radnja nije uspela: modScrAnaliza.Scr_Event chg:scrIzDoT -> 1004".
       `ActivateScreen` gradi zonu novog ekrana dok je `mScreen` još stari, pa `UiChange` šalje `chg:` ekranu
       ANALIZA, koji `Scr_Event` nema.

     **`pre-flight`.** `PLATFORM`, `DOWNSTREAM`, `INVARIANTS/OWNER` i `LANDING` su PROVEN. `WRITERS` je N/A
     (read-only, WARN ide u log), kao i `DOMAIN`, `IDENTITY`, `CARDINALITY`, `EVENTS` i `CAPABILITY`.
     `ACCEPTANCE CONTRACT` je zatvoren posle odluka ispod. Invarijanta `UI-ISHOD-01` se širi sa `ScrSave` na
     ceo ugovor ekrana; vlasnik je `modUiScreens.ScrRunIshod`.

     **Odluke operatera (14.10.2026), sve po preporuci:**
     1. **Obim:** ugovor ekrana (`modUiScreens` + `modUiPanel.PanelImaNesacuvano`). Ne-UI čitaoci (`InitApp`,
        Migracija, SEF recovery) su svaki svoj rez, a `modSelfUpdate` posebno (frozen).
     2. **Ugovor:** jedan pomoćnik za sve čitaoce i za `ScrSave`. `Empty`+1004 kod neobaveznog člana i kod
        `Scr_Event` znači „nije implementirano", pa time nestaje i izmeren toast ANALIZA → IZVEŠTAJI.
     3. **Izvori:** nijedan se ne dira, svi ostaju zapisani u `KI-014`. Netačni komentari se ispravljaju na
        izmereno; sam pad ekrana je nov `KI-016`.
     4. **`KI-010` → `CLOSED`.** Papir je ručno proveren u `#420` i `#421`. Odbijenice (nabavka, storniran
        revers) nisu ručno prikazane; to su kapije dokazane automatski.

     **Reprodukcija pre ispravke** (izvor `0ee25c29`: seam + testovi, čitaoci netaknuti;
     `run_vba --compile-only --keep`, pa samo 236–238 kroz `modTest.RunOne` u sopstvenoj instanci). Ispravka je
     `87c892b8`:

     ```
     TESTS=3 FAIL=3
     FAIL T_KI014_ZaostatakListeNijePad -- zaostala greska uz vracenu listu nije pad liste
          -- ocekivano [], dobijeno [-2147218582 Lista se nije ucitala]
     FAIL T_KI014_ZaostatakRadnjeNijePad -- zaostala greska uz obradjenu radnju nije pad radnje
          -- ocekivano [], dobijeno [? Radnja nije uspela: modScrMatPartneri.Scr_Event scrKI014Proba
             -> 9 Subscript out of range]
     FAIL T_KI014_ZaostatakNesacuvanogNijeOdbacivanje -- zaostala greska ne odbacuje nesacuvan unos
          ekrana -- ocekivano [True], dobijeno [False]
     ista sveska, tvrdnje iza prve (sonda_ostatak.py):
       IZVESTAJI posle ANALIZE  toast=[? Radnja nije uspela: modScrAnaliza.Scr_Event chg:scrIzDoT -> 1004 ...]
       panel: polje=cfg_SELLER_NAME pre=True, uz zaostatak=False   (fail-open)
     ```

     **Rez.**
     - `modUiScreens`:
       - `ScrRunIshod(v, poziv)` vraća `SCR_RUN_VRATIO` / `_PRAZNO` / `_NEMA` / `_PAO`; zaostatak
         upisuje u `ScrZaostalaGreska` + WARN i briše `Err`. Nema `On Error` naredbu, a `v` je `ByRef`.
       - Na njemu su `ScrPostoji`, `ScrSopstvenaBrana`, `ScrImaNesacuvano`, `ScrEvent`, `ScrGridData`,
         `ScrSort`, `ScrBrojac`, `ScrBrojiKomade` i `ScrSave`.
       - Seam `ScrCuriTestSet` / `ScrCuriAkoTest` je tvrdo gejtovan i gasi ga `ResetSeamova`.
       - Komentari „POZIV", `ScrGridData` i `ScrSopstvenaBrana` su ispravljeni na izmereno.
     - `modUiPanel.PanelImaNesacuvano` je na istom pomoćniku.
     - Kuke seam-a stoje na izlazu pravih članova: `modMaticniEkran.Redovi` / `Dogadjaj` / `ImaNesacuvano` i
       `modPodesavanja.Podesavanja_ImaNesacuvano`. Van test režima ne rade ništa i ne diraju `Err`.
     - `modOtkupUI`: samo komentari (`ScrAct`, `ReloadGrid`).

     **Acceptance contract.**

     | Čitalac | Vraćena vrednost (+ zaostatak) | `Empty` + 1004 | `Empty` + druga greška | `Empty` + 0 | Test |
     |---|---|---|---|---|---|
     | `ScrGridData` | lista važi, WARN | pad liste (`OTKUI_ERR_LISTA`) | pad liste | **nema liste**, nije pad | 236 |
     | `ScrEvent` | odgovor važi, WARN, bez toasta | **nije implementirano**: False, bez toasta | toast „Radnja nije uspela" | False | 237 |
     | `ScrImaNesacuvano` / `PanelImaNesacuvano` | odgovor važi (True ostaje True), WARN | False | False | False | 238 |
     | `ScrSopstvenaBrana` | odgovor važi, WARN | zabranjeno + `ScrLastErr` | zabranjeno + `ScrLastErr` | zabranjeno | 157 (postojeći) |
     | `ScrPostoji` | postoji, WARN | ne postoji | ne postoji | postoji | — |
     | `ScrSort` / `ScrBrojac` / `ScrBrojiKomade` | vrednost, WARN | `""` / 0 / False | isto | isto | postojeći |
     | `ScrSave` | upis važi, WARN (`UI-ISHOD-01`) | neuspeh | neuspeh | neuspeh | 227 |

     - **Netaknuto:** ponašanje `ScrSave` (227: isti tekst `ScrZaostalaGreska`, isti ishodi); `ScrBuild`, koji
       je Sub i ostaje na `Err` (sonda: Sub uz zaostatak = nepokrenut član), imenovano ograničenje; svi izvori
       (`HasCtl`, `ShowToast`, …); direktni `Run`-ovi u `modOtkupUI`.
     - **Edge:** `Scr_Rows` bez liste (ANALIZA, `Empty` uz `Err = 0`) nije pad (236, poslednja tvrdnja).
       `Empty` uz zaostatak se ne može razlikovati od nepokrenutog člana, pa važi strožiji ishod.
     - **Nemerena grana:** `ScrEvent` za `Empty` + druga greška (toast). Svi `Scr_Event` imaju isti potpis, pa
       produkcioni put do 449 ne postoji; isto mapiranje u `ScrGridData` meri 236.
     - `ScrZaostalaGreska` nosi **poslednji** zaostatak; prazni ga samo `ScrSave`, na početku upisa.

     **Sabotaže (768 → 777).** Nove:
     - `ki014-lista-zaostatak-je-pad` → 236 „zaostala greska uz vracenu listu nije pad liste";
     - `ki014-pomocnik-guta-zaostatak` → 236 „zaostala greska liste je zabelezena, ne progutana";
     - `ki014-pomocnik-nevracen-je-odgovor` → 236 „lista koju ekran ne vrati nije zaostatak nego neuspeh";
     - `ki014-lista-prazno-nije-pad` → 236 „lista koju ekran ne vrati uz gresku je pad liste";
     - `ki014-lista-bez-liste-je-pad` → 236 „ekran bez liste nije pad liste";
     - `ki014-radnja-zaostatak-je-pad` → 237 „zaostala greska uz obradjenu radnju nije pad radnje";
     - `ki014-radnja-nema-clana-je-pad` → 237 „ekran bez Scr_Event ne javlja pad tudjeg dogadjaja";
     - `ki014-nesacuvano-zaostatak-odbacuje` → 238 „zaostala greska ne odbacuje nesacuvan unos ekrana";
     - `ki014-panel-zaostatak-odbacuje` → 238 „zaostala greska ne odbacuje nesacuvanu izmenu panela".

     Kapija kataloga je tražila da pomoćnik i čitalac liste ne dele tvrdnju. Zato 236 ima tvrdnju koja ih
     razlikuje („nije zaostatak"), i ona stoji **pre** tvrdnje o padu.

     **Presidreno, isto značenje** (kod ispod sidra je zamenjen):
     - `brojac-ekran-po-imenu`, `brojac-nije-opcion` i `brojac-sentinel-umesto-nule`;
     - `mreza-podnozje-ljuska-ne-pita-ekran` i `mreza-podnozje-ugovor-fail-open`;
     - `ki011-scrsave-zaostala-je-neuspeh` i `ki011-scrsave-ne-javlja-neuspeh`, sada na pozivnom mestu u
       `ScrSave`, ne u pomoćniku.

     `--plan` za `ki014- ki011-scrsave brojac- mreza-podnozje-` daje 28 sabotaža u 7 prolaza; filter ne hvata
     nijedan od 9 `POZNATI_NALAZI`. **Očekivano zaklanjanje:** u grupi 3 `ki014-pomocnik-guta-zaostatak` obara i
     drugu tvrdnju 227 („zabelezena, ne progutana"), pa `ki011-scrsave-nepoznat-ekran-je-uspeh` (poslednja
     tvrdnja 227) u grupi ne vidi svoju i ponavlja se sam.

     **Očekivani brojevi** (izvedeni iz diff-a):

     | Suita | Pre | Posle | Razlog |
     |---|---|---|---|
     | `RunAllTests` | 235 | **238** | +236 … +238; ostali testovi ugovora (157, 227, brojač, podnožje, Palete) zadržavaju tvrdnje |
     | `RunBusinessFlowProSuite` | 2427 | **2427** | nijedna BFP tvrdnja ne prolazi kroz čitaoce ugovora |
     | `RunStornoTestSuite`, `RunBankaImportTestSuite` | 111, 241 | bez promene | — |

     **Verifikacija pre GO.** Svih 19 jeftinih kapija je zeleno:
     - `vba_check` (777 sabotaža, 0 nalaza, 9 poznatih) i `--self-test`;
     - `who_writes --check`, `--check-ownership` i `--self-test`;
     - `gen_schema_module --check` i `--self-test`;
     - `popis_citalaca --check`, `vba_gate --popis` i `--self-test`, `run_vba --self-test`;
     - paritet, self-update kapije i tvrdi popis, svaka sa `--self-test`, i `vba_import_marker_gate`.

     **`run_vba` posle ispravke, dokaz, compile i ručna provera NISU pušteni: čeka se reviewer GO.** Zelenilo
     236–238 posle ispravke je zato neverifikovano.

     **Otvoreno.**
     - Review.
     - Posle GO:
       - pun prolaz;
       - `dokaz.py ki014- ki011-scrsave brojac- mreza-podnozje- --grupe`;
       - compile;
       - ručna provera: ANALIZA → IZVEŠTAJI bez toasta, lista PREGLED, zatvaranje matičnog ekrana sa otvorenim
         editorom (pitanje mora da se pojavi), F7 upis kao kontrola `ScrSave`.
     - Seam-ovi `ScrCuriTestSet` / `ScrCuriAkoTest` treba da uđu u `.claude/rules/testovi.md` §4 — **zasebnim
       process PR-om**.

140. **Review `#422` (`d3c57cd7`): `CODE GO` za izvršne kapije, `P0 0`, `P1 0`, `P2 0`, `P3 2`, 9,2/10**
     (15.10.2026). Static CI #969 `SUCCESS`. Merge je `NO-GO` dok se ne završe suite, sabotaže, compile i ručna
     provera. Reviewer ne traži izmenu produkcionog koda. Oba `P3` su prihvaćena bez izmene koda, uz zapis:
     - **`P3 #1` — `SCR_RUN_NEMA` u `ScrEvent` ćuti i za korisnikov klik**, ne samo za sporedni `chg:`. Ekran čiji
       bi `Scr_Event` bio obrisan ili pogrešno registrovan tiho ne bi radio ništa. Danas `Scr_Event` imaju svi
       ekrani osim ANALIZE (13 od 14, popis u stavci 139), pa regresija nije izmerena. Razlikovanje opcione
       sistemske promene od obaveznog korisničkog događaja ide uz sledeću doradu ugovora (zapisano u `KI-014`).
     - **`P3 #2` — „jedan pomoćnik za ceo ugovor" je šire od stvarnog obima.** Pomoćnik pokriva sve **čitaoce
       `Err`-a** u ugovoru (devet u `modUiScreens` i `PanelImaNesacuvano`). `ScrMeta`, `ScrLayout`, `ScrGrid`,
       `ScrListe`, `ScrLista`, `ScrRadnje`, `ScrCipovi`, `ScrAktiviraj` i `ScrDeaktiviraj` i dalje zovu `Run`
       direktno, jer `Err` ne čitaju. Prvih šest ga i ne brišu, pa zaostatak ide pozivaocu. Statički merač nije
       našao nijednog pozivaoca koji ga posle toga čita. Formulacija u `KI-014` i u opisu PR-a je ispravljena.
     - Reviewer posebno traži da se nove `ki014-*` i presidrene `ki011-scrsave-*` sabotaže puste **zajedno**:
       pomoćnik je zajednički, pa ne sme da vrati lažan neuspeh upisa reversa. Filter dokaza iz stavke 139 ih
       već obuhvata.
     - `KI-016` je po reviewer-u visok prioritet: zaseban zadatak, pre produkcione stabilizacije. Predloženi
       redosled posle `#422` je `KI-016`, `KI-015`, `ODL-24`, pa S6.

     **Prvi izvršni prolaz** (head `d3c57cd7`): **11/12**. `RunAllTests` 238 / **1 palo**; BFP 2427/0, Banka 241/0,
     Storno 111/0. 236–238 su zeleni.

     ```
     FAIL T_Izv_MatricaVodiListe -- bez entiteta nema redova -- ocekivano [0], dobijeno [4]
     ```

     Tvrdnja zove `modScrIzvestaji.Scr_Rows` direktno, a taj modul rez nije dirao. Mereno kroz `modTest.RunOne`
     nad istim izvorom:

     ```
     123            OK
     237, 123       FAIL  (bez entiteta nema redova -- dobijeno [4])
     236, 123       OK
     238, 123       OK
     ```

     **Uzrok: izolacija testa, ne produkcioni kod.** Test 237 kroz pravu ljusku gradi zonu IZVEŠTAJA (ANALIZA →
     IZVEŠTAJI). Punjenje combo-a entiteta postavlja `mDefaultId` (prvi OM). `Scr_IzTestReset` je brisao
     `mComboTip`, ali ne i `mDefaultId`, pa je test 123 „bez entiteta" čitao podrazumevani entitet. Rupa je stara
     koliko i `mDefaultId`; dok nijedan test nije gradio tu zonu kroz ljusku, nije se videla.

     **Ispravka (`1d18023f`, samo test seam, gejtovan test režimom):** `Scr_IzTestReset` briše i `mDefaultId`,
     zajedno sa `mComboTip`, od koga zavisi. Posle nje `237, 123` daje OK, a `123` sam i dalje OK. Prvi prolaz
     zato ne važi za marker; pun prolaz se ponavlja nad novim izvorom.

     **Drugi izvršni prolaz** (head `1d18023f`, izvor `d10578d70bb0`): **12/12 ZELENO**, marker upisan.
     `RunAllTests` 238/0, BFP 2427/0, Banka 241/0, Storno 111/0, tačno očekivani brojevi.

     **Prvi dokaz** (`ki014- ki011-scrsave brojac- mreza-podnozje- --grupe`, potpis `45664203`): **27/28**, izvor
     pre i posle identičan.
     - `ki011-scrsave-nepoznat-ekran-je-uspeh` je u grupi 3 bio zaklonjen, kako je stavka 139 predvidela
       (`ki014-pomocnik-guta-zaostatak` obara ranije tvrdnju 227); ponovljen sam: OK.
     - **`ki014-lista-bez-liste-je-pad` → NE OBARA NIŠTA.** Poslednja tvrdnja 236 je merila ANALIZU kao
       `Empty` uz `Err = 0`, a ANALIZA vraća **niz**. To je dinamička sonda već bila izmerila
       (`rows{E=False niz=True}`, i to za svih 14 ekrana), a ja to nisam pročitao pre pisanja tvrdnje. Grana
       `SCR_RUN_PRAZNO` u `ScrGridData` nije imala slučaj, pa je tvrdnja bila placebo.

     **Ispravka (`0b1c5954`, test i seam):**
     - `ScrPraznoAkoTest` je kuka bez curenja;
     - `modMaticniEkran.Redovi` u obliku „REDOVI_BEZ_LISTE" vraća `Empty` uz `Err = 0`, kroz pravi `Run`;
     - 236 meri „bez liste" nad tim oblikom, ne nad ANALIZOM.

     Edge iz stavke 139 („ANALIZA") je time ispravljen: danas **nijedan ekran** ne vraća `Empty` u normalnom toku.
     Oblik ostaje u ugovoru jer ga zaglavlje modula zove „ekran nema listu".

     **Ispravka moje tvrdnje iz `KI-016` (prvo izdanje).** Tamo piše da `Err.Raise` sa korakom u izvoru, koji
     opisuje komentar u `ReloadGrid`, „ne postoji u kodu". To je pogrešno. Tražio sam doslovno `Scr_Rows[`, a
     izvor nosi ime funkcije liste (`"modScrFakture.RedoviSEF[" & mStep & "]"`). **30 funkcija lista**
     (`Rows*`/`Redovi*` u `modScr*`, sa `On Error GoTo EH`) grešku hvata i **ponovo diže** sa korakom, a
     `Scr_Rows` iznad njih handler nema (osim BankaNalozi i Izveštaja, koji imaju `On Error Resume Next`). Kroz
     `Application.Run` je to baš put do dijaloga iz sonde: toast sa korakom, koji je `ReloadGrid` obećavao, kroz
     ugovor ne stiže. `KI-016` je ispravljen i time pojačan; dinamički nije izmereno.

     **Izvršne kapije** (head `0b1c5954`, izvor `08757628f52a`, ugovor `9242b5a2e03f`, fixture `d3d3ff6d`):

     ```
     RunAllTests              238 / 0      (235 + 236..238)
     RunBusinessFlowProSuite 2427 / 0      RunStornoTestSuite 111 / 0
     RunBankaImportTestSuite  241 / 0      12 / 12 ZELENO, marker upisan
     dokaz.py ki014- ki011-scrsave brojac- mreza-podnozje- --grupe: crvenih 28 / 28,
       izvor pre/posle f13e981b010da820 IDENTICAN -> DOKAZANO (grupno);
       ki011-scrsave-nepoznat-ekran-je-uspeh zaklonjen u grupi 3 (predvidjeno), sam OK
     compile: NEJASNO (automatski) -- rucna kapija operatera
     ```

     Sveska za ručnu proveru je `tests/fixtures/rucna_provera_ki014_kod.xlsm`
     (`run_vba --compile-only --keep` nad `0b1c5954`).

     **Compile ✅** (operater, 15.10.2026, `rucna_provera_ki014_kod.xlsm`, kod sa `0b1c5954`):
     `vba_gate --mark-compile`, pa `--require-green --require-compile` daje RC=0 nad izvorom `08757628f52a`.
     `src-vba` je od `0b1c5954` do head-a nepromenjen; posle su samo docs.

     **Ručna provera ✅** (operater, 15.10.2026, ista sveska), svih šest stavki po očekivanju:

     ```
     1. ANALIZA -> IZVESTAJI (prvi ulazak)            bez toasta "Radnja nije uspela"    OK
     2. Maticni/Partneri "Novi" -> drugi ekran         pitanje o odbacivanju; "Ne" ostaje  OK
     3. Podesavanja, izmenjeno polje -> drugi ekran    pitanje se pojavljuje              OK
     4. prekidac lista: Fakture, Agrohemija, Uvoz      liste se ucitavaju, bez greske     OK
     5. F7 upis reversa (kontrola ScrSave)             toast "sacuvan", forma ociscena    OK
     6. start sveske (kontrola van reza)               bez laznog ERROR EnsureRuntimeSchema OK
     ```

     **Otvoreno.** Finalni exact-head review i merge.

141. **Merge `#422` — na izričit zahtev operatera, bez finalnog exact-head review-a** (15.10.2026).
     Zapis ide ovde, a ne u `#422`. Obrazac su stavke 128, 130, 132, 134 i 138.

     - Review (`d3c57cd7`): `CODE GO` za izvršne kapije, `P0 0`, `P1 0`, `P2 0`, `P3 2`, 9,2/10, static CI #969.
       Oba `P3` su prihvaćena bez izmene koda (stavka 140).
     - **Posle tog GO su ušla dva commit-a, samo test i test seam:**
       - `1d18023f`: `Scr_IzTestReset` briše i `mDefaultId`, jer je test 237 zagađivao test 123;
       - `0b1c5954`: 236 meri „bez liste" kroz seam, jer ANALIZA vraća niz.

       Posle njih su samo docs (`cf6a28d5`, `c4dd1d32`).
     - Izvršne kapije su nad `0b1c5954` (izvor `08757628f52a`): 12/12, dokaz 28/28 (grupno), compile i ručna
       provera 6/6.
     - Merge-ovano 15.10.2026 kao `fc216fc2`, na **izričit zahtev operatera**. Head `c4dd1d32`, static CI
       `SUCCESS`; `src-vba` je od `0b1c5954` do head-a nepromenjen.
     - **Finalni exact-head review nije tražen.** Reviewer nije video ni `1d18023f` ni `0b1c5954` ni docs posle
       njih; GO iz stavke 140 važi za `d3c57cd7`.
     - **`KI-014` → `CLOSED`** (odluka operatera 15.10.2026, po preporuci). Ugovor ekrana je rešen i dokazan u
       `#422`. Ne-UI čitaoci (`InitApp`, Migracija, SEF recovery) su različiti moduli i svaki je svoj rez, pa
       prelaze u nov `KI-017` sa istim popisom. `P3 #1` iz `#422` (`SCR_RUN_NEMA` ćuti i za korisnikov klik)
       ostaje zapisan i seli se u `KI-016`, kao otvorena stavka van reza.
     - Sledeći korak (stavke 142+): `KI-016`. `KI-015`, `ODL-24` i S6 dolaze posle njega.

142. **`KI-016` — pad člana ugovora ekrana ne izlazi kroz `Application.Run`** (15.10.2026). Grana
     `claude/ki016-pad-ekrana`, sa `main`-a `fc216fc2`. Skripte i izlazi su u scratchpad-u sesije `60ee0682`:
     `popis_run.py`, `sonda_ki016.py`, `repro_ki016.py`, `uredi_clanove.py`.

     **Merenja, pre koda.**
     - **Statički popis ciljeva `Run`-a** (`popis_run.py`: handler na vrhu cilja, ponovno dizanje u handleru,
       put do `Err.Raise` bez zaštite):
       - **`Scr_Rows`: 13 od 14** stiže do `Err.Raise` bez zaštite. Čista je samo ANALIZA. Ponovno dizanje
         (`Err.Raise Err.Number|errNum, ...` u handleru) ima 31 handler-a u `modScr*` i 4 u `modMaticni*`.
         **Ispravka prompta:** `On Error Resume Next` u `Scr_Rows` BankaNalozi i Izveštaja stoji **posle**
         poziva liste, pa ni oni nisu bili zaštićeni.
       - `Scr_Event`: Agro, BankaNalozi, BankaUvoz, Fakture, Izveštaji, Oporavak, Palete i Storno imaju handler
         (log i sopstveni toast); Dokumenti ima `Resume Next`; **četiri matična nemaju handler**.
       - `Scr_Build`: 13 od 13 bez handlera, ali bez puta do eksplicitnog `Err.Raise` (pada samo na kvaru koda).
       - Ostali ciljevi (Layout, Meta, Liste, Lista, Radnje, Cipovi, Sort, Brojac, ResetCache …) nemaju put do
         `Err.Raise`, a njihovi čitaoci pad ionako ne prikazuju.
       - **Paneli:** `BuildConfigEditor` i `BuildAdminPanel` hvataju sopstvenu grešku (LogErr + MsgBox) i ne
         dižu je; `_Release` / `_ImaNesacuvano` imaju `Resume Next`. **Puta do dijaloga nema.** Posledica: EH u
         `PanelOtvori` (`UIPAN_ERR_GRADNJA`) je mrtav, a pao graditelj se prijavljuje kao otvoren panel.
     - **Platformska sonda** (`sonda_ki016.py`, deo A, sopstvena instanca):

       ```
       A1  pozivalac Resume Next (COM)   -> dijalog, Continue UGASEN, End/Debug/Help; posle End: stanje
                                            prazno, forme=0
       A1g pozivalac On Error GoTo (COM) -> isto; handler pozivaoca je NE dobija
       A2  dogadjaj modeless forme,      -> isto; red posle Run-a se nije izvrsio; forma nestala,
           Excel sakriven                   Application.Visible=False, sveska otvorena
       A3  clan SAM uhvati gresku        -> Err=0; Variant/Sub=Empty, Boolean=False, String="",
                                            delimicno dodeljen Variant vraca NIZ; modul-level zapis prezivi
       A4  VBE Error Trapping            -> BreakOnAllErrors=0, BreakOnServerErrors=0 (podrazumevano)
       ```
       Zaključan projekat nije produkciona konfiguracija: nijedan pomen `Protection` u `src-vba`, `tools` i
       `docs`, a self-update uvozi module. To je pročitano, ne izmereno.
     - **Pravi put** (`sonda_ki016.py`, deo B, sveska `run_vba --compile-only --keep` nad `fc216fc2`):

       ```
       B1 kvar SEME (kolona Iznos u tblFakture), ScrGridData -> dijalog "Nedostaje kolona 'Iznos' ..."
       B2 kvar PODATKA (=NA() u ID celiji tblFakture)        -> NEMA dijaloga; citac toleriše, lista vracena
       B3 kvar seme, prava ljuska iz OnTime: IdiNaEkran FAKTURE prosao, UiEvent "lsSeg3" -> ReloadGrid ->
          dijalog; posle End forme=0, Excel sakriven
       ```

     **`pre-flight`.** `PLATFORM`, `DOWNSTREAM`, `INVARIANTS/OWNER` i `LANDING` su PROVEN. `WRITERS`, `DOMAIN`,
     `IDENTITY`, `CARDINALITY`, `EVENTS` i `CAPABILITY` su N/A: UI kanal i log, bez domenskih tabela. `Scr_Save`
     i transakcije se ne diraju (`UI-ISHOD-01`). `ACCEPTANCE CONTRACT` je zatvoren posle odluka ispod.
     **Nova invarijanta `UI-PAD-01`:** član ugovora koji ljuska zove kroz `Run` ne pušta grešku napolje, nego je
     beleži u kanal `modUiScreens`. Čitalac kanal čita **pre** vraćene vrednosti. Vlasnik je
     `modUiScreens.ScrPadZabelezi` (jedna implementacija), a član nosi samo branu i svoje ime.

     **Odluke operatera (15.10.2026), sve po preporuci:**
     1. **Obim:** lista, radnja i gradnja, ukupno 30 članova: 13 `Scr_Rows`, 4 `Scr_Event` matičnih i 13
        `Scr_Build`. Paneli ostaju van, jer puta do dijaloga nemaju; mrtav EH u `PanelOtvori` je zapisan u
        `KI-016`.
     2. **Šta operater vidi i kanal:**
        - lista: `OTKUI_MSG_MREZA_PALA`, sa imenom člana i korakom; mreža ostaje stara;
        - radnja: `OTKUI_ERR_RADNJA`, i na prekidaču liste i na izboru reda;
        - gradnja: `OTKUI_SCR_NEMA`, zona se uklanja.

        Uz to ERROR u log, bez dijaloga i bez reseta. Kanal je modul-level zapis u `modUiScreens`, ne
        vrednost (A3).
     3. **`KI-014` → `CLOSED`**, ne-UI čitaoci u nov `KI-017` (stavka 141).
     4. **`P3 #1` iz `#422`** ostaje zapisan i seli se u `KI-016`.

     **Reprodukcija pre ispravke** (izvor `2cac7cbf`: seam, kuke i testovi; brana i čitaoci netaknuti.
     `run_vba --compile-only --keep`, pa samo 239–242 kroz `modTest.RunOne`, svaki u svojoj instanci, sa
     hvatanjem dijaloga):

     ```
     239  TESTS=1 FAIL=1
          FAIL T_KI016_ClanHvataSvojPad -- svaka lista ekrana hvata svoj pad -- ocekivano [], dobijeno
          [ modScrAgro.Scr_Rows{9} modScrBankaNalozi.Scr_Rows{9} ... modScrStorno.Scr_Rows{9}]   (13 od 13)
     240  dijalog "Run-time error '9': podmetnut pad clana ugovora (KI-016)" [Continue ugasen]; posle End forme=0
     241  isto (UiEvent -> ScrAct -> ScrEvent -> Run)
     242  isto (IdiNaEkran -> ActivateScreen -> ScrBuild -> Run)
     ```
     239 staje na prvoj tvrdnji, pa radnje i gradnje u njemu nisu ispisane. 240–242 su pad kroz `Run`, pa se
     pre ispravke ne mogu izmeriti testom u suite-u, nego samo sondom.

     **Rez** (`949d1ad6`).
     - `modUiScreens`:
       - `ScrPadZabelezi(clan)` upisuje „`clan -> broj opis [izvor]`" i ERROR u log. Izvor ide u zapis kad
         nosi ime, jer funkcija liste diže sa korakom.
       - `ScrPadPreuzmi` vraća zapis i briše ga. Čitalac ga zove pre `Run`-a, da odbaci tuđ pad, i posle njega.
       - `ScrGridData`, `ScrEvent` i `ScrBuild` čitaju kanal pre vrednosti; `ScrRunIshod` je netaknut.
       - Seam `ScrPadTestSet` / `ScrPadAkoTest` je tvrdo gejtovan i gasi ga `ResetSeamova`, koji prazni i
         kanal.
     - **Brana u 30 članova:** `On Error GoTo EH` + kuka seam-a na vrhu, a handler zove `ScrPadZabelezi`
       (`Scr_Event` matičnih šalje i dogadjaj). `Scr_Build` Izveštaja se posle `Resume Next` segmenta vraća na
       branu (`On Error GoTo EH`), ne na `GoTo 0`.
     - `modOtkupUI`: prekidač liste i izbor reda idu kroz `ScrAct`, ne direktno kroz `ScrEvent`; komentari
       `ReloadGrid` i `ScrAct` su ispravljeni.

     **Acceptance contract.**

     | Pad u | Pre (izmereno) | Posle | Test |
     |---|---|---|---|
     | `Scr_Rows` (13) | dijalog, reset, Excel sakriven | `ScrLastErr` = zapis → toast „Učitavanje mreže nije uspelo na koraku: `modScrX.Scr_Rows -> … [korak]`" | 239, 240 |
     | `Scr_Event` matičnih (4) | dijalog, reset | toast „Radnja nije uspela: `modScrX.Scr_Event <dogadjaj> -> …`", i na prekidaču liste i na izboru reda | 239, 241 |
     | `Scr_Build` (13) | dijalog, reset | ekran se ne menja, toast `OTKUI_SCR_NEMA`, zona uklonjena; sledeći pokušaj gradi iznova | 239, 242 |

     - **Netaknuto:** `ScrRunIshod` i `UI-ISHOD-01`. `Scr_Save` i svih 8 `Scr_Event` koji već imaju handler.
       Zaostatak uz vraćenu listu i radnju ostaje WARN (236, 237). „Nema liste" (Empty uz `Err = 0`, bez
       zapisa) i dalje nije pad (236).
     - **Edge:** pad zabeležen mimo čitaoca, kad je član pozvan direktno, ne obara sledeće čitanje (240,
       poslednja tvrdnja). Pad uz vraćenu vrednost je pad: `Scr_Event` je Boolean, pa je posle pada „vratio
       False", a 241 tvrdi toast.
     - **Zašto nijedna sabotaža ne visi:** 239 zove članove **direktno**. Sabotaže brane gađaju Fakture
       `Scr_Rows`, MatRoba `Scr_Event` i Storno `Scr_Build`, a 240–242 kroz `Run` zovu MatPartneri i ANALIZU.
       Ostale sabotaže brane ne skidaju, nego gase kanal ili čitaoca. Višenje `dokaz.py` ionako beleži kao
       `TIMEOUT`, ne kao crveno.
     - **Imenovana ograničenja:**
       - Povratak `Scr_Build` Izveštaja na branu (umesto `GoTo 0`) kuka na vrhu ne meri, pa je bez sabotaže.
       - Pad uz delimično dodeljenu listu nije zaseban slučaj. Čitalac kanal gleda bez uslova nad vrednošću, a
         isti oblik meri `ScrEvent` (Boolean).
       - Direktni pozivi u testovima gube podizanje greške. Od 102 direktna `Scr_Rows` u `modTest`, 11 prvo
         tvrdi `IsArray`, a ostali čitaju `d(2)`, pa se crvene na `Empty`. `BuBrojRedova` pretvara `Empty` u 0,
         ali svaki njegov blok prvo tvrdi broj različit od nule (15, 1, 7, 5).

     **Sabotaže (777 → 789).** Nove:
     - `ki016-lista-pusta-pad` → 239 „svaka lista ekrana hvata svoj pad";
     - `ki016-radnja-pusta-pad` → 239 „svaka radnja maticnih ekrana hvata svoj pad";
     - `ki016-gradnja-pusta-pad` → 239 „svaka gradnja ekrana hvata svoj pad";
     - `ki016-pomocnik-ne-belezi` → 239 „pad clana je zabelezen u kanalu sa imenom clana";
     - `ki016-lista-pad-ignorisan` → 240 „pad liste je javljen kao neuspeh liste, sa imenom clana";
     - `ki016-pomocnik-gubi-izvor` → 240 „pad liste nosi izvor greske (korak)";
     - `ki016-kanal-ne-brise` → 240 „posle pada sledece citanje liste nije pad";
     - `ki016-lista-stari-pad` → 240 „pad zabelezen van citaoca ne obara sledece citanje liste";
     - `ki016-radnja-pad-ignorisan` → 241 „pad radnje je vidljiv operateru, sa imenom clana";
     - `ki016-prekidac-liste-cuti` → 241 „pad radnje na prekidacu liste je vidljiv";
     - `ki016-izbor-reda-cuti` → 241 „pad radnje na izboru reda je vidljiv";
     - `ki016-gradnja-pad-je-uspeh` → 242 „pad gradnje ne prebacuje ekran nego javlja da ekran nije dostupan".

     `--plan` za `ki016- ki014- ki011-scrsave` daje 24 sabotaže u 5 prolaza; filter ne hvata nijedan od 9
     `POZNATI_NALAZI`. `ki014-*` i `ki011-scrsave-*` idu zajedno sa novima jer su čitaoci dirani, a obrazac je
     memorija „centralna kapija zastareva dokaze". **Očekivano zaklanjanje:**
     - u grupi 4 `ki016-pomocnik-ne-belezi` gasi kanal, pa 240 i 241 padaju na prvoj tvrdnji, a
       `ki016-lista-stari-pad` i `ki016-prekidac-liste-cuti` se ponavljaju sami;
     - u grupi 3, kao u stavci 139, `ki011-scrsave-nepoznat-ekran-je-uspeh`.

     **Očekivani brojevi** (izvedeni iz diff-a):

     | Suita | Pre | Posle | Razlog |
     |---|---|---|---|
     | `RunAllTests` | 238 | **242** | +239 … +242; ostali testovi ne menjaju tvrdnje |
     | `RunBusinessFlowProSuite` | 2427 | **2427** | nijedna BFP tvrdnja ne ide kroz čitaoce ugovora; brana menja samo ishod pada |
     | `RunStornoTestSuite`, `RunBankaImportTestSuite` | 111, 241 | bez promene | — |

     **Verifikacija pre GO.** Sve jeftine kapije su zelene (20/20): `vba_check` (789 sabotaža, 0 nalaza,
     9 poznatih) i `--self-test`; `who_writes --check`, `--check-ownership` i `--self-test`;
     `gen_schema_module --check` i `--self-test`; `popis_citalaca --check`; `vba_gate --popis` i `--self-test`;
     `run_vba --self-test`; paritet, self-update kapije i tvrdi popis, svaka sa `--self-test`;
     `vba_import_marker_gate` i njegov `--self-test`; `sabotaza --self-test`.

     **`run_vba` posle ispravke, dokaz, compile i ručna provera NISU pušteni: čeka se reviewer GO.** Zelenilo
     239–242 posle ispravke je zato neverifikovano.

     **Otvoreno.**
     - Review (`#423`).
     - Posle GO:
       - pun prolaz;
       - `dokaz.py ki016- ki014- ki011-scrsave --grupe`;
       - compile;
       - ručna provera: izazvan pad liste, radnje i gradnje daje toast, ne dijalog, i aplikacija ostaje živa;
         kontrola su normalna lista, F7 upis i ANALIZA → IZVEŠTAJI.
     - Seam-ovi `ScrPadTestSet` / `ScrPadAkoTest` treba da uđu u `.claude/rules/testovi.md` §4, zajedno sa onima
       iz `#419` i `#422` — **zasebnim process PR-om**.

143. **Review `#423` (`ac921b4d`): `CODE GO`, `P0 0`, `P1 0`, `P2 0`, `P3 1`, 9,1/10** (15.10.2026). Static CI
     #974 `SUCCESS`. Merge je `NO-GO` dok se ne izvrše suite, sabotaže, compile i ručna provera; reviewer ne
     nalazi dokazan P1/P2 kvar u izvoru.
     - **`P3-1` — test ne dokazuje pad USRED gradnje.** Kuka seam-a stoji na vrhu člana, pa 242 (i 183, pre
       `Run`-a) obaraju gradnju pre nego što zona dobije ijednu kontrolu. Ne dokazuje se ni da brana važi posle
       lokalnog `Resume Next` segmenta u `Scr_Build` Izveštaja. To je bilo i imenovano ograničenje iz stavke 142.
       **Odluka operatera (15.10.2026, po preporuci): test PRE izvršnih kapija**, da se kapije puste jednom,
       nad konačnim izvorom. Razlog za „pre": pad usred gradnje je do KI-016 bio dijalog, pa je čišćenje posle
       njega tek sada dostižno u produkciji.
     - **Urađeno (`00ba4c26`, samo test i seam):**
       - druga kuka seam-a u `modScrIzvestaji.Scr_Build`, odmah iza povratka na branu posle `Resume Next`
         segmenta;
       - test 243 `T_KI016_PadUsredGradnje` prvo **direktno** meri da brana tu važi. Ako ne važi, ljuska se ne
         vozi, jer bi pad kroz `Run` otvorio dijalog. Zatim kroz ljusku meri da zona i omotači nestaju, ekran
         ostaje, a sledeći ulazak gradi **potpunu** zonu (`izHint` nastaje tek posle kuke);
       - sabotaža `ki016-izvestaji-gradnja-goto-0` → 243 „brana pokriva gradnju i posle Resume Next segmenta"
         (789 → 790). Čišćenje (tvrdnje 4 i 5) grizu i zatečene `zona-posle-pada-gradnje-ostaje` /
         `omotaci-posle-pada-gradnje-ostaju`, deklarisane na 183, pa idu u isti dokaz.
     - **Napomena reviewer-a, nije nalaz:** kanal je jedno polje (`mPadClana`). Koliziju kod ugnježdenih poziva
       reviewer nije dokazao. Kanal se preuzima pre i posle svakog `Run`-a, pa ugnježden čitalac preuzme svoj
       pad; zaseban test za ugnježden slučaj ne postoji.
     - Provera višenja pre dokaza: pod `ki016-gradnja-pad-je-uspeh` ljuska aktivira delimičnu zonu Izveštaja.
       Njen `Scr_Layout` ide kroz `RasporediPolja`, koji ima `Resume Next` na vrhu, a `Scr_Rows` je pod branom,
       pa nijedan poziv ne diže grešku kroz `Run`.
     - Jeftine kapije posle `00ba4c26`: 20/20 zeleno (790 sabotaža, 0 nalaza). Brza provera pre punog prolaza
       (`run_vba --compile-only --keep`, pa 239–243 kroz `RunOne`): `TESTS=5 FAIL=0`, bez dijaloga.

     **Očekivani brojevi:** `RunAllTests` 238 → **243**; BFP 2427, Banka 241, Storno 111 bez promene.

     **Prvi izvršni prolaz** (head `00ba4c26`): **11/12**. BFP 2427/0, Banka 241/0, Storno 111/0, ostale suite OK;
     `RunAllTests` 243 / **1 palo**:

     ```
     FAIL T_Fak_SefLogJeOpsegFakture -- pad citanja dnevnika se PRIJAVLJUJE, ne izgleda kao prazan log
          -- ocekivano [True], dobijeno [False]
     ```

     **Uzrok: premisa testa, ne produkcioni kod.** Pomoćnik `PadDnevnika` (test 162) je „prijavu" merio kao
     grešku podignutu iz `modScrFakture.Scr_Rows` pri direktnom pozivu. Po `UI-PAD-01` član grešku hvata i beleži
     u kanal, a `ScrGridData` je javlja kao neuspeh liste, ne kao praznu listu. Namera testa ostaje (AUD trag: pad
     čitanja nije prazan dnevnik); menja se samo mesto merenja. **Ovo je rupa u mom DOWNSTREAM popisu iz stavke
     142:** tražio sam testove koji bi prošli naprazno (`IsArray`, `d(2)`), a ne i testove koji MERE podizanje.
     Statički sken posle pada (procedure u test modulima koje zovu člana iz obima direktno i čitaju `Err.Number`)
     nalazi **samo** `PadDnevnika`, što se slaže sa prolazom.

     **Ispravka (`8c41e572`, samo test):** `PadDnevnika` meri kanal (`ScrPadPreuzmi`), a člana i dalje zove
     **direktno**: kroz `Run` bi `Scr_Rows` bez brane (`ki016-lista-pusta-pad`) otvorio dijalog. Tekst tvrdnji se
     ne menja. Postojeće `seflog-pad-citanja-cuti` i `seflog-nepostojeca-tabela-je-prazna` moraju i dalje da grizu
     (čitač tada vraća Empty, pa kanal ostaje prazan), pa `seflog-` ide u dokaz. Provera kroz `RunOne` nad novim
     izvorom: 162 i 239–243, `TESTS=6 FAIL=0`. Prvi prolaz zato ne važi za marker, pa se pun prolaz ponavlja.

     **Drugi izvršni prolaz** (head `8c41e572`, izvor `48932e3e721a`, ugovor `bcc264b6993f`, fixture `d3d3ff6d`):
     **12/12 ZELENO**, marker upisan.

     ```
     RunAllTests              243 / 0      (238 + 239..243)
     RunBusinessFlowProSuite 2427 / 0      RunStornoTestSuite 111 / 0
     RunBankaImportTestSuite  241 / 0      ostalih 8 suita OK
     ```

     **Dokaz** (`dokaz.py ki016- ki014- ki011-scrsave zona-posle-pada-gradnje omotaci-posle-pada-gradnje seflog-
     --grupe`; 36 sabotaža u 9 prolaza, filter ne hvata nijedan od 9 `POZNATI_NALAZI`):

     ```
     crvenih: 36 / 36    izvor pre/posle e59a78ee3c9822c9 IDENTICAN    bez ijednog TIMEOUT-a
     === DOKAZANO (grupno) ===   (grupno izmereno 30 / 36, 3 van grupa, 3 ponovljena sama)
     ```
     Sabotaže zaklonjene u grupi, pa ponovljene same (sve tri OK):
     - `ki011-scrsave-nepoznat-ekran-je-uspeh`, kao u stavci 139;
     - `ki016-lista-stari-pad`, zaklanja je `ki016-pomocnik-ne-belezi` (predviđeno u stavci 142);
     - `ki016-izbor-reda-cuti`, NIJE predviđeno: zaklonio ju je `ki016-kanal-ne-brise`. Zastareo pad u kanalu
       obara i `ScrBuild`, pa 241 pada već na preduslovu („ljuska je na ekranu MAT_PARTNERI").

     Sveska za compile i ručnu proveru je `tests/fixtures/rucna_provera_ki016_kod.xlsm`
     (`run_vba --compile-only --keep` nad `8c41e572`).

     **Compile ✅** (operater, 15.10.2026, `rucna_provera_ki016_kod.xlsm`, kod sa `8c41e572`): `vba_gate
     --mark-compile`, pa `--require-green --require-compile` daje RC=0 nad izvorom `48932e3e721a`.

     **Ručna provera ✅** (operater, 15.10.2026, ista sveska, snimci ekrana). Nijednom nije bilo dijaloga,
     aplikacija je ostala živa, a sledeći pokušaj je radio:

     ```
     1. Fakture uz preimenovanu kolonu (IznosKI016)  crveni toast "Ucitavanje mreze nije uspelo na koraku:
        modScrFakture.Scr_Rows -> -2147211703 Nedostaje kolona 'Iznos' ..."; posle vracanja imena
        lista ima 11 faktura i toast je nestao                                                      OK
     2. Partneri, seam na Scr_Event, izbor reda    "Radnja nije uspela: modScrMatPartneri.Scr_Event row:1
                                                    -> 9 podmetnut pad clana ugovora (KI-016) [...]"  OK
     3. seam #usred, klik IZVESTAJI                  ostaje Otkupni list, "Ovaj ekran jos nije prebacen..." OK
     4. seam ugasen, klik IZVESTAJI                  zona potpuna (Entitet, Od/Do, hint, Stampaj, Saldo)  OK
     ```
     Prekidač liste matičnih (stavka 2) nije na snimku; meri ga 241.

     **Dva prezentaciona nalaza ručne provere** (odluka operatera 15.10.2026: zapisati, merge bez izmene koda):
     - **(a) Korak je odsečen u toastu.** Zapis kanala je „`clan -> broj opis [izvor]`". Opis greške šeme
       nabraja zaglavlje, pa traka odseče `[modScrFakture.RedoviFakture[fakture]]`. Korak postoji u zapisu i u
       logu, ali ga operater na ekranu ne vidi, iako ga je acceptance contract iz stavke 142 obećavao.
       Kandidat za popravku: `[izvor]` ispred opisa u `ScrPadZabelezi`.
     - **(b) Poruka o padu gradnje ostaje posle uspešnog ulaska.** Crveni toast nema tajmer. `ReloadGrid` briše
       svoj (`mGridToast`), a uspešna aktivacija ekrana ne briše poruku `OTKUI_SCR_NEMA`, pa ispravan ekran
       IZVEŠTAJI i dalje „nije prebačen". Ponašanje je staro, ali je pre `KI-016` bilo praktično nedostižno,
       jer je pad gradnje gasio aplikaciju. Dira životni vek toasta u `ActivateScreen`, pa traži svoj test.

     Oba su zapisana u `KI-016` kao otvorena stavka van reza i idu u mali zaseban rez posle merge-a.

     **Otvoreno.** Merge na izričit zahtev, posle zelenog CI na tačnom head-u.

144. **Merge `#423` — na izričit zahtev operatera, na tačnom head-u posle zelenog static CI** (15.10.2026).
     Zapis ide ovde, a ne u `#423`. Obrazac su stavke 128, 130, 132, 134, 138 i 141.

     - Review (`ac921b4d`): `CODE GO`, `P0 0`, `P1 0`, `P2 0`, `P3 1`, 9,1/10, static CI #974 (stavka 143). `P3-1` je
       zatvoren testom 243 pre izvršnih kapija (odluka operatera).
     - **Posle tog GO su ušli samo test i seam:** `00ba4c26` (test 243 i druga kuka seam-a u
       `modScrIzvestaji.Scr_Build`) i `8c41e572` (`PadDnevnika` meri kanal pada). Posle njih su samo docs
       (`f769ef80`, `1252f79e`). Produkcioni kod van seam kuke nije diran.
     - Izvršne kapije su nad `8c41e572` (izvor `48932e3e721a`): 12/12, dokaz 36/36 (grupno), compile i ručna
       provera 4/4.
     - Merge-ovano 15.10.2026 kao `66f7897f`, na **izričit zahtev operatera**, sa `--match-head-commit` na head
       `1252f79e` (static CI `success` na tom head-u, run 38717568835).
     - **Finalni exact-head review nije tražen.** Reviewer nije video `00ba4c26`, `8c41e572` ni docs posle njih;
       GO iz stavke 143 važi za `ac921b4d`.
     - `KI-016` ostaje `OPEN` zbog stavki van reza (a)–(e) iz `KNOWN_ISSUES`; (d) i (e) su ovaj rez (stavke 145+).

145. **`KI-016` (d) i (e) — dva prezentaciona nalaza ručne provere `#423`** (15.10.2026). Grana
     `claude/ki016-poruke-pada`, sa `main`-a `66f7897f`. Skripte i izlazi su u scratchpad-u sesije `23d2fe62`:
     `repro_ki016.py` i `sonda_ki016.py` (preuzeti iz `60ee0682`), `fix.py`, `sab.py`, `sabB.py`, `kapije.sh`.

     **Popis pre koda.**
     - **(d) Ko čita oblik zapisa kanala.** Nijedan produkcioni čitalac ga ne rastavlja: `ScrAct`,
       `LoadGridFromScreen` (izvor greške → `ReloadGrid`) i `CommitDokument` ga prenose ceo. Nijedan
       `Split`/`InStr` po `" -> "` u `src-vba` ne čita kanal. Tačan oblik mere samo testovi:
       - 240 je tražio `"modScrMatPartneri.Scr_Rows -> "`, a 243 `Left$(…, 28) = "modScrIzvestaji.Scr_Build ->"`.
         Oba posle ispravke mere ime člana kao **početak** zapisa i `" -> "` posebno;
       - 239 (`KI016Clan`: početak + `" -> "`), 241 (`KI016ToastPada`: sadrži člana i događaj) i 162
         (`PadDnevnika`: dužina) ne zavise od redosleda.

       Sidro `ki016-pomocnik-gubi-izvor` je red `If InStr(errSrc, ".")`, koji ostaje. `ki016-pomocnik-ne-belezi`
       je gađao red koji se menja, pa dobija novo sidro.
     - **(e) Izlazi `ActivateScreen` sa porukom:** `OTKUI_SCR_ZABRANJEN` (crven), `OTKUI_SCR_NEMA` kad modula
       nema (zelen, tajmer od 4 s), `OTKUI_SCR_NEMA` posle pada gradnje (crven) i greška panela (crven).
       - Prigušena stavka sidebara je i dalje klikabilna (`PaintNav` samo boji), pa je zabranjen prelazak
         dostižan jednim klikom.
       - Greška panela je praktično nedostižna: `UIPAN_ERR_NEPOZNAT` seče `PanelPostoji`, `SAMO_ADMIN` seče
         `ScrDozvoljen` pre nje, a `UIPAN_ERR_GRADNJA` je mrtav EH (`KI-016` (b)).
       - `PrimeniNovaPrava` posle zamene operatera **namerno** ostavlja `OTKUI_SCR_ZABRANJEN` pa zove
         `NaPrviDozvoljenEkran` → uspešan `ActivateScreen`. Zastavica koja preživi tuđu poruku sakrila bi to
         objašnjenje, pa je `ShowToast` spušta.

     **Odluka operatera (15.10.2026), po preporuci: varijanta B.** Uspešan prelazak gasi crvenu poruku
     odbijenog prelaska iz `ActivateScreen`-a, i to i `OTKUI_SCR_NEMA` posle pada gradnje i
     `OTKUI_SCR_ZABRANJEN`. Gasi je **svaki** uspešan prelazak, ne samo ulazak u isti ekran, kao `mGridToast`.
     - Greška panela ostaje van (varijanta C), zapisana kao nedostižna. Zeleni `OTKUI_SCR_NEMA` ima tajmer.
     - Zastavicu spušta i `ShowToast`, kad nova poruka zameni staru.
     - Ista latentna rupa u `mGridToast` se samo zapisuje (ispod), ne dira se.

     **Reprodukcija pre ispravke** (`run_vba --compile-only --keep`, pa `RunOne` kroz `repro_ki016.py`, sopstvena
     instanca, bez dijaloga):

     ```
     e500390b (novi testovi, stara produkcija)            TESTS=4 FAIL=2
       FAIL 240 -- korak stoji ispred opisa greske, pa ga dug opis ne odseca u toastu
       FAIL 242 -- uspesan ulazak posle pada gradnje gasi poruku o padu
                   -- ocekivano [], dobijeno [Ovaj ekran jos nije prebacen na novi izgled.]
       OK   241, 243
     d9e0a257 (test zabrane, ispravka A bez B)             TESTS=1 FAIL=1
       FAIL 242 -- uspesan prelazak gasi poruku o zabranjenom prelasku
                   -- ocekivano [ANALIZA|], dobijeno [ANALIZA|Nemate dozvolu za ovu oblast.]
     ```
     Drugi prolaz meri i obe tvrdnje pre nove: posle ispravke A poruka o padu nestaje, a tuđa poruka ostaje.
     Prvi prolaz meri i instrument: `ToastNaEkranu` je pročitao vidljivu traku, pa `.Visible` nad formom bez
     `.Show` radi.

     **Ispravka (`952fc745`, `b24053ac`).**
     - **(d)** `ScrPadZabelezi` piše „`clan [izvor] -> broj opis`". Opis nema granicu dužine, a traka poruke
       odseca kraj, pa korak ide pre opisa. Log je nepromenjen (izvor ide u svoju kolonu).
     - **(e)** `mScrToast` po obrascu `mGridToast`:
       - zastavicu podižu zabrana i pad gradnje, odmah posle svog `ShowToast`-a;
       - `ActivateScreen` je postao omotač oko `ActivateScreenCore`. Uspešan prelazak gasi poruku na **jednom**
         mestu, jer jezgro ima pet izlaza sa uspehom;
       - `ShowToast` zastavicu spušta, jer nova poruka zamenjuje staru.

     **Testovi i sabotaže** (bez novog testa; `RunAllTests` ostaje 243):
     - 240 + „korak stoji ispred opisa greske, pa ga dug opis ne odseca u toastu" ← `ki016-korak-iza-opisa`;
     - 242 + „uspesan ulazak posle pada gradnje gasi poruku o padu" ← `ki016-poruka-pada-ostaje`;
     - 242 + „uspesan prelazak ne gasi poruku koja je zamenila poruku o padu" ← `ki016-tudja-poruka-nestaje`.
       Drugi pad je pre `Run`-a (`ScrGradnjuOboriTest`), nad ekranom čija zona ne postoji, pa nijedan modul
       ekrana ne dobija stanje;
     - 242 + „uspesan prelazak gasi poruku o zabranjenom prelasku" ← `ki016-poruka-zabrane-ostaje`. Zabrana ide
       kroz `ScrSekcijuZabraniTest`, pa se zona ni ne pravi;
     - novi pomoćnik `ToastNaEkranu`: `ToastTekst` čita natpis, a `HideToast` traku sakriva i natpis ne briše.

     Nijedna sabotaža ne diže grešku kroz `Run`. Katalog 790 → 794.

     **Provera pre review-a.**
     - Jeftine kapije 20/20 (794 sabotaže, 0 nalaza, 9 poznatih).
     - Brza provera, **nije izvršna kapija** (kept sveska nad `b24053ac`, `RunOne`): 1, 84, 103, 109, 153, 159,
       162, 236, 237 i 239–243, `TESTS=14 FAIL=0`, bez dijaloga.
     - `run_vba` i `dokaz` čekaju reviewer GO. Dokaz, 22 sabotaže:

       ```
       dokaz.py ki016- ljuska-deaktivira-pre-provere zona-posle-pada-gradnje omotaci-posle-pada-gradnje
                bu-ulazak-ne-zove-kuku bu-kuka-ignorise-argument --grupe
       ```
       - `ki016-` je 17 sabotaža: četiri nove, `ki016-pomocnik-ne-belezi` sa novim sidrom i 12 zatečenih;
       - ostalih pet imaju sidro u telu koje je postalo `ActivateScreenCore`, ili u grani pada gradnje.

     **Očekivani brojevi:** `RunAllTests` 243 / 0; BFP 2427, Banka 241, Storno 111 bez promene.

     **Uočeno van reza (zapisano, ne dirano):**
     - `mGridToast` nema spuštanje u `ShowToast`. Zastarela zastavica mreže može na sledećem uspelom čitanju
       da sakrije tuđu poruku, i to baš objašnjenje zabrane posle zamene operatera (`NaPrviDozvoljenEkran` →
       `ReloadGrid`).
     - Direktan `HideToast` briše `mToastPending`, a tajmer ne otkazuje. Tajmer zelene poruke tada ostaje
       siroče, a `CancelToastTimer` ga više ne vidi, pa može da sakrije crvenu poruku podignutu u ta 4 s.
       Ovaj rez to ne pojačava: `mScrToast` se diže crvenom porukom, koja tajmer otkazuje.

     **Otvoreno.** Review, pa izvršne kapije, compile i ručna provera.

146. **Review `#425` (`d0c6864e`): `CODE GO`, `P0 0`, `P1 0`, `P2 0`, `P3 1`, 9,3/10** (15.10.2026). Static CI
     #981 `SUCCESS`. `MERGE HOLD` dok se ne izvrše kapije. Produkcioni kod posle GO nije menjan.
     - **`P3-1`, nije blokirajući:** test meri redosled izvora i opisa, ali ne i stvarno odsecanje na formi. To
       meri ručna provera sa pravom greškom šeme Faktura (tačka 1 ispod).

     **Izvršni prolaz** (head `d0c6864e`, izvor `c35dd64b9cc8`, ugovor `75272720898e`, sveska `0bafa704`,
     fixture `d3d3ff6d`): **12/12 ZELENO**, marker upisan.

     ```
     RunAllTests              243 / 0      RunBusinessFlowProSuite 2427 / 0
     RunBankaImportTestSuite  241 / 0      RunStornoTestSuite       111 / 0      ostalih 8 suita OK
     ```

     **Dokaz** (filter ispod; 22 sabotaže u 5 prolaza). Filter ne hvata nijedan od 9 `POZNATI_NALAZI` ni od 2
     `POZNATI_NALAZI_DOKAZ`.

     ```
     dokaz.py ki016- ljuska-deaktivira-pre-provere zona-posle-pada-gradnje omotaci-posle-pada-gradnje
              bu-ulazak-ne-zove-kuku bu-kuka-ignorise-argument --grupe

     crvenih: 22 / 22    izvor pre/posle 3d1bbea1dd0c0248 IDENTICAN    bez ijednog TIMEOUT-a
     grupno izmereno 20 / 22 -> prvi verdikt NIJE DOKAZANO (jedan PROBLEM)
     ```
     Dva člana nisu priznata u grupi:
     - **`ki016-kanal-ne-brise`, zaklonjena u grupi 4, predviđeno iz plana.** `ki016-pomocnik-ne-belezi` gasi
       kanal, pa 240 pada već na prvoj tvrdnji. Sama: OK.
     - **`ljuska-deaktivira-pre-provere`, PROBLEM i kad je sama.** Pada preduslov „ekran je pre panela imao
       svoju zonu". **Zastarela je od `ba91e66e` (03.09.2026), ne od ovog reza:**
       - ista sabotaža nad `main`-om `66f7897f` (privremeni worktree, ista fixture sveska, izvor
         `e59a78ee3c9822c9`) daje isti pad;
       - sidro je bio red posle `jePanel = …`, dakle pre provere „već smo tu" za isti ekran, a ta provera je
         ušla posle sabotaže;
       - korak 2 testa ide `IdiNaEkran` na ekran na kome sekcija već stoji, pa je sabotaža brisala zonu pre
         tvrdnje.

       **Ispravka (posle GO, samo `tools/sabotaza.py`):** sidro je provera prava (`If Not
       modUiScreens.ScrDozvoljen(kljuc) Then`). To je staro mesto poziva, kako i komentar sabotaže kaže. Sama
       nad `d0c6864e`: `DOKAZANO`, izvor identičan. Marker i dalje važi, jer `sabotaza.py` nije deo ugovora.

     **Ukupno: 22 / 22 dokazano** (20 grupno, 2 sama).

     **Nalaz uz scenario 4 iz review-a (pročitano u kodu, ne izmereno).** Review traži da objašnjenje
     `OTKUI_SCR_ZABRANJEN`, koje `PrimeniNovaPrava` ostavlja posle zamene operatera, ostane na ekranu. U
     produkciji se to ne vidi:
     - `DoSwitchOperater` posle `PrimeniNovaPrava` odmah zove `ShowToast OTKUI_MSG_OPERATER`, zelenu poruku sa
       tajmerom od 4 s;
     - na neuspeloj prijavi `ShowToast OTKUI_MSG_ODJAVLJEN`;
     - preostali poziv je samo test seam.

     Objašnjenje zabrane zato prepiše sledeća poruka, sa spuštanjem zastavice i bez njega, kao i pre ovog reza.
     Komentar uz `mScrToast` taj primer navodi kao razlog spuštanja u `ShowToast`, što je netačno kao opis
     onoga što operater vidi.

     Pravi slučaj koji spuštanje štiti je poruka koju podigne **sam prelazak**: pad mreže novog ekrana u
     `ReloadGrid`, ili poruka kuke ulaska (`ScrAktiviraj`, uvoz izvoda u Banci). Omotač gasi poruku tek posle
     jezgra, pa bi bez spuštanja stara zastavica sakrila baš poruku o padu mreže ekrana u koji se upravo ušlo.
     - Test meri isti mehanizam: nova poruka između pada i uspeha, sabotaža `ki016-tudja-poruka-nestaje`.
     - Komentar se ispravlja posle merge-a, u zasebnom malom rezu. Izmena `src-vba` sada bi oborila marker i
       dokaz, a ponašanje se ne menja.
     - Scenario 4 checkliste je zato zamenjen ovim slučajem; scenario iz review-a ostaje kao opcioni.

     **Sveska za compile i ručnu proveru:** `tests/fixtures/rucna_provera_ki016de_kod.xlsm`
     (`run_vba --compile-only --keep` nad `d0c6864e`, izvor `c35dd64b9cc8`). Operater uvozi
     `src-vba/modVbaTools.bas`, kompajlira, zatvara i otvara svesku. Checklista je u opisu `#425`:
     1. greška šeme Faktura: korak se vidi u traci;
     2. pad gradnje Izveštaja, pa uspešan ulazak: traka nestaje;
     3. zabranjen prelazak, pa uspešan prelazak: traka nestaje;
     4. pad gradnje, pa ulazak u ekran čija mreža padne: poruka o padu mreže ostaje.

     **Compile ✅** (operater, 19.10.2026, `rucna_provera_ki016de_kod.xlsm`, kod sa `d0c6864e`): `vba_gate
     --mark-compile`, pa `--require-green --require-compile` daje RC=0 nad izvorom `c35dd64b9cc8`.

     **Ručna provera ✅** (operater, 18.–19.10.2026, ista sveska, snimci ekrana). Nijednom nije bilo dijaloga:

     ```
     1. Fakture, preimenovana kolona   traka nosi korak [modScrFakture.RedoviFakture[fakture]] ISPRED
                                       opisa -- vidljiv (u #423 odsecen)                              OK
     2. pad gradnje IZVESTAJA          ostaje Fakturisanje i "Ovaj ekran jos nije prebacen..."; sledeci
                                       ulazak gradi pun ekran BEZ trake                               OK
     3. zabrana PALETA                 ostaje IZVESTAJI i "Nemate dozvolu za ovu oblast."; sledeci
                                       prelazak otvara PALETE BEZ trake                               OK
     4. Fakturisanje, preimenovana     "Ucitavanje mreze nije uspelo" OSTAJE posle prelaska           OK
        kolona
     ```
     Uz tačku 4: snimak ne pokazuje da je pad gradnje ponovljen na svežoj svesci, jer su IZVEŠTAJI već bili
     izgrađeni. Potvrđeno je krajnje stanje, ne i to da je zastavica prelaska stajala u trenutku prelaska.
     Isti mehanizam, spuštanje zastavice u `ShowToast`, automatski meri kontra-tvrdnja u 242 („uspesan
     prelazak ne gasi poruku koja je zamenila poruku o padu"), sa sabotažom `ki016-tudja-poruka-nestaje`.

     **Otvoreno.** Merge na izričit zahtev, posle zelenog CI na tačnom head-u.

147. **Merge `#424` i `#425`** (18.–19.10.2026). Zapis ide ovde, a ne u PR-ove.

     - **`#424`** (process, samo `.claude/rules/testovi.md`): review krug 1 (`9160d3cf`) `NO-GO`, `P2 1` (ugovor
       `Run`-a bez `UI-PAD-01`), `P3 1` (tvrdnja da `modTest` ne pokriva mrežu). Oba su zatvorena u `88366cf5`,
       zajedno sa KI-016 odeljkom i dovođenjem grane na `main`. Merge-ovao **operater** 18.10.2026 (13:36 UTC)
       kao `33896eb9`, na head `88366cf5`, uz static CI `success`. **Rezultat ponovljenog review-a nije zapisan
       u ovom toku.**
     - **`#425`** (KI-016 (d) i (e)): review (`d0c6864e`) `CODE GO`, `P0 0`, `P1 0`, `P2 0`, `P3 1` (vizuelno
       odsecanje meri ručna provera, scenario 1), 9,3/10, static CI #981 (stavka 146).
     - **Posle tog GO** su ušli `ef0ab8a3` (samo `tools/sabotaza.py`: sidro `ljuska-deaktivira-pre-provere`,
       zastarelo od `ba91e66e`) i docs (`2ea80dd1`, `eeffddd2`). `src-vba` je od `d0c6864e` nepromenjen.
     - Izvršne kapije su nad `d0c6864e` (izvor `c35dd64b9cc8`): 12/12, dokaz 22/22, compile (19.10.2026) i
       ručna provera 4/4.
     - Merge-ovano 19.10.2026 (22:15 UTC 18.10.) kao `06d4a09f`, na **izričit zahtev operatera**, sa
       `--match-head-commit` na head `eeffddd2` (static CI `success`, run 39171262798).
     - **Finalni exact-head review nije tražen.** Reviewer nije video `ef0ab8a3` ni docs posle njega; GO važi
       za `d0c6864e`.
     - `KI-016` ostaje `OPEN`: (a) `P3 #1` (`SCR_RUN_NEMA` ćuti za korisnikov klik), (b) mrtav EH u
       `PanelOtvori`, (c) ostali ciljevi `Run`-a bez brane, (f) životni vek poruke mreže (stavka 148).

148. **`KI-016` (f) — život poruke mreže i primer uz `mScrToast`** (18.10.2026). Grana
     `claude/ki016-toast-zivotni-vek`. Rađena je nad granom `#425` (`2ea80dd1`) dok `#425` nije bio merge-ovan
     (odluka korisnika 18.10.2026: lokalno odmah, PR tek posle merge-a). Posle merge-a je grana napravljena
     iznova sa `main`-a `06d4a09f`: prvo zapis merge-a (stavka 147), pa cherry-pick oba commit-a. `src-vba` i
     `tools` su identični vrhu pre rebase-a (`38e8ff30`).

     **Popis pre koda.**
     - Direktni pozivaoci `HideToast` su samo dva: omotač `ActivateScreen` (`mScrToast`) i `ReloadGrid`
       (`mGridToast`). Treći je `Application.OnTime`, tajmer zelene poruke.
     - `mGridToast` se podizao **pre** `ShowToast`-a i ništa ga nije spuštalo kad druga poruka zameni poruku
       mreže.
     - Za gašenje poruke mreže nije postojao **nijedan** test.
     - Stvaran put je `DoSync`: `ShowToast "sync počinje"` (zelena, tajmer 4 s), pa `RefreshFromData` →
       `ReloadGrid`. Uz zastavicu zaostalu od ranijeg pada liste, direktan `HideToast` gasi „sync počinje"
       i briše `mToastPending` bez otkazivanja tajmera. Tajmer bez vlasnika zatim može da sakrije i crveni
       ishod sync-a: `CancelToastTimer` ga više ne vidi.

     **Reprodukcija pre ispravke** (`fa080084`, nov test 244 uz staru produkciju, `RunOne`, bez dijaloga):

     ```
     TESTS=1 FAIL=1
     FAIL T_KI016_PorukaMrezeNeGasiTudju -- uspelo citanje ne gasi poruku koja je zamenila poruku mreze
          -- ocekivano [ANALIZA|KI016 tudja poruka], dobijeno [ANALIZA|]
     ```
     Tri tvrdnje pre nje su zelene: oba preduslova i kontra-tvrdnja da uspelo čitanje i dalje gasi poruku
     mreže.

     **Ispravka (`787fa46b`).** `mGridToast` po obrascu `mScrToast`: `ShowToast` ga spušta, a `ReloadGrid` ga
     podiže tek **posle** svog `ShowToast`-a.

     Tajmer bez vlasnika posle toga nema puta. Direktan `HideToast` ide samo kad je poslednja poruka crvena
     poruka vlasnika zastavice, a ona je u `ShowToast`-u otkazala svaki tajmer na čekanju.
     - To je **pročitano, ne izmereno**: `Application.OnTime` ne okida dok test radi, pa se tajmer u harness-u
       ne može posmatrati.
     - `HideToast` zato nije menjan.

     **Komentar uz `mScrToast`** (nalaz iz stavke 146): primer „objašnjenje zabrane posle zamene operatera"
     zamenjen je pravim slučajem, porukom koju podigne sam prelazak (pad mreže novog ekrana, kuka ulaska).

     **Test i sabotaže:** nov test 244 `T_KI016_PorukaMrezeNeGasiTudju` (`RunAllTests` 243 → 244):
     - „uspelo citanje ne gasi poruku koja je zamenila poruku mreze" ← `ki016-mreza-gasi-tudju`;
     - „uspelo citanje posle pada liste gasi poruku mreze" ← `ki016-mreza-bez-zastavice`.

     Katalog 794 → 796.

     **Provera pre review-a.**
     - Jeftine kapije 20/20 (796 sabotaža, 0 nalaza).
     - Brza provera, **nije izvršna kapija** (`RunOne` nad `38e8ff30`, isti izvor kao `787fa46b`): 1, 84, 103, 109, 153, 159, 162, 236,
       237 i 239–244, `TESTS=15 FAIL=0`, bez dijaloga.
     - `run_vba` i `dokaz` čekaju reviewer GO. Predlog dokaza: `dokaz.py ki016-mreza- ki016-tudja-poruka-nestaje
       ki016-poruka-pada-ostaje ki016-poruka-zabrane-ostaje --grupe`.

     **Očekivani brojevi:** `RunAllTests` **244** / 0; BFP 2427, Banka 241, Storno 111 bez promene.

     **Otvoreno.** Review, pa izvršne kapije, compile i ručna provera.

149. **Review `#427` (`360f8b6b`): `CODE GO`, `P0 0`, `P1 0`, `P2 0`, `P3 1`, 9,3/10** (19.10.2026). Static CI
     #991 `SUCCESS` na tačnom head-u. `MERGE HOLD` dok se ne izvrše kapije. Produkcioni kod posle GO nije menjan.
     - **`P3-1`:** stvarno ponašanje `Application.OnTime` nije dokazano. Test 244 koristi drugu **crvenu** poruku,
       a `DoSync` zelenu sa tajmerom od 4 s, koji u automatskom testu ne okida. `CancelToastTimer` ima `On
       Error Resume Next`, pa neuspešno otkazivanje ne bi bilo prijavljeno. Reviewer traži ručni acceptance.

     **Izvršni prolaz** (head `360f8b6b`, izvor `c259577ff714`, ugovor `8e1d808997fe`, sveska `0bafa704`,
     fixture `d3d3ff6d`): **12/12 ZELENO**, marker upisan.

     ```
     RunAllTests              244 / 0      (243 + 244)
     RunBusinessFlowProSuite 2427 / 0      RunStornoTestSuite 111 / 0
     RunBankaImportTestSuite  241 / 0      ostalih 8 suita OK
     ```

     **Dokaz** (5 sabotaža u 3 prolaza). Filter ne hvata nijedan od 9 `POZNATI_NALAZI` ni od 2
     `POZNATI_NALAZI_DOKAZ`.

     ```
     dokaz.py ki016-mreza- ki016-tudja-poruka-nestaje ki016-poruka-pada-ostaje ki016-poruka-zabrane-ostaje --grupe

     crvenih: 5 / 5    izvor pre/posle 6e8779c278b21e8c IDENTICAN    bez ijednog TIMEOUT-a
     === DOKAZANO (grupno) ===   (grupno izmereno 4 / 5, peta je po planu sama u grupi; bez zaklanjanja)
     ```

     **Sonda tajmera za `P3-1`** (`sonda_ontime.py`, scratchpad `23d2fe62`).
     - Sopstvena instanca, kept sveska, prava ljuska, put `DoSync` u jednom makrou: pad liste, zelena poruka
       sa tajmerom, uspelo čitanje, crvena poruka.
     - Makro se zatim vrati, a Excel miruje 8 s; `OnTime` okida tek tada.
     - Kontrola je ista sonda nad `main`-om `06d4a09f` (stari kod, privremeni worktree, ista fixture sveska).

     ```
     korak                         STARI 06d4a09f                 NOVI 360f8b6b
     1 posle pada liste            vidljiva, poruka mreze         vidljiva, poruka mreze
     2 posle zelene (tajmer 4 s)   vidljiva, zelena               vidljiva, zelena
     3 posle uspelog citanja       SAKRIVENA (zelena)             vidljiva, zelena
     4 posle crvene                vidljiva, crvena               vidljiva, crvena
     5 posle 8 s mirovanja         SAKRIVENA (crvena)             vidljiva, crvena          bez dijaloga
     ```
     Na starom kodu je tajmer bez vlasnika **stvarno okinuo** i sakrio crvenu poruku. Na novom je crvena
     ostala, jer je njen `ShowToast` otkazao tajmer zelene. Kvar iz `KI-016` (f) je time izmeren u oba smera,
     stvarnim `OnTime`-om, a ne samo pročitan (stavka 148).

     Granica: sonda ide kroz `ShowToast` i `IdiNaEkran`, ne kroz sam `DoSync` (Sync zahteva Google vezu).
     Pravi `DoSync` ostaje ručnoj proveri, tačka 1.

     **Sveska za compile i ručnu proveru:** `tests/fixtures/rucna_provera_ki016f_kod.xlsm`
     (`run_vba --compile-only --keep` nad `360f8b6b`, izvor `c259577ff714`). Operater uvozi
     `src-vba/modVbaTools.bas`, kompajlira, zatvara i otvara svesku. Checklista je u opisu `#427`:
     1. `DoSync` posle pada liste: crvena poruka o neuspelom sync-u ostaje i posle 10 s;
     2. pad liste, pa uspelo čitanje bez druge poruke: crvena poruka mreže nestaje.

     **Ispravka postupka tačke 1 (19.10.2026).** Prvo izdanje je tražilo pad liste na Fakturama, pa Sync. To se
     ne može izvesti: dugme Sync postoji samo u desnom panelu ekrana Unos dokumenata (`zRight`, vidljiv samo
     na tom ekranu). Prelazak tamo ponovo učita mrežu, a to uredno ugasi poruku mreže i njenu zastavicu, pa Sync
     ne bi merio ništa. Pad liste zato ide baš na Unosu dokumenata, kroz seam
     (`ScrPadTestSet "modScrDokumenti.Scr_Rows"`). Zatim se seam i test režim gase, pa sledi Sync bez interneta.

     **Ručna provera, tačka 1 ✅** (operater, 19.10.2026, snimci 07:20 i 07:21, internet isključen):
     - pad liste Unosa dokumenata daje traku „Učitavanje mreže nije uspelo na koraku: modScrDokumenti.Scr_Rows
       [modUiScreens.ScrPadAkoTest[modScrDokumenti.Scr_Rows]] -> 9 …";
     - zatim „Sinhronizuj", pa crvena „Sinhronizacija nije uspela — vidi Immediate/log." **ostaje** na ekranu i
       posle 10 s (potvrda operatera), a panel sinhronizacije kaže „Sinhronizacija nije uspela";
     - bez dijaloga.

     **Uočeno uz tačku 1, nije istraženo:** posle Sync-a zaglavlje prelazi sa „Sve izmene sačuvane" na
     „Nesačuvan unos", iako je forma prazna. `RefreshFromData` ne zove `SelectModeCore`, pa to nije doslovno
     put iz duga „badge nesacuvano na novoj formi". Kandidat je ponovno punjenje combo-a (`FillZbirneCombo`,
     `RefillMaticniCombos`) koje okida `MarkDirty`, ali to nije provereno. Ide uz taj dug, ne uz ovaj rez.

     **Ručna provera, tačka 2 ✅** (operater, 19.10.2026, snimci 07:49 i 07:50, Unos dokumenata u režimu
     Otpremnica):
     - seam uključen, pa klik na drugu listu: crvena „Učitavanje mreže nije uspelo na koraku:
       modScrDokumenti.Scr_Rows […]";
     - seam ugašen, pa klik na „Sve": lista se učita (27 otpremnica) i crvena traka **nestane**;
     - bez druge poruke između i bez dijaloga.

     **Compile ✅** (operater, 19.10.2026, `rucna_provera_ki016f_kod.xlsm`, kod sa `360f8b6b`): `vba_gate
     --mark-compile`, pa `--require-green --require-compile` daje RC=0 nad izvorom `c259577ff714`.

     **Otvoreno.** Merge na izričit zahtev, posle zelenog CI na tačnom head-u.

## Dug sa imenom (posle S5-5b)

| Stavka | Zašto stoji, a ne „kasnije ćemo“ |
|---|---|
| **`MsgBox` u pisac-putanji visi u `run_vba` prolazu** | Protokol potvrde deficita pita operatera na **tri** mesta (`modOtkupUnos` od 03.10.2026, `modNovacUnos` i `modDokUnos` od 07.10.2026). Test koji uđe u tu granu ne pada nego **visi do timeout-a** i ostavlja Excel u `[break]` — ista cena kao compile greška (585 s + ubijen Excel). Danas to drže samo komentari uz tri grane; kapija bi morala da zna koji su pozivi iz suite-a dostupni, pa traži svoj rez i svoj dvosmerni dokaz |
| **`dokaz.py` ne dosegne `modIzvestajTests`** | kapija kataloga sabotaža priznaje samo `modTest`, `modTestBanka` i `modBusinessFlowProTests`, pa tvrdnja iz `modIzvestajTests` **ne može da se obori** — a tamo živi najdetaljnije merenje ambalažnog pregleda. Posledica je izmerena u `10c-2`: tvrdnja o identitetu dokumenta je **preseljena** u `#204`, a tvrdnja o spajanju dve vrste istog dokumenta ostala **bez sabotaže**. Proširenje kapije dira sam alat, pa ide **zaseban process PR** sa svojim dvosmernim dokazom |
| **kanonska knjiga nema PUNU integritetnu proveru** | `modIntegritet` ima **nula** referenci na `Od_Tip` / `Na_Tip` / `VrstaKretanja` / `tblAmbalazaDokument` (merenje 08.10.2026). Posle cutovera proverava samo legacy redove, koje produkcija ne piše — pa operateru kaže „nema nalaza" ne pogledavši kanonsku knjigu. Kandidati za provere: noga bez zaglavlja, zaglavlje bez noge, `StornoOd` koji ne pokazuje nigde, par koji nije neuređen (`AMB-INV-10`), obaveza partneru < 0 (`AMB-INV-09`). Od stavke 113 ugovor reda, dupli `AmbID` i `StornoOd` meri `Chk_AMB_KnjigaIntegritet`; od `10e-3` (stavka 129) i red koji ne dotiče knjigu (red starog oblika) je nalaz po imenu, jer „legacy" stanja više nema. Ostali kandidati (noga bez zaglavlja, zaglavlje bez noge, par koji nije neuređen, obaveza < 0) su i dalje otvoreni. `Chk_B10_ReversBezID` je obrisan u `10e-1` |
| **`GetAmbalazeStanje` guta grešku i vraća prazno** | `On Error GoTo EH → LogErr → Empty` je fail-open na putanji **štampe i izveštaja**: saldo koji tiho postane 0 je netačna tvrdnja operateru, ne odsustvo podatka. Komentar uz `AmbSaldoNaloga` to već imenuje („zatečen `GetStanicaAmbSaldo` tako radi i to je fail-open koji ovde ne sme da postoji"). Nije dirano u `10c-2` jer je to politika greške, ne model podatka — promena bi oborila štampu tamo gde danas štampa nulu; traži svoj rez i odluku šta operater vidi kad knjiga ne može da se pročita. **Odluka operatera (13.10.2026): sledeći korak posle `KI-012`/`KI-013` (stavka 133)**, zajedno sa `GetStanicaAmbSaldo` i istorijskim početnim stanjem na ponovljenom papiru reversa. **Odluke operatera za taj rez (13.10.2026, stavka 135), sve po preporuci:** (1) kad knjiga ne može da se pročita, papir **ne nastaje** (fail-closed, kao otkupni list po ODL-27), a oba čitaoca gube handler; (2) istorijsko stanje dobijaju **revers i otpremnica**, a grupni otkupni list ide u S6 (danas je nedostižan i čita pogrešan nalog); (3) kod pokrića deficita važi **pravilo reversa** (novo = stanje posle dokumenta, početno = novo − kretanje), a prikaz pokrića ide uz `ODL-24`; (4) **kupčev revers se iz F7 ne štampa**, kao iz liste (7507) |
| **grupni otkupni list čita pogrešan nalog** | `FillGrupniOtkupSablon` čita `GetStanicaAmbSaldo(VozacID)`, dakle `Stanica:<VozacID>`, i to današnji saldo, dok prijemnica od 10b-2 pomera nalog `Vozac`. Danas je nedostižan (štampa se samo posle upisa prijemnice, a unos je pauziran do S6), pa se u stavci 135 menja samo politika greške. Koji je nosilac salda na tom papiru i da li dobija stanje oko svog reda odlučuje **S6** (`AMBALAZA.md` 6.13g). Uz to ide i `GetAmbalazeStanje`, koji tip grupiše u rečniku sa **binarnim** poređenjem (`letvarica` i `LETVARICA` su dva reda), dok kanonski saldo i istorijski čitalac porede bez obzira na velika slova (review `#421`, P2 #2, stavka 136). Posle stavke 135 ga zovu samo ovaj list, golden i dva test pomoćnika |
| **badge "nesacuvano" na novoj formi** | `SelectModeCore` pise u polje broja POSLE `MarkClean` i `mLoading = False`, pa programski upis prodje kroz `MarkDirty` i prazna forma tvrdi da ima neupisanih izmena (review 08.10.2026, `P3`). Zatecen obrazac -- vazio je i pre P1 ispravke, za svaki rezim sa auto-brojem. Jedan premesten red, ali izmena `src-vba` obara compile i zeleni marker, a nov test (natpis u zaglavlju) trazi svoj dvosmeran dokaz; zato **svoj rez posle merge-a**, pre `ODL-24` |
| **`AMB-10-ODL-24`: pozajmica ambalaže od kupca nema svoj događaj** | Operater (08.10.2026): kupci **često** pre sezone predaju **svoje** prazne gajbe. Brojke su danas tačne — kroz potvrdu manjka nastaje `ULAZ_TUDJE` (obaveza +N) i `POVRAT_PRAZNE` — ali **planirana pozajmica i neobjašnjeno odstupanje ostavljaju isti trag**, pa se posle ne razlikuju; operater za redovan posao dobija pitanje o „manjku". Da postane svoj događaj traži izmenu **zatvorenog** `VrstaKretanja` enuma, formule obaveze (`AMB-INV-09`) i čitalaca u `10c` — i rešenje čvora: eksplicitan ulaz tuđe ambalaže bi sa **pokrićem deficita** delio par i vrstu na istom dokumentu. Puna merenja: `AMBALAZA.md` 6.12k. **Redosled je operaterov: posle `10c` i merge-a** |
| **nema kapije „modul ne sme da koristi tuđ `Private` simbol"** | VBA kompajlira **na zahtev**, pa `Sub or Function not defined` pukne tek kad neki test prvi put pozove baš tu proceduru — i to posle **600 s i ubijenog Excela**, uz poruku bez fajla i linije (06.10.2026: `MarkRowStornirano`, `Private` u `modStorno`, pozvan iz `modAmbalaza`). Ime **postoji** u projektu, samo nije vidljivo — pa ga nijedna jeftina kapija ne vidi. Jednokratni merač je napisan i dao **1 nalaz sa fajlom i linijom nad pokvarenim izvorom, 0 posle** — dvosmeran dokaz. Tri lažna nalaza prvog izdanja su i sama merenje: repni komentar, labela (`Resume CleanUp`) i **LF kopija iz git-a** (split po `
` dao je ceo fajl kao jedan red, pa je prvi „čist" prolaz bio lažan). Kao trajna kapija ide u `vba_check`, dakle **zaseban process PR** |
| **nema kapije „nastavak reda je poslednji znak u redu"** | regex nad VBA izvorom ume da pojede **prelom reda** (`\s*` hvata i `\r\n`), pa `_` ostane usred linije — sintaksna greška koju nijedna jeftina kapija ne vidi. Cena je nesrazmerna: **585 s i ubijen Excel**, a poruka je samo „Compile error: Syntax error" bez mesta (06.10.2026, rez reversa). Jednokratni merač je napisan i dao **2 nalaza pre ispravke, 0 posle** — dvosmeran dokaz po konstrukciji, nad celim `src-vba`. Kao trajna kapija ide u `vba_check` i traži svoj dvosmerni dokaz, dakle **zaseban process PR** |
| **nema kapije za „snimak testa mora pokriti sve što pisac piše"** | cutover proširi šta jedan pisac upisuje, a testovi sa **svojom** transakcijom i dalje snimaju stari skup tabela — pa inner commit preživi outer rollback i ostavi **sirotana** koji padne u tuđem testu (05.10.2026, stavka 65). Merenje je jednokratno napisano i dalo **1** nalaz u testovima i **0** u produkciji; kao trajna kapija traži mapu „pisac → tabele" iz `WHO_WRITES` i svoj dvosmerni dokaz, dakle **zaseban process PR** |
| **redosled u VLASNICKOJ grani kaskade nema test** | eksterna grana je pokrivena od P1 ispravke (`Test_PRJ_EksternaPrijemnicaBlokiraPonistenje` nad pravom kaskadom, kroz test seam). Vlasnička (`ownsChain = True`) nije: `ZbirnaOwnsExternalChain` je istina samo kad je kupac **konfigurisana** hladnjača (`CFG_MALINA_DEFAULT_KUPAC`), pa bi test morao da menja podesavanja — mutacija configa u suite-u je sama rizik. Za tu granu je izmereno **svojstvo** (`Test_PRJ_LanacSeOdmotavaObrnuto`), ne redosled; sabotaža koja bi vratila stari red **nije upisana** jer se ne bi videla, a sabotaža koja ne obara ništa je placebo |
| **bruto grana otkupa/otpremnice bez testa** | posledica pina `OTKUP_BRUTO_UNOS = NO` u `make_fixture` (KI-008): tara, odbijanje kad `tara >= kolicina` i zamrzavanje `BrutoKg` nemaju **ni jedan** test. Njen test mora sam da postavi zastavicu, kao `modIzvestajTests` za `MALINA_MODE` — nasleđivanje od donora je ono što je pet padova i napravilo |
| `dispecer.js` alokacija po klasama | **poslovna odluka**, ne prevod: raspodela količine na više klasa traži pravilo od operatera. Dok je N=1 ponašanje je identično |
| `OTKUP_CONFLICT` lifecycle | deterministički konflikt ostaje retryable pending — vidljivo i bezbedno, ali traži svoj rez |
| `DEGRADIRANO` grana ciklusa | `modGoogleSyncOrchestrator:384` — poslednji OTK razlog je nestao, sama grana nije |
| `BuildOTKFixtureData` (smoke) | gradi pre-S5-5b oblik žice; suite je zatečeno crven i van FULL prolaza, a izmena se **ne može izmeriti** bez živog Google-a |
| **pet test suite-ova koje nijedna kapija ne pokreće** | `RunHttpUtilsSmokeSuite`, `RunSEFDocumentIdShapeSuite`, `RunSEFStateTransitionSuite`, `RunSEFClientParserSmokeSuite`, `RunSEFOfflineSuite`. Nađene `vba_gate.py --popis`-om; do tada nisu bile zapisane nigde (petu je našao razlagač deklaracije — ima jedan **opcioni** argument, pa ju je izraz koji je tražio praznu listu preskakao). Nijedna nema `Err.Raise` u telu, pa bi i priključena bila `BLIND` — zato unos u `SUITES` nije dovoljan posao. Razlog i šta ga zatvara stoje u `SUITE_VAN_KAPIJA` |
| `.claude/rules/testovi.md` ne zna za JS kapiju | **samo process PR**, nikad uz feature izmenu |
| Node 20 deprecation u tri GitHub akcije | process PR |
| `popis_citalaca` javlja UPOZORENJE za `IzvedeniLanacIzPwaDostupan` | kapija ne postoji od #388 — očekivanje alata je zastarelo |
| **69 BFP tvrdnji nije bilo dokazivo — kapija ih je merila kao PODNIZ** | `dokaz.py` za BFP belezi **naziv tvrdnje** (runner ne ispisuje ime Sub-a), pa je tvrdnja **ključ** i mora biti **ceo statičan literal**. `vba_check` ju je merio kao podniz, uz obrazloženje u kodu da „dokaz.py isto radi podniz" — **netačno**. Posledica: odrezana (`'NACRT otpremnice se ne nudi'` umesto `"ZBR nevezane: NACRT otpremnice se ne nudi"`) i dinamička tvrdnja prolaze kapiju, a `dokaz.py` ih 25 minuta kasnije javi kao **NE OBARA SVOJ TEST**. Nađeno prvim prolazom dokaza nad `10b-2`, 04.10.2026: **5 od 8 problema** je bilo tačno to. Kapija je **pojačana** (claim mora biti ceo statičan literal), merenje je dalo **69** takvih; **57** je sireno do celog literala mehanički, **4** ambalažna su ispravljena u testu (vrednosti u svoju tvrdnju; dva literala spojena u jedan), a **8** traži izmenu tuđeg testa (OTP/ZBR/OTK) i stoji imenovano u `POZNATI_NALAZI` sa receptom. **10e-2 (stavka 127):** `otk-storno-izvora-aktivne-otpremnice` izmenu testa nije tražila. Test je već imao pun literal, pa je ispravljen samo katalog i ostalo ih je **7**. Popis je zatvoren — nova ne može tiho da uđe |
| **`NEDEKLARISAN` ne vidi provučenu transakciju** | `OtpIspravi` je koristio `tx` iz scope-a **pozivaoca** — uz `Option Explicit` to je `Variable not defined`, dakle glava se ne kompajlira. Ni jedna jeftina kapija to nije prijavila: arity sweep je **zelen** jer je broj argumenata tačan (`StornoOtpremnica(staraID, tx)` = 2), a nedeklarisana promenljiva je **semantika**, ne arnost. Našao ga je **review**, 04.10.2026 — drugi put u istom rezu da se `tx` provlači pogrešno. Jednokratni merač (`scratchpad/scope_scan.py`) je pušten nad celim `src-vba`: **5522 procedure, tačno 1 nalaz** (baš taj), i **0 posle ispravke** — dvosmeran dokaz po konstrukciji. Obim je namerno uzak: samo imena negde deklarisana kao `clsTransaction`; pun resolver bi tražio `With`, `For Each`, implicitne tipove. Dva lažna nalaza prvog izdanja su i sama merenje: `Dim tx As **New** clsTransaction` (12× u `modNovac`) i polje klase `TX As MSForms.TextBox` (`clsFlatBtn`) — zato merač sada skuplja **svako** deklarisano ime, bez obzira na tip. Usvajanje kao trajne kapije (ili jačanje `NEDEKLARISAN`-a) je **zaseban process PR** |
| **`ARNOST` ne vidi poziv u izraznoj poziciji** | kapija gleda samo statement pozive, pa je promena potpisa `CreateOtkup` prošla zelena dok se projekat ne bi kompajlirao — nađeno **čitanjem**, 03.10.2026. Jednokratni sweep koji gleda i izraznu poziciju (`scratchpad/sweep_arnost.py`) je pušten nad celim `src-vba`: **299 pozivnih mesta, 22 imena, nula nalaza**, uz dvosmerni dokaz nad **kopijom** izvora (tri oblika slomljenog poziva daju crveno po imenu i fajlu). **Granica je izmerena, ne tvrđena:** sweep vidi samo **arnost**, pa zamena tipa pri istom broju argumenata (`tx` izbačen, a poziv popunjava opcione argumente) ostaje zelena — to je slučaj koji **samo Compile** hvata, i tačno taj kvar je i nastao. Usvajanje kao trajne kapije je **zaseban process PR**: nad svim procedurama (ne samo nad 22) može da iznese zatečene nalaze, pa nije posao feature reza |
| **`vba_gate` pamti JEDAN compile, a više suita** | `marker["suites"]` je rečnik i svaka suita nosi svoj `izvor`, a `marker["compile"]` je **jedan objekat** — pa `--mark-compile` na drugoj grani pregazi potvrdu prve. Nađeno 03.10.2026 na #405/#406: operater je kompajlirao oba izvora, a marker je zadržao samo zadnji. Kapija to **tačno prijavljuje** (`<-- DRUGI IZVOR`), pa nema lažnog zelenog — ali rad na dve grane šalje compile u ping-pong. Ispravka je `compile` po `izvor`-u, kao `suites`; menja `kapija` deo ugovora, pa traži svoj dvosmerni dokaz i obara ZELENO suita (ne i compile, koji ključa samo na `izvor`). **Zaseban process PR**, ne uz feature rez |


## Alati i kapije

- Popis starog modela i DUAL READ: `python tools/popis_citalaca.py` (prag slajsa: nula živih referenci na kolone starog modela).
- **Prag po grupi (od S3b-1): `python tools/popis_citalaca.py --check`** — živa mesta po grupi ne smeju preko praga iz `PRAGOVI`;
  merenje ISPOD praga takođe pada (zastareo prag pušta grupu da naraste nazad bez ijednog crvenog).
- Pre push-a: `python tools/vba_check.py`, `python tools/who_writes.py --check` i `--check-ownership`,
  `python tools/gen_schema_module.py --check`; ponašanje: `python tools/run_vba.py --suite <ime>`.
- **Grupni dokaz: `python tools/dokaz.py <filter> --grupe`** — do 6 mutacija u **jednom** prolazu suite-a.
  Nad celim katalogom 649 → **118** prolaza (5,5×), na rezu `amb-pisac` 20 → **9** (2,2×); `--plan` ispiše
  podelu bez Excela. Verdikt je `DOKAZANO (grupno)` jer je tvrdnja sprovedena nad mutacijama grupe
  **zajedno**: u rezu se pušta grupno, **pred release pojedinačno** (isti poziv bez `--grupe`). Član koji
  u grupi ne obori **svoju** tvrdnju ne dobija priznanje nego se ponavlja sam. Pravila i cena svakog od
  njih: `docs/EXCEL_TEST_HARNESS.md` → „Grupni dokaz".
- **Popis test suita: `python tools/vba_gate.py --popis`** — suite u kodu vs katalog `SUITES` vs registar
  `SUITE_VAN_KAPIJA`. Ide i kroz `vba_check` (dakle kroz hook) i kroz CI. Hvata napisanu suite koju
  nijedna kapija ne pokreće, fantom u katalogu, zastareo unos u registru, i unos koji **postoji ali nije
  ulazna tačka**. „Javna deklaracija negde" nije isto što i „`Application.Run` to može pozvati": pojam
  `je_ulazna_tacka` traži **`.bas` + javna + `Sub`/`Function` + nula obaveznih argumenata + bezuslovna**
  (deklaracija u klasi, formi ili u `#If` grani zato ne zadovoljava katalog), a nalaz imenuje koji uslov
  je pao. Deklaracije se pamte kao **lista po imenu**, pa ishod ne zavisi od redosleda čitanja fajlova.
  Deklaracije idu kroz
  **dva deljena sloja** u `vba_check` — `logicke_izjave` (spaja nastavke ` _`, skida komentar van string
  literala, trpi uvlačenje) pa `deklaracija_procedure` (vidljivost / vrsta / ime / argumenti / broj
  **obaveznih**). Oba deli i kapija `DUPLIKAT`; prelazak je izmeren, skup javnih imena je identičan.
  Uslov za kandidata je „javna + `.bas` + nula obaveznih argumenata + ime po konvenciji" — ne sve širi
  izraz, jer je tri kruga review-a pokazalo da je problem bio **sloj**, ne izraz.
- **Marker zelenog: `python tools/vba_gate.py --require-green`** — rezultat **po suite-u** iz poslednjeg
  run-a (`tests/last_green.json`, gitignored; piše ga `run_vba.py`), uz **dva otiska**: `izvor` (`src-vba`)
  i `ugovor` (`izvor` + imenovani delovi `runner` / `fixture` / `kapija` / `golden` + verzija markera).
  Odgovara na „da li je **baš ovaj** izvor dokazan **pod ovim test sistemom**" — što je do sada bila
  rečenica uz PR. `kapija` je sam `vba_gate.py`: promena onoga što odlučuje **šta se priznaje kao dokaz**
  mora da obori stare dokaze bez ručnog bumpa verzije. Compile je vezan **samo za izvor**
  (`--mark-compile`), pa promena golden fajla ne obara potvrdu. Kontekst sveske je **identitet** (putanja
  + heš sadržaja izvorne sveske), ne ime — fixture je gitignored, pa bi zamena sveske inače prošla pod
  starim dokazom. Kontekst se **snima pre run-a** (nad temp kopijom sveske, onom koju Excel otvara) i
  upis je **fail-closed**: ako se `src-vba`, ugovor ili sveska promene **tokom** run-a, prolaz može biti
  zelen a marker se ne upisuje — inače bi GREEN bio pripisan stanju koje Excel nikad nije video, a
  prolaz traje 20–60 min uz paralelan razvoj. Ne upisuje ni: pao run, `--no-import` run, `BLIND` suite
  kao dokazanu, ni upis bez snimka konteksta.
- `vba_check` kapije nad **alatima** (katalog sabotaža, pravila grupisanja, popis suita) idu **ispred** izlaza
  `if not files: return 0` — hook sa putanjom koja nije VBA fajl ih je dotad preskakao, uključujući
  baš `tools/sabotaza.py`, gde se greška u katalogu i pravi.
- Poznati živi kvarovi van refaktora: `docs/KNOWN_ISSUES.md` AUD-055..057.
