# Stanje refaktora „dokument = header + stavke“

> **Ulaz za svaku novu sesiju.** Kratko i ažurno — čita se umesto celog plana. Pun plan i istorijat odluka:
> `docs/REFAKTOR_DOKUMENT_HEADER_STAVKE.md` (odluke po datumu u §14.x; važeće: §14.7 „Odluke operatera 16.09“).
> Ažurira se na kraju svakog koraka, u istom commit-u.

**Ažurirano:** 27.09.2026 (S5-5b spojen, #394 `561fcad9`).

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
| **AMB-10 ambalaza kao knjiga prenosa** | ⏳ **ide PRED S6** -- model ✅ · `10a` ugovor ✅ (#398) · `10-DOK` zaglavlje ✅ (#399) · **`10b-1` pisac ✅** · `10b-2` cutover ⏳; `docs/DOMEN/AMBALAZA.md` |
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

## Dug sa imenom (posle S5-5b)

| Stavka | Zašto stoji, a ne „kasnije ćemo“ |
|---|---|
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
| **69 BFP tvrdnji nije bilo dokazivo — kapija ih je merila kao PODNIZ** | `dokaz.py` za BFP belezi **naziv tvrdnje** (runner ne ispisuje ime Sub-a), pa je tvrdnja **ključ** i mora biti **ceo statičan literal**. `vba_check` ju je merio kao podniz, uz obrazloženje u kodu da „dokaz.py isto radi podniz" — **netačno**. Posledica: odrezana (`'NACRT otpremnice se ne nudi'` umesto `"ZBR nevezane: NACRT otpremnice se ne nudi"`) i dinamička tvrdnja prolaze kapiju, a `dokaz.py` ih 25 minuta kasnije javi kao **NE OBARA SVOJ TEST**. Nađeno prvim prolazom dokaza nad `10b-2`, 04.10.2026: **5 od 8 problema** je bilo tačno to. Kapija je **pojačana** (claim mora biti ceo statičan literal), merenje je dalo **69** takvih; **57** je sireno do celog literala mehanički, **4** ambalažna su ispravljena u testu (vrednosti u svoju tvrdnju; dva literala spojena u jedan), a **8** traži izmenu tuđeg testa (OTP/ZBR/OTK) i stoji imenovano u `POZNATI_NALAZI` sa receptom. Popis je zatvoren — nova ne može tiho da uđe |
| **nevaljani ambalažni fixture redovi u `modTestStornoCentar`** | `KnjigaIntegritet` odbija red koji **ne dotiče knjigu** a nije ni **valjan stari red** — takav bi „tiho nestao iz svakog salda". Prvi Windows prolaz 10b-2 ga je i našao: `Test_StornoJournalUndo_Auto` je sejao `tblAmbalaza` red sa samo tri kolone, pa je `StornoOtkup_TX` pao i **šest** provera je palo iz **jednog** raise-a. Kapija je zatečena (10b-1) i u pravu; cutover je samo prvi put stavio kanonskog čitaoca knjige na storno putanju. Taj red je **ispravljen**. Ali `TcSeedRevNoga` i `colsV` (`modTestStornoCentar.bas:465`) seju redove bez `Smer`/`TipAmbalaze`/`Kolicina` — **isto nevaljane**, danas **nedostižne** jer na tim putanjama nema čitaoca knjige. Ne diraju se u ovom rezu: dodavanje `Smer`-a može da promeni verdikt `ReversIDGranica` i obori 127 provera koje prolaze, a nalaz je latentan. **IZMERENO uz cutover otpremnice (04.10.2026): još su nedostižni.** `StornoOtpremnica` sada čita knjigu, ali nijedan od pet testova koji seju te redove (`Test_StornoRevers*`, `Test_StornoJournalReversGuard_Auto`, `Test_UndoReverseGuard_Auto`) ne stornira otpremnicu; a `AmbImaKontraStav` čita tabelu **direktno** (`GetColumnIndex`), ne kroz `KnjigaZaCitanje`, pa ni undo garda ne uvodi integritetsku kapiju na tu putanju. Zatvara ih **prijemnica** ili `10c` — šta god prvo stavi `KnjigaZaCitanje` u isti tx sa njima |
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
