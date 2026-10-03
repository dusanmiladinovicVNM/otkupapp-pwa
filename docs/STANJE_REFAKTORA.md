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
    **Dokaz je PLANIRAN, ne izmeren:** `dokaz.py` i `run_vba.py` čekaju reviewer GO
    (v. `skupe-kapije-cekaju-reviewer-go`). Statički: 19/19 kapija `rc=0`, katalog
    649 → **654**, a `vba_gate --require-green` tačno javlja `rc=2` — izvor je
    promenjen, pa nijedna suite nije dokazana nad njim.

## Dug sa imenom (posle S5-5b)

| Stavka | Zašto stoji, a ne „kasnije ćemo“ |
|---|---|
| **bruto grana otkupa/otpremnice bez testa** | posledica pina `OTKUP_BRUTO_UNOS = NO` u `make_fixture` (KI-008): tara, odbijanje kad `tara >= kolicina` i zamrzavanje `BrutoKg` nemaju **ni jedan** test. Njen test mora sam da postavi zastavicu, kao `modIzvestajTests` za `MALINA_MODE` — nasleđivanje od donora je ono što je pet padova i napravilo |
| `dispecer.js` alokacija po klasama | **poslovna odluka**, ne prevod: raspodela količine na više klasa traži pravilo od operatera. Dok je N=1 ponašanje je identično |
| `OTKUP_CONFLICT` lifecycle | deterministički konflikt ostaje retryable pending — vidljivo i bezbedno, ali traži svoj rez |
| `DEGRADIRANO` grana ciklusa | `modGoogleSyncOrchestrator:384` — poslednji OTK razlog je nestao, sama grana nije |
| `BuildOTKFixtureData` (smoke) | gradi pre-S5-5b oblik žice; suite je zatečeno crven i van FULL prolaza, a izmena se **ne može izmeriti** bez živog Google-a |
| **pet test suite-ova koje nijedna kapija ne pokreće** | `RunHttpUtilsSmokeSuite`, `RunSEFDocumentIdShapeSuite`, `RunSEFStateTransitionSuite`, `RunSEFClientParserSmokeSuite`, `RunSEFOfflineSuite`. Nađene `vba_gate.py --popis`-om; do tada nisu bile zapisane nigde (petu je našao razlagač deklaracije — ima jedan **opcioni** argument, pa ju je izraz koji je tražio praznu listu preskakao). Nijedna nema `Err.Raise` u telu, pa bi i priključena bila `BLIND` — zato unos u `SUITES` nije dovoljan posao. Razlog i šta ga zatvara stoje u `SUITE_VAN_KAPIJA` |
| `.claude/rules/testovi.md` ne zna za JS kapiju | **samo process PR**, nikad uz feature izmenu |
| Node 20 deprecation u tri GitHub akcije | process PR |
| `popis_citalaca` javlja UPOZORENJE za `IzvedeniLanacIzPwaDostupan` | kapija ne postoji od #388 — očekivanje alata je zastarelo |


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
