# Stanje refaktora „dokument = header + stavke“

> **Ulaz za svaku novu sesiju.** Kratko i ažurno — čita se umesto celog plana. Pun plan i istorijat odluka:
> `docs/REFAKTOR_DOKUMENT_HEADER_STAVKE.md` (odluke po datumu u §14.x; važeće: §14.7 „Odluke operatera 16.09“).
> Ažurira se na kraju svakog koraka, u istom commit-u.

**Ažurirano:** 21.09.2026 (S4-2b).

## Pravila koja važe (16.09.2026)

- **Nema produkcije ni podataka koje treba štititi — radi se iznova.** Bez migracija, backfill-a, fallback-a.
- **Legacy kod ne mora da radi između faza.** Ne gradi se ništa što ga čuva živim: pauze, podela kapija, test-only pisci,
  kolone ostavljene za pauzirane čitaoce, mostovi.
- **Apsolutno se čuva samo mapa sposobnosti** — sve što operater danas može da uradi ili dobije.
- **Jedino tehničko ograničenje:** posle svakog PR-a projekat se kompajlira (inače pada ceo `run_vba`). Legacy se briše,
  ne ostavlja polomljen.
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
| Kod slajsova (otpremnica, zbirna, prijemnica, faktura, paleta, sledljivost, brisanje) | ⏳ |

## Sledeći korak: S4-2c — ekrani zbirne (F3 forma + radni sto u F2)

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
30. **Sledeće:** **S4-2c/2** (ekrani: F3 forma nad nacrtom + pregled svih zbirnih, radni sto za izvore u F2,
    direktan izbor, skidanje pauze, brisanje starog pisca), pa S4-3 (članstvo, storno okvir, ZBR-KANON-03),
    S4-4 (malina auto-zbirna), pa **S5** (PWA sync, pre njega otkup u `PROSLEDJENO` kao izvor), pa **S3e-2**
    (brisanje kolona kad popis pokaže nulu), pa S6 (prijemnica, F4).

## Alati i kapije

- Popis starog modela i DUAL READ: `python tools/popis_citalaca.py` (prag slajsa: nula živih referenci na kolone starog modela).
- **Prag po grupi (od S3b-1): `python tools/popis_citalaca.py --check`** — živa mesta po grupi ne smeju preko praga iz `PRAGOVI`;
  merenje ISPOD praga takođe pada (zastareo prag pušta grupu da naraste nazad bez ijednog crvenog).
- Pre push-a: `python tools/vba_check.py`, `python tools/who_writes.py --check` i `--check-ownership`,
  `python tools/gen_schema_module.py --check`; ponašanje: `python tools/run_vba.py --suite <ime>`.
- Poznati živi kvarovi van refaktora: `docs/KNOWN_ISSUES.md` AUD-055..057.
