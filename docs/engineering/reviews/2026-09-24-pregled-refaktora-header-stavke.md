# Pregled rada 05.09 – 24.09.2026: refaktor „dokument = header + stavke"

> Spoljni code review, urađen kao „elite VBA developer" pregled. Opseg: sve što je
> spojeno u `main` posle PR #263 (`34aeedc`, 04.09.2026) do PR #388 (`6b23dd0`,
> 24.09.2026) — 124 PR-a, 380 commit-a, 166 fajlova, +80.097 / −29.864 linija.
>
> **Metod.** Svaki nalaz označen „POTVRĐENO" proveren je čitanjem koda na
> `origin/main` `6b23dd0`, uz grep pozivalaca. Statičke kapije su izvršene u
> Linux sesiji. **Excel suite i `Debug → Compile` nisu izvršeni** — ta strana je
> preuzeta iz PR opisa i označena kao tvrdnja, ne kao dokaz. Delovi pregleda su
> delegirani u šest dubinskih čitanja (kanonski pisci, self-update, PWA sync,
> jezgra zbirne, storno okvir, GAS/PWA uvoz); njihove P0/P1/P2 tvrdnje su ponovo
> proverene u kodu pre nego što su ušle ovde.

---

## 1) Verdikt

Ovo više nije isti projekat kao 4. septembra. Odluka od 16.09 („nema produkcije,
radi se iznova, bez migracija; legacy ne mora da radi između faza") pretvorila je
rad u fundamentalni refaktor domena. Kvalitet inženjerstva unutar tog okvira je
visok: GUID identitet, jedan pisac po tabeli sa snimkom svake tabele koju dira,
šema iz koda sa CI kapijom, osam novih statičkih kapija sa self-testovima,
spoljni review po PR-u koji stvarno obara. Cena je isto tako visoka:

- aplikacija danas **ne može da završi lanac posle zbirne** (F4 prijemnica je
  pauzirana do S6, pa nema paleta, GP fakture ni utovara);
- 13 komponenti iz prethodne dve nedelje rada je obrisano (ceo ekran Sledljivost,
  ZBR-IDENT/ZBR-CHILD aparatura, `modOtkupBlok`, `modMarza`…);
- self-heal briše podatke otkupa bez ikakve kapije;
- nijedan od pet P1 nalaza iz pregleda od 04.09 nije dirnut.

Nema potvrđenog bezuslovnog P0. Tri nalaza su P0 **pod uslovom da postoji
instalacija sa podacima** — a ta pretpostavka nije kapija u kodu nego odluka u
dokumentu.

## 2) Faze rada

| Faza | PR | Šta | Verifikacija po PR opisima |
|---|---|---|---|
| Poslednje forme, jedna forma, logo u kodu | #264–#273 | `frmOtkupAPP`, `frmSEF`, `frmLogin`, `frmSplash`, `frmExcelMini`, `frmMarza` obrisani; `modUiFaze`, `modLogo` | #271 (start lanac) NEVERIFIKOVANO |
| Self-update motor SU-1..SU-11, import marker, backup | #274–#287, #312–#316, #329 | `modVbaTools` +2470, 4 nove CI kapije | 9 SU commit-a NEVERIFIKOVANO; meren samo dev put |
| ZBR-IDENT, ZBR-CHILD (GeneracijaID na deci) | #288–#301 | identitet zbirne po generaciji | FULL zeleno, compile potvrđen |
| Šema iz koda, skele, otkup cutover | #302–#308 | `schema.json`, `modSchema`, golden 12, `CreateOtkup_TX` | FULL zeleno, compile NEJASNO |
| Alati, kapije, brojevi, revers | #309–#332 | dev sveska, `test_grana`, autosave/import/backup kapije, ugovor o formatu, brojevi po nizu, `ReversID` | FULL zeleno |
| Mapa sposobnosti, odluke 16.09 | #333–#351 | 387 sposobnosti, nova tabela slajsova | docs |
| Slajsovi S1–S5 | #352–#388 | otkup, banka po ID, otpremnica nacrt/izdavanje/ispravka, zbirna nacrt/radni sto/traka, PWA predaja i VOZ | FULL zeleno, BFP 336 → 1866 tvrdnji; compile potvrđen u #352 i #356 |

Autori: web sesije („Claude") 05–09.09 i 16.09; sve ostalo lokalne Windows sesije.
Nema direktnih commit-a na `main`. Četiri process PR-a diraju samo `.claude/`;
nijedan feature PR ne dira pravila — pravilo ispoštovano.

## 3) Strateška slika

- **ZBR-IDENT i ZBR-CHILD su građeni pa obrisani u istom opsegu.** #291–#301 su
  uveli `GeneracijaID` na decu zbirne sa desetak testova; S4-2a do S4-3c su to
  uklonili kao „dvostruki identitet". Isto važi za ekran Sledljivost (#246) i
  lanac GP (#248), obrisane 17.09 („vraća S9"). Oko 12.000 linija rada je živelo
  manje od tri nedelje.
- **Lanac je prekinut na prijemnici.** `PrijemnicaValidiraj` vraća pauzu
  bezuslovno do S6 (`modDokUnos.bas:820`), a ostatak validatora stoji kao mrtav
  kod iza `Exit Function`. Pauzirani su i ispravka zbirne, storno tok zbirne
  (S5-3b), ispravka hladnjačkog otkupa i auto-lanac hladnjače
  (`AUTO_PRIJEMNICA_HLADNJACA` default OFF do S6). Golden scenariji A1–G1 su
  obrisani „posle S6"; ostala su dva od dvanaest.
- **Odluka „nema produkcije" nije kapija u kodu.** Sve iz §4 pod P0 važi samo dok
  je ta pretpostavka tačna. `APP_VERSION` je i dalje `2.28.4` (od 19.08), pa
  klijent sa self-update-om ništa ne povlači — slučajna zaštita, ne namerna.
  Release notes idu do `vba-v2.93.0 — u pripremi` i opisuju jednu formu, ne
  refaktor.

## 4) Nalazi

### P0 — potvrđeno u kodu, uslovljeno postojanjem instalacije sa podacima

1. **Self-heal briše podatke otkupa bez ikakve kapije.**
   `modSetup.EnsureSledljivostSchema` (iz `EnsureRuntimeSchema`, svaki start) briše
   iz `tblOtkup` kolone `Isplaceno`, `DatumIsplate`, `Kolicina`, `Cena`,
   `KolAmbalaze`, `Novac`, `PrimalacNovca`, `Klasa`, `VremeUnosa`, `BrutoKg`
   (`modSetup.bas:1296-1310`). `ObrisiKolonuAko` proverava samo da kolona postoji
   (`:1362-1372`). Nema migracije u `tblOtkupStavke`, nema backup-a, nema uslova.
   Zatečena sveska gubi količine i cene svih otkupa na prvom startu; strogi
   čitači (`RequireZaglavljaSaStavkama`) posle toga padaju po imenu — sveska je
   glasno neupotrebljiva. Test `Test_OTK_SelfHealMigracijeKolona` meri samo broj
   kolona. **Minimalna ispravka:** u grani `If t = TBL_OTKUP` preskočiti brisanja
   uz `LogError` kad `tblOtkup` ima redove a `tblOtkupStavke` nema. Ako je
   pretpostavka „nema produkcije" tačna, kapija ne košta ništa.
2. **Klijentski self-update radi sve merge-ove u jednom makrou uprkos izmerenom
   padu.** `modVbaTools.bas:60-68` beleži merenje od 14.09: posle 5
   `DeleteLines` + `AddFromString` u jednom makrou Excel pada u VBE7.DLL.
   Popravka (jedan merge po OnTime tiku, `ImportAllVBA_MergeStep`) postoji samo u
   dev alatu. Zamrznuti `modSelfUpdate.ImportFromFolder` (`SKIP_MODULES`) i dalje
   ide kroz sve fajlove u jednoj petlji (`modSelfUpdate.bas:983-1095`). Svaki
   release sa više od 5 izmenjenih modula može oboriti Excel usred faze 1.
   `docs/SELF_UPDATE.md` ne pominje ovo merenje.
3. **Self-update nikad ne uklanja komponente kojih više nema u release-u.** Od
   04.09 obrisano je 13 komponenti (`clsBlokUI`, `clsWheelList`, 6 formi,
   `modKarticaDetalji`, `modMarza`, `modMouseWheel`, `modScrSledljivost`,
   `modSledljivost`). Klijent koji primi nov kod ih zadržava, a one referenciraju
   simbole kojih više nema (`modOtkupBlok` 8, `modKarticaDetalji` 7). Stari
   zamrznuti updater ih zove direktno (`modSelfUpdate@34aeedc:366-370`), pa je to
   compile pad pri startu sa sakrivenim Excelom. Brisanje viškova postoji samo u
   dev `ImportAllVBA` (#274, `modVbaTools.bas:1080-1145`), a
   `docs/UI_MIGRACIJA_KATALOG.md` §27.4 i `.claude/rules/otkup-i-dokumenta.md:96`
   i dalje tvrde da ni on ne briše.

### P1 — potvrđeno

4. **PWA red sa vozačem obara ceo ciklus sinhronizacije.** PWA i GAS ne šalju
   `PredajaID`; `modMasterSync.bas:1448` red sa `VozacID` bez `PredajaID`
   proglašava konfliktom → fatal flag (`:2449`) → `ImportOtkupFromPWA_Core =
   False` → izvoz ka PWA preskočen. Otkupac dobija grešku za već uvezen zapis,
   operater vidi generičan FAIL. Docs beleže zahtev prema PWA; kod ga ne kaže
   glasno nego ruši ciklus. Ispravka: degradacija sa porukom, ne fatal.
5. **Red koji desktop odbije otkupac vidi kao sinhronizovan.** GAS terminalne
   statuse uključuje `SyncError` (`gas/Code.gs:1543-1550`) i na postojeći red
   vraća `status: 'existing'` (`:1658-1660`). PWA `isSuccessStatus` tretira
   `existing` kao uspeh i upisuje `synced` (`src/js/utils/sync-engine.js:59-69`,
   `:133-141`); polje `terminal` niko ne čita. Mehanizam je zatečen, ali ga
   refaktor čini čestim (kolizija broja, predaja bez `PredajaID`, kultura).
6. **PWA numeracija ne vidi desktop unose.** Izvoz OtkupiAll upisuje broj
   dokumenta pod „Napomena" (`modStammdatenSync.bas:900`), a PWA
   `generateBrojDokumenta` čita `r.BrojDokumenta` (`otkup-form.js:568`). Isti broj
   istog dana na istoj stanici završava kao `SyncError` kroz
   `RequireBrojJedinstven` — pa po t.5 nestaje. Zatečeno (postoji i na 34aeedc),
   bez testa.
7. **Prvi izvor sa nepraznim poljem definiše činjenice zbirne, ne prvi izvor.**
   `ZbrPreuzmiPolje` puni zaglavlje samo kad je polje prazno a izvor ga ima
   (`modDokumenta.bas:2148-2164`); `ZbrRequireIstiAko` preskače proveru dok je
   zaglavlje prazno (`:2103-2118`); čišćenje ide samo na nula članova
   (`:2227-2244`). Izvor bez sorte pa izvor sa sortom → zaglavlje nosi sortu
   drugog; uklanjanje drugog je ostavlja, prvi izvor postaje nevaljan pri
   izdavanju. Jednopotezni `CreateZbirna` isti sastav odbija. Nema testa za
   prazno-pa-neprazno.
8. **Prozor između faze 1 i 2 self-update-a nije zaštićen od ručnog snimanja.**
   `ThisWorkbook.Saved = True` postoji samo u abort putu
   (`modSelfUpdate.bas:545`). Sa ugašenim eventima i vidljivim Excelom, zatvaranje
   fajla u tom prozoru nudi „Save changes?" i upisuje projekat bez tvrdih modula.
9. **Hardening self-update-a ne stiže klijentima.** Ceo SU niz menja
   `modSelfUpdate`, koji je u `SKIP_MODULES`; stari `ImportFromFolder` preskače
   svoj modul iz release foldera (`@34aeedc:25`). Klijent dobija novi updater samo
   punom reinstalacijom; `RELEASE_PROCEDURE.md` to pominje samo kao napomenu.
10. **Popustljiv skupni čitač članstva zbirne.** `AktivnoClanstvoPoKanonu`
    (`modDokumenta.bas:~1414-1470`) preskače prazan ID i ne proverava postojanje
    zbirne ni otpremnice, dok isti čitač za otpremnicu
    (`AktivnoOtpClanstvoPoKanonu`) diže 1331–1335. Red članstva ka nepostojećoj
    zbirnoj čini otpremnicu trajno „zauzetom" (nestaje iz NEVEZANE, `Dodaj` i
    `Storno` odbijaju) bez izlaza iz UI. Nema testa.
11. **Spisak testova koje S4 mora da vrati nije zatvoren.** Plan §14.14 nabraja
    14 obrisanih testova i 3 sabotaže kao obavezu S4; S4 je proglašen završenim,
    a nijedno ime nije vraćeno ni eksplicitno prevaziđeno. Tvrdnje „zbir zbirne =
    zbir otpremnica", „vozač je prva provera" i „prosek gajbe bez storniranih"
    nemaju naslednika među 39 novih `Test_ZBR_*`.

### P2 — potvrđeno

12. Izvoz OtkupiAll (`modStammdatenSync.bas:780-789`) nema `Klasa`, `Kolicina`,
    `Cena`, `KolAmbalaze`, a PWA pregled i menadžment ih čitaju
    (`otkup-pregled.js:181-185`). Novi tab `OtkupiAllStavke` nema nijednog čitaoca
    u GAS ni PWA. Svaki master otkup u PWA prikazuje 0 kg i 0 RSD.
13. Poništenje zbirne izvršava fazu B (skidanje paletnih stavki) **posle**
    `CommitTx` (`modStornoFlow.bas:2143-2152`); EH te rutine vraća 0
    (`modPaletniList.bas:2173-2176`); `res("ok")` je bezuslovno True. Pad ostavlja
    palete vezane, operater vidi „paletne stavke: 0", kontekst COMPLETED.
14. F8 za zbirnu i dalje nudi REŠI KASNIJE (`modScrStorno.bas:770`) iako STANJE
    t.36 kaže DUPLI i PONIŠTENJE; grana (`modStornoFlow.bas:411-415`) pravi samo
    PENDING kontekst bez nastavka.
15. `clsTransaction.RollbackTx` nema EH oko `RestoreTable`
    (`clsTransaction.cls:83-91`). Pad jedne tabele preskače ostale i `CleanUp`,
    pa Excel ostaje bez događaja i sa ručnim preračunom; svi pozivaoci ga zovu pod
    `On Error Resume Next` i prijavljuju rollback kao prošao.
16. `CompleteCorrectionContext` vraća False na pad (`modStornoContext.bas:99`),
    a `modOtkup.bas:480` i `modDokUnos.bas:1257` ga zovu kao Sub. Ispravka uspe,
    kontekst ostane PENDING i pojavi se u Oporavku.
17. Adapter otpremnice čita gajbe kroz `CLng` (`modDokUnos.bas:107-112`,
    `:151-152`, `:393-396`), pa decimala nestane pre `RequireCeoBroj`. Isti kvar
    koji je review #375 za zbirnu ocenio kao P1.
18. F8 lista za storno dedupira po broju (`modStornoFlow.bas:1298-1305`,
    `seen(broj)`). Niz zbirne je (vozač, dan), pa je isti broj drugog dana
    legalan, a druga aktivna zbirna je nevidljiva za storno iz UI.
19. Automatski put banke prosleđuje `dozvoliVisakKaoAvans:=True` bez pitanja
    operatera (`modBankaMapiranje.bas:891-896`), suprotno zapisu u STANJE „samo uz
    saglasnost pozivaoca". Bez testa za višak na auto putu.
20. Batch auto-zbirne diže grešku unutar petlje (`modMasterSync.bas:1785`):
    jedna otpremnica bez vozača ostavlja pola zbirnih i obara ciklus u svakom
    sledećem prolazu.
21. Logika „PROSLEDJENO je izdato" na tri mesta sa dve semantike:
    `IzdatoStatusJeIzdato` (`modDokumenta.bas:5346`), inline kopija
    `modOtkup.bas:278`, i `DocIsIssued` (`modDokumentInvariant.bas:210`, prazno =
    izdato, bez produkcionog pozivaoca).
22. Statička kapija je propustila četiri compile pada u dve sesije (#371
    preimenovan parametar, #374 obrisane javne funkcije, #376 `Private` iz drugog
    modula, #388 kvalifikovan poziv) — backlog §15, nije urađeno. Nezavisan skener
    kvalifikovanih poziva na `main`-u ne nalazi nijedan živ kvar.
23. `RunMasterSyncSmokeSuite` je mrtva: fixture ima 22 kolone sa dijakritikom
    „Količina" (`modGoogleSyncSmokeTests.bas:1096`), validator traži ASCII i 23
    kolone (`GS_BROJ_DOKUMENTA = 23`); poruka „Expected=22"
    (`modMasterSync.bas:1998`) je zastarela. Jedini test koji meri pravi
    round-trip ka Sheets ne može da prođe; suite je `default: False`.
24. `KRAJ_REDA` kapija je masinski zavisna: u web sesiji radno stablo ima LF u 96
    fajlova, pa `vba_check` uvek daje nalaze i hook blokira svaki Edit. Isti
    nalaz kao 04.09; u međuvremenu su dva incidenta sa dvostrukim CR (#378, #380)
    dobila kapiju.
25. Nov dupli ključ poruke `OTKUI_MSG_FK_SEF_OPORAVLJENO` sa različitim tekstom
    (`modPoruke.bas:1645` i `:1688`) — `MsgBox` u `modScrFakture.bas:1744` dobija
    tekst toasta. Stari dupli `OTKUI_SCRMK_SUB` je i dalje tu; `PORUKA` provera ne
    hvata duplikate.
26. `.claude/rules/testovi.md` je zastareo: tvrdi da `gen_schema_module` nema
    self-test (a CI ga vrti od #321); ne pominje `vba_import_marker_gate`,
    `popis_citalaca --check`, `schema_diff`, `test_grana`, `dokaz.py`,
    `RunGoldenSuite`, ni pravila `MRTAV_LOG`, `KOPIJA_NIZA`, `ODSECEN`.
27. Pun `dokaz.py` katalog košta 20–30 sati mašinskog vremena (589 sabotaža, po
    jedna suite od 2–3 min svaka); pravilo „pun katalog pred release" se u praksi
    neće izvršiti. Placebo sabotaža `ljuska-rez-bez-potvrde` stoji od 22.08.
28. GAS `getOrCreateSheet` prepisuje header i puni SheetRegistry po svakom zapisu
    (`gas/Code.gs:2946-2966`) — drift u prvih N kolona postaje nevidljiv.

### P3

- `PrintUtovar` ima četiri tiha `Exit Sub`, a pozivalac (`modScrFakture.bas:1338`)
  uvek pokazuje toast uspeha.
- Legacy kaskada po broju u `StornoOtkup_TX` (`modStorno.bas:61-88`, `370-400`)
  je nedostižna (`Otkup.BrojZbirne` niko ne piše); `FreeOtkupBloksInline` i
  `DetachOtpremniceInline` uvek daju prazan skup — S5-3b „16 mesta".
- Zastarele poruke: „pauziran do S4" (`modStorno.bas:244`); zaglavlja `modAdmin` i
  `modPodesavanja` opisuju `frmStammdaten`; `modAutoHladnjaca.bas:66-68` „F3/F4
  pauzirani".
- `PrijemnicaValidiraj` nosi ~180 linija mrtvog koda iza `Exit Function`
  (`modDokUnos.bas:823-1008`) sa identitetom po broju.
- `OtpremnicaValidiraj` piše `brutoKgI/II` koje adapter nikad ne šalje.
- Dve tolerancije za istu jednakost: `0.001` (`modDokumenta.bas:1285-1292`) i
  `0.0001` (`:2368`).
- `Diag_Hladnjaca` (`modOtkupUI.bas:7561`, javna, 60 linija `Debug.Print`) —
  dijagnostika u produkcionom modulu.
- Ugnježdena transakcija u `IspravkaOtkupa_TX` (`CreateCorrectionContext` /
  `CompleteCorrectionContext` otvaraju svoj `BeginTx` unutar spoljnog) —
  funkcionalno bezopasna (unutrašnji `CleanUp` vraća već spoljno stanje, spoljni
  snapshot pokriva `tblStornoVeze`), ali krši obrazac „core unutar tuđe
  transakcije".
- Docs pominju `ZameniOtpremnicaIzvor` i `ZbrIdIliGreska` koji ne postoje
  (stvarno: `IzvadiIzvorIzNacrta`/`UvediIzvorUNacrt`, `ZbrIdPoBroju`).

## 5) Status nalaza iz pregleda 04.09.2026

| Nalaz (04.09) | Stanje 24.09 |
|---|---|
| P1 `KapacitetPakovanja` čita `TezinaKg` (taru) kao kapacitet (`modUtovar.bas:185`) | **neispravljeno** |
| P1 anti-lockout: admin može da deaktivira/degradira sebe (`modMaticniKorisnici`) | **neispravljeno** |
| P1 `PostojiPK` binarno vs `MatRedPoID` `vbTextCompare` | **neispravljeno** |
| P1 korak 5 obara compile zatečene sveske | prevaziđeno: #274 ImportAllVBA briše viškove (samo dev put — v. P0 t.3) |
| P1 `KRAJ_REDA` masinski zavisna | **neispravljeno** (+ kapija za `\r\r\n`) |
| P2 dupli `OTKUI_SCRMK_SUB` | **neispravljeno** + nov duplikat (t.25) |
| P2 `T_FakturaGP_WriterKapijeIStorno` 439 linija | neispravljeno |
| P2 `APP_VERSION 2.28.4`, bez tagova | neispravljeno |
| P2 `Scr_SlTestReset` van `ResetSeamova` | prevaziđeno (modul obrisan) |

## 6) Verifikacija, testovi i proces

- **Statičke kapije na `main`-u:** `vba_check --self-test` 130/130, sabotaža
  `--proveri-sidra` 589 (2 poznata placeba), `who_writes --check` i
  `--check-ownership`, `gen_schema_module --check`, `vba_parity_check`,
  `vba_selfupdate_gates`, `vba_hard_census`, `vba_import_marker_gate`,
  `popis_citalaca --check` — svi exit 0. Jedini crveni je `vba_check` u web
  sesiji zbog KRAJ_REDA (t.24). CI `static.yml` ima osam novih kapija, svaka sa
  self-testom.
- **Verifikacija na Windows-u** je od 09.09 dosledna: FULL prolaz sa brojkama u
  51 PR opisu, BFP raste 336 → 485 → 608 → 679 → 872 → 1075 → 1207 → 1360 → 1792
  → 1866 tvrdnji. Compile eksplicitno potvrđen u 14 PR-ova; 19 kaže „ostaje na
  operateru"; 21 su NEVERIFIKOVANI — gotovo svi PR-ovi web sesija 05–09.09 (#271
  start lanac, #276–#286 self-update motor, #292, #297, #303). Prvi Windows prolaz
  posle toga (#289) našao je četiri crvena testa iz web sesija.
- **Testovi:** `modTest` 181 → 200 (29 obrisano, 48 dodato); BFP 55 → 252
  procedura (30 obrisano, 227 dodato); `modTestStornoCentar` 24 → 22; golden
  12 → 2. Dispatch tabele `RunOne`/`InvokeTest`/`TestName` konzistentne
  200/200/200. Obrasci koje su reviewi hvatali (tvrdnja po poruci, `IIf` koji
  evaluira obe grane, seed koji ćuti) zamenjeni su merenjem stanja.
- **Proces:** nema direktnih commit-a na `main`; pravila idu kroz zasebne PR-ove;
  spoljni review po PR-u sa P1 i NO-GO ishodima (#361, #362 ×4, #363 ×2, #364,
  #367 NO-GO, #370, #372–#375, #381 ×2, #383–#388).

## 7) Provereno i drži

- Svi pisci zbirne idu kroz jednu transakciju sa snimkom zaglavlja, stavki i
  izvora; `RequireZbrDraft` pred svakom mutacijom; izdavanje traži tačnu
  jednakost kilaže i gajbi po klasi uz revalidaciju; dupli izvor, drugi vozač i
  stornirana otpremnica se odbijaju; ZBR-KANON-04 idempotentno.
- GUID identitet (`NewEntityID`) u svim kanonskim piscima; `SchemaReadyOrFail`
  pred svakim pozicijskim upisom (uključujući banku kroz `modNovac`); jedan
  `Build*RowData` po tabeli.
- Strogi čitač stavki otpremnice: 18 pozivalaca, meko ga čitaju samo dva namerno
  meka mesta (`modIntegritet.ZbirOtpremnicaMeko`,
  `modStornoFlow.GetActiveDocumentsForStorno`). `ZbrIdPoBroju` i
  `RequireStornoAllowed` fail-closed i glasni. Faza A poništenja atomska.
- Hladnjački lanac: jedan prekidač, grana pre nacrta, nema dupliranja bloka
  između radnog stola i automatike.
- Uvoz iz PWA idempotentan po `ClientRecordID` (OTK i VOZ); snimci po redu
  potpuni; `ValidatePWAOtkup` odbija sve očekivano; `PrevezaNovacNaOtkup` u
  transakciji ispravke, idempotentan.
- Stari zamrznuti updater ignoriše svoj modul iz release foldera; abort put
  zatvara bez snimanja; retention ne može da obriše jedini validan backup.
- SEF validator ne odbija GP fakture; mapper deli `PrijemnicaID`/`PreradaID`.
- Locale rizik `CDbl`/`CDate` nad vrednostima iz Sheets bez `valueRenderOption`
  (`modGoogleSheets.bas:1220`, `modMasterSync.bas:2795-2818`) je zatečen od
  19.08, nije nov — banka je isti problem rešila typed datumom, PWA put nije.

## 8) Preporuke, po redu

1. Kapija u `EnsureSledljivostSchema` koja odbija brisanje kolona nad sveskom sa
   otkupima bez stavki. Jedan uslov, nula rizika.
2. Zamrznuti release do odluke o klijentima: ili reinstall flote sa novim
   `modSelfUpdate` (lanac merge-ova + brisanje viškova), ili eksplicitan zapis da
   self-update više nije podržan put isporuke.
3. PWA: red sa vozačem bez `PredajaID` kao degradacija; `terminal` +
   `SyncError` u `applyServerResults` kao neuspeh; `BrojDokumenta` kao svoja
   kolona u OtkupiAll; agregati po redu ili GAS join `OtkupiAllStavke`.
4. Zbirna: preuzimanje činjenica samo pri nula članova, stroga jednakost čim
   postoji član; strog skupni čitač članstva po obrascu za otpremnicu.
5. Storno: faza B unutar transakcije ili `MarkCorrectionManual` na manjak;
   ukloniti REŠI KASNIJE za zbirnu; `RollbackTx` sa EH po tabeli i bezuslovnim
   `CleanUp`; proveravati povratnu vrednost `CompleteCorrectionContext`.
6. Zatvoriti spisak §14.14 stavku po stavku (ime naslednika ili razlog
   prevaziđenosti).
7. Vratiti se na pet nalaza iz 04.09 — nijedan nije skup.
8. `vba_check`: provera vidljivosti i kvalifikovanih poziva iz backloga §15;
   `PORUKA` da hvata duplikate; `check_eol` da čita git atribute.
9. Politika za `dokaz.py`: uzorak po rezu plus pun katalog jednom nedeljno u
   pozadini. Osvežiti `testovi.md` katalog kapija.
10. Ograničiti PR na jednu isporučivu celinu; `modOtkupUI` (9.6k linija),
    `modDokumenta` (8.6k), `modTest` (16.7k) i BFP (18k) rastu bez plana podele.

## 9) Neproverljivo iz ove sesije

Excel suite i `Debug → Compile`; živi Google nalog (locale spreadsheet-a, oblik
koji `FORMATTED_VALUE` stvarno vraća); da li `EnsureStornoVezeSchemaCore` ikad
menja kolone posle snimka (okidač za t.15); `getMgmtOtkupiAll` merge stavki;
`NovacIDsZaOtkup` filter storniranih.
