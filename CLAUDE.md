# OtkupApp — operativna pravila (Codebase Guardian)

> Ova pravila Claude Code učitava na startu svake sesije. Cilj: čuvati postojeću
> arhitekturu, sprečiti dupliranje, svesti izmene na najmanji potreban delta.

**Default stav:** `reuse > new` · `extend > duplicate` · `verify > conclude` ·
`inspect before propose` · `minimal change over idealized redesign`.

---

## 1) Pre svake izmene (obavezno)

1. **Reference-first.** Pogledaj izvore istine:
   - `docs/ARCHITECTURE_REFERENCE.md`, `docs/ARCHITECTURE_CHANGELOG.md`
   - `instructions/AGRIX_ARCHITECTURE_REFERENCE_FILLED_v6_12_DRAFT.md`
   - `instructions/DOMAIN_MODELS_REVIEW_DRAFT_v6_21_WITH_AGROHEMIJA.md`
2. **Pretraži postojeće** u `src-vba/` (VBA/Excel app) i `src/` (PWA) PRE nego što
   predložiš novi fajl / komponentu / hook / service / helper / tip / konstantu /
   validaciju / state / API sloj.
3. Ako ekvivalent (ili delimičan ekvivalent) postoji — **koristi ga ili proširi
   minimalno**. Novo uvodi SAMO uz jasan razlog (postojeće objektivno ne podržava
   scenario; proširenje bi napravilo veći tehnički dug). Ako nešto nije provereno,
   reci da nije provereno — ne popunjavaj rupe pretpostavkama.

## 2) Anti-duplication

Ne praviti paralelnu implementaciju za nešto što već postoji; ne uvoditi novi
naming ako naming pattern već postoji; ne praviti novi shared helper ako sličan
postoji; ne uvoditi novi sloj apstrakcije bez jasnog razloga („rule of three").

## 3) Mapa koda (gde šta živi — ne praviti paralele)

| Oblast | Gde |
|---|---|
| Tabele / kolone / konstante | `modConfig.bas` (`TBL_*`, `COL_*`) |
| Pristup podacima | `modDataAccess.bas` (`GetTableData/GetColumnIndex/UpdateCell/AppendRow/GetNextID/LookupValue`) |
| Filter/sort/util nad nizovima | `modArrayUtils.bas` (`FilterArray`, `SortArray`), `modHelpers.bas` (`Nz/NzToText/ExcludeStornirano/FillCmb`) |
| Maticni podaci (UI) | `frmMaticniPodaci` + `frmStammdaten` (`Select Case Me.Tag`) + `modMaticniLookups` (data-driven meni) |
| Otkup / dokumenta | `frmOtkup`+`modOtkup`, `frmDokumenta`+`modDokumenta` |
| Agrohemija / magacin | `frmAgrohemija`+`modAgrohemija` (`SaveMagacin` ledger `MAG_ULAZ`/`MAG_IZLAZ`, `GetMagacinStanje`, `GetAgrohemijaDug`). Izlaz **opciono bez parcele** kad je `PRACENJE_PARCELA` OFF (`IsPracenjeParcela`, isti flag kao `frmOtkup`; smart-doza se preskače). **Početni dug kooperanta (migracija)** = rezervisani virtuelni artikal `ART_POCETNI_DUG` (`modConfig`) + `BookPocetniDug` → `SaveMagacin(... allowNoStock:=True)`; artikal je izuzet iz combo-lista i iz `GetMagacinStanje` (NE dirati to izuzimanje). **PWA `ExportMagacinKoop` ga još NE izuzima** (KI-006). |
| Ambalaza ledger | `modAmbalaza` · **Cenovnik (append-only):** `modCenovnik` (`GetVazecaCena/AddCena`) |
| Cena — DVA modela (ne mešati) | single-current po artiklu = `tblArtikli.CenaPoJedinici` (inline `LookupValue`, agrohemija); append-only istorija za otkup voća = `tblCenovnik` |
| Dinamičke kontrole (bez `.frx`) | `Controls.Add` + WithEvents klasa (`clsBlokUI`/`modOtkupBlok`, `clsLookupMenuBtn`/`modMaticniLookups`) · za form-hostovane runtime kontrole **generički `clsUiSink`** (`WireSink` + `UiSinkEvent` dispatcher; vidi `frmDokumenta`/`frmOtkupAPP`). **NOVE `Private WithEvents` deklaracije u FORMAMA su ZABRANJENE** — dodavanje event-sink deklaracije u formu lomi self-update code-merge te forme (krivac za crash 2.16.1→2.21.0; `docs/SELF_UPDATE.md` zamka #11). Zatečeni pre-2.16.1 form-WithEvents (`frmDokumenta` m_tgl*/storno-pregled/recovery dugmad, `frmAgrohemija` „Pocetni dug", `frmPalete`, `frmBankaExportPregled`, `frmIzvestaj`) su **zamrznuti** — ne diraj i ne dodaji nove |
| Banka import (izvodi) | `modBankaImport` (pull+`ImportBankaInbox_TX`), **multi-bank dispatch** `DetectBank`+`Select Case` u `ParseBankaIzvodForImport` (deljeni 4-nivo integritet+17-kol staging; parser po banci — `modBankaImportParserPdfToText`=Komercijalna, `modBankaProCredit`, `modBankaHalk`, `modBankaAlta` (`190-`, fingerprint naslov „IZVOD BR."); svi preko `pdftotext`/Poppler), mapiranje `modBankaMapiranje`→`tblNovac`, forma `frmBankaImport` (**jaki ključevi** — `poziv na broj`=otkup/faktura, `tekući račun` — od v2.38.0/RF-09 iza dugmeta „Mapiraj jake ključeve (N)", NE na `_Activate`; `_Activate` samo prebroji preko read-only `CountStrongKeyReadyBankaImport`). Dedupe ključ uključuje **broj računa**; `Map*` imaju **smer guard** (`RequireBimSmer`); blok sa 3+ otvorenih stavki diže `ERR_BMAP_MANUAL_REQUIRED` koju batch guta **po redu** (`AutoMapBankaImportRowBatch`), ne obara ceo `AutoMapAll`. Datumi izvoda se validiraju pre staging-a (`TryParseDateValue` round-trip, AUD-007). Testovi: `RunBankaImportTestSuite` (`modTestBanka`). GAS `gas/bank-pdf-downloader/` (Gmail→Drive). Runbook: `docs/production-runbook-banka-import-setup.md`; novi parser: `docs/development-banka-parser.md`. |
| Banka nalozi (isplate) | `frmBankaExportPregled`+`modBankaExportPregled` (pregled otvorenih blokova, per-blok „Isplatiti“ override; runtime combo — `.frx` se ne dira — „Kooperant“ filter radi i na unos i kao dd, substring, prune override-a protiv PUNE liste; „Sa računa“ = izbor računa firme (do 4 zasebna polja `BANKA_NALOG_RACUN_1..4` u Podešavanjima, spojena kroz `BankaNalogRacuniCSV`; legacy `;`-lista `BANKA_NALOG_RACUNI` + `SELLER_ACCOUNT` kao fallback), prikaz banke `BankaNazivZaRacun`). **CSV nalozi za prenos:** `GenerisiNalogeCSV` → `Nalozi za banku\` (platilac `SELLER_NAME`/`SELLER_ACCOUNT`; **poziv na broj = broj bloka** → auto-map pri uvozu izvoda; šifra/svrha `BANKA_NALOG_*`, Podešavanja grupa „Banka / nalozi“). **PDF specifikacija isplata:** `PrintIsplataSpecifikacija` → `modPrint.FillIsplataSpecSablon` (`ISPLATA_SPEC_PRINT_MODE`, default PDF). **Vezivanje virman avansa:** dugmad „Primeni avans na blok"/„(sel.)" → postojeći `ApplyAvansToOtkup_TX` (dotad samo auto pri snimanju novog otkupa u `modOtkup`; sad i za već otvorene blokove). **Bez upisa u `tblNovac` za isplate** — isplata se knjiži tek uvozom izvoda (avans upis je zaseban, veže `OtkupID` na postojeći `NOV_VIRMAN_AVANS_KOOP`). **Saldo je fail-closed (v2.39.0/RF-10, AUD-026):** override preživljava reload ali se pri svakom rebuild-u usklađuje sa otvorenim (`ClampOverridesToOpen` — nestao/zatvoren blok → briše, veći → spušta, **manji ostaje**), a `GenerisiNalogeCSV` pred upis čita **svež** saldo i kroz `ValidateNalogSaldo` odbija CEO fajl kad ijedan nalog prelazi otvoreno (razlog kroz `outOdbijeno`; blok kog nema među otvorenima = otvoreno 0). Iznosi se porede u **cent-domenu** (`ZaokruziNovac`, half-up), **bez epsilon tolerancije** — prag `+ 0.01` je propuštao preplatu od punog centa; isto pravilo za unos u formi i `CsvIznos`. NE uvoditi novu granicu ni nov helper. Avansi se broje po **stvarno proknjiženom iznosu** (`ApplyAvansToOtkup_TX` `ByRef`, RF-02) — `True` sam po sebi ne znači da je nešto vezano. |
| Sync / PWA | `modStammdatenSync`, `modMasterSync`, `gas/` · Google/PWA kredencijali žive u **`tblSEFConfig`** (`GetConfigValue`), auth `modGoogleAuth.RunGoogleAuthSetup` |
| Self-update (kod) | klijent `modSelfUpdate` (`CheckForUpdateOnOpen`/`RunSelfUpdate` dvofazni) · build `modRelease.PublishReleaseToDrive` · Drive REST `modDrive` · **vidi `docs/SELF_UPDATE.md` (zamke!)**. **`modSelfUpdate` je u `SKIP_MODULES` (frozen) → updatable moduli (`modMain`…) NE smeju early-bind-ovati NOV `modSelfUpdate` simbol** (star klijent posle self-update-a = nov `modMain` + star `modSelfUpdate` → `Compile error: Sub or Function not defined` obori `StartApp`; zamka #19). Nov cross-poziv sakrij iza postojećeg stabilnog simbola ili late-bound (`Application.Run`). |
| Setup / šeme | `modSetup` (`SetupNewPC`, `Ensure*Schema`; `SetupPopplerInteractive`/`SetupBankFoldersInteractive` pickeri; `RunSetupHealthCheck` uklj. živi `CheckServerLink`/`TestServerLink`), first-run kapija u `StartApp` (nudi `SetupNewPC` dok `APP_SETUP_COMPLETED!=DA`), Admin dugmad `modAdmin` (health/googleauth/ensure), dijagnostika `DebugKoloneTabele` |

## 4) VBA / Excel — specifična pravila (naučeno)

- **Ne zaključuj iz par linija.** Logika je raspoređena kroz module/forme/klase/
  evente — traži pun set relevantnih (`frm*`, `mod*`, `cls*`, `ThisWorkbook`) pre
  procene reuse-a / refaktora.
- **Šema tabela je izvor istine, ne kod.** Realne kolone se razlikuju po
  instalaciji (schema drift). PRE upisa proveri stvarne nazive kolona
  (`Alt+F8 → DebugKoloneTabele`). Primeri naučeni:
  - `tblStanice`: telefon je u koloni `Kontakt` (NE `Telefon`); kontakt = `Ime`/`Prezime`/`PIN`.
  - `tblKulture`: `KulturaID | VrstaVoca | SortaVoca | GajbicaPoPaleti` (NEMA `Aktivan`).
  - `tblOtkup/Otpremnica/Prijemnica/FakturaStavke`: količina je ASCII `Kolicina` (NE `Količina`); koristi `COL_*_KOLICINA`, ne hardkoduj dijakritiku (bio `RunProductionHealthCheck` bug).
- **TRI config tabele — ČITANJE i UPIS moraju u ISTU tabelu** (inače polje „ne radi"):
  `tblSEFConfig` (poslovni + **Google/PWA + SEF**; `Get/SetConfigValue`), `tblLocalConfig`
  (per-mašina: `PDFTOTEXT_EXE_PATH`, `BANKA_*_PATH`, `APP_SETUP_COMPLETED`; `Get/SetLocalConfigValue`),
  `tblConfig` (**legacy**, ne koristi se). Podešavanja editor rutira po `store` ("sef"/"local")
  u `CfgAdd`; path polja imaju inline „…" browse dugme. Naučene greške: poppler upisan u
  SEFConfig a čitan iz Local; Google/`APP_SETUP_COMPLETED` čitani iz pogrešne tabele.
  `GetLocalConfigValue` na **praznu** vrednost vraća **default** (pa prazan `PDFTOTEXT_EXE_PATH`
  = auto `<xlsm>\Tools\poppler\Library\bin\pdftotext.exe`).
- **Pozicijski `AppendRow` zavisi od redosleda kolona** — bezbedan samo ako je
  redosled potvrđen. Za polja čiji redosled nije siguran koristi upis **po imenu**
  (`UpdateCell`/`GetColumnIndex`).
- **Kontrole formi su u binarnom `.frx`** (`.frm` ima samo `OleObjectBlob`). Nove
  kontrole se NE dodaju editovanjem teksta — dodaj ih u runtime-u
  (`Controls.Add` + WithEvents), kao postojeći `modOtkupBlok`/`clsBlokUI`.
- Pri merge-u: čist git-merge može dati VBA compile grešku (dupli `Public`
  `Sub`/`Function`/`Const` → „Ambiguous name"). Posle merge-a uradi
  `Debug → Compile VBAProject` i proveri duple definicije.
- **Modul-level deklaracije IDU U DEKLARACIONU SEKCIJU** (vrh modula, posle
  `Option Explicit`, **pre prve procedure**): `Public`/`Private Const`,
  `Public`/`Private` promenljive, `Declare`, `Type`, `Enum`. VBA **ne kompajlira**
  `Const` ubačen između dve procedure — a to je prirodno mesto na koje padne kad
  se konstanta piše „uz funkciju koja je koristi" (RF-07: `IZV_TAB_*` stavljene
  iznad `IzvestajTabDostupan`, na sredini `modIzvestaj`). Grep pre commita:
  `Public|Private Const` posle prve `Sub`/`Function` linije = greška. Konstante
  grupiši uz postojeće na vrhu i objasni ih komentarom tamo, ne kod korisnika.
- **Rezervisane reči — VBA je case-insensitive.** Ime promenljive koje se
  case-insensitive poklapa sa ključnom reči obara compile, i kad se razlikuje po
  velikim slovima: `Dim eNum As Long` = `Enum` → greška (RF-06). Isto važi za
  `type`, `error`, `name`, `line`, `date`, `len`, `input`, `print`, `set`, `get`,
  `event`, `property`, `option`, `base`, `text`, `time`, `mid`, `local`, `read`…
  Za EH varijable koristi postojeću konvenciju projekta: **`errNum` / `errDesc` /
  `errSrc`** (`modStorno.LogAndReraise`, `modAgrohemija`, `modBankaImport`), ne
  izmišljaj `eNum`/`eSrc`. Grep pre commita nad novim `Dim`/`Const`/`ByVal`
  imenima — CI ne kompajlira VBA, pa ovo hvata tek operater u VBE-u.
- **Encoding (posle lokalizacije — `1jj9xw` / v2.6): VBA izvori (`.bas`/`.cls`/`.frm`/`.doccls`) su sada 100% ASCII i MORAJU ostati ASCII.** Sva dijakritika je izmeštena u runtime katalog (`modPoruke` → `Poruka("KLJUC")`, tekst se gradi sa `ChrW`), pa u izvoru nema ne-ASCII bajtova koje bi Edit/Write iskvario.
  - Pošto su fajlovi ASCII, **Edit/Write je sada bezbedan** na `.bas`/`.cls`/`.frm` (latin-1 round-trip više nije potreban). `file <fajl>` treba da kaže „ASCII text“.
  - **NIKAD ne upisuj ne-ASCII znak direktno u VBA izvor** — ni `š ž č ć đ Š Ž Č Ć Đ`, ni nemačke `ä ö ü ß`, ni tipografske `— « » • „ “`. Time fajl ponovo postaje UTF-8/mešan, a `ImportAllVBA` ga učita kao smeće (ista klasa greške kao `f08a0ee`).
  - **Korisnički tekst sa dijakritikom ide ISKLJUČIVO kroz katalog:** dodaj red u `modPoruke.UpsertPoruke` (`UpsertRow lo, existing, "KLJUC", "Gre" & ChrW(353) & "ka..."`), a na mestu prikaza koristi `Poruka("KLJUC")`. Dijakritika nastaje tek u runtime-u.
  - **NE „sređuj radi čitljivosti“** vraćanjem `ChrW` u literal (`"Gre" & ChrW(353) & "ka"` → `"Greška"`) — to je tačno reintrodukcija greške.
  - ChrW kodovi: `š=353 Š=352 ž=382 Ž=381 č=269 Č=268 ć=263 Ć=262 đ=273 Đ=272` · em-dash `—=8212` · `«=171 »=187 •=8226`. Interne/nemačke (dev) stringove transliteruj u ASCII (`ü→ue ö→oe ä→ae ß→ss`), ne u ChrW.
  - **`.frx` ostaje binarni Windows-1250 — i dalje se NE dira kao tekst** (nepromenjeno). Caption koji živi samo u `.frx` menja se u dizajneru forme.
  - **Verifikacija posle SVAKE VBA izmene:** `file` = „ASCII text“ (NE „UTF-8“, NE „Non-ISO extended-ASCII“); grep ne-ASCII nad `src-vba/*.bas src-vba/*.frm src-vba/*.cls` = prazno; svaki novi `Poruka("KLJUC")` ima par u `UpsertPoruke` (0 orphan-a); posle import-a `Alt+F8 → EnsurePoruke`.
  - Prelazno: ako `file` na nekom VBA fajlu i dalje kaže „Non-ISO extended-ASCII“ (nije transliterovan), za njega važi STARO pravilo (latin-1 round-trip, vidi git istoriju ovog odeljka) dok se ne prebaci na ASCII.
  - `.md` / `.js` / ostali UTF-8 fajlovi: Edit/Write je bezbedan (nema konverzije).

## 5) Verifikacija (CI ne pokreće Excel)

- VBA se ne kompajlira/pokreće u ovom okruženju. Verifikuj **statički**: balans
  `Sub`/`Function`/`Select Case`, nema duplih `Public` definicija, **nema
  modul-level deklaracije (`Const`/promenljive/`Declare`/`Type`/`Enum`) posle
  prve procedure** (§4 — VBA to ne kompajlira), `git merge-tree` za konflikte.
  Finalni smoke-test radi korisnik u Excelu.
- Forme: izmene su u kodu; `.frx` se ne dira. Pri re-importu, `.frm` ide sa svojim
  `.frx` parom.

## 6) Git / PR

- Razvoj na zadatoj feature grani; commit poruke jasne i opisne. Ne praviti PR bez
  eksplicitnog zahteva. Pre merge-a u `main` proveriti konflikte (`git merge-tree`)
  i preklapanja fajlova.
- **Integracija ažuriranog `main`-a u feature granu = UVEK „Opcija 3":** `fetch` →
  proveri preklapanja + `git merge-tree` → **rebase lokalno** na `origin/main` →
  pokaži rezultat (log, diff vs `main`, statičke provere) → `push --force-with-lease`
  TEK po eksplicitnom odobrenju. Nikad force-push pre pokazivanja.
- **Posle kreiranja PR-a ka `main`:** podseti korisnika na release/verzionisanje —
  `tools/release.sh <verzija>` → Excel `ImportAllVBA` → `Compile` → snimi → ship →
  `Fleet` provera, da se novi `AgriX_OtkupApp.xlsm` pravilno verzioniše. Vidi
  `docs/RELEASE_PROCEDURE.md` i dopuni `docs/RELEASE_NOTES.md`.
- **Na kraju SVAKE izmene koda (posle commit/push):** UVEK daj git bash komandu za
  preuzimanje feature grane radi testa kroz `ImportAllVBA`. Lokalni klon je
  `~/Documents/GitHub/otkupapp-pwa` (= `ImportAllVBA` folder). Šablon:
  `cd ~/Documents/GitHub/otkupapp-pwa` → `git fetch origin <grana>` →
  `git checkout <grana>` → `git pull --ff-only origin <grana>`; zatim u Excelu
  `Alt+F8 → ImportAllVBA → Debug → Compile → snimi → test`.
- **Uz to, svaki rad završi kratkom test-checklistom direktno u chatu:** numerisani,
  konkretni koraci šta operater proba u Excelu (klik po klik + očekivani rezultat),
  fokusiran na ono što je u toj izmeni dodato/promenjeno. Kratko i praktično.

---

_Detaljnu „Codebase Guardian" doktrinu (reference-first, anti-duplication, format
odgovora po sekcijama) primenjivati i kad nije eksplicitno ponovljena u promptu._
