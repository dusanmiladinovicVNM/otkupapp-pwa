# OtkupApp — operativna pravila

> Učitava se na startu SVAKE sesije, pa drži samo ono što važi uvek. Detalji po
> oblastima žive u `.claude/rules/` (tabela u §4) i čitaju se kad se ta oblast
> dira. Cilj: čuvati postojeću arhitekturu, sprečiti dupliranje, svesti izmenu na
> najmanji potreban delta.

**Default stav:** `reuse > new` · `extend > duplicate` · `verify > conclude` ·
`inspect before propose` · `minimal change over idealized redesign` — uz
`model > kompenzacija`: „minimalno“ se meri preko **svih** zakrpa koje isti
nedostatak modela traži, ne preko jedne (`.claude/rules/model-pre-zakrpe.md`).

**Ljuska korisnika je Git Bash, i samo Git Bash.** Svaka komanda koja mu se da
ide u **bash obliku sa `/`** — nikad PowerShell oblik, nikad `\` u putanji.
Backslash u Git Bash-u **nestaje bez greške**: `python tools\run_vba.py` postane
`toolsrun_vba.py`, a `cd ~\Documents\GitHub\...` postane `~DocumentsGitHub...`.
Poruka o grešci pokazuje na fajl koji ne postoji, ne na pravi uzrok. Važi i za
poziv PowerShell skripte: `powershell -File tools/check_merge.ps1`.

## 1) Pre izmene

1. **Reference-first.** Izvori istine: `docs/DOMEN/README.md` (šta dokumenti jesu
   i koje invarijante drže), `docs/DOMEN/WHO_WRITES.md` (ko piše koju tabelu),
   `docs/ARCHITECTURE_REFERENCE.md`, `docs/UI_MIGRACIJA_KATALOG.md` (prelazak na
   ljusku `frmOtkupUI` — šta je preneto i kako; §27 je **završen**, a od §27.18
   je `frmOtkupUI` **jedina forma** — splash, prijava i „Excel je otvoren" su
   faze istog prozora).
2. **Pretraži postojeće** u `src-vba/` (VBA/Excel) i `src/` (PWA) PRE nego što
   predložiš nov fajl, komponentu, helper, tip, konstantu, validaciju ili sloj.
   Ako ekvivalent postoji — koristi ga ili proširi minimalno. Novo samo uz jasan
   razlog (postojeće objektivno ne podržava scenario).
3. **Ne zaključuj iz par linija.** Logika je raspoređena kroz `frm*`/`mod*`/`cls*`
   i evente — traži pun set relevantnih pre procene reuse-a ili refaktora.
4. Ako nešto nije provereno, **reci da nije provereno**. Ne popunjavaj rupe
   pretpostavkama.

## 2) Bug: prvo dokaz, pa zakrpa

Reprodukuj (test u `modTest` koji pada baš iz tog razloga, ili merenje sa
priloženim izlazom) → nađi prekršenu invarijantu i **sve** koji pišu taj podatak
(`WHO_WRITES.md`; isto polje često piše više modula, pa zakrpa na jednom mestu
ostavlja ostale) → tek onda ispravka.

Ako se ne može reprodukovati, to je nalaz — prijavi da nije reprodukovano i ne
nagađaj ispravku. Izuzetak su očigledne mehaničke greške koje `vba_check` hvata.

## 3) VBA / Excel — uvek važi

Puni tekst i primeri: `.claude/rules/vba-izvor.md`.

- **VBA izvor je 100% ASCII i mora ostati ASCII.** Nijedan `š ž č ć đ`, em-dash
  ni tipografski navodnik u `.bas`/`.cls`/`.frm`/`.doccls`. Korisnički tekst sa
  dijakritikom ide kroz `modPoruke.UpsertPoruke` + `Poruka("KLJUC")`.
- **Modul-level deklaracije** (`Const`, promenljive, `Declare`, `Type`, `Enum`)
  idu u deklaracionu sekciju na vrhu, **pre prve procedure**.
- **Rezervisane reči su case-insensitive** — `Dim eNum As Long` = `Enum` =
  compile error. Za EH koristi `errNum` / `errDesc` / `errSrc`.
- **`.frx` se ne dira kao tekst.** Nove kontrole → runtime (`Controls.Add`).
  **Nove `Private WithEvents` deklaracije u formama su ZABRANJENE.**
- **Šema tabela dolazi iz koda: `schema/schema.json` je kanon.** `modSchema.bas`
  je njegov generisan artefakt (`tools/gen_schema_module.py`), sveska je
  posledica. Menjaš šemu → menjaš `schema.json`, pa regenerišeš. Nikad obrnuto.
- **Redosled kolona je deo šeme, ne kozmetika.** `AppendRow` piše **poziciono**,
  a pisci grade goli `Array(...)` — kolona ubačena u sredinu tiho šalje vrednosti
  u pogrešne kolone. Nove kolone idu **na kraj**.
- `.frm` uvek ide u commit sa svojim `.frx` parom.

## 4) Mapa koda + gde su detaljna pravila

| Oblast | Gde | Detaljna pravila |
|---|---|---|
| Domen: šta dokumenti jesu, invarijante | `docs/DOMEN/` | `docs/DOMEN/README.md` |
| Ko piše koju tabelu | generisano iz `src-vba/` | `docs/DOMEN/WHO_WRITES.md` |
| Šema tabela — kanon, self-heal, provera | `schema/schema.json`, `modSchema.bas` | `.claude/rules/podaci-i-config.md` |
| Tabele / kolone / konstante, pristup podacima | `modConfig.bas`, `modDataAccess.bas` | `.claude/rules/podaci-i-config.md` |
| Arhitektonski ugovor (A1–A12), vlasništvo upisa | `docs/DOMEN/ARCHITECTURE_CONTRACT.md`, `docs/DOMEN/WRITE_OWNERSHIP.json` | isti fajl |
| Otkup / dokumenta | `modScrDokumenti`+`modOtkupUnos`/`modDokUnos`, `modOtkup`+`modDokumenta` | `.claude/rules/otkup-i-dokumenta.md` |
| Forme, `.frx`, runtime kontrole | `frmOtkupUI`, `clsFlatBtn`, `clsUiSink` | `.claude/rules/forme-i-kontrole.md` |
| Agrohemija / ambalaža / cenovnik | `modAgrohemija`, `modAmbalaza`, `modCenovnik` | `.claude/rules/agrohemija-i-cene.md` |
| Banka — izvodi i nalozi | `modBankaImport`+parseri, `modBankaMapiranje` | `.claude/rules/banka.md` |
| Sync / PWA, self-update, release build | `mod*Sync`, `modSelfUpdate`, `modRelease`, `gas/` | `.claude/rules/sync-i-self-update.md` |
| VBA izvor (ASCII, deklaracije, duplikati) | ceo `src-vba/` | `.claude/rules/vba-izvor.md` |
| Testovi i verifikacija | `mod*Tests`, `tools/vba_check.py`, `tools/run_vba.py` | `.claude/rules/testovi.md` |
| Git, PR, release | `tools/release.sh`, `docs/RELEASE_*` | `.claude/rules/git-i-release.md` |
| Kompenzacija: zakrpa umesto ispravke modela | `docs/DOMEN/KOMPENZACIJE.md` | `.claude/rules/model-pre-zakrpe.md` |
| Kapija pre veće izmene (nova tabela/kolona, nov writer, identitet, storno, schema, SEF/banka) | — | skill `pre-flight` |

Fajlovi u `.claude/rules/` imaju `paths:` frontmatter (koja putanja ih aktivira).
Ako oblast nema svoj fajl, važi samo ovo ovde.

## 5) Verifikacija — tri nivoa

| Nivo | Kada | Komanda |
|---|---|---|
| **FAST** | posle svake VBA izmene; ide i automatski kroz `PostToolUse` hook | `python tools/vba_check.py` |
| **TARGETED** | za feature ili bug — suite koja pokriva to područje | `python tools/run_vba.py --suite <ime>` |
| **FULL** | pred release i za rizične izmene u jezgru | `python tools/run_vba.py` |

- `vba_check` radi svuda, i u Linux sesiji. `run_vba` traži **Windows + Excel +
  `pywin32`** i u web sesiji se **ne izvršava** — tamo se izmena ponašanja
  prijavljuje kao **neverifikovana**, nikad kao zelena.
- **Diraš šemu ili upis?** Dodaj i ove dve, obe rade bez Excela:
  `python tools/gen_schema_module.py --check` (kanon i `modSchema` u koraku) i
  `python tools/who_writes.py --check-ownership` (A11 — nov pisač domen-tabele
  obara CI). Pred uvoz u zatečenu svesku: `python tools/schema_diff.py <sveska>`.
- **Compile je ručna kapija pred release:** `Alt+F11 → Debug → Compile
  VBAProject`. Automatski verdikt je često `NEJASNO` i tako se i prijavljuje.
- **Izmena ponašanja nosi test u `modTest`**, ne checklistu. Checklista u chatu je
  samo za ono što se ne može automatizovati: izgled forme, štampa, PDF, ponašanje
  nad pravim podacima.
- **Dokaz u oba smera** (pokvari kod → provera pukne **po imenu** → vrati → opet
  zeleno) obavezan je kad menjaš **sam test ili checker**, i kod naročito kritične
  poslovne invarijante. Razlog: zelena suite koja nikad nije pokazana crvena ne
  dokazuje da išta meri. Za običnu funkcionalnu izmenu se ne traži.
- „Nejasno" se prijavljuje kao nejasno. Zadatak se ne preformuliše u uži koji je
  uspeo.

Detalji i katalog: `.claude/rules/testovi.md`. Zašto su pravila ovakva:
`docs/engineering/postmortems/2026-08-verifikacija.md`.

## 6) Git / PR

Detalji: `.claude/rules/git-i-release.md`.

- Razvoj na zadatoj feature grani. **Ne praviti PR bez eksplicitnog zahteva.**
- Integracija `main`-a: `git fetch` → `powershell -File tools/check_merge.ps1` →
  rebase lokalno → **pokaži rezultat** → `push --force-with-lease` tek po
  eksplicitnom odobrenju. Nikad force-push pre pokazivanja.
- **Izmene u `.claude/` idu isključivo kroz zaseban process PR, jedan po jedan** —
  nikad zajedno sa feature izmenom. Paralelne sesije nad istim `.claude/` su već
  proizvele tri sudarena PR-a.
- Komande zovi iz root-a repoa; `cd ... &&` prefiks obara permission match.
- Na kraju izmene koda daj komande za preuzimanje grane (**bash oblik**;
  `~/Documents/GitHub/otkupapp-pwa` = `ImportAllVBA` folder).
