# OtkupApp — operativna pravila (Codebase Guardian)

> Ovaj fajl se učitava na startu SVAKE sesije, pa drži samo ono što važi uvek.
> Detalji po oblastima žive u `.claude/rules/` i čitaju se kad se ta oblast dira
> (tabela u §3). Cilj je isti kao ranije: čuvati postojeću arhitekturu, sprečiti
> dupliranje, svesti izmene na najmanji potreban delta.

**Default stav:** `reuse > new` · `extend > duplicate` · `verify > conclude` ·
`inspect before propose` · `minimal change over idealized redesign`.

---

## 0) Dijagnoza pre zakrpe

**Kad se prijavi bug, prvo se pravi dokaz da postoji, pa tek onda ispravka.**
Redosled nije stvar stila — bez njega se popravlja simptom, a pravilo koje je
prekršeno ostaje prekršeno.

1. **Reprodukuj.** Test u `modTest` koji pada iz tog razloga, ili merenje sa
   priloženim izlazom. Ako se ne može reprodukovati, to je nalaz — reci da nije
   reprodukovano i ne nagađaj ispravku.
2. **Nađi pravilo.** Koja invarijanta je prekršena (`docs/DOMEN/README.md`) i ko
   sve piše taj podatak (`docs/DOMEN/WHO_WRITES.md`). Isto polje često piše više
   modula; zakrpa na jednom mestu ostavlja ostale.
3. **Tek onda ispravka**, pa isti dokaz u oba smera (§5).

U plan-modu plan mora da počne dokazom, ne rešenjem. Predlog ispravke bez
reprodukcije je pretpostavka i tako se prijavljuje.

Izuzetak su očigledne mehaničke greške (tipfeler, nedostajući argument) koje
`vba_check` ionako hvata.

## 1) Pre svake izmene (obavezno)

1. **Reference-first.** Pogledaj izvore istine:
   - `docs/DOMEN/README.md` — šta dokumenti jesu i koje invarijante drže
   - `docs/DOMEN/WHO_WRITES.md` — ko piše koju tabelu (generisano iz koda)
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
4. **Ne zaključuj iz par linija.** Logika je raspoređena kroz module/forme/klase/
   evente — traži pun set relevantnih (`frm*`, `mod*`, `cls*`, `ThisWorkbook`) pre
   procene reuse-a / refaktora.

## 2) Anti-duplication

Ne praviti paralelnu implementaciju za nešto što već postoji; ne uvoditi novi
naming ako naming pattern već postoji; ne praviti novi shared helper ako sličan
postoji; ne uvoditi novi sloj apstrakcije bez jasnog razloga („rule of three").

## 3) Mapa koda + gde su detaljna pravila

Kratka mapa „gde šta živi" — ne praviti paralele:

| Oblast | Gde | Detaljna pravila |
|---|---|---|
| Domen: šta dokumenti jesu, invarijante | `docs/DOMEN/` | `docs/DOMEN/README.md` |
| Ko piše koju tabelu (vlasništvo) | generisano iz `src-vba/` | `docs/DOMEN/WHO_WRITES.md` |
| Tabele / kolone / konstante | `modConfig.bas` (`TBL_*`, `COL_*`) | `.claude/rules/podaci-i-config.md` |
| Pristup podacima | `modDataAccess.bas` (`GetTableData`/`GetColumnIndex`/`UpdateCell`/`AppendRow`/`GetNextID`/`LookupValue`) | ↑ isto |
| Filter/sort/util nad nizovima | `modArrayUtils.bas`, `modHelpers.bas` | ↑ isto |
| Config tabele (SEF / Local / legacy) | `tblSEFConfig`, `tblLocalConfig`, `tblConfig` | ↑ isto |
| Setup / šeme / health-check | `modSetup`, `modAdmin`, `DebugKoloneTabele` | ↑ isto |
| Otkup / dokumenta | `frmOtkup`+`modOtkup`, `frmDokumenta`+`modDokumenta` | `.claude/rules/otkup-i-dokumenta.md` |
| Matični podaci (UI), forme, `.frx`, runtime kontrole | `frmMaticniPodaci`/`frmStammdaten`/`modMaticniLookups`, `clsBlokUI`, `clsUiSink` | `.claude/rules/forme-i-kontrole.md` |
| Agrohemija / magacin, ambalaža, cenovnik | `modAgrohemija`, `modAmbalaza`, `modCenovnik` | `.claude/rules/agrohemija-i-cene.md` |
| Banka — import izvoda i nalozi za isplatu | `modBankaImport`+parseri, `modBankaMapiranje`, `modBankaExportPregled` | `.claude/rules/banka.md` |
| Sync / PWA, Google auth, self-update, release build | `mod*Sync`, `modGoogleAuth`, `modSelfUpdate`, `modRelease`, `modDrive`, `gas/` | `.claude/rules/sync-i-self-update.md` |
| VBA izvor (ASCII, deklaracije, rezervisane reči, duplikati) | ceo `src-vba/` | `.claude/rules/vba-izvor.md` |
| Test suite i verifikacija | `mod*Tests`, `tools/vba_check.py`, `tools/run_vba.py` | `.claude/rules/testovi.md` |
| Git, PR, release procedura | `tools/release.sh`, `docs/RELEASE_*` | `.claude/rules/git-i-release.md` |

**Kako se koristi:** pre nego što diraš neku oblast, pročitaj njen rules fajl.
Fajlovi u `.claude/rules/` imaju `paths:` frontmatter (koja putanja ih aktivira),
pa je i ručno biranje jednoznačno. Ako oblast nema svoj fajl, važi samo ovo ovde.

## 4) VBA / Excel — pravila koja UVEK važe

Puni tekst i primeri: `.claude/rules/vba-izvor.md`. Minimum koji ne smeš zaboraviti:

- **VBA izvor je 100% ASCII i mora ostati ASCII.** Nijedan `š ž č ć đ`, nijedan
  em-dash, nijedan tipografski navodnik u `.bas`/`.cls`/`.frm`/`.doccls`.
  Korisnički tekst sa dijakritikom ide kroz `modPoruke.UpsertPoruke` + `Poruka("KLJUC")`.
- **Modul-level deklaracije** (`Const`, promenljive, `Declare`, `Type`, `Enum`)
  idu u deklaracionu sekciju na vrhu, **pre prve procedure**.
- **Rezervisane reči su case-insensitive** — `Dim eNum As Long` = `Enum` = compile
  error. Za EH koristi `errNum` / `errDesc` / `errSrc`.
- **`.frx` se ne dira kao tekst.** Nove kontrole → runtime (`Controls.Add`).
  **Nove `Private WithEvents` deklaracije u formama su ZABRANJENE.**
- **Šema tabela je izvor istine, ne kod** (schema drift po instalaciji). Pre upisa
  proveri stvarne nazive kolona.

## 5) Verifikacija — definicija gotovog

> **Gotovo = zeleno nad ispravnim i dokazano crveno nad namerno pokvarenim kodom,
> izlaz priložen.** Zadatak se ne preformuliše u uži koji je uspeo. „Nejasno" se
> prijavljuje kao nejasno, nikad kao zeleno.

Suite koja je zelena nad ispravnim kodom, a nije pokazana crvena nad pokvarenim,
ne dokazuje da išta meri — to je u PR #181 bio ishod četiri puta. Kad dodaš ili
menjaš proveru, pokvari kod namerno, pokaži da baš ta provera pukne po imenu, pa
vrati kod i pokaži da je opet zeleno. Bez oba smera nije gotovo.

**Posle svake izmene u `src-vba/` obavezno:**

```bash
python3 tools/vba_check.py        # exit 0 = čisto, 2 = ima nalaza
```

Isti checker se vrti i kao PostToolUse hook (`.claude/hooks/vba-check.sh`) nad
fajlom koji si upravo izmenio, pa nalaz stiže odmah. Ne prijavljuj izmenu kao
gotovu dok nije zelen.

Checker **hvata i tri najčešće compile greške** iz samog izvora, bez Excela:
nepostojeći simbol (`NEDEFINISAN`), pogrešnu arnost (`ARNOST`) i duplu `Public`
definiciju (`DUPLIKAT`). Namerno je usko — samo `.bas` i samo poziv u poziciji
naredbe — jer je lažan nalaz u hook-u gori od propuštenog.

Šta NE hvata: tip-greške, nedeklarisane promenljive, i bilo šta u `.frm`/`.cls`
(tamo se nasleđeni članovi zovu bez kvalifikatora, pa bi lažni nalazi bili
pravilo). Za to treba Excel.

**Ponašanje se ne proverava statički.** Izmena koja se uredno kompajlira, a menja
ponašanje, hvata se samo test suite-om: `python tools/run_vba.py`. To traži
**Windows + Excel + `pywin32`** i zato **ne radi u Linux/macOS sesiji** (Claude
Code na webu). U takvoj sesiji `Stop` hook prolazi tiho — što znači da sesija
može da se završi „zeleno" **bez ijednog testa ponašanja**. Ne čitaj to kao
verifikovano: za VBA izmene koje diraju ponašanje, sesija ide na Windows mašinu,
ili izmena ostaje neverifikovana i tako se prijavljuje.

Detalji, katalog suite-ova i `gate` vs „blind": `.claude/rules/testovi.md`.

Uz to i dalje: balans `Sub`/`Function`/`Select Case`, `git merge-tree` za
konflikte. Forme: izmene su u kodu, `.frm` ide sa svojim `.frx` parom.

## 6) Git / PR / release

Detalji: `.claude/rules/git-i-release.md`. Uvek važi:

- Razvoj na zadatoj feature grani. **Ne praviti PR bez eksplicitnog zahteva.**
- Integracija `main`-a u feature granu = **„Opcija 3"**: fetch → `git merge-tree`
  → rebase lokalno → **pokaži rezultat** → `push --force-with-lease` tek po
  odobrenju. Nikad force-push pre pokazivanja.
- **Na kraju svake izmene koda:** git bash komande za preuzimanje grane
  (`~/Documents/GitHub/otkupapp-pwa` = `ImportAllVBA` folder).
- **Izmena ponašanja nosi test u `modTest`** — ne checklistu. Checklista u chatu
  je samo za ono što se ne može automatizovati: izgled forme, štampa, PDF,
  ponašanje nad pravim podacima. Checklista NIJE zamena za test koji je moguće
  napisati; ako je moguć, piše se.

---

_Detaljnu „Codebase Guardian" doktrinu (reference-first, anti-duplication, format
odgovora po sekcijama) primenjivati i kad nije eksplicitno ponovljena u promptu._
