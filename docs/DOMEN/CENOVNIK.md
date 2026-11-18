# Cenovnik otkupa — red je odluka firme o ceni

> **Status (18.11.2026):** jezgro (`modCenovnik`, šema, Ensure, poruke) i ekran
> (matični podaci, forma otkupa) su u kodu na grani
> `claude/elegant-mccarthy-jqlkk6` (od `f154530`: `7d5e7f5`, `1338d9b`). Testovi
> 366–376 (`T_Cen_*`, §9) su imenovani unapred. **Ništa nije izvršeno u Excelu:**
> `run_vba` i Compile nisu pušteni (Linux sesija), pa je svako ponašanje ovde
> **neverifikovano**. Izvršena je samo statička kapija `python3 tools/vba_check.py`.

## 1) Svrha i izvor

Ovaj dokument je autoritet za cenovnik otkupa voća u VBA svesci (`tblCenovnik`,
`modCenovnik`). Ako se razilazi sa kodom, kod je u pravu i ovaj fajl je bug
(`docs/DOMEN/README.md` §5).

Model je preuzet iz AgriXSystems (`agrixsystems`, `main` = `41c35b6`):

- `docs/adr/0019-f1-pakovanje-cenovnik-teren.md:28-38` — red je globalna ili
  posebna cena stanice, od dana; dopuna 11.10.2026: rezerva i cena sa PDV-om;
- `docs/adr/0027-f1-cenovnik-i-storno-maticnog-reda.md:15-49` — red
  samo-dodavanje, storno, ispravka, pravilo važeće cene, granica datuma;
- `packages/domain/src/maticni/cena.ts` — `pobednik:148`, `vazecaCena:169`,
  `stornirajCenu:93`, `ispraviCenu:101`, `vaziOdUGranici:82`;
- `packages/domain/test/maticni/cena.test.ts:79-137` — tabela slučajeva C1–C19;
- `supabase/migrations/20261011100713_maticni_cenovnik.sql` — ograničenja;
- `docs/STANJE.md:12008-12660` — pre-flight F1.8, koji je čitao stari VBA sistem.

Ovo je **prevod** na VBA, ne kopija. VBA nema bazu sa ograničenjima, tenant,
uloge po stanici ni `KulturaID`. Šta je drugačije i zašto, piše §7.

Šta se menja prema starom VBA modelu (`GetVazecaCena`, više ne postoji): stara
važeća cena nije imala datum ni stanicu. Pobeđivao je najnoviji red tog ključa
(vrsta, sorta, klasa), i budući; istog dana pobeđivao je kasniji red tabele; bez
pogotka cena je bila 0 (pročitano u AgriXSystems `docs/STANJE.md:12039-12044`, na
`otkupapp-pwa@ede1ff1`).

## 2) Model reda

Red je odluka firme o otkupnoj ceni u din/kg, **sa PDV-om**, za robu
(`VrstaVoca` + `SortaVoca`) i klasu, od dana `Datum`. Prazan `StanicaID` znači
**globalnu** cenu (cela firma). `StanicaID` znači **posebnu** cenu te stanice.

Kanon je `schema/schema.json:119-141`. Nove kolone 12–17 su **na kraju**, jer
`AppendRow` piše poziciono (`modCenovnik.UpisiCenu:636-649`).

| # | Kolona | Značenje | Ko piše |
|---|---|---|---|
| 1 | `CenaID` | Identitet reda (`CEN-…`, `GetNextID`). Storno i ispravka idu po njemu. | `UpisiCenu:629` |
| 2 | `Datum` | Dan od kog cena važi, bez vremena. Jedini datum koji ulazi u pravilo. | `UpisiCenu:638` |
| 3 | `VrstaVoca` | Vrsta iz `tblKulture`, kao tekst (trim). | `UpisiCenu:639` |
| 4 | `SortaVoca` | Sorta iz `tblKulture`; prazno = kultura bez sorte (DOP-ODL-02). | `UpisiCenu:640` |
| 5 | `Klasa` | `I` ili `II`, kanonski oblik (`KlasaKanon:734`). | `UpisiCenu:641` |
| 6 | `Cena` | din/kg sa PDV-om, > 0, dve decimale (`Round(cena, 2)`). | `UpisiCenu:642` |
| 7 | `CreatedAt` | Trag unosa (`Now`). Ne ulazi u pravilo. | `UpisiCenu:643` |
| 8 | `Stornirano` | Tačno `Da` = storniran; sve ostalo = red učestvuje u pravilu. Isto čitanje kao `modHelpers.ExcludeStornirano`, kroz koje filtrira lista (CEN-ODL-10), i kao stari `GetVazecaCena` (`modCenovnik.JeStorno`; revizija 18.11.2026 je našla da su pravilo i lista ranije čitali različito). | `StornirajRedNaIndeksu:669` |
| 9 | `CreatedBy` | Pečat `AppendRow` (`modDataAccess.StampRowAudit`). | `modDataAccess` |
| 10 | `ModifiedAt` | Pečat; menja ga i storno i upis veze ispravke. | `modDataAccess` |
| 11 | `ModifiedBy` | Isto. | `modDataAccess` |
| 12 | `StanicaID` | Prazno = globalna; ID iz `tblStanice` = posebna. | `UpisiCenu:646` |
| 13 | `StorniranoAt` | Vreme storna (`Now`). Trag, ne poslovni datum. | `StornirajRedNaIndeksu:670` |
| 14 | `StorniranoBy` | Autor storna (`AuditUser()`, isti lanac kao `CreatedBy`). | `StornirajRedNaIndeksu:671` |
| 15 | `StornoRazlog` | Razlog, 1–500 znakova, bez belina na krajevima. | `StornirajRedNaIndeksu:672` |
| 16 | `IspravkaOdID` | Na novom redu ispravke: `CenaID` starog reda. | `UpisiCenu:648` |
| 17 | `ZamenjenSaID` | Na starom redu: `CenaID` novog reda. | `IspraviCenu:455` |

Konstante: `modConfig.bas:481-507` (`COL_CEN_*`, `CEN_DANA_UNAZAD = 31`,
`CEN_DANA_UNAPRED = 366`, `CEN_RAZLOG_MAX = 500`). Veza ispravke koristi
postojeće `COL_TRACE_ISPRAVKA_OD_ID` / `COL_TRACE_ZAMENJEN_SA_ID`. Klase su
`KLASA_I` / `KLASA_II` (`modConfig.bas:951-952`).

Red je u memoriji normalizovan u niz `CEN_R_*` (`modCenovnik.bas:44-52`), koji
grade `CenovnikRedovi:145` i testovi. Red bez čitljivog datuma dobija dan 0
(`DanVrednosti:758`).

## 3) Invarijante i vlasnici

Linije su `src-vba/modCenovnik.bas` ako nije drugačije navedeno. Testovi su
`modTest` indeksi iz §9; nijedan nije pušten.

| ID | Tvrdnja | Vlasnik | Test |
|---|---|---|---|
| **CEN-01** | Red je odluka: ključ je `VrstaVoca` + `SortaVoca` + `Klasa`, nivo je globalna (prazan `StanicaID`) ili posebna, a `Datum` je dan od kog važi. Cena je din/kg sa PDV-om; ništa se na nju ne dodaje. | `UpisiCenu:610` (oblik reda); kanon `schema/schema.json:119-141` | 368 |
| **CEN-02** | Red se posle unosa ne menja. Posle unosa se pišu samo storno kolone (8, 13–15), `ZamenjenSaID` i pečat `Modified*`. Jedini pisac je `modCenovnik` (A11). `MatIzmeni` za cenovnik odbija (`MATU_ERR_CENOVNIK_APPEND`), a ekran nema radnju „Izmeni". Unos, storno i ispravka idu svaki u svojoj transakciji: `AppendRow` doda red pre upisa ćelija i pri grešci ga ne skida, pa neuspeo unos bez snimka ostavlja delimičan red (CEN-REV-01). | `modCenovnik.AddCena` (`clsTransaction`); `docs/DOMEN/WRITE_OWNERSHIP.json:61-67` + `python tools/who_writes.py --check-ownership`; `modMaticniUnos.MatIzmeni`; `modMaticniEkran.Radnje` | 369, 370, 375; `T_MatEkran_RadnjeIRezim` (nema `izmeni`) |
| **CEN-03** | Najviše jedan nestorniran red po (vrsta, sorta, klasa, stanica, dan); poređenje je trim, bez razlike u veličini slova. Pisac duplikat odbija (`CEN_ERR_DUPLIKAT`). Čitalac ga ne razrešava nego diže `ERR_CEN_DUPLI_KLJUC`. Storno oslobađa ključ, i u istoj transakciji (ispravka: `izuzmiID`). | pisac `CenaUnosProblem:282-299`; čitalac `CenPobednik:103-109`; zatečeno `MigrirajDupleCene:482` (CEN-ODL-04) | 367, 368, 371 |
| **CEN-04** | Provere unosa, ovim redom: vrsta obavezna; vrsta (i sorta kad je data) postoji u `tblKulture`; klasa `I`/`II`; cena > 0, ≤ 9999999999.99, najviše dve decimale; stanica (kad je data) postoji u `tblStanice`; datum u poslovnoj godini (1900–2199, `modParse.bas:7-8`); granica (CEN-09); duplikat (CEN-03). Ishod je ključ poruke `CEN_ERR_*`; editor po ključu bira polje za fokus. | `CenaUnosProblem:231`; drugi zid `UpisiCenu:617-618`; prevod u fokus `modMaticniUnos.CenaOdbij` | 368 (provere i postojanje svih ključeva u katalogu); fokus samo za datum ispravke (370), ostala polja **bez testa** |
| **CEN-05** | Važeća cena za (stanica, roba, klasa, dan): kandidati su nestornirani redovi tog ključa sa danom od 1 do dana upita, globalni ili baš te stanice. Pobeđuje najkasniji dan; istog dana posebna pobeđuje globalnu. Prazna stanica upita vidi samo globalne. Nova globalna cena time prekida raniju posebnu. | `CenPobednik:80` | 366, 367 |
| **CEN-06** | Rezerva: kad sorta nema nijedan važeći red za tu stanicu i dan, isto pravilo važi za istu vrstu bez sorte. Sorta pobeđuje i noviju rezervu. Kultura bez sorte nema rezervu. | `CenVazeciRed:126` | 366 (C1, C2, C11, C15, C16, C17) |
| **CEN-07** | Bez reda nema cene — nikad 0. `TryVazecaCena` tada vraća `False`, a forma polje ne dira. Pobednik sa cenom ≤ 0 (zatečen ili nečitljiv broj) znači „nema cene"; stariji red se tada ne traži (`:204-207`). Kvar čitanja (nema tabele ili kolone, prekršen CEN-03) ide u log i znači „nema predloga": predlog ne sme da obori unos dokumenta. | `TryVazecaCena:183`; pozivalac `modOtkupUI.AutoFillCena` | 367, 373 |
| **CEN-08** | Buduća cena ne važi pre svog dana. Predlog na formi se računa za datum dokumenta (`modOtkupUI.DatumIzPolja`; danas dok polje nije čitljivo), ne za danas. Stari bag „buduća cena važi odmah" ne prelazi. | `CenPobednik:93` (`d <= danUpita`); `modOtkupUI.AutoFillCena` | 366 (C10), 373 |
| **CEN-09** | `Datum` novog reda je od danas − 31 do danas + 366, ivice uključene. „Danas" je lokalni `Date` računara. Ispravka koja zadržava stari `Datum` nije pod granicom; kad ispravka menja datum, nov datum jeste. Prazan ili nečitljiv datum u editoru znači danas **samo pri unosu** (PLT-52). Ispravka ga odbija (`CEN_ERR_DATUM`, fokus na datum), kao AgriX (`IspravkaCeneInput.vaziOd` je obavezan): tiho „danas" bi stornirao stari red i preselio cenu na danas, bez povratka za red stariji od granice. | `CenaUnosProblem:274-280`; izuzetak `CenaIspravkaProblem:355-356` i `IspraviCenu:444`; datum `modMaticniUnos.CenaIzPolja` (`datumObavezan`) | 368 (granica), 370 (granica ispravke; prazan i nečitljiv datum ispravke); „prazno = danas" pri unosu **bez testa** |
| **CEN-10** | Cena na stavci dokumenta je snimak (`docs/DOMEN/DOCUMENT_HEADER_LINES.md:271-279`, §4.1d). Cenovnik daje samo predlog u polje forme. Stavka ne pamti red cenovnika. Storno ili ispravka cenovnika ne menja izdat dokument. | struktura: redove `tblCenovnik` čitaju samo `modCenovnik` i lista `modMaticniIzvor`; ostali pomeni `TBL_CENOVNIK` (`modSetup`, `modSchema`, `modSchemaGuard`, `modPregledListova`) su šema i administracija (grep, 18.11.2026); pisac dokumenta traži samo `Cena > 0` | nema posebnog testa (v. §10) |
| **CEN-11** | Storno nosi razlog, autora i vreme zajedno: `Stornirano = Da`, `StorniranoAt = Now`, `StorniranoBy = AuditUser()`, `StornoRazlog`. Razlog je obavezan, 1–500 znakova posle skidanja belina na krajevima (razmak, tab, CR, LF, NBSP). Sve četiri ćelije idu u jednoj transakciji. `StornoRazlog` je pod ugovorom formata (`schema/schema.json` `formats`: tekst), pa razlog koji počinje sa `=` ostaje doslovan tekst, ne formula. | provera `CenaStornoProblem:305` → `RazlogProblem:695`, `TrimBeline:707`; upis `StornirajRedNaIndeksu:660-673` | 369 |
| **CEN-12** | Storno je jednokratan i konačan. Storniran red se ne stornira ni ne ispravlja ponovo (`CEN_ERR_STORNIRANA`). Drugi zid čita stanje ćelije u istoj transakciji, ne iz provere pre nje (drugi zid je **bez testa**: na sva tri puta ga zaklanja provera pre njega). Storno se ne poništava: `modCenovnik` ne piše u žurnal storna i radnje vraćanja nema. Pogrešan storno se popravlja novim unosom. | `CenaStornoProblem:314-317`; drugi zid `StornirajRedNaIndeksu:667` | 369 |
| **CEN-13** | Ispravka je u **jednoj** transakciji: storno starog reda (oslobađa ključ), nov red sa `IspravkaOdID` = stari, pa `ZamenjenSaID` = nov na starom. Nosi pun nov red (vrsta, sorta, klasa, stanica, datum, cena). Ispravka bez ijedne promene se odbija **pre** storna (`CEN_ERR_ISPRAVKA_BEZ_PROMENE`): isti dan, tekst bez razlike u veličini slova, klasa kanonski, cena u parama; razlog se ne računa. Pri grešci se tabela vraća, a `noviID` je prazan. | `IspraviCenu:421-470`; `CenaIspravkaProblem:325-357` | 370 |

**Napomena uz CEN-03.** Čitalac duplikat traži **nezavisno od izbora pobednika**,
nad svim kandidatima upita (nivo i dan): svaki par nestorniranih redova istog
ključa, nivoa i dana `<=` dana upita diže `ERR_CEN_DUPLI_KLJUC`, kojim god redom
stoje u tabeli (`CenPobednik`, rečnik `nivoDan`). Pada i par starijeg dana koji
kasnija cena zaklanja: preduslov ključa koji se čita je prekršen, a ishod mora
biti isti za svaki redosled. Do review-a PR #491 (CEN-REV-02) par se poredio
samo sa trenutnim pobednikom, pa je globalna–posebna–globalna istog dana prolazila
ili padala zavisno od redosleda. Duplikati drugog ključa i druge stanice ne
obaraju čitanje. Celu tabelu proverava samo migracija. Ovo je strože od
AgriXSystems-a, gde duplikat ne može ni da nastane (ograničenje u bazi).

**Napomena uz CEN-02.** Sveska nema grant po koloni kao baza AgriXSystems. Ručna
izmena ćelije na listu ostaje moguća i ništa je ne hvata u trenutku izmene.
Pravilo drže jedini pisac i kapija vlasništva upisa. Admin „Očisti tabele"
(`modPregledListova.OcistiTabele`) briše sve redove `tblCenovnik` zajedno sa
ostalim tabelama; to je zatečeno administrativno pražnjenje, ne pisac reda.

## 4) Tok upisa i čitanja

**Čitanje (predlog na formi).** `modOtkupUI.AutoFillCena` → `TryVazecaCena` →
`CenovnikRedovi` → `CenVazeciRed` → `CenPobednik`. Stanica ulazi u upit samo u
F1 i F2 (`modOtkupUI.CenaPoStanici`; CEN-ODL-08). Polje se puni samo kad je
ishod `True` (CEN-07).

**Unos.** Ekran `MAT_ROBA`, lista `CENOVNIK` → `modMaticniEkran.Sacuvaj` →
`CenaSmeNaUnos` (provera `modMaticniUnos.MatCenaProblem`, pa pregled
`MatCenaPregled`) → `MatDodaj` → `DodajCenu` → `AddCena:366` → `UpisiCenu:610`.

**Ispravka.** Radnja „Ispravi" → `OtvoriIzmenu` sa svim poljima reda →
`SacuvajIspravkuCene`: provera (`MatIspravkaCeneProblem`), razlog (InputBox),
pregled, `MatIspraviCenu` → `IspraviCenu:421`.

**Storno.** Radnja „Storniraj" → `StornirajCenuKlik`: razlog (InputBox) →
`MatStornirajCenu` → `StornirajCenu:385`. Prazan razlog (Cancel ili same beline)
je odustajanje, ne greška (`modMaticniEkran.RazlogPrazan`).

Sva tri pisca ekrana (`MatDodaj`, `MatIspraviCenu`, `MatStornirajCenu`) prvo zovu
`MatBranaUpisa` (CEN-ODL-01). Pisci u `modCenovnik` (`AddCena`, `IspraviCenu`,
`StornirajCenu`) sami ponove sve provere, pa poziv mimo ekrana ne zaobilazi
nijednu.

**Pregled pre upisa nove globalne cene** (`PrekinutePosebneCene:545`,
CEN-ODL-03). Suvo izvršenje **istog** pravila: `pre` su redovi kakvi jesu,
`posle` isti redovi plus nov globalni red, bez reda koji ispravka stornira. Za
svaku stanicu iz `tblStanice` i svaku robu (kad sorta nije data, i svaku sortu te
vrste koja bi prešla na rezervu) pita se da li u `pre` važi posebna, a u `posle`
nov red. Ishod je spisak `Naziv (ID): vrsta sorta / klasa -- 0.00 (od
dd.mm.yyyy)` ili prazno. Posebna cena ne prekida ništa, pa za nju pregleda nema
(`MatCenaPregled`).

**Test seam.** Svako pitanje cenovnika (razlog storna, razlog ispravke, potvrda
prekida) u test režimu se dopisuje u `modMaticniEkran.CenPitanjeTest` i odgovara
se odbijanjem, pa se ništa ne upisuje. Seam samo posmatra i zabranjuje.

## 5) Tabela slučajeva

Test 366 prenosi tabelu C1–C19 iz AgriXSystems 1:1
(`docs/STANJE.md:12459-12493`, `cena.test.ts:79-106`): stanice A i B; Malina bez
sorte, Polka i Miker; redovi r1–r11. Zato se primeri ovde ne prepisuju. Primeri
koje ta tabela drži: C6 (nova globalna prekida raniju posebnu), C7 (istog dana
posebna pobeđuje), C10 (buduća ne važi), C15 (sorta sa posebnom samo na A koristi
rezervu na B), C16 (sorta pobeđuje noviju rezervu), C18 (storno globalne vraća
raniju posebnu).

## 6) Odluke

Odluke vlasnika su od 18.11.2026 (pre-flight u sesiji). Brojevi prate komentare
u kodu: migracija je `CEN-ODL-04` u `modCenovnik.bas:22`, `:66-69` i
`modSetup.bas:994`, `:1019`.

| ID | Odluka |
|---|---|
| **CEN-ODL-01** | *(vlasnik)* Pravo je postojeće `MaticniPodaci` (`OBL_MATICNI`) kroz `modMaticniUnos.MatBranaUpisa`. Nema novog prava `cene:write` (AgriXSystems O179 se ne prenosi). |
| **CEN-ODL-02** | *(vlasnik)* Predlog cene na formi otkupa i dokumenata se osvežava na promenu vrste i sorte i na uključivanje II klase (kao do sada) **i** na promenu datuma i otkupnog mesta (`modOtkupUI.OnDatumChanged`, `OnStanicaChanged`). Računa se iznova i na kraju prefill-a kad dokument ne donosi cenu za neku klasu (nad konačnim datumom i mestom; cena sa dokumenta se zatim vraća i ima poslednju reč) i pri stvarnoj promeni režima (`SelectModeCore`), jer opseg upita zavisi od režima (CEN-ODL-08). Upisuje se samo kad cena postoji (CEN-07). |
| **CEN-ODL-03** | *(vlasnik)* Pre upisa nove **globalne** cene operater vidi spisak stanica čija ranija posebna cena od tog dana prestaje i bira Da/Ne (`MsgBox` sa podrazumevanim „Ne", `modMaticniEkran.PotvrdiPrekidPosebnih`). Važi za unos i za ispravku. `MsgBox` odseca tekst duži od oko 1024 znaka, pa `TekstPrekidaPosebnih` skraćuje spisak na 700 znakova, uz ukupan broj stanica u naslovu i broj neprikazanih; pitanje je uvek vidljivo. |
| **CEN-ODL-04** | *(vlasnik)* Zatečeni duplikati (stari model je dozvoljavao više nestorniranih redova istog ključa i dana i birao kasniji red tabele) rešava migracija: kasniji red tabele ostaje, raniji se storniraju sa razlogom koji imenuje pobednika („Zamenjen kasnijim redom CEN-… istog kljuca i dana (migracija CEN-ODL-04)."). To je isto pravilo koje planira AgriXSystems F6 (`docs/STANJE.md:12639-12640`). Ključ je (vrsta, sorta, klasa, stanica, dan); red bez čitljivog datuma se ne dira. Jedna transakcija; drugi prolaz ne upisuje ništa. Zove je `modSetup.EnsureCenovnikSchema` (`modSetup.bas:1015`), a nju `EnsurePaletniListSchema` i Admin „Ensure" (`modAdmin.AdminEnsureEverything`). `EnsureRuntimeSchema` je namerno **ne** zove: tamo nema backfill-a. **Posledica:** self-heal pri startu (`modMain.InitApp` → `EnsureRuntimeSchema` → `modSchema.EnsureAllTables`, `modSetup.bas:1235`) dopuni nove kolone, ali duplikate ostavi (čitanje koda). Do pokretanja Ensure (Admin → Ensure), kad je dan upita na dan duplikata ili kasniji (napomena uz CEN-03): (1) predlog cene pada za ključ duplikata i za svaku sortu te vrste koja tog dana na toj stanici nema sopstveni važeći red, jer čita rezervu (CEN-06); (2) unos i ispravka **globalne** cene te robe u Matičnim padaju i kad je datum novog reda kasniji, jer pregled (`MatCenaPregled` → `PrekinutePosebneCene`) grešku čitanja namerno ne guta, pa se ništa ne upisuje — posebna cena se upisuje normalno; (3) nova kasnija cena ne vraća predlog, jer stariji par i dalje obara čitanje. Lek je Admin → Ensure ili ručni storno jednog od duplikata; Ensure treba ponoviti i posle `modMigracija.MigrirajPodatkeIzStarog`, koja kopira stari `tblCenovnik` bez migracije. Kad migracija ne uspe, `EnsureCenovnikSchema` diže grešku (`ERR_CEN_MIGRACIJA`, `CEN_ERR_MIGRACIJA`), pa Ensure ne javlja uspeh. |
| **CEN-ODL-05** | *(prevod)* Ključ robe je tekst (`VrstaVoca` + `SortaVoca`), ne `KulturaID` — isti ključ kao doplata (DOP-ODL-02). ⚠ IDENTITY RISK, prihvaćen: preimenovanje vrste ili sorte u `tblKulture` odvaja stare redove cenovnika od robe. Isto preslikavanje teksta u kulturu planira AgriXSystems F6 (`docs/STANJE.md:12637-12638`). |
| **CEN-ODL-06** | *(prevod)* `StanicaID` je strani ključ na `tblStanice` (`COL_STA_ID`). Proverava se pri unosu (`StanicaPostoji:863`); aktivnost stanice se ne gleda ni pri unosu ni pri čitanju. |
| **CEN-ODL-07** | *(prevod)* Red bez čitljivog datuma (ili sa godinom van 1900–2199) ne učestvuje: ni u pravilu, ni u proveri duplikata, ni u migraciji (`DanVrednosti:758`, `JeAktivanSaDanom:793`). |
| **CEN-ODL-08** | *(prevod)* Zbirna (F3) i prijemnica (F4) nemaju `StanicaID` (`schema/schema.json`), pa njihov predlog vidi samo globalne cene. Izbor otkupnog mesta tamo nije podatak dokumenta (`modOtkupUI.CenaPoStanici`). |
| **CEN-ODL-09** | *(prevod)* PWA `Cena<Vrsta>` je van obima. To je zaseban izvor cene na terenu (§8, tačka 3). |
| **CEN-ODL-10** | *(prevod)* Stornirani redovi se ne prikazuju u listi (`modMaticniIzvor.CenovnikPoredak` → `ExcludeStornirano`). Trag (razlog, autor, vreme, veza) ostaje u tabeli. Lista nema „prikaži stornirane". |
| **CEN-ODL-11** | *(prevod)* Storno se ne poništava (CEN-12; AgriXSystems O176). Pogrešan storno se popravlja novim unosom iste cene. |
| **CEN-ODL-12** | *(odstupanje forme, prihvaćeno 18.11.2026)* Okidači datuma i otkupnog mesta osvežavaju **samo cenu** (`AutoFillCena False`). Tip ambalaže zavisi od kulture, koja se tada nije promenila, pa se ne prepisuje; inače bi se pregazio tip koji je operater ručno ispravio. |
| **CEN-ODL-13** | *(odstupanje forme, prihvaćeno 18.11.2026)* Promena otkupnog mesta okida predlog samo u F1 i F2 (`CenaPoStanici`). U F3/F4 mesto ne ulazi u upit, pa bi poziv samo ponovo upisao istu globalnu cenu preko ručno unete. Promena datuma okida predlog u F1–F4. |

**Posledice CEN-ODL-02, koje treba znati (čitanje koda, nisu izmerene):**

- Posle prefill-a (F1 sa otpremnice, ispravka posle storna, nacrt F2) ručna
  promena datuma ili mesta pregazi prepisanu cenu cenom iz cenovnika, kad ona
  postoji. Forma ne pamti poreklo cene.
- Kad za nov datum ili mesto cene nema, u polju ostaje prethodna cena (drugog
  dana ili drugog mesta). To sledi iz „prazan cenovnik ne briše unetu cenu", a
  ista klasa je `AUD-036` (`docs/KNOWN_ISSUES.md:352`).

## 7) Preslikavanje AgriXSystems → VBA

| AgriXSystems | Šta drži | VBA | Razlika |
|---|---|---|---|
| O175 | tabela `maticni.cenovnik` | §2, `schema/schema.json` | Nema `tenant_id`, RLS ni grant. ID je `CEN-…` iz `GetNextID`, ne UUIDv7. Autor je tekst (`AuditUser`), ne FK na članstvo. |
| MAT-18 | red se ne menja (grant samo storno kolone i `zamenjen_sa_id`) | CEN-02 | Nema granta po koloni. Drže ga jedini pisac i `who_writes --check-ownership`; ručna izmena lista ostaje moguća. |
| MAT-19, O185 | cena `numeric(12,2)` > 0, ne `NaN`; klasa I/II; din/kg sa PDV-om | CEN-01, CEN-04 | Cena je `Double`, ne `Novac` u parama. Dve decimale proverava `JeUParama:749` sa tolerancijom 1e-6, a upis je `Round(cena, 2)`. Red mimo pisca niko ne proverava; čitalac cenu ≤ 0 tumači kao „nema cene" (CEN-07). |
| MAT-20 | jedan nestorniran red po (kultura, klasa, nivo, dan), jedinstveni indeks | CEN-03 | Ključ je tekst. Indeksa nema: pisac odbija, čitalac diže grešku. Zatečeni duplikati → CEN-ODL-04. |
| MAT-21 | storno potpun: vreme, autor, razlog 1–500 bez belina | CEN-11 | Nema check-a „sve tri ili nijedna": četiri ćelije piše jedan pisac u jednoj transakciji. Skup belina je uži (razmak, tab, CR, LF, NBSP; AgriXSystems skida ceo JS skup). |
| MAT-22 | storno jednokratan i konačan; drugi zid `update … where stornirano_at is null` | CEN-12 | Drugi zid je ponovno čitanje ćelije u istoj transakciji. Testa trke nema: svesku piše jedan Excel proces. |
| MAT-23 | veza ispravke u oba smera, odloženi složeni FK | CEN-13 | Nema FK. Vezu piše jedan pisac u jednoj transakciji; nijedan čitalac je ne proverava. |
| MAT-24, O177 | pravilo važeće cene u domenu; SQL bira samo kandidate | CEN-05..08 | **Nema SQL kandidata:** pravilo i čitanje su u istoj VBA funkciji, nad svim redovima tabele (`CenovnikRedovi` → `CenPobednik`). Rezerva je po tekstu vrste, ne po `vrsta_voca_id`, i ne traži red „vrsta bez sorte" u `tblKulture`. Prazna stanica upita je dozvoljena i vidi samo globalne; u AgriXSystems upit uvek ima stanicu. |
| MAT-25 | ispravka bez promene se odbija pre storna | CEN-13 | Poredi tekst bez razlike u veličini slova (AgriXSystems poredi ID-eve). Pisac proverava razlog pre provere promene; ekran pita razlog tek posle nje. |
| MAT-10 (proširen) | nova referenca na kulturu i stanicu mora biti aktivna | CEN-04, CEN-ODL-06 | VBA proverava postojanje, ne aktivnost. |
| O176 | obrazac storna i ispravke; pregled je suvo izvršenje komande (PLT-53) | CEN-11..13, §4 | Pregled je isto pravilo nad kopijom redova u memoriji, ne izvršenje pisca. Postoji samo za unos i ispravku **globalne** cene i pokazuje samo prekid posebnih. Za storno pregleda nema (C18: storno globalne vraća raniju posebnu, a ekran to ne kaže). |
| O178 | komande `cena.create` (1–100 redova), `cena.storno`, `cena.correct`; opseg po stanici | `AddCena`, `StornirajCenu`, `IspraviCenu` | Unos je jedan red. Opsega po stanici nema (CEN-ODL-01). |
| O179 | pravo `cene:write` | CEN-ODL-01 | Pravo je `MaticniPodaci`. |
| O180, PLT-51 | oblik čitanja po ulozi | — | Nema. Istoriju vidi ko vidi ekran `MAT_ROBA`. |
| O181 | granica datuma od danas u zoni firme | CEN-09 | „Danas" je lokalni `Date` računara, ne `Europe/Belgrade`. Prazan datum je `Date` (PLT-52 se ne prenosi). |
| O181, O186 | tekst pre storna i ispravke da uneti otkupi zadržavaju svoju cenu | — | Nema tog teksta. Pitanje za razlog kaže samo da je storno konačan, odnosno da nastaje nov red (`CEN_ASK_RAZLOG_*`). |
| O182 | lista istorije po stranama, filteri, dan unosa i storna, oznaka „unazad" | `modMaticniIzvor.MatKolone`, `CenovnikPoredak` | Lista je datum opadajuće, pa `CenaID`; stornirani se ne vide (CEN-ODL-10). Dana unosa, storna, razloga i oznake „unazad" u listi nema. |
| O183 | teren F2, keš važećih cena | CEN-ODL-09 | PWA je van obima. |
| O184 | audit sa vrednostima po polju | — | Nema audit zapisa po polju. Trag su sami redovi: `CreatedAt/By`, storno kolone, veza ispravke. |
| O186 | ekran „Cenovnik" | `modMaticniEkran` (`Ispravi`, `Storniraj`, kolona `Stanica`, napomena `CEN_MSG_NAPOMENA`) | Nema prikaza „važeće cene" po stanici i danu. Posle unosa editor ostaje otvoren i prazni samo cenu (zatečeno). |
| O187 | sposobnosti F-017, D-061 → `REPLACED`; D-059 `OPEN` do F2 | §9 | Ovde je D-059 u kodu odmah, jer VBA predlog na formi već postoji. |
| O188 | imena, rute, kodovi grešaka | `CEN_ERR_*` u `modPoruke` (`UpsertPoruke03`) | Ruta nema. Ishod provere je ključ poruke, ne kod greške. |

## 8) DOWNSTREAM

1. **Snimak cene.** Otkup, otpremnica i prijemnica čuvaju svoju cenu na stavci
   (`tblOtkupStavke.Cena`, `tblOtpremnicaStavke.PredlogCena`,
   `tblPrijemnicaStavke.Cena`; CEN-10).
   Stavka ne nosi `CenaID`, pa se ne zna iz kog reda je predlog došao
   (⚠ IDENTITY RISK, isto kao AgriXSystems DOWNSTREAM 1,
   `docs/STANJE.md:12611-12615`). Storno i ispravka cenovnika ne menjaju dokument
   ni dinar.
2. **Doplata** koristi samo ključ (DOP-ODL-02, `docs/DOMEN/DOPLATA.md`), ne
   pravilo važeće cene: nema datuma, stanice ni rezerve. Stope ne idu u
   `tblCenovnik`.
3. **PWA** ne čita `tblCenovnik`. Predlog na terenu je Config ključ
   `Cena<Vrsta>` (`src/js/features/otkup/otkup-form.js:372-384`), po vrsti, bez
   sorte, klase, datuma i stanice; kooperant vidi sve `Cena*` ključeve
   (`src/js/features/kooperant/koopinfo.js:14`). Izvor je `tblSEFConfig` kroz
   `modStammdatenSync.ExportConfig` / `IsPwaConfigKey` (`modStammdatenSync.bas:2571`).
   Teren i desktop zato mogu predložiti različitu cenu za isti dan.
4. **AgriXSystems F6 uvoz.** `docs/STANJE.md:12636-12643` (DOWNSTREAM 8)
   pretpostavlja da stari cenovnik nema stanicu („globalni (stari nema
   stanicu)"), da je storno samo „Da" u svesci i da su duplikati istog dana još
   u tabeli. Sada sveska nosi `StanicaID` (posebne cene), trag storna
   (`StorniranoAt`, `StorniranoBy`, `StornoRazlog`) i vezu ispravke
   (`IspravkaOdID`, `ZamenjenSaID`). Duplikate posle Ensure stornira migracija,
   po istom pobedniku koji F6 planira. **Treba javiti AgriXSystems-u pre
   pre-flight-a F6** — u ovom rezu nije urađeno.
5. **`tools/migracija/provera/otkup.py`** (`_cenovnik`, `_cena`, `:1040-1090`)
   proverava istorijske otkupe po ključu bez stanice. Komentari na `:1062` i
   `:1156` imenuju `GetVazecaCena`. Da li posle posebnih cena daje lažne nalaze
   nije provereno.
6. **`.claude/rules/agrohemija-i-cene.md:46`** i dalje imenuje `GetVazecaCena` /
   `AddCena`. Ispravka ide kroz zaseban process PR (`CLAUDE.md` §6).
7. **Admin — kapija migracije (CEN-REV-04).** Ensure posle uvoza koda u
   zatečenu svesku je obavezan korak: bez njega zatečeni duplikati obaraju
   predlog za svoj ključ i za sorte koje čitaju rezervu, i zaustavljaju unos
   globalne cene te robe (CEN-ODL-04, tačke 1–3). Isto posle
   `modMigracija.MigrirajPodatkeIzStarog`. Dok ova četiri koraka nisu
   završena, novi cenovnik nije spreman za rad operatera:
   1. kopija sveske pre uvoza koda (`.xlsm` sa strane);
   2. Admin → Ensure; greška znači da migracija nije uspela i da je tabela vraćena;
   3. u setup logu red `EnsureCenovnikSchema: stornirano N zatecenih duplikata`
      (bez tog reda: duplikata nije bilo);
   4. drugi Admin → Ensure ne sme da upiše taj red: drugi prolaz migracije ne
      stornira ništa, pa aktivnih duplikata više nema.

## 9) Sposobnosti i testovi

| ID | Promena | Status |
|---|---|---|
| **D-059** | Pravilo **zamenjeno**: datum dokumenta, stanica (F1/F2), istog dana posebna pobeđuje, rezerva, bez reda nema cene; okidači i datum i mesto. | CEN-05..08, CEN-ODL-02 |
| **D-061** | **Prošireno**: stanica, provere unosa (CEN-04), granica datuma, odbijanje duplikata, pregled pre upisa globalne cene. | CEN-01..04, CEN-09, CEN-ODL-03 |
| **F-017** | **Prošireno**: radnje „Ispravi" i „Storniraj", kolona „Stanica", napomena u zoni. „Izmeni" i dalje nema. | CEN-02, CEN-11..13 |
| nova (u F-017) | Storno reda sa obaveznim razlogom; jednokratan i konačan. | CEN-11, CEN-12 |
| nova (u F-017) | Ispravka reda: storno + nov red sa vezom u oba smera. | CEN-13 |
| nova (u D-061) | Pregled pre upisa nove globalne cene, sa Da/Ne. | CEN-ODL-03 |

Nove sposobnosti nemaju svoj ID u `MAPA_SPOSOBNOSTI.md`; zapisane su u redovima
F-017 i D-061.

Testovi (`modTest`, indeksi fiksirani unapred; nijedan nije pušten):

| # | Test | Meri |
|---|---|---|
| 366 | `T_Cen_VazecaCenaTabelaC1C19` | CEN-05, -06, -08 nad čistim funkcijama, tabela C1–C19 1:1 |
| 367 | `T_Cen_PreduslovIBezCene` | CEN-03 (čitalac diže `ERR_CEN_DUPLI_KLJUC` nezavisno od redosleda: globalna–posebna–globalna, par starijeg dana iza kasnije cene), CEN-07 (bez reda 0 = nema), prazna stanica vidi samo globalne |
| 368 | `T_Cen_UnosOdbijaIPusta` | CEN-01, -03 (pisac), -04, -09 nad tabelom, uz rollback |
| 369 | `T_Cen_StornoJednokratanSaRazlogom` | CEN-11, CEN-12 |
| 370 | `T_Cen_IspravkaVezaUObaSmera` | CEN-13, izuzetak granice iz CEN-09 |
| 371 | `T_Cen_MigracijaDuplikataIdempotentna` | CEN-ODL-04 |
| 372 | `T_Cen_PregledPrekidaPosebne` | pregled pre upisa (§4, CEN-ODL-03) |
| 373 | `T_Cen_PredlogCenePratiDatumIOM` | CEN-ODL-02 (datum, OM, prefill sa i bez cene, promena režima F1 → F4); CEN-ODL-13 samo krak F1; CEN-07 na formi, kroz pravi `UiEvent`. CEN-ODL-12 **bez testa** |
| 374 | `T_Cen_PitanjaSuGejtovana` | seam: pitanje se zabeleži, ništa se ne upiše |
| 375 | `T_Cen_UnosSeVracaPriPaduUpisa` | unos u svojoj transakciji: pad posle dodatog reda ne ostavlja red (CEN-REV-01) |
| 376 | `T_Cen_PotvrdaPrekidaStajeUDijalog` | potvrda prekida staje u `MsgBox`: ukupan broj, broj preostalih i pitanje uvek vidljivi (CEN-ODL-03) |

**Premisa postojećeg testa se menja:** `T_MatEkran_RadnjeIRezim` (indeks 149)
tvrdi `Radnje("CENOVNIK") = ""` („cenovnik nema radnji nad redom"). Sada
`Radnje` vraća `ispravi` i `storno`; tvrdnja da `izmeni` nema i dalje važi.

## 10) Nije provereno

- `run_vba` nije pušten (Linux sesija). Nijedan test 366–376 nije izvršen, a ni
  Compile (`Alt+F11 → Debug → Compile VBAProject`). Ponašanje je neverifikovano.
- UI tokovi na pravom Excelu: InputBox za razlog, MsgBox pregleda, editor
  ispravke (datum kao `dd.mm.yyyy`, cena kao `CStr(Double)` — čitanje nazad uz
  drugačiji decimalni separator nije izmereno), toast poruke, širina kolone
  „Stanica".
- Change na polju datuma ide na svaki otkucaj, a svaki poziv čita ceo
  `tblCenovnik`. Međudatumi (npr. godina od dve cifre) mogu da puste predlog za
  prošli dan. Brzina i ishod na pravim podacima nisu mereni.
- Stanica čiji naziv sadrži „ - " ne prevodi se u ID kroz `IdZaCombo`, pa unos
  pada sa `CEN_ERR_STANICA_NEPOZNATA` (odbija, ne upisuje pogrešno). Zatečeno
  ponašanje combo-a, nije izmereno.
- CEN-10 nema svoj test u ovom rezu; tvrdnja stoji na grep-u i na
  `DOCUMENT_HEADER_LINES.md` §4.1d.
- Bez testa (revizija 18.11.2026): CEN-ODL-12 (okidači ne diraju tip
  ambalaže), CEN-ODL-13 van F1 (OM ne okida u F3/F4, datum okida u F2–F4),
  fokus po ključu za polja osim datuma ispravke, „prazno = danas" pri unosu,
  drugi zid storna (CEN-12) i greška Ensure-a na neuspeloj migraciji
  (CEN-ODL-04; okidač bi tražio pokvaren upis usred Ensure-a).
- Pregled pre upisa: grana rezerve (nova globalna cena sorte prekida posebnu
  cenu vrste bez sorte) i poziv sa `izuzmiID` (ispravka) nemaju svoj slučaj u
  372.
- **UX dug (CEN-REV-03, zapisan, nije rešen).** Kad cenovnik za nov kontekst
  (stanica, datum, režim) nema cenu, polje zadrži prethodni predlog (CEN-07: bez
  reda se polje ne dira, da ne bi pao ručni unos ili cena prepisana sa
  dokumenta). Operater tako može videti cenu koja ne odgovara aktivnom
  kontekstu, bez oznake da je nasleđena. Brisanje polja nije rešenje; rešenje
  je razlikovati predloženu od ručno unete ili prepisane cene i upozoriti kad
  predlog prestane da važi. Van ovog reza.
- Promena režima uvek računa predlog iznova, pa ručno uneta cena ne preživi
  prelazak F1 ↔ F2. Kad za nov opseg cene nema, ostaje cena iz prethodnog
  režima (polje se ne prazni).
- Hipoteza: kod ispravke posle storna dvoklasnog dokumenta `dveklase` stiže u
  prefill posle cena, pa `SetKlasa 2` → `AutoFillCena` može da pregazi cenu
  storniranog dokumenta cenom iz cenovnika. **Nije reprodukovana.** Od revizije
  18.11.2026 `ApplyPrefill` na kraju vraća cene koje spec nosi, pa taj put, ako
  postoji, više ne pregazi cenu dokumenta; dvoklasan slučaj nema test.
