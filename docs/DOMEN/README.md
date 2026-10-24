# Domen — šta dokumenti jesu, nezavisno od koda

> Svrha: da se domen ne rekonstruiše iz poziva. Agent (i čovek) koji prvi put
> dira lanac dokumenata treba da zna šta je invarijanta, a šta samo trenutna
> implementacija — inače „popravi" simptom i razbije pravilo.
>
> Ovo **nije** još jedan opis koda. Gde autoritativni dokument već postoji,
> ovde stoji samo pokazivač.

## 1) Lanac dokumenata

```
Otkupni blok  ->  Otpremnica  ->  Zbirna  ->  Prijemnica  ->  Faktura
 (tblOtkup)     (tblOtpremnica) (tblZbirna) (tblPrijemnica)  (tblFakture
                                                              + tblFakturaStavke)
```

- **Otkupni blok** — jedan otkup od jednog kooperanta, na jednom otkupnom mestu,
  jednog dana. Nosi `BrojDokumenta`, veže se na otpremnicu preko `OtpremnicaID` /
  `BrojOtpremnice`, i nasleđuje `BrojZbirne`.
- **Otpremnica** — roba koja fizički ide sa otkupnog mesta. Više blokova → jedna
  otpremnica.
- **Zbirna** — agregat više otpremnica koje idu istom kupcu/hladnjači. U kanonu je
  to **zaglavlje `tblZbirna` + stavke `tblZbirnaStavke` (po klasi) + članstvo
  `tblZbirnaIzvori`**; kilaža, gajbe i klasa **nisu na zaglavlju** (S4-1).
- **Prijemnica** — prijem robe na odredištu.
- **Faktura** — obračun prema kupcu.

Uporedo, ne u lancu: **`tblAmbalaza`** (ledger kretanja gajbi), **`tblNovac`**
(uplate/isplate), **`tblPaleta` / `tblPaletaStavka`** (paletizacija).

## 2) Invarijante koje kod stvarno drži

Ove su kodirane, ne dogovorene usmeno — izvor je naveden uz svaku.

**Zbirna je agregat, otpremnice su izvor istine.**

> ZBIRNA = tačno zbir svih svojih AKTIVNIH otpremnica.

**ZBR-KANON-01 — članstvo je zapis, ne labela (S4-1).** Koje otpremnice ulaze u
zbirnu piše u `tblZbirnaIzvori`, po `ZbirnaID` i `OtpremnicaID`. `BrojZbirne`
nikad nije bio veza nego **labela** dokumenta: isti broj sme da nose dva vozača,
a storniran vlasnik broja i dalje ima aktivnu decu. Nov kod ne sme da izvodi
pripadnost iz broja.

**ZBR-KANON-02 — sadržaj se čita sa stavki (S4-1).** Kilaža, gajbe i klasa zbirne
su u `tblZbirnaStavke`, jedna stavka po klasi. `CreateZbirna_TX` ih **izvodi iz
izvornih otpremnica** i na zaglavlju ostavlja prazno
`UkupnoKolicina`/`UkupnoAmbalaze`/`Klasa` (te kolone odlaze u S4-3/S3e-2). Čitalac
je strog nad celom tabelom: zaglavlje bez stavki, stavka bez zaglavlja i dve
stavke iste klase padaju **po imenu** — `modDokumenta.StavkeZbirneRedovi`.

**ZBR-KANON-03 — izveden dokument nije mutabilni keš (odluka operatera,
21.09.2026).** Kad se izvor promeni ili stornira, zbirna se **ne prepravlja u
mestu**: nastaje **nova verzija** — storno stare + nova zbirna sa preostalim
izvorima, jedan potez i jedna transakcija, sa tragom po ID-u. Isto pravilo koje
A13 već drži za otpremnicu (S3c).

> **S4-3a je taj okvir obrisao** (`RecalculateZbirnaFromOtpremnice_TX`,
> `CompleteZbirnaIspravka`, relink po broju, `Test_ZbirnaRecalcInPlace_Auto`).
>
> **Ali zamena još ne postoji, i to je namerno.** Jedan potez za zbirnu traži da
> se sa starog dokumenta prenesu i **deca**, a zbirnu vezuju **prijemnice**. Nova
> zbirna dobija **nov broj** (storno ne oslobađa broj, A9), pa svaka prijemnica
> mora da se preveže na novi dokument, inače ostaje siroče. Prijemnica postaje
> jedan dokument (header + stavke) tek u **S6**.
>
> *Ispravljeno 20.10.2026 (pre-flight S6, plan §14.55):* ovde je stajalo da
> prijemnice zbirnu vezuju „kolonom `BrojZbirne`, jer `ZbirnaID` im nije strani
> ključ nigde u šemi“. To je netačno od S4-3c: `tblPrijemnica.ZbirnaRoditeljID` je
> u kanonu, a `PoveziDeteNaZbirnu` ga puni kanonskim `ZbirnaID`-om, fail-closed
> (prazno kad broj nije jednoznačan). `BrojZbirne` je labela.
>
> **Odluka operatera (22.09.2026): ispravka zbirne se ODLAŽE do S6**, umesto da
> se sada piše relink po broju koji S6 odmah briše. Do tada F8 nad zbirnom nudi
> `DUPLI` (razveži otpremnice, one prežive) i `PONIŠTENJE` (obori lanac) — obe
> imenovane radnje, nijedna polovična.
>
> **Odluka operatera (20.10.2026, odluka 5 u §14.55): ispravka zbirne NIJE deo
> S6.** F8 i posle S6 nudi `DUPLI` i `PONIŠTENJE`. `IspravkaZbirne_TX`, po ogledalu
> `IspravkaOtpremnice_TX`, je svoj rez **posle** S6 i KI-025/026 — tada je
> prijemnica jedan dokument sa kanonskom vezom, pa je prevezivanje dece mali posao.

**ZBR-KANON-04 — izvedena činjenica živi tačno koliko i njen izvor (odluka
operatera, 22.09.2026).** `VrstaVoca`, `SortaVoca` i `TipAmbalaze` na nacrtu
zbirne **ne bira operater** — donosi ih **prvi izvor** (`ZbrPreuzmiCinjenice`),
jer su činjenica robe, a robu donosi otpremnica. Zato, kad članstvo padne na
**nulu**, te tri vrednosti se **brišu** (`ZbrOcistiCinjeniceBezClanstva`): iza
njih više ne stoji nijedna otpremnica, a ostavljene bi tiho sužavale prazan nacrt
na vrstu koju operater nikad nije izabrao niti je na ekranu vidi. Dok ima **bar
jednog** člana se ne diraju — izvor ih i dalje pokriva, a svaki sledeći se meri
prema njima (`ZbrRequireIstiAko`).

**Storno nije brisanje.** Dokument-tabele imaju `Stornirano` kolonu; storniran red
ostaje u tabeli i izlazi iz svih agregata. Zato „aktivan" nije isto što i
„postoji". Kaskade (šta storno jednog dokumenta povlači nizvodno) su opisane u
`docs/STORNO_BACKLOG.md` i release notes za `v2.4.0`.

**Koje tabele nose storno je DEKLARISANO, ne pogađa se.** `modSchemaGuard` drži dva
spiska: `STORNO_TABELE` (moraju imati kolonu) i `BEZ_STORNA` (matični podaci, koji
je nemaju). `ExcludeStornirano` je do `v2.78.0` na nenađenu kolonu tiho vraćao
**nefiltrirane** podatke — pa je storniran dokument izlazio kao živ, iz 183 poziva.
Sada se za tabelu iz prvog spiska pada glasno, a `vba_check` pravilom
`STORNO_REGISTAR` ne pušta poziv nad tabelom koju registar ne poznaje.

**Ambalaža je ledger, ne saldo-polje.** `tblAmbalaza` čuva kretanja
(`Smer` = Ulaz/Izlaz, `EntitetID`/`EntitetTip`), a saldo se **izvodi pri čitanju**.
Ne dodavati kolonu sa saldom. Puni model: `docs/AMBALAZA_MODEL.md`.

**Kontekst otpremnice preživljava snimanje otkupnog bloka.** Datum i broj zbirne
ostaju u formi posle snimanja (sledeći blok ide u niz iste otpremnice), kooperant
se briše. Ugovor i testovi: `.claude/rules/otkup-i-dokumenta.md`.

**Registar u `modSchema` je izvor istine za šemu — ne sveska.** Obrnuto je važilo
do PR1: spiskovi kolona osnovnih tabela živeli su isključivo u `.xlsm`, pa se
prazna sveska nije mogla rekonstruisati, a obrisana kolona se videla tek kao pad
upisa satima kasnije.

Sada: `modSchema` deklariše svih 41 tabelu i 590 kolona, `EnsureAllTables` ih
pravi i dopunjava, `VerifySchema` prijavljuje odstupanje (i vrti se u health
check-u), a `SchemaReadyOrFail` je tvrda kapija pred upis. Registar je generisan
iz stvarne sveske i regeneriše se:

```
python tools/dump_schema.py <sveska> --json <put.json>
python tools/gen_schema_module.py --json <put.json>
```

Statička kapija `SEMA_REGISTAR` (`vba_check`) ne pušta `TBL_*` konstantu koje
nema u registru. Suprotan smer — tabela u svesci bez konstante — hvata generator.

**Vlasništvo nad upisom je deklarisano** (`WRITE_OWNERSHIP.json`, ugovor A11).
`python tools/who_writes.py --check-ownership` obara CI na svakog novog pisca
domen-tabele. Lista je zamrznuto zatečeno stanje — račna, ne cilj; skraćuje se
kroz PR-ove.

**KART-HRON-01 — kartica je hronologija, a tekući saldo pripada njenom redosledu
(P1, 21.10.2026).** Kartica kooperanta (`KARTICA`) i kartica ambalaže
(`AMBKARTICA`) računaju tekući saldo redom kojim su redovi poređani, pa za svaki
prikazani red važi `Saldo(i) = Saldo(i-1) + Ulaz(i) - Izlaz(i)` — i na ekranu, i u
PDF-u, i u izvozu za PWA. Pre ispravke je ljuska redove kartice **ponovo
sortirala** („druga kolona opadajuće" = broj dokumenta kao tekst; klik na
zaglavlje = bilo koja kolona), a saldo je ostajao iz drugog reda: dva kretanja
istog otkupa (izdata prazna 30, uz robu 13) imala su salde 17 i −13.

- **Ljuska ne sortira karticu.** Ekran to izjavljuje kroz `Scr_Sort` =
  `SCR_SORT_POSLOVNI` (`modScrIzvestaji`), a ljuska pita ekran u trenutku upotrebe
  (`modOtkupUI.RedosledJePoslovni`) — zaostao sort druge liste ne preuređuje
  karticu. Klik na zaglavlje ne sortira, nego javlja `OTKUI_MSG_KART_HRONOLOGIJA`.
- **Redosled određuje izveštaj, ključem** (`modIzvestaj.KarticaDokKljuc` +
  `KarticaKretanjeKljuc`, stabilno binarno sortiranje
  `modArrayUtils.SortArrayPoKljucu`): dan → izvor → broj u nizu **kao broj** →
  identitet dokumenta → unutar dokumenta: original pre kontra-stava, ulaz pre
  izlaza, vrsta kretanja, tip gajbe, `AmbID`. Fizički redosled redova tabele ne
  ulazi nigde, pa isti događaji daju istu karticu bez obzira na poreklo upisa.
- **Šta je konvencija, a šta podatak.** Broj `x/ddmmyy[-rb]` je niz po (vlasnik,
  dan) i ide +1, pa unutar niza redosled **jeste** podatak. Brojevi različitih
  nizova (otkup ↔ revers, novac, agrohemija) ne kažu šta je bilo pre: istog dana
  ide otkup → otpremnica → prijemnica → ambalažni dokument → agrohemija → novac.
  Redosled kretanja unutar transakcije knjiga ne nosi (`AMBALAZA.md` §6.13:
  „slobodan"), pa je „ulaz pre izlaza" konvencija; uz nju međusaldo dokumenta ne
  pada ispod manjeg od stanja pre i posle dokumenta. Promena konvencije je jedan
  `Select Case` (`KarticaIzvorRang`) ili jedan znak u ključu.
- **Pretraga** ne preračunava saldo isečka: red zadržava svoj istorijski saldo, a
  red početnog stanja se pod pretragom ne prikazuje (zatečeno pravilo).
- **Opis** imenuje vrstu kretanja (`modAmbalazaUgovor.AmbKretanjeNaziv`) i storno
  (`OTKUI_KART_STORNO`) — `AmbKretanjaNaloga` zato vraća i `VrstaKretanja`,
  `AmbID` i `StornoOd` (kolone 9–11, dodate na kraj).

Mere ga `modTest` 263–266 (`T_KartHron_*`), sabotaže `kart-hron-*`.

**AMB-POSL-01 — poslovni saldo ambalaže je treći broj iz iste knjige
(24.10.2026).** `ULAZ_TUDJE_AMBALAZE` nije zahtev (`AmbVrstaJeZahtev`) nego
posledica koju pisac dopisuje sam: pokriće deficita (`AMB-INV-07`) i ulaz
pozajmice (`AMB-INV-13`). To su partnerove sopstvene gajbe — fizički ih drži,
ali firma ga njima nije zadužila. Zato iz iste knjige postoje tri broja:

| Broj | Šta znači | Ko ga čita |
|---|---|---|
| fizički saldo | gde su gajbe | `AmbSaldoNaloga`, `GetAmbalazeStanje`, `ReportAmbalaza` (pregled kretanja), provere |
| poslovni saldo | šta je firma partneru izdala minus šta je vratio | kartica kooperanta, kartica ambalaže, PDF otkupnog lista, kolona Ambalaža u pregledu po OM |
| obaveza | koliko tuđe ambalaže firma duguje | `AmbObavezaPartneru` (`AMB-INV-09`) |

Primer A: kooperant drži 11, vraća 23 → pokriće 12; fizički 0, poslovno −23 na
tom dokumentu. Primer B: drži 0, vraća 3 → pokriće 3; poslovno −3.

- Projekcija je **izbor čitaoca, ne nov pisac**: `AmbKretanjaNaloga`,
  `AmbSaldoPoNalogu` i `AmbStanjeOkoDokumenta` dobijaju opcioni
  `samoPoslovno` (podrazumevano `False` = zatečeno fizičko ponašanje). Knjiga,
  pisci, storno i šema se ne menjaju.
- **Kontra-stav se sudi po vrsti originala** (`AmbRedJeTehnicki`, preko
  `StornoOd`): storno pokrića je i sam tehnički, storno pravog prenosa poslovni.
- **PDF otkupnog lista** čita poslovno početno stanje
  (`GetKooperantAmbOpening`); granica (`ODL-26/31`), stabilnost ponovne štampe i
  `outPoznato` (`ODL-27`) ostaju iste. Ostali papiri čitaju fizičko stanje istog
  jezgra.

**AMB-POSL-02 — kolona Ambalaža u pregledu po OM je poslovni saldo prema TOJ
stanici, na `datumDo`.** Bila je fizički, današnji i globalni saldo kooperanta:
pokriće se brojalo kao zaduženje, pregled za prošlost je nosio današnje stanje, a
kooperant sa dve stanice je u `ReportSaldoOMZbirni` ulazio dvaput. Svaki red
knjige imenuje obe strane, pa je pripadnost stanici podatak, ne alokacija — nema
prelivanja. Fizički pregled stanice ostaje `ReportAmbalaza("OM")`.

**DOC-SORT-01 — dokumentna lista ide po danu i broju kao broju.** Otkupni
listovi, roba OM (otpremnice), roba kupca (prijemnice) i pojedinačan pregled
ambalaže slaže izveštaj ključem `KarticaDokKljuc` (`x/080726` → `-2` → `-3` →
`-10` → `x/090726`), a ekran to izjavljuje kroz `Scr_Sort` = `SCR_SORT_POSLOVNI`
kao i kartice. Stavke istog dokumenta zadržavaju redosled papira; `UKUPNO` i
kontrolni redovi ostaju na kraju (`IzvSortTelo`). Agregati (saldo OM, isplata,
zbirni, roba vozača po vrsti) broja nemaju: izveštaj ih slaže po nazivu pa ID-u
(`IzvNazivKljuc`), a ljuska ih i dalje sme sortirati.

**REPORT-DESC-01 — opis reda otkupa na kartici kooperanta** je
`vrsta sorta klase količina kg` („Malina Willamette I 108,40 kg"), bez reči
„Otkup" (`KarticaOtkupOpis`). Prazan deo se preskače, pa nema duplog razmaka.
Ekran i PDF čitaju isti `ReportKarticaKooperanta`.

Mere ih `modTest` 267–270, sabotaže `amb-posl-*`, `kart-opis-*`, `doc-sort-*`.

## 3) Ko šta piše

`WHO_WRITES.md` u ovom folderu — generisana mapa vlasništva nad tabelama
(`python3 tools/who_writes.py --out docs/DOMEN/WHO_WRITES.md`).

Koristi je pre nego što promeniš pravilo upisa: `tblOtkup` piše **12**
produkcionih modula, `tblFakture` **9**. Kad isto polje piše više mesta po
različitim pravilima, to je klasa buga koju test hvata tek posle nastanka.

## 4) Gde je šta autoritativno

| Tema | Autoritet |
|---|---|
| Arhitektura, moduli, tokovi | `docs/ARCHITECTURE_REFERENCE.md`, `docs/ARCHITECTURE_CHANGELOG.md` |
| Ambalaža (ledger, saldo, revers) | `docs/AMBALAZA_MODEL.md` |
| Funkcionalna mapa ekrana | `docs/AgriX_Functional_Map_v142.md` |
| Funkcionalne bele zone i product roadmap | `docs/DOMEN/FUNKCIONALNE_BELE_ZONE_I_ROADMAP.md` |
| Storno i kaskade | `docs/STORNO_BACKLOG.md`, `docs/STORNO_CENTAR_PLAN_RADA.md` |
| Prerada 2.0 — proizvodno jezgro (model, faze, odluke) | `docs/PRERADA_2_MODEL_I_PLAN.md` |
| SEF (e-fakture) | `docs/SEF_LIFECYCLE_MANUAL.md` |
| Provere integriteta | `docs/INTEGRITET_PROVERE.md` |
| Arhitektonski ugovor (A1–A12, kapije, Pre-Flight) | `docs/DOMEN/ARCHITECTURE_CONTRACT.md` |
| Ciljni model dokumenata (header + stavke, PK/FK, kardinaliteti) | `docs/DOMEN/DOCUMENT_HEADER_LINES.md` |
| Vlasništvo nad upisom (A11) | `docs/DOMEN/WRITE_OWNERSHIP.json` |
| Plan refaktora, redosled PR-ova, kapija odluke | `docs/REFAKTOR_DOKUMENT_HEADER_STAVKE.md` |
| **Ambalaza — model podataka, AMB-01..04** | `docs/DOMEN/AMBALAZA.md` — jednostrana knjiga, nosioci salda, matrica knjiženja; **tri predloga čekaju operatera** |
| Identitet zbirne, vezivanje prijemnice (ZBR-IDENT-01) | `docs/DOMEN/ZBR_IDENTITET.md` — **superseded posle refaktora** |
| Novac pri stornu / ispravci otkupa | `docs/DOMEN/ODLUKA_NOVAC_PRI_STORNU.md` — **OTVORENO, čeka operatera** |
| Poznata ograničenja | `docs/KNOWN_ISSUES.md` |
| Otkup / dokumenta — pravila izmene | `.claude/rules/otkup-i-dokumenta.md` |
| Verifikacija i definicija gotovog | `CLAUDE.md` §5, `.claude/rules/testovi.md` |

## 5) Šta ovaj folder namerno NE radi

Ne duplira postojeće dokumente i ne opisuje implementaciju. Ako se nešto ovde
razilazi sa kodom, kod je u pravu i **ovaj fajl je bug** — prijavi ga kao i svaki
drugi.
