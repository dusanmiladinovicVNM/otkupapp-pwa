# Doplata kooperantu — knjižno odobrenje po kulturi i klasi

> **⚠ Revizija 04.11.2026 (§2a):** doplata se radi **po bloku**, **bruto** (sa PDV
> nadoknadom), sa dva načina unosa i zbirnim prikazima. §1 i §3 opisuju
> prvobitni model; gde se razlikuju, važi §2a. Prerada P1–P4 je završena (§7a).
>
> **Status (06.11.2026):** F1–F4b i prerada P1–P4 su u kodu i potvrđeni u Excelu
> (§7a). Pravilo „PDV nadoknada nikad gotovinom" (`ISPLATA_GOTOVINA.md`) je
> zatvoreno. **Plaćanje po specifikaciji** (rez A: nalozi, izvod, avans; rez B:
> gotovina na F5, ručno mapiranje) i dugme „Uključi filtrirane" su u kodu i
> potvrđeni u Excelu (05.11.2026: 12 suita zeleno, RunAllTests 346/0, dokaz rez A
> 11/11, rez B i dugme 9/9), ručni Compile PASS, „Uključi filtrirane" ručno
> proveren. P4 (kartica zbirno po obračunu) i ploča „KEŠOM NAJVIŠE" na F5 su
> potvrđeni u Excelu (RunAllTests 348/0, dokaz 5/5). Brojevi bez slova
> (DOP-ODL-30) i pitanje o ponavljanju iz kataloga potvrđeni u Excelu 06.11.2026
> (12 suita zeleno, RunAllTests 349/0, dokaz 9/9). Ručno 06.11.2026: Compile
> PASS, F5 po specifikaciji ispravan, natpis ploče staje, knjiženje nad pravim
> podacima dalo specifikaciju `1/061126-001` i dokumente `…-04`. **Ručno još
> nije provereno:** mapiranje stavke izvoda na specifikaciju, štampa
> specifikacije iz kartice, poziv na broj u CSV-u naloga (probni nalog kroz
> e-banking). Otvoreno iz §8.
>
> **U `main`-u** od 06.11.2026 (PR #460 = `95f07e6`; pravila u `.claude/rules`
> kroz #470 = `1caa86f`).
>
> **Doplata samo za jednog kooperanta** (DOP-ODL-31, 12.11.2026): u kodu, na
> grani `claude/cool-bohr-68hi14`; statičke kapije zelene, **Excel prolaz
> neverifikovan** (§7a, na kraju).

## 1) Šta je doplata

Naknadno odobrenje kooperantu. Firma **globalno** zada stopu u **din/kg** po
ključu cenovnika (vrsta + sorta + klasa) za period. Sistem sabere kilograme
aktivnih otkupa svakog kooperanta po tom ključu i knjiži **jedan iznos u korist
kooperanta**.

Vodi se kao **knjižno odobrenje**: iznos koji se obračuna je iznos koji se
isplaćuje, **bez PDV nadoknade**.

Doplata **ne menja**:

- robu (ne nosi kilograme ni u jedan robni zbir);
- cenu na stavkama otkupa (`tblOtkupStavke.Cena`);
- otkupni list (pravni dokument je već odštampan; ponovna štampa mora dati isto);
- hladnjački lanac (cena stavke prijemnice je i dalje cena stavke otkupa, HLD-LANAC-01).

Odbačene alternative:

| Alternativa | Zašto ne |
|---|---|
| Podići `Cena` na stavkama otkupa | menja već odštampan otkupni list i hladnjačku prijemnicu; nije „jedan unos" |
| Red u `tblNovac` | `Tip` u `tblNovac` nosi **kanal plaćanja** (`modConfig.bas`, „Novac Tipovi"); doplata je obaveza, ne novac — svaki čitač uplata/isplata bi morao da je isključuje |

## 2) Odluke operatera

| ID | Odluka |
|---|---|
| **DOP-ODL-01** | ~~zamenjeno sa DOP-ODL-14~~ Stopa je **din/kg**. Iznos stavke = `ZaokruziNovac(kg × stopa)`. Bez PDV nadoknade — knjižno odobrenje. |
| **DOP-ODL-02** | Ključ kulture je **ključ cenovnika**: `VrstaVoca + SortaVoca + Klasa`, tačno poklapanje kao `modCenovnik.GetVazecaCena` (trim, bez razlike u veličini slova). Prazna sorta znači **kultura bez sorte**, ne „sve sorte". |
| **DOP-ODL-03** | Obuhvat: **svi aktivni otkupi** — sva otkupna mesta, uključujući hladnjačke blokove — u periodu **od–do po datumu otkupa**. |
| **DOP-ODL-03a** | Period ima i **sat** (odluka 03.11.2026). Dan je `Datum` dokumenta; sat je vreme nastanka po pravilu otkupnog lista (`modOtkup.OtkupVremeNastanka`: `SourceCreatedAt`, pa `CreatedAt`). Sat se gleda **samo kad granica seče dan** (Od posle 00:00 prvog dana, Do pre 23:59 poslednjeg). Obe granice su uključene, na minut. Kad `SourceCreatedAt` postoji a nije vreme (PWA ISO tekst), sat je **nepoznat** — ne uzima se `CreatedAt`, jer je to za PWA red trenutak sinhronizacije — i otkup na graničnom danu obara obračun po imenu (§8, stavka 1). |
| **DOP-ODL-04** | Datum zaduženja = **datum knjiženja** doplate. Na kartici kooperanta: **jedan red po kooperantu** po obračunu (važi i posle revizije, v. DOP-ODL-18). |
| **DOP-ODL-05** | ~~deo o matičnoj stanici zamenjen sa DOP-ODL-19~~ Doplata se **isplaćuje kroz postojeće naloge za banku** i automatsko mapiranje izvoda (faza F4). Gotovinska isplata KO ide na **matičnu stanicu** kooperanta (`tblKooperanti.StanicaID`), i kad KO ima stavke na više stanica. |
| **DOP-ODL-06** | *(revidirano: po bloku i klasi, v. §2a)* **Ponovljena doplata** nad istim ključem i preklopljenim periodom je **dozvoljena**, uz upozorenje koje imenuje raniji KO i izričitu potvrdu operatera. |
| **DOP-ODL-07** | Knjižena doplata je **snimak**: ne prati kasniji storno ili ispravku otkupa. Promena osnovice je **signal** (integritet / Oporavak, faza F3), ne automatska prepravka. Ispravka = storno KO + nov obračun (obrazac ZBR-KANON-03). |
| **DOP-ODL-08** | Izveštaj **Prosečna cena** se ne menja. |
| **DOP-ODL-09** | ~~zamenjeno sa DOP-ODL-17~~ Broj KO je **`KO-<n>/<godina>`**, niz po godini datuma knjiženja, preko svih KO (i storniranih — storno ne oslobađa broj, A9). |
| **DOP-ODL-10** | *Potvrđeno 06.11.2026 (sva četiri pravila važe).* Kraj perioda nije u budućnosti; datum knjiženja nije pre poslednjeg dana perioda ni posle današnjeg; svaka stopa mora da pogodi bar jedan aktivan otkup (tipfeler u sorti ne sme da da doplatu nula za ceo ključ); prazan obračun se ne knjiži. |
| **DOP-ODL-11** | ~~zamenjeno sa DOP-ODL-18~~ Kooperantu se **štampa KO** (odluka 03.11.2026) — faza F3: `modPrint.StampajKO`, šablon `KOSablon`, PDF u „Knjizna odobrenja", podešavanje `KO_PRINT_MODE` (podrazumevano PDF). Storniran KO se ne štampa. |
| **DOP-ODL-12** | *Potvrđeno 06.11.2026:* ekran Doplata je u grupi FINANSIJE, pod pravom **Fakturisanje** (`OBL_FAKTURISANJE`) — KO je knjigovodstveni dokument kao faktura, a nova kolona prava bila bi izmena šeme bez poslovnog razloga. Promena je jedan red u `modUiScreens.ScrRows`. |
| **DOP-ODL-13** | Storno KO je **na ekranu Doplata** (radnje „Storniraj KO" i „Storniraj ceo obračun"), ne na ekranu Storno: tamo prekidač već nosi 11 lista = `MAX_SEG`. KO nema nizvodni lanac robe, pa mu ne treba pregled posledica po modovima. |

## 2a) Revizija 04.11.2026 — doplata po bloku, bruto, dva načina unosa

Operater je posle F4b ispravio model (njegova reč: „moja greška"): doplata se
računa po kilaži, ali se **radi po bloku**, po istom modelu kao otkup — sa PDV
nadoknadom. Odluke ispod **imaju prednost** nad odlukama iz §2 koje zamenjuju i
nad strukturom iz §3 (KO po kooperantu); §3–§7 opisuju kod pre revizije dok
prerada (§7, P1–P4) ne bude u kodu. Ništa od toga nije u `main`-u ni u
produkciji (PR #460, draft), pa se kanon šeme sme menjati.

| ID | Odluka (04.11.2026) |
|---|---|
| **DOP-ODL-14** | *(zamenjuje 01)* Stopa je **bruto din/kg**, isti model kao `Cena` otkupa: sadrži PDV nadoknadu (`PDV_NADOKNADA_STOPA`, podrazumevano 8%). Primer: 10,80 = osnovica 10,00 + nadoknada 0,80. Iznos stavke = `ZaokruziNovac(kg × bruto doplata/kg)` — to se isplaćuje. Osnovica i nadoknada se izvode kao na otkupnom listu (neto = bruto / (1 + stopa)). |
| **DOP-ODL-15** | **Način unosa, jedan za ceo obračun:** (1) **bruto doplata po kg**; (2) **konačna bruto cena po kg** — ono što kooperant ukupno dobija po kg. U načinu 2 doplata po kg za svaku stavku = konačna cena − bruto cena stavke sa otkupnog lista − već knjižena aktivna doplata po kg na istom bloku i klasi (ponovljen obračun ne isplaćuje razliku dvaput). Stavka sa rezultatom ≤ 0 je **„bez doplate"**: prikazuje se, ne knjiži (negativna doplata nije doplata nego umanjenje). |
| **DOP-ODL-16** | *(zamenjuje strukturu §3)* **Jedan dokument doplate po otkupnom bloku** (veza `OtkupID`), stavke po klasi. Obračun (`ObracunID`) i dalje grupiše jedan globalni unos i storno celog obračuna. Pre knjiženja operater **isključuje neželjene blokove** (podrazumevano su svi uključeni); pisac izabrane ponovo proverava (aktivan, u periodu, ključ). |
| **DOP-ODL-17** | *(zamenjuje 09)* Broj dokumenta = **broj bloka + „D"**: `1/220526D`, `1/220526-2D`; ponovljena doplata na istom bloku `…D2`, `…D3` (stornirani se broje — storno ne oslobađa broj). Kanonski broj bloka je jedinstven: prefiks je stanica, a oba pisca otkupa (ručni unos i PWA uvoz) tvrdo odbijaju broj tuđe stanice ili dana (`modBrojevi.RequireBrojUKontekstu`). **Uslov pre prve isplate:** proveriti da e-banking prihvata i vraća slovo u pozivu na broj — bez slova poziv postaje broj bloka i automatika bi novac knjižila na otkupni blok (§8). |
| **DOP-ODL-18** | *(zamenjuje 11; dopunjuje 04)* **Prikazi i PDF-ovi su zbirni.** Ekran: pregled po kooperantu, po kulturi i klasi. Kartica: jedan red po kooperantu po obračunu, opis po kulturi i klasi. Detalj po otkupu je samo u **Specifikaciji doplate** (PDF: kooperant → otkupi → stavke; na kraju rekapitulacija po kulturi i klasi; za obračun ili za kooperanta). **Papir bloka** (po uzoru na otkupni list: neto, PDV nadoknada, ukupno, klauzula čl. 34) je radnja nad jednim blokom, ne masovna štampa. |
| **DOP-ODL-19** | *(zamenjuje deo 05)* Doplata pripada **stanici bloka**: Saldo OM, isplata kroz banku (OMID vezanog reda) i gotovina (F5 na toj stanici) — kao otkupni blok. |
| **DOP-ODL-20** | Doplata je **zaseban dokument sa svojim otvorenim iznosom** (ne povećava dug bloka). Nalozi za banku **po dokumentu**, kao otkupni blok; poziv na broj = broj doplate. Izvod, gotovina i primena avansa rade kao u F4/F4b, po dokumentu. |
| **DOP-ODL-06** *(rev.)* | Blok koji već ima aktivnu doplatu za istu klasu traži **upozorenje i potvrdu** — u oba načina unosa. |
| **DOP-ODL-21** *(05.11.2026)* | **Specifikacija doplate ima svoj broj** i obuhvata **jednog kooperanta u jednom obračunu**: svaki obračun daje kooperantu novu specifikaciju, broj teče po kooperantu u sezoni (1., 2., …). Sadržaj se posle štampe i plaćanja ne menja. Štampa se i papir bloka (radnja nad jednim blokom) i specifikacija po kooperantu. |
| **DOP-ODL-22** *(05.11.2026)* | **Specifikacija je osnov plaćanja**, pored bloka: nalog za banku može po specifikaciji (jedan nalog, poziv = broj specifikacije) ili po bloku kao danas. Izvod po broju specifikacije raspoređuje iznos na njene dokumente redom (najstariji blok prvi), višak ide u avans; isto za gotovinu. Radi se **posle P2, uz P3**, kao zaseban rez sa pre-flight kapijom (dira banku i identitet dokumenta). Format broja još nije određen — mora preživeti e-banking (v. §8, stavka 6). |
| **DOP-ODL-23** *(05.11.2026)* | **Broj specifikacije** = `S` + redni broj kooperanta u sezoni + `/` + godina + `-` + šifra kooperanta (broj iz `KOOP-123`), npr. `S1/2026-123`. Slovo ga odvaja od brojeva blokova; provera e-bankinga za slovo važi zajedno sa „D" (§8, stavka 6). |
| **DOP-ODL-24** *(05.11.2026)* | **Sezona** za redni broj specifikacije je **kalendarska godina datuma knjiženja** obračuna. Storno ne oslobađa broj (kao `…D2`). |
| **DOP-ODL-25** *(05.11.2026)* | **Storno jednog dokumenta posle štampe specifikacije je dozvoljen** (kao u P2): specifikacija zadržava broj, ponovna štampa pokazuje samo aktivne dokumente i napomenu koji su stornirani; plaćen pa storniran dokument odvezuje novac u avans (DOP-05). |
| **DOP-ODL-26** *(05.11.2026)* | **Papir bloka** ima raspored otkupnog lista: **dva primerka** (kooperant i firma) jedan ispod drugog, neto po klasi, osnovica, PDV nadoknada, UKUPNO = iznos dokumenta, klauzula iz istog podešavanja kao otkupni list, potpisi. Stopa nadoknade se čita iz podešavanja u trenutku štampe, kao na otkupnom listu. |
| **DOP-ODL-27** *(05.11.2026)* | **Nalozi za banku: specifikacija UMESTO njenih dokumenata** — jedan red i jedan nalog po specifikaciji (poziv = broj specifikacije, iznos = otvoreni ostatak specifikacije); dokumenti iste specifikacije se ne nude posebno, pa isti dug ne može da se plati dvaput. Dokument bez broja (knjižen pre P3) ostaje po dokumentu. |
| **DOP-ODL-28** *(05.11.2026)* | **Raspodela po specifikaciji**: iznos ide na otvorene dokumente specifikacije **od najstarijeg bloka** (datum bloka, pa broj bloka); svaki red novca je vezan za svoj dokument i nosi **stanicu bloka**; višak je avans na matičnoj stanici. **Gotovina** po specifikaciji raspoređuje **samo na dokumente čiji je blok na stanici isplate**; ostali ostaju otvoreni. |
| **DOP-ODL-29** *(05.11.2026)* | Specifikacija je cilj plaćanja i u **ručnom mapiranju izvoda** (lista ciljeva kooperanta) i u **primeni avansa** (nalozi) — ista raspodela kao DOP-ODL-28. Redosled rada: rez A = automatski izvod + nalozi; rez B = ručno mapiranje, gotovina, avans. |
| **DOP-ODL-30** *(06.11.2026; zamenjuje oblik broja iz DOP-ODL-17 i DOP-ODL-23/24)* | **Brojevi bez slova.** Dokument doplate = **broj bloka + `-0z`**, z = redni broj doplate na tom bloku (`5/250626-13-01`, `-02`…; prvi blok dana `5/280626` → `5/280626-01`). Specifikacija = **`x/ddmmyy-00y`**: x broj matične stanice kooperanta, ddmmyy datum knjiženja, y redni broj specifikacije te stanice tog dana (`5/051126-001`). Storno ne oslobađa broj. Vrste razdvaja **nula posle datuma**: broj bloka posle datuma nema sufiks ili ima `-2`, `-3`… (nikad nulu), doplata ima `-0` + cifra 1–9, specifikacija `-00`; razlika ostaje i kad banka iz poziva izbaci separatore. Uvoz izvoda traži **tačno jedan pogodak** preko blokova, dokumenata doplate i specifikacija; poziv koji pogodi dve vrste ide na ručno. Stari brojevi sa `D`/`S` (samo u test sveskama) ostaju čitljivi kao zauzeti. |
| **DOP-ODL-31** *(12.11.2026; dopunjuje 03 i 16)* | **Obračun može biti samo za jednog kooperanta (obuhvat).** Podrazumevano obuhvat ostaje „svi aktivni otkupi" (DOP-ODL-03). Kad je zadat kooperant, plan se računa **samo nad njegovim blokovima**: tuđi ne ulaze ni u jedan korak, pa stopa koja ne pogađa nijedan njegov blok pada kao stopa bez otkupa (DOP-ODL-10) iako pogađa tuđe, a izabran tuđ blok obara knjiženje po imenu. Obuhvat je **osobina obračuna** i piše se na svaki njegov dokument (`tblDoplata.ObuhvatKooperantID`, na kraju tabele; prazno = svi). **Signal promene osnovice** (DOP-ODL-07) za takav obračun traži retro unos samo među blokovima tog kooperanta — kasnije unet blok drugog kooperanta mu nikad nije pripadao. Obuhvat se **ne izvodi iz podataka**: običan obračun koji je pogodio samo jednog kooperanta ostaje obračun za sve. Dva ulaza, jedan mehanizam: polje „Samo kooperant" na ekranu Doplata i dugme „Doplata kooperantu" na kartici kooperanta u Izveštajima (samo prelazak — Izveštaji i dalje ništa ne upisuju). Tekst u polju koji nije izabran iz liste **nije „svi"**: pregled i knjiženje staju. Posle knjiženja obuhvat se prazni, kao stope. „Uključi filtrirane" ostaje; takav obračun nema obuhvat. |

## 3) Struktura dokumenta (potvrđena 03.11.2026; **revidirana u §2a**)

**Jedno knjižno odobrenje (KO) po kooperantu.** Globalni unos je **obračun**,
koji u jednoj transakciji pravi po jedan KO za svakog kooperanta sa kg > 0.

| Tabela | Kolone (kanon: `schema/schema.json`) |
|---|---|
| `tblDoplata` (zaglavlje = jedan KO) | DoplataID, BrojDokumenta, Datum, KooperantID, ObracunID, PeriodOd, PeriodDo, Napomena, Stornirano, CreatedAt, CreatedBy, ModifiedAt, ModifiedBy |
| `tblDoplataStavke` | DoplataStavkaID, DoplataID, RedniBroj, StanicaID, VrstaVoca, SortaVoca, Klasa, Kolicina, Cena, Iznos, CreatedAt, CreatedBy, ModifiedAt, ModifiedBy |

- **Kardinalnost:** 1 obračun : N KO (jedan po kooperantu) ; 1 KO : M stavki
  (stanica × vrsta × sorta × klasa). Kooperant koji je predavao na dva otkupna
  mesta ima **jedan** KO sa stavkama po stanici; Saldo OM čita stavke svoje stanice.
- **Identitet:** `DoplataID` (`DOP-…`) je kanon; `BrojDokumenta` je labela.
  `BrojDokumenta` je pod ugovorom formata `text`, kao u svim tabelama koje ga imaju.
  Slovni prefiks `KO-` drži poziv na broj odvojen od brojeva otkupa:
  `NormalizePozivKey` (`modBankaMapiranje`) zadržava slova i cifre. **Provereno u
  F4:** broj otkupa je od cifara (`x/ddmmyy`, `modBrojevi`), faktura se razrešava
  samo na uplati, pa se `KO…` ne može sudariti ni sa jednim nizom koji banka
  razrešava na isplati. `modDoplata.KOPoPozivu` ni ne čita tabelu za poziv bez
  prefiksa.
- **Isplata (F4):** `tblNovac.DoplataID` (kolona **na kraju**; `SaveNovac` piše
  poziciono, pa pisac vezu postavlja po imenu, kao `OtkupID`). Red isplate je
  `VirmanFirmaKoop` na **matičnoj** stanici kooperanta, bez `OtkupID`-a.
- **`ObracunID`** (`DOB-…`) je zajednički ključ jednog globalnog unosa (obrazac
  `GeneracijaID`) — za pregled i za storno celog obračuna. Nije zasebna tabela.
- **Stope** ne idu u `tblCenovnik` i nisu treći model cene
  (`.claude/rules/agrohemija-i-cene.md`): žive samo na stavkama dokumenta.
- **Kartica (F2):** red nosi opis sa razlaganjem („Doplata Malina Willamette I
  1.250 kg × 10") i ref-ključ `DOP|<DoplataID>` za „Detalje", kao red otkupa.

## 4) Invarijante i vlasnici

| ID | Pravilo | Vlasnik (kod · test) |
|---|---|---|
| **DOP-01** | Doplata ne nosi robu: ne ulazi ni u jedan kg zbir; cena na stavkama otkupa se ne menja. | `modDoplata` ne piše `tblOtkup*` · 315 („plan ne piše"), 316 („cena na stavkama otkupa ostaje netaknuta") |
| **DOP-02** | `Iznos` stavke = `ZaokruziNovac(Kolicina × Cena)`; `Kolicina` = zbir aktivnih `tblOtkupStavke` tog kooperanta, stanice i ključa u periodu, u trenutku knjiženja (čitač `modOtkup.StavkeOtkupaRedovi`, ne nov). | `modDoplata.DoplataPlan` · 315, 317 |
| **DOP-03** | Storno nije brisanje: `tblDoplata` je u `STORNO_TABELE`, `tblDoplataStavke` u `BEZ_STORNA` (stavka prati zaglavlje). Storniran KO izlazi iz svih zbirova; broj ostaje zauzet. | `modStorno.StornoDoplata_TX` / `StornoObracunDoplate_TX`, čitač `modDoplata.DoplataStavkeRedovi` · 319 |
| **DOP-04** | Pisci: nastanak samo `modDoplata` (`KnjiziDoplatu_TX`, obe tabele); storno zaglavlja `modStorno`, kao kod ostalih dokumenata. | `WRITE_OWNERSHIP.json` (A11), `who_writes.py --check-ownership` |
| **DOP-05** | Novac vezan za KO se pri stornu KO **odvezuje, ne stornira** — isto pravilo kao kod otkupa (`ODLUKA_NOVAC_PRI_STORNU.md`): virman opet postaje slobodan avans, keš ostaje zatvoren posao. Odvezivanje je u istoj transakciji sa stornom zaglavlja. | `modNovac.ResetNovacDoplataLink` iz `modStorno.StornoDoplata` · 330 |
| **DOP-06** | Jedan KO po kooperantu po obračunu; brojevi uzastopni u nizu godine. | `modDoplata.KnjiziPlan` · 316 |
| **DOP-07** | Doplata je **zaduženje u korist kooperanta** svuda gde se vidi saldo: na kartici je jedan red po KO na datum knjiženja (i u početnom stanju), a u Saldo OM i PWA izvozu saldo = vrednost + doplata − isplaćeno − agro. U Saldo OM stavka KO pripada stanici na kojoj je roba predata, a period je po datumu knjiženja. | `modDoplata.DoplataIznosPoKooperantu` / `DoplataKOKooperanta`, `modIzvestaj`, `modStammdatenSync.SaldoOMDetailRedovi` · 321, 322, 323 |
| **DOP-09** | **Isplata KO nije slobodan avans.** Slobodan avans = avansni tip (`VirmanAvansKoop`/`VirmanFirmaKoop`) bez `OtkupID`-a **i bez `DoplataID`-a** — jedan predikat za sva tri čitača avansa (primena na blok, zbir raspoloživog, lista naloga). Bez toga bi primena avansa povukla isplatu KO na sledeći otkupni blok i isti dinar platio dvaput. | `modNovac.JeSlobodanAvansKooperanta` · 329, 330 |
| **DOP-10** | KO se isplaćuje kao **obaveza po dokumentu**, kao otkupni blok: otvoreno = zbir stavki aktivnog KO − aktivan novac vezan `DoplataID`-em, u cent-domenu. Nalozi ga nose pod ključem bloka (`DoplataID`) sa **brojem KO kao pozivom na broj**; izvod sa tim pozivom je jak ključ na taj KO (plan `BIM_CILJ_DOPLATA`) i knjiži se istim piscem kao blok (`MapBankaImportAsKooperantBlockCore`, grana `doplataID`): KO dobija najviše otvoreno, višak je avans, plaćen KO ceo iznos šalje u avans. Slobodan avans se na KO primenjuje **istim motorom** kao na blok (`VeziSlobodanAvans`). | `modDoplata.OtvoreniKORedovi` / `OtvorenoNaKO` / `KOPoPozivu`, `modBankaExportPregled.BuildBlokIsplataList`, `modBankaMapiranje` · 328, 329, 331 |
| **DOP-11** | **Jedna kapija i jedna veza za sve pisce novca nad KO** — izvod (auto i ručno), gotovina, primena avansa: KO postoji tačno jednom, aktivan je i kooperantov (`modDoplata.KOKapija`), a veza se piše samo kroz `modNovac.VeziNovacZaKO`. Gotovina ide **samo na stanicu bloka** (od P1; ranije matična) i najviše do trenutnog ostatka (`IsplataKOProblem`, u modulu unosa i u piscu). Izabran plaćen KO u ručnom mapiranju se ne knjiži (protivrečnost, ne avans). | `KOKapija`, `VeziNovacZaKO`, `modNovac.IsplataKOProblem`, `modDokumenta.SaveOMUlaz_TX`, `modBankaMapiranje.MapBankaImportAsKooperantBlockCore` · 331, 332, 333 |
| **DOP-08** | Ekran nema svoju računicu: stopa prolazi istu proveru kao pisac (`DoplataStopeProblem` = `UcitajStope`), pregled je isti plan koji pisac izvršava (`DoplataPlan`), a upis i storno idu samo kroz `KnjiziDoplatu_TX` / `modStorno`. Svako pitanje (knjiženje, ponavljanje, storno) ide kroz `Pitaj` — u test režimu beleži i odgovara „Ne" (KI-015). | `modScrDoplata` · 325, 326 |
| **DOP-12** | **Obuhvat je osobina obračuna**: svi dokumenti jednog obračuna nose isti `ObuhvatKooperantID`, a kad nije prazan, jednak je `KooperantID`-u dokumenta. Piše ga samo pisac doplate, pri knjiženju; storno ga ne menja. | pisac `modDoplata.KnjiziPlan` (kapija `DOP_ERR_BASE + 36` iza plana suženog u `DoplataPlan`), strog čitalac `DoplataStavkeRedovi` (`+59` tuđ kooperant, `+60` različit obuhvat u obračunu), jedno pravilo obuhvata `UObuhvatu` za plan, pisca i signal · 351 |

**Prikaz i pisac dele jedan plan** (obrazac `BimAutoPlan`): `DoplataPlan` računa
stavke bez upisa, ekran ga prikazuje (F3), `KnjiziDoplatu_TX` ga izvršava.

## 5) Poslovni događaji

| Događaj | Datum / tabela |
|---|---|
| fizički | nema — roba se ne kreće |
| poslovni / pravni | KO nastaje na datum knjiženja (`tblDoplata.Datum`) |
| finansijski | obaveza firme prema kooperantu nastaje istog dana (zaduženje na kartici); isplata je poseban, kasniji događaj u `tblNovac` |

Poslovni i finansijski događaj dele datum zato što je to pravilo knjižnog
odobrenja (DOP-ODL-04), ne zbog lakše implementacije.

## 6) Nizvodno — ko čita činjenicu

| Čitač | Promena | Gde se vidi | Faza |
|---|---|---|---|
| `modIzvestaj.ReportKarticaKooperanta` | **urađeno:** red po KO (`DOP\|<DoplataID>`), zaduženje, ulazi u početno stanje; opis `KarticaDoplataOpis` („Doplata … kg x … din"); rang izvora: doplata 6, novac 7 (KART-HRON-01) | ekran kartice, PDF po šablonu, PWA `getKartica` | F2 |
| `modIzvestaj.ReportSaldoOM` | **urađeno:** kolona **8 „Doplata"** (na kraju, pozicioni čitaoci); saldo (6) je uključuje | ekran, tabelarni PDF, KPI pločica (kolona 6) | F2 |
| `modIzvestaj.ReportSaldoOMZbirni` | **urađeno:** kolona **9 „Doplata"**; ekran sabira saldo po indeksu (OM 7, kupci 6) umesto „pretposlednje kolone" | ekran, tabelarni PDF | F2 |
| `modStammdatenSync.SaldoOMDetailRedovi` (seam izdvojen iz `ExportSaldoOMDetail`) | **urađeno:** kolona **10 „Doplata"**; saldo je uključuje; PWA kartica kooperanta u menadžmentu prikazuje doplatu | PWA saldo | F2 |
| `modDoplata.OtvoreniKORedovi` (par `GetOpenOtkupi`), `modBankaExportPregled.BuildBlokIsplataList` → `ValidateNalogSaldo` / CSV / specifikacija, `BimAutoPlanIzVrednosti` + `KOPoPozivu`, tri čitača avansa u `modNovac` | **urađeno:** KO kao otvorena obaveza (`clsBlokIsplata.doplataID` je razlikovač, ključ bloka je `DoplataID`); poziv na broj = broj KO; izvod po broju KO; `tblNovac.DoplataID` **na kraju**; isplata KO nije avans | nalozi za banku, specifikacija isplata (PDF), uvoz izvoda (kolona Predlog: „knjižno odobrenje …") | F4 |

**Ne menja se:** `ReportIsplata*`, `ReportKarticaRobaRekap`, `ReportProsecnaCena`,
otkupni list, lanac dokumenata. Ekran kartice ne treba izmenu: red bez otkupa
(`DOP|`) već dobija detalj iz opisa i poruku „štampa nedostupna" (štampa KO je F3).

**Raspored promenjen u F2, testovi prilagođeni:** formula u testu Saldo OM
(„saldo = vrednost + doplata − isplaćeno − agro") i ref-kolona zbirnog salda u
dva testa (8 → 9). Sidro sabotaže `kart-hron-izvor-rang` prati novi rang novca (7).

**Nije pročitano (pre F2/F3):** unutrašnjost `ReportSaldoOMZbirni` i
`ReportProsecnaCena`, registracija novog tipa dokumenta u storno toku i ekranu
Storno, PWA `pregled.js`.

## 7) Faze

1. **F1 — šema, pisac, storno, testovi. U KODU (03.11.2026).**
   - `schema/schema.json` (+ `modSchema` regenerisan); `modConfig` (`TBL_DOPLATA*`,
     `COL_DOP*`, `DOP_BROJ_PREFIKS`); `modSchemaGuard` (storno registar);
     `WRITE_OWNERSHIP.json`; `WHO_WRITES.md` regenerisan.
   - `modDoplata`: `DoplataPlan`, `DoplataPreklapanja`, `KnjiziDoplatu_TX`,
     `DoplataStavkeRedovi`, `AktivneDoplateObracuna`, `BrojKO`, `PoslednjiBrojKO`.
   - `modStorno`: `StornoDoplata_TX`, `StornoDoplata`, `StornoObracunDoplate_TX`.
   - `modOtkup`: `OtkupVremeNastanka` (pravilo sata, sada jedno za štampu i
     doplatu — `modPrint.FillOtkupSablon` ga zove, ponašanje štampe isto);
     `RequireValidOtkupClass` je `Public`.
   - `modTest` 315–320 (`T_Doplata_*`), sabotaže `dop-*` u `tools/sabotaza.py`.
   - Nove tabele pravi `EnsureAllTables` pri startu (`modMain.StartApp`).
2. **F2 — čitaoci. U KODU (03.11.2026).** Kartica, Saldo OM (pojedinačno i
   zbirno), PWA izvoz (§6); `modTest` 321–323, sabotaže `dop-kartica-*`,
   `dop-saldo-om-*`, `dop-zbirni-*`, `dop-pwa-*`.
3. **F3 — ekran i papir. U KODU (04.11.2026).**
   - Ekran `DOPLATA` (`modScrDoplata`, registar `modUiScreens`, ikonica
     `IC_DOPLATA` = MDL2 Calculator, *za proveru*): datum knjiženja, period sa
     satom, stope (vrsta/sorta/klasa/din-kg), liste STOPE / PREGLED / KNJIŽENA KO,
     „Pregled obračuna", „Knjiži doplatu" (pitanje sa brojkama, pa pitanje o
     ponavljanju), radnje nad KO: štampa, storno KO, storno obračuna (DOP-ODL-13).
     Runtime kontrole kroz `NewFieldG` / `BtnV`, bez novih `WithEvents`.
   - Papir KO (`modPrint.EnsureKOSablon` / `FillKOSablon` / `StampajKO`); kartica:
     „Štampaj dokument" nad redom `DOP|` štampa KO.
   - Signal osnovice (DOP-ODL-07): `modDoplata.DoplataOdstupanjaOsnovice`, provera
     **DOP1** u `modIntegritet` → lista INTEGRITET na Oporavku
     (`docs/INTEGRITET_PROVERE.md`).
   - `modTest` 324–327 (+ ekran u testu 239, seam u `ResetSeamova`), sabotaže
     `dop-papir-*`, `dop-ekran-*`, `dop-signal-nema`.
4. **F4 — isplata, bankarski deo. U KODU (04.11.2026).**
   - Šema: `tblNovac.DoplataID` na kraju (`schema.json` → `modSchema`;
     `COL_NOV_DOPLATA_ID`). Svesku dopunjava `EnsureAllTables` na startu.
   - `modDoplata`: `OtvoreniKORedovi`, `OtvorenoNaKO`, `KOPoPozivu`, `JeKO`.
   - `modNovac`: `JeSlobodanAvansKooperanta` (DOP-09) u `ApplyAvansToOtkup`,
     `GetKooperantUnallocatedAvans`, `BuildKooperantUnallocatedAvansDict`;
     `BuildIsplataDictByDoplata`; `ResetNovacDoplataLink` (DOP-05).
   - `modBankaMapiranje`: `BIM_CILJ_DOPLATA`; `ResolveStrongKooperantByKeys`
     vraća i `DoplataID`; `BimJakiKljucInfo` / `BimAutoPlanIzVrednosti` /
     `BimPlanOpis` / mreža uvoza; pisac `MapBankaImportAsKooperantKO` →
     `MapBankaImportAsKooperantBlockCore` (grana `doplataID`, kapija
     `PotvrdiKOZaKooperanta`, veza `LinkNovacToDoplataStrict`).
   - `modBankaExportPregled.BuildBlokIsplataList` nosi KO; `modScrBankaNalozi`:
     avans se ne nudi na KO (`AvansKandidat`, `PrimeniAvans`).
   - `modStorno.StornoDoplata` odvezuje novac; oba `_TX` snapshot-uju `tblNovac`.
   - `modTest` 328–330, sabotaže `dop-nalozi-*`, `dop-nalog-*`, `dop-plan-ko-*`,
     `dop-izvod-*`, `dop-avans-*`, `dop-storno-ne-odvezuje`; dve zatečene
     sabotaže banke dobile su nova sidra (isti kod, ista tvrdnja).
5. **F4b — gotovina, ručno mapiranje, avans na KO. U KODU (04.11.2026).**
   - Kapija `modDoplata.KOKapija` (`KO_KAPIJA_*`) i veza `modNovac.VeziNovacZaKO`
     — jedno pravilo za sve pisce (DOP-11); `PotvrdiKOZaKooperanta` i
     `LinkNovacToDoplataStrict` u banci ih samo prevode / zovu.
   - **(c) Avans na KO:** petlja `ApplyAvansToOtkup` izdvojena u
     `modNovac.VeziSlobodanAvans` bez promene pravila; `ApplyAvansToKO(_TX)` je
     drugi pozivalac (deljen avans: nov red vezan `VeziNovacZaKO`). Nalozi:
     „Primeni avans" (red i korpa) bira motor po razlikovaču
     (`PrimeniAvansNaDokument`).
   - **(b) Ručno mapiranje izvoda:** lista cilja (`GetBlokoviZaBimMapiranje`) posle
     otkupa nudi KO kooperanta; `BimEfektivniKO` (izabran KO, ili prazan izbor uz
     poziv KO); `BimOtvorenoNaDokumentu` hrani kapiju zatvorenog izbora, pitanje o
     višku i tekst podele; `MapBankaImportAsKooperantBlockManual_TX` nosi
     `doplataID` do istog pisca.
   - **(a) Gotovina (F5):** polje „Otvoreni blok" nudi otvorena KO kad je aktivno
     OM matična stanica (`modOtkupUI.FillOpenBlokovi`); `IsplataValidiraj`
     prebacuje izabran KO u `doplataID`; kapija `modNovac.IsplataKOProblem` u
     modulu unosa i u `SaveOMUlaz_TX` (nov opcioni parametar `doplataID`).
   - `modTest` 331–333; sabotaže `dop-avans-split-*`, `dop-ko-kapija-*`,
     `dop-rucno-*`, `dop-gotovina-*`, `dop-nalog-avans-ko-ide-na-otkup` (zamenila
     `dop-nalog-avans-na-ko`); sidro `banka-uvoz-placen-blok-postaje-avans` prati
     zajednički `BimOtvorenoNaDokumentu`.

**Spuštanje je serijsko:** F1 i F4 diraju centralne registre (`schema.json`,
`modSchemaGuard`, `WRITE_OWNERSHIP.json`, brojevi testova u `modTest`). Na dan
F1 otvoren je plan S8 paleta (PR #456, samo docs); grana koja posle doda testove
od 309 nadalje sudara se sa ovom. **Desilo se:** S8 (#458 i dalje) je uzeo 309–314,
pa su testovi doplate pri rebase-u na `main` prenumerisani na **315–333** (brojevi
u ovom dokumentu su već novi; poruke commit-ova F1–F4b nose stare 309–327).

## 7a) Prerada po §2a — faze

**P1 — model i pisac. U KODU (05.11.2026).**
- Šema na kraj tabela: `tblDoplata.OtkupID`, `StanicaID`, `NacinUnosa`;
  `tblDoplataStavke.CenaOtkupa`, `KonacnaCena` (`COL_DOP_*`, `COL_DOPS_*`,
  `DOP_NACIN_PO_KG` / `DOP_NACIN_KONACNA`, `DOP_BROJ_SUFIKS`). Kolona `KO-…`
  prefiksa više nema.
- `modDoplata.DoplataPlan(…, nacin)`: **jedan red po bloku i klasi** (kolone
  11–17: blok, broj bloka, cena otkupa, konačna cena, već doplaćeno po kg, status,
  datum bloka), redosled po upisu bloka; način 2 oduzima cenu otkupa i već
  knjiženu aktivnu doplatu po kg; red ≤ 0 je `DOP_PLAN_BEZ_DOPLATE`.
- `KnjiziDoplatu_TX(…, nacin, izabrani)`: `PlanZaKnjizenje` knjiži **tačno**
  izabrane blokove (izabran a više ne u obračunu → greška po imenu; prazan izbor
  → ništa), ponavljanje po bloku i klasi (`PonavljanjaUPlanu`), jedan dokument
  po bloku, broj `BrojDoplate(broj bloka, n)` = `…D`, `…D2` (stornirani se
  broje; zauzet broj → greška).
- Signal osnovice po bloku: knjižena kg bloka prema sadašnjoj, plus retro unos
  (otkup u periodu nastao posle knjiženja, bez aktivne doplate u bilo kom
  obračunu); isključen blok nije nalaz, a ni retro blok plaćen kasnijim
  obračunom (odluka 05.11.2026, DOP-P1-REV-04; storno te doplate vraća nalaz;
  test 337, sabotaža `dop-signal-placen-retro-ostaje`).
- Isplata na **stanicu bloka**: `StanicaKO`, vezan red izvoda (`MapBankaImportAs…Core`),
  gotovina (`IsplataKOProblem`, poruka „blok sa drugog otkupnog mesta"), nalozi
  i F5 lista (`OtvoreniKORedovi` kolona 5); poziv na broj prepoznaje broj sa „D"
  (`KOPoPozivu`).
- `modTest` 315–335 (316 i 333 preimenovani u `T_Doplata_JedanDokumentPoBloku` i
  `T_Doplata_GotovinaNaStaniciBloka`; nova 334 način 2 i 335 izbor blokova);
  sabotaže `dop-dokument-po-kooperantu`, `dop-broj-bez-rednog`,
  `dop-signal-retro-nema`, `dop-signal-iskljuceni-blok`,
  `dop-konacna-bez-vec-doplaceno`, `dop-konacna-negativna-se-knjizi`,
  `dop-izbor-ignorisan` (stare `dop-ko-po-stavci` i `dop-broj-isti` otišle).
- **Review P1 (05.11.2026), zatvoreno u kodu:** DOP-P1-REV-01 — signal
  osnovice poredi svaki DOKUMENT sa svojim snimkom (ponovljena doplata D/D2 nije
  dupla kolicina; test 334, sabotaža `dop-signal-sabira-ponovljene`). REV-02 —
  dve stavke iste klase u bloku kanon otkupa ne dozvoljava (pisac 1865, čitalac
  1924, pada pre plana), pa grana u planu ostaje odbrana u dubinu. REV-03 —
  čitalac stavki proverava kg × cena = iznos (`DOP_ERR_BASE + 55`); motor avansa
  odbija dupli NovacID (`RequireJedanAvansRed`, i za otkup); lista cilja ručnog
  mapiranja nudi doplate i kad nema nijednog aktivnog otkupa (bez testa: fixture
  uvek ima otkupe). Test 336, sabotaže `dop-citalac-bez-provere-iznosa`,
  `avans-motor-dupli-novacid`. Poziv na broj bez slova „D" ostaje **release
  kapija** (§8, stavka 6).
- **Prvi Excel prolaz (05.11.2026):** padovi su bili u PREMISAMA testova, ne u
  kodu — 317 je sat zadavao kroz `SourceCreatedAt` (pisac ga prima kao tekst;
  sat desktop otkupa je `CreatedAt`); 324 je i dalje tvrdio papir po kooperantu
  (od P1 dokument je blok); 329/330/332 su ostavljali stavke izvoda koje 96/97
  broje (sada snimak i rollback, obrazac banka suite); 334/335 su pravili blok i
  knjiženje u istoj sekundi, koju čitalac signala računa kao „posle" (blokovi sad
  nastaju minut ranije). Ponovni prolaz je zelen: RunAllTests 336/0 (215,8 s),
  banka 241/0, storno 111/0, faktura 39/0, palete 115 provera, JSON 72/72,
  StornoCentar, licenca; BFP 3014/0, agrohemija 25/0 i golden u posebnom pozivu
  (`--timeout 1500`: ceo poziv sa BFP-om prelazi 600 s, kao u S8, stavka 301).
  Compile je `NEJASNO` (ručno). Dokaz `python tools/dokaz.py dop-
  avans-motor-dupli-novacid --grupe`: **crvenih 40 / 40**, svaka svojom
  tvrdnjom u svom testu, izvor pre/posle identičan (`df90d24f076db264`), 8
  prolaza; grupno izmereno 39, jedna sama. Pun pojedinačni dokaz je isti poziv
  bez `--grupe` (pred release).
- **REV-04 (05.11.2026):** test 337 u Excelu zelen (RunAllTests 337/0), sabotaža
  `dop-signal-placen-retro-ostaje` crvena po imenu, izvor pre/posle identičan
  (`ea8401c7716eb570`).
- **Ručni Compile (05.11.2026): PASS** (`Debug → Compile VBAProject` nad temp
  kopijom iz `run_vba.py --compile-only --keep`). Time je P1 zatvoren.
- **P2 U KODU (05.11.2026); Excel: RunAllTests 338/0, dokaz `dop-ekran-` 8/8
  (grupno 4, pojedinačno 4, izvor identičan `6d50eabf6a893fc2`); ručni Compile
  PASS; ručni pregled na kopiji radne sveske ispravan (Pregled, Blokovi,
  Knjiženo, „već ima doplatu", isključenje smanjuje zbir, pitanje o ponavljanju
  posle pitanja o knjiženju, storno dokumenta `…D2` smanjuje zbir tačno za svoj
  iznos), storno celog obračuna i način „Konačna cena" rukom ispravni; naslovi
  lista skraćeni jer su bili odsečeni. **P2 ZATVOREN (05.11.2026).** Ostaje:
  tekst pitanja o ponavljanju dolazi iz pisca bez dijakritike (od F1) — uz P3:**
  ekran `modScrDoplata` — polje „Način unosa" (menja se samo bez unetih stopa,
  natpis iznosa prati način); pet lista: Stope, **Pregled** zbirno po
  kooperantu, kulturi i klasi (samo uključeni blokovi), **Blokovi** (red po
  bloku i klasi, ✓ po bloku, klik uključuje/isključuje, „Uključi sve /
  Isključi sve", status „bez doplate / isključen / već ima doplatu"),
  **Knjiženo** zbirno po obračunu i kooperantu (storno obračuna), **Dokumenti**
  po bloku (štampa, storno dokumenta, storno obračuna). „Knjiži" šalje piscu
  način i tačno uključene blokove; pitanje broji samo njih. Isključenje ostaje
  kroz nov pregled, prazni ga knjiženje ili „Isprazni stope". Test 338 (325
  prilagođen: lista „Plan" je sada zbirni Pregled), sabotaže `dop-ekran-*`
  (6 novih, 2 presidrene).
  - **„Uključi filtrirane"** (zahtev operatera 05.11.2026), treće dugme liste
    Blokovi pored „Uključi sve / Isključi sve", koja ostaju kakva su. Služi za
    doplatu jednom kooperantu: Isključi sve → pretraga po imenu → Uključi
    filtrirane. Uključuje **sve** blokove koji prolaze pretragu, na svim
    stranama, jer čita ceo pogled mreže, a ne prikazanu stranu. Ostali blokovi
    zadržavaju svoje stanje. Poruka kaže koliko blokova filter pokriva. Test
    344, sabotaže `dop-filt-*`. Excel (05.11.2026): test 344 zelen u punom
    prolazu, dokaz `dop-filt-*` 2/2 DOKAZANO. Ručna provera operatera
    (05.11.2026): radi kako treba.
- **P3 U KODU (05.11.2026), ponašanje neverifikovano dok se ne pusti u Excelu.**
  Pre-flight: odluke DOP-ODL-23..26.
  - **P3a — papir bloka** (`modPrint.FillKOSablon`, isti ulazi: lista Dokumenti i
    kartica): dva primerka po 1/3 A4 kroz isti `WriteOtkupCopy` kao otkupni list
    (h nosi tekstove i okvir doplate; otkupni list, otpremnica i grupni list
    ostaju kakvi su). UKUPNO = iznos dokumenta, osnovica = zaokruženo(UKUPNO /
    (1 + stopa)), nadoknada = razlika. Test 324 prepravljen; sabotaže
    `dop-papir-*`.
  - **P3b — broj specifikacije** `S<redni>/<godina>-<šifra>`: kolona
    `tblDoplata.BrojSpecifikacije` na kraju (kanon + `modSchema`), dodeljuje je
    samo pisac, jedan broj po kooperantu u obračunu, redni broj po kooperantu u
    godini datuma knjiženja, storno ga ne oslobađa (`+48` zauzet, `+49` van
    oblika); čitač: kolona 20, `+56` dva broja u istoj specifikaciji, `+57` isti
    broj u dve. Test 339; sabotaže `dop-spec-niz-*`, `dop-spec-storno-*`,
    `dop-spec-citalac-*`.
  - **P3c — PDF specifikacije** (`modPrint.FillSpecDoplateSablon` /
    `StampajSpecifikacijuDoplate`): detalj po dokumentu, rekapitulacija po
    kulturi i klasi, osnovica / PDV nadoknada / UKUPNO, napomena o storniranim
    dokumentima (DOP-ODL-25); storniran obračun nema specifikaciju. Lista
    Knjiženo nosi kolonu broja i radnju „Štampaj specifikaciju". Test 340;
    sabotaže `dop-spec-papir-*`.
  - Excel (05.11.2026): RunAllTests 340/0; dokaz `dop-papir- dop-spec-` 8/9 —
    `dop-spec-storno-oslobadja` je bila crvena, ali na tuđoj tvrdnji: sabotaža kvari
    pisca za ceo prolaz, raniji test sa stornom i ponovnim knjiženjem napravi dupli
    broj, a čitač (`+57`) obori svako sledeće knjiženje pre tvrdnje testa 339. Test
    339 sada ide PRVI u bloku doplate (redosled nije identitet testa). Ponovljen
    dokaz: `dop-spec-storno-oslobadja` crvena svojom tvrdnjom (uz kaskadu od 39
    testova preko čitača), izvor identičan `ae29aaa0af242bcf` — **P3 dokaz 9/9**.
    Ručni Compile PASS. Ručni pregled PDF-a (05.11.2026): iznosi i raspored
    ispravni; ispravljeno u zajedničkom kodu papira (važi i za otkupni list,
    otpremnicu i grupni list): natpis stope „8,%" → „8%" (`StopaTekst`) i
    kodiran CR „_x000D_" u klauzuli iz podešavanja (`KlauzulaZaStampu`);
    rekapitulacija specifikacije dobila prosek din/kg. Specifikacija dokumenata
    knjiženih pre P3 nosi „(bez broja)" — broj dobija samo dokument knjižen
    novim kodom (pisac pada po imenu kad kolone nema). Novi obračuni na kopiji
    radne sveske imaju broj i u listi Knjiženo i na PDF-u (potvrda operatera
    05.11.2026). **P3 ZATVOREN.** Pun prolaz sa BFP-om posle ispravke štampe
    (`a5f2e9c`) nije javljen.
  - Uzgredni nalaz (van doplate, NIJE reprodukovan u Excelu): otkupni list upisuje
    tekući račun pod ključem `racun` (`modPrint.FillOtkupSablon`), a primerak ga
    čita pod `račun` (`WriteOtkupCopy`, zamena `{RAČUN}` u klauzuli) — po kodu
    „TR:" na otkupnom listu ostaje prazno. Papir doplate piše ključ koji se čita.
- **Plaćanje po specifikaciji — rez A** (DOP-ODL-27..29; pre-flight 05.11.2026):
  - **Nalozi:** dokument sa brojem specifikacije se ne nudi posebno; nudi se
    specifikacija — jedan red (`clsBlokIsplata.specifikacija`, ključ reda = broj),
    iznos = zbir otvorenog po dokumentu (`modDoplata.OtvoreneSpecifikacijeRedovi`),
    poziv na broj naloga = broj specifikacije. Stanica reda je **matična stanica
    kooperanta** (dokumenti mogu biti sa više stanica), datum = datum knjiženja.
    Dokument bez broja (knjižen pre P3) ostaje po dokumentu (test 328).
  - **Izvod:** poziv sa brojem specifikacije je jak ključ
    (`modDoplata.SpecPoPozivu`, cilj plana `BIM_CILJ_SPEC`). Razrešava se samo
    **tačno jedan** pogodak, jer se normalizovan broj može sudariti
    (`S1/2020-26123` / `S120/2026-123`). Specifikacija bez aktivnih dokumenata
    nije cilj. Pisac je isti (`MapBankaImportAsKooperantBlockCore`, grana
    `specBroj`): iznos ide na otvorene dokumente **od najstarijeg bloka**
    (`SpecDokumentiZaRaspodelu`: datum bloka, pa broj bloka sa ciframa kao
    brojem). Svaki dokument dobija red novca vezan za njega, na stanici svog bloka.
    Višak je avans na matičnoj; isplaćena specifikacija ceo iznos šalje u avans
    (kao plaćen KO).
  - **Primena avansa** na red specifikacije (nalozi): `modNovac.ApplyAvansToSpec_TX`,
    ista raspodela, svaki dokument kroz `ApplyAvansToKO`, jedna transakcija.
    Ušla je u rez A jer red specifikacije zamenjuje dokumente u nalozima: bez nje
    bi avans nad tim dokumentima nestao (CAPABILITY).
  - Novac se nikad ne vezuje za specifikaciju, nego za dokument (`DoplataID`), pa
    ostatak, storno (DOP-05), kartica i saldo OM ostaju po dokumentu.
  - Testovi 342–343, test 328 meri dokument bez broja. Sabotaže `dop-spec-nalog-*`,
    `dop-spec-raspodela-*`, `dop-spec-plan-kao-blok`, `dop-spec-avans-*`,
    `dop-spec-poziv-storniran`.
  - Excel (05.11.2026): pun prolaz 12 suita ZELENO (RunAllTests 343/0, BFP
    3014/0, banka 241/0, storno 111/0); dokaz 7 novih i 4 presidrene sabotaže
    **11/11 DOKAZANO**, izvor pre/posle identičan (`fa03821d4083eaa1`). Ručni
    Compile i pregled naloga nad pravim podacima nisu javljeni. **Rez A
    potvrđen.**
- **Plaćanje po specifikaciji — rez B** (DOP-ODL-28/29):
  - **Gotovina / F5:** polje „otvoreni blok" nudi i specifikaciju, sa ostatkom
    njenih dokumenata **na izabranom otkupnom mestu**. Pojedinačni dokumenti
    ostaju u listi: isplata na otkupnom mestu je trenutna i kapija čita stanje
    SADA, pa ista obaveza ne može da se plati dvaput (za razliku od dva naloga za
    banku). Raspodela je `modNovac.RaspodelaSpecNaStanici`:
    - samo dokumenti čiji je blok na toj stanici, od najstarijeg bloka;
    - svaki dokument najviše svoj ostatak, a gotovinom najviše do osnovice
      (ISP-ODL-02);
    - viška nema, iznos preko dozvoljenog se odbija.

    Kapija `IsplataSpecProblem` je ista na ekranu i u piscu. Pisac
    `modDokumenta.SaveOMUlazSpec_TX` svaki deo knjiži kroz `SaveOMUlaz_TX`, sa
    istim kapijama po dokumentu, u jednoj transakciji.
  - **Ručno mapiranje izvoda:** lista cilja kooperanta nudi i njegove otvorene
    specifikacije. Izabrana specifikacija ili prazan izbor uz poziv na broj
    specifikacije (`BimEfektivnaSpec`) ide **pre** KO i otkupa. Pitanje o višku i
    tekst podele računaju iz ostatka specifikacije, red po dokumentu, istim
    redosledom kao pisac. Pisac je isti (`MapBankaImportAsKooperantBlockCore`,
    `specBroj`): tuđa specifikacija se odbija, a izabrana plaćena se ne knjiži
    (ni kao avans).
  - Testovi 345–346, sabotaže `dop-spec-got-*` i `dop-spec-rucno-*`; presidrena
    `banka-uvoz-placen-blok-postaje-avans`.
  - Excel (05.11.2026): pun prolaz 12 suita ZELENO (RunAllTests 346/0, BFP
    3014/0, banka 241/0, storno 111/0); dokaz rez B i dugmeta **9/9 DOKAZANO**,
    izvor pre/posle identičan (`d5ccebf87e8a68e2`). Ručni Compile PASS. F5 po
    specifikaciji i ručno mapiranje nad pravim podacima operater proverava
    kasnije. **Rez B potvrđen automatski.**
- **P4 — kartica zbirno po obračunu** (DOP-ODL-04/18; 05.11.2026):
  - Kartica kooperanta nosi **jedan red po obračunu**, a ne po dokumentu.
    Obračun je kooperantova specifikacija, ma koliko blokova i stanica nosio.
  - Broj reda je broj specifikacije; obračun knjižen pre P3 nosi broj svog prvog
    dokumenta. Zaduženje je zbir aktivnih dokumenata.
  - Opis je „Doplata" pa po kulturi i klasi kg × din/kg: kg i iznos se sabiraju
    preko blokova, a u načinu „konačna cena" din/kg je prosek, kao u
    rekapitulaciji specifikacije. Čitač je `modDoplata.DoplataObracuniKooperanta`.
  - Ref-ključ reda je `DOP|<ObracunID>|<KooperantID>`. „Štampaj dokument" nad
    redom kartice sada štampa **specifikaciju** (stari oblik `DOP|<DoplataID>`
    i dalje štampa papir dokumenta).
  - Iz istog izveštaja čitaju ekran, PDF kartice i PWA (`getKartica`), pa P4
    važi svuda.
  - `DoplataKOKooperanta` (po dokumentu) ostaje za listu ručnog mapiranja.
  - Test 321 prilagođen (broj = specifikacija, novi ref-ključ), novi test 348.
    Sabotaže `dop-kartica-po-dokumentu` i `dop-kartica-opis-po-bloku`;
    presidrena `dop-kartica-bez-doplate`.
  - Excel (05.11.2026): pun prolaz 12 suita ZELENO (RunAllTests 348/0, BFP
    3014/0); dokaz `dop-kartica-*` + `isp-ploca-bez-osnovice` **5/5 DOKAZANO**,
    izvor pre/posle identičan (`1b93767f970b4162`). Ručni Compile PASS; PDF
    kartice (06.11.2026) pokazuje red po obračunu (`S1/2026-00040`, `S2…`,
    `S3…`) sa opisom po kulturi i klasi. **P4 zatvoren.**
- **Brojevi bez slova (DOP-ODL-30, 06.11.2026):** `BrojDoplate` = broj bloka +
  `-0z`, `BrojSpecifikacije` = `x/ddmmyy-00y` (matična stanica kooperanta, dan
  knjiženja, niz te stanice tog dana; `KnjiziPlan`). Parser prihvata stari
  oblik `S…` kao zauzet broj (ne ulazi u niz). `KOPoPozivu` / `SpecPoPozivu` više
  ne traže slovo. `ResolveStrongKooperantByKeys` traži tačno jedan pogodak preko
  blok / doplata / specifikacija, a `BimOtkupIzPozivaNaBroj` i `BimEfektivnaSpec`
  dižu ručno mapiranje kad poziv pogodi dve vrste (`PozivViseVrsta`).
  `JeSpecifikacija` prepoznaje broj po kosoj crti. Test 339 prepisan (oblik,
  niz stanice tog dana, storno, stari oblik); 316/319/334/335 (`-01`, `-02`) i
  343 prilagođeni; nov test 349. Sabotaže `dop-spec-oblik-sa-slovom`,
  `dop-spec-isti-broj-u-obracunu`, `dop-broj-sa-slovom`,
  `dop-poziv-prvi-pobedjuje`, `dop-poziv-blok-bez-provere`; presidrena
  `dop-broj-bez-rednog`. **Prvi Excel prolaz (06.11.2026): RunAllTests 349 / 2
  palo** — 334 i 335 su u očekivanju još lepili stari sufiks (`D2`, `D`), dok je
  stvarni broj već bio tačan (`…-25-02`, `…-29-01`); ostalih 11 suite-ova
  zeleno (BFP 3014/0, banka 241/0). Očekivanje ispravljeno (`90c23d5`).
  **Ponovni prolaz (06.11.2026): 12 suita zeleno, RunAllTests 349/0, BFP
  3014/0; dokaz 9/9 DOKAZANO** (`dop-broj-bez-rednog`, `dop-ponavljanje-van-kataloga`,
  `dop-spec-oblik-sa-slovom`, `dop-spec-isti-broj-u-obracunu`, `dop-broj-sa-slovom`,
  `dop-poziv-prvi-pobedjuje`, `dop-poziv-blok-bez-provere`, `dop-spec-storno-oslobadja`,
  `dop-spec-citalac-bez-provere`), izvor pre/posle identičan `f7d90bf38df19db5`.
  Ručno 06.11.2026 nad pravim podacima: specifikacija `1/061126-001` (matična
  stanica 1, dan knjiženja, prva tog dana), dokumenti `1/250626-16-04`,
  `1/280626-8-04`… — `-04` jer niz broji SVE dosadašnje dokumente bloka, i
  stornirane i stare sa „D" (`UcitajBrojeveDoplata`, A9). Stara specifikacija
  `S1/2026-00025` ostaje vidljiva pod svojim brojem. Poziv na broj u CSV-u
  naloga ručno nije pogledan.
- **Pitanje o ponavljanju kroz katalog poruka** (05.11.2026): tekst je iz
  `modPoruke` (`DOP_ASK_PONAVLJANJE`, `DOP_ASK_PONAVLJANJE_POTVRDI`), sa
  dijakritikom, isti u grešci pisca i u pitanju ekrana. Test 318 dopunjen,
  sabotaža `dop-ponavljanje-van-kataloga`. Excel 06.11.2026: 349/0, dokaz
  DOKAZANO (u istom prolazu kao DOP-ODL-30); tekst na ekranu ručno nije pregledan.
- **Doplata samo za jednog kooperanta (DOP-ODL-31, DOP-12; pre-flight 12.11.2026).**
  Zahtev operatera: sa kartice kooperanta u Izveštajima dati doplatu samo njemu,
  bez dupliranja koda. Nema novog ekrana, pisca ni plana:
  - **Šema:** `tblDoplata.ObuhvatKooperantID` na kraju (kanon `schema.json` →
    `modSchema`, otisak `4CCA6C14` → `A2A085EB`); bez ugovora o formatu, ID nije
    vrednost koju Excel konvertuje. Postojeća sveska dobija kolonu kroz
    `EnsureAllTables` na startu (isti put kao `BrojSpecifikacije`).
  - **Motor (`modDoplata`):** opcioni `kooperantID` na kraju `DoplataPlan`,
    `DoplataPreklapanja` i `KnjiziDoplatu_TX`; filter je jedan uslov u prvoj
    petlji plana (`UObuhvatu`), pre provere perioda — tuđ blok na graničnom danu
    bez poznatog sata ne obara ovaj obračun. Pisac proverava da kooperant
    postoji tačno jednom, računa plan u obuhvatu i upisuje ga na zaglavlje;
    `KnjiziPlan` ne veruje pozivaocu (`+36`). Čitalac nosi kolonu 21 i dve
    provere DOP-12. `DoplataOdstupanjaOsnovice` preskače otkupe van obuhvata.
  - **Ekran (`modScrDoplata`):** polje „Samo kooperant" u redu A (lista sa
    ID-em u prikazu — dva kooperanta istog imena se razlikuju); stanje prati
    polje u trenutku pregleda i knjiženja (`ObuhvatSaPolja`, čita izabranu
    stavku, ne tekst); promena poništava pregled i izbor blokova; tri postojeća
    poziva motora prosleđuju obuhvat; pitanje pred knjiženje imenuje kooperanta;
    `Scr_ResetCache` osvežava listu kooperanata. Liste, zbirovi i izbor blokova
    nisu menjani — plan već nosi samo njegove blokove.
  - **Kartica (`modScrIzvestaji`):** dugme „Doplata kooperantu" na finansijskoj
    kartici, samo za tip Kooperant i samo uz pravo na ekran Doplata. Redosled
    kao kod predaje ispravke (`modScrStorno.OtvoriIspravku`): prvo
    `IdiNaEkran "DOPLATA"`, pa `Scr_DopOtvoriZaKooperanta(KooperantID)`, pa
    `RefreshFromData` (ljuska posle radnje koja promeni ekran ne osvežava
    ništa). Putuje ID kartice koja je prikazana, ne ime. Ako polje ne može da
    pokaže kooperanta, obuhvat se ne postavlja — prazno polje bi „samo ovaj"
    tiho pretvorilo u „svi".
  - **Ljuska (`modOtkupUI`):** `PanelIzboraMiruj` — javni gard za programski
    upis u combo ekrana (isti koji ljuska koristi za svoja punjenja, `mPopMute`).
    ⚠ PLATFORM UNKNOWN: da li bi bez garda panel izbora ostao otvoren posle
    dolaska sa kartice nije izmereno; gard je uveden da ishod ne zavisi od toga.
  - Testovi 351 (`T_Doplata_ObuhvatKooperanta`: plan, pisac, signal, čitalac) i
    352 (`T_Doplata_ObuhvatNaEkranu`: neizabran kooperant, dolazak sa kartice
    pravim klikom kroz ljusku, pregled, pitanje, upis, pražnjenje); 14 sabotaža
    `dop-obuhvat-*`; presidrena `dop-ekran-izbor-ne-ide-piscu`. Kapija u
    `KnjiziPlan` (`+36`) je odbrana u dubinu i nema svoju sabotažu — sama se
    testom ne može oboriti. Katalog 1265 → 1279.
  - **Ne meri se testom (ručno):** da li polje staje u red A, da li se dugme
    vidi samo na kartici kooperanta, i da li postojeća sveska dobije kolonu po
    restartu.
  - `STANJE_REFAKTORA.md` namerno nije diran: otvoren PR #469 je rezervisao
    stavke 324–327.
  **Statičke kapije zelene; Excel prolaz neverifikovan.**
**P2 — ekran.** Izbor načina unosa, zbirni pregled po kooperantu, lista „Blokovi"
sa kvačicama, knjižena lista zbirno. Odobreno 05.11.2026: pet lista (Stope,
Pregled zbirno, Blokovi sa ✓, Knjiženo zbirno po obračunu i kooperantu,
Dokumenti po bloku za štampu i storno); promena načina se odbija dok ima unetih
stopa; isključeni blokovi ostaju isključeni do knjiženja ili čišćenja; „Uključi
sve / Isključi sve" na listi Blokovi.
**P3 — papir bloka i Specifikacija doplate** (PDF), sa brojem specifikacije
(DOP-ODL-21); zatim plaćanje po specifikaciji (DOP-ODL-22, pre-flight).
**P4 — kartica zbirno po obračunu**, opis po kulturi i klasi.

## 8) Otvoreno

1. **⚠ PLATFORM UNKNOWN — sat PWA otkupa.** *(Preneto u `docs/ROADMAP.md` §9.2
   kao PWA stavka niskog prioriteta, odluka operatera 05.11.2026.)* PWA piše `CreatedAtClient` kao
   `new Date().toISOString()` (UTC, npr. `2026-11-02T07:00:00.000Z`), a uvoz ga
   prenosi u `SourceCreatedAt` kao tekst (`modMasterSync`, `GS_CREATED_AT`). Ako
   VBA `IsDate` taj oblik ne prihvata — što se očekuje, ali **nije izmereno** —
   štampa otkupnog lista za PWA otkup danas pokazuje vreme **sinhronizacije**
   (`CreatedAt`), a doplata takav otkup na graničnom danu odbija (DOP-ODL-03a).
   Test 317 to meri; sonda nad pravom sveskom je u izveštaju uz F1. Ako se
   potvrdi, odluka je operaterova: parsirati ISO + UTC → lokalno vreme u
   `OtkupVremeNastanka` (popravlja i štampu), ili ostaviti granice na celim danima.
2. ~~DOP-ODL-10 je predlog.~~ Potvrđeno 06.11.2026.
3. ~~DOP-ODL-12 je predlog.~~ Potvrđeno 06.11.2026 (grupa FINANSIJE, pravo Fakturisanje).
4. ~~Ikonica `IC_DOPLATA` nije proverena na pravom fontu.~~ Vidi se u meniju i u
   zaglavlju ekrana (snimak 06.11.2026).
5. ~~`.claude/rules/testovi.md` (32 člana u testu 239) i `.claude/rules/banka.md`
   (doplata i specifikacija kao dokumenti naloga, `DoplataID` u pravilu
   avansa).~~ Urađeno zasebnim process PR-om #470 (`1caa86f`, 06.11.2026).
6. **~~RELEASE KAPIJA — slovo u pozivu na broj.~~ Rešeno brojevima bez slova
   (DOP-ODL-30, 06.11.2026).** Poziv sada nosi samo cifre, `/` i `-`, kao broj
   bloka; probni nalog kroz e-banking ostaje kao opšta provera izvoza naloga.
   Istorija: Nalog za
   doplatu nosi poziv „1/220526D" (broj bloka + D, DOP-ODL-17), a nalog
   specifikacije „S1/2026-123" (DOP-ODL-27; šifra kooperanta koji nema oblik
   `KOOP-<broj>` je ceo njegov ID, sa slovima). Ako banka slova ne vrati, izvod po
   specifikaciji ne prepoznaje cilj i ide putem kooperanta (račun → avans), pa se
   avans primenjuje na specifikaciju u nalozima. **Nije
   izmereno** da li e-banking prihvata slovo i da li ga izvod vraća. Ako ga banka
   odbaci, poziv postaje TAČNO broj otkupnog bloka, pa automatika isplatu
   doplate prepoznaje kao isplatu **otkupnog bloka**, ne doplate (da li završi na
   bloku ili u avansu kad je blok već plaćen — nije provereno testom), a doplata
   ostaje otvorena. Pre prve isplate doplate kroz banku: jedan pravi nalog
   kroz e-banking i provera izvoda (uvesti CSV u e-banking, pa pogledati izvod).
   Ranije, sa „KO-n/godina", isti gubitak slova je vodio samo u avans. Ako se
   desi: storno pogrešnog vezivanja, pa doplata ručno (Uvoz izvoda → kooperant →
   izbor doplate) ili primenom avansa na doplatu u nalozima.
7. ~~Gotovina i Saldo OM — asimetrija stanica.~~ Rešeno revizijom (DOP-ODL-19,
   P1): doplata i njena isplata pripadaju stanici bloka, kao otkup.
8. **Naziv u meniju se seče kad je izabran** (snimak 06.11.2026, odluka operatera:
   zabeležiti, sada ne dirati). „Doplata kooperantima" je najduži naziv u meniju;
   podebljan (izabran ekran) ne staje i prikazuje se „Doplata kooperantim".
   Isti ključ `OTKUI_NAV_DOPLATA` je i naslov ekrana (`modScrDoplata.Scr_Meta`,
   `naslov=`), pa skraćenje samo u meniju traži da naslov ekrana pređe na
   postojeći `OTKUI_SCRDOP_CAP`. Predlog: u meniju „Doplata".
