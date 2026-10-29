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

**PRJ-ZBR-01 — prijemnica se vezuje samo na IZDATU, jednoznačnu zbirnu (odluka
operatera, 24.10.2026, S6 korak 6).** Zbirna je prevozni spisak tek kad je izdata
(`IzdajZbirnu_TX`): nacrt je najava, pa bi manjak prijemnice bio meren prema
neproverenoj najavi, a nacrt se posle sme menjati. Vlasnik je **jedna kapija**,
`modDokumenta.PrijemnicaRoditeljRazlog` (identitet: postoji, jednoznačna, istog vlasnika,
istorija jednog vlasnika; pa status, sa razrešenog dokumenta), sa dva pozivaoca:
validator F4 (`modDokUnos.PrijemnicaValidiraj`, poruka uz polje) i pisac
(`SavePrijemnicaDokument`, greška pre ijednog upisa) — od S6 koraka 8; tekst razloga je
jedan (`PrijemnicaRoditeljPoruka`). Zadat roditelj po ID-u (hladnjački lanac) dobija istu
kapiju IZDATE zbirne (`ZbirnaRoditeljPrijemnice`). Pravilo važi i kad zbirne pod tim
brojem uopšte nema: odbija se **uvek** (review #440, krug 1). Do tada je uz
`PRIJEMNICA_ZBIRNA_PROVERA = UPOZORENJE` operater potvrdom pravio prijemnicu bez
roditelja (`ZbirnaRoditeljID` prazan; sonda, stavka 195), a pisac mimo F4 vezivao šta
zatekne; podešavanje, konstanta i blaža grana pisca su obrisani u koraku 8.
Plan i dokaz: `docs/REFAKTOR_DOKUMENT_HEADER_STAVKE.md` §14.55, „Korak 6" i „Korak 8".

**PRJ-ZBR-02 — prijemnica prima robu svoje zbirne (odluka operatera, 27.10.2026, S6
korak 8, KI-047).** Vrsta i sorta prijemnice su vrsta i sorta zbirne na koju se vezuje;
podela na I i II klasu sme da se razlikuje, jer je klasifikacija procena pri prijemu.
Poredi se kao kod zbirne prema izvorima: bez razmaka na krajevima, bez razlike u veličini
slova. Sorta sme da bude prazna (kultura bez sorte), i tada je prazna i na prijemnici;
izdata zbirna bez vrste je kvar i odbija se. Vlasnik je jedna kapija,
`modDokumenta.PrijemnicaRobaRazlog`, sa tri pozivaoca: F4 (poruka uz polje vrste ili
sorte), pisac (`SavePrijemnicaDokument`, i za zadat roditelj lanca, greška pre ijednog
upisa) i prevezivanje na drugu zbirnu (`ReassignPrijemnicaToZbirna_TX`, Oporavak).

**HLD-LANAC-01 — hladnjački blok je jedan lanac, 1:1 po klasi na četiri nivoa
(specifikacija operatera 20.09.2026; odluke 24.10.2026, S6 korak 7).** Blok upisan na
stanici sa `JeHladnjaca = Da` sam pravi `OTK → OTP → ZBR → PRJ`: kilaža i gajbe po
klasi jednake su na sva četiri dokumenta, svaki dokument izvodi stavke iz svog
kanonskog roditelja i vezuje se za njega **po ID-u**, a jedan blok je tačno jedan
lanac. Roba se meri pri prijemu u hladnjaču, pa sva četiri dokumenta nose datum
otkupa; vozač je ogledalo stanice; kupac je `MALINA_DEFAULT_KUPAC`; cena stavke
prijemnice je cena stavke otkupa te klase. Automatika je **obavezna**: stanica je
jedini autoritet (prekidač `AUTO_PRIJEMNICA_HLADNJACA` se briše), a ručni putevi ne
primaju hladnjački blok ni hladnjačku otpremnicu. Lanac je **nastavljiv**: svaki
korak je idempotentan po kanonskom roditelju, pa „Ponovi auto-lanac" kreće od prvog
koraka koji fali. Vlasnik: `modAutoHladnjaca.AutoLanacHladnjaca`; kapije ručnih puteva
`modScrDokumenti.VeziZaAktivnu` i `modDokumenta.DodajZbirnaIzvor_TX`. Prijemnicu
pisac vezuje za zbirnu po ID-u (`SavePrijemnicaMulti_TX`, `zbirnaRoditeljID`): broj
zbirne ume da bude dvosmislen (KR-001). Nedovršen lanac (otpremnica bez zbirne ili
zbirna bez prijemnice) vidi Oporavak, lista „Nedovršeno", red `LANAC`; PWA malina
batch hladnjačke blokove preskače, a sync ih vodi kroz lanac blok po blok.
Kupac je `MALINA_DEFAULT_KUPAC` **samo dok zbirna ne postoji**: nastavak posle zbirne
uzima kupca **roditeljske zbirne**, a pisac prijemnice zadatog roditelja drugog
vlasnika (kupac ili vozač) odbija (review #442, P1-a). Zbirna lanca nosi **tačno
jednu** aktivnu prijemnicu: dve, ili aktivna bez `PrijemnicaID`, su kvar koji
nastavak javlja razlogom sa imenom zbirne, a Oporavak redom `GREŠKA` — ne prvi
pogodak (review #442, P2). **Otpremnica pripada lancu samo ako joj je vozač ogledalo
stanice** (`VozacID = StanicaID`): oznaka stanice je trenutno stanje, pa ručna otpremnica
(pravi vozač) stanice koja je posle označena kao hladnjača nije lanac — Oporavak je ne
nudi, „Ponovi" je ne dira, a ručna F3 zbirna je prima (N1, ručna provera #442). Blok
**bez** otpremnice i dalje ide po trenutnoj oznaci: to je odluka o dokumentu koji tek
nastaje.
Plan i dokaz: `docs/REFAKTOR_DOKUMENT_HEADER_STAVKE.md` §14.55, „Korak 7" (kod:
„Tok koda, deo 2", stavka 202 u `docs/STANJE_REFAKTORA.md`; review #442: stavka 203).

**ZBR-OWN-01 — da li je zbirna interni hladnjački tok je činjenica zbirne, upisana
pri nastanku (review #442, P1-b; odluka operatera 25.10.2026).** `tblZbirna.KupacHladnjaca`
(`Da`/`Ne`, poslednja kolona kanona) piše **samo pisac zbirne**
(`BuildZbirnaHeaderRowData` pri nastanku, izmena kupca u nacrtu), po pravilu
`IsHladnjacaKupac(kupac)` u tom trenutku. PONIŠTENJE zbirne po njoj odlučuje sme li da
stornira prijemnicu i palete (`modStornoFlow.ZbirnaOwnsExternalChain`). Do review-a se
to pitalo „da li je kupac zbirne **danas** `MALINA_DEFAULT_KUPAC`", pa je promena
configa retroaktivno menjala pravilo storna: interni lanac je postajao „eksterni", a
eksterna zbirna kupca koji postane hladnjača „naša". **Prazna ili nepoznata oznaka
znači „ne zna se"** i PONIŠTENJE staje bez mutacije, sa razlogom koji imenuje kolonu;
posledice javljaju „spisak posledica nedostupan". Zatečene zbirne bez oznake se ne
dopunjavaju (nema produkcionih podataka). Dokaz: `Test_S6K7_VlasnistvoLancaNeCitaKonfig`.

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

Mere ga `modTest` 279–282 (`T_KartHron_*`), sabotaže `kart-hron-*`.

**AMB-POSL-01 — poslovni saldo ambalaže je treći broj iz iste knjige
(24.10.2026).** `ULAZ_TUDJE_AMBALAZE` nije zahtev (`AmbVrstaJeZahtev`) nego
posledica koju pisac dopisuje sam: pokriće deficita (`AMB-INV-07`) i ulaz
pozajmice (`AMB-INV-13`). To su partnerove sopstvene gajbe — fizički ih drži,
ali firma ga njima nije zadužila. Zato iz iste knjige postoje tri broja:

| Broj | Šta znači | Ko ga čita |
|---|---|---|
| fizički saldo | gde su gajbe | `AmbSaldoNaloga`, `GetAmbalazeStanje`, `ReportAmbalaza` (pregled kretanja), provere |
| poslovni saldo | šta je firma partneru izdala minus šta je vratio | kartica kooperanta, kartica ambalaže, PDF otkupnog lista, revers kooperantu (`AMB-POSL-03`), kolona Ambalaža u pregledu po OM |
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
Red prema granici (`SpoljniSvet`) nema drugu realnu stranu: otpis sa naloga
kooperanta (`AMB-10-ODL-34`, #445) ide `Kooperant → SpoljniSvet`, a stanica je
samo vlasnik dokumenta. Takav red pripada stanici koja je vlasnik ambalažnog
dokumenta (`BrojOwnerTip/ID` sa zaglavlja, isti OM kao na zapisniku), i nijednoj
drugoj — pa je zbir po stanicama jednak kartici.

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

Mere ih `modTest` 283–288 (287: otkupni list i kartica kažu isti broj; 288: otpis u pregledu po OM),
sabotaže `amb-posl-*`,
`kart-opis-*`, `doc-sort-*`.

**AMB-POSL-03 — revers kooperantu nosi poslovno stanje, kao otkupni list
(odluka operatera 28.10.2026).** Pravilo reversa iz stavke 135 (odluka 3) ostaje:
novo = stanje posle dokumenta, početno = novo − promet. Promet je i ranije bio
poslovni red dokumenta (pokriće nije poslovni red), ali je stanje bilo fizičko —
pa je pokriće ulazilo u novo i naduvavalo početno: povrat 23 uz pokriće 12
štampao je 23 → 0, a kooperant je pre povrata držao 11; sledeći revers je počinjao
od 3, ne od kraja prethodnog. Za kooperanta se stanje sada čita poslovno
(`samoPoslovno`): 11 → −12, pa −12 → −15 — isto kao kartica i otkupni list, i
papiri se nastavljaju. Revers firme (stanje OM-a) i zapisnik o otpisu ostaju
fizički. Meri ga `modTest` 289, sabotaža `amb-posl-revers-fizicki`.

**AMB-BIL-01 — bilans ambalaže je pogled firme na celu knjigu, po tipu gajbe
(odluka operatera 29.10.2026).** Izveštaji → OM → „Bilans amb.", bez izbora
entiteta, isto u oba režima. Tri odeljka, iz iste knjige, bez novog pisca:

| Odeljak | Stavke | Pravilo |
|---|---|---|
| Gde su gajbe | otkupna mesta, vozači, kooperanti, kupci (Firma kad nije 0), ukupno u opticaju | fizički saldo po vrsti naloga na kraju perioda; opticaj = zbir realnih naloga (`AMBALAZA.md` 6.14) |
| Čije su | sopstveni saldo, tuđe (i po partneru), neusklađeno | sopstveni saldo = nabavka − otpis (vrste na granici koje su zahtev) — **obračunski**, ne broj gajbi u vlasništvu: gajbe se ne vode po vlasniku, pa sme biti negativan (pozajmica 10, otpis 2 → −2: firma duguje više nego što drži); tuđe = obaveza (`AmbDoprinosObavezi`, `AMB-INV-09`) |
| Ulaz i izlaz u periodu | početak, nabavka, tuđe: partner doneo svoje, tuđe: pozajmica kupca, otpis, kraj | kretanje preko `SpoljniSvet`; kraj = početak + kretanje = ukupno u opticaju |

Red ulazi po svom danu, a kontra-stav nosi dan originala, pa je storno u svakom
periodu „kao da nije bilo" — ista istorija kao kartica (`AMB-10-ODL-26`). Ulaz
pozajmice se od pokrića razlikuje po identitetu dokumenta (`AMB-INV-13`), ne po
vrsti. Bilans je u komadima i nema finansijsku stranu. Čitalac je
`modAmbalaza.AmbBilansPodaci` (jedan prolaz kroz knjigu, kvar ide pozivaocu),
izveštaj `ReportBilansAmbalaze`. Koju vrstu na granici bilans zna, drži ugovor
(`AmbBilansStavkaGranice`), uz kapiju potpunosti `AmbBilansGranicaNepotpuna`.

**AMB-BIL-02 — „neusklađeno" je signal, ne invarijanta.** Opticaj − naše −
tuđe je 0 dok vraćene tuđe gajbe ne ostaju u opticaju. `VRACANJE_TUDJE` gasi
obavezu, a gajbe ostavlja na saldu partnera (`KI-024`). Danas ga ne zove nijedan
pisac, pa je zbir 0. Kad KI-024 dobije ulaz, ovaj red pokazuje razliku dok se
model ne ispravi. Ni jedan pisac ga ne drži; vlasnik je čitalac sa testom 291.

Mere ih `modTest` 290 (tri odeljka, granica perioda, storno, paritet sa
`AmbSaldoNaloga` i `AmbObavezaPartneru`), 291 (kapija potpunosti, KI-024, kvar
knjige) i 292 (lista na ekranu), sabotaže `amb-bil-*`.

**AMB-BIL-03 — ambalaža po partneru: bruto i neto (odluka operatera
29.10.2026).** Izveštaji → OM → „Amb. partneri", bez izbora entiteta. Red =
partner (kooperant, kupac) × tip gajbe, stanje na `datumDo`, tri već postojeća
broja iz iste knjige:

| Kolona | Pravilo | Isto kao |
|---|---|---|
| Drži | fizički saldo naloga | `AmbSaldoNaloga` |
| Tuđe | obaveza firme prema partneru | `AmbObavezaPartneru`, bilans „Tuđe" |
| Neto | poslovni saldo: + partner drži naše, − firma duguje | kartica ambalaže (`AMB-POSL-01`) |

Zbir „Tuđe" je bilans „Tuđe", a zbir „Drži" je bilans „Gde su" za kooperante i
kupce. Drži − Tuđe = Neto dok `VRACANJE_TUDJE` nema pisca (`KI-024`). Red bez
ijednog broja se ne prikazuje. Čitalac `modAmbalaza.AmbPartneriPodaci` (jedan
prolaz; stranu partnera za obavezu bira isti `BilStranaDuga` kao bilans),
izveštaj `ReportAmbalazaPoPartnerima` (po imenu, ljuska sme da sortira), podnožje
nosi zbir „Drži" i „Tuđe". Mere ga `modTest` 293 i 294, sabotaže `amb-par-*`.

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
