# Migracija stare sveske u novu (C002, kraj sezone 2026)

> Kontekst: main je građen pod pretpostavkom da starih podataka nema
> (`DOCUMENT_HEADER_LINES.md` uvod, komentari uz `modSetup.ObrisiStareKolone*`).
> Ovaj dokument je **izuzetak** od te pretpostavke za jednokratan prenos podataka
> klijenta C002 iz sveske starog modela (red po klasi, knjiga nogu ambalaže) u
> svežu svesku iz gita. Alat: `tools/migracija/` (zamenjuje lanac MIG-01…MIG-22
> sa grane `chore/import-migration-tools`). Odluke operatera: 28.10.2026.

## Odluke

| ID | Odluka | Gde u alatu |
|---|---|---|
| **MIG-ODL-01** | Ambalaža: **puna istorija** red po red u AMB-10. Pravila prevoda po starom tipu dokumenta su ista kao u MIG-04…MIG-19 (tabela u zaglavlju `ambalaza.py`). Stornirani stari redovi idu samo u arhivu, ne u knjigu (MIG-05.7/05.8). | `ambalaza.prevedi` |
| **MIG-ODL-02** | Knjiženje „ispravno”: istorija + ono što main traži za rad (AMB-INV-07: nijedan realan nalog ni jednog dana ispod nule). Kooperant → `ULAZ_TUDJE` / `VRACANJE_TUDJE` po dnevnom zaključku (MIG-16/17); kupac → `ULAZ_TUDJE` u trenutku manjka; stanica/vozač → jedna `NABAVKA` (dokument vrste NABAVKA) na prvi dan rada, u iznosu najdubljeg minusa. Sve dodato nosi `CreatedBy = MIGRACIJA`. | `ambalaza.ISPRAVNO` |
| **MIG-ODL-03** | Pojedinačni redovi: AMB-00700 izuzet (aktivna noga stornirane otpremnice, MIG-19.12); OMUlaz AMB-02317/03849/03939 → `PRENOS_INTERNO` sa ST-00001 (MIG-19.11); OMUlaz bez izvora AMB-01377/07115/07815 → `NABAVKA` na ST-00001. Svaka odluka važi samo dok red izgleda kao u potpisu. | `izuzeci.json` |
| **MIG-ODL-04** | Dokumenti: stari red = jedno zaglavlje + **jedna** stavka (izmereno nad sveskom: 0 dokumenata sa više redova). Zaglavlje zadržava stari ID. Kolone koje današnji pisac ne piše ostaju prazne (`tblOtkup.OtpremnicaID/BrojOtpremnice/BrojZbirne/ZbirnaRoditeljID/VozacID`, `tblOtpremnica.BrojZbirne/ZbirnaRoditeljID`, `GeneracijaID`). | `cilj.DOKUMENTI` |
| **MIG-ODL-05** | Članstvo: otkup → otpremnica iz starog `tblOtkup.OtpremnicaID` (MIG-11); otpremnica → zbirna iz `BrojZbirne`, po broju pa po vozaču i datumu (MIG-11). Veza aktivne otpremnice bez otkupa popravlja se samo strogim 1:1 dokazom (MIG-21, bez para-poteza); OTP-00666 se nikad ne popravlja. Prijemnica → zbirna kao pisac prijemnice: `ZbirnaRoditeljID` + zadržan `BrojZbirne`. | `cilj._clanstvo`, `_popravi_veze`, `_prijemnica_roditelj` |
| **MIG-ODL-06** | Status: otpremnica i zbirna `IZDATO` (stari program nije imao nacrt, MIG-14). Otkup `IZDATO`, osim **aktivnog otkupa bez aktivne otpremnice** — ostaje prazan, jer bi ga sync sam stavio u automatsku otpremnicu (`modMasterSync.GrupeZaAutoOtpremnicu`, `modAutoHladnjaca.AutoLanacZaNevezaneBlokove`). Aktivna otpremnica bez dokazanog otkupa ostaje neizdata (izveštaji robe bi inače pali, `VrednostIzvoraZaOtpremnicu`). | `cilj._status_*`, `_otpremnice_bez_izvora` |
| **MIG-ODL-07** | Podešavanja: vrednost iz stare sveske prepisuje test vrednost sveže sveske (`make_fixture`); aktivacija licence (`LICENSE_KEY/TOKEN/BOUND_PARTS/HWM/STATUS/NEXT_CHECK`, `TRIAL_START/HWM`) se ne prenosi — licenca se aktivira ponovo, kao `modMigracija` na drugoj mašini. | `cilj._podesavanja`, `upis.upisi_podesavanja` |
| **MIG-ODL-08** | Ne prenosi se: `tblStornoVeze`, `tblStornoZurnal` (pokazuju na kolone starog modela), `tblPoruke` (katalog iz nove sveske). | `cilj.NE_PRENOSI` |
| **MIG-ODL-09** | **Storniran otkup bez količine (0 kg) ne ide u novu svesku**: main ga ne ume da pročita (`modOtkup.RequireStavkaUgovor`, Kolicina > 0 važi i za storno), a storniran dokument ne ulazi ni u jedno saldo ni izveštaj — ostaje u staroj svesci kao arhivi. Aktivan otkup sa 0 kg i dalje BLOKIRA. Odluka operatera 28.10.2026 (umesto zakrpe MIG-10, grana `migration/mig10-cancelled-zero-readmodel`). | `cilj._izostavi_stornirane_bez_kolicine` |
| **MIG-ODL-10** | Isplate kooperantu bez `OtkupID` (5 redova, `VirmanFirmaKoop`/`VirmanAvansKoop`) su **stvarni avansi** — ostaju bez veze, a prvi sledeći otkup kooperanta ih troši (`modNovac.JeAvansKooperanta`). Odluka operatera 28.10.2026. | — |
| **MIG-ODL-11** | Otvorene palete stare sezone se **zatvaraju** (`Status = Zatvorena`), inače ih nova sezona puni gajbicama iste vrste/sorte/klase (`modPaletniList`). | `cilj._zatvori_otvorene_palete` |
| **MIG-ODL-12** | Prihvaćeno nad sveskom s kraja sezone: 17 aktivnih otpremnica bez dokazivog otkupa ostaje neizdato (MIG-ODL-06), 3 aktivna otkupa bez otpremnice ostaju neizdata, aktivacija licence se ne prenosi (MIG-ODL-07). | — |
| **MIG-ODL-13** | **Vraćanje tuđe ambalaže ima svoj izlaz** (`AMB-INV-15`, #448; `KI-024` deo 2). Uz svako migraciono `VRACANJE_TUDJE` ide `IZLAZ_TUDJE_AMBALAZE` partner → SpoljniSvet, isti dokument, dan, tip i količina — isti red koji piše `PrenesiAmbalazu`. Fizički saldo kooperanta na kraju dana je max(stari, 0), obaveza max(−stari, 0), a poslovni saldo (kartica) = stari. Bilans „Neusklađeno" je 0 (`AMB-BIL-02`). | `ambalaza.primeni_politiku_duga`, `_izlaz_tudje` |
| **MIG-ODL-14** | **Knjiga robnog dokumenta je ono što bi upisao pisac main-a** (`KI-065`, odluka operatera 31.10.2026 za otpremnicu). Otpremnica: gajbe = zbir gajbi blokova istog tipa, i u dokumentu i u knjizi (`OtpRequireJednakost`); **neizdata otpremnica ne knjiži** (`OtpIzdaj`) — njene stare noge su van knjige (`DOKUMENT_NE_KNJIZI`); blok drugog tipa ambalaže ne ulazi u zbir (`OtpRequireIzvorValjan`, ODLUKA). Zbirna: gajbe = zbir izdatih otpremnica (`ZbrRequireJednakost`). Otkup i prijemnica: **knjiga je istorija** (MIG-ODL-01), pa se gajbe dokumenta (`KolAmbalaze` stavke, `KolAmbIzdata`, `KolAmbVracena`) usklađuju sa njom; dokument sa nerazrešenom nogom se ne dira. Kg se ne menja nigde. Svaka izmena je nalaz u probi i red na listu „Uskladjivanje" pregleda. | `cilj._robni_dokumenti`, `ambalaza.uskladi_sa_dokumentima` |
| **MIG-ODL-15** | **NABAVKA pokriva i prazne** (`AMB-INV-16`, #448; vozač od `AMB-10-ODL-37`, #449). Jedna NABAVKA sopstvenog naloga prvog dana rada, u iznosu dubljeg od dva minusa: ukupnog salda i praznih (svaki red naloga osim `AMBALAZA_UZ_ROBU`, `AmbPrazneNaloga`). Pune ispod nule nabavka ne može da popravi — proba ih javlja kao ODLUKA (`AMB_PUNE_ISPOD_NULE`). | `ambalaza.nabavka_na_pocetku`, `pune_i_prazne`, `provere._ambalaza` |
| **MIG-ODL-16** | **Linija fakture nema svoj storno** (S7 rez 2, `FAK-ODL-03`; šema gubi `tblFakturaStavke.Stornirano/OsirocenoOd` i trace kolone fakture). Stara linija sa `Stornirano = Da` na **aktivnoj** fakturi ne ide u novu svesku (inače bi oživela i ponovo fakturisala stavku) — ostaje u staroj kao arhiva, ODLUKA. Provere prate `FakturaLinijePoStavci`: linija prijemnice mora imati fakturu bez obzira na storno; stavka prijemnice samo na aktivnoj liniji, i najviše na jednoj (`FAK-LIN-01/02`). Sveska C002 nema nijednu fakturu. | `cilj._linije_fakture`, `provere._fakture` |
| **MIG-ODL-17** | **Pune gajbe vozača ispod nule se prihvataju** (odluka operatera 03.11.2026). Uzrok: 17 otpremnica bez dokazivog otkupa po main-u ne knjiži gajbe (MIG-ODL-14), a prijemnice kupca su ih primile, pa vozač po knjizi predaje više punih nego što je dobio — main to dozvoljava (prijemnicu sudi samo ukupni saldo, AMB-INV-07). Nad C002: vozač ST-00001 0,4kg −125, ST-00002 0,4kg −6, ST-00003 0,4kg −669, ST-00003 0,5kg −47, ST-00007 0,4kg −7; za toliko su mu prazne iznad stvarnih, pa ih „Izdato OM" (`AMB-INV-16`, ODL-37) pušta više nego što fizički drži. Odbijena alternativa: gajbe neizdatih otpremnica u knjizi — manjak bi se vratio na stanice (KI-065), gde kapija praznih utiče na svakodnevni otkup. Otkupi za tih 17 otpremnica u staroj svesci ne postoje (datumi ne odgovaraju ni jednom od 3 otkupa bez otpremnice). | `provere._ambalaza` (ODLUKA `AMB_PUNE_ISPOD_NULE`) |
| **MIG-ODL-18** | **Otpremnica bez ijednog bloka je prava otpremnica kojoj fale blokovi, ne greška** (odluka operatera 03.11.2026: „otkupi su obavljeni, samo blokovi nisu uneti"; stari tok je prvo unosio otpremnicu i prijemnicu, pa blokove). main je izdaje tek sa blokovima (izveštaji robe je vrednuju iz njih), pa ostaje **nacrt**: stavke su očekivanje (stari kg i gajbe), ne knjiži gajbe; operater u novoj svesci unosi blokove i izdaje je (tada se knjiže i gajbe, smanjuje MIG-ODL-17). Bez aktivne zbirne sa prijemnicom nije dokazana kao prava — ODLUKA. **Otpremnica sa bar jednim blokom ostaje izdata i kad se kg ne slaže sa blokovima** — to operater rešava sa otkupcem za svaki otkup i otpremnicu; alat ne uvodi pravilo, samo daje spisak (UPOZORENJE `OTPREMNICA_PREOSTALO_KG`). Spisak nacrta (kg, zbirna, prijemnice, kandidati sa viškom blokova) je u `proba_sazetak.txt`. | `cilj._otpremnice_bez_izvora`, `provere._otpremnice_preostalo` |
| **MIG-ODL-19** | **OTP-00666 (2/050726, ST-00002) se prenosi kao stornirana** (odluka operatera 03.11.2026): nema blokova ni prijemnice, a zbirna S2/050726 joj je stornirana. Odluka je u `izuzeci.json` (sekcija `dokumenti`, radnja `storniraj`) i važi samo dok otpremnica u staroj svesci ima isti broj, datum i stanicu. Stornirana otpremnica nema neto knjigu, pa njena stara noga ambalaže izlazi iz knjige. | `cilj._storniraj_po_odluci`, `izuzeci.json` |
| **MIG-ODL-20** | **Prijemnica na pogrešnoj zbirnoj je greška stare forme i ispravlja se pre migracije** (odluka operatera 07.11.2026). `frmDokumenta` (v2.2.3–v2.28.4) posle „Unesi prijemnicu" nije praznila broj zbirne, a promena vozača ga nije dirala; pisac do v2.21.0 nije proveravao ništa, od v2.21.0 samo da zbirna postoji i meko pitanje „1 zbirna = 1 prijemnica". Sledeća prijemnica je zato nasleđivala zbirnu prethodne. Trag je uvek par: aktivna zbirna sa dve prijemnice + aktivna zbirna bez ijedne. **Pogrešan je kasniji unos.** Kad je druga stanica, prefiks broja se ne slaže sa vozačem (migracija javlja `PRIJEMNICA_BEZ_JEDNOZNACNE_ZBIRNE`); kad je ista stanica, ne javlja ništa i veže prijemnicu na pogrešnu zbirnu. Operater je nad C002 našao: PRJ-00177 S5/180626 → S5/180626-2, PRJ-00185 S4/100626 → S4/120626, PRJ-00195 S4/200626-2 → S4/210626, PRJ-00280 S5/240626 → S5/240626-2, PRJ-00313 S5/250626-2 → S8/250626, PRJ-00520 S5/010726-2 → S8/010726, PRJ-00549 S4/020726 → S4/020726-2, PRJ-00584 S3/030726 → S2/030726. Zbirne čije su otpremnice stornirane, a same nisu — treba ih stornirati: S3/220626, S3/220626-2, S5/210626, S5/130626. S3/040726-3 nema prijemnicu (otvoreno). **Način primene: odluke u `izuzeci.json`, original ostaje netaknut** (operater je izbor prepustio 07.11.2026). Radnja `prevezi_prijemnicu` menja `BrojZbirne` prijemnice samo dok je trenutni broj kao u potpisu i dok prava zbirna postoji kao aktivna; vezu (`ZbirnaRoditeljID`) posle pravi isto pravilo kao pisac (broj, kupac, vozač). Radnja `storniraj` sa `"tabela": "tblZbirna"` stornira sve aktivne redove tog broja, i to tek posle prevezivanja i članstva, i samo kad nijedna aktivna otpremnica nije član te zbirne i nijedna aktivna prijemnica ne nosi taj broj. Inače odluka ne važi i stavka je ODLUKA (`IZUZETAK_DOKUMENTA_POTPIS_SE_NE_SLAZE`, `PREVEZIVANJE_NA_ZBIRNU_KOJE_NEMA`). | `lanac.py` (`migracija.py lanac`), `cilj._prevezi_prijemnice_po_odluci`, `cilj._storniraj_zbirne_po_odluci`, `izuzeci.json` |
| **MIG-ODL-21** | **`KulturaID` otkupnog lista se određuje po vrsti I sorti** (KI-093; operater je izbor prepustio 07.11.2026). Stari pisci su je tražili samo po vrsti — i desktop `SaveOtkup` (`vba-v2.28.4:src-vba/modOtkup.bas:543`) i PWA uvoz (`modMasterSync.bas:1583`): prvi red vrste, a bez pogotka izmišljen `"vrsta-sorta"`. Zato je svaki list čija sorta nije prva sorta vrste nosio tuđu kulturu (npr. „Willamette teren" sa KUL-00001 = Willamette). Pravilo je isto kao `modOtkup.RazresiKulturuIzVrsteSorte` u main-u: **tačno jedan** red `tblKulture` iste vrste i sorte, bez obzira na velika slova i spoljne razmake; nula i više od jednog su ista greška. Tekst lista (`VrstaVoca`/`SortaVoca`) i `tblKulture` se ne menjaju — samo `KulturaID`, i aktivnom i storniranom listu. Bez tačno jednog pogotka `KulturaID` ostaje kakav je i to je ODLUKA. Zaglavlje otpremnice stare sveske kulturu nema i ostaje prazno (stari oblik, KI-091). | `cilj._kultura_otkupa` (INFO `KULTURA_PO_SORTI` po paru, ODLUKA `KULTURA_NE_RAZRESAVA`) |
| **MIG-ODL-22** | **Storniran otkup nije član AKTIVNE otpremnice** (operater je izbor prepustio 07.11.2026). Stara sveska ga je ostavljala na otpremnici; main to stanje nikad ne pravi (storno bloka u lancu pravi novu verziju otpremnice, `modStorno.StornoOtkup` odbija storno člana) i ne podnosi ga: „U blokovima" ga sabira, a specifikacija pada („Izvor otpremnice … je storniran"). Knjiga ambalaže ga ni ranije nije brojala (`_otpremnice_po_blokovima`), pa se salda ne menjaju; kg otpremnice ostaje (stavke su očekivanje), pa razliku pokazuje `OTPREMNICA_PREOSTALO_KG`. Stornirana otpremnica zadržava sve svoje članove kao istoriju. | `cilj._clanstvo` (INFO `STORNIRAN_OTKUP_NIJE_CLAN_AKTIVNE_OTPREMNICE`) |
| **MIG-ODL-23** | **Ispravke unosa koje menjaju i knjigu ambalaže se primenjuju na staru svesku u memoriji, pre prevoda** (original se ne menja; odluke operatera 08.11.2026). Vozač i storno prijemnice moraju preći i na njene noge u staroj knjizi, jer se knjiga prevodi iz nogu. Odluke u `izuzeci.json`, svaka uz potpis: **(1) tipfeler sorte** „Mekker" (5 listova) i „Miker" (1): stanica ST-00001 → Meeker, ostale → Meeker teren — samo dok je broj takvih listova tačno kao u odluci i nova sorta postoji u šifarniku; kultura se zatim razrešava po sorti (MIG-ODL-21). **(2) vozač prijemnice**: PRJ-00749 ST-00005 → ST-00008 (zamena storniranih PRJ-00745/746 za S8/100726, uneta odmah posle prijemnica stanice 5 — zbirna tačna, vozač zaostao), PRJ-00908 ST-00006 → ST-00003 (S3/260726-2) — i na prijemnici i na njenim nogama. **(3) storno prijemnice** PRJ-00830 (S2/150726-2): zbirna ZBR-00844 i otpremnica OTP-00844 su stornirane, a blokova nema — prijemnica i njene noge se storniraju; ne kad je na aktivnoj paleti ili fakturi. | `cilj._ispravke_izvora` (INFO `SORTA_ISPRAVLJENA_PO_ODLUCI`, `PRIJEMNICA_VOZAC_ISPRAVLJEN_PO_ODLUCI`, `PRIJEMNICA_STORNIRANA_PO_ODLUCI`) |

## MIG-WR-01 — izuzetak od vlasništva upisa

Masovni upis u **praznu** svesku iz gita je jednokratno punjenje pre rada, ne
pisac u radu. Zato je izuzetak od AMB-INV-08 (upis u knjigu samo u transakciji
izvornog dokumenta) i od A11 (`WRITE_OWNERSHIP.json`, red `tblAmbalaza` →
`modAmbalaza`). Presedan: `modMigracija.MigrirajPodatkeIzStarog`.

Uslovi koje alat drži:

- **pre upisa**: `provere.py` meri ista pravila po kojima main čita celu tabelu
  (`Stavke*Redovi`, članstvo, izdate otpremnice bez izvora, fakture) i knjigu
  ambalaže (invarijante, salda, `AMB-INV-07/15/16`, `AMB-BIL-02`). Nalaz
  BLOKIRA odbija upis; ODLUKA traži `--prihvatam-odluke`.
- **upis**: svaka tabela u koju se piše mora biti prazna; izvorna sveska se ne
  otvara u Excelu; makroi i događaji ugašeni.
- **posle upisa**: svaka ćelija se poredi sa planom, a salda ambalaže nove sveske
  sa saldima prevoda.

## Stanje

- **01.11.2026 — sveska iz `7c12aa2e` više NIJE sveska za C002.** Main
  `fd39a716` (#446–#450) menja pisca ambalaže (MIG-ODL-13..15); migracija se
  ponavlja nad svežom sveskom iz aktuelnog main-a (pregled → proba → upis).
- **Proba 01.11.2026 nad C002 (MIG-ODL-13..15):** 0 BLOKIRA; AMB-INV-07 i
  AMB-INV-16 (prazne) drže; 54 vraćanja = 54 izlaza; otpremnice −1042 gajbe
  (6 svedeno na blokove, 17 neizdatih van knjige); nabavka 11 dok. / 44391;
  5 naloga sa punim ispod nule — svi vozači, prihvaćeno 03.11.2026 (MIG-ODL-17).
- **Proba 03.11.2026 nad main-om `a668c5e6` (S7):** isti ishod kao 01.11.; sve
  ODLUKE prihvaćene (MIG-ODL-06/09/10/11/12/17).
- **Upis 03.11.2026 u svežu svesku iz `a668c5e6` (worktree, grana `b025c47c`):
  prošao.** Po MIG-ODL-18 se knjiga i dokumenti ne menjaju (iste otpremnice su nacrt kao i pre), samo nalazi nose drugo ime. 34 tabele + podešavanja; `tblAmbalaza` 5168 red., `tblAmbalazaDokument`
  82; povratno poređenje svake ćelije: 0 razlika; salda ambalaže naspram prevoda:
  0 razlika.
- **Upis 03.11.2026 posle MIG-ODL-18/19 (grana `1aa07632`, main `a668c5e6`):
  prošao** — 0 razlika, salda ambalaže 0; isti broj redova (`tblAmbalaza` 5168).
  **To je sveska za C002** (dok main ne promeni šemu ili pisca — sledeći je S8,
  palete).
- **Upis 28.10.2026 nad sveskom C002 (SHA256 `d5e1f612…`): prošao.** 34 tabele
  + podešavanja; povratno poređenje svake ćelije: 0 razlika; salda ambalaže nove
  sveske naspram prevoda: 0 razlika. Usput nađene i ispravljene tri greške
  samog upisa (pywin32 `Range.Resize`, tabela sa jednim redom, datum pomeren u
  UTC) — svaka ima test.
- **Ponovljen isti dan nad svežom sveskom iz main-a `7c12aa2e` (#445 + #441):**
  isti ishod — 0 razlika, salda 0 (zamenjeno gore).

## Usklađenost sa main-om posle #445 i #441

Provereno nad kodom 28.10.2026 (main `7c12aa2e` = #445 + #441 sa glavom
`7274598c`); šema se ni u jednom ne menja, pa `cilj.py`/`provere.py` važe bez
izmene. Sveska za C002 se pravi iz ovog main-a.

- **#445 (nabavka i otpis sa F7, AMB-10-ODL-34):** migraciona NABAVKA ima isti
  oblik kao `NabaviAmbalazu_TX` — zaglavlje vrste `NABAVKA` čiji je vlasnik broja
  nalog, red `SpoljniSvet → nalog`, `DokumentTip = AmbalazaDokument`. NABAVKA na
  **vozača** (ako je stara sveska imala vozača u minusu) ugovor dozvoljava
  (`AmbKlaseVrste`: GRANICA → SOPSTVENI), ali je ekran od ODL-34 (b) ne nudi —
  to je zapis istorije, ne tok koji se nastavlja. Nova numeracija
  (`AmbDokNizSken`) gleda samo isti nalog i isti dan, a `MIG-POC/...` čita kao
  redni broj 0, bez greške.
- **#441 (poslovna projekcija kartice, AMB-POSL-01/02):** kartica kooperanta i
  otkupni list ne broje `ULAZ_TUDJE` (`AmbVrstaJeZahtev`). Migracioni
  `VRACANJE_TUDJE` nije dodat red nego deo starog izdavanja (`primeni_politiku_duga`),
  pa je **poslovni saldo partnera u novoj svesci jednak starom saldu** — ista
  jednačina koju alat već meri (stari − van knjige + izuzeci + pokriće = novi;
  pokriće partnera je samo `ULAZ_TUDJE`).

## Usklađenost sa main-om `fd39a716` (#446–#450)

Provereno nad kodom 01.11.2026; šema se ne menja (otisak `87AB4F9B`).

- **#446 (revers kooperantu poslovno), #447 (bilans):** samo čitaoci. Bilans
  „Neusklađeno" je signal za vraćanje bez izlaza — posle MIG-ODL-13 je 0.
- **#448:** pisac deli izdavanje kooperantu (`AMB-10-ODL-35`: prvo vraća tuđe)
  i piše izlaz (`AMB-INV-15`) → MIG-ODL-13; migracija deli po dnevnom zaključku
  (MIG-16/17), što je na kraju dana isto. Kapija praznih stanice (`AMB-INV-16`)
  → MIG-ODL-15.
- **#449:** kapija praznih i za vozača (`ODL-37`), otpis kod vozača je lom
  praznih (`ODL-38`, stara sveska nema otpis), `KI-065` izmeren nad sveskom
  C002 (pune stanica ispod nule zbog otpremnica sa više gajbi od otkupa) i
  pravilo za ponovni prevod → MIG-ODL-14/15.
- **#450:** vraćanje kupcu (`ODL-39`) — novi tok, stara sveska ga nema; obaveza
  prema kupcu iz migracionog pokrića ostaje i sada se može vratiti sa F7.

## Usklađenost sa main-om `a668c5e6` (S7, #451–#454)

Provereno 03.11.2026; grana rebase-ovana na `a668c5e6`. S7 menja šemu fakture
(otisak `B22A1572`) → MIG-ODL-16. Ostalo (SEF linija, storno fakture, health,
utovar) su pisci i čitaoci nad fakturama, kojih u svesci C002 nema. Ambalaža i
robni dokumenti se ne menjaju, pa MIG-ODL-13..15 važe.

## Otvoreno

- ~~Storniran blok kao aktivan član otpremnice~~ — rešeno odlukom MIG-ODL-22.
- **Proba 03.11.2026 (MIG-ODL-18):** 16 od 17 otpremnica bez bloka ima izdatu
  zbirnu i tačno jednu prijemnicu, bez kandidata sa viškom blokova — prave, čekaju
  blokove (ukupno 2340,6 kg, 1007 gajbi; ST-00001 6, ST-00003 9, ST-00007 1).
  Gajbe stanica ST-00001 (125) i ST-00007 (7) su tačno manjak punih tih vozača
  (MIG-ODL-17), pa se on gasi kad operater izda te otpremnice. OTP-00666 nema
  prijemnicu, zbirna joj je stornirana — prenosi se kao stornirana (MIG-ODL-19). 37 izdatih otpremnica ima razliku
  kg prema blokovima (30 do 1,2 kg; veće: OTP-00289 −44,6, OTP-00329 −30,4 i
  OTP-00676 +30,4 na istoj stanici ST-00005, OTP-00424 −12,4, OTP-00174 +4,4,
  OTP-00184 −3,7, OTP-00393 −2,0) — rešava operater sa otkupcem.
- Posle upisa: operater unosi blokove otpremnica koje ih čekaju (MIG-ODL-18) i
  rešava sa otkupcem otpremnice čiji se kg ne slaže sa blokovima (spisak:
  `proba_sazetak.txt`, odeljak „Otpremnice — preostalo kg i neizdate").
- Prazna salda vozača na kraju sezone su velika (npr. ST-00002 0,4kg 3462) —
  istorija stare sveske; vredi ih proveriti popisom pre nove sezone.
- Provera u pravom Excelu nad novom sveskom (Compile, otvaranje ekrana,
  kartice, bilans „Neusklađeno" = 0, izveštaji robe, aktivacija licence) —
  NEVERIFIKOVANO dok je operater ne prođe.
- `docs/KNOWN_ISSUES.md` `KI-024` (deo 2) i `KI-065` se zatvaraju tek posle te
  provere.
- `docs/DOMEN/README.md` §1 još kaže da se otkup vezuje preko
  `OtpremnicaID`/`BrojOtpremnice`; kod ga vezuje preko `tblOtpremnicaIzvori`.
