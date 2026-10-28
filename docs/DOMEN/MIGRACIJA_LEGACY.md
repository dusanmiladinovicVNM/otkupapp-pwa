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

## MIG-WR-01 — izuzetak od vlasništva upisa

Masovni upis u **praznu** svesku iz gita je jednokratno punjenje pre rada, ne
pisac u radu. Zato je izuzetak od AMB-INV-08 (upis u knjigu samo u transakciji
izvornog dokumenta) i od A11 (`WRITE_OWNERSHIP.json`, red `tblAmbalaza` →
`modAmbalaza`). Presedan: `modMigracija.MigrirajPodatkeIzStarog`.

Uslovi koje alat drži:

- **pre upisa**: `provere.py` meri ista pravila po kojima main čita celu tabelu
  (`Stavke*Redovi`, članstvo, izdate otpremnice bez izvora, fakture). Nalaz
  BLOKIRA odbija upis; ODLUKA traži `--prihvatam-odluke`.
- **upis**: svaka tabela u koju se piše mora biti prazna; izvorna sveska se ne
  otvara u Excelu; makroi i događaji ugašeni.
- **posle upisa**: svaka ćelija se poredi sa planom, a salda ambalaže nove sveske
  sa saldima prevoda.

## Stanje

- **Upis 28.10.2026 nad sveskom C002 (SHA256 `d5e1f612…`): prošao.** 34 tabele
  + podešavanja; povratno poređenje svake ćelije: 0 razlika; salda ambalaže nove
  sveske naspram prevoda: 0 razlika. Usput nađene i ispravljene tri greške
  samog upisa (pywin32 `Range.Resize`, tabela sa jednim redom, datum pomeren u
  UTC) — svaka ima test.
- **Ponovljen isti dan nad svežom sveskom iz main-a `7c12aa2e` (#445 + #441):**
  isti ishod — 0 razlika, salda 0. To je sveska za C002.

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

## Otvoreno

- Provera u pravom Excelu nad novom sveskom (Compile, otvaranje ekrana,
  kartice, izveštaji robe, aktivacija licence) — NEVERIFIKOVANO dok je operater
  ne prođe.
- `docs/DOMEN/README.md` §1 još kaže da se otkup vezuje preko
  `OtpremnicaID`/`BrojOtpremnice`; kod ga vezuje preko `tblOtpremnicaIzvori`.
