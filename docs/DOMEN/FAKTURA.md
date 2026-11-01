# Faktura — identitet, linija, storno i SEF

> **Izvor:** pre-flight S7 (plan `docs/REFAKTOR_DOKUMENT_HEADER_STAVKE.md` §14.57,
> `docs/STANJE_REFAKTORA.md` stavka 251), odluke operatera od 01.11.2026. Mereno nad
> `main` = `fd39a716`. Linije koda u ovom fajlu važe za taj commit.
>
> Pravila linije fakture nad prijemnicom (`FAK-LIN-01..03`) su starija i stoje u
> `DOCUMENT_HEADER_LINES.md` §6.6. Ovde se ne prepisuju.

## 1) Šta je već odlučeno (ne otvara se ponovo)

| Pravilo | Gde |
|---|---|
| Stavka prijemnice → linija fakture **1:1, puna količina**; stavka na najviše jednoj aktivnoj liniji | `DOCUMENT_HEADER_LINES.md` §3, §6.6 (`FAK-LIN-01`) |
| Linija nosi `PrijemnicaStavkaID`; fakturisanost se izvodi iz aktivnih linija, keša nema | §6.6 (`FAK-LIN-02`, `FAK-LIN-03`), plan §14.55 „Korak 5" |
| Fakturisana (i delimično fakturisana) prijemnica se ne stornira i ne ispravlja; put je storno fakture, pa prijemnice | plan §14.55 „Korak 4", odluka 1; `Test_S6K4_FakturisanaSeOdbija` |
| Datum fakture (dan izdavanja) i datum prijema su dva datuma | plan §14.55 „Korak 5", BUSINESS EVENTS |
| Ispravka fakture = storno + nova faktura (nov ID, nov broj); F8 za fakturu nudi samo običan storno | `modStornoDok.TipUFlowDoc:513` (faktura nije tip okvira) |

## 2) Identitet

**`FAK-IDENT-01` — identitet fakture je `FakturaID`; `BrojFakture` je labela.**
Mutacija, storno, SEF i novac rade po `FakturaID`-u. Broj je oblik `n/gggg`; vlasnik
niza je firma, kontekst je kalendarska godina, a stornirane fakture se broje, pa
storno ne oslobađa broj (`modFaktura.GenerateBrojFakture:588`). Spolja broj stiže na
dva mesta i tamo se razrešava **jednom** u `FakturaID`:

- SEF ga dobija kao `InvoiceNumber` (izlaz, ne ulaz);
- izvod ga nosi u pozivu na broj. Jak ključ važi samo kad je pogodak jednoznačan
  (`modBankaMapiranje.TryResolveFakturaByPoziv:3304`, `hits = 1`). Slabi put (kupac
  iz imena ili mape, pa faktura) poredi podniz — `KI-075`.

Posle izbora reda nijedan pisac ne traži fakturu po broju. Jedini izuzetak danas je
`modStornoDok.StornoIzvrsi` bez `docID`-a (`:416`, `KI-041`); briše ga S7 rez 1.

**`FAK-IDENT-02` — linija fakture nad prijemnicom.** Kanonski ključ je
`PrijemnicaStavkaID`. `PrijemnicaID` je dokument: labela za čitaoce i ključ za
činjenice zaglavlja (vrsta, sorta, datum, kupac, storno). `BrojPrijemnice`, `Klasa`,
`Kolicina` i `Cena` na liniji su **snapshot** (§4.5): papir i ponovna štampa ih čitaju
takve kakve su bile pri izdavanju, i nijedan nije ključ. Izmereno: nijedno
produkciono mesto ne vezuje fakturu za prijemnicu preko broja (plan §14.57,
„Merenje").

## 3) Odluke S7 (operater, 01.11.2026)

**`FAK-ODL-01` — SEF linija nosi stavku i klasu.** `SellersItemIdentification/ID` linije
nad prijemnicom je `PrijemnicaStavkaID`, a naziv dobija klasu
(„<vrsta> <sorta> kl. <klasa> po prijemnici <broj>"). JSON zapis slanja dobija
`PrijemnicaStavkaID`. GP linija zadržava `PreradaID`. Razlog: danas su dve klase iste
prijemnice na e-fakturi dve linije istog naziva i istog ID-a dokumenta
(`modSEFMapper:268-276`, `:783-789`), a papir ih razlikuje po klasi
(`modFaktura.PrintFaktura:746-747`). Papir i e-faktura posle S7 nose isti sadržaj.

**`FAK-ODL-02` — lokalni storno fakture se odbija dok faktura živi na SEF-u.**
Faktura „živi na SEF-u" kad je `SEFWorkflowState` jedno od `SEF_SENDING`, `SEF_SENT`,
`SEF_ACCEPTED`, `SEF_SYNC_ERROR`, `SEF_UNKNOWN`, `SEF_TECH_FAILED`, **osim** kad
spoljni `SEFStatus` po `modSEFStatusSync.ClassifySEFExternalStatus` pada u klasu
`STORNO` ili `TERMINAL` (otkazan / obrisan). Prazno, `LOCAL_*`, `SEF_READY`,
`SEF_REJECTED` i `SEF_STORNO` puštaju storno kao danas.

- Izuzetak za otkazanu fakturu nije proširenje odluke nego njen uslov: otkaz na SEF-u
  ostavlja lokalni workflow u `SEF_SENT` (`docs/SEF_LIFECYCLE_MANUAL.md` §6, nema
  lokalnog `CANCELLED`). Bez izuzetka bi takva faktura ostala zauvek nestornirljiva.
- `SEF_TECH_FAILED` se odbija (dopunska odluka istog dana): 409 CONFLICT završava kao
  tehnički neuspeh, a dokument možda postoji na SEF-u (AUD-030, `modSEFService:364`).
  Put je „Osveži status"; kad pokaže da dokumenta nema ili je otkazan, storno ide.
- Kapija stoji u jezgru `modStorno.StornoFaktura` (pre prve mutacije), a F8 razlog i
  uvid pitaju isti predikat. SEF storno (`sfstorno`) i dalje ne stornira lokalno:
  posle njega lokalni storno ide jer je stanje `SEF_STORNO`.

**`FAK-ODL-03` — linija je aktivna ako i samo ako je njena faktura aktivna.**
`tblFakturaStavke` gubi `Stornirano` i `OsirocenoOd`, a `tblFakture` gubi
`OsirocenoOd`. Storno fakture više ne piše linije. Time važi
`DOCUMENT_HEADER_LINES.md` §6.5 („storno je dokument-level") i za fakturu, a uslov iz
`FAK-LIN-02` („aktivna linija nad storniranom fakturom je KVAR") postaje nemoguć po
konstrukciji. `OsirocenoOd` od S6 koraka 4 niko ne piše, a grana koja ga čita je
nedostižna: fakturisana prijemnica se ne stornira (`Test_S6K4_FakturisanaSeOdbija`
meri tu kapiju). Važi i za GP linije (`modUtovar`). Storno GP fakture i dalje oslobađa utovar
(`modStorno.ReleaseUtovarFromFaktura:1692`) u istoj transakciji: linija se ne piše, marker utovara
da (`S7-REV-01`, plan §14.57).

**`FAK-ODL-04` — saldo kupca ostaje robni.** C-003 i C-004 ostaju vrednost stavki
prijemnica minus uplate kupca (prevedeno u S6 koraku 3, `modIzvestaj.ReportSaldoKupci:2687`).
„Kartica kupca" iz tabele slajsova nije postojeća sposobnost (KARTICA postoji samo za
kooperanta, C-017), pa se u S7 ne gradi; nova sposobnost ide posle refaktora, po
pravilu §14.8 t. 8.

**`FAK-ODL-05` — uplata odvezana stornom fakture je avans kupca.** Storno fakture
prazni `FakturaID` na njenim uplatama (`modStorno.ResetNovacFakturaLink:1735`).
Čitač avansa kupca (`modNovac.ApplyAvansToFaktura`, danas samo `Tip = KupciAvans`,
`:628`) prima i `KupciUplata` bez `FakturaID`-a, pa se uplata sama prebija na sledeću
fakturu kupca (`CreateFaktura` → `ApplyAvansToFaktura`). Analog odluke „D suženo" za
kooperanta (`ODLUKA_NOVAC_PRI_STORNU.md` §0): oba tipa su isti novac kupca, razlika
je samo da li je faktura bila poznata u trenutku uplate. Delimično prebijanje ide postojećim
splitom (`modNovac:660-692`): ostatak ostaje na izvornom redu, a split nosi marker izvoda
(`BuildAvansSplitNapomena`).

**`FAK-ODL-06` — odvojeni deo uplate je `KupciAvans`, a PWA broji svaku uplatu vezanu za
fakturu (operater, 01.11.2026, pre koda S7 reza 1).** Kad odvezana uplata pokrije samo deo nove
fakture, `ApplyAvansToFaktura` deli red. Primenjeni deo je nov red tipa `KupciAvans`, sa
`FakturaID`-om nove fakture i BIM markerom izvoda. Izvorni red zadržava svoj tip (`KupciUplata`),
umanjen iznos i marker. PWA izvoz plaćenog po fakturi (`modStammdatenSync.UplacenoPoFakturi`)
broji svaku nestorniranu uplatu vezanu za fakturu, bez obzira na tip — isto pravilo kao
`modNovac.GetUplataForFaktura`. Do S7 reza 1 je brojao samo `KupciUplata`, pa je avans primenjen
na fakturu u PWA izgledao kao neplaćen deo.

Izmereno pre odluke: `KupciUplata` bez fakture nastaje samo stornom, jer je oba pisca upisuju
samo uz fakturu (`modNovacUnos.UplataValidiraj`, `modBankaMapiranje.MapBankaImportAsKupac`).
Zato proširen čitač avansa tačno odgovara `FAK-ODL-05`. Tip u ovom toku čita samo taj izvoz;
zbir po kupcu (`ExportSaldoKupci`) broji oba tipa.
