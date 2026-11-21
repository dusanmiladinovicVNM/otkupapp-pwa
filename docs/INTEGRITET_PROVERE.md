# Integritet provere (tabela ↔ tabela)

Konsolidovana, **read-only** revizija integriteta podataka kroz ceo lanac
dokumenata: `Otkup → Otpremnica → Zbirna → Prijemnica → Paleta → Prerada`, uz
knjigu ambalaže, doplatu, gotovinu i kulturu otkupnog lista. Ništa ne menja —
samo izlistava neusklađene zapise. Kod: `src-vba/modIntegritet.bas`.

> **Od S9-5 (plan §14.69, 21.11.2026)** revizija meri **model**: ivice grafa
> sledljivosti E2, E3 i E4 (`docs/DOMEN/SLEDLJIVOST.md`, `SLED-INV-01`) istim
> sudom koji drže strogi čitači, a ne kolone bez pisca ni labelu broja zbirne
> (`SLED-INV-02`). Provere starog modela — A1, B1a/B1b, B4a/B4b, B5, B5b, B6 i
> health 7 — su **obrisane, ne prevedene** (plan §14.8 t. 11): čitale su
> `Otkup.BrojZbirne`, `Otkup.OtpremnicaID`, `Otpremnica.BrojZbirne` i labelu
> `Prijemnica.BrojZbirne`, pa su bile ili uvek tihe (A1) ili uvek glasne (B5b:
> u C002 921 nalaz = svaka aktivna otpremnica), a B5 nije video prijemnicu sa
> labelom bez veze (C002: 6 od 905). Kodovi B2, B3 i C2 ne postoje od ranije.

## Pokretanje

- **Oporavak → lista INTEGRITET** (ljuska `frmOtkupUI`, `modScrOporavak`):
  mreža `Problem | Detalj`, bez radnje (nalaz je opis neslaganja; popravka ide
  kroz svoj tok — Nedovršeno, prevezivanje, storno). Snimak se računa kad se
  lista otvori i drži do sledećeg upisa. Prvi red je **UKUPNO** (dole), a lista
  ide **redosledom revizije**: naslov bloka (kod provere), zaglavlje kolona, pa
  nalazi. Ljuska je ne sortira (`Scr_Sort` = poslovni redosled), a klik na
  zaglavlje to kaže (`Scr_SortRazlog`) — snimak je izveštaj, i složen po tekstu
  Detalja gubi vezu nalaza sa proverom (`KI-100`).
- **Admin panel** → „Integritet provere (tabele)" (`RunIntegritetProvere`) →
  upis u sheet **`INTEGRITET_PROVERE`** (obriše se i iznova napiše pri svakom
  pokretanju): svaki blok naslov + broj zapisa, pa „OK – nema" ili tabela, na
  kraju „UKUPNO" + zbirni `MsgBox`.

**Red UKUPNO** (prvi red liste) nosi broj **nalaza** istog prolaza
(`GetIntegritetRows`, izlaz `outUkupno`), ne broj redova snimka — snimak nosi i
naslov bloka, zaglavlje kolona i prazan red između blokova (`KI-099`, do S9-5 je
broj za dva nalaza u dva bloka bio „7"). Red ne prolazi kroz pretragu: „koliko ih
ima" ne zavisi od toga šta je operater ukucao. Do ručne provere S9-5 broj je
stajao u dopuni naslova, ali je naslov labela fiksne širine (180pt) i dopuna se
odsecala cela — broj je bio tačan, a nije se video; naslov liste zato dopunu
nema. **Pala provera** (red `GREŠKA: …` — čitač je odbio tabelu) nije nalaz o
podacima: ne broji se u nalaze, a red UKUPNO je kaže posebno („…, neizvršene
provere: N"); isto `MsgBox` posle Admin prolaza. **Pad cele revizije** (van
handlera pojedinačnih provera) je red `INTEGRITET` sa „GREŠKA: revizija nije
izvršena — …" i jedna neizvršena provera, nikad prazan snimak sa nulom
(`S95-REV-02`).

Sve provere isključuju stornirane redove (`ExcludeStornirano`), osim gde je
stornirani dokument baš predmet nalaza (C4, KUL1/KUL2).

## Provere (redosled kao u `RunAllChecks`)

### E — Ivice modela (`SLED-INV-01`: svaka ivica razrešava tačno jednom)

| Kod | Značenje | Sud |
|---|---|---|
| E2 | Članstvo otpremnice (`tblOtpremnicaIzvori`) ne razrešava: prazan ID, otpremnica ili otkup ne postoji tačno jednom, dupli par, otkup u dve aktivne otpremnice | `modDokumenta.ClanstvoOtpremnicaNalazi` — **isto pravilo** kao strog čitač `AktivnoOtpClanstvoPoKanonu`, samo sakupljeno red po red (`S95-ODL-01`) |
| E3 | Članstvo zbirne (`tblZbirnaIzvori`) ne razrešava: ista pravila kao E2 (od S9-5 i u čitanju, `S95-ODL-02`) | `modDokumenta.ClanstvoZbirneNalazi` ↔ `AktivnoClanstvoPoKanonu` |
| E4 | Aktivna prijemnica bez aktivne zbirne: `ZbirnaRoditeljID` prazan, zbirna ne postoji, **nije jednoznačna** (isti `ZbirnaID` u dva zaglavlja) ili je stornirana. I prijemnica **bez broja** je nalaz | `modDokumenta.PrijemnicaE4Razlog` — **jedno pravilo** za integritet, listu siročadi u Oporavku i pisca prevezivanja (`S95-ODL-07`); popravka je „Prevezi prijemnicu" |

Čitač bez kolekcije **diže** na prvom kvaru (ekrani, izveštaji i jezgro grafa
`modSledljivost` ponašaju se kao i pre); revizija isti sud **sakuplja**, pa vidi
sve pokvarene redove (obrazac `modAmbalaza.AmbKnjigaNalazi`).

**Ostale ivice:** E1 (stavke otkupa) i E5 (stavke prijemnice) drže strogi čitači
stavki — njihov kvar je **GREŠKA bloka** koji ih čita (A2, A3), ne red po red
(`KI-081`, imenovana granica). E6 (linija fakture) meri health 3 i 4, E7 (paletna
stavka) C1/C4, E8 (prerada) D1/D2, E9/E10 (utovar, GP faktura) health (GP pravilo).

### A — Konzervacija količine (kg, sa **stavki**)

| Kod | Značenje | Napomena |
|---|---|---|
| A2 | Manjak/višak `zbirna → prijemnica` **po dokumentu** (`ZbirnaID`): **VIŠAK > 5%**, **NIŠTA PRIMLJENO**, **MANJAK > 10%**, **PRIJEM NEPOZNAT** (kvar E4) | isti agregat kao izveštaj manjka i ROBA (OM): `modHelpers.BuildManjakDict` (`S95-ODL-03`, `KI-097`); `ZbirnaID` je poslednja kolona; zbirna bez broja (nacrt) se ne meri |
| A3 | `Σ paleta-stavke.NetoKg` po stavci prijemnice vs `Kolicina` te stavke | samo paletizovane; tol 0.5 kg |
| A4 | `paleta.NetoKg`/`BrojGajbica` (header) vs `Σ stavke` | tol 0.5 kg / 0.001 gajbe |
| A5 | prerada `NetoUlazKg` vs `Σ stavke.NetoKg`; `NetoIzlaz ≤ NetoUlaz` | tol 0.5 kg |

### B — Identitet i dokument

| Kod | Značenje |
|---|---|
| B7 | Aktivna zbirna bez kilaže na stavkama (kanonski pisac je ne pravi — svaki nalaz je kvar podatka) |
| B8 | Broj zbirne nosi **više aktivnih dokumenata** — presuda `ZbirnaIdentResolve`, pa dvoklasna zbirna nije nalaz. Čita **sopstvenu** labelu zbirne (ima pisca). Ostaje dok F4 i F8 ispravka prijemnice razrešavaju zbirnu po broju (`KI-098`, `S95-ODL-05`); posle toga je savet |
| B9 | Aktivna zbirna bez `ZbirnaID`-a ili sa duplim (identitet dokumenta) |
| B11 | Jedan utovar (`PredajaID`) nosi **dva** aktivna dokumenta otpremnice |

### AMB — Knjiga ambalaže

| Kod | Značenje |
|---|---|
| AMB1 | Kanonska knjiga: ugovor reda, jedinstven `AmbID`, `StornoOd` — isti sud kao čitanje knjige (`modAmbalaza.AmbKnjigaNalazi`) |

### C — Palete

| Kod | Značenje |
|---|---|
| C1 | Paletna stavka bez žive prijemnice (prazan/nepostojeći `PrijemnicaID`) ili sa pogrešnim ključem linije (`PAL-INV-01`, isto pravilo kao pisac) |
| C4 | Paletna stavka ka **storniranoj** prijemnici (kaskadni storno ne dira `tblPaletaStavka`). Kontekst reda **po ivicama** (`S95-ODL-06`): broj, kg i gajbe zbirne po E4; otkupno mesto iz otpremnica zbirne (E3; fallback mirror-stanica po `VozacID` zbirne); proizvođač iz izvora tih otpremnica (E2). Kontekst je **mek**: kvar čitača je tekst u koloni, ne pad bloka |
| C3 | Paleta (header) bez ijedne aktivne stavke |
| C5 | Dupli `BrojPalete` unutar iste godine |

### D — Prerada

| Kod | Značenje |
|---|---|
| D1 | Paleta `Preradjeno=Da` bez ijedne aktivne prerada-stavke |
| D2 | Prerada-stavka ka nevalidnoj paleti (nesveža/stornirana/nepostojeća) — očekivano prazno zbog storno kapije |

### DOP / GOT — Novac

| Kod | Značenje |
|---|---|
| DOP1 | Osnovica knjižene doplate se promenila posle knjiženja (`modDoplata.DoplataOdstupanjaOsnovice`; `docs/DOMEN/DOPLATA.md`, DOP-ODL-07) |
| GOT1 | Gotovinska isplata preko osnovice (`docs/DOMEN/ISPLATA_GOTOVINA.md`, ISP-ODL-01..03) |

### KUL — Kultura otkupnog lista (`KI-092`, `KI-093`)

| Kod | Značenje |
|---|---|
| KUL1 | Otkupni list čiji `KulturaID` nije ono što `(VrstaVoca, SortaVoca)` daje resolveru `modOtkup.RazresiKulturuIzVrsteSorte` |
| KUL2 | Verzija otpremnice/zbirne (po `IspravkaOdID`, ID-trag) sa drugim tekstom sorte od prethodne. Rezervna grana po broju-obliku traga je obrisana u S9-5 (N16) |

## Podesive tolerancije

U `modIntegritet.bas`, vrh modula:
- `PRAG_MANJAK_PCT` (10) — manjak% iznad ovoga se prijavljuje (A2).
- `PRAG_VISAK_PCT` (5) — višak do ovoga se **ne** prijavljuje (A2).
- kg tolerancije (0.5) su inline u A3/A4/A5.

## Health check (`RunProductionHealthCheck`) — šta je otišlo u S9-5

- **Provera 7** (`Check_OtkupOtpremnicaCrossZbirnaLinks`) je obrisana: poredila je
  `Otkup.BrojZbirne` sa `Otpremnica.BrojZbirne` preko `Otkup.OtpremnicaID` — tri
  kolone bez pisca, pa je uvek javljala OK. Sa njom je obrisan i BFP
  `Test_NoCrossZbirnaLinksAudit` (ista provera nad fixture-om).
- **Provera 13, deo 1** (aktivan otkup → stornirana otpremnica preko
  `Otkup.OtpremnicaID`) je obrisana iz istog razloga; deo 2 (linija fakture nad
  storniranom prijemnicom) ostaje. Članstvo otkupa u otpremnici meri E2.

## Odnos prema invarijantama

- Lanac `otkup ↔ otpremnica ↔ zbirna ↔ prijemnica` = **E2, E3, E4** (postojanje i
  aktivnost veze); količina zbirna → prijemnica = **A2**.
- Paleta mora imati živu prijemnicu i tačnu liniju = **C1**; paleta nad
  storniranom prijemnicom = **C4** (popravka: „Skini palete" / prevezivanje u Oporavku).
- „Prerađeno bez sveže palete" = **D2**.

Storno **ne kaskadira** globalno, pa viseći dokumenti i orphan stavke nastaju
legitimno — zato ova revizija postoji.
