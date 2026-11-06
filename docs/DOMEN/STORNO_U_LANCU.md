# Storno i ispravka otkupnog lista u lancu (`LNC-*`)

> **Status: PRE-FLIGHT 06.11.2026 — odluke operatera zapisane. U kodu: deo 1 (šema
> `LNC-ODL-16`, ambalaža `LNC-ODL-15` / `AMB-10-ODL-40`), deo 2 (jezgro storna,
> `modDokumenta.LncPlanLista` / `StornoListaULancu_TX`) i deo 3 (ispravka,
> `IspravkaListaULancu_TX` nad `modOtkup.IspravkaOtkupaUTransakciji`, ruta u
> `modOtkupUnos`), NEVERIFIKOVANO u Excelu.**
> Mereno čitanjem nad `main` = `1caa86f`; Excel nije pokrenut, pa je svaka tvrdnja o
> ponašanju ovde **NEVERIFIKOVANA** dok je ne potvrdi test. Verdikt po osama je §7.
>
> Ovaj fajl je **spec za rez**, ne opis postojećeg koda. Kad se kod i ovaj fajl
> raziđu posle reza, kod je u pravu i ovaj fajl je bug.

## 0) Problem

Otkupni list koji je u **izdatoj** otpremnici, a ona u **aktivnoj** zbirnoj, danas se
ne može ni stornirati ni ispraviti bez ručnog rasklapanja lanca:

| Kapija | Gde |
|---|---|
| storno lista u aktivnoj otpremnici se odbija | `modStorno.StornoOtkup` (`modStorno.bas:143`), isti razlog na F8 pre potvrde (`modStornoDok.bas:204`) |
| ispravka lista u izdatoj otpremnici se odbija | `modOtkup.IspravkaPrepreka` (`modOtkup.bas:306`) |
| storno otpremnice koja je izvor aktivne zbirne se odbija | `modStorno.StornoOtpremnica` (`modStorno.bas:241`) |
| zbirna nema ispravku, F8 nudi samo DUPLI / PONIŠTENJE / REŠI KASNIJE | `modScrStorno.AkcijeRacun` (`modScrStorno.bas:794`), test `T_Zbirna_NemaIspravku` |

Ručni put je osam koraka na četiri ekrana (F8 zbirna DUPLI → F2 „Ispravi" → ukloni list
→ ručno umanji očekivanje → izdaj → F8 storno lista → F3 nova zbirna → Oporavak
„Prevezi prijemnicu"), sa prozorom u kome aktivne zbirne nema. Uz to, prema čitanju
koda, korak 2 pada na `AMB-INV-07` čim je vozač gajbe predao dalje (§3, `LNC-ODL-15`).

Pravilo koje rez sprovodi **već je odlučeno**: ZBR-KANON-03 (`README.md` §2), A13
(`ARCHITECTURE_CONTRACT.md`), scenariji H1/H2 (`GOLDEN_SCENARIJI.md` §12). Ovde se
zapisuje kako, i šta je operater odlučio tamo gde pravilo ćuti.

## 1) Odluke operatera (06.11.2026)

| ID | Odluka | Obrazloženje / izvor |
|---|---|---|
| **`LNC-ODL-01`** | Storno i ispravka lista u lancu su **jedan potez i jedna transakcija**: list, nova verzija otpremnice, nova verzija zbirne i prevezivanje prijemnica — sve ili ništa | A13 „Redosled sam po sebi nije dovoljan — cela korekcija je JEDNA transakcija"; ZBR-KANON-03 |
| **`LNC-ODL-02`** | Aktivne prijemnice stare zbirne se **prevezuju** na novu zbirnu. Količine, stavke, `PrijemnicaID`, palete i faktura se **ne diraju** | A13 „propagacija staje pred prvom nezavisnom činjenicom": prijemnica je kupčevo merenje. H1 „prijemnica se ne dira" = ne menja se njena količina; veza ide na aktivnu verziju, inače prijemnica postaje siroče |
| **`LNC-ODL-03`** | Očekivanje nove otpremnice se **izvodi**: staro − list (+ novi list kod ispravke), po klasi, za kg i gajbe | Stara je bila izdata, a izdavanje je dokazalo očekivano = povezano po klasi (`OtpIzdaj`, `modDokumenta.bas:4923`). Umanjenje za jedan poznat list nije „izjednačavanje" koje je S3b-2a odbio |
| **`LNC-ODL-04`** | Novac kao i danas: storno odvezuje (`ResetNovacOtkupLink`, virman postaje slobodan avans); ispravka prenosi na nov list i prijavljuje preplatu (PUT A) | `ODLUKA_NOVAC_PRI_STORNU.md` (A + D suženo); bez novog koda za novac |
| **`LNC-ODL-05`** | Ulaz za storno: **F8 nad listom**. Umesto današnjeg odbijanja nudi radnju sa uvidom posledica pre potvrde | F8 već ima obrazac „prvo vidi posledice, pa odluči" (`GradiUvid`, `modStornoImpact`) |
| **`LNC-ODL-06`** | Obim prvog reza: **storno + ispravka** lista u lancu. Premeštanje lista na drugu otpremnicu je naredni rez, istom mehanikom | — |
| **`LNC-ODL-07`** | Poslednji izvor: otpremnica bez preostalih listova se **samo stornira** (bez naslednika); zbirna bez preostalih otpremnica isto. **Odbija se** kad bi aktivna prijemnica ostala bez zbirne — tada F8 PONIŠTENJE nad zbirnom | `GOLDEN_SCENARIJI.md` §10 t. 3 („prazna nova verzija nema smisla"); otpremnica bez izvora nije isporuka (`OtpIzdaj`) |
| **`LNC-ODL-08`** | Fakturisana prijemnica se **prevezuje** kao i nefakturisana. Faktura i SEF se ne diraju | `PrijemnicaID` se ne menja, a linija fakture vezuje stavku prijemnice (`FAK-LIN-*`); ni faktura ni SEF ne vezuju zbirnu |
| **`LNC-ODL-09`** | **Nema automatske štampe.** Poruka posle poteza imenuje stari → novi broj otpremnice i zbirne | Danas ni izdavanje ne štampa samo; papir ide iz Izveštaja |
| **`LNC-ODL-10`** | Ispravka koja menja **stanicu, kulturu ili tip ambalaže** se **odbija**, sa porukom koja upućuje na storno lista u lancu i nov list | Ispravljen list ne može u istu otpremnicu: `OtpRequireIzvorValjan` proverava baš ta tri |
| **`LNC-ODL-11`** | Ispravka u lancu traži **potvrdu pri snimanju**, sa imenovanim posledicama (otpremnica, zbirna, prijemnice, novac). „Ne" = ništa nije upisano, forma ostaje popunjena | Ispravka se i danas upisuje tek pri snimanju (`OtvoriIspravkuBloka`, B-040) |
| **`LNC-ODL-12`** | Ulaz za ispravku: postojeće **F1 „Ispravi blok"** | `modScrDokumenti.OtvoriIspravkuBloka` (`modScrDokumenti.bas:355`) |
| **`LNC-ODL-13`** | **Nacrti se menjaju u mestu** (A14/A15): list izlazi iz nacrta otpremnice, odnosno otpremnica se u nacrtu zbirne menja novom verzijom, bez storna i bez novog broja. **Očekivanje nacrta se ne dira** | Očekivanje nacrta je operaterova prijava o kamionu, nezavisna od listova; „preostalo" poraste i vidi se u panelu, a izdavanje drži postojeća kapija jednakosti |
| **`LNC-ODL-14`** | **Hladnjački lanac je van poteza**: radnja ga odbija i upućuje na F8 PONIŠTENJE (pa storno lista) | `HLD-LANAC-01`: jedan blok je ceo lanac 1:1; prepoznaje ga `modAutoHladnjaca.OtpremnicaJeHladnjacka` |
| **`LNC-ODL-15`** | Ambalaža: `AMB-INV-07`, `AMB-INV-09` i `AMB-INV-16` se za ovaj potez mere **posle cele zamene**, nad svim nalozima koje je potez dotakao (neto), a ne posle svakog kontra-stava | Vidi §3 i `AMBALAZA.md` 6.14c (`AMB-10-ODL-40`) |
| **`LNC-ODL-16`** | `tblZbirna` dobija **`IspravkaOdID` i `ZamenjenSaID`, na kraju kanona** (`schema/schema.json` → `gen_schema_module.py`) | A9: veza verzija je po ID-u. `tblOtkup` i `tblOtpremnica` ih imaju, `tblZbirna` ne (samo broj-oblik `IspravkaOd`/`ZamenjenSa`, koji S9 briše, `SLED-ODL-08`) |
| **`LNC-ODL-17`** | Broj nove verzije **uvek iz generatora**, i kad je auto-broj isključen: otpremnica `GenerateBrojOtpremnice(stanica, datum)`, zbirna niz `KIND_ZBR` (vozač, datum), bez udaljene provere | A9 (nov broj iz istog niza). Prekidač auto-broja određuje **predlog u formi**; `SuggestNextBroj` ga poštuje (`modBrojevi.bas:64`), a `GenerateBrojOtpremnice` ne. Bez ovoga bi nova zbirna ostala bez broja |
| **`LNC-ODL-18`** | **Poništavanje storna (UNDO) se odbija** za storno lista koji je deo poteza u lancu. Greška se ispravlja novim potezom, ne vraćanjem dela | Žurnal pamti samo otkup i novac (`modStornoZurnal.PkColForTable`, `:326`); otpremnicu i zbirnu ne. Vraćen stari list bi stajao van svake otpremnice: izgubljen blok, nova malina auto-otpremnica, a kod ispravke dupli zbir |
| **`LNC-ODL-19`** | Zbirna sa vozačevog PWA: potez **prolazi**, a poruka kaže da telefon prikazuje stari broj zbirne. Upis novog broja na telefon je poseban rez | Povratni upis u VOZ sheet ide jednom, pri uvozu (`modMasterSync.WriteBackVOZSyncStatus`); ponovno slanje stare zbirne ostaje duplikat (`ZbirnaPoClientRecordID`), ne pravi novu |
| **`LNC-ODL-20`** | Korekcija = **red po paru u `tblStornoVeze`** (list, otpremnica, zbirna — svaki svoj `CorrectionID`), vezani **zajedničkim korenom**: svi redovi poteza nose `ParentDocType = LANAC`; red lista je koren (`ParentDocID` prazan), a redovi otpremnice i zbirne nose `ParentDocID` = `CorrectionID` korena. Dokument nosi `CorrectionID` svog para | `tblStornoVeze` drži jedan par po ključu (`CreateCorrectionContext`), a S9 istoriju čita par po par (`SLED-ODL-03`). **Precizira A13:** „isti CorrectionID" znači „isti koren korekcije" |

Izvedeno iz postojećih pravila, bez nove odluke:

- **`KupacHladnjaca` nove zbirne se nasleđuje sa stare** (ZBR-OWN-01: oznaka je činjenica
  nastanka lanca). Pisac je danas računa iz trenutnog configa
  (`modAutoHladnjaca.OznakaKupacHladnjaca`, `modDokumenta.bas:1476`), pa verzija mora da
  je prenese. Inače bi promena configa između dve verzije prevrnula pravilo PONIŠTENJA.
- **`PredajaID` se prenosi, `ClientRecordID` ne** (A13; presedan `OtpIspravi`). Bez
  `PredajaID` zakasneli PRED red istog utovara pravi drugu otpremnicu.
- **Datum nove verzije = datum stare.** Kontra-stav ambalaže nosi datum originala, izveštaji
  filtriraju po datumu, a broj nosi dan.

## 2) Invarijante poteza (`LNC-INV-*`)

| ID | Tvrdnja | Vlasnik (gde se drži) |
|---|---|---|
| **`LNC-INV-01`** | **Atomičnost:** pad bilo koje kapije ostavlja svih ~13 tabela kakve su bile (otkup i stavke, otpremnica, stavke i izvori, zbirna, stavke i izvori, prijemnica, ambalaža, novac, `tblStornoVeze`, `tblStornoZurnal`) | koordinator (jedan snapshot); presedan veličine: PONIŠTENJE zbirne snima 11 tabela |
| **`LNC-INV-02`** | **Redosled (A13):** stara pada pre nove, odozgo: zbirna → otpremnica → list; nove nastaju odozdo: (nov list) → otpremnica' → zbirna' → prijemnice | kapije jezgra već odbijaju obrnut redosled (`modStorno.bas:143`, `:241`, `AktivnoClanstvoPoKanonu`) |
| **`LNC-INV-03`** | **Istorija nepromenjena:** stara otpremnica i stara zbirna zadržavaju svoje članstvo i stavke (A15, H1 „stari lanac rekonstruišljiv") | `tblOtpremnicaIzvori` / `tblZbirnaIzvori` su bez storna (`BEZ_STORNA`) |
| **`LNC-INV-04`** | **Sestre ostaju:** otpremnice zbirne koje potez ne dira ulaze u zbirnu' nepromenjene (H2) | pisac zbirne iz izvora |
| **`LNC-INV-05`** | **Trag po ID-u na oba kraja** (`IspravkaOdID` na novoj, `ZamenjenSaID` na staroj) za otpremnicu i zbirnu; broj-oblik se **ne piše** (`SLED-ODL-08`). `ZamenjenSaID <> "" ⇒ Stornirano = Da` (A13) | `modDokumenta` (vlasnik `tblOtpremnica`, `tblZbirna`) |
| **`LNC-INV-06`** | **Očekivanje otpremnice' = povezano otpremnice'** po klasi, kg i gajbe, i to dokazuje **isti** `OtpIzdaj`; stavke zbirne' se izvode iz izvora (ulaz bez nezavisnog očekivanja) | `OtpIzdaj` / `ZbirnaUpis` |
| **`LNC-INV-07`** | **Prijemnice se biraju po ID-u:** tačno skup aktivnih prijemnica sa `ZbirnaRoditeljID = stara ZbirnaID`. Svaka se prevezuje oba polja zajedno (`PoveziDeteNaZbirnu`), uz PRJ-ZBR-01 (cilj izdat) i PRJ-ZBR-02 (ista roba). Aktivna prijemnica sa **labelom** stare zbirne a bez `ZbirnaRoditeljID` → **odbijanje** | `modDokumenta` (vlasnik `tblPrijemnica`). Današnji izbor dece kreće od labele (`ActivePrijIDsByZbirna`, `modStornoFlow.bas:1904`) — ⚠ IDENTITY RISK, ne koristi se |
| **`LNC-INV-08`** | **Ambalaža neto:** posle cele zamene nijedan realan nalog nije u minusu, nijedna partnerska obaveza nije negativna, a pokriće praznih stanice važi (`AMB-INV-07/09/16`) | `modAmbalaza` (`LNC-ODL-15`) |
| **`LNC-INV-09`** | **Nijedan ostatak za automatiku:** posle poteza nijedan aktivan list iz stare otpremnice nije van aktivne otpremnice, i nijedna aktivna izdata otpremnica iz stare zbirne nije van aktivne zbirne | inače ga sledeći sync obuhvati: malina auto-otpremnica, auto-zbirna (`NevezaneOtpremnice`) ili hladnjački auto-lanac |
| **`LNC-INV-10`** | **Korekcija je završena u istoj transakciji:** nijedan red `tblStornoVeze` ovog poteza ne ostaje `PENDING` | pending kontekst curi u Nedovršeno i u poruke F2/F3 („ispravka pauzirana", `modDokUnos.bas:352`) |
| **`LNC-INV-11`** | **UNDO ne vraća deo poteza** (`LNC-ODL-18`) | `modStornoZurnal` (kapija u `UndoOperation_TX` / `GetUndoableStornoOperations`) |

## 3) Ambalaža — zašto se meri na kraju (`LNC-ODL-15`)

Izdavanje otpremnice knjiži `Stanica → Vozač` za gajbe stavki (`OtpKnjiziAmbalazu`).
Storno upisuje kontra-stav i **odmah** meri `AMB-INV-07`
(`modAmbalaza.StornirajAmbalazuDokumenta`, `modAmbalaza.bas:2680`; komentar `:2510`:
„AMB-INV-07 se meri kao i do sada, posle kontra-stava").

Posle prijemnice je vozač gajbe već predao kupcu (`Vozač → Kupac`), a kupac ne vraća
uvek sve prazne (ili ih je vozač već vratio stanici). Kontra-stav stare otpremnice
(`Vozač → Stanica`, N) tada traži N gajbi koje vozač više nema, pa **pada pre nego što nova otpremnica proknjiži
N − k**. Novu nije moguće izdati prvu: izvor sme biti u tačno jednoj aktivnoj otpremnici,
a izvor aktivne zbirne se ne stornira. PONIŠTENJE taj problem nema samo zato što prijemnicu
stornira prvu (`modStornoFlow.bas`, AMB-10-ODL-21) — ovaj potez je ne stornira.

Neto efekat storna lista sa k gajbi je tačan: stanica 0, vozač −k, kooperant +k. Kad su
podaci dosledni (prijemnica meri ono što je kamion stvarno nosio), vozač na kraju drži bar
k, i neto merenje prolazi. Kad ne drži, odbijanje na kraju je **tačan** signal: dokumenti
tvrde da je kupac primio robu koje po stornu nije bilo.

Presedan za odloženo merenje već postoji: `odloziObavezu` odlaže `AMB-INV-09` u ispravci
otkupa (`modAmbalaza.bas:2505`) iz istog razloga — međukorak zamene ne sme da obori
operaciju čije je posle-stanje ispravno.

> **NEMERENO.** Tvrdnja „današnje F2 Ispravi pada kad je vozač predao gajbe" izvedena je
> čitanjem. Prvi test reza je baš taj crveni slučaj nad današnjim kodom (§8, T1).

## 4) Matrica stanja — šta potez radi

| List je u… | Otpremnica | Zbirna | Potez |
|---|---|---|---|
| nijednoj otpremnici | — | — | van obima: današnji storno / ispravka |
| otpremnici | **nacrt** | (nacrt otpremnice ne ulazi u zbirnu, `modDokumenta.bas:2154`) | u mestu: list izlazi (storno) ili se menja (ispravka) u članstvu nacrta; očekivanje se ne dira (`LNC-ODL-13`) |
| otpremnici | **izdata** | nema | otpremnica' (izvedeno očekivanje) + list |
| otpremnici | **izdata** | **nacrt** | otpremnica' + u nacrtu zbirne stara otpremnica → otpremnica' (u mestu) |
| otpremnici | **izdata** | **izdata** | otpremnica' + zbirna' + prijemnice prevezane |
| hladnjački lanac | — | — | odbija se → F8 PONIŠTENJE (`LNC-ODL-14`) |
| poslednji list / poslednja otpremnica | — | — | bez naslednika (`LNC-ODL-07`) |

`PROSLEDJENO` se ponaša kao `IZDATO` (`IzdatoStatusJeIzdato`; danas ga ne piše nijedan put).

## 5) Poslovni događaji

| Događaj | Šta se dešava | Datum | Status / tabela |
|---|---|---|---|
| **fizički** | **ništa novo.** Kamion je već otišao, kupac je već merio. Potez ispravlja **zapis** o tome: list nije postojao (dupli unos) ili je pogrešno izmeren | ostaje datum originala na svim verzijama | — |
| **poslovni / pravni** | nove verzije otpremnice i zbirne, novi brojevi iz istog niza (A9), stare stornirane sa `ZamenjenSaID`; kod ispravke i nov list. Prijemnica (kupčev dokument) ostaje, menja joj se samo roditelj | datum stare verzije | `IZDATO` (kao stara); `tblOtpremnica`, `tblZbirna`, `*Izvori`, `*Stavke`, `tblStornoVeze` |
| **finansijski** | dug firme prema kooperantu se menja (storno: nestaje; ispravka: novi iznos, preplata se prijavljuje). Potraživanje od kupca se **ne** menja: faktura i prijemnica ostaju. Zbirna i otpremnica nemaju finansijski efekat | datum lista | `tblOtkup`, `tblNovac` (`LNC-ODL-04`); ambalaža neto (`LNC-ODL-15`) |

Nijedan od tri događaja ne deli datum ili status sa drugim zbog lakše implementacije:
fizički događaj se ne ponavlja, a dokument nasleđuje datum svog fizičkog događaja.

## 6) Nizvodno — ko kasnije zavisi od onoga što potez menja

Sweep po kategorijama (06.11.2026). Stavke sa „✔" su ponovo pročitane u kodu; ostale su iz
sweep-a i proveravaju se u planu reza.

| Potrošač | Posle poteza | Uslov |
|---|---|---|
| prijemnica: dve veze (`BrojZbirne` labela + `ZbirnaRoditeljID`) ✔ `PoveziDeteNaZbirnu` | labelu čitaju lista siročadi (`GetOsirocenePrijemnice`), B1b, A2, izveštaji manjka, PWA status „primljeno", F4 kapija duplikata; ID čitaju hladnjački „Nedovršeno" i papir paletnog lista | **oba polja zajedno** (`LNC-INV-07`). Bez labele: siroče u Oporavku; bez ID-a: „Ponovi auto-lanac" pravi drugu prijemnicu |
| malina auto-otpremnica, auto-zbirna, hladnjački auto-lanac | obuhvate svaki aktivan list ili otpremnicu ostavljenu van nove verzije | `LNC-INV-09` |
| PRED uvoz (`OtpremnicaPoPredaji`) | traži 0 ili 1 aktivnu otpremnicu po `PredajaID` | otpremnica' nosi `PredajaID` ✔ (`OtpIspravi` presedan) |
| PWA zbirna (VOZ): uvoz po `ClientRecordID` | ponovno slanje stare zbirne = duplikat | `ClientRecordID` se **ne** prenosi (A13); telefon pokazuje stari broj (`LNC-ODL-19`) |
| izvoz vozaču / otkupcu (`modStammdatenSync`) | prepisuje se pri svakom sync-u iz aktivnog članstva | bez izmene |
| žurnal i UNDO ✔ | pamti samo otkup i novac | `LNC-ODL-18` / `LNC-INV-11` |
| `tblStornoVeze` ✔ | jedan par po ključu; nezavršen kontekst curi u Nedovršeno i F2/F3 | `LNC-ODL-20`, `LNC-INV-10`; spoljna transakcija snima `tblStornoVeze` (presedan `IspravkaOtkupa_TX`) |
| integritet A2 (zbirna kg ↔ prijemnica kg, po labeli) | manja zbirna' uz nepromenjenu prijemnicu prijavi VIŠAK | **očekivano**: A13 traži da razlika ostane vidljiva |
| izveštaji robe i manjka | isključuju stornirano, spajaju po labeli i datumu | labela prevezana, datum isti |
| papir otpremnice „ispravka od" | čita broj-oblik, koji se ne piše (`SLED-ODL-08`) | ostaje prazan, kao danas posle F2 „Ispravi"; S9 prebacuje papir na `tblStornoVeze` |
| faktura, SEF, palete, utovar | vezuju prijemnicu / stavku prijemnice, ne zbirnu ni otpremnicu | bez izmene (`LNC-ODL-08`) |
| strogi čitači članstva ✔ `AktivnoClanstvoPoKanonu` | dva aktivna članstva iste otpremnice i storniran izvor u aktivnoj izdatoj zbirnoj su tvrda greška | redosled `LNC-INV-02` |
| doplata (signal DOP1) | prijavi storniran ili ispravljen list u već obračunatom periodu | postojeći signal, bez izmene |

## 7) Verdikt po osama (pre-flight skill §1)

| Osa | Status | Dokaz |
|---|---|---|
| **DOMAIN** | **PROVEN** | ZBR-KANON-03, A13, A14, A15, H1/H2, D1 §10 t. 3; odluke `LNC-ODL-01..20`; matrica §4 |
| **IDENTITY** | **PROVEN uz uslov** | `OtkupID` putuje od reda F8 do pisca (`modStornoDok.bas:204`, `:358`), od F1 kroz `mIspravkaOtkupID`; otpremnica i zbirna iz kanonskog članstva (`OtpremnicaZaOtkup`, `AktivnaZbirnaZaOtpremnicu`). **Uslov:** prijemnice po `ZbirnaRoditeljID`, ne po labeli (`LNC-INV-07`); trag po ID-u (`LNC-ODL-16`, `-20`) |
| **CARDINALITY** | **PROVEN** | list ∈ 0..1 aktivna otpremnica (`OtpRequireIzvorValjan`); otpremnica ∈ 0..1 aktivna zbirna (`AktivnoClanstvoPoKanonu`, dva = greška); nacrt otpremnice ∉ zbirna (`:2154`); **zbirna → 0..N prijemnica** van hladnjače — S9 sonda nad C002: 5 od 894 zbirnih ima dve (PR #469, merenje 5); hladnjača tačno 1 |
| **INVARIANTS/OWNER** | **PROVEN** | `LNC-INV-01..11`, svaka sa vlasnikom (§2); `AMB-10-ODL-40` u `AMBALAZA.md` |
| **WRITERS** | **PROVEN uz uslov** | `WHO_WRITES.md` / `WRITE_OWNERSHIP.json`: otpremnica, zbirna, izvori, stavke i prijemnica → `modDokumenta`; otkup → `modOtkup`, `modDokumenta`, `modStornoFlow`; `tblStornoVeze` → `modStornoContext`; novac → `modNovac` (+ `modStorno`); ambalaža → `modAmbalaza`. Zastavicu `Stornirano` piše `modStorno.MarkRowStornirano` kroz promenljivu (mapa je ne vidi — imenovano u `WHO_WRITES.md`). **Uslov:** koordinator u `modDokumenta`, sve ostalo kroz API vlasnika → nijedan nov pisac (`who_writes.py --check-ownership` zelen) |
| **DOWNSTREAM** | **PROVEN** | §6; svaka stavka ima uslov ili odluku (`LNC-ODL-18/19`, `LNC-INV-07/09/10`) |
| **CAPABILITY** | **PROVEN** | nova: F8 nad listom u lancu (danas odbijanje); proširena: F1 „Ispravi blok" (danas odbijanje); **ostaje netaknuto:** F8 zbirna DUPLI/PONIŠTENJE/REŠI KASNIJE (`T_Zbirna_NemaIspravku` i dalje važi — premisa testa se ne menja), F2 „Ispravi", Oporavak „Prevezi" |
| **ACCEPTANCE CONTRACT** | **PROVEN (plan)** | §8 |
| **PLATFORM** | **N/A** | nema novog Excel/VBE/MSForms ponašanja: dijalog ide kroz postojeći seam `modOtkupUI.Dijalog`, F8 kroz `StornoIzvrsiMod`. Cena snapshot-a ~13 tabela **nemerena** (presedan 11) |
| **LANDING** | **RIZIK — serijski, ovaj rez pre S9** | §9 |

**⚠ ARCHITECTURAL EDGE CASE** (zatvoren odlukom): ambalaža u međukoraku (`LNC-ODL-15`).
**⚠ IDENTITY RISK** (zatvoren invarijantom): izbor prijemnica po labeli (`LNC-INV-07`).
**⚠ FALSE-GREEN RISK:** test nad fixture-om starog oblika (`Otkup.OtpremnicaID`, kolone bez
pisca — S9 nalaz N3) bi merio tautologiju; testovi reza grade lanac **kanonskim piscima**.

## 8) Acceptance contract — plan dokaza (kod još ne postoji)

**Šta važi kad je gotovo:** potez iz §4, sa invarijantama §2, kroz F8 (storno) i F1
(ispravka), jedan potez i jedna transakcija.

**Šta ostaje netaknuto:** F8 zbirna (`T_Zbirna_NemaIspravku`), F2 „Ispravi"
(`IspravkaOtpremnice_TX`, BFP `modBusinessFlowProTests:15369`, `:17519`), PONIŠTENJE i
DUPLI, ispravka lista u nacrtu (S3d-2), storno lista van otpremnice, hladnjački lanac,
prijemnica / faktura / palete.

| # | Slučaj | Tvrdnja | Gde |
|---|---|---|---|
| T1 | **crveno pre reza:** izdata otpremnica, vozač predao gajbe kupcu, F2 „Ispravi" | danas pada na `AMB-INV-07` (potvrđuje §3) | BFP |
| T2 | H1: storno lista, izdata otpremnica u izdatoj zbirnoj, jedna prijemnica | otpremnica' i zbirna' nastale, novi brojevi, isti datum, stare stornirane sa `ZamenjenSaID`; prijemnica prevezana oba polja, količina ista; ambalaža neto | BFP |
| T3 | H2 sestre: zbirna sa tri otpremnice | zbirna' = otpremnica' + dve netaknute; stara zbirna i dalje sastavljena od tri (`IzvoriZbirne`) | BFP |
| T4 | ispravka lista (kg I → II) u lancu | otpremnica' očekivanje = staro − stari + novi po klasi; novac prenesen, preplata prijavljena | BFP |
| T5 | poslednji list; poslednja otpremnica bez prijemnice | bez naslednika (`LNC-ODL-07`) | BFP |
| T6 | **negativno:** poslednja otpremnica, aktivna prijemnica | odbijanje, **nijedna** tabela promenjena | BFP |
| T7 | **negativno:** pad posle zbirne' (seam) | rollback svih tabela (`LNC-INV-01`) | BFP |
| T8 | **negativno:** ispravka menja kulturu / stanicu / ambalažu | odbijanje sa porukom (`LNC-ODL-10`) | BFP |
| T9 | **negativno:** hladnjački list | odbijanje, upućuje na PONIŠTENJE | BFP |
| T10 | **negativno:** aktivna prijemnica sa labelom stare zbirne bez `ZbirnaRoditeljID` | odbijanje (`LNC-INV-07`) | BFP |
| T11 | dve prijemnice iste zbirne | obe prevezane | BFP |
| T12 | fakturisana prijemnica | prevezana, faktura netaknuta | BFP |
| T13 | nacrt otpremnice; izdata u nacrtu zbirne | u mestu, bez novog broja, očekivanje nacrta isto | BFP |
| T14 | UNDO nad storno lista iz poteza | odbijen (`LNC-INV-11`) | `RunStornoTestSuite` |
| T15 | auto-broj isključen | brojevi verzija ipak nastaju (`LNC-ODL-17`) | BFP |
| T16 | `tblStornoVeze` posle poteza | red po paru, zajednički koren, nijedan `PENDING` (`LNC-ODL-20`, `LNC-INV-10`) | BFP |
| T17 | F8 uvid i ponuda radnje za list u lancu; F1 potvrda pri snimanju, „Ne" ne piše ništa | ekran | `modTest` |
| T18 | sync posle poteza: malina auto-otpremnica i auto-zbirna ne prave ništa | `LNC-INV-09` | BFP |

Šema i vlasništvo bez Excela: `gen_schema_module.py --check`, `who_writes.py
--check-ownership`, `vba_check`. Sabotaže (dvosmerni dokaz) za kritične invarijante:
ambalaža neto, izbor prijemnica po ID-u, redosled, rollback, UNDO kapija. Lanac u testovima
se gradi **kanonskim piscima**, ne fixture redovima starog oblika (§7, FALSE-GREEN).

## 9) Landing

- **Prioritet: ovaj rez ide PRE S9** (operater, 06.11.2026: „prioritet najveći, iznad S9").
  PR #469 (pre-flight S9) se preslaže na ovaj rez, ne obrnuto. S9 dira iste ivice
  (`SLED-ODL-03` istorija iz `tblStornoVeze`, `SLED-ODL-08` brisanje broj-oblika traga, S9-6
  kanon), pa važi u oba smera: ovaj rez **ne piše** broj-oblik koji S9 briše, a S9-6 **ne
  briše** nove ID kolone zbirne (`LNC-ODL-16`) i čita red po paru u `tblStornoVeze`
  (`LNC-ODL-20`). Pre-flight S9 posle preslaganja treba ponovo da izmeri merenje 2
  („kolone bez pisca"): `tblZbirna.IspravkaOdID` / `ZamenjenSaID` tada imaju pisca.
- **Šema** (`LNC-ODL-16`) je centralni registar → serijski, sama u svom commitu.
- PR #457 (ambalaža, otvoren) dira `AMBALAZA.md`, `README.md`, `modTest.bas`: brojevi
  testova i odeljaka se dodeljuju **posle** njegovog merge-a.
- Numeracija plana (`REFAKTOR_DOKUMENT_HEADER_STAVKE.md` §14.x) i stavki
  (`STANJE_REFAKTORA.md`) se ovde namerno ne dodeljuje: #469 zauzima §14.64 i stavke
  323–324.
- Predlog sečenja (svaki deo zelen sam za sebe): **1** šema + ambalaža na kraju
  (`AMB-10-ODL-40`) + T1 · **2** jezgro poteza za storno (BFP T2–T3, T5–T7, T9–T12, T16, T18)
  · **3** ispravka (T4, T8, T13, T15) · **4** ekrani F8/F1 + UNDO kapija (T14, T17).
