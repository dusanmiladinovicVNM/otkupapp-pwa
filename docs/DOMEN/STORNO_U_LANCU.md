# Storno i ispravka otkupnog lista u lancu (`LNC-*`)

> **Status: PRE-FLIGHT 06.11.2026 — odluke operatera zapisane. U kodu: deo 1 (šema
> `LNC-ODL-16`, ambalaža `LNC-ODL-15` / `AMB-10-ODL-40`), deo 2 (jezgro storna,
> `modDokumenta.LncPlanLista` / `StornoListaULancu_TX`) i deo 3 (ispravka,
> `IspravkaListaULancu_TX` nad `modOtkup.IspravkaOtkupaUTransakciji`, ruta u
> `modOtkupUnos`) i deo 4 (F8, F1, UNDO kapija). NEVERIFIKOVANO u Excelu.**
> **§10 (pre-flight 08.11.2026):** „Dodaj blok" / „Izvadi blok" u izdatoj otpremnici
> (`LNC-ODL-22..38`, `LNC-INV-12..21`) — odluke operatera zapisane, koda nema.
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

Ručni put je osam koraka na četiri ekrana (F8 zbirna DUPLI → F1 „Ispravi" → ukloni list
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
| **`LNC-ODL-03`** | Očekivanje nove otpremnice se **izvodi**: staro − list (+ novi list kod ispravke), po klasi, za kg i gajbe | Stara je bila izdata, a izdavanje je dokazalo očekivano = povezano po klasi (`OtpIzdaj`, `modDokumenta.bas:4923`). Umanjenje za jedan poznat list nije „izjednačavanje" koje je S3b-2a odbio. **Dopuna 08.11.2026:** premisa važi samo za usklađenu otpremnicu. Otpremnica starog oblika može da nosi očekivanje različito od zbira svojih listova, a nova izdata verzija sme da nosi samo zbir; zato potvrda takvog poteza počinje upozorenjem (`LNC-ODL-21`, KI-094) |
| **`LNC-ODL-04`** | Novac kao i danas: storno odvezuje (`ResetNovacOtkupLink`, virman postaje slobodan avans); ispravka prenosi na nov list i prijavljuje preplatu (PUT A) | `ODLUKA_NOVAC_PRI_STORNU.md` (A + D suženo); bez novog koda za novac |
| **`LNC-ODL-05`** | Ulaz za storno: **F8 nad listom**. Umesto današnjeg odbijanja nudi radnju sa uvidom posledica pre potvrde | F8 već ima obrazac „prvo vidi posledice, pa odluči" (`GradiUvid`, `modStornoImpact`) |
| **`LNC-ODL-06`** | Obim prvog reza: **storno + ispravka** lista u lancu. Premeštanje lista na drugu otpremnicu je naredni rez, istom mehanikom | **Nastavak:** §10 — „Dodaj blok" / „Izvadi blok" u izdatoj otpremnici, premeštanje u dva poteza (`LNC-ODL-22..38`, pre-flight 08.11.2026) |
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
| **`LNC-ODL-20`** | Korekcija = **red po paru u `tblStornoVeze`** (list, otpremnica, zbirna — svaki svoj `CorrectionID`), vezani **zajedničkim korenom**: svi redovi poteza nose `ParentDocType = LANAC`; red lista je koren (`ParentDocID` prazan), a redovi otpremnice i zbirne nose `ParentDocID` = `CorrectionID` korena. Dokument nosi `CorrectionID` svog para | `tblStornoVeze` drži jedan par po ključu (`CreateCorrectionContext`), a S9 istoriju čita par po par (`SLED-ODL-03`). **Precizira A13:** „isti CorrectionID" znači „isti koren korekcije". **Precizirano 08.11.2026:** potez bez para lista („Dodaj blok" / „Izvadi blok") ima koren na paru otpremnice (`LNC-ODL-27`) |
| **`LNC-ODL-21`** | Stari dokument koji se **ne slaže sa svojim izvorima** — otpremnica sa listovima, izdata zbirna sa otpremnicama (očekivano ≠ zbir izvora po klasi, za kg ili gajbe) — se **ne odbija**: posledice u potvrdi poteza (Da/Ne na F8, F1 „Storno" i F1 „Ispravi blok") počinju upozorenjem `PAZNJA:` sa **svim** razlikama i brojevima, a dijalog tada nudi **„Ne" kao podrazumevano**. Posle „Da" nova verzija nosi zbir svojih izvora; upozorenje ide u log posle commit-a. Bruto se uvek računa iz listova (S4.2) | Odluka operatera 08.11.2026 (KI-094). Izdata verzija sme da nosi samo zbir izvora (`OtpRequireJednakost`, `ZbrRequireJednakost`), pa razlika ne može da se prenese; operater odlučuje da li je zbir tačan. **Dopuna 09.11.2026:** upozorenje o aktivnoj prijemnici zbirne (X kg zbirne' prema Y kg prijemnica, i gajbe) važi za sve poteze u lancu (`LNC-ODL-24`) |

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
| **`LNC-INV-09`** | **Nijedan ostatak za automatiku:** posle poteza nijedan aktivan list iz stare otpremnice nije van aktivne otpremnice, i nijedna aktivna izdata otpremnica iz stare zbirne nije van aktivne zbirne | inače ga sledeći sync obuhvati: malina auto-otpremnica, auto-zbirna (`NevezaneOtpremnice`) ili hladnjački auto-lanac. **Suženo za „Izvadi blok"** (`LNC-INV-14`, `LNC-ODL-36`): izvađeni blok je namerno slobodan |
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

> **NEMERENO.** Tvrdnja „današnje F1 Ispravi pada kad je vozač predao gajbe" izvedena je
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
| papir otpremnice „ispravka od" | čita broj-oblik, koji se ne piše (`SLED-ODL-08`) | ostaje prazan, kao danas posle F1 „Ispravi"; S9 prebacuje papir na `tblStornoVeze` |
| faktura, SEF, palete, utovar | vezuju prijemnicu / stavku prijemnice, ne zbirnu ni otpremnicu | bez izmene (`LNC-ODL-08`) |
| strogi čitači članstva ✔ `AktivnoClanstvoPoKanonu` | dva aktivna članstva iste otpremnice i storniran izvor u aktivnoj izdatoj zbirnoj su tvrda greška | redosled `LNC-INV-02` |
| doplata (signal DOP1) | prijavi storniran ili ispravljen list u već obračunatom periodu | postojeći signal, bez izmene |

## 7) Verdikt po osama (pre-flight skill §1)

| Osa | Status | Dokaz |
|---|---|---|
| **DOMAIN** | **PROVEN** | ZBR-KANON-03, A13, A14, A15, H1/H2, D1 §10 t. 3; odluke `LNC-ODL-01..21`; matrica §4 |
| **IDENTITY** | **PROVEN uz uslov** | `OtkupID` putuje od reda F8 do pisca (`modStornoDok.bas:204`, `:358`), od F1 kroz `mIspravkaOtkupID`; otpremnica i zbirna iz kanonskog članstva (`OtpremnicaZaOtkup`, `AktivnaZbirnaZaOtpremnicu`). **Uslov:** prijemnice po `ZbirnaRoditeljID`, ne po labeli (`LNC-INV-07`); trag po ID-u (`LNC-ODL-16`, `-20`) |
| **CARDINALITY** | **PROVEN** | list ∈ 0..1 aktivna otpremnica (`OtpRequireIzvorValjan`); otpremnica ∈ 0..1 aktivna zbirna (`AktivnoClanstvoPoKanonu`, dva = greška); nacrt otpremnice ∉ zbirna (`:2154`); **zbirna → 0..N prijemnica** van hladnjače — S9 sonda nad C002: 5 od 894 zbirnih ima dve (PR #469, merenje 5); hladnjača tačno 1 |
| **INVARIANTS/OWNER** | **PROVEN** | `LNC-INV-01..11`, svaka sa vlasnikom (§2); `AMB-10-ODL-40` u `AMBALAZA.md` |
| **WRITERS** | **PROVEN uz uslov** | `WHO_WRITES.md` / `WRITE_OWNERSHIP.json`: otpremnica, zbirna, izvori, stavke i prijemnica → `modDokumenta`; otkup → `modOtkup`, `modDokumenta`, `modStornoFlow`; `tblStornoVeze` → `modStornoContext`; novac → `modNovac` (+ `modStorno`); ambalaža → `modAmbalaza`. Zastavicu `Stornirano` piše `modStorno.MarkRowStornirano` kroz promenljivu (mapa je ne vidi — imenovano u `WHO_WRITES.md`). **Uslov:** koordinator u `modDokumenta`, sve ostalo kroz API vlasnika → nijedan nov pisac (`who_writes.py --check-ownership` zelen) |
| **DOWNSTREAM** | **PROVEN** | §6; svaka stavka ima uslov ili odluku (`LNC-ODL-18/19`, `LNC-INV-07/09/10`) |
| **CAPABILITY** | **PROVEN** | nova: F8 nad listom u lancu (danas odbijanje); proširena: F1 „Ispravi blok" (danas odbijanje); **ostaje netaknuto:** F8 zbirna DUPLI/PONIŠTENJE/REŠI KASNIJE (`T_Zbirna_NemaIspravku` i dalje važi — premisa testa se ne menja), F1 „Ispravi", Oporavak „Prevezi" |
| **ACCEPTANCE CONTRACT** | **PROVEN (plan)** | §8 |
| **PLATFORM** | **N/A** | nema novog Excel/VBE/MSForms ponašanja: dijalog ide kroz postojeći seam `modOtkupUI.Dijalog`, F8 kroz `StornoIzvrsiMod`. Cena snapshot-a ~13 tabela **nemerena** (presedan 11) |
| **LANDING** | **RIZIK — serijski, ovaj rez pre S9** | §9 |

**⚠ ARCHITECTURAL EDGE CASE** (zatvoren odlukom): ambalaža u međukoraku (`LNC-ODL-15`).
**⚠ IDENTITY RISK** (zatvoren invarijantom): izbor prijemnica po labeli (`LNC-INV-07`).
**⚠ FALSE-GREEN RISK:** test nad fixture-om starog oblika (`Otkup.OtpremnicaID`, kolone bez
pisca — S9 nalaz N3) bi merio tautologiju; testovi reza grade lanac **kanonskim piscima**.

## 8) Dokaz — testovi i sabotaže u kodu (NEVERIFIKOVANO u Excelu)

**Šta važi:** potez iz §4, sa invarijantama §2, kroz F8 (storno) i F1 (ispravka), jedan
potez i jedna transakcija.

**Šta ostaje netaknuto** (premisa testova se ne menja): F8 zbirna DUPLI/PONIŠTENJE/REŠI
KASNIJE (`T_Zbirna_NemaIspravku`), F1 „Ispravi" (`IspravkaOtpremnice_TX`, sada kroz
zajednički `OtpZaglavljeVerzije`), `IspravkaOtkupa_TX` (i dalje odbija izdatog roditelja),
pisac `StornoOtkup_TX` (i dalje odbija list u aktivnoj otpremnici), storno lista van
otpremnice, hladnjački lanac.

**Premise koje se MENJAJU, namerno** (odluke operatera 06.11.2026; review `#471`, nalazi 2 i 3):

| Test | Bilo | Sada |
|---|---|---|
| `Test_OTK_IzvorAktivneOtpremniceSeNeStornira` (pomoćnik `IzvorOtpremniceOdbijen`) | F8 odbija list u otpremnici i imenuje je | F8 **pušta** potez u lancu (`LNC-ODL-05`); tvrdnje pisca ostaju |
| `Test_OTK_IspravkaBlokaUNacrtu`, A13 test ispravke izdate (`"... poruka imenuje put koji postoji"`) | odbijenica imenuje ručni put „ispravi otpremnicu" | odbijenica imenuje **potez u lancu** (`LNC-ODL-12`) |

Lanac u testovima grade **kanonski pisci**, sa sopstvenim tipom gajbe po scenariju (§7,
FALSE-GREEN). Svi testovi su u `RunBusinessFlowProSuite`.

| Plan | Test (BFP) | Sabotaža |
|---|---|---|
| ambalaža na kraju (`AMB-10-ODL-40`): A → B, bez odlaganja pada | `Test_AMB_ODL40_MeriPosleZamene` | `odl40-kontra-meri-odmah`, `amb-storno-bez-inv07` |
| zamena bez pokrića pada na kraju, rollback | `Test_AMB_ODL40_BezPokricaPadaNaKraju` | `odl40-zavrsi-bez-inv07` |
| zaboravljen završetak — commit odbijen; zaostalo stanje se zamenjuje | `Test_AMB_ODL40_CommitBezZavrsetkaOdbijen` | `odl40-commit-bez-kapije`, `odl40-zaostalo-ne-zamenjuje` |
| ugovor početka odlaganja | `Test_AMB_ODL40_UgovorPocetka` | — |
| H1: storno u izdatoj zbirnoj, prijemnica, trag, kontekst, ambalaža neto; `KupacHladnjaca` nasleđen i kad se config promeni posle nastanka | `Test_LNC_StornoListaUIzdatojZbirnoj` | `lnc-prijemnica-ostaje-na-staroj`, `lnc-kupac-hladnjaca-preracun` |
| H2: sestre | `Test_LNC_SestrePrezivljavaju` | `lnc-sestre-ispadaju` |
| poslednji izvor (bez naslednika / odbijen uz prijemnicu bez gajbi, da kapija bude jedini razlog) | `Test_LNC_PoslednjiIzvor` | `lnc-poslednji-izvor-bez-kapije` |
| rollback bez seam-a (neslaganje ambalaže) | `Test_LNC_NeslaganjeAmbalazeVracaSve` | (pokriva ga `odl40-*`) |
| prijemnica samo po labeli (sintetička anomalija) | `Test_LNC_PrijemnicaBezIDaOdbijena` | `lnc-prijemnica-po-labeli-tiho` |
| nacrt otpremnice u mestu | `Test_LNC_NacrtOtpremniceUMestu` | — |
| izdata otpremnica bez zbirne | `Test_LNC_IzdataOtpBezZbirne` | — |
| izdata otpremnica u nacrtu zbirne (u mestu, ZBR-KANON-04) | `Test_LNC_NacrtZbirneUMestu` | `lnc-nacrt-zbirne-bez-nove` |
| ispravka u lancu (staro − stari + novi; keš prelazi na nov list, preplata prijavljena) | `Test_LNC_IspravkaListaULancu` | `lnc-ispravka-bez-novog-lista` |
| ruta F1 snimanja: `modOtkupUnos.OtkupUpisi` → potez u lancu | `Test_LNC_IspravkaKrozUnos` | `lnc-unos-ruta-bez-lanca` |
| ispravka menja ambalažu — odbijena | `Test_LNC_IspravkaMenjaAmbalazuOdbijena` | `lnc-ispravka-menja-lanac-tiho` |
| broj verzije bez prekidača auto-broja (generator i sam potez) | `Test_LNC_BrojVerzijeBezPrekidaca` | — |
| UNDO poteza odbijen | `Test_LNC_UndoPotezaOdbijen` | `lnc-undo-vraca-deo-poteza` |
| F8: kapija, potvrda sa posledicama, storno potezom, poslednji izvor | `Test_LNC_F8StornoListaULancu` | `lnc-f8-obican-storno` |
| F1: otvaranje ispravke, potvrda sa posledicama | `Test_LNC_F1IspravkaULancu` | `lnc-f1-bez-potvrde` |

Postojeća sabotaža `ispravka-gubi-identitet-utovara` sada gađa zajednički
`OtpZaglavljeVerzije`, a `ispravka-bloka-*` i `ki024-*` premešteno jezgro
`IspravkaOtkupaUTransakciji` — tekst sidara je isti. `ispravka-bloka-izdata-prolazi` je
popravljena: stara zamena je izdatog roditelja slala u granu „nepoznat status", pa tvrdnja
nije mogla da padne (zatečeno pre `#471`, review nalaz 19).

**Review `#471` (06.11.2026)** — 7 dimenzija, svaki nalaz kroz dva skeptika: 27 potvrđenih ili
verovatnih, svi ispravljeni osim jednog zatečenog. Najteži: potez sa preostalim listovima je
UVEK padao — strog čitač stavki otpremnice je video novo zaglavlje bez stavki (nalaz 1). Zatečen
kvar potvrde manjka gajbi u ispravci sa gajbama (nalaz 16) je `KI-089`, svoj rez. Ispravljen je u
rezu KI-089 (07.11.2026) za obe rute ispravke, ali u Excelu još nije proveren: ekran meri manjak
bez starog lista (`AMBALAZA.md` 6.5). Storno otpremnice i zbirne u potezu ne dira nalog
kooperanta, pa isti račun važi i u lancu (`Test_LNC_KI089_PotvrdaIspravkeULancu`).

**Ručna provera operatera (07.11.2026).** Potez je prošao na svežem lancu: storno lista
4/071126-2 je napravio novu verziju otpremnice i zbirne i prevezao prijemnicu. Nad listom
5/250626-13 je pao sa porukom „Header polje je prazno: KulturaID". Njegova otpremnica je zapis
**starog oblika**: verzija programa koja ju je napravila nije pisala kulturu na zaglavlje.
Potez je zaglavlje nove verzije prepisivao sa stare, a pisac otpremnice prazno zaglavlje odbija.
Transakcija se vratila, pa podaci nisu dirani. Isto je padao i F1 „Ispravi" nad takvom
otpremnicom.

Pravilo od tada: kultura otpremnice mora biti kultura svakog lista (`OtpRequireIzvorValjan`),
pa se nova verzija starog zaglavlja vodi **kulturom listova koji u nju ulaze**
(`OtpKulturaVerzije`, isti za potez i za F1 „Ispravi"). Kad listovi nemaju jednu, nepraznu kulturu, plan
odbija **pre potvrde** i imenuje listove. Plan ispravke (`LncPlanLista(…, zaIspravku:=True)`,
F1 „Ispravi blok") meri i list koji se ispravlja, jer nov list nosi njegovu kulturu
(`LncIspravkaMenjaLanac`). Zato je storno lista druge kulture iz takve otpremnice dozvoljen
(ostaju listovi jedne kulture), a njegova ispravka nije. Kultura se ne izvodi iz
`VrstaVoca`/`SortaVoca` starog zaglavlja, jer bi to bio drugi izvor istine. Test:
`Test_LNC_StaraOtpremnicaBezKulture` (F8, F1 „Ispravi", list bez kulture, listovi dve kulture);
`KI-091`.

**Ručna provera operatera (08.11.2026).** Ispravka lista 5/250626-13 za +1 kg dala je novu verziju
otpremnice od 650,1 kg umesto 680,5. Stara 5/060726 je nosila 679,5 kg, a njeni listovi 649,1 (gajbe
su se slagale). `OtpUpisiOcekivanjeVerzije` piše zbir listova nove verzije, što je jedino što izdata
verzija sme da nosi (`OtpRequireJednakost`). Razlika od 30,4 kg je zato nestala tiho. Od tada plan
poredi staru otpremnicu sa njenim listovima, po klasi, za kg i gajbe, istim telom pravila kao izdavanje
(`OtpRazlike`). Kad se ne slažu, potez se **ne odbija**: posledice u potvrdi na F8 i F1 počinju upozorenjem
`PAZNJA:` sa svim razlikama, po klasi, za kg i gajbe („klasa I: na otpremnici 679.50 kg, zbir listova
649.10 kg (30.40 kg vise)"), a operater bira Da/Ne (`LNC-ODL-21`, odluka operatera 08.11.2026; dijalog tada
nudi „Ne" kao podrazumevano). Isto važi za **izdatu zbirnu**: nova verzija je zbir otpremnica, pa se stara
zbirna poredi sa **zbirom svojih izvora** — stavke ostalih otpremnica plus listovi ove otpremnice — što je tačno
neto promena zbirne (review KI-094: isti kvar jedan nivo iznad, ranije nevidljiv i za `LNC-INV-06`; zbirna
koja se sa listovima slaže ne dobija upozorenje ni kad se otpremnica ne slaže). Posle „Da" nova verzija nosi zbir izvora, a upozorenje
ide u log posle commit-a. Test: `Test_LNC_OtpremnicaNeuskladjenaSaListovima`; `KI-094`.

**Otvoreno, neizmereno (`KI-092`).** Nova verzija otpremnice upisuje `VrstaVoca`/`SortaVoca`
iz šifarnika kultura, a ne prepisuje tekst starog zaglavlja. Kad je zbirna izdata, nova verzija
zbirne preuzima te podatke od prvog izvora. Ako sestara nema, prvi izvor je nova otpremnica, pa
prevezivanje prijemnice (`PrijemnicaRobaRazlog`) poredi robu prijemnice sa tim tekstom. Sestra
starog oblika poredi se isto (`ZbrRequireIzvorValjan`, `ZbrRequireIstiAko`). Kad se tekst
starog oblika razlikuje od šifarnika, potez pada **posle potvrde**, u transakciji (rollback), i
to **i bez sestre**. Bez podataka iz sveske to nije izmereno.

**Nije pokriveno testom:** sam poziv potvrde iz `Scr_Save` (odluka se meri kroz
`Scr_PotvrdaLancaProbaj`, isto telo) i sync posle poteza (`LNC-INV-09` drži konstrukcija:
preostali listovi i otpremnice ulaze u nove verzije). Oba idu u ručnu proveru.

Kapije bez Excela (zelene): `vba_check`, `gen_schema_module.py --check`,
`who_writes.py --check` i `--check-ownership`, `sabotaza.py --proveri-sidra`, hard census,
parity, self-update kapije. **Excel nije pokrenut:** `RunBusinessFlowProSuite`, `RunAllTests`,
`dokaz.py lnc` / `odl40` i Compile čekaju operatera.

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

## 10) Dodaj blok / Izvadi blok u izdatoj otpremnici (`LNC-ODL-22..38`, `LNC-INV-12..21`)

> **Status: PRE-FLIGHT 08.11.2026 — odluke operatera zapisane (08. i 09.11.2026). U kodu (deo (a)):
> a1 — T21 crven nad starim kodom, potvrđeno u Excelu 09.11.2026 (BFP RunID 20261109212916-2353,
> 18 padova, svi u T21 (a)/(b)); a2 — kapija `LNC-ODL-37` (`LncNedovrsenKontekstRazlog`, granice u
> `Test_LNC_NedovrsenKontekstGranice`); a3 — trag upozorenja `LNC-ODL-38`
> (`modStornoContext.DopisiPorukuKonteksta`, T23); a4 — upozorenje o aktivnoj prijemnici `LNC-ODL-24`
> za sve poteze (`LncNapomenaPrijemnice`, T22; F1 potvrda dobija polja forme kroz
> `modOtkupUnos.OtkupStavkeIzForme`); a5 — preslaganje bez promene ponašanja (`LncPlanOtpremnice`,
> `OtpProveriIzvor` / `OtpIzvorRazlog`, `LncKontekstPar` kao funkcija, koren na paru otpremnice);
> a6 — „Izvadi blok" (`LncPlanBlok`, `LncPlanOpisZaBlok`, `IzvadiBlokULancu_TX`, testovi
> `Test_LNC_Blok_Izvadi*` — 7, sabotaže `lncb-izvadi-*`, `lncb-koren-na-bloku`, `lncb-zastareo-cilj`,
> `lncb-trag-izvadi-bez-upozorenja`, `lncb-sufiks-izvadi`). Pravi hladnjački lanac → PONIŠTENJE,
> N1 (ručna otpremnica stanice posle označene kao hladnjača) → F8 storno lista u lancu.
> a2–a6 NEVERIFIKOVANO u Excelu.**
> Mereno čitanjem nad `main` = `1fa4937` (merge #473); Excel nije pokrenut, pa je svaka tvrdnja
> o ponašanju **NEVERIFIKOVANA**. Nastavak `LNC-ODL-06` („premeštanje lista na drugu otpremnicu
> je naredni rez, istom mehanikom"). Važi sve iz §1–§9 osim onoga što ovaj odeljak izričito
> menja; kad se kod i ovaj odeljak raziđu posle reza, kod je u pravu i ovaj odeljak je bug.

### 10.1 Problem

Blok (otkupni list) je zaboravljen na izdatoj otpremnici, ili je na nju upisan a nije putovao
tim kamionom. Izdata otpremnica koja je izvor aktivne zbirne (nacrta ili izdate) danas nema
nijedan put za promenu sastava; van zbirne postoji samo posredan put u više koraka (F1 „Ispravi"
→ nacrt → „vezi" / „ukloni" → izdaj), bez reda u `tblStornoVeze`:

| Kapija | Gde |
|---|---|
| svaki upis i brisanje članstva otpremnice traži NACRT | `RequireOtpDraft` (`modDokumenta.bas:6950`), zvan iz `OtpRequireIzvorValjan` (`:5557`), `OtpUkloniIzvor` (`:5722`), `IzvadiIzvorIzNacrta` / `UvediIzvorUNacrt` (`:6567`, `:6582`) |
| F1 „Ispravi" izdate (B-022) pravi NACRT nove verzije, ali se odbija kad je otpremnica izvor aktivne zbirne | `OtpIspravi` → `modStorno.StornoOtpremnica` (`modStorno.bas:241`) |
| klik na izdatu otpremnicu (F1 lista „Otpremnice", F2 lista „Svi") javlja „izdata — ne prima nove blokove" | `NacrtRazlog` (`modScrDokumenti.bas:305-320`), `OTKUI_ERR_OTP_IZDATA` |

Kad je otpremnica u izdatoj zbirnoj, ostaje samo ručno rasklapanje lanca, kao u §0.

### 10.2 Odluke (`LNC-ODL-22..38`)

| ID | Odluka | Obrazloženje / izvor |
|---|---|---|
| **`LNC-ODL-22`** | **„Dodaj blok" i „Izvadi blok"** nad IZDATOM otpremnicom su potez u lancu po `LNC-ODL-01..21`: jedna transakcija, stara otpremnica (i zbirna) padaju, nove verzije, prijemnice prevezane po ID-u, trag po ID-u, hladnjački lanac van poteza, upozorenje `LNC-ODL-21`. **Blok se ne stornira**: posle „Izvadi" je slobodan. Važi i bez zbirne (tada samo otpremnica'). F1 „Ispravi" (B-022, `IspravkaOtpremnice_TX`) ostaje **netaknut** — ostaje put za izdatu otpremnicu van zbirne, sa svojim tragom; ovaj rez ga ne menja i ne ujednačava | Operater 08.11.2026 (ulaz nad izdatom otpremnicom, obim dodavanje + uklanjanje). Skelet: `StornoListaULancu_TX` (`modDokumenta.bas:3996-4049`) bez koraka lista. „Minimal change": dva puta za izdatu van zbirne su svesna razlika (F1 „Ispravi" ne piše `tblStornoVeze` i meri ambalažu odmah) |
| **`LNC-ODL-23`** | **Premeštanje A → B su dva poteza:** „Izvadi" iz A, pa „Dodaj" u B. Obrnut redosled nije moguć (Dodaj traži slobodan blok, 1291). U istoj izdatoj zbirnoj to daje Z → Z' → Z'' i dva nova broja. Međustanje je commit-ovano, pa „Izvadi" može da padne na `AMB-INV-07` kad je vozač gajbe bloka već predao kupcu, iako bi oba koraka zajedno bila neto nula — poruka tada imenuje vozača, tip i manjak i kaže da premeštanje u jednom potezu ne postoji. Potvrda „Izvadi" kaže da blok ostaje slobodan i da se u drugu otpremnicu dodaje posebnim potezom. Jednopotezno „Premesti A → B" je zaseban rez; **okidač** za njegov pre-flight: ručna provera (UAT) pokaže da je premeštanje u malina režimu često ili da sync preuzima blok između dva koraka | Operater 08.11.2026 (dva koraka). Odlaganje merenja važi unutar jedne transakcije (`modAmbalaza.bas:2882-2924`); odbijanje međustanja je tačan signal |
| **`LNC-ODL-24`** | **Zbirna sa aktivnom prijemnicom:** potez je dozvoljen. Potvrda počinje `PAZNJA:` i nabraja: zbirna Z → X kg; prijemnice (broj: kg, …), ukupno Y kg; razlika — i isti red za gajbe. „Ne" je podrazumevano (`LncTekstImaUpozorenje`). Važi za **sve poteze u lancu** (i F8 storno lista, F1 „Storno" i F1 „Ispravi blok"), jednim telom. Upozorenje stoji **samo** kad postoji aktivna prijemnica ili razlika po `LNC-ODL-21`; usklađen potez bez prijemnice ostaje bez njega. X se računa istim primitivom kao pisac zbirne (`StavkeZaOtpremnicu` po sestri, `CreateZbirna`, `modDokumenta.bas:1155-1171`) plus listovi otpremnice', Y je zbir stavki prijemnica izabranih po `ZbirnaRoditeljID`, nikad po labeli. Klasa se ne poredi | Operater 08.11.2026 (upozorenje, „Ne" podrazumevano) i 09.11.2026 („kg i gajbe, svi potezi"). Klasu kupac procenjuje pri prijemu (`PRJ-ZBR-02`). Premisa `Test_LNC_OtpremnicaNeuskladjenaSaListovima` („usklađena nema upozorenje", `modBusinessFlowProTests.bas:27500`) ostaje: fixture je bez prijemnice. Uz aktivnu prijemnicu potez traži ispravan `tblPrijemnicaStavke` (strogi čitač `ZbirStavkiPoPrijemnici`, isto kao F4); kad posledice ne mogu da se izračunaju, potvrda nosi `PAZNJA` i nudi „Ne". Potvrda je MsgBox (~1024 znaka): `modDokumenta.LncDijalogTekst` drži sva tri dijaloga (F8, F1 „Storno", F1 „Ispravi blok") ispod granice — uvod i pitanje se uvek vide, opis se skraćuje na kraju uz napomenu; pun tekst ide u trag i log |
| **`LNC-ODL-25`** | **Poslednji blok = `LNC-ODL-07`:** otpremnica se stornira bez naslednika. Kad je ona jedini izvor izdate zbirne sa aktivnom prijemnicom → odbijanje **pre potvrde**, uputi na F8 PONIŠTENJE. Sa sestrama: zbirna' = sestre. Bez sestara i bez prijemnice: i zbirna bez naslednika. Nacrt zbirne: otpremnica izlazi u mestu (`ZBR-KANON-04`) | Operater 08.11.2026. Telo postoji: `LncPlanLista` stornoRazlog (`modDokumenta.bas:3910-3920`), `LncSastavi` za praznu otpremnicu (`:4273-4296`) |
| **`LNC-ODL-26`** | **Potez ne dira list:** nema `StornoOtkup` ni `IspravkaOtkupaUTransakciji`, nema upisa u `tblOtkup`, `tblOtkupStavke`, `tblNovac`, `tblStornoZurnal`. Novac, isplate, avans, kartica i DOP1 ostaju na bloku. Potez ne otvara operaciju žurnala, pa UNDO nema šta da vrati (`LNC-INV-11` po konstrukciji); inverz poteza je suprotan potez | Izvedeno: žurnal otvara samo `StornoOtkup` (`modStorno.bas:151-156`); `LNC-ODL-04` je pravilo storna lista, ne ovog poteza |
| **`LNC-ODL-27`** | **Koren traga bez para lista** je red para OTPREMNICE u `tblStornoVeze` (`OldDocType` Otpremnica, `ParentDocType` LANAC, `ParentDocID` prazan); red zbirne nosi `ParentDocID` = `CorrectionID` korena. Mod ISPRAVKA kad otpremnica' postoji, PONISTENJE kad se vadi poslednji blok. Na blok se `CorrectionID` **ne** upisuje (današnji koren ga upisuje na list, `modDokumenta.bas:4503`), pa `LancaRazlog` (`modStornoZurnal.bas:174-186`) ostaje nepromenjen i kasnijem običnom stornu bloka ne dodaje odbijanje (blok sa gajbama je ionako konačan: `KontraStavRazlog` ide prvi, `AMB-10-ODL-25`, `modStornoZurnal.bas:364-367`). Koji blok je ušao ili izašao čita se iz razlike članstva (A15) i iz teksta reda | Izvedeno; **precizira `LNC-ODL-20`** („red lista je koren" važi kad potez ima list). S9 čita par po par (`SLED-ODL-03`) |
| **`LNC-ODL-28`** | **Jedan plan** za potvrdu i potez, ponovo računat u transakciji (kao `LNC-ODL-11`). „Izvadi" ulazi po (`OtpremnicaID`, `OtkupID`) i blok mora pripadati baš toj otpremnici; „Dodaj" po (`OtpremnicaID`, `OtkupID` slobodnog bloka). Jezgro po otpremnici (status, hladnjača, članovi, kultura, zbirna, sestre, prijemnice, stornoRazlog) izdvaja se iz `LncPlanLista` uz parametar „izvori nove verzije"; kultura se meri i nad dodatim blokom. Potez **odbija NACRT** otpremnice — nacrt se menja postojećim „vezi" / „ukloni" (A-024, A-025) | Izvedeno: `modDokumenta.bas:3748-3750` (jedan odgovor), ulaz po listu `:3802-3806`, `:3850-3862`; bez odbijanja bi grana nacrta u `LncRasturi` (`:4218-4221`) „Izvadi" tiho pretvorila u „ukloni" u mestu, a „Dodaj" bi pala na tuđoj grešci 1296 (`OtpUkloniIzvor`, `:5745-5748`) |
| **`LNC-ODL-29`** | **Kapije dodatog bloka su ISTO TELO kao `OtpRequireIzvorValjan`** (bez `RequireOtpDraft`), u obliku koji vraća razlog: postoji jednom, nije storniran, ima stavke, ista stanica, ista kultura, izdat, isti tip ambalaže kad nosi gajbe, slobodan u svakoj nestorniranoj otpremnici (zauzet → poruka imenuje otpremnicu). Dodaju se dve: hladnjački ili neizvestan blok (`modAutoHladnjaca.LanacVaziZaBlok`, fail-closed) i hladnjačka ciljna otpremnica (`LNC-ODL-14`). Otpremnica **bez tipa ambalaže** + blok sa gajbama → odbijanje **čitljivim** razlogom; to je postojeća kapija „isti tip ambalaže", koja danas javlja tehničku 1239 („ocekivano '', naslo 'X'"); tip zaglavlja se ne izvodi iz bloka. „Izvadi" odbija blok za koji važi `LanacVaziZaBlok` (N1) i upućuje na F8 storno lista u lancu. Plan meri pre potvrde, pisac ponovo | Izvedeno: `.claude/rules/testovi.md` §5 (jedna implementacija, dva pozivaoca); kapija hladnjače danas postoji samo na ekranu (`modScrDokumenti.bas:659-668`); tip: `OtpZaglavljeVerzije` prepisuje prazan tip (`modDokumenta.bas:5856`), a `OtpRequireIzvorValjan` → `RequireIstoPolje` odbija (`:5620-5627`, 1239) u `LncSastavi` (`:4283`), pre očekivanja i knjiženja |
| **`LNC-ODL-30`** | **Datum:** „Dodaj" odbija blok čiji je datum **posle** datuma otpremnice — pre potvrde i u piscu, istim telom. Isti ili raniji dan prolazi, a potvrda tada imenuje oba datuma. Pravilo važi samo za ovaj potez; „vezi" u nacrtu ostaje kakav jeste. Otpremnica' nosi datum stare | Operater 09.11.2026. Roba kupljena posle polaska kamiona nije mogla njime da putuje, a izveštaji i knjiga gajbi idu po datumu otpremnice. Danas datum ne meri nijedan pisac (`modDokumenta.bas:5547-5646`); predaja sme da nosi više dana (odluka 23.09.2026, `REFAKTOR_DOKUMENT_HEADER_STAVKE.md:5065-5069`) |
| **`LNC-ODL-31`** | **`AMB-10-ODL-40` važi i za „Dodaj" i „Izvadi"** (ista zamena verzije izdate otpremnice u jednoj transakciji): `AMB-INV-07/09/16` se mere posle cele zamene. Običan storno otpremnice meri odmah | Izvedeno; `AMBALAZA.md` §6.14c. Bez odlaganja „Dodaj" lažno pada čim je vozač predao gajbe (§3) |
| **`LNC-ODL-32`** | **Brojevi, datum i utovar** kao postojeći potez: otpremnica' dobija broj iz `GenerateBrojOtpremnice(stanica, datum stare)` (`modDokumenta.bas:5883`), zbirna' iz `BrojNoveVerzije(KIND_ZBR, vozač, datum stare)` (`:4403`); datum stare verzije; `PredajaID` i `KupacHladnjaca` se nasleđuju, `ClientRecordID` ne | Izvedeno iz `LNC-ODL-17` i §1 („izvedeno iz postojećih pravila") |
| **`LNC-ODL-33`** | **Ulaz: F1, lista „Otpremnice".** Klik na IZDATU otpremnicu je bira kao **cilj** (umesto današnjeg toasta). Cilj je zaseban kontekst, **ne** `mOtpID`: radni sto u svakom trenutku ima najviše jedan kontekst — aktivan nacrt ili izabranu izdatu — i izbor jednog napušta drugi. Uz cilj: „Blokovi otpremnice" pokazuje njene blokove sa „Izvadi", „Bez otpremnice" nudi „Dodaj"; „vezi" / „ukloni" / „izdaj" se ne nude. `RutaPosleUpisa` i `VeziZaAktivnu` cilj ne vide (nov list ne ide u izdatu). Posle uspeha cilj prelazi na otpremnicu' (presedan `IspraviIzdatu`, `modScrDokumenti.bas:1419-1425`); bez naslednika se briše. F2 nad izdatom otpremnicom upućuje na F1. Izbor cilja se grana u `Scr_Event`; `AktivirajOtpremnicu` / `NacrtRazlog` se **ne** olabavljuju (`mOtpID` ostaje samo nacrt, `Test_OTP_RadniStoBiraSamoNacrt`). Nijedna lista ne nudi više od `MAX_ACT` = 5 radnji | Operater 09.11.2026. Odluka od 08.11 je glasila „F2"; merenje je pokazalo da F2 nema nijednu listu blokova (`modScrDokumenti.bas:97-104`), a da radni sto otpremnice i „Ispravi" izdate stoje u F1 (`:107-128`, `:203-214`). Višak radnji se tiho ne crta (`modOtkupUI.bas:194`, `BazenStaje`) |
| **`LNC-ODL-34`** | **Identitet putuje redom:** `OtpremnicaID` i `OtkupID` iz nevidljivih kolona reda, nikad po broju. Pisac proverava: „Izvadi" → `OtpremnicaZaOtkup(blok)` = cilj; „Dodaj" → `OtpremnicaZaOtkup(blok)` = "". Cilj koji je u međuvremenu storniran ili zamenjen → odbijanje po ID-u (pisac ne prati `ZamenjenSaID`) | Izvedeno: pre-flight §5; presedan `IspraviIzdatu` (`modScrDokumenti.bas:1401`) |
| **`LNC-ODL-35`** | **Tekst poteza je njegov:** blok „izlazi iz" / „ulazi u" otpremnicu; nema „storniran" ni reda o novcu koji se odvezuje, nego „isplate ostaju na listu". `LncNapomenaKolicine` dobija vrstu poteza (bez „(bez storniranog lista)"). Potvrda imenuje posledice iz §10.6 koje operater ne vidi: broj zbirne' za prijemnicu koja još nije uneta, vozačevu zbirnu sa telefona koja nosi staru otpremnicu, i u malina režimu sledeći sync (`LNC-ODL-36`). Potvrda ide kroz **postojeći** seam potvrde poteza u lancu (`mPotvrdaLancaTest`, obrazac `PotvrdiIspravkuULancu`: beleži, odgovara „Ne"; gasi ga `Scr_PotvrdaLancaTestReset`), a kontekst cilja briše `Scr_OtpTestReset` — oba već zove `modTest` (`ResetSeamova`, `CleanupPosleTesta`), pa `modTest.bas` ostaje netaknut. Goli `MsgBox` ne. Rečenica `LncPoruka` o PWA zbirni bez naslednika se ispravlja | Izvedeno: današnji tekstovi govore jezikom storna (`modDokumenta.bas:3976-3985`, `:4566-4570`); `.claude/rules/testovi.md` §4 |
| **`LNC-ODL-36`** | **Malina:** prihvata se da sledeći sync od izvađenog bloka napravi auto-otpremnicu; sync se ne menja. U malina režimu potvrda „Izvadi" kaže izričito: dodaj ga u ciljnu otpremnicu odmah, pre sledećeg sync-a. `LNC-INV-09` se za izvađeni blok svesno sužava (`LNC-INV-14`) | Operater 09.11.2026. `AutoCreateOtpremniceFromPWA_TX` (`modGoogleSyncOrchestrator.bas:240-245`) grupiše svaki slobodan izdat nehladnjački blok (`modMasterSync.bas:1325-1345`). Ista izloženost već postoji za „ukloni" iz nacrta i storno otpremnice |
| **`LNC-ODL-37`** | **Nedovršen kontekst nad onim što potez menja → odbijanje pre potvrde** (za sve poteze u lancu, u zajedničkom planu, i ponovo u transakciji). Odbija se kad postoji red `tblStornoVeze` sa `NeedsRecovery = Da` nad starom otpremnicom ili starom zbirnom (F8 „Reši kasnije", PONIŠTENJE / DUPLI koji je pao ili čeka ručno), ili **nedovršena ispravka prijemnice** te zbirne (mod ISPRAVKA, **status PENDING** — zamena nije uneta; stara prijemnica stornirana; veza po `ZbirnaRoditeljID` stornirane prijemnice). Ispravka u MANUAL_REQUIRED znači da je zamena već uneta (palete nisu prevezane): ona je aktivna, prevezuje se po ID-u i potez **ne** zaustavlja — F4 je ne bi ni prepoznao kao ispravku. Stari oblik podataka se poredi po broju i poklapanje odbija (fail-closed), uz napomenu „poklopljen po BROJU" u odbijenici: red konteksta bez `OldDocID` (osim starog MANUAL_REQUIRED, koji je uvek upisan posle storna svoje zbirne), ispravka prijemnice bez `OldDocID` (po `ParentBroj`), prijemnica bez `ZbirnaRoditeljID` (po `BrojZbirne`). Red koji ne može da se proveri odbija i odbijenica imenuje njega. Poruka imenuje `CorrectionID` i vrstu i upućuje: Oporavak → „Odbaci", ili završi ispravku prijemnice u F4 (T21 meri tokene „F4" i „Oporavak"). Nad otpremnicom i zbirnom kapija gleda `NeedsRecovery = Da` (svaki status); za prijemnicu traži da je **stornirana** (pala ispravka nad aktivnom prijemnicom ne smeta). Spisak se čita jednom po planu; greška čitanja ili kolona koje nema odbija potez. „Reši kasnije" / pao red nad **aktivnom** prijemnicom ne smeta: `PrijemnicaID` se ne menja i prevezuje se po ID-u | Operater 09.11.2026 („odbij, uputi na Oporavak"). Merenje T21 (čitanjem, 09.11.2026): plan i pisci danas ne čitaju `tblStornoVeze` (`LncPlanLista`, `LncRequirePlan`, `LncRasturi`, `modStornoDok.bas:200-239`); posle poteza podsetnik ostaje `PENDING` nad storniranom zbirnom, sa „Odbaci" kao jedinom radnjom (`modStornoRecovery.bas:60-73`, `modScrOporavak.bas:625-630`). Štetan slučaj: nedovršena ispravka prijemnice + storno jedinog izvora — stornirana prijemnica ne ulazi u `LNC-ODL-07` (`LncPrijemniceZbirne`, `modDokumenta.bas:4444`), zbirna ostaje bez naslednika i zamena prijemnice više nema roditelja (`ZBR_PARENT_NEMA`, `:524-528`) — **KI-095**. Nad otpremnicom F8 nedovršen kontekst ne pravi (`modStornoDok.bas:530-545`) |
| **`LNC-ODL-38`** | **Potvrđeno upozorenje je trajan trag:** tekst `PAZNJA` koji je operater potvrdio (`LNC-ODL-21`, `-24`) upisuje se i u `Message` korenskog reda poteza u `tblStornoVeze`, **dopisan** uz postojeći opis poteza (ne umesto njega), u istoj transakciji, za sve poteze u lancu. Log posle commit-a ostaje | Operater 09.11.2026. Kolona postoji; dopisuje je `modStornoContext.DopisiPorukuKonteksta` (dopisuje iza opisa, ne prepisuje — `CompleteCorrectionContext` bi opis zamenio), vlasnik tabele, bez promene šeme i bez novog pisca. S9 istoriju čita iz `tblStornoVeze` (`SLED-ODL-03`) |

### 10.3 Invarijante (`LNC-INV-12..21`)

| ID | Tvrdnja | Vlasnik (gde se drži) · nivo merenja |
|---|---|---|
| **`LNC-INV-12`** | **List je netaknut:** red `tblOtkup` bloka je isti u svim kolonama (`Stornirano`, `CorrectionID`, `IzdatoStatus`…), isti su i stavke, `tblNovac` i `tblStornoZurnal`; redovi knjige ambalaže sa dokumentom Otkup se ne menjaju | koordinator u `modDokumenta` ne zove `StornoOtkup`, `IspravkaOtkupaUTransakciji` ni `modNovac`; `BeginStornoOp` postoji samo u `modStorno.StornoOtkup` (`modStorno.bas:154`) · fizički red |
| **`LNC-INV-13`** | **Članstvo posle poteza:** „Dodaj" → blok je u `IzvoriOtpremnice(otpremnica')`; „Izvadi" → `OtpremnicaZaOtkup(blok)` = "" i blok je u `NevezaniOtkupi`, a „bila u" **sadrži** broj stare otpremnice (spaja sve ranije, `BivseOtpremniceIzvora`, `modDokumenta.bas:6695`). Staro članstvo ostaje (`LNC-INV-03`); `AktivnoOtpClanstvoPoKanonu` ne diže 1305 ni 1333 | pisci `OtpUpisiClanstvo` / `LncSastavi`; strogi čitač `AktivnoOtpClanstvoPoKanonu` · logički dokument |
| **`LNC-INV-14`** | **Suženje `LNC-INV-09` za „Izvadi":** posle poteza je **tačno** izvađeni blok van svake aktivne otpremnice; svaki drugi list stare otpremnice je u otpremnici', a svaka izdata sestra stare zbirne u zbirnoj'. Za „Dodaj" `LNC-INV-09` važi neizmenjen | `LncSastavi` (izvori = preostali ± blok; zbirna' = sestre + otpremnica') · logički dokument |
| **`LNC-INV-15`** | **Trag poteza:** tačno jedan red otpremnice (koren, `LNC-ODL-27`) i najviše jedan red zbirne (dete), oba završena u istoj transakciji (`LNC-INV-10`); nijedan red sa `OldDocType` Otkup. Trag po ID-u na oba kraja; `ZamenjenSaID <> ""` ⇒ `Stornirano = Da`. Kad je potvrđeno upozorenje, `Message` korena ga sadrži (`LNC-ODL-38`) | `modStornoContext` kroz `LncKontekstPar` (postaje funkcija koja vraća `CorrectionID`); trag piše `LncSastavi`; `LancaRazlog` nepromenjen · fizički red |
| **`LNC-INV-16`** | **Upozorenje se ne razilazi sa piscem:** X iz potvrde = zbir stavki zbirne' posle commit-a; Y = zbir prijemnica sa `ZbirnaRoditeljID` = zbirna' posle commit-a (prijemnice nepromenjene). `PAZNJA` ⇔ aktivna prijemnica ili razlika po `LNC-ODL-21` | plan u `modDokumenta`: X = `StavkeZaOtpremnicu` po sestri (isti primitiv kao `CreateZbirna`) + listovi otpremnice' (skup izvora, jer otpremnica' u planu još ne postoji, `LNC-ODL-28`); Y = zbir stavki prijemnica po `ZbirnaRoditeljID`. Jednakost sa piscem drži T5 · poslovni broj (kg) |
| **`LNC-INV-17`** | **Ambalaža neto:** menjaju se samo redovi dokumenta Otpremnica — kontra-stav stare i knjiženje nove. Za otpremnicu usklađenu sa listovima: „Izvadi k": stanica +k, vozač −k; „Dodaj k": stanica −k, vozač +k; kooperant i kupac 0. Neusklađena (`LNC-ODL-21`) pomera i razliku između gajbi koje je stara knjižila i zbira listova — tu razliku već imenuje `PAZNJA`. Posle cele zamene nijedan realan nalog nije u minusu (`AMB-INV-07`). `AMB-INV-09` i `-16` su ovde prazne po konstrukciji i test **ne sme** da tvrdi da ih pokriva | `modAmbalaza` (`StornirajAmbalazuDokumenta` u grani odlaganja, `AmbOdloziInvarijanteZavrsi` → `RequireSaldoNaloga`); odlaganje otvara koordinator · saldo naloga po tipu |
| **`LNC-INV-18`** | **Kapije plana = kapije pisca** za sve što plan može da izmeri — nacrt i status, hladnjača i N1, kapije dodatog bloka i datum, poslednji blok uz prijemnicu, identitet cilja i bloka, prijemnica samo po labeli: razlog stiže **pre potvrde**, a `_TX` vraća isti bez ijednog upisa. Dva odbijanja postoje **samo u transakciji**, posle „Da": ambalaža na kraju (`AMB-INV-07`, `AmbOdloziInvarijanteZavrsi`) i provera robe prijemnice (`PrijemnicaRobaRazlog`); za njih je jemstvo potpun povraćaj ~13 tabela (`LNC-INV-01`) | jedno telo u `modDokumenta` (razlog-oblik `OtpRequireIzvorValjan` + jezgro plana), dva pozivaoca; atomičnost drži `LncSnimi` |
| **`LNC-INV-19`** | **Radnja se nudi samo u kontekstu u kome ima smisla:** „Dodaj" samo uz izabran izdat cilj i samo nad listom slobodnih blokova; „Izvadi" samo nad listom blokova izabranog cilja. Radnje su po listi, ne po redu (ugovor ljuske, `modScrDokumenti.bas:217-221`), pa neprimeren blok odbija plan imenom (`LNC-INV-18`). Nijedna lista ni u jednom kontekstu ne nudi više od `MAX_ACT` radnji | `modScrDokumenti.Scr_Radnje`; granica `modOtkupUI.MAX_ACT` |
| **`LNC-INV-20`** | **Utovar i izvoz:** otpremnica' nosi `PredajaID` stare, pa `Chk_B11` ne nalazi utovar sa dva aktivna dokumenta; izvoz vozaču nema staru verziju a ima otpremnicu' sa njenom aktivnom zbirnom; `TekucaPredajaOtkupa(blok)` daje otpremnicu' posle „Dodaj" i "" posle „Izvadi"; ponovljen PRED red istog utovara za blok u otpremnici' je Duplicate | `OtpZaglavljeVerzije`; `modMasterSync.OtpremnicaPoPredaji`; `modIntegritet.Chk_B11`; `modStammdatenSync` (`:828-885`, `:1006-1034`) |
| **`LNC-INV-21`** | **Nijedan nedovršen kontekst ne pokazuje na verziju koju je potez zamenio:** posle poteza nema reda sa `NeedsRecovery = Da` čiji je `OldDocID` stara otpremnica ili stara zbirna, niti nedovršene ispravke prijemnice čija je zbirna stornirana bez naslednika | kapija plana (`LNC-ODL-37`, `LncNedovrsenKontekstRazlog`), jedno čitanje po planu kroz `modStornoContext.GetPendingCorrections` (vraća `oldDocID`, a grešku čitanja ili kolone koje nema javlja kroz `outGreska`, koju kapija čita fail-closed) · fizički red |

`LNC-INV-01..08` i `-10` važe neizmenjeni; `-09` sužava `-14`, `-11` važi po konstrukciji (`LNC-ODL-26`). `LNC-INV-21` i `LNC-ODL-37` važe i za postojeće poteze (§1–§9).

### 10.4 Matrica stanja

| Otpremnica | Zbirna | Prijemnica | Potez | Ishod |
|---|---|---|---|---|
| nacrt | — | — | bilo koji | radnja se ne nudi, plan odbija; nacrt ostaje na „vezi" / „ukloni" (nacrt ne može biti izvor zbirne, `RequireOtpValidanIzvorZbirne`) |
| izdata | nema | — | Dodaj / Izvadi | otpremnica' = listovi ± blok, nov broj, isti datum i `PredajaID`; 1 red traga (ISPRAVKA, koren) |
| izdata | nema | — | Izvadi poslednji | otpremnica bez naslednika; 1 red PONISTENJE (koren); blok slobodan |
| izdata | nacrt | — | Dodaj / Izvadi | otpremnica' kao gore; u nacrtu zbirne stara izlazi a otpremnica' ulazi **u mestu**, bez novog broja zbirne (`LNC-ODL-13`); trag samo za otpremnicu |
| izdata | nacrt | — | Izvadi poslednji | otpremnica bez naslednika izlazi iz nacrta; prazan nacrt gubi vrstu, sortu i tip (`ZBR-KANON-04`) |
| izdata | izdata | nema | Dodaj / Izvadi | otpremnica' + zbirna' (sestre + otpremnica'), nov broj zbirne, `KupacHladnjaca` nasleđen; trag: koren + dete |
| izdata | izdata | nema | Izvadi poslednji | otpremnica bez naslednika; sa sestrama zbirna' = sestre; bez sestara i zbirna bez naslednika |
| izdata | izdata | 1..N aktivnih | Dodaj / Izvadi | **dozvoljeno**: `PAZNJA` sa X i Y (`LNC-ODL-24`), „Ne" podrazumevano; posle „Da" sve prijemnice prevezane po `ZbirnaRoditeljID`, kg i stavke im se ne menjaju; `AMB-INV-07` na kraju može da odbije uz povraćaj |
| izdata | izdata | 1..N aktivnih | Izvadi poslednji, uz sestre | dozvoljeno sa `PAZNJA`; zbirna' = sestre, prijemnice prevezane |
| izdata | izdata | 1..N aktivnih | Izvadi poslednji, bez sestara | **odbija se pre potvrde** → F8 PONIŠTENJE (`LNC-ODL-25`) |
| hladnjačka | — | — | bilo koji | odbija se pre potvrde → F8 PONIŠTENJE (`LNC-ODL-14`) |
| izdata | bilo koja | — | Dodaj bloka: zauzetog (nacrt ili izdata), storniranog, bez stavki, neizdatog, druge stanice / kulture / tipa gajbi, hladnjačkog ili neizvesnog, sa datumom posle otpremnice; ili sa gajbama u otpremnicu bez tipa | odbija se pre potvrde i ponovo u piscu (`LNC-ODL-29`, `-30`) |
| izdata (ručni vozač, N1) | bilo koja | — | Izvadi bloka za koji važi `LanacVaziZaBlok` | odbija se → F8 storno lista u lancu |
| bilo koja | bilo koja | — | blok više nije u cilju / više nije slobodan; cilj storniran ili zamenjen posle liste | odbija se po ID-u (`LNC-ODL-34`) |
| neusklađena sa izvorima | bilo koja | — | dozvoljen potez | `PAZNJA` po `LNC-ODL-21` sa tekstom ovog poteza; nova verzija nosi zbir izvora |
| bilo koja | izdata | prijemnica samo po labeli | bilo koji | odbija se (`LNC-INV-07`) |
| bilo koja | nacrt ili izdata, sa nedovršenim kontekstom (`NeedsRecovery = Da`) ili nedovršenom ispravkom prijemnice | — | bilo koji | odbija se pre potvrde, imenuje `CorrectionID` → Oporavak / F4 (`LNC-ODL-37`) |

Stanja „izdata otpremnica bez zbirne uz prijemnicu" i „nacrt zbirne uz prijemnicu" ne postoje
(`PrijemnicaRoditeljRazlog`, `PRJ-ZBR-01`).

### 10.5 Poslovni događaji

| Događaj | Šta se dešava | Datum | Status / tabela |
|---|---|---|---|
| **fizički** | **ništa novo.** Potez ispravlja zapis o tome kojim kamionom je blok išao: putovao je ovim (Dodaj) ili nije (Izvadi) | blok zadržava svoj; „Dodaj" ne prima blok kasnijeg datuma od otpremnice (`LNC-ODL-30`) | — |
| **poslovni / pravni** | otpremnica' i zbirna' sa novim brojevima, stare stornirane sa `ZamenjenSaID`; prijemnici se menja samo roditelj | datum stare verzije | `IZDATO`; `tblOtpremnica`, `tblZbirna`, `*Izvori`, `*Stavke`, `tblPrijemnica` (veza), `tblStornoVeze` |
| **finansijski** | **za novac ga nema:** dug prema kooperantu i isplate su na `OtkupID` (`LNC-ODL-26`), potraživanje od kupca je na stavci prijemnice (`LNC-ODL-08`). Knjiga ambalaže: kontra-stav stare i knjiženje nove otpremnice, neto ±k stanica / vozač (`LNC-INV-17`, `LNC-ODL-31`) | datum stare verzije | `tblAmbalaza` |

Dokument nasleđuje datum svog fizičkog događaja; nijedan datum ili status nije izjednačen zbog
lakše implementacije.

### 10.6 Nizvodno

| Potrošač | Posle poteza | Uslov / odluka |
|---|---|---|
| malina auto-otpremnica | sledeći sync od izvađenog bloka pravi auto-otpremnicu; „Dodaj u B" posle toga odbija (blok nije slobodan) | `LNC-ODL-36`, poruka u potvrdi |
| PRED predaja otkupca | otkupac vidi izvađen blok kao slobodan (`TekucaPredajaOtkupa`); nova predaja je **izričit izbor** otkupca (`src/js/features/otkup/otpremnice.js:873-888`) i pravi otpremnicu — legitiman fizički događaj | bez izmene; automatike ovde nema |
| vozačeva PWA zbirna koja nosi staru otpremnicu | uvoz odbija storniran izvor (`RequireOtpValidanIzvorZbirne`), red dobija `SyncError`, a `MarkPWAFatalSyncError` prekida **taj** ciklus pre izlaznog sync-a (`modMasterSync.bas:4400-4420`, `modGoogleSyncOrchestrator.bas:358-366`); sledeći ciklus red ne ponavlja (uvoz čita samo PENDING, `modMasterSync.bas:4348-4350`), pa vozač sastavlja zbirnu ponovo | **zatečeno ponašanje**, isto posle F1 „Ispravi", storna otpremnice i postojećeg poteza; potvrda ga imenuje (`LNC-ODL-35`) |
| F4 prijemnica po broju zbirne sa papira | roditelj se razrešava po broju nad **aktivnim** zbirnim (`ZbirnaIdentResolve`), pa stari broj više ne vodi na zbirnu' | potvrda i poruka posle poteza imenuju nov broj zbirne (`LNC-ODL-09`, `-35`) |
| Oporavak, izgubljeni blokovi | izvađen blok se vodi kao `IZGUBLJEN_BLOK` sa „bila u" (`modStornoRecovery.bas:136-145`) dok se ne doda | **tačan signal**, bez izmene |
| integritet A2 i izveštaj manjka | kg zbirne' se pomera za ± kg bloka uz nepromenjenu prijemnicu → VIŠAK ili MANJAK | **očekivano** (A13: razlika ostaje vidljiva), kao u §6 |
| nedovršen kontekst (F8 „Reši kasnije", pao PONIŠTENJE / DUPLI, nedovršena ispravka prijemnice) nad starom zbirnom | **izmereno čitanjem (T21, 09.11.2026):** danas ga potez ne vidi; posle poteza podsetnik pokazuje na storniranu zbirnu (Oporavak nudi samo „Odbaci", podaci nisu pogođeni), a uz storno jedinog izvora nedovršena ispravka prijemnice ostaje bez zbirne (KI-095). Nad otpremnicom ga F8 ne pravi | **`LNC-ODL-37`**: plan odbija i upućuje na Oporavak; važi i za postojeći potez. Isti nedostatak u F8 običnom stornu zbirne je KI-095, zaseban rez |
| izvoz vozaču / otkupcu, PRED istog `PredajaID`, faktura, SEF, palete, novac | bez izmene | `LNC-INV-20`, `LNC-ODL-08`, `LNC-ODL-26` |
| telefon za PWA zbirnu | i dalje pokazuje stari broj (`LNC-ODL-19`) | tekst poruke za zbirnu bez naslednika se ispravlja (`LNC-ODL-35`) |

### 10.7 Verdikt po osama (pre-flight skill §1)

„PROVEN" ovde znači da je osa **zatvorena u specifikaciji** (odluka, vlasnik, dokaz u kodu koji postoji);
**ne** znači da je novo ponašanje dokazano — testovi T1–T23 još ne postoje.


| Osa | Status | Dokaz |
|---|---|---|
| **DOMAIN** | **PROVEN** | `LNC-ODL-06`, `-07`, A13/A14/A15, H1/H2; odluke operatera 08. i 09.11.2026 (`LNC-ODL-22..25`, `-30`, `-33`, `-36`); matrica 10.4 |
| **IDENTITY** | **PROVEN uz uslov** | `OtkupID` i `OtpremnicaID` putuju redom do pisca, pisac proverava članstvo (`LNC-ODL-34`); prijemnice po `ZbirnaRoditeljID`; **uslov:** koren na paru otpremnice, ne na bloku (`LNC-ODL-27`), i cilj odvojen od `mOtpID` (`LNC-ODL-33`) |
| **CARDINALITY** | **PROVEN** | blok ∈ 0..1 nestornirane otpremnice (1305 / 1333, `AktivnoOtpClanstvoPoKanonu`); otpremnica ∈ 0..1 aktivne zbirne; nacrt ∉ zbirna; zbirna → 0..N prijemnica (dve kod 5 od 894 na C002, §7) — 1:N ide u test (T6) |
| **INVARIANTS/OWNER** | **PROVEN** | `LNC-INV-12..20` sa vlasnikom; suženje `-09` (`-14`); `AMB-10-ODL-40` proširen (`LNC-ODL-31`, `AMBALAZA.md` §6.14c) |
| **WRITERS** | **PROVEN uz uslov** | potez piše otpremnicu, zbirnu, izvore, stavke, vezu prijemnice (`modDokumenta`), ambalažu (`modAmbalaza` kroz `StornoOtpremnica` / `OtpIzdaj`), `tblStornoVeze` (`modStornoContext`); **ne** piše otkup, novac ni žurnal. **Uslov:** koordinator u `modDokumenta`, sve kroz API vlasnika; ne zovu se pisci sa sopstvenom transakcijom (`DodajOtpremnicaIzvor_TX`, `IspravkaOtpremnice_TX`); `who_writes.py --check-ownership` ostaje zelen bez izmene `WRITE_OWNERSHIP.json` |
| **DOWNSTREAM** | **PROVEN** | 10.6; nedovršen kontekst izmeren čitanjem (T21) i zatvoren odlukom `LNC-ODL-37`; izvršna potvrda je T21 |
| **CAPABILITY** | **PROVEN** | nova sposobnost B-049 (`MAPA_SPOSOBNOSTI.md`); **ostaje netaknuto:** F1 „Ispravi" (B-022), „vezi" / „ukloni" / „izdaj" nad nacrtom (A-024, A-025), F8 zbirna DUPLI / PONIŠTENJE / REŠI KASNIJE |
| **ACCEPTANCE CONTRACT** | **PROVEN (plan)** | 10.8 |
| **PLATFORM** | **N/A** | nema novog Excel/VBE/MSForms ponašanja: dugmad kroz postojeći bazen `Scr_Radnje`, potvrda kroz seam; izgled ostaje ručnoj proveri |
| **LANDING** | **PROVEN uz uslov** | 10.9 |
| **EVENTS** | **PROVEN** | 10.5 |

⚠ **IDENTITY RISK** (zatvoren odlukom): ponovna upotreba `LncKontekstKoren` bi upisala `CorrectionID`
na blok koji ostaje aktivan, a `LancaRazlog` bi onda odbio UNDO njegovog kasnijeg običnog storna
(`LNC-ODL-27`; meri se nad blokom bez gajbi — sa gajbama je storno ionako konačan, `AMB-10-ODL-25`).
⚠ **CAPABILITY RISK** (zatvoren odlukom): izdata otpremnica u `mOtpID` bi svaki nov list
pokušala da veže za sebe, a pisac bi odbio (1309) — blok bi ostao slobodan uz grešku vezivanja
(`LNC-ODL-33`).
⚠ **FALSE-GREEN RISK:** test nad `NextTestDate` fixture-om (+1 dan po zaglavlju) ne meri pravilo
datuma; datumi bloka i otpremnice se zadaju izričito.

### 10.8 Dokaz — plan (testovi još ne postoje)

**Šta važi:** potez iz 10.4 sa invarijantama 10.3, jedan potez i jedna transakcija, sa ulazom
iz `LNC-ODL-33`. Testovi su u `RunBusinessFlowProSuite`; lanac grade kanonski pisci (`LncTipAmb`,
`LncOtkup`, `LncOtp`, `LncZbr`, `LncPrj`), svaki scenario sa svojim tipom gajbe; fixture je
*normal production state*, osim gde je izričito označeno. Svaki test imenuje nivo koji meri.

| Test (BFP) | Tvrdnja |
|---|---|
| T1 `Test_LNC_Blok_DodajIzdataBezZbirne` | otpremnica' = stari listovi + blok (logički dokument); nov broj iz generatora, isti datum i `PredajaID`; stara storno + `ZamenjenSaID`; tačno 1 red traga (koren LANAC, `OldDocType` Otpremnica); red `tblOtkup` bloka identičan (fizički red) |
| T2 `_IzvadiIzdataBezZbirne` | `OtpremnicaZaOtkup(blok)` = ""; „bila u" sadrži stari broj; blok nije storniran; isplata ista; staro članstvo isto (A15) |
| T3 `_IzvadiPoslednjiBezZbirne` | otpremnica bez naslednika, 1 red PONISTENJE, blok slobodan |
| T4 `_NacrtZbirneUMestu` | nacrt zbirne menja člana u mestu; broj i najava isti; nova zbirna ne nastaje. Varijanta „Izvadi poslednji": prazan nacrt gubi vrstu, sortu i tip (`ZBR-KANON-04`) |
| T5 `_DodajUZbirnuSaPrijemnicom` | zbirna' = sestre + otpremnica'; prijemnica prevezana (oba polja), kg i stavke isti; **opis plana** (iz `modDokumenta`, kao `LncPlanOpisZaOtkup` u `Test_LNC_OtpremnicaNeuskladjenaSaListovima`) počinje `PAZNJA`; **X iz opisa = zbir stavki zbirne' posle commit-a**; trag: tačno 2 reda — otpremnica je koren (LANAC, prazan `ParentDocID`), zbirna nosi `ParentDocID` = `CorrectionID` korena, oba završena |
| T6 `_IzvadiDvePrijemnice` | obe prijemnice prevezane, Y = njihov zbir; treća prijemnica sa istom labelom a drugim roditeljem ne ulazi u Y i ne seli se. Varijanta: aktivna prijemnica samo po labeli → odbijanje (`LNC-INV-07`) |
| T7 `_IzvadiPoslednjiJediniIzvorSaPrijemnicom` | plan vraća razlog pre potvrde; `_TX` odbija; ~13 tabela nepromenjeno |
| T8 `_IzvadiPoslednjiUzSestre` | otpremnica bez naslednika, zbirna' = sestre, prijemnica prevezana. Varijanta bez sestara i bez prijemnice: i zbirna bez naslednika |
| T9 `_DodajOdbijen` | za svaki slučaj iz 10.4 (zauzet u nacrtu i u izdatoj — poruka imenuje otpremnicu; storniran; bez stavki; neizdat; druga stanica / kultura / tip; hladnjački; neizvesnost; hladnjački cilj; nacrt kao cilj; datum posle otpremnice; gajbe u otpremnicu bez tipa): razlog iz plana pre potvrde **i** isti razlog iz `_TX` bez upisa |
| T10 `_IzvadiOdbijen` | nacrt kao cilj; hladnjačka otpremnica; N1 blok; blok koji nije u cilju; cilj storniran ili zamenjen |
| T11 `_AmbalazaNeto` | „Dodaj k" i „Izvadi k" po nalozima (saldo po tipu), nad otpremnicom usklađenom sa listovima; kontrola: bez odlaganja „Dodaj" pada kad je vozač predao gajbe, sa odlaganjem prolazi |
| T12 `_IzvadiPosleIsporukeGajbiVracaSve` | `AMB-INV-07` na kraju odbija, sve se vraća (ogledalo `Test_LNC_NeslaganjeAmbalazeVracaSve`) |
| T13 `_DodajStanicaBezGajbi` | stanica nema k gajbi tog tipa → odbijanje uz povraćaj |
| T14 `_NeuskladjenaStaraOtpremnica` | `PAZNJA` po `LNC-ODL-21` sa tekstom ovog poteza, bez „storniranog lista" |
| T15 `_BezUndoOperacije` | nijedan nov red `tblStornoZurnal`; kasniji običan storno izvađenog bloka ostaje povratan — blok **bez gajbi** (kao `Test_LNC_UndoPotezaOdbijen`), inače `KontraStavRazlog` odbija iz drugog razloga |
| T16 `_PremestiADoB` | dve zbirne: između koraka blok je slobodan, na kraju nove verzije sa novim brojevima; ista zbirna: Z → Z' → Z'', prijemnica na Z'' |
| T17 `_UtovarIIzvoz` | `LNC-INV-20` |
| T18 `_EkranPotvrdaIRuta` | cilj se bira produkcionim putem (događaj nad mrežom, ne `Scr_OtpTestSet`); radnje samo u pravom kontekstu; tekst potvrde kroz seam (`PAZNJA` na početku uz prijemnicu); „Ne" ne upisuje; ID iz nevidljivih kolona; posle snimanja novog lista uz izdat cilj `RutaPosleUpisa` **ne** vraća `NACRT` i pokušaja vezivanja nema (meri se ruta, ne članstvo — pisac bi odbio ionako, pa bi tvrdnja nad članstvom bila placebo); posle uspeha cilj = otpremnica'; `Scr_PotvrdaLancaTestReset` gasi seam, `Scr_OtpTestReset` briše cilj |
| T19 `_GranicaRadnji` | `Scr_Radnje` za svaku listu F1 i F2, sa ciljem i bez njega, ≤ `MAX_ACT` (u BFP, ne u `modTest`) |
| T20 `_MalinaPoruka` | u malina režimu potvrda „Izvadi" imenuje sledeći sync, van njega ne |
| T21 `Test_LNC_NedovrsenKontekstOdbija` (meri postojeći potez; „Izvadi" dobija svoj slučaj u a6) | **prvo, crven nad današnjim kodom** (KI-095, `CLAUDE.md` §2): (S7) nedovršena ispravka prijemnice + storno jedinog lista u lancu — danas zbirna pada bez naslednika i zamena prijemnice nema roditelja; posle popravke plan odbija pre potvrde. (S1) „Reši kasnije" nad zbirnom + potez sa naslednikom — posle popravke odbija i imenuje `CorrectionID`. Meri i da stornirana prijemnica zadržava `ZbirnaRoditeljID` (osnov veze, nemereno). Isto za „Izvadi" kad postoji. „Reši kasnije" nad aktivnom prijemnicom ne smeta. Granice (sintetički i stari oblici) meri `Test_LNC_NedovrsenKontekstGranice`: PENDING ispravka nad aktivnom prijemnicom, zamena već uneta (MANUAL_REQUIRED), kontekst nad otpremnicom, red bez `OldDocID` (i stari MANUAL istog broja koji NE odbija), ispravka prijemnice bez `OldDocID`, prijemnica bez `ZbirnaRoditeljID`, red koji ne može da se proveri (odbija i imenuje sebe) |
| T23 `_UpozorenjeUTragu` | posle potvrđenog `PAZNJA` poteza `Message` korenskog reda sadrži tekst upozorenja **i** opis poteza (`LNC-ODL-38`), za storno i za ispravku lista u lancu (F1 „Ispravi blok"); „Dodaj" / „Izvadi" dobijaju svoj slučaj u a6/a7 (koren na paru otpremnice, `LNC-ODL-27`); bez upozorenja `Message` je samo opis |
| T22 `_UpozorenjePrijemniceSviPotezi` | `LNC-ODL-24` važi i za F8 storno lista, F1 „Storno" i F1 „Ispravi blok" nad zbirnom sa prijemnicom; bez prijemnice usklađen potez ostaje bez `PAZNJA` |

**Mora ostati netaknuto:** svi postojeći `Test_LNC_*` i `Test_AMB_ODL40_*`;
`Test_OTP_RadniStoBiraSamoNacrt` (izdata se ne bira kao `mOtpID`);
`T_Zbirna_NemaIspravku`; F1 „Ispravi" (`IspravkaOtpremnice_TX` i dalje odbija izvor aktivne
zbirne); „vezi" / „ukloni" nad nacrtom; `StornoOtkup` i dalje odbija list u aktivnoj
otpremnici; `StornoOtpremnica` i dalje odbija izvor aktivne zbirne; semantika `LancaRazlog`;
kapija hladnjače u `VeziZaAktivnu`; malina sync.

**Premisa koja se meri ponovo:** tvrdnje o tekstu potvrde u
`Test_LNC_OtpremnicaNeuskladjenaSaListovima` (`modBusinessFlowProTests.bas:27487-27503`) —
upozorenje o prijemnici ide **posle** bloka `LNC-ODL-21`, a oznaka `PAZNJA:` ostaje jedna, na
početku.

**Sabotaže** (samo nove, prefiks `lncb-`, pa `python tools/dokaz.py lncb`):
`lncb-dodaj-bez-odlaganja` → T11 · `lncb-izvadi-stornira-list` → T2, T15 ·
`lncb-koren-na-bloku` → polje CorrectionID bloka, `lncb-izvadi-koren-lista` → T15 (UNDO) · `lncb-dodaj-bez-kapije-hladnjace` → T9 ·
`lncb-kapija-samo-u-tx` → T9 (tvrdnja pre potvrde) · `lncb-dodaj-bez-datuma` → T9 ·
`lncb-dete-bez-korena` → T5 · `lncb-prijemnica-ostaje-na-staroj` → T6 · `lncb-y-po-labeli` → T6 · `lncb-x-iz-drugog-tela` → T5 ·
`lncb-poslednji-izvadi-bez-kapije` → T7 · `lncb-ne-upisuje-na-ne` → T18 ·
`lncb-izdata-u-ruti` → T18 · `lncb-upozorenje-samo-novi`, `lncb-x-bez-preostalih`, `lncb-x-bez-novog-lista`, `lncb-x-bez-sestara`, `lncb-tx-ispravka-bez-stavki`, `lncb-f1-bez-polja`, `lncb-f1-save-bez-polja`, `lncb-upozorenje-bez-prijemnice`, `lncb-upozorenje-druga-oznaka`, `lncb-dijalog-bez-granice` → T22 (`Test_LNC_UpozorenjePrijemniceSviPotezi`) · `lncb-pending-*` (jedanaest: zbirna, otpremnica, prijemnica nevidljiva, samo PENDING, prijemnica u svakom statusu, aktivna prijemnica, bez broja, stari MANUAL, prijemnica bez roda, prijemnica bez OldDocID, nepoznata prijemnica) → T21 i granice · `lncb-trag-storno-bez-upozorenja`, `-ispravka-bez-upozorenja`, `-prepisuje-opis`, `-uvek` → T23 (`Test_LNC_UpozorenjeUTragu`).

**FALSE-GREEN pravila:** nijedan test ne tvrdi `AMB-INV-09` ni `-16` (ovde prazne); članstvo se
ne meri kroz `Otkup.OtpremnicaID` ni broj otpremnice; pravilo datuma se ne meri nad
`NextTestDate`.

**Nemereno do Excela:** sve ponašanje. Redosled posle koda: `vba_check`,
`gen_schema_module.py --check`, `who_writes.py --check-ownership`, `run_vba.py --suite` za BFP,
pa jednom `dokaz.py lncb`; Compile i ručna provera (izgled dugmadi, dijalozi nad pravim
podacima) ostaju operateru.

**Ručna provera (UAT), posle oba dela:** (1) premeštanje A → B u malina režimu sa sync-om između
koraka — da li poruka kaže šta da se uradi i koliko se to dešava (okidač `LNC-ODL-23`); (2) operater
razlikuje F1 „Ispravi" (izdata van zbirne, zaglavlje) od „Dodaj" / „Izvadi" (članstvo, i u zbirnoj)
i ne bira pogrešnu radnju; (3) dijalozi `PAZNJA` nad pravim prijemnicama i starim neslaganjima, Enter =
„Ne"; (4) odbijanje zbog podsetnika u Oporavku vodi do pravog reda i posle „Odbaci" potez prolazi.

### 10.9 Landing

- **Prvo ovaj zapis (samo docs)**, pa kod u dva dela, svaki zelen sam za sebe:
  **(a)** `modDokumenta` — jezgro plana po otpremnici, razlog-oblik `OtpRequireIzvorValjan`,
  `LncKontekstPar` kao funkcija, dva `_TX` ulaza po skeletu `StornoListaULancu_TX`, tekstovi
  poteza i telo X/Y (`LNC-ODL-24` za sve poteze) — BFP T1–T17, T21–T23 (tekst se meri nad opisom plana, ne kroz seam); T21 ide **prvi**, crven nad
  današnjim kodom, pa popravka. **Odstupanje (a4, 09.11.2026):** tačan X u F1 potvrdi ispravke traži
  polja forme, pa deo (a) dira i `modScrDokumenti.PotvrdiIspravkuULancu` / `Scr_PotvrdaLancaProbaj`
  (opcioni parametar polja) i `modOtkupUnos.OtkupStavkeIzForme` (jedno telo za upis i potvrdu);
  sidro sabotaže `lnc-f1-bez-potvrde` je pomereno, pa ona ide u dokaz;
  **(b)** `modScrDokumenti` — kontekst cilja, radnje, potvrda kroz postojeći seam (reset kroz
  `Scr_PotvrdaLancaTestReset` / `Scr_OtpTestReset`), ključevi u `modPoruke` — T18–T20.
- `modDokumenta`, `modScrDokumenti`, `modBusinessFlowProTests` i ovaj fajl ne dira nijedan
  otvoren PR. Sa #457 se dele `modPoruke.bas`, `tools/sabotaza.py`, `AMBALAZA.md`, `README.md`
  (`README.md` dira i #469); #457 dira i `modOtkupUI.bas` (mreža, ne `MAX_ACT`), `modTest.bas` i
  `KNOWN_ISSUES.md` — zato T19 ide u BFP i samo čita `MAX_ACT`, a ovaj rez `modOtkupUI.bas` i
  `modTest.bas` ne dira.
- Ovde se ne dodeljuju KI, broj `modTest` testa ni stavka `STANJE_REFAKTORA.md`. Testovi
  `Test_LNC_Blok_*`, sabotaže `lncb-*`.
- **S9:** ovaj rez ide pre S9, kao i §9. PR #469 posle preslaganja meri i koren LANAC čiji je par
  otpremnica (`LNC-ODL-27`, `SLED-ODL-03`).
- **Ispravka naziva (09.11.2026):** ovaj fajl je „Ispravi" izdate otpremnice zvao „F2 Ispravi";
  radnja stoji u F1, lista „Otpremnice" (`modScrDokumenti.bas:203-214`), pa je naziv ispravljen
  u §0, §3, §6, §7, §8.
