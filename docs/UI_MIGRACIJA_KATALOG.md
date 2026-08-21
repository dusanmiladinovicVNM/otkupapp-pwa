# Katalog funkcija legacy formi i plan prelaska na novi UI

> Izvor istine za migraciju `frmOtkup` i `frmDokumenta` na `frmOtkupUI`
> (ljuska `modOtkupUI` + ekranski moduli `modScr*`).
>
> Cilj dokumenta: **ništa se ne gubi u prelasku.** Svaka sposobnost legacy
> forme je ovde popisana, sa oznakom da li je u novom UI-ju već obezbeđena,
> delimično obezbeđena ili nije. Plan na kraju radi samo po ovom spisku.

Stanje na dan `v6-ui-171`.

---

## 0. Šta još NIJE pokriveno (sažetak)

Faza A je pokrila **pravila unosa**. Ostalo je, po veličini:

1. ~~**Upis** — svi unosni režimi (F1–F7) sada knjiže.~~ **ZATVORENO**
   (v6-ui-117). Posao je izvučen iz formi u `modOtkupUnos` (F1, v6-ui-106),
   `modDokUnos` (F2 v6-ui-115, F3/F4 v6-ui-116) i `modNovacUnos` (F5/F6/F7,
   v6-ui-117): provere, bruto→neto, `Save*_TX`, štampa, auto-lanac hladnjače,
   auto-zbirna, završetak ispravke.
   **Legacy zadržava svoju kopiju te logike — namerno.** `frmOtkup` i
   `frmDokumenta` ostaju potpuno operativni dok novi UI ne bude umeo sve; do
   tada se pravilo menja u zajedničkom modulu pa **ručno preslikava** u legacy.
2. ~~**Storno bilo kog tipa dokumenta**~~ **ZATVORENO** (v6-ui-119). F8 je
   postao storno centar: prekidač bira **tip dokumenta** (devet tipova, kroz
   sedam tabela), radnja nad redom zove tačan `Storno*_TX` kroz nov
   `modStornoDok`. Do tada je F8 čitao samo `tblOtpremnica` i pokazivao samo
   već stornirane.
   **Ostatak storno okvira je zatvoren posle toga:** ispravka i dupli unos
   posle storna (`modStornoFlow`, Z10) u v6-ui-120, a Undo operacija,
   „Nedovršeno" i Recovery u v6-ui-121, kroz nov ekran **Oporavak**
   (`modScrOporavak`). **Faza D je zatvorena tek od `v6-ui-130`**, ne od
   `v6-ui-121` kako je ovde ranije pisalo: ekran je bio gotov, ali je F8 do
   tada gubio identitet izabranog reda i dokument je nizvodno biran po
   poslovnom broju.
   **Od `v6-ui-143` storno više nije režim F8 nego SVOJ EKRAN** (`modScrStorno`,
   grupa OPERACIJE, oblast `OBL_DOKUMENTA`) — v. §3.2. Taster F8 je ostao, ali
   sada bira ekran, ne režim. Unosni ekran time ima **sedam** režima.
   Nevidljiva kolona identiteta se od tada traži **argumentom**
   (`GridCols(tip, saIdentitetom)`), a ne uslovom `ActiveMode = "F8"` — ekran
   nema režim, pa bi taj uslov ćutke bio `False` i ceo lanac iz #198 bi pao na
   biranje po broju. To meri test 57.
3. **Pomoćni delovi režima** koji nisu pravila nego zaseban posao: lista
   zbirnih za izbor (F3), manjak prijemnice vs zbirna (F4). Upis F3/F4 od
   v6-ui-116 radi i bez njih — to su prikazi, ne kapije. Ostatak te tačke je
   zatvoren u v6-ui-117: avans saldo OM, otvorene fakture (F6) i otvoreni
   otkupi (F5) sada postoje, jer bez njih upis ne bi bio tačan (v. §3.1).
4. **Sitno:** peščanik za vreme upisa i dva nevezana KPI-ja. Prefill iz
   storniranog (Z10) je zatvoren u v6-ui-120; filtriranje kooperanata po
   otkupnom mestu u v6-ui-113 (`KOOP_FILTER_BY_OM`).

**Faza D je zatvorena** (v6-ui-130). Ostaju 3 i 4 — ostaci ranijih faza — plus
Faze C i E, koje nisu počele.

**Identitet nije bio dovoljan samo za sam dokument.** Do `v6-ui-136` je storno
otpremnice mutirao **roditeljsku zbirnu po golom `BrojZbirne`** — rekalkulacija,
storno prazne zbirne, relink prijemnica. Nad dvosmislenim brojem roditelja to je
moglo da ažurira zaglavlje jednog dokumenta zbirom otpremnica oba. Kapija
`ZbirnaBrojJeDvosmislenIkad` stoji na četiri mesta, uključujući **završetak
ispravke** — jer correction context je persistentan i zatečen context preživljava
upgrade, pa kapija samo na startu ne pokriva njega.

Od `v6-ui-137` ta kapija proverava **istu vrednost koju kod mutira**: do tada je
roditelja tražila po poslovnom broju (`LookupValue` po `BrojOtpremnice`), a
mutacije su išle nad `ParentBroj` iz context-a — pa je proveravala zbirnu
siblinga. Roditelj se sada uzima iz context-a, fallback ide isključivo preko
tačnog `OldDocID`, a nerazrešen roditelj je MANUAL. Kapija je i **fail-closed na
sopstvenu grešku**: schema drift znači „ne mutiraj", ne „jednoznačno je".

Od `v6-ui-138` kapija stoji na **obe strane**: i nad ciljnom zbirnom, ne samo nad
izvornom. Zatečena kapija u writeru (`RequireJedanVlasnikPoBroju`) to ne pokriva
jer broji samo **aktivne** vlasnike — a storniran vlasnik i dalje ima aktivnu decu.
Ista rupa je zatvorena i u `CompleteZbirnaIspravka`, gde po broju idu i izvor i
cilj. Sam primitiv (`RecalculateZbirnaFromOtpremnice_TX`,
`ReassignPrijemnicaToZbirna_TX` bez generacije) ostaje number-based — tu bi jednog
dana trebala centralna kapija umesto zaštite po call-site-u.

Od `v6-ui-140` identitet nosi i **dodatni storno otkupnih blokova**
(`StornirajBlokoveAko → GetStornoBlockRows → ActiveBlocksForFlow`), plus pregledi
u `ScanOtpremnica` i `ScanPrijemnica`. Do tada je spisak blokova nastajao po
poslovnom broju i mogao je da obori blok drugog dokumenta — a kapija
`BlockStornoDriftReason` se na toj putanji ne izvršava, jer `ModeStornoBlokParent`
vraća `True` za `PONISTENJE` i za `OTPREMNICA+DUPLI/ISPRAVKA`, to jest za jedine
modove koji tu i dolaze. Grana `FLOW_DOC_ZBIRNA` ostaje po broju jer `tblOtkup`
nosi `BrojZbirne`, ne `ZbirnaID`; taj put je zaštićen uzvodno.

---

## 1. Zajednička pravila (važe u obe forme)

Ovo nisu dugmad nego **ponašanja** koja se u legacy formama ponavljaju kroz
event-handlere. U novom UI-ju svako od njih mora imati tačno jedno mesto.

| # | Pravilo | Legacy mesto | Poslovna rutina | Novi UI |
|---|---|---|---|---|
| Z1 | **Cena iz cenovnika** po vrsti+sorti+klasi, na svaku promenu vrste/sorte | `AutoFillCenaOtkup`, `AutoFillCenaDok` | `modCenovnik.GetVazecaCena(vrsta, sorta, klasa)` | **IMA** (`AutoFillCena`, v6-ui-102) |
| Z2 | **Tip ambalaže iz kulture** po vrsti+sorti | isto | `GetKulturaTipAmbalaze` | **IMA** (v6-ui-102) |
| Z3 | **Predlog broja dokumenta** — po režimu (tabela ispod); poštuje toggle auto-broja | isto | `modBrojevi.SuggestNextBroj(kind, entityID, datum)`, `GenerateBrojPrijemnice` | **IMA** (`RefreshBrojPredlog`, v6-ui-104; pravila po režimu v6-ui-112) |
| Z4 | **Živi zbir kg** uz polje količine; u bruto režimu prikazuje neto posle tare | `UpdateUkupnoKg`, `UpdateUkupnoKgOtp`, `UpdateUkupnoKgPrij` | `GetTezinaGajbice`, `OtkupBrutoUnos()` | **IMA** (`SetKgLine`, v6-ui-102) |
| Z5 | **Dve klase** — prekidač otvara drugi red količine/cene/ambalaže | `chkDveKlase*_Click`, `ShowKolAmbalazeII`, `ShowKlIIAmb` | — | **IMA** (segment I/II uz KLASA I CENA) |
| Z6 | **Parcele** — lista zavisi od kooperanta; celo polje gasi `PRACENJE_PARCELA` | `cmbKooperant_Change`, `ApplyOtkupTogglesState` | `IsPracenjeParcela()` | **IMA** |
| Z7 | ~~Keš isplate uz otkupni list~~ **NE PRENOSI SE** — keš isplate idu isključivo kroz F5 (Isplate) i F6 (Kupci-uplate); otkupni list ih više ne nosi | `ApplyOtkupTogglesState` | — | **NAMERNO IZOSTAVLJENO** (v6-ui-105); pri upisu `SaveOtkupMulti_TX` dobija `novac=0`, `primalac=""` |
| Z8 | **Blokada praznih polja** pri snimanju, gejtovana `VALIDACIJA_UNOSA` | `btnUnos*_Click` | `IsValidacijaUnosa()` | **IMA** (v6-ui-104); datum se proverava uvek |
| Z9 | **Info o paleti** — „još N gajbica do zatvaranja" uz izabranu robu | `UpdatePaletaInfo` | `GajbeDoZatvaranjaPaleteInfo` | **IMA** (`RefreshPaletaInfo`, v6-ui-104) |
| Z10 | **Prefill iz storniranog dokumenta** — ispravka posle storna | `PrefillOtkupFromStornirano`, `PrefillOtpremnicaFromStornirana`, `PrefillZbirnaFromStornirana`, `PrefillPrijemnicaFromStornirana` | `modDokumenta.PickPrefillRows` | **IMA** (`modStornoDok.PrefillIzStorniranog`, v6-ui-120) — jedan račun umesto četiri kopije vezane za kontrole |
| Z11 | **F-tasteri i Enter/Exit ivice** polja | `SetupFkeyAccelerators`, `HandleFkey`, `txt*_Enter/_Exit` | — | **IMA** (F1–F8 globalno, fokus ivice u `clsFlatBtn`) |
| Z12 | **KPI traka** iznad forme | `LayoutTopKpis`, `RefreshTopKpis`, `SumOtkupKgToday` | `GetOMAvansSaldo` | **DELIMIČNO** — traka postoji, dva KPI-ja nisu vezana |
| Z13 | **Podrazumevani proizvod** po otvaranju/resetu | `ResetProizvodNaDefault` | `ApplyDefaultProizvod` | **IMA** (`ApplyDefaultRoba`, v6-ui-103) |
| Z14 | **Kontekst datuma i otkupnog mesta** se pamti između dokumenata | `txtDatum_AfterUpdate`, `cmbOtkupnoMesto_Change` | `AcquireStanicaLock`, `GetActiveStanica/Datum` | **IMA** (v6-ui-104) |

### Z3a — brojevni niz po režimu (poslovno pravilo)

Svaki režim ima svoj niz i svoj **entitet** po kome se broji. Pogrešan entitet
ne daje samo pogrešan broj — `ApplyMirrorPrefix` gleda da li je entitet
mirror-vozač, pa je stanica podmetnuta kao vozač davala `S…` van zbirnih.

| Režim | Niz | Entitet (po čemu se broji) | Oblik |
|---|---|---|---|
| F1 otkupni list | `KIND_OTK` | otkupno mesto (`cbOM`) | `st/ddmmyy[-n]` |
| F2 otpremnica | `KIND_OTP` | otkupno mesto | `st/ddmmyy[-n]` |
| F3 zbirna | `KIND_ZBR` | **vozač** (`cbVozac`) | `st/ddmmyy[-n]`, sa `S` prefiksom kad je vozač mirror-vozač OM-a — **`S` postoji samo ovde** |
| F4 prijemnica | `GenerateBrojPrijemnice` | **kupac**; auto **samo** za hladnjaču (`MALINA_DEFAULT_KUPAC`) | `1/ddmmyy[-n]` (x-deo fiksno `1`); ostali kupci → **slobodan unos**, polje se ne dira |
| F5/F6 isplate/uplate | — | — | **slobodan unos** |
| F7 revers | `KIND_REV` | otkupno mesto | isto kao otpremnica, skenira `tblAmbalaza` |
| F8 storno | — | — | nema broj (nije nov dokument) |

Predlog se preračunava na promenu: otkupnog mesta, datuma, režima, vozača (F3),
kupca (F4) i posle svakog upisa.

### Z3c — isporuka: novi UI NE ide kroz online self-update

`frmOtkupUI.frm/.frx` je **nova forma**. `modSelfUpdate` namerno razlikuje dva
slučaja: postojeća forma → code-merge; **nova forma ili sheet → `needsReinstall`**.
Zaštita postoji jer runtime `Remove`/`Import` forme ume da korumpira workbook —
i **ne sme se slabiti** da bi ovaj UI prošao kroz update.

Posledica za rollout: prelazak na novi UI je **jednokratna puna isporuka**
(nov `AgriX_OtkupApp.xlsm` ili `ImportAllVBA` na svakoj mašini), ne online update.
Sve kasnije izmene shell-a i ekrana (`modOtkupUI`, `modScr*`, `modUiKit`,
`modUiScreens`) su „meki" moduli i idu normalnim self-update-om.

Zato `OtkupUI_SelfCheck` od v6-ui-114 **ne traži jednake verzije** nego
`>= OTKUI_MIN_BUILD`: forma stiže punom isporukom, moduli se posle menjaju
nezavisno, a `clsFlatBtn` je namerno zamrznut — jednakost bi prijavljivala
neispravnu instalaciju i kad je ispravna.

### Z3b — revizija ulaznog sloja 1:1 (v6-ui-113)

Kompletan popis legacy handlera koji nešto **automatski popune ili preračunaju**,
i status u novom UI. Ovo je izvor istine za „šta još nije preneto" — ne
zaključivati iz koda.

**`frmOtkup` (F1)**

| Legacy | Šta radi | Novi UI |
|---|---|---|
| `chkDveKlase_Click` | II klasa on/off + auto-cena II + paleta info | **IMA** (`SetKlasa`) |
| `txtKolicina/KolAmbalaze/TipAmbalaze/KolicinaKlII_Change` → `UpdateUkupnoKg` | živi zbir / neto | **IMA** (`RecalcVrednost`, `SetKgLine`) |
| `cmbVrstaVoca_Change` / `cmbSortaVoca_Change` | kaskada sorte + auto-cena | **IMA** |
| `AutoFillCenaOtkup` | cena I (uvek), cena II (samo uz II klasu), tip ambalaže, paleta info | **IMA** (`AutoFillCena`) |
| `ApplyOtkupTogglesState` | parcele / keš | parcele **IMA**; keš **namerno izostavljen** (Z7) |
| `cmbOtkupnoMesto_Change` | briše kooperanta+parcelu, filtrira kooperante po OM, lock, MALINA vozač, predlog broja (remote), prazna stanica briše broj, izlazak iz konteksta otpremnice | **IMA** (v6-ui-111 + v6-ui-113) |
| `cmbKooperant_Change` | puni parcele | **IMA** (`FillParcele`) |
| `cmbKooperant_Change` → `OtkupBlok_RefreshKoopTotal` | inline „ukupan iznos otk. listova" za kooperanta | **NEMA** (informativno) |
| `cmbParcela_Change` | parcela → vrsta/sorta iz kulture | **IMA** (v6-ui-113) |
| `txtDatum_AfterUpdate` | re-lock na nov datum + predlog broja (remote) | **IMA** (v6-ui-113) |
| `ResetDatumKontekst` / `ResetProizvodNaDefault` | izlazak iz konteksta otpremnice | **IMA** (`NapustiOtpremnicu`, v6-ui-111) |
| `ClearOtkupFields` | reset posle snimanja | **IMA** (v6-ui-108/109/110) |
| `FillKooperantCombo` | `KOOP_FILTER_BY_OM` | **IMA** (v6-ui-113) |
| `btnUnos_Click` | snimanje | **IMA** (`modOtkupUnos`) |
| `UserForm_QueryClose`, `btnPovratak_Click` | `ReleaseStanicaLock` pri izlasku | **IMA** (`OtkupUI_Sakrij`, v6-ui-114) |
| `ShowLockStatus` / `HideLockStatus` | peščanik za vreme sinhronizacije | **NEMA** |

**`frmDokumenta` (F2–F7)**

| Legacy | Šta radi | Novi UI |
|---|---|---|
| `cmbOtkupnoMesto_Change` | primalac po stanici, predlog otpremnice i reversa, MALINA vozač | **IMA** (isti handler) |
| `cmbVozac_Change` | predlog zbirne; prazan vozač briše broj | **IMA** (v6-ui-112/113) |
| `txtDatum_AfterUpdate` | predlozi otp/zbr/prij | **IMA** |
| `cmbVrstaVoca/SortaVoca_Change` → `AutoFillCenaDok` | cena + tip ambalaže za otp/zbr/prij | **IMA** |
| `chkDveKlaseOtp/Zbr/Prij_Click` | II klasa | **IMA** (`SetKlasa`) |
| `RefreshBrojOtp/Zbirne/Prij/ReversSuggestion` | predlozi po nizu | **IMA** (v6-ui-112) |
| `cmbKupac_Change` → broj prijemnice | briše pa predlaže | **IMA** (v6-ui-113) |
| `cmbKupac_Change` → `cmbHladnjaca` / `cmbPogon` | odredište otpremnice | **NEMA** — zato `SaveZbirnaMulti_TX` iz novog UI-ja dobija prazne `hladnjaca`/`pogon` |
| `cmbKupac_Change` → `FillOpenFakture`, `cmbFakturaIzlaz_Change` | otvorene fakture (F6) | **IMA** (v6-ui-117, `FillOpenFakture` uz polje `fgFaktura`) |
| `txtBrojZbirnePrij_AfterUpdate` → `UpdateManjak` | manjak prijemnice vs zbirna | **NEMA** — Faza B (živi prikaz; upis F4 ne zavisi od njega) |
| `UpdateValidacija` (živi prikaz + kapija pri upisu zbirne) | poklapanje zbirne sa otpremnicama | **kapija IMA** (`ZbirnaValidiraj`, v6-ui-116), **živi prikaz NEMA** |
| `lstZbirne_Click` | izbor zbirne iz liste (F4) | **NEMA** — Faza B (F3 klikom na red pamti aktivnu zbirnu, pa je F4 nasledi) |
| `cmbPrimalacOMUlaz_Change` → `UpdateOMAvansSaldo` | avans saldo OM (F5) | **IMA** (v6-ui-117, u natpisu polja „ISPLATA IZ") |
| `tglIzOMAvansa` | keš iz OM avansa vs virman firme (F5) | **IMA** (v6-ui-117, polje `fgAvans`) |
| `btnUnosOtp_Click` | upis otpremnice (F2) | **IMA** (`modDokUnos`, v6-ui-115) |
| `btnUnosZbr_Click` | upis zbirne (F3) | **IMA** (`modDokUnos`, v6-ui-116) |
| `btnUnosPrij_Click` | upis prijemnice (F4) | **IMA** (`modDokUnos`, v6-ui-116) — **bez ispravke posle storna** (v. §3.1) |
| `btnUnosOMUlaz_Click` | upis F5 (isplate) i F7 (reversi) | **IMA** (`modNovacUnos`, v6-ui-117) |
| `btnUnosIzlaz_Click` | upis F6 (uplate kupaca) | **IMA** (`modNovacUnos`, v6-ui-117) |
| `Prefill*FromStornirana` | ispravka posle storna | **IMA** (`modStornoDok.PrefillIzStorniranog`, v6-ui-120) |
| `btnStorno_Click` (storno po tipu i broju) | storno bilo kog dokumenta | **IMA** (`modStornoDok` + F8, v6-ui-119) |
| `TryRunCorrectionFramework` (četiri moda) | ISPRAVKA / DUPLI / PONIŠTENJE / REŠI KASNIJE | **IMA** (v6-ui-120) — kroz pitanja, ne kroz overlay panel |
| Undo operacija, „Nedovršeno", Recovery | — | **NEMA** — Faza D, stavka 14 |

---

## 2. frmOtkup (1.294 linije) — otkupni list

Jedan režim, jedna forma. Sve što radi:

| Procedura | Šta radi | Poslovna rutina | Novi UI |
|---|---|---|---|
| `UserForm_Initialize` | puni combo-e (vrsta, sorta, tip ambalaže, vozači, kooperanti) | `GetLookupList`, `GetTipAmbalazeOptions`, `GetVozacDisplayList` | IMA (`FillCombos`) |
| `SetupAmbIzdataField`, `SetupKolAmbalazeIIField` | runtime polja koja nisu u `.frx` | — | IMA (cela forma je runtime) |
| `cmbVrstaVoca_Change` | puni sorte za vrstu | `GetLookupList` | IMA (`RefillSorta`) |
| `cmbSortaVoca_Change` → `AutoFillCenaOtkup` | **cena iz cenovnika + tip ambalaže iz kulture + paleta info** | Z1, Z2, Z9 | IMA |
| `UpdateUkupnoKg` | živi zbir kg / neto iz bruta | Z4 | IMA |
| `cmbOtkupnoMesto_Change` | pamti aktivnu stanicu, osvežava predlog broja, MALINA auto-vozač | Z3, Z14 | IMA (filtriranje kooperanata po stanici od v6-ui-113) |
| `txtDatum_AfterUpdate` | pamti aktivni datum, osvežava predlog broja | Z3, Z14 | IMA |
| `cmbKooperant_Change` | puni parcele; osvežava ukupan iznos kooperanta u panelu blokova | Z6, `OtkupBlok_RefreshKoopTotal` | DELIMIČNO — parcele ima, ukupan iznos nema |
| `cmbParcela_Change`, `ExtractParcelaID` | ID parcele iz prikaza | — | IMA |
| `chkDveKlase_Click` | druga klasa | Z5 | IMA |
| `ApplyOtkupTogglesState` | parcele / keš isplate | Z6, Z7 | IMA (keš namerno izostavljen) |
| `btnUnos_Click` | **snimanje otkupa**; relink paleta hladnjače ako je bio storno | `SaveOtkupMulti_TX`, `ReassignPaleteToPrijemnica_TX`, `GetHladnjacaRelinkPending` | **IMA** (`modOtkupUnos`, v6-ui-106); legacy forma još nije prebačena na isti modul |
| `ClearOtkupFields` | reset forme posle snimanja | — | IMA (`ClearForm`); v6-ui-108: čisti i **partnera** (`cbKupac`) i **ne vraća datum na danas dok je otpremnica aktivna** — blokovi otpremnice nose njen datum i njen broj |
| `btnStornoOtkup_Click` | storno otkupa iz forme | `modStorno` | IMA (radnja nad redom) |
| `ShowLockStatus` / `HideLockStatus` | status bar + peščanik za vreme upisa | — | **NEMA** |
| `UserForm_QueryClose`, `btnPovratak_Click` | pamćenje stanice pri izlasku | `GetActiveStanica` | **NEMA** |

**Panel „Otkupni blokovi"** (`modOtkupBlok` + `clsBlokUI`) je zaseban i već je
prenet u F1: lista otpremnica, blokovi, izgubljeni, kooperanti, štampa,
specifikacija, storno, preuzimanje, prefill sa otpremnice.

---

## 3. frmDokumenta (6.500 linija) — pet režima + storno okvir

### 3.1 Režimi unosa

| Režim | Dugme | Poslovna rutina | Novi UI |
|---|---|---|---|
| Otpremnica | `btnUnosOtp_Click` | `SaveOtpremnicaMulti_TX` | F2 — **upis IMA** (`modDokUnos`, v6-ui-115) |
| Zbirna | `btnUnosZbr_Click` | `SaveZbirnaMulti_TX` + `ValidateZbirnaPreUnosa` | F3 — **upis IMA** (`modDokUnos`, v6-ui-116) |
| Prijemnica | `btnUnosPrij_Click` | `SavePrijemnicaMulti_TX`, `GetPaletaStatusForPrijemnica`, `ReassignPaleteToPrijemnica_TX` | F4 — **upis IMA** (v6-ui-116) **osim ispravke posle storna** (relink paleta) |
| OM ulaz — novac | `btnUnosOMUlaz_Click` | `SaveOMUlaz_TX`, `GetOMAvansSaldo`, `GetOpenOtkupi` | F5 — **upis IMA** (`modNovacUnos`, v6-ui-117) |
| OM ulaz — ambalaža | `btnUnosOMUlaz_Click` | `SaveOMUlaz_TX` (`koopSmer`), `OutputIzdavanjeAmbalaze` | F7 — **upis IMA** (v6-ui-117) |
| Kupci izlaz (uplate) | `btnUnosIzlaz_Click` | `SaveKupciIzlaz_TX`, `GetUplataForFaktura`, `GetOpenFakture` | F6 — **upis IMA** (v6-ui-117) |

Jedan legacy handler (`btnUnosOMUlaz_Click`) pokriva **dva** režima novog UI-ja:
novac i ambalaža su tamo mogli u isti dokument, ovde ne mogu — F5 nema polja
ambalaže, F7 nema polje iznosa (`ApplyFormFields`). Zato je i podeljen na dva.

**Šta upis F5/F6/F7 nosi, a šta namerno ne (v6-ui-117):**

| Pravilo iz legacy | Gde je sad | Napomena |
|---|---|---|
| F5: primalac + izabran blok + „iz OM avansa" biraju **tip novca** | `IsplataValidiraj` | četiri grane: `KesOtkupacKoop`, `VirmanFirmaKoop`, `VirmanAvansKoop`, `KesFirmaOtkupac`. Pogrešan tip se ne vidi u formi nego tek u saldu — zato tri sabotaže baš o njemu |
| F5: iznos ne sme preko neisplaćenog ostatka bloka | `IsplataValidiraj` | ostatak računa `GetOpenOtkupi`, ekran ga samo prosleđuje |
| F5: iz OM avansa ne više nego što ga ima | `IsplataValidiraj` → `GetOMAvansSaldo` | |
| F6: izabrana faktura → `KupciUplata` + napomena; bez nje → `KupciAvans` | `UplataValidiraj` | bez izbora fakture nijedna faktura iz novog UI-ja ne bi bila zatvorena (`UpdateFakturaStatus`) |
| F6: uplata ne sme preko preostalog iznosa fakture | `UplataValidiraj` | preostalo daje `GetOpenFakture` |
| F7: smer je **obavezan** uz količinu | `ReversValidiraj` | prazan smer je ranije tiho knjižio „OM prima od vozača" |
| F7: auto-broj reversa kad je polje prazno | `ReversValidiraj` → `SuggestNextBroj(KIND_REV, …)` | posle izbora smera, kao u legacy |
| F7: PDF revers i završetak ispravke posle upisa | `ReversUpisi` | best-effort, ne obara potvrdu upisa |
| Broj dokumenta je zajednički namespace | `DuplBroj` | duplikat se traži u **obe** tabele (`tblAmbalaza`, `tblNovac`) |
| **Vozač se za čist novac NE traži** | — | legacy ga traži uz `VALIDACIJA_UNOSA`, ali samo zbog ambalaže u istom dokumentu; `SaveNovac` ga nema, pa se odbacuje. U F7, gde ambalaža postoji, vozač je obavezan i **bez** `VALIDACIJA_UNOSA` (firma↔OM ide preko vozača) |
| **F5: partner koji je otkupno mesto JESTE entitet novca** | `IsplataValidiraj` | polje se u F5 zove „Primalac". Legacy tu mogućnost nije imao — primalac je bio samo kooperant, a otkupno mesto se podrazumevalo iz konteksta forme. Kad je partner kooperant, entitet ostaje kontekst — tačno kao legacy |
| **F7 ne prima kupca kao partnera** | `ReversValidiraj` | četiri smera idu isključivo kooperant ↔ OM ↔ firma; ambalaža kupca u legacy ide kroz prijemnicu (povrat) i kupci-izlaz, ne kroz revers |
| F5/F6: broj dokumenta je slobodan unos | Z3a | prazan je dozvoljen samo bez `VALIDACIJA_UNOSA`; tada upis vraća „(bez broja)", jer je prazan povratak rezervisan za neuspeh |
| **Ukucan a nerazrešen izbor zaustavlja dokument** | `NerazresenIzbor` | combo dopušta kucanje, a ID dolazi iz skrivene kolone koja postoji samo uz stvarno izabranu stavku. Tekst bez ID-a bi tiho promenio značenje: partner → isplata otkupnom mestu, blok → avans kooperantu, faktura → avans kupca. Sve tri se knjiže kao **ispravan** dokument, samo pogrešan |
| **Vlasništvo i trenutni ostatak proverava CORE** | `IsplataBlokProblem`, `UplataFakturaProblem` (`modNovac`) | blok mora pripadati tom kooperantu i tom otkupnom mestu, faktura tom kupcu; ostatak se čita **u trenutku upisa**, ne iz snimka koji je ekran poslao. Istu kapiju diže i `SaveOMUlaz_TX` / `SaveKupciIzlaz_TX`, pa važi i za legacy formu i za svakog drugog pozivaoca |
| **F5: lista kooperanata i blokova je sužena na aktivno OM** | `FillFormPartner`, `FillOpenBlokovi` | legacy `frmDokumenta` je taj combo sužavao **bezuslovno** (ne kroz `KOOP_FILTER_BY_OM`), pa isto važi i ovde; blokovi se filtriraju po `StanicaID` iz `GetOpenOtkupi` |

**Šta upis F3/F4 nosi, a šta namerno ne (v6-ui-116):**

| Pravilo iz legacy | Gde je sad | Napomena |
|---|---|---|
| Zbirna: vozač → kupac → broj → roba, pa kapija | `ZbirnaValidiraj` | vozač je entitet niza (Z3a), zato prvi |
| Zbirna se mora poklopiti sa svojim otpremnicama (kg **i** ambalaža) | `ZbirnaValidiraj` → `ValidateZbirnaPreUnosa` | hard-kapija, **ne zavisi** od `VALIDACIJA_UNOSA` — kao u legacy |
| Zbirna: izvor ima Kl.II a prekidač isključen → blokada | `ZbirnaValidiraj` → `ZbirnaIzvorImaKlasuII` | inače bi se Kl.II tiho izgubila |
| Zbirna **nema** bruto→neto ni cenu | — | `tblZbirna` nema ni `BrutoKg` ni `Cena`; zbirna je zbir već netiranih otpremnica |
| Zbirna: `Hladnjaca` / `Pogon` | — | novi UI nema ta polja (Z3b) → upisuje se prazno |
| Prijemnica: kupac → vozač → broj → broj zbirne → zbirna postoji | `PrijemnicaValidiraj` | ponašanje po `PRIJEMNICA_ZBIRNA_PROVERA` (BLOK / UPOZORENJE) |
| Prijemnica: bruto→neto po klasama, `BrutoKg` zamrznut | `PrijemnicaValidiraj` | isto kao otkup i otpremnica |
| Prijemnica: 1 zbirna = 1 prijemnica (pitanje, ne greška) | `PrijemnicaValidiraj` → `LookupActiveID` | |
| Prijemnica: auto-štampa + grupni otkupni list samo za hladnjaču | `PrijemnicaUpisi` | best-effort, ne obara potvrdu upisa |
| Prijemnica: status palete uz potvrdu | `PrijemnicaUpisi` → `GetPaletaStatusForPrijemnica` | |
| Zbirna: završetak ispravke posle storna | `ZbirnaUpisi` → `ZavrsiIspravkuAko` → `CompleteZbirnaIspravka` | samo nad **persistentnom** ispravkom (`tblStornoVeza`) |
| **Prijemnica: ispravka posle storna (relink paleta)** | `PrepoznajIspravkuPrijemnice` (u `PrijemnicaValidiraj`) + `PreveziPaleteIspravke` (u `PrijemnicaUpisi`) | **PRENETO u v6-ui-120.** `SetPaletizeSkip` ide **pre** upisa, pa `ReassignPaleteToPrijemnica_TX` + `PaletaAdjustPrompt` posle njega. Ispravka na čekanju traži se u `tblStornoVeza`, ne u stanju sesije: storno se pokreće u F8, unos u F4, a između to dvoje sme da se zatvori Excel. **Safe-stop:** dve ili više ispravki na čekanju → ne bira se naslepo. |

Uz njih: `LoadZbirneListbox` / `lstZbirne_Click` (izbor zbirne iz liste),
`UpdateManjak` (manjak prijemnice vs zbirna), `UpdateValidacija`,
`UpdateOMAvansSaldo`, `FillOpenFakture`, `FillOpenOtkupi`,
`SetupOMIzdavanjeToggle` (četiri smera reversa — **IMA** u F7).

### 3.2 Storno okvir — najveći deo forme

> **v6-ui-143: storno je SVOJ EKRAN, ne više režim F8.** Do tada je bio osmi
> režim unosnog ekrana, što je imalo dve posledice: crtao je unosnu formu koju
> ne koristi (`Scr_Save` za `STORNO` je padao u `Case Else` → **primarno dugme
> mrtvo**), a pregled posledica je bio niz `MsgBox`-ova. #201 je prvo sakrio
> grid-maxom; `modScrStorno` to rešava tako što forme nema, a posledice stoje u
> zoni **pre** odluke. Redovi liste i dalje dolaze iz istog čitača koji puni
> unosni ekran (`modScrDokumenti.RedoviZaTip`) — kopije nema.

Sedam panela, svaki sa svojim `Ensure/Layout/Populate/Set*Visible`:

| Panel | Šta radi | Poslovna rutina | Novi UI |
|---|---|---|---|
| **Storno** (`btnStorno_Click`) | storno bilo kog tipa dokumenta po broju | `StornoOtkup_TX`, `StornoOtpremnica_TX`, `StornoZbirna_TX`, `StornoPrijemnica_TX`, `StornoFaktura_TX`, `StornoNovac_TX`, `StornoIzvod_TX`, `StornoOMKoopByBrDok_TX` | **IMA** (`modStornoDok`, v6-ui-119; ekran `modScrStorno` od v6-ui-143) — **običan storno**, bez framework-a ispravke |
| **Ispravka / dupli unos** (`TryRunCorrectionFramework`, `PromptCorrectionMode`) | posle storna nudi ispravku (prefill) ili dupli unos | `GetCorrectionField`, `ApplySelectedBlockStorno`, `StornoSelectedBlocks_TX` | **IMA** (v6-ui-120) — sva četiri moda, od v6-ui-143 kao **četiri dugmeta sa objašnjenjem** umesto dva `MsgBox` pitanja; **bez** multiselect storna otkupnih blokova (deo legacy overlay panela) |
| **Storno pregled** (`m_btnStornoPregled_Click`) | pregled storniranih, grupisano | `GetStorniraniGrupisano` | **IMA** — čip „Otkazane" nad svakim tipom |
| **Undo operacija** (`ShowUndoOpsPanel`) | poništavanje storna | `GetUndoableStornoOperations` | **IMA** (ekran Oporavak → lista „Vrati storno", v6-ui-121) |
| **Nađi dokument** (`m_btnStornoFind_Click`) | pretraga „toplih" dokumenata po tipu i tekstu | `GetWarmStornoDocs` | **IMA** — od v6-ui-143 i **navigacioni čip „Svi"** preko tipova, uz prekidač tipa, pretragu i filtere mreže |
| **Uvid pre odluke** (`m_sc_*`, `frmDokumenta.frm:4662`) | ceo lanac, palete i blokovi PRE storna | `modStornoImpact.BuildStornoImpact` | **IMA** (v6-ui-143) — zona ekrana Storno; do tada ga je renderovao **samo** legacy. Model je identity-scoped u celosti (palete prijemnice preko `PrijemnicaID`), nosi `valid`, i **bez valjanog uvida ekran ne nudi nijednu mutaciju**. Izuzetak je zbirna: `tblPaletaStavka` nosi `BrojZbirne`, ne `ZbirnaID` — ista granica šeme kao `FLOW_DOC_ZBIRNA` u `ActiveBlocksForFlow` |
| **Nedovršeno** (`m_btnNedovrseno_Click`) | lanci koji nisu dovršeni | `GetNedovrseno`, `CancelCorrectionContext` | **IMA** (ekran Oporavak → lista „Nedovršeno", v6-ui-121). Od **v6-ui-144** lista nije samo pregled: zaostao context ispravke se odbacuje radnjom nad redom, po **CorrectionID** iz nevidljive kolone — više contexta može da deli isti poslovni broj |
| **Recovery** (`m_btnRecovery_Click`) | osirotele prijemnice i palete, prevezivanje | `GetOsirocenePrijemnice`, `GetPrijemniceSaOsirocenimPaletama`, `ReassignPaleteToPrijemnica_TX`, `ReassignPrijemnicaToZbirna_TX` | **IMA** (ekran Oporavak, četiri liste, v6-ui-121) |
| `CheckVerwaisteDokumente` | upozorenje na siročiće pri otvaranju | `GetVerwaisteDokumente` | **NAMERNO IZOSTAVLJENO** — zamenjeno stalnom listom „Nedovršeno" i brojkom u zoni; modalni dijalog pri otvaranju se zatvara i zaboravlja, lista ne može |

**Ekran „Oporavak" (v6-ui-121) — šta nosi:**

Četiri legacy panela radila su istu stvar: pokazivala šta je ostalo nedovršeno i
nudila prevezivanje. Ovde su to **šest lista** istog prekidača koji F1 i Palete
već koriste (`modScrOporavak`, registrovan u `modUiScreens.ScrRows`).

| Lista | Izvor | Radnja nad redom |
|---|---|---|
| Nedovršeno | `GetNedovrseno` | **Odbaci ispravku** → `CancelCorrectionContext` po **CorrectionID** (v6-ui-144) |
| Osirotele prijemnice | `GetOsirocenePrijemnice` | Prevezi → `ReassignPrijemnicaToZbirna_TX` |
| Zbirne (cilj) | aktivne `tblZbirna` | klik bira CILJ |
| Osirotele palete | `GetPrijemniceSaOsirocenimPaletama` | Prevezi → `ReassignPaleteToPrijemnica_TX` |
| Prijemnice (cilj) | aktivne `tblPrijemnica` | klik bira CILJ |
| Vrati storno | `GetUndoableStornoOperations` | `UndoOperation_TX` po **OperationID** |

| Pravilo | Napomena |
|---|---|
| Cilj se bira klikom na red i stoji u zoni gore | isti obrazac kao aktivna otpremnica u F1 i aktivna paleta na ekranu Palete; legacy je za to imao combo u panelu — ovde je lista, pa se cilj može i pretražiti i sortirati |
| Liste ciljeva nude **samo aktivne** dokumente, jedan red po broju | prevezivanje na storniran cilj bi napravilo drugu siroticu umesto da reši prvu; klase I i II dele broj, a cilj JESTE broj |
| Liste ciljeva nude **samo aktivne** dokumente | prevezivanje na storniran cilj bi napravilo drugu siroticu umesto da reši prvu |
| Jedan red po **dokumentu** (broj + vlasnik), ne po broju | klase I i II dele broj **i vlasnika** → jedan dokument, jedan red. Dva kupca sa istim brojem → **dva** dokumenta, dva reda: `BrojPrijemnice` se računa po kupcu, pa je kolizija svakodnevna. Kolona VLASNIK je zato vidljiva. |
| Liste nose **GeneracijaID** i prosleđuju ga u akciju — i izvorne i **ciljne** | broj je labela, identitet je generacija; `Reassign*_TX` po njoj bira redove, pa dokument koji deli broj ne može biti zahvaćen |
| Cilj se bira po identitetu, ne po broju (`newGeneracijaID`, `zbirnaGeneracijaID`) | `BrojPrijemnice` se generiše po kupcu: kod kolizije je `newById(klasa)` uzimao red koji je slučajno poslednji u tabeli, pa je roba mogla da ode tuđem kupcu (v6-ui-124) |
| Labela se čita iz izabranog dokumenta, ne od pozivaoca | neusklađen par (broj jednog, generacija drugog) inače tiho upisuje tuđi broj |
| Propagacija u `tblPaletaStavka` ide po `PrijemnicaID` | prvi upis je bio po identitetu a drugi po broju, pa je tuđi dokument ostajao sam sebi protivrečan — prijemnica na staroj zbirni, njena paleta na novoj (v6-ui-125) |
| Zadata generacija koje nema → **STOP**, ne fallback po broju | prazan argument (legacy zapis) i „baš taj dokument, a nema ga” su dva različita stanja |
| Ciljna lista zbirnih grupiše po **generaciji**, vlasnik je vozač + kupac | sa samim kupcem bi dva dokumenta istog broja pala u jedan red i operater ne bi mogao da izabere pravi |
| „Jedini vlasnik" zbirne se meri **dokumentima**, ne distinct brojevima | zbirna je zbir svih svojih otpremnica, a broj otpremnice je scoped po stanici — dve otpremnice istog broja sa različitih stanica davale su jedan distinct broj, pa je PONIŠTENJE ulazilo u punu kaskadu i obaralo tuđu |
| Deca zbirne **nisu** nerešiva, samo još nisu scoped | otpremnica kaskada već ume `BrojZbirne + VozacID`, prijemnica `+ KupacID`, palete nose `PrijemnicaID`. Fail-closed je bezbedan izbor **dok se child mutacije ne dovedu dotle**, ne dokaz nemogućnosti |
| **Broj zbirne je jedinstven, broj prijemnice nije** | `SuggestNextBroj` za `ZBR` bumpuje sekvencu dok `BrojZbirneExists` ne kaže da je slobodan; `GenerateBrojPrijemnice` ima fiksan prefiks `1`, broji po kupcu i **nema takvu proveru**. Kod zbirne je identitet pojas za ručni unos, kod prijemnice je nužnost |
| Presuda o relabelu ide nad **već razrešenim** dokumentima (`PresudiPaletaReassign`) | writer je birao po generaciji, a `EvaluatePaletaReassign` ga je ponovo tražila po broju — kod kolizije je presuda opisivala tuđi dokument i relabel se tiho preskakao (v6-ui-126) |
| „Isti dokument” u ekranu se meri **generacijom**, ne brojem | ispravka koja menja kupca dobija isti poslovni broj kao original — poređenje po broju je odbijalo potpuno ispravnu operaciju |
| Su-stanar na deljenoj paleti je **drugi dokument**, ne drugi broj | dva kupca istog broja i iste robe smeju da dele paletu; poređenje po broju ih je videlo kao istu prijemnicu, pa bi relabel prepravio header cele palete a tuđa roba ostala pogrešno označena (v6-ui-127) |
| Ista kapija „isti dokument” stoji i u **writeru**, ne samo u ekranu | popravka samo u UI-ju je pravilo preselila u core umesto da ga zatvori |
| Bez generacije (stari zapisi) → **fail-closed** | `RequireJedanVlasnikPoBroju` / `VlasniciPoBroju`, sa kompozitnim vlasništvom po tipu — prijemnica kupac, zbirna **vozač + kupac** (isti par koji koriste `StornoZbirna` i `ApplyGeneracijaID`) |
| „Vrati storno" cilja **OperationID**, ne poslednju operaciju po broju | isti broj dokumenta može imati više generacija; zato je prva kolona baš `OperationID` |
| Kapija je `UndoGuardReason` (fail-closed) | ista koju diže i legacy dugme |
| Prevezivanje paleta ide sa `force=True` | ovaj ekran postoji baš da razreši ono što automatika nije umela; razlika u broju gajbica se prijavljuje i koriguje u mestu (`PaletaAdjustPrompt`), ne blokira |

**Šta F8 nosi, a šta namerno ne (v6-ui-119):**

F8 više nije pogled na jednu tabelu. Prekidač bira **tip dokumenta**, a sve što
pita „koja tabela" i „koje kolone" razrešava `modScrDokumenti.EffKey` — do tada
je `"STORNO"` bio tih sinonim za `"OTPREMNICA"` u desetak `Col*` funkcija.

| Pravilo iz legacy | Gde je sad | Napomena |
|---|---|---|
| Tip dokumenta bira šta se stornira | prekidač lista (`Scr_Liste` za F8) | devet tipova kroz sedam tabela; legacy `cmbStornoDokument` ima jedanaest stavki, jer četiri smera reversa broji zasebno |
| Otkup i otpremnica: klase I i II dele broj → stornira se **ceo** dokument | `StornoIzvrsi` → `*ByBrDok_TX` / `*ByBroj_TX` | isto pravilo koje F1 lista već poštuje |
| Novac: izvod se ne stornira parcijalno; broj sa više aktivnih redova traži `NovacID` | `ResolveNovacForStorno` (`modStorno`) | avans raspodela deli broj originalne stavke — **posledica:** takav red se iz F8 **ne može** stornirati, isto kao iz legacy forme bez ukucanog `NovacID` |
| Izvod: „broj" ili „broj/račun" → jedan izvod | `ResolveIzvodZaStorno` | broj računa se čita iz **treće kolone reda**, ne iz mape — dva izvoda istog broja tako ostaju razlučiva |
| Izvod: preflight blokada pre potvrde | `GetIzvodStornoBlokade` | razlog se vidi pre „Da", ne kao tih neuspeh posle |
| Izvod: ishod REMAP vs REIMPORT | ekran (`StornoRedF8`) | to je **odluka operatera o PDF-u**, ne pravilo — zato je u ekranu, a ne u `modStornoDok` |
| Revers: broj + **smer** | `StornoRazlog` → `ActiveAmbalazaDokExists` | četiri smera dele `KIND_REV`, pa broj sam ne kaže koji je red u `tblAmbalaza`; ekran smer nalazi tako što pita koji od četiri ima aktivan red |
| Zbirna: upozorenje da aktivna prijemnica ostaje vezana | `StornoIzvrsi` | `StornoZbirna` namerno ne kaskadira na prijemnice; bez poruke operater ne zna da mu je ostao posao |
| **Storno palete i prerade** | **NIJE preneto** | `StornoPaleta_TX` / `StornoPrerada_TX` pripadaju ekranu Palete (F8 nema tip „paleta"); tamo su i danas, kroz `modScrPalete` |

**Šta nosi framework ispravke (v6-ui-120):**

Za četiri tipa sa nizvodnim tokom storno nije Da/Ne nego izbor **šta storno
poslovno znači**. Smart trigger je isti kao u legacy: pun izbor se nudi samo kad
`CorrectionNeedsDialog` kaže da ima o čemu da se odlučuje.

| Pravilo iz legacy | Gde je sad | Napomena |
|---|---|---|
| Četiri moda (ISPRAVKA / DUPLI / PONIŠTENJE / REŠI KASNIJE) | `modScrStorno.AkcijeZaTip` → `modStornoDok.StornoIzvrsiMod` → `modStornoFlow.Run*Correction` | **odluka izmenjena u v6-ui-143.** Do tada: dva `MsgBox` pitanja, uz obrazloženje „četiri odgovora ne staju u jedan `MsgBox`" — tačno, ali je zaključak bio pogrešan. Sada su **četiri dugmeta u zoni**, svako sa objašnjenjem ispod, a iznad njih stoje posledice (lanac, palete, blokovi). Operater vidi sva četiri ishoda **istovremeno**, umesto da drugi izbor otkrije tek pošto odgovori na prvi. Prekidač „Ne diraj palete" je iz istog razloga izašao iz `MsgBox`-a (`STORNO_ASK_PALETE`) i stoji uz palete na koje se odnosi |
| Framework važi **samo** za otpremnicu, zbirnu, prijemnicu i revers | `TipUFlowDoc` | ostalih pet tipova nema nizvodni tok — njihov storno je običan, kao i u legacy |
| Revers: kratko pitanje storno vs ispravka | `IspravkaPreuzela` | revers je list u lancu, pa nikad ne traži pun izbor (legacy `RunReversStornoUI`) |
| PONIŠTENJE se izvršava u **dva** poziva | `IzvrsiMod` | prvi vrati `blocked=True` i pun spisak posledica; drugi ide tek po svesnoj potvrdi, sa `forceConfirm` — spisak se tako pravi PRE nego što se išta promeni |
| „Ne diraj palete" (prijemnica, DUPLI/PONIŠTENJE) | `IzvrsiMod` | uz ISPRAVKU se ne pita: tamo se palete prevezuju na novi dokument |
| ISPRAVKA → prefill + prelazak u režim unosa | `OtvoriIspravku` | režim se menja **pre** prefilla — `SelectMode` čisti formu, pa bi obrnut redosled obrisao upravo prepisane vrednosti |
| Storno otkupnih blokova uz DUPLI/PONIŠTENJE | `StornirajBlokoveAko` (v6-ui-121) | **multiselect od `v6-ui-149`** (lista „Blokovi“, podrazumevano nijedan — kao legacy). Do tada **sve-ili-ništa**: pre pitanja se ispiše pun spisak (broj, klasa, kg, kooperant), pa operater vidi nad čim odlučuje. Delimičan izbor ostaje na ekranu Oporavak, gde izgubljeni blokovi imaju svoju listu i radnju po redu. Kapija `BlockStornoDriftReason` (ADR-0001) je ista. |

### 3.3 Ostalo

`RefreshTopKpis` + `SumOtkupKgToday` (KPI traka, Z12), `SetupFkeyAccelerators`
/ `HandleFkey` (Z11 — **IMA**), ~60 `_Enter`/`_Exit`/`_KeyDown` handlera po
polju (Z11 — **IMA**, rešeno u `clsFlatBtn` jednom za sve).

---

## 4. Šta novi UI već ima, a legacy nema

Da se ne izgubi iz vida — prelazak nije samo prepisivanje:

- jedna forma umesto devetnaest, sve kontrole runtime (`.frx` se ne dira);
- mreža sa sortiranjem, pretragom, stranama i označavanjem redova nad **svakom**
  listom (legacy listbox-i to nemaju);
- prekidač lista po ekranu (`Scr_Liste`) i radnje nad redom (`Scr_Radnje`) bez
  ijednog imena liste u ljusci;
- „Izgubljeni" i „Kooperanti" kao obične liste umesto skrivenih režima;
- specifikacija po filtriranoj listi, ne po zasebnom izveštaju;
- prefill sa otpremnice pri izboru (F1), sa fokusom na kooperanta.

---

## 5. Plan — redosled

Rangirano po tome koliko svaka stavka blokuje **stvarni rad**, ne po veličini.

### Faza A — pravila unosa (bez upisa, mala i vidljiva)
1. ~~**Z1 + Z2**: cena iz cenovnika i tip ambalaže iz kulture na promenu
   vrste/sorte.~~ **URAĐENO** (v6-ui-102)
2. ~~**Z4**: živi zbir kg / neto iz bruta.~~ **URAĐENO** (v6-ui-102)
3. ~~**Z9**: „još N gajbica do zatvaranja palete".~~ **URAĐENO** (v6-ui-104)
4. ~~**Z3 + Z14**: predlog broja i kontekst stanice/datuma.~~ **URAĐENO** (v6-ui-104)
5. ~~**Z8 + Z13**: toggle-i.~~ **URAĐENO** (v6-ui-103, v6-ui-104)
   Z7 otpada — keš isplate ne idu kroz otkupni list (v6-ui-105).

**Faza A je time zatvorena.** Ostaje iz nje samo `ShowLockStatus` (peščanik za
vreme upisa); filtriranje kooperanata po otkupnom mestu je urađeno u v6-ui-113
(`KOOP_FILTER_BY_OM`).

### Faza B — upis (`CommitDokument`)
6. ~~F1 → `SaveOtkupMulti_TX` (+ relink paleta hladnjače).~~ **URAĐENO**
   (v6-ui-106, `modOtkupUnos`).
7. ~~F2 → `SaveOtpremnicaMulti_TX` (+ auto-zbirna MALINA, ispravka).~~
   **URAĐENO** (v6-ui-115, `modDokUnos`).
8. ~~F3/F4 → `SaveZbirnaMulti_TX` (+ `ValidateZbirnaPreUnosa`),
   `SavePrijemnicaMulti_TX` (+ status palete).~~ **URAĐENO** (v6-ui-116,
   `modDokUnos`). Ostaje iz te stavke: **živi prikaz manjka** prijemnice vs
   zbirna (`UpdateManjak`) i **lista zbirnih za izbor**, oboje prikaz a ne upis;
   i **ispravka prijemnice posle storna** (relink paleta), koja pripada Fazi D.
9. ~~F5/F6/F7 → `SaveOMUlaz_TX`, `SaveKupciIzlaz_TX`, novac.~~ **URAĐENO**
   (v6-ui-117, `modNovacUnos`). Uz upis su došla i dva polja bez kojih upis ne
   bi bio tačan: prekidač „ISPLATA IZ" sa avans saldom OM (F5) i lista
   otvorenih faktura (F6) — v. §3.1.
10. Z12: preostali KPI-jevi (`GetOMAvansSaldo`, otvoreno kg).

**Legacy se NE gasi i NE menja.** `frmOtkup` i `frmDokumenta` ostaju potpuno
operativni dok novi UI ne bude umeo sve što one umeju; do tada obe kopije
poslovne logike postoje **namerno**. Pravilo za taj period: pravilo unosa se
menja u `modOtkupUnos` / `modDokUnos`, pa se **ručno preslika** u legacy formu,
i to se zapiše uz izmenu. Prebacivanje legacy formi na zajedničke module je
odluka koja dolazi tek kad novi UI prođe rad u pogonu (i moguće nikad, ako se
legacy do tada penzioniše).

### Faza C — Palete P2
10. ~~Unos prerade (traži prosleđivanje događaja sopstvenih kontrola ekranu).~~
    **URAĐENO** (`v6-ui-159`). Četvrta lista ekrana Palete „Nova prerada":
    stikliranje paleta u mreži + polja u zoni. Ljuska je dobila granu za promenu
    u polju ugovornog ekrana (`chg:`), simetričnu sa `act:` i `row:`.
11. ~~Čipovi po ekranu (godina/status/prerađeno) — generalizacija čipova.~~
    **URAĐENO** (`v6-ui-169`). Ekran prijavljuje svoje čipove kroz `Scr_Cipovi`
    (`kljuc:KATALOG:sirina`), ljuska pozajmljuje slotove svog bazena i vraća
    ključ kroz `Scr_Rows`. Time je nestalo i poslednje mesto na kom je ljuska
    znala jedan ekran po imenu (`akt = "OTPREMNICE"`). Palete su dobile pet
    čipova: Sve · Ova godina · Otvorene · Zatvorene · Prerađene.

**Uz Fazu C je došlo i:** dvoklik na red se prosleđuje ekranu (`dbl:<red>`), pa
klik na paletu vodi pravo na njene stavke; panel za unos prerade pokazuje i **neto
ulaz** izabranih paleta pored neto izlaza (`v6-ui-169`). Traženi **padajući redovi
detalja ispod izabranog reda** su odloženi sa razlogom: mrežu koriste svi ekrani, a
za detalj ispod reda ugovor bi morao da dobije **vrstu reda** i opis kolona po
vrsti, uz renderer promenljive visine — posao reda veličine migracije jednog celog
ekrana. Stoji kao prioritet za kasnije.

**Faza C je time ZATVORENA** (uz Compile i smoke kao poslednju kapiju).

**Uz Fazu C je došlo i:** dvoklik na red se prosleđuje ekranu (`dbl:<red>`), pa
klik na paletu vodi pravo na njene stavke (`v6-ui-162`). Traženi **padajući redovi
detalja ispod izabranog reda** su odloženi sa razlogom: mrežu koriste svi ekrani, a
za detalj ispod reda ugovor bi morao da dobije **vrstu reda** i opis kolona po
vrsti, uz renderer promenljive visine — posao reda veličine migracije jednog
celog ekrana. Stoji kao prioritet za kasnije, jer znači na više mesta.

### Faza D — storno okvir (najveći, ide poslednji)
12. ~~Storno svih tipova dokumenata iz F8.~~ **URAĐENO** (v6-ui-119,
    `modStornoDok` + prekidač tipa u F8). Prenet je **običan** storno — onaj
    koji legacy radi kad `TryRunCorrectionFramework` ne preuzme tip. Uz njega su
    došli i „Storno pregled" (čip „Otkazane" po tipu) i „Nađi dokument"
    (prekidač tipa + pretraga mreže), koji su bili zasebni paneli.
13. ~~Ispravka / dupli unos posle storna (Z10) + „hladnjača ispravka" +
    ispravka prijemnice (relink paleta).~~ **URAĐENO** (v6-ui-120). Sva
    četiri moda (`modStornoFlow`) idu iz F8; prefill iz storniranog je jedan
    račun (`modStornoDok.PrefillIzStorniranog`) umesto četiri kopije vezane
    za kontrole; ispravka prijemnice preskače svežu paletizaciju i prevezuje
    palete; hladnjača ispravka se nudi posle storna otkupa iz F1.
14. ~~Recovery, Nedovršeno, Undo operacija (+ storno blokova iz stavke 13).~~
    **URAĐENO** (v6-ui-121). Nov ekran **Oporavak** (`modScrOporavak`) sa šest
    lista zamenjuje četiri legacy panela; storno otkupnih blokova uz
    DUPLI/PONIŠTENJE je vezan u F8, sve-ili-ništa uz pun spisak.

**Faza D je time ZATVORENA.**

### Faza E — ostali ekrani
15. ~~Agrohemija~~ **URAĐENO** (v6-ui-171, dorada `v6-ui-172`, `modScrAgro`) — v. §7.
16. ~~Fakture~~ **URAĐENO** (`v6-ui-176`, `modScrFakture`) — v. §8.
17. ~~Banka uvoz~~ **URAĐENO** (`v6-ui-177`, `modScrBankaUvoz`) — v. §9.
    Ostaju: Banka nalozi, Marža, Izveštaji, Sledljivost — svaki po istom
    obrascu.

---

## 6. Pravilo koje važi za sve faze

Ekran **nikad** ne računa i ne upisuje sam. Svaka stavka iz plana se rešava
pozivom postojeće rutine (`modCenovnik`, `modBrojevi`, `modOtkup`,
`modDokumenta`, `modPaletniList`, `modStorno`); ako rutina postoji ali je
`Private` i vezana za formu, prvo se **izdvaja račun** iz prikaza (kao
`KoopRangRows` iz `LoadKoopRang`), pa je koriste i legacy forma i novi ekran.
Duplirana logika se ne piše ni u jednom slučaju.

---

## 7. Agrohemija — šta je preneto (v6-ui-171, dorada `v6-ui-172`)

Prvi ekran **Faze E**. Zona nosi celu unosnu formu — što je moguće tek od
`v6-ui-159`, kad je Faza C/10 (unos prerade na Paletama) otvorila polja i
promenu teksta ugovornim ekranima.

### 7.1 Gde je šta završilo

| Legacy (`frmAgrohemija`) | Novo mesto |
|---|---|
| korpa izlaza / ulaza, `tKorpaItem` | `modAgroUnos` — korpa je `Collection` rečnika |
| `BuildArtikalStanjeDict` | `modAgroUnos.AgroStanjeMapa` |
| `GetKorpaIzlazKolicinaZaArtikal` | `modAgroUnos.AgroKorpaKolicina` |
| `ValidateKorpaIzlazStanje` | `modAgroUnos.AgroProveriKorpuIzlaz` |
| `btnDodajIzlaz_Click` provere | `modAgroUnos.AgroDodajIzlaz` |
| `btnDodajUlaz_Click` provere (+ potvrda cene 0) | `modAgroUnos.AgroDodajUlaz` |
| `btnZavrsiIzlaz_Click` transakcija | `modAgroUnos.AgroUpisiIzlaz` |
| `btnZavrsiUlaz_Click` transakcija | `modAgroUnos.AgroUpisiUlaz` |
| `UpdatePreporuka` (smart doza → pakovanja) | `modAgroUnos.AgroPreporukaInfo` |
| tri kopije invarijante nad `Pakovanje` | `modAgroUnos.AgroArtikalInfo` (jedna) |
| `m_btnPocetniDug_Click` | `modScrAgro.PocetniDug` → `BookPocetniDug` (nepromenjen) |
| KPI traka (4 broja) | zona ekrana, ista četiri broja |

Nove rutine za mrežu (ekran ne čita tabele sam):
`modAgrohemija.GetMagacinPrometForGrid`, `modAgrohemija.GetAgroDugoviForGrid`,
`modNovac.GetAgroAbzugMapa` (jednoprolazna mapa umesto `GetAgroAbzug` u petlji).

### 7.2 Šta ekran uzima od ljuske

Ništa od ovoga nije napravljeno za agrohemiju — sve je došlo uz **Fazu C** i
ovde se samo koristi. To je i bila poenta: drugi korisnik istog ugovora ne sme
da izmišlja svoju varijantu.

| Potreba | Ljuskin ugovor | Otkad |
|---|---|---|
| polje (natpis + okvir + kontrola) | `modOtkupUI.NewFieldG` | `v6-ui-159` |
| raspored unutar polja | `modOtkupUI.LayoutFieldInner` | `v6-ui-159` |
| promena teksta stiže ekranu | `Scr_Event("chg:<kontrola>")` | `v6-ui-159` |
| padajuća lista nad poljem u zoni | `FindCombo` gleda i `zScr_<ekran>` | `v6-ui-159` |
| klik na kontrolu u zoni | `Scr_Event("<tag>")`, prefiks `scr` | `v6-ui-143` |

Zbog toga kombo u zoni **mora** biti polje (okvir `nm` + kontrola `nmT`), a ne
gola kontrola — panel za izbor traži baš taj oblik.

Ekran prijavljuje i **čipove**, **brojač** i **dvoklik** — sve tri kuke koje je
ugovor dobio uz Fazu C:

| Lista | Čipovi (`Scr_Cipovi`) |
|---|---|
| Korpa | — (nekoliko upravo unetih redova; tu se ne traži nego se gleda) |
| Stanje | Sve · Ima na stanju · Bez zaliha |
| Promet | Sve · Ulazi · Izlazi · Ova godina |
| Dugovi | Sve · Duguju |

`Scr_Brojac` broji **korpu** — jedino što na ovom ekranu čeka operatera; sve
ostalo je već u tabelama. Bez te brojke neproknjižena korpa nestane bez traga
čim se pređe na drugi ekran.

`dbl:<red>` preuzima red u unos: iz **Dugova** kooperanta (i prebacuje u
IZDAVANJE — dug se izdaje, ne prima), iz **Stanja** artikal. Jedan potez umesto
tri (zapamti ime → pređi na korpu → nađi ga u padajućoj listi).

**Dvoklik bira po identitetu, ne po tekstu reda.** Lista pokazuje ime, a bira se
kooperant; dva kooperanta istog imena daju dvosmislen prikaz i tada dvoklik
**odbija** da bira umesto da pogodi — isto pravilo kao „dvosmislen broj →
MANUAL" u storno okviru. Pogađanje bi ovde izdalo robu pogrešnom čoveku.

### 7.3 Šta je namerno drugačije od legacy-ja

- **Dve sekcije → prekidač režima u zoni** (IZDAVANJE / PRIJEM). Ljuska ima
  jednu mrežu i jednu zonu; dve forme jedna pored druge se ne uklapaju.
  Obe korpe žive istovremeno — prelazak režima ne prazni ništa.
- **Multiselect parcela → sakupljanje dugmetom „+ Parcela".** Mreža bira jedan
  red, combo jednu stavku; zbir ha (koji smart doza računa) drži ekran uz
  spisak. Rezultat je isti `parcelaID` niz razdvojen `;` i isti zbir ha.
- **Četiri liste u mreži** kojih legacy nema: korpa, stanje magacina, promet i
  dug po kooperantu. Mreža je već tu — legacy je za isto morao u Izveštaje.
- **„Ukloni stavku" i „Isprazni korpu".** Legacy pogrešnu stavku nije umeo da
  izbaci — jedini izlaz je bio zatvaranje forme. Ekran ostaje otvoren, pa bez
  toga ne bi bio upotrebljiv.

### 7.4 Šta NIJE preneto

- `frmAgrohemija` se **ne gasi i ne menja** — isto pravilo kao za `frmOtkup` i
  `frmDokumenta` (§5, Faza B). Legacy zadržava svoju kopiju logike; pravilo se
  menja u `modAgroUnos` pa se **ručno preslikava** u formu.
- **Dobavljač je slobodan tekst**, kao i u legacy-ju (`cmbDobavljac` se nigde
  ne puni iz tabele). Šifarnik dobavljača ne postoji i ovde se ne uvodi.
- **Storno magacin stavke** nije ovde — to je posao ekrana Storno.
- KI-006 (`ExportMagacinKoop` ne izuzima `ART_POCETNI_DUG`) je **netaknut**:
  PWA izvoz nije deo ovog prelaska.

### 7.5 Verifikacija

Testovi 82–90 u `modTest`, uz nove fixture redove `tblArtikli` / `tblMagacin`
u `tools/make_fixture.py`. Fixture je namešten tako da zaokruženje **nagore**
ima gde da padne: doza 2 l/ha, pakovanje 5 l, stanje 15 l.

| Test | Šta meri | Sabotaža |
|---|---|---|
| `T_Agro_UgovorEkrana` | registar, četiri liste, radnja samo nad korpom | `agro-modul-ime` |
| `T_Agro_KapijaStanjaBrojiKorpu` | kapija broji korpu; kapija pred upis agregira po artiklu | `agro-korpa-se-ne-broji`, `agro-agregat-po-redu` |
| `T_Agro_SmartDozaZaokruzujeNagore` | doza → cela pakovanja, nagore | `agro-doza-nanize` |
| `T_ZonaAgro_PoljaPostojeIPrateRezim` | sve kontrole zone postoje; prekidač režima pali i **gasi** prava polja | `agro-rezim-ne-gasi-polja` |
| `T_Agro_CipoviSuzavajuListu` | ugovor čipa (`kljuc:KATALOG:sirina`) i pravilo svakog | `agro-cip-ne-suzava` |
| `T_Agro_BrojacIDvoklikPoIdentitetu` | brojač vidi korpu; dvosmislen prikaz nosi **prazan** identitet | `agro-brojac-ne-vidi-korpu`, `agro-dvosmislen-prvi-pobedjuje` |
| `T_Agro_AbzugMapaPratiPojedinacni` | mapa odbitaka i pojedinačni račun daju **isto**, nad svim kooperantima | `agro-abzug-mapa-ne-sabira` |
| `T_ZonaAgro_PrekidacRezimaZadrzavaBoju` | izabran režim ostaje zelen i kad pokazivač ode | `agro-prekidac-bez-rebase` |
| `T_Agro_TrakaKorpe_NajnovijePrvoIPreliv` | traka korpe: najnovije prvo, preliv se **prijavljuje** | `agro-traka-najstarije-prvo`, `agro-traka-bez-preliva` |

Granice bazena ljuske (`MaxPrekidaca`, `MAX_ACT`, `MAX_CHIP`, kolone) tvrdi
`T_Agro_UgovorEkrana`, sa sabotažom `agro-cipova-preko-bazena`. Višak se inače
**tiho odseca** — operater vidi ekran kome fali dugme, bez ijedne poruke.

### 7.6 Otvoreno

- **Suite je puštena i zelena.** Prvo izvršavanje je bilo na rebase-u na
  `main` (`v6-ui-171`), na mašini sa Excelom. `RunAllTests` **ZELENO (88)**,
  pun set **ZELENO** (72 · 189 · 35 · 181 · 97 · 336 · 25), svih **devet**
  agro sabotaža (sada ih je **jedanaest**) obara **imenovani** test i uredno se vraća.
- **Prvo puštanje je oborilo dva testa** — oba pisana nad fixture-om kakav
  nije, produkcioni kod je bio ispravan:
  - `T_Agro_KapijaStanjaBrojiKorpu` je kontrolni izlaz upisivao sa **praznom
    parcelom** (šesti pozicioni argument `SaveMagacinCore`), a `PRACENJE_PARCELA`
    je u fixture-u ON → 4215. Test je padao na svom **čistaču**, ne na kapiji.
  - `T_Agro_BrojacIDvoklikPoIdentitetu` je tražio identitet kooperanta
    `KOOP-TEST-2`, **koga u listi dugova nije bilo** (mapu puni čitač liste, a
    lista se gradi samo iz `MAG_IZLAZ` redova). Tvrdnja je merila **odsustvo
    reda**. Fixture zato dobija `MAG-TEST-4`, dug preko rezervisanog
    `ART_POCETNI_DUG` — da stanje `ART-TEST-1` ostane tačno 15.

  To je i cena pisanja testa bez izvršavanja: obe greške bi pale na prvom
  puštanju, a nijedna se ne vidi čitanjem.
- **Compile** (`Alt+F11 → Debug → Compile VBAProject`) ostaje operateru.

### 7.7 Prvi smoke: prekidač režima je belio

Smoke je našao kvar koji suite tada nije mogao da vidi: izabran režim je bio
zelen **samo dok je pokazivač nad njim**; čim pređe dalje, ispuna se vrati na
belu, a natpis ostane krem — pa aktivno dugme postane skoro nečitljivo.

Uzrok nije bojenje nego **pamćenje**. `clsFlatBtn` zapamti osnovnu boju pri
`Bind`-u i vraća je u `ResetVisual` kad pokazivač ode. `BoxState` menja kontrolu,
ali ne i tu zapamćenu osnovu. Dva su leka, i trebala su oba:

1. **`RebaseSink`** posle svakog `BoxState` — render koji promeni boju javlja novu
   osnovu. Isti kvar i ista popravka kao `StilDugmeta` u `modScrStorno`, gde se
   videlo samo na jednom od četiri dugmeta jer su ostala tri ionako tamna na belom.
2. **Vrsta `"seg"`** umesto `"btn"` — prekidač režima *jeste* segmentni prekidač,
   isti kao onaj nad mrežom, pa se i pravi istom fabrikom (`NewSegBtn`).
   `clsFlatBtn.IsSelected` priznaje izabrano stanje (`Font.Bold`) samo za
   `"nav"`, `"chip"` i `"seg"`. Kao `"btn"` je izabran režim bio obično dugme.

Bojenje prekidača time seli iz osvežavanja zone u `RasporediPolja`, uz vidljivost
polja: koji je režim izabran je **jedna** odluka, pa boja i raspored ne mogu da
se raziđu — i `Scr_Layout` dobija zonu argumentom, pa se može izmeriti u testu.

Čipovi (`ChipV`, vrsta `"chip"`) i prekidač lista (`NewSegBtn`) su bili zaštićeni
od početka — zato su na istom ekranu radili ispravno. Ostali ugovorni ekrani
nemaju ovaj kvar: `modScrPalete`, `modScrDokumenti` i `modScrOporavak` uopšte ne
prefarbavaju kontrole, a `modScrStorno` pokriva sva svoja mesta.

Test `T_ZonaAgro_PrekidacRezimaZadrzavaBoju` reprodukuje kvar **bez miša**:
`ResetVisual` se zove direktno nad sink-om, što je tačno ono što se desi kad
pokazivač napusti dugme, pa se boja čita pre i posle — u **oba** režima, da
sabotaža koja zamrzne boje na prvoj vrednosti ne prođe.

> Deo popravke se ne može izmeriti headless: prelazak na vrstu `"seg"` menja
> **hover-in** ponašanje (izabrano dugme se više ne zatamnjuje pod pokazivačem,
> kao ni prekidač lista ispod njega). To ostaje na smoke listi, uz ponovno
> puštanje suite — test 89 i sabotaža `agro-prekidac-bez-rebase` **nisu
> izvršeni**, pisani su u sesiji bez Excela.

### 7.8 Drugi smoke: korpa se nije videla

Korpa se videla **samo dok je izabrana lista „Korpa"**. Operater koji gleda
stanje, promet ili dugove nema nijedan znak šta je upravo dodao — a desna
polovina reda polja svejedno stoji prazna (parcele su često isključene, pa
slotovi 3 i 4 nikad ništa ne nose).

Zona je zato dobila **traku korpe** uz desnu ivicu: naslov, poslednje stavke i
zbir. Polja uzimaju **ostatak** širine — isti raspored kao `PRE_DESNO` na ekranu
Palete. Na uskom ekranu traka nestaje i polja uzimaju celu zonu, isto pravilo
kao KPI brojke koje bi nalegle na prekidač režima.

Dva pravila, oba se tiho kvare:

- **Najnovije prvo.** Operater upravo nešto doda, pa mu je potvrda ono što traži.
  Obrnut redosled izgleda ispravno dok se korpa ne napuni preko četiri reda.
- **Preliv se prijavljuje** (`… još N`). Lista koja se tiho odseca izgleda kao
  cela — isto pravilo koje ljuska nad sobom već ima (`BazenStaje`).

Račun (`TrakaRed`) je odvojen od crtanja, pa se meri bez forme. Fixture je dobio
`ART-TEST-Z` sa velikom zalihom i pakovanjem od 1: preko `ART-TEST-1` se preliv
ne može izmeriti jer mu kapija stanja (15 kg, pakovanje 5) propušta najviše tri
pakovanja.
- **Dupla implementacija odbitka je ZATVORENA.** `GetAgroAbzugMapa` ostaje
  brza kopija pravila iz `GetAgroAbzug` — obe su žive u istoj funkciji (mapu
  zove lista dugova, pojedinačnu keš ekrana), pa se mogu razići. Fixture je
  dobio pet `AgroAbzug` redova (dva za istog kooperanta, jedan storniran, jedan
  drugog tipa), a `T_Agro_AbzugMapaPratiPojedinacni` tvrdi slaganje nad **svakim**
  kooperantom koga mapa zna — i tačne zbirove, da ih ne obori isti kvar na obe
  strane. Sabotaža `agro-abzug-mapa-ne-sabira` (500 → 200).

### 7.9 Review PR #213: identitet stavke i kanal za značku (`v6-ui-172`)

Dva nalaza iz review-a merged PR-a #213. Oba su ista klasa greške — **stanje koje
se čita iz onoga što se vidi, umesto iz onoga što jeste**.

#### P1 — „Ukloni stavku" je birao po prikazu

Stavka korpe je tražena po **nazivu artikla i količini** iz prikazanog reda.
Dve iste stavke su tada nerazlučive, a to nije izmišljen slučaj: *„dva pakovanja
sada, dva kasnije"* daje dva reda iste robe i iste količine. Klik na drugi red je
izbacivao **prvi** — tiho, jer red koji nestane izgleda isto kao onaj koji je
trebalo da nestane.

Isto pravilo kao „dvosmislen broj → MANUAL" u storno okviru, samo što se ovde
dvosmislenost može **sprečiti** umesto prijaviti:

- `modAgroUnos.NoviRed` svakoj stavci daje `stavkaID` (`NovaStavkaId`). ID je
  **prolazan** — živi koliko i korpa, nikad ne ide u tabelu. Brojač je
  modul-level jer dve korpe (izdavanje i prijem) žive istovremeno.
- Grid korpe dobija **osmu kolonu** sa tim ID-jem, prioriteta **4**. Mreža crta
  do prioriteta 3 (`LayoutGrid`), pa vrednost postoji u modelu a ćelija se nikad
  ne pravi — `GridCell(red, 8)` je čita, operater je ne vidi. **Ljuska se ovim ne
  menja**: mehanizam je već postojao.
- Identitet ide **u red**, ne pored njega: mreža redove sortira i deli na strane,
  pa bi svaka mapa „prikaz → stavka" koju ekran drži sa strane zastarela na prvi
  klik po zaglavlju.
- `UkloniPoId` zamenjuje `UkloniPoPrikazu`; uklanjanje je u
  `modAgroUnos.AgroUkloniStavku`, jer je korpa struktura tog modula. Prazan ili
  nepoznat ID **ne uklanja ništa** i javlja se porukom — ne pogađa se.

#### P2 — značka je ćutala dok korpa nije prikazana lista

Ljuska brojače uz stavke menija pita **samo** kroz `RefreshFromData`, a nju zove
tek kad ekran na klik javi `True` = „podaci su promenjeni". Ekran to javlja samo
kad je korpa **prikazana** lista — inače bi terao ponovno čitanje stanja ili
prometa koje se nije menjalo.

Posledica u pogonu: operater gleda Stanje, doda tri stavke, a značka i dalje piše
nulu — pa pređe na drugi ekran misleći da nema šta da proknjiži. Korpa **nije**
podatak u tabeli, pa „podaci su promenjeni" i „korpa je promenjena" nisu ista
stvar i ne smeju da dele isti kanal.

Ekran zato dobija `KorpaPromenjena` — jedno mesto za obe posledice promene korpe
(zona + `modOtkupUI.OsveziNavBrojace`, koji je već `Public`). Zove se sa **sva
četiri** mesta gde se korpa menja: dodavanje, uklanjanje, pražnjenje, upis.
Prekidač režima tu **ne spada**: značka sabira obe korpe, pa prelazak ne menja
broj — menja samo koja se korpa vidi u traci.

> Cena: `OsveziNavBrojace` pita svaki ekran, a većina brojača je prolaz kroz
> tabele. Zato ovo ide na **klik** (Dodaj / Ukloni / Završi), a nikad iz
> `OsveziZonu` — zonu osvežava i svaki otkucaj u polju.

#### Test 89 — nalaz, ne zakrpa

Review javlja `T_ZonaAgro_PrekidacRezimaZadrzavaBoju` kao jedini crveni test.
**Nije reprodukovan** — u ovoj sesiji nema Excela (`run_vba.py` odbija, `pywin32`
nema), pa se ne zna koja je tvrdnja pala.

Ono što se **može** utvrditi čitanjem: test je tvrdnje postavljao **dok forma
živi**, a `T_ZonaAgro_PoljaPostojeIPrateRezim` u istom fajlu dokumentuje zašto to
ne valja — dok forma živi, njena mašinerija briše `Err` između `Err.Raise` i
omotnice testa, pa pad stiže kao **`greska bez opisa`** i ne kaže koja tvrdnja je
pala. Test je zato prestrojen na isti oblik: forma se prvo **izmeri**, pa se tvrdi
**posle `ReleaseOtkupUIForm`**. (Usput je dobio i uredno otpuštanje forme — do
sada je zvao goli `Unload`, pa je `mFrm` ostajao na oborenoj formi.)

To ne popravlja uzrok ako uzrok postoji — **popravlja dijagnostiku**: sledeće
puštanje ili je zeleno, ili imenuje tvrdnju i vrednost (`AssertEq` već ispisuje
`ocekivano [X], dobijeno [Y]`).

#### Verifikacija

Testovi **91** (`T_Agro_KorpaUklanjaPoIdentitetu`) i **92**
(`T_Agro_ZnackaPratiKorpuVanKorpeListe`); četiri nove sabotaže
(`agro-korpa-bez-identiteta`, `agro-identitet-ne-stize-do-mreze`,
`agro-identitet-vidljiv`, `agro-znacka-ne-prati-korpu`) — ukupno **sedamnaest**
agro sabotaža.

Test 91 meri **bez mreže**: mreža bi uvela sortiranje i stranice u tvrdnju koja
je o identitetu. Ono što mreža mora da uradi — da identitet **prenese** i da ga
**ne nacrta** — tvrdi se nad opisom kolona i nad redovima koje `Scr_Rows` vraća.

Test 92 ne pokriva da baš `DodajUKorpu` / `IsprazniKorpu` / `ZavrsiUnos` zovu
`KorpaPromenjena` — te tri rutine čitaju zonu, a zone u testu nema. To stoji u
kodu (nigde na tim mestima nije ostao goli `OsveziZonu`) i mereno je sabotažom.

**Neverifikovano:** `RunAllTests` nije puštan (nema Excela) — testovi 91 i 92 i
sve četiri nove sabotaže **nisu izvršeni**. Compile nije prošao.

> **Naknadno pušteno i zeleno.** `RunAllTests` **96 testova**, testovi 91 i 92
> prolaze, sve četiri sabotaže obaraju imenovani test. Numeracija je pomerena:
> Palete su posle rebase-a uzele 93–96.

### 7.10 Rez fonta: nalaz iz testa 89 (`v6-ui-175`)

Test 89 je bio jedini crveni test kroz tri release-a. Uzrok nije bio u
Agrohemiji nego u **ljusci** — samo se video prvi put tek kad je jedna tvrdnja
zatražila da **neizabran** segment **nije** bold.

`modUiKit.NewLbl` je ignorisao traženi rez: kontrola građena sa `bold=False`
izlazila je sa `Font.Weight = 700`. Kvar je **uniforman** — svaka runtime
kontrola je bila bold — pa se nije ni primetio. Izgledao je kao odluka dizajna.

#### Šta je merenje isključilo

Sonda u testu 89, šest krugova nad živom formom. Svaki krug je gasio po jedno
objašnjenje, i nijedno nije preživelo:

| Kandidat | Merenje | Presuda |
|---|---|---|
| raspored / `RebaseSink` / test-seam | `gradnja U=1`, `rez=IZLAZ`, `bgU=B` | kontrola **nastaje** bold |
| osobina segmenta | `nova0=1` — i gola `NewLbl` je bold | univerzalno za sve runtime kontrole |
| artefakt čitanja `Font.Bold` | `tezinaU=700` | rez je **stvarno** bold |
| font forme (nasleđivanje) | `forma=0` | ne |
| ožičenje (`clsFlatBtn`) | `ozicen=400` | ne |
| povratna vrednost (`Set x = NewLbl(...)`) | `ivicaB=700`, a nije povratna | ne |
| redosled upisa | `r1=r2=r3=700` | ni pozadina pre fonta, ni ponovni lookup, ni dvostruki upis |
| `BackColor` obara font | `s2=400` | ne |

Poslednje merenje je pokazalo i zašto ništa od toga nije radilo:

```
s1=400   upis rez=False nad izgrađenom kontrolom  -> PROLAZI
s2=400   upis BackColor nad istom kontrolom       -> font nedirnut
s3=700   JOŠ JEDAN isti takav upis rez=False      -> vrati 700
```

Upis `Font.Bold` **nije ni pouzdan ni idempotentan**.

#### Popravka: tvrdi se ishod, ne mehanizam

`modUiKit.PostaviRez(ctl, bold)` — upiši, pročitaj, i ako nije ono što je
traženo, upiši opet; najviše tri puta. Merilo je **`Font.Weight`** (400
normalan, 700 bold), jer je `Font.Bold` iz nje izveden i sam ume da prevari.
Petlja je ograničena: ekran koji se zavrti je gori kvar od pogrešnog reza.

Koriste ga `NewLbl`, `NewTxt` i `BoxState` — sva tri mesta na kojima se rez
uopšte postavlja.

Test 89 od tada čita `Font.Weight`, ne `Font.Bold`, i tvrdi ga za ispunu **i**
natpis oba segmenta, u gradnji **i** u rasporedu, jednom tvrdnjom. Sonde su
uklonjene — ostavljene bi tvrdile zatečena ponašanja MSForms-a kao da su ugovor.

#### Dve posledice kroz ceo UI

1. **Izgled.** Bold sada nosi samo ono što ga i traži: naslovi, izabrani
   segmenti, čipovi, brojčana polja, zbirovi. Ostalo prelazi u normalan rez.
2. **Ponašanje.** `clsFlatBtn.IsSelected` čita baš taj rez i za `"nav"`,
   `"chip"` i `"seg"` je do sada bio **uvek True**, pa hover nije prefarbavao
   nijedno od njih. Sada razlikuje izabrano od neizabranog, kako je i
   projektovano.

Sabotaža `ljuska-rez-bez-potvrde` vraća upis na jedan pokušaj bez čitanja.

---

## 8. Fakture — šta je preneto (`v6-ui-176`)

Drugi ekran **Faze E**, stavka 16. Red u registru (`modUiScreens.ScrRows`)
je postojao od `S3a` — stavka menija se do sada crtala prigušena jer modula
nije bilo. Ovim se piše modul koji taj red već očekuje; **registar se ne dira**.

### 8.1 Gde je šta završilo

| Legacy (`frmFakturisanje`) | Novo mesto |
|---|---|
| `cmbKupac` + `FillComboDisplayID` | polje zone `scrFkKup` (`NewFieldG`) |
| `btnUnesi_Click` (punjenje `lstPrijemnice`) | lista **ZAFAKT**, čitač `modFaktura.GetPrijemniceZaFakturisanjeForGrid` |
| `lstPrijemnice` MultiSelect | **korpa** ekrana + kolona oznake + traka u zoni |
| `chkPrikaziFakturisane` | čip **Fakturisane** na listi ZAFAKT |
| `btnIzradiFakturu_Click` provere | `modFaktura.CreateFaktura_TX` (nepromenjen) |
| `CalculateTotal` | `modScrFakture.FkZbirKorpe` (samo prikaz — iznos računa transakcija) |
| `cmbFaktura` + `FillFaktureZaKupca` | lista **FAKTURE**, čitač `modFaktura.GetFaktureForGrid` |
| `btnStampaj_Click` | radnja nad redom `fkprint` → `modFaktura.PrintFaktura` |
| `btnSEF_Click` (otvara `frmSEF`) | lista **SEF** sa pet radnji nad redom |
| status plaćanja (nigde u legacy-ju) | radnja `fkstat` → `modFaktura.UpdateFakturaStatus` |

Nove rutine za mrežu (ekran ne čita tabele sam):
`modFaktura.GetPrijemniceZaFakturisanjeForGrid`, `GetFaktureForGrid`,
`GetFaktureSEFForGrid`, `SEFKonfigurisan`.

Uz njih je iz `IsPrijemnicaAvailableForFaktura` izdvojeno pravilo
**`modFaktura.PrijemnicaDostupna`** — jedno mesto koje sada dele kapija
transakcije i čitač mreže. Do sada je stajalo samo u kapiji, pa bi ga čitač
morao prepisati, a prepisana kopija se razilazi (isti obrazac kao
`GetAgroAbzugMapa` u §7.8).

### 8.2 Šta ekran uzima od ljuske

Ništa nije napravljeno za ovaj ekran — sve postoji od Faze C i Agrohemije.
**Diff u `modOtkupUI` je NULA.**

| Potreba | Ljuskin ugovor | Otkad |
|---|---|---|
| polje (natpis + okvir + kontrola) | `modOtkupUI.NewFieldG` | `v6-ui-159` |
| raspored unutar polja | `modOtkupUI.LayoutFieldInner` | `v6-ui-159` |
| identitet iz reda mreže | `modOtkupUI.GridCell` | `v6-ui-143` |
| značka uz stavku menija | `modOtkupUI.OsveziNavBrojace` | `v6-ui-172` |
| skrivena kolona (prioritet 4) | `LayoutGrid` crta do 3 | postoji od početka |
| status kao znak u redu | kolona vrste `paypill` | `v6-ui-113` |
| prijava prekoračenja bazena | `modOtkupUI.BazenStaje` | `v6-ui-170` |
| osvežavanje mreže na zahtev ekrana | `modOtkupUI.RefreshFromData` | `S4b` |

**Jedno mesto gde ugovor ne daje ono što bi se očekivalo:** ljuska **ne gleda**
povratnu vrednost `chg:` događaja — `UiClick` zove `ScrAct "chg:" & tag` pa
odmah `Exit Sub`. Za Agrohemiju to ne smeta: nijedna njena lista ne zavisi od
polja zone. Ovde lista prijemnica **jeste** lista jednog kupca, pa bi bez
reakcije ostala na prethodnom kupcu do sledećeg klika bilo gde — što izgleda
kao da izbor kupca ne radi.

Ekran to rešava **kod sebe**: kad se ukucano razreši u *drugog* kupca, sam
zove `modOtkupUI.RefreshFromData` (javnu od `S4b`). Provera „drugi kupac
nego prošli put" je bitna — `chg:` stiže na **svaki otkucaj**, pa bi bez nje
svaki znak povukao pun prolaz kroz tabele. **Ugovor ljuske se ovim ne menja.**

### 8.3 Odluka o SEF-u: radnje nad redom, ne ekran

SEF **ne postaje svoj ekran**. Četiri operacije koje operater radi nad
**jednom** fakturom (`SendInvoiceToSEF_TX`, `RefreshSEFStatus_TX`,
`CancelInvoiceOnSEF_TX`, `StornoInvoiceOnSEF_TX`,
`RecoverStuckSEFSendingInvoice`) su po obliku radnje nad redom, pa tu i idu.

Ali **ne na listu FAKTURE** — nego na **svoju listu**. `MAX_ACT` je 5, a
lista faktura već nosi dve radnje; pet SEF operacija bi dalo sedam ukupno i
**višak bi se tiho odsekao** (`RefreshRowActions` radi `Exit For`). Isti kvar
je već plaćen na listi paleta (`v6-ui-162`). Zasebna lista drži pet radnji
**tačno na granici bazena**, i to tvrdi test.

`frmSEF` ostaje **operativan i nepromenjen**, i nosi ono što lista ne može:
event log po fakturi, `PrepareResubmit`, i batch radnje
(`RecoverAllStuckSEFSendingInvoices`, refresh pending) — nijedna od njih nije
radnja nad jednim redom. Nijedan `modSEF*` modul nije diran.

#### Kapija na radnji, ne na listi — ispravka posle prvog smoke-a

Prva verzija je listu SEF-a **krila** kad `SEF_BASE_URL` / `SEF_API_KEY` nisu
upisani u `tblSEFConfig`. Smoke je pokazao zašto to ne valja: na radnoj svesci
sa fakturama a bez SEF naloga segmenta jednostavno **nema**, bez ijednog
objašnjenja.

Greška u proceni je bila u tome što lista ima **dva dela, a kapiju traži samo
jedan**:

| Deo liste | Traži podešen SEF? |
|---|---|
| čitanje stanja (`SEFWorkflowState`, SEF ID, poslato, greška) | **ne** — to su kolone `tblFakture` |
| radnje (pošalji, osveži, otkaži, storno, oporavi) | **da** |

Skrivanje cele liste zbog drugog dela je novi UI činilo **užim od legacy-ja**:
`frmFakturisanje.btnSEF_Click` otvara `frmSEF` **bezuslovno**, bez ijedne
provere configa. Operater je i pre ovoga mogao da vidi stanje bez naloga.

Sada: **lista postoji uvek**, a kapija stoji na jednom mestu kroz koje prolazi
svih pet radnji (`SefID` → `OTKUI_ERR_FK_SEF_OFF`, uz poruku koja kaže i **gde**
se podešava). `SEFKonfigurisan` je ostao — samo se više ne pita za listu.

Sporedna dobit: test 97 je time postao **jači**. Uslovna lista se mogla tvrditi
samo granom po `SEFKonfigurisan()`, a fixture je donor-zavisan — pa je ta grana
bila lutrija. Bezuslovna lista se tvrdi jednom brojkom.

### 8.4 Šta je namerno drugačije od legacy-ja

- **Multiselect → korpa.** Mreža bira jedan red, pa se prijemnice sakupljaju
  dugmetom. Korpa se vidi na tri mesta: kolona sa kvačicom u samoj listi,
  traka uz desnu ivicu zone, i značka uz stavku menija. Dvoklik na red
  prebacuje red u korpu i iz nje — najbliži parnjak klikanju po listi.
- **`NEPLACENE` je ČIP, ne lista.** To je lista FAKTURE sa filterom po
  statusu — iste kolone, isti čitač, isti identitet, iste radnje. Zasebna
  lista bi bila druga kopija istog čitača koja može da se raziđe.
  (`modNovac.GetOpenFakture` je uz to **po kupcu**, pa ni nije čitač za
  listu preko svih kupaca — koristi se tamo gde jeste po kupcu: KPI brojka
  „Neplaćeno" izabranog kupca.)
- **Nema polja za broj fakture.** Broj dodeljuje transakcija
  (`CreateFaktura` sam zove `GenerateBrojFakture`, koji je `Private`),
  operater ga ne bira. Polje sa „predlogom" koji transakcija ignoriše bio bi
  prikaz koji se garantovano razilazi sa upisanim. Broj stiže u poruci posle
  upisa i u listi faktura. (`modBrojevi.SuggestNextBroj` fakture ne poznaje —
  zna `OTK/OTP/ZBR/REV`, i format mu je `N/ddmmyy` po stanici, drugi niz.)
- **Stavka korpe nosi SAMO `PrijemnicaID`.** `CreateFaktura` svaku drugu
  vrednost iznova izvodi iz `tblPrijemnica` i eksplicitno veruje samo
  `stavka(0)`; dodatna polja bi bila mrtav teret koji navodi da se u njih
  veruje. Legacy prosleđuje pet polja, od kojih se četiri ignorišu.
- **Korpa se prazni pri promeni kupca**, uz poruku. `CreateFaktura` odbija
  prijemnicu drugog kupca (greška 1721), pa korpa koja preživi promenu kupca
  može samo da padne u transakciji.
- **Status plaćanja se može osvežiti iz ekrana** (`UpdateFakturaStatus`).
  Legacy tu funkciju ima u modulu, ali je nijedno dugme ne zove.
- **Kolona „Fakturisano" pokazuje samo broj fakture.** Legacy je u istu
  kolonu pakovao i `uplaćeno/iznos`; ta dva broja sada imaju svoje kolone u
  listi FAKTURE, gde im je mesto.

### 8.5 Identitet — pravilo koje je već dvaput plaćeno

Svaka od tri liste nosi identitet u **poslednjoj koloni, prioriteta 4**;
`LayoutGrid` crta do 3, pa vrednost postoji u modelu a ćelija se nikad ne
pravi. Radnja je čita kroz `GridCell`. Mape „prikaz → ID" sa strane nema:
mreža sortira i deli na strane, pa bi svaka takva mapa zastarela na prvi
klik po zaglavlju.

**Broj fakture se NE resetuje po godini** — za razliku od broja palete.
`GenerateBrojPalete` vraća goli `Long`, a godina živi u zasebnoj koloni
`Godina`, pa broj `1` stvarno postoji dvaput. `GenerateBrojFakture` vraća
`"N/YYYY"` — godina je **u stringu**, pa kolizije po godini nema.

Identitet je svejedno `FakturaID`, iz drugih razloga: ništa u kodu ne brani
dva reda istog `BrojFakture` (`RequireSingleFakturaRow` čuva **FakturaID**,
ne broj), a redove bez `/` (uvoz, ručni unos) skener maksimuma tiho
preskače, pa se broj može ponoviti sa ručno unetim.

**Dvosmislen ID nosi PRAZAN identitet.** `modFaktura.BrojacIdova` broji
pojave svakog ID-a nad **sirovom** tabelom (i stornirani red čini ID
dvosmislenim — `FindRows`, koji na kraju odlučuje, i njega vidi), a
`IdIliPrazno` vraća prazno za sve što nije tačno jedno. Radnja tada
**odbija** da bira i javlja porukom. Bez toga bi radnja svakako pukla —
`RequireSingleFakturaRow` i `CreateFaktura` fail-close-uju na duplikat — ali
kao greška transakcije umesto kao poruka operateru.

**Dostupnost se takođe PRENOSI u redu** (kolona 11, prioritet 4), ne izvodi
iz prikaza. Prijemnica obeležena kao fakturisana a **bez** `FakturaID` ima
praznu kolonu fakture: iz prikaza izgleda slobodna, a kapija je odbija. Ko
dostupnost čita iz onoga što se vidi, ponudi je operateru pa padne u
transakciji.

### 8.6 Prolazno stanje ima svoj kanal

Korpa **nije** podatak u tabeli, pa „podaci su promenjeni" i „korpa je
promenjena" nisu ista stvar i ne dele isti kanal — isto pravilo kao §7.9/P2.
`KorpaPromenjena` je jedno mesto za obe posledice (zona +
`modOtkupUI.OsveziNavBrojace`) i zove se sa **sva četiri** mesta gde se korpa
menja: dodavanje, uklanjanje, pražnjenje, upis. Promena kupca je peto —
ona korpu prazni, pa i ona ide kroz isti kanal.

### 8.7 Šta NIJE preneto

- `frmFakturisanje` i `frmSEF` se **ne gase i ne menjaju** — isto pravilo kao
  za `frmOtkup`, `frmDokumenta` i `frmAgrohemija` (§5, Faza B; §7.4).
  Dve kopije poslovne logike postoje namerno.
- **Avans se i dalje obračunava sam**, unutar `CreateFaktura`
  (`ApplyAvansToFaktura`). Ekran o tome ne zna ništa i ne prikazuje ga —
  isto kao legacy.
- **`PrepareResubmit` i batch SEF radnje** ostaju u `frmSEF` (v. §8.3).
- **SEF event log** nije prenet: to je istorija po fakturi, a mreža ljuske
  ima jedan nivo. `frmSEF` ga i dalje pokazuje.
- **Storno fakture** nije ovde — to je posao ekrana Storno.

### 8.8 Verifikacija

Testovi **97–103** u `modTest` i **šesnaest** sabotaža, uz nove fixture redove
u `tools/make_fixture.py`:
tri fakture (`FAK-TEST-N` neplaćena drugog kupca, `FAK-TEST-P` plaćena u
celosti, `FAK-TEST-X` stornirana), jedna uplata po fakturi (jedina u fixture-u
koja nosi `FakturaID`), i tri prijemnice (`PRJ-FAK-1` uredno fakturisana,
`PRJ-FAK-2` obeležena **bez** `FakturaID`, `PRJ-FAK-3` slobodna).

Do sada `tblFakture` nije imao ni broj, ni datum, ni status, a nijedna uplata
nije bila vezana za fakturu — pa je svaka tvrdnja o listi faktura i čipovima
radila nad praznim skupom i bila zelena bez pokrića.

| Test | Šta meri | Sabotaža |
|---|---|---|
| `T_Fak_UgovorEkrana` | registar, **tri liste bezuslovno**, granice bazena (`MAX_ACT` tačno 5 na SEF-u, `MAX_CHIP`, `MAX_COLS`, `MaxPrekidaca`), prvi čip je najširi | `fakture-sef-lista-uslovna`, `fakture-sef-sesta-radnja`, `fakture-cip-sve-nije-prvi` |
| `T_Fak_IdentitetURedu_NeCrtaSe` | identitet u poslednjoj koloni prioriteta 4; dvosmislen ID → prazno | `fakture-identitet-vidljiv`, `fakture-dvosmislen-prvi-pobedjuje` |
| `T_Fak_DostupnostSePrenosiURedu` | pravilo `PrijemnicaDostupna`; red **prenosi** dostupnost umesto da je izvodi iz prikaza | `fakture-dostupnost-iz-prikaza`, `fakture-dostupnost-bez-oznake` |
| `T_Fak_KorpaZnackaITraka` | uklanjanje po identitetu, značka van korpe-liste, traka: najnovije prvo + preliv se prijavljuje | `fakture-korpa-uklanja-prvu`, `fakture-znacka-ne-prati-korpu`, `fakture-traka-najstarije-prvo`, `fakture-traka-bez-preliva` |
| `T_Fak_CipoviPrateStatusFakture` | `paypill` šifre, čip „plaćene" se slaže sa znakom u redu, čip „neplaćene" se slaže sa `GetOpenFakture` **po svakom kupcu**, stornirana nije u listi | `fakture-prazna-je-placena`, `fakture-nepl-ignorise-status`, `fakture-stornirana-u-listi` |

Tvrdnja koja nosi najviše: **skup faktura koje čip „neplaćene" propušta za
datog kupca mora biti identičan onome što `modNovac.GetOpenFakture` vraća** —
isti oblik kao `T_Agro_AbzugMapaPratiPojedinacni`. Pravilo „otvorena faktura"
živi na dva mesta i može da se raziđe; ovo je jedino što bi to primetilo.

**Čipovi i radnje se čitaju po KLJUČU liste** (`FkCipoviZaListu`,
`FkRadnjeZaListu`, `FkKoloneZaListu`), ne kroz `Scr_Lista`. Razlog:
`Scr_Lista` je gate-ovana SEF konfiguracijom, a fixture je **donor-zavisan**
(`make_fixture` u `KEEP_ROWS` ne briše `tblConfig`), pa bi test vezan za nju
bio lutrija — zelen na jednom donoru, neizvršen na drugom.

### 8.9 Nalaz iz sesije: `vba_check` ne vidi nedeklarisanu promenljivu

Tokom rada je jedna python patch-skripta nezaštićenim `str.replace` pogodila
i **tuđu** proceduru (`T_Agro_UgovorEkrana`) i tamo uvela `kv`, koje u njoj
nije deklarisano. `modTest` ima `Option Explicit`, pa se modul više nije
kompajlirao.

Simptom nije ličio na uzrok: `vba_check` **zelen**, a `run_vba` visi ~230s pa
vrati `Exception occurred`, uz Excel zaostao u `[break]`. Isto što
`sabotaza.py` opisuje kao zamku 4 (komentar posle `_`), samo iz drugog izvora.

Merenje koje je to razrešilo: pokretanje testova **pojedinačno** kroz
`Application.Run` (svi prolaze) naspram **jednog dugog poziva** (staje na 82).
Razlika nije bila u testu 82 nego u tome da modul uopšte ne kompajlira, a
compile dijalog u nevidljivom Excelu nema ko da zatvori.

`vba_check` ovo **ne hvata i ne tvrdi da hvata** — `.claude/rules/testovi.md`
izričito kaže „ne kompajlira VBA: ne hvata tip-greške ni nedeklarisane
promenljive". Nalaz se ovde beleži, ne krpi: proširenje checkera je zaseban
posao sa svojim dvosmernim dokazom, a ne prilog uz ekran.

**Pouka za patch-skripte nad `src-vba/`:** svaka zamena mora da tvrdi broj
pogodaka. Nezaštićen `replace` je isti razred greške kao `sed -i` nad CRLF —
tiho pogodi više nego što je traženo.

### 8.10 Review PR #217: tri ispravke (`v6-ui-176`)

Tri nalaza iz review-a. Prva dva su moja regresija, treći je nasleđen i ovim se
konačno zatvara.

#### R1 — `LogErr` briše `Err`, pa se greška mora čitati **pre** njega

Tri nova čitača su završavala sa:

```vb
EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.Description
```

`modLogError.LogError` počinje sa `On Error Resume Next`, a **svaka `On Error`
naredba u VBA briše `Err`**. Posle `LogErr` je `Err.Number = 0` i opis prazan, pa
gornje postane `Err.Raise 0, SRC, ""` — originalna greška prestane da postoji.

Posledica nije rušenje nego **tihi gubitak**: `RequireColumnIndex` uredno digne
grešku zbog nedostajuće kolone, čitač je proguta u prazno, i `LoadGridFromScreen`
to vidi kao „nema redova" umesto kao pad šeme. To ruši fail-closed read model
koji je Storno već dobio.

Ispravan oblik je **već bio u istom fajlu** (`CreateFaktura`, `CreateFaktura_TX`,
`UpdateFakturaStatus`) — čitači ga nisu ponovili.

Uz tri čitača popravljen je i **`PrintFaktura`**, koji je isti kvar nosio od
ranije. Ulazi u istu ispravku jer ga sada zove radnja ekrana `fkprint`, pa bi
operater na pad štampe dobio poruku bez razloga.

#### R2 — kucanje po polju kupca je praznilo celu korpu

Ljuska `Change` šalje ekranu **na svaki znak**, a `GetComboID` daje stabilan ID
samo dok je stavka stvarno izabrana (`ListIndex >= 0`). Čim operater krene da
kuca, `ListIndex` padne na `-1` i fallback iz parcijalnog teksta vrati `""`.

Uslov je bio `If nov = mKorpaKupac Then Exit Function` — a `"" <> "KUP-A"`, pa je
**prvo otkucano slovo bacilo celu neproknjiženu korpu**, i to a da drugi kupac
nije ni izabran. Komentar iznad koda je opisivao pravilo koje kod nije sprovodio.

Pravilo je izdvojeno u `FkKupacPromenjen(nov, stari)`: **prazan ID nije „drugi
kupac" nego nerazrešen unos**. Ako operater obriše tekst do praznog, korpa ostaje
vezana za prethodnog kupca — bezbedno, jer `IzradiFakturu` ionako odbija rad bez
razrešenog `KupacID`-a, a kad se stvarno izabere drugi kupac korpa se tada uredno
prazni.

#### R3 — pečat verzije je konačno podignut

`OTKUI_BUILD` je stajao na `v6-ui-173` dok je `modUiKit` bio `v6-ui-175`, a ovaj
ekran `v6-ui-176`. Pečat postoji da bi se u smoke-u odmah videlo **da li je pravi
kod uopšte uvezen** — a takav je tvrdio treću stvar. Podignut je na `v6-ui-176`.

To je **jedina linija diffa u `modOtkupUI`** u ovom PR-u, i razlog joj nije
„ekran Fakture" nego „pečat mora da govori istinu". `StaraKomponenta` poredi sa
`OTKUI_MIN_BUILD`, ne sa `OTKUI_BUILD`, pa promena ništa ne pomera.

#### Fixture više ne nasleđuje test-kritičan config

`make_fixture` čuva `tblSEFConfig` (`KEEP_ROWS`), pa je svaki ključ koji fixture
ne postavi ostajao **donorov**. Ista suite je zato davala različit rezultat na
dve sveske, i to je nekoliko PR-ova nosilo kao „dva crvena ali nisu moja":

| Ključ | Šta je kvario |
|---|---|
| `DEFAULT_SORTA_VOCA` | `ApplyDefaultProizvod` napuni combo, a golden očekuje prazan |
| `KES_ISPLATE` | `IsKesIsplate` gasi granu „isplata iz OM avansa", pa kapija ćuti |

Oba se sada **pinuju** (prazno, odnosno `YES`). Prazan string je i dalje „nije
postavljeno" za `ApplyDefaultProizvod`, ali je sada **zapisano** prazno, pa
donorova vrednost ne može da procuri.

Rezultat: `RunAllTests` **103 testa, 0 palih** — prvi put bez zatečenih crvenih.

Pušten je i **pun set** (`run_vba.py --all`), jer ovaj PR dira i `modOtkupUI`
(pečat) i putanje grešaka u `modFaktura`: **15 suite-ova OK**, tri `BLIND`
(bez fail-gate-a, po katalogu), i dve crvene — `RunGoogleSyncSmokeSuite` i
`RunMasterSyncSmokeSuite`. Obe padaju **identično sa izmenama sklonjenim**
(`git stash`), traže Google kredencijale kojih u headless runu nema, i ne
dodiruju nijedan fajl iz ovog PR-a.

#### Zašto static provera za R1 **nije** ušla

Review je predložio i usku static proveru: `LogErr` pa čitanje `Err.*` u istom
`EH` bloku = nalaz. Napisana je i **puštena nad celim repoom: 135 nalaza.**

Uzorak pokazuje da su i stvarni i lažni:

- `modStorno.bas:1149`, `modScrOporavak.bas:288` — **stvarni**; opis greške ide u
  poruku operateru i biće prazan.
- `modSetup.bas:1213` — **lažan**; `Err.Description` je tamo *argument* samog
  `LogError` poziva, samo u nastavku reda (`_`), pa se izračunava **pre** poziva.

135 nalaza sa nepoznatim udelom lažnih ne sme u feature PR — `vba_check` bi
postao crven za sve, a CI bi pao na `main`-u. To je isti precedent kao „406
lažnih nalaza" pri pokušaju širenja `ARNOST`-a (`.claude/rules/testovi.md`).

Provera je zato **vraćena**, a nalaz ostaje zapisan: traži spajanje nastavaka reda
pre skeniranja i trijažu ~135 mesta, sa svojim `--self-test` slučajevima. To je
zaseban PR.

Umesto nje, R1 čuva **test 103** (`T_Fak_GreskaNePreziviLogErr`), koji meri pravi
put: štampa nepostojeće fakture mora da stigne do pozivaoca sa brojem i opisom
koji imenuje fakturu, ne kao nula i prazan string.

#### 8.11 Kapija operatera: prošla

`Alt+F11 → Debug → Compile VBAProject` je **čist**, i smoke nad pravim podacima
je prošao: izbor kupca, sakupljanje u fakturu, izrada, lista faktura i SEF lista.

Time je stavka 16 Faze E zatvorena u oba smera — i u onome što headless meri i
u onome što ne vidi. Automatski verdikt compile-a je i ovaj put bio `NEJASNO`
(`run_vba` ga preko SendKeys ne ume da pročita), pa je ručna kapija bila jedini
izvor istine — kao i uvek.

Otvoreno ostaje samo ono što se ovog ekrana ne tiče: `RunGoogleSyncSmokeSuite` i
`RunMasterSyncSmokeSuite` u punom setu, koje traže Google kredencijale i padaju
identično na netaknutom `main`-u.

---

## 9. Banka uvoz — šta je preneto (`v6-ui-177`)

Treći ekran **Faze E**, stavka 17. Red u registru (`modUiScreens.ScrRows`) je
postojao od `S3a` — stavka menija se do sada crtala prigušena jer modula nije
bilo. Ovim se piše modul koji taj red već očekuje; **registar se ne dira**.

### 9.1 Gde je šta završilo

| Legacy (`frmBankaImport`) | Novo mesto |
|---|---|
| `lstBanka` + `LoadBankaRows` | lista **STAVKE**, čitač `modBankaMapiranje.GetBankaImportForGrid` |
| kolona `BIM` u `lstBanka` | **ne prikazuje se** — interna šifra (v. §9.5) |
| `lblIzvodSummary` (jedan, najnoviji izvod) | lista **IZVODI**, čitač `modBankaImport.GetBankaIzvodiForGrid` |
| KPI traka (`RefreshTopKpis` + `ComputeBankaMapState`) | `modBankaMapiranje.GetBankaImportKpi` (jedan prolaz umesto dva) |
| `btnAutoJedan_Click` | radnja nad redom `bmauto` → `AutoMapBankaImportRow_TX` |
| `btnAutoSve_Click` | radnja `bmsve` → `AutoMapAllBankaImport_TX` |
| runtime dugme „Mapiraj jake ključeve (N)" | radnja `bmjaki` → `AutoMapStrongKeysBankaImport_TX` |
| `btnSkip_Click` | radnja `bmskip` → `SkipBankaImportRow_TX` |
| `btnSacuvajRucno_Click` (tri grane) | radnja `bmrucno` + tri polja zone |
| `cmbMapTip` / `cmbPartner` / `cmbFaktura` / `cmbOtkupBlok` | polja zone `scrBuTip` / `scrBuPartner` / `scrBuCilj` |
| `ConfirmManyCandidatesSplit` + `SplitPreviewText` | `modScrBankaUvoz.PitajZaPodelu` + `TekstPodele` (isti `PlanBlokRaspodela`) |
| `ShowSelectedRow` (uplata → Kupac, isplata → Kooperant) | `PredloziTipZaRed` na `row:` događaj |
| `btnOsvezi_Click` | posao ljuske (`RefreshFromData`) |

Nove rutine za mrežu (ekran ne čita tabele sam):
`modBankaMapiranje.GetBankaImportForGrid`, `GetBankaImportKpi`,
`GetFaktureZaBimMapiranje`, `GetBlokoviZaBimMapiranje`;
`modBankaImport.GetBankaIzvodiForGrid`.

Uz njih su **izdvojena četiri pravila** koja su do sada bila zaključana:

| Pravilo | Bilo | Sada |
|---|---|---|
| „da li bi jak ključ zatvorio ovaj red" | dve grane u petlji `CountStrongKeyReadyBankaImport` | `modBankaMapiranje.BimJakiKljucInfo` / `BimJakKljucSpreman` |
| „stavka je još u redu za mapiranje" | uslov unutar `GetBankaImportOpen` | `BimOtvoren` |
| „koji blok ručno mapiranje stvarno uzima" | `frmBankaImport.EffectiveManualBlockNo` (Private) | `BimEfektivniBlok` |
| „traži li blok potvrdu podele" | `frmBankaImport.SafeBlockCandidates` (Private) | `BimBlokTraziPotvrdu` |
| „smer stavke odgovara tipu mapiranja" | samo u writeru (`RequireBimSmer`) | `BimSmerOdgovaraTipu` (uz writera, ne umesto njega) |
| integritet izvoda (`početno + potražuje − duguje`) | inline u `UpdateIzvodSummaryLabel` | `modBankaImport.BimSaldoStatus` / `BimSaldoRazlika` |

Prvo od njih je isti obrazac kao `PrijemnicaDostupna` izdvojena iz
`IsPrijemnicaAvailableForFaktura` (§8.1): dok je pravilo stajalo samo u brojaču,
čitač mreže bi ga morao **prepisati**, a prepisana kopija se razilazi. Sada ga
brojač i čitač dele, i test to tvrdi jednom brojkom.

### 9.2 Šta ekran uzima od ljuske

Ništa nije napravljeno za ovaj ekran. Diff u `modOtkupUI` su **dve linije**, i
nijedna nije nova mogućnost:

1. **Pečat verzije**, `OTKUI_BUILD` → `v6-ui-177`. Razlog nije „ekran Uvoz
   izvoda" nego isti kao u §8.10/R3: pečat postoji da bi se u smoke-u odmah
   videlo **da li je pravi kod uopšte uvezen**. Sa `v6-ui-176` u sidebaru, a
   `v6-ui-177` u ekranskom modulu, tvrdio bi treću stvar. `StaraKomponenta`
   poredi sa `OTKUI_MIN_BUILD`, pa promena ništa ne pomera.
2. **`zOtp` dopisan u spisak zona koje pripadaju samo ekranu Dokumenta**
   (`ShowZones`) — v. §9.10.

| Potreba | Ljuskin ugovor | Otkad |
|---|---|---|
| polje (natpis + okvir + kontrola) | `modOtkupUI.NewFieldG` | `v6-ui-159` |
| raspored unutar polja | `modOtkupUI.LayoutFieldInner` | `v6-ui-159` |
| promena teksta stiže ekranu | `Scr_Event("chg:<kontrola>")` | `v6-ui-159` |
| identitet iz reda mreže | `modOtkupUI.GridCell` | `v6-ui-143` |
| skrivena kolona (prioritet 4) | `LayoutGrid` crta do 3 | postoji od početka |
| značka uz stavku menija | `ScrBrojac` → `Scr_Brojac` | `S4b` |
| lista bez ijedne radnje | `ActDefs` vrati `Empty`, raspored **sakrije** zaostalu dugmad | postoji od početka |

**Ono što ovaj ekran NE koristi, a Fakturisanje mora:** `RefreshFromData` iz
`chg:` grane. Ljuska ne gleda povratnu vrednost `chg:` događaja (§8.2), pa ekran
čija lista zavisi od polja zone mora sam da zatraži osvežavanje. Ovde nijedna
lista ne zavisi od polja — polja biraju **cilj** ručnog mapiranja, ne skup
redova — pa se `RefreshFromData` iz `chg:` grane **namerno ne zove**.

### 9.3 Odluka: UVOZ ne ulazi u ekran

Povlačenje PDF-ova, parsiranje i staging (`ImportBankaInbox_TX`,
`PullBankPdfsFromDriveProduction`) **ostaju van ekrana**. Razlog nije dužina
posla nego **ishod**:

- `ImportBankaInbox_TX` je `Sub` koji **ne vraća ništa**. `SaveBankaImportRowsCore`
  prebroji i upisane (`savedCount`) i duplikate (`duplicateCount`), ali
  `duplicateCount` završi u `Debug.Print`, a wrapper odbaci i `savedCount`.
  Dugme koje ne može da kaže „uvezeno N, duplikata M" bilo bi **tiho knjiženje**;
  da bi moglo, morao bi se menjati javni potpis jezgra uvoza — a to nije izmena
  koja pripada PR-u o UI-ju.
- Uvoz uz to **pomera fajlove po disku** (`ExecutePendingBankaFileMoves`:
  Inbox → Processed/Error) i zavisi od `pdftotext`/Poppler. Mreža ljuske nema ni
  progres ni otkazivanje.
- I najvažnije za pravilo „ne praviti novi UI užim od legacy-ja" (§8.3):
  **`frmBankaImport` uvozno dugme nema.** Uvoz se oduvek pokreće zasebnom
  komandom (`ImportBankaInbox` / `_WithDrivePull`), a forma samo mapira. Ekran
  bez uvoza je **tačno širine legacy forme**, ne uži.

### 9.4 Odluka: RUČNO mapiranje ulazi u ekran

Suprotna odluka, i razlog joj nije „može da stane" nego **značka**. Auto i jaki
ključevi obrade lako; ono što ostane je po definiciji ono što traži ručno.
Ekran koji taj ostatak vidi a ne može da ga zatvori ima brojku koja iz njega
**nikad ne pada na nulu**, i šalje operatera u legacy formu baš za slučajeve
zbog kojih je red i napravljen vidljivim. To je isti kvar kao „Ukloni stavku"
koje Agrohemija nije imala (§7.3), samo skuplji.

Ručno mapiranje traži **tri spregnuta polja**, ne četiri: TIP, PARTNER i CILJ
(faktura za kupca, blok za kooperanta; za OM cilja nema, pa se polje gasi).
**Vrsta voća se NE bira** — ona je izlaz `PlanBlokRaspodela` (kolona 3) i vidi
se u predlogu podele; operater je nikad ne unosi.

Dve granice su postavljene namerno:

1. **Potvrda podele ostaje `MsgBox` sa ista tri ishoda** (DA = knjiži podelu,
   NE = ceo iznos kao avans kooperanta, OTKAZI = ne diraj stavku). Tri ishoda
   nad izračunatim tekstom nisu forma nego pitanje, a podelu računa **isti**
   `PlanBlokRaspodela` po kome se i knjiži.
2. **Ekran ne drži nijedno pravilo.** Sve što je bilo `Private` u formi je
   izdvojeno u `modBankaMapiranje` (tabela u §9.1). `frmBankaImport` zadržava
   svoju kopiju i **ne dira se** — isto pravilo kao za `frmOtkup`,
   `frmDokumenta`, `frmAgrohemija` i `frmFakturisanje`.

Najosetljivije od izdvojenog je **fail-closed nad listom faktura**: prazna lista
i **pad** učitavanja izgledaju isto, a znače suprotno — prazan izbor fakture
knjiži **AVANS** umesto zatvaranja duga. Zato `GetFaktureZaBimMapiranje` vraća
zastavicu `outOK`, a `BuSmeMapiranjeKupca` je imenovana odluka koju radnja čita
(inače bi brisanje jednog `If`-a prošlo neprimećeno).

### 9.5 Šta je namerno drugačije od legacy-ja

- **IZVODI su lista, a ne jedna labela.** `UpdateIzvodSummaryLabel` pokazuje
  **jedan** izvod — onaj sa najvećim `DatumIzvoda` — i nad njim radi proveru
  `početno + potražuje − duguje = završno`. Ekran to radi nad **svim** izvodima,
  grupisanim po `(BrojDokumenta + BrojRacuna)`. To je jedino mesto na kom se
  vidi da li se izvod slaže, i ovde je **šire** od legacy-ja.
- **„Obrađeno" i „preskočeno" su ČIPOVI, ne liste** — ista lista sa filterom po
  statusu, isti čitač, isti identitet, iste radnje. Zasebna lista bi bila druga
  kopija istog čitača koja može da se raziđe (§8.4).
- **„Za obradu" i „za ručno" nisu isto.** `GetBankaImportOpen` izbacuje samo
  `"Da"` i `"Skip"`, pa je red sa statusom `"Error"` (auto pokušao i odbio) i
  dalje **otvoren**. Bez oba čipa se ne vidi razlika između „još nije probano" i
  „probano pa vraćeno operateru".
- **Kolona „Predlog" umesto preview panela.** Legacy za IZABRANU stavku gradi
  višeredni tekst (`BuildAutoPreviewText`/`BuildManualPreviewText` i dve grane
  ispod njih, oko 350 linija). Ekran umesto toga nosi **jednu ćeliju po redu** —
  „faktura 2/2026", „blok 1/TEST", „avans kupca X", „nema jakog ključa",
  „nejasan smer" — pa se predlog vidi za **sve** redove odjednom, ne samo za
  izabrani. Cilj i njegovu oznaku računa čitač; ekran samo formuliše tekst.
- **Smer-kapija se vidi PRE klika.** `RequireBimSmer` u writeru ostaje, ali bi je
  operater osetio tek kao grešku transakcije. Ekran istu odluku
  (`BimSmerOdgovaraTipu`) postavlja pre poziva — isto što je legacy preview
  radio (AUD-025).
- **Dvoklik namerno ne radi ništa.** Na Fakturisanju i Agrohemiji dvoklik
  prebacuje red u korpu i iz nje — povratna radnja nad prolaznim stanjem. Ovde
  bi svaka radnja nad redom bila **knjiženje u `tblNovac`**, a knjiženje se ne
  pokreće promašenim dvoklikom.
- **Partner combo svuda prikazuje i ID** (`ShowIDInComboDisplay`). Dva partnera
  istog naziva su u ovim šifarnicima obična pojava (fixture ima dva istoimena
  kooperanta), a izbor pogrešnog šalje novac pogrešnom čoveku (FM-0024 #7).
- **`BankaImportID` se NE prikazuje.** Legacy ga ima kao prvu kolonu (`BIM`), ali
  to je interna šifra: operater ne zna čemu služi i ne može ništa s njom. Prva
  kolona je **broj izvoda** — jedini poslovni broj koji stavka nosi, i ono što
  `StyleGridCell` u prvoj koloni ionako crta kao broj dokumenta. Šifra ostaje u
  **pretrazi** (ko je ima iz loga ili poruke o grešci mora moći da nađe red) i u
  skrivenoj koloni identiteta.
- **Podnožje mreže pokazuje PROMET, ne neto.** Neto (uplate − isplate) je na
  čipu „obrađeno" davao **negativan** broj, koji nad izvodom ne znači ništa.
  Razdvojene brojke — koliko uplata, koliko isplata — stoje u traci iznad mreže;
  podnožje ljuske ima samo **jedan** slot (`grdFoot.ftVal`), pa dva broja u njemu
  traže dopunu ugovora ljuske i idu u zaseban PR.
- **Nema `Scr_NaslovDopuna`.** Naslov mreže je labela fiksne širine (`grdTitle`,
  180pt), pa se dopuna odsecala usred reči („— 29 z"). Broj koji je nosila već
  stoji u brojci OTVORENO iznad mreže i u čipu „za obradu"; odsečen tekst je
  gori od nikakvog.
- **Objašnjenje stoji UZ polja, ne ispod njih.** Ispod je nalegalo na traku koju
  ljuska crta odmah po završetku zone. Staje u prostor između poslednjeg polja i
  brojki; kad tog prostora nema, sklanja se — `Label` ne prelama, pa bi inače
  istekao preko brojki.

### 9.6 Identitet — pravilo koje je već triput plaćeno

Obe liste nose identitet u **poslednjoj koloni, prioriteta 4**; `LayoutGrid`
crta do 3, pa vrednost postoji u modelu a ćelija se nikad ne pravi. Radnja je
čita kroz `GridCell`. Mape „prikaz → ID" sa strane nema.

**Broj izvoda NIJE identitet.** Dedupe ključ (`IsDuplicateBankaImport`) počinje
od **broja računa** — „Drugi racun = druga transakcija, bez obzira na broj
izvoda i iznos" — pa dva računa firme legitimno nose izvod **istog broja**.
Identitet stavke je `BankaImportID`; identitet izvoda je
`BimIzvodKljuc(BrojDokumenta, BrojRacuna)`.

**Identitet nije ni u jednoj vidljivoj koloni.** `BankaImportID` je interna
šifra i iz prikaza je izbačen (§9.5); u redu postoji samo u skrivenoj koloni, i
to onakav kakav ga je čitač **proverio** (`modFaktura.IdIliPrazno` nad sirovom
tabelom). Dvosmislen ID nosi **prazno**, i radnja tada odbija da bira — bez toga
bi svakako pukla (`RequireSingleRow` fail-close-uje na duplikat), ali kao greška
transakcije umesto kao poruka.

Red sa dvosmislenim ID-em se pri tom **i dalje vidi u listi** — po datumu,
partneru i iznosu — pa operater zna koji je red odbijen. To i tvrdi test.

**Isto važi za sve što red PRENOSI a ne prikazuje jednoznačno:**

| Kolona (prio 4) | Zašto se ne izvodi iz prikaza |
|---|---|
| `OTVOREN` | nov red ima **prazan** status, pa se u mreži ne razlikuje od reda kome status nije upisan |
| `SMER` | red sa **i** uplatom **i** isplatom izgleda kao uplata (kolona uplate je popunjena), a writer ga odbija kao `NEJASAN` |

### 9.7 Značka ide kroz ljusku, bez privatnog kanala

Ovde je **obrnuto** od Agrohemije i Fakturisanja. Tamo je korpa prolazno stanje
van tabele, pa je morala sama da zove `OsveziNavBrojace` (§8.6). Ovde je red za
mapiranje **podatak u tabeli**: svaka radnja ga menja upisom, ljuska posle upisa
ionako zove `RefreshFromData`, i brojač je time već pokriven. Privatan kanal se
ne uvodi jer ne treba.

Broj koji značka nosi čita se iz **iste brojke** koju vidi i čip „za obradu"
(`GetBankaImportKpi`), ne iz zasebnog prolaza koji bi se s njim mogao razići.

### 9.8 Šta NIJE preneto

- **Uvoz** (v. §9.3). `frmBankaImport` i put `ImportBankaInbox` ostaju jedini
  način da se izvod uveze.
- **`frmBankaImport` i `frmBankaExportPregled` se ne gase i ne menjaju.** Dve
  kopije poslovne logike postoje namerno; pravilo se menja u `modBankaMapiranje`
  pa se ručno preslikava u formu.
- **Parseri** (`modBankaImportParserPdfToText`, `modBankaProCredit`,
  `modBankaHalk`, `modBankaAlta`) nisu dirani.
- **Veliki preview panel** — zamenjen kolonom „Predlog" (§9.5).
- **Nalozi za isplatu** (`frmBankaExportPregled`, CSV, specifikacija) nisu ovde —
  to je zaseban ekran (`BANKA_NALOZI`), sledeći na redu u Fazi E.
- **Storno stavke izvoda** nije ovde — to je posao ekrana Storno
  (`modStorno.StornoIzvod_TX`).

### 9.9 Verifikacija

Testovi **104–110** u `modTest` i **dvadeset jedna** sabotaža, uz nove fixture redove
u `tools/make_fixture.py`: **jedanaest** stavki izvoda u **četiri** grupe
`(broj + račun)` i tri otkupne stavke istog bloka.

Do sada `tblBankaImport` nije imao **nijedan** red, pa je svaka tvrdnja o listi,
čipovima, jakim ključevima i integritetu izvoda radila nad praznim skupom.

Svaki fixture red ima razlog:

| Red | Zašto postoji |
|---|---|
| `BIM-FIX-1` | jak ključ preko **fakture** (poziv na broj = broj fakture `2/2026`) |
| `BIM-FIX-2` | jak ključ preko **bloka** (poziv na broj = `BrojDokumenta` `1/TEST`) |
| `BIM-FIX-3` | bez ijednog jakog ključa → traži ručno |
| `BIM-FIX-K` | **drugi račun pod istim brojem izvoda** — kolizija koja dokazuje da broj izvoda nije identitet |
| `BIM-FIX-3K` | blok sa **tri** otvorene stavke → `ERR_BMAP_MANUAL_REQUIRED` |
| `BIM-FIX-ER` | `Obradjeno = "Error"` — auto pokušao i vratio operateru |
| `BIM-FIX-DA` / `BIM-FIX-SK` | obrađen i preskočen |
| `BIM-FIX-ST` | storniran — ne sme ni u jednu listu |
| `BIM-FIX-DUP` ×2 | **isti `BankaImportID` dvaput** — bez njih bi „dvosmislen ID nosi prazan identitet" merilo odsustvo reda |
| `IZV-FIX-1` / `IZV-FIX-2` | izvod koji se slaže i izvod kome fali 100 |

| Test | Šta meri | Sabotaža |
|---|---|---|
| `T_BankaUvoz_UgovorEkrana` | registar, dve liste, granice bazena (`MAX_ACT` tačno 5 na stavkama, `MAX_CHIP`, `MAX_COLS`, `MaxPrekidaca`), prvi čip je najširi, izvodi bez radnji, **datum stiže kao broj** | `banka-uvoz-sesta-radnja`, `banka-uvoz-cip-sve-nije-prvi`, `banka-uvoz-izvodi-imaju-radnju`, `banka-uvoz-datum-nije-broj` |
| `T_BankaUvoz_IdentitetURedu_NeCrtaSe` | identitet u prenosnoj koloni prioriteta 4; **interne šifre nema među vidljivim kolonama**; dvosmislen ID → prazno, a red se i dalje vidi; kolizija broja izvoda | `banka-uvoz-identitet-vidljiv`, `banka-uvoz-dvosmislen-prvi-pobedjuje` |
| `T_BankaUvoz_RedNosiSmerIOtvorenost` | red **prenosi** smer i otvorenost umesto da ih izvodi iz prikaza; `"Error"` je i dalje otvoren | `banka-uvoz-red-ne-nosi-otvorenost`, `banka-uvoz-red-ne-nosi-smer`, `banka-uvoz-predlog-i-za-zatvorene` |
| `T_BankaUvoz_CipJakihPratiBrojac` | čip „jaki ključevi" i `CountStrongKeyReadyBankaImport` vide **isti** skup; „sve" je unija tri stanja; značka = čip „za obradu" = `GetBankaImportOpen` | `banka-uvoz-cip-jaki-prolazi-sve`, `banka-uvoz-znacka-broji-mapirane`, `banka-uvoz-obradjeno-guta-preskoceno` |
| `T_BankaUvoz_IzvodiSuAgregatPoRacunu` | grupa je `(broj + račun)`; zbirovi se **uzimaju sa reda, ne sabiraju**; legacy red bez saldo podataka nije neslaganje | `banka-uvoz-izvod-kljuc-bez-racuna`, `banka-uvoz-saldo-se-sabira`, `banka-uvoz-legacy-red-je-razlika` |
| `T_BankaUvoz_RucnoMapiranjePravila` | smer-kapija se slaže sa writerom; prazan izbor bloka uzima poziv na broj; blok preko granice traži potvrdu; fail-closed nad listom faktura | `banka-uvoz-om-prima-nejasan-smer`, `banka-uvoz-prazan-blok-ostaje-prazan`, `banka-uvoz-fakture-fail-open`, `banka-uvoz-fakture-i-zatvorene` |
| `T_ZonaBankaUvoz_PoljaIRaspored` | zona se STVARNO gradi i raspoređuje; sve kontrole postoje; kombo je polje (`nm` + `nmT`); polje cilja je ugašeno za OM | `banka-uvoz-om-polje-cilja-radi` |

Tvrdnja koja nosi najviše: **broj redova koje propušta čip „jaki ključevi" mora
biti identičan onome što vraća `CountStrongKeyReadyBankaImport`** — isti oblik
kao `T_Agro_AbzugMapaPratiPojedinacni` i `T_Fak_CipoviPrateStatusFakture`.
Pravilo živi na dva mesta (čitač mreže i natpis dugmeta) i može da se raziđe;
ovo je jedino što bi to primetilo.

**Čipovi, radnje i kolone se čitaju po KLJUČU liste** (`BuCipoviZaListu`,
`BuRadnjeZaListu`, `BuKoloneZaListu`), ne kroz `Scr_Lista` — isti razlog kao
§8.8: ugovor svake liste mora da se meri bez prebacivanja stanja ekrana.

**Dvosmerni dokaz je pušten za svih dvadeset jednu**: svaka sabotaža obara
**tačno jedan** imenovani test i vraća se bit-identično. Bazna vrednost pre i posle je
`RunAllTests` **110 / 0**, a `RunBankaImportTestSuite` (tvrd fail-gate nad ovim
područjem) **PASS=189, FAIL=0**.

### 9.10 Nalazi iz sesije

#### `Dim src` i `Const SRC` su ISTO IME

Prvi run je pao ovako: `vba_check` **zelen**, `run_vba` visi **225 s** pa vrati
`Exception occurred`, a Excel ostane u `[break]` sa dijalogom
**„Duplicate declaration in current scope"**.

Uzrok: u `RedoviStavke` i `RedoviIzvodi` su stajali i `Dim src As Variant` i
`Const SRC As String`. **VBA je case-insensitive**, pa su to dva imena istog
identifikatora u istom opsegu — modul se ne kompajlira, a modul koji se ne
kompajlira obara **ceo projekat**.

To je isti simptom kao §8.9, iz trećeg izvora. `vba_check` ga **ne hvata i ne
tvrdi da hvata**: `DUPLIKAT_LOKALNI` izričito ne gleda `Const`/`Dim` **unutar**
procedure. Pokušaj uske statičke provere je puštan nad celim repoom i dao
**302 nalaza**, praktično sve lažne (hvata i argumente, i brojeve, i imena
konstanti u pozivima) — dakle isti razred kao „135 nalaza" za `LogErr` (§8.10) i
„406 lažnih" za `ARNOST`. **Nije ušla.** Nalaz se beleži: prava provera traži
parsiranje deklaracione liste, ne regex, i zaseban PR sa svojim `--self-test`.

Ispravka je bila da izvor greške ide kao literal u `Err.Raise`, kao u
`modScrFakture`.

#### `BIM_MAPTIP_*` konstante su mrtve

`modConfig` nosi `BIM_MAPTIP_FAKTURA` / `_KOOPERANT` / `_NEP` / `_PROVIZIJA` uz
komentar „MapTip values suggested". **Ne koristi ih nijedan modul.** Operativne
vrednosti su literali `"Kupac"` / `"Kooperant"` / `"OM"` iz
`frmBankaImport.cmbMapTip`, koji biraju koji se writer zove. Ekran zato uvodi
`BIM_TIP_*` u `modBankaMapiranje` sa **operativnim** vrednostima; mrtve
konstante se **ne diraju** (brisanje je zaseban, mehanički posao).

#### Test pisan nad fixture-om kakav nije

Prvo puštanje je oborilo `T_BankaUvoz_RucnoMapiranjePravila` na tvrdnji „kupac
ima bar jednu otvorenu fakturu". Produkcioni kod je bio ispravan: **raniji test
u istoj suite** (uplata na fakturu) zatvori `FAK-TEST-1` u celosti, pa
`FX_KUPAC` do testa 109 nema nijednu otvorenu fakturu. Tvrdnja je merila
**posledicu redosleda testova**, ne pravilo.

Ispravljeno tako da test koristi `FX_KUPAC2` i njegovu `FAK-TEST-N`, koju
nijedan test ne dira, a „zatvorena faktura ne ulazi u listu" se tvrdi nad
`FAK-TEST-P` — koja je plaćena **u samom fixture-u**, pa ne zavisi od redosleda.
Ista klasa greške kao dva pada u §7.6.

#### Dve sabotaže koje nisu oborile ništa — i šta su otkrile

Prvi prolaz dvosmernog dokaza dao je **16 / 18**. Nijedna od dve nije bila
greška u kodu; obe su bile greška **u dokazu**, i svaka je otkrila po nešto:

1. **`banka-uvoz-identitet-vidljiv` je otkrila grešku u samom testu.** Sabotaža
   pomera prioritet kolone identiteta sa 4 na 3, a test je proveravao
   **poslednju** kolonu liste. Identitet stavke **nije poslednji** — iza njega
   stoje još dve kolone koje red takođe samo *prenosi* (`OTVOREN`, `SMER`) — pa
   je tvrdnja merila **susednu** kolonu i pomeren prioritet identiteta bi prošao
   neprimećeno. Test sada gađa baš tu kolonu (po ključu `OTKUI_HDB_BIMKEY`) i uz
   to tvrdi da **sve tri** prenosne kolone ostaju van prikaza. Ovo je tačno ono
   zbog čega dvosmerni dokaz postoji: zelena tvrdnja koja nikad nije pokazana
   crvena ne dokazuje da išta meri.

   Ista sabotaža je uz to bila napisana sa **komentarom posle `_`**, što je
   syntax error, pa je obarala **compile** umesto testa (run visi, izlaz je
   `Exception occurred`). To je zamka 4 koju `sabotaza.py` opisuje **u svom
   docstring-u** — i svejedno je pokupljena.

2. **`banka-uvoz-saldo-se-sabira` je sabirala nulu.** Stajala je u grani koja se
   izvršava samo za **prvi** red grupe, gde je akumulator još prazan, pa
   „sabiranje" nije menjalo ništa. Premeštena je uz brojač stavki, koji radi za
   svaki red grupe.

Posle ispravki: **18 / 18** (sa kasnije dodatom `banka-uvoz-datum-nije-broj`
ukupno **19 / 19**).

#### Smoke nad pravim podacima: šest nalaza koje suite nije mogao da vidi

Prvi smoke je oborio šest stvari. Vredi ih razdvojiti po tome **zašto** ih
headless nije uhvatio:

**Nijedna tvrdnja nije čitala datum.** Kolona DATUM je bila prazna u svakom
redu. Ekran je mreži predavao vrednost ćelije kakva jeste — `Date` — a ljuskin
`FmtDatumKratko` počinje sa `If Not IsNumeric(v) Then Exit Function`, a
**`IsNumeric` je nad `Date`-om `False`**. Ćelija ostane prazna, bez ijedne
greške i bez traga u logu. Ostali ekrani datum konvertuju u serijski broj
(`modScrDokumenti.DatSerijski`, `modUiData.CellDate`) — ovaj sada takođe.
Test 104 od sada prolazi kroz opis kolona, nalazi svaku `date` kolonu i tvrdi da
je vrednost `IsNumeric`; sabotaža `banka-uvoz-datum-nije-broj` to obara.

**Prikaz se ne meri tvrdnjom nego okom.** Objašnjenje ispod polja naleglo je na
traku koju ljuska crta odmah po završetku zone; naslov mreže se odsekao usred
reči („— 29 z") jer je `grdTitle` fiksnih 180pt; podnožje je na čipu „obrađeno"
pokazivalo **negativan** zbir. Sve tri su popravljene (§9.5).

**Šta je operateru korisno nije stvar koda.** `BankaImportID` u prvoj koloni je
interna šifra — operater ne zna čemu služi. Izbačena je iz prikaza; prva kolona
je sada broj izvoda.

**Pečat je opet lagao.** Sidebar je pisao `v6-ui-176` dok je ekranski modul bio
`v6-ui-177`, pa se iz smoke-a nije moglo videti da li je nov kod uopšte uvezen —
tačno ono što je §8.10/R3 zatvarao. `OTKUI_BUILD` je podignut.

#### Sedmi nalaz: `Private Const` iz tuđeg modula, i rupa koju je otvorio

Posle prve runde ispravki compile je pukao na uvozu: **`Variable not defined`**
nad `GAP` u `RasporediPolja`. `modOtkupUI.GAP` je **`Private Const`** — susedni
`PAD` je `Public`, pa je propust bio lak. Ispravka je jedna reč (ekran ima svoj
`BU_FLD_GAP`), ali zanimljivo je **zašto je prošlo kroz sve kapije**:

- `vba_check` zna da ime `GAP` u repou postoji, ali **ne prati vidljivost** —
  nema pojam „`Private` u drugom modulu".
- `RunAllTests` je bio **zelen**. VBA „`Variable not defined`" prijavljuje tek
  kad se procedura **prvi put izvrši**, a ovo je greška u telu `RasporediPolja`
  — procedure koju **nijedan test nije zvao**, jer sve ostale tvrdnje o ekranu
  rade nad čitačima i pravilima, gde zone nema. (Za razliku od dupliranog
  `Const SRC`, koji je greška na nivou modula i obara ga odmah.)

Zato uz ispravku ide **test 110**, koji zonu stvarno gradi (`Scr_Build`) i
raspoređuje (`Scr_Layout`) nad pravom formom — isti obrazac kao
`T_ZonaAgro_PoljaPostojeIPrateRezim`. Time je taj put od sada pokriven, a uz
njega se tvrdi i pravilo koje se drugačije ne može izmeriti: **polje cilja je
ugašeno za OM**.

Provera vidljivosti je puštena i nad celim ekranskim modulom (skripta u
scratchpad-u, ne u repou): posle ispravke **nula** identifikatora koje modul
koristi a koji su `Private` drugde. Proširenje `vba_check`-a na tu proveru je
zaseban posao, sa svojim dvosmernim dokazom.

#### Drugi smoke: dve stvari koje je razrešilo tek MERENJE

Posle prve runde ispravki kolona DATUM je bila **gora**: umesto prazne, u njoj
je pisalo `OSIROCENE_PAL` — što je `OSIROCENE_PALETE`, **vrsta reda sa ekrana
Oporavak**, odsečena na širinu kolone. Ostale kolone su bile tačne.

Čitanje koda nije moglo da razreši: opis kolone je identičan onome koji koristi
`modScrDokumenti` (`"OTKUI_HD_DATUM||date|NN|1"`, druga kolona), vrednost ide
kroz isti `modUiData.CellDate`, `LayoutGrid` koloni daje širinu, a test nad
fixture-om tvrdi `IsNumeric` — i prolazi. Zato je u modul dodata dijagnostika
**`Diag_BuRedovi`** (presedan: `modBankaImport.Diag_DumpPdfTextAroundStanje`),
koja ispisuje šta ekran **predaje** mreži i šta mreža **drži**.

Merenje nad pravom sveskom:

```
EKRAN red 1 kol2: tip=Double vred=[26062026] IsNumeric=True
```

**`26062026` nije serijski broj datuma nego `ddmmyyyy` upisan kao BROJ.**
Ljuska nad kolonom tipa `date` radi `CDate` (`FmtDatumKratko`), `CDate` van
opsega **pukne**, `RenderGrid` radi pod `On Error Resume Next` — pa upis ćelije
bude preskočen i u njoj **ostane natpis od ranijeg crtanja**. Bez greške, bez
traga u logu, sa tuđim tekstom u koloni.

Ekran zato dobija kapiju **`BuDatumUOpsegu`**: u kolonu datuma ulazi samo ono
što `CDate` sme da primi; sve ostalo je 0, tj. prazna ćelija. Vrednost se
**ne tumači** — `ddmmyyyy` nije oblik koji `modParse.TryParseDateValue` poznaje,
pa bi tumačenje bilo izmišljanje pravila koje domen nema.

**Nalaz veći od ovog ekrana:** takav red je posejan u fixture da bi tvrdnja imala
nad čim da padne — i oborio je **sedam** testova sa `Overflow`, među njima i
`T_StornoEkran_SvakaListaVracaRedove`. Dakle tu vrstu podatka **ne podnosi samo
ovaj ekran**. Sejanje je vraćeno (potpis fixture-a je posle vraćanja
bit-identičan, što dokazuje da je uzrok bila isključivo ta vrednost), pravilo se
tvrdi direktno, a nalaz je otišao u zaseban posao. Prava kapija verovatno ne
pripada ekranu nego `modUiData.CellDate` (vraća bilo koji broj kao „datum") ili
`FmtDatumKratko` (čuva samo donju granicu) — to je odluka, ne pretpostavka.

#### Traka „Nema izabrane otpremnice…" — moja prva dijagnoza je bila pogrešna

Prvo sam je pripisao `modOtkupUI:1735` (`modeKey(ActiveMode) = "OTKUP"`, uslov
koji ne gleda aktivan ekran) i predložio zaseban PR. Operater je onda javio
podatak koji to obara: traka se vidi **samo** na Uvozu izvoda i na Dokumentima,
ne i na Agrohemiji ili Fakturisanju.

Pravi uzrok je jednostavniji: `ShowZones` gasi `zKpi`, `zCtx`, `zForm` i
`zRight` za ne-Dokumenta ekrane, a **`zOtp` je iz tog spiska ispao**. Njegovu
vidljivost postavlja samo `LayoutAll` (grana ekrana dokumenata), pa na ugovornim
ekranima ostaje onakav kakav ga je Dokumenta ostavila — a **vidi se ili ne vidi
zavisno od toga da li ga zona tog ekrana slučajno pokriva**. Agro i
Fakturisanje imaju više zone; Uvoz izvoda ima najnižu (104pt), pa je ostao
otkriven.

Ispravka je dopisivanje `zOtp` u već postojeći spisak — jedna reč. **Nije
pokrivena testom:** `ShowZones` je `Private`, a da bi se izvršila treba pokrenuti
prebacivanje ekrana kroz celu ljusku; test bi bio veći i krhkiji od same
ispravke. To se ovde beleži kao neizmereno, ne prećutkuje.

**U zaseban PR ide ono što je ostalo:** `ModeBrojiKomade(ActiveMode)` u podnožju
(ista klasa kao pogrešna dijagnoza gore — za poređenje, `ModeHasValCol()` i
`ModeHasKgCol()` su **ispravne**, jer se izvode iz `mCols`), i podnožje sa samo
jednim slotom za novčani zbir.
