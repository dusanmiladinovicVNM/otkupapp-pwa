# Katalog funkcija legacy formi i plan prelaska na novi UI

> Izvor istine za migraciju `frmOtkup` i `frmDokumenta` na `frmOtkupUI`
> (ljuska `modOtkupUI` + ekranski moduli `modScr*`).
>
> Cilj dokumenta: **ništa se ne gubi u prelasku.** Svaka sposobnost legacy
> forme je ovde popisana, sa oznakom da li je u novom UI-ju već obezbeđena,
> delimično obezbeđena ili nije. Plan na kraju radi samo po ovom spisku.

Stanje na dan `v6-ui-167`.

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
    **URAĐENO** (`v6-ui-167`). Ekran prijavljuje svoje čipove kroz `Scr_Cipovi`
    (`kljuc:KATALOG:sirina`), ljuska pozajmljuje slotove svog bazena i vraća
    ključ kroz `Scr_Rows`. Time je nestalo i poslednje mesto na kom je ljuska
    znala jedan ekran po imenu (`akt = "OTPREMNICE"`). Palete su dobile pet
    čipova: Sve · Ova godina · Otvorene · Zatvorene · Prerađene.

**Uz Fazu C je došlo i:** dvoklik na red se prosleđuje ekranu (`dbl:<red>`), pa
klik na paletu vodi pravo na njene stavke (`v6-ui-162`). Traženi **padajući redovi
detalja ispod izabranog reda** su odloženi sa razlogom: mrežu koriste svi ekrani, a
za detalj ispod reda ugovor bi morao da dobije **vrstu reda** i opis kolona po
vrsti, uz renderer promenljive visine — posao reda veličine migracije jednog
celog ekrana. Stoji kao prioritet za kasnije, jer znači na više mesta.

**Faza C je time ZATVORENA.**

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
15. Agrohemija, Fakture, Banka uvoz, Banka nalozi, Marža, Izveštaji,
    Sledljivost — svaki po istom obrascu.

---

## 6. Pravilo koje važi za sve faze

Ekran **nikad** ne računa i ne upisuje sam. Svaka stavka iz plana se rešava
pozivom postojeće rutine (`modCenovnik`, `modBrojevi`, `modOtkup`,
`modDokumenta`, `modPaletniList`, `modStorno`); ako rutina postoji ali je
`Private` i vezana za formu, prvo se **izdvaja račun** iz prikaza (kao
`KoopRangRows` iz `LoadKoopRang`), pa je koriste i legacy forma i novi ekran.
Duplirana logika se ne piše ni u jednom slučaju.
