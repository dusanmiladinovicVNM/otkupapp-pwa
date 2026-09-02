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

> **Pola je urađeno u `v2.84.0` — v. §19.** `RecalculateZbirnaFromOtpremnice_TX`
> sada nosi kapiju **u sebi** (broji vlasnike IKAD). Za
> `ReassignPrijemnicaToZbirna_TX` je izmereno da nijedan test ne razlikuje
> aktivne od IKAD na toj putanji, pa nije dirana — neizmerena izmena ponašanja
> je gora od otvorenog nalaza.
>
> „Umesto zaštite po call-site-u" se **ne izvršava** i to je namerno: kapije po
> call-site-u staju **pre transakcije** i kažu razlog, dok centralna staje
> iznutra i daje samo „nije uspelo". Centralna je **mreža ispod**, ne zamena.

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
18. ~~Banka nalozi~~ **URAĐENO** (`v6-ui-185`, `modScrBankaNalozi`) — v. §22.
19. ~~Izveštaji~~ **URAĐENO** (`v6-ui-186`, `modScrIzvestaji`) — v. §23.
    Ostaju: Marža, Sledljivost — svaki po istom obrascu.

### Faza F — Matični podaci

20. ~~**Matični podaci**~~ **URAĐENO** (`v6-ui-187` … `v6-ui-194`, §24.13–§24.19);
    `btnMatic` je u novom UI-ju bio nevezan. Nije stajalo u Fazi E jer nije „još jedan
    ekran po obrascu": donosi **sekciju ljuske** (drugi skup stavki sidebara)
    i prevezuje jedini operativan put upisa matičnih podataka. Redosled
    M0–M5 je u §24.10.

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

Ništa **novo** nije napravljeno za ovaj ekran: sve što mu treba postoji od Faze
C i Fakturisanja. Ali za razliku od §8.2, diff u `modOtkupUI` ovde **nije nula** —
smoke je našao četiri kvara u samoj ljusci, i oni su popravljeni **u ljusci**, ne
zaobiđeni u ekranu:

1. **Pečat verzije**, `OTKUI_BUILD` → `v6-ui-177`. Razlog nije „ekran Uvoz
   izvoda" nego isti kao u §8.10/R3: pečat postoji da bi se u smoke-u odmah
   videlo **da li je pravi kod uopšte uvezen**.
2. **`zOtp` dopisan u spisak zona koje pripadaju samo ekranu Dokumenta**
   (`ShowZones`).
3. **Geometrija kolona prati opis kolona** (`SetGridColsArr` →
   `OsveziGeometriju` na početku `RenderGrid`-a).
4. **`FmtDatumKratko` čuva i gornju granicu** datuma.

Uz njih, pravilo „broj koji nije datum" živi u `modUiData.CellDate`. Sve četiri
su opisane u §9.10; nijedna nije stvar ovog ekrana i sve pogađaju i ostale.

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

Testovi **104–112** u `modTest` i **trideset tri** sabotaže, uz nove fixture redove
u `tools/make_fixture.py`: **dvanaest** stavki izvoda u **pet** grupa
`(broj + račun + datum)`, tri otkupne stavke istog bloka i jedan broj bloka koji
postoji na **tri** otkupna mesta — od kojih jedno **nije upisano**.

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
| `BIM-FIX-PY` | **isti broj izvoda i isti račun, prethodni ciklus** — banke numeraciju ponavljaju po godini |
| `OTK-BIM-OMA` / `OTK-BIM-OMB` | **isti broj bloka na dva otkupna mesta** — broj otkupa je jedinstven po stanici |
| `OTK-BIM-OMX` | isti blok **bez upisanog `StanicaID`** — legacy oblik koji današnji pisci odbijaju, a zatečene sveske ga imaju |
| `OTK-BIM-PLAC` + `NOV-BIM-PLAC` | blok **u celosti plaćen** — lista ga i dalje nudi, a kandidata nema; bez ovog para ručni izbor takvog bloka ne bi imao nad čim da se izmeri |

| Test | Šta meri | Sabotaža |
|---|---|---|
| `T_BankaUvoz_UgovorEkrana` | registar, dve liste, granice bazena (`MAX_ACT` tačno 5 na stavkama, `MAX_CHIP`, `MAX_COLS`, `MaxPrekidaca`), prvi čip je najširi, izvodi bez radnji, **datum stiže kao broj** | `banka-uvoz-sesta-radnja`, `banka-uvoz-cip-sve-nije-prvi`, `banka-uvoz-izvodi-imaju-radnju`, `banka-uvoz-datum-nije-broj` |
| `T_BankaUvoz_IdentitetURedu_NeCrtaSe` | identitet u prenosnoj koloni prioriteta 4; **interne šifre nema među vidljivim kolonama**; dvosmislen ID → prazno, a red se i dalje vidi; kolizija broja izvoda | `banka-uvoz-identitet-vidljiv`, `banka-uvoz-dvosmislen-prvi-pobedjuje` |
| `T_BankaUvoz_RedNosiSmerIOtvorenost` | red **prenosi** smer i otvorenost umesto da ih izvodi iz prikaza; `"Error"` je i dalje otvoren | `banka-uvoz-red-ne-nosi-otvorenost`, `banka-uvoz-red-ne-nosi-smer`, `banka-uvoz-predlog-i-za-zatvorene` |
| `T_BankaUvoz_CipJakihPratiBrojac` | čip „jaki ključevi" i `CountStrongKeyReadyBankaImport` vide **isti** skup; „sve" je unija tri stanja; značka = čip „za obradu" = `GetBankaImportOpen`; **neuspeh čitanja zadržava poslednju poznatu brojku** | `banka-uvoz-cip-jaki-prolazi-sve`, `banka-uvoz-znacka-broji-mapirane`, `banka-uvoz-obradjeno-guta-preskoceno`, `banka-uvoz-kpi-greska-je-nula` |
| `T_BankaUvoz_IzvodiSuAgregatPoRacunu` | grupa je `(broj + račun + datum)`, mereno **direktno nad `BimIzvodKljuc`**; isti dan kao broj i kao `Date` je isti izvod; zbirovi se **uzimaju sa reda, ne sabiraju**; legacy red bez saldo podataka nije neslaganje | `banka-uvoz-izvod-kljuc-bez-racuna`, `banka-uvoz-izvod-kljuc-bez-datuma`, `banka-uvoz-saldo-se-sabira`, `banka-uvoz-legacy-red-je-razlika` |
| `T_BankaUvoz_RucnoMapiranjePravila` | smer-kapija se slaže sa writerom; prazan izbor bloka uzima poziv na broj; blok preko granice traži potvrdu; fail-closed nad listom faktura; **izabran blok nosi svoje otkupno mesto do writera**, a scope sužava kandidate u oba smera | `banka-uvoz-om-prima-nejasan-smer`, `banka-uvoz-prazan-blok-ostaje-prazan`, `banka-uvoz-fakture-fail-open`, `banka-uvoz-fakture-i-zatvorene`, `banka-uvoz-blok-bez-om-scope` |
| `T_ZonaBankaUvoz_PoljaIRaspored` | zona se STVARNO gradi i raspoređuje; sve kontrole postoje; kombo je polje (`nm` + `nmT`); polje cilja je ugašeno za OM | `banka-uvoz-om-polje-cilja-radi` |

Tvrdnja koja nosi najviše: **broj redova koje propušta čip „jaki ključevi" mora
biti identičan onome što vraća `CountStrongKeyReadyBankaImport`** — isti oblik
kao `T_Agro_AbzugMapaPratiPojedinacni` i `T_Fak_CipoviPrateStatusFakture`.
Pravilo živi na dva mesta (čitač mreže i natpis dugmeta) i može da se raziđe;
ovo je jedino što bi to primetilo.

**Čipovi, radnje i kolone se čitaju po KLJUČU liste** (`BuCipoviZaListu`,
`BuRadnjeZaListu`, `BuKoloneZaListu`), ne kroz `Scr_Lista` — isti razlog kao
§8.8: ugovor svake liste mora da se meri bez prebacivanja stanja ekrana.

Testovi **111** (`T_MrezaDatum_BrojKojiNijeDatum`) i **112**
(`T_MrezaGeometrija_PratiOpisKolona`) mere **ljusku**, ne ovaj ekran — nastali su
iz njegovog smoke-a, ali pravilo koje tvrde deli ceo UI.

**Dvosmerni dokaz je pušten za svih trideset tri**: svaka sabotaža obara
**tačno jedan** imenovani test i vraća se bit-identično. Bazna vrednost pre i posle je
`RunAllTests` **112 / 0**, a `RunBankaImportTestSuite` (tvrd fail-gate nad ovim
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

**Pravilo je LJUSKINO, ne ekranovo.** `modUiData.CellDate` je do sada svaki broj
propuštao kao „datum"; sada odbija sve što `CDate` ne sme da primi
(`DatumSerijskiValidan`, gornja granica 31.12.9999). `FmtDatumKratko` dobija isti
štitnik na samom mestu crtanja, jer tamo stiže i ono što nije prošlo kroz
`CellDate`. Ekran **ne drži svoju kopiju** tog pravila.

Vrednost se **ne tumači** — `ddmmyyyy` nije oblik koji `modParse.TryParseDateValue`
poznaje, pa bi tumačenje bilo izmišljanje pravila koje domen nema. Prazna ćelija
je istina; tuđi tekst nije.

**Nalaz je bio veći od ovog ekrana**, i to se videlo na fixture-u: takav red
posejan u `tblBankaImport` oborio je **sedam** testova sa `Overflow`, među njima i
`T_StornoEkran_SvakaListaVracaRedove`.

Prvo sejanje je pri tom promašilo metu: vrednost je upisana u ćeliju koja je
**nasledila datumski format** od reda iznad, pa je Excel pri čitanju pokušavao da
je vrati kao `Date` i obarao **celo** čitanje tabele (`GetTableData` → Overflow) —
grublji kvar od onog koji zatečene sveske imaju. Merenje na pravoj svesci
(`Diag_BuRedovi`) pokazuje `tip=Double`, dakle tamo ćelija ima običan format.
`make_fixture` zato dobija `Sirovo(...)`: vrednost koja se upisuje **bez**
nasleđenog formata. Red od tada **ostaje u fixture-u** — on je jedino što bi
povratak ove greške primetilo.

#### Treći smoke: mreža crta sa širinama PRETHODNE liste

Na listi IZVODI je kolona „OTVORENIH" (deseta vidljiva) bila **prazna u svakom
redu**. `FmtBroj(0, 0)` vraća `"0"`, ne prazno — dakle opet **preskočen upis**,
ne nula. Susedna kolona istog tipa (`num`) crtala se uredno.

Merenje (`Diag_BuRedovi`, proširen da ispisuje ceo red):

```
EKRAN red 1 kol10: tip=Long vred=[10]
MREZA red 1 kol10: tip=Long vred=[10]
```

Vrednost stiže do `mView` i `GridCell` je vraća — **samo se ne nacrta**. Time su
i čitač i ekran isključeni; ostaje crtanje.

**Uzrok je u ljusci.** `LayoutGrid` (koji puni `mColX` / `mColW`) zove se iz
**rasporeda** ekrana — `LayoutScreenZone`, odnosno `LayoutAll`. `ReloadGrid`
(promena liste, čipa, pretrage) zove samo `LoadGridFromScreen` + `RenderGrid`.
Posle promene liste `RenderGrid` zato crta sa `mColW` **prethodne** liste, a na
`mColW(k) = 0` radi `.Visible = False` i preskače upis.

Lista STAVKE ima 9 vidljivih kolona i tri skrivene (prioritet 4), pa je
`mColW(9) = 0`. Lista IZVODI je imala **10** vidljivih — i njena deseta kolona
je nasleđivala tu nulu. Zaglavlje je pri tom bilo vidljivo, jer ga osvežava
zaseban prolaz (`RefreshGridHeaders`) koji se pokreće kasnije, kad je `mColW`
već preračunat: otud najgori mogući izgled — **zaglavlje stoji, ćelije prazne**.

**Popravljeno u ljusci.** `SetGridColsArr` poredi **sadržaj** opisa kolona
(`ColsPotpis` — ekran vraća nov niz pri svakom čitanju, pa poređenje referenci ne
bi valjalo) i na promenu diže `mGeomStara`. `RenderGrid` na samom početku zove
`OsveziGeometriju`, koja preračuna raspored **samo ako** je opis stvarno drugi —
pa promena čipa ili strane ne plaća raspored. `LayoutGrid` na kraju briše
zastavicu, jer je upravo preračunao.

Mereno **na Fakturisanju**, ne na ovom ekranu: njegove liste su različite širine
(FAKTURE ima sedam vidljivih kolona, ZAFAKT devet), pa je prelazak sa uže na širu
baš onaj smer u kom su se kolone gubile. Time je i potvrđeno da nalaz nije bio
samo ovdašnji.

Broj otvorenih i broj stavki na listi izvoda ipak **ostaju spojeni** u jednu
kolonu (`„10 / 16"`, isti zapis koji traka iznad mreže već koristi za
`MAPIRANO 11 / 40`). To više nije zaobilazak nego izbor: dve susedne brojke bez
konteksta čitaju se gore od jedne sa kosom crtom.

#### Sabotaža koja je RASLA: zamena ne sme biti podniz sidra

Prva verzija sabotaže `mreza-geometrija-ne-prati-kolone` uklanjala je red
`mGeomStara = True` tako što je trolinijsko sidro menjala **prve dve** njegove
linije. `sabotaza.py --vrati` traži **zamenu** u fajlu i vraća sidro — a pošto je
zamena bila **podniz sidra**, nalazio ju je i u zdravom kodu. Svaki ciklus
apply→revert je zato dodavao još jedan primerak reda: posle proverâ ih je u
`modOtkupUI.bas` bilo **dvadeset četiri**.

To je zamka 8, opisana **u docstring-u same `sabotaza.py`** — i svejedno
pokupljena. Utoliko gore što je simptom nem: duplirano `mGeomStara = True` je
idempotentno, pa je kod radio, suite je bila zelena, a izvor je tiho rastao.

Ispravno: sabotaža gađa **jednu** liniju i zamenjuje je nečim **jedinstvenim**
(`mGeomStara = mGeomStara` uz oznaku), pa ni sidro ni zamena nisu podniz jedno
drugog. Dokazano tako što tri uzastopna ciklusa apply→revert ostavljaju fajl
bit-identičnim.

Uz to: prva verzija **nije obarala ništa** (21/22), i to je bio jedini znak da
nešto nije u redu. Sabotaža koja ne obara svoj test uvek je nalaz — ovog puta o
sebi samoj.

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

#### Code review: tri nalaza koja su ušla PRE merge-a

Review je dao `REQUEST CHANGES` sa tri nalaza o kojima suite nije imala šta da
kaže — ne zato što su tvrdnje bile slabe, nego zato što **fixture nije imao
podatak nad kojim bi se videli**. Sva tri su zatvorena u istoj grani.

**P1 — ručni kooperantski blok nije nosio otkupno mesto.** Broj otkupa je
jedinstven **po stanici**, pa isti broj bloka legitimno pripada dvama različitim
blokovima. `GetOtkupCandidatesForKooperantBlock` je filtrirao samo po
`KooperantID + BrojDokumenta`, pa bi u jednu raspodelu ušle stavke sa **oba**
otkupna mesta — novac na dva različita poslovna lanca, bez ijedne poruke.

Scope ide kroz ceo lanac kao **opcioni** argument (`stanicaID` →
`MapBankaImportAsKooperantBlockCore` / `Manual` / `_TX`, `BimBlokTraziPotvrdu`),
pa automatsko mapiranje — koje otkupno mesto nema odakle da zna — ostaje
nepromenjeno. `GetBlokoviZaBimMapiranje` zato više ne vraća niz brojeva nego
tabelu `BrojBloka | StanicaID | prikaz`, a kombo cilja dobija **treću, skrivenu**
kolonu. Prikaz (`12 · OM Naziv`) se **ne parsira** — isto pravilo kao identitet u
redu (§9.6): ono što čovek čita sme da se menja, podatak ne.

Kad blok dolazi iz **poziva na broj** (prazan izbor u kombou), scope-a nema i
ponašanje ostaje kao kod automatskog mapiranja. To je zaseban slučaj i meri se
zasebnom tvrdnjom (`Scr_BuScopeBlokaTest`).

**P2 — identitet izvoda nije nosio datum.** `BimIzvodKljuc` je bio
`(broj + račun)`. Banke numeraciju izvoda ponavljaju po ciklusu, pa izvod 15 na
istom računu postoji i 2025. i 2026 — i spajali bi se u jedan red, i to na
najgori način: saldo i datum sa **prvog** reda, a broj stavki **sabran preko
oba**. Sintetički izvod koji nikad nije postojao, i to na jedinom mestu gde se
vidi da li se izvod slaže.

Datum se u ključu normalizuje u serijski broj kad god može (`IzvodDatumKljuc`),
da isti dan zapisan kao `Date` i kao broj ne bi dao dve grupe; neispravna
vrednost ide kao tekst — ne sme da **spoji** dva izvoda, nego da ostane svoja.

**P2 — pad čitanja je postajao legitimna nula.** `Kpi()` je na grešku vraćao
`Array(0, 0, 0, 0#, 0#)`. Značka uz stavku menija odgovara na pitanje „ima li
finansijskih stavki koje čekaju čoveka" — pa je operater dobijao **„nema posla"
umesto „ne znam"**, i to baš kad je nešto sa šemom ili kesom pošlo naopako. Isti
fail-open je jednom već plaćen u Stornu.

Greška se sada **loguje** (`LogErr`), kes se **ne** proglašava važećim (sledeći
poziv pokušava ponovo), a vraća se **poslednja poznata** vrednost. Nula ide samo
dok validne vrednosti još nije ni bilo — tada ni značke nema, pa nema ni čega
lažnog. `Scr_ResetCache` zato više ne briše `mKpi`, samo ga proglašava
zastarelim.

**Zamka pri pisanju dokaza.** Prve dve sabotaže ključa izvoda obarale su
**istu** tvrdnju („isti broj daje tri reda"), jer se pravilo merilo samo preko
broja redova u mreži — zamka 5 iz `sabotaza.py`. Test zato sada meri
`BimIzvodKljuc` **direktno**, jednom tvrdnjom po polovini ključa; agregat ispod
ostaje kao integracioni dokaz. Sa tim, svaka od četiri nove sabotaže obara tačno
svoju imenovanu tvrdnju.

#### Drugi krug review-a: tri stanja koja su izgledala isto

Prvi krug je zatvorio „ručni blok ne nosi otkupno mesto". Drugi je pokazao da
**samo uvođenje scope-a nije dovoljno** — jer prazan `stanicaID` do sada znači
tri različite stvari, a izgleda kao jedna:

| Stanje | Šta znači | Šta sme |
|---|---|---|
| scope nije ni tražen | blok dolazi iz **poziva na broj** — automatsko mapiranje | prolazi bez scope-a, kao i pre |
| scope tražen, red nema stanicu | legacy/uvezen red bez `StanicaID` | **STOP** |
| scope tražen, kolone nema | schema drift | **greška**, nikad tihi nastavak |

**Drugo stanje je opasno baš zato što liči na prvo.** Operater je *birao* blok iz
liste; da je prazan scope prošao, writer bi raspodelio novac preko **svih**
otkupnih mesta sa tim brojem. `GetBlokoviZaBimMapiranje` takav red uredno nudi
(ključ je `broj & "|" & sta`, prazan `sta` prolazi), a prikaz mu je pao nazad na
goli broj — pa bi operater u listi video „12" i „12 · OM B" bez ijednog traga
zašto se razlikuju.

Red **ostaje u listi** — postoji u podacima i prećutati ga značilo bi lagati o
tome šta je u tabeli — ali se **označava** (`bez otkupnog mesta`), a radnja nad
njim staje uz objašnjenje. Pravilo je jedna funkcija (`BuScopeNedostaje`), pa
oznaka i kapija ne mogu da se raziđu.

**Test seam je pri tom morao da prestane da PONAVLJA izraz.** Prva verzija
`Scr_BuStopBezOmTest`-a je sama računala `BuScopeNedostaje(IzabraniCiljID(),
IzabranaStanicaCilja())` — isti izraz koji stoji u `RucnoKooperant`. Dok je tako,
sabotaža obara **kopiju u testu**, a radnja bi i dalje slala prazan scope; test
bi bio zelen nad kodom koji je pokvaren. Sada oboje zovu `ScopeIzbora`, koji
scope i odluku računa na jednom mestu. Ono što ostaje neizmereno je **jedan red**
vezivanja u `RucnoKooperant` (`If stani Then ... Exit Function`) — provereno
čitanjem, i tako se i beleži.

**Treće stanje je najtiši mogući kvar.** Filtar je glasio
`If Len(Trim$(stanicaID)) > 0 And colSta > 0` — dakle scope je zadat, kolona se
zbog drifta ne nađe, uslov otpadne, i resolver vrati kandidate sa **svih**
otkupnih mesta. Bez ijedne poruke, u listi koja izgleda savršeno ispravno. Sada
je pravilo u `BimScopeKolona`: zadat scope traži kolonu kroz `RequireColumnIndex`
i puca ako je nema; bez scope-a kolona ostaje opciona kao i pre.

**Ime kolone je argument te funkcije, a ne konstanta u telu** — to je jedini
način da se grana „kolone nema" izmeri a da se ne razbije šema fixture-a. Test
zove `BimScopeKolona(STA, "NemaOvakveKoloneUOtkupu")` i tvrdi da **puca**, pa
istu nedokazivu kolonu bez scope-a i tvrdi da **prolazi**.

#### Značka koja ume da kaže „ne znam"

Prvi krug je popravio *drugi* pad čitanja (zadrži poslednju poznatu brojku), ali
ne i **prvi u sesiji**: tada poslednje vrednosti nema, pa je vraćana nula.
Obrazloženje je bilo „tada ni značke nema, pa nema ni čega lažnog" — i bilo je
naopako. `BrojacTekst` za `n <= 0` vraća prazno, a **prazna značka u ovom UI-ju
nije odsustvo poruke nego poruka**, i glasi „nema šta da čeka". Prvi pad je zato
i dalje bio fail-open, samo tiši.

Ugovor `Scr_Brojac() As Long` nema treći kanal, a **brojač ne može legitimno da
bude negativan** — pa je to slobodan kanal koji ne traži izmenu ugovora:
`BuKpiNepoznato` vraća `-1`, ljuska ga crta kao `!`, a četiri KPI pločice ekrana
pokazuju crtu umesto brojki. Pola tačnih brojki uz dve nule bilo bi gore od
iskrenog „nemam podatak".

Izmena je **u ljusci** (`BrojacTekst`), pa važi za svaki ekran koji je poželi —
kao i geometrija mreže i granica datuma pre nje.

#### Šta je iz review-a svesno ODLOŽENO, i zašto

- ~~**Invarijanta metapodataka izvoda.**~~ **URAĐENO** — v. §12.
- **`RenderGrid` stale-safe.** Popravka datuma je zatvorila **jedan ulaz**, ne
  klasu: `CDbl` u `kg`/`num`/`rsd`/`mult` i `CLng` u `pill`/`paypill` nad
  neispravnom vrednošću imaju identičan ishod — greška se proguta, ćelija zadrži
  natpis prethodnog ekrana. Ispravno je prazniti ćeliju **pre** upisa, pa je pad
  rendera prazna ćelija i jedan log. **Zaseban PR**, i vredniji od svega
  odloženog jer zatvara klasu.
- **Checker tri registra testova** (`RunOne` / `TestName` / `InvokeTest`).
  Trenutno su usklađeni — 112/112/112, bez rupa i bez razlike — pa je ovo
  preventiva, ne živ bug. Uzak checker sa svojim `--self-test` slučajevima,
  **zaseban PR.**

#### Treći krug: ista greška, jedan nivo iznad

Prva dva kruga su zatvorila scope **unutar** ručnog mapiranja. Treći je pokazao
da je ista klasa ostala **iznad** njega — u trenutku kad se lista uopšte puni.

`PuniCiljCombo` puni obe liste cilja i ima jedan `On Error GoTo EH`, koji je
postavljao zastavicu `mFaktureOK = False`. Ime nije bilo slučajno: kapiju je
čitao **samo** `RucnoKupac`. `RucnoKooperant` je odmah išao na
`BimEfektivniBlok`, pa je pad učitavanja blokova izgledao ovako:

```
GetBlokoviZaBimMapiranje  ->  greška (schema drift, nedostupna tabela...)
PuniCiljCombo EH          ->  zastavica postavljena, ali je niko ne čita
prazan combo              ->  "operater nije birao blok"
BimEfektivniBlok          ->  fallback na PozivNaBroj
ScopeIzbora               ->  scope = "", stani = False
GetOtkupCandidates        ->  BEZ scope-a
```

To je tačno ono što je `c8b7a32b` zabranio, samo dosegnuto drugim putem.

**A ako iz poziva na broj ne ispadne nijedan kandidat, ishod je gori od
pogrešne raspodele.** `MapBankaImportAsKooperantBlockCore` na
`If IsEmpty(kandidati)` **ne prijavljuje grešku** — ceo iznos knjiži kao
`NOV_VIRMAN_AVANS_KOOP` i stavku označava obrađenom
(`UpdateBankaImportStatus ... "Da"`). Neuspeh čitanja tako postaje **uspešno
knjiženje drugog poslovnog ishoda**.

Kapija je zato sada **zajednička** (`CiljUcitan`), a zastavica se zove `mCiljOK`
— jer prazna lista na **obe** rute nosi poslovno značenje: prazan izbor fakture
je avans, prazan izbor bloka je poziv na broj.

#### Prazna tabela i nepostojeća tabela nisu isti ishod

`GetTableData` vraća `Empty` za oba. Čitač koji gleda samo `IsEmpty(data)` zato
nedostajuću tabelu tumači kao „nema redova" — a tamo gde prazna lista nosi
poslovno značenje to je fail-open. `RequireColumnIndex` to **ne pokriva**: do
provere kolona se ne bi ni stiglo, jer čitač izađe ranije.

Novi `modSchemaGuard.RequireTable` stoji **pre** `GetTableData` u
`GetFaktureZaBimMapiranje` i `GetBlokoviZaBimMapiranje`. Isti odnos kao
`TabelaCitljiva` u `modUiData`, samo na domenskoj strani i kao greška umesto
zastavice.

#### Sabotaža koja nije oborila ništa — i šta je tražila

`banka-uvoz-cilj-kapija-ne-puni` uklanja poziv `PuniCiljCombo` iz kapije. Prvi
put je vratila **112 / 0**.

Uzrok nije bio u sabotaži nego u tome da se pravilo **ne može videti**: bez forme
`PuniCiljCombo` izađe na `If c Is Nothing Then Exit Sub`, pa uklonjen poziv ne
menja nijedan merljiv ishod. A baš taj nedostajući poziv je bio kvar — kapija bi
sudila po zastavici **prethodnog** izbora.

Zato modul dobija `mCiljPunjenja`: broji se **poziv**, ne uspeh. To je jedina
stvar koju je posle uklonjenog poziva moguće opaziti bez forme, i sada sabotaža
obara svoju tvrdnju. Isti obrazac kao test 110 (zona se stvarno gradi) — kad se
pravilo ne vidi, ne izmišlja se tvrdnja koja prolazi, nego se napravi vidljivim.

#### Zastareli komentari

Dva komentara su zaostala iza koda i oba bi za tri meseca proizvela pogrešan
mentalni model: zaglavlje modula je i dalje tvrdilo da su IZVODI agregat po
`(broj + račun)` bez datuma, a `IzvodiKolone` je opisivao spojenu kolonu kao
privremeni zaobilazak „dok se to ne popravi u ljusci" — a popravljeno je **u
ovom istom PR-u**. Spojena kolona ostaje, ali sada kao **izbor**: dve susedne
brojke bez konteksta čitaju se gore od jedne sa kosom crtom.

#### Četvrti krug: neuspeh se ne pamti

Dve grane punjenja liste cilja **greške javljaju različito**, i to je pravilo
jedne od njih prećutno prekršilo:

| Grana | Kako javlja pad | Gde završi |
|---|---|---|
| blokovi (`GetBlokoviZaBimMapiranje`) | **diže** grešku | `EH` → `mCiljPunjen = ""` |
| fakture (`GetFaktureZaBimMapiranje`) | vraća **zastavicu** `outOK = False` | procedura mirno stigne do kraja |

Za fakture je punjenje zato stizalo do `mCiljPunjen = kljuc` i **keširalo
neuspeh**. Kapija je radnju tačno blokirala — fail-closed je držao, nijedno
pogrešno knjiženje nije bilo moguće — ali sledeći klik na *isti* izbor izlazio je
odmah na `If mCiljPunjen = kljuc Then Exit Sub` i **nije ni pokušavao ponovo**.
Prolazan kvar bi tako zaključao izbor do `ResetCache`: fail-closed koji izgleda
kao pokvaren ekran.

Kešira se sada samo uspešno čitanje (`CiljKesKljuc`) — isto pravilo koje `mKpiOK`
u istom modulu već ima. Sabotaža: `banka-uvoz-kes-pamti-neuspeh`.

#### Šta je i posle četiri kruga OSTALO neizmereno — i jedno nezatvoreno

Da se ne pročita kao „sve je pokriveno":

- **Vezivanje kapije u dve rute.** `RucnoKupac` i `RucnoKooperant` zovu
  `CiljUcitan` sa po jednim redom. Test dokazuje *pravilo* i da kapija zove
  punjenje (`mCiljPunjenja`), ali ne i da ta dva reda postoje — to je provereno
  čitanjem.
- **Samo punjenje kombo-a nikad se ne izvršava headless.** Bez kontrole
  `PuniCiljCombo` izađe odmah, pa tri kolone, skrivena kolona scope-a i oznaka
  „bez otkupnog mesta" postoje samo u kodu i u smoke-u. Test 110 dokazuje da se
  zona **gradi i raspoređuje**, ne i da se lista puni.
- **`MapBankaImportAsKooperantBlockCore` je i dalje fail-open — u writeru.** Na
  `IsEmpty(kandidati)` knjiži ceo iznos kao avans kooperanta i stavku označava
  obrađenom. Ovaj PR je zatvorio **putanje kojima se do tog stanja pogrešno
  dolazilo sa ovog ekrana**, ali to je ublažavanje, ne ispravka: `frmBankaImport`
  i automatsko mapiranje i dalje ulaze u istu granu.

  Za **automatsko** mapiranje je avans namerno i dokumentovano ponašanje
  („bezbedan izlaz" dok je poreklo dvosmisleno). Za **ručno izabran blok** to je
  nešto drugo: operater je rekao *koji* blok, pa „nijedna otkupna stavka" nije
  bezbedan ishod nego protivrečnost. Writer to danas **ne može da razlikuje** —
  `MapBankaImportAsKooperantBlockManual` samo prosleđuje argumente u `Core` bez
  ijedne oznake da je poziv ručan. Razdvajanje toga je **zaseban PR**, i dira
  legacy formu.

#### Peti krug: ručno izabran plaćen blok tiho postaje avans

Ovo je bilo u prethodnom krugu zapisano kao „writer je fail-open, zaseban PR".
Prigovor je bio da **novi ekran uvodi novu dostupnu putanju do tog writera**, pa
mora sam da je zatvori — i taj prigovor je tačan. Legacy dug ne opravdava novu
rutu do njega.

Putanja je u celosti dostupna sa ovog ekrana:

```
GetBlokoviZaBimMapiranje   ->  nudi SVAKI nestorniran broj otkupa
                               (ne proverava da li blok jos duguje)
operater bira blok 125     ->  koji je u celosti placen
GetOtkupCandidates...      ->  bira samo "otvoreno > 0.009"  ->  Empty
MapBanka...BlockCore       ->  IsEmpty(kandidati)
                               -> SaveNovac NOV_VIRMAN_AVANS_KOOP
                               -> UpdateBankaImportStatus "Da"
```

Operater je rekao **koji dug plaća**; sistem je knjižio **avans** i stavku
označio obrađenom, bez pitanja. Transakcija uspeva, a semantika je druga od
izabrane.

**Razlika koja pravilo čini tačnim je ručno vs. automatsko.** Za automatsko
mapiranje avans **jeste** namerno i dokumentovano ponašanje — bezbedan izlaz dok
je poreklo dvosmisleno. Za izričito izabran blok „nema šta da se plati" nije
bezbedan ishod nego **protivrečnost**. Writer to danas ne može da razlikuje
(`MapBankaImportAsKooperantBlockManual` samo prosleđuje argumente u `Core`), pa
odluku donosi pozivalac koji **zna** da je izbor bio ručan:
`BimBlokBezOtvorenih` je domensko pravilo, `BuBlokZatvoren` je ekranska kapija
koja se primenjuje **samo** kad `izabranBlok` nije prazan.

Blok **ostaje u listi** — postoji u podacima — po istom pravilu kao blok bez
otkupnog mesta.

Dve sabotaže, jer su dva pravila: `banka-uvoz-placen-blok-postaje-avans` (kapija
uopšte ne vidi da je blok zatvoren) i `banka-uvoz-kapija-bloka-i-za-poziv`
(kapija se proširi i na poziv na broj i ugasi legitimnu granu). Druga postoji
zato što je „popravka" koja gasi namerno ponašanje takođe regresija.

**Writer i dalje ostaje fail-open za svoje ostale pozivaoce** — `frmBankaImport`
i automatsko mapiranje ulaze u istu granu. Razdvajanje ručnog od automatskog
*unutar* writera je zaseban PR i dira legacy formu; ovaj PR zatvara samo putanju
koju sam uvodi.


---


#### Writer je zatvoren i za svoje ostale pozivaoce (`v6-ui-179`)

PR ekrana je ovo ostavio zapisano kao **ublažavanje, ne ispravku**: zatvorene su
putanje kojima se do fail-open grane dolazilo *sa novog ekrana*, ali
`frmBankaImport` i automatsko mapiranje su i dalje ulazili u istu granu.

`MapBankaImportAsKooperantBlockCore` na `IsEmpty(kandidati)` knjiži ceo iznos kao
`NOV_VIRMAN_AVANS_KOOP` i radi `UpdateBankaImportStatus … "Da"`. Za **automatsko**
mapiranje je to namerno i dokumentovano — bezbedan izlaz dok je poreklo
dvosmisleno. Za **izabran** blok nije.

Writer to nije mogao da razlikuje: `MapBankaImportAsKooperantBlockManual` je samo
prosleđivao argumente. Sada nosi `blokIzabran`, a `Core` na
`IsEmpty(kandidati) And blokIzabran` diže `ERR_BMAP_BLOK_PRAZAN` umesto da knjiži.

**Isti uslov postoji i na ekranu** (`BuBlokZatvoren`), i to nije duplikat nego
pravilo iz `.claude/rules/testovi.md` §5: modul unosa sudi po **snimku iz
trenutka kad je lista punjena**, a writer po stanju **u trenutku upisa** — i
legacy forma ovamo ulazi bez ijedne UI provere. Isti obrazac kao
`ApplyAvansToOtkup` i `UplataFakturaProblem`.

`frmBankaImport` dobija jedan `Private Function ManualBlokIzabran()` koji čita
**iste kontrole** iz kojih već čita `EffectiveManualBlockNo`, i prosleđuje ga.
Razlika koju forma već pravi samo je dobila ime.

**Tvrdnja živi u `RunBankaImportTestSuite`, ne u `RunAllTests`** — writer piše, a
`RunAllTests` je nemutirajuća suite (mutirajući test tamo obara CI kapiju
`who_writes`). `T21_IzabranPlacenBlokNijeAvans` seje sopstveni blok plaćen do
nule i meri **oba** ishoda nad istim podacima, razlikujući ih samo poslednjim
argumentom:

| Ulaz | Ishod |
|---|---|
| `blokIzabran = True` | ništa se ne knjiži, `tblNovac` prazan, stavka **ostaje otvorena** |
| `blokIzabran = False` (poziv na broj) | avans, stavka zatvorena, celim iznosom |

Druga polovina postoji zato što bi se pravilo moglo „ispraviti" gašenjem
legitimne grane. Sabotaža `banka-writer-placen-blok-je-avans` obara **tri**
provere tog testa.

**Ali sama zaštita u writeru nije bila dovoljna za legacy formu.** Writer radi
sa informacijom koju dobije, a `frmBankaImport` ju je davao pogrešno u jednom
slučaju: njen učitavač blokova nije razlikovao „kooperant nema blokova" od
„nisam uspeo da pročitam blokove". Prazan kombo → `ManualBlokIzabran = False` →
poziv na broj → prazan skup kandidata → **avans**, uz stavku označenu obrađenom.

Ista klasa, i forma je za nju **već imala rešenje jednu funkciju dalje**:
`m_FaktureLoadOk` / `m_FaktureLoadErr` postoje od ranije, ali samo za fakture.
Blokovi su ih sada dobili, sa istim `EH` obrascem i istom kapijom pre knjiženja.

Uz to su **oba** učitavača dobila `RequireTable`: i za fakture je nedostajuća
tabela do sada prolazila kao „nema faktura" (zastavica bi ostala `True`), što je
isti fail-open samo na drugoj listi.

**A pravu granicu drži domen.** `GetOtkupCandidatesForKooperantBlock` je i sam
na `IsEmpty(data)` izlazio praznim skupom, pa bi nedostupna `tblOtkup` završila
kao avans **za svakog pozivaoca**, uključujući automatsko mapiranje koje UI
proveru i nema. Sada i on ide kroz `RequireTable`.

Greška se pri tom **propagira**, ne pretvara u `Error` na tom redu:
`AutoMapBankaImportRowBatch` guta **samo** „ovaj red mora ručno"
(`IsManualRequiredBankaError` pokriva četiri broja), pa nedostupna tabela obara i
rollback-uje **ceo** batch. Tako i treba — nedostupna `tblOtkup` nije svojstvo
jednog bankarskog reda nego kvar instalacije.

| Sloj | Šta hvata |
|---|---|
| forma / ekran | pad učitavanja liste → **STOP**, sa objašnjenjem |
| domen (`GetOtkupCandidates…`) | nedostupna tabela → **greška**, za sve pozivaoce |
| writer (`blokIzabran`) | izabran blok bez otvorenih stavki → **greška** |

**To je bilo zapisano kao neizmereno — i onda izmereno.** Kapija u
`frmBankaImport` i `ManualBlokIzabran()` sada imaju test; v. §11.

> Seed je pri tom morao da bude prava uplata, ne status kolona:
> `SeedOtkupIsplacen` piše `Isplaceno`, a resolver računa
> `vrednost − GetUplataForOtkup(OtkupID)` — dakle sabira `Isplata` iz `tblNovac`.
> Prvi pokušaj je zato pao na sopstvenom preduslovu, što je i bila poenta
> preduslova.


---

## 10. Mreža: ćelija koja ne ume da se prikaže (`v6-ui-178`)

Ovo nije ekran nego **ljuska**, i nastalo je iz nalaza Banka uvoza (§9.10): datum
oblika `26062026` je u ćeliji ostavljao **tuđi tekst** — natpis sa prethodnog
ekrana. Tada je zatvoren taj jedan ulaz. Ovo zatvara **klasu**.

### 10.1 Zašto je preskočen upis, a ne samo račun

`RenderGrid` radi pod `On Error Resume Next`, i to je namerno: pad jedne ćelije
ne sme da obori crtanje cele mreže. Ali tekst se računao **u samom upisu**:

```vb
.caption = FmtBroj(CDbl(mView(r, k + 1)), 0)
```

Kad `CDbl` pukne, ne preskače se samo račun nego **i dodela**. U kontroli ostaje
natpis od ranijeg crtanja — vrednost sa prethodnog ekrana, bez greške i bez
traga.

Datum je bio samo prvi ulaz. Isti ishod daju `CDbl` nad tekstom u
`kg`/`num`/`sum0`/`rsd`/`mult` i `CLng` u `pill`/`paypill` kolonama.

### 10.2 Šta je promenjeno

Konverzija je izvučena u **funkciju koja ne može da pukne** (`CelijaTekst`,
`CelijaBroj`), a upis se radi **uvek**. Neuspeh je **prazna** ćelija — prazno je
istina, tuđi tekst nije.

| Pravilo | Zašto baš tako |
|---|---|
| `ok` je zasebna zastavica, ne „prazan rezultat" | Kolona `rest` namerno ne crta `0,00`; prazno je tu istina o podatku. Da se to broji kao kvar, log bi bio pun poruka o urednim redovima i prestao bi da se čita. |
| Pilula na neuspeh **ne dobija nulu** | Nula je kod `PaintPill` „Sačuvana", a kod `PaintPayPill` „Neplaćeno" — dakle **određen status**. To je svoja vrsta laži; ćelija se zato prazni. |
| Prava pilula (`pill`) na neuspeh se **briše cela** | `PaintPill` menja i `BackColor`, `BackStyle` i `width`, a `PaintRow` pri vraćanju pozadine reda `pill` kolone **namerno preskače**. Ćelija kojoj je obrisan samo natpis ostala bi kao **prazna obojena kutija**. |
| `paypill` na neuspeh gubi **samo natpis** | v. §10.4 — dve vrste, dva ugovora |
| Stil ostalih ćelija se **ne dira** | v. §10.3 |
| Broji se i prijavljuje **jednom po crtanju** | Prazna ćelija je istina, ali **tiha** prazna ćelija je bila pola problema: prvi nalaz ove vrste tražio je zasebnu dijagnostiku (`Diag_BuRedovi`) da bi se uopšte video. |

### 10.3 Reset stila je bio pogrešan lek — i sam je bio regresija

Prva verzija ove ispravke je pred svaki upis „vraćala ćeliju u neutralno":
`Font.bold = False` i `TextAlign = fmTextAlignLeft`. Namera je bila da bold sa
pilule ne procuri u običnu kolonu.

**To je obaralo ono što je `LayoutGrid` upravo postavio**, i to na svakom ekranu:
`LayoutGrid` brojčane kolone poravnava **desno**, a `StyleGridCell` prvoj koloni
i kolonama novca daje **bold**. Render koji to resetuje pomera sve brojeve levo i
skida bold — a nijedna tvrdnja o **natpisu** to ne primećuje, pa suite ostaje
zelena. Nađeno je u review-u, ne u testu.

Ispravno je da render stil **uopšte ne dira**. Procurivanje sa pilule rešava
mehanizam koji već postoji: promena liste menja opis kolona → `mGeomStara` →
`LayoutGrid` ponovo prođe kroz `StyleGridCell` za svaku ćeliju. Pilula i običan
broj nikad nisu ista kolona *u istoj listi*, pa drugog puta nema.

Ostaje samo slučaj u kom **ista** pilula ne može da se naslika — i to je jedino
mesto koje render čisti (`OcistiPilulu`).

Uz to: `Font.bold = False` je mehanizam koji je `v6-ui-175` već izmerio kao
**nepouzdan** u MSForms — zato postoji `modUiKit.PostaviRez`, koji proverava
`Font.Weight` i ponavlja upis. Rešenje u kom render bold uopšte ne dira to
pitanje zaobilazi u celosti.

### 10.4 Dve vrste pilule su dva ugovora

Prva verzija ispravke je obe vrste čistila istim helperom. To je za `paypill`
pogrešno, i to je greška napravljena **baš u PR-u koji tu klasu zatvara**:

| | ko drži širinu | ko drži pozadinu |
|---|---|---|
| `pill` | `PaintPill` (računa je po tekstu, `PillW`); `LayoutGrid` je **preskače** | `PaintPill`; `PaintRow` je **preskače** |
| `paypill` | `LayoutGrid` (`mColW - 16`) | `PaintRow` |

`PaintPayPill` menja **samo** natpis, boju teksta, poravnanje i bold. Čišćenje
koje bi `paypill` tretiralo kao pravu pilulu postavilo bi joj **punu** širinu
kolone — i ona bi takva ostala, jer `PaintPayPill` širinu ne vraća, a
`LayoutGrid` se ponovo pušta tek kad se promeni opis kolona:

```
paypill validan      -> width = mColW - 16
sledeci render, kvar -> "ciscenje" postavi width = mColW
sledeci render, opet
validna vrednost     -> PaintPayPill sirinu NE vraca  ->  celija ostaje 16pt sira
```

Zato render `paypill` koloni briše **samo natpis**.

**Ovo je izmereno round-trip-om**, ne samo „valid → invalid": zaostalo stanje se
vidi tek kad se posle kvara opet crta uredan podatak. Sabotaža
`mreza-paypill-kao-pill` vraća staro ponašanje i obara test sa
`očekivano [76], dobijeno [92]` — tačno 16pt.

### 10.5 Nalaz koji je brojač odmah izbacio

Čim je crtanje počelo da broji ćelije koje ne ume da prikaže, test nad **urednim
fixture podacima** prijavio je dva kvara:

```
date/1 tip=Date vred=[15.3.2026.]
```

**`IsNumeric` nad `Date`-om je `False`**, a `FmtDatumKratko` je počinjao baš tom
proverom. Ekran koji vrednost preda onakvu kakva u tabeli jeste — dakle kao
`Date` — dobijao je **praznu** ćeliju.

Banka uvoz je to zaobišao tako što svoj datum konvertuje u serijski broj
(`modUiData.CellDate`). **Lista FAKTURA to nije radila: njena kolona DATUM je
bila prazna u svakom redu.** Niko to nije prijavio, a nijedan test nije mogao da
vidi — do sada nijedan nije čitao *nacrtan* datum.

`FmtDatumKratko` sada prima `Date` direktno; serijski broj ide istim putem, sa
istom gornjom granicom.

### 10.6 Verifikacija

Test **113** (`T_MrezaCelija_NeostavljaTudjiTekst`) meri **oba nivoa**:

- **pravilo**, bez forme — `CelijaTekst` / `CelijaBroj` nad neispravnim
  vrednostima, uz razliku „prazno zbog nule" (nije kvar) i „prazno zbog
  neuspeha" (jeste);
- **crtanje**, nad pravom formom — `GridRenderTest` pušta `LayoutGrid` i
  `RenderGrid` koje ljuska i inače zove, pa se čita **caption same kontrole**.

To je bilo potrebno jer je pravilo tačno i kad se upis preskoči — a ceo kvar je
bio baš u tome. Kolona nad kojom se meri se **ne pretpostavlja** nego se traži
(`GridKindKoloneTest`): nad `txt` kolonom svaka vrednost prolazi, pa bi tvrdnja
tamo merila ništa.

Vrednost koja se ne može prikazati ubacuje se **posle učitavanja**
(`GridTestVrednost`), ne u tabelu: takav red je jednom već oborio **sedam** tuđih
testova sa `Overflow`, a izmišljen niz bi merio izmišljotinu umesto puta kojim
ide operater.

Test tvrdi i **stil pre i posle crtanja** — da render ne pokvari poravnanje i
bold koje je layout postavio. Bez toga je regresija iz §10.3 prolazila zeleno.

Šest sabotaža: `mreza-celija-prazno-ne-prepisuje` (vraća tačno stari oblik — prazan
rezultat ne prepisuje staru vrednost), `mreza-crtanje-kvari-stil` (vraća reset iz
§10.3; obori se sa `očekivano [3], dobijeno [1]` — desno vs. levo),
`mreza-pilula-ostaje`, `mreza-paypill-kao-pill`, `mreza-datum-nije-date`,
`mreza-kvar-celije-se-ne-broji`.

**Bilo je neizmereno:** čišćenje **pozadine** prave `pill` ćelije. Tada je ovde
pisalo da se lista Dokumenata „bez izabranog režima ne puni", pa da se ta ćelija
ne može doseći bez forme.

**Taj zaključak je bio netačan — v. §15.** Lista se puni i bez forme; rupa je
zatvorena testom 118.

**Širina jeste izmerena** (§10.4) — ona se na `paypill` koloni vidi.


---

## 11. Legacy forme dobijaju test seam-ove (`v6-ui-180`)

Tri uzastopna PR-a našla su **istu klasu greške** — „prazna lista je protumačena
kao izbor" — i svaki put ju je našao **review, ne suite**. Razlog nije bio u
tvrdnjama nego u tome gde pravila žive: u `Private` stanju forme i u click
handleru, dakle proverljivo samo rukom.

Ovo to menja za `frmBankaImport`.

### 11.1 Zašto je uopšte izvodljivo

`frmBankaImport` **nema `UserForm_Initialize`** — ima samo `UserForm_Activate`,
koji ide tek na `.Show`. `New frmBankaImport` je zato jeftin i **ne čita nijednu
tabelu**: test dobija formu, a težak posao se nikad ne pokrene.

To je razlika u odnosu na `frmOtkupUI`, gde `UserForm_Initialize` gradi ceo UI
(zato ga `NewOtkupUIForm` namerno okida čitanjem `.Controls.Count`).

### 11.2 Pravilo izlazi iz handlera

Kapija je stajala inline u `btnSacuvajRucno_Click`. Izdvojena je u
`KooperantRucnoSme(ByRef outPoruka)`, koji **vraća** poruku umesto da je prikaže
— time funkcija ostaje bez dijaloga i postaje pozivljiva iz testa. Handler zove
istu funkciju, pa se odluka i njena provera ne mogu razići.

Isti obrazac kao `ScopeIzbora` i `CiljUcitan` na novom ekranu.

### 11.3 Seam-ovi

Tvrdo gejtovani (`If Not IsTestMode() Then Exit`), po ugledu na
`modScrDokumenti.Scr_OtpTestSet`:

| Seam | Čemu služi |
|---|---|
| `BiTestSetUcitanost` | postavi `m_BlokoviLoadOk` / `_Err` — pad učitavanja se drugačije ne može izazvati bez lomljenja šeme |
| `BiTestSetFaktureUcitanost` | isto za fakture |
| `BiTestKooperantSme` / `BiTestKooperantPoruka` | **odluka koju handler stvarno čita**, ne njena kopija |
| `BiTestSetIzbor` | upis u kombo ide **kroz formu**, ne iz testa — kontrole su `Private`, a i operater prolazi tim putem |
| `BiTestBlokIzabran` | šta bi forma prijavila writeru kao „blok je izabran" |

### 11.4 Šta test tvrdi

`T_LegacyBanka_PadUcitavanjaNijePraznaLista` (114):

- pad učitavanja liste blokova **zaustavlja** ručno mapiranje, i operater dobija
  objašnjenje — ne ćutanje;
- uredno učitana lista **pušta** dalje (kapija ne sme da bude preširoka);
- izabran blok se writeru prijavljuje kao izabran;
- **prazan kombo nije izbor** — tada blok dolazi iz poziva na broj, gde je avans
  legitiman;
- kod kupca se blok **uopšte** ne prijavljuje.

Dve sabotaže, i to su **prve nad `.frm` fajlom**:
`banka-legacy-pad-liste-prolazi` i `banka-legacy-prazan-combo-je-izbor`. Druga
postoji zato što bi se pravilo moglo „ispraviti" tako što se prazan kombo
proglasi izborom — čime bi se ugasila legitimna grana.

### 11.5 Šta ovo ne pokriva

Sam `btnSacuvajRucno_Click` se i dalje ne izvršava u testu — pokriveno je
**pravilo** i to da ga handler zove (jedan red, proveren čitanjem). Klik, dijalozi
i redosled poziva ostaju smoke.

`frmDokumenta` nije diran. Ista tehnika bi radila i tamo, ali to je znatno veća
forma i zaseban posao.

> **Dopuna:** taj posao je urađen — v. §16. Ista greška je tamo i **nađena**, ne
> samo testabilnost.


---

## 12. Saldo izvoda: prvi red nije istina o celom izvodu (`v6-ui-181`)

Nalaz iz review-a #218, tada svesno odložen: `GetBankaIzvodiForGrid` uzima saldo
grupe sa **prvog** reda i pretpostavlja da su ostali isti.

### 12.1 Zašto je pretpostavka uopšte tu

Parser upisuje saldo izvoda na **svaki** red grupe. Agregat ih zato **uzima**, ne
sabira — sabiranje bi ih pomnožilo brojem stavki. To je ispravno **dok su svi
redovi saglasni**, a to niko nije proveravao.

Današnji parser ih ne može razići — kopira isti `saldo` u petlji
(`modBankaImport:614`). Ali ručno editovan red, delimičan re-import ili budući
parser mogu. Tada bi mreža prikazala brojku **prvog** reda kao istinu o celom
izvodu, bez ijednog traga.

### 12.2 Novo stanje, ne novo neslaganje

`BIM_SALDO_NEKONZISTENTAN` je **treće** stanje, ne varijanta „ne slaže se":

| Stanje | Šta znači |
|---|---|
| `_OK` | zna se šta piše, i slaže se |
| `_RAZLIKA` | zna se šta piše, i **ne** slaže se |
| `_NEMA` | legacy red, saldo metapodataka nema |
| `_NEKONZISTENTAN` | **ne zna se ni šta piše** — redovi nose različite zbirove |

**Nesaglasnost nadjačava i „slaže se".** To nije formalnost: prvi red fixture
para (`4500 + 500 − 0 = 5000`) **sam za sebe daje slaganje**, pa bi bez pravila u
koloni stajalo „slaže se" — tvrdnja o tačnosti brojki kojih zapravo nema.
Sabotaža `banka-izvod-saldo-prvi-red-pobedjuje` to i pokazuje:
`očekivano [3], dobijeno [1]`.

### 12.3 Zašto NE ulazi u čip „ne slaže se"

Čip nosi **jedno** tvrđenje. Nesaglasan izvod to tvrđenje ne podnosi — o
njegovim brojkama se ne zna ništa. Spajanje dva stanja pod jedan broj učinilo bi
brojku neupotrebljivom za oba. Nesaglasnost se vidi **u koloni**.

To je svestan izbor, ne previd: takav izvod i dalje traži čoveka, ali ga ne
traži iz istog razloga.

### 12.4 Prag je isti kao kod slaganja

`0.01`, isti koji koristi `BimSaldoStatus`. Uže poređenje bi zaokruženja
proglašavalo nesaglasnošću i brojka bi postala neupotrebljiva — sabotaža
`banka-izvod-saldo-prag-preuzak` meri baš to (`polovina centa nije`).

### 12.5 Dve stvari koje je dokaz našao na testu

**Tautološka tvrdnja.** Prva verzija je poredila `BuSlaganjeTekst` **samu sa
sobom** (`NEKONZISTENTAN <> RAZLIKA`). Pod sabotažom se obe strane menjaju isto,
pa je prolazila. Sada se poredi sa **katalogom poruka**.

**Pogrešan izvor statusa.** Prva verzija je status čitala iz reda **ekrana**, gde
je deseta kolona identitet — `CLng` nad njim daje `Type mismatch`. Status se čita
iz **čitača** (`GetBankaIzvodiForGrid`), a red ekrana se koristi za ono što on
zaista nosi: **tekst** kolone „Slaganje".

### 12.6 Verifikacija

Fixture dobija par redova istog izvoda sa različitim `ZavrsnoStanje`
(`BIM-FIX-NS1` / `NS2`). Tri sabotaže:
`banka-izvod-saldo-prvi-red-pobedjuje`, `banka-izvod-saldo-prag-preuzak`,
`banka-izvod-nesaglasno-je-razlika`.

Poruka je namerno kratka (`zbirovi se razlikuju`) — kolona je 132pt, a ostale
vrednosti u njoj su kratke i malim slovom. Prva verzija je bila rečenica od
trideset osam znakova, što je u toj koloni odsečen tekst.

### 12.7 Status bez posledice je pola posla

Prva verzija je menjala **samo status**. Četiri novčane ćelije i podnožje su i
dalje koristile vrednost **prvog** reda — dakle kod je pisao „ne zna se koji
zbirovi važe", a pored toga prikazivao `5.000` kao uredan saldo i **sabirao ga u
promet**. To je ista klasa koju §10 zove „tuđi podatak koji izgleda kao svoj".

Sada nesaglasan izvod nema brojke **nigde**:

| Gde | Pre | Sada |
|---|---|---|
| četiri novčane ćelije | vrednost prvog reda | **prazno** |
| promet u podnožju | ulazi | **ne ulazi** |
| kolona Slaganje | — | `zbirovi se razlikuju` |

**Kolone su zato `rest`, ne `rsd`.** `rest` prazni ćeliju kad je vrednost nula,
`rsd` bi napisao `0,00`. Ovde nula znači „nema podatka", ne „nula dinara" —
legacy uvoz saldo metapodatke nema, a nesaglasan izvod ih ima ali se ne zna koji
važe. Cena je **bold**: `StyleGridCell` ga daje `rsd` koloni, `rest` ne.
Prihvaćeno — četiri podebljane novčane kolone u istom redu ionako nisu davale
hijerarhiju. Poravnanje ostaje desno (`ColIsNum` poznaje `rest`).

### 12.8 `rest` je ugasio podnožje — promena vrste kolone je promena ugovora

Prelazak na `rest` je imao posledicu koju nisam predvideo: ljuska odlučuje da li
uopšte **crta** zbir vrednosti preko `ModeHasValCol()`, a taj spisak je znao samo
`rsd`, `mult` i `sum0`.

```
Scr_Rows(...)(4)  ->  promet uredno izračunat
ModeHasValCol()   ->  False
ftVal.Visible     ->  False
                  ->  operater ne vidi promet
```

Test je bio zelen jer je tvrdio **podatak ekrana**, a ne **odluku ljuske**.

**`rest` jeste novčana kolona**, pa je dopisan na spisak. Pre toga je urađen
audit, jer se time menja ugovor ljuske i za druge ekrane: jedina druga `rest`
kolona u repou je `OTKUI_HD_OSTATAK` na Dokumentima, a ona stoji **uz `mult`
kolonu** — tamo je `ModeHasValCol` već bio `True`. Proširenje je dakle **no-op**
za sve postojeće ekrane i menja samo listu izvoda.

Uz to je uveden seam `GridImaValKolonuTest`, pa test od sada tvrdi i **da ljuska
prikazuje** zbir, ne samo da ga ekran izračuna. Sabotaža
`mreza-rest-nije-novcana-kolona` skida `rest` sa spiska i obara baš tu tvrdnju.

**Cena koju `rest` nosi, i koja ostaje:** nula se u te četiri kolone ne
razlikuje od „nema podatka" — uredan izvod sa `Isplate = 0` prikazuje **prazno**,
ne `0,00`. Na izvodu je to gotovo uvek isto značenje; jedini slučaj u kom nije je
nov račun sa nultim početnim stanjem. Prihvaćeno svesno i stavljeno u smoke
listu, jer je to stvar oka a ne tvrdnje.

### 12.9 Zbir podnožja nije poštovao pretragu

Nađeno uz isti nalaz, i **starije** od njega: akumulacija je stajala **između**
čipa i pretrage. Izvod koji pretraga sakrije i dalje je ulazio u promet — traka
je tvrdila promet redova kojih na ekranu nema.

Zbir sada ide **posle oba filtera**.

Sabotaža `banka-izvod-promet-ne-postuje-pretragu` vraća stari redosled. Prva
verzija te sabotaže je zbir **duplirala** umesto da ga premesti, pa je merenje
bilo zamućeno i obarala je tuđu tvrdnju (zamka 5). Uz to je redosled tvrdnji u
testu izmenjen: provera pretrage ide **pre** provere nesaglasnog izvoda, jer bi
inače obe sabotaže padale na istoj tvrdnji.

### 12.10 Fixture koji izjednači dve brojke ubija tuđu sabotažu

Nova dva reda su prvo oba nosila `Obradjeno = "Da"`. Suite je bila zelena, ali je
dvosmerni dokaz prijavio **43 crvene od 44**: `banka-uvoz-znacka-broji-mapirane`
prestala je da obara išta.

Uzrok nije bio ni u sabotaži ni u kodu koji ona gađa. Ta sabotaža znački podmeće
`k(1)` (mapirane) umesto `k(0)` (otvorene) — a nova dva reda su te brojke
**izjednačila na 6 i 6**. Zamena tada ne menja ništa.

Najgore je što je to *tiho*: čovek koji doda fixture red uredno ažurira
`FX_BIM_OTVORENIH` i `FX_BIM_OBRADJENIH` na nove vrednosti (upravo to sam i
uradio), sve prođe, a jedan dokaz je nestao.

Zato jedan od novih redova ostaje **otvoren**, a test dobija tvrdnju koja čuva
**sam dokaz**, ne ponašanje:

```vb
AssertEq (nZa <> nObr), True, _
         "otvorenih i mapiranih MORA biti razlicito -- inace sabotaza znacke ne meri nista"
```

Provereno simulacijom cele greške — red vraćen na `"Da"` **i** sve konstante
usklađene: pada tačno jedna tvrdnja, ova.

To je nova vrsta nalaza u ovom projektu: do sada je zamka 9 hvatala sabotažu koja
je ostala bez sidra, a ovde je sidro bilo ispravno — **podatak** ju je obesmislio.
---

## 13. Podnožje mreže je brojalo tuđim režimom (`v6-ui-182`)

Mreža je zajednička: nose je i ekran dokumenata i svaki ugovorni ekran. Njeno
podnožje je jedinicu i broj decimala biralo iz **`ActiveMode`** — a to je režim
**unosa dokumenata** (`F1`…`F7`), koji ugovorni ekran nema.

### 13.1 Šta je operater video

Ko je bio na **F7 (Reversi)** pa prešao na **Uvoz izvoda**, u podnožju je dobio:

| | |
|---|---|
| pisalo je | `Ukupno 8.950 kom` |
| trebalo je | `Vrednost 8.950,00 RSD` |

Dakle **novac izbrojan kao komadi**, i još bez para. Broj je bio tačan — zbir
prometa — ali je nosio tuđu jedinicu, pa je čitanje bilo pogrešno u oba smera:
ni „8.950 komada" nije istina, ni „8.950 dinara" (bilo je `8.950,00`).

Ista klasa kao traka `zOtp` koja je na tuđem ekranu ostajala upaljena sa
porukom „Nema izabrane otpremnice": stanje ekrana dokumenata koje curi na ekran
koji o njemu ne zna ništa.

### 13.2 Zašto je `ActiveMode` uopšte bio tu

Ljuska je počela **kao ekran dokumenata** — mreža je tada imala tačno jednog
korisnika, pa je „režim" bio isto što i „ekran". Ugovorni ekrani su došli
kasnije i nasledili mrežu, ali ne i to pitanje.

Zato ovo nije previd u jednom redu nego **poslednji ostatak** pretpostavke da
mreža pripada dokumentima.

### 13.3 Ugovor umesto globala

Ljuska više ne čita `ActiveMode` — **pita ekran**:

```vb
Public Function ScrBrojiKomade(ByVal kljuc As String) As Boolean
```

Isti oblik kao `ScrBrojac`: kasno vezano preko `Application.Run`, greška se guta,
a ekran koji `Scr_BrojiKomade` **ne implementira** dobija `False` — dinare.

**Fail-closed je ovde bitniji nego što izgleda.** Da nepoznat odgovor znači
„komadi", svaki budući ekran bi novac prikazivao u komadima dok se neko ne seti
da doda ugovor. Ovako novi ekran ćuti i dobija ispravno; komade traži samo onaj
ko ih stvarno broji.

Dokumenta odgovaraju iz **svog** stanja:

```vb
Public Function Scr_BrojiKomade() As Boolean
    Scr_BrojiKomade = ModeBrojiKomade(ActiveMode)
End Function
```

`ActiveMode` time nije nestao — **preselio se tamo gde znači nešto**.

### 13.4 Storno je drugi korisnik reversa — nađeno u review-u

Prva verzija ovog rada je „ekran koji ne odgovori dobija dinare" tretirala kao
bezbedan podrazumevan odgovor. Za Banku i Fakture jeste. Za **Storno nije**, i to
determinističi:

- storno ekran bira **tip dokumenta**, i među osam tipova je i `STIP_REVERSI`;
- redove ne pravi sam nego ih uzima od **`modScrDokumenti.RedoviZaTip`** — istog
  čitača koji puni ekran dokumenata;
- kolone za `REVERSI` nose `OTKUI_HD_KOMADA|COL_AMB_KOLICINA|sum0`, dakle peta
  kolona je **broj komada**, i taj broj stiže u `mSumVal`.

`modScrStorno` ugovor nije implementirao, pa bi za 125 reversa u podnožju stajalo
`Vrednost 125,00 RSD`. Pre ove izmene je isti ekran zavisio od zatečenog
`ActiveMode` i mogao **slučajno** biti tačan (ako je iza operatera ostao `F7`);
fail-closed ga je učinio **uvek** pogrešnim. Popravka jedne klase greške je
otvorila drugu na istom mestu — zato ugovor mora da pokrije **svakog** korisnika
liste, ne samo prvog.

**Audit se zatvara:** `RedoviZaTip` u celom repou zovu tačno dva ekrana —
`modScrDokumenti` i `modScrStorno`. Trećeg nema.

Storno odgovara iz **svoje** aktivne liste:

```vb
Public Function Scr_BrojiKomade() As Boolean
    Scr_BrojiKomade = modScrDokumenti.TipBrojiKomade(Scr_Lista())
End Function
```

**Zašto novi primitiv `TipBrojiKomade`, a ne poziv `ModeBrojiKomade(Scr_Lista())`:**
`ModeBrojiKomade` prima **F-ključ** i vrti ga kroz `modeKey`, a `modeKey`
nepoznat ključ svodi na `"OTKUP"` (`Case Else`). Poziv sa tip-ključem
(`"REVERSI"`) zato izgleda ispravno a **tiho vraća `False`**. Poređenje je
izdvojeno u `TipBrojiKomade(tk)` da literal `"REVERSI"` ostane na jednom mestu i
da oba pozivaoca pitaju istu stvar; `ModeBrojiKomade` je sada tanak omotač nad
njim.

**`vba_check` je odmah tražio svoje:** drugi `Public Scr_BrojiKomade` je pao kao
`DUPLIKAT` („Ambiguous name detected"), jer ime nije bilo u `SCR_UGOVOR` — spisku
procedura ugovora ekrana koje smeju da postoje u više `modScr*` modula. Time je
dodavanje druge implementacije i formalno pretvorilo `Scr_BrojiKomade` u član
ugovora, a ne u „još jednu funkciju".

### 13.5 Natpis na jednom mestu

Crtanje traži pravu formu, a jedinica i decimale se vide tek u **gotovom
natpisu**. Zato je natpis izdvojen iz `RenderGrid` u `PodnozjeValTekst(iznos)`,
a test seam `GridPodnozjeValTest` zove **baš tu funkciju**.

Prva verzija seam-a je logiku **prepisala** — što je greška koju je review na
ovom projektu već dvaput hvatao: kopija se s vremenom raziđe sa originalom, i
tvrdnja tiho počne da meri kopiju.

### 13.6 Šta test tvrdi

`T_Mreza_PodnozjeJedinicaIdeIzUgovoraEkrana` (115) meri na **dva nivoa**:

| Nivo | Tvrdnja |
|---|---|
| ugovor | Dokumenta na reversima **i dalje** broje komade |
| ugovor | ...a na ambalaži (`F5`) ne — ekran prati **svoj** režim |
| ugovor | ugovorni ekran **ne nasleđuje** režim unosa dokumenata |
| ugovor | **Storno** kad prikazuje reverse broji komade |
| ugovor | ...a na ostalim tipovima na Stornu ne broji |
| ljuska | podnožje ugovornog ekrana ne pominje komade |
| ljuska | ...nego dinare |
| ljuska | ...i zove se „Vrednost", ne „Ukupno" |
| ljuska | novac ide **sa parama** |

**Zašto sam nivo ljuske nije dovoljan:** tvrdnja „na ugovornom ekranu su dinari"
prolazi i kad se komadi ugase **svima**. Reversi bi tiho izgubili svoju jedinicu,
a suite bi ostala zelena. Zato prve tri tvrdnje čuvaju **postojeće** ponašanje.

Jedinica i decimale se tvrde **odvojeno** iako su ista `If` grana: `1234.56` bez
para nema `56` nigde u natpisu, pa ta tvrdnja pada sama za sebe.

Dve tvrdnje o Stornu, ne jedna: „Storno broji komade" prošlo bi i da ekran
odgovara `True` **uvek**, čime bi fakture i izvodi na njemu postali komadi. Ista
simetrija kao kod Dokumenata.

Test se izvršava sa **zatrovanim** `ActiveMode = "F7"` i vraća ga na zatečenu
vrednost pre nego što išta tvrdi — pad tvrdnje ne sme da ostavi ljusku u tuđem
režimu. Isto važi i za aktivnu listu Storna (`Scr_TipTestSet`).

### 13.7 `ModeValUnit` je ostao bez posla

Jedinicu je do sada davao `modScrDokumenti.ModeValUnit(mode)`. Posle izmene je
`PodnozjeValTekst` uzima iz kataloga poruka, pa je ta funkcija ostala bez
ijednog pozivaoca — i **obrisana** je.

Nije kozmetika: `Public` funkcija koja jedinicu računa iz globalnog režima je
tačno ona zamka koja je i napravila ovaj kvar. Sledeći koji bi je našao pomislio
bi da je to živi put.

### 13.8 Ostala čitanja `ActiveMode` u ljusci — pregledana

Pošto je kvar bio klasa, a ne red, pregledana su **sva** čitanja `ActiveMode` u
`modOtkupUI`:

| Mesto | Presuda |
|---|---|
| `RefreshGridTitle` (naslov mreže) | ograđeno sa `ElseIf mScreen = "DOKUMENTI"` |
| `LayoutOtkup` (traka `zOtp`, visina) | nedostižno — `LayoutOtkup` za ugovorni ekran izlazi ranije (`Exit Sub`) |
| `LoadRowIntoForm` (dupli klik) | ograđeno sa `If mScreen = SCR_POCETNI` |
| `LayoutGrid`: `If Not IsArray(mCols) Then SetGridCols modeKey(ActiveMode)` | **jedino preostalo** — v. ispod |

`LayoutGrid` se **zove i na ugovornom ekranu** (iz `LayoutScreenZone`), pa je taj
red dostupan. Pali se samo ako opis kolona nikad nije postavljen, a punjenje
odmah zatim ga prepiše. **Nije mereno** i nije dirano u ovom PR-u — promena tog
reda traži odgovor na pitanje šta uopšte znači raspored mreže pre prvog punjenja.

### 13.9 Uz to: u sidebar se vraća verzija programa

Od `v6-ui-154` je na nereleasovanoj svesci na tom mestu stajao `OTKUI_BUILD`,
**izričito privremeno** — do kraja rada na storno ekranu, koji je završen još u
Fazi D. Razlog je bio dobar: svaka nereleasovana sveska nosi isti `v0.0.0-dev`,
pa se sa ekrana nije videlo da li je posle `ImportAllVBA` u njoj nov ili star UI
kod.

Dijagnostika se ne gubi nego se veže za **postojeći** prekidač: `UI_DEBUG=DA` u
`tblLocalConfig` (`IsDebugUI`). Ko meri — upali ga; ko radi — vidi verziju
programa.

**Nije pokriveno testom, i to namerno:** provera bi tražila upis u
`tblLocalConfig`, a `RunAllTests` je nemutirajuća suite (mutirajući test obara
`who_writes` kapiju u CI-ju). Ide u smoke listu.

### 13.10 Verifikacija

| Šta | Ishod |
|---|---|
| `RunAllTests` | **115 testova, 0 palih** |
| Šest novih sabotaža | **6 / 6** — svaka obara svoj test i **svoju** tvrdnju |
| Dvosmerni dokaz nad izmenjenim fajlovima (39 + 16 sabotaža) | **rupe samo tamo gde su i zatečene na `main`-u** (§13.11) |
| Sva 222 sidra posle izmena | ni jedno **novo** zastarelo |
| `vba_check` · `--self-test` | čisto (195 fajlova) · 64 slučaja |
| izvor posle svih vraćanja | bit-identičan |

Nove sabotaže, po jedna na svako svojstvo:
`mreza-podnozje-ljuska-ne-pita-ekran`, `mreza-podnozje-ugovor-fail-open`,
`mreza-podnozje-jedinica-iz-globalnog-rezima`, `mreza-podnozje-novac-bez-para`,
`mreza-podnozje-storno-ne-prijavljuje-komade`, `mreza-podnozje-storno-uvek-komadi`.

Izmena u `tools/vba_check.py` (`scr_brojikomade` u `SCR_UGOVOR`) dokazana je u
oba smera: bez tog imena `vba_check` prijavi `DUPLIKAT` **po imenu** za
`modScrDokumenti.bas` i `modScrStorno.bas`, sa njim je čisto.

### 13.11 Šta je pun prolaz našao o samom dokazu

Dvosmerni dokaz je do sada vrćen nad **podskupom** kataloga (banka/mreža, 48 od
220). Prvi put je pušten nad svim sabotažama koje gađaju izmenjene fajlove — i
odmah je pokazao tri vrste rupe. **Sve su zatečene:** provereno tako što su ista
sidra i isti ciljni kod uzeti iz `origin/main`.

**1) Sidra koja su zastarela — 10 od 220.** Kod ispod njih je u međuvremenu
popravljen, pa se sidro više ne nalazi. Sabotaža tada ne javlja „test je prošao"
nego „nisam našao sidro" — a to u petlji od pola sata lako prođe kao zeleno. To
je zamka 9, i sada se vidi da nije bila jedan slučaj nego klasa:

`uplata-preko-fakture`, `f8-jedna-tabela`, `storno-nema-dok`,
`relink-izvor-po-broju`, `relink-ignorise-generaciju` (2 pogotka, dvosmisleno),
`relink-cilj-bez-kapije`, `ljuska-odseca-liste`, `blokovi-po-broju`,
`f8-identitet-po-broju`, `oporavak-cid-ne-stize-u-red`.

Za tih deset **dokaza trenutno nema**: tvrdnje koje su nekad bile pokazane
crvenim danas nisu.

**2) Sabotaža koja ne obara ništa.** `parcela-tekst` se uredno primeni, a suite
ostane zelena — dakle `T_ParcelaID_IzSkriveneKolone` ne meri baš ono što ta
sabotaža kvari. Ciljni kod i telo testa su bajt-identični `main`-u.

**3) Polje „očekivana tvrdnja" je parafraza, ne merena vrednost.** Katalog uz
svaku sabotažu nosi tekst tvrdnje, ali ga niko nije poredio sa onim što stvarno
padne. Poređenje sada pokazuje da se u većini slučajeva radi o prepričavanju iste
tvrdnje, ali u dva slučaja pada **druga**:

| Sabotaža | katalog kaže | stvarno padne |
|---|---|---|
| `parse-cdate` | „godina van poslovnog opsega" | „mesec 13 se odbija, ne preliva u sledeću godinu" |
| `mreza-crtanje-kvari-stil` | „crtanje ne menja poravnanje koje je layout postavio" | **„preduslov:** brojčana kolona je poravnata DESNO" |

Drugi red je zamka 6 u čistom obliku: pada **preduslov**, pa poslovna tvrdnja
ispod njega uopšte i ne dođe na red.

**4) Dve sabotaže dele tvrdnju.** `zatecen-context-bez-kapije` i
`stale-parent-po-broju` gađaju isti red u `modStornoFlow.bas` i oslanjaju se na
istu tvrdnju testa `T_ZatecenContext_NePrevezujeTudjePrijemnice`. Test ih ne
razlikuje.

**Ispravka mog ranijeg izveštaja:** „mašinski provereno da nijedna tvrdnja nije
deljena" važilo je za podskup koji se tada vrteo, ne za katalog.

**Šta iz ovoga sledi.** Pravilo iz `CLAUDE.md` — „posle izmene pusti ceo dvosmerni
dokaz i tvrdi `crvenih == sabotaža`" — nad 222 sabotaže traje oko dva i po sata,
pa se u praksi nije puštalo celo. Jeftina polovina tog pravila je **statička**:
proveriti da se svako sidro nalazi tačno jednom, istim poređenjem koje koristi
`sabotaza.py`. To traje sekundu i uhvatilo bi svih deset.

**Zatvoreno:** `python tools/sabotaza.py --proveri-sidra` (ide i kroz `vba_check`,
dakle posle svake VBA izmene) i `python tools/dokaz.py [filter]` za pun dokaz.
Svih deset sidara je popravljeno i dokazano ponovo. Ceo tok, sa jednom „mrtvom"
sabotažom koja to nije bila, zapisan je u
`docs/engineering/postmortems/2026-08-verifikacija.md` §10.

### 13.12 Šta OSTAJE

Podnožje i dalje nosi **jedan** novčani podatak. Lista izvoda ima dva koja
operater traži (uplate i isplate), ali drugi slot je **promena ugovora**
(`Scr_Rows` bi vraćao par natpis/vrednost, plus nova kontrola `ftVal2`) i ide
zasebno.
---

## 14. Podnožje nosi dva novčana broja (`v6-ui-183`)

Zbir vrednosti u podnožju je **jedan** broj. Lista izvoda ima dva koja operater
traži: **uplate** i **isplate**. Do sada je videla samo njihov zbir — promet.

### 14.1 Zašto zbir nije dovoljan

Promet se ne može rastaviti unazad. Operater koji drži izvod u ruci ima na njemu
uplate i isplate odvojeno, i to su brojke koje poredi. `Promet 12.400,00 RSD` mu
ne kaže ništa o tome da li se izvod slaže — a upravo je zbog toga i otvorio tu
listu.

### 14.2 Sedmi član ugovora

`Scr_Rows` je vraćao šest članova; sedmi je **neobavezan** i nosi parove
`Array(kljucPoruke, iznos)`:

```vb
RedoviIzvodi = Array(IzvodiKolone(), outA, n, 0#, zbirU + zbirI, Array(0, 0, 0), _
                     Array(Array("OTKUI_FT_UPLATE", zbirU), _
                           Array("OTKUI_FT_ISPLATE", zbirI)))
```

**Zašto od ekrana, a ne računanje u ljusci:** brojke moraju biti izbrojane pod
**istim** filterima pod kojima su i redovi. Ljuska filtere ne primenjuje — čip i
pretragu razrešava ekran — pa bi svaki račun sa strane opisivao drugu listu od
one koja se vidi. Peti član (zbir vrednosti) ostaje promet, da ekran koji slotove
ne šalje i dalje ima smislen broj.

Ekran koji sedmi član ne pošalje ponaša se tačno kao pre. `LoadGridFromScreen` ga
čita odbranjeno (`If UBound(d) >= 6`), istim oblikom kojim već čita šesti.

### 14.3 Bazen, kao i svuda u ljusci

`MAX_FT_VAL = 2`. Ekran koji zatraži više ne greši, ali se višak **ne crta** i
`BazenStaje` to prijavi — isti mehanizam kao za čipove, radnje, liste i kolone.
Tiho odsecanje je već dvaput koštalo pun krug (jedanaesti čip, šesta radnja).

**Geometrija:** desna ivica novčanog dela je fiksna, drugi slot se dodaje
**ulevo** — da se jedan slot ne pomeri kad se pojavi drugi. Taj prostor deli sa
zbirom kilograma; ako su oba tu, crta se samo prvi, i to opet kroz `BazenStaje`
(`podnozje-uz-kg`), ne prećutno. Danas se ne dešava: obe liste koje traže slotove
nemaju kilograme.

### 14.4 Šta test tvrdi

`T_Mreza_PodnozjeDvaNovcanaSlota` (116), redom:

| Nivo | Tvrdnja |
|---|---|
| ugovor | sedmi član postoji i nosi **dva** slota, sa svojim ključevima |
| ugovor | **zbir slotova je promet** — dakle brojani su pod istim filterima |
| ugovor | pretraga smanjuje slotove, i to **oba**, ne samo prvi |
| ugovor | i u suženoj listi je zbir slotova **njen** promet |
| ugovor | **lista Stavke** nosi isti ugovor — ekran ima dva čitača |
| ljuska | ljuska je preuzela **oba** |
| ljuska | svaki nosi svoj natpis (Uplate / Isplate) i dinare |
| ljuska | **slotovi ne nose isti IZNOS** |
| ljuska | ekran bez slotova ostaje na zbiru vrednosti |
| crtanje | drugi slot je **stvarno nacrtan** (`ftVal2.Visible`), uz prvi |
| crtanje | na ekranu bez slotova se **gasi** — ne ostaje tuđa brojka |

Test **izričito** postavlja listu (`Scr_BuListaTestSet "IZVODI"`) pre prvog
čitanja. Prva verzija je merila ono što je prethodni test ostavio — zelen dok se
ne promeni redosled, a zove se po izvodima.

Crtanje se meri nad **pravom formom** (`GridRenderTest`), jer model i natpis nisu
isto što i nacrtan slot: kod ove iste liste je već jednom bio tačan zbir uz
nevidljivu kontrolu (prelazak novčanih kolona na `rest`, `v6-ui-181`).

Najvažnija je treća: **dva broja koja ne prate filtere gora su od jednog koji ih
prati**, jer izgledaju preciznije. Zato se ne tvrdi „slotovi postoje" nego da im
je zbir tačno onaj promet koji lista i prikazuje.

Poslednja postoji zbog najtišeg mogućeg kvara: oba slota crtaju **prvi** iznos.
Podnožje tada pokazuje dva broja, izgleda savršeno ispravno, a to je jedan isti.

**Prva verzija te tvrdnje nije merila ništa.** Poredila je cele natpise
(`t0 <> t1`) — a oni se razlikuju već po natpisu (`Uplate` vs `Isplate`), pa bi
prošla i kad oba slota nose isti iznos. Nađeno tako što je sabotaža oborila
**drugu** tvrdnju od očekivane: strogo pravilo iz #227 traži da padne baš ona
imenovana, pa se razmimoilaženje vidi umesto da prođe kao zeleno.

Sada se poredi ono što ostane kad se natpis skine, uz preduslov da se uplate i
isplate u test-svesci uopšte razlikuju — bez njega bi i oštra tvrdnja bila prazna.

### 14.5 Verifikacija

| Šta | Ishod |
|---|---|
| `RunAllTests` | **116 testova, 0 palih** |
| `tools/dokaz.py mreza-podnozje` | svaka obara **svoju** tvrdnju, izvor bit-identičan |
| `sabotaza.py --proveri-sidra` | čisto (229 sabotaža) |
| `vba_check` · `--self-test` | čisto (195) · 64 |

Sedam novih sabotaža. Dve od njih postoje zato što je review pokazao da bez njih
PR tvrdi više nego što dokazuje:

- `mreza-podnozje-stavke-nema-slotova` — lista **Stavke** je imala nov ugovor a
  nijedan dokaz;
- `mreza-podnozje-slot-ignorise-pretragu` — slot koji je **tačan nad punom** a
  pogrešan nad **suženom** listom. Najteži oblik, jer prva provera prolazi.
  Promet se namerno ne dira: da je diran, pao bi test prometa iznad i sabotaža ne
  bi merila svoje (zamka 5).

### 14.6 Zamka 4 uživo: sabotaža koja ne pada nego VISI

Prva verzija sabotaže `mreza-podnozje-slot-mimo-filtera` nosila je oznaku na
kraju reda koji se **nastavlja**:

```vb
Array(Array("OTKUI_FT_UPLATE", zbirU + zbirI), _   ' SABOTAZA: slot broji mimo liste
```

U VBA `_` mora biti **poslednji** znak u redu. Ovo je zato syntax error, a
posledica nije pao test nego **compile** — run visi do timeout-a, Excel ostaje u
`[break]`, a dokaz stoji na drugoj sabotaži i ne javlja ništa.

To je zamka 4 iz `tools/sabotaza.py`, zapisana odavno — i svejedno pokupljena
ponovo, jer je ništa nije proveravalo. Statička provera koja je hvata (`sabotaza.py
--proveri-sidra`) stiže tek u #227; do tada je jedini znak bio taj što run stoji.

Oznaka je premeštena u red **iznad**. Provereno mašinski da nijedna od 227 zamena
u katalogu nema komentar posle `_`.

### 14.7 Self-test iz #227 pao je na sopstvenom fixture-u

Podizanje builda na `v6-ui-183` oborilo je `sabotaza.py --self-test`:

```
SELF-TEST: zdrav unos je prijavljen kao nalaz
```

Njegov „zdrav" fixture je stajao na redu `Public Const OTKUI_BUILD ... "v6-ui-182"` —
**vrednosti koja se menja sa svakim izdanjem**. Čim se build podigao, sidro se više
nije nalazilo, pa je provera ispravno prijavila ono što joj je podmetnuto.

Fixture je premešten na `Option Explicit` i prvi red ispod njega — redove koji se
ne menjaju. Provereno da izmena nije oslabila proveru: gašenje četiri pravila
redom i dalje obara self-test **po imenu**.

Pouka je ista kao za sidra sabotaža: **fixture koji stoji na promenljivoj
vrednosti zastareva isto kao i sidro**, samo tiše — jer ga niko ne gleda dok ne
pukne.

Iz review-a je stigla još jedna sitnica istog reda: višeredni sintetički slučaj
stajao je na `Option Explicit` i prvom redu palete, između kojih u fajlu stoji
**prazan red**. Pored ciljanog nalaza davao je i višak (zastarelo sidro).
Self-test je prolazio jer traži baš ciljani, ali fixture koji uz tačan nalaz
nosi i netačan meri manje nego što izgleda. Sada stoji na dva stvarno susedna
reda i daje **tačno jedan** nalaz.

### 14.8 Šta OSTAJE

**Bazen nije izmeren.** Da se višak preko `MAX_FT_VAL` stvarno odseca vidi se
samo ako neki ekran zatraži tri slota — a nijedan ne traži. Sabotaža koja spusti
konstantu obara istu tvrdnju kao ona koja ljusci oduzme čitanje slotova (zamka
5), pa nije dodata. Ostaje kao poznata rupa, ne kao prećutana.
---

## 15. Rupa iz §10.6 nije ni postojala

§10.6 je čišćenje pozadine prave `pill` ćelije zapisalo kao **neizmereno**, uz
obrazloženje: jedina lista koja se puni bez forme a ima statusnu oznaku je
FAKTURE, a njena je `paypill`; prava `pill` kolona živi na Dokumentima, „čija se
lista bez izabranog režima ne puni".

Drugi deo te rečenice nije tačan.

### 15.1 Šta je stvarno stajalo na putu

`modScrDokumenti.Scr_Rows` prvo gleda **podlistu** (`Scr_Lista()`), jer režim F1
ima tri svoje liste (otpremnice, blokovi, izgubljeni, kooperanti) pored zatečene
liste dokumenata. Podrazumevano je `"SVI"` — dakle **lista dokumenata**.

Prva sonda je svejedno dobila otpremnice, jer je `mLista` ostao od ranijeg testa.
To je izgledalo kao „lista se ne puni", a bilo je „puni se druga lista".

Ništa novo nije trebalo: podlista se bira **produkcionim putem**, klikom na čip
(`Scr_Event "lsSVI"`), isto kako je bira operater. Mereno: **14 redova, 13 kolona,
trinaesta je baš `pill`**.

### 15.2 Zašto sonda nije odmah rekla istinu

Prva verzija je čitala vrste kolona i zaključivala iz njih. To je dalo devet
kolona bez `pill`-a i izgledalo kao potvrda da lista nema statusnu oznaku — a
zapravo je opisivalo **drugu listu**. Tek kad je sonda počela da poredi
`GridCols("OTKUP")` (13 kolona, poslednja `pill`) sa onim što je učitano, razlika
se videla.

Pouka je ista kao kod poruke o koloni (§11 postmortema): **kad zaključak zavisi od
stanja, izmeri i stanje**, ne samo ishod.

### 15.3 Šta test 118 tvrdi

`T_MrezaPilula_PozadinaSeCisti` radi nad **pravom formom** (`GridRenderTest`), nad
listom Dokumenata:

| Tvrdnja | |
|---|---|
| preduslov | lista se puni bez forme i ima `pill` kolonu |
| preduslov | uredna pilula je **naslikana** (neprozirna pozadina) |
| **glavna** | pilula koja se ne može prikazati gubi i **pozadinu**, ne samo natpis |
| | ...i natpis |
| kontrola | pre kvara je natpis postojao |
| round-trip | uredna vrednost **vraća** pozadinu |

Kolona se **traži** (`GridKindKoloneTest`), ne pretpostavlja. Neispravna vrednost
ide **posle** učitavanja (`GridTestVrednost`), ne u tabelu — takav red je jednom
već oborio sedam tuđih testova sa `Overflow`.

**Round-trip tvrdnja nema svoju sabotažu**, i to je izmereno, ne pretpostavljeno:
pozadinu i prvi put i posle popravke slika **ista** rutina (`PaintPill`), pa svaka
sabotaža nad njom obara **preduslov** umesto te tvrdnje (zamka 6). Prvi pokušaj je
to i pokazao, pa je sabotaža uklonjena umesto da joj se tekst „namesti".

### 15.4 Verifikacija

| Šta | Ishod |
|---|---|
| `RunAllTests` | **118 testova, 0 palih** |
| `tools/dokaz.py mreza-pilula` | **2 / 2**, izvor bit-identičan, `DOKAZANO` |
| `sabotaza --proveri-sidra` | čisto |
| `vba_check` | čisto (195 fajlova) |

Nova sabotaža: `mreza-pilula-pozadina-ostaje` — vraća stanje u kom se briše natpis
a pozadina ostaje, dakle obojen pravougaonik koji i dalje tvrdi stanje, samo bez
slova.

### 15.5 Šta i dalje NIJE urađeno

`frmDokumenta` **nije diran**. Ovo je zatvorilo rupu u mreži, ne u legacy formi;
seam-ovi za nju (po ugledu na `frmBankaImport`, §11) ostaju zaseban posao.
---

## 16. `frmDokumenta`: prazna lista blokova je bila avans

§11 je istu klasu greške zatvorio u `frmBankaImport` i zapisao da `frmDokumenta`
„nije diran, ista tehnika bi radila i tamo". Kad se tamo pogledalo, greška nije
bila samo moguća — bila je **prisutna**.

### 16.1 Šta je bilo

`btnUnosOMUlaz_Click`, grana koja odlučuje **na šta ide novac**:

```vb
If cmbOtkupBlok.ListIndex >= 0 Then
    ...                                  ' knjizi na izabran blok
Else
    tipNovca = NOV_VIRMAN_AVANS_KOOP     ' AVANS
End If
```

`ListIndex = -1` znači dve različite stvari:

| Stanje | Da li je avans tačan |
|---|---|
| kooperant nema otvorenih blokova | **da** |
| lista blokova se **nije učitala** | **ne** |

A `FillOpenOtkupi` je pad čitanja gubio bez traga: prvo `cmbOtkupBlok.Clear`, pa
`GetOpenOtkupi` — bez `On Error`. Pozivalac (`cmbPrimalacOMUlaz_Change`) **takođe
nema rukovaoca**, pa se od celog neuspeha vidi samo prazan kombo.

Ishod: novac tiho postaje **avans kooperanta** umesto da se knjiži na blok.

### 16.2 Šta je urađeno

Isti oblik kao u §11.2, jer je i greška ista:

- `FillOpenOtkupi` beleži ishod (`m_BlokoviOk` / `m_BlokoviErr`) i više ne guta pad;
- pravilo je izdvojeno u `BlokIzborSme(ByRef outPoruka)`, koje **vraća** poruku
  umesto da je prikaže — time ostaje bez dijaloga i pozivljivo iz testa;
- handler ga pita **baš u toj grani**, pre nego što izabere avans.

„Prazno je istina" ostaje: posle **uspešnog** čitanja prazna lista stvarno znači
da kooperant nema otvorenih blokova, i avans je tada ispravan. Kapija je uska
koliko i kvar.

### 16.2.1 Prva verzija je propuštala nedostajuću tabelu — nađeno u review-u

Beleženje pada hvata samo ono što **pukne**. `GetOpenOtkupi` čita kroz
`GetTableData`, a on vraća **isti `Empty`** i kad tabele nema i kad je prazna —
bez ijedne greške:

```
tblOtkup NEDOSTAJE -> GetTableData = Empty -> GetOpenOtkupi = Empty
                   -> nema Err -> m_BlokoviOk ostaje True -> AVANS
```

Dakle baš poslovna greška koju ovaj posao zatvara ostajala je otvorena kroz jedan
poznat put. To nije nijansa: `RequireTable` je i napravljen zbog te klase
(„prazna tabela i nepostojeća tabela nisu isti ishod"), a `frmBankaImport` — koji
je ovde naveden kao uzor — to **već radi** pre svog čitanja. Prekopiran je oblik,
ali ne i ta linija.

`FillOpenOtkupi` sada traži tabelu pre čitanja:

```vb
RequireTable TBL_OTKUP, "frmDokumenta.FillOpenOtkupi"
```

Time se tri stanja razdvajaju kako treba:

| Stanje | Ishod |
|---|---|
| tabela postoji, nema redova | `m_BlokoviOk = True` → avans je **legitiman** |
| tabela ne postoji | `RequireTable` diže → `EH` → `m_BlokoviOk = False` → **STOP** |
| čitanje pukne iz drugog razloga | `EH` → `m_BlokoviOk = False` → **STOP** |

### 16.2.2 Kapija je vazila samo kad izbora nema — takođe iz review-a

Beleženje pada rešava **stanje**, ali ne i **mesto** provere. Kapija je prvo
stajala u `Else` grani, dakle pitala se samo kad `ListIndex = -1`:

```
pad usred petlje -> kombo ostane DELIMICNO napunjen
                 -> operater izabere red iz nepotpune liste
                 -> ListIndex >= 0 -> kapija se NIKAD ne pita -> knjizenje ide dalje
```

Obećanje „ako učitavanje zakaže, unos staje" tada važi **samo za pola slučajeva**.

Dve izmene:

1. **Kapija ide iznad grananja** — `KnjizenjeSme(blokIzabran, outPoruka)` se pita
   pre nego što se uopšte gleda ima li izbora. Parametar se prima namerno iako se
   ne koristi: tvrdi se da odluka od njega **ne sme** da zavisi, i test to meri za
   obe vrednosti (sabotaža `dok-izbor-zaobilazi-kapiju`).
2. **`EH` prazni kombo i niz ID-jeva** — delimično napunjena lista je gora od
   prazne, jer izgleda kao potpuna.

Usput je `vba_check` uhvatio grešku u prvoj verziji tog čišćenja: `LogErr` je
stajao posle `On Error GoTo 0`, koji **resetuje `Err`**, pa ne bi upisao ništa
(`MRTAV_LOG`). Log sada ide pre čišćenja.

### 16.3 Seam-ovi

Tvrdo gejtovani, po ugledu na `frmBankaImport.BiTest*`:

| Seam | Čemu služi |
|---|---|
| `DokTestSetBlokUcitanost` | postavi ishod učitavanja — pad se drugačije ne može izazvati bez lomljenja šeme |
| `DokTestBlokSme` | **odluka koju handler stvarno čita**, ne njena kopija |
| `DokTestBlokPoruka` | šta bi operater video |

Forma se u testu **ne prikazuje**: `frmDokumenta` nema `UserForm_Initialize`, pa je
`New frmDokumenta` jeftin i ne čita nijednu tabelu — isti razlog kao §11.1. Težak
posao je u `UserForm_Activate`, koji ide tek na `.Show`.

### 16.4 Šta test tvrdi

`T_LegacyDok_PadListeBlokovaNijeAvans` (119):

- pad učitavanja **zaustavlja** knjiženje avansa, i operater dobija objašnjenje;
- u poruci stoji i **šta** je puklo;
- uredno učitana lista **pušta** avans, bez poruke.

Dve sabotaže: `dok-pad-liste-blokova-prolazi` i `dok-kapija-blokova-presiroka` —
druga postoji jer je kapija šira od kvara isto tako greška.

### 16.5 Šta ovo NE pokriva

**Sam `btnUnosOMUlaz_Click` se u testu ne izvršava** — pokriveno je *pravilo* i to
da ga handler zove (jedan red, proveren čitanjem). Isto ograničenje kao §11.5.

**Da `FillOpenOtkupi` stvarno beleži pad nije mereno**: test postavlja stanje kroz
seam, pa zaobilazi samu funkciju. Izazvati pravi pad čitanja iz testa značilo bi
lomiti šemu.

To je i razlog zašto je rupa iz §16.2.1 preživela prvu verziju: `2 / 2 DOKAZANO`
je dokazivalo **„ako je `m_BlokoviOk = False`, kapija radi"** — a ne „svaki stvarni
put neuspeha zaista postavlja `False`". Razlika je tiha i skupa.

Ni `RequireTable` u `FillOpenOtkupi` **nije izmeren** iz istog razloga. Ono što
jeste: samo pravilo `RequireTable` ima svoje testove i sabotažu
(`banka-uvoz-nema-tabele-je-prazna`), pa je nemereno ostalo **da ga ova funkcija
zove** — jedan red, proveren čitanjem. Isti oblik kao §11.5.

Zato je scenario „tabela nedostaje" u smoke listi ovog posla **obavezan**, ne
opcion.

### 16.6 Nalaz sa strane: `vba_check` ne vidi nedeklarisanu modul-promenljivu

Prvi patch je pao na drugom paru zamena, a skripta piše fajl tek kad **svi**
prođu — pa je i prvi deo otkotrljan. Drugi patch je zatim upisao kod koji koristi
`m_BlokoviOk`, dok deklaracije nije bilo.

`vba_check` je na to rekao **`cisto`** (`rc=0`), iako `.frm` ima `Option Explicit`.
Pad se video tek kao `run_vba` koji visi, sa Excelom u `[break]` — dakle kroz
najskuplji mogući kanal.

Provereno namerno: uklanjanje `Private m_BlokoviOk As Boolean` iz forme i dalje
daje `vba_check: cisto`. To je rupa u checkeru, zapisana ovde i ostavljena kao
zaseban posao — nije deo ovog PR-a.

> **ZATVORENO** u `v2.80.0`, pravilom `NEDEKLARISAN`. Isti test je i dokaz:
> nad zdravim `frmDokumenta.frm` daje **0** nalaza, a čim se ta jedna
> deklaracija ukloni — **1**, i imenuje `m_BlokoviOk`. Detaljno:
> `docs/engineering/postmortems/2026-08-verifikacija.md` §13.

---

## 17. Ista greška je stajala na još tri mesta (`v6-ui-184`)

§16 je zatvorio **jednu instancu**, ne klasu. Oblik je uvek isti:

```
Fill*  guta pad  ->  prazna lista  ->  grana bira AVANS
```

Traženje ostalih instanci polazi od pitanja „ko još iz **praznog** polja
zaključuje **tip novca**". Odgovor su bila tri mesta, sva tri nad novcem:

| Gde | Šta je postajalo | Kako se pad video |
|---|---|---|
| `frmDokumenta.FillOpenFakture` (F6) | `NOV_KUPCI_AVANS` | `LogErr` + `Clear` |
| `modOtkupUI.FillOpenBlokovi` (F5) | `NOV_VIRMAN_AVANS_KOOP` | **`Debug.Print`** |
| `modOtkupUI.FillOpenFakture` (F6) | `NOV_KUPCI_AVANS` | **`Debug.Print`** |

### 17.1 Nova ljuska je bila tiša od legacy-ja

`modOtkupUI` je imao **osam** `Debug.Print ... PAO` rukovalaca i **jedan**
`LogErr`. Immediate prozor u pogonu niko ne gleda, pa pad koji menja **tip
knjiženja** nije ostavljao ni trag u logu — legacy je bar pisao `LogErr`.

Dva od tih osam stoje tačno ispred novčane grane i ovim poslom su prešla na
`LogErr`. Preostalih šest nisu dirana: nisu deo ove klase, i menjati ih usput
značilo bi izmenu koju ništa ne meri.

### 17.2 Nedostajuća tabela je prolazila i ovde

`GetOpenFakture` vraća `Empty` i kad `tblFakture` nema — isto što i
`GetOpenOtkupi` za `tblOtkup` (§16.2.1). Sva tri mesta su zato dobila
`RequireTable` **pre** čitanja.

### 17.3 Kapija je široka tačno koliko i kvar

Ovo je deo koji nosi najviše tvrdnji, jer je i najlakše preterati. Kapija se
**ne pita**:

- kad je iznos nula — unos same ambalaže nema odluku faktura/avans, pa lista
  koja ga se ne tiče ne sme da ga zaustavi;
- kad je primalac otkupno mesto (F5) — blok se bira samo kooperantu, isplata OM-u
  ide drugim tipom novca i listu blokova ne dodiruje;
- u režimima koji te liste nemaju (F1–F4, F7).

Svako od ta tri ograničenja ima svoju sabotažu, jer je svako **tvrdnja**, a ne
izuzetak. Bez njih bi jedan pad čitanja zaustavio i posao koji nikad ne bi bio
pogrešno proknjižen.

### 17.4 Zašto pravilo NIJE izvučeno u zajednički modul

Legacy forma i ljuska drže **svoje kopije** — `frmDokumenta.ListaSme` i
`modOtkupUI.LjuskaListaSme`. To je namerno, po pravilu iz §5: dok novi UI ne ume
sve, obe kopije poslovne logike postoje, a legacy se ne vezuje za nove module.

Unutar **svakog** domaćina rule je izvučeno jednom: `BlokIzborSme` i
`FakturaIzborSme` zovu isti `ListaSme`, a u ljusci obe grane zovu isti
`LjuskaListaSme`. Podeljen je i **oblik poruke**, da operater u obe forme vidi
istu rečenicu.

Cena: dve kopije umesto jedne. Dobitak: sabotaža po mestu, pa svaka obara **svoju**
tvrdnju — kad bi tekst bio jedan zajednički, jedna sabotaža bi obarala četiri
pravila i nijedno ne bi bilo izmereno posebno.

### 17.5 Kapija ljuske stoji nad REČNIKOM, ne nad formom

`modOtkupUI.NovacListaSme(p, outPoruka)` čita isti rečnik koji ide u
`modScrDokumenti` → `modNovacUnos`. Dve posledice:

1. test ga zove bez forme — nema `.Show`, nema čitanja tabela;
2. režim se čita iz `p("rezim")`, dakle iz vrednosti po kojoj se **stvarno** bira
   `Save` rutina, a ne iz `ActiveMode`.

Druga tačka traži preciznost, jer je prva verzija ovog teksta tvrdila više nego
što stoji. `p("rezim")` **nije** nezavisan izvor istine — `SkupiPolja` ga pravi kao
`modeKey(ActiveMode)`, dakle kao snimak istog `ActiveMode` po kome `Fill*` odlučuje
da li će listu uopšte čitati.

Razlog za `p("rezim")` je uži i tačan: kapija treba da odlučuje nad **istim
payload-om** koji se predaje `Save` putu, a ne nad stanjem koje se do tada moglo
promeniti. Njihova međusobna usklađenost ostaje **integraciona pretpostavka**
sinhronog toka. Test pravi sintetički `p("rezim")` nezavisno od `ActiveMode`, pa
meri `NovacListaSme` izolovano — ali tu nezavisnost u stvarnom UI putu **ne meri**,
i ne treba je ni čitati kao dokazanu.

### 17.6 Šta testovi tvrde

`T_LegacyDok_PadListeFakturaNijeAvans` (120):

| Tvrdnja | Sabotaža |
|---|---|
| pad učitavanja **zaustavlja** avans kupca | `dok-pad-liste-faktura-prolazi` |
| ...uz poruku u kojoj stoji šta je puklo | — |
| **ni izabrana faktura** ne prolazi kad je pad | `dok-faktura-izbor-zaobilazi` |
| unos **bez novca** NE staje zbog te liste | `dok-uplata-kapija-siri-se-na-ambalazu` |
| uredna lista **pušta** uplatu | `dok-kapija-faktura-presiroka` |

`T_Ljuska_PadListeNovcaNijeAvans` (121):

| Tvrdnja | Sabotaža |
|---|---|
| pad liste blokova zaustavlja isplatu kooperantu | `ljuska-pad-liste-blokova-prolazi` |
| isplata **otkupnom mestu** ne zavisi od te liste | `ljuska-kapija-hvata-i-otkupno-mesto` |
| režim bez tih listi kapiju **ne oseća** | `ljuska-kapija-hvata-sve-rezime` |
| unos **bez novca** NE staje zbog liste | `ljuska-kapija-hvata-i-bez-novca` |
| pad liste faktura zaustavlja uplatu kupca | `ljuska-pad-liste-faktura-prolazi` |

**Redosled tvrdnji je deo konstrukcije.** `AssertEq` staje na prvoj paloj, pa
sabotaža koja kapiju širi na sve režime usput obara i `padF6` — zato `drugiRezim`
mora doći **pre** njega. Bez toga bi test pao po imenu tuđe tvrdnje (zamka 6).

Iz istog razloga `DokTestUplataPoruka` zove `FakturaIzborSme` **direktno**, a ne
kroz `UplataSme`: inače bi sabotaža kapije oborila i tvrdnju o poruci, pa bi prva
pala tvrdnja bila pogrešna (zamka 5). Isti oblik kao `DokTestBlokPoruka`.

### 17.7 Šta ovo NE pokriva

**Nijedan od dva handlera se u testu ne izvršava** — ni `btnUnosIzlaz_Click` ni
`CommitDokument`. Pokriveno je *pravilo* i to da ga handler zove iznad grananja
(po jedan red, proveren čitanjem). Isto ograničenje kao §11.5 i §16.5.

**Da `Fill*` stvarno beleži pad nije mereno** ni na jednom od tri mesta: test
postavlja stanje kroz seam. Isti razlog zbog kog je rupa iz §16.2.1 preživela.

**Ni tri nova `RequireTable` poziva nisu izmerena** — sabotaža koja ih ukloni bila
bi mrtva u katalogu.

**Da liste u ljusci uvek budu napunjene pre potvrde** nije mereno. Kapija se
oslanja na to što `FillOpenBlokovi` / `FillOpenFakture` idu na svaku promenu
režima i partnera, pa je „nikad učitano" nedostižno pre nego što partner postoji —
a bez partnera se ni jedna od dve grane ne bira. Pročitano, ne izmereno.

### 17.8 Nalaz sa strane: mrtva sabotaža `ljuska-rez-bez-potvrde`

`dokaz.py` je prijavio da `ljuska-rez-bez-potvrde` (`modUiKit.bas`, potvrda
`Font.bold` u tri pokušaja) **ne obara ništa** — sidro se i dalje nalazi, ali
`T_ZonaAgro_PrekidacRezimaZadrzavaBoju` tu invarijantu ne meri.

Provereno da je zatečeno: isti nalaz daje i `origin/main`, a ovaj PR ne dira ni
`modUiKit.bas` ni taj test. Ostavljeno kao zaseban posao — ista klasa koju je
#227 čistio.

---

## 18. Filter storniranih je znao ŠTA filtrira samo napola (`v2.78.0`)

Nastavak iste klase iz §16 i §17, ali izvan UI-ja i sa najširim dometom u
projektu: **183 poziva**, 21 tabela.

```vb
colStorno = GetColumnIndex(tblName, COL_STORNIRANO)
If colStorno = 0 Then
    ExcludeStornirano = data      ' vraca SVE, i stornirane
    Exit Function
End If
```

### 18.1 Zašto je ovo gore od §16 i §17

Tamo je novac dobijao **pogrešan tip** (avans umesto razduženja). Ovde storniran
dokument dobija **pogrešno postojanje**: `GetOpenFakture` bi otkazanu fakturu
vratio kao otvorenu, pa bi uplata otišla na dokument koji više ne važi. Isto za
`GetOpenOtkupi` i otkupne blokove.

Domenska invarijanta koju to krši stoji zapisana u `docs/DOMEN/README.md`:
*„storniran red ostaje u tabeli i izlazi iz svih agregata"*. Fail-open je značio
da iz agregata **ne** izlazi, i to ćutke.

### 18.2 Nula je imala dva značenja

Isti oblik kao §16.2.1, samo jedan nivo niže:

| `colStorno = 0` znači | Ispravan ishod |
|---|---|
| tabela storno pojam **nema** (matični podaci) | prolaz — filtrirati se nema šta |
| kolona **nije nađena** (drift) | pad — filter se NIJE izvršio |

Kod tu razliku nije znao, pa je uzimao blažu. Izmerio sam šemu
(`tools/dump_schema.py` nad `tests/fixtures/otkup_test.xlsm`): od 38 tabela
**17 nosi** `Stornirano`, 21 ne. Od 21 tabele koja stvarno prolazi kroz
`ExcludeStornirano`, **15 mora** imati kolonu, a šest su matični podaci.

### 18.3 Registar je deklaracija, ne snimak

`modSchemaGuard` drži `STORNO_TABELE` i `BEZ_STORNA`. Ključna razlika: to **nije**
snimak zatečene sveske. Šema je izvor istine po instalaciji (CLAUDE.md §3), pa
tabela iz prvog spiska koja nema kolonu znači **drift** — i tada se pada, umesto
da se tiho ne filtrira.

Pad ide kroz postojeći `RequireColumnIndex`, bez novog koda za grešku: njegova
poruka od #229 već razlikuje *„kolone nema"*, *„tabele nema"* i *„čitanje je
puklo"*, i kaže da li je traženu kolonu **videla** u svežem prolazu. Za drift je
baš to razlika koja se traži.

**Dve deklaracije šeme ne smeju da se raziđu.** `modProductionHealthCheck` je za
`tblFakturaStavke` tražio sve kolone osim `Stornirano`, dok registar tu kolonu od
sada zahteva — pa bi health check javljao zdravo stanje za svesku nad kojom
`ExcludeStornirano` pada. Kolona je dodata i tamo: bolje da se vidi pre rada nego
u radu.

### 18.4 Treće stanje: tabela koju registar ne poznaje

Za nju **nema tačnog odgovora** — ni pad ni prolaz nisu opravdani — pa mora da
padne, i to na **dva** mesta.

**Statički**, gde je najjeftinije: novo `vba_check` pravilo `STORNO_REGISTAR` traži
da svaki `ExcludeStornirano(..., TBL_X)` imenuje tabelu iz jednog od dva spiska.

**I u izvršavanju**, jer statička provera namerno preskače pozive sa promenljivim
imenom tabele — a takvih stvarno ima (`modIntegritet.CollectBrojZbirne`,
`modDokumenta.SumByBroj` i slični, koji `tbl` primaju kao argument). Bez runtime
kapije bi `TabelaNosiStorno = False` opet značilo **dve stvari**: „eksplicitno
`BEZ_STORNA`" i „niko je nije klasifikovao". To je ista bolest zbog koje je ceo
posao i nastao, samo pomerena jedan nivo dalje — i prva verzija ovog PR-a je baš
tu grešku ponovila: `StornoRegistarZna` je bio **dodat, pa nekorišćen**.

`RequireStornoKlasifikaciju` sada stoji kao prva naredba u `ExcludeStornirano`.

**Preostala granica:** statička provera i dalje ne vidi promenljivo ime tabele —
ali runtime ga sada hvata, pa fail-closed ugovor ne zavisi od nje.

### 18.4.1 Kapiju je jedan red IZNAD nje mogao da poništi

Najozbiljniji nalaz iz review-a, i nije bio teorijski. `modKarticaDetalji`:

```vb
Private Function PrijemnicaBrojZaOtpremnicu(ByVal otpID As String) As String
    On Error Resume Next
    ...
    d = ExcludeStornirano(d, TBL_PRIJEMNICA)   ' <- pukne
    If Not IsArray(d) Then Exit Function       ' <- d je jos ORIGINALNI niz
```

Kad kapija padne, greška se proguta, **dodela se ne izvrši**, i `d` ostaje
nefiltriran — pa stornirana prijemnica ide dalje kao živa. Fail-closed primitiv
koji pozivalac proguta nije fail-closed.

Revizija svih 183 poziva dala je **jedan** takav produkcioni pozivalac (plus dva
namerna hvatanja u testu 122). Prepravljen je na `On Error GoTo EH`: nema broja je
bolje nego pogrešan broj.

Da revizija ne bi zavisila od toga da se neko seti da je ponovi, uvedeno je i
pravilo `STORNO_PROGUTAN`: u **produkcionom** modulu je `ExcludeStornirano` pod
aktivnim `On Error Resume Next` uvek nalaz.

**Prva verzija tog pravila merila je pogrešnu stvar**, i to je nalaz iz review-a
koji vredi zapisati. Izuzetak je bio „sledeća naredba pominje `Err.`" — što ne
dokazuje da je greška **obrađena**. Kroz njega je prolazio baš kvar koji pravilo
sprečava:

```vb
On Error Resume Next
d = ExcludeStornirano(d, TBL_PRIJEMNICA)   ' pukne, dodela se ne izvrsi
Err.Clear                                  ' obrise DOKAZ, checker zelen
If IsArray(d) Then ...                     ' d je jos NEFILTRIRAN
```

`Debug.Print Err.Number` je prolazio isto. Dokazivati statički da je `Err` stvarno
obrađen znači pisati mini analizu toka — a nije potrebno: revizija je pokazala da
u produkciji **nema nijednog** legitimnog takvog poziva. Zato pravilo nema
heuristiku, a namerno hvatanje sme samo u **test modulu**, gde je test sam sebi
dokaz (pao bi da greške nema).

Predikat testnog modula ima svoju zamku: `modTestMode.bas` **nije** test nego
produkcijski `IsTestMode()`, uprkos imenu. I to ima svoj slučaj u self-testu.

Pravilo je pušteno nad **`origin/main` verzijom** `modKarticaDetalji.bas` i
prijavilo je tačno taj poziv, na liniji 316. To je jači dokaz od fixture-a: hvata
kod koji je stvarno bio u repou.

**Granica:** hvata se samo **direktan** poziv pod aktivnim `Resume Next`. Ako `A`
zove `ExcludeStornirano` bez rukovaoca, a `B` zove `A` pod `Resume Next`, greška
se penje do `B` — to traži analizu celog grafa poziva i ne pokušava se.

### 18.5 Dvosmerni dokaz checkera

Pravilo menja **checker**, pa po CLAUDE.md §5 nosi dokaz u oba smera. Dva
sabotiranja razdvajaju dva različita kvara:

| Sabotaža | Šta padne |
|---|---|
| `check_storno_registar` vraća `[]` | **3** slučaja + CLI slučaj |
| `main()` je više ne zove | **samo** CLI slučaj |
| `check_storno_progutan` isključen iz `check_file` | **5** slučajeva |
| izuzetak za test modul uklonjen | **samo** slučaj „isti kod u TEST modulu" |

Poslednji red je bitan: i sam **izuzetak** ima crveni dokaz, ne samo pravilo.
Izuzetak koji nikad nije pokazan kao potreban je isto što i pravilo koje nikad
nije pokazano kao crveno.

Druga je ona koju self-test bez CLI prolaza ne bi video — „provera grize, ali nije
priključena". To je ista klasa greške koju komentar u `vba_check` već opisuje kao
placebo test.

### 18.6 Self-test je pao na MOM fixture-u, i to je bio prvi nalaz

Slučaj „prelomljen poziv se i dalje vidi" je pao pre nego što je bilo šta
sabotirano. Uzrok nije bio parser nego fixture: pisao sam `\r\n` u string koji se
na disk piše sa `newline="\r\n"`, pa je svaki red dobio prazan red iza sebe — a
nastavak reda (` _`) se onda spajao **sa prazninom**.

Vredi zapisati jer je isti oblik kao zamka iz §14: fixture koji preskače prazan
red menja ono što se meri, a da testu ništa ne izgleda sumnjivo.

### 18.7 Šta testira test 122

| Tvrdnja | Sabotaža |
|---|---|
| dokument tabela **je** u registru | `storno-registar-prazan` |
| matični podaci **nisu** | `storno-registar-hvata-i-maticne` |
| nedostajuća kolona **pada** i imenuje kolonu | `storno-filter-nedostajuca-kolona-prolazi` |
| ...i kad je tabela **prazna** | `storno-filter-prazna-tabela-preskace-kapiju` |
| **neklasifikovana** tabela pada | `storno-nepoznata-tabela-prolazi` |
| tabela bez storno pojma **prolazi** | `storno-filter-hvata-i-tabele-bez-storna` |

Nula se izaziva kroz keš kolona (`KesKoloneTestSet`), istim putem kao test 117.
Time se **ne tvrdi** da je keš uzrok ijednog pada iz pogona — tvrdi se da
`ExcludeStornirano` na nulu više ne propušta nefiltrirano.

### 18.8 Šta ovo NE pokriva

**Da spisak od 15 tabela odgovara PRODUKCIJI nije izmereno.** Izmeren je fixture;
produkciona sveska može imati drugu šemu. Ako u njoj neka od tih 15 nema kolonu,
ovaj PR pretvara tihu grešku u **glasan zastoj** — što je namera, ali je i razlog
zašto smoke nad kopijom prave sveske ovde nije formalnost.

**Promenljivo ime tabele** i dalje izmiče **statičkoj** proveri (§18.4), ali ga
runtime kapija hvata, pa ugovor od nje ne zavisi.

**Transitivni `Resume Next`** — v. §18.4.1: pozivalac pozivaoca se ne prati.

**Prazna tabela više ne preskače kapiju** (to je bio nalaz iz review-a):
klasifikacija i provera kolone stoje **iznad** `If IsEmpty(data)`. Ugovor je
„tabela iz `STORNO_TABELE` bez kolone znači drift", a ne „drift se prijavljuje
samo dok tabela ima redova". Fixture ima redove, pa tu granu prva verzija testa
nije ni takla.

**`mColCache` i dalje pamti nulu.** Jedna neuspela pretraga kolone ostaje
zapamćena za ceo `BeginTableCache` prozor, pa bi se nova kapija posle nje držala
zatvorenom do kraja prozora. To je i dalje bolje od tihog prolaza, ali nije
rešeno — a razlikovanje „kolone stvarno nema" od „čitanje je puklo" traži oslonac
na broj greške koji **nisam izmerio**, pa ga ne uvodim na pretpostavku. Ostaje
otvoreno, kao i pre ovog posla.

---

## 19. Kapija je ušla u primitiv (`v2.84.0`)

Katalog je ovo tražio rečima „tu bi jednog dana trebala centralna kapija umesto
zaštite po call-site-u". Urađena je **polovina**, i to je merena polovina.

### Šta je bilo

`RecalculateZbirnaFromOtpremnice_TX` mutira **sve** `tblZbirna` redove sa datim
brojem. Broj nije identitet — dva vlasnika mogu nositi isti. Zaštita je stajala
isključivo po call-site-u: `ZbirnaBrojJeDvosmislenIkad` na **šest** mesta u
`modStornoFlow`, dok sam primitiv nije imao **nijednu**. Nov pozivalac je bio
bezbedan tek ako se autor kapije seti.

### Mereno stanje pre izmene

Sonda na kraju niza testova, nad `ZB-TEST-KASK`:

| | |
|---|---|
| vlasnika **IKAD** | 2 |
| vlasnika **aktivnih** | 1 |
| `RecalculateZbirnaFromOtpremnice_TX` vraća | **True** — prolazi i mutira |

Drugi red je ono što ovaj slučaj čini vrednim: aktivan je **jedan**, pa kapija
koja broji samo aktivne ovde **ne bi okinula**. Test time dokazuje i da se mora
brojati IKAD, bez posebne sabotaže za tu zastavicu.

### Šta je urađeno

Kapija je ušla u primitiv i broji IKAD. Uz nju je u `modStorno` izdvojen **jedan
račun** za obe kapije (`BrojVlasnikaPoBroju`), pa se aktivna i IKAD varijanta ne
mogu razići — isti potez kao `_pogodaka` u `v2.82.0`, i iz istog razloga.

Redosled po `CLAUDE.md` §2 — test 124 pisan i pušten **pre** izmene:

```
FAIL T_RekalkZbirne_KapijaJeUPrimitivu -- rekalkulacija po dvosmislenom broju
     ne prolazi kroz sam primitiv -- ocekivano [False], dobijeno [True]
```

Posle izmene pun set: `RunAllTests 124/0`, Banka 196/0, svih jedanaest suite-ova
OK. To je i odgovor na jedini pravi rizik — stroža kapija ne obara nijedan
postojeći tok.

### Šta NIJE urađeno, i zašto

**`ReassignPrijemnicaToZbirna_TX` nije dirana.** Njena kapija i dalje broji samo
aktivne. Probano je i izmereno: prebacivanje na IKAD ostavlja **ceo set zelen** —
što znači da nijedan test tu razliku ne vidi. Druga sonda je pokazala i da poziv
bez generacije na dvosmislen cilj već vraća `False`, ali **nije izolovano zbog
čega** (moguće raniji uslov, ne kapija).

> **Zatvoreno na strani CILJA u `v2.85.0`** (§20). Na strani **izvorne
> prijemnice** kapija ostaje po AKTIVNIMA, i to **namerno** — §21 objašnjava
> zašto šira zabrana tamo nije popravka nego šteta.

Izmena ponašanja bez testa koji je meri je tačno ono što `CLAUDE.md` §2 zabranjuje,
pa je ostavljena otvorena umesto da se progura kao „i to je popravljeno".

**Šest kapija po call-site-u je ostalo.** Katalog je pisao „umesto", ali bi
uklanjanje bilo **nazadovanje u dijagnostici**: te kapije staju pre transakcije i
kažu razlog („Zamena bi prevezala decu"), a centralna staje iznutra i pozivaocu
daje samo neuspeh. Zato je centralna **mreža ispod**, a ne zamena — i katalog je
gore ispravljen da to više ne obećava.

---

## 20. Cilj-zbirna se birala po redu koji je slučajno prvi (`v2.85.0`)

### Šta je bilo

`ReassignPrijemnicaToZbirna_TX` bez generacije bira cilj ovako:

```vb
tId = LookupValue(TBL_ZBIRNA, COL_ZBR_BROJ, targetBrZbirne, COL_ZBR_ID)
If IsEmpty(tId) Then Exit Function                    ' zbirna ne postoji
tStor = NzToText(LookupValue(TBL_ZBIRNA, COL_ZBR_BROJ, targetBrZbirne, COL_STORNIRANO))
If UCase$(Trim$(tStor)) = "DA" Then Exit Function     ' cilj-zbirna stornirana
```

`LookupValue` vraća **prvi** pogodak i ne gleda ni identitet ni storno. Posle
storna jednog vlasnika prvi red sa tim brojem može biti storniran dok pod istim
brojem stoji **aktivna** zbirna — a tada prevezivanje **tiho** stane: bez poruke,
bez loga, samo `False`.

Kod je hazard i sam opisivao („LookupValue po broju uzima prvi pogodak"), ali
**samo za granu sa generacijom**. Fallback grana je ostala po broju.

### Popravka

Dva pogrešna pitanja („postoji li ijedan red" + „da li je **prvi** storniran")
zamenjena su jednim tačnim — **brojem aktivnih vlasnika pod tim brojem**:

```vb
Dim aktivniVlasnici As Long
aktivniVlasnici = VlasniciPoBroju(TBL_ZBIRNA, COL_ZBR_BROJ, targetBrZbirne, _
                                  SRC, False, _
                                  Array(COL_ZBR_VOZAC, COL_ZBR_KUPAC)).count
If aktivniVlasnici = 0 Then Exit Function
```

Nula znači „nema aktivnog cilja", a više od jedan hvata kapija ispod — koja ostaje
zbog svoje poruke.

### Prva verzija popravke je uvela novu grešku

Prvo sam napisao `If Not ZbirnaPostoji(targetBrZbirne) Then Exit Function`. Pita
pravu stvar, ali **drugim poređenjem**:

| | poređenje |
|---|---|
| `ZbirnaPostoji` | `StrComp(..., vbTextCompare)` — bez obzira na veličinu slova |
| `VlasniciPoBroju` (iza kapije) | `Trim$(...) = Trim$(broj)` — **tačno** |

Sa `zb-test-kask` umesto `ZB-TEST-KASK` postojanje bi reklo **da**, kapija bi
videla **nula** vlasnika — a ona hvata samo `n > 1`, pa bi propustila — i u
`tblPrijemnica` i `tblPaletaStavka` bi se upisala **labela pozivaoca** umesto one
iz tabele. Stari kod to nije dopuštao, jer je `LookupValue` poredio tačno; dakle
regresija koju je uvela **sama popravka**.

Našla je recenzija. Popravljeno tako što postojanje i vlasništvo idu kroz **isti**
račun, pa ne mogu da govore o dve različite stvari — isti obrazac kao
`BrojVlasnikaPoBroju` u §19 i `_pogodaka` u `v2.82.0`.

Uz to ide i regresija u testu 125: poziv sa `LCase$(FX_ZBIRNA_KASK)` mora da vrati
`False`, a u prijemnici mora da ostane kanonska labela. Sabotaža
`cilj-zbirna-case-mesano` vraća baš prvu verziju i obara tu tvrdnju po imenu.

> Ovo je bio **drugi** put u istom PR-u da merenje obori moju procenu — prvi put
> vozilo testa, drugi put semantika poređenja.

### Merenje je oborilo prvu dijagnozu — i to je glavni nalaz

Test 125 je napisan i pušten pre izmene, i pao je. Popravka je primenjena — i
**test je i dalje padao, isto**.

Uzrok: za vozilo sam uzeo `FX_PRIJ_BROJ`, a sonda je pokazala da taj broj ima
**dva aktivna reda i dva vlasnika**, pa poziv obara kapija na strani **prijemnice**
(`RequireJedanVlasnikPoBroju`) — pre nego što razrešenje cilja uopšte dođe do
izražaja. Prvobitno „tiho odbijanje" koje sam pripisao prvom redu bilo je, u tom
pozivu, nešto sasvim drugo.

| | |
|---|---|
| moja dijagnoza | razrešenje cilja po prvom redu |
| šta je stvarno blokiralo taj poziv | kapija po vlasnicima prijemnice |
| koliko razloga je bilo | **dva**, a pokrio sam pogrešan |

Tek posle treće sonde — koja traži prijemnicu sa **tačno jednim** vlasnikom —
nađeno je vozilo (`FX_PRIJEMNICA_STALE`, već na toj zbirnoj, pa poziv prolazi
putanju a ne pomera ništa semantički). Sa njim sabotaža `cilj-zbirna-po-prvom-redu`
obara test po imenu, a bez nje je ceo set zelen.

Pouka je stara i skupa: **zelena popravka nad pogrešnim vozilom ne dokazuje
ništa, a crvena isto tako ne optužuje ono što misliš.**

### Druga runda recenzije: test je cementirao ono što §19 zabranjuje

Prvi popravljeni test je nad `ZB-TEST-KASK` tvrdio da prevezivanje **prolazi** —
a to je isti broj za koji test 124, neposredno pre njega, dokazuje `ikad = 2`,
`aktivnih = 1`. Testom bi time bilo ozvaničeno tačno ono što je `v2.84.0` uveo
kao zabranu: mutacija po golom broju nad istorijski dvosmislenim ciljem.

Nije bio teorijski dug. Recovery panel poziva baš tu putanju:

```vb
' frmDokumenta.frm:5603
If ReassignPrijemnicaToZbirna_TX(brp, brz) Then
```

Dva argumenta — bez generacije, bez `ZbirnaBrojJeDvosmislenIkad` iznad, bez ičega.
Putanje iz `modStornoFlow` imaju kapiju iznad sebe; **ova nema**. A veza koja se
upisuje je gola labela (`COL_PRJ_BROJ_ZBIRNE`, `COL_PALS_BROJ_ZBIRNE`), pa bi dete
završilo vezano za broj koji nose dva vlasnička toka.

Zato je u tu granu dodat `RequireJedanVlasnikIkadPoBroju`, a test 125 sada tvrdi
**suprotno**: `ikad = 2 / aktivnih = 1` → prevezivanje bez generacije **ne
prolazi**. Time ovaj PR zatvara **ciljnu polovinu** rupe iz §19 umesto da je
cementira.

### Prvobitni kvar je subsumiran, i to se ne može izmeriti zasebno

Pitanje više nije „da li je **prvi** red storniran" nego „ima li **aktivnog**
cilja" i „da li je broj **ikad** pripadao dvama vlasnicima". Stari oblik time
nestaje kao zasebna grana — i njegova sabotaža bi bila **mrtva**, pa je uklonjena.

Da bi se prvobitni kvar dokazao zasebno, trebalo bi vozilo sa:

```
prvi red = storniran, aktivan red = postoji, IKAD vlasnika = 1
```

Izmereno nad fixture-om — takvog nema:

| broj | prvi red | IKAD | aktivnih |
|---|---|---|---|
| `ZB-TEST-KASK` | storniran | 2 | 1 |
| `ZB-TEST-OLDU` | storniran | 1 | **0** |
| `ZB-TEST-STORNO` | storniran | 1 | **0** |

A napraviti ga ne mogu iz izvora: redovi fixture-a žive u `.xlsm`, ne u repou —
`src-vba/` te redove samo **žigoše** generacijom, ne kreira ih.

Zato se ne tvrdi da je taj oblik izmeren. Umesto njega mere se **dve kapije koje
su ga zamenile**, svaka svojom sabotažom:

| Sabotaža | Pada, po imenu |
|---|---|
| `cilj-zbirna-kapija-samo-aktivni` | `istorijski dvosmislen broj ne prolazi bez generacije` |
| `cilj-zbirna-bez-provere-postojanja` | `prevezivanje na broj bez ijedne aktivne zbirne ne prolazi` |
| `cilj-zbirna-case-mesano` | `broj sa drugom velicinom slova nije isti broj` |

### Šta ostaje otvoreno

`RequireJedanVlasnikPoBroju` na strani **prijemnice** (drugi poziv u istoj
funkciji) i dalje broji **samo aktivne**. Ciljna strana je zatvorena; ta nije.

> **Ne zatvara se.** `v2.86.0` (§21) je izmerio da je zaštita tamo **slojevita**
> i da bi jedna šira kapija na ulazu odbila legitiman oporavak.

**Order-dependency testova 124 i 125** ostaje test-dug: oba moraju biti poslednja
u nizu jer diraju fixture. Nije rešeno ovde, da se diff ne širi.

---

## 21. Šira zabrana koja je bila predložena kao popravka (`v2.86.0`)

Ovo izdanje je **povuklo** izmenu koju je isti PR prvo uveo. Zapis postoji zato
što je nalaz opšti: *konzervativnija* provera nije automatski bolja provera.

### Šta sam predložio

Kapija na ulazu u `ReassignPrijemnicaToZbirna_TX`, po uzoru na stranu cilja:

```vb
RequireJedanVlasnikIkadPoBroju TBL_PRIJEMNICA, COL_PRJ_BROJ, brPrijemnice, _
                               SRC, COL_PRJ_KUPAC
```

Obrazloženje je bilo simetrija: ako je broj **ikad** pripadao dvama kupcima,
izbor po broju pomera i tuđi red.

### Zašto je to bilo pogrešno

Kod ispod toga **to ne radi**. Izmereno čitanjem, red po red:

| Sloj | Šta ga štiti |
|---|---|
| zaglavlje prijemnice | u `targetRows` ulaze **samo aktivni** redovi — storniran tuđi dokument ne može ni da uđe u izbor |
| paletna stavka sa `PrijemnicaID` | odlučuje **identitet**: `pripada = docIds.Exists(pidS)` |
| legacy stavka **bez** `PrijemnicaID` | `brojDvosmislen` računa **IKAD** i puca uz `Err.Raise`, a cela transakcija se povlači |

Istorijska kapija, dakle, **već postoji** — tačno tamo gde je nužna, i samo tamo.
Moja bi je duplirala na ulazu i time zabranila slučajeve koje kod već bezbedno
razdvaja.

Poslovna cena nije teorijska: broj prijemnice je numerisan **po kupcu**, pa je
kolizija očekivana, a ne egzotika. Legacy panel u `frmDokumenta` (bez generacije)
bi posle te izmene odbio **potpuno rešiv** oporavak čim je isti broj nekad nosio
storniran dokument drugog kupca — i oterao operatera na ručni rad.

### Prvi test je bio kružan

Vredi zapisa isto koliko i sama izmena.

Test je tražio da poziv vrati `False`, a za cilj je uzeta zbirna na kojoj
dokument **već stoji**. Zatečeno „`True`" zato nije bio dokaz pogrešne mutacije
nego uspešan **no-op**. Sabotaža je onda vraćala kapiju i test je postajao crven —
čime je dokazano samo da je *nova politika potrebna novoj tvrdnji*, ne da bez nje
nešto pogrešno mutira.

> Mutacioni dokaz koji meri sopstvenu politiku umesto poslovne greške je zelen iz
> pogrešnog razloga — ista klasa kao §12, samo na nivou tvrdnje a ne teksta.

### Šta je stvarno urađeno

Kapija vraćena na aktivne, a test prepisan da meri **posledicu**:

```
3/150326:  PRJ-TEST-I1 storniran (KUP-TEST-1)   PRJ-TEST-I2 aktivan (KUP-TEST-2)
           obe na ZB-TEST-4
```

Prevezivanje na drugu, jednoznačnu aktivnu zbirnu mora da:

1. **prođe** — legitiman oporavak se ne odbija;
2. pomeri **aktivan** dokument;
3. **ne** dirne storniran dokument drugog kupca.

Sabotaža `prijemnica-izvor-i-stornirani` uklanja filter po storniranom, pa tuđi
dokument **stvarno** biva pomeren — i obara treću tvrdnju po imenu. To je dokaz
pogrešne mutacije, a ne prekršaja politike.

### Šta ovo ne pokriva

**Paletne stavke.** Pod tim brojem ih u fixture-u **nema nijedne**, pa se zaštita
po `PrijemnicaID` i fail-closed grana za legacy stavku bez ID-a **ne mere** ovim
testom. Rečeno kao neizmereno.

**Broj vlasnika nije broj dokumenata.** `VlasniciPoBroju` broji **vlasnike**, pa
isti kupac sa dva aktivna dokumenta pod istim brojem i dalje daje `1`. Nijedna od
kapija u ovoj funkciji taj slučaj ne hvata. Nalaz stoji otvoren.

---

## 22. Banka nalozi — šta je preneto (`v6-ui-185`)

Četvrti ekran **Faze E**. Red u registru (`modUiScreens.ScrRows`) je postojao
od `S3a` — stavka menija se do sada crtala prigušena jer modula nije bilo.
Ovim se piše modul koji taj red već očekuje; **registar se ne dira**.

Bitna razlika u odnosu na Banka uvoz: najskuplji deo migracije je ovde već bio
odrađen. `modBankaExportPregled` je od AUD-026 nosio deset javnih funkcija
(lista otvorenih blokova sa fail-closed identitetom, klamp, saldo kapija, CSV
writer, specifikacija, računi firme) — pa je ekran uglavnom **prikaz nad
postojećim računima**, a izdvajanja iz forme skoro da nije bilo.

### 22.1 Gde je šta završilo

| Legacy (`frmBankaExportPregled`) | Novo mesto |
|---|---|
| `LoadBlokovi` + `RenderListbox` | lista **NALOZI**, čitač `modBankaExportPregled.GetBlokIsplataForGrid` |
| `lstBlokovi` MultiSelect (čekiranje) | **korpa „U NALOZIMA"** + kolona oznake + traka u zoni + dvoklik |
| filter Kooperant / datum / stanica | **pretraga ljuske** (haystack: broj, kooperant, stanica, račun, OtkupID) |
| KPI traka (`RefreshTopKpis`) | `modBankaExportPregled.NalogeKpi` (jedan prolaz; avans pool po kooperantu) |
| `CollectIsplataBlokovi` (selekcija za izvoz) | `modBankaExportPregled.OdaberiBlokoveZaNaloge` |
| combo „Sa računa" (`PopulateRacunCombo`) | polje zone `scrBnRacun` (isti `BankaNalogRacuniCSV` + `BankaNazivZaRacun`) |
| `btnGenerisiCSV_Click` | dugme zone `scrBnCsv` → `GenerisiNalogeCSV` (nepromenjen) |
| `btnExport_Click` (PDF specifikacija) | dugme zone `scrBnSpec` → `PrintIsplataSpecifikacija` (nepromenjen) |
| `mBtnAvansBlok` („Primeni avans na blok") | radnja nad redom `bnavans` → `ApplyAvansToOtkup_TX` (nepromenjen) |
| `txtIsplatiti` override + `ClampOverridesToOpen` | **NIJE preneto** — v. 22.3 |
| `btnOsvezi_Click` | posao ljuske (`RefreshFromData`) |

Nove rutine za mrežu (ekran ne čita tabele sam): sve tri u
`modBankaExportPregled` i sve tri **nad** `BuildBlokIsplataList` — pa
nasleđuju njegove fail-closed kapije identiteta (dupli/prazan `OtkupID`,
dupli `KooperantID` obaraju čitanje; testovi T17–T20 banka suite-a).
`OdaberiBlokoveZaNaloge` je izdvojen iz `CollectIsplataBlokovi` obrazac
(legacy forma je čitala ListBox; sada pravilo „korpa ili svi + bez TR se
preskače i broji + `IsplatitiIznos = ZaokruziNovac(OtvorenIznos)` PRE praga"
živi na jednom mestu) — isto kao `PrijemnicaDostupna` u §8.1.

### 22.2 Tri odluke

**(A) Export CSV ULAZI u ekran.** Merilo iz §9.3 (uvoz nije ušao) ovde daje
suprotan odgovor po sve tri tačke: `GenerisiNalogeCSV` **vraća ishod** (putanju
fajla i `outOdbijeno` razlog — dugme ume da kaže „N naloga, ukupno X, fajl Y");
piše **jedan** fajl atomično, bez pomeranja tuđih fajlova i bez potrebe za
progresom; i legacy forma **ima** to dugme, pa bi ekran bez njega bio uži od
legacy-ja (§8.3). Finalna kapija ostaje u domenu: `BuildNalogCsvPayload` čita
svež saldo i `ValidateNalogSaldo` odbija ceo fajl. Potvrda pre upisa nosi broj,
ukupan iznos, račun (sa bankom), datum valute i broj preskočenih — ništa tiho.

**(B) Korigovani iznos po redu NE ulazi.** Legacy `txtIsplatiti` override je
prolazno stanje sa celim §8.6 aparatom (klamp na svaki reload, čišćenje na
promenu konteksta, kanal za značku), a mreža ljuske nema detail panel u kom bi
unos po redu živeo — polje u zoni vezano za „izabran red" bi zastarevalo na
svaki sort/stranu/filter. v1 zato izvozi **pun otvoren iznos** (identično
legacy ponašanju kad operater ništa ne kuca); delimična isplata se i dalje
radi u `frmBankaExportPregled`, koja ostaje operativna. Posledica na §8.6:
jedino prolazno stanje ekrana je korpa identiteta — i ona NAMERNO **ne nosi
iznos**, pa zastareo snimak ne može da naruči uplatu (iznos se čita svež u
trenutku izvoza).

> **Oborena na prvom smoke-u** — v. §22.9: „nekad se iznos ne isplaćuje u
> potpunosti" je stvaran tok, ne izuzetak. Delimična isplata je ušla kao
> radnja „Iznos…" (InputBox, presedan SEF komentara), sa zadatim iznosima u
> zasebnom rečniku koji održava **isti** klamp kao legacy override
> (`ClampOverridesToOpenDict`, jezgro izdvojeno iz `ClampOverridesToOpen`
> baš za ovo — dva pozivaoca, jedan račun). Deo odluke koji je ostao: polja
> za unos u zoni i dalje **nema** — unos vezan za „izabran red" ne živi u
> zoni nego u dijalogu radnje, pa sort/strana/filter nemaju šta da zastare.

**(C) Štampa specifikacije ULAZI.** Jedan poziv postojeće rutine
(`PrintIsplataSpecifikacija` → `modPrint`, režim `ISPLATA_SPEC_PRINT_MODE`),
isti izbor blokova i isti iznosi kao CSV. Ne može se verifikovati automatski —
ide na smoke checklistu. Režim `OFF` se **prijavljuje** („štampa isključena"),
jer bi klik bez ijedne poruke izgledao kao dugme koje ne radi.

### 22.3 Šta je namerno drugačije od legacy-ja

- **JEDNA lista.** Predložena druga lista RAČUNI je izbačena: sav njen sadržaj
  (računi firme + naziv banke) nosi combo „Sa računa" u zoni, pa bi lista bila
  pregled bez ijednog posla nad redom — §8.4 formalno prolazi (druga forma
  podatka), ali ne prolazi test vrednosti.
- **Čip „iznad praga" nije ušao.** Prag ne postoji ni u legacy formi ni u
  configu — uvođenje bi bilo novo poslovno pravilo, ne prelazak ekrana.
  Čipovi su: Sve · Ima račun · Bez računa · Avans (poslednji: kooperant ima
  neraspoređen avans — te blokove pre naloga treba vezati radnjom).
- **Multiselect → korpa identiteta** („U NALOZIMA"): radnje `bnadd`/`bniznos`/
  `bndel`, dvoklik prebacuje, kolona kvačice, traka u zoni (najnovije prvo +
  preliv se prijavljuje), KPI pločica. Korpa se **usklađuje sa svežom listom
  pri svakom čitanju** (`BnUskladiKorpu`): stavka čiji blok više nije otvoren
  izlazi uz poruku — isti razlog kao legacy klamp („tiho spuštanje bi operater
  lako promašio") — a živima se osvežava snimak za traku; od druge recenzije
  izlazi i blok koji je u međuvremenu **ostao bez računa** (izvoz bi ga
  ionako preskočio, ali traka, zbir i potvrda ne smeju da ga pokazuju kao
  spreman). **Prazan izbor NE izvozi ništa** — v. §22.10/R1: prvi ugovor
  („prazno = svi", preuzet od legacy „nema selekcije = svi") je recenzija
  oborila kao merge blocker; „svi" postoji samo kao izričita peta radnja
  **„+ Svi sa računom"** (`bnsve`, tačno `MAX_ACT` radnji — granica bazena
  koju test tvrdi).
- **Delimična isplata = radnja „Iznos…"** (v. §22.9): zadati iznos po bloku,
  legacy pravila unosa (cent-domen, > 0, nikad preko otvorenog; jednak
  otvorenom briše zadato), vidljiva kolona ISPLATITI uz OTVORENO, klamp
  zadatih pri svakom čitanju. Blok kome se zada iznos automatski ulazi u
  naloge; izbacivanje iz naloga briše i zadati iznos.
- **Blok bez tekućeg računa ne može u korpu** (`BnDodaj` odbija) — legacy je
  isto radio na check-u reda. U CSV ionako ne sme (nema primaoca); ovim se ne
  broje nalozi koji nikad ne nastanu.
- **Posle uspešnog izvoza korpa se prazni** — drugi klik ne sme tiho da
  napravi iste naloge još jednom.
- **Datum od/do i stanica filter nisu preneti kao polja.** Pretraga ljuske
  pokriva stanicu (i sve ostalo iz haystack-a); vremenski opseg u praksi ne
  sužava isplatu (plaća se sve otvoreno), a `BuildBlokIsplataList` filtere i
  dalje nosi za legacy formu.

### 22.4 Identitet

Identitet reda je **`OtkupID`**, u koloni prioriteta 4 (ključ
`OTKUI_HDN_OTKID`), čitan kroz `GridCell` — mape „prikaz → ID" nema. Broj
bloka NIJE identitet (jedinstven je samo po stanici — fixture drži isti broj
na dva otkupna mesta), a žiro račun dele svi blokovi istog kooperanta.

**Dvosmislen `OtkupID` ovde ne stiže do mreže.** Za razliku od §8.5/§9.6
(`IdIliPrazno` → prazan identitet po redu), domenski čitač
(`BuildBlokIsplataList`) na dupli/prazan `OtkupID` među otvorenima **obara
celo čitanje** (AUD-026, `ERR_ISPLATA_*`) — jer bi na ovoj putanji pogrešan
identitet značio nalog pogrešnom primaocu, a to pravilo postoji od ranije i
ima svoje testove (T17–T20). `IdReda` u ekranu ostaje kao poslednja linija
(prazan → odbij), ne kao prva.

Red **prenosi** i ono što radnje moraju da znaju a iz prikaza se ne vidi
jednoznačno (§8.5): `KooperantID` (avans se knjiži na taj par), „ima račun"
(prazna ćelija računa liči na kolonu koja se nije nacrtala) i avans saldo —
sve prioriteta 4.

### 22.5 Brojač i prolazno stanje

Značka menija broji **otvorene blokove** (podatak u tabeli, `NalogeKpi`);
svaka promena nastaje upisom, pa je `RefreshFromData` pokriva — privatan kanal
ka `OsveziNavBrojace` ne postoji (§9.7 obrazac, ne §8.6). Korpa se u znački
**ne broji**: ona je izbor za izvoz, a ne posao koji čeka — napuštanje ekrana
sa punom korpom ne gubi ništa (izvoz je eksplicitan, tihog knjiženja nema).
Neuspeh čitanja KPI-ja nije nula: poslednja poznata vrednost, a pre prve
`-1` = „ne znam" (ljuska crta `!`) — isto pravilo kao §9.

### 22.6 Šta NIJE preneto

- **`frmBankaExportPregled` se ne gasi i ne menja** — dve kopije žive namerno
  (§5, Faza B). `ValidateNalogSaldo`, `ClampOverridesToOpen`, CSV writer i
  cent-domen pravilo **nisu dirani**.
- **„Isplatiti" override** (delimična isplata) — v. odluku (B).
- **„Primeni avans (sel.)"** (batch nad čekiranim) — v1 nosi avans samo po
  redu; batch ostaje u legacy formi. Razlog: batch knjiženje traži zbirni
  izveštaj ishoda (ok/no-op/greška po bloku), a toast nosi jedan red.
- **Storno isplate / izvoda** nije ovde — posao ekrana Storno.

### 22.7 Fixture

`tblOtkup`/`tblKooperanti`/`tblNovac` do sada nisu imali nijedan slučaj za
ovaj ekran: nijedan kooperant nije imao tekući račun, pa bi svaka tvrdnja o
„sme u CSV" merila prazan skup. Dodato (uz razlog u komentaru):

| Red | Zašto postoji |
|---|---|
| `KOOP-TEST-1.TekuciRacun` | jedini kooperant SA računom — „ima račun" polovina svih tvrdnji |
| `KOOP-TEST-2/3` bez računa | „bez računa" polovina: blok se vidi, ne sme u naloge |
| `OTK-NAL-DELIM` + `NOV-NAL-DELIM` | delimično isplaćen blok (1000 − 400): otvoreno = ostatak, ne pun iznos |
| `OTK-NAL-STOR` | storniran blok sa „otvorenim" iznosom — ne sme u listu |
| `OTK-BIM-PLAC` (postojeći) | u celosti plaćen — ne sme u listu |
| `KOOP-TEST-1` sa ≥2 bloka + avans 1000 | vozilo za dedup avans pool-a po kooperantu |

Pet config ključeva ekrana se **pinuje** u `SEF_CONFIG`
(`BANKA_NALOG_RACUN_1..4`, `BANKA_NALOG_RACUNI`, šifra, svrha,
`ISPLATA_SPEC_PRINT_MODE=OFF`) — ista klasa kao `DEFAULT_SORTA_VOCA` u §8.10:
donor-zavisan config je već nosio „dva crvena ali nisu moja".

### 22.8 Verifikacija

Testovi **127–132** u `modTest`, **T22** u `RunBankaImportTestSuite`, i
**sedamnaest** sabotaža. U `RunAllTests` se 127–132 izvršavaju **pre**
124–126 (mutirajući testovi ostaju poslednji u nizu; redosled izvršavanja ne
mora da prati brojeve).

| Test | Šta meri | Sabotaža |
|---|---|---|
| `T_BankaNalozi_UgovorEkrana` | registar, JEDNA lista, tri radnje, granice bazena, prvi čip najširi, datum kao broj | `banka-nalozi-cip-sve-nije-prvi` |
| `T_BankaNalozi_IdentitetURedu_NeCrtaSe` | identitet po ključu kolone, prio 4; sve prenosne kolone van prikaza; delimičan blok = ostatak; storniran/plaćen nisu u listi | `banka-nalozi-identitet-vidljiv`, `banka-nalozi-red-ne-nosi-koopid` |
| `T_BankaNalozi_CipoviIKpiPratePravila` | čipovi particija „sve"; slaganje sa čitačem po svakom bloku; avans pool po kooperantu (uz dokaz da vozilo postoji); KPI posle greške | `banka-nalozi-cip-imarac-pusta-sve`, `banka-nalozi-kpi-avans-po-bloku`, `banka-nalozi-kpi-greska-je-nula` |
| `T_BankaNalozi_KorpaIIzvoz` | korpa po identitetu (prazan/bez TR/dupli se odbijaju), traka najnovije prvo + preliv, usklađivanje sa svežom listom, izbor za izvoz (korpa/svi/nepoznat), svež iznos u cent-domenu | `banka-nalozi-bez-racuna-u-naloge`, `banka-nalozi-prazan-id-ulazi`, `banka-nalozi-usklad-ne-cisti`, `banka-nalozi-izvoz-ignorise-izbor`, `banka-nalozi-izvoz-sirov-iznos` |
| `T_ZonaBankaNalozi_PoljaIRaspored` | zona se stvarno gradi i raspoređuje; combo je polje (`nm`+`nmT`); tvrdi se posle `Unload`-a | `banka-nalozi-zona-bez-dugmeta` |
| `T_BankaNalozi_IznosPoBloku` | pravila unosa zadatog iznosa (legacy `txtIsplatiti`), klamp pri čitanju, kolona ISPLATITI, izvoz nosi zadato, čišćenje uz korpu | `banka-nalozi-iznos-preko-otvorenog`, `banka-nalozi-citanje-ne-klampuje`, `banka-nalozi-izvoz-ignorise-iznos` |
| `T22_RacunUCsvJeExcelSafe` (banka suite) | goli 18-cifreni račun se u kolonama CSV-a kanonizuje u NBS 3-13-2; sve drugo netaknuto | `banka-csv-racun-goli-broj` |

Tvrdnja koja nosi najviše: **avans pool po kooperantu** ima i tvrdnju da
vozilo postoji (kooperant sa avansom i ≥2 bloka, i da se dve politike stvarno
razlikuju u zbiru) — bez nje bi dedup tvrdnja bila zelena i nad fixture-om
koji razliku ne može da pokaže.

Tvrdnje „storniran nije u listi" i „plaćen nije u listi" **nemaju svoju
sabotažu**: čuva ih nasleđeni sloj (`ExcludeStornirano` / `GetOpenOtkupi` u
`modNovac`), koji nije diran u ovom PR-u — sabotaža nad njim bi obarala tuđe,
već pokrivene testove.

Diff u `modOtkupUI` je **jedna linija — pečat** (`OTKUI_BUILD` →
`v6-ui-185`), isti razlog kao §8.10/R3.

**Pušteno i zeleno:** `RunAllTests` **131 / 0** (prvo puštanje novih testova),
`RunBankaImportTestSuite` **196 / 0**, `vba_check` + `--self-test`,
`sabotaza --self-test`, `who_writes --check`. Dvosmerni dokaz svih dvanaest
sabotaža: v. ispod.

**Nalaz iz sesije — dokaz i uređivanje izvora se ne mešaju.** Prvi prolaz
`dokaz.py banka-nalozi` oboren je mojom greškom u redosledu, ne kodom: usred
prvog ciklusa podignut je pečat u `modOtkupUI`, a dokaz posle svakog vraćanja
poredi potpis **celog** `src-vba` — pa je uredno stao uz `REVERT-FAIL`
(tačno ono za šta ta provera postoji). Izvor je vraćen, pečat podignut PRE
drugog prolaza, i pravilo zapisano: dok dokaz radi, u `src-vba` se ne dira
ništa.

**Ručna kapija operatera (traži se izričito):** `Alt+F11 → Debug → Compile
VBAProject`, pa smoke nad pravim podacima — v. checklistu u PR-u (izgled
zone, traka korpe, potvrda i ishod CSV-a, PDF specifikacija, avans).

### 22.9 Prvi smoke: dva nalaza (obe ispravke u istom PR-u)

Compile je prošao, ekran radi, CSV je nastao — i doneo dva nalaza koje suite
nije mogla da vidi.

#### N1 — goli 18-cifreni račun se u Excelu raspada

U generisanom fajlu su dva od tri računa primaoca stajala kao `3,25934E+17`
i `2,059E+17`: računi uneti u matične podatke kao **golih 18 cifara**. CSV na
disku je bio tačan — ali Excel pri otvaranju niz duži od 15 cifara čita kao
BROJ, prikaže ga u naučnoj notaciji i **drži samo 15 značajnih cifara** —
pa bi snimanje iz Excela (banner „possible data loss" je tačno to) račun
primaoca **uništilo** pre uvoza u e-banking. Treći račun, unet sa crticama,
ostao je tekst i ceo.

Kolona je, dakle, i do sada nosila mešavinu dva oblika (račun ide u fajl
onako kako je unet). Ispravka: `FormatRacunZaNalog` — **tačno 18 golih
cifara** se kanonizuje u NBS oblik `3-13-2`; sve ostalo prolazi netaknuto
(format koji domen nema se ne izmišlja). Primenjeno na obe kolone računa u
`BuildNalogCsvPayload`, pa ispravku dobija i legacy putanja (isti writer).

Redosled po `CLAUDE.md` §2: test **T22** je prvo pisan nad no-op verzijom
funkcije (bit-identično današnjem ponašanju) i pušten — **4 tvrdnje crvene**
(`dobijeno=205000000012345678, ocekivano=205-0000000123456-78`) — pa je
kanonizacija ušla i suite je zelena. Kvar nije ovog ekrana: isti fajl pravi
i legacy forma od AUD-026.

#### N2 — „nekad se iznos ne isplaćuje u potpunosti"

Operatersko pitanje „gde se definiše iznos po bloku" je podatak koji obara
odluku (B): delimična isplata je **stvaran tok**, a ekran ju je slao u
legacy formu. Ušla je kao radnja **„Iznos…"** (v. §22.2, oboreni deo):

- unos kroz `InputBox` radnje (presedan: SEF komentar na Fakturisanju) — ne
  kroz polje zone, pa nema stanja vezanog za „izabran red" koje sort/strana/
  filter zastarevaju;
- pravila unosa su legacy `txtIsplatiti_Exit` pravila, u čistoj funkciji
  (`BnPostaviIznos`): cent-domen pre svake provere, `> 0`, nikad preko
  otvorenog, jednak otvorenom **briše** zadato;
- zadate iznose pri svakom čitanju liste usklađuje **isti račun** kao legacy
  override: `ClampOverridesToOpenDict`, jezgro izdvojeno iz
  `ClampOverridesToOpen` (wrapper za legacy formu nepromenjen, T12 zelen) —
  nestao/zatvoren blok gubi zadato, veće se spušta, manje ostaje, uz poruku;
- nova vidljiva kolona **ISPLATITI** uz OTVORENO (podrazumevano isti broj;
  razlikuju se tačno gde je operater zadao manje); `OdaberiBlokoveZaNaloge`
  dobija opcioni rečnik zadatih iznosa — bez njega ponašanje nepromenjeno;
- namerno BEZ nove kapije u izvozu (§21 lekcija): klamp drži zadato ≤
  otvoreno pri čitanju, a svežu preplatu između čitanja i klika i dalje
  preseca `ValidateNalogSaldo` za ceo fajl — isti slojevi kao legacy.

**Ekransko vezivanje `BlokoviZaIzvoz` → `Iznosi()` je jedan red, provereno
čitanjem** (domenska polovina je pod testom i sabotažom) — isti oblik
beleške kao §9 „vezivanje kapije u dve rute".

#### N3 — pretraga „ne radi": kvake u podacima, DE tastatura kod operatera

Drugi krug smoke-a. Mehanizam pretrage je ljuskin i isti za sve ekrane
(`txtSearch` → `Scr_Rows(filter, q)` — provereno čitanjem), a upit uredno
stiže ekranu. Ali **prava imena nose dijakritiku** (Petrović, Đerić,
Dželebdžić), a operater — na DE tastaturi, koja te znakove nema — kuca
„petrovic": `InStr` nad sirovim haystack-om ne nalazi ništa, i pretraga
izgleda mrtva. Fixture to nije mogao da vidi: sva njegova imena su ASCII.

Popravka: `modUiData.TekstZaPretragu` — obe strane (haystack i upit) se
svode na ASCII istom transliteracijom koju repo već ima
(`SanitizeFileNamePart`: š→s, đ→dj…). Živi u ljuskinom data sloju da je
ekrani dele; fixture dobija kooperanta sa kvakama (`KOOP-NAL-DJ`, „Đorđe
Šarčević" + blok `OTK-NAL-DJ`) — bez njega bi tvrdnja merila prazan skup.
**Zatečeni ekrani (Uvoz izvoda, Fakturisanje) imaju istu rupu u svojim
haystack-ovima** — dopisivanje `TekstZaPretragu` tamo je zaseban, mehanički
posao i ovde se samo beleži.

#### N4 — isti Excel-broj kvar i u PDF specifikaciji

Kolona „Tekući račun" u specifikaciji je pokazivala `3,4E+17`: šablon
(`FillIsplataSpecSablon`) je Excel sheet, pa golih 18 cifara u ćeliji prolazi
kroz istu konverziju kao CSV otvoren u Excelu. Procena iz N1 da „PDF nema
taj problem" bila je pogrešna. Popravka: `spec(i,5)` ide kroz **isti**
`FormatRacunZaNalog` (pravilo pod T22; vezivanje jedan red, čitanjem).

#### N5 — red trake i zbir ispod njega su govorili dva različita broja

Stavka sa zadatih 10.000 (otvoreno 21.798) je u traci „U NALOZIMA" stajala
kao `broj · 21.798`, a zbir ispod kao `10.000 RSD`. `KorpaRedPrikaz` sada
ide kroz `BnIznosZa` — isti račun kao zbir, pa se ne mogu razići. Tvrdnja u
testu 132, sabotaža `banka-nalozi-traka-nosi-otvoreno`.

#### §22.10 Recenzija PR-a: prazan izbor je bio merge blocker

Recenzija posle drugog kruga (#244) — četiri prihvaćene tačke, sve u istom
PR-u:

**R1 (blocker) — „prazna korpa = svi" + „clear sprečava dupli export" je
bila NETAČNA tvrdnja.** CSV ne knjiži isplatu: blokovi ostaju otvoreni i
posle fajla, a izbor se posle izvoza prazni — pa bi drugi klik tiho izvezao
naloge za **sve** otvorene, uključujući **pun** iznos bloka čiji je zadati
deo upravo izvezen. Ugovor promenjen: **prazan izbor ne izvozi ništa** (gate
u `BlokoviZaIzvoz`, pre ijednog čitanja tabela; poruka razlikuje „izbor
prazan" od „izabrani ne mogu u nalog"), a „svi" je izričita radnja
**„+ Svi sa računom"** (`bnsve`, `trebaRed=0`) koja korpu puni kroz istu
domensku Nothing-granu `OdaberiBlokoveZaNaloge` — grana ostaje i testirana
i korišćena, samo više nikad implicitno.

**R2 — osnova i verzija.** Grana je bila iza `main`-a (#243 je u
međuvremenu ušao i zauzeo `v2.87.0`): `origin/main` je unesen merge-om
(konflikt samo u release notes — #243-ov `v2.87.0` ostaje, ovaj unos
postaje **`v2.88.0`**; `sabotaza.py` se spojio čisto, #243 korekcije
`cilj-bez-istorijske-kapije` / `zbirna-ispravka-cilj-bez-kapije` očuvane
uz nove Banka unose). #243 nije dirao numeraciju testova (staje na 126) ni
pečat (184), pa 127–132 i `v6-ui-185` ostaju važeći.

**R3 — ekransko vezivanje iznosa nije bilo dokazano.** Domenska polovina
(`OdaberiBlokoveZaNaloge` + sabotaža) jeste, ali bi uklonjen argument
`, Iznosi()` u `BlokoviZaIzvoz` ostavio sve zeleno — a UI bi pokazivao 250
dok fajl nosi 600. Za „koliko novca ide u nalog" beleška „provereno
čitanjem" nije dovoljna: seam **`Scr_BnBlokoviZaIzvozTest`** meri istu
putanju koju zovu CSV i PDF (uključujući i gate praznog izbora), sabotaže
`banka-nalozi-ekran-ne-salje-iznose` i `banka-nalozi-prazan-izbor-izvozi-sve`.

**R4 — blok koji izgubi račun ostajao je u izboru.** Nije finansijski kvar
(izvoz ga preskače), ali traka/zbir/potvrda ne smeju da ga pokazuju kao
spreman: „živa" mapa za usklađivanje sada sadrži samo otvorene **sa
računom**, pa takav blok izlazi uz istu poruku (i njegov zadati iznos s
njim). Sabotaža `banka-nalozi-korpa-drzi-bez-racuna`.

Uz to je ublažen komentar zaglavlja modula: `BnPostaviIznos` **jeste**
pravilo unosa (preslikana kopija legacy `txtIsplatiti_Exit`, obrazac §5/
Faza B — dve kopije žive namerno); „ovde nema pravila" više nije doslovno
tačno i sada je zapisano šta jeste a šta nije ekranovo.

#### N7 — „filter ne radi": model je bio tačan, cena po otkucaju nije

Treći smoke. Popravka kvaka (N3) NIJE bila ceo uzrok — na pravoj svesci su
imena već ASCII (`Zecevic Gvozden`), pa je transliteracija tamo no-op.
Dijagnoza je bila nepotpuna, pa je merena, ne doterivana (§2): prošireni
`Diag_BnRedovi` pamti poslednji `(filter, q, n)` koji je `Scr_Rows` primio.

Merenje na pravoj svesci (1.595 otvorenih blokova):

```
POSLEDNJI POZIV: filter=[sve] q=[gvozden] vraceno redova=38
MREZA red 1..3 = Zecevic Gvozden (tacno filtrirani redovi)
```

Upit **stiže**, ekran **vraća 38**, mreža ih **drži** — ceo lanac je
ispravan. Ali svaki otkucaj plaća **pun `BuildBlokIsplataList`** (ceo
`tblOtkup` + isplate + računi + avansi + vlasnici), po nekoliko sekundi po
slovu: operater kuca, ne vidi ništa, i „posle ~10 s se pojavi celo ime" —
doživljaj mrtvog filtera uz savršeno tačan model.

Legacy formu je isto ovo već naučilo: `LoadBlokovi` (skupo) ide samo kad se
izvor menja, a kooperant-filter je „LAGANI re-filter nad već učitanom
`m_FullBlokovi`, bez čitanja tabela". Taj obrazac je sada prenet: **snimak
liste se kešira u ekranu** (`Snimak()`), pretraga i čipovi filtriraju nad
njim trenutno, a invalidira ga `Scr_ResetCache` — koji ljuska ionako zove
posle svakog upisa. *(Dopuna `v6-ui-186`: `Scr_ResetCache` stiže samo
AKTIVNOM ekranu, pa su snimak i KPI preživljavali upis sa drugog ekrana —
od recenzije PR #245 oba proveravaju i generaciju podataka,
`modUiData.DataGeneracija` — v. §23.10/R1.)* **Izvoz keš ne koristi**: `BlokoviZaIzvoz` i finalna
kapija čitaju svež saldo, kao i do sada. Merljivo brojačem stvarnih čitanja
(`mSnimakPunjenja`, obrazac `mCiljPunjenja` iz §9): tri uzastopne pretrage
i promena čipa = **jedno** čitanje tabela; posle `Scr_ResetCache` sledeće
čitanje ide u tabele. Sabotaža `banka-nalozi-pretraga-puni-iznova`.

#### N6 — otvoren nalaz: ćelija ISPLATITI bez vrednosti na jednom redu

Na screenshotu prve strane red `2/020726-4` (kooperant bez računa) deluje
kao da ima **praznu ćeliju ISPLATITI**, dok model tu uvek predaje broj
(`BnIznosZa` vraća `Double` bez izuzetka). Čitanjem koda uzrok nije nađen, a
fixture ga ne reprodukuje (test tvrdi `IsNumeric` nad svakim redom i zelen
je) — pa se **ne krpi pretpostavkom**. To je tačno klasa „ćelija prazna ili
tuđa = preskočen upis pod `On Error Resume Next`" iz §9.10, za koju postoji
presedan-alat: ekran dobija **`Diag_BnRedovi`** (Alt+F8, ispisuje šta ekran
predaje i šta mreža drži). Ako se na pravoj svesci potvrdi — merenje kaže
gde.

---

## 23. Izveštaji — šta je preneto (`v6-ui-186`)

Peti ekran **Faze E**, stavka 19. Red u registru (`modUiScreens.ScrRows`) je
postojao od `S3a` — stavka menija se do sada crtala prigušena jer modula nije
bilo. Ovim se piše modul koji taj red već očekuje; **registar se ne dira**.

Ovo je poslovno najvidljiviji ekran do sada: izveštaji se pokazuju
knjigovođi, banci i vlasniku, pa je merilo prelaska **SLAGANJE, ne izgled** —
svaka lista ima test koji njen sadržaj veže za nezavisan read-model (v.
§23.6). Kao i kod Banka naloga, najskuplji deo migracije je već bio odrađen:
`modIzvestaj` od RF-06/RF-07 nosi čist API `Report*(entitet, opseg) → 2D niz`
sa dokumentovanim kolonama, matricu dostupnosti i deljene seam-ove — ekran je
zato **prikaz nad postojećim računima**; nijedan `Report*`, matrica ni
štampa se ne menjaju.

### 23.1 Gde je šta završilo

| Legacy (`frmIzvestaj`) | Novo mesto |
|---|---|
| 9 statičkih tabova `mpReports` + 2 runtime taba | **10 lista** deljene mreže (v. odluku A u §23.2) |
| `tglOM/tglKupci/tglVozaci/tglKooperanti` | 4 seg prekidača zone (`NewSegBtn`, vrsta `"seg"` — §7.7) |
| `tglPojedinacni/tglZbirni` | 2 seg prekidača režima |
| `cmbEntitet` + `LoadEntiteti` (po tipu) | polje zone `scrIzEnt` (obrazac `PuniPartnerCombo` iz §9: 2 kolone, čist ID u drugoj, prikaz sa ID-jem; kooperanti samo aktivni, kao legacy) |
| `txtDatumOd/Do` (default 1.1. — danas) | polja zone `scrIzOd`/`scrIzDo`; parsiranje = `DatGranica` pravilo iz `modScrDokumenta` (nepotpun unos = „još nema granice") |
| `AutoRefresh` + lazy `m_genTabs` po tabu | keš snimka po ključu konteksta (§23.5) + `RefreshFromData` koju ekran sam zove (§8.2 — sve liste zavise od polja zone) |
| matrica `UpdateReportMode` → `IzvestajTabDostupan` | `IzListaDostupna` — pita **istu matricu**; runtime tabovi preslikan legacy uslov (§23.3) |
| `btnStampaj` (tabelarni PDF aktivnog taba) | dugme zone „Štampaj izveštaj" → `PrintIzvestaj` (nepromenjen) |
| `btnStampajKarticu` (tab-aware) | dugme zone „Štampaj karticu (PDF)", vidljivo samo na listama kartica → `PrintKarticaPDF` / `PrintKarticaAmbalazePDF` (nepromenjeni) |
| `m_btnStampajOtk` / `m_btnStampajOtkRoba` / `m_btnStampajAmb` | **jedna radnja nad redom** `izprint` („Štampaj dokument") na 4 liste, ruta po listi i tipu dokumenta |
| `StampajReversAmbDok` + `AmbRedStorniran` (račun u formi!) | **izdvojeno** u `modIzvestaj.StampajReversAmbalaze` (+ `IzvAmbRedStorniran`); forma zadržava svoju kopiju i ne menja se (§5/Faza B) |
| headers po tabu u `btnStampaj_Click` | `IzHeaderiZaListu` (izvedeno iz opisa kolona ekrana) |
| `lblStatus` (4 stanja) | hint zone: „izveštaj ne postoji za kombinaciju" ≠ „izaberi entitet" ≠ opis prikazanog konteksta |
| specijalni redovi lista (v. §23.4) | **brojke zone** (OM avans / agro nerasp.; primljeno / kod otkupca) |
| detail panel „Detalji otkupa" (`modKarticaDetalji`) | **NIJE preneto** — v. §23.7 |

### 23.2 Odluke

**(A) 10 lista, ne 9+2 ni 11.** Legacy tabovi SALDO OM (0) i SALDO KUPCI (1)
postaju **jedna lista SALDO** sa dispatch-om po tipu — tačno ono što legacy
`GenerateSaldoReport` već radi, a matrica garantuje da operater **nikad ne
vidi oba** (tab 0 samo OM, tab 1 samo Kupac). Kolone-po-kontekstu su
presedan iz samog legacy-ja (tab ROBA menja kolone po tipu), a mreža
geometriju prati iz opisa kolona pri svakom crtanju (§9.2/3). Dva runtime
taba („Otkupni listovi", „Pregled ambalaže") su pune liste: lista BLOKOVI na
Dokumentima **nije** isto (blokovi jedne otpremnice, ne cele stanice u
periodu — §8.4 prolazi u oba smera). Bazen `MAX_SEG` (11) zadržava jedan
slobodan slot; test to tvrdi.

**(B) Matrica → prazna lista sa objašnjenjem** (opcija i). Čipovi ne mogu da
nose entitet-tip (tip je polje zone deljeno svim listama; čip pada na „Sve"
pri promeni liste, a „Sve" za tip ne znači ništa), a dinamičko sklanjanje
segmenata je tiho nestajanje liste. Svih 10 segmenata postoji uvek;
`IzListaDostupna` pita `IzvestajTabDostupan` (matrica se ne širi i ne
„popravlja" — prazan tab kooperanta u zbirnom je nepostojeći izveštaj);
nedostupna kombinacija = 0 redova + hint koji imenuje razlog, **različit** od
„izaberi entitet" i od „nema podataka" (FM-0029 merilo: ni pun naslov nad
trajno praznom listom, ni tiho nestajanje).

**(C) Štampe: dva dugmeta zone + jedna radnja nad redom.** „Štampaj
izveštaj" štampa tačno ono što operater vidi (čip + pretraga — `PrintSpecDat`
presedan), sa naslovom iz **konteksta snimka** (AUD-024: naslov opisuje
prikazano, ne trenutno stanje polja) i vidljivom napomenom kad je filter
aktivan — papir bez nje izgleda kao ceo izveštaj. „Štampaj karticu (PDF)" je
tab-aware kao legacy `btnStampajKarticu`. Radnja `izprint` živi samo na 4
liste čiji red ima dokument (OTK_LISTE, ROBA/OM, AMBALAZA, KARTICA); ruta po
tipu dokumenta je legacy `m_btnStampajAmb_Click` pravilo, revers kroz
izdvojeni račun. Agregatne liste nemaju nijednu radnju (ljuska krije dugmad
— IZVODI presedan, §9.2).

**(D) Dvoklik namerno ne radi ništa; detail panel se odlaže.** Jedina radnja
je štampa — promašen dvoklik koji pokrene PDF je gori od nikakvog (§9.5
princip). Pregled stavki bloka na klik (detail panel) čeka padajuće redove
(§5/Faza C — poznat odložen posao); bitno iz legacy-ja (štampa dokumenta iz
reda) jeste preneto.

**(E) Keš od prvog dana** — v. §23.5.

### 23.3 Matrica dostupnosti — ekran je ne prepisuje

`IzListaTab(lista, tip)` mapira listu na legacy `IZV_TAB_*` indeks (SALDO
bira tab 0/1 po tipu); `IzListaDostupna` za 8 statičkih lista **zove**
`IzvestajTabDostupan`, za 2 runtime liste preslikava uslov iz
`UpdateReportMode` (Otkupni listovi samo OM-pojedinačno, Pregled ambalaže
samo Kooperant-pojedinačno). Test tvrdi poslovne činjenice matrice
(kooperant+zbirno nema ništa; vozač pojedinačno tačno AMBALAZA i MANJAK;
ISPLATA samo OM; ROBA za vozača **ne postoji kao tab** iako
`ReportOtkupRobaVozac` postoji — matrica prati dispatch, ekran matricu), pa
refactor koji bi je zaobišao pada po imenu.

### 23.4 Identitet po listi — i šta je izdvojeno u zonu

Identitet u poslednjoj koloni **prioriteta 4** (mreža crta do 3), čitan kroz
`GridCell`; mapa „prikaz → ID" ne postoji. Red bez identiteta nosi prazno i
radnja **odbija** porukom.

| Lista | Identitet reda |
|---|---|
| OTK_LISTE | `OTK\|<OtkupID>` (štampa celog bloka, obe klase) |
| ROBA (OM oblik) | `OTP\|<OtpremnicaID>`; Kupac/Vozac oblik: **prazno** (agregat po vrsti) |
| AMBALAZA | `DokumentTip` + `DokumentID` u dve prenosne kolone (ključ reversa traži oba; tip ambalaže je vidljiva kolona istog reda) |
| KARTICA | legacy ref-ključ: `OTK\|<id>` ima dokument; `NOV`/`MAG`/`AMB` nemaju → radnja odbija (legacy `Case Else` poruka) |
| SALDO / ISPLATA / ZBIRNI / CENA / MANJAK / AMB_KARTICA | **prazno** — agregati; `Report*` povratci ne nose ID (i legacy nije imao radnju nad tim redovima) |

**Dva pravila prikaza istine** (odstupanja od legacy-ja, namerna):

1. **Specijalni redovi ne idu u mrežu nego u brojke zone**: „OM AVANS
   (nerasporedjen)" i „AGROHEMIJA (nerasporedjena, van UKUPNO)" iz
   `ReportSaldoOM`, tri kontrolna reda iz `ReportIsplata`. U tipiziranim
   kolonama mreže njihove prazne ćelije bi postale „0,00" — FM-0028 #5 klasa
   laži. Izdvajanje i prikaz dele isto mesto (`VrstaReda`), a testovi tvrde i
   da brojka u zoni nosi vrednost i da reda u mreži **nema**.
2. **UKUPNO red nikad ne ide u mrežu, POČETNO STANJE živi samo u
   nefiltriranom prikazu**: mreža sortira po koloni pa UKUPNO pluta (prvi
   smoke — v. §23.9/S4), a pod filterom bi tvrdio zbir koji ne odgovara
   vidljivim redovima. Zbir prikazanih uvek daje podnožje mreže (računat pod
   istim filterima kao redovi — §13); tabelarna štampa dodaje svoj izračunat
   UKUPNO nad tačno štampanim redovima.

Uz to: **Manjak kg i % su razdvojene kolone** (legacy ih je spajao zbog
ListBox limita 10; `MAX_COLS` je 14); kolone kod kojih je „prazno poruka"
(prijemnica/manjak bez prijema) idu kao tekst koji formatira ekran — nikad
„0,00" umesto oznake; ćelije gajbi prikazuju prazno umesto nule (legacy
pregled). **Running saldo kartice je semantika reda, ne prikaza**: pretraga
seče redove, ali saldo kolona ostaje „saldo posle tog dokumenta u punoj
kartici" — zato KARTICA **nema čipove** (filter po vrsti reda bi trajno
pokazivao isečen saldo uz kumulativnu kolonu). Čipovi postoje samo gde je
filter prirodan podatak reda: MANJAK (`sve · bez prijema`) i AMBALAZA
(`sve · ulaz · izlaz`); prvi je uvek najširi.

### 23.5 Keš snimka (N7 pravilo od prvog dana)

Sirov `Report*` povratak se kešira po **ključu konteksta**
(`lista|tip|režim|entitet|od|do`). Pretraga i čipovi su re-filter nad
snimkom — **nula čitanja tabela po otkucaju** (§22.9/N7: pun prolaz po
otkucaju je plaćen kvar); promena entiteta/opsega/liste/režima legitimno
čita ponovo — to su dva svesno razdvojena slučaja. `Scr_ResetCache`
(ljuska ga zove posle svakog upisa) snimak proglašava zastarelim. Broj
stvarnih čitanja meri `mSnimakPunjenja` (test: tri pretrage + čip = jedno
punjenje; posle reset-a novo), a `Diag_IzRedovi` (Alt+F8) od prvog dana
pamti poslednji `(filter, q, n)` + ključ snimka — obrazac `Diag_BnRedovi`.

Datumska polja koriste **postojeće pravilo** `DatGranica` (nepotpun unos
„21." = još nema granice — ne prazni listu i ne čita tabele; potpuno prazno
polje = pun opseg, jer `Report*` primaju `Date`). Ljuskina `specOd/specDo`
polja **nisu** upotrebljena: `SpecDatLista` ih tvrdo veže za
DOKUMENTI/OTPREMNICE, a diff ljuske je pečat i ništa više.

### 23.6 Slaganja — srce zadatka

Sve tvrdnje su **relacije** (golden brojke zabranjene), nad nezavisnim
read-modelima; testovi **133–141** u `modTest` (izvršavaju se pre
mutirajućih 124–126):

| Lista | Tvrdnja | Nezavisan read-model |
|---|---|---|
| ZBIRNI (OM) | kg/vrednost po (stanica, vrsta) | `ReportProsecnaCena("OM", S)` po svakoj stanici — dva različita čitača |
| ZBIRNI (Kupac) | UKUPNO = Σ pojedinačnih po **svakom** kupcu | `ReportOtkupRoba("Kupac", K)` + ručni prolaz kroz `tblPrijemnica` |
| ZBIRNI (Vozac) | amb izlaz po vozaču | ručni zbir `tblZbirna` po vozaču |
| CENA | cena × kg = vrednost po redu; sume ↔ `ReportSaldoKupci` | **dva određenja vrste** (kolona prijemnice vs. vrsta zbirne) |
| SALDO (OM) | po redu saldo = vrednost − isplaćeno − agro; isplaćeno ↔ ručni zbir sa `NovacRedPripadaStanici`; ambalaža ↔ `GetAmbalazeStanje` | ručni prolazi + kanonski saldo |
| SALDO (Kupac) | novac ↔ ručni zbir uplata; saldo = vrednost − novac | ručni prolaz |
| ISPLATA | po redu ukupno = keš + virman firma + virman avans (sva tri kanala pod naponom); UKUPNO ↔ ručni zbir tri tipa; zona primljeno/kod ↔ ručni zbir Firma→Otkupac | ručni prolazi kroz `tblNovac` |
| ROBA (OM) | UKUPNO otpremljeno ↔ `tblOtpremnica`; blokovi ↔ `tblOtkup` vezan za te otpremnice | ručni prolazi |
| MANJAK | red ↔ `ManjakStavka` aritmetika; UKUPNO samo nad redovima **sa prijemom**; red bez prijema nosi oznaku, ne nulu | čist seam + vlasnike čuvaju postojeći E2E u `modIzvestajTests` |
| AMBALAZA | **zbirni režim = Σ pojedinačnih po tipu** (dva agregatna puta iste funkcije); UKUPNO ↔ ručni prolaz kroz ledger; ≥2 tipa ambalaže kao vozilo | ručni prolaz |
| KARTICA | running red-po-red; UKUPNO = početno + promet; Σ zaduženja ↔ ručni kg×cena; rekap ↔ ručni Σ kg; amb saldo ↔ `GetAmbalazeStanje`; **komplementarne granice**: završni saldo (1.1–31.3) = početno stanje (od 1.4), za novac i ambalažu | ručni prolazi + kanonski saldo + sama kartica preko granica |
| AMB_KARTICA | running; završni saldo ↔ `GetAmbalazeStanje` | kanonski saldo |
| OTK_LISTE | broj redova + Σ kg/vrednost ↔ ručni prolaz kroz `tblOtkup` | ručni prolaz |

**Tri nalaza iz crvenih krugova** (testovi i dokaz su radili pre nego što su
postali zeleni):

1. **Fixture je bio domenski nekonzistentan**: otkupi sa `KolAmbalaze` bez
   uz-otkup ledger parova koje `SaveOtkup` piše — kartica (kolone
   `tblOtkup` + samostalna kretanja) i kanonski saldo (`GetAmbalazeStanje`)
   su se legitimno razilazili (−47 vs 25). Parovi su dodati u
   `make_fixture.py` za sve nestornirane otkupe KOOP-TEST-1; tri puta
   (finansijska kartica, amb kartica, kanonski saldo) sada mere isto.
2. **Tvrdnja o početnom stanju mora biti vremenski robusna**: raniji test u
   nizu (`T_WriterGuard_AvansSaldoOM`) legitimno upisuje virman sa
   *današnjim* datumom, pa „početno stanje od 1.4 = završni saldo punog
   opsega" ne stoji. Prepravljena na komplementarne granice (gore) — što je
   ujedno i jača formulacija FM-0028 #1.
3. **Prvi dokaz je vratio `NE OBARA NISTA` za `izvestaji-ukupno-prezivi-
   filter`** — i bio je u pravu: pod *pretragom* je UKUPNO red dvostruko
   zaštićen (haystack „UKUPNO" ionako ne sadrži upit), pa uklanjanje
   vrsta-filtera ništa ne menja; pod **čipom** je taj filter jedina brana
   (čip „ulaz" bi UKUPNO red ambalaže propustio — zbirni ulaz > 0), a baš
   taj slučaj tvrdnja nije merila. Test je dobio čip-granu, sabotaža sada
   obara nju. Placebo tvrdnju je otkrio dokaz, ne čitanje.

Tvrdnje slaganja nad samim `Report*` funkcijama **nemaju zasebnu sabotažu**:
mutacija `modIzvestaj` bi obarala i `RunIzvestajTests` (tuđu, postojeću
suite) — isto pravilo kao „storniran nije u listi" u §22.8. Sabotaže (16,
prefiks `izvestaji-`) gađaju ekransku polovinu: izdvajanje u zonu, keš,
matricu, identitet, prikaz istine, zonu.

### 23.7 Šta NIJE preneto, i zašto

- **`frmIzvestaj` se ne gasi i ne menja** — dve kopije žive namerno (§5,
  Faza B). `Report*`, `PrintIzvestaj`, `PrintKartica*PDF`,
  `IzvestajTabDostupan` — nijedno pravilo nije dirano.
- **Detail panel „Detalji otkupa"** — prenet u krugu 4 kao detalj traka u
  zoni (§23.11/S7), ne kao padajući redovi; padajući redovi u mreži ostaju
  Faza C. Štampa dokumenta iz reda jeste preneta.
- **Rang kooperanata** deli jedan račun sa Dokumentima
  (`KoopRangRows`); od kruga 8 postoji i ovde kao tab „Rang“ uz period
  zone — račun nije prepisan (Optional granice).
- **Zbirni oblik ambalažnog pregleda kroz ekran**: matrica tab 3 nudi samo
  pojedinačno (i u legacy-ju), pa lista AMBALAZA uvek koristi pojedinačni
  oblik; zbirna grana `ReportAmbalazeZbirni` ostaje pokrivena testom
  slaganja kao API.
- **Agro (nerasporedjena) brojka** nema fixture vozilo (nema magacin izlaza
  bez kooperanta) — izdvajanje postoji i testira se kroz OM avans polovinu
  istog mehanizma; tvrdnja nad agro brojkom bi merila prazan skup.
- **Marža i Sledljivost** su zasebni ekrani (stavke §5).

### 23.8 Poznati nalazi van ovog PR-a

- **`PrintIzvestaj`/`OutputToSheet` upisuju sirove stringove u ćelije**
  print sheeta — string koji Excel ume da protumači (npr. tip ambalaže
  `12/1` kao datum) menja oblik u štampanom PDF-u. Za **novi UI** zatvoreno
  u krugu 4: `PrintIzvestajHouse` piše sve data ćelije sa
  `NumberFormat="@"` (§23.11/S8). **Legacy putanja** (`frmIzvestaj` →
  `PrintIzvestaj`/`OutputToSheet`) i dalje nosi kvar — pripada zajedničkom
  prolazu, ne ovom PR-u. Golih nizova >15 cifara u izveštajima nema
  (provereno po listama — nijedna ne nosi račun), pa N1 klasa ne nastaje.
- **Kursor preko placeholder-a pretrage** — poznat estetski backlog svih
  ekrana, ne dira se (§22.9).

### 23.9 Prvi smoke: tri prijave, četiri nalaza (ispravke u istom PR-u)

Compile je prošao, ekran radi na pravoj svesci — i doneo nalaze koje suite
nije mogla da vidi:

**S1 — „dropdown prikazuje samo prvu stavku."** Ljuskin panel izbora
filtrira stavke po **tekućem tekstu comba** (`PopIndex`: sužavanje po
podnizu — to je i smisao kucanja). Ekran je, kao legacy, auto-birao prvu
stavku (`ListIndex = 0`) → combo od prvog trenutka drži pun tekst → panel
zauvek nudi samo tu stavku. Banka uvoz nema auto-izbor, pa se tamo nije
videlo. Ispravka je **ekranska** (ljuska netaknuta): podrazumevani entitet
živi u stanju ekrana (`mDefaultId` = prvi entitet tipa; `IzabraniEntitet`
ga vraća dok izbora nema), combo ostaje prazan sa placeholder-om, a hint
ispod polja kaže koji je entitet **stvarno** prikazan. Izbor operatera i
dalje preživljava refill. Legacy „odmah vidiš podatke" ponašanje je
zadržano — bez trovanja panela.

**S2 — „sve je sporo."** Dva pojačivača u ekranu: (1) `chg:` stiže i tokom
**programskog** punjenja comba, a handler nije imao `mFill` guard — refill
usred `Scr_Rows` je okidao ugnežđen `RefreshFromData` i dupli `Report*`
prolaz; (2) keš je držao **jedan** snimak, pa je svaki klik na drugu listu
plaćao pun prolaz ispočetka — a šetnja po 10 lista je osnovni tok ekrana.
Sada: guard u oba `chg:` handlera + **mapa snimaka po ključu konteksta**
(kapa 16; `Scr_ResetCache` prazni sve) — povratak na viđenu listu je
trenutan, upis i dalje invalidira sve.

**S3 — „na kartici nema gde da se izabere kooperant."** Posledica S1
(combo je delovao mrtav) + hint je govorio samo zašto, ne i kuda. Hint na
listama kartica sada kaže: „Kartica postoji samo za kooperante — klikni
'Kooperanti' pa izaberi kooperanta."

**S4 — UKUPNO red je plutao po mreži** (vidljivo na screenshotu Isplate:
UKUPNO usred liste). Mreža sortira po koloni, a legacy UKUPNO je „poslednji
red" samo u ListBox-u bez sortiranja. UKUPNO red zato **nikad ne ide u
mrežu**: zbir prikazanih daje podnožje (računat pod istim filterima), a
tabelarna štampa dodaje **svoj izračunat UKUPNO red** — nad tačno onim
redovima koji su na papiru, po tipu kolone. POČETNO STANJE ostaje red
podataka (živi u punom prikazu, nestaje pod filterom). Time je pravilo
„UKUPNO samo u nefiltriranom prikazu" iz §23.4 postalo strože: „UKUPNO
nikad u mreži" — testovi i sabotaža su prepravljeni na taj oblik.

### 23.10 Recenzija PR-a #245: tri nalaza (krug 3)

Recenzija posle drugog kruga — dva blocker-a i jedan manji, sva tri
prihvaćena i zatvorena u istom PR-u:

**R1 (blocker) — izvedeni keš je preživljavao upis sa DRUGOG ekrana.**
`RefreshFromData` resetuje keš samo **aktivnog** ekrana, a `ActivateScreen`
pri povratku ne resetuje ništa — pa je sekvenca „Izveštaji → drugi ekran →
upis → nazad" pogađala stari snimak pod istim ključem i pokazivala **stare
brojke**. Isti kvar je od `v6-ui-185` nosio i snimak liste (i značka) na
Platnim nalozima. Rešeno **deljenim ugovorom invalidacije**, bez TTL-a i
bez imena ekrana u ljusci: `modUiData.ResetCache` (jedina tačka kroz koju
prolaze svi upisi novog UI-ja) podiže **generaciju podataka**
(`DataGeneracija`), a ekran uz svoj keš pamti generaciju punjenja i pri
čitanju odbacuje stariju — Izveštaji (mapa snimaka), Platni nalozi (snimak
+ KPI značke). Povratak na ekran **bez** upisa i dalje ide iz keša (S2
dobit ostaje). Test 142 upis simulira tačno onim pozivom koji
`RefreshFromData` radi; diff ljuske je `modUiData` (+generacija, pečat) —
`modOtkupUI` i dalje netaknut (do kruga 5 — v. §23.12/S9, tri opšte
linije u `RefreshFromData` za kontekstne tabove).

**R2 (blocker) — štampani UKUPNO je sabirao i nesabirljivo.** Generička
suma „svaka numerička kolona" je sabirala prosečne cene, prosek gajbi i
**running saldo kartice** (zbir međustanja nije poslovna vrednost) — tip
kolone opisuje prikaz, ne aditivnost. Uvedena je **politika sabirljivosti
po listi** (`IzSabirljive`, 1-based indeksi vidljivih kolona, verna legacy
UKUPNO redovima): sabira se promet, nikad prosek i nikad stanje; time se
sabiraju i txt kolone koje **jesu** promet (ulaz/izlaz gajbi — generička
suma ih je preskakala). Podnožje kartice uz to dobija neto promet
prikazanih redova umesto „Vrednost 0,00".

**R3 — „Štampaj dokument" se nudio i gde red nema dokument.** Radnja je
bila po listi; ROBA u kupac/vozač obliku je agregat po vrsti bez
ref-kolone, a nedostupna kombinacija nema ni redove. `Scr_Radnje` sada ide
kroz `IzRadnjeZaKontekst(lista, tip, režim)`: bez radnje kad je kombinacija
nedostupna ili kad oblik nema dokument-grain.

### 23.11 Smoke krug 3: četiri prijave (krug 4 ispravki)

Posle čistog Compile-a operater je u trećem krugu prijavio četiri stvari —
sve četiri su ispravke ponašanja, ne kozmetika:

**S5 — prelaz na Ambalažu „katastrofalno spor".** Isti oblik kvara kao S2,
ali u **jezgru izveštaja**, ne u ekranu: `ReportAmbalazePojedinacni` je za
svaki red ledger-a zvao `ResolveDokBroj`, a ovaj **tri LookupValue-a po
redu** — pri 1.596 redova ledger-a to je O(n·m) prolaza kroz tri
dokument-tabele po svakom otvaranju liste. Sada se **pre** izlazne petlje
jednom grade tri mape (`BuildLookupDict`: otpremnice, prijemnice, otkupi
po broju) i red plaća O(1) (`ResolveDokBrojMape`). Pravilo prevoda je
verno starom (`BuildLookupDict` je „prvi pojav pobeđuje", identično
`LookupValue`-u), pa je ovo čist perf fix — stari `ResolveDokBroj` više
nema pozivalaca i uklonjen je. Ekranski keš (§23.5) ovaj trošak amortizuje
tek od drugog otvaranja; prvo otvaranje mora biti brzo u izvoru.

**S6 — čipovi na nemogućim kombinacijama.** Čip nad listom koja za
kombinaciju tip×režim ne postoji je filter praznog skupa — zbunjuje, kao
i radnja iz R3. `Scr_Cipovi` sada ide kroz `IzCipoviZaKontekst`, koji prvo
pita `IzListaDostupna` (istu matricu §23.3) pa tek onda vraća čipove
liste: nedostupna kombinacija nema ni čipove, hint ostaje jedini sadržaj.

**S7 — drill-down detalj reda.** Legacy `frmIzvestaj` je imao panel
„Detalji otkupa" (dvoklik na red) — to u prenosu nije postojalo (§23.7 ga
je vodio kao „nije preneto"). Sada: klik na red u zoni otvara **detalj
traku** (desno od polja, `izDetCap` + do 6 linija, nestaje na uskom
prozoru umesto da se preklapa). Račun je čist i testabilan bez forme:
`IzDetaljOtkupLista(otkupID)` vraća **sve nestornirane stavke istog
dokumenta** (broj + stanica — isto pravilo po kom reprint štampa ceo
list): „Vrsta Klasa kg × cena = vrednost" po liniji, kooperant na vrhu,
UKUPNO na dnu; `IzDetaljOtpremnice(otpID)` vozača, kg i broj vezanih
blokova. Liste bez dokument-graina ne pune traku; promena liste je čisti.

**S8 — štampe u house obrascu.** Tabelarna štampa je do sada išla kroz
legacy `PrintIzvestaj` (goli grid bez zaglavlja firme). Dokumenti štampani
iz novog UI-ja sada dele isti obrazac kao računi/otpremnice:
`modPrint.PrintIzvestajHouse` — `DocSellerHeader` (podaci firme) +
`DocTitleBlock` (naslov + kontekst-linija „entitet · opseg") + siva
header traka (`DocColHeaderFill`) + UKUPNO red bold. Sve data ćelije se
pišu sa `NumberFormat "@"` **pre** upisa — bez toga Excel „12/1" (tip
gajbe) parsira u datum, ista klasa kvara kao §22 „broj dokumenta kao
datum". Landscape se pali automatski preko 7 kolona; PDF ide u isti
folder i otvara se kao i ostale štampe. Legacy `PrintIzvestaj` ostaje
netaknut za `frmIzvestaj`.

### 23.12 Smoke krug 4: pet prijava (krug 5 ispravki)

Četvrti krug nad pravim podacima — svih pet su doterivanja ponašanja, i
jedno od njih je prva svesna dopuna ljuske za ovaj ekran:

**S9 — tabovi lista su kontekstni.** Kartica i Amb. kartica su stajale u
traci i kad je izabran OM/Kupac/Vozač — mrtva dugmad (isti princip kao S6
čipovi). Sada `Scr_Liste` vraća samo liste koje za izabrani **tip** postoje
u bar jednom režimu (`IzListeZaTip` pita istu matricu §23.3): OM ima 8
tabova (sa Otk. listovima, bez kartica), Kooperant 2 (obe kartice — i prva
dostupna postaje aktivna, pa klik na „Kooperanti" odmah otvara karticu),
Vozač 3, Kupac 6. Lista dostupna samo u drugom režimu istog tipa (Manjak
za OM) ostaje vidljiva — režim je jedan klik, hint vodi. Aktivna lista
kojoj tab nestane prelazi na prvu dostupnu (`PostaviTip`).

*Dopuna ljuske (izuzetak od „diff NULA", prijavljen izričito):*
`RefreshFromData` sada radi i `mGeomStara = True` + `RefreshListSeg` +
`RefreshGridTitle` (tri linije, bez imena ijednog ekrana) — skup tabova sme
da zavisi od konteksta ekrana, pa posle promene konteksta geometrija mora
da se preračuna, a highlight i naslov prate aktivnu listu koju je ekran
mogao da prebaci. Za ekrane sa statičkim listama sve tri linije su
idempotentne. Bez ovoga bi prekrojeni tabovi zadržali boje starog skupa.

**S10 — detalj bez dupliranja.** Princip iz prijave: „dodati samo podatke
ako su novi, nikako duplirati postojeće". Detalj otkupnog lista više ne
ponavlja kooperanta (kolona reda) i UKUPNO dodaje samo kad dokument ima
više linija; detalj otpremnice zadržava vozača (izričito traženo), broj
otkupnih listova, i dobija **zbirnu** i **prijemnice te zbirne** (broj +
kg; vezivanje kao prvi korak `ReportOtkupRobaOM` — broj zbirne pa vozač;
detalj je pregled pa sme fail-open) umesto otpremljenih kg i kg blokova
koji su već kolone reda. Fixture nema otpremnicu čija zbirna nosi
nestorniranu prijemnicu, pa je prijemnica-linija u detalju bez test-vozila
(zapisano, ne tvrdi se) — vozač/blokovi/zbirna jesu pod testom.

**S11 — Roba za kupca = prijemnice.** Agregat po vrsti je već posao taba
Zbirni; operater nad kupcem traži **dokumenta**. Nova
`ReportPrijemniceKupca` u `modIzvestaj` (jedini novi Report ovog kruga)
normalizuje `GetPrijemniceByKupac` — isti read-model kao korpa
fakturisanja — u fiksne kolone sa UKUPNO redom; lista dobija identitet
`PRJ|` (radnja „Štampaj dokument" → `PrintPrijemnica`), detalj reda
(vozač, sorta, ambalaža, status fakturisanja — ništa što kolone već kažu)
i slaganje sa ručnim prolazom kroz `tblPrijemnica`. Stari kupac-agregat
`ReportOtkupRoba("Kupac")` ostaje živ kao API i dalje pod testom
slaganja (T136). Vozačka roba po matrici ne postoji ni u jednom režimu,
pa joj tab (S9) i ne nudi listu.

**S12 — završni saldo kartica u zoni i štampi.** Kolona salda u mreži je
running po redu; „gde smo na kraju" je jedna brojka i sada živi u zoni kao
KPI (novčana kartica: saldo + amb. saldo; amb. kartica: saldo gajbi) —
puni se iz UKUPNO reda Report* kartica (koji u mrežu ionako ne ide), pa se
prikaz i izvor ne mogu razići. Ista brojka ide u kontekst-liniju house
štampe (UKUPNO red štampe sabira samo promet — R2 politika netaknuta).
Test veže zonu za završni running red mreže (novac) i za kanonski
`GetAmbalazeStanje` saldo (ambalaža, pun opseg).

### 23.13 Smoke krug 5: sledljivost u detalju (krug 6)

**S13 — detalj reda dobija karike sledljivosti.** Dve dopune iz petog
smoke-a, obe u istom principu „samo ono što red ne kaže":

- **Prijemnica-linija nosi kupca** („firmu koja je izdala prijemnicu") —
  `DodajPrijemniceZbirne` dodaje `EntitetNaziv("Kupac", …)` u liniju, pa
  je ista dopuna stigla i u detalj otpremnice (Roba/OM) i u detalj
  otkupnog lista.
- **Detalj otkupnog lista** (Otk. listovi i kartica kooperanta) posle
  stavki i UKUPNO dobija kontekst liniju „Vozač … · Zbirna …" (delovi koji
  postoje na listu; ništa se ne izmišlja) i prijemnice te zbirne sa kg i
  kupcem — puna vertikala otkup → zbirna → prijemnica → kupac na jedan
  klik.
- Fixture: dodata **OTP-IZV-Z** (jedina otpremnica čija zbirna,
  ZB-TEST-4, nosi nestornirane prijemnice) — time je zatvorena rupa iz
  §23.12/S10 i linija „prijemnica + kupac" je pod testom. Kupac se u
  tvrdnji meri po ID-u: fixture namerno nema red u `tblKupci` (kupac živi
  samo kao ID na fakturi), pa `EntitetNaziv` pada na ID; na pravoj svesci
  ista linija nosi naziv firme.

**Krug 7 (dopune po odluci posle petog smoke-a):**

- **Jedna kartica, legacy šablon.** Odluka operatera: kartica sa
  rekapitulacijom robe, BPG-om i potpisima („Štampaj karticu (PDF)") je
  „ono što je ispravno i potrebno" — na listama kartica se generički
  tabelarni PDF više ne nudi (`scrIzPrint` sakriven, `scrIzKartPdf` na
  njegovom mestu). Dugmad štampe su sada **komplementarna po listi**:
  kartice → samo legacy kartica; sve ostale liste → samo tabelarni house
  izveštaj. Legacy `PrintKartica*PDF` i dalje netaknut.
- **Širine kolona house štampe po sadržaju.** `EntireColumn.AutoFit` je
  merio i zaglavlje firme / kontekst-liniju u koloni A, pa je prva kolona
  (Datum) dobijala širinu najdužeg teksta strane. Sada se AutoFit radi
  SAMO nad opsegom tabele (header + podaci), a header više nije WrapText
  (AutoFit wrap ćelije ignoriše i lomio je reč — „OTPREMNIC/A"): kolona
  je tačno max(naslov, sadržaj).
- Radno stablo vraćeno na granu Izveštaja; zatečeni polu-gotov rad
  paralelne sesije (Sledljivost) sačuvan je kao WIP commit na
  `claude/sledljivost-ekran` i tamo se nastavlja — jedna sesija po
  radnom stablu.

**Krug 8 — Rang kooperanata (smoke krug 6):**

Za tip Kooperanti treći tab, **Rang** — legacy „Lista kooperanata" /
„Kooperanti po iznosu otkupa" sa Unosa dokumenata, ovde uz **period
zone** umesto fiksne tekuće godine. Račun ostaje jedan:
`modOtkupBlok.KoopRangRows` je dobio Optional granice (bez njih staro
ponašanje — legacy panel i lista na Dokumentima bit-identični; Izveštaji
šalju pun opseg, nikad 0/0). Lista: Rang | Kooperant | Otkupno mesto |
Iznos + `KOP|` identitet (prio 4, za budući drill na karticu); rang broj
je pozicija na celoj listi (pretraga ga ne prepakuje — isto pravilo kao
na Dokumentima); dostupna u oba režima (rang ne zavisi od izabranog
entiteta, pa je guard entiteta preskočen). Test 145: broj redova = broj
kooperanata sa otkupom u opsegu (ručni prolaz `tblOtkup`), zbir = ručni
Σ kg×cena, sortiranost opadajuća, rang 1 na vrhu, period se poštuje
(1990. opseg = prazno); sabotaža `izvestaji-rang-mimo-perioda` (grana
perioda u `KoopRangRows` — legacy pozivaoci je ne dodiruju).

**Krug 9 — Zbirni sadržaj („fali sadržaj za zbirne izveštaje"):**

Nalaz operatera posle kruga 8 — priznat: prenos je verno pratio legacy
matricu, pa je zbirni režim nudio samo Zbirni/Pros. cenu/Manjak, iako je
`ReportAmbalazeZbirni` sve vreme postojao neponuđen (§23.7), a klik na
„Zbirno" ostavljao pojedinačnu listu na hintu. Zatvoreno u tri poteza,
**uz svesnu izmenu matrice** (odluka operatera — izuzetak od „matrica se
ne širi"):

- **AMBALAŽA zbirno** → postojeća legacy grana `ReportAmbalazeZbirni`:
  agregat po tipu gajbe **za izabranog entiteta** (OM/Kupac/Vozač) — zato
  zbirni režim na toj listi jedini zadržava combo entiteta
  (`IzTrebaEntitet`).
- **SALDO i ISPLATA zbirno (OM)** → novi `ReportSaldoOMZbirni` /
  `ReportIsplataZbirniOM`: red = stanica, kolone = UKUPNO red
  pojedinačnog izveštaja te stanice (isti račun, ništa se ne prepisuje;
  stanica čiji su svi brojevi nula se preskače, stanica sa saldom bez
  prometa ostaje), UKUPNO preko svih + `OM|` identitet za budući drill.
  SALDO kupaca zbirno se NE dodaje — tab Zbirni to već daje.
- **Auto-prelaz pri promeni režima** (`PostaviRezim`, isto pravilo kao
  prelaz tipa iz S9; i `PostaviTip` sada bira prvu dostupnu za tekući
  režim): nikad prazan ekran sa hintom kao prvi utisak. Kooperant +
  „Zbirno" prelazi na Rang.

Test 146 (registar ne trpi rupe — CI kapija; grana Sledljivost svoje
testove numeriše od 147 pri rebase-u): red
STA-TEST-2 = ručni prolaz `tblOtkup` (kg) + tri kanala `tblNovac`
(isplaćeno, obrazac T138) + sve kolone = UKUPNO pojedinačnog; zbirna
ambalaža mreže = API zbirni red; auto-prelaz u oba smera. Tri sabotaže:
`izvestaji-zbirno-van-matrice`, `izvestaji-zbirni-saldo-tudji-red`,
`izvestaji-rezim-bez-prelaza`. Usput: `NumVal` je Private u
`modOtkupBlok` — novi Report-i dobili lokalni `IzvNum` (poziv u izrazu je
poznata `vba_check` rupa, uhvatio ga je tek `[break]` na suite-u).

**Krug 11 (smoke kruga 9, nastavak — „fale salda po kupcima, u robi
roba po kupcu"):** isti obrazac kao stanice, sada za kupce: **SALDO
zbirno** = red po kupcu (kg, vrednost, uplaćeno, saldo, amb — iz UKUPNO
reda `ReportSaldoKupci`; prosečna cena se u zbir ne prenosi) i **ROBA
zbirno** = roba po kupcu (UKUPNO kupčevog agregata preko svih vrsta) —
`ReportSaldoKupciZbirni` / `ReportRobaKupciZbirni`, `KUP|` identitet.
Spisak kupaca dolazi **iz podataka** (distinct po nestorniranim
prijemnicama, `IzvKupciIzPodataka`), ne iz šifarnika — kupac sa prometom
bez reda u `tblKupci` mora da se vidi (fixture to namerno drži tako);
naziv iz šifarnika sa fallback-om na ID. Radnja „Štampaj dokument" na
zbirnoj robi je ugašena (agregat bez dokumenta — ista R3 klasa). Matrica:
Kupac-zbirno grana razdvojena od vozačke (+`SALDO_KUPCI`, `OTKUP_ROBA`);
sabotaža `izvestaji-kupci-zbirno-van-matrice`. T146 dopune: red kupca =
UKUPNO pojedinačnog salda (kg, saldo), roba po kupcu = ručni zbir
prijemnica.

**Krug 12 (smoke kruga 11 — „roba po OM u zbirnom; i vozači"):**
**ROBA zbirno za OM** = projekcija zbirnog salda (`ReportRobaOMZbirni`
vraća kolone 1–4 `ReportSaldoOMZbirni` — kg i vrednost su isti izvor,
ne drugi račun, tačno kako je operater primetio); **ROBA zbirno za
VOZAČE** = otpremljeno po vozaču (`ReportRobaVozaciZbirni`: Σ kg i
Σ kg×cena nestorniranih otpremnica u opsegu, naziv iz `tblVozaci` sa
fallback-om na ID, `VOZ|` identitet). Kooperanti zbirno ostaju bez
robe/salda — Rang je njihov zbirni pogled (potvrđena odluka). Matrica:
OM-Z i Vozač-Z + `OTKUP_ROBA`; sabotaže
`izvestaji-vozaci-roba-van-matrice` i `izvestaji-roba-vozaci-storno`
(storno filter koji tiho nestane duplira prevoz). T144 obrnuta istina
(vozač IMA robu zbirno; pojedinačno i dalje ne); T146: roba po OM =
isti ručni prolaz kao zbirni saldo, roba po vozaču = ručni zbir
otpremnica.

**Krug 14 (smoke kruga 13 — „sumarno stanje po tipu za svakog vozača,
dropdown je besmislen"):** zbirna Ambalaža više nije legacy „agregat za
izabranog" nego **svi entiteti tipa × tip gajbe**: red = Entitet | Tip |
Ulaz | Izlaz | Saldo (`ReportAmbalazaZbirnoSvi` — distinct entiteti iz
nestorniranog ledgera, uz isti `DOK_TIP_OTKUP` izuzetak za vozače; po
entitetu se zove postojeći legacy zbirni račun, smerovi se ne
prepisuju). Combo entiteta se u zbirnom režimu krije bez izuzetka
(`IzTrebaEntitet` vraćen na čisto pravilo), kontekst je sada stvarno
„Svi". T141/T146 obrnute istine; `OM|`/`KUP|`/`VOZ|` identitet po redu
za budući drill.

**Krug 18 (poslednji dodatak pre merge-a — filteri vrste i sorte):**
dinamički čipovi **VRSTE** (vrednosti iz podataka: distinct
`tblOtkup`+`tblPrijemnica`, nestornirano, keš po generaciji) na robnim
listama — Otk. listovi, Roba (pojedinačni oblici), Zbirni (ne-Vozač),
Pros. cena, Saldo-Kupac; **SORTE** samo na prijemnicama kupca (jedina
lista čiji snimak nosi sortu — `ReportPrijemniceKupca` dobio sortu kao
10. skrivenu kolonu; ostali Report* je ne vraćaju pa se tamo sorta ne
laže čipom). Ključ čipa nosi vrednost (`vr<vrednost>`/`so<vrednost>`),
poređenje `StrComp vbTextCompare`; ljuska dobila **sirovi natpis čipa**
kroz `~` prefiks (dinamičke vrednosti nemaju kataloški ključ — mala
opšta dopuna, S9-stil). Ne-robne liste i zbirni oblici po entitetu čip
vrste ne nude. Test 149 (spisak čipova po listi + filtriranje = ručni
prolaz + nepostojeća vrsta = 0), sabotaža
`izvestaji-cip-vrste-ne-filtrira` (322 sidra).

**Krug 17 (recenzija, drugi prolaz — lifecycle blocker + hardening):**

- **Aktivacija ekrana primenjuje podrazumevani sort aktivne liste.**
  Povratak na Izveštaje sa aktivnim Rangom je vraćao sort po imenu:
  `ActivateScreen` je tvrdo resetovao 2/desc mimo `SortZaListu` ugovora
  i mimo `mSortLista`. Sada aktivacija zove isti `PrimeniSortZaListu
  ActiveLista()` kao klik i auto-prelaz. T147 (1b) meri aktivacioni
  korak kroz gejtovane seam-ove (`GridScreenSetTest`, `GridSortSetTest`,
  `GridSortAktivacijaTest` — ista procedura, ne kopija); veza
  `ActivateScreen` → procedura ostaje na smoke koraku.
- **`IzvStaniceUnion` fail-visible** (`RequireColumnIndex` za obavezne
  kolone, bez `On Error Resume Next`) — i to je odmah isplivalo pravu
  minu: **VBA `Or` nema kratki spoj**, pa je `cF = 0 Or CStr(d(i, cF))`
  evaluiralo `d(i, 0)` i pucalo za pozive bez filter kolone; stari OERN
  je grešku gutao (uz Resume-Next slučajni ulazak u telo — rezultat je
  bio tačan iz pogrešnih razloga). Uslovi prepisani ugnježdeno.
- Sabotaža `izvestaji-aktivacija-gazi-sort` (aktivacioni korak).

**Krug 16 (recenzija — REQUEST CHANGES, sva četiri zahteva):**

- **R1 (blocker) — Rang se otvara po rangu i u stvarnom UI-ju.** Izbor
  podrazumevanog sorta izvučen u čist ljuskin ugovor
  `modOtkupUI.SortZaListu` (rang-liste → kolona 1 rastuće; ostale →
  kolona 2 opadajuće), koji dele klik na tab **i** `RefreshFromData` pri
  auto-prelazu liste (bez ovoga bi Kooperanti+Zbirno→Rang zadržao tuđ
  sort). Test 147 tvrdi shell sort contract direktno, ne `Scr_Rows`;
  seam `GridSortTest`.
- **R2 — kontekst „Svi" prati LISTU, ne režim:**
  `EntitetNaziv(tip, iD, Not IzTrebaEntitet(kljuc, zbirni))` — Rang je
  „Svi" i u pojedinačnom; „Kooperanti: ()" više ne postoji (test 147).
- **R3 (P1) — orphan stanice ne ispadaju iz „Svi OM":** univerzum
  zbirnog Salda/Isplate sada dolazi **iz podataka**
  (`IzvStaniceIzPodataka`: union `tblOtkup` + OMID iz `tblNovac` +
  Stanica-entiteti iz `tblAmbalaza`, nestornirano; šifarnik samo
  imenuje, fallback ID) — isti princip kao kupci. Fixture dobio
  `OTK-ORPH-1` na `STA-ORPHAN` (bez reda u `tblStanice`); test 148
  tvrdi da se orphan vidi u zbirnom Saldu i Robi pod svojim ID-em.
- **R4 — docs kontradikcije očišćene** (jedanaest lista; Rang deli
  jedan račun sa Dokumentima, nije dupliran).
- Tri nove sabotaže: `izvestaji-rang-sort-ime`,
  `izvestaji-rang-kontekst-prazan`, `izvestaji-om-univerzum-sifarnik`.
- **Presek:** ovo je poslednji krug PR-a #245 — analytics faza (Pregled,
  Poređenje, Pažnja, 360, izvoz) ide kao novi PR.

**Krug 15 (smoke kruga 14 — čipovi i podnožje zbirne ambalaže):**

- **Čipovi Ulaz/Izlaz se na zbirnoj ambalaži ne nude**: red je agregat
  entitet × tip pa skoro svaki ima oba smera — čip po smeru ništa ne
  razdvaja („Sve i Ulaz daju iste brojke"). Pojedinačni ledger ih
  zadržava (tamo filtriraju transakcije).
- **Podnožje zbirne ambalaže**: slotovi **Ulaz / Izlaz u komadima**
  (isti mehanizam kao Uplate/Isplate na izvodima); kg/vrednost se
  nuliraju da se gajbe ne potpišu kao „kg"/„RSD". Uz to mala opšta
  dopuna ljuske (S9-stil, prijavljena): slot podnožja dobio **opcioni
  treći član — ključ jedinice** (default ostaje RSD; postojeći ekrani
  netaknuti), jer je jedinica bila tvrdo „RSD".

**Krug 13 (smoke kruga 12 — saldo po vozaču + brzina P→Z):**

- **Zbirna Ambalaža dobila SALDO kolonu** (ulaz − izlaz po tipu; smer je
  već entitetski jer `isVozac` obrće u Report-u) — „mora da se prikaže
  saldo po vozaču"; nula se prikazuje (izravnat entitet JE informacija);
  saldo je sabirljiv preko tipova (bilans gajbi).
- **Prelaz Pojedinačno→Zbirno ubrzan:** ceo `PuniSnimak` sada ide pod
  postojeći `BeginTableCache`/`EndTableCache` (modDataAccess, ref-counted
  — isti obrazac kao Storno uvid). Zbirni oblici zovu pojedinačni Report
  **po entitetu**, pa su bez keša istu tabelu čitali sa lista N puta —
  sada jednom po snimku (i pojedinačne liste dobijaju: SaldoOM čita 3+
  tabele). `EndTableCache` ide i kroz grešku, inače keš preživi upis.

**Krug 10 (smoke kruga 9):** kontekst-linija zbirne Ambalaže je pisala
„OM: Svi" dok je prikaz bio za podrazumevanog (prvog) entiteta —
`EntitetNaziv` je za zbirni režim vraćao „Svi" pre provere da li lista
traži entitet. Sada ime „Svi" nosi samo lista koja je stvarno preko svih
(`Not IzTrebaEntitet`); tvrdnja u T146 preko novog seam-a
`Scr_IzCtxNazivTest`.

### 23.14 Verifikacija

- `RunAllTests` **144 / 0** (dvanaest novih testova; prva dva runa su
  bila crvena — v. §23.6); `RunBankaImportTestSuite` **205 / 0**
  (bit-identičan BN ekran posle fixture izmena — nova vozila su vezana za
  **zatvoren** blok i kooperanta bez blokova, pa KPI/čipovi/korpa Platnih
  naloga ne vide ništa novo); `vba_check` + `--self-test`,
  `sabotaza --self-test`, `who_writes --check` čisti.
- Fixture: `tblAmbalaza` prvi put ima redove (do sada bi svaka ambalažna
  tvrdnja merila prazan skup); `tblNovac` prvi put ima OMID, sva tri
  kanala isplate i Firma→Otkupac avans — **na STA-TEST-2**, jer
  `T_WriterGuard_AvansSaldoOM` traži da STA-TEST-1 ima avans saldo tačno 0;
  pin `MALINA_MODE=NO` + `KARTICA*_PRINT_MODE=OFF` u `SEF_CONFIG` (ista
  klasa kao `KES_ISPLATE` u §8.10).
- Dvosmerni dokaz: `python tools/dokaz.py izvestaji` — 25 sabotaža (16 iz
  prvog kruga, po jedna za S1/R1/R3, tri za krug 4: čip na nedostupnoj,
  detalj bez stavki, ambalažni broj ostaje ID; tri za krug 5: tabovi ne
  slušaju tip, roba kupca opet agregat, saldo zone iz pogrešne kolone —
  uz `radnja-na-agregatu` prepravljenu da meri novo pravilo „kupac ima
  radnju"), svaka obara tačno jedan imenovani test i vraća se
  bit-identično; plus `banka-nalozi-kes-ignorise-generaciju` za BN stranu
  R1 ugovora.
- **Ručna kapija operatera (traži se izričito):** `Alt+F11 → Debug →
  Compile VBAProject`, pa smoke nad pravim podacima u više krugova
  (checklista u PR-u): izgled zone i prekidača, sve četiri kombinacije
  tip×režim, štampe (tabelarna, kartica, dokument iz reda, revers),
  ponašanje datumskih polja, pretraga na velikoj svesci, brzina Ambalaže
  na punom ledger-u, detalj traka.

---

## 24. Sledljivost — šta je preneto (`v6-ui-188`)

Šesti ekran **Faze E**. Red u registru (`modUiScreens.ScrRows`) je postojao
od `S3a` — stavka menija ANALITIKA → „Sledljivost" se do sada crtala
prigušena jer modula nije bilo. Ovim se piše modul koji taj red već očekuje;
**registar se ne dira**.

Merilo zadatka: **LANAC KOJI SE NE IZMIŠLJA.** Ekran odgovara na dva
pitanja — „od ovog otkupnog lista, gde je roba završila?" (napred:
otkup → otpremnica → zbirna → prijemnica → faktura/kupac) i „od ove
fakture/prijemnice, od kojih kooperanata i parcela je roba došla?"
(nazad). Nepotpun ili višesmislen lanac se prikazuje kao takav
(fail-closed, kao `ReportOtkupRobaOM` pravilo #V>1 i `IZV_NEMA_PRIJEMA`
oznake); kg koji „nestaje" niz lanac je vidljiva razlika sa oznakom,
nikad prećutana.

### 24.1 Gde je šta završilo

| Legacy (`frmSledljivost` + `modSledljivost`) | Novo mesto |
|---|---|
| `TraceByZbirna` (zbirna → otkupi → kooperanti/parcele) | lista **PARCELE** (ista zrna kao LANAC, sertifikaciona projekcija: kooperant, BPG, kat. broj, kultura, ha, GGAP) |
| `cmbZbirna` filter po zbirnoj | **pretraga ljuske** — haystack nosi SVE brojeve lanca (v. 24.2/B) |
| `GetUnlinkedOtkupi` (NEPOVEZANI OTKUPI lista) | klasa `OTKUP-BEZ-OTPREMNICE` u listi **NEPOTPUNI** + oznaka `nepovezan` u LANAC-u |
| `PrintTracePDF` (FillSledljivostSablon) | dugme zone **„Lanac (PDF)"** — house PDF lanca IZABRANOG reda sa kontekst-linijom (koren · opseg · kompletnost) — i, od drugog/trećeg smoke-a (S4/S5), dugme **„Sledljivost (PDF)"**: mete po tome gde roba jeste (`ReportSledljivostMete`) — zbirna kroz POSTOJEĆI šablon (izvučena ruta `modIzvestaj.StampajSledljivostZbirne`), sveža paleta kroz paletni list, prerada kroz preradni list. `SLEDLJIVOST_PRINT_MODE` se poštuje, OFF se prijavljuje. Forma zadržava svoju kopiju rute, netaknuta (Faza B) |
| `btnAutoLink` / `btnPovezi` (upis!) | preneto u drugom krugu (S3): dugme **„Poveži automatski"** na NEPOTPUNI (isti `AutoLinkOtkupOtpremnica_TX`, toast sa brojem) + radnja **„Poveži…"** nad redom klase `OTKUP-BEZ-OTPREMNICE` (kandidati `GetOtpremnicaKandidatiZaOtkup` — ista stanica + isti datum; upis `ReassignOtkupToOtpremnica_TX`) |
| — (legacy nema) | lista **LANAC**: 1 red = 1 otkupni list sa razrešenim karikama kao kolonama; lista **NEPOTPUNI**: 1 red = 1 problem karike (7 klasa); „Štampaj izveštaj" (house PDF aktivne liste) |

Novi računi žive u `modIzvestaj` (obrazac `ReportPrijemniceKupca`):
**`ReportSledljivostLanac(od, do)`** — zrno otkup, 26 dokumentovanih kolona
(karike + parcela projekcija + kg po karici) — i
**`ReportSledljivostProblemi(od, do)`** — zrno problem, 8 kolona (klasa,
karika, nosilac, kg, detalj sa brojkama, DokTip+DokID za rutu štampe).
Ekran je prikaz nad ta dva računa i ne čita tabele sam (detalj i PDF idu
iz snimka).

### 24.2 Odluke (A–E iz zadatka)

**(A) Inventar veza — iz `modConfig`/`WHO_WRITES`, ne iz sećanja.**
otkup→otpremnica = `tblOtkup.OtpremnicaID` (piše 12 modula, među njima
`modSledljivost` auto-link; jednoznačan ID; prazan = nepovezan);
otpremnica→zbirna i zbirna→prijemnica = `BrojZbirne` — broj NIJE identitet
(dve aktivne zbirne legitimno dele broj), vlasnik je broj+vozač+kupac, a
razrešenje je ISTO pravilo koje dele `ReportOtkupRobaOM` i `ReportManjak`:
`BuildManjakDict` (#V/#1/#O ključevi) + `PrijemZaZbirnu` — **poziva se,
nikad ne prepisuje** (u lancu kroz izdvojeni `SledResolveZbirna` koji
zove iste ključeve). prijemnica→faktura = denorm `FakturaID` (isti čitač
kao ekran Fakturisanja; `tblFakturaStavke` je normativ i ne čita se).
otkup→parcela = `ParcelaID` → `tblParcele` (aktivnost se ne filtrira —
sledljivost je istorijska). Prag kg = **0.01**, ista vrednost kao
(privatni) `modDokumentInvariant.EPS_KG` — dva mesta, jedan prag, uz
komentar na oba. **Otkupov denorm `BrojZbirne` se NE koristi za lanac** —
služi samo za proveru saglasnosti (raskorak = oznaka `veza neusaglasena`);
premošćenje kroz njega je tačno klasa laži koju merilo zabranjuje (fixture
vozilo: `OTK-TEST-2` tvrdi ZB-TEST-3 koju njegova otpremnica nema).

**(B) Tri liste, bez tipa/režima/entiteta.** LANAC (default) · PARCELE ·
NEPOTPUNI. Oba smera pitanja odgovara ISTO zrno (otkupni list sa
karikama kao kolonama), pa smer ne traži prekidač: **pretraga ljuske**
nosi sve brojeve lanca (otkup, otpremnica, zbirna, prijemnica, faktura,
kooperant, kupac, stanica, oznaka) — ukucaš broj fakture i čitaš
kooperante/parcele iz redova (nazad); ukucaš kooperanta i čitaš karike
udesno (napred). Bez entitet comboa nema ni S1 klase zamki (PopIndex).
Sve tri liste su UVEK dostupne — matrica nedostupnosti ne postoji, a
prazna lista kaže zašto i kuda: „nema otkupa u periodu → proširi period"
≠ „sve karike potpune" (dobro stanje na NEPOTPUNI listi, i tako se i
kaže — nikad pun naslov nad trajno praznom listom).

**(C) Lanac u ravnoj mreži = kolona-nivo.** 1 red = 1 otkupni list;
karike su kolone sa brojevima dokumenata (Otpremnica, Zbirna, Prijem,
Faktura, Kupac) + kolona **OZNAKA** (prva prekinuta/višesmislena karika
po poziciji u lancu: `nepovezan` → `otpremnica stornirana` → `veza
neusaglasena` → kg razlika blok↔otp → `bez zbirne`/`zbirna ne postoji`/
`nejasan vlasnik` → kg otp↔zbirna → `nema prijema` →
`faktura neusaglasena` (krug 9: SAMO lažna tvrdnja „Da“ — `Ne` je
legitiman tok u hladnjaču); prazno = potpun; kg zbirna↔prijem se ne proverava —
to je transportno kalo, v. §24.7/S1). Prijem ćelija nosi broj prijemnice,
„N prij." kad ih je više, ili ostaje prazna (razlog je u OZNAKA koloni —
nikad izmišljen broj, nikad „0,00" umesto poruke). **Detalj traka** desno
u zoni (obrazac §23.11/S7): pun lanac izabranog reda, karika po karika
**sa kg po karici** — samo ono što red ne pokazuje (S10): vozač, stanica,
parcela (na LANAC listi), kg otpremnice/zbirne/prijema, kupac uz zbirnu.
Padajući redovi mreže ostaju Faza C.

**(D) Identitet + radnje.** LANAC/PARCELE: `OTK|<OtkupID>` u poslednjoj
koloni prio 4 (`GridCell`, mapa „prikaz → ID" ne postoji); NEPOTPUNI:
`DokTip` + `DokID` u dve prenosne kolone (obrazac AMBALAZA §23.4). Jedna
radnja nad redom — „Štampaj dokument": LANAC/PARCELE štampaju otkupni
list (zrno reda, `ReprintOtkupniListByOtkupID`); NEPOTPUNI rutira po
vrsti karike (OTK→otkupni list, OTP→`OutputOtpremnicaPDF`,
PRJ→`PrintPrijemnica`; **zbirna nema svoju štampu** — vrsta `Zbirna`
postoji da radnja ume da ODBIJE s razlogom, legacy Case Else obrazac).
Zona: „Štampaj izveštaj" (house PDF aktivne liste, tačno ono što se
vidi, sa politikom sabirljivosti: LANAC/PARCELE sabiraju samo kg;
površina parcele se NE sabira — atribut koji se ponavlja po redu iste
parcele; NEPOTPUNI ne sabira ništa — kg meša zrna karika) i „Lanac
(PDF)". Brojač menija = 0 (read-only pregled kao Izveštaji §23 — ništa
ovde ne čeka operatera; broj nepotpunih karika je KPI zone, uz broj
potpunih lanaca, oba iz snimka, `—` pre prvog čitanja).

**(E) Keš od prvog dana.** JEDNO punjenje po ključu konteksta (`od|do`)
puni **obe** polovine snimka (lanac + problemi) → prelaz na bilo koju
listu, čip i svaki otkucaj pretrage su re-filter nad snimkom, nula
čitanja tabela (§22.9/N7). Invalidacija: `Scr_ResetCache` +
**generacija podataka** (`modUiData.DataGeneracija`, §23.10/R1) od
prvog dana. `mSnimakPunjenja` broji punjenja (test: tri pretrage + čip +
sve tri liste = jedno); `Diag_SlRedovi` (Alt+F8) od prvog dana. U
Reportima nijedan `LookupValue` po redu — mape pre petlje
(`BuildLookupDict` ×6 + rečnici otpremnica/parcela/prijemnica,
`BuildManjakDict` jednom po punjenju).

### 24.3 Klase problema (NEPOTPUNI) i čipovi

Sedam klasa, svaka fail-closed nalaz sa tačnom karikom (DokTip+DokID) i
detaljem sa brojkama: `OTKUP-BEZ-OTPREMNICE` (i veza na storniranu/
nepostojeću — detalj razlikuje), `VEZA-NEUSAGLASENA`,
`OTPREMNICA-BEZ-ZBIRNE` (i broj bez aktivne zbirne), `BROJ-ZBIRNE-
DVOSMISLEN` (JEDNOM po broju), `ZBIRNA-BEZ-PRIJEMA` (po stavki; uz #V>1
sa nepripisivim prijemnicama se NE tvrdi — prijem možda postoji a ne sme
se pripisati), `FAKTURA-VEZA-NEISPRAVNA` (krug 9: SAMO „Fakturisano=Da“
bez validne aktivne fakture — `Ne` NIJE klasa; pokriva i poznato stanje
„Fakturisano=Da bez FakturaID" — PRJ-FAK-2 klasa), `KG-RAZLIKA` (SAMO
podatkovne karike: blokovi↔otpremnica i otpremnice↔zbirna — roba se
nije mrdala pa brojevi moraju biti isti; razlika zbirna↔prijem je
TRANSPORTNO KALO i ne prijavljuje se — meri je Manjak, v. §24.7/S1; uz
#V>1 se ne računa — fail-closed i za kg). Čipovi: LANAC
sve·potpun·nepotpun (lanac koji curi NIJE potpun); PARCELE sve·bez
parcele (obrazac MANJAK čipa); NEPOTPUNI sve·veze·prijem·fakture·kg
(grupe klasa). Prvi čip je svuda najširi.

### 24.4 Slaganja — srce zadatka

Testovi **150–155** u `modTest` (izvršavaju se pre mutirajućih 124–126;
Izveštaji su kroz krugove 8–18 zauzeli 145–149, pa numeracija ide odmah
iza njih — registar ne trpi rupe),
sve tvrdnje su relacije nad SLED-* vozilima koja nijedan drugi test ne
dira:

| Test | Tvrdnja | Nezavisan read-model |
|---|---|---|
| 151 `T_Sled_LanacSlaganje` | potpun lanac karika po karika == ručni prolaz kroz `tblOtkup`→`tblOtpremnica`→`tblZbirna`→`tblPrijemnica`→`tblFakture` (samo `GetTableData`+`GetColumnIndex`); kg se slaže niz CEO lanac (blokovi = otpremnica = zbirna = prijem); nazad: pretraga po broju fakture vraća OBA kooperanta lanca; PARCELE projekcija == `tblParcele`; blok bez parcele nosi oznaku; KPI == isti snimak | ručni prolazi |
| 152 `T_Sled_FailClosed` | dvosmislen broj → `nejasan vlasnik` (kg OSTAJE prazan); lažna tvrdnja fakturisanosti → `faktura neusaglasena` (a `Fakturisano=Ne` je POTPUN lanac — krug 9); kg curenje → `kg razlika`; raskorak denorma → `veza neusaglasena`; nepovezan; veza na storniranu; storniran dokument nije NIGDE; svaka klasa problema sa tačnom karikom; detalj kg razlike nosi obe brojke | pokvarena vozila fixture-a |
| 150/153/154/155 | ugovor ekrana; identitet prio 4 + DokTip ruta; keš/generacija/kvake pretrage/hint; zona posle Unload-a | — |

Tvrdnje nad deljenim pravilom vlasnika (`BuildManjakDict`/
`PrijemZaZbirnu`) i dalje čuvaju postojeći testovi (`RunIzvestajTests`,
T136/139) — sabotaža nad njima bi obarala tuđe. **Nove** Report funkcije
nemaju tuđe pokriće, pa 19 sabotaža (`sledljivost-*`) gađa i Report
polovinu i ekransku: premošćenje zbirne, sabiranje dvosmislenog broja,
gutanje kg praga (lanac i problemi zasebno), storniran ulazi, klasa
ispada iz problema, identitet se crta, tuđa vrsta karike, keš puni
iznova, generacija se ignoriše, sirov haystack, tri čipa, zona bez
dugmeta, PDF bez oznake, KPI iz pogrešnog izvora, detalj bez karika.

### 24.5 Fixture

SLED-* vozila (v. blok konstanti u `make_fixture.py`), sva na
STA-TEST-2, svi blokovi ZATVORENI (`NOV-SLED-*` virmani, OMID=STANICA2 —
novi novčani redovi samo na STA-TEST-2; `T_WriterGuard_AvansSaldoOM`
preduslov netaknut, virman firma→koop ne dira avans pool; Platni nalozi
bit-identični, isti razlog kao `OTK_IZV_ZATVOREN`):

| Vozilo | Zašto postoji |
|---|---|
| OTK-SLED-1 (KOOP-TEST-2) + OTK-SLED-2 (KOOP-TEST-IME) → OTP-SLED-1 → ZB-TEST-SLED → PRJ-SLED-1 → FAK-SLED-1 | **prvi POTPUN lanac u fixture-u** (zatvara rupu iz §23.12/S10 — do sada nijedna otpremnica sa blokovima nije imala zbirnu sa nestorniranom prijemnicom); dva kooperanta na istoj otpremnici („nazad" mora vratiti oba); kg 300+200=500 slaže se niz ceo lanac; OTK-SLED-2 je bez parcele (oznaka na PARCELE listi) i nosi ga ISTOIMENI kooperant (identitet je OTK\|id, ne prikazano ime). KOOP-TEST-3 se ne sme koristiti ni za jedan SLED blok: `T_BankaUvoz_RucnoMapiranjePravila` broji njegove blokove apsolutno (=5) — prvi crveni krug je to i pokazao |
| OTK-SLED-D → OTP-SLED-D (vozač PRAZAN) → ZB-TEST-SLDD (svoj par, dva vozača) | dvosmislen broj koji se NE može razrešiti po vozaču otpremnice → `nejasan vlasnik`. Svoj par, ne ZB-TEST-DUPL: DUPL troši raniji storno test (ostane jedan aktivan vlasnik), pa u trenutku sledljivost testova više nije dvosmislen |
| OTK-SLED-N → OTP-SLED-N → ZB-TEST-SLN → PRJ-SLED-N (Fakturisano=Ne) | krug 9: LEGITIMAN tok u hladnjaču — lanac POTPUN, bez oznake i bez problema (kvar fakture mere F-vozila) |
| OTK-SLED-R (100) → OTP-SLED-R (250!) → ZB-TEST-SLR (bez prijemnice) | kg curi na prvoj karici → `kg razlika` + `KG-RAZLIKA` problem sa obe brojke; zbirna bez prijema na svojoj mirnoj zbirni |
| pin `SLEDLJIVOST_PRINT_MODE=OFF` u SEF_CONFIG | klik „Lanac (PDF)" u testu/smoke-u ne pravi PDF; ekran OFF prijavljuje porukom |

Vozila bez ambalaže (KolAmbalaze se ne seje) — kanonski amb saldo i
kartice ne dobijaju kretanja bez ledger parova (§23.6 nalaz 1).

### 24.6 Šta NIJE preneto, i zašto

- ~~Povezivanje NE ulazi u v1~~ — **oborio drugi smoke (S3)**: pregled
  bez alata terao je operatera nazad u staru formu. Povezivanje je sada
  na ekranu (auto dugme + radnja „Poveži…", v. 24.7/S3), ali **upis i
  dalje ide isključivo kroz postojeće TX kapije**
  (`AutoLinkOtkupOtpremnica_TX`, `ReassignOtkupToOtpremnica_TX` — nisu
  dirani); legacy `frmSledljivost` ostaje operativna i nepromenjena (dve
  kopije žive namerno — §5/Faza B).
- **Oznaka `zbirna ne postoji`** (broj na otpremnici bez ijedne aktivne
  zbirne) postoji u kodu, ali **nema fixture vozilo** — ne tvrdi se
  testom (zapisano, kao prijemnica-linija u §23.12/S10).
- **Oznaka `nema prijema` u LANAC koloni nema vozilo kod koga je PRVA
  anomalija** (SLED-R curi kg pre prijema; ZB-TEST-1 lanac isto) — klasa
  `ZBIRNA-BEZ-PRIJEMA` u problemima jeste pod testom (ZBI-SLED-R).
- **Detalj trake za ne-otkup karike NEPOTPUNI liste** (OTP/ZBR/PRJ
  redovi) ostavlja traku praznu — uzvodni sažetak iz snimka je zapisan
  kao dorada, ne izmišlja se za v1.
- **Vrednost (RSD) niz lanac** — v1 prati robu (kg), ne novac; podnožje
  nosi kg, „Vrednost 0,00" slot ostaje kao na ambalažnim listama.
- ~~Paletna sledljivost ostaje na ekranu Palete~~ — **oborio treći smoke
  (S5)**: izveštaj o sledljivosti se dokazuje dokumentom na kome roba
  STVARNO jeste (zbirna / paleta / prerada), pa dugme „Sledljivost (PDF)"
  sada nudi sve tri mete (v. 24.7/S5). Paletna OPERATIVA (zatvaranje,
  storno, prevezivanje, fizička štampa) i dalje živi isključivo na
  ekranu Palete; palete nisu postale karika lanca u mreži.
- **Owner-scoped trace kroz palete/preradu** (krug 8 R3): paletna
  stavka nosi zbirnu samo kao BROJ, pa se za dvosmislen broj paletni
  tok ne može pripisati vlasniku — v1 ga fail-closed odbija
  (`ZBIRNA-NEJASNA`). Pravo rešenje traži jači identitet u
  `tblPaletaStavka` (npr. ZbirnaID), što je izmena šeme i writer-a —
  ne izmišlja se ovde.

### 24.7 Smoke: nalazi operatera (ispravke u istom PR-u)

Compile je prošao, ekran radi na pravoj svesci (malina sveska, 1.620
lanaca) — i doneo nalaze koje suite nije mogla da vidi:

**S1 — „kg razlika između zbirne i prijemnice je transportno kalo“.**
Operater je odmah pročitao ono što je dizajn prevideo: razlika
zbirna↔prijem *skoro uvek postoji* — roba putuje, kalo je poslovna
veličina (meri je izveštaj Manjak), ne kvar podataka. Provera te karike
je uklonjena i iz oznake lanca i iz liste NEPOTPUNI; kg po karici ostaje
vidljiv u detalj traci (razlika se VIDI, ali se ne optužuje). Kg provere
ostaju samo na podatkovnim karikama gde se roba nije mrdala: blokovi↔
otpremnica i otpremnice↔zbirna (invarijanta `modDokumentInvariant`).
Brojka „Nepotpune karike“ time pada na stvaran posao. Prag se NE
uvodi (dozvoljeno kalo bi bilo novo poslovno pravilo — §22.2 merilo).

**S2 — detalj traka je vizuelno curila u susedni blok.** Redovi trake
(do 6 linija ispod naslova) izlazili su ispod bele podloge zone i
vizuelno ulazili u naredni blok ekrana. Bela kartica sada obuhvata celu
traku (dno ide do pred donju liniju zone), pa traka ima svoj okvir.

**S3 — „gde je nestalo povezivanje nepovezanih?"** (drugi krug).
Odluka „pregled bez upisa" (bivši §24.6 prvi red) u praksi znači:
ekran ti pokaže nepovezan otkup, a za popravku moraš u staru formu.
Vraćeno, uz istu podelu koju drži ceo projekat — **UX na ekranu, upis u
postojećim TX kapijama**: dugme „Poveži automatski" (samo na listi
NEPOTPUNI; zove `AutoLinkOtkupOtpremnica_TX`, toast kaže koliko je
povezano) i radnja „Poveži…" nad redom klase `OTKUP-BEZ-OTPREMNICE`
(svaki drugi red odbija porukom). Kandidati su legacy pravilo stare
forme, izvučeno kao read-only `GetOtpremnicaKandidatiZaOtkup` (ista
stanica + isti datum, bez storniranih); izbor kroz InputBox radnje
(presedan „Iznos…" na Platnim nalozima), do 15 kandidata uz prijavljen
preliv; upis kroz `ReassignOtkupToOtpremnica_TX` (kapije cilja ostaju u
writeru); posle upisa liste se odmah preračunaju. Test 156 +
sabotaža `sledljivost-kandidati-bez-stanice` čuvaju pravilo kandidata.

**S4 — „sledljivost ima već definisanu formu za PDF."** Tačno:
`SledljivostSablon` (radni list šablona) je štampa koju operater već
poznaje. Dodato dugme „Sledljivost zbirne (PDF)" — za zbirnu izabranog
reda (iz bilo koje liste; red bez zbirne odbija porukom) puni ISTI
šablon kroz izvučenu rutu `modIzvestaj.StampajSledljivostZbirne`
(`TraceByZbirna` + `FillSledljivostSablon`, zaglavlje iz prve aktivne
zbirne tog broja, zbir kg prijemnica). House „Lanac (PDF)" ostaje —
dva pogleda: karike jednog otkupa vs. cela zbirna po šablonu.

**S5 — „izveštaj o sledljivosti ide za zbirnu, paletu ili prerađenu
paletu — po tome gde je roba"** (treći krug). Šablon zbirne dokazuje
sledljivost robe prodate dalje kao SVEŽE; roba u magacinu sveže robe
dokazuje se PALETNIM listom, a roba prerađena (prodata kao prerađena
ili u magacinu prerađene robe) PRERADNIM listom — ta dva dokumenta su u
legacy toku već nosila kooperante. Dugme (preimenovano u **„Sledljivost
(PDF)"**) sada za zbirnu izabranog reda skuplja **mete** kroz novi
read-only `modIzvestaj.ReportSledljivostMete`: red 1 je uvek zbirna
(šablon), zatim nestornirane NEprerađene palete čija stavka nosi taj
`BrojZbirne` (podatkovna veza — paletna stavka od prvog dana čuva broj
zbirne), pa prerade čija stavka pokazuje na nađene palete (join po
`PaletaID`, isto pravilo kao `modIntegritet` D2). Ništa se ne
premošćuje: prerađena paleta bez preradne stavke ne izmišlja metu
(D2 je prijavljuje), stornirana paleta nije meta ni kad joj je stavka
aktivna. Jedna meta ide odmah; više njih se bira u POLJU (v. S6 —
InputBox verzija izbora je živela pola dana i oborena je u istom
smoke-u). Rutiranje: zbirna kroz postojeću rutu (poštuje ceo
`SLEDLJIVOST_PRINT_MODE`); paletni/preradni list kao PDF
(`ExportPaletniListPDF` / `ExportPreradaPDF` — dugme je „(PDF)",
fizička štampa tih listova ostaje na ekranu Palete). OFF režim se
prijavljuje na ulazu. Test 157 + sabotaže
`sledljivost-mete-storniranu-paletu` i
`sledljivost-mete-preradjena-kao-sveza`; fixture dobija zatvorenu svežu
paletu na SLED lancu, prerađenu paletu sa preradom na SLN lancu i
storniranu paletu kao negativ (sve ZATVORENE — otvorene bi ušle u
gajbe-do-zatvaranja račun ljuske).

**S6 — „mora da postoji jasno polje za izbor sa dropdown i filter
poljem"** (isti dan, nad InputBox izborom meta iz S5). Zona dobija
polje **„Dokument sledljivosti"** (`NewFieldG "cmb"` — isti okvir +
combo + strelica kao kooperant na Agrohemiji; strelicu ljuska već
vezuje na `DropDown`). Ponudu daje novi read-only
`modIzvestaj.ReportSledljivostDokumenti(od, do)`: SVI dokumenti
sledljivosti perioda — zbirne po DISTINCT broju (štampa šablona je po
broju — legacy `cmbZbirna` semantika; dvosmislen broj se nudi jednom),
nestornirane NEprerađene palete, nestornirane prerade; dokument bez
validnog datuma se UKLJUČUJE (vidljiva anomalija bolja od tiho
sakrivene). Ponuda ide u treći slot ISTOG snimka konteksta (lanac +
problemi + dokumenti = jedno punjenje — §23.10/R1 važi i ovde), pa je
kucanje 0 čitanja. Kucanje filtrira substring pretragom sa kvake-foldom
(`modUiData.TekstZaPretragu`, kao mreža — broj, datum, tip, status…) i
drži dropdown otvoren da se sužavanje vidi; `MatchEntry` je isključen
(prefix-autocomplete bi se tukao sa substring filterom). Dugme
„Sledljivost (PDF)": razrešen izbor polja ide PRVI; kucan a nerazrešen
tekst ODBIJA porukom (pravilo nerazrešenog izbora — ne pogađa se);
prazno polje pada na izabrani red (jedna meta odmah, više njih poruka
upućuje na polje). Test 158 (`SlDokPrikazi` je čist sklop nad snimkom,
testabilan bez kontrole) + sabotaže
`sledljivost-dokumenti-storniranu-paletu` /
`sledljivost-dokumenti-preradjena-kao-sveza`.

**S7 — „vidi se samo jedan red, ispod je sivo"** (četvrti krug, nad
poljem iz S6). Dva kvara istog polja: tvrda širina kolone (`244 pt`)
bila je šira od liste → horizontalni klizač i odsečen tekst; a
programski `DropDown` na svako slovo ostavljao je suženu ponudu u
prozoru STARE visine (MSForms ne preračuna visinu već otvorene liste) —
otud jedan red i sivo. Ispravka: širina kolone prati kontrolu, a
dropdown se više ne otvara programski — kucanje sužava listu u mestu,
strelica je otvara (tada je visina uvek tačna).

**S8 — „filter u redu nepotpune karike ne radi."** Provereno celim
putem: čipovi (ljuskin `ChipFilter` → `SlCipProblemi`), pretraga i
paginacija (`ReloadGrid` vraća stranu 1) — ispravni. Stvarni raskorak:
period-linija obećava „pretraga nalazi svaki broj u lancu", a haystack
NEPOTPUNIH je nosio samo broj/nosioca/detalj sopstvenog reda — broj
zbirne ukucan na Nepotpunima nije nalazio njene nefakturisane
prijemnice. Ispravka: `ReportSledljivostProblemi` dobija kolonu 9
(lanac-brojevi reda koji nisu već u kolonama 3/6 — npr. broj zbirne
nefakturisane prijemnice, denorm broj zbirne nepovezanog bloka) i ona
ulazi u haystack. Tvrdnja u testu 154 + sabotaža
`sledljivost-nep-pretraga-bez-lanca`.

**S9 — „povezivanje treba lepše rešiti od postojećeg stanja."** InputBox
kandidata (krug 2) zamenjen istim obrascem kao S6: polje **„Otpremnica
za povezivanje"** (combo, vidljivo samo na NEPOTPUNIMA — isti uslov kao
dugme „Poveži automatski"). Izbor reda klase `OTKUP-BEZ-OTPREMNICE`
puni polje kandidatima (legacy pravilo, `GetOtpremnicaKandidatiZaOtkup`;
stavka se NE bira automatski — S1 klasa zamki), radnja „Poveži…" čita
RAZREŠEN izbor polja (delimičan tekst odbija porukom; polje čuva i
odbranu od ustajalog para red/polje po `OtkupID`). Upis nepromenjen:
`ReassignOtkupToOtpremnica_TX`; posle upisa polje se prazni i liste se
preračunaju.

**S10 — doterivanje polja kandidata** (peti krug; dropdown iz S7 je
potvrđen kao dobar). Dva nalaza: „čudne razlike u fontu" — font oba
comba sada EKSPLICITNO prati datumska polja (ne oslanja se na default
kontrole), a stavke kandidata dobijaju isti ritam separatora kao polje
dokumenta (dva razmaka pre tačke, jedan posle — mešani razmaci su
čitani kao „drugi font"); i „kad se klikne na Poveži automatski otvori
dropdown" — „Poveži…" bez razrešenog izbora sada uz poruku i OTVORI
listu kandidata (fokus + `DropDown`; lista je tada zatvorena pa S7
zamka ne važi; `SetFocus` iza `IsTestMode` — u nevidljivom Excelu visi).

**S11 — lanac PDF: isečen naslov, tuđe zaglavlje, dvotačke** (peti
krug, sa priloženog PDF-a — operater nije morao ni da kaže). Tri kvara
jednog dokumenta: (1) naslov i kontekst-linija su MERGE preko kolona
tabele (`DocTitleBlock`) i seku se na njenu širinu — uska 4-kolonska
tabela im je sekla tekst sa obe strane („EDLJIVOST — LANAC DOKUMENT");
kompozer je zajednički i ne dira se, pa tabelu širi SIMETRIČAN pad
heder ćelija (AutoFit meri max(header, sadržaj); centriran tekst +
jednak pad levo/desno = isto zaglavlje na oko). (2) Kolona karika je
nosila heder „PROBLEM" (pozajmljen od liste NEPOTPUNI) — sada „KARIKA"
(novi ključ). (3) Karike „Zbirna:"/„Faktura:" su stajale sa dvotačkom
pored „Otkup" bez nje — deljeni detalj-ključevi se ne menjaju (nose ih
i trake Izveštaja), dvotačka se skida na mestu upotrebe
(`BezDvotacke`); tvrdnja u testu 153 + sabotaža
`sledljivost-pdf-dvotacka`.

**S12 — „vidi se stari dropdown preko modernog"** (šesti krug). Ljuska
za SVE combo kontrole ima svoj panel (`zPop`: kucanje ga otvara i
sužava po podnizu — `PopFromTyping`/`PopIndex`; strelica ga otvara;
izbor postavlja `ListIndex`), a nativna MSForms lista je penzionisana
(„nativna lista se više NE koristi — OpenPopup pokriva i duge liste").
Moj `CB.DropDown` iz kruga 5 crtao je baš tu penzionisanu listu PREKO
panela. Auto-otvaranje na „Poveži…" sada ide kroz FRONT DOOR ljuske:
`modOtkupUI.UiEvent "scrSlPov", "Drop"` (javni dispečer → `OpenPopupFor`
→ panel); `UiEvent` sam izlazi bez forme, pa test-režim nije ugrožen.

**S13 — „ekran se prvi put dosta sporije učitava od ostalih"** (šesti
krug). Dva troška istog polja: punjenje ponude od 1000+ dokumenata išlo
je kroz 2N `AddItem` COM poziva — sada JEDAN upis liste
(`CB.List = matrica`); i ekranski chg-filter je na svako slovo iznova
prepisivao listu (opet 2N poziva) — a ljuskin panel isto suženje već
radi sam, po podnizu nad prikazom. Ekranski filter je UKLONJEN
(`SlDokPrikazi(q)` → `SlDokPonuda()` bez filtera): prikaz stavke zato
nosi sve po čemu se traži (tip, broj, datum, opis), što test 158 sada i
tvrdi. Kvake-fold pretrage ovim prelazi na ljuskino pravilo
(`vbTextCompare` podniz — isto kao kooperant combo na Agrohemiji).

**S14 — „lanac mora ozbiljnije formatirati — ovako kao detalj na A4
listu je smešno"** (šesti krug). „Lanac (PDF)" više ne ide kroz house
kompozer lista (on je za ŠIROKE liste; uskoj 5-rednoj tabeli je i posle
pad-a davao izgled zalutalog detalja) nego kroz novi dokument-builder
`modIzvestaj.StampajSledljivostLanacDoc` po ugledu na
`SledljivostSablon`: zaglavlje firme + naslov (`DocSellerHeader`/
`DocTitleBlock`), info blok korena (otkupni list, kooperant, stanica,
datum · vozač, kupac, period), tabela karika sa NOSIOCIMA (kolona 3:
kooperant/vozač/kupac po karici — `SlLanacZaPdf` sada 5 kolona + info
blok u paketu), red kompletnosti (LANAC POTPUN / LANAC STAO NA: …) i
podnožje datum štampe + potpis/pečat. Sopstveni list `_SlLanacPrint`,
eksplicitne širine kolona, A4 `FitToPagesWide=1`; pad-headeri iz S11
uklonjeni (builder ima prave širine). Test 153 dopunjen (nosilac +
info blok + oznaka na koloni 5).

**S15 — „u dropdown-u dokumenta je 8. red uvek većeg fonta"** (sedmi
krug) — **NIJE reprodukovano iz izvora** (CLAUDE.md §2: nalaz se
prijavljuje, ne nagađa se zakrpa). Statički: `BuildPopup` gradi svih 14
`pop*` redova identično, `PopRender` menja samo
caption/boju/top/širinu/vidljivost, hover (`clsFlatBtn`) samo boju —
font niko ne dira. Empirijski (COM sonda nad fixture-om): posle gradnje
forme svih 14 redova je identično `size=10 Segoe UI bold=False h=21` —
mutacija se, ako postoji, dešava tek u ŽIVOJ upotrebi (headless se
panel ne može otvoriti: `mFrm` postavlja tek `ShowOtkupUI`). Isporučena
dijagnostika **`Alt+F8 → Diag_SlPopFont`** — čita stvarno stanje svih
redova otvorene forme (Immediate ispis); pušta se ODMAH pošto se veći
red vidi. Ako ispis pokaže odstupanje svojstva, kvar je u LJUSCI
(deljeni `zPop` panel) — ispravka tamo je odstupanje od „`modOtkupUI` =
pečat-only" pravila ove migracije i ide uz izričitu odluku operatera;
ako su svojstva uniformna a red vizuelno veći, uzrok je u iscrtavanju
(DPI/font-substitucija) i traži drugi alat.

**Krug 8 — spoljna revizija (REQUEST CHANGES → završni paket R1–R7,
bez novih funkcija):**

- **R1 — fakturisanost je ALL, ne ANY.** Lanac je govorio „potpun" čim
  JEDNA prijemnica zbirne ima fakturu; sada SVAKA razrešena prijemnica
  mora imati `Fakturisano=Da`, neprazan `FakturaID` i taj ID mora
  pripadati POSTOJEĆOJ AKTIVNOJ fakturi (`SledFakMapa` sa
  `ExcludeStornirano` — `BuildLookupDict` stornirane nije izbacivao).
  Jedna neispravna obara celu kariku, a NEPOTPUNI dobija i treću granu
  („FakturaID … nije među aktivnim fakturama") — dva read-modela istog
  ekrana se SLAŽU. Fixture: F-lanac (2 prijemnice: aktivna + mrtva
  veza); test 152 + sabotaža `sledljivost-fakture-any-umesto-all`.
- **R2 — SearchRefs kolona 27.** Prikaz „N prij."/„N fakt." guta
  pojedinačne brojeve, pa smer nazad po drugoj fakturi nije radio.
  `ReportSledljivostLanac` sada u koloni 27 nosi SVE brojeve prijemnica
  i faktura reda; haystack LANAC i PARCELE je koriste (PARCELE dobija i
  kupca). Fixture: M-lanac (1 zbirna → 2 prijemnice → 2 fakture,
  potpun); testovi 152/154 + sabotaža
  `sledljivost-lanac-pretraga-bez-refs`.
- **R3 — dvosmislen broj zbirne je fail-closed SVUDA.** Šablon:
  vlasnička kapija (`SledVlasnikaBroja` — broj RAZLIČITIH aktivnih
  vlasnika vozač+kupac) PRE režima štampe vraća `DVOSMISLEN`, a header
  ide iz AKTIVNE zbirne (`ExcludeStornirano` — prvi red je umeo biti
  stornirana generacija). Mete: jedina meta je `ZBIRNA-NEJASNA`, bez
  štampe (i bez paleta po prostom broju — i stavka palete ga nosi kao
  BROJ). Dropdown: dvosmislen broj je stavka „Zbirna — nejasan vlasnik"
  BEZ kg (nikakvog sabiranja tuđih vlasnika) i bez print akcije.
  Testovi 157/158 + sabotaža `sledljivost-sablon-dvosmislen-broj`.
  Pravi owner-scoped trace kroz palete tražio bi jači identitet u
  `tblPaletaStavka` — zapisano u §24.6, ne izmišlja se u v1.
- **R4 — `IIf(IsDate, CDate, …)` mina uklonjena** (tri mesta u
  `ReportSledljivostDokumenti`): `IIf` računa OBE grane, pa bi nevalidan
  datum pukao baš na dokumentu koji ugovor obećava da ostane vidljiv.
  Sada ugnježdeni `If`. Fixture: `PAL-SLED-B` (Datum „nevalidan");
  test 158 + sabotaža `sledljivost-dokumenti-nevalidan-datum-skriven`.
- **R5 — `zPop` tekst-traka (S15 zatvoren, u LJUSCI kao izričito
  odobren izuzetak).** Revizija je izmerila: nije `Font.Size` nego GDI
  rasterizacija — label PUNE visine (21pt) menja baseline fazu po redu,
  pa „8. red izgleda kao veći font". Isto pravilo kao mreža: pozadina/
  hit ostaje pun red, TEKST ide u unutrašnji label visine
  `TxtH(TS_BODY)` centriran kroz `CenterY`; hover se boji CENTRALNO
  (`PopHoverRed`, kao `mHoverRow` mreže), klik radi preko istog taga.
  Pečat `v6-ui-188`. Test 159 + sabotaža `ljuska-popup-tekst-pun-red`.
- **R6 — `BeginTableCache`/`EndTableCache`** oko sva tri read-modela u
  snimku (isti veliki listovi tri puta); kes se zatvara i na grešci.
- **R7 — auto-link je GLOBALAN i greška ≠ 0.** Toast sada kaže „SVI
  periodi, ne samo prikazani"; `AutoLinkOtkupOtpremnica_TX` dobija
  `Optional ByRef outGreska` (legacy pozivi netaknuti), pa ekran
  razlikuje rollback (poruka greške) od legitimne nule.

**Krug 9 — ispravka poslovne semantike fakture (finalni review nalaz):**
pravilo iz kruga 8 („SVE prijemnice moraju biti fakturisane") bilo je
pogrešno za AgriX domen — legitiman tok je prijem u SOPSTVENU hladnjaču
(`Fakturisano=Ne`), pa paleta/prerada/kasnija prodaja; faktura NIJE
obavezna karika sledljivosti. Zato je i ~938 „nepotpunih" na pravoj
svesci bilo uglavnom normalna roba u hladnjači. Pravi invariant:
**„Ne" = normalno; „Da" mora imati neprazan `FakturaID` ka postojećoj
aktivnoj fakturi** (ALL nad prijemnicama koje TVRDE fakturisanost).
Preimenovano po istini: klasa `FAKTURA-VEZA-NEISPRAVNA` (prikaz
„Neispravna veza fakture"), oznaka lanca `faktura neusaglasena`.
Testovi obrnuti: `PRJ-SLED-N` (Ne) sada dokazuje POTPUN lanac bez
problema; F-lanac ostaje negativ (Da + mrtav ID); NEP pretraga po
zbirnoj prelazi na F-vozilo. Usput `SledParceleMapa`:
`IIf(c > 0, src(i, c), …)` mina (IIf računa obe grane → `src(i, 0)`
puca) zamenjena `RequireColumnIndex` + direktan pristup — kolone su
ionako obavezne (`TraceByZbirna` ih tako tretira).

### 24.8 Verifikacija

- `RunAllTests` (deset novih testova 150–159, registrovani u sva tri
  registra, izvršavaju se PRE 124–126; 156 je NAMERNO samo čitanje —
  upis povezivanja bi pojeo vozilo `OTK-NAL-DJ` koje test 152 meri kao
  „nepovezan"), `RunBankaImportTestSuite` (Platni nalozi bit-identični —
  SLED blokovi su zatvoreni; ponovljeno posle svake izmene fixture-a),
  `vba_check` + `--self-test`, `sabotaza --proveri-sidra` +
  `--self-test`, `who_writes --check` — rezultati u PR-u.
- Dvosmerni dokaz: `python tools/dokaz.py sledljivost ljuska-popup` —
  31 sabotaža, svaka obara tačno svoj imenovani test i vraća se
  bit-identično.
- Diff ljuske: `modOtkupUI` = pečat (`OTKUI_BUILD` → `v6-ui-188`) + `zPop`
  tekst-traka iz kruga 8 (S15 — IZRIČITO odobren izuzetak od pečat-only
  pravila u spoljnoj reviziji: kvar je zajedničke ljuske), ništa
  više; kontekstni tabovi nisu ni trebali (liste su statične — S9 dopuna
  iz kruga 5 je idempotentna).
- **Ručna kapija operatera (traži se izričito):** `Alt+F11 → Debug →
  Compile VBAProject`, pa smoke nad pravim podacima (checklista u PR-u):
  izgled zone, tri liste, pretraga po broju fakture/zbirne na velikoj
  svesci (smer nazad), brzina prvog otvaranja na punoj svesci, detalj
  traka, obe štampe, ruta „Štampaj dokument" po vrsti karike, NEPOTPUNI
  nad pravim podacima (krug 9: prijemnice u sopstvenu hladnjaču sa
  `Fakturisano=Ne` NISU na listi — očekuje se drastičan pad broja
  „problema" na stvarne kvarove); iz drugog kruga još: „Poveži automatski" i „Poveži…" nad
  pravim nepovezanima (upis!) i šablon zbirne protiv iste štampe iz
  stare forme; iz trećeg kruga: polje „Dokument sledljivosti" — kucanje
  broja/datuma sužava listu (strelica je otvara), izbor
  zbirne/palete/prerade daje tačan šablon / paletni / preradni list;
  dugme bez polja i dalje radi za red sa jednoznačnom metom; iz četvrtog
  kruga: pretraga po broju zbirne na NEPOTPUNIMA nalazi njene
  nefakturisane prijemnice, a povezivanje ide kroz polje „Otpremnica za
  povezivanje" (izbor reda ga puni, „Poveži…" čita izbor — InputBox-a
  više nigde nema).

## 25. Sledljivost GP lanac + fakturisanje gotove robe (`v6-ui-189`)

Nastavak §24 po planu Faza 1 + Faza 2 (dogovor posle merge-a #246):
lanac se produžava POSLE prijemnice (paleta → prerada → GP faktura), a
ekran Fakturisanje dobija fakturisanje gotove robe — time završna
faktura zatvara i hladnjača-tok, kroz podatkovnu vezu.

### 25.1 Faza 1 — LANAC kolone 28–30 (čista projekcija)

- `ReportSledljivostLanac` 27 → 30 kolona: 28 „Palete", 29
  „Prerada / GP", 30 „Stanje". `SledGpMape` gradi pet mapa JEDNIM
  prolazom kroz `tblPaleta(Stavka)`/`tblPrerada(Stavka)` pre petlje
  (S5 pravilo — nijedan `LookupValue` po redu).
- Iste podatkovne veze kao mete (§24.2 E): paletna stavka vezuje BROJ
  zbirne, prerada join `PaletaID`, GP faktura **isključivo**
  `tblPrerada.FakturaID`. Heuristika „isti artikal + datum + kupac ≈
  ta faktura" je izmišljanje sledljivosti i ZABRANJENA (advisor uslov).
- Fail-closed: `nVl > 1` (dvosmislen broj) ne pripisuje ništa — kolone
  28–30 prazne; pripisivanje po broju bez vozača bi sabralo tuđe tokove.
- Prerada `Fakturisano=Da` bez validne aktivne fakture = ista klasa
  kvara kao prijemnica (krug 9): oznaka `faktura neusaglasena`; na
  NEPOTPUNIMA red `FAKTURA-VEZA-NEISPRAVNA` sa `DokTip=Prerada`
  (nova konstanta `SLED_DOK_PRERADA`; ruta štampe → `ExportPreradaPDF`,
  ista kao meta PRERADA). `Fakturisano=Ne`/prazno je legitiman tok.
- Stanje (kol. 30, `SLED_ST_*`): najdalja dostignuta karika po
  prioritetu GP faktura > prerada > sveža faktura > paleta > otvoren
  tok; piše se SAMO kad red nema oznaku. „Otvoren tok" nema fixture
  vozilo — ne tvrdi se testom (zapisano, svesno).
- SearchRefs (kol. 27) += brojevi paleta/prerada/GP faktura (princip
  R2 — prikaz „N pal."/„N pre." guta brojeve). Haystack LANAC += kol.
  30 (pretraga po stanju).
- Ekran: LANAC 11 → 14 kolona (ref pomeren 11 → 14 u `OsveziDetalj` i
  `StampajDokumentReda`); hederi `OTKUI_HDS_PALETE/PRERADAGP/STANJE`.
  Brojevi paleta/prerada su sada u haystacku, pa fixture brojevi tih
  dokumenata moraju da se završavaju na 0/1 i budu daleko od
  `GenerateBroj*` sekvence — substring pretraga inače guta brojeve
  faktura (`38/2026` sadrži `8/2026`; naučeno padom `T_Sled_KesPretragaIHint`).
  PDF dokument lanca (šablon iz kruga 6) NIJE širen GP karikama u ovom
  krugu — ekran ih nosi; širenje šablona po smoke potrebi.

### 25.2 Faza 2 — šema, writer, ekran Fakturisanje

- Šema (`EnsurePaletniListSchema`, kolone na KRAJ tabela):
  `tblFakturaStavke.PreradaID/BrojPrerade`,
  `tblPrerada.Fakturisano/FakturaID`. Postojeći pozicijski `AppendRow`
  upisi (header 21 / stavka 9) netaknuti; GP identitet ide PO IMENU
  (`RequireUpdateCell` posle `AppendRow`).
- `modFaktura.CreateFakturaGP_TX(kupacID, stavke)` — stavka =
  `Array(preradaID, cena)`; kapije 1731–1745 u BASE writeru pod TX
  (postoji/jedinstvena/nije stornirana/nije već fakturisana/izlaz > 0/
  cena > 0; pre-validacija SVIH stavki pre ijednog upisa). Cena je
  UNOS OPERATERA — GP nema evidentiranu cenu nigde (legacy nema GP
  prodaju; „izvedena" cena bi bila izmišljanje). Prerada se markira
  `Fakturisano=Da + FakturaID` — isti obrazac kao prijemnica, pa storno
  i sledljivost rade istim pravilima.
- `GetGPZaFakturisanjeForGrid` (1..9) — GP kolone čita meko
  (`GetColumnIndex`): sveska pre nadogradnje = niko nije fakturisan.
- Storno simetrija (`modStorno`): storno fakture →
  `ReleasePreradaFromFaktura` (reset markera; fail-closed na dupli ID);
  storno prerade → `MarkFakturaOrphaned` + `MarkFakturaStavkeOrphanedGP`
  (postojeći `MarkFakturaStavkeOrphaned` filtrira po `PrijemnicaID` pa
  bi za GP bio no-op). `GetColumnIndex` + guard svuda u storno putu —
  sveska pre nadogradnje ne sme da pukne.
- Ekran `modScrFakture` (`v6-ui-189`): četvrta lista `GOTOVA` (druga po
  redu, uz ZAFAKT), čipovi sve/čeka/fakturisano, kolone GP mreže
  (identitet `PreradaID` prio 4), polje `scrFkCenaGP` vidljivo samo na
  GOTOVA. Korpa: `KorpaTip()` (""/"PRJ"/"GP" po prvoj stavci) — mešanje
  odbijaju OBE strane (`OTKUI_ERR_FK_MESANJE`); GP stavka bez cene ne
  ulazi (`OTKUI_ERR_FK_CENA_GP`); `IzradiFakturu` rutira po tipu korpe
  (GP → `CreateFakturaGP_TX`, ostalo → `CreateFaktura_TX`). Test
  seam-ovi `Scr_FkKorpaTestDodajGP`/`Scr_FkUkloniStavkuGPTest` (gate
  `IsTestMode`).

### 25.3 Fixture i verifikacija

- `make_fixture`: **`ENSURE_COLS`** — generator dodaje kolone koje
  donor nema (sejanje ide PO IMENU pa bi red sa novom kolonom oborio
  build; fixture = sveska POSLE nadogradnje šeme). U potpisu.
- Tri nova lanca: G (prodato GP — `PRE-SLED-G` → aktivna `9/2026`),
  H (u hladnjači — sveža `PAL-SLED-H`), K (kontradikcija —
  `PRE-SLED-K` → mrtav `FAK-NEMA-GP`); potrošne prerade `PRE-GP-W1`
  (writer uspeh, mutira se poslednja), `PRE-GP-X` (stornirana),
  `PRE-GP-W0` (0 kg).
- Testovi 160–162 (sva tri registra; 160–161 read-only pre mutirajućih
  124–126, 162 MUTIRA i ide poslednji): `T_Sled_GpLanacIStanja` (ručni
  prolaz kroz tabele = Report; stanja; fail-closed D; problemi;
  smer nazad), `T_Fak_GpListaIKorpa` (read-model; korpa pravila u oba
  smera), `T_FakturaGP_WriterKapijeIStorno` (svih 6 kapija + rollback
  dokaz + stavka po imenu + storno oslobađa). `T_Fak_UgovorEkrana`
  ažuriran (3 → 4 liste, GOTOVA druga).
- 7 novih sabotaža (`sledljivost-gp-stanje-placebo`,
  `sledljivost-gp-kontradikcija-cuti`, `sledljivost-gp-refs-progutani`,
  `faktura-gp-dupla-prodaja`, `storno-gp-ne-oslobadja`,
  `fakture-gp-korpa-mesa-u-prj/-u-gp`) — katalog 360; dva stara sidra
  ažurirana (kolone/haystack pomereni), `fakture-sef-lista-uslovna`
  tvrdnja prati novi tekst.
- `who_writes` ažuran (modFaktura piše `tblPrerada` od Commit A).
- **Ručna kapija operatera:** Compile + smoke checklista u PR-u
  (lanac kolone i stanja na pravim podacima, GP lista, izrada i storno
  GP fakture, mešanje korpe).

### 25.4 Smoke GP-1: nalazi operatera (ispravke u istom PR-u)

- **Prekidači lista Fakturisanja imenuju ROBU, ne proces:** „Za
  fakturisanje" uz „Gotova roba" je bio nespojiv par — prvi je
  preimenovan u **„Sveža roba"** (`OTKUI_SEG_FK_ZAFAKT`; ključ i lista
  ostaju `ZAFAKT`).
- **Hederi GP kolona u operaterovom rečniku:** „PAL. SVEŽE ROBE" /
  „PAL. GOTOVOG PROIZV." umesto internih „PALETE" / „PRERADA / GP"
  (prerada = paleta gotovog proizvoda).
- **Kolone LANAC prate tok robe:** posle prijemnice idu dve paletne
  kolone, pa faktura i kupac; oznaka pretposlednja, **stanje
  poslednje** (zaključak reda). Grid raspored: Datum | Br. dok |
  Kooperant | Kg | Otpremnica | Zbirna | Prijem | Pal. sveže | Pal.
  gotovog | Faktura | Kupac | Oznaka | Stanje | [ref]. SRC kolone
  reporta (1–30) se NE menjaju — samo mrežni raspored (`UpisiRed`).
- **„Prvi put laguje" (lag prvog otvaranja):** GP sekcija lanca je
  pravila **dva `CreateObject("Scripting.Dictionary")` po redu
  otkupa** — na velikoj svesci desetine hiljada COM kreacija u jednom
  punjenju snimka (kasniji ulasci su keš hit, zato „na dalje je OK").
  Rečnici se sada prave jednom pre petlje i prazne `RemoveAll`-om po
  redu; `SledFakMapa` se u problemima gradi jednom (prijemnički i GP
  prolaz dele mapu). Za merenje na pravoj svesci isporučen
  **`Diag_SlPerf`** (Alt+F8 → Immediate): ms po read-modelu pod istim
  TableCache uslovima kao ekran — brojke presuđuju, ne nagađanje.

### 25.5 Spoljna revizija #248 (R1–R6): GP faktura kao first-class

Verdikt revizije: core i data-model dobri, ali GP faktura nije bila
first-class kroz Print/SEF/storno/deployment. Sve tvrdnje proverene u
kodu — svih pet tehničkih je stajalo. Potvrđene poslovne odluke
operatera: **1 prerada = ceo prodajni lot** (prodaje se u celosti;
zrno se ne spušta) i **GP ima isti poreski tretman kao sveža roba**.

- **R1 Print:** `PrintFaktura` GP grana (meke kolone) — GP stavka na
  papiru nosi **broj prerade + TipGotovogProizvoda** umesto praznih
  prijemničkih polja; `FillFakturaSablon` dobio `gp` parametar i piše
  hedere kolona 2/3 na SVAKOM renderu („Broj prerade"/„Proizvod" ↔
  „Broj prijemnice"/„Klasa") jer je šablon perzistentan.
- **R2 SEF:** `BuildSEFInvoiceDto` polimorfni izvor stavke — tačno
  jedan od `PrijemnicaID`/`PreradaID`, fail-closed na nijedan ili oba.
  GP linija: naziv = TipGotovogProizvoda + „po preradi N", UBL
  `SellersItemIdentification` = `PreradaID` (stabilni source ID),
  datum isporuke = **datum prerade**. `clsSEFLine` +preradaID/
  brojPrerade; porez nepromenjen (ista stopa za sve linije).
- **R3 Storno atomarnost:** `StornoFaktura_TX` snapshot **+tblPrerada**
  (release grana je piše — bez snapshota bi pukli rollback ostavio
  preradu slobodnu = dvostruka prodaja); `StornoPrerada_TX` snapshot
  **+tblFakture/tblFakturaStavke** (orphan grana). Rollback dokaz kroz
  test seam `StornoTestFailPosleRelease` (gejtovan `IsTestMode`) —
  namerna greška posle release-a, tvrdi se da rollback vraća SVE.
- **R4 Runtime šema:** četiri GP kolone dodate u `EnsureRuntimeSchema`
  (svaki start) — klijent koji dobije kod self-update-om dobija i
  kolone; ručni Admin korak više nije preduslov.
- **R5 Locale:** `CDbl(Val(nz(...)))` u GP read-modelu zamenjen
  `IsNumeric/CDbl` obrascem — `Val("50,5")` na srpskom locale-u čita
  50, pa bi grid/korpa pokazali jedno a writer upisao drugo. Fixture
  `PRE-GP-W1` sada ima decimalan izlaz (50.5).
- Fixture: `tblKupci` se sada SEJE (2 test kupca sa PIB-om — SEF DTO
  ih traži), pin `SELLER_NAME`/`SELLER_PIB`/`FAKTURA_PRINT_MODE=OFF`.
- Test 162 proširen: SEF DTO + UBL tvrdnje, print tvrdnje (GP hederi i
  stavka + povratak svežih hedera), rollback atomarnost; 4 nove
  sabotaže (katalog 364). Smoke ciklus revizije: GP faktura → PDF →
  SEF payload → storno → prerada dostupna → lanac „preradjeno".
- Odloženo (P1 revizije, svesno): oznaka „mešovit tok" kad potomci
  jedne zbirne imaju više završnih stanja — stanje je dokumentovano
  kao „najdalja karika".

### 25.6 Spoljna revizija #248, krug 3 (B1–B3 + 2×P1): finalni lanac

Verdikt kruga: tehnički GP tok zatvoren, ali vizuelni lanac nije
stizao do finalnog kupca i „prodato" se tvrdilo markerom umesto
podatkom. Potvrđeno ranije: prerada = finalni lot (kese/kutije),
posle nje nema proizvodnih karika.

- **B1 — GP faktura zauzima kolone Faktura i Kupac:** kolona 29 nosi
  SAMO preradu (heder „GOTOV PROIZVOD" — „PAL. GOTOVOG PROIZV." je
  sugerisao identitet koji ne postoji); završna GP faktura ide u
  kolonu 12, a **finalni kupac sa fakture** u kolonu 13 (do tada je
  uz prodatu robu stajalo odredište prijema — „Hladnjača"). Više
  završnih faktura/kupaca → „N fakt."/„N kup." + brojevi i nazivi u
  SearchRefs. Red evoluira: `103/2026 | 55/2026 | [prazno] |
  Hladnjača | preradjeno` → `103/2026 | 55/2026 | 91/2026 | Metro |
  prodato GP`.
- **B2 — kanonski GP invoice-link contract:** VALIDNO PRODATO =
  `Fakturisano=Da` **AND** `FakturaID` aktivan **AND** tačno JEDNA
  aktivna `FakturaStavka(PreradaID, FakturaID)` — `FakturaStavka.
  PreradaID` je time postao ono zbog čega je uveden: **dokaz** prodajne
  veze, ne dekoracija. Marker bez stavke / višak stavki / stale
  `FakturaID` uz `Fakturisano=Ne` → „faktura neusaglasena" (lanac +
  NEPOTPUNI). DOSTUPNO za novu fakturu = nije fakturisana AND veza
  prazna AND nema aktivne stavke — writer (`CreateFakturaGP` kapije
  1748/1749) i grid dele isti contract (`GPAktivneStavkePoPreradi`).
- **B3 — uslovni snapshot:** `StornoFaktura_TX` snapshotuje
  `tblPrerada` samo ako tabela postoji — stariji workbook bez Prerade
  ponovo može da stornira običnu svežu fakturu (isti meki ugovor kao
  `GetColumnIndex` kapije).
- **P1 — proizvod je writer invariant:** prazan `TipGotovogProizvoda`
  ne prolazi u fakturu (kapija 1750) — print bez imena robe i SEF
  fallback „Gotov proizvod" su bili fail-open.
- **P1 — dupli `PreradaID` guard u gridu:** `BrojacIdova`/`IdIliPrazno`
  (isti obrazac kao prijemnice/fakture) — korupcija prazni identitet
  umesto generičke greške tek pri izradi.
- Fixture: `FST-SLED-GP` stavka (dokaz za G lanac), `FAK-SLED-GP2`
  (aktivna bez stavke), vozila `PRE-GP-B2`/`PRE-GP-WL`/`PRE-GP-WT`/
  `PRE-GP-DUP×2`; +6 sabotaža (katalog 370).

### 25.7 Spoljna revizija #248, krug 4: totalCount — contract kompletan

Krug 3 je proveravao samo PAR (`preradaID|fakturaID`), što ne vidi
dve rupe: **istu preradu sa aktivnom stavkom na DRUGOJ fakturi**
(dvostruka prodaja koju bi lanac prikazao kao uredno „prodato GP") i
**stavku bez markera** (writer vidi prodajnu vezu i odbija novu
prodaju, a Sledljivost je tvrdila da veze nema). Grain potvrđen još
jednom: 1 prerada = cela se prodaje jednom kupcu na jednoj fakturi.

- `SledGpStavkeMapa` sada nosi i **`#T|preradaID` total ključ**;
  validno „prodato GP" = pair(P|F)=1 **AND total(P)=1**. Lanac i
  problemi dele mapu; problemi dobili grane „stavke na više faktura
  (dvostruka prodaja)" i „aktivna stavka bez markera na preradi";
  lanac obe klase vodi u „faktura neusaglasena".
- Writer/grid su total već pokrivali (`GPAktivneStavkePoPreradi` je
  po preradi) — dokazna strana je sada ista.
- P1: grid `Dostupna` traži i neprazan `TipGotovogProizvoda` (UI ne
  sme da kaže „može" pa writer „ne može"); P1: writer proverava da
  kupac postoji tačno jednom u `tblKupci` (kapija 1751 — GP nema
  prijemnicu za implicitnu proveru vlasništva).
- Vozila: `PRE-GP-2F` + stavke na `FAK-SLED-GP3/GP4` (par uredan,
  total=2), `PRE-GP-SB` + stavka bez markera. +4 sabotaže (katalog
  374; peta — „total ključ mrtav" — uklonjena jer joj je tvrdnja
  subsumirana postojećom sabotažom: checker `deli tvrdnju` odbio).

### 25.8 Krug 5: UTOVARNA LISTA — pravi prodajni grain

Poslovna presuda operatera: parcijalna prodaja MORA biti moguća (500
kg sa palete od 1.000), a datum prodaje NIJE datum prerade nego datum
utovara/slanja (obrazac utovarne liste sledi). Time je 1:1 model
(marker na preradi) postao netačan i **zamenjen je pre merge-a**.

- **Model:** `tblUtovar` (UtovarID | BrojUtovara | Godina |
  DatumUtovara | KupacID | Fakturisano | FakturaID | Napomena |
  Stornirano) + `tblUtovarStavke` (UtovarStavkaID | UtovarID |
  PreradaID | BrojPrerade | KolicinaKg | Stornirano);
  `tblFakturaStavke` +`UtovarID`. **Prerada = proizvodni lot** (koliko
  je proizvedeno); **utovar skida robu sa stanja** (koliko, kome,
  KADA); **faktura finansijski prati utovar** (v1: 1 utovar = 1 GP
  faktura, ista TX). `Prerada.Fakturisano/FakturaID` su UKLONJENI iz
  svega (ništa nije bilo deploy-ovano). Bez `tblGPPaleta` — operater
  nema poseban fizički ID (prerada mu je „paleta"); `GPPaletaID`
  denorm se može dodati kasnije bez loma.
- **Na stanju** = `NetoIzlazKg` − SUM aktivnih utovarenih kg
  (`modUtovar.UtovarenoPoPreradi` — JEDNA funkcija za grid, writer i
  storno kapiju). Writer `CreateUtovarSaFakturom_TX`: kapije stanja
  (kolicina ≤ na stanju), proizvoda, kupca, cene; utovar + faktura +
  FST stavke (PreradaID/BrojPrerade/UtovarID po imenu) u jednoj TX.
- **SEF:** `DeliveryDate` GP stavke = **`Utovar.DatumUtovara`** (v1 =
  datum izrade fakture; kad stigne poseban ekran/obrazac utovara,
  datum ostaje isti izvor istine).
- **Storno lanac:** faktura → oslobađa utovar (roba OSTAJE utovarena);
  utovar (samo nefakturisan) → vraća robu na stanje (header + stavke);
  prerada sa aktivnim utovarom se NE stornira (fail-closed, nema
  orphan kaskade). Snapshot: `StornoFaktura_TX` +`tblUtovar` uslovno.
- **Sledljivost:** stanja `prodato GP` (sve fakturisano) /
  **`delimicno prodato`** (deo) / `preradjeno` (ništa); kolone 12/13
  nose fakture i finalne kupce validno fakturisanih utovara („N
  fakt."/„N kup." za više — parcijalno je LEGALNO, ne kvar).
  Neusaglašenosti: utovar marker bez fakture/FST dokaza, stale veza,
  stavka na storniran utovar, FST siroče, utovareno > proizvedeno —
  „faktura neusaglasena" + NEPOTPUNI (utovar klase nose DokTip
  `Utovar`; štampa odbija s razlogom dok obrazac ne stigne).
- **Ekran:** GP kolona kg = „NA STANJU (kg)"; polje „Količina za
  utovar (prazno = sve)" — ceo-lot klik ostaje default, parcijala je
  upis broja.
- Fixture: `ENSURE_TABLES` mehanizam (generator pravi tabele koje
  donor nema); P lanac („delimično prodato": 50 od 120 kg); contract
  vozila preseljena na utovar nivo (B2/WL/WM/OV/SB). Sabotaže: 15
  mrtvih (marker model) zamenjeno sa 14 utovar sabotaža (katalog 373).
- **Proces (novi režim):** subset dokaz novih sabotaža PRE punog
  kataloga našao je oba problema za 12 min (defense-in-depth: stanje
  kapija brani i W0; storno header već vraća stanje pa je stavkama
  dodata sopstvena tvrdnja) — pun prolaz ide tek jednom, na kraju.

### 25.9 Krug 5b: UI vidljivost utovara + štampani obrazac

Operater: model bez ekrana i papira je bezvredan — utovar mora da se
VIDI i da se PREDA sa robom. Isporučeno u istom PR-u:

- **Peta lista „Utovari"** na ekranu Fakturisanje (redosled = tok
  robe: Sveža roba | Gotova roba | **Utovari** | Fakture | SEF):
  broj, datum, kupac, roba (broj prerade / „N pre."), kg, broj
  fakture; čipovi sve/čeka(nefakturisan)/fakturisan; radnje
  **„Utovarna lista"** (štampa) i **„Storniraj"** (potvrda pa
  `StornoUtovar_TX` — writer kapija odbija fakturisan utovar).
- **Štampani obrazac** (`EnsureUtovarSablon`/`FillUtovarSablon`, house
  stil kao faktura; perzistentan list `UtovarSablon`): zaglavlje
  prodavca + „UTOVARNA LISTA" + broj/datum utovara/kupac; stavke
  `Rb | Broj prerade | Proizvod | Pakovanje (kutije/kese) | Kolicina
  (kg)`; UKUPNO; potpisi „Robu predao/preuzeo". Režim
  `UTOVAR_PRINT_MODE` (default PRINT; OFF u fixture-u).
- NEPOTPUNI ruta za DokTip `Utovar` sada ŠTAMPA listu (umesto
  odbijanja); `modUtovar.PrintUtovar` je zajednička ruta.
- Nova mina u katalogu znanja: `cFid`/`cFId` u ISTOJ proceduri =
  „Duplicate declaration" (case-insensitive) — uhvaćeno compile-only
  runom, skener dupli-Dim proširen na modUtovar/modPrint.

### 25.10 Smoke 5c: režim štampe, migracija starih GP faktura, tekstovi

- **Režim štampe utovarne liste ide kroz Podešavanja** (kartica
  Štampa, ključ `UTOVAR_PRINT_MODE`, PDF;PRINT;PREVIEW;OFF) — kao svi
  dokumenti; **default je PDF, ne PRINT**: bez podešenog ključa
  dokument ne sme tiho da ode na štampač.
- **Migracija:** GP fakture napravljene verzijama PRE utovarne liste
  imaju stavke bez `UtovarID` pa ih lista Utovari ne vidi (operaterov
  nalaz). `Alt+F8 → BackfillUtovariIzGPFaktura` (jednokratno,
  idempotentno) pravi utovar IZ postojećih podataka fakture (datum,
  kupac, stavke — ništa se ne izmišlja) i vezuje stavke. NAMERNO nije
  u `EnsureRuntimeSchema`: automatski bi „legalizovao" svaku buduću
  siroče-stavku umesto da je NEPOTPUNI prijavi.
- **Tekstovi:** naslovi mreže se seku pored prekidača — skraćeni na
  „Gotova roba" / „Utovarne liste" (kontekst nosi segment); label
  „Kol. za utovar (kg)" skraćen, a pravilo „prazno = celo stanje"
  seli u hint liniju GOTOVA liste (`OTKUI_LBL_FK_HINT_GP`).

### 25.11 Krug 5d: profesionalni obrazac utovarne liste + podaci prevoza

Operaterov zahtev: „uradi sve iz preporuke, ne želim minimalno nego
profesionalno" — obrazac utovarne liste dignut na pun logistički
dokument, sa unosom podataka prevoza na listi Utovari.

- **Šema `tblUtovar` +8 kolona:** `Prevoznik`, `Vozac`, `Registracija`,
  `Plomba`, `TemperaturniRezim`, `MestoIstovara`, `VremeUtovara`,
  `BrojNarudzbenice` (PO broj kupca). `CreateUtovarSaFakturom_TX`
  odmah upisuje vreme utovara (`hh:mm`); ostalo unosi operater.
- **Unos prevoza na listi Utovari:** 8 polja (prevoznik, vozač,
  registracija, plomba, temp. režim, mesto istovara, PO broj,
  napomena) vidljivih samo na toj listi + radnja „Sačuvaj prevoz" po
  redu. Writer `UpdateUtovarPrevoz_TX` (transakcioni, kapije: utovar
  postoji tačno jednom, nije storniran): **prazno polje ne dira
  postojeću vrednost, „-" briše** — dopuna posle delimičnog unosa ne
  gazi ranije uneto. Mreža Utovari dobila kolone PREVOZNIK i REG.
- **Obrazac (LAYOUT_VER 2, 9 kolona):** identitet dokumenta (broj,
  datum, vreme, faktura) + primalac (kupac, mesto istovara, PO broj);
  blok PREVOZ (prevoznik/vozač/registracija/plomba/temp. režim);
  stavke Rb | Lot | Proizvod | Proizvedeno | **Rok trajanja** |
  Pakovanje | **Pal.** („1" cela / „deo") | Neto (kg) | **Bruto (kg)**;
  zbirni red „UKUPNO · paleta: N [+ M deo]" + neto + bruto; napomena;
  tri potpisa (Robu predao — pečat / Vozač / Robu preuzeo — pečat).
- **Rok trajanja je IZVEDEN podatak:** datum prerade + N meseci iz
  Podešavanja (`GP_ROK_TRAJANJA_MESECI`, default 24 za smrznuto) — uz
  sanity opseg 1–600 celih meseci, jer datumski formatirana config
  ćelija kroz srpski locale (`CStr` → „23.1.1900." → `CLng` čita tačke
  kao hiljade = 2311900) obara `DateAdd` preko godine 9999 (greška 5).
  Posebna kolona roka po preradi ostaje budući korak.
- **Bruto po stavci:** cela paleta = bruto prerade; parcijala =
  srazmeran udeo neto količine (aritmetika nad stvarnim merenjima);
  bez bruto podatka = neto (ne izmišlja se).
- **Fixture generator:** config vrednosti se upisuju kao TEKST
  (`NumberFormat "@"`) — nov red tabele nasleđuje format reda iznad,
  pa je datumski format pretvarao pin „24" u datum 1900-01-23;
  `ENSURE_TABLES` sada i postojećoj tabeli dopunjava kolone koje fale
  (donor iz prethodnog kruga nema prevoz kolone).
- **Verifikacija:** test 161 tvrdi prevoz kolone u mreži; test 162
  pun obrazac (vreme/faktura/mesto/PO/prevoz/plomba posle dopune,
  datum proizvodnje, rok +24m, „deo", neto 20, bruto ≥ neto) + 
  semantiku prazno/„-" u writeru. Dve nove sabotaže:
  `utovar-gp-prevoz-prazno-brise` (prazno pregazi vrednost) i
  `utovar-gp-rok-placebo` (rok = datum proizvodnje) — obaraju tačno
  svoju tvrdnju. Katalog: 378 sabotaža.

### 25.12 Revizija #6: integrity/lifecycle krug (3 blockera + P1)

Završni krug pre merge-a — bez feature širenja, samo zatvaranje
lifecycle rupa. Verifikacioni ciklus (RunAll + dokaz + Banka) ide
ODLOŽENO, po operaterovom „tek kad budem zadovoljan modulom".

- **BLOCKER 1 — faktura iz postojećeg utovara:** posle storna GP
  fakture utovar ostaje aktivan a roba je fizički otišla — novi utovar
  bi tvrdio da je izašla dva puta. Novi writer
  `CreateFakturaIzUtovara_TX(utovarID)`: koristi postojeći
  nefakturisani utovar, NE dira stanje, čita njegove aktivne stavke,
  pravi samo novu fakturu + stavke, ponovo markira utovar, ostavlja
  originalni `DatumUtovara`. Cene iz poslednje (stornirane) fakture
  tog utovara — postoje uvek jer se utovar rađa uz fakturu; bez cene
  za neku preradu writer blokira s uputstvom. Radnja **„Fakturiši"**
  na listi Utovari (već fakturisan red — jasna poruka bez writera).
- **BLOCKER 2 — SEF kapija GP linije:** GP stavka bez `UtovarID` više
  NIKAD ne ide na SEF (ranije fail-open na datum fakture). Sve GP
  linije fakture moraju nositi ISTI utovar (1 faktura = 1 utovar), a
  utovar mora: postojati tačno jednom, biti aktivan, `Fakturisano=Da`
  i tvrditi baš tu fakturu (`ValidateGpUtovarZaSEF`) — inače BLOCK.
- **BLOCKER 3 — sledljivost dokaz po stavci:** dokaz fakturisanosti
  više nije „postoji bar jedna FST na utovaru" (ANY-vs-ALL rupa kod
  više prerada na istom utovaru) nego mapa `UtovarID|PreradaID →
  zbir kg` — stavka je fakturisana samo ako njen par postoji i
  količina se poklapa; sve drugo je „faktura neusaglasena".
- **Pakovanja parcijalnog utovara (t.4):** `tblUtovarStavke` +
  `BrojKutija`/`BrojKesa` — broj pakovanja koji je STVARNO ušao u
  kamion. Writer puni samo dokazivo: cela paleta = puni brojevi;
  parcijala = samo ako prerada ima jednu vrstu pakovanja i količina
  je celobrojan umnožak kg/pakovanju (500/10 = 50); inače prazno.
  Obrazac čita iz stavke (fallback na punu preradu samo za celu
  paletu) — transportni dokument radije prazan nego pogrešan.
- **P1:** datum i vreme utovara editabilni na listi Utovari (kroz
  „Sačuvaj prevoz"; prazno = ne diraj, datum se ne može obrisati jer
  je SEF datum isporuke); `tblUtovar`/`tblUtovarStavke`/`tblPrevoznici`
  u `AuditableTables()` (CreatedBy/ModifiedBy kao ostali dokumenti);
  migracija starih GP faktura numeriše utovar po godini DATUMA FAKTURE
  (`GenerateBrojUtovara(godina)`) — bez „3/2026 sa datumom 2025".

### 25.13 Revizija #7: završni integrity krug + pun recert

- **B1 — datum utovara zaključan posle SEF slanja:** `DeliveryDate` je
  poreski podatak. `UpdateUtovarPrevoz_TX` menja datum/vreme samo dok
  faktura utovara NIJE dokazano otišla spolja (`LOCAL_FINALIZED` /
  `SEF_READY` / `SEF_TECH_FAILED` + prazno); u svakom drugom SEF
  stanju promena je blokirana s porukom. Transportni tekst (prevoznik,
  plomba...) ostaje editabilan. Datum unos toleriše srpski format sa
  završnom tačkom („2.6.2026." — `IsDate` baš nju ne prima).
- **B2 — „nefakturisan" utovar sa aktivnom FST je kontradikcija:**
  `CreateFakturaIzUtovara` u istom prolazu kroz FST (cene) broji i
  AKTIVNE stavke — ako postoje, re-fakturisanje je blokirano (dupla
  prodaja); sledljivost isti slučaj markira „faktura neusaglasena"
  (header grana: bez markera + bez FakturaID + aktivna FST → lose).
- **B3 — SEF količinski dokaz 1:1:** pored postojanja utovara,
  `ValidateGpUtovarZaSEF` sada poredi zbir kg aktivnih GP linija po
  preradi sa zbirom kg aktivnih utovarnih stavki — u OBA smera
  (faktura od 400 kg nad utovarom od 500 kg = BLOCK; utovarna stavka
  nepredstavljena u fakturi = BLOCK).
- **P1:** radnja preimenovana u **„Ponovi fakturu"** — eksplicitno
  koristi cene prethodne stornirane fakture (za storno iz tehničkog
  razloga); korekcija cene je poseban budući rad (potvrda to i kaže).
- **Obrazac v5 — merge mina:** tri pojedinačna row-merge-a (F7:I7,
  F8:I8, F9:I9) Excel je slepio u JEDNU oblast F7:I9, pa je upis u
  F8 (`UtMestoIst`) tiho nestajao. Header/prevoz blokovi su sada BEZ
  merge-ova — tekst se prirodno preliva preko praznih susednih ćelija
  (v2 se sekao jer su susedne bile pune). POUKA: uzastopni row-merge
  isti raspon kolona = jedna velika oblast, imenovana ćelija unutar
  nje gubi upis bez greške.
- **Testovi (162):** ponovna faktura iz istog utovara (marker, ista
  količina na stanju, originalan datum), odbijanje već fakturisanog,
  rogue aktivna FST blok, SEF količinski mismatch sa ISTIM totalom
  (25 kg × 122 = 30.5 kg × 100 — vozilo kome je količinska kapija
  jedina brana), datum lock (SENT → odbij; transport tekst prolazi;
  LOCAL_FINALIZED → prolazi + upis). **Sabotaže 381** (+3: refakt-
  aktivna-fst, sef-kolicina-placebo, datum-lock-placebo).
- **Recert na finalnom SHA:** vba_check čisto; RunAllTests 162/0;
  Banka ZELENO; pun dokaz svih 8 prefiksa; compile ručna kapija.

### 25.14 Revizija #8: dokaz vezan za fakturu + korekcija cene

- **Sledljivost — troделни ključ dokaza:** FST dokaz je sada
  `UtovarID|FakturaID|PreradaID → kg` — aktivna stavka TUĐE fakture
  na utovaru više ne može da „dokaže" prodaju header fakture koja tu
  robu ne nosi. Uz to: aktivne FST dva različita FakturaID na istom
  utovaru = neusaglašeno (1 utovar = 1 faktura); aktivna FST za
  preradu koje na utovaru NEMA = neusaglašeno (višak). SEF mapper i
  sledljivost su sada jednako fail-closed.
- **Korekcija cene bez falsifikovanja utovara:** pogrešna cena više
  ne traži storno fizičkog utovara (roba JE otišla — taj događaj se
  ne poništava). „Ponovi fakturu" sada pita za novu cenu (prazno =
  cene prethodne fakture, za tehnički storno); v1 ograničenje: jedna
  cena za ceo utovar, pa se utovar sa više prerada odbija umesto da
  se cene tiho izravnaju. Poruka koja je savetovala storno utovara
  UKLONJENA.
- Testovi: korekcija cene (30.5 kg × 150 = novi iznos, datum utovara
  netaknut). Recert: 162/0; ciljani dokaz sledljivost-gp/utovar-gp/
  sef-gp.

### 25.14 Revizija #9: model B — utovar može postojati PRE fakture

Poslovna odluka operatera (uz revizorsku preporuku): utovarna lista je
događaj magacina, faktura događaj računovodstva — životni ciklusi im
ne moraju početi istog trenutka.

- **Writer razdvojen na dva core-a:** `CreateUtovarCore` (kapije +
  tblUtovar + stavke: količina, pakovanja, **dogovorena cena `CenaKg`**)
  i postojeći `CreateFakturaIzUtovara`. `CreateUtovarSaFakturom_TX`
  (brzi put) je sada tanak kompozit ta dva; novi `CreateUtovar_TX`
  pravi SAMO utovar. Lanac cene u fakturisanju: uneta korekcija >
  poslednja faktura utovara > **dogovorena cena sa stavke** (B-utovar
  bez FST istorije) > blok.
- **UI:** dva dugmeta na listi Gotova roba — postojeće „Izradi
  fakturu" (utovar+faktura odmah) i novo **„Napravi utovar"** (samo
  utovar; roba se skida sa stanja, lista se štampa, čip „čeka").
  Radnja na listi Utovari vraćena na generičko **„Fakturiši"** — sada
  pokriva prvo fakturisanje, ponavljanje posle storna i korekciju cene.
- **Rok trajanja po vrsti proizvoda** (poslovno potvrđeno: rokovi se
  razlikuju): `tblVrstaGotovihProizvoda` + kolona `RokMeseci` (unos u
  Matičnim podacima; sanity 1–600 celih); obrazac je koristi po
  stavci, prazno = globalni `GP_ROK_TRAJANJA_MESECI`.
- Šema: `tblUtovarStavke` + `CenaKg`; `EnsureRuntimeSchema` self-heal
  za obe kolone. Test 162: B tok (utovar bez markera skida stanje,
  cena na stavci, kasnije fakturisanje nosi tu cenu, storno lanac
  vraća stanje). SEF kapije netaknute — B-utovar bez fakture ne
  postoji na SEF putu dok se ne fakturiše.
- Revizorov audit P1 bio je zastareo nalaz: utovar tabele su u
  `AuditableTables()` od revizije #6.

### 25.15 Revizija #10: lanac razume model B + storno integritet

- **B1 — nova stanja lanca:** „utovareno / ceka fakturu" (sve fizički
  otišlo, faktura ne postoji) i „delimicno utovareno" (deo otišao,
  ništa fakturisano) — utovar je samostalan događaj i lanac ga
  prijavljuje i PRE fakture; finalni kupac (kolona 13) dolazi iz
  aktivnog utovara, faktura ga kasnije samo potvrđuje. Fixture vozilo:
  U lanac (120 proizvedeno / 50 utovareno / 0 fakturisano,
  `UT-SLED-U` bez markera).
- **B2 — storno utovara ne veruje samo header markeru:** kanonski
  helper `modUtovar.AktivnihFstZaUtovar` (dele ga
  `CreateFakturaIzUtovara` i `StornoUtovar`) — aktivna faktura-stavka
  koja tvrdi utovar blokira storno (dupla zaliha) isto kao što blokira
  re-fakturisanje (dupla prodaja).
- **B3 — release proverava vlasnika:**
  `ReleaseUtovarFromFaktura(utovarID, fakturaID)` resetuje marker SAMO
  ako utovar tvrdi baš fakturu koja se stornira; korumpirana stavka
  tuđe fakture se loguje i NE dira tuđ marker (lanac je prijavljuje
  kao neusaglašenost).
- **P1 audit self-heal:** `EnsureRuntimeSchema` dopunjava audit kolone
  (CreatedAt/By, ModifiedAt/By) na tri nove tabele — klijent na
  self-update putu dobija ko/kada trag bez ručnog `EnsureAuditColumns`.
- **Test 163** (`T_UtovarB_SledIStornoKapije`, poseban broj po
  revizorskoj preporuci): U lanac stanje+kupac, rogue FST blokira
  storno i re-fakturu, storno žrtvene fakture sa korumpiranom stavkom
  ne oslobađa tuđ utovar, pun cleanup. RunAllTests **163/0**.
- POUKA: oslanjanje na `Dim i` niže u proceduri obara compile
  („Variable not defined" modalni dijalog = suite VISI headless);
  blok uvek deklariše svoju petlja-promenljivu.

### 25.16 Revizija #11: kupac invariant + čišćenja — završni kod krug

- **BLOCKER — kupac utovara = kupac fakture:** kanonski GP contract
  proširen poslednjom karikom. SEF (`ValidateGpUtovarZaSEF`) blokira
  slanje kad `Utovar.KupacID ≠ Faktura.KupacID`; sledljivost isti
  mismatch markira „faktura neusaglasena" (ne „2 kup.") — faktura
  kupcu B za robu koja je fizički otišla kupcu A nikad nije validna
  prodaja. Test 163 dokazuje SEF blok (korupcija kupca → blok →
  restore → prolaz).
- **P1 — parcijalni bruto se ne štampa kao procena:** cela paleta
  nosi izmereni bruto prerade; parcijala nema svoj izmeren bruto pa
  ostaje PRAZNA (isti princip kao pakovanja); zbir bruto samo od
  izmerenih stavki. Stvarni `UtovarStavka.BrutoKg` unos je budući rad.
- **P1 — release jednom po utovaru:** storno fakture sa više stavki
  istog utovara više ne pravi lažne „tvrdi drugu fakturu" warninge
  (dedup po UtovarID pre `ReleaseUtovarFromFaktura`).
- **Poslovna odluka (operater):** cena OSTAJE obavezna pri izradi
  utovara — dogovara se pre utovara; „Fakturiši" je čita sa stavke.
- RunAllTests **163/0**; pun recert (dokaz + Banka + compile + smoke)
  na finalnom SHA pred merge — bez novih feature-a posle ovog kruga.

### 25.17 Revizija #12: migracija uklonjena — završni rez pre merge-a

- **`BackfillUtovariIzGPFaktura` OBRISAN** (revizorski blocker): rutina
  je iz datuma FAKTURE izvodila datum FIZIČKOG utovara — a upravo je
  ovaj PR utvrdio da su to dva različita događaja i da je datum
  utovara poreska činjenica (SEF DeliveryDate). Uz to nije bila
  transakciona (pad na pola fakture = parcijalno stanje). Stari GP
  model nikad nije deploy-ovan, pa je migracija podržavala samo
  neobjavljenu međufazu ovog istog PR-a. Ako backfill ikad stvarno
  zatreba: zasebna migracija u kojoj operater EKSPLICITNO unosi
  stvarni datum utovara, uz transakciju po celoj fakturi (zabeleženo
  u modSetup komentaru).
- Uklonjeni i: migracijski blok testa 162 (SB vozilo ostaje u
  fixture-u — sledljivost siroče prijavljuje kao neusaglašenost) i
  sabotaža `utovar-gp-migracija-bez-veze`. Katalog: 380 sabotaža.
- Zabeleženo za kasnije (revizorske ne-blocker napomene): snapshot
  `OpisProizvoda` na FakturaStavci (reprint ne sme da zavisi od
  kasnije promene naziva na preradi), snapshot roka na utovarnoj
  stavci, i prikaz utovara u detail stripu Sledljivosti (mreža je na
  MAX_COLS=14).
- PR opis #248 prepisan: dokumentuje model B / utovar grain umesto
  obrisanog `Prerada.Fakturisano` modela.

### 25.18 Smoke #12: pakovanja prate stanje (kapacitet izveden iz lota)

Operaterov nalaz posle finalnog smoke-a (koji je ceo revizorov tok
prošao): roba se smanji, a kutije/kese ne — a pakovanje JE definisano
(uniformnih 10 kg), pa je računica prosta.

- **`modUtovar.PakovanjaZaKg(kg, lotNeto, lotBroj, samoTacno)`** —
  jedna formula za obe potrebe; kapacitet pakovanja se IZVODI iz
  samog lota (`neto ÷ broj` = kg po pakovanju), bez nove šeme:
  - **grid Gotova roba** (`samoTacno=False`): broj CELIH pakovanja na
    stanju (`Fix`) — 733 kg uz 10 kg/kutiji = 73; hederi vraćeni na
    „KUTIJE"/„KESE" (LOT sufiks više ne važi jer kolone prate stanje);
  - **utovarna lista** (`samoTacno=True`): broj SAMO kad je količina
    celobrojan umnožak kapaciteta — dokument i dalje ne nosi procene.
- **Parcijala sada dobija OBE vrste pakovanja** (500 kg = 50 kutija i
  50 kesa) — ranije je uslov „jedina vrsta pakovanja" ostavljao prazno
  kad lot ima i kutije i kese.
- **Read-time fallback u štampi:** stare parcijalne stavke (napravljene
  pre ovog fixa, bez upisanih pakovanja) na reprint dobijaju istu
  strogu računicu iz kapaciteta lota (`PakTekst` helper).

### 25.19 Kapacitet pakovanja kao poslovna postavka

Operaterov contract: kapaciteti se rešavaju u Podešavanjima, a
kapacitet = **neto kg robe** po pakovanju (težina ambalaže ne ulazi —
ona živi samo u brutu palete).

- Nova podešavanja (kartica Štampa, uz rok trajanja):
  `GP_KG_PO_KUTIJI` „Kapacitet kutije (kg neto)" i `GP_KG_PO_KESI`
  „Kapacitet kese (kg neto)".
- `KapacitetPakovanja` (modUtovar): postavka iz Podešavanja (sanity
  0.1–1000 kg, ista odbrana od datumski formatirane ćelije kao rok) →
  fallback izvedeno iz lota (neto ÷ broj) → 0. `PakovanjaZaKg` sada
  prima ključ postavke; svi pozivaoci (grid Gotova roba, upis stavke,
  reprint fallback) prosleđuju kutija/kesa ključ.
- Bruto logika netaknuta (izmereni bruto prerade za celu paletu,
  parcijala prazna).

### 25.20 Revizija #13: završni correctness commit

- **B1 — SEF lock po durable dokazu:** `SEF_TECH_FAILED` može doći i
  POSLE `SEF_SENT` (Mistake), pa workflow stanje nije dovoljan dokaz
  da dokument nije stigao spolja. Datum/vreme utovara je sada
  zaključan i kad faktura ima `SEFDocumentId` (isti princip kao
  `CanSendSEFInvoice`), bez obzira na stanje.
- **B2 — rok trajanja je SNAPSHOT na lotu:** nova kolona
  `tblPrerada.DatumIsteka` nastaje pri preradi po TADAŠNJEM pravilu
  (`RokIstekaZaTip`: vrsta → RokMeseci → globalni CFG → 24); štampa
  čita snapshot, a trenutna pravila koristi SAMO kao fallback za
  stare lotove bez snapshota. Kasnija promena podešavanja više ne
  menja rok na reprintu; otvara FEFO/isticanje analitiku. Globalni
  rok preseljen u grupu „Otkup / dokumenta" (nije štampa opcija).
- **B3 — UKUPNO BRUTO samo kad je kompletan:** ako ijedna stavka nema
  izmeren bruto, zbir ostaje prazan (`svaBruto`) — 930 kg pored
  1.400 kg neto više ne može da izgleda kao bruto celog kamiona.
- **P1 writer fail-closed:** `RequireStavkeKonzistentne` — oštećena
  aktivna utovarna stavka (nenumerička/nepozitivna količina, bez
  PreradaID, header nepostojeći/dupli/storniran) BLOKIRA dalju
  prodaju te prerade: tiho preskakanje u `UtovarenoPoPreradi` bi dalo
  previsoko stanje (korupcija → moguća dupla prodaja).
- **Kapacitet pakovanja — operaterov presek:** kapaciteti se NE
  dupliraju u Podešavanjima (polja iz prošlog commita uklonjena) —
  već postoje kao šifarnici `tblKutije`/`tblKese` (tip → NETO kg,
  Matični podaci), a prerada nosi tip: lanac je šifarnik po tipu →
  lot fallback (neto/broj). Testovi računice u T163 (lot fallback,
  strogo/floor, prioritet šifarnika).
- **Odluka (audit transporta):** plomba/registracija/prevoz ostaju
  editabilni i posle SEF-a (nisu poreski podatak) uz ModifiedBy trag;
  append-only istorija korekcija = budući rad za cold-chain dokaz.
## 26. Matični podaci — plan prelaska (M0–M6 urađeni)

> Ovo je jedina sekcija kataloga koja je napisana **pre** posla koji opisuje.
> Piše se pre koda namerno: prelazak Matičnih podataka nije „još jedan ekran po
> obrascu iz §7–§9, §22, §23" nego donosi **nov pojam u navigaciju** (sekcija
> ljuske) i dira sidebar, koji do sada nijedan ekran nije dirao. Odluka o tome
> se donosi jednom i pre pisanja, ne u toku.
>
> Stanje na dan `v6-ui-186`.

### 26.0 Zatečeno stanje — šta dugme danas sadrži

Dva puta, po jedan za svaku ljusku:

| Ljuska | Dugme | Šta se otvara |
| legacy `frmOtkupAPP` | `btnMaticni` (zlatno, gore desno) | `frmMaticniPodaci` — popup meni; klik na sekciju → `frmStammdaten` kao `mActiveContent` |
| novi `frmOtkupUI` | `btnMatic` (zlatno, gore desno) | **ništa** — `modOtkupUI.bas:3847` javlja `OTKUI_TODO_NEVEZANO` |

Meni je data-driven (`modMaticniLookups.MaticniSekcijeGrupisano`): **16 sekcija
u 4 grupe**, svaka sekcija = jedan `Tag` za `frmStammdaten`.

| Grupa | Sekcije (`Tag`) | Tabela | Oblik |
|---|---|---|---|
| Šifarnici | Kooperanti, Stanice, Kupci, Vozaci, Parcele | `tblKooperanti`, `tblStanice`, `tblKupci`, `tblVozaci`, `tblParcele` | lista + polja |
| Proizvodi i cene | Artikli, Kulture, Cenovnik, VrstaGP | `tblArtikli`, `tblKulture`, `tblCenovnik`, `tblVrstaGP` | lista + polja |
| Ambalaža i pakovanje | TipAmbalaze, TipPalete, Kutije, Kese | `tblTipAmbalaze`, `tblTipPalete`, `tblKutije`, `tblKese` | lista + polja (2 polja) |
| Sistem | Podešavanja, Admin, Korisnici | `tblSEFConfig`, — , `tblKorisnici` | **panel**, **panel**, lista + matrica prava |

`frmStammdaten` (3.674 linije) je jedna univerzalna forma: naslov + podnaslov,
`lstData` (do 10 kolona), `lblField1..10` + `txtField1..10` + `cmbField1..6`,
dugmad **Dodaj** / **Izmeni** / **Deaktiviraj-Aktiviraj** (runtime, `clsStmBtn`)
/ **Povratak**. Po sekciji se granaju: `Setup*` (13 procedura), `LoadList`
(351 linija), `lstData_Click`, `btnDodaj_Click` (**544 linije**),
`btnIzmeni_Click` (**463 linije**), `ClearFields`, `ResetFieldVisibility`.

Tri sekcije nisu obična CRUD:

- **Cenovnik** je append-only — `btnIzmeni` se sakriva, izmena je nov važeći red.
- **Parcele** nose GEO blok: GeoSrbija, paste koordinata, Google Maps, polygon
  editor, „obriši geo" sa potvrdom u dva koraka (`btnGeo*`, ~600 linija).
- **Korisnici** nose **dinamičku matricu prava po oblasti** (`BuildKorisniciOblasti`,
  combo po oblasti iz `modAuth.OblastiList`) i upis u transakciji, po imenu kolone.

Dve sekcije uopšte nisu liste: **Podešavanja** (`modPodesavanja.BuildConfigEditor`,
847 linija — grupisani editor `tblSEFConfig` sa sklapanjem grupa) i **Admin**
(`modAdmin.BuildAdminPanel`, 363 linije — deset dugmadi koja zovu postojeće
ulazne tačke: self-update, setup, health check, Google auth, release, VBA
import/export, migracija).

### 26.1 Mera koja je oblikovala rešenje

Sidebar **nema skrol**. `BuildNav` postavlja stavke apsolutno, `LayoutOtkup`
mu daje visinu `hTot - HEADER_H - STATUS_H`, a profil i podnožje su prikovani
za donjih 62 pt. Budžet na `MIN_H` (620):

```
620 - HEADER_H(42) - STATUS_H(24) = 554
554 - profil/podnozje(62)          = 492  slobodno
zauzeto danas: 12 + 3 grupe*(20+12) + 11 stavki*27 = 405   -> ostalo ~87
```

Ostalo je **~87 pt, to jest tri stavke — i to samo ako nema nove grupe**.
Matični podaci donose pet stavki i dve grupe (`5*27 + 2*32 = 199`), ukupno
**604 > 492**. Stavke bi na najmanjem prozoru nestale ispod profila, tiho —
tačno klasa kvara koju §10 i §13 već plaćaju. Granica se probija i pri
najštedljivijem rezu: već **četiri** stavke sa jednom grupom (`140`) prelaze
slobodnih 87.

**Zato Matični podaci ne mogu prosto da se dopišu u sidebar.** Sve što sledi
je posledica ove mere, a ne ukusa.

### 26.2 Odluka A — Matični podaci su SEKCIJA ljuske, ne još stavki u sidebaru

Zlatno dugme već nosi tu nameru. `modOtkupUI.BuildHeader` kaže doslovno:
*„Sve tri su ghostdark … a „Maticni podaci" ostaje jedino zlatno (primarno)
dugme."* Ono nije alatka nego **odredište** — isto što je u legacy ljusci
otvaralo ceo popup meni.

Predlog: `btnMatic` **prebacuje sekciju ljuske**. Sidebar tada pokazuje grupe
matičnih podataka, dugme menja natpis u „Nazad na rad" i vraća prethodni ekran.
Dva skupa stavki nikad ne stoje zajedno, pa budžet iz §26.1 nije ni blizu
granice (matična sekcija: `12 + 2*32 + 5*27 = 211`, radna ostaje `405`).

Mehanika, najmanja moguća:

1. `modUiScreens.ScrRows` dobija **sedmo polje** `SCR_SEKCIJA` (`RAD` | `MATICNI`);
   `ScrGroups` isto. Postojeći redovi dobijaju `RAD` — ponašanje se ne menja.
2. `BuildNav` gradi stavke za **obe** sekcije jednom, pri gradnji forme. Nema
   pravljenja i rušenja kontrola u radu, pa nema ni ponovnog ožičavanja
   (`WireBtn` → `clsFlatBtn` `WithEvents`); to je isti razlog zbog kog
   `modMaticniLookups` drži `mWrappers` živim.
3. Nov `LayoutNav(frm)` (izdvojen iz `BuildNav`) **postavlja Y i vidljivost samo
   aktivnoj sekciji**. `NavCollapse` se ne dira.
4. `Case "btnMatic"` obrće `mSekcija`, zove `LayoutNav` i prelazi na prvi
   **dozvoljen** ekran te sekcije (`ScrAktivan`).

Zašto ne popup meni ispod zlatnog dugmeta, kao u legacy-ju: dok si u matičnom
ekranu sidebar bi pokazivao **izabranu stavku iz radne sekcije** — ljuska bi
tvrdila da si na Dokumentima dok gledaš Kupce. Prekidač sekcije čuva da sidebar
govori istinu, a to je jedina stvar koju sidebar radi.

Zašto ne dodatna grupa u istom sidebaru: mera iz §26.1.

### 26.3 Odluka B — pet ekrana, ne šesnaest i ne jedan

`MAX_SEG = 11` je gornja granica lista po ekranu, a sekcija sa podacima ima
**trinaest**. Jedan ekran otpada po meri, isto kao sidebar.

Granica ekrana je **legacy grupa**. To nije lenjost nego reuse: grupisanje je u
`modMaticniLookups` uvedeno svesno („Pakovanje … je svoja grupa - vizuelno
podredjena osnovnim sifarnicima, a ne ravnopravno sa Kooperanti/Kupci/Stanice"),
operater ga već zna, i preživelo je do danas.

| # | Ključ | Naslov | Grupa sidebara | Liste | Oblast |
|---|---|---|---|---|---|
| 1 | `MAT_PARTNERI` | Partneri | ŠIFARNICI | Kooperanti · Stanice · Kupci · Vozači · Parcele | `OBL_MATICNI` |
| 2 | `MAT_ROBA` | Proizvodi i cene | ŠIFARNICI | Artikli · Kulture · Cenovnik · Vrsta GP | `OBL_MATICNI` |
| 3 | `MAT_PAKOVANJE` | Ambalaža i pakovanje | ŠIFARNICI | Ambalaža · Palete · Kutije · Kese | `OBL_MATICNI` |
| 4 | `MAT_KORISNICI` | Korisnici | SISTEM | Korisnici (+ prava) | `OBL_MATICNI` + admin |
| 5 | `MAT_SISTEM` | Podešavanja i alati | SISTEM | — (v. §26.9) | `OBL_MATICNI` + admin |

Redovi registra (isti oblik kao postojeći, sa dodatim sedmim poljem):

```
c.Add "MAT_PARTNERI|modScrMatPartneri|OTKUI_NAV_MAT_PARTNERI|" & IC_MAT_PARTNERI & _
      "|SIFARNICI|" & OBL_MATICNI & "|MATICNI"
```

Grupe: `"SIFARNICI|OTKUI_NAVG_SIFARNICI|MATICNI"`, `"SISTEM|OTKUI_NAVG_SISTEM|MATICNI"`.

**Prava.** `OBL_MATICNI` već postoji (`modConfig.bas:677`) i već je oblast u
matrici prava (`frmStammdaten.OblastCaption` zna `OBL_MATICNI`). `ScrDozvoljen`
→ `modAuth.KorisnikImaPravo` time pokriva tri ekrana bez ijedne nove provere.
Grupa SISTEM zadržava **dodatnu** admin branu koju `MaticniMenu_OnClick` danas
primenjuje (`modAuth.MozeAdministraciju`) — `OBL_MATICNI` sam po sebi nije
administracija. Tvrde brane u `modAdmin`/`modPodesavanja` ostaju gde jesu; ovo
je i dalje samo UI brana.

### 26.4 Odluka C — editor je u zoni ekrana i otvara se samo kad se uređuje

Obrazac postoji i proveren je: `modScrPalete` lista „Nova prerada" drži polja u
**svojoj zoni**, gradi ih `modOtkupUI.NewFieldG` (prefiks `scr` je obavezan —
bez njega promena ide ljusci), a `Scr_Layout` ih pali, gasi i raspoređuje; zona
raste samo za tu listu (`PAL_ZONA_H` → `PAL_ZONA_NOVA_H`).

Matični ekran ima dva stanja zone:

- **zatvorena** (`KPI_H`, 56): naziv sekcije, koliko zapisa, koliko aktivnih /
  neaktivnih, dugme **„+ Nova stavka"**;
- **otvorena** (≈ 56 + 3 reda po 46 + 34): do **deset** polja u redovima po
  četiri (isti budžet koji `frmStammdaten` ima kroz `txtField1..10`), linija
  koja kaže da li je unos nov ili izmena kog ID-ja, i dugmad **Sačuvaj** /
  **Odustani**.

Tok je isti kao u legacy formi, samo bez druge forme: klik na red u mreži bira
zapis; **Izmeni** puni polja iz reda i otvara zonu; **+ Nova stavka** otvara
praznu; **Sačuvaj** upisuje i zatvara; **Odustani** zatvara bez upisa.

Izbor reda vraća `False` iz `Scr_Event` (mreža se ne čita ponovo, operater ne
gubi mesto u listi) — pravilo iz `modScrOporavak` i `modScrPalete`. Upis vraća
`True`.

**Podnožje mreže ne treba nijednu izmenu ljuske.** `ModeHasKgCol` i
`ModeHasValCol` čitaju **opis kolona ekrana** (`mCols`), ne `ActiveMode`; lista
koja ne prijavi nijednu kolonu vrste `kg` ni `rsd`/`mult`/`sum0`/`rest` dobija
podnožje bez oba zbira. Za šifarnike se time samo prijavljuju vrste kolona
pošteno — „Težina gajbice" je `num`, ne `kg`, jer nije zbirna veličina. Ovo je
provereno u kodu, ne pretpostavljeno; §13 i §14 su isti kvar plaćali dvaput.

### 26.5 Odluka D — jedan pisac, i zašto ovde drugačije nego u Fazi B

Faza B je **namerno** ostavila legacy formama njihovu kopiju upisne logike
(v. §5). Ovde se to **ne ponavlja**, iz dva merljiva razloga:

1. Za dokumente je postojao operativan put (legacy) i neproveren nov; kod
   matičnih podataka postoji **samo jedan** put — forma. Druga kopija bi bila
   prva prilika za razlaz, a ne osiguranje od njega.
2. `docs/DOMEN/WHO_WRITES.md` pokazuje da matične tabele imaju **0–2** modula
   koji ih pišu (`tblKulture`, `tblKupci`, `tblStanice`, `tblTipAmbalaze`,
   `tblSEFConfig` — nula van testova; `tblKooperanti` — `modKooperant`;
   `tblArtikli` — `modAgrohemija`; `tblParcele` — `modGeoParcele`,
   `modMasterSync`; `tblKorisnici` — `modAuth`, `modSetup`). To je tačno
   situacija iz CLAUDE.md §2 („isto polje često piše više modula") u kojoj
   dodavanje još jednog pisca najviše košta — pogotovo što su ove tabele
   **površina sinhronizacije** (`modStammdatenSync`, `modMasterSync`).

Zato: provere i upis se **izdvajaju** iz `frmStammdaten.btnDodaj_Click` /
`btnIzmeni_Click` / `OnSoftDeleteClick` u nov `modMaticniUnos`, a **legacy forma
se prevezuje na njega**. To je pravilo §6 kataloga primenjeno doslovno
(„ako rutina postoji ali je `Private` i vezana za formu, prvo se izdvaja račun
iz prikaza"), isti potez kao `KoopRangRows` iz `LoadKoopRang` i
`modIzvestaj.StampajReversAmbalaze` iz §23.1.

Rizik se ne prećutkuje: prevezivanje dira **jedini operativan put** za matične
podatke. Zato M2 (v. §26.10) ide sam, sa dvosmernim dokazom, i pre njega stoji
M1 koji ne dira upis uopšte.

### 26.6 Šta ekran uzima od ljuske — i šta ljuska mora da dobije

Uzima, bez izmene: mrežu (sortiranje, pretraga, strane, označen red), prekidač
lista (`Scr_Liste`), čipove (`Scr_Cipovi`), radnje nad redom (`Scr_Radnje`),
naslovnu traku, toast, `NewFieldG` polja u zoni, granu `chg:` za promenu u
polju, `row:` i `dbl:`.

Ljuska dobija **tačno tri** izmene, sve u §26.2:

| Izmena | Gde | Veličina |
| `SCR_SEKCIJA` u registru + u `ScrGroups` | `modUiScreens` | ~15 linija |
| `LayoutNav` izdvojen iz `BuildNav`, gradnja obe sekcije | `modOtkupUI.BuildNav` | ~40 linija |
| `Case "btnMatic"` → prekidač sekcije | `modOtkupUI` | ~10 linija |

Ništa od toga ne menja ponašanje radne sekcije: dok je `mSekcija = "RAD"`,
`LayoutNav` postavlja iste stavke na iste Y koje `BuildNav` danas računa.

### 26.7 Liste, kolone, čipovi i radnje — po ekranu

**Čipovi** su svuda isti i nose jedinu poslovnu osu koju šifarnik ima —
soft-delete: `sve` · `aktivni` · `neaktivni`. Sekcija bez kolone statusa ne
prijavljuje nijedan čip; koja je to sekcija ne pogađa se nego se **pita**, istom
proverom koju legacy koristi (`AktivanColName` → `Aktivan` ili `Aktivna`,
`GetColumnIndex`). `tblCenovnik`, `tblTipAmbalaze` i `tblTipPalete` po zatečenoj
šemi nemaju status; `tblKutije`, `tblKese`, `tblVrstaGP` ga imaju.

**Radnje nad redom** (`Scr_Radnje`, budžet `MAX_ACT = 5`):

| Lista | Radnje |
| sve osim Cenovnika i Parcela | `izmeni` (soft) · `status` — „Deaktiviraj/Aktiviraj" (danger) |
| Cenovnik | `novacena` — „Nova cena" (soft); **nema `izmeni`** — append-only, isto pravilo po kom legacy krije `btnIzmeni` |
| Parcele | `izmeni` · `status` · `geo` — „Geo" (soft), otvara postojeće ulazne tačke (v. §26.9) |

**Kolone lista** se ne izmišljaju — preslikavaju se iz `LoadList`, koja je za
svaku sekciju već izabrala šta operater vidi (npr. Kooperanti: ID, ime+prezime,
telefon, **naziv** stanice a ne `StanicaID`, BPG, tekući račun, PIN,
adresa+mesto, JMBG, status). Prioriteti kolona (peto polje opisa) prate pravilo
iz §9.2: identitet i naziv su 1, ostalo 2–3, tehnička polja 4 (nevidljiva).

**Identitet reda.** Svaka lista nosi PK tabele u koloni 1 (`KooperantID`,
`StanicaID`, …) i radnja bira **po njemu**, nikad po nazivu. Za matične podatke
to nije ista klasa opasnosti kao kod dokumenata (nema generacija), ali istoimeni
zapisi su ovde **obična pojava** — §18 to već beleži („duplikati istog naziva su
u ovim šifarnicima obična pojava"). `frmStammdaten` to danas rešava
`m_RowMap`-om iz pozicije u listboxu; posle sortiranja i pretrage u novoj mreži
pozicija ne znači ništa, pa PK u koloni 1 nije udobnost nego uslov.

### 26.8 Šta je namerno drugačije od legacy-ja

- **Nema modalnih `MsgBox`-eva za ishod.** Uspeh ide u toast, greška u crveni
  toast — kao na svih pet postojećih ekrana. `MsgBox` ostaje samo tamo gde je
  **potvrda pre radnje koja menja podatke** (deaktivacija), po obrascu
  `OdbaciIspravku` iz `modScrOporavak`.
- **Pretraga i sortiranje nad svakom listom.** `lstData` ih nema; šifarnik od
  hiljadu kooperanata se danas pregleda skrolovanjem.
- **Status je pilula, ne tekst „DA/NE".** Mreža to već ume (`PaintPill`).
- **Prelazak između sekcija ne zatvara i ne otvara formu.** Danas svaki izbor
  radi `Unload frmStammdaten` pa `Show` — otud i `m_SetupDone` i
  `mChromeRemoved` čuvari. Prekidač lista je promena jednog polja.
- **Naziv umesto ID-ja u prikazu, ID u podacima.** Legacy to već radi za stanicu
  kooperanta (`LookupValue`); pravilo se primenjuje na sve strane ključeve, uz
  obrazac `PuniPartnerCombo` iz §9 za combo-e (dve kolone, čist ID u drugoj).

### 26.9 Šta NIJE preneto (i zašto)

- **Admin panel** (deset radnji). Nije lista nego spisak dugmadi; ugovorni ekran
  uvek dobija mrežu (`LayoutScreenZone` bezuslovno slaže `zTitle` + `zGrid`), pa
  bi ekran bez mreže tražio zahvat u ljusci — a dobio bi ništa. Ostaje
  `modAdmin.BuildAdminPanel` na `frmStammdaten`; ekran `MAT_SISTEM` ga otvara
  dugmetom zone (`frmStammdaten.Tag = "Admin"`, pa `Show vbModeless`).
- **Podešavanja.** `modPodesavanja` radi, ima grupisanje i sklapanje grupa, i
  nosi bezbednosno pravilo (interni/anti-tamper ključevi se **namerno** ne
  prikazuju). Prelazak bi bio redizajn onoga što radi — `minimal change over
  idealized redesign`. Otvara se iz `MAT_SISTEM` isto kao Admin.
- **GEO alatke Parcela** (GeoSrbija, Google Maps, polygon editor). To su spoljni
  prozori i pozivi; ostaju gde jesu i zovu se iz radnje `geo` nad redom, kroz
  iste ulazne tačke (`OpenGoogleMaps`, `OpenParcelPolygonEditor`,
  `modGeoParcele`). U zonu ulazi samo unos/paste koordinata, jer je to polje.
- **Legacy popup `frmMaticniPodaci` i `frmStammdaten`** — ne diraju se i ostaju
  potpuno operativni za `frmOtkupAPP`, po pravilu §5. Jedina izmena u
  `frmStammdaten` je prevezivanje upisa na `modMaticniUnos` (M2).

### 26.10 Redosled

Manji-pa-veći, i tako da svaki korak sam po sebi ima smisla ako se stane.

| Korak | Šta | Zašto tim redom |
| **M0** | `SCR_SEKCIJA`, `LayoutNav`, prekidač na `btnMatic`, grupa SISTEM sa dva dugmeta na legacy panele | Dugme prestaje da laže. Nijedan matični podatak se još ne dira. |
|---|---|---|
| **M0** | `SCR_SEKCIJA`, `LayoutNav`, prekidač na `btnMatic`, ekran `MAT_SISTEM` kao pokretač legacy panela | **URAĐENO** (`v6-ui-187`) — v. §26.13. Dugme prestaje da laže. Nijedan matični podatak se još ne dira. |
| **M1** | `modMaticniIzvor` (opis 13 sekcija) + `modMaticniEkran` (zajedničko telo) + tri tanka ekrana kao **pregled** | **URAĐENO** (`v6-ui-188`) — v. §26.14. Ceo čitalački deo bez ijednog upisa. Legacy netaknut. |
| **M2a** | `modMaticniUnos` — izdvajanje provera i upisa iz forme; **legacy prevezan** | **URAĐENO** (`v6-ui-189`) — v. §26.15. Jedan pisac (§26.5). |
| **M2b** | editor u zoni ekrana, radnje `izmeni` i `status` | **URAĐENO** (`v6-ui-190`) — v. §26.16. |
| **M3** | Cenovnik (append-only) i Parcele (GEO panel u zoni) | **URAĐENO** (`v6-ui-191`) — v. §26.17. Cenovnik je stigao već uz M2b. |
| **M4** | `MAT_KORISNICI` — lista + matrica prava | **URAĐENO** (`v6-ui-193`) — v. §26.18. Zatvara i nalaz o deaktivaciji korisnika. |
| **M5** | Odluka o Podešavanjima i Adminu | **URAĐENO** (`v6-ui-194`) — v. §26.19. Odgovor: **ostaju**, sa izmerenim razlozima; test 160 drži pokrivenost menija. |

### 26.11 Verifikacija — šta se meri

Merilo prelaska je **SLAGANJE sa legacy prikazom i upisom**, ne izgled:

- za svaku od 13 sekcija test tvrdi da lista ekrana ima **iste redove i iste
  vrednosti** kao `LoadList` iste sekcije (nezavisan read-model, po obrascu
  §23.6); istoimeni zapisi u fixture-u su **obavezni**, da bi test uopšte umeo
  da razlikuje izbor po PK od izbora po nazivu;
- upis: test tvrdi da `modMaticniUnos` posle izdvajanja piše **iste ćelije**
  koje je forma pisala (snimak reda pre/posle), za svaku sekciju posebno;
- soft-delete: `AktivanColName` mora naći kolonu tamo gde je legacy nalazi, i
  ne naći je tamo gde je nema — inače čip „aktivni" tiho pokazuje sve;
- čipovi: ugovor „svaki čip koji ekran prijavi ume da se izabere"
  (`SegIndeksIzTaga` presedan);
- podnožje: test tvrdi da nijedna matična lista ne prijavljuje kolonu `kg` ni
  novčanu vrstu, pa podnožje ne crta zbirove (§26.4);
- **dvosmerni dokaz** (`tools/dokaz.py`) je ovde **obavezan**, ne opcion: M2
  menja jedini operativan put za matične podatke;
- ručna kapija: `Alt+F11 → Debug → Compile VBAProject`, pa smoke nad pravim
  podacima — prekidač sekcije na najmanjem i na maksimizovanom prozoru, sve
  liste, unos i izmena po sekciji, GEO alatke, i **legacy meni i dalje radi**.

`run_vba` traži Windows + Excel, pa se u web sesiji izmena ponašanja prijavljuje
kao **neverifikovana**, nikad kao zelena (CLAUDE.md §5).

### 26.12 Otvoreno

- **Deset polja u tri reda po četiri** je gornja granica zatečenog modela
  (`txtField1..10`). Kupci realno traže više (ulica, mesto, poštanski broj,
  država, PIB, MB, e-mail, hladnjača, tekući račun, naziv = 10 tačno). Ako se
  neka sekcija proširi, red se dodaje u zoni, ne u ljusci.
- **Značka uz stavku menija** (`Scr_Brojac`) za šifarnike nema šta da broji.
  Ne implementira se — ekran koji je ne implementira dobija nulu, a nula se ne
  crta.
- **Padajući redovi detalja** (parcele kooperanta ispod kooperanta, cene ispod
  kulture) — isti odložen posao kao u §5/Faza C; ne uvodi se ovde.
- **Sinhronizacija.** `modStammdatenSync` i `modMasterSync` čitaju iste tabele.
  M2 ne menja šta se upisuje, pa ne bi smeo da ih dotakne — ali to je tvrdnja
  koju test mora da drži, ne pretpostavka.
  **M4 jeste promenio šta se upisuje** (`tblKorisnici.Aktivan`: `DA`/`NE` umesto
  `Aktivan`/`Neaktivan`). Provereno je da tu tabelu **ne dodiruje nijedan**
  sinhronizacioni put — nema pogodaka u `modStammdatenSync`, `modMasterSync`,
  `gas/` ni `src/`. Jedini čitač je `modAuth`, i on je namerno ostavljen kakav
  jeste. **Ostaje otvoreno:** migracija zatečenih `"Aktivan"`/`"Neaktivan"`
  vrednosti u rečnik `DA`/`NE`. Dok se ne uradi, takav zapis se **čita kao
  aktivan** — što je isto ponašanje kao pre M4, ne regresija.

### 26.13 M0 — šta je stvarno urađeno (`v6-ui-187`)

Nosač je gotov. Tri stvari su u izvedbi ispale **drugačije nego u planu**, i
razlog je u svakom slučaju nešto što se videlo tek u kodu.

**(1) `MAT_SISTEM` je EKRAN sa listom, ne zonsko dugme.** §26.9 je govorio da
Podešavanja i Admin otvara „dugme zone". Ali ugovorni ekran **uvek** dobija
mrežu (`LayoutScreenZone` bezuslovno slaže `zTitle` + `zGrid`), pa bi ekran sa
dugmetom u zoni imao ispod sebe praznu mrežu koja izgleda kao pad. Lista alatki
(dve, sa kolonama *Alatka · Šta radi · Otvara se* i radnjom „Otvori") je jedina
lista koja tu ima smisla. Uz to daje matičnoj sekciji **bar jedan ekran od
prvog dana** — bez njega bi zlatno dugme vodilo u sidebar u kom je sve
prigušeno. Oba panela ostaju netaknuta: `modScrMatSistem.OtvoriPanel` radi tačno
ono što `frmMaticniPodaci.OpenSekcija` radi (`Unload` → `Tag` → `Show`).

**(2) Ugovor ekrana je dobio `Scr_Dozvoljen`.** Grupa SISTEM traži
**administraciju**, a to nije oblast iz `tblKorisnici` pa ne staje u
`SCR_OBLAST`. Umesto da ljuska dobije izuzetak po imenu ekrana, `ScrDozvoljen`
sada pita i sam ekran; ekran koji tu funkciju nema se ponaša kao pre. Keširа se
**samo činjenica da funkcija postoji**, nikad njen odgovor — prava se menjaju
zamenom operatera.

**(3) Popravljen zatečen kvar koji bi M0 obesmislio.** `PaintNav` je svakoj
neizabranoj stavci vraćao punu boju, pa su Marža i Sledljivost izgledale
prigušeno **samo do prvog klika** bilo gde u meniju. U matičnoj sekciji bi to
bilo **tri od pet** stavki: operater bi video ekrane koji izgledaju kao da rade.
Stanje se sada čita iz mape napravljene pri gradnji (`mNavOff`) — tamo gde je
odgovor ionako već izračunat, pa se ne može razići sa onim što `BuildNav`
nacrta.

**Uz to:** „Nazad na rad" vraća **ekran sa kog se otišlo** (pamti se po
sekciji), a ne uvek Unos dokumenata; `ModeHasKgCol` / `ModeHasValCol` su
izdvojeni u čiste funkcije nad opisom kolona (`OpisImaKgKolonu`,
`OpisImaValKolonu`), da bi se „ova lista nema zbirove" moglo tvrditi **pre**
crtanja; gola `62` u `LayoutOtkup` je postala `PROFIL_H`, jer ulazi u meru iz
§26.1 i mora da se može tvrditi u testu.

**Mera iz §26.1 je sada test, ne komentar.** `LayoutNav` pamti visinu koju je
sekcija zauzela, a test 150 je čita iz **stvarnog rasporeda** i tvrdi tri
stvari: radna sekcija staje (405 ≤ 492), matična staje (211 ≤ 492), a **zajedno
ne staju** (616 > 492) — to poslednje je jedini razlog zbog kog sekcije postoje.
Test zato hvata i izmenu koja nema veze sa matičnim podacima: dodavanje šestog
ekrana u OPERACIJE.

**Verifikacija:** `vba_check` čist; pet novih sabotaža pod prefiksom `maticni-`.
Testovi 150 i 151 su napisani ali **nisu izvršeni** — `run_vba` traži Windows +
Excel, pa se u web sesiji ponašanje prijavljuje kao **neverifikovano**. Ostaje
ručna kapija: `Alt+F11 → Debug → Compile VBAProject`, `RunAllTests`,
`dokaz.py maticni`, pa smoke: prekidač sekcije na najmanjem i maksimizovanom
prozoru, suženi sidebar, otvaranje oba panela, i **legacy meni i dalje radi**.
Tri MDL2 glifa (`IC_MAT_PARTNERI`, `IC_MAT_ROBA`, `IC_MAT_PAKOVANJE`) su izabrana
iz kataloga a **nisu vizuelno proverena** — provera je `DumpMdl2Used`.

### 26.14 M1 — pregled bez ijednog upisa (`v6-ui-188`)

Tri ekrana rade. **Nijedna linija upisa nije napisana** i `frmStammdaten` nije
dirana: unos i izmena i dalje idu kroz staru formu, a to zona i kaže naglas
(„Pregled. Unos i izmena i dalje idu kroz stari ekran matičnih podataka.") —
u zoni, ne u dijalogu, jer se dijalog zatvori i zaboravi.

**Podela na tri modula, ne na tri ekrana.**

| Modul | Zna | Ne zna |
|---|---|---|
| `modMaticniIzvor` | tabele, PK, kolone, izvedene vrednosti, čipove, sort | ništa o kontrolama i ljusci |
| `modMaticniEkran` | zonu, mrežu, prekidač lista | ništa o tabelama |
| `modScrMat*` (3×) | **samo** koje sekcije nosi | sve ostalo |

Ekranski moduli su po ~70 linija. Tri kopije istog rasporeda bi se razišle prvom
doradom — to je već plaćeno na četiri kopije `Prefill*` (Z10).

**Šta je preneto 1:1 iz `LoadList`:** koje kolone operater vidi i kojim redom, i
svih šest izvedenih vrednosti (puno ime, **naziv** stanice umesto ID-ja, spojena
adresa, adresa kupca sa poštanskim brojem, kooperant kao „Ime Prezime (ID)",
geo status/izvor). Cenovnik zadržava svoj poredak (`ExcludeStornirano` →
`SortArray` po datumu opadajuće, tie-break `CenaID`) — to je poslovno pravilo,
ne prikaz: važeća cena je poslednja.

**Tri stvari su namerno drugačije od legacy-ja:**

1. **Kolona statusa se TRAŽI u šemi, ne pogađa.** Parcele nose `Aktivna`, ostali
   `Aktivan`, a Cenovnik, Ambalaža i Palete je po zatečenoj šemi nemaju uopšte.
   Isti probe koji legacy radi u `AktivanColName`, i isti razlog: sekcija bez te
   kolone **ne dobija čipove**. Čip „Aktivni" koji tiho pokazuje sve je gori od
   nikakvog čipa. Test to tvrdi po imenu (`maticni-status-pogadja-umesto-da-trazi`).
2. **Identitet je PK u koloni 1, uvek.** Legacy bira red po poziciji u listboxu
   (`m_RowMap`); u novoj mreži pozicija ne znači ništa posle sortiranja i
   pretrage, a istoimeni zapisi su u šifarnicima obična pojava (§18).
3. **Status ostaje tekst, ne pilula.** Vrsta ćelije `pill` ima **zatvoren rečnik**
   od tri stanja dokumenta (Otkazana / Poslato / Sačuvana) — aktivnom kooperantu
   bi pisalo „Sačuvana". Pilula za Aktivan/Neaktivan je nova vrsta ćelije i nije
   M1.

**Ugovor ekrana je dobio `Scr_Sort`.** Ljuskino pravilo je „druga kolona
opadajuće" — napisano za datum dokumenta, a nad Stanicama znači nazive unazad.
Umesto da se `SortZaListu` dopisuje sa još trinaest imena lista pored zatečena
dva, ekran kaže svoj sort sam. Pravilo je jedno: **rastuće po koloni koju opis
označi kao rastegljivu** (`part`) — računa se iz opisa kolona, pa dodavanje
sekcije ne traži još jedan spisak koji bi se razišao. Cenovnik je jedini izuzetak
i to je poslovno pravilo.

**Podnožje nema lažne zbirove** — nijedna matična kolona nije `kg` ni novčana;
težine su `num`, jer jesu broj ali nisu zbirna veličina. Test to tvrdi nad
opisom, pre crtanja (§26.4).

**Verifikacija:** `vba_check` čist (203 fajla, 331 sabotaža). Dva nova testa
(152 — opis svih 13 sekcija je potpun i saglasan sa šemom; 153 — čip, identitet
i pretraga nad stvarnim podacima) i četiri nove sabotaže. Test 153 **pada ako
nijedna sekcija nema redove u fixture-u** — zelena tvrdnja nad praznim skupom ne
dokazuje ništa. Kao i M0: **testovi nisu izvršeni** u web sesiji.

**Otvoreno posle M1:** poređenje 1:1 sa `LoadList` po sekciji (traži Excel);
pilula za status; `Scr_Brojac` se namerno ne implementira (šifarnik nema šta da
broji, a nula se ne crta).

### 26.15 M2a — jedan pisac (`v6-ui-189`)

Provere i upis su izašli iz `frmStammdaten` u `modMaticniUnos`, **i forma je
prevezana na njega**. `btnDodaj_Click` je bio 544 linije, `btnIzmeni_Click` 463;
danas su po ~30, i ne sadrže nijedno poslovno pravilo. Formi je ostalo samo ono
što forma jeste — koja kontrola nosi koje polje (`KontrolaZaPolje`, jedan
`Select Case`).

**Ulaz pisca je rečnik, ne kontrole.** `MatDodaj(kljuc, polja, noviID)` /
`MatIzmeni(kljuc, red, polja)` / `MatPromeniStatus(kljuc, red, noviStatus)` —
isti obrazac koji ljuska već koristi za `Scr_Save`. Modul u isti rečnik upisuje
`fokus` = ključ odbijenog polja, pa forma ume da vrati kursor tamo gde je bila
greška, kao i pre.

**Opis polja je jedan, i služi dvoma.** `modMaticniIzvor.MatPolja` opisuje šta
se unosi (ključ, natpis, vrsta, obavezno, kolona, izvor combo-a). Iz njega čita
pisac; iz njega će u M2b čitati i editor. Dva spiska bi se razišla.

**Jedna namerna izmena ponašanja: unos ide PO IMENU KOLONE.** Legacy je za deset
sekcija zvao `AppendRow` sa nizom vrednosti, pa je tačnost zavisila od redosleda
kolona — a šema se razlikuje po instalaciji; baš zato `LoadList` za Kupce već
ima toleranciju, a `btnIzmeni` za Stanice alias-probe. Sada je svuda obrazac koji
legacy već koristi za Korisnike i Kulture: prazan red pa `RequireUpdateCell` po
imenu, sve u `clsTransaction`. **Upisuju se iste ćelije; menja se samo kako se
adresiraju.** Alias-probe (`@alias:Kontakt,Telefon`) time važi i za unos, ne samo
za izmenu.

**Šta je namerno OSTALO nedosledno** — jer bi poravnanje bilo neizmerena izmena
na jedinom operativnom putu upisa:

- **Parcele pri unosu dobijaju status `"Da"`**, a soft-delete u istu kolonu piše
  `"Aktivan"`/`"Neaktivan"`. Čitač aktivnim smatra sve što nije `"Neaktivan"`, pa
  oba oblika prolaze. Zapisano u `MatStatusNaUnosu`.
- **Prazan broj:** pri unosu ćelija ostaje prazna, pri izmeni se upisuje nula.
  To je zatečeno (`btnDodaj` za Kulture preskače praznu gajbicu, `btnIzmeni`
  upisuje 0) i čuva se doslovno.

**Šta NIJE prešlo:** Korisnici. `tblKorisnici` nosi matricu prava i
`PreparePin`, i ide u M4 zajedno sa svojim ekranom; do tada forma zadržava svoje
`DodajKorisnika` / `IzmeniKorisnika` / `SoftDeleteLegacy` — tela su **prenesena
neizmenjena** iz nekadašnjih `Case "Korisnici"` grana.

**Nalaz usput, NIJE popravljen** (dira prijavu, pa ne ide slepo): dugme
„Deaktiviraj" nad Korisnicima upisuje `"Neaktivan"` u `Aktivan`, a `modAuth`
deaktiviranim smatra **samo** vrednost `"NE"` (`modAuth.bas:87`). Deaktivacija
korisnika kroz to dugme dakle **ne sprečava prijavu** — u listi piše
„Neaktivan", a login i dalje prolazi. Ovo je zatečeno stanje, ne posledica M2a
(putanja je prenesena netaknuta u `SoftDeleteLegacy`). Rešava se u **M4**, gde
se odlučuje i šta se upisuje i šta `modStammdatenSync` i PWA očekuju.

**Verifikacija:** `vba_check` čist (204 fajla, 335 sabotaža). Dva nova testa —
154 (opis polja svih 13 sekcija razrešava se u kolonu koja **postoji u šemi**;
PK nije polje za ručni unos; alias se stvarno razrešava) i 155 (provere odbijaju
prazno obavezno, ne-broj, negativno, nulu tamo gde nema smisla, i ukršteno
pravilo pragova — uz to imenuju **koje** polje). Oba mere **bez ijednog upisa**
(`MatProveriTest`), pa ne diraju tabele. Četiri nove sabotaže pod istim
prefiksom. Kao i ranije: **testovi nisu izvršeni** u web sesiji.

**Ručna kapija za M2a je stroža nego dosad**, jer je ovo jedini put upisa:
`Compile`, `RunAllTests`, `dokaz.py maticni`, pa smoke nad **kopijom** radne
sveske — unos i izmena po svakoj od 13 sekcija, deaktivacija/aktivacija,
nova cena, nova stanica u MALINA režimu (ogledalo vozača), i Korisnici (koji
nisu dirani).

### 26.16 M2b — editor u zoni (`v6-ui-190`)

Tri matična ekrana sada **unose i menjaju**. Nijedno pravilo nije napisano
ponovo: sve ide u `modMaticniUnos`, istog pisca koga zove i legacy forma.

**Zona ima dva stanja.** Zatvorena je pregled (naziv sekcije, koliko zapisa,
aktivni/neaktivni) plus dugme **„+ Nova stavka"**. Otvorena nosi polja sekcije,
**Sačuvaj** i **Odustani**, i naslov koji kaže da li je nov zapis ili izmena
kog ID-ja. Obrazac je iz liste „Nova prerada" na Paletama — polja postoje uvek,
`Scr_Layout` ih pali, gasi i raspoređuje.

**Polja su BAZEN, ne po sekciji.** `NewFieldG` pri gradnji odlučuje pravi li
tekst ili combo, a sekcije se razlikuju — pa se gradi **deset tekstualnih i šest
combo** polja (tačno budžet koji `frmStammdaten` ima kroz `txtField1..10` /
`cmbField1..6`), a raspored svakoj sekciji dodeljuje sledeće slobodno polje
njene vrste. Kupci troše svih deset tekstualnih; ništa ne prelazi bazen.

> **Zašto baš to ima svoju sabotažu:** da dva polja dobiju istu kontrolu, jedno
> bi pri upisu tiho pregazilo drugo — operater unese telefon, a snimi se kao BPG
> broj. Upis prolazi, poruka je „Dodato", ništa ne prijavljuje grešku. Test 156
> tvrdi da su sve dodeljene kontrole **različite**, za svih 13 sekcija.

**Ljuska je dobila jedan izlaz:** `OsveziRasporedEkrana`. Otvaranje editora menja
**visinu zone**, a `LayoutScreenZone` se izvršava isključivo iz `LayoutOtkup` —
ni `RefreshFromData` ni `ReloadGrid` ga ne diraju (oni preračunavaju *mrežu*).
Bez toga bi polja bila nacrtana ispod mreže, što se vidi kao „dugme ne radi".
Ekran to traži sam; ljuska i dalje ne zna nijedan ekran po imenu.

**Otvaranje i zatvaranje editora vraćaju `False`** iz `Scr_Event` — to su
promene prikaza, ne podataka. Da vraćaju `True`, svaki klik na „Nova stavka"
platio bi pun `RefreshFromData` (KPI, keš, combo-i). `True` vraćaju samo upis i
promena statusa.

**Radnje nad redom prate ono što sekcija stvarno ume:** `Izmeni` svuda osim na
Cenovniku (append-only — isto pravilo po kom legacy krije to dugme), a
`Deaktiviraj / Aktiviraj` **samo gde kolona statusa postoji**. Dugme koje tiho ne
radi ništa je gore od dugmeta kog nema. Promena statusa traži potvrdu, po
obrascu iz Oporavka.

**Cenovnik je dobio nešto što legacy nema:** klik na `Izmeni` (ili dvoklik)
otvara **nov unos sa vrednostima izabranog reda** — proizvod, sorta i klasa
ostaju, cena i datum se unose novi. Legacy je za isto krio dugme i tražio da se
sve otkuca ponovo. Zapis i dalje ide kroz `AddCena`, dakle append-only; menja se
samo koliko se kuca.

**Identitet:** red se bira po **PK iz kolone 1**, a `MatRedPoID` ga prevodi u red
tabele. Pozicija u mreži se ne koristi nigde — posle sortiranja i pretrage ona
ne znači ništa.

**Verifikacija:** `vba_check` čist (204 fajla, 338 sabotaža). Dva nova testa —
156 (bazen polja se ne preklapa; combo polje dobija combo kontrolu; otvorena
zona je viša od zatvorene) i 157 (radnje prate mogućnosti sekcije; oblik opisa
radnje). Tri nove sabotaže. Kao i ranije: **testovi nisu izvršeni** u web sesiji.

**Otvoreno posle M2:** Parcele još nemaju GEO radnju (M3); Korisnici i njihova
matrica prava idu u M4, zajedno sa nalazom iz §26.15 o deaktivaciji koja ne
sprečava prijavu.

### 26.17 M3 — GEO parcele (`v6-ui-191`)

Cenovnik je zatvoren već uz M2b (append-only + prefill iz izabranog reda), pa je
M3 sveden na **GEO**. Šest legacy dugmadi uz listu parcela postalo je
`modMaticniGeo` — operacije **po `ParcelaID`-ju**, bez ijedne kontrole.

**Zašto po ID-ju:** legacy je slao `m_SelectedRow`, redni broj izveden iz
pozicije u listboxu. U mreži koja se sortira i pretražuje pozicija ne znači
ništa — a `modGeoParcele` **već ima** `*ByID` varijantu svake operacije, sa
`RequireSingleParcelaRow` kapijom. Koristi se ona; ništa se nije prepisivalo.

**Zona je dobila treće stanje.** Zatvorena / editor / **geo**. Editor i geo se
isključuju — jedna stvar u zoni u isto vreme; dva panela jedan preko drugog su
ista klasa kvara kao traka `zOtp` koja je ostajala upaljena na tuđem ekranu.
Prelazak na drugu listu zatvara oba.

GEO panel nosi koordinate izabrane parcele, red opisa (status / izvor / ima li
poligon) i šest alatki: **Sačuvaj geo · Nalepi koordinate · GeoSrbija · Google
Maps · Poligon · Obriši geo**, plus Zatvori. Radnja `geo` nad redom postoji
**samo na Parcelama** — jedina sekcija sa koordinatama; dugme na ostalima bi
otvaralo panel koji nema šta da pokaže.

**Dva pravila su izašla iz forme, i forma ih sada zove:**

1. **Prepoznavanje koordinata iz nalepljenog teksta** (`GeoIzTeksta`). Prag
   `|d| > 1000` nije ukras: UTM34 nad Srbijom je sedmocifren, pa se time
   odbacuju broj parcele, godina i sve ostalo iz reda prekopiranog sa portala.
   Bez praga bi „Parcela 123" dala koordinatu 123. Test to tvrdi po imenu.
2. **Adrese** (`GeoUrlMape`, `GeoUrlPoligon`, `GEO_URL_SRBIJA`). Adresa mape
   mora imati decimalnu **tačku**: na mašini sa zarezom kao separatorom `CStr`
   daje `44,81`, pa bi URL bio neispravan — mapa bi se otvorila na pogrešnom
   mestu ili nikako. `frmStammdaten` više nema nijedan URL u sebi.

**Brisanje traži potvrdu.** Legacy je to rešavao dugmetom u dva koraka
(„Potvrdi brisanje"); ovde je `MsgBox` potvrda — ista brana, čitljivija, i ista
kao kod promene statusa. Tačka i poligon se gube, a poligon se ručno crta.

**Verifikacija:** `vba_check` čist (205 fajlova, 341 sabotaža). Nov test 158 —
prepoznavanje koordinata (sa oznakama `N=`/`E:`, zagradama, tabovima; mali
brojevi se preskaču; manje od dva velika broja = nema koordinata) i adrese
(decimalna tačka, prazan ID ne daje adresu poligona) — plus tvrdnja da geo
radnja postoji samo na Parcelama. Tri nove sabotaže. **Testovi nisu izvršeni**
u web sesiji.

**Otvoreno posle M3:** samo **M4** — Korisnici, matrica prava, i nalaz iz
§26.15 (deaktivacija korisnika ne sprečava prijavu). Posle toga ostaje M5
(odluka o Podešavanjima i Adminu), gde je podrazumevani odgovor „ostaju".

### 26.18 M4 — Korisnici i prava (`v6-ui-193`)

Poslednja sekcija menija koja je ostajala u formi. Ostala je namerno: `tblKorisnici`
nosi PIN i matricu prava po oblasti, pa je čekala svoj ekran.

**Ekran ima dve liste, i druga zavisi od prve.** `MAT_KORISNICI` nosi
**Korisnici** (nalozi) i **Prava** (matrica po oblasti izabranog korisnika).
Zato su na istom ekranu, a ne dva: prekidač je jedan potez, a promena ekrana bi
izgubila izbor. Izabrani korisnik se pamti **odvojeno** od izabranog reda
(`mKorisnikID` vs `mIzabranID`) — čim se na listi Prava izabere red, `mIzabranID`
postaje oblast, a korisnik mora da ostane.

**Zatvoren je nalaz iz §26.15.** Rečnik kolone `Aktivan` u `tblKorisnici` je
`"DA"`/`"NE"` i tako ga čita `modAuth` (`modAuth.bas:87` neaktivnim smatra **samo**
`"NE"`). Generičko dugme „Deaktiviraj" je u nju upisivalo `"Aktivan"`/`"Neaktivan"`
— u listi je pisalo *Neaktivan*, a **prijava je prolazila**. Od M4 sve sekcije
idu kroz `modMaticniUnos`, koji za Korisnike delegira `modMaticniKorisnici`, a
taj piše isključivo `DA`/`NE`.

**Čitač se NIJE dirao.** `modAuth` ostaje kakav jeste: zapisi koji već nose
`"Neaktivan"` i dalje se čitaju kao aktivni, tačno kao do sada. Menja se samo šta
se **upisuje** — nijedan korisnik ovom izmenom ne gubi pristup. Poravnanje
zatečenih vrednosti je odvojen posao (migracija podataka), ne usputna izmena.

**Drugi nalaz, zatvoren usput: spuštanje sa admina.** Adminov red nosi `DA` na
svih dvanaest oblasti zato što je to **posledica uloge**, a ne izbor koji je neko
napravio (`OblComboVal` je adminu uvek pisao `DA`). Kad uloga padne na *Korisnik*,
prenos tih `DA` daje pun pristup koji niko nije dodelio — a `modAuth` ga čita
bukvalno. Pravilo je izdvojeno u `GasiSvaPrava` i glasi: kad prava nisu izričito
poslata, **unos i spuštanje sa admina** gase sve pa se dodeljuje ponovo.

**Prava se ne uređuju u editoru.** Red liste Prava je oblast, ne zapis, i jedina
radnja je *Uključi / isključi*. Zato ta lista nema ni „Izmeni" ni „Deaktiviraj"
ni „Nova stavka" — uslov se čita iz opisa polja (`MatPolja` za nju ne daje
editor), ne iz spiska sekcija, pa nova sekcija bez editora ne može da dobije
dugme koje ne vodi nikuda. Dvanaest combo-a prava ne bi ni stalo u bazen od šest.

**Identitet reda prava je skriven.** Prva kolona nosi *lokalizovan naziv* oblasti,
a ključ (naziv kolone prava) stoji u koloni 4, prioritet 4 = nikad se ne crta.
Zato je uveden `MatKolonaID(kljuc)`: sekcija **prijavljuje** koja kolona nosi
identitet, a `modMaticniEkran` je pita umesto da pretpostavlja kolonu 1. Bez toga
bi promena prevoda menjala to koje se pravo uključuje — kvar bez ijedne greške u
logu. Isti obrazac koji `modScrMatSistem` već ima kroz `MS_COL_TAG`.

**Brana je dvostruka.** Pored oblasti `MaticniPodaci`, ekran traži i
administraciju (`Scr_Dozvoljen` → `modAuth.MozeAdministraciju`). Prava pristupa su
jedina lista na kojoj se može dodeliti pristup samom sebi. Stavka u sidebaru
ostaje **vidljiva ali prigušena** — operater treba da vidi da to postoji i da mu
je zabranjeno, ne da mu nestane. Test seam sme **samo da zatvori** branu.

**Šta je forma izgubila:** `btnDodaj`/`btnIzmeni` grane za Korisnike (~130 linija)
sada su dva poziva `modMaticniUnos`; `NormalizeUloga` je obrisan (pravilo živi u
piscu); `OblComboVal` je ostao bez `isAdmin` parametra — pravilo „admin dobija
sve" bilo je na dva mesta, sada je na jednom. **Zatečena putanja soft-delete-a
(`SoftDeleteLegacy`) je obrisana**: nijedna sekcija je više ne dostiže, a tiho
upisivanje pogrešnog rečnika je upravo ono što se ovim zatvara — nepoznat `Tag`
se sada odbija porukom, isto kao u `btnDodaj`/`btnIzmeni`.

**Forma i ekran zovu istog pisca, ali ne šalju isto.** Forma šalje svih dvanaest
combo-a prava pod prefiksom `obl:`; ekran ne šalje nijedan, pa mu prava ostaju
netaknuta i menjaju se iz svoje liste. Pisac zato upisuje **samo one oblasti koje
je stvarno dobio**.

**Verifikacija:** `vba_check` čist (207 fajlova, 349 sabotaža). Nov test 159 —
rečnik `DA`/`NE` (uključujući da zatečeni tekst ne prolazi kao dozvola), prazan
PIN, obavezno i jedinstveno korisničko ime, pravilo spuštanja sa admina, skrivena
kolona identiteta prava, radnje liste prava, naziv svake oblasti u katalogu, i
brana ekrana. Testovi 152/154/157 prošireni na `MAT_KORISNICI`; tvrdnja „kolona 1
nosi PK" generalizovana u „kolona identiteta koju sekcija prijavi stvarno ga
nosi". Sedam novih sabotaža. **Testovi nisu izvršeni** u web sesiji — nema
Excela; `Debug → Compile` ostaje ručna kapija.

**Otvoreno posle M4:** samo **M5** — odluka o Podešavanjima i Adminu, gde je
podrazumevani odgovor „ostaju" (to su alatke, ne matični podaci; već ih otvara
ekran *Podešavanja i alati*). Van plana ostaje migracija zatečenih
`"Aktivan"`/`"Neaktivan"` vrednosti u `tblKorisnici` u rečnik `DA`/`NE`.

### 26.19 M5 — odluka o Podešavanjima i Adminu (`v6-ui-194`)

Odgovor je **ostaju kao legacy paneli**, što je i bio podrazumevani odgovor iz
§26.10 — ali sada sa izmerenim razlozima, a ne kao neispitana pretpostavka.
Nijedan red produkcione logike se zbog M5 nije premestio; jedino što je M5
proizveo je **test koji odluku drži** (160) i ovaj zapis.

#### Podešavanja — ne staju u ugovor ekrana, i ne treba da stanu

| Mera | Vrednost | Ograničenje ljuske |
|---|---|---|
| polja | **97** | bazen zone je **10 tekstualnih + 6 combo = 16** |
| grupa | **11** | `MAX_SEG = 11` — prekidač lista je **već pun** |
| najveća grupa | 17 (`Otkup / dokumenta`) | i ta jedna grupa ne staje u zonu |
| tipovi | `text · bool · int · list:… · secret · memo` | opis polja zna `txt · num · cmb · date` |

Tri nezavisna razloga, svaki sam za sebe dovoljan:

1. **Ne staje.** Jedanaest grupa je tačno `MAX_SEG`; sledeća dodata grupa obara
   prekidač. Ni jedna jedina grupa ne staje u zonu od 16 kontrola.
2. **Radnja je druga.** Config editor pokazuje **celu grupu odjednom** i čuva
   jednim klikom — tako se instalacija i podešava, red po red. Editor u zoni bi
   isto značio jedno polje po otvaranju. To nije prenos nego pogoršanje.
3. **Rečnik ne postoji.** `secret` i `memo` nemaju par ni u vrsti kolone
   (`txt|num|date|part|kg|rsd|mult|sum0|rest|pill|paypill`) ni u vrsti polja
   (`txt|num|cmb|date`). Uveli bi se dva nova pojma u ugovor koji ih nigde
   drugde ne koristi.

Prepisivanje bi značilo ~850 novih linija umesto ~850 koje rade — suprotno od
`minimal change` (CLAUDE.md).

#### Admin — staje po obliku, ali ne sme po sadržaju

Dvanaest komandi u pet grupa **jeste** lista i legla bi u `MAT_SISTEM` kao još
redova. Ne radi se, i razlog nije oblik nego **šta komande diraju**:
`AdminVbaImport` ponovo učitava module u kojima živi ekran koji crta tu listu;
`AdminEnsureEverything` i `MigrirajPodatkeIzStarog` menjaju šemu ispod nje;
`OcistiTabele` briše podatke; self-update menja samu svesku. Red u mreži nove
ljuske značio bi da ljuska poziva nešto što ruši ljusku **usred sopstvenog
događaja**. Legacy panel koji preuzme prozor tu nije zaostatak nego tačan oblik.

#### Brana je proverena, i nije u ljusci

Da „ostaju" ne bi značilo „nisu pogledani": oba panela imaju **tvrdu branu u
sebi**, nezavisno od sidebara — `modPodesavanja.BuildConfigEditor` i
`modAdmin.AdminPanel_OnClick` oba odbijaju ne-admina (`AUD-033`), a Admin je
gejtovan i na **svakoj akciji**, ne samo pri izgradnji panela. Ekran
`MAT_SISTEM` dodaje treću branu (`Scr_Dozvoljen`). Prigušena stavka u sidebaru
je time samo najspoljašnji sloj, ne jedini.

#### Šta M5 ostavlja iza sebe: test 160

Dve registracije istih sekcija žive odvojeno i to je namerno — stari meni pamti
**grupisanje i natpise** (`modMaticniLookups`), nova sekcija pamti **ekrane i
kolone** (`modMaticniIzvor`). Cena je da mogu da se raziđu: sekcija dopisana u
jedan spisak a zaboravljena u drugom **nestaje iz te ljuske bez ijedne greške**.

Test 160 meri pokrivenost **u oba smera**: svaki `Tag` starog menija mora biti
dostižan u novoj ljusci — ili kao lista matičnog ekrana
(`MatKljucIzLegacyTag`), ili kao alatka na *Podešavanja i alati*
(`MsAlatkaTagovi`) — i nijedna lista nove ljuske ne sme postojati mimo starog
menija. Jedini zabeležen izuzetak je **PRAVA**: matrica prava je u formi bila
deo sekcije Korisnici, a ovde je svoja lista (§26.18).

Brojke koje test drži: **16** sekcija u starom meniju, **15** lista u novoj
sekciji (14 sa parom u meniju + PRAVA), **2** legacy panela na ekranu alatki.

Za ovu tvrdnju stoji **jedna** sabotaža, ne dve: oba smera mere isti raskorak,
pa bi svaka stvarna izmena oborila obe — a sabotaža čija imenovana tvrdnja ne
padne prva gora je od nijedne.

#### Čime je §26 zatvoren

`btnMatic` više ne laže, svih 16 sekcija starog menija dostižno je iz nove
ljuske, jedini put upisa matičnih podataka je jedan pisac koga zovu obe strane,
i dva nalaza o pravima pristupa su zatvorena (§26.18). Ostaje jedan posao van
plana: **migracija zatečenih `"Aktivan"`/`"Neaktivan"` vrednosti** u
`tblKorisnici.Aktivan` u rečnik `DA`/`NE` (§26.12).

### 26.20 Pred brisanje legacy formi — šta je nedostajalo (`v6-ui-195`)

§26 je pisan pod pretpostavkom **koegzistencije**: legacy forma ostaje, cilj je
*jedan pisac koga zovu obe strane* (§26.5). Kad je cilj postao **brisanje svih
`frm*` fajlova**, ta pretpostavka je pala — i sa njom su se videle tri rupe koje
pod starim ciljem nisu bile rupe.

#### Šta jeste, a šta nije bilo premešteno

| | Stanje posle M5 |
|---|---|
| **Premešteno** (jedna implementacija, oba pozivaoca) | upis i sve provere (`modMaticniUnos`), GEO (`modMaticniGeo`), Korisnici i prava (`modMaticniKorisnici`), prevod Tag→sekcija |
| **Napravljeno paralelno** (dve implementacije) | **čitanje** — `LoadList` (355), `SetupColumnHeaders` (136), `lstData_Click` (159) i dalje samo u formi; nova ljuska ima `MatRedovi` / `MatKolone` / `MatVrednostiReda` |
| **Umire sa formom, i to je u redu** | 13 `Setup*`, stilizovanje, raspored kontrola, `m_RowMap`, vidljivost geo kontrola — to su ruke te forme |

Paralelno čitanje nije greška pod starim ciljem: pisac je deljen, pa se **upis**
ne može razići, a čitač je imao legacy kao referencu. Ali referenca nestaje sa
formom — pa se slaganje mora izmeriti **pre** brisanja, ne posle.

#### 1. Test slaganja čitača (161) — napisan sada

Tvrdnja iz §26.11 koja je kroz M1–M5 ostala nenapisana. Meri, nad svih 14 lista:
**skup zapisa** (broj, višak, manjak — kao *jednu* tvrdnju, jer su to tri lica
istog kvara) i **svaku golu vrednost** novog čitača protiv legacy reda istog
identiteta. Izvedene vrednosti se preskaču i to je **granica tvrdnje, ne rupa**:
`@status` novi čitač svodi na Aktivan/Neaktivan (parcele u šemi nose `"Da"`), a
`@geo` legacy nije ni pokazivao — zabeležene odluke iz §26.8.

Test i njegov seam (`frmStammdaten.StmTestLista`) su **privremeni**: umiru
zajedno sa formom. To je svrha, ne propust.

#### 2. Kaskada zavisnog combo-a — bila prekinuta

`MatComboStavke("@sorte", kontekst)` **jeste** bio prenet iz forme. **Ožičavanje
nije.** `OcistiPolja` je punio combo-e samo pri otvaranju editora, kad vrsta još
nema vrednost, pa je spisak sorti ostajao prazan — operater je kucao naslepo.
Događaj promene je stizao (`NewFieldG` → `WireInput` → `UiChange` →
`ScrAct "chg:"`), ali ga je `modMaticniEkran` ispuštao.

Popravka je zavisnost izmestila **u opis**, umesto tvrdo kodiranog
`If izvor = "@sorte"`: `MatComboZavisi(izvor)` je jedan spisak koji koriste tri
posla — punjenje pri otvaranju, ponovno punjenje na `chg:`, i punjenje posle
učitavanja postojećeg zapisa. Dok je zavisnost stajala u punjenju, druga dva
posla nisu ni postojala.

Uz to ide straža od povratnog udara (`mUZavisnima`): `FillCmb` i `PostaviPolje`
menjaju kontrolu, što opet okida `Change`. Kaskada je danas plitka pa bi se sama
zaustavila — straža je uslov da druga zavisnost sutra ne postane petlja.

#### 3. Podešavanja i Admin nemaju host bez forme — otvoreno

`modPodesavanja.BuildConfigEditor(frm)` i `modAdmin.BuildAdminPanel(frm)` grade
runtime kontrole **na `frmStammdaten`** (`Private mFrm As Object ' host`).
Brisanjem forme oba panela umiru.

To ne obara §26.19 — razlozi zašto **ne** postaju ekrani po ugovoru (97 polja,
`MAX_SEG` pun, `secret`/`memo` bez para u rečniku) i dalje stoje. Menja se samo
**gde žive**: treba im nov domaćin, a to je odluka koja nije doneta i nije bila u
planu. Zove se **M6** i predstoji.

### 26.21 M6 — Podešavanja i Admin u radnoj površini (`v6-ui-196`)

§26.19 je zaključio da ta dva **ostaju paneli** — 97 polja u 11 grupa i 12
sistemskih komandi ne staju u ugovor ekrana, i ne treba da stanu. §26.20 je
pokazao da ih hostuje `frmStammdaten`, koja se briše. M6 im daje nov dom, **bez
menjanja ijednog reda njihove logike**: `BuildConfigEditor` i `BuildAdminPanel`
su oduvek primali `frm As Object` — menja se samo **ko je domaćin**.

#### Ljuska daje prazan okvir i ništa više

`modOtkupUI` dobija `zPanel` — okvir preko radne površine — i dve javne
procedure: `PanelHost()` (okvir) i `PanelRezim(ukljuci)` (ustupanje). Ljuska
**ne zna nijedan panel po imenu**, isto kao što ne zna nijedan ekran. Dok je
površina ustupljena, `zScr_*`, `zTitle` i `zGrid` se ne crtaju; sidebar,
zaglavlje i statusna traka ostaju.

**Zašto okvir, a ne sama forma:** oba graditelja prvo sakriju **sve** kontrole
domaćina. Nad formom bi to ugasilo celu ljusku; nad namenskim okvirom gase samo
ono što je njihovo.

#### Logika je u standardnom modulu, ne u formi

`frmOtkupUI` nije dobio ništa. Registar panela je **`modUiPanel`** — nov
standardni modul, isti obrazac koji `modUiScreens` ima za ekrane: red opisuje
`KLJUC|modul|graditelj|naslov`, a poziv je isključivo **kasno vezan i
kvalifikovan** (`Application.Run "modPodesavanja.BuildConfigEditor", host`).

Zatvaranje ide obrnutim redom od otvaranja: prvo `<Modul>_Release` (omotači
`WithEvents`), pa tek onda pražnjenje okvira. Omotač koji preživi svoju kontrolu
je mrtva referenca koja puca tek pri sledećem kliku, daleko od uzroka. Ime
procedure se izvodi iz imena modula po dogovoru (`modAdmin` → `Admin_Release`),
pa nov panel ne traži red više u registru — samo da poštuje isti dogovor.
`modAdmin` je taj dogovor dobio sada; `modPodesavanja` ga je već imao.

#### Šta se promenilo u ponašanju

Panel se otvara **unutar ljuske, preko mreže** — sidebar i zaglavlje ostaju,
povratak je jedan klik. Klik u sidebaru dok je panel otvoren **prvo zatvara
panel**: bez toga bi se ekran ispod promenio, operater bi video isti panel, a
ljuska bi mislila da je negde drugde.

Skrivena kolona identiteta na ekranu *Podešavanja i alati* više ne nosi legacy
`Tag` za `frmStammdaten` nego **ključ panela**. Time nova ljuska nema više
nijednu vezu sa tom formom.

Izlaz iz panela je **host-svestan**: `CloseConfigEditor` / `CloseAdminPanel`
pitaju da li je domaćin okvir ili forma. Ta grana nestaje zajedno sa legacy
formom.

#### Brana je trostruka, i to nije višak

Ekran odgovara kroz `Scr_Dozvoljen`, registar proverava pre ustupanja površine,
a graditelj još jednom (`AUD-033`). Prava se menjaju zamenom operatera, pa
nijedan sloj ne veruje prethodnom. Provera u registru postoji da se površina ne
ustupi panelu koji će odmah odbiti da se izgradi — operater bi ostao pred
praznim okvirom.

**Verifikacija:** `vba_check` čist (208 fajlova, 357 sabotaža). Nov test 163 —
ugovor registra (modul, graditelj i naslov postoje; svaki modul poštuje dogovor
o `_Release`), ustupanje i vraćanje radne površine, i odbijanje nepoznatog
ključa bez uzimanja površine. Test 151 meri ključ panela umesto legacy `Tag`-a.
Tri nove sabotaže. **Testovi nisu izvršeni** u web sesiji.

**Posle M6 `frmStammdaten` više nema nijednog korisnika u novoj ljusci** — ostaje
samo legacy meni (`frmMaticniPodaci`) i test 161, koji je i napisan da umre s
njom.

### 26.22 „Procedure too large" — katalog poruka je prerastao VBA granicu

Prvi `Debug → Compile` posle spajanja sa `main`-om nije prošao:
**`Procedure too large`** na `modPoruke.UpsertPoruke`.

VBA odbija da prevede proceduru čija prevedena veličina prelazi **~64 KB**.
`UpsertPoruke` je bila jedna procedura sa **1657** `UpsertRow` poziva; obe grane
su nezavisno dopisivale poruke (Sledljivost, GP lanac, utovar sa jedne strane,
matični sa druge) i spajanjem je prag probijen. Nijedna strana sama nije bila
preko njega.

**Zašto to nije uhvaćeno ranije:** nije sintaksna greška, pa `vba_check` ćuti;
testovi je ne dotiču jer se modul ne prevede pa *ništa* ne radi; CI ne pokreće
Excel. Jedina kapija koja je vidi je ručni `Debug → Compile` — a on je poslednji
korak, ne prvi.

**Rez.** `UpsertPoruke` je podeljena na **12 blokova** (`UpsertPoruke01`…`12`),
a javna procedura zadržava izgradnju rečnika `existing` i redom ih zove. Podela
je **mehanička, po granici naredbe**: redosled poziva je nepromenjen, prelomljene
naredbe (`_`) nisu presečene, skup ključeva je dokazano identičan pre i posle.
Najveći blok je ~17 KB izvora.

**Kapija, da se ne ponovi.** `vba_check` je dobio `PROCEDURA_VELIKA`: procedura
preko **40 KB izvora** je nalaz. Prag je izmeren, ne pogođen — najveća procedura
u repou koja se uredno prevodi ima ~27 KB, nijedna nema preko 30 KB, a pala je
imala ~200 KB. Meri se izvor, jer se prevedena veličina odavde ne može izračunati;
mera je gruba namerno — posao provere je da javi „ova je prerasla", ne da predvidi
tačan bajt na kom VBA staje.

Dokaz u oba smera stoji u `vba_check --self-test` (`PROC_SIZE_CASES`), kako
`.claude/rules/testovi.md` §6 traži kad se menja sam checker: jedan slučaj koji
**mora** da zapišti i jedan veliki koji **ne sme**. Provereno i unazad — podignut
prag obara self-test po imenu.

### 26.23 Smoke posle integracije: decimale, radnja alatki, stil panela (`v6-ui-198`)

Tri nalaza sa prvog pravog prolaza kroz nove ekrane.

#### 1. Decimala je nestajala — i to na više mesta nego što je prijavljeno

Težina gajbice `0,5 kg` se crtala kao **`1`**. Vrsta ćelije `num` formatira sa
**nula decimala**, pa je `Format$` zaokruživao: `0,4 → "0"`, `0,5 → "1"`.

Podatak u tabeli je bio **ispravan sve vreme** — pisac upisuje pun `Double`.
Grešio je samo prikaz, što je gore od pada: broj izgleda kao podatak.

Isti kvar je pogađao i **cenu u Cenovniku** i cenu po jedinici artikla —
`152,50` se crtalo kao `153`. To niko nije prijavio, ali je ista greška.

Uvedena je vrsta ćelije **`dec`**: dve decimale, **bez zbira u podnožju**. Nije
upotrebljeno `rsd` iako ima dve decimale — novčane vrste pale podnožje sa
zbirom, a šifarnik nema šta da sabira (§26.4). Šest kolona prebačeno; polja
**unosa** ostaju `num`, jer provera već prima decimale i upisuje pun `Double`.

#### 2. Radnja „Otvori" nije se videla

Prigušenim (`soft`) dugmetom se nije videlo da je to **jedini ulaz** u alatku —
lista je izgledala kao pregled bez radnje. Sada je `primary` i šira. Uz to je
popravljen zastareo natpis kolone: od M6 se alatka otvara **u radnoj površini**,
ne „u svom prozoru".

#### 3. Paneli su i dalje izgledali kao stari ekrani

M6 ih je preselio u radnu površinu, ali su i dalje crtani **legacy `Style*`
helperima** — siva podloga, sistemski font, gusti redovi. To je bio pravi
prigovor: mesto je novo, izgled nije.

**Šta je promenjeno:** samo **prezentacija**. Registar polja
(`ConfigEditorFields`, 97 redova), čitanje, `SaveConfigEditor` i ožičavanje
(`clsConfigBtn`) su netaknuti.

**Šta NIJE promenjeno, i zašto:** kontrole ostaju `CommandButton` / `TextBox` /
`ComboBox`. `modUiKit.BtnV` pravi stilizovanu **ploču**, ne dugme, a
`clsConfigBtn` hvata evente preko `WithEvents btn As CommandButton` — zamena tipa
kontrole bi pokidala svaki klik u panelu. Zato se postojeće kontrole **boje**,
ne zamenjuju.

Stilovi stoje u **`modUiKit`** (`PanelStil*`), a ne u modulu panela: dva panela
sa svojom kopijom istog pravila su dva mesta koja se prvom doradom raziđu. Oba
panela sada dele paletu, tipografiju i ritam sa ostatkom ljuske; grupe u
Podešavanjima su peščane trake sa strelicom stanja, a u Adminu **`Očisti tabele`
i `Migracija` nose `danger` boju** — jedine dve komande koje diraju podatke, pa
se to vidi *pre* klika, a ne tek u dijalogu potvrde.

**Zapamćena zamka:** u MSForms se `SpecialEffect` i `BorderStyle` isključuju —
postavljanje `SpecialEffect`-a *posle* `BorderStyle`-a gasi ivicu. Redosled je
zato Flat → ivica → boja ivice.

**Verifikacija:** `vba_check` čist (210 fajlova, 417 sabotaža). Nov test 178 za
decimale (formatiranje, `dec` ne pali podnožje, nijedna matična kolona sa težinom
ili cenom nije `num`) i dve sabotaže. **Stil panela nije pokriven testom** — to je
vizuelni ishod koji se meri jedino smoke-om u Excelu; ne piše se placebo test
koji bi tvrdio da „boja je postavljena".
