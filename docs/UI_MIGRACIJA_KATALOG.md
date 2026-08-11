# Katalog funkcija legacy formi i plan prelaska na novi UI

> Izvor istine za migraciju `frmOtkup` i `frmDokumenta` na `frmOtkupUI`
> (ljuska `modOtkupUI` + ekranski moduli `modScr*`).
>
> Cilj dokumenta: **ništa se ne gubi u prelasku.** Svaka sposobnost legacy
> forme je ovde popisana, sa oznakom da li je u novom UI-ju već obezbeđena,
> delimično obezbeđena ili nije. Plan na kraju radi samo po ovom spisku.

Stanje na dan `v6-ui-101`.

---

## 1. Zajednička pravila (važe u obe forme)

Ovo nisu dugmad nego **ponašanja** koja se u legacy formama ponavljaju kroz
event-handlere. U novom UI-ju svako od njih mora imati tačno jedno mesto.

| # | Pravilo | Legacy mesto | Poslovna rutina | Novi UI |
|---|---|---|---|---|
| Z1 | **Cena iz cenovnika** po vrsti+sorti+klasi, na svaku promenu vrste/sorte | `AutoFillCenaOtkup`, `AutoFillCenaDok` | `modCenovnik.GetVazecaCena(vrsta, sorta, klasa)` | **IMA** (`AutoFillCena`, v6-ui-102) |
| Z2 | **Tip ambalaže iz kulture** po vrsti+sorti | isto | `GetKulturaTipAmbalaze` | **IMA** (v6-ui-102) |
| Z3 | **Predlog broja dokumenta** iz otkupnog mesta + datuma; poštuje toggle auto-broja | `RefreshBrojDokumentaSuggestion`, `RefreshBrojOtpSuggestion`, `RefreshBrojZbirneSuggestion`, `RefreshBrojPrijSuggestion`, `RefreshBrojReversSuggestion` | `modBrojevi.SuggestNextBroj(kind, entityID, datum)` | **DELIMIČNO** — samo pri izboru otpremnice (prefill); ne i na promenu OM/datuma |
| Z4 | **Živi zbir kg** uz polje količine; u bruto režimu prikazuje neto posle tare | `UpdateUkupnoKg`, `UpdateUkupnoKgOtp`, `UpdateUkupnoKgPrij` | `GetTezinaGajbice`, `OtkupBrutoUnos()` | **IMA** (`SetKgLine`, v6-ui-102) |
| Z5 | **Dve klase** — prekidač otvara drugi red količine/cene/ambalaže | `chkDveKlase*_Click`, `ShowKolAmbalazeII`, `ShowKlIIAmb` | — | **IMA** (segment I/II uz KLASA I CENA) |
| Z6 | **Parcele** — lista zavisi od kooperanta; celo polje gasi `PRACENJE_PARCELA` | `cmbKooperant_Change`, `ApplyOtkupTogglesState` | `IsPracenjeParcela()` | **IMA** |
| Z7 | **Keš isplate** — polja Novac/Primalac gasi `KES_ISPLATE` | `ApplyOtkupTogglesState`, `ApplyKesIsplateState` | `IsKesIsplate()` | **DELIMIČNO** — polja postoje u F5/F6, toggle se ne primenjuje |
| Z8 | **Blokada praznih polja** pri snimanju, gejtovana `VALIDACIJA_UNOSA` | `btnUnos*_Click` | `IsValidacijaUnosa()` | **DELIMIČNO** — `CommitDokument` proverava broj, količinu, cenu i datum; toggle se ne gleda |
| Z9 | **Info o paleti** — „još N gajbica do zatvaranja" uz izabranu robu | `UpdatePaletaInfo` | `GajbeDoZatvaranjaPaleteInfo` | **NEMA** |
| Z10 | **Prefill iz storniranog dokumenta** — ispravka posle storna | `PrefillOtkupFromStornirano`, `PrefillOtpremnicaFromStornirana`, `PrefillZbirnaFromStornirana`, `PrefillPrijemnicaFromStornirana` | — | **NEMA** |
| Z11 | **F-tasteri i Enter/Exit ivice** polja | `SetupFkeyAccelerators`, `HandleFkey`, `txt*_Enter/_Exit` | — | **IMA** (F1–F8 globalno, fokus ivice u `clsFlatBtn`) |
| Z12 | **KPI traka** iznad forme | `LayoutTopKpis`, `RefreshTopKpis`, `SumOtkupKgToday` | `GetOMAvansSaldo` | **DELIMIČNO** — traka postoji, dva KPI-ja nisu vezana |
| Z13 | **Podrazumevani proizvod** po otvaranju/resetu | `ResetProizvodNaDefault` | `ApplyDefaultProizvod` | **IMA** (`ApplyDefaultRoba`, v6-ui-103) |
| Z14 | **Kontekst datuma i otkupnog mesta** se pamti između dokumenata | `txtDatum_AfterUpdate`, `cmbOtkupnoMesto_Change`, `ResetDatumKontekst` | `GetActiveDatum`, `GetActiveStanica` | **NEMA** |

---

## 2. frmOtkup (1.294 linije) — otkupni list

Jedan režim, jedna forma. Sve što radi:

| Procedura | Šta radi | Poslovna rutina | Novi UI |
|---|---|---|---|
| `UserForm_Initialize` | puni combo-e (vrsta, sorta, tip ambalaže, vozači, kooperanti) | `GetLookupList`, `GetTipAmbalazeOptions`, `GetVozacDisplayList` | IMA (`FillCombos`) |
| `SetupAmbIzdataField`, `SetupKolAmbalazeIIField` | runtime polja koja nisu u `.frx` | — | IMA (cela forma je runtime) |
| `cmbVrstaVoca_Change` | puni sorte za vrstu | `GetLookupList` | IMA (`RefillSorta`) |
| `cmbSortaVoca_Change` → `AutoFillCenaOtkup` | **cena iz cenovnika + tip ambalaže iz kulture + paleta info** | Z1, Z2, Z9 | **NEMA** |
| `UpdateUkupnoKg` | živi zbir kg / neto iz bruta | Z4 | **NEMA** |
| `cmbOtkupnoMesto_Change` | pamti aktivnu stanicu, osvežava predlog broja | Z3, Z14 | DELIMIČNO |
| `txtDatum_AfterUpdate` | pamti aktivni datum, osvežava predlog broja | Z3, Z14 | **NEMA** |
| `cmbKooperant_Change` | puni parcele; osvežava ukupan iznos kooperanta u panelu blokova | Z6, `OtkupBlok_RefreshKoopTotal` | DELIMIČNO — parcele ima, ukupan iznos nema |
| `cmbParcela_Change`, `ExtractParcelaID` | ID parcele iz prikaza | — | IMA |
| `chkDveKlase_Click` | druga klasa | Z5 | IMA |
| `ApplyOtkupTogglesState` | parcele / keš isplate | Z6, Z7 | DELIMIČNO |
| `btnUnos_Click` | **snimanje otkupa**; relink paleta hladnjače ako je bio storno | `SaveOtkupMulti_TX`, `ReassignPaleteToPrijemnica_TX`, `GetHladnjacaRelinkPending` | **NEMA** (`CommitDokument` je šav) |
| `ClearOtkupFields` | reset forme posle snimanja | — | IMA (`ClearForm`) |
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
| Otpremnica | `btnUnosOtp_Click` | `SaveOtpremnicaMulti_TX` | F2 — forma IMA, upis NEMA |
| Zbirna | `btnUnosZbr_Click` | `SaveZbirnaMulti_TX` + `ValidateZbirnaPreUnosa` | F3 — forma IMA, upis i validacija NEMA |
| Prijemnica | `btnUnosPrij_Click` | `SavePrijemnicaMulti_TX`, `GetPaletaStatusForPrijemnica`, `ReassignPaleteToPrijemnica_TX` | F4 — forma IMA, upis NEMA |
| OM ulaz (revers/avans) | `btnUnosOMUlaz_Click` | `SaveOMUlaz_TX`, `GetOMAvansSaldo`, `GetOpenOtkupi` | F7 — forma IMA, upis NEMA |
| Kupci izlaz (uplate) | `btnUnosIzlaz_Click` | `SaveKupciIzlaz_TX`, `GetUplataForFaktura`, `GetOpenFakture` | F6 — forma IMA, upis NEMA |

Uz njih: `LoadZbirneListbox` / `lstZbirne_Click` (izbor zbirne iz liste),
`UpdateManjak` (manjak prijemnice vs zbirna), `UpdateValidacija`,
`UpdateOMAvansSaldo`, `FillOpenFakture`, `FillOpenOtkupi`,
`SetupOMIzdavanjeToggle` (četiri smera reversa — **IMA** u F7).

### 3.2 Storno okvir — najveći deo forme

Sedam panela, svaki sa svojim `Ensure/Layout/Populate/Set*Visible`:

| Panel | Šta radi | Poslovna rutina | Novi UI |
|---|---|---|---|
| **Storno** (`btnStorno_Click`) | storno bilo kog tipa dokumenta po broju | `StornoOtkup_TX`, `StornoOtpremnica_TX`, `StornoZbirna_TX`, `StornoPrijemnica_TX`, `StornoFaktura_TX`, `StornoNovac_TX`, `StornoIzvod_TX`, `StornoOMKoopByBrDok_TX` | DELIMIČNO — F8 lista postoji, storno radi samo nad otkupom |
| **Ispravka / dupli unos** (`TryRunCorrectionFramework`, `PromptCorrectionMode`) | posle storna nudi ispravku (prefill) ili dupli unos | `GetCorrectionField`, `ApplySelectedBlockStorno`, `StornoSelectedBlocks_TX` | **NEMA** |
| **Storno pregled** (`m_btnStornoPregled_Click`) | pregled storniranih, grupisano | `GetStorniraniGrupisano` | DELIMIČNO — F8 lista |
| **Undo operacija** (`ShowUndoOpsPanel`) | poništavanje storna | `GetUndoableStornoOperations` | **NEMA** |
| **Nađi dokument** (`m_btnStornoFind_Click`) | pretraga „toplih" dokumenata po tipu i tekstu | `GetWarmStornoDocs` | DELIMIČNO — brza pretraga u mreži |
| **Nedovršeno** (`m_btnNedovrseno_Click`) | lanci koji nisu dovršeni | `GetNedovrseno` | **NEMA** |
| **Recovery** (`m_btnRecovery_Click`) | osirotele prijemnice i palete, prevezivanje | `GetOsirocenePrijemnice`, `GetPrijemniceSaOsirocenimPaletama`, `ReassignPaleteToPrijemnica_TX`, `ReassignPrijemnicaToZbirna_TX` | **NEMA** |
| `CheckVerwaisteDokumente` | upozorenje na siročiće pri otvaranju | `GetVerwaisteDokumente` | **NEMA** |

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
3. **Z9**: „još N gajbica do zatvaranja palete" uz izabranu robu.
4. **Z3 + Z14**: predlog broja na promenu OM/datuma; pamćenje aktivne stanice i
   datuma.
5. **Z7 + Z8**: toggle-i (keš isplate, validacija unosa).
   ~~Z13 podrazumevani proizvod~~ **URAĐENO** (v6-ui-103)

### Faza B — upis (`CommitDokument`)
6. F1 → `SaveOtkupMulti_TX` (+ relink paleta hladnjače).
7. F2/F3/F4 → `SaveOtpremnicaMulti_TX`, `SaveZbirnaMulti_TX` (+
   `ValidateZbirnaPreUnosa`), `SavePrijemnicaMulti_TX` (+ status palete).
8. F5/F6/F7 → `SaveOMUlaz_TX`, `SaveKupciIzlaz_TX`, novac.
9. Z12: preostali KPI-jevi (`GetOMAvansSaldo`, otvoreno kg).

### Faza C — Palete P2
10. Unos prerade (traži prosleđivanje događaja sopstvenih kontrola ekranu).
11. Čipovi po ekranu (godina/status/prerađeno) — generalizacija čipova.

### Faza D — storno okvir (najveći, ide poslednji)
12. Storno svih tipova dokumenata iz F8.
13. Ispravka / dupli unos posle storna (Z10) + „hladnjača ispravka".
14. Recovery, Nedovršeno, Undo operacija, Nađi dokument.

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
