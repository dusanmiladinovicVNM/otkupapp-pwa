# Katalog funkcija legacy formi i plan prelaska na novi UI

> Izvor istine za migraciju `frmOtkup` i `frmDokumenta` na `frmOtkupUI`
> (ljuska `modOtkupUI` + ekranski moduli `modScr*`).
>
> Cilj dokumenta: **ništa se ne gubi u prelasku.** Svaka sposobnost legacy
> forme je ovde popisana, sa oznakom da li je u novom UI-ju već obezbeđena,
> delimično obezbeđena ili nije. Plan na kraju radi samo po ovom spisku.

Stanje na dan `v6-ui-117`.

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
2. **Storno okvir frmDokumenta** — sedam panela i ~2/3 te forme. Novi UI danas
   ume da stornira **samo otkupni list** (iz F1 liste); otpremnica, zbirna,
   prijemnica, faktura, novac i izvod — ne.
3. **Pomoćni delovi režima** koji nisu pravila nego zaseban posao: lista
   zbirnih za izbor (F3), manjak prijemnice vs zbirna (F4). Upis F3/F4 od
   v6-ui-116 radi i bez njih — to su prikazi, ne kapije. Ostatak te tačke je
   zatvoren u v6-ui-117: avans saldo OM, otvorene fakture (F6) i otvoreni
   otkupi (F5) sada postoje, jer bez njih upis ne bi bio tačan (v. §3.1).
4. **Sitno:** peščanik za vreme upisa, dva nevezana KPI-ja, prefill iz
   storniranog (Z10). Filtriranje kooperanata po otkupnom mestu je urađeno
   (v6-ui-113, `KOOP_FILTER_BY_OM`).

Tačka 2 je Faza D iz plana; 3 i 4 su ostaci.

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
| Z10 | **Prefill iz storniranog dokumenta** — ispravka posle storna | `PrefillOtkupFromStornirano`, `PrefillOtpremnicaFromStornirana`, `PrefillZbirnaFromStornirana`, `PrefillPrijemnicaFromStornirana` | — | **NEMA** |
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
| `Prefill*FromStornirana` | ispravka posle storna | **NEMA** — Faza D |
| storno paneli (7 kom.) | `btnStorno_Click` i dalje | **NEMA** — Faza D |

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
| **Prijemnica: ispravka posle storna (relink paleta)** | **NIJE preneto** | traži `SetPaletizeSkip` **pre** upisa + `ReassignPaleteToPrijemnica_TX` + `PaletaAdjustPrompt` — to je storno okvir (Faza D). Novi UI još ne ume da stornira prijemnicu, pa takva ispravka može da nastane samo u `frmDokumenta`, gde se i završava. Prijemnica upisana iz novog UI-ja dok takva ispravka visi dobija **sveže** palete, a stare ostaju osirotele. |

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
