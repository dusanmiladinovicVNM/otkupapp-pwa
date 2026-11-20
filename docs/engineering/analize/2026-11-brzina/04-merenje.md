# Merenje brzine pre i posle — `modPerf`

Uputstvo za merenje dobiti iz rezova u PR #497 (IZV-1, OTK-1, HLD-1,
IZV-3/IZV-5, IZV-2, IZV-9, IZV-7). Sva vremena u `README.md` su **procene**;
ovo je način da se zamene merenjem.

## 1. Dve grane, isti merni kod

| Grana | Šta je | Za šta |
|---|---|---|
| `claude/merenje-modperf-pre` | `main` + `modPerf` + kukice. Bez testova merača. | merenje **pre** |
| `claude/merenje-modperf` | PR #497 (optimizacije) + `modPerf` + kukice + testovi 393–399 | merenje **posle**; kasnije zaseban PR |

Kukice stoje samo na linijama koje optimizacije ne diraju, pa obe grane mere
**istim** kodom. Razlika u izmerenom vremenu je razlika u kodu aplikacije, a ne
u meraču.

## 2. Šta merač meri

Spisak pozivnih mesta je u zaglavlju `src-vba/modPerf.bas`.

- **Akcija:** svaka radnja operatera (`UI.<događaj>.<tag>`, npr.
  `UI.Click.scrIzRezZ`), plus `ONTIME.AUTOSAVE` i `ONTIME.WARM`.
  - Programski `Change` unutar radnje postaje sekcija `AKCIJA.UI.Change.*`.
    Tako se vidi i kaskada iz nalaza UI-7, ako postoji.
  - `REFRESH` je `RefreshFromData`.
- **Sekcije:**
  - posle „Sačuvaj": `UI.Commit.ScrSave`, `UI.Commit.ClearForm`;
  - štampa: `PRINT.OtkupniList`;
  - hladnjački lanac: `OTK.AutoLanacHladnjaca`, `DOK.NevezaneOtpremnice`;
  - jedan snimak Izveštaja: `IZV.PuniSnimak|<lista>|<P/Z>`;
  - snimak tabele pre upisa: `TX.Snapshot|<tabela>`.
- **Čekanje na dijalog** (`UI.CekaOperatera`) se oduzima od aktivnog vremena.
- **Brojači:**
  - `DA.GTD|<tabela>`: svako čitanje tabele sa lista, uz broj ćelija;
  - `DA.GTD.KES|<tabela>`: pogodak keša.

  Brojači su deterministički, pa tvrdnju „tabela se čita 8 puta po snimanju"
  potvrđuju ili obaraju tačnim brojem, nezavisno od šuma.
- **Sat:** `QueryPerformanceCounter` (test 398); kad nije dostupan, rezerva je
  `Timer`.
- **Merač ništa ne piše u tabele sveske.**
  - Izlaz je `<folder sveske>/Log/Perf_YYYY-MM-DD.csv`, separator `;`,
    decimalna tačka.
  - Uz to ide po jedan red u `OtkupApp_*.log` po akciji iznad praga.

## 3. Priprema (jednom)

1. **Samo kopija prave sveske, nikad produkciona.**

   ```bash
   mkdir -p /c/Merenje/pre /c/Merenje/posle
   ```

   ```bash
   cp "/c/<folder>/<sveska>.xlsm" /c/Merenje/pre/sveska.xlsm
   ```

   ```bash
   cp "/c/<folder>/<sveska>.xlsm" /c/Merenje/posle/sveska.xlsm
   ```

2. **U obe kopije ugasi spoljni saobraćaj**, inače kopija može da gura podatke
   u produkciju:
   - Podešavanja → Sinhronizacija → „Cloud sync uključen" = NE
     (`CLOUD_SYNC_ENABLED` = `NO`);
   - `MONITORING_ENDPOINT` isprazniti (bez njega se ništa ne šalje).
3. **Uvezi kod:** `pre` kopija dobija granu `claude/merenje-modperf-pre`, a
   `posle` kopija granu `claude/merenje-modperf`. Obe ulaze kroz uobičajeni
   `ImportAllVBA`, pa `Debug → Compile VBAProject`.
4. **Uključi merač u obe kopije** (Immediate, `Ctrl+G`):

   ```
   SetLocalConfigValue "PERF_MERENJE", "DA"
   SetLocalConfigValue "PERF_PRAG_MS", "0"
   ```

   - Prag 0 upisuje svaku akciju; podrazumevano je 10 ms.
   - Za jednu sesiju bez upisa u konfiguraciju: `Alt+F8 → Perf_Ukljuci`.
   - Nivo 2 (`PERF_MERENJE` = `DETALJ`) meri i vreme svakog čitanja tabele, ali
     je skuplji; koristi se samo za raspodelu.

## 4. Scenariji

- **Ulaz:** isti u obe kopije (isti kooperant, OM, količine, gajbe). Sveska
  se pre svakog prolaza vraća na polaznu kopiju.
- **Ponavljanja:** svaka radnja 1× „hladno" (prva posle otvaranja sveske) i 5–10×
  „toplo".

| # | Scenario | Rez koji se meri |
|---|---|---|
| S1 | Izveštaji → Saldo, OM: Zbirno, pa Pojedinačno, pa promena OM-a i povratak | IZV-1, IZV-2 |
| S2 | Izveštaji → zbirni saldo i roba po OM i po kupcima | IZV-3, IZV-5 |
| S3 | Izveštaji → Kooperant → RANG, period prazan | IZV-7 |
| S4 | Izveštaji → otkupni listovi jednog OM-a: klik na čip vrste, pa „Sve" | IZV-9 |
| S5 | Snimanje otkupa (sa PDF-om) na običnoj stanici | OTK-1 |
| S6 | Snimanje otkupa na hladnjačkoj stanici | HLD-1, lanac |
| S7 | Posle jednog snimanja 70 s mirovanja | AutoSave i warm |

Na kraju svakog prolaza: `Alt+F8 → Perf_Sazetak` napravi
`Log/Perf_sazetak_*.txt` (po akciji: broj ponavljanja, prvo merenje, prosek
ostalih, min i max; top sekcije; brojači; veličine tabela).

## 5. Čitanje (Git Bash)

```bash
cd /c/Merenje/pre/Log
```

Aktivno vreme po akciji, toplo (prosek, broj, akcija):

```bash
awk -F';' '$6=="AKCIJA" && $5>1 {n[$4]++; s[$4]+=$10} END {for (a in n) printf "%10.1f %4d %s\n", s[a]/n[a], n[a], a}' Perf_*.csv | sort -rn
```

Čitanja tabela sa lista po akciji (zbir ćelija i broj čitanja):

```bash
awk -F';' '$6=="BROJ" && $7 ~ /^DA\.GTD\|/ {c[$4" "$7]+=$12; p[$4" "$7]+=$8} END {for (k in c) printf "%12.0f %6d %s\n", c[k], p[k], k}' Perf_*.csv | sort -rn | head -30
```

Isto u `/c/Merenje/posle/Log`, pa se uporede iste akcije.

## 6. Šta je prihvatljiv dokaz

- **Dobit postoji** kad je medijana toplih ponavljanja „posle" manja od „pre"
  za više od rasipanja između ponavljanja iste akcije, **i** kad brojač pada
  koliko rez tvrdi. Primeri:
  - IZV-1: povratak na viđen kontekst bez novog `IZV.PuniSnimak`;
  - IZV-3: `DA.GTD|tblOtkupStavke` po zbirnom prikazu manje.
- **META** redovi moraju biti uredni: `TableCacheDepth` = 0, a
  `Nezatvorenih`/`Neuparenih` = 0 ili objašnjeni.
- **Neočekivan nalaz se beleži, ne prepravlja.** Primer je sekcija
  `AKCIJA.UI.Change.*` unutar „Sačuvaj" (UI-7).
