# Maticni podaci: sifarnici + cenovnik (dorade)

Ovaj dokument opisuje dorade na desktop (Excel/VBA) aplikaciji vezane za
maticne podatke, sifarnike i cenovnik.

## Sta je uradjeno

1. **Sifarnici preko "Maticni podaci"** — `frmMaticniPodaci` sada nudi i:
   - **Kulture** (`tblKulture`)
   - **Ambalaza** (`tblTipAmbalaze`)
   - **Palete** (`tblTipPalete`)
   - **Cenovnik** (`tblCenovnik`, nova tabela)

   Ceo meni je sada **data-driven**: gradi se iz jedne registracije sekcija
   (`modMaticniLookups.MaticniSekcije`), a dugmad se prave dinamicki
   (`Controls.Add` + `clsLookupMenuBtn` WithEvents), tako da se
   `frmMaticniPodaci.frx` **ne dira**. Postojeca staticna dugmad ostaju kao
   fallback (sakrivena pri uspesnoj dinamickoj izgradnji).

   Unos/izmena svih sekcija ide kroz postojeci univerzalni editor
   `frmStammdaten` (nove `Case` grane).

2. **Izmena za "Stanice" popravljena** — `frmStammdaten` je pri izmeni
   stanice pisao u nepostojece kolone `KontaktIme`/`KontaktPrezime`, pa je
   izmena padala (rollback). Sada koristi prave kolone **`Ime` / `Prezime` /
   `PIN`** (kako ih vodi i sync sloj). Dodavanje je i ranije radilo jer je
   pozicijsko.

3. **Tip ambalaze iz maticnih podataka** — u `frmOtkup` i `frmDokumenta`
   combo "Tip ambalaze" se sada puni iz `tblTipAmbalaze`
   (`GetTipAmbalazeOptions`), uz fallback na 12/1 i 6/1 ako je sifarnik
   prazan.

4. **Cenovnik (cena po proizvodu, append-only)** — nova tabela `tblCenovnik`.
   Vazeca cena = poslednji (najnoviji po `Datum`) ne-stornirani red za
   kljuc **VrstaVoca + SortaVoca + Klasa**. Stari redovi ostaju (kretanje
   cena). U `frmOtkup` i `frmDokumenta` cena se **automatski popunjava** pri
   izboru vrste/sorte (Klasa I -> polje cene, Klasa II -> polje cene II);
   rucni unos i dalje moguc.

## Sema tblCenovnik

| Kolona      | Tip     | Opis                                  |
|-------------|---------|---------------------------------------|
| CenaID      | tekst   | PK, "CEN-00001"                       |
| Datum       | datum   | datum vazenja cene                    |
| VrstaVoca   | tekst   | npr. "Malina"                         |
| SortaVoca   | tekst   | npr. "Willamette" (moze i prazno)     |
| Klasa       | tekst   | "I" ili "II" (KLASA_I / KLASA_II)     |
| Cena        | broj    | cena po jedinici                      |
| CreatedAt   | datum   | timestamp unosa                       |
| Stornirano  | tekst   | "Da" iskljucuje red                   |

Konstante: `TBL_CENOVNIK`, `COL_CEN_*` u `modConfig.bas`.
Logika: `modCenovnik.GetVazecaCena` / `modCenovnik.AddCena`.

## Setup (jednokratno, na master workbook-u)

`tblCenovnik` se kreira automatski u sklopu `EnsurePaletniListSchema`, ili
samostalno:

```
Alt+F8 -> EnsureCenovnikSchema
```

Idempotentno je (kreira tabelu ili dopuni kolone).

> Napomena: `tblKulture` se pri dodavanju iz forme pise pozicijski i
> ocekuje kolone redom: `KulturaID, VrstaVoca, SortaVoca, Aktivan,
> GajbicaPoPaleti`. Izmena ide po imenu kolone (bezbedno).

## VAZNO — PWA sinhronizacija cenovnika (TODO)

Trenutno cenovnik radi **samo u desktop (Excel/VBA) aplikaciji**.

**Kada krene rad na PWA / mobilnom otkupu, OBAVEZNO sinhronizovati
`tblCenovnik` na Google Sheets** (kroz `modStammdatenSync` /
`modMasterSync`), da bi otkupci na stanicama dobijali istu vazecu cenu kao
u desktopu. Bez toga, auto-cena na mobilnom nece postojati ili ce se
razlikovati od desktopa.

Predlog obima za PWA fazu:
- Export `tblCenovnik` u poseban sheet (npr. "Cenovnik") — analogno
  `ExportStanice` / `ExportKupci`.
- Na PWA strani: ucitati cenovnik i primeniti istu logiku "poslednji red
  vazi" po kljucu VrstaVoca + SortaVoca + Klasa.

## Rucni test (u Excelu, posle importa modula)

1. **Setup:** pokrenuti `EnsureCenovnikSchema` (ili `EnsurePaletniListSchema`).
2. **Maticni podaci -> Ambalaza:** dodaj tip (npr. "12/1", 0.9 kg), izmeni,
   proveri da se cuva.
3. **Maticni podaci -> Palete:** dodaj tip palete + tezina.
4. **Maticni podaci -> Kulture:** dodaj vrstu/sortu (+ gajbica/paleti).
5. **Maticni podaci -> Stanice:** izaberi stanicu, izmeni Ime/Prezime/PIN ->
   mora da se sacuva (ranije nije radilo).
6. **Maticni podaci -> Cenovnik:** dodaj cenu (Vrsta+Sorta+Klasa I),
   pa jos jednu sa drugom cenom -> u listi se vide oba reda (najnoviji gore).
7. **frmOtkup:** izaberi tu vrstu+sortu -> polje "Cena" se auto-popuni
   poslednjom cenom. "Tip ambalaze" lista = ono iz `tblTipAmbalaze`.
8. **frmDokumenta:** isto za otpremnicu/prijemnicu (Cena I/II).
