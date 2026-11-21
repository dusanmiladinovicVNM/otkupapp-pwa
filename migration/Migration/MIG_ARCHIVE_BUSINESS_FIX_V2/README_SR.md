# MIG-ARCHIVE — poslovna kartica kooperanta, patch v2 (stvarni VBA snapshot)

## Polazna osnova

Originalni izvoz stvarnih VBA modula iz `ACTUAL_VBA_SNAPSHOT_RECOVERED_20261023_175339_339025.zip`, iz korisnikove `_05` arhivske sveske. Ne koristi se GitHub baseline. `modIzvestaj` zadrzava postojece MIG-18 dorade (`AmbKarticaSortKey`, opisi, sortiranje). `modScrIzvestaji` i `modAmbalazaUgovor` ne menjaju se.

## Zasto nastaje razlika

`AmbKretanjaNaloga` trenutno vraca i poslovne prenose i automatsko pokrice deficita `ULAZ_TUDJE_AMBALAZE`. `ReportKarticaAmbalaze` tumaci svaki pozitivan potpisani iznos kao ulaz kooperanta. Primer `KOOP-00019`: UI prikazuje konacnih +5; tehnicka pokrica 8.7. (+12) i 11.7. (+3) ne pripadaju poslovnom prometu. Ocekivani rezultat nakon izuzimanja 15: **-10**. Stvarni transferi 23 i 3 ostaju kao izlazi. FIZICKA knjiga i njen saldo ostaju neizmenjeni.

## Sta patch radi

- `modAmbalaza.AmbKretanjaNaloga` dobija `Optional ByVal samoPoslovno As Boolean = False`. Svi stari pozivaoci ostaju identicni, a izbor `True` filtrira samo `ULAZ_TUDJE_AMBALAZE` i njegov `StornoOd` inverz. Koristi stvarne `VrstaKretanja` i mapu originalnih `AmbID`-jeva koju proverava `KnjigaZaCitanje`.
- `modIzvestaj.ReportKarticaAmbalaze` i `ReportKarticaKooperanta` traze poslovni rezim `True`.
- Ne menja `AmbSaldoNaloga`, `AmbSaldoPoNalogu`, `GetAmbalazeStanje`, `AmbObavezaPartneru`, poslovne podatke, `tblAmbalaza`, obrazac za upis, `modScrIzvestaji` ni PDF renderer. Obe stampe koriste te iste postojece funkcije izvestaja.
- 7 preciznih izmena VBA koda, 2 modula, bez zamene kompletnog modula. Snapshot i patch preview su u paketu.

## Pokretanje (Git Bash, Windows)

1. Raspakuj ceo ovaj ZIP u `C:\Users\Dusan\Desktop\AgriX DEV\Migration\MIG_ARCHIVE_BUSINESS_FIX_V2\`.
2. Otvori samo `C:\Users\Dusan\Desktop\AgriX DEV\MigrationArchiveXLSM_05\AgriX_2026_ARHIVA_BUSINESS_VIEW.xlsm` u Excelu. Zatvori druge Excel instance ako su otvorene. Ne menjaj istorijske podatke. U Trust Center ukljuci privremeno `Trust access to the VBA project object model`.
3. U Git Bash-u:

```bash
ROOT="$HOME/Desktop/AgriX DEV"
PS1="$ROOT/Migration/MIG_ARCHIVE_BUSINESS_FIX_V2/install_open_workbook.ps1"
BOOK="$ROOT/MigrationArchiveXLSM_05/AgriX_2026_ARHIVA_BUSINESS_VIEW.xlsm"

powershell.exe -NoProfile -ExecutionPolicy Bypass \
  -File "$(cygpath -w "$PS1")" -Workbook "$(cygpath -w "$BOOK")" -DryRun
```

Samo ako terminal prijavi `PREFLIGHT_PASS` + `DRY_RUN_OK`, ponovi bez `-DryRun`:

```bash
powershell.exe -NoProfile -ExecutionPolicy Bypass \
  -File "$(cygpath -w "$PS1")" -Workbook "$(cygpath -w "$BOOK")"
```

Ocekivano `BACKUP_SAVED`, `VERIFIED modAmbalaza`, `VERIFIED modIzvestaj`, `ARCHIVE_BUSINESS_CARD_PATCH_SAVED`.

4. U VBA Editor-u `Debug -> Compile VBAProject`. Osvezi karticu KOOP-00019 za `01.01.2026-23.10.2026` i proveri:
   - 8.7.: ne postoji fiktivni ulaz 12; izlaz 23 ostaje.
   - 11.7.: ne postoji fiktivni ulaz 3; izlaz 3 ostaje.
   - Poslednji saldo na ambalaznoj kartici i UI KPI: **-10** (prethodno +5).
   - Finansijska kartica (saldo ambalaze), stampa/PDF: isti poslovni rezultat.
   - Proveri jos jedan drugog kooperanta i normalna izdavanja/povrate.
5. Iskljuci `Trust access to the VBA project object model` posle provere.

## Stroge zastite

- Instalater nikada ne pokrece novu Excel COM instancu; koristi vec otvorenu.
- Pre izmene **uporedjuje kompletan stvarni kod** oba modula sa source snapshotom koji si dostavio; bilo kakva razlika znaci STOP.
- Radi iskljucivo nad `_05/AgriX_2026_ARHIVA_BUSINESS_VIEW.xlsm`, nikad masterom ili `_04`.
- Pre izmene eksportuje `.bas` i `.code.txt` kao backup na disk.
- `-DryRun` ne menja VBA i ne snima XLSM.
- Ako se prijavi `SOURCE_MISMATCH`, nemoj pokretati Apply i nemoj prepisivati module; prvo novi read-only snapshot.
- Ako padne tokom apply, nemoj snimati izmene i zatvori `_05` **bez Save**, jer bi delom primenjen patch inace ostao u svesci. Kopija `_04` postoji.
- Zahteva da otvorena sveska nema prethodno nesacuvane izmene pre Apply.

## Status testiranja

Python source-level i simulacioni testovi: 4/4 PASS. Stvarna VBA compile/Windows COM instalacija i stvarni ekran nisu testirani u ovom okruzenju. Rezultat `PATCH_SAVED` i Excel test su obavezni za finalni GO.
