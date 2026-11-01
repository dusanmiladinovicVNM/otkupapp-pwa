# Migracija stare sveske — jedan alat

Zamenjuje lanac MIG-01 … MIG-22 sa grane `chore/import-migration-tools`
(182 fajla, 15 međufoldera sa SHA zaključavanjem). Ista pravila, jedan prolaz,
uvek od nule nad celom sveskom — nema delte.

| Fajl | Šta radi |
|---|---|
| `migracija.py` | komande `pregled`, `proba`, `upis`, `--self-test` |
| `citac.py` | čita Excel tabele iz `.xlsm` bez Excela, samo za čitanje |
| `ambalaza.py` | stara knjiga (noge Ulaz/Izlaz) → AMB-10 kretanja, dug kooperanta, salda |
| `dokumenti.py` | tabele naspram `schema/schema.json`, dokumenti sa više redova |
| `pregled.py` | Excel pregled (prave Excel tabele sa filterom) |
| `izuzeci.json` | ručne odluke po starom `AmbID`, svaka sa potpisom reda |
| `test_migracija.py` | testovi nad izmišljenom sveskom (CI) |

## Faza 1 — pregled (samo čita)

Iz root-a repoa, u Git Bash-u:

```bash
python tools/migracija/migracija.py pregled --izvor "C:/Users/Dusan/Desktop/AgriX - C002/Venivo/Original kraj sezone pred konacno stanje/AgriX_C002_2.28.4_2026-10-16.xlsm"
```

Pregled ide u `Migracija_pregled/` pored izvorne sveske (nikad u repo — nosi
stvarne podatke). Listovi:

- **Sazetak** — koliko redova, koji ishod, šta je za odluku, provere.
- **Salda** — svaki nalog × tip ambalaže: stari i novi saldo, dopisano pokriće
  tuđe ambalaže i nabavka, obaveza firme, najmanji dnevni saldo.

Dva režima knjiženja (`--knjizenje`):

- **`ispravno`** (podrazumevano) — istorija + ono što `main` traži za dalji rad
  (nijedan realan nalog ni jednog dana ispod nule, AMB-INV-07):
  kooperant → `ULAZ_TUDJE` / `VRACANJE_TUDJE` + `IZLAZ_TUDJE` po dnevnom
  zaključku (MIG-16/17, AMB-INV-15), kupac → `ULAZ_TUDJE` u trenutku manjka,
  stanica/vozač → jedna `NABAVKA` na prvi dan rada, u iznosu najdubljeg minusa
  ukupnog salda ili praznih (AMB-INV-16). Knjiga robnog dokumenta je ono što bi
  upisao pisac main-a: otpremnica = zbir gajbi blokova i samo kad je izdata,
  zbirna = zbir otpremnica, a gajbe otkupa i prijemnice prate staru knjigu
  (MIG-ODL-13..15). Sve dodato nosi `CreatedBy = MIGRACIJA`.
- **`verno`** — tačno kao stara sveska, sa negativnim saldima; list Salda tada
  predlaže događaj preseka iz salda na kraju.
- **Stavke** — svaki kooperant, stanica, vozač i kupac, stavka po stavka, sa
  tekućim novim i starim saldom. Filter na koloni *Naziv* ili *ID*.
- **Kretanja** — nacrt nove knjige u redosledu kolona `tblAmbalaza`.
- **Reversi** — nacrt zaglavlja `tblAmbalazaDokument`.
- **Uskladjivanje** — svaki red knjige koji je odredio dokument (stara i nova
  količina, razlog).
- **ZaOdluku**, **Upozorenja**, **Invarijante** — šta alat neće sam da reši.
- **Dokumenti**, **DokumentiViseRedova**, **Tabele** — ulaz za odluku 1.

Izvorna sveska se ne menja: SHA256 pre i posle čitanja mora biti isti, inače
alat javi grešku.

## Faza 2 — nova sveska

Pravila i odluke: `docs/DOMEN/MIGRACIJA_LEGACY.md`.

**Korak 1 — proba (samo čita):**

```bash
python tools/migracija/migracija.py proba --izvor "C:/.../stara.xlsm"
```

Pravi `proba_upisa.xlsx`: svaka tabela nove sveske kao Excel tabela + list
**Provere**. `BLOKIRA` = posle upisa bi main pao na čitanju cele tabele;
`ODLUKA` = ishod zavisi od tebe; `UPOZORENJE` = akcija nad tim jednim
dokumentom ne bi prošla.

**Korak 2 — upis (Windows + Excel + pywin32):**

```bash
python tools/migracija/migracija.py upis --izvor "C:/.../stara.xlsm" --cilj "C:/.../nova.xlsm" --prihvatam-odluke
```

Ako `--cilj` ne postoji, pravi ga `tools/make_dev_workbook.py` iz gita (treba
„Trust access to the VBA project object model”). Upis odbija probu sa BLOKIRA
nalazom, piše samo u prazne tabele, pa svaku ćeliju poredi sa planom i salda
ambalaže sa prevodom. Izveštaj: `<nova>_upis_izvestaj.txt`.

Posle svake tabele, pre čuvanja, prvi i poslednji red se čitaju nazad iz
Excela: pomeren datum ili tekst koji je Excel pretvorio zaustavlja upis dok je
sveska još prazna (nije sačuvana — samo ponovi komandu). Ako izveštaj posle
čuvanja prijavi razlike, sveska je sačuvana sa greškom: obriši je i ponovi
komandu (alat pravi svežu iz gita).

| Fajl | Šta radi |
|---|---|
| `cilj.py` | sve tabele nove sveske u memoriji (zaglavlje + stavka, izvori, statusi, podešavanja) |
| `provere.py` | ista pravila po kojima main čita — pre upisa |
| `upis.py` | jedini deo koji piše (COM), plus poređenje posle upisa |
