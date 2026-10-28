# Migracija stare sveske — jedan alat

Zamenjuje lanac MIG-01 … MIG-22 sa grane `chore/import-migration-tools`
(182 fajla, 15 međufoldera sa SHA zaključavanjem). Ista pravila, jedan prolaz,
uvek od nule nad celom sveskom — nema delte.

| Fajl | Šta radi |
|---|---|
| `migracija.py` | komanda (`pregled`, `--self-test`) |
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
  kooperant → `ULAZ_TUDJE` / `VRACANJE_TUDJE` po dnevnom zaključku (MIG-16/17),
  kupac → `ULAZ_TUDJE` u trenutku manjka, stanica/vozač → jedna `NABAVKA` na
  prvi dan rada, u iznosu najdubljeg minusa. Sve dodato nosi `CreatedBy =
  MIGRACIJA`.
- **`verno`** — tačno kao stara sveska, sa negativnim saldima; list Salda tada
  predlaže događaj preseka iz salda na kraju.
- **Stavke** — svaki kooperant, stanica, vozač i kupac, stavka po stavka, sa
  tekućim novim i starim saldom. Filter na koloni *Naziv* ili *ID*.
- **Kretanja** — nacrt nove knjige u redosledu kolona `tblAmbalaza`.
- **Reversi** — nacrt zaglavlja `tblAmbalazaDokument`.
- **ZaOdluku**, **Upozorenja**, **Invarijante** — šta alat neće sam da reši.
- **Dokumenti**, **DokumentiViseRedova**, **Tabele** — ulaz za odluku 1.

Izvorna sveska se ne menja: SHA256 pre i posle čitanja mora biti isti, inače
alat javi grešku.

## Faza 2 — upis u novu svesku

Još ne postoji. Čeka odluke 1–5 (pre-flight: `main` je građen bez starih
podataka — `DOCUMENT_HEADER_LINES.md`, AMB-INV-07/08).
