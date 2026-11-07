# Provera unosa stare sveske

Alat **samo čita** staru svesku i prijavljuje greške koje je stari unos ostavio u
podacima. Ne menja ništa: SHA256 sveske pre i posle čitanja mora biti isti.

Iz root-a repoa, u Git Bash-u:

```bash
python tools/migracija/provera_unosa.py "C:/Users/Dusan/Desktop/AgriX - C002/Venivo/Original kraj sezone pred konacno stanje/AgriX_C002_2.28.4_2026-10-16.xlsm"
```

Izveštaj ide u folder pored sveske: `Migracija_pregled/<ime>_provera_unosa.xlsx`
(nikad u repo — nosi stvarne podatke). Listovi:

- **Sazetak** — koliko nalaza po kodu i ozbiljnosti;
- **Nalazi** — jedan red po nalazu: ozbiljnost, kod, tabela, ID, opis, predlog ispravke.

| Ozbiljnost | Značenje |
|---|---|
| `GRESKA` | podaci sami dokazuju grešku |
| `SUMNJA` | potpis greške, ali postoji i legitimno objašnjenje — operater proverava |
| `INFO` | skreće pažnju; nije greška sama po sebi |

## Odakle klase grešaka

Svaka klasa je nađena čitanjem koda koji je C002 koristio (tagovi `vba-v2.2.3`
… `vba-v2.28.4`): polje koje posle snimanja ostaje popunjeno, predlog koji
zastari kad se promeni vozač ili stanica, veza koju pisac nije proveravao, storno
koji nije kaskadirao. Za svaku je zapisan potpis u podacima i legitimni slučajevi
koji liče na grešku, pa ih je nezavisan skeptik proverio ponovnim čitanjem koda.

Polazni slučaj: prijemnica nasledi broj zbirne od prethodnog unosa
(`btnUnosPrij_Click` ne prazni `txtBrojZbirnePrij`, `vba-v2.11.0`), pa jedna
zbirna ima dve prijemnice, a prava nijednu.

## Veze u staroj svesci

Stara sveska nema tabele članstva ni `GeneracijaID`; veze su kolone na samim
dokumentima:

| Veza | Kako |
|---|---|
| otkup → otpremnica | `tblOtkup.OtpremnicaID` (po ID-ju) |
| otpremnica → zbirna | `tblOtpremnica.BrojZbirne` (po broju) |
| prijemnica → zbirna | `tblPrijemnica.BrojZbirne` (po broju) |
| novac → otkup | `tblNovac.OtkupID` |
| ambalaža → dokument | `DokumentTip` + `DokumentID` (reversi: broj dokumenta) |
| paletna stavka → prijemnica | `tblPaletaStavka.PrijemnicaID` (+ broj prijemnice i zbirne) |
| faktura stavka → prijemnica | `tblFakturaStavke.PrijemnicaID` |

Brojevi: otpremnica `x/ddmmyy[-n]` (x = broj stanice), zbirna `S` + broj
otpremnice, prijemnica `1/ddmmyy[-n]` (niz po kupcu i danu), otkup `x/ddmmyy[-n]`.

## Kod

| Fajl | Šta |
|---|---|
| `../provera_unosa.py` | komanda |
| `pokretac.py` | učita svesku, pusti module, skupi nalaze; pad modula je nalaz `MODUL_PAO` |
| `zajednicko.py` | `Kontekst` (tabele, indeksi), čitanje vrednosti, `parse_broj`, `nalaz` |
| `fixture.py` | čista izmišljena stara sveska — nad njom nijedna provera ne sme ništa da prijavi |
| `prijemnica.py`, `otpremnica_zbirna.py`, `otkup.py`, `novac.py`, `amb_pal_fak.py`, `storno.py` | provere po oblasti, svaka sa svojim `test_*.py` |

Testovi idu uz `python tools/migracija/migracija.py --self-test`.
