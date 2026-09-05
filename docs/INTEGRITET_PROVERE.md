# Integritet provere (tabela ↔ tabela)

Konsolidovana, **read-only** revizija integriteta podataka kroz ceo lanac
dokumenata: `Otkup → Otpremnica → Zbirna → Prijemnica → Paleta → Prerada`.
Ništa ne menja — samo izlistava neusklađene zapise.

## Pokretanje

- **Crveni baner na startu** (`frmOtkupAPP`): klik → **in-app pregled** (ListBox
  overlay, kolone `Provera | Detalj`, dugme „Zatvori"). Ne dira sheet.
- **Admin panel** → grupa „Setup i provere" → **„Integritet provere (tabele)"**, ili
- `Alt+F8 → RunIntegritetProvere` → upis u **sheet**.

Baner i in-app pregled koriste `GetIntegritetRows()` (samo problemi, u memoriji).
Admin/Alt+F8 put upisuje u sheet **`INTEGRITET_PROVERE`** (obriše se i iznova napiše
pri svakom pokretanju): svaki blok naslov + broj problema, pa „OK – nema" ili
tabela zapisa, na kraju „UKUPNO" + zbirni `MsgBox`. Sheet se filtrira/sortira/
štampa u Excelu; in-app pregled je brz uvid bez izlaska iz app-a.

Sve provere isključuju stornirane redove (`ExcludeStornirano`) i agregiraju po
`BrojZbirne` (Klasa I + II dele isti broj → zaseban red).

## Provere

### A — Konzervacija količine (kg)
| Kod | Značenje | Napomena |
|---|---|---|
| A1 | `Σ otpremnica.Kolicina` vs `Σ zbirna.UkupnoKolicina` po `BrojZbirne` | reuse `ValidateZbirna` (prag 0.01 kg) |
| A2 | Manjak/višak `zbirna → prijemnica`: **VIŠAK > 5%** (prijemnica > zbirna), **NIŠTA PRIMLJENO**, **MANJAK > 10%** | pragovi `PRAG_VISAK_PCT` (5%) / `PRAG_MANJAK_PCT` (10%) |
| A3 | `Σ paleta-stavke.NetoKg` po prijemnici vs `prijemnica.Kolicina` | samo paletizovane; tol 0.5 kg |
| A4 | `paleta.NetoKg`/`BrojGajbica` (header) vs `Σ stavke` | tol 0.5 kg / 0.001 gajbe |
| A5 | prerada `NetoUlazKg` vs `Σ stavke.NetoKg`; `NetoIzlaz ≤ NetoUlaz` | tol 0.5 kg |

### B — Referencijalni integritet lanca
| Kod | Značenje |
|---|---|
| B1a/B1b | **Verwaist** otpremnice/prijemnice: živ dokument, `BrojZbirne` **potpuno stornirane** zbirne (reuse `GetVerwaisteDokumente`) |
| B2 | **Otkupi bez otpremnice** (`OtpremnicaID` prazan) — reuse `GetUnlinkedOtkupi` |
| B3 | **Izgubljeni otkup blokovi**: otkup čiji `OtpremnicaID` → stornirana/nepostojeća otpremnica (reuse `GetLostOtkupBlokovi`) |
| B4a/B4b | Otpremnica/prijemnica sa `BrojZbirne` koji **uopšte ne postoji** u `tblZbirna` (različito od B1) |
| B5 | Prijemnica **bez `BrojZbirne`** (obavezna veza) |
| B5b | Otpremnica **bez `BrojZbirne`** (nije vezana za zbirnu) |
| B6 | `BrojZbirne` se poklapa sa zbirnom **samo do velikog/malog slova** (npr. `s5/…` vs `S5/…`) — advisory za normalizaciju; skenira otpremnicu/prijemnicu/paleta-stavku/otkup |
| B7 | **Zbirna sa 0** (ili prazan) `UkupnoKolicina` — sama po sebi anomalija (komplement A2: A2 hvata zbirne-sa-kg-bez-prijema) |

### C — Palete
| Kod | Značenje |
|---|---|
| C1 | Paleta-stavka bez žive prijemnice (prazan/nepostojeći `PrijemnicaID`) |
| C2 | Paleta-stavka bez ispravne zbirne (prazan/nepostojeći `BrojZbirne`) |
| C3 | Paleta (header) bez ijedne aktivne stavke (orphan header) |
| C4 | Paleta-stavka ka **storniranoj** prijemnici (kaskadni storno ne dira `tblPaletaStavka`) |
| C5 | Dupli `BrojPalete` unutar iste `Godina` |

### D — Prerada
| Kod | Značenje |
|---|---|
| D1 | Paleta `Preradjeno=Da` bez ijedne aktivne prerada-stavke (reset flaga izostao) |
| D2 | Prerada-stavka ka nevalidnoj paleti (nesveža/stornirana/nepostojeća) — **očekivano prazno** zbog storno guarda |

### P — Prerada 2.0: lager jedinica (Faza A)
| Kod | Značenje | Napomena |
|---|---|---|
| P4 | Prerada ↔ lager jedinica: bez LJ / više LJ (`IzvorTip=PRERADA`, isti `IzvorID`) / `KgPocetno ≠ NetoIzlazKg` (aktivna) / LJ bez `ProizvodID` iako prerada ima tip / obrnuti pokazivač `tblPrerada.LagerJedinicaID ≠ LJ` | tol 0.01 kg; sve prerade, i stornirane (LJ nasleđuje `Stornirano`) |
| P5 | Aktivna utovarna/fakturna GP stavka bez `LagerJedinicaID`, sa nepostojećom LJ, ili sa LJ koja pripada drugoj preradi | meko po tabeli (sveska pre nadogradnje) |
| P6 | Lager jedinica: prazan/dupli `LagerJedinicaID`, nepoznat `IzvorTip`, prazan `IzvorID`, aktivna jedinica bez `StanicaID` | `StanicaID` je prazan kad sveska nema tačno jednu stanicu `JeHladnjaca=Da` |

## Podesive tolerancije

U `modIntegritet.bas`, vrh modula:
- `PRAG_MANJAK_PCT` (10) — manjak% iznad ovoga se prijavljuje (A2).
- `PRAG_VISAK_PCT` (5) — višak do ovoga se **ne** prijavljuje (A2).
- kg tolerancije (0.5) su inline u A3/A4/A5 — lako promeniti ako po-gajbi
  zaokruživanje pravi šum.

**Poređenje `BrojZbirne` je case-insensitive** (`s5/…` = `S5/…`) — u
`AllBrojeviInZbirna` i `AggByBroj` (`CompareMode = vbTextCompare`), pa razlika u
velikom/malom slovu ne pravi lažni „ne postoji" (B4/C2) ni razdvojenu grupu (A1/A2).

## Odnos prema invarijantama (potvrđeno iz koda)

- Razlika `Otkup ↔ Otpremnica` **nije** samo „verwaist": čine je **B2** (unlinked)
  + broken FK (`Check_OtkupOtpremnicaCrossZbirnaLinks`). Verwaist se tiče
  `otpremnica/prijemnica ↔ zbirna`.
- `Otpremnica ↔ Zbirna`: kg = **A1**, reference = **B1** + **B4**.
- Prijemnica: manjak/višak = **A2**, obavezna zbirna = **B5**.
- Palete moraju imati prijemnicu = **C1**; bez zbirne = **C2**.
- „Prerađeno bez sveže palete" = **D2** (očekivano prazno — već-prerađena paleta
  se ne može stornirati; oporavak samo preko `StornoPrerada`).

Storno **ne kaskadira** globalno (osim hladnjača-blok i malina putanje), pa viseći
dokumenti i orphan stavke nastaju legitimno — zato ova revizija postoji.
