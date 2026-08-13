---
paths:
  - "src-vba/frmOtkup.frm"
  - "src-vba/modOtkup.bas"
  - "src-vba/modOtkupBlok.bas"
  - "src-vba/frmDokumenta.frm"
  - "src-vba/modDokumenta.bas"
  - "src-vba/modDokumentInvariant.bas"
---

# Otkup i dokumenta

> Najaktivnija oblast projekta. Do sada je bila jedini red u CLAUDE.md §3 bez
> svog rules fajla — a baš tu su živela tri buga zbog kojih postoji `modTest`.

## 1) `ClearOtkupFields` ima ugovor, ne „čisti formu"

`frmOtkup.ClearOtkupFields` se izvršava posle **svakog** snimanja otkupnog lista
(`btnUnos_Click`, po uspehu `SaveOtkupMulti_TX`). Šta se briše a šta ne — nije
stvar ukusa, nego radnog toka operatera na terenu:

| Polje | Posle snimanja | Zašto |
|---|---|---|
| `txtDatum` | **ostaje** | sledeći blok ide u niz istog datuma otpremnice |
| `txtBrojZbirne` | **ostaje** | sledeći blok iste otpremnice nosi istu zbirnu |
| `cmbKooperant` | **briše se** | sledeći unos je nov partner |

Sva tri su pokrivena testovima (`modTest`), i sva tri su dokazana u oba smera —
vrati brisanje datuma ili zbirne, odnosno ukloni brisanje kooperanta, i pukne
tačno taj test po imenu. **Ako menjaš ovu rutinu, prvo pročitaj šta test tvrdi.**

Ta rutina se dodatno oslanja na to da je forma živa: zove `AutoFillCenaOtkup`
(vraća auto-cenu za i dalje izabran proizvod) i `RefreshBrojDokumentaSuggestion`.
Obe su pod `On Error Resume Next` odnosno tolerantne na prazan kontekst.

## 2) Test seam u formi — dve linije koje se ne diraju

- `ClearOtkupFields` je **`Public`**, ne `Private`. To je namerno: `modTest` je
  zove direktno, bez vožnje celog `btnUnos_Click` (koji traži stanica-lock, PDF
  izlaz i auto-lanac hladnjače).
- `cmbKooperant.SetFocus` je iza **`If Not IsTestMode()`**. Forma koja nije
  `.Show`-ovana ne može da primi fokus; bez garda bi svi testovi padali na fokusu
  umesto na ponašanju. U produkciji je `IsTestMode()` uvek `False`.

Obe linije nose komentar u samoj formi. „Čišćenje" koje ih vrati u prethodno
stanje obara test suite — i to je jedini način da se to primeti.

## 3) Kontekst otpremnice se resetuje na tačno dva mesta

Datum i broj zbirne se nasleđuju sa izabrane otpremnice (panel „Otkupni
blokovi"). Vraćaju se na „danas / prazno" samo kad se taj kontekst stvarno
napušta:

- `ResetDatumKontekst` (`Public`, zove `modOtkupBlok` kad se blokovi sakriju)
- promena otkupnog mesta **van** konteksta otpremnice, uz `OtkupBlok_IsPrefilling`
  gard — prefill sam postavlja stanicu i NE sme da se resetuje

Ako dodaješ treće mesto, pitanje nije „da li da očistim polja" nego „da li se
kontekst otpremnice zaista napušta".

## 4) Snimanje ide kroz transakciju, ne kroz formu

`SaveOtkupMulti_TX` je jedina ulazna tačka za upis otkupa. `btnUnos_Click` iznad
nje radi samo validaciju i prikupljanje vrednosti; ispod nje idu best-effort
koraci koji **ne smeju** da obore potvrdu snimanja (`OutputOtkupniList`,
`AutoChainHladnjaca`, prevezivanje paleta) — svi su pod `On Error Resume Next`.

Ne dodavati upis u tabele mimo `SaveOtkupMulti_TX`, i ne premeštati best-effort
korake iznad nje.

## 5) Verifikacija

Izmena u ovoj oblasti nosi test u `modTest` — vidi `.claude/rules/testovi.md`.
Checklista u chatu je samo za ono što se ne može automatizovati (izgled forme,
štampa, PDF, ponašanje nad pravim podacima).
