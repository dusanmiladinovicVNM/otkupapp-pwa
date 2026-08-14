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

## 3) Kontekst otpremnice se resetuje samo kad se stvarno napušta

Datum i broj zbirne se nasleđuju sa izabrane otpremnice (panel „Otkupni
blokovi"). Vraćaju se na „danas / prazno" samo kad se taj kontekst stvarno
napušta.

**Legacy (`frmOtkup`) — dva mesta:**

- `ResetDatumKontekst` (`Public`, zove `modOtkupBlok` kad se blokovi sakriju)
- promena otkupnog mesta **van** konteksta otpremnice, uz `OtkupBlok_IsPrefilling`
  gard — prefill sam postavlja stanicu i NE sme da se resetuje

**Novi UI (`modOtkupUI`) — jedna rutina, tri poziva.** Otpuštanje je izdvojeno u
`OtpustiOtpremnicu(refresh)`; ona i zove `Scr_OtpOtkazi`, i vraća datum, zbirnu
(`mAktivnaZbirna`) i robu. Poziva se sa:

| Poziv | Kada | `refresh` |
|---|---|---|
| `NapustiOtpremnicu` | promena OM-a na **drugu** stanicu | `True` |
| `SelectModeCore` | izlazak iz **F1** u drugi režim | `False` |
| `ActivateScreen` | izlazak sa ekrana **DOKUMENTI** (Palete, Agrohemija…) | `False` |

`refresh=False` je zato što ta dva pozivaoca mrežu i raspored osvežavaju sama —
inače bi lista bila pročitana dva puta.

Uz to, datum pripada **režimu**: `SetDatumPoRezimu` (u `SelectModeCore`, samo na
stvarnu promenu režima) daje datum otpremnice u F1 dok je aktivna, a **danas** u
svakom drugom slučaju.

> Do `v6-ui-116` je `Scr_OtpOtkazi` imao **jednog jedinog pozivaoca**
> (`NapustiOtpremnicu`), pa se kontekst puštao isključivo na promenu OM-a. F1 → F2
> → F1 i odlazak na Palete ostavljali su datum otpremnice u polju i traku na
> ekranu, a **nova otpremnica u F2 nasleđivala je datum stare** — pogrešan datum u
> dokumentu, bez ijednog znaka operateru. Prijavljeno sa terena, a ne testom: u
> novi UI je bio prenet samo drugi red legacy tabele iznad.

Ako dodaješ novo mesto, pitanje nije „da li da očistim polja" nego „da li se
kontekst otpremnice zaista napušta" — i ide kroz `OtpustiOtpremnicu`, ne kao nova
kopija tih pet upisa.

## 4) Snimanje ide kroz transakciju, ne kroz formu

`SaveOtkupMulti_TX` je jedina ulazna tačka za upis otkupa. `btnUnos_Click` iznad
nje radi samo validaciju i prikupljanje vrednosti; ispod nje idu best-effort
koraci koji **ne smeju** da obore potvrdu snimanja (`OutputOtkupniList`,
`AutoChainHladnjaca`, prevezivanje paleta) — svi su pod `On Error Resume Next`.

Ne dodavati upis u tabele mimo `SaveOtkupMulti_TX`, i ne premeštati best-effort
korake iznad nje.

## 5) Novi UI (`frmOtkupUI`) postoji paralelno — legacy se NE gasi

Otkup i dokumenta se prenose na jednu runtime formu (`frmOtkupUI` + ljuska
`modOtkupUI` + ekranski moduli `modScr*`). **Dok oba sistema ne budu potpuna,
`frmOtkup` i `frmDokumenta` ostaju potpuno operativni i ne diraju se.**

- **Izvor istine za stanje prelaska:** `docs/UI_MIGRACIJA_KATALOG.md` — pravila
  Z1–Z14, brojevni niz po režimu (Z3a), revizija ulaznog sloja 1:1 (Z3b),
  isporuka (Z3c) i plan po fazama. „Šta još nije preneto" se čita odatle, ne
  zaključuje iz koda.
- **Poslovna logika unosa je izdvojena iz formi:** `modOtkupUnos` (otkupni list)
  i `modDokUnos` (otpremnica…) — bez ijedne kontrole, zovu ih i ekran i, kad za
  to dođe red, forma.
- **Prelazno pravilo:** pravilo unosa se menja u tim modulima, pa se **ručno
  preslika** u legacy formu, i to se zabeleži uz izmenu. Dve kopije postoje
  namerno.
- Ugovor `ClearOtkupFields` iz §1 važi i za `modOtkupUI.ClearForm` — ista tri
  ponašanja, ista tri razloga. **Pokriven je testom** (`T_ClearForm_Ugovor` u
  `modTest`), zajedno sa `ParseDatum` i `ParcelaID`. Novi UI ima i seam-ove iz §2
  u svom obliku (`ClearForm`/`ParseDatum`/`ParcelaID` su `Public`, tri `SetFocus`-a
  su iza `IsTestMode`, `Scr_OtpTestSet` je gejtovan) — detalji i sabotaže:
  `.claude/rules/testovi.md` §4.
- Razlika u odnosu na legacy koju test fiksira: **bez** aktivne otpremnice novi UI
  vraća datum na danas (legacy ga uopšte ne dira). Uz aktivnu otpremnicu ponašanje
  je isto — datum ostaje njen.

## 6) Verifikacija

Izmena u ovoj oblasti nosi test u `modTest` — vidi `.claude/rules/testovi.md`.
Checklista u chatu je samo za ono što se ne može automatizovati (izgled forme,
štampa, PDF, ponašanje nad pravim podacima).
