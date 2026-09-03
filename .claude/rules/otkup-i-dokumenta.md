---
paths:
  - "src-vba/frmOtkup.frm"
  - "src-vba/modOtkup.bas"
  - "src-vba/modOtkupBlok.bas"
  - "src-vba/frmDokumenta.frm"
  - "src-vba/modDokumenta.bas"
  - "src-vba/modDokumentInvariant.bas"
---

<!-- `frmOtkup.frm` i `frmDokumenta.frm` u `paths` odlaze SA formama (korak 2,
     `docs/UI_MIGRACIJA_KATALOG.md` §27.3). Do tada moraju da stoje: agent koji
     dira legacy formu mora da vidi §1–§3. -->

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

## 5) Novi UI je ljuska — legacy se penzioniše po koracima

Otkup i dokumenta su preneti na jednu runtime formu (`frmOtkupUI` + ljuska
`modOtkupUI` + ekranski moduli `modScr*`).

Zato `frmOtkup` i `frmDokumenta` više ne stoje „paralelno dok oba sistema ne
budu potpuna" nego se **penzionišu po koracima**: `docs/UI_MIGRACIJA_KATALOG.md`
§27 nosi inventar formi (§27.2), redosled uklanjanja (§27.3) i verifikaciju po
koraku (§27.5). Obe odlaze u **koraku 2**.

Redosled nije stvar ukusa: **prvo se seku reference, pa forma.** Forma bez
referenci se kompajlira i ne smeta; forma koja referencira obrisano obara
compile cele sveske. I `git rm` nije kraj posla — ni self-update ni
`ImportAllVBA` ne brišu komponente, pa svaki korak nosi ručni `Remove` u VBE
po instalaciji (§27.4).

**Dok njihov korak ne bude isporučen, obe forme ostaju potpuno operativne i ne
diraju se.** „Ne diraju se" znači: ne menjaju se mimo koraka koji ih uklanja.

- **Izvor istine za stanje prelaska:** `docs/UI_MIGRACIJA_KATALOG.md` — pravila
  Z1–Z14, brojevni niz po režimu (Z3a), revizija ulaznog sloja 1:1 (Z3b),
  isporuka (Z3c) i plan po fazama. „Šta još nije preneto" se čita odatle, ne
  zaključuje iz koda.
- **Poslovna logika unosa je izdvojena iz formi:** `modOtkupUnos` (otkupni list),
  `modDokUnos` (otpremnica, zbirna, prijemnica) i `modNovacUnos` (isplate, uplate
  kupaca, reversi) — bez ijedne kontrole, zovu ih i ekran i, kad za to dođe red,
  forma. Ekran (`modScrDokumenti.Scr_Save`) samo prevodi polja u rečnik;
  **nijedna provera ne živi u ekranu.**
- **Prelazno pravilo:** pravilo unosa se menja u tim modulima, pa se **ručno
  preslika** u legacy formu, i to se zabeleži uz izmenu. Dve kopije postoje
  namerno — i **prestaju da postoje sa korakom 2**, kad forme odu. Od tada
  pravilo živi na jednom mestu i preslikavanja nema.
- Ugovor `ClearOtkupFields` iz §1 važi i za `modOtkupUI.ClearForm` — ista tri
  ponašanja, ista tri razloga. **Pokriven je testom** (`T_ClearForm_Ugovor` u
  `modTest`), zajedno sa `ParseDatum` i `ParcelaID`. Upis zbirne i prijemnice
  (`ZbirnaValidiraj` / `PrijemnicaValidiraj`) ima svojih pet testova u istoj
  suite-i, a upis novca i ambalaže (`IsplataValidiraj` / `UplataValidiraj` /
  `ReversValidiraj`) svoja tri — spisak i sabotaže: `.claude/rules/testovi.md` §4. Novi UI ima i seam-ove iz §2
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
