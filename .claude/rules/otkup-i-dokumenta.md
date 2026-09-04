---
paths:
  - "src-vba/modOtkup.bas"
  - "src-vba/modOtkupBlok.bas"
  - "src-vba/modOtkupUnos.bas"
  - "src-vba/modDokumenta.bas"
  - "src-vba/modDokUnos.bas"
  - "src-vba/modDokumentInvariant.bas"
  - "src-vba/modScrDokumenti.bas"
---

# Otkup i dokumenta

> Najaktivnija oblast projekta. Do sada je bila jedini red u CLAUDE.md §3 bez
> svog rules fajla — a baš tu su živela tri buga zbog kojih postoji `modTest`.

## 1) `ClearForm` ima ugovor, ne „čisti ekran"

`modOtkupUI.ClearForm` se izvršava posle **svakog** snimanja otkupnog lista
(`modScrDokumenti.Scr_Save`, po uspehu `SaveOtkupMulti_TX`). Šta se briše a šta
ne — nije stvar ukusa, nego radnog toka operatera na terenu:

| Polje | Posle snimanja | Zašto |
|---|---|---|
| `fgDatum` | **ostaje** | sledeći blok ide u niz istog datuma otpremnice |
| `fgBrZbir` | **ostaje** | sledeći blok iste otpremnice nosi istu zbirnu |
| `cbKupac` | **briše se** | sledeći unos je nov partner |
| `fgKgI`, `fgKolAmb` | **brišu se** | podaci bloka odlaze sa partnerom |

Bez aktivne otpremnice datum se **vraća na danas** — prazno ili staro polje bi
bila greška koju operater ispravlja pri svakom novom dokumentu. To je jedina
razlika u odnosu na legacy, koji datum uopšte nije dirao.

Sve to meri `T_ClearForm_Ugovor` u `modTest`, i dokazano je u oba smera — vrati
brisanje datuma ili zbirne, odnosno ukloni brisanje partnera, i pukne tačno taj
test po imenu. **Ako menjaš ovu rutinu, prvo pročitaj šta test tvrdi.**

> Do koraka 2 je isti ugovor nosio `frmOtkup.ClearOtkupFields`, sa poljima
> `txtDatum` / `txtBrojZbirne` / `cmbKooperant`. Forma je obrisana
> (`docs/UI_MIGRACIJA_KATALOG.md` §27.10); ugovor je ostao.

## 2) Test seam-ovi žive u ljusci, ne više u formi

Dve linije koje su ovde stajale (`ClearOtkupFields` kao `Public`, `SetFocus` iza
`IsTestMode`) otišle su sa `frmOtkup` u koraku 2. Ljuska nosi svoje, i one su
popisane na jednom mestu — `.claude/rules/testovi.md` §4: `ClearForm` /
`ParseDatum` / `ParcelaID` su `Public`, tri `SetFocus`-a su iza `IsTestMode`, a
`modScrDokumenti.Scr_OtpTestSet` je tvrdo gejtovan.

Pravilo je isto kao i pre: **„čišćenje" koje ih vrati u prethodno stanje obara
suite — i to je jedini način da se to primeti.**

## 3) Kontekst otpremnice se napušta na jednom mestu

Datum i broj zbirne se nasleđuju sa izabrane otpremnice. Vraćaju se na „danas /
prazno" tek kad se taj kontekst stvarno napusti — u ljusci je to
`modScrDokumenti.Scr_OtpOtkazi` (prazni `mOtpID` / `mOtpBroj` i vraća listu na
otpremnice), a `ClearForm` posle njega vraća datum na danas (§1).

Legacy par je bio `OtkupBlok_ClearActiveOtp` + `frmOtkup.ResetDatumKontekst`.
Prvi postoji i dalje; drugog nema, ali ga `modOtkupBlok` i dalje zove
(`mForm.ResetDatumKontekst`). **To je mrtav kod, ne kvar:** `mForm` postavlja
`AttachOtkupBlokPanel`, a on posle koraka 2 nema nijednog pozivaoca — jedini je
bio `frmOtkup.UserForm_Initialize`. `mForm As Object` znači da poziv ne obara
compile. Čišćenje te polovine `modOtkupBlok`-a je zasebna odluka
(`docs/UI_MIGRACIJA_KATALOG.md` §27.10).

Ako dodaješ mesto koje resetuje kontekst, pitanje nije „da li da očistim polja"
nego „da li se kontekst otpremnice zaista napušta".

## 4) Snimanje ide kroz transakciju, ne kroz formu

`SaveOtkupMulti_TX` je jedina ulazna tačka za upis otkupa. `Scr_Save` →
`modOtkupUnos` iznad nje radi validaciju i prikupljanje vrednosti; ispod nje
idu best-effort koraci koji **ne smeju** da obore potvrdu snimanja (`OutputOtkupniList`,
`AutoChainHladnjaca`, prevezivanje paleta) — svi su pod `On Error Resume Next`.

Ne dodavati upis u tabele mimo `SaveOtkupMulti_TX`, i ne premeštati best-effort
korake iznad nje.

## 5) Novi UI je ljuska — legacy se penzioniše po koracima

Otkup i dokumenta su preneti na jednu runtime formu (`frmOtkupUI` + ljuska
`modOtkupUI` + ekranski moduli `modScr*`).

Zato `frmOtkup` i `frmDokumenta` više ne stoje „paralelno dok oba sistema ne
budu potpuna" nego se **penzionišu po koracima**: `docs/UI_MIGRACIJA_KATALOG.md`
§27 nosi inventar formi (§27.2), redosled uklanjanja (§27.3) i verifikaciju po
koraku (§27.5). Obe su otišle u **koraku 2** (§27.10) — ovaj fajl od tada
opisuje ljusku, a legacy se pominje samo tamo gde objašnjava zašto je nešto
ovakvo kakvo jeste.

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
- Ugovor iz §1 nosi `modOtkupUI.ClearForm` — ista tri ponašanja, isti razlozi
  kao u legacy formi koje više nema. **Pokriven je testom** (`T_ClearForm_Ugovor` u
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
