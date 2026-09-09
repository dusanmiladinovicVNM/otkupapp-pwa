# Architecture Contract

> Pravila koja važe za **svaki** budući domen u AgriX-u, ne samo za lanac
> dokumenata. Kratko namerno — ugovor koji se ne pamti se ne poštuje.
>
> Svako pravilo ima **kako se proverava**. Pravilo bez provere je želja.
> Status: usvojeno 2026-09-09. Model lanca dokumenata: `DOCUMENT_HEADER_LINES.md`.

---

## A1 — jedan logički dokument = jedan ID

`ZbirnaID` je cela zbirna, nikad njena klasa ili stavka. Public create API vraća
**tačno jedan** ID.

*Provera:* `vba_check` pravilo `NEMA_ID_PLUS_ID` — nijedan `Create*_TX` ne vraća
konkatenaciju, nigde `Split(..., " + ")` nad ID stringom.

## A2 — poslovni broj nije identitet

`BrojDokumenta`, `BrojOtpremnice`, `BrojZbirne`, `BrojPrijemnice`, `BrojFakture`
su **labele**. Smeju za: pretragu, prikaz, generator sledećeg broja, validaciju
duplikata pri unosu. Ne smeju za: FK, storno target, izbor „svih redova
dokumenta", rekalkulaciju roditelja, correction identitet, ownership guard.

*Provera:* `vba_check` pravilo `NEMA_BROJA_KAO_FK` (allowlist dozvoljenih
konteksta) + acceptance test `BrojNijeIdentitet`.

## A3 — header + stavke

Ako dokument može da nosi ponovljive podatke, nosi ih kao stavke.

> **Polje koje pripada dokumentu kao celini → header. Polje koje poslovno može
> imati različitu vrednost među stavkama istog dokumenta → stavka.**

Ne kopirati header polja u svaku stavku.

**Postojeći potpis ulaza je dokaz trenutne poslovne kardinalnosti, ne definicija
domena.** Za refaktor dokumenata je `Save*Multi_TX` bio vrlo dobar dokaz — autor
je za svako polje već odlučio da li varira po klasi — ali taj signal ne sme da
postane pravilo. Inače jedan ekran koji danas nudi samo jedan proizvod sutra
natera model Prodaje da tvrdi kako proizvod pripada zaglavlju.

Kad se signal iz UI-ja i poslovna kardinalnost razilaze, **poslovna odlučuje**, a
razlika se zapisuje kao **ODLUKA** u modelu.

*Provera:* review + `DOCUMENT_HEADER_LINES.md` tabela grain-a.

## A4 — odnos se zapisuje, ne nagađa

Ako roba A ulazi u dokument B, `A_ID → B_ID` je upisana poslovna činjenica.
Nikad `datum + vozač + klasa + broj ≈ verovatno B`.

*Provera:* acceptance test `TraceBezPogadjanja` — lanac unazad prolazi samo kroz
FK, bez ijednog poređenja po datumu/broju/vozaču.

## A5 — jedan kanonski izvor istine

Ako se vrednost može izvesti iz stavki, ona je **izvedena** ili **eksplicitno
označen keš**. Nikad dva ravnopravna izvora. Keš se imenuje kao keš u nazivu
kolone ili komentaru, i ima test koji dokazuje da se poklapa sa izvorom.

*Provera:* invarijantni testovi po dokumentu (`DOCUMENT_HEADER_LINES.md` §6).

## A6 — identitet zaliha je `LagerJedinicaID`

Kad roba postane lager roba, njen identitet je lager jedinica — ne paleta, ne
prijemnica, ne tekst. *(Primenjuje se od Milestone 2; ovde je zapisano da se
budući model ne bi ponovo vezao za paletni red.)*

## A7 — proces je događaj

`N ulaza → Proces → M izlaza`, sa eksplicitnim vezama na obe strane i masa-balansom
kao invarijantom writer-a. *(Prerada 2.0; spec postoji, kod ne.)*

## A8 — storno nije brisanje

Storniran red ostaje i izlazi iz svih agregata. „Aktivan" nije isto što i
„postoji". Koje tabele nose storno je **deklarisano**, ne pogađa se.

*Provera:* postojeće `vba_check` pravilo `STORNO_REGISTAR` + `modSchemaGuard`
registri `STORNO_TABELE` / `BEZ_STORNA`.

## A9 — ispravka je nov ID

Nova verzija dokumenta dobija **nov** `DocumentID` i kad poslovni broj ostaje isti.
Veza je `IspravkaOdID` / `ZamenjenSaID`, po ID-u — nikad po broju.

*Provera:* acceptance test `IspravkaID`.

## A10 — sync ima nepromenljiv eksterni ID

Retry ne pravi duplikat. Eksterni identitet stoji na nivou na kom stvarno
identifikuje eksterni entitet, i idempotentnost se proverava baš nad njim.

*Provera:* acceptance test `PWAReimport` (isti record dvaput → jedan dokument).

## A11 — jedna poslovna operacija ima jednog vlasnika upisa

Svaka domen-tabela ima **imenovanu listu modula koji smeju da je pišu**. Svi
ostali zovu API vlasnika.

```
danas:  modNovac ─────────> UpdateCell tblOtkup
target: modNovac ──> Otkup_RecomputePaymentStatus ──> tblOtkup
```

Vlasnički API sme fizički da živi u istom modulu — ne traži se nova topologija
modula, traži se jedan ulaz.

*Provera:* `docs/DOMEN/WRITE_OWNERSHIP.json` + `who_writes.py --check-ownership`.
Nov pisač koji nije na listi = pad CI-ja.

**Snapshot nije vlasništvo.** `AddTableSnapshot` znači „moja transakcija mora da
ume da vrati ovu tabelu", ne „ja sam pišem". Ciljna arhitektura ima koordinatora
koji snapshotuje tuđu tabelu i zove API njenog vlasnika. Kapija zato meri
**mutatore** (`AppendRow` / `UpdateCell` / `RequireUpdateCell`), a učesnici
transakcije se prikazuju odvojeno.

**Kapija proverava isključivo `row_owner`.** `schema_owner` je zaseban pojam
(ko sme da napravi tabelu ili kolonu) i **ne učestvuje** u proveri mutacije reda
— unija dva spiska bi bila poznat bypass: propuštala bi baš ono što ugovor
zabranjuje. Ko sme da menja šemu je druga provera, ne ova.

`modSetup` je danas u `row_owner` baseline-u zato što **stvarno piše** redove
(backfill `BrojOtpremnice` nad `tblOtkup`, admin nalog nad `tblKorisnici`) — ne
zato što sme. Cilj ga isključuje.

**Zatečeno stanje koje ovo pravilo cilja** (mereno nad mutatorima,
`WRITE_OWNERSHIP.json`):

| Tabela | `row_owner` danas | Cilj |
|---|---|---|
| `tblOtkup` | 9 | `modOtkup` |
| `tblFakturaStavke` | 3 | — |
| `tblFakture` | 3 | — |
| `tblKorisnici` | 3 | — |
| `tblNovac` | 3 | — |
| `tblPrijemnica` | 3 | `modDokumenta` |

> Do PR1 je kapija merila i snapshot i upis pomešano, a `RequireUpdateCell` joj
> je bio **nevidljiv** (regex je tražio granicu reči pred `UpdateCell`). Time je
> 220 poziva prolazilo neopaženo — uključujući `modSEFPersistance` nad
> `tblFakture`, koji u mapi uopšte nije postojao.

## A12 — UI ne zna detalje persistencije

UI kaže `CreateZbirna(header, stavke)`, ne `AppendRow TBL_ZBIRNA, Array(...)`.
Adapter (`modOtkupUnos` / `modDokUnos`) prevodi polja forme u DTO; dalje ide
domen.

*Provera:* `who_writes.py` — nijedan `modScr*` / `frm*` / `mod*UI` nije na
ownership listi domen-tabela.

---

## Četiri tvrde kapije

| Gate | Tvrdnja | Kako se meri |
|---|---|---|
| **G1 — reproduktivna sveska** | prazan `.xlsm` + uvoz koda + `SetupNewPC` = spremna aplikacija, bez ručnog koraka. Kanon je `schema/schema.json` u gitu; sveska je posledica. | **delimično dokazano** — v. G1a |
| **G1a — rekonstrukcija tabele iz koda** | obrisana tabela se vraća iz registra, bez donora i bez ručnog koraka | test 201 `T_Sema_SamoLeci`: obriši `tblMGMT` → `VerifySchema` je prijavi → `EnsureAllTables` je vrati → šema čista |
| **G1b — redosled kolona je deo šeme** | upis je pozicion (`AppendRow`), pa svako razilaženje **pre kraja** — premeštena kolona, izbačena iz sredine, ubačena u sredinu, ili **produženo ime poslednje kanonske kolone** — tiho šalje vrednosti u pogrešna polja | poređenje po **indeksu kolone** (`PrefiksNeslaganje`, jedan helper za obe kapije), otisak nad uređenim prefiksom uz self-heal, `SchemaReadyOrFail` pred pozicionim upisom, `schema_diff` blokira uvoz. Testovi 199, 202. |

> **G1 još NIJE dokazana u celini.** Dokazana je njena *rekonstrukciona*
> komponenta (G1a): jedna nestala tabela se vraća iz registra. Nije mereno:
> prazan `.xlsm` + uvoz svih modula + `.frm`/`.frx` + `SetupNewPC` + config
> bootstrap = upotrebljiv AgriX. Dok to nema svoj test, ugovor ne sme da tvrdi
> više nego što meri.
>
> **Gde kapija `SchemaReadyOrFail` stoji danas** (10 poziva): `SaveOtkup_TX`,
> `SaveOtkupMulti_TX`, `SaveOtpremnica_TX`, `SaveOtpremnicaMulti_TX`,
> `SaveZbirna_TX`, `SaveZbirnaMulti_TX`, `SavePrijemnica_TX`,
> `SavePrijemnicaMulti_TX`, `SaveNovac_TX`, `CreateFaktura_TX`.
> **Nisu još gejtovani** pisci van lanca dokumenata — agrohemija, banka, geo,
> palete, utovar. Dobijaju kapiju kad im dođe red u refaktoru; do tada ugovor to
> ne sme da tvrdi.
| **G2 — duplikat broja** | dva dokumenta sa istim poslovnim brojem ne prave **nijedan** poseban code path u jezgru | test `BrojNijeIdentitet` + pravilo `NEMA_BROJA_KAO_FK` |
| **G3 — sledljivost bez pogađanja** | lanac unazad ide samo kroz ID/FK graf | test `TraceBezPogadjanja` |
| **G4 — vlasništvo upisa** | nov pisač domen-tabele obara CI | `who_writes.py --check-ownership` |

---

## Pre-Flight — obavezan pre implementacije bilo kog novog domena

1. Grain — šta je jedan red
2. Identitet — šta je PK, i zašto baš to
3. Header/stavke — šta varira
4. FK kardinaliteti — 1:1, 1:N, N:M, delimična alokacija
5. Izvor istine — šta je izvedeno, šta je keš
6. Invarijante — šta mora uvek da važi
7. Prelazi stanja
8. Storno
9. Ispravka
10. Granica transakcije — koje tabele operacija dira
11. Nizvodni potrošači
12. Ivični slučajevi
13. Acceptance testovi

**Ako bilo koja od prvih šest stavki nije jasna — nema implementacije.**
To je lek protiv onoga što se dogodilo sa `GeneracijaID`: kolona je nastala zato
što grain i identitet nisu bili odlučeni pre writer-a.

---

## Šta ovaj ugovor namerno NE traži

- Generički Repository / App framework
- `Cmd` / `Qry` / `Rules` sloj po dokumentu
- Jedan `tblDokumenti` sa EAV-om
- Novo stablo koda

A11 se dobija listom dozvoljenih pisača, ne novom topologijom modula. Ako se
posle refaktora pojavi potreba za slojevima, to je zasebna odluka sa zasebnim
dokazom.
