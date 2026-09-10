# Registar kompenzacija — gde se nedostatak modela plaća kodom

> Pravilo koje ovaj registar vodi: `.claude/rules/model-pre-zakrpe.md`.
> Zašto postoji: `docs/engineering/postmortems/2026-09-model-vs-zakrpa.md`.
>
> **Šta ovde ide:** zakrpa koja je izabrana **umesto** ispravke modela podataka,
> uz alternativu koja je odbačena i cenu koja je plaćena.
> **Šta ne ide:** običan bug (`KNOWN_ISSUES.md`), otvoreno pitanje domena
> (`docs/DOMEN/`), tehnički dug bez veze sa modelom (`KNOWN_ISSUES.md` TL-\*).
>
> Registar prati **obrazac**, ne procedure. Red se ne otvara zato što je pravilo
> pokrenuto — otvara se kad je zakrpa stvarno izabrana nad modelom.

## Kako se čita

- **Pogodaka** = koliko je puta ISTI nedostatak modela tražio novu zakrpu.
- **Treći pogodak** znači da zakrpa više nije podrazumevana
  (`model-pre-zakrpe.md` §7).
- ID format: `KMP-<broj>`, redom, bez ponovne upotrebe.

---

## KMP-001 — dokument nema zaglavlje (`header + stavke`)

| | |
|---|---|
| **Status** | `MODEL USVOJEN, NIJE IMPLEMENTIRAN` — 4 pogotka pre nego što je model postavljen |
| **Simptom** | dvoklasni dokument = dva fizička reda sa dva ID-a; logički dokument nema gde da živi |
| **Gde je model kriv** | `jedan fizički red = jedan dokument` nije tačno ni za jedan od 4 dokumenta u lancu |
| **Izabrane zakrpe** | `GeneracijaID` kao surogat identiteta · `ZbirnaIdentResolve` (10-poljni resolver) · kapije `ZBR-ACTIVE-NUMBER-01`, `ZBR-PARENT-01`, `ZBR-MUT-01`, `ZBR-CHILD-01`, `ZBR-NORM-02` · `ApplyGeneracijaID` u 3 pisca · read-model `AktivniBrojeviZbirne` · MIG-005b dedup u pickeru · `"OTK-1 + OTK-2"` konkatenacija kao de facto ID |
| **Odbačena alternativa** | `tbl*Stavke` — zaglavlje nosi dokument, stavka nosi klasu |
| **Zašto je odbačena** | nikad nije bila predložena; nijedna kapija nije postavila pitanje modela |
| **Cena do sada** | 9 PR-ova (#291–#301), 53 commita kroz 3 dana, od toga 22 `fix` (review treadmill); `Generacija*` 355 pojava u `src-vba/`, `ZbirnaIdent*` 74, `VlasniciPoBroju` 43, `RequireJedanVlasnik*` 20, `Split(x, " + ")` na 9 mesta |
| **Šta nestaje sa modelom** | sve iz reda „izabrane zakrpe"; `ZBR_IDENTITET.md` postaje `SUPERSEDED` (`DOCUMENT_HEADER_LINES.md` §10) |
| **Cena modela** | nema legacy transakcionih podataka → stari model se briše, ne prevodi (`REFAKTOR_DOKUMENT_HEADER_STAVKE.md`, uvod); šema je kanon u gitu → diff, ne migracija |
| **Pogodaka** | **4** — MIG-005a/b (dedup u pickeru) · `ZBR-IDENT-01` (identitet) · `ZBR-MUT-01` (storno po broju hvata dva dokumenta) · `ZBR-CHILD-01` (deca ne znaju svoj dokument) |
| **Detektor** | K1 ✓ (9 call site-ova) · K2 ✓ (`GeneracijaID`) · K3 ✓ (3 pisca pečate) · K4 ✓ (9 PR-ova) · K5 ✓ (broj+vlasnik+generacija kao „scope") — **5/5** |
| **Model** | `docs/DOMEN/DOCUMENT_HEADER_LINES.md` · plan: `docs/REFAKTOR_DOKUMENT_HEADER_STAVKE.md` |

> Nijedna od pet osa detektora nije bila skrivena. Sve su bile merljive `grep`-om
> u trenutku kad je prva zakrpa pisana.

---

## KMP-002 — poslovni broj kao korelacioni ključ (`BrDok`, poziv na broj)

| | |
|---|---|
| **Status** | `OTVOREN` |
| **Simptom** | poslovni broj se koristi za spajanje zapisa preko stanica i generacija — storno-po-broju, banka auto-map, relink ispravke, grupisanje za štampu |
| **Gde je model kriv** | isti nedostatak kao KMP-001: labela radi posao identiteta jer identiteta nema |
| **Izabrana zakrpa** | kapije po toku (`per-flow scope guards`), bez globalnog ugovora o jedinstvenosti |
| **Odbačena alternativa** | kanonski ID koji putuje od izabranog reda do pisca (`pre-flight` §5) |
| **Cena do sada** | nije merena u ovoj sesiji |
| **Pogodaka** | najmanje 1 (`BrojZbirne` krak je zatvoren kroz KMP-001) |
| **Izvor** | `docs/KNOWN_ISSUES.md` TL-008; `AUDIT_FM_TRIJAZA.md` FM-0012/0013/0021/0023/0031 |

> **Nije nezavisno provereno u ovoj sesiji** — red je preuzet iz TL-008 da obrazac
> ne bi ostao zapisan samo za slučaj koji je već rešen. Pre rada po njemu izmeri
> detektor (§2 pravila) nad stvarnim `BrDok` / poziv-na-broj putevima.
