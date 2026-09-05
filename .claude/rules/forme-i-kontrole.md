---
paths:
  - "src-vba/frm*.frm"
  - "src-vba/cls*.cls"
  - "src-vba/modOtkupBlok.bas"
  - "src-vba/modMaticniLookups.bas"
  - "src-vba/modPaletniListUI.bas"
  - "src-vba/mod*UI.bas"
---

# Forme, `.frx` i runtime kontrole

> Preseljeno iz `CLAUDE.md` §3/§4. Učitava se kad se dira UI sloj.

## Koje forme ostaju

Plan iz `docs/UI_MIGRACIJA_KATALOG.md` §27 je **ispunjen** (korak 7, §27.17).
Formi ima **četiri** i to su sve:

| Forma | Uloga |
|---|---|
| `frmOtkupUI` | ljuska — jedini operaterski ekran; sadržaj daju `modScr*` |
| `frmLogin` | prijava |
| `frmSplash` | start |
| `frmExcelMini` | povratak iz Excela |

**Legacy formi više nema**, pa nema ni izuzetka: nova kontrola ide na **ekran**
(`modScr*`), tačka. Ranije je ovde stajalo „osim kad §5 `otkup-i-dokumenta.md`
izričito traži" — taj izuzetak je postojao zbog preslikavanja pravila unosa u
formu koja ga je duplirala i otišao je sa korakom 2.

## Kako se forma penzioniše (obrazac iz §27)

Vredi i za sve što tek dolazi:

1. **Prvo se seku reference, pa forma** (§27.3). Forma koja ostane a nema
   referenci se kompajlira i ne smeta; forma koja referencira obrisano **obara
   compile cele sveske**.
2. **Forma bez ekrana se ne briše nego joj se NAPIŠE zamena** — makar ta zamena
   zasad pošteno rekla da još ne radi (`frmMarza` → ekran `ANALIZA` „U IZRADI",
   §27.15). Prazan ekran koji to kaže bolji je od prigušene stavke koja na klik
   javlja da ekrana nema.
3. **Pre brisanja se izmeri šta se gubi.** Tvrdnje koje forma nosi se
   **presele**, ne obrišu — a ako ljuska tu tvrdnju već ima, proverava se da
   stvarno grize (§27.14: jedna je bila bez ijedne sabotaže).
4. **Po instalaciji ide `Remove`.** Self-update nikad ne uklanja komponente
   (§27.4), pa zaostala forma živi dok je operater ručno ne skine.

## `.frx` se ne dira kao tekst

Kontrole formi žive u binarnom `.frx` (`.frm` ima samo `OleObjectBlob`). Nove
kontrole se **NE** dodaju editovanjem teksta — dodaju se u runtime-u
(`Controls.Add` + WithEvents), kao postojeći `modOtkupBlok`/`clsBlokUI`.

Caption koji živi samo u `.frx` menja se u dizajneru forme, ne u kodu.
Pri re-importu `.frm` uvek ide sa svojim `.frx` parom — `ImportAllVBA` preskače
formu bez `.frx` para.

## Dinamičke kontrole (bez `.frx`)

`Controls.Add` + WithEvents klasa: `clsBlokUI`/`modOtkupBlok`. Drugi primer je
bio `clsLookupMenuBtn`/`modMaticniLookups` — klasa je otišla u koraku 5 sa
`frmMaticniPodaci`, jedinom formom koja je taj meni gradila.

Za form-hostovane runtime kontrole koristi **generički `clsUiSink`** (`WireSink`
+ `UiSinkEvent` dispatcher; živ primer je `frmOtkupUI` sa `modOtkupUI`).

## NOVE `Private WithEvents` deklaracije u FORMAMA su ZABRANJENE

Dodavanje event-sink deklaracije u formu **lomi self-update code-merge te forme**
(krivac za crash 2.16.1 → 2.21.0; `docs/SELF_UPDATE.md` zamka #11).

Zatečeni pre-2.16.1 form-WithEvents su bili **zamrznuti** — spisak se samo
skraćivao, red po red, kako je koja forma odlazila: `frmDokumenta` (korak 2),
`frmAgrohemija` / `frmPalete` / `frmIzvestaj` (korak 3), `frmBankaExportPregled`
(korak 4).

**Spisak je sada PRAZAN.** Zabrana iznad time ne slabi nego postaje apsolutna:
nijedna preostala forma nema nijednu `WithEvents` deklaraciju, i tako mora da
ostane. Izuzetka više nema — svaki novi form-`WithEvents` je nov kvar, ne
nastavak zatečenog stanja.

Runtime kontrole u formama idu kroz omotače: `clsFlatBtn` (ljuska), `clsUiSink`
(generički sink), `clsBlokUI` (`modOtkupBlok`).

## Matični podaci (UI)

**Ljuska (važi):** ekrani `MAT_PARTNERI` / `MAT_ROBA` / `MAT_PAKOVANJE` /
`MAT_KORISNICI` (`modMaticni*` + `modScrMat*`) i paneli `MAT_PODESAVANJA` /
`MAT_ADMIN` (`modUiPanel`). Ulaz je prekidač sekcija u zaglavlju ljuske; pravo
na oblast `MaticniPodaci` odlučuje šta je od toga dostupno.

**Legacy je otišao u koraku 5** (§27.13): `frmMaticniPodaci`, `frmStammdaten`,
`clsStmBtn` i `clsLookupMenuBtn`. `modMaticniLookups` **ostaje** — ljuska koristi
`MaticniSekcije`, `MaticniSekcijeGrupisano` i `MaticniMenu_Release`; otišla je
samo njegova polovina koja je gradila stari meni.
