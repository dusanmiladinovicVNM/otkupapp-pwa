---
paths:
  - "src-vba/frm*.frm"
  - "src-vba/cls*.cls"
  - "src-vba/modOtkupBlok.bas"
  - "src-vba/modMaticniLookups.bas"
  - "src-vba/modPaletniListUI.bas"
  - "src-vba/mod*UI.bas"
  - "src-vba/modUiFaze.bas"
  - "src-vba/modLogo.bas"
  - "tools/logo_to_vba.py"
---

# Forme, `.frx` i runtime kontrole

> Preseljeno iz `CLAUDE.md` §3/§4. Učitava se kad se dira UI sloj.

## Forma je JEDNA

`frmOtkupUI`. Nijedna druga ne postoji (§27.18). Plan §27 je bio ispunjen na
četiri forme (korak 7, §27.17); preostale tri su posle toga prešle u **faze**
istog prozora.

| Faza | Šta je bilo | Ko je gradi |
|---|---|---|
| `APP` | ljuska sa ekranima | `modOtkupUI` + `modScr*` |
| `BOOT` | `frmSplash` | `modUiFaze` |
| `LOGIN` | `frmLogin` | `modUiFaze` |
| `MINI` | `frmExcelMini` | `modUiFaze` |

**Nova forma se ne pravi.** Nov operaterski sadržaj je **ekran** (`modScr*`);
nešto što pokriva ceo prozor i nije navigaciono (nema stavku sidebara, oblast ni
pravo) je **faza** (`modUiFaze`). Treće ne postoji.

Ranije je ovde stajao izuzetak „osim kad §5 `otkup-i-dokumenta.md` izričito
traži" — postojao je zbog preslikavanja pravila unosa u formu koja ga je
duplirala i otišao je sa korakom 2.

### Šta faza mora da poštuje

Ovo nisu preporuke nego cena koju je §27.18 već platio:

1. **Tagovi kontrola faze počinju sa `fz`.** `modOtkupUI.UiEvent` ih prosleđuje u
   `modUiFaze.FazaEvent` **pre** nego što dodirne ijednu kontrolu ljuske — u fazi
   prijave ljuska još nije ni izgrađena.
2. **Ljuska se NE gradi dok traje prijava.** `BuildOtkupScreen` čita registar
   ekrana i prava operatera, a operatera tada nema; izgrađena ljuska bi dobila
   prava prazne sesije i zapamtila ih. Odluku nosi `FazaGradiLjusku`, gradnju
   `OtkupUI_EnsureShellBuilt`.
3. **Sinkovi faze preživljavaju gradnju ljuske** (`SinkoviFaze`). `BuildOtkupScreen`
   resetuje `Btns`; kad bi tu pale i kontrole kartice prijave, prva zamena
   operatera bi **visila zauvek** — petlju čekanja prekida baš klik na to dugme.
4. **Zone ljuske se GASE, ne prekrivaju** (`OtkupUI_ZoneUstupi`). Prekriven
   `Frame` i dalje prima `Tab`, pa bi se ispod kartice prijave moglo dotaći polje
   ekrana koji prijava još nije odobrila.
5. **Forma se između faza ne prikazuje ponovo**, pa `UserForm_Activate` ne puca —
   posao aktivacije radi `OtkupUI_AktivirajLjusku`, sa dva pozivaoca.

### Start i vidljivost sveske

`Application.Visible = False` je **prva naredba `Workbook_Open`-a**, a splash se
diže pre svih kapija. Sveska se otkriva na **tačno tri mesta**: odbijena kapija
(licenca / verzija / prijava — te grane same zovu `Visible = True` pa gase
aplikaciju), first-run setup (`SetupNewPC` bira foldere kroz `FileDialog`) i
dugme „Otvori Excel" iza `OBL_OTVORI_EXCEL`.

Kapija koja otkrije svesku pa ipak **propusti** (self-update na „Ne",
min-version na `WARN`) mora da bude praćena vraćanjem `Visible = False` u
`StartApp` — te kapije su u tuđim modulima i ne diraju se (`modSelfUpdate` je
frozen). Ceo lanac je u `docs/production-runbook-startup-autosave-journal.md` §2.2.

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
5. **Forma ljuske se ne penzioniše brisanjem nego pretvaranjem u fazu** (§27.18).
   Splash, prijava i mini kartica nisu bile legacy — nosile su svoju kopiju istog
   jezika (gradijent, zlatna nit, znak, `NewShell` polja, primarno dugme). Razlog
   zbog koga legacy odlazi važi i za njih: kopija istog pravila na četiri mesta
   se razilazi prvom doradom.

## `.frx` se ne dira kao tekst

Kontrole formi žive u binarnom `.frx` (`.frm` ima samo `OleObjectBlob`). Nove
kontrole se **NE** dodaju editovanjem teksta — dodaju se u runtime-u
(`Controls.Add` + WithEvents), kao postojeći `modOtkupBlok`/`clsBlokUI`.

Caption koji živi samo u `.frx` menja se u dizajneru forme, ne u kodu.
Pri re-importu `.frm` uvek ide sa svojim `.frx` parom — `ImportAllVBA` preskače
formu bez `.frx` para.

### Slike NE idu u `.frx` nego u `modLogo`

`.frx` **ne putuje kroz self-update** (`ImportFromFolder` uvozi kod), pa bi
svaka izmena znaka tražila REINSTALL na svakoj mašini. Zato je logotip
**generisan kod**: `src-vba/modLogo.bas`, Base64 GIF, dekodiran u privremeni
fajl i učitan `LoadPicture`-om.

- **Ne menja se rukom** — generiše ga `python tools/logo_to_vba.py` iz
  `img/AgriX-Otkup-Logo-Final.png`.
- **GIF, ne PNG:** `LoadPicture` čita BMP, RLE, ICO, WMF, EMF, GIF i JPEG —
  PNG ne.
- **Pozadina se peče u sliku**, jer MSForms ne zna per-pixel alfu. Ista boja
  izlazi kao `LOGO_BG_*` i boji ploču iza slike, pa se crtanje i slika ne mogu
  razići.
- **Crta se 1:1, ne skalira se.** `PictureSizeMode = Zoom` ide kroz `StretchBlt`
  sa `COLORONCOLOR` — bez uglačavanja, smanjivanje prosto **ispušta** redove i
  kolone. Piksel mera se računa iz okvira u tačkama; varijantu (1x / 2x) bira
  runtime (`LogoKljuc`).
- **Neuspeh je očekivan ishod, ne greška:** `LogoSlika` vraća `Nothing` kad nema
  MSXML/ADODB ili je TEMP nedostupan, i tada se crta tekstualni znak. Rezerva se
  ne briše.

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
