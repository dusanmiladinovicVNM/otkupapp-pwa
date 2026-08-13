# `Ensure*` funkcije — analiza i plan profesionalizacije

> Status: analiza (bez izmena koda). Snimljeno stanje: grana
> `claude/ensure-functions-analysis-n9c81u`.
> Metod: pun inventar `src-vba/` (`.bas` + `.frm`) + `src/js/` + `gas/`, čitanje
> tela svake funkcije, provera pozivalaca i kontrole grešaka.

---

## 1) Inventar

**61 `Ensure*` procedura u `src-vba/`** (29 `Public`, 32 `Private`), raspoređenih
u 18 modula/formi. Uz njih još **6 `ensure*` u `src/js/`** i **5 u `gas/`**.

Raspodela po obradi grešaka (VBA):

| Obrada greške | Broj | Ishod |
|---|---:|---|
| `On Error GoTo EH` + interni `Resume Next` | 20 | uglavnom `LogErr` |
| `On Error GoTo EH` | 12 | `MsgBox` / `Err.Raise` / `LogSetup` |
| samo `On Error Resume Next` (tiho) | 12 | **bez ishoda (20x `-`)** |
| bez ikakve obrade | 11 | — |
| `On Error GoTo <drugi label>` | 6 | `LogErr` |

Ishod na grešku: 27x `LogErr`, 5x `Err.Raise`+`LogSetup`, 4x `MsgBox`+`LogSetup`,
3x samo `LogSetup`, **20x ništa**.

---

## 2) Taksonomija — jedno ime, šest različitih ugovora

Ovo je centralni nalaz. `Ensure` u ovom kodu znači šest različitih stvari, sa
šest različitih ugovora o grešci i vidljivosti:

| # | Familija | Broj | Šta zapravo radi | Ugovor danas |
|---|---|---:|---|---|
| **A** | Šema tabela/kolona | 20 | DDL nad `ListObject` | mešano: `Raise`, `MsgBox`, tiho |
| **B** | Folderi / FS | 5 | `CreateFolder` / `MkDir` rekurzivno | tiho + `LogSetup` |
| **C** | Excel šabloni (`*Sablon`) | 13 | kreira/regeneriše radni list | `LogErr`, best-effort |
| **D** | Runtime UI kontrole | 17 | `Controls.Add` na formi (`.frx` se ne dira) | `LogErr`, best-effort |
| **E** | Poslovni seed/mirror | 2 | **upisuje poslovne podatke** | 1x `Raise`, 1x `LogErr` |
| **F** | Assertion (pogrešno ime) | 1 | **ne menja ništa, samo puca** | `Err.Raise` |
| **G** | Remote (Google Sheets) tabovi | 3 | best-effort mrežni poziv | tiho |

**Zašto je to problem u praksi:** pozivalac po imenu ne može da zna
(a) da li poziv može da pukne, (b) da li može da otvori `MsgBox` usred toka,
(c) da li menja podatke ili samo proverava. Zbog toga pozivaoci defanzivno
obmotavaju sve u `On Error Resume Next` — što onda guta i one greške koje su
stvarno važne.

### Konkretno: familija F je pogrešno imenovana

`modPaletniList.EnsurePrijemnicaNotAlreadyPaletized` (`modPaletniList.bas:130`)
**ne obezbeđuje ništa** — čita `tblPaleteStavke` i puca ako nađe aktivnu stavku.
To je assertion. Projekat **već ima konvenciju za to**: familija `Require*`
(20+ funkcija — `RequireSingleRow`, `RequireColumnIndex`, `RequireBimSmer`,
`RequireFakturaZaKupca`, `RequireAmbalazaSchema`…). Ime je jedini problem.

---

## 3) Nalazi (rangirano po riziku)

### N1 — `EnsureRuntimeSchema` guta svaku grešku na svakom startu ⚠️ visok

`modSetup.bas:1135-1169`: `On Error Resume Next` preko celog tela, 12 koraka
(kolone, formati, `EnsureStornoVezeSchemaCore`, `EnsureStornoZurnalSchemaCore`,
`EnsureSledljivostSchema`), **nijedan `LogSetup`, nijedan per-korak check**.

Pozivalac `modMain.InitApp` (`modMain.bas:211-216`) ima naizgled zaštitu:

```vba
On Error Resume Next
EnsureRuntimeSchema
If Err.Number <> 0 Then LogErr "modMain.InitApp.EnsureRuntimeSchema": Err.Clear
```

Ta provera je nepouzdana: `Err` se u VBA resetuje na izlasku iz procedure, a i u
najboljem slučaju bi se video samo *poslednji* error — bez podatka **koji** od 12
koraka je pao. Praktična posledica: ako self-heal padne (zaključana tabela,
zaštićen sheet, tabela obrisana), aplikacija se digne bez kolone i bez ijednog
traga; greška isplivava kasnije, kao pogrešan podatak u dokumentu.

Isto važi za `EnsureSledljivostSchema` (`modSetup.bas:1180-1197`).

### N2 — Agregat „Ensure (setup + šeme)" laže o uspehu ⚠️ visok

`modAdmin.AdminEnsureEverything` (`modAdmin.bas:271-282`) zove 7 koraka i na
kraju uvek javi *„Ensure završen (setup + sve šeme provereno)."*

Ali koraci imaju **nekompatibilnu semantiku greške**:

| Korak | Na grešku |
|---|---|
| `EnsurePoruke`, `EnsureCenovnikSchema`, `EnsureKorisniciSchema` | `Err.Raise` → **prekida ceo agregat** |
| `EnsurePaletniListSchema`, `EnsureDoradeSchema`, `EnsureAuditColumns` | `MsgBox` pa **normalan povratak** → agregat nastavlja i prijavi uspeh |

Dakle: pad `EnsureDoradeSchema` završi sa dva dijaloga — prvo „Greška…", pa
odmah „Ensure završen, sve provereno". Pad `EnsureCenovnikSchema` tiho preskoči
preostala 4 koraka.

Uzgred: `EnsureCenovnikSchema` se u toj putanji izvršava **dvaput** —
direktno (`modAdmin.bas:275`) i iz `EnsurePaletniListSchema` (`modSetup.bas:959`).

### N3 — Agregat ispali 4 `MsgBox`-a ⚠️ srednji

`EnsurePaletniListSchema`, `EnsureDoradeSchema`, `EnsureAuditColumns` svaki nosi
svoj `MsgBox` uspeha, plus finalni agregatni. Admin klikne jednom → klikće četiri
puta. Uzrok: `MsgBox` je u *jezgru* umesto u ulaznoj tački.

Projekat je već otkrio pravi obrazac — `EnsureAuditColumns` (`MsgBox`) /
`EnsureAuditColumnsCore` (tiho, vraća broj), i `EnsureStornoVezeSchema` /
`…SchemaCore`. Samo nije primenjen svuda: `EnsurePaletniListSchema` i
`EnsureDoradeSchema` nemaju `Core` blizanca, pa su neupotrebljivi iz koda.

### N4 — Primitivi su `Private`, pa su forkovani ⚠️ srednji

`EnsureColumnOnTable` (`modSetup.bas:1618`) i `EnsureDataTable`
(`modSetup.bas:1580`) su `Private`. Posledica — doslovan re-implement:

| Fork | Original | Razlika |
|---|---|---|
| `modPaletniList.EnsurePreradaCol` (`:2445`) | `modSetup.EnsureColumnOnTable` | **nema** — isti kod, 7 linija |
| `modPaletniList.EnsurePreradaCols` (`:2433`) | ono što radi `EnsureDataTable` nad postojećom tabelom | ručni spisak 6 kolona |
| `modBankaImport.EnsureFolderExists` (`:1153`) → `BankaEnsureFolderExistsRecursive` (`:1659`) | `modSetup.EnsureFolder` (`:1718`) | banka varijanta ima `BankaNormalizeFolderPath` (Drive virtuelne putanje); detekcija je u oba FSO `FolderExists` |

Komentar u `modPaletniList.bas:2431` to i priznaje: *„Rešava 0 u sažetku paletnog
lista kada `EnsurePaletniListSchema` nije pokrenut posle nadogradnje."* — tj.
fork postoji jer se centralni Ensure nije mogao pozvati odatle.

### N5 — Šabloni: dve legitimne podvrste pod istim imenom ⚠️ nizak-srednji

13 `Ensure*Sablon` funkcija se deli na dva stvarno različita mehanizma:

- **Verzionisani layout** (8): `FakturaSablon`, `KarticaSablon`,
  `SledljivostSablon`, `KarticaAmbalazeSablon`, `SpecifikacijaSablon`,
  `IsplataSpecSablon`, `PaletaSablon` (v3), `PreradaSablon` (v5) — marker u
  `H1`/`N1`, stara verzija se briše i pregrađuje. **Auto-upgrade radi.**
- **Prazno platno** (5): `Otpremnica`, `Otkup`, `GrupniOtkup`, `Prijemnica`,
  `IzdavanjeAmbalaze` — `If Not ws Is Nothing Then Exit Sub`, samo `Sheets.Add` +
  širine kolona; sadržaj crta `Fill*` pri svakoj štampi.

Za drugu grupu ovo **nije bug** (sadržaj se ionako pregrađuje), ali jeste latentna
zamka: **širine kolona se postavljaju samo pri kreiranju**. Promena
`ws.columns("B").ColumnWidth` u kodu nikad ne stigne do postojećih instalacija.
Iz imena se ne vidi kojoj grupi funkcija pripada.

### N6 — Familija E menja poslovne podatke pod imenom „Ensure" ⚠️ nizak

`modMalina.EnsureVozacMirrorForStanica` (103 linije, `Public`) i
`modAgrohemija.EnsureArtikalPocetniDug` upisuju redove u `tblVozaci` / `tblArtikli`.
Prvi je već prošao korekciju (AUD-046: re-raise umesto gutanja, komentar
`modMalina.bas:58-64`) — što potvrđuje da ime nije komuniciralo težinu operacije.

### N7 — UI familija: dva različita guard obrasca ⚠️ nizak (kozmetika)

- boolean flag: `If m_undoOpsBuilt Then Exit Sub` … `m_undoOpsBuilt = True`
  (`frmDokumenta`, 6 panela)
- null-check objekta: `If Not mCmbKooperant Is Nothing Then Exit Sub`
  (`frmBankaExportPregled`, `frmIzvestaj`)

Oba rade. Boolean flag traži dodatno modul-level stanje i može da se raziđe sa
stvarnošću (flag `True`, kontrola `Nothing` posle greške na pola izgradnje —
`EnsureUndoOpsPanel` postavlja flag tek na kraju, pa je tu OK, ali obrazac nije
zaštićen pravilom).

### N8 — Nema automatske provere ⚠️ nizak

`tools/vba_check.py` proverava ASCII, deklaracije, rezervisane reči, duplikate i
nedefinisane simbole — ali ništa o `Ensure*` ugovoru. Testovi dodiruju samo 4
funkcije (`EnsureVozacMirrorForStanica`, `EnsurePoruke`,
`EnsureStornoVezeSchemaCore`, `EnsureStornoZurnalSchemaCore`); ostalih 57 nema
nijedan test — uključujući idempotenciju, koja je njihova jedina obećana osobina.

---

## 4) Predlog profesionalizacije

Princip: **ne uvoditi novi sloj — formalizovati onaj koji kod već ima.**
Sve četiri konvencije ispod već postoje u `src-vba/`, samo se ne primenjuju
dosledno.

### 4.1 Ugovor po prefiksu (jedno pravilo, četiri imena)

| Prefiks | Ugovor | Sme `MsgBox`? | Sme da puca? | Menja stanje? |
|---|---|---|---|---|
| `Ensure*` | idempotentno dovede u željeno stanje; no-op ako već jeste | **NE** | da, `Err.Raise` uz `SRC` | da |
| `Require*` | provera preduslova | **NE** | **da, to mu je posao** | **NE** |
| `Setup*` / `Admin*` / `Run*` | interaktivna ulazna tačka (Alt+F8 / dugme) | **DA** | ne — hvata i prikazuje | delegira `Ensure*` |
| `Fill*` / `Build*` | render u već obezbeđen kontejner | ne | da | ne dira šemu |

Posledica: **`*Core` sufiks nestaje kao pojam.** `EnsureXCore` postaje `EnsureX`
(tiho jezgro), a `MsgBox` omotač se zove `SetupX`. Time
`EnsureAuditColumns`/`EnsureAuditColumnsCore` postaje
`SetupAuditColumns`/`EnsureAuditColumns`.

Migracija imena (5 preimenovanja, mehanička):

| Sada | Posle | Razlog |
|---|---|---|
| `EnsurePrijemnicaNotAlreadyPaletized` | `RequirePrijemnicaNotPaletized` | assertion, ne mutacija (N-F) |
| `EnsureAuditColumns` / `…Core` | `SetupAuditColumns` / `EnsureAuditColumns` | MsgBox u ulaznu tačku |
| `EnsureStornoVezeSchema` / `…Core` | `SetupStornoVezeSchema` / `EnsureStornoVezeSchema` | isto |
| `EnsurePaletniListSchema` | `SetupPaletniListSchema` + novo tiho `EnsurePaletniListSchema` | agregat mora da ga zove bez dijaloga (N2/N3) |
| `EnsureDoradeSchema` | `SetupDoradeSchema` + tiho `EnsureDoradeSchema` | isto |

`EnsureVozacMirrorForStanica` / `EnsureArtikalPocetniDug` (familija E) **ostaju** —
posle preimenovanja `Ensure*` znači „idempotentna mutacija", što je tačno ono što
one rade; a `EnsureVozacMirrorForStanica` već ima ispravan ugovor greške (AUD-046).

### 4.2 Jedan sloj primitiva (ubija forkove iz N4)

`EnsureDataTable`, `EnsureColumnOnTable` i `EnsureFolder` iz `Private` → `Public`
u `modSetup` (tabela u `CLAUDE.md` §3 ionako već kaže da šeme žive tamo — dakle
nema novog modula, nema nove apstrakcije).

Zatim:
- `modPaletniList.EnsurePreradaCol` → **obrisati**, zvati `EnsureColumnOnTable`
- `modPaletniList.EnsurePreradaCols` → zadržati kao tanak spisak kolona, telo
  delegira `EnsureColumnOnTable`
- `modBankaImport.EnsureFolderExists` → **obrisati**; `BankaNormalizeFolderPath`
  ugraditi u `modSetup.EnsureFolder` (Drive-safe varijanta je bolja, ne obrnuto),
  pozivi idu na `EnsureFolder`

Provera posle: `python3 tools/vba_check.py` (DUPLIKAT check hvata sudare `Public`
imena; `Ensure*` imena su jedinstvena, sudar se ne očekuje).

### 4.3 Rezultat umesto tišine (ubija N1)

`Ensure*` jezgra vraćaju rezultat umesto da ćute. Minimalna varijanta, bez novog
tipa — po uzoru na već postojeći `EnsureAuditColumnsCore() As Long`:

```vba
' Svaki korak: greska se hvata i LOGUJE, tok se nastavlja, broj padova se vraca.
Public Function EnsureRuntimeSchema() As Long
    Dim fails As Long
    fails = fails + EnsureStep("KulturePrag", ...)
    ...
    If fails > 0 Then LogSetup "ERROR", "EnsureRuntimeSchema: " & fails & " koraka palo"
    EnsureRuntimeSchema = fails
End Function
```

Ako se `EnsureStep` helper smatra novim slojem (pravilo „rule of three"),
jeftinija varijanta bez ijedne nove funkcije — per-korak `Err.Number` provera po
obrascu koji `modMain.InitApp` već koristi:

```vba
On Error Resume Next
Err.Clear: EnsureColumnOnTable TBL_KULTURE, COL_KUL_PRAG_PROSEK_UPOZ
If Err.Number <> 0 Then LogSetup "ERROR", "RuntimeSchema/PragUpoz: " & Err.description
```

To je isti obrazac koji `modMigracija.bas:106-107` već primenjuje na
`EnsureAuditColumnsCore`. **Reuse, ne novo.**

### 4.4 Vidljivost u familiji C (N5)

Iz imena mora da se vidi mehanizam. Dve opcije, po ceni:

- **jeftino (preporuka):** zadržati imena, dodati obavezan blok-komentar iznad
  svake `Ensure*Sablon` sa jednom od dve oznake — `' LAYOUT: verzionisan (H1=n)`
  ili `' LAYOUT: prazno platno, sadrzaj crta Fill*`. Nula rizika, rešava čitljivost.
- **skuplje:** i pet „praznih platna" dobiju `LAYOUT_VER` marker, pa promena
  širina kolona stigne do postojećih instalacija. Cena: brisanje i pregradnja
  lista pri prvom startu posle nadogradnje na 5 dokumenata — treba smoke-test po
  dokumentu. Uraditi **samo ako** se širine zaista menjaju.

### 4.5 Zaštita od regresije (N8)

Dva jeftina koraka:

1. **`tools/vba_check.py` — novo pravilo `ENSURE`:** `Public Sub Ensure*` u
   `.bas` ne sme da sadrži `MsgBox`. Statički proverljivo, hvata tačno N3, i
   sprečava povratak obrasca. (~20 linija, po uzoru na postojeći `check_reserved`.)
2. **Test idempotencije:** jedan test u `modBusinessFlowProTests` koji dvaput
   zove svako tiho `Ensure*Schema` jezgro i tvrdi da drugi poziv ne menja broj
   kolona. Idempotencija je jedina osobina koju sve ove funkcije obećavaju, a
   nijedna je ne dokazuje.

---

## 5) Plan po fazama (minimalan delta prvo)

| Faza | Sadržaj | Fajlovi | Rizik | Vrednost |
|---|---|---|---|---|
| **F1** | N1 + N2 + N3: per-korak logovanje u `EnsureRuntimeSchema`/`EnsureSledljivostSchema`; `MsgBox` iz jezgara u `Setup*` omotače; `AdminEnsureEverything` skuplja i prijavljuje stvarni rezultat; ukloniti dupli `EnsureCenovnikSchema` | `modSetup`, `modAdmin` | nizak | **visoka** — kraj tihim padovima i lažnom „sve OK" |
| **F2** | N4: primitivi `Public`, brisanje 3 forka | `modSetup`, `modPaletniList`, `modBankaImport` | nizak-srednji | visoka — anti-duplication |
| **F3** | 4.1 preimenovanja (5 komada) + `RequirePrijemnicaNotPaletized` | ~6 fajlova | nizak (mehanički, `vba_check` hvata promašaje) | srednja — čitljivost ugovora |
| **F4** | 4.5: `ENSURE` pravilo u checkeru + test idempotencije | `tools/vba_check.py`, `modBusinessFlowProTests` | nizak | srednja — sprečava povratak |
| **F5** *(opciono)* | 4.4 skuplja varijanta; deklarativni registar šeme (tabela → kolone → format) koji vozi sve `Ensure*Schema` | `modSetup` + `modConfig` | **srednji-visok** | visoka, ali je to redizajn — tek kad F1–F4 legnu |

F5 je namerno poslednja i označena kao opciona: `AuditableTables()`
(`modSetup.bas:1090`) je već mikro-registar i pokazuje da obrazac radi, ali
prevođenje svih 8 `Ensure*Schema` na deklarativni model je idealizovan redizajn,
ne minimalna izmena — protiv default stava iz `CLAUDE.md`.

---

## 6) Šta NE dirati

- **`.frx` i `Private WithEvents`** — familija D (17 funkcija) postoji upravo zato
  što se kontrole dodaju u runtime-u. Taj obrazac je ispravan i ostaje.
- **`EnsureVozacMirrorForStanica`** — ugovor greške je već saniran (AUD-046) i
  pokriven testovima; samo se uklapa u novu konvenciju, ne prepisuje se.
- **`gas/` i `src/js/` `ensure*`** — druga sredina, druge norme (JS `ensure*` je
  ustaljena konvencija). Van opsega ove sanacije; jedina napomena je da
  `ensureMasterSyncNotActive` (`src/js/utils/master-sync-guard.js:224`) vraća
  `false` umesto da baca — što je po gornjoj taksonomiji zapravo `Require*`
  semantika sa soft ishodom, i to je za PWA u redu.
- **Familija B izvan `modSetup`** — `modJournaling`, `modLogError`,
  `modSelfUpdate` imaju svoje `MkDir` pozive, ali oni nisu `Ensure*` funkcije i
  nisu deo ovog opsega.

---

## 7) Sažetak u tri rečenice

`Ensure*` je u ovom kodu postao skupno ime za šest različitih ugovora, pa
pozivalac ne zna da li poziv puca, ćuti ili otvara dijalog — i defanzivno guta
sve. Najskuplja posledica je da `EnsureRuntimeSchema` na svakom startu tiho
proguta svaki pad self-heal-a, a admin agregat prijavi „sve provereno" i kad nije.
Profesionalizacija ne traži novi sloj: dovoljno je formalizovati četiri prefiksa
koje kod već koristi (`Ensure` / `Require` / `Setup` / `Fill`), izbaciti `MsgBox`
iz jezgara, otvoriti tri primitiva da forkovi nestanu, i to zaključati jednim
pravilom u `vba_check.py`.
