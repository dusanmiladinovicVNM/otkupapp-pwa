---
paths:
  - "src-vba/**"
---

# VBA izvor — pravila koja obaraju compile ili import

> Preseljeno iz `CLAUDE.md` §4. Ovo su greške koje CI ne vidi (VBA se ovde ne
> kompajlira) — hvata ih `python3 tools/vba_check.py` i, na kraju, operater u VBE-u.

**Posle svake izmene u `src-vba/`:** `python3 tools/vba_check.py` (exit 0 = čisto).
Isti checker se vrti kao PostToolUse hook nad izmenjenim fajlom, pa nalaz stiže
odmah — ne ignoriši ga.

## 1) Izvor mora ostati 100% ASCII

Posle lokalizacije (`1jj9xw` / v2.6) svi `.bas`/`.cls`/`.frm`/`.doccls` su ASCII i
MORAJU ostati ASCII. Zato je Edit/Write na njima bezbedan (latin-1 round-trip više
nije potreban), ali samo dok ne upišeš ne-ASCII bajt.

- **NIKAD** ne piši `š ž č ć đ Š Ž Č Ć Đ`, nemačke `ä ö ü ß`, ni tipografske
  `— « » • „ "` u VBA izvor. Fajl time postaje UTF-8/mešan i `ImportAllVBA` ga
  učita kao smeće (ista klasa greške kao `f08a0ee`).
- **Korisnički tekst sa dijakritikom ide ISKLJUČIVO kroz katalog:** red u
  `modPoruke.UpsertPoruke`
  (`UpsertRow lo, existing, "KLJUC", "Gre" & ChrW(353) & "ka..."`), a na mestu
  prikaza `Poruka("KLJUC")`. Dijakritika nastaje tek u runtime-u.
- **NE „sređuj radi čitljivosti"** vraćanjem `ChrW` u literal
  (`"Gre" & ChrW(353) & "ka"` → `"Greška"`) — to je tačno reintrodukcija greške.
- ChrW kodovi: `š=353 Š=352 ž=382 Ž=381 č=269 Č=268 ć=263 Ć=262 đ=273 Đ=272` ·
  em-dash `—=8212` · `«=171 »=187 •=8226`. Interne/nemačke (dev) stringove
  transliteruj u ASCII (`ü→ue ö→oe ä→ae ß→ss`), ne u `ChrW`.
- `.frx` ostaje binarni Windows-1250 — **ne dira se kao tekst**.
- Verifikacija: `file <fajl>` = „ASCII text"; `tools/vba_check.py` proverava i
  ASCII i orphan `Poruka("KLJUC")` ključeve; posle importa `Alt+F8 → EnsurePoruke`.
- Ovo važi SAMO za VBA izvor. `.md` / `.js` / ostali UTF-8 fajlovi u repou nemaju
  ovo ograničenje — tamo je Edit/Write bezbedan sa dijakritikom (nema konverzije).
- Prelazno: ako `file` na nekom VBA fajlu i dalje kaže „Non-ISO extended-ASCII"
  (nije transliterovan), za njega važi STARO pravilo (latin-1 round-trip, vidi git
  istoriju `CLAUDE.md`) dok se ne prebaci na ASCII.

## 2) Modul-level deklaracije idu u deklaracionu sekciju

Vrh modula, posle `Option Explicit`, **pre prve procedure**: `Public`/`Private
Const`, `Public`/`Private` promenljive, `Declare`, `Type`, `Enum`.

VBA **ne kompajlira** `Const` ubačen između dve procedure — a to je prirodno mesto
na koje padne kad se konstanta piše „uz funkciju koja je koristi" (RF-07: `IZV_TAB_*`
iznad `IzvestajTabDostupan`, na sredini `modIzvestaj`). Konstante grupiši uz
postojeće na vrhu i objasni ih komentarom tamo.

## 3) Rezervisane reči — VBA je case-insensitive

`Dim eNum As Long` = `Enum` → compile error (RF-06). Za EH varijable koristi
konvenciju projekta: **`errNum` / `errDesc` / `errSrc`** (`modStorno.LogAndReraise`,
`modAgrohemija`, `modBankaImport`), ne izmišljaj `eNum`/`eSrc`.

`tools/vba_check.py` proverava compile-hard podskup (ključne reči + imena tipova).
Šira lista iz starog `CLAUDE.md` (`name`, `line`, `text`, `date`, `base`, `time`,
`mid`, `local`, `read`…) je **stilska** — zatečeni kod ih koristi i kompajlira se;
ne menjaj ih usput.

## 4) Duple `Public` definicije = „Ambiguous name"

Čist git-merge može dati VBA compile grešku (dupli `Public` `Sub`/`Function`/`Const`).
Posle merge-a: `Debug → Compile VBAProject`. Statički to hvata `vba_check.py`
(DUPLIKAT) — samo nad `.bas`, jer `Public` član forme/klase nije globalno ime.

**Jedan izuzetak: ugovor ekrana novog UI-ja.** Ljuska `modOtkupUI` ne poznaje
nijedan ekran po imenu — svaki `modScr*` modul implementira isti skup procedura
(`Scr_Meta`, `Scr_Rows`, `Scr_Event`…), a ljuska ih zove **isključivo kasno
vezano i kvalifikovano** (`Application.Run "modScrDokumenti.Scr_Rows"`). Poziv
nikad nije nekvalifikovan, pa VBA nema šta da razrešava i „Ambiguous name" ne
nastaje. Spisak je `SCR_UGOVOR` u `vba_check.py`.

Deo ugovora je **neobavezan** — ekran koji ga ne implementira ponaša se kao pre:
`Scr_Dozvoljen` (dodatna brana), `Scr_Sort`, `Scr_ImaNesacuvano`,
`Scr_Deaktiviraj` (izlazak) i `Scr_Aktiviraj` (ulazak). Kad dodaješ nov član,
dodaj ga i u `SCR_UGOVOR`, inače `DUPLIKAT` prijavi drugi `modScr*` koji ga
implementira.

Izuzetak važi **samo kad su svi definicioni fajlovi `modScr*`**. Ista procedura
u bilo kom drugom modulu i dalje pada — što je i bio smisao provere. Dokazano u
oba smera: `Scr_Rows` prekopiran u `modUiData` puca, obična dupla definicija u
dva `modScr*` modula puca, čist kod je zelen.

## 4a) Duplo ime unutar JEDNOG modula — drugi simptom, druga provera

`DUPLIKAT` gleda **globalni** imenski prostor. Dva ista imena u **istom** modulu
su mu nevidljiva, a obaraju compile isto tako — i javljaju se drugačije:

> Modul koji se ne kompajlira obara **ceo projekat**, pa greška stigne kao
> **`Cannot run the macro`** na bilo kom makrou, ne kao „Ambiguous name". Simptom
> ne pokazuje na krivca; izgleda kao da je pukao harness ili instalacija.

Hvata to `DUPLIKAT_LOKALNI` (radi i nad `.frm`/`.cls` — unutar modula je sudar
sudar bez obzira na vrstu fajla). Jedini izuzetak je `Property Get/Let/Set`
trojka nad istim imenom. Ne gleda `Const`/`Dim` **unutar** procedure — isto ime u
dve procedure je legalno i uobičajeno.

Najčešći ulaz nije merge nego **neuspeo pokušaj izmene**: python heredoc koji je
„pukao" već je upisao konstantu, pa je `Edit` doda još jednom.

## 5) Ostalo

- **Ne zaključuj iz par linija.** Logika je raspoređena kroz module/forme/klase/
  evente — traži pun set relevantnih (`frm*`, `mod*`, `cls*`, `ThisWorkbook`) pre
  procene reuse-a / refaktora.
- **Pozicijski `AppendRow` zavisi od redosleda kolona** — bezbedan samo ako je
  redosled potvrđen. Inače upis **po imenu** (`UpdateCell`/`GetColumnIndex`).
  Detalji o šemi: `.claude/rules/podaci-i-config.md`.
