# Excel test harness — `tools/run_vba.py`

Headless `import src-vba/` → `Debug > Compile` → test suite, bez ijednog klika.
Radi **samo na Windows mašini sa instaliranim Excelom** (COM). Na Linux/macOS
nema `VBProject` interfejsa — tamo postoji samo `tools/vba_check.py` (statičke
provere, radi svuda).

## Jednokratna priprema

```bash
pip install pywin32
```

Excel: **File → Options → Trust Center → Trust Center Settings → Macro Settings**
→ čekiraj **„Trust access to the VBA project object model"**. Bez toga skripta ne
može da uveze module (`VBProject` je zaključan).

## Upotreba

```bash
python tools/run_vba.py --compile-only     # najbrže i najstabilnije
python tools/run_vba.py                    # + podrazumevani set suite-ova
python tools/run_vba.py --suite RunBankaImportTestSuite
python tools/run_vba.py --all
python tools/run_vba.py --workbook "C:\putanja\AgriX_OtkupApp.xlsm"
```

Izlazni kod: `0` = zeleno, `2` = palo (compile greška, pala `gate` suite, ili
neočekivan modalni dijalog).

## Koja sveska se koristi

| Slučaj | Sveska |
|---|---|
| bez `--workbook` | `tests/fixtures/otkup_test.xlsm` |
| tog fajla nema | skripta ga **sama napravi** kao praznu `.xlsm` |
| `--workbook PUT` | ta sveska |

Sveska se **uvek kopira u temp folder** pre rada — original se ne dira. Zato je
bezbedno proslediti i pravi radni `AgriX_OtkupApp.xlsm`.

`tests/fixtures/` je u `.gitignore` — fixture je lokalni artefakt, ne ide u repo.

### Zašto je prazna sveska dovoljna za compile

Nijedan modul ne referencira sheet `CodeName` rano-vezano — sve `SHT_*` konstante
u `modConfig` su **string literali** (`"sOtkup"`, `"sConfig"`…), a nijedan
`.doccls` nema `Public` član koji bi neko spolja zvao. Sheet `.doccls`-ovi se u
praznoj svesci uredno preskoče (`SKIP ... (nema komponente 'sOtkup')`) i compile
i dalje pokriva ceo `src-vba/`.

### Kada prazna sveska NIJE dovoljna

Test suite čitaju `tblOtkup`, `tblKooperanti`, `tblArtikli`… Nad praznom sveskom
padaju na podacima, ne na kodu. Za suite:

```bash
python tools/run_vba.py --suite RunIzvestajTests --workbook "C:\...\AgriX_OtkupApp.xlsm"
```

## Zašto skripta ne zove `modVbaTools.ImportAllVBA`

`ImportAllVBA` ima **hardkodiran folder**
(`C:\Users\Dusan\Documents\GitHub\otkupapp-pwa\src-vba\`) i završava se
`MsgBox`-om. Oboje je smrt za headless run — `MsgBox` u nevidljivom Excelu visi
zauvek (COM poziv ne puca, samo stoji).

Import se zato radi istom logikom preko COM-a, sa folderom izvedenim iz lokacije
skripte. Pravila su identična:

- `.doccls` se **merge-uje** u postojeću komponentu (ne briše se pa uvozi)
- `.bas` / `.cls` / `.frm` se uklone pa uvezu
- `.frm` mora imati `.frx` par, inače se preskače
- `modVbaTools` se preskače (on je taj koji se izvršava)

Modalne dijaloge zatvara watchdog nit (prozori klase `#32770` u Excel procesu,
klik na podrazumevano dugme). Svaki uhvaćen dijalog se prijavljuje u izveštaju —
neočekivan dijalog je nalaz, ne šum.

## Kako se compile stvarno meri (i zašto ne preko menija)

Prva verzija je zaključivala iz `Enabled` stanja stavke **Debug → Compile
VBAProject** („posle uspešnog compile-a postaje siva"). To **ne radi** u
nevidljivom Excelu: VBE osvežava enabled-stanje kontrola tek kad se meni iscrta,
pa stavka ostaje aktivna i kad je projekat uredno kompajliran. Rezultat je bio
`COMPILE NEJASNO` na svakom pokretanju — a, gore, `NEJASNO` je prolazilo kao
`REZULTAT: ZELENO`.

Druga verzija je verdikt tražila od **probe-a**: doda se modul sa funkcijom koja
vraća `42`, pozove se, i ako `Application.Run` prođe — projekat se kompajlira.
Ni to ne radi: **VBA kompajlira lenjo**. Prevede modul koji se zove i ono što taj
modul stvarno referencira. Probe modul ne referencira ništa, pa je namerno
ubačena greška u `modConfig` prošla netaknuta — treće lažno zeleno.

Ispostavilo se i zašto meni-compile nije prijavio grešku: **`Execute()` bez
vidljivog VBE prozora ne uradi ništa.** Kontrola „Compile VBAProject" je tada
disabled, `Execute` tiho prođe i nikakav dijalog se ne pojavi — pa je izostanak
dijaloga izgledao kao uspeh.

Sada:

1. **VBE prozor se otvara** (minimizovan) i projekat se postavlja kao aktivan —
   bez toga meni kontrola ne reaguje.
2. **`Enabled` stanje kontrole je verdikt.** `True` = projekat nije kompajliran,
   `False` = jeste. Pre compile-a mora biti `True` (tek smo uvezli module); ako
   nije, alat kaže `NEJASNO` umesto da tvrdi ishod.
3. **Dijalozi** koje watchdog uhvati sa „compile"/„error" u tekstu = `FAIL`.
4. **Probe** ostaje kao dodatna kontrola (hvata projekat koji uopšte nije
   izvršiv), ali više nije izvor verdikta.

Ispis uvek sadrži i sirovo stanje (`enabled_before`, `enabled_after`, `dialogs`,
`probe`), da se sledeći pogrešan ishod ne mora pogađati.

Pravilo: **sve što nije eksplicitno `OK` je pad** (izlazni kod 2). Alat koji ne
zna ishod mora to da kaže glasno, ne da ćuti u zeleno.

## `gate` vs „blind" suite

Katalog `SUITES` u skripti nosi zastavicu `gate`:

- `gate: True` — suite **podiže grešku** kad provera padne → runner je vidi kao crvenu
- `gate: False` — rezultat postoji samo u Immediate prozoru → runner je prijavljuje
  kao **`blind`**, što znači „prošla bez greške", a **ne** „sve provere prošle"

Kad pišeš novu suite, napravi je `gate` i upiši je u katalog. Puna tabela:
`.claude/rules/testovi.md`.

## Provera bez Excela

```bash
python3 tools/run_vba.py --self-test
```

Radi svuda (i u Claude Code sesiji na Linuxu). Proverava da strip VBA header-a ne
propušta header u kod — greška koja je jednom već prošla neopaženo i videla se
tek kao `[break]` u naslovu VBE prozora na Windows mašini.

## Ako zapne

- Skripta ima **tvrdi prekid** (`--timeout`, default 600 s): ako Excel prestane da
  odgovara, proces se ubija i run pada sa `FATAL`. Bez toga COM poziv ne puca —
  samo stoji zauvek.
- **Break mode** (`[break]` u naslovu VBE prozora) znači da je VBA stao usred
  izvršavanja. Najčešći uzrok je bio pokvaren import (header u kodu → naredba
  `End`), a drugi klik na **Debug** u dijalogu greške. Watchdog sada nikad ne
  klika `Debug` ni `Help` — bira `End` / `OK` / `Cancel`.
- `Ctrl+C`, pa proveri Task Manager za zaostali `EXCEL.EXE` — skripta pokreće
  zaseban proces (`DispatchEx`), pa ne dira tvoj otvoreni Excel.
- `--keep` ostavlja temp kopiju sveske da možeš da je otvoriš i vidiš stanje.
- Ako `import` prijavi puno `SKIP ... (nema komponente ...)` nad **pravom**
  sveskom — to znači da sveska nema te listove, tj. da nije AgriX radna sveska.
