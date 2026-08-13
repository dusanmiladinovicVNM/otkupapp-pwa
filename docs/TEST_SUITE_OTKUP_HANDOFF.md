# Handoff: test suite nad .xlsm koja dokazano pada na pravom bugu

> Status: **nije započeto — čeka Windows sesiju.** Ovaj fajl je ulaz za tu sesiju.
> Nastao u Linux sesiji (`claude/test-suite-otkup-proof-lsx0ky`) u kojoj se
> kriterijum prihvatanja ne može izvršiti. Kod nije pisan namerno — vidi §1.

## 1) Zašto ovde nije rađeno

Kriterijum prihvatanja zadatka je **dvosmerni dokaz**: `python tools/run_vba.py`
→ `exit 0` nad čistim kodom, pa sabotaža → `exit 2` sa imenom baš tog testa u
ispisu, pa `git checkout` → ponovo `exit 0`. To traži Excel preko COM-a.

Izmereno u toj sesiji:

```
uname -a                             → Linux vm 6.18.5-fc-v20 x86_64
python3 -c "import win32com.client"  → ModuleNotFoundError: No module named 'win32com'
ls /mnt/c → nema           which excel → nema
```

Suite koja je zelena nad čistim kodom, a nije dokazano crvena nad pokvarenim, ne
dokazuje ništa — to je već četiri puta bio ishod u PR #181. Zato kod nije pisan
„naslepo" da ga neko drugi verifikuje; posao se nastavlja na mašini gde dokaz može
da se izvede.

## 2) Imena iz zadatka ≠ imena u kodu

Zadatak je pisan nad simbolima kojih u `src-vba/` nema. Provereno nad svih 192
fajla:

| Zadatak kaže | Stvarno |
|---|---|
| `src-vba/modOtkupUI.bas` | ne postoji |
| `modOtkupUI.ClearForm` | `ClearOtkupFields` — `frmOtkup.frm:1186`, `Private` |
| `CommitDokument` | `SaveOtkupMulti_TX`, zvano iz `btnUnos_Click` (`frmOtkup.frm:772`, `Private`) |
| `cbKupac` | `cmbKooperant` |
| `v6-ui-108` | nema takvog taga/commita ni u istoriji ni u `docs/` |
| `UI_MIGRACIJA_KATALOG.md` | ne postoji (faza 4 zadatka nema gde da upiše red) |
| `txtBrojZbirne`, `SortArray`, `tblLocalConfig`, `APP_SETUP_COMPLETED` | postoje |

**Ponašanje opisano u zadatku je tačno** — samo su imena druga. `ClearOtkupFields`
zaista ne dira `txtDatum`, ne dira `txtBrojZbirne`, i zaista radi
`cmbKooperant.value = ""` (linija 1194). Sva tri testa imaju smisla nad ovim kodom.

Napomena: repo je shallow (230 commita), pa odsustvo `v6-ui-108` nije dokaz da ga
nikad nije bilo.

## 3) Merenja koja menjaju obim isporuke

```
MsgBox: 976 poziva   InputBox: 49   FileDialog/GetOpenFilename: 5
fajlova koje bi pun wrapper dirao: 67      postojeći jedinstveni wrapper: nema ga
od 976 MsgBox: 58 troši povratnu vrednost, 59 koristi vbYesNo/vbOKCancel
```

Zato `If gTestMode Then Exit Sub` (naredbeni oblik) ne pokriva sve — za tih ~59
treba *Function* wrapper sa podrazumevanim odgovorom. Pun wrapper nad 67 fajlova je
direktno protiv ograničenja „`src-vba` se menja samo tamo gde zadatak traži".

## 4) Odluke operatera (donete, ne otvarati ponovo)

1. **Sesija se seli na Windows** (Excel + pywin32). Dokaz radi ta sesija.
2. **Public seam u `frmOtkup.frm` je dozvoljen.** `ClearOtkupFields` →
   `Public Sub` (ili tanak `Public Sub Test_*` omotač). Bez toga `modTest` ne može
   da ga pozove i tri testa ne postoje. Ovo je izuzetak od „ništa drugo".
3. **`gTestMode` gard je uzak** — samo pozivi koje tri testa stvarno pogode, uz
   postojeći watchdog u `run_vba.py` koji zatvara modalne dijaloge. Bez refaktora
   976 mesta.

## 5) Recepture sabotaže — ispravljene na stvarna imena

Svaki test mora da se pokaže u **oba** smera. Revert je uvek
`git checkout -- src-vba/frmOtkup.frm` (**ne** `modOtkupUI.bas`).

| Test | Sabotaža u `ClearOtkupFields` (`frmOtkup.frm:1186`) |
|---|---|
| `T_PosleSnimanja_ZadrzavaKontekstOtpremnice` | dodaj `txtDatum.value = ""` |
| `T_PosleSnimanja_ZadrzavaZbirnu` | dodaj `txtBrojZbirne.value = ""` |
| `T_ClearForm_BrisePartnera` | ukloni liniju 1194 `cmbKooperant.value = ""` |

Očekivano po sabotaži: `exit 2` **i ime baš tog testa kao FAIL u ispisu**. Exit 2
bez imena testa (npr. compile pukao) ne važi kao dokaz.

`.frm` se menja samo u kodnom delu; `.frx` se ne dira i ide u paru.

## 6) Ostaje da se napravi

Isporuke iz zadatka, sa korekcijama iz §2–§4:

1. Uzak `gTestMode` gard (§4.3).
2. `tools/make_fixture.py` → `tests/fixtures/otkup_test.xlsm`.
   `tests/fixtures/` je već u `.gitignore:12`; `tests/golden/` **nije** ignorisan —
   golden fajlovi idu u repo na pregled, to je namerno.
3. `src-vba/modTest.bas` (ASCII): `RunAllTests`, `AssertEq`, `DumpKontrole(f)`
   (koristi postojeći `modArrayUtils.SortArray`, ne pisati nov sort),
   `AssertSnapshot`.
4. Tri testa iz §5. Realan oblik: postavi vrednosti kontrola → pozovi Public
   `ClearOtkupFields` → proveri. Vožnja punog `btnUnos_Click` povlači stanica-lock,
   Google sync, PDF i auto-hladnjaču; a sve tri sabotaže ionako ciljaju
   `ClearOtkupFields`, pa je to tačno mesto gde bug živi.
   Instanciranje bez prikaza: `Set f = New frmOtkup`, pa odmah `f.Controls.Count`
   (bez toga se `Initialize` ne okine). Bez `.Show`.
5. `run_vba.py` — dopuna, ne prepisivanje. `last_run.txt` već postoji
   (`tools/run_vba.py:768`). Nedostaje: poziv `RunAllTests`, `FAIL>0` → exit 2,
   `last_run.txt` ne postoji → exit 2 (ne 0), golden fajlovi u temp i nazad.
6. `Stop` hook — tek posle dokazana oba smera. Dotle se skripta zove ručno.
7. `.claude/rules/testovi.md`: nova suite mora biti `gate` i upisana u `SUITES`
   katalog (`tools/run_vba.py:55`). „blind" suite je test koji niko neće videti
   kad pukne.

## 7) Već dokazano u PR #181 — ne ponavljati

- `Application.Run` NE kompajlira ceo projekat (VBA kompajlira lenjo).
- `FindControl(Id=578).Execute()` bez aktivnog VBE prozora ne uradi ništa.
- `Enabled` stanje te kontrole ima dva uzroka i ne razlikuje ih — nije verdikt.
- Header `.doccls` fajla se skida pozicijski, ne po obliku linije.
- Watchdog nikad ne klikće „Debug".

Compile signal ne treba goniti posebno: `Run("RunAllTests")` kompajlira `modTest` i
sve što on referencira, a to je baš kod pod testom.
