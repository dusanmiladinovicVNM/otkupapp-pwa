# Main review — AMB-10b trenutni `main`

Datum: 2026-10-02  
Repo: `dusanmiladinovicVNM/otkupapp-pwa`  
Pregledani head: `25e482c1df45809f2198f913c65fc88c00139f92`  
Merge: PR #400 — `feat(ambalaza): AMB-10b-1 -- knjiga dobija pisca, bez cutovera`

## Executive summary

Trenutni `main` je značajno pomeren u odnosu na prethodni pregled. Fokus više nije primarno self-update / banka / UI, nego **AMB-10b pisac za append-only knjigu ambalaže**.

Ukupni verdict:

```text
P0: 0

P1 otvoreno:
1. `modStorno` rollback failure se i dalje guta.
2. `modBankaMapiranje` rollback failure se i dalje guta.
3. `modMain.InitApp` i dalje hardkoduje `Application` state.
4. `modAmbalaza.PrenesiAmbalazu` je public multi-write pisac bez sopstvenog TX wrappera — nije bug danas jer još nema cutover, ali mora biti gate pre prvog produkcionog poziva.

P2:
1. Windows-only zavisnosti ostaju.
2. Storno public non-TX core funkcije i dalje postoje.
3. AMB-10b reader/writer performance kasnije izmeriti na većoj knjizi.
```

Novi AMB-10b sloj je domen-modelski jak. Nema razloga vraćati ovaj `main`. Sledeći stvarno opasan trenutak je **10b-2 cutover**, kada se `PrenesiAmbalazu` prvi put zakači na realne poslovne dokumente.

---

## 1. AMB-10b — opšta ocena

### Arhitektura

Dobro je odvojeno:

```text
modAmbalazaUgovor = šta je legalno
modAmbalaza        = kada i kako se upisuje
```

`modAmbalazaUgovor` nosi:

- zatvorene liste naloga;
- zatvorene liste vrsta kretanja;
- matricu dozvoljenih strana;
- `AmbNalogProblem` / `RequireAmbPrenos`;
- ugovor o ambalažnom dokumentu;
- doprinos obavezi.

Važno: modul ne piše u tabele. To je ispravno, jer ugovor mora biti jedan izvor istine za pisca, čitaoce i testove.

### Model

Novi model je append-only prenos:

```text
OdNalog -> NaNalog
Kolicina > 0
VrstaKretanja objašnjava poslovni razlog
Storno kasnije ide kao kontra-stav preko StornoOd
```

To je značajno čistije od starog `Smer + Entitet` modela, jer jedan red sada zna obe strane transfera.

Dobro pogođene invarijante:

```text
AMB-INV-07: nijedan realan nalog ne sme ispod nule
AMB-INV-09: obaveza partnera ne sme ispod nule
AMB-INV-10: jedan dokument = jedan poslovni par naloga
AMB-INV-11: AmbID jedinstven, StornoOd pokazuje na tačno jedan red
```

### Čitanje knjige

Dobro je što čitalac nije fail-open. Tokom prelaza tabela nosi dva oblika reda, pa svaki čitalac mora da razlikuje:

- stari validan red;
- novi validan red;
- prazan / artefakt red;
- pokvaren red koji ne pripada nijednom modelu.

Ovo je bitno jer bi red koji se tiho preskoči promenio saldo i sledeće odluke pisca.

---

## 2. Novi P1: `PrenesiAmbalazu` je public multi-write pisac bez sopstvene TX granice

`PrenesiAmbalazu` može da upiše do tri reda:

```text
1. traženi prenos
2. pokriće deficita
3. ostatak podele kod vraćanja tuđe ambalaže
```

Poslovni model je dobar. Problem je granica transakcije.

Funkcija je `Public Function PrenesiAmbalazu(...)`, direktno radi upise kroz `UpisiRedKnjige`, a nema svoj `PrenesiAmbalazu_TX`. Ako budući caller pozove ovu funkciju van `clsTransaction` wrappera, može nastati parcijalan ledger:

```text
glavni red upisan
↓
ostatak/pokriće pukne
↓
nema rollback-a ako caller nije u TX
```

Kod već ispravno kaže da AMB-INV-08, odnosno „upis u istoj transakciji sa izvornim dokumentom“, ne može da se dokaže iznutra i da mora biti statička kapija nad pozivnim mestima.

### Verdict

```text
Nije blocker za AMB-10b-1 jer nema cutover.
Biće blocker za prvi PR koji zakači PrenesiAmbalazu na realne dokumente.
```

### Required pre-cutover gate

Pre 10b-2 mora važiti:

```text
Nijedan produkcioni poziv `PrenesiAmbalazu` ne sme postojati van TX wrappera izvornog dokumenta.
```

Minimalni uslovi:

- `who_writes` / static gate mora znati da je `PrenesiAmbalazu` writer;
- svaki caller mora snapshotovati `TBL_AMBALAZA`;
- ako caller kreira `tblAmbalazaDokument` + `tblAmbalaza`, mora snapshotovati obe tabele;
- test mora namerno sabotirati drugi/treći upis i dokazati rollback celog izvornog dokumenta.

---

## 3. P1: `modStorno` rollback failure se i dalje guta

`HandleStornoTxError` i dalje koristi obrazac:

```vba
On Error Resume Next
LogErr procedureName
If Not tx Is Nothing Then tx.RollbackTx
...
On Error GoTo 0
```

Ako `RollbackTx` pukne, greška rollback-a se gubi. To je ozbiljno kod storna, jer korisnik dobija poruku da su promene vraćene, a sistem možda ostaje parcijalan.

### Verdict

```text
P1 otvoreno.
```

### Preporuka

U handleru treba posebno uhvatiti rollback grešku:

```text
originalErr = Err
attempt rollback
if rollback failed:
    Monitor_Critical / Monitor_Event severity=CRITICAL
    Debug.Print rollback error
    do not silently report ordinary failure only
```

---

## 4. P1: `modBankaMapiranje` rollback failure se i dalje guta

`AutoMapBankaImportRow_TX` i ostali `_TX` handleri i dalje imaju isti obrazac:

```vba
LogErr ...
On Error Resume Next
Monitor_BankaMapFail ...
If Not tx Is Nothing Then tx.RollbackTx
```

Ako rollback pukne, nema posebnog signala.

### Verdict

```text
P1 otvoreno.
```

Banka logika je zrelija nego ranije, naročito oko batch/manual-required grešaka, ali rollback failure visibility nije rešena.

---

## 5. P1: `modMain.InitApp` hardkoduje `Application` state

`InitApp` i dalje radi:

```vba
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
...
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
```

Ne pamti prethodno stanje. Za veliki Excel sistem ovo je stability issue, jer aplikacija može promeniti šire Excel stanje umesto da ga samo privremeno suspenduje.

### Verdict

```text
P1 otvoreno, mali patch.
```

### Minimalni patch

```vba
Dim oldScreen As Boolean
Dim oldCalc As XlCalculation
Dim oldEvents As Boolean

oldScreen = Application.ScreenUpdating
oldCalc = Application.Calculation
oldEvents = Application.EnableEvents

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

...

CleanUp:
    Application.ScreenUpdating = oldScreen
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
```

---

## 6. Self-update status

Ovaj AMB-10b delta ne menja prethodni self-update verdict.

Prethodna ocena ostaje:

```text
Self-update = P1 operational risk po prirodi,
ali prethodno dosta hardenovan.
```

Nema novog self-update blockera iz trenutnog `main`.

---

## 7. `modDataAccess` cache / audit side-effect

Prethodni zaključak ostaje:

```text
Cache rizik je smanjen.
Audit side-effect na svaki write ostaje stvar discipline TX snapshot-a.
```

AMB-10b dodatno povećava važnost tog pravila, jer novi writer intenzivno koristi centralne write primitive.

---

## 8. AMB-10b test pokrivenost

Dobro je što su AMB testovi ubačeni u `RunBusinessFlowProSuite`:

```text
Test_Amb_UgovorPrenosa
Test_Amb_DoprinosObavezi
Test_Amb_NalogDvosmislenPada
Test_Amb_DokumentUgovor
Test_Amb_PisacKnjige
Test_Amb_DeficitIObaveza
Test_Amb_ZaglavljeDokumentaPisac
Test_Amb_JedanProtivpartnerPoDokumentu
```

To je dobar minimum za 10b-1.

Ali još ne dokazuje kompletan produkcioni tok, jer cutover nije urađen. Testira se ugovor i pisac, ne kompletno knjiženje iz svih realnih dokumenata.

---

## Heatmap

| Oblast | Status | Ocena |
|---|---|---|
| `modSelfUpdate` | bez nove promene | P1 inherentni operational risk |
| `modMain.StartApp` | kompleksan, ali stabilan | P2 refactor kasnije |
| `modMain.InitApp` | hardkoduje Application state | P1 |
| `modStorno` rollback | rollback failure se guta | P1 |
| `modBankaMapiranje` rollback | rollback failure se guta | P1 |
| `modDataAccess` cache | ranije poboljšan | P2 |
| `modAuth` Windows hash | Windows-only ostaje | P2 |
| `AMB-10b ugovor` | jak domain model | GOOD |
| `AMB-10b pisac` | dobar, ali public multi-write | P1 pre cutovera |
| `AMB-10b cutover` | još nije urađen | sledeći rizični trenutak |

---

## Konačni verdict

```text
P0: 0

P1:
- rollback failure visibility u `modStorno`
- rollback failure visibility u `modBankaMapiranje`
- `InitApp` hardkoduje Application state
- `PrenesiAmbalazu` mora dobiti TX/caller gate pre prvog produkcionog poziva

P2:
- Windows-only platforma
- public non-TX storno core funkcije
- buduće performanse AMB knjige
```

Ukupna ocena trenutnog `main`:

```text
Code quality:            4.2 / 5
Domain architecture:     4.4 / 5
Regression discipline:   4.1 / 5
Operational safety:      3.8 / 5
Platform portability:    1.8 / 5
```

Zaključak:

```text
Ne bih vraćao AMB-10b-1 nazad.
Ne vidim novi P0.
Sledeći blocker mora biti postavljen na 10b-2 cutover:
`PrenesiAmbalazu` ne sme ući u produkcioni tok van transakcije izvornog dokumenta.
```
