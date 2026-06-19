# PREDLOG — Modul „Admin + korisnici sa pravima po oblastima" (VBA/Excel) — v1 DRAFT

> Status: **predlog / nije implementirano**. Po Codebase Guardian doktrini
> (`reuse > new`, `extend > duplicate`, `minimal change`, `inspect before propose`).
> **Cilj:** dodati u VBA/Excel aplikaciju login + admina sa svim pravima + korisnike
> kojima admin odobrava pristup **po oblastima** (Otkup, Dokumenta, Agrohemija, …),
> tako što se **kači na postojeće tačke** (startup gate, jedan launcher chokepoint,
> Maticni podaci CRUD), bez novog data sloja i bez diranja `.frx`-a.
>
> PWA je van opsega — tamo je pristup već rešen kroz `UserRole` (GAS + `role-nav.js`).

---

## 0) TL;DR + preporuke

- VBA app **nema pojam korisnika/role** — postoji samo licenca/trial **po uređaju**
  (`modLicense`/`modTrial`) i PIN **po stanici** (`tblStanice`, `modStanicaLock`). Dakle
  ovo gradimo iznad postojećeg, ne diramo licencu ni stanica-lock.
- **Dve idealne tačke kačenja (potvrđene u kodu):**
  1. **Login gate** ide u `modMain.StartApp`, **između** `AccessGateOrQuit()`
     (`modMain.bas:32`) i `frmSplash.Show` (`modMain.bas:36`).
  2. **Jedna provera prava za SVE oblasti** ide na početak `OpenContentForm(...)`
     (`frmOtkupAPP.frm:904`) — kroz njega prolaze sva dugmad sekcija
     (`frmOtkupAPP.frm:562–679`).
- **Admin UI bez novog CRUD ekrana i bez `.frx`:** „Korisnici" se dodaje kao nova
  sekcija u **postojeći** Maticni podaci meni — `jedan red` u
  `modMaticniLookups.MaticniSekcije()` + `Case "Korisnici"` u `frmStammdaten`
  (isti mehanizam kojim su dodate Kulture/Cenovnik).
- **Bez novog data sloja:** sve čitanje/pisanje preko postojećih
  `FindRows/GetTableData/GetColumnIndex/AppendRow/UpdateCell/GetNextID/LookupValue`
  (`modDataAccess`). Šema preko `EnsureKorisniciSchema` (mirror `EnsureCenovnikSchema`).
- **Opt-in rollout (preporuka):** `AUTH_ENABLED` flag (default `NE`), mirror
  `LICENSE_ENABLED`/`TRIAL_ENABLED` → dok admin ne uključi i ne kreira prvog admina,
  ništa se ne menja → **niko se ne može slučajno zaključati**.
- **Otvorene odluke** u §9 (model prava, login UI, opt-in, PIN hash) — tražim potvrdu pre koda.

---

## 1) Šta VEĆ postoji (inventar za reuse) — potvrđeno u kodu

### 1.1 Startup sekvenca (gde ide login)
- `ThisWorkbook.Workbook_Open` → `StartApp` (`ThisWorkbook.doccls:6`), pa
  `If AccessWasDenied() Then Exit Sub` (`:10`), pa `CleanupOrphanedLocks` (`:16`).
- `modMain.StartApp`: `InitApp` (`modMain.bas:26`) → **`AccessGateOrQuit()`**
  (`modMain.bas:32`, licenca+trial) → `Application.Visible = False` (`:34`) →
  **`frmSplash.Show`** (`:36`) → splash → `frmOtkupAPP.Show` (`frmSplash.frm`).
- **Hook za login:** odmah posle `:32` (gate prošao), pre `:36` (splash).

### 1.2 Launcher (jedan chokepoint za sve oblasti)
- `frmOtkupAPP` je glavni shell; dugmad sekcija zovu `OpenContentForm frmX, btn, "Naslov"`
  (`frmOtkupAPP.frm:562–679`): `btnBlocks→frmOtkup`, `btnPurchase→frmDokumenta`,
  `btnAgro→frmAgrohemija`, `btnReports→frmIzvestaj`, `btnInvoicing→frmFakturisanje`,
  `btnBanka→frmBankaImport`, `btnMargin→frmMarza`, `btnTrace→frmSledljivost`.
- **`OpenContentForm(contentForm, activeBtn, sectionTitle)`** @ `frmOtkupAPP.frm:904`,
  prikaz preko `mActiveContent.Show vbModeless` @ `:932`. → **jedna guard tačka na `:907`.**
- `btnMaticni → OpenMaticniForm()` (`frmOtkupAPP.frm:~745`) je zaseban put → `frmMaticniPodaci`.

### 1.3 Šema/tabele (idempotentni obrazac)
- `EnsureDataTable(tblName, sheetName, headers())` @ `modSetup.bas:767` — kreira ListObject;
  ako postoji, dopuni kolone preko `EnsureColumnOnTable` (schema-drift safe).
- Javni mirror primer: `EnsureCenovnikSchema` @ `modSetup.bas:732`.
- Dijagnostika `DebugKoloneTabele` @ `modSetup.bas:750` (Alt+F8).

### 1.4 Pristup podacima (reuse, bez novog sloja)
- `modDataAccess`: `GetTable/GetTableData/GetTableHeaders/GetColumnIndex/GetColumnData/`
  `AppendRow/UpdateCell/FindRows/LookupValue/GetLookupList/GetNextID`.

### 1.5 Konstante (modConfig obrazac)
- `TBL_*` blok `modConfig.bas:13–59`; `COL_*` blokovi `:70–401`; `SHT_*` `:61–68`.

### 1.6 „Ko koristi app" — danas NE postoji
- Nema per-user login-a. `gActiveStanica` (Private) + `GetActiveStanica()`
  (`modStanicaLock.bas:269`) prate aktivnu **stanicu**, ne korisnika. PIN u `tblStanice`
  je **stanica-level**. Identitet korisnika je danas konstantno `"Operator"` u monitoringu.

### 1.7 Admin/Setup danas
- Sve preko Alt+F8 (`SetupNewPC`, `EnsureCenovnikSchema`, `ActivateLicensePrompt`,
  `DebugKoloneTabele`, `EnableDesktopOnlyMode`…). Nema menija ni role.
- `Monitor_Event` (audit) već se koristi u startup-u → reuse za login/deny događaje.

---

## 2) Oblasti (jedinice prava) — iz launcher-a

| Oblast (vrednost prava) | Forma | Dugme | Modul |
|---|---|---|---|
| `Otkup` | frmOtkup | btnBlocks | modOtkup |
| `Dokumenta` | frmDokumenta | btnPurchase | modDokumenta |
| `Agrohemija` | frmAgrohemija | btnAgro | modAgrohemija |
| `Izvestaji` | frmIzvestaj | btnReports | modIzvestaj |
| `Fakturisanje` | frmFakturisanje | btnInvoicing | modFaktura |
| `Banka` | frmBankaImport / frmBankaExportPregled | btnBanka | modBankaImport |
| `Marza` | frmMarza | btnMargin | modMarza |
| `Sledljivost` | frmSledljivost | btnTrace | — |
| `MaticniPodaci` | frmMaticniPodaci | btnMaticni | modMaticniLookups |

> `Novac`/`Ambalaza` su sheet-based (bez dugmeta) → u v1 nisu zasebne oblasti; mogu
> kasnije kao kolone. „Korisnici" (administracija) = podsekcija Maticnih, vidljiva samo Adminu.

---

## 3) Predlog — minimalni delta

### 3.1 Model podataka — `tblKorisnici` (jedan red = jedan korisnik)

Kolone (kreira `EnsureKorisniciSchema`, idempotentno):

| Kolona | Vrednosti | Svrha |
|---|---|---|
| `KorisnikID` | KOR-00001 (`GetNextID`) | PK |
| `Username` | tekst | login |
| `ImePrezime` | tekst | prikaz |
| `PIN` | tekst | login (parity sa `tblStanice.PIN`) |
| `Uloga` | `Admin` / `Korisnik` | Admin = sva prava (bypass) |
| `Aktivan` | `DA` / `NE` | blok logina bez brisanja reda |
| `StanicaID` | (opciono) | veza ka `tblStanice` |
| `Otkup, Dokumenta, Agrohemija, Izvestaji, Fakturisanje, Banka, Marza, Sledljivost, MaticniPodaci` | `DA`/`NE` | **prava po oblasti** |
| `CreatedAt` | datum | trag |

**Zašto kolone-po-oblasti (a ne matrica):** Excel-native, admin vidi/edituje prava
očima u gridu, nova oblast = `EnsureColumnOnTable` (jedan red u array-u). Admin red ima
sve `DA` i `Uloga=Admin` → uvek prolazi. *(Alternativa = matrica `tblKorisniciPrava`,
§9 odluka 1, ako zatreba read/write granularnost.)*

> Ovo **ne duplira** `tblStanice.PIN` — to je identitet/lock stanice; `tblKorisnici` je
> app-login + prava po oblasti. Različita svrha → opravdano nova tabela.

### 3.2 Novi modul `modAuth.bas` (stanje + guard) — mirror `modStanicaLock` globala
```vba
Private gCurrentUser As String
Private gCurrentUserUloga As String

Public Function AuthEnabled() As Boolean          ' cita AUTH_ENABLED (opt-in)
Public Function Login() As Boolean                ' frmLogin/InputBox -> validacija nad tblKorisnici -> set global + Monitor_Event
Public Function GetCurrentUser() As String
Public Function CurrentUserIsAdmin() As Boolean   ' gCurrentUserUloga = "Admin"
Public Function KorisnikImaPravo(ByVal oblast As String) As Boolean
    ' Admin -> True; inace LookupValue(TBL_KORISNICI, Username->oblast) = "DA"
Public Sub Logout()
```

### 3.3 Login gate — `modMain.StartApp` (između `:32` i `:36`)
```vba
If Not AccessGateOrQuit() Then Exit Sub           ' postojece :32
If modAuth.AuthEnabled() Then
    If Not modAuth.Login() Then Exit Sub          ' fail/cancel -> quit (mirror license gate)
End If
Application.Visible = False                        ' postojece :34
frmSplash.Show                                     ' postojece :36
```

### 3.4 Per-oblast guard — JEDNA tačka u `OpenContentForm` (`frmOtkupAPP.frm:907`)
```vba
Private Sub OpenContentForm(ByVal contentForm As Object, _
                            ByVal activeBtn As MSForms.CommandButton, _
                            ByVal sectionTitle As String)
    On Error GoTo EH
    If modAuth.AuthEnabled() Then
        If Not modAuth.KorisnikImaPravo(OblastZaFormu(contentForm.name)) Then
            MsgBox "Nemate dozvolu za oblast: " & sectionTitle, vbExclamation, APP_NAME
            Exit Sub
        End If
    End If
    ' ... postojeci kod ...
```
`OblastZaFormu("frmOtkup")="Otkup"` itd. Isti uslov i u `OpenMaticniForm` za `MaticniPodaci`.

### 3.5 Admin UI — bez novog ekrana, bez `.frx`
- `modMaticniLookups.MaticniSekcije()`: dodati **jedan red** `Array("Korisnici","Korisnici")`,
  prikazan samo ako `CurrentUserIsAdmin()` (filter u `AttachMaticniMenu`).
- `frmStammdaten`: dodati `Case "Korisnici"` (veže `tblKorisnici`, headeri, CRUD) — reuse
  cele Stammdaten grid/save mašinerije. Admin tu dodaje/menja korisnike i čeklira oblasti.

### 3.6 Login UI
- **Preporuka:** mali **novi** `frmLogin` (TextBox `PasswordChar="●"` za maskiran PIN).
  To je NOVA forma (svoj `.frx`), dodaje se u VBA IDE — **ne** edituje se postojeći `.frx`.
- **Fallback bez `.frx`:** `InputBox` za Username + PIN u `modAuth.Login()` (PIN vidljiv).

### 3.7 Šema + bootstrap (modSetup, mirror postojećih `Ensure*`/`Enable*`)
```vba
Public Sub EnsureKorisniciSchema()                 ' mirror EnsureCenovnikSchema
    EnsureDataTable TBL_KORISNICI, "Korisnici", _
        Array(COL_KOR_ID, COL_KOR_USERNAME, COL_KOR_IME, COL_KOR_PIN, _
              COL_KOR_ULOGA, COL_KOR_AKTIVAN, COL_KOR_STANICA, COL_KOR_CREATED, _
              "Otkup","Dokumenta","Agrohemija","Izvestaji","Fakturisanje", _
              "Banka","Marza","Sledljivost","MaticniPodaci")
End Sub
Public Sub KreirajPrvogAdmina()  ' Alt+F8: EnsureKorisniciSchema + AppendRow Admin (sve DA)
Public Sub EnableAuth()          ' Alt+F8: AUTH_ENABLED=DA (mirror EnableDesktopOnlyMode)
Public Sub DisableAuth()         ' Alt+F8: AUTH_ENABLED=NE
```
`AUTH_ENABLED` u `tblLocalConfig` (`GetLocalConfigValue/SetLocalConfigValue`) ili
`tblSEFConfig` (`GetConfigValue`) — kao postojeći flagovi.

### 3.8 Konstante (modConfig — aditivno)
```vba
' uz TBL_* blok (~:59)
Public Const TBL_KORISNICI As String = "tblKorisnici"
' uz COL_* blokove (~:401) — KorisnikID/Username/ImePrezime/PIN/Uloga/Aktivan/StanicaID/CreatedAt
```

---

## 4) Mapa delte (fajlovi + tačne tačke)

| Fajl | Tip | Šta | Hook |
|---|---|---|---|
| `modConfig.bas` | aditivno | `TBL_KORISNICI`, `COL_KOR_*` | `:59`, `:401` |
| `modAuth.bas` | **nov modul** | login state + `Login/KorisnikImaPravo/...` | mirror `modStanicaLock` globala |
| `modSetup.bas` | aditivno | `EnsureKorisniciSchema`, `KreirajPrvogAdmina`, `EnableAuth/DisableAuth` | mirror `EnsureCenovnikSchema`/`EnableDesktopOnlyMode` |
| `modMain.bas` | mala izmena (code) | login poziv (opt-in) | između `:32` i `:36` |
| `frmOtkupAPP.frm` | mala izmena (code, bez `.frx`) | guard + `OblastZaFormu` + Maticni guard | `OpenContentForm` `:907`, `OpenMaticniForm` |
| `modMaticniLookups.bas` | +1 red | „Korisnici" sekcija (Admin-gated) | `MaticniSekcije()` |
| `frmStammdaten.frm` | code, bez `.frx` | `Case "Korisnici"` (CRUD reuse) | `Select Case Me.Tag` |
| `frmLogin` | **nova forma** (opciono) | maskiran PIN | ili InputBox fallback |
| `instructions/…md` | doc | ovaj predlog | — |

**Bez novih:** data-access funkcija, parsing/schema mašinerije, role-tabela, diranja `.frx`.

---

## 5) Tok (startup → login → oblast)
1. `Workbook_Open → StartApp → AccessGateOrQuit()` (licenca/trial, nepromenjeno).
2. Ako `AUTH_ENABLED=DA`: `modAuth.Login()` (Username+PIN nad `tblKorisnici`, `Aktivan=DA`);
   3 pokušaja pa quit (mirror license). Postavlja `gCurrentUser`/`Uloga` + `Monitor_Event AUTH_LOGIN`.
3. `frmSplash → frmOtkupAPP`. Klik na sekciju → `OpenContentForm` guard: Admin uvek prolazi;
   korisnik prolazi ako je oblast `DA`; inače MsgBox + `Monitor_Event AUTH_DENIED`.
4. „Korisnici" u Maticnim podacima vidljiv samo Adminu → CRUD + čekiranje oblasti.

---

## 6) Bezbednost / verifikacija
- **PIN:** plaintext u v1 = parity sa postojećim `tblStanice.PIN` (desktop, lokalni workbook).
  Opcioni hash kasnije (§9 odluka 4).
- **Lockout-safety:** `AUTH_ENABLED` default `NE` + `KreirajPrvogAdmina` → ne može se
  niko slučajno zaključati; postupno uvođenje.
- **Audit:** reuse `Monitor_Event` (AUTH_LOGIN / AUTH_DENIED / AUTH_USER_CHANGED).
- **VBA verifikacija (CLAUDE.md §4–5):** posle dodavanja `modAuth` uraditi
  `Debug → Compile VBAProject` (nema duplih `Public` imena → „Ambiguous name"); statički
  balans `Sub/Function/Select Case`; finalni smoke-test u Excelu radi korisnik.

---

## 7) Faze isporuke
- **Faza 1 (MVP):** `tblKorisnici` + `modAuth` + login gate (opt-in) + guard u
  `OpenContentForm` + `KreirajPrvogAdmina`/`EnableAuth`. Login InputBox ili `frmLogin`.
- **Faza 2:** „Korisnici" sekcija u Maticnim (CRUD + čekiranje oblasti) + guard za
  `OpenMaticniForm`.
- **Faza 3 (opciono):** PIN hash + migracija; per-oblast read/write (matrica) ako zatreba;
  Alt+F8 setup/admin makroi ograničeni na Admin.

---

## 8) Rizici / napomene
- `frmOtkupAPP`/`frmStammdaten` izmene su **samo u kodu** (`.frm`), `.frx` se ne dira
  (CLAUDE.md §4). `frmLogin` je nova forma sa sopstvenim `.frx` (ako se izabere taj UI).
- `OblastZaFormu` mapiranje mora pokriti sve forme iz §2 (default: ako oblast nije
  mapirana → tretiraj kao dozvoljeno ili kao MaticniPodaci? → §9 nije nužno, default
  „dozvoljeno" da se ne blokira nepoznata/buduća sekcija; Admin svejedno prolazi).

---

## 9) Otvorene odluke (molim potvrdu pre koda)
1. **Model prava:** kolone-po-oblasti `DA/NE` na `tblKorisnici` *(preporuka — Excel-native,
   minimalno)* **vs** matrica `tblKorisniciPrava` (za buduće read/write po oblasti).
2. **Login UI:** novi `frmLogin` sa maskiranim PIN-om *(preporuka)* **vs** `InputBox`
   (bez `.frx`, ali PIN vidljiv).
3. **Rollout:** `AUTH_ENABLED` opt-in flag, default `NE` *(preporuka — bez rizika lockout-a)*
   **vs** odmah obavezan login.
4. **PIN:** plaintext kao sad *(preporuka v1)* **vs** hash odmah.
5. **Alt+F8 setup/admin makroi** (SetupNewPC, Ensure*, License): ostaviti kao IT/power-user
   van auth-a *(preporuka)* **vs** i njih zaključati na Admina.

---

_Po potvrdi §9 krećem od Faze 1 (login + guard u jednoj chokepoint tački), strogo
proširujući postojeće obrasce (modSetup `Ensure*`, modDataAccess, Maticni meni) bez novog
data sloja i bez diranja `.frx`._
