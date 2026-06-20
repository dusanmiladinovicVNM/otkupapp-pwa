# Dorade — checklista za proveru (6 funkcionalnosti)

> Grana: `claude/loving-fermi-nzvof0`. VBA se ne kompajlira u CI — verifikovano
> statički (balans, nema duplih `Public` definicija). Finalni test u Excelu.

## 0) Priprema (jednom)

- [ ] **Re-import VBA** (forme idu sa `.frx` parom; modul/klasu Remove pa Import):
  - Novi: `modAutoHladnjaca.bas`, `clsStmBtn.cls`
  - Izmenjeni: `frmStammdaten.frm`, `frmOtkup.frm`, `frmDokumenta.frm`,
    `modConfig.bas`, `modPodesavanja.bas`, `modSetup.bas`, `modAmbalaza.bas`,
    `modHelpers.bas`, `modComboBinding.bas`, `modDataAccess.bas`
- [ ] **`Alt+F8 → EnsureDoradeSchema`** (kreira kolone + decimalni format)
- [ ] **Debug → Compile VBAProject** (mora bez greške)
- [ ] **Maticni podaci → Podešavanja** — proveri da se vide novi ključevi
  (grupa „Otkup / dokumenta"): `DEFAULT_VRSTA_VOCA`, `DEFAULT_SORTA_VOCA`,
  `KOOP_FILTER_BY_OM`, `AUTO_PRIJEMNICA_HLADNJACA`

## 1) Soft-delete (Aktivan ⇄ Neaktivan) u svim šifarnicima
- [ ] U `EnsureDoradeSchema` dodate kolone `Aktivan` u Kulture/Ambalaza/Palete,
      postojeći redovi popunjeni „Aktivan"
- [ ] Maticni podaci → bilo koji šifarnik sa statusom (Kooperanti/Stanice/Kupci/
      Vozaci/Kulture/Ambalaza/Palete): pojavljuje se dugme **„Deaktiviraj/Aktiviraj"**
- [ ] Izaberi red → klik → status se promeni (red ostaje u listi, ne briše se)
- [ ] Deaktiviran zapis **NESTAJE** iz padajućih lista u **frmOtkup/frmDokumenta**
      (vrsta/sorta, tip ambalaže, otkupno mesto, kupac, kooperant), a ostaje u
      Maticnim podacima (može da se reaktivira)

## 2) Default vrsta/sorta (Podešavanja)
- [ ] U Podešavanjima upiši `DEFAULT_VRSTA_VOCA` (npr. „Malina") i
      `DEFAULT_SORTA_VOCA` (npr. „Willamette")
- [ ] Otvori **frmOtkup** → vrsta i sorta su već popunjene; auto se popuni i
      cena (cenovnik) i tip ambalaže (kultura)
- [ ] Otvori **frmDokumenta** → isto (vrsta/sorta popunjene, samo ako prazno)

## 3) Hladnjača → auto otpremnica+zbirna+prijemnica
- [ ] Maticni podaci → **Stanice**: postavi `Hladnjača? = Da` za jednu stanicu
- [ ] Podešavanja: `MALINA_DEFAULT_KUPAC = <KupacID hladnjače>`,
      `AUTO_PRIJEMNICA_HLADNJACA = Da`
- [ ] frmOtkup: izaberi tu (hladnjača) stanicu, unesi otkup (Klasa I, po želji II)
      → klik **Unos**
- [ ] Provera tabela: automatski kreirani redovi u **tblOtpremnica**,
      **tblZbirna**, **tblPrijemnica** (cena = cena iz otkupa)
- [ ] **Izvorni `tblOtkup` red** dobija nazad `OtpremnicaID`, `BrojZbirne` i
      `VozacID` (ako je otkup bio bez vozača → upiše se `VozacID = StanicaID`;
      operaterov vozač se ne gazi)
- [ ] **BrojPrijemnice**: prva tog dana = `1/DDMMYY` (npr. `1/200626`), sledeća
      `1/200626-2`, pa `-3` …
- [ ] Ako stanica NIJE hladnjača ILI je toggle OFF → ništa se ne kreira automatski

## 4) Toggle filtera kooperanata po otkupnom mestu
- [ ] Podešavanja: `KOOP_FILTER_BY_OM = Da` (default) → frmOtkup prikazuje samo
      kooperante izabranog OM
- [ ] `KOOP_FILTER_BY_OM = Ne` → frmOtkup prikazuje **sve** kooperante

## 5) Decimalna količina
- [ ] frmOtkup/frmDokumenta: unesi količinu sa decimalom („12,5" ili „12.5")
- [ ] Sačuvaj → vrednost u `tblOtkup.Kolicina` je 12,5 (format `0.00` posle
      `EnsureDoradeSchema`), ne zaokružena

## 6) Podrazumevani tip ambalaže po kulturi
- [ ] Maticni podaci → **Kulture**: za kulturu izaberi „Tip ambalaže (podraz.)"
- [ ] frmOtkup: izbor te vrste/sorte → polje **Tip ambalaže** se auto-popuni
- [ ] frmDokumenta: isto (otpremnica + prijemnica tip ambalaže)

## Napomene / zavisnosti
- **Malina mod**: u `frmOtkup`/`frmDokumenta`, izbor OM odmah auto-bira vozača
  (par-vozač, `VozacID == StanicaID`). Tako otkup red dobija `VozacID` već pri
  snimanju → radi auto-povezivanje u sledljivosti. (Ako par-vozač nije u listi —
  ostaje prazno; mirror se pravi `BackfillVozacMirrorsForMalina` / pri unosu stanice.)
- #3 zavisi od postavljenog `MALINA_DEFAULT_KUPAC` (kupac-hladnjača) — ako je
  prazno, auto-lanac se preskače (zabeleži se u log).
- #3 koristi broj otkupnog dokumenta kao broj otpremnice/zbirne (ako otkup nema
  broj → generiše `HL-DDMMYY-hhnnss`).
- **Predlog `BrojZbirne`** (`SuggestNextBroj` ZBR): nikad ne nudi već zauzet broj.
  Test regresije: u malina modu napravi zbirnu (npr. `1/DDMMYY`), ugasi malina mod,
  unesi otpremnicu sa vozačem istog numeričkog dela (npr. `VOZ-00001`) → predlog
  mora biti `1/DDMMYY-2`, ne ponovo `1/DDMMYY`.
- #1: dugme se vidi samo gde tabela ima kolonu `Aktivan`/`Aktivna` (Artikli i
  Cenovnik nemaju → nema dugmeta).
