# Release notes — AgriX / OtkupApp (VBA)

Za svaku release verziju (`vba-vX.Y.Z`) — par rečenica šta je urađeno.
Dopunjava se pri svakom `tools/release.sh` (korak B-11 u `RELEASE_PROCEDURE.md`).

> Razlika od `ARCHITECTURE_CHANGELOG.md`: tamo je detaljna arhitektonska istorija
> (interno `v6.xx`); ovde su **release-tagovi** (`vba-vX.Y.Z`) sa kratkim opisom za
> korisnika — „šta je novo u ovom .xlsm-u".

**Format:**
```
## vba-vX.Y.Z — YYYY-MM-DD
- promena 1
- promena 2
```

---

## vba-v2.2.3 — 2026-06-25
- Pregled storniranih: dugme u Dokumentima -> panel sa svim storniranim (po tipu) + lanac zavisnih (Zbirna/Otpremnica/Faktura).
- Izgubljeni otkup blokovi: upozorenje na dashboardu + sekcija „Izgubljeni" u Otkupnim blokovima sa „Preuzmi" (vrati blok na ispravnu otpremnicu posle storna; cuva OtkupID/uplate/ambalazu).
- Otkupni blokovi: kolona „Kupac" (firma) umesto prazne „Hladnjaca".
- Storno dvoklasne otpremnice/prijemnice sada stornira OBE klase (ranije samo jednu).
- Validacija zbirne vise ne blokira unos kada ima ambalaze Klase II.
- Prosek gajbi uracunava obe klase.
  
---

## vba-v2.3.0 — 2026-06-25
- **Grupni otkupni list (malina mod):** posle snimanja prijemnice (za default hladnjaču) izlazi na obrascu otkupnog lista (dva primerka, 1/3 A4) — umesto proizvođača prikazuje **otkupno mesto** (stanica = vozač u malina modu); proizvod/klasa/količina/ambalaža/cena se čitaju iz prijemnice (cena BRUTO → neto + PDV nadoknada); ambalaža = saldo na nivou stanice (entitet „Stanica"). Prijemnica i pojedinačni otkupni list ostaju nepromenjeni.
- **Revers izdavanja ambalaze:** PDF revers pri izdavanju prazne ambalaže kooperantu, na 1/3 A4 sa dva primerka (kao otkupni list).
- **Verzionisanje koda:** novi `modBuildInfo` (`BUILD_SHA` / `BUILD_VERSION` / `BUILD_DATE`), stamp pri buildu (`tools/stamp-build`).
- **Auto verzija:** `BUILD_VERSION` iz `git describe` (na tagu čisto, između tagova se sam diže).
- **Telemetrija builda:** `modMonitoring` i `modLicense` šalju `buildSha`/`buildVersion`/`buildDate`.
- **Fleet pregled „ko ima koju verziju":** GAS `Events`/`Fleet` + `rebuildMonitoringFleet` (auto na sat preko `installMonitoringTriggers`).
- **Release rutina:** `tools/release.sh|ps1` (jedna komanda) + procedura `docs/RELEASE_PROCEDURE.md`.
- **Min-version gate:** `modUpdateGate` na startu pita GAS (`checkVersion`) za minimalnu dozvoljenu verziju; zastarela verzija dobija upozorenje (ili blok uz `VERSION_ENFORCE=YES`). Opt-in + fail-open. Podešava se preko GAS Script Properties (`VERSION_MIN`/`VERSION_LATEST`/`VERSION_ENFORCE`/`VERSION_MESSAGE`), bez redeploy-a.
- **Blanko build guard:** `modBuildGuard.AssertBlankBuild` (ručno, build-only) proveri da fajl pre distribucije nema podataka — sprečava da blanko `AgriX_x.x.x` ode klijentima sa tuđim podacima.
- **Distribucija:** konvencija `builds\AgriX_x.x.x.xlsm` (blanko, isti za sve) dokumentovana u proceduri.

---

## vba-v2.4.0 — 2026-06-26
- Storno kaskade: malina mod — storno otpremnice automatski stornira i njenu 1:1 zbirnu; autohladnjača — storno otkupnog bloka kaskadno stornira ceo auto-generisani lanac (otpremnica + zbirna + prijemnica). Faktura se NE dira.
- Recovery panel „Osiroćeni dokumenti" (Dokumenti): re-point osiroćene prijemnice (zbirna stornirana) na novu aktivnu zbirnu — biranje iz liste, bez kucanja; menja se samo BrojZbirne (PrijemnicaID ostaje → faktura/palete ispravne).
- Palete re-point (isti panel, „Palete" mod): posle storno+ponovni unos u autohladnjači, prevezivanje paleta sa stornirane na novu prijemnicu — poništava dvostruku auto-paletizaciju, čuva fizičke (zatvorene) palete, KG-sync kad se broj gajbica poklapa; razlika u broju gajbica → upozorenje (aneks ručno).

---

## sledeći release (u pripremi) — planirano vba-v2.5.0
Tačan broj/datum se postavlja pri `tools/release.sh` (planirano: **2.5.0**).

- **Izveštaji — „Detalji" panel + štampa na Otkupljena roba i Ambalaza:** klik na red prikazuje read-only pregled desno (kao Kartica/Otkupni listovi), sa dugmetom za štampu. Otkupljena roba → **„Štampaj otpremnicu"** (PDF u stilu otkupnog lista, podaci iz `tblOtpremnice`). Ambalaza → **„Štampaj dokument"** rutiran po tipu (Prijemnica/Otkup/Otpremnica/Revers).
- **Otkupljena roba (po otpremnici):** nova kolona **„Prijemnica kg"** (malina = direktno iz prijemnice; inače srazmerno udelu otpremnice u zbirnoj); „Manjak kg" i „Manjak %" spojeni u jednu kolonu (ListBox limit 10 kolona).
- **Ambalaza tab (preimenovan iz „Primljena ambalaza"):** ako isti dokument ima i Ulaz i Izlaz, prikazuju se u istom redu; kolona „Dokument" prikazuje **poslovni broj** (ne interni ID).
- **Revers OM↔kooperant:** štampa reversa za izdavanje/povrat prazne ambalaže (OM-Izlaz-Koop / OM-Ulaz-Koop), rekonstruisan iz ledgera.
- **Detalji otpremnice:** Cena, Vrednost i **Broj prijemnice** (poslovni broj iz zbirne).
- **Saldo OM:** header-i poravnati tačno nad kolonama; kolona **„Ambalaža" = aktivni neto saldo** kooperanta iz ledgera (Ulaz − Izlaz), umesto zbira predatih gajbica.
- **Štampa izveštaja (`PrintIzvestaj`):** izlaz kao PDF koji se otvori (pouzdanije, bez zavisnosti od podrazumevanog štampača).
- **Otkup — auto-kreiranje kooperanta (toggle):** novo podešavanje „Auto-kreiraj kooperanta iz unetog imena" (default DA). Kad je NE, unos imena koje ne postoji u bazi više ne pravi tiho novog kooperanta — `frmOtkup` javi da kooperant nije pronađen.
- **Otkup — auto-cena se ne gubi:** posle snimanja prvog otkupnog lista cena/tip ambalaže se automatski ponovo popune za i dalje izabranu vrstu/sortu (ranije je polje ostajalo prazno do ponovnog otvaranja forme).
- **Otkup — info o aktivnoj paleti:** ispod dugmeta „Povratak" prikazuje se koliko gajbi još treba da se zatvori aktivna (otvorena) paleta za trenutno izabranu vrstu/sortu (Klasa I, i Klasa II kad je uključeno „Dve klase"); read-only, ne kreira paletu.
- **Otkupni blokovi (panel) — zbirna + prijemnica:** klik na otpremnicu sada prikazuje broj **zbirne** i broj **prijemnice** za koju je ta otpremnica vezana (veza preko `BrojZbirne`).
- **Podešavanja — toggle „Praćenje parcela":** kad je NE, polje za parcelu u `frmOtkup` je vidljivo ali onemogućeno (bez unosa; tab ga preskače).
- **Podešavanja — toggle „Postoje keš isplate proizvođačima":** kad je NE, u `frmOtkup` su Novac i Primalac onemogućeni, a u Dokumentima → „Ulaz OM" je onemogućeno polje „Br. otk. blk." (sve ostaje vidljivo/sivo; tab preskače).
- **Paletni list gotovih proizvoda — šifarnici:** novi šifarnici u Matičnim podacima — **Kutije** (tip + težina), **Kese** (tip + težina) i **Vrsta gotovog proizvoda**. Prozor „Matični podaci" se sam proširuje da stanu sve sekcije.
- **Paletni list gotovih proizvoda — prerada (`frmPalete`):** uz preradu se sada unose **težina palete**, **bruto**, **tip + broj kutija**, **tip + broj kesa** i bira **gotov proizvod**; **neto se računa automatski** (bruto − težina palete − težina ambalaže = broj·težina po tipu). Filteri Godina/Vrsta/Sorta/Status/Prerađeno u jednom redu; nova desna lista **„Prerađene palete"** (istorija) sa **dvoklik = (re)štampa PDF**.
- **Paletni list (PDF) — preimenovan iz „Preradni list":** naslov **„Paletni list gotovih proizvoda"**; vrsta = **„DZ" + vrsta + sorta + tip gotovog proizvoda**; desni sažetak ima 6 redova (težina palete / bruto / broj kutija / broj kesa / težina ambalaže / neto).
