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

## vba-v2.5.0 — 2026-06-27

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

---

## vba-v2.6.0 — 2026-06-28

- **Lokalizacija — ASCII-only VBA izvori (kraj `š/ž/č` korupcije):** sva dijakritika izmeštena iz koda u runtime katalog (`tblPoruke` → `modPoruke.Poruka("KLJUC")`, tekst se gradi `ChrW`-om). Izvori (`.bas/.cls/.frm/.doccls`) su sada **100% ASCII** → bezbedni za bilo koji editor (nema više tihog kvarenja `š/ž/č/ć` pri snimanju). 284 string-literala migrirano u `Poruka()`; statički natpisi formi se auto-koriguju pri otvaranju (`FixFormCaptions`); invarijanta dokumentovana u `CLAUDE.md` (sekcija 4).
- **Self-update (auto-ažuriranje koda preko Drive-a):** klijent na `Workbook_Open` proveri `AgriX_Release/version.json`, i ako postoji novija verzija ponudi ažuriranje — uz potvrdu povuče nov kod sa Drive-a i uveze ga u sebe, **bez migracije podataka** (isti `.xlsm`; šema se self-heal-uje kroz `InitApp` posle restarta). Lokalna backup kopija se napravi pre ažuriranja. Opt-in (`REL_FOLDER_ID` u `modConfig`); zahteva uključen „Trust access to the VBA project object model" na klijentu. Build objava: `Alt+F8 → PublishReleaseToDrive` (`modRelease`). Detalji i naučene zamke: `docs/SELF_UPDATE.md`.
- **Release procedura:** novi korak 7b (`PublishReleaseToDrive`) u `docs/RELEASE_PROCEDURE.md` — objava `src-vba` + `version.json` u `AgriX_Release` posle `AssertBlankBuild`.
- **Self-update poruke lokalizovane:** dijalozi ažuriranja (prompt, rezultat, greške, Trust-access) idu kroz `Poruka()` katalog (`SU_*` ključevi) — pun `š/ž/č`.

---

## vba-v2.7.0 — u pripremi
Tačan broj/datum se postavlja pri `tools/release.sh` (planirano: **2.7.0**).

- **PDF dokumenti — namenski podfolderi:** generisani PDF-ovi se više ne mešaju u root folderu pored radne sveske, već svaki tip ide u svoj podfolder — `Otkupni listovi`, `Prijemnice`, `Otpremnice`, `Revers ambalaze`, `Kartice kooperanata`, `Paletni listovi`, `Preradni listovi`, `Specifikacije`, `Izvestaji`. Folderi se prave automatski (pri prvom generisanju i u setup-u, pored `Backups`/`Journal`). Centralni helper `EnsureDocFolder` + `PDF_DIR_*` konstante (`modConfig`/`modSetup`); imena fajlova i vremenski pečat ostaju nepromenjeni (npr. `Izveštaj_…pdf` u folderu `Izvestaji`).
- **Matični podaci (meni) — grupisane sekcije:** umesto ravne liste dugmadi, sekcije su grupisane pod naslovima — **Šifarnici** (Kooperanti, Stanice, Kupci, Vozači, Parcele), **Proizvodi i cene** (Artikli, Kulture, Cenovnik, Vrsta got. proizvoda), **Ambalaža i pakovanje** (Ambalaža, Palete, Kutije, Kese) i **Sistem** (Podešavanja) — pa ambalaža/palete/kutije/kese više nisu ravnopravne sa osnovnim šifarnicima, već svoja podgrupa. Sve runtime (`modMaticniLookups.MaticniSekcijeGrupisano`), `.frx` se ne dira; dodavanje sekcije = i dalje jedan red.
- **Matični podaci → „Admin" (operativne/razvojne komande):** nova sekcija **Admin** u grupi „Sistem" (ispod „Podešavanja") otvara runtime panel sa komandama grupisanim po nameni — **Ažuriranje** (ručna provera self-update-a, sa porukom i kad nema novije verzije), **Setup i provere** (agregatni „Ensure" = `SetupNewPC` + sve šeme; Health check setup; Production health check), **Google/Drive** (Google autorizacija; Objavi release na Drive), **VBA (dev)** (Import/Export; Otvori VBA editor) i **Podaci (oprezno)** (Migracija iz starog fajla; Očisti tabele — uz potvrdu). Svako dugme poziva **postojeću** ulaznu tačku (bez nove logike); panel je runtime (`Controls.Add` + `clsAdminBtn` WithEvents, isti obrazac kao Podešavanja), `.frx` se ne dira. Ručna provera ažuriranja čita `AgriX_Release/version.json` preko `modUpdateGate.ReleaseManifestVersion` (ne dira `modSelfUpdate`/`SKIP_MODULES` → self-update-safe), a samo ažuriranje ide preko `Application.OnTime "RunSelfUpdate"` (nikad direktno).
- **Podešavanja — sklopive sekcije + dvokolonski raspored:** umesto jedne duge kolone svih polja, svaka grupa je sada sklopiva sekcija (header `[+]/[-]`, uz „Raširi sve"/„Skupi sve"); polja unutar grupe idu u 2 kolone (memo preko celog reda). Početno su sve grupe sklopljene (pregledan „sadržaj"); čuvanje radi nepromenjeno bez obzira na stanje sklapanja.
- **Podešavanja — preuređene grupe:** redosled „operativno gore → tehnika dole" (Prodavac, Otkup/dokumenta, Štampa, Malina režim, Management/Klijent, SEF, pa Monitoring, Sinhronizacija, Google, Alati, Napredno); **Licenca** i **Probni period** spojeni u **Management/Klijent**. „Otkup/dokumenta" raspoređen u dve kolone — vrednosti/podešavanja levo (klauzula, rok, vrsta/sorta, tip palete, bruto, PDV, praćenje), toggle-i desno (filter/auto-kreiraj kooperanata, auto-broj, keš, paletiranje, panel). „Auto otpremnica+zbirna+prijemnica" (autohladnjača) prebačen u grupu „Malina režim" (uz Malina toggle i podrazumevanog kupca/hladnjaču).

- **Svi štampani dokumenti — jedinstven izgled:** faktura, kartica kooperanta, kartica ambalaže, sledljivost i specifikacija sada imaju isto zaglavlje firme + naslov + stilizovanu tabelu kao otkupni/paletni list (ranije je svaki izgledao drugačije, neki bez zaglavlja firme).
- **Generisani šabloni (kraj „ručnih" sheet-ova):** `FakturaSablon`, `KarticaSablon`, `KarticaAmbalazeSablon`, `SledljivostSablon` i `SpecifikacijaSablon` se prave automatski iz koda i same se obnavljaju na promenu rasporeda — nema više greške „sheet ne postoji" niti ručnog pravljenja šablona; print logika sledljivosti izmeštena iz forme u `modPrint`.
- **Faktura — broj se prikazuje ispravno:** broj fakture (npr. `1/2026`) više se ne tumači kao datum („jan.26").
- **Konfigurabilan izlaz za sve dokumente:** faktura/kartica/kartica ambalaže/sledljivost/specifikacija poštuju `*_PRINT_MODE` (PDF/PRINT/PREVIEW/OFF) kao otkupni/paletni; podrazumevano ponašanje nepromenjeno.
- **Podešavanja — grupa „Štampa" (svi dokumenti):** režim štampe po dokumentu (PDF/PRINT/PREVIEW/OFF) na jednom mestu, za **svih 12** dokumenata: Otkupni list, Grupni otkupni list, Prijemnica, Otpremnica, Paletni list, Paletni list got. proizvoda, Revers ambalaže, Faktura, Kartica kooperanta, Kartica ambalaže, Sledljivost, Specifikacija. Otpremnica i paletni list gotovih proizvoda (ranije „uvek PDF") i grupni otkupni list (ranije delio režim sa otkupnim) sad imaju svoj `*_PRINT_MODE` ključ kroz centralni izlazni dispečer (`DocResolveMode`/`DocPrintWs`); podrazumevano ponašanje nepromenjeno (grupni: prazno → prati otkupni). Eksplicitni dvoklik-reprint preradnog lista i dalje je uvek PDF.
- **Prijemnica — labela:** „Vraćena ambalaža (kom)" preimenovana u **„Izdata ambalaža (kom)"** (samo natpis na štampanom dokumentu; interno polje `KolAmbVracena` i logika nepromenjeni).
- **Interno (čišćenje):** zajednički `modDocStyle` helperi (`DocExportPdf`, `DocPageSetupThirdA4`, `DocResolveMode`, `DocPrintWs`) uklanjaju dupliranje PDF-izvoza/PageSetup-a/izlaznog dispečera; imena šablona kao `WS_*_SABLON` konstante; sav print/template kod u modPrint grupi.
- **Revers OM↔firma (preko vozača) — „Izdato OM" / „Prijem od OM":** revers za kretanje prazne ambalaže između firme (hladnjače) i OM-a, koje uvek ide preko vozača. Dva nova toggle-a u „Ulaz OM" frejmu (uz postojeće koop), sva četiri smera međusobno isključiva. Knjiži se **OM (Stanica) noga + vozač** (vozač se razdužuje pri raspodeli praznih na OM, odn. zadužuje pri povratu sa OM-a; utovar praznih kod kupca već postoji kroz prijemnica-povrat / „Kupci izlaz", pa se hladnjača ne knjiži duplo). PDF u istom 1/3-A4 formatu kao koop revers (protivpartner = **vozač**, potpisi OM/vozač, saldo = stanje praznih gajbi na OM-u), broj = `OM/ddmmyy[-N]` (deli `KIND_REV` namespace); poštuje `OM_IZDAVANJE_PRINT_MODE` (Podešavanja → Štampa → „Revers ambalaže"). Storno kategorije **„Revers izdato OM (firma)"** / **„Revers prijem od OM (firma)"** + reprint iz Izveštaja (vozač rekonstruisan iz ledgera). Novi tipovi `DOK_TIP_OM_ULAZ_FIRMA` / `DOK_TIP_OM_IZLAZ_FIRMA`.
- **„Ulaz OM" frejm — preuređen raspored + sopstvene labele:** redosled je sada Broj dokumenta → **Kooperant** (labela umesto „Primalac novca") → Tip ambalaže → Količina ambalaže → **[4 revers toggle-a 2×2]** → Novac → Preostali keš (Br. otk. blk.) → Iz OM avansa → Unos. Stare `.frx` labele se kriju, prave se sopstvene (runtime, `RelayoutOMUlaz`/`MakeLbl`), font usaglašen sa ostatkom forme; frejm se produži da stane celo dugme „Unos OM ulaz"; **tab order prati raspored**.
- **Keš isključen → ceo keš-paket onemogućen:** kad je „Postoje keš isplate" NE, u „Ulaz OM" su uz „Br. otk. blk." (preostali keš) sada onemogućeni i **Novac** i **Iz OM avansa** (sivo/zaključano), umesto samo „Br. otk. blk.".
- **Otkup — datum se ne „zaglavljuje" sa izabrane otpremnice:** klik na otpremnicu u panelu „Otkupni blokovi" popuni levu formu njenim podacima (datum, broj zbirne, vrsta/sorta…). Sada se pri **napuštanju te otpremnice** — *Sakrij blokove* ili promena otkupnog mesta na drugu stanicu (npr. direktan unos na hladnjači) — **datum vraća na današnji**, **broj zbirne briše**, a **vrsta/sorta voća vraćaju na podrazumevani proizvod** (iz Podešavanja). Ranije je svež unos mogao tiho da nasledi datum sa stare otpremnice — tada je „odlazio u prošlost" i u koloni Datum i u `ddmmyy` delu brojeva otkupnog lista/otpremnice/zbirne/prijemnice (prefiks i sekvenca su ostajali ispravni za stanicu, pa je delovalo da je „sve usklađeno sa hladnjačom osim datuma"). Reset se okida samo dok postoji veza sa otpremnicom iz panela; normalan i namerno unazad-datiran unos su netaknuti.
- **Otkup — kooperant se čisti posle „Unos":** posle snimanja otkupa polje Kooperant se prazni i fokus ide na njega (sledeći unos = novi kooperant); ranije je ostajao popunjen.
- **Otkupni blokovi (panel) — kolona „Kupac" prikazuje kooperanta (malina mod, OM-hladnjača):** u listi OTPREMNICE, kada je **malina mod** uključen i otkupno mesto otpremnice je označeno kao **hladnjača** (`tblStanice.JeHladnjaca = "Da"`), kolona **„Kupac"** umesto naziva firme-hladnjače prikazuje **ime kooperanta** vezanog otkupnog bloka (svaka takva otpremnica = 1 kooperant po otpremnici/prijemnici). Ostale otpremnice (drugi OM) i rad **van malina moda** ostaju nepromenjeni. Lookup (hladnjača-stanice + `OtpremnicaID→KooperantID` + ime) gradi se samo u malina modu; schema-robustno (ako `JeHladnjaca` ili koop veza fali → ostaje kupac/firma kao i pre).

---

## vba-v2.8.0 — 2026-06-30
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Donosi modul **Korisnici** (prijava + prava po oblasti) i **audit trag**. Sve je **opt-in** — dok admin ne uključi, aplikacija radi kao i pre (bez prijave).

- **Modul „Korisnici" (admin + korisnici, prava po oblasti):** novi sistem prijave gde **admin** ima sva prava, a običnim korisnicima admin odobrava pristup **po oblasti** (12: Otkup, Dokumenta, Agrohemija, Izveštaji, Fakturisanje, Banka, Marža, Sledljivost, Matični podaci, Palete, Otvori Excel, Sinhronizuj PWA). Model „kolona po oblasti = DA/NE" (`tblKorisnici`); korisnik vidi/otvara samo ono za šta ima pravo, admin sve (bypass). Uključuje se preko `Alt+F8 → EnsureKorisniciSchema → KreirajPrvogAdmina → EnableAuth` (uključenje je blokirano dok ne postoji bar jedan aktivan admin — zaštita od zaključavanja).
- **Prijava (frmLogin):** na startu (kad je prijava uključena) traži korisničko ime + **PIN** (maskiran), uz limit od **3 pokušaja**. Neuspela/otkazana prijava zatvara aplikaciju.
- **Administracija korisnika kroz UI (Matični → Sistem → Korisnici):** dodavanje/izmena korisnika; **Uloga**, **Aktivan**, **Stanica** su padajuće liste, a **oblasti** su desna kolona **DA/NE** padajućih lista. Uloga **Admin** automatski postavlja sve oblasti na DA i zaključava ih. Sekciju „Korisnici" vidi/otvara samo admin.
- **PIN hashing (podrazumevano UKLJUČENO):** PIN se čuva kao SHA-256 heš; postojeći „goli" PIN-ovi se transparentno migriraju na heš pri prvoj prijavi. Bezbedan fallback: ako SHA (.NET) nije dostupan, upis/provera padaju na plaintext (bez rizika od zaključavanja). Po potrebi se isključuje sa `Alt+F8 → DisablePinHash` (ne kvari već hešovane PIN-ove).
- **Audit trag (timestamp + userstamp) na svim glavnim tabelama:** `Alt+F8 → EnsureAuditColumns` dodaje kolone **CreatedAt/CreatedBy** (unos) i **ModifiedAt/ModifiedBy** (svaka izmena) na 26 tabela. Pečat se upisuje automatski iz centralnog sloja podataka (`AppendRow`/`UpdateCell`) — vidi se **ko** je i **kada** uneo/izmenio red. Korisnik = prijavljeni app-korisnik → (ako prijava nije uključena) Windows nalog.
- **Operater u gornjoj traci = ime i prezime:** umesto Windows naloga, gore piše **ime i prezime** prijavljenog korisnika.
- **Odjava / zamena korisnika u toku rada:** klik na „Operator: …" u gornjoj traci odjavljuje trenutnog i otvara prijavu za drugog korisnika (isti tok kao paljenje aplikacije) — bez ručnog gašenja.
- **„Ensure" dugme (Admin panel) priprema i Korisnike:** agregatni Ensure dodatno radi `EnsureKorisniciSchema` + `EnsureAuditColumns` (jednim klikom napravi tabelu korisnika i audit kolone).
- **Migracija prenosi korisnike (foolproof):** „Migracija iz starog fajla" sada na početku sama osigura `tblKorisnici` + audit kolone, pa se korisnici (i PIN, uloga, prava po oblasti) prenose iz starog u novi fajl i bez ručnog „Ensure".
- **Robusnost:** unos i izmena korisnika su atomični (transakcija — nema „pola reda" ako neki upis padne) i pišu **po imenu kolone** (drift-safe, otporno na dodavanje oblasti/audit kolona); duplikat korisničkog imena je sprečen i pri unosu i pri izmeni.
- **Uputstvo:** `docs/UPUTSTVO_KORISNICI.md` (uključivanje, rad, zaboravljen PIN, admin lockout, oporavak).

---

## vba-v2.8.1 — 2026-06-30
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Patch: **blokator obaveznih polja** na unosima — sprečava snimanje/preradu dok sva relevantna polja nisu popunjena (operater dobije jasnu poruku i fokus na polje koje fali).

- **Otkup (`frmOtkup`):** uz postojeće obavezne (otkupno mesto, kooperant, vrsta voća, datum, količina, cena) sada su obavezni i **sorta voća**, **broj gajbi (ambalaža)** za svaku unetu klasu (ranije se tražio samo u bruto režimu) i **broj dokumenta**.
- **Otpremnica (`frmDokumenta`):** obavezni **vrsta + sorta voća**, **cena I** (Klasa I, > 0), **broj gajbi (I/II) + tip ambalaže** i **broj dokumenta**.
- **Zbirna:** obavezni **vrsta + sorta voća** i **tip ambalaže** (kad je uneta ambalaža); **Hladnjača i Pogon ostaju opcioni**.
- **Prijemnica:** obavezni **vrsta + sorta voća**, **cena I** (> 0) i **broj gajbi (I/II) + tip ambalaže** (vraćena ambalaža ostaje opciona).
- **OM Ulaz / Izlaz (tok ambalaže/novca):** obavezni **vozač**, **vrsta voća** i **broj dokumenta** (kod OM Ulaza smer-tokovi i dalje sami predlažu broj; pun blok tek ako ostane prazno).
- **Palete — prerada (`frmPalete` → „Preradi izabrane"):** uz već postojeći izbor palete sada blokira ako su prazni **bruto**, **težina palete**, **gotov proizvod**, **broj + tip kutija** i **broj + tip kesa** — sprečava preradu sa neto = 0 i neoznačen izlazni proizvod.
- **Bez novih zavisnosti:** sve poruke su **ASCII-only** u izvoru (inline `ChrW` za dijakritiku), uz reuse postojećeg `Poruka("DOK_MSG_IZABERITE_TIP_AMBALAZE")` — nema novih katalog ključeva, pa posle importa **ne treba `EnsurePoruke`**.

---

## vba-v2.8.2 — 2026-06-30
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Patch: **specifikacija dnevnog/periodičnog otkupa dobija kolonu „Kupac"** (firma kome ide roba).

- **Specifikacija otkupnih blokova — kolona „Kupac":** dnevna/periodična specifikacija (Otkup → „Štampaj po datumu") i specifikacija ručno izabranih otpremnica („Biraj otpremnice" → „Štampaj specifikaciju") sada uz „Broj zbirne"/„Broj otpremnice" prikazuju i **kupca (firmu) kome ide roba** — izvedeno lancem `BrojZbirne → tblZbirna.KupacID → tblKupci.Naziv` (reuse postojećeg `KupacNazivZaZbirnu`, isti podatak kao kolona „Kupac" u listi otpremnica). Ubačena je kao **druga kolona** (uz „Broj zbirne").
- **Raspored i dalje staje na 1 A4 landscape:** šablon već koristi `FitToPagesWide = 1` (garantovano 1 strana po širini); uz blagi trim širina ostalih kolona (Ime i Prezime 24→20, Otkupno mesto 18→16, Broj otpremnice 14→12, Datum 11→10…) ukupna širina je ~152 jedinice (pre 147), pa su auto-skaliranje i čitljivost praktično nepromenjeni. `SpecifikacijaSablon` se sam obnavlja (LAYOUT_VER 1→2 → stari šablon se prepravlja pri prvom otvaranju).
- **Bez novih zavisnosti:** izvori (`modOtkupBlok`, `modPrint`) ostaju **ASCII-only**; nema novih katalog ključeva (`Poruka()`), pa posle importa **ne treba `EnsurePoruke`**.

---

## vba-v2.8.3 — 2026-06-30
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Patch: **performanse** — brže otvaranje/zatvaranje formi i generisanje izveštaja, **bez promene rezultata**; uz uklanjanje mrtvog koda (`tblRpt*`).

- **Brže otvaranje i zatvaranje sekcija (glavni meni):** prebacivanje između sekcija (`OpenContentForm`) i gašenje aplikacije (`ShutdownApp`) sada gase `Application.ScreenUpdating` dok se forme prikazuju/uklanjaju — ranije su se runtime kontrole iscrtavale jedna po jedna. Pri zatvaranju `frmOtkup` poziva se `OtkupBlok_Release` (postojeća rutina) da oslobodi ~35 dinamičkih kontrola i `WithEvents` objekata pre `Unload` (ranije se čistilo samo pre self-update-a → gomilali su se).
- **Izveštaji — prelaz „Otkupna mesta ↔ Kooperanti" više ne radi dupli posao:** `LoadEntiteti` je postavljanjem `ListIndex=0` okidao `AutoRefresh`, a `tgl*_Click` ga je zvao još jednom → **ceo izveštaj se generisao dvaput** po prelazu. Sada se generiše jednom (ručni izbor kooperanta i dalje osvežava).
- **Izveštaji — brže generisanje:** dok traje jedan „Prikaži", svaka tabela se čita **jednom** (request-scoped keš u `modDataAccess`) umesto 17–18 puta (npr. `tblOtkup` se čitao 5×). Dodatno, imena (kooperant/vozač/artikal) se rešavaju iz mape napravljene jednim prolazom, umesto linearnom pretragom po svakom redu rezultata. Rezultat izveštaja je **identičan** — menja se samo koliko puta se podaci čitaju.
- **KPI na dashboardu („Današnji otkup kg") se osvežava samo kad treba:** posle unosa u `frmOtkup`/`frmDokumenta` ili PWA importa (sve ide kroz `AppendRow`), a ne pri svakom povratku na dashboard.
- **Uklonjen mrtav kod — `tblRpt*` izveštajne tabele:** `WriteReportTables` i `WriteMarza` su pri svakom „Prikaži" pisali u `tblRptSaldoOM/Kupci/Marza/Zbirni` **redom-po-red kroz `AppendRow`** (uz CSV-journal na disk po redu), a te tabele se **nigde ne čitaju** (`modMigracija` ih ionako preskače kao izvedene). Upis je uklonjen; same tabele/sheetovi su ostavljeni netaknuti.
- **Bez promene podataka i bez novih zavisnosti:** journaling **stvarnih** unosa (`tblOtkup`, `tblNovac`, `tblOtpremnica`…) je nepromenjen; izvori ostaju **ASCII-only**, nema novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.

---

## vba-v2.8.4 — 2026-06-30
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Agrohemija: **izdavanje bez parcele** kad je praćenje parcela isključeno + **unos početnog duga kooperanta** (migracija) bez vezivanja za artikal.

- **Izdavanje robe bez parcele (kad je praćenje parcela OFF):** ako je u Podešavanjima `Praćenje parcela` isključeno (`PRACENJE_PARCELA`), Agrohemija sada prihvata izdavanje robe **bez odabira parcele** — lista parcela je zaključana, a „smart" preporuka doze po hektaru se preskače (broj pakovanja se unosi ručno). Kad je praćenje uključeno, sve radi kao i pre (obavezna parcela + preporuka). Isti prekidač koji već postoji u Otkupu (`IsPracenjeParcela`).
- **Početni dug kooperanta (migracija) — dugme „Početni dug":** novo dugme iznad „Završi izdavanje" knjiži **čist iznos duga u RSD** za izabranog kooperanta, bez biranja realnog artikla i bez greške „Nedovoljno stanje". Knjiži se kao jedna stavka izdavanja na rezervisani interni artikal (`ART_POCETNI_DUG`), pa se dug **konzistentno** vidi i u Agrohemiji i u kartici/saldo izveštaju kooperanta. **Reverzibilno** preko storniranja te stavke. Rezervisani artikal je sakriven iz lista artikala i iz pregleda stanja magacina (nema fantomskog negativnog stanja).
- **Dug se prikazuje i za kooperanta bez parcela:** ranije se, kad praćenje parcela radi, dug nije prikazivao ako kooperant nema unetu parcelu — sada se prikazuje uvek.
- **Bez novih zavisnosti:** izvori (`frmAgrohemija`, `modAgrohemija`, `modConfig`) ostaju **ASCII-only** (dijakritika u prikazu preko inline `ChrW`); nema novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.
- **Poznato ograničenje:** stavke početnog duga se za sada **šalju i na PWA** (`ExportMagacinKoop`) sa fantomskom količinom 1 — biće filtrirano u sledećem patch-u (vidi ROADMAP / KI-006).

---

## vba-v2.8.5 — 2026-06-30
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Patch: **performanse (nastavak 2.8.3)** — dalje ubrzanje izveštaja i otvaranja formi, **bez promene rezultata**; uz ispravku KPI „kg danas".

- **Izveštaji — punjenje liste jednim prenosom (`.List = arr`):** sve liste u `frmIzvestaj` se grade kao 2D niz u memoriji pa prenose **jednom** u ListBox, umesto ćeliju-po-ćeliju (`.List(i, j)` = stotine/hiljade COM poziva po izveštaju). Mereno na „Otkupljena roba": **922 ms → 16 ms**; otvaranje `frmIzvestaj` sa ~15 s na **~0,3 s**.
- **Izveštaji — lazy generisanje po tabu:** pri otvaranju i promeni entiteta generiše se **samo aktivan tab**; ostali tek kad korisnik klikne na njih (keširani do sledećeg „Prikaži"). Najskuplji izveštaj (Ambalaža, ~550 ms nad celim ledgerom) se više ne računa na svakom otvaranju ako se taj tab i ne gleda. Rezultat svakog taba je **identičan**.
- **Izveštaji — manje skeniranja u istom „Prikaži" prozoru:** `GetColumnIndex` i `ExcludeStornirano` se sada keširaju u istom request-scoped bloku (uz `tblData` keš iz 2.8.3). `ReportOtkupListe`/`ReportAmbalaza` više ne prave **dodatnu kopiju cele tabele** za storno — storno se gleda inline u petlji, odnosno kao filter u jednom prolazu.
- **Forme — tema/stilizacija jednom po instanci:** `frmIzvestaj` je modeless, pa je `UserForm_Activate` na **svaki povratak fokusa** radio pun (i dupli) obilazak stabla kontrola — sada jednom po instanci, uz uklanjanje duplog `StyleControls`. `frmDokumenta`: ista dedup teme (~70 ms brže otvaranje).
- **Ispravka — KPI „kg danas" (`SumOtkupKgToday`):** jedna ćelija sa Excel greškom (`#N/A`/`#VALUE!`) u `tblOtkup` je obarala **ceo** dnevni zbir na 0 (Type mismatch). Sada se takve ćelije preskaču po redu (`IsError` guard) — KPI je pouzdan.
- **Bez promene podataka i bez novih zavisnosti:** izvori ostaju **ASCII-only**, nema novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.

---

## vba-v2.8.6 — 2026-07-01
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Podešavanje: **prekidač „Kompletna validacija unosa"** — pali/gasi utegnutu proveru obaveznih polja uvedenu u 2.8.1.

- **Podešavanja — toggle „Kompletna validacija unosa (obavezna polja pre snimanja)":** novo podešavanje u grupi **„Otkup / dokumenta"** (DA/NE, default **DA**). Kontroliše „Blokator obaveznih polja" iz 2.8.1 u `frmOtkup`, `frmDokumenta` i `frmPalete`.
  - **DA (kao dosad):** pre snimanja su obavezni svi polja iz 2.8.1 — sorta voća, broj gajbi i tip ambalaže po klasi, cena I klase, broj dokumenta, vozač; u preradi (`frmPalete`) bruto / težina palete / gotov proizvod / broj + tip kutija i kesa.
  - **NE (kao pre 2.8.1):** minimalna validacija — ta polja više ne blokiraju snimanje. U `frmOtkup` broj gajbi ostaje obavezan **isključivo u BRUTO režimu** (bez toga se bruto ne pretvara u neto) — kao i pre te izmene.
- **Pre-postojeće provere se ne diraju:** obavezni kupac / otkupno mesto / datum / količina i drugi ranije postojeći uslovi ostaju **uvek aktivni**, nezavisno od prekidača — prekidač gasi samo ono što je 2.8.1 dodala.
- **Default DA → postojeće instalacije rade identično** kao na 2.8.1 dok se prekidač ručno ne postavi na NE.
- **Bez novih zavisnosti:** data-driven kroz postojeći editor podešavanja (`modPodesavanja`) + `IsValidacijaUnosa()` u `modConfig` (`ConfigFlag`, default ON); izvori ostaju **ASCII-only**, nema novih `Poruka()` ključeva (reuse postojećih) → posle importa **ne treba `EnsurePoruke`**.

---

## vba-v2.8.7 — 2026-07-01
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Fokus: **puštanje bankarskog importa u rad + zaokruženo prvo podešavanje računara** (config putanje, first-run, provere veze). Novi `Poruka()` ključ → posle importa **`EnsurePoruke`**.

- **Banka import (izvodi) — end-to-end:** Gmail → Drive → lokalni disk → `pdftotext` → `tblBankaImport` → `tblNovac`. GAS downloader (multi-client) `gas/bank-pdf-downloader/` puni deljeni `01_Bank` folder; VBA povlači, parsira (Komercijalna banka, saldo-integrity) i stage-uje.
  - **Dva prioritetna auto-map ključa (jača od imena):** `poziv na broj` (= broj otkupnog lista za isplate / broj fakture za uplate) i `tekući račun` partnera. Normalizacija tolerantna na format (`205-...-XX`, gole cifre, model).
  - `frmBankaImport`: **auto-map na otvaranje** po jakim ključevima, brojač **„Mapirano X / Y"** prati stvarno stanje, zaglavlja kolona iznad liste.
- **Podešavanja — grupa „Banka / lokalno" + inline „…" browse dugmad:** per-mašina putanje (`PDFTOTEXT_EXE_PATH`, `BANKA_DRIVE_SOURCE_PATH`, Inbox/Processed/Error, `BANKA_DRIVE_*`) sada se rutiraju u **`tblLocalConfig`** (ranije su odlazile u `tblSEFConfig` a čitane iz Local → polje „nije radilo"). Svako path-polje ima folder picker; poppler dugme = `SetupPopplerInteractive` (auto pored xlsm-a ili picker). Poppler default se računa relativno na radnu svesku.
- **Prvi start / setup:** na otvaranju, ako računar nije podešen (`APP_SETUP_COMPLETED != DA`), aplikacija ponudi `SetupNewPC` (jednokratno). **SEF je opcion** — ako sva SEF polja prazna, provera se preskače (ne blokira „zeleno" setup). **Google config** se čita iz `tblSEFConfig` (kao runtime) — nestao lažan „Nedostaje GOOGLE_… u tblConfig".
- **Provera veze desktop↔server (advisory):** `RunSetupHealthCheck` / `TestServerLink` / Admin dugme „Health check (setup)" proveravaju živi Google OAuth token, GAS monitoring i banka Drive folder. U `SetupNewPC` je samo NAPOMENA — offline ne obara setup.
- **Fix:** `RunProductionHealthCheck` je tražio kolonu `Količina` (dijakritika) umesto ASCII `Kolicina` → lažan „Missing column: tblOtkup.Kolicina" na ispravnoj šemi.
- **Usklađena folder struktura + docs:** `00_Inbox/01_Bank` + `Downloaded` (GAS `DriveFolder.gs`, uklonjen mrtvi `02_Bank_Izvodi`), `Processed` umesto `Verarbeitet`, `Setup-OtkupApp.ps1` kopira `Tools\poppler` pored sveske; runbook/onboarding/`CLAUDE.md` usklađeni.
- **Encoding:** izvori ostaju **ASCII-only**; jedini novi `Poruka()` ključ je `SETUP_MSG_FIRSTRUN_PONUDA` (dijakritika kroz `ChrW`) → posle importa pokrenuti **`EnsurePoruke`**.

---

## vba-v2.8.8 — 2026-07-02
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Fokus: **pregled prometa kooperanta u panelu „Otkupni blokovi"**. Bez promene podataka; izvori ostaju **ASCII-only**, nema novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.

- **Otkup / „Otkupni blokovi" panel — „Otk.listovi: <iznos> RSD":** na liniji „Ostatak" (između sažetka i dugmeta „Sakrij"), nova info pokazuje **ukupan iznos izdatih otkupnih listova** (Σ Količina × Cena = „Ukupna vrednost", bruto sa PDV nadoknadom) za **trenutno izabranog kooperanta u tekućoj godini**; osvežava se na promenu kooperanta i pri otvaranju panela. Slobodan unos / bez izbora → prazno (bez skeniranja, bez auto-kreiranja kooperanta).
- **Dugme „Lista kooperanata":** ispod info; otvara overlay preko **celog panela** sa **svim kooperantima firme, sortiranim opadajuće** po istom iznosu u tekućoj godini. Kolone: `# | Kooperant | OM | Iznos (RSD)` (OM = matična stanica kooperanta). „Zatvori" vraća na panel.
- **Bez novih zavisnosti / bez promene podataka:** sve kontrole su dinamičke (`Controls.Add` + `clsBlokUI`, `frmOtkup.frx` se ne dira); reuse postojećih helpera (`ExcludeStornirano`, `BuildKoopNames`, `BuildLookup`, `FmtRsd`, `GetComboID`). Sažetak (Ukupno / U blokovima / Ostatak) blago sužen i pomeren levo da oslobodi mesto za info na istoj liniji. Izvori **ASCII-only**, nema novih `Poruka()` ključeva.

---

## vba-v2.9.0 — 2026-07-02
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Fokus: **skrolovanje listi točkićem miša** na svim formama sa `ListBox`-om. Izvori ostaju **ASCII-only**, bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.

- **Točkić miša skroluje liste:** MSForms `ListBox` po defaultu ne prima `WM_MOUSEWHEEL` (traka radi, točkić ne) — sada radi na svim ekranima sa listama: Izveštaji (kartica + pod-tabovi), Dokumenti (zbirne + storno/recovery), Palete, Otkup „Otkupni blokovi" panel (otpremnice / blokovi / rang kooperanata), Matični podaci, Agrohemija, Fakturisanje, Banka export, Sledljivost, SEF. Novi `modMouseWheel` + `clsWheelList` (isti „`WithEvents` omotač oko dinamičke kontrole" idiom kao `clsBlokUI`).
- **Uključivanje kroz Podešavanja (per-mašina):** nova grupa **„Interfejs / lokalno"** → „Skrolovanje listi točkićem miša" (DA/NE, `tblLocalConfig`, prazno = **DA**). `StartApp` ga čita na svakom pokretanju i pali/gasi; snimanje u Podešavanjima primenjuje **odmah** (bez restarta). Ručno / dijagnostika (Alt+F8): `MouseWheel_On` / `MouseWheel_Off` / `MouseWheel_Reset`.
- **Bezbedan dizajn (low-level hook, fail-safe):** hook se diže **lenjivo** — tek kad miš pređe preko liste (ne pri otvaranju forme → nema „belog ekrana") — i **sam se skine** čim miš siđe s liste (kursor van liste ostaje gladak) ili kad forma izgubi fokus (`Deactivate`/`QueryClose`/`Terminate`). **Ne diže se dok je VBE otvoren** (izbegava zamrzavanje tokom razvoja). Bira listu pod mišem preko `MouseMove` (bez geometrije/DPI). `.frx` se ne dira; OFF-safe dok se ne „naoruža".
- **Self-update kompatibilno:** `modMouseWheel` / `clsWheelList` (module-level `MSForms` deklaracije / `WithEvents`) idu kroz postojeći dvofazni `Remove`+`Import` (rutiranje je error-driven, nema hardkodirane liste); higijena `MouseWheel_Off` u `PrepareRuntimeForSelfUpdate` i `ShutdownApp`. `docs/SELF_UPDATE.md` dopunjen + smoke-test checklist posle release-a.

---

## vba-v2.10.0 — 2026-07-02
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Fokus: **multi-bank parser dispatch za bankarski uvoz izvoda (ProCredit + Halkbank)**. Izvori ostaju **ASCII-only**, bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.

- **Multi-bank uvoz izvoda:** `ParseBankaIzvodForImport` sada prepoznaje banku iz `pdftotext` teksta (`DetectBank`) i rutira na parser te banke (`Select Case`), uz **deljeni 4-nivo integrity + 17-kolonski staging** koji se ne diraju. Pored **Komercijalne** (`modBankaImportParserPdfToText`), dodati **ProCredit** (`modBankaProCredit`, računi `220-…`) i **Halkbank** (`modBankaHalk`, računi `155-…`).
  - **ProCredit:** validiran na stvarnom izvodu — integritet (sume uplata/isplata + broj naloga) se poklapa sa bančinim sopstvenim totalima.
  - **Halkbank:** rukuje **dvama datumima** (izvršenja/prijema), saldo blokom **u sredini** dokumenta, i **isključuje sekciju `NEIZVRŠENI NALOZI`** (nalozi na čekanju koji ne ulaze u saldo) — bound na „Ukupno na računu". Validiran.
- **„Banka uvoz izvoda" jednim klikom povlači sa Drive-a:** dugme (`frmOtkupAPP.btnBanka`) sada zove `ImportBankaInbox_WithDrivePull` — prvo **dovuče nove PDF-ove sa Drive-a** u lokalni Inbox pa uveze. **Backward-safe:** pull samo ako je `BANKA_DRIVE_SOURCE_PATH` podešen, inače identično dosadašnjem lokalnom uvozu.
- **Operater-alat `Diag_DumpFullPdfText` (Alt+F8):** izvuče **pun `pdftotext` izlaz** izabranog PDF-a u `.txt` pored njega (isti flagovi kao uvoz) — za slanje uzorka nove banke bez komandne linije. Uz bank-agnostic `Test_BankParse` (detekcija + integritet + per-red dump).
- **Dev vodič `docs/development-banka-parser.md`:** korak-po-korak dodavanje parsera za novu banku (ugovor 5 funkcija, dispatch, integritet=validacija, naučene zamke, referentni primeri). CLAUDE.md „code-map" dopunjen multi-bank dispatch-om.
- **Bez novih zavisnosti / bez promene podataka:** novi `.bas` moduli se automatski kreiraju kroz self-update (`VBComponents.Add`, bez ručnog transfera); izvori **ASCII-only**, nema novih `Poruka()` ključeva.

---

## vba-v2.11.0 — 2026-07-03
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Fokus: **priprema naloga za plaćanje (banka)** — CSV nalozi za prenos + PDF specifikacija isplata iz sekcije „Banke platni nalozi". Jedan postojeći `Poruka()` ključ menja tekst → posle importa pokrenuti **`EnsurePoruke`**.

- **„Generiši CSV naloge"** (`frmBankaExportPregled`): od selektovanih blokova (ili svih prikazanih) pravi **CSV naloga za prenos** za uvoz u e-banking. Platilac = firma (`SELLER_NAME`/`SELLER_ACCOUNT` iz Podešavanja → „Prodavac (firma)"); primalac = kooperant + tekući račun; **iznos po bloku = „Isplatiti"** (unos operatera u detalju bloka, ili otvoreni iznos); **poziv na broj = broj otkupnog bloka** — jaki ključ za auto-map pri kasnijem uvozu izvoda; šifra plaćanja (prazno = 221) i svrha iz nove grupe Podešavanja **„Banka / nalozi"**. Potvrda pre upisa (broj naloga + ukupna suma); fajl ide u `Nalozi za banku\` pored radne sveske (UTF-8, `;` separator, decimalna tačka), Explorer se otvori sa označenim fajlom. Blokovi bez tekućeg računa se preskaču i prijave. **Ništa se ne knjiži u `tblNovac`** — isplata se knjiži tek kroz uvoz izvoda (postojeći tok).
- **„PDF specifikacija"** (dugme koje je ranije bilo „Export u clipboard"): **specifikacija isplata kooperantima** — house-style obrazac (zaglavlje firme + naslov), A4 portrait, kolone `R.br | Datum | Br. bloka | Kooperant | Tekući račun | Ukupno | Isplaćeno | Otvoreno | Za isplatu` + UKUPNO red; **isti izbor blokova i isti iznosi kao CSV**. Izlaz po novom podešavanju „Specifikacija isplata (banka nalozi)" u grupi Štampa (`ISPLATA_SPEC_PRINT_MODE`; prazno = PDF → otvori se; folder `Specifikacije\`). TSV clipboard export je uklonjen.
- **Filter „Kooperant"** (runtime combo u filter redu, `.frx` se ne dira): radi **i na unos i kao padajuća lista** — kucanje autocomplete-uje i filtrira listu po delu imena (substring), izbor iz liste filtrira tačno; prazno = svi. Lista imena se puni iz otvorenih blokova (poštuje datum/stanica filter). **„Isplatiti" unosi se NE gube** pri prebacivanju kooperant-filtera (prune ide protiv pune liste).
- **Combo „Sa računa"** (uz action dugmad): firma može imati **više računa u raznim bankama** — novo podešavanje `BANKA_NALOG_RACUNI` („Banka / nalozi", više računa odvojenih `;`); combo prikazuje račun + ime banke po NBS prefiksu (205 NLB Komercijalna, 155 Halkbank, 220 ProCredit…); default = `SELLER_ACCOUNT`. Izabrani račun ide u CSV kao **RacunPlatioca**, u potvrdu pre generisanja („Sa računa: …") i u zaglavlje PDF specifikacije.
- **Vezivanje virmanskog avansa na blokove:** kad je bankarska isplata kooperantu ušla kao avans (`NOV_VIRMAN_AVANS_KOOP`, bez otvorenog otkupa da se poklopi), sada se **ručno vezuje na konkretne otvorene blokove** iz iste forme. Dva dugmeta (zakačena na postojeći transakcioni motor `ApplyAvansToOtkup_TX`): **„Primeni avans na blok"** (u detalju izabranog bloka; aktivno samo kad kooperant ima avans) i **„Primeni avans (sel.)"** (na sve čekirane blokove). Potvrda pre upisa. Po vezivanju avans dobija `OtkupID` → „Otvoreno"/„Za isplatu" bloka pada, pa se nalozi generišu samo za ostatak; ako avans premaši blok, višak ostaje kao avans (postojeći FIFO + split). Ranije se avans vezivao samo automatski pri snimanju **novog** bloka (`modOtkup`) — za već otvorene blokove nije postojala UI akcija.
- **Performanse i raspored „Banke platni nalozi":** biranje/menjanje kooperant filtera je sada **trenutno** — filter je čist pregled nad već učitanim blokovima (ne čita više tabele na svaku promenu), lista se puni jednim upisom (`.List=arr`), a `BuildBlokIsplataList` gradi `KooperantID` mapu u O(n) umesto O(n²). Dugme **„Osveži"** premešteno desno u filter redu i vertikalno centrirano (više ne naleže na kooperant combo).
- **Podešavanja:** nova grupa **„Banka / nalozi"** (`BANKA_NALOG_SIFRA_PLACANJA`, `BANKA_NALOG_SVRHA`, `BANKA_NALOG_RACUNI` — poslovni config, `tblSEFConfig`) + novi red u grupi Štampa.
- **Encoding:** izvori ostaju **ASCII-only**; ključ `BANKA_LBL_GENERISI_CSV_COMMIT` sada glasi „Generiši CSV naloge" → posle importa pokrenuti **`EnsurePoruke`**.
