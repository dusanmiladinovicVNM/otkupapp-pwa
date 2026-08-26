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

---

## vba-v2.12.0 — 2026-07-03
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Fokus: **Drive pull radi na deljenom shortcut folderu** (klijentske mašine, `.shortcut-targets-by-id`). Izvori ostaju **ASCII-only**, bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.

- **Drive pull na deljenom shortcut folderu (`.shortcut-targets-by-id`):** klijentske mašine pristupaju `01_Bank`-u preko **deljenog shortcut-a** (mail nalog firme; vlasnički `ops@agrix` nikad nije na klijentu — multi-tenant izolacija), a Google Drive tu putanju izlaže kao `G:\.shortcut-targets-by-id\<id>\01_Bank`. Legacy VBA `Dir$`/`MkDir`/`Name`/`FileCopy`/`FileLen`/`FileDateTime` na toj virtuelnoj putanji **lažu ili pucaju** (greške **75** „Path/File access" i **76** „Path not found") i obarale su povlačenje sa Drive-a. Sve file/folder operacije u pull-u (`PullBankPdfsFromDriveProduction`, `BankaEnsureFolderExistsRecursive`, `MoveFileSafe`, provera spremnosti/kopiranje) prebačene su na **`Scripting.FileSystemObject`** (pouzdan na Drive virtuelnim/online-only putanjama). Radi **bez `Available offline`** i **bez logovanja vlasničkog naloga**.
- **Dugme „Banka uvoz izvoda" — fail-soft na pull:** ako Drive privremeno nije dostupan (offline, pogrešan/nedostupan `BANKA_DRIVE_SOURCE_PATH`), pull se **preskoči uz WARN** i uveze se lokalni Inbox umesto da klik padne (`ImportBankaInbox_WithDrivePull`). Sam uvoz (`ImportBankaInbox_TX`) ostaje hard.
- **Runbook `docs/production-runbook-banka-import-setup.md` (Faza 2):** klijentski shared-shortcut setup korak-po-korak — mail nalog, **Editor na `01_Bank`+`Downloaded`** (oba u `00_Inbox`), `BANKA_DRIVE_DOWNLOADED_PATH` prazno = default; nikad `ops@agrix` na klijentu. + troubleshooting za greške 75/76.
- **Bez novih zavisnosti / bez promene podataka:** samo `modBankaImport` (pull sloj) + docs; izvori **ASCII-only**, nema novih `Poruka()` ključeva.

---

## vba-v2.12.1 — 2026-07-03
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Fokus: **fix tihe korupcije `tblOtkup.OtpremnicaID`** (otkup se vezivao za storniranu otpremnicu). Izvori ostaju **ASCII-only**, bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.

- **Fix — tiha korupcija `OtpremnicaID` (otkup vezan za storniranu otpremnicu):** posle unosa otkupa panel „Otkupni blokovi" (`OtkupBlok_AfterUnos` → `LinkOtkupIDsToOtpremnica`) je vezivao svež otkup za izabranu otpremnicu iz zaostale promenljive `mActiveOtpID`, pišući **samo** `OtpremnicaID` bez provere cilja — pa je na hladnjači **pregazio** ispravnu vezu koju je auto-lanac upravo upisao, a kad je `mActiveOtpID` pokazivao na storniranu otpremnicu, otkup je završavao vezan za mrtav dokument (`BrojZbirne` ostaje tačan → nesklad se ne vidi golim okom). `LinkOtkupIDsToOtpremnica` sada: **(1)** ne vezuje na storniranu/nepostojeću otpremnicu i **(2)** vezuje samo redove sa praznim `OtpremnicaID` (ne pregazi vezu auto-lanca) — isti idiom kao `ReassignOtkupToOtpremnica_TX`. Zatečeni slučaj: OTK-01110/01111 → stornirana OTP-00577.
- **Fix — Sledljivost „Poveži":** `frmSledljivost.btnPovezi` je ručno vezivao otkup→otpremnica sirovim upisom iz modularnog niza `m_CandidateOtpIDs` (ista klasa greške, bez provere cilja u trenutku upisa); sada ide kroz proverenu, transakcionu `ReassignOtkupToOtpremnica_TX` (validira cilj + drži `OtpremnicaID`/`BrojZbirne` konzistentnim).
- **Popravka zatečenih redova (nepromenjeno, već u kodu):** detekcija kroz `RunProductionHealthCheck` (`Check_OtkupOtpremnicaCrossZbirnaLinks` + `Check_DocumentSoftDeleteReferences`) i dashboard `CheckVerwaisteDokumente` (sekcija 5); popravka kroz panel „Otkupni blokovi" → „Izgubljeni" → „Preuzmi" (re-point na ispravnu otpremnicu, čuva OtkupID/uplate/ambalažu).
- **Bez promene podataka / bez novih zavisnosti:** samo `modOtkupBlok` + `frmSledljivost` (+ backlog beleška **P1-16** za srodne MEDIUM storno-guard rupe u `LinkNovacToOtkupStrict` / `ApplyAvans*`); izvori **ASCII-only**, nema novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.

---

## vba-v2.13.0 — 2026-07-04
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Fokus: **više tekućih računa firme za isplate** (zasebna polja) + **čitljiviji nazivi bankarskih izlaza** (datum plaćanja + banka). Izvori ostaju **ASCII-only**, bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.

- **Podešavanja — do 4 računa firme za isplate (zasebna polja):** u grupi „Banka / nalozi" je jedno polje `BANKA_NALOG_RACUNI` (više računa u jednom polju, odvojeno `;`) zamenjeno **četirima zasebnim textbox-ovima** — „Račun firme 1 / 2 / 3 / 4 (isplate)" — radi preglednijeg i simetričnog unosa. Combo **„Sa računa"** u „Banke platni nalozi" (`frmBankaExportPregled`) i dalje nudi sve unete račune (uz naziv banke po NBS prefiksu), preskačući prazna polja; ako su sva tri prazna pada na stari `;`-spisak, pa na `SELLER_ACCOUNT` (**kompatibilno unazad**). Postojeća `;`-vrednost se pri prvom otvaranju Podešavanja **automatski razbije** u zasebna polja (jednokratna migracija). Bez izmena na dokumentima/SEF/PWA — glavni račun ostaje `SELLER_ACCOUNT`.
- **Bankarski izlazi — naziv fajla sadrži datum plaćanja i banku:** CSV nalozi za prenos i PDF specifikacija isplata (`frmBankaExportPregled`) su se imenovali **samo timestamp-om kreiranja** (`yyyymmdd_hhnnss`), pa je snalaženje u folderu bilo teško. Sada naziv **počinje datumom plaćanja** (`yyyy-mm-dd`, = `DatumValute` u nalozima) **i nazivom banke platioca**, uz kratak `hhnnss` na kraju radi jedinstvenosti (regeneracija istog dana istom bankom **ne pregazi** raniji fajl):
  - CSV: `Nalozi_za_prenos_2026-07-04_Banca-Intesa_143022.csv`
  - PDF: `Specifikacija_isplata_2026-07-04_Banca-Intesa_143022.pdf`

  Banka se izvodi iz izabranog računa „Sa računa" (`BankaNazivZaRacun`); **nepoznat/prazan račun → naziv sadrži samo datum**. Naziv banke se sanitizuje za ime fajla (SR dijakritika → ASCII, razmaci → `-`), npr. „Poštanska štedionica" → `Postanska-stedionica`.
- **Bez promene podataka / bez novih zavisnosti:** samo `modConfig`, `modPodesavanja`, `modBankaExportPregled`, `frmBankaExportPregled`; izvori **ASCII-only**, nema novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.

---

## vba-v2.14.0 — 2026-07-04
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Fokus: **zaštita dve osetljive Admin komande od slučajnog klika** — šifra za objavu release-a, potvrda kucanjem za čišćenje tabela. Izvori ostaju **ASCII-only**, bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.

- **„Objavi release na Drive" — traži šifru:** dugme u Admin panelu (grupa „Google / Drive") je build/dev komanda koja objavljuje kod **celom fleetu**; ranije ga je štitio samo Da/Ne dijalog, pa je jedan slučajan klik operatera mogao da pokrene objavu. Sada `AdminPublishToDrive` traži **šifru** (`RELEASE_PUBLISH_SIFRA`, nova konstanta u `modConfig`) preko `InputBox`-a — unos šifre je ujedno potvrda (jedan dijalog): prazno/Cancel = tiho odustaje, pogrešna šifra = poruka „Pogrešna šifra. Objava je otkazana." + prekid. **Napomena za build/dev:** default `agrix-release` promeniti u `modConfig` pre isporuke (izvor se objavljuje fleetu → šifra štiti od slučajnog klika, nije prava tajna).
- **„Očisti tabele od podataka" / „Obriši sve" — potvrda kucanjem:** destruktivno brisanje svih unosa iz ~23 tabele (`OcistiTabele`, dostupno iz Admin panela i sa lista „Pregled listova") je ranije tražilo samo Da/Ne. Sada traži da operater **ukuca „OBRIŠI"** (`InputBox` + helper `PotvrdaObrisi`, prihvata `OBRISI`/`OBRIŠI`, bez razlike u velikim/malim slovima i dijakritici); bilo šta drugo (prazno/Cancel/pogrešno) = prekid. Zaglavlja i sami ListObject-i ostaju kao i pre.
- **Bez promene podataka / bez novih zavisnosti:** samo `modConfig`, `modAdmin`, `modPregledListova`; izvori **ASCII-only**, nema novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.

---

## vba-v2.15.0 — 2026-07-04
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Fokus: **kontrola proseka neto kg po gajbici pri otkupu** (upozorenje/blokada, pragovi po kulturi, podesivi kroz Matične podatke) + **popravka pada „Kupci" taba** na schema-drift. Izvori ostaju **ASCII-only**, bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.

- **Otkup — upozorenje/blokada po proseku gajbice:** pri unosu otkupnog lista (`frmOtkup`), posle bruto→neto konverzije, računa se prosek **neto kg po gajbici** (`Kolicina / KolAmbalaze`) za svaku klasu. Prosek **iznad praga upozorenja** → upozorenje uz „Da/Ne" (može da nastavi); prosek **iznad praga blokade** → **tvrda blokada** celog otkupnog lista (poruka imenuje spornu klasu). Kod dvoklasnog otkupa, ako **bilo koja** klasa probije prag blokade, ceo list se blokira (atomično — ni ispravna klasa se ne snima). Pragovi su **po kulturi**; prazno = provera isključena za tu kulturu (opt-in, npr. malina `2.1` / `2.2`).
- **Pragovi po kulturi — podesivi kroz Matične podatke → Kulture:** dva nova polja u editoru kultura (**„Prag upozorenja"** i **„Prag blokade"**, kg po gajbici) i **dve nove kolone u listi** (sa zaglavljima); validacija (prag blokade ≥ prag upozorenja). Čitanje i upis su **po imenu kolone** (otporno na audit/drift kolone `tblKulture`). Labeli editora prošireni ~50% radi čitljivosti dužih natpisa.
- **Pragovi — sami se pojave posle update-a (bez ručnog koraka):** nove kolone `PragProsekUpoz` / `PragProsekBlok` na `tblKulture` dodaje **`EnsureRuntimeSchema`** na startu (`InitApp`, isti obrazac kao `EnsurePoruke`), pa nastaju automatski i posle self-update-a **koda** (self-update ne migrira podatke). Do tada je kontrola **fail-safe**: prazna/nepostojeća kolona = bez provere, ne obara unos.
- **Fix — „Kupci" tab više ne ruši otvaranje (schema-drift):** `frmStammdaten.LoadList` je tvrdo padao (`Err.Raise` „Nedostaju kolone u tblKupci") ako **bilo koja** od 12 kolona `tblKupci` fali — kolona **„Država"** se tražila sa dijakritikom (`ChrW(382)`) dok su ostale ASCII, pa je na instalaciji gde je kolona `Drzava` (bez dijakritike, ili neka druga nedostaje) ceo tab padao. Sada je učitavanje **tolerantno** (obavezan je samo PK `KupacID`; kolone koje fale ostaju prazne) + **ASCII fallback za „Drzava"** i pri prikazu i pri izmeni.
- **Bez promene podataka / bez novih zavisnosti:** `modConfig`, `modMain`, `modOtkup`, `modSetup`, `frmOtkup`, `frmStammdaten`; **`.frx` netaknut** (nova polja/kolone u editoru su postojeće kontrole — samo runtime `Visible`/caption); izvori **ASCII-only**, bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**. **Šema:** nove kolone na `tblKulture` nastaju automatski na startu (ili ručno `EnsureDoradeSchema`).

---

## vba-v2.16.0 — 2026-07-04
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Fokus: **storno palete gotovih proizvoda (prerade) direktno iz forme Palete** — dosad prerađena paleta nije mogla da se stornira kroz UI, pa je klik na „Storniraj" vraćao samo neprozirno „Storno nije uspeo (vidi log)". Izvori ostaju **ASCII-only**, bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.

- **Palete — novo dugme „Storniraj preradu":** prerađena paleta (gotov proizvod, kolona `Prer.=Da`) se namerno ne stornira direktno, ali dosad **nije postojao nijedan put kroz formu** da se stornira sama **prerada** — pa je „Storniraj" nad takvom paletom davao samo neprozirno „Storno nije uspeo (vidi log)". Sada desna kolona (lista „Prerađene palete") ima dugme **„Storniraj preradu"** koje nad izabranom preradom zove postojeći `modStorno.StornoPrerada_TX`: prerada se stornira, a njene **palete se vraćaju u lager** (`Prerađeno=""`) i mogu se ponovo obraditi ili stornirati. Posle storna se osvežavaju i lista paleta i lista prerada.
- **Palete — jasna poruka pri stornu prerađene palete:** klik na „Storniraj" nad paletom `Prer.=Da` sada javlja **šta da se uradi** („Ova paleta je prerađena u gotov proizvod. Prvo stornirajte preradu u desnoj listi…") umesto neprozirne poruke o neuspehu.
- **Bez promene podataka / bez novih zavisnosti:** samo `frmPalete` (reuse `modStorno.StornoPrerada_TX`, koji je i ranije bio dostupan samo kao Alt+F8 makro `StornoPrerada_Prompt`); **`.frx` netaknut** (dugme je runtime `Controls.Add` + `WithEvents`, isti obrazac kao postojeća lista „Prerađene palete"); izvori **ASCII-only**, bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.

---

## vba-v2.16.1 — 2026-07-04
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Fokus: **popravka zaglavlja firme na paletnom listu gotovih proizvoda (prerada)** — pri (re)štampi/PDF-u kroz dvoklik podaci o firmi (naziv, adresa, PIB/MB/žiro + logo) nisu se popunjavali. Izvori ostaju **ASCII-only**, bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.

- **Paletni list gotovih proizvoda — zaglavlje firme se sada popunjava:** PDF „PALETNI LIST GOTOVIH PROIZVODA" (dvoklik na prerađenu paletu u desnoj listi → `ExportPreradaPDF`) izlazio je **bez podataka o firmi** (naziv/adresa/PIB-MB-žiro, logo). Uzrok: zaglavlje se pisalo **samo pri izgradnji šablona** (`EnsurePreradaSablon`, koji se preskače dok se `LAYOUT_VER` ne promeni), pa je ostajalo zamrznuto/prazno od prve izgradnje `PreradaSablon` lista (npr. ako `SELLER_*` config tada nije bio popunjen). Sada se zaglavlje čita **uživo iz `SELLER_*` configa na svako punjenje** (`FillPreradaSablon`), isti obrazac kao prijemnica (`FillPrijemnicaSablon`) — logo se prvo skida (jer `DocDrawLogo` samo dodaje) da se ne gomila pri reprintu. Šablon se **ne ruši** (nazvani opsezi, naslov i težine ostaju), samo se osvežava zaglavlje; nema potrebe za ručnim brisanjem `PreradaSablon` lista — popuni se na sledeći izlaz.
- **Bez promene podataka / bez novih zavisnosti:** samo `modPaletniList` (`FillPreradaSablon`, reuse `modDocStyle.DocSellerHeader`); **`.frx` netaknut**; izvori **ASCII-only**, bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.

---

## vba-v2.16.2 — 2026-07-04
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Fokus: **specifikacija otkupnih blokova — nova kolona „Vrsta i sorta" između `Datum` i `Kolicina`**. Vrsta i sorta voća spajaju se u jednu ćeliju iz podataka koji **već postoje** na otkupu (`tblOtkup.VrstaVoca` + `SortaVoca`), bez schema izmene i bez migracije. Izvori ostaju **ASCII-only**, bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.

- **Specifikacija otkupa — nova kolona „Vrsta i sorta":** dnevna/periodična specifikacija (Od/Do → dugme) i ručna selekcija otpremnica („Biraj otpremnice" → „Štampaj specifikaciju") sada, **između kolona `Datum` i `Kolicina`**, prikazuju vrstu i sortu voća **spojene u jednu kolonu** (npr. „Jabuka Ajdared"; ako sorta nije uneta, prikazuje se samo vrsta). Podatak se čita direktno iz otkupnog reda (`VrstaVoca` + `SortaVoca`), koji se popunjava još pri snimanju otkupa — **nema promene podataka ni migracije**. Raspored ostaje A4 landscape (sada 13 kolona) sa auto-uklapanjem po širini; `UKUPNO` red i dalje sabira Količinu, Vrednost, PDV i Ukupnu vrednost u ispravnim kolonama.
- **Šablon se sam regeneriše:** `SpecifikacijaSablon` je podignut na `LAYOUT_VER 3`, pa se novi raspored kolona primeni **automatski na prvom sledećem izveštaju** — bez ručnog brisanja lista.
- **Bez promene podataka / bez novih zavisnosti:** samo `modOtkupBlok` (`RenderSpec`) i `modPrint` (`EnsureSpecifikacijaSablon`/`FillSpecifikacijaSablon`, reuse `modDocStyle.DocSellerHeader`/`DocTitleBlock`); **`.frx` netaknut**; izvori **ASCII-only**, bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.

---

## vba-v2.16.3 — 2026-07-04
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Fokus: **ujednačavanje prikaza decimala na tačno dve — u formama Palete, Izveštaji i Dokumenta**. Neki brojevi su se prikazivali sa promenljivim brojem decimala (npr. cena `12.345678`, neto palete `123.4500001`, ili `150.5` pored `150`); sada svi novčani/kg iznosi imaju **tačno dve decimale**. Gajbe/kapacitet ostaju **celi brojevi**. Bez promene podataka; izvori ostaju **ASCII-only**, bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.

- **Palete — grid, stavke i prerade na 2 decimale:** kolone **Neto** i **Bruto** u pregledu paleta, kao i **Neto** u listi stavki i u listi „Prerađene palete", prikazivale su se kao sirov broj s punom preciznošću (npr. `123.4500001`). Sada se formatiraju na tačno dve decimale (`modPaletniList` grid funkcije). **Gajbe** i **kapacitet** ostaju celi brojevi.
- **Izveštaji — jedinstven kg prikaz (2 decimale):** `FmtKolicina` (jedinstveni izvor istine za prikaz kg u izveštajima; koristi ga `frmIzvestaj` i deljeni panel „Detalji otkupa") sada **uvek** daje dve decimale umesto ranijih 0/1/2 (`#,##0` ili `#,##0.##`). Time su usklađene sve „Količina"/„Manjak kg"/„Prosek" kolone i vrednosti u izveštajima. Fiksni format je lokalno-bezbedan (bez „500," repa koji je pravio opcioni `.##`).
- **Dokumenta — labeli, storno i recovery panel na 2 decimale:** validacione i statusne oznake (ukupno kg, manjak, prosek gajbe, saldo kupca/avans, KPI trake), cene iz cenovnika (bilo do 6 decimala), storno pregled (Količina/Cena/Iznos) i „Osiročeni dokumenti" (kolona Kol) — sve na tačno dve decimale. **Iznosi reversa u gajbama** i dalje su celi brojevi.
- **Bez promene podataka / bez novih zavisnosti:** samo `frmIzvestaj`, `frmDokumenta`, `modHelpers` (`FmtKolicina`), `modPaletniList` i `modDokumenta` (storno/recovery, reuse postojećeg `StornoNumText`); **`.frx` netaknut** (menjani su samo format-stringovi u kodu); izvori **ASCII-only**, bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.

---

## vba-v2.16.4 — 2026-07-04
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. **Docs-only** — bez promene VBA koda (FSO pull sloj je već isporučen u v2.12.0, nijedan `.bas`/`.frm`/`.cls` nije diran). Fokus: **ispravno dokumentovan stvarni uzrok banka „75/76 na Drive-u" = nepokrenut `SetupNewPC` na klijentu.** Izvori ostaju **ASCII-only**, bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.

- **Stvarni uzrok „75/76 na Drive-u" — `SetupNewPC` nije pokrenut na klijentu:** `tblLocalConfig` (per-mašina putanje `BANKA_DRIVE_SOURCE_PATH`, Inbox/Processed/Error, `PDFTOTEXT_EXE_PATH`, `APP_SETUP_COMPLETED`) **putuje unutar distribuiranog `.xlsm`-a** sa build/dev mašine. Ako se `.xlsm` samo prekopira a `SetupNewPC` se ne pokrene, klijent **nasleđuje dev putanje** (a `APP_SETUP_COMPLETED=DA` je „upečen" pa se first-run kapija preskače) → pull gađa nepostojeći folder = greška 75/76. Fix: `Alt+F8 → SetupNewPC`.
- **Runbook dopunjen (`docs/production-runbook-banka-import-setup.md`):** Faza 4.5 dobila **kritičnu napomenu** „`SetupNewPC` obavezan po svakom klijentu"; troubleshooting red za 75/76 sada šalje **prvo na `SetupNewPC`** (pa Editor, pa build); Faza 2 „Zašto FSO" preformulisana — **FSO (v2.12.0) je defanziva, ne primarni lek** (u produkciji je 75/76 skoro uvek bilo zbog nasleđenih dev putanja, a ne zbog shortcut-a samog).
- **Ispravka ranije formulacije:** v2.12.0 nota je 75/76 pripisivala shortcut/legacy uzroku; taj shipovani zapis ostaje kao istorija, a ova verzija dodaje ispravku unapred (ne menja se retroaktivno već isporučena verzija).

---

## vba-v2.16.5 — 2026-07-04
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Fokus: **sistemski (trajni) format kolona koje su se posle reinstalacije/self-update-a vraćale na `General`** — pa je Excel dug broj prikazivao kao naučnu notaciju (`1.23E+11`) ili tarabe (`####`), a na dokumentima je vrednost bila pogrešna ili prazna. Bez promene podataka; izvori ostaju **ASCII-only**, bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.

- **BPG (broj gazdinstva) se više ne prikazuje kao `1.23E+11` ni tarabe:** kolona `BPGBroj` u `tblKooperanti` i BPG ćelija na **kartici kooperanta** dobijaju **Text** format (`@`), pa dug identifikator ostaje čitljiv. Uzrok je bio što posle reinstalacije/self-update-a kolona nasledi `General`, gde Excel dug broj prikazuje u naučnoj notaciji, a na kartici se string koji liči na broj auto-konvertovao u broj pri upisu.
- **Prerada — težine na tačno dve decimale (bez `E`/tarabe/blank):** kolone `TezinaPaleteKg`, `BrutoKg`, `AmbalazaKg` u `tblPrerada` i odgovarajuće ćelije na **paletnom listu gotovih proizvoda** (`FillPreradaSablon`) dobijaju format `0.00`. Vrednosti ostaju **broj** (računske, `NzD`) — samo je prikaz fiksiran.
- **Sistemski — preživljava reinstall i self-update:** format se nameće kroz `modSetup.EnsureRuntimeSchema`, koji se izvršava na **SVAKI start** (`modMain.InitApp`) — pa kad god se šema tabele iznova napravi (nova instalacija, self-update koda, dodavanje kolone), format se **automatski ponovo postavi**, bez ručnog `Alt+F8`. Reuse postojećeg helpera `SetColumnNumberFormat` (isti obrazac kao pragovi kulture / decimalna količina).
- **Bez promene podataka / bez novih zavisnosti:** samo `modSetup` (`EnsureRuntimeSchema`), `modPrint` (`EnsureKarticaSablon`) i `modPaletniList` (`FillPreradaSablon`); komplementarno sa v2.16.3 (tamo grid data-source `Format$`, ovde format ćelija/kolona); **`.frx` netaknut**; izvori **ASCII-only**, bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.

---

## vba-v2.16.6 — 2026-07-04
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Fokus: **paletni list gotovih proizvoda (prerada) — red „Vrsta voća" i novi izbor rasporeda lista.** Red „Vrsta voća" sada prikazuje isključivo tekst iz combo-a „Gotov proizvod" (bez „DZ"/vrste/sorte iz `tblPaleta`), a novo Podešavanje bira detaljan ili zbirni raspored. Bez promene podataka; izvori ostaju **ASCII-only**, bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.

- **„Vrsta voća" = samo gotov proizvod:** u redu „Vrsta voća" na paletnom listu gotovih proizvoda (`FillPreradaSablon`) više se **ne čita vrsta ni sorta iz `tblPaleta`**, niti se dodaje prefiks „DZ". Prikazuje se **isključivo tekst izabran u combo-u „Gotov proizvod"** (`tblPrerada.TipGotovogProizvoda`, iz šifarnika „Vrsta gotovog proizvoda"). Ranije (v2.5.0) je red bio „DZ + vrsta + sorta + tip gotovog proizvoda"; sada je čist tip. Težine/ambalaža i stavke nepromenjene.
- **Novo Podešavanje „Detaljni prikaz sledljivosti (paletni list got. proizvoda)" (grupa „Otkup / dokumenta", DA/NE, default DA):** bira raspored lista:
  - **DA (kao do sada):** puna tabela stavki `Rb | Kooperant | Neto kg | Ambalaža` (jedan red po otkupu) + desni sažetak težina/ambalaže.
  - **NE:** umesto detaljne tabele — **samo lista šifri kooperanata** (bez ponavljanja, u jednom redu razdvojene zarezom); sažetak **težina/ambalaže se centrira i uvećava** da popuni oslobođeni prostor.
- **Šablon se sam regeneriše po režimu:** `PreradaSablon` je podignut na `LAYOUT_VER 5`, a verzija u `H1` nosi i režim (`-D`/`-N`), pa promena toggle-a **automatski ponovo izgradi** list u odgovarajućem rasporedu — bez ručnog brisanja lista. Kompatibilno sa v2.16.5 (težine i dalje `0.00`).
- **Bez promene podataka / bez novih zavisnosti:** samo `modConfig` (novi flag `PRERADA_SLEDLJIVOST_DETALJ` + `IsPreradaSledljivostDetalj`), `modPodesavanja` (jedan `bool` red) i `modPaletniList` (`EnsurePreradaSablon`/`FillPreradaSablon` granaju layout; privatni `BuildPreradaSablon{Detalj,Zbirno}` i `FillPrerada{StavkeDetalj,SifreZbirno}`, reuse `ConfigFlag`/`CfgAdd`/`GetOtkupiZaPalete`); **`.frx` netaknut**; izvori **ASCII-only**, bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.

---

## vba-v2.16.7 — 2026-07-04
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Fokus: **auto-kreiran kooperant se odmah vidi u `cmbKooperant` (bez zatvaranja forme).** Kad je uključeno „Auto-kreiraj kooperanta iz unetog imena" (v2.5.0), unos novog imena pri snimanju otkupa napravi red u `tblKooperanti`, ali se novi kooperant nije pojavljivao u padajućoj listi dok se `frmOtkup` ne zatvori i ponovo otvori. Bez promene podataka; izvori ostaju **ASCII-only**, bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.

- **`cmbKooperant` se osvežava odmah posle auto-kreiranja:** posle uspešnog snimanja otkupa u kojem je `ResolveKooperantByName` napravio **novog** kooperanta, lista combo-a se ponovo puni (`FillKooperantCombo`) pa je novi kooperant odmah dostupan za sledeći unos — bez zatvaranja/otvaranja forme. Ranije se lista punila **samo** pri izboru otkupnog mesta (`cmbOtkupnoMesto_Change`), pa je auto-kreiran kooperant ostajao „nevidljiv" do reload-a forme.
- **Osvežava se samo kad je stvarno kreiran nov:** `ResolveKooperantByName` je dobio opcioni izlazni parametar `created` (`True` isključivo kad je napravljen novi red u `tblKooperanti`); običan izbor **postojećeg** kooperanta iz liste ne dira combo (nema nepotrebnog ponovnog čitanja `tblKooperanti` na svakom snimanju otkupa).
- **Bez promene podataka / bez novih zavisnosti:** samo `modKooperant` (`ResolveKooperantByName` + opcioni `created`) i `frmOtkup` (novi privatni helper `FillKooperantCombo` — jedno mesto za `KoopFilterByOM` punjenje, reuse postojećeg `FillComboKooperantiByStanica`; poziva ga i `cmbOtkupnoMesto_Change` i osvežavanje posle unosa u `btnUnos`); **`.frx` netaknut** (menjan samo kod u `.frm`); izvori **ASCII-only**, bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.

---

## vba-v2.16.8 — 2026-07-04
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Fokus: **broj zbirne otpremnice unet u prijemnicu mora da postoji u sistemu.** Do sada je prijemnica prihvatala bilo koji broj zbirne — i kad ta zbirna nije uneta — pa je nastajala viseća („orphan") referenca koja **tiho nestaje iz obračuna manjka** (izveštaj `ReportManjak` je vođen tabelom `tblZbirna`, pa prijemnica koja pokazuje na nepostojeću zbirnu nije nigde uračunata). Sada se broj zbirne proverava, a ponašanje bira operater u Podešavanjima. Bez promene podataka; izvori ostaju **ASCII-only**, bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.

- **Prijemnica — broj zbirne mora da postoji u `tblZbirna`:** pri snimanju prijemnice (`frmDokumenta`) broj zbirne se proverava naspram postojećih (ne-storniranih) zbirni. Ranije se proveravalo **samo da polje nije prazno**, pa se mogla snimiti prijemnica koja pokazuje na nepostojeću zbirnu (greška u kucanju ili unos van redosleda) — takva prijemnica ne ulazi u obračun manjka jer je izveštaj vođen listom zbirni. Provera je **storno-aware** (stornirana zbirna se tretira kao nepostojeća).
- **Novo Podešavanje „Prijemnica: kad zbirna nije u sistemu" (grupa „Otkup / dokumenta", BLOK/UPOZORENJE, default BLOK):** bira ponašanje kada uneta zbirna ne postoji:
  - **BLOK (default):** tvrda greška, snimanje prijemnice se prekida uz poruku „Zbirna '<broj>' ne postoji u sistemu".
  - **UPOZORENJE:** upozorenje sa potvrdom (Da/Ne) — operater može svesno da snimi prijemnicu sa visećom referencom (npr. kada roba stigne pre nego što je zbirna uneta u sistem).
- **Backend sigurnosna mreža:** provera postojanja je i u `ValidatePrijemnicaInput` (a ne samo u formi), pa hvata i ne-form pozivaoce — ali **samo u BLOK modu** (u UPOZORENJE modu forma traži potvrdu, pa backend ne sme tvrdo da padne). Malina / auto-hladnjača lanac je bezbedan jer se zbirna kreira **pre** prijemnice (guard prolazi).
- **Bez promene podataka / bez novih zavisnosti:** samo `modConfig` (novi ključ `PRIJEMNICA_ZBIRNA_PROVERA` + `PrijemnicaZbirnaBlokira`, reuse `ConfigFlag` obrasca), `modDokumenta` (novi `ZbirnaPostoji` storno-aware, reuse `ExcludeStornirano`, + guard u `ValidatePrijemnicaInput`), `frmDokumenta` (provera u `btnUnosPrij_Click`), `modPodesavanja` (jedan `list:BLOK;UPOZORENJE` red) i `modBusinessFlowProTests` (novi `Test_PrijemnicaMissingZbirnaDoesNotAppend` + fixture zbirne u dva postojeća testa koja su se oslanjala na orphan prijemnicu); **`.frx` netaknut** (menjan samo kod u `.frm`); izvori **ASCII-only**, bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.

---

## gas · bank-pdf-downloader — 2026-07-04
> **GAS-only** (Apps Script na bankovnom nalogu koji skida PDF izvode Gmail→Drive) — **ne menja `AgriX_OtkupApp.xlsm`**, ne prolazi kroz `tools/release.sh` ni `ImportAllVBA`. Deploy = zameniti `Code.gs` u Apps Script projektu na tom nalogu i re-run `setupDailyBankPdfImportTrigger`. Fokus: **češće preuzimanje izvoda tokom dana + manji lookback prozor.**

- **Raspored 1×/dan (07h) → 6×/dan (07/08/10/12/14/16):** nova konstanta `BANK_IMPORT_TRIGGER_HOURS`; `setupDailyBankPdfImportTrigger` sada instalira po jedan dnevni okidač za svaki sat (re-run briše stare pa postavlja nove, idempotentno). Izvod koji stigne posle jutarnjeg pokretanja se pokupi isti dan. Bezbedno: `LockService` serijalizuje pokretanja, a dedup po stabilnom imenu (msgID + provera postojanja u folderu) čini dodatne prolaze no-op-om.
- **`searchDays` default 30 → 7:** sa 6×/dan rasporedom nov izvod se uhvati za par sati, pa lookback prozor **nije za svežinu** nego je **outage buffer** (koliko dana GAS-nerada da se auto-nadoknadi pri oporavku). Manji prozor ujedno bounduje re-download churn (~4× manje duplikata u `Downloaded`) i čini svaki run jeftinijim. Ređi duži prekid: `runBankPdfImportBackfill` (eksplicitan opseg, ne zavisi od `searchDays`). Postojeći klijenti sa eksplicitnim `searchDays` u `BANK_IMPORT_CLIENTS_JSON` nisu dirnuti — treba im ručno spustiti vrednost na tom nalogu.
- **Dokumentovan re-download churn (README „Dedupe"):** pošto VBA `PullBankPdfsFromDriveProduction` posle povlačenja **premesti** PDF u `Downloaded` (van vidokruga GAS naloga), naredno pokretanje ga ne nađe u korenu i **ponovo ga skine** dok mu je mejl u prozoru. Bezopasno (staging-dedup `IsDuplicateBankaImport` ga odbije, ne knjiži se dvaput) i ne usporava VBA (`Downloaded` se ne enumeriše). Svesno prihvaćeno umesto Gmail labela; `searchDays` bounduje obim.

---

## vba-v2.17.0 — 2026-07-05
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Fokus: **jedinstvena provera integriteta podataka kroz ceo lanac dokumenata** (otkup → otpremnica → zbirna → prijemnica → palete → prerada) + par novih reconciliation provera. Sve **read-only** — ništa se ne menja u podacima; izvori ostaju **ASCII-only**, bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**. `.frx` netaknut (sve nove UI kontrole su runtime `Controls.Add`).

- **Integritet pregled (`modIntegritet`) — 22 provere u jednom prolazu:** konsolidovana revizija koja izlistava sve neusklađene zapise kroz lanac dokumenata:
  - **A (količina/kg):** A1 Σotpremnica vs Σzbirna po broju zbirne (reuse `ValidateZbirna`); A2 manjak/višak zbirna→prijemnica (višak > 5%, „ništa primljeno", manjak > 10%); A3/A4 palete-stavke vs prijemnica i paleta-header vs stavke; A5 prerada ulaz vs Σstavke, izlaz ≤ ulaz.
  - **B (veze lanca):** B1 verwaist otpremnice/prijemnice (reuse `GetVerwaisteDokumente`); B2 otkupi bez otpremnice (`GetUnlinkedOtkupi`); B3 izgubljeni blokovi — otkup → stornirana/nepostojeća otpremnica (`GetLostOtkupBlokovi`); B4 „viseći" broj zbirne (ne postoji u `tblZbirna`); B5/B5b prijemnica/otpremnica bez broja zbirne; B6 broj zbirne se poklapa samo do velikog/malog slova (advisory za normalizaciju); B7 zbirna sa 0 (ili prazan) `UkupnoKolicina`.
  - **C (palete):** C1 stavka bez žive prijemnice; C2 stavka bez zbirne; C3 paleta-header bez stavke; C4 stavka ka storniranoj prijemnici; C5 dupli `BrojPalete` u istoj godini.
  - **D (prerada):** D1 paleta `Preradjeno=Da` bez aktivne prerade; D2 prerada ka nevalidnoj (nesvežoj/storniranoj) paleti.
  - Poređenje `BrojZbirne` je **case-insensitive** (`s5` = `S5`) — usklađeno sa ostatkom app-a; B6 posebno izlistava case-mismatch zapise za čišćenje (jer case-senzitivni `ReportManjak`/`GetVerwaisteDokumente` ih tiho razdvajaju).
- **Dva prikaza, isti engine:**
  - **In-app pregled** — klik na crveni upozorenje-baner (`frmOtkupAPP`) otvara ListBox pregled preko content-zone (sidebar i header ostaju vidljivi), sa podnaslovima po bloku, poravnatim kolonama (monospace) i dugmetom „Zatvori". Runtime kontrole (`Controls.Add` + form-local `WithEvents`), stilizovano po `modTheme`; navigacija ga gasi.
  - **Sheet** — Admin panel → „Integritet provere (tabele)" ili `Alt+F8 → RunIntegritetProvere`; upisuje sheet `INTEGRITET_PROVERE` sa punim listama (filter/sort/print u Excelu). Audit je **strogi nadskup** postojećeg startup banera (`CheckVerwaisteDokumente`).
- **Nove health provere (`modProductionHealthCheck`):** `Check_KooperantOtkupReconciliation` (Σ `Kolicina×Cena` po kooperantu vs sirov `tblOtkup`; hvata prazan/orphan `KooperantID`, loš `StanicaID`/`Datum` koji „ispadaju" iz per-kooperant izveštaja) i `Check_FakturaIznosReconciliation` (denormalizovani `tblFakture.Iznos` vs živi `Σ prijemnica.Kolicina×Cena` — hvata drift kad se prijemnica izmeni posle fakturisanja).
- **„Lista kooperanata po iznosu" (otkupni blok) — UKUPNO footer:** overlay dobio red **UKUPNO (prikazano)** sa sumom vrednosti + kg, a kad postoje redovi sa praznim `KooperantID` i redove „Prazan KooperantID (van liste)" + „tblOtkup UKUPNO" — operater vidi da li se prikaz slaže sa tabelom (prikazano + prazan = tblOtkup).
- **Bez promene podataka / bez novih zavisnosti:** novi `modIntegritet`; dopune `modProductionHealthCheck` (2 provere), `modOtkupBlok` (footer), `frmOtkupAPP` (runtime overlay + klikabilan baner), `modAdmin` (Admin dugme); dokumentacija `docs/INTEGRITET_PROVERE.md`. `.frx` netaknut; ASCII-only.

---

## vba-v2.18.0 — 2026-07-06
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Fokus: **ispravka stornirane paletizovane prijemnice/otkupa bez ponovne paletizacije.** Kad se roba fizički nije pomerila sa palete, storno + ispravka menja **samo dokument**, a palete se prevežu na novu prijemnicu umesto da se ponovo paletizuju (što je pravilo fantom-palete i trošilo brojeve). Uz to: korekcija broja gajbica/kg direktno na paleti (uz izbor preliva), rukovanje duplim unosom, i tvrde zaštite protiv tihog kvarenja evidencije. **Bez ručnog migracionog koraka:** nova kolona `Istorija` na `tblPaleta` se dodaje **automatski pri startu** (`InitApp → EnsureRuntimeSchema`), pa preživljava i self-update koda (`Alt+F8 → EnsurePaletniListSchema` i dalje radi, ali nije obavezan). Izvori **ASCII-only**, bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**; `.frx` netaknut (nove UI kontrole su runtime `Controls.Add`).

- **Storno paletizovane prijemnice → 3 izbora (Dokumenti i Otkupni blokovi/autohladnjača):** posle storna prijemnice koja je nosila palete, operater bira:
  - **ISPRAVKA** — polja se sama popune iz stornirane (uklj. **datum**, kupac/vozač, vrsta/sorta, obe klase, količine i cena; broj se NE preuzima, predlaže se nov). Operater izmeni samo grešku i snimi; palete stare prijemnice se **prevežu** na novu — bez ponovne paletizacije, roba ostaje na svojim (zatvorenim) paletama.
  - **DUPLI UNOS** — roba nije stvarno primljena drugi put: fantomske paleta-stavke se skinu sa paleta, a paleta koja time ostane prazna se **stornira** (kroz kanonski `StornoPaleta`). Druge prijemnice na istoj paleti se ne diraju.
  - **NIŠTA** — palete ostaju osiroćene (rešava se kasnije ručno kroz recovery panel).
- **Ispravka identiteta (vrsta/sorta/tip ambalaže) — relabel u mestu:** kad ispravka menja identitet, paleta i njene stavke se **preoznače** (roba se ne pomera, menja se etiketa). **Zaštita:** ako paleta nosi robu **više prijemnica**, promena identiteta bi iskvarila i tuđu robu → operacija se **tvrdo blokira** (ne upozorenje), uz uputstvo za bezbedan izlaz (Skini stavke → nov unos). Blokada važi i kad je operater već potvrdio.
- **Korekcija broja gajbica / kg direktno na paleti:** kad ispravka menja količinu, evidencija se usklađuje **u mestu**: manjak se skida sa poslednje stavke (stavka na 0 se stornira, zatvorena paleta se reopen-uje); višak koji ne staje nudi izbor **PRELIJ** (na sledeću/novu paletu) ili **PREKO kapaciteta** (svesno slaganje preko maksimuma); neto i ambalaža-kg se proporcionalno preračunaju, a header palete iz stavki (self-healing, uračunava i su-stanare na rubnoj paleti). Prerađena paleta se ne dira (blok uz poruku). Dostupno i ručno: `Alt+F8 → PaletaAdjust_Prompt`.
- **Recovery panel (Osiroćeni dokumenti → Mod: Palete) — dorada:** ciljna lista se filtrira po **broju zbirne** izabrane stornirane prijemnice (zbirna = ključ sledljivosti, obično tačno 1 cilj), sa kolonom **„Ocena"** (Prevezi / +koriguj / +etiketa); dugme **„Cilj: svi"** za izlaz iz filtera (npr. ispravka snimljena pod novom zbirnom); dugme **„Skini stavke"** (detach za dupli unos); potvrda relabela pri prevezivanju.
- **Sledljivost — gap-fix:** re-point prijemnice na drugu zbirnu (Mod: Prijemnice) sada prevodi i **paletne stavke** na novu zbirnu — ranije su ostajale sa mrtvom zbirnom, pa se lanac paleta → zbirna → kooperanti prekidao (hvatao ga je Integritet C2).
- **Storno zbirne — upozorenje:** kad se stornira zbirna koja ima aktivnu prijemnicu, operater dobija upozorenje da je sledljivost prekinuta + uputstvo da prijemnicu preveže na novu zbirnu (paletne stavke tada automatski dobijaju novu zbirnu).
- **Audit trag na paleti:** nova kolona **`Istorija`** na `tblPaleta` beleži svaku izmenu (RELABEL / DETACH / ADJUST sa **vidljivom deltom gajbica** `+3` / `-2` / `0 = samo kg` / STORNO_PRAZNA) uz Monitor event. Neuspelo auto-prevezivanje se beleži kao trajni `PALETA_RELINK_FAIL` (WARN) — stanje ostaje vidljivo i posle klika na poruku, a osiroćene palete se i dalje vide u recovery panelu i hvata ih Integritet C4.
- **Rizik za podatke:** sve izmene pišu samo u `tblPaleta` / `tblPaletaStavka` (+ `BrojZbirne` na stavkama), **transakciono** (rollback na svaku grešku); podrazumevano ponašanje (bez ispravke) je nepromenjeno. Naplata/faktura/ambalaža idu nepromenjenim postojećim putevima.
- **Regresioni test (`modTestPalete`, `Alt+F8 → RunPaleteTestSuite`):** 11 integracionih testova / ~75 provera nad **pravim** tabelama; ceo run je u jednoj transakciji koja se **uvek poništava** (ne ostavlja podatke; test-identitet `TST-*`). Pokriva paletizaciju + spill i idempotency guard, CLEAN re-point + KG-sync, relabel gejt, per-klasa verdikt na istom agregatu, korekciju (potreba-izbora → PRELIJ, PREKO, smanjenje kroz više paleta), detach su-stanar + storno prazne palete, blokadu na prerađenoj, zbirna re-point, i co-tenant RELABEL blokadu.
- **Dodirnuti moduli:** `modPaletniList` (engine: `EvaluatePaletaReassign`, `ReassignPaleteToPrijemnica_TX` + relabel/co-tenant guard, `DetachOsirocenePaletaStavke_TX`, `AdjustPaletaGajbiceZaPrijemnicu_TX`, `SpillGajbice` konsolidovan, `PaletaLog`/`LogRelinkFailure`), `modPaletniListUI` (korekcija prompt), `modDokumenta` (zbirna re-point + `GetAktivnePrijemnice` uklj. nepaletizovane), `modAutoHladnjaca` + `frmOtkup` + `modOtkupBlok` (autohladnjača ispravka + prefill), `frmDokumenta` (storno tok + prefill + recovery panel), `modConfig` + `modSetup` (`Istorija` šema), `modTestPalete` (nov). `.frx` netaknut; ASCII-only.

---

## vba-v2.19.0 — 2026-07-06
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Fokus: **centralni, poslovno-svestan storno/ispravka framework za Otpremnicu, Zbirnu i Revers** — storno više nije „Da/Ne" nego prvo pita **šta storno poslovno znači**, uz tvrdu garanciju invarijante **zbirna = zbir svojih aktivnih otpremnica** (KG po klasi + ambalaža) i trajni trag staro→novo. Nadograđuje paletni sloj iz v2.18.0 (samo ga poziva, ne dira). Nova tabela `tblStornoVeze` se dodaje **automatski pri startu** (`InitApp → EnsureRuntimeSchema`, preživljava self-update). Izvori **ASCII-only**, `.frx` netaknut (UI kontrole su runtime `Controls.Add`).

- **Pametan okidač (smart trigger):** za obično storno (dokument bez nizvodnog toka — nema prijemnice/paleta; malina 1:1) → jedan `Stornirati X?` i gotovo, kao pre. Pun poslovni dijalog sa 4 moda se pokaže **samo** kad postoji zavisni tok koji traži odluku operatera. Motor (rekalkulacija/odvezivanje) radi tiho u oba slučaja.
- **Četiri poslovna moda storna** (`modStornoFlow`): **ISPRAVKA_ODMAH** (pogrešan unos, isti fizički događaj → storno stare, nova, prevezivanje + rekalkulacija), **DUPLI_FANTOM** (nikad nije trebalo da postoji → skini/odveži posledice bez naslednika; saldo se ne duplira), **PONIŠTENJE_BEZ_ZAMENE** (fizički tok se poništava; blokada uz svesnu potvrdu ako postoje zavisni dokumenti), **REŠI_KASNIJE** (persistent recovery zapis, ne samo MsgBox).
- **Invarijant engine** (`modDokumentInvariant`): `ValidateZbirnaInvariant` (KG **po klasi** = tvrdo, ambalaža ukupno = tvrdo, ambalaža po klasi = meko uz jasno označen limit postojećeg unosa), `RecalculateZbirnaFromOtpremnice_TX` (rewrite zbirne iz otpremnica — otpremnice su izvor istine), `ValidateOtpremnicaZbirnaImpact` (stara + nova zbirna kod prevezivanja).
- **Persistent correction context** (`tblStornoVeze` + `modStornoContext`): svaka storno/ispravka dobija red staro→novo sa modom, statusom (`PENDING/COMPLETED/FAILED/MANUAL_REQUIRED/CANCELLED`) i `NeedsRecovery`. Kontekst **preživljava** zatvaranje forme/Excela; svaki prelaz ide i u **Monitor** (`Monitor_Event`).
- **ISPRAVKA otpremnice — dovršetak AUTOMATSKI po snimanju nove** (nema dugmeta): otkupni listovi se prevežu na novu otpremnicu, **prijemnica i paletne stavke se presele na novu zbirnu** (`ReassignPrijemnicaToZbirna_TX`), nova zbirna se rekalkuliše, a **stara zbirna se STORNIRA** kad ostane prazna (ranije je pogrešno ostajala aktivna 0/0 sa zaglavljenom prijemnicom/paletama). Ako downstream relink ne uspe → kontekst `MANUAL_REQUIRED` (ne `COMPLETED`).
- **Bezbednost auto-dovršetka:** pre nego što se pending ISPRAVKA veže za upravo snimljeni dokument, sistem pita „da li je ovo zamena?" — sprečava pogrešan relink ako je operater napustio ispravku pa uneo drugi dokument.
- **Storno zbirne / reversa bez tihog mismatch-a:** storno zbirne odvezuje otpremnice u „čeka zbirnu"; storno reversa se oslanja na postojeći saldo koji **već isključuje stornirano** (bez kontra-stavke, bez duplog salda).
- **Revers ispravka kroz UI + hardening:** revers dobija kratak izbor (obično storniranje / ispravka / odustani) umesto slepog simple-storno puta; ispravka reversa se dovršava **automatski po snimanju novog reversa**. `CompleteReversIspravka` sada **validira** da novi revers stvarno postoji kao aktivan (inače `MANUAL_REQUIRED`, ne `COMPLETED`). Auto-dovršetak ispravke ima **safe-stop**: ako postoji više otvorenih ispravki istog tipa dokumenta, ne bira se naslepo najnovija — sistem staje i upućuje na recovery (sprečava pogrešan relink).
- **DUPLI vs PONIŠTENJE — sada stvarno funkcionalno različiti:**
  - **DUPLI/FANTOM = razveži.** Stornira **samo** ciljni dokument i **odveže** sve povezano — deca **prežive** nevezana, spremna za reveze: otkupni blokovi se **oslobode** (`OtpremnicaID`/`BrojZbirne` prazno, blok ostaje aktivan — realna kupovina), otpremnice zbirne se vrate u „čeka zbirnu", prijemnica/palete ostanu **osiroćene** (recovery zabeleška, ne blokira). Tiho, bez pitanja.
  - **PONIŠTENJE = kaskadni storno.** Stornira ciljni dokument **i ceo nizvodni tok iz spiska** (zbirna → sve otpremnice → prijemnica → paletne stavke), kroz kanonska storno jezgra (prijemnica ide kroz `StornoPrijemnica` → faktura se **osiroti**). **Paletne stavke idu kroz paletni motor** (`DetachOsirocenePaletaStavke_TX`, isti put kao recovery „Skini stavke"): gajbe/neto/ambalaža se **skidaju sa palete**, paleta se **reopen-uje** ispod kapaciteta, a paleta koja ostane **prazna se STORNIRA** — su-stanari (druge prijemnice/zbirne na istoj paleti) **netaknuti**. (Raniji naivni flag-flip stavke je ostavljao paletu „punom" — ispravljeno; motor se samo poziva, ne dira se.) Otkupni blokovi se i ovde **samo oslobode** (nikad ne storniraju). Uvek prvo pun spisak posledica + svesna potvrda.
  - **Ownership pravilo:** dokument kaskadira/razvezuje **samo ono što ekskluzivno poseduje**. Otpremnica poseduje zbirnu **samo ako je jedina** (malina 1:1 ili poslednja) → tek tada PONIŠTENJE jedne otpremnice obara ceo tok; **deljena** zbirna (više otpremnica) se ne obara (sestre) → samo rekalkulacija (tada je PONIŠTENJE = DUPLI na nivou otpremnice).
  - **Hladnjača vs eksterni kupac (normalni mod):** nizvodni tok (prijemnica/palete) pripada zbirni i kaskadira se **samo za hladnjača-kupca** (`kupac == MALINA_DEFAULT_KUPAC`, ista detekcija kao auto-predlog broja prijemnice). Za **eksternog** kupca je zbirna poslednji **interni** dokument, a prijemnica **prvi i poslednji eksterni** → framework je **ne dira** (ide svojim faktura-mehanizmom).
- **Ispravka „nuliranja umesto storna" (bug iz produkcije):** kad zbirna ostane bez ijedne aktivne otpremnice, sada se **STORNIRA** (ne ostaje aktivna 0/0 sa zaglavljenom prijemnicom/paletama) — kroz zajednički `RecalcOrStornoEmptyZbirna_TX` u svim storno modovima.
- **Regresioni test (`modTestStorno`, `Alt+F8 → RunStornoTestSuite`):** 22 scenarija nad **pravim** tabelama u jednoj transakciji koja se **uvek poništava** (test-identitet `SVT-*`; config `MALINA_DEFAULT_KUPAC` se snapshotuje pa vraća rollback-om): rekalkulacija zbirne posle storna otpremnice, validacija obe zbirne pri prevezivanju, storno zbirne bez mismatch-a, ispravka zbirne prevezuje otpremnice+prijemnicu, paletne stavke dobijaju novu zbirnu, revers ispravka ne duplira saldo, revers poništenje uklanja saldo, pending ostaje vidljiv na fail, smart-trigger gate, T11 = pun scenario ispravke otpremnice (prevezuje prijemnicu/palete, stara zbirna stornirana, context COMPLETED samo na uspeh), T12–T13 = revers complete traži aktivan novi revers (odbija → `MANUAL_REQUIRED` / COMPLETED + saldo samo novi), T14 = safe-stop pri više otvorenih ispravki, T15 = PONIŠTENJE uvek traži svesnu potvrdu (spisak posledica) i bez zavisnih, **T16 = DUPLI otpremnice oslobodi blok (ostaje aktivan) + ne dira prijemnicu + prazna zbirna se STORNIRA (nulling fix), T17 = PONIŠTENJE zbirne (hladnjača) kaskadno stornira otpremnice+prijemnicu + skida paletne stavke kroz motor (deljena paleta reopen + gajbe umanjene + co-tenant druge zbirne netaknut; prazna paleta stornirana), T18 = PONIŠTENJE zbirne (eksterni kupac) NE dira prijemnicu, T19 = DUPLI odvezuje vs PONIŠTENJE stornira otpremnice, T20 = PONIŠTENJE jedne otpremnice deljene zbirne ne obara zbirnu (sestra + prijemnica prežive), T21 = malina 1:1 (blok=okidač, stanica=hladnjača → nikad dve otpremnice/dva bloka) — PONIŠTENJE jedine otpremnice kaskadira ceo tok (zbirna+prijemnica+palete), blok oslobođen ali aktivan, T22 = SIMPLE storno otpremnice (bez nizvodnog toka) NE ostavlja aktivnu zbirnu 0/0 nego je STORNIRA**.
- **Pre-merge doslednost (review fixevi):** (a) **simple** storno otpremnice sada koristi isti `RecalcOrStornoEmptyZbirna_TX` kao DUPLI/PONIŠTENJE → prazna zbirna se stornira (ranije je simple putanja mogla ostaviti aktivnu **0/0**); (b) **DUPLI zbirne je atomaran** — storno + odvezivanje otpremnica u **jednoj** transakciji (deljeni `StornoZbirnaIDetach_TX`, isti kao simple) + guard: ako su postojale otpremnice a nijedna nije odvezana → `MANUAL_REQUIRED` (ne lažni COMPLETED); (c) **DUPLI/PONIŠTENJE otpremnice** — guard: ako su postojali otkupni blokovi a nijedan nije oslobođen → `MANUAL_REQUIRED` (blok ne ostaje tiho na storniranoj otpremnici); (d) **revers preview** više ne duplira količinu dvojnog upisa (uzima jednu nogu + prikazuje broj knjižnih redova).
- **Rizik za podatke:** aditivno (4 nova modula + tabela); jedine izmene postojećeg su storno-rutiranje + save-hook u `frmDokumenta`, konstante u `modConfig`, šema u `modSetup`. Sve mutacije transakciono (rollback na grešku); paletni engine samo pozvan (netaknut, njegovi testovi i dalje prolaze).
- **Dodirnuti moduli:** `modDokumentInvariant` (nov), `modStornoContext` (nov), `modStornoFlow` (nov, + DUPLI-razveži / PONIŠTENJE-kaskada primitive + hladnjača/eksterni gate + nulling fix), `modTestStorno` (nov, 22 testova), `modAutoHladnjaca` (`IsHladnjacaKupac` — detekcija hladnjača-kupca), `frmDokumenta` (storno UX + auto-dovršetak), `modConfig` + `modSetup` (`tblStornoVeze` šema). Reuse: `modStorno` (storno jezgra `StornoZbirna`/`StornoOtpremnica`/`StornoPrijemnica`), `modDokumenta` (`ReassignOtkupToOtpremnica_TX`/`ReassignPrijemnicaToZbirna_TX`/`ZbirnaPostoji`), `modAmbalaza` (saldo), `modPaletniList` (`DetachOsirocenePaletaStavke_TX` — skidanje paletnih stavki + storno praznih paleta u PONIŠTENJE kaskadi, isti put kao recovery panel; motor se ne dira/širi). `.frx` netaknut; ASCII-only; bez novih `Poruka()` ključeva.

---

## vba-v2.20.0 — 2026-07-06
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Fokus: **Integritet audit C4 („Paleta-stavke ka storniranoj prijemnici") prikazuje ceo dokumentni tok u istoj tabeli** — do sada je nalaz imao samo `StavkaID / PaletaID / PrijemnicaID / BrojPrijemnice`, pa je operater morao ručno da traži čemu ta stornirana prijemnica pripada. Sada su, u istim redovima, i **BrojZbirne, količina i ambalaža zbirne, otkupno mesto i proizvođač(i)**. Read-only provera (ništa se ne upisuje), bez izmene šeme; nadovezuje se na re-point/sledljivost iz v2.18.0/v2.19.0.

- **C4 dobija 5 novih kolona (dokumentni tok problematične paleta-stavke):** `BrojZbirne` · `ZbirnaKg` · `ZbirnaAmb` · `OtkupnoMesto` · `Proizvodjac`. Put: paleta-stavka → prijemnica (na koju pokazuje) → zbirna → otkupno mesto + kooperanti.
  - **BrojZbirne** — sa prijemnice na koju stavka pokazuje; čita se preko **svih** redova `tblPrijemnica` (uklj. stornirane — C4 baš cilja stornirane prijemnice, pa se broj zbirne ne bi video da se gledaju samo aktivne).
  - **ZbirnaKg / ZbirnaAmb** — Σ `UkupnoKolicina` / `UkupnoAmbalaze` aktivnih redova te zbirne (reuse `AggByBroj`; dvoklasna zbirna Kl I+II se sabira, isto kao A1/A2).
  - **OtkupnoMesto** — preko **aktivnih otpremnica** te zbirne (`StanicaID → tblStanice.Naziv`); fallback za mirror-stanica zbirne (malina mod, „S" prefiks) direktno iz `tblZbirna.VozacID` → pokriva i potpuno stornirane lance bez žive otpremnice.
  - **Proizvodjac** — kooperanti **aktivnih otkupa** te zbirne (primarno `tblOtkup.BrojZbirne`; fallback `OtpremnicaID → tblOtpremnica.BrojZbirne` kad kolona `BrojZbirne` na otkupu ne postoji — schema drift). Imena kao „Ime Prezime" iz `tblKooperanti`; nepoznat ID ostaje ID (bez tihe rupe).
- **Više vrednosti = „; " separator:** kad jednu zbirnu vuče više otkupnih mesta ili više proizvođača, `OtkupnoMesto` i `Proizvodjac` prikazuju distinct listu odvojenu sa `; ` (npr. `Petar Petrović; Marko Marković`). Sheet autofit proširen `A:H → A:I` (9 kolona); in-app monospace pregled se sam proširi. C1 nalaz ostaje 4 kolone (nema prijemnice → nema toka).
- **Rizik za podatke:** nema — provera je čisto read-only (piše samo u `INTEGRITET_PROVERE` sheet / in-app ListBox); ne dodiruje poslovne tabele. Ostali blokovi (A/B/C/D) nepromenjeni.
- **Dodirnuti moduli:** `modIntegritet` (C4 blok + novi privatni helperi `PrijemnicaZbirnaMap` / `OtkupnoMestoByZbirna` / `ProizvodjacByZbirna` / `AddDistinctJoin`), `modDokumenta` (`BuildIdNameDict` Private→**Public** radi reuse za nazive stanica/kooperanata — bez duplog helpera). `.frx` netaknut; ASCII-only; bez novih `Poruka()` ključeva.

---

## vba-v2.21.0 — 2026-07-06
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Fokus: **nova banka ALTA u uvozu izvoda** — Alta izvodi (račun `190-`) se sada prepoznaju i parsiraju kroz postojeću multi-bank arhitekturu. Jedan nov parser-modul + jedna grana u dispečeru; deljeni 4-nivo integritet, 17-kolonski staging, mapiranje (`tblNovac`) i forma `frmBankaImport` ostaju **netaknuti** (reuse > new). Bez izmene šeme, bez migracije podataka.

- **Nov parser `modBankaAlta`** (5 ugovornih funkcija `ExtractIzvodBroj/Datum/Racun/Saldo Alta` + `ParseBankaIzvodAlta`, isti ugovor kao ProCredit/Halk): naslov „IZVOD BR.", 2 datuma (knjiženja/prijema), STANJE data-linija (4 iznosa + 2 broja), **smer transakcije po „Obr. naknada"** (zaduženje = standalone iznos pre; odobrenje = `<iznos> <šifra>` posle), referenca = 15-cifreni „Podaci za reklamaciju"; parsiranje ograničeno na „PROMENE"…„Ukupno za ra". Broj-format `1,234.56` (isti kao ostale banke). Uključen `Test_AltaParse` (Alt+F8).
- **`DetectBank` otisak = naslov „IZVOD BR." + prefiks računa `190-`** — razlikuje Altu od Komercijalna/Halk („Izvod broj", 9. znak `.` vs `o`) i ProCredit („IZVOD NNN"). Ne skreće pogrešno kad se `220-`/`155-` (ProCredit/Halk) pojave kao partner unutar Alta izvoda (traži i njihov specifičan header, kog Alta nema).
- **Verifikovano nad stvarnim izvodom** (BR. 110, 17 naloga): `DetectBank=ALTA`, sva 4 nivoa integriteta prolaze (Početno 11.346,21 + uplate 2.504.625,00 − isplate 1.654.811,11 = Novo 861.160,10; broj naloga 14/2 se slaže do na paru sa bančinim totalima); poreklo naloga (naziv banke, npr. „PROCREDIT BANKA") ispravno izbačeno iz partnera i svrhe.
- **Rizik za podatke:** nema — bez izmene šeme (reuse postojećeg `tblBankaImport`, 17-kolonski staging od v6.18+), parser + dispečer su aditivni. Runtime zaštita: pri drugačijem layout-u integritet **glasno rollback-uje** batch (nema tihe korupcije). Self-update sam kreira nov modul (čist `.bas`, bez formi/`WithEvents`/`MSForms` deklaracija → faza 1, ne pogađa zamke #3/#4).
- **Dodirnuti moduli:** `modBankaAlta` (nov), `modBankaImport` (`DetectBank` grana + `Case "ALTA"` u dispečeru). Docs: `development-banka-parser.md` (referentna tabela) + `CLAUDE.md` (mapa parsera). `.frx` netaknut; ASCII-only; bez novih `Poruka()` ključeva.

---

## vba-v2.22.0 — 2026-07-15
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Fokus: **objedinjeni Storno centar** — sav storno ide kroz jedno dugme „Storno", sa jasnim uvidom u efekat pre potvrde i sledljivošću ispravki (append-only smer: storno starog + novi red koji nosi vezu na stari). Bez izmene poslovnih tokova; PWA/sync netaknut (ključ ostaje `BrojZbirne`).

- **Jedan ulaz za storno:** uklonjena dva direktna dugmeta („Pregled storniranih", „Osiroćeni dokumenti") sa glavne forme — sve ide kroz „Storno" → browse (bez kucanja broja). Panel „Osiroćeni / za doradu" i „Stornirani" žive unutar tog ekrana.
- **Uvid pre potvrde (panel „Efekat storna"):** poravnati naslovi kolona (zaseban header-listbox), opis svakog moda iznad svog dugmeta, i po-modu upozorenja. Stil efekta ujednačen: uvek prvo „Duplikat" pa „Poništenje" (a kad je efekat isti, kaže se jednom). Zaglavlje prikazuje **naziv** partnera/stanice (ne ID) + indikator „[ispravka dokumenta X]" / „[zamenjen dokumentom Y]".
- **Sledljivost (append-only):** nove trace-kolone (`IspravkaOd`, `ZamenjenSa`, `CorrectionID`, `IzdatoStatus`) + denorm poslovni ključ `BrojOtpremnice` na otkupnim blokovima. Ispravka utiskuje vezu stari↔novi red; štampa otpremnice/prijemnice nosi indikator ispravke u naslovu. Šema se sama dopunjava (`EnsureSledljivostSchema`), backfill `BackfillOtkupBrojOtpremnice` (Alt+F8). Vidi ADR-0001/0002.
- **Prefill forme pri ispravci:** klik na „Ispravka" za otpremnicu i zbirnu (ranije samo prijemnica) sada popuni novu formu podacima storniranog dokumenta — operater ne mora da ima papirni original ispred sebe.
- **Guard C (nepromenljivost izdatih):** blok-storno nad **živom** otpremnicom se odbija sa jasnim razlogom (spreči preračun već izdatog/prosleđenog dokumenta bez svesne odluke).
- **Ispravke iz internog review-a:** (1) `Reši kasnije` / `Ispravka` više **ne** storniraju čekirane blokove — blok-storno je moguć samo uz Duplikat/Poništenje (čekirani se inače ispišu kao ignorisani); (2) sažetak obuhvata je neutralan pre izbora moda (upućuje na kolonu „Efekat storna"); (3) prijemnica-correction **ne prijavljuje „gotovo"** ako palete nisu stvarno skinute — nepotpun detach ili „ne diraj palete" → recovery (Osiroćeni dokumenti), inače završeno; (4) **objedinjeni „Nedovršeno / recovery" centar** — jedan panel prikazuje i persistentne context-e (RESI_KASNIJE/MANUAL, sa `CorrectionID`) i pojedinačne osiroćene stavke (deduplikovano), sa akcijom po redu (2×klik: context → otvori storno / odbaci; osiroćeno → „Osiroćeni dokumenti" panel). Ranije su parkirani context-i bili nevidljivi u recovery toku.
- **Performanse:** idle pre-warm keš browse-a (`Application.OnTime`, ~60s debounce) — otvaranje instant umesto 7–10 s.
- **Regres-testovi:** `modTestStornoCentar` (`Test_StornoCentar_All`, Alt+F8) — rollback-safe auto-testovi (trace utiskivanje, Guard C, DocIsIssued, mrtvi roditelj bloka, impact agregat, aktivni dokumenti, blok-storno) + `Test_FindSingleActiveRow`.
- **Rizik za podatke:** nizak — dopuna šeme je aditivna (nove kolone), postojeći tokovi netaknuti. Odloženo za sledeći PR: append-only re-verzionisanje zbirne (korak 3.2), „Vrati storno" (undo, review #5) dorada — trenutno konzervativno i verifikuje se kroz `Test_UndoStorno`. ASCII-only izvori; `.frx` netaknut.

---

## vba-v2.23.0 — 2026-07-15
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Fokus: **kaljenje storno recovery-ja** — sledljivost izmene izdate zbirne (audit-trag) + zaštita „Vrati storno" (undo). Nastavak na v2.22.0; bez izmene poslovnih tokova, PWA/sync netaknut.

- **Zbirna — audit-trag za in-place preračun izdate zbirne:** kad se **izdata** zbirna promeni automatskim preračunom (npr. storno jedne otpremnice iz zbirne sa više otpremnica), upisuje se durabilan trag u Monitoring (`ZBIRNA_IZDATA_RECALC`, WARN): `CorrectionID` + razlog + **stara→nova** vrednost (kg/amb) po klasi. Zbirna je izveden agregat (suma aktivnih otpremnica) pa se **ne re-verzionise** (nov broj bi razbio interne veze lanca), ali izmena izdatog dokumenta više nije tiha (ADR-0001). Eksplicitna **ISPRAVKA** zbirne i dalje ide punim append-only putem (nov `BrojZbirne` + relink + trace).
  - Odluka doneta na osnovu analize sync-a: `ZbirnaID` je master-interni (sync ključa na `ClientRecordID`), pa je nov `ZbirnaID` sync-u nevidljiv — ali **isti `BrojZbirne`** na dva živa reda razbija interne lookup-e. Detalji: ADR-0002 sekcija 3 (revidirana).
- **Undo („Vrati storno") — zaštita i sakrivanje iz produkcije:** motor `UndoStorno_TX` je konzervativan (ne vraća `tblNovac` vezu, ne journališe konkretan row-set — review #5), pa je **dugme sakriveno iz produkcije** (`UNDO_UI_ENABLED=False`); undo ostaje dostupan kroz `Test_UndoStorno` (Alt+F8) dok pun storno-journal ne stigne. Reverse (revers) grana dobija **dup-guard** (`ActiveAmbalazaDokExists`): odbija undo ako već postoji aktivan revers istog broj+tip (ranije bi duplirao).
- **Regres-testovi:** dodati u `Test_StornoCentar_All` — `Test_ZbirnaRecalcInPlace_Auto` (recalk ostaje in-place, bez novog reda), `Test_UndoReverseGuard_Auto` (reverse dup-guard), `Test_GetNedovrseno_Auto` (recovery dedup/CorrectionID).
- **Rizik za podatke:** nizak — nema izmene šeme; audit je dodatni upis u Monitoring, recalk ponašanje **nepromenjeno** (i dalje in-place). Odloženo: pun storno-journal (StornoOperationID + novac veza) + ponovno uključivanje undo dugmeta. ASCII-only izvori; `.frx` netaknut.

---

## vba-v2.24.0 — 2026-07-17
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Fokus: **storno operacioni žurnal → lossless undo motor (Otkup + Revers).** Rešava review #5 na nivou motora: undo sada vraća i `tblNovac.OtkupID` (koji je storno nepovratno brisao) i cilja baš tu operaciju. Produkcijsko dugme „Vrati storno" ostaje SAKRIVENO do operation-centric UI-a (sledeći PR); motor se verifikuje kroz test-suite.

- **Novi append-only ledger `tblStornoZurnal`** (`OperationID | Timestamp | DocType | Broj | Tabela | RowID | Kolona | StaraVrednost | NovaVrednost`): svaka storno operacija zabeleži **staru i novu** vrednost svake dirnute ćelije PRE mutacije. Šema se sama dopunjava (`EnsureStornoZurnalSchemaCore`, u `EnsureRuntimeSchema`).
- **Instrumentacija (ambient op-kontekst, `modStornoZurnal`):** `StornoOtkup` i `StornoOMKoopByBrDok` (revers) otvore operaciju **po broju** (dvoklasni dokument → **jedan** `OperationID`); primitive (`StornoAmbalazaByDokument`, `ResetNovacOtkupLink`) usput žurnališu. Žurnal upisi teku **unutar iste storno TX** → rollback storna povlači i žurnal red **u Excel tabeli** (napomena: eksterni CSV crash-log `modJournaling` nije transakcion, pa CSV linija ostane — ne dira poslovne podatke, ali je van rollback-a). Ne-instrumentirane putanje (otpremnica/zbirna/prijemnica/faktura) su no-op.
- **`UndoOperation_TX(opID)` — pravi inverz sa optimistic-concurrency zaštitom:** cilja **samo tu operaciju**, i **pre svake mutacije** proverava (sve-ili-ništa): podržana tabela + kolona, TAČNO jedan ciljni red, i **drift** (trenutna vrednost ćelije == `NovaVrednost`; ako je stanje promenjeno posle storna — npr. novac re-linkovan — undo se **odbija** i ne gazi noviju izmenu). Vraća `tblNovac.OtkupID` i rešava reused-broj rizik.
- **Fail-closed sigurnost:** `BeginStornoOp` (bez `OperationID` → prekid storna; nested drugi dokument → greška), `JournalCell` (neuspeo upis → rollback), obavezan PK (`RequireColumnIndex`), i undo garde (`OtkupBlockDeadParent`, `ActiveAmbalazaDokExists`) — sve fail-closed (provera koja ne uspe **blokira** undo, ne propušta ga).
- **Precizne garde:** otkup active-dup je **po (broj, klasa) reda** (parcijalni storno jedne klase se može vratiti iako je druga aktivna); revers dup-garda (#134) sada važi i na žurnal-putu.
- **„Vrati storno" dugme za sada OSTAJE SAKRIVENO** (`UNDO_UI_ENABLED=False`): motor je lossless, ali je Stornirani panel document-centric (nema `OperationID` po redu), pa bi kod reused poslovnog broja `LatestOpFor` mogao vratiti pogrešnu generaciju. Dugme se uključuje tek uz **operation-centric UI** (lista undoable operacija → `UndoOperation_TX(opID)` direktno) — sledeći mali PR. Do tada undo se verifikuje kroz `Test_StornoCentar_All` / `Test_UndoStorno`.
- **Regres-testovi** (`Test_StornoCentar_All`): journal undo (novac vraćen), dvoklasni (jedan op), revers-guard, pre-validacija, **drift**, **parcijalna klasa**, pomešan-op, prazan-BrDok (odvojeni op).
- **„Lossless" — precizno:** važi za storna **posle** ovog builda, pod uslovom da stanje ćelija nije promenjeno posle storna (drift-guard odbija inače). Stara storna nisu lossless-undo-abilna (legacy). **Odloženo:** chain instrumentacija; `UndoneAt` status operacije. Rizik: nizak/srednji (aditivna tabela; dira `modStorno` motor — verifikovano suite-om). ASCII-only; `.frx` netaknut.

---

## vba-v2.25.0 — 2026-07-18
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Fokus: **operation-centric „Vrati storno" UI** — dugme se uključuje (`UNDO_UI_ENABLED=True`) i undo ide po **konkretnom `OperationID`**, ne po poslovnom broju. Nastavak na v2.24.0 (žurnal motor).

- **Operation-centric panel „Vrati storno":** dugme u pregledu storniranih otvara **listu undoable operacija** (`OperationID | Datum | Tip | Broj | #redova | Status`) iz `GetUndoableStornoOperations`; klik/„Vrati" → `UndoOperation_TX(izabrani opID)` **direktno**. Time reused poslovni broj više ne može da vrati pogrešnu generaciju (razne generacije = različiti `OperationID`) — glavni P1 iz review-a. Status kolona: `moguce` / `vec vraceno` / `izmenjeno`.
- **Dead-parent guard po redu operacije** (`OtkupBlockDeadParentByID`): mrtav roditelj **druge, nepovezane generacije** istog broja više **ne preblokira** undo bezbedne operacije (ranije broj-level `OtkupBlockDeadParent` je gledao sve stornirane redove istog broja).
- **`OtkupReissueDupExists` fail-closed** (`RequireColumnIndex` + raise na grešku).
- **Prazan `BrDok` (unbound blok)** se sada može vratiti kroz UI (undo po `OperationID`, ne treba poslovni broj).
- **Testovi (+3):** reused-broj (undo starog op vraća staru generaciju), dead-parent druge generacije (undo prolazi), prazan-BrDok end-to-end undo.
- **Rizik za podatke:** nizak — UI je runtime overlay (`.frx` netaknut); motor nepromenjen osim per-red garde (uža, sigurnija). ASCII-only. **Odloženo:** chain instrumentacija; `UndoneAt` status operacije.

---

## vba-v2.26.0 — 2026-07-19
> Tačan broj/datum se potvrđuje pri `tools/release.sh`. Male, realne UX/robusnost ispravke (uvid + status čitanje).

- **Uvid pre storna — `Kolicina` je sada SUMA aktivnih Klasa I+II** (`modStornoImpact.ImpactHeader`): ranije je čitala samo prvi (jedan) red pa je **potcenjivala** dvoklasni dokument (otpremnica/zbirna/prijemnica). Sada sabira količine aktivnih redova istog broja (storniran red se ne broji). Test: `Test_ImpactHeaderSum_Auto`.
- **`DocIsIssued` čita status sa AKTIVNOG reda** (`modDokumentInvariant`): pređeno sa `LookupValue` na `LookupActiveID` za `IzdatoStatus` — ranije je za broj sa storniranom generacijom mogao pokupiti status STORNIRANOG reda; sada gleda aktivni red (konzervativno „izdato" ako aktivnog nema). Test: prošireni `Test_DocIsIssued_Auto` (broj sa storniran=DRAFT + aktivan=IZDATO → `True`).
- **Regres-testovi** (`Test_StornoCentar_All`, Alt+F8): **88 OK, 0 FAIL** (uz prethodni `EnsureRuntimeSchema` da trace kolone postoje). Dodati `Test_ImpactHeaderSum_Auto` + prošireni `Test_DocIsIssued_Auto`.
- **Rizik:** nizak — read-only agregacija/čitanje statusa u uvidu (ne dira storno/undo motor). ASCII-only; `.frx` netaknut.
## vba-v2.27.0 — 2026-07-20
> Verzija/datum se **finalizuju pri `tools/release.sh`** (i `APP_VERSION` u `modConfig` — trenutno 2.21.0 — mora se dići na 2.27.0 da self-update komparator vidi novo). Fokus: **self-update otporan na crash i ATOMSKI** (ne snima delimičnu verziju). Istraga crash-a posle v2.16.1: mehanizam se od pre 2.16.1 **nije menjao** — polomile su ga **nove `WithEvents`/event-sink deklaracije dodavane u FORME** kroz release-e v2.17.0+ (utvrđeni krivac; ista klasa kvara kao zamka #3), u sadejstvu sa nezaštićenim rupama u updateru. Zatvoreno sa obe strane: sadržaj (WithEvents seli iz formi u `clsUiSink`) + transport (hardening `modSelfUpdate`). **Domet:** cilj nije apsolutno „ne može da sruši", nego da **nijedan neuspeh ne ostavi snimljenu polu-novu / neoperativnu aplikaciju** (disk ostaje netaknut dok update nije 100% uspešan). `.frm` migracija je pokrivena statički, ali **code-merge put nije još odigran na klijentu** — obavezan smoke-test pre flote (dole). Detalji i zamke #7–#15: `docs/SELF_UPDATE.md`.

- **KRIVAC UKLONJEN — `WithEvents` seli iz formi u `clsUiSink` (nova generička sink klasa):** svih 25 post-2.16.1 event-sink deklaracija (`frmDokumenta`: storno centar, finder, undo-ops, „nedovršeno", recovery — 24 kontrole; `frmOtkupAPP`: integritet overlay close — 1) postaje **obična referenca bez `WithEvents`**, a događaje (Click/Change/DblClick) hvata `clsUiSink` instanca po kontroli (`WireSink` posle `Controls.Add` + jedan Public `UiSinkEvent` dispatcher po formi → postojeći handleri, netaknuta tela). Deklaracioni blok formi je time vraćen na **2.16.1-kompatibilan oblik** (samo inertni dodaci), pa code-merge formi pri sledećem update-u više ne dodaje nijedan event-sink u `.frm`. Isti obrazac kao `clsBlokUI`/`clsWheelList`/`clsAdminBtn` (WithEvents u klasi, ne u formi); **novo pravilo** u `CLAUDE.md` + `SELF_UPDATE.md` zamka #11: novi `Private WithEvents` u formama je zabranjen. Pre-2.16.1 form-WithEvents su zamrznuti (klijenti ih već imaju). `clsUiSink` stiže kao nov `.cls` **hard putem (faza 2 `Import`)** — **nije još proveren na klijentu**, vidi smoke-test. **Ovo je nova runtime event-arhitektura, ne samo transport** (dve produkcijske forme + nova klasa).
- **Forme/sheet komponente NIKAD ne idu u fazu 2 (`VBComponents.Remove`):** do sada je SVAKA komponenta čiji code-merge padne u sva 3 prolaza išla u `Remove` — a uklanjanje FORME u runtime-u je poznata zamka #1 (korupcija + „Document Recovery" = **crash Excela**), pri čemu faza 2 uvozi samo `.bas`/`.cls` pa bi forma i **trajno nestala** iz projekta. Sada: u `failed` ulaze samo `.bas`/`.cls` (+ dodatni type-guard na samom `Remove`); forma čiji merge padne dobija **best-effort rollback na stari kod** i jasnu poruku „potreban reinstall" — update se završi bez crash-a.
- **Delta-skip:** komponenta čiji je kod identičan novom telu se **ne dira** (ranije se na svaki update prepisivao ceo projekat ~90 komponenti). Manje COM edita = manji rizik + brži update; dvofazni `Remove`+`Import` se sada dešava samo kad je neki „tvrd" modul (module-level `MSForms` deklaracije) stvarno izmenjen.
- **ATOMARNOST — snima se SAMO pri punom uspehu:** bilo koji fatalni ishod (forma se ne može azurirati, faza-2 `Import` padne, `Save` ne uspe, **download nepotpun**) → **ništa se ne snima**. Disk ostaje **stara ISPRAVNA verzija**; poruka traži „zatvori bez snimanja". `APP_VERSION` se diže tek pri snimanju → **nema tihe polu-nove verzije** (koja bi prestala da nudi update). `Save` je **verifikovan** (`Err` + `ThisWorkbook.Saved`) — neuspeo save = neuspeo update.
- **Kompletnost download-a:** `DownloadReleaseFiles` poredi **preuzeto vs očekivano**; `n <> expected` → prekid (ranije je i 1/95 fajlova prolazilo — gledalo se samo `n=0`).
- **Tvrdi moduli se PREPOZNAJU UNAPRED** (`IsHardModuleBody`: module-level `WithEvents`/`As MSForms.`, uz strip komentara) i idu **pravo u fazu 2** — nad njima se `AddFromString` (koji diskonektuje `CodeModule`) **nikad ne poziva**. Pokriva i `clsUiSink`. Zamena za „error-driven" rutiranje; 11 modula (7 `WithEvents` klasa + `modOtkupBlok`/`modKarticaDetalji`/`modPodesavanja`/`modMouseWheel`).
- **Faza 2 uvozi TAČNU listu** (imena fajlova iz faze 1), ne skenira ceo temp → nema `SKIP_MODULES` bypass-a ni uvoza sirovo-palih/dev modula.
- **Otkaz SVIH `OnTime` tikova pre importa:** `StopScheduledSync` + `StopAutoSaveTimer` + `StopHeartbeatTimer` + **`StopStornoWarm`** (dodat — stigao u `main` posle prvobitnog rada). Tik između faza forsira compile polomljenog projekta, a AutoSave/StornoWarm bi i **snimio polu-verziju**.
- **Startup watchdog** `RecoverPendingSelfUpdate` (iz `StartApp`): ako faza 2 nikad ne opali (prekinuta sesija), na sledećem startu čisti stale stanje + temp i obavesti (disk je stara verzija; ne „dovršava" fazu 2 nad starim projektom — mešalo bi verzije).
- **`clsUiSink` lifecycle:** `Release` + `Class_Terminate` + `ReleaseUiSinks` (QueryClose/Terminate obe forme) raskidaju krug forma↔sink i otpuštaju reference kontrola pre nego što `PrepareRuntime` edituje VBProject. **`WireSink` fail-visible** (log umesto tihog gutanja).
- **`EnableEvents`/`ScreenUpdating` se vraćaju na svakom izlazu** (tokom prozora faza 1→2 namerno off, vraća ih faza 2). Ranije su ostajali ugašeni → posle „zatvori i otvori" u istoj instanci `Workbook_Open` ne opali.
- **Sitne ispravke:** `Err` iz lookup-a ne „boji" Add put; prazno/nečitljivo telo ne prazni postojeći modul; jedinstven temp folder po operaciji (ne kolidira sa drugim update-om).
- **Faza 2 FAIL-CLOSED:** stara komponenta koja je **još prisutna** (Remove nedovršen) više se ne preskače tiho (bio je mešan build: stari modul + nov ostatak) → fatalno; `imported = expected` obavezno; posle `Import` verifikuje se ime + tip komponente (`ImportedOk`). **Ojačano (review):** `expected = 0` u fazi 2 (izgubljen registry state posle faze-1 Remove-ova) → fatalno (ranije bi `0<>0=False` pa bi se snimio build bez uklonjenih modula); prazan `dir` u fazi 2 → fatalno auto-close (ne tihi izlaz sa polovnim projektom).
- **Review hardening (dodatno):** (1) **No-op guard fail-SAFE** — `AnyUpdatePending` na bilo koju grešku vraća „ima izmena" (pun atomski put), nikad lažni „već ažurni"; (2) **delta-skip case-precizniji** — `SameCode` lowercase-uje samo kod **izvan** string-literala/komentara pa poredi **binarno** (apsorbuje VBE identifier re-casing ali **hvata** izmenu case-a u stringu, npr. `"DA"`→`"da"`); (3) **forma sa module-level `WithEvents`/`MSForms` → reinstall** (fail-closed) umesto opasnog `AddFromString` (i za novo tvrdo telo i za zatečenu tvrdu/zamrznutu formu; zamka #1/#3).
- **Review hardening (2. krug):** (1) **phase-2 integritet handoff-a** — faza 1 čuva `expected`=broj tvrdih modula + `fhash`=`Sha256Hex(files)`; faza 2 zahteva `pending="1"`, `expected=savedExpected` i `Sha256Hex(filesCsv)=fhash` (hvata parcijalno izgubljen/pokvaren registry state, ne samo prazan); (2) **tri-state manifest** — `ParseManifestFiles`→`hadFilesKey`, `ResolveReleaseSource`→`isVersioned`: prazna kolekcija je legacy listing **samo** za flat kanal bez `"files"` ključa, versioned ili prisutan-ali-prazan `files[]` = **INVALID** (prekid, bez listing fallback-a); (3) **versioned `manifest_sha256` obavezan + 64-hex** (`IsSha256Hex`) — prazan/malformiran `current.json` = fatalno (SHA-nedostupan **na klijentu** i dalje degradira, PIN presedan); (4) **immutable republish HARD odbijen** — `PublishReleaseToDrive` odbija re-objavu verzije koja već ima `manifest.json` (override `ALLOW_REPUBLISH=False`, `True` samo za TEST/hitno); `current.json` se piše samo uz validan 64-hex `manifest_sha256`.
- **WithEvents → `clsUiSink` (URAĐENO, follow-up grana):** preostale zamrznute `WithEvents` u formama (dodatne u `frmDokumenta`, `frmPalete`, `frmIzvestaj`, `frmAgrohemija`, `frmBankaExportPregled`; uz 25 ranije migriranih u `frmDokumenta`/`frmOtkupAPP`) izmeštene su u `clsUiSink` → **nijedna forma više nema module-level `WithEvents`**, čime je **uklonjena klasa hard-crash-a** koja je obarala update 2.16.1→2.21.0. **VAŽNO — forme i dalje NISU self-updatable (ostaju reinstall-only):** posle migracije zadržavaju module-level `As MSForms.*` reference (runtime kontrole), koje `IsHardModuleBody`/form-guard **namerno** i dalje hvata → forma ide na **reinstall** (empirijski potvrđeno `RunSelfUpdateDev`-om nad `frmAgrohemija` = „Preskočeno, reinstall"). Migracija je uklonila **crash**, ne i reinstall-only status; forme se distribuiraju bootstrap-om (`ImportAllVBA`/nov `.xlsm`). Ne blokira ovaj release jer se flota **ionako bootstrap-uje** (vidi „rollout").
- **Prazan stub modul više ne obara update + mrtva klasa obrisana:** prazan `.bas/.cls` (samo header) je `ExtractModuleCode`-om davao prazno telo uz `Err=0` → faza 1 je dizala „prazno telo" (`[-2147218703]`) i **forsirala fazu 2** na SVAKOM update-u (krivac: prazan orphan `clsSEFValidationResult.cls`). Sada: prazno telo uz `Err=0` = **`„same"`/skip** (no-op; prazan izvor nikad ne briše zatečen kod); genuina greška ekstrakcije (`Err<>0`) i dalje → faza 2/reinstall. Mrtva `clsSEFValidationResult.cls` (0 referenci u projektu) **obrisana**. `clsUiSink.Bind` sada **`Err.Raise`** na nepodržan tip kontrole (fail-fast umesto tihog no-op-a).
- **Auto-close na fatalni ishod:** atomičnost više nije proceduralna („zatvorite bez snimanja") nego **tehnička** — `AbortSelfUpdateClose` (`Saved=True` + `Close SaveChanges:=False`) sam zatvori svesku bez snimanja, pa ni `Ctrl+S`/OneDrive AutoSave ne upiše polu-nov projekat. **Ojačano (review):** `AbortSelfUpdate` postavi `Saved=True` **odmah** i zove close **direktno iz tekućeg stack-a** (ne preko novog `Application.OnTime` makroa — polomljen/nekompajlabilan projekat mu ne bi razrešio name-dispatch pa se close nikad ne bi desio).
- **ROLLOUT (kritično):** `modSelfUpdate` je u `SKIP_MODULES` → **star klijent se self-update-uje SVOJIM starim updater-om** (bez ijedne ove ispravke: ne gasi tajmere → `StornoWarmTick` „Cannot run macro"; `AddFromString` nad tvrdim modulima → crash). Zato star→nov **mora** ići jednokratnim ručnim **bootstrap-om** (`ImportAllVBA`/nov `.xlsm`), NE self-update-om. Ove ispravke štite **buduće** self-update-ove (sa novog updater-a), ne star→nov hop. Redosled: 1) bootstrap sve klijente; 2) potvrdi da čitaju `current.json`; 3) tek onda normalna objava.
- **Multi-copy izolacija:** oba `Application.OnTime` poziva su **workbook-qualified** (`'Ime.xlsm'!Proc`; `modMain`/`modAdmin`/`modSelfUpdate`), a `phase2` registarsko stanje **scope-ovano po workbook imenu** — dve otvorene kopije (npr. DEV test) ne gaze jedna drugu. **Redosled:** self-update ide **pre** min-version enforce gate-a u `StartApp` (inače `enforce=YES` ugasi baš klijenta kome update treba).
- **Release = kompletan snapshot (`modRelease`/`modDrive`):** `PublishReleaseToDrive` objavljuje `version.json` **tek pošto SVI code fajlovi stignu** (jedan pao → manifest se ne dira); **prune** zastarelih fajlova iz `AgriX_Release` (`DriveTrashFile` — inače ih klijent ponovo skida; npr. obrisani test moduli); manifest nosi `files` (ime+veličina+**sha256**). Uz klijentski „preuzeto = očekivano", release je atomski na oba kraja.
- **SHA-256 verifikacija sadržaja (F0+F1 iz snapshot plana):** manifest nosi `sha256` svakog fajla (`modDrive.Sha256File` — reuse dokazanog `.NET SHA256Managed` iz PIN hasha); klijent je **manifest-driven** i **verifikuje heš** svakog skinutog fajla pre importa — tiha korupcija/stale (HTTP 200 ali pogrešni bajtovi) → fatalno (`AbortSelfUpdate`). Fallback: SHA nedostupan **na klijentu** → prisustvo/broj (kao pre); stari publisher (bez `files[]`) → legacy. **Ojačano (review):** hashed manifest na SHA-sposobnoj mašini zahteva **validan 64-hex `sha256` po fajlu** (`IsSha256Hex`) — prazan/nevalidan heš više nije „prošao bez provere" nego fatalno. Self-test `Alt+F8 → Test_Sha256File`.
- **Versioned folderi + `current.json` (F2) — IMPLEMENTIRANO** (`docs/SELF_UPDATE_SNAPSHOT_PLAN.md`): `releases/<verzija>/` snapshot + `manifest.json`; **`current.json`** pokazivač (app_version + release_folder_id + **`manifest_sha256`**) upisan **poslednji** (atomski „go live"). Klijent: `current.json` → versioned folder → provera `manifest_sha256` **pre ijednog fajla** → per-file SHA (fail-closed na svakom nivou); **flat fallback** ako nema `current.json` (pun backward-compat). Dual-write flat (`version.json`) za stare klijente. Retention 10 (`PruneOldReleases`); `RollbackReleaseTo`/`ListReleases` (Alt+F8). Re-objava iste verzije **upozori** (prepisuje snapshot).
- **DEV test** `RunSelfUpdateDev` (Alt+F8): najlakši lokalni test — code-merge iz lokalnog git klona kroz **isti atomski core**. **Guard = zaštita od slučajnog klika** (git klon `src-vba` + `.git`), **ne bezbednosna granica** (ko može `Alt+F8` može i `Alt+F11`/`ImportAllVBA`).
- **OPERATIVNO — kako fix stiže na klijente:** `modSelfUpdate` je u `SKIP_MODULES`, pa **ne stiže self-update-om**. Jednokratno po mašini: `git pull` + `Alt+F8 → ImportAllVBA → Compile → snimi` (ili nov `.xlsm`). Tek posle toga self-update opet važi. Crash-ovane mašine: `Backup\AgriX_pre-update_*.xlsm` pa isto.
- **Rizik za podatke:** nema — izmene u transportu koda + UI event-dostavi + build-objavi; poslovne tabele/šema se ne diraju. `.frx` netaknut; ASCII-only; bez novih `Poruka()` ključeva. **Rizik za sam update (kod):** nije nula dok se ne odigra smoke-test (dole) — zato obavezan pre flote.
- **Dodirnuti moduli:** `modSelfUpdate` (rewrite jezgra + atomarnost/auto-close/faza-2 fail-closed + manifest-driven SHA verify + **F2 klijent: `ResolveReleaseSource`/`DownloadNamedText` + review hardening: no-op fail-safe, `SameCode` string-case, forma-tvrda→reinstall, `IsSha256Hex`, direktan close**), `clsUiSink` (nov + lifecycle), `frmDokumenta` + `frmOtkupAPP` (WithEvents→sink + cleanup), `modMain` (watchdog + redosled), `modAdmin` (qualified OnTime), `modDrive` (`DriveTrashFile` + `Sha256File`/`Test_Sha256File` + **`DriveEnsureFolder`**), `modRelease` (atomska objava + prune + manifest sa sha256 + **F2: versioned dual-write + `current.json` + `PruneOldReleases`/`RollbackReleaseTo`/`ListReleases` + republish upozorenje**). Docs: `SELF_UPDATE.md` (zamke #7–#18 + F2 lanac), `SELF_UPDATE_SNAPSHOT_PLAN.md`, `SELF_UPDATE_SMOKE.md`, `CLAUDE.md`.

---

## vba-v2.28.0 — 2026-07-21
> Verzija/datum se **finalizuju pri `tools/release.sh`** (uz `APP_VERSION` u `modConfig`). Nadovezuje se na (još neobjavljeni) 2.27.0.

- **Kartica kooperanta — rekapitulacija robe (kg):** u PDF kartici kooperanta, ispod finansijske tabele, novi blok **„REKAPITULACIJA ROBE (kg)"** — zbir otkupljene kilaže grupisan po **vrsti, sorti i klasi** voća za izabrani period, sa UKUPNO redom. Isti obuhvat kao sama kartica (bez storniranih, isti datumski opseg, isti kooperant). Roba bez sorte prikazuje se korektno (prazna Sorta ćelija).
- **Nedirano (namerno):** ekranska lista Kartice, Google izvoz (`ExportKarticeToGoogle_Core`) i PWA kartica ostaju nepromenjeni — UKUPNO red glavne kartice i dalje nosi tačan string „UKUPNO", pa PWA filter radi. PWA prikaz rekapitulacije odložen u backlog (`P3-PWA-1`).
- **Rizik za podatke:** nema — samo novi read-only prikaz u PDF izveštaju; poslovne tabele/šema i `.frx` se ne diraju; ASCII-only izvor; bez novih `Poruka()` ključeva.
- **Dodirnuti moduli:** `modIzvestaj` (nova `ReportKarticaRobaRekap` + `PrintKarticaPDF` prosleđuje rekap u šablon), `modPrint` (`FillKarticaSablon` dobio Optional `rekapData` + nova `FillKarticaRobaRekap` renderuje blok). Docs/backlog: `backlog/backlokg.md` (`P3-PWA-1`).

---

## vba-v2.28.1 — 2026-07-21
> Verzija/datum se **finalizuju pri `tools/release.sh`** (uz `APP_VERSION` u `modConfig`). Patch: **KPI „kg danas" više ne puca** na pokvarenom/čudnom zapisu u `tblOtkup`.

- **Ispravka — KPI „Otvoreno kg" / „Današnji otkup kg" (`SumOtkupKgToday`):** ulazak u `frmDokumenta` je znao da izbaci `SumOtkupKgToday | 13 | Type mismatch` **čak i kada za taj dan nema otkupa** — jer se skenira **svaki istorijski red** `tblOtkup`, a ne samo današnji. Datum se poredio inline (`IsError`/`IsDate`/`CDate`); `IsError` hvata `#N/A`/`#REF!` ćelije, ali je ostajao uzak prolaz (zapis gde `IsDate=True` a `CDate` ipak pukne, ili tip van očekivanja) koji je probijao do `EH` → **ceo dnevni zbir je padao na 0** uz log greške.
- **Rešenje:** svaki red se sada čita **isključivo kroz deljene bezbedne parsere** — `NzToText` (Variant/Error → `""`), `TryParseDateValue` i `TryParseDouble` (oba sa sopstvenim `On Error`, nikad ne bacaju). **Nijedan pojedinačan zapis** (Excel greška, prazno, tekst, neočekivan tip) ne može više da obori KPI. Isti obrazac koji `frmOtkupAPP.SumOtkupKgForDate` već koristi (`SafeDateKey`/`SafeKpiDouble`) — dve KPI funkcije su usklađene. Dodat i `IsArray(data)` guard (skalar/degenerisan `DataBodyRange` → čist 0 umesto `UBound` greške) u `SumOtkupKgToday`, `SumOtkupKgForDate` i `CountDocsForDate`.
- **Napomena o podacima:** loš zapis ostaje u `tblOtkup` — sada se samo **preskače** (ne ulazi u zbir); ako treba da se uračuna, ispraviti tu ćeliju (Excel: `Go To Special → Formulas → Errors`).
- **Rizik za podatke:** nema — samo robusnije čitanje u KPI izračunu; poslovne tabele/šema i `.frx` se ne diraju; ASCII-only izvor; bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.
- **Dodirnuti moduli:** `frmDokumenta` (`SumOtkupKgToday` → bezbedni parseri + `IsArray` guard), `frmOtkupAPP` (`SumOtkupKgForDate` + `CountDocsForDate` → `IsArray` guard).

---

## vba-v2.28.2 — 2026-07-21
> Verzija/datum se **finalizuju pri `tools/release.sh`** (uz `APP_VERSION` u `modConfig`). **RF-01 (M0) — „brisanje balasta" (Wave 0 iz plana sanacije):** čisto uklanjanje mrtvog koda, bez ijedne promene ponašanja. Svaki cilj re-verifikovan protiv aktuelnog `main`-a (v2.28.1) pre brisanja.

- **Obrisani mrtvi moduli (0 živih poziva):**
  - `modBankaImportParserClipboard` — legacy parser za ručno kopiran tekst izvoda; javni `ParseBankaIzvod`/`TestParser` bez ijednog produkcionog caller-a (glavni import ide preko `ParseBankaIzvodForImport` + bank-specific parsera).
  - `modLicenceTests` — **nekanonski spelling-duplikat** (britansko „Licence") kanonskog `modLicenseTests` (američko „License", **zadržan**). Time je **uklonjena realna „Ambiguous name" pretnja**: oba modula su izlagala identične `Public TestLicense_All/_SplitParts/_PartsMatch/_NonEmptyParts/_DeviceFingerprint`. Impl `modLicense` i svi runbook-ovi ionako pokazuju na `modLicenseTests`.
- **Obrisani mrtvi članovi (0 upotrebe):**
  - `modBankaImport.GetFileNameFromPath2` — bajt-identična kopija postojećeg `GetFileNameFromPath`; jedini poziv (u dev testu `Test_SaldoIntegrityOnSamplePDF`) preusmeren na `GetFileNameFromPath`.
  - `modArrayUtils.GroupBySum` / `SumColumn` — generički array helperi bez ijednog poziva (aspiraciona „zamena za Zbirni-Reports" koja se nikad nije zakačila).
  - `modIzvestaj` enum `IzvestajTip` — tip + svih 7 članova bez upotrebe (dispatch izveštaja ide preko tabova forme, ne preko enum-a).
  - `frmIzvestaj` — mrtve `UpdateUnosButtonState` i `PrijemniceZaOtpremnicu` (nijedan `clsUiSink`/`UiSinkEvent` dispatcher ih ne zove; veza prijemnica↔otpremnica preko `BrojZbirne` je živa u `modAutoHladnjaca`/`modBrojevi`).
- **VAŽNO — operativno (build):** `ImportAllVBA` **ne briše komponente**. Pre importa u master `.xlsm` **ručno ukloniti u VBE** dve komponente: `modBankaImportParserClipboard` i `modLicenceTests` (britansko „Licence" — ostaviti `modLicenseTests`). Zatim `ImportAllVBA → Debug→Compile` (mora bez „Ambiguous name") → `AssertBlankBuild` → snimi. Flota izmenu dobija kroz redovan bootstrap/self-update snapshot.
- **Rizik za podatke:** nema — isključivo uklanjanje mrtvog koda; poslovne tabele/šema i `.frx` se ne diraju; ASCII-only izvor; bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**. Statičke provere: 0 referenci na sva obrisana imena; `modE2EReleaseGate` `Application.Run` i dalje jednoznačan; Sub/Function/End balans; jedina preostala dupla `Public` linija je pre-postojeći `#If VBA7` `MouseWheel_*` par (benigna conditional-compilation, van obima).
- **Dodirnuti moduli:** obrisani `modBankaImportParserClipboard`, `modLicenceTests`; izmenjeni `modBankaImport` (uklonjen `GetFileNameFromPath2` + preusmeren jedini poziv), `modArrayUtils` (uklonjeni `GroupBySum`/`SumColumn`), `modIzvestaj` (uklonjen enum `IzvestajTip`), `frmIzvestaj.frm` (uklonjene 2 mrtve procedure; `.frx` netaknut). Prateći plan/status: `docs/REFAKTOR_PLAYBOOK.md` (RF-01) — ažurirati po merge-u #141.

---

## vba-v2.28.3 — 2026-07-22
> Verzija/datum se **finalizuju pri `tools/release.sh`** (uz `APP_VERSION` u `modConfig`). Fokus: **`modMigracija` (jednokratna migracija podataka iz starog `.xlsm` u novi prazan) — bezbednosno ojačanje da „success" ne sakrije izgubljene podatke i da destruktivan alat ne upiše polu-migriran fajl.** Alat je dev/admin (Matični podaci → Admin → „Podaci"; „Pregled listova" → „Migracije"); poslovni moduli i runtime ponašanje se ne diraju.

- **Glasne provere integriteta (`PROBLEMI: N` + upozoravajuća ikonica):** migracija ide kroz tabele NOVOG fajla i povlači istoimene iz starog — što je ostavljalo tihe rupe. Sada se prijavljuje: (a) tabela koja postoji **samo u starom** (nema je u novom → redovi bi tiho ostali); (b) **stara kolona sa podatkom bez cilja** u novom (preimenovana/izbačena); (c) **nova vezna kolona bez izvora** (`*ID` / `BrojZbirne` / `Klasa` / `DokumentTip` … → red bi stigao razvezan; audit i kalkulisane kolone se preskaču da brojač ne šumi); (d) **zbir čisto numeričkih kolona** (količine/vrednosti) staro vs novo pročitano nazad = da li je upis legao (ne hvata pogrešno mapiranje).
- **Fail-closed provere:** provera koja **nije izvedena** se prijavljuje (nije isto što i „prošla") — i za `Ensure*` korake na početku i za svaku jedinicu provere po tabeli.
- **Obavezan backup pre izmena:** `\Backup\AgriX_pre-migracija_*.xlsm` (isti obrazac kao pre-update backup). Ako fajl nije snimljen na disk ili backup ne uspe → **migracija se PREKIDA** (bez potvrđenog backup-a nema mutacije).
- **Nema tihog gubitka rezultata:** uklonjen `ThisWorkbook.Saved = True` (koji je zapravo **potiskivao close-prompt** → zatvaranje bez pitanja tiho gubi migraciju). Sada fajl ostaje „prljav" (Excel normalno pita „Snimi?" = svesna kapija); `modJournaling` AutoSave tajmer i **Excelov cloud AutoSave** (`AutoSaveOn`, OneDrive/SharePoint, zaseban mehanizam) se ugase za sesiju i vraćaju **samo na čistom uspehu** (na problem/grešku ostaju ugašeni da auto-save ne persistuje pre svesne odluke).
- **Format kolona (datumi/iznosi):** posle array-upisa preuzima se `NumberFormat` **cele stare kolone** ako je nova `General` — datumi/iznosi se više ne prikazuju kao goli brojevi (ne dira namerni format novog šablona; ranije se ručno prepodešavalo).
- **Sanity gate + progres:** ako izabrani stari fajl nema `tblOtkup` (potpis baze) → prekid uz prijavu (pogrešan fajl više ne izgleda kao „0 redova = uspeh"). `StatusBar` po tabeli i po 1000 redova + `Cursor = xlWait` — velika migracija više ne izgleda zamrznuto (da operater ne ubije Excel usred upisa).
- **Aktivacija licence se ne migrira:** machine-bound aktivacioni ključevi (`LICENSE_KEY/TOKEN/BOUND_PARTS/HWM/STATUS/NEXT_CHECK`, `TRIAL_START/HWM`) se **ne prenose** (nova mašina re-aktivira), a config licence (`LICENSE_ENABLED/ENDPOINT`, `TRIAL_ENABLED/DAYS`) **prelazi** (nema nepotrebnog re-setup-a na istoj mašini). Nepoznat `LICENSE_*`/`TRIAL_*` ključ se **prenese ali prijavi** (da budući aktivacioni ključ ne otputuje tiho).
- **Robusnost > brzina:** redovi se dodaju kroz `ListRows.Add/Delete` (namerno **ne** `ListObject.Resize`, koji ume tiho da proguta sadržaj ispod tabele).
- **Rizik za podatke:** nema promene poslovne logike — alat je jednokratna migracija; sve provere su read-only; obavezan backup + svesno snimanje štite rezultat. `.frx` netaknut; VBA izvor **ASCII-only**; bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.
- **Otvoreno (backlog):** verzijski-svesna migracija (`StaroImeKolone` čita `APP_VERSION` starog + rename mapa po verziji — rename se sada glasno prijavi, ali migracija nije automatska); opcioni **suvi prolaz** (preview mapiranja + provere bez upisa); licencni **migration mode** (same-machine vs new-machine, target-clean).
- **Dodirnuti moduli:** `modMigracija` (jedini izmenjen). Koristi postojeći `modJournaling.StopAutoSaveTimer` (bez self-update zamke). Test: `git pull` grane → `ImportAllVBA → Debug→Compile → snimi`.

---

## vba-v2.29.0 — 2026-07-22
> Verzija/datum se **finalizuju pri `tools/release.sh`** (uz `APP_VERSION` u `modConfig`). **RF-27 (M2) — Agrohemija: cena + validacija (AUD-040, P1 finansijski/audit).** Zatvara milestone M2. Poslovna šema se ne dira; promene su u knjiženju magacina i validaciji.

- **Izlazna cena se knjiži tačno (AUD-040):** pri „Završi izdavanje" magacin je do sada **ponovo čitao master cenu artikla** umesto cene iz korpe (ulaz je to radio ispravno — asimetrija je bila dokaz previda). Sada se knjiži **cena iz korpe** (snapshot pri dodavanju), pa `tblMagacin.Cena`/`Vrednost` odgovaraju onome što je operater video — dug kooperanta se više ne potcenjuje.
- **Nulta/nevalidna cena se ne knjiži tiho (fail-closed):** realan artikal sa cenom ≤ 0 (nenumerička/prazna master cena) je ranije upisivao `Cena=0/Vrednost=0` i tiho umanjivao dug. Sada takav izlaz **pada** — red se ne upisuje. Rezervisani `ART_POCETNI_DUG` (migracija početnog duga) je izuzet i radi kao pre.
- **Operater vidi tačan razlog:** knjiženje sada javlja konkretnu poruku („Cena za artikal … mora biti veca od 0", „Artikal ne postoji", „Kooperant ne postoji", „Parcela … ne pripada kooperantu") umesto generičke greške pri čuvanju (forma zove `SaveMagacinCore` koji diže tipiziranu grešku; omotač `SaveMagacin` za stare pozivaoce ostaje).
- **Referencijalne provere pri izlazu:** artikal mora postojati; kooperant mora postojati; **parcela mora pripadati tom kooperantu i biti aktivna** (`;`-lista se proverava po stavci; kad je `PRACENJE_PARCELA` uključeno parcela je obavezna, isključeno → prazna dozvoljena). Sprečava upis magacinskog reda na tuđu/nepostojeću/neaktivnu parcelu ili nepostojeći entitet.
- **Besplatan/korektivni ULAZ (novo):** prijem sa cenom 0 je sada moguć **samo uz izričitu potvrdu** („Cena je 0. Proknjižiti besplatan/korektivni prijem?") — za dokumentovane besplatne/korektivne prijeme. Izlaz i negativna cena ostaju strogi.
- **Rizik za podatke:** nema promene šeme; poslovne tabele i `.frx` se ne diraju; VBA izvor **ASCII-only**. **Nov `Poruka()` ključ** `AGRO_MSG_POTVRDI_BESPLATAN_ULAZ` → posle importa **obavezno pokreni `EnsurePoruke`**.
- **Testovi:** nov `modAgrohemijaTests.RunAgrohemijaSmokeSuite` — izolovan (dev-guard, `modJournaling` test-mode koji gasi journaling+AutoSave, i TX rollback → **bez ijednog traga** u ledgeru/journalu/AutoSave-u; log samo u Immediate): snapshot cena, fail-closed (0/nenumerička/negativna), nepostojeći artikal/kooperant, parcela (tuđa/nepostojeća/neaktivna/`;`-lista), multi-stavka rollback, početni dug, zero-value ULAZ sa/bez potvrde, i provera da forma prosleđuje korpa cenu (VBProject wiring, uz „Trust access to VBA project").
- **Napomena:** „aktivan" status ne postoji kao kolona za `tblArtikli`/`tblKooperanti` u šemi (samo `tblParcele`), pa se aktivnost proverava samo za parcele.
- **Dodirnuti moduli:** `modAgrohemija` (`SaveMagacinCore`/`SaveMagacin` split, `ValidateMagacinInput` + parcela↔koop), `frmAgrohemija.frm` (izlaz/ulaz pozivi, zero-value potvrda; `.frx` netaknut), `modPoruke` (nov ključ), `modJournaling` (test-mode toggle), nov `modAgrohemijaTests`. Prateći: `docs/KNOWN_ISSUES.md` (AUD-040 zatvoren), `docs/REFAKTOR_PLAYBOOK.md` (RF-27 status). Test: `git pull` grane → `ImportAllVBA → Debug→Compile → EnsurePoruke → snimi`.

---

## vba-v2.29.1 — 2026-07-23
> Verzija/datum se **finalizuju pri `tools/release.sh`** (uz `APP_VERSION` u `modConfig`). Dopuna `modMigracija` paketa (v2.28.3) — dva fixa: (1) **compile na starijem Excel-u** i (2) **licenca se ipak prenosi na ISTOJ mašini**.

- **Fix — `member not found` na starijem Excel-u (kritično):** v2.28.3 je uveo **early-bound** `ThisWorkbook.AutoSaveOn` (property tek u Excel 2016+/365). Na starijem Excel-u to je **compile greška** koju `On Error` ne hvata i koja obori **ceo** VBA projekat (`Debug→Compile` puca; migracija se ne pokreće). Sada je **late-bound** (preko `Object` promenljive) → kompajlira se svuda; na starom Excel-u runtime `438` se preskoči (cloud AutoSave tamo ionako ne postoji, a `StopAutoSaveTimer` + prljav fajl + backup i dalje štite).
- **Licenca — same-machine gate (ispravka ponašanja iz v2.28.3):** v2.28.3 je **uvek** preskakao machine-bound aktivacione ključeve (`LICENSE_KEY/TOKEN/BOUND_PARTS/HWM/STATUS/NEXT_CHECK`, `TRIAL_START/HWM`) → migracija na istoj mašini je nepotrebno tražila re-aktivaciju. Sada se aktivacija **prenosi na ISTOJ mašini** — otisak stare vezane mašine se poredi sa ovom mašinom **istim pragom kao sama provera licence** (`modLicense.LicPartsMatch(GetDeviceParts(), LICENSE_BOUND_PARTS) >= LIC_MIN_MATCH`). Na **drugoj/neutvrđenoj** mašini se i dalje preskače (jer bi `BOUND_PARTS` sa tuđim otiskom **zaključao** novu mašinu). Config licence (`LICENSE_ENABLED/ENDPOINT`, `TRIAL_ENABLED/DAYS`) prelazi uvek. Rezultat migracije prijavljuje odluku („licenca: ISTA masina → prenosi" / „druga → re-aktivacija").
- **Rizik za podatke:** nema — jedina izmena koda je `modMigracija.bas` (reuse postojećih Public `modLicense.GetDeviceParts`/`LicPartsMatch`); ostalo netaknuto. `.frx` netaknut; VBA izvor **ASCII-only**; bez novih `Poruka()` ključeva → posle importa **ne treba `EnsurePoruke`**.
- **Dodirnuti moduli:** `modMigracija` (late-bind `AutoSaveOn`; `sameMachine` gate + helperi `JeIstaMasina`/`StaroConfigVrednost`). Koristi postojeće Public `GetDeviceParts`/`LicPartsMatch` iz `modLicense`.

---

## vba-v2.29.2 — 2026-08-04
> Verzija/datum se **finalizuju pri `tools/release.sh`** (uz `APP_VERSION` u `modConfig`). **RF-03 (M3) — Storno correctness (AUD-020, AUD-021, AUD-049).** Jezgro storno sloja + storno celog bankovnog izvoda. Poslovna šema se ne dira.

- **Storno novca iz Dokumenata je uopšte proradio:** grana „Novac" je prosleđivala ukucani **broj** dokumenta tamo gde se očekuje `NovacID`, pa je storno novca iz UI-ja **uvek** padao greškom iz dubine. Sada se broj prvo razreši u `NovacID`, a nepostojeći/već storniran red daje jasnu poruku pre potvrde — isti obrazac kao grane Otpremnica i Faktura.
- **Storno uplate vezane za otkup više ne ostavlja blok kao „Isplaćeno" (AUD-021):** čitao se samo `FakturaID` i osvežavao status fakture, pa je otkupni blok ostajao plaćen sa ustajalim `DatumIsplate`. Sada se čita i `OtkupID` i status otkupa se osvežava; rollback pokriva i tu izmenu.
- **Lažno „završena" ispravka zbirne (AUD-020):** povratna vrednost relinka otpremnica se odbacivala, pa je pad relinka davao rekalkulaciju 0/0, invarijanta je prolazila (0=0) i kontekst se zatvarao kao uspešan. Sada se broj prevezanih otpremnica proverava; 0 uz aktivne otpremnice na staroj zbirni → ispravka se označava za ručnu obradu.
- **Prazan correction context više ne guta recovery red:** 6 grana je menjalo podatke i kad kontekst nije napravljen, pa se gubio red u `tblStornoVeze` i MANUAL flag. Dodat hard-stop na svih 6 mesta.
- **Storno celog bankovnog izvoda (AUD-049, novo):** pojedinačni storno izvodnog reda je zabranjen, pa pogrešno mapiran izvod dosad nije imao putanju ispravke. Sada `StornoIzvod_TX` u **jednoj** transakciji obara sav novac izvoda i vraća staging, sa dva ishoda — **remap** (stavke nazad „za obradu", izvod ostaje uvezen, PDF se ne uvozi ponovo) i **ponovni uvoz** (izvod se gasi, isti PDF se može uvesti opet). Novac i staging padaju zajedno; inače bi ponovni uvoz + mapiranje dali dvostruko knjiženje.
- **Pripadnost izvodu se više ne pogađa:** uklonjena heuristika „isti `BrojDokumenta` + `PartnerID`" (mogla je da obori tuđi ručni red) i broj-bazirana zaštita porekla. Red pripada izvodu isključivo po **markeru** koji direktan upis dobija pri knjiženju a split nasleđuje od roditelja. Dodata rekonsilijacija iznosa po stavci — zbir aktivnog novca pod markerom mora odgovarati iznosu stavke (uplate i isplate zasebno), inače se ceo storno odbija sa iznosima u poruci.
- **Kanal plaćanja je sada eksplicitan:** `Tip` nosi keš odvojeno od virmana, pa avans po otkupnom mestu broji **oba** kanala. Izvod bez broja se odbija pri uvozu (ukinut „IZVOD" fallback).
- **Rizik za podatke:** nema promene šeme; `.frx` netaknut; VBA izvor ASCII-only.
- **Dodirnuti moduli:** `modStorno` (nova sekcija IZVOD), `modStornoFlow`, `modNovac`, `modBankaMapiranje`, `modConfig`, `modIzvestaj`, `frmDokumenta`, `frmBankaImport`, `modTestStorno` (T25/T28/T29/T32/T35/T36). Prateći: `docs/KNOWN_ISSUES.md` (AUD-049 zatvoren).

---

## vba-v2.30.0 — 2026-08-05
> Verzija/datum se **finalizuju pri `tools/release.sh`** (uz `APP_VERSION` u `modConfig`). **RF-04 (M3) — Hladnjača auto-lanac: propagacija neuspeha (AUD-005 / FM-0010 #1, #2, #4, #5).** Poslovna šema se ne dira; promene su u orkestraciji lanca otpremnica→zbirna→prijemnica i u backfill makrou.

- **Lanac se više ne prijavljuje kao uspešan kad nešto padne:** rezultati `SaveOtpremnica_TX` i `SaveZbirna_TX` su se ranije odbacivali (zbirna se zvala kao naredba, povrat potpuno ignorisan), pa je operater dobijao potvrdu i kad otpremnica ili zbirna nisu nastale. Sada se prati **svaki** korak po klasi, a upozorenje nabraja tačno šta je palo — otpremnica / zbirna / prijemnica / veza sa otkupom. Stara poruka je tvrdila „otpremnica i zbirna su kreirane" i onda kad nisu.
- **Pad koraka zaustavlja lanac (fail-fast):** ranije je posle pada otpremnice svejedno nastajala zbirna, pa i prijemnica — dakle dokumenti bez uzvodnog dokumenta, koje postojeći backfill (usidren na otpremnicu) ne ume da sanira. Sada pad otpremnice ili zbirne **preskače ostatak te klase**. Bitna posledica: prijemnica u već palom lancu je paletizovala broj i time **blokirala ponovni unos** („broj prijemnice je već paletizovan") — to više ne može da se desi. Klase ostaju nezavisne: pad Klase I ne obara Klasu II.
- **Palete se ne prevezuju na nepostojeću prijemnicu:** broj nove prijemnice se izlaže pozivaocu tek **posle stvarno kreirane** prijemnice. Ranije se postavljao odmah po generisanju, pa je „Unos ispravke" mogao da relinkuje osirotele palete na prijemnicu koja nikad nije nastala.
- **Otkup bez veze sa dokumentom se prijavljuje:** upis `OtpremnicaID`/`BrojZbirne` nazad u otkup red je radio bez signala o ishodu (rollback + log, bez re-raise), pa je neuspeh prolazio nezapaženo. Sada ulazi u upozorenje. Prazan `OtkupID` se više ne broji kao uspeh — `SaveOtkupMulti_TX` garantuje ID za svaku aktivnu klasu, pa je prazan ID prekršen ugovor, ne „nema šta da se veže".
- **Backfill: obe klase istog dokumenta dele broj prijemnice.** Ako je jedna klasa već imala prijemnicu, druga je dobijala **nov** broj umesto postojećeg. Uz to su mape sada ograničene na **hladnjača-kupca** — numeracija prijemnica je per-kupac, pa je prijemnica drugog kupca sa istim `BrojZbirne` mogla i da pozajmi broj i (gore) da preskoči legitimnog kandidata, tj. da backfill tiho ne odradi posao.
- **Rizik za podatke:** nema promene šeme; poslovne tabele i `.frx` se ne diraju; VBA izvor **ASCII-only**; **bez novih `Poruka()` ključeva → posle importa ne treba `EnsurePoruke`**. Promena je vidljiva operateru samo kao precizniji tekst upozorenja i kao izostanak polovičnih dokumenata.
- **Testovi:** `RunBusinessFlowProSuite` prošireno sa 7 testova hladnjača lanca (kompletan lanac, pad otpremnice/zbirne/prijemnice/veze, backfill deljenje broja, backfill izolacija po kupcu) — **164/164**. Uveden test seam `ArmHladnjacaTestFail` (jednokratan, troši se na ulazu u lanac, u produkciji uvek prazan) i `BackfillPrijemniceHladnjacaCore` (silent varijanta bez prompta; `Alt+F8 → BackfillPrijemniceHladnjaca` ostaje nepromenjen). Silent backfill se u testovima ograničava na jedan `BrojZbirne` da suite ne bi dirao prave dokumente, a suite na startu ispisuje upozorenje da piše u svesku i da se pokreće nad test kopijom.
- **Dodirnuti moduli:** `modAutoHladnjaca` (fail-fast lanac, `LinkOtkupRedNaDokument` → `Boolean`, kvalifikovane backfill mape, seam + `…Core`), `modBusinessFlowProTests` (7 testova, fixtures `SeedHladnjacaStanica`/`SeedKupac2`, disclaimer u `BeginRun`). Prateći: `docs/KNOWN_ISSUES.md` (AUD-005), `docs/REFAKTOR_PLAYBOOK.md` (RF-04 status).

---

## vba-v2.31.0 — 2026-08-05
> Verzija/datum se **finalizuju pri `tools/release.sh`** (uz `APP_VERSION` u `modConfig`). **RF-05 (M3) — frmDokumenta unos + storno set (AUD-008 vidi RF-03, AUD-009, AUD-022, deo AUD-003).** Poslednji paket M3. **Ima promenu šeme** — vidi „Rizik za podatke".

- **Stornirana faktura više ne ulazi u listu za plaćanje/avans (AUD-009):** jedini filter je bio `Status <> "Placeno"`, a storno postavlja `Status = "Stornirano"` — dakle stornirana faktura je ostajala izborna i na nju se mogla proknjižiti uplata. Forma sada koristi centralni read-model `modNovac.GetOpenFakture` (izbacuje stornirane, traži `Neplaceno`, vraća samo preostalo > 0) umesto sopstvene petlje; helper je dopunjen kolonom `Datum` za prikaz. **Posledica:** fakture sa nestandardnim statusom ili bez preostalog duga se više ne pojavljuju u listi.
- **Klasa II se ne može tiho izgubiti (AUD-022):** ako izvorne otpremnice imaju Klasu II a čekboks „Dve klase" je isključen, snimanje je slalo `hasKlasaII:=False` i Kl.II se odbacivala bez ijedne poruke. Sada validacija označava stanje kao neispravno (uz iznos Kl.II u labeli), a „Unesi" blokira snimanje i vodi fokus na čekboks.
- **Smer ambalaže je obavezan uz količinu (AUD-022):** unos količine bez izabranog smera je padao u `Case Else` i tiho knjižio legacy „OM prima od vozača" (`Stanica ULAZ`). Sada UI blokira takav unos, a `SaveOMUlaz_TX` odbija prazan/nepoznat smer i na nivou servisa. Operater koji je koristio tihi default bira **„Izdato OM"** (vozač predaje na OM). Istorijski redovi ostaju netaknuti.
- **Malina auto-zbirna više ne pada nevidljivo (AUD-022):** povratna vrednost `AutoCreateZbirnaFromOtpremnice_TX` se odbacivala, a greška je završavala samo u logu — operater je video „Otpremnica sačuvana" i ništa više. Sada se hvata i povrat i greška; ako zbirna nije nastala, poruka to izričito kaže.
- **Prefill ispravke uzima pravu generaciju dokumenta (AUD-022):** ranije je uzimao **prvi** red pronađen po broju, dakle najstariju generaciju, a Klasu I i Klasu II je birao nezavisno (mogle su doći iz različitih ispravki). Uvedena je eksplicitna kolona **`GeneracijaID`**: svi redovi jednog upisa je dele, ispravka posle storna dobija novu. Prefill polazi od PK-a stornirane (`OldDocID` iz correction context-a) i uzima obe klase samo iz njene generacije. Poslovni `Datum` se ne koristi kao kriterijum — ispravka može nositi raniji datum od originala.
- **Prosek gajbe ne računa stornirane redove:** `SumByBroj` (izvor za prosek po otpremnici i po zbirnoj) sabirao je i stornirane redove.
- **`SaveZbirna` upisuje po imenu kolone (deo AUD-003):** pozicijski `Array(...)` je zamenjen `BuildZbirnaRowData` (isti obrazac kao prijemnica), pa promena redosleda kolona ili kolona umetnuta u sredinu ne mogu tiho iskriviti red.
- **Storno po broju ne dira tuđi dokument (novo, otkriveno tokom review-a):** broj dokumenta nije globalno jedinstven — `GenerateBrojPrijemnice` računa sekvencu po kupcu a x-deo je fiksno „1", pa dva kupca istog dana oba dobiju `1/ddmmyy`. Storno po broju zahvata sve aktivne redove tog broja (Kl.I + Kl.II dele broj), pa je storno jednog dokumenta mogao da obori i tuđi. Sada `RequireJedanVlasnikPoBroju` odbija takav storno **unutar transakcije** (rollback, nijedan red se ne menja) na svim putanjama — SIMPLE, ISPRAVKA, DUPLI, kao i u malina/autohladnjača **kaskadama**. Kaskade dodatno razrešavaju vlasnika lanca (`ResolveZbirnaChainScope`) jednom pre prve mutacije i obaraju isključivo redove tog vlasnika — tuđa prijemnica/otpremnica pod istim `BrojZbirne` ostaje netaknuta, a osiroteli nizvodni dokument bez aktivne zbirne zaustavlja storno umesto da bude tiho oboren. Vlasnik: otpremnica → `StanicaID`, prijemnica → `KupacID`, zbirna → `VozacID` + `KupacID`.
- **Rizik za podatke:** **promena šeme** — nova kolona `GeneracijaID` na `tblOtkup`, `tblOtpremnica`, `tblZbirna`, `tblPrijemnica`, `tblFakture`, `tblNovac`; dodaje je `EnsureSledljivostSchema` automatski na svakom startu (nije potreban ručni korak). Postojeći redovi ostaju bez vrednosti — prefill za njih koristi konzervativni fallback (samo poslednji red, druga klasa prazna). Poslovni podaci se ne migriraju; `.frx` netaknut; VBA izvor **ASCII-only**; **bez novih `Poruka()` ključeva → posle importa ne treba `EnsurePoruke`**.
- **Testovi:** `RunBusinessFlowProSuite` **276/276** (+ kaskadni guard test) (dodato 10 RF-05 testova: prosek gajbe, read-model otvorenih faktura, Kl.II blokada, generacija i prefill, malina signal pada, column-mapped zbirna, obavezan smer ambalaže, storno guard na svim putanjama), `RunStornoTestSuite` **181/181**. Uveden test seam: `ZbirnaIzvorImaKlasuII`, `PickPrefillRows` i `SaveOMUlaz_TX` premešteni iz forme u `modDokumenta` (nema referenci na kontrole), `AutoCreateZbirnaFromOtpremnice(_TX)` dobio opcioni scope po `BrojOtpremnice` da testovi ne diraju nepovezane dokumente.
- **Dodirnuti moduli:** `frmDokumenta`, `modDokumenta` (generacija + prefill picker + OM ulaz servis + `BuildZbirnaRowData`), `modStorno` (guard vlasnika), `modStornoFlow` (`OldDocID` za zbirnu, guard u atomic otpremnici), `modNovac` (`GetOpenFakture` + `Datum`), `modMasterSync` (PWA import generacija, scoped auto-zbirna), `modDokumentInvariant` (generacija pri rekalkulaciji), `modConfig` (`COL_GENERACIJA_ID`), `modSetup` (`EnsureSledljivostSchema`), `modBusinessFlowProTests`. Prateći: `docs/KNOWN_ISSUES.md` (AUD-009/AUD-022 zatvoreni, AUD-052 nov), `docs/REFAKTOR_PLAYBOOK.md` (RF-05 status).

---

## vba-v2.32.0 — 2026-08-06
> Verzija/datum se **finalizuju pri `tools/release.sh`** (uz `APP_VERSION` u `modConfig`). **RF-14 (M4a) — MasterSync / Sheets JSON (AUD-001, AUD-002, deo AUD-018).** Prva dva su preostala P0 cele sanacije. Poslovna šema se ne dira; promene su u čitanju Google odgovora i u tome kako se greška razlikuje od praznog rezultata.

- **Tekst iz PWA više ne stiže iskvaren (AUD-001):** čitač Google odgovora je pre parsiranja radio globalni `Replace(", " → ",")`, koji nije bio svestan navodnika — pa je brisao razmak posle zapete u **svakoj** tekst-ćeliji (adrese, imena, napomene: „Ulica 5, Beograd" → „Ulica 5,Beograd"). Uz to je escape-ovani navodnik lomio praćenje stringa i cepao red na pogrešnom mestu, `\uXXXX` nikad nije dekodovan (srpska slova i `& < > ' "` prolazili doslovno), a red se delio literalnim traženjem `],[` — pa je ćelija koja sadrži taj niz znakova pravila lažan red. Sve je zamenjeno jednim prolazom kroz ceo odgovor koji ispravno prati navodnike i escape-ove.
- **Skraćen ili pokvaren odgovor se više ne prihvata kao podatak:** ranije se nedovršen red vraćao kao da je ceo. Red kome fali rep može proći validaciju, biti lokalno upisan i na Google-u označen kao `Synced>Master` — i time **trajno izgubiti polja sa kraja**. Sada je čitanje fail-closed: nezatvoren string, neuravnotežene zagrade, smeće posle kraja odgovora, nepoznat escape (`\x`), neispravan `\u12G4` i `values` koji nije niz (`null`/`"x"`/`{}`/broj/bool) obaraju čitanje umesto da vrate pola reda. **Prazan sheet i dalje prolazi normalno** — razlikuje se „prazno" od „nije pročitano".
- **Pad jednog sheeta više ne briše ono što je već uvezeno (AUD-002):** uvoz OTK-a je ceo batch držao u jednoj transakciji, a Google potvrda (`Synced>Master`) se ne može poništiti. Pad kasnijeg sheeta je rollback-om obrisao **lokalne** redove ranijih sheetova, dok su njihovi redovi na Google-u ostajali potvrđeni — pa ih sledeći ciklus preskače i **nikad se više ne isporuče**. Batch transakcija je uklonjena (isti model koji VOZ uvoz već koristi); svaki red i dalje ide kroz sopstvenu transakciju.
- **Neuspelo čitanje sheeta zaustavlja uvoz umesto da izgleda kao prazan sheet:** OTK i VOZ uvoz sada prekidaju **pre** obrade ijednog reda i pre upisa statusa nazad na Google, i to se broji kao fatalna greška ciklusa.
- **Preko 100 vozača se više ne gubi tiho (deo AUD-018):** listanje VOZ sheetova je radilo jedan zahtev bez paginacije, pa je sve preko prvih 100 nestajalo bez ijedne poruke.
- **Predlog broja dokumenta se ne pogađa kad Google ne odgovori:** remote provera zauzetih brojeva je svaku grešku (Drive lookup, čitanje, nedostajuće kolone, bilo koja neočekivana greška) tretirala kao „na Google-u nema brojeva" i mogla predložiti **već zauzet** broj. Sada se greška propagira, forma ostaje prazna i operater unosi broj ručno. Sheet koji stvarno ne postoji i dalje legitimno daje nulu. Prazan rezultat se više ne pamti do restarta, pa se sheet koji PWA napravi u međuvremenu odmah vidi.
- **Nema više duplih Google tabela:** obrazac „nađi pa ako nema kreiraj" je grešku pri traženju čitao kao „ne postoji" i pravio **drugu** tabelu istog imena. Najopasnije na masovnoj putanji za stanice (duplikat `OTK-*` za više stanica odjednom) i na first-run putanjama za `Stammdaten`/`Kartice`/`MgmtReports`. Sada se kreira isključivo posle uspešne provere.
- **Lock tabela se ne prepisuje nepotpuna:** `SyncControl` se čita pa upisuje nazad; neuspelo čitanje je davalo prazan skup, pa je upis brisao lockove drugih stanica i ostala podešavanja. Sada se upis prekida.
- **Health check više ne laže:** provera dostupnosti Google foldera je prijavljivala „dostupno" i onda kada je lookup padao.
- **Rizik za podatke:** nema promene šeme; poslovne tabele i `.frx` se ne diraju; VBA izvor **ASCII-only** (dekodovanje `\uXXXX` je runtime `ChrW`, ne literal). **Izmenjen je tekst poruke `SYNC_MSG_PWA_UVOZ_NIJE`** (više ne tvrdi da su promene vraćene, jer batch rollback-a nema) → posle importa pokrenuti **`EnsurePoruke`**.
- **Testovi:** nova offline suite `RunSheetsJsonParserTests` (18 grupa: razmak posle zapete, `\"`, `\uXXXX`, embedded `],[`, prettyPrint/`\\`/`\t`/`\n`, brojevi bez navodnika, pa negativne — prekid usred stringa, prekid posle nekoliko ćelija, neuravnotežen dokument, pogrešan tip `values`, HTML telo) — **sve zeleno**. `RunMasterSyncSmokeSuite` prošireno sa 6 testova (read-failure je fatal bez importa i writeback-a, prazan sheet nije greška, remote scan brojeva fail-closed, get-or-create bez duplikata, cross-sheet pad ne poništava prvi sheet, VOZ listing preko 100 sheetova). Suite-ovi su sada **tvrd gate**: `RunSheetsJsonParserTests`, `RunGoogleSyncSmokeSuite` i `RunMasterSyncSmokeSuite` podižu grešku kad interno padnu, pa `RunE2EReleaseGate_v610` više ne može da ih prijavi kao PASS.
- **Dodirnuti moduli:** `modGoogleSheets` (parser + `TryParseValuesJson`/`TryReadSheetData`/`TryGetSpreadsheetID`/`TryGetOrCreateSpreadsheetID`/`IsWellFormedJsonDocument`), `modMasterSync` (batch bez outer TX, fail-closed read u OTK/VOZ, VOZ paginacija, `ExtractNextPageToken`, test seam-ovi), `modBrojevi` (remote scan fail-closed, bez negativnog keša), `modStanicaLock` (`TryReadSyncControlAsDict`), `modStammdatenSync` + `modGoogleSyncOrchestrator` (get-or-create), `modProductionHealthCheck` (probe), `modPoruke` (tekst poruke), `modGoogleSyncSmokeTests`, `modE2EReleaseGate`. Prateći: `docs/KNOWN_ISSUES.md` (AUD-001/AUD-002 zatvoreni, AUD-018 delimično), `docs/REFAKTOR_PLAYBOOK.md` (RF-14 status — posle merge-a).

---

## vba-v2.33.0 — 2026-08-06
> Verzija/datum se **finalizuju pri `tools/release.sh`** (uz `APP_VERSION` u `modConfig`). **RF-28 (M4b) — MasterSync integritet (AUD-041, AUD-042, AUD-043, AUD-046).** Nastavak RF-14: tamo se popravljalo *čitanje* Google odgovora, ovde *šta se sa pročitanim redom radi* — grupisanje u otpremnice, dodela brojeva, vezivanje zbirne i FK vozača. Poslovna šema se ne dira.

- **Auto-otpremnica više ne meša vrste i cene (AUD-043):** otpremnice su se pravile grupisanjem po `Stanica|Datum|Vozač|Klasa`, količine su se sabirale, a **vrsta, sorta, cena i tip ambalaže su se čitali sa prvog reda grupe**. Jabuke i kruške istog vozača istog dana su tako išle kao **jedna** otpremnica „sve jabuke", po ceni jabuka i u ambalaži prvog reda — pogrešna roba i pogrešan novac na dokumentu. Grupni ključ sada sadrži i vrstu, sortu, cenu i tip ambalaže, pa svaka kombinacija dobija svoju otpremnicu. Okidalo se rutinski, na svakom mešanom danu.
- **Zbirna više ne dobija već zauzet broj (AUD-041):** generator broja zbirne pri uvozu je **brojao redove** umesto da traži najveću sekvencu, pa je svaka rupa u nizu davala duplikat — ako u bazi postoje `1/ddmmyy` i `1/ddmmyy-3`, predlog je ponovo bio `-3`. Sada koristi isti kanonski generator kao ostatak aplikacije (najveća sekvenca + provera da broj nije zauzet). Kad broj ne može da se dodeli, red se **vidljivo** označava greškom umesto da dobije pogrešan broj.
- **Vezivanje zbirne ne prepisuje tuđu vezu (AUD-043):** uvoz zbirne je upisivao `BrojZbirne` na otkup i otpremnicu **bezuslovno**, pa je jedna zbirna mogla „preuzeti" otkupe koji već pripadaju drugoj — roba dvostruko obračunata, a prva zbirna ostaje bez stavki. Sada se upisuje samo ako je polje prazno ili već nosi isti broj; sve ostalo je konflikt koji zaustavlja taj red (bez ikakve izmene u bazi).
- **Zbirna prima samo svoje otkupe (AUD-043):** lista otkupa je dolazila iz PWA reda i prihvatala se bez provere — pogrešan identifikator je vezivao **tuđi** otkup (drugi vozač, drugi dan) u zbirnu i time kvario i broj (`ddmmyy`) i obračun po vozaču. Sada se proverava da otkup pripada vozaču zbirne i njenom poslovnom danu; utovar posle ponoći (susedni dan) prolazi uz upozorenje u logu, veća razlika zaustavlja red. Otkup kojem vozač još nije sinhronizovan sa terena i dalje normalno prolazi.
- **Neuspeo upis vozača se više ne prijavljuje kao uspešan sync (AUD-042):** kad je red već u bazi a sa terena dođe vozač, upis se radio **bez provere da li je prošao**, i red je na Google-u ipak dobijao završni status. Posledica: `VozacID` ostaje prazan, red se **nikad više ne isporučuje**, a ciklus se završava kao zelen. Sada se ishod razlikuje — upisano / nema šta da se menja / upis pao / **na terenu je drugi vozač nego u bazi**. Poslednja tri označavaju red greškom, broje se i obaraju ceo ciklus na crveno, a različit vozač se tretira kao konflikt podataka i **ne** prepisuje bazu.
- **Neispravan datum više ne postaje današnji (AUD-042):** ako datum iz PWA reda nije mogao da se pročita, oba uvoza (otkup i zbirna) su ga **tiho zamenjivala današnjim** — dokument je dobijao pogrešan poslovni dan i pogrešan `ddmmyy` u broju, a red je izgledao uspešno uvezen, pa se greška nije mogla ni naći. Sada takav red ide u grešku sa objašnjenjem. Prazno, tekst, samo vreme bez datuma i 1899-baseline vrednost nisu datum; realni datumi i unos sa zakasnelim datumom prolaze normalno.
- **Nedovršena Google tabela se čisti sama (AUD-042):** ako se `OTK-*` tabela napravi a upis zaglavlja padne, sledeći ciklus ju je po imenu našao kao „postoji", preskočio — i zaglavlje **nikad** nije upisano, dok PWA piše u tabelu bez šeme. Takva tabela sada odlazi u Drive korpu, pa sledeći ciklus pravi čistu. Ako brisanje ne prođe, u logu stoji izričito uputstvo da se obriše ručno.
- **Broj prijemnice se ne izmišlja posle greške (AUD-041):** kad bi računanje broja palo, vraćao se `1/ddmmyy` — broj koji **izgleda** kao regularan prvi broj dana, a već postoji. Sada se u tom slučaju ne vraća broj (isto ponašanje kao kod otkupa i otpremnice), pa korak lanca prijavi pad umesto da napravi duplikat.
- **Dokument ne dobija vozača koji ne postoji (AUD-046):** u malina/hladnjača konvenciji je „vozač = stanica", pa se `VozacID` postavljao na `StanicaID` **bezuslovno** — i kad par-vozač u `tblVozaci` ne postoji. Dokument time dobija vezu bez pokrića i svako spajanje na vozača (izveštaji, ambalaža, nalozi za banku) vraća prazno ime. Sada postoji jedna provera para (`tblStanice` + `tblVozaci`) koju oba mesta pitaju **pre** upisa: uvoz preskoči red uz upozorenje u logu (jedna problematična stanica ne obara ceo prolaz), a auto-lanac hladnjače se ne pokreće i operater dobija poruku. Pravljenje par-vozača više ne prolazi tiho — traži da stanica postoji i da je jedinstvena, a neuspeh se prijavljuje.
- **Greške u ovim tokovima se više ne gube u prolazu:** na tri mesta se greška logovala pa **ponovo dizala iz već obrisanog stanja** (`Err.Raise 0`), zbog čega je propadala bez dijagnostike — uključujući guard koji zaustavlja pogrešno vezivanje zbirne. Sada se podatak o grešci čuva pre logovanja, pa poruka stiže do pozivaoca i transakcija se stvarno vraća.
- **Storniranje na SEF-u sada menja i lokalno stanje fakture.** Do sada je uspešan storno menjao samo spoljni status, pa je faktura trajno ostajala u stanju „poslata"/„prihvaćena" uz status `STORNO` — a takvu fakturu grupna obrada preskače kao završenu, pa je niko više nije ispravljao. Sada prelazi u lokalno stanje `SEF_STORNO`. **Otkazivanje (`Cancelled`) namerno ostaje samo spoljni podatak** — lokalno stanje za otkazivanje ne postoji, pa se beleži status, a faktura se samo izvlači iz „šalje se" ako je tamo bila zaglavljena.
- **Greške pri upisu SEF stanja sada stvarno stižu do operatera.** U celom sloju koji upisuje i proverava SEF stanje (22 mesta) rukovanje greškama je gutalo originalnu grešku — prijava je odlazila u log, a poziv je nastavljao kao da je sve prošlo. To je posebno važilo za samu proveru duplikata i za upis promene stanja, dakle za provere na kojima počivaju sve gornje garancije. Sada se greška hvata pre logovanja i propagira, pa transakcija zaista pada i vraća izmene.
- **Rizik za podatke:** **nema promene šeme**; poslovne tabele i `.frx` se ne diraju; VBA izvor **ASCII-only**; **bez novih `Poruka()` ključeva → posle importa ne treba `EnsurePoruke`**. Ponašanje je strože nego pre: redovi koji su ranije tiho prolazili (neispravan datum, konfliktna veza, neuspeo upis vozača) sada završavaju kao vidljiva greška i traže reakciju operatera — to je i svrha izmene. **Otvorena pretpostavka:** zbirna se tretira kao jednodnevni dokument (+ utovar posle ponoći); ako u praksi legitimno obuhvata više dana, prozor se podiže jednom konstantom (`MASTER_SYNC_MEMBERSHIP_DAY_TOLERANCE`) — zapisano u `KNOWN_ISSUES`.
- **Testovi:** `RunBusinessFlowProSuite` dopunjen sa 7 RF-28 testova (otpremnica se cepa po **svakom** segmentu ključa zasebno — cena, vrsta, sorta, tip ambalaže; rupa u nizu daje najveću sekvencu + 1; konflikt veze ne prepisuje; vezivanje po primarnom ključu i kad dve zbirne dele poslovni broj; prozor dana; neispravan datum je greška a ne današnji — uključujući stvarni ISO format koji PWA šalje; svi ishodi upisa vozača). Svaki test radi u sopstvenoj transakciji i vraća je, pa ne ostavlja redove u svesci — bilo **326/326** pre dopune, očekivano **~336** posle. `RunMasterSyncSmokeSuite` dobio 2 end-to-end testa (neuspeo upis vozača završava kao `SyncError` + fatalna greška ciklusa; nedovršena tabela je stvarno u Drive korpi). Uveden jednokratni test seam (`TestHook_ArmFailSeam`) jer se neuspeh upisa u tabelu i na Google ne može izazvati „prirodno" — u produkciji je prazan.
- **Dodirnuti moduli:** `modMasterSync` (grupni ključ, delegacija broja zbirne, membership + konflikt guard, ishodi upisa vozača, strogi datum, `CreateOTKSheetWithHeader`, test hookovi), `modBrojevi` (`GenerateBrojPrijemnice` bez izmišljenog broja), `modMalina` (`IsManagedStationMirror`, stroži `EnsureVozacMirrorForStanica`), `modAutoHladnjaca` (provera para pre stampanja vozača), `modBusinessFlowProTests`, `modGoogleSyncSmokeTests`. Prateći: `docs/KNOWN_ISSUES.md` (AUD-041/042/043/046 zatvoreni, uz zapisane rezidue), `docs/REFAKTOR_PLAYBOOK.md` (RF-28 status — posle merge-a).

---

## vba-v2.34.0 — 2026-08-06
> Verzija/datum se **finalizuju pri `tools/release.sh`** (uz `APP_VERSION` u `modConfig`). **RF-06 (M5) — ispravnost brojki u izveštajima (AUD-023).** Prvi paket M5 milestone-a. Ne dira poslovnu šemu ni `.frx`; menja se **šta izveštaj računa i prikazuje**, ne kako se podaci upisuju.

- **Isplate se vode na onom otkupnom mestu gde su i izvršene:** „Saldo OM" je isplatu kooperanta pripisivao njegovom **matičnom** otkupnom mestu iz šifarnika, a ne mestu na kom je novac stvarno isplaćen. Kooperant koji preda robu na dva OM-a je time na jednom izveštaju imao isplatu koja tamo ne pripada, a na drugom je nedostajala. Sada odlučuje otkupno mesto **samog reda u tblNovac** — isti ključ koji izveštaj „Isplata" već koristi, pa se dva izveštaja konačno slažu. Stariji redovi koji uopšte ne nose otkupno mesto i dalje idu na matično (da novac ne bi nestao ni sa jednog izveštaja).
- **Kartica kooperanta počinje od stanja duga, ne od nule:** kolona „Saldo" je zapravo prikazivala **neto promenu perioda** — promet pre početnog datuma se nije uzimao u obzir, pa je kartica za jul kod kooperanta koji duguje od juna pokazivala pogrešan (obično manji) dug. Sada kartica ima prvi red **„POCETNO STANJE"** sa saldom duga i saldom gajbi zatečenim na dan `datumOd`, a red UKUPNO prikazuje **završni saldo** (početno + promet perioda); kolone zaduženja i razduženja i dalje pokazuju samo promet perioda. Isto važi i za tab „Pregled ambalaže" i za oba PDF-a (kartica i kartica ambalaže) i za karticu koja se izvozi u PWA.
- **Manjak se više ne računa preko tuđe prijemnice:** izveštaji „Manjak" i „Otkupljena roba (OM)" su zbirne i prijemnice spajali **isključivo po broju zbirne**, a taj broj nije jedinstven — dve aktivne zbirne mogu ga deliti (drugi vozač, drugi kupac). Posledica: obe su videle **zbir obe prijemnice**, pa su i primljena količina i manjak i procenat bili pogrešni. Spajanje sada ide po vlasniku **i klasi** (broj + vozač + kupac + Klasa) — isto pravilo vlasništva koje storno već koristi. Klasa je bitna jer hladnjača vodi Klasu I i Klasu II kao **odvojene** otpremnice, zbirne i prijemnice pod istim brojem: bez nje se primljena količina obe klase sabirala i taj zbir pripisivao **svakoj** klasi (za robu 1.000 kg Klase I i 200 kg Klase II ukupan prijem se prikazivao kao 2.100 kg umesto 1.050 kg). Kad se vlasnik ne može dokazati, red **ne dobija izmišljenu brojku** nego oznaku **„nejasan vlasnik"** i ostaje van UKUPNO. Zbirne čiji je broj jedinstven (ogromna većina) računaju se kao i pre, pa starije prijemnice bez upisanog vozača/kupca i dalje normalno prolaze.
- **„nema prijema" umesto dva različita odgovora na isto pitanje:** otpremnica/zbirna kojoj još nije stigla prijemnica se u „Otkupljena roba (OM)" prikazivala kao **0 kg / 0,00% manjka**, a u izveštaju „Manjak" istovremeno kao **100% manjka** — isti podatak, dva suprotna zaključka, i oba pogrešna (roba nije nestala, samo još nije primljena). Sada oba izveštaja ostavljaju brojke prazne i pišu oznaku **„nema prijema"**. Takvi redovi **ne ulaze u UKUPNO manjak** ni u osnovicu procenta, pa ukupan manjak više ne skače zbog pošiljki koje su tek na putu.
- **Uplata kupca se deli po vrstama voća srazmerno:** za fakturu sa više vrsta (npr. malina + kupina) cela uplata je knjižena na vrstu **prve stavke** fakture, pa je saldo kupca pokazivao dug na jednoj vrsti i višak na drugoj. Sada se uplata deli srazmerno vrednosti stavki (količina × cena). Svaki deo je **zaokružen na paru kako se i prikazuje**, a višak para se raspoređuje po najvećim ostacima — bez toga bi uplata od 100 podeljena na tri vrste prikazala 33,33 + 33,33 + 33,33 = 99,99 uz UKUPNO 100,00. Metod ujedno garantuje da **nijedna vrsta ne dobije negativan iznos**: kod sitne uplate raspoređene na više vrsta (npr. 0,03 na pet vrsta) raniji način računanja je poslednjoj vrsti upisivao −0,01, iako je zbir bio tačan. Fakture sa jednom vrstom se ponašaju kao i pre.
- **Nevalidna kombinacija izveštaja daje praznu listu, ne tuđe brojke:** u zbirnom modu su tabovi „Prosečna cena" i „Manjak" vidljivi i za Kooperante i Vozače, a izveštaj je za takav izbor tiho prikazivao **globalni** rezultat pod naslovom izabranog entiteta. Isto je važilo za ambalažni izveštaj i „Otkupljenu robu". Sada takva kombinacija daje čistu praznu listu. (Vidljiva poruka i uklanjanje nevalidnih tabova iz menija dolaze u sledećem paketu, RF-07.)
- **Neraspoređena agrohemija se više ne broji u svakom otkupnom mestu:** izlaz agrohemije bez kooperanta se ne može pripisati nijednom OM-u (magacin nema kolonu otkupnog mesta), a ulazio je u UKUPNO **svake** stanice — zbir po stanicama je isti trošak brojao više puta. Red ostaje vidljiv radi informacije, labela sada nosi „van UKUPNO", i u UKUPNO se ne uračunava.
- **Rizik za podatke:** **nema promene šeme**; poslovne tabele se ne diraju i nema nijednog novog upisa — izmene su isključivo u sloju čitanja/računanja izveštaja. `.frx` nedirnut (u `frmIzvestaj` je promenjen samo kod prikaza, da tekstualna oznaka „nema prijema" ne bi bila pretvorena u 0). VBA izvor **ASCII-only**; **bez novih `Poruka()` ključeva → posle importa ne treba `EnsurePoruke`**. Očekivane razlike u brojkama posle importa su namerne i navedene gore; kartice i izveštaj Manjak će za iste datume dati **druge** iznose nego pre — to je i svrha izmene.
- **Testovi:** nova assert suite **`RunIzvestajTests`** (**pada glasno** — podiže grešku kad provera padne, pa ne može da se prijavi kao uspešna) (`modIzvestajTests`) — fiksira svaku ispravljenu brojku nad čistim računskim funkcijama (bez tabela, pa je deterministična na svakoj instalaciji): pripadnost isplate stanici, početno stanje i running saldo obe kartice, oznaka „nema prijema" (prazne brojke, ne 0 i ne 100%), srazmerna podela uplate (zbir podele == iznos **i** nijedan deo nije negativan — obe invarijante se proveravaju zajedno) i dispatch matrica izveštaja. Uz to **tri end-to-end testa nad stvarnim tabelama** (dve zbirne istog broja kod dva kupca sa različitim primljenim količinama — zasebno iz ugla „Manjak" i iz ugla „Otkupljena roba (OM)"; plus Klasa I + Klasa II istog dokumenta, gde svaka klasa mora dobiti svoj prijem a ukupan zbir ostati jednostruk) — provere nad izdvojenim funkcijama ne mogu da uhvate grešku u samom spajanju tabela. End-to-end testovi rade u transakciji i **uvek se poništavaju**, ne ostavljaju redove. Postojeći `SmokeTest_modIzvestaj` ostaje kao provera oblika nad živim podacima.
- **Dodirnuti moduli:** `modHelpers` (`BuildManjakDict` scoped na vlasnika+klasu, `ZbirnaVlasnikKljuc`/`ZbirnaStavkaKljuc`/`KlasaOrDefault`), `modIzvestaj` (novi deljeni seam-ovi `PrijemZaZbirnu`, `NovacRedPripadaStanici` / `ManjakStavka` / `KarticaRezultatSaPocetnim` / `KarticaAmbRezultatSaPocetnim`; `ReportSaldoOM`, `ReportKarticaKooperanta`, `ReportKarticaAmbalaze`, `ReportOtkupRobaOM`, `ReportManjak`, `ReportProsecnaCena`, `ReportAmbalaza`, `ReportOtkupRoba`), `modNovac` (`BuildVrstaFakturaCache` → `BuildFakturaVrstaUdeoCache`, nove `RaspodeliPoUdelima` i `ZaokruziNovac`, `GetUplataByVrsta`), `modAutoHladnjaca` (`ClassOrDefault` delegira na deljeno pravilo — jedina izmena), `frmIzvestaj` (prikaz oznake, samo kod), `modIzvestajTests`. Prateći: `docs/KNOWN_ISSUES.md` (AUD-023 zatvoren; uz AUD-013 zapisan nalaz provere), `docs/REFAKTOR_PLAYBOOK.md` i `docs/PLAN_SANACIJE.md` (RF-06 / M5 status — posle merge-a).

---

## vba-v2.35.0 — 2026-08-07
> Verzija/datum se **finalizuju pri `tools/release.sh`** (uz `APP_VERSION` u `modConfig`). **RF-07 (M5) — freshness izveštaja, vidljive greške i revers ambalaže (AUD-024, AUD-012, deo AUD-027).** Drugi paket M5 milestone-a; nastavak RF-06. RF-06 je popravljao *šta izveštaj računa*, RF-07 popravlja *šta izveštaj tvrdi da prikazuje* i **koji revers se odštampa**. Poslovna šema se ne dira, `.frx` netaknut.

- **Izveštaj više ne tvrdi period koji nije prikazan:** status linija i zaglavlje štampe (uključujući PDF kartice kooperanta) su čitali **trenutni sadržaj polja „Od"/„Do"**, dok su podaci na ekranu bili generisani za period koji je važio u trenutku klika na „Prikaži". Promena datuma bez ponovnog „Prikaži" je time davala **staru listu pod novim periodom** — i na ekranu i na odštampanom papiru. Sada i status i sve štampe nose period **stvarno prikazanih podataka**.
- **„NIJE OSVEŽENO" upozorenje:** čim se promeni „Od" ili „Do", status linija odmah pređe u amber upozorenje „NIJE OSVEŽENO — kliknite 'Prikaži'" i uz njega ispiše period podataka koji su još na ekranu. Upozorenje nestaje tek kad se izveštaj stvarno regeneriše. Štampa pre ijednog „Prikaži" je blokirana porukom (nema perioda za zaglavlje).
- **Greška pri generisanju izveštaja se više ne gubi:** kad bi generisanje taba puklo (npr. nedostajuća kolona u tabeli), greška se samo upisivala u log — na ekranu je ostajala **prethodna lista uz zelen status „N redova učitano"**, pa je operater čitao tuđe brojke kao rezultat. Sada se lista **obriše**, panel „Detalji" se resetuje, status postaje crven („GREŠKA: izveštaj nije generisan") i izlazi poruka sa opisom. Tab ostaje neoznačen kao generisan, pa se pokušava ponovo pri sledećem prelasku na njega.
- **Zbirni mod više ne nudi kombinacije koje ne postoje:** tabovi „Zbirni", „Prosečna cena" i „Manjak" su u zbirnom režimu bili vidljivi **svim** tipovima entiteta. RF-06 je takve kombinacije doveo do čiste prazne liste; sada se **uopšte ne nude** — Vozači nemaju „Prosečnu cenu" (taj izveštaj nema vozačku granu), a Kooperanti nemaju nijedan zbirni tab. Kad izabrana kombinacija nema nijedan dostupan izveštaj, status to i kaže umesto da forma ostane prazna bez objašnjenja.
- **Revers ambalaže: stornirano se više ne štampa:** rekonstrukcija reversa je čitala ceo ledger ambalaže **bez storno filtera**, pa su poništeni redovi ulazili u količinu na papiru.
- **Revers ambalaže: tipovi gajbica se više ne mešaju (isti dokument, dve vrste):** tip ambalaže se uzimao sa **prvog** reda dokumenta, a količina sabirala preko **svih** tipova. Dokument sa 25 letvarica i 15 plastičnih gajbica je davao jedan revers na **40 „letvarica"** — pogrešan tip i pogrešna količina na dokumentu koji kooperant potpisuje. Sada je tip ambalaže deo ključa: štampa se revers **tačno onog reda** koji je izabran u pregledu (pregled je i inače grupisan po dokumentu i tipu), pa dve vrste daju dva ispravna reversa.
- **Upozorenje na duplirane noge reversa:** ako isti dokument i tip nose više od jedne noge po strani (duplikat ili više generacija), zbir na reversu je verovatno naduvan — sada se to prijavi pre štampe, uz mogućnost odustajanja.
- **Nema više „cross-tab" štampe iz panela „Detalji":** izabrani red panela je bio zapamćen na nivou modula, nevezano za tab. Klik na red u „Kartici", pa prelazak na „Ambalažu" i klik na „Štampaj dokument" je štampao **dokument sa prethodnog taba**. Promena taba sada čisti izabrani red.
- **Promena entiteta bez dostupnog izbora više ne ostavlja tuđe podatke na ekranu:** kad se pređe na tip entiteta koji nema nijedan aktivan zapis (npr. „Vozači" kad nema aktivnog vozača), osvežavanje se preskakalo — a na ekranu su ostajali **redovi prethodnog entiteta**, sa zelenim statusom, i štampali bi se pod novim naslovom. Sada svaka promena entiteta ili režima **poništava kontekst**: liste se prazne, panel „Detalji" se čisti, status kaže „Prvo kliknite 'Prikaži'". Uz to štampa proverava da je **baš aktivni tab** stvarno generisan, a zaglavlje nosi **zapamćeni entitet** (kao što već nosi zapamćeni period), ne trenutno stanje dugmadi i padajuće liste.
- **„Prosečna cena" se više ne nudi za Kupce u zbirnom režimu:** taj izveštaj za kupca traži prijemnice po `KupacID`, a zbirni režim ne šalje kupca — upit je time tražio prijemnice **bez kupca** i vraćao prazno i kad prijemnice postoje. Tab je uklonjen iz menija. (Za otkupna mesta kombinacija radi i ostaje.) Globalni prosek preko **svih** kupaca nije implementiran — to je nov izveštaj, ne UI podešavanje; zapisano je kao otvoreno poslovno pitanje u `KNOWN_ISSUES`.
- **Tip ambalaže se isto normalizuje u pregledu i na reversu:** pregled je grupisao redove po sirovom tekstu tipa, a revers ih poredio bez obzira na velika/mala slova — pa su „Letvarica" i „letvarica" davali dva reda u pregledu, ali je svaki revers sabirao oba. Sada obe putanje koriste isti ključ.
- **Revers izdate ambalaže uz otkup se konačno može odštampati iz pregleda:** kad otkupni list i primi pune gajbe i izda prazne (ista vrsta gajbica), obe stavke se u ledgeru vode pod **istim brojem dokumenta** ali kao **dva različita dokumenta** — otkup i revers. Pregled ambalaže ih je spajao u jedan red koji je nosio tip prvog zapisa, pa je dugme „Štampaj dokument" uvek štampalo **otkupni list**, a revers za izdate prazne gajbe nije imao svoj red i **nije se mogao odštampati**. Sada su to dva reda: jedan sa ulazom (otkup) i jedan sa izlazom (revers), svaki sa svojom štampom. Ukupni zbirovi se ne menjaju — menja se samo broj prikazanih redova.
- **Rizik za podatke:** **nema promene šeme**; poslovne tabele se ne diraju i nema nijednog novog upisa — izmene su u UI sloju izveštaja i u rekonstrukciji podataka za štampu. `.frx` netaknut (menjan je samo kod forme; nove `Private WithEvents` deklaracije nisu uvedene). VBA izvor **ASCII-only**. **Dodato je 7 novih `Poruka()` ključeva (`RPT_MSG_*`, `RPT_ERR_GENERISANJE_IZVESTAJA`) → posle importa OBAVEZNO pokrenuti `EnsurePoruke`** (bez toga status i poruke prikazuju `[KLJUC]`). Očekivana vidljiva promena: revers za dokument sa više tipova ambalaže sada daje **manje** gajbica po reversu nego ranije — to je ispravka, ne regresija. Uz to pregled ambalaže daje **više redova** za otkupe koji su i primili i izdali gajbice (dva dokumenta, dva reda) — ukupni zbirovi ostaju isti.
- **Testovi — `RunIzvestajTests` zeleno, 192 provere** (**tvrd gate**: suite podiže grešku kad ijedna provera padne, pa ne može da se prijavi kao uspešna). Pet novih grupa uz postojeće RF-06 provere: `T_TabMatrica` (matrica dostupnih tabova — pored nevalidnih kombinacija fiksira i **postojeći pojedinačni režim**, da se meni ne suzi slučajno), `T_EntitetKod` (mapiranje UI labele u kod entiteta — sada jedno mesto istine), `T_ReversKljucPoTipu` (ključ reversa: isti dokument + drugi tip ambalaže **ne** pripada; razmaci i velika/mala slova ne razbijaju poklapanje; prazan tip nije wildcard), i **dva end-to-end testa nad stvarnim tabelama** (u izolovanom test-prozoru, u transakciji koja se **uvek poništava**): `T_E2E_ProsecnaCenaZbirniKupac` — prijemnice dva različita kupca, pojedinačno po kupcu izveštaj vraća redove a zbirno vraća prazno, čime je dokazano da uklanjanje taba nije proizvoljno (test **pada u oba smera**: i ako se prazan tab vrati u meni, i ako neko implementira globalni prosek a tab ostane skriven); `T_E2E_AmbPregledRazdvajaTipDokumenta` — dva ledger reda sa **istim DokumentID-om i istim tipom ambalaže** ali različitim tipom dokumenta (`Otkup` i `OM-Izlaz-Koop`, tačno ono što `SaveOtkup` upisuje) moraju dati dva reda + UKUPNO, sa dva različita skrivena ključa. Freshness indikator, vidljiva greška, tab-meni i rute štampe su **operater-smoke** (UI stanje, ne računska funkcija).
- **Dodirnuti moduli:** `frmIzvestaj` (`UpdateStatusLabel` + novi `ActiveReportList`/`PrikazanPeriod`, `UpdateReportMode` na matricu, `GenerateActivePage` + `ShowGenFailure`, `txtDatumOd_Change`/`txtDatumDo_Change`, `mpReports_Change`, `btnStampaj_Click`, `btnStampajKarticu_Click`, `StampajReversAmbDok` + novi `AmbRedStorniran`, `AutoRefresh` + novi `InvalidateReportContext`/`ClearAllReportLists`/`AktivanTabGenerisan`/`SyncDetaljiVisibility`), `modIzvestaj` (novi seam-ovi `IzvestajTabDostupan`, `IzvestajEntitetKod`, `ReversRedPripada`, `AmbTipKljuc` + `IZV_TAB_*` konstante; `ReportAmbalazePojedinacni` grupni ključ), `modKarticaDetalji` (`KarticaDetalji_CurrentAmbTip`), `modPoruke` (7 novih ključeva), `modIzvestajTests`. Prateći: `docs/KNOWN_ISSUES.md` (AUD-024 i AUD-012 zatvoreni, AUD-027 delimično), `docs/REFAKTOR_PLAYBOOK.md` i `docs/PLAN_SANACIJE.md` (RF-07 / M5 status — posle merge-a), `CLAUDE.md` (§4/§5 — novo pravilo: modul-level `Const` ide u deklaracionu sekciju, VBA ne kompajlira `Const` između procedura).

---

## vba-v2.36.0 — 2026-08-08
> Verzija/datum se **finalizuju pri `tools/release.sh`** (uz `APP_VERSION` u `modConfig`). **RF-08 (M5) — faktura i štampa (AUD-011, ostatak AUD-027).** Poslednji paket M5 milestone-a; time je M5 (izveštaji + faktura) zatvoren. Ne dira poslovnu šemu ni `.frx`.

- **Faktura više ne može da pokupi prijemnicu drugog kupca:** `CreateFaktura` je verovao pozivaocu da su prosleđene prijemnice zaista kupčeve — jedina zaštita je bio filter u formi za fakturisanje, a forma nije sigurnosna granica. Sada se za **svaku** stavku poredi `KupacID` prijemnice sa kupcem fakture; neslaganje zaustavlja ceo posao i transakcija se vraća, pa ne ostaje ni pola upisane fakture. **Vidljiva posledica:** stara prijemnica kojoj je polje kupca prazno (redovi upisani pre nego što se kolona popunjavala) sada biva odbijena uz poruku — takav red se prvo mora dopuniti u tabeli pre fakturisanja.
- **Dvostruki red iste prijemnice više ne prolazi tiho:** ako u tabeli postoje dva aktivna reda sa **istim `PrijemnicaID`**, uzimao se prvi pogodak i fakturisala njegova količina i cena — deterministički pogrešan iznos, bez ijedne poruke. Sada takav slučaj zaustavlja fakturisanje sa jasnom greškom (koji ID i koliko redova), po istom pravilu koje faktura već primenjuje na svoj broj.
- **Faktura se pravi isključivo kroz transakciju:** osnovna funkcija za pravljenje fakture (zaglavlje, stavke i označavanje prijemnica su tri odvojena upisa) više nije dostupna spolja — jedini ulaz je transakciona verzija sa snapshot-om, pa greška na pola posla više ne može da ostavi fakturu bez stavki ili prijemnice označene kao fakturisane bez fakture. Nijedan postojeći poziv nije menjan: forma i test suite su i ranije išli kroz transakcionu verziju.
- **Stornirani otkupni list se više ne štampa:** ponovna štampa iz izveštaja je stornirani otkup **preskakala u filtriranju pa ga ipak štampala** kroz rezervnu granu („bar taj red") — na papiru je izlazio dokument koji je poništen, bez ikakve oznake. Sada se pre svega ostalog proverava sirova tabela i štampa se blokira vidljivom porukom. **Provera je fail-closed:** štampa prolazi tiho samo kad se otkup može dokazati kao jedinstven i aktivan; ako podatak nedostaje (nema kolone `Stornirano`, otkup nije pronađen, ili postoje dva reda sa istim `OtkupID`) — štampa se **takođe zaustavlja**, uz poruku sa razlogom. Ranija verzija je u takvim slučajevima puštala štampu, što je najopasnije baš kod oštećene tabele.
- **Faktura sa više stavki se više ne renderuje kroz zaostali spoj ćelija:** red „UKUPNO:" se spaja na poziciji koja zavisi od **broja stavki**, a čišćenje šablona pre punjenja je brisalo sadržaj i ivice, ali ne i spojene ćelije. Faktura sa 3 stavke odštampana posle fakture sa 1 stavkom je time pisala preko spoja iz prethodne (pomeren/izgubljen sadržaj reda). Sada se opseg stavki razdvaja pre punjenja, i to **na opsegu koji prati stvaran broj stavki i stvaran sadržaj lista**, ne na fiksnih 80 redova kao ranije: faktura sa preko 80 stavki je inače nastavljala da lomi isti scenario (81 → 82 stavke: štampa bi tiho izostala, a broj prijemnice tipa `1/2026` od 81. stavke Excel bi mogao da protumači kao datum). Broj stavki nije ograničen — bira se koliko prijemnica ima. Zaglavlje šablona (kupac, podaci prodavca, naslov) ima svoje namerne spojeve i **ostaje netaknuto** — razdvaja se samo tabela stavki.
- **Rizik za podatke:** **nema promene šeme**; poslovne tabele se ne diraju (izmene su u validaciji pre upisa i u pripremi lista za štampu); `.frx` netaknut. VBA izvor **ASCII-only**. **Dodata su 2 nova `Poruka()` ključa (`PRINT_ERR_STORNIRAN_OTKUP`, `PRINT_ERR_OTKUP_BLOKIRAN`) → posle importa pokrenuti `EnsurePoruke`** (bez toga poruke prikazuju `[KLJUC]`). Ponašanje je strože nego pre: fakturisanje koje je ranije tiho prolazilo (prijemnica bez kupca ili sa drugim kupcem, duplirani `PrijemnicaID`) sada završava kao vidljiva greška i traži reakciju operatera — to je i svrha izmene.
- **Testovi:** `RunFakturaSmokeSuite` je dobio **tvrd gate** (podiže grešku kad ijedna provera padne, kao `RunIzvestajTests`; ranije je samo prikazivao zbir) i četiri nova testa: prijemnica tuđeg kupca mora biti odbijena i ostati neoznačena kao fakturisana; dva aktivna reda sa istim `PrijemnicaID` moraju zaustaviti fakturisanje; štampa mora biti odbijena za stornirani, duplirani i nepoznat otkup a dozvoljena za aktivni (redovi u `tblOtkup` se seed-uju u transakciji koja se **uvek poništava**); i faktura sa **82 stavke odštampana posle fakture sa 81 stavkom** mora da se renderuje čisto (82. red nije u zaostalom spoju, stvarno je upisan, broj prijemnice mu je i dalje tekst, zaglavlje netaknuto), uz proveru da čišćenje ostaje ograničeno na stvaran sadržaj — prazna formatirana ćelija hiljadama redova ispod fakture ne sme da ga razvuče (inače bi svaka štampa obrađivala milione ćelija i praktično zamrzla Excel). Provera storna se radi nad izdvojenom funkcijom, jer sama blokada otvara poruku operateru. Postojeći testovi fakture su dopunjeni kupcem u test prijemnici (pomoćna funkcija sada zahteva `KupacID`). **Merge test malog obima ostaje operater-smoke:** faktura sa 3 stavke odštampana posle fakture sa 1 stavkom. Napomena: taj test na kraju **briše pomoćni list `FakturaSablon`** (generisani šablon, ne podaci) da ne ostavi test stavke za sobom — prva sledeća štampa fakture ga automatski napravi ponovo.
- **Dodirnuti moduli:** `modFaktura` (`CreateFaktura` — vlasništvo prijemnice, `Count=1` guard, prelazak na `Private`), `modPrint` (`ReprintOtkupniListByOtkupID` + nova fail-closed kapija `RequireOtkupAktivanZaStampu`, `FillFakturaSablon` dinamičan cleanup + `.UnMerge`, novi `SablonLastContentRow`), `modPoruke` (2 nova ključa), `modFakturaTests`. Prateći: `docs/KNOWN_ISSUES.md` (AUD-011 i AUD-027 zatvoreni), `docs/REFAKTOR_PLAYBOOK.md` i `docs/PLAN_SANACIJE.md` (RF-08 / M5 status — posle merge-a).

---

## vba-v2.37.0 — 2026-08-08
> Verzija/datum se **finalizuju pri `tools/release.sh`** (uz `APP_VERSION` u `modConfig`). **RF-22 (M6) — SEF UX i lifecycle (AUD-032).** Prvi paket M6 milestone-a. Ne dira poslovnu šemu, `.frx` ni SEF submit/validate/JSON logiku (to je bio RF-21).
>
> **Suština:** SEF ekran je na više mesta prijavljivao uspeh kojeg nije bilo — odbijena faktura je izgledala kao poslata, prazan status kao „poslato", pali refresh kao osvežen, a zaglavljena faktura kao oporavljena. Ova verzija svaku od tih poruka vezuje za stvarni ishod.

- **Poruka posle slanja govori šta se stvarno desilo.** Do sada je posle svakog slanja pisalo „Faktura poslata", i kad je SEF fakturu **odbio** (REJECTED) i kad je slanje **tehnički palo** (TECH_FAILED). Sada se poruka bira po stvarnom stanju: poslata / prihvaćena / odbijena (uz uputstvo da se proveri „Poslednja greška" i pripremi ponovno slanje) / tehnička greška. Stanje fakture se u svim slučajevima čuva pre poruke, kao i do sada.
- **Prazan ili nepoznat status sa SEF-a više se ne upisuje kao „poslato".** Kad SEF na proveru statusa ne vrati status, ranije se to tiho upisivalo kao `SENT` — nepotvrđena faktura je izgledala kao uredno predata. Sada se upisuje `UNKNOWN_STATUS`, faktura ide na **ručnu proveru** (žuto u pregledu, upozorenje u monitoringu), a lokalno stanje se ne pomera napred. Faktura zaglavljena u slanju prelazi u stanje `SEF_UNKNOWN` iz kog operater može da ponovi „Osveži status" — ranije je to stanje postojalo, ali iz njega nije bilo izlaza.
- **Statusi se čitaju po zvaničnom SEF spisku.** SEF prihvaćenu fakturu zove **`Approved`**, a program je poznavao samo `Accepted`; nije poznavao ni `Seen`, `Sending`, `Paid`, `OverDue`, `Archived`, `Mistake`, `Deleted`. Sada postoji jedan spisak koji svaki zvanični status prevodi u značenje (odobreno / odbijeno / u obradi / poništeno / informativno / greška slanja / nepoznato), i koriste ga prikaz, boja u formi, batch obrada i sve provere — pa ne mogu da se raziđu. **Posledica koju ćete videti:** kod prihvaćene fakture u polju statusa sada piše `APPROVED`.
- **Faktura više ne ostaje zaglavljena u stanju „šalje se".** Tri slučaja su je tamo držala: (a) SEF javi da je dokument storniran/otkazan — faktura je ostajala „u slanju" i pri **svakom pokretanju** programa upisivala lažan zapis o oporavku; (b) status je plaćeno/dospelo/arhivirano — što ne govori da je kupac odobrio fakturu, ali dokazuje da dokument više nije u slanju; (c) provera statusa padne — ranije je pokušavala prelaz koji sam program ne dozvoljava, pa je pucala i vraćala izmene. Sva tri sada izvode fakturu iz „šalje se", bez proglašavanja prihvaćenom.
- **Neuspelo slanje se ne prikazuje kao poslato.** SEF status `Mistake` znači **grešku prilikom slanja dokumenta**; program ga je svrstavao među poništene, pa je faktura kojoj slanje nije uspelo lokalno postajala „poslata", batch bi je preskakao, a otkazivanje i ponovno slanje bili bi nedostupni. Sada prelazi u stanje tehničke greške. **Putanja za takvu fakturu je „Otkaži slanje na SEF" + ručna provera, a ne ponovno slanje** — da li SEF prihvata ponovnu predaju istog dokumenta nije moguće utvrditi iz koda, a slanje fakture je pravni čin, pa se ne pretpostavlja.
- **„Osveži status", „Recover sending" i batch akcije govore istinu.** Prve dve su ranije uvek javljale uspeh, i kad je poziv ka SEF-u pao; sada prikazuju stvaran rezultat, a „Recovery završen" samo ako je faktura zaista izašla iz stanja slanja. „Osveži sve Pending" i „Recover sve sending" daju sažetak (pregledano / osveženo / nerazrešeno / preskočeno / palo) umesto jedne fiksne poruke; brojači više ne računaju „nije puklo" kao uspeh.
- **Dugmad prate isti spisak kao provera koja ih propušta.** Forma i provera su imale odvojene liste statusa, pa je forma mogla da ponudi akciju koju provera odbija. Sada „Otkaži slanje", „Storniraj" i „Pošalji" koriste isti spisak: `Mistake` je dobio otkazivanje, `APPROVED` storniranje, a **dugme za slanje se gasi kad dokument na SEF-u postoji** — uključujući posle uspešnog otkazivanja i posle pada mreže pri proveri statusa (odluka se vezuje za broj dokumenta na SEF-u, podatak koji pad mreže ne briše). Kad je slanje blokirano, poruka nudi otkazivanje, storniranje ili proveru na portalu — u zavisnosti od toga šta je za taj status zaista dozvoljeno.
- **Odbijena faktura sada stvarno može ponovo da se pošalje.** Kad SEF fakturu odbije tek pri proveri statusa (a ne odmah pri slanju), „Pripremi za ponovno slanje" je vraćalo fakturu u pripremljeno stanje, ali bi samo slanje palo uz poruku o duplikatu — jer je zapis o prethodnoj predaji ostajao označen kao uspešan. Sada priprema razdužuje i taj zapis, u istoj transakciji. Ako fakturu i dalje blokira neka ranija uspešna predaja, priprema **staje sa jasnom porukom** umesto da ostavi fakturu koju bi slanje odbilo.
- **Promena izabrane fakture briše prikazan SEF kontekst.** Biranjem druge fakture u padajućoj listi status, SEF broj dokumenta i event log prethodne fakture više ne ostaju na ekranu (ranije su stajali dok se ne klikne „Učitaj fakturu", pa se tuđi status mogao pročitati kao status izabrane).
- **Opasni test makroi više nisu u Alt+F8 listi.** `Test_CancelInvoiceOnSEF_TX` i `Test_StornoInvoiceOnSEF_TX` su otkazivali/stornirali stvarnu fakturu na SEF-u (pravni čin) nad hardkodovanim brojem fakture i bez ijedne potvrde, a stajali su u listi makroa uz obične alatke. Uklonjeni su — isti scenariji već postoje u test modulu iza tri kapije (dozvola za žive testove, posebna dozvola za cancel/storno, tipkanje potvrde). Ostali razvojni makroi koji zovu živi SEF prebačeni su u `Private`, pa ih nema u listi, ali se i dalje pokreću iz VBE-a.
- **Rizik za podatke:** **nema promene šeme**; poslovne tabele se ne diraju izvan postojećih SEF kolona; `.frx` netaknut; nema novih `WithEvents` deklaracija u formi. VBA izvor **ASCII-only**. **Dodato je 7 novih `Poruka()` ključeva → posle importa obavezno pokrenuti `EnsurePoruke`** (bez toga poruke prikazuju `[KLJUC]`). **Ponašanje je namerno strože nego pre:** situacije koje su ranije prolazile kao uspeh (odbijena faktura, prazan status, pali refresh, greška slanja) sada traže reakciju operatera.
- **Testovi:** nov **`RunSEFTestSuite`** — ciljani suite za SEF milestone, sa **tvrdim gate-om** (podiže grešku kad ijedna provera padne). Ne poziva pravi SEF. Pokriva: klasifikaciju svakog statusa iz zvaničnog SEF spiska; matricu **12 lokalnih stanja × 9 klasa statusa** (svaki predlog prelaza mora biti dozvoljen ili nikakav); ishod slanja; ugovor recovery-ja; šta se sme raditi nad kojim statusom (otkazivanje / storniranje / ponovno slanje). **Dva** testa seed-uju redove u SEF tabele i **uvek ih poništavaju** (uz ugašen journal i AutoSave, da ne ostave trag): odbijena faktura posle pripreme stvarno prolazi proveru duplikata, i uspešan storno stvarno pomera lokalno stanje fakture. Suite je bio referenciran u planu sanacije kao gate za SEF milestone, ali do sada **nije postojao**. Živi SEF testovi ostaju iza postojećih kapija. Uz suite ide i `tools/check-sef-asserts.py` — statička provera koja **čita produkcioni izvor** (klasifikator, ciljna stanja po klasi, tabela dozvoljenih prelaza, konstante) i javlja ako neki test tvrdi nešto što kod više ne radi. CI ne pokreće Excel, pa se takva neusklađenost inače vidi tek kad operater pokrene suite. Provera je **parcijalno** izvedena iz izvora — dva pravila planera su u njoj ručno preslikana i to je zapisano u samoj skripti.
- **Stanje verifikacije:** `RunSEFTestSuite` je pokrenut na operaterskoj mašini posle `ImportAllVBA` + `Compile` — **prošao all green** (`Failed=0`). `tools/check-sef-asserts.py` prijavljuje 0 neslaganja assert-a sa produkcionim izvorom. **Operaterski smoke (slanje na demo SEF, storno, recovery posle restarta) još nije izvršen** — koraci A2–E2 u `docs/SEF_LIFECYCLE_MANUAL.md` §8; najvažniji preostali su D4 (storno pomera lokalni workflow u `SEF_STORNO`) i D7 (dupli restart ne nalazi istu fakturu opet).
- **Uputstvo:** novo `docs/SEF_LIFECYCLE_MANUAL.md` — šta koja poruka i koje dugme sada znače (dva stanja, tabela SEF status → klasa → lokalno stanje, kada je koje dugme aktivno, tipične situacije) + kompletna test-lista za ovu verziju. Incident runbook (`docs/production-runbook-sef-slanje-faktura.md`) ostaje nepromenjen za dijagnostiku „ne mogu da pošaljem fakturu".
- **Dodirnuti moduli:** `modSEFService` (ishod slanja, recovery koji vraća rezultat, batch sažetak, dev makroi), `modSEFStatusSync` (adapter statusa, planer prelaza, stvaran rezultat refresh-a, batch sažetak), `modSEFClient` (`ParseStatusResponse` + test proxy), `modSEFValidator` (prelazi, capability spiskovi za cancel/storno/slanje, priprema resubmita), `modSEFPersistance` (razduženje submisije), `frmSEF` (poruke po ishodu, reset pri promeni fakture, dugmad po capability spisku), `modConfig` (11 novih konstanti), `modPoruke` (7 novih ključeva), `modSEFTests`. Prateći: `docs/KNOWN_ISSUES.md` (AUD-032 zatvoren; jedan AUD-054 site usput popravljen), `docs/REFAKTOR_PLAYBOOK.md`, `docs/PLAN_SANACIJE.md`, `docs/ARCHITECTURE_CHANGELOG.md`, novo `docs/SEF_LIFECYCLE_MANUAL.md`.

---

## vba-v2.38.0 — 2026-08-09
> Verzija/datum se **finalizuju pri `tools/release.sh`** (uz `APP_VERSION` u `modConfig`). **RF-09 (M6) — banka: uvoz izvoda i mapiranje u novac.** Zatvara `AUD-014` + `AUD-025`, i preuzima **P0 `AUD-007`** (datum) koji nije bio zaveden ni na jednom paketu. Drugi paket M6 milestone-a (posle RF-22). Ne dira banka **naloge/isplate** (`frmBankaExportPregled` — to je RF-10), parsere po banci, ni `.frx`.
>
> **Suština:** uvoz i knjiženje izvoda su na više mesta radili nešto što operater ne vidi — datum je mogao da uđe pomeren, transakcija sa drugog računa da nestane kao „duplikat", jedan problematičan blok da poništi ceo prolaz, a samo otvaranje ekrana da proknjiži novac. Ova verzija svaku takvu situaciju ili odbija glasno, ili stavlja pod klik uz prikaz onoga što će se stvarno knjižiti.

**Datum (P0)**

- **Nemoguć datum se više ne uvozi pomeren.** `DateSerial` u Excelu ne odbija `30.02.2026` nego ga tiho pretvori u `02.03.2026` (isto važi za dan 32 i mesec 13), pa je datum iz izvoda mogao da završi u pogrešnom mesecu bez ijedne poruke. Provera datuma sada radi round-trip — što je uneto mora i da izađe — i takav datum odbija.
- **Datum se ne tumači po podešavanjima računara.** Format `dd.mm.yyyy` (koji daju svi parseri banaka) čita se deterministički — dan pa mesec — pre nego što se uopšte pokuša Windows/Excel tumačenje. Na računaru podešenom na `MM/DD` je `01.02.2026` ranije postajalo **2. januar** umesto 1. februara. Uz to, vrednost tog oblika koju provera odbije **više se ne „spasava"** Windows tumačenjem: `01.13.2026` (mesec 13) je VBA na en-US mašini prihvatao tako što zameni dan i mesec i da 13. januar — sada je odbijena. Isto pravilo znači da se dotted-ISO oblik (`2026.02.01`) ne prihvata; podržani su `d.m.gggg`, `d/m/gggg` i `gggg-mm-dd`.
- **Jednom pročitan datum ostaje datum.** Upisuje se u staging kao pravi datum, a ne kao tekst koji se pri svakom čitanju iznova tumači. Uvoz uz to ima novu, **nultu** kapiju: datum izvoda i datum **svake** transakcije provere se pre nego što ijedan red uđe u staging; ako neki nije stvaran datum (ili je van poslovnog opsega godina), uvoz tog PDF-a staje sa jasnom porukom — koji red i koja vrednost — a fajl ide u `Error` folder. Upis batch-a je atomičan: ako je peta transakcija neispravna, ni prve četiri ne ostaju.

**Novac i identitet**

- **Transakcija sa drugog računa firme se više ne gubi.** Provera duplikata je gledala broj izvoda, datum, iznos, partnera i referencu — ali **ne i broj računa**. Kako je broj izvoda jedinstven po računu (a ne globalno), ista suma istog dana na drugom računu je odbacivana kao „već uvezeno" i tiho nestajala.
- **Uplata veća od duga po fakturi se deli.** Ako je na fakturi otvoreno 1.000, a iz banke stigne 1.500, ranije je svih 1.500 knjiženo na fakturu — faktura preplaćena, a avans nigde. Sada ide 1.000 na fakturu i 500 kao avans kupca, u istoj transakciji; stavka izvoda se zatvara tek kada je ceo iznos proknjižen. Ako je faktura u međuvremenu plaćena, mapiranje se odbija uz poruku (osveži listu).
- **Ručno mapiranje kupca može da zatvori fakturu.** Polje za fakturu je postojalo na formi ali se nikad nije punilo, pa je svaka ručna uplata kupca završavala kao **avans** — i onda kad je faktura postojala i bila prikazana u pregledu. Sada se nudi lista otvorenih (nestorniranih) faktura tog kupca sa preostalim iznosom; bez izabrane fakture program traži potvrdu i jasno kaže da se knjiži kao avans. Veza faktura↔kupac se proverava i pri upisu (tuđa ili stornirana faktura se odbija).
- **Kupac i OM se biraju po identitetu, ne po imenu.** Padajuća lista je spajala partnere istog naziva u jednu stavku i uzimala prvi pogodak — dva kupca istog imena operater nije mogao ni da razlikuje, a uplata je mogla da završi na pogrešnom. Sada izbor nosi `KupacID`/`StanicaID` i **prikazuje ga uz naziv** (`Agro Trade  [KUP-284]`).
- **Pogrešan smer se odbija.** Izbor tipa (Kupac / Kooperant / OM) nije bio vezan za smer stavke, pa je isplata mogla ručno da se proknjiži kao uplata kupca (i obrnuto). Sada: Kupac traži uplatu, Kooperant isplatu, OM prihvata oba smera ali ne i „nečist" red (i uplata i isplata, ili nijedno). Odbijanje je vidljivo, a ne kao ranije — tiho „ništa se nije desilo".

**Blok sa 3+ otvorenih stavki**

- **Nijedan problematičan red ne ruši ceo „Automatski mapiraj sve".** Isto važi i kada se red automatski poveže sa fakturom koja je **već plaćena** — takav red ide „za ručno", a ostatak prolaza se knjiži. Greška koja nastane **posle** prvog upisa (neproknjižen avansni deo) i dalje namerno obara i vraća ceo prolaz.
- **Jedan problematičan blok više ne ruši ceo „Automatski mapiraj sve".** Ako je otkupni blok imao **tri ili više** otvorenih stavki, program bi pukao na internoj grešci i **poništio ceo prolaz** — i sve redove koje je do tada uredno mapirao. Sada takav red dobija jasnu poruku, označi se za ručno mapiranje, a batch nastavlja; poruka na kraju kaže i koliko redova traži ručno. Svaka **druga** greška i dalje zaustavlja i vraća ceo prolaz — to se namerno nije popuštalo.
- **Takav red se može i završiti — i bez biranja bloka iz liste.** Lista blokova se puni ali se ne selektuje sama, pa je „ništa izabrano" najčešći slučaj; ranije je baš tu ručno knjiženje završavalo generičkom greškom umesto ponuđenom podelom. Sada ručno mapiranje uvek koristi isti blok koji pregled prikazuje (izbor iz liste ako postoji, inače poziv na broj). Ručno mapiranje po bloku ide preko granice **uz izričitu potvrdu**: program prikaže tačnu podelu iznosa po otkupima (veći otvoreni prvi, ostatak u avans) i pita „Knjižiti ovako?". Uz to postoji i **„knjiži ceo iznos kao avans kooperanta"** — kada je poreklo dvosmisleno (recikliran broj bloka, dupliran unos), ništa se ne vezuje za otkup, a avans se kasnije precizno primenjuje dugmetom „Primeni avans na blok".

**Ekran „Banka uvoz izvoda"**

- **Otvaranje ekrana više ne knjiži novac.** Do sada je samo otvaranje pregleda pokretalo automatsko mapiranje po jakim ključevima: redovi u novcu su nastajali bez potvrde i bez poruke, a ako bi prolaz pukao — greška je bila progutana i sve vraćeno, pa je izgledalo kao da nije ni bilo šta da se mapira. Sada se pri otvaranju **samo prebroji** šta bi se mapiralo (bez ijednog upisa), broj se vidi u statusu i na novom dugmetu **„Mapiraj jake ključeve (N)"**, a knjiženje ide na klik uz potvrdu i prikaz rezultata.
- **Pregled i dugme govore isto — i to svako za svoju komandu.** Pregled je kandidate računao po pozivu na broj **iz izvoda**, dok „Ručno mapiraj red" knjiži po bloku izabranom u padajućoj listi. Sada ispod automatskog pregleda stoji sekcija **„RUČNO"** koja pokazuje šta bi tačno uradilo to dugme (tip, partner, blok, faktura, tip knjiženja, uključujući „ODBIJENO" i razlog). Automatski deo pregleda i dalje prati **poziv na broj**, jer „Automatski mapiraj red" tu listu uopšte ne gleda — ručni izbor bloka menja samo ručni deo.
- **Rezultat akcije se ne ignoriše.** Posle ručnog mapiranja dobija se poruka je li knjiženo ili nije. Ako lista faktura ne može da se učita, mapiranje kupca **staje** uz stvarnu grešku — prazan spisak posle greške izgleda isto kao „nema otvorenih faktura", a to bi značilo avans.

**Kad nešto pukne, to se vidi**

- **Kad uvoz padne, vidi se ZAŠTO.** Na vrhu lanca uvoza poruka o grešci je mogla da ostane bez uzroka, jer se uzrok čitao tek pošto ga je logovanje već obrisalo. Sada se prenosi originalni broj i opis greške — uvoz je i ranije korektno vraćao promene, ali operater nije video razlog.
- **Neuspelo „Automatski mapiraj sve" ne izgleda kao uspeh.** Ako batch padne i sve se vrati, ranije je posle crvene poruke stizala i poruka „Automatski mapirano: 0" — što se čita kao uredno završen prolaz bez pogodaka. Sada je jasno rečeno da mapiranje **nije izvršeno** i da su promene vraćene.
- **Knjiženje se ne može pozvati bez transakcije.** Od ove verzije jedan poziv mapiranja može da napravi više redova u novcu (uplata na fakturu + avans; raspodela po bloku), pa su „goli" upisivači zatvoreni, a spolja su dostupni samo transakcioni ulazi — pad drugog upisa vraća i prvi, umesto da ostane pola knjiženja uz otvorenu stavku izvoda.

**Rizik za podatke i verifikacija**

- **Rizik za podatke:** **nema promene šeme**; `.frx` netaknut; nema novih `WithEvents` deklaracija u formi (novo dugme je runtime kontrola preko `clsUiSink`); VBA izvor **ASCII-only**; **nema novih `Poruka()` ključeva** (`EnsurePoruke` nije potreban za ovu verziju). **Ponašanje je namerno strože nego pre:** izvod sa nevalidnim datumom se odbija u celosti, a pogrešan smer i blok sa 3+ otvorenih stavki se odbijaju umesto da prođu ili puknu.
- **Testovi:** nov **`RunBankaImportTestSuite`** (`modTestBanka`, Alt+F8) — **11 grupa provera** sa tvrdim gate-om (podiže grešku kada ijedna padne, ne samo crveni `MsgBox`). Pokriva: nemoguće i locale datume (`01.02.2026` = 1. februar); multi-account dedupe; blok sa 3+ stavki (batch ne pada, zdrav red ostaje mapiran, anomalan bez parcijalnog knjiženja, pa se ručnom putanjom završi tako da zbir knjiženja odgovara iznosu iz izvoda); pogrešan smer; typed staging + atomičan mešovit batch; podelu preplate (1.500 na dug 1.000 → 1.000 + 500 avans) i odbijanje već plaćene fakture; istoimene partnere po stabilnom ID-u; batch koji padne (diže grešku i vraća i prethodno uspešan red); to da automatsko knjiženje pogađa blok iz poziva na broj, a ne drugi blok istog kooperanta; već plaćenu fakturu u „Auto sve" (samo taj red ide „za ručno", zdrav red ostaje knjižen, faktura se ne preplaćuje); i ručno mapiranje kooperanta **bez izbora bloka** (prolazi kroz potvrđenu podelu i knjiži tačan iznos). Svi podaci (prefiks `BIT-`) prave se u transakciji koja se **uvek** poništava, uz utišan journal. Postojeći `Test_BankParse` ostaje ručna provera parsera nad stvarnim PDF-om.
- **Statička kapija:** nov `tools/check-banka-eh.py` (ne traži Excel) — čuva da se ne vrate obrasci u kojima greška nestaje: čitanje `Err` posle `LogErr`, `Err.Raise` nad živim `Err`, i `Err.Raise` pod aktivnim `On Error Resume Next` (takav raise se tiho guta). Pri uvođenju je odmah našao tri stvarna mesta; širok skan celog `src-vba/` pokazuje da drugih progutanih `Err.Raise` u projektu nema.
- **Stanje verifikacije:** statički — ASCII, balans `Sub`/`Function`/`If`/`Select Case`, nema modul-level deklaracija posle prve procedure, nema duplih `Public` definicija, `git merge-tree` i `git diff --check` čisti, `check-banka-eh.py` čist. **`ImportAllVBA` + `Compile` + `RunBankaImportTestSuite` pokrenuti su na operaterskoj mašini → `PASS=110 FAIL=0`, svih 11 grupa zeleno.** Put do toga je našao dva stvarna nalaza: prvi run (`PASS=98 FAIL=1`) je oborio **produkcioni bug** — `CDate` fallback je prihvatao `01.13.2026` kao 13. januar (zamena dana i meseca), pa je taj put zatvoren za `d.m.g` oblik; drugi run (`PASS=106 FAIL=4`) je oborio **curenje između test grupa**, ne produkciju — red bez broja izvoda, koji jedna grupa namerno ostavlja da obori batch, ulazio je u batch sledeće grupe (izolacija rešena whitelist-om otvorenih redova). **Operaterski smoke `frmBankaImport` + uvoz stvarnog PDF-a još nisu izvršeni** — koraci su u test-listi uz PR.
- **Dodirnuti moduli:** `modParse` (`TryParseBankaDateDMY`, `TryParseDateValue`, `IsPoslovnaGodina`), `modBankaImport` (Level 0 kapija datuma, `RequireBimDatum`, `SaveBankaImportRows_TX` + private core, dedupe sa brojem računa, EH na import lancu), `modBankaMapiranje` (`RequireBimSmer`/`ClassifyBimSmer`, `RequireFakturaZaKupca`, `GetOtvorenoNaFakturi`, `GetOtkupCandidatesForKooperantBlock` + `PlanBlokRaspodela`, `AutoBlockNoForBim`, `AutoMapBankaImportRowBatch`, `CountStrongKeyReadyBankaImport`, TX-only javni API), `frmBankaImport` (runtime dugme, `cmbFaktura`, `EffectiveManualBlockNo`, sekcija „RUČNO", EH grane), `modComboBinding` (`ComboDisplayWithID`/`ShowIDInComboDisplay`), nov `modTestBanka`, `modTestStorno` (dva poziva dedupe-a usklađena sa novim potpisom). Prateći: `docs/KNOWN_ISSUES.md` (AUD-007/014/025 zatvoreni, KI-BANKA-DEV zaveden kao follow-up), `docs/REFAKTOR_PLAYBOOK.md`, `docs/PLAN_SANACIJE.md`, `docs/ARCHITECTURE_CHANGELOG.md` (v6.47), `docs/production-runbook-banka-import-setup.md`, `docs/development-banka-parser.md`, `CLAUDE.md`, nov `tools/check-banka-eh.py`.

---

## vba-v2.39.0 — 2026-08-11
> Verzija/datum se **finalizuju pri `tools/release.sh`** (uz `APP_VERSION` u `modConfig`). **RF-10 (M6) — banka: nalozi za isplatu.** Zatvara `AUD-026`. **Treći i poslednji paket M6 milestone-a** (posle RF-22 i RF-09) — time je M6 kompletan. Menja ekran „Banka — nalozi za isplatu" (`frmBankaExportPregled` + `modBankaExportPregled`) i jednu proveru u knjiženju avansa (`modNovac`). Ne dira uvoz izvoda (RF-09), format i šifre naloga, ni `.frx`.
>
> **Suština:** ekran koji priprema naloge za banku mogao je, na dva različita načina, da napravi nalog koji nije tačan — da naruči **veći iznos** nego što je dug, ili da novac pošalje **pogrešnoj osobi**. Nijedan od ta dva slučaja nije bio vidljiv operateru: nalog izgleda uredno, iznos i broj bloka su „normalni". Ova verzija zatvara oba, po istom pravilu: **kada iznos ili primalac ne mogu da se dokažu, fajl se ne pravi.**

**Novac ide tačnoj osobi**

- **Nalog više ne može otići pogrešnom kooperantu.** Ovo je najvažnija ispravka u verziji. Ime i tekući račun primaoca su se nalazili tako što se u tabeli otkupa uzme **prvi** red sa datom šifrom otkupa (`OtkupID`). Ako su dva otkupa delila istu šifru, a prvi je već bio isplaćen — pa se u pregledu uopšte ne vidi — otvoreni blok drugog kooperanta je dobijao **račun onog prvog**. Iznos i poziv na broj su pritom tačni, pa se na nalogu ne vidi ništa sumnjivo. Sada se vlasnik prihvata samo kada je šifra otkupa jednoznačna; u suprotnom pregled i generisanje staju uz poruku koja imenuje spornu šifru.
- **Ni dvojnik u matičnim podacima ne bira račun.** Isto važi korak dalje: ako u kooperantima postoje **dva reda sa istom šifrom kooperanta** a različitim tekućim računom, program više ne uzima onaj koji prvi naiđe. Ovo postojeća provera integriteta podataka nije hvatala — ona gleda šifre otkupa, otpremnica, faktura i slično, ali ne i kooperante.
- **„Primeni avans" se ne izvršava nad nejednoznačnim otkupom.** Ista zaštita važi i za drugu akciju ovog ekrana: ako dva otkupa dele istu šifru, klik na jedan blok je mogao da se izvrši nad drugim. Sada se akcija odbija, ne knjiži se ništa i avans ostaje nevezan. Provera je u samom knjiženju, ne u ekranu, pa važi i za svaki drugi poziv.
- Dvojnik koji ne stoji iza nijednog otvorenog bloka (npr. istorijski, potpuno zatvoren) **ne smeta radu** — iz njega nalog ionako ne može nastati.

**Nalog ne može da naruči više nego što je otvoreno**

- **Zaostao „Isplatiti" iznos se sam usklađuje.** Ručno unet iznos po bloku ostaje zapamćen kroz osvežavanje liste — to je namerno, da se unos ne gubi. Ali se do sada brisao samo za blok koji je iz liste potpuno nestao, dok se sam iznos nije poredio ni sa čim. Sada se pri svakom osvežavanju iznos veći od otvorenog **spušta na otvoreno**, unos za zatvoren ili storniran blok se **briše**, a iznos **manji** od otvorenog ostaje netaknut — delimična isplata je vaša odluka, ne greška. Koliko je unosa dirnuto piše u statusnoj liniji, da spuštanje ne prođe neprimećeno.
- **Pred generisanje CSV-a se saldo proverava iznova.** Između trenutka kad ste videli listu i klika na „Generiši" stanje se moglo promeniti — uvezen je izvod, vezan avans, nešto stornirano u drugom delu programa. Zato se pre pisanja fajla čita **sveže** stanje i svaki nalog poredi sa trenutno otvorenim iznosom. Ako ijedan traži više, **nijedan nalog se ne generiše**: dobijate poruku koja imenuje blok i kaže koliko traži naspram koliko je otvoreno, a pregled se odmah osvežava. Blok koji u međuvremenu više nije otvoren računa se kao **otvoreno = 0** — namerno stroža strana.
- **Preplata od jednog centa se više ne provlači.** Iznosi se porede zaokruženi na dinarske pare, bez ikakve tolerancije. Ranija provera je propuštala tačno jedan cent viška (`600,01` na dug od `600,00`), a to je iznos koji banka stvarno isplati.
- **Nema naloga na 0,00 dinara.** Otvoreni iznos se računa kao količina × cena − isplaćeno, pa u njemu može ostati sitan ostatak reda nekoliko hiljaditih dinara. Takav blok je prolazio proveru „iznos veći od nule", a u fajl bi otišao kao nalog na `0,00` — banka takav red odbija i sa njim može odbiti **ceo** uvezeni paket. Sada se zaokružuje pre te provere, pa blok sa sub-cent ostatkom jednostavno ne dobija nalog, a ostali prolaze normalno.

**Ekran ne tvrdi više nego što je proverio**

- **Otvoren dug ne može tiho da nestane sa spiska.** Ako otkupu nedostaje šifra ili kooperant, program ga do sada jednostavno **nije prikazivao** — obaveza od npr. 100.000 dinara bi izostala iz pregleda bez ijedne poruke, a ostali nalozi bi se uredno generisali. Pošto se tada ne može utvrditi ni obaveza ni primalac, pregled se zaustavlja uz poruku koja imenuje sporan dokument.
- **Neuspelo osvežavanje briše listu.** Ako učitavanje padne, prethodni spisak blokova ne ostaje na ekranu — inače bi izgledao kao proveren, a upravo je provera pukla.
- **Razlog greške se vidi.** Poruke o grešci su čitale opis **posle** logovanja, koje ga može obrisati — pa je operater mogao dobiti praznu poruku umesto razloga. Opis se sada hvata pre logovanja.

**Primena avansa prijavljuje ono što je stvarno knjiženo**

- **Broj „primenjenih avansa" više nije naduvan.** Vezivanje avansa uspeva i kada nije bilo šta da se veže (kooperant nema slobodan avans, blok više nije otvoren) — transakcija prođe, ali se ne knjiži ni dinar. Do sada je batch „Primeni avans (sel.)" i takav ishod brojao kao primenjen avans. Sada se broji **stvarno proknjižen iznos**: rezultat razdvaja „primenjeno" (uz zbir u dinarima), „bez promene" i „greška". Isto važi za dugme po bloku — kad ništa nije knjiženo, program to kaže umesto da javi uspeh.

**Rizik za podatke i verifikacija**

- **Rizik za podatke:** **nema promene šeme**; `.frx` netaknut; nema novih `WithEvents` deklaracija u formi; VBA izvor **ASCII-only**; **nema novih `Poruka()` ključeva** (`EnsurePoruke` nije potreban za ovu verziju). Ekran i dalje **ne piše u `tblNovac`** — isplata se knjiži tek uvozom izvoda; jedini upis je postojeće vezivanje avansa.
- **Ponašanje je namerno strože nego pre.** Tamo gde se iznos ili primalac ne mogu dokazati, program staje umesto da pogađa. Ako u podacima postoji oštećena šifra otkupa ili dvojnik kooperanta iza otvorenog bloka, **ekran isplata će prijaviti grešku i neće raditi dok se to ne sredi** — prvi korak je `RunProductionHealthCheck`, uz napomenu da on duplikate kooperanata ne pokriva, pa se oni traže po poruci koja imenuje spornu šifru.
- **Testovi:** `modTestBanka` (Alt+F8 → `RunBankaImportTestSuite`) proširen sa uvoza na celu banku — **9 novih grupa, ukupno 20**, sa tvrdim gate-om (podiže grešku kad ijedna provera padne). Pokrivaju: usklađivanje zaostalog iznosa; odbijanje preplate sa granicama na cent (`600,00` prolazi, `600,01` pada, `600,006` → `600,01` pada, `600,004` → `600,00` prolazi; za vrednost tačno na pola pare tvrdi se samo da iznos u fajlu nikad ne prelazi otvoreno, jer smer zaokruživanja tu zavisi od binarne reprezentacije); sadržaj samog CSV-a (kapija je unutar funkcije koja gradi fajl, pa preplata daje prazan sadržaj); sub-cent ostatak koji ne sme da napravi red `0,00`; oštećenu šifru otkupa (dupla, prazna, i **sakrivena** — jedan red isplaćen, drugi otvoren na drugog kooperanta sa drugim računom); blok bez kooperanta; dvojnika u kooperantima sa dva različita računa; i primenu avansa nad nejednoznačnim otkupom. Svaka grupa koja dira tabele radi u **izolovanoj transakciji koja se uvek poništava**, pa ne zavisi od redosleda pokretanja i ne ostavlja podatke.
- **Stanje verifikacije:** statički — ASCII, balans `Sub`/`Function`/`Select Case`, nema modul-level deklaracija posle prve procedure, nema duplih `Public` definicija, arnost svih test seed-ova proverena mehanički, `git merge-tree` vs `main` čist. Granice zaokruživanja proverene i van Excela za sva četiri slučaja. **`ImportAllVBA` + `Compile` + `RunBankaImportTestSuite` pokrenuti su na operaterskoj mašini.** Prvi run: `PASS=186 FAIL=1` — pao je jedan **test vektor**, ne produkcija: `600,005` se u `Double`-u čuva ispod pola pare, pa ga zaokruživanje spušta na `600,00`; vektor je zamenjen jednoznačnim `600,006`, a granični slučaj sada tvrdi invarijantu (iznos u fajlu ne prelazi otvoreno) umesto smera zaokruživanja. Sve ostale 186 provere zelene, uključujući sve nove kapije. **Preostaje re-run posle ove ispravke i smoke ekrana** (unos → izmena podataka → osvežavanje → odbijeno generisanje).
- **Napomena za smoke:** kontrolni delovi novih testova čitaju **stvarne** podatke te mašine. Ako neki padne, to je po pravilu **stvaran nalaz u podacima** (oštećena šifra otkupa, blok bez kooperanta, dvojnik kooperanta), a ne greška testa.
- **Dodirnuti moduli:** `modNovac` (`ApplyAvansToOtkup` — fail-closed guard na dupli `OtkupID`), `modBankaExportPregled` (novi `ClampOverridesToOpen`, `ValidateNalogSaldo`, `BuildOpenAmountDict`, `BuildNalogCsvPayload`, `BuildOtkupOwnerIndex`; `GenerisiNalogeCSV` dobio finalnu kapiju i `outOdbijeno`; `BuildKooperantTekuciRacunCache` vraća i duplikate), `frmBankaExportPregled` (`PruneStaleOverrides` → `Function`, status o usklađivanju, cent-domen u `txtIsplatiti_Exit`, `PrimeniAvansTX` + oba avans dugmeta, `btnGenerisiCSV_Click`, čišćenje liste na grešci, `Err` pre `LogErr`), `modTestBanka` (T12–T20 + izolovane transakcije). Prateći: `docs/KNOWN_ISSUES.md` (AUD-026 zatvoren; FM-0021 #5 zatvoren strože nego što je katalog predlagao; FM-0021 #6 ostaje otvoren i zaveden), `docs/REFAKTOR_PLAYBOOK.md`, `docs/PLAN_SANACIJE.md` (M6 ✅), `docs/ARCHITECTURE_CHANGELOG.md` (v6.48), `CLAUDE.md`.

---

## vba-v2.39.1 — 2026-08-13
> Verzija/datum se **finalizuju pri `tools/release.sh`**. **Razvojna verzija — ne menja `.xlsm`.**
> Nijedan `.bas`/`.cls`/`.frm` nije dirnut, nema promene šeme, nema promene ponašanja
> aplikacije. Operater ovde nema šta da testira; entry postoji da se u istoriji vidi
> **kada je verifikacija pre commita prestala da bude „na dobru volju"**.
>
> **Suština:** dve compile greške koje su ranije čekale operatera u VBE-u (`Alt+F11 →
> Compile`) sada se hvataju iz izvora, u milisekundama, pre commita — bez Excela.

**Compile greške se hvataju pre commita**

- **„Sub or Function not defined" i „Wrong number of arguments" više ne čekaju Excel.** To su, uz „Ambiguous name", tri najčešće compile greške u ovom projektu, i sve tri se vide iz samog izvora. `tools/vba_check.py` dobija dve nove provere — `NEDEFINISAN` (poziv procedure koja nigde u `src-vba/` nije definisana) i `ARNOST` (poziv sa brojem argumenata van deklarisanog opsega) — uz postojeći `DUPLIKAT`. Ranije se za to pravio headless compile gate; posle **četiri pokušaja** koji su svaki put lagali (v. `docs/EXCEL_TEST_HARNESS.md`) zaključak je da za ove tri greške Excel uopšte ne treba.
- **Nalaz stiže odmah, ne pred commit.** Checker se vrti kao PostToolUse hook (`.claude/hooks/vba-check.sh`) nad fajlom koji je upravo izmenjen, pa greška ne stiže posle importa u Excel nego u trenutku izmene.
- **Provera je namerno uska, jer je lažan nalaz gori od propuštenog.** Gleda samo `.bas` module (u `.frm`/`.cls` se nasleđeni članovi zovu bez kvalifikatora — `Repaint`, `Show`, `SetFocus` — pa bi lažni nalazi bili pravilo, ne izuzetak) i samo poziv u poziciji naredbe (`x = Foo(1)` se ne dira: bez tipova se poziv funkcije ne razlikuje od indeksiranja niza). Ime sa zagradom bez `Call` prefiksa računa se kao indeks kolekcije/niza — svih 8 prvih lažnih nalaza pri uvođenju bilo je upravo to. Ime definisano na više mesta sa različitom arnošću se isključuje.
- **Šta i dalje traži Excel:** tipovi, nedeklarisane promenljive, greške u `.frm`/`.cls`. Za to ostaje `Alt+F11 → Debug → Compile VBAProject` na operaterskoj mašini.

**Harness za Excel — četiri laži i šta je od njih ostalo**

- **`run_vba.py` je tražio fixture koji nikad nije postojao**, pa se nikad nije ni pokrenuo do kraja.
- **„COMPILE NEJASNO" je prolazilo kao „REZULTAT: ZELENO".** Nejasan ishod je izveštaj prikazivao kao uspeh — najgora vrsta greške u kapiji, jer se crveno stanje čita kao zeleno.
- **Import je upisivao VBA header u kod**, pa je sveska završavala u break modu.
- **Verdikt se čitao iz `Enabled` stanja stavke `Debug → Compile`**, što je treći put dalo lažno zeleno — provera ne kompajlira projekat. Uz to: minimizovan VBE prozor znači da meni uopšte ne reaguje, pa verdikt više ne laže ni u crveno.
- Ono što je od harnessa preživelo i dalje radi (watchdog za modalne dijaloge, merenje compile-a mimo menija) i dokumentovano je u `docs/EXCEL_TEST_HARNESS.md`; zaključak „pravo rešenje je statičko, ne headless" stoji na vrhu tog dokumenta.

**Kontekst pravila (`CLAUDE.md`)**

- **`CLAUDE.md` je sveden na ono što važi uvek**, a detalji po oblastima preseljeni u `.claude/rules/` (9 fajlova sa `paths:` frontmatter-om — podaci i config, forme i kontrole, agrohemija i cene, banka, sync i self-update, VBA izvor, testovi, git i release). Pravila se učitavaju kad se ta oblast dira, umesto da svaka sesija nosi ceo tekst.

**Rizik za podatke i verifikacija**

- **Rizik za podatke: nikakav.** Nema promene šeme, `.frx` netaknut, nijedna VBA linija nije izmenjena, nema novih `Poruka()` ključeva. Menjaju se samo `tools/`, `docs/`, `.claude/` i `.gitignore`.
- **Stanje verifikacije:** `python3 tools/vba_check.py` nad celim `src-vba/` — **175 fajlova, čisto, exit 0**, dakle nijedan lažan nalaz od dve nove provere. Negativna proba (namenski modul sa 2 pogrešne arnosti i 2 nedefinisana poziva) daje **tačno ta 4 nalaza** i ne prijavljuje ispravan poziv u istom modulu. `git merge-tree` vs `main` čist.
- **Dodirnuti fajlovi:** `tools/vba_check.py` (`NEDEFINISAN`, `ARNOST`, `collect_definitions`, `collect_arities`, `check_undefined`), `tools/run_vba.py` (harness, četiri ispravke verdikta i importa), `.claude/hooks/vba-check.sh`, `.claude/settings.json`, `.claude/rules/*` (9 novih), `CLAUDE.md` (sveden), `docs/EXCEL_TEST_HARNESS.md` (nov), `.gitignore`.

---

## vba-v2.39.2 — 2026-08-13
> Verzija/datum se **finalizuju pri `tools/release.sh`**.
> **Rizik za aplikaciju: praktično nikakav** — u `frmOtkup.frm` su promenjene dve
> linije, obe bez efekta na produkciju (v. „Rizik za podatke i verifikacija").
> Nema promene šeme, `.frx` netaknut, nema novih `Poruka()` ključeva.
>
> **Suština:** do sada je verifikacija hvatala sintaksu i compile greške. Sada
> postoji i suite koja hvata izmenu koja se **uredno kompajlira, a menja
> ponašanje** — i koja je **dokazano crvena** nad namerno pokvarenim kodom, ne
> samo zelena nad ispravnim.

**Test suite koja pada na ponašanju**

- **Tri testa nad `frmOtkup.ClearOtkupFields`**, rutinom koja se izvršava posle svakog snimanja otkupnog lista: datum otpremnice se NE briše (sledeći blok ide u niz istog datuma), broj zbirne ostaje popunjen (i drugi blok dobija istu zbirnu), a kooperant se briše (sledeći unos je nov partner). To je klasa buga koja nastaje pri radu na UI-ju: ukloni se ili doda jedna linija, kod se uredno kompajlira, `vba_check` je zelen — a operater od sledećeg dana kuca datum i broj zbirne iznova na svaki unos.
- **Dokazano u OBA smera, sva tri testa.** Nad čistim kodom `python tools/run_vba.py --suite RunAllTests` → exit 0, `TESTS 3 ukupno, 0 palo`. Nad namerno vraćenim brisanjem datuma → exit 2 uz `FAIL T_PosleSnimanja_ZadrzavaKontekstOtpremnice -- ocekivano [15.3.2026], dobijeno []`; isto i za zbirnu i za partnera, svaki sa svojom porukom. Posle `git checkout` ponovo exit 0. Svaka sabotaža obara **samo svoje** testove — ostali ostaju `OK`, dakle testovi su specifični, ne padaju u gomili. (Zelena suite koja nije dokazano crvena ne dokazuje ništa; to je u PR #181 bio ishod četiri puta.)
- **Snapshot hvata i polja koja niko nije tražio da se provere.** `DumpKontrole` snima svih 43 kontrole forme kao sortirano `ime=vrednost` (sortira postojećim `modArrayUtils.SortArray`) i poredi sa `tests/golden/*.txt`. Kad golden ne postoji, test ga upiše i **padne** — nov golden mora proći ljudski pregled pre nego što postane merilo. U dokazu je snapshot samostalno uhvatio obe regresije (`golden [cmbKooperant=] vs tekuci [cmbKooperant=KOOP-TEST-1]`).

**Fixture bez ijednog klijentskog podatka**

- **`tools/make_fixture.py`** pravi `tests/fixtures/otkup_test.xlsm` iz **donor** sveske (npr. `builds/AgriX_2.28.4.xlsm`): obriše redove iz svih tabela osim kataloga, pa poseje samo sintetiku — 3 kooperanta, 2 parcele, tri otpremnice (sa zbirnom i ostatkom 600, bez zbirne, bez zbirne ali sa blokom koji zbirnu nosi), `APP_SETUP_COMPLETED=DA`, licenca off. Donor se nikad ne dira. Time nijedan pravi kooperant ne može da završi u golden fajlu koji ide na GitHub.
- **Zašto donor a ne „od nule":** osnovnu šemu ne pravi nijedan kod — `Ensure*` rutine u `modSetup` samo **dodaju kolone** na postojeće tabele, a spiskovi kolona osnovnih tabela žive isključivo u `.xlsm`. Zakucavanje tih spiskova u Python napravilo bi drugi izvor istine koji konkuriše svesci (`CLAUDE.md` §4).
- **`tools/dump_schema.py`** ispisuje šemu bilo koje sveske (sheetovi, `CodeName`-ovi, tabele, kolone, broj redova) — samo čitanje, sveska se ne snima. Batch varijanta onoga što `modSetup.DebugKoloneTabele` radi interaktivno kroz `InputBox`, jednu tabelu po pozivu. Korisno i mimo testova, za dijagnozu schema drift-a po instalaciji.

**Tri kvara koja bi suite lažno zelenila ili trajno crvenila**

Sva tri su izašla tek kroz pokretanje, ne kroz čitanje koda:

- **Cleanup je visio na „Want to save your changes?" bez ijednog čuvara.** U `run_vba.py` je redosled bio `killer.cancel()` → `watchdog.stop()` → `Workbooks.Close()`. Goli `Close()` pita za snimanje kad je sveska prljava, a `DisplayAlerts=False` ne pomaže jer ga suite u svom čišćenju vrati na `True` — u tom trenutku nema ko da klikne dijalog ni ko da ubije proces. Sveska se sada zatvara **pre** gašenja čuvara, uz eksplicitan `SaveChanges=False`.
- **Golden je bio neuporediv zbog dijakritike.** VBA `Print #` piše u ANSI kodnu stranu (cp1252 na ENG Windows-u) koja `ć` nema, pa se golden snimi osakaćen i svako sledeće poređenje pada — a poruka o razlici izgleda besmisleno („golden [Vrsta voca] vs tekuci [Vrsta voca]"), jer se i ona gubi na istom mestu. `DumpKontrole` sada escape-uje sve van štampanog ASCII-ja u `\uXXXX`, isto pravilo koje već važi za VBA izvor.
- **Golden bi pukao na svakom svežem klonu.** VBA ga piše sa `vbLf`; git na Windows-u konvertuje u CRLF pri checkout-u i pročitani golden prestaje da bude jednak dump-u. `.gitattributes` sada drži `tests/golden/*.txt` na `eol=lf`, a `ReadTextFile` izbacuje `CR` pri čitanju.

**Regresija se hvata bez sećanja operatera**

- **`Stop` hook** (`.claude/hooks/vba-test.sh`) pušta suite na kraju sesije kad je `src-vba/` diran — u radnom stablu ili u poslednjem commit-u. Bez `pywin32`/Excela prolazi **tiho**, pa Claude Code sesija na webu nije ometana; tamo i dalje radi `vba_check` kroz `PostToolUse`.
- **`run_vba.py` je dopunjen, ne prepisan:** `RunAllTests` u `SUITES` katalogu, verdikt iz `last_run.txt` pored sveske (a ne iz „`Run()` nije pukao", jer `modTest` hvata grešku po testu da jedan pad ne obori ostale), **nema `last_run.txt` → exit 2**, golden fajlovi u temp pre rana i nazad posle.
- **Compile verdikt više ne obara run kad suite-ovi idu.** `COMPILE NEJASNO` se i dalje ispisuje nepromenjen, ali odgovor nose testovi: da bi se `RunAllTests` uopšte pokrenuo, VBA mora da kompajlira `modTest` i sve što on referencira — a to je baš kod pod testom. Eksplicitan compile `FAIL` i dalje pada, kao i `NEJASNO` uz `--compile-only`, gde je probe jedini izvor istine.

**Zatečeni padovi — prijavljeni, ne popravljani u ovom PR-u**

Prvo pokretanje golog `python tools/run_vba.py` (pun set suite-ova) dalo je dva pada nezavisna od ove suite. Nijedan nije diran ovde; jedan je u međuvremenu rešen zasebno:

- **`RunBankaImportTestSuite` — `PASS=186 FAIL=1`** na prvom pokretanju. **Rešeno u #183, i to nije bio bug u produkciji nego u test vektoru:** `600.005` se u `Double`-u čuva kao `600.00499999...`, dakle ispod pola pare, pa ga half-up zaokruživanje korektno spušta na `600.00` — `ValidateNalogSaldo` je bio u pravu. Vektor je zamenjen jednoznačnim `600.006`, a za vrednost tačno na pola pare se sada tvrdi invarijanta (iznos u fajlu ne prelazi otvoreno) umesto smera zaokruživanja. Vredi zabeležiti kako je nađen: suite koja je puštena zbog sasvim drugog posla iznela je dvosmislen vektor koji je stajao u repou.
- **`TestLicense_All` — „Cannot run the macro".** Makro postoji (`modLicenseTests.bas:18`, `Public Sub`) i import prolazi bez primedbe, pa je najverovatnije compile greška u `modLicenseTests` (VBA kompajlira lenjo i odbija da pokrene makro iz modula koji ne prolazi). **Nije potvrđeno** — za potvrdu treba `Alt+F11 → Debug → Compile` ručno.

Dok `TestLicense_All` stoji, akceptaciona komanda za ovu suite je `--suite RunAllTests`, a ne goli poziv. Kad se i to raščisti, u `.claude/hooks/vba-test.sh` se `--suite RunAllTests` menja golim pozivom i hook počinje da vrti ceo podrazumevani set — blizu 300 provera pod gate-om umesto tri.

**Šta je zapravo najveća promena**

Vredi izdvojiti, jer se lako previdi pored tri nova testa: do sada se **nijedna** postojeća suite nije pokretala kroz `run_vba.py`. Compile probe je vraćao `NEJASNO`, `rc = 2` je padao pre suite petlje, i petlja se nikad nije dosegla — suite su postojale, ali samo kao ručni `Alt+F8`. Otkad probe više ne obara run, jedna komanda vrti ceo podrazumevani set: `RunSheetsJsonParserTests` (72), `RunBankaImportTestSuite` (187), `RunFakturaSmokeSuite` (35), `RunIzvestajTests`, `RunAllTests` (3) pod gate-om, plus `Test_StornoCentar_All` (88) kao blind. Tri nova testa su manji deo dobitka od toga.

**Rizik za podatke i verifikacija**

- **Rizik za podatke: nikakav.** Nema promene šeme, `.frx` netaknut, nema novih `Poruka()` ključeva. Fixture je lokalan artefakt (`.gitignore`) sa isključivo sintetičkim podacima; golden fajlovi ne sadrže nijedan klijentski podatak.
- **Rizik za aplikaciju: dve linije u `frmOtkup.frm`.** `ClearOtkupFields` je `Private` → `Public` (test seam; poziva je samo forma i `modTest`), a `cmbKooperant.SetFocus` je dobio `If Not IsTestMode() Then` — forma koja nije `.Show`-ovana ne može da primi fokus, pa bi test padao na fokusu umesto na ponašanju. U produkciji `IsTestMode()` je uvek `False` (flag postavlja isključivo test modul), pa je ponašanje identično. **To je jedina tačka koju operater treba da proveri u Excelu:** posle snimanja otkupnog unosa fokus mora i dalje da skoči na polje kooperanta.
- **Novi moduli u `.xlsm`:** `modTestMode.bas` (mora da se isporučuje — referencira ga `frmOtkup`) i `modTest.bas` (test modul, kao postojeći `mod*Tests`; ne radi ništa dok se ne pozove).
- **Stanje verifikacije:** `python3 tools/vba_check.py` — **177 fajlova, čisto, exit 0**. `tools/run_vba.py --self-test` čist. `--suite RunAllTests` **3/3 OK, exit 0**; sve tri sabotaže daju exit 2 sa imenom ciljanog testa; posle reverta ponovo 3/3.
- **Dodirnuti fajlovi:** `src-vba/modTest.bas` (nov), `src-vba/modTestMode.bas` (nov), `src-vba/frmOtkup.frm` (2 linije), `tools/make_fixture.py` (nov), `tools/dump_schema.py` (nov), `tools/run_vba.py` (dopuna), `.claude/hooks/vba-test.sh` (nov), `.claude/settings.json`, `.claude/rules/testovi.md`, `.gitattributes` (nov), `tests/golden/PosleSnimanja_KontekstOtpremnice.txt` (nov), `docs/TEST_SUITE_OTKUP_HANDOFF.md` (nov).

---

## vba-v2.39.3 — 2026-08-14
> Verzija/datum se **finalizuju pri `tools/release.sh`**.
> **Rizik za aplikaciju: nikakav** — dirani su isključivo test moduli
> (`mod*Tests`, `modTest*`), i to samo dodavanjem `Err.Raise` na kraju suite-a.
> Nijedna produkciona rutina, nema promene šeme, `.frx` netaknut.
>
> **Suština:** zaštita po sesiji je otišla sa **3 provere na ~1050**, i to ne
> pisanjem novih testova nego time što su postojeći prestali da lažu.

**„Blind" suite nisu bile zaštita nego privid**

- **Suite sa `gate: False` je runner prijavljivao kao „prošla bez greške", što NIJE „sve provere prošle".** Rezultat je postojao samo u Immediate prozoru; pale provere niko nije video. Pet takvih suite-ova prevedeno je u `gate`: `RunStornoTestSuite` (181), `Test_StornoCentar_All` (88), `RunPaleteTestSuite` (97), `RunAgrohemijaSmokeSuite` (25), `RunBusinessFlowProSuite` (336), `TestLicense_All` (23). Konverzija je po tri linije jer su sve već brojale padove — samo nisu podizale grešku.
- **Svaka je dokazana u oba smera:** namerno oborena jedna provera → `exit 2` sa imenom baš te suite → `git checkout` → ponovo zeleno. Bez tog dokaza konverzija se ne prijavljuje kao gotova (`CLAUDE.md` §5).
- **„Suite se nije pokrenuo" nije „prošlo".** Uz konverziju su zatvorene četiri putanje tihog `Exit Sub` pre nego što ijedna provera krene — paletiranje isključeno, zatečen `TST-` ostatak, operater odustao na potvrdi, dev-guard odbijen. Sve su runneru izgledale kao `OK`; sada podižu grešku sa porukom koja počinje `suite NIJE pokrenut:`.

**Dva „nalaza" koja to nisu bila**

- **147 palih provera u `RunBusinessFlowProSuite` nije bila regresija nego nepripremljena sveska.** Fixture nastaje iz starijeg donora (2.28.4), a kod je noviji; kolone dodate u međuvremenu ne postoje dok se ne pusti schema upgrade. `run_vba.py` sada **uvek** pušta `EnsureRuntimeSchema` posle importa a pre suite-ova i ispisuje `SCHEMA OK` / `SCHEMA FAIL`. Posle toga suite daje `336/336`. Pala priprema šeme obara run i kad su sve suite zelene.
- **`TestLicense_All` („Cannot run the macro") nije bila compile greška nego zaostali duplikat u svesci.** Fixture je nasleđivao **131** VBA modul iz donora; import prepisuje samo ono što repo ima, pa zaostali modul ostaje i izvršava se — a duplo `Public` ime daje „Ambiguous name". Otud je ručno pokretanje prolazilo (druga sveska), driver padao (fixture), a `vba_check` bio zelen s pravom (duplikata u repou nema). `make_fixture.py` sada uklanja sav kod iz donora; za sveske iz `--workbook` driver ih ne briše nego prijavljuje kao `ORPHAN`.

**Alati**

- **`tools/read_test_log.py`** — čita log sheet suite-a i grupiše padove po temi i po razlogu (`pao / ukupno`), da se masovan pad razlikuje od pojedinačnog. `run_vba.py --keep` sada i **snima** temp kopiju; ranije ju je čuvao u stanju pre rana, pa je trijaža čitala stariji, tuđi run.
- **`Stop` hook pušta goli `run_vba.py`.** Katalog `SUITES` je jedini izvor istine — nova suite ulazi u gate time što je upisana sa `default: True`, bez diranja hook-a.

**Stanje verifikacije**

- Golo `python tools/run_vba.py` → **`EXIT=0`**, 11 suite-ova, `SCHEMA OK`, **bez `BLIND` reda**. Banka 189, BFP 336, storno 181, palete 97, centar 88, json 72, faktura 35, agro 25, licenca 23, `modTest` 3, plus `RunIzvestajTests`.
- Svih šest konverzija pokazano crveno pod sabotažom i vraćeno u zeleno.
- `python3 tools/vba_check.py` — 177 fajlova, čisto.

---

## vba-v2.40.0 — 2026-08-14
> Verzija/datum se **finalizuju pri `tools/release.sh`**.
> **Isporuka: NIJE online update.** Paket nosi **novu formu** (`frmOtkupUI.frm/.frx`),
> a `modSelfUpdate` novu formu namerno rutira na `needsReinstall` (runtime
> `Remove`/`Import` forme ume da korumpira svesku). Prelazak traži **jednokratnu
> punu isporuku** — nov `AgriX_OtkupApp.xlsm` ili `ImportAllVBA` po mašini.
>
> **Legacy se NE gasi.** `frmOtkup` i `frmDokumenta` rade nepromenjeno; novi UI se
> gradi paralelno i preuzima posao tek kad u kompletu bude umeo sve što one umeju.

**Novi ekran otkupa i dokumenata (`frmOtkupUI`)**

- Jedna runtime forma umesto 19 zatečenih: ljuska `modOtkupUI` (zaglavlje, KPI, kontekst, forma, mreža, sidebar) + ekranski moduli `modScr*` koji se zovu kasno vezano. Kontrole se grade u runtime-u (`Controls.Add`), `.frx` se ne dira, `WithEvents` živi isključivo u `clsFlatBtn`.
- **F1 Otkupni list** — pet lista (svi listovi, otpremnice, blokovi, izgubljeni, kooperanti), radnje nad redom (štampa, storno, specifikacija, preuzmi), višestruki izbor otpremnica za štampu, filter po opsegu datuma i štampa svega filtriranog, traka aktivne otpremnice sa ostatkom.
- **Prefill sa otpremnice:** klik na otpremnicu prepisuje sve što ona zna (datum, roba, otkupno mesto, vozač, tip ambalaže, cena, broj zbirne), operateru ostaju kooperant i količine.
- **Palete** — tri liste (palete, stavke, prerade) sa radnjama (štampa, PDF, zatvaranje, storno, nepotpune).
- Pravila unosa iz legacy formi: cena iz cenovnika, tip ambalaže iz kulture, živi zbir kg / neto iz bruta, info o paleti, podrazumevani proizvod, predlog broja dokumenta, kontekst otkupnog mesta i datuma.

**Upis (Faza B, u toku)**

- Poslovna logika unosa **izdvojena iz formi u module bez kontrola**: `modOtkupUnos` (otkupni list) i `modDokUnos` (otpremnica, zbirna, prijemnica). Provere u legacy redosledu, bruto→neto sa zamrzavanjem `BrutoKg`, `Save*_TX`, štampa, auto-lanac hladnjače, auto-zbirna (MALINA), završetak ispravke.
- Knjiže se **otkupni list (F1)**, **otpremnica (F2)**, **zbirna (F3)** i **prijemnica (F4)**. OM ulaz i kupci izlaz još ne upisuju.
- **Zbirna (F3)** nosi svoju kapiju iz legacy: kilogrami i ambalaža moraju da se poklope sa nestorniranim otpremnicama te zbirne (`ValidateZbirnaPreUnosa`), i to **bez obzira na podešavanje `VALIDACIJA_UNOSA`**. Uz to: blokada kad izvor ima Klasu II a prekidač je isključen (inače bi se Kl.II tiho izgubila). Zbirna **nema** bruto→neto ni cenu — `tblZbirna` ih nema, a otpremnice koje zbraja su već u netu.
- **Prijemnica (F4)**: broj zbirne je obavezan i zbirna mora da postoji (ponašanje po `PRIJEMNICA_ZBIRNA_PROVERA`), bruto→neto po klasama, pravilo „1 zbirna = 1 prijemnica" kao pitanje, auto-štampa i grupni otkupni list samo za default hladnjaču, status palete uz potvrdu upisa.
- **Šta prijemnica još ne radi:** ispravku posle storna (prevezivanje paleta stare prijemnice). To traži storno okvir (Faza D) i ostaje u `frmDokumenta`, gde takva ispravka i može da nastane — novi UI još ne ume da stornira prijemnicu.
- Keš isplate uz otkupni list **namerno ne prelaze** — idu isključivo kroz F5/F6.

**Brojevni niz po režimu — ispravljena tri kvara**

- **Zbirna se broji po vozaču**, ne po otkupnom mestu. Predlog je slao stanicu kao entitet, a `ApplyMirrorPrefix` gleda da li je entitet mirror-vozač — pa se `S` prefiks pojavljivao i van zbirnih.
- **Reversi nisu dobijali broj:** `modeKey` vraća `"REVERSI"`, a uslov je proveravao `"REVERS"`.
- **Prijemnica je dobila predlog** (`GenerateBrojPrijemnice`, `1/ddmmyy`), gejtovan na hladnjača-kupca; ostali kupci nose eksterni broj i polje se ne dira.
- Predlog prati promenu otkupnog mesta, datuma, režima, vozača i kupca. Pad zaključavanja stanice više ne preskače preračun.

**Ispravke koje se tiču podataka**

- **`ParcelaID` je u `tblOtkup` upisivao ceo prikazni string** umesto ID-a: vadio ga je iz teksta tražeći `" - "`, a lista parcela gradi prikaz sa `ChrW(183)`. Sada se čita skrivena kolona combo-a, kao kod svih ostalih.
- **Datum u putu upisa bio locale-zavisan** (`IsDate`/`CDate`), dok ga predlog broja i zaključavanje već čitaju kroz `TryParseDateValue` — isti tekst je mogao dati dva datuma. Sada ide isključivo kroz deterministički parser.
- **Kontekst otpremnice se gubio posle snimanja:** datum se vraćao na danas, pa je drugi blok otpremnice od 22.07 dobijao današnji datum i današnji broj. Isto i broj zbirne, koji je usput brisalo i ponovno punjenje liste (`ComboBox.Clear` briše i vrednost).
- **Zaključana stanica se nije puštala** pri zatvaranju ekrana (forma se sakriva, `Terminate` ne puca).

**Revizija ulaznog sloja 1:1 sa legacy**

Prođen ceo reaktivni sloj obe legacy forme i upoređen sa novim UI-jem; zatvoreno: parcela postavlja vrstu/sortu, promena stanice briše kooperanta i parcelu, lista kooperanata se sužava na otkupno mesto (`KOOP_FILTER_BY_OM`), prazan entitet briše predlog broja, promena kupca u prijemnici briše broj pre predloga, predlog uvažava i udaljeni maksimum (Google) na promenu stanice i datuma. Popis svih legacy handlera sa statusom je `docs/UI_MIGRACIJA_KATALOG.md` (Z3b).

**Alat**

- `tools/vba_check.py`: uzak izuzetak od `DUPLIKAT` za ugovor ekrana (`Scr_*` u `modScr*`) — ljuska ih zove isključivo kvalifikovano i kasno vezano, pa „Ambiguous name" ne nastaje. Dokazano u oba smera: ista procedura u bilo kom drugom modulu i dalje pada.

- `tools/sabotaza.py`: sedam imenovanih sabotaža nad `modOtkupUI` (`--lista` / `--vrati`), svaka obara tačno jedan test — druga polovina dokaza iz `CLAUDE.md` §5. Rešava tri zamke koje su već ujedale: CRLF vs LF sidro, uvlačenje, i vraćanje obrnutom zamenom umesto `git checkout --` (koji je jednom pojeo test seam-ove).

- **Oba hooka su na Windows mašini bila mrtva — ispravljeno.** Interpreter se birao po `command -v python3`, a na Windows-u PATH sadrži Microsoft Store execution alias `python3` koji **postoji kao fajl**, ali svaki poziv ispiše „Python was not found" i izađe sa 49. Posledice su bile različite i obe tihe u pogrešnom smeru: `vba-check.sh` (PostToolUse) nije mogao da isparsira payload, `file_path` je ostajao prazan i hook je izlazio **0** — provera VBA izvora **nikad se nije izvršila, bez ijedne poruke**; `vba-test.sh` (Stop) je na tom interpreteru obarao `who_writes.py --check` i lažno prijavljivao „`WHO_WRITES.md` je zastareo" uz `exit 2` na svakom Stop-u, pa se do žiga i Excela nije ni stizalo — ceo brzi set iz v2.40.0 bio je nedostižan. Sada se proverava **izvršavanje** (`"$PY" -c ""`), ne postojanje fajla.

- `.claude/settings.json`: allow lista prebačena sa „pravilo po skripti" na prefiks po familiji (`python tools\`, `powershell -File tools\`, read-only `git`, read-only `gh`), i to u obe forme — `Bash(...)` i `PowerShell(...)`. Nova skripta u `tools/` više ne traži novo pravilo. Write operacije (`commit`, `push`, `reset`) ostaju namerno van liste.

**Testovi ponašanja za novi UI**

- Tri nova testa u `modTest` (`RunAllTests` sada vrti šest): `T_ClearForm_Ugovor` (datum i broj zbirne su kontekst i ostaju, partner se briše, a bez aktivne otpremnice datum se vraća na danas), `T_ParseDatum_Ugovor` (prazno/nečitljivo je `0`, `d.m.yyyy` bez `CDate`, trailing tačka se skida, nemoguć datum se odbija umesto da se prelije) i `T_ParcelaID_IzSkriveneKolone` (ID iz skrivene kolone, sakriveno polje ne šalje parcelu u dokument).
- Tri **test seam-a** koje novi UI zbog toga nosi u isporuci: `ClearForm`/`ParseDatum`/`ParcelaID` su `Public` umesto `Private`; tri `SetFocus`-a su iza `IsTestMode()` (forma bez `.Show` ne može da primi fokus, a u nevidljivom Excelu `SetFocus` ne puca nego trajno visi); `modScrDokumenti.Scr_OtpTestSet` je jedini način da test dobije aktivnu otpremnicu i **tvrdo je gejtovan** — van test-režima ne radi ništa. U produkciji je `IsTestMode()` uvek `False`, pa je ponašanje nepromenjeno.
- **Izolacija suite-a ispravljena:** test koji padne nije stizao do svog `ReleaseOtkupUIForm`, pa su `mFrm`/`Btns`/keš u `modOtkupUI` i aktivna otpremnica u `modScrDokumenti` ostajali sledećem testu — jedna sabotaža obarala je dva testa, a drugi pad je bio lažan trag (`Err.Number=0`, prazan opis). `modTest.RunOne` sada čisti iz `EH` grane, a `Err` čita pre čišćenja.

**Kapija nad `.claude/settings*.json`**

- **Jedan izostavljen zarez je oborio ceo `settings.json`, i to je prošlo neprimećeno pola dana.** Merge `5b3777c` spojio je dve grane koje su obe dopisivale na kraj `allow` niza — **bez ijednog konflikt markera** — i rezultat nije bio validan JSON. Claude Code takav fajl odbacuje **u celini**: nije važilo nijedno od 29 permission pravila ni jedan od dva hook-a, pa se tražilo odobrenje za svaku komandu a `PostToolUse vba_check` se nije palio. Za konfiguraciju koja upravlja celim development workflow-om nije postojao ni `json.load()`.
- **Provera je sada na dva ulaza, jer jedan ne bi bio dovoljan.** `PostToolUse` hvata slučaj kad taj fajl menja `Edit` (odgovor stiže u istom turnu, pre nego što sledeći alat krene bez pravila). `Stop` proverava na **svakom** prolazu, bez obzira šta je dirano — merge/rebase ne prolazi kroz `Edit`, a `check_merge.ps1` nad takvim spojem javlja `CISTO`. Provera je instant, ne traži ni Excel ni `win32com`, i stoji pre svih ostalih kapija.
- **Dokazano u oba smera, šest slučajeva** (u peščaniku, pravi `settings.json` nije diran): pokvaren JSON kroz `Edit` → `exit 2` uz tačan `Expecting ',' delimiter: line 5 column 7`; ispravan → `exit 0` bez poruke; `.bas` sa ne-ASCII bajtom → `exit 2` sa nalazom `ASCII`; čist `.bas` → `exit 0`; `Stop` nad pokvarenim → `exit 2`; `Stop` nad ispravnim → propušta i stiže do sledeće kapije.

**Rezanje pravila — `testovi.md` sa 527 na 316 linija**

- `.claude/rules/testovi.md` je narastao na najveći fajl instrukcija u repou — veći od celog `CLAUDE.md`. Istorija incidenata (147 lažnih padova zbog šeme, `TestLicense_All` / 131 zaostali modul, curenje izolacije, PR #181, četiri putanje tihog `Exit Sub`, `T13`, `Stop` hook sa punim gate-om, zarez u `settings.json`, `python3` alias) preseljena je u **`docs/engineering/postmortems/2026-08-verifikacija.md`**, sa poukom uz svaki slučaj. U pravilima je ostalo samo pravilo.
- **Izbačeno i ono što je bilo doslovan duplikat koda:** tabela `gate`/`blind` suita i tabela sa brojem provera po suite-u (izvor istine je `SUITES` katalog u `run_vba.py`), i tri tabele sabotaža (izvor istine je `sabotaza.py --lista`, koji ispisuje isto). Duplikat nije bezopasan: red `RunAllTests` je u jednom danu bio netačan dva puta (3 → 6 → 11), a sekcija „Stop hook" je i posle ispravke tabele i dalje tvrdila „6 provera".
- Svih 11 testova i 14 sabotaža koje je donela ova verzija **ostaju zapisani** — sečena je arheologija, ne pokrivenost.
- `CLAUDE.md` §7: dekoracija komande ne sme ni na kraj (`; echo "rc=$?"`), jer „always allow" nad compound komandom upisuje pravila **po segmentu, doslovnim tekstom** — pa allow lista dobije `Bash(echo "rc=$?")` ili putanju sa UUID-om sesije, pravila koja se nikad više neće poklopiti. Jedan compound poziv tako proizvede odobrenje sada i odobrenje svaki sledeći put.

**Stanje verifikacije**

Pokrenuto na Windows mašini (Excel + `pywin32`), 14.08.2026:

- `python tools/vba_check.py` → **čisto (187 fajlova)**, exit 0.
- `python tools/run_vba.py --suite RunAllTests` → **TESTS=11, FAIL=0** (šest testova UI ugovora + pet nad upisom zbirne i prijemnice).
- **Dokaz u oba smera:** svih 14 sabotaža iz `tools/sabotaza.py` obara test po imenu, pa se vraća i suite je opet zelena.
- `python tools/run_vba.py` (pun podrazumevani set) → **`EXIT=0`**, 11 suite-ova zeleno, bez `BLIND` reda (~1055 provera).
- **Hookovi, dokaz u oba smera** (posle ispravke izbora interpretera): `.bas` sa `š` → `vba-check.sh` staje sa `ASCII: ne-ASCII bajt (\xc5\xa1)` i `exit 2`; isti fajl u čistom ASCII → `exit 0`. `vba-test.sh`: prvi prolaz odradi `RunAllTests` (**TESTS=11, FAIL=0**) i upiše žig, drugi prolaz nad istim stanjem staje **pre** Excela (`tests/last_run.txt` nepromenjen).
- **`COMPILE` je ostao `NEJASNO`** — `Compile nije izvrsen (nema dijaloga, kontrola ostala aktivna), ishod NEPOZNAT`. Automatski Compile nije prošao, pa se compile status i dalje potvrđuje **ručno**: `Alt+F11 → Debug → Compile VBAProject`. Prijavljuje se kako jeste, ne kao zeleno.

Ograničenje se i dalje prijavljuje kako jeste: testovi pokrivaju `ClearForm`/`ParseDatum`/`ParcelaID` i **provere + bruto→neto** puta upisa zbirne i prijemnice — **ne i sam transakcioni upis** (`Save*_TX`, koji pokrivaju `RunStornoTestSuite` i `RunBusinessFlowProSuite`), mrežu i storno. Forma se gradi bez `.Show`, pa `UserForm_Activate` (raspored, `GoFullScreen`, punjenje mreže) nikad ne ide.

---

## vba-v2.41.0 — 2026-08-14
> Verzija/datum se **finalizuju pri `tools/release.sh`**.
> **Isporuka: običan online update** — za razliku od 2.40.0. Ovaj paket ne nosi
> nijednu novu formu; `frmOtkupUI.frm/.frx` je nepromenjen, dirani su samo
> „meki" moduli. Ko je već dobio 2.40.0 punom isporukom, ovo dobija normalnim
> self-update-om.
>
> **Obavezno posle uvoza:** `Alt+F8 → EnsurePoruke`. Petnaest novih ključeva
> poruka do tada postoji samo u kodu, pa bi se prikazali kao `[KLJUC]`.
>
> **Legacy se i dalje NE gasi.** `frmOtkup` i `frmDokumenta` rade nepromenjeno.

**Upis novca i ambalaže — novi UI sada knjiži i F5, F6 i F7**

Time je upis zatvoren za **sve unosne režime** novog ekrana (F1–F7). Posao je,
kao i do sada, izvučen iz forme u modul bez ijedne kontrole — novi
`modNovacUnos`, treći uz `modOtkupUnos` (F1) i `modDokUnos` (F2–F4):

- **F5 Isplate** → `SaveOMUlaz_TX`. Nosi sve četiri grane **tipa novca** iz
  legacy `btnUnosOMUlaz_Click`: keš iz avansa otkupnog mesta, virman firme,
  virman avansa kooperantu i keš firma→otkupac. To nije formalnost — pogrešan
  tip se ne vidi u formi nego tek u saldu: ne razdužuje otkupni blok i ne skida
  avans otkupnog mesta.
- **F6 Uplate kupaca** → `SaveKupciIzlaz_TX`. Izabrana faktura daje „uplata po
  fakturi" (i zatvara je kroz `UpdateFakturaStatus`); bez nje je red avans kupca.
- **F7 Reversi** → `SaveOMUlaz_TX` sa smerom, auto-broj po `KIND_REV`, PDF
  revers i završetak ispravke posle storna (oba best-effort, ne obaraju potvrdu
  upisa).

Jedan legacy handler pokrivao je dva današnja režima: tamo su novac i ambalaža
mogli u isti dokument, ovde ne mogu — F5 nema polja ambalaže, F7 nema polje
iznosa.

**Dva polja bez kojih upis ne bi bio tačan**

- **„ISPLATA IZ" (F5)** — legacy `tglIzOMAvansa`, sa raspoloživim avansom
  otkupnog mesta u natpisu polja (legacy `UpdateOMAvansSaldo`). Bez prekidača bi
  **svaka** isplata po otkupnom bloku bila virman, pa avans otkupnog mesta nikad
  ne bi bio razdužen.
- **„OTVORENA FAKTURA" (F6)** — legacy `cmbFakturaIzlaz`. Bez izbora fakture bi
  svaka uplata iz novog UI-ja bila avans kupca i **nijedna faktura ne bi bila
  zatvorena**.

Oba čitaju postojeće read-modele (`GetOpenOtkupi`, `GetOpenFakture`,
`GetOMAvansSaldo`) — nijedna računica se ne ponavlja u ekranu.

**Tri namerne razlike u odnosu na legacy** (zapisane i u kodu i u
`docs/UI_MIGRACIJA_KATALOG.md`)

- **Vozač se za čist novac ne traži.** Legacy ga traži uz `VALIDACIJA_UNOSA`,
  ali samo zato što je isti dokument mogao da nosi i ambalažu; `SaveNovac`
  vozača uopšte nema, pa bi provera zaustavljala operatera na podatku koji se
  odbacuje. U F7, gde ambalaža postoji, vozač je obavezan i **bez** stroge
  validacije (firma↔OM ide preko vozača).
- **U F5 partner koji je otkupno mesto jeste entitet novca.** Polje se u tom
  režimu zove „Primalac". Legacy tu mogućnost nije imao — primalac je bio samo
  kooperant, a otkupno mesto se podrazumevalo iz konteksta forme. Kad je partner
  kooperant, entitet ostaje kontekst, tačno kao pre.
- **F7 ne prima kupca kao partnera.** Četiri smera reversa idu isključivo
  kooperant ↔ OM ↔ firma; ambalaža kupca i u legacy ide kroz prijemnicu (povrat)
  i kupci-izlaz, ne kroz revers. Izbor kupca se odbija porukom umesto da se tiho
  proknjiži na pogrešan entitet.

**Novac se više ne knjiži na osnovu onoga što piše u polju**

Tri kapije koje su došle iz pregleda PR-a, sve na istom mestu: greška tu nije
UI bug nego pogrešno proknjižen novac.

- **Ukucano ime koje nije izabrano iz liste zaustavlja dokument.** Padajuće
  liste dozvoljavaju kucanje, a veza ka partneru/bloku/fakturi postoji samo kad
  je stavka stvarno izabrana. Operater je mogao da vidi ime kooperanta u polju,
  pritisne Sačuvaj — i da isplata bude proknjižena na otkupno mesto umesto na
  njega. Isto je važilo za blok (postajao avans) i fakturu (postajala avans
  kupca, faktura ostajala otvorena). Sva tri slučaja su izgledala kao uredan
  dokument. Sada se unos zaustavlja dok se stavka ne izabere; **prazno** polje i
  dalje znači „nije izabrano" i prolazi.
- **Isplate (F5) su vraćene u granice otkupnog mesta.** Lista kooperanata se u
  tom režimu sužava na aktivno otkupno mesto **uvek** (kao u staroj formi, gde
  to nije zavisilo od podešavanja), a lista otvorenih blokova se filtrira po
  istom mestu. Ranije je bila moguća kombinacija „blok sa jednog otkupnog
  mesta, novac knjižen na drugo".
- **Ograničenja iznosa se proveravaju u trenutku upisa, ne u trenutku otvaranja
  liste.** Između punjenja liste i potvrde stanje se može promeniti — drugim
  unosom, uvozom izvoda ili drugim delom programa. Sada se pred svaki upis
  ponovo čita ko je vlasnik bloka/fakture, da nije storniran i koliko je
  **stvarno** ostalo. Provera stoji u samom upisu, pa važi i za staru formu i
  za svaki drugi put do knjiženja, ne samo za novi ekran.

**Dva kvara zatečena usput**

- **Polje „Novac" se nije praznilo posle snimanja.** `ClearForm` ga nije imao u
  spisku, pa je zatečen iznos ostajao u formi — sledeća potvrda bi isplatila isti
  novac drugi put. Najskuplja greška ovog režima.
- **Forma je pokazivala smer reversa koji dokument nije imao.** Prvi segment je
  izgledao izabrano, a interno stanje je bilo „nije izabrano" (0). Sada nijedan
  nije unapred obeležen — smer se bira eksplicitno, jer je prazan smer ranije
  tiho knjižio „OM prima od vozača".

**Ostalo**

- Partner combo dobio **treću skrivenu kolonu sa tipom partnera**
  (kooperant / otkupno mesto / kupac). U mešovitoj listi se tip ne može
  zaključiti iz ID-ja, a od njega zavisi kako se dokument knjiži.
- `modDokUnos.ZavrsiIspravkuAko` postao `Public` i dobio granu za revers —
  umesto da se isto pravilo prepiše u treći modul.

**Stanje verifikacije**

Pokrenuto na Windows mašini (Excel + `pywin32`), 14.08.2026:

- `python tools\vba_check.py` → **čisto (188 fajlova)**, exit 0.
- `python tools\run_vba.py --suite RunAllTests` → **TESTS=17, FAIL=0** (šest
  novih testova: tri nad `modNovacUnos`, tri nad kapijama vlasništva i
  trenutnog ostatka — uključujući jedan koji zove sam upis, bez UI sloja).
- `python tools\run_vba.py` (pun set) → **`EXIT=0`**, 11 suite-ova zeleno, bez
  `BLIND` reda.
- **Dokaz u oba smera:** četrnaest novih sabotaža, svaka obara test **po imenu**,
  pa se vraća i suite je opet zelena. Jedna je popravka postojećeg sidra
  (`clear-zbirna`) — razvezalo se čim je `ClearForm` dobio novo polje, i tek bi
  pri sledećem pokretanju tiho prijavilo da sidro nije jednoznačno. Pravilo je
  zapisano u `.claude/rules/testovi.md`: kad menjaš red koji je nečije sidro,
  promeni i sidro pa ponovo pokaži crveno.
- **Fixture je proširen jednom fakturom** — bez nje kapija nad fakturom nema nad
  čim da radi. Ko vrti testove lokalno mora da regeneriše `otkup_test.xlsm`
  (uputstvo u `.claude/rules/testovi.md`; donor može biti i postojeći fixture).
- **`COMPILE` je i dalje `NEJASNO`** — automatski Compile ne prolazi, pa se
  potvrđuje **ručno**: `Alt+F11 → Debug → Compile VBAProject`. Prijavljuje se
  kako jeste.

Ograničenje: testovi pokrivaju **provere i izbor tipa novca**, ne i sam
transakcioni upis (`SaveOMUlaz_TX` / `SaveKupciIzlaz_TX`, koje pokrivaju
`RunStornoTestSuite` i `RunBusinessFlowProSuite`). Izgled novih polja, PDF revers
i ponašanje nad pravim podacima ostaju na operateru.

**Šta i dalje nedostaje novom UI-ju:** storno okvir (sedam panela — Faza D),
živi prikaz manjka prijemnice i lista zbirnih za izbor, peščanik za vreme upisa,
dva nevezana KPI-ja i prefill iz storniranog dokumenta.

---

## vba-v2.41.1 — 2026-08-16 (hotfix: kapije nad novcem)
> Follow-up na 2.41.0 (PR #190) po pregledu. Dva finansijska nalaza, oba u
> jezgru — važe i za legacy `frmDokumenta` i za novi UI.

**Potpuno plaćena faktura mogla je da primi još jednu uplatu**

`UplataFakturaProblem` je od 2.41.0 ponovo čitao stvarno stanje fakture pred
knjiženje — ali je poslednji uslov glasio:

```
If preostalo > 0 And iznos > preostalo Then
```

Kad je `preostalo = 0`, `0 > 0` je `False` i **cela kapija ćuti**. Scenario je
realan i vodi pravo kroz mehanizam koji je uveden da spreči zastarelo stanje:

1. faktura 10.000, operater otvori F6 dok je preostalo 500;
2. u međuvremenu neko fakturu zatvori u celosti;
3. operater snimi svojih 500 → `preostalo = 0` → prolazi.

Isto je važilo za **preplaćenu** fakturu (`preostalo < 0`).

Uslov koji je tu zaista trebao je *„faktura ima iznos"* — faktura kojoj iznos
nije evidentiran ne sme da blokira uplatu, i to je razlog zbog koga je provera
uopšte bila uslovna. Sada:

```
If iznosFak <= 0 Then Exit Function        ' bez iznosa -> bez kapije
preostalo = ZaokruziNovac(iznosFak - GetUplataForFaktura(fakturaID))
If preostalo <= 0 Then       -> "faktura je već u potpunosti plaćena"
ElseIf ZaokruziNovac(iznos) > preostalo Then -> "veće od preostalog"
```

Poređenje je u **cent-domenu**, bez epsilon tolerancije — isto pravilo koje već
važi za naloge za banku. Bez zaokruženja bi ostatak od 0,000001 zbog float
aritmetike prijavljivao „preostalo 0,00" a ipak blokirao.

**Avans otkupnog mesta nije bio zaštićen u writer-u**

Blok i faktura su od 2.41.0 imali kapiju i u `SaveOMUlaz_TX` / `SaveKupciIzlaz_TX`,
a avans je ostajao samo na UI sloju (`modNovacUnos.IsplataValidiraj`). Writer je
time bio poslednja linija za **dve od tri** stvari koje isti dokument može da
prekorači.

`SaveOMUlaz_TX` sada odbija keš isplatu kooperantu (`NOV_KES_OTKUPAC_KOOP`) koja
prelazi `GetOMAvansSaldo(stanicaID)` — istu vrednost koju UI već proverava.
Virman firme ne troši avans i prolazi nepromenjen.

**Verifikacija**

- `python tools\vba_check.py` → **čisto (188 fajlova)**, exit 0.
- `python tools\run_vba.py --suite RunAllTests` → **TESTS=19, FAIL=0** (dva nova
  testa, oba idu kroz **pravi writer**, ne kroz validator).
- `python tools\run_vba.py` (pun set) → **`EXIT=0`**, 11 suite-ova zeleno.
- **Tri nove sabotaže**, svaka obara test po imenu: `faktura-preostalo-nula`,
  `faktura-bez-iznosa`, `avans-bez-writer-kapije`.

Test za avans je **diferencijalan**: ista suma i isti prazan saldo, dva tipa
novca — keš isplata se odbija, virman firme prolazi. Bez te druge grane test ne
bi razlikovao ciljanu kapiju od opšte blokade svake isplate.

**Fixture** je dobio drugu fakturu, **bez iznosa** (`FAK-TEST-0`). Ona postoji
zbog jednog pravila koje se lako izgubi: faktura kojoj iznos nije evidentiran ne
sme da blokira uplatu. Bez tog reda popravka gornje kapije mogla bi tiho da
ukine i to pravilo — sabotaža `faktura-bez-iznosa` to sada drži.

**Usput, o alatu:** compile grešku koja je nastala u radu (dupli
`Private Const` u istom modulu) `vba_check` **nije** uhvatio — `DUPLIKAT` gleda
samo `Public` imena između modula. Videlo se tek kao „Cannot run the macro" nad
celim projektom. Podsetnik da `COMPILE = NEJASNO` nije formalnost.
## vba-v2.42.0 — 2026-08-15
> Verzija/datum se **finalizuju pri `tools/release.sh`**.
> **Isporuka: običan online update.** Paket ne nosi nijednu novu formu ni sheet;
> `frmOtkupUI.frm/.frx` je nepromenjen — nova dugmad prekidača se prave u
> runtime-u, kao i sve ostalo u toj formi. Nov modul `modStornoDok.bas` je
> običan `.bas` i self-update ga dodaje sam.
>
> **Obavezno posle uvoza:** `Alt+F8 → EnsurePoruke`. Trideset osam novih
> ključeva poruka do tada postoji samo u kodu, pa bi se prikazali kao `[KLJUC]`.
>
> **Legacy se i dalje NE gasi.** `frmOtkup` i `frmDokumenta` rade nepromenjeno;
> `frmDokumenta.btnStorno_Click` nije ni dirnut.

**F8 je postao storno centar — stornira svih devet tipova dokumenata**

Do sada je F8 bio pregled: čitao je isključivo `tblOtpremnica` i pokazivao samo
već stornirane otpremnice. Storno je iz novog UI-ja radio jedino nad otkupnim
listom, iz F1 liste. Sve ostalo je moralo u `frmDokumenta`.

Sada F8 ima **prekidač tipa dokumenta** (isti izbor koji legacy traži kroz
`cmbStornoDokument`, samo se lista bira pa red u njoj — umesto da se broj kuca
napamet):

| Tip | Tabela | Rutina koja radi posao |
|---|---|---|
| Otkupni list | `tblOtkup` | `StornoOtkupByBrDok_TX` |
| Otpremnica | `tblOtpremnica` | `StornoOtpremnicaByBroj_TX` |
| Zbirna | `tblZbirna` | `StornoZbirna_TX` |
| Prijemnica | `tblPrijemnica` | `StornoPrijemnicaByBroj_TX` |
| Isplate / Uplate | `tblNovac` | `StornoNovac_TX` |
| Reversi | `tblAmbalaza` | `StornoOMKoopByBrDok_TX` |
| Fakture | `tblFakture` | `StornoFaktura_TX` |
| Izvodi (ceo) | `tblBankaImport` | `StornoIzvod_TX` |

Fakture i izvodi su tu iako novi UI **nema režim koji ih kreira** — stornirati
se moraju, a legacy ih ima u istom combo-u.

**Kapije su ostale gde jesu — u `modStorno`, ne u ekranu**

Nov `modStornoDok` je četvrti modul bez ijedne kontrole, uz `modOtkupUnos`,
`modDokUnos` i `modNovacUnos`. Izvučen je iz `frmDokumenta.btnStorno_Click`, gde
je jedan `Select Case` radio tri pomešane stvari: razrešavanje broja u ID,
kapije, i poziv pravog `Storno*_TX`. Sada su to tri javne rutine, pa ekran može
da **pita „sme li" pre nego što uopšte ponudi potvrdu**:

- izvod se ne stornira parcijalno (`ResolveNovacForStorno`);
- broj sa više aktivnih novac-redova (avans raspodela deli isti broj) traži
  `NovacID` umesto tihog storna jednog reda;
- `GetIzvodStornoBlokade` je preflight — razlog se vidi pre potvrde, ne kao tih
  neuspeh posle „Da";
- revers bez smera se odbija: četiri smera dele isti brojevni niz, pa broj sam
  ne kaže koji je red u `tblAmbalaza`.

Nijedna od tih provera nije prepisana — sve su već bile javne u `modStorno` i
diže ih i legacy forma.

**Dve razlike u odnosu na legacy — namerne**

- **Storno se bira iz liste, ne kuca.** Legacy traži tip + broj napamet; ovde je
  red u mreži, sa pretragom, filterima i opsegom datuma. Za izvod se broj računa
  čita iz same liste, pa „isti broj na dve banke" više nije dvosmislenost koju
  operater mora da razreši kucanjem.
- **F8 se otvara nad AKTIVNIM dokumentima**, ne nad storniranima — to je lista
  nad kojom se radi. Pregled storniranih ostaje, kroz čip „Otkazane", isti onaj
  koji rade svi ostali režimi. Čipovi su u F8 bili sakriveni dok je taj režim
  umeo samo da prikazuje.

**Šta Faza D još NIJE donela**

Storno iz novog UI-ja je **običan storno** — isti onaj koji legacy radi kad
`TryRunCorrectionFramework` ne preuzme tip. Ispravka i dupli unos posle storna
(`modStornoFlow`, prefill iz storniranog — Z10), Undo operacija, „Nedovršeno" i
Recovery panel i dalje postoje samo u `frmDokumenta`.

**Verifikacija**

Pokrenuto na Windows mašini (Excel + `pywin32`), 15.08.2026:

- `python tools\vba_check.py` → **čisto (189 fajlova)**, exit 0.
- `python tools\run_vba.py --suite RunAllTests` → **TESTS=19, FAIL=0** (dva nova
  testa: mapa tip→tabela→kolone za svih devet tipova, i kapije dispečera pre
  upisa).
- `python tools\run_vba.py` (pun set) → **`EXIT=0`**, 11 suite-ova zeleno, bez
  `BLIND` reda. `Test_StornoCentar_All` i `RunStornoTestSuite` (181 tvrdnja)
  zeleni — postojeći storno nije pomeren.
- **Dokaz u oba smera:** četiri nove sabotaže (`f8-jedna-tabela`,
  `f8-tabela-tipa`, `storno-nema-dok`, `storno-revers-smer`), svaka obara test
  **po imenu**, pa se vraća i suite je opet zelena.
- **`COMPILE` je i dalje `NEJASNO`** — automatski Compile ne prolazi, pa se
  potvrđuje **ručno**: `Alt+F11 → Debug → Compile VBAProject`. Prijavljuje se
  kako jeste.

Ograničenje: testovi pokrivaju **rutiranje po tipu i kapije pre upisa**, ne i
sam storno iz F8 nad pravim podacima (transakcije pokrivaju `RunStornoTestSuite`
i `Test_StornoCentar_All`, ali kroz svoje ulazne tačke, ne kroz ekran). Izgled
prekidača sa devet dugmadi na užem ekranu, tri dijaloga storna izvoda i ponašanje
nad pravim podacima ostaju na operateru — checklista je u PR-u.

---

## vba-v2.43.0 — 2026-08-15
> Verzija/datum se **finalizuju pri `tools/release.sh`**.
> **Isporuka: običan online update.** Nijedna nova forma ni sheet;
> `frmOtkupUI.frm/.frx` je nepromenjen.
>
> **Obavezno posle uvoza:** `Alt+F8 → EnsurePoruke`. Dvadeset novih ključeva
> poruka do tada postoji samo u kodu.
>
> **Legacy se i dalje NE gasi.** `frmDokumenta` i `frmOtkup` rade nepromenjeno.

**Storno posle ovoga nije Da/Ne — bira se šta poslovno znači**

Prethodni paket (2.42.0) dao je F8 sposobnost da stornira svih devet tipova, ali
uvek kao **običan** storno. Ovaj dodaje ono što legacy zove „centralni storno /
ispravka framework": za četiri tipa sa nizvodnim tokom (otpremnica, zbirna,
prijemnica, revers) bira se **šta storno znači**:

| Mod | Šta radi |
|---|---|
| **ISPRAVKA** | pogrešan unos, isti fizički događaj — storniraj stari, unesi ispravan; veze i preračun idu automatski |
| **DUPLIKAT** | dokument nikad nije trebalo da postoji — skini posledice, nema zamene |
| **PONIŠTENJE** | fizički tok se poništava — blokada ako postoje zavisni dokumenti, osim uz svesnu potvrdu |
| **REŠI KASNIJE** | trajan recovery zapis, ne samo poruka koja prođe |

Izbor se nudi **samo kad ima o čemu da se odlučuje** (`CorrectionNeedsDialog`) —
isti smart trigger koji legacy koristi. Revers dobija kratko pitanje storno vs
ispravka, jer je list u lancu.

**Z10: polja se pune iz storniranog dokumenta**

Posle ISPRAVKE operater više ne kuca dokument iznova napamet — menja samo
grešku. Legacy to radi u **četiri kopije** istog računa, svaka vezana za svoje
kontrole (`PrefillOtkupFromStornirano` u `modOtkupBlok`, tri
`Prefill*FromStornirana` u `frmDokumenta`). Ovde je izdvojen jedan račun
(`modStornoDok.PrefillIzStorniranog`), koji vraća opis vrednosti — ekran ne zna
nijednu kolonu tabele, a modul nijednu kontrolu.

Tri pravila iz legacy koja test drži:

- polazi se od **PK-a** stornirane (`OldDocID`), ne od broja — broj nije globalno
  jedinstven (`GenerateBrojPrijemnice` broji po kupcu, pa dva kupca istog dana
  dobiju „1/ddmmyy");
- **datum** se preuzima iz stornirane — ispravka sutradan ne sme da promeni dan;
- **broj** se NE preuzima — ispravka je nov dokument sa novim brojem, a veza na
  stari živi u `tblStornoVeza`.

**Ispravka prijemnice — zatvorena poznata rupa iz Faze B**

Do sada je prijemnica upisana iz novog UI-ja dok visi ispravka dobijala **sveže**
palete, a stare ostajale osirotele. Sada `PrijemnicaValidiraj` prepoznaje da je
unos zamena (traži ispravku na čekanju u `tblStornoVeza` — ne u stanju sesije,
jer se storno pokreće u F8 a unos u F4 i između to dvoje sme da se zatvori
Excel), `PrijemnicaUpisi` preskače svežu paletizaciju (`SetPaletizeSkip` **pre**
upisa) i prevezuje palete stare na novu.

**Safe-stop:** dve ili više ispravki na čekanju → ne bira se naslepo. Pogrešno
pogođena veza bi palete jedne prijemnice prevezala na tuđu robu.

**Hladnjača ispravka radi i iz novog UI-ja**

Završetak je postojao od v6-ui-106 (`modOtkupUnos`), ali ga niko iz novog UI-ja
nije mogao pokrenuti. Sada se posle storna otkupa iz F1 nudi isti izbor kao u
legacy: ISPRAVKA (palete idu na nov lanac) / DUPLI UNOS (skini fantomske stavke)
/ OTKAZI (palete ostaju osirotele, rešava se ručno).

**Šta Faza D još NIJE donela**

Recovery panel, „Nedovršeno" i Undo operacija (stavka 14). Uz njih ide i
**multiselect storna otkupnih blokova** uz DUPLIKAT/PONIŠTENJE — to je deo legacy
overlay panela, koji nije prenet: izbor moda su ovde pitanja, ne panel.

**Verifikacija**

Pokrenuto na Windows mašini (Excel + `pywin32`), 15.08.2026:

- `python tools\vba_check.py` → **čisto (189 fajlova)**, exit 0.
- `python tools\run_vba.py --suite RunAllTests` → **TESTS=21, FAIL=0** (dva nova
  testa: prefill čita tabelu svog tipa, i framework važi samo za četiri tipa).
- `python tools\run_vba.py` (pun set) → **`EXIT=0`**, 11 suite-ova zeleno.
- **Dokaz u oba smera:** četiri nove sabotaže (`prefill-zbirna-kolona`,
  `prefill-tabela`, `prefill-broj`, `framework-otkup`), svaka obara test **po
  imenu**.
- **`COMPILE` je i dalje `NEJASNO`** — potvrđuje se ručno.

**Jedna greška uhvaćena pre nego što je ušla:** `tblZbirna` kolonu količine zove
`UkupnoKolicina`, a ambalaže `UkupnoAmbalaze` — ne `Kolicina`/`KolAmbalaze` kao
ostale tri tabele. Literal bi tiho vratio nulu i prefill ispravke zbirne došao bi
prazan, bez ijedne poruke o grešci. Zato ime kolone nigde nije literal, a
sabotaža `prefill-zbirna-kolona` to drži.

Ograničenje: testovi pokrivaju **prefill i rutiranje po modu**, ne i sam relink
paleta iz ekrana. Sam `ReassignPaleteToPrijemnica_TX` je pokriven
(`RunPaleteTestSuite`, devet slučajeva uključujući razliku u broju gajbica), ali
lepak koji odlučuje **kada** se zove — prepoznavanje ispravke na čekanju i
`SetPaletizeSkip` oko upisa — proverava se **ručno**, po checklisti u PR-u.
Fixture nema nijednu prijemnicu ni paletu, pa se taj tok ne može odvrteti bez
proširenja fixture-a.

---

## vba-v2.44.0 — 2026-08-15
> Verzija/datum se **finalizuju pri `tools/release.sh`**.
> **Isporuka: običan online update.** Nijedna nova forma ni sheet;
> `frmOtkupUI.frm/.frx` je nepromenjen. Nov modul `modScrOporavak.bas` je
> običan `.bas` i self-update ga dodaje sam.
>
> **Obavezno posle uvoza:** `Alt+F8 → EnsurePoruke`. Četrdeset pet novih
> ključeva poruka do tada postoji samo u kodu.
>
> **Legacy se i dalje NE gasi.** `frmDokumenta` i `frmOtkup` rade nepromenjeno.

**Nov ekran „Oporavak" — Faza D je zatvorena**

Četiri legacy panela iz `frmDokumenta` radila su istu stvar: pokazivala šta je
ostalo nedovršeno i nudila prevezivanje. Sada je to jedan ekran u sidebaru, sa
šest lista istog prekidača koji F1 i Palete već koriste:

| Lista | Šta pokazuje | Radnja |
|---|---|---|
| **Nedovršeno** | sve što čeka: pending/manual konteksti + osirotele prijemnice, palete i izgubljeni blokovi | pregled |
| **Osirotele prijemnice** | zbirna im je stornirana ili je nema | Prevezi na ciljnu zbirnu |
| **Zbirne (cilj)** | aktivne zbirne | klik bira cilj |
| **Osirotele palete** | prijemnice sa osirotelim paletnim stavkama | Prevezi na ciljnu prijemnicu |
| **Prijemnice (cilj)** | aktivne prijemnice | klik bira cilj |
| **Vrati storno** | storno operacije koje se mogu vratiti | Vrati storno |

**Zašto dva cilja umesto dijaloga sa listom:** prevezivanje uvek ima izvor i
cilj, a novi UI za izbor cilja već ima obrazac — aktivna otpremnica u F1 i
aktivna paleta na ekranu Palete. Isti obrazac: cilj se bira klikom na red u
svojoj listi i stoji u zoni gore, gde se vidi sve vreme. Legacy je za to imao
combo u panelu; ovde je lista, pa se cilj može i **pretražiti i sortirati**.

**Tri pravila koja test drži**

- liste ciljeva nude **samo aktivne** dokumente — prevezivanje na storniran cilj
  napravilo bi drugu siroticu umesto da reši prvu;
- **jedan red po broju** — klase I i II dele broj, a cilj prevezivanja JESTE broj;
- „Vrati storno" cilja **OperationID**, ne poslednju operaciju po broju: isti broj
  dokumenta može imati više generacija. Zato je prva kolona baš `OperationID`.

**Razlika u odnosu na legacy — namerna**

Upozorenje na siročiće pri otvaranju forme (`CheckVerwaisteDokumente`) se **ne
prenosi kao modalni dijalog**. Umesto njega stalno stoji lista „Nedovršeno" i
brojka u zoni gore. Dijalog pri otvaranju se zatvarao i zaboravljao; lista ne
može da se zaboravi.

**Storno otkupnih blokova uz DUPLIKAT/PONIŠTENJE**

Poslednji deo stavke 13. Kad dokument nestaje bez naslednika, blokovi koji o
njemu vise mogu da padnu s njim. Legacy nudi multiselect u overlay panelu; ovde
je **sve-ili-ništa**, ali se pre pitanja ispiše pun spisak (broj, klasa,
kilogrami, kooperant), pa operater vidi tačno nad čim odlučuje. Delimičan izbor
ostaje na ekranu Oporavak, gde izgubljeni blokovi imaju svoju listu i radnju po
redu. Kapija `BlockStornoDriftReason` (ADR-0001) je ista: blok vezan za **živu**
otpremnicu se ne stornira, jer bi je ostavio precenjenu.

**Verifikacija**

Pokrenuto na Windows mašini (Excel + `pywin32`), 15.08.2026:

- `python tools\vba_check.py` → **čisto (190 fajlova)**, exit 0.
- `python tools\run_vba.py --suite RunAllTests` → **TESTS=23, FAIL=0** (dva nova
  testa: ugovor ekrana i radnje po listi, i ponašanje ciljnih lista).
- `python tools\run_vba.py` (pun set) → **`EXIT=0`**, 11 suite-ova zeleno.
- **Dokaz u oba smera:** tri nove sabotaže (`oporavak-registar`,
  `oporavak-cilj-radnja`, `oporavak-stornirani-cilj`), svaka obara test **po
  imenu**.
- **`COMPILE` je i dalje `NEJASNO`** — potvrđuje se ručno.

**Fixture je proširen jednom storniranom zbirnom** (`ZB-TEST-STORNO`). Razlog je
sam po sebi nalaz: sabotaža `oporavak-stornirani-cilj` je nad starim fixture-om
ostajala **zelena** — u njemu nije bilo nijednog storniranog dokumenta, pa
tvrdnja „lista ciljeva nudi samo aktivne" nije imala nad čim da padne. Test koji
ne može da pocrveni ne meri ništa. Ko vrti testove lokalno mora da regeneriše
`tests/fixtures/otkup_test.xlsm` (`python tools\make_fixture.py --donor <put>
--force`; donor može biti i postojeći fixture).

Ograničenje: testovi pokrivaju **ugovor ekrana i čitanje lista**, ne i sama
prevezivanja iz ovog ekrana. Jezgro je pokriveno drugde
(`ReassignPaleteToPrijemnica_TX` u `RunPaleteTestSuite`, `UndoOperation_TX` u
`Test_StornoCentar_All`); lepak — izbor cilja i redosled potvrda — proverava se
**ručno**, po checklisti u PR-u.

---

## vba-v2.44.1 — 2026-08-16 (hardening po pregledu)
> Bez novih sposobnosti. Zatvara **fail-open** putanje i **dvosmislenost
> identiteta dokumenta** koje je pregled našao u 2.43.0 i 2.44.0.

**Neizvesnost sada zaustavlja upis, ne propušta ga**

Tri mesta su na grešku birala „nastavi":

- **detekcija ispravke prijemnice** — ako čitanje `tblStornoVeze` pukne dok
  ispravka možda čeka, unos je nastavljao kao običan: nova prijemnica dobije
  **sveže** palete, stare ostanu osirotele, a korekcija ostane `PENDING` i čeka
  još jednu prijemnicu. Sada je detekcija izdvojena u `NadjiIspravku`
  (0 / 1 / −1 = STOP) i fail-closed;
- **neočekivana greška u prevezivanju paleta** — prijemnica je već snimljena a
  paletizacija preskočena; bez oznake `MANUAL` korekcija bi ostala `PENDING` i
  sledeći unos bi opet bio ponuđen kao zamena. Sada i ta grana ide u `MANUAL`;
- **`StornoTraziIzborModa`** — `CorrectionNeedsDialog` je sam fail-closed (na
  grešku vraća `True`), a omotač sa `On Error Resume Next` je tu zaštitu
  poništavao: rezultat ostaje `False` i storno prelazi u običan, bez pitanja o
  nizvodnom toku.

**Broj dokumenta nije identitet**

`BrojPrijemnice` se računa **po kupcu**, pa dva kupca istog dana dobiju isti
`1/ddmmyy`. Iz toga su sledile dve greške:

- **prefill je padao nazad na broj.** `FindAnchorRow` je, kad je `OldDocID`
  zadat ali ga u tabeli nema, uzimao „poslednji red istog broja" — dakle tuđi
  dokument. Fallback ostaje samo za stare kontekste koji `OldDocID` uopšte
  nemaju.
- **ciljne liste u Oporavku su deduplikovale po broju**, pa su dva dokumenta
  postajala jedan red. Sada je ključ **broj + vlasnik**, kolona VLASNIK je
  vidljiva, a prevezivanje na dvosmislen broj se **odbija**
  (`modDokumenta.AktivnihVlasnikaPoBroju`) umesto da ode na onaj koji zatekne
  poslednji.

> Prava popravka je da `ReassignPaleteToPrijemnica_TX` i
> `ReassignPrijemnicaToZbirna_TX` prime **identitet dokumenta** umesto broja.
> One su i danas broj-bazirane i zovu ih i legacy forma i `modOtkupUnos`, pa je
> to zaseban zahvat u jezgru. Dok se ne uradi, dvosmislen slučaj staje.

**Verifikacija**

- `python tools\vba_check.py` → **čisto (190 fajlova)**, exit 0.
- `python tools\run_vba.py --suite RunAllTests` → **TESTS=25, FAIL=0**.
- `python tools\run_vba.py` (pun set) → **`EXIT=0`**, 11 suite-ova zeleno.
- **Pet novih sabotaža**, svaka obara test po imenu: `prefill-fallback-po-broju`,
  `prefill-anchor-broj`, `vlasnik-broji-stornirane`, `ispravka-fail-open`,
  `oporavak-cilj-po-broju`.

**Fixture je ponovo proširen** — dve aktivne prijemnice sa **istim brojem i
različitim kupcem**, jedna stornirana sa paletom, i dve ispravke na čekanju.
Bez tog para nijedna tvrdnja oblika „dokument se jednoznačno razrešava po broju"
nije imala nad čim da padne; testovi su bili zeleni jer kolizija u fixture-u
nije postojala. Regeneracija:

```
python tools\make_fixture.py --donor tests\fixtures\otkup_test.xlsm --force
```

---

## vba-v2.44.2 — 2026-08-16 (end-to-end pokriće ispravke prijemnice)
> Bez izmena u ponašanju. Zatvara poslednju rupu u pokriću koju su prethodna
> dva paketa prijavila kao „ostaje na operaterskoj checklisti".

**Najrizičniji put je sada odvrćen automatski**

`SetPaletizeSkip` + prevezivanje paleta + zatvaranje korekcije — do sada
proveravano samo ručno. Novi `T_IspravkaPrijemnice_SkipIRelink` vrti ceo tok.

**Bez novog seam-a u produkcionom kodu.** Rečnik koji `PrijemnicaUpisi` prima
već je javni ulazni ugovor (`NoviPrijemnicaUnos` ga i objavljuje), pa test
postavlja `ispravkaID` direktno. `MsgBox` iz `PrepoznajIspravkuPrijemnice` tako
uopšte nije na putanji, a odluka koju taj dijalog donosi već je pokrivena kroz
`NadjiIspravku`.

Test je **diferencijalan**: isti upis se izvrši dvaput — jednom bez ispravke
(kontrola: sveža paletizacija *mora* da odradi) i jednom kao ispravka. Bez tog
para „nema svežih paleta" ne bi dokazivalo ništa: isto bi se videlo da je
paletiranje ugašeno u Podešavanjima.

**Dve verzije testa bile su placebo — sabotaža ih je otkrila**

Prva je brojala aktivne paletne stavke. Druga je merila gajbice. **Obe su
ostajale zelene** kad se `SetPaletizeSkip` ukloni.

Izmereno stanje pokazuje zašto — bez preskakanja se sveža paleta ipak napravi,
a `ReassignPaleteToPrijemnica_TX` je odmah **stornira**:

```
sa preskakanjem : [ST T-ISPR-1 gaj=40]
bez preskakanja : [ST T-ISPR-1 gaj=40] + [ST T-ISPR-1 gaj=40 st=Da]
```

U aktivnom preseku razlike nema. Tvrdnja zato broji **sve** stavke, uključujući
stornirane — što je i tačno ono što komentar uz `SetPaletizeSkip` opisuje kao
štetu: „kreirale bi se palete koje se odmah storniraju (prazna otvorena paleta +
potrošen broj)". Broj palete se ne vraća.

**Verifikacija**

- `python tools\vba_check.py` → **čisto (190 fajlova)**, exit 0.
- `python tools\run_vba.py --suite RunAllTests` → **TESTS=26, FAIL=0**.
- `python tools\run_vba.py` (pun set) → **`EXIT=0`**, 11 suite-ova zeleno.
- **Tri nove sabotaže**, svaka obara test po imenu: `ispravka-bez-skipa`,
  `ispravka-bez-relinka`, `ispravka-context-ostaje`.

Time sa operaterske checkliste ispada tačka 7 iz PR #194 — ostaje samo ono što
se zaista ne može automatizovati (izgled, štampa, ponašanje nad pravim podacima).

---

## vba-v2.45.0 — 2026-08-16 (identitet dokumenta na granici prevezivanja)
> Završnica Faze D po pregledu. Bez novih sposobnosti — izbacuje **goli poslovni
> broj** sa write granice recovery/relink operacija.

**Broj nije identitet**

`BrojPrijemnice` se računa **po kupcu**, broj zbirne **po vozaču**. Dva dokumenta
lako dele broj. Rutine prevezivanja su do sada primale samo broj i skenirale
tabelu po njemu — pa su kod kolizije zahvatale i tuđi dokument.

**Infrastruktura je već postojala.** `GeneracijaID` (identitet logičkog
dokumenta, Kl.I + Kl.II zajedno) pečate svi writeri kroz `ApplyGeneracijaID`, i
to sa već tačnim kompozitnim vlasništvom po tipu — otpremnica `StanicaID`,
prijemnica `KupacID`, zbirna `VozacID + KupacID`. Kolonu pravi
`EnsureSledljivostSchema` na svakom startu. Nedostajalo je samo da je
recovery/relink putanja **koristi**.

| Sloj | Pre | Sada |
|---|---|---|
| `ReassignPaleteToPrijemnica_TX` | izvor po `bp = oldBroj` | opcioni `oldGeneracijaID`; izvor po `PrijemnicaID` te generacije |
| `ReassignPrijemnicaToZbirna_TX` | svi aktivni redovi broja | opcioni `prijemnicaGeneracijaID`; + kapija nad **ciljnom zbirnom** (vozač + kupac) |
| `GetOsirocenePrijemnice` | grupa po broju | grupa po **broj + generacija**, generacija kao 8. kolona |
| `GetPrijemniceSaOsirocenimPaletama` | grupa po broju | grupa po **broj + generacija**, generacija kao 7. kolona |
| ekran Oporavak | šalje broj | šalje i generaciju iz reda |
| `modDokUnos` ispravka | broj + ad-hoc kapija | generacija iz correction context-a |

Legacy panel čita fiksne indekse (1..6 / 1..7), pa je dodavanje kolone **na kraj**
za njega nevidljivo.

**Bez generacije → fail-closed, ne fail-open**

Stari zapisi generaciju nemaju. Tada se pada na kapiju nad jednoznačnošću broja —
ali kroz **postojeći** `RequireJedanVlasnikPoBroju`, koji već nosi kompozitno
vlasništvo po tipu. Moja ranija `AktivnihVlasnikaPoBroju` (jedan vlasnik, i
fail-open na nedostajuću kolonu) je **obrisana**; jedini račun je sada
`VlasniciPoBroju`, sa prekidačem „broji li i stornirane" — jer je izvor
prevezivanja baš storniran dokument, pa bi brojanje samo aktivnih tu uvek dalo
nulu i kapija ne bi radila.

**Sitno iz istog pregleda**

- ciljna lista je sabirala samo Kl.I; sada obe klase idu u isti red i u zbir;
- filter pretrage se primenjuje **pre** nego što se dokument upamti, pa dokument
  koji ne pogađa prvim redom može da pogodi drugim.

**Verifikacija**

- `python tools\vba_check.py` → **čisto (190 fajlova)**, exit 0.
- `python tools\run_vba.py --suite RunAllTests` → **TESTS=27, FAIL=0**.
- `python tools\run_vba.py` (pun set) → **`EXIT=0`**, 11 suite-ova zeleno.
- Nove sabotaže: `relink-izvor-po-broju`, `relink-ignorise-generaciju`,
  `vlasnik-broji-stornirane`.

**Fixture** je dobio **kolizioni par storniranih** prijemnica (`8/150326`, dva
kupca, svaka sa svojom paletom) i drugu aktivnu zbirnu. Novi
`T_RelinkPoGeneraciji_NeDiraTudjDokument` prevezuje jedan dokument i tvrdi da
drugi **ostaje na svom mestu** — što je tvrdnja koju raniji E2E test nije mogao
da napravi, jer je koristio jedinstven broj.

Par je namerno na **zasebnom** broju od `9/150326`: testovi dele svesku, pa test
koji prevezuje jedan dokument ne sme da potroši podatke onome koji dokazuje
izolaciju.

> **Fixture je gitignored.** Prelazak grane ga ne regeneriše — posle
> `git checkout` pokreni:
> `python tools\make_fixture.py --donor tests\fixtures\otkup_test.xlsm --force`

## v2.45.1 — `v6-ui-124` · identitet i na CILJNOJ strani prevezivanja

Nastavak v2.45.0 po pregledu. Prethodna verzija je izvor prevezivanja prebacila
na `GeneracijaID`, a **cilj je ostao goli poslovni broj** — pola posla.

### Zašto je to bio otvoren nalaz

`BrojPrijemnice` se generiše **po kupcu**, pa dva aktivna dokumenta lako dele
broj. Ciljni dokument se biran po tom broju, a mapa `newById(klasa)` prima jedan
ID po klasi — pobeđivao je **onaj red koji je slučajno poslednji u tabeli**.
Palete su tako mogle da odu tuđem kupcu; isti kvar kao na izvornoj strani, samo
na drugom kraju.

Fixture je to već modelovao (`1/150326` nose i `PRJ-TEST-A` i `PRJ-TEST-B`), a
raniji test je uz taj cilj imao komentar „svejedno koja, bitno je da postoji".
Nije bilo svejedno.

### Izmene

| Sloj | Sada |
|---|---|
| `ReassignPaleteToPrijemnica_TX` | `newGeneracijaID` uz `oldGeneracijaID`; cilj po identitetu |
| `ReassignPrijemnicaToZbirna_TX` | `zbirnaGeneracijaID` uz `prijemnicaGeneracijaID` |
| ciljne mreže (`ZBIRNE`, `CILJPRIJ`) | nose generaciju, kolona prioriteta 3 |
| izbor cilja u ekranu | pamti `mCiljZbirnaGen` / `mCiljPrijemnicaGen` |
| `modDokUnos` ispravka | cilj je upravo upisana prijemnica — PK je poznat |
| `JeIzvornaStavka` → `PripadaDokumentu` | isto pravilo za obe strane, jedna funkcija |

**Labela se čita iz izabranog dokumenta**, ne veruje se pozivaocu: kad generacija
odlučuje, `newBroj` / `targetBrZbirne` se preuzimaju sa tog reda. Neusklađen par
(broj jednog, generacija drugog dokumenta) bi inače tiho upisao tuđi broj.

**Bez generacije — fail-closed.** Dvosmislen ciljni broj zaustavlja prevezivanje
uz razlog za operatera, kroz isti `VlasniciPoBroju` primitiv. Za cilj se broje
samo **aktivni** dokumenti (cilj ne sme biti storniran), za izvor i stornirani.

### Verifikacija

- `python tools\vba_check.py` → **čisto (190 fajlova)**, exit 0.
- `python tools\run_vba.py --suite RunAllTests` → **TESTS=27, FAIL=0**.
- `python tools\run_vba.py` (pun set) → **`EXIT=0`**, 11 suite-ova zeleno.
- Nove sabotaže, obe obaraju `T_RelinkPoGeneraciji_NeDiraTudjDokument` po imenu:
  - `relink-cilj-po-broju` → „roba je stigla na dokument kupca 1 (40 gajbica) —
    očekivano [40], dobijeno [0]"
  - `relink-cilj-bez-kapije` → „bez generacije CILJA dvosmislen broj se odbija —
    očekivano [False], dobijeno [True]"
- `COMPILE` → **`NEJASNO`**, ručna kapija pred release.

## v2.45.2 — `v6-ui-125` · recovery: identitet do kraja lanca

Treća runda po pregledu. v2.45.1 je zatvorila **izbor** dokumenta na obe strane;
ovo zatvara **upis koji je posle tog izbora išao svojim putem**.

### P1 — prevezivanje na zbirnu je vuklo tuđu paletu

`ReassignPrijemnicaToZbirna_TX` je redove `tblPrijemnica` birala po generaciji —
tačno. Zatim je novu `BrojZbirne` propagirala u `tblPaletaStavka` **po
`BrojPrijemnice`**, čime je poništavala ceo taj izbor.

Posledica nije bila „prevezano malo više" nego dokument koji **sam sebi
protivreči**: prijemnica drugog kupca ostaje na staroj zbirni, a njena paleta
završi na novoj. Sledljivost paleta → zbirna → kooperanti tada laže.

Sada se, posle izbora `targetRows`, iz njih čitaju `PrijemnicaID` i upis ide po
njima. Stavka **bez** `PrijemnicaID` (zatečen zapis) sme po broju samo ako taj
broj nosi jedan dokument; kad ga nose dva, transakcija se prekida uz poruku, jer
se ne može utvrditi čija je. Isto pravilo dobio je i `PripadaDokumentu`.

### Identity downgrade — zadata generacija koje nema je greška

Razdvojena su dva stanja koja su se ponašala isto:

| Argument | Ponašanje |
|---|---|
| `""` | pozivalac ne zna identitet (legacy zapis) → fallback po broju, kroz kapiju |
| `"GEN-X"`, a nema ga | **STOP** — pad na broj bi značio da se dira nešto drugo |

Važi za izvor i cilj, u obe rutine.

### Ciljna lista zbirnih je gubila deo identiteta

`RowsAktivni` je vlasnikom smatrala **samo kupca**, a broj zbirne se generiše
**po vozaču**. Dve zbirne istog broja i istog kupca a različitih vozača — u
jezgru dva dokumenta — padale su u **jedan red**, pa operater nije mogao ni da
izabere onaj koji mu treba, a skrivena generacija je nosila generaciju reda koji
je slučajno pobedio.

Sada: grupisanje po **generaciji** kad postoji, vlasnik je **niz kolona**
(`VozacID` + `KupacID` za zbirnu) i prikazuje se ceo.

### Verifikacija

- `python tools\vba_check.py` → **čisto (190 fajlova)**, exit 0.
- `python tools\run_vba.py --suite RunAllTests` → **TESTS=29, FAIL=0**.
- `python tools\run_vba.py` (pun set) → **`EXIT=0`**, 11 suite-ova zeleno.
- Nove sabotaže, svaka obara **baš svoju** tvrdnju:
  - `zbirna-paleta-po-broju` → „tuđa paleta OSTAJE na staroj zbirni — očekivano
    [ZB-TEST-4], dobijeno [ZB-TEST-2]"
  - `generacija-nema-pa-po-broju` → „zadata generacija prijemnice koje nema
    zaustavlja upis — očekivano [False], dobijeno [True]"
  - `zbirna-vlasnik-samo-kupac` → „isti broj zbirne kod dva vozača daje DVA
    ciljna dokumenta — očekivano [2], dobijeno [1]"
- `COMPILE` → **`NEJASNO`**, ručna kapija pred release.

## v2.45.3 — `v6-ui-126` · presuda o relabelu ide nad izabranim dokumentom

Poslednji ostatak starog identity modela u recovery putanji.

### P1 — verdikt je ponovo razrešavao dokument po broju

Writer je već birao po `GeneracijaID` — `srcIds`, `tgtIds`, konkretni
`PrijemnicaID`. Onda je pozivao `EvaluatePaletaReassign(oldBroj, newBroj)`, koja
je dokumente **tražila iznova, po poslovnom broju**, uzimajući prvi red.

Kvar je **tiši** od pogrešnog prevezivanja. Izvor A (jabuka) i tuđi dokument B
(kruška) dele broj, cilj X je kruška: presuda po B vidi kruška → kruška i vrati
`CLEAN`, pa writer **preskoči relabel**. Paleta završi vezana za kruška-prijemnicu
a i dalje označena kao jabuka. Upis je tačan — laže samo etiketa.

Presuda je izdvojena u `PresudiPaletaReassign`, koja **ne čita nijednu tabelu**,
pa ne može da izabere drugi dokument nego onaj koji joj je dat. Writer je zove sa
onim što već ima (identitet izvora se čita u istom prolazu kroz `tblPrijemnica`);
nema drugog čitanja. `EvaluatePaletaReassign` ostaje javna zbog legacy panela, ali
je sada adapter koji prvo razreši identitet — uz opcione `oldGeneracijaID` /
`newGeneracijaID` — pa pozove isto jezgro.

### P2 — „isti broj" više nije „isti dokument"

Guard u `PreveziPalete` je odbijao prevezivanje kad se brojevi poklope. Ali broj
nastaje **po kupcu**, pa ispravka koja menja kupca lako dobije isti poslovni broj
kao original — a to su dva dokumenta i operacija je legitimna. Sada se poredi
generacija kad je obe strane imaju; broj ostaje fallback samo za zapise bez nje.

### Verifikacija

- `python tools\vba_check.py` → **čisto (190 fajlova)**, exit 0.
- `python tools\run_vba.py --suite RunAllTests` → **TESTS=30, FAIL=0**.
- `python tools\run_vba.py` (pun set) → **`EXIT=0`**, 11 suite-ova zeleno.
- Nova sabotaža `verdikt-po-broju` → „stavka je prelabelirana na vrstu ciljnog
  dokumenta — očekivano [TESTVOCE], dobijeno [TESTVOCE2]" — doslovna reprodukcija
  kvara: stavka ostaje označena starom robom.
- `COMPILE` → **`NEJASNO`**, ručna kapija pred release.

**Nije pokriveno testom:** P2 guard u `PreveziPalete` — ta putanja otvara `MsgBox`
potvrde, pa se headless ne vozi. Ide na operatersku checklistu.

## v2.45.4 — `v6-ui-127` · deljena paleta i „isti dokument" u writeru

Dva mesta na kojima je staro pravilo **„broj = dokument"** još curilo.

### P1 — su-stanar na deljenoj paleti se tražio po broju

Pred relabel se proverava da li fizička paleta nosi i tuđu robu; ako nosi,
promena headera bi iskvarila i nju, pa se operacija blokira. Ideja je tačna, ali
se „tuđa" merilo poređenjem brojeva:

```vb
If bpg <> oldBroj And bpg <> newBroj Then ...
```

Dva kupca istog broja i **iste robe** smeju legitimno da dele paletu — roba im je
identična, nema šta da se razlikuje. Za tu kapiju su izgledali kao ista prijemnica
(`bpg = oldBroj`), pa nije okidala: `STEP 2b` bi prepravio header **cele** palete
na novu robu, a su-stanar ostaje stara. Paleta i njena stavka bi od tog trenutka
tvrdile različito.

Sada se pripadnost meri kroz `PripadaDokumentu` nad `srcIds` / `tgtIds`, uz već
postojeći fail-closed za zapise bez `PrijemnicaID`.

### P2 — kapija „isti dokument" je bila popravljena samo u ekranu

`PreveziPalete` je dobio ispravnu logiku prošlu rundu, ali je na ulazu u writer
ostalo staro poređenje, i to **pre** razrešavanja generacija:

```vb
If StrComp(oldBroj, newBroj, vbTextCompare) = 0 Then Exit Function
```

Pravilo je time bilo samo preseljeno iz UI-ja u core. Sada isti princip stoji i u
writeru: generacije kad ih obe strane imaju, broj kao fallback samo bez njih.

Testira se **direktno na writeru** — ekranska putanja otvara `MsgBox` potvrde i
headless se ne vozi, pa bi test kroz UI bio nemoguć a kroz checklistu nepouzdan.

### Ostaje otvoreno (P3)

Legacy poziv `EvaluatePaletaReassign(oldBroj, newBroj)` bez generacija nema
ambiguity guard i može dati **pogrešan preview** kad brojevi nisu jedinstveni.
Nije data-integrity problem — writer taj preview više ne koristi za odluku o
upisu — ali panel u `frmDokumenta` može prikazati pogrešnu ocenu.

### Verifikacija

- `python tools\vba_check.py` → **čisto (190 fajlova)**, exit 0.
- `python tools\run_vba.py --suite RunAllTests` → **TESTS=32, FAIL=0**.
- `python tools\run_vba.py` (pun set) → **`EXIT=0`**, 11 suite-ova zeleno.
- Nove sabotaže:
  - `cotenant-po-broju` → „relabel deljene palete se odbija i uz potvrdu —
    očekivano [False], dobijeno [True]"
  - `writer-isti-broj-odbija` → „isti broj a različite generacije PROLAZI —
    očekivano [True], dobijeno [False]"
- `COMPILE` → **`NEJASNO`**, ručna kapija pred release.

## v2.45.5 — `v6-ui-128` · obim co-tenant provere

Sitna izmena po pregledu: u `ReassignPaleteToPrijemnica_TX` se **prvo** proverava
da li je stavka uopšte na paleti koju relabel dira, pa se tek onda računa
pripadnost dokumentu. Ranije je obrnut redosled radio analizu identiteta nad
svakom aktivnom stavkom u tabeli.

**Ovo nije ispravka buga.** Opisani scenario — zatečena stavka bez
`PrijemnicaID` pod dvosmislenim brojem, na paleti bez veze sa operacijom, koja
prekida relabel — nije dostižan: ista provera sa istim argumentima već se izvrši
u ranijoj petlji koja skuplja `oldRows`/`freshRows` nad **svim** aktivnim
stavkama. Red koji bi ovde podigao grešku podigao bi je tamo, jedan korak pre.

Izmena je svejedno urađena: obim kapije postaje očigledan iz koda, ne radi se
posao koji ne može ništa da promeni, i mogućnost nestaje ako se ranija petlja
ikad promeni.

- `python tools\vba_check.py` → **čisto (190 fajlova)**, exit 0.
- `python tools\run_vba.py --suite RunAllTests` → **TESTS=32, FAIL=0**.
- `python tools\run_vba.py` (pun set) → **`EXIT=0`**, 11 suite-ova zeleno.
- Sabotaža `cotenant-po-broju` i dalje obara `T_DeljenaPaleta_SuStanarPoIdentitetu`.

## v2.46.0 — `v6-ui-129` · F8 nosi identitet izabranog reda

Poslednje mesto na kome je storno centar vraćao kanonski izbor reda nazad u
poslovni broj. **Tek ovim je Faza D stvarno zatvorena** — katalog je do sada
tvrdio da je zatvorena od `v6-ui-121`, što nije bilo tačno.

### Šta je bilo

`StornoRedF8` je izabran red svodio na `GridCell(red, 1)` — broj. Revers je uz
njega nosio smer, izvod broj računa, ostalih šest tipova ništa. Niže je
`modStornoDok` dokument tražio **iznova, po broju**.

Za običan storno to je od #194/#195 uglavnom hvatao owner guard u writeru. Ali
mod **`REŠI KASNIJE` guarded writer uopšte ne zove** — napravi samo trajan
recovery zapis. Taj zapis je mogao zauvek da pokazuje na tuđi dokument, i ništa
to nije prijavljivalo. `ScanPrijemnica` je uz to palete brojao po
`BrojPrijemnice`, pa su i brojke u pregledu mogle biti tuđe.

### Kako identitet putuje

Nevidljiva kolona u `GridCols`, **samo za F8**:

```
"OTKUI_HD_IDENT|" & IdKolonaTipa(mk) & "|txt|0|4"
```

Tri činjenice iz ljuske koje to čine ispravnim — sve tri proverene u kodu:

| Mehanizam | Posledica |
|---|---|
| `SortedView` kopira tačno `mColN` kolona | kolona van `cols` ne preživi sortiranje |
| `mColN = UBound(mCols) + 1` | deklarisana kolona uvek putuje |
| `For pass = 3 To 1 Step -1` | prioritet **4** se nikad ne renderuje |

Mapa `red → identitet` sa strane bi bila **pogrešna**: ljuska sortira posle
`Scr_Rows`, pa izlazni indeks nije indeks u mreži.

Identitet po tipu: `GeneracijaID` za otkup/otpremnicu/zbirnu/prijemnicu,
`FakturaID`, `NovacID`. Revers i izvod već idu uz smer, odnosno broj računa.

### Lanac

`StornoRedF8` → `StornoRazlog` / `StornoTraziIzborModa` / `StornoPregledLanca` /
`StornoIzvrsi` / `StornoIzvrsiMod` → `Scan*` i writeri.

`PkPoIdentitetu` zamenjuje `LookupActiveID(... broj ...)` u sva tri `Scan*`:
generacija bira dokument; bez nje se pada na broj **tek pošto se dokaže da broj
nosi jednog vlasnika**, inače prazno → `exists=False` i flow staje. Writeri
(`StornoOtkupByBrDok_TX`, `StornoOtpremnicaByBroj_TX`, `StornoPrijemnicaByBroj_TX`)
dobili su opcionu generaciju i biraju redove po njoj.

### Test je prvo bio placebo

Prva verzija je birala dokument i poredila `OldDocID` sa njim — i **prolazila je
i kad se identitet potpuno ignoriše**, jer je razrešavanje po broju slučajno
davalo baš taj dokument. Sabotaža je to pokazala.

Prepravljen da meri **razliku u ponašanju**, ne konkretan PK: bez identiteta se
nad dvosmislenim brojem recovery zapis **ne pravi**, sa identitetom se pravi i
pokazuje na izabran dokument. Prva tvrdnja pada čim se identitet zaobiđe, bez
obzira na redosled redova.

```
FAIL … bez identiteta se NE pravi recovery zapis nad dvosmislenim brojem
       -- ocekivano [0], dobijeno [9]        (f8-identitet-po-broju)
```

### Sitno, iz istog pregleda

- `Err.description` se u `StornoRazlog` i `StornoIzvrsi` čita **pre** `LogErr`-a:
  `LogError` ima `On Error Resume Next` i fajl I/O, pa greška u samom logovanju
  prepiše `Err` i operater vidi pogrešnu poruku.
- `docs/EXCEL_TEST_HARNESS.md`: izričita napomena da je `--out` obavezan kad je
  donor postojeći fixture — komanda bez njega je već dvaput napisana kao da radi.

### Verifikacija

- `python tools\vba_check.py` → **čisto (190 fajlova)**, exit 0.
- `python tools\run_vba.py --suite RunAllTests` → **TESTS=35, FAIL=0**.
- `python tools\run_vba.py` (pun set) → **`EXIT=0`**, 11 suite-ova zeleno.
- `COMPILE` → **`NEJASNO`**. Posle ovog PR-a ručni `Alt+F11 → Debug → Compile
  VBAProject` je prava kapija, ne formalnost — u prethodnom rebase-u se pokazalo
  da statički checker propušta nedefinisan simbol i pogrešnu arnost.

## v2.46.1 — `v6-ui-130` · identitet do kraja lanca, i gde ne može

Dopuna v2.46.0 po pregledu. Identitet je stizao do prve putanje, pa sam ga
proglasio provučenim. **Nije bio** — `REŠI KASNIJE` je bio identity-aware, a
`ISPRAVKA`, `DUPLI` i deo `PONIŠTENJA` su se vraćali na broj.

### Preflight je primao identitet pa ga ignorisao

`StornoRazlog` je dobio `docID` i nije ga koristio. Za novac je i dalje zvao
`ResolveNovacForStorno(broj)`, koji kod dva aktivna reda istog broja kaže „treba
`NovacID`" — **iako mu je F8 `NovacID` upravo poslao**. `StornoIzvrsi` niže je
već bio ispravan, ali se do njega nije stizalo: kapija iznad je zaustavljala
operaciju.

To je poučan oblik greške — popravka jednog sloja bez drugog izgleda kao da radi,
jer je donji sloj tačan.

### Šta je sada identity-aware

| Putanja | Bilo | Sada |
|---|---|---|
| `StornoRazlog` (svih 6 tipova) | uvek po broju | `AktivanPoIdentitetu` |
| prijemnica ISPRAVKA/DUPLI/PONIŠTENJE | `…ByBroj_TX(broj)` | `(broj, docID)` |
| otpremnica ISPRAVKA/DUPLI | `…Atomic_TX(broj)` | `(broj, gen)`, `GetOtpremnicaIDsByBroj` filtrira |
| zbirna — **zaglavlje** | po broju | `StornoZbirna_TX(broj, gen)` |

Kapije nad brojem su sada **uslovne**: kad je identitet poznat, ne primenjuju se.
Bez toga su obarale potpuno legitimnu operaciju — storno je bio bezbedan, ali
funkcija nije radila.

### Zbirna: granica koja se ne može preći

Otpremnice, prijemnice i paletne stavke vezuju zbirnu **kolonom `BrojZbirne`** —
`ZbirnaID` im nije strani ključ nigde u šemi. **Deca dva dokumenta istog broja su
nerazlučiva podatkom koji postoji.**

Zato: zaglavlje se stornira po generaciji (tačno), a putanje koje bi menjale decu
(`PonistiZbirnaChain_TX`, `StornoZbirnaIDetach_TX`) **staju** kad broj nose dve
aktivne zbirne. To nije previd nego jedina poštena opcija — i tako je zapisano u
kodu, ne samo ovde.

### Verifikacija

- `python tools\vba_check.py` → **čisto (190 fajlova)**, exit 0.
- `python tools\run_vba.py --suite RunAllTests` → **TESTS=38, FAIL=0**.
- `python tools\run_vba.py` (pun set) → **`EXIT=0`**, 11 suite-ova zeleno.
- Tri nove sabotaže, svaka obara svoju tvrdnju:
  - `preflight-ignorise-id` → „sa NovacID-em preflight propušta izabran red"
  - `kapija-i-uz-identitet` → „ISPRAVKA pod kolizijom broja PROLAZI…"
  - `zbirna-zaglavlje-po-broju` → „zbirna drugog vozača istog broja OSTAJE aktivna"
- `COMPILE` → **`NEJASNO`**.

### Dva compile pada koje headless harness nije video

`IsStorniranoValue` je `Private` u `modStorno`, a pozvao sam je iz `modStornoDok`;
i `RunSimpleStornoZbirna` je dobio `docID` u telu bez parametra u potpisu. Oba su
prošla `vba_check` — pozivi su u izrazu, a to je rupa svesno ostavljena u #199.
Oba je pokazao **screenshot VBE-a**, ne harness.

## v2.46.2 — `v6-ui-131` · ispravka obrazloženja: broj zbirne JESTE jedinstven

Ne menja ponašanje. Ispravlja **tvrdnju** na kojoj je deo prethodnog rada
obrazložen — a pogrešno obrazloženje je gore od suvišne provere, jer sledeći
čovek na osnovu njega donese isti zaključak.

### Šta je bilo pogrešno

Pisao sam da se „broj zbirne generiše po vozaču, pa ga dva dokumenta lako dele".
**Ne dele ga.** `SuggestNextBroj` za `KIND_ZBR` ima eksplicitnu petlju:

```vb
Do While BrojZbirneExists(SuggestNextBroj)
    nextSeq = nextSeq + 1
    SuggestNextBroj = ApplyMirrorPrefix(entityID, FormatBroj(entityID, datum, nextSeq))
Loop
```

`BrojZbirneExists` skenira **celu** `tblZbirna`, bez opsega po vozaču, a
`ApplyMirrorPrefix` dodaje `S` baš da se mirror-vozač ne sudari sa realnim.
Generator ne može da izda zauzet broj.

### Zašto prijemnica jeste drugačija

```vb
maxSeq = MaxSeqFromTable(TBL_PRIJEMNICA, ..., COL_PRJ_KUPAC, kupacID, datum)
GenerateBrojPrijemnice = FormatBroj("1", datum, maxSeq + 1)
```

Fiksan prefiks `"1"`, sekvenca **po kupcu**, i **nema provere jedinstvenosti**.
Uz to auto-broj postoji samo za hladnjaču — ostali kupci unose slobodno. Tu je
kolizija stvarna, i tamo identitet nije pojas nego nužnost.

Grešku sam napravio izvodeći pravilo iz **oblika broja** umesto iz generatora.

### Šta ostaje i zašto

Zbirna zadržava `generacijaID` i fail-closed kaskade: broj koji generator drži
jedinstvenim i dalje može ući **mimo generatora** — ručnim unosom kad je
auto-broj isključen u Podešavanjima, uvozom, ili ispravkom u tabeli. Zaštita od
toga ne škodi.

Ispravljena su obrazloženja u `modStorno`, `modStornoFlow`, `modScrDokumenti`,
`modScrOporavak`, `modDokumenta`, `modTest`, `make_fixture.py` i katalogu — svuda
gde je stajalo „po vozaču, pa se dele".

Fixture `ZBI-DUPL-1/2` i test 38 ostaju, ali sada kažu šta stvarno brane:
**ručni unos**, ne redovan tok.

## v2.46.3 — `v6-ui-132` · identitet i na putanjama koje su ostale

Prethodni commit je tvrdio da je „svih pet nalaza zatvoreno". **Nije bilo** —
zatvorena je prva putanja svakog nalaza. Ovo zatvara ostale.

### P1 — otkup bez generacije je mogao da stornira dva otkupna mesta

`BrojDokumenta` otkupa je scoped **po otkupnom mestu**, pa isti broj na dva OM-a
postoji legitimno. Writer je bez generacije skupljao **sve** aktivne redove tog
broja. Prijemnica je taj obrazac već imala; otkup nije.

### P1 — „jedini vlasnik" zbirne merio se distinct brojevima

`OtpremnicaIsSoleOwner` je pitao `DistinctActiveValues(COL_OTP_BROJ)`. Zbirna je
po invarijanti zbir **svih** svojih aktivnih otpremnica, a broj otpremnice je
scoped po stanici — pa dve otpremnice istog broja sa različitih stanica u istoj
zbirni daju **jedan** distinct broj. Odgovor je bio „jedini vlasnik", `PONIŠTENJE`
izabrane je ulazilo u punu kaskadu i obaralo i tuđu.

Sada se broje **logički dokumenti** (generacija, PK kao fallback), i „jedini
vlasnik" znači: tačno jedan aktivan dokument, i to baš izabrani.

### P1 — završetak ispravke otpremnice vraćao se na broj

`CompleteOtpremnicaIspravka` je imao `docID`, ali ga pravi pozivaoci
(`modDokUnos`, `frmDokumenta`) nisu slali. Umesto da to traži od njih, completion
sada **izvodi identitet iz `correctionID`**: context nosi `OldDocID`, iz njega se
čita generacija stare otpremnice. Persistentan je i preživljava restart Excela.

Cilj (upravo snimljena zamena) dobio je kapiju nad jednoznačnošću broja — nova
otpremnica još nema generaciju u contextu, pa je to najuža provera koja se tu
može postaviti.

### P2

- `StornoIzvrsi` je bacao `docID` neposredno pre `StornoZbirna_TX`.
- `StornoOtpremnicaByBroj_TX` je imao **bezuslovnu** kapiju nad brojem — odbijao
  je tačno zadat `GEN-A` samo zato što `GEN-B` iste oznake postoji na drugoj stanici.
- `PkPoIdentitetu` je primao jednu kolonu vlasnika, dok je `ScanZbirna` odmah
  zatim merio dvosmislenost sa `VozacID + KupacID`. Sada prima niz.
- `RedJeGeneracije` je kod zadate generacije bez kolone **tiho** padao na broj;
  sada diže grešku, isto kao `RedJeIzabranogDokumenta`.

### Ispravka jedne prejake tvrdnje

Napisao sam da se deca zbirne „ne mogu razdvojiti ni u principu". Netačno:
otpremnica kaskada već ume `BrojZbirne + VozacID`, prijemnica `+ KupacID`, palete
nose `PrijemnicaID`. Scope se **može** izvesti — child mutacije samo još nisu sve
dovedene dotle. Fail-closed ostaje kao bezbedan izbor, ali više nije opisan kao
nemogućnost.

### Verifikacija

- `python tools\vba_check.py` → **čisto (190 fajlova)**.
- `python tools\run_vba.py --suite RunAllTests` → **TESTS=41, FAIL=0**.
- `python tools\run_vba.py` (pun set) → **`EXIT=0`**, 11 suite-ova zeleno.
- Tri nove sabotaže, svaka obara svoju tvrdnju.
- `COMPILE` → **`NEJASNO`**.

**Test 41 je prvo bio placebo** i to je vredno zapisati: tvrdio je da kaskada
staje, a stajala je od **zatečene** kapije u `StornoZbirna` — sabotaža moje
provere ništa nije menjala. Prepravljen da meri ono što moja provera stvarno
dodaje: **razlog** koji stiže do operatera umesto generičkog „nije uspelo".

## v2.46.4 — `v6-ui-133` · completion sloj: gde se identitet gubio na kraju

Dva P1 su bila u **completion** putanjama — onome što se izvršava tek **posle**
snimanja zamenskog dokumenta. Tamo testova nije bilo, i to nije slučajnost:
početak operacije je izgledao ispravno, pa se dalje nije gledalo.

### P1 — zamena zbirne je mogla da odnese decu tuđe zbirne

Početak ISPRAVKE je tačan: `StornoZbirna_TX(broj, docID)` stornira samo izabrano
zaglavlje, tuđe ostaje aktivno. Ali `CompleteZbirnaIspravka` — koja ide tek posle
snimanja zamene — prevezuje otpremnice i prijemnice **po `BrojZbirne`**.

Ishod: storniram tačno **svoje** zaglavlje, pa **tuđoj** zbirni odnesem decu.
Ništa ne izgleda pokvareno u trenutku storna.

Dok child mutacije ne budu scoped, modovi koji diraju decu (`ISPRAVKA`, `DUPLI`,
`PONIŠTENJE`) **staju pre nego što se išta promeni**. `REŠI KASNIJE` prolazi —
on ne dira decu.

### P1 — završetak ispravke otpremnice degradirao je tačan `OldDocID`

Zatečen dokument nema `GeneracijaID`. Completion je iz contexta čitao `OldDocID`,
`GeneracijaPoID` je vraćao `""`, i prazan opseg je značio **„izaberi po
poslovnom broju"**. Broj otpremnice je scoped po stanici, pa su blokovi dokumenta
sa **druge stanice** ulazili u relink.

`OldDocID` je bio tačan sve vreme — gubio se jedan korak kasnije. Sada se, kad
generacije nema, čita `StanicaID` baš tog `OldDocID` i opseg je **broj + stanica**.

### P2

Dvosmislen cilj u `CompleteOtpremnicaIspravka` sada ide u **`MANUAL`**, ne u tiho
`PENDING` — inače sledeći unos otpremnice ponovo pokreće pitanje „je li ovo
zamena?".

Uz to, **razlog iz kaskade sada stiže do operatera** u obe grane (zbirna i
prijemnica) umesto generičkog „nije uspelo (kaskada)".

### Testovi 42 i 43 — tamo gde ih nije bilo

- **42** — dve aktivne zbirne istog broja, svaka sa decom → ISPRAVKA staje,
  forma za zamenu se **ne otvara**, nijedno zaglavlje nije stornirano, tuđa
  otpremnica ostaje na svojoj zbirni.
- **43** — zatečen par bez generacije na dve stanice → completion prevezuje samo
  blok svog dokumenta.

Sabotaža za 43 reprodukuje kvar doslovno: blok `OTK-LEG-B` završi na zamenskoj
otpremnici `OTP-LEG-N` umesto da ostane na `OTP-LEG-B`.

### Test 41 je premešten, jer ga je nova kapija učinila nedostižnim

Kapija na nivou moda staje pre kaskadne, pa sabotaža kaskade više nije obarala
test 41. Umesto da ga ostavim kao ukras, premešten je na **PONIŠTENJE
prijemnice** — jedini put koji kaskadnu kapiju stvarno dohvata, jer ide nad
roditeljskom zbirnom.

### Verifikacija

- `python tools\vba_check.py` → **čisto (190 fajlova)**.
- `python tools\run_vba.py --suite RunAllTests` → **TESTS=43, FAIL=0**.
- `python tools\run_vba.py` (pun set) → **`EXIT=0`**, 11 suite-ova zeleno.
- Pet sabotaža iz poslednje dve runde, **svaka obara svoju tvrdnju**.
- `COMPILE` → **`NEJASNO`**.

Usput nađen i ispravljen poziv `MarkCorrectionManual` sa **četiri argumenta za
tri parametra**. Suite je bio zelen — VBA ne kompajlira modul dok ga ne dotakne.
Preostalih 20 poziva iste rutine je prebrojano mehanički.

## v2.46.5 — `v6-ui-134` · storniran vlasnik i dalje ima aktivnu decu

### P1 — kapija je brojala samo AKTIVNE vlasnike

`StornoZbirna_TX` stornira **samo redove `tblZbirna`** — otpremnice, prijemnice i
palete ne dira. Zato je ovo dostižno stanje, ne teorija:

```
Zbirna A  broj Z-10  STORNIRANA   ali OTP-A i PRJ-A još AKTIVNI
Zbirna B  broj Z-10  AKTIVNA
```

Sa brojanjem samo aktivnih vlasnika, izbor B daje „broj je jednoznačan" — pa
`DetachOtpremniceInline` i kaskada, koje idu **po broju**, odvežu i decu
stornirane A. **Storniran vlasnik nestaje iz računa, njegova deca ne.**

Sve tri kapije koje rade child mutaciju sada broje i stornirane vlasnike:
guard na nivou moda, `StornoZbirnaIDetach_TX` i `PonistiZbirnaChain_TX`.
Konzervativnije nego što je nužno — u skladu sa strategijom „fail-closed dok se
deca ne scope-uju po owneru".

### P2 — `stanicaID` opseg je bio fail-open na schema drift

`Or cSta = 0` je značilo: kolone nema → **propusti sve stanice**. Tačno suprotno
od razloga zbog kog opseg postoji. Sada diže grešku.

### P2 — test 43 dobio pozitivnu kontrolu

Tvrdio je samo „tuđ blok nije pomeren" — što prolazi i kod verzije koja **ne
preveže nijedan** blok. Sada tvrdi oba smera: moj blok **jeste** prevezan na
zamensku otpremnicu, tuđi **nije**. Sabotaža `completion-ne-prevezuje` obara
pozitivnu polovinu.

### Test 44 i jedna stvar koju je otkrio

Test počinje od storniranog zaglavlja A sa aktivnim detetom, pa traži da `DUPLI`
nad B stane.

Prve tri sabotaže **nisu ugrizle**, i razlog je vredan zapisa: ishod čuvaju
**dve nezavisne kapije** (na nivou moda i u detach-u), pa ga nijedna pojedinačna
sabotaža ne može oboriti. To je dobra odbrana, ali test koji tvrdi samo ishod ne
može da pokaže koja kapija radi.

Zato test sada tvrdi i **koja** je stala: kapija na nivou moda staje **pre
transakcije** i objašnjava razlog, dok bi detach pukao iznutra i dao samo „Storno
zbirne nije uspeo". Isto rešenje kao kod testa 41.

### Verifikacija

- `python tools\vba_check.py` → **čisto (190 fajlova)**.
- `python tools\run_vba.py --suite RunAllTests` → **TESTS=44, FAIL=0**.
- `python tools\run_vba.py` (pun set) → **`EXIT=0`**, 11 suite-ova zeleno.
- Sedam sabotaža iz poslednje tri runde, **svaka obara svoju tvrdnju**.
- `COMPILE` → **`NEJASNO`** — ostaje ručna kapija pred merge.

## v2.46.6 — `v6-ui-135` · kapija koja se sama guta nije kapija

### P2 blocker — `Err.Raise` u funkciji koja sve guta

Prethodna runda je u `GetOtpremnicaIDsByBroj` dodala fail-closed proveru: zadat
opseg stanice a kolone nema → greška. Ali ista funkcija završava sa:

```vb
EH:
    LogErr MOD_NAME & ".GetOtpremnicaIDsByBroj"
End Function
```

Dakle greška se digne, EH je proguta, pozivalac dobije **praznu kolekciju**,
petlja se preskoči — i completion završi kao **uspeh nad neprevezanim
blokovima**, uz `COMPLETED` context. Zaštita je postojala samo na papiru.

Sada se `Err.Number` / `Description` / `Source` sačuvaju, loguju i **ponovo
dignu**.

### Invarijanta kod pozivaoca

Context tvrdi da stari dokument postoji. Nula razrešenih ID-eva zato nije prazan
posao nego **nerazrešen izvor** — i `CompleteOtpremnicaIspravka` sada tu staje.
Štiti i buduće greške resolvera, ne samo nedostajuću kolonu.

### P3 — poruka je protivrečila tabeli

Kapije od `v6-ui-134` broje i **stornirane** vlasnike, a poruke su i dalje
govorile „nose dva **aktivna** dokumenta". Test 44 upravo dokazuje suprotan
slučaj: A stornirana, B aktivna, operacija svejedno staje. Operater bi dobio
poruku koja protivreči tabeli koju gleda.

Sve tri poruke sada kažu da je broj **pripadao više vlasnika** i da storniran
vlasnik i dalje može imati aktivnu decu.

### Verifikacija

- `python tools\vba_check.py` → **čisto (190 fajlova)**.
- `python tools\run_vba.py --suite RunAllTests` → **TESTS=44, FAIL=0**.
- `python tools\run_vba.py` (pun set) → **`EXIT=0`**, 11 suite-ova zeleno.
- **Sedam sabotaža** iz poslednje četiri runde, svaka obara svoju tvrdnju.
- `COMPILE` → **`NEJASNO`** — ostaje ručna kapija.

### Dve sabotaže su usput bile mrtve, i to se skoro nije videlo

Izmena poruka je zastarela sidra dvema sabotažama. `sabotaza.py` to prijavljuje
glasno („sidro nadjeno 0 puta"), ali sam u sweep-u imao `2>&1 >/dev/null` — pa je
izgledalo kao da sabotaža ne grize, umesto da nije ni primenjena.

**Kad se menja tekst poruke koja je deo sidra, sidro se menja s njom.** Sweep
odsad ne guta `stderr`.

## v2.46.7 — `v6-ui-136` · roditeljska zbirna je bila poslednji broj u lancu

Identitet je dotad bio rešen za **sam** dokument koji se stornira. Otpremnica ima
roditelja, i taj roditelj se sve vreme mutirao **po golom `BrojZbirne`**.

### P1 — mutacija roditelja po broju

`RunSimpleStornoOtpremnica`, `RunOtpremnicaCorrection` i
`CompleteOtpremnicaIspravka` sve tri dohvataju roditeljsku zbirnu po broju, pa
nad njom rade rekalkulaciju, storno prazne zbirne i relink prijemnica.

`RecalculateZbirnaFromOtpremnice_TX` je najgori slučaj: sabere otpremnice po
broju, pa tim zbirom ažurira **jedan** nađen red. Nad dvosmislenim brojem to
znači zaglavlje jednog dokumenta ažurirano zbirom otpremnica **oba**.

Nova kapija `ZbirnaBrojJeDvosmislenIkad(broj)` stoji na četiri mesta: pred
običnim stornom, pred svim modovima ispravke osim `RESI_KASNIJE`, u završetku
ispravke, i kao poslednja odbrana u `RecalcOrStornoEmptyZbirna_TX`.

### P2 — obična F8 putanja je i dalje ispuštala `docID`

Prosti storno otpremnice je zvao `StornoOtpremnicaByBroj_TX` bez identiteta —
isti oblik greške koji je već dva puta zatvaran na drugim tipovima.

### P2 — `Err.Description` posle `LogErr`

`LogErr` interno diže sopstveni EH i time briše `Err`. Novi re-raise je stizao u
blok koji je opis čitao **posle** `LogErr` — pa bi poruka bila prazna. Opis se
sada čita pre.

### Zatečen context preživljava upgrade

Kapija na startu ne pomaže za correction context napravljen **pre** nje.
Context je persistentan; posle upgrade-a bi završetak ispravke prošao bez ijedne
provere. Zato završetak pita ponovo, a ne veruje da je start pitao.

### Verifikacija

- `python tools\vba_check.py` → **čisto (190 fajlova)**.
- `python tools\run_vba.py --suite RunAllTests` → **TESTS=46, FAIL=0**.
- `python tools\run_vba.py --all` → **11 suite-ova punog seta zeleno**.
- **Dve nove sabotaže** (`otpremnica-bez-kapije-nad-zbirnom`,
  `zatecen-context-bez-kapije`), svaka obara svoju tvrdnju i samo svoj test.
- `COMPILE` → **`NEJASNO`** — ostaje ručna kapija.

### Dve sync suite padaju, i to nije od ovoga

`RunGoogleSyncSmokeSuite` (4/81) i `RunMasterSyncSmokeSuite` (9/26) padaju u
`--all`. **Nisu u punom setu** (`default: False`) i traže mrežu. Provereno na
worktree-u nad čistim `main`-om: identični brojevi padova, bez ijedne izmene sa
grane. Zatečeno stanje, prijavljeno kao zatečeno — ne kao zeleno.

## v2.46.8 — `v6-ui-137` · jedna linija je poništavala celu prethodnu rundu

Kapija iz `v6-ui-136` je proveravala **drugu zbirnu** od one koju kod mutira.

### P1 — roditelj se opet tražio po poslovnom broju

```
staraZbirna = LookupValue(TBL_OTPREMNICA, COL_OTP_BROJ, oldBroj, COL_OTP_BROJ_ZBIRNE)
```

To je tačno obrazac koji ceo PR uklanja: *imam identitet → odbacim ga → ponovo
biram prvi red po poslovnom broju*. Kapija je gledala `staraZbirna`, a relink
prijemnica, rekalkulacija i storno prazne zbirne su išli nad `oldZbirna` iz
context-a. Dve različite promenljive — pa je kapija mogla proveriti jednoznačnu
zbirnu **siblinga** i pustiti mutaciju nad dvosmislenom zbirnom izabranog
dokumenta.

Sada postoji **jedna** promenljiva: `ParentBroj` iz context-a → fallback
isključivo preko tačnog `OldDocID` → inače MANUAL. Nerazrešen roditelj se
razlikuje od **nepostojećeg**: otpremnica bez zbirne nema šta da mutira i to nije
greška, a nestao red jeste.

### P2 — kapija je bila fail-open na sopstvenu grešku

`On Error Resume Next` je pod schema drift-om davao `False`, to jest „broj je
jednoznačan, mutiraj" — baš kad se ništa ne zna. Sada `EH:` vraća `True`: za
kapiju je „ne mogu da dokažem jednoznačnost" isto što i „ne mutiraj".

### P2 — test 46 nije reprodukovao stvarni kvar

Stari test 46 je pravio context bez `ParentBroj`, pa completion nije ni imao
roditelja preko kog bi prevezao prijemnicu. Sabotaža je obarala uzgrednu tvrdnju
o `success`-u, a poslovnu nije ni doticala.

Nov scenario: dve otpremnice **istog broja** sa **različitim** roditeljima
(`OTP-STL-A` → dvosmislena, `OTP-STL-B` → jednoznačna, i B je prvi red), namenska
ciljna zbirna i namenska tuđa prijemnica. Sabotaža sada pokazuje **štetu**:

```
ocekivano [ZB-TEST-KASK], dobijeno [ZB-TEST-STL]
```

Prijemnica drugog dokumenta prevezana na zbirnu koja joj ne pripada.

### Test 47 — kapija mora blokirati kad ne može da dokaže

Drift se pravi stvarno (preimenovanje `VozacID` u `tblZbirna`), pa se meri kroz
seam: nad zdravom šemom `False`, pod driftom `True`. Pozitivna kontrola postoji
jer bi test inače prošao i sa kapijom koja blokira sve.

### Šesta zamka sabotaže: `AssertEq` diže grešku

Test se **prekida** na prvom padu, pa tvrdnje ispod ostaju neizvršene. Sabotaža
koja obori uzgrednu tvrdnju ostavlja poslovnu **nemerenom**, a to izgleda kao
uspešan dvosmerni dokaz. **Redosled tvrdnji je deo dokaza — najvažnija ide prva.**
Zapisano u `tools/sabotaza.py`.

### Verifikacija

- `python tools\vba_check.py` → **čisto (190 fajlova)**.
- `python tools\run_vba.py --suite RunAllTests` → **TESTS=47, FAIL=0**.
- `python tools\run_vba.py --all` → **11 suite-ova punog seta zeleno**.
- **Četiri sabotaže** nad ovim područjem, svaka obara **svoju** tvrdnju.
- Dve sync suite (van punog seta, traže mrežu) padaju identično i na čistom
  `main`-u — zatečeno.
- `COMPILE` → **`NEJASNO`** — ostaje ručna kapija.

## v2.46.9 — `v6-ui-138` · zaštita je bila jednostrana: izvor da, cilj ne

Kapija iz `v6-ui-137` je stajala samo nad **starom** zbirnom. Ciljna nije imala
nijednu — a nizvodne operacije nad ciljem idu po golom broju:
`ReassignPrijemnicaToZbirna_TX`, `RecalculateZbirnaFromOtpremnice_TX`,
`ValidateZbirnaInvariant`.

### P1 — zatečena kapija u writeru ovo ne pokriva

`ReassignPrijemnicaToZbirna_TX` bez generacije zove
`RequireJedanVlasnikPoBroju`, a taj broji **samo aktivne** vlasnike. Kad je jedan
vlasnik broja storniran a njegovo dete aktivno (upravo ono što test 44 dokazuje),
writer vidi jednog vlasnika i pusti relink.

Posle toga `SumOtpremniceByKlasa` sabira po broju, bez ownera — pa aktivno
zaglavlje dobije zbir dece **oba** vlasnika. Sabotaža to pokazuje kao broj:

```
ocekivano [100], dobijeno [400]
```

100 je količina njegovog jedinog deteta, 400 je 300 (dete storniranog vlasnika) +
100. Kapija sada stoji i nad `newZbirna`, **pre relinka blokova** — inače se
blokovi prevežu pa se tek onda otkrije da ostatak ne može bezbedno da se završi.

### Ista rupa je bila i u ispravci zbirne, na obe strane

`CompleteZbirnaIspravka` nije proveravala ni izvor ni cilj, a po broju ide sve:
`RelinkOtpremniceToZbirna_TX(oldBroj, newBroj)`, `DistinctActiveValues` po
`oldBroj`, relink prijemnica na `newBroj`, rekalkulacija `newBroj`. Dvosmislen
cilj znači „čije zaglavlje dobija zbir", dvosmislen izvor „čija deca se sele".
Obe grane su sada zatvorene i obe imaju sabotažu: cilj `100 → 500`, izvor —
otpremnica tuđeg dokumenta odseljena sa dvosmislenog broja.

### Zašto je ovo bilo najgore od svega

Invarijanta je i sama po broju. U pokvarenom stanju bi oba iznosa bila 400 i
`ValidateZbirnaInvariant` bi rekla **ISPRAVNO**. Test 48 to fiksira kao tvrdnju:
u zdravom stanju invarijanta kaže *neispravno* za `ZB-TEST-TGT`, jer sabira decu
oba vlasnika protiv jednog zaglavlja. Validacija koja potvrđuje kontaminaciju je
gora od validacije koje nema.

### Zatečena provera koja štiti slučajno, ne po pravilu

Sabotaža je prvo izgledala kao da ne grize. Uzrok: `ReassignPrijemnicaToZbirna_TX`
bez generacije čita `Stornirano` **prvog reda po broju** — a u prvoj verziji
fixture-a je prvi red slučajno bio stornirani vlasnik, pa je relink odbijen. Ne
zato što proverava vlasništvo, nego zato što je prvi red slučajno bio storniran.
Redosled redova u fixture-u je zato deo scenarija i tako je i zapisan.

To je i mesto gde bi jednog dana trebala **centralna** kapija: ovaj PR štiti svoje
call-site-ove, ali sam primitiv ostaje number-based.

### Verifikacija

- `python tools\vba_check.py` → **čisto (190 fajlova)**.
- `python tools\run_vba.py --suite RunAllTests` → **TESTS=49, FAIL=0**.
- `python tools\run_vba.py --all` → **12 suite-ova zeleno** (11 punog seta + SEF).
- **Tri nove sabotaže**, svaka obara svoju tvrdnju i pokazuje štetu brojem.
- `COMPILE` → **`NEJASNO`** — ostaje ručna kapija.

## v2.46.10 — `v6-ui-139` · `Err.Description` posle `LogErr`, svih deset mesta

`LogErr` interno zove `LogError`, a taj ima `On Error Resume Next` i fajl I/O — pa
briše `Err`. Svaki EH blok koji opis čita **posle** `LogErr`-a prijavljuje
`"Greska: "` i ništa više.

Prijavljeno je bilo jedno mesto (`CompleteZbirnaIspravka`). Mehaničkim skeniranjem
`modStornoFlow` ih je **deset**: četiri `RunSimpleStorno*`, četiri `Run*Correction`,
`CompleteZbirnaIspravka` i `CompleteReversIspravka`. Sva su ispravljena istim
obrascem — opis se čita u prvi red EH bloka, pre `LogErr`-a.

Sređivanje samo prijavljenog mesta bi ostavilo devet identičnih, a to je već tri
puta bio problem u ovom PR-u.

### Ovo NIJE pokriveno testom, i to je namerno

Do tih EH blokova se iz testa ne može deterministički doći: svaki sloj ispod
(`ZbirnaPostoji`, `RecalculateZbirnaFromOtpremnice_TX`, `ReassignPrijemnica…`) ima
svoj `On Error` i grešku vraća kao `False`, pa operacija stane u redovnoj putanji
sa svojom porukom. Test bi zahtevao raise-seam u produkcionom kodu — što je gore
od same ispravke.

Obrazac je **mehanički prepoznatljiv u izvoru** i tu mu je mesto: provera u
`vba_check` („`Err.description` čitan posle `LogErr`-a u istom EH bloku") ide kao
zaseban posao, kad se `tools/vba_check.py` oslobodi (menja ga #199).

### Verifikacija

- `python tools\vba_check.py` → **čisto (190 fajlova)**.
- `python tools\run_vba.py --all` → **12 suite-ova zeleno, TESTS=49 FAIL=0** —
  dokaz da ispravka ništa nije pokvarila, ne da je EH putanja izvršena.
- `COMPILE` → **`NEJASNO`** — ostaje ručna kapija.

## v2.46.11 — `v6-ui-140` · poslednja putanja koja je birala po broju: spisak blokova

Identitet je stigao do svakog dokumenta i do njihovih roditelja, ali **dodatni
storno otkupnih blokova** je i dalje birao po poslovnom broju.

### P1 — spisak blokova je nosio blokove tuđeg dokumenta

`ActiveBlocksForFlow` je za otpremnicu radila
`GetBlokOtkupIDs(GetOtpremnicaIDsByBroj(broj))` — bez generacije. Isti
`BrojOtpremnice` na dve stanice je legitiman, a `GetOtpremnicaIDsByBroj` namerno
uključuje i **stornirane** otpremnice, jer njihovi blokovi još mogu pokazivati na
njih.

Taj spisak nije pregled: iz njega se pravi `ids` i ide pravo u
`StornoSelectedBlocks_TX`.

### Zašto zatečena kapija tu nije pomagala

`BlockStornoDriftReason` počinje sa:

```vb
If ModeStornoBlokParent(docType, mode) Then Exit Function     ' roditelj umire -> ok
```

a `ModeStornoBlokParent` je `True` za **svaki** `PONISTENJE` i za
`OTPREMNICA + DUPLI/ISPRAVKA` — dakle za tačno one modove koji jedini i stižu do
dodatnog storna blokova. Kapija se ne izvršava. Njena pretpostavka („roditelj
umire, pa je blok-storno bezbedan") važi samo za blokove **izabranog** dokumenta.

Sabotaža pokazuje mutaciju, ne pregled:

```
FAIL T_StorniranSibling_ZadrzavaSvojBlok
     blok storniranog siblinga je ostao AKTIVAN
     ocekivano [False], dobijeno [True]
```

### Isti kvar u pregledu, na dva mesta

- `ScanOtpremnica` je razrešila dokument po identitetu pa `blockCount` računala po
  broju — pregled bi pokazao tuđe blokove i correction dijalog bi se otvorio i nad
  dokumentom koji blokove nema.
- `ScanPrijemnica` je imala tačan `prijID`, ali je `blockCount` išao kroz
  `ActiveBlocksForFlow(PRIJEMNICA, broj)`, a ta je roditeljsku zbirnu izvodila iz
  **prvog reda tog broja**. `BrojPrijemnice` nije globalno jedinstven (sekvenca po
  kupcu), pa je to bio verovatniji ulaz od otpremnice.

Sada: `StornirajBlokoveAko → GetStornoBlockRows → ActiveBlocksForFlow` nose
`docID`, prijemnica čita roditelja iz tačnog `prijID`, a `ScanOtpremnica` broji po
generaciji.

### Šta ostaje number-based, i zašto

Grana `FLOW_DOC_ZBIRNA` u `ActiveBlocksForFlow`: `tblOtkup` nosi denormalizovan
`BrojZbirne`, ne `ZbirnaID`, pa se deca po generaciji zbirne **ne mogu** razdvojiti.
Taj put je zaštićen uzvodno — kapije nad dvosmislenim brojem zbirne obore mode
operaciju, a dodatni storno blokova ide samo posle uspešne. Zapisano u kodu, jer
ako se te kapije jednog dana suze, mesto se otvara.

`modStornoImpact.BuildStornoImpact` (ekran „Uvid") i dalje zove
`GetStornoBlockRows` bez identiteta. To je read-only pregled i njegovi pozivaoci
identitet ne nose, pa nisam dodavao parametar koji niko ne prosleđuje.

### Verifikacija

- `python tools\vba_check.py` → **čisto (190 fajlova)**.
- `python tools\run_vba.py --suite RunAllTests` → **TESTS=51, FAIL=0**.
- `python tools\run_vba.py --all` → **12 suite-ova zeleno**.
- **Dve nove sabotaže**: `blokovi-po-broju` (mutacija) i `blockcount-po-broju`
  (pregled), svaka obara svoju tvrdnju.
- `COMPILE` → **`NEJASNO`** — ostaje ručna kapija.

## v2.46.12 — `v6-ui-141` · compile greška koja je živela pet PR-ova

Operater je ručnim `Debug → Compile VBAProject` našao ono što 51 zelen test nije:

```
Compile error: Expected array
    poruka = Poruka("STORNO_MSG_ZBIRNA_PRIJ") & " " & vezPrij
```

`StornoIzvrsi` ima izlazni parametar `ByRef poruka As String`. VBA je
case-insensitive, pa nekvalifikovan `Poruka("KLJUC")` unutar te procedure **nije
poziv funkcije nego indeksiranje tog String parametra**. Šest poziva u istoj
proceduri, svi pogrešni od `v6-ui-119` (#193).

### Zašto je pet PR-ova prošlo pored ovoga

**VBA kompajlira proceduru tek kad se pozove.** `StornoIzvrsi` je zvao samo UI
(`modScrDokumenti`), nijedan test — pa je ceo blok bio mrtav za suite-ove.
Statički ga ne vidi ni `vba_check`: poziv je u **poziciji izraza**
(`x = Foo(...)`), a to je dokumentovana rupa checkera (pokušaj proširenja je dao
406 lažnih nalaza).

Dakle: ni suite, ni checker, ni CI. Samo `Debug → Compile`.

### Ispravka

Svih šest poziva je sada **kvalifikovano** (`modPoruke.Poruka(...)`) i u kodu stoji
zašto to mora ostati tako.

Mehanički sweep nad celim `src-vba` (`.bas`, `.frm`, `.cls`) traži isti obrazac —
lokalna skalarna promenljiva ili parametar čije ime zaklanja `Public Function`, pa
se poziva sa zagradom i string literalom. Posle ispravke: **0 nalaza**. Devet
kandidata koje sweep prijavi bez tipa su lažni — niz i objekat se legitimno zovu sa
zagradom; compile obara samo skalar.

### Test 52 — procedura mora da se IZVRŠI, ne samo da postoji

`T_StornoIzvrsi_ZbirnaImenujeVezanuPrijemnicu` zove `StornoIzvrsi` za
`STIP_ZBIRNA` i traži da poruka imenuje prijemnicu koja je ostala vezana za
storniranu zbirnu (`StornoZbirna` namerno ne kaskadira).

Da test radi to što treba, dokazano je **slučajno i najuverljivije**: prva verzija
je puštena kad je bio ispravljen samo jedan od šest poziva, i suite je pao —
`SUITE FAIL RunAllTests (91.1s)`, dijalog `Compile error`. Isti kod je do tada bio
51/51 zelen.

Sabotaža (`zbirna-poruka-bez-prijemnice`) obara tvrdnju o poruci. Sama compile
greška se sabotažom **ne može** dokazati imenovanom tvrdnjom — takva sabotaža
obara compile, pa izlaz bude „Exception occurred" umesto imena tvrdnje (zamka 4).
To je i zapisano u katalogu, uz ono što test 52 stvarno dodaje: proceduru koja se
izvršava.

### Drugi nalaz iz istog Compile-a: `StornoRedF8`

Operater je posle prve ispravke pustio Compile ponovo i dobio **isto** u
`modScrDokumenti.StornoRedF8` — još šest poziva, isti obrazac. Ukupno dvanaest, u
dve procedure.

Moj prvi sweep ih nije prijavio zbog svoje greške: čitao je samo **prvu**
deklaraciju u `Dim` redu, a ovde je `poruka` druga:

```vb
Dim razlog As String, poruka As String, odg As VbMsgBoxResult
```

Ispravljen sweep je zatim dao 20 nalaza od kojih **14 lažnih** — ime unutar
**string literala** (`"...bez OtkupID (dokument: ..."`). To je tačno ona klasa
lažnih nalaza zbog koje je proširenje `ARNOST`-a odbijeno u #199. Posle skidanja
literala i komentara: **6 nalaza, svih 6 stvarnih, 0 lažnih.** Svi ispravljeni,
ponovni sweep daje **0**.

### Šta ovo znači za verifikaciju

Nijedan od dvanaest pogrešnih poziva nije bio dohvatljiv suite-om: `StornoIzvrsi` i
`StornoRedF8` zove **samo UI**, a VBA kompajlira proceduru tek kad se pozove. Test
52 pokriva prvu; `StornoRedF8` je `Private` i otvara `MsgBox`, pa se iz testa ne
može pozvati.

Zato ovaj obrazac ide u `vba_check` kao provera, kroz **zaseban process PR**.
Uslovi su ispunjeni: mehanički je prepoznatljiv, ima **0 lažnih nalaza** kad se
literali i komentari skinu, i dva puta je našao stvarne greške koje ništa drugo
nije videlo.

### Verifikacija

- `python tools\vba_check.py` → **čisto (190 fajlova)**.
- `python tools\run_vba.py --suite RunAllTests` → **TESTS=52, FAIL=0**.
- Sweep nad `src-vba` za zaklonjena imena → **0** (bio 12).
- `COMPILE` → i dalje `NEJASNO` iz runnera; **stvarna kapija je bila operaterova, dva puta.**

## v2.47.0 — `v6-ui-142` · keš koji je pamtio da podataka nema

Operater je prijavio dve stvari koje su izgledale nevezano: **prazne liste za
svaki tip dokumenta** i **`KOOP-00022` umesto imena kooperanta**. Koren je jedan.

### P1 — keš je memoisao neuspeh

```vb
If Not mCache.Exists(tblName) Then mCache(tblName) = GetTableData(tblName)
```

`GetTableData` vraća `Empty` i kad je tabela prazna. Ako je pri **prvom** čitanju
bila prazna — a podaci stižu posle, kroz sync, uvoz ili legacy formu — `Empty`
ostaje u kešu **do kraja sesije**. `ResetCache` se zove samo pri gradnji ekrana i
posle upisa kroz novi UI, a nijedan od tih puteva nije prošao.

Dijagnostika sa prave instalacije je to pokazala u jednom redu:

```
IsArray(modUiData.CachedTable("tblOtkup"))  ->  False
IsArray(GetTableData("tblOtkup"))           ->  True
GetColumnIndex("tblOtkup", "KooperantID")   ->  3      (nema schema drift-a)
```

Potvrda: ručni `modUiData.ResetCache: modOtkupUI.RefreshFromData` je vratio sve
liste.

**Sada se kešira samo uspeh.** Neuspeh je „ne znam još", ne „nema ništa" — pa se
sledeći poziv ponovo pita. Cena je jedan promašen sken po tabeli koja je zaista
prazna.

### P1 — tiho `Exit Function` je krilo razliku

```vb
src = modUiData.CachedTable(tblName)
If Not IsArray(src) Then Exit Function      ' bez ijedne reci
```

Vraćao je nedodeljen `Variant`, pa je mreža crtala praznu listu **bez greške** —
operater nije imao način da razlikuje „nema dokumenata" od „ne umem da pročitam
tabelu". Sada: prazna tabela prolazi kao prazna, **nepostojeća je greška sa imenom
tabele** (`modUiData.TabelaCitljiva`).

To je isti fail-open obrazac koji je devet rundi vađen iz storna, samo u sloju
prikaza.

### P2 — `KOOP-00022` umesto imena, dva razloga

`PartnerMap` čita isti keš. Kad dobije `Empty`, napravi **prazan** rečnik — i
onda ga **kešira**, pa svako ime do kraja sesije pada na goli ID. Prazna mapa se
sada ne kešira.

Uz to je ključ keša bio **samo ime tabele**, iako mapa zavisi i od kolona: prvi
pozivalac je time odlučivao šta svi ostali dobijaju (`"Ime"` bez `"Prezime"`).
Ključ sada nosi tabelu **i** kolone.

Zašto su liste „Kooperanti" i „Izgubljeni po kg" ipak radile: one čitaju
`GetTableData` **direktno**, mimo tog keša. Ta asimetrija je i bila trag.

### Storno više ne izgleda kao unos

F8 je crtao celu unosnu formu i primarno dugme „Storniraj dokument", a `Scr_Save`
za `STORNO` pada u `Case Else` i vraća „Nije vezano na postojeću rutinu" — dakle
**dugme je bilo mrtvo**, a forma je pozivala operatera da ukuca podatke dokumenta
koji hoće da stornira.

Ulazak u F8 sada uključuje grid-max: forma i kontekstni red se sklanjaju, mreža
dobija ceo prostor — isti raspored koji imaju Palete i Oporavak. Operaterov izbor
grid-maxa se pamti i vraća pri izlasku, a prekidač `⤢` je u F8 bez posla pa se
tiho ignoriše.

**Storno ostaje u F-traci**, nije izdvojen u svoj ekran. Izdvajanje je veći posao
(~800 linija kroz pet modula) i njegova oštra ivica je predaja `ISPRAVKA`/`DUPLI`
unosnom ekranu — tok koji je upravo stabilizovan. Vizuelni ishod je isti.

### Verifikacija

- `python tools\vba_check.py` → **čisto (190 fajlova)**.
- `python tools\run_vba.py --suite RunAllTests` → **TESTS=55, FAIL=0**.
- `python tools\run_vba.py --all` → **12 suite-ova zeleno**.
- **Četiri nove sabotaže**, svaka obara svoju tvrdnju.
- `COMPILE` → **`NEJASNO`** — ostaje ručna kapija.

Testovi 53–55 mere baš ono što se ne vidi kroz vraćenu vrednost: sam keš (kroz
seam), jer `Empty` izgleda isto i kad je neuspeh keširan i kad nije.


## v2.48.0 — `v6-ui-143` · Storno dobija svoj ekran

Korak 1 od dva. Storno prestaje da bude osmi režim unosnog ekrana i postaje
zaseban ekran u navigaciji, sa pregledom posledica **pre** odluke — ekvivalent
legacy ekrana „Storno / potvrda".

### Zašto

F8 je crtao unosnu formu koju ne koristi. `Scr_Save` za `STORNO` je padao u
`Case Else` i vraćao „Nije vezano na postojeću rutinu" — dakle **primarno dugme
je bilo mrtvo**, a forma je pozivala operatera da ukuca podatke dokumenta koji
hoće da stornira. `v6-ui-142` je to sakrio grid-maxom; forma je i dalje
postojala, samo se nije videla. Sada je nema.

Drugo: pregled posledica je bio niz `MsgBox`-ova. `modStornoImpact` vraća
**sedam sekcija** (zaglavlje, lanac, blokovi, zastavice, palete, faktura,
sažetak) — tačno sekcije sa legacy ekrana — ali ga je do sada renderovao **samo
legacy** (`frmDokumenta.frm:4662`). Novi UI ga nije zvao ni sa jednog mesta.

### Odluka koja se menja

Katalog (`UI_MIGRACIJA_KATALOG.md`, §3.2) je beležio: „četiri odgovora ne staju
u jedan `MsgBox`". Zapažanje je bilo tačno, zaključak nije. Sada su to **četiri
dugmeta** — Pogrešan unos / Duplikat / Ništa se nije desilo / Reši kasnije —
svako sa objašnjenjem ispod, a iznad njih stoje lanac, palete i broj blokova.
Operater vidi sva četiri ishoda **istovremeno**, umesto da drugi izbor otkrije
tek pošto odgovori na prvi. Prekidač „Ne diraj palete" je iz istog razloga izašao
iz `MsgBox`-a (`STORNO_ASK_PALETE`) i stoji uz palete na koje se odnosi.

### Najopasnija linija u migraciji

```vb
If modOtkupUI.ActiveMode = "F8" Then          ' <- do v6-ui-142
    c.Add "OTKUI_HD_IDENT|" & IdKolonaTipa(mk) & "|txt|0|4"
```

Nevidljiva kolona kanonskog identiteta dodavala se pod uslovom da je aktivan
režim F8. **Ekran nema režim.** Da je taj uslov ostao, bio bi ćutke `False`,
kolona bi nestala, `IdentIzReda` bi vraćao prazno i ceo lanac iz #198
(`correctionID` / `OldDocID` / `GeneracijaID`) pao bi na fail-closed po broju.

Nijedna postojeća suite to ne bi videla: testovi identiteta (35, 45, 46, 48–52)
mere sloj **ispod** mreže, kome se `docID` prosleđuje direktno. Zato je kapija
sada **argument** (`GridCols(tip, saIdentitetom)`), i zato test 57 meri baš taj
spoj — u oba smera: kolona mora biti tu kad se traži, i **ne sme** biti tu za
unosni režim.

### Šta je još u ovom koraku

- **`modScrStorno`** po ugovoru ekrana (`oblik=lista|upis=ne`), grupa OPERACIJE,
  oblast `OBL_DOKUMENTA`. Deset čipova: devet tipova iz F8 plus **navigacioni
  „Svi"**, koji legacy ima („Nađi dokument") a novi UI nije imao.
- **„Svi" je namerno samo navigacioni.** Izvor (`GetActiveDocumentsForStorno`)
  pokriva tri framework tipa i **nema kolonu identiteta**; radnja nad takvim
  redom vratila bi se na biranje po broju. Klik zato prebacuje na tipiziranu
  listu, gde identitet postoji, i odluka se donosi tamo.
- **`RedoviZaTip`**: `RowsDokumenti` (~250 linija) je bio `Private` i čitao
  `ActiveMode` na tri mesta. Sada prima **ključ tipa** kao argument — jedan
  čitač i za unosni ekran i za Storno, bez kopije. `EffKey` i `mStTip` su
  nestali s njim; `ModeTable` je sada samo `TabelaTipa(modeKey(mode))`.
- **`BuildStornoImpact` dobija `docID`** i provlači ga do `GetStornoChainRows` /
  `GetStornoBlockRows` / `GetChainFlags` (koji su ga primali od `v6-ui-136/140`
  — nedostajao je samo sloj koji ih spaja). Zaglavlje uvida ide kroz nov `HLI`:
  `LookupValue` vraća **prvi** red po broju, pa je pod kolizijom uvid opisivao
  dokument koji se ne stornira.
- **Predaja ISPRAVKA/DUPLI** je sada tri koraka a ne dva: `IdiNaEkran` →
  `IdiNaRezim` → `ApplyPrefill`. Redosled je obavezan — `SelectMode` **čisti**
  formu, pa bi prefill pre njega bio obrisan; a bez prvog koraka prefill upisuje
  u zonu koja je sakrivena. `ActivateScreen` je ostao `Private`; nov javni ulaz
  je `IdiNaEkran`.
- **Ljuska**: `rezima=8` → `7`, osma kartica iz `zRight`, taster F8 sada bira
  **ekran**, grid-max izuzetak iz `v6-ui-142` uklonjen (postao je bespredmetan),
  i nov generički prolaz za kontrole iz zone ugovornog ekrana (prefiks `scr`) —
  do sada nijedan ekran nije imao kliktaču kontrolu u zoni.

### Šta je našao smoke test (i šta je iz toga naučeno)

Pet nalaza, od kojih **tri nisu bila u Stornu nego u ljusci**. Svi imaju istu
osobinu: kvar bez ijedne poruke.

1. **`MAX_SEG = 9`, a ekran ima deset lista.** `LayoutGrid` nacrta prvih devet i
   stane — bez greške i bez traga — pa čip „Izvodi" **nije postojao**. Granica je
   podignuta na 10, a tvrdnja se sada meri kroz nov seam `MaxPrekidaca()`, i to
   tako da važi za **svaki budući ekran**.
2. **`ScrGridData` je gutao grešku iz `Scr_Rows` golim `On Error Resume Next`.**
   Greška postaje `Empty`, `LoadGridFromScreen` na ne-niz radi `Exit Sub`, i
   mreža ostane na **prethodnoj listi sa prethodnim naslovom**. Prekidač izgleda
   kao da ne radi. To je isti obrazac zbog kog `ScrEvent` već ima `ScrLastErr`
   („Po datumu"), samo u sloju redova — gušenje ostaje, trag se dodaje.
3. **Prefill je gubio baš ono što je lista.** Novi `scr` prolaz je posle radnje
   bezuslovno zvao `RefreshFromData`, a on radi `FillZbirneCombo` i
   `mPartnerFor = ""` — pa briše izbor koji je prefill upravo postavio. Radnja
   sada osvežava **samo ako je ostala na istom ekranu**.
4. **Broj dokumenta je posle ispravke ostajao prazan.** Prefill ga namerno ne
   donosi (stari broj pripada storniranom), ali predlog se ni **nije računao**:
   `RefreshBrojPredlog` visi o promeni stanice ili datuma, a `ApplyPrefill` oba
   postavlja pod `mLoading = True`. Računa se sada na kraju prefilla, i **samo**
   kad broj nije donet. Remote provera se preskače u test-režimu (isti gard kao
   kod `SetFocus`-a) — suite ne sme da zavisi od mreže.
5. **Keš je stajao oko pogrešnog dela.** `BeginTableCache` je obuhvatao samo
   `BuildStornoImpact`, a red odluke (`AkcijeZaTip` → `StornoTraziIzborModa` →
   `CorrectionNeedsDialog`) je **još jedan prolaz kroz isti lanac** i stajao je
   izvan keša. Baš to operater vidi kao „sporo se puni efekat po modu". Keš sada
   obuhvata ceo izbor reda, uz `EndTableCache` i u error handleru.

**Test 59 je prvo pao na PREDUSLOVU, ne na pravilu** — `ApplyPrefill` u testu
nije mogao da izabere stanicu jer combo-i nisu bili punjeni. Bez te tvrdnje bi
test merio prazan combo i „prošao" bi i nad neispravnim kodom.

**Nije defekt:** prazan KUPAC na otpremnici. `tblOtpremnica` nema kolonu kupca —
partner otpremnice je stanica.

### Šta NIJE u obimu (Korak 2)

Ponovo upotrebljiva tabela (ljuska ima **jednu** mrežu, pa lanac i blokovi idu u
zonu kao labele) i kolona sa checkbox-om (mreža bira **jedan** red, pa dodatni
storno blokova ostaje sve-ili-ništa). Legacy `frmDokumenta` se ne dira.

### Verifikacija

- `python tools\vba_check.py` → **čisto (191 fajl)**; `--self-test` → **29**.
- `python tools\run_vba.py --suite RunAllTests` → **TESTS=62, FAIL=0** (bilo 55).
- `python tools\run_vba.py --all` → **12 suite-ova OK**, dve sync suite
  (`RunGoogleSyncSmokeSuite`, `RunMasterSyncSmokeSuite`) crvene **zatečeno**.
### Šta je zatvorila recenzija PR-a

Recenzija je našla tri P1 — sva tri o istome: **novi ekran još nije imao
fail-closed standard koji je poslovni sloj već dostigao.**

1. **Uvid nije bio identity-aware u celosti.** `docID` je stizao do zaglavlja,
   lanca, blokova i zastavica, a `ImpactPalete` i `ImpactFaktura` su i dalje
   tražili po **broju**. Pod kolizijom broja to znači: zaglavlje pokazuje
   izabran dokument, palete pokazuju palete **oba** — a writer nizvodno mutira
   samo izabrani. Ekran koji obećava „ovo su posledice" tvrdio bi posledice koje
   se neće desiti.
   Sada: prijemnica se sužava kroz `tblPaletaStavka.PrijemnicaID` (skup ID-jeva,
   jer Klasa I i II dele generaciju a imaju različit PK), otpremnica kroz `HLI`
   nad svojom zbirnom, faktura kroz `HLI`. **Zbirna ostaje po broju** — stavke
   nose `BrojZbirne`, ne `ZbirnaID`; ista granica šeme kao kod `FLOW_DOC_ZBIRNA`
   u `ActiveBlocksForFlow`. To je prijavljeno kao granica, ne pokriveno
   pogađanjem.
2. **Ako uvid pukne, dugmad za mutaciju su se i dalje nudila.** `BuildStornoImpact`
   je rezultat dodeljivao **pre** nego što ga izgradi, pa je pad na pola davao
   parcijalan rečnik koji spolja izgleda ispravno; a red odluke se računao
   nezavisno od uvida. Model sada nosi `valid`, koji se postavlja tek na kraju, i
   `AkcijeRacun` je fail-closed: framework dokument bez valjanog uvida ne nudi
   **nijednu** radnju. Tip koji uvid nema po prirodi (revers, otkup, novac) time
   nije zaključan.
3. **Keš odluke je preživljavao promenu podataka.** Ključ je `tip|broj|docID` —
   dakle dokument, ne stanje podataka. `Scr_ResetCache` je brisao samo uvid, pa
   je posle sync-a važila odluka od pre njega: dokument koji je u 10:00 bio bez
   nizvodnog toka zadržao bi „običan storno" i pošto mu je sync doneo zbirnu,
   prijemnicu i palete. `StornoRazlog` to ne hvata — on pita sme li se dokument
   stornirati, ne da li sada treba framework ispravke. `Scr_ResetCache` sada čisti
   **ceo izbor**: posle promene podataka operater bira dokument ponovo.

Uz to, tri P2:

- **`Err.description` posle `LogErr`-a** se vratio u novi modul, iako je iz
  desetak mesta vađen u `d60b6706`. `LogErr` ima svoj `On Error Resume Next` i
  usput briše stanje greške, pa je poruka operateru ostajala prazna. Opis se sada
  čita **pre** loga, u svim handlerima; `IzborReda` grešku i **prikazuje**, umesto
  da klik na red tiho ne uradi ništa.
- **„Svi" nije bio svi.** Izvor (`GetActiveDocumentsForStorno`) pokriva tri
  framework tipa, a `TipIzNaziva` ume da razreši ista tri — otkup, novac, revers,
  faktura i izvod nisu tu. Naziv je obećavao globalnu pretragu. Lista se sada zove
  **„Lanac robe"** (`OTKUI_SEG_ST_LANAC`), što je tačno ono što pokriva.
- **Build pečati su lagali.** `OTKUI_BUILD`, `SCRDOK_BUILD` i `UISCR_BUILD` su
  ostali na `v6-ui-121` iako su ta tri modula najviše promenjena — a baš njih
  `OtkupUI_SelfCheck` prijavljuje na pitanje „da li je klijentu uvezena nova
  verzija". Podignuti su na `v6-ui-143`.

**Tri nova behavioral testa** (60–62) mere baš tu granicu, koju 59 dotadašnjih
nije diralo: nijedan nije ni pozivao `BuildStornoImpact`.

- **60** — dva dokumenta istog broja, različite palete: uvid izabranog sadrži
  isključivo njegove posledice (preduslov dokazuje da po broju vidi oba).
- **61** — odluka izračunata pre promene podataka ne sme da preživi `Scr_ResetCache`.
- **62** — framework dokument bez uvida ne nudi nijednu radnju; revers i otkup,
  koji uvid nemaju po prirodi, i dalje nude svoje.

### `valid=True` je sada stvaran ugovor

Druga runda recenzije je pokazala da je `valid` bio flag oko spoljnog `On Error`,
a ne ugovor: **više čitalaca je samo gutalo grešku**, kontrola bi se uredno
vratila u `BuildStornoImpact`, i on bi postavio `valid = True`.

```
ne mogu da procitam palete  ->  prazna Collection
                            ->  ImpactPalete misli da je sve OK
                            ->  valid = True
                            ->  ekran kaze "nema paleta" i nudi mutaciju
```

A tačan odgovor nije „nema paleta" nego **„ne znam da li ih ima"**.

Zatvoreno na dva pravila:

1. **Čitaocima uvida je gutanje zabranjeno.** `GetPaleteImpactByField` je dobio
   `strict` (podrazumevano `False`, pa zatečenim pozivaocima ponašanje ostaje
   isto); u strict režimu nedostajuća kolona, nečitljiva tabela i greška u
   prolazu **dižu** grešku. `SumActiveNum` i `HLI` isto. `ImpactHeader` i
   `ImpactFaktura` više ne gutaju — njihov `EH` **propušta** grešku dalje, jer bi
   inače pojeli baš ono što je ispod dignuto.
2. **Zadat identitet nikad ne degradira na broj.** Prazan `docID` i dalje sme da
   radi po broju — zatečen zapis nema generaciju, pa je broj sve što postoji. Ali
   `docID` koji je zadat a ne može da se razreši **obara uvid**; tihi povratak na
   broj vratio bi tačno ono što je #198 vadio, i to unutar modela koji se posle
   označava kao valjan.

**Zbirna i dalje ostaje po broju** — `tblPaletaStavka` nosi `BrojZbirne`, ne
`ZbirnaID`. To je granica šeme, prijavljena kao granica.

Dva nova testa mere sam ugovor, a ne njegovu posledicu:

- **63** — kolona `PrijemnicaID` se **stvarno** preimenuje (isti obrazac kao test
  47), pa se traži `valid = False`. Pozitivna kontrola pre toga dokazuje da nad
  zdravom šemom uvid jeste valjan i da paleta stvarno postoji — bez nje bi test
  prošao i kad `BuildStornoImpact` uvek vraća `False`.
- **64** — identitet koji ne postoji obara uvid; bez identiteta uvid i dalje radi
  po broju i tada legitimno vidi oba dokumenta.

**Strict ide do dna, ne do granice modula.** Prva verzija `valid` ugovora je
zaustavljala gutanje u `modStornoImpact`, a tri sekcije dolaze iz `modStornoFlow`
— i tamo je fail-open živeo još jednu rundu:

```
ScanPrijemnica pukne  ->  GetChainFlags proguta  ->  default hasDependents=False
                      ->  BuildStornoImpact nastavi  ->  valid = True
```

Takav uvid je opasniji od praznog: većina podataka je tačna, pa nema razloga za
sumnju. `strict` se sada provlači kroz **svih sedam putanja** — `GetChainFlags`,
`GetStornoChainRows`, `GetStornoBlockRows`, `ScanOtpremnica`, `ScanZbirna`,
`ScanPrijemnica`, `ActiveBlocksForFlow` i `GetBlokOtkupIDs`. Podrazumevano je
`False`, pa legacy `frmDokumenta`, koji zove isti model za svoj panel i **ne
čita** `valid`, ostaje netaknut; ekran Storno traži `strict = True`.

- **65** — kolona `OtkupID` se stvarno preimenuje, pa se traži `valid = False`.
  Test koristi generaciju **B**, ne A: test 51 stornira dokument generacije A
  zajedno sa njegovim blokom, pa bi spisak za A do ovog testa bio prazan — a
  prazan i nečitljiv spisak su baš ono što ovaj test razlikuje.

### Crveni toast preko uspešne ispravke

Operater je posle **uspešne** ispravke i dalje dobijao `X Radnja nije uspela:
modScrStorno.Scr_Event scrStA0` — preko uredno popunjene forme.

Uzrok nije bio pad radnje nego **život `Err`-a**: `On Error Resume Next`
prigušuje grešku, ali je **ne briše**. `OtvoriIspravku` ima baš takav gard, pa je
prigušena greška preživela povratak kroz `StornoPoModu` i `PokreniAkciju` sve do
`modUiScreens.ScrEvent` — a on posle `Application.Run` čita `Err.Number` i, ako
nije nula, javlja neuspeh.

Prethodna runda je dodala `Err.Clear` u `EH` handlere i to **nije bilo dovoljno**:
`EH` se na uspešnom putu uopšte ne izvršava. `Scr_Event` sada ima **jedan izlaz**,
i na njemu čisti `Err`; prava greška i dalje ide kroz `EH`, koji je prijavljuje,
pa čišćenje ne može da sakrije pad.

- **66** — meri baš to curenje: posle `Scr_Event`-a `Err.Number` mora biti nula,
  uz kontrolu u drugom smeru da je događaj stvarno obrađen. Sabotaža vraća stanje
  i test pada sa `dobijeno [-2147024809]`.

**Strict je stao na dispečeru, ne na dnu.** Prethodni commit je tvrdio „do dna",
a `ActiveBlocksForFlow` ga je prosleđivao samo za **otpremnicu**; zbirna i
prijemnica idu kroz `ActiveOtkupIDsByZbirna`, koji je ostao fail-open:

```
tblOtkup.BrojZbirne drift  ->  ActiveOtkupIDsByZbirna vrati prazno
                           ->  GetStornoBlockRows izadje na ids.count = 0
                           ->  dakle PRE svoje kapije
                           ->  blocks = 0, valid = True
```

Isti kvar kao u testu 65, samo **druga grana istog `Select Case`-a**. Zatvoreno u
`ActiveOtkupIDsByZbirna`, `CountActive` (koji je činio `Scan*` strict spolja a
slep iznutra) i `PkPoIdentitetu`. Uz to i legacy rupa u `HLI`: `docID = ""` je i
pod `strict` išao kroz `HL`, koji ima svoj `On Error Resume Next`.

- **67** — pokriva **obe** grane koje idu kroz `ActiveOtkupIDsByZbirna`. Pozitivna
  kontrola je nad **zbirnom**, ne prijemnicom: prijemnične blokove raniji testovi
  u istom prolazu storniraju, pa bi kontrola merila redosled testova umesto
  pravila.

### Nalaz van obima: log je slep na 112 mesta

Smoke je otkrio da pad upisa otpremnice **ne ostavlja nijednu `ERROR` liniju** u
`Log\OtkupApp_<datum>.log`. Uzrok nije u upisu nego u dijagnostici:

```vb
LogErr pise samo  If Err.Number <> 0
a EH blokovi rade:  errDesc = Err.Description
                    On Error Resume Next          ' <- resetuje Err
                    LogErr "SaveOtpremnicaMulti_TX"   ' <- vidi 0, ne pise nista
```

Isti oblik postoji na **112 mesta** kroz ceo `src-vba`, uključujući sve `Save*_TX`
writere. `Monitor_Error` prima opis eksplicitno, pa telemetrija ima podatak — ali
ona ide na udaljeni endpoint i samo ako je monitoring uključen; lokalno se gubi.

**Test 68** ne meri naš kod nego **semantiku VBA** na kojoj taj zaključak stoji,
da tvrdnja padne u suite-u a ne kroz prazan log posle incidenta.

Devet writera dokumenata (Otpremnica / Zbirna / Prijemnica, sve tri varijante)
**već je hvatalo** `errNum` i `errDesc` u lokale pre tog reda — podatak je bio tu,
samo se bacao. Sada se predaje izričito (`LogError "X", errDesc, errNum`). To je
najuža kriška koja operatera odblokira; preostalih ~103 mesta ostaje za zaseban
PR. `SaveKupciIzlaz_TX` i `SaveOMUlaz_TX` opis ne hvataju u lokale, pa ih nisam
dirao napamet.

**Šta je pad zapravo bio:** `Nedostaje kolona 'GeneracijaID' u tabeli
'tblOtpremnica'` iz `GeneracijaIDZaBroj` — schema drift na instalaciji, ne greška
u ovoj grani. Kolonu dodaje `EnsureSledljivostSchema` kroz `EnsureRuntimeSchema`,
a **i taj poziv je jedno od 112 mesta** sa mrtvim `LogErr`-om: ako je self-heal
šeme ikad pao, pao je nečujno. Posle ručnog `EnsureRuntimeSchema` upis prolazi.

### Upozorenje uz uspešan upis se više ne gubi

Dokument može biti snimljen, a da uz njega nešto ne prođe — prevezivanje paleta,
auto-zbirna, ili završetak ispravke koji stane na safe-stopu („više ispravki na
čekanju"). Toast to gubi **dvostruko**: piše u usko polje pa seče rep, a uspešan
toast se još i sam sakrije posle četiri sekunde.

Razlikovanje ide po oznaci koju katalog **već nosi**: `ChrW(10007)` = upozorenje
→ ide i u `MsgBox`; `ChrW(10003)` = informacija → ostaje u toastu. **Test 69**
čuva baš tu podelu, jer se iz koda ne vidi — oba su samo stringovi.

Sam popravak (112 mesta, zajednički writeri koje koristi i legacy) **nije u ovom
PR-u** — ide zasebno.

**Poslednja rupa u strict ugovoru: `PkPoIdentitetu`.** Dobio je parametar
`strict`, ali ga **nije koristio** — komentar iznad koda je tvrdio jedno, kod
radio drugo:

```vb
' ZADATA generacija koja se ne razresava je greska, ne poziv na fallback.
If ids.count = 0 Then Exit Function     ' ...a vracalo se prazno
```

Nizvodno je to izgledalo kao „dokument ne postoji" umesto „ne mogu da ga
razrešim" — a model se posle svega označavao kao `valid`. Uz to je **`ScanZbirna`
prekidao propagaciju** baš na PK resolveru, pa je zbirna prolazila i kad
otpremnica nije.

Sada: zadata generacija koje nema **diže** grešku u strict režimu, a `EH` je
re-raise. Van strict-a ostaje prazno, zbog zatečenih zapisa bez generacije.

- **70** — nestao identitet obara uvid, mereno **odvojeno** za otpremnicu i za
  zbirnu (dve grane, dve sabotaže), uz pozitivnu kontrolu i uz suprotan smer:
  bez identiteta oba i dalje rade po broju.

- **Devetnaest novih sabotaža**, svaka oborila svoju tvrdnju **po imenu**.

Zapisana je i **osma zamka** u `tools/sabotaza.py`: zamena ne sme biti podniz
sidra — `--vrati` je tada nalazi i u zdravom kodu, pa umesto vraćanja dodaje još
jedan primerak razlike (kod nas tri uzastopna `Err.Clear`).
- `COMPILE` → **`NEJASNO`** — ostaje ručna kapija.

Dve zamke iz `tools/sabotaza.py` naplaćene su ponovo, pa su obe sada zapisane u
zaglavlju tog fajla: sabotaža sa **praznom zamenom** se ne može vratiti
(`--vrati` traži zamenu u fajlu), a kod još nekomitovanog fajla ni
`git checkout` nije mreža; i oznaka `' SABOTAZA` **posle** `_` je syntax error,
pa run visi do timeout-a umesto da prijavi tvrdnju.

---

## v2.49.0 — `v6-ui-144` · log koji nije pisao, i lista koja nije imala izlaz

> Dva nalaza koja je ostavio smoke-test PR-a #202, oba van tadašnjeg obima. Nema
> nove poslovne logike — jedan vraća dijagnostiku, drugi otvara izlaz iz ekrana
> koji ga nije imao.

### 1) `LogErr` koji ne može ništa da zapiše — 91 mesto

Operateru je upis otpremnice pao, a `Log\OtkupApp_<datum>.log` nije imao **nijednu**
`ERROR` liniju. Uzrok nije u upisu nego u obliku handlera:

```vb
EH:
    errDesc = Err.Description
    On Error Resume Next            ' <- OVA naredba resetuje Err
    LogErr "SaveOtpremnicaMulti_TX" ' <- vidi Err.Number = 0, ne pise NISTA
```

`LogErr` piše samo kad je `Err.Number <> 0`, a **svaka** `On Error` naredba resetuje
`Err` objekat. Podatak je u tom trenutku već u lokalnoj promenljivoj — samo se
bacao.

PR #202 je popravio devet writera dokumenata (najuža kriška koja je operatera
odblokirala). Ostatak je ovde: **91 poziv u 24 fajla**, uključujući `modSEFPersistance`
(15), `modMasterSync` i `modBankaImport` (po 11), `modBankaMapiranje` (10),
`modSEFValidator` (9). Popravka je svuda ista i minimalna — `LogErr` ide **iznad**
`On Error`, dok je `Err` još živ. Ponašanje otpornosti se ne menja, samo prestaje
da bude nemo.

**Nova statička provera `MRTAV_LOG`** u `tools/vba_check.py` čuva to stanje. Gleda
**samo unutar handlera** (posle labele `EH:` / `ErrHandler:` / `Fin:` / `VRATI:`) —
van njega je `On Error Resume Next` **pre** poziva koji sme da pukne legitiman i čest
obrazac, pa bi lažan nalaz nad njim naučio da se checker ignoriše. Ta razlika je
kodifikovana u `MRTAV_LOG_CASES`: dva slučaja koja **moraju** da zapište i dva koja
**ne smeju**. Provera je odmah našla i tri mesta van `.bas` fajlova
(`ThisWorkbook.doccls`, `frmBankaImport.frm`, `frmStammdaten.frm`) koja ručna pretraga
nije obuhvatila.

`ThisWorkbook` je bio najgori slučaj: njegov handler je i `Monitor_Error` zvao sa
već resetovanim `Err`-om, pa je i **telemetrija** slala prazan opis. Sada se
`errNum`/`errDesc`/`errSrc` hvataju prvi, pre bilo čega drugog.

### 2) `EnsureSledljivostSchema` — self-heal šeme koji je ćutke odustajao

Pravi uzrok operaterovog pada bio je `Nedostaje kolona 'GeneracijaID' u tabeli
'tblOtpremnica'` — **schema drift na instalaciji**, ne greška u kodu. Kolonu dodaje
`EnsureSledljivostSchema` na svakom startu. Ta rutina je počinjala blanket
`On Error Resume Next`, pa je **prvi** pad ćutke preskakao ostatak posla: sveska bez
kolone, a upis koji na nju računa pada satima kasnije, sa porukom koja o uzroku ne
kaže ništa.

Sada svaka kolona ide kroz svoj gard (`EnsureKolonaSaTragom`): pad jedne se **zapiše**
i ne zaustavlja ostale. `EnsureColumnOnTable` uz to `LogWarn`-uje kad tabele nema —
tihi `Exit Sub` je bio tačno mesto na kom se `GeneracijaID` gubila.

Otpornost je nepromenjena; razlika je što se sada vidi.

### 3) Ekran Oporavak: „Odbaci ispravku"

Lista „Nedovršeno" je bila **čist pregled**. Operater vidi da ga safe-stop blokira
(„više ispravki na čekanju" — pravilo koje odbija da nagađa **koju** od više ispravki
novi dokument zamenjuje), a nema čime da to razreši. Jedini izlaz je bila legacy
`frmDokumenta`. `CancelCorrectionContext` je postojao sve vreme — falio mu je ulaz iz
novog UI-ja.

Radnja gađa **`CorrectionID`, ne poslovni broj.** Nad istim brojem može da stoji više
contexta (storno, pa opet storno istog dokumenta), pa bi izbor po broju zatvorio onaj
koji zatekne prvi — a operater je gledao drugi red. Red zato nosi identitet u
**nevidljivoj koloni** (prioritet 4, isti obrazac koji ekran Storno koristi za
`GeneracijaID`; petlja vidljivosti ide 3 → 1, pa 4 nikad ne prolazi).

Redovi koji **nisu** context (osirotele prijemnice, palete, izgubljeni blokovi) nemaju
šta da odbace — oni se rešavaju prevezivanjem, pa radnja nad njima **odbija** umesto
da tiho ne uradi ništa. Potvrda je `MsgBox` sa `vbDefaultButton2` i izričitim „dokumenti
se NE diraju"; dugme nosi `danger` stil, kao „Vrati storno".

### Šta je zatvorila recenzija (9.0/10)

**Test je merio transport, ne posledicu.** Test 71 dokazuje da `CorrectionID`
stigne do reda mreže — ali nijednom ne poziva samu mutaciju. Hard-kodovan
`CancelCorrectionContext("SV-TEST-1")`, ili `GridCell(red - 1, ...)`, prošli bi
71/71 netaknuti, dok PR tvrdi baš suprotno: „radnja gađa `CorrectionID`“.

Radnja je zato dobila jezgro bez UI-ja (`OdbaciIspravkuCore`) — sve osim potvrde i
toast-a; `MsgBox` u headless runu visi, pa se kroz UI ne može izmeriti. **Test 72**
meri posledicu, i najvažniju tvrdnju stavlja **prvu** (`AssertEq` prekida test na
prvom padu): odbaci `SV-TEST-2` → **`SV-TEST-1` ostaje `PENDING` / `NeedsRecovery=Da`**.
Tek onda da je izabrani zaista `CANCELLED` / `Ne`, pa da je nestao iz liste a sused
ostao.

Test **mutira** podatke, pa ih i **vraća** — fixture nosi tačno dve ispravke na
čekanju, a test 25 na tome meri safe-stop. Vraćanje se i **proverava**: neprovereno
čišćenje je isto što i nikakvo, jer bi sledeći test nasledio tiho izmenjen fixture i
pao po tuđem imenu.

**Indeks kolone više ne može da se raziđe.** `NED_COL_CID` je sada jedan broj koji
vezuje opis kolona, punjenje reda i radnju. Da je radnja imala svoj indeks, drift bi
bio nevidljiv: mreža bi izgledala ispravno, a radnja bi čitala tuđu kolonu.

**`MRTAV_LOG` je bio case-sensitive, a VBA nije.** `EH:` je hvatao, `eh:` / `Eh:` /
`errHandler:` nije — isti program za VBA, nevidljiv za checker. To je tačno ona
kategorija koju je ovaj PR i trebalo da zatvori: zelen checker koji cela jedna
legitimna sintaksa zaobilazi. Sada nosi `re.IGNORECASE` i svoj self-test slučaj
(labela malim slovima).

`SCROPO_BUILD` je podignut na `v6-ui-144` — `modScrOporavak` je dobio mutacionu
radnju, pa je stari pečat `v6-ui-135` lagao. `UISCR_BUILD` ostaje `v6-ui-143`:
`modUiScreens` u ovom PR-u nije menjan.

### Verifikacija

- `python tools\vba_check.py` → **čisto (191 fajlova)**, exit 0.
- `python tools\vba_check.py --self-test` → **čisto (34 slučaja)**.
- `python tools\run_vba.py --suite RunAllTests` → **ZELENO**, 72 testa.
- `python tools\run_vba.py` (pun set) → **ZELENO**, svi suite-ovi.
- **Test 71** (`T_Oporavak_OdbaciIspravku_PoIdentitetu`) tvrdi da svaki context red
  nosi **svoj** `CorrectionID` u koloni `NED_COL_CID` — ne samo „nije prazno", nego baš oba
  ID-ja iz fixture-a, jer bi test koji meri praznoću prošao i kad bi svi redovi
  nosili isti CID. **Test 26** je dopunjen: „Nedovršeno je samo pregled" više nije
  tačno.
- **Pet novih sabotaža**, svaka oborila svoju tvrdnju **po imenu**:
  `oporavak-nema-odbaci`, `oporavak-cid-ne-stize-u-red`, `oporavak-cid-kolona-vidljiva`,
  `oporavak-odbacuje-prvi-a-ne-izabrani` (pada baš na „SV-TEST-1 ostaje netaknut“),
  `oporavak-cid-kolona-drift`.
- `COMPILE` → **`NEJASNO`** — ostaje ručna kapija (`Alt+F11 → Debug → Compile VBAProject`).

### Provera `ODSECEN` — prazan fajl više nije „čist“

`vba_check` je **prazan fajl prijavljivao kao čist**: prazan fajl nema šta da prekrši,
pa nijedna provera nije imala reč. Zelen izlaz nad izbrisanim modulom je gori od
crvenog — `ImportAllVBA` ga uveze kao prazan i sve što je u njemu bilo nestane, bez
ijedne poruke.

Ulaz je uvek isti obrazac:

```python
io.open(P, "wb").write(s.encode("ascii"))   # open() odsece PRE nego sto encode pukne
```

U ovoj sesiji je tri puta ostavio `.bas` od nula bajtova. Bezbedno je
`data = s.encode(...)` **pa** upis.

Nova provera traži `Attribute VB_Name = "..."` — red koji nosi **svaki** izvoz iz
VBE-a, svih 191 fajlova bez izuzetka. Ništa strožije ne prolazi nad zatečenim
izvorom: najmanji legalan modul ima 154 bajta i **nema** `Option Explicit`
(`modMeteo.bas`), a `.frm` nosi `VB_Name` tek posle `Begin` bloka — oba su među
self-test slučajevima koji **ne smeju** da zapište.

Self-test je sa 34 na **39 slučajeva**. Dokaz u oba smera: sa ugašenom proverom pada
tačno tri slučaja, svaki po imenu (`prazan fajl`, `samo beline`,
`kod bez VB_Name zaglavlja`), dok dva negativna ostaju na nuli.

---

## v2.50.0 — `v6-ui-145` · toast koji se nikad nije video

> Nađeno smoke-testom PR-a #203, a starije je od njega: pogađa **svaki** ekran
> novog UI-ja osim unosa dokumenata.

Operater je prijavio da je dugme „Odbaci ispravku" mrtvo — aktivno, klik ne radi
ništa, log prazan. Instrumentacija je pokazala suprotno: radnja je radila, i to
tačno. Nevidljiv je bio **odgovor**.

`ShowZones` sakriva zonu unosnog ekrana na svakom ugovornom ekranu:

```vb
' Ovo je samo ekran dokumenata: KPI traka, kontekstni red, forma, kartice.
nmv = Array("zKpi", "zCtx", "zForm", "zRight")
    frm.Controls(CStr(nmv(i))).Visible = dok      ' dok = (mScreen = "DOKUMENTI")
```

a `ShowToast` je pisao baš tamo:

```vb
Set fr = mFrm.Controls("zForm").Controls("tstOk")
fr.Visible = True          ' kontrola u SKRIVENOM roditelju -- ne prikazuje nista
```

`Visible = True` nad kontrolom u skrivenom roditelju ne prikazuje ništa, a
`On Error Resume Next` na vrhu `ShowToast`-a guta i eventualnu grešku. Rezultat:
na ekranima **Storno, Palete, Oporavak i Agrohemija nijedna poruka nikad nije
stigla do operatera** — ni potvrde, ni odbijanja, ni `ScrLastErr`, kanal kojim
ljuska prijavljuje da je ekran pukao.

U logu se to lepo vidi: tri klika na „Prevezi" u četiri sekunde. Radnja je svaki
put odgovorila, samo nemo, pa je operater kliktao dalje.

**Popravka:** toast se seli u **mrezu** (`zGrid`) — traka preko cele širine, tačno
iznad podnožja (`Prikazano / Ukupno / strane`).

Prva verzija ga je stavila u naslovnu traku: video se, ali je **pokrivao naslov
ekrana**. Mreža je bolji domaćin — vidljiva je na svim ekranima isto kao i naslov,
ali joj je dno prazno kad lista ne popuni stranu, pa poruka ništa ne zaklaňa. A i
odgovor na radnju nad redom stoji uz same redove.

Puna širina je namerna: razlog odbijanja je često duga rečenica, a poruka koja se
seče ne vredi više od one koje nema.

`grdBody` se pravi **posle** trake, pa `ShowToast` radi `ZOrder 0` — bez toga bi
poruka stajala ispod redova i opet se ne bi videla, druga varijanta istog kvara.

Time otpada i sprega sa `KgLineVisible`: toast je ranije delio prostor sa zbirom
kilograma u akcionom redu, pa ih je sakrivao dok stoji. Sada su kilogrami stalno
vidljivi, a `KgLineVisible` i stari `tstOk` su uklonjeni kao mrtvi.

**Verifikacija:** `vba_check` čisto (191), `who_writes` ažuran,
`RunAllTests` **ZELENO (72)**, pun set **ZELENO** (11 suite-ova). `COMPILE` →
`NEJASNO`, ostaje ručna kapija.

Automatski test ovde **ne postoji i ne može da postoji**: tvrdnja je „kontrola je
vidljiva operateru", a forma se u harnessu gradi bez `.Show`. Ostaje smoke: poruka
mora da se vidi na Oporavku, Storno i Paletama, i dalje da radi na unosu.

---

## v2.50.1 — `v6-ui-146` · uvid o stornu: 1969 → merenje pa popravka

Operater je prijavio zastoj od 2–3 sekunde po kliku na red na ekranu Storno.
Ranije je ista prijava zatvorena rečenicom „nisam merio, ne znam gde odlazi" —
ovog puta je prvo dodato **merenje po sekciji**, pa tek onda popravka.

### Merenje

`BuildStornoImpact` meri svih sedam sekcija i piše **jednu** liniju, i to samo kad
ukupno pređe **400 ms** — granicu na kojoj čekanje postaje primetno. Brz put ostaje
tih, spor sam sebe prijavi. Nije privremeni debug nego trajna kapija.

Sa terena, dva različita tipa dokumenta:

```
Uvid Otpremnica 5/210726-2 trajao 1969 ms: zaglavlje 12, lanac 51, blokovi 8,
zastavice 20, palete 1879, faktura 0, zbir 0 ms.

Uvid Zbirna S1/220726 trajao 1977 ms: zaglavlje 12, lanac 27, blokovi 27,
zastavice 16, palete 1895, faktura 0, zbir 0 ms.
```

**95% vremena u jednoj sekciji, i to istoj za oba tipa.** Jedna popravka, ne dve —
što se bez merenja ne bi znalo.

### Uzrok

`GetPaleteImpactByField` je za **svaku** paletu u rezultatu radio:

| Poziv | Šta radi |
|---|---|
| `FindRowIndexByID(TBL_PALETA, ...)` | linearni prolaz kroz `tblPaleta` |
| `PaletaLabel(pid)` | **još jedan** isti prolaz + `GetTableData` |
| `IsPaletaPreradjena(pid)` | **treći** isti prolaz + `GetTableData` |
| `GetPaletaAggregates(palRow, ...)` | `GetTableData` + 5 × `GetColumnIndex` |

Dakle **tri linearna prolaza i tri kopije cele tabele po paleti**. Batch keš
(`BeginTableCache`) sprečava ponovno **čitanje** iz Excela, ali ne i kopiranje:
`d = GetTableData(...)` dodeljuje niz `Variant`-u, a VBA tada kopira ceo niz.

### Popravka

Tabela se čita **jednom**, a red se nalazi kroz rečnik `PaletaID → red`. Prvi red
pobeđuje, isto kao `FindRowIndexByID` — da se ponašanje nad duplim ID-jem ne
promeni usput. Pomoćne rutine ostaju za svoje ostale pozivaoce; ovde se čitaju
ista polja, iz istog reda, samo bez ponovnog traženja tog reda.

### Test 73 — jer zatečena suita ovo ne bi videla

Postojeći testovi tvrde samo **koliko** paleta uvid nosi i **koliki im je zbir**, a
zbir dolazi iz druge petlje koju izmena ne dira. Polja iz zaglavlja palete —
oznaka, popunjenost, kapacitet, neto, prerađenost — nije merilo **ništa**, a baš
njih izmena preračunava. Zeleno posle refaktora zato ne bi značilo ništa.

`T_ImpactPalete_ZaglavljeIzPraveVrste` tvrdi da zaglavlje dolazi iz reda **baš te**
palete (`PAL-TEST-Z2` = 12/2026, 20 gajbi od 100, 200 kg), i razdvaja ta dva
računa: koliko je na paleti **ukupno** i koliko od toga nosi **ovaj** dokument.

Sabotaža `palete-zaglavlje-prvi-red` uzima prvi red tabele umesto traženog i pada
po imenu: `popunjenost je iz reda BAS te palete — očekivano [20], dobijeno [25]`.

### Verifikacija

`vba_check` čisto (191), self-test čisto (39), `who_writes` ažuran,
`RunAllTests` **ZELENO (73)**, pun set **ZELENO** (11 suite-ova).
`COMPILE` → `NEJASNO`, ostaje ručna kapija.

Ubrzanje se **ne prijavljuje kao izmereno** — merenje sa terena postoji samo za
stanje pre popravke. Novi broj daje isti `WARN`, ili njegov izostanak.

---

## v2.50.2 — `v6-ui-147` · `ByVal` na nizu je kopirao celu tabelu, po ćeliji

Nastavak prethodnog: **prva popravka nije pomogla**, i to je bio podatak. Vreme je
ostalo isto (1918 / 1895 / 1906 ms) bez obzira na tip dokumenta i na broj paleta —
a trošak po paleti bi morao da varira. Konstanta znači fiksni trošak.

Merenje unutar sekcije je onda pokazalo tačno mesto:

```
[v6-ui-146] 1063 stavki, 1 paleta, ukupno 1918 ms:
citanje tabele 0, prolaz kroz stavke 1883, obrada paleta 35 ms.
```

Čitanje tabele **0 ms** (batch keš radi), obrada palete **35 ms** — a prolaz kroz
1063 reda **1883 ms**. To je **1,8 ms po redu** za čitanje dva polja iz niza koji
je već u memoriji. Toliko ne traje pristup nizu; toliko traje kopiranje.

### Uzrok

```vb
Public Function SafeCell(ByVal d As Variant, ByVal r As Long, ByVal idx As Long)
```

`ByVal` na `Variant`-u koji **sadrži niz** znači da VBA kopira ceo niz pri svakom
pozivu. `SafeCell` je čitač **po ćeliji** — u toj petlji se zove dvaput po redu.
Dakle 2126 poziva × kopija tabele od 1063 reda.

Popravka je jedna reč: **`ByRef`**. Funkcija `d` samo čita, nikad ne piše, pa je
razlika isključivo u tome što se niz ne umnožava. `SafeCell` ima **206 pozivalaca**
u `modPaletniList` — ubrzanje ne pripada samo uvidu o stornu.

Isti obrazac je popravljen i u `modStornoFlow.NzTxC`, koji je u istoj putanji
(lanac i blokovi). Tamo se nije video jer su te tabele manje — ali greška je ista.

### Šta je ovo ostavilo za kasnije

Isti potpis — čitač po ćeliji koji prima niz `ByVal` — postoji i u
`modDokumenta.StornoCellRaw` / `StornoCellText` i `modKarticaDetalji.CellVal`.
Nisu mereni i nisu u ovoj putanji, pa ih ne diram napamet; zapisani su kao nalaz.

### Zašto ovde nema testa

Tvrdnja je „isto ponašanje, manje kopiranja". Vrednosti se ne menjaju, pa test koji
bi ih proveravao ništa ne bi dokazao, a test koji meri **vreme** u harnessu je
neupotrebljiv — fixture ima desetak stavki, tamo je i stara verzija brza.

Kapija je **merenje ugrađeno u kod**: obe rutine se prijave same kad pređu 400 ms.
Regresija ovog tipa ubuduće stiže kao `WARN` linija, ne kao pritužba operatera.

### Verifikacija

`vba_check` čisto (191), self-test čisto (39), `who_writes` ažuran,
`RunAllTests` **ZELENO (73)**, pun set **ZELENO** (11 suite-ova).
`COMPILE` → `NEJASNO`, ostaje ručna kapija.

Ubrzanje se **ne prijavljuje kao izmereno** dok ne stigne merenje sa terena.

### Napomena o efektu se više ne seče

Tabela „Efekat storna po modu" ima tri kolone, a komentar iznad rasporeda je već
opisivao pravo ponašanje: dokument fiksno, info fiksno, **napomena uzima ostatak**.
Kod je ipak sve tri delio na trećine, pa je napomena dobijala trećinu i sekla se
na pola rečenice:

```
DUPLIKAT i PONISTENJE: preracun, storno ako ostane prazn
DUPLIKAT: ostaje osirocena (rucno)  |  PONISTENJE: stornir
```

Prve dve kolone nose kratak i predvidiv sadržaj — naziv dokumenta i broj ili
brojač u zagradi — pa im fiksna širina dostaje. Trećoj je dužina neograničena i
ona je jedina koja opisuje **posledicu**; zbog nje ekran i postoji.

Automatskog testa nema: raspored se meri u pikselima nad formom koju harness
gradi bez `.Show`. Ostaje smoke — nijedna napomena se ne sme završiti sečenjem.

---

## v2.51.0 — `v6-ui-148` · poslovni jezik sekcije „Posledice po osnovu storna"

Operater: *„previše je laički napisano, nije uopšte poslovno."* Tačno — i uzrok
nije bio u izboru reči nego u tome **gde su reči stajale**.

Svi ti tekstovi bili su **tvrdo ukucani ASCII literali u `modStornoFlow`**. VBA
izvor mora ostati ASCII, pa se poslovna rečenica u njemu ne može ni napisati:
bez č/ć/š/ž/đ ostaje telegrafski zapis („preracun, NE pada", „ako ih cekiras").
Selidba u katalog (`modPoruke` + `ChrW`) je zato preduslov, ne kozmetika.

### Osnovi storna — imenovani po poslovnom događaju

| Pre | Sada |
|---|---|
| Pogrešan unos | **Ispravka dokumenta** |
| Duplikat | **Dupli unos** |
| Ništa se nije desilo | **Poništenje prometa** |
| Reši kasnije | **Odloženo rešavanje** |
| Storniraj | **Storno bez zamene** |
| Ispravka *(revers)* | **Zamena reversa** |

Objašnjenja ispod dugmadi su **namerno kratka** — labela ima 168pt, a duga
rečenica bi se sekla. Puna formulacija posledice stoji u tabeli iznad, koja širinu
ima.

### Posledice po dokumentu

| Pre | Sada |
|---|---|
| `stornira se (uz ambalazu)` | stornira se, sa pripadajućom ambalažom |
| `preracun, storno ako ostane prazna (jedini vlasnik)` | preračunava se; stornira se ako ostane bez otpremnica |
| `preracun, NE pada (deljena - sestre ostaju)` | preračunava se; ostaje aktivna jer nosi i druge otpremnice |
| `preracun, storno ako padne na 0` | preračunava se; stornira se ako ostane bez količine |
| `ostaje osirocena (rucno)` | ostaje osirotela; prevezuje se ručno |
| `odvezuju se (prezivljavaju)` | odvezuju se od zbirne i ostaju aktivne |
| `ostaje netaknuta` | ostaje nepromenjena |
| `skidaju se` | odvezuju se sa palete |
| `oslobadja se (stavke osirocene)` | oslobađa se; stavke ostaju osirotele |
| `Samostalni - storniraju se samo ako ih cekiras (svaki mod)` | evidentiraju se zasebno; storniraju se samo po izboru iz liste |
| `Stornira se (saldo se koriguje, bez kontra-stavke)` | stornira se; saldo se koriguje bez kontra-stavke |

**Prefiksi se sada zovu isto kao dugmad**: `DUPLIKAT:` / `PONISTENJE:` →
`DUPLI UNOS:` / `PONIŠTENJE:`. Do sada se nisu poklapali — u tabeli je pisalo
`PONISTENJE`, a dugme je glasilo „Ništa se nije desilo", pa je operater morao sam
da poveže redak sa dugmetom na koje se odnosi.

Naslov sekcije: `EFEKAT STORNA PO MODU` → **`POSLEDICE PO OSNOVU STORNA`**.
„Mod" je programerski; četiri izbora su poslovni **osnovi**.

**Zatečeni domenski termini ostaju** — „osirotela", „prevezivanje", „zbirna",
„otkupni blok" su jezik ove aplikacije i menjati ih bilo bi štetnije od dobitka.

### Test 74 — selidba u katalog uvodi nov tihi kvar

Ključ koji katalog ne zna vraća **prazan string**: najvažnija kolona ekrana ostane
prazna, bez greške i bez traga. `vba_check` (provera `PORUKA`) hvata ključ bez para
u `UpsertPoruke` — ali ne i katalog koji nije osvežen, a to je baš ono što se
dešava posle importa.

`T_StornoEfekat_TekstIzKataloga` zato traži da katalog stvarno nosi tekst (ne
prazan ključ), pa da se napomena sklopi iz njega, i to u **oba** oblika: jedan
spojen prefiks kad je efekat isti za oba osnova, dva kad se razlikuju.

Sabotaža `efekat-uvek-spojen-prefiks` uklanja poređenje u `ChainEff` i pada po
imenu: *„razlicit efekat nosi OBA prefiksa u istom redu — očekivano [True],
dobijeno [False]"*.

Slučaj „ključ ne postoji u katalogu" **nema** sabotažu, i to namerno: hvata ga
`vba_check` još pre nego što suite krene, pa bi sabotaža pala na tuđoj kapiji i
lažno tvrdila da je meri test.

### Verifikacija

`vba_check` čisto (191) · self-test čisto (39) · `who_writes` ažuran ·
`RunAllTests` **ZELENO (74)** · pun set **ZELENO** (11 suite-ova) ·
`COMPILE` → `NEJASNO`, ostaje ručna kapija.

Isti tekst vidi i **legacy panel** `frmDokumenta` — model je zajednički
(`BuildStornoImpact`), pa se promena ne razilazi između dva ekrana.

---

## v2.52.0 — `v6-ui-149` · otkupni blokovi se ponovo biraju, a ne storniraju svi

Operater: *„fali check lista sa otkupnim listovima pri stornu uzvodnih dokumenata.
to postoji kod legacy, mora i ovde."* Tačno — i poređenje sa legacy-jem otkrilo je
da nije reč samo o prikazu.

| | Legacy panel | Nov ekran do sada |
|---|---|---|
| Podrazumevano | **nijedan blok nije čekiran** | — |
| Na potvrdu | stornira **samo čekirane** | stornira **SVE** |

Legacy: *„default oslobodjeni/netaknuti, cekiran = dodatno storniran."* Nov ekran je
ispisivao spisak u `MsgBox`-u i na „Da" stornirao sve. Bio je dakle **destruktivniji
od legacy-ja**, i to ne namerno nego zato što multiselect nije bio prenet — što je
katalog migracije i beležio kao svesnu privremenu odluku („sve-ili-ništa").

### Lista „Blokovi"

Jedanaesti čip ekrana Storno. Nije tip dokumenta — kao ni „Lanac robe" — nego
**pogled nad već izabranim dokumentom**:

| Kolona | |
|---|---|
| ✓ | izbor; klik na red uključuje/isključuje |
| Broj otkupa, Količina, Klasa, Kooperant | isto što je legacy prikazivao |
| *(nevidljiva)* | `OtkupID` |

Izbor se drži **po `OtkupID`-u, ne po broju otkupa**: broj se računa po kooperantu,
pa dva bloka lako dele isti — a spisak završava u `StornoSelectedBlocks_TX`, dakle
u mutaciji. Isti razlog zbog kog `GeneracijaID` postoji na ekranu Storno i
`CorrectionID` na Oporavku.

Redovi dolaze iz **već izgrađenog uvida** (`mImpact`), ne iz novog skeniranja: model
na osnovu koga zona tvrdi posledice mora biti i izvor spiska nad kojim se bira —
inače ekran pokazuje jedno, a stornira drugo.

Prelazak na taj čip **ne poništava izbor dokumenta**, za razliku od prelaska između
tipova. Poništio bi baš ono što lista treba da prikaže.

### Dva nalaza iz istog posla

**Test seam je lagao.** `Scr_IzborTestSet` je postavljao polja izbora, ali **nije
gradio uvid** — što produkcija radi pri svakom kliku na red. Test 62 je na tome
gradio ceo svoj slučaj: „framework dokument bez uvida ne nudi radnju" prolazilo je
zato što uvida u testu nikad nije ni bilo, a ne zato što kapija radi.

Seam sada radi ono što radi i produkcija, a test 62 neuspeh uvida pravi **stvarno** —
zadatom generacijom koju nijedan red ne nosi — i dobio je pozitivnu kontrolu: isti
dokument sa razrešivim identitetom radnje **ima**, pa se vidi da kapija nije
zaglavljena na nuli.

**`TabelaTipa` je fail-open.** Nepoznat ključ tiho vraća `tblOtkup` (`Case Else`), pa
bi lista sa pogrešnim ključem prikazala otkupne listove pod svojim naslovom — tačno
ono na šta komentar u testu 56 upozorava. Nije popravljeno u ovom PR-u (dira sve
liste, ne samo ovu); test 56 zato meri **posledicu** — koje kolone lista vrati.

### Verifikacija

`vba_check` čisto (191) · self-test (39) · `who_writes` ažuran ·
`RunAllTests` **ZELENO (75)** · pun set **ZELENO** · `COMPILE` `NEJASNO`.

**Test 75** tvrdi da je podrazumevano stanje prazno i da označavanje pogađa baš taj
blok, po `OtkupID`-u. Dve sabotaže, obe oborile svoju tvrdnju po imenu:
`blokovi-svi-oznaceni` i `blokovi-oznake-prezive-izbor`.

Druga je usput ponovo naplatila **zamku 8** iz `tools/sabotaza.py`: prva verzija joj
je uklanjala red i ostavljala `mSelTip = ""`, koji postoji i u zdravom kodu — pa ga
je `--vrati` našao tamo i dodao još jedan `Set mBlokOznaceni = Nothing`. Zamena sada
nosi oznaku `' SABOTAZA`, pa je jedinstvena.

---

## v2.52.1 — `v6-ui-150` · vodeće „?" u dijalogu i uput na ekran koji ne postoji

Smoke posle upisa otpremnice pokazao je poruku:

```
? Vise ispravki na cekanju za ovaj tip -- prevezivanje NIJE uradeno.
  Resi kroz Osiroceni dokumenti.
```

Tri greške u jednoj rečenici, i sve tri različite vrste.

### 1. Oznaka je signal, ne tekst

`ChrW(10007)` (✗) uz poruku znači **„ovo idi u dijalog"** — razlikovanje uvedeno u
`v6-ui-143`, jer se upozorenje uz uspešan upis gubilo u traci. Ali `MsgBox` crta
kroz **ANSI kodnu stranu**, u kojoj tog znaka nema, pa ga je operater video kao
vodeće `?`.

Oznaka se sada skida pred dijalogom (`PorukaZaDijalog`). U traci poruka, koja je
Unicode, **ostaje** — tamo nosi značenje (crveno = nešto nije prošlo), dok u
dijalogu istu ulogu već igra `vbExclamation`.

Isti tekst u istoj traci se video ispravno — zato je greška i preživela: proverom
jednog kanala izgleda tačno.

### 2. Pogrešno slovo

`Osiro` & `ChrW(269)` & `eni` daje **„Osiročeni"**. Ostale četiri poruke istog
sadržaja koriste `ChrW(263)` (ć). Jedan ključ je odstupao.

### 3. Uput na ekran koji u novom UI-ju ne postoji

„Reši kroz **Osiroćeni dokumenti**" je naziv **legacy panela**. U novom UI-ju taj
ekran se zove **Oporavak**, a lista **Nedovršeno**. Operater je upućivan na nešto
što na ekranu ne piše nigde.

Ispravljene su sve četiri poruke koje su vodile na legacy imena:

| Pre | Sada |
|---|---|
| Reši kroz Osiroćeni dokumenti. | Reši na ekranu Oporavak → Nedovršeno. |
| Reši kroz: Osiroćeni dokumenti. | Reši na ekranu Oporavak → Osiroćene prijem. |
| uradi ručno (Osiroćeni dokumenti → Palete) | uradi ručno (Oporavak → Osiroć.palete) |
| OTKAZI = REŠI KASNIJE … (Osiroćeni dokumenti) | OTKAZI = ODLOŽENO REŠAVANJE … (Oporavak) |

Poslednja je usklađena i sa novim imenima osnova storna iz `v6-ui-148`.

### Verifikacija

`vba_check` čisto (191) · self-test (39) · `who_writes` ažuran ·
`RunAllTests` **ZELENO (75)** · pun set **ZELENO** · `COMPILE` `NEJASNO`.

**Test 69** je dobio treću tvrdnju: poruka u dijalogu ide **bez** oznake, nije
prazna posle skidanja, i počinje slovom a ne razmakom. Skidanje je zato izdvojeno
u `PorukaZaDijalog` — da se može izmeriti; dijalog u headless runu visi, pa se sam
`MsgBox` ne može testirati.

Sabotaža `dijalog-nosi-oznaku` vraća oznaku u dijalog i pada po imenu:
*„DOKUNOS_MSG_VISE_ISPRAVKI u dijalogu ide BEZ oznake — očekivano [False],
dobijeno [True]"*.

### Jedanaesti čip se crtao, ali klik na njega nije radio ništa

Operater: *„čip Blokovi postoji, ali koja mu je svrha? mrtav je…"* — i bio je u pravu.

Dispečer klika je glasio:

```vb
If Left$(tag, 5) = "lsSeg" And Len(tag) = 6 Then
```

`Len(tag) = 6` pokriva **samo `lsSeg0`…`lsSeg9`**. Jedanaesti čip je `lsSeg10` —
sedam znakova — pa je propadao kroz granu i **nije radio ništa**: crta se, boji se
na hover, a klik nema kome da stigne.

Ljuska ima **dve kapije** nad čipovima, i ovo je bila druga:

| Kapija | Šta odlučuje | Kad je pukla |
|---|---|---|
| `MAX_SEG` | da li se čip **crta** | `v6-ui-143` — „Izvodi" su nestali |
| dispečer klika | da li klik **stigne** | sada — „Blokovi" su mrtvi |

Obe daju isti simptom za operatera i obe ćute. Redni broj se sada čita kroz
`SegIndeksIzTaga`, ne merenjem dužine taga.

**Test 56** meri obe odvojeno: crtanje jednom tvrdnjom, razrešavanje klika drugom —
i to za **poslednji** čip, jer je prvi radio i pre.

Sabotaža postoji **samo za prvu** kapiju; njeno sidro je usput popravljeno (gađalo
je `MAX_SEG = 10`, a sada je 11). Za drugu je namerno nema: test može da tvrdi da
`SegIndeksIzTaga` razrešava poslednji čip, ali ne i da ga dispečer zaista **zove** —
klik kroz formu se u harnessu ne može odigrati. Sabotaža nad dispečerom bi ostavila
suite zelen i lažno tvrdila da je tvrdnja pokrivena (zamka 5). Ta kapija ostaje na
smoke-u.

### Build UI-ja se sada vidi na ekranu

Sporedni nalaz iz iste runde, ali je koštao dva puna kruga: svaka nereleasovana
sveska nosi isti `v0.0.0-dev`, pa se sa ekrana **nije moglo videti da li je posle
`ImportAllVBA` u njoj nov ili star UI kod**.

Dvaput se zbog toga merilo nad neuvezenim buildom i dvaput se nije moglo
razlikovati *„nije pomoglo“* od *„nije uvezeno“* — jednom kod ubrzanja paleta,
jednom kod mrtvog čipa.

Na **nereleasovanoj** svesci se u sidebaru sada prikazuje `OTKUI_BUILD` umesto
verzije sveske. Prvi pokušaj je pisao oba, ali se odsekao: raspored drži tu labelu
na 55pt uz desnu ivicu sidebara, pa je od `v0.0.0-dev  v6-ui-153` ostajalo
`v0.0.0-d`. Na releasovanoj svesci ostaje verzija sveske — tamo ona jeste podatak.

Merenja u logu istu oznaku nose od `v6-ui-147`; ovo zatvara i drugi kanal, bez
Immediate prozora.

### Red o blokovima nosi STANJE, ne pravilo

Operater, pošto je čip prorad io: *„nije ni ovo rešenje loše, ali kako operater da
vidi da u čip Blokovi treba da odluči šta da radi sa blokovima?“*

Legacy je listu blokova imao **unutar** panela odluke, pa se videla sama od sebe.
U novom UI-ju je iza čipa — što je dobro za prostor, ali ništa nije govorilo da tamo
ima šta da se reši.

Red „Otkupni blokovi" u tabeli posledica je **jedini koji traži odluku**, a odluka
se donosi na drugom mestu. Zato taj red više ne opisuje pravilo nego **stanje**, i
menja se pri svakom štikliranju:

| Stanje | Tekst |
|---|---|
| nema blokova | nema samostalnih otkupnih blokova |
| ima ih, nijedan izabran | **nijedan nije izabran — biraju se u listi „Blokovi"** |
| izabrani | izabrano 2 od 3 — storniraju se uz Dupli unos ili Poništenje prometa |

Red se prepoznaje po tome što mu je napomena baš taj ključ iz kataloga, a ne po
nazivu dokumenta: naziv je tekst za prikaz i sme da se menja.

**Test 75** tvrdi sva tri stanja; `BlokStatusTekst` je zato javna — zona se crta nad
formom koju harness gradi bez `.Show`, pa se sam natpis ne može pročitati.
Sabotaža `blok-status-ne-prati-izbor` pada po imenu.

---

## v2.53.0 — `v6-ui-155` · Oporavak nosi brojku, jer nije radnja nego zaostatak

Operater: *„ima li poslovno i UX smisla da Oporavak i Storno stoje ravnopravno jedan
ispod drugoga?"*

Nemaju, i to se videlo u ponašanju aplikacije pre nego u teoriji:

- **Storno je radnja**, Oporavak je **posledica** — spisak onoga što je ostalo
  nedovršeno, uglavnom zato što je neki storno stao na safe-stopu.
- **Oporavak nema svoj ulaz.** Niko ne odluči „idem da radim oporavak" — tamo ga
  pošalje poruka. U `v6-ui-150` su četiri poruke prepravljene da glase „Reši na
  ekranu Oporavak → Nedovršeno".
- **Sve u njemu je broj koji bi trebalo da bude nula.** Na instalaciji se nakupilo
  **44 / 32 / 2**, a sidebar je izgledao isto i kad je iza stavke nula.

Stavka ostaje u OPERACIJE — posao u njoj **jeste** operativan (prevezivanje i
vraćanje storna su mutacije, ne izveštaji) — ali sada nosi **brojku**:

```
Storno
Oporavak            44
```

Nula se ne crta. Značka postoji da bi se video **zaostatak**; nula uz svaku stavku
bi je pretvorila u ukras koji se prestane primećivati.

### Ljuska i dalje ne poznaje nijedan ekran po imenu

Najlakše rešenje bilo bi da sidebar pozove `GetNedovrseno`. To je tačno ono što ceo
ugovor ekrana izbegava — sledeći ekran sa zaostatkom morao bi da se doda **u
ljusku**, a ne u svoj modul.

Zato je `Scr_Brojac` **opcion član ugovora**: ekran koji ima šta da broji ga
implementira, ostali ga nemaju i dobijaju nulu. Ljuska pita sve redom i ne zna ko je
odgovorio ni šta se broji.

`Application.Run` na nepostojeću proceduru **diže grešku**, pa je bez gutanja te
greške sidebar ne bi ni iscrtao — ali greška se i **čisti**, da ne procuri u ljusku
(ista klasa nalaza kao test 66).

### Cena je držana pod kontrolom

`Scr_Brojac` se zove iz `RefreshFromData` **i iz `EnsureGridLoaded`** — dakle na
promenu podataka i jednom na otvaranju aplikacije, ne pri
svakom crtanju sidebara. `GetNedovrseno` je više prolaza kroz tabele; na crtačkoj
učestalosti bio bi to isti kvar zbog koga su dva PR-a vadila sekunde iz uvida o
stornu.

Broji se **ista lista** koju ekran prikazuje kao „Nedovršeno", pa se broj u meniju i
broj na ekranu ne mogu raziću.
## v2.54.0 — `v6-ui-157` · svi čitači ćelije primaju niz `ByRef`

Zatvara nalaz iz `v6-ui-147`: `ByVal` na `Variant`-u koji **sadrži niz** tera VBA da
kopira ceo niz pri svakom pozivu. Kad je funkcija čitač **po ćeliji**, to je kopija
cele tabele po pročitanom polju.

Mereno na terenu, na `modPaletniList.SafeCell`:

```
1063 stavki, 1 paleta, 1918 ms: citanje tabele 0, prolaz kroz stavke 1883,
                                obrada paleta 35 ms.
```

**1,8 ms po redu** za čitanje dva polja iz niza koji je već u memoriji. Popravka je
jedna reč.

### Tri zapisana, dvadeset šest popravljenih

Zapisana su bila tri (`StornoCellRaw`, `StornoCellText`, `CellVal`). Sistematičan
pretres po istom potpisu — funkcija koja prima niz `ByVal` i indeksira ga sa dva
indeksa — dao je **jedanaest**, a nova statička provera još petnaest:

| Modul | Čitači |
|---|---|
| `modDokumenta` | `StornoCellRaw` · `StornoCellText` |
| `modKarticaDetalji` | `CellVal` |
| `modPrint` | `OtpC` · `OtpN` |
| `modStornoContext` | `TxCell` |
| `modStornoDok` | `KolicinaReda` · `CeliBroj` · `CeliBrojD` · `CelijaAko` |
| `modStornoZurnal` | `TxZ` |

`OtpN` je bio dupla kazna: prosleđivao je isti niz dalje u `OtpC`, dakle druga kopija
po pozivu.

Svaki je **pojedinačno pročitan pre izmene** — svi samo indeksiraju niz, nijedan u
njega ne piše, pa je `ByRef` bez ijedne posledice po ponašanje. Posle izmene u
`src-vba` nema više nijednog čitača ćelije sa `ByVal` nizom.

### Zašto ovde nema testa

Vrednosti se **ne menjaju** — test koji bi ih proveravao ništa ne bi dokazao. Test
koji meri **vreme** u harnessu je neupotrebljiv: fixture ima desetak redova, tamo je
i stara verzija brza.

Kapija ostaje merenje **ugrađeno u kod**: `BuildStornoImpact` i
`GetPaleteImpactByField` se same prijave u log kad pređu 400 ms. Regresija ovog tipa
stiže kao `WARN` linija, ne kao pritužba operatera.

### Provera `KOPIJA_NIZA` — da dvanaesti ne uđe neprimećen

Popravka jedanaest mesta ništa ne sprečava sutra. Zato provera, i to **namerno uska**,
sa tri granice:

| Granica | Zašto |
|---|---|
| samo **dvoindeksni** pristup `a(r, c)` | jednoindeksni je isti trošak, ali `Split()` rezultat, kolekcija i default-member poziv izgledaju isto |
| telo **ne sme da piše** u niz | rutina koja u njega upisuje mora ostati `ByVal`, inače bi menjala pozivaočev niz |
| telo **bez petlje**, do 12 naredbi | rutina koja niz sama iterira plaća **jednu** kopiju za ceo prolaz — to je zanemarljivo |

Bez treće granice je provera nad zatečenim kodom dala **51 nalaz** umesto 26 — većina
bezopasnih. Lažan nalaz uči da se checker ignoriše, pa je to bila razlika između
upotrebljive i štetne provere.

Osam self-test slučajeva: dva „mora da zapišti" i šest „ne sme". Dokaz u oba smera —
sa ugašenom proverom padaju tačno ta dva, po imenu.

**Provera je usput našla još 15 čitača** istog oblika — među njima `RowRank`, koji je
komparator sortiranja, dakle kopija cele tabele **po poređenju**. Popravljeni su svi:
checker ne sme da uđe crven.

Ukupno **26 čitača** u ovom PR-u.

Usput naplaćena i jedna zamka izvan VBA: `\b` u regexu je kroz shell heredoc postao
**pravi backspace znak** (`0x08`), pa `PETLJA` nikad nije poklapala. Sužavanje sa 51
na 15 je zato uradila samo granica dužine, a ne i granica petlje — i to se videlo tek
kad je self-test slučaj „rutina koja niz sama iterira" pao. Bez tog slučaja bi provera
ušla tiho pokvarena.

### Verifikacija

`vba_check` čisto (191) · self-test (39) · `who_writes` ažuran ·
`RunAllTests` **ZELENO (76)** · pun set **ZELENO** · `COMPILE` `NEJASNO`.

**Test 76** tvrdi da ljuska dobija baš ono što ekran broji (bez posrednika), da je
član ugovora opcion, i da poziv ekrana bez brojača **ne ostavlja `Err` postavljen**.

Dve sabotaže, obe po imenu:

| Sabotaža | Obara |
|---|---|
| `brojac-ekran-po-imenu` | ljuska dobija BAS ono što ekran broji (12 → 0) |
| `brojac-nije-opcion` | poziv ekrana bez brojača ne ostavlja `Err` (0 → 1004) |

Druga je usput otkrila da je moja prva verzija tvrdnje bila prazna: `On Error Resume
Next` već guta grešku i vrednost ostaje 0, pa sabotaža nad tim blokom nije obarala
ništa. Ono što blok stvarno radi je `Err.Clear` — i tek kad je test to počeo da meri,
sabotaža je dobila šta da obori.

Prva verzija je zvala samo `RefreshFromData`, pa su značke bile **prazne do prve
promene podataka** — a zaostatak koji se vidi tek pošto nešto uradiš ne rešava
ništa. Ceo smisao je da se vidi čim se aplikacija otvori.
`RunAllTests` **ZELENO (76)** · pun set **ZELENO** (Banka 189 · Storno 181 ·
Flow 336 · Palete 97 · Agrohemija 25).

Pun set je ovde obavezan, ne formalnost: izmenjeni čitači su u putanjama štampe,
kartice kooperanta i storno žurnala, koje `RunAllTests` ne dodiruje.

---

## v2.55.0 — `v6-ui-159` · unos prerade: Faza C, stavka 10

Poslednja stavka koja je tražila **proširenje ugovora ekrana**, a ne samo novu rutinu.

### Zašto je tražila seam

`UiChange` je ceo poznavao polja **unosnog** ekrana po imenu — `fgKgIT`, `cbKupac`,
`fgBrZbirT`. To je i tačno, jer su ta polja njegova. Ali ekran koji ima **svoja**
polja tu nije imao šta da traži: ljuska ne zna šta ona znače, a ne sme ni da sazna.

Klik je tu granu već imao (prefiks `scr` u `UiClick`); **promena teksta nije** — pa
ekran sa poljem za unos nije mogao ni da postoji. Sada se `scr*` promena prosleđuje
kao `chg:<tag>`, istim putem kao `act:` i `row:`.

`NewFieldG` je postao javan iz istog razloga: bez toga bi svaki ekran sa unosom
crtao svoju verziju istog polja, pa bi se razišli u izgledu i ponašanju.

### Četvrta lista: „Nova prerada"

Nije pregled kao ostale tri nego **radni ekran**: mreža služi da se palete označe,
zona nosi polja.

| Deo | Kako |
|---|---|
| Mreža | isti izvor kao lista „Palete", plus ✓ kolona napred i `PaletaID` nevidljiv pozadi |
| Izbor | klik na red štiklira — isti obrazac koji su otkupni blokovi dobili u `v6-ui-149` |
| Zona | **naraste** samo za ovu listu; `Scr_Layout` vraća visinu, pa je to već bilo u ugovoru |
| Polja | bruto · težina palete · gotov proizvod · kutije + tip · kese + tip · napomena |
| Neto | računa se **uživo**, kroz `chg:` seam |
| Upis | `SavePrerada_TX` — isti writer koji zove i legacy panel |

Izbor se drži **po `PaletaID`**, ne po broju palete: spisak završava u mutaciji.

Odlazak sa liste **poništava izbor** — palete pripadaju preradi koja se sprema, a
ostavljene bi sledeći put ušle u spisak koji operater nije video. Ista lekcija kao
kod označenih otkupnih blokova.

**Nijedna nova poslovna kapija.** Svih sedam provera je preneto iz legacy panela i i
dalje su pod prekidačem `VALIDACIJA_UNOSA`. Lista pokazuje **sve** palete, kao i
legacy — sužavanje ide kroz pretragu.

### Račun izdvojen iz prikaza

`NetoIzracun(bruto, težinaPalete, kutije, tipKutija, kese, tipKesa)` je javan i
čist. Bez tog izdvajanja jedina poslovna formula na ekranu ne bi mogla da se izmeri:
zona se crta nad formom koju harness gradi bez `.Show`.

**Test 77** tvrdi ugovor liste, kolone, izbor po identitetu i sam račun — uključujući
donju granicu: negativan neto nije podatak nego znak da unos nije potpun, pa je nula
iskrenija od minusa.

Dve sabotaže, obe po imenu: `prerada-sve-palete` i `prerada-neto-bez-ambalaze`
(*neto je bruto minus težina palete — očekivano [80], dobijeno [100]*).

### Compile je zaradio svoje mesto

Prvi run posle ovog koda je **visio 585 sekundi** i pao sa `The remote procedure call
failed`. Nije bio pad testa nego **projekat koji se ne kompajlira**:

| Simbol | Problem |
|---|---|
| `GAP` | `Private` u `modOtkupUI`, korišćen iz `modScrPalete` |
| `FmtBroj` | isto |
| `Zona()` | postoji u `modScrStorno`, ne i u `modScrPalete` |

`vba_check` nijedan ne vidi: `NEDEFINISAN` traži simbol koji **nigde** nije definisan,
a ova dva jesu — samo ne odavde. To je granica statičke provere i tačno onaj razlog
zbog kog `Debug → Compile VBAProject` ostaje ručna kapija.

Popravka nije bila da se `GAP` i `FmtBroj` otvore: ekran nema pravo da otvara ljuskine
konstante zbog sopstvenog rasporeda. Lokalni `PRE_GAP` i `Format$` — isti idiom koji
ostatak ekrana već koristi — manjeg su dometa.

### Verifikacija

`vba_check` čisto (191) · self-test (47) · `who_writes` ažuran ·
`RunAllTests` **ZELENO (77)** · pun set **ZELENO** · `COMPILE` — **ostaje na
operateru**, i ovog puta se videlo zašto.

### Smoke: padajuće liste i „koje palete su izabrane"

Dva nalaza iz prve probe, oba iz istog korena — polje u zoni ekrana nije bilo
ravnopravno sa poljem unosnog ekrana.

**1. Padajuće liste nisu radile.** Panel izbora je ljuskin i generičan (renderuje
`.List` bilo kog kombo polja), ali `FindCombo` ga je tražio samo u `zCtx` i
`zForm` — zonama unosnog ekrana. Kombo u zoni ugovornog ekrana nije nalazio, pa je
`PopKeyFor` vraćao prazno i klik na strelicu nije radio ništa.

Drugi uzrok je bio moj: grana `scr` u `UiClick` je gutala i klik na **strelicu**
(`scrPreGPD` počinje sa `scr` kao i sve ostalo što ekran nosi), pa je odlazio ekranu
— a on o panelu izbora ne zna ništa, niti treba. Strelica se sada rešava **pre**
grane ekrana. Isto i za kucanje: `PopFromTyping` se zove i za `scr` polja, pa kombo
ekrana ima strelicu koja radi **i** kucanje koje sužava listu.

**2. Nigde nije pisalo koje palete ulaze.** Zona je pokazivala samo brojku
(`2 izabrano paleta`) — a baš je izbor paleta odluka koju operater donosi.
Sada stoji i spisak **po broju palete**: `2 izabrano paleta:  185, 184`.

Izbor se i dalje **drži po `PaletaID`** — broj se ponavlja kroz godine — ali zona
ume da ga imenuje, kroz mapu `PaletaID → broj` koju punjenje redova ionako vidi.
Dug spisak se skraćuje: zona ima jedan red, a poenta je prepoznavanje, ne inventar.

### Panel prerade: jedinica nasred polja i „ostrva" belog

Dva izgledačka nalaza iz druge probe, i oba su bila ista greška: ekran je polje
**napravio**, ali ga nije **rasporedio**.

`NewFieldG` crta unutrašnjost prema početnih 180pt. Kad zona polje proširi na ~300pt,
labela, ivica, jedinica i sam unos ostaju na starim merama — pa `kg` završi nasred
polja, a unos izgleda odsečen sa leve strane. Unosni ekran to rešava
`LayoutFieldInner`-om; bio je `Private`, pa ga ekran nije mogao pozvati.

Sada je javan, iz istog razloga kao i `NewFieldG`: **ekran koji polje napravi mora i
da ga rasporedi.**

Drugi nalaz: svako polje je svoj **beli okvir**, a zona je krem — pa se između njih
videla pozadina i panel je izgledao kao niz odvojenih ostrva. Ispod celog panela
sada stoji jedna bela podloga, napravljena **pre** polja (u MSForms kasnija kontrola
stoji iznad), sa malim uvlačenjem levo i desno da prva labela ne bude zalepljena za
ivicu.

---

## v2.56.0 — `v6-ui-162` · dvoklik otvara stavke palete

Operater je tražio **padajuće redove ispod izabranog reda**: klik na paletu otvara
njene stavke, sa svojim zaglavljem, kao prozor ispod tog reda. UX je bolji i to
ostaje **zapisan prioritet za kasnije** — ali traži *vrstu reda* u ugovoru ekrana i
renderer promenljive visine, jer mrežu koriste **svi** ekrani, uključujući unos
dokumenata. To je posao reda veličine migracije jednog celog ekrana.

Do tada: **jedan potez umesto dva**, bez ijedne izmene u načinu na koji mreža crta
redove.

### Zašto dvoklik, a ne klik

Klik na red **bira** paletu, a radnje nad redom — *zatvori paletu*, *storniraj*,
*štampaj* — rade baš nad izabranom. Da klik prebacuje listu, do tih radnji se ne bi
moglo doći: prekidač bi se pretvorio pre nego što operater stigne da pritisne dugme.

Zato:

| Potez | Šta radi |
|---|---|
| klik | bira paletu — zona pokazuje koja je, radnje nad redom rade nad njom |
| **dvoklik** | otvara „Stavke palete", već sužene na tu paletu |

Naslov liste stavki i dalje nosi broj otvorene palete, pa se sa te liste zna čije su.

### Novi deo ugovora: `dbl:<red>`

Ljuska je dvoklik do sada trošila sama (`LoadRowIntoForm`) — što je tačno za **unosni**
ekran, koji jedini ima formu ispod mreže. Ugovorni ekran je ostajao bez njega.

Sada `UiDblClick` prosleđuje `dbl:<red>` ekranu, istim putem kao `row:` i `act:`.
Ljuska i dalje **ne zna nijedan ekran po imenu** — zna samo svoj početni.

### Bag koji je test našao pre operatera

Test je prvo pao sa **„Lista se nije učitala"**, iako se lista čitala ispravno.

`ScreenZone` vraća zonu ekrana, a kad je nema — `Nothing`. Grešku je preskakao
(`On Error Resume Next`), ali je **nije brisao**, pa je `Err` ostajao postavljen i
posle povratka. `ScrGridData` to čita kao **pad ekrana**: prazna mreža i poruka o
grešci nad podacima koji su potpuno ispravni.

U produkciji se ne vidi dok zona postoji. Vidi se tačno onda kad je najgore — kad
ekran **padne u gradnji**, jer tada zona nestane, a poruka o pravom uzroku se izgubi
iza lažne. Popravka je jedan red na izvoru (`If Err.Number <> 0 Then Err.Clear`), a
pogađa **svaki** ekran koji u `Scr_Rows` dopunjava svoju zonu.

Isti obrazac kao `T_LogErr_NeVidiErrPosleResumeNext`: `On Error` ne čisti za sobom
samo zato što je greška „obrađena".

### Verifikacija

**Test 78** tvrdi oba smera poteza (dvoklik otvara, klik ne), naslov otvorene palete,
postojanje radnje i to da pročitana lista **nije** pad ekrana. Mreža se puni pravim
podacima kroz `GridTestLoad` — tvrdo gejtovan seam koji zamenjuje samo `.Show`, ne
i put kojim klik ide.

Tri sabotaže, sve tri padaju po **svojoj** tvrdnji:

| Sabotaža | Pada na |
|---|---|
| `paleta-dvoklik-ne-otvara` | *dvoklik na paletu otvara njene stavke* — `[STAVKE]` vs `[PALETE]` |
| `paleta-klik-otvara` | *izbor reda ne traži ponovno čitanje mreže* — `[False]` vs `[True]` |
| `zona-curi-gresku` | *pročitana lista se ne prijavljuje kao pad ekrana* |

`vba_check` čisto (191) · `who_writes` ažuran · `RunAllTests` **ZELENO (78)** ·
`COMPILE` — **ostaje na operateru**.

## v2.57.0 — `v6-ui-169` · čipovi po ekranu, neto ulaz, i jedan skup dan

Faza C je zatvorena. Uz nju ide i najskuplja lekcija ove grane.

### Čipovi pripadaju ekranu (Faza C, stavka 11)

Ovo je bilo **poslednje mesto na kom je ljuska znala jedan ekran po imenu**:

```vba
ElseIf akt = "OTPREMNICE" Then
    ShowChip frm, k, (k = "chipSve" Or k = "chipOtvorene")
```

Sada svaki ekran prijavi svoje čipove kroz `Scr_Cipovi()`, istim oblikom kao radnje:
`kljuc:KATALOG:sirina`. Ljuska ključ **ne tumači** — vraća ga kroz
`Scr_Rows(filter, q)` onakvog kakvog ga je dobila.

Kontrole se ne prave iznova: ljuska ima **bazen od sedam slotova** (`ChipRow`), prvih
šest pripada režimima unosnog ekrana, ostatak se pozajmljuje. Tri mesta su morala da
nauče da slot može biti pozajmljen — `ChipFilter` (klik vraća *ekranov* ključ),
`RenderChipCounts` (brojač bi pregazio tuđi natpis) i `ApplyChipVisual` (crvena boja
pripada čipu „Otkazane", ne slotu).

Palete su dobile pet: **Sve · Ova godina · Otvorene · Zatvorene · Prerađene**.
Pravilo je izdvojeno u `PalCipProlaz` — čist račun bez mreže; nepoznat filter **pušta
sve**, da lista ne ostane prazna.

### Neto ulaz izabranih paleta

Panel je pokazivao samo **izlaz**. Sada levo od njega stoji **NETO ULAZ** — zbir neto
kilaže izabranih paleta sveže robe. Dve brojke jedna uz drugu su ceo račun prerade:
koliko ulazi i koliko gotovog izlazi.

Zbir ide iz **zapamćenih** vrednosti po `PaletaID`, ne iz mreže: mreža je filtrirana i
straničena, pa bi zbir po njoj zavisio od toga šta je trenutno na ekranu.

### Bela podloga panela je bila Frame — i to je koštalo ceo dan

Operater je prijavio panel u kom rade samo **prvo i poslednje polje**, bez naslova,
bez NETO, bez spiska izabranih i bez dugmeta.

`preBg` — bela podloga uvedena u `v6-ui-161` da se između polja ne vidi krem
pozadina — bio je **Frame**. U MSForms su Frame-ovi *prozorske* kontrole i crtaju se
**iznad** bezprozorskih (labela i svega sklopljenog od njih), **bez obzira na
z-order**. Komentar uz taj kod je čak tvrdio suprotno („kasnija kontrola stoji
iznad") — što važi među istom vrstom, ne između Frame-a i labele.

Podloga pokriva ceo panel, pa su ispod nje završili naslov panela, `NETO`, spisak
izabranih paleta i dugme „Preradi". Polja su se probijala jer su i sama Frame-ovi.

Popravka je **zamena vrste kontrole**, ne uklanjanje podloge: `preBg` je sada labela
sa belom pozadinom. Labela poštuje z-order, pa napravljena prva ostaje ispod svega.

**Šta je koštalo:** tri pokušaja popravke pre ovoga (`v6-ui-165` do `v6-ui-167`) — svi
zasnovani na *pretpostavljenom* mehanizmu, nijedan na reprodukciji. Sve tri su
oborene. Pravilo iz `CLAUDE.md` je jasno i bilo je prekršeno tri puta zaredom: kad se
ne može reprodukovati, **to je nalaz**, a ne osnov za zakrpu. Ono što je slučaj na
kraju rešilo bilo je operaterovo zapažanje — „radilo je dok smo sređivali vertikalni
prostor" — dakle **raspon commit-a**, a ne još jedna teorija.

### Bazen radnji: ista greška koju su čipovi imali izmerenu

`MAX_ACT = 5` — ljuska pravi tačno pet dugmadi za radnje nad redom, a lista paleta ih
je već koristila svih pet. Dodata šesta („Stavke") je tiho izbacila „Nepotpune
palete", jer `RefreshRowActions` radi `If i >= MAX_ACT Then Exit For`.

Šesta radnja se **ne vraća** — stavke se otvaraju dvoklikom. `MAX_ACT` je sada javan i
test tvrdi da lista ne traži više radnji nego što ljuska ima dugmadi. Ista kapija koju
su čipovi dobili kroz `MAX_CHIP` — da je postojala ujutru, greška ne bi ni nastala.

### Šta se namerno nije vratilo

Dve izmene rasporeda iz neuspelih pokušaja (`z.Height` pri rasporedu zone,
`Repaint`/`RenderGrid` pri ulasku na ekran). Pravi uzrok je bio drugde, pa im nema
osnova — a menjaju put kojim se crta **svaki** ekran.

### Granica testa, izmerena

Proba je namerno pozvala `Scr_Layout` nad okvirom od 10 poena i suite je ostao
**zelen**: `.Visible` je i dalje `True` kad MSForms ne prefarba. Kapija za tu klasu
kvara je **smoke**, ne `RunAllTests` — i to je sada zapisano u testu 80, da se ne
pokušava ponovo.

Uz to: dok `frmOtkupUI` živi, mašinerija forme briše `Err` između `Err.Raise` i
omotnice testa, pa pad stiže kao „greška bez opisa". Test 80 zato **skuplja** nalaze u
spisak imena i tvrdi tek **posle `Unload`-a** — pad tada kaže i koje kontrole fale.

### Verifikacija

**Testovi 78–80**: dvoklik u oba smera (otvara, a klik i dalje samo bira) + bazen
radnji · ugovor čipova, veličina bazena i to da čip *stvarno* sužava
(`Otvorene + Zatvorene = sve palete`) · zona: sve kontrole postoje, na listi za unos
upaljene, u pregledu ugašene. Test 77 je dobio i **neto ulaz**.

Osam sabotaža, svaka pada po **svojoj** tvrdnji — uključujući `radnji-vise-nego-dugmadi`
i `ulaz-bez-kilaze` (*neto ulaz je zbir neto izabranih paleta* — `[250]` vs `[0]`).

`vba_check` čisto (191) · self-test (47) · `who_writes` ažuran · `RunAllTests`
**ZELENO (80)** · pun set **ZELENO** (336 + 97 + 25) · **COMPILE i smoke prošli kod
operatera**.

---

## v2.58.0 — `v6-ui-170` · bazen ljuske više ne ćuti

Ljuska ima **bazene**: segmenata prekidača (`MAX_SEG`), dugmadi radnji nad redom
(`MAX_ACT`), čipova (`MAX_CHIP`) i kolona mreže (`MAX_COLS`). Svaki je tačno onoliko
veliki koliko je kontrola napravljeno pri gradnji.

Ekran koji zatraži više nije grešio — ali se višak gubio **bez ijedne poruke**. To se
dogodilo **dvaput**:

| Kada | Šta se videlo |
|---|---|
| `v6-ui-143` | jedanaesti čip se **crtao**, a klik nije imao kome da stigne |
| `v6-ui-162` | šesta radnja nad redom je tiho izbacila „Nepotpune palete“ |

Oba puta je kvar našao **operater**, ne suite. Oba puta je uzrok bio isti obrazac:
`If i >= MAX_X Then Exit For`.

### Šta se menja

Odsecanje **ostaje** — višak nema gde da se nacrta. Menja se to što odsecanje sada
dobija **ime**: koji ekran, koja lista, koja vrsta i koliko je traženo.

```
WARN modOtkupUI.BazenStaje | PALETE/PALETE/radnje: trazeno 6, bazen ima 5 -- visak se NE crta
```

Sva četiri mesta idu kroz isti čuvar (`BazenStaje`). Prijava je **jednom po**
**(ekran, lista, vrsta)** — ovo se zove pri svakom crtanju, pa bi prijava po pozivu
napunila log i sakrila ono što se stvarno dešava. Toast ide samo u dev buildu:
operater od te brojke nema šta da uradi, ali trag mora da ostane uvek.

### Test 81

Tvrdi **oboje**: da čuvar odseca **na veličinu bazena** (ne na nulu, ne na traženo) i
imenuje prekoračenje **jednom**, i da **nijedan današnji ekran** ne prekoračuje.

Druga tvrdnja ide kroz **registar ekrana**, ne kroz spisak imena — pa pokriva i
ekrane Faze E koji tek dolaze, bez ijedne dopune testa. To je i tvrdnja koja će pasti
onog dana kad neko doda osmi čip.

Sabotaže: `bazen-cuti-visak` (*prekoračenje se prijavljuje*) i `bazen-odseca-na-nulu`
(*višak se odseca na veličinu bazena* — `[5]` vs `[0]`).

`vba_check` čisto (191) · self-test (47) · `who_writes` ažuran · `RunAllTests`
**ZELENO (81)** · pun set **ZELENO** · `COMPILE` — **ostaje na operateru**.

---

## v2.59.0 — `v6-ui-171` · Agrohemija na novom UI-ju (Faza E, stavka 15)

Prvi ekran **Faze E**. `frmAgrohemija` ostaje operativna i **nepromenjena** — dve
kopije poslovne logike postoje namerno, kao kod `frmOtkup` i `frmDokumenta`.

### Šta je operater dobio

Dve legacy sekcije (izlaz i ulaz) bile su dve forme jedna pored druge. Ljuska ima
jednu zonu i jednu mrežu, pa su postale **prekidač režima**: IZDAVANJE / PRIJEM.
Obe korpe žive istovremeno — prelazak režima ne prazni ništa.

Mreža je dobila **četiri liste kojih legacy nema**: Korpa · Stanje magacina ·
Promet · Dug po kooperantu. Legacy je za isto morao u Izveštaje.

- **Traka korpe** uz desnu ivicu zone: naslov, poslednje stavke, zbir. Korpa se
  inače vidi samo dok je izabrana lista „Korpa" — operater koji gleda stanje ili
  dugove nema nijedan znak šta je upravo dodao. Polja uzimaju **ostatak** širine,
  isti raspored kao `PRE_DESNO` na Paletama; na uskom ekranu traka nestaje i polja
  uzimaju celu zonu.
- **Čipovi** — Stanje: Ima na stanju · Bez zaliha. Promet: Ulazi · Izlazi · Ova
  godina. Dugovi: Duguju. Negativno stanje **prolazi** kroz „Bez zaliha": greška u
  knjiženju koja se ne vidi ostaje neispravljena.
- **Brojač u meniju** broji korpu — jedino što na ovom ekranu čeka operatera.
- **Dvoklik** preuzima red u unos: iz Dugova kooperanta (i prebacuje u IZDAVANJE
  — dug se izdaje, ne prima), iz Stanja artikal.
- **„Ukloni stavku" / „Isprazni korpu"** — legacy pogrešnu stavku nije umeo da
  izbaci; jedini izlaz je bio zatvaranje forme.
- **Multiselect parcela** zamenjen sakupljanjem dugmetom „+ Parcela"; zbir ha koji
  smart doza računa je isti, kao i `parcelaID` niz razdvojen `;`.

Dva mesta u ovom ekranu poštuju isto pravilo: **lista koja se tiho odseca izgleda
kao cela.** Traka korpe kaže `… još N` kad stavke ne stanu, a granice bazena ljuske
(`MaxPrekidaca`, `MAX_ACT`, `MAX_CHIP`, kolone) sada tvrdi i ovaj ekran.

### Dvoklik bira po identitetu, ne po tekstu reda

Lista pokazuje **ime**, a bira se **kooperant**. Čitač liste gradi mapu
`prikaz → ID` i na koliziji pamti **prazno**; dvoklik na dvosmislen red tada
**odbija** da bira umesto da pogodi — isto pravilo kao „dvosmislen broj → MANUAL"
u storno okviru. Pogađanje bi ovde izdalo robu pogrešnom čoveku.

Fixture je zato dobio par istoimenih (`KOOP-TEST-1` i `KOOP-TEST-IME`, oba „Prvi
Testni"). Dug blizanca ide preko rezervisanog virtuelnog artikla, koji
`GetMagacinStanje` izuzima — pa se stanje `ART-TEST-1` ne pomera i dva testa se ne
vezuju jedan za drugi.

### Poslovna logika je izašla iz forme

`modAgroUnos` (novo) drži korpu, obe kapije stanja i transakciju — isti oblik kao
`modOtkupUnos` (F1), `modDokUnos` (F2–F4), `modNovacUnos` (F5–F7):

- kapija pri dodavanju **broji i ono što je već u korpi** — ista roba se ne može
  dodati dva puta preko stanja;
- kapija pred upis **agregira po artiklu preko cele korpe** — stanje se moglo
  promeniti između dodavanja i upisa (drugi operater, sync). Bez nje upis pukne na
  pola petlje i vrati se rollback-om, a operater dobije `4301` umesto rečenice;
- invarijanta nad `Pakovanje` svedena iz **tri kopije u formi** na jednu
  (`AgroArtikalInfo`);
- smart doza se zaokružuje **nagore**, na cela pakovanja — pola pakovanja se ne
  izdaje.

Čitači za mrežu su u domenskim modulima, ne u ekranu: `GetMagacinPrometForGrid`,
`GetAgroDugoviForGrid`, `modNovac.GetAgroAbzugMapa`.

### Odbitak duga se računa na dva mesta, pa se sada i poredi

`GetAgroAbzugMapa` je brza kopija pravila iz `GetAgroAbzug` — jedan prolaz umesto
`O(n·m)`. Obe su **žive u istoj funkciji**: mapu zove lista dugova, pojedinačnu keš
ekrana. Dve kopije istog pravila se tiho raziđu, i ista aplikacija na dva mesta
pokaže **različit dug istom čoveku**.

Fixture je zato dobio `AgroAbzug` redove (dva za istog kooperanta, jedan
**storniran**, jedan **drugog tipa**), a `T_Agro_AbzugMapaPratiPojedinacni` tvrdi
slaganje nad **svakim** kooperantom koga mapa zna — uz tačne zbirove, da ih ne
obori isti kvar na obe strane.

### Ljuska se ovim ne menja

Diff nad `src-vba/modOtkupUI.bas` je **jedna linija** — pečat verzije. Ekran je
*drugi korisnik* ugovora koji je Faza C otvorila (`NewFieldG`, `LayoutFieldInner`,
`chg:`, `FindCombo` nad zonom, `dbl:`, `Scr_Cipovi`, `Scr_Brojac`). Namerno: dve
fabrike polja znače dva izgleda istog polja.

### Šta je smoke našao

Dva kvara, oba uhvaćena **pre** nego što je išta isporučeno:

**Prekidač režima je belio.** Izabran režim je bio zelen samo dok je pokazivač nad
njim. Uzrok nije bojenje nego **pamćenje**: `clsFlatBtn` zapamti osnovnu boju pri
`Bind`-u i vraća je u `ResetVisual` kad pokazivač ode, a `BoxState` tu osnovu ne
dira. Lek je dvostruk — `RebaseSink` posle bojenja (isto kao `StilDugmeta` u
Stornu, gde je već jednom plaćeno) i vrsta `"seg"` umesto `"btn"`, jer prekidač
režima jeste segmentni prekidač kao onaj nad mrežom. Čipovi i prekidač lista su
bili zaštićeni od početka — zato su na istom ekranu radili ispravno.

**Korpa se nije videla.** Vidi „Traka korpe" gore. Desna polovina reda polja je
svejedno stajala prazna — parcele su često isključene, pa slotovi 3 i 4 nikad
ništa ne nose.

Nijedan od dva se ne vidi čitanjem koda, i nijedan suite tada nije mogao da uhvati.
Oba sada imaju test koji ih reprodukuje **bez miša**.

### Verifikacija

Testovi **82–90** (numeracija pomerena za jedan: `v6-ui-170` je uzeo 81), **trinaest**
agro sabotaža.

Pušteno na mašini sa Excelom, **zaključno sa testom 88**:

`vba_check` čisto (193) · self-test (47) · `who_writes` ažuran ·
`RunAllTests` **ZELENO (88)** · pun set **ZELENO** · deset agro sabotaža obara
**imenovani** test.

**Prvo puštanje je oborilo dva testa**, oba pisana nad fixture-om kakav nije —
produkcioni kod je bio ispravan u oba slučaja:

| Test | Zašto je pao |
|---|---|
| `T_Agro_KapijaStanjaBrojiKorpu` | kontrolni izlaz je upisivan sa **praznom parcelom**, a `PRACENJE_PARCELA` je ON → `4215`. Test je padao na svom **čistaču**, ne na kapiji koju meri. |
| `T_Agro_BrojacIDvoklikPoIdentitetu` | tražio je identitet kooperanta **koga u listi dugova nije bilo** — tvrdnja je merila odsustvo reda i prolazila bi samo greškom. |

Obe greške bi pale na prvom puštanju, a nijedna se ne vidi čitanjem — to je cena
testa koji je napisan bez izvršavanja.

> **Nije prošlo.** Testovi **89** i **90** (prekidač režima, traka korpe) i njihove
> **tri** sabotaže su pisani u sesiji bez Excela i **nisu izvršeni**. Fixture se
> mora **regenerisati** (`make_fixture.py`) — dobio je `ART-TEST-Z` sa velikom
> zalihom, bez ijednog izlaza, jer se preliv trake preko `ART-TEST-1` ne može
> izmeriti (kapija stanja propušta najviše tri pakovanja).
>
> **Compile** (`Alt+F11 → Debug → Compile VBAProject`) nije prošao nijednom.
> Na smoke listi ostaje i jedna stvar koju headless ne vidi ni u principu: vrsta
> `"seg"` menja **hover-in** — izabrano dugme se više ne zatamnjuje pod
> pokazivačem, kao ni prekidač lista ispod njega.

---

## v2.60.0 — `v6-ui-172` · Agrohemija: identitet stavke korpe i značka koja ne ćuti

Dorada ekrana iz `v6-ui-171` po nalazima review-a merged PR-a #213. Ljuska se ovim
ne menja — diff nad `src-vba/modOtkupUI.bas` je **prazan**.

### „Ukloni stavku" je izbacivao pogrešan red

Stavka korpe je tražena po **nazivu artikla i količini** iz prikazanog reda. Dve
iste stavke su tada nerazlučive, a to nije izmišljen slučaj: *„dva pakovanja sada,
dva kasnije"* daje dva reda iste robe i iste količine. Klik na drugi red je
izbacivao **prvi** — tiho, jer red koji nestane izgleda isto kao onaj koji je
trebalo da nestane.

Svaka stavka sada nosi **svoj identitet** (`stavkaID`), prolazan — živi koliko i
korpa, nikad ne ide u tabelu. Identitet putuje **u redu mreže**, u koloni
prioriteta 4 koju mreža nikad ne crta: sortiranje i stranice ga ne mogu razdvojiti
od stavke. Prazan ili nepoznat identitet **ne uklanja ništa** i javlja se porukom
— ne pogađa se.

Isto pravilo kao „dvosmislen broj → MANUAL" u storno okviru, samo što se ovde
dvosmislenost može **sprečiti** umesto prijaviti.

### Značka u meniju prati korpu i kad korpa nije prikazana lista

Operater gleda Stanje, doda tri stavke, a značka uz „Agrohemija" i dalje piše nulu
— pa pređe na drugi ekran misleći da nema šta da proknjiži.

Uzrok: ljuska brojače pita samo kad ekran javi „podaci su promenjeni", a ekran to
javlja samo kad je korpa prikazana lista (inače bi terao ponovno čitanje Stanja ili
Prometa koje se nije menjalo). **Korpa nije podatak u tabeli**, pa te dve stvari ne
smeju da dele isti kanal. Promena korpe sada ima svoj: dodavanje, uklanjanje,
pražnjenje i upis osvežavaju značku bez obzira na to koja je lista u mreži.

Prekidač režima tu ne spada — značka sabira obe korpe, pa prelazak sa izdavanja na
prijem ne menja broj.

### Verifikacija

Testovi **91** (`T_Agro_KorpaUklanjaPoIdentitetu`) i **92**
(`T_Agro_ZnackaPratiKorpuVanKorpeListe`); četiri nove sabotaže — ukupno
**sedamnaest** agro sabotaža.

Urađeno i prijavljeno kao zeleno: `vba_check` čisto (193) · self-test (47) ·
`who_writes` ažuran · sve četiri nove sabotaže se primenjuju i **uredno vraćaju**,
a izvor posle vraćanja ostaje bit-identičan.

> **Neverifikovano.** `RunAllTests` **nije puštan** — u sesiji nema Excela.
> Testovi 91 i 92 i sve četiri nove sabotaže **nisu izvršeni**. **Compile** nije
> prošao. Ništa od ovoga se ne prijavljuje kao zeleno.

### Test 89 — nalaz, ne zakrpa

Review javlja `T_ZonaAgro_PrekidacRezimaZadrzavaBoju` kao jedini crveni test.
**Nije reprodukovan** — nema Excela u ovoj sesiji, pa se ne zna koja je tvrdnja
pala i uzrok nije pogađan.

Ono što se **može** utvrditi čitanjem: test je tvrdnje postavljao **dok forma
živi**, a susedni test u istom fajlu dokumentuje zašto to ne valja — tada mašinerija
forme obriše `Err` između `Err.Raise` i omotnice testa, pa pad stiže kao
`greska bez opisa`. Test je zato prestrojen: forma se prvo **izmeri**, pa se tvrdi
**posle** njenog otpuštanja (koje je usput i dobio — do sada je zvao goli `Unload`).

To ne popravlja uzrok ako uzrok postoji, nego **dijagnostiku**: sledeće puštanje
je ili zeleno, ili imenuje tvrdnju i vrednost.

---

## v2.61.0 — `v6-ui-173` · paleta se bira po identitetu, ne po broju

### Broj palete nije identitet

Broj palete i broj prerade **resetuju se po godini** — `GenerateBrojPalete` i
`GenerateBrojPrerade` računaju `maxN + 1` unutar `Year(Date)`. Zato `12/2025` i
`12/2026` postoje istovremeno.

Ekran Palete je identitet rešavao kroz rečnik `broj → ID` (`mPalIds`, `mPreIds`).
Takav rečnik za dva zapisa istog broja ima **tačno jedan unos** — pa je radnja nad
starijom paletom pogađala **noviju**:

| Radnja | Šta se dešavalo |
|---|---|
| `palprint` / `palpdf` | štampa se pogrešan paletni list |
| `palzatvori` | zatvara se tuđa paleta |
| `palstorno` | **stornira se pogrešan zapis** |
| `preprint` / `prestorno` | isto, nad preradom |

Kvar je tih: operater vidi red koji je izabrao, a radnja ode na drugi zapis.

### Šta se menja

Obe liste dobijaju **nevidljivu kolonu identiteta** (`PaletaID`, `PreradaID`) —
isti oblik koji lista za unos prerade (`PAL_NOVA_COL_ID`) i ekran Storno
(`ST_BLOK_COL_ID`) već nose: širina `0`, prioritet `4`, uvek poslednja.

`PostaviAktivnu` i sve radnje čitaju ID **iz izabranog reda**. Kolona putuje sa
redom, pa preživi i sortiranje (`SortedView` kopira svih `mColN` kolona) i
filtriranje — jer je deo reda, a ne pogled sa strane. Rečnici `mPalIds` /
`mPreIds` su **obrisani**: rečnik koji izgleda kao izvor identiteta, a nije, samo
čeka sledećeg čitaoca.

### Telo mreže više ne ulazi u traku poruka

Traka stoji tačno iznad podnožja (`footTop - TOAST_H - 4`), a telo je računato sa
rezervom od **6pt** — pa je poslednji red ulazio **24pt u traku**. Poruka se
crtala preko reda i držala se samo `ZOrder`-om, što rešava **redosled crtanja**, ne
**prostor**: red ispod poruke je bio nečitljiv.

`LayoutGrid` sada rezerviše `TOAST_H`: `body.Bottom <= toast.Top`. Ispod ~195pt
i dalje pobeđuje pod od tri reda — mreža koja pokaže manje od tri reda nije
upotrebljiva.

### `Scr_Event` vraća čist `Err`

Ista obaveza koju `modScrStorno.Scr_Event` već drži. Ovde je cela funkcija stajala
pod `On Error Resume Next` i `Err` nikad nije čistila, pa je progutana greška
ostajala postavljena i posle povratka — **ljuska je prijavljivala neuspeh za radnju
koja je prošla**. Telo je izdvojeno u `ObradiDogadjaj`, omotnica loguje, javlja
toast i čisti `Err` u oba smera.

### Verifikacija

Testovi **93–96**, uz fixture redove koji koliziju uopšte prave: `PAL-TEST-Y25`
(`12/2025`, uz postojeću `12/2026`) i dve prerade istog broja u dve godine.
`tblPrerada` do sada nije imala **nijedan** red.

| Test | Šta meri | Sabotaža |
|---|---|---|
| `T_PaleteIdentitet_PoIDNePoBroju` | svaki red daje svoj `PaletaID`; drugi ostaje netaknut | `palete-id-po-broju` |
| `T_PreradeIdentitet_PoIDNePoBroju` | isto nad preradom, druga grana resolvera | `palete-id-po-broju` |
| `T_GridTelo_NePokrivaToast` | `body.Bottom <= toast.Top` na šest visina | `grid-telo-preko-toasta` |
| `T_PaleteScrEvent_NeCuriGreska` | `Err.Number = 0` i kad događaj iznutra pukne | `palete-event-curi-err` |

`vba_check` čisto (193) · self-test (47) · `who_writes` ažuran ·
sve četiri sabotaže obaraju **imenovani** test i vraćaju se bit-identično.

> **Suite je puštena PRE rebase-a na `v6-ui-172`.** Tada je `RunAllTests` imao
> **94 testa, 93 prolaze** (uz ostale suite ZELENO). Posle rebase-a ih je **96** —
> agro rad je uzeo 91 i 92, pa su testovi Paleta prenumerisani na **93–96**.
> **Nije ponovo puštena.** Ni renumeracija ni ta dva nova testa nisu mereni na
> ovoj grani; to ostaje uz compile.

> **Jedan test pada, i nije iz ovog rada.** `T_ZonaAgro_PrekidacRezimaZadrzavaBoju`
> (test 89, `v6-ui-171`) pada i na čistom `main`-u, pre ovog rebase-a —
> `neizabran rezim nije Bold: ocekivano [False], dobijeno [True]`. Došao je uz
> `bab97df7`, zajedno sa `OsveziPrekidacRezima` koji popravlja. Baseline nad
> `main`-om je snimljen pre rebase-a baš zato da se pad ne pripiše ovoj grani.

**Compile** (`Alt+F11 → Debug → Compile VBAProject`) **ostaje operateru** —
automatski verdikt je `NEJASNO`.

---

## v2.62.0 — `v6-ui-175` · rez fonta se potvrđuje, ne veruje mu se

Jedan kvar u ljusci koji je tri release-a držao jedan test crvenim — i sve vreme
menjao izgled **celog** novog UI-ja, a da se nije video.

### Šta je bilo

`modUiKit.NewLbl` je ignorisao traženi rez: kontrola građena sa `bold=False`
izlazila je sa `Font.Weight = 700`. Kvar je **uniforman** — svaka runtime
kontrola je bila bold — pa se nije ni primetio; izgledao je kao odluka dizajna.

Izašao je tek kad je jedna tvrdnja zatražila da **neizabran** segment
Agrohemije **nije** bold.

### Kako je nađen

Sonda u testu 89, šest krugova merenja nad živom formom. Svaki krug je gasio po
jedno objašnjenje: raspored, ožičenje, povratnu vrednost, nasleđivanje fonta
forme, artefakt čitanja, redosled upisa, `BackColor`. Nijedno nije preživelo.

Poslednje merenje je pokazalo zašto:

```
s1=400   upis rez=False nad izgrađenom kontrolom  -> PROLAZI
s2=400   upis BackColor nad istom kontrolom       -> font nedirnut
s3=700   JOŠ JEDAN isti takav upis rez=False      -> vrati 700
```

Upis `Font.Bold` **nije ni pouzdan ni idempotentan**.

### Popravka

`modUiKit.PostaviRez` — upiši, pročitaj, i ako nije ono što je traženo, upiši
opet; najviše tri puta. Merilo je `Font.Weight` (400 normalan, 700 bold), jer je
`Font.Bold` iz nje izveden i sam ume da prevari. Petlja je ograničena: ekran koji
se zavrti je gori kvar od pogrešnog reza.

Koriste ga `NewLbl`, `NewTxt` i `BoxState` — sva tri mesta na kojima se rez
uopšte postavlja.

### Šta operater vidi

- **Izgled.** Bold sada nosi samo ono što ga i traži: naslovi, izabrani segmenti,
  čipovi, brojčana polja, zbirovi. Ostalo prelazi u normalan rez. Vidljivo je
  kroz **ceo** novi UI.
- **Ponašanje.** `clsFlatBtn.IsSelected` čita baš taj rez i za `"nav"`, `"chip"`
  i `"seg"` je do sada bio **uvek True**, pa hover nije prefarbavao nijedno od
  njih. Sada razlikuje izabrano od neizabranog, kako je i projektovano.

### Verifikacija

`RunAllTests` **96 testova**, `T_ZonaAgro_PrekidacRezimaZadrzavaBoju` **prolazi**
prvi put otkad postoji. Test od sada čita `Font.Weight`, ne `Font.Bold`, i tvrdi
ga za ispunu **i** natpis oba segmenta, u gradnji **i** u rasporedu, jednom
tvrdnjom. Sonde su uklonjene — ostavljene bi tvrdile zatečena ponašanja
MSForms-a kao da su ugovor.

Sabotaža `ljuska-rez-bez-potvrde` vraća upis na jedan pokušaj bez čitanja i
obara test po imenu.

```
vba_check  cisto (193)  ·  self-test (47)  ·  who_writes azuran
RunAllTests 96 testova, 2 pala -- oba na PODACIMA DONORA, ne na kodu
```

> **Dva preostala pada nisu kod.** `make_fixture.py` u `KEEP_ROWS` ne briše
> `tblConfig`, pa fixture nasleđuje podešavanja donora:
> `T_PosleSnimanja_ZadrzavaKontekstOtpremnice` pada na `DEFAULT_SORTA_VOCA`
> (donor ima `Willamette`, golden je snimljen bez njega), a
> `T_IsplataValidiraj_TipNovcaPoIzboru` na `KES_ISPLATE` — isključen kod donora,
> pa grana `If B(p, "izAvansa") And IsKesIsplate()` ne uđe i validator vrati
> prazno. Oba padaju identično i pre ovih izmena.

> **Compile** (`Alt+F11 → Debug → Compile VBAProject`) ostaje operateru, i
> **smoke** nad novim rezom kroz sve ekrane — to je promena koju headless ne
> vidi.

---

## v2.63.0 — `v6-ui-176` · Fakturisanje na novom UI-ju (Faza E, stavka 16)

Drugi ekran Faze E. Stavka **Fakturisanje** u meniju više nije prigušena.

### Šta operater dobija

**Jedan ekran umesto forme i njenog SEF-a.** Levo se bira kupac, u mreži se vide
njegove prijemnice, a sve ostalo — izdate fakture i stanje elektronskih faktura —
stoji pod prekidačem iznad mreže. Nema više „vrati se pa otvori drugu formu".

**Tri liste:**

| Lista | Šta pokazuje |
|---|---|
| **Za fakturisanje** | prijemnice izabranog kupca; kvačica označava šta je već ubačeno u fakturu, poslednja kolona pokazuje broj fakture ako je prijemnica već fakturisana |
| **Fakture** | izdate fakture: broj, datum, kupac, iznos, **uplaćeno**, **ostatak** i status u boji |
| **SEF** | stanje elektronskih faktura: šta je poslato, šta je odbijeno, šta je zaglavljeno |

**Sakupljanje umesto obeležavanja.** Stara forma je tražila da se u listi
označi više redova pa se pritisne „Izradi fakturu". Mreža bira jedan red, pa se
prijemnice sada **skupljaju**: dugme „Dodaj u fakturu", ili prosto **dvoklik na
red** — isti potez i za dodavanje i za vraćanje.

Šta je u fakturi vidi se na **tri mesta** istovremeno: kvačica u samoj listi,
traka „Stavke fakture" uz desnu ivicu (poslednje dodato je gore, sa zbirom
ispod), i **broj uz stavku menija** — pa se neproknjižena faktura ne može
zaboraviti prelaskom na drugi ekran.

**Četiri brojke uz naslov:** koliko prijemnica čeka fakturu, koliko ih je u
korpi, koliki je iznos korpe i koliko izabrani kupac ukupno duguje po
neplaćenim fakturama.

**Radnje nad izabranim redom:**

- lista *Za fakturisanje* — **Dodaj u fakturu**, **Ukloni iz fakture**
- lista *Fakture* — **Štampaj**, **Osveži status** (proverava uplate i prevodi
  fakturu u „plaćeno" kad je pokrivena; stara forma to nikad nije nudila)
- lista *SEF* — **Pošalji**, **Osveži SEF**, **Otkaži**, **Storno**, **Oporavi**

**Filteri (čipovi) po listi:** prijemnice se sužavaju na *Za fakturisanje* ili
*Fakturisane* (zamena za staru kućicu „Prikaži fakturisane"), fakture na
*Neplaćene* / *Plaćene* / *Ova godina*, SEF na *Za slanje* / *U slanju* /
*Odbijeno* / *Greška*.

### Šta se nije promenilo

- **Stara forma `frmFakturisanje` i `frmSEF` rade i dalje, netaknute.** Kao i kod
  otkupa i dokumenata, dve kopije postoje namerno.
- **Račun fakture je isti** — `Prijemnica.Količina × Prijemnica.Cena`, i radi ga
  ista transakcija (`CreateFaktura_TX`), sa istim proverama: vlasništvo
  prijemnice, storno, dupli izbor, avans. Ekran ne računa ništa sam.
- **SEF moduli nisu dirani.** Ekran zove postojeće funkcije; `frmSEF` i dalje
  drži ono što ekran ne nosi — istoriju događaja po fakturi, pripremu ponovnog
  slanja i grupne radnje nad svim zaglavljenim fakturama.
- **Lista SEF se vidi i bez SEF naloga.** Stanje faktura je zapisano u samoj
  tabeli, pa se čita i bez veze; tek radnja (Pošalji, Otkaži…) traži upisane
  `SEF_BASE_URL` i `SEF_API_KEY` u `tblSEFConfig` i to kaže porukom.

### Sitnice koje se primete tek u pogonu

- **Broj fakture se ne unosi.** Dodeljuje ga upis; stiže u poruci čim faktura
  nastane i vidi se u listi. Polje sa „predlogom" bi umelo da pokaže jedan broj
  a upiše drugi.
- **Korpa se prazni kad se promeni kupac**, uz poruku. Faktura ne sme da meša
  prijemnice dva kupca, pa bi takva korpa svakako pukla pri upisu — bolje odmah
  i glasno. **Kucanje po polju kupca korpu ne dira** — tek stvarno izabran drugi
  kupac je promena; dok se pretražuje, sve što je sakupljeno ostaje.
- **Prijemnica koja je već fakturisana se odbija sa porukom**, i onda kad joj
  kolona sa brojem fakture izgleda prazna. Takvi redovi postoje u podacima i iz
  liste izgledaju slobodni.
- **Stornirana faktura se ne prikazuje** — pa joj se ne može ni ponuditi štampa
  ni slanje na SEF.
- **Faktura iznosa nula nije „plaćena"**, nego neplaćena. Prazna faktura i
  namirena faktura nisu isto stanje.

### Verifikacija

Testovi **97–103** u `modTest` i **šesnaest** novih sabotaža. `RunAllTests`
**103 testa, 0 palih.**

Dva pada koja su ranije stajala uz svaki PR (`T_PosleSnimanja_...` i
`T_IsplataValidiraj_...`) nisu bila ničija regresija nego **donorov config
koji je curio u fixture**: `make_fixture` čuva `tblSEFConfig`, pa je svaki
ključ koji fixture ne postavi ostajao onakav kakav je u svesci od koje se
polazi. `DEFAULT_SORTA_VOCA` i `KES_ISPLATE` se sada **pinuju**, pa ista suite
daje isti rezultat na svakoj svesci.

Fixture je dobio ono što do sada nije imao: fakture sa brojem, datumom i
statusom, uplatu vezanu za fakturu, storniranu fakturu i prijemnicu obeleženu
kao fakturisanu **bez** broja fakture. Bez tih redova su filteri radili nad
praznim skupom i bili zeleni bez pokrića.

**Compile i smoke su prošli.** `Alt+F11 → Debug → Compile VBAProject` je čist,
a ekran je proveren nad pravim podacima: izbor kupca, sakupljanje u fakturu,
izrada, lista faktura i SEF lista. Time je i poslednja kapija koju headless ne
vidi zatvorena — ovaj unos nema nijednu otvorenu stavku.

> Iz punog seta ostaju crvene `RunGoogleSyncSmokeSuite` i
> `RunMasterSyncSmokeSuite`. Ne tiču se ovog rada: traže Google kredencijale
> kojih u headless runu nema i padaju identično na netaknutom `main`-u.

---

## v2.64.0 — `v6-ui-177` · Uvoz izvoda na novom UI-ju (Faza E, stavka 17)

Stavka menija **„Uvoz izvoda"** se više ne crta prigušena — otvara se kao pun
ekran novog UI-ja, sa dve liste umesto jednog spiska. Uz njega su ispravljena i
**tri kvara u samoj mreži**, koji pogađaju svaki ekran novog UI-ja.

### Šta operater dobija

**Red za mapiranje u mreži.** Stavke izvoda se sortiraju klikom na zaglavlje,
pretražuju (partner, poziv na broj, broj izvoda, broj računa) i sužavaju
čipovima: *Sve · Za obradu · Jaki ključevi · Za ručno · Obrađeno · Preskočeno*.
„Za ručno" su stavke koje je automatika probala i vratila — do sada su se od
netaknutih razlikovale samo po slovu u koloni statusa.

**Kolona „Predlog" na svakom redu.** Šta bi automatika uradila sa tom stavkom —
*faktura 2/2026*, *blok 1/TEST*, *avans kupca X*, *nema jakog ključa*, *nejasan
smer* — vidi se za **sve** redove odjednom. Stara forma je isto to pokazivala
samo za izabranu stavku, u panelu ispod liste.

**Nova lista „Izvodi".** Po svakom uvezenom izvodu: početno stanje, uplate,
isplate, završno stanje, koliko stavki nosi i koliko ih je još otvoreno, i **da
li se slaže**. Stara forma je isti račun radila u jednom redu teksta i samo za
**najnoviji** izvod — izvod od pre dve nedelje koji ne štima nije se video
nigde.

**Pet radnji nad redom:** automatski mapiraj red, ručno mapiraj, preskoči,
mapiraj jake ključeve, automatski mapiraj sve.

**Ručno mapiranje u zoni:** tip (Kupac / Kooperant / OM), partner i cilj
(faktura ili blok). Izbor tipa se **predlaže sam** iz smera stavke — uplata
predlaže kupca, isplata kooperanta — kao i u staroj formi.

**Značka uz stavku menija** pokazuje koliko stavki još čeka, pa se to vidi i sa
drugog ekrana.

### Šta je zaštićeno

Ovo je ekran na kom se novac vezuje za dugove, pa je najveći deo posla otišao na
to da **nijedan kvar ne može tiho da postane drugo knjiženje**.

- **Stavka pogrešnog smera se odbija pre klika**, sa objašnjenjem, a ne kao
  greška pri knjiženju.
- **Blok sa više otvorenih stavki nego što automatika sme da podeli traži
  izričitu potvrdu:** prikaže se tačna podela koja bi bila proknjižena, uz izbor
  DA / avans / odustani.
- **Ručno mapiranje se zaustavlja ako lista za izbor nije učitana.** Prazna
  lista i neuspelo čitanje izgledaju isto, a prazan izbor nosi značenje: kod
  kupca je to **avans**, kod kooperanta **poziv na broj iz izvoda**. Važi za obe
  ručne rute.
- **Nedostupna tabela više ne prolazi kao prazna.** Excel na oba slučaja vraća
  isto, pa je nedostupna tabela faktura ili otkupa izgledala kao „nema redova".
- **Ručno mapiranje na blok pamti i otkupno mesto.** Broj otkupa je jedinstven
  po stanici, pa isti broj bloka može da postoji na dva otkupna mesta; lista ih
  prikazuje odvojeno (`12 · OM Naziv`), a raspodela ide samo na stavke izabranog
  mesta.
- **Blok kome otkupno mesto nije upisano se vidi, ali se ne knjiži.** U listi
  stoji označen (`12 · bez otkupnog mesta`), a ručno mapiranje ga odbija — inače
  bi raspodela zahvatila sva mesta sa tim brojem. Takve redove današnji unos ne
  pravi; nalaze se u starijim podacima.
- **Blok koji je već u celosti plaćen se odbija.** I dalje se vidi u listi, ali
  ako bi ga operater izabrao, ceo iznos bi se proknjižio kao **avans
  kooperanta** i stavka bi bila označena obrađenom — bez pitanja. Operater je
  rekao *koji* dug plaća, pa „nema šta da se plati" nije bezbedan ishod.
  Kad blok dolazi iz **poziva na broj** (operater nije birao), avans ostaje
  namerno ponašanje.
- **Stavka čiji se interni broj u tabeli pojavljuje dvaput vidi se u listi**,
  ali radnja nad njom odbija da pogađa koju od dve misli.
- **Isti broj izvoda sa dva različita računa firme su dva izvoda**, ne jedan —
  banke te brojeve dodeljuju po računu. Isto važi i za **isti broj na istom
  računu iz drugog ciklusa**: numeracija se ponavlja po godini, pa bi se dva
  izvoda spojila u jedan sa zbirom stavki preko oba.
- **Značka koja ne može da pročita podatak ne pokazuje nulu.** Ostaje poslednja
  poznata brojka, a greška se upisuje u log. Ako poznate brojke još nema — prvi
  pokušaj u sesiji — uz stavku menija stoji **`!`**, a brojke na ekranu su crte.
  Prazna značka znači „nema šta da čeka" i ne sme da se pojavi zato što čitanje
  nije uspelo.
- **Neuspelo čitanje se ne pamti.** Prolazan kvar je ranije mogao da zaključa
  isti izbor: knjiženje je bilo bezbedno blokirano, ali sledeći klik nije ni
  pokušavao ponovo.

### Tri ispravke u mreži — vide se na svakom ekranu

- **Kolone više ne kasne jednu listu.** Mreža je posle prebacivanja liste crtala
  sa širinama prethodne, pa je kolona koja je tamo bila skrivena ostajala prazna
  i kad joj je vrednost tačna — zaglavlje vidljivo, ćelije prazne. Vidi se na
  svakom ekranu čije se liste razlikuju po broju kolona, na primer na
  Fakturisanju pri povratku sa SEF-a na „Za fakturisanje".
- **Datum koji nije datum više ne kvari prikaz.** Na zatečenim sveskama datum
  transakcije ume da bude upisan kao broj (`26062026` umesto 26.06.2026). Takva
  vrednost je rušila ispis ćelije, tiho, i u njoj je ostajao tekst sa
  **prethodnog ekrana**. Sada se odbija i polje ostaje prazno.
- **Traka „Nema izabrane otpremnice…"** se više ne vidi tamo gde ne pripada —
  ostajala je upaljena sa ekrana Unos dokumenata.

### Šta se nije promenilo

- **Uvoz izvoda (povlačenje PDF-ova i parsiranje) ostaje gde je bio.** Novi ekran
  radi sa onim što je već uvezeno. Razlog nije dužina posla nego ishod:
  `ImportBankaInbox_TX` ne vraća broj uvezenih ni broj duplikata, pa bi dugme
  koje to ne može da kaže bilo tiho knjiženje. Ni stara forma uvozno dugme nema.
- **Stara forma `frmBankaImport` radi kao i do sada** i nije menjana — kao i sve
  druge stare forme tokom prelaska na novi UI.
- **Nalozi za isplatu (`Banka izveštaj`) nisu dirani.**
- **Račun se ne menja.** Ekran zove postojeće writer-e; podela po bloku i dalje
  ide kroz `PlanBlokRaspodela`, istu koju koristi i prikaz, pa se prikaz i
  knjiženje ne mogu razići.

### Sitnice koje se primete tek u pogonu

- **Interne šifre nema u prikazu.** Prva kolona je broj izvoda; `BankaImportID`
  operateru ne znači ništa. Identitet reda i dalje postoji, samo se ne crta.
- **Broj otvorenih i ukupan broj stavki stoje u jednoj koloni** (`10 / 16`),
  isto kao brojka „Mapirano" iznad mreže. Dve susedne brojke bez konteksta
  čitaju se gore od jedne sa kosom crtom.
- **„Obrađeno" i „preskočeno" su čipovi, ne liste.**
- **Izvod bez saldo podataka nije neslaganje nego odsustvo podatka.** Uvozi
  stariji od v6.18 te brojke nemaju, pa se ne prikazuju kao greška.

### Verifikacija

Testovi **104–112** u `modTest`, `RunAllTests` **112 testova, 0 palih**
(baseline je bio 103). **Trideset tri** sabotaže; svaka obara **tačno jedan**
imenovani test, a izvor se posle vraćanja vraća bit-identično.

`RunBankaImportTestSuite` (tvrd fail-gate nad ovim područjem) ostaje zelena:
**PASS=189, FAIL=0**. Pun podrazumevani set je zelen.

Fixture je dobio ono što do sada nije imao: **`tblBankaImport` je bio potpuno
prazan**, pa su sve tvrdnje o listama, čipovima i jakim ključevima radile nad
praznim skupom. Sada nosi dvanaest stavki izvoda u pet grupa, blok koji postoji
na tri otkupna mesta (od kojih jedno nije upisano), blok koji je u celosti
plaćen, i red čiji je datum transakcije upisan kao `ddmmyyyy` broj.

**Sedam nalaza je došlo iz smoke-a nad pravim podacima, a pet iz code
review-a** — među njima i tri koja bi svako moglo da završi kao pogrešno
knjiženje. Ceo tok, sa uzrocima i merenjima, zapisan je u
`docs/UI_MIGRACIJA_KATALOG.md` §9.10.

> **Compile i smoke još nisu izvršeni na finalnom SHA.** To je jedina otvorena
> stavka ovog unosa i namerno se ne prećutkuje: u ovom istom radu je ručni
> `Debug → Compile VBAProject` već jednom našao grešku (`Variable not defined`)
> **posle potpuno zelenog `RunAllTests`** — u telu procedure koju nijedan test
> nije zvao.

> Iz punog seta ostaju crvene `RunGoogleSyncSmokeSuite` i
> `RunMasterSyncSmokeSuite`. Ne tiču se ovog rada: traže Google kredencijale
> kojih u headless runu nema i padaju identično na netaknutom `main`-u.


---

## v2.65.0 — `v6-ui-178` · ćelija mreže ne ostavlja tuđi tekst

Ispravka u **samoj mreži**, pa se vidi na svakom ekranu novog UI-ja.

### Šta se moglo videti

Ćelija koja iz bilo kog razloga nije mogla da se ispiše **zadržavala je natpis od
ranijeg crtanja** — najčešće vrednost sa *prethodnog ekrana*. U koloni datuma je
tako umeo da stoji naziv vrste reda, u koloni iznosa broj iz druge liste.

Prazna ćelija je istina; tuđi tekst nije. Sada, kad se vrednost ne može
prikazati, ćelija **ostaje prazna**, a u dnevnik se upisuje jedan red sa
ekranom i kolonom.

Ranije je isto popravljeno samo za jedan slučaj (datum oblika `26062026` na
Uvozu izvoda). Ovim je zatvoren ceo razred: isto se dešavalo i sa kilogramima,
iznosima i statusnim oznakama kad vrednost nije bila onakva kakvu kolona
očekuje.

### Kolona DATUM na Fakturisanju je bila prazna

Nalaz je došao iz same ispravke, čim je crtanje počelo da broji ćelije koje ne
ume da prikaže: **lista faktura je imala prazan datum u svakom redu.**

Uzrok je Excel-ova sitnica sa dalekim posledicama — provera „da li je vrednost
broj" nad pravim datumom vraća **ne**. Uvoz izvoda je to zaobilazio pretvarajući
datum u broj; Fakturisanje nije, pa je kolona ostajala prazna. Nije bilo prijave
jer ništa nije pucalo, a nijedan test do sada nije čitao **nacrtan** datum.

Sada se datum prima i kao datum i kao broj.

### Sitnica koja se primeti tek u pogonu

Statusna oznaka (pilula) koja se ne može naslikati sada **nestaje** umesto da
ostane kao prazna obojena oznaka nad novim podatkom — i kad se podatak popravi,
vraća se tačno onakva kakva je bila, iste širine.

### Verifikacija

Test **113** (`T_MrezaCelija_NeostavljaTudjiTekst`) i tri nove sabotaže.
`RunAllTests` **113 testova, 0 palih.**

Test meri **oba nivoa**: pravilo (bez forme) i **samo crtanje** — nad pravom
formom, kroz isti raspored i isto iscrtavanje koje ide u pogonu, pa se čita
natpis same kontrole. To je bilo potrebno jer je pravilo tačno i kad se upis
preskoči, a ceo kvar je bio baš u tome.

Detaljno: `docs/UI_MIGRACIJA_KATALOG.md` §10.

> **Compile i smoke još nisu izvršeni na finalnom SHA.**

> Iz punog seta ostaju crvene `RunGoogleSyncSmokeSuite` i
> `RunMasterSyncSmokeSuite`. Ne tiču se ovog rada: traže Google kredencijale
> kojih u headless runu nema i padaju identično na netaknutom `main`-u.


---

## v2.66.0 — `v6-ui-179` · izabran blok koji nema šta da plati se odbija svuda

Ispravka u **knjiženju**, pa važi i za novi ekran i za staru formu.

### Šta se moglo desiti

Blok koji je već u celosti plaćen i dalje stoji u listi blokova — ona ne
proverava dug. Ako bi ga operater izabrao, ceo iznos bi se proknjižio kao
**avans kooperanta**, a stavka izvoda bila označena obrađenom. Bez pitanja, uz
uspešnu transakciju.

Operater je rekao *koji* dug plaća; „nema šta da se plati" tu nije bezbedan
ishod nego protivrečnost. Takav izbor se sada odbija uz objašnjenje i **ništa se
ne knjiži** — stavka ostaje otvorena.

**Kad blok dolazi iz poziva na broj** (operater ga nije birao), avans ostaje
namerno ponašanje — to je bezbedan izlaz dok je poreklo uplate dvosmisleno.

Novi ekran je ovo već odbijao od `v2.64.0`; sada isto važi i za **staru formu**
`frmBankaImport`, koja je do sada ulazila u tu granu bez ikakve provere.

### I kad se lista uopšte ne učita

Stara forma je do sada imala i tiši oblik istog problema: ako **učitavanje liste
blokova ne uspe**, prazan spisak je izgledao kao „kooperant nema blokova" — pa bi
se uzeo poziv na broj, a odatle bi ceo iznos opet završio kao avans.

Sada se to zaustavlja uz poruku. Forma je isto pravilo već imala za fakture;
blokovi su ga dobili. Uz to, **nedostupna tabela više ne prolazi kao prazna
lista** — ni za fakture ni za blokove, ni u samom knjiženju.

Pri odbijanju se sada dobija **jedna** poruka umesto dve: konkretna iz knjiženja,
bez generičke „nije izvršeno" preko nje.

### Verifikacija

`T21_IzabranPlacenBlokNijeAvans` u `RunBankaImportTestSuite` — suita koja piše
kroz transakciju sa rollback-om, pa se meri **stvarno knjiženje**, ne samo
odluka. `PASS=196, FAIL=0` (bilo 189). `RunAllTests` **113 testova, 0 palih.**

Test meri oba ishoda nad istim podacima: izabran blok → ništa se ne knjiži i
stavka ostaje otvorena; isti blok iz poziva na broj → avans, celim iznosom.

Detaljno: `docs/UI_MIGRACIJA_KATALOG.md` §9.10.

> **Compile i smoke još nisu izvršeni na finalnom SHA.**


---

## v2.67.0 — `v6-ui-180` · pravila stare forme se sada proveravaju automatski

Bez promene ponašanja za operatera. Ovo je posao na **proverama**.

### Zašto

Tri uzastopne ispravke u uvozu izvoda ticale su se iste stvari: **prazna lista je
bila protumačena kao izbor** — pa je uplata umela da postane avans. Sva tri puta
je to našao pregled koda, a ne automatski testovi, jer su ta pravila živela u
staroj formi, gde ih ništa osim čoveka nije moglo proveriti.

Sada mogu. Stara forma `frmBankaImport` dobila je uske provere koje rade **samo u
test režimu** i u radu ne postoje.

### Šta se sada proverava samo

- pad učitavanja liste blokova **zaustavlja** ručno mapiranje, uz objašnjenje;
- uredno učitana lista i dalje prolazi;
- **prazan izbor bloka nije izbor** — tada se koristi poziv na broj, gde je avans
  legitiman;
- kod kupca se blok uopšte ne prijavljuje knjiženju.

### Verifikacija

`RunAllTests` **114 testova, 0 palih.** Dve nove sabotaže — prve koje diraju
samu formu, a ne modul.

Detaljno: `docs/UI_MIGRACIJA_KATALOG.md` §11.

> **Compile i smoke još nisu izvršeni na finalnom SHA.**


---

## v2.68.0 — provera registra testova

Bez promene ponašanja aplikacije. Ovo je posao na **alatu za proveru koda**.

### Šta je bio rizik

Svaki test je upisan na **tri** mesta u `modTest`: šta se pušta, pod kojim se
imenom prijavljuje pad, i šta se stvarno zove. Ta tri spiska se održavaju rukom i
već su se razišla pri spajanju grana — a nijedan razlaz se ne vidi:

- test se broji, ali se **ne izvršava** — suite ostaje zelena;
- pad se prijavi pod imenom `T_Nepoznat_114`;
- test postoji i prolazi, ali se nikad ne pušta;
- `Case 114` zove telo testa 113 — oba „prolaze", jedan se ne izvrši.

### Šta se sada proverava

`vba_check` (isti alat koji već brani od ne-ASCII znakova i duplih imena) traži
da sva tri spiska nose **isti skup**, bez rupa u numeraciji i bez dupliranog
indeksa, i da se **ime slaže sa pozivom**.

Provera se okida sadržajem, ne imenom fajla, pa modul koji takav registar nema
ostaje netaknut.

### Verifikacija

Devet novih slučajeva u `--self-test` (56 ukupno, bilo 47), i **svaki je dokazan
u oba smera**: kad se pravilo namerno isključi, self-test pocrveni po imenu tog
slučaja. Uz to je provera puštena i nad **pravim** `modTest.bas` — uklonjen
`RunOne` i zamenjen poziv u `Case` grani oba puta prijavljeni sa tačnim brojem
linije.


---

## v2.69.0 — provera registra testova, drugi deo

Nastavak `v2.68.0`. Bez promene ponašanja aplikacije.

### Šta prvi deo nije video

Prva verzija je poredila **tri spiska međusobno**. Ali test koji je napisan a nije
upisan **nigde** ostavlja sva tri savršeno saglasna — i nikad se ne izvrši. Suite
ostane „114 testova, 0 palih", a testa u njoj nema.

To je poslednja klasa greške iz ove porodice: *zeleno, a nije ni pušteno.*

### Šta se sada proverava

| Pravilo | Šta hvata |
|---|---|
| telo nije registrovano | test napisan, zaboravljen u sva tri spiska |
| isti cilj pod dva indeksa | jedan test se izvrši dvaput, drugi nikad |

Drugo pravilo postoji jer provera duplog **indeksa** to ne vidi — indeksi su
različiti, a telo isto.

### Šta se namerno NE proverava

Da `Case` grana zove proceduru koje nema — to već hvata postojeća provera
`NEDEFINISAN`. Dva nalaza za isti kvar bi bila šum, a šum uči da se alat
ignoriše.

### Verifikacija

Četiri nova slučaja u `--self-test` (**64** ukupno, bilo 60), svaki izoluje tačno
jedno pravilo, i tri sabotaže nad samim alatom — od kojih jedna proverava da se
**ne javi** lažan nalaz nad pomoćnom procedurom.

Provera je puštena i nad pravim `modTest`: test uklonjen iz sva tri spiska i test
registrovan pod dva broja — oba prijavljena, sa tačnim brojem linije.


---

## v2.70.0 — `v6-ui-181` · zbirovi izvoda se proveravaju po redovima

Lista **Izvodi** je jedino mesto na kom se vidi da li se izvod slaže. Ovo izdanje
zatvara slučaj u kom je ta lista mogla da pokaže brojku koja ne važi za ceo
izvod, i sređuje dve stvari oko podnožja.

### Šta se moglo desiti

Podaci o izvodu — početno stanje, uplate, isplate, završno — upisani su na
**svaki** red tog izvoda. Lista ih zato **uzima** sa prvog reda, a ne sabira
(sabiranje bi ih pomnožilo brojem stavki).

To važi dok su svi redovi isti. Ako nisu — ručno ispravljen red, delimično
ponovljen uvoz — lista bi prikazala brojku **prvog** reda kao istinu o celom
izvodu. Bez ikakvog znaka.

Vredi znati koliko je podmuklo: u primeru iz testa **prvi red se sam za sebe
slaže** (4.500 + 500 − 0 = 5.000). Bez ove provere bi u koloni stajalo „slaže
se" — dakle tvrdnja o tačnosti brojki kojih zapravo nema.

### Šta se sada vidi

Takav izvod dobija svoj status: **`zbirovi se razlikuju`**.

To **nije** isto što i „ne slaže se" — tamo se zna šta piše pa se ne slaže, ovde
se ne zna ni šta piše. Zato takav izvod **ne ulazi** u čip „Ne slaže se": čip
nosi jedno tvrđenje, a o njemu se ne zna ništa. Vidi se u koloni.

**I nema brojke nigde.** Četiri novčane kolone su prazne, i **ne ulazi u promet**
u podnožju. Prikazati vrednost jednog od dva neslažuća reda značilo bi ponuditi
tuđi podatak kao saldo.

### Dve stvari oko podnožja

- **Zbir sada poštuje i pretragu.** Ranije je izvod koji pretraga sakrije i dalje
  ulazio u promet, pa je traka tvrdila promet redova kojih na ekranu nema. Stariji
  propust, nađen uz ovaj rad.
- **Nula se u te četiri kolone više ne ispisuje.** Uredan izvod bez isplata
  pokazuje praznu ćeliju umesto `0,00`. Na izvodu su to gotovo uvek ista stvar;
  jedini izuzetak je nov račun sa nultim početnim stanjem.

### Sitnica koja se primeti tek u pogonu

Te četiri kolone više nisu podebljane. Cena je toga što sada umeju da ostanu
prazne; četiri podebljane novčane kolone u istom redu ionako nisu davale
hijerarhiju. Poravnanje ostaje desno.

### Verifikacija

`RunAllTests` **114 testova, 0 palih**, uz **48** sabotaža u dvosmernom dokazu —
svaka obara **tačno jednu** imenovanu tvrdnju, i mašinski je provereno da nijedna
tvrdnja nije deljena između dve sabotaže.

`RunBankaImportTestSuite` ostaje **PASS=196, FAIL=0**; pun podrazumevani set je
zelen.

Test-sveska je dobila par redova istog izvoda koji nose različite zbirove — i po
stanju i po prometu, jer bez ovog drugog podnožje ne bi imalo šta da meri.

Ceo tok, sa tri greške koje je dokaz našao na samom testu, zapisan je u
`docs/UI_MIGRACIJA_KATALOG.md` §12.

> **Compile i smoke još nisu izvršeni na finalnom SHA.**

> Iz punog seta ostaju crvene `RunGoogleSyncSmokeSuite` i
> `RunMasterSyncSmokeSuite`. Ne tiču se ovog rada: traže Google kredencijale
> kojih u headless runu nema i padaju identično na netaknutom `main`-u.

## v2.71.0 — `v6-ui-182` · podnožje mreže više ne broji novac u komadima

Traka ispod liste je jedinicu birala prema režimu **unosa dokumenata** — a taj
režim ugovorni ekrani (Uvoz izvoda, Fakturisanje, Palete…) nemaju.

### Šta se videlo

Ko je radio reverse (`F7`) pa otvorio **Uvoz izvoda**, u podnožju je pisalo:

> `Ukupno 8.950 kom`

umesto

> `Vrednost 8.950,00 RSD`

Zbir je bio **tačan** — promet izvoda — ali sa tuđom jedinicom i bez para. Traka
se popravljala sama od sebe čim se uđe u neki od dokumentnih režima koji broji
dinare, pa je izgledala kao nasumična.

Isto je nekad radila i traka otpremnice koja je ostajala upaljena na tuđem
ekranu: stanje jednog ekrana koje se vidi na drugom.

### Šta je sada

Traka pita **ekran** šta broji, umesto da gleda u zajedničku promenljivu.
Reversi i dalje broje komade, sve ostalo dinare — a ekran koji o tome ne kaže
ništa dobija **dinare**, ne komade.

### Isto je važilo i za Storno

Ekran **Storno** ume da prikaže listu reversa, i tada mu u podnožju stoji broj
komada — a i on je do sada zavisio od zatečenog režima. Sada i on kaže šta broji,
pa `125` reversa više ne može da se prikaže kao `Vrednost 125,00 RSD`. Na ostalim
tipovima (fakture, izvodi, otkup) ostaje dinar.

### I jedna sitnica u sidebaru

Dole levo se opet vidi **verzija programa**. Tu je privremeno stajala oznaka UI
builda, dok je trajao rad na storno ekranu; taj je odavno gotov. Kome oznaka
builda treba za merenje, dobija je uz `UI_DEBUG=DA` u `tblLocalConfig`.

### Verifikacija

`RunAllTests` **115 testova, 0 palih**. Novi test meri obe strane: da ugovorni
ekran ne nasledi tuđu jedinicu, i da reversi svoju ne izgube — jer „svuda
dinari" bi prošlo prvu proveru, a tiho pokvarilo drugu.

Dvosmerni dokaz je pušten nad **39** sabotaža koje gađaju izmenjene fajlove.
Četiri nove obaraju tačno svoju tvrdnju; tri zatečene rupe u katalogu sabotaža
(nastale ranije, na `main`-u) zapisane su u `docs/UI_MIGRACIJA_KATALOG.md` §13.11
i idu kao zaseban posao.

Detalji: `docs/UI_MIGRACIJA_KATALOG.md` §13.

> **Compile i smoke još nisu izvršeni na finalnom SHA.**

## v2.72.0 — dokazni aparat: deset mrtvih sidara

Nema izmene u programu — ovo je izdanje o **proverama**.

### Šta je bilo pokvareno

Uz svaku automatsku proveru u projektu ide i „sabotaža": namerno pokvaren kod
koji **mora** da obori baš tu proveru. To je jedini dokaz da provera nešto meri.

Ispostavilo se da je **deset od 222** sabotaža prestalo da radi: kod koji su
gađale je u međuvremenu popravljen, pa se sabotaža više nije mogla ni primeniti.
Za tih deset provera dokaza više nije bilo — a to se nije videlo, jer takva
sabotaža ne javlja „prošlo je" nego „ne nalazim mesto", usred izlaza koji traje
pola sata.

Uz to: jedna sabotaža koja stvarno ništa nije obarala, jedno zastarelo ime testa,
i jedna provera koja je bila **ispravna** ali je izgledala mrtva jer je izveštaj
pucao na jednom slovu sa kvačicom.

### Šta je urađeno

Svih deset je popravljeno i dokazano ponovo. Dodata je provera koja ovo ubuduće
hvata **za sekundu** umesto za dva i po sata, i vezana je za istu proveru koja se
ionako pušta posle svake izmene koda — pa mrtvo sidro više ne može da prođe
neprimećeno.

Izveštaj o rezultatu više ne puca zbog znaka koji konzola ne ume da ispiše.

### Zašto je ovo vredelo

Zelena provera koja nikad nije pokazana crvenom ne dokazuje ništa. Deset takvih
je stajalo u projektu i izgledalo kao pokrivenost.

Ceo tok je zapisan u `docs/engineering/postmortems/2026-08-verifikacija.md` §10.
## v2.73.0 — `v6-ui-183` · podnožje liste izvoda pokazuje uplate i isplate

Traka ispod liste **Izvodi** je do sada imala jedan novčani podatak — **promet**,
dakle uplate i isplate sabrane u jedan broj.

### Šta je bio problem

Taj zbir se ne može rastaviti unazad. Izvod koji operater drži u ruci ima uplate
i isplate odvojeno, i to su brojke koje poredi. `Promet 12.400,00 RSD` ne govori
ništa o tome da li se izvod slaže — a zbog toga se ta lista i otvara.

### Šta je sada

Podnožje pokazuje **oba broja**:

> `Uplate 9.100,00 RSD     Isplate 3.300,00 RSD`

Isto važi i za listu **Stavke**. Brojke prate čip i pretragu isto kao i redovi —
kad se lista suzi, suze se i one.

### Verifikacija

`RunAllTests` **116 testova, 0 palih**, uz pet novih sabotaža. Jedna od njih
namerno crta **isti** broj u oba slota — najtiši mogući kvar, jer podnožje tada
izgleda savršeno ispravno.

Detalji: `docs/UI_MIGRACIJA_KATALOG.md` §14.

> **Compile i smoke još nisu izvršeni na finalnom SHA.**

## v2.74.0 — poruka o nedostajućoj koloni kaže i šta je videla

Nema izmene u radu programa — ovo je o **poruci koja se pojavi kad nešto pođe po zlu**.

### Šta je bilo

U logu se javilo `Nedostaje kolona 'VozacID' u tabeli 'tblZbirna'` — nad sveskom u
kojoj ta kolona **postoji**. Ista rečenica opisuje tri različita stanja: kolone
stvarno nema, tabele nema, ili je zaglavlje drugačije od očekivanog. Iz nje se ne
može znati koje je od ta tri.

### Šta je sada

Uz poruku ide i zaglavlje koje je program **stvarno pročitao**:

> `Nedostaje kolona 'VozacID' u tabeli 'tblZbirna'.`
> `Vidjeno zaglavlje: ZbirnaID, Datum, VozacID, BrojZbirne, ... (+21).`
> `Trazena kolona VIDJENA, pozicija 3.`

Poslednji red je onaj koji zapravo odgovara: ako program kaže da kolone nema,
a sveže čitanje je **vidi**, problem nije u tabeli nego u putu do nje.

Kad se sledeći put pojavi, iz same poruke se vidi da li je problem u šemi ili u
čitanju — bez ponovnog pokretanja i pogađanja.

### Šta NIJE urađeno

Uzrok nije reprodukovan i nije popravljan naslepo. Zapisan je ceo tok, sa dve
dijagnoze koje su merenjem oborene, u
`docs/engineering/postmortems/2026-08-verifikacija.md` §11.

## v2.75.0 — statusna oznaka se briše cela, i to je sada provereno

Nema izmene u radu programa — provera je ta koja je dopunjena.

### Šta je bilo zapisano

Statusna oznaka („pilula") u listi se, kad se vrednost ne može prikazati, briše
**cela** — i natpis i podloga. Da podloga stvarno nestaje **nije bilo provereno**;
u zapisu je stajalo da se ta lista ne može otvoriti bez ekrana, pa da se ne može
ni izmeriti.

To nije bilo tačno. Lista se otvara, samo je trebalo izabrati je onako kako je
bira i operater.

### Šta je sada

Provera radi nad **pravom listom dokumenata**: pilula se naslika, vrednost se
pokvari, i traži se da nestane i podloga — ne samo slova. Obojena podloga bez
natpisa i dalje tvrdi stanje koje nema pokriće.

Provereno je i da se posle ispravke vrednosti podloga **vrati**.

Detalji: `docs/UI_MIGRACIJA_KATALOG.md` §15.

## v2.76.0 — novac više ne postaje avans kad lista blokova zakaže

### Šta se moglo desiti

Na unosu ambalaže i novca za kooperanta (`Izlaz OM`), program bira gde novac ide
ovako: ako je iz liste izabran **otkupni blok**, knjiži na njega; ako nije,
knjiži kao **avans kooperanta**.

Ta lista je ostajala prazna i onda kad se **nije uspela učitati** — a prazna lista
je čitana kao „nema bloka". Novac je tada tiho odlazio u avans, bez ijedne poruke.
Razlika se videla tek kasnije, u saldu.

Tri različita stanja završavala su isto:

| Stanje | Šta je značilo | Šta bi trebalo |
|---|---|---|
| kooperant nema otvorenih blokova | avans | avans ✔ |
| lista se nije učitala | **avans** | stop |
| **tabela otkupa nedostaje** | **avans** | stop |

Poslednje je bilo najtiše: kad tabele nema, program dobija **istu praznu listu**
kao kad blokova stvarno nema — nijedna greška se ne javi.

### Šta je sada

Ako učitavanje zakaže, unos **staje** i operater dobija objašnjenje, sa razlogom
zbog kojeg lista nije učitana. To važi i kad tabela otkupa nedostaje.

Kad je lista uredno učitana a prazna, ništa se ne menja: to stvarno znači da
kooperant nema otvorenih blokova, pa je avans ispravan.

### I kad je učitavanje palo na pola

Ako punjenje liste pukne **usred posla**, u njoj bi ostalo nekoliko blokova — pa
bi izgledala kao potpuna. Operater bi izabrao jedan od njih, a pošto izbor
postoji, provera se ranije uopšte ne bi ni postavila.

Sada je obrnuto: prvo se pita da li je lista uopšte u redu, pa tek onda da li je
nešto izabrano. Nepotpuna lista se uz to **prazni**, da ne bi glumila potpunu.

### Odakle je došlo

Ista greška je ranije nađena i zatvorena na uvozu izvoda (`v6-ui-180`), uz belešku
da isto važi i za ovaj ekran. Kad se pogledalo, tamo je stvarno stajala — i to nad
novcem.

### Verifikacija

`RunAllTests` **119 testova, 0 palih**, uz tri sabotaže koje svaka obara svoju
tvrdnju: pad učitavanja mora da zaustavi, uredna lista mora da prođe, i **izabran
blok ne sme da zaobiđe proveru**.

Detalji: `docs/UI_MIGRACIJA_KATALOG.md` §16.

> **Compile i smoke još nisu izvršeni na finalnom SHA.** Ovo dira tok novca, pa
> smoke uključuje i slučaj sa nedostajućom tabelom.

## v2.77.0 — ista provera i na strani kupca, i u novom ekranu

### Šta se moglo desiti

Prošla verzija je zaustavila jedan slučaj: novac koji postaje **avans kooperanta**
kad se lista otkupnih blokova ne učita. Ista stvar je stajala na još tri mesta —
i sva tri su nad novcem.

| Gde | Šta je postajalo |
|---|---|
| Izlaz kupcu, stara forma | **avans kupca** umesto uplate po fakturi |
| Isplate (F5), novi ekran | **avans kooperanta** umesto razduženja bloka |
| Uplate (F6), novi ekran | **avans kupca** umesto uplate po fakturi |

Svuda isti niz: lista se ne učita → ostane prazna → prazno se pročita kao „nema
fakture" ili „nema bloka" → novac ode u avans. Bez ijedne poruke; razlika se vidi
tek u saldu partnera.

**U novom ekranu je bilo i tiše nego u staroj formi.** Tamo se pad učitavanja nije
upisivao ni u dnevnik grešaka, pa nije ostajao nikakav trag — ni na ekranu, ni
kasnije u logu.

**Tri stanja su i dalje završavala isto**, kao i prošli put:

| Stanje | Šta je značilo | Šta bi trebalo |
|---|---|---|
| partner nema otvorenih stavki | avans | avans ✔ |
| lista se nije učitala | **avans** | stop |
| tabela nedostaje | **avans** | stop |

### Šta je sada

Ako se lista ne učita, unos **staje** i operater dobija objašnjenje sa razlogom —
na sva tri mesta, i u staroj formi i u novom ekranu. Pad se od sada i **upisuje u
dnevnik**, pa postoji i posle zatvaranja programa. Isto važi kad odgovarajuća
tabela nedostaje.

Kad je lista uredno učitana a prazna, ništa se ne menja: to stvarno znači da
partner nema otvorenih stavki, pa je avans ispravan.

### Provera je uska koliko i kvar

Namerno **ne** staje:

- unos same ambalaže, bez novca — tu nema odluke faktura/avans;
- isplata otkupnom mestu — ona otkupne blokove ne dodiruje;
- režimi koji te liste uopšte nemaju.

Bez tih ograničenja bi jedan pad čitanja zaustavio i posao koji nikad ne bi bio
pogrešno proknjižen. Svako od tih ograničenja ima svoju proveru — i to je ono što
je review na ovom poslu i našao: jedno od njih je bilo tvrdnja bez provere, pa je
provera dopisana.

### Odakle je došlo

Prošla verzija je zatvorila jedno mesto. Ovo je bilo pitanje **ko još iz praznog
polja zaključuje tip novca** — i odgovor su bila preostala tri.

### Verifikacija

`RunAllTests` **121 test, 0 palih**, uz dvanaest sabotaža — svaka obara tačno
svoju tvrdnju, i ni jednu tuđu.

Detalji: `docs/UI_MIGRACIJA_KATALOG.md` §17.

> **Compile i smoke još nisu izvršeni na finalnom SHA.** Ovo dira tok novca u dve
> forme, pa smoke ide na kopiji sveske i uključuje slučaj sa nedostajućom tabelom.

## v2.78.0 — otkazan dokument više ne može da se pojavi kao živ

### Šta se moglo desiti

Storniranje ne briše dokument — red ostaje u tabeli, obeležen kao storniran, i
program ga izbacuje iz svih pregleda i zbirova. To izbacivanje radi **jedno
mesto**, kroz koje prolazi skoro svako čitanje u programu.

To mesto je prvo tražilo kolonu „Stornirano". Ako je nije našlo, tiho je vraćalo
**sve redove, zajedno sa storniranim** — kao da nema šta da se izbaci.

Posledica: otkazana faktura mogla je da se pojavi u listi otvorenih faktura, pa da
uplata ode na dokument koji više ne važi. Isto za otkazan otkupni blok.

Ovo je šire od prethodne dve verzije. Tamo je novac dobijao **pogrešnu vrstu**;
ovde otkazan dokument dobija **pogrešno postojanje**.

### Šta je sada

Program zna **koje tabele moraju** da nose kolonu „Stornirano", a koje je nemaju
uopšte:

| Vrsta tabele | Kad kolona nedostaje |
|---|---|
| dokumenti (otkup, novac, otpremnice, fakture, palete…) | **staje uz grešku** |
| matični podaci (kooperanti, kupci, vozači, stanice…) | prolazi — nikad je i nisu imali |

Za prvu grupu poruka kaže i **da li je kolonu videla** u svežem pogledu, pa se
razlikuje „kolone stvarno nema" od „šema se razišla".

Provera radi i kad je tabela **prazna**: nedostatak kolone je isti problem bez
obzira na to ima li trenutno redova.

### Jedno mesto je i dalje moglo da progura storniran red

Provera je bila na pravom mestu, ali je jedan ekran (detalji kartice) bio napisan
tako da **prećutkuje svaku grešku** — pa bi prećutao i ovu, i nastavio sa
nefiltriranim spiskom. Prepravljen je: radije bez podatka nego sa pogrešnim.

Pregledani su **svi** ostali pozivi te provere; drugih takvih nije bilo. Uz to je
dodata kontrola koja ne dozvoljava da se isti propust ponovo napiše neprimećeno.

### Provera zdravlja sveske to sada vidi unapred

„Provera zdravlja" je za stavke faktura tražila sve kolone osim „Stornirano" — pa
je mogla da javi zdravo stanje za svesku nad kojom program sada staje. Sada traži
i nju: bolje da se razilaženje vidi **pre** rada nego usred njega.

### Šta uraditi ako program stane

Poruka imenuje **tabelu i kolonu**, i kaže da li je kolonu videla. Dva slučaja:

- **kolona stvarno nedostaje** — vratiti je u tabelu (ili vratiti svesku iz
  rezervne kopije); dok je nema, ti pregledi ne rade;
- **kolona postoji, a poruka kaže da je viđena** — to je razilaženje šeme; javiti
  uz tekst poruke, jer ona nosi celo zaglavlje koje je program pročitao.

Prećutan rad sa netačnim podacima više nije opcija — to je i bila poenta ove
verzije.

### Verifikacija

`RunAllTests` **122 testa, 0 palih**, uz šest sabotaža — svaka obara tačno svoju
tvrdnju.

Uz to su i alati dobili dva nova pravila: jedno ne pušta čitanje nad tabelom koju
program nije klasifikovao, drugo ne dozvoljava da se pad ove provere prećuti. Oba
su dokazana i u „crvenom" smeru, uključujući i njihove izuzetke.

Detalji: `docs/UI_MIGRACIJA_KATALOG.md` §18. Domen: `docs/DOMEN/README.md`.

> **Compile i smoke još nisu izvršeni na finalnom SHA.** Ova verzija menja
> ponašanje čitanja kroz koje prolazi skoro sve, pa smoke ide na **kopiji**
> sveske, i prvi korak je običan rad — on dokazuje da vaša šema ima kolonu tamo
> gde je program očekuje.
