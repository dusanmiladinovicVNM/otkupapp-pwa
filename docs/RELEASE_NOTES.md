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
