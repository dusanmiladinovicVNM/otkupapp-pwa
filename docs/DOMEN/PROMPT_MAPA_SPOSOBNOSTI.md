# Prompt: mapa sposobnosti (AgriX / OtkupApp)

Ti si analitičar koda u repou otkupapp-pwa (Excel/VBA aplikacija za otkup voća; VBA u src-vba/, PWA u src/, Google Apps Script u gas/).
Radiš SAMO ČITANJE. Svaka tvrdnja nosi fajl:linija. Ako nešto nisi proverio, piši NEPROVERENO — nikad ne nagađaj.

## Zašto
Aplikacija se prepisuje na nov model podataka: dokument = zaglavlje + stavke (tblOtkup/tblOtkupStavke,
tblOtpremnica/tblOtpremnicaStavke + tblOtpremnicaIzvori, tblZbirna/tblZbirnaStavke + tblZbirnaIzvori, prijemnica, faktura, paleta).
Stari kod NE mora da ostane funkcionalan. Mora se sačuvati SAMO SPISAK SPOSOBNOSTI — sve što operater danas može da uradi ili dobije.
Tvoj izlaz je taj spisak. Ne predlaži kako čuvati stari kod.

## Oblast ove sesije
<OBLAST> — jedna od:
 A ekran DOKUMENTI (modScrDokumenti, modOtkupUnos, modDokUnos, modOtkupUI) — unos/izmena/štampa otkupa, otpremnice, zbirne, prijemnice, reversi, blokovi, izgubljeni blokovi
 B STORNO i OPORAVAK (modScrStorno, modStornoDok, modStornoFlow, modStorno, modStornoRecovery, modStornoZurnal, modScrOporavak)
 C IZVEŠTAJI, SLEDLJIVOST, PALETE (modScrIzvestaji, modIzvestaj, modScrSledljivost, modSledljivost, modScrPalete, modPaletniList, modPrint)
 D FAKTURE, BANKA, NOVAC, AGRO, ANALIZA (modScrFakture, modFaktura, modSEF*, modScrBanka*, modBanka*, modNovac*, modScrAgro, modAgrohemija, modScrAnaliza, modMarza, modAmbalaza)
 E SYNC, PWA, GAS, IZVOZI (modGoogleSyncOrchestrator, modMasterSync, modStammdatenSync, modStanicaLock, gas/Code.gs, src/js)
 F ADMIN, MAKROI, INTEGRITET, HEALTH, SETUP (modAdmin, modPodesavanja, modIntegritet, modProductionHealthCheck, modSetup, modAutoHladnjaca, javni Sub-ovi bez argumenata = Alt+F8)

## Kako naći sposobnosti
- Ekrani: registar u modUiScreens (c.Add "KLJUC|modScrX|..."); u svakom modScr* procedure Scr_Liste, Scr_Radnje, Scr_Event, Scr_Save, Scr_Rows.
- Paneli: modUiPanel.PanelRedovi. Tajmeri i kasno vezani pozivi: Application.OnTime / Application.Run / OnAction stringovi.
- Postojeći opisi (koristi, ali proveri u kodu): docs/UI_MIGRACIJA_KATALOG.md, docs/AgriX_Functional_Map_v142.md.
- Mrtav kod (bez pozivaoca) NIJE sposobnost — navedi ga posebno, kratko.

## Za svaku sposobnost jedan red
| ID | Sposobnost (glagol + predmet) | Ulaz (ekran/lista/radnja ili makro/sync, fajl:linija) | Izlaz (ekran, PDF, izvoz, sheet, poruka) | Pravilo/ishod koji operater očekuje | Podaci (tabele.kolone koje čita/piše) | Legacy implementacija (procedure, fajl:linija) | Zavisi od starog modela? | Testovi koji tvrde ishod (ime + tekst poruke) |

- ID: <OBLAST>-NNN.
- „Zavisi od starog modela?“: da/ne, i ako da — koje: Otkup.OtpremnicaID / BrojZbirne / VozacID / BrojOtpremnice; linijska polja na zaglavlju otkupa (Kolicina, Cena, Klasa, KolAmbalaze, BrutoKg); po-klasna polja na zaglavlju otpremnice (Kolicina, Klasa, KolAmbalaze, BrutoKg, Cena); Otpremnica.BrojZbirne ili poslovni broj kao veza; GeneracijaID; SaveOtpremnica*/SaveZbirnaMulti_TX i drugi pisci „red po klasi“; Split(ID, " + ").
- Pravilo/ishod pišeš jezikom operatera („storno otpremnice oslobađa njene otkupne blokove“), ne imenom procedure.

## Na kraju
1. Pokrivenost: spisak svih ulaznih tačaka oblasti koje si našao (radnje, liste, dugmad, makroi) i za svaku ID sposobnosti ili „mrtvo“/„tehničko“.
2. Sposobnosti koje postoje samo u PWA/GAS, a VBA ih ne vidi (za oblast E obavezno).
3. NEPROVERENO sa razlogom.
Izlaz: samo markdown (tabela + tri liste), bez uvoda.
