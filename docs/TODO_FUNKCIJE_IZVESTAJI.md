# AgriX / OtkupApp — Pregled funkcija i izveštaja

*Radni dokument — jul 2026.*

---

## A. Knjigovodstveno-poreski sloj — najveća rupa, direktno prodaje

| # | Stavka | Opis | Napomena |
|---|--------|------|----------|
| A1 | **Specifikacija po priznanicama** | Broj priznanice, kooperant (BPG/JMBG), datum otkupa, osnovica, PDV nadoknada 8%, datum i način isplate vrednosti i nadoknade posebno | Dizajnirati tako da služi kao formalna PDV evidencija |
| A2 | **Zbirni nalog za knjiženje** | Izveden iz A1; konta Soll/Haben (241/243/430/130...) | Razdvojiti „obračunato" (otkup) od „isplaćeno" (pravo na odbitak nadoknade — samo virmanski isplaćene stavke) — veza sa postojećim P4 (Kontni plan + Buchhalter-Export) |
| A3 | **POPDV-ready izveštaj** | Polja koja knjigovođa direktno prepisuje u POPDV | Automatizovan preko BankaImport razvođenja |
| A4 | **RPG validacija / porez i doprinosi za ne-RPG lica** | Upozorenje na formi ako lice nije nosilac/član RPG; ako klijenti otkupljuju od takvih lica → obračun poreza i doprinosa (PPP-PD) | **Prvo proveriti sa klijentima da li se slučaj uopšte javlja** |
| A5 | **IOS po kooperantu** | Izvod otvorenih stavki za kraj godine | Podaci već postoje u SaldoOM |
| A6 | **Blagajnički dnevnik** | Formalan dnevni izveštaj za keš isplate | Kategorija KesOtkupacKoop već postoji |
| A7 | **BankaImport — razvođenje izvoda** | Automatsko vezivanje isplata sa izvoda za priznanice/kooperante | Već u backlogu (P2); postaje temelj za A2 i A3 |

## B. Operativa — feature-parity, traži se u prodaji

| # | Stavka | Opis | Napomena |
|---|--------|------|----------|
| B1 | **Ambalaža** | Zaduženje/razduženje gajbica po kooperantu, revers, trenutno stanje, lager ambalaže | Konkurencija ovo ima kao standard — prvo pitanje klijenta br. 3 |
| B2 | **Usluga skladištenja za treća lica** | Ležarina, ulaz/izlaz tuđe robe po komorama, obračun i fakturisanje usluge | Otvara novi segment kupaca (hladnjače-uslužne) |
| B3 | **Kompenzacija repromaterijala** | Đubrivo/sadnice/zaštita dato kooperantu → prebijanje sa otkupom | Proveriti da li Avans logika već pokriva u celosti |

## C. Izveštaji za gazdu — retencija

| # | Stavka | Opis | Napomena |
|---|--------|------|----------|
| C1 | **Marža po kulturi / kooperantu / periodu** | Otkupna vs prodajna strana | Osnov za gazdine odluke |
| C2 | **Cash-flow projekcija** | Dospele obaveze prema kooperantima po nedeljama | Kritično u sezoni |
| C3 | **RZS statistički izveštaj o otkupu** | Mesečni obrazac koji otkupljivači predaju | Proveriti sa klijentima koji tačno obrazac predaju |

---

## D. Nice-to-have — 10 potencijalnih funkcija i izveštaja

| # | Stavka | Opis | Zašto |
|---|--------|------|-------|
| D1 | **Paket virmana za banku** | Automatsko generisanje naloga za prenos (isplate kooperantima) u formatu za Halcom / OfficeBanking / FX Client | Skida sate ručnog kucanja virmana; prirodan par sa BankaImport (izvod ulazi → virmani izlaze) |
| D2 | **Potvrda otkupa Viberom/SMS-om** | Kooperant posle vaganja odmah dobija poruku: količina, klasa, cena, novi saldo | Poverenje kooperanata = lojalnost hladnjači; koristi već planiranu Viber infrastrukturu (modMeteo, P5) |
| D3 | **Kooperant-portal (PWA pogled)** | Kooperant vidi svoje priznanice (PDF), isplate i saldo | Smanjuje pozive „koliko mi duguješ"; diferencijator na tržištu |
| D4 | **Godišnja potvrda o otkupu za kooperanta** | Zbirni dokument po kooperantu za godinu | Kooperantu treba za subvencije/kredite — traže je svake godine |
| D5 | **Rang lista kooperanata** | Po količini, vrednosti, % škarta, redovnosti | Osnova za odluke o avansima i bonusima pred sezonu |
| D6 | **Istorija cena + tržišno poređenje** | Kretanje otkupnih cena po kulturi kroz sezone; poređenje sa STIPS/berzanskim cenama | Gazdi argument u pregovorima; lep grafički izveštaj |
| D7 | **Zauzetost komora / lager mapa** | Vizuelni prikaz popunjenosti hladnjače po komorama i kulturama | Operativno planiranje; osnova i za B2 |
| D8 | **Etikete sa barkodom po gajbi/paleti** | Štampa nalepnica pri prijemu; skeniranje pri izlazu | Diže Sledljivost v2.0 na nivo skeniranja — jak GGAP argument |
| D9 | **Ugovori o otkupu** | Evidencija + generisanje ugovora sa kooperantom (mail-merge iz podataka) | Imaš mail-merge iskustvo; hladnjače sa ugovorenom proizvodnjom ovo traže |
| D10 | **Dnevni KPI dashboard** | Otkup danas (kg/RSD), isplate danas, dospele obaveze, zauzetost — jedan ekran na frmMain ili PWA | Gazda otvara jedan pogled ujutru umesto pet izveštaja |

---

## Predlog redosleda

1. **A-blok** (A1→A2→A3 kao jedna celina, oslonjena na A7/BankaImport) — pretvara program za vagu u sistem koji knjigovođa brani umesto da mu smeta.
2. **B1 Ambalaža** — feature-parity pitanje koje će iskrsnuti u prvoj sledećoj prodaji.
3. **D1 Virmani** — mali obim, veliki wow-efekat, zaokružuje novčani tok u oba smera.
4. Ostalo po prilici i zahtevima klijenata.

*Napomena: format A1/A2 pre kodiranja validirati sa knjigovođom jednog od postojećih klijenata.*
