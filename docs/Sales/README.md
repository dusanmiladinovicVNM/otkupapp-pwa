# Sales

Prodajni playbook, discovery, demo, kvalifikacija, ponude, prigovori, pipeline pravila i win/loss proces.

Poverljive ponude, kontakt podaci i ugovorni detalji ne čuvaju se u javnom repozitorijumu.

## Dokumenti

| Dokument | Verzija / datum | Sadržaj |
|---|---|---|
| `AgriX_Cenovnik_2027.html` | 27.07.2026. | **Izvor istine za cenovnik.** Cene se menjaju isključivo ovde, u `data-eur` atributima. Koristi AgriX brand tokene iz `src/styles/base.css`, fontove iz `vendor/fonts/` i logo iz `img/` — relativnim putanjama, pa fajl mora ostati u `docs/Sales/`. |
| `AgriX_Cenovnik_2027.pdf` | važi od sezone 2027 · 27.07.2026. | Generisani cenovnik za klijenta, 9 strana: paketi Desktop/Mobile i all-in varijante sa izričitim sastavom, moduli sa obračunskom jedinicom, stanice i dodatna instanca, Gazdinstvo, **dve tarife Savetnika** (standalone i Enterprise), **dve satnice** (razvojna 50 €/h i implementaciona 30 €/h), primeri obračuna, šta je uključeno u pretplatu. |
| `AgriX_Materijal_za_prvi_kontakt.html` | v2 · 28.07.2026. | **Izvor istine za materijal.** Svaka tvrdnja o ceni, popustu, probi ili roku nosi ID odluke iza sebe. |
| `AgriX_Materijal_za_prvi_kontakt.pdf` | v2 · 28.07.2026. | Generisani interni prodajni dokument, 8 strana: prodajni prozori po kulturama, tri tira i tri poruke, skripta poziva, email šabloni, prigovori i odgovori, šta se nikada ne obećava, evidencija posle poziva. **Ne šalje se klijentu.** |
| `_brand.css` | — | Zajednički brand tokeni i fontovi za oba dokumenta. Izvor: `src/styles/base.css`. |
| `AgriX_Sablon_ponude.xlsx` | v1 · 26.07.2026. | Radni šablon ponude sa listom `Cenovnik` kao jedinim mestom za cene. Ponuda povlači vrednosti iz cenovnika; cene se ne kucaju u ponudu. |

Napomene:

- **Cenovnik se ne menja u PDF-u** — menja se `AgriX_Cenovnik_2027.html` pa se PDF regeneriše:

  ```bash
  tools/cenovnik.sh build|check     # cenovnik
  tools/materijal.sh build|check    # materijal za prvi kontakt
  ```

  Oba koriste `tools/render-pdf.sh`, koji normalizuje vremenske pečate u PDF-u — dva build-a istog izvora daju **identičan** fajl, pa nema lažnih git diff-ova. Zavisnost je Chromium/Chrome; `CHROME_BIN` nadjačava automatsko pronalaženje.

  **Šta `cenovnik.sh check` stvarno proverava:** sve iznose iz `data-cena` atributa (ne samo pakete), poklapanje atributa sa prikazanim tekstom, strukturu cena iz odluka 414 i 415 (Mobile dodatak 1.000 €, all-in doplata 700 €), odsustvo mrtve tačke u kojoj à la carte košta koliko all-in, izvedene zbirove u primerima protiv njihovih komponenti, i poklapanje sa šablonom ponude, Prilogom 1 ugovora i finansijskim modelom.

  **Šta `materijal.sh check` proverava:** da nijedna formulacija ne obećava popust (odluka 418), da svaki navedeni ID odluke postoji u decision logu i nije obrisan, i da dokument nosi oznaku da je interni.

- **Cene moraju biti identične na četiri mesta:** `AgriX_Cenovnik_2027.html`, list `Cenovnik` u šablonu ponude, Prilog 1 ugovora (`docs/Legal/AgriX_Ugovor_o_licenciranju.md`) i finansijski model. `tools/cenovnik.sh check` to proverava programski;
- cene se menjaju samo kada se promeni odluka o ceni (izvor: odluke 339, 341, 349–358, 409–422);
- šablon ponude je prazan obrazac — popunjene ponude sa podacima klijenta se ne commit-uju;
- hardverska podrška (odluka 357) i cena po gazdinstvu kod Savetnika (odluka 341) potvrđene su 27.07.2026.;
- Dispatch se nudi samo uz Mobile paket (odluka 293). **GGAP se sme prikazati samo uz vidljivu oznaku „na upit, uz potvrdu obima — nije deo standardne ponude“** (odluka 417); ostaje van redovne komercijalne ponude do validacije (odluka 405);
- **Savetnik nosi oznaku „u pripremi“ i ne kotira se kao redovna stavka** (odluka 423). Ima dve objavljene tarife — standalone 150 €/15 € i Enterprise 100 €/10 € (odluke 419, 420) — ali se ne ugovara dok proizvod ne bude stabilan (odluka 217). U cenovniku ne sme nositi zlatni okvir ni drugu vizuelnu oznaku preporuke, i ne uvrštava se automatski u ponudu;
- **nema pregovaračkih ni individualnih popusta** (odluka 418). Jedina cenovna razlika unutar istog obima je −50 % na drugu i svaku narednu instancu (odluka 413); objavljene razlike iz cenovnika nisu popusti;
- marža na hardver (~100 €/stanici, odluke 356 i 407) je **interni podatak i ne prikazuje se klijentu**.
