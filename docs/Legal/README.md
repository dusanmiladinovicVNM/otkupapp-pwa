# Legal

Ugovorni modeli, privacy, data processing, SLA granice, odgovornost za hardver, softverska ograničenja i regulatorne teme.

Poverljivi ugovori i lični podaci ostaju van javnog repozitorijuma.

## Dokumenti

| Dokument | Verzija / datum | Sadržaj |
|---|---|---|
| `AgriX_Mapa_tokova_podataka.pdf` | v1 radni nacrt · 26.07.2026. | Učesnici, kategorije podataka K1–K11, tokovi T01–T19, podobrađivači, predlog raspodele uloga rukovalac/obrađivač i pitanja za pravnika. |
| `AgriX_Ugovor_o_licenciranju.md` | nacrt · 27.07.2026. | **Izvor istine za tekst ugovora.** Isti tekst kao `.docx`, plus mapa članova na ID-jeve odluka i pravilo sinhronizacije cena. |
| `AgriX_Ugovor_o_licenciranju.docx` | nacrt · 27.07.2026. | Deliverable za pravnika i klijenta: 15 članova + Prilog 1 (obim i cene), Prilog 2 (podrška), Prilog 3 (obrada podataka o ličnosti). |

**Tok rada sa ugovorom** — `.md` je izvor, `.docx` je izlaz:

```bash
tools/ugovor.sh check    # da li .docx i .md i dalje govore isto
tools/ugovor.sh build    # .md -> docs/Legal/build/AgriX_Ugovor_o_licenciranju.docx
```

Menjaš uslove → menjaš `.md`, pa `build`. Pravnik vrati izmenjen `.docx` → prvo `check`, pa izmene preneti u `.md`. `build` namerno **ne prepisuje** kurirani `.docx`, nego piše u `docs/Legal/build/`. Zavisnost je `pandoc`.

**Status ugovora — nacrt, nije za potpisivanje u ovom obliku:**

- pravni pregled još nije obavljen; nacrt piše osnivač, pravnik radi pregled gotovog teksta (odluka 376);
- **Prilog 3 nije dovršen** — uloge rukovaoca i obrađivača utvrđuju se tek posle mape tokova i pravne analize (LEG1). U dokumentu stoji izričita napomena da se u ovom obliku ne potpisuje;
- **popunjeno 27.07.2026.:** vremenski prozor za rok reakcije od jednog sata (odluka 359, Prilog 2) — tokom sezone svakog dana 08.00–20.00, van sezone radnim danima 08.00–16.00; mesto nadležnog suda (član 15) — **Niš**, jer je sedište AgriX-a Merošina;
- **i dalje nepopunjeno:** spisak drugih obrađivača i lokacija obrade (Prilog 3, odluka 373);
- cene u Prilogu 1 preuzete su iz odluka 349–357 i **409–422** i moraju ostati identične sa `docs/Sales/AgriX_Cenovnik_2027.pdf`; hardverska podrška (odluka 357) potvrđena je 27.07.2026.

Popunjeni i potpisani ugovori sa podacima klijenta ne commit-uju se u repozitorijum — ovde ostaje samo prazan nacrt.

Mapa tokova je ulaz za razrešenje odluke **LEG1** (formalne uloge u zaštiti podataka o ličnosti) i osnov za Prilog 3 ugovora. Predložene uloge u koloni „Predlog uloge“ **nisu pravno potvrđene** i ne smeju se koristiti u ugovorima pre pravne provere. Od nje zavise i **LEG5** (rokovi obaveštavanja) i politike privatnosti za Gazdinstvo i Savetnik.
