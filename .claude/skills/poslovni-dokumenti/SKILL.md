---
name: poslovni-dokumenti
description: Pravila za rad sa AgriX poslovnom dokumentacijom u docs/ — Master Plan, decision log, cenovnik, ponuda, ugovor, finansijski model, sales i legal materijali. Koristi kad se dodaje ili menja bilo koji dokument u docs/Master Plan/, docs/Sales/, docs/Legal/, docs/Finance/, docs/Product/, kad se unosi nova odluka ili menja postojeća, kad se menja cena, kad se uploaduje PDF/DOCX/XLSX iz chata u repo, ili kad treba proveriti da li se dokumenti međusobno slažu. Ne odnosi se na VBA/PWA kod — za to važi CLAUDE.md.
---

# AgriX poslovni dokumenti — operativna pravila

Ovaj skill pokriva `docs/`. Za kod važi `CLAUDE.md`. Isti default stav:
`reuse > new`, `verify > conclude`, `inspect before propose`.

---

## 1. Gde je istina

| Šta | Merodavan fajl |
|---|---|
| Numerisane odluke | `docs/Master Plan/09_QA_DECISION_LOG.md` |
| Tematski indeks odluka | `docs/Master Plan/09B_ODLUKE_PO_OBLASTIMA.md` |
| STR/GOV odluke i njihovi razlozi | `docs/Master Plan/DECISION_LOG.md` |
| Istorija izmena Master Plana | `docs/Master Plan/CHANGELOG.md` |
| Cene | `docs/Sales/AgriX_Cenovnik_2027.pdf` |
| Tekst ugovora | `docs/Legal/AgriX_Ugovor_o_licenciranju.md` |
| Cilj rasta | `docs/Finance/AgriX_Finansijski_model.xlsx`, list `Pretpostavke` red 42 |

**Pravilo prvenstva:** kad se dokument i decision log razlikuju, važi decision log.
Kad se dve odluke razlikuju, važi kasnija — i to se izričito označi.

---

## 2. Pre svake izmene

1. Pročitaj `09_QA_DECISION_LOG.md` i nađi odluke koje temu dodiruju. Ne piši
   poslovnu tvrdnju bez ID-ja odluke iza nje.
2. Ako tvrdnja nema odluku — reci da nema. Ne izmišljaj ID i ne pretpostavljaj
   sadržaj odluke koja nije u repou.
3. Grep-uj pre pisanja: ista tvrdnja često živi na 3–4 mesta (poglavlje, CSV
   matrica, README, PDF). Menja se svuda ili nigde.

---

## 3. Numeracija odluka

- Zauzeto: **1–321**, **323–378**, **401–408**. Slobodno: 322, 379–400 se **ne koriste**.
- Uz brojeve postoje serije: A, BC, C, D, I, IP, L, LEG, M, MKT, ML, ON, P, PRT, Q, S.
- Nove numerisane odluke idu na kraj `09_QA_DECISION_LOG.md` kao nov odeljak sa
  datumom i obuhvatom, pa se opseg dopiše u odeljak „Napomena o kontinuitetu numeracije“.
- **Odluke se ne brišu.** Kad prestanu da važe: `Superseded`, datum, i veza ka novoj.
- Kad odluka menja raniju, u tekstu nove piše šta se menja, a u staroj se doda oznaka.

Provera pokrivenosti posle unosa:

```bash
python3 - <<'EOF'
import re
t=open('docs/Master Plan/09_QA_DECISION_LOG.md',encoding='utf-8').read()
n=set(int(x) for x in re.findall(r'^(\d{1,3})\. \*\*', t, re.M))
miss=[i for i in list(range(1,322))+list(range(323,379))+list(range(401,409)) if i not in n]
print('nedostaje:', miss or 'nista')
EOF
```

---

## 4. Sinhronizacija cena — četiri mesta

Svaka cena mora biti identična u:

1. `docs/Sales/AgriX_Cenovnik_2027.pdf`
2. `docs/Sales/AgriX_Sablon_ponude.xlsx`, list `Cenovnik`
3. `docs/Legal/AgriX_Ugovor_o_licenciranju.md`, Prilog 1
4. `docs/Finance/AgriX_Finansijski_model.xlsx`, list `Pretpostavke`, odeljak A

Menjaš cenu → menjaš sva četiri + odluku u logu. Ako neko od četiri ne možeš da
izmeniš (npr. PDF je dizajniran van repoa), **to se izričito kaže korisniku**,
ne prećuti.

Čitanje binarnih fajlova:

```bash
pdftotext -layout docs/Sales/AgriX_Cenovnik_2027.pdf -   # PDF
python3 -c "import openpyxl,sys; wb=openpyxl.load_workbook(sys.argv[1]); [print(ws.title,[c.value for r in ws.iter_rows() for c in r]) for ws in wb]" fajl.xlsx
pandoc -f docx -t gfm --wrap=none fajl.docx              # DOCX
```

Ako `pdftotext`/`pandoc` nedostaju: `apt-get update && apt-get install -y poppler-utils pandoc`.

---

## 5. Klase dokumenata

| Klasa | Primeri | Pravilo |
|---|---|---|
| **Content-primary** | decision log, Master Plan poglavlja, ugovor | Markdown u repou je izvor; PDF/DOCX se generiše |
| **Design-primary** | cenovnik, materijal za prvi kontakt | dizajn ostaje van repoa, ali sadržaj mora imati parnjak u repou |
| **Model** | finansijski model, šablon ponude | XLSX je izvor; formule se ne diraju bez razloga |

**Ugovor** ima poseban tok — `.md` je izvor, `.docx` je deliverable:

```bash
tools/ugovor.sh check    # da li .docx i .md i dalje govore isto
tools/ugovor.sh build    # .md -> docs/Legal/build/*.docx
```

Kad pravnik vrati izmenjen `.docx`: prvo `check`, pa izmene ručno preneti u `.md`.

---

## 6. Uploadovani dokumenti iz chata

Redosled: **pročitaj sadržaj → tek onda odluči folder.** Ime fajla nije dovoljno.

| Sadržaj | Folder |
|---|---|
| definicija proizvoda, specifikacije, procesne mape | `docs/Product/` |
| ugovori, privacy, mapa tokova podataka, regulatorno | `docs/Legal/` |
| cenovnik, ponude, outreach, playbook | `docs/Sales/` |
| modeli, budžeti, unit economics | `docs/Finance/` |
| strategija, odluke, tržište, portfolio | `docs/Master Plan/` |

Posle smeštanja **uvek** dopuni README tog foldera: tabela dokumenata sa verzijom,
datumom i kratkim sadržajem, plus uslovi korišćenja — šta je nacrt, šta čeka
potvrdu, šta se ne sme slati klijentu.

Hash prefiks iz upload putanje se skida iz imena fajla.

---

## 7. Provera pre commit-a

1. `grep` za svaku promenjenu tvrdnju kroz ceo `docs/` — nema zaostalih starih vrednosti.
2. Ako je diran CSV: broj kolona po redu je konstantan.
   ```bash
   python3 -c "import csv,sys; r=list(csv.reader(open(sys.argv[1],encoding='utf-8'))); print('bad:',[i+1 for i,x in enumerate(r) if x and len(x)!=len(r[0])])" 'docs/Master Plan/07A_PRODUCT_STATUS_MATRIX.csv'
   ```
3. Ako je diran `.docx`: zip integritet.
   ```bash
   python3 -c "import zipfile,sys; print('bad:',zipfile.ZipFile(sys.argv[1]).testzip())" fajl.docx
   ```
4. Ako je dirana cena ili ugovor: `tools/ugovor.sh check`.
5. `CHANGELOG.md` dobija unos: Added / Changed / Superseded / Open.
6. Otvorene stavke idu u tabelu u `09_QA_DECISION_LOG.md`, ne u prozu.

---

## 8. Naučene greške — ne ponavljati

- **Ne izvoditi broj iz konteksta.** Raspon „12–15“ u jednom dokumentu bio je broj
  *novih* klijenata, a prepisan je kao *ukupan*. Uvek proveri jedinicu: novi ili ukupno,
  po stanici ili po pravnom licu, maloprodajna ili kanalska cena.
- **Ne proglašavati nešto nedefinisanim pre nego što se pročitaju svi dokumenti.**
  Zapisano je da Savetniku „packaging i cena nisu zaključani“, a cena je već bila
  objavljena u cenovniku.
- **Status proizvoda nije mišljenje.** `Pilot only` / `Standard offer` / `Not for sale`
  menja se samo odlukom sa ID-jem.
- **Tri stuba su Enterprise, Gazdinstvo i Savetnik** (odluke 269, 401). GGAP je
  modul Enterprise-a (odluka 402), ne stub — STR-012 je povučena.
- **PDF snimak i `.md` se razilaze.** Kad `.md` ima parnjak u PDF-u, u `.md` ide
  napomena šta je od snimka zastarelo.

---

## 9. Git

Grana, commit, push kao u `CLAUDE.md`. PR ka `main` tek na eksplicitan zahtev, uz
`git merge-tree` proveru. Docs-only izmene **ne traže** release ciklus ni
`ImportAllVBA` — to važi samo kad je diran `src-vba/` ili `src/`.
