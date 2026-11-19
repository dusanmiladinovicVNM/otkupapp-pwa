# Plan rada — AgriX VBA (od 19.11.2026)

> **Status: PREDLOG.** Redosled važi kad operater odgovori na tri odluke iz §5.
> Ovaj dokument drži samo redosled i uslove završetka. Detalji po rezu ostaju u
> svojim izvorima: `docs/STANJE_REFAKTORA.md` (refaktor), `docs/DOMEN/SLEDLJIVOST.md`
> (S9), `docs/DOMEN/MIGRACIJA_LEGACY.md` (C002), `docs/KNOWN_ISSUES.md` (AUD/KI),
> `docs/DOMEN/STORNO_U_LANCU.md` §12 (izmena izdate zbirne i prijemnice).

**Sidro:** `main` `1a2293d5` (posle #495), 19.11.2026.

**Kako je nastao:** čitanjem repoa, dokumenata i otvorenih PR-ova u web sesiji.
Ništa nije pušteno u Excelu. Klon je bio plitak (istorija od 22.10.2026), pa su
brojevi o aktivnosti mereni samo u tom prozoru.

---

## 1) Zašto se redosled menja

Dosadašnji red je bio: P1 → S9 → R1 → R2 → R3, uz talase brzine i dalje faze
optimizacije verifikacije. Ovaj plan ispred svega stavlja **izdanje i prelazak
C002 na novi model**. Razlozi:

1. **Izdanje kasni.** Poslednje je `vba-v2.92.0` od 03.09.2026. `v2.93.0` je u
   `docs/RELEASE_NOTES.md` i dalje „u pripremi", a nosi sve od #333 do #495.
2. **C002 juri pokretnu metu.** Sveska je poslednji put migrirana 03.11.2026 nad
   `a668c5e6`. Posle toga su S8, S9 i cenovnik menjali šemu i pisce, pa se
   migracija ponavlja. Provera migrirane sveske u pravom Excelu je
   NEVERIFIKOVANA (`MIGRACIJA_LEGACY.md`, „Otvoreno").
3. **Rupa u mapi sposobnosti.** `src-vba/modScrSledljivost.bas` ne postoji.
   Ekran SLEDLJIVOST i GlobalGAP dokument zbirne (C-031 do C-044) vraća tek
   S9-7 (`SLED-ODL-04`). Mapa se čuva apsolutno, pa je S9-7 uslov za izdanje.
4. **Težište je otišlo sa proizvoda na alat.** Od 22.10.2026: 288 commit-a, od
   toga 48 dira `src-vba`, 102 dira `tools`, 125 su samo dokumentacija.
5. **Registri ne govore istinu.** Primeri: AUD-030 se vodi kao aktivan P0, a
   ispravka je u kodu (`modSEFClient.bas:483`); AUD-034 je rešen
   (`ThisWorkbook.doccls:30`); AUD-005 opisuje obrisan `AutoChainHladnjaca`.
   `PLAN_SANACIJE.md` (M0–M12) nema status. Zato se danas ne može reći koliko
   je P0 otvoreno, a Master Plan (faza A0) traži „nema nepotvrđenog P0".
6. **Preskočen test se ne vidi.** U `src-vba` nema nijednog `SKIP=`, pa suite u
   kojoj sve preskoči piše `FAIL=0` (nalaz N1 iz PR #410).
7. **PWA i sync su bez pažnje.** U merenom prozoru je jedan commit u `src/` i
   `gas/`. Tok teren → centrala nad novim modelom nije izmeren uživo.

---

## 2) Redosled

| # | Posao | Kraj je kad |
|---|---|---|
| 1 | **S9 do kraja (rezovi 3–7)** — poslednji posao koji menja šemu | mapa sposobnosti bez rupe; `popis_citalaca --check` na pragu 0 za grupe koje S9 gasi |
| 2 | **Istina o registrima** (bez Excela) | svaki AUD i KI je potvrđen nad `main`-om, zatvoren ili obrisan; jedna lista stvarno otvorenih P0/P1; `PLAN_SANACIJE.md` i `ROADMAP.md` §1 nose status |
| 3 | **`SKIP` vidljiv u zaglavlju rezultata i u kapiji** (process PR, dokaz u oba smera) | suite sa preskočenim testom nije zelena |
| 4 | **Zamrzavanje i izdanje** | FULL zelen, pun `dokaz.py`, ručni Compile, release notes, potvrđeni P0 iz tačke 2 zatvoreni |
| 5 | **C002 na novi model** | migracija nad tagom izdanja; operater prošao proveru u Excelu; `KI-024` (deo 2) i `KI-065` zatvoreni |
| 6 | **PWA/sync nad novim modelom, uživo** | tok stanica → centrala prolazi nad test Google okruženjem; AUD-019 (PIN i JMBG) i KI-001 (`saveParcelPolygon`) imaju odluku |
| 7 | **Izmena izdate zbirne i prijemnice** — P1, R1, R2; R3 ostaje HOLD za dizajn | po uslovima iz `STORNO_U_LANCU.md` §12.3 |
| 8 | **Brzina** — dalji talasi tek uz tajmer | vreme izmereno pre i posle |

Tačke 2 i 3 ne zavise od tačke 1 i mogu ići paralelno sa njom. Tačka 4 čeka
sve tri.

Pre svakog reza koji dira šemu, pisca, identitet ili storno važi kapija
`pre-flight`, kao i do sada.

---

## 3) Zamrznuto do izdanja

- Faze 5–7 iz `docs/engineering/PLAN_OPTIMIZACIJE_VERIFIKACIJE.md`. Cilj je
  postignut prijemom G1. U `tools/` ulazi samo ono što blokira rez, i tačka 3.
- Novi talasi brzine posle #497.
- Nove sposobnosti i „bele zone" (`DOMEN/FUNKCIONALNE_BELE_ZONE_I_ROADMAP.md`).

Od zamrzavanja do završetka tačke 5 `main` prima samo ispravke.

---

## 4) Otvoreni PR-ovi (stanje 19.11.2026)

| PR | Predlog |
|---|---|
| #497 brzina | spojiti pre zamrzavanja, kad prođu njegove četiri ručne provere |
| #496 process pravilo | spojiti |
| #477 plan LZP | spojiti kao plan; izvršenje je tačka 7 |
| #457 AMB po OM | prenumerisati test (350 se sudara sa `main`-om) i izmeriti, ili zatvoriti |
| #410 plan test platforme | zatvoriti; iz njega živi samo tačka 3 |
| #444 arhiva migracionih alata | ostaje draft |
| #7, #42 | zatvoriti kao zastarele (#42 nije upoređen sa `modLicense` na `main`-u) |

---

## 5) Odluke koje čekaju operatera

1. **S9-7 pre izdanja?** Alternativa je izdanje koje glasno kaže da sledljivosti
   nema. Predlog: S9-7 ide pre, zbog GlobalGAP dokumenta.
2. **Kako postojeći klijenti osim C002 prelaze na novi model?** Ako idu kroz
   self-update, AUD-035 (komponenta nestaje pri neuspelom uvozu forme) postaje
   P0 za izdanje. U kodu nije proveren.
3. **Trenutak zamrzavanja.** Predlog: vezati ga za kraj S9, ne za kalendar.

---

## 6) Organizacija rada

Usko grlo je jedna Windows mašina sa Excelom i pažnja operatera, ne pisanje
koda. Pravila koja iz toga slede:

1. **Rez koji menja ponašanje radi se u lokalnoj sesiji na Windows mašini**, koja
   sama pušta `run_vba` i `dokaz.py`. Web sesija dobija samo posao bez Excela:
   registri, review, pre-flight, PWA i GAS, migracioni alat.
2. **Jedna serijska traka za šemu, pisce i `modTest`**; pored nje najviše dve
   paralelne za posao koji se sa njom ne sudara. Svaka traka ima svoj worktree.
3. **Broj testa se dodeljuje pri spajanju, ne pri pisanju.**
4. **Najviše dva otvorena feature PR-a.** PR stariji od nedelju dana se zatvara
   ili odmah preslaže.
5. **FULL i pun `dokaz.py` idu svake noći** nad `main`-om (Task Scheduler);
   rezultat se čita ujutru.
6. **Ručne provere operatera idu u jedan dnevni blok**: Compile, provera ekrana,
   čitanje diff-a, merge.
7. **Nezavisni review samo za rizične rezove**: šema, storno, novac, identitet.
8. **Jedan zapis po rezu, u istom commit-u sa kodom.** `STANJE_REFAKTORA.md` se
   svodi na trenutno stanje, a istorija stavki ide u arhivu (zaseban rez).
9. **Izdanje ima ritam** (predlog: svake dve nedelje, sa oznakom).

---

## 7) Šta ovaj plan ne tvrdi

- Da je spisak otvorenih P0 poznat — to je ishod tačke 2.
- Da je obim S9-3 do S9-7 procenjen; rezovi imaju pre-flight, ne procenu trajanja.
- Da PWA tok nad novim modelom ne radi — samo da nije izmeren uživo.
- Da brojevi o aktivnosti važe pre 22.10.2026.
