# Postmortem: devet PR-ova zakrpe za jednu grešku u modelu, septembar 2026.

Pravilo koje iz ovoga sledi živi u `.claude/rules/model-pre-zakrpe.md`; registar u
`docs/DOMEN/KOMPENZACIJE.md`. Ovde je **zašto** — arheologija koja ne mora da se
učitava u svakoj sesiji.

Čita se kad neko pita „zašto se traži poređenje zakrpa/model" ili kad se isti
obrazac ponovi.

---

## 1) Šta se dogodilo

Dokument u lancu (otkup, otpremnica, zbirna, prijemnica) modelovan je kao
`jedan fizički red = jedan dokument`. To nije tačno: dvoklasni dokument su dva
reda sa dva ID-a. Logički dokument postoji u domenu, ali u šemi nema gde da živi,
pa se **rekonstruiše pri svakom čitanju**.

Za tu rekonstrukciju je izgrađeno: `GeneracijaID` kao surogat identiteta,
`ZbirnaIdentResolve` sa deset izlaznih polja, pet imenovanih kapija
(`ZBR-ACTIVE-NUMBER-01`, `ZBR-PARENT-01`, `ZBR-MUT-01`, `ZBR-CHILD-01`,
`ZBR-NORM-02`), pečaćenje generacije u tri pisca, poseban read-model, dedup u
pickeru, i konkateniran string `"OTK-1 + OTK-2"` koji se parsira na devet mesta.

Mera, iz `git log` i `grep` nad `src-vba/`:

| | |
|---|---|
| PR-ova | 9 (#291, #292, #293, #295, #296, #298, #299, #300, #301) |
| commita | 53 kroz 3 dana (07.09. — 14, 08.09. — 25, 09.09. — 14) |
| od toga `fix(...)` | 22 |
| `Generacija*` u `src-vba/` | 355 pojava, 25 modula |
| `ZbirnaIdent*` / `VlasniciPoBroju` / `RequireJedanVlasnik*` | 74 / 43 / 20 |
| `Split(x, " + ")` | 9 call site-ova |

Dana **09.09.** napisan je `docs/REFAKTOR_DOKUMENT_HEADER_STAVKE.md`, čija prva
rečenica glasi da je sve to „kompenzacija za nepostojeće zaglavlje", i
`docs/DOMEN/DOCUMENT_HEADER_LINES.md` §10, koji kaže da sa headerom
`GeneracijaID`, `ZbirnaIdent`, resolveri i vlasničke kapije **nemaju posao**.

> Ispravno rešenje nije bilo teško naći. Nikad nije bilo traženo.

**Opseg merenja:** brojevi gore pokrivaju `BrojZbirne` krak, koji je u git prozoru
ovog repoa cele 3 dana. Obrazac je stariji — `KNOWN_ISSUES.md` TL-008 opisuje istu
bolest za `BrDok` i poziv na broj, i ta dva kraka su **i dalje otvorena**.

---

## 2) Pet uzroka

### 2.1 Podrazumevani stav je nagrađivao zakrpu

`CLAUDE.md` je nosio `reuse > new` · `extend > duplicate` ·
`minimal change over idealized redesign`. Naspram pogrešnog modela, „minimalna
izmena" je **uvek** sledeća zakrpa: izmena modela nije minimalna ni u jednom
pojedinačnom tiketu — minimalna je tek u **zbiru** svih tiketa koje isti
nedostatak traži. Nijedno pravilo taj zbir nije računalo.

**Pravilo koje je ostalo:** „minimalno" se meri preko svih rata koje isti
nedostatak modela traži (`CLAUDE.md` §0, `model-pre-zakrpe.md` §8).

### 2.2 Sve kapije su merile ispravnost UNUTAR modela

`pre-flight` traži verdikt po osama `DOMAIN` / `IDENTITY` / `CARDINALITY` /
`INVARIANTS` naspram `docs/DOMEN/README.md`. `GeneracijaID` prolazi svaku kao
`PROVEN` — ima ugovor, ID, testove i vlasnika. Kapija je kompenzaciju **overila**
kao dobar inženjering, jer je to i bila, unutar zadatog modela.

Nijedna osa ne pita: **da li bi ova osa bila prazna da je šema drugačija.**

**Pravilo koje je ostalo:** detektor kompenzacije (`model-pre-zakrpe.md` §2) meri
oblik rešenja, ne njegovu ispravnost.

### 2.3 Izlaz iz petlje vodio je nazad u zamku

`pre-flight` §10 („review loop breaker") je jedini mehanizam koji je mogao da
okine — i okidao je: 22 `fix` commita nad jednom invarijantom je tačno taj signal.
Ali njegova instrukcija glasi „imenuj uzrok i **vrati se na taj contract**", a
contract je bio `ZBR-IDENT-01` — sama kompenzacija. Povratak na ugovor
kompenzacije proizvodi bolju kompenzaciju.

**Pravilo koje je ostalo:** `K4` (isti invariant, više krugova) vodi na pitanje
modela, ne na ugovor koji se krpi.

### 2.4 Kapija se nije ni učitala, jer je posao počeo kao „mali"

`pre-flight` opis izuzima „lokalni bug sa jasnom reprodukcijom". ZBR posao je
počeo tačno tako — MIG-005a, jednoredni `ExcludeStornirano` fix. Skill sam to
priznaje u §2 („trigger gleda početak posla, a posao mutira") i **dodat je
08.09.**, usred kompleksa, pa pooštren 09.09. — nijedanput nije zaustavio posao.

**Pravilo koje je ostalo:** `model-pre-zakrpe.md` je `paths:` pravilo, ne skill —
okida na dodirnut fajl, ne na procenu obima.

### 2.5 Cena modela nikad nije izmerena, a bila je blizu nule

Jedina činjenica koja je izmenu modela činila jeftinom — **nema legacy
transakcionih podataka**, sezona je prošla, nijedan klijent nije na starom
programu — stajala je u repou i **nijedan korak je nije tražio**. Da je bila na
stolu u PR-u #291, ZBR kompleks najverovatnije nikad ne bi bio napisan.

**Pravilo koje je ostalo:** „model je preskup" je pretpostavka dok se ne odgovori
na tri pitanja iz `model-pre-zakrpe.md` §5.

---

## 3) Uzrok koji nije bio u pravilima

Obrazac se ne vidi u jednom PR-u. Vidi se u zbiru — a zbir niko ne nosi između
sesija: svaka sesija kreće bez sećanja, svaki pojedinačan predlog je bio lokalno
razuman, i devet lokalno razumnih odluka je dalo globalno pogrešan rezultat.

Zato je jedini deo mehanizma koji stvarno menja ishod **registar u fajlu**
(`docs/DOMEN/KOMPENZACIJE.md`), a ne bolja procena. Registar pretvara „primeti
obrazac kroz mesece" (nemoguće bez sećanja) u „pročitaj fajl" (trivijalno), i
uvodi tvrdu tačku: treći pogodak nad istim redom skida zakrpu sa podrazumevanog.

> Nije nedostajala pamet u trenutku odluke. Nedostajalo je mesto na kome se
> odluke sabiraju.

---

## 4) Kad je detektor mogao da okine

Nad kodom kakav je bio **pre PR-a #291**, svih pet signala je bilo merljivo
`grep`-om:

| Signal | Vrednost tada | Komanda |
|---|---|---|
| K1 rekonstrukcija | 9 call site-ova | `grep -rn '" + "' src-vba/ \| grep -i split` |
| K2 surogat identiteta | `GeneracijaID` bez kolone koja nosi dokument | `grep -rn 'GeneracijaID' src-vba/` |
| K3 kapija umesto strukture | 3 pisca moraju da pečate svaki red | `ZBR_IDENTITET.md` §2 |
| K4 isti invariant, više krugova | KI-007 otvoren, MIG-005a već drugi krug | `git log --grep=ZBR` |
| K5 duplirana činjenica | broj + vlasnik + generacija kao „scope" | `ZBR_IDENTITET.md` §1 |

**5/5.** Ni jedan jedini nije zahtevao znanje koje tada nije postojalo.

---

## 5) Šta ovo NIJE

Nije nalaz da je `GeneracijaID` bio loše napisan — bio je dobro napisan i tačno je
rešavao problem koji mu je zadat. `DOCUMENT_HEADER_LINES.md` §10 ga zato i ne
briše iz istorije: bio je „ispravno rešenje za multi-row model dokumenta".

Nije ni licenca za redizajn. Jedan signal je zapažanje. Granice iz
`DOCUMENT_HEADER_LINES.md` §8 (bez `tblDokumenti`/EAV, bez framework slojeva)
ostaju na snazi — refaktor koji iz ovoga sledi **briše** kod, ne dodaje sloj.
