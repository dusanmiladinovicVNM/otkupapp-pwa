# AgriX — Sažetak Q&A Decision Log-a

**Izvor:** `09_QA_DECISION_LOG.md` (odluke 1–260, sesija 24.07.2026.)
**Status:** sažetak radnog strateškog zapisa
**Namena:** brz pregled ključnih odluka; za tačnu formulaciju i kontekst uvek važi pun log.

> Ovo je destilacija 260 odluka u ključne zaključke po temama. Gde je kasnija
> odluka ispravila raniju, ovde stoji kasnija. Tačka konflikta → prednost ima
> kasnija odluka (vidi §10 osnovnog loga).

---

## 1. Proizvod i paketi (odluke 1–24)

- **AgriX = jedinstven sistem teren–centrala.** Do fakture su PWA i VBA Desktop ravnopravni; od fakture nadalje proces je za sada Desktop-only.
- **Dva paketa, jasni tehnički nazivi:** **AgriX Desktop** (baza) i **AgriX Mobile** (Desktop + PWA Otkupac + PWA Vozač).
  - **Desktop Core:** otkup i osnovna dokumenta, prijemnice/fakture, ambalaža i repromaterijal, skladište, standardni izveštaji i kontrole.
  - **Mobile nije samostalan** — uvek uključuje kompletan Desktop; nema standalone Mobile.
  - **Management PWA** je uključen u svaki paket (i Desktop-only), bez veštačkog ograničavanja funkcija.
- **Posebno plaćeni moduli:** **SEF, Banka, Dispatch** (napredni). Osnovni PWA Vozač je u Mobile-u; napredni Dispatch (rute, kapaciteti, dispečerski pregled) je poseban modul.
- **Banka:** Core dozvoljava ručni unos novca; modul automatizuje izvode, povezivanje uplata, rasknjižavanje, avanse i naloge. Kartice/salda ostaju u Core-u.
- **Core ostaje širok:** sledljivost i WMS (palete, lotovi/serije, jedinice), repromaterijal/agrohemija (zalihe, dugovanja, doziranje, parcele, tretmani), kreiranje faktura, svi postojeći standardni izveštaji. **SEF integracija** je poseban modul.
- **Klijentske specifičnosti** se procenjuju i naplaćuju odvojeno; ako imaju opštu tržišnu vrednost, ulaze u zajednički proizvod. **Nema trajnih klijentskih forkova.**
- **Održavanje** (bug-fix u scope-u, security, redovna ažuriranja) je u godišnjoj pretplati; **novi zahtevi** (procesi, integracije, funkcije) se naplaćuju posebno.

## 2. Komercijalni model, ugovor i podrška (odluke 25–57)

- **Samo godišnja pretplata**, 12 meseci od aktivacije svakog klijenta; nema mesečnog modela. Sezonski klijenti plaćaju punu godinu.
- **Cenovna jedinica:** po pravnom licu + naknada po aktivnim stanicama. Osnovni paket uključuje **do 5 stanica**; svaka preko 5 ima istu fiksnu godišnju cenu (nema tier-ova). **Nema per-user/per-uređaj naplate.** Više pravnih lica → svako puna osnovna cena (grupni popust nije podrazumevan).
- **Hardver** = odvojena jednokratna stavka; hardverska podrška je posebna godišnja naknada; svaki **izlazak na teren** se posebno naplaćuje.
- **Onboarding:** početna politika — prvim klijentima se uglavnom ne naplaćuje. Jednostavan uvoz šifarnika u početku besplatan; **složena migracija** (istorija, salda, veze, finansije) se procenjuje i naplaćuje.
- **Podaci i infrastruktura:** Desktop podaci primarno lokalni + periodične kopije na AgriX Drive; PWA/GAS/Sheets obezbeđuje i kontroliše AgriX; **silo po klijentu**. Klijent je vlasnik podataka, AgriX tehnički obrađivač. Prestanak ugovora → pun izvoz u standardnom formatu + **30 dana** tranzicije.
- **Backup/retention:** backup posle svake Journal promene + pun fajl pri svakom otvaranju + dnevni off-site na Drive. Dnevne kopije 30 dana, mesečne ≥12 meseci. **RPO** dnevni, **RTO** 24h, potvrda prijema kritičnog incidenta u roku od **1 sata**.
- **Podrška:** radnim danima 08–16; u sezoni i vikendom 08–16 (uključeno u pretplatu Desktop i Mobile). Kritični incidenti pokriveni i van radnog vremena; nekritični — odgovor u roku od jednog radnog dana.
- **Kritičan incident** = većina korisnika/stanica ne radi · nemoguć otkup/osnovni dokument · ozbiljan rizik gubitka podataka · centralni VBA neupotrebljiv · sync potpuno blokiran bez fallback-a.

## 3. Tržište, pozicioniranje i granice (odluke 58–76)

- **Jedan proizvod** — nema small/mid/large izdanja; razlike se rešavaju paketima, modulima, brojem stanica i konfiguracijom.
- **Fokus:** voće i povrće su javni fokus; duvan i žitarice imaju zasebne materijale. **Hladnjače** su ključna ciljna grupa (AgriX već pokriva preradu, palete, sledljivost).
- **Cilj sezone 2027:** puniji proizvodni sistem — radni nalozi, norme ulaza/izlaza, proizvodne partije, ambalaža, otpad/prinos, kapaciteti linija, smene/radnici, učinak, direktne integracije s vagama/PLC/senzorima (samo za unapred odobrene uređaje; instalacija se naplaćuje, **kod ostaje AgriX-u**).
- **Trajne strateške granice:** AgriX **nije** računovodstveni ni generički ERP i ne zamenjuje BizniSoft/Pantheon. Podela: AgriX = operacije, dokumenti, sledljivost, logistika, hladnjača, proizvodnja, upravljački pregledi, operativne finansije; **ERP** = glavna knjiga, PDV, završni račun, zarade.
- **Petogodišnji ekosistem:** Enterprise (glavni B2B) + Gazdinstvo (proizvođači/kooperanti) + GGAP (compliance sloj koji ih povezuje). **North Star:** digitalna platforma koja povezuje ceo lanac (proizvođači–otkup–hladnjače–logistika–prerada–sertifikacija–finansije).

## 4. Gazdinstvo, pricing mehanika i strategija (odluke 77–195)

**Gazdinstvo — model i kanali**
- Finansiranje: hladnjača plaća Basic za kooperante · proizvođač kupuje Pro direktno · hladnjača plaća Pro i odbija kroz robu/saldo · prodaja i nezavisno od Enterprise-a. **Dva growth engine-a:** B2B2C preko Enterprise-a i direktna prodaja proizvođačima.
- **Primarni korisnik = proizvođač** (nije white-label hladnjače). Brend: **AgriX Gazdinstvo** uz prikaz povezane hladnjače. Preko Enterprise-a podrazumeva se **Basic**; napredne funkcije = **Pro**. Basic je stvarno upotrebljiv, bez veštačkih limita.
- **Prvih 50 Basic naloga** partner (hladnjača) dobija besplatno; preko 50 plaća po korisniku.

**Identiteti i sinhronizacija**
- Dugoročno globalni identiteti (proizvođač, parcela, katalog proizvoda); trenutno lokalni `KOOP-xxxxx`/`PAR-xxxxx` + kasnije mapiranje. Firma već globalna kroz `Cxxx`. Desktop broj dokumenta je kanonski i trajan; PWA daje privremeni broj → konačan pri sync-u. **Multi-Enterprise** je dugoročni cilj (sad jedna veza). **Sheets baza ostaje.**

**Strategija, prioriteti, rizici**
- Prioritet do 2027: **pun proizvodni sistem hladnjače je najviši prioritet**; Gazdinstvo — ključne Pro funkcije bez velikih arhitektonskih projekata koji bi ga usporili.
- Konkurentska prednost: pokrivanje cele firme + unos jednom na izvoru + brz/fleksibilan razvoj. Glavna pretnja: dobro finansiran konkurent s jakim sales timom. **Bottleneck na 30–50 klijenata = prodaja**; prvi prodavac tek tada (posle standardizacije proizvoda/procesa/referenci). Uloga osnivača: arhitektura i strategija ostaju kod osnivača; razvoj/podrška/onboarding/prodaja se postepeno delegiraju.
- **Custom rad:** time-and-materials (ne fiksna cena); klijent dobija procenu sati + maksimalni budžet, prekoračenje uz novo pisano odobrenje; jedna standardna satnica uz mogući popust. Promena roadmap-a za prospect zahteva **pisanu nameru/prihvaćenu ponudu** (usmeno interesovanje nije dovoljno).

**Pricing mehanika (obnove, stanice, moduli, instance)**
- Obnova nije automatska; obaveštenje 30 dana pre isteka; bez obnove → **read-only** 30 dana (pregled/izvoz, bez novog unosa).
- Usred godine: Desktop→Mobile i nova stanica = proporcionalna doplata do isteka; Mobile→Desktop downgrade tek pri obnovi, bez refundacije.
- **Mobile multiplikator:** Desktop Otkup + Mobile ≥ 2× cena Desktop Otkup-a. Moduli (SEF/Banka/Dispatch) — fiksna godišnja cena po pravnom licu, ista bez obzira na Desktop/Mobile; plaćaju se jednom i koriste kroz sve instance pravnog lica.
- **Proizvodni dodatak (Hladnjača/Proizvodnja):** nezavisan Desktop dodatak, jedan pogon; dodatni pogon = dodatna Desktop instanca (+ dodatni dodatak). Sve instance istog pravnog lica — isti ugovor i datum obnove.
- Desktop bez limita korisnika; trenutno single-active-user, Management PWA podržava više pregleda; budući multi-user ulazi u postojeću pretplatu (ne kao modul), prioritet tek ako konkretan ugovor zahteva.
- Promena cena: novi klijenti odmah novu cenu; postojeći — jedna prelazna godina po staroj/prelaznoj ceni. Osnovna pretplata se ne menja tokom plaćenog perioda.

**Core vs modul + pilot**
- **Core kriterijum:** funkcija je nužna da AgriX ispuni osnovno obećanje proizvoda. **Modul:** jasno merljiva dodatna vrednost koju klijent može posebno da plati. Pri izdvajanju modula postojeći korisnici zadržavaju funkciju besplatno do isteka ugovorne godine, plaćaju od obnove.
- **Pilot:** kritične funkcije prvo kod jednog klijenta (onaj koji ju je tražio i aktivno testira); besplatna tokom pilota i do kraja te ugovorne godine; posle → Core ili modul.

**Životni ciklus i privatnost (Gazdinstvo)**
- Naplata: samo godišnja pretplata. **Probe:** 30 dana Basic bez kartice → posle mora kupiti Basic ili read-only. **Pro proba:** 30 dana na poverenje, samo uz plaćen Basic, **jednom**; kasnije uz unapred evidentiranu uplatu. **Aktivacija na poverenje:** pristupni kod odmah, ali ako uplata ne stigne za 7 dana → blokada.
- Upgrade Basic→Pro proporcionalno; obnova Pro = puna godišnja cena Pro (uključuje Basic); downgrade Pro→Basic pri obnovi (Pro podaci ostaju vidljivi/zaključani). Prekid saradnje s hladnjačom → hladnjačom finansiran Pro se odmah deaktivira, nalog na Basic. Prestanak Enterprise ugovora → finansirani nalozi 30 dana za samostalnu obnovu, zatim read-only.
- **Podaci proizvođača:** istorija otkupa/dokumenata/ambalaže/salda trajno dostupna proizvođaču; samostalan izvoz u standardnim formatima. Brisanje naloga briše lične podatke, ali zajednički poslovni dokumenti ostaju kod hladnjače (strana u dokumentu, obaveza čuvanja). Hladnjača vidi **samo podatke svog odnosa**; dodatni podaci (plan, tretmani, prinos) samo uz jasnu saglasnost.
- **Cene se objavljuju:** Basic i Pro tačne godišnje cene; Enterprise i Mobile rasponi; dodatna stanica tačan iznos; Hladnjača/Proizvodnja cenovni raspon.

## 5. AgriX Savetnik (odluke 196–231)

- **Poseban proizvod:** savetnik/agronom iz jednog interfejsa vodi više gazdinstava; **naplata po broju aktivnih gazdinstava** (ista tarifa za nezavisne savetnike i interne agronomske službe). Samo godišnja pretplata; proba 30 dana do 10 gazdinstava.
- Cena Savetnika pokriva aktivna gazdinstva (ona ne plaćaju zaseban Pro). Proizvođač zadržava svoj nalog i pristup podacima; prekid → proizvođač čuva sve, savetnik odmah gubi pristup.
- Funkcije: savetnik šalje obavezujući radni nalog ili neobaveznu preporuku (stiže u Pro nalog proizvođača), prati status/kašnjenja/odstupanja. **GGAP granica:** GGAP minimum je u GGAP modulu; agrosaveti iznad toga pripadaju Savetniku.
- **Prioritet:** osnovna verzija do 2027, bez usporavanja Enterprise proizvodnog sistema (prvo jedan savetnik → više gazdinstava; timovi kasnije). Lansiranje javno čim je stabilno (zatvoreni pilot nije obavezan). Kanal: direktno obraćanje agronomima/firmama.
- **Partnerstvo/provizije:** savetnik može preporučivati Gazdinstvo/Enterprise uz **fiksnu proviziju** (ne procenat). Gazdinstvo: provizija za prvu prodaju i svaku obnovu dok je korisnik aktivan; Enterprise: samo jednokratno pri prvom ugovoru. Cena Savetnika se ne objavljuje (individualna ponuda).
- **Dugoročno:** od alata ka platformi za pronalaženje/ugovaranje/plaćanje savetnika (marketplace **ne pre kraja sezone 2027**).

## 6. Hladnjača/Proizvodnja i buduće vertikale (odluke 232–251)

- **Aktivna prodaja čim osnovni tok bude stabilan** (ne čeka se ceo roadmap). Palete sveže i prerađene robe su već u produkciji; početni prioritet postojećim klijentima radi lakše validacije.
- **Red razvoja:** radni nalozi → norme → ambalaža → otpad/prinos → integracije s vagama/opremom → kapaciteti linija, smene, učinak radnika. **Klijent kupuje samo ono što postoji na dan prodaje** — roadmap nije ugovorno obećanje.
- Primarna grupa: hladnjače (prijem, klasiranje, zamrzavanje, pakovanje, palete). Dugoročno drugi agro-prehrambeni prerađivači samo kad se prirodno uklope (ne graditi generički proizvodni ERP). Vertikale: zajedničko jezgro, posebni komercijalni paketi/prezentacija/cena; ozbiljan razvoj tek uz konkretnog kupca. Mogući red posle 2027: žitarice/silosi/mlinovi, duvan, sušare, vinarije.
- **Naknada modula:** dok se razvija/standardizuje/proverava kod prvog klijenta — nema posebne godišnje naknade; posle uspešne provere počinje godišnja naplata. **Uvođenje modula postojećem klijentu je uvek besplatno** (naplaćuje se samo početni onboarding celog sistema kod novog klijenta).
- **Ciljevi do 2027:** 10–20 aktivnih pravnih lica; proizvodni modul očekivano koristi >80% Enterprise klijenata; geografski fokus **samo Srbija**; reference primarno hladnjače za voće i povrće. Modul je standardni deo svake ponude za hladnjače (klijent može da ga ne kupi).

## 7. Prodaja, demonstracija i reference (odluke 252–260)

- **Demo** kreće od konkretnog problema klijenta i završava prikazom celog toka (otkup → prerada → palete → zalihe → dokumenti → upravljački pregled). Pre demoa kratak razgovor o procesima/prioritetima; poseban demo samo za kvalifikovan lead s pristupom donosiocu odluke.
- **Enterprise nema probni period ni produkcioni pilot** — postoji samo demo sa dummy podacima; kvalifikovan lead može dobiti vremenski ograničen pristup standardizovanoj dummy demo instanci (jedan scenario za sve). Demo prikazuje ceo ekosistem i sve module; ponuda jasno odvaja kupljeno od opcionog. Razvojne funkcije mogu se prikazati, ali jasno označene kao nedostupne za ugovaranje.
- **Reference:** AgriX može javno navesti klijenta ako to ugovorom nije izričito zabranjeno.

## 8. GGAP (raniji deo sesije)

- **GGAP nije deo Pro-a** — poseban je Enterprise dodatak koji kupuje hladnjača za mrežu kooperanata; **jedna fiksna godišnja cena po pravnom licu**, pokriva sve GGAP kooperante (bez per-user tier-ova). Korisnik u GGAP-u dobija sve potrebne Gazdinstvo funkcije za usklađenost bez dodatne Pro naknade samo zbog GGAP-a.
- Softverska cena = platforma + tehnička podrška + compliance workflow; **stručno savetovanje, priprema dokumentacije, audit i konsulting se naplaćuju odvojeno**. Redosled: prvo softver → mreža eksternih stručnjaka → dugoročno moguća interna ekipa. Softver sam nikada ne garantuje sertifikat.
- Redovna prodaja tek posle validacije sadržaja od kompetentnog konsultanta i **najmanje jednog uspešnog realnog projekta**. Do 2027 GGAP je ograničen na konceptualnu pripremu i stručnu validaciju (ozbiljan razvoj posle proizvodnog sistema i stabilizacije Enterprise-a).

---

## Otvorena pitanja (nerešeno u sesiji)

1. Precizna pravna pravila za povlačenje saglasnosti proizvođača za dodatno deljenje podataka.
2. Konačan redosled velikih post-2027 inicijativa: GGAP, marketplace Savetnika, multi-Enterprise arhitektura.
3. Konkretni cenovnici i apsolutni iznosi (Enterprise, Mobile, dodatne stanice, moduli, Gazdinstvo, Savetnik).
4. Tačan trenutak prelaska sa besplatnog na plaćeni onboarding novih Enterprise klijenata.
5. Formalni uslovi partnerskog programa, provizije i atribucija lead-a.

## Pravilo tumačenja

- **Kasnija odluka ima prednost** nad ranijom u konfliktu.
- Već implementirane i produkciono potvrđene funkcije ne predstavljati kao buduće.
- Roadmap nije prodajno obećanje ni ugovorna obaveza bez posebnog pisanog ugovora.
- AgriX ostaje **zajednički proizvod bez trajnih klijentskih forkova**.
