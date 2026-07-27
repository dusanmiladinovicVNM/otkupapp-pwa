# AgriX — Q&A Decision Log

**Datum sesije:** 24.07.2026.  
**Status:** radni strateški zapis  
**Obuhvat:** odluke, korekcije, pretpostavke i otvorena pitanja iz Q&A sesija — numerisane odluke 1–321, 323–378 i 401–422, uz serije A, BC, C, D, I, IP, L, LEG, M, MKT, ML, ON, P, PRT, Q, S. Brojevi 322 i 379–400 se ne koriste; videti odeljak 27.  
**Tematski indeks:** `09B_ODLUKE_PO_OBLASTIMA.md`

> Ovaj dokument čuva kompletan suštinski sadržaj sesije. Pomoćne formulacije asistenta koje nisu menjale odluke nisu prenete. Kada je kasnija odluka ispravila raniju, važi kasnija formulacija i to je izričito označeno.

---

## 1. Osnovna definicija proizvoda i paketi

1. **Definicija AgriX-a.** AgriX je jedinstven sistem teren–centrala. Do fakture su PWA i VBA Desktop ravnopravno podržani načini unosa. PWA štedi vreme jer se podatak unosi jednom na mestu nastanka i sinhronizuje sa centralom. Od fakture nadalje proces je trenutno Desktop-only.
2. **Desktop-only paket.** Desktop-only je legitimno kompletan proizvod, ali je operativno manje efikasan od kombinacije Desktop + PWA i značajno je jeftiniji.
3. **Desktop Core.** Obuhvata otkup i osnovna dokumenta, prijemnice i fakture, ambalažu i repromaterijal, skladište, standardne izveštaje i kontrole.
4. **Posebno plaćeni moduli.** SEF, Banka i Dispatch prodaju se odvojeno.
5. **Management PWA.** Uključen je u svaki paket, uključujući Desktop-only, bez veštačkog ograničavanja funkcija. U Desktop-only modelu nema stvarnog terenskog real-time otkupa, ali uprava ima mobilni pregled količina, zaliha, finansijske pozicije, kartica i izveštaja.
6. **Nazivi paketa.** Koriste se jasni tehnički nazivi: **AgriX Desktop** i **AgriX Mobile**.
7. **AgriX Mobile.** Standardno uključuje i PWA Otkupac i PWA Vozač.
8. **Mobile nije samostalan proizvod.** AgriX Mobile uvek uključuje kompletan Desktop; ne postoji standalone Mobile.
9. **Osnovni PWA Vozač.** Uključuje preuzimanje sa stanica, zbirni/transportni dokument, količine i status, kao i centralnu sinhronizaciju.
10. **Napredni Dispatch.** Poseban plaćeni modul koji uključuje raspoređivanje vozila i vozača, rute, kapacitete, neraspoređenu robu i dispečerski pregled.
11. **SEF i Banka.** U početku se prodaju kao odvojeni moduli. Kasnije je moguć bundle, ali početna ponuda ostaje jednostavna.
12. **Banka — Core granica.** Desktop Core dozvoljava ručni unos novčanih transakcija.
13. **Banka — automatizacija.** Banka modul automatizuje uvoz izvoda, povezivanje uplata, raspoređivanje/rasknjižavanje, avanse i platne naloge.
14. **Kartice i salda.** Partner kartice i salda ostaju u Core-u; Banka automatizuje njihovo popunjavanje i obradu.
15. **SEF modul.** Obuhvata slanje izlaznih faktura, praćenje statusa, storniranje prema procesu, preuzimanje ulaznih faktura i njihovo povezivanje sa AgriX evidencijom.
16. **Sledljivost i WMS.** Napredne funkcije zaliha, prijema, ambalaže, paleta, skladišnih jedinica, lotova/serija i sledljivosti proizvođač–parcela–prijem–prerada–kupac ostaju u Core-u.
17. **Repromaterijal i agrohemija.** Za sada kompletan tok zaliha, izdavanja, dugovanja, salda, doziranja, parcela, kultura i tretmana ostaje u Core-u. Odvajanje u budućnosti moguće je samo ako složenost i podrška to opravdaju.
18. **Fakture.** Kreiranje i evidencija faktura su Core; SEF integracija je poseban modul.
19. **Standardni izveštaji.** Svi postojeći standardni operativni, finansijski, skladišni, sledljivostni i upravljački izveštaji su Core.
20. **Klijentski izveštaji.** Novi izveštaji specifični za jednog klijenta procenjuju se i naplaćuju odvojeno.
21. **Klijentske funkcije.** Specifična funkcionalnost se posebno procenjuje i naplaćuje. Ako ima opštu tržišnu vrednost, ulazi u zajednički proizvod.
22. **Bez trajnih forkova.** Ne održavaju se trajne klijentske grane proizvoda.
23. **Održavanje.** Bug-fix u ugovorenom scope-u, bezbednosne ispravke i redovna ažuriranja uključeni su u godišnju pretplatu.
24. **Novi zahtevi.** Novi procesi, integracije, funkcije i klijentske izmene naplaćuju se odvojeno.

---

## 2. Komercijalni model, ugovor i podrška

25. **Osnovni model.** AgriX se prodaje kroz godišnju pretplatu koja uključuje korišćenje, ažuriranja, bug-fix i standardnu podršku.
26. **Nema mesečnog modela.** Enterprise nema mesečnu pretplatu niti mesečno plaćanje.
27. **Trajanje ugovora.** Standardni ugovor traje 12 meseci.
28. **Sezonski klijenti.** Plaćaju punu godišnju pretplatu.
29. **Osnovna cenovna jedinica.** Cena se formira po pravnom licu, uz dodatnu naknadu prema broju aktivnih otkupnih stanica/lokacija.
30. **Više pravnih lica.** Svako pravno lice plaća punu osnovnu cenu; grupni popust nije podrazumevan.
31. **Nema per-user naplate.** Ne naplaćuje se po korisniku ili uređaju.
32. **Hardver.** Prodaje se kao odvojena jednokratna stavka.
33. **Hardverska podrška.** Posebna godišnja naknada može obuhvatiti prioritetnu intervenciju, daljinsku dijagnostiku, reinstalaciju/konfiguraciju i zamenski uređaj prema uslovima.
34. **Izlazak na teren.** Svaka fizička intervencija na lokaciji naplaćuje se posebno.
35. **Onboarding.** Početni onboarding/migracija/integracije mogu se posebno proceniti, ali je početna go-to-market politika da se prvim klijentima onboarding uglavnom ne naplaćuje.
36. **Osnovni uvoz podataka.** Jednostavan uvoz šifarnika iz Excel-a, Infosys-a ili drugog sistema može u početku biti besplatan.
37. **Složena migracija.** Istorijski dokumenti, salda, kartice, veze i finansijski podaci procenjuju se i naplaćuju posebno.
38. **Standardna podrška.** Uključuje email/poruke/telefon, daljinsku dijagnostiku, pomoć u korišćenju, bug-fix i sezonski prioritet.
39. **Kritični incidenti.** Incidenti koji blokiraju rad pokriveni su i van radnog vremena.
40. **Podaci klijenta.** Desktop podaci su primarno lokalni, uz periodične kopije na AgriX Drive.
41. **PWA infrastruktura.** PWA/GAS/Sheets infrastrukturu obezbeđuje i kontroliše AgriX.
42. **Silo arhitektura.** Svaki klijent/pravno lice ima odvojeni Drive silo.
43. **Vlasništvo nad podacima.** Klijent je vlasnik poslovnih podataka; AgriX je tehnički administrator/obrađivač, ne vlasnik podataka.
44. **Prestanak ugovora.** Klijent dobija kompletan izvoz u standardnom formatu i opcionu plaćenu pomoć pri migraciji na drugi sistem.
45. **Tranzicioni rok.** Posle prestanka ugovora postoji 30 dana za izvoz i prelazak.
46. **Retention rezervnih kopija.** Dnevne kopije čuvaju se 30 dana, mesečne najmanje 12 meseci.
47. **Lokalni backup.** Backup se pravi nakon svake Journal promene, kompletna kopija fajla pri svakom otvaranju, a najmanje jednom dnevno off-site kopija na AgriX Drive.
48. **RPO.** Cloud RPO može biti dnevni jer su lokalne kopije češće.
49. **RTO.** Cilj za kritičan incident je povrat funkcionalnog stanja u roku od 24 sata.
50. **Kritični odziv.** Potvrda prijema i početak dijagnostike u roku od jednog sata.
51. **Definicija kritičnog incidenta.** Većina korisnika/stanica ne može da radi; nije moguće evidentirati otkup ili izdati osnovni dokument; postoji ozbiljan rizik gubitka/oštećenja podataka; centralni VBA je neupotrebljiv; sinhronizacija je potpuno blokirana bez fallback-a.
52. **Nekritični incidenti.** Odgovor u roku od jednog radnog dana; vreme rešavanja zavisi od prioriteta, uticaja i složenosti.
53. **Standardno radno vreme podrške.** Radnim danima 08:00–16:00.
54. **Sezonska podrška.** Tokom definisane sezone otkupa redovna podrška postoji i vikendom.
55. **Vikend radno vreme.** 08:00–16:00.
56. **Definicija sezone.** Određuje se po klijentu prema kulturama i stvarnom periodu otkupa.
57. **Cena sezonske podrške.** Vikend podrška u sezoni uključena je u standardnu pretplatu za Desktop i Mobile.

---

## 3. Tržište, pozicioniranje i granice proizvoda

58. **Širina ciljnog tržišta.** Cilj su i mali otkupljivači sa jednom stanicom i veliki sistemi sa više stanica.
59. **Jedan proizvod.** Ne postoje posebna small/mid/large izdanja. Razlike se rešavaju paketima, modulima, brojem stanica i konfiguracijom.
60. **Operativni režimi.** Klijent bira način rada kroz podešavanja istog proizvoda.
61. **Primarno pozicioniranje.** Voće i povrće su glavni javni fokus; duvan i žitarice imaju posebne stranice/materijale.
62. **Hladnjače.** Ključna su ciljna grupa jer AgriX već pokriva dokumentaciju prerade i sledljivosti.
63. **Proizvodnja u Core proizvodnom domenu.** Tok sirovine, klasiranje, otpad/kalo, proizvodne partije, ambalaža, palete, skladište i sledljivost čine osnovu proizvodnog sistema za hladnjače.
64. **Današnje pozicioniranje.** AgriX je kompletan operativni sistem za otkup, dokumentaciju prerade, skladište, sledljivost i upravljanje.
65. **Cilj za sezonu 2027.** Razviti puniji proizvodni sistem sa planiranjem, normama, kapacitetom linija, radnicima, učinkom i integracijom sa opremom.
66. **Minimalni scope proizvodnje 2027.** Radni nalozi, norme ulaza/izlaza, proizvodne partije, utrošak ambalaže, otpad/prinos, gotovi proizvodi i sledljivost, kapaciteti linija, smene/radnici, učinak i direktne integracije sa vagama/mašinama/senzorima.
67. **Integracije sa opremom.** Standardizuju se samo za unapred odobrene vage, PLC-eve, senzore i mašine.
68. **Naplaćivanje integracija.** Svaka konkretna instalacija, konfiguracija, testiranje i puštanje u rad naplaćuju se klijentu.
69. **Vlasništvo integracionog koda.** AgriX zadržava framework i kod; podrška za uređaj ulazi u zajednički proizvod.
70. **Finansiranje nove integracije.** Ako ima široku tržišnu vrednost, razvoj finansira AgriX. Ako je vrlo specifična, razvoj plaća prvi klijent, ali kod ostaje AgriX-u.
71. **Željena reakcija posle demo-a.** „Ovo je sistem koji pokriva celu firmu.“
72. **Granica prema ERP-u.** AgriX nije računovodstveni ERP i ne zamenjuje BizniSoft, Pantheon i slične sisteme.
73. **Podela odgovornosti.** AgriX pokriva operacije, dokumente, sledljivost, logistiku, kooperante, hladnjaču, proizvodnju, upravljačke preglede i operativne finansijske integracije. ERP pokriva glavnu knjigu, PDV, završni račun, zarade i zakonsko računovodstvo.
74. **Trajne strateške granice.** Ne praviti opšti računovodstveni ERP, ne praviti generički ERP za sve industrije i ostati fokusiran na agroindustriju.
75. **Petogodišnji ekosistem.** Enterprise je glavni B2B proizvod, Gazdinstvo je proizvod za proizvođače/kooperante, a GGAP je compliance sloj koji ih povezuje.
76. **North Star.** Digitalna platforma koja povezuje proizvođače, otkupljivače, hladnjače, logistiku, preradu, sertifikaciju i finansijske tokove.

---

## 4. Gazdinstvo — proizvod, poslovni model i privatnost

77. **Kanali finansiranja Gazdinstva.** Hladnjača može finansirati Basic za kooperante; proizvođač može direktno kupiti Pro; hladnjača može platiti Pro i odbiti iznos kroz robu/saldo; Gazdinstvo se prodaje i nezavisno od Enterprise-a.
78. **Samostalna vrednost.** Radna pretpostavka je da Gazdinstvo mora imati vrednost i bez Enterprise veze; Enterprise veza dodaje funkcije saradnje.
79. **Dva growth engine-a.** Enterprise-to-Gazdinstvo B2B2C i direktna prodaja proizvođačima.
80. **Basic preko Enterprise-a.** Preko Enterprise-a se podrazumeva Basic; napredne funkcije zahtevaju Pro.
81. **Primarni korisnik.** Proizvođač je primarni korisnik Gazdinstva. Hladnjača dobija transparentnost, lojalnost kooperanata i bolju kontrolu, ali aplikacija nije njen white-label proizvod.
82. **Multi-Enterprise dugoročno.** Jedan proizvođač će moći da bude povezan sa više Enterprise sistema. Svaka hladnjača vidi samo podatke svog odnosa, a proizvođač konsolidovan pregled.
83. **Trenutno ograničenje.** Trenutno je jedna Enterprise veza; multi-Enterprise nije kratkoročni prioritet.
84. **Globalni identitet proizvođača.** Dugoročno jedan globalni proizvođački identitet; trenutni `KOOP-xxxxx` ostaje lokalni po Desktop-u i kasnije se mapira.
85. **Globalni identitet firme.** Već postoji kroz `Cxxx`, koji se koristi za Drive i ostale resurse.
86. **Globalni identitet parcele.** Dugoročno postoji globalni identitet; trenutni `PAR-xxxxx` ostaje lokalni i kasnije se mapira.
87. **Globalni katalog proizvoda.** Dugoročno zajednički katalog uz lokalne alias-e i mogućnost lokalnih nezavisnih proizvoda.
88. **Dokumentni identitet.** Desktop broj dokumenta je kanonski i trajan. PWA kreira privremeni broj, a pri sinhronizaciji zapis dobija konačan Desktop broj.
89. **Robusnost identiteta.** Duplikati se sprečavaju, svi brojevi su trajni, a pomoćni PWA/server identifikatori ostaju zbog robusne sinhronizacije.
90. **Sheets ostaje.** Sheets baza se ne briše.
91. **Prioritet do 2027.** Puni proizvodni sistem hladnjače ima najviši razvojni prioritet.
92. **Konkurentska prednost.** Kombinacija pokrivanja cele firme, unosa jednom na izvoru i fleksibilnog/brzog razvoja.
93. **Glavna pretnja.** Dobro finansiran konkurent sa jakim industrijskim sales timom može brzo zauzeti tržište.
94. **Bottleneck pri 30–50 klijenata.** Prodaja.
95. **Prvi prodavac.** Tek oko 30–50 klijenata, nakon standardizacije proizvoda, referenci i prodajnog procesa.
96. **Uloga osnivača.** Arhitektura proizvoda i strateške odluke ostaju kod osnivača; razvoj, podrška, onboarding i prodaja postepeno se delegiraju.
97. **Izvor vrednosti.** Kratkoročno Enterprise klijenti; srednjoročno Enterprise + Gazdinstvo; veoma dugoročno podaci i mrežni efekat.
98. **Rizici po prioritetu.** Prespor razvoj/osvajanje tržišta; odlazak u generički ERP; previše custom razvoja.
99. **Prodajni kanali.** Direktno obraćanje i lični demo; zatim SEO/digitalni oglasi; zatim partneri.
100. **Lead scoring 2027.** Svi segmenti mogu u scoring, ali veliki sistemi trenutno nisu primarni zbog kapaciteta za specifične zahteve.
101. **Rang lead-a.** Verovatnoća brzog zatvaranja, prihod, fit sa postojećim funkcijama/količina custom rada, referentna vrednost.
102. **Nema automatskog odbijanja segmenta.** Svaki lead se procenjuje prema prihodu, kapacitetu i strateškoj vrednosti.
103. **Veliki custom zahtev.** Prihod mora opravdati odlaganje roadmap-a; zatim se gleda vrednost za buduće klijente.
104. **Procena odlaganja.** Za veliki zahtev mora se eksplicitno navesti šta se odlaže i zašto je prihod/strateška vrednost dovoljna.
105. **Prioritet razvoja.** Kritični incidenti postojećih klijenata; funkcije koje direktno donose ugovor; strateški roadmap; ostala poboljšanja.
106. **Promena roadmap-a za prospect-a.** Potrebna je pisana namera ili prihvaćena ponuda; usmeno interesovanje nije dovoljno.
107. **Custom naplata.** Time-and-materials, ne fiksna cena.
108. **Kontrola budžeta.** Klijent dobija procenu sati i maksimalni budžet; prekoračenje zahteva novo pisano odobrenje.
109. **Fakturisanje custom rada.** Mesečno ili po završetku, prema dokumentovanim satima i unutar odobrenog limita.
110. **Satnica.** Jedna standardna satnica, uz mogući individualni popust većim ili dugoročnim klijentima. *`Superseded` 27.07.2026. odlukom 409 — dve satnice, bez individualnog popusta.*
111. **Osnova popusta na custom rad.** Veći unapred dogovoreni obim i dugoročna ukupna vrednost odnosa. *`Deleted` 27.07.2026. odlukom 418 — pregovaračkih popusta nema.*
112. **Maksimalni custom popust.** Ne postoji fiksni maksimum; odlučuje se pojedinačno. *`Deleted` 27.07.2026. odlukom 418 — pregovaračkih popusta nema.*
113. **Pretplata bez popusta.** Isti paket i broj stanica znače istu godišnju cenu.
114. **Godišnji obračun.** Pretplata traje 12 meseci od datuma aktivacije svakog klijenta.
115. **Obnova.** Nije automatska; zahteva potvrdu i plaćanje.
116. **Obaveštenje o obnovi.** Šalje se 30 dana pre isteka.
117. **Bez obnove.** Sistem prelazi u read-only režim tokom 30 dana; moguć je pregled i izvoz, bez novog unosa i obrade.
118. **Desktop → Mobile usred godine.** Plaća se proporcionalna razlika do isteka ugovora.
119. **Mobile → Desktop downgrade.** Moguć tek pri obnovi, bez refundacije.
120. **Nova stanica usred godine.** Proporcionalna naknada do isteka ugovora, puna pri obnovi.
121. **Uključene stanice.** Osnovni paket uključuje do pet aktivnih stanica.
122. **Preko pet stanica.** Svaka dodatna stanica ima istu fiksnu godišnju cenu; nema tier-ova.
123. **Ista cena stanice.** Dodatna stanica košta isto u Desktop i Mobile paketu; razlika je u osnovnoj ceni paketa.
124. **Moduli po pravnom licu.** SEF, Banka i Dispatch imaju fiksnu godišnju cenu po pravnom licu.
125. **Ista cena modula.** Cena modula ne zavisi od Desktop/Mobile paketa.
126. **Mobile multiplikator.** Cena Mobile-a definiše se kao unapred utvrđen odnos prema Desktop-u. *`Superseded` 27.07.2026. odlukom 414 — aditivni Mobile dodatak od 1.000 €, ne multiplikator.*
127. **Minimalni odnos.** Desktop Otkup + Mobile treba da bude najmanje dva puta cena Desktop Otkup-a. *`Superseded` 27.07.2026. odlukom 414.*
128. **Struktura ponude.** Desktop Otkup je baza; Mobile je dodatak; Hladnjača/Proizvodnja je nezavisan Desktop dodatak; SEF, Banka i Dispatch su odvojeni.
129. **Mobile i proizvodnja.** Mobile pokriva teren/transport; proizvodnja ostaje Desktop funkcionalnost.
130. **Proizvodni dodatak — početno.** Nakon standardizacije ima fiksnu godišnju naknadu. *`Closed` 27.07.2026. odlukom 421 — 400 € po proizvodnom pogonu, uslov standardizacije ispunjen.*
131. **Jedan pogon.** Proizvodni dodatak pokriva jedan proizvodni pogon.
132. **Više pogona.** Dodatni pogon istog pravnog lica zahteva dodatnu Desktop instancu i dodatni proizvodni dodatak.
133. **Dodatna Desktop instanca.** Cena se određuje individualno prema razlogu, scope-u i složenosti. *`Superseded` 27.07.2026. odlukom 413 — fiksnih −50 %, bez individualnog određivanja.*
134. **Jedan ugovor.** Sve instance istog pravnog lica imaju isti ugovor i datum obnove.
135. **Moduli kroz instance.** SEF, Banka i Dispatch plaćaju se jednom po pravnom licu i koriste kroz sve njegove instance.
136. **Desktop korisnici.** Nema limita/licence po korisniku.
137. **Trenutni Desktop concurrency.** Desktop je tehnički single-active-user; Management PWA podržava više istovremenih pregleda.
138. **Budući multi-user.** Biće uključen u postojeću Desktop pretplatu, ne kao poseban modul.
139. **Prioritet multi-user razvoja.** Pre sezone 2027. samo ako konkretan ugovor to zahteva; inače posle sezone, iza proizvodnog sistema.
140. **Inkrementalna isporuka proizvodnje.** Funkcionalne celine puštaju se kada postanu stabilne; ne čeka se završetak celog roadmap-a.
141. **Rizik rollout-a.** Niskorizične funkcije mogu odmah šire; kritične funkcije prvo kod jednog klijenta.
142. **Uslov za širi rollout kritične funkcije.** Tehnička provera, realan rad bez kritičnih grešaka i potvrda klijenta da proces odgovara praksi.
143. **Prvi pilot klijent.** Onaj koji je tražio funkciju i spreman je aktivno da testira.
144. **Cena kritične pilot funkcije.** Besplatna tokom pilota i do kraja te ugovorne godine.
145. **Posle pilota.** Na sledećoj obnovi funkcija ulazi u Core ili se plaća kao modul, zavisno od klasifikacije.
146. **Pilot ugovori.** Uslovi se dogovaraju pojedinačno; nije obavezan standardni pilot ugovor.
147. **Core kriterijum.** Funkcija ulazi u Core ako je potrebna da AgriX ispuni osnovno obećanje proizvoda.
148. **Modul kriterijum.** Poseban modul je funkcija sa jasno merljivom dodatnom vrednošću koju klijent može posebno da plati, osim ako je neophodna za osnovno obećanje.
149. **Postojeći korisnici kod izdvajanja modula.** Zadržavaju funkciju besplatno do isteka ugovorne godine; plaćaju od obnove.
150. **Povećanje cena.** Novi klijenti odmah dobijaju novu cenu; postojeći prelazni period.
151. **Zaštita postojeće cene.** Jedna naredna ugovorna godina po staroj/prelaznoj ceni, zatim puna aktuelna cena.
152. **Cena tokom ugovora.** Osnovna pretplata se ne menja tokom plaćenog perioda; novi moduli/stanice/paketi se doplaćuju proporcionalno.
153. **Plaćanje novih klijenata.** Godišnja pretplata se plaća pre aktivacije.
154. **Obnova postojećih.** Faktura sa rokom plaćanja 30 dana.
155. **Kašnjenje pri obnovi.** Rešava se individualno prema istoriji i razlogu, bez automatskog trenutnog blokiranja.
156. **Javne cene Enterprise-a.** Objavljuju se rasponi za Desktop i Mobile. *`Superseded` 27.07.2026. odlukom 416.*
157. **Javna cena proizvodnje.** Objavljuje se cenovni raspon za Hladnjača/Proizvodnja. *`Superseded` 27.07.2026. odlukom 416.*
158. **Javna cena dodatne stanice.** Objavljuje se tačan godišnji iznos za svaku stanicu preko pet.
159. **Gazdinstvo cene.** Objavljuju se tačne godišnje cene Basic i Pro.
160. **Basic samostalno.** Može se kupiti bez Enterprise-a.
161. **Enterprise Basic benefit.** Prvih 50 Basic korisnika partner dobija bez dodatne naknade; preko 50 plaća po korisniku.
162. **Gazdinstvo prioritet 2027.** Razvijaju se ključne Pro funkcije i aktivno pridobijaju proizvođači, ali bez velikih arhitektonskih projekata koji bi usporili Enterprise proizvodni sistem.
163. **Kanal rasta Gazdinstva.** Enterprise i direktni kanal paralelno, sa prioritetom Enterprise kanala.
164. **Glavni cilj Gazdinstva 2027.** Rast broja plaćenih Pro korisnika, ne broj besplatnih Basic naloga.
165. **Plaćanje Pro-a.** Direktno plaćanje proizvođača i plaćanje preko hladnjače su ravnopravni; bira se put sa najmanjom preprekom.
166. **Pro preko hladnjače.** Licenca traje do isteka Enterprise ugovora hladnjače i obnavlja se zajedno sa njim.
167. **Pro aktivacija usred godine.** Hladnjača plaća proporcionalno do isteka ugovora.
168. **Prekid saradnje.** Pro licenca koju finansira hladnjača odmah se deaktivira kada proizvođač prestane da sarađuje sa njom.
169. **Posle gašenja Pro-a.** Nalog se vraća na Basic; Pro podaci ostaju vidljivi, ali nema novih Pro unosa.
170. **Basic proba.** Svaki novi korisnik dobija 30 dana besplatnog Basic-a bez kartice i automatske naplate.
171. **Posle Basic probe.** Samostalni korisnik mora kupiti Basic; bez uplate nalog prelazi u read-only.
172. **Gazdinstvo naplata.** Basic i Pro imaju samo godišnju pretplatu.
173. **Aktivacija na poverenje.** Korisnik se registruje, dobija pristupni kod i nakon izjave da je uplatio odmah dobija funkcionalnost. Ako uplata ne stigne u roku od sedam dana, nalog se blokira.
174. **Blokada zbog neplaćanja.** Pristup se potpuno blokira do evidentirane uplate.
175. **Pro proba.** Dostupna je samo korisniku sa plaćenim Basic-om i pokreće se kada korisnik sam odluči.
176. **Jedna Pro proba.** Jednom se daje 30 dana Pro na poverenje; svaka kasnija Pro aktivacija zahteva unapred evidentiranu uplatu.
177. **Po isteku Pro probe.** Povratak na plaćeni Basic; Pro podaci vidljivi, ali bez izmene i novog unosa.
178. **Upgrade na Pro.** Plaća se proporcionalna doplata do isteka aktivne Basic licence.
179. **Obnova Pro-a.** Plaća se jedinstvena puna godišnja cena Pro paketa, koja uključuje Basic.
180. **Downgrade Pro → Basic.** Moguć pri obnovi; Pro podaci ostaju vidljivi i zaključani.
181. **Prestanak Enterprise ugovora.** Finansirani Gazdinstvo nalozi dobijaju 30 dana za samostalnu obnovu, zatim read-only.
182. **Tih 30 dana.** Pristup je read-only; puna funkcionalnost se vraća nakon samostalne uplate.
183. **Istorija saradnje.** Otkupi, dokumenta, ambalaža i saldo ostaju trajno dostupni proizvođaču za pregled jer su i njegovi lični/poslovni podaci.
184. **Izvoz podataka.** Proizvođač može samostalno izvesti kompletnu istoriju u standardnim formatima.
185. **Brisanje naloga.** Brišu se samostalni podaci proizvođača i deaktivira nalog; zajednički poslovni dokumenti ostaju kod hladnjače jer je ona strana u dokumentima i ima pravo/obavezu čuvanja.
186. **Nalog preko hladnjače.** Hladnjača priprema nalog, a proizvođač ga pri prvom otvaranju aktivira i prihvata uslove.
187. **Podaci koje hladnjača vidi.** Samo podatke svog poslovnog odnosa sa proizvođačem.
188. **Dobrovoljno deljenje dodatnih podataka.** Proizvođač može posebno odobriti deljenje plana proizvodnje, tretmana, očekivanog prinosa i drugih grupa podataka.
189. **Povlačenje saglasnosti.** Detaljna pravna razrada je ostavljena otvorenom; trenutno važi da se dodatni podaci dele samo uz jasnu saglasnost.
190. **Validacija i marketing.** Širenje i validacija tržišta idu paralelno; ne čeka se unapred dokazana retencija.
191. **Poruka Gazdinstva.** Preko Enterprise-a naglasak je i na saradnji sa hladnjačom. Samostalno: vođenje proizvodnje, troškovi, profitabilnost, Smart Dosage, weather alerts i ostale napredne funkcije.
192. **Brend Gazdinstva.** Jasno **AgriX Gazdinstvo**, uz prikaz povezane hladnjače; nije white-label.
193. **Basic ograničenja.** Basic je stvarno upotrebljiv bez veštačkih limita broja parcela, evidencija, dokumenata ili količine podataka.
194. **Sajt Gazdinstva.** Trenutno posebna sekcija glavnog sajta; moguć zaseban sajt kasnije. Gazdinstvo mora biti opisano i na stranicama za hladnjače zbog transparentnosti, pregleda stanja i GGAP-a.
195. **Kada izdvojiti sajt.** Kada količina sadržaja i različita ciljna grupa počnu da narušavaju jasnoću glavnog sajta.

---

## 5. AgriX Savetnik

196. **Ciljni direktni korisnici.** Komercijalni proizvođači i poljoprivredni savetnici/konsultanti; veoma mala gazdinstva nisu primarni platioci.
197. **Model Savetnika.** Savetnik kupuje poseban profesionalni nalog i iz jednog interfejsa upravlja većim brojem gazdinstava.
198. **Naplata.** Prema broju aktivnih gazdinstava. *`Superseded` 27.07.2026. odlukom 420 — osnovica plus iznos po gazdinstvu preko 10.*
199. **Naziv proizvoda.** Poseban proizvod: **AgriX Savetnik**.
200. **Licence gazdinstava.** Cena Savetnika pokriva aktivna gazdinstva; ona ne plaćaju zaseban Pro.
201. **Pristup proizvođača.** Proizvođač zadržava sopstveni nalog i pristup podacima.
202. **Ciljni kupci.** Samostalni savetnici/agronomi i savetodavne firme; dugoročno i timski rad.
203. **Prioritet razvoja.** Osnovna verzija do sezone 2027, bez usporavanja Enterprise proizvodnog sistema.
204. **Prva verzija.** Jedan savetnik vodi više gazdinstava; timovi i raspodela dolaze kasnije.
205. **Veza sa Enterprise-om.** Isto gazdinstvo može biti povezano i sa Savetnikom i sa jednom ili više hladnjača.
206. **Interne agronomske službe.** Veće poljoprivredne firme sa sopstvenim agronomima su ravnopravna ciljna grupa.
207. **Ista tarifa.** Isti model cene po aktivnom gazdinstvu za nezavisne savetnike i interne agronomske službe. *`Superseded` 27.07.2026. odlukom 419 — dve tarife, standalone i Enterprise.*
208. **Period naplate.** Samo godišnja pretplata.
209. **Probni period.** Besplatnih 30 dana, uz ograničen broj gazdinstava.
210. **Limit probe.** Do 10 aktivnih gazdinstava.
211. **Odnos prema Pro-u.** Gazdinstvo Pro ostaje isti proizvod. Savetnik dobija dodatne planerske i kontrolne funkcije, a planovi i nalozi stižu u Pro naloge proizvođača.
212. **Vrsta sadržaja savetnika.** Savetnik bira da li šalje obavezujući radni nalog ili neobaveznu preporuku.
213. **Praćenje izvršenja.** Automatski vidi status, kašnjenja, utrošene količine i odstupanja.
214. **Odstupanje proizvođača.** Proizvođač može evidentirati odstupanje i razlog; savetnik dobija upozorenje.
215. **Enterprise klijent sa agronomima.** Posebno kupuje AgriX Savetnik prema broju aktivnih gazdinstava.
216. **GGAP granica.** Sve neophodno za GGAP je u GGAP modulu. Agrosaveti koji prelaze GGAP minimum pripadaju AgriX Savetniku.
217. **Lansiranje.** Javno se nudi čim osnovna verzija bude stabilna; zatvoreni pilot nije obavezan.
218. **Prodajni kanal.** Primarno direktno obraćanje agronomima, savetnicima i savetodavnim firmama.
219. **Savetnik kao partner.** Može formalno prodavati/preporučivati Gazdinstvo uz proviziju.
220. **Gazdinstvo provizija.** Provizija za prvu prodaju i svaku godišnju obnovu dok je korisnik aktivan.
221. **Obračun provizije.** Fiksni iznos po aktivaciji i obnovi, ne procenat pretplate.
222. **Obuhvat provizije.** I Gazdinstvo i Enterprise preporuke, uz različite fiksne iznose.
223. **Enterprise provizija.** Samo jednokratno pri prvom zaključenju ugovora; nema provizije za obnove.
224. **Internet prezentacija.** Posebna stranica unutar glavnog AgriX sajta.
225. **Javna cena.** Ne objavljuje se; individualna ponuda prema broju aktivnih gazdinstava.
226. **Samostalna proba.** Svaki savetnik može sam da se registruje i odmah pokrene probu.
227. **Posle probe.** Nalog prelazi u read-only do prihvatanja ponude i plaćanja.
228. **Prekid saradnje.** Proizvođač zadržava sve svoje podatke, planove, preporuke i istoriju; savetnik odmah gubi pristup.
229. **Dugoročna uloga.** Trenutno softverski alat; dugoročno platforma za pronalaženje, ugovaranje i plaćanje savetnika.
230. **Marketplace timing.** Ne razvijati pre završetka sezone 2027.
231. **Prioriteti posle 2027.** Redosled GGAP-a, marketplace-a i multi-Enterprise arhitekture trenutno je ostavljen otvorenim.

---

## 6. Hladnjača/Proizvodnja i buduće vertikale

232. **Aktivna prodaja.** Proizvodni dodatak se prodaje čim osnovni tok bude stabilan; ne čeka se ceo roadmap.
233. **Postojeće funkcije.** Palete sveže i prerađene robe već su u produkciji. Modul se prodaje postojećim i novim klijentima, uz početni prioritet postojećima zbog lakše validacije.
234. **Red razvoja proizvodnje.** Prvo radni nalozi, norme, ambalaža, otpad i prinos; zatim integracije sa vagama/opremom; zatim kapaciteti linija, smene i učinak radnika.
235. **Buduće funkcije u prodaji.** Klijent kupuje samo ono što postoji na dan prodaje. Roadmap nije ugovorno obećanje niti osnov prodaje.
236. **Današnje pozicioniranje modula.** Sistem za evidenciju prerade, paleta sveže i prerađene robe, zaliha i sledljivosti, uz precizno navođenje postojećih funkcija i napomenu da je dalji razvoj aktivan.
237. **Primarna ciljna grupa.** Hladnjače sa prijemom, klasiranjem, zamrzavanjem, pakovanjem i paletama.
238. **Dugoročna širina.** Mogući su drugi agro-prehrambeni prerađivači samo kada se prirodno uklapaju u AgriX arhitekturu; ne praviti generički proizvodni ERP.
239. **Vertikalni paketi.** Zajedničko proizvodno jezgro, ali posebni komercijalni paketi i prezentacija po važnom segmentu.
240. **Cena vertikala.** Svaki paket ima sopstveni scope, pozicioniranje i cenu.
241. **Kada razviti vertikalu.** Prezentacija može nastati unapred; ozbiljan razvoj tek uz konkretnog kupca ili jasno potvrđenu tražnju.
242. **Mogući red posle 2027.** Žitarice/silosi/mlinovi; duvan; sušare/prerada voća i povrća; vinarije. Nije sadašnji prioritet i ne razrađuje se pre 2027.
243. **Početna naknada za proizvodni modul.** Dok se modul razvija, standardizuje i proverava kod prvog stvarnog klijenta, nema posebne godišnje naknade.
244. **Početak godišnje naknade.** Nakon standardizacije i uspešne provere kod jednog klijenta počinje godišnja naplata modula.
245. **Uvođenje modula.** Uvođenje pojedinačnih modula postojećem klijentu uvek je besplatno. Samo početni onboarding celog AgriX sistema kod novog klijenta može biti naplaćen.
246. **Glavni poslovni cilj do sezone 2027.** Ravnoteža razvoja proizvodnje i rasta klijenata; prednost funkcijama koje neposredno omogućavaju konkretan novi ugovor i ostaju deo zajedničkog proizvoda.
247. **Reference do 2027.** Primarno hladnjače za voće i povrće, radi standardizacije proizvoda, implementacije, prodaje i podrške.
248. **Geografski fokus.** Aktivna prodaja samo u Srbiji do sezone 2027.
249. **Cilj Enterprise klijenata.** 10–20 aktivnih pravnih lica do sezone 2027.
250. **Udeo proizvodnog modula.** Očekivanje je da će ga koristiti više od 80% Enterprise klijenata.
251. **Mesto u prezentaciji.** Hladnjača/Proizvodnja je standardni deo svake početne prezentacije i ponude za hladnjače; klijent može da ga ne kupi.

---

## 7. Prodaja, demonstracija i reference

252. **Tok demonstracije.** Počinje konkretnim problemom klijenta, a završava prikazom celog toka kroz firmu: otkup, prerada, palete, zalihe, dokumenti i upravljački pregled.
253. **Priprema demo-a.** Pre demonstracije održava se kratak razgovor o procesima, problemima i prioritetima klijenta.
254. **Kvalifikacija.** Posebno prilagođen demo radi se samo za lead sa stvarnom potrebom, odgovarajućim profilom i pristupom donosiocu odluke.
255. **Enterprise proba.** Ne postoji probni period niti produkcioni pilot. Postoji samo demo verzija sa dummy podacima.
256. **Samostalni demo pristup.** Kvalifikovani lead posle vođenog demo-a može dobiti vremenski ograničen pristup dummy demo sistemu.
257. **Standardizovan demo.** Samostalna demo instanca ima jedan standardni scenario i iste dummy podatke za sve.
258. **Obim demo-a.** Prikazuje ceo AgriX ekosistem i sve dostupne module; ponuda jasno odvaja kupljeno od opcionog.
259. **Razvojne funkcije u demo-u.** Funkcionalni prototipovi mogu biti prikazani, ali moraju biti jasno označeni kao razvojni i nedostupni za ugovaranje.
260. **Javne reference.** AgriX može javno navesti klijenta kao referencu ako ugovorom to nije izričito zabranjeno.

---

## 8. Posebne GGAP odluke iz ranijeg dela sesije

- Gazdinstvo Basic je primarno veza sa hladnjačom: kartica/saldo, otkupi, ambalaža, dokumenti, obaveštenja, GIS i osnovni prikaz parcela, osnovne vremenske funkcije; ostalo je preview/zaključano/ograničeno.
- Gazdinstvo Pro obuhvata napredno vođenje proizvodnje, tretmane i karencu, agrohemiju, doziranje, troškove i profit, radove i mehanizaciju, napredno vreme, prognoze i analitiku.
- GGAP nije deo Pro-a, već poseban Enterprise dodatak koji kupuje hladnjača za mrežu kooperanata.
- Korisnik uključen u GGAP dobija sve Gazdinstvo funkcije potrebne za usklađenost bez dodatne Pro naknade samo zbog GGAP-a.
- GGAP ima jednu fiksnu godišnju cenu po pravnom licu i pokriva sve njegove GGAP kooperante; nema per-user tier-ova.
- Softverska cena obuhvata platformu, tehničku podršku i compliance workflow. Stručno savetovanje, priprema dokumentacije, pregled, audit podrška i konsulting naplaćuju se odvojeno.
- Trenutno se nudi softver; zatim se razvija mreža eksternih stručnjaka; dugoročno je moguća interna konsultantska ekipa.
- Eksterni konsultant može biti direktno angažovan od klijenta ili podugovoren kroz objedinjenu AgriX uslugu.
- Odgovornost i eventualna garancija sertifikacije određuju se ugovorom i nivoom kontrole. Softver sam po sebi nikada ne garantuje sertifikat.
- Redovna prodaja GGAP modula počinje tek posle validacije sadržaja od kompetentnog konsultanta i najmanje jednog uspešnog realnog projekta.
- Do sezone 2027. GGAP je ograničen na konceptualnu pripremu i stručnu validaciju; ozbiljan razvoj dolazi posle glavnog proizvodnog sistema i stabilizacije Enterprise-a.

---

## 9. Otvorena pitanja

1. Precizna pravna pravila za povlačenje saglasnosti proizvođača za dodatno deljenje podataka.
2. Konačan redosled velikih post-2027 inicijativa: GGAP, marketplace Savetnika i multi-Enterprise arhitektura.
3. Konkretni cenovnici i apsolutni iznosi za Enterprise, Mobile, dodatne stanice, module, Gazdinstvo i Savetnik.
4. Tačan trenutak prelaska sa besplatnog na plaćeni početni onboarding novih Enterprise klijenata.
5. Formalni uslovi partnerskog programa, provizije i pravila atribucije lead-a.

---

## 10. Pravilo tumačenja

- Kasnija odluka ima prednost nad ranijom kada postoji konflikt.
- Funkcije koje su već implementirane i produkciono potvrđene ne treba u dokumentima predstavljati kao buduće.
- Roadmap nije prodajno obećanje niti ugovorna obaveza bez posebnog pisanog ugovora.
- AgriX ostaje zajednički proizvod bez trajnih klijentskih forkova.

---

## 11. Dopuna Q&A sesije — odluke 261–321

**Datum dopune:** 26.07.2026.  
**Obuhvat:** reference, brend, Enterprise sajt i paketi, onboarding, upravljanje razvojem i Multi-Enterprise vlasništvo podataka.

261. **Najvrednija referenca.** Vredne su i javno prepoznatljivo ime/logo klijenta i dokazivi poslovni rezultati. Idealna referenca kombinuje oba, ali je svaka od te dve vrste i samostalno korisna.
262. **Pokazatelji studije slučaja.** Mere se ušteda vremena i administracije, smanjenje grešaka, kontrola robe i sledljivost, kao i kvalitet upravljačkog pregleda. Za svakog klijenta ističu se pokazatelji najvažniji za njegov proces.
263. **Priprema studije slučaja.** AgriX priprema nacrt na osnovu podataka i pokazatelja iz sistema, a klijent potvrđuje tačnost pre objave.
264. **Vreme objave studije slučaja.** Objavljuje se čim postoje dovoljno jasni i merljivi rezultati; ne čeka se nužno završetak cele sezone.
265. **Format reference.** Najvrednija referenca kombinuje pisanu studiju slučaja sa konkretnim pokazateljima i kratku video izjavu vlasnika ili menadžera.
266. **Prikaz rezultata.** Koriste se precizne brojke kada nisu poslovno osetljive i klijent ih odobri, a procenti i relativna poboljšanja kada apsolutni podaci treba da ostanu poverljivi.
267. **Bez podsticaja za referencu.** Za javnu referencu se ne nude finansijske pogodnosti, besplatni moduli niti produženje licence. Referenca treba da bude dobrovoljna posledica zadovoljstva i stvarnih rezultata.
268. **Javno lice proizvoda.** Autoritet i poverenje grade se oko brenda AgriX, proizvoda i rezultata, a ne oko osnivača kao javnog lica.
269. **Arhitektura brenda.** AgriX je krovni brend sa proizvodima **AgriX Enterprise**, **AgriX Gazdinstvo** i **AgriX Savetnik**. Proizvodi ostaju deo jedinstvenog ekosistema.
270. **Enterprise, Desktop i Mobile.** AgriX Enterprise je naziv glavnog B2B proizvoda, dok su AgriX Desktop i AgriX Mobile njegovi komercijalni paketi.
271. **Glavna poruka Enterprise-a.** Suština proizvoda je kompletan operativni sistem za otkup, preradu, skladište i sledljivost. Prodajni ulaz može biti jednostavniji — „softver za otkup i hladnjače“. „Digitalna platforma za kompletno poslovanje agroindustrijske firme“ ostaje približno desetogodišnji North Star.
272. **Naslov i SEO poruka.** Za oglase, SEO i prvi kontakt koristi se „Softver za otkup i hladnjače“. Na glavnoj prodajnoj stranici koristi se „Kompletan operativni sistem za otkup i hladnjače“, uz prikaz šireg toka.
273. **Poziv na akciju.** Primarni CTA je „Zakažite demonstraciju“, a sekundarni „Zatražite ponudu“.
274. **Forma za demonstraciju.** Forma je kvalifikaciona i prikuplja tip firme, broj stanica, vrste robe, postojeći softver i glavni poslovni problem.
275. **Postupanje sa prijavom.** AgriX prvo pregleda prijavu. Jasno kvalifikovan kontakt odmah zakazuje demonstraciju, a kod nejasne prijave prvo sledi kratak kvalifikacioni razgovor.
276. **Struktura Enterprise sajta.** Postoji glavna prodajna stranica i posebne funkcionalne stranice za otkup, hladnjaču/proizvodnju, skladište i sledljivost.
277. **Prioritet funkcionalnih stranica.** Otkup i terenski rad, hladnjača/proizvodnja i skladište/palete/sledljivost razvijaju se paralelno. Blagu prednost ima „Otkup i terenski rad“ zbog neposredne uštede vremena i QR povezivanja proizvođača sa automatskim unosom podataka u dokumente.
278. **Poruka stranice Otkup.** QR kod je konkretan dokaz brzine, a jednokratan unos na mestu nastanka šire obećanje sistema.
279. **Poruka stranice Hladnjača/Proizvodnja.** Centralno obećanje je potpuna sledljivost od prijema sirovine do gotove palete; poslovna vrednost je kontrola prerade, utroška, otpada, prinosa i zaliha.
280. **Poruka stranice Skladište/Palete/Sledljivost.** Kombinuju se trenutna kontrola lokacije i sadržaja svake palete sa potpunom istorijom porekla i kretanja robe od proizvođača do kupca.
281. **Prikaz celine sistema.** Sajt prvo prikazuje povezani tok kroz firmu — otkup, prijem, prerada, palete, skladište, prodaja i dokumenti — a zatim pakete i module.
282. **Glavni dokaz proizvoda.** Dijagram objašnjava kompletan tok robe i podataka, dok stvarni snimci ekrana i kratki video-prikazi dokazuju da funkcije postoje i rade.
283. **Prvi vizuelni sloj.** Prvo se prikazuje jednostavan dijagram kompletnog toka, a zatim stvarni video rada sistema. Već postoji mobilni JSX video-demo prolaska kroz ceo tok.
284. **Mobile i Desktop u videu.** Mobilni video odmah prati prikaz centralnog Desktop sistema, da se ne stvori utisak da je AgriX samo mobilna aplikacija.
285. **Prvi Desktop dokaz.** Najpre se pokazuje da se podatak unet na terenu pojavljuje u centrali i nastavlja kroz dokumente, zalihe i sledljivost; zatim se prikazuje širi Desktop sistem, uključujući dashboarde, izveštaje, SEF, Banku i druge module.
286. **Desktop-only na sajtu.** AgriX Desktop se jasno predstavlja kao potpuno validan, kompletan i povoljniji paket za firme koje ne žele terenski Mobile rad.
287. **Poređenje paketa.** Desktop i Mobile prikazuju se kao dve ravnopravne kolone. Desktop je kompletno rešenje, a Mobile uključuje Desktop i dodaje terenski rad.
288. **Sadržaj Mobile paketa.** Mobile kolona navodi: sve iz Desktop paketa, Mobile Otkupac, Mobile Vozač, sinhronizaciju u realnom vremenu i „Otkup uživo“, gde se najnoviji otkupi prikazuju na vrhu.
289. **Uloga Otkupa uživo.** Korisna je i atraktivna funkcija, ali nije centralni razlog kupovine. Glavna vrednost Mobile-a ostaje unos na mestu nastanka, automatski tok podataka i ušteda vremena.
290. **Management Mobile u oba paketa.** Upravljački mobilni pregled ističe se i uz Desktop i uz Mobile kao zajednička funkcija.
291. **Obim Management Mobile-a.** U Desktop paketu omogućava preglede po kupcu, stanici i kooperantu, zbirno ili pojedinačno, kao i fakture i magacin. U Mobile paketu koristi punu snagu real-time ekosistema. „Otkup uživo“ postoji samo uz Mobile, a Dispatch je dodatni modul koji takođe zahteva Mobile.
292. **Javni nazivi mobilnih uloga.** Koriste se nazivi **Mobile Otkupac**, **Mobile Vozač** i **Mobile MGMT**. Tehnički termin PWA ne koristi se u javnoj i prodajnoj komunikaciji.
293. **Dostupnost Dispatch-a.** Dispatch je funkcionalno dostupan isključivo uz AgriX Mobile. Njegovo mesto na sajtu dodatno je precizirano odlukom 294.
294. **Sekcija dodatnih modula.** Posle poređenja paketa prikazuje se posebna sekcija: Hladnjača/Proizvodnja, SEF, Banka i Dispatch. Prva tri mogu uz Desktop i Mobile; Dispatch se jasno označava kao Mobile-only.
295. **Redosled dodatnih modula.** Redosled se prilagođava stranici i ciljnoj grupi. Na glavnoj Enterprise stranici Hladnjača/Proizvodnja ima prednost.
296. **Bez posebnih stranica modula za sada.** Dodatni moduli ostaju sekcije unutar Enterprise sajta, različite dužine prema značaju.
297. **Prikaz cena.** U kolonama paketa prikazuje se početna cena „od … godišnje“, a detaljni raspon i logika obračuna niže u posebnoj sekciji.
298. **Obuhvat početne cene.** Početna cena obuhvata jedno pravno lice i do pet aktivnih otkupnih stanica. Dodatne stanice imaju javno navedenu godišnju cenu.
299. **Odnos prema računovodstvenom ERP-u.** Ne ističe se rano na prodajnoj stranici. Objašnjava se u FAQ-u i tokom demonstracije: AgriX vodi operativni tok, ali ne zamenjuje računovodstveni ERP.
300. **Prvo FAQ pitanje.** Najistaknutije pitanje je „Da li AgriX radi bez Mobile paketa?“ Odgovor mora potvrditi da je Desktop samostalno i kompletno rešenje.
301. **Drugo FAQ pitanje.** „Koliko traje uvođenje sistema?“
302. **Odgovor o trajanju uvođenja.** Objavljuje se tipičan raspon standardnog uvođenja, a svaki klijent nakon analize dobija poseban plan i dogovoreni datum početka rada.
303. **Uslovi uspešnog uvođenja.** Potrebni su odgovorna osoba kod klijenta, dovoljno vremena pre sezone i dovoljno sređeni matični podaci.
304. **Hijerarhija uslova.** Odgovorna osoba → dovoljno vremena pre sezone → sređeni matični podaci.
305. **Klijent neposredno pred sezonu.** Standardni paket može brzo da krene jer onboarding tipično traje oko pola dana. Pre starta se ne radi custom razvoj.
306. **Minimalni obim pred sezonu.** Nema posebnog „skraćenog“ proizvoda: klijent kreće sa standardnim paketom, a nestandardni zahtevi se odlažu za kasniju fazu. Ova odluka zajedno sa 305 zamenjuje raniju ideju ubrzanog custom uvođenja.
307. **Custom razvoj tokom sezone.** Prihvata se samo kada neposredno omogućava konkretan novi ugovor ili rešava ozbiljan operativni problem, bez ugrožavanja stabilnosti i podrške postojećim klijentima.
308. **Razvoj i release tokom sezone.** Razvoj se nastavlja na odvojenim granama. Svim produkcionim klijentima direktno se puštaju samo bug ispravke, bezbednosne/stabilnosne izmene i kritične funkcije; ostalo ide kroz proveru i kontrolisan rollout.
309. **Postsezonski razgovor.** Sa svakim Enterprise klijentom održava se strukturisan pregled rezultata, problema, zahteva, potrebnih modula i pripreme za narednu sezonu.
310. **Komercijalna uloga pregleda.** Postsezonski razgovor je istovremeno godišnji poslovni pregled i osnova za obnovu, proširenje paketa i plan sledeće sezone.
311. **Pisani ishod pregleda.** Klijent dobija rezime rezultata, problema, dogovorenih obaveza, preporučenih modula i priprema za narednu sezonu.
312. **Jedinstveni backlog.** Svi zahtevi se vode zajedno i klasifikuju kao bug, standardno poboljšanje, strateška funkcija ili plaćeni custom razvoj.
313. **Periodični pregled backloga.** Prioritete formalno pregledaju osnivač, razvoj i podrška prema uticaju na klijente, stabilnost, prodaju, kapacitet i strateški pravac.
314. **Konačna odluka.** Kada nema saglasnosti, osnivač donosi konačnu odluku, posebno za arhitekturu, pozicioniranje i dugoročni pravac.
315. **Roadmap prema klijentima.** Roadmap se ne objavljuje. Budući razvoj se pominje individualno kada je relevantan, bez formalne obaveze i obećanja rokova.
316. **Bez glasanja klijenata.** Klijenti mogu predlagati funkcije, ali AgriX sam određuje prioritete; nema javnog glasanja.
317. **Prioritet reinvestiranja do sezone 2027.** Prvo podrška i onboarding, zatim prodaja i marketing. Razvoj proizvoda nastavlja se svojim planiranim tempom. Raniji princip ostaje: ne uzima se investitor samo radi novca; strateški investitor ima smisla samo ako donosi tržišni pristup, prodajnu mrežu ili industrijsku prednost.
318. **Izričita saglasnost za javnu referencu — ispravka odluke 260.** Ime, logo, studija slučaja, fotografije i video klijenta mogu se javno koristiti samo uz njegovu izričitu saglasnost. Sama činjenica da ugovor to ne zabranjuje nije dovoljna.
319. **Vlasništvo u Multi-Enterprise modelu.** Podaci pripadaju onome ko ih je stvorio. Proizvođač je vlasnik samostalnih podataka Gazdinstva; svaka hladnjača vodi samo podatke svog poslovnog odnosa. Hladnjače međusobno ne vide podatke. Proizvođač u Gazdinstvu vidi odvojene odnose sa svakom hladnjačom.
320. **Master ličnih podataka proizvođača.** Kada proizvođač aktivno koristi Gazdinstvo, ono je master za njegove zajedničke lične/matične podatke. Proizvođač može da ih menja, ali svaka firma mora da odobri primenu promene u svom Enterprise sistemu. Interne šifre, poslovni podaci, finansije i istorijski dokumenti ostaju u nadležnosti svake firme. Ako proizvođač nema aktivno Gazdinstvo, firma vodi svoje matične podatke.
321. **Samostalna vrednost Gazdinstva.** AgriX Gazdinstvo mora biti kompletan i vredan proizvod bez ijedne povezane AgriX hladnjače. Enterprise povezivanje prvenstveno donosi dodatnu korist hladnjači; pogodnosti za proizvođača nisu uslov osnovne vrednosti Gazdinstva.

---

## 12. Arhitektonski audit — vlasništvo, workflow i poslovna pravila

A1. **Istorija kada hladnjača napusti AgriX.** Proizvođač u Gazdinstvu zadržava kompletnu istoriju saradnje koja se na njega odnosi, uključujući otkupe, dokumente i analitiku. Hladnjača izvozi svoje poslovne podatke i može nastaviti u drugom sistemu.
A2. **Nepromjenjiva poslovna istorija.** Važeći istorijski dokumenti se ne menjaju direktno. Ispravke se rade kroz storno, korekciju ili novi dokument, uz potpun audit trail.
A3. **Enterprise matični podaci.** Matični podaci Enterprise-a uređuju se kroz Desktop. Mobile aplikacije ih koriste i preuzimaju, ali ih ne menjaju.
A4. **Workflow odobravanja ličnih podataka.** Gazdinstvo je master zajedničkih ličnih podataka aktivnog proizvođača, ali promena ne postaje automatski važeća u Enterprise-u. Svaka firma prihvata, odbija ili odlaže primenu.
A5. **Plan nije otkupni dokument.** Očekivana količina, plan proizvodnje i najava nikada ne smeju biti osnova za izradu otkupnog lista. Otkupni list nastaje iz stvarno izmerenih i utvrđenih podataka.
A6. **Plansko korišćenje najava.** Najave se mogu i treba da koriste za planiranje vozila, ruta, prijema, radnika, smena, komora, proizvodnje i drugih kapaciteta.
A7. **Pouzdanost najava.** Čuvaju se prijavljena količina, stvarna količina, odstupanje i tačnost po proizvođaču, kulturi i periodu. To služi logistici i budućim ML modelima; razvoj je planiran posle 2028.
A8. **Preporuke i hard business rules.** Poslovne preporuke upozoravaju, dok bezbednosna i regulatorna pravila mogu blokirati nastavak procesa dok se uslov ne ispuni ili ne odobri dozvoljeni izuzetak.
A9. **Kontekstualno odobravanje izuzetka.** Svako pravilo definiše uloge koje smeju da odobre izuzetak. Beleže se korisnik, vreme, razlog, pravilo i povezani dokument.
A10. **Apsolutna i uslovna pravila.** Neka pravila nemaju override; druga dozvoljavaju kontrolisani override uz odgovarajuću ulogu, obrazloženje i audit.
A11. **Product-driven poslovna pravila.** Invarijante su ugrađene u proizvod. Administrator klijenta može aktivirati/deaktivirati konfigurabilna pravila i menjati dozvoljene parametre. Nova pravila i promena njihove logike dolaze isključivo kroz razvoj AgriX-a.
A12. **Verzionisano ponašanje.** Dokument i obrada tumače se prema verziji proizvoda i pravila važećoj u trenutku nastanka. BUILD_VERSION/Git SHA identifikuju build koji je proizveo rezultat.

### Osnovni arhitektonski principi izvedeni iz audita

1. **Business Ownership.** Podatkom upravlja strana kojoj on poslovno pripada; međusistemske promene prolaze kroz definisan workflow i odobrenje.
2. **Immutable Business History.** Istorija se ne prepisuje, već koriguje novim poslovnim događajem.
3. **Product-Driven Business Rules.** Klijent konfiguriše dozvoljene parametre, ali ne menja poslovnu logiku proizvoda.
4. **Planning ≠ Execution ≠ Legal.** Planovi, izvršenje i pravno/računovodstveni dokumenti ostaju odvojeni slojevi.
5. **Versioned Behavior.** Rezultat je vezan za konkretnu verziju sistema i pravila.

### Strateška ML napomena

Dugoročna prednost nije pojedinačan model, već kontinuitet i povezanost podataka kroz Gazdinstvo → Otkup → Hladnjaču → Proizvodnju → Logistiku. U skup ulaze Agromero i drugi agrometeorološki izvori, primenjena agrohemija i đubrenje, tretmani, radovi, vreme, prinos, kvalitet, randman i pouzdanost planiranja. ML se razvija tek kada postoji više sezona kvalitetnih podataka.

---

## 13. Monitoring, incidenti i operabilnost

M1. **Vidljivost monitoringa.** Za sada je Monitoring isključivo interni alat AgriX podrške. Klijent se obaveštava kada treba da reaguje ili kada značajan incident utiče na njegov rad/podatke.
M2. **Kontrolisani self-healing.** Automatski se pokreću samo unapred definisane, bezbedne i idempotentne recovery procedure, uz audit svakog pokušaja i rezultata.
M3. **Centralna matrica incidenata.** Svaki tip incidenta ima ozbiljnost, kanal, odgovornu ulogu, maksimalno vreme reakcije, dozvoljeni recovery i put eskalacije.
M4. **SLA od detekcije.** Interni SLA sat počinje pouzdanom automatskom detekcijom, bez čekanja prijave klijenta.
M5. **Agregirana telemetrija.** Anonimizovani i agregirani monitoring podaci svih klijenata mogu se koristiti za otkrivanje sistemskih problema, merenje pouzdanosti verzija i unapređenje recovery procedura.
M6. **Retention telemetrije.** Detaljni operativni događaji čuvaju se ograničeno; dugoročno ostaju agregati, trendovi i kritični audit događaji.
M7. **Ručna kontrola rollouta i rollbacka.** Monitoring šalje kritični alert, ali zaustavljanje rollouta i rollback pokreće AgriX ručno nakon procene.
M8. **Obaveštavanje prema ozbiljnosti.** Manji potpuno rešeni incidenti mogu ostati interni. Klijent se obaveštava o značajnom uticaju na dostupnost, obradu, tačnost ili integritet, uz razumljiv sažetak uzroka i rešenja.
M9. **Obavezni postmortem.** Svaki značajan produkcioni incident dobija formalni interni postmortem sa uzrokom, vremenskom linijom, uticajem, oporavkom, merama, vlasnicima i rokovima.
M10. **AuditCritical je odvojen od telemetrije.** To je neizmenjiv, dugoročno čuvan audit trag povezan sa korisnikom, firmom, entitetom/dokumentom, pravilom, vremenom i buildom.
M11. **Pristup klijenta audit zapisu.** Za sada se zapis dostavlja na zahtev, tokom kontrole ili posle značajnog incidenta. Dugoročni cilj je read-only pregled i izvoz audit zapisa sopstvene firme.
M12. **Jedinstveni standard podrške.** Svi Enterprise klijenti koriste istu matricu ozbiljnosti, eskalacije i vremena reakcije; nema posebnih SLA nivoa po klijentu.
M13. **Bez javnog procenta dostupnosti.** AgriX za sada ne ugovara niti javno garantuje procenat dostupnosti. Planirano održavanje se najavljuje, a neplanirani prekidi tretiraju kao incidenti.
M14. **Interni cilj dostupnosti.** Dostupnost ključnih servisa i poslovnih tokova meri se interno radi upravljanja kvalitetom.
M15. **Dvonivojsko merenje.** Glavna mera su stvarni poslovni tokovi, a komponente se mere radi dijagnostike uzroka.
M16. **Spoljni servisi.** Nedostupnost SEF-a, Google servisa, bankarskog API-ja i drugih eksternih sistema ne ulazi u AgriX metriku dostupnosti, ali se posebno prati i klasifikuje.
M17. **Kontrolisana degradacija.** Za svaki tok se definiše da li se zahtev bezbedno stavlja u red i ponavlja ili se proces blokira da ne bi nastalo nevažeće stanje.
M18. **Odvojeni lokalni i integracioni status.** Lokalno uspešna transakcija ostaje završena, dok se spoljna obrada vodi zasebno kao Pending.
M19. **Nastavak dok je Pending.** Bezbedni operativni koraci mogu da se nastave; pravni, finansijski i nepovratni koraci čekaju spoljnu potvrdu.

---

## 14. Bezbednost i izolacija podataka

S1. **Pristup podrške.** Ovlašćeni članovi AgriX podrške mogu imati stalni pristup produkcionim podacima radi podrške i održavanja, ograničen prema ulozi i potpuno auditovan.
S2. **Bez produkcionih podataka u razvoju.** Razvojna i testna okruženja koriste sintetičke ili nepovratno anonimizovane podatke.
S3. **Imenovani nalozi.** Svaki korisnik ima sopstveni nalog; deljeni nalozi nisu dozvoljeni.
S4. **MFA.** Obavezan je za administratore, AgriX podršku i druge privilegovane uloge; nije obavezan za standardne operativne korisnike.
S5. **Izolacija pravnih lica.** Korisnik pristupa samo izričito dodeljenim firmama i uvek radi u jasno označenom kontekstu jedne firme. Zbirni pregled postoji samo kroz posebno odobrenu grupnu ulogu.

---

## 15. ML i AgriX Intelligence

ML1. **Zajedničko treniranje.** Nepovratno anonimizovani podaci različitih klijenata i Gazdinstava mogu se koristiti za zajedničke modele bez otkrivanja identiteta ili poslovnih podataka učesnika.
ML2. **Bez samostalnog poslovnog ovlašćenja.** ML daje procene, preporuke i upozorenja; ne određuje sam cenu, ne odbija proizvođača i ne blokira otkup.
ML3. **Objašnjivost i audit.** Značajna preporuka čuva verziju modela, relevantne ulaze, nivo pouzdanosti i razumljivo objašnjenje glavnih faktora.
ML4. **Jedinstveni modeli.** Razvija se jedan zajednički model po nameni. Kultura, region, gazdinstvo i uslovi su ulazne karakteristike, a ne osnov za zaseban model po klijentu.
ML5. **Komercijalni model.** Napredne ML funkcije prodaju se kroz poseban plaćeni modul **AgriX Intelligence**, tek kada postoje dovoljan kvalitet podataka, više sezona istorije i dokaziva pouzdanost.

---

## 16. Integracije i autoritativni izvori

I1. **Zatvoren integracioni model.** Integracije razvija i kontroliše AgriX. Klijentima i trećim stranama ne daje se opšti API za samostalni upis ili pokretanje poslovnih procesa.
I2. **Standardni konektori.** Za svaki podržani spoljni sistem razvija se jedan konektor koji se konfiguriše po klijentu; ne prave se zasebne integracije za svaku firmu.
I3. **Autoritativni izvor po domenu.** AgriX je autoritativan za otkup, proizvodnju, skladište, logistiku i sledljivost. Spoljni sistem je autoritativan za sopstveni domen, npr. glavnu knjigu.
I4. **Konflikti bez automatskog prepisivanja.** Neslaganje se označava i prosleđuje na korisničku kontrolu; vrednosti se ne prepisuju automatski.
I5. **Naplata integracija.** Osnovni standardizovani uvoz/izvoz je uključen. Puna automatska ili dvosmerna integracija je dodatno plaćeni konektor/modul.

---

## 17. Životni ciklus podataka

D1. **Podaci posle prestanka ugovora.** Neograničeno arhiviranje je dozvoljeno samo uz izričitu saglasnost bivšeg klijenta. Bez saglasnosti primenjuje se brisanje ili posebno dogovoreno arhiviranje, uz zakonske izuzetke.
D2. **Završni izvoz.** Standardni izvoz podataka i dokumenata u definisanim formatima je uključen. Posebno mapiranje i direktna migracija u drugi sistem dodatno se ugovaraju i naplaćuju.
D3. **Potpuno brisanje.** Na zahtev se podaci uklanjaju iz aktivnih sistema i rezervnih kopija, osim zapisa koje izričita zakonska obaveza zahteva da se zadrže.
D4. **Anonimizovani agregati.** Nepovratno anonimizovani podaci koji se više ne mogu povezati sa klijentom, firmom ili osobom mogu se trajno zadržati za statistiku, pouzdanost i ML.
D5. **Izvoz tokom ugovora.** Klijent sam izvozi osnovne podatke i dokumente; kompletan arhivski izvoz svih veza i priloga pruža AgriX podrška na zahtev.

### Potvrđena offline arhitektura

Mobile je offline-first i sinhronizuje čim dobije vezu. Zaštite brojeva dokumenata, stanica i lock mehanizmi sprečavaju konflikte. Desktop nastavlja da radi i kada Mobile ili sinhronizacija privremeno nisu dostupni; kvar jednog kanala ne sme zaustaviti drugi.

---

## 18. Dugoročna platformska arhitektura

P1. **Migracija samo po objektivnim pragovima.** Ne postoji migracija radi tehnologije. Komponente se postepeno zamenjuju kada broj klijenata, podaci, performanse, pouzdanost, konkurentnost ili regulativa to zahtevaju.
P2. **Stabilan poslovni ugovor.** Tokom tehnološke zamene ostaju stabilni poslovni model, identiteti entiteta, istorijski podaci i očekivano ponašanje. Menja se implementacija, ne značenje procesa.
P3. **Poslovna logika kroz kanale.** Trenutno Desktop i Mobile prate jednu kanonsku specifikaciju i moraju davati isti rezultat. Dugoročno se logika premešta na centralni backend, a Desktop i Mobile postaju klijentski kanali.
P4. **Jedan autoritativni izvršilac po domenu.** Tokom migracije svaki domen ima samo jedan sistem koji potvrđuje konačnu promenu; paralelni upis u isti domen nije dozvoljen.
P5. **Postepena migracija klijenata.** Klijenti prelaze pojedinačno ili u grupama, uz kompatibilnost podataka i jasno ograničen kraj podrške stare generacije.

---

## 19. Intelektualna svojina i finansirani razvoj

IP1. **Vlasništvo AgriX-a.** Kod, arhitektura i poslovno rešenje ostaju AgriX-u i kada klijent finansira razvoj; klijent dobija pravo korišćenja.
IP2. **Poverljivo znanje klijenta.** Ugovorom se određuje šta AgriX sme da generalizuje i ponovo koristi, a šta ostaje poverljiva procedura ili poslovna tajna klijenta.
IP3. **Bez izvornog koda.** Klijent ne dobija izvorni kod, posebnu source licencu niti source-code escrow.
IP4. **Pogodnost finansijeru razvoja.** Klijent dobija prioritetnu izradu/prilagođavanje i besplatno korišćenje funkcionalnosti tokom prve godine od produkcijskog puštanja. Posle toga važe standardni uslovi. *`Rewritten` 27.07.2026. odlukom 422 — ista prva besplatna godina kao 367, ne dodatna.*
IP5. **Ekskluzivnost.** Moguća je samo uz poseban, vremenski i funkcionalno ograničen ugovor i znatno višu cenu.

---

## 20. Kontinuitet poslovanja i zavisnost od osnivača

BC1. **Trenutno prihvaćena zavisnost.** U sadašnjoj fazi ključne produkcione operacije, deployment, incidenti i napredna podrška mogu zavisiti od osnivača.
BC2. **Bez vanrednog continuity paketa za sada.** Ne postoji poseban paket pristupa i ovlašćenja za slučaj dugotrajne nedostupnosti osnivača.
BC3. **Okidač za drugu tehničku osobu.** Druga tehnički ovlašćena osoba uvodi se najkasnije pri 15–20 aktivnih firmi ili zapošljavanju prve tehničke osobe — šta nastupi ranije.

---

## 21. Zastarele funkcije, kvalitet i release

L1. **Ukidanje starog toka.** Kada novi proces zameni stari, prethodna funkcionalnost se uklanja nakon prelaznog perioda i migracije aktivnih klijenata. Ne održavaju se trajno paralelni načini za isti proces.

Q1. **Poznati nekritični nedostaci.** Release je dozvoljen ako nedostatak ne ugrožava podatke, zakonsku ispravnost ni ključne tokove, postoji bezbedan workaround i plan otklanjanja.
Q2. **Obavezni release gate-ovi.** Svaki produkcijski release prolazi propisane provere kritičnih tokova, migracija, oporavka, integriteta podataka i monitoringa.
Q3. **Rollout prema riziku.** Velike, kritične i arhitektonski značajne funkcije prvo idu ograničenom broju klijenata; male niskorizične izmene mogu ići svima posle standardnih provera.
Q4. **Hotfix bez preskakanja.** I hitna ispravka prolazi iste release gate-ove, samo prioritetno i ubrzano.

---

## 22. Onboarding i produkcijski početak

ON1. **Početni podaci.** Klijent priprema i potvrđuje poslovnu tačnost šifarnika, proizvođača, stanica, artikala, cena i početnih stanja. AgriX daje obrasce, uvoz, tehničku validaciju i prijavljuje nelogičnosti.
ON2. **Obuka.** Početna implementaciona obuka deo je onboardinga. Kasnije obuke novih zaposlenih i dodatne radionice posebno se ugovaraju i naplaćuju.
ON3. **Zajednički produkcijski sign-off.** AgriX potvrđuje tehničku spremnost, uspešnost uvoza i ključne tokove. Odgovorna osoba klijenta potvrđuje poslovnu tačnost podataka, podešavanja i spremnost zaposlenih. Produkcijski rad počinje tek nakon obe potvrde.
ON4. **Period pojačane podrške nakon puštanja.** Pojačana podrška neposredno nakon produkcijskog početka nije automatski deo standardnog onboardinga; obezbeđuje se samo kada je posebno ugovorena.

---

## 23. Komercijalni model, tržište i portfolio — dodatne odluke

C1. **Trial režim.** AgriX ima trial režim koji omogućava punu funkcionalnost proizvoda. Trial je standardni način probnog korišćenja pre pune komercijalne aktivacije.

C2. **Naplata onboardinga i instalacija.** Instalacija AgriX-a uvek je uključena u cenu. Onboarding se posebno naplaćuje kada zahteva više od najosnovnijeg prikaza toka kroz program.

C3. **Dozvoljeni popusti.** Popusti se mogu odobriti za više pravnih lica, veliki broj otkupnih stanica i druge opravdane specifične situacije. Ne postoji automatsko pravo na popust van tih poslovno obrazloženih slučajeva. *`Deleted` 27.07.2026. odlukom 418 — pregovaračkih i individualnih popusta nema.*

C4. **Promena cena.** Cene pri obnovi i buduće promene cenovnika određuju se diskreciono, u skladu sa poslovnom procenom AgriX-a.

C5. **Prevremeni raskid godišnjeg ugovora.** Klijent može prevremeno raskinuti godišnji ugovor samo kada AgriX učini bitnu povredu svojih ugovornih obaveza.

C6. **Kašnjenje koje izazove klijent.** Kada klijent ne dostavi potrebne podatke, ne odredi odgovornu osobu ili ne pripremi zaposlene, implementacija se pauzira, ali ugovorni period nastavlja da teče.

C7. **Granica standardne podrške.** Čišćenje i ispravljanje podataka klijenta, masovne korekcije nastale greškom korisnika, posebni izveštaji, dolazak na lokaciju, podešavanje računara/mreže/štampača i savetovanje o internim poslovnim procesima predstavljaju dodatno plaćeni rad.

MKT1. **Primarni tržišni segment.** Hladnjače i otkupljivači predstavljaju prioritetnu ciljnu grupu AgriX-a.

MKT2. **Geografski fokus.** Primarno tržište je Srbija. Regionalno širenje nije trenutni prioritet.

MKT3. **Partnerski model.** Partneri mogu imati referral ili posredničku ulogu. Ne uvodi se model ovlašćenih implementacionih partnera niti punih resellera.

MKT4. **Gate za povećanje marketing budžeta.** Odluka o većem ulaganju u marketing zasniva se na kombinovanoj proceni svih ključnih pokazatelja: broja aktivnih Enterprise klijenata, demo-to-contract konverzije, vremena zatvaranja prodaje, troška sticanja klijenta, stope obnove i broja/kvaliteta preporuka i referenci.

MKT5. **Prva nova zaposlena uloga.** Prva nova uloga treba da bude fokusirana na onboarding i podršku.

PRT1. **Granica Gazdinstvo Basic i Pro.** Već je definisana u postojećim proizvodnim odlukama i ne duplira se u ovoj dopuni.

PRT2. **Plaćanje Gazdinstvo Pro.** Proizvođač je primarni korisnik i može da plati Pro direktno ili preko hladnjače. Hladnjača može finansirati Pro za svoje kooperante kada to želi kao deo sopstvenog poslovnog modela.

PRT3. **AgriX Savetnik.** AgriX Savetnik je realan budući proizvod, praktično management sloj koji omogućava istovremeni pregled više gazdinstava, slanje naloga i preporuka gazdinstvima i kontrolu njihovog rada.

PRT4. **GGAP/compliance pozicioniranje.** GGAP je dodatni paket u okviru AgriX Enterprise-a. Aktivacija tog paketa otvara dodatne funkcije u AgriX Gazdinstvu koje su potrebne za GGAP procese i evidencije.

PRT5. **Gate za AgriX Intelligence.** Spremnost za ML/Intelligence određuju kvalitet i kompletnost podataka i stvarni broj sezona potreban da model bude upotrebljiv. Ne postoji jedan univerzalni fiksni broj sezona za sve use-case-ove.

---

## 24. Pravni i podatkovni okvir — dodatne odluke

LEG1. **Formalne uloge u zaštiti podataka.** Uloge rukovaoca i obrađivača ne zaključavaju se unapred. Biće određene uz pravnika nakon izrade kompletne mape tokova podataka za Enterprise, Gazdinstvo, povezivanje sa hladnjačama, sistemske logove, naplatu i agregiranu analitiku.

LEG2. **Brisanje podataka iz rezervnih kopija.** Aktivni podaci brišu se odmah. Rezervne kopije ostaju neizmenjene do isteka definisanog retention perioda, bez obaveze naknadne obrade ili ponovnog brisanja podataka ako se stara kopija privremeno vrati.

LEG3. **Korišćenje anonimizovanih podataka za agregate i ML.** Anonimizovani podaci mogu se koristiti za agregatnu analitiku, AgriX Intelligence i ML samo uz prethodnu izričitu saglasnost klijenta ili proizvođača, čak i kada je anonimizacija nepovratna.

LEG4. **Ograničenje odgovornosti.** Ukupna ugovorna odgovornost AgriX-a prema konkretnom klijentu ograničava se na iznos koji je taj klijent platio AgriX-u tokom prethodnih 12 meseci, uz izuzetke i ograničenja koja se po zakonu ne mogu ugovorom isključiti.

LEG5. **Obaveštavanje o bezbednosnom incidentu.** Klijenti se obaveštavaju bez nepotrebnog odlaganja, u zavisnosti od ozbiljnosti incidenta i potvrđenog uticaja na njihove podatke ili poslovanje. Ne propisuje se jedinstven fiksni rok za sve vrste incidenata.

---

## 25. Dopuna sesije 26.07.2026. — odluke 323–378

**Datum dopune:** 26.07.2026.  
**Obuhvat:** cene i obračun, ugovor i obnova, podrška, onboarding, Gazdinstvo, Savetnik, tržišni cilj, bezbednost i razrešenje četrnaest konflikata iz arhitektonskog audita.  
**Izvor:** `09B_ODLUKE_PO_OBLASTIMA.md` (tematski indeks, verzija 2), uz detalje iz `docs/Sales/AgriX_Cenovnik_2027.pdf`, `docs/Legal/AgriX_Ugovor_o_licenciranju.docx`, `docs/Product/AgriX_Definicija_proizvoda.pdf` i `docs/Finance/AgriX_Finansijski_model.xlsx`.

> Numeracija 323–378 nastavlja niz posle 321; broj **322 se ne koristi**. Gde odluka menja raniju, to je izričito označeno.

### 25.1 Cene i obračun

323. **Aktivna otkupna stanica.** Aktivnom stanicom smatra se svaka stanica na kojoj je u toku ugovorne godine evidentiran najmanje jedan otkupni blok.
324. **Prijava i usklađenje broja stanica.** Stanice se prijavljuju unapred; stvarni broj se utvrđuje po završetku sezone i usklađuje prilikom obnove, bez povraćaja. Menja odluku 120.
325. **Bez plaćanja usred sezone.** Ugovori se sklapaju i naknada plaća pre početka sezone, nikada tokom nje.
326. **Jedinstveni predsezonski datum obnove.** Datum obnove se usklađuje tako da uvek pada pre početka sezone; prvi ugovorni period je srazmeran. Zamenjuje odluku 114.
329. **Valuta i kurs.** Cene su u EUR; plaćanje je u dinarima po srednjem kursu Narodne banke Srbije na dan uplate. Cene ne sadrže PDV.
334. **Zaštita cene postojećih klijenata.** Postojeći klijenti zadržavaju svoje cene za ovu sezonu.
337. **Smanjenje broja stanica.** Smanjenje broja stanica u toku ugovorne godine ne menja cenu ako je stanica u toj godini bila aktivna.
349. **Cene paketa.** AgriX Desktop 500 €, Desktop all-in 1.200 €, AgriX Mobile 1.500 €, Mobile all-in 2.200 € — godišnje, po pravnom licu, do pet aktivnih stanica. Struktura cena definisana je odlukama 414 i 415; sastav all-in paketa odlukom 415.
350. **Cene modula.** SEF, Banka i Dispatch po 200 €; Hladnjača/Proizvodnja 400 € — godišnje, po pravnom licu.
351. **Cena dodatne stanice.** Svaka aktivna stanica preko pet — 50 € godišnje. Ista cena u oba paketa (potvrđuje odluku 123).
352. **Cena GGAP modula.** Od 1.000 € godišnje po pravnom licu; jedna cena pokriva sve GGAP kooperante tog lica.
353. **Satnica za razvoj i migraciju.** 50 € po satu za razvoj po zahtevu i složenu migraciju podataka.
354. **Obuka.** Pet sati implementacione obuke uključeno u onboarding; preko toga 30 € po satu.
355. **Izlazak na teren.** 50 € po izlasku, uvećano za gorivo, vreme puta i vreme rada na lokaciji.
356. **Marža na hardver.** Hardver se prodaje sa oko 100 € marže po stanici.
357. **Hardverska podrška.** 40 € po stanici godišnje, minimum 200 € po pravnom licu. `DECISION` — potvrđeno 27.07.2026. Ne pokriva fizička oštećenja, potrošni materijal ni opremu nabavljenu van AgriX-a.
358. **Dodatna instanca.** Druga instanca dobija −50 % na sve što ta instanca dodatno koristi. Moduli koji se po odluci 135 plaćaju jednom po pravnom licu ne dupliraju se.
367. **Jedno gratis pravilo.** Postoji samo jedan gratis period: prva godina od produkcijskog puštanja funkcionalnosti. Zamenjuje odluke 144 i 149; usklađuje IP4. Razrešava konflikt K-05.
368. **Granica besplatnog uvođenja modula.** Softverski modul i njegova konfiguracija su besplatni; fizički rad na lokaciji i puštanje opreme u rad se naplaćuju. Razrešava konflikt K-04 (odluka 68 protiv odluke 245).

### 25.2 Ugovor, obnova i raskid

338. **Kašnjenje krivicom klijenta.** Kada uvođenje kasni iz razloga na strani klijenta, neiskorišćeni deo perioda prenosi se **jednokratno** u narednu ugovornu godinu. Dopunjuje C6.
361. **Read-only režim posle isteka.** Režim pregleda i izvoza bez unosa je razvojni prioritet sa rokom pre 1. juna 2027. Do isporuke se u ugovoru tretira kao postojeći. Razrešava konflikt K-03 (odluka 117 protiv stvarnog `LicenseBlock` ponašanja).
376. **Izrada ugovora.** Nacrt ugovora piše osnivač; pravnik radi pregled gotovog nacrta, ne izradu od nule.

### 25.3 Podrška i SLA

327. **Kontrola dokumenata.** Klijent je dužan da kontroliše sadržaj dokumenata koje izdaje; AgriX ispravlja potvrđen bug u najkraćem roku i ne odgovara za posledice izostale kontrole.
331. **Definicija sezone.** Sezona se definiše jedinstveno na nivou AgriX-a, ne po klijentu. Zamenjuje odluku 56.
332. **Vikend podrška u sezoni.** Tokom sezone vikend podrška postoji **samo za kritične incidente**. Zamenjuje odluke 54 i 57.
359. **Rok reakcije od jednog sata.** Rok od jednog sata za kritične incidente važi unutar definisanog proširenog vremenskog prozora; van njega je best effort. Menja odluku 50 i razrešava konflikt K-01.

> **Prozor utvrđen 27.07.2026.:** tokom sezone (1. jun — 30. novembar) **svakog dana od 08.00 do 20.00** časova, uključujući vikend; van sezone **radnim danima od 08.00 do 16.00** časova. Upisano u Prilog 2 ugovora. Usklađeno sa odlukama 332 (vikend podrška u sezoni samo za kritične incidente) i 378 (trajanje sezone). Van prozora važi best effort, bez ugovorenog roka.
378. **Trajanje sezone.** Sezona traje od 1. juna do 30. novembra. Precizira odluku 331.

### 25.4 Onboarding i implementacija

362. **Modul i obuka.** Uvođenje modula postojećem klijentu je besplatno uz uključenu obuku; preko toga se naplaćuje. Razrešava konflikt K-12 (odluka 245 protiv ON2). **Precizirano 27.07.2026.:** uključeno je **pet sati ukupno**, i za početni onboarding i za uvođenje modula; preko toga 30 € po satu. Time je odluka 362 usklađena sa odlukama 354 i 365 i nema odvojene kvote sati po modulu.
363. **Skraćeni sign-off.** Za predsezonski start potvrda iz ON3 može se dati u skraćenom obliku, uz izričito prihvatanje rizika od strane klijenta; tada se pre produkcijskog starta ne izvodi razvoj po zahtevu. Razrešava konflikt K-13.
365. **Fiksni uključeni obim.** U cenu je uključeno: instalacija, povezivanje svih komponenti i pet sati obuke. Sve preko toga se naplaćuje. Razrešava konflikt K-11 (C2 protiv odluke 35).

### 25.5 Prodaja i marketing

371. **Trial.** Trial režim postoji i dolazi **posle** vođene demonstracije i kvalifikacije, ne umesto njih. Razrešava konflikt K-09 (C1 protiv odluke 255).
374. **Gate za marketing budžet.** Uvode se tri nivoa potrošnje; svaki ima dvostruki uslov — kanal dokazano konvertuje **i** postoji kapacitet za isporuku — uz pravilo povratka na niži nivo kada uslov prestane da važi. Operacionalizuje MKT4.
377. **Referral provizije.** Nema referral provizija, osim u slučaju Savetnika.
333. **Bez ekskluzivnosti.** Ne ugovara se teritorijalna ni segmentna ekskluzivnost; direktni konkurenti mogu istovremeno biti klijenti.

### 25.6 Tržišni cilj

375. **Scenario rasta C.** Izabran je scenario C. **Precizirano 27.07.2026.: planska vrednost je 14 novih Enterprise klijenata do sezone 2027, ukupno 17 aktivnih firmi.** Raniji zapis kao raspona (12–15 novih, 15–18 ukupno) zamenjen je jednom vrednošću, saglasno finansijskom modelu. Odluka 375 zamenjuje odluku 249 (cilj 10–20 aktivnih pravnih lica do 2027).

> Merodavan izvor je `docs/Finance/AgriX_Finansijski_model.xlsx`, list `Pretpostavke`, red 42. Isti broj mora stajati u `02_STRATEGY.md` §9 i `04_MARKET.md` §9.1. Kolona od 18 klijenata na listu `Kapacitet` je **stress-test scenario**, ne cilj — služi da pokaže kada osnivač postaje usko grlo, i zadržava se kao takva.

### 25.7 AgriX Gazdinstvo

339. **Kanalska cena.** Jedinstvena kanalska cena za sve partnerski posredovane naloge: **10 € Basic, 20 € Pro**. Maloprodajna cena ostaje 19 € Basic i 39 € Pro. Prvih 50 Basic naloga partner dobija bez naknade (odluka 161).
343. **Jedan Pro po proizvođaču.** Proizvođač ima jedan Pro nalog — ko ga prvi aktivira, taj ga plaća; druga strana ne plaća ponovo.
369. **Povlačenje saglasnosti.** Povlačenje saglasnosti proizvođača deluje samo ubuduće; već izdati dokumenti ostaju nepromenjeni. Dopunjuje odluku 189.

### 25.8 AgriX Savetnik

340. **Gazdinstva nisu uključena u cenu.** Gazdinstva u portfelju savetnika nisu uključena u cenu Savetnika; svako drži sopstvenu Pro pretplatu po kanalskoj ceni. Zamenjuje odluku 200.
341. **Cena Savetnika.** Osnovica 150 € godišnje, uključeno do 10 gazdinstava; svako gazdinstvo preko 10 — 15 €. Potvrđeno 27.07.2026. Odluka 419 uvodi drugu, Enterprise tarifu (100 € / 10 €), a odluka 420 definiše model naplate.
342. **Aktivno gazdinstvo.** Aktivnim se smatra gazdinstvo kojem je savetnik u toku godine poslao makar jedan nalog ili preporuku.
344. **Savetnik kao platilac.** Savetnik može platiti Pro u ime proizvođača i ugraditi to u svoju naknadu — posrednička uloga u skladu sa MKT3.
345. **Bez cashbacka u portfelju.** Nema provizije ni cashbacka za gazdinstva u sopstvenom portfelju; podsticaj je sam alat, koji bez Pro naloga ne funkcioniše. Odluka 221 ostaje samo za preporuke van portfelja.
346. **Proba Savetnika.** Probni period obuhvata i Pro za do 10 gazdinstava. Dopunjuje odluku 209.
347. **Objavljivanje cene.** Cena Savetnika se objavljuje — osnovica i cena po gazdinstvu. Zamenjuje odluku 225.
348. **Interne agronomske službe.** Interne agronomske službe plaćaju samo alat kada su njihovi kooperanti već pokriveni partnerskim paketom.

### 25.9 Podaci, bezbednost i platforma

330. **Hosting.** Hosting van Google infrastrukture nije u ponudi do 2028.
364. **Anonimizovani podaci.** Nepovratno anonimizovani podaci ne traže posebnu saglasnost; upotreba se transparentno navodi u ugovoru. Revidira LEG3 i razrešava konflikt K-10.
370. **Restore i brisanje.** Ako se rezervna kopija privremeno vrati u upotrebu, izvršena brisanja se ponovo primenjuju i o tome se sačinjava zabeleška. Razrešava konflikt K-14 uz LEG2.
372. **Numeracija dokumenata.** Prelazak privremenog broja u konačan rešen je u kodu; pravilo treba zapisati u dokumentaciju. Dopunjuje odluku 88, razrešava konflikt K-02.
373. **Podobrađivači.** Jedini podobrađivač je Google. `OPEN`: lokaciju obrade podataka treba verifikovati u Workspace konzoli i zapisati.
366. **Sheets kao backend.** Google Sheets ostaje PWA backend dok se ne dostignu pragovi iz P1. Preformuliše odluku 90 i razrešava konflikt K-07.

### 25.10 Razvoj, release i organizacija

328. **Zakonske izmene.** Prilagođavanje izmenama propisa AgriX radi o svom trošku, nezavisno od broja pogođenih klijenata. Ograničava odluku 24.
360. **Emergency release gate.** Tokom incidenta važi smanjen obavezan skup provera; puna validacija i dokumentacija obavljaju se u roku od 24 sata po stabilizaciji. Menja Q4 i razrešava konflikt K-08.
335. **Odgovor na pitanje kontinuiteta.** Usvojen je standardni odgovor na prodajno pitanje „šta ako vas sutra nema“: podaci su klijentovi i izvoze se u standardnom formatu; Desktop radi lokalno do isteka licence; uvođenje druge tehnički ovlašćene osobe je ugovorna obaveza.
336. **BC3 kao ugovorna obaveza.** Uvođenje druge tehnički ovlašćene osobe najkasnije pri 15–20 aktivnih firmi postaje ugovorna obaveza, ne interni cilj.

### 25.11 Otvorene stavke iz ove dopune

Stanje na dan 27.07.2026.

| # | Stavka | Vezano za | Status |
|---|---|---|---|
| 1 | Pravni pregled ugovora nije obavljen | 376 | otvoreno |
| 2 | Prilog 3 ugovora nije dovršen — čeka mapu tokova i LEG1 | LEG1 | otvoreno |
| 3 | Vremenski prozor za rok od jednog sata | 359 | **zatvoreno 27.07.** — sezona 08–20 svakog dana, van sezone 08–16 radnim danima; upisano u Prilog 2 |
| 4 | Mesto nadležnog suda u ugovoru | član 15 | **zatvoreno 27.07.** — Niš; sedište AgriX-a je Merošina |
| 5 | Cena po gazdinstvu kod Savetnika čeka potvrdu | 341 | **zatvoreno 27.07.** — potvrđeno 150 € osnovica / 15 € po gazdinstvu; uvedena Enterprise tarifa 100 € / 10 € (419) |
| 6 | Hardverska podrška 40 €/stanici | 357 | **zatvoreno 27.07.** — potvrđeno kako stoji u cenovniku i Prilogu 1 |
| 7 | Lokacija podataka u Google Workspace konzoli nije verifikovana | 373 | otvoreno |
| 8 | Redosled post-2027 inicijativa svesno otvoren | 231 | otvoreno svesno |
| 9 | Vertikalni paketi protiv odluke 59 — odloženo do prve vertikale | K-06 | odloženo |
| 10 | Broj uključenih sati obuke u 362 | 362 | **zatvoreno 27.07.** — pet sati ukupno, isto kao 354 i 365 |
| 11 | Raspon protiv tačke u cilju rasta | 375 | **zatvoreno 27.07.** — merodavno je 14 novih / 17 ukupno |
| 12 | `09B_ODLUKE_PO_OBLASTIMA.md` nije obuhvatao 401–408 ni 409–422 | — | **zatvoreno 27.07.** — indeks regenerisan na verziju 3 |
| 13 | Desktop all-in štedi klijentu samo 100 € — slaba bundle poruka | 415 | komercijalno, nije konflikt |

---

## 26. Odluke 401–408 — portfolio, rast i cene

**Datum dopune:** 27.07.2026.  
**Obuhvat:** treći proizvodni stub, status GGAP-a, model rasta, komercijalna spremnost Gazdinstva, cena po stanici, hardverska marža i redosled cena i unit economics-a.

401. **Savetnik je treći stub.** AgriX Savetnik je treći ravnopravan proizvodni stub uz Enterprise i Gazdinstvo. Potvrđuje odluku 269 (krovni brend sa tri proizvoda) i koriguje ranije formulacije u `02_STRATEGY.md` i `07_PRODUCT_PORTFOLIO.md`, gde je treći stub bio GGAP.
402. **GGAP je modul, ne stub.** GGAP je modul u okviru AgriX Enterprise-a i koriste ga isključivo hladnjače koje su već Enterprise klijenti. Aktivacija modula otključava dodatne funkcije u Gazdinstvu (PRT4). Ovom odlukom se **ukida STR-012** (GGAP kao treći proizvodni stub).
403. **Fiksan ciljni broj klijenata.** Rast se planira prema fiksnom ciljnom broju klijenata, ne prema readiness cap-u. Ovom odlukom se **povlači STR-001** (sezonski cap određuje readiness score). Readiness ostaje operativni preduslov kvaliteta isporuke i može zaustaviti pojedinačan onboarding, ali više ne određuje ciljni broj. Aktuelan cilj je scenario C iz odluke 375.
404. **Gazdinstvo je launch ready.** Gazdinstvo prelazi iz statusa `Pilot only` u `Standard offer`. Menja status iz `07_PRODUCT_PORTFOLIO.md` §8 i §11.
405. **GGAP van komercijalne ponude.** GGAP ostaje van komercijalne ponude do validacije. Ne prodaje se kao redovna stavka; nudi se samo kroz kontrolisan pilot uz potvrdu obima.
406. **Jedinstvena cena stanice.** Cena po otkupnoj stanici je ista bez obzira na režim rada (Desktop-only ili PWA-led). Razliku u vrednosti pokriva cena Mobile paketa. Ovom odlukom se **zatvara odluka 9 iz `07B_ENTERPRISE_OPERATING_MODES.md`** (predlog da pricing razlikuje desktop-only od PWA-led vrednosti po stanici).
407. **Hardverska marža.** Hardver ostaje na planiranoj marži do izbora dobavljača. Marža se ne prepravlja pre nego što postoje stvarne nabavne cene.
408. **Redosled cena i unit economics-a.** Cene su određene. Unit economics je kasnije fino podešavanje, ne preduslov za izlazak sa ponudom.

---

### 26.1 Dopuna 27.07.2026. — odluke 409–422

**Datum dopune:** 27.07.2026.  
**Obuhvat:** satnice, obračunska jedinica modula, dodatna instanca, formiranje i prikaz cena, politika popusta, tarife Savetnika, i razrešenje pet preostalih konflikata iz cenovnog i komercijalnog dela.

> Numeracija nastavlja niz posle 408. Brojevi 322 i 379–400 se ne koriste.

#### Satnice i usluge

409. **Dve standardne satnice.** Postoje tačno dve standardne satnice:

- **razvojna, 50 €/h** — razvoj po zahtevu, složena migracija podataka, novi adapteri, posebni izveštaji, masovne korekcije podataka;
- **implementaciona, 30 €/h** — obuka preko uključenih pet sati, konfiguracija, čišćenje podataka, IT setup, procesni konsalting, rad na lokaciji.

Satnica se određuje prema **prirodi posla, ne prema mestu izvođenja**. Satnice su fiksne i nepregovaračke (vidi 418). *Menja odluku 110. Potvrđuje 353 i 354.*

410. **Obračun izlaska na teren.** 50 € po izlasku, uvećano za gorivo, vreme puta i vreme rada. Vreme puta obračunava se **uvek po implementacionoj satnici (30 €/h)**, bez obzira na vrstu posla. Vreme rada obračunava se po satnici koja odgovara prirodi posla iz odluke 409. *Precizira odluku 355.*

411. **Raspoređivanje usluga iz C7 po satnicama.** Čišćenje podataka, IT setup i procesni konsalting obračunavaju se po implementacionoj satnici. Masovne korekcije podataka i posebni izveštaji obračunavaju se po razvojnoj satnici, jer zahtevaju programsku intervenciju. *Precizira C7 u skladu sa 409.*

#### Moduli i instance

412. **Obračunska jedinica modula.** SEF, Banka i Dispatch plaćaju se **jednom po pravnom licu** i važe kroz sve njegove instance. Hladnjača/Proizvodnja plaća se **po proizvodnom pogonu** — svaki pogon koji koristi modul plaća ga. *Precizira 350. Potvrđuje 124, 131, 132 i 135.*

> `REVIEW 27.07.2026.` — poziv na odluku 135 traži proveru. Odluka 135 glasi „moduli se plaćaju jednom po pravnom licu, važe kroz sve instance“, a 412 izuzima Hladnjača/Proizvodnja i naplaćuje ga po pogonu. Odluka 412 time **ograničava** 135, ne potvrđuje je. Napomena: odluka 132 („dodatni pogon: dodatna instanca i dodatni proizvodni dodatak“) već je implicirala isto, pa 412 zapravo razrešava zatečeni sudar 135 protiv 132. Operativni tekst odluke 412 je nedvosmislen i primenjen je kako je napisan; koriguje se samo pozivanje.

413. **Dodatna instanca.** Druga i svaka naredna instanca istog pravnog lica dobija **−50 %** na sve što ta instanca dodatno koristi, uključujući modul Hladnjača/Proizvodnja (**200 € po dodatnom pogonu**). Moduli koji se po odluci 412 plaćaju po pravnom licu ne dupliraju se.

**Osnovica za obračun −50 % je lista cena pojedinačnih stavki koje instanca koristi, nikada bundle cena all-in paketa.**

Broj uključenih aktivnih stanica ostaje pet po pravnom licu i ne uvećava se otvaranjem dodatne instance. *Zamenjuje odluku 133. Precizira 358 i 132.*

#### Formiranje i prikaz cena

414. **Formiranje cene Mobile paketa.** Cena Mobile paketa jednaka je ceni odgovarajućeg Desktop paketa uvećanoj za **fiksni Mobile dodatak od 1.000 €**. Dodatak je isti na baznom i na all-in nivou i ne može biti niži od bazne cene Desktop paketa. *Zamenjuje odluke 126 i 127.*

415. **Formiranje all-in cene.** All-in varijanta paketa jednaka je baznoj ceni uvećanoj za **fiksnu all-in doplatu od 700 €**, istu u oba paketa. Doplata mora ostati niža od zbira liste cena uključenih modula i različita od zbira svakog mogućeg podskupa modula, kako nijedna kombinacija à la carte ne bi koštala isto kao all-in.

**Sastav all-in paketa:** Desktop all-in obuhvata SEF, Banku i Hladnjača/Proizvodnja. Mobile all-in dodatno obuhvata i Dispatch, koji je po odluci 293 dostupan isključivo uz Mobile. Ovaj sastav mora biti izričito naveden u cenovniku i u Prilogu 1 ugovora. *Dopunjuje odluku 349.*

416. **Politika prikaza cena.** Bazne cene paketa objavljuju se kao „od X € godišnje“, gde X obuhvata jedno pravno lice i do pet aktivnih stanica. Cene modula, dodatne stanice, Gazdinstva i Savetnika objavljuju se kao tačni fiksni iznosi. *Zamenjuje odluke 156 i 157. Potvrđuje 158, 159, 297 i 298.*

417. **Prikaz GGAP-a u cenovniku.** GGAP ostaje u cenovniku sa cenom „od 1.000 € godišnje po pravnom licu“, uz **obaveznu vidljivu oznaku „na upit, uz potvrdu obima — nije deo standardne ponude“**. Bez te oznake stavka se ne sme prikazati, jer bi je prodaja mogla kotirati kao redovnu. *Precizira 352 u skladu sa 405.*

#### Popusti

418. **Politika popusta.** Ne postoje pregovarački ni individualni popusti. Cena zavisi isključivo od paketa, broja aktivnih stanica, izabranih modula i broja instanci. Satnice iz odluke 409 su fiksne bez obzira na obim ili trajanje odnosa.

Jedina dozvoljena cenovna razlika unutar istog obima je **−50 % na drugu i svaku narednu instancu istog pravnog lica** (odluka 413).

Cenovne razlike koje su **objavljene u cenovniku** nisu popusti u smislu ove odluke i ostaju na snazi: kanalska cena Gazdinstva (339), all-in bundle doplata (415), Enterprise tarifa Savetnika (419), prvih 50 Basic naloga za partnera (161) i prva godina od produkcijskog puštanja (367, 422).

Prelazne odredbe 151 i 334 nisu popusti nego zaštita zatečene cene i ostaju na snazi.

*Briše C3, 111 i 112. Uklanja oznaku ⚠ sa odluka 30 i 113, koje od sada važe bez izuzetka.*

#### Savetnik

419. **Dve tarife Savetnika.** AgriX Savetnik ima dve objavljene tarife:

| Tarifa | Osnovica (do 10 aktivnih gazdinstava) | Svako gazdinstvo preko 10 |
|---|---|---|
| **Standalone** — savetnik bez drugog ugovornog odnosa sa AgriX-om | 150 € | 15 € |
| **Enterprise** — savetnik ili interna agronomska služba pravnog lica sa aktivnim Enterprise ugovorom | **100 €** | **10 €** |

Enterprise tarifa je uslovljena aktivnim Enterprise ugovorom. Prestankom tog ugovora korisnik prelazi na standalone tarifu pri prvoj narednoj obnovi, bez retroaktivnog obračuna.

Kada su gazdinstva u portfelju već pokrivena partnerskim paketom hladnjače, savetnik plaća samo alat po Enterprise tarifi; Pro pretplate tih gazdinstava se ne plaćaju ponovo.

*Zamenjuje odluku 207. Precizira 215 i 348. Ne dira 340 — gazdinstva koja nisu pokrivena partnerskim paketom i dalje drže sopstveni Pro po kanalskoj ceni.*

420. **Model naplate Savetnika.** Savetnik se naplaćuje kao **osnovica koja uključuje do 10 aktivnih gazdinstava, uvećana za fiksni iznos po svakom aktivnom gazdinstvu preko 10**, prema tarifi iz odluke 419. Aktivno gazdinstvo definisano je odlukom 342. *Zamenjuje odluku 198, koja je opisivala čist obračun po broju gazdinstava bez osnovice.*

#### Zatvaranje preostalih stavki

421. **Cena modula Hladnjača/Proizvodnja.** Fiksna godišnja naknada iznosi **400 € po proizvodnom pogonu**. Uslov standardizacije iz odluke 130 smatra se ispunjenim; cena više nije uslovljena. *Zatvara odluku 130. Potvrđuje 350 i 412.*

422. **Jedno gratis pravilo.** Prva godina besplatnog korišćenja od produkcijskog puštanja (odluka 367) je **jedino** gratis pravilo u modelu. Pogodnost finansijera razvoja iz IP4 je ta ista pogodnost, ne dodatna. Finansijer zadržava prioritet u redosledu razvoja, ali ne dobija drugu besplatnu godinu. *Prepisuje IP4 i uklanja njegovu oznaku ⚠. Potvrđuje 367.*

---

## 27. Napomena o kontinuitetu numeracije

Stanje numerisanih odluka u repou:

| Opseg | Status |
|---|---|
| 1–321 | uneto (odeljci 1–11) |
| 322 | **ne koristi se** — niz se nastavlja od 323 |
| 323–378 | uneto (odeljak 25); izvor `09B_ODLUKE_PO_OBLASTIMA.md` |
| 379–400 | **ne koriste se** — nisu dodeljene |
| 401–408 | uneto (odeljak 26) |
| 409–422 | uneto (odeljak 26.1) — cenovna revizija 27.07.2026. |

Uz numerisane odluke važe i serije A, BC, C, D, I, IP, L, LEG, M, MKT, ML, ON, P, PRT, Q i S.

Napomene o izvoru: odluke 323–378 unete su iz tematskog indeksa, koji je sam po sebi sažetak („jedna linija po odluci“). Gde je bio dostupan detaljniji izvor — cenovnik, ugovor, definicija proizvoda ili finansijski model — formulacija je dopunjena i izvor je naveden. Odluka **370** ne postoji u tematskim tabelama indeksa; rekonstruisana je iz tabele razrešenih konflikata (K-14) i člana 8. stav 7. ugovora.

Odluke koje su **obrisane** cenovnom revizijom 27.07.2026. (odluka 418) i više se ne vode: **C3**, **111**, **112**. Ne zamenjuju se novim brojem — politika popusta je od tada u odluci 418.

Tematski indeks `09B_ODLUKE_PO_OBLASTIMA.md` verzije 3 obuhvata ceo opseg, zaključno sa 422.
