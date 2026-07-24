# AgriX — Q&A Decision Log

**Datum sesije:** 24.07.2026.  
**Status:** radni strateški zapis  
**Obuhvat:** odluke, korekcije, pretpostavke i otvorena pitanja iz Q&A sesije 1–260.

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
110. **Satnica.** Jedna standardna satnica, uz mogući individualni popust većim ili dugoročnim klijentima.
111. **Osnova popusta na custom rad.** Veći unapred dogovoreni obim i dugoročna ukupna vrednost odnosa.
112. **Maksimalni custom popust.** Ne postoji fiksni maksimum; odlučuje se pojedinačno.
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
126. **Mobile multiplikator.** Cena Mobile-a definiše se kao unapred utvrđen odnos prema Desktop-u.
127. **Minimalni odnos.** Desktop Otkup + Mobile treba da bude najmanje dva puta cena Desktop Otkup-a.
128. **Struktura ponude.** Desktop Otkup je baza; Mobile je dodatak; Hladnjača/Proizvodnja je nezavisan Desktop dodatak; SEF, Banka i Dispatch su odvojeni.
129. **Mobile i proizvodnja.** Mobile pokriva teren/transport; proizvodnja ostaje Desktop funkcionalnost.
130. **Proizvodni dodatak — početno.** Nakon standardizacije ima fiksnu godišnju naknadu.
131. **Jedan pogon.** Proizvodni dodatak pokriva jedan proizvodni pogon.
132. **Više pogona.** Dodatni pogon istog pravnog lica zahteva dodatnu Desktop instancu i dodatni proizvodni dodatak.
133. **Dodatna Desktop instanca.** Cena se određuje individualno prema razlogu, scope-u i složenosti.
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
156. **Javne cene Enterprise-a.** Objavljuju se rasponi za Desktop i Mobile.
157. **Javna cena proizvodnje.** Objavljuje se cenovni raspon za Hladnjača/Proizvodnja.
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
198. **Naplata.** Prema broju aktivnih gazdinstava.
199. **Naziv proizvoda.** Poseban proizvod: **AgriX Savetnik**.
200. **Licence gazdinstava.** Cena Savetnika pokriva aktivna gazdinstva; ona ne plaćaju zaseban Pro.
201. **Pristup proizvođača.** Proizvođač zadržava sopstveni nalog i pristup podacima.
202. **Ciljni kupci.** Samostalni savetnici/agronomi i savetodavne firme; dugoročno i timski rad.
203. **Prioritet razvoja.** Osnovna verzija do sezone 2027, bez usporavanja Enterprise proizvodnog sistema.
204. **Prva verzija.** Jedan savetnik vodi više gazdinstava; timovi i raspodela dolaze kasnije.
205. **Veza sa Enterprise-om.** Isto gazdinstvo može biti povezano i sa Savetnikom i sa jednom ili više hladnjača.
206. **Interne agronomske službe.** Veće poljoprivredne firme sa sopstvenim agronomima su ravnopravna ciljna grupa.
207. **Ista tarifa.** Isti model cene po aktivnom gazdinstvu za nezavisne savetnike i interne agronomske službe.
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
