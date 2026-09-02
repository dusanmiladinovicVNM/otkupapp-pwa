"""Generise tests/fixtures/otkup_test.xlsm iz donor sveske.

Zasto donor a ne "od nule": osnovna sema (sheetovi + ListObject-i sa kolonama)
ne postoji nigde u kodu -- Ensure* rutine u modSetup samo DODAJU kolone na
postojece tabele, a spiskovi kolona osnovnih tabela zive iskljucivo u .xlsm.
Zakucavanje tih spiskova u Python napravilo bi drugi izvor istine koji konkurise
svesci (CLAUDE.md S4: "Sema tabela je izvor istine, ne kod"). Zato: struktura se
uzima iz donora, a podaci su 100% sinteticki.

Donor se NIKAD ne menja -- radi se nad kopijom.

Rezultat: sveska u kojoj su svi redovi obrisani (osim kataloga -- vidi KEEP_ROWS)
i posejani samo test unosi. Nijedan klijentski podatak ne moze da zavrsi u
tests/golden/*.txt koji idu u git.

    python tools/make_fixture.py --donor "C:/.../AgriX_2.28.4.xlsm"
    python tools/make_fixture.py --donor <put> --out tests/fixtures/otkup_test.xlsm --force

Windows + Excel + pywin32. Semu donora ispisuje tools/dump_schema.py.
"""

import argparse
import datetime
import hashlib
import os
import shutil
import sys

MSO_AUTOMATION_SECURITY_LOW = 1

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(ROOT, "tests", "fixtures", "otkup_test.xlsm")

# Tabele kojima se redovi NE brisu: katalog poruka (Poruka("KLJUC") bez njega
# vraca prazno) i config tabele (iz njih citaju GetConfigValue/GetLocalConfigValue).
KEEP_ROWS = {"tblporuke", "tblsefconfig", "tblconfig", "tbllocalconfig"}

# FIKSAN datum, ne "danas": golden snapshot hvata txtDatum, pa bi fixture vezan
# za danasnji dan obarao golden fajlove svaki sledeci dan.
FIXTURE_DATE = datetime.date(2026, 3, 15)

STATUS_AKTIVAN = "Aktivan"          # modConfig.STATUS_AKTIVAN
AMB_12_1 = "12/1"                   # modConfig.AMB_12_1

STANICA = "STA-TEST-1"
# Druga stanica: BrojDokumenta otkupa i BrojOtpremnice su scoped PO STANICI, pa
# isti broj na dva OM-a postoji legitimno. Bez druge stanice se ta kolizija ne
# moze ni napisati kao test.
STANICA2 = "STA-TEST-2"
VOZAC = "VOZ-TEST-1"
# Drugi vozac postoji zbog CILJNE liste zbirnih. Broj zbirne generator DRZI
# JEDINSTVENIM (SuggestNextBroj za ZBR bumpuje sekvencu dok ne nadje slobodan),
# pa dve zbirne istog broja mogu nastati samo RUCNIM UNOSOM ili uvozom -- a
# tada lista ciljeva mora da ponudi oba dokumenta, ne jedan spojen red.
VOZAC2 = "VOZ-TEST-2"
ZBIRNA_DUPL = "ZB-TEST-DUPL"        # isti broj, isti kupac, dva vozaca
# Zbirna u koju NIJEDAN test ne upisuje. Kolizioni par aktivnih prijemnica mora
# da pocne na takvoj: na ZB-TEST-1 je okidao dijalog "dupla prijemnica" iz testa
# koji tamo upisuje, a MsgBox u headless runu je visenje koje watchdog samo
# maskira.
ZBIRNA_MIRNA = "ZB-TEST-4"
# Svez par zbirnih za KASKADU: test 38 stornira jedno zaglavlje, pa bi posle
# njega ostao jedan aktivan vlasnik i kapija ne bi imala sta da detektuje.
ZBIRNA_KASK = "ZB-TEST-KASK"
OTKUP_KOLIZIJA = "7/150326"          # isti BrDok na dva otkupna mesta
OTPREMNICA_KOLIZIJA = "8/TEST"       # isti broj otpremnice, dve stanice, ista zbirna
# ZATECEN par BEZ generacije, za zavrsetak ispravke: OldDocID je tacan, pa se ne
# sme degradirati u prazan opseg i zavrsiti na golom broju. Testovi koji pecate
# generacije ne smeju da ga dodirnu, pa ima svoj broj.
OTPREMNICA_LEGACY = "6/TEST"
# ZATECEN ("stale") context za zavrsetak ispravke. Dve otpremnice ISTOG broja sa
# RAZLICITIM roditeljima: prva u tabeli visi na JEDNOZNACNOJ zbirni, a izabrana
# (ona iz context-a) na DVOSMISLENOJ. Lookup po poslovnom broju zato vraca
# POGRESNOG roditelja -- kapija proveri jednoznacnu zbirnu siblinga, a kod mutira
# dvosmislenu zbirnu izabranog dokumenta. Bez ovog para se ta razlika ne meri.
OTPREMNICA_STALE = "9/TEST"
OTPREMNICA_STALE_NOVA = "10/TEST"    # cilj zamene; mora imati NEPRAZNU zbirnu
# Namenska zbirna i prijemnica za stale scenario. NIJEDAN drugi test ih ne
# dira: naslanjanje na ZB-TEST-4 i PRJ-KASK-1 je poslovnu tvrdnju cinilo
# vakuumskom -- dotle bi ih drugi testovi vec pomerili ili stornirali, pa
# relink nije imao sta da preveze i sabotaza nije obarala svoju tvrdnju.
ZBIRNA_STALE = "ZB-TEST-STL"         # cilj relinka; jednoznacna, mirna
PRIJEMNICA_STALE = "12/TEST"         # TUDJA prijemnica na dvosmislenoj zbirni
# F8 DODATNI STORNO BLOKOVA. Isti BrojOtpremnice na dve stanice je legitiman,
# a spisak blokova se pravio po tom broju -- pa je u korpu ulazio i blok
# drugog dokumenta. Zato je B STORNIRANA a njen blok AKTIVAN: tacno to stanje
# gasi kapiju BlockStornoDriftReason (nema zivog roditelja), a i inace je ona
# na ovoj putanji preskocena za DUPLI/PONISTENJE.
OTPREMNICA_BLOK = "18/TEST"
# CILJNA zbirna sa istorijski dvosmislenim brojem: owner A je STORNIRAN a
# njegovo dete je AKTIVNO (test 44 dokazuje da storno zaglavlja ne stornira
# dete), owner B je aktivan i njegovo dete je zamena. Zatecena kapija u
# writeru broji samo AKTIVNE vlasnike, pa ovde vidi jednog i pusti relink --
# a rekalkulacija po broju onda sabere decu OBA scope-a.
ZBIRNA_TGT = "ZB-TEST-TGT"
ZBIRNA_OLDU = "ZB-TEST-OLDU"        # izvorna zbirna; jednoznacna, mirna
OTPREMNICA_OLD_U = "13/TEST"        # izvorna otpremnica koja se ispravlja
OTPREMNICA_HIST = "14/TEST"         # aktivno dete STORNIRANOG vlasnika cilja
OTPREMNICA_NEW_T = "15/TEST"        # zamena; dete AKTIVNOG vlasnika cilja
PRIJEMNICA_OLD_U = "16/TEST"        # izvorna prijemnica; ne sme se prevezati
OTPREMNICA_ZAMENA = "7/TEST"
VRSTA = "TESTVOCE"
# Druga vrsta postoji zbog jedne tvrdnje koju ranije nije bilo cime napisati:
# presuda o RELABEL-u mora da opisuje BAS izabran dokument. PRJ-TEST-C2 je zato
# druge vrste od svog kolizionog blizanca C1 i od cilja -- kad se presuda racuna
# po broju, ona vidi C1 (ista vrsta kao cilj) i kaze CLEAN, pa se relabel
# preskoci i paleta ostane pogresno oznacena.
VRSTA2 = "TESTVOCE2"
SORTA = "TESTSORTA"
ZBIRNA = "ZB-TEST-1"
ZBIRNA_U_BLOKU = "ZB-TEST-3"        # zbirnu nosi otkupni blok, ne otpremnica
ZBIRNA_STORNIRANA = "ZB-TEST-STORNO"  # ne sme se pojaviti u listi ciljeva
# Druga AKTIVNA zbirna: nosi kolizioni par prijemnica. Da su oni na ZB-TEST-1,
# provera "1 zbirna = 1 prijemnica" bi u testovima otvarala dijalog, a dijalog
# u headless runu nema ko da zatvori.
ZBIRNA2 = "ZB-TEST-2"
# Kupac postoji SAMO kao ID na fakturi -- red u tblKupci ne treba: kapije koje
# ga koriste porede identifikatore, ne citaju karticu kupca.
KUPAC = "KUP-TEST-1"
KUPAC2 = "KUP-TEST-2"
FAKTURA = "FAK-TEST-1"
FAKTURA_IZNOS = 10000
# Iznos = 0 -> kapija nad uplatom se na nju ne primenjuje (v. tblFakture dole).
FAKTURA_BEZ_IZNOSA = "FAK-TEST-0"
# EKRAN FAKTURISANJE (Faza E/16). Do sada su tblFakture imale dva reda BEZ
# broja, datuma i statusa, i nijedna uplata nije bila vezana za fakturu, pa je
# svaka tvrdnja o listi faktura, cipovima i slaganju sa GetOpenFakture bila
# zelena bez pokrica -- svi filteri su radili nad praznim skupom.
FAKTURA_NEPL = "FAK-TEST-N"      # KUPAC2, 5000, Neplaceno, bez uplate
FAKTURA_PLAC = "FAK-TEST-P"      # KUPAC, 4000, Placeno, uplata pokriva ceo iznos
FAKTURA_STORNO = "FAK-TEST-X"    # KUPAC, 7000, STORNIRANA -- ne sme u listu
FAKTURA_NEPL_IZNOS = 5000
FAKTURA_PLAC_IZNOS = 4000
# Dva AKTIVNA reda tblNovac pod ISTIM brojem -- avans raspodela to radi
# svakodnevno. Bez NovacID-a je broj dvosmislen i storno se odbija; sa njim se
# stornira bas izabran red. Preflight je do sada odbijao i kad ID postoji.
NOVAC_DUPLI_BROJ = "NOV-DUPLI-1"

# BANKA UVOZ (Faza E/17). tblBankaImport je do sada bio PRAZAN, pa je svaka
# tvrdnja o listi stavki, cipovima, jakim kljucevima i integritetu izvoda radila
# nad praznim skupom -- zelena bez pokrica.
#
# BROJ IZVODA NIJE IDENTITET. Dedupe kljuc (IsDuplicateBankaImport) pocinje od
# BROJA RACUNA: "Drugi racun = druga transakcija, bez obzira na broj izvoda i
# iznos". Dva racuna firme zato legitimno nose izvod ISTOG broja, i bez takvog
# para u fixture-u tvrdnja "identitet je BankaImportID, ne BrojDokumenta" nema
# nad cim da padne.
BIM_IZVOD_1 = "IZV-FIX-1"                 # isti broj na dva racuna -> kolizija
BIM_IZVOD_2 = "IZV-FIX-2"                 # izvod kome se saldo NE slaze
BIM_IZVOD_3 = "IZV-FIX-3"                 # nosi DVOSMISLEN BankaImportID
BIM_RACUN_1 = "160-0000000111111-11"
BIM_RACUN_2 = "265-0000000222222-22"
BIM_DATUM_1 = datetime.date(2026, 3, 16)
BIM_DATUM_2 = datetime.date(2026, 3, 17)
# Blok sa TRI otvorene otkupne stavke -- preko MAX_BLOK_KANDIDATA (2), pa
# automatska raspodela dize ERR_BMAP_MANUAL_REQUIRED i red ide na rucno.
BIM_BLOK_3 = "BLK-BIM-3"
# ISTI kooperant, ISTI broj bloka, DVA otkupna mesta. Broj otkupa je
# jedinstven PO STANICI, pa je ovo legitiman podatak -- i jedini nacin da
# se izmeri da rucno mapiranje nosi scope otkupnog mesta, a ne samo broj.
# Bez ovog para bi novac mogao da ode na pogresan otkupni lanac, a nijedna
# tvrdnja to ne bi primetila.
BIM_BLOK_OM = "BLK-BIM-OM"
# POTPUNO PLACEN blok. Lista blokova nudi SVAKI nestorniran broj otkupa i ne
# proverava da li blok jos duguje, a kandidati za placanje se biraju samo ako
# je "otvoreno > 0.009" -- pa placen blok legitimno postoji u listi a daje
# NULA kandidata. Writer to ne prijavljuje kao gresku nego knjizi AVANS i
# stavku oznaci obradjenom. Bez ovog para redova (otkup + uplata koja ga
# zatvara) nijedna tvrdnja ne bi mogla da vidi da rucni izbor takvog bloka
# mora da STANE.
BIM_BLOK_PLACEN = "BLK-BIM-PLAC"
BIM_OTK_PLACEN = "OTK-BIM-PLAC"
BIM_OTK_PLACEN_IZNOS = 500.0      # 10 * 50
# Isti broj izvoda i isti racun, ali DRUGI ciklus. Banke numeraciju
# ponavljaju po godini; bez datuma u kljucu bi se dva izvoda spojila u
# jedan sinteticki red koji nikad nije postojao.
BIM_DATUM_PY = datetime.date(2025, 3, 16)
# Izvod cija DVA REDA nose RAZLICITE zbirove. Danasnji parser to ne moze da
# napravi -- kopira isti saldo u petlji -- ali rucno editovan red, delimican
# re-import ili buduci parser mogu. Bez ovog para agregat bi brojku PRVOG
# reda prikazao kao istinu o celom izvodu, a nijedna tvrdnja to ne bi videla.
BIM_IZVOD_NES = "IZV-FIX-NES"
# Saldo izvoda 2 je NAMERNO za 100 veci od tacnog (8000 + 950 - 3000 = 5950).
# Bez reda koji se ne slaze, provera integriteta bi u fixture-u bila uvek OK.
BIM_IZVOD_2_ZAVRSNO = 6050

# KOLIZIJA BROJEVA -- srce ovog fixture-a.
#
# BrojPrijemnice NIJE globalno jedinstven: GenerateBrojPrijemnice racuna
# sekvencu PO KUPCU, pa dva kupca istog dana dobiju isti "1/ddmmyy". Dokle god
# u fixture-u nije postojao takav par, svaka tvrdnja oblika "cilj/izvor se
# jednoznacno razresava po broju" bila je zelena bez pokrica -- i sabotaza koja
# bi je oborila prolazila je neprimeceno.
PRIJEMNICA_BROJ = "1/150326"        # isti broj kod KUPAC i KUPAC2
PRIJEMNICA_STORNO = "9/150326"      # stornirana; njene palete su osirocene
PRIJEMNICA_STORNO2 = "8/150326"     # kolizioni par storniranih (dva kupca)
# Kolizioni par AKTIVNIH prijemnica sa svojim paletama. Postoji zbog jednog
# propusta koji se video tek kad se tvrdnja napisala: prevezivanje prijemnice na
# zbirnu menjalo je tblPrijemnica po identitetu, a tblPaletaStavka jos po BROJU
# -- pa je dokument drugog kupca ostajao sam sebi protivrecan (prijemnica na
# staroj zbirni, njena paleta na novoj).
PRIJEMNICA_ZBR_KOLIZIJA = "6/150326"
# DELJENA FIZICKA PALETA. Dva kupca istog broja i ISTE robe smeju legitimno da
# dele paletu -- roba im je identicna, pa nema sta da se razlikuje. Kapija koja
# su-stanara trazi po BROJU tu ne okine (bpg == oldBroj), pa relabel prepravi
# header cele palete i tudja roba ostane pogresno oznacena.
PRIJEMNICA_DELJENA = "5/150326"
# Aktivan cilj DRUGE vrste -- da prevezivanje uopste bude RELABEL.
PRIJEMNICA_CILJ_V2 = "4/150326"
# AGROHEMIJA. Magacin do sada nije imao nijedan red u fixture-u, pa je svaka
# tvrdnja o stanju, dugu i smart dozi bila zelena bez pokrica.
#
# ART-TEST-1 nosi Pakovanje 5 i DozaPoHa 2. Te dve vrednosti su izabrane tako da
# se ZAOKRUZENJE NAGORE vidi: 1.5 ha * 2 = 3 l, a pakovanje je 5 l -> jedno
# pakovanje. Da je pakovanje 1, ceo racun bi izgledao ispravno i kad bi se
# zaokruzivalo nanize ili matematicki.
#
# ART-TEST-2 je BEZ Pakovanja -- invarijanta "svaki artikal ima Pakovanje" je
# kapija izdavanja, pa mora da postoji red nad kojim ona pada.
#
# ART-TEST-3 ima Pakovanje ali NEMA nijedan magacin red -> stanje 0, pa kapija
# stanja ima nad cim da padne i kad artikal postoji.
ARTIKAL = "ART-TEST-1"
ARTIKAL_BEZ_PAK = "ART-TEST-2"
ARTIKAL_BEZ_STANJA = "ART-TEST-3"
# Artikal sa VELIKOM zalihom i pakovanjem od 1. Postoji zbog trake korpe: da bi
# se izmerio preliv ("i jos N"), u korpu mora da udje vise stavki nego sto traka
# ima redova -- a ART-TEST-1 to ne dozvoljava, jer mu kapija stanja (15 kg,
# pakovanje 5) propusta najvise tri pakovanja. Nema nijedan IZLAZ, pa ne ulazi
# ni u jedan dug.
ARTIKAL_ZALIHA = "ART-TEST-Z"
ARTIKAL_PAKOVANJE = 5
ARTIKAL_DOZA = 2
ARTIKAL_CENA = 500
# ULAZ 20 l, pa IZLAZ 5 l kooperantu KOOP-TEST-1 -> stanje 15, dug 2500.
ARTIKAL_STANJE = 15
AGRO_DUG_KOOP1 = 2500
# Odbitak duga: 300 + 200 (storniranih 999 i tudji tip 777 se NE broje).
AGRO_ABZUG_KOOP1 = 500
AGRO_ABZUG_KOOP2 = 100

# Kolizija broja po godini -- isti broj, dve godine, dva razlicita identiteta.
PALETA_KOLIZIJA_BROJ = 12
PALETA_KOLIZIJA_ID = "PAL-TEST-Y25"      # 12/2025; PAL-TEST-Z2 je 12/2026
PRERADA_KOLIZIJA_BROJ = 7
PRERADA_NOVA_ID = "PRE-TEST-Y26"
PRERADA_STARA_ID = "PRE-TEST-Y25"
# Dve AKTIVNE prijemnice istog broja za ISPRAVKU. Zaseban broj: test 35 pravi
# RESI KASNIJE context nad 6/150326, a pending ispravka nad istim brojem bi
# zaustavila ISPRAVKU (safe-stop) i test bi merio pogresnu stvar.
PRIJEMNICA_ISPRAVKA = "3/150326"

# EKRAN IZVESTAJI (Faza E/19). Dve rupe koje su tvrdnje slaganja do sada cinile
# zelenima bez pokrica:
#
# 1) tblNovac nije imao NIJEDAN red sa OMID-om, nijedan KesOtkupacKoop i
#    nijedan Firma->Otkupac avans -- ReportIsplata("OM", ...) je bio prazan,
#    GetOMAvansSaldo svuda 0, pa slaganje "Ukupno = Kes + VirmanFirma +
#    VirmanAvans" i "OM AVANS red = GetOMAvansSaldo" nije imalo nad cim da
#    padne. Novi redovi su vezani za OTK-IZV-1, blok koji je U CELOSTI placen
#    (200 kes + 300 virman = 500): GetOpenOtkupi ga zato NE vidi, pa ekran
#    Platni nalozi (KPI, cipovi, korpa) ostaje bit-identican.
# 2) tblAmbalaza je bio PRAZAN (nije ni u KEEP_ROWS ni u SEED-u) -- sve
#    ambalazne tvrdnje (SALDO kolona, lista AMBALAZA, kartica ambalaze,
#    running saldo) radile bi nad praznim skupom.
#
# KOOP-IZV-AV je kooperant SA avansom (OMID=STANICA2) a BEZ ijednog otkup
# bloka: ReportIsplata dobija VirmanAvans kanal, a avans pool ekrana Platni
# nalozi se ne pomera (pool se sabira po kooperantima OTVORENIH blokova).
OTK_IZV_ZATVOREN = "OTK-IZV-1"
IZV_KES = 200.0
IZV_VIRMAN = 300.0
IZV_AVANS_KOOP = 700.0
IZV_OM_AVANS = 5000.0            # KesFirmaOtkupac; GetOMAvansSaldo = 5000 - 200
AMB_LETVA = "Letvarica"          # drugi tip ambalaze; slobodan unos kao u pogonu

# EKRAN SLEDLJIVOST (Faza E/20, v6-ui-187). Do sada fixture NIJE imao nijednu
# otpremnicu cija zbirna nosi NESTORNIRANU prijemnicu (rupa zapisana u
# par. 23.12/S10), pa se potpun lanac otkup -> otpremnica -> zbirna ->
# prijemnica -> faktura nije mogao ni napisati kao tvrdnja -- svaka provera
# slaganja "napred == nazad == rucni prolaz" merila bi prazan skup.
#
# SVA vozila zive na STA-TEST-2 i SVI blokovi lanca su U CELOSTI placeni
# (VirmanFirmaKoop redovi dole, OMID=STANICA2): GetOpenOtkupi ih ne vidi, pa
# KPI/cipovi/korpa ekrana Platni nalozi ostaju bit-identicni (isti razlog kao
# OTK_IZV_ZATVOREN), a T_WriterGuard_AvansSaldoOM preduslov (STA-TEST-1 avans
# saldo 0) ostaje netaknut -- virman firma->koop ne dira avans pool.
#
# Vozila (bez ambalaze -- KolAmbalaze se ne seje, da kanonski amb saldo i
# kartice ne dobiju kretanja bez ledger parova; v. par. 23.6 nalaz 1):
#   POTPUN LANAC: OTK-SLED-1 (KOOP-TEST-2, parcela PAR-TEST-2, 300 kg) +
#     OTK-SLED-2 (KOOP-TEST-IME, BEZ parcele, 200 kg) -> OTP-SLED-1 (500 kg)
#     -> ZB-TEST-SLED (500 kg, VOZAC2, KUPAC) -> PRJ-SLED-1 (500 kg,
#     fakturisana) -> FAK-SLED-1. Kg se slaze niz CEO lanac; blok bez
#     parcele je ujedno vozilo za oznaku "bez parcele" na listi PARCELE.
#     KOOP-TEST-3 se NE sme koristiti ni za jedan SLED blok:
#     T_BankaUvoz_RucnoMapiranjePravila broji NJEGOVE blokove apsolutno
#     (GetBlokoviZaBimMapiranje = 5), pa bi svaki nov blok oborio tudji
#     test. KOOP-TEST-IME je i namerno: dva istoimena kooperanta u istom
#     lancu dokazuju da je identitet reda OTK|id, ne prikazano ime.
#   DVOSMISLEN BROJ: OTK-SLED-D -> OTP-SLED-D (VozacID PRAZAN) ->
#     ZB-TEST-SLDD, broj koji dele DVA aktivna vlasnika (dva vozaca).
#     SVOJ par, ne ZB-TEST-DUPL: DUPL par trosi test 22 (StornoZbirna_TX
#     stornira ZBI-DUPL-2), pa u trenutku sledljivost testova ima JEDNOG
#     aktivnog vlasnika i vise nije dvosmislen. Bez vozaca otpremnice se
#     vlasnik ne moze razresiti -> IZV_VLASNIK_NEJASAN, fail-closed.
#   DO PRIJEMNICE BEZ FAKTURE: OTK-SLED-N -> OTP-SLED-N -> ZB-TEST-SLN
#     (KUPAC2) -> PRJ-SLED-N (Fakturisano=Ne). Krug 9: to je LEGITIMAN
#     tok (roba u hladnjaci) -- lanac je POTPUN, bez oznake i bez problema.
#   KG RAZLIKA + BEZ PRIJEMA: OTK-SLED-R (100 kg) -> OTP-SLED-R (250 kg!
#     kg curi na prvoj karici) -> ZB-TEST-SLR (250 kg, bez prijemnice).
SLED_ZBIRNA = "ZB-TEST-SLED"
SLED_ZBIRNA_N = "ZB-TEST-SLN"
SLED_ZBIRNA_R = "ZB-TEST-SLR"
SLED_ZBIRNA_D = "ZB-TEST-SLDD"     # dvosmislen broj; nijedan drugi test ga ne dira
SLED_PRIJ_BROJ = "30/150326"       # sekvenca KUPAC (1..22 zauzeti)
SLED_PRIJ_BROJ_N = "31/150326"     # sekvenca KUPAC2
SLED_FAKTURA = "FAK-SLED-1"
SLED_KG_1 = 300.0
SLED_KG_2 = 200.0
# Krug 8 (review paket):
#   F-LANAC (R1, ALL-pravilo fakturisanosti): zbirna sa DVE prijemnice --
#     F1 fakturisana na aktivnu FAK-SLED-2, F2 "Fakturisano=Da" ali
#     FakturaID pokazuje na nepostojecu fakturu -> lanac je NEFAKTURISANO
#     (jedna neispravna obara celu kariku), NEPOTPUNI prijavljuje F2.
#   M-LANAC (R2, SearchRefs): POTPUN lanac sa 2 prijemnice i 2 fakture --
#     prikaz "2 prij."/"2 fakt." guta brojeve, pretraga po FAK-SLED-3B
#     mora da nadje red i na LANAC i na PARCELE (ima parcelu PAR-TEST-2).
#   PAL-SLED-B (R4): paleta sa NEVALIDNIM datumom -- ugovor kaze da u
#     ponudi polja izbora ostaje VIDLJIVA (IIf mina bi ovde pukla).
SLED_ZBIRNA_F = "ZB-TEST-SLF"
SLED_ZBIRNA_M = "ZB-TEST-SLM"
SLED_FAKTURA_2 = "FAK-SLED-2"
SLED_FAKTURA_3 = "FAK-SLED-3"
SLED_FAKTURA_3B = "FAK-SLED-3B"
SLED_KG_F = 120.0                  # 60 + 60 po prijemnicama
SLED_KG_M = 200.0                  # 100 + 100 po prijemnicama

# GP grana lanca (Faza 1+2, v6-ui-189) -- tri nova lanca + potrosne prerade:
#   G "PRODATO GP":  OTK-SLED-G -> OTP-SLED-G (38/TEST) -> ZB-TEST-SLG ->
#     PRJ-SLED-G (37/150326, NEfakturisana -- hladnjaca tok) -> PAL-SLED-G
#     (50, ZATVORENA, Preradjeno=Da) -> PRE-SLED-G (51, Fakturisano=Da na
#     AKTIVNU FAK-SLED-GP 9/2026). Lanac kolone 28-30: "50/2026",
#     "51/2026 -> 9/2026", stanje "prodato GP".
#   H "U HLADNJACI": isti tok do PAL-SLED-H (60, ZATVORENA, SVEZA), bez
#     prerade i bez faktura -- stanje "u hladnjaci".
#   K KONTRADIKCIJA: PRE-SLED-K (61) tvrdi Fakturisano=Da na NEPOSTOJECU
#     FAK-NEMA-GP -> oznaka "faktura neusaglasena" + NEPOTPUNI red
#     (FAKTURA-VEZA-NEISPRAVNA, DokTip Prerada).
#   PRE-GP-W1 (71): cista nefakturisana -- potrosno vozilo writer testa
#     (CreateFakturaGP_TX je fakturise; mutirajuci test ide POSLE read-only).
#   PRE-GP-X (81, Stornirano=Da) i PRE-GP-W0 (91, NetoIzlazKg=0): negativi
#     kapija writera (stornirana / bez izlaza).
# "Otvoren tok" NEMA fixture vozilo -- ne tvrdi se testom (svaki cist lanac
# vec nosi neku dalju kariku).
# BROJEVI paleta (50/60/70) i prerada (51..91) se ZAVRSAVAJU na 0/1:
# brojevi paleta/prerada sada zive u SearchRefs/haystack-u LANCA, a
# pretraga je substring -- "NN/2026" sadrzi "N/2026", pa bi poslednja
# cifra 2-9 gutala broj neke fixture fakture (2/2026..9/2026). Iz istog
# razloga su DALEKO od generatorske sekvence: GenerateBrojPalete = max+1,
# pa je test u RunAll dobijao "38/2026" koje upit "8/2026" nalazi.
SLED_ZBIRNA_G = "ZB-TEST-SLG"
SLED_ZBIRNA_H = "ZB-TEST-SLH"
SLED_ZBIRNA_K = "ZB-TEST-SLK"
SLED_FAKTURA_GP = "FAK-SLED-GP"
SLED_KG_G = 80.0
SLED_KG_H = 90.0
SLED_KG_K = 70.0
SLED_GP_IZLAZ_KG = 64.0            # izlaz prerade G (kalo prerade je legalan)
SLED_GP_CENA = 200.0               # rucna cena GP -> iznos fakture 12.800
# Krug 5: P lanac -- DELIMICNA prodaja (prerada 120 kg izlaza, utovar
# 50 kg na fakturu 12/2026 -> stanje "delimicno prodato", 70 kg na
# stanju). Brojevi po 0/1 pravilu.
SLED_ZBIRNA_P = "ZB-TEST-SLP"
SLED_KG_P = 150.0

class Sirovo:
    """Vrednost koja se upisuje BEZ formata koji bi red nasledio od kolone.

    Nov `ListRows.Add` nasledi format prethodnog reda. Za broj koji stoji u
    datumskoj koloni to znaci da ga Excel pri citanju pokusa da vrati kao
    `Date` -- a broj van opsega datuma tada obara CELO citanje tabele
    (`GetTableData` -> Overflow), sto NIJE ono sto se desava na zatecenim
    svescima: tamo takva celija ima obican format i cita se kao broj.
    """

    def __init__(self, v):
        self.v = v

    def __repr__(self):
        # POTPIS FIXTURE-a hashira repr() posejanih vrednosti. Podrazumevani
        # repr objekta nosi adresu iz memorije, pa bi hash bio drugaciji u
        # svakom prolazu i run_vba bi svaki put javljao "USTAJAO FIXTURE".
        return "Sirovo(%r)" % (self.v,)


# Sejanje ide PO IMENU KOLONE -- ako donor nema neku od ovih kolona, skripta
# pukne glasno umesto da tiho napravi fixture nad kojim testovi lazu.
SEED = {
    "tblStanice": [
        {"StanicaID": STANICA, "Naziv": "Test Otkupno Mesto", "Mesto": "Test Mesto",
         "Aktivan": STATUS_AKTIVAN, "JeHladnjaca": "NE"},
        {"StanicaID": STANICA2, "Naziv": "Drugo Otkupno Mesto", "Mesto": "Test Mesto",
         "Aktivan": STATUS_AKTIVAN, "JeHladnjaca": "NE"},
    ],
    "tblVozaci": [
        {"VozacID": VOZAC, "Ime": "Test", "Prezime": "Vozac",
         "Aktivan": STATUS_AKTIVAN, "KapacitetKG": 5000},
        {"VozacID": VOZAC2, "Ime": "Drugi", "Prezime": "Vozac",
         "Aktivan": STATUS_AKTIVAN, "KapacitetKG": 4000},
    ],
    "tblKulture": [
        {"KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "GajbicaPoPaleti": 100, "Aktivan": STATUS_AKTIVAN, "TipAmbalaze": AMB_12_1},
    ],
    "tblTipAmbalaze": [
        {"TipAmbalaze": AMB_12_1, "TezinaGajbiceKg": 1.0, "Aktivan": STATUS_AKTIVAN},
    ],
    "tblKooperanti": [
        # KOOP-TEST-1 IMA tekuci racun: ekran Platni nalozi bez ijednog
        # kooperanta sa racunom nema nijedan blok koji sme u CSV, pa bi cip
        # "ima racun", kapija korpe i izbor za izvoz merili prazan skup.
        # KOOP-TEST-2 i KOOP-TEST-3 NAMERNO ostaju bez racuna -- to je "bez
        # racuna" polovina istih tvrdnji (blok se vidi, ne sme u naloge).
        {"KooperantID": "KOOP-TEST-1", "Ime": "Prvi", "Prezime": "Testni",
         "Mesto": "Test Mesto", "StanicaID": STANICA, "Aktivan": STATUS_AKTIVAN,
         "TekuciRacun": "205-0000000123-45"},
        {"KooperantID": "KOOP-TEST-2", "Ime": "Drugi", "Prezime": "Testni",
         "Mesto": "Test Mesto", "StanicaID": STANICA, "Aktivan": STATUS_AKTIVAN},
        {"KooperantID": "KOOP-TEST-3", "Ime": "Treci", "Prezime": "Testni",
         "Mesto": "Test Mesto", "StanicaID": STANICA, "Aktivan": STATUS_AKTIVAN},
        # ISTO IME kao KOOP-TEST-1, drugi identitet. Postoji zbog jednog pravila
        # koje se drugacije ne moze napisati: lista dugova pokazuje IME, a
        # dvoklik bira KOOPERANTA. Dok u fixture-u nije bilo dva istoimena,
        # tvrdnja "dvosmislen prikaz se odbija" nije imala nad cim da padne, a
        # pogadjanje bi izdalo robu pogresnom coveku.
        {"KooperantID": "KOOP-TEST-IME", "Ime": "Prvi", "Prezime": "Testni",
         "Mesto": "Test Mesto", "StanicaID": STANICA, "Aktivan": STATUS_AKTIVAN},
        # IME SA DIJAKRITIKOM (smoke 28.08.2026, Platni nalozi): prava imena
        # nose kvake, a operater na DE/EN tastaturi kuca bez njih -- pretraga
        # je "ne radila". Svi ostali fixture kooperanti su ASCII, pa tvrdnja
        # "ASCII upit nalazi dijakriticno ime" bez ovog reda nema nad cim da
        # padne. Ima i tekuci racun + blok (OTK-NAL-DJ), da red bude u listi.
        {"KooperantID": "KOOP-NAL-DJ", "Ime": "Đorđe", "Prezime": "Šarčević",
         "Mesto": "Test Mesto", "StanicaID": STANICA, "Aktivan": STATUS_AKTIVAN,
         "TekuciRacun": "205-0000000999888-77"},
        # EKRAN IZVESTAJI: kooperant sa neraspodeljenim avansom (NOV-IZV-A1,
        # OMID=STANICA2) a BEZ ijednog otkup bloka. ReportIsplata("OM") tako
        # dobija kanal VirmanAvans, a avans pool ekrana Platni nalozi se ne
        # pomera -- pool se sabira po kooperantima OTVORENIH blokova, kojih
        # ovaj nema.
        {"KooperantID": "KOOP-IZV-AV", "Ime": "Avram", "Prezime": "Avansni",
         "Mesto": "Test Mesto", "StanicaID": STANICA2, "Aktivan": STATUS_AKTIVAN},
    ],
    "tblParcele": [
        {"ParcelaID": "PAR-TEST-1", "KooperantID": "KOOP-TEST-1", "KatBroj": "1001",
         "KatOpstina": "Test Opstina", "Kultura": VRSTA, "PovrsinaHa": 1.5,
         "Aktivna": STATUS_AKTIVAN},
        {"ParcelaID": "PAR-TEST-2", "KooperantID": "KOOP-TEST-2", "KatBroj": "1002",
         "KatOpstina": "Test Opstina", "Kultura": VRSTA, "PovrsinaHa": 2.25,
         "Aktivna": STATUS_AKTIVAN},
    ],
    # Druga zbirna je STORNIRANA i postoji samo zbog ekrana Oporavak: lista
    # ciljeva prevezivanja sme da nudi iskljucivo AKTIVNE dokumente. Bez
    # storniranog reda ta tvrdnja nema nad cim da padne (sabotaza
    # "oporavak-stornirani-cilj" je nad starim fixture-om ostajala zelena).
    "tblZbirna": [
        {"ZbirnaID": "ZBI-KASK-1", "Datum": FIXTURE_DATE, "VozacID": VOZAC,
         "BrojZbirne": ZBIRNA_KASK, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "UkupnoKolicina": 300, "TipAmbalaze": AMB_12_1, "UkupnoAmbalaze": 30,
         "Klasa": "I", "KupacID": KUPAC},
        {"ZbirnaID": "ZBI-KASK-2", "Datum": FIXTURE_DATE, "VozacID": VOZAC2,
         "BrojZbirne": ZBIRNA_KASK, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "UkupnoKolicina": 400, "TipAmbalaze": AMB_12_1, "UkupnoAmbalaze": 40,
         "Klasa": "I", "KupacID": KUPAC},
        # ISTI BrojZbirne, ISTI kupac, DVA vozaca -> u jezgru dva dokumenta.
        # Ciljna lista Oporavka ih je spajala u jedan red jer je vlasnikom
        # smatrala samo kupca, pa operater nije mogao da izabere pravi.
        # Mirna, jednoznacna zbirna: cilj relinka u stale scenariju.
        # Ciljna zbirna, dva vlasnika ISTOG broja: B aktivan, A storniran.
        # UkupnoKolicina zaglavlja B je namerno = kolicina SAMO njegovog deteta,
        # da kontaminacija (dete A + dete B) bude vidljiva kao promena broja.
        #
        # REDOSLED JE DEO SCENARIJA i ne sme se menjati. Aktivan vlasnik mora biti
        # PRVI red tog broja: ReassignPrijemnicaToZbirna_TX bez generacije cita
        # Stornirano PRVOG reda po broju, pa bi sa storniranim prvim slucajno
        # odbio relink -- ne zato sto proverava vlasnistvo, nego zato sto je prvi
        # red slucajno bio storniran. Test bi tada bio zelen bez pokrica.
        {"ZbirnaID": "ZBI-TGT-B", "Datum": FIXTURE_DATE, "VozacID": VOZAC,
         "BrojZbirne": ZBIRNA_TGT, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "UkupnoKolicina": 100, "TipAmbalaze": AMB_12_1, "UkupnoAmbalaze": 10,
         "Klasa": "I", "KupacID": KUPAC},
        {"ZbirnaID": "ZBI-TGT-A", "Datum": FIXTURE_DATE, "VozacID": VOZAC2,
         "BrojZbirne": ZBIRNA_TGT, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "UkupnoKolicina": 300, "TipAmbalaze": AMB_12_1, "UkupnoAmbalaze": 30,
         "Klasa": "I", "KupacID": KUPAC, "Stornirano": "Da"},
        # Izvorna zbirna: jednoznacna, da test meri BAS cilj a ne izvor.
        {"ZbirnaID": "ZBI-OLDU-1", "Datum": FIXTURE_DATE, "VozacID": VOZAC,
         "BrojZbirne": ZBIRNA_OLDU, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "UkupnoKolicina": 100, "TipAmbalaze": AMB_12_1, "UkupnoAmbalaze": 10,
         "Klasa": "I", "KupacID": KUPAC},
        {"ZbirnaID": "ZBI-STL-1", "Datum": FIXTURE_DATE, "VozacID": VOZAC,
         "BrojZbirne": ZBIRNA_STALE, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "UkupnoKolicina": 200, "TipAmbalaze": AMB_12_1, "UkupnoAmbalaze": 20,
         "Klasa": "I", "KupacID": KUPAC},
        {"ZbirnaID": "ZBI-TEST-4", "Datum": FIXTURE_DATE, "VozacID": VOZAC,
         "BrojZbirne": ZBIRNA_MIRNA, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "UkupnoKolicina": 300, "TipAmbalaze": AMB_12_1, "UkupnoAmbalaze": 30,
         "Klasa": "I", "KupacID": KUPAC},
        {"ZbirnaID": "ZBI-DUPL-1", "Datum": FIXTURE_DATE, "VozacID": VOZAC,
         "BrojZbirne": ZBIRNA_DUPL, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "UkupnoKolicina": 100, "TipAmbalaze": AMB_12_1, "UkupnoAmbalaze": 10,
         "Klasa": "I", "KupacID": KUPAC},
        {"ZbirnaID": "ZBI-DUPL-2", "Datum": FIXTURE_DATE, "VozacID": VOZAC2,
         "BrojZbirne": ZBIRNA_DUPL, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "UkupnoKolicina": 200, "TipAmbalaze": AMB_12_1, "UkupnoAmbalaze": 20,
         "Klasa": "I", "KupacID": KUPAC},
        {"ZbirnaID": "ZBI-TEST-1", "Datum": FIXTURE_DATE, "VozacID": VOZAC,
         "BrojZbirne": ZBIRNA, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "UkupnoKolicina": 1000, "TipAmbalaze": AMB_12_1, "UkupnoAmbalaze": 100,
         "Klasa": "I"},
        {"ZbirnaID": "ZBI-TEST-2", "Datum": FIXTURE_DATE, "VozacID": VOZAC,
         "BrojZbirne": ZBIRNA2, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "UkupnoKolicina": 950, "TipAmbalaze": AMB_12_1, "UkupnoAmbalaze": 95,
         "Klasa": "I"},
        {"ZbirnaID": "ZBI-TEST-STOR", "Datum": FIXTURE_DATE, "VozacID": VOZAC,
         "BrojZbirne": ZBIRNA_STORNIRANA, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "UkupnoKolicina": 500, "TipAmbalaze": AMB_12_1, "UkupnoAmbalaze": 50,
         "Klasa": "I", "Stornirano": "Da"},
        # SLEDLJIVOST vozila -- v. blok konstanti SLED_* gore.
        {"ZbirnaID": "ZBI-SLED-1", "Datum": FIXTURE_DATE, "VozacID": VOZAC2,
         "BrojZbirne": SLED_ZBIRNA, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "UkupnoKolicina": SLED_KG_1 + SLED_KG_2, "Klasa": "I", "KupacID": KUPAC},
        {"ZbirnaID": "ZBI-SLED-N", "Datum": FIXTURE_DATE, "VozacID": VOZAC2,
         "BrojZbirne": SLED_ZBIRNA_N, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "UkupnoKolicina": 150, "Klasa": "I", "KupacID": KUPAC2},
        {"ZbirnaID": "ZBI-SLED-R", "Datum": FIXTURE_DATE, "VozacID": VOZAC2,
         "BrojZbirne": SLED_ZBIRNA_R, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "UkupnoKolicina": 250, "Klasa": "I", "KupacID": KUPAC2},
        # Krug 8: F-lanac (ALL fakturisanost) i M-lanac (SearchRefs).
        {"ZbirnaID": "ZBI-SLED-F", "Datum": FIXTURE_DATE, "VozacID": VOZAC2,
         "BrojZbirne": SLED_ZBIRNA_F, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "UkupnoKolicina": SLED_KG_F, "Klasa": "I", "KupacID": KUPAC2},
        {"ZbirnaID": "ZBI-SLED-M", "Datum": FIXTURE_DATE, "VozacID": VOZAC2,
         "BrojZbirne": SLED_ZBIRNA_M, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "UkupnoKolicina": SLED_KG_M, "Klasa": "I", "KupacID": KUPAC2},
        # GP grana (v. blok konstanti SLED_ZBIRNA_G/H/K gore).
        {"ZbirnaID": "ZBI-SLED-G", "Datum": FIXTURE_DATE, "VozacID": VOZAC2,
         "BrojZbirne": SLED_ZBIRNA_G, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "UkupnoKolicina": SLED_KG_G, "Klasa": "I", "KupacID": KUPAC2},
        {"ZbirnaID": "ZBI-SLED-H", "Datum": FIXTURE_DATE, "VozacID": VOZAC2,
         "BrojZbirne": SLED_ZBIRNA_H, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "UkupnoKolicina": SLED_KG_H, "Klasa": "I", "KupacID": KUPAC2},
        {"ZbirnaID": "ZBI-SLED-K", "Datum": FIXTURE_DATE, "VozacID": VOZAC2,
         "BrojZbirne": SLED_ZBIRNA_K, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "UkupnoKolicina": SLED_KG_K, "Klasa": "I", "KupacID": KUPAC2},
        {"ZbirnaID": "ZBI-SLED-P", "Datum": FIXTURE_DATE, "VozacID": VOZAC2,
         "BrojZbirne": SLED_ZBIRNA_P, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "UkupnoKolicina": SLED_KG_P, "Klasa": "I", "KupacID": KUPAC2},
        # Dvosmislen par za sledljivost (v. SLED_ZBIRNA_D): dva aktivna
        # vlasnika (dva vozaca) dele broj; nijedan drugi test ih ne dira.
        {"ZbirnaID": "ZBI-SLED-D1", "Datum": FIXTURE_DATE, "VozacID": VOZAC,
         "BrojZbirne": SLED_ZBIRNA_D, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "UkupnoKolicina": 100, "Klasa": "I", "KupacID": KUPAC2},
        {"ZbirnaID": "ZBI-SLED-D2", "Datum": FIXTURE_DATE, "VozacID": VOZAC2,
         "BrojZbirne": SLED_ZBIRNA_D, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "UkupnoKolicina": 100, "Klasa": "I", "KupacID": KUPAC2},
    ],
    # Tri slucaja koje zadatak trazi:
    #   OTP-TEST-1  datum iz proslosti + poznata zbirna + ostatak != 0 (1000 - 400)
    #   OTP-TEST-2  bez zbirne
    #   OTP-TEST-3  bez zbirne, ali blok u tblOtkup nosi zbirnu (ZB-TEST-3)
    "tblOtpremnica": [
        {"OtpremnicaID": "OTP-LEG-A", "Datum": FIXTURE_DATE, "StanicaID": STANICA,
         "VozacID": VOZAC, "BrojOtpremnice": OTPREMNICA_LEGACY, "BrojZbirne": "",
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 100, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 10, "Klasa": "I"},
        {"OtpremnicaID": "OTP-LEG-B", "Datum": FIXTURE_DATE, "StanicaID": STANICA2,
         "VozacID": VOZAC, "BrojOtpremnice": OTPREMNICA_LEGACY, "BrojZbirne": "",
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 200, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 20, "Klasa": "I"},
        {"OtpremnicaID": "OTP-LEG-N", "Datum": FIXTURE_DATE, "StanicaID": STANICA,
         "VozacID": VOZAC, "BrojOtpremnice": OTPREMNICA_ZAMENA, "BrojZbirne": "",
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 100, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 10, "Klasa": "I"},
        # Dve otpremnice ISTOG broja sa RAZLICITIH stanica u ISTOJ zbirni.
        # Zbirna je po invarijanti zbir svih svojih otpremnica, pa je ovo
        # normalno stanje -- a "jedini vlasnik" po distinct BROJU tu laze.
        {"OtpremnicaID": "OTP-KOL-A", "Datum": FIXTURE_DATE, "StanicaID": STANICA,
         "VozacID": VOZAC, "BrojOtpremnice": OTPREMNICA_KOLIZIJA, "BrojZbirne": ZBIRNA_KASK,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 100, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 10, "Klasa": "I"},
        {"OtpremnicaID": "OTP-KOL-B", "Datum": FIXTURE_DATE, "StanicaID": STANICA2,
         "VozacID": VOZAC, "BrojOtpremnice": OTPREMNICA_KOLIZIJA, "BrojZbirne": ZBIRNA_KASK,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 200, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 20, "Klasa": "I"},
        # ZATECEN CONTEXT: redosled je deo scenarija. OTP-STL-B je PRVI red tog
        # broja i visi na JEDNOZNACNOJ zbirni; izabrana OTP-STL-A je druga i visi
        # na DVOSMISLENOJ. Ne menjati redosled -- test 46 meri bas to da kod ne
        # sme da uzme prvog po broju.
        {"OtpremnicaID": "OTP-OLD-U", "Datum": FIXTURE_DATE, "StanicaID": STANICA,
         "VozacID": VOZAC, "BrojOtpremnice": OTPREMNICA_OLD_U, "BrojZbirne": ZBIRNA_OLDU,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 100, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 10, "Klasa": "I"},
        # AKTIVNO dete STORNIRANOG vlasnika cilja -- srce scenarija.
        {"OtpremnicaID": "OTP-HIST", "Datum": FIXTURE_DATE, "StanicaID": STANICA2,
         "VozacID": VOZAC2, "BrojOtpremnice": OTPREMNICA_HIST, "BrojZbirne": ZBIRNA_TGT,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 300, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 30, "Klasa": "I"},
        {"OtpremnicaID": "OTP-NEW-T", "Datum": FIXTURE_DATE, "StanicaID": STANICA,
         "VozacID": VOZAC, "BrojOtpremnice": OTPREMNICA_NEW_T, "BrojZbirne": ZBIRNA_TGT,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 100, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 10, "Klasa": "I"},
        {"OtpremnicaID": "OTP-BLK-A", "Datum": FIXTURE_DATE, "StanicaID": STANICA,
         "VozacID": VOZAC, "BrojOtpremnice": OTPREMNICA_BLOK, "BrojZbirne": "",
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 100, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 10, "Klasa": "I"},
        # STORNIRANA, ali njen blok ostaje aktivan (recovery stanje).
        {"OtpremnicaID": "OTP-BLK-B", "Datum": FIXTURE_DATE, "StanicaID": STANICA2,
         "VozacID": VOZAC, "BrojOtpremnice": OTPREMNICA_BLOK, "BrojZbirne": "",
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 200, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 20, "Klasa": "I",
         "Stornirano": "Da"},
        {"OtpremnicaID": "OTP-STL-B", "Datum": FIXTURE_DATE, "StanicaID": STANICA2,
         "VozacID": VOZAC, "BrojOtpremnice": OTPREMNICA_STALE, "BrojZbirne": ZBIRNA_STALE,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 200, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 20, "Klasa": "I"},
        {"OtpremnicaID": "OTP-STL-A", "Datum": FIXTURE_DATE, "StanicaID": STANICA,
         "VozacID": VOZAC, "BrojOtpremnice": OTPREMNICA_STALE, "BrojZbirne": ZBIRNA_KASK,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 100, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 10, "Klasa": "I"},
        # Cilj zamene. Zbirna mora biti NEPRAZNA i jednoznacna: relink prijemnica
        # se radi samo kad nova zbirna postoji, pa bi prazna napravila placebo
        # test -- tudja prijemnica se ne bi prevezala ni bez kapije.
        {"OtpremnicaID": "OTP-STL-N", "Datum": FIXTURE_DATE, "StanicaID": STANICA,
         "VozacID": VOZAC, "BrojOtpremnice": OTPREMNICA_STALE_NOVA, "BrojZbirne": ZBIRNA_STALE,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 100, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 10, "Klasa": "I"},
        {"OtpremnicaID": "OTP-TEST-1", "Datum": FIXTURE_DATE, "StanicaID": STANICA,
         "VozacID": VOZAC, "BrojOtpremnice": "1/TEST", "BrojZbirne": ZBIRNA,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 1000, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 100, "Klasa": "I"},
        {"OtpremnicaID": "OTP-TEST-2", "Datum": FIXTURE_DATE, "StanicaID": STANICA,
         "VozacID": VOZAC, "BrojOtpremnice": "2/TEST", "BrojZbirne": "",
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 500, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 50, "Klasa": "I"},
        {"OtpremnicaID": "OTP-TEST-3", "Datum": FIXTURE_DATE, "StanicaID": STANICA,
         "VozacID": VOZAC, "BrojOtpremnice": "3/TEST", "BrojZbirne": "",
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 800, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 80, "Klasa": "I"},
        # Otpremnica cija zbirna (ZB-TEST-4) nosi NESTORNIRANE prijemnice --
        # do sada nijedna nije postojala (par. 23.12/S10), pa se linija
        # "prijemnica + kupac" u detalju otpremnice/otkupa nije mogla
        # tvrditi. ZBIRNA_MIRNA je namerno: njene prijemnice (PRJ-FAK-1/2/3,
        # PRJ-TEST-Z2) su vec vozila fakturisanja i ne diraju se.
        {"OtpremnicaID": "OTP-IZV-Z", "Datum": FIXTURE_DATE, "StanicaID": STANICA,
         "VozacID": VOZAC, "BrojOtpremnice": "Z/TEST", "BrojZbirne": ZBIRNA_MIRNA,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 700, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 70, "Klasa": "I"},
        # SLEDLJIVOST vozila -- v. blok konstanti SLED_* gore.
        {"OtpremnicaID": "OTP-SLED-1", "Datum": FIXTURE_DATE, "StanicaID": STANICA2,
         "VozacID": VOZAC2, "BrojOtpremnice": "31/TEST", "BrojZbirne": SLED_ZBIRNA,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": SLED_KG_1 + SLED_KG_2,
         "Cena": 50.0, "Klasa": "I"},
        # VozacID PRAZAN namerno: broj ZB-TEST-SLDD dele dva vozaca, pa se bez
        # vozaca otpremnice vlasnik ne moze razresiti (fail-closed vozilo).
        {"OtpremnicaID": "OTP-SLED-D", "Datum": FIXTURE_DATE, "StanicaID": STANICA2,
         "VozacID": "", "BrojOtpremnice": "32/TEST", "BrojZbirne": SLED_ZBIRNA_D,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 100,
         "Cena": 50.0, "Klasa": "I"},
        {"OtpremnicaID": "OTP-SLED-N", "Datum": FIXTURE_DATE, "StanicaID": STANICA2,
         "VozacID": VOZAC2, "BrojOtpremnice": "33/TEST", "BrojZbirne": SLED_ZBIRNA_N,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 150,
         "Cena": 50.0, "Klasa": "I"},
        # Krug 8: F i M lanci (v. blok konstanti).
        {"OtpremnicaID": "OTP-SLED-F", "Datum": FIXTURE_DATE, "StanicaID": STANICA2,
         "VozacID": VOZAC2, "BrojOtpremnice": "36/TEST", "BrojZbirne": SLED_ZBIRNA_F,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": SLED_KG_F,
         "Cena": 50.0, "Klasa": "I"},
        # GP grana: tri lanca G/H/K (v. blok konstanti gore).
        {"OtpremnicaID": "OTP-SLED-G", "Datum": FIXTURE_DATE, "StanicaID": STANICA2,
         "VozacID": VOZAC2, "BrojOtpremnice": "38/TEST", "BrojZbirne": SLED_ZBIRNA_G,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": SLED_KG_G,
         "Klasa": "I"},
        {"OtpremnicaID": "OTP-SLED-H", "Datum": FIXTURE_DATE, "StanicaID": STANICA2,
         "VozacID": VOZAC2, "BrojOtpremnice": "39/TEST", "BrojZbirne": SLED_ZBIRNA_H,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": SLED_KG_H,
         "Klasa": "I"},
        {"OtpremnicaID": "OTP-SLED-K", "Datum": FIXTURE_DATE, "StanicaID": STANICA2,
         "VozacID": VOZAC2, "BrojOtpremnice": "40/TEST", "BrojZbirne": SLED_ZBIRNA_K,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": SLED_KG_K,
         "Klasa": "I"},
        {"OtpremnicaID": "OTP-SLED-P", "Datum": FIXTURE_DATE, "StanicaID": STANICA2,
         "VozacID": VOZAC2, "BrojOtpremnice": "50/TEST", "BrojZbirne": SLED_ZBIRNA_P,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": SLED_KG_P,
         "Klasa": "I"},
        {"OtpremnicaID": "OTP-SLED-M", "Datum": FIXTURE_DATE, "StanicaID": STANICA2,
         "VozacID": VOZAC2, "BrojOtpremnice": "37/TEST", "BrojZbirne": SLED_ZBIRNA_M,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": SLED_KG_M,
         "Cena": 50.0, "Klasa": "I"},
        # 250 kg nad blokom od 100 kg -- kg curi na prvoj karici (namerno).
        {"OtpremnicaID": "OTP-SLED-R", "Datum": FIXTURE_DATE, "StanicaID": STANICA2,
         "VozacID": VOZAC2, "BrojOtpremnice": "34/TEST", "BrojZbirne": SLED_ZBIRNA_R,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 250,
         "Cena": 50.0, "Klasa": "I"},
    ],
    "tblOtkup": [
        # Po jedan blok na svakoj legacy otpremnici -- zavrsetak ispravke sme da
        # preveze SAMO blok dokumenta ciji je OldDocID u context-u.
        {"OtkupID": "OTK-LEG-A", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-1",
         "StanicaID": STANICA, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": 100, "Cena": 50.0, "TipAmbalaze": AMB_12_1,
         "KolAmbalaze": 10, "VozacID": VOZAC, "BrojDokumenta": "L1/TEST", "Klasa": "I",
         "OtpremnicaID": "OTP-LEG-A", "BrojOtpremnice": OTPREMNICA_LEGACY},
        {"OtkupID": "OTK-LEG-B", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-2",
         "StanicaID": STANICA2, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": 200, "Cena": 50.0, "TipAmbalaze": AMB_12_1,
         "KolAmbalaze": 20, "VozacID": VOZAC, "BrojDokumenta": "L2/TEST", "Klasa": "I",
         "OtpremnicaID": "OTP-LEG-B", "BrojOtpremnice": OTPREMNICA_LEGACY},
        # Isti BrojDokumenta na DVA otkupna mesta -- legitimno, broj je scoped
        # po stanici. Bez generacije storno po broju bi zahvatio oba.
        {"OtkupID": "OTK-KOL-A", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-1",
         "StanicaID": STANICA, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": 100, "Cena": 50.0, "TipAmbalaze": AMB_12_1,
         "KolAmbalaze": 10, "VozacID": VOZAC, "BrojDokumenta": OTKUP_KOLIZIJA, "Klasa": "I"},
        {"OtkupID": "OTK-KOL-B", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-2",
         "StanicaID": STANICA2, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": 200, "Cena": 50.0, "TipAmbalaze": AMB_12_1,
         "KolAmbalaze": 20, "VozacID": VOZAC, "BrojDokumenta": OTKUP_KOLIZIJA, "Klasa": "I"},
        # Po jedan AKTIVAN blok na svakoj od dve otpremnice istog broja.
        {"OtkupID": "OTK-BLK-A", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-1",
         "StanicaID": STANICA, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": 100, "Cena": 50.0, "TipAmbalaze": AMB_12_1,
         "KolAmbalaze": 10, "VozacID": VOZAC, "BrojDokumenta": "B1/TEST", "Klasa": "I",
         "OtpremnicaID": "OTP-BLK-A", "BrojOtpremnice": OTPREMNICA_BLOK},
        {"OtkupID": "OTK-BLK-B", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-2",
         "StanicaID": STANICA2, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": 200, "Cena": 50.0, "TipAmbalaze": AMB_12_1,
         "KolAmbalaze": 20, "VozacID": VOZAC, "BrojDokumenta": "B2/TEST", "Klasa": "I",
         "OtpremnicaID": "OTP-BLK-B", "BrojOtpremnice": OTPREMNICA_BLOK},
        {"OtkupID": "OTK-TEST-1", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-1",
         "StanicaID": STANICA, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": 400, "Cena": 50.0, "TipAmbalaze": AMB_12_1,
         "KolAmbalaze": 40, "VozacID": VOZAC, "BrojDokumenta": "1/TEST",
         "Klasa": "I", "BrojZbirne": ZBIRNA, "OtpremnicaID": "OTP-TEST-1",
         "BrojOtpremnice": "1/TEST", "ParcelaID": "PAR-TEST-1"},
        {"OtkupID": "OTK-TEST-2", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-2",
         "StanicaID": STANICA, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": 200, "Cena": 50.0, "TipAmbalaze": AMB_12_1,
         "KolAmbalaze": 20, "VozacID": VOZAC, "BrojDokumenta": "3/TEST",
         "Klasa": "I", "BrojZbirne": ZBIRNA_U_BLOKU, "OtpremnicaID": "OTP-TEST-3",
         "BrojOtpremnice": "3/TEST", "ParcelaID": "PAR-TEST-2"},
        # TRI otvorene stavke ISTOG bloka, isti kooperant. Poziv na broj iz
        # izvoda ga razresava jednoznacno (jedan kooperant = jedan pogodak), ali
        # GetOtkupCandidatesForKooperantBlock preko MAX_BLOK_KANDIDATA dize
        # ERR_BMAP_MANUAL_REQUIRED -- red koji je "spreman po jakom kljucu" a
        # automatski se ipak NE moze zavrsiti. Bez ovog para stanja se cip
        # "jaki kljucevi" i stvarni ishod auto-mapiranja ne mogu razlikovati.
        {"OtkupID": "OTK-BIM-3A", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-3",
         "StanicaID": STANICA, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": 20, "Cena": 50.0, "TipAmbalaze": AMB_12_1,
         "KolAmbalaze": 2, "VozacID": VOZAC, "BrojDokumenta": BIM_BLOK_3, "Klasa": "I"},
        {"OtkupID": "OTK-BIM-3B", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-3",
         "StanicaID": STANICA, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": 20, "Cena": 40.0, "TipAmbalaze": AMB_12_1,
         "KolAmbalaze": 2, "VozacID": VOZAC, "BrojDokumenta": BIM_BLOK_3, "Klasa": "II"},
        {"OtkupID": "OTK-BIM-3C", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-3",
         "StanicaID": STANICA, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": 20, "Cena": 30.0, "TipAmbalaze": AMB_12_1,
         "KolAmbalaze": 2, "VozacID": VOZAC, "BrojDokumenta": BIM_BLOK_3, "Klasa": "I"},
        # ISTI kooperant, ISTI broj bloka, DVE stanice -- v. BIM_BLOK_OM.
        {"OtkupID": "OTK-BIM-OMA", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-3",
         "StanicaID": STANICA, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": 10, "Cena": 50.0, "TipAmbalaze": AMB_12_1,
         "KolAmbalaze": 1, "VozacID": VOZAC, "BrojDokumenta": BIM_BLOK_OM, "Klasa": "I"},
        {"OtkupID": "OTK-BIM-OMB", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-3",
         "StanicaID": STANICA2, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": 10, "Cena": 60.0, "TipAmbalaze": AMB_12_1,
         "KolAmbalaze": 1, "VozacID": VOZAC, "BrojDokumenta": BIM_BLOK_OM, "Klasa": "I"},
        # ISTI blok, ali BEZ upisanog otkupnog mesta. Danasnji pisci ovo odbijaju
        # (SaveOtkup / SaveOtkupMulti_TX traze StanicaID), pa je red legacy oblik
        # -- a zatecene sveske takve redove imaju (v. datum 26062026). Bez njega
        # se ne moze izmeriti da rucno mapiranje STAJE umesto da posalje prazan
        # scope i raspodeli novac preko sva tri otkupna mesta.
        {"OtkupID": "OTK-BIM-OMX", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-3",
         "StanicaID": "", "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": 10, "Cena": 70.0, "TipAmbalaze": AMB_12_1,
         "KolAmbalaze": 1, "VozacID": VOZAC, "BrojDokumenta": BIM_BLOK_OM, "Klasa": "I"},
        # Blok koji je u CELOSTI placen -- v. BIM_BLOK_PLACEN i red u tblNovac.
        {"OtkupID": BIM_OTK_PLACEN, "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-3",
         "StanicaID": STANICA, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": 10, "Cena": 50.0, "TipAmbalaze": AMB_12_1,
         "KolAmbalaze": 1, "VozacID": VOZAC, "BrojDokumenta": BIM_BLOK_PLACEN, "Klasa": "I"},
        # DELIMICNO ISPLACEN blok (1000, od cega 400 kroz NOV-NAL-DELIM):
        # ekran Platni nalozi mora da pokaze otvoreno = ostatak (600), ne pun
        # iznos -- bez ovog reda bi "otvoreno < ukupno" bilo nemerljivo, jer su
        # svi ostali otvoreni blokovi bez ijedne knjizene isplate.
        {"OtkupID": "OTK-NAL-DELIM", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-1",
         "StanicaID": STANICA, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": 20, "Cena": 50.0, "TipAmbalaze": AMB_12_1,
         "KolAmbalaze": 2, "VozacID": VOZAC, "BrojDokumenta": "NAL1/TEST", "Klasa": "I"},
        # Blok kooperanta sa dijakriticnim imenom -- v. KOOP-NAL-DJ.
        {"OtkupID": "OTK-NAL-DJ", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-NAL-DJ",
         "StanicaID": STANICA, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": 10, "Cena": 50.0, "TipAmbalaze": AMB_12_1,
         "KolAmbalaze": 1, "VozacID": VOZAC, "BrojDokumenta": "NAL2/TEST", "Klasa": "I"},
        # STORNIRAN blok sa "otvorenim" iznosom: ne sme ni u listu ni u naloge
        # (ExcludeStornirano u GetOpenOtkupi). Bez njega bi tvrdnja "storniran
        # nije u listi" merila odsustvo reda, ne filter.
        {"OtkupID": "OTK-NAL-STOR", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-1",
         "StanicaID": STANICA, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": 10, "Cena": 50.0, "TipAmbalaze": AMB_12_1,
         "KolAmbalaze": 1, "VozacID": VOZAC, "BrojDokumenta": "NALX/TEST", "Klasa": "I",
         "Stornirano": "Da"},
        # EKRAN IZVESTAJI: blok U CELOSTI placen kroz DVA kanala (NOV-IZV-K1
        # kes 200 + NOV-IZV-V1 virman 300 = 500) -- vozilo za slaganje
        # ReportIsplata (Ukupno = Kes + VirmanFirma + VirmanAvans po redu).
        # Zatvoren je NAMERNO: GetOpenOtkupi ga ne vidi, pa KPI/cipovi/korpa
        # ekrana Platni nalozi ostaju bit-identicni (T129/T130 tvrde relacije
        # nad otvorenima).
        # ORPHAN STANICA (recenzija #245/krug 16): otkup na stanici koje
        # NEMA u tblStanice -- zbirni "Svi OM" mora da je prikaze pod ID-em,
        # ne da je tiho izostavi (univerzum iz podataka, ne iz sifarnika).
        {"OtkupID": "OTK-ORPH-1", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-2",
         "StanicaID": "STA-ORPHAN", "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": 10, "Cena": 50.0, "TipAmbalaze": AMB_12_1,
         "KolAmbalaze": 1, "VozacID": VOZAC, "BrojDokumenta": "ORPH/1", "Klasa": "I"},
        {"OtkupID": OTK_IZV_ZATVOREN, "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-2",
         "StanicaID": STANICA2, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": 10, "Cena": 50.0, "TipAmbalaze": AMB_12_1,
         "KolAmbalaze": 1, "VozacID": VOZAC, "BrojDokumenta": "IZV1/TEST", "Klasa": "I"},
        # EKRAN SLEDLJIVOST -- v. blok konstanti SLED_* gore. Svi blokovi su
        # ZATVORENI (NOV-SLED-* redovi), pa Platni nalozi ne vide nista novo.
        # Dva kooperanta na ISTOJ otpremnici: "nazad" pitanje (od fakture ka
        # kooperantima) mora da vrati OBA.
        {"OtkupID": "OTK-SLED-1", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-2",
         "StanicaID": STANICA2, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": SLED_KG_1, "Cena": 50.0,
         "VozacID": VOZAC2, "BrojDokumenta": "S1/TEST", "Klasa": "I",
         "BrojZbirne": SLED_ZBIRNA, "OtpremnicaID": "OTP-SLED-1",
         "BrojOtpremnice": "31/TEST", "ParcelaID": "PAR-TEST-2"},
        # BEZ parcele -- vozilo za oznaku "bez parcele" na listi PARCELE.
        # KOOP-TEST-IME, ne KOOP-TEST-3 (v. komentar bloka konstanti).
        {"OtkupID": "OTK-SLED-2", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-IME",
         "StanicaID": STANICA2, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": SLED_KG_2, "Cena": 50.0,
         "VozacID": VOZAC2, "BrojDokumenta": "S2/TEST", "Klasa": "I",
         "BrojZbirne": SLED_ZBIRNA, "OtpremnicaID": "OTP-SLED-1",
         "BrojOtpremnice": "31/TEST"},
        {"OtkupID": "OTK-SLED-D", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-2",
         "StanicaID": STANICA2, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": 100, "Cena": 50.0,
         "VozacID": VOZAC2, "BrojDokumenta": "S3/TEST", "Klasa": "I",
         "BrojZbirne": SLED_ZBIRNA_D, "OtpremnicaID": "OTP-SLED-D",
         "BrojOtpremnice": "32/TEST"},
        {"OtkupID": "OTK-SLED-N", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-2",
         "StanicaID": STANICA2, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": 150, "Cena": 50.0,
         "VozacID": VOZAC2, "BrojDokumenta": "S4/TEST", "Klasa": "I",
         "BrojZbirne": SLED_ZBIRNA_N, "OtpremnicaID": "OTP-SLED-N",
         "BrojOtpremnice": "33/TEST"},
        # Krug 8: F-lanac (bez parcele; KOOP-TEST-IME -- v. komentar bloka
        # konstanti) i M-lanac (KOOP-TEST-2 sa parcelom -> PARCELE red).
        {"OtkupID": "OTK-SLED-F", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-IME",
         "StanicaID": STANICA2, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": SLED_KG_F, "Cena": 50.0,
         "VozacID": VOZAC2, "BrojDokumenta": "S6/TEST", "Klasa": "I",
         "BrojZbirne": SLED_ZBIRNA_F, "OtpremnicaID": "OTP-SLED-F",
         "BrojOtpremnice": "36/TEST"},
        {"OtkupID": "OTK-SLED-M", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-2",
         "StanicaID": STANICA2, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": SLED_KG_M, "Cena": 50.0,
         "VozacID": VOZAC2, "BrojDokumenta": "S7/TEST", "Klasa": "I",
         "BrojZbirne": SLED_ZBIRNA_M, "OtpremnicaID": "OTP-SLED-M",
         "BrojOtpremnice": "37/TEST", "ParcelaID": "PAR-TEST-2"},
        {"OtkupID": "OTK-SLED-R", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-2",
         "StanicaID": STANICA2, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": 100, "Cena": 50.0,
         "VozacID": VOZAC2, "BrojDokumenta": "S5/TEST", "Klasa": "I",
         "BrojZbirne": SLED_ZBIRNA_R, "OtpremnicaID": "OTP-SLED-R",
         "BrojOtpremnice": "34/TEST", "ParcelaID": "PAR-TEST-2"},
        # GP grana: G/H/K lanci (v. blok konstanti gore).
        {"OtkupID": "OTK-SLED-G", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-2",
         "StanicaID": STANICA2, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": SLED_KG_G, "Cena": 50.0,
         "VozacID": VOZAC2, "BrojDokumenta": "S8/TEST", "Klasa": "I",
         "BrojZbirne": SLED_ZBIRNA_G, "OtpremnicaID": "OTP-SLED-G",
         "BrojOtpremnice": "38/TEST", "ParcelaID": "PAR-TEST-2"},
        {"OtkupID": "OTK-SLED-H", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-2",
         "StanicaID": STANICA2, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": SLED_KG_H, "Cena": 50.0,
         "VozacID": VOZAC2, "BrojDokumenta": "S9/TEST", "Klasa": "I",
         "BrojZbirne": SLED_ZBIRNA_H, "OtpremnicaID": "OTP-SLED-H",
         "BrojOtpremnice": "39/TEST"},
        {"OtkupID": "OTK-SLED-K", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-2",
         "StanicaID": STANICA2, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": SLED_KG_K, "Cena": 50.0,
         "VozacID": VOZAC2, "BrojDokumenta": "S10/TEST", "Klasa": "I",
         "BrojZbirne": SLED_ZBIRNA_K, "OtpremnicaID": "OTP-SLED-K",
         "BrojOtpremnice": "40/TEST"},
        {"OtkupID": "OTK-SLED-P", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-2",
         "StanicaID": STANICA2, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": SLED_KG_P, "Cena": 50.0,
         "VozacID": VOZAC2, "BrojDokumenta": "S11/TEST", "Klasa": "I",
         "BrojZbirne": SLED_ZBIRNA_P, "OtpremnicaID": "OTP-SLED-P",
         "BrojOtpremnice": "50/TEST"},
    ],
    # Jedna faktura, samo zato da kapija UplataFakturaProblem ima nad cim da
    # radi: vlasnistvo (KupacID), trenutni preostali iznos (Iznos - uplate) i
    # razlika "postoji / ne postoji". Namerno samo tri kolone -- sejanje ide PO
    # IMENU, pa svaka dodatna kolona koju donor nema obara generator.
    # Druga faktura je BEZ IZNOSA i postoji zbog jednog pravila koje se lako
    # izgubi: faktura kojoj iznos nije evidentiran NE SME da blokira uplatu.
    # To je razlog zbog koga je kapija u UplataFakturaProblem uopste uslovna;
    # bez ovog reda popravka te kapije mogla bi tiho da ukine i to pravilo.
    "tblNovac": [
        # Zatvara BIM_OTK_PLACEN u celosti: GetUplataForOtkup sabira Isplata po
        # OtkupID-u, pa je "otvoreno" tacno nula i kandidata nema.
        {"NovacID": "NOV-BIM-PLAC", "BrojDokumenta": "BIM-PLAC-1",
         "Datum": FIXTURE_DATE, "Tip": "VirmanFirmaKoop",
         "Isplata": BIM_OTK_PLACEN_IZNOS, "KooperantID": "KOOP-TEST-3",
         "OtkupID": BIM_OTK_PLACEN},
        # Delimicna isplata bloka OTK-NAL-DELIM (400 od 1000) -- druga polovina
        # para; v. komentar uz taj red u tblOtkup.
        {"NovacID": "NOV-NAL-DELIM", "BrojDokumenta": "NAL-PLAC-1",
         "Datum": FIXTURE_DATE, "Tip": "VirmanFirmaKoop",
         "Isplata": 400, "KooperantID": "KOOP-TEST-1",
         "OtkupID": "OTK-NAL-DELIM"},
        {"NovacID": "NOV-TEST-D1", "BrojDokumenta": NOVAC_DUPLI_BROJ,
         "Datum": FIXTURE_DATE, "Tip": "VirmanAvansKoop", "Isplata": 1000,
         "KooperantID": "KOOP-TEST-1"},
        {"NovacID": "NOV-TEST-D2", "BrojDokumenta": NOVAC_DUPLI_BROJ,
         "Datum": FIXTURE_DATE, "Tip": "VirmanAvansKoop", "Isplata": 2000,
         "KooperantID": "KOOP-TEST-2"},
        # ODBITAK AGRO DUGA. Pravilo zivi u DVE implementacije: GetAgroAbzug
        # (po kooperantu, zove ga kes ekrana) i GetAgroAbzugMapa (jednim
        # prolazom, zove je lista dugova). Dok u fixture-u nije bilo nijednog
        # AgroAbzug reda, obe su vracale nulu i tvrdnja da se slazu nije imala
        # nad cim da padne.
        #
        # DVA reda za KOOP-TEST-1 (300 + 200 = 500) -- da se vidi da se
        # SABIRA, a ne da pobedjuje poslednji.
        {"NovacID": "NOV-TEST-AB1", "BrojDokumenta": "AB-1",
         "Datum": FIXTURE_DATE, "Tip": "AgroAbzug", "Uplata": 300,
         "KooperantID": "KOOP-TEST-1"},
        {"NovacID": "NOV-TEST-AB2", "BrojDokumenta": "AB-2",
         "Datum": FIXTURE_DATE, "Tip": "AgroAbzug", "Uplata": 200,
         "KooperantID": "KOOP-TEST-1"},
        # Drugi kooperant -- mapa mora da razdvaja, ne da sve slije u jedan zbir.
        {"NovacID": "NOV-TEST-AB3", "BrojDokumenta": "AB-3",
         "Datum": FIXTURE_DATE, "Tip": "AgroAbzug", "Uplata": 100,
         "KooperantID": "KOOP-TEST-2"},
        # STORNIRAN odbitak: obe implementacije ga izuzimaju (ExcludeStornirano).
        # Da jedna prestane, zbir KOOP-TEST-1 postaje 1499 i test pukne.
        {"NovacID": "NOV-TEST-AB4", "BrojDokumenta": "AB-4",
         "Datum": FIXTURE_DATE, "Tip": "AgroAbzug", "Uplata": 999,
         "KooperantID": "KOOP-TEST-1", "Stornirano": "Da"},
        # Uplata DRUGOG tipa istom kooperantu -- ni jedna ni druga je ne broje.
        {"NovacID": "NOV-TEST-AB5", "BrojDokumenta": "AB-5",
         "Datum": FIXTURE_DATE, "Tip": "UplataKoop", "Uplata": 777,
         "KooperantID": "KOOP-TEST-1"},
        # UPLATA PO FAKTURI. Jedina u fixture-u koja nosi FakturaID, pa je i
        # jedina koju BuildUplataDictByFaktura ima sta da sabere. Zatvara
        # FAKTURA_PLAC tacno na iznos -- ni dinar vise, da se "placeno" i
        # "preplaceno" ne mesaju.
        {"NovacID": "NOV-TEST-UF1", "BrojDokumenta": "UF-1",
         "Datum": FIXTURE_DATE, "Tip": "KupciUplata", "Uplata": FAKTURA_PLAC_IZNOS,
         "PartnerID": KUPAC, "Partner": KUPAC, "FakturaID": FAKTURA_PLAC},
        # EKRAN IZVESTAJI -- cetiri reda sa OMID-om (v. blok konstanti gore).
        # Do sada NIJEDAN red nije nosio OMID, pa su ReportIsplata("OM"),
        # GetOMAvansSaldo i atribucija isplate stanici (NovacRedPripadaStanici)
        # radili nad praznim skupom.
        #
        # Kes i virman ZATVARAJU OTK-IZV-1 (200 + 300 = 500 = vrednost bloka):
        # tri kanala u istom izvestaju, nula novih otvorenih blokova.
        {"NovacID": "NOV-IZV-K1", "BrojDokumenta": "IZV-K1",
         "Datum": FIXTURE_DATE, "Tip": "KesOtkupacKoop", "Isplata": IZV_KES,
         "KooperantID": "KOOP-TEST-2", "OMID": STANICA2,
         "OtkupID": OTK_IZV_ZATVOREN},
        {"NovacID": "NOV-IZV-V1", "BrojDokumenta": "IZV-V1",
         "Datum": FIXTURE_DATE, "Tip": "VirmanFirmaKoop", "Isplata": IZV_VIRMAN,
         "KooperantID": "KOOP-TEST-2", "OMID": STANICA2,
         "OtkupID": OTK_IZV_ZATVOREN},
        # VirmanAvans kanal: kooperant BEZ blokova (v. KOOP-IZV-AV), pa avans
        # pool Platnih naloga ostaje netaknut.
        {"NovacID": "NOV-IZV-A1", "BrojDokumenta": "IZV-A1",
         "Datum": FIXTURE_DATE, "Tip": "VirmanAvansKoop", "Isplata": IZV_AVANS_KOOP,
         "KooperantID": "KOOP-IZV-AV", "OMID": STANICA2},
        # Avans Firma -> Otkupac (bez kooperanta): red "OM AVANS (nerasporedjen)"
        # u SaldoOM i kontrolni redovi ReportIsplata. Sve na STANICA2:
        # T_WriterGuard_AvansSaldoOM trazi da STA-TEST-1 NEMA avans salda
        # (preduslov 0), pa vozilo izvestaja zivi na drugoj stanici.
        # GetOMAvansSaldo(STANICA2)
        # = 5000 - 200 (kes podeljen kooperantu) = 4800 -- prvi put != 0.
        {"NovacID": "NOV-IZV-FA", "BrojDokumenta": "IZV-FA",
         "Datum": FIXTURE_DATE, "Tip": "KesFirmaOtkupac", "Isplata": IZV_OM_AVANS,
         "OMID": STANICA2},
        # EKRAN SLEDLJIVOST: pet virmana ZATVARA svih pet SLED blokova u
        # celosti (kg * 50), pa GetOpenOtkupi ne vidi nijedan i Platni nalozi
        # ostaju bit-identicni. VirmanFirmaKoop ne dira avans pool; OMID je
        # STANICA2 (novi novcani redovi samo na STA-TEST-2 --
        # T_WriterGuard_AvansSaldoOM trazi STA-TEST-1 saldo 0).
        {"NovacID": "NOV-SLED-1", "BrojDokumenta": "SLED-P1",
         "Datum": FIXTURE_DATE, "Tip": "VirmanFirmaKoop", "Isplata": SLED_KG_1 * 50,
         "KooperantID": "KOOP-TEST-2", "OMID": STANICA2, "OtkupID": "OTK-SLED-1"},
        {"NovacID": "NOV-SLED-2", "BrojDokumenta": "SLED-P2",
         "Datum": FIXTURE_DATE, "Tip": "VirmanFirmaKoop", "Isplata": SLED_KG_2 * 50,
         "KooperantID": "KOOP-TEST-IME", "OMID": STANICA2, "OtkupID": "OTK-SLED-2"},
        # Krug 8: zatvaranje F i M blokova (ista pravila kao NOV-SLED-1/2).
        {"NovacID": "NOV-SLED-F", "BrojDokumenta": "SLED-P6",
         "Datum": FIXTURE_DATE, "Tip": "VirmanFirmaKoop", "Isplata": SLED_KG_F * 50,
         "KooperantID": "KOOP-TEST-IME", "OMID": STANICA2, "OtkupID": "OTK-SLED-F"},
        {"NovacID": "NOV-SLED-M", "BrojDokumenta": "SLED-P7",
         "Datum": FIXTURE_DATE, "Tip": "VirmanFirmaKoop", "Isplata": SLED_KG_M * 50,
         "KooperantID": "KOOP-TEST-2", "OMID": STANICA2, "OtkupID": "OTK-SLED-M"},
        {"NovacID": "NOV-SLED-D", "BrojDokumenta": "SLED-P3",
         "Datum": FIXTURE_DATE, "Tip": "VirmanFirmaKoop", "Isplata": 5000,
         "KooperantID": "KOOP-TEST-2", "OMID": STANICA2, "OtkupID": "OTK-SLED-D"},
        {"NovacID": "NOV-SLED-N", "BrojDokumenta": "SLED-P4",
         "Datum": FIXTURE_DATE, "Tip": "VirmanFirmaKoop", "Isplata": 7500,
         "KooperantID": "KOOP-TEST-2", "OMID": STANICA2, "OtkupID": "OTK-SLED-N"},
        # GP grana: zatvaranje G/H/K blokova (isto pravilo kao NOV-SLED-1/2).
        {"NovacID": "NOV-SLED-G", "BrojDokumenta": "SLED-P8",
         "Datum": FIXTURE_DATE, "Tip": "VirmanFirmaKoop", "Isplata": SLED_KG_G * 50,
         "KooperantID": "KOOP-TEST-2", "OMID": STANICA2, "OtkupID": "OTK-SLED-G"},
        {"NovacID": "NOV-SLED-H", "BrojDokumenta": "SLED-P9",
         "Datum": FIXTURE_DATE, "Tip": "VirmanFirmaKoop", "Isplata": SLED_KG_H * 50,
         "KooperantID": "KOOP-TEST-2", "OMID": STANICA2, "OtkupID": "OTK-SLED-H"},
        {"NovacID": "NOV-SLED-K", "BrojDokumenta": "SLED-P10",
         "Datum": FIXTURE_DATE, "Tip": "VirmanFirmaKoop", "Isplata": SLED_KG_K * 50,
         "KooperantID": "KOOP-TEST-2", "OMID": STANICA2, "OtkupID": "OTK-SLED-K"},
        {"NovacID": "NOV-SLED-PP", "BrojDokumenta": "SLED-P11",
         "Datum": FIXTURE_DATE, "Tip": "VirmanFirmaKoop", "Isplata": SLED_KG_P * 50,
         "KooperantID": "KOOP-TEST-2", "OMID": STANICA2, "OtkupID": "OTK-SLED-P"},
        {"NovacID": "NOV-SLED-R", "BrojDokumenta": "SLED-P5",
         "Datum": FIXTURE_DATE, "Tip": "VirmanFirmaKoop", "Isplata": 5000,
         "KooperantID": "KOOP-TEST-2", "OMID": STANICA2, "OtkupID": "OTK-SLED-R"},
    ],
    # AMBALAZNI LEDGER -- do sada PRAZAN (nije ni u KEEP_ROWS ni u SEED-u), pa
    # bi svaka ambalazna tvrdnja ekrana Izvestaji bila zelena nad praznim
    # skupom. Redovi su SAMOSTALNA kretanja (DokumentID nije otkupID): bas njih
    # ReportKarticaKooperanta uzima kao amb redove kartice, a uz-otkup parove
    # vec pokrivaju kolone tblOtkup.
    "tblAmbalaza": [
        # Revers REV-IZV-1: OM izdao KOOP-TEST-1 prazne gajbe, DVE NOGE ISTOG
        # DokumentID-a (Kooperant Ulaz + Stanica Izlaz) -- oblik koji
        # rekonstrukcija reversa trazi. Uz 12/1 ide i DRUGI TIP na ISTOM
        # dokumentu (AMB_LETVA): kljuc reversa je DokumentID + DokumentTip +
        # TIP AMBALAZE, pa pregled mora da ih drzi u DVA reda (AUD-012); bez
        # ovog para bi se spajanje tipova vratilo neprimeceno.
        {"AmbID": "AMB-IZV-K1", "Datum": FIXTURE_DATE, "TipAmbalaze": AMB_12_1,
         "Kolicina": 30, "Smer": "Ulaz", "EntitetID": "KOOP-TEST-1",
         "EntitetTip": "Kooperant", "DokumentID": "REV-IZV-1",
         "DokumentTip": "OM-Izlaz-Koop"},
        {"AmbID": "AMB-IZV-S1", "Datum": FIXTURE_DATE, "TipAmbalaze": AMB_12_1,
         "Kolicina": 30, "Smer": "Izlaz", "EntitetID": STANICA,
         "EntitetTip": "Stanica", "DokumentID": "REV-IZV-1",
         "DokumentTip": "OM-Izlaz-Koop"},
        {"AmbID": "AMB-IZV-K2", "Datum": FIXTURE_DATE, "TipAmbalaze": AMB_LETVA,
         "Kolicina": 5, "Smer": "Ulaz", "EntitetID": "KOOP-TEST-1",
         "EntitetTip": "Kooperant", "DokumentID": "REV-IZV-1",
         "DokumentTip": "OM-Izlaz-Koop"},
        {"AmbID": "AMB-IZV-S2", "Datum": FIXTURE_DATE, "TipAmbalaze": AMB_LETVA,
         "Kolicina": 5, "Smer": "Izlaz", "EntitetID": STANICA,
         "EntitetTip": "Stanica", "DokumentID": "REV-IZV-1",
         "DokumentTip": "OM-Izlaz-Koop"},
        # Povrat REV-IZV-2: kooperant vratio 10 gajbi -> saldo kooperanta
        # 30 + 5 - 10 = 25; kartica ambalaze ima i Ulaz i Izlaz redove.
        {"AmbID": "AMB-IZV-K3", "Datum": FIXTURE_DATE, "TipAmbalaze": AMB_12_1,
         "Kolicina": 10, "Smer": "Izlaz", "EntitetID": "KOOP-TEST-1",
         "EntitetTip": "Kooperant", "DokumentID": "REV-IZV-2",
         "DokumentTip": "OM-Ulaz-Koop"},
        {"AmbID": "AMB-IZV-S3", "Datum": FIXTURE_DATE, "TipAmbalaze": AMB_12_1,
         "Kolicina": 10, "Smer": "Ulaz", "EntitetID": STANICA,
         "EntitetTip": "Stanica", "DokumentID": "REV-IZV-2",
         "DokumentTip": "OM-Ulaz-Koop"},
        # STORNIRAN red sa velikom kolicinom: i izvestaj i kanonski saldo
        # (GetAmbalazeStanje) ga izuzimaju (ExcludeStornirano). Da jedna
        # strana prestane, saldo kooperanta postane 124 i slaganje pukne --
        # isti obrazac kao NOV-TEST-AB4.
        {"AmbID": "AMB-IZV-KS", "Datum": FIXTURE_DATE, "TipAmbalaze": AMB_12_1,
         "Kolicina": 99, "Smer": "Ulaz", "EntitetID": "KOOP-TEST-1",
         "EntitetTip": "Kooperant", "DokumentID": "REV-IZV-X",
         "DokumentTip": "OM-Izlaz-Koop", "Stornirano": "Da"},
        # Ulaz od firme na OM: lista AMBALAZA za Stanicu ima i Ulaz i Izlaz
        # redove, pa cipovi ulaz/izlaz ne mere prazan skup.
        {"AmbID": "AMB-IZV-S4", "Datum": FIXTURE_DATE, "TipAmbalaze": AMB_12_1,
         "Kolicina": 100, "Smer": "Ulaz", "EntitetID": STANICA,
         "EntitetTip": "Stanica", "DokumentID": "REV-IZV-3",
         "DokumentTip": "OM-Ulaz-Firma"},
        # KUPAC red: lista AMBALAZA za Kupca nije prazna. DokumentTip
        # Prijemnica -> ResolveDokBroj razresava broj iz tblPrijemnica.
        {"AmbID": "AMB-IZV-KP1", "Datum": FIXTURE_DATE, "TipAmbalaze": AMB_12_1,
         "Kolicina": 10, "Smer": "Ulaz", "EntitetID": KUPAC,
         "EntitetTip": "Kupac", "DokumentID": "PRJ-FAK-3",
         "DokumentTip": "Prijemnica"},
        # UZ-OTKUP PAROVI za KOOP-TEST-1 -- onako kako ih SaveOtkup pise
        # (primljene pune gajbe: Kooperant-Izlaz + Stanica-Ulaz, DokTip=Otkup,
        # DokumentID = otkupID). Bez njih su kartica kooperanta (tblOtkup
        # kolone + samostalna kretanja) i kanonski ledger saldo
        # (GetAmbalazeStanje) dva read-modela nad NEKONZISTENTNOM sveskom i
        # slaganje nema smisla -- prvi crveni run T_Izv_SlaganjeKartica je
        # bio tacno to (kartica -47, ledger 25). Parovi idu za SVAKI
        # nestorniran otkup KOOP-TEST-1 sa KolAmbalaze > 0; storniran
        # OTK-NAL-STOR se preskace (storno flow bi stornirao i par).
        {"AmbID": "AMB-OTK-K1A", "Datum": FIXTURE_DATE, "TipAmbalaze": AMB_12_1,
         "Kolicina": 10, "Smer": "Izlaz", "EntitetID": "KOOP-TEST-1",
         "EntitetTip": "Kooperant", "DokumentID": "OTK-LEG-A", "DokumentTip": "Otkup"},
        {"AmbID": "AMB-OTK-S1A", "Datum": FIXTURE_DATE, "TipAmbalaze": AMB_12_1,
         "Kolicina": 10, "Smer": "Ulaz", "EntitetID": STANICA,
         "EntitetTip": "Stanica", "DokumentID": "OTK-LEG-A", "DokumentTip": "Otkup"},
        {"AmbID": "AMB-OTK-K1B", "Datum": FIXTURE_DATE, "TipAmbalaze": AMB_12_1,
         "Kolicina": 10, "Smer": "Izlaz", "EntitetID": "KOOP-TEST-1",
         "EntitetTip": "Kooperant", "DokumentID": "OTK-KOL-A", "DokumentTip": "Otkup"},
        {"AmbID": "AMB-OTK-S1B", "Datum": FIXTURE_DATE, "TipAmbalaze": AMB_12_1,
         "Kolicina": 10, "Smer": "Ulaz", "EntitetID": STANICA,
         "EntitetTip": "Stanica", "DokumentID": "OTK-KOL-A", "DokumentTip": "Otkup"},
        {"AmbID": "AMB-OTK-K1C", "Datum": FIXTURE_DATE, "TipAmbalaze": AMB_12_1,
         "Kolicina": 10, "Smer": "Izlaz", "EntitetID": "KOOP-TEST-1",
         "EntitetTip": "Kooperant", "DokumentID": "OTK-BLK-A", "DokumentTip": "Otkup"},
        {"AmbID": "AMB-OTK-S1C", "Datum": FIXTURE_DATE, "TipAmbalaze": AMB_12_1,
         "Kolicina": 10, "Smer": "Ulaz", "EntitetID": STANICA,
         "EntitetTip": "Stanica", "DokumentID": "OTK-BLK-A", "DokumentTip": "Otkup"},
        {"AmbID": "AMB-OTK-K1D", "Datum": FIXTURE_DATE, "TipAmbalaze": AMB_12_1,
         "Kolicina": 40, "Smer": "Izlaz", "EntitetID": "KOOP-TEST-1",
         "EntitetTip": "Kooperant", "DokumentID": "OTK-TEST-1", "DokumentTip": "Otkup"},
        {"AmbID": "AMB-OTK-S1D", "Datum": FIXTURE_DATE, "TipAmbalaze": AMB_12_1,
         "Kolicina": 40, "Smer": "Ulaz", "EntitetID": STANICA,
         "EntitetTip": "Stanica", "DokumentID": "OTK-TEST-1", "DokumentTip": "Otkup"},
        {"AmbID": "AMB-OTK-K1E", "Datum": FIXTURE_DATE, "TipAmbalaze": AMB_12_1,
         "Kolicina": 2, "Smer": "Izlaz", "EntitetID": "KOOP-TEST-1",
         "EntitetTip": "Kooperant", "DokumentID": "OTK-NAL-DELIM", "DokumentTip": "Otkup"},
        {"AmbID": "AMB-OTK-S1E", "Datum": FIXTURE_DATE, "TipAmbalaze": AMB_12_1,
         "Kolicina": 2, "Smer": "Ulaz", "EntitetID": STANICA,
         "EntitetTip": "Stanica", "DokumentID": "OTK-NAL-DELIM", "DokumentTip": "Otkup"},
        # VOZAC ruta (filter po VozacID koloni): utovar na otpremnici pa
        # predaja na prijemnici -- kompletna ruta, vozacev saldo 0. DokTip
        # Otkup se NE koristi (vozacki izvestaj ga izuzima).
        {"AmbID": "AMB-IZV-V1", "Datum": FIXTURE_DATE, "TipAmbalaze": AMB_12_1,
         "Kolicina": 40, "Smer": "Izlaz", "EntitetID": STANICA,
         "EntitetTip": "Stanica", "VozacID": VOZAC, "DokumentID": "OTP-TEST-1",
         "DokumentTip": "Otpremnica"},
        {"AmbID": "AMB-IZV-V2", "Datum": FIXTURE_DATE, "TipAmbalaze": AMB_12_1,
         "Kolicina": 40, "Smer": "Ulaz", "EntitetID": KUPAC,
         "EntitetTip": "Kupac", "VozacID": VOZAC, "DokumentID": "PRJ-TEST-A",
         "DokumentTip": "Prijemnica"},
    ],
    "tblFakture": [
        {"FakturaID": FAKTURA, "KupacID": KUPAC, "Iznos": FAKTURA_IZNOS},
        {"FakturaID": FAKTURA_BEZ_IZNOSA, "KupacID": KUPAC, "Iznos": 0},
        # Neplacena faktura DRUGOG kupca: bez nje bi tvrdnja "cip neplacenih se
        # slaze sa GetOpenFakture" merila slaganje dva prazna skupa.
        {"FakturaID": FAKTURA_NEPL, "BrojFakture": "2/2026", "Datum": FIXTURE_DATE,
         "KupacID": KUPAC2, "Iznos": FAKTURA_NEPL_IZNOS, "Status": "Neplaceno"},
        # Placena u celosti (uplata je u tblNovac). Cip "placene" i znak u redu
        # (paypill) moraju da je vide isto.
        {"FakturaID": FAKTURA_PLAC, "BrojFakture": "3/2026", "Datum": FIXTURE_DATE,
         "KupacID": KUPAC, "Iznos": FAKTURA_PLAC_IZNOS, "Status": "Placeno"},
        # STORNIRANA: citac je izbacuje (ExcludeStornirano). Da prestane, pojavila
        # bi se u listi i operater bi joj nudio stampu i SEF.
        {"FakturaID": FAKTURA_STORNO, "BrojFakture": "4/2026", "Datum": FIXTURE_DATE,
         "KupacID": KUPAC, "Iznos": 7000, "Status": "Neplaceno", "Stornirano": "Da"},
        # EKRAN SLEDLJIVOST: kraj potpunog lanca (PRJ-SLED-1 nosi ovaj ID).
        # Iznos = 500 kg * 50; Neplaceno je sveze, posteno stanje -- naplata
        # je tudji tok i ne dira lanac.
        {"FakturaID": SLED_FAKTURA, "BrojFakture": "5/2026", "Datum": FIXTURE_DATE,
         "KupacID": KUPAC, "Iznos": (SLED_KG_1 + SLED_KG_2) * 50,
         "Status": "Neplaceno"},
        # Krug 8: aktivne fakture F i M lanaca (R1/R2). FAK-NEMA-GA se
        # NAMERNO ne dodaje -- PRJ-SLED-F2 pokazuje u prazno.
        {"FakturaID": SLED_FAKTURA_2, "BrojFakture": "6/2026", "Datum": FIXTURE_DATE,
         "KupacID": KUPAC2, "Iznos": 60 * 50, "Status": "Neplaceno"},
        {"FakturaID": SLED_FAKTURA_3, "BrojFakture": "7/2026", "Datum": FIXTURE_DATE,
         "KupacID": KUPAC2, "Iznos": 100 * 50, "Status": "Neplaceno"},
        {"FakturaID": SLED_FAKTURA_3B, "BrojFakture": "8/2026", "Datum": FIXTURE_DATE,
         "KupacID": KUPAC2, "Iznos": 100 * 50, "Status": "Neplaceno"},
        # GP grana: AKTIVNA zavrsna faktura gotove robe (PRE-SLED-G pokazuje
        # na nju; stavka FST-SLED-GP je DOKAZ prodajne veze -- kanonski
        # contract B2). FAK-NEMA-GP se NAMERNO ne dodaje -- PRE-SLED-K
        # pokazuje u prazno (kontradikcija).
        {"FakturaID": SLED_FAKTURA_GP, "BrojFakture": "9/2026", "Datum": FIXTURE_DATE,
         "KupacID": KUPAC2, "Iznos": SLED_GP_IZLAZ_KG * SLED_GP_CENA,
         "Status": "Neplaceno"},
        # AKTIVNA faktura BEZ stavke prerade PRE-GP-B2 koja na nju
        # pokazuje -- marker bez stavke = "faktura neusaglasena" (B2).
        {"FakturaID": "FAK-SLED-GP2", "BrojFakture": "10/2026", "Datum": FIXTURE_DATE,
         "KupacID": KUPAC2, "Iznos": 4000, "Status": "Neplaceno"},
        # Krug 5: faktura DELIMICNE prodaje P lanca (50 kg x 200 preko
        # utovara UT-SLED-P).
        {"FakturaID": "FAK-SLED-GP4", "BrojFakture": "12/2026", "Datum": FIXTURE_DATE,
         "KupacID": KUPAC2, "Iznos": 50 * 200, "Status": "Neplaceno"},
    ],
    # Prve GP stavke fakture u fixture-u: podatkovni DOKAZ da faktura
    # 9/2026 STVARNO sadrzi preradu G (PreradaID/BrojPrerade su
    # ENSURE_COLS kolone). PrijemnicaID/Klasa/BrojPrijemnice prazni --
    # GP stavka nema prijemnicu (isti oblik kao writer).
    "tblFakturaStavke": [
        {"StavkaID": "FST-SLED-GP", "FakturaID": SLED_FAKTURA_GP,
         "PrijemnicaID": "", "Kolicina": SLED_GP_IZLAZ_KG, "Cena": SLED_GP_CENA,
         "Klasa": "", "BrojPrijemnice": "",
         "PreradaID": "PRE-SLED-G", "BrojPrerade": "51/2026",
         "UtovarID": "UT-SLED-G"},
        # Krug 5, P lanac: parcijalna prodaja -- 50 od 120 kg, preko
        # utovara UT-SLED-P na fakturu 12/2026.
        {"StavkaID": "FST-SLED-P", "FakturaID": "FAK-SLED-GP4",
         "PrijemnicaID": "", "Kolicina": 50, "Cena": 200,
         "Klasa": "", "BrojPrijemnice": "",
         "PreradaID": "PRE-SLED-P", "BrojPrerade": "181/2026",
         "UtovarID": "UT-SLED-P"},
        # Krug 5, SB siroce: aktivna PRODAJNA stavka bez utovara --
        # fizicka isporuka ne postoji, veza je neusaglasena.
        {"StavkaID": "FST-GP-SB", "FakturaID": "FAK-SLED-GP2",
         "PrijemnicaID": "", "Kolicina": 5, "Cena": 150,
         "Klasa": "", "BrojPrijemnice": "",
         "PreradaID": "PRE-GP-SB", "BrojPrerade": "171/2026"},
    ],
    # UTOVARNE LISTE (krug 5): dokument fizicke isporuke. G/K/P na
    # lancima; B2/WL/WM/OV su contract negativi (v. komentar prerada).
    "tblUtovar": [
        # G nosi i PUNE podatke prevoza (krug 5d) -- vozilo za test
        # profesionalnog obrasca.
        {"UtovarID": "UT-SLED-G", "BrojUtovara": 1, "Godina": 2026,
         "DatumUtovara": FIXTURE_DATE, "KupacID": KUPAC2,
         "Fakturisano": "Da", "FakturaID": SLED_FAKTURA_GP,
         "Prevoznik": "Test prevoz DOO", "Vozac": "Petar Petrovic",
         "Registracija": "BG-123-AB", "Plomba": "PL-0042",
         "TemperaturniRezim": "-18 C", "MestoIstovara": "Beograd, Skladiste 2",
         "VremeUtovara": "08:30", "BrojNarudzbenice": "PO-7788",
         "Napomena": "fixture prevoz"},
        {"UtovarID": "UT-SLED-K", "BrojUtovara": 2, "Godina": 2026,
         "DatumUtovara": FIXTURE_DATE, "KupacID": KUPAC2,
         "Fakturisano": "Da", "FakturaID": "FAK-NEMA-GP"},
        {"UtovarID": "UT-SLED-P", "BrojUtovara": 3, "Godina": 2026,
         "DatumUtovara": FIXTURE_DATE, "KupacID": KUPAC2,
         "Fakturisano": "Da", "FakturaID": "FAK-SLED-GP4"},
        {"UtovarID": "UT-GP-B2", "BrojUtovara": 4, "Godina": 2026,
         "DatumUtovara": FIXTURE_DATE, "KupacID": KUPAC2,
         "Fakturisano": "Da", "FakturaID": "FAK-SLED-GP2"},
        {"UtovarID": "UT-GP-WL", "BrojUtovara": 5, "Godina": 2026,
         "DatumUtovara": FIXTURE_DATE, "KupacID": KUPAC2,
         "FakturaID": "FAK-STALE-GP"},
        {"UtovarID": "UT-GP-WM", "BrojUtovara": 6, "Godina": 2026,
         "DatumUtovara": FIXTURE_DATE, "KupacID": KUPAC2,
         "Fakturisano": "Da"},
        {"UtovarID": "UT-GP-OV", "BrojUtovara": 7, "Godina": 2026,
         "DatumUtovara": FIXTURE_DATE, "KupacID": KUPAC2},
    ],
    "tblUtovarStavke": [
        {"UtovarStavkaID": "UTS-SLED-G", "UtovarID": "UT-SLED-G",
         "PreradaID": "PRE-SLED-G", "BrojPrerade": "51/2026",
         "KolicinaKg": SLED_GP_IZLAZ_KG},
        {"UtovarStavkaID": "UTS-SLED-K", "UtovarID": "UT-SLED-K",
         "PreradaID": "PRE-SLED-K", "BrojPrerade": "61/2026",
         "KolicinaKg": 56},
        {"UtovarStavkaID": "UTS-SLED-P", "UtovarID": "UT-SLED-P",
         "PreradaID": "PRE-SLED-P", "BrojPrerade": "181/2026",
         "KolicinaKg": 50},
        {"UtovarStavkaID": "UTS-GP-B2", "UtovarID": "UT-GP-B2",
         "PreradaID": "PRE-GP-B2", "BrojPrerade": "101/2026",
         "KolicinaKg": 10},
        {"UtovarStavkaID": "UTS-GP-WL", "UtovarID": "UT-GP-WL",
         "PreradaID": "PRE-GP-WL", "BrojPrerade": "111/2026",
         "KolicinaKg": 5},
        {"UtovarStavkaID": "UTS-GP-WM", "UtovarID": "UT-GP-WM",
         "PreradaID": "PRE-GP-WM", "BrojPrerade": "151/2026",
         "KolicinaKg": 5},
        {"UtovarStavkaID": "UTS-GP-OV", "UtovarID": "UT-GP-OV",
         "PreradaID": "PRE-GP-OV", "BrojPrerade": "191/2026",
         "KolicinaKg": 30},
    ],
    # STAVKE IZVODA -- devet redova u TRI izvoda, svaki sa svojim razlogom:
    #   BIM-FIX-1   jak kljuc preko FAKTURE (poziv na broj = broj fakture 2/2026)
    #   BIM-FIX-2   jak kljuc preko BLOKA   (poziv na broj = BrojDokumenta 1/TEST)
    #   BIM-FIX-3   bez ijednog jakog kljuca -> trazi rucno mapiranje
    #   BIM-FIX-K   DRUGI racun pod ISTIM brojem izvoda -> kolizija broja izvoda
    #   BIM-FIX-3K  blok sa 3 otvorene stavke -> ERR_BMAP_MANUAL_REQUIRED
    #   BIM-FIX-ER  Obradjeno="Error" -- auto pokusao i odbio (cip "za rucno")
    #   BIM-FIX-DA  Obradjeno="Da"    -- obradjen, van reda za mapiranje
    #   BIM-FIX-SK  Obradjeno="Skip"  -- preskocen
    #   BIM-FIX-ST  Stornirano="Da"   -- ne sme ni u jednu listu
    #
    # PartnerKonto je svuda prazan namerno: fixture nema tekuce racune u
    # sifarnicima, pa bi grana "jak kljuc preko racuna" bila lazno zelena.
    # Zbirovi izvoda (UkupanDuguje / UkupanPotrazuje) se slazu sa zbirom
    # NESTORNIRANIH redova tog izvoda -- tako ih parser i upisuje.
    "tblBankaImport": [
        # --- Izvod 1, racun 1: pocetno 10000 + 1500 - 2000 = 9500 (SLAZE SE)
        {"BankaImportID": "BIM-FIX-1", "BrojDokumenta": BIM_IZVOD_1,
         "DatumIzvoda": BIM_DATUM_1, "BrojRacuna": BIM_RACUN_1,
         "DatumTransakcije": BIM_DATUM_1, "Partner": "Kupac Prvi doo",
         "PartnerKonto": "", "Opis": "Uplata po fakturi", "Uplata": 1000, "Isplata": 0,
         "Valuta": "RSD", "PozivNaBroj": "2/2026", "SvrhaPlacanja": "Uplata po fakturi",
         "BankaReferenz": "REF-FIX-1", "IzvorFajl": "fixture.pdf",
         "ImportVreme": BIM_DATUM_1, "Obradjeno": "",
         "PocetnoStanje": 10000, "ZavrsnoStanje": 9500,
         "UkupanDuguje": 2000, "UkupanPotrazuje": 1500},
        {"BankaImportID": "BIM-FIX-2", "BrojDokumenta": BIM_IZVOD_1,
         "DatumIzvoda": BIM_DATUM_1, "BrojRacuna": BIM_RACUN_1,
         "DatumTransakcije": BIM_DATUM_1, "Partner": "Prvi Testni",
         "PartnerKonto": "", "Opis": "Isplata po bloku", "Uplata": 0, "Isplata": 2000,
         "Valuta": "RSD", "PozivNaBroj": "1/TEST", "SvrhaPlacanja": "Isplata kooperantu",
         "BankaReferenz": "REF-FIX-2", "IzvorFajl": "fixture.pdf",
         "ImportVreme": BIM_DATUM_1, "Obradjeno": "",
         "PocetnoStanje": 10000, "ZavrsnoStanje": 9500,
         "UkupanDuguje": 2000, "UkupanPotrazuje": 1500},
        {"BankaImportID": "BIM-FIX-3", "BrojDokumenta": BIM_IZVOD_1,
         "DatumIzvoda": BIM_DATUM_1, "BrojRacuna": BIM_RACUN_1,
         "DatumTransakcije": BIM_DATUM_1, "Partner": "Nepoznat Platilac",
         "PartnerKonto": "", "Opis": "Bez poziva na broj", "Uplata": 500, "Isplata": 0,
         "Valuta": "RSD", "PozivNaBroj": "", "SvrhaPlacanja": "",
         "BankaReferenz": "REF-FIX-3", "IzvorFajl": "fixture.pdf",
         "ImportVreme": BIM_DATUM_1, "Obradjeno": "",
         "PocetnoStanje": 10000, "ZavrsnoStanje": 9500,
         "UkupanDuguje": 2000, "UkupanPotrazuje": 1500},
        # --- Izvod 1, racun 2: ISTI broj izvoda, drugi racun. Pocetno 5000 + 700 = 5700
        {"BankaImportID": "BIM-FIX-K", "BrojDokumenta": BIM_IZVOD_1,
         "DatumIzvoda": BIM_DATUM_1, "BrojRacuna": BIM_RACUN_2,
         "DatumTransakcije": BIM_DATUM_1, "Partner": "Drugi Platilac",
         "PartnerKonto": "", "Opis": "Uplata na drugi racun", "Uplata": 700, "Isplata": 0,
         "Valuta": "RSD", "PozivNaBroj": "", "SvrhaPlacanja": "",
         "BankaReferenz": "REF-FIX-K", "IzvorFajl": "fixture2.pdf",
         "ImportVreme": BIM_DATUM_1, "Obradjeno": "",
         "PocetnoStanje": 5000, "ZavrsnoStanje": 5700,
         "UkupanDuguje": 0, "UkupanPotrazuje": 700},
        # --- Izvod 2, racun 1: 8000 + 950 - 3000 = 5950, a upisano je 6050 -> RAZLIKA 100
        {"BankaImportID": "BIM-FIX-3K", "BrojDokumenta": BIM_IZVOD_2,
         "DatumIzvoda": BIM_DATUM_2, "BrojRacuna": BIM_RACUN_1,
         "DatumTransakcije": BIM_DATUM_2, "Partner": "Treci Testni",
         "PartnerKonto": "", "Opis": "Isplata po bloku sa 3 stavke",
         "Uplata": 0, "Isplata": 3000,
         "Valuta": "RSD", "PozivNaBroj": BIM_BLOK_3, "SvrhaPlacanja": "Isplata kooperantu",
         "BankaReferenz": "REF-FIX-3K", "IzvorFajl": "fixture3.pdf",
         "ImportVreme": BIM_DATUM_2, "Obradjeno": "",
         "PocetnoStanje": 8000, "ZavrsnoStanje": BIM_IZVOD_2_ZAVRSNO,
         "UkupanDuguje": 3000, "UkupanPotrazuje": 950},
        {"BankaImportID": "BIM-FIX-ER", "BrojDokumenta": BIM_IZVOD_2,
         "DatumIzvoda": BIM_DATUM_2, "BrojRacuna": BIM_RACUN_1,
         "DatumTransakcije": BIM_DATUM_2, "Partner": "Sporni Platilac",
         "PartnerKonto": "", "Opis": "Auto odbio", "Uplata": 250, "Isplata": 0,
         "Valuta": "RSD", "PozivNaBroj": "", "SvrhaPlacanja": "",
         "BankaReferenz": "REF-FIX-ER", "IzvorFajl": "fixture3.pdf",
         "ImportVreme": BIM_DATUM_2, "Obradjeno": "Error",
         "PocetnoStanje": 8000, "ZavrsnoStanje": BIM_IZVOD_2_ZAVRSNO,
         "UkupanDuguje": 3000, "UkupanPotrazuje": 950},
        {"BankaImportID": "BIM-FIX-DA", "BrojDokumenta": BIM_IZVOD_2,
         "DatumIzvoda": BIM_DATUM_2, "BrojRacuna": BIM_RACUN_1,
         "DatumTransakcije": BIM_DATUM_2, "Partner": "Obradjeni Platilac",
         "PartnerKonto": "", "Opis": "Vec proknjizeno", "Uplata": 400, "Isplata": 0,
         "Valuta": "RSD", "PozivNaBroj": "", "SvrhaPlacanja": "",
         "BankaReferenz": "REF-FIX-DA", "IzvorFajl": "fixture3.pdf",
         "ImportVreme": BIM_DATUM_2, "Obradjeno": "Da",
         "PocetnoStanje": 8000, "ZavrsnoStanje": BIM_IZVOD_2_ZAVRSNO,
         "UkupanDuguje": 3000, "UkupanPotrazuje": 950},
        # DATUM KOJI NIJE DATUM -- regresioni red.
        #
        # Na zatecenim svescima tblBankaImport nosi DatumTransakcije kao BROJ
        # oblika ddmmyyyy (26.06.2026 -> 26062026), a ne kao datum. Mreza nad
        # kolonom tipa "date" radi CDate, koji van opsega baca Overflow, a
        # RenderGrid to guta (On Error Resume Next) -- pa u celiji ostane natpis
        # od RANIJEG crtanja.
        #
        # Kad je ovaj red prvi put posejan, oborio je SEDAM testova, ukljucujuci
        # tudje (T_StornoEkran_SvakaListaVracaRedove). Posle ispravke u ljusci
        # (modUiData.CellDate odbija broj van opsega) suite je opet zelena, pa
        # red OSTAJE -- on je jedino sto bi povratak te greske primetilo.
        {"BankaImportID": "BIM-FIX-SK", "BrojDokumenta": BIM_IZVOD_2,
         "DatumIzvoda": BIM_DATUM_2, "BrojRacuna": BIM_RACUN_1,
         "DatumTransakcije": Sirovo(26062026), "Partner": "Preskoceni Platilac",
         "PartnerKonto": "", "Opis": "Operater preskocio", "Uplata": 300, "Isplata": 0,
         "Valuta": "RSD", "PozivNaBroj": "", "SvrhaPlacanja": "",
         "BankaReferenz": "REF-FIX-SK", "IzvorFajl": "fixture3.pdf",
         "ImportVreme": BIM_DATUM_2, "Obradjeno": "Skip",
         "PocetnoStanje": 8000, "ZavrsnoStanje": BIM_IZVOD_2_ZAVRSNO,
         "UkupanDuguje": 3000, "UkupanPotrazuje": 950},
        # ISTI broj izvoda i ISTI racun kao izvod 1, ali PRETHODNI ciklus.
        # Bez datuma u kljucu grupe, ova dva izvoda bi se spojila: saldo i datum
        # bi se uzeli sa prvog reda, a broj stavki sabrao preko oba.
        # Obradjeno="Da" da ne pomera brojku otvorenih.
        {"BankaImportID": "BIM-FIX-PY", "BrojDokumenta": BIM_IZVOD_1,
         "DatumIzvoda": BIM_DATUM_PY, "BrojRacuna": BIM_RACUN_1,
         "DatumTransakcije": BIM_DATUM_PY, "Partner": "Prosla Godina doo",
         "PartnerKonto": "", "Opis": "Izvod iz proslog ciklusa", "Uplata": 100, "Isplata": 0,
         "Valuta": "RSD", "PozivNaBroj": "", "SvrhaPlacanja": "",
         "BankaReferenz": "REF-FIX-PY", "IzvorFajl": "fixture0.pdf",
         "ImportVreme": BIM_DATUM_PY, "Obradjeno": "Da",
         "PocetnoStanje": 1000, "ZavrsnoStanje": 1100,
         "UkupanDuguje": 0, "UkupanPotrazuje": 100},
        # DVA reda istog izvoda sa RAZLICITIM zbirovima -- v. BIM_IZVOD_NES.
        # Prvi kaze zavrsno 5000, drugi 9999. Ko uzme "prvi red pobedjuje",
        # prikazace 5000 kao istinu o celom izvodu.
        {"BankaImportID": "BIM-FIX-NS1", "BrojDokumenta": BIM_IZVOD_NES,
         "DatumIzvoda": BIM_DATUM_1, "BrojRacuna": BIM_RACUN_1,
         "DatumTransakcije": BIM_DATUM_1, "Partner": "Nesaglasan doo",
         "PartnerKonto": "", "Opis": "Prvi red", "Uplata": 500, "Isplata": 0,
         "Valuta": "RSD", "PozivNaBroj": "", "SvrhaPlacanja": "",
         "BankaReferenz": "REF-NS1", "IzvorFajl": "fixture9.pdf",
         "ImportVreme": BIM_DATUM_1, "Obradjeno": "Da",
         "PocetnoStanje": 4500, "ZavrsnoStanje": 5000,
         "UkupanDuguje": 0, "UkupanPotrazuje": 500},
        {"BankaImportID": "BIM-FIX-NS2", "BrojDokumenta": BIM_IZVOD_NES,
         "DatumIzvoda": BIM_DATUM_1, "BrojRacuna": BIM_RACUN_1,
         "DatumTransakcije": BIM_DATUM_1, "Partner": "Nesaglasan doo",
         "PartnerKonto": "", "Opis": "Drugi red, drugi zbirovi",
         "Uplata": 500, "Isplata": 0,
         "Valuta": "RSD", "PozivNaBroj": "", "SvrhaPlacanja": "",
         "BankaReferenz": "REF-NS2", "IzvorFajl": "fixture9.pdf",
         # OTVOREN namerno: da oba nova reda ne budu "Da". Sa oba mapirana bi
         # broj otvorenih i broj mapiranih ispali JEDNAKI (6 i 6), pa sabotaza
         # koja znacki podmetne mapirane umesto otvorenih ne bi obarala nista --
         # dokaz bi tiho nestao. v. tvrdnja "otvorenih <> mapiranih" u modTest.
         "ImportVreme": BIM_DATUM_1, "Obradjeno": "",
         "PocetnoStanje": 4500, "ZavrsnoStanje": 9999,
         # i PROMET se razlikuje, ne samo stanje -- inace podnozje nema sta da meri
         "UkupanDuguje": 0, "UkupanPotrazuje": 700},
        # Storniran red nosi ISTE zbirove izvoda kao ostali -- da agregat ne
        # zavisi od toga koji je red grupe procitan.
        {"BankaImportID": "BIM-FIX-ST", "BrojDokumenta": BIM_IZVOD_2,
         "DatumIzvoda": BIM_DATUM_2, "BrojRacuna": BIM_RACUN_1,
         "DatumTransakcije": BIM_DATUM_2, "Partner": "Stornirani Platilac",
         "PartnerKonto": "", "Opis": "Storniran uvoz", "Uplata": 900, "Isplata": 0,
         "Valuta": "RSD", "PozivNaBroj": "", "SvrhaPlacanja": "",
         "BankaReferenz": "REF-FIX-ST", "IzvorFajl": "fixture3.pdf",
         "ImportVreme": BIM_DATUM_2, "Obradjeno": "", "Stornirano": "Da",
         "PocetnoStanje": 8000, "ZavrsnoStanje": BIM_IZVOD_2_ZAVRSNO,
         "UkupanDuguje": 3000, "UkupanPotrazuje": 950},
        # --- Izvod 3, racun 1: DVA reda pod ISTIM BankaImportID-em.
        #
        # Nista u kodu ne brani duplikat, a RequireSingleRow (koja na kraju
        # odlucuje) fail-close-uje na njega. Bez ovog para bi tvrdnja
        # "dvosmislen ID nosi PRAZAN identitet" merila odsustvo reda -- bila bi
        # zelena i kad bi radnja pogadjala prvi pogodak.
        #
        # Oba su Obradjeno="Da": ostaju vidljivi pod cipom "sve" (gde ih
        # identitetska tvrdnja i trazi), a ne ulaze u red za mapiranje, pa je
        # izvod 3 jedini BEZ otvorenih stavki i cip "sa otvorenim" ima sta da
        # iskljuci. Saldo se slaze: 2000 + 200 - 0 = 2200.
        {"BankaImportID": "BIM-FIX-DUP", "BrojDokumenta": BIM_IZVOD_3,
         "DatumIzvoda": BIM_DATUM_2, "BrojRacuna": BIM_RACUN_1,
         "DatumTransakcije": BIM_DATUM_2, "Partner": "Dvojnik Prvi",
         "PartnerKonto": "", "Opis": "Dvosmislen ID (1)", "Uplata": 100, "Isplata": 0,
         "Valuta": "RSD", "PozivNaBroj": "", "SvrhaPlacanja": "",
         "BankaReferenz": "REF-FIX-D1", "IzvorFajl": "fixture4.pdf",
         "ImportVreme": BIM_DATUM_2, "Obradjeno": "Da",
         "PocetnoStanje": 2000, "ZavrsnoStanje": 2200,
         "UkupanDuguje": 0, "UkupanPotrazuje": 200},
        {"BankaImportID": "BIM-FIX-DUP", "BrojDokumenta": BIM_IZVOD_3,
         "DatumIzvoda": BIM_DATUM_2, "BrojRacuna": BIM_RACUN_1,
         "DatumTransakcije": BIM_DATUM_2, "Partner": "Dvojnik Drugi",
         "PartnerKonto": "", "Opis": "Dvosmislen ID (2)", "Uplata": 100, "Isplata": 0,
         "Valuta": "RSD", "PozivNaBroj": "", "SvrhaPlacanja": "",
         "BankaReferenz": "REF-FIX-D2", "IzvorFajl": "fixture4.pdf",
         "ImportVreme": BIM_DATUM_2, "Obradjeno": "Da",
         "PocetnoStanje": 2000, "ZavrsnoStanje": 2200,
         "UkupanDuguje": 0, "UkupanPotrazuje": 200},
    ],
    # Tri prijemnice, sve tri sa razlogom:
    #   PRJ-TEST-A i PRJ-TEST-B  ISTI broj, RAZLICIT kupac -> kolizija identiteta
    #   PRJ-TEST-S               stornirana; nosi paletu koja time postaje osirocena
    # Sve tri imaju aktivnu zbirnu, pa NISU osirocene prijemnice -- lista
    # osirocenih ostaje prazna i meri bas ono sto treba (zbirna, ne kupac).
    "tblPrijemnica": [
        # Prijemnica ciji je RODITELJ zbirna sa dvosmislenim brojem. Kroz nju se
        # dohvata kaskadna kapija: PONISTENJE prijemnice zove PonistiZbirnaChain_TX
        # nad roditeljem, a taj put ne prolazi kroz kapiju na nivou moda zbirne.
        # TUDJA prijemnica na dvosmislenoj zbirni. Nijedan test je ne stornira
        # ni ne pomera, pa relink po BROJU stare zbirne ima sta da zahvati.
        # Izvorna prijemnica: bez kapije bi presla na ciljnu zbirnu.
        {"PrijemnicaID": "PRJ-OLD-U", "Datum": FIXTURE_DATE, "KupacID": KUPAC,
         "VozacID": VOZAC, "BrojPrijemnice": PRIJEMNICA_OLD_U, "BrojZbirne": ZBIRNA_OLDU,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 100, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 10, "Klasa": "I"},
        {"PrijemnicaID": "PRJ-STL-T", "Datum": FIXTURE_DATE, "KupacID": KUPAC2,
         "VozacID": VOZAC, "BrojPrijemnice": PRIJEMNICA_STALE, "BrojZbirne": ZBIRNA_KASK,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 100, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 10, "Klasa": "I"},
        {"PrijemnicaID": "PRJ-KASK-1", "Datum": FIXTURE_DATE, "KupacID": KUPAC,
         "VozacID": VOZAC, "BrojPrijemnice": "2/150326", "BrojZbirne": ZBIRNA_KASK,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 100, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 10, "Klasa": "I"},
        {"PrijemnicaID": "PRJ-TEST-I1", "Datum": FIXTURE_DATE, "KupacID": KUPAC,
         "VozacID": VOZAC, "BrojPrijemnice": PRIJEMNICA_ISPRAVKA, "BrojZbirne": ZBIRNA_MIRNA,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 120, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 12, "Klasa": "I"},
        {"PrijemnicaID": "PRJ-TEST-I2", "Datum": FIXTURE_DATE, "KupacID": KUPAC2,
         "VozacID": VOZAC, "BrojPrijemnice": PRIJEMNICA_ISPRAVKA, "BrojZbirne": ZBIRNA_MIRNA,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 180, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 18, "Klasa": "I"},
        # Deljena paleta: D1 i D2 nose isti broj i istu robu, svaki svom kupcu.
        {"PrijemnicaID": "PRJ-TEST-D1", "Datum": FIXTURE_DATE, "KupacID": KUPAC,
         "VozacID": VOZAC, "BrojPrijemnice": PRIJEMNICA_DELJENA, "BrojZbirne": ZBIRNA_MIRNA,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 100, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 10, "Klasa": "I", "Stornirano": "Da"},
        {"PrijemnicaID": "PRJ-TEST-D2", "Datum": FIXTURE_DATE, "KupacID": KUPAC2,
         "VozacID": VOZAC, "BrojPrijemnice": PRIJEMNICA_DELJENA, "BrojZbirne": ZBIRNA_MIRNA,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 150, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 15, "Klasa": "I", "Stornirano": "Da"},
        # EKRAN FAKTURISANJE. Tri reda istog kupca, svaki nosi po jedno stanje
        # koje citac mora da razlikuje -- bez njih lista prijemnica ima samo
        # slobodne redove, pa se kolona "dostupna" nikad ne razlikuje od prazne
        # kolone broja fakture i tvrdnja o njoj nema nad cim da padne.
        #
        # 1) VEC FAKTURISANA -- uredno obelezena: i oznaka i FakturaID.
        {"PrijemnicaID": "PRJ-FAK-1", "Datum": FIXTURE_DATE, "KupacID": KUPAC,
         "VozacID": VOZAC, "BrojPrijemnice": "20/150326", "BrojZbirne": ZBIRNA_MIRNA,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 100, "Cena": 40.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 10, "Klasa": "I",
         "Fakturisano": "Da", "FakturaID": FAKTURA_PLAC},
        # 2) NEPOTPUNO OBELEZENA: Fakturisano="Da", a FakturaID PRAZAN. Iz prikaza
        #    (kolona broja fakture je prazna) izgleda slobodna, a
        #    IsPrijemnicaAvailableForFaktura je odbija. Tacno taj raskorak deli
        #    "citam pravilo" od "citam ono sto se vidi".
        {"PrijemnicaID": "PRJ-FAK-2", "Datum": FIXTURE_DATE, "KupacID": KUPAC,
         "VozacID": VOZAC, "BrojPrijemnice": "21/150326", "BrojZbirne": ZBIRNA_MIRNA,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 100, "Cena": 40.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 10, "Klasa": "I",
         "Fakturisano": "Da"},
        # 3) SLOBODNA -- referentna tacka; jedina od tri sme u korpu.
        {"PrijemnicaID": "PRJ-FAK-3", "Datum": FIXTURE_DATE, "KupacID": KUPAC,
         "VozacID": VOZAC, "BrojPrijemnice": "22/150326", "BrojZbirne": ZBIRNA_MIRNA,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 100, "Cena": 40.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 10, "Klasa": "I"},
        # Cilj druge vrste -- jedinstven broj, bez kolizije.
        {"PrijemnicaID": "PRJ-TEST-T2", "Datum": FIXTURE_DATE, "KupacID": KUPAC,
         "VozacID": VOZAC, "BrojPrijemnice": PRIJEMNICA_CILJ_V2, "BrojZbirne": ZBIRNA_MIRNA,
         "VrstaVoca": VRSTA2, "SortaVoca": SORTA, "Kolicina": 100, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 10, "Klasa": "I"},
        # Kolizioni par AKTIVNIH: prevezivanje na zbirnu sme da dira SAMO Z1.
        {"PrijemnicaID": "PRJ-TEST-Z1", "Datum": FIXTURE_DATE, "KupacID": KUPAC,
         "VozacID": VOZAC, "BrojPrijemnice": PRIJEMNICA_ZBR_KOLIZIJA, "BrojZbirne": ZBIRNA_MIRNA,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 100, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 10, "Klasa": "I"},
        {"PrijemnicaID": "PRJ-TEST-Z2", "Datum": FIXTURE_DATE, "KupacID": KUPAC2,
         "VozacID": VOZAC, "BrojPrijemnice": PRIJEMNICA_ZBR_KOLIZIJA, "BrojZbirne": ZBIRNA_MIRNA,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 200, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 20, "Klasa": "I"},
        {"PrijemnicaID": "PRJ-TEST-A", "Datum": FIXTURE_DATE, "KupacID": KUPAC,
         "VozacID": VOZAC, "BrojPrijemnice": PRIJEMNICA_BROJ, "BrojZbirne": ZBIRNA2,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 300, "Cena": 60.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 30, "Klasa": "I"},
        {"PrijemnicaID": "PRJ-TEST-B", "Datum": FIXTURE_DATE, "KupacID": KUPAC2,
         "VozacID": VOZAC, "BrojPrijemnice": PRIJEMNICA_BROJ, "BrojZbirne": ZBIRNA2,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 700, "Cena": 80.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 70, "Klasa": "I"},
        {"PrijemnicaID": "PRJ-TEST-S", "Datum": FIXTURE_DATE, "KupacID": KUPAC,
         "VozacID": VOZAC, "BrojPrijemnice": PRIJEMNICA_STORNO, "BrojZbirne": ZBIRNA,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 400, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 40, "Klasa": "I", "Stornirano": "Da"},
        # KOLIZIONI PAR: dve STORNIRANE prijemnice ISTOG broja, dva kupca,
        # svaka sa svojom paletom. Bez njega se ne moze dokazati da prevezivanje
        # dira SAMO svoj dokument -- a to je bio otvoren nalaz: izvor se birao
        # po broju, pa bi ovaj par bio zahvacen zajedno.
        #
        # Na ZASEBNOM broju (8/150326), ne na 9/150326: testovi dele svesku, pa
        # test koji prevezuje 9/150326 ne sme da potrosi podatke onome koji
        # dokazuje izolaciju.
        {"PrijemnicaID": "PRJ-TEST-C1", "Datum": FIXTURE_DATE, "KupacID": KUPAC,
         "VozacID": VOZAC, "BrojPrijemnice": PRIJEMNICA_STORNO2, "BrojZbirne": ZBIRNA2,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 400, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 40, "Klasa": "I", "Stornirano": "Da"},
        {"PrijemnicaID": "PRJ-TEST-C2", "Datum": FIXTURE_DATE, "KupacID": KUPAC2,
         "VozacID": VOZAC, "BrojPrijemnice": PRIJEMNICA_STORNO2, "BrojZbirne": ZBIRNA2,
         "VrstaVoca": VRSTA2, "SortaVoca": SORTA, "Kolicina": 250, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 25, "Klasa": "I", "Stornirano": "Da"},
        # EKRAN SLEDLJIVOST -- v. blok konstanti SLED_* gore. PRJ-SLED-1 je
        # PRVA nestornirana prijemnica na zbirni koju nosi otpremnica sa
        # blokovima (zatvara rupu iz par. 23.12/S10): fakturisana, kg = zbir
        # oba bloka. PRJ-SLED-N je Fakturisano=Ne -- od kruga 9 LEGITIMNO
        # stanje (roba u hladnjaci), lanac potpun; kvar fakture mere F-vozila.
        {"PrijemnicaID": "PRJ-SLED-1", "Datum": FIXTURE_DATE, "KupacID": KUPAC,
         "VozacID": VOZAC2, "BrojPrijemnice": SLED_PRIJ_BROJ, "BrojZbirne": SLED_ZBIRNA,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": SLED_KG_1 + SLED_KG_2,
         "Cena": 50.0, "Klasa": "I", "Fakturisano": "Da", "FakturaID": SLED_FAKTURA},
        {"PrijemnicaID": "PRJ-SLED-N", "Datum": FIXTURE_DATE, "KupacID": KUPAC2,
         "VozacID": VOZAC2, "BrojPrijemnice": SLED_PRIJ_BROJ_N, "BrojZbirne": SLED_ZBIRNA_N,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 150,
         "Cena": 50.0, "Klasa": "I"},
        # Krug 8 R1: F-lanac -- F1 na AKTIVNU fakturu, F2 "Da" na
        # NEPOSTOJECU (ALL-pravilo obara celu kariku; NEPOTPUNI je vidi).
        {"PrijemnicaID": "PRJ-SLED-F1", "Datum": FIXTURE_DATE, "KupacID": KUPAC2,
         "VozacID": VOZAC2, "BrojPrijemnice": "33/150326", "BrojZbirne": SLED_ZBIRNA_F,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 60,
         "Cena": 50.0, "Klasa": "I", "Fakturisano": "Da", "FakturaID": SLED_FAKTURA_2},
        {"PrijemnicaID": "PRJ-SLED-F2", "Datum": FIXTURE_DATE, "KupacID": KUPAC2,
         "VozacID": VOZAC2, "BrojPrijemnice": "34/150326", "BrojZbirne": SLED_ZBIRNA_F,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 60,
         "Cena": 50.0, "Klasa": "I", "Fakturisano": "Da", "FakturaID": "FAK-NEMA-GA"},
        # Krug 8 R2: M-lanac -- dve prijemnice na DVE aktivne fakture
        # (potpun; "2 prij."/"2 fakt." prikaz, brojevi u SearchRefs).
        {"PrijemnicaID": "PRJ-SLED-M1", "Datum": FIXTURE_DATE, "KupacID": KUPAC2,
         "VozacID": VOZAC2, "BrojPrijemnice": "35/150326", "BrojZbirne": SLED_ZBIRNA_M,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 100,
         "Cena": 50.0, "Klasa": "I", "Fakturisano": "Da", "FakturaID": SLED_FAKTURA_3},
        {"PrijemnicaID": "PRJ-SLED-M2", "Datum": FIXTURE_DATE, "KupacID": KUPAC2,
         "VozacID": VOZAC2, "BrojPrijemnice": "36/150326", "BrojZbirne": SLED_ZBIRNA_M,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 100,
         "Cena": 50.0, "Klasa": "I", "Fakturisano": "Da", "FakturaID": SLED_FAKTURA_3B},
        # GP grana: sve tri NEfakturisane (hladnjaca tok, krug 9 -- to je
        # legitimno stanje; dalje karike su paleta/prerada/GP faktura).
        {"PrijemnicaID": "PRJ-SLED-G", "Datum": FIXTURE_DATE, "KupacID": KUPAC2,
         "VozacID": VOZAC2, "BrojPrijemnice": "37/150326", "BrojZbirne": SLED_ZBIRNA_G,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": SLED_KG_G,
         "Cena": 50.0, "Klasa": "I"},
        {"PrijemnicaID": "PRJ-SLED-H", "Datum": FIXTURE_DATE, "KupacID": KUPAC2,
         "VozacID": VOZAC2, "BrojPrijemnice": "38/150326", "BrojZbirne": SLED_ZBIRNA_H,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": SLED_KG_H,
         "Cena": 50.0, "Klasa": "I"},
        {"PrijemnicaID": "PRJ-SLED-K", "Datum": FIXTURE_DATE, "KupacID": KUPAC2,
         "VozacID": VOZAC2, "BrojPrijemnice": "39/150326", "BrojZbirne": SLED_ZBIRNA_K,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": SLED_KG_K,
         "Cena": 50.0, "Klasa": "I"},
        {"PrijemnicaID": "PRJ-SLED-P", "Datum": FIXTURE_DATE, "KupacID": KUPAC2,
         "VozacID": VOZAC2, "BrojPrijemnice": "40/150326", "BrojZbirne": SLED_ZBIRNA_P,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": SLED_KG_P,
         "Cena": 50.0, "Klasa": "I"},
    ],
    # Paleta i njena stavka vise o STORNIRANOJ prijemnici -> tacno ono sto
    # GetPrijemniceSaOsirocenimPaletama treba da nadje.
    "tblPaleta": [
        {"PaletaID": "PAL-TEST-D", "BrojPalete": 21, "Godina": 2026,
         "Datum": FIXTURE_DATE, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "Klasa": "I", "TipAmbalaze": AMB_12_1, "BrojGajbica": 25,
         "KapacitetGajbica": 100, "NetoKg": 250, "Status": "OTVORENA"},
        {"PaletaID": "PAL-TEST-Z1", "BrojPalete": 11, "Godina": 2026,
         "Datum": FIXTURE_DATE, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "Klasa": "I", "TipAmbalaze": AMB_12_1, "BrojGajbica": 10,
         "KapacitetGajbica": 100, "NetoKg": 100, "Status": "OTVORENA"},
        {"PaletaID": "PAL-TEST-Z2", "BrojPalete": 12, "Godina": 2026,
         "Datum": FIXTURE_DATE, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "Klasa": "I", "TipAmbalaze": AMB_12_1, "BrojGajbica": 20,
         "KapacitetGajbica": 100, "NetoKg": 200, "Status": "OTVORENA"},
        # ISTI BROJ, RANIJA GODINA. Broj palete se resetuje po godini
        # (GenerateBrojPalete racuna maxN+1 unutar Year(Date)), pa 12/2025 i
        # 12/2026 postoje istovremeno. Dok je ekran identitet resavao preko
        # broja, jedan od ta dva zapisa je bio NEDOSTUPAN: radnja nad starijom
        # paletom je pogadjala noviju. Bez ovog reda tvrdnja nema nad cim da
        # padne -- sve palete u fixture-u su bile iz iste godine.
        {"PaletaID": PALETA_KOLIZIJA_ID, "BrojPalete": PALETA_KOLIZIJA_BROJ,
         "Godina": 2025,
         "Datum": FIXTURE_DATE, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "Klasa": "I", "TipAmbalaze": AMB_12_1, "BrojGajbica": 15,
         "KapacitetGajbica": 100, "NetoKg": 150, "Status": "OTVORENA"},
        {"PaletaID": "PAL-TEST-1", "BrojPalete": 1, "Godina": 2026,
         "Datum": FIXTURE_DATE, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "Klasa": "I", "TipAmbalaze": AMB_12_1, "BrojGajbica": 40,
         "KapacitetGajbica": 100, "NetoKg": 400, "Status": "OTVORENA"},
        {"PaletaID": "PAL-TEST-2", "BrojPalete": 2, "Godina": 2026,
         "Datum": FIXTURE_DATE, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "Klasa": "I", "TipAmbalaze": AMB_12_1, "BrojGajbica": 40,
         "KapacitetGajbica": 100, "NetoKg": 400, "Status": "OTVORENA"},
        {"PaletaID": "PAL-TEST-3", "BrojPalete": 3, "Godina": 2026,
         "Datum": FIXTURE_DATE, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "Klasa": "I", "TipAmbalaze": AMB_12_1, "BrojGajbica": 25,
         "KapacitetGajbica": 100, "NetoKg": 250, "Status": "OTVORENA"},
        # EKRAN SLEDLJIVOST, smoke krug 3 (mete sledljivosti): roba potpunog
        # SLED lanca lezi i na ZATVORENOJ svezoj paleti (meta "u magacinu
        # sveze robe"). ZATVORENA namerno: otvorene palete iste vrste ulaze u
        # GajbeDoZatvaranjaPaleteInfo racun ljuske, zatvorene ne diraju nista.
        # Brojevi 31-33 su van svih postojecih (1,2,3,11,12,21) i ne pomeraju
        # nijedan kolizioni par.
        {"PaletaID": "PAL-SLED-1", "BrojPalete": 31, "Godina": 2026,
         "Datum": FIXTURE_DATE, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "Klasa": "I", "TipAmbalaze": AMB_12_1, "BrojGajbica": 25,
         "KapacitetGajbica": 100, "NetoKg": 250, "Status": "ZATVORENA"},
        # Roba SLN lanca (nefakturisana prijemnica) je PRERADJENA -- paleta
        # postoji ali NIJE meta "sveze robe"; njena sledljivost je prerada.
        {"PaletaID": "PAL-SLED-2", "BrojPalete": 32, "Godina": 2026,
         "Datum": FIXTURE_DATE, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "Klasa": "I", "TipAmbalaze": AMB_12_1, "BrojGajbica": 15,
         "KapacitetGajbica": 100, "NetoKg": 150, "Status": "ZATVORENA",
         "Preradjeno": "Da"},
        # STORNIRANA paleta na istoj SLED zbirnoj -- ne sme biti meta
        # (negativ za filter storna; stavka joj NIJE stornirana, filter mora
        # da padne na paleti).
        {"PaletaID": "PAL-SLED-X", "BrojPalete": 33, "Godina": 2026,
         "Datum": FIXTURE_DATE, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "Klasa": "I", "TipAmbalaze": AMB_12_1, "BrojGajbica": 10,
         "KapacitetGajbica": 100, "NetoKg": 99, "Status": "ZATVORENA",
         "Stornirano": "Da"},
        # Krug 8 R4: NEVALIDAN datum -- dokument MORA ostati vidljiv u
        # ponudi polja izbora (IIf mina bi na njemu pukla). ZATVORENA iz
        # istog razloga kao ostale SLED palete.
        {"PaletaID": "PAL-SLED-B", "BrojPalete": 34, "Godina": 2026,
         "Datum": "nevalidan", "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "Klasa": "I", "TipAmbalaze": AMB_12_1, "BrojGajbica": 8,
         "KapacitetGajbica": 100, "NetoKg": 77, "Status": "ZATVORENA"},
        # GP grana: G preradjena (roba prodata kao GP), H SVEZA (stanje
        # "u hladnjaci"), K preradjena (kontradiktorna prerada nad njom).
        # Brojevi 50/60/70 -- v. komentar bloka SLED_* konstanti.
        {"PaletaID": "PAL-SLED-G", "BrojPalete": 50, "Godina": 2026,
         "Datum": FIXTURE_DATE, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "Klasa": "I", "TipAmbalaze": AMB_12_1, "BrojGajbica": 8,
         "KapacitetGajbica": 100, "NetoKg": SLED_KG_G, "Status": "ZATVORENA",
         "Preradjeno": "Da"},
        {"PaletaID": "PAL-SLED-H", "BrojPalete": 60, "Godina": 2026,
         "Datum": FIXTURE_DATE, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "Klasa": "I", "TipAmbalaze": AMB_12_1, "BrojGajbica": 9,
         "KapacitetGajbica": 100, "NetoKg": SLED_KG_H, "Status": "ZATVORENA"},
        {"PaletaID": "PAL-SLED-K", "BrojPalete": 70, "Godina": 2026,
         "Datum": FIXTURE_DATE, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "Klasa": "I", "TipAmbalaze": AMB_12_1, "BrojGajbica": 7,
         "KapacitetGajbica": 100, "NetoKg": SLED_KG_K, "Status": "ZATVORENA",
         "Preradjeno": "Da"},
        {"PaletaID": "PAL-SLED-P", "BrojPalete": 80, "Godina": 2026,
         "Datum": FIXTURE_DATE, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "Klasa": "I", "TipAmbalaze": AMB_12_1, "BrojGajbica": 15,
         "KapacitetGajbica": 100, "NetoKg": SLED_KG_P, "Status": "ZATVORENA",
         "Preradjeno": "Da"},
    ],
    "tblPaletaStavka": [
        # ISTA fizicka paleta, dva dokumenta istog broja.
        {"StavkaID": "PST-TEST-D1", "PaletaID": "PAL-TEST-D",
         "BrojPrijemnice": PRIJEMNICA_DELJENA, "BrojZbirne": ZBIRNA_MIRNA,
         "BrojGajbica": 10, "NetoKg": 100, "PrijemnicaID": "PRJ-TEST-D1",
         "Klasa": "I", "VrstaVoca": VRSTA, "SortaVoca": SORTA},
        {"StavkaID": "PST-TEST-D2", "PaletaID": "PAL-TEST-D",
         "BrojPrijemnice": PRIJEMNICA_DELJENA, "BrojZbirne": ZBIRNA_MIRNA,
         "BrojGajbica": 15, "NetoKg": 150, "PrijemnicaID": "PRJ-TEST-D2",
         "Klasa": "I", "VrstaVoca": VRSTA, "SortaVoca": SORTA},
        {"StavkaID": "PST-TEST-Z1", "PaletaID": "PAL-TEST-Z1",
         "BrojPrijemnice": PRIJEMNICA_ZBR_KOLIZIJA, "BrojZbirne": ZBIRNA_MIRNA,
         "BrojGajbica": 10, "NetoKg": 100, "PrijemnicaID": "PRJ-TEST-Z1",
         "Klasa": "I", "VrstaVoca": VRSTA, "SortaVoca": SORTA},
        {"StavkaID": "PST-TEST-Z2", "PaletaID": "PAL-TEST-Z2",
         "BrojPrijemnice": PRIJEMNICA_ZBR_KOLIZIJA, "BrojZbirne": ZBIRNA_MIRNA,
         "BrojGajbica": 20, "NetoKg": 200, "PrijemnicaID": "PRJ-TEST-Z2",
         "Klasa": "I", "VrstaVoca": VRSTA, "SortaVoca": SORTA},
        {"StavkaID": "PST-TEST-1", "PaletaID": "PAL-TEST-1",
         "BrojPrijemnice": PRIJEMNICA_STORNO, "BrojZbirne": ZBIRNA,
         "BrojGajbica": 40, "NetoKg": 400, "PrijemnicaID": "PRJ-TEST-S",
         "Klasa": "I", "VrstaVoca": VRSTA, "SortaVoca": SORTA},
        # Kolizioni par: dve palete pod ISTIM brojem prijemnice, dva dokumenta.
        {"StavkaID": "PST-TEST-C1", "PaletaID": "PAL-TEST-2",
         "BrojPrijemnice": PRIJEMNICA_STORNO2, "BrojZbirne": ZBIRNA2,
         "BrojGajbica": 40, "NetoKg": 400, "PrijemnicaID": "PRJ-TEST-C1",
         "Klasa": "I", "VrstaVoca": VRSTA, "SortaVoca": SORTA},
        {"StavkaID": "PST-TEST-C2", "PaletaID": "PAL-TEST-3",
         "BrojPrijemnice": PRIJEMNICA_STORNO2, "BrojZbirne": ZBIRNA2,
         "BrojGajbica": 25, "NetoKg": 250, "PrijemnicaID": "PRJ-TEST-C2",
         "Klasa": "I", "VrstaVoca": VRSTA2, "SortaVoca": SORTA},
        # SLEDLJIVOST mete: stavke nose BrojZbirne -- podatkovna veza
        # zbirna -> paleta koju ReportSledljivostMete cita (bez premoscivanja).
        {"StavkaID": "PST-SLED-1", "PaletaID": "PAL-SLED-1",
         "BrojPrijemnice": SLED_PRIJ_BROJ, "BrojZbirne": SLED_ZBIRNA,
         "BrojGajbica": 25, "NetoKg": 250, "PrijemnicaID": "PRJ-SLED-1",
         "Klasa": "I", "VrstaVoca": VRSTA, "SortaVoca": SORTA},
        {"StavkaID": "PST-SLED-2", "PaletaID": "PAL-SLED-2",
         "BrojPrijemnice": SLED_PRIJ_BROJ_N, "BrojZbirne": SLED_ZBIRNA_N,
         "BrojGajbica": 15, "NetoKg": 150, "PrijemnicaID": "PRJ-SLED-N",
         "Klasa": "I", "VrstaVoca": VRSTA, "SortaVoca": SORTA},
        {"StavkaID": "PST-SLED-X", "PaletaID": "PAL-SLED-X",
         "BrojPrijemnice": SLED_PRIJ_BROJ, "BrojZbirne": SLED_ZBIRNA,
         "BrojGajbica": 10, "NetoKg": 99, "PrijemnicaID": "PRJ-SLED-1",
         "Klasa": "I", "VrstaVoca": VRSTA, "SortaVoca": SORTA},
        # GP grana: stavke vezuju BROJ zbirne (podatkovna veza lanca).
        {"StavkaID": "PST-SLED-G", "PaletaID": "PAL-SLED-G",
         "BrojPrijemnice": "37/150326", "BrojZbirne": SLED_ZBIRNA_G,
         "BrojGajbica": 8, "NetoKg": SLED_KG_G, "PrijemnicaID": "PRJ-SLED-G",
         "Klasa": "I", "VrstaVoca": VRSTA, "SortaVoca": SORTA},
        {"StavkaID": "PST-SLED-H", "PaletaID": "PAL-SLED-H",
         "BrojPrijemnice": "38/150326", "BrojZbirne": SLED_ZBIRNA_H,
         "BrojGajbica": 9, "NetoKg": SLED_KG_H, "PrijemnicaID": "PRJ-SLED-H",
         "Klasa": "I", "VrstaVoca": VRSTA, "SortaVoca": SORTA},
        {"StavkaID": "PST-SLED-K", "PaletaID": "PAL-SLED-K",
         "BrojPrijemnice": "39/150326", "BrojZbirne": SLED_ZBIRNA_K,
         "BrojGajbica": 7, "NetoKg": SLED_KG_K, "PrijemnicaID": "PRJ-SLED-K",
         "Klasa": "I", "VrstaVoca": VRSTA, "SortaVoca": SORTA},
        {"StavkaID": "PST-SLED-P", "PaletaID": "PAL-SLED-P",
         "BrojPrijemnice": "40/150326", "BrojZbirne": SLED_ZBIRNA_P,
         "BrojGajbica": 15, "NetoKg": SLED_KG_P, "PrijemnicaID": "PRJ-SLED-P",
         "Klasa": "I", "VrstaVoca": VRSTA, "SortaVoca": SORTA},
    ],
    # DVE ispravke na cekanju, i to NAD OTPREMNICOM -- namerno ne nad
    # prijemnicom: detekcija ispravke prijemnice pita operatera kroz MsgBox, a
    # MsgBox u headless runu visi. Ovako se safe-stop pravilo ("dve ili vise na
    # cekanju = ne biraj naslepo") proverava nad istom deljenom rutinom, bez
    # ijednog dijaloga.
    "tblArtikli": [
        {"ArtikalID": ARTIKAL, "Naziv": "Test Preparat", "Tip": "Zastita",
         "JedinicaMere": "l", "CenaPoJedinici": ARTIKAL_CENA,
         "DozaPoHa": ARTIKAL_DOZA, "Kultura": VRSTA,
         "Pakovanje": ARTIKAL_PAKOVANJE, "Aktivan": STATUS_AKTIVAN},
        {"ArtikalID": ARTIKAL_BEZ_PAK, "Naziv": "Test Bez Pakovanja",
         "Tip": "Zastita", "JedinicaMere": "kg", "CenaPoJedinici": 100,
         "DozaPoHa": 1, "Kultura": VRSTA, "Aktivan": STATUS_AKTIVAN},
        {"ArtikalID": ARTIKAL_BEZ_STANJA, "Naziv": "Test Bez Stanja",
         "Tip": "Zastita", "JedinicaMere": "l", "CenaPoJedinici": 200,
         "DozaPoHa": 1, "Kultura": VRSTA, "Pakovanje": 1,
         "Aktivan": STATUS_AKTIVAN},
        {"ArtikalID": ARTIKAL_ZALIHA, "Naziv": "Test Zaliha",
         "Tip": "Zastita", "JedinicaMere": "kg", "CenaPoJedinici": 100,
         "DozaPoHa": 1, "Kultura": VRSTA, "Pakovanje": 1,
         "Aktivan": STATUS_AKTIVAN},
    ],
    "tblMagacin": [
        {"MagacinID": "MAG-TEST-1", "Datum": FIXTURE_DATE, "ArtikalID": ARTIKAL,
         "Tip": "Ulaz", "Kolicina": 20, "BrojDokumenta": "AGRO-ULAZ-1",
         "CenaPoJedinici": ARTIKAL_CENA, "Vrednost": 10000,
         "DobavljacID": "DOB-TEST"},
        {"MagacinID": "MAG-TEST-2", "Datum": FIXTURE_DATE, "ArtikalID": ARTIKAL,
         "Tip": "Izlaz", "Kolicina": 5, "KooperantID": "KOOP-TEST-1",
         "ParcelaID": "PAR-TEST-1", "BrojDokumenta": "AGRO-IZLAZ-1",
         "CenaPoJedinici": ARTIKAL_CENA, "Vrednost": AGRO_DUG_KOOP1},
        # Istoimeni kooperant takodje ima dug -- inace se ne bi ni pojavio u
        # listi dugova, pa kolizije prikaza ne bi ni bilo.
        #
        # Dug mu ide preko REZERVISANOG virtuelnog artikla (pocetni dug), ne
        # preko prave robe: GetMagacinStanje ga izuzima, pa stanje ART-TEST-1
        # ostaje tacno 15 i preduslov testa kapije stanja se ne pomera. Da je
        # ovde stajao ART-TEST-1, dva testa bi se tiho vezala jedan za drugi.
        {"MagacinID": "MAG-TEST-3", "Datum": FIXTURE_DATE,
         "ArtikalID": "ART-POC-DUG", "Tip": "Izlaz", "Kolicina": 1,
         "KooperantID": "KOOP-TEST-IME", "BrojDokumenta": "AGRO-POC-2",
         "CenaPoJedinici": 1, "Vrednost": 1},
        # KOOP-TEST-2 ("Drugi Testni") -- JEDNOZNACNO ime, sa dugom. Kapija
        # dvosmislenosti mora da odbije samo istoimene, a ne sve redom; bez
        # ovog reda "Drugi Testni" uopste nije u listi dugova, pa je tvrdnja
        # "jednoznacan prikaz daje svoj identitet" merila odsustvo reda.
        #
        # Isti rezervisani artikal kao gore, iz istog razloga: stanje
        # ART-TEST-1 mora da ostane tacno 15.
        {"MagacinID": "MAG-TEST-4", "Datum": FIXTURE_DATE,
         "ArtikalID": "ART-POC-DUG", "Tip": "Izlaz", "Kolicina": 1,
         "KooperantID": "KOOP-TEST-2", "BrojDokumenta": "AGRO-POC-3",
         "CenaPoJedinici": 1, "Vrednost": 1},
        # Zaliha za traku korpe. Samo ULAZ -- ne ulazi ni u jedan dug.
        {"MagacinID": "MAG-TEST-5", "Datum": FIXTURE_DATE,
         "ArtikalID": ARTIKAL_ZALIHA, "Tip": "Ulaz", "Kolicina": 1000,
         "BrojDokumenta": "AGRO-ULAZ-Z", "CenaPoJedinici": 100,
         "Vrednost": 100000, "DobavljacID": "DOB-TEST"},
    ],
    # PRERADE. Do sada nijedan red -- lista prerada je bila prazna, pa se
    # identitet reda nije mogao ni izmeriti. Isti obrazac kao palete: broj
    # prerade se resetuje po godini (GenerateBrojPrerade), pa dva zapisa nose
    # isti broj i razlikuju se samo po PreradaID.
    "tblPrerada": [
        {"PreradaID": PRERADA_NOVA_ID, "BrojPrerade": PRERADA_KOLIZIJA_BROJ,
         "Godina": 2026, "Datum": FIXTURE_DATE, "NetoIzlazKg": 300,
         "BrojKutija": 30, "BrojKesa": 60, "TipGotovogProizvoda": "Rinfuz"},
        {"PreradaID": PRERADA_STARA_ID, "BrojPrerade": PRERADA_KOLIZIJA_BROJ,
         "Godina": 2025, "Datum": FIXTURE_DATE, "NetoIzlazKg": 200,
         "BrojKutija": 20, "BrojKesa": 40, "TipGotovogProizvoda": "Rinfuz"},
        # SLEDLJIVOST mete: prerada nad PAL-SLED-2 -- roba SLN lanca prodata/
        # uskladistena kao PRERADJENA. Broj 41 van kolizionog para (7).
        {"PreradaID": "PRE-SLED-1", "BrojPrerade": 41,
         "Godina": 2026, "Datum": FIXTURE_DATE, "NetoIzlazKg": 150,
         "BrojKutija": 15, "BrojKesa": 30, "TipGotovogProizvoda": "Rinfuz"},
        # GP grana (v. blok konstanti SLED_ZBIRNA_G/K gore): G je PRODATA
        # kao GP (Fakturisano=Da na AKTIVNU GP fakturu), K je KONTRADIKCIJA
        # (Da na nepostojecu). Fakturisano/FakturaID kolone dodaje
        # ENSURE_COLS -- donor ih nema.
        {"PreradaID": "PRE-SLED-G", "BrojPrerade": 51,
         "Godina": 2026, "Datum": FIXTURE_DATE, "NetoIzlazKg": SLED_GP_IZLAZ_KG,
         "BrojKutija": 8, "BrojKesa": 16, "TipGotovogProizvoda": "Rinfuz"},
        {"PreradaID": "PRE-SLED-K", "BrojPrerade": 61,
         "Godina": 2026, "Datum": FIXTURE_DATE, "NetoIzlazKg": 56,
         "BrojKutija": 7, "BrojKesa": 14, "TipGotovogProizvoda": "Rinfuz"},
        # Krug 5: P lanac -- DELIMICNA prodaja (50 od 120 kg utovareno
        # i validno fakturisano -> stanje "delimicno prodato").
        {"PreradaID": "PRE-SLED-P", "BrojPrerade": 181,
         "Godina": 2026, "Datum": FIXTURE_DATE, "NetoIzlazKg": 120,
         "BrojKutija": 12, "BrojKesa": 24, "TipGotovogProizvoda": "Rinfuz"},
        # Potrosna vozila writer testa CreateFakturaGP_TX (nisu ni na
        # jednom SLED lancu -- mutacija ne dira tvrdnje lanca).
        {"PreradaID": "PRE-GP-W1", "BrojPrerade": 71,
         # DECIMALAN izlaz (R5 revizije #248): Val("50,5") na srpskom
         # locale-u cita 50 -- read-model mora IsNumeric/CDbl putem,
         # inace korpa i writer upisu razlicite vrednosti.
         "Godina": 2026, "Datum": FIXTURE_DATE, "NetoIzlazKg": 50.5,
         "BrojKutija": 5, "BrojKesa": 10, "TipGotovogProizvoda": "Rinfuz"},
        {"PreradaID": "PRE-GP-X", "BrojPrerade": 81,
         "Godina": 2026, "Datum": FIXTURE_DATE, "NetoIzlazKg": 40,
         "BrojKutija": 4, "BrojKesa": 8, "TipGotovogProizvoda": "Rinfuz",
         "Stornirano": "Da"},
        {"PreradaID": "PRE-GP-W0", "BrojPrerade": 91,
         "Godina": 2026, "Datum": FIXTURE_DATE, "NetoIzlazKg": 0,
         "BrojKutija": 0, "BrojKesa": 0, "TipGotovogProizvoda": "Rinfuz"},
        # Contract vozila (krug 5 -- neusaglasenosti zive na UTOVARU,
        # v. tblUtovar dole; brojevi i dalje 0/1 pravilo):
        #   B2 (101): utovar Da na AKTIVNU 10/2026 BEZ FST stavke.
        #   WL (111): utovar sa zaostalim FakturaID bez markera.
        #   WM (151): utovar Da bez FakturaID.
        #   OV (191): utovareno 30 kg od proizvedenih 20 (prekomerno).
        #   SB: FST stavka bez utovara (siroce).
        #   WT (121): prazan TipGotovogProizvoda -> writer odbija.
        #   DUP (131/141): dupli PreradaID -> grid prazni identitet.
        {"PreradaID": "PRE-GP-B2", "BrojPrerade": 101,
         "Godina": 2026, "Datum": FIXTURE_DATE, "NetoIzlazKg": 20,
         "BrojKutija": 2, "BrojKesa": 4, "TipGotovogProizvoda": "Rinfuz"},
        {"PreradaID": "PRE-GP-WL", "BrojPrerade": 111,
         "Godina": 2026, "Datum": FIXTURE_DATE, "NetoIzlazKg": 25,
         "BrojKutija": 2, "BrojKesa": 5, "TipGotovogProizvoda": "Rinfuz"},
        {"PreradaID": "PRE-GP-OV", "BrojPrerade": 191,
         "Godina": 2026, "Datum": FIXTURE_DATE, "NetoIzlazKg": 20,
         "BrojKutija": 2, "BrojKesa": 4, "TipGotovogProizvoda": "Rinfuz"},
        {"PreradaID": "PRE-GP-WT", "BrojPrerade": 121,
         "Godina": 2026, "Datum": FIXTURE_DATE, "NetoIzlazKg": 30,
         "BrojKutija": 3, "BrojKesa": 6, "TipGotovogProizvoda": ""},
        {"PreradaID": "PRE-GP-WM", "BrojPrerade": 151,
         "Godina": 2026, "Datum": FIXTURE_DATE, "NetoIzlazKg": 15,
         "BrojKutija": 1, "BrojKesa": 3, "TipGotovogProizvoda": "Rinfuz"},
        {"PreradaID": "PRE-GP-DUP", "BrojPrerade": 131,
         "Godina": 2026, "Datum": FIXTURE_DATE, "NetoIzlazKg": 10,
         "BrojKutija": 1, "BrojKesa": 2, "TipGotovogProizvoda": "Rinfuz"},
        {"PreradaID": "PRE-GP-DUP", "BrojPrerade": 141,
         "Godina": 2026, "Datum": FIXTURE_DATE, "NetoIzlazKg": 10,
         "BrojKutija": 1, "BrojKesa": 2, "TipGotovogProizvoda": "Rinfuz"},
        {"PreradaID": "PRE-GP-SB", "BrojPrerade": 171,
         "Godina": 2026, "Datum": FIXTURE_DATE, "NetoIzlazKg": 5,
         "BrojKutija": 1, "BrojKesa": 1, "TipGotovogProizvoda": "Rinfuz"},
    ],
    # Prve stavke prerade u fixture-u: kanonski join je PaletaID (kao
    # modIntegritet D2), BrojPalete je samo labela.
    "tblPreradaStavka": [
        {"StavkaID": "PRS-SLED-1", "PreradaID": "PRE-SLED-1",
         "PaletaID": "PAL-SLED-2", "BrojPalete": 32, "NetoKg": 150},
        # GP grana: join prerada -> paleta (kanonski PaletaID).
        {"StavkaID": "PRS-SLED-G", "PreradaID": "PRE-SLED-G",
         "PaletaID": "PAL-SLED-G", "BrojPalete": 50, "NetoKg": SLED_KG_G},
        {"StavkaID": "PRS-SLED-K", "PreradaID": "PRE-SLED-K",
         "PaletaID": "PAL-SLED-K", "BrojPalete": 70, "NetoKg": SLED_KG_K},
        {"StavkaID": "PRS-SLED-P", "PreradaID": "PRE-SLED-P",
         "PaletaID": "PAL-SLED-P", "BrojPalete": 80, "NetoKg": SLED_KG_P},
    ],
    # SEF DTO nad GP fakturom (R2 revizije #248): mapper trazi Naziv i
    # PIB kupca iz tblKupci -- do sada je kupac ziveo samo kao ID na
    # fakturi. Sejanjem se donorski kupci BRISU (deterministicki fixture,
    # kapije porede ID-eve pa im redovi ne trebaju).
    "tblKupci": [
        {"KupacID": KUPAC, "Naziv": "Test kupac 1", "PIB": "100000001"},
        {"KupacID": KUPAC2, "Naziv": "Test kupac 2", "PIB": "100000002"},
    ],
    "tblStornoVeze": [
        {"CorrectionID": "SV-TEST-1", "Mode": "ISPRAVKA_ODMAH", "Status": "PENDING",
         "OldDocType": "Otpremnica", "OldDocID": "OTP-TEST-2", "OldBroj": "2/TEST",
         "NeedsRecovery": "Da", "Message": "fixture: ispravka na cekanju"},
        {"CorrectionID": "SV-TEST-2", "Mode": "ISPRAVKA_ODMAH", "Status": "PENDING",
         "OldDocType": "Otpremnica", "OldDocID": "OTP-TEST-3", "OldBroj": "3/TEST",
         "NeedsRecovery": "Da", "Message": "fixture: druga ispravka na cekanju"},
    ],
}

# Kolone koje DONOR (produkcijska sveska pre nadogradnje) nema, a fixture
# mora da ih ima: sejanje ide PO IMENU, pa red sa novom kolonom obara
# generator; testovi writera (RequireUpdateCell) takodje traze kolonu.
# U aplikaciji ih dodaje modSetup.EnsurePaletniListSchema (EnsureColumnOnTable
# -> na KRAJ tabele); generator radi ISTO, pa je fixture = sveska POSLE
# nadogradnje. Kolona koja vec postoji se ne dira.
ENSURE_COLS = {
    "tblFakturaStavke": ["PreradaID", "BrojPrerade", "UtovarID"],
    # Revizija #9: rok trajanja po vrsti GP (prazno = globalni).
    "tblVrstaGotovihProizvoda": ["RokMeseci"],
}

# Tabele koje donor NEMA (krug 5: utovarna lista) -- generator ih pravi
# isto kao modSetup.EnsureUtovarSchemaCore (EnsureDataTable): novi sheet
# + ListObject sa ovim kolonama. Redosled = redosled u modSetup Array.
ENSURE_TABLES = {
    "tblUtovar": ("Utovar",
                  ["UtovarID", "BrojUtovara", "Godina", "DatumUtovara",
                   "KupacID", "Fakturisano", "FakturaID", "Napomena",
                   "Stornirano",
                   # krug 5d: profesionalna utovarna lista (prevoz).
                   "Prevoznik", "Vozac", "Registracija", "Plomba",
                   "TemperaturniRezim", "MestoIstovara", "VremeUtovara",
                   "BrojNarudzbenice"]),
    "tblUtovarStavke": ("UtovarStavke",
                        ["UtovarStavkaID", "UtovarID", "PreradaID",
                         "BrojPrerade", "KolicinaKg", "Stornirano",
                         # revizija #6 t.4: stvarno utovarena pakovanja;
                         # revizija #9: dogovorena cena (model B).
                         "BrojKutija", "BrojKesa", "CenaKg"]),
    # Smoke 5d: sifarnik EKSTERNIH prevoznika/vozaca (nije tblVozaci).
    "tblPrevoznici": ("Prevoznici",
                      ["PrevoznikID", "Naziv", "Vozac", "Registracija",
                       "Aktivan"]),
}

# tblLocalConfig (Kljuc | Vrednost | Opis)
LOCAL_CONFIG = {
    "APP_SETUP_COMPLETED": "DA",
}

# tblSEFConfig -- licenca off. LICENSE_ENABLED=NO nije dovoljno: modLicense ima
# LATCH (vidi modLicense.bas:21) -- gate radi i bez YES ako postoje LICENSE_KEY i
# LICENSE_BOUND_PARTS. Zato se ti kljucevi prazne.
SEF_CONFIG = {
    "LICENSE_ENABLED": "NO",
    "LICENSE_KEY": "",
    "LICENSE_TOKEN": "",
    "LICENSE_BOUND_PARTS": "",
    "LICENSE_NEXT_CHECK": "",
    "LICENSE_STATUS": "",
    "LICENSE_HWM": "",
    # TEST-KRITICAN CONFIG SE PINUJE, ne nasledjuje se od donora.
    #
    # tblSEFConfig je u KEEP_ROWS (brise se sve osim kataloga), pa je do sada
    # svaki kljuc koji fixture NE postavi ostajao onakav kakav je bio u donoru.
    # Ista suite je zato davala razlicit rezultat na dve sveske, i to je vise
    # PR-ova nosio kao 'dva crvena ali nisu moja':
    #   DEFAULT_SORTA_VOCA -> ApplyDefaultProizvod popuni combo, golden ocekuje prazan
    #   KES_ISPLATE        -> IsKesIsplate gasi granu 'isplata iz OM avansa'
    # Prazan string je i dalje 'nije postavljeno' za ApplyDefaultProizvod, ali je
    # sada ZAPISANO prazno, pa donorska vrednost ne moze da procuri.
    "DEFAULT_VRSTA_VOCA": "",
    "DEFAULT_SORTA_VOCA": "",
    "KES_ISPLATE": "YES",
    # Ekran Platni nalozi (v6-ui-185): pet kljuceva od kojih ekran zavisi se
    # PINUJE, ne nasledjuje od donora -- ista klasa kao DEFAULT_SORTA_VOCA
    # iznad. RACUN_1 je jedini racun firme (BankaNalogRacuniCSV ga vraca bez
    # fallback-a); SIFRA/SVRHA prazno = DocConfigOr pada na default konstante,
    # pa CSV payload ne zavisi od donora; ISPLATA_SPEC_PRINT_MODE=OFF da klik
    # na specifikaciju u testu/smoke-u nad fixture-om ne pravi PDF-ove.
    "BANKA_NALOG_RACUN_1": "160-1111111111-11",
    "BANKA_NALOG_RACUN_2": "",
    "BANKA_NALOG_RACUN_3": "",
    "BANKA_NALOG_RACUN_4": "",
    "BANKA_NALOG_RACUNI": "",
    "BANKA_NALOG_SIFRA_PLACANJA": "",
    "BANKA_NALOG_SVRHA": "",
    "ISPLATA_SPEC_PRINT_MODE": "OFF",
    # Ekran Izvestaji (v6-ui-186): ReportOtkupRobaOM u malina modu racuna
    # prijem DIREKTNO (1 otpremnica = 1 prijemnica po klasi) umesto srazmerno,
    # pa bi donor sa ukljucenim modom davao druge brojke manjka -- ista klasa
    # kao KES_ISPLATE. Pinuje se OFF; malina grana ima svoj E2E u
    # modIzvestajTests koji rezim postavlja sam.
    "MALINA_MODE": "NO",
    # Stampe iz izvestaja u testu/smoke-u nad fixture-om ne prave PDF-ove --
    # isti razlog kao ISPLATA_SPEC_PRINT_MODE iznad.
    "KARTICA_PRINT_MODE": "OFF",
    "KARTICA_AMB_PRINT_MODE": "OFF",
    # GP faktura (R1/R2 revizije #248): print test puni sablon bez
    # izlaza (OFF), a SEF DTO test trazi seller podatke -- pinovano da
    # ne zavisi od donora (ista klasa kao DEFAULT_SORTA_VOCA).
    "FAKTURA_PRINT_MODE": "OFF",
    "UTOVAR_PRINT_MODE": "OFF",
    "GP_ROK_TRAJANJA_MESECI": "24",
    "SELLER_NAME": "Test prodavac DOO",
    "SELLER_PIB": "100000000",
    # Ekran Sledljivost (v6-ui-187): "Lanac (PDF)" postuje ovaj rezim; OFF da
    # klik u testu/smoke-u nad fixture-om ne pravi PDF (ekran OFF prijavljuje
    # porukom, pa dugme ne izgleda mrtvo).
    "SLEDLJIVOST_PRINT_MODE": "OFF",
}

# DEFAULT_VRSTA_VOCA / DEFAULT_SORTA_VOCA se PINUJU NA PRAZNO (v. SEF_CONFIG):
# ApplyDefaultProizvod tada ostavlja combo-e prazne (frmOtkup ga zove pod
# On Error Resume Next), pa Initialize ne okida auto-cenu i stanje forme je
# deterministicno za golden snapshot.
#
# Do sada su bili samo IZOSTAVLJENI, a to nije isto: tblSEFConfig je u
# KEEP_ROWS, pa je donorova vrednost prezivljavala i golden je padao na
# svesci gde je sorta bila podesena.

EXCEL_EPOCH = datetime.date(1899, 12, 30)


# --- potpis: koji podaci su u fixture-u --------------------------------------
#
# Fixture je gitignored (.gitignore: tests/fixtures/), pa `git checkout` NE menja
# fajl na disku. Prelazak na granu koja seje nove redove ostavlja fixture
# prethodne grane, testovi padnu na podacima, a pad izgleda kao regresija koda.
# To je vec pojelo pola sata trijaze.
#
# Zato generator pored sveske ostavlja `otkup_test.sig` sa hash-om PODATAKA koje
# je posejao. run_vba.py ga poredi sa tekucim generatorom i staje pre Excela.
#
# Hash pokriva SAMO deklarativne podatke (SEED, config, datum, KEEP_ROWS) -- to
# je ono sto se menja od grane do grane. Izmena LOGIKE upisa (add_row, strip_rows)
# se ne vidi; nju operater regenerise namerno. Bolje uzak i tacan potpis nego
# sirok koji trazi regeneraciju na svaku izmenu komentara.
FIXTURE_SIG_EXT = ".sig"

# Rucna poluga za ono sto hash ne vidi. Kad se promeni SEMANTIKA generatora
# (add_row, strip_rows, upsert_config, ili sadrzaj tabela koje se cuvaju iz
# donora preko KEEP_ROWS), podaci u svesci se promene a deklarativni blokovi
# ostanu isti -- potpis bi tvrdio da je stari fixture i dalje dobar. Tada se
# ovaj broj podigne za jedan. Jeftinije i tacnije nego hashirati ceo .py, koji
# bi trazio regeneraciju i na izmenu komentara.
FIXTURE_FORMAT_VERSION = 1


def signature() -> str:
    payload = "\n".join([
        "FORMAT=" + str(FIXTURE_FORMAT_VERSION),
        "FIXTURE_DATE=" + FIXTURE_DATE.isoformat(),
        "KEEP_ROWS=" + repr(sorted(KEEP_ROWS)),
        "LOCAL_CONFIG=" + repr(sorted(LOCAL_CONFIG.items())),
        "SEF_CONFIG=" + repr(sorted(SEF_CONFIG.items())),
        "SEED=" + repr([(t, [sorted(r.items()) for r in rows])
                        for t, rows in sorted(SEED.items())]),
        "ENSURE_COLS=" + repr(sorted((t, cols) for t, cols in ENSURE_COLS.items())),
        "ENSURE_TABLES=" + repr(sorted((t, sh, cols) for t, (sh, cols) in ENSURE_TABLES.items())),
    ])
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()[:16]


def sig_path(workbook: str) -> str:
    return os.path.splitext(workbook)[0] + FIXTURE_SIG_EXT


def read_sig(workbook: str) -> str:
    """Potpis zapisan uz svesku; "" ako ga nema (fixture od starijeg generatora)."""
    try:
        with open(sig_path(workbook), "r", encoding="ascii") as fh:
            return fh.read().strip()
    except OSError:
        return ""


class SchemaError(Exception):
    pass


def xl_serial(d: datetime.date) -> int:
    """Excel serijski broj datuma -- bez vremenske zone, za razliku od datetime preko COM-a."""
    return (d - EXCEL_EPOCH).days


def iter_tables(wb):
    for ws in wb.Worksheets:
        for lo in ws.ListObjects:
            yield lo


def find_table(wb, name: str):
    target = name.strip().lower()
    for lo in iter_tables(wb):
        if str(lo.Name).strip().lower() == target:
            return lo
    return None


def header_index(lo) -> dict:
    return {str(c.Name).strip().lower(): int(c.Index) for c in lo.ListColumns}


def strip_vba(wb) -> list:
    """Izbaci SAV standardni/klasni/form kod iz donora.

    Kod u fixture-u je balast: run_vba.py na svakom pokretanju uveze svez
    src-vba/ preko njega. Ali uvozi samo ono sto repo IMA -- modul zaostao iz
    starijeg donora ostaje i izvrsava se. Ako nosi Public ime koje postoji i u
    svezem kodu, VBA to vidi kao "Ambiguous name" i odbija da pokrene makro iz
    njega, uz poruku "Cannot run the macro" koja ne lici na compile gresku.
    Tako je TestLicense_All bio mrtav dok je vba_check bio uredno zelen --
    duplikat nije bio u repou nego u svesci.

    Document moduli (listovi, ThisWorkbook) se NE mogu ukloniti; njihov kod
    run_vba merge-uje iz .doccls fajlova.

    Trazi "Trust access to the VBA project object model".
    """
    STD, CLS, FRM = 1, 2, 3
    removed = []
    try:
        proj = wb.VBProject
        comps = [c for c in proj.VBComponents]      # snapshot: brisemo iz kolekcije
    except Exception as exc:
        raise SchemaError(
            f"nema pristupa VBA projektu ({exc}). Ukljuci: File > Options > "
            "Trust Center > Trust Center Settings > Macro Settings > "
            "'Trust access to the VBA project object model'")

    for comp in comps:
        try:
            if int(comp.Type) not in (STD, CLS, FRM):
                continue
            name = str(comp.Name)
            proj.VBComponents.Remove(comp)
            removed.append(name)
        except Exception:
            pass                                    # zakljucan projekat/komponenta
    return sorted(removed)


def strip_rows(wb) -> list:
    cleared = []
    for lo in iter_tables(wb):
        name = str(lo.Name)
        if name.strip().lower() in KEEP_ROWS:
            continue
        try:
            if int(lo.ListRows.Count) > 0:
                n = int(lo.ListRows.Count)
                lo.DataBodyRange.Delete()
                cleared.append((name, n))
        except Exception as exc:
            raise SchemaError(f"{name}: brisanje redova nije uspelo ({exc})")
    return cleared


def add_row(lo, values: dict, table_name: str) -> None:
    idx = header_index(lo)
    missing = [k for k in values if k.strip().lower() not in idx]
    if missing:
        raise SchemaError(
            f"{table_name}: donor nema kolone {missing}. "
            f"Postojece: {sorted(idx)}"
        )
    row = lo.ListRows.Add()
    for key, val in values.items():
        cell = row.Range.Cells(1, idx[key.strip().lower()])
        if isinstance(val, datetime.date):
            cell.NumberFormat = "dd.mm.yyyy"
            cell.Value = xl_serial(val)
        elif isinstance(val, Sirovo):
            cell.NumberFormat = "General"
            cell.Value = val.v
        else:
            cell.Value = val


def upsert_config(wb, table_name: str, pairs: dict,
                  key_col: str = "Kljuc", val_col: str = "Vrednost") -> int:
    """Kljuc/vrednost tabele: postojeci kljuc se azurira, novi se dodaje.

    Imena kolona nisu ista svuda -- tblConfig/tblLocalConfig imaju Kljuc|Vrednost,
    a tblSEFConfig ConfigKey|ConfigValue.
    """
    lo = find_table(wb, table_name)
    if lo is None:
        raise SchemaError(f"{table_name} ne postoji u donoru")
    idx = header_index(lo)
    for needed in (key_col, val_col):
        if needed.strip().lower() not in idx:
            raise SchemaError(f"{table_name}: nema kolonu '{needed}' ({sorted(idx)})")
    kcol, vcol = idx[key_col.strip().lower()], idx[val_col.strip().lower()]
    akt = idx.get("aktivan")          # tblSEFConfig ima Aktivan; nov red mora biti aktivan

    existing = {}
    if int(lo.ListRows.Count) > 0:
        for r in range(1, int(lo.ListRows.Count) + 1):
            key = lo.ListRows(r).Range.Cells(1, kcol).Value
            if key is not None:
                existing[str(key).strip().upper()] = r

    for key, val in pairs.items():
        r = existing.get(key.strip().upper())
        # Vrednost ide kao TEKST: nov red nasledjuje format reda iznad,
        # pa datumski format pretvori "24" u datum 1900-01-23 --
        # GetConfigValue onda vrati smece (krug 5d: rok 2311900 meseci).
        if r is None:
            row = lo.ListRows.Add()
            row.Range.Cells(1, kcol).Value = key
            row.Range.Cells(1, vcol).NumberFormat = "@"
            row.Range.Cells(1, vcol).Value = val
            if akt:
                row.Range.Cells(1, akt).Value = STATUS_AKTIVAN
        else:
            cell = lo.ListRows(r).Range.Cells(1, vcol)
            cell.NumberFormat = "@"
            cell.Value = val
    return len(pairs)


def build(donor: str, out: str, force: bool) -> int:
    try:
        import win32com.client as win32
    except ImportError:
        print("pywin32 nije instaliran: python -m pip install pywin32", file=sys.stderr)
        return 2

    donor = os.path.abspath(donor)
    out = os.path.abspath(out)

    if not os.path.exists(donor):
        print(f"Donor ne postoji: {donor}", file=sys.stderr)
        return 2
    if os.path.normcase(donor) == os.path.normcase(out):
        print("Donor i izlaz su ista putanja -- odbijam (donor se ne dira).", file=sys.stderr)
        return 2
    if os.path.exists(out) and not force:
        print(f"Izlaz vec postoji: {out}\nDodaj --force da ga prepisem.", file=sys.stderr)
        return 2

    os.makedirs(os.path.dirname(out), exist_ok=True)

    # Potpis pada PRE nego sto se sveska dirne. Bez ovoga neuspeo build (donor bez
    # kolone -> SchemaError) ostavlja prepisanu svesku uz stari .sig, pa run_vba
    # cita potpis koji vise ne opisuje nista. Bolje "nema potpisa" nego lazan.
    try:
        os.remove(sig_path(out))
    except OSError:
        pass

    shutil.copy2(donor, out)          # radi se nad kopijom; donor ostaje netaknut

    xl = win32.DispatchEx("Excel.Application")
    wb = None
    try:
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.AutomationSecurity = MSO_AUTOMATION_SECURITY_LOW
        xl.EnableEvents = False       # KLJUCNO: Workbook_Open (StartApp) se ne pokrece

        wb = xl.Workbooks.Open(out, UpdateLinks=0)

        stripped = strip_vba(wb)
        if stripped:
            print(f"Uklonjeno {len(stripped)} VBA modula iz donora: "
                  + ", ".join(stripped[:8])
                  + (f" ... (+{len(stripped) - 8})" if len(stripped) > 8 else ""))

        cleared = strip_rows(wb)
        print(f"Obrisani redovi u {len(cleared)} tabela"
              + (": " + ", ".join(f"{n}({c})" for n, c in cleared) if cleared else ""))

        # Nove TABELE koje donor nema (krug 5) -- isto sto radi
        # modSetup.EnsureDataTable: sheet + ListObject sa kolonama.
        created_tables = []
        for table_name, (sheet_name, headers) in ENSURE_TABLES.items():
            lo_ex = find_table(wb, table_name)
            if lo_ex is None:
                ws_new = wb.Worksheets.Add()
                ws_new.Name = sheet_name
                for ci, h in enumerate(headers, start=1):
                    ws_new.Cells(1, ci).Value = h
                lo_new = ws_new.ListObjects.Add(
                    1, ws_new.Range(ws_new.Cells(1, 1),
                                    ws_new.Cells(1, len(headers))), None, 1)
                lo_new.Name = table_name
                created_tables.append(table_name)
            else:
                # Postojecoj tabeli (donor = prosli fixture) dopuni
                # kolone koje fale -- isto sto radi EnsureDataTable.
                idx_ex = header_index(lo_ex)
                for h in headers:
                    if h.strip().lower() not in idx_ex:
                        lo_ex.ListColumns.Add().Name = h
                        created_tables.append(f"{table_name}.{h}")
        if created_tables:
            print("Kreirano (tabele/kolone): " + ", ".join(created_tables))

        # Nadogradnja seme PRE sejanja (v. ENSURE_COLS): nove kolone na KRAJ,
        # isto sto radi modSetup.EnsureColumnOnTable na startu aplikacije.
        added_cols = []
        for table_name, cols in ENSURE_COLS.items():
            lo = find_table(wb, table_name)
            if lo is None:
                raise SchemaError(f"{table_name} ne postoji u donoru (ENSURE_COLS)")
            idx = header_index(lo)
            for col in cols:
                if col.strip().lower() not in idx:
                    lo.ListColumns.Add().Name = col
                    added_cols.append(f"{table_name}.{col}")
        if added_cols:
            print("Dodate kolone (nadogradnja seme): " + ", ".join(added_cols))

        seeded = []
        for table_name, rows in SEED.items():
            lo = find_table(wb, table_name)
            if lo is None:
                raise SchemaError(f"{table_name} ne postoji u donoru")
            for values in rows:
                add_row(lo, values, table_name)
            seeded.append((table_name, len(rows)))
        print("Posejano: " + ", ".join(f"{n}({c})" for n, c in seeded))

        upsert_config(wb, "tblLocalConfig", LOCAL_CONFIG)
        upsert_config(wb, "tblSEFConfig", SEF_CONFIG,
                      key_col="ConfigKey", val_col="ConfigValue")
        print(f"Config: tblLocalConfig({len(LOCAL_CONFIG)}), tblSEFConfig({len(SEF_CONFIG)}) -- licenca OFF")

        wb.Save()

        # Potpis se pise TEK posle uspesnog Save-a: sig uz svesku koja nije do
        # kraja napravljena tvrdio bi da je fixture svez.
        sig = signature()
        with open(sig_path(out), "w", encoding="ascii") as fh:
            fh.write(sig + "\n")

        print(f"\nFixture: {out}")
        print(f"Potpis:  {sig}  ({os.path.basename(sig_path(out))})")
        return 0
    except SchemaError as exc:
        print(f"\nSEMA: {exc}", file=sys.stderr)
        return 2
    except Exception as exc:
        print(f"\nGRESKA: {exc}", file=sys.stderr)
        return 2
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        xl.Quit()


def main(argv) -> int:
    ap = argparse.ArgumentParser(description="Pravi tests/fixtures/otkup_test.xlsm iz donor sveske.")
    ap.add_argument("--donor", required=True, help="putanja do .xlsm koja daje semu (ne menja se)")
    ap.add_argument("--out", default=DEFAULT_OUT, help=f"izlaz (podrazumevano {DEFAULT_OUT})")
    ap.add_argument("--force", action="store_true", help="prepisi postojeci izlaz")
    args = ap.parse_args(argv)
    return build(args.donor, args.out, args.force)


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
