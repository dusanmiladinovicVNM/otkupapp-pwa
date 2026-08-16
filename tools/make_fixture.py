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
# Dva AKTIVNA reda tblNovac pod ISTIM brojem -- avans raspodela to radi
# svakodnevno. Bez NovacID-a je broj dvosmislen i storno se odbija; sa njim se
# stornira bas izabran red. Preflight je do sada odbijao i kad ID postoji.
NOVAC_DUPLI_BROJ = "NOV-DUPLI-1"

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
# Dve AKTIVNE prijemnice istog broja za ISPRAVKU. Zaseban broj: test 35 pravi
# RESI KASNIJE context nad 6/150326, a pending ispravka nad istim brojem bi
# zaustavila ISPRAVKU (safe-stop) i test bi merio pogresnu stvar.
PRIJEMNICA_ISPRAVKA = "3/150326"

# Sejanje ide PO IMENU KOLONE -- ako donor nema neku od ovih kolona, skripta
# pukne glasno umesto da tiho napravi fixture nad kojim testovi lazu.
SEED = {
    "tblStanice": [
        {"StanicaID": STANICA, "Naziv": "Test Otkupno Mesto", "Mesto": "Test Mesto",
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
        {"KooperantID": "KOOP-TEST-1", "Ime": "Prvi", "Prezime": "Testni",
         "Mesto": "Test Mesto", "StanicaID": STANICA, "Aktivan": STATUS_AKTIVAN},
        {"KooperantID": "KOOP-TEST-2", "Ime": "Drugi", "Prezime": "Testni",
         "Mesto": "Test Mesto", "StanicaID": STANICA, "Aktivan": STATUS_AKTIVAN},
        {"KooperantID": "KOOP-TEST-3", "Ime": "Treci", "Prezime": "Testni",
         "Mesto": "Test Mesto", "StanicaID": STANICA, "Aktivan": STATUS_AKTIVAN},
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
        # ISTI BrojZbirne, ISTI kupac, DVA vozaca -> u jezgru dva dokumenta.
        # Ciljna lista Oporavka ih je spajala u jedan red jer je vlasnikom
        # smatrala samo kupca, pa operater nije mogao da izabere pravi.
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
    ],
    # Tri slucaja koje zadatak trazi:
    #   OTP-TEST-1  datum iz proslosti + poznata zbirna + ostatak != 0 (1000 - 400)
    #   OTP-TEST-2  bez zbirne
    #   OTP-TEST-3  bez zbirne, ali blok u tblOtkup nosi zbirnu (ZB-TEST-3)
    "tblOtpremnica": [
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
    ],
    "tblOtkup": [
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
        {"NovacID": "NOV-TEST-D1", "BrojDokumenta": NOVAC_DUPLI_BROJ,
         "Datum": FIXTURE_DATE, "Tip": "VirmanAvansKoop", "Isplata": 1000,
         "KooperantID": "KOOP-TEST-1"},
        {"NovacID": "NOV-TEST-D2", "BrojDokumenta": NOVAC_DUPLI_BROJ,
         "Datum": FIXTURE_DATE, "Tip": "VirmanAvansKoop", "Isplata": 2000,
         "KooperantID": "KOOP-TEST-2"},
    ],
    "tblFakture": [
        {"FakturaID": FAKTURA, "KupacID": KUPAC, "Iznos": FAKTURA_IZNOS},
        {"FakturaID": FAKTURA_BEZ_IZNOSA, "KupacID": KUPAC, "Iznos": 0},
    ],
    # Tri prijemnice, sve tri sa razlogom:
    #   PRJ-TEST-A i PRJ-TEST-B  ISTI broj, RAZLICIT kupac -> kolizija identiteta
    #   PRJ-TEST-S               stornirana; nosi paletu koja time postaje osirocena
    # Sve tri imaju aktivnu zbirnu, pa NISU osirocene prijemnice -- lista
    # osirocenih ostaje prazna i meri bas ono sto treba (zbirna, ne kupac).
    "tblPrijemnica": [
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
    ],
    # DVE ispravke na cekanju, i to NAD OTPREMNICOM -- namerno ne nad
    # prijemnicom: detekcija ispravke prijemnice pita operatera kroz MsgBox, a
    # MsgBox u headless runu visi. Ovako se safe-stop pravilo ("dve ili vise na
    # cekanju = ne biraj naslepo") proverava nad istom deljenom rutinom, bez
    # ijednog dijaloga.
    "tblStornoVeze": [
        {"CorrectionID": "SV-TEST-1", "Mode": "ISPRAVKA_ODMAH", "Status": "PENDING",
         "OldDocType": "Otpremnica", "OldDocID": "OTP-TEST-2", "OldBroj": "2/TEST",
         "NeedsRecovery": "Da", "Message": "fixture: ispravka na cekanju"},
        {"CorrectionID": "SV-TEST-2", "Mode": "ISPRAVKA_ODMAH", "Status": "PENDING",
         "OldDocType": "Otpremnica", "OldDocID": "OTP-TEST-3", "OldBroj": "3/TEST",
         "NeedsRecovery": "Da", "Message": "fixture: druga ispravka na cekanju"},
    ],
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
}

# DEFAULT_VRSTA_VOCA / DEFAULT_SORTA_VOCA se NAMERNO ne postavljaju: bez njih
# ApplyDefaultProizvod ostavlja combo-e prazne (frmOtkup ga zove pod
# On Error Resume Next), pa Initialize ne okida auto-cenu i stanje forme je
# deterministicno za golden snapshot.

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
        if r is None:
            row = lo.ListRows.Add()
            row.Range.Cells(1, kcol).Value = key
            row.Range.Cells(1, vcol).Value = val
            if akt:
                row.Range.Cells(1, akt).Value = STATUS_AKTIVAN
        else:
            lo.ListRows(r).Range.Cells(1, vcol).Value = val
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
