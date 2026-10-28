"""Testovi alata za migraciju -- bez Excela, nad sintetickom sveskom.

Fixture oznake (pre-flight par. 5):
  NORMALNO  -- normal production state: oblik koji stari pisci stvarno prave
               (dvostrani Otkup/OM-*-Koop, jednostrani sa vozacem, storno).
  ANOMALIJA -- synthetic anomaly / fault injection: iz njih se ne izvodi
               poslovna invarijanta, samo se meri da alat ne cuti.

    python tools/migracija/migracija.py --self-test
"""

import contextlib
import io
import os
import sys
import tempfile
import unittest
from datetime import date
from decimal import Decimal

OVDE = os.path.dirname(os.path.abspath(__file__))
if OVDE not in sys.path:
    sys.path.insert(0, OVDE)

import ambalaza  # noqa: E402
import cilj  # noqa: E402
import dokumenti  # noqa: E402
import provere  # noqa: E402
import upis  # noqa: E402
import migracija  # noqa: E402
import pregled  # noqa: E402
from citac import datum, polje, procitaj_svesku, sha256_fajla, tabela, tekst  # noqa: E402

D1, D2, D3 = date(2026, 6, 20), date(2026, 6, 21), date(2026, 6, 22)

AMB_KOL = ["AmbID", "Datum", "TipAmbalaze", "Kolicina", "Smer", "EntitetID", "EntitetTip",
           "VozacID", "DokumentID", "DokumentTIP", "Stornirano", "CreatedAt", "CreatedBy"]


def noga(i, d, kol, smer, ent, tip_ent, vozac, dok_id, dok_tip, storno="", tip="12/1"):
    return ["AMB-%05d" % i, d, tip, kol, smer, ent, tip_ent, vozac, dok_id, dok_tip, storno, "", "test"]


def normalno():
    """NORMALNO: jedna mala sezona u starom modelu."""
    return [
        # dan 1: otkup, kooperant donosi 10 (svoje) gajbe, dobija 4 prazne uz otkup
        noga(1, D1, 10, "Izlaz", "K1", "Kooperant", "", "OTK-1", "Otkup"),
        noga(2, D1, 10, "Ulaz", "S1", "Stanica", "", "OTK-1", "Otkup"),
        noga(3, D1, 4, "Ulaz", "K1", "Kooperant", "", "OTK-1", "OM-Izlaz-Koop"),
        noga(4, D1, 4, "Izlaz", "S1", "Stanica", "", "OTK-1", "OM-Izlaz-Koop"),
        # dan 2: revers -- stanica izdaje 8 praznih; otpremnica 10 sa vozacem
        noga(5, D2, 8, "Ulaz", "K1", "Kooperant", "", "R-1", "OM-Izlaz-Koop"),
        noga(6, D2, 8, "Izlaz", "S1", "Stanica", "", "R-1", "OM-Izlaz-Koop"),
        noga(7, D2, 10, "Izlaz", "S1", "Stanica", "D1", "OTP-1", "Otpremnica"),
        # dan 3: prijemnica (pune 10, vraca 7 praznih), vozac vraca 7 na stanicu
        noga(8, D3, 10, "Ulaz", "C1", "Kupac", "D1", "PRJ-1", "Prijemnica"),
        noga(9, D3, 7, "Izlaz", "C1", "Kupac", "D1", "PRJ-1", "Prijemnica"),
        noga(10, D3, 7, "Ulaz", "S1", "Stanica", "D1", "F-1", "OM-Ulaz-Firma"),
        # storniran otkup: obe noge u arhivu
        noga(11, D3, 5, "Izlaz", "K2", "Kooperant", "", "OTK-2", "Otkup", "Da"),
        noga(12, D3, 5, "Ulaz", "S1", "Stanica", "", "OTK-2", "Otkup", "Da"),
    ]


def osnovne_tabele(amb_redovi, dodatno=None):
    t = {
        "tblKooperanti": (["KooperantID", "Ime", "Prezime"],
                          [["K1", "Pera", "Peric"], ["K2", "Mika", "Mikic"]]),
        "tblStanice": (["StanicaID", "Naziv"], [["S1", "Stanica Jedan"], ["S2", "Stanica Dva"]]),
        "tblKupci": (["KupacID", "Naziv"], [["C1", "Hladnjaca"]]),
        "tblVozaci": (["VozacID", "Ime", "Prezime"], [["D1", "Vlada", "Vozac"]]),
        "tblTipAmbalaze": (["TipAmbalaze"], [["12/1"], ["6/1"]]),
        "tblOtkup": (["OtkupID", "Datum", "KooperantID", "StanicaID", "BrojDokumenta", "Klasa",
                      "Kolicina", "Cena", "KolAmbalaze", "KolAmbIzdata", "TipAmbalaze", "Stornirano",
                      "OtpremnicaID", "BrojZbirne", "Novac", "Isplaceno", "CreatedAt", "CreatedBy"],
                     [["OTK-1", D1, "K1", "S1", "101", "I", 100, 80, 10, 4, "12/1", "", "OTP-1", "301",
                       0, "", "2026-06-20 08:00:00", "op"],
                      ["OTK-2", D3, "K2", "S1", "102", "I", 50, 80, 5, 0, "12/1", "Da", "", "", 0, "", "", ""],
                      ["OTK-3", D3, "K2", "S1", "103", "I", 40, 80, 3, 0, "12/1", "", "", "", 0, "Da", "", ""]]),
        "tblOtpremnica": (["OtpremnicaID", "Datum", "StanicaID", "VozacID", "BrojOtpremnice", "Klasa",
                           "Kolicina", "KolAmbalaze", "Cena", "Stornirano", "GeneracijaID", "BrojZbirne"],
                          [["OTP-1", D2, "S1", "D1", "201", "I", 100, 10, 0, "", "G-OTP-1", "301"],
                           ["OTP-9", D2, "S1", "D1", "209", "I", 30, 3, 75, "Da", "G-OTP-9", ""]]),
        "tblZbirna": (["ZbirnaID", "Datum", "VozacID", "KupacID", "BrojZbirne", "Klasa", "UkupnoKolicina",
                       "UkupnoAmbalaze", "Stornirano", "GeneracijaID"],
                      [["ZBR-1", D2, "D1", "C1", "301", "I", 100, 10, "", "G-ZBR-1"]]),
        "tblPrijemnica": (["PrijemnicaID", "Datum", "KupacID", "VozacID", "BrojPrijemnice", "BrojZbirne",
                           "Klasa", "Kolicina", "Cena", "KolAmbalaze", "KolAmbVracena", "Stornirano",
                           "GeneracijaID", "Fakturisano"],
                          [["PRJ-1", D3, "C1", "D1", "401", "301", "I", 90, 60, 10, 7, "", "G-PRJ-1", "Da"],
                           ["PRJ-2", D3, "C1", "D1", "401", "301", "II", 10, 40, 0, 0, "", "G-PRJ-1", ""]]),
        "tblNovac": (["NovacID", "Datum", "Tip", "KooperantID", "OtkupID", "Isplata", "Uplata", "Stornirano"],
                     [["NOV-00001", D2, "VirmanFirmaKoop", "K1", "OTK-1", 8000, 0, ""]]),
        "tblPaleta": (["PaletaID", "BrojPalete", "Godina", "Status", "Stornirano", "Preradjeno"],
                      [["PAL-00001", "1", 2026, "Zatvorena", "", "Da"]]),
        "tblPaletaStavka": (["StavkaID", "PaletaID", "PrijemnicaID", "BrojPrijemnice", "Klasa", "Stornirano"],
                            [["PLS-00001", "PAL-00001", "PRJ-1", "401", "I", ""]]),
        "tblFakture": (["FakturaID", "BrojFakture", "Datum", "KupacID", "Stornirano"],
                       [["FAK-00001", "1/2026", D3, "C1", ""]]),
        "tblFakturaStavke": (["StavkaID", "FakturaID", "PrijemnicaID", "Kolicina", "Cena", "Klasa", "Stornirano"],
                             [["FST-00001", "FAK-00001", "PRJ-1", 90, 60, "I", ""]]),
        "tblSEFConfig": (["ConfigKey", "ConfigValue"], [["MALINA_DEFAULT_KUPAC", "C1"]]),
        "tblAmbalaza": (AMB_KOL, amb_redovi),
    }
    t.update(dodatno or {})
    return t


def snimi_svesku(putanja, tabele):
    sv = pregled.Sveska()
    for ime, (kol, redovi) in tabele.items():
        sv.dodaj_list(ime.replace("tbl", "")[:28], kol, redovi, tabela=ime)
    sv.snimi(putanja)


class SaSveskom(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.dir.cleanup)

    def prevedi(self, amb_redovi, izuzeci=(), dodatno=None, partneri=None, knjizenje="verno"):
        put = os.path.join(self.dir.name, "stara.xlsx")
        snimi_svesku(put, osnovne_tabele(amb_redovi, dodatno))
        self.tabele = procitaj_svesku(put)
        rez = ambalaza.prevedi(self.tabele, izuzeci, knjizenje=knjizenje, politika_partneri=partneri)
        self.salda = {(s["nalog_tip"], s["nalog_id"], s["tip"]): s for s in ambalaza.izracunaj_salda(rez)}
        return rez

    def razlog_za(self, rez, redni):
        for z in rez.za_odluku:
            if redni in z["noge"]:
                return z["razlog"]
        return None

    def osnovne_tvrdnje(self, rez):
        self.assertFalse(any(ambalaza.proveri_pokrivenost(rez).values()),
                         ambalaza.proveri_pokrivenost(rez))
        for k, s in self.salda.items():
            self.assertEqual(s["razlika"], 0, "salda se ne zatvaraju za %s" % (k,))


class Citac(SaSveskom):
    def test_cita_tabele_datume_i_kolone_bez_obzira_na_velika_slova(self):
        put = os.path.join(self.dir.name, "s.xlsx")
        snimi_svesku(put, osnovne_tabele(normalno()))
        t = procitaj_svesku(put)
        amb = tabela(t, "TBLAMBALAZA")
        self.assertEqual(len(amb.redovi), 12)
        r = amb.redovi[0]
        self.assertEqual(polje(r, "DokumentTip"), "Otkup")       # kolona je DokumentTIP
        self.assertEqual(datum(polje(r, "Datum")), D1)          # Excel serijski broj
        self.assertEqual(r["_redni"], 1)


class NormalnoStanje(SaSveskom):
    def test_kretanja_vrste_i_zaglavlja(self):
        rez = self.prevedi(normalno())
        self.osnovne_tvrdnje(rez)
        self.assertEqual(rez.oblik, "STARI")
        self.assertEqual(rez.za_odluku, [])
        vrste = sorted((k.pravilo, k.vrsta, k.od_tip, k.na_tip) for k in rez.kretanja)
        self.assertEqual(vrste, sorted([
            ("OTKUP_PAR", "AMBALAZA_UZ_ROBU", "Kooperant", "Stanica"),
            ("OTKUP_IZDATE_PRAZNE", "IZDATA_PRAZNA", "Stanica", "Kooperant"),
            ("REVERS_IZDATE_PRAZNE", "IZDATA_PRAZNA", "Stanica", "Kooperant"),
            ("OTPREMNICA_VOZAC", "AMBALAZA_UZ_ROBU", "Stanica", "Vozac"),
            ("PRIJEMNICA_PUNE", "AMBALAZA_UZ_ROBU", "Vozac", "Kupac"),
            ("PRIJEMNICA_POVRAT", "POVRAT_PRAZNE", "Kupac", "Vozac"),
            ("REVERS_VOZAC_STANICA", "PRENOS_INTERNO", "Vozac", "Stanica"),
        ]))
        self.assertEqual(sorted(z["BrojDokumenta"] for z in rez.zaglavlja.values()), ["F-1", "R-1"])
        self.assertEqual({k.dok_tip for k in rez.kretanja if k.pravilo.startswith("OTKUP")}, {"Otkup"})
        self.assertEqual(ambalaza.proveri_invarijante(rez), [])

    def test_istorija_je_verna_bez_dopisanog_pokrica(self):
        """Podrazumevano (odluka 28.10.2026): stari minus ostaje, nista se ne dopisuje."""
        rez = self.prevedi(normalno())
        self.assertFalse([k for k in rez.kretanja if k.generisano])
        k1 = self.salda[("Kooperant", "K1", "12/1")]
        self.assertEqual((k1["stari"], k1["novi"], k1["pokrice"]), (Decimal(2), Decimal(2), Decimal(0)))
        self.assertEqual((k1["najmanji"], k1["najmanji_dan"]), (Decimal(-6), D1))

    def test_dug_kooperanta_po_dnevnom_zakljucku_na_zahtev(self):
        """MIG-16/17: dan 1 obaveza raste 6 (pokrice), dan 2 pada 6 (vracanje)."""
        rez = self.prevedi(normalno(), partneri=("Kooperant",))
        self.assertIn(("PODELA_VRACANJE_TUDJE", Decimal(6)),
                      [(k.pravilo, k.kolicina) for k in rez.kretanja])
        k1 = self.salda[("Kooperant", "K1", "12/1")]
        self.assertEqual((k1["stari"], k1["novi"], k1["pokrice"], k1["obaveza"]),
                         (Decimal(2), Decimal(8), Decimal(6), Decimal(0)))
        self.assertGreaterEqual(k1["najmanji"], 0)              # AMB-INV-07 posle pokrica
        self.assertGreaterEqual(k1["najmanja_obaveza"], 0)      # AMB-INV-09
        pok = [k for k in rez.kretanja if k.pravilo == "POKRICE_DUGA"]
        self.assertEqual([(k.kolicina, k.dok_tip, k.dok_id, k.datum) for k in pok],
                         [(Decimal(6), "Otkup", "OTK-1", D1)])

    def test_presek_iz_salda_na_kraju(self):
        """Presek dovodi na nulu SALDO NA KRAJU (-5), ne najdublji dan (-12)."""
        self.prevedi(normalno())
        s1 = self.salda[("Stanica", "S1", "12/1")]
        self.assertEqual((s1["stari"], s1["novi"]), (Decimal(-5), Decimal(-5)))
        self.assertEqual((s1["najmanji"], s1["najmanji_dan"]), (Decimal(-12), D2))
        self.assertEqual(pregled._predlog_preseka(s1), ("NABAVKA", Decimal(5)))
        k1 = self.salda[("Kooperant", "K1", "12/1")]
        self.assertEqual(pregled._predlog_preseka(k1), ("", Decimal(0)))

    def test_partner_u_minusu_na_kraju_dobija_ulaz_tudje(self):
        self.prevedi(normalno() + [noga(13, D3, 9, "Izlaz", "C1", "Kupac", "D1", "KO-1", "Kupci-Otpremnica")])
        c1 = self.salda[("Kupac", "C1", "12/1")]
        self.assertEqual(c1["novi"], Decimal(-6))
        self.assertEqual(pregled._predlog_preseka(c1), ("ULAZ_TUDJE_AMBALAZE", Decimal(6)))

    def test_vozac_je_ogledalo_starog_salda(self):
        self.prevedi(normalno())
        d1 = self.salda[("Vozac", "D1", "12/1")]
        self.assertEqual((d1["stari"], d1["novi"]), (Decimal(0), Decimal(0)))

    def test_storno_ide_u_arhivu(self):
        rez = self.prevedi(normalno())
        self.assertEqual([n.ishod for n in rez.noge if n.redni in (11, 12)], ["STORNO", "STORNO"])
        self.assertNotIn(("Kooperant", "K2", "12/1"), {k for k, s in self.salda.items() if s["novi"]})

    def test_isti_ulaz_isti_id(self):
        a = sorted(k.amb_id for k in self.prevedi(normalno()).kretanja)
        b = sorted(k.amb_id for k in self.prevedi(normalno()).kretanja)
        self.assertEqual(a, b)
        self.assertTrue(all(x.startswith("AMB-") and len(x) == 36 for x in a))

    def test_stavke_po_nalogu_zavrsavaju_konacnim_saldom(self):
        rez = self.prevedi(normalno())
        st = ambalaza.stavke_po_nalogu(rez)
        for kl, lst in st.items():
            s = next(v for k, v in self.salda.items() if ambalaza._kl(k) == kl)
            self.assertEqual(lst[-1]["saldo_novi"], s["novi"], kl)
            self.assertEqual(lst[-1]["saldo_stari"], s["stari"], kl)


KUPAC_VRACA_VISE = noga(13, D3, 9, "Izlaz", "C1", "Kupac", "D1", "KO-1", "Kupci-Otpremnica")


class Ispravno(SaSveskom):
    """--knjizenje ispravno: sveska spremna za rad -- nijedan nalog ni dana ispod nule."""

    def test_nijedan_nalog_ni_jednog_dana_ispod_nule(self):
        rez = self.prevedi(normalno() + [KUPAC_VRACA_VISE], knjizenje="ispravno")
        self.osnovne_tvrdnje(rez)
        self.assertEqual([k for k, s in self.salda.items() if s["najmanji"] < 0], [])
        self.assertEqual(ambalaza.proveri_invarijante(rez), [])

    def test_stanica_dobija_jednu_nabavku_prvog_dana(self):
        rez = self.prevedi(normalno(), knjizenje="ispravno")
        nab = [k for k in rez.kretanja if k.vrsta == "NABAVKA"]
        self.assertEqual([(k.na_tip, k.na_id, k.kolicina, k.datum) for k in nab],
                         [("Stanica", "S1", Decimal(12), D1)])
        z = rez.zaglavlja[nab[0].dok_id]
        self.assertEqual((z["Vrsta"], z["BrojOwnerTip"], z["BrojOwnerID"]), ("NABAVKA", "Stanica", "S1"))
        s1 = self.salda[("Stanica", "S1", "12/1")]
        self.assertEqual((s1["stari"], s1["nabavka"], s1["novi"]), (Decimal(-5), Decimal(12), Decimal(7)))

    def test_kupac_dobija_pokrice_u_trenutku_manjka(self):
        rez = self.prevedi(normalno() + [KUPAC_VRACA_VISE], knjizenje="ispravno")
        pok = [k for k in rez.kretanja if k.pravilo == "POKRICE_MANJKA"]
        self.assertEqual(sum(k.kolicina for k in pok), Decimal(6))
        self.assertTrue(all(k.na_tip == "Kupac" and k.od_tip == "SpoljniSvet" for k in pok))
        c1 = self.salda[("Kupac", "C1", "12/1")]
        self.assertEqual((c1["stari"], c1["pokrice"], c1["novi"], c1["obaveza"]),
                         (Decimal(-6), Decimal(6), Decimal(0), Decimal(6)))

    def test_bez_nabavke_provera_pada(self):
        """Dvosmerni dokaz: bez NABAVKE stanica je u minusu i provera to vidi."""
        staro = ambalaza.nabavka_na_pocetku
        ambalaza.nabavka_na_pocetku = lambda rez, tipovi: []
        self.addCleanup(setattr, ambalaza, "nabavka_na_pocetku", staro)
        self.prevedi(normalno(), knjizenje="ispravno")
        self.assertEqual([k for k, s in self.salda.items() if s["najmanji"] < 0], [("Stanica", "S1", "12/1")])


class Anomalije(SaSveskom):
    def test_noga_bez_para_ide_za_odluku(self):
        rez = self.prevedi(normalno() + [noga(13, D3, 3, "Izlaz", "K2", "Kooperant", "", "OTK-3", "Otkup")])
        self.osnovne_tvrdnje(rez)
        self.assertEqual(self.razlog_za(rez, 13), "NOGA_BEZ_PARA")
        k2 = self.salda[("Kooperant", "K2", "12/1")]
        self.assertEqual((k2["stari"], k2["van_knjige"], k2["novi"]), (Decimal(-3), Decimal(-3), Decimal(0)))

    def test_nepoznat_tip_dokumenta(self):
        rez = self.prevedi(normalno() + [noga(13, D3, 2, "Ulaz", "S1", "Stanica", "", "X", "Nesto")])
        self.osnovne_tvrdnje(rez)
        self.assertEqual(self.razlog_za(rez, 13), "NEPOZNAT_TIP_DOKUMENTA")

    def test_omulaz_bez_odluke(self):
        rez = self.prevedi(normalno() + [noga(13, D3, 2, "Ulaz", "S1", "Stanica", "", "OM-5", "OMUlaz")])
        self.assertEqual(self.razlog_za(rez, 13), "OMULAZ_IZVOR_NEPOZNAT")

    def test_aktivna_noga_storniranog_dokumenta(self):
        rez = self.prevedi(normalno() + [noga(13, D2, 3, "Izlaz", "S1", "Stanica", "D1", "OTP-9", "Otpremnica")])
        self.osnovne_tvrdnje(rez)
        self.assertEqual(self.razlog_za(rez, 13), "AKTIVNA_NOGA_STORNIRANOG_DOKUMENTA")

    def test_dvosmisleno_sparivanje(self):
        dodatno = [noga(13, D3, 6, "Ulaz", "K1", "Kooperant", "", "R-9", "OM-Izlaz-Koop"),
                   noga(14, D3, 6, "Ulaz", "K2", "Kooperant", "", "R-9", "OM-Izlaz-Koop"),
                   noga(15, D3, 6, "Izlaz", "S1", "Stanica", "", "R-9", "OM-Izlaz-Koop"),
                   noga(16, D3, 6, "Izlaz", "S2", "Stanica", "", "R-9", "OM-Izlaz-Koop")]
        rez = self.prevedi(normalno() + dodatno)
        self.osnovne_tvrdnje(rez)
        self.assertEqual({self.razlog_za(rez, r) for r in (13, 14, 15, 16)}, {"DVOSMISLENO_SPARIVANJE"})

    def test_nalog_van_maticne(self):
        rez = self.prevedi(normalno() + [noga(13, D3, 2, "Ulaz", "S7", "Stanica", "D1", "F-2", "OM-Ulaz-Firma")])
        self.assertEqual(self.razlog_za(rez, 13), "NALOG_NIJE_U_MATICNOJ")

    def test_dva_ista_reversa_u_istom_zaglavlju_krse_amb_inv_04(self):
        dodatno = [noga(13, D3, 2, "Ulaz", "K1", "Kooperant", "", "R-2", "OM-Izlaz-Koop"),
                   noga(14, D3, 2, "Izlaz", "S1", "Stanica", "", "R-2", "OM-Izlaz-Koop"),
                   noga(15, D3, 3, "Ulaz", "K1", "Kooperant", "", "R-2", "OM-Izlaz-Koop"),
                   noga(16, D3, 3, "Izlaz", "S1", "Stanica", "", "R-2", "OM-Izlaz-Koop")]
        rez = self.prevedi(normalno() + dodatno)
        inv = {x["inv"] for x in ambalaza.proveri_invarijante(rez)}
        self.assertIn("AMB-INV-04", inv)


class Izuzeci(SaSveskom):
    OMU = noga(13, D3, 30, "Ulaz", "S1", "Stanica", "", "OM-7", "OMUlaz")

    def izuzetak(self, **izmena):
        iz = {"AmbID": "AMB-00013", "radnja": "kretanje",
              "provera": {"Datum": D3.isoformat(), "DokumentTip": "OMUlaz", "EntitetID": "S1",
                          "Smer": "Ulaz", "TipAmbalaze": "12/1", "Kolicina": 30},
              "kretanje": {"Od": ["Stanica", "S2"], "Na": ["Stanica", "S1"],
                           "VrstaKretanja": "PRENOS_INTERNO", "Dokument": "REVERS"},
              "razlog": "test"}
        iz.update(izmena)
        return iz

    def test_kretanje_iz_izuzetka(self):
        rez = self.prevedi(normalno() + [self.OMU], izuzeci=[self.izuzetak()])
        self.osnovne_tvrdnje(rez)
        k = [x for x in rez.kretanja if x.pravilo == "IZUZETAK"]
        self.assertEqual([(x.od_id, x.na_id, x.vrsta, x.dok_tip) for x in k],
                         [("S2", "S1", "PRENOS_INTERNO", "AmbalazaDokument")])
        self.assertEqual(self.salda[("Stanica", "S2", "12/1")]["izuzeci"], Decimal(-30))

    def test_promenjen_red_ponistava_odluku(self):
        prom = list(self.OMU)
        prom[3] = 31
        rez = self.prevedi(normalno() + [prom], izuzeci=[self.izuzetak()])
        self.osnovne_tvrdnje(rez)
        self.assertEqual(self.razlog_za(rez, 13), "IZUZETAK_POTPIS_SE_NE_SLAZE")

    def test_izuzmi(self):
        rez = self.prevedi(normalno() + [self.OMU], izuzeci=[self.izuzetak(radnja="izuzmi")])
        self.osnovne_tvrdnje(rez)
        self.assertEqual([n.ishod for n in rez.noge if n.redni == 13], ["IZUZETO"])

    def test_nabavka_iz_izuzetka_ima_nabavka_dokument(self):
        iz = self.izuzetak(kretanje={"Od": ["SpoljniSvet", ""], "Na": ["Stanica", "S1"],
                                     "VrstaKretanja": "NABAVKA", "Dokument": "NABAVKA"})
        rez = self.prevedi(normalno() + [self.OMU], izuzeci=[iz])
        self.osnovne_tvrdnje(rez)
        k = [x for x in rez.kretanja if x.pravilo == "IZUZETAK"][0]
        z = rez.zaglavlja[k.dok_id]
        self.assertEqual((k.vrsta, z["Vrsta"], z["BrojOwnerTip"], z["BrojOwnerID"]),
                         ("NABAVKA", "NABAVKA", "Stanica", "S1"))
        self.assertEqual(self.salda[("Stanica", "S1", "12/1")]["izuzeci"], Decimal(0))

    def test_isporuceni_izuzeci_su_ispravni(self):
        lista = ambalaza.ucitaj_izuzetke(migracija.IZUZECI)
        self.assertEqual(len(lista), 7)
        self.assertTrue(all(iz.get("provera") for iz in lista))
        self.assertEqual([d["OtpremnicaID"] for d in cilj.ucitaj_izuzetke_dokumenata(migracija.IZUZECI)],
                         ["OTP-00666"])


class Dokumenti(SaSveskom):
    def test_dokumenti_sa_vise_redova(self):
        self.prevedi(normalno())
        saz, grupe = dokumenti.identitet_dokumenata(self.tabele)
        prj = next(d for d in saz if d["tabela"] == "tblPrijemnica")
        self.assertEqual((prj["redova"], prj["logickih"], prj["sa_vise_redova"]), (2, 1, 1))
        self.assertEqual([g["id"] for g in grupe], ["PRJ-1, PRJ-2"])

    def test_popis_tabela_naspram_kanona(self):
        self.prevedi(normalno())
        popis = {p["tabela"]: p for p in dokumenti.popis_tabela(self.tabele, dokumenti.ucitaj_kanon())}
        self.assertIn("Smer", popis["tblAmbalaza"]["visak"])
        self.assertIn("OdNalogTip", popis["tblAmbalaza"]["fali"])
        self.assertEqual(popis["tblOtkupStavke"]["u_svesci"], "NE")
        self.assertEqual(dokumenti.oblik_dokumenata(self.tabele), "STARI")


class Amb10Oblik(SaSveskom):
    def test_vec_prevedena_knjiga_se_samo_sabira(self):
        kol = ["AmbID", "Datum", "TipAmbalaze", "Kolicina", "DokumentID", "DokumentTIP", "OdNalogTip",
               "OdNalogID", "NaNalogTip", "NaNalogID", "VrstaKretanja", "StornoOd"]
        red = [["AMB-A", D1, "12/1", 5, "OTK-1", "Otkup", "Kooperant", "K1", "Stanica", "S1",
                "AMBALAZA_UZ_ROBU", ""]]
        rez = self.prevedi([], dodatno={"tblAmbalaza": (kol, red)})
        self.assertEqual(rez.oblik, "AMB10")
        self.assertEqual(self.salda[("Stanica", "S1", "12/1")]["novi"], Decimal(5))
        self.assertEqual(self.salda[("Stanica", "S1", "12/1")]["razlika"], 0)


class NovaSveska(SaSveskom):
    """Faza 2: nova sveska u memoriji (cilj.py) i provere pre upisa (provere.py)."""

    def sagradi(self, amb=None, dodatno=None):
        put = os.path.join(self.dir.name, "stara.xlsx")
        snimi_svesku(put, osnovne_tabele(normalno() if amb is None else amb, dodatno))
        self.tabele = procitaj_svesku(put)
        c = cilj.izgradi(self.tabele, dokumenti.ucitaj_kanon())
        provere.proveri(c)
        provere.proveri_isplate_iz_starog(c, self.tabele)
        return c

    def nalazi(self, c, tezina=None):
        return sorted((n["pravilo"], tekst(n["id"])) for n in c.nalazi if tezina in (None, n["tezina"]))

    def test_zaglavlje_i_jedna_stavka_po_dokumentu(self):
        c = self.sagradi()
        self.assertEqual([z["OtkupID"] for z in c.tabele["tblOtkup"]], ["OTK-1", "OTK-2", "OTK-3"])
        self.assertNotIn("Kolicina", c.tabele["tblOtkup"][0])
        st = {s["OtkupID"]: s for s in c.tabele["tblOtkupStavke"]}
        self.assertEqual((st["OTK-1"]["Klasa"], st["OTK-1"]["Kolicina"], st["OTK-1"]["Cena"]),
                         ("I", Decimal(100), Decimal(80)))
        self.assertTrue(st["OTK-1"]["OtkupStavkaID"].startswith("OKS-"))
        self.assertEqual(len(st["OTK-1"]["OtkupStavkaID"]), 36)
        self.assertEqual(c.tabele["tblOtkup"][0]["OtpremnicaID"], "")      # pisac ga ne pise
        zst = c.tabele["tblZbirnaStavke"][0]
        self.assertEqual((zst["Kolicina"], zst["KolAmbalaze"]), (Decimal(100), Decimal(10)))
        self.assertEqual(c.tabele["tblZbirna"][0]["UkupnoKolicina"], "")
        ost = {s["OtpremnicaID"]: s for s in c.tabele["tblOtpremnicaStavke"]}
        self.assertEqual((ost["OTP-1"]["PredlogCena"], ost["OTP-9"]["PredlogCena"]), ("", Decimal(75)))
        self.assertEqual(len(c.tabele["tblPrijemnicaStavke"]), 2)
        self.assertNotIn("GeneracijaID", c.tabele["tblPrijemnica"][0])

    def test_clanstvo_status_i_veze(self):
        c = self.sagradi()
        self.assertEqual([(v["OtkupID"], v["OtpremnicaID"]) for v in c.tabele["tblOtpremnicaIzvori"]],
                         [("OTK-1", "OTP-1")])
        self.assertEqual([(v["OtpremnicaID"], v["ZbirnaID"]) for v in c.tabele["tblZbirnaIzvori"]],
                         [("OTP-1", "ZBR-1")])
        status = {z["OtkupID"]: z["IzdatoStatus"] for z in c.tabele["tblOtkup"]}
        self.assertEqual(status, {"OTK-1": "IZDATO", "OTK-2": "IZDATO", "OTK-3": ""})
        self.assertEqual({z["IzdatoStatus"] for z in c.tabele["tblOtpremnica"]}, {"IZDATO"})
        self.assertEqual({z["ZbirnaRoditeljID"] for z in c.tabele["tblPrijemnica"]}, {"ZBR-1"})
        self.assertEqual(c.tabele["tblZbirna"][0]["KupacHladnjaca"], "Da")
        sid = {s["PrijemnicaID"]: s["PrijemnicaStavkaID"] for s in c.tabele["tblPrijemnicaStavke"]}
        self.assertEqual(c.tabele["tblPaletaStavka"][0]["PrijemnicaStavkaID"], sid["PRJ-1"])
        self.assertEqual(c.tabele["tblFakturaStavke"][0]["PrijemnicaStavkaID"], sid["PRJ-1"])
        self.assertTrue(c.tabele["tblAmbalaza"])

    def test_normalna_sveska_nema_blokade(self):
        c = self.sagradi()
        self.assertEqual(self.nalazi(c, "BLOKIRA"), [])
        self.assertIn(("AKTIVAN_OTKUP_BEZ_OTPREMNICE", "OTK-3"), self.nalazi(c, "ODLUKA"))
        self.assertIn(("STARI_ISPLACENO_BEZ_NOVCA", "OTK-3"), self.nalazi(c, "ODLUKA"))

    def test_storniran_otkup_sa_nula_kg_se_izostavlja(self):
        """MIG-ODL-09: storniran 0 kg ne ide u novu svesku (ni zaglavlje ni stavka)."""
        t = osnovne_tabele(normalno())
        kol, redovi = t["tblOtkup"]
        redovi[1][6] = 0
        c = self.sagradi(dodatno={"tblOtkup": (kol, redovi)})
        self.assertEqual(self.nalazi(c, "BLOKIRA"), [])
        self.assertIn(("STORNIRAN_OTKUP_BEZ_KOLICINE_IZOSTAVLJEN", "OTK-2"), self.nalazi(c, "ODLUKA"))
        self.assertNotIn("OTK-2", [z["OtkupID"] for z in c.tabele["tblOtkup"]])
        self.assertNotIn("OTK-2", [s["OtkupID"] for s in c.tabele["tblOtkupStavke"]])

    def test_aktivan_otkup_sa_nula_kg_i_dalje_blokira(self):
        t = osnovne_tabele(normalno())
        kol, redovi = t["tblOtkup"]
        redovi[2][6] = 0
        c = self.sagradi(dodatno={"tblOtkup": (kol, redovi)})
        self.assertIn(("KOLICINA_NIJE_POZITIVNA", "OTK-3"), self.nalazi(c, "BLOKIRA"))

    def test_otvorena_paleta_se_zatvara(self):
        t = osnovne_tabele(normalno())
        kol, redovi = t["tblPaleta"]
        redovi.append(["PAL-00002", "2", 2026, "Otvorena", "", ""])
        c = self.sagradi(dodatno={"tblPaleta": (kol, redovi)})
        st = {p["PaletaID"]: p["Status"] for p in c.tabele["tblPaleta"]}
        self.assertEqual(st, {"PAL-00001": "Zatvorena", "PAL-00002": "Zatvorena"})
        self.assertIn(("PALETA_ZATVORENA_PRI_MIGRACIJI", "PAL-00002"), self.nalazi(c, "INFO"))
        self.assertNotIn("OTVORENA_PALETA", [p for p, _ in self.nalazi(c)])

    def test_klasa_iii_blokira(self):
        t = osnovne_tabele(normalno())
        kol, redovi = t["tblOtkup"]
        redovi[2][5] = "III"
        c = self.sagradi(dodatno={"tblOtkup": (kol, redovi)})
        self.assertIn(("KLASA_NIJE_I_II", "OTK-3"), self.nalazi(c, "BLOKIRA"))

    def otp2(self, otkup_202=None, izuzeci_dok=()):
        """OTP-2 (broj 202) bez veze; opciono otkup OTK-4 sa istim brojem i potpisom."""
        t = osnovne_tabele(normalno())
        okol, ored = t["tblOtkup"]
        okol = okol + ["VrstaVoca", "SortaVoca"]
        ored = [r + ["Malina", "Vilamet"] for r in ored]
        if otkup_202:
            ored.append(otkup_202)
        pkol, pred = t["tblOtpremnica"]
        pkol = pkol + ["VrstaVoca", "SortaVoca", "TipAmbalaze"]
        pred = [r + ["Malina", "Vilamet", "12/1"] for r in pred]
        pred.append(["OTP-2", D3, "S1", "D1", "202", "I", 20, 2, 0, "", "", "", "Malina", "Vilamet", "12/1"])
        put = os.path.join(self.dir.name, "stara.xlsx")
        snimi_svesku(put, osnovne_tabele(normalno(), {"tblOtkup": (okol, ored), "tblOtpremnica": (pkol, pred)}))
        self.tabele = procitaj_svesku(put)
        c = cilj.izgradi(self.tabele, dokumenti.ucitaj_kanon(), izuzeci_dok=izuzeci_dok)
        provere.proveri(c)
        return c

    OTK4 = ["OTK-4", D3, "K1", "S1", "202", "I", 20, 80, 2, 0, "12/1", "", "", "", 0, "", "", "",
            "Malina", "Vilamet"]

    def test_otpremnica_bez_otkupa_ostaje_neizdata(self):
        c = self.otp2()
        self.assertEqual(self.nalazi(c, "BLOKIRA"), [])
        self.assertIn(("OTPREMNICA_BEZ_OTKUPA_OSTAJE_NEIZDATA", "OTP-2"), self.nalazi(c, "ODLUKA"))
        self.assertEqual({z["OtpremnicaID"]: z["IzdatoStatus"] for z in c.tabele["tblOtpremnica"]}["OTP-2"], "")

    def test_veza_se_popravlja_kad_je_otkup_dokazan_1_1(self):
        c = self.otp2(self.OTK4)
        self.assertIn(("VEZA_POPRAVLJENA_MIG21", "OTP-2"), self.nalazi(c, "INFO"))
        self.assertIn(("OTK-4", "OTP-2"), [(v["OtkupID"], v["OtpremnicaID"])
                                           for v in c.tabele["tblOtpremnicaIzvori"]])
        st = {z["OtkupID"]: z["IzdatoStatus"] for z in c.tabele["tblOtkup"]}
        self.assertEqual(st["OTK-4"], "IZDATO")
        self.assertEqual(self.nalazi(c, "BLOKIRA"), [])

    def test_razlicita_kolicina_nije_dokaz(self):
        drugi = list(self.OTK4)
        drugi[6] = 21
        c = self.otp2(drugi)
        self.assertNotIn(("VEZA_POPRAVLJENA_MIG21", "OTP-2"), self.nalazi(c, "INFO"))
        self.assertIn(("OTPREMNICA_BEZ_OTKUPA_OSTAJE_NEIZDATA", "OTP-2"), self.nalazi(c, "ODLUKA"))

    def test_izuzetak_bez_popravke_veze(self):
        c = self.otp2(self.OTK4, izuzeci_dok=[{"OtpremnicaID": "OTP-2", "radnja": "bez_popravke_veze",
                                                "razlog": "test"}])
        self.assertIn(("VEZA_SE_NE_POPRAVLJA_IZUZETAK", "OTP-2"), self.nalazi(c, "INFO"))
        self.assertNotIn(("VEZA_POPRAVLJENA_MIG21", "OTP-2"), self.nalazi(c, "INFO"))

    def test_izdata_otpremnica_bez_otkupa_blokira_u_proveri(self):
        """Dvosmerni dokaz: provera i dalje hvata IZDATU otpremnicu bez izvora."""
        c = self.otp2()
        next(z for z in c.tabele["tblOtpremnica"] if z["OtpremnicaID"] == "OTP-2")["IzdatoStatus"] = "IZDATO"
        c.nalazi = []
        provere.proveri(c)
        self.assertIn(("IZDATA_OTPREMNICA_BEZ_AKTIVNOG_OTKUPA", "OTP-2"), self.nalazi(c, "BLOKIRA"))

    def test_broj_dokumenta_ide_kao_tekst(self):
        t = osnovne_tabele(normalno())
        kol, redovi = t["tblOtkup"]
        redovi[0][4] = 101
        c = self.sagradi(dodatno={"tblOtkup": (kol, redovi)})
        self.assertEqual(c.tabele["tblOtkup"][0]["BrojDokumenta"], "101")

    def test_isplata_bez_otkupa_je_avans(self):
        t = osnovne_tabele(normalno())
        kol, redovi = t["tblNovac"]
        redovi.append(["NOV-00002", D3, "VirmanAvansKoop", "K2", "", 500, 0, ""])
        c = self.sagradi(dodatno={"tblNovac": (kol, redovi)})
        self.assertIn(("ISPLATA_BEZ_OTKUPA_POSTAJE_AVANS", "NOV-00002"), self.nalazi(c, "ODLUKA"))

    def test_faktura_na_nepostojecu_prijemnicu_blokira(self):
        t = osnovne_tabele(normalno())
        kol, redovi = t["tblFakturaStavke"]
        redovi.append(["FST-00002", "FAK-00001", "PRJ-9", 5, 60, "I", ""])
        c = self.sagradi(dodatno={"tblFakturaStavke": (kol, redovi)})
        self.assertIn(("FAKTURA_STAVKA_BEZ_STAVKE_PRIJEMNICE", "FST-00002"), self.nalazi(c, "BLOKIRA"))

    def test_bez_stavke_provera_pada(self):
        """Dvosmerni dokaz: zaglavlje bez stavke mora da BLOKIRA."""
        c = self.sagradi()
        c.tabele["tblOtkupStavke"] = [s for s in c.tabele["tblOtkupStavke"] if s["OtkupID"] != "OTK-1"]
        c.nalazi = []
        provere.proveri(c)
        self.assertIn(("ZAGLAVLJE_BEZ_STAVKE", "OTK-1"), self.nalazi(c, "BLOKIRA"))

    def test_proba_komanda_pravi_paket(self):
        put = os.path.join(self.dir.name, "stara.xlsx")
        snimi_svesku(put, osnovne_tabele(normalno()))
        izlaz = os.path.join(self.dir.name, "proba")
        self.assertEqual(tiho(["proba", "--izvor", put, "--izlaz", izlaz, "--bez-izuzetaka"]), 0)
        t = procitaj_svesku(os.path.join(izlaz, "proba_upisa.xlsx"))
        self.assertEqual(len(tabela(t, "tblOtkupStavke").redovi), 3)
        self.assertEqual(datum(polje(tabela(t, "tblOtkup").redovi[0], "Datum")), D1)
        self.assertIsNotNone(tabela(t, "Provere"))


class Upis(SaSveskom):
    """Ciste funkcije upisa (bez Excela): pretvaranje, redosled kolona, poredjenje."""

    def test_com_vrednost_cuva_smisao(self):
        from datetime import datetime
        self.assertIsNone(upis.com_vrednost(""))
        self.assertEqual(upis.com_vrednost("101"), "'101")              # ostaje tekst
        self.assertEqual(upis.com_vrednost("2026-06-20 08:00:00"), "'2026-06-20 08:00:00")
        self.assertEqual(upis.com_vrednost("=1+1"), "'=1+1")             # ne postaje formula
        self.assertEqual(upis.com_vrednost("Pera"), "Pera")
        self.assertEqual(upis.com_vrednost(Decimal("100")), 100)
        self.assertEqual(upis.com_vrednost(Decimal("1.5")), 1.5)
        self.assertEqual(upis.com_vrednost(D1), datetime(2026, 6, 20))

    def test_niz_po_redosledu_kolona_u_svesci(self):
        redovi = [{"A": "x", "B": Decimal(2)}]
        self.assertEqual(upis.niz_za_tabelu(redovi, ["B", "A", "Runtime"]), [(2, "x", None)])
        with self.assertRaisesRegex(ValueError, "kolone sa podacima"):
            upis.niz_za_tabelu(redovi, ["A"])
        # tekstualna kolona: bez apostrofa (format "@" se postavlja pre upisa)
        self.assertEqual(upis.niz_za_tabelu([{"BrojDokumenta": "101", "Napomena": "101"}],
                                            ["BrojDokumenta", "Napomena"]), [("101", "'101")])

    def napravi_paket(self, c):
        put = os.path.join(self.dir.name, "nova.xlsx")
        sv = pregled.Sveska()
        for ime, redovi in sorted(c.tabele.items()):
            if redovi:
                kol = c.kanon[ime]
                sv.dodaj_list(ime[3:], kol, [[r.get(k, "") for k in kol] for r in redovi], tabela=ime)
        sv.snimi(put)
        return put

    def test_poredjenje_posle_upisa(self):
        put = os.path.join(self.dir.name, "stara.xlsx")
        snimi_svesku(put, osnovne_tabele(normalno()))
        c = cilj.izgradi(procitaj_svesku(put), dokumenti.ucitaj_kanon())
        paket = self.napravi_paket(c)
        self.assertEqual(upis.uporedi(c, paket), [])
        c.tabele["tblOtkupStavke"][0]["Kolicina"] = Decimal(101)        # dvosmerni dokaz
        r = upis.uporedi(c, paket)
        self.assertEqual(len(r), 1)
        self.assertIn("tblOtkupStavke red 1, Kolicina", r[0])

    def test_podesavanja_iz_stare_bez_aktivacije_licence(self):
        dodatno = {"tblSEFConfig": (["ConfigKey", "ConfigValue", "Aktivan"],
                                    [["MALINA_DEFAULT_KUPAC", "C1", "Aktivan"], ["MALINA_MODE", "YES", ""],
                                     ["LICENSE_KEY", "ABC", ""], ["LICENSE_ENABLED", "YES", ""]])}
        put = os.path.join(self.dir.name, "stara.xlsx")
        snimi_svesku(put, osnovne_tabele(normalno(), dodatno))
        c = cilj.izgradi(procitaj_svesku(put), dokumenti.ucitaj_kanon())
        kljucevi = [k for k, _ in c.podesavanja["tblSEFConfig"]]
        self.assertEqual(kljucevi, ["MALINA_DEFAULT_KUPAC", "MALINA_MODE", "LICENSE_ENABLED"])
        self.assertEqual(dict(c.podesavanja["tblSEFConfig"])["MALINA_DEFAULT_KUPAC"]["Aktivan"], "Aktivan")
        self.assertIn("AKTIVACIJA_LICENCE_SE_NE_PRENOSI", [n["pravilo"] for n in c.nalazi])

    def test_upis_odbija_blokadu_i_trazi_prihvatanje_odluka(self):
        put = os.path.join(self.dir.name, "stara.xlsx")
        snimi_svesku(put, osnovne_tabele(normalno()))
        meta = os.path.join(self.dir.name, "nova.xlsm")
        self.assertEqual(tiho(["upis", "--izvor", put, "--cilj", meta, "--bez-izuzetaka"]), 4)
        t = osnovne_tabele(normalno())
        kol, redovi = t["tblOtkup"]
        redovi[0][5] = "III"
        snimi_svesku(put, osnovne_tabele(normalno(), {"tblOtkup": (kol, redovi)}))
        self.assertEqual(tiho(["upis", "--izvor", put, "--cilj", meta, "--bez-izuzetaka",
                               "--prihvatam-odluke"]), 3)
        self.assertFalse(os.path.exists(meta))


class DvosmerniDokaz(SaSveskom):
    """Provere koje cuvaju pregled moraju umeti da padnu (CLAUDE.md par. 5)."""

    def test_pokrivenost_hvata_izgubljen_red(self):
        rez = self.prevedi(normalno())
        rez.noge[0].ishod = ""
        self.assertEqual(ambalaza.proveri_pokrivenost(rez)["bez_ishoda"], [1])

    def test_pokrivenost_hvata_nogu_u_dva_kretanja(self):
        rez = self.prevedi(normalno())
        rez.kretanja[1].noge = rez.kretanja[1].noge + rez.kretanja[0].noge[:1]
        self.assertTrue(ambalaza.proveri_pokrivenost(rez)["u_dva_kretanja"])

    def test_pogresan_smer_pravila_obara_proveru_efekta(self):
        staro = ambalaza.JEDNOSTRANE["otpremnica"]
        ambalaza.JEDNOSTRANE["otpremnica"] = {("Stanica", "izlaz"): (False, "AMBALAZA_UZ_ROBU", "SABOTAZA")}
        self.addCleanup(ambalaza.JEDNOSTRANE.__setitem__, "otpremnica", staro)
        with self.assertRaisesRegex(ambalaza.GreskaAmbalaze, "SABOTAZA"):
            self.prevedi(normalno())

    def test_salda_hvataju_kretanje_bez_traga(self):
        rez = self.prevedi(normalno())
        rez.kretanja[0].kolicina += 1
        razlike = [s for s in ambalaza.izracunaj_salda(rez) if s["razlika"] != 0]
        self.assertTrue(razlike)


def tiho(argv):
    with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
        return migracija.main(argv)


class Komanda(SaSveskom):
    def test_pregled_kraj_do_kraja_ne_menja_izvor(self):
        put = os.path.join(self.dir.name, "stara.xlsx")
        snimi_svesku(put, osnovne_tabele(normalno()))
        pre = sha256_fajla(put)
        izlaz = os.path.join(self.dir.name, "out")
        kod = tiho(["pregled", "--izvor", put, "--izlaz", izlaz, "--bez-izuzetaka"])   # ispravno
        self.assertEqual(kod, 0)
        self.assertEqual(sha256_fajla(put), pre)
        t = procitaj_svesku(os.path.join(izlaz, "pregled_migracije.xlsx"))
        for ime in ("Salda", "Stavke", "Kretanja", "Reversi", "ZaOdluku", "Dokumenti", "Tabele"):
            self.assertIsNotNone(tabela(t, ime), ime)
        k1 = [r for r in tabela(t, "Salda").redovi if polje(r, "ID") == "K1"]
        self.assertEqual(polje(k1[0], "Naziv"), "Pera Peric")
        self.assertEqual(len(tabela(t, "Kretanja").redovi), 10)       # 7 + pokrice + vracanje + nabavka
        with open(os.path.join(izlaz, "tabele.txt"), encoding="utf-8") as f:
            self.assertIn("tblAmbalaza (12 redova): AmbID, Datum", f.read())
        with open(os.path.join(izlaz, "sazetak.txt"), encoding="utf-8") as f:
            saz = f.read()
        self.assertRegex(saz, r"Nijedan nalog ni jednog dana ispod nule \(AMB-INV-07\) +DA")
        self.assertRegex(saz, r"dopisano NABAVKA \(stanica, vozac\) +1 dok\., 12 kom\.")

    def test_odbija_izlaz_u_repou(self):
        put = os.path.join(self.dir.name, "stara.xlsx")
        snimi_svesku(put, osnovne_tabele(normalno()))
        izlaz = os.path.join(migracija.KOREN_REPOA, "pregled_test_ne_sme")
        self.assertEqual(tiho(["pregled", "--izvor", put, "--izlaz", izlaz]), 2)
        self.assertFalse(os.path.exists(izlaz))


if __name__ == "__main__":
    unittest.main()
