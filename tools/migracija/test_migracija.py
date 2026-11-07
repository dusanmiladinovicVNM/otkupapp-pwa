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
from collections import Counter
from datetime import date
from decimal import Decimal

OVDE = os.path.dirname(os.path.abspath(__file__))
if OVDE not in sys.path:
    sys.path.insert(0, OVDE)

import ambalaza  # noqa: E402
import cilj  # noqa: E402
import dokumenti  # noqa: E402
import lanac  # noqa: E402
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
        """MIG-16/17: dan 1 obaveza raste 6 (pokrice), dan 2 pada 6 (vracanje + izlaz).

        Premisa promenjena sa AMB-INV-15 (#448): vracene gajbe izlaze iz opticaja,
        pa je fizicki saldo K1 = stari (2), a ne 8 -- 8 je bio KI-024 (vracene
        gajbe su ostajale na kooperantu, pa bi ih sledece donosenje prevelo u firmine).
        """
        rez = self.prevedi(normalno(), partneri=("Kooperant",))
        self.assertIn(("PODELA_VRACANJE_TUDJE", Decimal(6)),
                      [(k.pravilo, k.kolicina) for k in rez.kretanja])
        vr = next(k for k in rez.kretanja if k.vrsta == ambalaza.VRACANJE_TUDJE)
        iz = [k for k in rez.kretanja if k.vrsta == ambalaza.IZLAZ_TUDJE]
        self.assertEqual([(k.kolicina, k.dok_tip, k.dok_id, k.datum, k.od_id, k.na_tip) for k in iz],
                         [(vr.kolicina, vr.dok_tip, vr.dok_id, vr.datum, "K1", "SpoljniSvet")])
        k1 = self.salda[("Kooperant", "K1", "12/1")]
        self.assertEqual((k1["stari"], k1["novi"], k1["pokrice"], k1["izlaz"], k1["obaveza"]),
                         (Decimal(2), Decimal(2), Decimal(6), Decimal(-6), Decimal(0)))
        self.assertEqual([x for x in ambalaza.proveri_invarijante(rez)
                          if x["inv"] in ("AMB-INV-15", "AMB-BIL-02", "AMB-INV-10")], [])
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
        self.assertEqual([(d["OtpremnicaID"], d["radnja"]) for d in cilj.ucitaj_izuzetke_dokumenata(migracija.IZUZECI)
                          if d.get("OtpremnicaID")],
                         [("OTP-00666", "bez_popravke_veze"), ("OTP-00666", "storniraj")])


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


class SaNovomSveskom(SaSveskom):
    def sagradi(self, amb=None, dodatno=None, izuzeci_dok=()):
        put = os.path.join(self.dir.name, "stara.xlsx")
        snimi_svesku(put, osnovne_tabele(normalno() if amb is None else amb, dodatno))
        self.tabele = procitaj_svesku(put)
        c = cilj.izgradi(self.tabele, dokumenti.ucitaj_kanon(), izuzeci_dok=izuzeci_dok)
        provere.proveri(c)
        provere.proveri_isplate_iz_starog(c, self.tabele)
        return c

    def nalazi(self, c, tezina=None):
        return sorted((n["pravilo"], tekst(n["id"])) for n in c.nalazi if tezina in (None, n["tezina"]))


class NovaSveska(SaNovomSveskom):
    """Faza 2: nova sveska u memoriji (cilj.py) i provere pre upisa (provere.py)."""

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
        self.assertIn(("OTPREMNICA_CEKA_BLOKOVE", "OTP-2"), self.nalazi(c, "INFO"))
        self.assertEqual({z["OtpremnicaID"]: z["IzdatoStatus"] for z in c.tabele["tblOtpremnica"]}["OTP-2"], "")
        # bez zbirne i prijemnice nije dokazana kao prava -- odluka (MIG-ODL-18)
        self.assertIn(("OTPREMNICA_BEZ_BLOKOVA_BEZ_LANCA", "OTP-2"), self.nalazi(c, "ODLUKA"))

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
        self.assertIn(("OTPREMNICA_CEKA_BLOKOVE", "OTP-2"), self.nalazi(c, "INFO"))

    STORNIRAJ_OTP2 = {"OtpremnicaID": "OTP-2", "radnja": "storniraj", "razlog": "test",
                      "provera": {"BrojOtpremnice": "202", "Datum": D3.isoformat(), "StanicaID": "S1"}}

    def test_storniraj_po_odluci(self):
        """MIG-ODL-19: otpremnica bez blokova i bez prijemnice, proglasena storniranom,
        prenosi se sa Stornirano = Da -- nije vise nacrt koji ceka blokove."""
        c = self.otp2(izuzeci_dok=[self.STORNIRAJ_OTP2])
        otp = {z["OtpremnicaID"]: z for z in c.tabele["tblOtpremnica"]}
        self.assertEqual(otp["OTP-2"]["Stornirano"], "Da")
        self.assertIn(("OTPREMNICA_STORNIRANA_PO_ODLUCI", "OTP-2"), self.nalazi(c, "INFO"))
        pravila = [n[0] for n in self.nalazi(c) if n[1] == "OTP-2"]
        self.assertNotIn("OTPREMNICA_CEKA_BLOKOVE", pravila)
        self.assertNotIn("OTPREMNICA_BEZ_BLOKOVA_BEZ_LANCA", pravila)
        self.assertEqual(self.nalazi(c, "BLOKIRA"), [])

    def test_storniraj_ne_vazi_kad_se_potpis_ne_slaze(self):
        iz = dict(self.STORNIRAJ_OTP2, provera={"BrojOtpremnice": "202", "Datum": D2.isoformat(), "StanicaID": "S1"})
        c = self.otp2(izuzeci_dok=[iz])
        otp = {z["OtpremnicaID"]: z for z in c.tabele["tblOtpremnica"]}
        self.assertEqual(otp["OTP-2"]["Stornirano"], "")
        self.assertIn(("IZUZETAK_DOKUMENTA_POTPIS_SE_NE_SLAZE", "OTP-2"), self.nalazi(c, "ODLUKA"))

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

    def test_stornirana_linija_aktivne_fakture_ne_ozivljava(self):
        """S7 (FAK-ODL-03): linija nema svoj storno. Stara linija sa Stornirano = Da
        na aktivnoj fakturi ne ide u novu svesku -- inace bi ozivela."""
        c = self.sagradi(dodatno=izmeni("tblFakturaStavke", {"FST-00001": {"Stornirano": "Da"}}))
        self.assertEqual(c.tabele["tblFakturaStavke"], [])
        self.assertIn(("STORNIRANA_LINIJA_AKTIVNE_FAKTURE_IZOSTAVLJENA", "FST-00001"), self.nalazi(c, "ODLUKA"))
        self.assertNotIn("Stornirano", c.kanon["tblFakturaStavke"])

    def test_linija_stornirane_fakture_nije_aktivna(self):
        """Linija stornirane fakture bez stavke prijemnice ne obara strog citac
        (FakturaLinijePoStavci gleda samo aktivne); bez fakture -- obara uvek."""
        d = izmeni("tblFakture", {"FAK-00001": {"Stornirano": "Da"}})
        d.update(izmeni("tblFakturaStavke", novi=[
            dict(StavkaID="FST-00002", FakturaID="FAK-00001", PrijemnicaID="PRJ-9", Kolicina=5),
            dict(StavkaID="FST-00003", FakturaID="FAK-00009", PrijemnicaID="PRJ-1", Kolicina=5)]))
        c = self.sagradi(dodatno=d)
        blokira = self.nalazi(c, "BLOKIRA")
        self.assertNotIn(("FAKTURA_STAVKA_BEZ_STAVKE_PRIJEMNICE", "FST-00002"), blokira)
        self.assertIn(("FAKTURA_STAVKA_BEZ_FAKTURE", "FST-00003"), blokira)

    def test_stavka_prijemnice_na_dve_aktivne_fakture_blokira(self):
        d = izmeni("tblFakture", novi=[dict(FakturaID="FAK-00002", BrojFakture="2/2026", Datum=D3, KupacID="C1")])
        d.update(izmeni("tblFakturaStavke", novi=[
            dict(StavkaID="FST-00002", FakturaID="FAK-00002", PrijemnicaID="PRJ-1", Kolicina=90, Klasa="I")]))
        c = self.sagradi(dodatno=d)
        self.assertIn(("STAVKA_PRIJEMNICE_NA_DVE_AKTIVNE_FAKTURE", "FST-00002"), self.nalazi(c, "BLOKIRA"))

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
        with open(os.path.join(izlaz, "proba_sazetak.txt"), encoding="utf-8") as f:
            saz = f.read()
        # ODLUKA pojedinacno (dokument/nalog), ne samo broj -- o njima se odlucuje pre upisa
        self.assertIn("ODLUKA AKTIVAN_OTKUP_BEZ_OTPREMNICE tblOtkup OTK-3", saz.split("pojedinacno:")[1])


def izmeni(ime, izmene=None, novi=()):
    """Tabela iz osnovne sveske sa izmenjenim redovima ({ID: {kolona: vrednost}}) i
    dodatim redovima ({kolona: vrednost}), u obliku za `dodatno`."""
    kol, redovi = osnovne_tabele(normalno())[ime]
    redovi = [list(r) for r in redovi]
    for r in redovi:
        for c, v in (izmene or {}).get(r[0], {}).items():
            r[kol.index(c)] = v
    for n in novi:
        redovi.append([n.get(c, "") for c in kol])
    return {ime: (kol, redovi)}


class MainPosle448(SaNovomSveskom):
    """Knjiga ambalaze po pravilima main-a posle #445-#450 (MIG-ODL-13..15)."""

    def knjiga(self, c, dok_id):
        return [(r["VrstaKretanja"], r["OdNalogID"], r["NaNalogID"], r["Kolicina"])
                for r in c.tabele["tblAmbalaza"] if r["DokumentID"] == dok_id]

    def saldo(self, c, tip, idn, tip_amb="12/1"):
        return next(s for s in ambalaza.izracunaj_salda(c.amb)
                    if (s["nalog_tip"], s["nalog_id"], s["tip"]) == (tip, idn, tip_amb))

    def test_otpremnica_nosi_zbir_gajbi_blokova(self):
        """KI-065: stara otpremnica 12 gajbi, njen otkup 10 -> dokument i knjiga 10,
        zbirna 10; stanica nijednog dana nema pune ispod nule."""
        amb = normalno()
        amb[6] = noga(7, D2, 12, "Izlaz", "S1", "Stanica", "D1", "OTP-1", "Otpremnica")
        d = {}
        d.update(izmeni("tblOtpremnica", {"OTP-1": {"KolAmbalaze": 12}}))
        d.update(izmeni("tblZbirna", {"ZBR-1": {"UkupnoAmbalaze": 12}}))
        c = self.sagradi(amb, d)
        ost = {s["OtpremnicaID"]: s for s in c.tabele["tblOtpremnicaStavke"]}
        self.assertEqual(ost["OTP-1"]["KolAmbalaze"], Decimal(10))
        self.assertEqual(c.tabele["tblZbirnaStavke"][0]["KolAmbalaze"], Decimal(10))
        self.assertEqual(self.knjiga(c, "OTP-1"), [("AMBALAZA_UZ_ROBU", "S1", "D1", Decimal(10))])
        self.assertIn(("GAJBE_OTPREMNICE_PO_BLOKOVIMA", "OTP-1"), self.nalazi(c, "INFO"))
        s1 = self.saldo(c, "Stanica", "S1")
        self.assertEqual((s1["korekcija"], s1["razlika"], s1["najmanje_pune"]),
                         (Decimal(2), Decimal(0), Decimal(0)))
        self.assertEqual([n for n in self.nalazi(c) if n[0].startswith("AMB_")], [])

    def test_neizdata_otpremnica_ne_knjizi(self):
        """Otpremnica bez otkupa ostaje neizdata (MIG-ODL-06), a main knjizi gajbe tek
        pri izdavanju (OtpIzdaj): njene stare noge izlaze iz knjige, ali ne nestaju."""
        amb = normalno() + [noga(13, D2, 5, "Izlaz", "S1", "Stanica", "D1", "OTP-5", "Otpremnica")]
        d = izmeni("tblOtpremnica", novi=[dict(OtpremnicaID="OTP-5", Datum=D2, StanicaID="S1", VozacID="D1",
                                                BrojOtpremnice="205", Klasa="I", Kolicina=40, KolAmbalaze=5)])
        c = self.sagradi(amb, d)
        self.assertEqual(self.knjiga(c, "OTP-5"), [])
        self.assertIn(("NEIZDATA_OTPREMNICA_NE_KNJIZI", "OTP-5"), self.nalazi(c, "INFO"))
        noga13 = next(n for n in c.amb.noge if n.amb_id == "AMB-00013")
        self.assertEqual(noga13.ishod, ambalaza.ISHOD_NE_KNJIZI)
        self.assertFalse(any(ambalaza.proveri_pokrivenost(c.amb).values()))
        self.assertEqual(self.saldo(c, "Stanica", "S1")["razlika"], 0)

    def test_stornirana_po_odluci_ne_knjizi(self):
        """MIG-ODL-19 sa aktivnom starom nogom (kao OTP-00666, 41 gajba): stornirana
        otpremnica nema neto knjigu, pa noga izlazi iz knjige, sa svojim razlogom."""
        amb = normalno() + [noga(13, D2, 5, "Izlaz", "S1", "Stanica", "D1", "OTP-5", "Otpremnica")]
        d = izmeni("tblOtpremnica", novi=[dict(OtpremnicaID="OTP-5", Datum=D2, StanicaID="S1", VozacID="D1",
                                                BrojOtpremnice="205", Klasa="I", Kolicina=40, KolAmbalaze=5)])
        iz = [{"OtpremnicaID": "OTP-5", "radnja": "storniraj", "razlog": "test",
               "provera": {"BrojOtpremnice": "205", "Datum": D2.isoformat(), "StanicaID": "S1"}}]
        c = self.sagradi(amb, d, izuzeci_dok=iz)
        self.assertEqual(self.knjiga(c, "OTP-5"), [])
        self.assertIn(("STORNIRANA_OTPREMNICA_NE_KNJIZI", "OTP-5"), self.nalazi(c, "INFO"))
        self.assertFalse(any(ambalaza.proveri_pokrivenost(c.amb).values()))
        self.assertEqual([n for n in self.nalazi(c, "BLOKIRA")], [])

    def test_nabavka_pokriva_i_prazne(self):
        """AMB-INV-16: S2 primi 10 punih i izda 4 prazne -- ukupno +6, prazne -4.
        Nabavka 4 prvog dana (pre #448 ne bi je bilo, pa bi kapija odbila izdavanje)."""
        amb = normalno() + [
            noga(13, D1, 10, "Izlaz", "K2", "Kooperant", "", "OTK-7", "Otkup"),
            noga(14, D1, 10, "Ulaz", "S2", "Stanica", "", "OTK-7", "Otkup"),
            noga(15, D1, 4, "Ulaz", "K2", "Kooperant", "", "OTK-7", "OM-Izlaz-Koop"),
            noga(16, D1, 4, "Izlaz", "S2", "Stanica", "", "OTK-7", "OM-Izlaz-Koop")]
        d = izmeni("tblOtkup", novi=[dict(OtkupID="OTK-7", Datum=D1, KooperantID="K2", StanicaID="S2",
                                          BrojDokumenta="107", Klasa="I", Kolicina=100, Cena=80,
                                          KolAmbalaze=10, KolAmbIzdata=4, TipAmbalaze="12/1")])
        c = self.sagradi(amb, d)
        s2 = self.saldo(c, "Stanica", "S2")
        self.assertEqual((s2["nabavka"], s2["najmanje_prazne"], s2["najmanji"], s2["pune"], s2["prazne"]),
                         (Decimal(4), Decimal(0), Decimal(10), Decimal(10), Decimal(0)))
        self.assertNotIn("AMB_INV16_PRAZNE_ISPOD_NULE", [n[0] for n in self.nalazi(c, "BLOKIRA")])

    def test_pune_vozaca_ispod_nule_ide_na_odluku(self):
        """Prijemnica 12 punih, a vozac je otpremnicom primio 10: nabavka pokriva ukupno,
        ali pune ostaju -2 -- knjiga to ne ume da popravi, operater odlucuje (KI-065)."""
        amb = normalno()
        amb[7] = noga(8, D3, 12, "Ulaz", "C1", "Kupac", "D1", "PRJ-1", "Prijemnica")
        c = self.sagradi(amb, izmeni("tblPrijemnica", {"PRJ-1": {"KolAmbalaze": 12}}))
        d1 = self.saldo(c, "Vozac", "D1")
        self.assertEqual((d1["najmanje_pune"], d1["najmanji"], d1["najmanje_prazne"]),
                         (Decimal(-2), Decimal(0), Decimal(0)))
        self.assertIn("AMB_PUNE_ISPOD_NULE", [n[0] for n in self.nalazi(c, "ODLUKA")])

    def test_dokument_otkupa_prati_staru_knjigu(self):
        """Otkup kaze 5 izdatih, stara knjiga 4: main knjizi tacno polje dokumenta,
        a istorija je knjiga (MIG-ODL-01) -- dokument dobija 4, uz upozorenje."""
        c = self.sagradi(dodatno=izmeni("tblOtkup", {"OTK-1": {"KolAmbIzdata": 5}}))
        otk = {z["OtkupID"]: z for z in c.tabele["tblOtkup"]}
        self.assertEqual(otk["OTK-1"]["KolAmbIzdata"], Decimal(4))
        self.assertIn(("GAJBE_DOKUMENTA_PO_KNJIZI", "OTK-1"), self.nalazi(c, "UPOZORENJE"))
        self.assertEqual(sorted(self.knjiga(c, "OTK-1"))[1], ("IZDATA_PRAZNA", "S1", "K1", Decimal(4)))

    def test_dijagnostika_preostalo_kg_i_kandidati(self):
        """OTK-3 (40 kg) visi pod OTP-1 (100 kg): OTP-1 ostaje izdata, a spisak za
        operatera kaze preostalo -40; OTP-5 (40 kg, ista stanica, bez bloka) ceka
        blokove i dobija OTP-1 kao kandidata (blok moguce pod pogresnom otpremnicom)."""
        d = izmeni("tblOtkup", {"OTK-3": {"OtpremnicaID": "OTP-1"}})
        d.update(izmeni("tblOtpremnica", novi=[dict(OtpremnicaID="OTP-5", Datum=D2, StanicaID="S1",
                                                      VozacID="D1", BrojOtpremnice="205", Klasa="I",
                                                      Kolicina=40, KolAmbalaze=3)]))
        c = self.sagradi(dodatno=d)
        otp = {z["OtpremnicaID"]: z["IzdatoStatus"] for z in c.tabele["tblOtpremnica"]}
        self.assertEqual((otp["OTP-1"], otp["OTP-5"]), (cilj.IZDATO, ""))
        up = {n["id"]: n["opis"] for n in c.nalazi if n["pravilo"] == "OTPREMNICA_PREOSTALO_KG"}
        self.assertIn("ocekivano I 100 kg, povezano I 140 kg, preostalo I -40", up["OTP-1"])
        self.assertIn(("OTPREMNICA_PREOSTALO_KG", "OTP-1"), self.nalazi(c, "UPOZORENJE"))
        lanac = next(n["opis"] for n in c.nalazi
                     if n["pravilo"] == "NEIZDATA_OTPREMNICA_LANAC" and n["id"] == "OTP-5")
        self.assertIn("ocekivano I 40 kg", lanac)
        self.assertIn("OTP-1 201 (+40 kg)", lanac)

    def test_otpremnica_sa_delom_blokova_ostaje_izdata(self):
        """Odluka operatera 03.11.2026: neslaganje kg otkupa i otpremnice resava operater
        sa otkupcem -- alat ne pravi pravilo. Otpremnica 130 kg sa blokom od 100 kg ostaje
        izdata, kg dokumenta se ne dira, gajbe i knjiga idu po bloku (MIG-ODL-14)."""
        c = self.sagradi(dodatno=izmeni("tblOtpremnica", {"OTP-1": {"Kolicina": 130, "KolAmbalaze": 13}}))
        otp = {z["OtpremnicaID"]: z for z in c.tabele["tblOtpremnica"]}
        self.assertEqual(otp["OTP-1"]["IzdatoStatus"], cilj.IZDATO)
        ost = {s["OtpremnicaID"]: s for s in c.tabele["tblOtpremnicaStavke"]}
        self.assertEqual((ost["OTP-1"]["Kolicina"], ost["OTP-1"]["KolAmbalaze"]), (Decimal(130), Decimal(10)))
        self.assertEqual(self.knjiga(c, "OTP-1"), [("AMBALAZA_UZ_ROBU", "S1", "D1", Decimal(10))])
        self.assertIn(("OTPREMNICA_PREOSTALO_KG", "OTP-1"), self.nalazi(c, "UPOZORENJE"))
        self.assertEqual([n for n in self.nalazi(c, "ODLUKA") if n[1] == "OTP-1"], [])

    def test_bilans_neuskladjeno_nula_i_vracanje_ima_izlaz(self):
        """AMB-BIL-02 / AMB-INV-15 nad celom novom sveskom (kooperant sa dugom)."""
        c = self.sagradi()
        vrste = Counter(r["VrstaKretanja"] for r in c.tabele["tblAmbalaza"])
        self.assertEqual(vrste["VRACANJE_TUDJE_AMBALAZE"], vrste["IZLAZ_TUDJE_AMBALAZE"])
        self.assertGreater(vrste["IZLAZ_TUDJE_AMBALAZE"], 0)
        self.assertEqual([x for x in ambalaza.proveri_invarijante(c.amb)], [])


class Upis(SaSveskom):
    """Ciste funkcije upisa (bez Excela): pretvaranje, redosled kolona, poredjenje."""

    def test_com_vrednost_cuva_smisao(self):
        from datetime import datetime, timezone
        self.assertIsNone(upis.com_vrednost(""))
        self.assertEqual(upis.com_vrednost("101"), "'101")              # ostaje tekst
        self.assertEqual(upis.com_vrednost("2026-06-20 08:00:00"), "'2026-06-20 08:00:00")
        self.assertEqual(upis.com_vrednost("=1+1"), "'=1+1")             # ne postaje formula
        self.assertEqual(upis.com_vrednost("#N/A"), "'#N/A")             # ne postaje greska
        self.assertEqual(upis.com_vrednost("TRUE"), "'TRUE")             # ne postaje logicka
        self.assertEqual(upis.com_vrednost("Pera"), "Pera")
        self.assertEqual(upis.com_vrednost(Decimal("100")), 100)
        self.assertEqual(upis.com_vrednost(Decimal("1.5")), 1.5)
        # datum sa oznakom UTC: pywin32 tada ne pomera vreme (pad 28.10.2026: -2 h)
        d = upis.com_vrednost(D1)
        self.assertEqual((d.replace(tzinfo=None), d.tzinfo), (datetime(2026, 6, 20), timezone.utc))
        d = upis.com_vrednost(datetime(2026, 6, 20, 8, 30))
        self.assertEqual((d.replace(tzinfo=None), d.tzinfo), (datetime(2026, 6, 20, 8, 30), timezone.utc))

    def test_niz_po_redosledu_kolona_u_svesci(self):
        redovi = [{"A": "x", "B": Decimal(2)}]
        self.assertEqual(upis.niz_za_tabelu(redovi, ["B", "A", "Runtime"]), [(2, "x", None)])
        with self.assertRaisesRegex(ValueError, "kolone sa podacima"):
            upis.niz_za_tabelu(redovi, ["A"])
        # tekstualna kolona: bez apostrofa (format "@" se postavlja pre upisa)
        self.assertEqual(upis.niz_za_tabelu([{"BrojDokumenta": "101", "Napomena": "101"}],
                                            ["BrojDokumenta", "Napomena"]), [("101", "'101")])

    def test_upis_tabele_gradi_opseg_iz_ugaonih_celija(self):
        # Lazni Excel: Range postoji samo kao (r1, c1, r2, c2); nema Resize/Offset,
        # pa bi povratak na Range.Resize (pad 28.10.2026) ovde pao AttributeError-om.
        from datetime import datetime, timedelta, timezone

        class Opseg:
            def __init__(self, r1, c1, r2, c2):
                self.r1, self.c1, self.r2, self.c2 = r1, c1, r2, c2
                self.Row, self.Column = r1, c1
                self.Columns = type("K", (), {"Count": c2 - c1 + 1})()
                self.Address = "R%dC%d" % (r1, c1)

        class Celije:
            def __init__(self, lst, o):
                self.__dict__.update(lst=lst, opseg=o)

            def __setattr__(self, ime, v):
                o, lst = self.opseg, self.lst
                a = (o.r1, o.c1, o.r2, o.c2)
                if ime == "NumberFormat":
                    lst.formati.append((a, v))
                    return
                lst.upisi.append((a, v))
                for i, red in enumerate(v):
                    for j, x in enumerate(red):
                        if isinstance(x, datetime):
                            # kao pywin32: astimezone(UTC), naivno vreme = lokalno (CEST)
                            x = (x - lst.lokalno if x.tzinfo is None else x.astimezone(timezone.utc))
                            x = x.replace(tzinfo=None)
                        elif isinstance(x, str) and x.startswith("'"):
                            x = x[1:]
                        lst.celije[(o.r1 + i, o.c1 + j)] = x

            def __getattr__(self, ime):
                if ime != "Value":
                    raise AttributeError(ime)
                o = self.opseg
                # kao pywin32: datum se vraca kao upisano vreme sa oznakom UTC
                v = tuple(tuple((lambda x: x.replace(tzinfo=timezone.utc) if isinstance(x, datetime) else x)(
                    self.lst.celije.get((r, c))) for c in range(o.c1, o.c2 + 1)) for r in range(o.r1, o.r2 + 1))
                return v[0][0] if len(v) == 1 and len(v[0]) == 1 else v

        class List:
            def __init__(self):
                self.upisi, self.formati, self.celije = [], [], {}
                self.lokalno = timedelta(hours=2)

            def Cells(self, r, c):
                return (r, c)

            def Range(self, a, b):
                return Celije(self, Opseg(a[0], a[1], b[0], b[1]))

        class Tabela:
            # Kao Excel: tabela samo sa zaglavljem vec nosi jedan prazan red za unos,
            # a ListRows.Count je 0 dok se u njega ne upise (pad tblKese, 1 red).
            def __init__(self, r0, c0, m):
                self.HeaderRowRange = Opseg(r0, c0, r0, c0 + m - 1)
                self.redova = 1
                self.prosirenje = None

            @property
            def Range(self):
                return type("R", (), {"Worksheet": ws,
                                      "Rows": type("K", (), {"Count": self.redova + 1})()})()

            @property
            def ListRows(self):
                prazna = self.redova == 1 and not ws.upisi
                return type("L", (), {"Count": 0 if prazna else self.redova})()

            def Resize(self, cilj):
                o = cilj.__dict__["opseg"]
                self.prosirenje = (o.r1, o.c1, o.r2, o.c2)
                self.redova = o.r2 - o.r1

        kolone = ["Naziv", "Kolicina", "BrojDokumenta", "Datum"]
        ws = List()
        lo = Tabela(3, 2, 4)                                  # zaglavlje u B3:E3
        niz = upis.niz_za_tabelu([{"Naziv": "a%d" % i, "Kolicina": Decimal(i), "BrojDokumenta": "101",
                                   "Datum": D1} for i in range(5)], kolone)
        stari = upis.DELOVI
        upis.DELOVI = 2
        try:
            upis._upisi_tabelu(lo, niz, kolone)
        finally:
            upis.DELOVI = stari
        self.assertEqual(lo.prosirenje, (3, 2, 8, 5))        # zaglavlje + 5 redova
        self.assertEqual(ws.formati, [((4, 4, 8, 4), "@")])  # samo tekst kolona, ceo stub
        self.assertEqual([a for a, _ in ws.upisi], [(4, 2, 5, 5), (6, 2, 7, 5), (8, 2, 8, 5)])
        self.assertEqual([r for _, v in ws.upisi for r in v], niz)
        self.assertEqual(ws.celije[(8, 5)], datetime(2026, 6, 20))   # datum bez pomaka

        ws = List()
        lo = Tabela(1, 1, 4)
        upis._upisi_tabelu(lo, niz[:1], kolone)               # jedan red: Resize je no-op
        self.assertEqual([a for a, _ in ws.upisi], [(2, 1, 2, 4)])

        ws = List()
        lo = Tabela(1, 1, 4)
        lo.Resize = lambda cilj: None                         # Excel ne prosiri tabelu
        with self.assertRaisesRegex(RuntimeError, "zauzima 2 redova sa zaglavljem, ocekivano 6"):
            upis._upisi_tabelu(lo, niz, kolone)

        # Naivan datum (pre ispravke): Excel ga pomeri, uzorak zaustavi upis pre cuvanja.
        ws = List()
        lo = Tabela(1, 1, 4)
        naivno = [r[:3] + (datetime(2026, 6, 8),) for r in niz]
        with self.assertRaisesRegex(RuntimeError, "povratno citanje, red 1, Datum"):
            upis._upisi_tabelu(lo, naivno, kolone)

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
        # Masovna greska u jednoj tabeli ne sakriva drugu, ni drugu kolonu istog reda.
        from datetime import timedelta
        for red in c.tabele["tblAmbalaza"]:
            red["Datum"] = red["Datum"] + timedelta(days=1)
        c.tabele["tblAmbalaza"][0]["Kolicina"] = Decimal(9999)
        stari = upis.PRIMERA_PO_TABELI
        upis.PRIMERA_PO_TABELI = 2
        try:
            r = upis.uporedi(c, paket)
        finally:
            upis.PRIMERA_PO_TABELI = stari
        self.assertEqual(len([x for x in r if x.startswith("tblAmbalaza red")]), 2)
        self.assertIn("tblAmbalaza: ukupno 12 razlika -- po koloni: Datum 11, Kolicina 1", r)
        self.assertIn("tblOtkupStavke red 1, Kolicina", r[-1])

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


def druga_prijemnica(vozac_prj="D1", vozac_zbr2="D1", broj_zbr2="301-2"):
    """ANOMALIJA (greska unosa u staroj formi): PRJ-3 je nasledila broj zbirne
    301 od prethodne prijemnice, a roba je sa zbirne ZBR-2 (50 kg, 5 gajbi),
    koja je ostala bez prijemnice."""
    d = {}
    d.update(izmeni("tblZbirna", novi=[dict(ZbirnaID="ZBR-2", Datum=D2, VozacID=vozac_zbr2, KupacID="C1",
                                            BrojZbirne=broj_zbr2, Klasa="I", UkupnoKolicina=50,
                                            UkupnoAmbalaze=5, GeneracijaID="G-ZBR-2")]))
    d.update(izmeni("tblOtpremnica", novi=[dict(OtpremnicaID="OTP-2", Datum=D2, StanicaID="S1",
                                                VozacID=vozac_zbr2, BrojOtpremnice="202", Klasa="I",
                                                Kolicina=50, KolAmbalaze=5, GeneracijaID="G-OTP-2",
                                                BrojZbirne=broj_zbr2)]))
    d.update(izmeni("tblPrijemnica", novi=[dict(PrijemnicaID="PRJ-3", Datum=D3, KupacID="C1",
                                                VozacID=vozac_prj, BrojPrijemnice="402", BrojZbirne="301",
                                                Klasa="I", Kolicina=Decimal("49.5"), Cena=60, KolAmbalaze=5,
                                                GeneracijaID="G-PRJ-3")]))
    return d


class LanacStareSveske(SaSveskom):
    """lanac.py: prijemnica na pogresnoj zbirnoj i storno koji nije presao lanac."""

    def lanac(self, dodatno=None):
        put = os.path.join(self.dir.name, "stara.xlsx")
        snimi_svesku(put, osnovne_tabele(normalno(), dodatno))
        self.tabele = procitaj_svesku(put)
        return lanac.Lanac(self.tabele).proveri()

    def pravila(self, lc, tezina=None):
        return sorted((n["pravilo"], n["dok"].id) for n in lc.nalazi if tezina in (None, n["tezina"]))

    def test_normalna_sveska_nema_nalaza(self):
        """NORMALNO: dve klase prijemnice (isti GeneracijaID i broj) su JEDNA prijemnica."""
        lc = self.lanac()
        self.assertEqual(self.pravila(lc), [])
        self.assertEqual(lc.prijemnice, [])
        self.assertEqual([d.ids for d in lc.prj], [["PRJ-1", "PRJ-2"]])

    def test_ista_stanica_kasniji_unos_dobija_predlog(self):
        """Druga tura iste stanice: vozac se slaze, pa migracija prijemnicu tiho
        veze na pogresnu zbirnu -- trag je samo par 2 prijemnice / 0 prijemnica."""
        lc = self.lanac(druga_prijemnica())
        self.assertIn(("ZBR_VISE_PRIJEMNICA", "ZBR-1"), self.pravila(lc, lanac.ALARM))
        self.assertIn(("ZBR_BEZ_PRIJEMNICE", "ZBR-2"), self.pravila(lc, lanac.ALARM))
        self.assertNotIn("PRJ_VOZAC_NIJE_VOZAC_ZBIRNE", [p for p, _ in self.pravila(lc)])
        [r] = lc.prijemnice
        self.assertEqual((r["prj"].id, r["ostaje"].id, r["predlog"].id, r["pouzdanost"]),
                         ("PRJ-3", "PRJ-1", "ZBR-2", "VISOKA"))
        self.assertIn("kasniji unos (PRJ-3 posle PRJ-1)", r["signali"])
        self.assertIn("gajbe se slazu", r["signali"])
        self.assertEqual([b["predlozena_za"] for b in lc.bez_prijemnice], [["PRJ-3"]])
        c = cilj.izgradi(self.tabele, dokumenti.ucitaj_kanon())
        rod = {p["PrijemnicaID"]: p["ZbirnaRoditeljID"] for p in c.tabele["tblPrijemnica"]}
        self.assertEqual(rod["PRJ-3"], "ZBR-1")     # migracija ovde ne upozorava

    def test_ista_generacija_drugi_broj_su_dva_dokumenta(self):
        """ANOMALIJA: pogresna prijemnica nosi GeneracijaID prave -- broj je razlikuje."""
        d = druga_prijemnica()
        d["tblPrijemnica"][1][-1][d["tblPrijemnica"][0].index("GeneracijaID")] = "G-PRJ-1"
        lc = self.lanac(d)
        self.assertEqual(sorted(x.id for x in lc.prj), ["PRJ-1", "PRJ-3"])
        self.assertEqual([r["prj"].id for r in lc.prijemnice], ["PRJ-3"])

    def test_druga_stanica_vozac_nije_vozac_zbirne(self):
        lc = self.lanac(druga_prijemnica(vozac_prj="D2", vozac_zbr2="D2", broj_zbr2="S2/210626"))
        self.assertIn(("PRJ_VOZAC_NIJE_VOZAC_ZBIRNE", "PRJ-3"), self.pravila(lc, lanac.ALARM))
        [r] = lc.prijemnice
        self.assertEqual((r["predlog"].id, r["pouzdanost"]), ("ZBR-2", "VISOKA"))
        self.assertIn("vozac prijemnice D2 nije vozac zbirne D1", r["signali"])

    def test_predlog_prvo_istog_vozaca(self):
        """Zbirna drugog vozaca sa tacnim kg ne pobedjuje zbirnu vozaca prijemnice."""
        d = druga_prijemnica()
        d["tblZbirna"][1].append([dict(ZbirnaID="ZBR-3", Datum=D2, VozacID="D2", KupacID="C1",
                                       BrojZbirne="S2/210626", Klasa="I", UkupnoKolicina=Decimal("49.5"),
                                       UkupnoAmbalaze=5, GeneracijaID="G-ZBR-3").get(k, "")
                                  for k in d["tblZbirna"][0]])
        [r] = self.lanac(d).prijemnice
        self.assertEqual(r["predlog"].id, "ZBR-2")
        self.assertTrue(r["ostali"][0].startswith("S2/210626 D2"), r["ostali"])

    def test_isti_broj_stornirana_i_aktivna_zbirna(self):
        """NORMALNO (stari tok ispravke): storno zbirne pa nov unos pod ISTIM brojem.
        Prijemnica pripada aktivnoj zbirnoj, ne storniranoj."""
        lc = self.lanac(izmeni("tblZbirna", novi=[dict(ZbirnaID="ZBR-0", Datum=D1, VozacID="D1", KupacID="C1",
                                                       BrojZbirne="301", Klasa="I", UkupnoKolicina=100,
                                                       UkupnoAmbalaze=10, Stornirano="Da",
                                                       GeneracijaID="G-ZBR-0")]))
        self.assertEqual(self.pravila(lc), [])
        self.assertEqual([lc.zbr_od_prj[p].id for p in lc.prj], ["ZBR-1"])

    def test_nema_predloga_van_prozora_dana(self):
        d = druga_prijemnica()
        d.update(izmeni("tblZbirna", novi=[dict(ZbirnaID="ZBR-2", Datum=date(2026, 6, 1), VozacID="D1",
                                                KupacID="C1", BrojZbirne="301-2", Klasa="I", UkupnoKolicina=50,
                                                UkupnoAmbalaze=5, GeneracijaID="G-ZBR-2")]))
        [r] = self.lanac(d).prijemnice
        self.assertEqual((r["predlog"], r["pouzdanost"]), (None, "NEMA KANDIDATA"))

    def test_prefiks_broja_i_broj_vozaca(self):
        self.assertEqual([lanac.prefiks_broja(b) for b in ("S3/030726", "5/250626-2", "301")], [3, 5, None])
        self.assertEqual([lanac.broj_vozaca(v) for v in ("ST-00002", "D1", "")], [2, None, None])

    def test_storno_otpremnice_nije_presao_lanac(self):
        lc = self.lanac(izmeni("tblOtpremnica", {"OTP-1": {"Stornirano": "Da"}}))
        self.assertEqual(self.pravila(lc, lanac.ALARM),
                         [("OTK_U_STORNIRANOJ_OTPREMNICI", "OTK-1"), ("ZBR_SVE_OTPREMNICE_STORNIRANE", "ZBR-1")])
        z = next(n for n in lc.nalazi if n["pravilo"] == "ZBR_SVE_OTPREMNICE_STORNIRANE")
        self.assertIn("PAZNJA: aktivna prijemnica PRJ-1", z["opis"])

    def test_storniran_otkup_u_aktivnoj_otpremnici(self):
        lc = self.lanac(izmeni("tblOtkup", novi=[dict(OtkupID="OTK-4", Datum=D1, KooperantID="K2", StanicaID="S1",
                                                      BrojDokumenta="104", Klasa="I", Kolicina=36, Cena=80,
                                                      KolAmbalaze=0, TipAmbalaze="12/1", Stornirano="Da",
                                                      OtpremnicaID="OTP-1")]))
        self.assertEqual(self.pravila(lc), [("OTP_SA_STORNIRANIM_OTKUPOM", "OTP-1"),
                                            ("ZBR_SA_STORNIRANIM_OTKUPOM", "ZBR-1")])
        n = next(n for n in lc.nalazi if n["pravilo"] == "OTP_SA_STORNIRANIM_OTKUPOM")
        self.assertIn("stornirani blokovi OTK-4 (36 kg); aktivnih 1 (100 kg)", n["opis"])

    def test_storno_zbirne_nije_presao_lanac(self):
        lc = self.lanac(izmeni("tblZbirna", {"ZBR-1": {"Stornirano": "Da"}}))
        self.assertEqual(self.pravila(lc, lanac.ALARM),
                         [("OTK_U_STORNIRANOJ_ZBIRNOJ", "OTK-1"), ("OTP_U_STORNIRANOJ_ZBIRNOJ", "OTP-1"),
                          ("PRJ_ZBIRNA_STORNIRANA", "PRJ-1")])

    def test_storno_prijemnice_nije_presao_lanac(self):
        lc = self.lanac(izmeni("tblPrijemnica", {"PRJ-1": {"Stornirano": "Da"}, "PRJ-2": {"Stornirano": "Da"}}))
        self.assertEqual(self.pravila(lc, lanac.ALARM),
                         [("OTK_PRIJEMNICA_STORNIRANA", "OTK-1"), ("OTP_PRIJEMNICA_STORNIRANA", "OTP-1"),
                          ("ZBR_SAMO_STORNIRANE_PRIJEMNICE", "ZBR-1")])

    def test_lanac_komanda_pravi_izvestaj_i_ne_menja_izvor(self):
        put = os.path.join(self.dir.name, "stara.xlsx")
        snimi_svesku(put, osnovne_tabele(normalno(), druga_prijemnica()))
        pre = sha256_fajla(put)
        izlaz = os.path.join(self.dir.name, "lanac")
        self.assertEqual(tiho(["lanac", "--izvor", put, "--izlaz", izlaz]), 0)
        self.assertEqual(sha256_fajla(put), pre)
        t = procitaj_svesku(os.path.join(izlaz, "lanac.xlsx"))
        [r] = tabela(t, "Prijemnice").redovi
        self.assertEqual((polje(r, "PrijemnicaID"), polje(r, "SadaBrojZbirne"), polje(r, "PredlogBrojZbirne")),
                         ("PRJ-3", "301", "301-2"))
        self.assertEqual([polje(z, "BrojZbirne") for z in tabela(t, "ZbirneBezPrijemnice").redovi], ["301-2"])
        self.assertIsNotNone(tabela(t, "Nalazi"))
        with open(os.path.join(izlaz, "lanac_sazetak.txt"), encoding="utf-8") as f:
            self.assertIn("PRJ-3 402: sada 301 -> predlog 301-2  [VISOKA]", f.read())

    def test_lanac_odbija_izlaz_u_repou(self):
        put = os.path.join(self.dir.name, "stara.xlsx")
        snimi_svesku(put, osnovne_tabele(normalno()))
        izlaz = os.path.join(migracija.KOREN_REPOA, "lanac_test_ne_sme")
        self.assertEqual(tiho(["lanac", "--izvor", put, "--izlaz", izlaz]), 2)
        self.assertFalse(os.path.exists(izlaz))


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
        self.assertEqual(len(tabela(t, "Kretanja").redovi), 11)       # 7 + pokrice + vracanje + izlaz + nabavka
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


# ---------------------------------------------------------------------------
# MIG-ODL-20/21: prijemnica i zbirna po odluci, kultura po sorti
# ---------------------------------------------------------------------------

KUL_KOL = ["KulturaID", "VrstaVoca", "SortaVoca"]
KULTURE = (KUL_KOL, [["KUL-1", "Malina", "Willamette"], ["KUL-2", "Malina", "Willamette teren"],
                     ["KUL-3", "Malina", "Fertodi"]])


def otkup_sa_kulturom(vrednosti):
    """tblOtkup osnovne sveske uz VrstaVoca/SortaVoca/KulturaID ({OtkupID: (vrsta, sorta, kultura)})."""
    kol, redovi = osnovne_tabele(normalno())["tblOtkup"]
    kol = kol + ["VrstaVoca", "SortaVoca", "KulturaID"]
    out = []
    for r in redovi:
        out.append(list(r) + list(vrednosti.get(r[0], ("", "", ""))))
    return {"tblOtkup": (kol, out)}


class MigOdl21KulturaPoSorti(SaNovomSveskom):
    def kultura(self, c, oid):
        return next(z["KulturaID"] for z in c.tabele["tblOtkup"] if z["OtkupID"] == oid)

    def test_kultura_po_vrsti_i_sorti(self):
        """Stari pisac: prva kultura vrste. Nov: tacno jedan red iste vrste i sorte."""
        d = otkup_sa_kulturom({"OTK-1": ("Malina", "Willamette teren", "KUL-1"),
                               "OTK-2": ("Malina", "Willamette", "KUL-1"),
                               "OTK-3": ("malina", " Fertodi ", "Malina-Fertodi")})
        d["tblKulture"] = KULTURE
        c = self.sagradi(dodatno=d)
        self.assertEqual([self.kultura(c, o) for o in ("OTK-1", "OTK-2", "OTK-3")], ["KUL-2", "KUL-1", "KUL-3"])
        info = [n for n in c.nalazi if n["pravilo"] == "KULTURA_PO_SORTI"]
        self.assertEqual(len(info), 2, info)
        self.assertNotIn("KULTURA_NE_RAZRESAVA", [n["pravilo"] for n in c.nalazi])
        # tekst lista se ne menja
        z1 = next(z for z in c.tabele["tblOtkup"] if z["OtkupID"] == "OTK-1")
        self.assertEqual((z1["VrstaVoca"], z1["SortaVoca"]), ("Malina", "Willamette teren"))

    def test_nerazresivo_ostaje_i_odluka(self):
        d = otkup_sa_kulturom({"OTK-1": ("Malina", "Meeker", "KUL-1"), "OTK-2": ("Malina", "Willamette", "KUL-1"),
                               "OTK-3": ("Malina", "Fertodi", "KUL-3")})
        d["tblKulture"] = KULTURE
        c = self.sagradi(dodatno=d)
        self.assertEqual(self.kultura(c, "OTK-1"), "KUL-1")
        odl = [n for n in c.nalazi if n["pravilo"] == "KULTURA_NE_RAZRESAVA"]
        self.assertEqual(len(odl), 1, odl)
        self.assertEqual(odl[0]["tezina"], "ODLUKA")
        self.assertIn("Meeker", odl[0]["opis"])

    def test_dva_reda_iste_sorte_ne_razresava(self):
        """Nula i vise od jedan pogodak su ista greska -- kao resolver main-a."""
        d = otkup_sa_kulturom({"OTK-1": ("Malina", "Willamette", "KUL-1")})
        d["tblKulture"] = (KUL_KOL, KULTURE[1] + [["KUL-9", "Malina", "Willamette"]])
        c = self.sagradi(dodatno=d)
        self.assertEqual(self.kultura(c, "OTK-1"), "KUL-1")
        self.assertIn("KULTURA_NE_RAZRESAVA", [n["pravilo"] for n in c.nalazi])

    def test_bez_sifarnika_samo_info(self):
        c = self.sagradi()
        prav = [(n["tezina"], n["pravilo"]) for n in c.nalazi if "KULTUR" in n["pravilo"]]
        self.assertEqual(prav, [("INFO", "KULTURA_NIJE_PROVERENA")])


class MigOdl20PoOdluci(SaNovomSveskom):
    """Prijemnica PRJ-3 je nasledila zbirnu 301 od prethodnog unosa; prava je 302."""

    def dodatno(self, prijemnica_302=False, otp9_zbirna=""):
        z = osnovne_tabele(normalno())["tblZbirna"]
        zbr = (z[0], z[1] + [["ZBR-2", D3, "D1", "C1", "302", "I", 50, 5, "", "G-ZBR-2"],
                             ["ZBR-3", D2, "D1", "C1", "303", "I", 30, 3, "", "G-ZBR-3"]])
        d = {"tblZbirna": zbr}
        d.update(izmeni("tblPrijemnica", novi=[dict(
            PrijemnicaID="PRJ-3", Datum=D3, KupacID="C1", VozacID="D1", BrojPrijemnice="402",
            BrojZbirne="302" if prijemnica_302 else "301", Klasa="I", Kolicina=49, Cena=60, KolAmbalaze=5)]))
        d.update(izmeni("tblOtpremnica", {"OTP-9": {"BrojZbirne": otp9_zbirna}}))
        return d

    PREVEZI = {"PrijemnicaID": "PRJ-3", "radnja": "prevezi_prijemnicu", "BrojZbirne": "302",
               "provera": {"BrojZbirne": "301"}, "razlog": "test MIG-ODL-20"}

    def prj(self, c, pid):
        return next(p for p in c.tabele["tblPrijemnica"] if p["PrijemnicaID"] == pid)

    def test_prevezi_prijemnicu(self):
        c = self.sagradi(dodatno=self.dodatno(), izuzeci_dok=[self.PREVEZI])
        p = self.prj(c, "PRJ-3")
        self.assertEqual((p["BrojZbirne"], p["ZbirnaRoditeljID"]), ("302", "ZBR-2"))
        self.assertIn(("PRIJEMNICA_PREVEZANA_PO_ODLUCI", "PRJ-3"), self.nalazi(c, "INFO"))

    def test_bez_odluke_ostaje_na_pogresnoj(self):
        """Kontrola: bez odluke ista stanica prolazi tiho, na pogresnoj zbirnoj."""
        c = self.sagradi(dodatno=self.dodatno())
        self.assertEqual(self.prj(c, "PRJ-3")["ZbirnaRoditeljID"], "ZBR-1")

    def test_prevezi_potpis_se_ne_slaze(self):
        iz = dict(self.PREVEZI, provera={"BrojZbirne": "999"})
        c = self.sagradi(dodatno=self.dodatno(), izuzeci_dok=[iz])
        self.assertEqual(self.prj(c, "PRJ-3")["BrojZbirne"], "301")
        self.assertIn(("IZUZETAK_DOKUMENTA_POTPIS_SE_NE_SLAZE", "PRJ-3"), self.nalazi(c, "ODLUKA"))

    def test_prevezi_na_zbirnu_koje_nema(self):
        iz = dict(self.PREVEZI, BrojZbirne="777")
        c = self.sagradi(dodatno=self.dodatno(), izuzeci_dok=[iz])
        self.assertEqual(self.prj(c, "PRJ-3")["BrojZbirne"], "301")
        self.assertIn(("PREVEZIVANJE_NA_ZBIRNU_KOJE_NEMA", "PRJ-3"), self.nalazi(c, "ODLUKA"))

    STORNO_303 = {"tabela": "tblZbirna", "BrojZbirne": "303", "radnja": "storniraj",
                  "provera": {"BrojZbirne": "303"}, "razlog": "test: otpremnica stornirana"}

    def zbr(self, c, zid):
        return next(z for z in c.tabele["tblZbirna"] if z["ZbirnaID"] == zid)

    def test_storniraj_zbirnu_cija_je_otpremnica_stornirana(self):
        c = self.sagradi(dodatno=self.dodatno(otp9_zbirna="303"), izuzeci_dok=[self.STORNO_303])
        self.assertEqual(self.zbr(c, "ZBR-3")["Stornirano"], "Da")
        self.assertIn("ZBIRNA_STORNIRANA_PO_ODLUCI", [p for p, _ in self.nalazi(c, "INFO")])

    def test_storniraj_zbirnu_sa_aktivnom_prijemnicom_ne_vazi(self):
        iz = dict(self.STORNO_303, BrojZbirne="301", provera={"BrojZbirne": "301"})
        c = self.sagradi(dodatno=self.dodatno(), izuzeci_dok=[iz])
        self.assertNotEqual(self.zbr(c, "ZBR-1")["Stornirano"], "Da")
        n = [x for x in c.nalazi if x["pravilo"] == "IZUZETAK_DOKUMENTA_POTPIS_SE_NE_SLAZE"]
        self.assertTrue(n and "aktivna otpremnica" in n[0]["opis"] and "aktivna prijemnica" in n[0]["opis"], n)

    def test_prevezivanje_pa_storno_prazne_zbirne(self):
        """Redosled: prvo prevezivanje prijemnica, pa storno -- zbirna sa koje je
        prijemnica skinuta postaje prazna i sme da se stornira."""
        iz_p = dict(self.PREVEZI, PrijemnicaID="PRJ-3", BrojZbirne="303", provera={"BrojZbirne": "302"})
        c = self.sagradi(dodatno=self.dodatno(prijemnica_302=True), izuzeci_dok=[
            iz_p, dict(self.STORNO_303, BrojZbirne="302", provera={"BrojZbirne": "302"})])
        self.assertEqual(self.zbr(c, "ZBR-2")["Stornirano"], "Da")
        self.assertEqual(self.prj(c, "PRJ-3")["ZbirnaRoditeljID"], "ZBR-3")

    def test_ucitavanje_novih_radnji(self):
        import json
        put = os.path.join(self.dir.name, "iz.json")
        with open(put, "w", encoding="utf-8") as f:
            json.dump({"dokumenti": [self.PREVEZI, self.STORNO_303]}, f)
        self.assertEqual(len(cilj.ucitaj_izuzetke_dokumenata(put)), 2)
        with open(put, "w", encoding="utf-8") as f:
            json.dump({"dokumenti": [dict(self.PREVEZI, provera={})]}, f)
        with self.assertRaises(ValueError):
            cilj.ucitaj_izuzetke_dokumenata(put)

    def test_odluke_u_izuzecima_su_ispravnog_oblika(self):
        """izuzeci.json u repou: 8 prijemnica i 4 zbirne (MIG-ODL-20), svaka sa potpisom."""
        lista = cilj.ucitaj_izuzetke_dokumenata(migracija.IZUZECI)
        prev = [iz for iz in lista if iz["radnja"] == "prevezi_prijemnicu"]
        zbr = [iz for iz in lista if iz["radnja"] == "storniraj" and iz.get("tabela") == "tblZbirna"]
        self.assertEqual((len(prev), len(zbr)), (8, 4))
        self.assertTrue(all(iz["provera"].get("BrojZbirne") for iz in prev + zbr))

class MigOdl22StorniranBlok(SaNovomSveskom):
    def clanovi(self, c):
        return sorted((v["OtpremnicaID"], v["OtkupID"]) for v in c.tabele["tblOtpremnicaIzvori"])

    def test_storniran_blok_nije_clan_aktivne_otpremnice(self):
        c = self.sagradi(dodatno=izmeni("tblOtkup", {"OTK-2": {"OtpremnicaID": "OTP-1"}}))
        self.assertNotIn(("OTP-1", "OTK-2"), self.clanovi(c))
        self.assertIn(("OTP-1", "OTK-1"), self.clanovi(c))
        self.assertIn(("STORNIRAN_OTKUP_NIJE_CLAN_AKTIVNE_OTPREMNICE", "OTK-2"), self.nalazi(c, "INFO"))

    def test_stornirana_otpremnica_zadrzava_stornirane_clanove(self):
        c = self.sagradi(dodatno=izmeni("tblOtkup", {"OTK-2": {"OtpremnicaID": "OTP-9"}}))
        self.assertIn(("OTP-9", "OTK-2"), self.clanovi(c))
        self.assertNotIn("STORNIRAN_OTKUP_NIJE_CLAN_AKTIVNE_OTPREMNICE", [p for p, _ in self.nalazi(c)])


if __name__ == "__main__":
    unittest.main()
