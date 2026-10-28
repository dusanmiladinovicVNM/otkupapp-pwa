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
import dokumenti  # noqa: E402
import migracija  # noqa: E402
import pregled  # noqa: E402
from citac import datum, polje, procitaj_svesku, sha256_fajla, tabela  # noqa: E402

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
                      "Kolicina", "Cena", "KolAmbalaze", "KolAmbIzdata", "TipAmbalaze", "Stornirano"],
                     [["OTK-1", D1, "K1", "S1", "101", "I", 100, 80, 10, 4, "12/1", ""],
                      ["OTK-2", D3, "K2", "S1", "102", "I", 50, 80, 5, 0, "12/1", "Da"],
                      ["OTK-3", D3, "K2", "S1", "103", "I", 40, 80, 3, 0, "12/1", ""]]),
        "tblOtpremnica": (["OtpremnicaID", "Datum", "StanicaID", "VozacID", "BrojOtpremnice", "Klasa",
                           "Kolicina", "KolAmbalaze", "Stornirano", "GeneracijaID"],
                          [["OTP-1", D2, "S1", "D1", "201", "I", 100, 10, "", "G-OTP-1"],
                           ["OTP-9", D2, "S1", "D1", "209", "I", 30, 3, "Da", "G-OTP-9"]]),
        "tblZbirna": (["ZbirnaID", "Datum", "VozacID", "BrojZbirne", "Klasa", "UkupnoKolicina",
                       "Stornirano", "GeneracijaID"],
                      [["ZBR-1", D2, "D1", "301", "I", 100, "", "G-ZBR-1"]]),
        "tblPrijemnica": (["PrijemnicaID", "Datum", "KupacID", "VozacID", "BrojPrijemnice", "Klasa",
                           "Kolicina", "KolAmbalaze", "KolAmbVracena", "Stornirano", "GeneracijaID"],
                          [["PRJ-1", D3, "C1", "D1", "401", "I", 90, 10, 7, "", "G-PRJ-1"],
                           ["PRJ-2", D3, "C1", "D1", "401", "II", 10, 0, 0, "", "G-PRJ-1"]]),
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

    def prevedi(self, amb_redovi, izuzeci=(), dodatno=None, partneri=()):
        put = os.path.join(self.dir.name, "stara.xlsx")
        snimi_svesku(put, osnovne_tabele(amb_redovi, dodatno))
        self.tabele = procitaj_svesku(put)
        rez = ambalaza.prevedi(self.tabele, izuzeci, politika_partneri=partneri)
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
        """MIG-16/17 (--pokrice-duga): dan 1 obaveza raste 6 (pokrice), dan 2 pada 6."""
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

    def test_isporuceni_izuzeci_su_ispravni(self):
        lista = ambalaza.ucitaj_izuzetke(migracija.IZUZECI)
        self.assertEqual(len(lista), 7)
        self.assertTrue(all(iz.get("provera") for iz in lista))


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
        kod = tiho(["pregled", "--izvor", put, "--izlaz", izlaz, "--bez-izuzetaka"])
        self.assertEqual(kod, 0)
        self.assertEqual(sha256_fajla(put), pre)
        t = procitaj_svesku(os.path.join(izlaz, "pregled_migracije.xlsx"))
        for ime in ("Salda", "Stavke", "Kretanja", "Reversi", "ZaOdluku", "Dokumenti", "Tabele"):
            self.assertIsNotNone(tabela(t, ime), ime)
        k1 = [r for r in tabela(t, "Salda").redovi if polje(r, "ID") == "K1"]
        self.assertEqual(polje(k1[0], "Naziv"), "Pera Peric")
        self.assertEqual(len(tabela(t, "Kretanja").redovi), 7)

    def test_odbija_izlaz_u_repou(self):
        put = os.path.join(self.dir.name, "stara.xlsx")
        snimi_svesku(put, osnovne_tabele(normalno()))
        izlaz = os.path.join(migracija.KOREN_REPOA, "pregled_test_ne_sme")
        self.assertEqual(tiho(["pregled", "--izvor", put, "--izlaz", izlaz]), 2)
        self.assertFalse(os.path.exists(izlaz))


if __name__ == "__main__":
    unittest.main()
