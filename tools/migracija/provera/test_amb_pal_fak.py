"""Testovi provere amb_pal_fak: cista sveska bez nalaza, svaka klasa podmetnuta u
fixture.kopija(fixture.osnovna()) i njen glavni legitimni slicni slucaj bez nalaza.

Pokretanje (iz tools/migracija):  python3 -m unittest provera.test_amb_pal_fak -v
"""

import time
import unittest
from datetime import date, datetime, timedelta
from decimal import Decimal

from citac import Tabela

from provera import fixture
from provera.fixture import D5, D6, KUPAC, red
from provera.pokretac import pokreni_nad_tabelama
from provera.test_zajednicko import tabele_iz

MOD = ("amb_pal_fak",)
D7 = date(2026, 7, 7)
KUP2 = "KUP-00002"


def nalazi(t, kolone=None):
    ctx = pokreni_nad_tabelama(tabele_iz(t, kolone), moduli=MOD)
    pao = [x.opis for x in ctx.nalazi if x.kod == "MODUL_PAO"]
    if pao:
        raise AssertionError("modul pao:\n" + pao[0])
    return ctx.nalazi


def amb(t, dan, tip, kol, smer, ent, etip, voz, dok, dtip, st=""):
    i = max(int(r["AmbID"][4:]) for r in t["tblAmbalaza"]) + 1
    aid = "AMB-%05d" % i
    t["tblAmbalaza"].append(dict(AmbID=aid, Datum=dan, TipAmbalaze=tip, Kolicina=kol, Smer=smer, EntitetID=ent,
                                 EntitetTip=etip, VozacID=voz, DokumentID=dok, DokumentTIP=dtip, Stornirano=st,
                                 CreatedAt="", CreatedBy="op"))
    return aid


def noge(t, dok):
    return [r for r in t["tblAmbalaza"] if r["DokumentID"] == dok]


def nova_prijemnica(t, od, pid, **izmene):
    p = dict(red(t, "tblPrijemnica", "PrijemnicaID", od))
    p.update(PrijemnicaID=pid, **izmene)
    t["tblPrijemnica"].append(p)
    return p


def storno_prj(t, pid):
    """StornoPrijemnica: red + noge ambalaze (StornoAmbalazaByDokument); paletne stavke ostaju."""
    red(t, "tblPrijemnica", "PrijemnicaID", pid)["Stornirano"] = "Da"
    for r in noge(t, pid):
        r["Stornirano"] = "Da"


def kupac2(t):
    t["tblKupci"].append(dict(KupacID=KUP2, Naziv="Drugi kupac", Aktivan="Da", Hladnjaca=""))


def kod(nal, k_):
    return [x for x in nal if x.kod == k_]


class Osnova(unittest.TestCase):
    def ima(self, nal, k_, oz, id_):
        x = [n for n in nal if n.kod == k_ and n.id == id_]
        self.assertTrue(x, "nema %s na %s; nalazi: %s" % (k_, id_, [(n.kod, n.id, n.ozbiljnost) for n in nal]))
        self.assertEqual(x[0].ozbiljnost, oz, x[0].opis)
        self.assertEqual(len(x), 1, "vise nalaza istog koda na istom dokumentu")
        return x[0]

    def nema(self, nal, *kodovi):
        self.assertEqual([(n.kod, n.id, n.opis) for n in nal if n.kod in kodovi], [])


# ---------------------------------------------------------------------
class CistaSveska(Osnova):
    def test_bez_nalaza(self):
        self.assertEqual([(x.kod, x.id, x.opis) for x in nalazi(fixture.osnovna())], [])

    def test_nedostaju_tabele(self):
        t = fixture.osnovna()
        for ime in ("tblAmbalaza", "tblPaleta", "tblPaletaStavka", "tblFakture", "tblFakturaStavke", "tblNovac",
                    "tblStornoVeze", "tblVozaci"):
            del t[ime]
        self.assertEqual(nalazi(t), [])

    def test_prazne_tabele(self):
        t = fixture.osnovna()
        for ime in ("tblAmbalaza", "tblPaleta", "tblPaletaStavka", "tblPrijemnica", "tblZbirna"):
            t[ime] = []
        self.assertEqual(nalazi(t), [])

    def test_losi_podaci_bez_pada(self):
        t = fixture.kopija(fixture.osnovna())
        red(t, "tblAmbalaza", "AmbID", "AMB-00001").update(Kolicina="abc", Datum="")
        red(t, "tblAmbalaza", "AmbID", "AMB-00002").update(Datum="nije datum", Smer="")
        amb(t, "", "", "x", "Ulaz", "ST-00001", "Stanica", "ST-00002", "", "OMUlaz")
        amb(t, None, "12/1", "", "Izlaz", "", "Kupac", "", "", "Kupci-Otpremnica")
        red(t, "tblPaletaStavka", "StavkaID", "PLS-00001").update(BrojGajbica="", NetoKg="tekst")
        red(t, "tblPaleta", "PaletaID", "PAL-00001").update(BrojGajbica="?", Istorija="act=;;|")
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00003").update(Datum="", Kolicina="n/a")
        nalazi(t)   # bez MODUL_PAO

    def test_kolona_fali_preskoceno(self):
        t = fixture.osnovna()
        kol = {"tblPaletaStavka": [c for c in fixture.KOLONE["tblPaletaStavka"] if c != "BrojZbirne"],
               "tblAmbalaza": [c for c in fixture.KOLONE["tblAmbalaza"] if c != "VozacID"]}
        nal = nalazi(t, kol)
        x = self.ima(nal, "PAL_BROJ_ZBIRNE_ZASTAREO_PRESKOCENO", "INFO", "tblPaletaStavka.BrojZbirne")
        self.assertIn("BrojZbirne", x.opis)
        self.ima(nal, "AMB_REVERS_VOZAC_NIJE_PAR_STANICE_PRESKOCENO", "INFO", "tblAmbalaza.VozacID")
        self.assertTrue(all(n.ozbiljnost == "INFO" and n.kod.endswith("_PRESKOCENO") for n in nal),
                        [(n.kod, n.id) for n in nal])

    def test_losi_podaci_bez_pada_2(self):
        """AmbID prazni (redosled po _redni), datetime sa vremenom, None, tekst umesto broja, dupli storno veze."""
        t = fixture.kopija(fixture.osnovna())
        amb(t, datetime(2026, 7, 6, 23, 59), "12/1", Decimal("1.5"), "ulaz", "st-00001 ", "STANICA", "ST-00002",
            "1/060726", "omulaz")
        for r_ in t["tblAmbalaza"]:
            r_["AmbID"] = ""
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00002").update(Datum=None, Stornirano="Da", KolAmbalaze="x",
                                                                     Kolicina=None)
        red(t, "tblPaleta", "PaletaID", "PAL-00001").update(Istorija=None, BrutoKg="abc", PaletaKg=12)
        t["tblStornoVeze"].append(dict(CorrectionID="C1", OldDocID="PRJ-00002", NewDocID="PRJ-00003"))
        t["tblStornoVeze"].append(dict(CorrectionID="C2", OldDocID="", NewDocID=""))
        nalazi(t)   # bez MODUL_PAO

    def test_kljucne_kolone_fale(self):
        """DokumentTip/Stornirano/KupacID/PaletaID fale: samo PRESKOCENO, bez pada i bez drugih nalaza."""
        kol = {"tblAmbalaza": [c for c in fixture.KOLONE["tblAmbalaza"] if c not in ("DokumentTIP", "Stornirano")],
               "tblPrijemnica": [c for c in fixture.KOLONE["tblPrijemnica"] if c not in ("KupacID", "Stornirano")],
               "tblPaleta": [c for c in fixture.KOLONE["tblPaleta"] if c not in ("BrojGajbica", "PaletaID")]}
        nal = nalazi(fixture.osnovna(), kol)
        self.assertTrue(nal)
        self.assertTrue(all(n.ozbiljnost == "INFO" and n.kod.endswith("_PRESKOCENO") for n in nal),
                        [(n.kod, n.id) for n in nal])
        self.assertEqual(len({n.kod for n in nal}), len(nal), "jedan PRESKOCENO po klasi")

    def test_performanse_tip_i_otkup(self):
        """1000 otkupa i 2500 reversa iste stanice sa tri tipa: prozor tipa ne sme biti O(reversi x otkupi)."""
        t = fixture.kopija(fixture.osnovna())
        base = t["tblOtkup"][0]
        for i in range(1000):
            t["tblOtkup"].append(dict(base, OtkupID="OTK-%05d" % (100 + i), Datum=D5 + timedelta(days=i % 60),
                                      TipAmbalaze=("12/1", "10/1")[i % 2]))
        for i in range(2500):
            dan = D5 + timedelta(days=i % 60)
            amb(t, dan, ("12/1", "10/1", "8/1")[i % 3], 10, "Ulaz", "ST-00001", "Stanica", "ST-00002",
                "1/%s-%d" % (dan.strftime("%d%m%y"), i + 2), "OM-Ulaz-Firma")
        t0 = time.time()
        nalazi(t)
        self.assertLess(time.time() - t0, 20)

    def test_performanse(self):
        """~5000 redova tblAmbalaza (reversi i Kupci-Otpremnica istog naloga, najgori slucaj za grupisanje)
        i ~1000 prijemnica sa stavkama: bez O(n^2)."""
        def tab(ime, kolone, redovi):
            tb = Tabela(ime, ime, kolone)
            for i, r_ in enumerate(redovi, 1):
                x = {c.casefold(): r_.get(c, "") for c in kolone}
                x["_redni"] = i
                tb.redovi.append(x)
            return tb
        amb_r, prj_r, pls_r = [], [], []
        for i in range(1, 5001):
            dan = D5 + timedelta(days=i % 60)
            dtip = ("OMUlaz", "Kupci-Otpremnica", "OM-Ulaz-Firma", "OM-Izlaz-Firma")[i % 4]
            amb_r.append(dict(AmbID="AMB-%05d" % i, Datum=dan, TipAmbalaze=("12/1", "10/1")[i % 2], Kolicina=40,
                              Smer=("Ulaz", "Izlaz")[i % 2], EntitetID=("ST-00001", KUPAC)[i % 4 == 1],
                              EntitetTip=("Stanica", "Kupac")[i % 4 == 1], VozacID="ST-00002",
                              DokumentID="1/%s-%d" % (dan.strftime("%d%m%y"), i), DokumentTip=dtip,
                              Stornirano=""))
        for i in range(1, 1001):
            prj_r.append(dict(PrijemnicaID="PRJ-%05d" % i, Datum=D5 + timedelta(days=i % 60), KupacID=KUPAC,
                              VozacID="ST-00001", BrojPrijemnice="1/050726-%d" % i, BrojZbirne="S1/050726-%d" % i,
                              VrstaVoca="Malina", SortaVoca="Vilamet", Kolicina=100, KolAmbalaze=40, Klasa="I",
                              Stornirano="Da" if i % 10 == 0 else ""))
            pls_r.append(dict(StavkaID="PLS-%05d" % i, PaletaID="PAL-%05d" % (i // 3 + 1), PrijemnicaID="PRJ-%05d" % i,
                              BrojPrijemnice="1/050726-%d" % i, BrojZbirne="S1/050726-%d" % i, Klasa="I",
                              VrstaVoca="Malina", SortaVoca="Vilamet", BrojGajbica=40, NetoKg=100))
        tabele = {"tblambalaza": tab("tblAmbalaza", fixture.KOLONE["tblAmbalaza"][:9] + ["DokumentTip", "Stornirano"],
                                     amb_r),
                  "tblprijemnica": tab("tblPrijemnica", fixture.KOLONE["tblPrijemnica"], prj_r),
                  "tblpaletastavka": tab("tblPaletaStavka", fixture.KOLONE["tblPaletaStavka"], pls_r),
                  "tblstanice": tab("tblStanice", ["StanicaID"], [dict(StanicaID="ST-00001"),
                                                                  dict(StanicaID="ST-00002")])}
        t0 = time.time()
        ctx = pokreni_nad_tabelama(tabele, moduli=MOD)
        self.assertEqual([x.opis for x in ctx.nalazi if x.kod == "MODUL_PAO"], [])
        self.assertLess(time.time() - t0, 20)


# ---------------------------------------------------------------------
class AmbNogaStorniranog(Osnova):
    K = "AMB_AKTIVNA_NOGA_STORNIRANOG_DOKUMENTA"

    def test_greska(self):
        t = fixture.kopija(fixture.osnovna())
        red(t, "tblOtkup", "OtkupID", "OTK-00003")["Stornirano"] = "Da"
        ids_ = ", ".join(r["AmbID"] for r in noge(t, "OTK-00003"))
        x = self.ima(nalazi(t), self.K, "GRESKA", ids_)
        self.assertIn("OTK-00003", x.opis)

    def test_storno_pa_ponovni_unos_istog_broja(self):
        t = fixture.kopija(fixture.osnovna())
        o = red(t, "tblOtkup", "OtkupID", "OTK-00003")
        o["Stornirano"] = "Da"
        for r in noge(t, "OTK-00003"):
            r["Stornirano"] = "Da"
        novi = dict(o, OtkupID="OTK-00007", Stornirano="")
        t["tblOtkup"].append(novi)
        amb(t, D5, "12/1", 20, "Izlaz", "KOOP-00001", "Kooperant", "", "OTK-00007", "Otkup")
        amb(t, D5, "12/1", 20, "Ulaz", "ST-00001", "Stanica", "", "OTK-00007", "Otkup")
        self.nema(nalazi(t), self.K)

    def test_dve_klase_jedan_nalaz(self):
        t = fixture.kopija(fixture.osnovna())
        nova_prijemnica(t, "PRJ-00004", "PRJ-00006", Klasa="II", KolAmbalaze=5, Kolicina=10)
        a = amb(t, D6, "12/1", 5, "Ulaz", KUPAC, "Kupac", "ST-00001", "PRJ-00006", "Prijemnica")
        for pid in ("PRJ-00004", "PRJ-00006"):
            red(t, "tblPrijemnica", "PrijemnicaID", pid)["Stornirano"] = "Da"
        ids_ = ", ".join([r_["AmbID"] for r_ in noge(t, "PRJ-00004")] + [a])
        self.ima(nalazi(t), self.K, "GRESKA", ids_)

    def test_dupli_id_aktivan_red_ima_prednost(self):
        """Isti OtkupID dvaput (storniran i aktivan red): roditelj nije storniran."""
        t = fixture.kopija(fixture.osnovna())
        o = red(t, "tblOtkup", "OtkupID", "OTK-00003")
        t["tblOtkup"].insert(0, dict(o, Stornirano="Da"))
        self.nema(nalazi(t), self.K)


# ---------------------------------------------------------------------
class AmbReversVozac(Osnova):
    K = "AMB_REVERS_VOZAC_NIJE_PAR_STANICE"
    K2 = "AMB_REVERS_STANICA_ZASTARELA"

    def test_vozac_zastareo(self):
        t = fixture.kopija(fixture.osnovna())
        a = amb(t, D6, "12/1", 30, "Ulaz", "ST-00001", "Stanica", "ST-00002", "1/060726", "OM-Ulaz-Firma")
        x = self.ima(nalazi(t), self.K, "SUMNJA", a)
        self.assertIn("VozacID: ST-00002 -> ST-00001", x.predlog)

    def test_stanica_zastarela(self):
        t = fixture.kopija(fixture.osnovna())
        a = amb(t, D6, "12/1", 30, "Izlaz", "ST-00002", "Stanica", "ST-00001", "2/060726", "OM-Izlaz-Firma")
        nal = nalazi(t)
        x = self.ima(nal, self.K2, "SUMNJA", a)
        self.assertIn("EntitetID: ST-00002 -> ST-00001", x.predlog)
        self.nema(nal, self.K)

    def test_za_odluku(self):
        t = fixture.kopija(fixture.osnovna())
        a = amb(t, D5, "12/1", 30, "Ulaz", "ST-00001", "Stanica", "ST-00002", "1/050726", "OMUlaz")
        x = self.ima(nalazi(t), self.K, "SUMNJA", a)
        self.assertTrue(x.predlog.startswith("ZA ODLUKU"), x.predlog)

    def test_legitimno(self):
        t = fixture.kopija(fixture.osnovna())
        amb(t, D6, "12/1", 30, "Ulaz", "ST-00001", "Stanica", "ST-00001", "1/060726", "OM-Ulaz-Firma")
        amb(t, D6, "12/1", 10, "Ulaz", "KOOP-00002", "Kooperant", "", "1/060726-2", "OM-Izlaz-Koop")
        amb(t, D6, "12/1", 10, "Izlaz", "ST-00001", "Stanica", "", "1/060726-2", "OM-Izlaz-Koop")
        # stanica bez par-vozaca u tblVozaci: vozac je uvek rucni
        t["tblVozaci"] = [v for v in t["tblVozaci"] if v["VozacID"] != "ST-00003"]
        amb(t, D6, "12/1", 30, "Ulaz", "ST-00003", "Stanica", "ST-00001", "3/060726", "OM-Ulaz-Firma")
        self.nema(nalazi(t), self.K, self.K2)

    def test_bez_tabele_vozaca(self):
        """Bez tblVozaci se ne zna koja stanica ima par-vozaca: provera se tiho preskace."""
        t = fixture.kopija(fixture.osnovna())
        del t["tblVozaci"]
        amb(t, D6, "12/1", 30, "Ulaz", "ST-00001", "Stanica", "ST-00002", "1/060726", "OM-Ulaz-Firma")
        self.nema(nalazi(t), self.K, self.K2)


class AmbKupciOtpremnica(Osnova):
    KV = "AMB_REVERS_VOZAC_NIJE_PAR_STANICE"
    KK = "KUPAC_ZASTAREO_NA_KUPCI_OTPREMNICI_I_UPLATI"

    def test_vozac_bez_prijemnice_kod_kupca(self):
        t = fixture.kopija(fixture.osnovna())
        a = amb(t, D5, "12/1", 20, "Izlaz", KUPAC, "Kupac", "ST-00003", "K-1", "Kupci-Otpremnica")
        self.ima(nalazi(t), self.KV, "INFO", a)

    def test_kupac_bez_prijemnica(self):
        t = fixture.kopija(fixture.osnovna())
        kupac2(t)
        a = amb(t, D6, "12/1", 20, "Izlaz", KUP2, "Kupac", "ST-00001", "K-1", "Kupci-Otpremnica")
        nal = nalazi(t)
        self.ima(nal, self.KK, "INFO", a)
        self.nema(nal, self.KV)

    def test_kupac_samo_ambalaza_jedan_nalaz(self):
        """Kupac bez ijedne prijemnice u svesci: jedan INFO za sve njegove Kupci-Otpremnice (bez preplavljivanja)."""
        t = fixture.kopija(fixture.osnovna())
        kupac2(t)
        a1 = amb(t, D5, "12/1", 50, "Izlaz", KUP2, "Kupac", "ST-00001", "K-1", "Kupci-Otpremnica")
        a2 = amb(t, D6, "12/1", 30, "Izlaz", KUP2, "Kupac", "ST-00002", "K-2", "Kupci-Otpremnica")
        nal = nalazi(t)
        self.ima(nal, self.KK, "INFO", "%s, %s" % (a1, a2))
        self.assertEqual(len(kod(nal, self.KK)), 1)

    def test_pozajmica_pred_sezonu(self):
        """Hladnjaca daje prazne pre prve prijemnice (01.07, prijemnice od 05.07): nije zastareo kupac."""
        t = fixture.kopija(fixture.osnovna())
        amb(t, date(2026, 7, 1), "12/1", 50, "Izlaz", KUPAC, "Kupac", "ST-00001", "K-1", "Kupci-Otpremnica")
        self.nema(nalazi(t), self.KK, self.KV)

    def test_vozac_bez_poredjenja(self):
        """Kupac u poslednjih 7 dana nema nijednu prijemnicu: nema drugog vozaca za poredjenje (prazne kasnije)."""
        t = fixture.kopija(fixture.osnovna())
        amb(t, D6 + timedelta(days=10), "12/1", 20, "Izlaz", KUPAC, "Kupac", "ST-00003", "K-1", "Kupci-Otpremnica")
        self.nema(nalazi(t), self.KK, self.KV)

    def test_prazne_vracene_kasnije(self):
        t = fixture.kopija(fixture.osnovna())
        amb(t, D6 + timedelta(days=5), "12/1", 20, "Izlaz", KUPAC, "Kupac", "ST-00001", "K-1", "Kupci-Otpremnica")
        self.nema(nalazi(t), self.KV, self.KK)


# ---------------------------------------------------------------------
class AmbReversTip(Osnova):
    K = "AMB_REVERS_TIP_ZASTAREO"

    def test_tip_van_robnih(self):
        t = fixture.kopija(fixture.osnovna())
        a1 = amb(t, D5, "10/1", 10, "Izlaz", "KOOP-00001", "Kooperant", "", "1/050726-5", "OM-Ulaz-Koop")
        a2 = amb(t, D5, "10/1", 10, "Ulaz", "ST-00001", "Stanica", "", "1/050726-5", "OM-Ulaz-Koop")
        x = self.ima(nalazi(t), self.K, "INFO", "%s, %s" % (a1, a2))
        self.assertIn("10/1 -> 12/1", x.predlog)

    def test_tip_nasledjen(self):
        t = fixture.kopija(fixture.osnovna())
        red(t, "tblOtkup", "OtkupID", "OTK-00006")["TipAmbalaze"] = "10/1"
        for r in noge(t, "OTK-00006"):
            r["TipAmbalaze"] = "10/1"
        amb(t, D6, "10/1", 10, "Izlaz", "KOOP-00004", "Kooperant", "", "3/060726", "OM-Ulaz-Koop")
        amb(t, D6, "10/1", 10, "Ulaz", "ST-00003", "Stanica", "", "3/060726", "OM-Ulaz-Koop")
        a1 = amb(t, D6, "10/1", 8, "Izlaz", "KOOP-00001", "Kooperant", "", "1/060726", "OM-Ulaz-Koop")
        a2 = amb(t, D6, "10/1", 8, "Ulaz", "ST-00001", "Stanica", "", "1/060726", "OM-Ulaz-Koop")
        nal = nalazi(t)
        x = self.ima(nal, self.K, "SUMNJA", "%s, %s" % (a1, a2))
        self.assertIn("nasledjen", x.opis)
        self.assertEqual(len(kod(nal, self.K)), 1)

    def test_kooperant_koristi_oba_tipa(self):
        t = fixture.kopija(fixture.osnovna())
        red(t, "tblOtkup", "OtkupID", "OTK-00002")["TipAmbalaze"] = "10/1"
        amb(t, D6, "10/1", 10, "Izlaz", "KOOP-00002", "Kooperant", "", "1/060726", "OM-Ulaz-Koop")
        amb(t, D6, "10/1", 10, "Ulaz", "ST-00001", "Stanica", "", "1/060726", "OM-Ulaz-Koop")
        self.nema(nalazi(t), self.K)


# ---------------------------------------------------------------------
def novac_om(t, nid, broj, dan, st, koop):
    t["tblNovac"].append(dict(NovacID=nid, BrojDokumenta=broj, Datum=dan, Partner="Stanica", PartnerID=st,
                              EntitetTip="OM", OMID=st, KooperantID=koop, VrstaVoca="Malina", Tip="KesOtkupacKoop",
                              Uplata=0, Isplata=1000, Stornirano=""))


class AmbOMUlaz(Osnova):
    K = "AMB_OMULAZ_ZABORAVLJEN_SMER"

    def _izdato_koop3(self, t, dan=D5):
        # IZDAVANJE 42 praznih kooperantu KOOP-00003 (otkup 32): kooperant drzi +10
        amb(t, dan, "12/1", 42, "Ulaz", "KOOP-00003", "Kooperant", "", "2/%s" % dan.strftime("%d%m%y"),
            "OM-Izlaz-Koop")
        amb(t, dan, "12/1", 42, "Izlaz", "ST-00002", "Stanica", "", "2/%s" % dan.strftime("%d%m%y"), "OM-Izlaz-Koop")

    def test_kooperant_nije_razduzen(self):
        t = fixture.kopija(fixture.osnovna())
        self._izdato_koop3(t)
        a = amb(t, D6, "12/1", 10, "Ulaz", "ST-00002", "Stanica", "ST-00002", "2/060726", "OMUlaz")
        novac_om(t, "NOV-00003", "2/060726", D6, "ST-00002", "KOOP-00003")
        x = self.ima(nalazi(t), self.K, "SUMNJA", a)
        self.assertIn("OM-Ulaz-Koop", x.predlog)

    def test_kooperant_nije_zaduzen(self):
        t = fixture.kopija(fixture.osnovna())
        a = amb(t, D6, "12/1", 36, "Ulaz", "ST-00003", "Stanica", "ST-00003", "3/060726", "OMUlaz")
        novac_om(t, "NOV-00003", "3/060726", D6, "ST-00003", "KOOP-00004")
        x = self.ima(nalazi(t), self.K, "SUMNJA", a)
        self.assertIn("OM-Izlaz-Koop", x.predlog)

    def test_redovan_tok_firma_salje_prazne_i_novac(self):
        t = fixture.kopija(fixture.osnovna())
        amb(t, D6, "12/1", 10, "Ulaz", "ST-00001", "Stanica", "ST-00001", "1/060726", "OMUlaz")
        novac_om(t, "NOV-00003", "1/060726", D6, "ST-00001", "KOOP-00001")
        self.nema(nalazi(t), self.K)

    def test_pre_uvodjenja_smera(self):
        t = fixture.kopija(fixture.osnovna())
        rano = date(2026, 6, 20)
        self._izdato_koop3(t, rano)
        amb(t, rano, "12/1", 10, "Ulaz", "ST-00002", "Stanica", "ST-00002", "2/200626-2", "OMUlaz")
        novac_om(t, "NOV-00003", "2/200626-2", rano, "ST-00002", "KOOP-00003")
        self.nema(nalazi(t), self.K)


# ---------------------------------------------------------------------
class AmbBezStorna(Osnova):
    K = "AMB_BEZ_STORNA_KUPCI_OTPREMNICA_OMULAZ"

    def test_dupli_uzastopni(self):
        t = fixture.kopija(fixture.osnovna())
        a1 = amb(t, D5, "12/1", 20, "Izlaz", KUPAC, "Kupac", "ST-00001", "K-1", "Kupci-Otpremnica")
        a2 = amb(t, D5, "12/1", 20, "Izlaz", KUPAC, "Kupac", "ST-00001", "K-2", "Kupci-Otpremnica")
        self.ima(nalazi(t), self.K, "SUMNJA", "%s, %s" % (a1, a2))

    def test_dve_ture_istog_vozaca(self):
        """Isti vozac, kupac, kolicina i dan, ali drugi broj i drugi unosi izmedju: dve stvarne ture."""
        t = fixture.kopija(fixture.osnovna())
        amb(t, D5, "12/1", 20, "Izlaz", KUPAC, "Kupac", "ST-00001", "K-1", "Kupci-Otpremnica")
        amb(t, D5, "12/1", 5, "Ulaz", "ST-00001", "Stanica", "ST-00001", "1/050726-8", "OM-Ulaz-Firma")
        amb(t, D5, "12/1", 20, "Izlaz", KUPAC, "Kupac", "ST-00001", "K-2", "Kupci-Otpremnica")
        self.nema(nalazi(t), self.K)

    def test_dupli_bez_broja(self):
        t = fixture.kopija(fixture.osnovna())
        a1 = amb(t, D5, "12/1", 20, "Ulaz", "ST-00001", "Stanica", "ST-00001", "", "OMUlaz")
        amb(t, D5, "12/1", 5, "Ulaz", "ST-00001", "Stanica", "ST-00001", "1/050726-8", "OM-Ulaz-Firma")
        a2 = amb(t, D5, "12/1", 20, "Ulaz", "ST-00001", "Stanica", "ST-00001", "", "OMUlaz")
        self.ima(nalazi(t), self.K, "SUMNJA", "%s, %s" % (a1, a2))

    def test_razliciti_vozaci_nisu_duplikat(self):
        t = fixture.kopija(fixture.osnovna())
        amb(t, D5, "12/1", 20, "Izlaz", KUPAC, "Kupac", "ST-00001", "K-1", "Kupci-Otpremnica")
        amb(t, D5, "12/1", 20, "Izlaz", KUPAC, "Kupac", "ST-00002", "K-2", "Kupci-Otpremnica")
        self.nema(nalazi(t), self.K)

    def test_omulaz_dva_kooperanta(self):
        """Dva uzastopna OMUlaz iste kolicine, svaki uz novac DRUGOM kooperantu: odvojene isplate (redovan tok)."""
        t = fixture.kopija(fixture.osnovna())
        amb(t, D6, "12/1", 10, "Ulaz", "ST-00001", "Stanica", "ST-00001", "1/060726", "OMUlaz")
        novac_om(t, "NOV-00003", "1/060726", D6, "ST-00001", "KOOP-00001")
        amb(t, D6, "12/1", 10, "Ulaz", "ST-00001", "Stanica", "ST-00001", "1/060726-2", "OMUlaz")
        novac_om(t, "NOV-00004", "1/060726-2", D6, "ST-00001", "KOOP-00002")
        self.nema(nalazi(t), self.K)

    def test_omulaz_dupli_isti_kooperant(self):
        t = fixture.kopija(fixture.osnovna())
        a1 = amb(t, D6, "12/1", 10, "Ulaz", "ST-00001", "Stanica", "ST-00001", "1/060726", "OMUlaz")
        novac_om(t, "NOV-00003", "1/060726", D6, "ST-00001", "KOOP-00001")
        a2 = amb(t, D6, "12/1", 10, "Ulaz", "ST-00001", "Stanica", "ST-00001", "1/060726-2", "OMUlaz")
        novac_om(t, "NOV-00004", "1/060726-2", D6, "ST-00001", "KOOP-00001")
        self.ima(nalazi(t), self.K, "SUMNJA", "%s, %s" % (a1, a2))

    def test_storniran_tip_bez_storna(self):
        t = fixture.kopija(fixture.osnovna())
        a = amb(t, D5, "12/1", 20, "Izlaz", KUPAC, "Kupac", "ST-00001", "K-1", "Kupci-Otpremnica", st="Da")
        self.ima(nalazi(t), self.K, "INFO", a)


class AmbKompenzacija(Osnova):
    K = "AMB_KOMPENZACIJA_BEZ_STORNA"

    def test_omulaz_pa_vracanje_firmi(self):
        t = fixture.kopija(fixture.osnovna())
        a1 = amb(t, D5, "12/1", 15, "Ulaz", "ST-00001", "Stanica", "ST-00001", "1/050726-7", "OMUlaz")
        a2 = amb(t, D6, "12/1", 15, "Izlaz", "ST-00001", "Stanica", "ST-00001", "1/060726-2", "OM-Izlaz-Firma")
        self.ima(nalazi(t), self.K, "INFO", "%s, %s" % (a1, a2))

    def test_omulaz_pa_izdavanje_kooperantu(self):
        t = fixture.kopija(fixture.osnovna())
        amb(t, D5, "12/1", 15, "Ulaz", "ST-00001", "Stanica", "ST-00001", "1/050726-7", "OMUlaz")
        amb(t, D6, "12/1", 15, "Ulaz", "KOOP-00002", "Kooperant", "", "1/060726-2", "OM-Izlaz-Koop")
        amb(t, D6, "12/1", 15, "Izlaz", "ST-00001", "Stanica", "", "1/060726-2", "OM-Izlaz-Koop")
        self.nema(nalazi(t), self.K)


class AmbBrojReversa(Osnova):
    K = "AMB_BROJ_REVERSA_NE_PRATI_DATUM"

    def test_broj_drugog_dana(self):
        t = fixture.kopija(fixture.osnovna())
        a = amb(t, D6, "12/1", 30, "Ulaz", "ST-00001", "Stanica", "ST-00001", "1/050726-4", "OM-Ulaz-Firma")
        x = self.ima(nalazi(t), self.K, "INFO", a)
        self.assertIn("2026-07-05", x.opis)

    def test_rucni_broj_bez_predloga(self):
        t = fixture.kopija(fixture.osnovna())
        amb(t, D6, "12/1", 30, "Ulaz", "ST-00001", "Stanica", "ST-00001", "1/050726-4", "OMUlaz")
        amb(t, D6, "12/1", 30, "Izlaz", KUPAC, "Kupac", "ST-00001", "1/050726-5", "Kupci-Otpremnica")
        amb(t, D6, "12/1", 30, "Ulaz", "ST-00001", "Stanica", "ST-00001", "1/060726", "OM-Ulaz-Firma")
        self.nema(nalazi(t), self.K)


class AmbDatumIzIspravke(Osnova):
    KR = "AMB_REVERS_DATUM_ZASTAREO"
    KN = "AMB_DATUM_NASLEDJEN_IZ_ISPRAVKE"

    def _scena(self, t, r_dtip, storno_staru=True, pre_dan=D7):
        """Ispravka PRJ-00004 (D6) uneta 07.07: pre nje unos od 07.07, posle R sa Datum 06.07, pa opet 07.07."""
        if storno_staru:
            red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00004")["Stornirano"] = "Da"
            for r in noge(t, "PRJ-00004"):
                r["Stornirano"] = "Da"
        amb(t, pre_dan, "12/1", 5, "Ulaz", "ST-00001", "Stanica", "ST-00001", pre_dan.strftime("1/%d%m%y"),
            "OM-Ulaz-Firma")
        nova_prijemnica(t, "PRJ-00004", "PRJ-00006", BrojPrijemnice="1/060726-3", Stornirano="")
        amb(t, D6, "12/1", 28, "Ulaz", KUPAC, "Kupac", "ST-00001", "PRJ-00006", "Prijemnica")
        amb(t, D6, "12/1", 28, "Izlaz", KUPAC, "Kupac", "ST-00001", "PRJ-00006", "Prijemnica")
        if r_dtip == "Kupci-Otpremnica":
            a = amb(t, D6, "12/1", 7, "Izlaz", KUPAC, "Kupac", "ST-00001", "K-7", r_dtip)
        else:
            a = amb(t, D6, "12/1", 7, "Ulaz", "ST-00001", "Stanica", "ST-00001", "1/060726", r_dtip)
        amb(t, D7, "12/1", 6, "Ulaz", "ST-00001", "Stanica", "ST-00001", "1/070726-2", "OM-Ulaz-Firma")
        return a

    def test_revers_posle_ispravke(self):
        t = fixture.kopija(fixture.osnovna())
        a = self._scena(t, "OM-Ulaz-Firma")
        x = self.ima(nalazi(t), self.KR, "SUMNJA", a)
        self.assertIn("2026-07-06 -> 2026-07-07", x.predlog)

    def test_kupci_otpremnica_posle_ispravke(self):
        t = fixture.kopija(fixture.osnovna())
        a = self._scena(t, "Kupci-Otpremnica")
        nal = nalazi(t)
        self.ima(nal, self.KN, "SUMNJA", a)
        self.nema(nal, self.KR)

    def test_naknadni_unos_starih_papira(self):
        """Bez storniranog prethodnika nova prijemnica nije ispravka: stari datum je naknadni unos."""
        t = fixture.kopija(fixture.osnovna())
        self._scena(t, "OM-Ulaz-Firma", storno_staru=False)
        self.nema(nalazi(t), self.KR, self.KN)

    def test_ispravka_istog_dana(self):
        t = fixture.kopija(fixture.osnovna())
        self._scena(t, "OM-Ulaz-Firma", pre_dan=D6)
        self.nema(nalazi(t), self.KR, self.KN)


# ---------------------------------------------------------------------
class PalBrojZbirne(Osnova):
    K = "PAL_BROJ_ZBIRNE_ZASTAREO"

    def test_stara_zbirna_na_stavci(self):
        t = fixture.kopija(fixture.osnovna())
        red(t, "tblPaletaStavka", "StavkaID", "PLS-00002")["BrojZbirne"] = "S1/050726"
        x = self.ima(nalazi(t), self.K, "GRESKA", "PLS-00002")
        self.assertIn("S1/050726 -> S1/050726-2", x.predlog)

    def test_prazna_na_stavci(self):
        t = fixture.kopija(fixture.osnovna())
        red(t, "tblPaletaStavka", "StavkaID", "PLS-00002")["BrojZbirne"] = ""
        self.ima(nalazi(t), self.K, "SUMNJA", "PLS-00002")

    def test_razmak_i_velika_slova_i_uskladjeno(self):
        t = fixture.kopija(fixture.osnovna())
        red(t, "tblPaletaStavka", "StavkaID", "PLS-00002")["BrojZbirne"] = "s1/ 050726-2"
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00001")["BrojZbirne"] = "S1/050726-2"
        red(t, "tblPaletaStavka", "StavkaID", "PLS-00001")["BrojZbirne"] = "S1/050726-2"
        self.nema(nalazi(t), self.K)


class PalGajbeKg(Osnova):
    K = "PAL_GAJBE_KG_NESAGLASNI"

    def test_gajbe(self):
        t = fixture.kopija(fixture.osnovna())
        red(t, "tblPaletaStavka", "StavkaID", "PLS-00001")["BrojGajbica"] = 38
        nal = nalazi(t)
        self.ima(nal, self.K, "GRESKA", "PRJ-00001")
        x = self.ima(nal, self.K, "GRESKA", "PAL-00001")
        self.assertIn("BrojGajbica 60 -> 58", x.opis)

    def test_samo_kg(self):
        t = fixture.kopija(fixture.osnovna())
        red(t, "tblPaletaStavka", "StavkaID", "PLS-00001")["NetoKg"] = Decimal("97")
        nal = nalazi(t)
        self.ima(nal, self.K, "SUMNJA", "PRJ-00001")
        self.ima(nal, self.K, "SUMNJA", "PAL-00001")

    def test_preliv_na_dve_palete_sa_zaokruzivanjem(self):
        t = fixture.kopija(fixture.osnovna())
        red(t, "tblPaletaStavka", "StavkaID", "PLS-00001").update(BrojGajbica=30, NetoKg=Decimal("74.2"))
        red(t, "tblPaleta", "PaletaID", "PAL-00001").update(BrojGajbica=50, NetoKg=Decimal("123.7"))
        t["tblPaleta"].append(dict(red(t, "tblPaleta", "PaletaID", "PAL-00001"), PaletaID="PAL-00002",
                                   BrojPalete="2", BrojGajbica=10, NetoKg=Decimal("24.75"), Status="Otvorena"))
        t["tblPaletaStavka"].append(dict(red(t, "tblPaletaStavka", "StavkaID", "PLS-00001"), StavkaID="PLS-00003",
                                         PaletaID="PAL-00002", BrojGajbica=10, NetoKg=Decimal("24.75")))
        self.nema(nalazi(t), self.K)


class PalIdentitet(Osnova):
    K = "PAL_IDENTITET_RAZLICIT_OD_PRIJEMNICE"

    def test_sorta_stavke(self):
        t = fixture.kopija(fixture.osnovna())
        red(t, "tblPaletaStavka", "StavkaID", "PLS-00002")["SortaVoca"] = "Miker"
        x = self.ima(nalazi(t), self.K, "SUMNJA", "PAL-00001")
        self.assertIn("PLS-00002 SortaVoca", x.opis)

    def test_zaglavlje_palete(self):
        t = fixture.kopija(fixture.osnovna())
        red(t, "tblPaleta", "PaletaID", "PAL-00001")["SortaVoca"] = "Miker"
        x = self.ima(nalazi(t), self.K, "SUMNJA", "PAL-00001")
        self.assertIn("Relabel", x.predlog)

    def test_prazna_sorta_je_nepoznata(self):
        t = fixture.kopija(fixture.osnovna())
        red(t, "tblPaletaStavka", "StavkaID", "PLS-00002")["SortaVoca"] = ""
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00001")["TipAmbalaze"] = ""
        self.nema(nalazi(t), self.K)


class PalPrevezivanje(Osnova):
    K = "PAL_PREVEZIVANJE_PO_BROJU_TUDJI_KUPAC"
    KD = "PAL_DETACH_PO_BROJU_TUDJA_PRIJEMNICA"

    def test_stavka_dobila_zbirnu_tudje_prijemnice(self):
        t = fixture.kopija(fixture.osnovna())
        kupac2(t)
        nova_prijemnica(t, "PRJ-00001", "PRJ-00006", KupacID=KUP2, BrojZbirne="S2/050726", Stornirano="Da")
        red(t, "tblPaletaStavka", "StavkaID", "PLS-00001")["BrojZbirne"] = "S2/050726"
        nal = nalazi(t)
        x = self.ima(nal, self.K, "GRESKA", "PLS-00001")
        self.assertIn("PRJ-00006", x.predlog)
        self.nema(nal, "PAL_BROJ_ZBIRNE_ZASTAREO")

    def _bez_stavki(self, t, istorija="", paleta_storno=""):
        kupac2(t)
        nova_prijemnica(t, "PRJ-00001", "PRJ-00006", KupacID=KUP2, BrojZbirne="", KolAmbalaze=10,
                        Kolicina=Decimal("25"))
        t["tblPaleta"].append(dict(red(t, "tblPaleta", "PaletaID", "PAL-00001"), PaletaID="PAL-00002",
                                   BrojPalete="2", BrojGajbica=0, NetoKg=0, Status="Otvorena", Istorija=istorija,
                                   Stornirano=paleta_storno))
        t["tblPaletaStavka"].append(dict(red(t, "tblPaletaStavka", "StavkaID", "PLS-00001"), StavkaID="PLS-00003",
                                         PaletaID="PAL-00002", PrijemnicaID="PRJ-00006", BrojZbirne="",
                                         BrojGajbica=10, NetoKg=25, Stornirano="Da"))

    def test_prijemnica_tudjeg_kupca_ostala_bez_stavki(self):
        t = fixture.kopija(fixture.osnovna())
        self._bez_stavki(t)
        nal = nalazi(t)
        self.ima(nal, self.K, "SUMNJA", "PRJ-00006")
        self.nema(nal, self.KD)

    def test_detach_po_broju(self):
        t = fixture.kopija(fixture.osnovna())
        self._bez_stavki(t, istorija="act=DETACH;det=prij=1/050726 gajb=10;t=2026-07-08 10:00 | "
                                     "act=STORNO_PRAZNA;det=detach prij=1/050726;t=2026-07-08 10:00",
                         paleta_storno="Da")
        nal = nalazi(t)
        x = self.ima(nal, self.KD, "SUMNJA", "PRJ-00006")
        self.assertIn("DETACH", x.opis)
        self.nema(nal, self.K)

    def test_gajbe_nesaglasne_kod_kolizije_greska(self):
        """Kolizija broja ne sme da obori ozbiljnost: zbir gajbi po PrijemnicaID != KolAmbalaze je dokazan."""
        t = fixture.kopija(fixture.osnovna())
        kupac2(t)
        nova_prijemnica(t, "PRJ-00001", "PRJ-00006", KupacID=KUP2, BrojZbirne="S2/050726", Stornirano="Da")
        red(t, "tblPaletaStavka", "StavkaID", "PLS-00001")["BrojGajbica"] = 38
        red(t, "tblPaleta", "PaletaID", "PAL-00001")["BrojGajbica"] = 58
        nal = nalazi(t)
        self.ima(nal, self.K, "GRESKA", "PLS-00001")
        self.nema(nal, "PAL_GAJBE_KG_NESAGLASNI")

    def test_bez_kolizije_broja(self):
        """Isti kupac, drugi broj: detach duplog unosa je legitiman."""
        t = fixture.kopija(fixture.osnovna())
        nova_prijemnica(t, "PRJ-00001", "PRJ-00006", BrojPrijemnice="1/050726-4", KolAmbalaze=10)
        t["tblPaletaStavka"].append(dict(red(t, "tblPaletaStavka", "StavkaID", "PLS-00001"), StavkaID="PLS-00003",
                                         PrijemnicaID="PRJ-00006", BrojPrijemnice="1/050726-4", Stornirano="Da"))
        self.nema(nalazi(t), self.K, self.KD)


class PalZbirnaPoBroju(Osnova):
    K = "PAL_ZBIRNA_PREVEZANA_PO_BROJU"

    def test_ista_zbirna_dva_kupca_istog_broja(self):
        t = fixture.kopija(fixture.osnovna())
        kupac2(t)
        nova_prijemnica(t, "PRJ-00001", "PRJ-00006", KupacID=KUP2, BrojPrijemnice="1/050726 ")
        x = self.ima(nalazi(t), self.K, "SUMNJA", "PRJ-00006")
        self.assertIn("PRJ-00001", x.opis)

    def test_drugi_broj(self):
        t = fixture.kopija(fixture.osnovna())
        kupac2(t)
        nova_prijemnica(t, "PRJ-00001", "PRJ-00006", KupacID=KUP2, BrojPrijemnice="7/050726")
        self.nema(nalazi(t), self.K)


class PalMesaniKupci(Osnova):
    K = "PAL_MESANI_KUPCI"

    def test_dva_kupca_na_paleti(self):
        t = fixture.kopija(fixture.osnovna())
        kupac2(t)
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00002")["KupacID"] = KUP2
        self.ima(nalazi(t), self.K, "INFO", "PAL-00001")

    def test_jedan_kupac(self):
        self.nema(nalazi(fixture.osnovna()), self.K)

    def test_stavke_bez_palete(self):
        """Stavke bez PaletaID nisu jedna paleta."""
        t = fixture.kopija(fixture.osnovna())
        kupac2(t)
        for x in t["tblPaletaStavka"]:
            x["PaletaID"] = ""
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00002")["KupacID"] = KUP2
        self.nema(nalazi(t), self.K, "PAL_IDENTITET_RAZLICIT_OD_PRIJEMNICE")


class PalNaStorniranoj(Osnova):
    K = "PAL_STAVKA_NA_STORNIRANOJ_PRIJEMNICI"

    def test_osirocena(self):
        t = fixture.kopija(fixture.osnovna())
        storno_prj(t, "PRJ-00002")
        x = self.ima(nalazi(t), self.K, "GRESKA", "PLS-00002")
        self.assertIn("Nema zamene", x.predlog)

    def test_zamena_paletizovana_dvaput(self):
        t = fixture.kopija(fixture.osnovna())
        storno_prj(t, "PRJ-00002")
        nova_prijemnica(t, "PRJ-00002", "PRJ-00006", BrojPrijemnice="1/050726-4", Stornirano="")
        t["tblPaleta"].append(dict(red(t, "tblPaleta", "PaletaID", "PAL-00001"), PaletaID="PAL-00002",
                                   BrojPalete="2", BrojGajbica=20, NetoKg=Decimal("49.5"), Status="Otvorena"))
        t["tblPaletaStavka"].append(dict(red(t, "tblPaletaStavka", "StavkaID", "PLS-00002"), StavkaID="PLS-00003",
                                         PaletaID="PAL-00002", PrijemnicaID="PRJ-00006", BrojPrijemnice="1/050726-4"))
        x = self.ima(nalazi(t), self.K, "GRESKA", "PLS-00002")
        self.assertIn("dvaput", x.opis)

    def test_stavka_skinuta(self):
        t = fixture.kopija(fixture.osnovna())
        storno_prj(t, "PRJ-00002")
        red(t, "tblPaletaStavka", "StavkaID", "PLS-00002")["Stornirano"] = "Da"
        red(t, "tblPaleta", "PaletaID", "PAL-00001").update(BrojGajbica=40, NetoKg=Decimal("99"))
        self.assertEqual([(x.kod, x.id) for x in nalazi(t)], [])


# ---------------------------------------------------------------------
FAK_KOL = {"tblFakture": fixture.KOLONE["tblFakture"] + ["SEFStatus", "SEFDocumentId"]}


def faktura(t):
    """FAK-00001 (KUP-00001, 06.07) fakturise PRJ-00001 i PRJ-00002, kao CreateFaktura (fa11:290-325)."""
    t["tblFakture"].append(dict(FakturaID="FAK-00001", BrojFakture="1/2026", Datum=D6, KupacID=KUPAC,
                                Iznos=Decimal("44550"), Status="Neplaceno", Stornirano="", SEFStatus="Sent",
                                SEFDocumentId="sef-1"))
    for i, pid in enumerate(("PRJ-00001", "PRJ-00002"), 1):
        p = red(t, "tblPrijemnica", "PrijemnicaID", pid)
        t["tblFakturaStavke"].append(dict(StavkaID="FAK-00001-%02d" % i, FakturaID="FAK-00001", PrijemnicaID=pid,
                                          Kolicina=p["Kolicina"], Cena=p["Cena"], Klasa=p["Klasa"],
                                          BrojPrijemnice=p["BrojPrijemnice"], Stornirano="", OsirocenoOd=""))
        p.update(Fakturisano="Da", FakturaID="FAK-00001")


def storno_prijemnice(t, pid):
    """StornoPrijemnica (s11:576-584): prazni Fakturisano/FakturaID, linija i faktura ostaju aktivne."""
    storno_prj(t, pid)
    red(t, "tblPrijemnica", "PrijemnicaID", pid).update(Fakturisano="", FakturaID="")
    for ln in t["tblFakturaStavke"]:
        if ln["PrijemnicaID"] == pid:
            ln["OsirocenoOd"] = pid


class Fakture(Osnova):
    def test_cista_faktura(self):
        t = fixture.kopija(fixture.osnovna())
        faktura(t)
        self.assertEqual([(x.kod, x.id, x.opis) for x in nalazi(t, FAK_KOL)], [])

    def test_fakturisano_ne(self):
        """SavePrijemnica upisuje Fakturisano = 'Ne' (d28:1195): to nije 'fakturisano'."""
        t = fixture.kopija(fixture.osnovna())
        faktura(t)
        for p in t["tblPrijemnica"]:
            if not p["Fakturisano"]:
                p["Fakturisano"] = "Ne"
        self.assertEqual([(x.kod, x.id) for x in nalazi(t, FAK_KOL)], [])

    def test_stavka_tudjeg_kupca(self):
        t = fixture.kopija(fixture.osnovna())
        kupac2(t)
        faktura(t)
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00002")["KupacID"] = KUP2
        nal = nalazi(t, FAK_KOL)
        self.ima(nal, "FAK_STAVKA_TUDJEG_KUPCA", "SUMNJA", "FAK-00001-02")
        self.nema(nal, "FAK_RELINK_TUDJA_ILI_STORNIRANA", "FAK_STAVKA_ZASTARELA_POSLE_RELINKA")

    def test_relink_na_tudju_prijemnicu(self):
        t = fixture.kopija(fixture.osnovna())
        kupac2(t)
        faktura(t)
        nova_prijemnica(t, "PRJ-00002", "PRJ-00006", KupacID=KUP2, Datum=D7, Fakturisano="Da", FakturaID="FAK-00001")
        red(t, "tblFakturaStavke", "StavkaID", "FAK-00001-02")["PrijemnicaID"] = "PRJ-00006"
        nal = nalazi(t, FAK_KOL)
        x = self.ima(nal, "FAK_RELINK_TUDJA_ILI_STORNIRANA", "SUMNJA", "FAK-00001-02")
        self.assertIn("relinka", x.opis)
        self.nema(nal, "FAK_STAVKA_TUDJEG_KUPCA")

    def test_stavka_zastarela(self):
        t = fixture.kopija(fixture.osnovna())
        faktura(t)
        red(t, "tblFakturaStavke", "StavkaID", "FAK-00001-01")["Kolicina"] = 100
        nal = nalazi(t, FAK_KOL)
        self.ima(nal, "FAK_STAVKA_ZASTARELA_POSLE_RELINKA", "SUMNJA", "FAK-00001-01")
        self.ima(nal, "FAK_STAVKA_ZASTARELA_POSLE_RELINKA", "INFO", "FAK-00001")

    def test_linija_na_storniranoj(self):
        t = fixture.kopija(fixture.osnovna())
        faktura(t)
        storno_prijemnice(t, "PRJ-00002")
        nal = nalazi(t, FAK_KOL)
        self.ima(nal, "FAK_LINIJA_NA_STORNIRANOJ_PRIJEMNICI", "SUMNJA", "FAK-00001-02")
        self.nema(nal, "FAK_AKTIVNA_NA_STORNIRANOJ_PRIJEMNICI")

    def test_zamena_fakturisana_dvaput(self):
        t = fixture.kopija(fixture.osnovna())
        faktura(t)
        storno_prijemnice(t, "PRJ-00002")
        nova_prijemnica(t, "PRJ-00002", "PRJ-00006", BrojPrijemnice="1/050726-4", Stornirano="",
                        Fakturisano="Da", FakturaID="FAK-00002")
        t["tblFakture"].append(dict(red(t, "tblFakture", "FakturaID", "FAK-00001"), FakturaID="FAK-00002",
                                    BrojFakture="2/2026", Iznos=Decimal("14850")))
        t["tblFakturaStavke"].append(dict(StavkaID="FAK-00002-01", FakturaID="FAK-00002", PrijemnicaID="PRJ-00006",
                                          Kolicina=Decimal("49.5"), Cena=300, Klasa="I", BrojPrijemnice="1/050726-4",
                                          Stornirano="", OsirocenoOd=""))
        nal = nalazi(t, FAK_KOL)
        x = self.ima(nal, "FAK_AKTIVNA_NA_STORNIRANOJ_PRIJEMNICI", "SUMNJA", "FAK-00001-02")
        self.assertIn("FAK-00002", x.opis)
        self.nema(nal, "FAK_LINIJA_NA_STORNIRANOJ_PRIJEMNICI")

    def test_status_prijemnice(self):
        t = fixture.kopija(fixture.osnovna())
        faktura(t)
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00003").update(Fakturisano="Da", FakturaID="FAK-00009")
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00001")["Fakturisano"] = ""
        nal = nalazi(t, FAK_KOL)
        self.ima(nal, "FAK_STATUS_PRIJEMNICE_NESAGLASAN", "GRESKA", "PRJ-00003")
        self.ima(nal, "FAK_STATUS_PRIJEMNICE_NESAGLASAN", "SUMNJA", "PRJ-00001")

    def test_relink_ozivio_liniju_stornirane_fakture(self):
        t = fixture.kopija(fixture.osnovna())
        faktura(t)
        t["tblFakture"].append(dict(red(t, "tblFakture", "FakturaID", "FAK-00001"), FakturaID="FAK-00002",
                                    BrojFakture="2/2026", Status="Stornirano", Stornirano="Da"))
        t["tblFakturaStavke"].append(dict(StavkaID="FAK-00002-01", FakturaID="FAK-00002", PrijemnicaID="PRJ-00003",
                                          Kolicina=79, Cena=300, Klasa="I", BrojPrijemnice="1/050726-3",
                                          Stornirano="Da", OsirocenoOd=""))
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00003").update(Fakturisano="Da", FakturaID="FAK-00002")
        nal = nalazi(t, FAK_KOL)
        self.ima(nal, "FAK_RELINK_TUDJA_ILI_STORNIRANA", "GRESKA", "PRJ-00003")
        self.nema(nal, "FAK_STATUS_PRIJEMNICE_NESAGLASAN")

    def test_sef_storno(self):
        t = fixture.kopija(fixture.osnovna())
        faktura(t)
        red(t, "tblFakture", "FakturaID", "FAK-00001")["SEFStatus"] = "STORNO"
        self.ima(nalazi(t, FAK_KOL), "FAK_SEF_STORNO_BEZ_LOKALNOG", "GRESKA", "FAK-00001")

    def test_sef_storno_i_lokalno(self):
        t = fixture.kopija(fixture.osnovna())
        faktura(t)
        red(t, "tblFakture", "FakturaID", "FAK-00001").update(SEFStatus="Cancelled", Stornirano="Da",
                                                            Status="Stornirano")
        for pid in ("PRJ-00001", "PRJ-00002"):
            red(t, "tblPrijemnica", "PrijemnicaID", pid).update(Fakturisano="", FakturaID="")
        for ln in t["tblFakturaStavke"]:
            ln["Stornirano"] = "Da"
        self.assertEqual([(x.kod, x.id) for x in nalazi(t, FAK_KOL)], [])

    def test_bez_sef_kolone(self):
        t = fixture.kopija(fixture.osnovna())
        faktura(t)
        nal = nalazi(t)
        self.ima(nal, "FAK_SEF_STORNO_BEZ_LOKALNOG_PRESKOCENO", "INFO", "tblFakture.SEFStatus")
        self.assertEqual(len(nal), 1)


if __name__ == "__main__":
    unittest.main()
