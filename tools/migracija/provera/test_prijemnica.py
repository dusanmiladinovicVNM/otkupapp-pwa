"""Testovi provere prijemnice (provera/prijemnica.py).

Svaki test polazi od ciste fixture.osnovna() (nad njom modul ne sme nista da
prijavi), podmetne gresku stare aplikacije ili njen legitimni slicni obrazac, i
trazi kod na tacnom ID-ju i ozbiljnosti -- ili nula nalaza.
"""

import time
import unittest
from datetime import date, timedelta
from decimal import Decimal

from provera import fixture
from provera.pokretac import pokreni_nad_tabelama
from provera.test_zajednicko import tabele_iz
from provera.zajednicko import GRESKA, INFO, SUMNJA

D5, D6, D7 = fixture.D5, fixture.D6, date(2026, 7, 7)
KUP = fixture.KUPAC
KOL_PRJ_ISPR = fixture.KOLONE["tblPrijemnica"] + ["IspravkaOd", "ZamenjenSa", "CorrectionID"]
KOL_ZBR_ISPR = fixture.KOLONE["tblZbirna"] + ["IspravkaOd", "ZamenjenSa", "CorrectionID"]


def pokreni(t, kolone=None):
    return pokreni_nad_tabelama(tabele_iz(t, kolone), moduli=("prijemnica",))


def nalazi(ctx, kod=None):
    return [x for x in ctx.nalazi if x.modul == "prijemnica" and (kod is None or x.kod == kod)]


def stanica(t, i, hladnjaca=""):
    sid = "ST-%05d" % i
    if not any(r["StanicaID"] == sid for r in t["tblStanice"]):
        t["tblStanice"].append(dict(StanicaID=sid, Naziv="Stanica %d" % i, Mesto="Mesto", Aktivan="Da",
                                    JeHladnjaca=hladnjaca))
        t["tblVozaci"].append(dict(VozacID=sid, Ime="Stanica %d" % i, Aktivan="Da"))
    return sid


def lanac(t, otp, zbr, broj, st, dan, kg, gajbe, vozac=None, klasa="I", kupac=KUP, bz=None, storno_otp="",
          storno_zbr="", **zbr_dop):
    """Otpremnica + auto-zbirna (malina): BrojZbirne = 'S' + broj kad je vozac stanica."""
    vozac = vozac or st
    bz = bz or (("S" + broj) if vozac.startswith("ST-") else broj)
    t["tblOtpremnica"].append(dict(
        OtpremnicaID=otp, Datum=dan, StanicaID=st, VozacID=vozac, BrojOtpremnice=broj, BrojZbirne=bz,
        VrstaVoca="Malina", SortaVoca="Vilamet", Kolicina=kg, Cena=0, TipAmbalaze="12/1", KolAmbalaze=gajbe,
        Klasa=klasa, Stornirano=storno_otp, BrutoKg=kg))
    z = dict(ZbirnaID=zbr, Datum=dan, VozacID=vozac, BrojZbirne=bz, KupacID=kupac, VrstaVoca="Malina",
             SortaVoca="Vilamet", UkupnoKolicina=kg, TipAmbalaze="12/1", UkupnoAmbalaze=gajbe, Klasa=klasa,
             Stornirano=storno_zbr)
    z.update(zbr_dop)
    t["tblZbirna"].append(z)
    return bz


def prijemnica(t, prj, broj, bz, vozac, kg, gajbe, dan, klasa="I", kupac=KUP, **dop):
    r = dict(PrijemnicaID=prj, Datum=dan, KupacID=kupac, VozacID=vozac, BrojPrijemnice=broj, BrojZbirne=bz,
             VrstaVoca="Malina", SortaVoca="Vilamet", Kolicina=Decimal(str(kg)), Cena=300, TipAmbalaze="12/1",
             KolAmbalaze=gajbe, KolAmbVracena=gajbe, Klasa=klasa, Fakturisano="", FakturaID="", Stornirano="",
             BrutoKg=Decimal(str(kg)))
    r.update(dop)
    t["tblPrijemnica"].append(r)
    return r


def red(t, ime, kol, v):
    return fixture.red(t, ime, kol, v)


class Osnova(unittest.TestCase):
    def kodovi(self, ctx):
        return sorted((x.kod, x.id, x.ozbiljnost) for x in nalazi(ctx))

    def jedan(self, ctx, kod, id_, ozb):
        nn = [x for x in nalazi(ctx, kod) if x.id == id_]
        self.assertEqual(len(nn), 1, "%s na %s: %s" % (kod, id_, self.kodovi(ctx)))
        self.assertEqual(nn[0].ozbiljnost, ozb, nn[0].opis)
        return nn[0]

    def nista(self, ctx):
        self.assertEqual([(x.kod, x.id, x.opis[:160]) for x in nalazi(ctx)], [])

    def bez_koda(self, ctx, kod):
        self.assertEqual([(x.id, x.opis[:160]) for x in nalazi(ctx, kod)], [])


# ------------------------------------------------------------------------------------------
class Cista(Osnova):
    def test_cista_sveska_nula_nalaza(self):
        self.nista(pokreni(fixture.osnovna()))

    def test_ascii_izvor(self):
        import os
        p = os.path.join(os.path.dirname(__file__), "prijemnica.py")
        with open(p, "rb") as f:
            f.read().decode("ascii")

    def test_prazna_tabela_prijemnica(self):
        t = fixture.osnovna()
        t["tblPrijemnica"] = []
        self.nista(pokreni(t))

    def test_bez_tabele_zbirna_ne_pada(self):
        t = fixture.osnovna()
        del t["tblZbirna"]
        ctx = pokreni(t)
        self.assertEqual([x for x in ctx.nalazi if x.kod == "MODUL_PAO"], [])
        self.nista(ctx)

    def test_kolona_fali_info_preskoceno(self):
        t = fixture.osnovna()
        kol = [c for c in fixture.KOLONE["tblPrijemnica"] if c != "KolAmbalaze"]
        ctx = pokreni(t, {"tblPrijemnica": kol})
        self.assertEqual([x for x in ctx.nalazi if x.kod == "MODUL_PAO"], [])
        for kod in ("PRJ_KG_GAJBE_VAN_OPSEGA_PRESKOCENO", "PRJ_GAJBE_NULA_ILI_BEZ_TIPA_PRESKOCENO"):
            nn = nalazi(ctx, kod)
            self.assertEqual(len(nn), 1, kod)
            self.assertEqual(nn[0].ozbiljnost, INFO)
            self.assertIn("KolAmbalaze", nn[0].opis)
        self.assertEqual({x.kod for x in nalazi(ctx)},
                         {"PRJ_KG_GAJBE_VAN_OPSEGA_PRESKOCENO", "PRJ_GAJBE_NULA_ILI_BEZ_TIPA_PRESKOCENO"})

    def test_smece_u_celijama_ne_pada(self):
        t = fixture.osnovna()
        r = red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")
        r.update(Kolicina="abc", Datum="", BrojZbirne="", KolAmbalaze="x", Cena="?", Klasa="", BrojPrijemnice="")
        red(t, "tblZbirna", "ZbirnaID", "ZBR-00001").update(UkupnoKolicina="n/a", Datum="juce")
        ctx = pokreni(t)
        self.assertEqual([x.opis for x in ctx.nalazi if x.kod == "MODUL_PAO"], [])


# ------------------------------------------------------------------------------------------
class DveNaZbirnojIstaStanica(Osnova):
    """Varijanta (b): broj zbirne nasledjen od prethodne prijemnice iste stanice."""

    def test_druga_tura_istog_dana(self):
        t = fixture.osnovna()
        lanac(t, "OTP-00006", "ZBR-00006", "3/060726-2", "ST-00003", D6, 60, 24)
        prijemnica(t, "PRJ-00006", "1/060726-3", "S3/060726", "ST-00003", "59.5", 24, D6)
        ctx = pokreni(t)
        x = self.jedan(ctx, "PRJ_DVE_NA_ZBIRNOJ_ISTA_STANICA", "PRJ-00006", GRESKA)
        self.assertIn("BrojZbirne: S3/060726 -> S3/060726-2; kg 59.5 vs zbirna 60, gajbe 24 = 24", x.predlog)
        self.assertIn("NASLEDJEN od prethodne prijemnice PRJ-00005", x.opis)
        z = self.jedan(ctx, "ZBR_ZIVA_BEZ_PRIJEMNICE", "ZBR-00006", SUMNJA)
        self.assertIn("PRJ-00006", z.opis)
        self.assertEqual(self.kodovi(ctx), [("PRJ_DVE_NA_ZBIRNOJ_ISTA_STANICA", "PRJ-00006", GRESKA),
                                            ("ZBR_ZIVA_BEZ_PRIJEMNICE", "ZBR-00006", SUMNJA)])

    def test_primer_s5_180626(self):
        """Rucni nalaz: S5/180626 umesto S5/180626-2."""
        t = fixture.osnovna()
        st = stanica(t, 5)
        d = date(2026, 6, 18)
        lanac(t, "OTP-00006", "ZBR-00006", "5/180626", st, d, 100, 40)
        lanac(t, "OTP-00007", "ZBR-00007", "5/180626-2", st, d, 121, 50)
        prijemnica(t, "PRJ-00006", "1/180626", "S5/180626", st, 99, 40, d)
        prijemnica(t, "PRJ-00007", "1/180626-2", "S5/180626", st, "120.5", 50, d)
        ctx = pokreni(t)
        x = self.jedan(ctx, "PRJ_DVE_NA_ZBIRNOJ_ISTA_STANICA", "PRJ-00007", GRESKA)
        self.assertIn("BrojZbirne: S5/180626 -> S5/180626-2; kg 120.5 vs zbirna 121, gajbe 50 = 50", x.predlog)
        self.jedan(ctx, "ZBR_ZIVA_BEZ_PRIJEMNICE", "ZBR-00007", SUMNJA)
        self.bez_koda(ctx, "PRJ_KG_GAJBE_VAN_OPSEGA")

    def test_primer_sledeci_dan_s4(self):
        """Rucni nalaz: S4/200626-2 umesto S4/210626 (sledeci dan)."""
        t = fixture.osnovna()
        st = stanica(t, 4)
        d20, d21 = date(2026, 6, 20), date(2026, 6, 21)
        lanac(t, "OTP-00006", "ZBR-00006", "4/200626", st, d20, 80, 32)
        lanac(t, "OTP-00007", "ZBR-00007", "4/200626-2", st, d20, 50, 20)
        lanac(t, "OTP-00008", "ZBR-00008", "4/210626", st, d21, 140, 56)
        prijemnica(t, "PRJ-00006", "1/200626", "S4/200626", st, 79, 32, d20)
        prijemnica(t, "PRJ-00007", "1/200626-2", "S4/200626-2", st, "49.5", 20, d20)
        prijemnica(t, "PRJ-00008", "1/210626", "S4/200626-2", st, 139, 56, d21)
        ctx = pokreni(t)
        x = self.jedan(ctx, "PRJ_DVE_NA_ZBIRNOJ_ISTA_STANICA", "PRJ-00008", GRESKA)
        self.assertIn("S4/200626-2 -> S4/210626", x.predlog)
        self.assertEqual(self.kodovi(ctx), [("PRJ_DVE_NA_ZBIRNOJ_ISTA_STANICA", "PRJ-00008", GRESKA),
                                            ("ZBR_ZIVA_BEZ_PRIJEMNICE", "ZBR-00008", SUMNJA)])

    def test_negativ_vise_tura_istog_dana(self):
        """Legitimno: druga tura '-2' sa svojom prijemnicom."""
        t = fixture.osnovna()
        lanac(t, "OTP-00006", "ZBR-00006", "3/060726-2", "ST-00003", D6, 60, 24)
        lanac(t, "OTP-00007", "ZBR-00007", "3/060726-3", "ST-00003", D6, 60, 24)
        prijemnica(t, "PRJ-00006", "1/060726-3", "S3/060726-2", "ST-00003", "59.5", 24, D6)
        prijemnica(t, "PRJ-00007", "1/060726-4", "S3/060726-3", "ST-00003", "59.6", 24, D6)
        self.nista(pokreni(t))

    def test_deljenje_prijema_samo_info(self):
        """Hladnjaca izdala dve prijemnice za jedan kamion: zbir odgovara zbirnoj."""
        t = fixture.osnovna()
        r = red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")
        r.update(Kolicina=Decimal(50), KolAmbalaze=20, BrutoKg=Decimal(50))
        prijemnica(t, "PRJ-00006", "1/060726-3", "S3/060726", "ST-00003", 39, 16, D6)
        ctx = pokreni(t)
        self.assertEqual(self.kodovi(ctx), [("PRJ_DVE_NA_ZBIRNOJ_ISTA_STANICA", "PRJ-00006", INFO)])


# ------------------------------------------------------------------------------------------
class ZbirnaDrugeStanice(Osnova):
    """Varijanta (a): prijemnica stanice X na zbirnoj stanice Y."""

    def test_primer_s3_030726(self):
        t = fixture.osnovna()
        d3 = date(2026, 7, 3)
        lanac(t, "OTP-00006", "ZBR-00006", "3/030726", "ST-00003", d3, 90, 36)
        lanac(t, "OTP-00007", "ZBR-00007", "2/030726", "ST-00002", d3, 60, 24)
        prijemnica(t, "PRJ-00006", "1/030726", "S3/030726", "ST-00003", 89, 36, d3)
        prijemnica(t, "PRJ-00007", "1/030726-2", "S3/030726", "ST-00002", "59.4", 24, d3)
        ctx = pokreni(t)
        x = self.jedan(ctx, "PRJ_ZBIRNA_DRUGE_STANICE", "PRJ-00007", GRESKA)
        self.assertIn("BrojZbirne: S3/030726 -> S2/030726; kg 59.4 vs zbirna 60, gajbe 24 = 24", x.predlog)
        self.assertIn("NASLEDJEN", x.opis)
        self.jedan(ctx, "ZBR_ZIVA_BEZ_PRIJEMNICE", "ZBR-00007", SUMNJA)
        self.assertEqual(len(nalazi(ctx)), 2, self.kodovi(ctx))

    def test_negativ_otpremnica_sa_vozacem_druge_stanice(self):
        """Prefiks S2 a VozacID zbirne ST-00003: legitimno kad prijemnica nosi isti VozacID kao zbirna."""
        t = fixture.osnovna()
        lanac(t, "OTP-00006", "ZBR-00006", "2/060726", "ST-00002", D6, 60, 24, vozac="ST-00003")
        prijemnica(t, "PRJ-00006", "1/060726-3", "S2/060726", "ST-00003", "59.5", 24, D6)
        self.nista(pokreni(t))

    def test_negativ_realni_vozac(self):
        """Rucna zbirna realnog vozaca (broj bez S)."""
        t = fixture.osnovna()
        lanac(t, "OTP-00006", "ZBR-00006", "2/060726", "ST-00002", D6, 60, 24, vozac="VOZ-00001", bz="1/060726")
        prijemnica(t, "PRJ-00006", "1/060726-3", "1/060726", "VOZ-00001", "59.5", 24, D6)
        self.nista(pokreni(t))


# ------------------------------------------------------------------------------------------
class ZastareoVozac(Osnova):
    def test_vozac_zaostao(self):
        t = fixture.osnovna()
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")["VozacID"] = "ST-00002"
        ctx = pokreni(t)
        x = self.jedan(ctx, "PRJ_ZASTAREO_VOZAC", "PRJ-00005", GRESKA)
        self.assertIn("VozacID: ST-00002 -> ST-00003", x.predlog)
        self.assertIn("tblAmbalaza", x.predlog)
        self.assertEqual(len(nalazi(ctx)), 1, self.kodovi(ctx))

    def test_vozac_je_stanica_iz_broja_samo_sumnja(self):
        """Vozac prijemnice = stanica iz broja zbirne, a zbirna nosi vozaca druge stanice: ODLUKA."""
        t = fixture.osnovna()
        lanac(t, "OTP-00006", "ZBR-00006", "2/060726", "ST-00002", D6, 60, 24, vozac="ST-00003")
        prijemnica(t, "PRJ-00006", "1/060726-3", "S2/060726", "ST-00002", "59.5", 24, D6)
        ctx = pokreni(t)
        x = self.jedan(ctx, "PRJ_ZASTAREO_VOZAC", "PRJ-00006", SUMNJA)
        self.assertIn("ODLUKA", x.predlog)
        self.bez_koda(ctx, "PRJ_ZBIRNA_DRUGE_STANICE")

    def test_negativ_isti_vozac_kao_zbirna(self):
        t = fixture.osnovna()
        lanac(t, "OTP-00006", "ZBR-00006", "2/060726", "ST-00002", D6, 60, 24, vozac="ST-00003")
        prijemnica(t, "PRJ-00006", "1/060726-3", "S2/060726", "ST-00003", "59.5", 24, D6)
        self.bez_koda(pokreni(t), "PRJ_ZASTAREO_VOZAC")


# ------------------------------------------------------------------------------------------
class Kupac(Osnova):
    def test_kupac_zaostao(self):
        t = fixture.osnovna()
        t["tblKupci"].append(dict(KupacID="KUP-00002", Naziv="Drugi kupac", Aktivan="Da"))
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")["KupacID"] = "KUP-00002"
        ctx = pokreni(t)
        x = self.jedan(ctx, "PRJ_KUPAC_RAZLICIT_OD_ZBIRNE", "PRJ-00005", SUMNJA)
        self.assertIn("KupacID: KUP-00002 -> KUP-00001", x.predlog)
        self.assertEqual(len(nalazi(ctx)), 1, self.kodovi(ctx))

    def test_negativ_eksterni_kupac_sa_svojom_zbirnom(self):
        t = fixture.osnovna()
        t["tblKupci"].append(dict(KupacID="KUP-00002", Naziv="Drugi kupac", Aktivan="Da"))
        lanac(t, "OTP-00006", "ZBR-00006", "2/060726", "ST-00002", D6, 60, 24, kupac="KUP-00002")
        prijemnica(t, "PRJ-00006", "77/2026", "S2/060726", "ST-00002", "59.5", 24, D6, kupac="KUP-00002")
        self.nista(pokreni(t))


# ------------------------------------------------------------------------------------------
class Datum(Osnova):
    def test_prijem_pre_zbirne(self):
        t = fixture.osnovna()
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005").update(Datum=D5, BrojPrijemnice="1/050726-4")
        ctx = pokreni(t)
        x = self.jedan(ctx, "PRJ_DATUM_ZASTAREO_ILI_PRE_ZBIRNE", "PRJ-00005", SUMNJA)
        self.assertIn("PRE datuma zbirne", x.opis)

    def test_kasni_vise_od_praga(self):
        t = fixture.osnovna()
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005").update(Datum=D6 + timedelta(days=10),
                                                                     BrojPrijemnice="1/160726")
        x = self.jedan(pokreni(t), "PRJ_DATUM_ZASTAREO_ILI_PRE_ZBIRNE", "PRJ-00005", SUMNJA)
        self.assertIn("kasni 10 dana", x.opis)

    def test_datum_u_broju_hladnjace(self):
        t = fixture.osnovna()
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")["BrojPrijemnice"] = "1/070726"
        x = self.jedan(pokreni(t), "PRJ_DATUM_ZASTAREO_ILI_PRE_ZBIRNE", "PRJ-00005", SUMNJA)
        self.assertIn("datum u broju 1/070726", x.opis)

    def test_negativ_prijemnica_dan_posle_zbirne(self):
        t = fixture.osnovna()
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005").update(Datum=D7, BrojPrijemnice="1/070726")
        self.nista(pokreni(t))


# ------------------------------------------------------------------------------------------
class Roba(Osnova):
    def test_sorta_zaostala(self):
        t = fixture.osnovna()
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")["SortaVoca"] = "Polka"
        x = self.jedan(pokreni(t), "PRJ_ROBA_RAZLICITA_OD_ZBIRNE", "PRJ-00005", SUMNJA)
        self.assertIn("SortaVoca: Polka -> Vilamet", x.predlog)

    def test_vrsta_razlicita_greska(self):
        t = fixture.osnovna()
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")["VrstaVoca"] = "Kupina"
        self.jedan(pokreni(t), "PRJ_ROBA_RAZLICITA_OD_ZBIRNE", "PRJ-00005", GRESKA)

    def test_negativ_samo_velika_slova(self):
        t = fixture.osnovna()
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005").update(VrstaVoca="MALINA ", SortaVoca="vilamet")
        self.nista(pokreni(t))


# ------------------------------------------------------------------------------------------
class Klasa(Osnova):
    def test_klasa_ii_na_zbirnoj_klase_i(self):
        t = fixture.osnovna()
        lanac(t, "OTP-00006", "ZBR-00006", "3/060726-2", "ST-00003", D6, 30, 12, klasa="II")
        prijemnica(t, "PRJ-00006", "1/060726-3", "S3/060726", "ST-00003", "29.5", 12, D6, klasa="II")
        # PRJ-00006 (Klasa II) nasledila S3/060726: dve na zbirnoj, druga je pogresne klase
        ctx = pokreni(t)
        x = self.jedan(ctx, "PRJ_DVE_NA_ZBIRNOJ_ISTA_STANICA", "PRJ-00006", GRESKA)
        self.assertIn("KLASA_II_NA_ZBIRNOJ_KLASE_I", x.opis)
        self.assertIn("S3/060726 -> S3/060726-2", x.predlog)

    def test_klasa_jedina_prijemnica(self):
        t = fixture.osnovna()
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")["Klasa"] = "II"
        lanac(t, "OTP-00006", "ZBR-00006", "3/060726-2", "ST-00003", D6, 90, 36, klasa="II")
        ctx = pokreni(t)
        x = self.jedan(ctx, "PRJ_KLASA_RAZLICITA_OD_ZBIRNE", "PRJ-00005", SUMNJA)
        self.assertIn("S3/060726 -> S3/060726-2", x.predlog)
        self.jedan(ctx, "ZBR_ZIVA_BEZ_PRIJEMNICE", "ZBR-00006", INFO)
        self.bez_koda(ctx, "PRJ_KG_GAJBE_VAN_OPSEGA")

    def test_negativ_pwa_auto_otpremnica_po_klasi(self):
        """PWA: Klasa II iste ture dobija svoj broj (-2), svoju zbirnu i svoju prijemnicu."""
        t = fixture.osnovna()
        lanac(t, "OTP-00006", "ZBR-00006", "3/060726-2", "ST-00003", D6, 30, 12, klasa="II")
        prijemnica(t, "PRJ-00006", "1/060726-3", "S3/060726-2", "ST-00003", "29.5", 12, D6, klasa="II")
        self.nista(pokreni(t))


# ------------------------------------------------------------------------------------------
class ZbirnaNePostojiIliStornirana(Osnova):
    def test_tipfeler(self):
        t = fixture.osnovna()
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")["BrojZbirne"] = "S3/06O726"
        x = self.jedan(pokreni(t), "PRJ_ZBIRNA_NE_POSTOJI_ILI_STORNIRANA", "PRJ-00005", GRESKA)
        self.assertIn("TIPFELER", x.opis)

    def test_ne_postoji_sa_kandidatom(self):
        t = fixture.osnovna()
        lanac(t, "OTP-00006", "ZBR-00006", "3/060726-2", "ST-00003", D6, 60, 24)
        prijemnica(t, "PRJ-00006", "1/060726-3", "S3/060726-5", "ST-00003", "59.5", 24, D6)
        ctx = pokreni(t)
        x = self.jedan(ctx, "PRJ_ZBIRNA_NE_POSTOJI_ILI_STORNIRANA", "PRJ-00006", GRESKA)
        self.assertIn("S3/060726-5 -> S3/060726-2", x.predlog)

    def test_lanac_storniran(self):
        """(P1) storno otpremnice kaskadira na zbirnu, ne na prijemnicu."""
        t = fixture.osnovna()
        red(t, "tblOtpremnica", "OtpremnicaID", "OTP-00005")["Stornirano"] = "Da"
        red(t, "tblZbirna", "ZbirnaID", "ZBR-00005")["Stornirano"] = "Da"
        lanac(t, "OTP-00006", "ZBR-00006", "3/060726-2", "ST-00003", D6, 90, 36)
        ctx = pokreni(t)
        x = self.jedan(ctx, "PRJ_ZBIRNA_NE_POSTOJI_ILI_STORNIRANA", "PRJ-00005", GRESKA)
        self.assertIn("LANAC_STORNIRAN", x.opis)
        self.assertIn("S3/060726 -> S3/060726-2", x.predlog)
        self.bez_koda(ctx, "PRJ_AKTIVNA_NA_STORNIRANOJ_ZBIRNOJ")

    def test_dvosmislena(self):
        t = fixture.osnovna()
        z = dict(red(t, "tblZbirna", "ZbirnaID", "ZBR-00005"))
        z["ZbirnaID"] = "ZBR-00006"
        t["tblZbirna"].append(z)
        self.jedan(pokreni(t), "PRJ_ZBIRNA_NE_POSTOJI_ILI_STORNIRANA", "PRJ-00005", SUMNJA)

    def test_negativ_storno_pa_ponovni_unos_istog_broja(self):
        t = fixture.osnovna()
        red(t, "tblOtpremnica", "OtpremnicaID", "OTP-00005")["Stornirano"] = "Da"
        red(t, "tblZbirna", "ZbirnaID", "ZBR-00005")["Stornirano"] = "Da"
        lanac(t, "OTP-00006", "ZBR-00006", "3/060726", "ST-00003", D6, 90, 36)
        self.nista(pokreni(t))

    def test_negativ_storno_ispravka_bez_kaskade_na_zbirnu(self):
        """Stara zbirna ostaje aktivna (ne kaskadira), nova istog broja je napravljena posle."""
        t = fixture.osnovna()
        red(t, "tblOtpremnica", "OtpremnicaID", "OTP-00005")["Stornirano"] = "Da"
        lanac(t, "OTP-00006", "ZBR-00006", "3/060726", "ST-00003", D6, 90, 36)
        self.nista(pokreni(t))


# ------------------------------------------------------------------------------------------
class AktivnaNaStorniranojZbirnoj(Osnova):
    def test_zbirna_stornirana_otpremnica_ziva(self):
        t = fixture.osnovna()
        red(t, "tblZbirna", "ZbirnaID", "ZBR-00005")["Stornirano"] = "Da"
        x = self.jedan(pokreni(t), "PRJ_AKTIVNA_NA_STORNIRANOJ_ZBIRNOJ", "PRJ-00005", GRESKA)
        self.assertIn("ZBR-00005", x.opis)
        self.assertIn("OTP-00005", x.opis)

    def test_naslednik_iz_zamenjen_sa(self):
        t = fixture.osnovna()
        red(t, "tblZbirna", "ZbirnaID", "ZBR-00005").update(Stornirano="Da", ZamenjenSa="S3/060726-2")
        lanac(t, "OTP-00006", "ZBR-00006", "3/060726-2", "ST-00003", D6, 90, 36)
        ctx = pokreni(t, {"tblZbirna": KOL_ZBR_ISPR})
        x = self.jedan(ctx, "PRJ_AKTIVNA_NA_STORNIRANOJ_ZBIRNOJ", "PRJ-00005", GRESKA)
        self.assertIn("S3/060726 -> S3/060726-2 (naslednik: ZamenjenSa", x.predlog)
        z = self.jedan(ctx, "ZBR_ZIVA_BEZ_PRIJEMNICE", "ZBR-00006", SUMNJA)
        self.assertIn("PRJ-00005", z.opis)

    def test_negativ_stornirana_prijemnica_na_storniranoj_zbirnoj(self):
        t = fixture.osnovna()
        red(t, "tblZbirna", "ZbirnaID", "ZBR-00005")["Stornirano"] = "Da"
        red(t, "tblOtpremnica", "OtpremnicaID", "OTP-00005")["Stornirano"] = "Da"
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")["Stornirano"] = "Da"
        self.nista(pokreni(t))


# ------------------------------------------------------------------------------------------
class BrojNenormalizovan(Osnova):
    def test_razmak_i_mala_slova(self):
        t = fixture.osnovna()
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")["BrojZbirne"] = " s3/060726"
        ctx = pokreni(t)
        x = self.jedan(ctx, "PRJ_BROJ_ZBIRNE_NENORMALIZOVAN", "PRJ-00005", GRESKA)
        self.assertIn("-> 'S3/060726'", x.predlog)
        self.assertEqual(len(nalazi(ctx)), 1, self.kodovi(ctx))

    def test_negativ_zbirna_nenormalizovana_a_prijemnica_tacna(self):
        t = fixture.osnovna()
        red(t, "tblZbirna", "ZbirnaID", "ZBR-00005")["BrojZbirne"] = "s3/060726"
        self.nista(pokreni(t))


# ------------------------------------------------------------------------------------------
class DupliBroj(Osnova):
    def test_dupli_broj_hladnjace(self):
        t = fixture.osnovna()
        lanac(t, "OTP-00006", "ZBR-00006", "2/060726", "ST-00002", D6, 60, 24)
        prijemnica(t, "PRJ-00006", "1/060726-2", "S2/060726", "ST-00002", "59.5", 24, D6)
        x = self.jedan(pokreni(t), "PRJ_DUPLI_BROJ_PRIJEMNICE", "PRJ-00006", GRESKA)
        self.assertIn("BrojPrijemnice: 1/060726-2 -> 1/060726-3", x.predlog)

    def test_negativ_storno_pa_ponovni_unos_istog_broja(self):
        t = fixture.osnovna()
        r = red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")
        n = dict(r)
        r["Stornirano"] = "Da"
        n["PrijemnicaID"] = "PRJ-00006"
        t["tblPrijemnica"].append(n)
        self.nista(pokreni(t))

    def test_negativ_par_klasa_i_ii(self):
        """VBA forma: Klasa I i II isti broj otpremnice, zbirne i prijemnice -- jedan dokument."""
        t = fixture.osnovna()
        lanac(t, "OTP-00006", "ZBR-00006", "3/060726", "ST-00003", D6, 30, 12, klasa="II")
        prijemnica(t, "PRJ-00006", "1/060726-2", "S3/060726", "ST-00003", "29.7", 12, D6, klasa="II")
        self.nista(pokreni(t))


# ------------------------------------------------------------------------------------------
class CenaNula(Osnova):
    def test_cena_nula(self):
        t = fixture.osnovna()
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")["Cena"] = 0
        self.jedan(pokreni(t), "PRJ_CENA_NULA", "PRJ-00005", SUMNJA)

    def test_dokument_sa_dve_klase_jednom(self):
        t = fixture.osnovna()
        lanac(t, "OTP-00006", "ZBR-00006", "3/060726", "ST-00003", D6, 30, 12, klasa="II")
        prijemnica(t, "PRJ-00006", "1/060726-2", "S3/060726", "ST-00003", "29.7", 12, D6, klasa="II", Cena="")
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")["Cena"] = 0
        ctx = pokreni(t)
        self.jedan(ctx, "PRJ_CENA_NULA", "PRJ-00005, PRJ-00006", SUMNJA)
        self.assertEqual(len(nalazi(ctx)), 1, self.kodovi(ctx))

    def test_praksa_bez_cene_jedan_info(self):
        """Vecina prijemnica bez cene = praksa (cena na fakturi): jedan INFO, ne poplava SUMNJI."""
        t = fixture.osnovna()
        for r in t["tblPrijemnica"]:
            r["Cena"] = 0
        for j in range(25):
            dan = date(2026, 5, 1) + timedelta(days=j)
            bz = lanac(t, "OTP-%05d" % (100 + j), "ZBR-%05d" % (100 + j), "1/%s" % dan.strftime("%d%m%y"),
                       "ST-00001", dan, 50, 20)
            prijemnica(t, "PRJ-%05d" % (100 + j), "1/%s" % dan.strftime("%d%m%y"), bz, "ST-00001", "49.5", 20, dan,
                       Cena=0)
        ctx = pokreni(t)
        nn = nalazi(ctx, "PRJ_CENA_NULA")
        self.assertEqual([x.ozbiljnost for x in nn], [INFO])
        self.assertIn("30 od 30", nn[0].opis)
        self.assertEqual(len(nalazi(ctx)), 1, self.kodovi(ctx))

    def test_negativ_stornirana_bez_cene(self):
        t = fixture.osnovna()
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005").update(Cena="", Stornirano="Da")
        self.bez_koda(pokreni(t), "PRJ_CENA_NULA")


# ------------------------------------------------------------------------------------------
class KgGajbe(Osnova):
    def test_tipfeler_decimalni_zarez(self):
        t = fixture.osnovna()
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")["Kolicina"] = Decimal("8.9")
        x = self.jedan(pokreni(t), "PRJ_KG_GAJBE_VAN_OPSEGA", "PRJ-00005", GRESKA)
        self.assertIn("Kolicina: 8.9 -> 89", x.predlog)

    def test_gajbe_razlicite(self):
        t = fixture.osnovna()
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")["KolAmbalaze"] = 30
        x = self.jedan(pokreni(t), "PRJ_KG_GAJBE_VAN_OPSEGA", "PRJ-00005", SUMNJA)
        self.assertIn("gajbe 30 vs zbirna 36", x.opis)

    def test_negativ_kalo_u_pragu(self):
        t = fixture.osnovna()
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")["Kolicina"] = Decimal("87.3")
        self.nista(pokreni(t))


# ------------------------------------------------------------------------------------------
class GajbeNula(Osnova):
    def test_gajbe_nula(self):
        t = fixture.osnovna()
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")["KolAmbalaze"] = 0
        ctx = pokreni(t)
        x = self.jedan(ctx, "PRJ_GAJBE_NULA_ILI_BEZ_TIPA", "PRJ-00005", SUMNJA)
        self.assertIn("KolAmbalaze: 0 -> 36", x.predlog)
        self.bez_koda(ctx, "PRJ_KG_GAJBE_VAN_OPSEGA")

    def test_bez_tipa(self):
        t = fixture.osnovna()
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")["TipAmbalaze"] = ""
        x = self.jedan(pokreni(t), "PRJ_GAJBE_NULA_ILI_BEZ_TIPA", "PRJ-00005", SUMNJA)
        self.assertIn("bez TipAmbalaze", x.opis)

    def test_negativ_zbirna_bez_gajbi(self):
        t = fixture.osnovna()
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")["KolAmbalaze"] = 0
        red(t, "tblZbirna", "ZbirnaID", "ZBR-00005")["UkupnoAmbalaze"] = 0
        self.nista(pokreni(t))


# ------------------------------------------------------------------------------------------
class PrevezanaBezProvere(Osnova):
    def test_paletna_stavka_nosi_drugu_zbirnu(self):
        t = fixture.osnovna()
        red(t, "tblPaletaStavka", "StavkaID", "PLS-00001")["BrojZbirne"] = "S2/050726"
        x = self.jedan(pokreni(t), "PRJ_PREVEZANA_BEZ_PROVERE", "PRJ-00001", SUMNJA)
        self.assertIn("Ostaviti S1/050726", x.predlog)

    def test_prevezano_zajedno_preko_kupaca(self):
        t = fixture.osnovna()
        t["tblKupci"].append(dict(KupacID="KUP-00002", Naziv="Drugi kupac", Aktivan="Da"))
        prijemnica(t, "PRJ-00006", "1/060726-2", "S3/060726", "ST-00003", 40, 16, D6, kupac="KUP-00002")
        x = self.jedan(pokreni(t), "PRJ_PREVEZANA_BEZ_PROVERE", "PRJ-00006", SUMNJA)
        self.assertIn("PREVEZANO_ZAJEDNO", x.opis)

    def test_negativ_stavka_bez_broja_zbirne(self):
        t = fixture.osnovna()
        red(t, "tblPaletaStavka", "StavkaID", "PLS-00001")["BrojZbirne"] = ""
        self.nista(pokreni(t))


# ------------------------------------------------------------------------------------------
class IspravkaNaDrugojZbirnoj(Osnova):
    def test_pogresna_zamena(self):
        t = fixture.osnovna()
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")["Stornirano"] = "Da"
        lanac(t, "OTP-00006", "ZBR-00006", "2/060726", "ST-00002", D6, 60, 24)
        prijemnica(t, "PRJ-00006", "1/060726-3", "S2/060726", "ST-00002", "59.5", 24, D6, IspravkaOd="1/060726-2")
        ctx = pokreni(t, {"tblPrijemnica": KOL_PRJ_ISPR})
        x = self.jedan(ctx, "PRJ_ISPRAVKA_NA_DRUGOJ_ZBIRNOJ", "PRJ-00006", SUMNJA)
        self.assertIn("POGRESNA_ZAMENA", x.opis)

    def test_negativ_legitimna_popravka_pogresne_zbirne(self):
        t = fixture.osnovna()
        lanac(t, "OTP-00006", "ZBR-00006", "3/060726-2", "ST-00003", D6, 60, 24)
        prijemnica(t, "PRJ-00006", "1/060726-3", "S3/060726", "ST-00003", "59.5", 24, D6, Stornirano="Da")
        prijemnica(t, "PRJ-00007", "1/060726-4", "S3/060726-2", "ST-00003", "59.5", 24, D6,
                   IspravkaOd="1/060726-3")
        self.nista(pokreni(t, {"tblPrijemnica": KOL_PRJ_ISPR}))


# ------------------------------------------------------------------------------------------
class AutoLanacNepotpun(Osnova):
    def test_lanac_bez_prijemnice(self):
        t = fixture.osnovna()
        red(t, "tblStanice", "StanicaID", "ST-00003")["JeHladnjaca"] = "Da"
        t["tblPrijemnica"] = [r for r in t["tblPrijemnica"] if r["PrijemnicaID"] != "PRJ-00005"]
        ctx = pokreni(t)
        self.jedan(ctx, "PRJ_AUTO_LANAC_NEPOTPUN", "OTP-00005", SUMNJA)
        self.bez_koda(ctx, "ZBR_ZIVA_BEZ_PRIJEMNICE")

    def test_negativ_lanac_kompletan(self):
        t = fixture.osnovna()
        red(t, "tblStanice", "StanicaID", "ST-00003")["JeHladnjaca"] = "Da"
        self.nista(pokreni(t))


# ------------------------------------------------------------------------------------------
class ZivaBezPrijemnice(Osnova):
    def test_ziva_zbirna_bez_prijemnice(self):
        t = fixture.osnovna()
        lanac(t, "OTP-00006", "ZBR-00006", "2/060726", "ST-00002", D6, 60, 24)
        x = self.jedan(pokreni(t), "ZBR_ZIVA_BEZ_PRIJEMNICE", "ZBR-00006", SUMNJA)
        self.assertIn("OTP-00006", x.opis)

    def test_negativ_zombi_zbirna(self):
        """Aktivna zbirna cije su otpremnice stornirane nije 'ziva' (zombi: drugi modul)."""
        t = fixture.osnovna()
        lanac(t, "OTP-00006", "ZBR-00006", "2/060726", "ST-00002", D6, 60, 24, storno_otp="Da")
        self.nista(pokreni(t))


# ------------------------------------------------------------------------------------------
class NasledjenaZbirna(Osnova):
    def test_nasledjena_posle_storna_prethodne(self):
        t = fixture.osnovna()
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")["Stornirano"] = "Da"
        lanac(t, "OTP-00006", "ZBR-00006", "3/060726-2", "ST-00003", D6, 60, 24)
        prijemnica(t, "PRJ-00006", "1/060726-3", "S3/060726", "ST-00003", "59.5", 24, D6)
        ctx = pokreni(t)
        x = self.jedan(ctx, "PRJ_NASLEDJENA_ZBIRNA", "PRJ-00006", SUMNJA)
        self.assertIn("BrojZbirne: S3/060726 -> S3/060726-2", x.predlog)
        self.bez_koda(ctx, "PRJ_KG_GAJBE_VAN_OPSEGA")

    def test_negativ_storno_pa_ispravka_iste_robe(self):
        t = fixture.osnovna()
        r = red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")
        r["Stornirano"] = "Da"
        prijemnica(t, "PRJ-00006", "1/060726-3", "S3/060726", "ST-00003", 89, 36, D6)
        self.nista(pokreni(t))


# ------------------------------------------------------------------------------------------
class PokrivaDveTure(Osnova):
    def test_jedna_prijemnica_za_dve_ture(self):
        t = fixture.osnovna()
        lanac(t, "OTP-00006", "ZBR-00006", "3/060726-2", "ST-00003", D6, 60, 24)
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005").update(Kolicina=Decimal(149), KolAmbalaze=60)
        ctx = pokreni(t)
        x = self.jedan(ctx, "PRIJEMNICA_POKRIVA_DVE_TURE", "PRJ-00005", SUMNJA)
        self.assertIn("S3/060726-2", x.opis)
        self.bez_koda(ctx, "ZBR_ZIVA_BEZ_PRIJEMNICE")
        self.bez_koda(ctx, "PRJ_KG_GAJBE_VAN_OPSEGA")

    def test_negativ_svaka_tura_svoja_prijemnica(self):
        t = fixture.osnovna()
        lanac(t, "OTP-00006", "ZBR-00006", "3/060726-2", "ST-00003", D6, 60, 24)
        prijemnica(t, "PRJ-00006", "1/060726-3", "S3/060726-2", "ST-00003", "59.5", 24, D6)
        self.nista(pokreni(t))


# ------------------------------------------------------------------------------------------
class TestniRedovi(Osnova):
    def test_redovi_testne_suite(self):
        t = fixture.osnovna()
        prijemnica(t, "PRJ-00006", "TST-PRO-PRJ-1", "TST-PRO-ZBR-1", "VOZ-90001", 10, 4, D6, kupac="KUP-90001",
                   VrstaVoca="Test Jabuka")
        ctx = pokreni(t)
        self.jedan(ctx, "TESTNI_REDOVI_U_PRODUKCIJI", "PRJ-00006", SUMNJA)
        self.assertEqual(len(nalazi(ctx)), 1, self.kodovi(ctx))

    def test_negativ_realni_vozac_nije_test(self):
        t = fixture.osnovna()
        lanac(t, "OTP-00006", "ZBR-00006", "2/060726", "ST-00002", D6, 60, 24, vozac="VOZ-00001", bz="1/060726")
        prijemnica(t, "PRJ-00006", "1/060726-3", "1/060726", "VOZ-00001", "59.5", 24, D6)
        self.bez_koda(pokreni(t), "TESTNI_REDOVI_U_PRODUKCIJI")


# ------------------------------------------------------------------------------------------
class Recenzija(Osnova):
    """Slucajevi koje je nasla nezavisna recenzija (lazni pozitivi, propustene varijante, ozbiljnost)."""

    def test_ascii_testovi(self):
        import os
        with open(os.path.join(os.path.dirname(__file__), "test_prijemnica.py"), "rb") as f:
            f.read().decode("ascii")

    def test_negativ_jedan_dokument_klasa_i_ii_a_pwa_zbirna_ii(self):
        """Hladnjaca vodi Klasu II na istoj prijemnici (dva reda, BrojZbirne Klase I), PWA je Klasi II dala
        svoju zbirnu -2: zbirna -2 je 'bez prijemnice' samo formalno -> samo INFO, bez SUMNJE."""
        t = fixture.osnovna()
        lanac(t, "OTP-00006", "ZBR-00006", "3/060726-2", "ST-00003", D6, 30, 12, klasa="II")
        prijemnica(t, "PRJ-00006", "1/060726-2", "S3/060726", "ST-00003", "29.5", 12, D6, klasa="II")
        ctx = pokreni(t)
        self.assertEqual(self.kodovi(ctx), [("PRJ_KLASA_RAZLICITA_OD_ZBIRNE", "PRJ-00005, PRJ-00006", INFO),
                                            ("ZBR_ZIVA_BEZ_PRIJEMNICE", "ZBR-00006", INFO)])

    def test_klasa_jedina_prijemnica_bez_duplog_info(self):
        """Klasa II prijemnica na zbirnoj Klase I: jedan nalaz na prijemnici, ne i INFO na istoj zbirnoj."""
        t = fixture.osnovna()
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")["Klasa"] = "II"
        ctx = pokreni(t)
        self.assertEqual(self.kodovi(ctx), [("PRJ_KLASA_RAZLICITA_OD_ZBIRNE", "PRJ-00005", SUMNJA)])

    def test_negativ_odvojene_prijemnice_po_klasi_uz_zivu_zbirnu(self):
        """Forma: otpremnica Klasa I+II pod istim brojem, hladnjaca posebna prijemnica po klasi. Ziva zbirna
        Klase II druge ture (-2) bez prijemnice ne sme da 'otme' prijemnicu Klase II."""
        t = fixture.osnovna()
        lanac(t, "OTP-00006", "ZBR-00006", "3/060726", "ST-00003", D6, 30, 12, klasa="II")
        prijemnica(t, "PRJ-00006", "1/060726-3", "S3/060726", "ST-00003", "29.7", 12, D6, klasa="II")
        lanac(t, "OTP-00007", "ZBR-00007", "3/060726-2", "ST-00003", D6, 30, 12, klasa="II")
        ctx = pokreni(t)
        self.assertEqual(self.kodovi(ctx), [("ZBR_ZIVA_BEZ_PRIJEMNICE", "ZBR-00007", SUMNJA)])
        x = self.jedan(ctx, "ZBR_ZIVA_BEZ_PRIJEMNICE", "ZBR-00007", SUMNJA)
        self.assertNotIn("PRJ-00006", x.opis + x.predlog)

    def test_vozac_stanica_iz_broja_ali_odgovara_svojoj_zbirni(self):
        """VozacID prijemnice = stanica iz broja zbirne, ali prijemnica NE odgovara toj zbirnoj, a odgovara
        zivoj zbirnoj svog vozaca bez prijemnice: pogresan je broj zbirne, ne vozac."""
        t = fixture.osnovna()
        lanac(t, "OTP-00006", "ZBR-00006", "2/060726", "ST-00002", D6, 60, 24, vozac="ST-00003")
        lanac(t, "OTP-00007", "ZBR-00007", "2/060726-2", "ST-00002", D6, 40, 16)
        prijemnica(t, "PRJ-00006", "1/060726-3", "S2/060726", "ST-00002", "39.5", 16, D6)
        ctx = pokreni(t)
        x = self.jedan(ctx, "PRJ_ZBIRNA_DRUGE_STANICE", "PRJ-00006", SUMNJA)
        self.assertIn("BrojZbirne: S2/060726 -> S2/060726-2; kg 39.5 vs zbirna 40, gajbe 16 = 16", x.predlog)
        self.bez_koda(ctx, "PRJ_ZASTAREO_VOZAC")

    def test_ispravka_vozaca_otpremnice_prijemnica_zaostala(self):
        """Storno-ispravka otpremnice promenila VozacID, stara zbirna ostala aktivna (ne kaskadira): prijemnica
        sa starim VozacID pripada novoj (zivoj) zbirni -> zastareo vozac, ne tiho vezivanje na staru."""
        t = fixture.osnovna()
        red(t, "tblOtpremnica", "OtpremnicaID", "OTP-00005").update(Stornirano="Da", VozacID="ST-00002")
        red(t, "tblZbirna", "ZbirnaID", "ZBR-00005")["VozacID"] = "ST-00002"
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")["VozacID"] = "ST-00002"
        lanac(t, "OTP-00006", "ZBR-00006", "3/060726", "ST-00003", D6, 90, 36)
        ctx = pokreni(t)
        x = self.jedan(ctx, "PRJ_ZASTAREO_VOZAC", "PRJ-00005", GRESKA)
        self.assertIn("VozacID: ST-00002 -> ST-00003", x.predlog)

    def test_negativ_ispravka_vozaca_otpremnice_prijemnica_nova(self):
        t = fixture.osnovna()
        red(t, "tblOtpremnica", "OtpremnicaID", "OTP-00005").update(Stornirano="Da", VozacID="ST-00002")
        red(t, "tblZbirna", "ZbirnaID", "ZBR-00005")["VozacID"] = "ST-00002"
        lanac(t, "OTP-00006", "ZBR-00006", "3/060726", "ST-00003", D6, 90, 36)
        self.nista(pokreni(t))

    def test_negativ_ispravka_popravlja_datum_i_zbirnu(self):
        """Legitimna ispravka koja istovremeno popravlja zastareo datum (prefill) i pogresnu zbirnu."""
        t = fixture.osnovna()
        lanac(t, "OTP-00006", "ZBR-00006", "3/060726-2", "ST-00003", D6, 60, 24)
        prijemnica(t, "PRJ-00006", "1/060726-3", "S3/060726", "ST-00003", "59.5", 24, D6 - timedelta(days=5),
                   Stornirano="Da")
        prijemnica(t, "PRJ-00007", "1/060726-4", "S3/060726-2", "ST-00003", "59.5", 24, D6,
                   IspravkaOd="1/060726-3")
        self.nista(pokreni(t, {"tblPrijemnica": KOL_PRJ_ISPR}))

    def test_negativ_ispravka_popravlja_vozaca(self):
        t = fixture.osnovna()
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005").update(Stornirano="Da", VozacID="ST-00002")
        prijemnica(t, "PRJ-00006", "1/060726-3", "S3/060726", "ST-00003", 89, 36, D6, IspravkaOd="1/060726-2")
        self.nista(pokreni(t, {"tblPrijemnica": KOL_PRJ_ISPR}))

    def test_negativ_prijemnica_tacno_prepisuje_nekanonsku_zbirnu(self):
        """Zbirna i prijemnica imaju isti (nekanonski) broj: stari citaoci ih vezuju, greska nije na prijemnici."""
        t = fixture.osnovna()
        red(t, "tblZbirna", "ZbirnaID", "ZBR-00005")["BrojZbirne"] = "s3/060726"
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")["BrojZbirne"] = "s3/060726"
        self.bez_koda(pokreni(t), "PRJ_BROJ_ZBIRNE_NENORMALIZOVAN")

    def test_negativ_eksterni_broj_prijemnice_sa_razmakom(self):
        """Broj eksternog kupca je kupcev ('INV 77/2026'): unutrasnji razmak nije greska."""
        t = fixture.osnovna()
        t["tblKupci"].append(dict(KupacID="KUP-00002", Naziv="Drugi kupac", Aktivan="Da"))
        lanac(t, "OTP-00006", "ZBR-00006", "2/060726", "ST-00002", D6, 60, 24, kupac="KUP-00002")
        prijemnica(t, "PRJ-00006", "INV 77/2026", "S2/060726", "ST-00002", "59.5", 24, D6, kupac="KUP-00002")
        self.nista(pokreni(t))

    def test_unutrasnji_razmak_u_broju_zbirne(self):
        t = fixture.osnovna()
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")["BrojZbirne"] = "S3/ 060726"
        x = self.jedan(pokreni(t), "PRJ_BROJ_ZBIRNE_NENORMALIZOVAN", "PRJ-00005", GRESKA)
        self.assertIn("-> 'S3/060726'", x.predlog)

    def test_kolicina_tekst_nije_manjak(self):
        """Tekst umesto broja u Kolicina nije 'manjak 100%'."""
        t = fixture.osnovna()
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")["Kolicina"] = "abc"
        ctx = pokreni(t)
        self.assertEqual([x.opis for x in ctx.nalazi if x.kod == "MODUL_PAO"], [])
        self.bez_koda(ctx, "PRJ_KG_GAJBE_VAN_OPSEGA")

    def test_gajbe_tekst_u_opisu(self):
        t = fixture.osnovna()
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")["KolAmbalaze"] = "x"
        x = self.jedan(pokreni(t), "PRJ_GAJBE_NULA_ILI_BEZ_TIPA", "PRJ-00005", SUMNJA)
        self.assertIn("'x' (nije broj)", x.opis)

    def test_negativ_klasa_rucno_kucana(self):
        """'Klasa I' / '1' su ista klasa kao 'I'."""
        t = fixture.osnovna()
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")["Klasa"] = "Klasa I"
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00004")["Klasa"] = "1"
        self.nista(pokreni(t))

    def test_dve_nejasno_sumnjiva_ona_sa_kandidatom(self):
        """Nijedna sama ne odgovara zbirnoj: prijavljuje se prijemnica koja odgovara zivoj zbirnoj bez
        prijemnice, i kad ima nizi seq (pogresan klik pa tacan unos)."""
        t = fixture.osnovna()
        red(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005").update(Kolicina=Decimal("59.5"), KolAmbalaze=24)
        lanac(t, "OTP-00006", "ZBR-00006", "3/060726-2", "ST-00003", D6, 60, 24)
        prijemnica(t, "PRJ-00006", "1/060726-3", "S3/060726", "ST-00003", 80, 36, D6)
        ctx = pokreni(t)
        x = self.jedan(ctx, "PRJ_DVE_NA_ZBIRNOJ_ISTA_STANICA", "PRJ-00005", SUMNJA)
        self.assertIn("S3/060726 -> S3/060726-2", x.predlog)
        self.assertEqual([y.id for y in nalazi(ctx, "PRJ_DVE_NA_ZBIRNOJ_ISTA_STANICA")], ["PRJ-00005"])

    def test_negativ_hladnjaca_lanac_sa_pravim_vozacem(self):
        """Auto-lanac hladnjace sa VozacID pravog vozaca (bez S prefiksa): nista."""
        t = fixture.osnovna()
        red(t, "tblStanice", "StanicaID", "ST-00003")["JeHladnjaca"] = "Da"
        for tb in ("tblOtpremnica", "tblZbirna", "tblPrijemnica"):
            for r in t[tb]:
                if r.get("BrojZbirne") == "S3/060726":
                    r.update(VozacID="VOZ-00001", BrojZbirne="3/060726")
        red(t, "tblOtkup", "OtkupID", "OTK-00006").update(VozacID="VOZ-00001", BrojZbirne="3/060726")
        self.nista(pokreni(t))


# ------------------------------------------------------------------------------------------
class Performanse(Osnova):
    def test_1000_dokumenata_cisto_i_brzo(self):
        t = fixture.osnovna()
        for i in range(1, 11):
            stanica(t, i)
        br = {}
        nid = 6
        prj_po_danu = {}
        for j in range(1000):
            st_i = j % 10 + 1
            dan = date(2026, 5, 1) + timedelta(days=j // 20)
            st = "ST-%05d" % st_i
            kl = (st_i, dan)
            br[kl] = br.get(kl, 0) + 1
            broj = "%d/%s%s" % (st_i, dan.strftime("%d%m%y"), "-%d" % br[kl] if br[kl] > 1 else "")
            kg, gaj = 50 + (j % 37), 20 + (j % 15)
            bz = lanac(t, "OTP-%05d" % nid, "ZBR-%05d" % nid, broj, st, dan, kg, gaj)
            prj_po_danu[dan] = prj_po_danu.get(dan, 0) + 1
            q = prj_po_danu[dan]
            brp = "1/%s%s" % (dan.strftime("%d%m%y"), "-%d" % q if q > 1 else "")
            prijemnica(t, "PRJ-%05d" % nid, brp, bz, st, Decimal(kg) - Decimal("0.5"), gaj, dan)
            for a in range(5):
                t["tblAmbalaza"].append(dict(AmbID="AMB-9%05d%d" % (nid, a), Datum=dan, TipAmbalaze="12/1",
                                             Kolicina=gaj, Smer="Ulaz", EntitetID=st, EntitetTip="Stanica",
                                             VozacID=st, DokumentID="OTP-%05d" % nid, DokumentTIP="Otpremnica",
                                             Stornirano=""))
            nid += 1
        tabele = tabele_iz(t)
        t0 = time.time()
        ctx = pokreni_nad_tabelama(tabele, moduli=("prijemnica",))
        trajanje = time.time() - t0
        self.nista(ctx)
        self.assertLess(trajanje, 20, "trajanje %.1f s" % trajanje)


if __name__ == "__main__":
    unittest.main()
