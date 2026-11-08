"""Testovi provera otpremnice i zbirne (provera/otpremnica_zbirna.py).

Svaki test polazi od ciste fixture.osnovna(), podmetne gresku (ili legitiman
obrazac) i trazi kod na tacnom ID-ju i ozbiljnosti. Tabele idu kroz pravi .xlsx
(tabele_iz), kao u pravoj svesci.

    cd tools/migracija && python3 -m unittest provera.test_otpremnica_zbirna -v
"""

import unittest
from datetime import date

from provera import fixture
from provera import otpremnica_zbirna as oz
from provera.pokretac import pokreni_nad_tabelama
from provera.test_zajednicko import tabele_iz
from provera.zajednicko import GRESKA, INFO, SUMNJA

MOD = "otpremnica_zbirna"
D5, D6 = fixture.D5, fixture.D6


# ---------------------------------------------------------------- pomocne za izmene fixture
def kopiraj(t, ime, kol, vrednost, **izmene):
    r = dict(fixture.red(t, ime, kol, vrednost))
    r.update(izmene)
    t[ime].append(r)
    return r


def obrisi(t, ime, kol, vrednost):
    t[ime] = [r for r in t[ime] if r.get(kol) != vrednost]


def postavi(t, ime, kol, vrednost, **izmene):
    fixture.red(t, ime, kol, vrednost).update(izmene)


def preimenuj_zbirnu(t, stari, novi):
    """Broj zbirne je veza po stringu: menja se na svim dokumentima lanca."""
    for ime in ("tblOtpremnica", "tblZbirna", "tblPrijemnica", "tblOtkup", "tblPaletaStavka"):
        for r in t[ime]:
            if r.get("BrojZbirne") == stari:
                r["BrojZbirne"] = novi


def config(t, kljuc, vrednost):
    for r in t["tblSEFConfig"]:
        if r["ConfigKey"] == kljuc:
            r["ConfigValue"] = vrednost


def storno_veza(t, **polja):
    r = dict(CorrectionID="COR-%05d" % (len(t["tblStornoVeze"]) + 1), CreatedAt="2026-07-06 12:00:00",
             CreatedBy="op", NeedsRecovery="")
    r.update(polja)
    t["tblStornoVeze"].append(r)


class Baza(unittest.TestCase):
    def nalazi(self, t, kolone=None):
        ctx = pokreni_nad_tabelama(tabele_iz(t, kolone), moduli=(MOD,))
        nal = [x for x in ctx.nalazi if x.modul == MOD]
        self.assertEqual([(x.id, x.opis) for x in nal if x.kod == "MODUL_PAO"], [])
        for x in nal:
            self.assertTrue((x.opis + x.predlog).isascii(), x.opis + x.predlog)
        return nal

    @staticmethod
    def sazmi(nal):
        return [(x.kod, x.ozbiljnost, x.id) for x in nal]

    def jedan(self, nal, kod, id_, ozbiljnost):
        hit = [x for x in nal if x.kod == kod]
        self.assertEqual([(x.id, x.ozbiljnost) for x in hit], [(id_, ozbiljnost)], self.sazmi(nal))
        return hit[0]

    def nema(self, nal, *kodovi):
        self.assertEqual([x for x in self.sazmi(nal) if x[0] in kodovi], [])


# ---------------------------------------------------------------- cista sveska
class Cista(Baza):
    def test_cista_bez_nalaza(self):
        self.assertEqual(self.sazmi(self.nalazi(fixture.osnovna())), [])

    def test_vise_tura_istog_dana_legitimno(self):
        """Treca tura iste stanice i dana (-3), isti vozac: nista."""
        t = fixture.kopija(fixture.osnovna())
        kopiraj(t, "tblOtpremnica", "OtpremnicaID", "OTP-00002", OtpremnicaID="OTP-00006",
                BrojOtpremnice="1/050726-3", BrojZbirne="S1/050726-3", Kolicina=30)
        kopiraj(t, "tblZbirna", "ZbirnaID", "ZBR-00002", ZbirnaID="ZBR-00006", BrojZbirne="S1/050726-3",
                UkupnoKolicina=30)
        kopiraj(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00002", PrijemnicaID="PRJ-00006",
                BrojPrijemnice="1/050726-4", BrojZbirne="S1/050726-3", Kolicina=29)
        self.assertEqual(self.sazmi(self.nalazi(t)), [])

    def test_pwa_otpremnica_po_klasi_i_prijemnica_sutradan_legitimno(self):
        """PWA daje SVAKOJ klasi svoj broj i svoju zbirnu (-2 za Klasu II); prijemnica (Klasa I i II, jedan
        broj) vezana za zbirnu Klase I i datirana dan posle zbirne."""
        t = fixture.kopija(fixture.osnovna())
        kopiraj(t, "tblOtpremnica", "OtpremnicaID", "OTP-00004", OtpremnicaID="OTP-00006", Klasa="II",
                BrojOtpremnice="1/060726-2", BrojZbirne="S1/060726-2", Kolicina=10, KolAmbalaze=4)
        kopiraj(t, "tblZbirna", "ZbirnaID", "ZBR-00004", ZbirnaID="ZBR-00006", Klasa="II", BrojZbirne="S1/060726-2",
                UkupnoKolicina=10, UkupnoAmbalaze=4)
        kopiraj(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00004", PrijemnicaID="PRJ-00006", Klasa="II", Kolicina=9.9,
                KolAmbalaze=4)
        sutra = date(2026, 7, 7)
        for prj in ("PRJ-00004", "PRJ-00006"):
            postavi(t, "tblPrijemnica", "PrijemnicaID", prj, Datum=sutra)
        self.assertEqual(self.sazmi(self.nalazi(t)), [])


# ---------------------------------------------------------------- vozac
class Vozac(Baza):
    def _druga_stanica(self, t):
        for ime, kol, v in (("tblOtpremnica", "OtpremnicaID", "OTP-00005"), ("tblZbirna", "ZbirnaID", "ZBR-00005"),
                            ("tblPrijemnica", "PrijemnicaID", "PRJ-00005"), ("tblOtkup", "OtkupID", "OTK-00006")):
            postavi(t, ime, kol, v, VozacID="ST-00002")

    def test_vozac_druge_stanice_ceo_lanac(self):
        t = fixture.kopija(fixture.osnovna())
        self._druga_stanica(t)
        nal = self.nalazi(t)
        x = self.jedan(nal, oz.K_VOZAC, "OTP-00005", GRESKA)
        self.assertIn("ZBR-00005", x.opis)
        self.assertIn("PRJ-00005", x.opis)
        self.assertIn("izvor: otkup", x.opis)
        self.assertIn("VozacID: ST-00002 -> ST-00003 na OTP-00005, ZBR-00005, PRJ-00005", x.predlog)

    def test_samo_zbirna_ima_tudjeg_vozaca(self):
        t = fixture.kopija(fixture.osnovna())
        postavi(t, "tblZbirna", "ZbirnaID", "ZBR-00005", VozacID="ST-00002")
        x = self.jedan(self.nalazi(t), oz.K_VOZAC, "ZBR-00005", GRESKA)
        self.assertIn("ST-00002 -> ST-00003", x.predlog)

    def test_izvor_prefill_ispravke(self):
        """P7: stornirana otpremnica iste stanice i dana nosi isti pogresan vozac."""
        t = fixture.kopija(fixture.osnovna())
        postavi(t, "tblOtpremnica", "OtpremnicaID", "OTP-00005", VozacID="ST-00002")
        postavi(t, "tblZbirna", "ZbirnaID", "ZBR-00005", VozacID="ST-00002")
        kopiraj(t, "tblOtpremnica", "OtpremnicaID", "OTP-00005", OtpremnicaID="OTP-00009",
                BrojOtpremnice="3/060726-9", BrojZbirne="", VozacID="ST-00002", Stornirano="Da")
        x = self.jedan(self.nalazi(t), oz.K_VOZAC, "OTP-00005", GRESKA)
        self.assertIn("prefill ispravke", x.opis)
        self.assertIn("OTP-00009", x.opis)

    def test_bez_malina_moda_je_sumnja(self):
        t = fixture.kopija(fixture.osnovna())
        self._druga_stanica(t)
        config(t, "MALINA_MODE", "NO")
        self.jedan(self.nalazi(t), oz.K_VOZAC, "OTP-00005", SUMNJA)

    def test_pravi_vozac_dosledno_samo_info(self):
        t = fixture.kopija(fixture.osnovna())
        preimenuj_zbirnu(t, "S3/060726", "3/060726")
        for ime, kol, v in (("tblOtpremnica", "OtpremnicaID", "OTP-00005"), ("tblZbirna", "ZbirnaID", "ZBR-00005"),
                            ("tblPrijemnica", "PrijemnicaID", "PRJ-00005"), ("tblOtkup", "OtkupID", "OTK-00006")):
            postavi(t, ime, kol, v, VozacID="VOZ-00001")
        self.assertEqual(self.sazmi(self.nalazi(t)), [(oz.K_VOZAC, INFO, "VOZ-00001")])


# ---------------------------------------------------------------- P4 tura po vozacu
class TuraPoVozacu(Baza):
    def test_druga_tura_sa_vozacem_druge_stanice(self):
        t = fixture.kopija(fixture.osnovna())
        for ime, kol, v in (("tblOtpremnica", "OtpremnicaID", "OTP-00002"), ("tblZbirna", "ZbirnaID", "ZBR-00002"),
                            ("tblPrijemnica", "PrijemnicaID", "PRJ-00002"), ("tblOtkup", "OtkupID", "OTK-00003")):
            postavi(t, ime, kol, v, VozacID="ST-00002")
        nal = self.nalazi(t)
        x = self.jedan(nal, oz.K_TURA_VOZ, "OTP-00001, OTP-00002", SUMNJA)
        self.assertIn("OTK-00003 VozacID ST-00002", x.opis)
        self.jedan(nal, oz.K_VOZAC, "OTP-00002", GRESKA)

    def test_dve_ture_isti_vozac_nista(self):
        nal = self.nalazi(fixture.osnovna())     # OTP-00001 i OTP-00002: ista stanica, dan, klasa i vozac
        self.nema(nal, oz.K_TURA_VOZ)


# ---------------------------------------------------------------- datum u broju
class Datum(Baza):
    def _broj(self, t, otp, novi, stara_zbirna):
        postavi(t, "tblOtpremnica", "OtpremnicaID", otp, BrojOtpremnice=novi)
        preimenuj_zbirnu(t, stara_zbirna, "S" + novi)

    def test_broj_drugog_dana(self):
        t = fixture.kopija(fixture.osnovna())
        self._broj(t, "OTP-00004", "1/050726-3", "S1/060726")
        x = self.jedan(self.nalazi(t), oz.K_DATUM, "OTP-00004", GRESKA)
        self.assertIn("razlika +1 dana", x.opis)
        self.assertIn("ZBR-00004", x.opis)           # zbirna nasledjuje -- nema poseban nalaz
        self.assertIn("1/050726-3 -> 1/060726", x.predlog)

    def test_zamena_dan_mesec(self):
        t = fixture.kopija(fixture.osnovna())
        self._broj(t, "OTP-00004", "1/070626", "S1/060726")
        x = self.jedan(self.nalazi(t), oz.K_DATUM, "OTP-00004", GRESKA)
        self.assertIn("zamena dan/mesec", x.opis)

    def test_datum_pogresan_broj_tacan(self):
        t = fixture.kopija(fixture.osnovna())
        postavi(t, "tblOtpremnica", "OtpremnicaID", "OTP-00004", Datum=D5)
        x = self.jedan(self.nalazi(t), oz.K_DATUM, "OTP-00004", GRESKA)
        self.assertIn("Datum: 05.07.2026 -> 06.07.2026", x.predlog)

    def test_samo_zbirna_ima_drugi_datum(self):
        t = fixture.kopija(fixture.osnovna())
        postavi(t, "tblZbirna", "ZbirnaID", "ZBR-00004", Datum=D5)
        x = self.jedan(self.nalazi(t), oz.K_DATUM, "ZBR-00004", GRESKA)
        self.assertIn("-> 06.07.2026 (kao otpremnica OTP-00004)", x.predlog)

    def test_unazad_datiran_unos_legitimno(self):
        """Broj i Datum istog (starog) dana: nije nesklad."""
        t = fixture.kopija(fixture.osnovna())
        d4 = date(2026, 7, 4)
        self._broj(t, "OTP-00004", "1/040726", "S1/060726")
        for ime, kol, v in (("tblOtpremnica", "OtpremnicaID", "OTP-00004"), ("tblZbirna", "ZbirnaID", "ZBR-00004"),
                            ("tblPrijemnica", "PrijemnicaID", "PRJ-00004"), ("tblOtkup", "OtkupID", "OTK-00005")):
            postavi(t, ime, kol, v, Datum=d4)
        self.assertEqual(self.sazmi(self.nalazi(t)), [])

    def test_hladnjacki_lanac_nije_ovde(self):
        """Broj otpremnice hladnjace je broj otkupa -- nesklad pripada oblasti otkupa."""
        t = fixture.kopija(fixture.osnovna())
        postavi(t, "tblStanice", "StanicaID", "ST-00003", JeHladnjaca="Da")
        self._broj(t, "OTP-00005", "3/050726", "S3/060726")
        self.nema(self.nalazi(t), oz.K_DATUM)

    def test_datum_nije_datum(self):
        t = fixture.kopija(fixture.osnovna())
        postavi(t, "tblOtpremnica", "OtpremnicaID", "OTP-00004", Datum="abc")
        x = self.jedan(self.nalazi(t), oz.K_DATUM, "OTP-00004", SUMNJA)
        self.assertIn("06.07.2026", x.predlog)

    def test_nekanonski_broj_info(self):
        t = fixture.kopija(fixture.osnovna())
        self._broj(t, "OTP-00004", "RUCNO-7", "S1/060726")
        self.jedan(self.nalazi(t), oz.K_DATUM, "OTP-00004", INFO)


# ---------------------------------------------------------------- tudja stanica
class TudjaStanica(Baza):
    def test_prefiks_druge_stanice(self):
        t = fixture.kopija(fixture.osnovna())
        postavi(t, "tblOtpremnica", "OtpremnicaID", "OTP-00005", BrojOtpremnice="2/060726")
        preimenuj_zbirnu(t, "S3/060726", "S2/060726")
        nal = self.nalazi(t)
        x = self.jedan(nal, oz.K_TUDJA, "OTP-00005", GRESKA)
        self.assertIn("pripada ST-00002", x.opis)
        self.assertIn("2/060726 -> 3/060726", x.predlog)   # otkup OTK-00006 je stanice ST-00003
        self.nema(nal, oz.K_VOZAC)                         # vozac je potvrdjen otpremnicom

    def test_vodece_nule_ista_stanica(self):
        t = fixture.kopija(fixture.osnovna())
        postavi(t, "tblOtpremnica", "OtpremnicaID", "OTP-00005", BrojOtpremnice="03/060726")
        preimenuj_zbirnu(t, "S3/060726", "S03/060726")
        self.assertEqual(self.sazmi(self.nalazi(t)), [])


# ---------------------------------------------------------------- dupli brojevi
class Dupli(Baza):
    def test_isti_broj_otpremnice_drugi_dan(self):
        t = fixture.kopija(fixture.osnovna())
        postavi(t, "tblOtpremnica", "OtpremnicaID", "OTP-00004", BrojOtpremnice="1/050726")
        x = self.jedan(self.nalazi(t), oz.K_OTP_DUPLI, "OTP-00001, OTP-00004", GRESKA)
        self.assertIn("zadrzati broj na OTP-00001", x.predlog)
        self.assertIn("OTP-00004 -> 1/060726", x.predlog)

    def test_klasa_I_i_II_istog_dokumenta_legitimno(self):
        t = fixture.kopija(fixture.osnovna())
        kopiraj(t, "tblOtpremnica", "OtpremnicaID", "OTP-00004", OtpremnicaID="OTP-00006", Klasa="II", Kolicina=10,
                KolAmbalaze=4)
        kopiraj(t, "tblZbirna", "ZbirnaID", "ZBR-00004", ZbirnaID="ZBR-00006", Klasa="II", UkupnoKolicina=10,
                UkupnoAmbalaze=4)
        kopiraj(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00004", PrijemnicaID="PRJ-00006", Klasa="II", Kolicina=9.9,
                KolAmbalaze=4)
        self.assertEqual(self.sazmi(self.nalazi(t)), [])

    def test_storno_pa_ponovni_unos_istog_broja_legitimno(self):
        t = fixture.kopija(fixture.osnovna())
        postavi(t, "tblOtpremnica", "OtpremnicaID", "OTP-00004", Stornirano="Da")
        postavi(t, "tblZbirna", "ZbirnaID", "ZBR-00004", Stornirano="Da")
        kopiraj(t, "tblOtpremnica", "OtpremnicaID", "OTP-00004", OtpremnicaID="OTP-00006", Stornirano="")
        kopiraj(t, "tblZbirna", "ZbirnaID", "ZBR-00004", ZbirnaID="ZBR-00006", Stornirano="")
        postavi(t, "tblOtkup", "OtkupID", "OTK-00005", OtpremnicaID="OTP-00006")
        self.assertEqual(self.sazmi(self.nalazi(t)), [])

    def test_dve_aktivne_zbirne_istog_broja(self):
        t = fixture.kopija(fixture.osnovna())
        kopiraj(t, "tblZbirna", "ZbirnaID", "ZBR-00004", ZbirnaID="ZBR-00006")
        x = self.jedan(self.nalazi(t), oz.K_ZBR_DUPLI, "ZBR-00004, ZBR-00006", GRESKA)
        self.assertIn("ZBR-00004 (= OTP-00004)", x.predlog)
        self.assertIn("stornirati ZBR-00006", x.predlog)


# ---------------------------------------------------------------- zbirna sa storniranim otpremnicama
class StorniraneOtpremnice(Baza):
    def _naslednik(self, t, broj_zbirne):
        kopiraj(t, "tblOtpremnica", "OtpremnicaID", "OTP-00005", OtpremnicaID="OTP-00006",
                BrojOtpremnice="3/060726-2", BrojZbirne=broj_zbirne, Stornirano="")
        postavi(t, "tblOtkup", "OtkupID", "OTK-00006", OtpremnicaID="OTP-00006", BrojZbirne=broj_zbirne)

    def test_visi_sa_prijemnicom_i_naslednikom(self):
        t = fixture.kopija(fixture.osnovna())
        postavi(t, "tblOtpremnica", "OtpremnicaID", "OTP-00005", Stornirano="Da")
        self._naslednik(t, "S3/060726-2")
        kopiraj(t, "tblZbirna", "ZbirnaID", "ZBR-00005", ZbirnaID="ZBR-00006", BrojZbirne="S3/060726-2")
        nal = self.nalazi(t)
        x = self.jedan(nal, oz.K_VISI, "ZBR-00005", GRESKA)
        self.assertIn("PRJ-00005", x.opis)
        self.assertIn("pre v2.4.0", x.opis)
        self.assertIn("PRJ-00005: BrojZbirne S3/060726 -> S3/060726-2", x.predlog)
        self.nema(nal, oz.K_TURA, oz.K_ISPR)

    def test_storno_sa_kaskadom_nista(self):
        t = fixture.kopija(fixture.osnovna())
        for ime, kol, v in (("tblOtpremnica", "OtpremnicaID", "OTP-00005"), ("tblZbirna", "ZbirnaID", "ZBR-00005"),
                            ("tblPrijemnica", "PrijemnicaID", "PRJ-00005"), ("tblOtkup", "OtkupID", "OTK-00006")):
            postavi(t, ime, kol, v, Stornirano="Da")
        self.assertEqual(self.sazmi(self.nalazi(t)), [])

    def _ispravka_bez_nove_zbirne(self, t):
        postavi(t, "tblOtpremnica", "OtpremnicaID", "OTP-00005", Stornirano="Da")
        postavi(t, "tblZbirna", "ZbirnaID", "ZBR-00005", UkupnoKolicina=0, UkupnoAmbalaze=0)
        self._naslednik(t, "")

    def test_ispravka_zavrsena_bez_nove_zbirne_kontekst(self):
        t = fixture.kopija(fixture.osnovna())
        self._ispravka_bez_nove_zbirne(t)
        storno_veza(t, Mode="ISPRAVKA_ODMAH", Status="COMPLETED", OldDocType="Otpremnica", OldDocID="OTP-00005",
                    OldBroj="3/060726", NewDocType="Otpremnica", NewDocID="OTP-00006", NewBroj="3/060726-2",
                    ParentDocType="Zbirna", ParentBroj="S3/060726")
        nal = self.nalazi(t)
        x = self.jedan(nal, oz.K_ISPR, "ZBR-00005", GRESKA)
        self.assertIn("COR-00001", x.opis)
        self.assertIn("Napraviti zbirnu S3/060726-2", x.predlog)
        self.assertIn("PRJ-00005: BrojZbirne S3/060726 -> S3/060726-2", x.predlog)
        self.nema(nal, oz.K_VISI)
        self.jedan(nal, oz.K_BEZ, "OTP-00006", SUMNJA)      # naslednik bez zbirne -- drugi dokument

    def test_ispravka_bez_konteksta_je_sumnja(self):
        t = fixture.kopija(fixture.osnovna())
        self._ispravka_bez_nove_zbirne(t)
        self.jedan(self.nalazi(t), oz.K_ISPR, "ZBR-00005", SUMNJA)

    def test_ispravka_uspela_nista(self):
        t = fixture.kopija(fixture.osnovna())
        postavi(t, "tblOtpremnica", "OtpremnicaID", "OTP-00005", Stornirano="Da")
        postavi(t, "tblZbirna", "ZbirnaID", "ZBR-00005", Stornirano="Da")
        self._naslednik(t, "S3/060726-2")
        kopiraj(t, "tblZbirna", "ZbirnaID", "ZBR-00005", ZbirnaID="ZBR-00006", BrojZbirne="S3/060726-2",
                Stornirano="")
        postavi(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005", BrojZbirne="S3/060726-2")
        storno_veza(t, Mode="ISPRAVKA_ODMAH", Status="COMPLETED", OldDocType="Otpremnica", OldDocID="OTP-00005",
                    OldBroj="3/060726", NewDocID="OTP-00006", NewBroj="3/060726-2", ParentBroj="S3/060726")
        self.assertEqual(self.sazmi(self.nalazi(t)), [])

    def test_deo_storniran_zbir_nije_rekalkulisan(self):
        t = fixture.kopija(fixture.osnovna())
        kopiraj(t, "tblOtpremnica", "OtpremnicaID", "OTP-00002", OtpremnicaID="OTP-00006",
                BrojOtpremnice="1/050726-3", BrojZbirne="S1/050726", Kolicina=30, Stornirano="Da")
        postavi(t, "tblZbirna", "ZbirnaID", "ZBR-00001", UkupnoKolicina=130)
        x = self.jedan(self.nalazi(t), oz.K_DELIM, "ZBR-00001", SUMNJA)
        self.assertIn("130 -> 100", x.predlog)

    def test_deo_storniran_zbir_rekalkulisan_nista(self):
        t = fixture.kopija(fixture.osnovna())
        kopiraj(t, "tblOtpremnica", "OtpremnicaID", "OTP-00002", OtpremnicaID="OTP-00006",
                BrojOtpremnice="1/050726-3", BrojZbirne="S1/050726", Kolicina=30, Stornirano="Da")
        self.assertEqual(self.sazmi(self.nalazi(t)), [])


# ---------------------------------------------------------------- otpremnica na storniranoj zbirnoj
class Ceka(Baza):
    def test_zbirna_postoji_samo_stornirana(self):
        t = fixture.kopija(fixture.osnovna())
        postavi(t, "tblZbirna", "ZbirnaID", "ZBR-00003", Stornirano="Da")
        x = self.jedan(self.nalazi(t), oz.K_CEKA, "OTP-00003", GRESKA)
        self.assertIn("PRJ-00003", x.opis)
        self.assertIn("Napraviti novu zbirnu S2/050726", x.predlog)

    def test_zbirna_ponovo_napravljena_nista(self):
        t = fixture.kopija(fixture.osnovna())
        postavi(t, "tblZbirna", "ZbirnaID", "ZBR-00003", Stornirano="Da")
        kopiraj(t, "tblZbirna", "ZbirnaID", "ZBR-00003", ZbirnaID="ZBR-00006", Stornirano="")
        self.assertEqual(self.sazmi(self.nalazi(t)), [])

    def test_tacno_poredjenje_vidi_samo_storniranu(self):
        t = fixture.kopija(fixture.osnovna())
        postavi(t, "tblOtpremnica", "OtpremnicaID", "OTP-00003", BrojZbirne="s2/050726")
        kopiraj(t, "tblZbirna", "ZbirnaID", "ZBR-00003", ZbirnaID="ZBR-00006")
        postavi(t, "tblZbirna", "ZbirnaID", "ZBR-00003", BrojZbirne="s2/050726", Stornirano="Da")
        nal = self.nalazi(t)
        x = self.jedan(nal, oz.K_CEKA_TACNO, "OTP-00003", SUMNJA)
        self.assertIn("'s2/050726' -> 'S2/050726'", x.predlog)
        self.nema(nal, oz.K_CEKA)


# ---------------------------------------------------------------- malina: bez zbirne
class BezZbirne(Baza):
    def test_veza_izgubljena(self):
        t = fixture.kopija(fixture.osnovna())
        postavi(t, "tblOtpremnica", "OtpremnicaID", "OTP-00005", BrojZbirne="")
        nal = self.nalazi(t)
        x = self.jedan(nal, oz.K_BEZ, "OTP-00005", GRESKA)
        self.assertIn("'' -> S3/060726 (ZBR-00005)", x.predlog)
        self.nema(nal, oz.K_HL)

    def test_zbirna_nije_napravljena_prijemnica_na_tudjoj(self):
        t = fixture.kopija(fixture.osnovna())
        postavi(t, "tblOtpremnica", "OtpremnicaID", "OTP-00002", BrojZbirne="")
        obrisi(t, "tblZbirna", "ZbirnaID", "ZBR-00002")
        postavi(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00002", BrojZbirne="S1/050726")
        x = self.jedan(self.nalazi(t), oz.K_BEZ, "OTP-00002", SUMNJA)
        self.assertIn("PRJ-00002 (49.5 kg, na S1/050726", x.predlog)

    def test_pad_celog_prolaza(self):
        t = fixture.kopija(fixture.osnovna())
        for otp, zbr in (("OTP-00004", "ZBR-00004"), ("OTP-00005", "ZBR-00005")):
            postavi(t, "tblOtpremnica", "OtpremnicaID", otp, BrojZbirne="")
            obrisi(t, "tblZbirna", "ZbirnaID", zbr)
        nal = self.nalazi(t)
        for otp in ("OTP-00004", "OTP-00005"):
            x = [y for y in nal if y.kod == oz.K_BEZ and y.id == otp]
            self.assertEqual(len(x), 1, self.sazmi(nal))
            self.assertIn("pada celog prolaza", x[0].opis)

    def test_bez_malina_moda_nista(self):
        t = fixture.kopija(fixture.osnovna())
        config(t, "MALINA_MODE", "NO")
        postavi(t, "tblOtpremnica", "OtpremnicaID", "OTP-00005", BrojZbirne="")
        self.nema(self.nalazi(t), oz.K_BEZ, oz.K_HL)


# ---------------------------------------------------------------- nepotpun lanac
class Lanac(Baza):
    def test_otpremnica_na_nepostojecoj_zbirnoj(self):
        t = fixture.kopija(fixture.osnovna())
        obrisi(t, "tblZbirna", "ZbirnaID", "ZBR-00005")
        x = self.jedan(self.nalazi(t), oz.K_HL, "OTP-00005", GRESKA)
        self.assertIn("Napraviti zbirnu S3/060726", x.predlog)

    def test_zbirna_bez_otpremnice_hladnjaca_otkup_postoji(self):
        t = fixture.kopija(fixture.osnovna())
        postavi(t, "tblStanice", "StanicaID", "ST-00003", JeHladnjaca="Da")
        obrisi(t, "tblOtpremnica", "OtpremnicaID", "OTP-00005")
        x = self.jedan(self.nalazi(t), oz.K_HL, "ZBR-00005", GRESKA)
        self.assertIn("hladnjaca", x.opis)
        self.assertIn("Rekonstruisati otpremnicu iz otkupa OTK-00006", x.predlog)

    def test_zbirna_bez_otpremnice_bez_otkupa(self):
        t = fixture.kopija(fixture.osnovna())
        obrisi(t, "tblOtpremnica", "OtpremnicaID", "OTP-00005")
        obrisi(t, "tblOtkup", "OtkupID", "OTK-00006")
        self.jedan(self.nalazi(t), oz.K_HL, "ZBR-00005", SUMNJA)

    def test_pwa_zbirna_vozaca_nista(self):
        t = fixture.kopija(fixture.osnovna())
        kopiraj(t, "tblZbirna", "ZbirnaID", "ZBR-00004", ZbirnaID="ZBR-00006", BrojZbirne="1/060726",
                VozacID="VOZ-00001", Klasa="I/II", SyncSource="PWA", ClientRecordID="c-1")
        self.assertEqual(self.sazmi(self.nalazi(t)), [])

    def test_hladnjacki_lanac_klasa_I_i_II_legitimno(self):
        t = fixture.kopija(fixture.osnovna())
        postavi(t, "tblStanice", "StanicaID", "ST-00003", JeHladnjaca="Da")
        kopiraj(t, "tblOtpremnica", "OtpremnicaID", "OTP-00005", OtpremnicaID="OTP-00006", Klasa="II", Kolicina=10)
        kopiraj(t, "tblZbirna", "ZbirnaID", "ZBR-00005", ZbirnaID="ZBR-00006", Klasa="II", UkupnoKolicina=10)
        kopiraj(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005", PrijemnicaID="PRJ-00006", Klasa="II", Kolicina=10)
        kopiraj(t, "tblOtkup", "OtkupID", "OTK-00006", OtkupID="OTK-00007", Klasa="II", Kolicina=10,
                OtpremnicaID="OTP-00006")
        self.assertEqual(self.sazmi(self.nalazi(t)), [])


# ---------------------------------------------------------------- tura bez prijemnice
class Tura(Baza):
    def test_varijanta_b_ista_stanica(self):
        t = fixture.kopija(fixture.osnovna())
        postavi(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00002", BrojZbirne="S1/050726")
        x = self.jedan(self.nalazi(t), oz.K_TURA, "ZBR-00002", SUMNJA)
        self.assertIn("varijanta b", x.opis)
        self.assertIn("PRJ-00002 (1/050726-2): BrojZbirne S1/050726 -> S1/050726-2", x.predlog)
        self.assertIn("gajbe 20 = 20", x.predlog)

    def test_varijanta_a_druga_stanica(self):
        t = fixture.kopija(fixture.osnovna())
        postavi(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00003", BrojZbirne="S1/050726")
        x = self.jedan(self.nalazi(t), oz.K_TURA, "ZBR-00003", SUMNJA)
        self.assertIn("varijanta a", x.opis)
        self.assertIn("PRJ-00003 (1/050726-3): BrojZbirne S1/050726 -> S2/050726, VozacID ST-00002 (isti)", x.predlog)

    def test_prijemnica_pokriva_dve_ture_nije_ova_klasa(self):
        t = fixture.kopija(fixture.osnovna())
        obrisi(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00002")
        postavi(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00001", Kolicina=148.5)
        self.nema(self.nalazi(t), oz.K_TURA)

    def test_druga_klasa_nije_par(self):
        t = fixture.kopija(fixture.osnovna())
        postavi(t, "tblOtpremnica", "OtpremnicaID", "OTP-00002", Klasa="II")
        postavi(t, "tblZbirna", "ZbirnaID", "ZBR-00002", Klasa="II")
        postavi(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00002", BrojZbirne="S1/050726")
        self.nema(self.nalazi(t), oz.K_TURA)


# ---------------------------------------------------------------- clan van 1:1
class ClanVan(Baza):
    def test_otpremnica_na_tudjoj_zbirnoj(self):
        t = fixture.kopija(fixture.osnovna())
        postavi(t, "tblOtpremnica", "OtpremnicaID", "OTP-00002", BrojZbirne="S1/050726")
        nal = self.nalazi(t)
        x = self.jedan(nal, oz.K_CLAN, "OTP-00002", SUMNJA)
        self.assertIn("S1/050726 -> S1/050726-2 (ZBR-00002)", x.predlog)
        self.nema(nal, oz.K_HL)

    def test_zbirna_sa_otpremnicom_drugog_dana(self):
        t = fixture.kopija(fixture.osnovna())
        postavi(t, "tblOtpremnica", "OtpremnicaID", "OTP-00003", BrojOtpremnice="A/1")
        preimenuj_zbirnu(t, "S2/050726", "SA/1")
        postavi(t, "tblZbirna", "ZbirnaID", "ZBR-00003", Datum=D6)
        self.jedan(self.nalazi(t), oz.K_CLAN, "ZBR-00003", SUMNJA)

    def test_pravi_vozac_vise_stanica(self):
        """Jedan pravi vozac skuplja sa vise OM (period van maline): legitimno. MALINA_MODE je danasnja
        vrednost, istorija nije zapisana -- zato mora da prodje i kad je danas ukljucen."""
        for malina in ("YES", "NO"):
            with self.subTest(MALINA_MODE=malina):
                t = fixture.kopija(fixture.osnovna())
                config(t, "MALINA_MODE", malina)
                obrisi(t, "tblZbirna", "ZbirnaID", "ZBR-00003")
                preimenuj_zbirnu(t, "S2/050726", "1/050726")
                preimenuj_zbirnu(t, "S1/050726", "1/050726")
                postavi(t, "tblZbirna", "ZbirnaID", "ZBR-00001", UkupnoKolicina=180, VozacID="VOZ-00001")
                for otp in ("OTP-00001", "OTP-00003"):
                    postavi(t, "tblOtpremnica", "OtpremnicaID", otp, VozacID="VOZ-00001")
                nal = self.nalazi(t)
                self.nema(nal, oz.K_CLAN, oz.K_HL)
                self.assertEqual([x[1] for x in self.sazmi(nal) if x[0] == oz.K_VOZAC], [INFO])


# ---------------------------------------------------------------- stornirana i aktivna istog broja
class IstiBroj(Baza):
    def test_dupli_fantom_ponovo_ozivljena(self):
        t = fixture.kopija(fixture.osnovna())
        postavi(t, "tblZbirna", "ZbirnaID", "ZBR-00005", Stornirano="Da")
        kopiraj(t, "tblZbirna", "ZbirnaID", "ZBR-00005", ZbirnaID="ZBR-00006", Stornirano="")
        storno_veza(t, Mode="DUPLI_FANTOM", Status="COMPLETED", OldDocType="Zbirna", OldBroj="S3/060726")
        x = self.jedan(self.nalazi(t), oz.K_ISTI, "ZBR-00005, ZBR-00006", SUMNJA)
        self.assertIn("COR-00001", x.opis)

    def test_ponovo_napravljena_bez_dupli_konteksta_nista(self):
        t = fixture.kopija(fixture.osnovna())
        postavi(t, "tblZbirna", "ZbirnaID", "ZBR-00005", Stornirano="Da")
        kopiraj(t, "tblZbirna", "ZbirnaID", "ZBR-00005", ZbirnaID="ZBR-00006", Stornirano="")
        storno_veza(t, Mode="PONISTENJE_BEZ_ZAMENE", Status="COMPLETED", OldDocType="Zbirna", OldBroj="S3/060726")
        self.assertEqual(self.sazmi(self.nalazi(t)), [])


# ---------------------------------------------------------------- otpornost
class Otpornost(Baza):
    def test_nedostaje_kolona(self):
        t = fixture.kopija(fixture.osnovna())
        kol = [c for c in fixture.KOLONE["tblOtpremnica"] if c != "VozacID"]
        nal = self.nalazi(t, {"tblOtpremnica": kol})
        presk = [x for x in nal if x.kod.endswith("_PRESKOCENO")]
        self.assertTrue(presk)
        self.assertTrue(all(x.ozbiljnost == INFO and "tblOtpremnica.VozacID" in x.opis for x in presk))
        self.assertEqual(len({x.kod for x in presk}), len(presk))      # jedan po kodu
        self.assertIn(oz.K_VOZAC + "_PRESKOCENO", {x.kod for x in presk})
        self.assertEqual([x for x in self.sazmi(nal) if not x[0].endswith("_PRESKOCENO")], [])

    def test_nedostaje_tabela_tiho(self):
        t = fixture.kopija(fixture.osnovna())
        del t["tblZbirna"]
        del t["tblStornoVeze"]
        self.assertEqual(self.sazmi(self.nalazi(t)), [])

    def test_prazna_tabela_tiho(self):
        t = fixture.kopija(fixture.osnovna())
        t["tblOtpremnica"] = []
        nal = self.nalazi(t)
        self.assertFalse([x for x in nal if x.kod.endswith("_PRESKOCENO")])

    def test_lose_vrednosti_bez_pada(self):
        t = fixture.kopija(fixture.osnovna())
        postavi(t, "tblOtpremnica", "OtpremnicaID", "OTP-00004", Datum="", Kolicina="abc", BrojOtpremnice="")
        postavi(t, "tblZbirna", "ZbirnaID", "ZBR-00004", Datum="x", UkupnoKolicina="y", Klasa="")
        postavi(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00002", Kolicina="--", BrojZbirne="S1/050726")
        self.nalazi(t)        # samo bez MODUL_PAO (proverava Baza.nalazi)


# ---------------------------------------------------------------- recenzija: napadi na modul
class Recenzija(Baza):
    """Slucajevi koje je nezavisni recenzent nasao (pad, lazni pozitiv, propustena varijanta)."""

    def _dve_prijemnice_na_s1(self, t):
        """Varijanta (b): PRJ-00002 nasledila S1/050726, pa S1/050726-2 (ZBR-00002) ostaje bez prijemnice."""
        postavi(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00002", BrojZbirne="S1/050726")

    def test_datum_van_opsega_bez_pada(self):
        """Datum 31.12.9999 na zbirnoj bez prijemnice: dan +1 izlazi iz opsega -- ne sme MODUL_PAO."""
        t = fixture.kopija(fixture.osnovna())
        postavi(t, "tblZbirna", "ZbirnaID", "ZBR-00005", Datum="31.12.9999")
        obrisi(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00005")
        self.nalazi(t)                     # Baza.nalazi trazi da nema MODUL_PAO

    def test_prijemnica_predlozena_samo_jednoj_zbirnoj(self):
        """Dve zbirne bez prijemnice (ST-00001 -2 i ST-00002 -2) iste kg: PRJ-00002 se predlaze samo zbirnoj
        svoje stanice/vozaca, ne obema."""
        t = fixture.kopija(fixture.osnovna())
        self._dve_prijemnice_na_s1(t)
        kopiraj(t, "tblOtpremnica", "OtpremnicaID", "OTP-00003", OtpremnicaID="OTP-00006", BrojOtpremnice="2/050726-2",
                BrojZbirne="S2/050726-2", Kolicina=49, KolAmbalaze=20)
        kopiraj(t, "tblZbirna", "ZbirnaID", "ZBR-00003", ZbirnaID="ZBR-00006", BrojZbirne="S2/050726-2",
                UkupnoKolicina=49, UkupnoAmbalaze=20)
        postavi(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00003", Kolicina=128)    # pokriva obe ture ST-00002 (P6)
        self.jedan(self.nalazi(t), oz.K_TURA, "ZBR-00002", SUMNJA)

    def test_mrtva_zbirna_nije_tura(self):
        """Zbirna cije su sve otpremnice stornirane (ZBR_VISI) ne sme da primi prijemnicu, ni kad joj kg
        bolje odgovara."""
        t = fixture.kopija(fixture.osnovna())
        self._dve_prijemnice_na_s1(t)
        kopiraj(t, "tblOtpremnica", "OtpremnicaID", "OTP-00002", OtpremnicaID="OTP-00006", BrojOtpremnice="1/050726-3",
                BrojZbirne="S1/050726-3", Kolicina=49.5, Stornirano="Da")
        kopiraj(t, "tblZbirna", "ZbirnaID", "ZBR-00002", ZbirnaID="ZBR-00006", BrojZbirne="S1/050726-3",
                UkupnoKolicina=49.5)
        nal = self.nalazi(t)
        self.jedan(nal, oz.K_TURA, "ZBR-00002", SUMNJA)
        self.jedan(nal, oz.K_VISI, "ZBR-00006", GRESKA)

    def test_jednake_ture_kg_ne_presudjuje(self):
        """Varijanta (b) kad su obe ture iste kg: nalaz ostaje (SUMNJA), predlog kaze da kg ne presudjuje."""
        t = fixture.kopija(fixture.osnovna())
        self._dve_prijemnice_na_s1(t)
        postavi(t, "tblOtpremnica", "OtpremnicaID", "OTP-00002", Kolicina=100)
        postavi(t, "tblZbirna", "ZbirnaID", "ZBR-00002", UkupnoKolicina=100)
        postavi(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00002", Kolicina=99.5)
        x = self.jedan(self.nalazi(t), oz.K_TURA, "ZBR-00002", SUMNJA)
        self.assertIn("kg ne presudjuje", x.predlog)

    def test_prazna_klasa_ne_iskljucuje(self):
        t = fixture.kopija(fixture.osnovna())
        self._dve_prijemnice_na_s1(t)
        postavi(t, "tblZbirna", "ZbirnaID", "ZBR-00002", Klasa="")
        self.jedan(self.nalazi(t), oz.K_TURA, "ZBR-00002", SUMNJA)

    def test_varijanta_b_sledeci_dan(self):
        """Kao S4/200626-2 umesto S4/210626: prijemnica 06.07 nasledila broj zbirne 05.07."""
        t = fixture.kopija(fixture.osnovna())
        postavi(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00004", BrojZbirne="S1/050726-2")
        x = self.jedan(self.nalazi(t), oz.K_TURA, "ZBR-00004", SUMNJA)
        self.assertIn("PRJ-00004 (1/060726): BrojZbirne S1/050726-2 -> S1/060726", x.predlog)

    def test_klasa_I_i_II_pod_dva_broja_legitimno(self):
        """Zbirna sa Kl I i II, prijemnice I i II pod dva broja (rucni unos); sledeca tura jos bez prijemnice:
        prijemnica Klase II se ne predlaze turi Klase I."""
        t = fixture.kopija(fixture.osnovna())
        kopiraj(t, "tblOtpremnica", "OtpremnicaID", "OTP-00004", OtpremnicaID="OTP-00006", Klasa="II", Kolicina=20,
                KolAmbalaze=8)
        kopiraj(t, "tblZbirna", "ZbirnaID", "ZBR-00004", ZbirnaID="ZBR-00006", Klasa="II", UkupnoKolicina=20,
                UkupnoAmbalaze=8)
        kopiraj(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00004", PrijemnicaID="PRJ-00006", Klasa="II", Kolicina=19.8,
                KolAmbalaze=8, BrojPrijemnice="1/060726-3")
        kopiraj(t, "tblOtpremnica", "OtpremnicaID", "OTP-00004", OtpremnicaID="OTP-00007", BrojOtpremnice="1/060726-2",
                BrojZbirne="S1/060726-2", Kolicina=22, KolAmbalaze=9)
        kopiraj(t, "tblZbirna", "ZbirnaID", "ZBR-00004", ZbirnaID="ZBR-00007", BrojZbirne="S1/060726-2",
                UkupnoKolicina=22, UkupnoAmbalaze=9)
        self.assertEqual(self.sazmi(self.nalazi(t)), [])

    def test_zbirna_sa_jednom_prijemnicom_po_klasi_nije_z2(self):
        """S1/060726 ima Kl I i Kl II pod dva broja prijemnice (po jedna po klasi); Kl II prijemnica nosi vozaca
        ST-00002 (zaostao combo, oblast prijemnice). Tura Kl II stanice ST-00002 bez prijemnice nije par: Z2 nema
        dve prijemnice iste klase."""
        t = fixture.kopija(fixture.osnovna())
        kopiraj(t, "tblOtpremnica", "OtpremnicaID", "OTP-00004", OtpremnicaID="OTP-00006", Klasa="II", Kolicina=20,
                KolAmbalaze=8)
        kopiraj(t, "tblZbirna", "ZbirnaID", "ZBR-00004", ZbirnaID="ZBR-00006", Klasa="II", UkupnoKolicina=20,
                UkupnoAmbalaze=8)
        kopiraj(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00004", PrijemnicaID="PRJ-00006", Klasa="II", Kolicina=19.8,
                KolAmbalaze=8, BrojPrijemnice="1/060726-3", VozacID="ST-00002")
        kopiraj(t, "tblOtpremnica", "OtpremnicaID", "OTP-00003", OtpremnicaID="OTP-00007", BrojOtpremnice="2/060726",
                BrojZbirne="S2/060726", Datum=D6, Klasa="II", Kolicina=20, KolAmbalaze=8)
        kopiraj(t, "tblZbirna", "ZbirnaID", "ZBR-00003", ZbirnaID="ZBR-00007", BrojZbirne="S2/060726", Datum=D6,
                Klasa="II", UkupnoKolicina=20, UkupnoAmbalaze=8)
        self.nema(self.nalazi(t), oz.K_TURA)

    def test_prijemnica_druge_klase_ne_ide_na_turu(self):
        """S1/060726 (Kl I + II) ima dve prijemnice Kl II; tura S1/060726-2 je samo Kl I: prijemnica Kl II joj se
        ne predlaze, ni kad joj kg odgovara."""
        t = fixture.kopija(fixture.osnovna())
        kopiraj(t, "tblOtpremnica", "OtpremnicaID", "OTP-00004", OtpremnicaID="OTP-00006", Klasa="II", Kolicina=20,
                KolAmbalaze=8)
        kopiraj(t, "tblZbirna", "ZbirnaID", "ZBR-00004", ZbirnaID="ZBR-00006", Klasa="II", UkupnoKolicina=20,
                UkupnoAmbalaze=8)
        for prj, br, kg in (("PRJ-00006", "1/060726-3", 19.8), ("PRJ-00007", "1/060726-4", 9.9)):
            kopiraj(t, "tblPrijemnica", "PrijemnicaID", "PRJ-00004", PrijemnicaID=prj, Klasa="II", Kolicina=kg,
                    BrojPrijemnice=br)
        kopiraj(t, "tblOtpremnica", "OtpremnicaID", "OTP-00004", OtpremnicaID="OTP-00007", BrojOtpremnice="1/060726-2",
                BrojZbirne="S1/060726-2", Kolicina=10, KolAmbalaze=4)
        kopiraj(t, "tblZbirna", "ZbirnaID", "ZBR-00004", ZbirnaID="ZBR-00007", BrojZbirne="S1/060726-2",
                UkupnoKolicina=10, UkupnoAmbalaze=4)
        self.nema(self.nalazi(t), oz.K_TURA)

    def test_pravi_vozac_druga_tura_legitimno(self):
        """Prvu turu stanice ST-00001 vozi pravi vozac VOZ-00001, drugu sama stanica: nije P4."""
        for malina in ("YES", "NO"):
            with self.subTest(MALINA_MODE=malina):
                t = fixture.kopija(fixture.osnovna())
                config(t, "MALINA_MODE", malina)
                preimenuj_zbirnu(t, "S1/050726", "1/050726")
                for ime, kol, v in (("tblOtpremnica", "OtpremnicaID", "OTP-00001"),
                                    ("tblZbirna", "ZbirnaID", "ZBR-00001"),
                                    ("tblPrijemnica", "PrijemnicaID", "PRJ-00001"),
                                    ("tblOtkup", "OtkupID", "OTK-00001"), ("tblOtkup", "OtkupID", "OTK-00002")):
                    postavi(t, ime, kol, v, VozacID="VOZ-00001")
                self.assertEqual(self.sazmi(self.nalazi(t)), [(oz.K_VOZAC, INFO, "VOZ-00001")])

    def test_nepoznat_vozac_nije_pravi_vozac(self):
        """VozacID kog nema ni u tblVozaci ni u tblStanice: SUMNJA, ne INFO 'pravi vozac'."""
        t = fixture.kopija(fixture.osnovna())
        preimenuj_zbirnu(t, "S3/060726", "3/060726")
        for ime, kol, v in (("tblOtpremnica", "OtpremnicaID", "OTP-00005"), ("tblZbirna", "ZbirnaID", "ZBR-00005"),
                            ("tblPrijemnica", "PrijemnicaID", "PRJ-00005")):
            postavi(t, ime, kol, v, VozacID="VOZ-00099")
        x = self.jedan(self.nalazi(t), oz.K_VOZAC, "VOZ-00099", SUMNJA)
        self.assertIn("OTP-00005", x.opis)

    def test_ocekivana_zbirna_zauzeta_dupli_broj(self):
        """Druga otpremnica istog broja bez BrojZbirne: nalaz postoji (red se ne gubi u dokumentu prve), ali
        S1/060726 je zbirna OTP-00004 -- nije 'veza izgubljena'."""
        t = fixture.kopija(fixture.osnovna())
        kopiraj(t, "tblOtpremnica", "OtpremnicaID", "OTP-00004", OtpremnicaID="OTP-00006", BrojZbirne="", Kolicina=20)
        nal = self.nalazi(t)
        x = self.jedan(nal, oz.K_BEZ, "OTP-00006", SUMNJA)
        self.assertNotIn("'' -> S1/060726", x.predlog)
        self.assertIn(oz.K_OTP_DUPLI, x.predlog)
        self.jedan(nal, oz.K_OTP_DUPLI, "OTP-00004, OTP-00006", GRESKA)

    def test_otpremnica_na_nepostojecoj_zbirnoj_bez_maline_sumnja(self):
        """Van maline forma upise predlozen broj zbirne pre nego sto zbirna postoji: SUMNJA, ne GRESKA."""
        t = fixture.kopija(fixture.osnovna())
        config(t, "MALINA_MODE", "NO")
        obrisi(t, "tblZbirna", "ZbirnaID", "ZBR-00005")
        self.jedan(self.nalazi(t), oz.K_HL, "OTP-00005", SUMNJA)

    def test_neprocitan_datum_nije_drugi_kontekst(self):
        """Klasa II istog dokumenta sa Datumom koji ne moze da se procita: nije dupli broj (to je DATUM)."""
        t = fixture.kopija(fixture.osnovna())
        kopiraj(t, "tblOtpremnica", "OtpremnicaID", "OTP-00004", OtpremnicaID="OTP-00006", Klasa="II", Kolicina=10,
                Datum="xx")
        kopiraj(t, "tblZbirna", "ZbirnaID", "ZBR-00004", ZbirnaID="ZBR-00006", Klasa="II", UkupnoKolicina=10,
                Datum="xx")
        nal = self.nalazi(t)
        self.nema(nal, oz.K_OTP_DUPLI, oz.K_ZBR_DUPLI)
        self.jedan(nal, oz.K_DATUM, "OTP-00006", SUMNJA)

    def test_pwa_zbirna_bez_kolone_syncsource(self):
        """Starija sveska bez kolone SyncSource: PWA zbirnu vozaca prepoznaje ClientRecordID."""
        t = fixture.kopija(fixture.osnovna())
        kopiraj(t, "tblZbirna", "ZbirnaID", "ZBR-00004", ZbirnaID="ZBR-00006", BrojZbirne="1/060726",
                VozacID="VOZ-00001", Klasa="I/II", ClientRecordID="c-1")
        kol = [c for c in fixture.KOLONE["tblZbirna"] if c != "SyncSource"]
        self.nema(self.nalazi(t, {"tblZbirna": kol}), oz.K_HL)


if __name__ == "__main__":
    unittest.main()
