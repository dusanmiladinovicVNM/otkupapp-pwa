"""Testovi provera otkupa (provera/otkup.py).

Svaki test polazi od ciste fixture.osnovna(), podmetne gresku (ili legitiman
obrazac) i trazi kod na tacnom ID-ju i ozbiljnosti. Tabele idu kroz pravi .xlsx
(tabele_iz), kao u pravoj svesci.

    cd tools/migracija && python3 -m unittest provera.test_otkup -v
"""

import unittest
from datetime import date

from provera import fixture
from provera import otkup as ok
from provera.pokretac import pokreni_nad_tabelama
from provera.test_zajednicko import tabele_iz
from provera.zajednicko import GRESKA, INFO, SUMNJA

MOD = "otkup"
D5, D6 = fixture.D5, fixture.D6
D7 = date(2026, 7, 7)
JUN27 = date(2026, 6, 27)


# ---------------------------------------------------------------- pomocne za izmene fixture
def postavi(t, ime, kol, vrednost, **izmene):
    fixture.red(t, ime, kol, vrednost).update(izmene)


def kopiraj(t, ime, kol, vrednost, **izmene):
    r = dict(fixture.red(t, ime, kol, vrednost))
    r.update(izmene)
    t[ime].append(r)
    return r


def otk(t, od="OTK-00001", **izmene):
    return kopiraj(t, "tblOtkup", "OtkupID", od, **izmene)


def otp(t, od="OTP-00001", **izmene):
    return kopiraj(t, "tblOtpremnica", "OtpremnicaID", od, **izmene)


def config(t, kljuc, vrednost):
    for r in t["tblSEFConfig"]:
        if r["ConfigKey"] == kljuc:
            r["ConfigValue"] = vrednost
            return
    t["tblSEFConfig"].append(dict(ConfigKey=kljuc, ConfigValue=vrednost))


def hladnjaca(t, st="ST-00002"):
    postavi(t, "tblStanice", "StanicaID", st, JeHladnjaca="Da")


class Baza(unittest.TestCase):
    def setUp(self):
        self.t = fixture.kopija(fixture.osnovna())

    def nalazi(self, t=None, kolone=None):
        ctx = pokreni_nad_tabelama(tabele_iz(self.t if t is None else t, kolone), moduli=(MOD,))
        nal = [x for x in ctx.nalazi if x.modul == MOD]
        self.assertEqual([(x.id, x.opis) for x in nal if x.kod == "MODUL_PAO"], [])
        for x in nal:
            self.assertTrue((x.opis + x.predlog).isascii(), x.opis + x.predlog)
        return nal

    @staticmethod
    def sazmi(nal):
        return [(x.kod, x.ozbiljnost, x.tabela, x.id) for x in nal]

    def jedan(self, nal, kod, id_, ozbiljnost, tabela="tblOtkup"):
        hit = [x for x in nal if x.kod == kod]
        self.assertEqual([(x.tabela, x.id, x.ozbiljnost) for x in hit], [(tabela, id_, ozbiljnost)], self.sazmi(nal))
        return hit[0]

    def nema(self, nal, *kodovi):
        self.assertEqual([x for x in self.sazmi(nal) if x[0] in kodovi], [])


# ---------------------------------------------------------------- cista sveska i legitimni obrasci
class Cista(Baza):
    def test_cista_bez_nalaza(self):
        self.assertEqual(self.sazmi(self.nalazi()), [])

    def test_lanac_hladnjace_cist(self):
        """ST-00002 je hladnjaca: OTK-00004 je vezan za svoj lanac OTP-00003 -- nije nalaz."""
        hladnjaca(self.t)
        self.assertEqual(self.sazmi(self.nalazi()), [])

    def test_bez_kolone_jehladnjaca(self):
        """Bez JeHladnjaca se lanac prepoznaje po otisku; jedan blok po otpremnici istog broja nije lanac."""
        kol = {"tblStanice": ["StanicaID", "Naziv", "Mesto", "Aktivan"]}
        self.assertEqual(self.sazmi(self.nalazi(kolone=kol)), [])

    def test_vise_tura_istog_dana(self):
        """Treca tura stanice istog dana (-3), sa svojim blokom: nije nalaz."""
        otp(self.t, OtpremnicaID="OTP-00006", BrojOtpremnice="1/050726-3", BrojZbirne="S1/050726-3",
            Kolicina=25, KolAmbalaze=10)
        otk(self.t, OtkupID="OTK-00007", KooperantID="KOOP-00002", BrojDokumenta="1/050726-4",
            OtpremnicaID="OTP-00006", BrojZbirne="S1/050726-3", Kolicina=25, KolAmbalaze=10, BrutoKg=25)
        self.assertEqual(self.sazmi(self.nalazi()), [])


# ---------------------------------------------------------------- OTK_ZASTARELI_KONTEKST_DATUM
class Zastareli(Baza):
    def test_vreme_unosa(self):
        """Naknadni unos za OTP-00004 (06.07) 07.07; posle popune sledeci blok ostaje na 06.07 sa njenom zbirnom."""
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00005", CreatedAt="2026-07-07 08:50:00")
        otp(self.t, OtpremnicaID="OTP-00006", Datum=D7, BrojOtpremnice="1/070726", BrojZbirne="S1/070726",
            Kolicina=25, KolAmbalaze=10)
        otk(self.t, "OTK-00005", OtkupID="OTK-00007", KooperantID="KOOP-00001", BrojDokumenta="1/060726-2",
            OtpremnicaID="", Kolicina=25, KolAmbalaze=10, BrutoKg=25, CreatedAt="2026-07-07 09:00:00",
            VremeUnosa="2026-07-07 09:00:00")
        x = self.jedan(self.nalazi(), ok.K_ZAST, "OTK-00007", SUMNJA)
        self.assertIn("07.07.2026", x.predlog)
        self.assertIn("OTP-00006", x.predlog)
        self.assertIn("OTP-00004", x.opis)

    def test_redosled_bez_vremena(self):
        """Bez VremeUnosa/CreatedAt kasni unos se vidi iz redosleda ID-jeva (raniji unos ima Datum 06.07)."""
        postavi(self.t, "tblOtpremnica", "OtpremnicaID", "OTP-00002", Kolicina=75, KolAmbalaze=30)
        otk(self.t, "OTK-00003", OtkupID="OTK-00007", KooperantID="KOOP-00002", BrojDokumenta="1/050726-4",
            Kolicina=25, KolAmbalaze=10, BrutoKg=25, CreatedAt="", ModifiedAt="")
        otk(self.t, "OTK-00003", OtkupID="OTK-00008", BrojDokumenta="1/050726-5", OtpremnicaID="",
            Kolicina=20, KolAmbalaze=8, BrutoKg=20, CreatedAt="", ModifiedAt="")
        x = self.jedan(self.nalazi(), ok.K_ZAST, "OTK-00008", SUMNJA)
        self.assertIn("06.07.2026", x.predlog)

    def test_naknadni_unos_vezan_nije_nalaz(self):
        """Legitiman naknadni unos: blok od juce unet danas, VEZAN za izabranu otpremnicu."""
        postavi(self.t, "tblOtpremnica", "OtpremnicaID", "OTP-00002", Kolicina=75, KolAmbalaze=30)
        otk(self.t, "OTK-00003", OtkupID="OTK-00007", KooperantID="KOOP-00002", BrojDokumenta="1/050726-4",
            Kolicina=25, KolAmbalaze=10, BrutoKg=25, CreatedAt="2026-07-07 09:00:00")
        self.nema(self.nalazi(), ok.K_ZAST)

    def test_kasni_bez_dokaza_nije_nalaz(self):
        """Nevezan kasni blok bez zbirne te otpremnice, a otpremnica nije popunjena: samo kandidat, ne nalaz."""
        postavi(self.t, "tblOtpremnica", "OtpremnicaID", "OTP-00004", Kolicina=100, KolAmbalaze=40)
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00005", CreatedAt="2026-07-07 08:50:00")
        otk(self.t, "OTK-00005", OtkupID="OTK-00007", KooperantID="KOOP-00001", BrojDokumenta="1/060726-2",
            OtpremnicaID="", BrojZbirne="", Kolicina=25, KolAmbalaze=10, BrutoKg=25,
            CreatedAt="2026-07-07 09:00:00")
        self.nema(self.nalazi(), ok.K_ZAST)

    def test_isti_dan_nije_nalaz(self):
        """Posle popune OTP-00004 sledeci nevezan blok istog dana: datum je tacan, nije zastareli kontekst."""
        otk(self.t, "OTK-00005", OtkupID="OTK-00007", KooperantID="KOOP-00001", BrojDokumenta="1/060726-2",
            OtpremnicaID="", Kolicina=25, KolAmbalaze=10, BrutoKg=25, CreatedAt="2026-07-06 10:10:00")
        self.nema(self.nalazi(), ok.K_ZAST)

    def test_pwa_red_iskljucen(self):
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00005", CreatedAt="2026-07-07 08:50:00")
        otk(self.t, "OTK-00005", OtkupID="OTK-00007", KooperantID="KOOP-00001", BrojDokumenta="1/060726-2",
            OtpremnicaID="", Kolicina=25, KolAmbalaze=10, BrutoKg=25, CreatedAt="2026-07-07 09:00:00",
            SyncSource="PWA", ClientRecordID="c-7")
        self.nema(self.nalazi(), ok.K_ZAST)


# ---------------------------------------------------------------- OTK_VEZA_NESAGLASNA_SA_OTP
class Veza(Baza):
    def test_stanica(self):
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00002", OtpremnicaID="OTP-00003")
        x = self.jedan(self.nalazi(), ok.K_VEZA, "OTK-00002", GRESKA)
        self.assertIn("STANICA", x.opis)
        self.assertIn("OTP-00001", x.predlog)
        self.assertIn("zatvara obe", x.predlog)

    def test_datum_posle_otpremnice(self):
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00003", Datum=D6)
        x = self.jedan(self.nalazi(), ok.K_VEZA, "OTK-00003", GRESKA)
        self.assertIn("POSLE", x.opis)

    def test_datum_dan_pre_sumnja(self):
        """Vecernji otkup, otpremnica sutradan: 1 dan razlike je moguc -- SUMNJA."""
        postavi(self.t, "tblOtpremnica", "OtpremnicaID", "OTP-00002", Datum=D6)
        self.jedan(self.nalazi(), ok.K_VEZA, "OTK-00003", SUMNJA)

    def test_sorta(self):
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00002", SortaVoca="Polka")
        x = self.jedan(self.nalazi(), ok.K_VEZA, "OTK-00002", SUMNJA)
        self.assertIn("SORTA", x.opis)

    def test_klasa_I_II_istog_lista_nije_nalaz(self):
        """SaveOtkupMulti_TX: Klasa I i II istog lista vezane za JEDNU otpremnicu -- po dizajnu."""
        postavi(self.t, "tblOtpremnica", "OtpremnicaID", "OTP-00001", Kolicina=110, KolAmbalaze=44)
        otk(self.t, OtkupID="OTK-00007", Klasa="II", Kolicina=10, KolAmbalaze=4, BrutoKg=10)
        nal = self.nalazi()
        self.nema(nal, ok.K_VEZA, ok.K_DUPLI, ok.K_BROJ)


# ---------------------------------------------------------------- OTK_PANEL_PREGAZIO_VEZU
class Panel(Baza):
    def test_a_pregazena_veza_lanca(self):
        hladnjaca(self.t)
        otp(self.t, "OTP-00003", OtpremnicaID="OTP-00006", BrojOtpremnice="2/050726-2", BrojZbirne="S2/050726-2")
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00004", OtpremnicaID="OTP-00006")
        nal = self.nalazi()
        x = self.jedan(nal, ok.K_PANEL, "OTK-00004", GRESKA)
        self.assertIn("OTP-00003", x.opis)
        self.assertIn("OTP-00006", x.predlog)
        self.nema(nal, ok.K_BZ)

    def test_b_hladnjaca_mrtva_veza(self):
        hladnjaca(self.t)
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00004", OtpremnicaID="OTP-00099")
        nal = self.nalazi()
        x = self.jedan(nal, ok.K_PANEL, "OTK-00004", GRESKA)
        self.assertIn("-> OTP-00003", x.predlog)
        self.nema(nal, ok.K_ZIVA)

    def test_b_vezan_posle_storna(self):
        """Svez blok vezan za otpremnicu stornirana PRE njegovog unosa; stari blok je K_ZIVA."""
        postavi(self.t, "tblOtpremnica", "OtpremnicaID", "OTP-00002", Stornirano="Da",
                ModifiedAt="2026-07-05 09:00:00")
        otk(self.t, "OTK-00003", OtkupID="OTK-00007", KooperantID="KOOP-00002", BrojDokumenta="1/050726-4",
            Kolicina=20, KolAmbalaze=8, BrutoKg=20, CreatedAt="2026-07-05 10:00:00")
        nal = self.nalazi()
        self.jedan(nal, ok.K_PANEL, "OTK-00007", GRESKA)
        self.jedan(nal, ok.K_ZIVA, "OTK-00003", GRESKA)

    def test_storno_pa_ponovni_unos_lanca(self):
        """Storniran stari lanac + aktivan nov lanac istog broja: aktivan je samo jedan -- nije nalaz."""
        hladnjaca(self.t)
        otp(self.t, "OTP-00003", OtpremnicaID="OTP-00006", Stornirano="Da")
        otk(self.t, "OTK-00004", OtkupID="OTK-00007", OtpremnicaID="OTP-00006", Stornirano="Da")
        self.assertEqual(self.sazmi(self.nalazi()), [])


# ---------------------------------------------------------------- DUPLI_LANAC_HLADNJACE_POSLE_GUARD2
class DupliLanac(Baza):
    def test_c_lanac_i_izabrana_otpremnica(self):
        hladnjaca(self.t)
        otp(self.t, "OTP-00003", OtpremnicaID="OTP-00006", BrojOtpremnice="2/050726-2", BrojZbirne="S2/050726-2")
        x = self.jedan(self.nalazi(), ok.K_DLAN, "OTK-00004", SUMNJA)
        self.assertIn("OTP-00006", x.opis)
        self.assertIn("OTP-00003", x.opis)

    def test_c_rucna_otpremnica_druge_kolicine_nije_nalaz(self):
        """Hladnjaca sa rucnom otpremnicom za robu van otkupa: preostalo nije blizu kg lanca."""
        hladnjaca(self.t)
        otp(self.t, "OTP-00003", OtpremnicaID="OTP-00006", BrojOtpremnice="2/050726-2", BrojZbirne="S2/050726-2",
            Kolicina=300)
        self.nema(self.nalazi(), ok.K_DLAN)


# ---------------------------------------------------------------- OTK_BROJ_ZBIRNE_ZAOSTAO
class BrojZbirne(Baza):
    def test_zbirna_druge_otpremnice(self):
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00002", BrojZbirne="S1/050726-2")
        x = self.jedan(self.nalazi(), ok.K_BZ, "OTK-00002", SUMNJA)
        self.assertIn("S1/050726-2 -> S1/050726", x.predlog)

    def test_preraspodela_info(self):
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00002", BrojZbirne="S1/040726")
        self.jedan(self.nalazi(), ok.K_BZ, "OTK-00002", INFO)

    def test_razmak_i_velika_slova_nije_nalaz(self):
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00002", BrojZbirne=" s1/050726 ")
        self.nema(self.nalazi(), ok.K_BZ)

    def test_nevezan_zaostala_zbirna(self):
        otk(self.t, OtkupID="OTK-00007", BrojDokumenta="1/050726-4", OtpremnicaID="", BrojZbirne="S3/060726",
            Kolicina=20, KolAmbalaze=8, BrutoKg=20)
        x = self.jedan(self.nalazi(), ok.K_BZ, "OTK-00007", INFO)
        self.assertIn("ST-00003", x.opis)

    def test_stornirana_zbirna_je_za_modul_storno(self):
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00002", BrojZbirne="S1/040726")
        kopiraj(self.t, "tblZbirna", "ZbirnaID", "ZBR-00001", ZbirnaID="ZBR-00009", BrojZbirne="S1/040726",
                Stornirano="Da")
        self.nema(self.nalazi(), ok.K_BZ)


# ---------------------------------------------------------------- OTK_VOZAC_ZAOSTAO_ILI_LOSE_IZVUCEN / LITERAL
class Vozac(Baza):
    def test_nepostojeci_vozac(self):
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00002", VozacID="VOZ-00009")
        self.jedan(self.nalazi(), ok.K_VOZ, "OTK-00002", GRESKA)

    def test_druga_stanica_malina(self):
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00002", VozacID="ST-00002")
        x = self.jedan(self.nalazi(), ok.K_VOZ, "OTK-00002", SUMNJA)
        self.assertIn("-> ST-00001", x.predlog)

    def test_potvrdjen_otpremnicom_nije_nalaz(self):
        """Vozac druge stanice i na otpremnici: to prijavljuje otpremnica_zbirna (OTP_VOZAC_NIJE_PAR_STANICE)."""
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00002", VozacID="ST-00002")
        postavi(self.t, "tblOtpremnica", "OtpremnicaID", "OTP-00001", VozacID="ST-00002")
        self.nema(self.nalazi(), ok.K_VOZ)

    def test_van_maline_razlicit_od_otpremnice(self):
        config(self.t, "MALINA_MODE", "NO")
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00002", VozacID="VOZ-00001")
        x = self.jedan(self.nalazi(), ok.K_VOZ, "OTK-00002", SUMNJA)
        self.assertIn("-> ST-00001", x.predlog)

    def test_van_maline_pravi_vozac_nevezan_nije_nalaz(self):
        config(self.t, "MALINA_MODE", "NO")
        otk(self.t, OtkupID="OTK-00007", BrojDokumenta="1/050726-4", OtpremnicaID="", VozacID="VOZ-00001",
            Kolicina=20, KolAmbalaze=8, BrutoKg=20)
        self.nema(self.nalazi(), ok.K_VOZ)

    def test_ime_vozaca_sa_zagradom(self):
        self.t["tblVozaci"].append(dict(VozacID="VOZ-00002", Ime="Hladnjaca", Prezime="(hladnjaca)", Aktivan="Da"))
        x = self.jedan(self.nalazi(), ok.K_VOZ, "VOZ-00002", INFO, tabela="tblVozaci")
        self.assertIn("'hladnjaca'", x.opis)

    def test_literal_hladnjaca(self):
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00004", VozacID="hladnjaca")
        postavi(self.t, "tblOtpremnica", "OtpremnicaID", "OTP-00003", VozacID="hladnjaca")
        nal = self.nalazi()
        hit = [(x.tabela, x.id, x.ozbiljnost) for x in nal if x.kod == ok.K_LIT]
        self.assertEqual(hit, [("tblOtkup", "OTK-00004", GRESKA), ("tblOtpremnica", "OTP-00003", GRESKA)])
        self.nema(nal, ok.K_VOZ)
        self.assertIn("ST-00002", [x for x in nal if x.kod == ok.K_LIT][0].predlog)


# ---------------------------------------------------------------- OTK_KOOPERANT_ZAOSTAO
def _junski_par(t, druga_klasa=False):
    """Dva uzastopna unosa istog kooperanta 27.06 (pre v2.7.0), razmak 3 min."""
    otk(t, OtkupID="OTK-00101", Datum=JUN27, BrojDokumenta="1/270626", OtpremnicaID="", BrojZbirne="",
        Kolicina=30, KolAmbalaze=12, BrutoKg=30, Novac=0, PrimalacNovca="", Isplaceno="", DatumIsplate="",
        CreatedAt="2026-06-27 08:00:00", ModifiedAt="")
    if druga_klasa:
        otk(t, "OTK-00101", OtkupID="OTK-00102", Klasa="II", Kolicina=10, KolAmbalaze=4, BrutoKg=10,
            CreatedAt="2026-06-27 08:00:00")
    else:
        otk(t, "OTK-00101", OtkupID="OTK-00102", BrojDokumenta="1/270626-2", Kolicina=12.5, KolAmbalaze=5,
            BrutoKg=12.5, CreatedAt="2026-06-27 08:03:00")


class KooperantZaostao(Baza):
    def test_kandidat_bez_baze_info(self):
        _junski_par(self.t)
        nal = self.nalazi()
        hit = [(x.id, x.ozbiljnost) for x in nal if x.kod == ok.K_KOOP]
        self.assertEqual(hit, [("", INFO), ("OTK-00102", INFO)], self.sazmi(nal))

    def test_visak_pre_granice_sumnja(self):
        """Pre granice 2 para na 4 unosa, posle 0 na 37: visak -> parovi su SUMNJA."""
        _junski_par(self.t)
        otk(self.t, "OTK-00101", OtkupID="OTK-00103", StanicaID="ST-00002", KooperantID="KOOP-00003",
            VozacID="ST-00002", BrojDokumenta="2/270626", CreatedAt="2026-06-27 09:00:00")
        otk(self.t, "OTK-00103", OtkupID="OTK-00104", BrojDokumenta="2/270626-2", Kolicina=15, KolAmbalaze=6,
            BrutoKg=15, CreatedAt="2026-06-27 09:04:00")
        for i in range(32):
            otk(self.t, "OTK-00005", OtkupID="OTK-%05d" % (201 + i), KooperantID=("KOOP-00001", "KOOP-00002")[i % 2],
                BrojDokumenta="1/060726-%d" % (i + 2), OtpremnicaID="", Kolicina=20 + i / 2.0, KolAmbalaze=8,
                BrutoKg=20 + i / 2.0, CreatedAt="2026-07-06 11:%02d:00" % i)
        nal = self.nalazi()
        hit = [(x.id, x.ozbiljnost) for x in nal if x.kod == ok.K_KOOP]
        self.assertEqual(hit, [("", INFO), ("OTK-00102", SUMNJA), ("OTK-00104", SUMNJA)], self.sazmi(nal))

    def test_ista_stopa_pre_i_posle_bez_nalaza(self):
        """Kooperant koji redovno donosi dvaput zaredom: posle granice ista stopa parova kao
        pre nje -> nema viska, nema nalaza (ni zbirnog INFO, ni parova)."""
        _junski_par(self.t)
        otk(self.t, "OTK-00101", OtkupID="OTK-00103", StanicaID="ST-00002", KooperantID="KOOP-00003",
            VozacID="ST-00002", BrojDokumenta="2/270626", CreatedAt="2026-06-27 09:00:00")
        otk(self.t, "OTK-00103", OtkupID="OTK-00104", BrojDokumenta="2/270626-2", Kolicina=15, KolAmbalaze=6,
            BrutoKg=15, CreatedAt="2026-06-27 09:04:00")
        for i in range(40):
            otk(self.t, "OTK-00005", OtkupID="OTK-%05d" % (201 + i),
                KooperantID=("KOOP-00001", "KOOP-00002")[(i // 2) % 2],
                BrojDokumenta="1/060726-%d" % (i + 2), OtpremnicaID="", Kolicina=20 + i / 2.0, KolAmbalaze=8,
                BrutoKg=20 + i / 2.0, CreatedAt="2026-07-06 11:%02d:00" % i)
        self.nema(self.nalazi(), ok.K_KOOP)

    def test_klasa_I_II_istog_lista_nije_par(self):
        _junski_par(self.t, druga_klasa=True)
        self.nema(self.nalazi(), ok.K_KOOP)

    def test_posle_v270_nije_u_prozoru(self):
        otk(self.t, "OTK-00006", OtkupID="OTK-00007", BrojDokumenta="3/060726-2", OtpremnicaID="", Kolicina=20,
            KolAmbalaze=8, BrutoKg=20, CreatedAt="2026-07-06 10:40:00")
        self.nema(self.nalazi(), ok.K_KOOP)


# ---------------------------------------------------------------- OTK_CENA_TIP_ZAOSTALI / ZAOSTALA_CENA_POSLE_DESELEKCIJE
def _cenovnik(t):
    t["tblCenovnik"] = [
        dict(CenaID="CEN-00001", Datum=D5, VrstaVoca="Malina", SortaVoca="Vilamet", Klasa="I", Cena=250, Stornirano=""),
        dict(CenaID="CEN-00002", Datum=D6, VrstaVoca="Malina", SortaVoca="Vilamet", Klasa="I", Cena=270, Stornirano=""),
    ]


class Cena(Baza):
    def test_najnovija_cena(self):
        _cenovnik(self.t)
        otk(self.t, OtkupID="OTK-00007", BrojDokumenta="1/050726-4", OtpremnicaID="", Kolicina=20, KolAmbalaze=8,
            BrutoKg=20, Cena=270)
        x = self.jedan(self.nalazi(), ok.K_CENA, "OTK-00007", SUMNJA)
        self.assertIn("270 -> 250", x.predlog)

    def test_vezan_blok_ne_poredi_cenovnik(self):
        """Vezan blok nosi 'cenu po otpremnici' (bruto sa PDV) -- ne poredi se sa cenovnikom."""
        _cenovnik(self.t)
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00002", Cena=270)
        self.nema(self.nalazi(), ok.K_CENA)

    def test_cena_storniranog_prvog_bloka(self):
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00001", Stornirano="Da", Cena=260)
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00002", Cena=260)
        otk(self.t, OtkupID="OTK-00007", BrojDokumenta="1/050726-4", Stornirano="", Cena=250)
        x = self.jedan(self.nalazi(), ok.K_CENA, "OTK-00002", SUMNJA)
        self.assertIn("OTK-00001", x.opis)

    def test_vise_cena_bez_storna_nije_nalaz(self):
        """Cena po otpremnici promenjena usred dana: legitimno."""
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00002", Cena=260)
        self.nema(self.nalazi(), ok.K_CENA)

    def test_tip_ambalaze(self):
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00002", TipAmbalaze="6/1")
        x = self.jedan(self.nalazi(), ok.K_CENA, "OTK-00002", SUMNJA)
        self.assertIn("TIP", x.opis)

    def test_tip_bez_gajbi_nije_nalaz(self):
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00002", TipAmbalaze="6/1", KolAmbalaze=0)
        self.nema(self.nalazi(), ok.K_CENA)

    def test_cena_posle_deselekcije(self):
        """OTP-00005 popunjena blokom OTK-00006; sledeci blok (kupina) dobije njenu cenu."""
        otk(self.t, "OTK-00006", OtkupID="OTK-00007", BrojDokumenta="3/060726-2", OtpremnicaID="", BrojZbirne="",
            VrstaVoca="Kupina", SortaVoca="Tornfri", Kolicina=20, KolAmbalaze=8, BrutoKg=20)
        self.jedan(self.nalazi(), ok.K_CDES, "OTK-00007", SUMNJA)

    def test_isti_proizvod_posle_deselekcije_nije_nalaz(self):
        otk(self.t, "OTK-00006", OtkupID="OTK-00007", BrojDokumenta="3/060726-2", OtpremnicaID="", BrojZbirne="",
            Kolicina=20, KolAmbalaze=8, BrutoKg=20)
        self.nema(self.nalazi(), ok.K_CDES)


# ---------------------------------------------------------------- OTK_BROJ_VAN_NIZA_ILI_PRAZAN / PWA_STANICA_I_DATUM
class Broj(Baza):
    def test_prazan(self):
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00002", BrojDokumenta="")
        x = self.jedan(self.nalazi(), ok.K_BROJ, "OTK-00002", SUMNJA)
        self.assertIn("PRAZAN", x.opis)

    def test_prazan_pre_v283_info(self):
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00002", BrojDokumenta="", Datum=JUN27,
                CreatedAt="2026-06-27 08:06:00")
        self.jedan(self.nalazi(), ok.K_BROJ, "OTK-00002", INFO)

    def test_tudja_stanica(self):
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00002", BrojDokumenta="2/050726-5")
        x = self.jedan(self.nalazi(), ok.K_BROJ, "OTK-00002", SUMNJA)
        self.assertIn("TUDJA_STANICA", x.opis)

    def test_tudji_dan(self):
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00002", BrojDokumenta="1/040726-5")
        x = self.jedan(self.nalazi(), ok.K_BROJ, "OTK-00002", SUMNJA)
        self.assertIn("TUDJI_DAN", x.opis)

    def test_nekanonski(self):
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00002", BrojDokumenta="A-17")
        self.jedan(self.nalazi(), ok.K_BROJ, "OTK-00002", INFO)

    def test_duplikat(self):
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00002", BrojDokumenta="1/050726")
        x = self.jedan(self.nalazi(), ok.K_BROJ, "OTK-00001, OTK-00002", SUMNJA)
        self.assertIn("DUPLIKAT", x.opis)

    def test_storno_pa_ponovni_unos_istog_broja(self):
        otk(self.t, "OTK-00002", OtkupID="OTK-00007", Stornirano="Da")
        self.assertEqual(self.sazmi(self.nalazi()), [])

    def test_pwa_stanica_kooperanta(self):
        otk(self.t, OtkupID="OTK-00007", BrojDokumenta="2/050726-7", OtpremnicaID="", Kolicina=20, KolAmbalaze=8,
            BrutoKg=20, SyncSource="PWA", ClientRecordID="c-7")
        nal = self.nalazi()
        x = self.jedan(nal, ok.K_PWA, "OTK-00007", SUMNJA)
        self.assertIn("ST-00002", x.predlog)
        self.nema(nal, ok.K_BROJ)

    def test_pwa_oko_ponoci_info(self):
        otk(self.t, OtkupID="OTK-00007", BrojDokumenta="1/040726-3", OtpremnicaID="", Kolicina=20, KolAmbalaze=8,
            BrutoKg=20, SyncSource="PWA", ClientRecordID="c-7")
        self.jedan(self.nalazi(), ok.K_PWA, "OTK-00007", INFO)


# ---------------------------------------------------------------- OTK_DUPLI_UNOS_BLOKA
class Dupli(Baza):
    def test_nevezan_duplikat(self):
        otk(self.t, OtkupID="OTK-00007", BrojDokumenta="1/050726-4", OtpremnicaID="", CreatedAt="2026-07-05 08:13:00")
        x = self.jedan(self.nalazi(), ok.K_DUPLI, "OTK-00001, OTK-00007", SUMNJA)
        self.assertIn("OTK-00007", x.predlog)

    def test_prepunjena_otpremnica_greska(self):
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00001", Kolicina="60.35", BrutoKg="60.35")
        postavi(self.t, "tblOtpremnica", "OtpremnicaID", "OTP-00001", Kolicina="100.35")
        otk(self.t, OtkupID="OTK-00007", BrojDokumenta="1/050726-4", Novac=0, Isplaceno="")
        self.jedan(self.nalazi(), ok.K_DUPLI, "OTK-00001, OTK-00007", GRESKA)

    def test_dve_isporuke_bez_dokaza_nije_nalaz(self):
        """Isti kooperant, kg i gajbe, ali na razlicitim otpremnicama, sati razmaka, bez dvostruke isplate."""
        otk(self.t, OtkupID="OTK-00007", BrojDokumenta="1/050726-4", OtpremnicaID="OTP-00002",
            BrojZbirne="S1/050726-2", Novac=0, Isplaceno="", CreatedAt="2026-07-05 12:00:00")
        self.nema(self.nalazi(), ok.K_DUPLI)


# ---------------------------------------------------------------- OTK_KULTURAID_PO_VRSTI
def _kulture(t, polka=True):
    t["tblKulture"] = [dict(KulturaID="KUL-00001", VrstaVoca="Malina", SortaVoca="Vilamet")]
    if polka:
        t["tblKulture"].append(dict(KulturaID="KUL-00002", VrstaVoca="Malina", SortaVoca="Polka"))


class Kultura(Baza):
    def test_po_vrsti(self):
        _kulture(self.t)
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00002", SortaVoca="Polka")
        x = self.jedan(self.nalazi(), ok.K_KUL, "OTK-00002", GRESKA)
        self.assertIn("-> KUL-00002", x.predlog)

    def test_fallback_vrsta_sorta(self):
        _kulture(self.t)
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00003", KulturaID="Malina-Vilamet")
        x = self.jedan(self.nalazi(), ok.K_KUL, "OTK-00003", GRESKA)
        self.assertIn("-> KUL-00001", x.predlog)

    def test_jedna_sorta_po_vrsti_nije_nalaz(self):
        _kulture(self.t, polka=False)
        self.nema(self.nalazi(), ok.K_KUL)


# ---------------------------------------------------------------- KOOP_DUPLIKAT_AUTOKREIRANJEM / NEPOSTOJECI_ID
class Kooperant(Baza):
    def test_duplikat_imena(self):
        self.t["tblKooperanti"].append(dict(KooperantID="KOOP-00005", Ime="Pera", Prezime="Peric",
                                            StanicaID="ST-00001", Aktivan="Da"))
        self.jedan(self.nalazi(), ok.K_KDUP, "KOOP-00001, KOOP-00005", SUMNJA, tabela="tblKooperanti")

    def test_kvacice_i_razmaci_isti_kljuc(self):
        self.t["tblKooperanti"].append(dict(KooperantID="KOOP-00005", Ime="  pera ", Prezime="Peri\u0107",
                                            StanicaID="ST-00001", Aktivan="Da"))
        self.jedan(self.nalazi(), ok.K_KDUP, "KOOP-00001, KOOP-00005", SUMNJA, tabela="tblKooperanti")

    def test_razliciti_bpg_nije_nalaz(self):
        """Razlicite osobe istog imena (otac/sin) sa svojim BPG-om."""
        postavi(self.t, "tblKooperanti", "KooperantID", "KOOP-00001", BPGBroj="111")
        self.t["tblKooperanti"].append(dict(KooperantID="KOOP-00005", Ime="Pera", Prezime="Peric",
                                            StanicaID="ST-00001", Aktivan="Da", BPGBroj="222"))
        self.nema(self.nalazi(), ok.K_KDUP)

    def test_imenjak_druge_stanice(self):
        self.t["tblKooperanti"].append(dict(KooperantID="KOOP-00006", Ime="Zika", Prezime="Zikic",
                                            StanicaID="ST-00001", Aktivan="Da"))
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00002", KooperantID="KOOP-00003")
        nal = self.nalazi()
        hit = [(x.tabela, x.id, x.ozbiljnost) for x in nal if x.kod == ok.K_KDUP]
        self.assertEqual(hit, [("tblKooperanti", "KOOP-00003, KOOP-00006", SUMNJA),
                               ("tblOtkup", "OTK-00002", SUMNJA)], self.sazmi(nal))

    def test_nepostojeci_kooperant(self):
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00002", KooperantID="mladji")
        x = self.jedan(self.nalazi(), ok.K_KNEP, "OTK-00002", GRESKA)
        self.assertIn("zagrade", x.opis)


# ---------------------------------------------------------------- OTK_GAJBE_KG_NEPROVERENO
class Gajbe(Baza):
    def test_bez_gajbi_uz_manjak_otpremnice(self):
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00002", KolAmbalaze=0)
        x = self.jedan(self.nalazi(), ok.K_GAJ, "OTK-00002", SUMNJA)
        self.assertIn("0 -> 16", x.predlog)

    def test_bez_gajbi_nevezan_nije_nalaz(self):
        """Kooperant doneo robu u svojoj ambalazi: bez otpremnice nema dokaza."""
        otk(self.t, OtkupID="OTK-00007", BrojDokumenta="1/050726-4", OtpremnicaID="", Kolicina=20, KolAmbalaze=0,
            BrutoKg=20)
        self.nema(self.nalazi(), ok.K_GAJ)

    def test_previse_kg_po_gajbi(self):
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00005", KolAmbalaze=7)
        self.jedan(self.nalazi(), ok.K_GAJ, "OTK-00005", SUMNJA)

    def test_zamenjena_polja(self):
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00006", Kolicina=36, KolAmbalaze=90, BrutoKg=36)
        x = self.jedan(self.nalazi(), ok.K_GAJ, "OTK-00006", SUMNJA)
        self.assertIn("Zamenjena", x.predlog)

    def test_polu_pune_gajbe_nije_nalaz(self):
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00005", KolAmbalaze=45)
        self.nema(self.nalazi(), ok.K_GAJ)


# ---------------------------------------------------------------- OTK_AKTIVAN_BEZ_ZIVE_OTPREMNICE
class Ziva(Baza):
    def test_stornirana_otpremnica(self):
        postavi(self.t, "tblOtpremnica", "OtpremnicaID", "OTP-00002", Stornirano="Da")
        x = self.jedan(self.nalazi(), ok.K_ZIVA, "OTK-00003", GRESKA)
        self.assertIn("stornirana", x.opis)

    def test_nepostojeca_otpremnica(self):
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00002", OtpremnicaID="OTP-00099")
        x = self.jedan(self.nalazi(), ok.K_ZIVA, "OTK-00002", GRESKA)
        self.assertIn("ne postoji", x.opis)

    def test_ponovo_unet_blok_predlog_storno(self):
        postavi(self.t, "tblOtpremnica", "OtpremnicaID", "OTP-00002", Stornirano="Da")
        otp(self.t, "OTP-00002", OtpremnicaID="OTP-00006", Stornirano="")
        otk(self.t, "OTK-00003", OtkupID="OTK-00007", BrojDokumenta="1/050726-4", OtpremnicaID="OTP-00006")
        x = self.jedan(self.nalazi(), ok.K_ZIVA, "OTK-00003", GRESKA)
        self.assertIn("storno", x.predlog)

    def test_oslobodjen_stornom(self):
        postavi(self.t, "tblOtpremnica", "OtpremnicaID", "OTP-00002", Stornirano="Da")
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00003", OtpremnicaID="", BrojZbirne="")
        self.t["tblStornoVeze"].append(dict(CorrectionID="COR-00001", Mode="DUPLI_FANTOM", Status="COMPLETED",
                                            OldDocType="Otpremnica", OldDocID="OTP-00002", OldBroj="1/050726-2",
                                            CreatedAt="2026-07-05 12:00:00", CreatedBy="op"))
        self.jedan(self.nalazi(), ok.K_ZIVA, "OTK-00003", SUMNJA)

    def test_nikad_vezan_nije_nalaz(self):
        otk(self.t, OtkupID="OTK-00007", BrojDokumenta="1/050726-4", OtpremnicaID="", BrojZbirne="", Kolicina=20,
            KolAmbalaze=8, BrutoKg=20)
        self.nema(self.nalazi(), ok.K_ZIVA)


# ---------------------------------------------------------------- SYNC_AUTO_OTPREMNICA_ZA_NEVEZANE
def _sync_otpremnica(t, pwa=False, manjak=True):
    """Blok koji je pripadao OTP-00004 (manjak 30 kg) sync spakovao u svoju otpremnicu OTP-00006 bez prijemnice."""
    if manjak:
        postavi(t, "tblOtpremnica", "OtpremnicaID", "OTP-00004", Kolicina=100, KolAmbalaze=40)
    otk(t, "OTK-00005", OtkupID="OTK-00007", KooperantID="KOOP-00001", BrojDokumenta="1/060726-2",
        OtpremnicaID="OTP-00006", BrojZbirne="S1/060726-2", Kolicina=30, KolAmbalaze=12, BrutoKg=30,
        CreatedAt="2026-07-06 10:05:00", SyncSource="PWA" if pwa else "", ClientRecordID="c-7" if pwa else "")
    otp(t, "OTP-00004", OtpremnicaID="OTP-00006", BrojOtpremnice="1/060726-2", BrojZbirne="S1/060726-2",
        Kolicina=30, KolAmbalaze=12, Cena=250, BrutoKg=0, CreatedAt="2026-07-06 12:00:00")
    kopiraj(t, "tblZbirna", "ZbirnaID", "ZBR-00004", ZbirnaID="ZBR-00006", BrojZbirne="S1/060726-2",
            UkupnoKolicina=30, UkupnoAmbalaze=12)


class Sync(Baza):
    def test_sync_otpremnica_bez_prijemnice(self):
        _sync_otpremnica(self.t)
        x = self.jedan(self.nalazi(), ok.K_SYNC, "OTP-00006", SUMNJA, tabela="tblOtpremnica")
        self.assertIn("OTP-00004", x.predlog)

    def test_pwa_tok_nije_nalaz(self):
        """Legitiman PWA tok: clanovi auto-otpremnice su PWA redovi."""
        _sync_otpremnica(self.t, pwa=True)
        self.nema(self.nalazi(), ok.K_SYNC)

    def test_bez_manjka_druge_otpremnice_nije_nalaz(self):
        _sync_otpremnica(self.t, manjak=False)
        self.nema(self.nalazi(), ok.K_SYNC)


# ---------------------------------------------------------------- nezavisna recenzija (lazni pozitivi, varijante)
BEZ_AUDITA = {ime: [c for c in kol if c not in ("CreatedAt", "CreatedBy", "ModifiedAt", "ModifiedBy", "VremeUnosa")]
              for ime, kol in fixture.KOLONE.items()}


class Recenzija(Baza):
    def test_rucna_otpremnica_bez_audita_nije_sync(self):
        """Bez audit kolona popunjena rucna otpremnica (cena iz cenovnika = cena blokova, bez bruto-a) ima ceo
        otisak sync-a; ima prijemnicu, pa vezani blokovi NE idu na poredjenje sa cenovnikom."""
        for r in self.t["tblOtpremnica"]:
            r.update(Cena=250, BrutoKg="")
        self.t["tblCenovnik"] = [
            dict(CenaID="CEN-00001", Datum=D5, VrstaVoca="Malina", SortaVoca="Vilamet", Klasa="I", Cena=240,
                 Stornirano=""),
            dict(CenaID="CEN-00002", Datum=D7, VrstaVoca="Malina", SortaVoca="Vilamet", Klasa="I", Cena=250,
                 Stornirano="")]
        self.assertEqual(self.sazmi(self.nalazi(kolone=BEZ_AUDITA)), [])

    def test_dve_isporuke_na_dve_otpremnice_nije_dupli(self):
        """Dve stvarne isporuke istog kooperanta (50 kg / 20 gajbi) na dve rucne otpremnice, druga sa
        prijemnicom: druga samo lici na sync -- nije dokaz duplog unosa."""
        otp(self.t, "OTP-00002", OtpremnicaID="OTP-00006", BrojOtpremnice="1/050726-3", BrojZbirne="S1/050726-3",
            Cena=250, BrutoKg="")
        kopiraj(self.t, "tblZbirna", "ZbirnaID", "ZBR-00002", ZbirnaID="ZBR-00006", BrojZbirne="S1/050726-3")
        kopiraj(self.t, "tblPrijemnica", "PrijemnicaID", "PRJ-00002", PrijemnicaID="PRJ-00006",
                BrojPrijemnice="1/050726-9", BrojZbirne="S1/050726-3")
        otk(self.t, "OTK-00003", OtkupID="OTK-00007", BrojDokumenta="1/050726-4", OtpremnicaID="OTP-00006",
            BrojZbirne="S1/050726-3", CreatedAt="2026-07-05 15:00:00", ModifiedAt="")
        self.nema(self.nalazi(kolone=BEZ_AUDITA), ok.K_DUPLI)

    def test_zastareli_niz_prekinut_unosom_istog_dana(self):
        """Posle popune OTP-00004 nevezan blok unet ISTOG dana prekida niz: kasniji nevezan unos sutradan
        nije dokaz zaostalog datuma te otpremnice (ispravljen predikat skeptika: 'uzastopni')."""
        otk(self.t, "OTK-00005", OtkupID="OTK-00007", KooperantID="KOOP-00001", BrojDokumenta="1/060726-2",
            OtpremnicaID="", BrojZbirne="", Kolicina=20, KolAmbalaze=8, BrutoKg=20,
            CreatedAt="2026-07-06 11:00:00", VremeUnosa="2026-07-06 11:00:00")
        otk(self.t, "OTK-00005", OtkupID="OTK-00008", KooperantID="KOOP-00001", BrojDokumenta="1/060726-3",
            OtpremnicaID="", BrojZbirne="", Kolicina=21, KolAmbalaze=8, BrutoKg=21,
            CreatedAt="2026-07-07 09:00:00", VremeUnosa="2026-07-07 09:00:00")
        self.nema(self.nalazi(), ok.K_ZAST)

    def test_zastareli_niz_prekinut_unosom_drugog_datuma_druge_stanice(self):
        """txtDatum je jedno polje forme: izmedju L i O snimljen je blok druge stanice sa Datum 07.07, pa je
        06.07 na O ponovo upisan rucno -- nije zaostali kontekst."""
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00005", CreatedAt="2026-07-07 08:50:00")
        otk(self.t, "OTK-00006", OtkupID="OTK-00007", Datum=D7, BrojDokumenta="3/070726", OtpremnicaID="",
            BrojZbirne="", Kolicina=20, KolAmbalaze=8, BrutoKg=20, CreatedAt="2026-07-07 09:00:00")
        otk(self.t, "OTK-00005", OtkupID="OTK-00008", KooperantID="KOOP-00001", BrojDokumenta="1/060726-2",
            OtpremnicaID="", Kolicina=25, KolAmbalaze=10, BrutoKg=25, CreatedAt="2026-07-07 09:10:00")
        self.nema(self.nalazi(), ok.K_ZAST)

    def test_zastareli_vezan_za_sync_otpremnicu(self):
        """Varijanta (iv): zastareli blok nije ostao nevezan -- sync ga je spakovao u svoju otpremnicu
        OTP-00006 (zbirna bez prijemnice). Dokaz (a): nosi zbirnu S1/060726 popunjene OTP-00004."""
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00005", CreatedAt="2026-07-07 08:50:00")
        otk(self.t, "OTK-00005", OtkupID="OTK-00007", KooperantID="KOOP-00001", BrojDokumenta="1/060726-2",
            OtpremnicaID="OTP-00006", BrojZbirne="S1/060726", Kolicina=20, KolAmbalaze=8, BrutoKg=20,
            CreatedAt="2026-07-07 09:00:00", VremeUnosa="2026-07-07 09:00:00")
        otp(self.t, "OTP-00004", OtpremnicaID="OTP-00006", BrojOtpremnice="1/060726-2", BrojZbirne="S1/060726-2",
            Kolicina=20, KolAmbalaze=8, Cena=250, BrutoKg="", CreatedAt="2026-07-07 12:00:00")
        kopiraj(self.t, "tblZbirna", "ZbirnaID", "ZBR-00004", ZbirnaID="ZBR-00006", BrojZbirne="S1/060726-2",
                UkupnoKolicina=20, UkupnoAmbalaze=8)
        nal = self.nalazi()
        x = self.jedan(nal, ok.K_ZAST, "OTK-00007", SUMNJA)
        self.assertIn("sync", x.opis)
        self.nema(nal, ok.K_BZ)

    def test_pwa_stanica_kooperanta_veza_je_sumnja(self):
        """PWA otkup sa StanicaID maticnog OM kooperanta, vezan za otpremnicu stanice iz broja: veza moze biti
        tacna (pogresna je StanicaID otkupa) -- SUMNJA, ne GRESKA."""
        otk(self.t, "OTK-00004", OtkupID="OTK-00007", KooperantID="KOOP-00001", StanicaID="ST-00001",
            BrojDokumenta="2/050726-2", Kolicina=10, KolAmbalaze=4, BrutoKg=10, SyncSource="PWA",
            ClientRecordID="c-7")
        postavi(self.t, "tblOtpremnica", "OtpremnicaID", "OTP-00003", Kolicina=90, KolAmbalaze=36)
        nal = self.nalazi()
        x = self.jedan(nal, ok.K_VEZA, "OTK-00007", SUMNJA)
        self.assertIn("PWA", x.opis)
        self.jedan(nal, ok.K_PWA, "OTK-00007", SUMNJA)

    def test_cena_posle_deselekcije_cenovnik_novog_proizvoda(self):
        """Cenovnik ima cenu kupine: promena proizvoda okida AutoFillCenaOtkup, pa cena otpremnice ne moze da
        zaostane -- ista cena je rucni unos, nije nalaz."""
        self.t["tblCenovnik"] = [
            dict(CenaID="CEN-00001", Datum=D5, VrstaVoca="Kupina", SortaVoca="Tornfri", Klasa="I", Cena=180,
                 Stornirano="")]
        otk(self.t, "OTK-00006", OtkupID="OTK-00007", BrojDokumenta="3/060726-2", OtpremnicaID="", BrojZbirne="",
            VrstaVoca="Kupina", SortaVoca="Tornfri", Kolicina=20, KolAmbalaze=8, BrutoKg=20)
        self.nema(self.nalazi(), ok.K_CDES)

    def test_panel_b_rucna_otpremnica_istog_broja_druge_kg(self):
        """Hladnjaca: nevezan blok 30 kg i otpremnica istog broja 80 kg bez blokova -- nije lanac tog bloka."""
        hladnjaca(self.t)
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00004", OtpremnicaID="", Kolicina=30, KolAmbalaze=12,
                BrutoKg=30)
        self.nema(self.nalazi(), ok.K_PANEL)

    def test_vezan_pre_storna_ponovljenog_broja_nije_panel(self):
        """Storno pa ponovni unos istog broja: OTP-00002 storn. u 09:00, OTP-00006 (isti broj) uneta, blok
        OTK-00007 vezan u 10:00, OTP-00006 storn. u 11:00. Vreme storna je 11:00, ne 09:00."""
        postavi(self.t, "tblOtpremnica", "OtpremnicaID", "OTP-00002", Stornirano="Da")
        otp(self.t, "OTP-00002", OtpremnicaID="OTP-00006", Stornirano="Da", CreatedAt="2026-07-05 09:30:00",
            ModifiedAt="2026-07-05 11:00:00")
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00003", Stornirano="Da")
        otk(self.t, "OTK-00003", OtkupID="OTK-00007", BrojDokumenta="1/050726-4", OtpremnicaID="OTP-00006",
            Stornirano="", CreatedAt="2026-07-05 10:00:00")
        for i, (old, t_) in enumerate((("OTP-00002", "09:00"), ("OTP-00006", "11:00")), 1):
            self.t["tblStornoVeze"].append(dict(
                CorrectionID="COR-%05d" % i, Mode="STORNO", Status="Done", OldDocType="Otpremnica", OldDocID=old,
                OldBroj="1/050726-2", CreatedAt="2026-07-05 %s:00" % t_))
        nal = self.nalazi()
        self.nema(nal, ok.K_PANEL)
        self.jedan(nal, ok.K_ZIVA, "OTK-00007", GRESKA)

    def test_cena_lanca_hladnjace_najnovija(self):
        """Hladnjacki blok je vezan za svoj lanac, a cena mu dolazi iz txtCena/cenovnika: poredi se."""
        hladnjaca(self.t)
        _cenovnik(self.t)
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00004", Cena=270)
        x = self.jedan(self.nalazi(), ok.K_CENA, "OTK-00004", SUMNJA)
        self.assertIn("NAJNOVIJA", x.opis)

    def test_sync_otisak_sa_bruto_je_rucna(self):
        """BrutoKg > 0 pise samo rucni unos (sync zove SaveOtpremnica_TX bez bruto-a): nije sync otpremnica."""
        _sync_otpremnica(self.t)
        postavi(self.t, "tblOtpremnica", "OtpremnicaID", "OTP-00006", BrutoKg=30)
        self.nema(self.nalazi(), ok.K_SYNC)

    def test_losi_podaci_svuda(self):
        """Prazan ID, None/nemoguc datum, tekst umesto broja, prazni maticni redovi: bez pada."""
        t = self.t
        t["tblOtkup"].append(dict(OtkupID="", Datum=None, Kolicina="abc", KolAmbalaze="x", Cena="?",
                                  BrojDokumenta="1/990726", StanicaID="", OtpremnicaID="OTP-00001"))
        t["tblOtkup"].append(dict(OtkupID="OTK-00099", Datum="31.06.2026", Kolicina="1e5", KolAmbalaze="-3",
                                  Cena="NaN", BrojDokumenta="1/310626", StanicaID="ST-00001", VozacID="hladnjaca",
                                  VremeUnosa="99999999", OtpremnicaID="OTP-00098"))
        t["tblOtpremnica"].append(dict(OtpremnicaID="OTP-00099", Datum="", Kolicina="nan", BrojOtpremnice="",
                                       StanicaID=None))
        t["tblStornoVeze"].append(dict(Mode="DUPLI_FANTOM", OldDocType="Zbirna", OldBroj="S1/050726",
                                       CreatedAt="xx"))
        t["tblKooperanti"].append(dict(KooperantID="", Ime=None, Prezime=None))
        t["tblVozaci"].append(dict(VozacID="VOZ-00009", Ime="A - B", Prezime="(x)"))
        _kulture(t)
        _cenovnik(t)
        self.nalazi()
        self.nalazi(kolone=BEZ_AUDITA)

    def test_minimalne_kolone_jedan_preskoceno_po_kodu(self):
        kol = {"tblOtkup": ["OtkupID", "Datum", "Kolicina"], "tblOtpremnica": ["OtpremnicaID"]}
        nal = self.nalazi(kolone=kol)
        kodovi = [x.kod for x in nal]
        self.assertTrue(kodovi)
        self.assertTrue(all(kd.endswith("_PRESKOCENO") for kd in kodovi), kodovi)
        self.assertEqual(len(kodovi), len(set(kodovi)))


# ---------------------------------------------------------------- otpornost
class Otpornost(Baza):
    def test_kolona_fali_preskoceno(self):
        kol = {"tblOtkup": [c for c in fixture.KOLONE["tblOtkup"] if c != "OtpremnicaID"]}
        nal = self.nalazi(kolone=kol)
        kodovi = {x.kod for x in nal}
        self.assertIn(ok.K_ZAST + "_PRESKOCENO", kodovi)
        self.assertIn(ok.K_VEZA + "_PRESKOCENO", kodovi)
        self.assertTrue(all(x.ozbiljnost == INFO for x in nal if x.kod.endswith("_PRESKOCENO")))
        self.assertEqual(len([x for x in nal if x.kod == ok.K_ZAST + "_PRESKOCENO"]), 1)

    def test_prazne_i_nepostojece_tabele(self):
        self.t["tblOtpremnica"] = []
        del self.t["tblZbirna"]
        self.nalazi()
        t = fixture.kopija(fixture.osnovna())
        del t["tblOtkup"]
        self.nalazi(t)

    def test_losi_podaci(self):
        postavi(self.t, "tblOtkup", "OtkupID", "OTK-00002", Datum="abc", Kolicina="x", KolAmbalaze=None, Cena="",
                CreatedAt="nije vreme", VremeUnosa="45000.5")
        postavi(self.t, "tblOtpremnica", "OtpremnicaID", "OTP-00001", Datum=None, Kolicina="", BrojZbirne=None)
        _cenovnik(self.t)
        self.t["tblCenovnik"].append(dict(CenaID="CEN-00003", Datum="", VrstaVoca="Malina", SortaVoca="", Klasa="I",
                                          Cena="abc"))
        _kulture(self.t)
        self.t["tblKulture"].append(dict(KulturaID="", VrstaVoca="", SortaVoca=""))
        self.nalazi()


if __name__ == "__main__":
    unittest.main()
