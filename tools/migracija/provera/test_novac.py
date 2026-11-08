"""Testovi modula provera/novac.py: polazi se od ciste fixture sveske, podmetne se
trag koji je ostavljao stari unos novca i trazi kod na tacnom ID-ju i ozbiljnosti;
uz svaku klasu ide i njen glavni legitimni slican slucaj (bez nalaza)."""

import unittest
from datetime import date
from decimal import Decimal

from provera import fixture
from provera.fixture import D5, D6, KUPAC
from provera.pokretac import pokreni_nad_tabelama
from provera.test_zajednicko import tabele_iz
from provera.zajednicko import GRESKA, INFO, SUMNJA

D7 = date(2026, 7, 7)
D10 = date(2026, 7, 10)
KOL = {"tblBankaImport": fixture.KOLONE["tblBankaImport"] + ["SvrhaPlacanja", "DatumIzvoda"],
       "tblPartnerMap": ["BankaName", "PartnerID", "EntitetTip", "OMID"],
       "tblStornoZurnal": ["ZurnalID", "OperationID", "Timestamp", "DocType", "Broj", "Tabela", "RowID", "Kolona",
                           "StaraVrednost", "NovaVrednost"]}

K_DUPLA = "NOV_DUPLA_ISPLATA_RUCNO_I_IZVOD"
K_OSIR = "NOV_ISPLATA_OSIROCENA_STORNOM"
K_ODVZ = "NOV_ISPLATA_ODVEZANA_OD_STORNIRANOG_OTKUPA"
K_PREPL = "NOV_PREPLATA_OTKUPA"
K_ISPL = "OTK_ISPLACENO_POSLE_STORNA_NOVCA"
K_VEZA = "NOV_VEZA_NA_TUDJI_ILI_STORNIRAN_DOKUMENT"
K_FAKST = "FAK_UPLATA_NA_STORNIRANU_FAKTURU"
K_UPLK = "NOV_UPLATA_KUPCA_NA_STORNIRANU_ILI_PREPLACENU_FAKTURU"
K_ODVF = "NOV_UPLATA_KUPCA_ODVEZANA_STORNOM_FAKTURE"
K_KUPOM = "NOV_IZVOD_RUCNO_KUPAC_OM_POGRESNO"
K_KOOP = "NOV_IZVOD_POGRESAN_KOOPERANT"
K_TIP = "NOV_TIP_ISPLATE_PO_PODRAZUMEVANOJ"
K_DAT = "NOV_ZASTAREO_DATUM_FORME"
K_LOK = "NOV_IZVOD_DATUM_LOKAL"
K_STLAN = "NOV_STORNO_NOVCA_BEZ_LANCA"
K_RAZD = "NOV_AVANS_RASPODELA_RAZDVOJENA_STORNOM"
K_DUPRU = "NOV_DUPLI_RUCNI_UNOS"
K_VAN = "IZVOD_VAN_KNJIGE"
K_BEZP = "IZVOD_BEZ_POZIVA_ODE_U_AVANS"
K_REV = "REVERS_STORNIRAN_NOVAC_OSTAO"
K_UPLOM = "UPLATA_OM_NEVIDLJIVA_U_SALDU"
K_POZ = "POZICIONI_POMAK_TBLNOVAC"
K_DED = "IZVOD_DEDUPE_NEPROVERLJIV"


def nova():
    return fixture.kopija(fixture.osnovna())


def r(t, ime, id_):
    return fixture.red(t, ime, fixture.KOLONE[ime][0], id_)


def nov(t, id_, **kw):
    red = dict(NovacID=id_, BrojDokumenta="", Datum=D6, Partner="", PartnerID="", EntitetTip="OM", OMID="",
               KooperantID="", FakturaID="", VrstaVoca="Malina", Tip="", Uplata=0, Isplata=0, Napomena="",
               Stornirano="", OtkupID="", CreatedAt="", CreatedBy="op")
    red.update(kw)
    t["tblNovac"].append(red)
    return red


def izvod(t, id_, bim_id, match="Kooperant", **kw):
    kw.setdefault("BrojDokumenta", "12")
    return nov(t, id_, Napomena="BIM:%s; Match:%s" % (bim_id, match), **kw)


def bim(t, id_, **kw):
    red = dict(BankaImportID=id_, BrojRacuna="160-1", BrojDokumenta="12", DatumTransakcije="06.07.2026",
               Partner="", PartnerKonto="", PozivNaBroj="", Isplata=0, Uplata=0, Obradjeno="Da", Stornirano="")
    red.update(kw)
    t["tblBankaImport"].append(red)
    return red


def otk(t, base, id_, **kw):
    novi = dict(r(t, "tblOtkup", base))
    novi.update(OtkupID=id_, Novac=0, PrimalacNovca="", Isplaceno="", DatumIsplate="", Stornirano="")
    novi.update(kw)
    t["tblOtkup"].append(novi)
    return novi


def fak(t, id_, **kw):
    red = dict(FakturaID=id_, BrojFakture="1/2026", Datum=D5, KupacID=KUPAC, Iznos=10000, Status="Neplaceno",
               DatumPlacanja="", Stornirano="")
    red.update(kw)
    t["tblFakture"].append(red)
    return red


def kupac_upl(t, id_, **kw):
    kw.setdefault("Tip", "KupciUplata")
    return nov(t, id_, EntitetTip="Kupac", PartnerID=KUPAC, Partner="Hladnjaca Test", **kw)


class Osnova(unittest.TestCase):
    def nalazi(self, t, kolone=None):
        kol = dict(KOL)
        kol.update(kolone or {})
        ctx = pokreni_nad_tabelama(tabele_iz(t, kol), moduli=("novac",))
        self.assertEqual([x.opis for x in ctx.nalazi if x.kod == "MODUL_PAO"], [])
        return ctx.nalazi

    def jedan(self, nn, kod, id_, ozb):
        x = [y for y in nn if y.kod == kod]
        self.assertEqual(len(x), 1, [(y.kod, y.id, y.ozbiljnost, y.opis) for y in nn])
        self.assertEqual((x[0].id, x[0].ozbiljnost), (id_, ozb), x[0].opis)
        return x[0]

    def nema(self, nn, kod, ozb=None):
        self.assertEqual([(y.id, y.ozbiljnost, y.opis) for y in nn if y.kod == kod
                          and (ozb is None or y.ozbiljnost in ozb)], [])


class CistaSveska(Osnova):
    def test_nula_nalaza(self):
        self.assertEqual([(x.kod, x.id, x.opis[:120]) for x in self.nalazi(nova())], [])


# ---------------------------------------------------------------------------
def _rucni_virman_004(t):
    """NOV-00002 kao rucni Ulaz OM virman za blok OTK-00004 (KOOP-00003, 20000)."""
    x = r(t, "tblNovac", "NOV-00002")
    x.update(EntitetTip="OM", PartnerID="ST-00002", Partner="Stanica 2", Napomena="Zika Zikic")
    return x


class DuplaRucnoIIzvod(Osnova):
    def test_jak_pogodak_isti_blok(self):
        t = nova()
        _rucni_virman_004(t)
        bim(t, "BIM-00001", DatumTransakcije="07.07.2026", Partner="Zika Zikic", PozivNaBroj="2/050726",
            Isplata=20000)
        izvod(t, "NOV-00010", "BIM-00001", Datum=D7, Tip="VirmanAvansKoop", KooperantID="KOOP-00003",
              PartnerID="ST-00002", OMID="ST-00002", Isplata=20000)
        x = self.jedan(self.nalazi(t), K_DUPLA, "NOV-00002", GRESKA)
        self.assertIn("JAK POGODAK", x.opis)
        self.assertIn("BIM-00001", x.predlog)

    def test_avans_potrosen_na_sledeci_otkup(self):
        t = nova()
        _rucni_virman_004(t)
        otk(t, "OTK-00004", "OTK-00007", Datum=D10, BrojDokumenta="2/100726", Kolicina=40)   # 10000
        bim(t, "BIM-00001", DatumTransakcije="07.07.2026", PozivNaBroj="2/050726", Isplata=20000)
        izvod(t, "NOV-00010", "BIM-00001", Datum=D7, Tip="VirmanAvansKoop", KooperantID="KOOP-00003",
              PartnerID="ST-00002", OMID="ST-00002", Isplata=10000)
        nov(t, "NOV-00011", BrojDokumenta="12", Datum=D7, Tip="VirmanAvansKoop", KooperantID="KOOP-00003",
            EntitetTip="Kooperant", PartnerID="ST-00002", OMID="ST-00002", Isplata=10000, Napomena="Avans raspodela",
            OtkupID="OTK-00007")
        x = self.jedan(self.nalazi(t), K_DUPLA, "NOV-00002", GRESKA)
        self.assertIn("OTK-00007", x.opis)

    def test_bez_poziva_sumnja(self):
        t = nova()
        _rucni_virman_004(t)
        bim(t, "BIM-00001", DatumTransakcije="07.07.2026", Isplata=20000)
        izvod(t, "NOV-00010", "BIM-00001", Datum=D7, Tip="VirmanAvansKoop", KooperantID="KOOP-00003",
              PartnerID="ST-00002", OMID="ST-00002", Isplata=20000)
        self.jedan(self.nalazi(t), K_DUPLA, "NOV-00002", SUMNJA)

    def test_om_varijanta(self):
        t = nova()
        nov(t, "NOV-00010", BrojDokumenta="OM-7", Datum=D6, Tip="KesFirmaOtkupac", PartnerID="ST-00001",
            OMID="ST-00001", Isplata=50000)
        bim(t, "BIM-00001", DatumTransakcije="07.07.2026", Isplata=50000)
        izvod(t, "NOV-00011", "BIM-00001", match="OM", Datum=D7, Tip="KesFirmaOtkupac", PartnerID="ST-00001",
              OMID="ST-00001", Isplata=50000)
        self.jedan(self.nalazi(t), K_DUPLA, "NOV-00010", SUMNJA)

    def test_legitimno_blok_na_dve_polovine_nije_greska(self):
        """Izvod je platio SAM blok (VirmanFirmaKoop sa OtkupID: blok bio otvoren pri mapiranju), a blok nije
        preplacen: dve stvarne isplate po pola su moguce -> najvise SUMNJA."""
        t = nova()
        _rucni_virman_004(t)["Isplata"] = 10000
        bim(t, "BIM-00001", DatumTransakcije="07.07.2026", PozivNaBroj="2/050726", Isplata=10000)
        izvod(t, "NOV-00010", "BIM-00001", Datum=D7, Tip="VirmanFirmaKoop", KooperantID="KOOP-00003",
              PartnerID="ST-00002", OMID="ST-00002", Isplata=10000, OtkupID="OTK-00004")
        self.jedan(self.nalazi(t), K_DUPLA, "NOV-00002", SUMNJA)

    def test_blok_preplacen_greska(self):
        """Isti par, ali izvod je platio ceo blok pa je blok preplacen: podaci dokazuju duplikat."""
        t = nova()
        _rucni_virman_004(t)
        bim(t, "BIM-00001", DatumTransakcije="07.07.2026", PozivNaBroj="2/050726", Isplata=20000)
        izvod(t, "NOV-00010", "BIM-00001", Datum=D7, Tip="VirmanFirmaKoop", KooperantID="KOOP-00003",
              PartnerID="ST-00002", OMID="ST-00002", Isplata=20000, OtkupID="OTK-00004")
        self.jedan(self.nalazi(t), K_DUPLA, "NOV-00002", GRESKA)

    def test_legitimno_razliciti_blokovi(self):
        t = nova()
        _rucni_virman_004(t)
        otk(t, "OTK-00004", "OTK-00007", Datum=D6, BrojDokumenta="2/060726")      # drugi blok, 20000
        bim(t, "BIM-00001", DatumTransakcije="07.07.2026", PozivNaBroj="2/060726", Isplata=20000)
        izvod(t, "NOV-00010", "BIM-00001", Datum=D7, Tip="VirmanFirmaKoop", KooperantID="KOOP-00003",
              PartnerID="ST-00002", OMID="ST-00002", Isplata=20000, OtkupID="OTK-00007")
        self.nema(self.nalazi(t), K_DUPLA)


# ---------------------------------------------------------------------------
class OsirocenaStornom(Osnova):
    def test_prevezi_na_naslednika_zurnal(self):
        t = nova()
        r(t, "tblOtkup", "OTK-00004")["Stornirano"] = "Da"
        r(t, "tblNovac", "NOV-00002")["OtkupID"] = ""                   # ResetNovacOtkupLink
        otk(t, "OTK-00004", "OTK-00007")                                # ponovo unet, isti broj, bez isplate
        t["tblStornoZurnal"] = [dict(ZurnalID="Z-1", OperationID="OP-1", Timestamp="2026-07-07 10:00:00",
                                     DocType="Otkup", Broj="2/050726", Tabela="tblNovac", RowID="NOV-00002",
                                     Kolona="OtkupID", StaraVrednost="OTK-00004", NovaVrednost="")]
        x = self.jedan(self.nalazi(t), K_OSIR, "NOV-00002", SUMNJA)
        self.assertIn("PREVEZI", x.opis)
        self.assertIn("zurnal", x.opis)
        self.assertIn("-> OTK-00007", x.predlog)

    def test_pravo_sirce_bez_naslednika(self):
        t = nova()
        r(t, "tblOtkup", "OTK-00004")["Stornirano"] = "Da"
        r(t, "tblNovac", "NOV-00002")["OtkupID"] = ""
        x = self.jedan(self.nalazi(t), K_OSIR, "NOV-00002", SUMNJA)
        self.assertIn("PRAVO_SIROCE", x.opis)

    def test_om_ulaz_broj_nije_broj_bloka(self):
        """Ulaz OM nosi broj OM ulaza / reversa (isti oblik x/ddmmyy); poklapanje sa brojem storniranog bloka
        nije dokaz -- O se trazi samo slabim putem (d)."""
        t = nova()
        r(t, "tblOtkup", "OTK-00003")["Stornirano"] = "Da"              # KOOP-00001, 1/050726-3, 05.07.
        nov(t, "NOV-00010", BrojDokumenta="1/050726-3", Datum=D5, Tip="KesOtkupacKoop", PartnerID="ST-00001",
            OMID="ST-00001", KooperantID="KOOP-00001", Isplata=5000, Napomena="Pera Peric")
        r(t, "tblOtkup", "OTK-00001")["Stornirano"] = "Da"              # drugi stornirani blok istog dana
        r(t, "tblNovac", "NOV-00001").update(OtkupID="", Stornirano="Da")
        x = [y for y in self.nalazi(t) if y.kod == K_OSIR and y.id == "NOV-00010"]
        self.assertEqual(len(x), 1)
        self.assertNotIn("isti broj dokumenta", x[0].opis)
        self.assertIn("NEPOZNAT", x[0].opis)

    def test_veza_na_storniran_otkup(self):
        t = nova()
        r(t, "tblOtkup", "OTK-00004")["Stornirano"] = "Da"              # OtkupID na novcu ostao
        self.jedan(self.nalazi(t), K_OSIR, "NOV-00002", GRESKA)

    def test_dupla_kes(self):
        t = nova()
        r(t, "tblOtkup", "OTK-00001")["Stornirano"] = "Da"
        r(t, "tblNovac", "NOV-00001")["OtkupID"] = ""
        otk(t, "OTK-00001", "OTK-00007", Novac=15000)
        nov(t, "NOV-00010", BrojDokumenta="1/050726", Datum=D5, EntitetTip="Kooperant", PartnerID="KOOP-00001",
            OMID="ST-00001", KooperantID="KOOP-00001", Tip="KesOtkupacKoop", Isplata=15000, OtkupID="OTK-00007")
        nn = self.nalazi(t)
        x = self.jedan(nn, K_ODVZ, "NOV-00001", GRESKA)
        self.assertIn("NOV-00010", x.predlog)
        self.nema(nn, K_OSIR)

    def test_legitimno_storno_i_ponovni_unos(self):
        t = nova()
        r(t, "tblOtkup", "OTK-00001")["Stornirano"] = "Da"
        n1 = r(t, "tblNovac", "NOV-00001")
        n1.update(OtkupID="", Stornirano="Da")                          # i kes je storniran
        otk(t, "OTK-00001", "OTK-00007", Novac=15000)
        nov(t, "NOV-00010", BrojDokumenta="1/050726", Datum=D5, EntitetTip="Kooperant", PartnerID="KOOP-00001",
            OMID="ST-00001", KooperantID="KOOP-00001", Tip="KesOtkupacKoop", Isplata=15000, OtkupID="OTK-00007")
        nn = self.nalazi(t)
        self.nema(nn, K_OSIR)
        self.nema(nn, K_ODVZ)

    def test_legitimno_split_vracen_u_avans(self):
        t = nova()
        nov(t, "NOV-00010", BrojDokumenta="OM-5", Datum=D5, Tip="VirmanAvansKoop", KooperantID="KOOP-00002",
            PartnerID="ST-00001", OMID="ST-00001", Isplata=3000)
        nov(t, "NOV-00011", BrojDokumenta="OM-5", Datum=D5, Tip="VirmanAvansKoop", KooperantID="KOOP-00002",
            EntitetTip="Kooperant", PartnerID="ST-00001", OMID="ST-00001", Isplata=2000, Napomena="Avans raspodela")
        nn = self.nalazi(t)
        self.nema(nn, K_OSIR)
        self.nema(nn, K_ODVZ)


# ---------------------------------------------------------------------------
class Preplata(Osnova):
    def test_preplata_dokumenta(self):
        t = nova()
        r(t, "tblNovac", "NOV-00001")["Isplata"] = 16000                 # vrednost 60 * 250 = 15000
        x = self.jedan(self.nalazi(t), K_PREPL, "OTK-00001", GRESKA)
        self.assertIn("1000", x.opis)

    def test_klasa_artefakt(self):
        t = nova()
        otk(t, "OTK-00001", "OTK-00007", Klasa="II", Kolicina=20)        # isti dokument, Klasa II = 5000
        r(t, "tblNovac", "NOV-00001")["Isplata"] = 20000                 # sav kes na Klasi I
        # SaveOtkupMulti_TX knjizi ceo kes na primarni red: ponasanje stare aplikacije, samo kontekst
        self.jedan(self.nalazi(t), K_PREPL, "OTK-00001, OTK-00007", INFO)

    def test_legitimno_zaokruzivanje(self):
        t = nova()
        r(t, "tblNovac", "NOV-00001")["Isplata"] = Decimal("15000.6")
        self.nema(self.nalazi(t), K_PREPL)


class IsplacenoPosleStorna(Osnova):
    def test_isplaceno_a_isplata_stornirana(self):
        t = nova()
        r(t, "tblNovac", "NOV-00002")["Stornirano"] = "Da"               # StornoNovac ne dira Isplaceno
        self.jedan(self.nalazi(t), K_ISPL, "OTK-00004", GRESKA)

    def test_legitimno_zamenska_isplata(self):
        t = nova()
        r(t, "tblNovac", "NOV-00002")["Stornirano"] = "Da"
        nov(t, "NOV-00010", BrojDokumenta="V-2", Datum=D6, EntitetTip="Kooperant", PartnerID="KOOP-00003",
            OMID="ST-00002", KooperantID="KOOP-00003", Tip="VirmanFirmaKoop", Isplata=20000, OtkupID="OTK-00004")
        self.nema(self.nalazi(t), K_ISPL)


# ---------------------------------------------------------------------------
class VezaTudji(Osnova):
    def test_tudji_kooperant(self):
        t = nova()
        r(t, "tblNovac", "NOV-00002")["OtkupID"] = "OTK-00006"            # otkup KOOP-00004
        self.jedan(self.nalazi(t), K_VEZA, "NOV-00002", GRESKA)

    def test_nepostojeci_otkup(self):
        t = nova()
        r(t, "tblNovac", "NOV-00002")["OtkupID"] = "OTK-00099"
        self.jedan(self.nalazi(t), K_VEZA, "NOV-00002", GRESKA)

    def test_tudji_kupac_na_fakturi(self):
        t = nova()
        fak(t, "FAK-00001", KupacID="KUP-00002")
        kupac_upl(t, "NOV-00010", FakturaID="FAK-00001", Uplata=10000, Napomena="Uplata po fakturi: 1/2026")
        self.jedan(self.nalazi(t), K_VEZA, "NOV-00010", GRESKA)

    def test_legitimno_druga_stanica_samo_info(self):
        t = nova()
        _rucni_virman_004(t)["OMID"] = "ST-00003"                       # Ulaz OM sa druge stanice
        nn = self.nalazi(t)
        self.nema(nn, K_VEZA, (GRESKA, SUMNJA))
        self.jedan(nn, K_VEZA, "NOV-00002", INFO)

    def test_legitimno_izvod_omid_iz_sifarnika(self):
        t = nova()
        bim(t, "BIM-00001", Isplata=20000, PozivNaBroj="2/050726")
        r(t, "tblNovac", "NOV-00002").update(Napomena="BIM:BIM-00001; Match:Kooperant", OMID="ST-00003",
                                            EntitetTip="OM", BrojDokumenta="12")
        self.nema(self.nalazi(t), K_VEZA)


    def test_legitimno_split_omid_iz_sifarnika(self):
        t = nova()
        bim(t, "BIM-00001", DatumTransakcije="06.07.2026", Isplata=30000)
        izvod(t, "NOV-00010", "BIM-00001", Datum=D6, Tip="VirmanAvansKoop", KooperantID="KOOP-00002",
              PartnerID="ST-00003", OMID="ST-00003", Isplata=12500)
        nov(t, "NOV-00011", BrojDokumenta="12", Datum=D6, Tip="VirmanAvansKoop", KooperantID="KOOP-00002",
            EntitetTip="Kooperant", PartnerID="ST-00003", OMID="ST-00003", Isplata=17500, Napomena="Avans raspodela",
            OtkupID="OTK-00005")                                       # OTK-00005 je na ST-00001
        self.nema(self.nalazi(t), K_VEZA)


# ---------------------------------------------------------------------------
def _fs(t, id_, fid, prj):
    t["tblFakturaStavke"].append(dict(StavkaID=id_, FakturaID=fid, PrijemnicaID=prj, Kolicina=99, Cena=300,
                                      Klasa="I", BrojPrijemnice="1/050726", Stornirano="", OsirocenoOd=""))


class FakturaStornirana(Osnova):
    def test_uplata_na_storniranu(self):
        t = nova()
        fak(t, "FAK-00001", Stornirano="Da", Status="Stornirano")
        _fs(t, "FS-00001", "FAK-00001", "PRJ-00001")
        t["tblFakturaStavke"][-1]["Stornirano"] = "Da"
        fak(t, "FAK-00002", BrojFakture="2/2026", Datum=D6)
        _fs(t, "FS-00002", "FAK-00002", "PRJ-00001")
        kupac_upl(t, "NOV-00010", FakturaID="FAK-00001", Uplata=10000, Napomena="Uplata po fakturi: 1/2026")
        x = self.jedan(self.nalazi(t), K_FAKST, "NOV-00010", GRESKA)
        self.assertIn("-> FAK-00002", x.predlog)

    def test_legitimno_aktivna_faktura(self):
        t = nova()
        fak(t, "FAK-00001")
        kupac_upl(t, "NOV-00010", FakturaID="FAK-00001", Uplata=10000, Napomena="Uplata po fakturi: 1/2026")
        self.nema(self.nalazi(t), K_FAKST)


class UplataKupca(Osnova):
    def test_preplacena_faktura(self):
        t = nova()
        fak(t, "FAK-00001")
        kupac_upl(t, "NOV-00010", FakturaID="FAK-00001", Uplata=10000, Napomena="Uplata po fakturi: 1/2026")
        kupac_upl(t, "NOV-00011", FakturaID="FAK-00001", Uplata=2000, Napomena="Uplata po fakturi: 1/2026")
        self.jedan(self.nalazi(t), K_UPLK, "FAK-00001", GRESKA)

    def test_kupci_uplata_bez_fakture(self):
        t = nova()
        fak(t, "FAK-00001", BrojFakture="5/2026", Stornirano="Da", Status="Stornirano")
        kupac_upl(t, "NOV-00010", Uplata=10000, Napomena="Uplata po fakturi: 5/2026 | 05.07.2026 | 10.000")
        x = self.jedan(self.nalazi(t), K_UPLK, "NOV-00010", SUMNJA)
        self.assertIn("5/2026", x.opis)

    def test_legitimno_tacna_uplata_i_avans(self):
        t = nova()
        fak(t, "FAK-00001")
        kupac_upl(t, "NOV-00010", FakturaID="FAK-00001", Uplata=Decimal("10000.5"), Napomena="Uplata po fakturi: 1/2026")
        kupac_upl(t, "NOV-00011", Tip="KupciAvans", Uplata=3000, Napomena="Avans kupca")
        self.nema(self.nalazi(t), K_UPLK)


class OdvezanaStornomFakture(Osnova):
    def test_avans_posle_storna_fakture(self):
        t = nova()
        fak(t, "FAK-00001", Stornirano="Da", Status="Stornirano")
        kupac_upl(t, "NOV-00010", Tip="KupciAvans", Uplata=10000, Napomena="Avans kupca", Datum=D6)
        self.jedan(self.nalazi(t), K_ODVF, "NOV-00010", INFO)

    def test_legitimno_bez_stornirane_fakture(self):
        t = nova()
        fak(t, "FAK-00001")
        kupac_upl(t, "NOV-00010", Tip="KupciAvans", Uplata=10000, Napomena="Avans kupca", Datum=D6)
        self.nema(self.nalazi(t), K_ODVF)


# ---------------------------------------------------------------------------
class IzvodKupacOm(Osnova):
    def test_poziv_na_fakturu_a_avans(self):
        t = nova()
        fak(t, "FAK-00001", BrojFakture="7/2026", Iznos=29700)
        bim(t, "BIM-00001", Uplata=29700, PozivNaBroj="7/2026", Partner="Hladnjaca Test")
        izvod(t, "NOV-00010", "BIM-00001", match="Kupac", EntitetTip="Kupac", PartnerID=KUPAC, Tip="KupciAvans",
              Uplata=29700)
        x = self.jedan(self.nalazi(t), K_KUPOM, "NOV-00010", SUMNJA)
        self.assertIn("FAK-00001", x.predlog)

    def test_legitimno_poziv_podniz(self):
        t = nova()
        fak(t, "FAK-00001", BrojFakture="1/2026", Iznos=29700)
        bim(t, "BIM-00001", Uplata=29700, PozivNaBroj="11/2026", SvrhaPlacanja="avans po ugovoru 11/2026")
        izvod(t, "NOV-00010", "BIM-00001", match="Kupac", EntitetTip="Kupac", PartnerID=KUPAC, Tip="KupciAvans",
              Uplata=29700)
        self.nema(self.nalazi(t), K_KUPOM)

    def test_istoimeni_kupac_po_kontu(self):
        t = nova()
        t["tblKupci"].append(dict(KupacID="KUP-00002", Naziv="Hladnjaca Test", Aktivan="Da",
                                  TekuciRacun="160-0000000012345-67"))
        bim(t, "BIM-00001", Uplata=5000, PartnerKonto="160-12345-67", Partner="HLADNJACA TEST")
        izvod(t, "NOV-00010", "BIM-00001", match="Kupac", EntitetTip="Kupac", PartnerID=KUPAC, Tip="KupciAvans",
              Uplata=5000)
        x = self.jedan(self.nalazi(t), K_KUPOM, "NOV-00010", GRESKA)
        self.assertIn("KUP-00002", x.predlog)

    def test_isplata_kupcu(self):
        t = nova()
        bim(t, "BIM-00001", Isplata=5000, Partner="Hladnjaca Test")
        izvod(t, "NOV-00010", "BIM-00001", match="Kupac", EntitetTip="Kupac", PartnerID=KUPAC, Tip="KupciAvans",
              Isplata=5000)
        self.jedan(self.nalazi(t), K_KUPOM, "NOV-00010", SUMNJA)


class PogresanKooperant(Osnova):
    def _osnova(self, konto, poziv=""):
        t = nova()
        r(t, "tblKooperanti", "KOOP-00003")["TekuciRacun"] = "205-111-11"
        r(t, "tblKooperanti", "KOOP-00004")["TekuciRacun"] = "205-222-22"
        bim(t, "BIM-00001", Isplata=5000, PartnerKonto=konto, Partner="Zika Zikic", PozivNaBroj=poziv)
        izvod(t, "NOV-00010", "BIM-00001", Tip="VirmanAvansKoop", KooperantID="KOOP-00003", PartnerID="ST-00002",
              OMID="ST-00002", Isplata=5000)
        return t

    def test_konto_drugog_kooperanta(self):
        x = self.jedan(self.nalazi(self._osnova("205-0000000000222-22")), K_KOOP, "NOV-00010", GRESKA)
        self.assertIn("KOOP-00004", x.predlog)
        self.assertIn("ST-00003", x.predlog)

    def test_poziv_svoj_blok_porodica_sumnja(self):
        self.jedan(self.nalazi(self._osnova("205-222-22", "2/050726")), K_KOOP, "NOV-00010", SUMNJA)

    def test_legitimno_konto_van_sifarnika(self):
        self.nema(self.nalazi(self._osnova("205-999-99")), K_KOOP)


# ---------------------------------------------------------------------------
class TipIsplate(Osnova):
    def _pokriven(self, posle=True, izvod_koop=True):
        t = nova()
        _rucni_virman_004(t)
        bim(t, "BIM-00001", DatumTransakcije="01.07.2026", Isplata=150, Obradjeno="Skip", Partner="Banka")
        if posle:
            bim(t, "BIM-00002", DatumTransakcije="10.07.2026", Isplata=150, Obradjeno="Skip", Partner="Banka")
        if izvod_koop:      # isti kooperant se inace placa sa uvezenog racuna
            bim(t, "BIM-00003", DatumTransakcije="28.06.2026", Isplata=4000, Partner="Zika Zikic")
            izvod(t, "NOV-00010", "BIM-00003", Datum=date(2026, 6, 28), Tip="VirmanAvansKoop",
                  KooperantID="KOOP-00003", PartnerID="ST-00002", OMID="ST-00002", Isplata=4000)
        return t

    def test_virman_bez_izvoda_u_pokrivenom_periodu(self):
        self.jedan(self.nalazi(self._pokriven()), K_TIP, "NOV-00002", SUMNJA)

    def test_legitimno_period_nije_pokriven(self):
        self.nema(self.nalazi(self._pokriven(posle=False)), K_TIP)

    def test_legitimno_kooperant_van_uvezenog_racuna(self):
        self.nema(self.nalazi(self._pokriven(izvod_koop=False)), K_TIP)

    def test_sistemska_praksa_jedan_info(self):
        t = self._pokriven()
        for i in range(8):
            o = otk(t, "OTK-00004", "OTK-%05d" % (20 + i), BrojDokumenta="2/060726-%d" % (i + 2), Datum=D6)
            nov(t, "NOV-%05d" % (30 + i), BrojDokumenta="OM-%d" % i, Datum=D6, Tip="VirmanFirmaKoop",
                PartnerID="ST-00002", OMID="ST-00002", KooperantID="KOOP-00003", Isplata=20000 + i,
                OtkupID=o["OtkupID"])
        nn = self.nalazi(t)
        self.nema(nn, K_TIP, (GRESKA, SUMNJA))
        x = [y for y in nn if y.kod == K_TIP]
        self.assertEqual([y.ozbiljnost for y in x], [INFO])
        self.assertIn("9 od 9", x[0].opis)

    def test_kes_om_bez_primaoca_jednak_ostatku(self):
        t = nova()
        nov(t, "NOV-00010", BrojDokumenta="OM-1", Datum=D5, Tip="KesFirmaOtkupac", PartnerID="ST-00001",
            OMID="ST-00001", Isplata=10000)                              # OTK-00002: 40 * 250 = 10000 otvoreno
        self.jedan(self.nalazi(t), K_TIP, "NOV-00010", SUMNJA)

    def test_legitimno_kes_om_jednak_klasi_ii_placenog_dokumenta(self):
        t = nova()
        otk(t, "OTK-00001", "OTK-00007", Klasa="II", Kolicina=20)        # 5000, isti dokument
        r(t, "tblNovac", "NOV-00001")["Isplata"] = 20000                 # ceo kes na Klasi I
        r(t, "tblOtkup", "OTK-00002")["Kolicina"] = 60                   # OTK-00002 da ne bude kandidat 5000
        nov(t, "NOV-00010", BrojDokumenta="OM-1", Datum=D5, Tip="KesFirmaOtkupac", PartnerID="ST-00001",
            OMID="ST-00001", Isplata=5000)
        self.nema(self.nalazi(t), K_TIP)

    def _om_avans(self, stor):
        t = nova()
        nov(t, "NOV-00010", BrojDokumenta="OM-1", Datum=D5, Tip="KesFirmaOtkupac", PartnerID="ST-00001",
            OMID="ST-00001", Isplata=30000, Stornirano=stor)
        nov(t, "NOV-00011", BrojDokumenta="OM-2", Datum=D6, Tip="KesOtkupacKoop", PartnerID="ST-00001",
            OMID="ST-00001", KooperantID="KOOP-00002", Isplata=8000, OtkupID="OTK-00005", Napomena="Mika Mikic")
        return t

    def test_saldo_om_posle_storna_avansa(self):
        x = self.jedan(self.nalazi(self._om_avans("Da")), K_TIP, "NOV-00011", INFO)
        self.assertIn("NOV-00010", x.opis)

    def test_legitimno_saldo_pokriven(self):
        self.nema(self.nalazi(self._om_avans("")), K_TIP)

    def test_legitimno_kes_om_drugi_iznos(self):
        t = nova()
        nov(t, "NOV-00010", BrojDokumenta="OM-1", Datum=D5, Tip="KesFirmaOtkupac", PartnerID="ST-00001",
            OMID="ST-00001", Isplata=50000)
        self.nema(self.nalazi(t), K_TIP)


class ZastareoDatum(Osnova):
    def test_isplata_pre_otkupa(self):
        t = nova()
        nov(t, "NOV-00010", BrojDokumenta="OM-2", Datum=D5, Tip="KesOtkupacKoop", PartnerID="ST-00001",
            OMID="ST-00001", KooperantID="KOOP-00002", Isplata=5000, OtkupID="OTK-00005")   # otkup 06.07.
        x = self.jedan(self.nalazi(t), K_DAT, "NOV-00010", SUMNJA)
        self.assertIn("2026-07-06", x.predlog)

    def test_legitimno_avans_pre_otkupa(self):
        t = nova()
        nov(t, "NOV-00010", BrojDokumenta="OM-2", Datum=D5, Tip="VirmanAvansKoop", PartnerID="ST-00001",
            OMID="ST-00001", KooperantID="KOOP-00002", Isplata=5000, OtkupID="OTK-00005")
        self.nema(self.nalazi(t), K_DAT)


class DatumLokal(Osnova):
    def _t(self, datum_novca, dt="05.07.2026"):
        t = nova()
        bim(t, "BIM-00001", DatumTransakcije=dt, Isplata=5000)
        izvod(t, "NOV-00010", "BIM-00001", Datum=datum_novca, Tip="VirmanAvansKoop", KooperantID="KOOP-00002",
              PartnerID="ST-00001", OMID="ST-00001", Isplata=5000)
        return t

    def test_zamena_dan_mesec(self):
        x = self.jedan(self.nalazi(self._t(date(2026, 5, 7))), K_LOK, "NOV-00010", GRESKA)
        self.assertIn("ZAMENA_DAN_MESEC", x.opis)
        self.assertIn("-> 2026-07-05", x.predlog)

    def test_legitimno_isti_datum(self):
        nn = self.nalazi(self._t(D5))
        self.nema(nn, K_LOK)
        self.nema(self.nalazi(self._t(date(2026, 7, 18), "18.07.2026")), K_LOK)


# ---------------------------------------------------------------------------
def _izvodni_avans_sa_splitom(t, stor_kor="", stor_spl="", isp_kor=12500, isp_spl=17500, A=30000):
    bim(t, "BIM-00001", DatumTransakcije="05.07.2026", Isplata=A)
    izvod(t, "NOV-00010", "BIM-00001", Datum=D5, Tip="VirmanAvansKoop", KooperantID="KOOP-00002",
          PartnerID="ST-00001", OMID="ST-00001", Isplata=isp_kor, Stornirano=stor_kor)
    nov(t, "NOV-00011", BrojDokumenta="12", Datum=D5, Tip="VirmanAvansKoop", KooperantID="KOOP-00002",
        EntitetTip="Kooperant", PartnerID="ST-00001", OMID="ST-00001", Isplata=isp_spl, Napomena="Avans raspodela",
        OtkupID="OTK-00005", Stornirano=stor_spl)


class StornoNovcaBezLanca(Osnova):
    def test_storniran_original_split_aktivan(self):
        t = nova()
        _izvodni_avans_sa_splitom(t, stor_kor="Da")
        nn = self.nalazi(t)
        self.jedan(nn, K_STLAN, "NOV-00010", GRESKA)
        self.nema(nn, K_RAZD)

    def test_stavka_da_bez_reda(self):
        t = nova()
        bim(t, "BIM-00001", Isplata=7000, Partner="Laza Lazic")
        self.jedan(self.nalazi(t), K_STLAN, "BIM-00001", GRESKA)

    def test_stavka_da_zamenjena_rucnim(self):
        t = nova()
        bim(t, "BIM-00001", Isplata=20000, Partner="Zika Zikic")
        izvod(t, "NOV-00010", "BIM-00001", Tip="VirmanAvansKoop", KooperantID="KOOP-00003", PartnerID="ST-00002",
              OMID="ST-00002", Isplata=20000, Stornirano="Da")
        x = self.jedan(self.nalazi(t), K_STLAN, "BIM-00001", SUMNJA)
        self.assertIn("NOV-00002", x.opis)

    def test_stavka_da_zamenjena_rucnim_na_pravom_partneru(self):
        t = nova()
        bim(t, "BIM-00001", Isplata=7000, Partner="Laza Lazic")
        izvod(t, "NOV-00010", "BIM-00001", Tip="VirmanAvansKoop", KooperantID="KOOP-00003", PartnerID="ST-00002",
              OMID="ST-00002", Isplata=7000, Stornirano="Da")             # mapiran na pogresnog
        nov(t, "NOV-00011", Datum=D7, Tip="VirmanFirmaKoop", PartnerID="ST-00003", OMID="ST-00003",
            KooperantID="KOOP-00004", Isplata=7000, OtkupID="OTK-00006", Napomena="Laza Lazic")
        x = self.jedan(self.nalazi(t), K_STLAN, "BIM-00001", SUMNJA)
        self.assertIn("NOV-00011", x.opis)

    def test_zbir_lanca_razlicit(self):
        t = nova()
        _izvodni_avans_sa_splitom(t, isp_kor=10000)                      # 10000 + 17500 != 30000
        self.jedan(self.nalazi(t), K_STLAN, "BIM-00001", GRESKA)

    def test_legitimno_ceo_lanac(self):
        t = nova()
        _izvodni_avans_sa_splitom(t)
        nn = self.nalazi(t)
        self.nema(nn, K_STLAN)
        self.nema(nn, K_RAZD)


class RaspodelaRazdvojena(Osnova):
    def test_storniran_split_original_aktivan(self):
        t = nova()
        _izvodni_avans_sa_splitom(t, stor_spl="Da")
        nn = self.nalazi(t)
        self.jedan(nn, K_RAZD, "NOV-00011", GRESKA)
        self.nema(nn, K_STLAN)

    def _rucni_avans(self, nova_raspodela):
        t = nova()
        nov(t, "NOV-00010", BrojDokumenta="OM-5", Datum=D5, Tip="VirmanAvansKoop", KooperantID="KOOP-00002",
            PartnerID="ST-00001", OMID="ST-00001", Isplata=2500)
        nov(t, "NOV-00011", BrojDokumenta="OM-5", Datum=D5, Tip="VirmanAvansKoop", KooperantID="KOOP-00002",
            EntitetTip="Kooperant", PartnerID="ST-00001", OMID="ST-00001", Isplata=17500, Napomena="Avans raspodela",
            OtkupID="OTK-00005", Stornirano="Da")
        if nova_raspodela:
            nov(t, "NOV-00012", BrojDokumenta="OM-5", Datum=D5, Tip="VirmanAvansKoop", KooperantID="KOOP-00002",
                EntitetTip="Kooperant", PartnerID="ST-00001", OMID="ST-00001", Isplata=17500,
                Napomena="Avans raspodela", OtkupID="OTK-00005")
        return t

    def test_rucni_avans_storniran_split(self):
        self.jedan(self.nalazi(self._rucni_avans(False)), K_RAZD, "NOV-00011", SUMNJA)

    def test_legitimno_rucni_avans_ponovna_raspodela(self):
        self.nema(self.nalazi(self._rucni_avans(True)), K_RAZD)

    def test_legitimno_ponovna_raspodela(self):
        t = nova()
        _izvodni_avans_sa_splitom(t, stor_spl="Da", isp_kor=12500)
        nov(t, "NOV-00012", BrojDokumenta="12", Datum=D5, Tip="VirmanAvansKoop", KooperantID="KOOP-00002",
            EntitetTip="Kooperant", PartnerID="ST-00001", OMID="ST-00001", Isplata=17500, Napomena="Avans raspodela",
            OtkupID="OTK-00005")
        self.nema(self.nalazi(t), K_RAZD)


# ---------------------------------------------------------------------------
def _om_virman(t, id_, otk_id="OTK-00005", isp=5000, broj="", dan=D6):
    return nov(t, id_, BrojDokumenta=broj, Datum=dan, Tip="VirmanFirmaKoop", PartnerID="ST-00001", OMID="ST-00001",
               KooperantID="KOOP-00002", Isplata=isp, OtkupID=otk_id)


class DupliRucni(Osnova):
    def test_isti_blok_uzastopno(self):
        t = nova()
        _om_virman(t, "NOV-00010", broj="OM-3")
        _om_virman(t, "NOV-00011")                                       # bez broja
        self.jedan(self.nalazi(t), K_DUPRU, "NOV-00010, NOV-00011", SUMNJA)

    def test_isti_blok_preplata(self):
        t = nova()
        _om_virman(t, "NOV-00010", isp=10000, broj="OM-3")
        _om_virman(t, "NOV-00011", isp=10000, broj="OM-4")               # 20000 > 17500
        self.jedan(self.nalazi(t), K_DUPRU, "NOV-00010, NOV-00011", GRESKA)

    def test_legitimno_kes_dva_bloka(self):
        t = nova()
        r(t, "tblOtkup", "OTK-00003")["Kolicina"] = 60                   # 15000, kao OTK-00001
        nov(t, "NOV-00003", BrojDokumenta="1/050726-3", Datum=D5, EntitetTip="Kooperant", PartnerID="KOOP-00001",
            OMID="ST-00001", KooperantID="KOOP-00001", Tip="KesOtkupacKoop", Isplata=15000, OtkupID="OTK-00003")
        self.nema(self.nalazi(t), K_DUPRU)


# ---------------------------------------------------------------------------
class IzvodVanKnjige(Osnova):
    def test_skip_zamenjen_rucnim(self):
        t = nova()
        bim(t, "BIM-00001", Isplata=20000, Partner="Zika Zikic", Obradjeno="Skip")
        x = self.jedan(self.nalazi(t), K_VAN, "BIM-00001", SUMNJA)
        self.assertIn("NOV-00002", x.predlog)

    def test_skip_partnermap_blokiran(self):
        t = nova()
        t["tblPartnerMap"] = [dict(BankaName="PERA PERIC DOO", PartnerID="KOOP-00001", EntitetTip="Kooperant",
                                   OMID="ST-00001")]
        bim(t, "BIM-00001", Isplata=7777, Partner="Pera Peric doo", Obradjeno="Error")
        self.jedan(self.nalazi(t), K_VAN, "BIM-00001", SUMNJA)

    def test_legitimno_provizija_samo_info(self):
        t = nova()
        bim(t, "BIM-00001", Isplata=150, Partner="Banka Intesa", Obradjeno="Skip")
        nn = self.nalazi(t)
        self.nema(nn, K_VAN, (GRESKA, SUMNJA))
        self.jedan(nn, K_VAN, "BIM-00001", INFO)


class BezPozivaAvans(Osnova):
    def test_avans_potrosen_na_kasniji_a_stari_otvoren(self):
        t = nova()
        bim(t, "BIM-00001", DatumTransakcije="05.07.2026", Isplata=10000)
        izvod(t, "NOV-00010", "BIM-00001", Datum=D5, Tip="VirmanAvansKoop", KooperantID="KOOP-00002",
              PartnerID="ST-00001", OMID="ST-00001", Isplata=10000, OtkupID="OTK-00005")   # OTK-00005 je 06.07.
        x = self.jedan(self.nalazi(t), K_BEZP, "NOV-00010", SUMNJA)
        self.assertIn("OTK-00002", x.predlog)

    def test_legitimno_klasa_ii_dokument_placen(self):
        t = nova()
        otk(t, "OTK-00001", "OTK-00007", Klasa="II", Kolicina=20)        # 5000, isti dokument 1/050726
        r(t, "tblNovac", "NOV-00001")["Isplata"] = 20000                 # ceo kes na Klasi I
        nov(t, "NOV-00003", BrojDokumenta="1/050726-3", Datum=D5, EntitetTip="Kooperant", PartnerID="KOOP-00001",
            OMID="ST-00001", KooperantID="KOOP-00001", Tip="KesOtkupacKoop", Isplata=12500, OtkupID="OTK-00003")
        otk(t, "OTK-00001", "OTK-00008", Datum=D10, BrojDokumenta="1/100726", Kolicina=40)
        bim(t, "BIM-00001", DatumTransakcije="07.07.2026", Isplata=5000)
        izvod(t, "NOV-00010", "BIM-00001", Datum=D7, Tip="VirmanAvansKoop", KooperantID="KOOP-00001",
              PartnerID="ST-00001", OMID="ST-00001", Isplata=5000, OtkupID="OTK-00008")
        self.nema(self.nalazi(t), K_BEZP)

    def test_legitimno_stari_blok_placen(self):
        t = nova()
        nov(t, "NOV-00003", BrojDokumenta="1/050726-2", Datum=D5, EntitetTip="Kooperant", PartnerID="KOOP-00002",
            OMID="ST-00001", KooperantID="KOOP-00002", Tip="KesOtkupacKoop", Isplata=10000, OtkupID="OTK-00002")
        bim(t, "BIM-00001", DatumTransakcije="05.07.2026", Isplata=10000)
        izvod(t, "NOV-00010", "BIM-00001", Datum=D5, Tip="VirmanAvansKoop", KooperantID="KOOP-00002",
              PartnerID="ST-00001", OMID="ST-00001", Isplata=10000, OtkupID="OTK-00005")
        self.nema(self.nalazi(t), K_BEZP)


# ---------------------------------------------------------------------------
def _revers(t, stor):
    for i, (smer, eid, etip) in enumerate((("Ulaz", "KOOP-00001", "Kooperant"), ("Izlaz", "ST-00001", "Stanica")), 1):
        t["tblAmbalaza"].append(dict(AmbID="AMB-0090%d" % i, Datum=D7, TipAmbalaze="12/1", Kolicina=50, Smer=smer,
                                     EntitetID=eid, EntitetTip=etip, VozacID="", DokumentID="1/070726",
                                     DokumentTIP="OM-Izlaz-Koop", Stornirano=stor, CreatedAt="", CreatedBy="op"))
    nov(t, "NOV-00010", BrojDokumenta="1/070726", Datum=D7, Tip="VirmanAvansKoop", PartnerID="ST-00001",
        OMID="ST-00001", KooperantID="KOOP-00001", Isplata=3000, Napomena="Pera Peric")


class ReversStorniran(Osnova):
    def test_novac_ostao(self):
        t = nova()
        _revers(t, "Da")
        self.jedan(self.nalazi(t), K_REV, "NOV-00010", SUMNJA)

    def test_novac_ostao_i_ponovo_unet(self):
        t = nova()
        _revers(t, "Da")
        for i, (smer, eid, etip) in enumerate((("Ulaz", "KOOP-00001", "Kooperant"), ("Izlaz", "ST-00001", "Stanica")),
                                              1):
            t["tblAmbalaza"].append(dict(AmbID="AMB-0091%d" % i, Datum=D7, TipAmbalaze="12/1", Kolicina=40,
                                         Smer=smer, EntitetID=eid, EntitetTip=etip, VozacID="",
                                         DokumentID="1/070726-2", DokumentTIP="OM-Izlaz-Koop", Stornirano="",
                                         CreatedAt="", CreatedBy="op"))
        nov(t, "NOV-00011", BrojDokumenta="1/070726-2", Datum=D7, Tip="VirmanAvansKoop", PartnerID="ST-00001",
            OMID="ST-00001", KooperantID="KOOP-00001", Isplata=3000)
        x = self.jedan(self.nalazi(t), K_REV, "NOV-00010", GRESKA)
        self.assertIn("NOV-00011", x.opis)

    def test_legitimno_kasniji_isti_avans_nije_greska(self):
        """Isti iznos kasnije, bez novog reversa: moze biti redovan fiksni avans -> SUMNJA, ne GRESKA."""
        t = nova()
        _revers(t, "Da")
        nov(t, "NOV-00011", BrojDokumenta="", Datum=date(2026, 7, 20), Tip="VirmanAvansKoop", PartnerID="ST-00001",
            OMID="ST-00001", KooperantID="KOOP-00001", Isplata=3000)
        self.jedan(self.nalazi(t), K_REV, "NOV-00010", SUMNJA)

    def test_legitimno_revers_aktivan(self):
        t = nova()
        _revers(t, "")
        self.nema(self.nalazi(t), K_REV)


class UplataOm(Osnova):
    def test_izvodna_uplata_om(self):
        t = nova()
        bim(t, "BIM-00001", Uplata=5000, Partner="Stanica 1")
        izvod(t, "NOV-00010", "BIM-00001", match="OM", Tip="KesFirmaOtkupac", PartnerID="ST-00001", OMID="ST-00001",
              Uplata=5000)
        self.jedan(self.nalazi(t), K_UPLOM, "NOV-00010", SUMNJA)

    def test_legitimno_isplata_om(self):
        t = nova()
        bim(t, "BIM-00001", Isplata=5000, Partner="Stanica 1")
        izvod(t, "NOV-00010", "BIM-00001", match="OM", Tip="KesFirmaOtkupac", PartnerID="ST-00001", OMID="ST-00001",
              Isplata=5000)
        self.nema(self.nalazi(t), K_UPLOM)


class PozicioniPomak(Osnova):
    def _kol(self, a, b):
        kol = list(fixture.KOLONE["tblNovac"])
        i, j = kol.index(a), kol.index(b)
        kol[i], kol[j] = kol[j], kol[i]
        return {"tblNovac": kol}

    def test_zaglavlje_pomereno_i_vrednosti(self):
        t = nova()
        for x in t["tblNovac"]:                       # pisac je pisao poziciono: Tip <-> VrstaVoca
            x["Tip"], x["VrstaVoca"] = x["VrstaVoca"], x["Tip"]
        self.jedan(self.nalazi(t, self._kol("Tip", "VrstaVoca")), K_POZ, "", GRESKA)

    def test_zaglavlje_pomereno_vrednosti_ispravne(self):
        self.jedan(self.nalazi(nova(), self._kol("Uplata", "Isplata")), K_POZ, "", SUMNJA)

    def test_nepoznat_tip(self):
        t = nova()
        r(t, "tblNovac", "NOV-00002")["Tip"] = "Kes"
        self.jedan(self.nalazi(t), K_POZ, "NOV-00002", SUMNJA)

    def test_legitimno_otkupid_drugog_oblika_postoji(self):
        t = nova()
        r(t, "tblOtkup", "OTK-00004")["OtkupID"] = "PWA-00004"
        r(t, "tblNovac", "NOV-00002")["OtkupID"] = "PWA-00004"
        for a in t["tblAmbalaza"]:
            if a["DokumentID"] == "OTK-00004":
                a["DokumentID"] = "PWA-00004"
        self.nema(self.nalazi(t), K_POZ)

    def test_partner_map(self):
        t = nova()
        # savePartnerMap pise poziciono (BankaName, PartnerID, EntitetTip, OMID) u pomereno zaglavlje
        t["tblPartnerMap"] = [dict(BankaName="ZIKA", EntitetTip="KOOP-00003", PartnerID="Kooperant", OMID="ST-00002")]
        self.jedan(self.nalazi(t, {"tblPartnerMap": ["BankaName", "EntitetTip", "PartnerID", "OMID"]}),
                   K_POZ, "", GRESKA)

    def test_legitimno_audit_kolone_na_kraju(self):
        kol = list(fixture.KOLONE["tblNovac"]) + ["IspravkaOd", "ZamenjenSa"]
        self.nema(self.nalazi(nova(), {"tblNovac": kol}), K_POZ)


class DedupeInfo(Osnova):
    def test_racuni_i_rupe(self):
        t = nova()
        bim(t, "BIM-00001", BrojRacuna="160-1", BrojDokumenta="1", Isplata=150, Obradjeno="Skip")
        bim(t, "BIM-00002", BrojRacuna="160-1", BrojDokumenta="3", Isplata=150, Obradjeno="Skip")
        bim(t, "BIM-00003", BrojRacuna="205-9", BrojDokumenta="1", Isplata=150, Obradjeno="Skip")
        x = self.jedan(self.nalazi(t), K_DED, "", INFO)
        self.assertIn("= 2", x.opis)
        self.assertIn("rupe u nizu 2", x.opis)

    def test_bez_kolone_brojracuna(self):
        t = nova()
        bim(t, "BIM-00001", Isplata=150, Obradjeno="Skip")
        kol = [c for c in KOL["tblBankaImport"] if c != "BrojRacuna"]
        self.nema(self.nalazi(t, {"tblBankaImport": kol}), K_DED)


# ---------------------------------------------------------------------------
class Robusnost(Osnova):
    def test_kolona_fali_preskoceno(self):
        t = nova()
        kol = [c for c in fixture.KOLONE["tblNovac"] if c != "Napomena"]
        nn = self.nalazi(t, {"tblNovac": kol})
        x = [y for y in nn if y.kod == K_DUPRU + "_PRESKOCENO"]
        self.assertEqual(len(x), 1)
        self.assertEqual(x[0].ozbiljnost, INFO)
        self.assertIn("Napomena", x[0].opis)

    def test_bez_tabela_tiho(self):
        t = nova()
        del t["tblBankaImport"]
        t["tblNovac"] = []
        self.assertEqual([(x.kod, x.id) for x in self.nalazi(t)], [])

    def test_prljave_vrednosti_bez_pada(self):
        t = nova()
        nov(t, "NOV-00010", Datum="", Tip="VirmanFirmaKoop", KooperantID="KOOP-00002", Isplata="pet hiljada",
            OtkupID="OTK-00005")
        nov(t, "NOV-00011", Datum="nije datum", Tip="", Isplata="1.000,5", Napomena="BIM:")
        bim(t, "BIM-00001", DatumTransakcije="", Isplata="abc", Obradjeno="")
        bim(t, "", DatumTransakcije="32.13.2026", Isplata=5)
        r(t, "tblOtkup", "OTK-00002").update(Kolicina="", Cena="x", Datum="")
        self.nalazi(t)


if __name__ == "__main__":
    unittest.main()
