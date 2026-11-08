"""Testovi modula provera/storno.py: polazi se od ciste fixture sveske, podmetne
se trag storna kakav je ostavljala stara aplikacija i trazi kod na tacnom ID-ju."""

import random
import unittest
from datetime import date, datetime
from decimal import Decimal

from provera import fixture
from provera.fixture import D6, KUPAC
from provera.pokretac import pokreni_nad_tabelama
from provera.test_zajednicko import tabele_iz
from provera.zajednicko import GRESKA, INFO, SUMNJA

KOD_BLOK = "OTP_AKTIVNA_SA_STORNIRANIM_BLOKOM"
KOD_PRJ = "PRJ_STORNO_PO_BROJU_TUDJI_KUPAC"
KOD_KOL = "KOLATERALNI_STORNO_PO_BROJU"
KOD_OTK = "OTK_STORNO_PO_BROJU_TUDJI_BLOK"
KOD_PON = "ZBR_PONISTENJE_OBARA_TUDJU_PRIJEMNICU"
KOD_USK = "AUTO_USKRSNUCE_POSLE_STORNA"
KOD_DEL = "DOK_DELIMICNO_STORNIRAN"
KOD_KTX = "STORNO_KONTEKST_NEDOVRSEN"
KOD_BZ = "OTK_BROJZBIRNE_NA_STORNIRANOJ_ZBIRNOJ"
KOD_FAK = "FAK_OSIROCENO_ZASTARELO"

TREN = "2026-07-07 10:00:01"      # ModifiedAt jedne storno operacije
AUDIT = ("CreatedAt", "CreatedBy", "ModifiedAt", "ModifiedBy")
# sveska bez audit kolona (EnsureAuditColumns nikad pokrenut); tblStornoVeze ima svoj CreatedAt
BEZ_AUDITA = {ime: [c for c in kol if c not in AUDIT] for ime, kol in fixture.KOLONE.items()
              if ime != "tblStornoVeze"}
BEZ_VREMENA_OTP = {"tblOtpremnica": [c for c in fixture.KOLONE["tblOtpremnica"] if c not in AUDIT]}


def nova():
    return fixture.kopija(fixture.osnovna())


def r(t, ime, id_):
    return fixture.red(t, ime, fixture.KOLONE[ime][0], id_)


def dodaj_kopiju(t, ime, id_, **izmene):
    novi = dict(r(t, ime, id_))
    novi.update(izmene)
    t[ime].append(novi)
    return novi


def dodaj_sv(t, **kw):
    red = dict(CorrectionID="COR-00001", Mode="ISPRAVKA_ODMAH", Status="PENDING", OldDocType="Prijemnica",
               OldDocID="PRJ-00002", OldBroj="1/050726-2", NewDocType="Prijemnica", NewDocID="", NewBroj="",
               ParentDocType="Zbirna", ParentDocID="", ParentBroj="S1/050726-2", CreatedAt="2026-07-07 09:00:00",
               CreatedBy="op", CompletedAt="", Message="Ispravka prijemnice: storno stare, ceka snimanje nove.",
               NeedsRecovery="Da", RecoveryAction="")
    red.update(kw)
    t["tblStornoVeze"].append(red)
    return red


class Osnova(unittest.TestCase):
    def nalazi(self, t, kolone=None):
        ctx = pokreni_nad_tabelama(tabele_iz(t, kolone), moduli=("storno",))
        self.assertEqual([x.opis for x in ctx.nalazi if x.kod == "MODUL_PAO"], [])
        return ctx.nalazi

    def jedan(self, nn, kod, id_, ozb):
        x = [y for y in nn if y.kod == kod]
        self.assertEqual(len(x), 1, [(y.kod, y.id, y.ozbiljnost, y.opis) for y in nn])
        self.assertEqual((x[0].id, x[0].ozbiljnost), (id_, ozb), x[0].opis)
        return x[0]

    def nema(self, nn, kod):
        self.assertEqual([(y.id, y.opis) for y in nn if y.kod == kod], [])


class CistaSveska(Osnova):
    def test_nula_nalaza(self):
        self.assertEqual([(x.kod, x.id, x.opis[:120]) for x in self.nalazi(nova())], [])


# ---------------------------------------------------------------------------
class StorniranBlokNaAktivnojOtpremnici(Osnova):
    def test_storniran_blok_ostao_na_otpremnici(self):
        t = nova()
        r(t, "tblOtkup", "OTK-00002")["Stornirano"] = "Da"       # panel: StornoOtkupByBrDok_TX, OtpremnicaID ostaje
        x = self.jedan(self.nalazi(t), KOD_BLOK, "OTP-00001", GRESKA)
        self.assertIn("100 -> 60", x.predlog)
        self.assertIn("40 -> 24", x.predlog)
        self.assertIn("S1/050726", x.opis)
        self.assertIn("OTK-00002", x.opis)

    def test_ponovni_unos_nevezan_predlaze_vezu(self):
        t = nova()
        r(t, "tblOtkup", "OTK-00002")["Stornirano"] = "Da"
        dodaj_kopiju(t, "tblOtkup", "OTK-00002", OtkupID="OTK-00007", Stornirano="", OtpremnicaID="", BrojZbirne="",
                     BrojOtpremnice="", BrojDokumenta="1/060726-2", Datum=D6)
        x = self.jedan(self.nalazi(t), KOD_BLOK, "OTP-00001", GRESKA)
        self.assertIn("Vezati nevezan otkup OTK-00007", x.predlog)

    def test_ponovo_unet_na_drugu_otpremnicu_dva_puta(self):
        t = nova()
        r(t, "tblOtkup", "OTK-00002")["Stornirano"] = "Da"
        dodaj_kopiju(t, "tblOtkup", "OTK-00002", OtkupID="OTK-00007", Stornirano="", OtpremnicaID="OTP-00002",
                     BrojZbirne="S1/050726-2", BrojDokumenta="1/050726-4")
        x = self.jedan(self.nalazi(t), KOD_BLOK, "OTP-00001", GRESKA)
        self.assertIn("dva puta", x.opis)

    def test_druga_isporuka_istog_kooperanta_ne_podize_ozbiljnost(self):
        """Isti kooperant i kg na drugoj turi istog dana moze biti druga isporuka:
        gajbe se ne slazu, pa ostaje SUMNJA, a opis ne tvrdi 'dva puta'."""
        t = nova()
        r(t, "tblOtkup", "OTK-00002")["Stornirano"] = "Da"
        r(t, "tblOtpremnica", "OTP-00001")["KolAmbalaze"] = 41
        r(t, "tblOtpremnica", "OTP-00002").update(Kolicina=90, KolAmbalaze=36)
        dodaj_kopiju(t, "tblOtkup", "OTK-00002", OtkupID="OTK-00007", Stornirano="", OtpremnicaID="OTP-00002",
                     BrojZbirne="S1/050726-2", BrojDokumenta="1/050726-4")
        x = self.jedan(self.nalazi(t), KOD_BLOK, "OTP-00001", SUMNJA)
        self.assertIn("moze biti i druga isporuka", x.opis)

    def test_zamenjen_blok_vezan_za_istu_otpremnicu_nije_nalaz(self):
        t = nova()
        r(t, "tblOtkup", "OTK-00002")["Stornirano"] = "Da"
        dodaj_kopiju(t, "tblOtkup", "OTK-00002", OtkupID="OTK-00007", Stornirano="", BrojDokumenta="1/050726-4")
        self.nema(self.nalazi(t), KOD_BLOK)

    def test_blok_nikad_unet_nije_ova_klasa(self):
        t = nova()
        t["tblOtkup"] = [o for o in t["tblOtkup"] if o["OtkupID"] != "OTK-00002"]   # manjak bez storna (MIG-ODL-18)
        self.nema(self.nalazi(t), KOD_BLOK)


# ---------------------------------------------------------------------------
def kolizija_prijemnice(t, isti_trenutak=True):
    """Kupac 2 ima rucni broj 1/050726 kao prijemnica kupca hladnjace. Storno po
    broju (ispravka PRJ-00001) obori obe, a ponovo se unese samo PRJ-00001."""
    t["tblKupci"].append(dict(KupacID="KUP-00002", Naziv="Drugi kupac", Aktivan="Da", Hladnjaca=""))
    r(t, "tblZbirna", "ZBR-00003")["KupacID"] = "KUP-00002"
    p1, p3 = r(t, "tblPrijemnica", "PRJ-00001"), r(t, "tblPrijemnica", "PRJ-00003")
    p3.update(KupacID="KUP-00002", BrojPrijemnice="1/050726")
    dodaj_kopiju(t, "tblPrijemnica", "PRJ-00001", PrijemnicaID="PRJ-00006", CreatedAt="2026-07-07 10:05:00",
                 ModifiedAt="2026-07-07 10:05:00")
    p1.update(Stornirano="Da", ModifiedAt=TREN)
    p3.update(Stornirano="Da", ModifiedAt=TREN if isti_trenutak else "2026-07-06 18:00:00")


class StornoPrijemnicePoBroju(Osnova):
    def test_tudji_kupac_oboren_istom_operacijom(self):
        t = nova()
        kolizija_prijemnice(t)
        x = self.jedan(self.nalazi(t), KOD_PRJ, "PRJ-00003", GRESKA)
        self.assertIn("S2/050726", x.opis)
        self.assertIn("Vratiti PRJ-00003 u aktivno", x.predlog)

    def test_bez_dokaza_trenutka_je_sumnja(self):
        t = nova()
        kolizija_prijemnice(t, isti_trenutak=False)
        self.jedan(self.nalazi(t), KOD_PRJ, "PRJ-00003", SUMNJA)

    def test_masovni_pecat_nije_dokaz_trenutka(self):
        """Isti ModifiedAt na mnogo redova (kopiranje/backfill) nije dokaz jednog storna."""
        t = nova()
        kolizija_prijemnice(t)
        for p in t["tblPrijemnica"]:
            p["ModifiedAt"] = TREN
        self.jedan(self.nalazi(t), KOD_PRJ, "PRJ-00003", SUMNJA)

    def test_oba_kupca_legitimno_stornirala_i_ponovo_unela(self):
        t = nova()
        kolizija_prijemnice(t)
        dodaj_kopiju(t, "tblPrijemnica", "PRJ-00003", PrijemnicaID="PRJ-00007", Stornirano="",
                     CreatedAt="2026-07-07 10:06:00")
        self.nema(self.nalazi(t), KOD_PRJ)

    def test_storno_pa_ponovni_unos_istog_broja_jedan_kupac(self):
        t = nova()
        r(t, "tblPrijemnica", "PRJ-00001").update(Stornirano="Da", ModifiedAt=TREN)
        dodaj_kopiju(t, "tblPrijemnica", "PRJ-00001", PrijemnicaID="PRJ-00006", Stornirano="")
        self.assertEqual([x.kod for x in self.nalazi(t)], [])


# ---------------------------------------------------------------------------
def kolizija_otpremnice(t, isti_trenutak=True):
    """Otpremnica ST-00002 nosi broj 1/050726 (kao OTP-00001). Ispravka OTP-00001
    po broju obori obe; OTP-00001 se ponovo unese, blokovi ST-00002 ostaju na
    storniranoj OTP-00003."""
    o1, o3 = r(t, "tblOtpremnica", "OTP-00001"), r(t, "tblOtpremnica", "OTP-00003")
    o3["BrojOtpremnice"] = "1/050726"
    dodaj_kopiju(t, "tblOtpremnica", "OTP-00001", OtpremnicaID="OTP-00006", CreatedAt="2026-07-07 10:05:00",
                 ModifiedAt="2026-07-07 10:05:00")
    for otk in ("OTK-00001", "OTK-00002"):
        r(t, "tblOtkup", otk)["OtpremnicaID"] = "OTP-00006"
    o1.update(Stornirano="Da", ModifiedAt=TREN)
    o3.update(Stornirano="Da", ModifiedAt=TREN if isti_trenutak else "2026-07-06 18:00:00")


class KolateralniStornoPoBroju(Osnova):
    def test_otpremnica_druge_stanice_oborena(self):
        t = nova()
        kolizija_otpremnice(t)
        x = self.jedan(self.nalazi(t), KOD_KOL, "OTP-00003", GRESKA)
        self.assertIn("OTK-00004", x.opis)
        self.assertIn("ST-00002", x.opis)

    def test_otpremnica_bez_dokaza_trenutka_sumnja(self):
        t = nova()
        kolizija_otpremnice(t, isti_trenutak=False)
        self.jedan(self.nalazi(t), KOD_KOL, "OTP-00003", SUMNJA)

    def test_obe_stanice_ponovo_unele(self):
        t = nova()
        kolizija_otpremnice(t)
        dodaj_kopiju(t, "tblOtpremnica", "OTP-00003", OtpremnicaID="OTP-00007", Stornirano="",
                     BrojOtpremnice="2/050726")
        r(t, "tblOtkup", "OTK-00004")["OtpremnicaID"] = "OTP-00007"
        self.nema(self.nalazi(t), KOD_KOL)

    def test_zbirna_drugog_vozaca_oborena(self):
        t = nova()
        z1, z3 = r(t, "tblZbirna", "ZBR-00001"), r(t, "tblZbirna", "ZBR-00003")
        z3["BrojZbirne"] = "S1/050726"
        r(t, "tblOtpremnica", "OTP-00003")["BrojZbirne"] = "S1/050726"
        r(t, "tblPrijemnica", "PRJ-00003")["BrojZbirne"] = "S1/050726"
        dodaj_kopiju(t, "tblZbirna", "ZBR-00001", ZbirnaID="ZBR-00006")
        z1.update(Stornirano="Da", ModifiedAt=TREN)
        z3.update(Stornirano="Da", ModifiedAt=TREN)
        x = self.jedan(self.nalazi(t), KOD_KOL, "ZBR-00003", GRESKA)
        self.assertEqual(x.tabela, "tblZbirna")
        self.assertIn("OTP-00003", x.opis)

    def test_zbirna_ispravka_vozaca_pod_istim_brojem(self):
        """Zbirna sa vozacem druge stanice stornirana i ponovo uneta pod istim brojem sa
        pravim vozacem (pa jos jednom ispravljena): to je ispravka vlasnika, ne kolateral."""
        t = nova()
        r(t, "tblZbirna", "ZBR-00001").update(VozacID="ST-00002", Stornirano="Da", ModifiedAt="2026-07-05 12:00:00")
        dodaj_kopiju(t, "tblZbirna", "ZBR-00001", ZbirnaID="ZBR-00006", VozacID="ST-00001", Stornirano="Da",
                     ModifiedAt="2026-07-05 13:00:00")
        dodaj_kopiju(t, "tblZbirna", "ZBR-00001", ZbirnaID="ZBR-00007", VozacID="ST-00001", Stornirano="",
                     ModifiedAt="2026-07-05 13:00:01")
        self.nema(self.nalazi(t), KOD_KOL)

    def test_jedna_stanica_storno_pa_ponovni_unos(self):
        t = nova()
        r(t, "tblOtpremnica", "OTP-00001").update(Stornirano="Da", ModifiedAt=TREN)
        dodaj_kopiju(t, "tblOtpremnica", "OTP-00001", OtpremnicaID="OTP-00006", Stornirano="")
        for otk in ("OTK-00001", "OTK-00002"):
            r(t, "tblOtkup", otk)["OtpremnicaID"] = "OTP-00006"
        self.assertEqual([x.kod for x in self.nalazi(t)], [])


# ---------------------------------------------------------------------------
def kolizija_otkupa(t, isti_trenutak=True):
    """PWA broj 2/050726 dobio je i otkup KOOP-00004 na ST-00003. Storno bloka
    OTK-00004 po broju obori i njega; OTK-00004 se ponovo unese."""
    o4, o6 = r(t, "tblOtkup", "OTK-00004"), r(t, "tblOtkup", "OTK-00006")
    o6["BrojDokumenta"] = "2/050726"
    dodaj_kopiju(t, "tblOtkup", "OTK-00004", OtkupID="OTK-00007", BrojDokumenta="2/050726-2")
    o4.update(Stornirano="Da", ModifiedAt=TREN)
    o6.update(Stornirano="Da", ModifiedAt=TREN if isti_trenutak else "2026-07-06 18:00:00")


class StornoOtkupaPoBroju(Osnova):
    def test_blok_drugog_kooperanta_oboren(self):
        t = nova()
        kolizija_otkupa(t)
        x = self.jedan(self.nalazi(t), KOD_OTK, "OTK-00006", GRESKA)
        self.assertIn("KOOP-00004", x.opis)
        self.assertIn("OTP-00005", x.opis)

    def test_bez_dokaza_trenutka_sumnja(self):
        t = nova()
        kolizija_otkupa(t, isti_trenutak=False)
        self.jedan(self.nalazi(t), KOD_OTK, "OTK-00006", SUMNJA)

    def test_ispravka_kooperanta_pod_istim_brojem(self):
        """Otkup unet na pogresnog kooperanta, storniran i ponovo unet pod istim brojem na
        pravog (koji je kasnije jos jednom ispravljen): nije storno tudjeg bloka."""
        t = nova()
        r(t, "tblOtkup", "OTK-00004").update(Stornirano="Da", ModifiedAt="2026-07-05 12:00:00")
        dodaj_kopiju(t, "tblOtkup", "OTK-00004", OtkupID="OTK-00007", KooperantID="KOOP-00001", Stornirano="Da",
                     ModifiedAt="2026-07-05 13:00:00")
        dodaj_kopiju(t, "tblOtkup", "OTK-00004", OtkupID="OTK-00008", KooperantID="KOOP-00001", Stornirano="",
                     ModifiedAt="2026-07-05 13:00:05")
        self.nema(self.nalazi(t), KOD_OTK)

    def test_oba_kooperanta_ponovo_uneta(self):
        t = nova()
        kolizija_otkupa(t)
        dodaj_kopiju(t, "tblOtkup", "OTK-00006", OtkupID="OTK-00008", Stornirano="", BrojDokumenta="3/060726-2")
        self.nema(self.nalazi(t), KOD_OTK)


# ---------------------------------------------------------------------------
def ponistenje_tudje(t):
    """Prijemnica stanice 2 (PRJ-00006) nasledila je BrojZbirne S3/060726. Operater
    je PONISTI: kaskada po S3/060726 obori zbirnu, otpremnicu i pravu prijemnicu
    stanice 3 (PRJ-00005), a blok OTK-00006 oslobodi. Prava zbirna stanice 2
    (S2/060726) ostaje bez prijemnice."""
    dodaj_kopiju(t, "tblOtpremnica", "OTP-00003", OtpremnicaID="OTP-00006", Datum=D6, BrojOtpremnice="2/060726",
                 BrojZbirne="S2/060726", Kolicina=50, KolAmbalaze=20, BrutoKg=50,
                 CreatedAt="2026-07-06 12:00:00", ModifiedAt="2026-07-06 12:00:00")
    dodaj_kopiju(t, "tblZbirna", "ZBR-00003", ZbirnaID="ZBR-00006", Datum=D6, BrojZbirne="S2/060726",
                 UkupnoKolicina=50, UkupnoAmbalaze=20, CreatedAt="2026-07-06 12:00:00")
    dodaj_kopiju(t, "tblOtkup", "OTK-00004", OtkupID="OTK-00007", Datum=D6, Kolicina=50, KolAmbalaze=20, BrutoKg=50,
                 BrojDokumenta="2/060726", BrojZbirne="S2/060726", OtpremnicaID="OTP-00006",
                 BrojOtpremnice="2/060726")
    dodaj_kopiju(t, "tblPrijemnica", "PRJ-00005", PrijemnicaID="PRJ-00006", VozacID="ST-00002",
                 BrojPrijemnice="1/060726-3", Kolicina=Decimal("49.5"), KolAmbalaze=20, KolAmbVracena=20,
                 Stornirano="Da", ModifiedAt=TREN)
    for ime, id_ in (("tblZbirna", "ZBR-00005"), ("tblOtpremnica", "OTP-00005"), ("tblPrijemnica", "PRJ-00005")):
        r(t, ime, id_).update(Stornirano="Da", ModifiedAt=TREN)
    r(t, "tblOtkup", "OTK-00006").update(OtpremnicaID="", BrojZbirne="", BrojOtpremnice="", ModifiedAt=TREN)
    dodaj_sv(t, Mode="PONISTENJE_BEZ_ZAMENE", Status="COMPLETED", OldDocType="Prijemnica", OldDocID="PRJ-00006",
             OldBroj="1/060726-3", NewDocType="", ParentBroj="S3/060726", CompletedAt="2026-07-07 09:00:02",
             NeedsRecovery="Ne", Message="Ponistena prijemnica sa CELIM tokom zbirne S3/060726.")


def ponistenje_legit(t):
    """PONISTENJE prave prijemnice PRJ-00005: ceo tok stanice 3 otpada."""
    for ime, id_ in (("tblZbirna", "ZBR-00005"), ("tblOtpremnica", "OTP-00005"), ("tblPrijemnica", "PRJ-00005")):
        r(t, ime, id_).update(Stornirano="Da", ModifiedAt=TREN)
    r(t, "tblOtkup", "OTK-00006").update(OtpremnicaID="", BrojZbirne="", BrojOtpremnice="", ModifiedAt=TREN)
    dodaj_sv(t, Mode="PONISTENJE_BEZ_ZAMENE", Status="COMPLETED", OldDocType="Prijemnica", OldDocID="PRJ-00005",
             OldBroj="1/060726-2", NewDocType="", ParentBroj="S3/060726", NeedsRecovery="Ne")


def uskrsnuce_stanice_3(t):
    """Sledeci sync spakuje oslobodjen OTK-00006 u novu otpremnicu 3/060726-2 i
    zbirnu S3/060726-2 (bez prijemnice)."""
    dodaj_kopiju(t, "tblOtpremnica", "OTP-00005", OtpremnicaID="OTP-00007", BrojOtpremnice="3/060726-2",
                 BrojZbirne="S3/060726-2", Stornirano="", CreatedAt="2026-07-07 09:30:00",
                 ModifiedAt="2026-07-07 09:30:00")
    dodaj_kopiju(t, "tblZbirna", "ZBR-00005", ZbirnaID="ZBR-00007", BrojZbirne="S3/060726-2", Stornirano="",
                 CreatedAt="2026-07-07 09:30:00")
    r(t, "tblOtkup", "OTK-00006").update(OtpremnicaID="OTP-00007", BrojZbirne="S3/060726-2",
                                         BrojOtpremnice="3/060726-2")


def tura2_iste_kg(t, prijemnica=True):
    """Druga tura stanice 3 istog dana, istih kg, uneta PRE storna (6.7. 14:00)."""
    dodaj_kopiju(t, "tblOtpremnica", "OTP-00005", OtpremnicaID="OTP-00007", BrojOtpremnice="3/060726-2",
                 BrojZbirne="S3/060726-2", Stornirano="", CreatedAt="2026-07-06 14:00:00",
                 ModifiedAt="2026-07-06 14:00:00")
    dodaj_kopiju(t, "tblZbirna", "ZBR-00005", ZbirnaID="ZBR-00007", BrojZbirne="S3/060726-2", Stornirano="",
                 CreatedAt="2026-07-06 14:00:00")
    dodaj_kopiju(t, "tblOtkup", "OTK-00006", OtkupID="OTK-00007", OtpremnicaID="OTP-00007", BrojZbirne="S3/060726-2",
                 BrojDokumenta="3/060726-2", BrojOtpremnice="3/060726-2", Stornirano="")
    if prijemnica:
        dodaj_kopiju(t, "tblPrijemnica", "PRJ-00005", PrijemnicaID="PRJ-00007", BrojPrijemnice="1/060726-3",
                     BrojZbirne="S3/060726-2", Stornirano="")


class PonistenjeObaraTudjuPrijemnicu(Osnova):
    def test_ponistenje_pogresno_vezane_prijemnice(self):
        t = nova()
        ponistenje_tudje(t)
        x = self.jedan(self.nalazi(t), KOD_PON, "PRJ-00005", GRESKA)
        self.assertIn("COR-00001", x.opis)
        self.assertIn("ST-00002", x.opis)
        self.assertIn("S2/060726 (ZBR-00006", x.predlog)

    def test_legitimno_ponistenje_jedne_prijemnice(self):
        t = nova()
        ponistenje_legit(t)
        self.nema(self.nalazi(t), KOD_PON)

    def test_legitimno_ponistenje_zbirna_sa_vozacem_druge_stanice(self):
        """P4: otpremnica stanice 3 sa vozacem ST-00002 daje zbirnu S3/.. sa VozacID
        ST-00002; prijemnica nosi stanicu iz prefiksa (ST-00003). Nije tudja."""
        t = nova()
        r(t, "tblZbirna", "ZBR-00005")["VozacID"] = "ST-00002"
        r(t, "tblOtpremnica", "OTP-00005")["VozacID"] = "ST-00002"
        ponistenje_legit(t)
        self.nema(self.nalazi(t), KOD_PON)

    def test_bez_traga_ispravka_prijemnice_iste_stanice(self):
        """Prijemnica ispravljena pod novim brojem (ista stanica i kupac), pa zbirna
        oborena bez kaskade na prijemnicu: to je PRJ_AKTIVNA_NA_STORNIRANOJ_ZBIRNOJ
        (modul prijemnica), ne kaskada nad tudjom prijemnicom."""
        t = nova()
        for ime, id_ in (("tblZbirna", "ZBR-00002"), ("tblOtpremnica", "OTP-00002"), ("tblPrijemnica", "PRJ-00002")):
            r(t, ime, id_)["Stornirano"] = "Da"
        dodaj_kopiju(t, "tblPrijemnica", "PRJ-00002", PrijemnicaID="PRJ-00006", BrojPrijemnice="1/050726-4",
                     Stornirano="")
        self.nema(self.nalazi(t), KOD_PON)

    def test_bez_storno_veza_dve_prijemnice_na_oborenoj_zbirni(self):
        t = nova()
        ponistenje_tudje(t)
        t["tblStornoVeze"] = []
        r(t, "tblPrijemnica", "PRJ-00006")["Stornirano"] = ""        # tudja ostala aktivna na oborenoj zbirni
        x = self.jedan(self.nalazi(t), KOD_PON, "PRJ-00006", SUMNJA)
        self.assertIn("bez traga u tblStornoVeze", x.opis)


class AutoUskrsnucePosleStorna(Osnova):
    def test_sync_vratio_ponisten_tok(self):
        t = nova()
        ponistenje_tudje(t)
        uskrsnuce_stanice_3(t)
        nn = self.nalazi(t)
        x = self.jedan(nn, KOD_USK, "OTP-00007", SUMNJA)
        self.assertIn("OTK-00006", x.opis)
        self.assertIn("S3/060726-2 nema aktivnu prijemnicu", x.opis)
        self.jedan(nn, KOD_PON, "PRJ-00005", GRESKA)

    def test_dupli_roba_dva_puta(self):
        """Blok OTK-00005 bio je na duplikatu OTP-00006 umesto na OTP-00004. DUPLI
        stornira duplikat i oslobodi blok, sync ga spakuje u OTP-00007, a OTP-00004
        ostaje bez blokova: ista roba je dva puta u lancu."""
        t = nova()
        dodaj_kopiju(t, "tblOtpremnica", "OTP-00004", OtpremnicaID="OTP-00006", BrojOtpremnice="1/060726-2",
                     BrojZbirne="S1/060726-2", Stornirano="Da", CreatedAt="2026-07-06 12:00:00", ModifiedAt=TREN)
        dodaj_kopiju(t, "tblZbirna", "ZBR-00004", ZbirnaID="ZBR-00006", BrojZbirne="S1/060726-2", Stornirano="Da",
                     ModifiedAt=TREN)
        dodaj_kopiju(t, "tblOtpremnica", "OTP-00004", OtpremnicaID="OTP-00007", BrojOtpremnice="1/060726-3",
                     BrojZbirne="S1/060726-3", CreatedAt="2026-07-07 09:30:00")
        dodaj_kopiju(t, "tblZbirna", "ZBR-00004", ZbirnaID="ZBR-00007", BrojZbirne="S1/060726-3")
        r(t, "tblOtkup", "OTK-00005").update(OtpremnicaID="OTP-00007", BrojZbirne="S1/060726-3",
                                             BrojOtpremnice="1/060726-3")
        dodaj_sv(t, Mode="DUPLI_FANTOM", Status="COMPLETED", OldDocType="Otpremnica", OldDocID="OTP-00006",
                 OldBroj="1/060726-2", NewDocType="", ParentBroj="S1/060726-2", NeedsRecovery="Ne")
        x = self.jedan(self.nalazi(t), KOD_USK, "OTP-00007", GRESKA)
        self.assertIn("OTP-00004", x.predlog)

    def test_druga_tura_istog_dana_drugih_kg(self):
        t = nova()
        ponistenje_legit(t)
        dodaj_kopiju(t, "tblOtpremnica", "OTP-00005", OtpremnicaID="OTP-00007", BrojOtpremnice="3/060726-2",
                     BrojZbirne="S3/060726-2", Stornirano="", Kolicina=30, KolAmbalaze=12,
                     CreatedAt="2026-07-07 09:30:00")
        dodaj_kopiju(t, "tblOtkup", "OTK-00006", OtkupID="OTK-00007", Kolicina=30, KolAmbalaze=12,
                     OtpremnicaID="OTP-00007", BrojZbirne="S3/060726-2", BrojDokumenta="3/060726-2")
        self.nema(self.nalazi(t), KOD_USK)

    def test_uskrsnuce_bez_audit_kolona_sumnja(self):
        t = nova()
        ponistenje_tudje(t)
        uskrsnuce_stanice_3(t)
        self.jedan(self.nalazi(t, BEZ_VREMENA_OTP), KOD_USK, "OTP-00007", SUMNJA)

    def test_druga_tura_istih_kg_pre_storna(self):
        t = nova()
        ponistenje_legit(t)
        tura2_iste_kg(t)
        self.nema(self.nalazi(t), KOD_USK)

    def test_druga_tura_istih_kg_bez_vremena_sa_prijemnicom(self):
        """Bez CreatedAt na otpremnici: druga tura sa prijemnicom je isporucen tok."""
        t = nova()
        ponistenje_legit(t)
        tura2_iste_kg(t)
        self.nema(self.nalazi(t, BEZ_VREMENA_OTP), KOD_USK)

    def test_druga_tura_istih_kg_bez_vremena_blok_jos_slobodan(self):
        """Oslobodjeni blok je i dalje nevezan: sync ga nije spakovao, nema uskrsnuca."""
        t = nova()
        ponistenje_legit(t)
        tura2_iste_kg(t, prijemnica=False)
        self.nema(self.nalazi(t, BEZ_VREMENA_OTP), KOD_USK)

    def test_druga_tura_posle_storna_sa_prijemnicom(self):
        """Ponisten tok (i blok storniran), pa kasnije istog dana stvarna druga tura istih
        kg sa prijemnicom: zbirna od A ima prijemnicu, tok je isporucen."""
        t = nova()
        ponistenje_legit(t)
        r(t, "tblOtkup", "OTK-00006")["Stornirano"] = "Da"
        tura2_iste_kg(t)
        r(t, "tblOtpremnica", "OTP-00007")["CreatedAt"] = "2026-07-07 11:00:00"
        self.nema(self.nalazi(t), KOD_USK)

    def test_bez_vremena_samo_najnovija_kandidatkinja(self):
        """Bez CreatedAt: druga tura istih kg (bez prijemnice, uneta ranije) i prava
        uskrsla otpremnica OTP-00008 -- prijavljuje se samo najnovija."""
        t = nova()
        ponistenje_legit(t)
        tura2_iste_kg(t, prijemnica=False)
        dodaj_kopiju(t, "tblOtpremnica", "OTP-00005", OtpremnicaID="OTP-00008", BrojOtpremnice="3/060726-3",
                     BrojZbirne="S3/060726-3", Stornirano="")
        dodaj_kopiju(t, "tblZbirna", "ZBR-00005", ZbirnaID="ZBR-00008", BrojZbirne="S3/060726-3", Stornirano="")
        r(t, "tblOtkup", "OTK-00006").update(OtpremnicaID="OTP-00008", BrojZbirne="S3/060726-3",
                                             BrojOtpremnice="3/060726-3")
        self.jedan(self.nalazi(t, BEZ_VREMENA_OTP), KOD_USK, "OTP-00008", SUMNJA)

    def test_druga_tura_bez_storna_nije_uskrsnuce(self):
        t = nova()
        dodaj_kopiju(t, "tblOtpremnica", "OTP-00005", OtpremnicaID="OTP-00007", BrojOtpremnice="3/060726-2",
                     BrojZbirne="S3/060726-2", CreatedAt="2026-07-06 14:00:00")
        dodaj_kopiju(t, "tblOtkup", "OTK-00006", OtkupID="OTK-00007", OtpremnicaID="OTP-00007",
                     BrojZbirne="S3/060726-2", BrojDokumenta="3/060726-2")
        self.nema(self.nalazi(t), KOD_USK)


# ---------------------------------------------------------------------------
class DelimicnoStorniranDokument(Osnova):
    def test_klasa_ii_stornirana_klasa_i_aktivna(self):
        t = nova()
        dodaj_kopiju(t, "tblOtpremnica", "OTP-00001", OtpremnicaID="OTP-00006", Klasa="II", Kolicina=20,
                     KolAmbalaze=8, BrutoKg=20, Stornirano="Da")       # isti CreatedAt: jedan unos
        x = self.jedan(self.nalazi(t), KOD_DEL, "OTP-00001, OTP-00006", SUMNJA)
        self.assertIn("kl. II", x.opis)

    def test_storno_pa_ponovni_unos_iste_klase(self):
        t = nova()
        r(t, "tblOtpremnica", "OTP-00001")["Stornirano"] = "Da"
        dodaj_kopiju(t, "tblOtpremnica", "OTP-00001", OtpremnicaID="OTP-00006", Stornirano="",
                     CreatedAt="2026-07-05 11:00:00")
        self.nema(self.nalazi(t), KOD_DEL)

    def test_ponovni_unos_druge_klase_kasnije(self):
        t = nova()
        r(t, "tblOtpremnica", "OTP-00001")["Stornirano"] = "Da"
        dodaj_kopiju(t, "tblOtpremnica", "OTP-00001", OtpremnicaID="OTP-00006", Klasa="II", Stornirano="",
                     CreatedAt="2026-07-05 11:00:00")
        self.nema(self.nalazi(t), KOD_DEL)


# ---------------------------------------------------------------------------
class NedovrsenStornoKontekst(Osnova):
    def test_ispravka_bez_nove_prijemnice(self):
        t = nova()
        r(t, "tblPrijemnica", "PRJ-00002").update(Stornirano="Da", ModifiedAt="2026-07-07 09:00:01")
        dodaj_sv(t)
        x = self.jedan(self.nalazi(t), KOD_KTX, "COR-00001", SUMNJA)
        self.assertIn("novi nije snimljen", x.opis)
        self.assertIn("S1/050726-2", x.opis)

    def test_zavrsena_ispravka_nije_nalaz(self):
        t = nova()
        r(t, "tblPrijemnica", "PRJ-00002")["Stornirano"] = "Da"
        dodaj_kopiju(t, "tblPrijemnica", "PRJ-00002", PrijemnicaID="PRJ-00006", Stornirano="")
        dodaj_sv(t, Status="COMPLETED", NewBroj="1/050726-2", NeedsRecovery="Ne")
        dodaj_sv(t, CorrectionID="COR-00002", Mode="RESI_KASNIJE", Status="COMPLETED")
        dodaj_sv(t, CorrectionID="COR-00003", Status="CANCELLED")
        self.nema(self.nalazi(t), KOD_KTX)

    def test_pending_sa_naslednikom_je_info(self):
        t = nova()
        r(t, "tblPrijemnica", "PRJ-00002")["Stornirano"] = "Da"
        dodaj_kopiju(t, "tblPrijemnica", "PRJ-00002", PrijemnicaID="PRJ-00006", Stornirano="")
        dodaj_sv(t)
        x = self.jedan(self.nalazi(t), KOD_KTX, "COR-00001", INFO)
        self.assertIn("PRJ-00006", x.opis)

    def test_parkirano_pa_reseno_kasnijom_ispravkom(self):
        t = nova()
        dodaj_sv(t, Mode="RESI_KASNIJE", NewDocType="", Message="Prijemnica parkirana za kasnije.")
        r(t, "tblPrijemnica", "PRJ-00002")["Stornirano"] = "Da"
        dodaj_kopiju(t, "tblPrijemnica", "PRJ-00002", PrijemnicaID="PRJ-00006", Stornirano="")
        dodaj_sv(t, CorrectionID="COR-00002", Status="COMPLETED", NewBroj="1/050726-2", NeedsRecovery="Ne",
                 CreatedAt="2026-07-08 09:00:00")
        x = self.jedan(self.nalazi(t), KOD_KTX, "COR-00001", INFO)
        self.assertIn("COR-00002", x.opis)

    def test_failed_pa_uspeo_ponovni_pokusaj(self):
        t = nova()
        r(t, "tblPrijemnica", "PRJ-00002")["Stornirano"] = "Da"
        dodaj_kopiju(t, "tblPrijemnica", "PRJ-00002", PrijemnicaID="PRJ-00006", Stornirano="")
        dodaj_sv(t, Status="FAILED")
        dodaj_sv(t, CorrectionID="COR-00002", Status="COMPLETED", NewBroj="1/050726-2",
                 CreatedAt="2026-07-07 09:05:00")
        self.jedan(self.nalazi(t), KOD_KTX, "COR-00001", INFO)

    def test_parkirano_nikad_reseno(self):
        t = nova()
        dodaj_sv(t, Mode="RESI_KASNIJE", NewDocType="", Message="Prijemnica parkirana za kasnije.")
        self.jedan(self.nalazi(t), KOD_KTX, "COR-00001", SUMNJA)

    def test_manual_required_ukrsta_druge_nalaze(self):
        t = nova()
        r(t, "tblOtkup", "OTK-00002")["Stornirano"] = "Da"
        dodaj_sv(t, Mode="DUPLI_FANTOM", Status="MANUAL_REQUIRED", OldDocType="Otpremnica", OldDocID="OTP-00001",
                 OldBroj="1/050726", NewDocType="", ParentBroj="S1/050726",
                 RecoveryAction="Odluci o osirocenoj prijemnici/paletama (reveze ili storno).")
        x = self.jedan(self.nalazi(t), KOD_KTX, "COR-00001", SUMNJA)
        self.assertIn(KOD_BLOK, x.opis)
        self.assertIn("reveze ili storno", x.opis)


# ---------------------------------------------------------------------------
class BrojZbirneOtkupaNaStorniranojZbirni(Osnova):
    def test_zbirna_stornirana_otkup_zadrzao_broj(self):
        t = nova()
        r(t, "tblZbirna", "ZBR-00004")["Stornirano"] = "Da"
        r(t, "tblOtpremnica", "OTP-00004")["BrojZbirne"] = ""        # otpremnica odvezana, otkup nije
        x = self.jedan(self.nalazi(t), KOD_BZ, "OTK-00005", INFO)
        self.assertIn("S1/060726", x.opis)

    def test_zbirna_ponovo_uneta_istog_broja(self):
        t = nova()
        r(t, "tblZbirna", "ZBR-00004")["Stornirano"] = "Da"
        dodaj_kopiju(t, "tblZbirna", "ZBR-00004", ZbirnaID="ZBR-00006", Stornirano="")
        self.nema(self.nalazi(t), KOD_BZ)


class OsirocenoNaFakturiZastarelo(Osnova):
    KOLONE = {"tblFakture": fixture.KOLONE["tblFakture"] + ["OsirocenoOd"]}

    def faktura(self, t, osiroceno_linija):
        t["tblFakture"].append(dict(FakturaID="FAK-00001", BrojFakture="F-1/2026", Datum=D6, KupacID=KUPAC,
                                    Iznos=99 * fixture.CENA_PRJ, Status="Izdata", Stornirano="",
                                    OsirocenoOd="PRJ-00009"))
        t["tblFakturaStavke"].append(dict(StavkaID="FKS-00001", FakturaID="FAK-00001", PrijemnicaID="PRJ-00001",
                                          Kolicina=99, Cena=fixture.CENA_PRJ, Klasa="I", BrojPrijemnice="1/050726",
                                          Stornirano="", OsirocenoOd=osiroceno_linija))

    def test_relink_ocistio_liniju_ne_i_fakturu(self):
        t = nova()
        self.faktura(t, "")
        self.jedan(self.nalazi(t, self.KOLONE), KOD_FAK, "FAK-00001", INFO)

    def test_linija_jos_osirocena(self):
        t = nova()
        self.faktura(t, "PRJ-00009")
        self.nema(self.nalazi(t, self.KOLONE), KOD_FAK)


# ---------------------------------------------------------------------------
class Robusnost(Osnova):
    def test_kolona_fali_preskoceno(self):
        t = nova()
        dodaj_sv(t)
        kol = {"tblStornoVeze": [c for c in fixture.KOLONE["tblStornoVeze"] if c != "Status"],
               "tblOtkup": [c for c in fixture.KOLONE["tblOtkup"] if c != "OtpremnicaID"]}
        nn = self.nalazi(t, kol)
        x = self.jedan(nn, KOD_KTX + "_PRESKOCENO", "", INFO)
        self.assertIn("tblStornoVeze.Status", x.opis)
        x = self.jedan(nn, KOD_BLOK + "_PRESKOCENO", "", INFO)
        self.assertIn("tblOtkup.OtpremnicaID", x.opis)

    def test_prazna_tabela_bez_kolone_tiho(self):
        t = nova()
        kol = {"tblStornoVeze": ["CorrectionID", "OldBroj"]}
        self.assertEqual([x.kod for x in self.nalazi(t, kol)], [])

    def test_tabele_ne_postoje(self):
        t = nova()
        for ime in ("tblStornoVeze", "tblOtkup", "tblZbirna", "tblFakture", "tblFakturaStavke"):
            del t[ime]
        self.assertEqual([x.kod for x in self.nalazi(t)], [])

    def test_smece_u_podacima(self):
        t = nova()
        ponistenje_tudje(t)
        uskrsnuce_stanice_3(t)
        kolizija_otkupa(t)
        r(t, "tblOtkup", "OTK-00001").update(Kolicina="abc", Datum=None, KolAmbalaze="x")
        r(t, "tblOtpremnica", "OTP-00002").update(Datum="nije datum", Kolicina="", Stornirano="Da")
        r(t, "tblPrijemnica", "PRJ-00001").update(Datum=None, BrojZbirne="  s1/ 050726 ", Stornirano="Da")
        r(t, "tblZbirna", "ZBR-00001").update(UkupnoKolicina="1.2.3", Datum="")
        dodaj_sv(t, CorrectionID="", Mode="", Status="FAILED", OldDocType="Revers", OldBroj="", OldDocID="")
        dodaj_sv(t, CorrectionID="COR-00009", Mode="DUPLI_FANTOM", Status="PENDING", OldDocType="Zbirna",
                 OldBroj="nepostoji", CreatedAt="31.13.2026 99:99")
        t["tblOtkup"].append(dict(OtkupID="", BrojDokumenta="2/050726", Stornirano="Da", ModifiedAt=TREN))
        nn = self.nalazi(t)
        self.assertTrue(all(x.ozbiljnost in (GRESKA, SUMNJA, INFO) for x in nn))


class BezAuditKolona(Osnova):
    """Sveska bez CreatedAt/ModifiedAt: nula nalaza na cistoj, pozitivi kao SUMNJA."""

    def test_cista(self):
        self.assertEqual([x.kod for x in self.nalazi(nova(), BEZ_AUDITA)], [])

    def test_pozitivni(self):
        for gradi, kod, id_ in ((kolizija_prijemnice, KOD_PRJ, "PRJ-00003"),
                                (kolizija_otpremnice, KOD_KOL, "OTP-00003"),
                                (kolizija_otkupa, KOD_OTK, "OTK-00006"),
                                (ponistenje_tudje, KOD_PON, "PRJ-00005")):
            t = nova()
            gradi(t)
            x = [y for y in self.nalazi(t, BEZ_AUDITA) if y.kod == kod]
            self.assertEqual([y.id for y in x], [id_], kod)
            if kod != KOD_PON:
                self.assertEqual(x[0].ozbiljnost, SUMNJA, kod)


class Smece(Osnova):
    def test_nasumicno_smece_bez_pada(self):
        rnd = random.Random(7)
        junk = [None, "", "abc", "1e30", -5, 0, "12,5", datetime(2026, 7, 5, 10, 0), date(2026, 1, 1), "  ", "Da",
                "S9/999999", "31.13.2026 99:99", 45000.5, True]
        gradi = [kolizija_prijemnice, kolizija_otpremnice, kolizija_otkupa, ponistenje_tudje]
        for it in range(30):
            t = nova()
            for g in rnd.sample(gradi, 2):
                try:
                    g(t)
                except KeyError:
                    pass
            dodaj_sv(t, Mode=rnd.choice(["PONISTENJE_BEZ_ZAMENE", "DUPLI_FANTOM", "RESI_KASNIJE", "ISPRAVKA_ODMAH"]),
                     Status=rnd.choice(["PENDING", "FAILED", "MANUAL_REQUIRED", "COMPLETED"]),
                     OldDocType=rnd.choice(["Otpremnica", "Zbirna", "Prijemnica", "Revers", "Otkup", ""]),
                     OldBroj=rnd.choice(["1/050726", "S1/050726", "S3/060726", "", "x"]))
            for _ in range(25):
                ime = rnd.choice([i for i in t if t[i]])
                red = rnd.choice(t[ime])
                red[rnd.choice(list(red))] = rnd.choice(junk)
            with self.subTest(it=it):
                self.nalazi(t)          # nalazi() proverava da nema MODUL_PAO


if __name__ == "__main__":
    unittest.main()
