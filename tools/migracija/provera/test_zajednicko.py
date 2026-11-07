import os
import tempfile
import unittest
from datetime import date

from citac import procitaj_svesku

from provera import fixture
from provera.pokretac import pokreni, pokreni_nad_tabelama
from provera.zajednicko import Kontekst, cifre, norm_broj, parse_broj


def tabele_iz(t, kolone=None):
    """fixture dict -> tabele kakve vraca citac (preko pravog .xlsx)."""
    with tempfile.TemporaryDirectory() as dirn:
        p = os.path.join(dirn, "s.xlsx")
        fixture.snimi(p, t, kolone)
        return procitaj_svesku(p)


class Brojevi(unittest.TestCase):
    def test_otpremnica(self):
        b = parse_broj("1/180626-3")
        self.assertEqual((b.s, b.x, b.dan, b.seq), (False, 1, date(2026, 6, 18), 3))

    def test_zbirna_sa_s(self):
        b = parse_broj("S3/120626-2")
        self.assertEqual((b.s, b.x, b.dan, b.seq), (True, 3, date(2026, 6, 12), 2))

    def test_prvi_bez_nastavka(self):
        self.assertEqual(parse_broj("2/050726").seq, 1)

    def test_nije_broj(self):
        self.assertIsNone(parse_broj("HL-050726-101500"))
        self.assertIsNone(parse_broj(""))

    def test_nepostojeci_datum(self):
        self.assertIsNone(parse_broj("1/310626").dan)

    def test_norm(self):
        self.assertEqual(norm_broj(" s3/ 120626-2 "), "s3/120626-2")

    def test_cifre(self):
        self.assertEqual(cifre("ST-00002"), 2)


class CistaSveska(unittest.TestCase):
    def test_kontekst(self):
        ctx = Kontekst(tabele_iz(fixture.osnovna()))
        self.assertEqual(ctx.stanica_po_broju[2], "ST-00002")
        self.assertEqual(ctx.malina_kupac, fixture.KUPAC)
        self.assertEqual(len(ctx.t["tblPrijemnica"]), 5)
        self.assertEqual(ctx.stanica_iz_broja("S3/060726"), "ST-00003")

    def test_bez_nalaza(self):
        """Cista sveska: NIJEDAN modul ne sme da prijavi nista (lazni pozitivi)."""
        ctx = pokreni_nad_tabelama(tabele_iz(fixture.osnovna()))
        self.assertEqual([(x.modul, x.kod, x.id, x.opis[:200]) for x in ctx.nalazi], [])

    def test_pokreni_ne_menja_svesku(self):
        with tempfile.TemporaryDirectory() as dirn:
            p = os.path.join(dirn, "s.xlsx")
            fixture.snimi(p, fixture.osnovna())
            ctx = pokreni(p)
            self.assertEqual(len(ctx.sha), 64)
