"""Provere PRIJEMNICE prema zbirnoj (tblPrijemnica <-> tblZbirna) -- SAMO CITA.

Polazni slucaj: btnUnosPrij_Click ne prazni txtBrojZbirnePrij (vba-v2.2.3 ..
vba-v2.28.4), pa prijemnica nasledi broj zbirne od prethodnog unosa: jedna zbirna
dobije dve prijemnice, a prava nijednu.
  (a) druga stanica  -> PRJ_ZBIRNA_DRUGE_STANICE (ili PRJ_ZASTAREO_VOZAC kad je
      zbirna tacna, a vozac zaostao u zajednickom cmbVozac);
  (b) ista stanica   -> PRJ_DVE_NA_ZBIRNOJ_ISTA_STANICA (druga tura -2 ili
      sledeci dan);
  isti mehanizam bez druge aktivne prijemnice (prethodna stornirana)
                     -> PRJ_NASLEDJENA_ZBIRNA.
Prava zbirna je ZBR_ZIVA_BEZ_PRIJEMNICE. Predlog ispravke je globalna 1:1 dodela
pogresno vezanih prijemnica i zivih zbirnih bez prijemnice (isti VozacID i
KupacID, ista klasa, Datum zbirne u [Datum-3, Datum+1], gajbe jednake, manjak kg
u pragu koji se meri iz iste sveske).

Dokument = redovi iste (KupacID, BrojPrijemnice, Datum, BrojZbirne, storno); svaka
Klasa najvise jednom (SavePrijemnicaMulti_TX: Klasa I i II dele broj). Zbirna isto,
po (BrojZbirne, VozacID, Datum, storno). Veza prijemnica -> zbirna je PO BROJU.
Redovi testne suite (TST-PRO, KUP-90001, VOZ-90001, ST-90001, 'Test Jabuka') se
izdvajaju (TESTNI_REDOVI_U_PRODUKCIJI) i ne ulaze ni u jednu drugu proveru.
"""

import re
from bisect import bisect_left
from collections import defaultdict
from datetime import date, datetime, timedelta
from decimal import Decimal

from citac import kljuc, polje, tekst

from provera.zajednicko import GRESKA, INFO, SUMNJA, cifre, d, dan_txt, ima_broj, k, kg_txt, n, norm_broj, parse_broj, s, storno

TPRJ, TZBR, TOTP = "tblPrijemnica", "tblZbirna", "tblOtpremnica"
TPLS, TSTA, TFKS, TSV = "tblPaletaStavka", "tblStanice", "tblFakturaStavke", "tblStornoVeze"

PRAG_KG_PODR = Decimal(5)          # % manjka kad je uzorak cistih parova mali
PRAG_KG_MIN, PRAG_KG_MAX = Decimal(3), Decimal(10)
VISAK_TOL = Decimal(1)             # % viska (prijemnica > zbirna) koji se trpi
PRAG_DANA_PODR, PRAG_DANA_MAX = 3, 7
MIN_UZORAK = 20
PROZOR_PRE, PROZOR_POSLE = 3, 1    # Datum zbirne-kandidata u [Datum prijemnice - 3, + 1]
DATUM_ZASTITE = date(2026, 7, 1)   # vba-v2.8.6: cena i gajbe obavezni uz VALIDACIJA_UNOSA
CENA_SISTEMSKI = 20                # vise od pola prijemnica bez cene (i bar 20) = praksa, ne greska unosa
DA = ("da", "true", "1", "yes")

_PRJ = ["PrijemnicaID", "BrojZbirne", "Stornirano"]
_ZBR = ["ZbirnaID", "BrojZbirne", "Stornirano"]
ZAHTEVI = {
    "PRJ_NASLEDJENA_ZBIRNA": {TPRJ: _PRJ + ["BrojPrijemnice", "Kolicina"], TZBR: _ZBR + ["UkupnoKolicina"]},
    "PRJ_DVE_NA_ZBIRNOJ_ISTA_STANICA": {TPRJ: _PRJ + ["VozacID", "KupacID", "Kolicina"],
                                        TZBR: _ZBR + ["VozacID", "KupacID", "UkupnoKolicina"]},
    "PRJ_ZBIRNA_DRUGE_STANICE": {TPRJ: _PRJ + ["VozacID"], TZBR: _ZBR + ["VozacID"]},
    "PRJ_ZASTAREO_VOZAC": {TPRJ: _PRJ + ["VozacID"], TZBR: _ZBR + ["VozacID"]},
    "ZBR_ZIVA_BEZ_PRIJEMNICE": {TZBR: _ZBR, TOTP: ["BrojZbirne", "Stornirano"], TPRJ: _PRJ},
    "PRJ_KUPAC_RAZLICIT_OD_ZBIRNE": {TPRJ: _PRJ + ["KupacID"], TZBR: _ZBR + ["KupacID"]},
    "PRJ_DATUM_ZASTAREO_ILI_PRE_ZBIRNE": {TPRJ: _PRJ + ["Datum"]},
    "PRJ_ROBA_RAZLICITA_OD_ZBIRNE": {TPRJ: _PRJ + ["VrstaVoca", "SortaVoca", "TipAmbalaze"],
                                     TZBR: _ZBR + ["VrstaVoca", "SortaVoca", "TipAmbalaze"]},
    "PRJ_KLASA_RAZLICITA_OD_ZBIRNE": {TPRJ: _PRJ + ["Klasa"], TZBR: _ZBR + ["Klasa"]},
    "PRJ_ZBIRNA_NE_POSTOJI_ILI_STORNIRANA": {TPRJ: _PRJ, TZBR: _ZBR},
    "PRJ_AKTIVNA_NA_STORNIRANOJ_ZBIRNOJ": {TPRJ: _PRJ, TZBR: _ZBR},
    "PRJ_BROJ_ZBIRNE_NENORMALIZOVAN": {TPRJ: _PRJ},
    "PRJ_DUPLI_BROJ_PRIJEMNICE": {TPRJ: _PRJ + ["BrojPrijemnice", "KupacID"]},
    "PRJ_CENA_NULA": {TPRJ: _PRJ + ["Cena"]},
    "PRJ_KG_GAJBE_VAN_OPSEGA": {TPRJ: _PRJ + ["Kolicina", "KolAmbalaze"],
                                TZBR: _ZBR + ["UkupnoKolicina", "UkupnoAmbalaze"]},
    "PRJ_GAJBE_NULA_ILI_BEZ_TIPA": {TPRJ: _PRJ + ["KolAmbalaze", "TipAmbalaze"]},
    "PRJ_PREVEZANA_BEZ_PROVERE": {TPRJ: _PRJ, TPLS: ["PrijemnicaID", "BrojZbirne", "Stornirano"]},
    "PRJ_ISPRAVKA_NA_DRUGOJ_ZBIRNOJ": {TPRJ: _PRJ + ["BrojPrijemnice"]},
    "PRJ_AUTO_LANAC_NEPOTPUN": {TOTP: ["OtpremnicaID", "StanicaID", "BrojZbirne", "Stornirano"], TPRJ: _PRJ},
    "PRIJEMNICA_POKRIVA_DVE_TURE": {TPRJ: _PRJ + ["Kolicina"], TZBR: _ZBR + ["UkupnoKolicina"]},
}

# --- testna suite (vba-v2.11.0:src-vba/modBusinessFlowProTests.bas:42-54) ---------------
TEST_TABELE = {"tblOtkup": "OtkupID", "tblOtpremnica": "OtpremnicaID", "tblZbirna": "ZbirnaID",
               "tblPrijemnica": "PrijemnicaID"}
_TEST_ID = frozenset(("kup-90001", "voz-90001", "st-90001", "koop-90001", "kul-90001"))
_TEST_ID_KOL = ("KupacID", "VozacID", "StanicaID", "KooperantID", "KulturaID")
_TEST_BROJ_KOL = ("PrijemnicaID", "ZbirnaID", "OtpremnicaID", "OtkupID", "BrojPrijemnice", "BrojZbirne",
                  "BrojOtpremnice", "BrojDokumenta")


def _je_test(r):
    if any(k(r, c) in _TEST_ID for c in _TEST_ID_KOL):
        return True
    if any(k(r, c).startswith(("tst-", "stst-")) for c in _TEST_BROJ_KOL):
        return True
    return k(r, "VrstaVoca") == "test jabuka"


# --- pomocne --------------------------------------------------------------------------
def _seq(id_):
    c = cifre(id_)
    return c if c is not None else 10 ** 12


def _klase(r):
    v = re.sub(r"\s+", "", k(r, "Klasa"))
    if not v:
        return frozenset()
    return frozenset(x for x in re.split(r"[/+,]", v) if x)


def _kompat(pk, zk):
    """Klase prijemnice su pokrivene zbirnom (prazno = nepoznato = prolazi)."""
    return not pk or not zk or pk <= zk


def _kl(kl):
    return "/".join(sorted(x.upper() for x in kl)) or "?"


def _manjak(zkg, pkg):
    if zkg is None or zkg <= 0:
        return None
    return (zkg - pkg) / zkg * 100


def _pct(m):
    return "?" if m is None else format(round(m, 1), "f") + "%"


def _g(v):
    return kg_txt(v)


def _perc(vrednosti, p):
    v = sorted(vrednosti)
    if not v:
        return None
    return v[min(len(v) - 1, int(round((len(v) - 1) * p)))]


def _lev(a, b):
    if a == b:
        return 0
    if abs(len(a) - len(b)) > 2:
        return 3
    pre = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(pre[j] + 1, cur[j - 1] + 1, pre[j - 1] + (ca != cb)))
        pre = cur
    return pre[-1]


def _susedan(a, b):
    """b je 'sledeci' broj posle a: isti S{x}, isti dan i seq+1, ili sledeci dan sa seq 1."""
    pa, pb = parse_broj(a), parse_broj(b)
    if not pa or not pb or pa.s != pb.s or pa.x != pb.x or not pa.dan or not pb.dan:
        return False
    if pa.dan == pb.dan and pb.seq == pa.seq + 1:
        return True
    return pb.dan == pa.dan + timedelta(days=1) and pb.seq == 1


def _kanon(v):
    """Broj kakav pisac pravi: bez razmaka, vodece S veliko."""
    t = re.sub(r"\s+", "", tekst(v))
    return "S" + t[1:] if t[:1] == "s" else t


def _vreme(r):
    v = polje(r, "CreatedAt")
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, date):
        return v.isoformat()
    t = tekst(v)
    return t if re.match(r"^\d{4}-\d{2}-\d{2}", t) else ""


def _skrati(lista, mx=10):
    lista = list(lista)
    return ", ".join(lista[:mx]) + (" (+%d)" % (len(lista) - mx) if len(lista) > mx else "")


class Dok:
    """Logicki dokument (jedan ili vise redova, svaka klasa najvise jednom)."""

    def __init__(self, vrsta):
        self.vrsta_dok = vrsta
        self.redovi = []
        self.klase = frozenset()
        # prijemnica
        self.status, self.z, self.zkand = "", None, []
        self.prev = self.next = None
        self.kand = None
        self.dodela = None
        self.naslednik = None
        self.greska = None
        self.cist = self.cist_deo = False
        # zbirna
        self.prj = []
        self.otp = []
        self.ziva = self.zombi = self.hl = False
        self.dodeljen = self.objasnjen = self.klasa_kand = None

    def dodaj(self, r, kl):
        self.redovi.append(r)
        self.klase = self.klase | kl

    def zavrsi(self, pk):
        r0 = self.redovi[0]
        self.aktivan = not storno(r0)
        self.ids = ", ".join(s(r, pk) for r in self.redovi)
        self.seq = min(_seq(s(r, pk)) for r in self.redovi)
        self.datum = d(r0, "Datum")
        self.vozac, self.vozac_txt = k(r0, "VozacID"), s(r0, "VozacID")
        self.kupac, self.kupac_txt = k(r0, "KupacID"), s(r0, "KupacID")
        self.vrsta = next((s(r, "VrstaVoca") for r in self.redovi if s(r, "VrstaVoca")), "")
        self.sorta = next((s(r, "SortaVoca") for r in self.redovi if s(r, "SortaVoca")), "")
        self.tip = next((s(r, "TipAmbalaze") for r in self.redovi if s(r, "TipAmbalaze")), "")
        if self.vrsta_dok == "prj":
            self.broj, self.bz = s(r0, "BrojPrijemnice"), s(r0, "BrojZbirne")
            self._kg, self._gaj = "Kolicina", "KolAmbalaze"
        else:
            self.broj = self.bz = s(r0, "BrojZbirne")
            self._kg, self._gaj = "UkupnoKolicina", "UkupnoAmbalaze"
        self.bzn = norm_broj(self.bz)
        self.kg = sum((n(r, self._kg) for r in self.redovi), Decimal(0))
        self.gajbe = sum((n(r, self._gaj) for r in self.redovi), Decimal(0))

    def _za(self, kol, klase):
        if not klase:
            return sum((n(r, kol) for r in self.redovi), Decimal(0))
        return sum((n(r, kol) for r in self.redovi if not _klase(r) or (_klase(r) & klase)), Decimal(0))

    def kg_za(self, klase):
        return self._za(self._kg, klase)

    def gajbe_za(self, klase):
        return self._za(self._gaj, klase)


def _grupisi(redovi, pk, kljuc_fn, vrsta):
    docs, otvoreni = [], defaultdict(list)
    for r in sorted(redovi, key=lambda r: (_seq(s(r, pk)), polje(r, "_redni") or 0)):
        kl, key, cilj = _klase(r), kljuc_fn(r), None
        if key is not None and kl:
            cilj = next((x for x in otvoreni[key] if x.klase and not (x.klase & kl)), None)
        if cilj is None:
            cilj = Dok(vrsta)
            docs.append(cilj)
            if key is not None:
                otvoreni[key].append(cilj)
        cilj.dodaj(r, kl)
    for x in docs:
        x.zavrsi(pk)
    return docs


def _kljuc_prj(r):
    if not k(r, "BrojPrijemnice"):
        return None
    return (k(r, "KupacID"), k(r, "BrojPrijemnice"), d(r, "Datum"), norm_broj(polje(r, "BrojZbirne")), storno(r))


def _kljuc_zbr(r):
    b = norm_broj(polje(r, "BrojZbirne"))
    return (b, k(r, "VozacID"), d(r, "Datum"), storno(r)) if b else None


class Kand:
    __slots__ = ("z", "m", "gd", "prolazi", "sused", "dd", "score")

    def __init__(self, z, m, gd, prolazi, sused, dd, score):
        self.z, self.m, self.gd, self.prolazi, self.sused, self.dd, self.score = z, m, gd, prolazi, sused, dd, score


class Greska:
    __slots__ = ("kod", "ozb", "opis", "nacin", "uslovno")

    def __init__(self, kod, ozb, opis, nacin="zbirna", uslovno=False):
        self.kod, self.ozb, self.opis, self.nacin, self.uslovno = kod, ozb, opis, nacin, uslovno


# ======================================================================================
class Analiza:
    def __init__(self, ctx):
        self.ctx = ctx
        self._spremno, self._vec = {}, set()
        c = ctx
        self.ima_voz = c.ima_kolonu(TPRJ, "VozacID") and c.ima_kolonu(TZBR, "VozacID")
        self.ima_kup = c.ima_kolonu(TPRJ, "KupacID") and c.ima_kolonu(TZBR, "KupacID")
        self.ima_kg = c.ima_kolonu(TPRJ, "Kolicina") and c.ima_kolonu(TZBR, "UkupnoKolicina")
        self.ima_gaj = c.ima_kolonu(TPRJ, "KolAmbalaze") and c.ima_kolonu(TZBR, "UkupnoAmbalaze")
        self.ima_ispr = c.ima_kolonu(TPRJ, "IspravkaOd")
        self.prag, self.prag_dana, self.kasn_med = PRAG_KG_PODR, PRAG_DANA_PODR, 0

        self.hl_stanice = set()
        if c.ima_kolonu(TSTA, "JeHladnjaca"):
            self.hl_stanice = {k(r, "StanicaID") for r in c.tab(TSTA) if k(r, "JeHladnjaca") in DA}

        self.prj_svi = _grupisi([r for r in c.tab(TPRJ) if not _je_test(r)], "PrijemnicaID", _kljuc_prj, "prj")
        self.zbr_svi = _grupisi([r for r in c.tab(TZBR) if not _je_test(r)], "ZbirnaID", _kljuc_zbr, "zbr")
        self.otp = [r for r in c.tab(TOTP) if not _je_test(r)]

        self.zbr_po_bzn = defaultdict(list)
        self.zbr_po_id = {}
        for z in self.zbr_svi:
            if z.bzn:
                self.zbr_po_bzn[z.bzn].append(z)
            for r in z.redovi:
                self.zbr_po_id.setdefault(k(r, "ZbirnaID"), z)
        self.otp_po_bzn = defaultdict(list)
        self.otp_po_broju = defaultdict(list)
        for r in self.otp:
            b = norm_broj(polje(r, "BrojZbirne"))
            if b:
                self.otp_po_bzn[b].append(r)
            if k(r, "BrojOtpremnice"):
                self.otp_po_broju[norm_broj(polje(r, "BrojOtpremnice"))].append(r)
        self._zbirne_zive()

        self.prj_akt = [p for p in self.prj_svi if p.aktivan]
        self.prj_akt_po_bzn = defaultdict(list)
        for p in self.prj_svi:
            p.hl = p.vozac in self.hl_stanice
            if p.aktivan and p.bzn:
                self.prj_akt_po_bzn[p.bzn].append(p)
        niz = sorted((p for p in self.prj_svi if not p.hl), key=lambda p: p.seq)
        for i, p in enumerate(niz):
            p.prev = niz[i - 1] if i else None
            p.next = niz[i + 1] if i + 1 < len(niz) else None

        self.pls_po_prj = defaultdict(list)
        for r in c.tab(TPLS):
            if not storno(r):
                self.pls_po_prj[k(r, "PrijemnicaID")].append(r)
        self.otp_vreme = sorted((_vreme(r), s(r, "VozacID"), s(r, "OtpremnicaID")) for r in self.otp if _vreme(r))
        self.W = []

    # --- infrastruktura -----------------------------------------------------------------
    def spremno(self, kod):
        if kod in self._spremno:
            return self._spremno[kod]
        fali = []
        ok = True
        for tb, kolone in ZAHTEVI.get(kod, {}).items():
            if not self.ctx.tab(tb):
                ok = False
                fali = []
                break
            fali += ["%s.%s" % (tb, x) for x in kolone if not self.ctx.ima_kolonu(tb, x)]
        if ok and fali:
            ok = False
            self.ctx.nalaz(kod + "_PRESKOCENO", INFO, fali[0].split(".")[0], "",
                           "Provera %s preskocena: nedostaje kolona %s" % (kod, ", ".join(fali)))
        self._spremno[kod] = ok
        return ok

    def nalaz(self, kod, ozb, tabela, id_, opis, predlog=""):
        if not self.spremno(kod) or (kod, tabela, id_) in self._vec:
            return False
        self._vec.add((kod, tabela, id_))
        self.ctx.nalaz(kod, ozb, tabela, id_, opis, predlog)
        return True

    def _isti_voz(self, a, b):
        return not self.ima_voz or a.vozac == b.vozac

    def _isti_kup(self, a, b):
        return not self.ima_kup or a.kupac == b.kupac

    def _isti_vlasnik(self, a, b):
        return self._isti_voz(a, b) and self._isti_kup(a, b)

    def _vk(self, x):
        return (x.vozac if self.ima_voz else "", x.kupac if self.ima_kup else "")

    def _slaze(self, pkg, pg, zkg, zg):
        if self.ima_gaj and pg > 0 and pg != zg:
            return False
        if self.ima_kg:
            m = _manjak(zkg, pkg)
            if m is None or abs(m) > self.prag:
                return False
        return True

    def odgovara(self, p, z):
        if not _kompat(p.klase, z.klase):
            return False
        return self._slaze(p.kg, p.gajbe, z.kg_za(p.klase), z.gajbe_za(p.klase))

    def _poredi(self, p, z):
        zkg, zg = z.kg_za(p.klase), z.gajbe_za(p.klase)
        return "kg %s vs zbirna %s, gajbe %s %s %s" % (_g(p.kg), _g(zkg), _g(p.gajbe),
                                                       "=" if p.gajbe == zg else "!=", _g(zg))

    def _zopis(self, z):
        return "%s (%s, %s, vozac %s, %s kg, %s gajbi, Klasa %s)" % (
            z.broj, z.ids, dan_txt(z.datum), z.vozac_txt or "?", _g(z.kg), _g(z.gajbe), _kl(z.klase))

    def _popis(self, p):
        return "%s %s (Datum %s, vozac %s, kupac %s, %s kg, %s gajbi, Klasa %s)" % (
            p.ids, p.broj or "bez broja", dan_txt(p.datum), p.vozac_txt or "?", p.kupac_txt or "?",
            _g(p.kg), _g(p.gajbe), _kl(p.klase))

    # --- priprema ----------------------------------------------------------------------
    def _zbirne_zive(self):
        for b, docs in self.zbr_po_bzn.items():
            otps = self.otp_po_bzn.get(b, [])
            akt = [o for o in otps if not storno(o)]
            aktivne = [z for z in docs if z.aktivan]
            uzi = {}
            if len(aktivne) > 1:
                for z in aktivne:
                    uzi[id(z)] = [o for o in akt if k(o, "VozacID") == z.vozac and d(o, "Datum") == z.datum]
            neko = any(uzi.values())
            for z in docs:
                z.otp = otps
                z.hl = z.vozac in self.hl_stanice or any(k(o, "StanicaID") in self.hl_stanice for o in otps)
                z.zombi = z.aktivan and bool(otps) and not akt
                z.ziva = z.aktivan and (bool(uzi.get(id(z))) if neko else bool(akt))

    def _razresi(self, p):
        if not p.bzn:
            return "PRAZAN", None, []
        sve = self.zbr_po_bzn.get(p.bzn, [])
        if not sve:
            return "NE_POSTOJI", None, []
        akt = [z for z in sve if z.aktivan]
        if not akt:
            return "SAMO_STORNIRANA", None, sve
        if len(akt) == 1:
            return "OK", akt[0], akt
        kand = akt
        for f in (lambda z: _kompat(p.klase, z.klase), lambda z: self._isti_voz(p, z),
                  lambda z: self._isti_kup(p, z), lambda z: z.kg > 0, lambda z: z.ziva,
                  lambda z: self.odgovara(p, z)):
            uzi = [z for z in kand if f(z)]
            if len(uzi) == 1:
                return "OK", uzi[0], akt
            if uzi:
                kand = uzi
        # Storno-ispravka otpremnice ne kaskadira na zbirnu: stara zbirna ostaje aktivna, a ponovni unos
        # istog broja pravi novu (veci ZbirnaID). Jedna aktivna + bar jedna stornirana otpremnica tog
        # broja, a zbirne iste po vozacu i datumu -> vazi najnovija.
        otps = self.otp_po_bzn.get(p.bzn, [])
        if len({(z.vozac, z.datum) for z in kand}) == 1 and any(storno(o) for o in otps) \
                and len([o for o in otps if not storno(o)]) == 1:
            return "OK", max(kand, key=lambda z: z.seq), akt
        return "DVOSMISLENA", None, kand

    def _pragovi(self):
        mm, dd = [], []
        for z in self.zbr_svi:
            if not z.aktivan or len(z.prj) != 1:
                continue
            p = z.prj[0]
            if not self._isti_vlasnik(p, z) or not _kompat(p.klase, z.klase):
                continue
            if p.prev is not None and p.prev.bzn == p.bzn:
                continue
            if self.ima_gaj and p.gajbe > 0 and p.gajbe != z.gajbe_za(p.klase):
                continue
            m = _manjak(z.kg_za(p.klase), p.kg)
            if m is not None and abs(m) < 50:
                mm.append(abs(m))
                if p.datum and z.datum:
                    dd.append((p.datum - z.datum).days)
        if len(mm) >= MIN_UZORAK:
            self.prag = min(max(_perc(mm, 0.99), PRAG_KG_MIN), PRAG_KG_MAX)
        if len(dd) >= MIN_UZORAK:
            self.prag_dana = min(max(_perc(dd, 0.99), PRAG_DANA_PODR), PRAG_DANA_MAX)
        if dd:
            self.kasn_med = max(0, _perc(dd, 0.5))

    def _kandidati(self, p, nacin="zbirna"):
        if p.kand is not None:
            return p.kand
        out = []
        for z in self.U_idx.get(self._vk(p), []):
            if z is p.z or not _kompat(p.klase, z.klase):
                continue
            dd = None
            if p.datum and z.datum:
                dd = (z.datum - p.datum).days
                if dd < -PROZOR_PRE or dd > PROZOR_POSLE:
                    continue
            if p.vrsta and z.vrsta and kljuc(p.vrsta) != kljuc(z.vrsta):
                continue
            if p.sorta and z.sorta and kljuc(p.sorta) != kljuc(z.sorta):
                continue
            zkg, zg = z.kg_za(p.klase), z.gajbe_za(p.klase)
            m = _manjak(zkg, p.kg)
            gd = abs(zg - p.gajbe) if (self.ima_gaj and p.gajbe > 0) else Decimal(0)
            prolazi = gd == 0 and (not self.ima_kg or (m is not None and abs(m) <= self.prag))
            sused = _susedan(p.z.broj if p.z is not None else p.bz, z.broj)
            ed = _lev(p.bzn, z.bzn) if nacin == "broj" else 0
            score = (0 if prolazi else 1, 0 if ed <= 1 else 1, gd, abs(m) if m is not None else Decimal(999),
                     0 if sused else 1, abs(dd) if dd is not None else 99, z.seq)
            out.append(Kand(z, m, gd, prolazi, sused, dd, score))
        out.sort(key=lambda x: x.score)
        p.kand = out
        return out

    def _prolazni(self, p, nacin="zbirna"):
        return [x for x in self._kandidati(p, nacin) if x.prolazi]

    def _greska(self, p, kod, ozb, opis, nacin="zbirna", uslovno=False):
        p.greska = Greska(kod, ozb, opis, nacin, uslovno)
        self.W.append(p)

    # --- glavni tok ---------------------------------------------------------------------
    def pokreni(self):
        self._testni()
        if not self.prj_svi:
            return
        for kod in ZAHTEVI:
            self.spremno(kod)
        for p in self.prj_akt:
            p.status, p.z, p.zkand = self._razresi(p)
            if p.z is not None:
                p.z.prj.append(p)
        self._pragovi()
        self.U = [z for z in self.zbr_svi
                  if z.aktivan and z.ziva and not z.hl and z.bzn and not self.prj_akt_po_bzn.get(z.bzn)]
        self.U_idx = defaultdict(list)
        for z in self.U:
            self.U_idx[self._vk(z)].append(z)
        self._klasifikuj()
        self._dodela()
        self._prijavi_veze()
        self._ciste()
        self._zive_bez_prijemnice()
        self._nenormalizovan()
        self._dupli()
        self._cena()
        self._gajbe()
        self._prevezana()
        self._ispravka()
        self._auto_lanac()

    # --- klasifikacija veze -------------------------------------------------------------
    def _klasifikuj(self):
        for z in self.zbr_svi:
            if not z.prj:
                continue
            docs = sorted(z.prj, key=lambda p: p.seq)
            isti = [p for p in docs if self._isti_vlasnik(p, z)]
            for p in docs:
                if p in isti:
                    continue
                if not self._isti_voz(p, z):
                    self._var_a(p, z, docs, isti)
                else:
                    self._var_kupac(p, z, isti)
            if z.hl:
                if len(isti) == 1:
                    isti[0].cist = True
                continue
            if len(isti) >= 2:
                self._var_b(z, isti)
            elif len(isti) == 1:
                self._jedna(isti[0], z)
        for p in self.prj_akt:
            if p.status != "OK":
                self._bez_zbirne(p)

    def _nasl(self, p):
        return p.prev is not None and p.prev.bzn == p.bzn

    def _var_a(self, p, z, docs, isti):
        pb = parse_broj(z.broj)
        p_st, z_st = self.ctx.je_stanica(p.vozac_txt), self.ctx.je_stanica(z.vozac_txt)
        prefix_ok = bool(pb and pb.s and p_st and pb.x == cifre(p.vozac_txt))
        z_pref_ok = bool(pb and pb.s and z_st and pb.x == cifre(z.vozac_txt))
        nasl = self._nasl(p)
        drugi_ok = [q for q in isti if self.odgovara(q, z)]
        kand = self._prolazni(p)
        pz = self.odgovara(p, z)
        mirror = p_st and z_st
        dop = []
        if self.ima_kup and p.kupac != z.kupac:
            dop.append("i KupacID se razlikuje (%s vs %s)" % (p.kupac_txt, z.kupac_txt))
        if prefix_ok and not nasl and not drugi_ok:
            self.nalaz("PRJ_ZASTAREO_VOZAC", SUMNJA, TPRJ, p.ids,
                       "Prijemnica %s: VozacID %s je stanica iz broja zbirne %s, ali zbirna nosi VozacID %s "
                       "(otpremnica sa vozacem druge stanice?). %s" % (p.ids, p.vozac_txt, z.broj, z.vozac_txt,
                                                                        self._poredi(p, z)),
                       "ODLUKA: VozacID prijemnice %s -> %s (kao zbirna, i u tblAmbalaza nogama prijemnice), ili "
                       "ispraviti VozacID otpremnice/zbirne %s" % (p.vozac_txt, z.vozac_txt, z.broj))
            return
        uvod = "Prijemnica %s (VozacID %s) stoji na zbirnoj %s (VozacID %s)" % (p.ids, p.vozac_txt, z.broj,
                                                                               z.vozac_txt)
        if nasl or drugi_ok or kand:
            razl = []
            if nasl:
                razl.append("broj zbirne NASLEDJEN od prethodne prijemnice %s" % p.prev.ids)
            if drugi_ok:
                razl.append("zbirni odgovara %s (VozacID %s)" % (", ".join(q.ids for q in drugi_ok),
                                                                 z.vozac_txt))
            if not pz:
                razl.append("prijemnica ne odgovara zbirnoj: %s" % self._poredi(p, z))
            ozb = GRESKA if (mirror and not prefix_ok) else SUMNJA
            self._greska(p, "PRJ_ZBIRNA_DRUGE_STANICE", ozb, uvod + "; " + "; ".join(razl + dop))
        elif len(docs) == 1 and pz:
            ozb = GRESKA if (mirror and z_pref_ok) else SUMNJA
            self.nalaz("PRJ_ZASTAREO_VOZAC", ozb, TPRJ, p.ids,
                       uvod + "; zbirna je jedina i odgovara (%s), za VozacID %s nema zive zbirne bez prijemnice"
                       "%s%s" % (self._poredi(p, z), p.vozac_txt, self._voz_pojacanje(p, z),
                                 "; " + "; ".join(dop) if dop else ""),
                       "VozacID: %s -> %s; isto u tblAmbalaza.VozacID za DokumentTip=Prijemnica, DokumentID=%s"
                       "%s" % (p.vozac_txt, z.vozac_txt, p.ids, self._amb_noge(p)))
        else:
            self._greska(p, "PRJ_ZBIRNA_DRUGE_STANICE", SUMNJA,
                         uvod + "; nejasno da li je pogresan broj zbirne ili VozacID (%s)%s" % (
                             self._poredi(p, z), "; " + "; ".join(dop) if dop else ""))

    def _voz_pojacanje(self, p, z):
        out = []
        if p.prev is not None and p.prev.vozac == p.vozac and p.prev.ids:
            out.append("VozacID isti kao prethodna prijemnica %s (zaostao cmbVozac)" % p.prev.ids)
        t = next((_vreme(r) for r in p.redovi if _vreme(r)), "")
        if t and self.otp_vreme:
            i = bisect_left(self.otp_vreme, (t,))
            if i > 0 and kljuc(self.otp_vreme[i - 1][1]) == p.vozac:
                out.append("VozacID kao poslednja otpremnica uneta pre prijemnice (%s)" % self.otp_vreme[i - 1][2])
        return ("; " + "; ".join(out)) if out else ""

    def _amb_noge(self, p):
        noge = [a for r in p.redovi for a in self.ctx.amb_po_dokumentu.get(("prijemnica", k(r, "PrijemnicaID")), [])
                if not storno(a)]
        return " (%d nogu)" % len(noge) if noge else ""

    def _var_kupac(self, p, z, isti):
        nasl = self._nasl(p) and self._isti_vlasnik(p.prev, z)
        drugi_ok = [q for q in isti if self.odgovara(q, z)]
        if (nasl or drugi_ok) and self._prolazni(p):
            self._greska(p, "PRJ_NASLEDJENA_ZBIRNA", SUMNJA,
                         "Prijemnica %s kupca %s stoji na zbirnoj %s kupca %s%s%s" % (
                             p.ids, p.kupac_txt, z.broj, z.kupac_txt,
                             "; broj zbirne nasledjen od %s" % p.prev.ids if nasl else "",
                             "; zbirni odgovara %s" % ", ".join(q.ids for q in drugi_ok) if drugi_ok else ""))
            return
        pb = parse_broj(p.broj)
        hlad = bool(self.ctx.malina_kupac) and z.kupac == kljuc(self.ctx.malina_kupac)
        if hlad and pb and not pb.s and pb.x == 1:
            pred = "KupacID: %s -> %s (zbirna je auto-zbirna kupca hladnjace, broj %s je u nizu hladnjace); " \
                   "ispraviti i EntitetID nogu prijemnice u tblAmbalaza" % (p.kupac_txt, z.kupac_txt, p.broj)
        else:
            pred = "ODLUKA: kupac prijemnice ili zbirne je pogresan (proveri i fakturu: Fakturisano/FakturaID)"
        self.nalaz("PRJ_KUPAC_RAZLICIT_OD_ZBIRNE", SUMNJA, TPRJ, p.ids,
                   "Prijemnica %s: KupacID %s, a zbirna %s ima KupacID %s (%s)" % (
                       p.ids, p.kupac_txt, z.broj, z.kupac_txt, self._poredi(p, z)), pred)

    def _var_b(self, z, isti):
        docs = isti
        odg = [self.odgovara(p, z) for p in docs]
        kl_ok = all(_kompat(p.klase, z.klase) for p in docs)
        S = sum((p.kg for p in docs), Decimal(0))
        G = sum((p.gajbe for p in docs), Decimal(0))
        m_sum = _manjak(z.kg, S)
        sum_ok = kl_ok and (not self.ima_kg or (m_sum is not None and abs(m_sum) <= self.prag)) and \
            (not self.ima_gaj or any(p.gajbe <= 0 for p in docs) or G == z.gajbe)
        if sum_ok and not any(self._prolazni(p) for p in docs):
            for p in docs:
                p.cist_deo = True
            razdvojene = all(p.klase for p in docs) and \
                sum(len(p.klase) for p in docs) == len(frozenset().union(*(p.klase for p in docs)))
            if not razdvojene:
                for p in docs[1:]:
                    self.nalaz("PRJ_DVE_NA_ZBIRNOJ_ISTA_STANICA", INFO, TPRJ, p.ids,
                               "Zbirna %s ima %d prijemnice istog vozaca i kupca (%s); zbir kg %s vs zbirna %s, "
                               "gajbe %s vs %s -- izgleda kao deljenje prijema" % (
                                   z.broj, len(docs), "; ".join(q.ids for q in docs), _g(S), _g(z.kg), _g(G),
                                   _g(z.gajbe)), "")
            return
        desni = [p for p, o in zip(docs, odg) if o]
        visak = m_sum is not None and m_sum < -self.prag
        if desni:
            ostaje = desni[0]
            ostaje.cist_deo = True
            for p in docs:
                if p is ostaje:
                    continue
                razl = ["zbirni %s odgovara prijemnica %s" % (z.broj, ostaje.ids)]
                if self._nasl(p):
                    razl.append("broj zbirne NASLEDJEN od prethodne prijemnice %s" % p.prev.ids)
                if not _kompat(p.klase, z.klase):
                    razl.append("KLASA_%s_NA_ZBIRNOJ_KLASE_%s" % (_kl(p.klase), _kl(z.klase)))
                elif self.odgovara(p, z):
                    razl.append("obe odgovaraju zbirnoj (iste ture); po redosledu unosa %s ostaje" % ostaje.ids)
                else:
                    razl.append("ova ne odgovara: %s" % self._poredi(p, z))
                if visak:
                    razl.append("zbir prijemnica %s kg > zbirna %s kg" % (_g(S), _g(z.kg)))
                self._greska(p, "PRJ_DVE_NA_ZBIRNOJ_ISTA_STANICA", GRESKA if visak else SUMNJA,
                             "Prijemnica %s je druga aktivna prijemnica na zbirnoj %s; %s" % (
                                 p.ids, self._zopis(z), "; ".join(razl)))
        else:
            for p in docs[1:]:
                self._greska(p, "PRJ_DVE_NA_ZBIRNOJ_ISTA_STANICA", SUMNJA,
                             "Zbirna %s ima %d aktivne prijemnice istog vozaca i kupca (%s); nijedna sama ne "
                             "odgovara, ni zbir (kg %s vs %s, gajbe %s vs %s) -- nejasno" % (
                                 self._zopis(z), len(docs), "; ".join(q.ids for q in docs), _g(S), _g(z.kg), _g(G),
                                 _g(z.gajbe)))

    def _jedna(self, p, z):
        p.cist = True
        pr = p.prev
        if z.zombi:
            self._greska(p, "PRJ_ZBIRNA_NE_POSTOJI_ILI_STORNIRANA", SUMNJA,
                         "Prijemnica %s stoji na zbirnoj %s cije su sve otpremnice stornirane (%s) -- "
                         "LANAC_STORNIRAN, zbirna je ostala aktivna" % (
                             p.ids, z.broj, ", ".join(s(o, "OtpremnicaID") for o in z.otp)), uslovno=True)
        elif pr is not None and pr.bzn == p.bzn and not pr.aktivan and not self.odgovara(p, z) \
                and self.odgovara(pr, z) and not (pr.kg == p.kg and pr.gajbe == p.gajbe) \
                and not self._ispravka_od(p, pr):
            self._greska(p, "PRJ_NASLEDJENA_ZBIRNA", SUMNJA,
                         "Prijemnica %s nosi broj zbirne %s prethodne (stornirane) prijemnice %s; zbirni je "
                         "odgovarala prethodna (%s), a ova ne (%s)" % (
                             p.ids, z.broj, pr.ids, self._poredi(pr, z), self._poredi(p, z)), uslovno=True)

    def _ispravka_od(self, p, q):
        if not self.ima_ispr or not q.broj:
            return False
        return any(norm_broj(polje(r, "IspravkaOd")) == norm_broj(q.broj) for r in p.redovi)

    def _bez_zbirne(self, p):
        if p.status == "PRAZAN":
            self._greska(p, "PRJ_ZBIRNA_NE_POSTOJI_ILI_STORNIRANA", SUMNJA,
                         "Prijemnica %s nema BrojZbirne" % self._popis(p), nacin="broj")
        elif p.status == "NE_POSTOJI":
            pb = parse_broj(p.bz)
            tip = "" if (pb and pb.dan) else " (format nije [S]x/ddmmyy[-n] -- TIPFELER)"
            self._greska(p, "PRJ_ZBIRNA_NE_POSTOJI_ILI_STORNIRANA", GRESKA,
                         "Prijemnica %s: BrojZbirne %s ne postoji u tblZbirna%s" % (self._popis(p), p.bz, tip),
                         nacin="broj")
        elif p.status == "SAMO_STORNIRANA":
            sve = self.zbr_po_bzn.get(p.bzn, [])
            otps = self.otp_po_bzn.get(p.bzn, [])
            p.naslednik = self._naslednik(sve, otps)
            if otps and all(storno(o) for o in otps):
                self._greska(p, "PRJ_ZBIRNA_NE_POSTOJI_ILI_STORNIRANA", GRESKA,
                             "Prijemnica %s je aktivna, a ceo lanac %s je storniran: zbirne %s, otpremnice %s "
                             "(LANAC_STORNIRAN, nije pogresan broj)" % (
                                 self._popis(p), p.bz, ", ".join(z.ids for z in sve),
                                 ", ".join(s(o, "OtpremnicaID") for o in otps)))
            else:
                self._greska(p, "PRJ_AKTIVNA_NA_STORNIRANOJ_ZBIRNOJ", GRESKA,
                             "Prijemnica %s je aktivna, a zbirna %s postoji samo stornirana (%s; VozacID %s, "
                             "KupacID %s)%s" % (
                                 self._popis(p), p.bz, ", ".join(z.ids for z in sve),
                                 ", ".join(sorted({z.vozac_txt for z in sve})),
                                 ", ".join(sorted({z.kupac_txt for z in sve})),
                                 "; aktivne otpremnice: %s" % ", ".join(s(o, "OtpremnicaID") for o in otps
                                                                         if not storno(o)) if otps else ""))
        elif p.status == "DVOSMISLENA":
            self.nalaz("PRJ_ZBIRNA_NE_POSTOJI_ILI_STORNIRANA", SUMNJA, TPRJ, p.ids,
                       "Prijemnica %s: BrojZbirne %s ima %d aktivne zbirne istog vozaca i kupca (DVOSMISLENA): %s" % (
                           self._popis(p), p.bz, len(p.zkand), "; ".join(self._zopis(z) for z in p.zkand)),
                       "ODLUKA: stornirati visak zbirne istog broja ili odluciti koja pripada prijemnici")

    def _akt_zbirna(self, v):
        if not tekst(v):
            return None
        akt = [z for z in self.zbr_po_bzn.get(norm_broj(v), []) if z.aktivan]
        if len(akt) == 1:
            return akt[0]
        z = self.zbr_po_id.get(kljuc(v))
        return z if z is not None and z.aktivan else None

    def _naslednik(self, sve, otps):
        sv = [r for r in self.ctx.tab(TSV)]
        for z in sve:
            for r in z.redovi:
                x = self._akt_zbirna(s(r, "ZamenjenSa"))
                if x is not None:
                    return x, "ZamenjenSa zbirne %s" % z.ids
        brojevi = {z.bzn for z in sve}
        for r in sv:
            if k(r, "OldDocType") == "zbirna" and norm_broj(polje(r, "OldBroj")) in brojevi:
                x = self._akt_zbirna(s(r, "NewBroj")) or self._akt_zbirna(s(r, "NewDocID"))
                if x is not None:
                    return x, "tblStornoVeze %s" % s(r, "CorrectionID")
        for o in otps:
            if not storno(o):
                continue
            novi = [s(o, "ZamenjenSa")] + [s(r, "NewBroj") for r in sv if k(r, "OldDocType") == "otpremnica"
                                            and norm_broj(polje(r, "OldBroj")) == norm_broj(polje(o, "BrojOtpremnice"))]
            for v in novi:
                for o2 in self.otp_po_broju.get(norm_broj(v), []) if v else []:
                    if not storno(o2):
                        x = self._akt_zbirna(s(o2, "BrojZbirne"))
                        if x is not None:
                            return x, "zamena otpremnice %s -> %s" % (s(o, "OtpremnicaID"), s(o2, "OtpremnicaID"))
        return None

    # --- 1:1 dodela ---------------------------------------------------------------------
    def _dodela(self):
        zauzete = set()
        for p in self.W:
            if p.naslednik is not None and id(p.naslednik[0]) not in zauzete:
                zauzete.add(id(p.naslednik[0]))
                p.naslednik[0].dodeljen = p
        parovi = []
        for p in self.W:
            if p.naslednik is None:
                for x in self._prolazni(p, p.greska.nacin):
                    parovi.append((x.score, p.seq, x.z.seq, id(p), p, x))
        parovi.sort(key=lambda t: t[:4])
        for _sc, _ps, _zs, _i, p, x in parovi:
            if p.dodela is None and id(x.z) not in zauzete:
                p.dodela = x
                x.z.dodeljen = p
                zauzete.add(id(x.z))

    def _predlog_veze(self, p, prol):
        cilj, delovi = None, []
        stari = p.bz or "''"
        if p.naslednik is not None:
            cilj, izvor = p.naslednik
            delovi.append("BrojZbirne: %s -> %s (naslednik: %s); %s" % (stari, cilj.broj, izvor, self._poredi(p, cilj)))
        elif p.dodela is not None:
            cilj = p.dodela.z
            ostali = [y for y in prol if y is not p.dodela][:2]
            if not ostali:
                delovi.append("BrojZbirne: %s -> %s; %s" % (stari, cilj.broj, self._poredi(p, cilj)))
            else:
                delovi.append("BrojZbirne: %s -> %s (1. kandidat; %s); ostali kandidati: %s; ODLUKA" % (
                    stari, cilj.broj, self._poredi(p, cilj),
                    "; ".join("%d) %s (%s)" % (i, y.z.broj, self._poredi(p, y.z)) for i, y in enumerate(ostali, 2))))
        elif prol:
            delovi.append("ODLUKA: kandidati su dodeljeni drugim prijemnicama: %s" % "; ".join(
                "%s (%s, uzela %s)" % (y.z.broj, self._poredi(p, y.z), y.z.dodeljen.ids if y.z.dodeljen else "?")
                for y in prol[:3]))
        else:
            najbl = self._kandidati(p, p.greska.nacin)[:1] if p.greska else []
            delovi.append("ODLUKA: nema zive zbirne bez prijemnice (isti VozacID i KupacID, Datum -%d..+%d) koja "
                          "odgovara po gajbama i kg%s" % (
                              PROZOR_PRE, PROZOR_POSLE,
                              "; najbliza: %s (%s)" % (najbl[0].z.broj, self._poredi(p, najbl[0].z)) if najbl else ""))
        if cilj is not None:
            if cilj.datum and p.datum and cilj.datum > p.datum:
                delovi.append("Datum prijemnice %s je pre datuma zbirne %s -- proveri i datum" % (
                    dan_txt(p.datum), dan_txt(cilj.datum)))
            pls = [s(r, "StavkaID") for x in p.redovi for r in self.pls_po_prj.get(k(x, "PrijemnicaID"), [])]
            if pls:
                delovi.append("uskladiti i tblPaletaStavka.BrojZbirne (%s)" % _skrati(pls))
        return "; ".join(delovi)

    def _prijavi_veze(self):
        for p in self.W:
            g = p.greska
            prol = self._prolazni(p, g.nacin)
            if g.uslovno and not prol and p.naslednik is None:
                continue
            opis = g.opis
            if g.kod == "PRJ_ZBIRNA_DRUGE_STANICE" and not self._nasl(p) and p.dodela is not None and p.z is not None:
                dz = p.dodela.z
                if abs(dz.seq - p.z.seq) <= 2 or dz.datum == p.z.datum:
                    opis += "; moguc KLIK_U_LISTI zbirnih (%s i %s su susedi)" % (p.z.ids, dz.ids)
            if self.nalaz(g.kod, g.ozb, TPRJ, p.ids, opis, self._predlog_veze(p, prol)) and g.uslovno:
                p.cist = False

    # --- cisti parovi: dve ture, kg/gajbe, klasa, datum, roba ---------------------------
    def _ciste(self):
        for p in self.prj_akt:
            z = p.z if (p.cist or p.cist_deo) else None
            if p.cist and z is not None:
                if not _kompat(p.klase, z.klase):
                    self._klasa(p, z)
                elif not self._dve_ture(p, z):
                    self._kg(p, z)
            self._datum(p, z)
            if z is not None:
                self._roba(p, z)

    def _klasa(self, p, z):
        pb = parse_broj(z.broj)
        kand = []
        for z2 in self.U_idx.get(self._vk(p), []):
            p2 = parse_broj(z2.broj)
            if _kompat(p.klase, z2.klase) and z2.klase and pb and p2 and (p2.s, p2.x, p2.dan) == (pb.s, pb.x, pb.dan):
                kand.append(z2)
        if kand:
            kand[0].klasa_kand = p
            pred = "BrojZbirne: %s -> %s (zbirna Klase %s); %s" % (p.bz, kand[0].broj, _kl(kand[0].klase),
                                                                   self._poredi(p, kand[0]))
        else:
            pred = "ODLUKA: Klasa prijemnice ili veza na zbirnu je pogresna"
        self.nalaz("PRJ_KLASA_RAZLICITA_OD_ZBIRNE", SUMNJA, TPRJ, p.ids,
                   "Prijemnica %s Klase %s stoji na zbirnoj %s Klase %s" % (p.ids, _kl(p.klase), z.broj, _kl(z.klase)),
                   pred)

    def _dve_ture(self, p, z1):
        if not self.ima_kg or not self.spremno("PRIJEMNICA_POKRIVA_DVE_TURE"):
            return False
        z1kg, z1g = z1.kg_za(p.klase), z1.gajbe_za(p.klase)
        m = _manjak(z1kg, p.kg)
        if m is None or m >= -self.prag:
            return False
        for z0 in self.U_idx.get(self._vk(p), []):
            if z0 is z1 or z0.dodeljen is not None or z0.objasnjen is not None:
                continue
            if z0.datum != z1.datum or not _kompat(p.klase, z0.klase):
                continue
            tot, gt = z1kg + z0.kg_za(p.klase), z1g + z0.gajbe_za(p.klase)
            m2 = _manjak(tot, p.kg)
            if m2 is None or abs(m2) > self.prag or not abs(p.kg - tot) < abs(p.kg - z1kg):
                continue
            if self.ima_gaj and p.gajbe > 0 and p.gajbe != gt:
                continue
            z0.objasnjen = p
            self.nalaz("PRIJEMNICA_POKRIVA_DVE_TURE", SUMNJA, TPRJ, p.ids,
                       "Prijemnica %s (%s kg, %s gajbi) na zbirnoj %s (%s kg, %s gajbi) pokriva i susednu turu %s "
                       "(%s kg, %s gajbi) bez prijemnice: zbir %s kg, %s gajbi" % (
                           p.ids, _g(p.kg), _g(p.gajbe), z1.broj, _g(z1kg), _g(z1g), z0.broj,
                           _g(z0.kg_za(p.klase)), _g(z0.gajbe_za(p.klase)), _g(tot), _g(gt)),
                       "ODLUKA: podeliti prijemnicu na %s i %s (ili povezati obe zbirne sa ovom prijemnicom)" % (
                           z1.broj, z0.broj))
            return True
        return False

    def _kg(self, p, z):
        zkg, zg = z.kg_za(p.klase), z.gajbe_za(p.klase)
        m = _manjak(zkg, p.kg) if self.ima_kg else None
        razl, tipf = [], None
        if m is not None and (m > self.prag or m < -VISAK_TOL):
            razl.append("%s kg vs zbirna %s kg (%s %s)" % (_g(p.kg), _g(zkg), "manjak" if m > 0 else "visak",
                                                         _pct(abs(m))))
            for f in (Decimal(10), Decimal("0.1")):
                m2 = _manjak(zkg, p.kg * f)
                if m2 is not None and abs(m2) <= self.prag:
                    tipf = p.kg * f
        gaj_ok = not (self.ima_gaj and p.gajbe > 0 and p.gajbe != zg)
        if not gaj_ok:
            razl.append("gajbe %s vs zbirna %s" % (_g(p.gajbe), _g(zg)))
        if not razl:
            return
        if tipf is not None:
            ozb, pred = (GRESKA if gaj_ok else SUMNJA), "Kolicina: %s -> %s (decimalni zarez/tipfeler)" % (
                _g(p.kg), _g(tipf))
        else:
            ozb, pred = SUMNJA, "ODLUKA: kalo/odbitak hladnjace ili greska unosa (prag kg %s)" % _pct(self.prag)
        self.nalaz("PRJ_KG_GAJBE_VAN_OPSEGA", ozb, TPRJ, p.ids,
                   "Prijemnica %s na zbirnoj %s: %s" % (p.ids, z.broj, "; ".join(razl)), pred)

    def _datum(self, p, z):
        if not self.spremno("PRJ_DATUM_ZASTAREO_ILI_PRE_ZBIRNE") or p.datum is None:
            return
        razl, jak = [], False
        if z is not None and z.datum:
            dd = (p.datum - z.datum).days
            if dd < 0:
                jak = True
                t = "Datum %s je %d dan(a) PRE datuma zbirne %s (%s)" % (dan_txt(p.datum), -dd, z.broj, dan_txt(z.datum))
                bz, bp = parse_broj(z.broj), parse_broj(p.broj)
                if bz and bz.dan and bz.dan > p.datum and bp and bp.dan == p.datum:
                    t += "; i datum u broju zbirne je posle prijema -- kandidat je greska otpremnice"
                razl.append(t)
            elif dd > self.prag_dana:
                razl.append("prijemnica kasni %d dana za zbirnom %s (%s; prag %d)" % (
                    dd, z.broj, dan_txt(z.datum), self.prag_dana))
        hlad = self.ctx.malina_kupac and p.kupac == kljuc(self.ctx.malina_kupac)
        pb = parse_broj(p.broj)
        if hlad and pb and not pb.s and pb.x == 1 and pb.dan and pb.dan != p.datum:
            jak = True
            razl.append("datum u broju %s (%s) != Datum %s (broj rucno kucan ili datum menjan)" % (
                p.broj, dan_txt(pb.dan), dan_txt(p.datum)))
        if not razl:
            return
        pr, nx = p.prev, p.next
        if jak and pr is not None and nx is not None and pr.datum and nx.datum and p.datum < pr.datum \
                and p.datum < nx.datum and (not pr.aktivan or any(s(r, "IspravkaOd") for r in pr.redovi)):
            razl.append("datum mozda nasledjen iz prefill-a ispravke (%s)" % pr.ids)
        kand = []
        if z is not None and z.datum:
            kand.append(dan_txt(z.datum + timedelta(days=int(self.kasn_med))))
        if pb and pb.dan and (z is None or not z.datum or pb.dan >= z.datum) and pb.dan != p.datum:
            kand.append(dan_txt(pb.dan) + " (iz broja)")
        self.nalaz("PRJ_DATUM_ZASTAREO_ILI_PRE_ZBIRNE", SUMNJA, TPRJ, p.ids,
                   "Prijemnica %s: %s" % (p.ids, "; ".join(razl)),
                   "ODLUKA: datum je poslovni dogadjaj%s" % ("; kandidat: " + " ili ".join(kand) if kand else ""))

    def _roba(self, p, z):
        razl, ozb, pred = [], SUMNJA, []
        for kol, pv, zv in (("VrstaVoca", p.vrsta, z.vrsta), ("SortaVoca", p.sorta, z.sorta),
                            ("TipAmbalaze", p.tip, z.tip)):
            if not zv:
                continue
            if pv and kljuc(pv) == kljuc(zv):
                continue
            if not pv and kol == "TipAmbalaze":
                continue   # prazan tip: PRJ_GAJBE_NULA_ILI_BEZ_TIPA
            razl.append("%s %s vs zbirna %s" % (kol, pv or "prazno", zv))
            if kol == "VrstaVoca" and pv:
                ozb = GRESKA
            if kol == "TipAmbalaze":
                pred.append("TipAmbalaze: ODLUKA (menja saldo po tipu; noge u tblAmbalaza nose isti tip)")
            else:
                pred.append("%s: %s -> %s (kao zbirna)" % (kol, pv or "''", zv))
        if razl:
            self.nalaz("PRJ_ROBA_RAZLICITA_OD_ZBIRNE", ozb, TPRJ, p.ids,
                       "Prijemnica %s na zbirnoj %s: %s" % (p.ids, z.broj, "; ".join(razl)), "; ".join(pred))

    # --- zbirne bez prijemnice ----------------------------------------------------------
    def _zive_bez_prijemnice(self):
        for z in self.U:
            otp = ", ".join(s(o, "OtpremnicaID") for o in z.otp if not storno(o))
            opis = "Ziva zbirna bez aktivne prijemnice: %s; otpremnice %s" % (self._zopis(z), otp)
            if z.objasnjen is not None:
                continue
            if z.dodeljen is not None:
                q = z.dodeljen
                self.nalaz("ZBR_ZIVA_BEZ_PRIJEMNICE", SUMNJA, TZBR, z.ids,
                           opis + "; verovatno prava zbirna za prijemnicu %s (sada na %s)" % (q.ids, q.bz or "''"),
                           "Prevezati prijemnicu %s: BrojZbirne %s -> %s" % (q.ids, q.bz or "''", z.broj))
            elif z.klasa_kand is not None:
                q = z.klasa_kand
                self.nalaz("ZBR_ZIVA_BEZ_PRIJEMNICE", INFO, TZBR, z.ids,
                           opis + "; BEZ_PRIJEMNICE_KLASE: prijemnica %s Klase %s stoji na %s" % (
                               q.ids, _kl(q.klase), q.bz), "Videti PRJ_KLASA_RAZLICITA_OD_ZBIRNE za %s" % q.ids)
            else:
                self.nalaz("ZBR_ZIVA_BEZ_PRIJEMNICE", SUMNJA, TZBR, z.ids, opis,
                           "ODLUKA: uneti prijemnicu ili utvrditi da roba nije primljena/odbijena")
        for b, prj in self.prj_akt_po_bzn.items():
            pk = frozenset().union(*(p.klase for p in prj))
            if not all(p.klase for p in prj):
                continue
            for z in self.zbr_po_bzn.get(b, []):
                if z.aktivan and z.ziva and not z.hl and z.klase and not z.klase <= pk:
                    self.nalaz("ZBR_ZIVA_BEZ_PRIJEMNICE", INFO, TZBR, z.ids,
                               "BEZ_PRIJEMNICE_KLASE: zbirna %s ima Klasu %s, prijemnice istog broja (%s) samo %s" % (
                                   z.broj, _kl(z.klase - pk), "; ".join(p.ids for p in prj), _kl(pk)),
                               "Proveriti da li hladnjaca vodi Klasu %s na istoj prijemnici" % _kl(z.klase - pk))

    # --- nezavisne provere --------------------------------------------------------------
    def _nenormalizovan(self):
        for p in self.prj_akt:
            razl, pred = [], []
            for kol in ("BrojZbirne", "BrojPrijemnice"):
                for r in p.redovi:
                    v = polje(r, kol)
                    if not isinstance(v, str) or not v.strip():
                        continue
                    if v != v.strip() or re.search(r"\s", v.strip()) or (kol == "BrojZbirne" and _kanon(v) != v.strip()):
                        cilj = _kanon(v)
                        if kol == "BrojZbirne" and p.z is not None and _kanon(p.z.broj) == p.z.broj:
                            cilj = p.z.broj
                        razl.append("%s '%s' (%s)" % (kol, v, s(r, "PrijemnicaID")))
                        pred.append("%s: '%s' -> '%s'" % (kol, v, cilj))
                        break
            # Prijemnica u kanonskom obliku a zbirna ne: to nije greska prijemnice (ne prijavljuje se ovde).
            if razl:
                if p.redovi and any(self.pls_po_prj.get(k(r, "PrijemnicaID")) for r in p.redovi):
                    pred.append("isto u tblPaletaStavka.BrojZbirne")
                self.nalaz("PRJ_BROJ_ZBIRNE_NENORMALIZOVAN", GRESKA, TPRJ, p.ids,
                           "Prijemnica %s: nenormalizovan broj (razmak ili mala slova; stari citaoci porede tacno): %s"
                           % (p.ids, "; ".join(razl)), "; ".join(pred))

    def _dupli(self):
        grupe = defaultdict(list)
        for p in self.prj_akt:
            if norm_broj(p.broj):
                grupe[norm_broj(p.broj)].append(p)
        hlad = kljuc(self.ctx.malina_kupac) if self.ctx.malina_kupac else None
        sledeci = {}
        for b, g in grupe.items():
            if len(g) < 2:
                continue
            g = sorted(g, key=lambda p: p.seq)
            prvi = g[0]
            for p in g[1:]:
                razl_tekst = " (isti tek posle normalizacije: '%s' vs '%s')" % (p.broj, prvi.broj) \
                    if p.broj != prvi.broj else ""
                if p.kupac != prvi.kupac and p.bzn and p.bzn == prvi.bzn:
                    self.nalaz("PRJ_PREVEZANA_BEZ_PROVERE", SUMNJA, TPRJ, p.ids,
                               "PREVEZANO_ZAJEDNO: prijemnice %s (kupac %s) i %s (kupac %s) imaju isti broj %s i "
                               "istu zbirnu %s (prevezivanje po broju je globalno, preko kupaca)" % (
                                   prvi.ids, prvi.kupac_txt, p.ids, p.kupac_txt, p.broj, p.bz),
                               "ODLUKA: vratiti zbirnu prijemnici koja ne pripada %s" % p.bz)
                    continue
                isti_k = p.kupac == prvi.kupac
                ozb = GRESKA if (isti_k and hlad and p.kupac == hlad) else SUMNJA
                pred = "ODLUKA: broj je kupcev (eksterni kupac)"
                if isti_k and hlad and p.kupac == hlad and p.datum:
                    kl = (p.kupac, p.datum)
                    if kl not in sledeci:
                        mx = 0
                        for q in self.prj_svi:
                            qb = parse_broj(q.broj)
                            if q.kupac == p.kupac and qb and not qb.s and qb.x == 1 and qb.dan == p.datum:
                                mx = max(mx, qb.seq)
                        sledeci[kl] = mx
                    sledeci[kl] += 1
                    nb = "1/%s" % p.datum.strftime("%d%m%y") + ("-%d" % sledeci[kl] if sledeci[kl] > 1 else "")
                    pred = "BrojPrijemnice: %s -> %s (sledeci slobodan za kupca i dan, racunajuci i stornirane)" % (
                        p.broj, nb)
                self.nalaz("PRJ_DUPLI_BROJ_PRIJEMNICE", ozb, TPRJ, p.ids,
                           "Dve aktivne prijemnice sa brojem %s: %s i %s%s" % (
                               p.broj, self._popis(prvi), self._popis(p), razl_tekst), pred)

    def _cena(self):
        if not self.spremno("PRJ_CENA_NULA"):
            return
        nule = []
        for p in self.prj_akt:
            rr = [r for r in p.redovi if not ima_broj(r, "Cena") or n(r, "Cena") == 0]
            if rr:
                nule.append((p, rr))
        if len(nule) >= CENA_SISTEMSKI and len(nule) * 2 > len(self.prj_akt):
            self.nalaz("PRJ_CENA_NULA", INFO, TPRJ, _skrati([p.ids for p, _ in nule], 5),
                       "%d od %d aktivnih prijemnica nema cenu -- izgleda kao praksa (cena na fakturi), ne greska unosa"
                       % (len(nule), len(self.prj_akt)), "ODLUKA: potvrditi da se cena odredjuje na fakturi")
            return
        for p, rr in nule:
            fks = [f for r in rr for f in self.ctx.fks_po_prijemnici.get(k(r, "PrijemnicaID"), [])
                   if not storno(f) and n(f, "Cena") > 0]
            pre = p.datum is not None and p.datum < DATUM_ZASTITE
            pred = "Cena: 0 -> %s (iz tblFakturaStavke %s)" % (_g(n(fks[0], "Cena")), s(fks[0], "StavkaID")) if fks \
                else "ODLUKA: uneti cenu (ugovor/faktura) za Datum %s" % dan_txt(p.datum)
            self.nalaz("PRJ_CENA_NULA", SUMNJA, TPRJ, p.ids,
                       "Prijemnica %s bez cene (Cena prazno ili 0; Klasa %s)%s" % (
                           self._popis(p), _kl(frozenset().union(*(_klase(r) for r in rr))),
                           "; pre 2026-07-01 forma nije trazila cenu" if pre else ""), pred)

    def _gajbe(self):
        if not self.spremno("PRJ_GAJBE_NULA_ILI_BEZ_TIPA"):
            return
        for p in self.prj_akt:
            razl, pred = [], []
            z = p.z if p.status == "OK" else None
            if p.gajbe <= 0 and z is not None and self.ima_gaj and z.gajbe_za(p.klase) > 0:
                razl.append("KolAmbalaze prazno/0, a zbirna %s ima %s gajbi" % (z.broj, _g(z.gajbe_za(p.klase))))
                pred.append("KolAmbalaze: 0 -> %s (zbirna)" % _g(z.gajbe_za(p.klase)))
            bez_tipa = [s(r, "PrijemnicaID") for r in p.redovi if n(r, "KolAmbalaze") > 0 and not s(r, "TipAmbalaze")]
            if bez_tipa:
                razl.append("KolAmbalaze > 0 bez TipAmbalaze (%s)" % ", ".join(bez_tipa))
                pred.append("TipAmbalaze: '' -> %s" % (z.tip if z is not None and z.tip else "ODLUKA"))
            if not razl:
                continue
            noge = sum(len([a for a in self.ctx.amb_po_dokumentu.get(("prijemnica", k(r, "PrijemnicaID")), [])
                            if not storno(a)]) for r in p.redovi)
            if p.datum is not None and p.datum < DATUM_ZASTITE:
                razl.append("pre 2026-07-01 gajbe nisu bile obavezne")
            razl.append("nogu u tblAmbalaza: %d" % noge)
            self.nalaz("PRJ_GAJBE_NULA_ILI_BEZ_TIPA", SUMNJA, TPRJ, p.ids,
                       "Prijemnica %s: %s" % (p.ids, "; ".join(razl)), "; ".join(pred) + "; proveriti noge ambalaze")

    def _prevezana(self):
        if not self.spremno("PRJ_PREVEZANA_BEZ_PROVERE"):
            return
        for p in self.prj_akt:
            pls = [r for x in p.redovi for r in self.pls_po_prj.get(k(x, "PrijemnicaID"), [])]
            razl = sorted({s(r, "BrojZbirne") for r in pls if norm_broj(polje(r, "BrojZbirne"))
                           and norm_broj(polje(r, "BrojZbirne")) != p.bzn})
            if not razl:
                continue
            ids_pls = [s(r, "StavkaID") for r in pls if s(r, "BrojZbirne") in razl]
            cur_ok = p.z is not None and self._isti_vlasnik(p, p.z) and self.odgovara(p, p.z)
            alt = self._akt_zbirna(razl[0]) if len(razl) == 1 else None
            alt_ok = alt is not None and self._isti_vlasnik(p, alt) and self.odgovara(p, alt)
            if cur_ok and not alt_ok:
                pred = "Ostaviti %s; uskladiti tblPaletaStavka.BrojZbirne %s -> %s (%s)" % (
                    p.bz, ", ".join(razl), p.bz, _skrati(ids_pls))
            elif alt_ok and not cur_ok:
                pred = "BrojZbirne: %s -> %s (prema paletnoj stavci; %s)" % (p.bz, alt.broj, self._poredi(p, alt))
            else:
                pred = "ODLUKA: prikazane obe vrednosti; izabrati onu koja prolazi kg/gajbe/vozac prema zbirnoj"
            self.nalaz("PRJ_PREVEZANA_BEZ_PROVERE", SUMNJA, TPRJ, p.ids,
                       "Prijemnica %s ima BrojZbirne %s, a njene paletne stavke (%s) nose %s -- veza menjana posle "
                       "paletizacije (Prevezi bez provere vlasnika i 1 zbirna = 1 prijemnica)" % (
                           p.ids, p.bz or "''", _skrati(ids_pls), ", ".join(razl)), pred)

    def _ispravka(self):
        if not self.ima_ispr or not self.spremno("PRJ_ISPRAVKA_NA_DRUGOJ_ZBIRNOJ"):
            return
        stor = defaultdict(list)
        for q in self.prj_svi:
            if not q.aktivan and norm_broj(q.broj):
                stor[norm_broj(q.broj)].append(q)
        for p in self.prj_akt:
            od = next((s(r, "IspravkaOd") for r in p.redovi if s(r, "IspravkaOd")), "")
            if not od:
                continue
            kand = stor.get(norm_broj(od), [])
            if not kand:
                continue
            S = sorted(kand, key=lambda q: (q.kupac != p.kupac, q.seq > p.seq, -q.seq))[0]
            razl = []
            if S.bzn != p.bzn:
                p_ok = p.z is not None and self._isti_vlasnik(p, p.z) and self.odgovara(p, p.z)
                zs = [z for z in self.zbr_po_bzn.get(S.bzn, [])]
                s_ok = any(self.odgovara(S, z) for z in zs)
                if not (p_ok and not s_ok):
                    t = "ispravka %s premestila prijemnicu sa zbirne %s na %s" % (S.ids, S.bz or "''", p.bz or "''")
                    t += "; nova %s svojoj zbirnoj" % ("odgovara" if p_ok else "NE odgovara")
                    t += "; stara je %s svojoj" % ("odgovarala" if s_ok else "NE odgovarala")
                    pls = [r for x in p.redovi for r in self.pls_po_prj.get(k(x, "PrijemnicaID"), [])]
                    if pls and not p_ok:
                        pg = sum((n(r, "BrojGajbica") for r in pls), Decimal(0))
                        pk = sum((n(r, "NetoKg") for r in pls), Decimal(0))
                        if pg == S.gajbe and (not S.kg or abs(_manjak(S.kg, pk) or 0) <= self.prag) and \
                                not (pg == p.gajbe and p.kg == pk):
                            t += "; POGRESNO_VEZANA: palete (%s gajbi, %s kg) odgovaraju staroj, ne novoj" % (
                                _g(pg), _g(pk))
                    if p.z is not None and len([q for q in p.z.prj if self._isti_vlasnik(q, p.z)]) >= 2:
                        t += "; zbirna %s sada ima vise prijemnica (1 zbirna = 1 prijemnica preskocena)" % p.z.broj
                    razl.append(t)
            dif = [x for x, b in (("kupac", S.kupac != p.kupac), ("vozac", S.vozac != p.vozac),
                                  ("datum", bool(S.datum and p.datum and (p.datum - S.datum).days > 2)),
                                  ("zbirna", S.bzn != p.bzn)) if b]
            if len(dif) >= 2:
                razl.append("POGRESNA_ZAMENA: razlikuju se %s (crash-safe pitanje 'ZAMENA?' ima Da kao "
                            "podrazumevano)" % ", ".join(dif))
            if razl:
                self.nalaz("PRJ_ISPRAVKA_NA_DRUGOJ_ZBIRNOJ", SUMNJA, TPRJ, p.ids,
                           "Prijemnica %s (IspravkaOd %s): %s" % (self._popis(p), od, "; ".join(razl)),
                           "ODLUKA: prikazati staru %s (%s), novu %s (%s) i paletne stavke" % (
                               S.ids, S.bz or "''", p.ids, p.bz or "''"))

    def _auto_lanac(self):
        if not self.hl_stanice or not self.spremno("PRJ_AUTO_LANAC_NEPOTPUN"):
            return
        prj_kl = defaultdict(list)
        for p in self.prj_akt:
            for r in p.redovi:
                prj_kl[(p.bzn, _klase(r))].append(p)
        po_bzn = defaultdict(list)
        for o in self.otp:
            if not storno(o) and k(o, "StanicaID") in self.hl_stanice and norm_broj(polje(o, "BrojZbirne")):
                po_bzn[norm_broj(polje(o, "BrojZbirne"))].append(o)
        for b, otps in po_bzn.items():
            fali, dupl = [], []
            for o in otps:
                pp = {id(x): x for x in prj_kl.get((b, _klase(o)), [])}
                if not pp:
                    fali.append(o)
                elif len(pp) > 1:
                    dupl.append((o, list(pp.values())))
            if fali:
                o = fali[0]
                self.nalaz("PRJ_AUTO_LANAC_NEPOTPUN", SUMNJA, TOTP, ", ".join(s(x, "OtpremnicaID") for x in fali),
                           "Lanac hladnjace (stanica %s): otpremnica %s i zbirna %s, a nema aktivne prijemnice "
                           "Klase %s" % (s(o, "StanicaID"), s(o, "BrojOtpremnice"), s(o, "BrojZbirne"),
                                         _kl(frozenset().union(*(_klase(x) for x in fali)))),
                           "ODLUKA: prijemnica iz otpremnice (kg, gajbe, cena, bruto po klasi), broj = sledeci "
                           "1/ddmmyy-n kupca hladnjace za %s (kao BackfillPrijemniceHladnjaca)" % dan_txt(d(o, "Datum")))
            for o, pp in dupl[:1]:
                self.nalaz("PRJ_AUTO_LANAC_NEPOTPUN", SUMNJA, TOTP, s(o, "OtpremnicaID"),
                           "Lanac hladnjace: zbirna %s Klase %s ima %d aktivne prijemnice (%s) -- backfill duplikat?" % (
                               s(o, "BrojZbirne"), _kl(_klase(o)), len(pp), "; ".join(x.ids for x in pp)),
                           "ODLUKA: stornirati visak prijemnice")

    def _testni(self):
        for tb, pk in TEST_TABELE.items():
            rr = [r for r in self.ctx.tab(tb) if _je_test(r)]
            if rr:
                self.nalaz("TESTNI_REDOVI_U_PRODUKCIJI", SUMNJA, tb, _skrati([s(r, pk) for r in rr], 5),
                           "%d redova testne suite (TST-PRO / KUP-90001 / VOZ-90001 / ST-90001 / 'Test Jabuka') u %s; "
                           "izuzeti su iz ostalih provera" % (len(rr), tb),
                           "ODLUKA: stornirati/ukloniti testne redove pre migracije")


def provere(ctx):
    Analiza(ctx).pokreni()
