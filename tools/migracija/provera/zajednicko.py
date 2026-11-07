"""Zajednicki deo provere unosa stare sveske -- SAMO CITA.

Stara sveska (C002, vba-v2.2.3 .. vba-v2.28.4) nema tabele clanstva ni
GeneracijaID: veze su kolone na samim dokumentima (vidi tools/migracija/
provera/README.md, odeljak "Veze"). Ovaj modul daje:

  * Kontekst -- tabele po imenu, redovi kao dict (kolone bez obzira na velika
    slova), indeksi po ID-ju i po broju, mapa broj stanice <-> StanicaID;
  * citanje vrednosti: s() tekst, n() broj (Decimal), d() datum, storno();
  * parse_broj() -- "x/ddmmyy[-n]" i "Sx/ddmmyy[-n]" (modBrojevi);
  * nalaz() -- jedan nalaz provere (kod, ozbiljnost, tabela, ID, opis, predlog).

Moduli provera (prijemnica.py, otpremnica_zbirna.py, ...) imaju po jednu
funkciju `provere(ctx)` koja samo cita ctx i zove ctx.nalaz(...).
"""

import re
from collections import defaultdict
from datetime import date
from decimal import Decimal

from citac import broj as _broj
from citac import datum as _datum
from citac import je_storno, kljuc, polje, tabela, tekst

GRESKA = "GRESKA"      # podatak je sigurno pogresan (dokazano iz podataka)
SUMNJA = "SUMNJA"      # potpis greske, ali postoji legitimno objasnjenje -- operater proverava
INFO = "INFO"          # skrece paznju, nije greska sama po sebi
OZBILJNOSTI = (GRESKA, SUMNJA, INFO)

_BROJ = re.compile(r"^\s*(S?)\s*(\d+)\s*/\s*(\d{2})(\d{2})(\d{2})\s*(?:-\s*(\d+))?\s*$", re.IGNORECASE)


class Broj:
    """Rastavljen broj dokumenta "x/ddmmyy[-n]" (opciono sa vodecim S)."""

    __slots__ = ("s", "x", "dan", "seq", "izvor")

    def __init__(self, s, x, dan, seq, izvor):
        self.s, self.x, self.dan, self.seq, self.izvor = s, x, dan, seq, izvor

    def __repr__(self):
        return "Broj(%s%d/%s-%d)" % ("S" if self.s else "", self.x, self.dan, self.seq)


def parse_broj(v):
    """Broj ili None. dan = date iz ddmmyy (None ako datum ne postoji), seq = 1
    kad nema nastavka (ExtractSeqFromBroj)."""
    t = tekst(v)
    m = _BROJ.match(t)
    if not m:
        return None
    s, x, dd, mm, yy, seq = m.groups()
    try:
        dan = date(2000 + int(yy), int(mm), int(dd))
    except ValueError:
        dan = None
    return Broj(bool(s), int(x), dan, int(seq) if seq else 1, t)


def norm_broj(v):
    """Kljuc broja kakav bi trebalo da bude: bez ikakvih razmaka, bez obzira na
    velika slova. kljuc() skida samo spoljne razmake -- razlika izmedju njih je
    potpis nenormalizovanog unosa."""
    return re.sub(r"\s+", "", tekst(v)).casefold()


def cifre(v):
    """Broj iz ID-ja entiteta kao ExtractNumericFromEntityID: spoji SVE cifre."""
    c = "".join(ch for ch in tekst(v) if ch.isdigit())
    return int(c) if c else None


def s(r, kol):
    return tekst(polje(r, kol))


def k(r, kol):
    return kljuc(polje(r, kol))


def n(r, kol):
    v = _broj(polje(r, kol))
    return v if v is not None else Decimal(0)


def ima_broj(r, kol):
    return _broj(polje(r, kol)) is not None


def d(r, kol):
    return _datum(polje(r, kol))


def storno(r):
    return je_storno(r)


def dan_txt(v):
    x = v if isinstance(v, date) else _datum(v)
    return x.isoformat() if x else tekst(v)


def kg_txt(v):
    if v is None:
        return ""
    v = Decimal(v)
    return format(round(v, 2).normalize(), "f") if v else "0"


class Nalaz:
    __slots__ = ("kod", "ozbiljnost", "tabela", "id", "opis", "predlog", "modul")

    def __init__(self, kod, ozbiljnost, tabela, id_, opis, predlog, modul):
        if ozbiljnost not in OZBILJNOSTI:
            raise ValueError("ozbiljnost %r" % ozbiljnost)
        self.kod, self.ozbiljnost, self.tabela, self.id = kod, ozbiljnost, tabela, id_
        self.opis, self.predlog, self.modul = opis, predlog, modul

    def red(self):
        return [self.ozbiljnost, self.kod, self.tabela, self.id, self.opis, self.predlog, self.modul]


KOLONE_NALAZA = ["Ozbiljnost", "Kod", "Tabela", "ID", "Opis", "Predlog", "Modul"]


class Kontekst:
    """Sve sto provere citaju. Napravi se jednom po svesci."""

    TABELE = ("tblOtkup", "tblOtpremnica", "tblZbirna", "tblPrijemnica", "tblNovac", "tblAmbalaza",
              "tblPaleta", "tblPaletaStavka", "tblFakture", "tblFakturaStavke", "tblStanice",
              "tblVozaci", "tblKupci", "tblKooperanti", "tblStornoVeze", "tblBankaImport",
              "tblSEFConfig", "tblConfig")
    PK = {"tblOtkup": "OtkupID", "tblOtpremnica": "OtpremnicaID", "tblZbirna": "ZbirnaID",
          "tblPrijemnica": "PrijemnicaID", "tblNovac": "NovacID", "tblAmbalaza": "AmbID",
          "tblPaleta": "PaletaID", "tblPaletaStavka": "StavkaID", "tblFakture": "FakturaID",
          "tblFakturaStavke": "StavkaID", "tblStanice": "StanicaID", "tblVozaci": "VozacID",
          "tblKupci": "KupacID", "tblKooperanti": "KooperantID"}

    def __init__(self, tabele):
        self._tabele = tabele
        self.nalazi = []
        self._modul = ""
        self.t = {}
        self.kolone = {}
        for ime in self.TABELE:
            tb = tabela(tabele, ime)
            self.t[ime] = list(tb.redovi) if tb is not None else []
            self.kolone[ime] = [c.casefold() for c in tb.kolone] if tb is not None else None
        # indeksi po ID-ju (kljuc() oblik)
        self.po_id = {}
        for ime, pk in self.PK.items():
            idx = {}
            for r in self.t[ime]:
                idx.setdefault(k(r, pk), r)
            self.po_id[ime] = idx
        # stanice: broj -> StanicaID (skeniranjem, kao sto savetuje modBrojevi)
        self.stanica_po_broju = {}
        for r in self.t["tblStanice"]:
            x = cifre(s(r, "StanicaID"))
            if x is not None:
                self.stanica_po_broju.setdefault(x, s(r, "StanicaID"))
        self.stanice = {k(r, "StanicaID") for r in self.t["tblStanice"]}
        # brojevi -> redovi
        self.otp_po_broju = _grupe(self.t["tblOtpremnica"], "BrojOtpremnice")
        self.zbr_po_broju = _grupe(self.t["tblZbirna"], "BrojZbirne")
        self.prj_po_broju_zbirne = _grupe(self.t["tblPrijemnica"], "BrojZbirne")
        self.otp_po_broju_zbirne = _grupe(self.t["tblOtpremnica"], "BrojZbirne")
        self.otk_po_otpremnici = _grupe(self.t["tblOtkup"], "OtpremnicaID")
        self.nov_po_otkupu = _grupe(self.t["tblNovac"], "OtkupID")
        self.amb_po_dokumentu = defaultdict(list)
        for r in self.t["tblAmbalaza"]:
            self.amb_po_dokumentu[(k(r, "DokumentTip"), k(r, "DokumentID"))].append(r)
        self.pls_po_prijemnici = _grupe(self.t["tblPaletaStavka"], "PrijemnicaID")
        self.fks_po_prijemnici = _grupe(self.t["tblFakturaStavke"], "PrijemnicaID")
        self.malina_kupac = self.config("MALINA_DEFAULT_KUPAC")

    # --- citanje -------------------------------------------------------
    def tab(self, ime):
        """Redovi tabele (prazna lista ako tabele nema)."""
        return self.t.get(ime) if ime in self.t else _redovi(self._tabele, ime)

    def ima_tabelu(self, ime):
        return self.kolone.get(ime) is not None or tabela(self._tabele, ime) is not None

    def ima_kolonu(self, ime, kol):
        kolone = self.kolone.get(ime)
        if kolone is None:
            tb = tabela(self._tabele, ime)
            kolone = [c.casefold() for c in tb.kolone] if tb is not None else []
        return kol.casefold() in kolone

    def red(self, ime, id_):
        return self.po_id.get(ime, {}).get(kljuc(id_))

    def config(self, kljuc_cfg):
        for ime, kc, vc in (("tblSEFConfig", "ConfigKey", "ConfigValue"), ("tblConfig", "Kljuc", "Vrednost")):
            for r in self.t.get(ime, []):
                if k(r, kc) == kljuc_cfg.casefold():
                    return s(r, vc)
        return ""

    def stanica_iz_broja(self, b):
        """StanicaID za broj dokumenta (x iz "x/..." ili "Sx/..."), '' ako nema."""
        p = b if isinstance(b, Broj) else parse_broj(b)
        return self.stanica_po_broju.get(p.x, "") if p else ""

    def je_stanica(self, id_):
        return kljuc(id_) in self.stanice

    # --- nalazi ------------------------------------------------------
    def modul(self, ime):
        self._modul = ime

    def nalaz(self, kod, ozbiljnost, tabela_, id_, opis, predlog=""):
        self.nalazi.append(Nalaz(kod, ozbiljnost, tabela_, tekst(id_), opis, predlog, self._modul))


def _grupe(redovi, kol):
    g = defaultdict(list)
    for r in redovi:
        v = k(r, kol)
        if v:
            g[v].append(r)
    return g


def _redovi(tabele, ime):
    tb = tabela(tabele, ime)
    return list(tb.redovi) if tb is not None else []


def aktivni(redovi):
    return [r for r in redovi if not storno(r)]


def ids(redovi, pk):
    return ", ".join(s(r, pk) for r in redovi)
