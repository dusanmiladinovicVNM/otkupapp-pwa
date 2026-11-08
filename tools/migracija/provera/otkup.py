"""Provere OTKUPA stare sveske (C002, vba-v2.2.3 .. vba-v2.28.4) -- SAMO CITA.

Otkup se unosio kroz frmOtkup (od v2.7.0 uz panel blokova modOtkupBlok) i kroz PWA
uvoz (ImportRowToTblOtkup). Veza otkup -> otpremnica je tblOtkup.OtpremnicaID (po
ID-ju); BrojZbirne, VozacID i BrojOtpremnice na otkupu su kopije. Polja forme koja
se posle snimanja ne prazne (txtDatum, txtBrojZbirne, cmbVozac, cena, tip ambalaze,
do v2.7.0 i kooperant) ostavljaju potpis u podacima, a sync
(AutoCreateOtpremniceFromPWA, modMasterSync.bas:628-771) pakuje SVAKI nevezan otkup
sa VozacID-jem u novu otpremnicu, i desktop red.

Klase (kod nalaza = ID klase iz specifikacije):
  OTK_ZASTARELI_KONTEKST_DATUM     posle auto-deselekcije blok nosi datum/zbirnu stare otpremnice
  OTK_VEZA_NESAGLASNA_SA_OTP       otkup vezan za otpremnicu druge stanice, dana, vrste/sorte, klase
  OTK_PANEL_PREGAZIO_VEZU          panel pregazio vezu lanca hladnjace (A), blok na storniranoj (B)
  DUPLI_LANAC_HLADNJACE_POSLE_GUARD2  v2.16.1+: lanac hladnjace + panelom izabrana otpremnica (C)
  OTK_BROJ_ZBIRNE_ZAOSTAO          BrojZbirne otkupa != zbirna njegove otpremnice / nijedne
  OTK_VOZAC_ZAOSTAO_ILI_LOSE_IZVUCEN  VozacID nepostojeci, druge stanice, razlicit od otpremnice
  VOZAC_HLADNJACA_LITERAL          VozacID 'hladnjaca' ili bez cifara (ExtractIDFromDisplay)
  OTK_KOOPERANT_ZAOSTAO            do v2.7.0: uzastopni blokovi istog kooperanta (kandidati)
  OTK_CENA_TIP_ZAOSTALI            najnovija umesto vazece cene, cena storniranog prvog bloka, tip
  ZAOSTALA_CENA_POSLE_DESELEKCIJE  cena popunjene otpremnice presla na blok drugog proizvoda
  OTK_BROJ_VAN_NIZA_ILI_PRAZAN     broj prazan, nekanonski, tudja stanica/dan, duplikat
  PWA_STANICA_I_DATUM              PWA: stanica kooperanta / datum zamenjen danom uvoza
  OTK_DUPLI_UNOS_BLOKA             isti blok unet dvaput (samo uz jak dokaz)
  OTK_KULTURAID_PO_VRSTI           KulturaID razresen samo po vrsti (prva sorta)
  KOOP_DUPLIKAT_AUTOKREIRANJEM     kooperant auto-kreiran po imenu / otkup pripisan imenjaku
  OTK_KOOPERANT_NEPOSTOJECI_ID     KooperantID otkupa nije u tblKooperanti
  OTK_GAJBE_KG_NEPROVERENO         bez gajbi uz manjak gajbi otpremnice, kg po gajbi van opsega
  OTK_AKTIVAN_BEZ_ZIVE_OTPREMNICE  aktivan otkup na storniranoj/nepostojecoj otpremnici, oslobodjen stornom
  SYNC_AUTO_OTPREMNICA_ZA_NEVEZANE sync otpremnica od desktop blokova, bez prijemnice, uz manjak druge

Pravila: tabela koje nema ili je prazna -> provera se tiho preskace; tabela postoji
a kolona fali -> jedan INFO "<KOD>_PRESKOCENO". Jedan nalaz po dokumentu po klasi:
redovi istog dokumenta (Klasa I i II: KooperantID+StanicaID+Datum+BrojDokumenta) se
spajaju. Redosled unosa = broj iz OtkupID (GetNextID = MAX+1); vreme unosa =
VremeUnosa, pa CreatedAt (StampRowAudit pri AppendRow).
"""

import re
import traceback
import unicodedata
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from decimal import Decimal
from statistics import median

from citac import broj as _broj
from citac import datum as _datum
from citac import kljuc, polje, tekst

from provera.zajednicko import GRESKA, INFO, SUMNJA, aktivni, cifre, d, ima_broj, k, kg_txt, n, norm_broj, parse_broj, s, storno

OTK, OTP, ZBR, PRJ, NOV = "tblOtkup", "tblOtpremnica", "tblZbirna", "tblPrijemnica", "tblNovac"
STA, VOZ, KOOP, KUL, CEN, SV, AMB = ("tblStanice", "tblVozaci", "tblKooperanti", "tblKulture", "tblCenovnik",
                                     "tblStornoVeze", "tblAmbalaza")

K_ZAST = "OTK_ZASTARELI_KONTEKST_DATUM"
K_VEZA = "OTK_VEZA_NESAGLASNA_SA_OTP"
K_PANEL = "OTK_PANEL_PREGAZIO_VEZU"
K_DLAN = "DUPLI_LANAC_HLADNJACE_POSLE_GUARD2"
K_BZ = "OTK_BROJ_ZBIRNE_ZAOSTAO"
K_VOZ = "OTK_VOZAC_ZAOSTAO_ILI_LOSE_IZVUCEN"
K_LIT = "VOZAC_HLADNJACA_LITERAL"
K_KOOP = "OTK_KOOPERANT_ZAOSTAO"
K_CENA = "OTK_CENA_TIP_ZAOSTALI"
K_CDES = "ZAOSTALA_CENA_POSLE_DESELEKCIJE"
K_BROJ = "OTK_BROJ_VAN_NIZA_ILI_PRAZAN"
K_PWA = "PWA_STANICA_I_DATUM"
K_DUPLI = "OTK_DUPLI_UNOS_BLOKA"
K_KUL = "OTK_KULTURAID_PO_VRSTI"
K_KDUP = "KOOP_DUPLIKAT_AUTOKREIRANJEM"
K_KNEP = "OTK_KOOPERANT_NEPOSTOJECI_ID"
K_GAJ = "OTK_GAJBE_KG_NEPROVERENO"
K_ZIVA = "OTK_AKTIVAN_BEZ_ZIVE_OTPREMNICE"
K_SYNC = "SYNC_AUTO_OTPREMNICA_ZA_NEVEZANE"

EPS = Decimal("0.01")
CENA_EPS = Decimal("0.005")
MAX = 12
RANG = {GRESKA: 0, SUMNJA: 1, INFO: 2}
_MALINA_DA = ("yes", "true", "1", "on", "enabled")              # IsMalinaMode, modConfig.bas:909-919
_FLAG_NE = ("no", "false", "0", "off", "disabled", "ne")       # ConfigFlag, modConfig.bas:922-933
_NEAKTIVAN = ("ne", "neaktivan", "no", "false", "0")
V270_PRVI_DAN = date(2026, 6, 30)   # vba-v2.7.0 (29.06 11:55): ClearOtkupFields prazni kooperanta, panel blokova
V270_NAJRANIJE = date(2026, 6, 28)  # b3b209e2 (28.06): ranije nema panela, pa ni dokaza o njemu
V283 = date(2026, 6, 30)            # vba-v2.8.3 (30.06 11:47): "Unesite broj dokumenta" / "Unesite broj gajbi"
PONOC_SAT = 4                       # VremeUnosa je lokalni Now(): unos do 04:00 za prethodni dan nije "kasni"
KOOP_RAZMAK = timedelta(minutes=10)
DUPLI_RAZMAK = timedelta(minutes=30)
MIN_UZORAK = 5
GORNJI, DONJI = Decimal("1.6"), Decimal("0.4")
_KANON = re.compile(r"^(\d+)/(\d{2})(\d{2})(\d{2})(?:-(\d+))?$")    # bez "S": S je samo zbirna
_TS = ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S.%f", "%Y-%m-%d %H:%M",
       "%d.%m.%Y %H:%M:%S", "%d.%m.%Y %H:%M", "%d/%m/%Y %H:%M:%S")
_EXCEL_NULA = datetime(1899, 12, 30)


# ---------------------------------------------------------------- pomocne
def _dmy(x):
    return x.strftime("%d.%m.%Y") if isinstance(x, date) else "?"


def _num(id_):
    c = cifre(id_)
    return c if c is not None else -1


def _kljuc_otk(r):
    return (_num(s(r, "OtkupID")), r.get("_redni") or 0)


def _lista(vrednosti, maks=MAX):
    v = list(vrednosti)
    if len(v) <= maks:
        return ", ".join(v)
    return ", ".join(v[:maks]) + " ... (+%d)" % (len(v) - maks)


def _ids(redovi, pk):
    out = []
    for r in sorted(redovi, key=lambda r: (_num(s(r, pk)), s(r, pk))):
        x = s(r, pk)
        if x and x not in out:
            out.append(x)
    return ", ".join(out)


def _po(redovi, kol):
    g = defaultdict(list)
    for r in redovi:
        v = k(r, kol)
        if v:
            g[v].append(r)
    return g


def _suma(redovi, kol):
    return sum((n(r, kol) for r in redovi), Decimal(0))


def _cena_txt(v):
    return kg_txt(v) or "0"


def _vreme(v):
    """(datetime ili None, date ili None) za VremeUnosa / CreatedAt / ModifiedAt."""
    if isinstance(v, datetime):
        return v, v.date()
    if isinstance(v, date):
        return None, v
    t = tekst(v)
    if not t:
        return None, None
    b = _broj(t)
    if b is not None:
        if Decimal(20000) <= b < Decimal(100000):
            dt = _EXCEL_NULA + timedelta(days=float(b))
            return (dt.replace(microsecond=0) if b != b.to_integral_value() else None), dt.date()
        return None, None
    for f in _TS:
        try:
            dt = datetime.strptime(t, f)
            return dt, dt.date()
        except ValueError:
            continue
    return None, _datum(t)


def _kljuc_dok(o):
    """Logicki dokument otkupa (dokumenti.py): KooperantID+StanicaID+Datum+BrojDokumenta."""
    b = k(o, "BrojDokumenta")
    if not b:
        return ("#", k(o, "OtkupID"), o.get("_redni") or 0)
    return (k(o, "KooperantID"), k(o, "StanicaID"), d(o, "Datum"), b)


def _dokumenti(redovi):
    g = {}
    for r in sorted(redovi, key=_kljuc_otk):
        g.setdefault(_kljuc_dok(r), []).append(r)
    return list(g.values())


def _o_txt(o):
    return "%s %s (%s, %s, %s kg, %s gajbi)" % (
        s(o, "OtkupID"), s(o, "BrojDokumenta") or "bez broja", s(o, "StanicaID") or "?", _dmy(d(o, "Datum")),
        kg_txt(n(o, "Kolicina")) or "0", kg_txt(n(o, "KolAmbalaze")) or "0")


def _p_txt(P):
    return "%s %s (%s, %s, %s kg)" % (
        s(P, "BrojOtpremnice") or "bez broja", s(P, "OtpremnicaID"), s(P, "StanicaID") or "?", _dmy(d(P, "Datum")),
        kg_txt(n(P, "Kolicina")) or "0")


def _ime_kljuc(r):
    """Ime + Prezime: casefold, bez kvacica (c/c/s/z, dj), sazeti razmaci."""
    t = ("%s %s" % (s(r, "Ime"), s(r, "Prezime"))).casefold().replace("\u0111", "dj")
    t = unicodedata.normalize("NFKD", t)
    t = "".join(ch for ch in t if not unicodedata.combining(ch))
    return " ".join(t.split())


def _izvuceni_id(display):
    """ExtractIDFromDisplay (vba-v2.11.0 modHelpers.bas:17-35): levo od prvog ' - ', inace izmedju
    PRVE '(' i prve ')', inace ceo tekst."""
    i = display.find(" - ")
    if i >= 0:
        return display[:i]
    a, b = display.find("("), display.find(")")
    if a >= 0 and b > a:
        return display[a + 1:b]
    return display


class _Sabirac:
    """Jedan nalaz po (kod, dokument): redovi istog dokumenta (Klasa I i II) se spajaju."""

    def __init__(self, pg, kod, tabela, pk):
        self.pg, self.kod, self.tabela, self.pk = pg, kod, tabela, pk
        self.g = {}

    def dodaj(self, redovi, oz, opis, predlog=""):
        kl = _kljuc_dok(redovi[0])
        e = self.g.get(kl)
        if e is None:
            e = self.g[kl] = {"redovi": [], "oz": oz, "opis": [], "predlog": []}
        for r in redovi:
            if not any(r is x for x in e["redovi"]):
                e["redovi"].append(r)
        if RANG[oz] < RANG[e["oz"]]:
            e["oz"] = oz
        if opis and opis not in e["opis"]:
            e["opis"].append(opis)
        if predlog and predlog not in e["predlog"]:
            e["predlog"].append(predlog)

    def emituj(self):
        for e in self.g.values():
            self.pg.nalaz(self.kod, e["oz"], self.tabela, e["redovi"], self.pk, " ".join(e["opis"]),
                          " ".join(e["predlog"]))


class _Pogled:
    """Indeksi koje provere ovog modula dele (racuna se jednom)."""

    def __init__(self, ctx):
        self.ctx = ctx
        self.preskoceno = set()
        self.prijavljeno = defaultdict(set)      # kod -> {id(red)} vec prijavljenih redova
        self.panel_a, self.panel_b, self.literal = set(), set(), set()
        t = ctx.t
        self.otk, self.otp, self.zbr, self.prj = t[OTK], t[OTP], t[ZBR], t[PRJ]
        self.otk_akt = [o for o in self.otk if not storno(o)]
        self.otp_akt = [r for r in self.otp if not storno(r)]
        self.otp_svi_po_id = _po(self.otp, "OtpremnicaID")
        self.otp_akt_po_id = {}
        for r in self.otp_akt:
            self.otp_akt_po_id.setdefault(k(r, "OtpremnicaID"), r)
        self.clan = _po(self.otk_akt, "OtpremnicaID")
        self.clan_svi = _po(self.otk, "OtpremnicaID")
        self.otp_akt_po_st_dan = defaultdict(list)
        for r in self.otp_akt:
            self.otp_akt_po_st_dan[(k(r, "StanicaID"), d(r, "Datum"))].append(r)
        self.otp_akt_po_broju = _po(self.otp_akt, "BrojOtpremnice")
        self.otp_akt_po_bz = defaultdict(list)
        for r in self.otp_akt:
            if norm_broj(polje(r, "BrojZbirne")):
                self.otp_akt_po_bz[norm_broj(polje(r, "BrojZbirne"))].append(r)
        self.zbr_po_b = _po(self.zbr, "BrojZbirne")
        self.zbr_akt_po_b = _po([z for z in self.zbr if not storno(z)], "BrojZbirne")
        self.prj_akt_po_b = _po([p for p in self.prj if not storno(p)], "BrojZbirne")
        self.malina = kljuc(ctx.config("MALINA_MODE")) in _MALINA_DA
        self.stanice = set(ctx.stanice)
        self.ima_jehl = bool(t[STA]) and ctx.ima_kolonu(STA, "JeHladnjaca")
        self.hladnjace = {k(r, "StanicaID") for r in t[STA] if k(r, "JeHladnjaca") == "da"}   # StrComp "Da"
        self.vozaci = {k(r, "VozacID") for r in t[VOZ] if k(r, "VozacID")}
        self.koop = {}
        for r in t[KOOP]:
            if k(r, "KooperantID"):
                self.koop.setdefault(k(r, "KooperantID"), r)
        self.otk_dok = _dokumenti(self.otk_akt)
        self.desktop = sorted([o for o in self.otk if not self.pwa(o)], key=_kljuc_otk)
        # najkasniji Datum aktivnih desktop redova unetih PRE reda (ID je redosled unosa)
        self.proxy = {}
        mx = None
        for o in self.desktop:
            self.proxy[id(o)] = mx
            dan = d(o, "Datum")
            if not storno(o) and dan is not None and (mx is None or dan > mx):
                mx = dan
        self._ts, self._lanac, self._sync = {}, {}, {}
        self._seq = None

    # --- kolone -----------------------------------------------------
    def ima(self, kod, potrebno):
        """False (tiho) kad tabela ne postoji ili je prazna; False + jedan INFO
        <kod>_PRESKOCENO kad tabela postoji a kolona fali."""
        for tb in potrebno:
            if not self.ctx.tab(tb):
                return False
        fali = [(tb, c) for tb, kol in potrebno.items() for c in kol if not self.ctx.ima_kolonu(tb, c)]
        if fali:
            if kod not in self.preskoceno:
                self.preskoceno.add(kod)
                self.ctx.nalaz(kod + "_PRESKOCENO", INFO, fali[0][0], "",
                               "Provera %s preskocena: nedostaje kolona %s." % (
                                   kod, ", ".join("%s.%s" % x for x in fali)))
            return False
        return True

    def nalaz(self, kod, oz, tabela, redovi, pk, opis, predlog=""):
        for r in redovi:
            self.prijavljeno[kod].add(id(r))
        self.ctx.nalaz(kod, oz, tabela, _ids(redovi, pk), opis, predlog)

    # --- redovi -----------------------------------------------------
    def pwa(self, o):
        """PWA red: ImportRowToTblOtkup pise ClientRecordID i SyncSource (modMasterSync.bas:1623-1627)."""
        return k(o, "SyncSource") == "pwa" or bool(s(o, "ClientRecordID"))

    def ts(self, r, kolone=("VremeUnosa", "CreatedAt")):
        key = (id(r), kolone)
        if key not in self._ts:
            v = (None, None)
            for kol in kolone:
                x = _vreme(polje(r, kol))
                if x[1] is not None:
                    v = x
                    break
            self._ts[key] = v
        return self._ts[key]

    def kasni(self, o):
        """(kasni, dan unosa, izvor): unet je posle svog Datum-a."""
        dan = d(o, "Datum")
        if dan is None:
            return False, None, ""
        vr, du = self.ts(o)
        if du is not None:
            if du <= dan:
                return False, du, ""
            if vr is not None and du == dan + timedelta(days=1) and vr.hour < PONOC_SAT:
                return False, du, ""
            return True, du, "vreme unosa %s" % (vr.strftime("%d.%m.%Y %H:%M") if vr else _dmy(du))
        px = self.proxy.get(id(o))
        if px is not None and px > dan:
            return True, px, "procena iz redosleda unosa: raniji unos ima Datum %s" % _dmy(px)
        return False, None, ""

    def ziva(self, oid):
        return self.otp_akt_po_id.get(kljuc(oid)) if tekst(oid) else None

    def clanovi(self, P):
        return self.clan.get(k(P, "OtpremnicaID"), [])

    def preostalo(self, P):
        return n(P, "Kolicina") - _suma(self.clanovi(P), "Kolicina")

    def prvi_blok(self, P):
        """BuildFirstBlokCena: prvi red tabele sa tim OtpremnicaID, bez storno filtera (modOtkupBlok.bas:1417-1446)."""
        M = self.clan_svi.get(k(P, "OtpremnicaID"), [])
        return min(M, key=lambda r: (r.get("_redni") or 0, _num(s(r, "OtkupID")))) if M else None

    def bez_prijemnice(self, P):
        bz = k(P, "BrojZbirne")
        return not (bz and self.prj_akt_po_b.get(bz))

    def prijemnice_txt(self, bz):
        pr = self.prj_akt_po_b.get(kljuc(bz), []) if tekst(bz) else []
        if not pr:
            return "bez prijemnice"
        return "prijemnica " + _lista("%s %s" % (s(p, "PrijemnicaID"), s(p, "BrojPrijemnice")) for p in pr)

    # --- hladnjaca --------------------------------------------------
    def lanac_x(self, o):
        """Aktivna otpremnica lanca hladnjace za otkup o (modAutoHladnjaca.bas:118-176): BrojOtpremnice =
        BrojDokumenta, ista stanica i dan. Hladnjaca = tblStanice.JeHladnjaca 'Da'; bez te kolone
        otisak lanca (ista kg, zbirna S+broj kupca hladnjace i prijemnica istih kg bez povrata)."""
        key = id(o)
        if key in self._lanac:
            return self._lanac[key]
        X = None
        b, st, dan = k(o, "BrojDokumenta"), k(o, "StanicaID"), d(o, "Datum")
        if b and st and dan is not None:
            kand = [x for x in self.otp_akt_po_broju.get(b, []) if k(x, "StanicaID") == st and d(x, "Datum") == dan]
            kand = [x for x in kand if k(x, "Klasa") == k(o, "Klasa")] or kand
            if kand:
                if self.ima_jehl:
                    X = kand[0] if st in self.hladnjace else None
                else:
                    X = next((x for x in kand if self._otisak_lanca(o, x)), None)
        self._lanac[key] = X
        return X

    def _otisak_lanca(self, o, x):
        if abs(n(x, "Kolicina") - n(o, "Kolicina")) > EPS:
            return False
        b = s(o, "BrojDokumenta")
        kupac = kljuc(self.ctx.malina_kupac)
        for bz in (kljuc("S" + b), kljuc(b)):
            zs = [z for z in self.zbr_akt_po_b.get(bz, []) if not kupac or k(z, "KupacID") == kupac]
            if zs and any(abs(n(p, "Kolicina") - n(o, "Kolicina")) <= EPS and n(p, "KolAmbVracena") == 0
                          for p in self.prj_akt_po_b.get(bz, [])):
                return True
        return False

    def je_lanac(self, o, X):
        lx = self.lanac_x(o)
        return X is not None and lx is not None and k(lx, "OtpremnicaID") == k(X, "OtpremnicaID")

    # --- sync -------------------------------------------------------
    def sync_like(self, A):
        """Otisak AutoCreateOtpremniceFromPWA (modMasterSync.bas:667-743): clanovi iste grupe
        StanicaID|Datum|VozacID|Klasa, kg = zbir, vrsta/sorta/cena sa PRVOG bloka, gajbe = zbir;
        otpremnica nije postojala pre blokova."""
        key = id(A)
        if key not in self._sync:
            self._sync[key] = self._sync_like(A)
        return self._sync[key]

    def _sync_like(self, A):
        M = self.clanovi(A)
        if not M or not k(A, "VozacID"):
            return False
        grupa = (k(A, "StanicaID"), d(A, "Datum"), k(A, "VozacID"), k(A, "Klasa"))
        if any((k(o, "StanicaID"), d(o, "Datum"), k(o, "VozacID"), k(o, "Klasa")) != grupa for o in M):
            return False
        if abs(n(A, "Kolicina") - _suma(M, "Kolicina")) > EPS:
            return False
        prvi = min(M, key=_kljuc_otk)
        if k(prvi, "VrstaVoca") != k(A, "VrstaVoca") or k(prvi, "SortaVoca") != k(A, "SortaVoca"):
            return False
        if abs(n(A, "Cena") - n(prvi, "Cena")) > CENA_EPS:
            return False
        if self.ctx.ima_kolonu(OTP, "KolAmbalaze") and self.ctx.ima_kolonu(OTK, "KolAmbalaze"):
            if abs(n(A, "KolAmbalaze") - _suma(M, "KolAmbalaze")) > EPS:
                return False
        if n(A, "BrutoKg") > 0:
            return False        # sync zove SaveOtpremnica_TX bez bruto-a; BrutoKg pise samo rucni unos (> 0)
        va = self.ts(A, ("CreatedAt",))[0]
        vm = [self.ts(o)[0] for o in M]
        if va is not None and vm and all(v is not None for v in vm) and va < min(vm) - timedelta(minutes=1):
            return False        # otpremnica je postojala pre blokova: panel, ne sync
        return True

    def sync_auto(self, A):
        """Sync otpremnica za koju podaci govore da nije rucna: otisak sync-a i zbirna bez prijemnice.
        Bez CreatedAt je sam otisak dvosmislen: popunjena rucna otpremnica panela ima iste clanove,
        zbir, vozaca, klasu i cenu (blok i otpremnica dobijaju cenu iz istog cenovnika)."""
        return self.sync_like(A) and self.bez_prijemnice(A)

    # --- kandidati --------------------------------------------------
    def kandidati_otp(self, o, dan, iskljuci=(), tol_dana=0, maks=3):
        """Aktivne otpremnice iste stanice (i vrste/sorte/klase kad su poznate), dan +-tol, sa
        preostalo >= kg otkupa. Rang: dan, |preostalo - kg|, ID."""
        if dan is None:
            return []
        st, kg = k(o, "StanicaID"), n(o, "Kolicina")
        out = []
        for delta in range(-tol_dana, tol_dana + 1):
            for Q in self.otp_akt_po_st_dan.get((st, dan + timedelta(days=delta)), []):
                if k(Q, "OtpremnicaID") in iskljuci:
                    continue
                if any(k(Q, c) and k(o, c) and k(Q, c) != k(o, c) for c in ("VrstaVoca", "SortaVoca", "Klasa")):
                    continue
                pre = self.preostalo(Q)
                if pre >= kg - EPS:
                    out.append((abs(delta), abs(pre - kg), _num(s(Q, "OtpremnicaID")), Q, pre))
        out.sort(key=lambda x: x[:3])
        return [(Q, pre) for _a, _b, _c, Q, pre in out[:maks]]

    @staticmethod
    def kandidati_txt(kand):
        return _lista("%s %s (preostalo %s kg)" % (s(Q, "OtpremnicaID"), s(Q, "BrojOtpremnice") or "bez broja",
                                                   kg_txt(pre) or "0") for Q, pre in kand)

    def vreme_storna(self, oid):
        """Najranije poznato vreme storna otpremnice: tblStornoVeze.CreatedAt, pa ModifiedAt reda
        (StampRowAudit pri UpdateCell) ako je posle CreatedAt. None ako nije stornirana / nepoznato."""
        rows = self.otp_svi_po_id.get(kljuc(oid), [])
        if not rows or not all(storno(r) for r in rows):
            return None
        vremena = []
        brojevi = {k(r, "BrojOtpremnice") for r in rows if k(r, "BrojOtpremnice")}
        for R in self.ctx.t[SV]:
            # po broju samo kad red nema OldDocID: posle storna i ponovnog unosa isti broj nosi i ranije
            # stornirana otpremnica, pa bi njeno (ranije) vreme storna lazno "dokazalo" unos posle storna
            if k(R, "OldDocID") == kljuc(oid) or (not k(R, "OldDocID") and k(R, "OldDocType") == "otpremnica"
                                                   and k(R, "OldBroj") in brojevi):
                v = _vreme(polje(R, "CreatedAt"))[0]
                if v is not None:
                    vremena.append(v)
        if not vremena:
            for r in rows:
                vm, vc = _vreme(polje(r, "ModifiedAt"))[0], _vreme(polje(r, "CreatedAt"))[0]
                if vm is not None and (vc is None or vm > vc):
                    vremena.append(vm)
        return min(vremena) if vremena else None

    def sledeci_broj(self, st_id, dan):
        """Predlog kanona x/ddmmyy[-n]: n = MAX(seq)+1 nad tblOtkup za StanicaID+Datum, i stornirani
        (MaxSeqFromTable, modBrojevi.bas:346-388)."""
        x = cifre(st_id)
        if x is None or dan is None:
            return ""
        if self._seq is None:
            self._seq = defaultdict(int)
            for r in self.otk:
                p = parse_broj(s(r, "BrojDokumenta"))
                if p is not None and d(r, "Datum") is not None:
                    key = (k(r, "StanicaID"), d(r, "Datum"))
                    self._seq[key] = max(self._seq[key], p.seq)
        mx = self._seq.get((kljuc(st_id), dan), 0)
        b = "%d/%s" % (x, dan.strftime("%d%m%y"))
        return b if mx == 0 else "%s-%d" % (b, mx + 1)


# ---------------------------------------------------------------- VOZAC_HLADNJACA_LITERAL
def _literal(pg):
    ctx = pg.ctx
    naziv = {OTK: "otkup", OTP: "otpremnica", ZBR: "zbirna", PRJ: "prijemnica", AMB: "ambalaza"}
    uzrok = []
    for r in ctx.t[VOZ]:
        disp = "%s %s (%s)" % (s(r, "Ime"), s(r, "Prezime"), s(r, "VozacID"))
        if kljuc(_izvuceni_id(disp)) != k(r, "VozacID"):
            uzrok.append((kljuc(_izvuceni_id(disp)), s(r, "VozacID"), disp.strip()))
    for ime, pk in ((OTK, "OtkupID"), (OTP, "OtpremnicaID"), (ZBR, "ZbirnaID"), (PRJ, "PrijemnicaID"),
                    (AMB, "AmbID")):
        if not ctx.t[ime] or not pg.ima(K_LIT, {ime: [pk, "VozacID"]}):
            continue
        g = defaultdict(list)
        for r in ctx.t[ime]:
            v = s(r, "VozacID")
            if storno(r) or not v or kljuc(v) in pg.vozaci:
                continue
            if kljuc(v) == "hladnjaca" or not any(ch.isdigit() for ch in v):
                g[kljuc(v)].append(r)
        for kv, rr in g.items():
            st = Counter(_stanica_reda(pg, ime, r) or "?" for r in rr)
            opis = ("%d aktivnih redova (%s) ima VozacID '%s': ExtractIDFromDisplay uzima tekst iz PRVE zagrade "
                    "prikaza 'Ime Prezime (ID)', a bez zagrade ceo tekst (modHelpers.bas:17-35); par-vozac "
                    "hladnjace ima Prezime '(hladnjaca)' (modAutoHladnjaca.bas:79-85)." % (
                        len(rr), naziv[ime], s(rr[0], "VozacID")))
            izvor = [x for x in uzrok if x[0] == kv]
            if izvor:
                opis += " Izvor u tblVozaci: %s." % _lista("'%s' (%s)" % (x[2], x[1]) for x in izvor)
            predlog = ("VozacID '%s' -> StanicaID stanice reda (%s); otkup, otpremnicu, zbirnu i prijemnicu lanca "
                       "ispraviti zajedno (prijemnica se za zbirnu vezuje i po VozacID)." % (
                           s(rr[0], "VozacID"), _lista("%s: %d" % (x, c) for x, c in sorted(st.items()))))
            pg.nalaz(K_LIT, GRESKA, ime, rr, pk, opis, predlog)
            if ime == OTK:
                pg.literal.update(id(r) for r in rr)


def _stanica_reda(pg, ime, r):
    if ime in (OTK, OTP):
        return s(r, "StanicaID")
    if ime in (ZBR, PRJ):
        return pg.ctx.stanica_iz_broja(s(r, "BrojZbirne"))
    if ime == AMB and k(r, "EntitetTip") == "stanica":
        return s(r, "EntitetID")
    return ""


# ---------------------------------------------------------------- OTK_PANEL_PREGAZIO_VEZU (A, B)
def _panel(pg):
    if not pg.ima(K_PANEL, {OTK: ["OtkupID", "BrojDokumenta", "StanicaID", "Datum", "Kolicina", "BrojZbirne",
                                  "OtpremnicaID", "Stornirano"],
                            OTP: ["OtpremnicaID", "BrojOtpremnice", "BrojZbirne", "StanicaID", "Datum", "Kolicina",
                                  "Stornirano"]}):
        return
    sab = _Sabirac(pg, K_PANEL, OTK, "OtkupID")
    for dok in pg.otk_dok:
        for o in dok:
            oid = s(o, "OtpremnicaID")
            Y = pg.ziva(oid)
            X = pg.lanac_x(o)
            if X is not None and abs(n(X, "Kolicina") - n(o, "Kolicina")) > EPS:
                X = None              # isti broj, druga kg: rucna otpremnica, ne lanac ovog bloka
            if X is not None and not pg.clanovi(X):
                lan = "lanac %s, zbirna %s, %s" % (_p_txt(X), s(X, "BrojZbirne") or "-",
                                                    pg.prijemnice_txt(s(X, "BrojZbirne")))
                if Y is not None and k(Y, "OtpremnicaID") != k(X, "OtpremnicaID"):
                    bz = norm_broj(polje(o, "BrojZbirne"))
                    if not bz or bz != norm_broj(polje(X, "BrojZbirne")) or bz == norm_broj(polje(Y, "BrojZbirne")):
                        continue
                    opis = ("A (pregazeno): hladnjacki otkup %s ima svoj %s bez ijednog bloka, a OtpremnicaID "
                            "pokazuje na %s (zbirna %s, %s); otkup i dalje nosi BrojZbirne lanca %s. Panel "
                            "(LinkOtkupIDsToOtpremnica bez guard-a, do v2.11.0) je posle lanca prepisao vezu: "
                            "roba je dva puta u lancu." % (
                                _o_txt(o), lan, _p_txt(Y), s(Y, "BrojZbirne") or "-",
                                pg.prijemnice_txt(s(Y, "BrojZbirne")), s(o, "BrojZbirne")))
                    predlog = ("Smer se ne predlaze automatski: ako je %s prava tura (ima svoju zbirnu i prijemnicu "
                               "kod kupca), lanac %s je duplikat -> stornirati lanac (otpremnica, zbirna %s, "
                               "prijemnica); inace OtpremnicaID: %s -> %s (posle: preostalo %s %s kg, %s %s kg)." % (
                                   s(Y, "OtpremnicaID"), s(X, "OtpremnicaID"), s(X, "BrojZbirne") or "-",
                                   s(Y, "OtpremnicaID"), s(X, "OtpremnicaID"), s(X, "OtpremnicaID"),
                                   kg_txt(n(X, "Kolicina") - n(o, "Kolicina")) or "0", s(Y, "OtpremnicaID"),
                                   kg_txt(pg.preostalo(Y) + n(o, "Kolicina")) or "0"))
                    sab.dodaj([o], GRESKA, opis, predlog)
                    pg.panel_a.add(id(o))
                    continue
                if Y is None:
                    opis = ("B: hladnjacki otkup %s %s, a njegov %s nema nijedan blok." % (
                        _o_txt(o), "pokazuje na storniranu ili nepostojecu otpremnicu %s" % oid if oid
                        else "nije vezan ni za jednu otpremnicu", lan))
                    predlog = "OtpremnicaID: %s -> %s (lanac istog broja, dana i kg)." % (
                        oid or "(prazno)", s(X, "OtpremnicaID"))
                    sab.dodaj([o], GRESKA if oid else SUMNJA, opis, predlog)
                    pg.panel_b.add(id(o))
                    continue
            if oid and Y is None:
                t_st = pg.vreme_storna(oid)
                vo = pg.ts(o)[0]
                if t_st is not None and vo is not None and vo > t_st:
                    stara = pg.otp_svi_po_id.get(kljuc(oid), [])
                    opis = ("B: otkup %s je unet %s, POSLE storna otpremnice %s (%s), a vezan je za nju: panel je "
                            "do v2.16.1 vezivao blok za izabranu otpremnicu bez provere storna." % (
                                _o_txt(o), vo.strftime("%d.%m.%Y %H:%M"),
                                _p_txt(stara[0]) if stara else oid, t_st.strftime("%d.%m.%Y %H:%M")))
                    kand = pg.kandidati_otp(o, d(o, "Datum"), tol_dana=1)
                    predlog = ("OtpremnicaID: %s -> %s." % (oid, pg.kandidati_txt(kand)) if kand else
                               "OtpremnicaID isprazniti ili vezati za zivu otpremnicu iste stanice i dana "
                               "(kandidat sa dovoljno preostalog kg nije nadjen).")
                    sab.dodaj([o], GRESKA, opis, predlog)
                    pg.panel_b.add(id(o))
    sab.emituj()


# ---------------------------------------------------------------- DUPLI_LANAC_HLADNJACE_POSLE_GUARD2 (C)
def _dupli_lanac(pg):
    if not pg.ima_jehl or not pg.hladnjace:
        return
    if not pg.ima(K_DLAN, {OTK: ["OtkupID", "BrojDokumenta", "StanicaID", "Datum", "Kolicina", "OtpremnicaID",
                                 "Stornirano"],
                           OTP: ["OtpremnicaID", "BrojOtpremnice", "BrojZbirne", "StanicaID", "Datum", "Kolicina",
                                 "Stornirano"]}):
        return
    brojevi = {(k(o, "StanicaID"), k(o, "BrojDokumenta")) for o in pg.otk_akt if k(o, "BrojDokumenta")}
    lanci = defaultdict(list)
    for o in pg.otk_akt:
        if k(o, "StanicaID") not in pg.hladnjace:
            continue
        X = pg.ziva(s(o, "OtpremnicaID"))
        if X is not None and pg.je_lanac(o, X):
            lanci[(k(o, "StanicaID"), d(o, "Datum"))].append((o, X))
    for (st, dan), par in lanci.items():
        uzeti = set()
        for Y in sorted(pg.otp_akt_po_st_dan.get((st, dan), []), key=lambda r: _num(s(r, "OtpremnicaID"))):
            if (st, k(Y, "BrojOtpremnice")) in brojevi:
                continue                      # Y je lanac nekog otkupa
            pre = pg.preostalo(Y)
            if pre <= EPS:
                continue
            slobodni = [(o, X) for o, X in par if id(o) not in uzeti and (not k(Y, "Klasa") or k(Y, "Klasa") == k(o, "Klasa"))]
            pog = [p for p in slobodni if abs(pre - n(p[0], "Kolicina")) <= EPS][:1]
            if not pog and len(slobodni) > 1 and abs(pre - sum(n(o, "Kolicina") for o, _x in slobodni)) <= EPS:
                pog = slobodni
            if not pog:
                continue
            uzeti.update(id(o) for o, _x in pog)
            lanci_txt = "; ".join("%s -> lanac %s, zbirna %s, %s" % (
                _o_txt(o), _p_txt(X), s(X, "BrojZbirne") or "-", pg.prijemnice_txt(s(X, "BrojZbirne")))
                for o, X in pog)
            opis = ("C (dupli lanac, v2.16.1..v2.28.4): na hladnjaci %s, %s blok(ovi) su vezani za svoj lanac: %s. "
                    "Ista stanica i dan imaju i otpremnicu %s (zbirna %s, %s) bez blokova za tacno te kg "
                    "(preostalo %s kg): lanac se zove bezuslovno, a guard 2 ne vezuje blok za panelom izabranu "
                    "otpremnicu (frmOtkup.frm:1131)." % (
                        s(pog[0][0], "StanicaID"), _dmy(dan), lanci_txt, _p_txt(Y), s(Y, "BrojZbirne") or "-",
                        pg.prijemnice_txt(s(Y, "BrojZbirne")), kg_txt(pre)))
            predlog = ("Smer se ne predlaze automatski: ako je %s prava tura (ima svoju zbirnu i prijemnicu kod "
                       "kupca), lanac je duplikat -> stornirati lanac %s (otpremnica, zbirna, prijemnica) i "
                       "OtpremnicaID -> %s; ako %s nije isla kupcu -> stornirati %s." % (
                           s(Y, "OtpremnicaID"), _ids([X for _o, X in pog], "OtpremnicaID"), s(Y, "OtpremnicaID"),
                           s(Y, "OtpremnicaID"), s(Y, "OtpremnicaID")))
            pg.nalaz(K_DLAN, SUMNJA, OTK, [o for o, _x in pog], "OtkupID", opis, predlog)


# ---------------------------------------------------------------- OTK_ZASTARELI_KONTEKST_DATUM
def _zastareli(pg):
    if not pg.ima(K_ZAST, {OTK: ["OtkupID", "Datum", "StanicaID", "OtpremnicaID", "BrojZbirne", "Kolicina",
                                 "Stornirano"],
                           OTP: ["OtpremnicaID", "Datum", "StanicaID", "BrojZbirne", "Kolicina", "Stornirano"]}):
        return
    sab = _Sabirac(pg, K_ZAST, OTK, "OtkupID")
    # niz po stanici (ispravljen predikat skeptika), ali txtDatum je JEDNO polje forme: desktop unos bilo koje
    # stanice sa drugim Datum-om znaci da je datum forme u medjuvremenu promenjen, pa prekida i tudji niz.
    niz = {}                                  # StanicaID -> (P, L)
    for o in pg.desktop:
        if storno(o):
            continue
        st, dan = k(o, "StanicaID"), d(o, "Datum")
        for st2 in [x for x, (P2, _L2) in niz.items() if x != st and d(P2, "Datum") != dan]:
            del niz[st2]
        P, L = niz.get(st, (None, None))
        Po = pg.ziva(s(o, "OtpremnicaID"))
        if (P is not None and dan is not None and dan == d(P, "Datum")
                and (Po is None or k(Po, "OtpremnicaID") != k(P, "OtpremnicaID"))):
            tip = _tip_veze(pg, o, Po)
            if tip:
                if pg.kasni(o)[0]:
                    _zastareli_kandidat(pg, sab, o, P, L, Po, tip)
                else:
                    niz.pop(st, None)     # unet istog dana: datum forme je bio stvaran, niz posle L je prekinut
                continue
        if Po is not None and not pg.je_lanac(o, Po):
            niz[st] = (Po, o)
        elif P is not None and dan != d(P, "Datum"):
            niz.pop(st, None)
    sab.emituj()


def _tip_veze(pg, o, Po):
    oid = s(o, "OtpremnicaID")
    if not oid:
        return "nije vezan ni za jednu otpremnicu"
    if Po is None:
        return "vezan je za storniranu ili nepostojecu otpremnicu %s" % oid
    if pg.je_lanac(o, Po):
        return "vezan je za svoj lanac hladnjace %s" % _p_txt(Po)
    if (pg.sync_like(Po) and pg.bez_prijemnice(Po) and k(Po, "StanicaID") == k(o, "StanicaID")
            and d(Po, "Datum") == d(o, "Datum")):
        return "vezan je za sync auto-otpremnicu %s (zbirna %s bez prijemnice)" % (_p_txt(Po), s(Po, "BrojZbirne") or "-")
    return ""


def _zastareli_kandidat(pg, sab, o, P, L, Po, tip):
    kasni, du, izvor = pg.kasni(o)
    if not kasni:
        return
    bz_o = norm_broj(polje(o, "BrojZbirne"))
    a = bool(bz_o) and bz_o == norm_broj(polje(P, "BrojZbirne"))
    pre = pg.preostalo(P)
    b = pre <= EPS
    if not (a or b):
        return
    dokazi = []
    if a:
        dokazi.append("BrojZbirne %s je zbirna te otpremnice" % s(o, "BrojZbirne"))
    if b:
        dokazi.append("otpremnica je popunjena (preostalo %s kg), pa je panel skinuo izbor" % (kg_txt(pre) or "0"))
    prvi = pg.prvi_blok(P)
    if prvi is not None and n(o, "Cena") > 0 and abs(n(o, "Cena") - n(prvi, "Cena")) <= CENA_EPS:
        dokazi.append("cena %s je cena blokova te otpremnice" % _cena_txt(n(o, "Cena")))
    bp = parse_broj(s(o, "BrojDokumenta"))
    if bp is not None and bp.dan == d(o, "Datum"):
        dokazi.append("i broj %s je predlozen za stari dan" % s(o, "BrojDokumenta"))
    opis = ("Otkup %s je unet posle svog datuma (%s), posle bloka %s otpremnice %s, sa njenim datumom; %s. "
            "Dokaz: %s. Auto-deselekcija (modOtkupBlok AfterUnos) ne vraca txtDatum/txtBrojZbirne." % (
                _o_txt(o), izvor, s(L, "OtkupID"), _p_txt(P), tip, "; ".join(dokazi)))
    delovi = ["Datum: %s -> %s (dan unosa; potvrditi sa papirnim otkupnim listom)" % (_dmy(d(o, "Datum")), _dmy(du))]
    kand = pg.kandidati_otp(o, du, iskljuci={k(P, "OtpremnicaID")})
    if kand:
        delovi.append("kandidat otpremnica za %s: %s" % (_dmy(du), pg.kandidati_txt(kand)))
    else:
        delovi.append("aktivna otpremnica za %s sa dovoljno preostalog kg nije nadjena" % _dmy(du))
    if Po is not None and pg.je_lanac(o, Po):
        delovi.append("na hladnjaci ispraviti i Datum lanca %s (zbirna %s, %s)" % (
            s(Po, "OtpremnicaID"), s(Po, "BrojZbirne") or "-", pg.prijemnice_txt(s(Po, "BrojZbirne"))))
    delovi.append("broj dokumenta se ne prenumerise automatski")
    sab.dodaj([o], SUMNJA, opis, "; ".join(delovi) + ".")


# ---------------------------------------------------------------- OTK_VEZA_NESAGLASNA_SA_OTP
def _veza(pg):
    if not pg.ima(K_VEZA, {OTK: ["OtkupID", "OtpremnicaID", "StanicaID", "Datum", "VrstaVoca", "SortaVoca", "Klasa",
                                 "Kolicina", "Stornirano"],
                           OTP: ["OtpremnicaID", "StanicaID", "Datum", "VrstaVoca", "SortaVoca", "Klasa", "Kolicina",
                                 "Stornirano"]}):
        return
    sab = _Sabirac(pg, K_VEZA, OTK, "OtkupID")
    for dok in pg.otk_dok:
        klase = defaultdict(set)
        for o in dok:
            klase[k(o, "OtpremnicaID")].add(k(o, "Klasa"))
        for o in dok:
            P = pg.ziva(s(o, "OtpremnicaID"))
            if P is None:
                continue
            asp = []
            if k(o, "StanicaID") and k(P, "StanicaID") and k(o, "StanicaID") != k(P, "StanicaID"):
                if pg.pwa(o):
                    # PWA uvoz uzima StanicaID iz maticnog OM kooperanta (modMasterSync.bas:1575): veza moze biti
                    # tacna, a pogresna StanicaID otkupa (PWA_STANICA_I_DATUM).
                    asp.append(("STANICA: PWA otkup %s, otpremnica %s (PWA uvoz uzima StanicaID iz maticnog OM "
                                "kooperanta -- proveriti da li je pogresna stanica otkupa, a veza tacna)" % (
                                    s(o, "StanicaID"), s(P, "StanicaID")), SUMNJA))
                else:
                    asp.append(("STANICA: otkup %s, otpremnica %s" % (s(o, "StanicaID"), s(P, "StanicaID")), GRESKA))
            do, dp = d(o, "Datum"), d(P, "Datum")
            if do is not None and dp is not None and do != dp:
                raz = (do - dp).days
                if raz > 0:
                    asp.append(("DATUM: otkup je %d dan(a) POSLE otpremnice (%s vs %s) -- roba ne moze na raniju "
                                "otpremnicu" % (raz, _dmy(do), _dmy(dp)), GRESKA))
                else:
                    asp.append(("DATUM: otkup je %d dan(a) pre otpremnice (%s vs %s)%s" % (
                        -raz, _dmy(do), _dmy(dp), " -- moguc vecernji otkup i jutarnja otpremnica" if raz == -1 else ""),
                        SUMNJA))
            uzrok = " (uzrok: sync auto-otpremnica uzima vrstu/sortu sa prvog bloka grupe)" if pg.sync_auto(P) else ""
            if k(o, "VrstaVoca") and k(P, "VrstaVoca") and k(o, "VrstaVoca") != k(P, "VrstaVoca"):
                asp.append(("VRSTA: otkup %s, otpremnica %s%s" % (s(o, "VrstaVoca"), s(P, "VrstaVoca"), uzrok), GRESKA))
            elif k(o, "SortaVoca") and k(P, "SortaVoca") and k(o, "SortaVoca") != k(P, "SortaVoca"):
                asp.append(("SORTA: otkup %s, otpremnica %s%s" % (s(o, "SortaVoca"), s(P, "SortaVoca"), uzrok), SUMNJA))
            if (k(o, "Klasa") and k(P, "Klasa") and k(o, "Klasa") != k(P, "Klasa")
                    and len(klase[k(o, "OtpremnicaID")]) < 2):
                asp.append(("KLASA: otkup %s, otpremnica %s" % (s(o, "Klasa"), s(P, "Klasa")), SUMNJA))
            if not asp:
                continue
            oz = min((z for _t, z in asp), key=RANG.get)
            pre_p = pg.preostalo(P)
            opis = "Otkup %s vezan je za otpremnicu %s koja se ne slaze: %s. Preostalo na otpremnici: %s kg." % (
                _o_txt(o), _p_txt(P), "; ".join(t for t, _z in asp), kg_txt(pre_p) or "0")
            kand = pg.kandidati_otp(o, do, iskljuci={k(P, "OtpremnicaID")})
            if kand:
                Q, pre_q = kand[0]
                jak = abs(pre_q - n(o, "Kolicina")) <= EPS and abs(pre_p + n(o, "Kolicina")) <= EPS
                predlog = "OtpremnicaID: %s -> %s (%s, preostalo %s kg)%s." % (
                    s(P, "OtpremnicaID"), s(Q, "OtpremnicaID"), s(Q, "BrojOtpremnice") or "bez broja",
                    kg_txt(pre_q) or "0", "; premestaj zatvara obe otpremnice" if jak else "")
                if len(kand) > 1:
                    predlog += " Ostali kandidati: %s." % pg.kandidati_txt(kand[1:])
            else:
                predlog = ("Kandidat (ista stanica, dan, vrsta/sorta, dovoljno preostalog kg) nije nadjen -- "
                           "odluka operatera.")
            sab.dodaj([o], oz, opis, predlog)
    sab.emituj()


# ---------------------------------------------------------------- OTK_BROJ_ZBIRNE_ZAOSTAO
def _broj_zbirne(pg):
    if not pg.ima(K_BZ, {OTK: ["OtkupID", "BrojZbirne", "OtpremnicaID", "StanicaID", "Datum", "Stornirano"],
                         OTP: ["OtpremnicaID", "BrojZbirne", "StanicaID", "Datum", "Stornirano"]}):
        return
    sab = _Sabirac(pg, K_BZ, OTK, "OtkupID")
    vec = pg.prijavljeno[K_ZAST] | pg.panel_a
    for dok in pg.otk_dok:
        for o in dok:
            bz = norm_broj(polje(o, "BrojZbirne"))
            if not bz or id(o) in vec:
                continue
            zb = pg.zbr_po_b.get(k(o, "BrojZbirne"), [])
            if zb and all(storno(z) for z in zb):
                continue          # zbirna postoji samo stornirana: OTK_BROJZBIRNE_NA_STORNIRANOJ_ZBIRNOJ (storno)
            oid = s(o, "OtpremnicaID")
            P = pg.ziva(oid)
            iste = [Q for Q in pg.otp_akt_po_bz.get(bz, [])
                    if k(Q, "StanicaID") == k(o, "StanicaID") and d(Q, "Datum") == d(o, "Datum")]
            if P is not None:
                bp = norm_broj(polje(P, "BrojZbirne"))
                if not bp or bp == bz:
                    continue
                druge = [Q for Q in iste if Q is not P]
                predlog = "BrojZbirne: %s -> %s (zbirna otpremnice %s)." % (
                    s(o, "BrojZbirne"), s(P, "BrojZbirne"), s(P, "OtpremnicaID"))
                if druge:
                    sab.dodaj([o], SUMNJA, (
                        "Otkup %s je na otpremnici %s (zbirna %s), a nosi BrojZbirne %s druge aktivne otpremnice %s "
                        "iste stanice i dana: zaostalo polje txtBrojZbirne ili pregazena veza -- proveriti koja je "
                        "otpremnica prava." % (_o_txt(o), _p_txt(P), s(P, "BrojZbirne"), s(o, "BrojZbirne"),
                                                _ids(druge, "OtpremnicaID"))), predlog)
                else:
                    sab.dodaj([o], INFO, (
                        "Otkup %s nosi BrojZbirne %s, a njegova otpremnica %s ima zbirnu %s; nijedna druga otpremnica "
                        "iste stanice i dana nema %s -- verovatno kasnija preraspodela zbirne (backfill dopunjuje "
                        "samo prazno polje)." % (_o_txt(o), s(o, "BrojZbirne"), _p_txt(P), s(P, "BrojZbirne"),
                                                  s(o, "BrojZbirne"))), predlog)
            elif not oid and not iste:
                nap = ""
                p = parse_broj(s(o, "BrojZbirne"))
                if p is not None and p.s:
                    if cifre(s(o, "StanicaID")) is not None and p.x != cifre(s(o, "StanicaID")):
                        nap += " Prefiks S%d pripada stanici %s." % (p.x, pg.ctx.stanica_po_broju.get(p.x, "?"))
                    if p.dan is not None and d(o, "Datum") is not None and p.dan != d(o, "Datum"):
                        nap += " Dan u broju zbirne je %s." % _dmy(p.dan)
                sab.dodaj([o], INFO, (
                    "Nevezan otkup %s nosi BrojZbirne %s, a nijedna aktivna otpremnica stanice %s za %s nema tu "
                    "zbirnu: polje je zaostalo iz prethodnog unosa (ClearOtkupFields ne prazni txtBrojZbirne).%s" % (
                        _o_txt(o), s(o, "BrojZbirne"), s(o, "StanicaID") or "?", _dmy(d(o, "Datum")), nap)),
                    "Polje je samo dokaz konteksta (migracija ga prazni); utvrditi za koju otpremnicu je blok.")
    sab.emituj()


# ---------------------------------------------------------------- OTK_VOZAC_ZAOSTAO_ILI_LOSE_IZVUCEN
def _vozac(pg):
    ctx = pg.ctx
    if pg.otk and pg.ima(K_VOZ, {OTK: ["OtkupID", "VozacID", "StanicaID", "OtpremnicaID", "Stornirano"]}):
        sab = _Sabirac(pg, K_VOZ, OTK, "OtkupID")
        for dok in pg.otk_dok:
            for o in dok:
                voz, st = k(o, "VozacID"), k(o, "StanicaID")
                if not voz or id(o) in pg.literal:
                    continue
                P = pg.ziva(s(o, "OtpremnicaID"))
                pv = k(P, "VozacID") if P is not None else ""
                if pg.vozaci and voz not in pg.vozaci and voz not in pg.stanice:
                    cilj = s(o, "StanicaID") if pg.malina else (s(P, "VozacID") if pv else "vozac sa otpremnice")
                    sab.dodaj([o], GRESKA, (
                        "Otkup %s ima VozacID '%s' koji ne postoji u tblVozaci: ID je lose izvucen iz teksta combo-a "
                        "(ExtractIDFromDisplay / rucno ukucan vozac)." % (_o_txt(o), s(o, "VozacID"))),
                        "VozacID: %s -> %s." % (s(o, "VozacID"), cilj or "?"))
                    continue
                if pg.malina:
                    if voz == st or (P is not None and pv == voz):
                        continue      # vozac = stanica, ili potvrdjen otpremnicom (OTP_VOZAC_NIJE_PAR_STANICE)
                    if voz in pg.stanice or P is not None:
                        sab.dodaj([o], SUMNJA, (
                            "Malina: otkup %s ima VozacID %s%s; stanica otkupa je %s%s. cmbVozac se posle snimanja "
                            "ne prazni, a sync grupise auto-otpremnicu i po VozacID." % (
                                _o_txt(o), s(o, "VozacID"), " (par-vozac druge stanice)" if voz in pg.stanice else "",
                                s(o, "StanicaID") or "?",
                                "; otpremnica %s ima VozacID %s" % (s(P, "OtpremnicaID"), s(P, "VozacID") or "-")
                                if P is not None else "")),
                            "VozacID: %s -> %s (malina: vozac otkupa = stanica)." % (s(o, "VozacID"), s(o, "StanicaID")))
                elif P is not None and pv and pv != voz:
                    sab.dodaj([o], SUMNJA, (
                        "Otkup %s ima VozacID %s, a njegova otpremnica %s VozacID %s (panel prepisuje vozaca sa "
                        "otpremnice; razlika = zaostao izbor)." % (_o_txt(o), s(o, "VozacID"), _p_txt(P),
                                                                     s(P, "VozacID"))),
                        "VozacID: %s -> %s." % (s(o, "VozacID"), s(P, "VozacID")))
        sab.emituj()
    # uzrok: vozac cije ime ima zagradu daje pogresan izvuceni ID
    if not ctx.t[VOZ] or not ctx.ima_kolonu(VOZ, "VozacID"):
        return
    upotreba = Counter()
    for ime in (OTK, OTP, ZBR, PRJ):
        for r in ctx.t[ime]:
            if not storno(r) and k(r, "VozacID"):
                upotreba[k(r, "VozacID")] += 1
    for r in ctx.t[VOZ]:
        if k(r, "Aktivan") in _NEAKTIVAN or not s(r, "VozacID"):
            continue
        disp = "%s %s (%s)" % (s(r, "Ime"), s(r, "Prezime"), s(r, "VozacID"))   # GetVozacDisplayList
        izv = _izvuceni_id(disp)
        if kljuc(izv) == k(r, "VozacID"):
            continue
        ctx.nalaz(K_VOZ, INFO, VOZ, s(r, "VozacID"), (
            "Vozac %s: prikaz u combo-u '%s' daje izvuceni ID '%s' umesto %s (ExtractIDFromDisplay uzima levo od "
            "prvog ' - ', pa tekst iz PRVE zagrade); %d aktivnih dokumenata nosi '%s'." % (
                s(r, "VozacID"), disp.strip(), izv.strip(), s(r, "VozacID"), upotreba[kljuc(izv)], izv.strip())),
            "Ime/Prezime bez zagrade i bez ' - ' pre daljeg unosa; dokumente ispraviti po VOZAC_HLADNJACA_LITERAL i "
            "%s." % K_VOZ)


# ---------------------------------------------------------------- OTK_KOOPERANT_ZAOSTAO
def _koop_zaostao(pg):
    ctx = pg.ctx
    if not pg.ima(K_KOOP, {OTK: ["OtkupID", "KooperantID", "StanicaID", "Datum", "BrojDokumenta", "Stornirano"]}):
        return
    seq = pg.desktop
    if len(seq) < 2:
        return
    granica, izvor = V270_PRVI_DAN, "vba-v2.7.0 (29.06.2026)"
    dokaz = []
    for o in pg.otk_akt:
        if pg.pwa(o):
            continue
        P = pg.ziva(s(o, "OtpremnicaID"))
        vo, du = pg.ts(o)
        vp = pg.ts(P, ("CreatedAt",))[0] if P is not None else None
        if vo is not None and vp is not None and vp < vo and du >= V270_NAJRANIJE and not pg.je_lanac(o, P):
            dokaz.append(du)
    if dokaz:
        granica, izvor = min(dokaz), "prvi dan kad se vidi panel blokova (blok dodat na ranije unetu otpremnicu)"
    prvi_posle = min((_num(s(o, "OtkupID")) for o in seq if d(o, "Datum") is not None and d(o, "Datum") >= granica),
                     default=None)

    def pre(o):
        du = pg.ts(o)[1]
        if du is not None:
            return du < granica
        return prvi_posle is None or _num(s(o, "OtkupID")) < prvi_posle

    n_pre = n_posle = p_posle = 0
    parovi = []
    for i in range(1, len(seq)):
        r1, r2 = seq[i - 1], seq[i]
        if storno(r2):
            continue
        u = pre(r2)
        if u:
            n_pre += 1
        else:
            n_posle += 1
        if storno(r1) or not k(r1, "KooperantID") or k(r1, "KooperantID") != k(r2, "KooperantID"):
            continue
        if k(r1, "StanicaID") != k(r2, "StanicaID"):
            continue
        if norm_broj(polje(r1, "BrojDokumenta")) == norm_broj(polje(r2, "BrojDokumenta")):
            continue            # Klasa I + II istog lista
        t1, t2 = pg.ts(r1)[0], pg.ts(r2)[0]
        if t1 is not None and t2 is not None and not (timedelta(0) <= t2 - t1 <= KOOP_RAZMAK):
            continue
        if u:
            parovi.append((r1, r2, t1, t2))
        else:
            p_posle += 1
    if not parovi:
        return
    stopa_pre = len(parovi) / n_pre if n_pre else 0.0
    stopa_posle = p_posle / n_posle if n_posle >= 30 else None
    oz = INFO
    procena = "nema dovoljno unosa posle granice za poredjenje"
    if stopa_posle is not None:
        visak = len(parovi) - stopa_posle * n_pre
        procena = "posle granice %d/%d (%.1f%%); procena gresaka ~%d" % (
            p_posle, n_posle, 100 * stopa_posle, max(0, round(visak)))
        if len(parovi) >= 2 and stopa_pre > 1.5 * stopa_posle:
            oz = SUMNJA
        else:
            # Isto ponasanje kao posle ispravke forme: uzastopni unosi istog kooperanta su
            # njegove redovne isporuke, ne zaostao combo -- nema nalaza (ni zbirnog).
            return
    ctx.nalaz(K_KOOP, INFO, OTK, "", (
        "Pre granice %s (%s): %d uzastopnih unosa istog kooperanta i stanice sa razlicitim brojem lista na %d unosa "
        "(%.1f%%); %s. Do v2.7.0 ClearOtkupFields ne prazni cmbKooperant (v2.2.3 frmOtkup.frm:962-982)." % (
            _dmy(granica), izvor, len(parovi), n_pre, 100 * stopa_pre, procena)),
        "Parovi su pojedinacno navedeni; operater ih uporedjuje sa papirnim otkupnim listovima.")
    nov = ctx.nov_po_otkupu
    sab = _Sabirac(pg, K_KOOP, OTK, "OtkupID")
    for r1, r2, t1, t2 in parovi:
        razmak = " (razmak %d min)" % ((t2 - t1).total_seconds() // 60) if t1 is not None and t2 is not None else ""
        novac = aktivni(nov.get(k(r2, "OtkupID"), []))
        nov_txt = ""
        if novac or n(r2, "Novac") > 0:
            nov_txt = " Uz %s je isplaceno %s din (%s) istom kooperantu." % (
                s(r2, "OtkupID"), kg_txt(_suma(novac, "Isplata") or n(r2, "Novac")) or "0",
                _ids(novac, "NovacID") or "tblOtkup.Novac")
        sab.dodaj([r2], oz, (
            "Kandidat: %s pa %s pod istim kooperantom %s%s; do v2.7.0 kooperant ostaje izabran posle snimanja.%s" % (
                _o_txt(r1), _o_txt(r2), s(r2, "KooperantID"), razmak, nov_txt)),
            "Uporediti list %s sa papirom: ako je drugog kooperanta -> KooperantID: %s -> <kooperant sa lista> na %s "
            "(i tblNovac/tblAmbalaza tog otkupa)." % (s(r2, "BrojDokumenta") or s(r2, "OtkupID"),
                                                      s(r2, "KooperantID"), s(r2, "OtkupID")))
    sab.emituj()


# ---------------------------------------------------------------- OTK_CENA_TIP_ZAOSTALI
def _cenovnik(pg):
    """(vrsta, sorta, klasa) -> redovi tblCenovnik po Datum (pa po redu tabele), ili None."""
    cen = pg.ctx.tab(CEN)
    if not cen or not pg.ima(K_CENA, {CEN: ["Datum", "VrstaVoca", "SortaVoca", "Klasa", "Cena"]}):
        return None
    idx = defaultdict(list)
    for c in cen:
        if storno(c) or d(c, "Datum") is None or not ima_broj(c, "Cena"):
            continue
        idx[(k(c, "VrstaVoca"), k(c, "SortaVoca"), k(c, "Klasa"))].append(c)
    for lst in idx.values():
        lst.sort(key=lambda c: (d(c, "Datum"), c.get("_redni") or 0))
    return idx


def _cena(pg):
    if not pg.ima(K_CENA, {OTK: ["OtkupID", "Datum", "VrstaVoca", "SortaVoca", "Klasa", "Cena", "TipAmbalaze",
                                 "KolAmbalaze", "OtpremnicaID", "Stornirano"],
                           OTP: ["OtpremnicaID", "TipAmbalaze", "Stornirano"]}):
        return
    sab = _Sabirac(pg, K_CENA, OTK, "OtkupID")
    idx = _cenovnik(pg)
    # (a) NAJNOVIJA: GetVazecaCena vraca najnoviju cenu bez obzira na datum (modCenovnik.bas:31-80)
    if idx:
        for o in pg.otk_akt:
            oid = s(o, "OtpremnicaID")
            P = pg.ziva(oid)
            if oid and (P is None or not (pg.je_lanac(o, P) or pg.sync_auto(P))):
                continue        # vezan blok: cena je "cena po otpremnici", ne cenovnik
            lst = idx.get((k(o, "VrstaVoca"), k(o, "SortaVoca"), k(o, "Klasa")))
            dan = d(o, "Datum")
            if not lst or dan is None or not ima_broj(o, "Cena"):
                continue
            c_o = n(o, "Cena")
            vaz = [c for c in lst if d(c, "Datum") <= dan]
            ocek = vaz[-1] if vaz else None
            if ocek is not None and abs(n(ocek, "Cena") - c_o) <= CENA_EPS:
                continue
            pog = [c for c in lst if d(c, "Datum") > dan and abs(n(c, "Cena") - c_o) <= CENA_EPS]
            if not pog:
                continue
            opis = ("Otkup %s: NAJNOVIJA -- cena %s je cena iz cenovnika od %s, posle datuma otkupa; vazeca na dan "
                    "otkupa je %s." % (_o_txt(o), _cena_txt(c_o), _dmy(d(pog[0], "Datum")),
                                       "%s (od %s)" % (_cena_txt(n(ocek, "Cena")), _dmy(d(ocek, "Datum")))
                                       if ocek is not None else "nepostojeca"))
            predlog = ("Cena: %s -> %s (vazeca na %s); novac je mozda vec isplacen -- proveriti tblNovac." % (
                _cena_txt(c_o), _cena_txt(n(ocek, "Cena")), _dmy(dan)) if ocek is not None else
                "Na dan otkupa cena nije postojala u cenovniku -- operater utvrdjuje cenu.")
            sab.dodaj([o], SUMNJA, opis, predlog)
    # (b) prvi blok otpremnice storniran: njegova cena je default za ostale (BuildFirstBlokCena bez storno filtera)
    for P in pg.otp_akt:
        prvi = pg.prvi_blok(P)
        if prvi is None or not storno(prvi):
            continue
        c0 = n(prvi, "Cena")
        akt = pg.clanovi(P)
        sa_c0 = [o for o in akt if abs(n(o, "Cena") - c0) <= CENA_EPS]
        ostali = [o for o in akt if abs(n(o, "Cena") - c0) > CENA_EPS]
        if not sa_c0 or not ostali:
            continue
        ostale = sorted({n(o, "Cena").quantize(EPS) for o in ostali})
        for o in sa_c0:
            sab.dodaj([o], SUMNJA, (
                "Otkup %s: cena %s je cena storniranog prvog bloka %s otpremnice %s, a ostali aktivni blokovi imaju %s "
                "(panel uzima cenu prvog bloka bez storno filtera)." % (
                    _o_txt(o), _cena_txt(c0), s(prvi, "OtkupID"), s(P, "OtpremnicaID"),
                    "/".join(_cena_txt(c) for c in ostale))),
                "Cena: %s -> %s (cena ostalih blokova otpremnice); novac je mozda vec isplacen." % (
                    _cena_txt(c0), "/".join(_cena_txt(c) for c in ostale)))
    # (c) tip ambalaze bloka != tip otpremnice (cmbTipAmbalaze se ne prazni)
    for o in pg.otk_akt:
        P = pg.ziva(s(o, "OtpremnicaID"))
        if P is None or n(o, "KolAmbalaze") <= 0:
            continue
        to, tp = k(o, "TipAmbalaze"), k(P, "TipAmbalaze")
        if not to or not tp or to == tp:
            continue
        uzrok = "sync auto-otpremnica uzima tip sa prvog bloka" if pg.sync_auto(P) else \
            "cmbTipAmbalaze se posle snimanja ne prazni"
        sab.dodaj([o], SUMNJA, (
            "Otkup %s: TIP -- %s gajbi tipa %s na otpremnici %s tipa %s (%s)." % (
                _o_txt(o), kg_txt(n(o, "KolAmbalaze")), s(o, "TipAmbalaze"), s(P, "OtpremnicaID"),
                s(P, "TipAmbalaze"), uzrok)),
            "TipAmbalaze: %s -> %s ako su gajbe tog tipa; inace je zbir gajbi otpremnice po tipu pogresan." % (
                s(o, "TipAmbalaze"), s(P, "TipAmbalaze")))
    sab.emituj()


# ---------------------------------------------------------------- ZAOSTALA_CENA_POSLE_DESELEKCIJE
def _cena_deselekcija(pg):
    if not pg.ima(K_CDES, {OTK: ["OtkupID", "OtpremnicaID", "Cena", "VrstaVoca", "SortaVoca", "Stornirano"],
                           OTP: ["OtpremnicaID", "Kolicina", "VrstaVoca", "SortaVoca", "Stornirano"]}):
        return
    idx = _cenovnik(pg) or {}
    poz = {id(o): i for i, o in enumerate(pg.desktop)}
    for P in pg.otp_akt:
        M = [o for o in pg.clanovi(P) if not pg.pwa(o)]
        if not M or pg.preostalo(P) > EPS:
            continue
        L = max(M, key=_kljuc_otk)
        if pg.je_lanac(L, P):
            continue
        i = poz.get(id(L))
        if i is None or i + 1 >= len(pg.desktop):
            continue
        O = pg.desktop[i + 1]            # sledece snimanje posle bloka koji je popunio otpremnicu
        if storno(O) or k(O, "OtpremnicaID") == k(P, "OtpremnicaID"):
            continue
        prvi = pg.prvi_blok(P)
        c = n(prvi, "Cena") if prvi is not None else Decimal(0)
        if c <= 0 or abs(n(O, "Cena") - c) > CENA_EPS:
            continue
        if (k(O, "VrstaVoca"), k(O, "SortaVoca")) == (k(P, "VrstaVoca"), k(P, "SortaVoca")):
            continue
        lst = idx.get((k(O, "VrstaVoca"), k(O, "SortaVoca"), "i"))
        if lst and n(lst[-1], "Cena") > 0:
            continue                     # promena proizvoda okida AutoFillCenaOtkup: GetVazecaCena > 0 prepise
                                         # txtCena, pa cena otpremnice ne moze da zaostane (frmOtkup.frm:398-408)
        pg.nalaz(K_CDES, SUMNJA, OTK, [O], "OtkupID", (
            "Otkup %s (%s/%s) snimljen je odmah posle bloka %s koji je popunio otpremnicu %s (%s/%s), sa njenom cenom "
            "%s: AfterUnos postavi txtCena na cenu otpremnice, pa auto-deselekcija ostavi tu cenu sledecem bloku "
            "(modOtkupBlok.bas:252-267)." % (
                _o_txt(O), s(O, "VrstaVoca"), s(O, "SortaVoca") or "-", s(L, "OtkupID"), _p_txt(P),
                s(P, "VrstaVoca"), s(P, "SortaVoca") or "-", _cena_txt(c))),
            "Proveriti cenu %s za %s/%s sa otkupnim listom i cenovnikom; novac je mozda vec isplacen." % (
                s(O, "OtkupID"), s(O, "VrstaVoca"), s(O, "SortaVoca") or "-"))


# ---------------------------------------------------------------- OTK_BROJ_VAN_NIZA_ILI_PRAZAN / PWA_STANICA_I_DATUM
def _broj_dok(pg):
    ctx = pg.ctx
    if not pg.ima(K_BROJ, {OTK: ["OtkupID", "BrojDokumenta", "StanicaID", "Datum", "Stornirano"]}):
        return
    validacija_ne = kljuc(ctx.config("VALIDACIJA_UNOSA")) in _FLAG_NE
    auto_ne = kljuc(ctx.config("AUTO_BROJ_DOKUMENTA")) in _FLAG_NE
    po_broju = defaultdict(list)
    for dok in pg.otk_dok:
        if norm_broj(polje(dok[0], "BrojDokumenta")):
            po_broju[norm_broj(polje(dok[0], "BrojDokumenta"))].append(dok)
    dupli = {nb: doks for nb, doks in po_broju.items() if len(doks) > 1}
    emitovano = set()
    for dok in pg.otk_dok:
        desk, pwa = _broj_aspekti(pg, dok, validacija_ne, auto_ne)
        o = dok[0]
        if pwa:
            oz = min((z for _t, z, _p in pwa), key=RANG.get)
            pg.nalaz(K_PWA, oz, OTK, dok, "OtkupID", "PWA otkup %s: %s." % (_o_txt(o), "; ".join(t for t, _z, _p in pwa)),
                     " ".join(p for _t, _z, p in pwa if p))
        nb = norm_broj(polje(o, "BrojDokumenta"))
        if nb in dupli:
            if nb in emitovano:
                continue
            emitovano.add(nb)
            doks = dupli[nb]
            redovi = [r for x in doks for r in x]
            ostalo = []
            oz = SUMNJA
            for x in doks:
                dx, _px = _broj_aspekti(pg, x, validacija_ne, auto_ne)
                for t, z, _p in dx:
                    ostalo.append("%s: %s" % (s(x[0], "OtkupID"), t))
                    if RANG[z] < RANG[oz]:
                        oz = z
            kasniji = doks[-1][0]
            opis = ("DUPLIKAT: broj %s na %d aktivna dokumenta: %s. CheckDuplicate blokira samo isti string medju "
                    "aktivnima, pa ostaju razlike u razmacima/velikim slovima i PWA redovi.%s" % (
                        s(o, "BrojDokumenta"), len(doks), "; ".join(_o_txt(x[0]) + " kooperant " +
                                                                    (s(x[0], "KooperantID") or "?") for x in doks),
                        (" Uz to: " + "; ".join(ostalo) + ".") if ostalo else ""))
            predlog = ("Isti blok -> stornirati kasniji (%s); razliciti blokovi -> kasniji dobija sledeci slobodan broj "
                       "%s (prema papirnim listovima)." % (
                           s(kasniji, "OtkupID"), pg.sledeci_broj(s(kasniji, "StanicaID"), d(kasniji, "Datum")) or "?"))
            pg.nalaz(K_BROJ, oz, OTK, redovi, "OtkupID", opis, predlog)
            continue
        if desk:
            oz = min((z for _t, z, _p in desk), key=RANG.get)
            pg.nalaz(K_BROJ, oz, OTK, dok, "OtkupID", "Otkup %s: %s." % (_o_txt(o), "; ".join(t for t, _z, _p in desk)),
                     " ".join(p for _t, _z, p in desk if p))


def _broj_aspekti(pg, dok, validacija_ne, auto_ne):
    """([(tekst, oz, predlog)] desktop, [(tekst, oz, predlog)] PWA) za jedan dokument."""
    o = dok[0]
    broj, st, dan = s(o, "BrojDokumenta"), s(o, "StanicaID"), d(o, "Datum")
    je_pwa = any(pg.pwa(r) for r in dok)
    kanon = pg.sledeci_broj(st, dan)
    if not broj:
        du = pg.ts(o)[1] or dan
        oz = INFO if validacija_ne or (du is not None and du < V283) else SUMNJA
        t = "PRAZAN broj dokumenta%s" % (" (pre v2.8.3 broj nije bio obavezan)" if du is not None and du < V283 else
                                         " (VALIDACIJA_UNOSA = NE)" if validacija_ne else "")
        return [(t, oz, "BrojDokumenta: upisati broj sa papirnog lista (kanon za stanicu i dan: %s)." % (
            kanon or "?"))], []
    m = _KANON.match(broj)
    if not m:
        p = parse_broj(broj)
        if p is not None and not p.s and p.dan is not None:
            ocisceno = "%d/%s%s" % (p.x, p.dan.strftime("%d%m%y"), "-%d" % p.seq if p.seq > 1 else "")
            return [("NEKANONSKI broj '%s' (razmaci)" % broj, INFO, "BrojDokumenta: '%s' -> %s." % (broj, ocisceno))], []
        return [("NEKANONSKI broj '%s' nije oblik x/ddmmyy[-n]%s" % (
            broj, " (AUTO_BROJ_DOKUMENTA = NE: slobodan oblik je dozvoljen)" if auto_ne else ""), INFO, "")], []
    x = int(m.group(1))
    try:
        bdan = date(2000 + int(m.group(4)), int(m.group(3)), int(m.group(2)))
    except ValueError:
        return [("NEKANONSKI broj '%s': dan u broju ne postoji" % broj, SUMNJA,
                 "BrojDokumenta: proveriti sa papirnim listom (kanon: %s)." % (kanon or "?"))], []
    xs = cifre(st)
    desk, pwa = [], []
    if xs is not None and x != xs:
        if je_pwa:
            st_x = pg.ctx.stanica_po_broju.get(x)
            pwa.append(("STANICA_KOOPERANTA -- broj %s je sa uredjaja %d, a StanicaID %s je maticni OM kooperanta "
                        "(ImportRowToTblOtkup uzima stanicu iz tblKooperanti)" % (broj, x, st),
                        SUMNJA if st_x else INFO,
                        "Ako je otkup bio na %s: StanicaID %s -> %s (i noge ambalaze)." % (st_x, st, st_x) if st_x else ""))
        else:
            P = pg.ziva(s(o, "OtpremnicaID"))
            if not (P is not None and k(P, "StanicaID") != k(o, "StanicaID") and cifre(s(P, "StanicaID")) == x):
                desk.append(("TUDJA_STANICA -- broj %s je stanice %d, a otkup je na %s" % (broj, x, st), SUMNJA,
                             "Proveriti da li je pogresna stanica (OM promenjen posle predloga broja) ili broj; ne "
                             "prenumerisati automatski (kanon: %s)." % (kanon or "?")))
    if dan is not None and bdan != dan:
        raz = (dan - bdan).days
        if je_pwa:
            pwa.append(("DATUM_FALLBACK -- dan u broju %s je %s, a Datum %s%s" % (
                broj, _dmy(bdan), _dmy(dan), " (oko ponoci: PWA broj nosi dan uredjaja)" if abs(raz) == 1 else
                " (neprocitan datum PWA uvoz zamenjuje danasnjim)"), INFO if abs(raz) == 1 else SUMNJA,
                "Datum: %s -> %s (dan iz broja), ako papirni list to potvrdi." % (_dmy(dan), _dmy(bdan))))
        else:
            desk.append(("TUDJI_DAN -- dan u broju %s je %s, a Datum %s (predlog broja se ne osvezava kad se datum "
                         "promeni)" % (broj, _dmy(bdan), _dmy(dan)), SUMNJA,
                         "Uporediti Datum i broj sa papirnim listom; ne prenumerisati automatski."))
    return desk, pwa


# ---------------------------------------------------------------- OTK_DUPLI_UNOS_BLOKA
def _dupli(pg):
    if not pg.ima(K_DUPLI, {OTK: ["OtkupID", "KooperantID", "StanicaID", "Datum", "VrstaVoca", "SortaVoca", "Klasa",
                                  "Kolicina", "KolAmbalaze", "OtpremnicaID", "Stornirano"]}):
        return
    nov = pg.ctx.nov_po_otkupu
    g = defaultdict(list)
    for o in pg.otk_akt:
        dan = d(o, "Datum")
        if dan is None or n(o, "Kolicina") <= 0 or not k(o, "KooperantID"):
            continue
        g[(k(o, "KooperantID"), k(o, "StanicaID"), dan, k(o, "VrstaVoca"), k(o, "SortaVoca"), k(o, "Klasa"),
           n(o, "Kolicina").quantize(EPS), n(o, "KolAmbalaze"))].append(o)
    for rows in g.values():
        if len({_kljuc_dok(o) for o in rows}) < 2:
            continue
        rows.sort(key=_kljuc_otk)
        kg = n(rows[0], "Kolicina")
        dokazi, greska = [], False
        tt = [pg.ts(o)[0] for o in rows]
        if all(t is not None for t in tt) and max(tt) - min(tt) <= DUPLI_RAZMAK:
            dokazi.append("uneti u razmaku od %d min" % ((max(tt) - min(tt)).total_seconds() // 60))
        veze = [(o, pg.ziva(s(o, "OtpremnicaID"))) for o in rows]
        vezani = [(o, P) for o, P in veze if P is not None and not pg.sync_auto(P)]
        slobodni = [(o, P) for o, P in veze if P is None or pg.sync_auto(P)]
        if vezani and slobodni:
            dokazi.append("%s je na otpremnici %s, a %s %s" % (
                s(vezani[0][0], "OtkupID"), s(vezani[0][1], "OtpremnicaID"), _ids([o for o, _P in slobodni], "OtkupID"),
                "nije vezan / na sync auto-otpremnici" ))
        po_p = defaultdict(list)
        for o, P in vezani:
            po_p[k(P, "OtpremnicaID")].append((o, P))
        for lst in po_p.values():
            if len(lst) >= 2 and abs(pg.preostalo(lst[0][1]) + kg) <= EPS:
                dokazi.append("oba su na otpremnici %s koja je prepunjena tacno za jedan blok (preostalo %s kg)" % (
                    s(lst[0][1], "OtpremnicaID"), kg_txt(pg.preostalo(lst[0][1]))))
                greska = greska or kg != kg.to_integral_value()
        placeni = [o for o in rows if aktivni(nov.get(k(o, "OtkupID"), []))]
        if len(placeni) >= 2:
            dokazi.append("oba imaju isplatu u tblNovac (%s): rizik dvostruke isplate" % _ids(
                [x for o in placeni for x in aktivni(nov.get(k(o, "OtkupID"), []))], "NovacID"))
        if not dokazi:
            continue
        kandidat = slobodni[0][0] if slobodni else rows[-1]
        pg.nalaz(K_DUPLI, GRESKA if greska else SUMNJA, OTK, rows, "OtkupID", (
            "Isti blok %d puta: kooperant %s, %s, %s, %s/%s kl. %s, %s kg, %s gajbi (%s). Dokaz: %s." % (
                len(rows), s(rows[0], "KooperantID"), s(rows[0], "StanicaID") or "?", _dmy(d(rows[0], "Datum")),
                s(rows[0], "VrstaVoca"), s(rows[0], "SortaVoca") or "-", s(rows[0], "Klasa") or "?", kg_txt(kg),
                kg_txt(n(rows[0], "KolAmbalaze")) or "0",
                ", ".join("%s %s" % (s(o, "OtkupID"), s(o, "BrojDokumenta") or "bez broja") for o in rows),
                "; ".join(dokazi))),
            "Kandidat za storno: %s (%s); posle toga proveriti preostalo otpremnice i novac." % (
                s(kandidat, "OtkupID"), "nevezan" if slobodni else "kasniji unos"))


# ---------------------------------------------------------------- OTK_KULTURAID_PO_VRSTI
def _kultura(pg):
    kul = pg.ctx.tab(KUL)
    if not kul or not pg.ima(K_KUL, {KUL: ["KulturaID", "VrstaVoca", "SortaVoca"],
                                     OTK: ["OtkupID", "KulturaID", "VrstaVoca", "SortaVoca", "Stornirano"]}):
        return
    po_id, po_vs = {}, {}
    for r in kul:
        if k(r, "KulturaID"):
            po_id.setdefault(k(r, "KulturaID"), r)
            po_vs.setdefault((k(r, "VrstaVoca"), k(r, "SortaVoca")), s(r, "KulturaID"))
    g = defaultdict(list)
    for o in pg.otk_akt:
        kid, vr, so = k(o, "KulturaID"), k(o, "VrstaVoca"), k(o, "SortaVoca")
        if not kid or not vr:
            continue
        K = po_id.get(kid)
        if K is not None and k(K, "VrstaVoca") == vr and (not so or k(K, "SortaVoca") == so):
            continue
        g[(kid, vr, so)].append(o)
    for (kid, vr, so), rows in g.items():
        o0 = rows[0]
        K = po_id.get(kid)
        tacan = po_vs.get((vr, so)) if so else None
        if K is not None:
            stanje = "u tblKulture je %s/%s" % (s(K, "VrstaVoca"), s(K, "SortaVoca") or "-")
        else:
            stanje = "nema ga u tblKulture%s" % (" (fallback 'vrsta-sorta')" if "-" in s(o0, "KulturaID") else "")
        oz = GRESKA if tacan and kljuc(tacan) != kid else SUMNJA
        predlog = ("KulturaID: %s -> %s (red tblKulture za %s/%s), kao main RazresiKulturuIzVrsteSorte." % (
            s(o0, "KulturaID"), tacan, s(o0, "VrstaVoca"), s(o0, "SortaVoca")) if tacan else
            "U tblKulture nema reda za %s/%s: dodati ga, pa KulturaID na njegov ID." % (
                s(o0, "VrstaVoca"), s(o0, "SortaVoca") or "-"))
        pg.nalaz(K_KUL, oz, OTK, rows, "OtkupID", (
            "%d aktivnih otkupa %s/%s nosi KulturaID %s (%s): SaveOtkup i PWA uvoz traze KulturaID samo po VrstaVoca "
            "(prvi red vrste), a bez pogotka upisuju 'vrsta-sorta' (modOtkup.bas:462-467)." % (
                len(rows), s(o0, "VrstaVoca"), s(o0, "SortaVoca") or "-", s(o0, "KulturaID"), stanje)), predlog)


# ---------------------------------------------------------------- KOOP_DUPLIKAT_AUTOKREIRANJEM
def _koop_dup(pg):
    ctx = pg.ctx
    kp = ctx.t[KOOP]
    if not kp or not pg.ima(K_KDUP, {KOOP: ["KooperantID", "Ime", "Prezime", "StanicaID"]}):
        return
    ima_bpg, ima_tr = ctx.ima_kolonu(KOOP, "BPGBroj"), ctx.ima_kolonu(KOOP, "TekuciRacun")
    po_imenu = defaultdict(list)
    for r in kp:
        kl = _ime_kljuc(r)
        if kl and k(r, "KooperantID"):
            po_imenu[kl].append(r)
    broj_otk = Counter(k(o, "KooperantID") for o in pg.otk_akt)
    for kl, rr in po_imenu.items():
        if len({k(r, "KooperantID") for r in rr}) < 2:
            continue
        if ima_bpg and all(s(r, "BPGBroj") for r in rr) and len({k(r, "BPGBroj") for r in rr}) == len(rr):
            continue                      # razliciti BPG: razlicite osobe
        stavke, auto = [], []
        for r in rr:
            bez = (not ima_bpg or not s(r, "BPGBroj")) and (not ima_tr or not s(r, "TekuciRacun"))
            if bez:
                auto.append(r)
            stavke.append("%s (%s, %s%s, %d otkupa)" % (
                s(r, "KooperantID"), s(r, "StanicaID") or "?", "bez BPG/racuna -- otisak auto-kreiranja" if bez
                else "BPG %s" % (s(r, "BPGBroj") or "-"), ", neaktivan" if k(r, "Aktivan") in _NEAKTIVAN else "",
                broj_otk[k(r, "KooperantID")]))
        glavni = max(rr, key=lambda r: (bool(s(r, "BPGBroj")), bool(s(r, "TekuciRacun")), broj_otk[k(r, "KooperantID")]))
        pg.nalaz(K_KDUP, SUMNJA, KOOP, rr, "KooperantID", (
            "Isto ime '%s %s' na %d kooperanata: %s. FindKooperantIDByName poredi tacan tekst imena, bez filtera "
            "Aktivan, a KoopAutoCreate (podrazumevano DA) pravi novi red za ukucano ime." % (
                s(rr[0], "Ime"), s(rr[0], "Prezime"), len(rr), "; ".join(stavke))),
            "Ako je ista osoba: spojiti na %s uz preusmeravanje otkupa, novca i ambalaze -- samo odlukom operatera; "
            "ako su razlicite osobe, upisati BPG/JMBG da se razlikuju." % s(glavni, "KooperantID"))
    # imenjak: otkup na stanici S pripisan kooperantu druge stanice, a imenjak postoji na S
    po_st = {}
    for kl, rr in po_imenu.items():
        for r in rr:
            po_st.setdefault((kl, k(r, "StanicaID")), r)
    if not pg.otk or not ctx.ima_kolonu(OTK, "KooperantID") or not ctx.ima_kolonu(OTK, "StanicaID"):
        return
    g = defaultdict(list)
    for o in pg.otk_akt:
        if pg.pwa(o):
            continue
        K = pg.koop.get(k(o, "KooperantID"))
        st = k(o, "StanicaID")
        if K is None or not st or k(K, "StanicaID") == st:
            continue
        im = po_st.get((_ime_kljuc(K), st))
        if im is None or k(im, "KooperantID") == k(K, "KooperantID"):
            continue
        g[(k(K, "KooperantID"), st, k(im, "KooperantID"))].append(o)
    for (_kid, _st, _iid), rows in g.items():
        K = pg.koop[k(rows[0], "KooperantID")]
        im = po_st[(_ime_kljuc(K), _st)]
        pg.nalaz(K_KDUP, SUMNJA, OTK, rows, "OtkupID", (
            "Otkupi %s na stanici %s pripisani su kooperantu %s (maticna stanica %s), a imenjak %s je kooperant te "
            "stanice." % (_ids(rows, "OtkupID"), s(rows[0], "StanicaID"), s(K, "KooperantID"),
                          s(K, "StanicaID") or "?", s(im, "KooperantID"))),
            "Ako su to otkupi imenjaka: KooperantID: %s -> %s (i tblNovac, tblAmbalaza) -- prema papirnim listovima." % (
                s(K, "KooperantID"), s(im, "KooperantID")))


# ---------------------------------------------------------------- OTK_KOOPERANT_NEPOSTOJECI_ID
def _koop_nepostojeci(pg):
    ctx = pg.ctx
    if not ctx.t[KOOP] or not pg.ima(K_KNEP, {OTK: ["OtkupID", "KooperantID", "Stornirano"], KOOP: ["KooperantID"]}):
        return
    g = defaultdict(list)
    for o in pg.otk_akt:
        if k(o, "KooperantID") not in pg.koop:
            g[k(o, "KooperantID")].append(o)
    if not g:
        return
    br_nov, br_amb = Counter(), Counter()          # jedan prolaz (ne po grupi: ~5000 redova tblAmbalaza)
    for r in ctx.t[NOV]:
        if not storno(r):
            for x in {k(r, "KooperantID"), k(r, "PartnerID")} - {""}:
                br_nov[x] += 1
    for r in ctx.t[AMB]:
        if not storno(r) and k(r, "EntitetID"):
            br_amb[k(r, "EntitetID")] += 1
    for kid, rows in g.items():
        v = s(rows[0], "KooperantID")
        nov = br_nov[kid] if kid else 0
        amb = br_amb[kid] if kid else 0
        if not kid:
            opis = "%d aktivnih otkupa nema KooperantID (%s)." % (len(rows), _ids(rows, "OtkupID"))
        else:
            opis = "KooperantID '%s' ne postoji u tblKooperanti (%d aktivnih otkupa)%s." % (
                v, len(rows), ": verovatno tekst iz zagrade ukucanog imena (GetComboID uzima sadrzaj zagrade, "
                "ResolveKooperantByName ga ne proverava)" if not any(ch.isdigit() for ch in v) else "")
        pg.nalaz(K_KNEP, GRESKA, OTK, rows, "OtkupID", opis,
                 "Utvrditi kooperanta sa papirnih listova (%s) i postaviti pravi KooperantID; isto u tblNovac (%d) i "
                 "tblAmbalaza (%d redova)." % (_lista(s(o, "BrojDokumenta") or s(o, "OtkupID") for o in rows),
                                               nov, amb))


# ---------------------------------------------------------------- OTK_GAJBE_KG_NEPROVERENO
def _gajbe(pg):
    ctx = pg.ctx
    if not pg.ima(K_GAJ, {OTK: ["OtkupID", "Kolicina", "KolAmbalaze", "VrstaVoca", "TipAmbalaze", "OtpremnicaID",
                                "Stornirano"]}):
        return
    sab = _Sabirac(pg, K_GAJ, OTK, "OtkupID")
    ima_b = ctx.ima_kolonu(OTK, "BrutoKg")
    uz, uz_b = defaultdict(list), defaultdict(list)
    for o in pg.otk_akt:
        kg, g = n(o, "Kolicina"), n(o, "KolAmbalaze")
        if kg > 0 and g > 0:
            key = (k(o, "VrstaVoca"), k(o, "TipAmbalaze"))
            uz[key].append(kg / g)
            if ima_b and n(o, "BrutoKg") > 0:
                uz_b[key].append(n(o, "BrutoKg") / g)
    med = {key: median(v) for key, v in uz.items() if len(v) >= MIN_UZORAK}
    med_b = {key: median(v) for key, v in uz_b.items() if len(v) >= MIN_UZORAK}
    prag = {}
    if ctx.ima_kolonu(KUL, "PragProsekBlok"):
        for r in ctx.tab(KUL):                 # KulturaProsekPrag: prvi red vrste (LookupValue)
            prag.setdefault(k(r, "VrstaVoca"), n(r, "PragProsekBlok"))
    gajbe_otp = ctx.ima_kolonu(OTP, "KolAmbalaze")
    for dok in pg.otk_dok:
        for o in dok:
            kg, g = n(o, "Kolicina"), n(o, "KolAmbalaze")
            key = (k(o, "VrstaVoca"), k(o, "TipAmbalaze"))
            if kg > 0 and g <= 0:
                P = pg.ziva(s(o, "OtpremnicaID"))
                if P is None or not gajbe_otp:
                    continue           # bez otpremnice nema dokaza: roba moze biti u ambalazi kooperanta
                M = pg.clanovi(P)
                manjak = n(P, "KolAmbalaze") - _suma(M, "KolAmbalaze")
                if manjak <= 0:
                    continue
                bez = [x for x in M if n(x, "KolAmbalaze") <= 0]
                sab.dodaj([o], SUMNJA, (
                    "Otkup %s: BEZ_GAJBI -- %s kg bez gajbi, a otpremnica %s ima %s gajbi naspram %s u blokovima "
                    "(nedostaje %s)." % (_o_txt(o), kg_txt(kg), s(P, "OtpremnicaID"), kg_txt(n(P, "KolAmbalaze")),
                                         kg_txt(_suma(M, "KolAmbalaze")) or "0", kg_txt(manjak))),
                    "KolAmbalaze: 0 -> %s (prema otpremnici; potvrditi otkupnim listom)." % kg_txt(manjak)
                    if len(bez) == 1 else
                    "Manjak %s gajbi rasporediti na blokove bez gajbi (%s) prema otkupnim listovima." % (
                        kg_txt(manjak), _ids(bez, "OtkupID")))
                continue
            if kg <= 0 or g <= 0:
                continue
            p = kg / g
            m = med.get(key)
            van = []
            if m:
                if p > GORNJI * m:
                    van.append("%s kg po gajbi, a medijana za %s/%s je %s" % (
                        kg_txt(p), s(o, "VrstaVoca"), s(o, "TipAmbalaze") or "-", kg_txt(m)))
                elif p < DONJI * m:
                    van.append("samo %s kg po gajbi, a medijana za %s/%s je %s" % (
                        kg_txt(p), s(o, "VrstaVoca"), s(o, "TipAmbalaze") or "-", kg_txt(m)))
                if van and ima_b and n(o, "BrutoKg") > 0 and key in med_b:
                    mb = med_b[key]
                    if DONJI * mb <= n(o, "BrutoKg") / g <= GORNJI * mb:
                        van = []       # bruto po gajbi je u opsegu: neto posle tare, ne kucna greska
            pr = prag.get(k(o, "VrstaVoca"))
            if pr and pr > 0 and p > pr:
                van.append("prosek %s kg po gajbi je iznad PragProsekBlok %s" % (kg_txt(p), kg_txt(pr)))
            if not van:
                continue
            zamena = m and g > kg and Decimal("0.6") * m <= g / kg <= Decimal("1.4") * m
            predlog = ("Zamenjena polja? Kolicina %s <-> KolAmbalaze %s (proveriti otkupnim listom)." % (
                kg_txt(kg), kg_txt(g)) if zamena else
                "Proveriti otkupnim listom%s i zbirom gajbi otpremnice." % (
                    " (gajbi ~%d po medijani)" % int((kg / m).to_integral_value()) if m else ""))
            sab.dodaj([o], SUMNJA, "Otkup %s: %s." % (_o_txt(o), "; ".join(van)), predlog)
    sab.emituj()


# ---------------------------------------------------------------- OTK_AKTIVAN_BEZ_ZIVE_OTPREMNICE
def _ziva(pg):
    ctx = pg.ctx
    if not pg.ima(K_ZIVA, {OTK: ["OtkupID", "OtpremnicaID", "StanicaID", "Datum", "Kolicina", "Stornirano"],
                           OTP: ["OtpremnicaID", "Stornirano"]}):
        return
    sv = ctx.t[SV]
    nasl = defaultdict(list)
    for R in sv:
        if k(R, "OldDocID") and s(R, "NewDocID"):
            nasl[k(R, "OldDocID")].append(s(R, "NewDocID"))
    vezani = defaultdict(list)
    for o in pg.otk_akt:
        if pg.ziva(s(o, "OtpremnicaID")) is not None:
            vezani[(k(o, "KooperantID"), d(o, "Datum"))].append(o)
    for dok in pg.otk_dok:
        mrtvi = [o for o in dok if s(o, "OtpremnicaID") and pg.ziva(s(o, "OtpremnicaID")) is None
                 and id(o) not in pg.panel_b]
        if not mrtvi:
            continue
        delovi, predlozi = [], []
        for o in mrtvi:
            oid = s(o, "OtpremnicaID")
            stare = pg.otp_svi_po_id.get(kljuc(oid), [])
            if stare:
                delovi.append("%s: otpremnica %s je stornirana%s" % (
                    s(o, "OtkupID"), _p_txt(stare[0]),
                    " (zbirna %s)" % s(stare[0], "BrojZbirne") if s(stare[0], "BrojZbirne") else ""))
            else:
                delovi.append("%s: otpremnica %s ne postoji u tblOtpremnica" % (s(o, "OtkupID"), oid))
            nn = [x for x in nasl.get(kljuc(oid), []) if pg.ziva(x) is not None]
            strogo = [X for X in pg.otp_akt_po_broju.get(k(o, "BrojDokumenta"), [])
                      if k(X, "StanicaID") == k(o, "StanicaID") and d(X, "Datum") == d(o, "Datum")
                      and all(k(X, c) == k(o, c) for c in ("VrstaVoca", "SortaVoca", "Klasa", "TipAmbalaze"))
                      and abs(n(X, "Kolicina") - n(o, "Kolicina")) <= EPS
                      and abs(n(X, "KolAmbalaze") - n(o, "KolAmbalaze")) <= EPS] if k(o, "BrojDokumenta") else []
            ponovo = [c for c in vezani.get((k(o, "KooperantID"), d(o, "Datum")), [])
                      if c is not o and abs(n(c, "Kolicina") - n(o, "Kolicina")) <= EPS]
            if ponovo:
                predlozi.append("Isti kooperant, dan i kg ponovo unet i vezan (%s): %s je verovatno visak -> storno." % (
                    _ids(ponovo, "OtkupID"), s(o, "OtkupID")))
            elif nn:
                predlozi.append("OtpremnicaID: %s -> %s (naslednik iz tblStornoVeze)." % (oid, nn[0]))
            elif strogo:
                predlozi.append("OtpremnicaID: %s -> %s (isti broj, dan, stanica, roba, kg i gajbe)." % (
                    oid, s(strogo[0], "OtpremnicaID")))
            else:
                kand = pg.kandidati_otp(o, d(o, "Datum"), tol_dana=1)
                predlozi.append("Kandidati (ista stanica, klasa, vrsta, dan +-1, dovoljno preostalog kg): %s." % (
                    pg.kandidati_txt(kand)) if kand else "Kandidata nema -- odluka operatera (storno bloka ili nova "
                                                         "otpremnica).")
        pg.nalaz(K_ZIVA, GRESKA, OTK, mrtvi, "OtkupID", "Aktivan otkup %s pokazuje na otpremnicu koje nema medju "
                 "aktivnima (%s); blok tako nije ni u jednoj zivoj otpremnici." % (
                     _o_txt(mrtvi[0]), "; ".join(delovi)), " ".join(predlozi))
    # (b) blok oslobodjen stornom (DUPLI_FANTOM / PONISTENJE_BEZ_ZAMENE): kandidat, ne dokaz
    if not sv or not all(ctx.ima_kolonu(SV, c) for c in ("Mode", "OldDocType", "OldBroj")):
        return
    dog = defaultdict(list)
    for R in sv:
        mode, tip = k(R, "Mode"), k(R, "OldDocType")
        if mode not in ("dupli_fantom", "ponistenje_bez_zamene"):
            continue
        if tip == "otpremnica":
            S = ctx.otp_po_broju.get(k(R, "OldBroj"), [])
        elif tip == "zbirna":
            S = ctx.otp_po_broju_zbirne.get(k(R, "OldBroj"), [])
        elif tip == "prijemnica":
            S = ctx.otp_po_broju_zbirne.get(k(R, "ParentBroj"), [])
        else:
            continue
        for x in S:
            if storno(x):
                dog[(k(x, "StanicaID"), d(x, "Datum"))].append((R, x))
    if not dog:
        return
    for dok in pg.otk_dok:
        if any(s(o, "OtpremnicaID") or s(o, "BrojZbirne") for o in dok):
            continue
        o = dok[0]
        ev = dog.get((k(o, "StanicaID"), d(o, "Datum")))
        if not ev:
            continue
        kand = pg.kandidati_otp(o, d(o, "Datum"), tol_dana=1)
        pg.nalaz(K_ZIVA, SUMNJA, OTK, dok, "OtkupID", (
            "Nevezan aktivan otkup %s (bez OtpremnicaID i BrojZbirne): iste stanice i dana je storno %s -- storno "
            "oslobadja blokove bez traga po bloku, pa je ovo kandidat, ne dokaz." % (
                _o_txt(o), _lista("%s %s (%s)" % (s(R, "Mode"), s(x, "BrojOtpremnice") or s(R, "OldBroj"),
                                                  s(x, "OtpremnicaID")) for R, x in ev))),
            ("Ako je blok bio na storniranoj otpremnici: vezati za naslednika; kandidati: %s." % pg.kandidati_txt(kand))
            if kand else "Ako je blok bio na storniranoj otpremnici: vezati za naslednika ili novu otpremnicu.")


# ---------------------------------------------------------------- SYNC_AUTO_OTPREMNICA_ZA_NEVEZANE
def _sync(pg):
    if not pg.ima(K_SYNC, {OTP: ["OtpremnicaID", "StanicaID", "Datum", "VozacID", "BrojOtpremnice", "BrojZbirne",
                                 "Kolicina", "Klasa", "Stornirano"],
                           OTK: ["OtkupID", "OtpremnicaID", "StanicaID", "Datum", "VozacID", "Klasa", "Kolicina",
                                 "Stornirano"]}):
        return
    stornirane = defaultdict(list)
    for x in pg.otp:
        if storno(x):
            stornirane[(k(x, "StanicaID"), d(x, "Datum"), k(x, "Klasa"))].append(x)
    for A in sorted(pg.otp_akt, key=lambda r: _num(s(r, "OtpremnicaID"))):
        st, dan = k(A, "StanicaID"), d(A, "Datum")
        if pg.ima_jehl and st in pg.hladnjace:
            continue
        M = pg.clanovi(A)
        if not M or any(pg.pwa(o) for o in M) or not pg.sync_like(A):
            continue
        bz, broj = k(A, "BrojZbirne"), s(A, "BrojOtpremnice")
        if not bz or not pg.zbr_akt_po_b.get(bz) or norm_broj(bz) not in (norm_broj("S" + broj), norm_broj(broj)):
            continue
        if pg.prj_akt_po_b.get(bz):
            continue
        if any(abs(n(x, "Kolicina") - n(A, "Kolicina")) <= EPS and _num(s(x, "OtpremnicaID")) < _num(s(A, "OtpremnicaID"))
               for x in stornirane[(st, dan, k(A, "Klasa"))]):
            continue                # oslobodjeni blokovi posle storna: AUTO_USKRSNUCE_POSLE_STORNA (storno)
        Ys = [Y for Y in pg.otp_akt_po_st_dan.get((st, dan), [])
              if Y is not A and (not k(Y, "Klasa") or k(Y, "Klasa") == k(A, "Klasa")) and pg.preostalo(Y) > EPS]
        if not Ys:
            continue
        Ys.sort(key=lambda Y: (abs(pg.preostalo(Y) - n(A, "Kolicina")), _num(s(Y, "OtpremnicaID"))))
        jak = abs(pg.preostalo(Ys[0]) - n(A, "Kolicina")) <= EPS
        proizvodi = sorted({"%s/%s" % (s(o, "VrstaVoca"), s(o, "SortaVoca") or "-") for o in M})
        pg.nalaz(K_SYNC, SUMNJA, OTP, [A], "OtpremnicaID", (
            "Otpremnica %s ima otisak sync auto-otpremnice: clanovi %s su desktop blokovi (zbir %s kg = otpremnica; "
            "vrsta, sorta i cena sa prvog bloka)%s, a zbirna %s nema prijemnicu. Ista stanica i dan: %s%s. "
            "AutoCreateOtpremniceFromPWA pakuje SVAKI nevezan otkup sa VozacID-jem, i desktop." % (
                _p_txt(A), _ids(M, "OtkupID"), kg_txt(_suma(M, "Kolicina")),
                "; vise proizvoda: %s" % ", ".join(proizvodi) if len(proizvodi) > 1 else "", s(A, "BrojZbirne"),
                _lista("%s (preostalo %s kg)" % (_p_txt(Y), kg_txt(pg.preostalo(Y))) for Y in Ys),
                " -- manjak tacno odgovara ovoj otpremnici" if jak else "")),
            "Ako su blokovi fizicki isli na %s: OtpremnicaID %s -> %s, pa stornirati %s i zbirnu %s (bez prijemnice)." % (
                s(Ys[0], "OtpremnicaID"), _ids(M, "OtkupID"), s(Ys[0], "OtpremnicaID"), s(A, "OtpremnicaID"),
                s(A, "BrojZbirne")))


_PROVERE = (_literal, _panel, _dupli_lanac, _zastareli, _veza, _broj_zbirne, _vozac, _koop_zaostao, _cena,
            _cena_deselekcija, _broj_dok, _dupli, _kultura, _koop_dup, _koop_nepostojeci, _gajbe, _ziva, _sync)


def provere(ctx):
    pg = _Pogled(ctx)
    for f in _PROVERE:
        try:
            f(pg)
        except Exception:  # noqa: BLE001 -- pad jedne provere ne sme da obori ostale
            ctx.nalaz("MODUL_PAO", GRESKA, "", "otkup" + f.__name__, traceback.format_exc(limit=6))
