"""Provere OTPREMNICE i ZBIRNE stare sveske (C002) -- SAMO CITA.

Veze u staroj svesci: otpremnica -> zbirna i prijemnica -> zbirna idu PO BROJU
(BrojZbirne je tekst), otkup -> otpremnica po OtpremnicaID. U malina modu
auto-zbirna pravi zbirnu za svaki red otpremnice: BrojZbirne = "S" + BrojOtpremnice
kad je VozacID mirror stanice (ApplyMirrorPrefix, vba-v2.11.0 modBrojevi.bas:296-301),
a Datum, VozacID, kg i gajbe uzima sa otpremnice (modMasterSync.bas:940-976).

Klase (kod nalaza = ID klase iz specifikacije):
  OTP_VOZAC_NIJE_PAR_STANICE      vozac druge stanice na otpremnici/zbirnoj (+ izvor, P7)
  TURA_PO_VOZACU_IZ_OTKUPA        ista stanica/dan/klasa, otpremnice sa razlicitim vozacem (P4)
  OTP_BROJ_DATUM_NESLAGANJE       ddmmyy iz broja != Datum (otpremnica, pa zbirna)
  OTP_BROJ_TUDJA_STANICA          prefiks broja != broj stanice otpremnice
  OTP_DUPLI_BROJ_AKTIVAN          isti broj na vise aktivnih otpremnica (drugi dan/stanica/ista klasa)
  ZBR_DUPLI_BROJ_AKTIVAN          isti broj na vise aktivnih zbirnih (ista klasa, drugi dan ili vozac)
  ZBR_VISI_OTP_STORNIRANE         aktivna zbirna, sve otpremnice stornirane
  ISPRAVKA_ZAVRSENA_BEZ_NOVE_ZBIRNE  isto, 0 kg + prijemnica + naslednik bez zbirne (P5, v2.19+)
  ZBR_AKTIVNA_STORNIRANE_OTPREMNICE  aktivna zbirna sa delom storniranih otpremnica, zbir ne valja
  OTP_CEKA_NOVU_ZBIRNU            aktivna otpremnica, njen broj zbirne postoji samo storniran
  OTP_AKTIVNA_NA_STORNIRANOJ_ZBIRNOJ  isto, ali samo pri tacnom poredjenju (stari kod je case-sensitive)
  OTP_BEZ_ZBIRNE_MALINA           malina, aktivna otpremnica bez BrojZbirne
  HLADNJACA_LANAC_NEPOTPUN        broj zbirne bez zbirne, ili zbirna bez ijedne otpremnice
  ZBR_TURA_BEZ_PRIJEMNICE         zbirna bez prijemnice uz susednu sa dve (varijante a/b)
  ZBR_CLAN_VAN_1_1                malina: otpremnica na zbirnoj koja nije njena "S"+broj
  ZBR_ISTI_BROJ_STORNIRANA_I_AKTIVNA  dupli/fantom storno zbirne, pa ju je auto-zbirna ozivela

Jedan nalaz po dokumentu po klasi: redovi istog dokumenta (Klasa I i II) se
spajaju, a kad jedan uzrok pogadja vise dokumenata u lancu (otpremnica, njena
zbirna, prijemnica), nalaz ide na izvor i nabraja ostale.
"""

import re
import traceback
from collections import Counter, defaultdict
from datetime import date, timedelta
from decimal import Decimal

from citac import kljuc, tekst

from provera.zajednicko import GRESKA, INFO, SUMNJA, cifre, d, k, kg_txt, n, norm_broj, parse_broj, s, storno

OTP, ZBR, PRJ, OTK, SV = "tblOtpremnica", "tblZbirna", "tblPrijemnica", "tblOtkup", "tblStornoVeze"
_MALINA_DA = ("yes", "true", "1", "on", "enabled")      # IsMalinaMode, modConfig.bas:909-920
_DA = ("da", "yes", "true", "1")
_HL = re.compile(r"^\s*HL-(\d{2})(\d{2})(\d{2})-(\d{6})\s*$", re.IGNORECASE)   # modAutoHladnjaca.bas:93
_EPS = Decimal("0.01")
_MAX = 12   # najvise ID-jeva/stavki u jednom opisu

K_VOZAC = "OTP_VOZAC_NIJE_PAR_STANICE"
K_TURA_VOZ = "TURA_PO_VOZACU_IZ_OTKUPA"
K_DATUM = "OTP_BROJ_DATUM_NESLAGANJE"
K_TUDJA = "OTP_BROJ_TUDJA_STANICA"
K_OTP_DUPLI = "OTP_DUPLI_BROJ_AKTIVAN"
K_ZBR_DUPLI = "ZBR_DUPLI_BROJ_AKTIVAN"
K_VISI = "ZBR_VISI_OTP_STORNIRANE"
K_ISPR = "ISPRAVKA_ZAVRSENA_BEZ_NOVE_ZBIRNE"
K_DELIM = "ZBR_AKTIVNA_STORNIRANE_OTPREMNICE"
K_CEKA = "OTP_CEKA_NOVU_ZBIRNU"
K_CEKA_TACNO = "OTP_AKTIVNA_NA_STORNIRANOJ_ZBIRNOJ"
K_BEZ = "OTP_BEZ_ZBIRNE_MALINA"
K_HL = "HLADNJACA_LANAC_NEPOTPUN"
K_TURA = "ZBR_TURA_BEZ_PRIJEMNICE"
K_CLAN = "ZBR_CLAN_VAN_1_1"
K_ISTI = "ZBR_ISTI_BROJ_STORNIRANA_I_AKTIVNA"


# ---------------------------------------------------------------- pomocne
def _dmy(x):
    return x.strftime("%d.%m.%Y") if isinstance(x, date) else "?"


def _lista(vrednosti, maks=_MAX):
    v = list(vrednosti)
    if len(v) <= maks:
        return ", ".join(v)
    return ", ".join(v[:maks]) + " ... (+%d)" % (len(v) - maks)


def _ids(redovi, pk):
    return ", ".join(sorted({s(r, pk) for r in redovi if s(r, pk)}))


def _id_set(redovi, pk):
    return {k(r, pk) for r in redovi if k(r, pk)}


def _kg(redovi, kol):
    return sum((n(r, kol) for r in redovi), Decimal(0))


def _klase(r):
    v = k(r, "Klasa")
    return frozenset(x.strip() for x in v.split("/")) if "/" in v else frozenset([v])


def _klasa_txt(redovi):
    return "/".join(sorted({s(r, "Klasa") for r in redovi if s(r, "Klasa")})) or "?"


def _dan(redovi, kol="Datum"):
    for r in redovi:
        x = d(r, kol)
        if x is not None:
            return x
    return None


def _num(id_):
    c = cifre(id_)
    return c if c is not None else -1


def _broj_za(x, dan, seq):
    b = "%d/%s" % (x, dan.strftime("%d%m%y"))
    return b if seq <= 1 else "%s-%d" % (b, seq)


def _dokumenti(redovi, kljuc_fn):
    """Redovi istog dokumenta (Klasa I i II) zajedno, redosledom prvog ID-ja."""
    g = defaultdict(list)
    for r in redovi:
        g[kljuc_fn(r)].append(r)
    return list(g.values())


def _po(redovi, kol):
    g = defaultdict(list)
    for r in redovi:
        v = k(r, kol)
        if v:
            g[v].append(r)
    return g


class _Pogled:
    """Indeksi koje provere ovog modula dele (racuna se jednom)."""

    def __init__(self, ctx):
        self.ctx = ctx
        self.preskoceno = set()
        self.pokriveno = defaultdict(set)     # kod -> ID-jevi vec prijavljeni (za dedup izmedju klasa)
        self.zbr_dupli = set()                # brojevi zbirnih prijavljeni kao ZBR_DUPLI
        t = ctx.t
        self.otp, self.zbr, self.prj, self.otk = t[OTP], t[ZBR], t[PRJ], t[OTK]
        self.sv = t.get(SV, [])
        self.malina = kljuc(ctx.config("MALINA_MODE")) in _MALINA_DA
        # stanica = red u tblStanice; ako je tabela prazna, bar StanicaID-jevi sa otpremnica
        self.stanice = set(ctx.stanice) | {k(r, "StanicaID") for r in self.otp if k(r, "StanicaID")}
        self.hladnjace = {k(r, "StanicaID") for r in t["tblStanice"] if k(r, "JeHladnjaca") in _DA}
        self.vozaci = {k(r, "VozacID") for r in t["tblVozaci"] if k(r, "VozacID")}
        self.otp_akt = [r for r in self.otp if not storno(r)]
        self.zbr_akt = [r for r in self.zbr if not storno(r)]
        self.prj_akt = [r for r in self.prj if not storno(r)]
        self.zbr_po_b = _po(self.zbr, "BrojZbirne")
        self.zbr_akt_po_b = _po(self.zbr_akt, "BrojZbirne")
        self.otp_po_b = _po(self.otp, "BrojZbirne")
        self.otp_akt_po_b = _po(self.otp_akt, "BrojZbirne")
        self.otp_po_broju = _po(self.otp, "BrojOtpremnice")
        self.prj_akt_po_b = _po(self.prj_akt, "BrojZbirne")
        otk_akt = [r for r in self.otk if not storno(r)]
        self.otk_akt_po_otp = _po(otk_akt, "OtpremnicaID")
        self.otk_akt_po_b = _po(otk_akt, "BrojZbirne")
        # BrojZbirne je deo kljuca: red istog dokumenta sa drugim (ili praznim) brojem zbirne je drugi lanac
        # i mora ga videti svaka provera (inace se gleda samo prvi red dokumenta)
        self.otp_dok = _dokumenti(self.otp_akt, lambda r: (
            k(r, "BrojOtpremnice") or "#" + k(r, "OtpremnicaID"), k(r, "StanicaID"), k(r, "VozacID"), d(r, "Datum"),
            k(r, "BrojZbirne")))
        self.zbr_dok = _dokumenti(self.zbr_akt, lambda r: (
            k(r, "BrojZbirne") or "#" + k(r, "ZbirnaID"), k(r, "VozacID"), d(r, "Datum")))
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

    # --- izvedeno ---------------------------------------------------
    def ocek(self, vozac, broj):
        """ApplyMirrorPrefix: broj koji vec pocinje velikim S ostaje; inace S+broj
        kad je vozac bilo koja stanica, a broj kad je pravi vozac."""
        b = tekst(broj)
        if not b or b.startswith("S"):
            return b
        return "S" + b if kljuc(vozac) in self.stanice else b

    def stanica_zbirne(self, zs):
        """kljuc StanicaID zbirne: iz prefiksa S<x> (broj otpremnice), pa iz VozacID-mirror."""
        for z in zs:
            p = parse_broj(s(z, "BrojZbirne"))
            if p is not None and p.s and p.x in self.ctx.stanica_po_broju:
                return kljuc(self.ctx.stanica_po_broju[p.x])
        for z in zs:
            if k(z, "VozacID") in self.stanice:
                return k(z, "VozacID")
        return ""

    def stanica_prijemnice(self, prows):
        for r in prows:
            if k(r, "VozacID") in self.stanice:
                return k(r, "VozacID")
        return self.stanica_zbirne(prows)

    def otkupi(self, dok):
        return [o for r in dok for o in self.otk_akt_po_otp.get(k(r, "OtpremnicaID"), [])]

    def sledeci_broj(self, stanica_id, dan, rezervisano):
        """Prvi slobodan broj x/ddmmyy[-n] za (stanica, dan), racunajuci i stornirane
        redove (MaxSeqFromTable), a i brojeve vec predlozene u ovom nalazu."""
        x = cifre(stanica_id)
        if x is None or dan is None:
            return ""
        if self._seq is None:
            self._seq = defaultdict(int)
            for r in self.otp:
                p = parse_broj(s(r, "BrojOtpremnice"))
                if p is not None and not p.s and p.dan is not None:
                    self._seq[(p.x, p.dan)] = max(self._seq[(p.x, p.dan)], p.seq)
        seq = self._seq[(x, dan)] + 1
        while True:
            b = _broj_za(x, dan, seq)
            if kljuc(b) not in self.otp_po_broju and kljuc(b) not in rezervisano:
                rezervisano.add(kljuc(b))
                return b
            seq += 1

    def prijemnice_dok(self, prows):
        """Prijemnice na jednom broju zbirne, po dokumentu (BrojPrijemnice)."""
        return _dokumenti(prows, lambda r: k(r, "BrojPrijemnice") or "#" + k(r, "PrijemnicaID"))

    def prijemnice_txt(self, prows, sa_vozacem=False):
        out = []
        for pd in self.prijemnice_dok(prows):
            t = "%s %s %s kg" % (_ids(pd, "PrijemnicaID"), s(pd[0], "BrojPrijemnice"), kg_txt(_kg(pd, "Kolicina")))
            if sa_vozacem:
                t += " VozacID %s" % (s(pd[0], "VozacID") or "-")
            out.append(t)
        return _lista(out)


# ---------------------------------------------------------------- 1) vozac
def _otp_vozac(pg):
    ctx = pg.ctx
    if not pg.ima(K_VOZAC, {OTP: ["OtpremnicaID", "Datum", "StanicaID", "VozacID", "BrojOtpremnice", "BrojZbirne",
                                  "Stornirano"]}):
        return
    ima_zbr = bool(pg.zbr) and pg.ima(K_VOZAC, {ZBR: ["ZbirnaID", "VozacID", "BrojZbirne", "Stornirano"]})
    oz = GRESKA if pg.malina else SUMNJA
    pokr = set()
    pravi = defaultdict(list)
    stornirane = defaultdict(list)          # (stanica, dan) -> stornirane otpremnice (P7)
    for r in pg.otp:
        if storno(r):
            stornirane[(k(r, "StanicaID"), d(r, "Datum"))].append(r)
    for dok in pg.otp_dok:
        r0 = dok[0]
        voz, st = k(r0, "VozacID"), k(r0, "StanicaID")
        if not voz or not st or voz == st:
            continue
        B = k(r0, "BrojZbirne")
        zb = pg.zbr_akt_po_b.get(B, []) if (ima_zbr and B) else []
        pr = pg.prj_akt_po_b.get(B, []) if B else []
        pokr |= _id_set(zb, "ZbirnaID")
        oid = _ids(dok, "OtpremnicaID")
        if voz not in pg.stanice:
            if s(r0, "BrojZbirne")[:1] in ("S", "s"):
                ctx.nalaz(K_VOZAC, SUMNJA, OTP, oid,
                          "Otpremnica %s (%s, %s) ima pravog vozaca %s, a broj zbirne %s nosi S prefiks koji "
                          "ApplyMirrorPrefix daje samo vozacu-stanici." % (
                              s(r0, "BrojOtpremnice"), s(r0, "StanicaID"), _dmy(d(r0, "Datum")),
                              s(r0, "VozacID"), s(r0, "BrojZbirne")),
                          "Proveriti vozaca: ako je tura stanice, VozacID: %s -> %s na otpremnici, zbirnoj i "
                          "prijemnici; ako je pravi vozac, broj zbirne bez S." % (s(r0, "VozacID"), s(r0, "StanicaID")))
            else:
                pravi[s(r0, "VozacID")].append(dok)
            continue
        # vozac je DRUGA stanica (mirror): lanac otpremnica -> zbirna -> prijemnice
        delovi = ["Otpremnica %s (%s, %s) nosi vozaca druge stanice %s" % (
            s(r0, "BrojOtpremnice"), s(r0, "StanicaID"), _dmy(d(r0, "Datum")), s(r0, "VozacID"))]
        if zb:
            delovi.append("zbirna %s (%s) VozacID %s" % (
                s(zb[0], "BrojZbirne"), _ids(zb, "ZbirnaID"), "/".join(sorted({s(z, "VozacID") for z in zb}))))
        if pr:
            delovi.append("prijemnice: %s" % pg.prijemnice_txt(pr, sa_vozacem=True))
        izvor = ""
        st_par = [x for x in stornirane.get((st, d(r0, "Datum")), []) if k(x, "VozacID") == voz]
        otk = pg.otkupi(dok)
        if st_par:
            izvor = "prefill ispravke (v2.22+) preneo vozaca sa stornirane %s" % _ids(st_par, "OtpremnicaID")
        elif any(k(o, "VozacID") == voz for o in otk):
            izvor = "otkup (%s nosi isti VozacID; frmOtkup ne prazni cmbVozac)" % _ids(
                [o for o in otk if k(o, "VozacID") == voz], "OtkupID")
        elif otk and all(k(o, "VozacID") == st for o in otk):
            izvor = "forma otpremnice (cmbVozac zaostao; otkupi %s imaju VozacID = stanica)" % _ids(otk, "OtkupID")
        if izvor:
            delovi.append("izvor: " + izvor)
        svi = [oid, _ids([z for z in zb if k(z, "VozacID") != st], "ZbirnaID"),
               _ids([p for p in pr if k(p, "VozacID") != st], "PrijemnicaID")]
        predlog = "VozacID: %s -> %s na %s" % (s(r0, "VozacID"), s(r0, "StanicaID"), ", ".join(x for x in svi if x))
        if zb:
            predlog += " (broj zbirne %s ostaje)" % s(zb[0], "BrojZbirne")
        predlog += "."
        if pg.vozaci and st not in pg.vozaci:
            predlog += " U tblVozaci fali par-vozac %s -- dodati ga pre ispravke (to je i uzrok)." % s(r0, "StanicaID")
        ctx.nalaz(K_VOZAC, oz, OTP, oid, "; ".join(delovi) + ".", predlog)
    # pravi vozaci: dosledan obrazac (zbirna bez S) -- legitimno ako C002 ima tog vozaca
    for voz, doks in sorted(pravi.items()):
        otp_ids = [_ids(x, "OtpremnicaID") for x in doks]
        if pg.vozaci and kljuc(voz) not in pg.vozaci:
            # lazni pozitiv (a) vazi samo za vozaca koji POSTOJI u tblVozaci; nepoznat ID nije pravi vozac
            ctx.nalaz(K_VOZAC, SUMNJA, "tblVozaci", voz,
                      "VozacID %s na %d otpremnica (%s) ne postoji ni u tblVozaci ni u tblStanice." % (
                          voz, len(doks), _lista(otp_ids)),
                      "VozacID: %s -> StanicaID otpremnice (par-vozac), isto na njihovim zbirnim i prijemnicama; "
                      "ili dodati vozaca u tblVozaci ako stvarno postoji." % voz)
            continue
        ctx.nalaz(K_VOZAC, INFO, "tblVozaci", voz,
                  "Pravi vozac %s na %d otpremnica (%s); zbirne bez S prefiksa -- dosledno." % (
                      voz, len(doks), _lista(otp_ids)),
                  "Legitimno ako C002 stvarno ima ovog vozaca; inace VozacID -> StanicaID otpremnice.")
    if not ima_zbr:
        return
    # zbirne koje nisu vec u lancu neke otpremnice
    for zd in pg.zbr_dok:
        z0 = zd[0]
        if _id_set(zd, "ZbirnaID") & pokr:
            continue
        voz = k(z0, "VozacID")
        broj = s(z0, "BrojZbirne")
        p = parse_broj(broj)
        if not voz or p is None:
            continue
        zid = _ids(zd, "ZbirnaID")
        otps = pg.otp_akt_po_b.get(k(z0, "BrojZbirne"), [])
        if otps and all(k(o, "VozacID") == voz and k(o, "StanicaID") == voz for o in otps):
            continue                      # vozac potvrdjen otpremnicom: nesklad je u broju (TUDJA_STANICA / CLAN_VAN_1_1)
        otp_voz = sorted({s(o, "VozacID") for o in otps if s(o, "VozacID")})
        vlasnik = pg.ctx.stanica_po_broju.get(p.x, "stanica %d" % p.x)
        if p.s and voz in pg.stanice and cifre(voz) != p.x:
            cilj = otp_voz[0] if len(otp_voz) == 1 else vlasnik
            ctx.nalaz(K_VOZAC, oz, ZBR, zid,
                      "Zbirna %s (%s): broj pripada stanici %s, a vozac je %s%s." % (
                          broj, _dmy(d(z0, "Datum")), vlasnik, s(z0, "VozacID"),
                          "; otpremnice %s imaju VozacID %s" % (_ids(otps, "OtpremnicaID"), "/".join(otp_voz))
                          if otps else ""),
                      "VozacID zbirne: %s -> %s (i na njenim prijemnicama)." % (s(z0, "VozacID"), cilj))
        elif p.s and voz not in pg.stanice:
            ctx.nalaz(K_VOZAC, SUMNJA, ZBR, zid,
                      "Zbirna %s (%s) ima S prefiks (rezervisan za vozaca-stanicu), a vozac %s nije stanica." % (
                          broj, _dmy(d(z0, "Datum")), s(z0, "VozacID")),
                      "VozacID: %s -> %s, ili broj bez S ako je pravi vozac." % (s(z0, "VozacID"), vlasnik))
        elif not p.s and voz in pg.stanice and not broj.startswith("S"):
            ctx.nalaz(K_VOZAC, SUMNJA, ZBR, zid,
                      "Zbirna %s (%s) nema S prefiks, a vozac %s je stanica (ApplyMirrorPrefix bi dodao S)." % (
                          broj, _dmy(d(z0, "Datum")), s(z0, "VozacID")),
                      "Proveriti rucni unos broja: ocekivano %s." % ("S" + broj))


# ---------------------------------------------------------------- P4 ture po vozacu
def _tura_po_vozacu(pg):
    ctx = pg.ctx
    if not pg.ima(K_TURA_VOZ, {OTP: ["OtpremnicaID", "Datum", "StanicaID", "VozacID", "BrojOtpremnice", "Klasa",
                                     "Stornirano"]}):
        return
    g = defaultdict(list)
    for r in pg.otp_akt:
        st, dan = k(r, "StanicaID"), d(r, "Datum")
        if st and dan is not None:
            g[(st, dan, k(r, "Klasa"))].append(r)
    for (st, dan, _kl), rows in sorted(g.items(), key=lambda x: (x[0][1], x[0][0], x[0][2])):
        voz = {k(r, "VozacID") for r in rows if k(r, "VozacID")}
        if len(voz) < 2:
            continue
        if not any(v != st and v in pg.stanice for v in voz):
            continue                      # drugi vozac je pravi vozac (VOZ-...): legitimna druga tura, ne P4
        rows = sorted(rows, key=lambda r: _num(s(r, "OtpremnicaID")))
        st_id = s(rows[0], "StanicaID")
        stavke = ["%s %s VozacID %s %s kg" % (s(r, "OtpremnicaID"), s(r, "BrojOtpremnice"), s(r, "VozacID") or "-",
                                             kg_txt(n(r, "Kolicina"))) for r in rows]
        tudji = [r for r in rows if k(r, "VozacID") and k(r, "VozacID") != st]
        otk = pg.otkupi(tudji)
        opis = "Stanica %s, %s, Klasa %s: %d otpremnica sa razlicitim vozacem (%s)." % (
            st_id, _dmy(dan), s(rows[0], "Klasa") or "?", len(rows), _lista(stavke))
        if otk:
            opis += " Otkupi tih tura: %s." % _lista(
                "%s VozacID %s" % (s(o, "OtkupID"), s(o, "VozacID") or "-") for o in otk)
        predlog = ("Auto-otpremnica grupise otkupe i po VozacID (modMasterSync.bas:667-671): ako je vozac iz otkupa "
                   "zaostao, VozacID -> %s na %s i njihovim otkupima/zbirnim/prijemnicama; ture ostaju zasebni "
                   "dokumenti." % (st_id, _ids(tudji, "OtpremnicaID") or "-"))
        ctx.nalaz(K_TURA_VOZ, SUMNJA, OTP, _ids(rows, "OtpremnicaID"), opis, predlog)


# ---------------------------------------------------------------- 2) datum u broju
def _otp_datum(pg):
    ctx = pg.ctx
    if not pg.ima(K_DATUM, {OTP: ["OtpremnicaID", "Datum", "StanicaID", "BrojOtpremnice", "Stornirano"]}):
        return
    pokr = pg.pokriveno[K_DATUM]
    for dok in pg.otp_dok:
        r0 = dok[0]
        broj = s(r0, "BrojOtpremnice")
        if not broj or k(r0, "StanicaID") in pg.hladnjace:
            continue                     # hladnjacki lanac: broj je broj otkupa (oblast otkupa)
        oid = _ids(dok, "OtpremnicaID")
        dan = d(r0, "Datum")
        m = _HL.match(broj)
        if m:
            try:
                bdan = date(2000 + int(m.group(3)), int(m.group(2)), int(m.group(1)))
            except ValueError:
                bdan = None
        else:
            p = parse_broj(broj)
            if p is None or p.s:
                ctx.nalaz(K_DATUM, INFO, OTP, oid,
                          "Broj otpremnice '%s' nije kanonski (x/ddmmyy[-n]): datum i stanica iz broja se ne "
                          "proveravaju." % broj, "")
                continue
            bdan = p.dan
        if dan is None:
            B = k(r0, "BrojZbirne")
            zb = [z for z in pg.zbr_akt_po_b.get(B, []) if d(z, "Datum") is None] if B else []
            ctx.nalaz(K_DATUM, SUMNJA, OTP, oid,
                      "Otpremnica %s: Datum je prazan ili nije datum ('%s')%s." % (
                          broj, s(r0, "Datum"), ("; zbirna %s (%s) ima isti problem" % (
                              s(zb[0], "BrojZbirne"), _ids(zb, "ZbirnaID"))) if zb else ""),
                      ("Datum: -> %s (dan iz broja)%s." % (_dmy(bdan), " i na zbirnoj" if zb else ""))
                      if bdan else "Upisati Datum.")
            pokr |= _id_set(dok, "OtpremnicaID") | _id_set(zb, "ZbirnaID")
            continue
        if bdan == dan:
            continue
        pokr |= _id_set(dok, "OtpremnicaID")
        B = k(r0, "BrojZbirne")
        zb = [z for z in pg.zbr_akt_po_b.get(B, []) if d(z, "Datum") == dan] if B else []
        pokr |= _id_set(zb, "ZbirnaID")
        if bdan is None:
            ctx.nalaz(K_DATUM, GRESKA, OTP, oid,
                      "Otpremnica %s (%s): datum u broju ne postoji." % (broj, _dmy(dan)),
                      "Ispraviti broj na postojeci datum %s." % _dmy(dan))
            continue
        razlika = (dan - bdan).days
        opis = "Otpremnica %s (%s): broj nosi %s, Datum je %s (razlika %+d dana)" % (
            broj, s(r0, "StanicaID"), _dmy(bdan), _dmy(dan), razlika)
        if bdan.day <= 12 and (bdan.day, bdan.month, bdan.year) == (dan.month, dan.day, dan.year):
            opis += "; zamena dan/mesec (locale)"
        if zb:
            opis += "; zbirna %s (%s) nasledjuje isti nesklad" % (s(zb[0], "BrojZbirne"), _ids(zb, "ZbirnaID"))
        izvori = pg.otkupi(dok) + (pg.prj_akt_po_b.get(B, []) if B else [])
        za_dat = sum(1 for x in izvori if d(x, "Datum") == dan)
        za_broj = sum(1 for x in izvori if d(x, "Datum") == bdan)
        novi = pg.sledeci_broj(s(r0, "StanicaID"), dan, set())
        if za_broj > za_dat:
            predlog = "Datum: %s -> %s (otkupi/prijemnice su od dana iz broja: %d prema %d)." % (
                _dmy(dan), _dmy(bdan), za_broj, za_dat)
        elif za_dat > za_broj and novi:
            predlog = ("Broj je zaostao: BrojOtpremnice %s -> %s (otkupi/prijemnice su od %s: %d prema %d); "
                       "isto preneti na BrojZbirne zbirne i prijemnica (veza je po stringu)." % (
                           broj, novi, _dmy(dan), za_dat, za_broj))
        else:
            predlog = "Izvori ne presudjuju: ili Datum -> %s, ili broj -> %s (uz BrojZbirne zbirne/prijemnica)." % (
                _dmy(bdan), novi or "sledeci slobodan za stanicu i dan")
        ctx.nalaz(K_DATUM, GRESKA, OTP, oid, opis + ".", predlog)
    # zbirne ciji nesklad nije nasledjen od prijavljene otpremnice
    if not pg.zbr or not pg.ima(K_DATUM, {ZBR: ["ZbirnaID", "Datum", "BrojZbirne", "Stornirano"]}):
        return
    for zd in pg.zbr_dok:
        z0 = zd[0]
        if _id_set(zd, "ZbirnaID") & pokr:
            continue
        if pg.stanica_zbirne(zd) in pg.hladnjace:
            continue
        broj = s(z0, "BrojZbirne")
        p = parse_broj(broj)
        if p is None:
            continue
        zid = _ids(zd, "ZbirnaID")
        dan = d(z0, "Datum")
        otps = pg.otp_akt_po_b.get(k(z0, "BrojZbirne"), [])
        if dan is None:
            ctx.nalaz(K_DATUM, SUMNJA, ZBR, zid, "Zbirna %s: Datum je prazan ili nije datum ('%s')." % (
                broj, s(z0, "Datum")), ("Datum: -> %s (dan iz broja)." % _dmy(p.dan)) if p.dan else "Upisati Datum.")
        elif p.dan is None:
            ctx.nalaz(K_DATUM, GRESKA, ZBR, zid, "Zbirna %s (%s): datum u broju ne postoji." % (broj, _dmy(dan)),
                      "Ispraviti broj zbirne.")
        elif p.dan != dan:
            od = _dan(otps)
            if od == p.dan:
                predlog = "Datum zbirne: %s -> %s (kao otpremnica %s)." % (_dmy(dan), _dmy(p.dan),
                                                                         _ids(otps, "OtpremnicaID"))
            else:
                predlog = "Uskladiti Datum zbirne (%s) i broj (%s) sa otpremnicom%s." % (
                    _dmy(dan), _dmy(p.dan), (" " + _ids(otps, "OtpremnicaID")) if otps else "")
            ctx.nalaz(K_DATUM, GRESKA, ZBR, zid, "Zbirna %s: broj nosi %s, Datum je %s (razlika %+d dana)." % (
                broj, _dmy(p.dan), _dmy(dan), (dan - p.dan).days), predlog)
        else:
            continue
        pokr |= _id_set(zd, "ZbirnaID")


# ---------------------------------------------------------------- 3) tudja stanica
def _otp_tudja(pg):
    ctx = pg.ctx
    if not pg.ima(K_TUDJA, {OTP: ["OtpremnicaID", "StanicaID", "BrojOtpremnice", "Stornirano"]}):
        return
    for dok in pg.otp_dok:
        r0 = dok[0]
        st = k(r0, "StanicaID")
        if not st or st in pg.hladnjace:
            continue
        broj = s(r0, "BrojOtpremnice")
        p = parse_broj(broj)
        xs = cifre(s(r0, "StanicaID"))
        if p is None or p.s or xs is None or p.x == xs:
            continue
        oid = _ids(dok, "OtpremnicaID")
        vlasnik = ctx.stanica_po_broju.get(p.x, "")
        B = k(r0, "BrojZbirne")
        zb = pg.zbr_akt_po_b.get(B, []) if B else []
        opis = "Otpremnica %s je stanice %s, a prefiks broja (%d) pripada %s" % (
            broj, s(r0, "StanicaID"), p.x, vlasnik or "stanici koje nema u tblStanice")
        if zb:
            zv = sorted({s(z, "VozacID") for z in zb})
            opis += "; zbirna %s (%s) VozacID %s" % (s(zb[0], "BrojZbirne"), _ids(zb, "ZbirnaID"), "/".join(zv))
            if any(cifre(v) == p.x for v in zv) and pg.ctx.je_stanica(zv[0]):
                opis += " (vozac te druge stanice -- vidi i %s)" % K_VOZAC
        otk = pg.otkupi(dok)
        st_otk = {k(o, "StanicaID") for o in otk}
        if otk and vlasnik and st_otk == {kljuc(vlasnik)}:
            predlog = "StanicaID: %s -> %s (otkupi %s su te stanice)." % (s(r0, "StanicaID"), vlasnik,
                                                                       _ids(otk, "OtkupID"))
        elif otk and st_otk == {st}:
            novi = pg.sledeci_broj(s(r0, "StanicaID"), d(r0, "Datum"), set())
            predlog = ("Broj je zaostao (otkupi %s su stanice %s): BrojOtpremnice %s -> %s, pa isto na BrojZbirne "
                       "zbirne i prijemnica." % (_ids(otk, "OtkupID"), s(r0, "StanicaID"), broj,
                                                 novi or "sledeci slobodan"))
        else:
            predlog = "Presuditi po vezanim otkupima: ili StanicaID -> %s, ili nov broj za %s." % (
                vlasnik or "stanica %d" % p.x, s(r0, "StanicaID"))
        ctx.nalaz(K_TUDJA, GRESKA, OTP, oid, opis + ".", predlog)


# ---------------------------------------------------------------- 4) dupli broj otpremnice
def _otp_dupli(pg):
    ctx = pg.ctx
    if not pg.ima(K_OTP_DUPLI, {OTP: ["OtpremnicaID", "Datum", "StanicaID", "BrojOtpremnice", "Klasa",
                                      "Stornirano"]}):
        return
    for b, rows in sorted(_po(pg.otp_akt, "BrojOtpremnice").items()):
        if len(rows) < 2:
            continue
        konteksti = defaultdict(list)
        for r in rows:
            konteksti[(k(r, "StanicaID"), d(r, "Datum"))].append(r)
        kl = Counter(k(r, "Klasa") for r in rows)
        # Datum koji ne moze da se procita ne pravi poseban kontekst (to prijavljuje OTP_BROJ_DATUM_NESLAGANJE)
        razl = {(st, dn) for st, dn in konteksti if dn is not None}
        razl_st = {st for st, _dn in konteksti}
        if len(razl) < 2 and len(razl_st) < 2 and max(kl.values()) < 2:
            continue                      # Klasa I + II istog dokumenta (isti broj, stanica, dan)
        rows = sorted(rows, key=lambda r: _num(s(r, "OtpremnicaID")))
        pg.pokriveno[K_OTP_DUPLI] |= _id_set(rows, "OtpremnicaID")
        broj = s(rows[0], "BrojOtpremnice")
        p = parse_broj(broj)
        stavke = ["%s %s %s Kl %s %s kg" % (s(r, "OtpremnicaID"), s(r, "StanicaID"), _dmy(d(r, "Datum")),
                                           s(r, "Klasa") or "?", kg_txt(n(r, "Kolicina"))) for r in rows]
        opis = "Broj %s na %d aktivne otpremnice (%s)." % (broj, len(rows), _lista(stavke))
        rez = set()
        if len(konteksti) >= 2:
            slaze = [kx for kx in konteksti if p is not None and not p.s and kx[1] == p.dan
                     and cifre(s(konteksti[kx][0], "StanicaID")) == p.x]
            zadrzi = slaze[0] if len(slaze) == 1 else None
            if zadrzi is None:
                # nijedan ili vise kontekst odgovara broju: zadrzi onaj ciji lanac postoji (zbirna + prijemnica)
                def lanac(kx):
                    B = k(konteksti[kx][0], "BrojZbirne")
                    return bool(B and pg.zbr_akt_po_b.get(B)) + bool(B and pg.prj_akt_po_b.get(B))
                zadrzi = sorted(konteksti, key=lambda kx: (-lanac(kx), _num(s(konteksti[kx][0], "OtpremnicaID"))))[0]
            delovi = ["zadrzati broj na %s" % _ids(konteksti[zadrzi], "OtpremnicaID")]
            for kx, rr in sorted(konteksti.items(), key=lambda x: _num(s(x[1][0], "OtpremnicaID"))):
                if kx == zadrzi:
                    continue
                novi = pg.sledeci_broj(s(rr[0], "StanicaID"), kx[1], rez)
                delovi.append("%s -> %s" % (_ids(rr, "OtpremnicaID"), novi or "sledeci slobodan za stanicu i dan"))
            predlog = "; ".join(delovi) + "; nov broj preneti na BrojZbirne njene zbirne i prijemnica."
        else:
            dupla = [c for c, x in kl.items() if x > 1]
            vis = [r for r in rows if k(r, "Klasa") in dupla][1:]
            predlog = ("Ista klasa dvaput u istom dokumentu: ako je ista isporuka, stornirati visak %s; inace mu "
                       "dodeliti %s." % (_ids(vis, "OtpremnicaID"),
                                         pg.sledeci_broj(s(rows[0], "StanicaID"), d(rows[0], "Datum"), rez)
                                         or "sledeci slobodan broj"))
        ctx.nalaz(K_OTP_DUPLI, GRESKA, OTP, _ids(rows, "OtpremnicaID"), opis, predlog)


# ---------------------------------------------------------------- 5) dupli broj zbirne
def _zbr_dupli(pg):
    ctx = pg.ctx
    if not pg.ima(K_ZBR_DUPLI, {ZBR: ["ZbirnaID", "Datum", "VozacID", "BrojZbirne", "Klasa", "Stornirano"]}):
        return
    for B, zs in sorted(pg.zbr_akt_po_b.items()):
        if len(zs) < 2:
            continue
        kl = Counter(c for z in zs for c in _klase(z))
        dani = {d(z, "Datum") for z in zs if d(z, "Datum") is not None}
        voz = {k(z, "VozacID") for z in zs}
        if max(kl.values()) < 2 and len(dani) < 2 and len(voz) < 2:
            continue                      # Klasa I + II iste zbirne
        pg.zbr_dupli.add(B)
        zs = sorted(zs, key=lambda z: _num(s(z, "ZbirnaID")))
        otps = sorted(pg.otp_akt_po_b.get(B, []), key=lambda o: _num(s(o, "OtpremnicaID")))
        stavke = ["%s %s %s Kl %s %s kg" % (s(z, "ZbirnaID"), _dmy(d(z, "Datum")), s(z, "VozacID") or "-",
                                           s(z, "Klasa") or "?", kg_txt(n(z, "UkupnoKolicina"))) for z in zs]
        opis = "Broj zbirne %s na %d aktivne zbirne (%s)" % (s(zs[0], "BrojZbirne"), len(zs), _lista(stavke))
        opis += ("; aktivne otpremnice tog broja: %s" % _lista(
            "%s %s %s kg" % (s(o, "OtpremnicaID"), s(o, "BrojOtpremnice"), kg_txt(n(o, "Kolicina"))) for o in otps)
                 if otps else "; nijedna aktivna otpremnica nema taj broj")
        # uparivanje zbirna <-> otpremnica po kg, danu, vozacu i klasi
        slobodne = list(otps)
        par, visak = [], []
        for z in zs:
            nad = None
            for o in slobodne:
                if (abs(n(o, "Kolicina") - n(z, "UkupnoKolicina")) <= _EPS and d(o, "Datum") == d(z, "Datum")
                        and k(o, "VozacID") == k(z, "VozacID") and (k(o, "Klasa") in _klase(z))):
                    nad = o
                    break
            if nad is not None:
                slobodne.remove(nad)
                par.append((z, nad))
            else:
                visak.append(z)
        if par and visak:
            predlog = "Zadrzati %s; visak bez para u otpremnici (kg/dan/vozac): stornirati %s." % (
                ", ".join("%s (= %s)" % (s(z, "ZbirnaID"), s(o, "OtpremnicaID")) for z, o in par),
                _ids(visak, "ZbirnaID"))
        elif par:
            predlog = ("Svaka zbirna ima svoju otpremnicu istog broja (%s): prvo resiti %s, pa zbirnu druge "
                       "otpremnice preimenovati u S + njen nov broj." % (_ids([o for _z, o in par], "OtpremnicaID"),
                                                                        K_OTP_DUPLI))
        else:
            predlog = "Nijedna zbirna ne odgovara otpremnici po kg/danu/vozacu -- operater bira koja ostaje."
        ctx.nalaz(K_ZBR_DUPLI, GRESKA, ZBR, _ids(zs, "ZbirnaID"), opis + ".", predlog)


# ---------------------------------------------------------------- 6) zbirna sa storniranim otpremnicama
def _sv_ispravke(pg, B, sto):
    """Konteksti ISPRAVKA_ODMAH otpremnice koji vode na zbirnu B ili na njene stornirane otpremnice."""
    sto_ids = _id_set(sto, "OtpremnicaID")
    sto_br = _id_set(sto, "BrojOtpremnice")
    out = []
    for c in pg.sv:
        if k(c, "Mode") != "ispravka_odmah" or k(c, "OldDocType") != "otpremnica":
            continue
        if k(c, "ParentBroj") == B or k(c, "OldDocID") in sto_ids or k(c, "OldBroj") in sto_br:
            out.append(c)
    return out


def _naslednik_sv(pg, c):
    nid, nb = k(c, "NewDocID"), k(c, "NewBroj")
    o = pg.ctx.po_id[OTP].get(nid) if nid else None
    if o is not None and not storno(o):
        return o
    for x in pg.otp_akt:
        if nb and k(x, "BrojOtpremnice") == nb:
            return x
    return None


def _kandidati_naslednika(pg, sto, uslov_zbirne):
    """Aktivne otpremnice iste stanice, dana i klase kao stornirane, veceg ID-ja."""
    st = _id_set(sto, "StanicaID")
    dani = {d(o, "Datum") for o in sto if d(o, "Datum") is not None}
    kl = {k(o, "Klasa") for o in sto}
    maxid = max(_num(s(o, "OtpremnicaID")) for o in sto)
    out = []
    for o in pg.otp_akt:
        if (k(o, "StanicaID") in st and d(o, "Datum") in dani and k(o, "Klasa") in kl
                and _num(s(o, "OtpremnicaID")) > maxid and uslov_zbirne(o)):
            out.append(o)
    return sorted(out, key=lambda o: _num(s(o, "OtpremnicaID")))


def _zbr_stornirane_otp(pg):
    ctx = pg.ctx
    potrebno = {ZBR: ["ZbirnaID", "Datum", "BrojZbirne", "UkupnoKolicina", "Stornirano"],
                OTP: ["OtpremnicaID", "Datum", "StanicaID", "BrojZbirne", "Kolicina", "Stornirano"]}
    ok = [pg.ima(kod, potrebno) for kod in (K_VISI, K_ISPR, K_DELIM)]
    if not all(ok):
        return
    for B, zs in sorted(pg.zbr_akt_po_b.items()):
        sve = pg.otp_po_b.get(B, [])
        sto = [o for o in sve if storno(o)]
        if not sto:
            continue
        akt = [o for o in sve if not storno(o)]
        prj = pg.prj_akt_po_b.get(B, [])
        zid = _ids(zs, "ZbirnaID")
        broj = s(zs[0], "BrojZbirne")
        kgz = _kg(zs, "UkupnoKolicina")
        sv = _sv_ispravke(pg, B, sto)
        slova = sorted({s(o, "BrojZbirne") for o in sve if s(o, "BrojZbirne") != broj})
        nap_slova = (" Otpremnice nose broj '%s' koji se od zbirne razlikuje samo u velikim slovima -- stari kod ih "
                     "ne spaja." % "', '".join(slova)) if slova else ""
        if not akt:
            # --- P5: ispravka zavrsena bez nove zbirne (0 kg + prijemnica + naslednik bez zbirne)
            if kgz == 0 and prj:
                def bez_prijemnice(o):
                    b2 = k(o, "BrojZbirne")
                    return not b2 or not pg.prj_akt_po_b.get(b2)
                nasl, iz_sv = [], None
                for c in sv:
                    o = _naslednik_sv(pg, c)
                    if o is not None and bez_prijemnice(o) and k(o, "BrojZbirne") != B:
                        nasl, iz_sv = [o], c
                        break
                if not nasl:
                    nasl = _kandidati_naslednika(pg, sto, bez_prijemnice)
                if nasl:
                    o = nasl[0]
                    b2 = s(o, "BrojZbirne")
                    opis = ("Ispravka otpremnice zavrsena bez nove zbirne: aktivna zbirna %s (%s) ima 0 kg i "
                            "prijemnicu %s; stara otpremnica %s stornirana; naslednik %s %s %s" % (
                                broj, zid, pg.prijemnice_txt(prj), _ids(sto, "OtpremnicaID"), s(o, "OtpremnicaID"),
                                s(o, "BrojOtpremnice"), ("ima zbirnu %s bez prijemnice" % b2) if b2
                                else "nema BrojZbirne"))
                    if iz_sv is not None:
                        opis += " [kontekst %s, status %s]" % (s(iz_sv, "CorrectionID"), s(iz_sv, "Status") or "-")
                    cilj = b2 or pg.ocek(s(o, "VozacID"), s(o, "BrojOtpremnice"))
                    predlog = ""
                    if not b2:
                        predlog = "Napraviti zbirnu %s iz %s (kg %s, gajbe %s); " % (
                            cilj, s(o, "OtpremnicaID"), kg_txt(n(o, "Kolicina")), kg_txt(n(o, "KolAmbalaze")))
                    predlog += "prevezati %s: BrojZbirne %s -> %s, VozacID %s; zatim stornirati %s. kg prijemnice %s " \
                               "vs otpremnica %s." % (_ids(prj, "PrijemnicaID"), broj, cilj, s(o, "VozacID") or "-",
                                                      zid, kg_txt(_kg(prj, "Kolicina")), kg_txt(n(o, "Kolicina")))
                    ctx.nalaz(K_ISPR, GRESKA if iz_sv is not None else SUMNJA, ZBR, zid, opis + "." + nap_slova,
                              predlog)
                    continue
            # --- ZBR_VISI
            opis = "Aktivna zbirna %s (%s, %s kg): sve njene otpremnice su stornirane (%s); aktivnih prijemnica %d%s" % (
                broj, zid, kg_txt(kgz), _ids(sto, "OtpremnicaID"), len(pg.prijemnice_dok(prj)),
                (" (%s)" % pg.prijemnice_txt(prj)) if prj else "")
            if prj and kgz == 0:
                opis += "; 0 kg uz prijemnicu = potpis ISPRAVKA puta v2.19+ (naslednik nije nadjen)"
            elif kgz > 0:
                opis += "; zbirna nije kaskadno stornirana (potpis storna pre v2.4.0)"
            if not prj:
                predlog = "Stornirati zbirnu %s (nema prijemnicu)." % zid
            else:
                kand = []
                for c in sv:
                    o = _naslednik_sv(pg, c)
                    if o is not None and k(o, "BrojZbirne") and k(o, "BrojZbirne") != B:
                        kand.append(o)
                kgp = _kg(prj, "Kolicina")
                if not kand:
                    kand = _kandidati_naslednika(pg, sto, lambda o: bool(
                        k(o, "BrojZbirne") and pg.zbr_akt_po_b.get(k(o, "BrojZbirne"))
                        and not pg.prj_akt_po_b.get(k(o, "BrojZbirne"))))
                    kand.sort(key=lambda o: abs(n(o, "Kolicina") - kgp))
                if kand:
                    o = kand[0]
                    predlog = ("Prevezati %s: BrojZbirne %s -> %s, VozacID %s (naslednik %s, kg %s vs prijemnica %s), "
                               "pa stornirati %s." % (_ids(prj, "PrijemnicaID"), broj, s(o, "BrojZbirne"),
                                                      s(o, "VozacID") or "-", s(o, "OtpremnicaID"),
                                                      kg_txt(n(o, "Kolicina")), kg_txt(kgp), zid))
                    if len(kand) > 1:
                        predlog += " Ostali kandidati: %s." % _ids(kand[1:], "OtpremnicaID")
                else:
                    predlog = "Naslednik nije nadjen: stornirati %s; o prijemnici %s odlucuje operater." % (
                        zid, _ids(prj, "PrijemnicaID"))
            ctx.nalaz(K_VISI, GRESKA, ZBR, zid, opis + "." + nap_slova, predlog)
            continue
        # --- deo otpremnica storniran: zbir aktivnih mora biti = zbirna
        if B in pg.zbr_dupli:
            continue
        if any(len(_klase(z)) > 1 for z in zs):
            zk = {"*": kgz}
            ak = {"*": _kg(akt, "Kolicina")}
        else:
            zk, ak = defaultdict(Decimal), defaultdict(Decimal)
            for z in zs:
                zk[k(z, "Klasa")] += n(z, "UkupnoKolicina")
            for o in akt:
                ak[k(o, "Klasa")] += n(o, "Kolicina")
        razlike = [(c, zk.get(c, Decimal(0)), ak.get(c, Decimal(0))) for c in sorted(set(zk) | set(ak))
                   if abs(zk.get(c, Decimal(0)) - ak.get(c, Decimal(0))) > _EPS]
        if not razlike:
            continue
        opis = ("Aktivna zbirna %s (%s) ima aktivne otpremnice (%s) i stornirane (%s); UkupnoKolicina ne prati "
                "aktivne: %s" % (broj, zid, _ids(akt, "OtpremnicaID"), _ids(sto, "OtpremnicaID"),
                                 "; ".join("Kl %s zbirna %s vs otpremnice %s" % (
                                     c if c != "*" else "sve", kg_txt(a), kg_txt(b)) for c, a, b in razlike)))
        otv = [c for c in sv if k(c, "Status") != "completed"]
        if otv:
            opis += "; otvoren kontekst ispravke %s (%s)" % (
                ", ".join(s(c, "CorrectionID") for c in otv), "/".join(sorted({s(c, "Status") for c in otv})))
        predlog = "UkupnoKolicina rekalkulisati na zbir aktivnih otpremnica: %s." % "; ".join(
            "Kl %s %s -> %s" % (c if c != "*" else "sve", kg_txt(a), kg_txt(b)) for c, a, b in razlike)
        nb = [s(c, "NewBroj") for c in sv if s(c, "NewBroj")]
        if nb:
            predlog += " Naslednik iz ispravke: %s." % ", ".join(nb)
        ctx.nalaz(K_DELIM, SUMNJA, ZBR, zid, opis + "." + nap_slova, predlog)


# ---------------------------------------------------------------- 7) otpremnica na storniranoj zbirnoj
def _otp_ceka(pg):
    ctx = pg.ctx
    potrebno = {OTP: ["OtpremnicaID", "BrojOtpremnice", "BrojZbirne", "Stornirano"],
                ZBR: ["ZbirnaID", "BrojZbirne", "Stornirano"]}
    if not pg.ima(K_CEKA, potrebno) or not pg.ima(K_CEKA_TACNO, potrebno):
        return
    for dok in pg.otp_dok:
        r0 = dok[0]
        B = k(r0, "BrojZbirne")
        zall = pg.zbr_po_b.get(B, []) if B else []
        if not zall:
            continue                      # broj ne postoji ni storniran: HLADNJACA_LANAC_NEPOTPUN
        zakt = pg.zbr_akt_po_b.get(B, [])
        oid = _ids(dok, "OtpremnicaID")
        prj = pg.prj_akt_po_b.get(B, [])
        bt = s(r0, "BrojZbirne")
        if not zakt:
            opis = "Aktivna otpremnica %s pokazuje na zbirnu %s koja postoji samo stornirana (%s)%s." % (
                s(r0, "BrojOtpremnice"), bt, _ids(zall, "ZbirnaID"),
                ("; aktivne prijemnice tog broja su osirocene: %s" % pg.prijemnice_txt(prj)) if prj else "")
            oc = pg.ocek(s(r0, "VozacID"), s(r0, "BrojOtpremnice"))
            ozakt = pg.zbr_akt_po_b.get(kljuc(oc), []) if oc else []
            if ozakt and kljuc(oc) != B:
                predlog = "Vezati za aktivnu zbirnu %s (%s): BrojZbirne %s -> %s." % (oc, _ids(ozakt, "ZbirnaID"),
                                                                                    bt, oc)
            elif pg.malina and oc:
                predlog = ("Napraviti novu zbirnu %s (Datum %s, VozacID %s, kg %s, gajbe %s sa otpremnice)%s." % (
                    oc, _dmy(d(r0, "Datum")), s(r0, "VozacID") or "-", kg_txt(_kg(dok, "Kolicina")),
                    kg_txt(_kg(dok, "KolAmbalaze")),
                    "; prijemnice %s prate (isti broj)" % _ids(prj, "PrijemnicaID") if prj and kljuc(oc) == B else ""))
            else:
                predlog = "Napraviti novu zbirnu za otpremnicu ili isprazniti BrojZbirne (stanje 'nevezana')."
            ctx.nalaz(K_CEKA, GRESKA, OTP, oid, opis, predlog)
            continue
        tacno = [z for z in zall if s(z, "BrojZbirne") == bt]
        if tacno and all(storno(z) for z in tacno):
            ctx.nalaz(K_CEKA_TACNO, SUMNJA, OTP, oid,
                      "Otpremnica %s: stari kod (tacno poredjenje) za BrojZbirne '%s' vidi samo storniranu %s; "
                      "aktivna %s ima '%s' (razlika samo u velikim slovima)." % (
                          s(r0, "BrojOtpremnice"), bt, _ids(tacno, "ZbirnaID"), _ids(zakt, "ZbirnaID"),
                          s(zakt[0], "BrojZbirne")),
                      "BrojZbirne: '%s' -> '%s' (tacno kao aktivna zbirna)." % (bt, s(zakt[0], "BrojZbirne")))


# ---------------------------------------------------------------- 8) malina: otpremnica bez zbirne
def _z2_prijemnice(pg):
    """Prijemnice na zbirnim koje imaju >= 2 prijemnice: (B, prows, kg zbirne)."""
    out = []
    for B, prows in pg.prj_akt_po_b.items():
        pdok = pg.prijemnice_dok(prows)
        zs = pg.zbr_akt_po_b.get(B, [])
        if len(pdok) >= 2 and zs:
            kgz = _kg(zs, "UkupnoKolicina")
            out.extend((B, pd, kgz) for pd in pdok)
    return out


def _otp_bez_zbirne(pg):
    ctx = pg.ctx
    if not pg.malina:
        return
    if not pg.ima(K_BEZ, {OTP: ["OtpremnicaID", "Datum", "StanicaID", "VozacID", "BrojOtpremnice", "BrojZbirne",
                                "Stornirano"]}):
        return
    bez = [dok for dok in pg.otp_dok if not k(dok[0], "BrojZbirne")]
    if not bez:
        return
    dani = [d(x[0], "Datum") for x in bez if d(x[0], "Datum") is not None]
    d0 = min(dani) if dani else None
    neparsirani = [x for x in bez if d(x[0], "Datum") is None]
    blok = ""
    if d0 is not None and len(bez) >= 2:
        od = [x for x in pg.otp_dok if d(x[0], "Datum") is not None and d(x[0], "Datum") >= d0]
        if od and all(not k(x[0], "BrojZbirne") for x in od):
            blok = " Od %s su SVE aktivne otpremnice (%d) bez zbirne -- potpis pada celog prolaza auto-zbirne%s." % (
                _dmy(d0), len(od), ("; Datum koji ne moze da se procita: %s" % _lista(
                    _ids(x, "OtpremnicaID") for x in neparsirani)) if neparsirani else "")
    z2 = _z2_prijemnice(pg) if pg.zbr else []
    for dok in bez:
        r0 = dok[0]
        oid = _ids(dok, "OtpremnicaID")
        dan = d(r0, "Datum")
        broj = s(r0, "BrojOtpremnice")
        oc = pg.ocek(s(r0, "VozacID"), broj)
        zakt = pg.zbr_akt_po_b.get(kljuc(oc), []) if oc else []
        # ocekivanu zbirnu vec drzi DRUGA aktivna otpremnica (isti broj, OTP_DUPLI): veza nije "izgubljena"
        zauzeta = [o for o in pg.otp_akt_po_b.get(kljuc(oc), []) if k(o, "OtpremnicaID") not in
                   _id_set(dok, "OtpremnicaID")] if zakt else []
        kgo = _kg(dok, "Kolicina")
        opis = "Malina: aktivna otpremnica %s (%s, %s, %s kg) nema BrojZbirne" % (
            broj or "-", s(r0, "StanicaID"), _dmy(dan), kg_txt(kgo))
        if zakt and zauzeta:
            opis += ("; ocekivana zbirna %s (%s) postoji, ali je vec zbirna otpremnice %s (isti broj otpremnice)" % (
                oc, _ids(zakt, "ZbirnaID"), _ids(zauzeta, "OtpremnicaID")))
            zakt = []
        elif zakt:
            opis += "; ocekivana zbirna %s postoji aktivna (%s) -- veza je izgubljena" % (oc, _ids(zakt, "ZbirnaID"))
        elif oc:
            opis += "; zbirna %s ne postoji (auto-zbirna nije napravljena)" % oc
        if dan is None:
            opis += "; Datum '%s' ne moze da se procita (CDate obara ceo prolaz auto-zbirne)" % s(r0, "Datum")
        opis += "." + blok
        if zakt:
            predlog = "BrojZbirne: '' -> %s (%s)." % (oc, _ids(zakt, "ZbirnaID"))
        else:
            predlog = "Napraviti zbirnu %s (Datum %s, VozacID %s, kg %s, gajbe %s sa otpremnice) i upisati broj." % (
                ("S + nov broj otpremnice (prvo resiti %s)" % K_OTP_DUPLI) if zauzeta else (oc or "S+broj otpremnice"),
                _dmy(dan), s(r0, "VozacID") or "-", kg_txt(kgo),
                kg_txt(_kg(dok, "KolAmbalaze")))
            st = k(r0, "StanicaID")
            kand = []
            for B, pd, kgz in z2:
                pdan = d(pd[0], "Datum")
                if dan is None or pdan is None or (pdan - dan).days not in (0, 1):
                    continue
                if pg.stanica_prijemnice(pd) != st:
                    continue
                kgp = _kg(pd, "Kolicina")
                if abs(kgp - kgo) < abs(kgp - kgz):
                    kand.append((abs(kgp - kgo), B, pd, kgp))
            if kand:
                kand.sort(key=lambda x: x[0])
                predlog += " Pre toga: %s." % "; ".join(
                    "%s (%s kg, na %s sa 2+ prijemnice) je verovatno njena -- prevezati na %s" % (
                        _ids(pd, "PrijemnicaID"), kg_txt(kgp), s(pd[0], "BrojZbirne"),
                        "novu zbirnu" if zauzeta else (oc or "novu zbirnu"))
                    for _x, B, pd, kgp in kand[:3])
        ctx.nalaz(K_BEZ, GRESKA if zakt else SUMNJA, OTP, oid, opis, predlog)


# ---------------------------------------------------------------- 9) nepotpun lanac
def _hladnjaca_lanac(pg):
    ctx = pg.ctx
    if not pg.ima(K_HL, {OTP: ["OtpremnicaID", "StanicaID", "VozacID", "BrojOtpremnice", "BrojZbirne", "Stornirano"],
                         ZBR: ["ZbirnaID", "BrojZbirne", "Stornirano"]}):
        return
    # (a) otpremnica pokazuje na broj zbirne kog nema ni storniranog
    for dok in pg.otp_dok:
        r0 = dok[0]
        B = k(r0, "BrojZbirne")
        if not B or B in pg.zbr_po_b:
            continue
        hl = k(r0, "StanicaID") in pg.hladnjace
        prj = pg.prj_akt_po_b.get(B, [])
        oc = pg.ocek(s(r0, "VozacID"), s(r0, "BrojOtpremnice"))
        opis = "Aktivna otpremnica %s (%s) ima BrojZbirne %s, a tog broja nema u tblZbirna (ni storniranog)%s%s." % (
            s(r0, "BrojOtpremnice"), s(r0, "StanicaID"), s(r0, "BrojZbirne"),
            "; stanica je hladnjaca (auto-lanac)" if hl else "",
            ("; prijemnice tog broja: %s" % pg.prijemnice_txt(prj)) if prj else "")
        predlog = "Napraviti zbirnu %s iz otpremnice (Datum %s, VozacID %s, kg %s, gajbe %s)%s." % (
            s(r0, "BrojZbirne"), _dmy(d(r0, "Datum")), s(r0, "VozacID") or "-", kg_txt(_kg(dok, "Kolicina")),
            kg_txt(_kg(dok, "KolAmbalaze")),
            ("; ocekivani broj je %s" % oc) if oc and kljuc(oc) != B else "")
        # van maline forma upisuje predlozeni broj zbirne na otpremnicu PRE nego sto zbirna postoji
        # (RefreshBrojZbirneSuggestion -> txtBrojZbirneOtp): zbirna moze jos da ceka rucni unos
        if not pg.malina:
            predlog += " Van maline: proveriti da li je zbirna uneta pod drugim brojem ili jos ceka unos."
        ctx.nalaz(K_HL, GRESKA if pg.malina else SUMNJA, OTP, _ids(dok, "OtpremnicaID"), opis, predlog)
    # (b) aktivna zbirna bez ijedne otpremnice sa tim brojem
    ocekivane = {kljuc(pg.ocek(s(r, "VozacID"), s(r, "BrojOtpremnice"))) for r in pg.otp_akt}
    for B, zs in sorted(pg.zbr_akt_po_b.items()):
        if B in pg.otp_po_b or B in ocekivane:
            continue                      # ima otpremnicu, ili je ocekivana zbirna neke otpremnice (veza pogresna)
        # PWA zbirna vozaca nosi SyncSource=PWA i ClientRecordID (modMasterSync.bas:2648-2676, v2.11.0)
        zs = [z for z in zs if k(z, "SyncSource") != "pwa" and not k(z, "ClientRecordID")]
        if not zs:
            continue                      # zbirna iz PWA vozaca (VOZ list) nema otpremnicu po dizajnu
        zid = _ids(zs, "ZbirnaID")
        st = pg.stanica_zbirne(zs)
        hl = st in pg.hladnjace
        otk = pg.otk_akt_po_b.get(B, [])
        prj = pg.prj_akt_po_b.get(B, [])
        opis = "Aktivna zbirna %s (%s, %s, %s kg) nema nijednu otpremnicu (ni storniranu) sa tim BrojZbirne%s%s%s." % (
            s(zs[0], "BrojZbirne"), zid, _dmy(_dan(zs)), kg_txt(_kg(zs, "UkupnoKolicina")),
            "; stanica %s je hladnjaca: potpis auto-lanca (otpremnica pala, zbirna upisana)" % st.upper() if hl else "",
            ("; otkupi sa tim brojem zbirne: %s" % _ids(otk, "OtkupID")) if otk else "",
            ("; prijemnice: %s" % pg.prijemnice_txt(prj)) if prj else "")
        if otk:
            p = parse_broj(s(zs[0], "BrojZbirne"))
            predlog = "Rekonstruisati otpremnicu iz otkupa %s (kg %s, gajbe %s, broj %s)." % (
                _ids(otk, "OtkupID"), kg_txt(_kg(otk, "Kolicina")), kg_txt(_kg(otk, "KolAmbalaze")),
                s(zs[0], "BrojZbirne")[1:] if p is not None and p.s else s(zs[0], "BrojZbirne"))
        else:
            predlog = "Otkup sa tim brojem ne postoji: stornirati zbirnu %s%s." % (
                zid, ("; prijemnica %s tada ostaje bez zbirne -- odlucuje operater" % _ids(prj, "PrijemnicaID"))
                if prj else "")
        ctx.nalaz(K_HL, GRESKA if otk else SUMNJA, ZBR, zid, opis, predlog)


# ---------------------------------------------------------------- 10) tura bez prijemnice
def _pomeri(dan, dd):
    """dan + dd dana, ili None kad izlazi iz opsega datuma (npr. Datum 31.12.9999)."""
    try:
        return dan + timedelta(days=dd)
    except OverflowError:
        return None


def _klase_preklop(a, b):
    """Klase se preklapaju; prazna Klasa (stari red bez klase) ne iskljucuje par."""
    return bool(a & b) or "" in a or "" in b


def _kg_klase(zs, pk):
    """UkupnoKolicina redova zbirne cija se klasa preklapa sa klasom prijemnice (sve kad nijedan)."""
    rr = [z for z in zs if _klase_preklop(_klase(z), pk)]
    return _kg(rr or zs, "UkupnoKolicina")


def _zbr_tura(pg):
    ctx = pg.ctx
    if not pg.ima(K_TURA, {ZBR: ["ZbirnaID", "Datum", "VozacID", "BrojZbirne", "UkupnoKolicina", "Stornirano"],
                           PRJ: ["PrijemnicaID", "BrojPrijemnice", "BrojZbirne", "Kolicina", "Stornirano"]}):
        return
    info = {}
    for B, zs in pg.zbr_akt_po_b.items():
        dan = _dan(zs)
        if dan is None:
            continue
        info[B] = dict(zs=zs, dan=dan, kg=_kg(zs, "UkupnoKolicina"), gaj=_kg(zs, "UkupnoAmbalaze"),
                       klase=frozenset(c for z in zs for c in _klase(z)), voz=k(zs[0], "VozacID"),
                       p=parse_broj(s(zs[0], "BrojZbirne")), pdok=pg.prijemnice_dok(pg.prj_akt_po_b.get(B, [])),
                       # zbirna cije su SVE otpremnice stornirane nije tura (ZBR_VISI): ne prima prijemnicu
                       mrtva=bool(pg.otp_po_b.get(B)) and not pg.otp_akt_po_b.get(B))
    z2_po_danu = defaultdict(list)
    for B, i in info.items():
        if len(i["pdok"]) >= 2:
            z2_po_danu[i["dan"]].append(B)
    # svi parovi (Z0, prijemnica sa Z2); svaka prijemnica se predlaze samo JEDNOJ Z0 (najboljoj)
    najbolji = {}
    for B0 in sorted(info):
        i0 = info[B0]
        if i0["pdok"] or i0["mrtva"]:
            continue
        for dd in (-1, 0, 1):
            dan2 = _pomeri(i0["dan"], dd)
            if dan2 is None:
                continue
            for B2 in z2_po_danu.get(dan2, []):
                i2 = info[B2]
                if not _klase_preklop(i0["klase"], i2["klase"]):
                    continue              # druga klasa (npr. Klasa II iz PWA puta)
                p0, p2 = i0["p"], i2["p"]
                ista = (p0 is not None and p2 is not None and p0.s == p2.s and p0.x == p2.x) or i0["voz"] == i2["voz"]
                nerazl = ista and abs(i0["kg"] - i2["kg"]) <= _EPS      # ture iste kg: kg ne presudjuje
                for pd in i2["pdok"]:
                    pk = frozenset(c for r in pd for c in _klase(r))
                    if not _klase_preklop(pk, i0["klase"]):
                        continue          # prijemnica druge klase ne ide na ovu turu
                    if sum(1 for x in i2["pdok"] if _klase_preklop(frozenset(c for r in x for c in _klase(r)), pk)) < 2:
                        continue          # Z2 ima samo jednu prijemnicu te klase (npr. Kl I i II pod dva broja)
                    kgp = _kg(pd, "Kolicina")
                    d0, d2 = abs(kgp - _kg_klase(i0["zs"], pk)), abs(kgp - _kg_klase(i2["zs"], pk))
                    vozp = k(pd[0], "VozacID")
                    voz_ok = bool(vozp) and vozp == i0["voz"] and vozp != i2["voz"]
                    if not (d0 < d2 or voz_ok or nerazl):
                        continue
                    rang = (not voz_ok, not (d0 < d2 or voz_ok), vozp != i0["voz"], d0,
                            _num(s(i0["zs"][0], "ZbirnaID")), _num(s(pd[0], "PrijemnicaID")))
                    kl = (B2, _ids(pd, "PrijemnicaID"))
                    if kl not in najbolji or rang < najbolji[kl][0]:
                        najbolji[kl] = (rang, B0, B2, pd, kgp, ista, nerazl)
    po_z0 = defaultdict(list)
    for x in najbolji.values():
        po_z0[x[1]].append(x)
    for B0 in sorted(po_z0):
        kand = sorted(po_z0[B0], key=lambda x: x[0])
        i0 = info[B0]
        _r, _b0, B2, pd, kgp, ista, nerazl = kand[0]
        i2 = info[B2]
        z0, z2 = i0["zs"][0], i2["zs"][0]
        opis = ("Zbirna %s (%s, %s, %s kg, %s gajbi) nema prijemnicu, a susedna %s (%s, %s, %s kg) ima %d: %s "
                "(varijanta %s)" % (
                    s(z0, "BrojZbirne"), _ids(i0["zs"], "ZbirnaID"), _dmy(i0["dan"]), kg_txt(i0["kg"]),
                    kg_txt(i0["gaj"]), s(z2, "BrojZbirne"), _ids(i2["zs"], "ZbirnaID"), _dmy(i2["dan"]),
                    kg_txt(i2["kg"]), len(i2["pdok"]), pg.prijemnice_txt([r for x in i2["pdok"] for r in x], True),
                    "b: ista stanica, druga tura ili susedni dan" if ista else "a: druga stanica"))
        kup = s(z0, "KupacID")
        if kup and ctx.malina_kupac and kljuc(kup) != kljuc(ctx.malina_kupac):
            opis += "; KupacID %s nije MALINA_DEFAULT_KUPAC" % kup
        gp = _kg(pd, "KolAmbalaze")
        vozp = s(pd[0], "VozacID")
        predlog = "%s (%s): BrojZbirne %s -> %s, VozacID %s; kg %s vs zbirna %s, gajbe %s %s %s." % (
            _ids(pd, "PrijemnicaID"), s(pd[0], "BrojPrijemnice"), s(pd[0], "BrojZbirne"), s(z0, "BrojZbirne"),
            ("%s (isti)" % vozp) if kljuc(vozp) == i0["voz"] else "%s -> %s" % (vozp or "-", s(z0, "VozacID")),
            kg_txt(kgp), kg_txt(i0["kg"]), kg_txt(gp), "=" if gp == i0["gaj"] else "vs", kg_txt(i0["gaj"]))
        if nerazl:
            predlog += " Ture imaju istu kg (%s): kg ne presudjuje koja prijemnica je ove ture -- operater bira." % (
                kg_txt(i0["kg"]))
        ostali = ["%s (%s kg)" % (_ids(x[3], "PrijemnicaID"), kg_txt(x[4])) for x in kand[1:4]]
        if ostali:
            predlog += " Ostali kandidati: %s." % ", ".join(ostali)
        ctx.nalaz(K_TURA, SUMNJA, ZBR, _ids(i0["zs"], "ZbirnaID"), opis + ".", predlog)


# ---------------------------------------------------------------- 11) clan van 1:1
def _zbr_clan_van(pg):
    ctx = pg.ctx
    if not pg.malina:
        return
    if not pg.ima(K_CLAN, {OTP: ["OtpremnicaID", "Datum", "StanicaID", "VozacID", "BrojOtpremnice", "BrojZbirne",
                                 "Stornirano"],
                           ZBR: ["ZbirnaID", "Datum", "VozacID", "BrojZbirne", "Stornirano"]}):
        return
    pokr = pg.pokriveno[K_CLAN]
    for dok in pg.otp_dok:
        r0 = dok[0]
        B, bo = k(r0, "BrojZbirne"), s(r0, "BrojOtpremnice")
        zakt = pg.zbr_akt_po_b.get(B, []) if (B and bo) else []
        if not zakt:
            continue
        voz = k(r0, "VozacID")
        if voz not in pg.stanice and not (_id_set(zakt, "VozacID") & pg.stanice):
            continue                      # pravi vozac: zbirna sa vise OM je legitimna
        oc = pg.ocek(s(r0, "VozacID"), bo)
        if norm_broj(oc) == norm_broj(s(r0, "BrojZbirne")):
            continue
        pokr |= _id_set(dok, "OtpremnicaID")
        ozakt = pg.zbr_akt_po_b.get(kljuc(oc), [])
        clanovi = [o for o in pg.otp_akt_po_b.get(B, []) if k(o, "OtpremnicaID") not in _id_set(dok, "OtpremnicaID")]
        opis = "Otpremnica %s (%s, %s) je na zbirnoj %s (%s), a auto-zbirna bi joj dala %s%s; zbirna %s kg, " \
               "otpremnica %s kg." % (
                   bo, s(r0, "StanicaID"), _dmy(d(r0, "Datum")), s(r0, "BrojZbirne"), _ids(zakt, "ZbirnaID"), oc,
                   ("; na istoj zbirnoj i %s" % _lista("%s %s" % (s(o, "OtpremnicaID"), s(o, "BrojOtpremnice"))
                                                      for o in clanovi)) if clanovi else "",
                   kg_txt(_kg(zakt, "UkupnoKolicina")), kg_txt(_kg(dok, "Kolicina")))
        if ozakt:
            predlog = "BrojZbirne: %s -> %s (%s); rekalkulisati obe zbirne." % (s(r0, "BrojZbirne"), oc,
                                                                               _ids(ozakt, "ZbirnaID"))
        else:
            predlog = "Zbirna %s ne postoji: napraviti je iz otpremnice (kao %s) i upisati u BrojZbirne." % (oc, K_BEZ)
        ctx.nalaz(K_CLAN, SUMNJA, OTP, _ids(dok, "OtpremnicaID"), opis, predlog)
    # zbirna sa otpremnicama vise stanica ili dana (a otpremnice nisu vec prijavljene)
    vec = pokr | pg.pokriveno[K_DATUM] | pg.pokriveno[K_OTP_DUPLI]
    for B, zs in sorted(pg.zbr_akt_po_b.items()):
        if B in pg.zbr_dupli or not (_id_set(zs, "VozacID") & pg.stanice):
            continue
        otps = pg.otp_akt_po_b.get(B, [])
        if not otps or (_id_set(otps, "OtpremnicaID") | _id_set(zs, "ZbirnaID")) & vec:
            continue
        st = {k(o, "StanicaID") for o in otps}
        dani = {d(o, "Datum") for o in otps if d(o, "Datum") is not None}
        zdani = {d(z, "Datum") for z in zs if d(z, "Datum") is not None}
        if len(st) < 2 and len(dani) < 2 and not (dani and zdani and dani != zdani):
            continue
        ctx.nalaz(K_CLAN, SUMNJA, ZBR, _ids(zs, "ZbirnaID"),
                  "Zbirna %s (%s, vozac %s) ima otpremnice %s: stanice %s, dani %s; malina je 1 otpremnica = "
                  "1 zbirna." % (s(zs[0], "BrojZbirne"), _dmy(_dan(zs)), s(zs[0], "VozacID"),
                                 _ids(otps, "OtpremnicaID"), "/".join(sorted(x.upper() for x in st)),
                                 "/".join(_dmy(x) for x in sorted(dani))),
                  "Svaku otpremnicu vratiti na njenu zbirnu (S + broj otpremnice), pa rekalkulisati.")


# ---------------------------------------------------------------- 12) stornirana + aktivna istog broja
def _zbr_isti_broj(pg):
    ctx = pg.ctx
    if not pg.ima(K_ISTI, {ZBR: ["ZbirnaID", "Datum", "VozacID", "BrojZbirne", "Stornirano"],
                           SV: ["CorrectionID", "Mode", "OldDocType", "OldBroj"]}):
        return
    dupli = defaultdict(list)
    for c in pg.sv:
        if k(c, "Mode") == "dupli_fantom" and k(c, "OldDocType") == "zbirna" and k(c, "OldBroj"):
            dupli[k(c, "OldBroj")].append(c)
    for B, cs in sorted(dupli.items()):
        zall = pg.zbr_po_b.get(B, [])
        akt = [z for z in zall if not storno(z)]
        sto = [z for z in zall if storno(z)]
        par = [(a, x) for a in akt for x in sto
               if k(a, "VozacID") == k(x, "VozacID") and d(a, "Datum") == d(x, "Datum")]
        if not par:
            continue
        a_ids = {k(a, "ZbirnaID") for a, _x in par}
        akt_p = [a for a in akt if k(a, "ZbirnaID") in a_ids]
        sto_p = [x for _a, x in par]
        otps = pg.otp_akt_po_b.get(B, [])
        ctx.nalaz(K_ISTI, SUMNJA, ZBR, _ids(akt_p + sto_p, "ZbirnaID"),
                  "Zbirna %s je stornirana kao dupli/fantom (%s), a auto-zbirna je ponovo napravila isti broj: aktivna "
                  "%s (%s, %s), stornirana %s%s." % (
                      s(akt_p[0], "BrojZbirne"), ", ".join(s(c, "CorrectionID") for c in cs), _ids(akt_p, "ZbirnaID"),
                      _dmy(d(akt_p[0], "Datum")), s(akt_p[0], "VozacID") or "-", _ids(sto_p, "ZbirnaID"),
                      ("; otpremnice tog broja: %s" % _ids(otps, "OtpremnicaID")) if otps else ""),
                  "Ako je stornirana bila fantom (ne sme da postoji): stornirati i %s i isprazniti BrojZbirne na %s; "
                  "ako je bila duplikat: stanje je ispravno -- pri migraciji vezati otpremnicu i prijemnicu za "
                  "AKTIVNU %s." % (_ids(akt_p, "ZbirnaID"), _ids(otps, "OtpremnicaID") or "otpremnici",
                                   _ids(akt_p, "ZbirnaID")))


_PROVERE = (_zbr_dupli, _otp_dupli, _otp_vozac, _tura_po_vozacu, _otp_datum, _otp_tudja, _zbr_stornirane_otp,
            _otp_ceka, _otp_bez_zbirne, _hladnjaca_lanac, _zbr_tura, _zbr_clan_van, _zbr_isti_broj)


def provere(ctx):
    pg = _Pogled(ctx)
    for f in _PROVERE:
        try:
            f(pg)
        except Exception:  # noqa: BLE001 -- pad jedne provere ne sme da obori ostale
            ctx.nalaz("MODUL_PAO", GRESKA, "", "otpremnica_zbirna" + f.__name__, traceback.format_exc(limit=6))
