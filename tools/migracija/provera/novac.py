"""Provere novca u staroj svesci (C002, vba-v2.2.3 .. vba-v2.28.4) -- SAMO CITA.

tblNovac je knjiga (jedan red = jedna transakcija). Jedini pisac je SaveNovac
(modNovac.bas), a izvor reda se vidi po tragu koji pozivalac ostavlja:

  IZVOD        Napomena "BIM:<id>; ...; Match:<Kupac|Kooperant|Kooperant-visak|OM>"
               (BuildBIMNapomena, v2.11.0 modBankaMapiranje.bas:1766-1782). Red
               kooperanta ima EntitetTip = "OM", PartnerID = OMID = stanica iz
               sifarnika (:778, :797-801); BrojDokumenta = broj izvoda ili "IZVOD".
  SPLIT        Napomena "Avans raspodela": deo avansa potrosen na otkup/fakturu;
               original se SMANJUJE za taj iznos (ApplyAvansToOtkup modNovac.bas
               :1112-1166, ApplyAvansToFaktura :563-621). Split nosi isti
               BrojDokumenta, Datum i KooperantID (kupac: PartnerID) kao original.
  OTKUP_FORMA  kes uz otkup: EntitetTip "Kooperant", Tip KesOtkupacKoop, BrojDokumenta
               = broj bloka, OtkupID = primarni red (modOtkup.bas:255-279).
  OM_ULAZ      rucni "Ulaz OM": EntitetTip "OM", Tip po prekidacima (frmDokumenta.frm
               :1722-1774, SaveOMUlaz_TX :1286-1303).
  KUPAC_RUCNO  uplata kupca: EntitetTip "Kupac", Napomena "Uplata po fakturi: ..." /
               "Avans kupca" (frmDokumenta.frm:2749-2785).

Klase (kod nalaza = ID klase iz specifikacije spec_mod_novac.md):

  NOV_DUPLA_ISPLATA_RUCNO_I_IZVOD          rucna isplata + ista isplata kroz izvod
  NOV_ISPLATA_OSIROCENA_STORNOM            isplata bez OtkupID (ili na storniranom
                                           otkupu) posle storna; naslednik DUPLA/PREVEZI/
                                           PRAVO_SIROCE
  NOV_ISPLATA_ODVEZANA_OD_STORNIRANOG_OTKUPA  podklasa DUPLA KES: odvezan kes, naslednik
                                           ima svoj kes istog iznosa
  NOV_PREPLATA_OTKUPA                      zbir isplata > vrednost (po dokumentu)
  OTK_ISPLACENO_POSLE_STORNA_NOVCA         Isplaceno = Da, a isplata stornirana
  NOV_VEZA_NA_TUDJI_ILI_STORNIRAN_DOKUMENT OtkupID/FakturaID tudjeg / nepostojeceg
  FAK_UPLATA_NA_STORNIRANU_FAKTURU         uplata na storniranu / nepostojecu fakturu
  NOV_UPLATA_KUPCA_NA_STORNIRANU_ILI_PREPLACENU_FAKTURU  preplacena faktura; KupciUplata
                                           bez fakture
  NOV_UPLATA_KUPCA_ODVEZANA_STORNOM_FAKTURE  spisak avansa kupca posle storna fakture (INFO)
  NOV_IZVOD_RUCNO_KUPAC_OM_POGRESNO        poziv = faktura, a knjizeno kao avans; istoimeni
                                           partneri; isplata kupcu
  NOV_IZVOD_POGRESAN_KOOPERANT             konto / poziv / PartnerMap pokazuju drugog
  NOV_TIP_ISPLATE_PO_PODRAZUMEVANOJ        rucni "virman" bez izvoda; saldo OM; kes OM bez
                                           primaoca
  NOV_ZASTAREO_DATUM_FORME                 isplata datirana pre otkupa / uplata pre fakture
  NOV_IZVOD_DATUM_LOKAL                    datum izvodnog reda != datum transakcije
  NOV_STORNO_NOVCA_BEZ_LANCA               storno originala ostavio split; stavka "Da" bez
                                           reda u knjizi; zbir lanca != izvod
  NOV_AVANS_RASPODELA_RAZDVOJENA_STORNOM   storniran split, original aktivan, zbir manji
  NOV_DUPLI_RUCNI_UNOS                     ista rucna isplata/uplata dvaput
  IZVOD_VAN_KNJIGE                         stavka Skip/Error/'' a novac stvarno otisao
  IZVOD_BEZ_POZIVA_ODE_U_AVANS             izvod bez poziva u avans dok stari blok ceka
  REVERS_STORNIRAN_NOVAC_OSTAO             storno reversa ostavio novac istog broja
  UPLATA_OM_NEVIDLJIVA_U_SALDU             izvodna uplata OM-a (saldo je ne vidi)
  POZICIONI_POMAK_TBLNOVAC                 zaglavlje / vrednosti ne odgovaraju pozicionom
                                           piscu
  IZVOD_DEDUPE_NEPROVERLJIV                (INFO) broj racuna u izvodima i rupe u nizu

Pravila: tabela koje nema ili je prazna -> provera se tiho preskace; tabela
postoji a kolona fali -> jedan INFO "<KOD>_PRESKOCENO". Jedan nalaz po
transakciji / dokumentu po klasi.
"""

import re
import traceback
from collections import defaultdict
from datetime import date, timedelta
from decimal import Decimal

from citac import broj as _broj
from citac import datum as _datum
from citac import kljuc, polje, tabela, tekst

from provera.zajednicko import GRESKA, INFO, SUMNJA, aktivni, cifre, d, dan_txt, k, kg_txt, n, s, storno

TOL = Decimal(1)              # 1 RSD
TOL_PCT = Decimal("0.01")     # 1 % (podklasa DUPLA KES)
MAX_ID = 12

# Tip (modConfig.bas:463-470), casefold
T_BANKA_UPL = "bankauplata"
T_BANKA_ISPL = "bankaisplata"
T_KUP_UPL = "kupciuplata"
T_KUP_AV = "kupciavans"
T_KES_FO = "kesfirmaotkupac"
T_KES_OK = "kesotkupackoop"
T_VIR_FK = "virmanfirmakoop"
T_VIR_AV = "virmanavanskoop"
TIPOVI = {T_BANKA_UPL, T_BANKA_ISPL, T_KUP_UPL, T_KUP_AV, T_KES_FO, T_KES_OK, T_VIR_FK, T_VIR_AV}
TIP_ISPLATE_KOOP = {T_VIR_FK, T_VIR_AV, T_KES_OK}

IZVOD, SPLIT, OTKUP_FORMA, OM_ULAZ, KUPAC_RUCNO, OSTALO = (
    "IZVOD", "SPLIT", "OTKUP_FORMA", "OM_ULAZ", "KUPAC_RUCNO", "OSTALO")
RUCNI = (OTKUP_FORMA, OM_ULAZ, KUPAC_RUCNO, OSTALO)

SPLIT_NAPOMENA = "avans raspodela"
M_KUPAC, M_KOOP, M_VISAK, M_OM = "kupac", "kooperant", "kooperant-visak", "om"

# Pozicioni pisac SaveNovac (modNovac.bas v2.11.0:198-201, v2.28.4:219-222)
OCEK_NOV = ["novacid", "brojdokumenta", "datum", "partner", "partnerid", "entitettip", "omid", "kooperantid",
            "fakturaid", "vrstavoca", "tip", "uplata", "isplata", "napomena", "stornirano", "otkupid",
            "osirocenood"]
OCEK_PM = ["bankaname", "partnerid", "entitettip", "omid"]      # savePartnerMap modNovac.bas:320

NOV_BAZA = ["NovacID", "Datum", "EntitetTip", "KooperantID", "Tip", "Isplata", "Napomena", "Stornirano",
            "OtkupID"]
OTK_BAZA = ["OtkupID", "KooperantID", "BrojDokumenta", "Datum", "Kolicina", "Cena", "Stornirano"]

DOK_REVERS_KOOP = ("om-izlaz-koop", "om-ulaz-koop")      # modConfig.bas:418-419

_BIM = re.compile(r"^\s*BIM:\s*([^;]*)", re.IGNORECASE)
_DMY = re.compile(r"^\s*(\d{1,2})\.(\d{1,2})\.(\d{4})\.?\s*$")
_OTK_ID = re.compile(r"^otk-\d+$")


# --- normalizacije kao u starom VBA (modBankaMapiranje.bas v2.11.0) -------------------
def poziv_kljuc(v):
    """NormalizePozivKey (:2033-2058): skini [NN] model, ostavi samo [0-9A-Z]."""
    r = tekst(v).upper()
    while True:
        p1 = r.find("[")
        if p1 < 0:
            break
        p2 = r.find("]", p1)
        if p2 < 0:
            break
        r = r[:p1] + r[p2 + 1:]
    return "".join(ch for ch in r if ("0" <= ch <= "9") or ("A" <= ch <= "Z"))


def konto_kljuc(v):
    """NormalizeKonto (:1971-2003): 3+13+2 cifre sa vodecim nulama, inace sve cifre."""
    raw = tekst(v)
    if not raw:
        return ""
    g = re.findall(r"[0-9]+", raw)
    if len(g) == 3 and len(g[0]) <= 3 and len(g[1]) <= 13 and len(g[2]) <= 2:
        return g[0].zfill(3) + g[1].zfill(13) + g[2].zfill(2)
    return "".join(re.findall(r"[0-9]", raw))


def loose(v):
    """NormalizeLooseBIM (:1784-1797)."""
    t = tekst(v).upper().replace("\t", " ")
    for ch in ",.-/":
        t = t.replace(ch, " ")
    return " ".join(t.split())


def _seg(nap, ime):
    """Vrednost segmenta '<ime>:' iz Napomena izvodnog reda ('' ako ga nema)."""
    if ime == "Match":
        p = nap.rfind("; Match:")
        return tekst(nap[p + 8:]) if p >= 0 else ""
    m = re.search(r"(?:^|;\s*)" + ime + r":([^;]*)", nap)
    return tekst(m.group(1)) if m else ""


def _rsd(v):
    return kg_txt(v)


def _lista(vals):
    out, vid = [], set()
    for v in vals:
        v = tekst(v)
        if v and v.casefold() not in vid:
            vid.add(v.casefold())
            out.append(v)
    if len(out) > MAX_ID:
        return ", ".join(out[:MAX_ID]) + " (+%d)" % (len(out) - MAX_ID)
    return ", ".join(out)


def _vr(o):
    """Vrednost reda otkupa Kolicina * Cena (Decimal)."""
    return n(o, "Kolicina") * n(o, "Cena")


def _u_prozoru(x, od, do):
    return x is not None and od is not None and do is not None and od <= x <= do


def _dani(a, b):
    return abs((a - b).days) if a and b else 9999


# --- red novca ----------------------------------------------------------------------
class _NR:
    __slots__ = ("r", "id", "idk", "idn", "tip", "tip_txt", "nap", "bim", "match", "etip", "koop", "koop_txt",
                 "pid", "pid_txt", "omid", "omid_txt", "otk", "otk_txt", "fak", "fak_txt", "isp", "upl", "dan",
                 "broj", "broj_txt", "st", "izvor")

    def __init__(self, r):
        self.r = r
        self.id = s(r, "NovacID")
        self.idk = kljuc(self.id)
        self.idn = cifre(self.id) or 0
        self.tip = k(r, "Tip")
        self.tip_txt = s(r, "Tip")
        self.nap = s(r, "Napomena")
        m = _BIM.match(self.nap)
        self.bim = kljuc(m.group(1)) if m else ""
        self.match = kljuc(_seg(self.nap, "Match")) if m else ""
        self.etip = k(r, "EntitetTip")
        self.koop, self.koop_txt = k(r, "KooperantID"), s(r, "KooperantID")
        self.pid, self.pid_txt = k(r, "PartnerID"), s(r, "PartnerID")
        self.omid, self.omid_txt = k(r, "OMID"), s(r, "OMID")
        self.otk, self.otk_txt = k(r, "OtkupID"), s(r, "OtkupID")
        self.fak, self.fak_txt = k(r, "FakturaID"), s(r, "FakturaID")
        self.isp = n(r, "Isplata")
        self.upl = n(r, "Uplata")
        self.dan = d(r, "Datum")
        self.broj, self.broj_txt = k(r, "BrojDokumenta"), s(r, "BrojDokumenta")
        self.st = storno(r)
        if self.bim:
            self.izvor = IZVOD
        elif self.nap.casefold() == SPLIT_NAPOMENA:
            self.izvor = SPLIT
        elif self.etip == "kooperant" and self.tip == T_KES_OK:
            self.izvor = OTKUP_FORMA
        elif self.etip == "om":
            self.izvor = OM_ULAZ
        elif self.etip == "kupac":
            self.izvor = KUPAC_RUCNO
        else:
            self.izvor = OSTALO

    def iznos(self):
        return self.isp + self.upl

    def av_kljuc(self):
        """Kljuc lanca avansa: original i njegovi split redovi dele Tip, entitet, broj i datum."""
        if self.tip == T_VIR_AV:
            ent = self.koop
        elif self.tip == T_KUP_AV:
            ent = self.pid
        else:
            return None
        if not ent:
            return None
        return (self.tip, ent, self.broj, self.dan.isoformat() if self.dan else "")

    def kratko(self):
        return "%s (%s, %s, %s RSD%s)" % (self.id, self.tip_txt or "?", dan_txt(self.dan), _rsd(self.iznos()),
                                          ", " + self.otk_txt if self.otk_txt else "")


# --- kontekst modula ------------------------------------------------------------------
class _P:
    def __init__(self, ctx):
        self.ctx = ctx
        self._presk = set()
        self.nov = [_NR(r) for r in (ctx.tab("tblNovac") or [])]
        self.akt = [x for x in self.nov if not x.st]
        self.po_id = {}
        self.po_bim = defaultdict(list)
        self.po_otk = defaultdict(list)
        self.po_fak = defaultdict(list)
        self.av_kor = defaultdict(list)
        self.av_spl = defaultdict(list)
        for x in self.nov:
            self.po_id.setdefault(x.idk, x)
            if x.bim:
                self.po_bim[x.bim].append(x)
            if x.otk:
                self.po_otk[x.otk].append(x)
            if x.fak:
                self.po_fak[x.fak].append(x)
            kk = x.av_kljuc()
            if kk:
                (self.av_spl if x.izvor == SPLIT else self.av_kor)[kk].append(x)
        self.placeno = defaultdict(Decimal)
        for x in self.akt:
            if x.otk:
                self.placeno[x.otk] += x.isp
        self.uplaceno = defaultdict(Decimal)
        for x in self.akt:
            if x.fak:
                self.uplaceno[x.fak] += x.upl
        self.otk = ctx.tab("tblOtkup") or []
        self.otk_po_id = ctx.po_id.get("tblOtkup", {})
        self.bim = ctx.tab("tblBankaImport") or []
        self.bim_po_id = {}
        for b in self.bim:
            if k(b, "BankaImportID"):
                self.bim_po_id.setdefault(k(b, "BankaImportID"), b)
        self.fak_rows = ctx.tab("tblFakture") or []
        self.fak_po_id = ctx.po_id.get("tblFakture", {})
        self.reported = defaultdict(set)      # kod -> kljucevi vec prijavljenih (dedupe medju klasama)
        self._otk_idx = None
        self._dok = None
        self._mat = None
        self._rucni = None

    # --- citanje / zaglavlja ---
    def redovi(self, ime):
        return self.ctx.tab(ime) or []

    def spremno(self, kod, zahtevi):
        """True kad su sve tabele neprazne i imaju kolone. Prazna/nepostojeca tabela -> tiho
        False; kolona fali -> jedan INFO <kod>_PRESKOCENO."""
        for ime, _kol in zahtevi:
            if not self.redovi(ime):
                return False
        fale = ["%s.%s" % (ime, kol) for ime, kolone in zahtevi for kol in kolone
                if not self.ctx.ima_kolonu(ime, kol)]
        if not fale:
            return True
        if kod not in self._presk:
            self._presk.add(kod)
            self.ctx.nalaz(kod + "_PRESKOCENO", INFO, fale[0].split(".")[0], "",
                           "Provera %s preskocena: nedostaje kolona %s." % (kod, ", ".join(fale)),
                           "Proveriti zaglavlje tabele u svesci.")
        return False

    def ima(self, ime, kol):
        return bool(self.redovi(ime)) and self.ctx.ima_kolonu(ime, kol)

    # --- otkup ---
    def otk_idx(self):
        if self._otk_idx is None:
            po_poziv = defaultdict(list)      # poziv kljuc -> svi redovi (i stornirani)
            po_koop_broj = defaultdict(list)  # (koop, broj) -> svi redovi
            po_koop = defaultdict(list)
            for o in self.otk:
                kk = k(o, "KooperantID")
                pk = poziv_kljuc(polje(o, "BrojDokumenta"))
                if pk:
                    po_poziv[pk].append(o)
                if k(o, "BrojDokumenta"):
                    po_koop_broj[(kk, k(o, "BrojDokumenta"))].append(o)
                po_koop[kk].append(o)
            self._otk_idx = (po_poziv, po_koop_broj, po_koop)
        return self._otk_idx

    def otk_po_pozivu(self, poziv, koop=None, aktivni_samo=True):
        pk = poziv_kljuc(poziv)
        if not pk:
            return []
        out = self.otk_idx()[0].get(pk, [])
        if koop is not None:
            out = [o for o in out if k(o, "KooperantID") == koop]
        if aktivni_samo:
            out = [o for o in out if not storno(o)]
        return out

    def otvoreno(self, o):
        return _vr(o) - self.placeno.get(k(o, "OtkupID"), Decimal(0))

    def dok(self, o):
        """(vrednost, placeno, redovi) logickog dokumenta otkupa (Klasa I + II istog broja). SaveOtkupMulti_TX
        knjizi ceo kes na primarni red, pa je red Klase II 'otvoren' iako je dokument placen."""
        if self._dok is None:
            g = defaultdict(list)
            for r in aktivni(self.otk):
                if k(r, "OtkupID"):
                    g[_dok_kljuc(r)].append(r)
            self._dok = {}
            for kl, redovi in g.items():
                V = sum((_vr(r) for r in redovi), Decimal(0))
                Pd = sum((self.placeno.get(k(r, "OtkupID"), Decimal(0)) for r in redovi), Decimal(0))
                self._dok[kl] = (V, Pd, redovi)
        if storno(o):
            return _vr(o), self.placeno.get(k(o, "OtkupID"), Decimal(0)), [o]
        return self._dok.get(_dok_kljuc(o), (_vr(o), self.placeno.get(k(o, "OtkupID"), Decimal(0)), [o]))

    def otvoreno_dok(self, o):
        V, Pd, _r = self.dok(o)
        return V - Pd

    # --- lanac izvoda ---
    def lanac(self, bid):
        """(direktni redovi izvoda, split redovi lanca, dvosmislen). Split se pripisuje izvodu samo
        kad u kljucu avansa postoji JEDAN original (broj izvoda deli vise transakcija)."""
        direkt = [x for x in self.po_bim.get(bid, ()) if x.izvor == IZVOD]
        spl, dvo, vid = [], False, set()
        for x in direkt:
            kk = x.av_kljuc()
            if not kk or kk in vid:
                continue
            vid.add(kk)
            if self.av_spl.get(kk):
                if len(self.av_kor[kk]) == 1:
                    spl.extend(self.av_spl[kk])
                else:
                    dvo = True
        return direkt, spl, dvo

    def splitovi(self, x):
        """Split redovi originala x (None kad pripadnost nije jednoznacna)."""
        kk = x.av_kljuc()
        if not kk:
            return []
        if len(self.av_kor.get(kk, ())) != 1:
            return None if self.av_spl.get(kk) else []
        return self.av_spl.get(kk, [])

    def iznos_lanca(self, x):
        sp = self.splitovi(x) if x.izvor != SPLIT else []
        return x.iznos() + sum((y.iznos() for y in (sp or []) if not y.st), Decimal(0))

    # --- maticni ---
    def mat(self):
        if self._mat is None:
            koop_konto, koop_ime, kup_konto, kup_ime, st_ime = (defaultdict(set) for _ in range(5))
            koop = {}
            for r in self.redovi("tblKooperanti"):
                kid = k(r, "KooperantID")
                if not kid:
                    continue
                koop[kid] = r
                kt = konto_kljuc(polje(r, "TekuciRacun"))
                if kt:
                    koop_konto[kt].add(kid)
                ime = loose("%s %s" % (s(r, "Ime"), s(r, "Prezime")))
                if ime:
                    koop_ime[ime].add(kid)
                ime2 = loose("%s %s" % (s(r, "Prezime"), s(r, "Ime")))
                if ime2:
                    koop_ime[ime2].add(kid)
            kup = {}
            for r in self.redovi("tblKupci"):
                kid = k(r, "KupacID")
                if not kid:
                    continue
                kup[kid] = r
                kt = konto_kljuc(polje(r, "TekuciRacun"))
                if kt:
                    kup_konto[kt].add(kid)
                if loose(polje(r, "Naziv")):
                    kup_ime[loose(polje(r, "Naziv"))].add(kid)
            st = {}
            for r in self.redovi("tblStanice"):
                sid = k(r, "StanicaID")
                if sid:
                    st[sid] = r
                    if loose(polje(r, "Naziv")):
                        st_ime[loose(polje(r, "Naziv"))].add(sid)
            pm = {}
            for r in self.redovi("tblPartnerMap"):
                b = k(r, "BankaName")
                if b and b not in pm:
                    pm[b] = r
            self._mat = dict(koop=koop, koop_konto=koop_konto, koop_ime=koop_ime, kup=kup, kup_konto=kup_konto,
                             kup_ime=kup_ime, st=st, st_ime=st_ime, pm=pm)
        return self._mat

    def partner_bim(self, b):
        """(vrsta, id, kako) za stavku izvoda: konto -> PartnerMap -> ime; None ako nije jednoznacno."""
        m = self.mat()
        kt = konto_kljuc(polje(b, "PartnerKonto"))
        isplata = n(b, "Isplata") > 0
        if kt:
            if isplata and len(m["koop_konto"].get(kt, ())) == 1:
                return ("kooperant", next(iter(m["koop_konto"][kt])), "konto")
            if not isplata and len(m["kup_konto"].get(kt, ())) == 1:
                return ("kupac", next(iter(m["kup_konto"][kt])), "konto")
        pm = m["pm"].get(k(b, "Partner"))
        if pm is not None and k(pm, "PartnerID"):
            et = k(pm, "EntitetTip")
            vrsta = {"kooperant": "kooperant", "kupac": "kupac", "om": "om"}.get(et)
            if vrsta:
                return (vrsta, k(pm, "PartnerID"), "PartnerMap")
        ime = loose(polje(b, "Partner"))
        if ime:
            hits = [("kooperant", x) for x in m["koop_ime"].get(ime, ())] + \
                   [("kupac", x) for x in m["kup_ime"].get(ime, ())] + \
                   [("om", x) for x in m["st_ime"].get(ime, ())]
            if len(hits) == 1:
                return (hits[0][0], hits[0][1], "ime")
        return None

    def rucni(self):
        """(vrsta, id) -> aktivni rucni redovi (nisu izvod ni split)."""
        if self._rucni is None:
            g = defaultdict(list)
            for x in self.akt:
                if x.izvor not in RUCNI:
                    continue
                if x.koop:
                    g[("kooperant", x.koop)].append(x)
                elif x.etip == "kupac" and x.pid:
                    g[("kupac", x.pid)].append(x)
                elif x.omid:
                    g[("om", x.omid)].append(x)
            self._rucni = g
        return self._rucni

    def rucna_zamena(self, vrsta, pid, isplata, iznos, dan, posle=10):
        """Rucni red istog partnera i iznosa (+-1) u [dan-3, dan+posle], najblizi prvi."""
        if not dan:
            return []
        out = []
        for x in self.rucni().get((vrsta, pid), ()):
            a = x.isp if isplata else x.upl
            if a > 0 and abs(a - iznos) <= TOL and _u_prozoru(x.dan, dan - timedelta(days=3),
                                                               dan + timedelta(days=posle)):
                out.append(x)
        return sorted(out, key=lambda x: (_dani(x.dan, dan), x.idn))

    def bim_dan(self, b, direkt=()):
        dd = d(b, "DatumTransakcije")
        if dd is None:
            for x in direkt:
                if x.dan:
                    return x.dan
        return dd


def _otk_txt(o):
    return "%s (%s, %s, %s, vrednost %s)" % (s(o, "OtkupID"), s(o, "BrojDokumenta"), s(o, "KooperantID"),
                                             dan_txt(d(o, "Datum")), _rsd(_vr(o)))


def _bim_txt(b):
    return "%s (izvod %s, %s, %s, %s RSD)" % (
        s(b, "BankaImportID"), s(b, "BrojDokumenta") or "-", dan_txt(polje(b, "DatumTransakcije")),
        s(b, "Partner") or "-", _rsd(n(b, "Isplata") + n(b, "Uplata")))


# =====================================================================================
# 1) NOV_DUPLA_ISPLATA_RUCNO_I_IZVOD
# =====================================================================================
def _dupla(P):
    K = "NOV_DUPLA_ISPLATA_RUCNO_I_IZVOD"
    if not P.spremno(K, [("tblNovac", NOV_BAZA + ["BrojDokumenta", "OMID"]),
                         ("tblBankaImport", ["BankaImportID", "Isplata", "Obradjeno"])]):
        return
    rucni_koop = defaultdict(list)
    rucni_om = defaultdict(list)
    for x in P.akt:
        if x.izvor != OM_ULAZ or x.isp <= 0 or not x.dan:
            continue
        if x.koop and x.tip in TIP_ISPLATE_KOOP:
            rucni_koop[x.koop].append(x)
        elif not x.koop and x.tip == T_KES_FO and x.omid:
            rucni_om[x.omid].append(x)
    parovi = []
    for b in P.bim:
        bid = k(b, "BankaImportID")
        A = n(b, "Isplata")
        if storno(b) or k(b, "Obradjeno") != "da" or A <= 0 or not bid:
            continue
        direkt, spl, dvo = P.lanac(bid)
        direkt = [x for x in direkt if not x.st]
        if not direkt:
            continue
        D = P.bim_dan(b, direkt)
        if D is None:
            continue
        koops = {x.koop for x in direkt if x.koop}
        if len(koops) == 1:
            K_ = next(iter(koops))
            blok = P.otk_po_pozivu(polje(b, "PozivNaBroj"), K_)
            blok_ids = {k(o, "OtkupID") for o in blok}
            visak_red = [x for x in direkt if x.match == M_VISAK and x.tip == T_VIR_AV]
            visak = sum((P.iznos_lanca(x) for x in visak_red), Decimal(0))
            for M in rucni_koop.get(K_, ()):
                if not _u_prozoru(M.dan, D - timedelta(days=10), D + timedelta(days=30)):
                    continue
                izn = P.iznos_lanca(M) if M.tip == T_VIR_AV else M.isp
                pun = abs(izn - A) <= TOL
                deo = visak > TOL and bool(M.otk) and M.otk in blok_ids
                if not (pun or deo):
                    continue
                if pun and blok_ids and M.otk and M.otk not in blok_ids:
                    continue        # izvod i rucni placaju RAZLICITE blokove (legitimno)
                jak = bool(M.otk) and M.otk in blok_ids
                parovi.append(((jak, pun, -_dani(M.dan, D)), b, M, direkt, spl, dvo, blok, visak, "koop"))
        elif not koops:
            oms = {x.omid for x in direkt if x.tip == T_KES_FO and x.isp > 0 and x.omid}
            if len(oms) != 1:
                continue
            om = next(iter(oms))
            for M in rucni_om.get(om, ()):
                if abs(M.isp - A) <= TOL and _u_prozoru(M.dan, D - timedelta(days=10), D + timedelta(days=10)):
                    parovi.append(((False, True, -_dani(M.dan, D)), b, M, direkt, spl, dvo, [], Decimal(0), "om"))
    parovi.sort(key=lambda p: p[0], reverse=True)
    uzeti_m, uzeti_b = set(), set()
    for (jak, pun, _bl), b, M, direkt, spl, dvo, blok, visak, var in parovi:
        bid = k(b, "BankaImportID")
        if M.idk in uzeti_m or bid in uzeti_b:
            continue
        uzeti_m.add(M.idk)
        uzeti_b.add(bid)
        P.reported[K].add(bid)
        A = n(b, "Isplata")
        lazno = [y.otk_txt for y in direkt + spl if not y.st and y.otk and y.otk != M.otk]
        lanac_txt = _lista([y.id for y in direkt + spl if not y.st])
        if var == "om":
            ozb = SUMNJA
            opis = ("Rucni KesFirmaOtkupac %s (OM %s, %s, %s RSD) i izvodni KesFirmaOtkupac %s za isti OM i isti "
                    "iznos (stavka %s, %s)." % (M.id, M.omid_txt, dan_txt(M.dan), _rsd(M.isp), lanac_txt,
                                                s(b, "BankaImportID"), dan_txt(P.bim_dan(b, direkt))))
            predlog = ("Ako je rucni red knjizenje iste uplate OM-u: storno %s (izvodni red ostaje). Ako su to dve "
                       "isplate (fiksni dnevni avans), ostaviti. Odluka operatera." % M.id)
        else:
            # GRESKA samo kad podaci dokazuju da je blok vec bio zatvoren: izvod sa pozivom na M-ov blok je
            # knjizen kao avans (Match:Kooperant, GetOtkupCandidatesForKooperantBlock preskace zatvoren blok),
            # ili je blok posle obe isplate preplacen. Izvod koji je placao sam blok (VirmanFirmaKoop sa
            # OtkupID) znaci da je blok u trenutku mapiranja bio otvoren: dve isplate po pola su moguce.
            avans_zatvoren = any(y.tip == T_VIR_AV and y.match == M_KOOP for y in direkt)
            preplacen = False
            if M.otk and M.otk in P.otk_po_id:
                Vd, Pd, _rr = P.dok(P.otk_po_id[M.otk])
                preplacen = Pd - Vd > TOL
            dokazano = avans_zatvoren or preplacen
            ozb = GRESKA if (jak and pun and M.tip in (T_VIR_FK, T_VIR_AV) and dokazano) else SUMNJA
            vrsta = "pun" if pun else "delimican (izvod imao visak %s RSD, preklop %s RSD)" % (
                _rsd(visak), _rsd(min(visak, M.isp)))
            opis = ("Rucna isplata %s (%s, kooperant %s, %s, %s RSD%s) i izvod %s (lanac %s) placaju isto: %s "
                    "duplikat%s." % (M.id, M.tip_txt, M.koop_txt, dan_txt(M.dan), _rsd(M.isp),
                                     ", blok " + M.otk_txt if M.otk_txt else "", _bim_txt(b), lanac_txt, vrsta,
                                     "; JAK POGODAK: poziv %s pokazuje na isti blok" % s(b, "PozivNaBroj")
                                     if jak else ""))
            if jak and pun and not dokazano:
                opis += (" Izvod je placao sam blok (blok bio otvoren pri mapiranju) i blok nije preplacen: moguce "
                         "su i dve stvarne isplate po delovima.")
            if dvo:
                opis += " Split redovi lanca nisu jednoznacni (vise avansa istog kooperanta na istom izvodu)."
            if lazno:
                opis += " Izvodni avans je potrosen na: %s (lazno placeno ako je rucni red duplikat)." % _lista(lazno)
            predlog = ("Izvodni red ostaje (dokaz da je novac otisao). Kandidat za storno: %s; izvodni lanac "
                       "prevezati na %s%s. BankaImportID %s. Odluka operatera." % (
                           M.id, M.otk_txt or "blok iz poziva",
                           ("; vratiti u otvorene: " + _lista(lazno)) if lazno else "", s(b, "BankaImportID")))
        P.ctx.nalaz(K, ozb, "tblNovac", M.id, opis, predlog)


# =====================================================================================
# 2) NOV_ISPLATA_OSIROCENA_STORNOM + NOV_ISPLATA_ODVEZANA_OD_STORNIRANOG_OTKUPA
# =====================================================================================
def _zurnal_otkup(P):
    """NovacID -> [stari OtkupID] iz tblStornoZurnal (Kolona OtkupID obrisana stornom, v2.28.x)."""
    out = defaultdict(list)
    z = P.redovi("tblStornoZurnal")
    if not z:
        return out
    for kol in ("Tabela", "RowID", "Kolona", "StaraVrednost"):
        if not P.ctx.ima_kolonu("tblStornoZurnal", kol):
            return out
    for r in z:
        if k(r, "Tabela") == "tblnovac" and k(r, "Kolona") == "otkupid" and k(r, "StaraVrednost") \
                and not k(r, "NovaVrednost"):
            out[k(r, "RowID")].append(s(r, "StaraVrednost"))
    return out


def _nadji_o(P, x, zur):
    """(O, kako, kandidati). O = stornirani otkup koji je x placao."""
    po_poziv, po_koop_broj, po_koop = P.otk_idx()
    for oid in reversed(zur.get(x.idk, [])):
        o = P.otk_po_id.get(kljuc(oid))
        if o is not None and storno(o):
            return o, "zurnal storna", []
    kand = []
    # (b) isti broj: samo kes uz otkup (BrojDokumenta = broj bloka). Ulaz OM nosi broj OM ulaza / reversa
    #     (isti oblik x/ddmmyy, niz nad tblAmbalaza), pa bi poklapanje sa brojem bloka bilo slucajno -> (d).
    if x.izvor not in (IZVOD, OM_ULAZ) and x.broj:
        kand = [o for o in po_koop_broj.get((x.koop, x.broj), ()) if storno(o)]
        if len(kand) > 1 and x.dan:
            isti = [o for o in kand if d(o, "Datum") == x.dan]
            kand = isti or kand
        if len(kand) == 1:
            return kand[0], "isti broj dokumenta", []
        if kand:
            return None, "", kand
    if x.izvor == IZVOD:
        b = P.bim_po_id.get(x.bim)
        if b is not None:
            kand = P.otk_po_pozivu(polje(b, "PozivNaBroj"), x.koop, aktivni_samo=False)
            kand = [o for o in kand if storno(o)]
            if len(kand) == 1:
                return kand[0], "poziv na broj izvoda", []
            if kand:
                return None, "", kand
    if x.izvor in (OM_ULAZ, OSTALO) and x.dan:
        kand = [o for o in po_koop.get(x.koop, ()) if storno(o) and d(o, "Datum") and d(o, "Datum") <= x.dan
                and _vr(o) >= x.isp - TOL]
        if len(kand) == 1:
            return kand[0], "jedini stornirani otkup kooperanta do datuma isplate (slab pogodak)", []
        if kand:
            return None, "", kand
    return None, "", []


def _naslednik(P, O):
    """Aktivan otkup istog kooperanta koji je zamenio O (isti broj, ili isti dan i stanica; noviji ID)."""
    _pp, po_koop_broj, po_koop = P.otk_idx()
    on = cifre(s(O, "OtkupID")) or 0
    kk = k(O, "KooperantID")
    isti_broj = [o for o in po_koop_broj.get((kk, k(O, "BrojDokumenta")), ()) if not storno(o)
                 and (cifre(s(o, "OtkupID")) or 0) > on] if k(O, "BrojDokumenta") else []
    if isti_broj:
        return sorted(isti_broj, key=lambda o: cifre(s(o, "OtkupID")) or 0)[0], True
    isti_dan = [o for o in po_koop.get(kk, ()) if not storno(o) and (cifre(s(o, "OtkupID")) or 0) > on
                and d(o, "Datum") and d(o, "Datum") == d(O, "Datum") and k(o, "StanicaID") == k(O, "StanicaID")]
    if isti_dan:
        return sorted(isti_dan, key=lambda o: cifre(s(o, "OtkupID")) or 0)[0], False
    return None, False


def _osirocena(P):
    K1 = "NOV_ISPLATA_OSIROCENA_STORNOM"
    K2 = "NOV_ISPLATA_ODVEZANA_OD_STORNIRANOG_OTKUPA"
    if not P.spremno(K1, [("tblNovac", NOV_BAZA + ["BrojDokumenta", "OMID"]),
                          ("tblOtkup", OTK_BAZA + ["StanicaID"])]):
        return
    zur = _zurnal_otkup(P)
    for x in P.akt:
        if x.isp <= 0 or not x.koop:
            continue
        o_link = P.otk_po_id.get(x.otk) if x.otk else None
        veza_storn = o_link is not None and storno(o_link)
        sirce = not x.otk and x.tip in (T_KES_OK, T_VIR_FK) and x.izvor != SPLIT
        if not (sirce or veza_storn):
            if not x.otk and x.tip == T_VIR_AV and zur.get(x.idk):
                O = P.otk_po_id.get(kljuc(zur[x.idk][-1]))
                if O is not None and storno(O):
                    P.ctx.nalaz(K1, INFO, "tblNovac", x.id,
                                "%s %s vracen u avans stornom otkupa %s (zurnal): legitimno, prikaz uz otkup." % (
                                    x.tip_txt, x.kratko(), _otk_txt(O)),
                                "Nista; avans ce potrositi sledeci otkup kooperanta %s." % x.koop_txt)
            continue
        if veza_storn:
            O, kako, kand = o_link, "OtkupID i dalje pokazuje na storniran otkup", []
        else:
            O, kako, kand = _nadji_o(P, x, zur)
        S, isti_broj = (_naslednik(P, O) if O is not None else (None, False))
        # podklasa DUPLA KES bez poznatog O: aktivan otkup istog kooperanta, ista stanica, +-1 dan, svoj kes +-1%
        if S is None and O is None and x.tip == T_KES_OK and x.dan:
            for o in P.otk_idx()[2].get(x.koop, ()):
                if storno(o) or k(o, "StanicaID") != x.omid or _dani(d(o, "Datum"), x.dan) > 1:
                    continue
                svoj = [y for y in P.po_otk.get(k(o, "OtkupID"), ()) if not y.st and y.tip == T_KES_OK and y is not x
                        and abs(y.isp - x.isp) <= max(TOL, x.isp * TOL_PCT)]
                if svoj:
                    S = o
                    break
        klasa, dup = "PRAVO_SIROCE", []
        if S is not None:
            sid = k(S, "OtkupID")
            dup = [y for y in P.po_otk.get(sid, ()) if not y.st and y is not x and y.tip == x.tip
                   and abs(y.isp - x.isp) <= max(TOL, x.isp * TOL_PCT if x.tip == T_KES_OK else TOL)]
            ost = P.otvoreno(S)
            if dup:
                klasa = "DUPLA"
            elif x.isp <= ost + TOL:
                klasa = "PREVEZI"
            else:
                klasa = "NASLEDNIK_PLACEN"
        elif O is None:
            klasa = "NEPOZNAT"
        o_txt = _otk_txt(O) if O is not None else (
            "nepoznat (kandidati: %s)" % _lista(s(o, "OtkupID") for o in kand) if kand else "nepoznat")
        s_txt = _otk_txt(S) if S is not None else "nema"
        if klasa == "DUPLA" and x.tip == T_KES_OK:
            kod = K2
            ozb = GRESKA if (isti_broj or (S is not None and k(S, "BrojDokumenta") == x.broj and x.broj)) else SUMNJA
        else:
            kod = K1
            ozb = GRESKA if veza_storn else SUMNJA
        opis = ("%s %s %s; stornirani otkup: %s [%s]; naslednik: %s; klasa %s." % (
            "Isplata" if not veza_storn else "Isplata vezana za STORNIRAN otkup", x.kratko(),
            "bez OtkupID" if not x.otk else "", o_txt, kako or "-", s_txt, klasa))
        if klasa == "DUPLA":
            opis += " Naslednik vec ima svoju isplatu istog iznosa: %s." % _lista(y.kratko() for y in dup)
            predlog = ("Dupla isplata: jedan od %s / %s je kandidat za storno ili potrazivanje od kooperanta%s. "
                       "Odluka operatera." % (x.id, _lista(y.id for y in dup),
                                              " (kes istog broja: verovatno unet dvaput)" if kod == K2 else ""))
        elif klasa == "PREVEZI":
            predlog = "OtkupID: %s -> %s (ostatak naslednika %s RSD, isplata %s RSD)." % (
                x.otk_txt or "''", s(S, "OtkupID"), _rsd(P.otvoreno(S)), _rsd(x.isp))
        elif klasa == "NASLEDNIK_PLACEN":
            predlog = ("Naslednik %s je vec placen (ostatak %s RSD); prevezivanje bi napravilo preplatu %s RSD. "
                       "Isplata ostaje dug kooperanta ili storno, odluka operatera." % (
                           s(S, "OtkupID"), _rsd(P.otvoreno(S)), _rsd(x.isp - P.otvoreno(S))))
        elif klasa == "PRAVO_SIROCE":
            predlog = ("Placen posao je storniran i nije ponovo unet: isplata ostaje potrazivanje od kooperanta %s "
                       "(prikazati, ne brisati). Odluka operatera." % x.koop_txt)
        else:
            predlog = ("Nije nadjen otkup koji je isplata placala (moguca i rucna izmena OtkupID u Excelu). "
                       "Operater bira otkup kooperanta %s ili ostavlja kao potrazivanje." % x.koop_txt)
        P.reported[K1].add(x.idk)
        P.ctx.nalaz(kod, ozb, "tblNovac", x.id, opis, predlog)


# =====================================================================================
# 3) NOV_PREPLATA_OTKUPA, OTK_ISPLACENO_POSLE_STORNA_NOVCA
# =====================================================================================
def _dok_kljuc(o):
    dd = d(o, "Datum")
    return (k(o, "BrojDokumenta") or "#" + k(o, "OtkupID"), k(o, "KooperantID"), dd.isoformat() if dd else "")


def _preplata(P):
    K = "NOV_PREPLATA_OTKUPA"
    if not P.spremno(K, [("tblOtkup", OTK_BAZA), ("tblNovac", ["OtkupID", "Isplata", "Stornirano"])]):
        return
    ima_novac = P.ctx.ima_kolonu("tblOtkup", "Novac")
    dok = defaultdict(list)
    for o in aktivni(P.otk):
        if k(o, "OtkupID"):
            dok[_dok_kljuc(o)].append(o)
    for _kl, redovi in dok.items():
        V = sum((_vr(o) for o in redovi), Decimal(0))
        Pd = sum((P.placeno.get(k(o, "OtkupID"), Decimal(0)) for o in redovi), Decimal(0))
        prepl = [o for o in redovi if P.placeno.get(k(o, "OtkupID"), Decimal(0)) - _vr(o) > TOL]
        nov_preko = ima_novac and sum((n(o, "Novac") for o in redovi), Decimal(0)) > V + TOL
        if not prepl and not (Pd - V > TOL) and not nov_preko:
            continue
        ids = _lista(s(o, "OtkupID") for o in redovi)
        isplate = [x for o in redovi for x in P.po_otk.get(k(o, "OtkupID"), ()) if not x.st]
        po_izv = defaultdict(Decimal)
        for x in isplate:
            po_izv[x.izvor] += x.isp
        izv_txt = ", ".join("%s %s" % (i, _rsd(v)) for i, v in sorted(po_izv.items()))
        o0 = redovi[0]
        glava = "Otkup %s (%s, kooperant %s, %s)" % (ids, s(o0, "BrojDokumenta") or "-", s(o0, "KooperantID"),
                                                     dan_txt(d(o0, "Datum")))
        if Pd - V > TOL:
            ozb = SUMNJA if V <= 0 else GRESKA
            opis = "%s: isplaceno %s RSD > vrednost %s RSD (preplata %s RSD); isplate po izvoru: %s." % (
                glava, _rsd(Pd), _rsd(V), _rsd(Pd - V), izv_txt or "-")
            if V <= 0:
                opis += " Kolicina ili cena je nula/prazna."
            if any(i in (OM_ULAZ, IZVOD) for i in po_izv):
                opis += " Preplatu cini i isplata iz Ulaza OM/izvoda: proveriti prevezivanje posle storna."
            predlog = ("Visak %s RSD je potrazivanje od kooperanta; ako je isplacen virmanom, odvojiti ga kao "
                       "VirmanAvansKoop bez OtkupID (kes ne moze u avans). Isplate: %s." % (
                           _rsd(Pd - V), _lista(x.id for x in isplate)))
        elif prepl:
            # SaveOtkupMulti_TX namerno knjizi ceo kes na primarni red (modOtkup.bas:255-279): ovo je
            # ponasanje stare aplikacije, ne greska unosa -> INFO (kontekst za raspodelu po klasi u migraciji).
            ozb = INFO
            o1 = prepl[0]
            drugi = [o for o in redovi if o is not o1]
            opis = ("%s: red %s (Klasa %s) placen %s RSD > svoja vrednost %s RSD, ali dokument nije preplacen "
                    "(%s / %s RSD): kes knjizen na primarni red (ponasanje SaveOtkupMulti_TX)." % (
                        glava, s(o1, "OtkupID"), s(o1, "Klasa") or "?", _rsd(P.placeno[k(o1, "OtkupID")]),
                        _rsd(_vr(o1)), _rsd(Pd), _rsd(V)))
            predlog = "Visak %s RSD sa %s prebaciti na %s (split isplate ili promena OtkupID)." % (
                _rsd(P.placeno[k(o1, "OtkupID")] - _vr(o1)), s(o1, "OtkupID"), _lista(s(o, "OtkupID") for o in drugi))
        else:
            ozb = INFO
            opis = "%s: polje Novac %s RSD > vrednost %s RSD, a isplate u tblNovac (%s RSD) nisu preko vrednosti." % (
                glava, _rsd(sum((n(o, "Novac") for o in redovi), Decimal(0))), _rsd(V), _rsd(Pd))
            predlog = "Proveriti stvarno isplacen kes; staro polje Novac se ne prenosi."
        P.ctx.nalaz(K, ozb, "tblOtkup", ids, opis, predlog)


def _isplaceno_posle_storna(P):
    K = "OTK_ISPLACENO_POSLE_STORNA_NOVCA"
    if not P.spremno(K, [("tblOtkup", OTK_BAZA + ["Isplaceno"]), ("tblNovac", ["OtkupID", "Isplata", "Stornirano"])]):
        return
    dok = defaultdict(list)
    for o in aktivni(P.otk):
        if k(o, "Isplaceno") != "da" or not k(o, "OtkupID"):
            continue
        oid = k(o, "OtkupID")
        Pp = P.placeno.get(oid, Decimal(0))
        V = _vr(o)
        stor = [x for x in P.po_otk.get(oid, ()) if x.st]
        if Pp < V - Decimal("0.01") and stor:
            dok[_dok_kljuc(o)].append((o, Pp, V, stor))
    for _kl, lst in dok.items():
        ids = _lista(s(o, "OtkupID") for o, *_r in lst)
        Pp = sum((x[1] for x in lst), Decimal(0))
        V = sum((x[2] for x in lst), Decimal(0))
        stor = [y for x in lst for y in x[3]]
        o0 = lst[0][0]
        P.ctx.nalaz(K, GRESKA, "tblOtkup", ids,
                    "Otkup %s (%s, kooperant %s): Isplaceno = Da, a aktivne isplate %s RSD < vrednost %s RSD; "
                    "stornirana isplata: %s." % (ids, s(o0, "BrojDokumenta") or "-", s(o0, "KooperantID"), _rsd(Pp),
                                                 _rsd(V), _lista(y.kratko() for y in stor)),
                    "Isplaceno: Da -> '' i DatumIsplate -> '' (kao UpdateOtkupStatus); otkup je otvoren za ostatak "
                    "%s RSD." % _rsd(V - Pp))


# =====================================================================================
# 4) NOV_VEZA_NA_TUDJI_ILI_STORNIRAN_DOKUMENT
# =====================================================================================
def _veza(P):
    K = "NOV_VEZA_NA_TUDJI_ILI_STORNIRAN_DOKUMENT"
    if not P.redovi("tblNovac"):
        return
    if P.otk and P.spremno(K, [("tblNovac", NOV_BAZA + ["OMID", "BrojDokumenta"]),
                               ("tblOtkup", ["OtkupID", "KooperantID", "StanicaID", "BrojDokumenta", "Stornirano"])]):
        _pp, po_koop_broj, _pk = P.otk_idx()
        tudji_om = defaultdict(list)
        for x in P.akt:
            if not x.otk:
                continue
            o = P.otk_po_id.get(x.otk)
            if o is None:
                P.ctx.nalaz(K, GRESKA, "tblNovac", x.id,
                            "%s: OtkupID %s ne postoji u tblOtkup." % (x.kratko(), x.otk_txt),
                            "Veza se brise; operater bira otkup kooperanta %s (isti broj %s ako postoji)." % (
                                x.koop_txt or "-", x.broj_txt or "-"))
                continue
            if storno(o):
                continue        # -> NOV_ISPLATA_OSIROCENA_STORNOM
            ok = k(o, "KooperantID")
            if x.koop and ok and ok != x.koop:
                kand = [c for c in po_koop_broj.get((x.koop, k(o, "BrojDokumenta")), ()) if not storno(c)]
                P.ctx.nalaz(K, GRESKA, "tblNovac", x.id,
                            "%s (kooperant %s) placa otkup %s TUDJEG kooperanta %s." % (
                                x.kratko(), x.koop_txt, s(o, "OtkupID"), s(o, "KooperantID")),
                            "OtkupID: %s -> %s." % (x.otk_txt, _lista(s(c, "OtkupID") for c in kand))
                            if kand else "Veza se brise (OtkupID -> ''); trazi otkup kooperanta %s, odluka "
                                         "operatera." % x.koop_txt)
            elif not x.koop:
                P.ctx.nalaz(K, SUMNJA, "tblNovac", x.id,
                            "%s nema KooperantID, a vezan je za otkup %s (kooperant %s)." % (
                                x.kratko(), s(o, "OtkupID"), s(o, "KooperantID")),
                            "KooperantID: '' -> %s ili veza se brise; odluka operatera." % s(o, "KooperantID"))
            elif x.izvor not in (IZVOD, SPLIT) and x.omid and k(o, "StanicaID") and k(o, "StanicaID") != x.omid:
                # split nasledjuje OMID originala (izvod: stanica iz sifarnika), a ApplyAvansToOtkup ne gleda
                # stanicu -> razlika je legitimna i ne prijavljuje se
                tudji_om[x.koop].append((x, o))
        for koop, lst in tudji_om.items():
            P.ctx.nalaz(K, INFO, "tblNovac", _lista(x.id for x, _o in lst),
                        "TUDJI_OM: isplata kooperantu %s sa jednog OM placa otkup sa druge stanice: %s." % (
                            lst[0][0].koop_txt, _lista("%s %s -> %s %s" % (x.id, x.omid_txt, s(o, "OtkupID"),
                                                                       s(o, "StanicaID")) for x, o in lst)),
                        "Informativno (kooperant predaje na vise stanica); nista ako je tako dogovoreno.")
    if P.fak_rows and P.spremno(K, [("tblNovac", ["NovacID", "FakturaID", "PartnerID", "EntitetTip", "Stornirano"]),
                                    ("tblFakture", ["FakturaID", "KupacID"])]):
        for x in P.akt:
            if not x.fak or x.etip != "kupac" or not x.pid:
                continue
            f = P.fak_po_id.get(x.fak)
            if f is None or storno(f) or k(f, "Status") == "stornirano":
                continue        # -> FAK_UPLATA_NA_STORNIRANU_FAKTURU
            if k(f, "KupacID") and k(f, "KupacID") != x.pid:
                P.ctx.nalaz(K, GRESKA, "tblNovac", x.id,
                            "%s (kupac %s) vezana za fakturu %s %s TUDJEG kupca %s." % (
                                x.kratko(), x.pid_txt, s(f, "FakturaID"), s(f, "BrojFakture"), s(f, "KupacID")),
                            "FakturaID: %s -> '' (uplata postaje KupciAvans kupca %s) ili PartnerID -> %s; odluka "
                            "operatera." % (x.fak_txt, x.pid_txt, s(f, "KupacID")))


# =====================================================================================
# 5) Fakture: FAK_UPLATA_NA_STORNIRANU_FAKTURU, NOV_UPLATA_KUPCA_..., NOV_UPLATA_KUPCA_ODVEZANA_...
# =====================================================================================
def _fak_storn(f):
    return storno(f) or k(f, "Status") == "stornirano"


def _zamenska_faktura(P, f):
    """Aktivna faktura istog kupca cije stavke nose iste PrijemnicaID kao stornirana f."""
    st = P.redovi("tblFakturaStavke")
    if not st:
        return []
    po_fak = defaultdict(set)
    for r in st:
        if not k(r, "PrijemnicaID"):
            continue
        # stavke stornirane fakture su i same stornirane (StornoFakturaStavkeAndReleasePrijemnice)
        if k(r, "FakturaID") == k(f, "FakturaID") or not storno(r):
            po_fak[k(r, "FakturaID")].add(k(r, "PrijemnicaID"))
    prj = po_fak.get(k(f, "FakturaID"), set())
    if not prj:
        return []
    out = []
    for g in P.fak_rows:
        if g is f or _fak_storn(g) or k(g, "KupacID") != k(f, "KupacID"):
            continue
        if po_fak.get(k(g, "FakturaID"), set()) & prj:
            out.append(g)
    return out


def _fak_txt(f):
    return "%s %s (%s, kupac %s, %s RSD%s)" % (s(f, "FakturaID"), s(f, "BrojFakture"), dan_txt(d(f, "Datum")),
                                              s(f, "KupacID"), _rsd(n(f, "Iznos")),
                                              ", stornirana" if _fak_storn(f) else "")


def _fak_uplata_storn(P):
    K = "FAK_UPLATA_NA_STORNIRANU_FAKTURU"
    if not P.fak_rows or not P.spremno(K, [("tblNovac", ["NovacID", "FakturaID", "Uplata", "Stornirano", "Napomena"]),
                                           ("tblFakture", ["FakturaID", "KupacID", "Stornirano"])]):
        return
    for x in P.akt:
        if not x.fak or x.upl <= 0:
            continue
        f = P.fak_po_id.get(x.fak)
        if f is not None and not _fak_storn(f):
            continue
        if f is None:
            opis = "%s vezana za fakturu %s koja ne postoji." % (x.kratko(), x.fak_txt)
            predlog = "FakturaID: %s -> '' (uplata postaje avans kupca %s)." % (x.fak_txt, x.pid_txt)
        else:
            zam = _zamenska_faktura(P, f)
            opis = "%s vezana za STORNIRANU fakturu %s." % (x.kratko(), _fak_txt(f))
            if x.izvor == IZVOD:
                opis += " Izvor je izvod: storno brise FakturaID, pa je ovo rucna izmena."
            predlog = ("FakturaID: %s -> %s (zamenska faktura istog kupca sa istim prijemnicama)." % (
                x.fak_txt, s(zam[0], "FakturaID")) if len(zam) == 1 else
                "FakturaID: %s -> '' i Tip -> KupciAvans (uplata postaje avans kupca %s)%s." % (
                    x.fak_txt, x.pid_txt, "; kandidati: " + _lista(s(g, "FakturaID") for g in zam) if zam else ""))
        P.ctx.nalaz(K, GRESKA, "tblNovac", x.id, opis, predlog)


def _broj_iz_napomene(nap):
    t = tekst(nap)
    p = t.casefold().find("uplata po fakturi:")
    if p < 0:
        return ""
    rest = t[p + len("uplata po fakturi:"):]
    q = rest.find(" | ")
    return tekst(rest[:q] if q >= 0 else rest)


def _uplata_kupca(P):
    K = "NOV_UPLATA_KUPCA_NA_STORNIRANU_ILI_PREPLACENU_FAKTURU"
    if not P.spremno(K, [("tblNovac", ["NovacID", "Tip", "FakturaID", "PartnerID", "Uplata", "Stornirano",
                                       "Napomena", "Datum"])]):
        return
    # (b) preplacena faktura
    if P.fak_rows and P.spremno(K, [("tblFakture", ["FakturaID", "KupacID", "Iznos", "Stornirano"])]):
        for f in P.fak_rows:
            if _fak_storn(f) or not k(f, "FakturaID"):
                continue
            fid = k(f, "FakturaID")
            U = P.uplaceno.get(fid, Decimal(0))
            Iz = n(f, "Iznos")
            if U - Iz <= TOL:
                continue
            red = sorted((x for x in P.po_fak.get(fid, ()) if not x.st and x.upl > 0), key=lambda x: x.idn)
            preko = [x for x in red if x.izvor == IZVOD]
            visak = U - Iz
            druga = [g for g in P.fak_rows if g is not f and not _fak_storn(g) and k(g, "KupacID") == k(f, "KupacID")
                     and n(g, "Iznos") - P.uplaceno.get(k(g, "FakturaID"), Decimal(0)) > TOL
                     and any(abs(x.upl - n(g, "Iznos")) <= TOL for x in red)]
            P.ctx.nalaz(K, GRESKA, "tblFakture", s(f, "FakturaID"),
                        "Faktura %s preplacena: uplate %s RSD > iznos %s RSD (visak %s RSD); uplate: %s%s." % (
                            _fak_txt(f), _rsd(U), _rsd(Iz), _rsd(visak), _lista(x.kratko() for x in red),
                            "; preko je presao izvod: " + _lista(x.id for x in preko) if preko else ""),
                        "Visak %s RSD odvojiti kao KupciAvans%s." % (
                            _rsd(visak), "; uplata pogodjena po iznosu verovatno pripada nezatvorenoj fakturi " +
                            _lista(_fak_txt(g) for g in druga) if druga else ""))
    # (c) KupciUplata bez fakture
    po_broju = defaultdict(list)
    for f in P.fak_rows:
        if k(f, "BrojFakture"):
            po_broju[(k(f, "KupacID"), poziv_kljuc(polje(f, "BrojFakture")))].append(f)
    for x in P.akt:
        if x.tip != T_KUP_UPL or x.fak or x.upl <= 0:
            continue
        broj = _broj_iz_napomene(x.nap)
        if not broj and x.izvor == IZVOD:
            b = P.bim_po_id.get(x.bim)
            if b is not None:
                broj = s(b, "PozivNaBroj")
        kand = po_broju.get((x.pid, poziv_kljuc(broj)), []) if broj else []
        predlog = "Uplata ce posle migracije potrositi PRVU sledecu fakturu kupca; odluka operatera."
        if kand:
            f = kand[0]
            if _fak_storn(f):
                zam = _zamenska_faktura(P, f)
                predlog = ("FakturaID: '' -> %s (zamenska faktura sa istim prijemnicama)." % s(zam[0], "FakturaID")
                           if len(zam) == 1 else "Tip: KupciUplata -> KupciAvans (zamenska faktura nije nadjena%s)." % (
                               ": kandidati " + _lista(s(g, "FakturaID") for g in zam) if zam else ""))
            else:
                predlog = "FakturaID: '' -> %s." % s(f, "FakturaID")
        P.ctx.nalaz(K, SUMNJA, "tblNovac", x.id,
                    "KupciUplata %s (kupac %s) bez FakturaID%s: odvezana stornom fakture." % (
                        x.kratko(), x.pid_txt, (", originalna faktura '%s'%s" % (
                            broj, " = " + _fak_txt(kand[0]) if kand else "")) if broj else ""),
                    predlog)


def _odvezana_storno_fakture(P):
    K = "NOV_UPLATA_KUPCA_ODVEZANA_STORNOM_FAKTURE"
    if not P.fak_rows or not P.spremno(K, [("tblNovac", ["NovacID", "Tip", "FakturaID", "PartnerID", "Uplata",
                                                         "Stornirano", "Datum"]),
                                           ("tblFakture", ["FakturaID", "KupacID", "Datum", "Stornirano"])]):
        return
    storn = defaultdict(list)
    for f in P.fak_rows:
        if _fak_storn(f) and k(f, "KupacID"):
            storn[k(f, "KupacID")].append(f)
    po_kupcu = defaultdict(list)
    for x in P.akt:
        if x.upl <= 0 or x.fak or x.tip == T_KUP_UPL or not x.pid or x.pid not in storn:
            continue
        fs = [f for f in storn[x.pid] if d(f, "Datum") and x.dan and d(f, "Datum") <= x.dan] or \
             ([] if x.izvor != SPLIT else storn[x.pid])
        if fs:
            po_kupcu[x.pid].append((x, fs))
    for kup, lst in po_kupcu.items():
        spl = [x for x, _f in lst if x.izvor == SPLIT]
        fs = {s(f, "FakturaID"): f for _x, l in lst for f in l}
        zam = [g for f in fs.values() for g in _zamenska_faktura(P, f)]
        P.ctx.nalaz(K, INFO, "tblNovac", _lista(x.id for x, _f in lst),
                    "Kupac %s ima stornirane fakture %s, a avanse bez fakture: %s%s." % (
                        lst[0][0].pid_txt, _lista(fs), _lista(x.kratko() for x, _f in lst),
                        "; sigurno odvezani split redovi: " + _lista(x.id for x in spl) if spl else ""),
                    ("Vezati za zamensku fakturu %s (kao ApplyAvansToFaktura)." % _lista(_fak_txt(g) for g in zam))
                    if zam else "Pregled: avans ostaje avans kupca osim ako operater ne nadje zamensku fakturu.")


# =====================================================================================
# 6) NOV_IZVOD_RUCNO_KUPAC_OM_POGRESNO
# =====================================================================================
def _sadrzi_token(hay, needle):
    h, nd = " %s " % loose(hay), loose(needle)
    return bool(nd) and (" %s " % nd) in h


def _izvod_kupac_om(P):
    K = "NOV_IZVOD_RUCNO_KUPAC_OM_POGRESNO"
    if not P.bim or not P.spremno(K, [("tblNovac", NOV_BAZA + ["PartnerID", "Uplata", "FakturaID"]),
                                      ("tblBankaImport", ["BankaImportID"])]):
        return
    # (a) poziv = faktura istog kupca, a knjizeno kao avans
    if P.fak_rows and P.ctx.ima_kolonu("tblFakture", "BrojFakture") and P.ctx.ima_kolonu("tblFakture", "KupacID"):
        po_kupcu = defaultdict(list)
        for f in P.fak_rows:
            if not _fak_storn(f):
                po_kupcu[k(f, "KupacID")].append(f)
        for x in P.akt:
            if x.izvor != IZVOD or x.etip != "kupac" or x.tip != T_KUP_AV or not x.pid:
                continue
            b = P.bim_po_id.get(x.bim)
            if b is None:
                continue
            pk = poziv_kljuc(polje(b, "PozivNaBroj"))
            svrha = polje(b, "SvrhaPlacanja")
            hits = [f for f in po_kupcu.get(x.pid, ()) if (pk and poziv_kljuc(polje(f, "BrojFakture")) == pk)
                    or (tekst(svrha) and _sadrzi_token(svrha, polje(f, "BrojFakture")))]
            if len(hits) != 1:
                continue
            F = hits[0]
            spl = P.splitovi(x) or []
            faks = {y.fak for y in [x] + spl if not y.st and y.fak}
            if faks == {k(F, "FakturaID")}:
                continue
            opis = "Izvodna uplata %s kupca %s knjizena kao KupciAvans, a poziv/svrha '%s' pokazuje na fakturu %s" % (
                x.kratko(), x.pid_txt, s(b, "PozivNaBroj") or tekst(svrha), _fak_txt(F))
            if not faks:
                opis += "; avans ceka (faktura nije zatvorena tom uplatom)."
            else:
                opis += "; avans potrosen na DRUGU fakturu %s." % _lista(faks)
            P.ctx.nalaz(K, SUMNJA, "tblNovac", x.id, opis,
                        "FakturaID -> %s i Tip -> KupciUplata do otvorenog iznosa; visak ostaje KupciAvans%s." % (
                            s(F, "FakturaID"), "; pogresno zatvorenu fakturu %s ponovo otvoriti" % _lista(faks)
                            if faks else ""))
    # (b) istoimeni partneri
    m = P.mat()
    grupe_k = defaultdict(set)
    for kid, r in m["kup"].items():
        nz = kljuc(polje(r, "Naziv"))
        if nz:
            grupe_k[nz].add(kid)
    grupe_s = defaultdict(set)
    for sid, r in m["st"].items():
        nz = kljuc(polje(r, "Naziv"))
        if nz:
            grupe_s[nz].add(sid)
    clan_k = {kid: g for g in grupe_k.values() if len(g) > 1 for kid in g}
    clan_s = {sid: g for g in grupe_s.values() if len(g) > 1 for sid in g}
    slabi = defaultdict(list)
    for x in P.akt:
        if x.izvor != IZVOD:
            continue
        if x.etip == "kupac" and x.pid in clan_k:
            g, vrsta = clan_k[x.pid], "kupac"
        elif x.etip == "om" and not x.koop and x.pid in clan_s:
            g, vrsta = clan_s[x.pid], "OM"
        else:
            continue
        b = P.bim_po_id.get(x.bim)
        kt = konto_kljuc(polje(b, "PartnerKonto")) if b is not None else konto_kljuc(_seg(x.nap, "Konto"))
        drugi = [c for c in g if c != x.pid and vrsta == "kupac" and kt
                 and konto_kljuc(polje(m["kup"][c], "TekuciRacun")) == kt]
        if len(drugi) == 1:
            P.ctx.nalaz(K, GRESKA, "tblNovac", x.id,
                        "Izvodni red %s mapiran na %s %s, a konto %s pripada istoimenom %s." % (
                            x.kratko(), vrsta, x.pid_txt, kt, s(m["kup"][drugi[0]], "KupacID")),
                        "PartnerID: %s -> %s." % (x.pid_txt, s(m["kup"][drugi[0]], "KupacID")))
            continue
        if vrsta == "kupac" and kt and konto_kljuc(polje(m["kup"][x.pid], "TekuciRacun")) == kt:
            continue        # konto potvrdjuje izabranog
        slabi[tuple(sorted(g))].append((x, vrsta))
    for g, lst in slabi.items():
        P.ctx.nalaz(K, SUMNJA, "tblNovac", _lista(x.id for x, _v in lst),
                    "Istoimeni partneri %s (%s, isti Naziv): rucno mapiranje bira prvi pogodak po nazivu; izvodni "
                    "redovi %s." % (", ".join(g).upper(), lst[0][1], _lista(x.id for x, _v in lst)),
                    "Operater potvrdjuje PartnerID po kontu/PIB-u za svaki red; bez konta odluka ostaje operateru.")
    # (c) smer: isplata kupcu
    for x in P.akt:
        if x.izvor == IZVOD and x.etip == "kupac" and x.isp > 0:
            P.ctx.nalaz(K, SUMNJA, "tblNovac", x.id,
                        "Izvodni red %s je ISPLATA kupcu %s knjizena kao %s." % (x.kratko(), x.pid_txt, x.tip_txt),
                        "Proveriti smer i partnera u tblBankaImport (%s): povracaj kupcu ili pogresno mapiran "
                        "partner." % x.bim.upper())


# =====================================================================================
# 7) NOV_IZVOD_POGRESAN_KOOPERANT
# =====================================================================================
def _pogresan_kooperant(P):
    K = "NOV_IZVOD_POGRESAN_KOOPERANT"
    if not P.bim or not P.redovi("tblKooperanti") or not P.spremno(
            K, [("tblNovac", NOV_BAZA), ("tblBankaImport", ["BankaImportID", "Partner"]),
                ("tblKooperanti", ["KooperantID", "Ime", "Prezime"])]):
        return
    m = P.mat()
    ima_konto = P.ctx.ima_kolonu("tblKooperanti", "TekuciRacun")
    for b in P.bim:
        bid = k(b, "BankaImportID")
        direkt, spl, _dvo = P.lanac(bid)
        direkt = [x for x in direkt if not x.st]
        koops = {x.koop for x in direkt if x.koop}
        if len(koops) != 1:
            continue
        K_ = next(iter(koops))
        kt_b = konto_kljuc(polje(b, "PartnerKonto")) or konto_kljuc(_seg(direkt[0].nap, "Konto"))
        r_k = m["koop"].get(K_)
        kt_k = konto_kljuc(polje(r_k, "TekuciRacun")) if (r_k is not None and ima_konto) else ""
        k2 = None
        if kt_b and kt_k and kt_b != kt_k:
            drugi = m["koop_konto"].get(kt_b, set()) - {K_}
            if len(drugi) == 1:
                k2 = next(iter(drugi))
        poz = P.otk_po_pozivu(polje(b, "PozivNaBroj"))
        poz_k = {k(o, "KooperantID") for o in poz}
        k3 = next(iter(poz_k)) if len(poz_k) == 1 and K_ not in poz_k else None
        poziv_svoj = K_ in poz_k
        slab = None
        pm = m["pm"].get(k(b, "Partner"))
        if pm is not None and k(pm, "PartnerID") == K_:
            imenjaci = m["koop_ime"].get(loose(polje(b, "Partner")), set())
            if len(imenjaci) >= 2 and K_ in imenjaci and not poziv_svoj and not (kt_b and kt_k == kt_b):
                slab = sorted(imenjaci - {K_})
        if not (k2 or k3 or slab):
            continue
        cilj = k2 or k3
        if k2 and (k3 is None or k3 == k2) and not poziv_svoj:
            ozb = GRESKA
        else:
            ozb = SUMNJA
        dokazi = []
        if k2:
            dokazi.append("konto %s pripada kooperantu %s (sifarnik %s: %s)" % (kt_b, k2.upper(), K_.upper(), kt_k))
        if k3:
            dokazi.append("poziv %s pogadja otkup %s kooperanta %s" % (
                s(b, "PozivNaBroj"), _lista(s(o, "OtkupID") for o in poz), k3.upper()))
        if slab:
            dokazi.append("PartnerMap '%s' -> %s, a isto ime ima i %s (slab pogodak)" % (
                s(b, "Partner"), K_.upper(), ", ".join(x.upper() for x in slab)))
        if poziv_svoj:
            dokazi.append("poziv ipak pogadja blok kooperanta %s (moguca isplata na racun clana porodice)" % K_.upper())
        potroseno = [y.otk_txt for y in direkt + spl if not y.st and y.otk]
        st2 = s(m["koop"][cilj], "StanicaID") if cilj and cilj in m["koop"] else ""
        predlog = ("KooperantID: %s -> %s, OMID -> %s%s." % (
            direkt[0].koop_txt, cilj.upper(), st2 or "?",
            "; veze na otkupe %s se brisu (tudji otkupi), potroseni avans se vraca" % _lista(potroseno)
            if potroseno else "") if cilj else
            "Operater potvrdjuje kooperanta po kontu; proveriti zapis tblPartnerMap '%s'." % s(b, "Partner"))
        if pm is not None and cilj and k(pm, "PartnerID") == K_:
            predlog += " Pogresan zapis tblPartnerMap '%s' -> %s prijaviti posebno." % (s(b, "Partner"), K_.upper())
        P.ctx.nalaz(K, ozb, "tblNovac", _lista(x.id for x in direkt),
                    "Izvod %s pripisan kooperantu %s: %s." % (_bim_txt(b), direkt[0].koop_txt, "; ".join(dokazi)),
                    predlog)


# =====================================================================================
# 8) NOV_TIP_ISPLATE_PO_PODRAZUMEVANOJ
# =====================================================================================
def _tip_isplate(P):
    K = "NOV_TIP_ISPLATE_PO_PODRAZUMEVANOJ"
    if not P.spremno(K, [("tblNovac", NOV_BAZA + ["OMID"])]):
        return
    # (1) rucni VirmanFirmaKoop bez izvoda, a period je pokriven izvodima
    if P.bim and P.ctx.ima_kolonu("tblBankaImport", "DatumTransakcije"):
        racuni = defaultdict(list)
        for b in P.bim:
            dd = d(b, "DatumTransakcije")
            if dd and not storno(b):
                racuni[k(b, "BrojRacuna") if P.ctx.ima_kolonu("tblBankaImport", "BrojRacuna") else ""].append(dd)
        izv = defaultdict(list)     # koop -> [(dan, iznos)]
        for b in P.bim:
            direkt = [x for x in P.po_bim.get(k(b, "BankaImportID"), ()) if not x.st and x.izvor == IZVOD]
            kk = {x.koop for x in direkt if x.koop}
            if len(kk) == 1:
                izv[next(iter(kk))].append((P.bim_dan(b, direkt), n(b, "Isplata")))
        rucni_vir = [x for x in P.akt if x.izvor == OM_ULAZ and x.tip == T_VIR_FK and x.otk and x.dan and x.isp > 0]
        pogoci = []
        for x in rucni_vir:
            if not izv.get(x.koop):
                continue        # kooperant se ne placa sa uvezenog racuna: virman sa drugog racuna je moguc
            if any(_u_prozoru(dd, x.dan - timedelta(days=3), x.dan + timedelta(days=15)) and abs(a - x.isp) <= TOL
                   for dd, a in izv.get(x.koop, ())):
                continue        # -> NOV_DUPLA_ISPLATA_RUCNO_I_IZVOD
            pokriven = bool(racuni) and all(
                any(x.dan - timedelta(days=30) <= dd < x.dan for dd in dani) and
                any(x.dan < dd <= x.dan + timedelta(days=30) for dd in dani) for dani in racuni.values())
            if pokriven:
                pogoci.append(x)
        if len(pogoci) > max(5, len(rucni_vir) * 3 // 10):
            # sistemska praksa (rucno knjizeni virmani sa drugog racuna), ne zaboravljen prekidac
            P.ctx.nalaz(K, INFO, "tblNovac", _lista(x.id for x in pogoci),
                        "%d od %d rucnih VirmanFirmaKoop nema izvodnu isplatu iako su izvodi za period uvezeni: "
                        "verovatno virmani sa racuna ciji izvod nije uvezen; pojedinacno se ne prijavljuju." % (
                            len(pogoci), len(rucni_vir)),
                        "Proveriti sa kog racuna se placaju kooperanti; ako je kes, Tip -> KesOtkupacKoop.")
        else:
            for x in pogoci:
                P.ctx.nalaz(K, SUMNJA, "tblNovac", x.id,
                            "Rucni VirmanFirmaKoop %s (kooperant %s) nema izvodnu isplatu istog iznosa u [-3, +15] "
                            "dana, iako su izvodi za period uvezeni i kooperant se placa sa tog racuna: moguc kes "
                            "zapisan kao virman (prekidac 'iz OM avansa' vracen na podrazumevano)." % (
                                x.kratko(), x.koop_txt),
                            "Tip: VirmanFirmaKoop -> KesOtkupacKoop samo ako saldo OM %s ne ide u minus i ako virman "
                            "nije isplacen sa drugog racuna; odluka operatera." % x.omid_txt)
    # (2) saldo OM po hronologiji NovacID (informativno). GetOMAvansSaldo je provera samo na Ulazu OM uz
    #     prekidac 'iz OM avansa' (frmDokumenta.frm:1747-1756), pa je saldo < 0 POSLE takvog reda trag da
    #     se saldo kasnije promenio (storno KesFirmaOtkupac, rucna izmena) -- kes uz otkup ga legitimno obara.
    po_om = defaultdict(list)
    for x in P.akt:
        if x.omid and x.tip in (T_KES_FO, T_KES_OK):
            po_om[x.omid].append(x)
    for om, lst in po_om.items():
        if not any(x.tip == T_KES_OK and x.izvor == OM_ULAZ for x in lst):
            continue
        saldo, minimum, prvi, upl = Decimal(0), Decimal(0), None, Decimal(0)
        for x in sorted(lst, key=lambda y: y.idn):
            if x.tip == T_KES_FO:
                saldo += x.isp
                upl += x.upl
            else:
                saldo -= x.isp
            if saldo < -TOL and prvi is None and x.tip == T_KES_OK and x.izvor == OM_ULAZ:
                prvi = (x, saldo)
            minimum = min(minimum, saldo)
        if prvi is None:
            continue
        stor = [y for y in P.nov if y.st and y.omid == om and y.tip == T_KES_FO and y.idn < prvi[0].idn]
        P.ctx.nalaz(K, INFO, "tblNovac", prvi[0].id,
                    "Saldo OM %s (KesFirmaOtkupac - KesOtkupacKoop po NovacID) je %s RSD posle isplate 'iz OM "
                    "avansa' %s, koju bi stara provera odbila; minimum %s RSD%s%s. Saldo je informativan." % (
                        lst[0].omid_txt, _rsd(prvi[1]), prvi[0].kratko(), _rsd(minimum),
                        "; kasnije stornirani KesFirmaOtkupac: " + _lista(y.id for y in stor) if stor else "",
                        "; uplate OM-a (ne ulaze u saldo): %s RSD" % _rsd(upl) if upl else ""),
                    "Proveriti avanse OM %s pre %s (storno ili rucna izmena posle isplate)." % (
                        lst[0].omid_txt, dan_txt(prvi[0].dan)))
    # (3) rucni KesFirmaOtkupac bez primaoca = ostatak otvorenog otkupa istog OM istog dana
    if P.otk and P.ctx.ima_kolonu("tblOtkup", "StanicaID"):
        po_st_dan = defaultdict(list)
        for o in aktivni(P.otk):
            if k(o, "StanicaID") and d(o, "Datum"):
                po_st_dan[(k(o, "StanicaID"), d(o, "Datum"))].append(o)
        for x in P.akt:
            if x.izvor != OM_ULAZ or x.tip != T_KES_FO or x.koop or x.isp <= 0 or not x.dan:
                continue
            kand, vid_dok = [], set()
            for o in po_st_dan.get((x.omid, x.dan), ()):
                dk = _dok_kljuc(o)
                if dk in vid_dok:
                    continue
                vid_dok.add(dk)
                otv = P.otvoreno_dok(o)     # po dokumentu (kes Klase I/II je na primarnom redu)
                if otv > TOL and abs(otv - x.isp) <= TOL:
                    kand.append(max(P.dok(o)[2], key=lambda rr: P.otvoreno(rr)))
            if len(kand) != 1:
                continue
            o = kand[0]
            P.ctx.nalaz(K, SUMNJA, "tblNovac", x.id,
                        "Rucni KesFirmaOtkupac %s bez primaoca jednak je ostatku otvorenog otkupa %s istog OM i dana "
                        "(slab pogodak: isplata kooperantu uneta bez primaoca?)." % (x.kratko(), _otk_txt(o)),
                        "Tip -> KesOtkupacKoop, KooperantID -> %s, OtkupID -> %s, uz potvrdu operatera." % (
                            s(o, "KooperantID"), s(o, "OtkupID")))


# =====================================================================================
# 9) NOV_ZASTAREO_DATUM_FORME
# =====================================================================================
def _zastareo_datum(P):
    K = "NOV_ZASTAREO_DATUM_FORME"
    if not P.spremno(K, [("tblNovac", NOV_BAZA)]):
        return
    om = sorted((x for x in P.nov if x.izvor == OM_ULAZ), key=lambda y: y.idn)
    sused = {}
    for i, x in enumerate(om):
        sused[x.idk] = [y.dan for y in om[max(0, i - 1):i + 2] if y is not x and y.dan]
    if P.otk and P.ctx.ima_kolonu("tblOtkup", "Datum"):
        for x in P.akt:
            if x.izvor != OM_ULAZ or x.tip not in (T_KES_OK, T_VIR_FK) or not x.otk or not x.dan:
                continue
            o = P.otk_po_id.get(x.otk)
            if o is None or not d(o, "Datum") or x.dan >= d(o, "Datum"):
                continue
            kand = max([d(o, "Datum")] + sused.get(x.idk, []))
            P.ctx.nalaz(K, SUMNJA, "tblNovac", x.id,
                        "Rucna isplata %s datirana %s, PRE otkupa koji placa %s (Ulaz OM nudi samo postojece "
                        "otkupe): zastareo datum forme ili pogresan datum otkupa." % (
                            x.kratko(), dan_txt(x.dan), _otk_txt(o)),
                        "Datum: %s -> %s (max datuma otkupa i susednog rucnog reda), ili ispraviti datum otkupa; "
                        "odluka operatera." % (dan_txt(x.dan), dan_txt(kand)))
    if P.fak_rows and P.ctx.ima_kolonu("tblFakture", "Datum"):
        for x in P.akt:
            if x.izvor != KUPAC_RUCNO or x.tip != T_KUP_UPL or not x.fak or not x.dan:
                continue
            f = P.fak_po_id.get(x.fak)
            if f is None or not d(f, "Datum") or x.dan >= d(f, "Datum"):
                continue
            P.ctx.nalaz(K, SUMNJA, "tblNovac", x.id,
                        "Rucna uplata %s datirana %s, PRE fakture %s." % (x.kratko(), dan_txt(x.dan), _fak_txt(f)),
                        "Datum: %s -> %s ili proveriti datum fakture; odluka operatera." % (
                            dan_txt(x.dan), dan_txt(d(f, "Datum"))))


# =====================================================================================
# 10) NOV_IZVOD_DATUM_LOKAL
# =====================================================================================
def _datum_lokal(P):
    K = "NOV_IZVOD_DATUM_LOKAL"
    if not P.bim or not P.spremno(K, [("tblNovac", ["NovacID", "Datum", "Napomena", "Stornirano"]),
                                      ("tblBankaImport", ["BankaImportID", "DatumTransakcije"])]):
        return
    ima_di = P.ctx.ima_kolonu("tblBankaImport", "DatumIzvoda")
    pogoci = []
    for b in P.bim:
        direkt = [x for x in P.po_bim.get(k(b, "BankaImportID"), ()) if not x.st and x.izvor == IZVOD and x.dan]
        if not direkt:
            continue
        raw = polje(b, "DatumTransakcije")
        tekstualan = isinstance(raw, str)
        D0 = None
        if tekstualan:
            mm = _DMY.match(raw)
            if mm:
                try:
                    D0 = date(int(mm.group(3)), int(mm.group(2)), int(mm.group(1)))
                except ValueError:
                    D0 = None
        else:
            D0 = _datum(raw)
        if D0 is None:
            continue
        nd = direkt[0].dan
        lose = [x for x in direkt if x.dan != D0]
        if lose:
            nd = lose[0].dan
            zam = nd.day == D0.month and nd.month == D0.day and nd.year == D0.year
            pogoci.append((b, direkt, D0, nd, "ZAMENA_DAN_MESEC" if zam else "DRUGO", tekstualan))
            continue
        if ima_di and not tekstualan and D0.day <= 12 and D0.day != D0.month:
            DI = d(b, "DatumIzvoda")
            if DI:
                try:
                    sw = date(D0.year, D0.day, D0.month)
                except ValueError:
                    sw = None
                u = DI - timedelta(days=7) <= D0 <= DI
                if not u and sw and DI - timedelta(days=7) <= sw <= DI:
                    pogoci.append((b, direkt, sw, nd, "POTENCIJALNO_ZAMENJEN", tekstualan))
    if not pogoci:
        return
    sistem = len(pogoci) >= 2 and all(p[4] in ("ZAMENA_DAN_MESEC", "POTENCIJALNO_ZAMENJEN") for p in pogoci)
    for b, direkt, D0, nd, klasa, tekstualan in pogoci:
        if klasa == "POTENCIJALNO_ZAMENJEN":
            ozb = SUMNJA
            opis = ("Izvodni red %s: datum transakcije %s (vec datum u stagingu) je van [DatumIzvoda-7, DatumIzvoda] "
                    "= %s, a zamenjen dan/mesec %s upada u prozor." % (
                        _lista(x.id for x in direkt), dan_txt(nd), dan_txt(d(b, "DatumIzvoda")), dan_txt(D0)))
        else:
            ozb = GRESKA if (klasa == "ZAMENA_DAN_MESEC" and tekstualan) else SUMNJA
            opis = "Izvodni red %s ima Datum %s, a stavka %s ima DatumTransakcije '%s' (%s): %s." % (
                _lista(x.id for x in direkt), dan_txt(nd), s(b, "BankaImportID"), tekst(polje(b, "DatumTransakcije")),
                dan_txt(D0), klasa if tekstualan or klasa != "DRUGO" else "DRUGO (moguca rucna izmena)")
        if sistem:
            opis += " Svi pogoci imaju dan <= 12: sistematska zamena po lokalu verovatna."
        P.ctx.nalaz(K, ozb, "tblNovac", _lista(x.id for x in direkt), opis,
                    "Datum: %s -> %s na %s (i split redovima istog lanca)." % (
                        dan_txt(nd), dan_txt(D0), _lista(x.id for x in direkt)))


# =====================================================================================
# 11) NOV_STORNO_NOVCA_BEZ_LANCA, NOV_AVANS_RASPODELA_RAZDVOJENA_STORNOM
# =====================================================================================
def _storno_bez_lanca(P):
    K = "NOV_STORNO_NOVCA_BEZ_LANCA"
    KR = "NOV_AVANS_RASPODELA_RAZDVOJENA_STORNOM"
    if not P.spremno(K, [("tblNovac", NOV_BAZA + ["BrojDokumenta", "PartnerID", "Uplata"])]):
        return
    # (a) storniran original, aktivni split redovi
    for kk, kor in P.av_kor.items():
        spl_akt = [y for y in P.av_spl.get(kk, ()) if not y.st]
        if not kor or not spl_akt or not all(x.st for x in kor):
            continue
        izvod = any(x.izvor == IZVOD for x in kor)
        potr = [y.otk_txt or y.fak_txt for y in spl_akt if y.otk or y.fak]
        P.reported[KR].update(x.idk for x in kor)
        P.ctx.nalaz(K, GRESKA if izvod else SUMNJA, "tblNovac", _lista(x.id for x in kor),
                    "Storniran %s %s, a split redovi istog lanca su aktivni: %s%s." % (
                        "izvodni avans" if izvod else "avans", _lista(x.kratko() for x in kor),
                        _lista(y.kratko() for y in spl_akt),
                        " (potroseni na %s)" % _lista(potr) if potr else ""),
                    "Storno i split redova %s, ili ponistiti storno %s; zavisi od toga sta je operater hteo." % (
                        _lista(y.id for y in spl_akt), _lista(x.id for x in kor)))
    # (b) stavka izvoda 'Da' bez aktivnog reda u knjizi, (c) rucna zamena, (d) zbir lanca
    if not P.bim or not P.spremno(K, [("tblBankaImport", ["BankaImportID", "Isplata", "Uplata", "Obradjeno"])]):
        return
    for b in P.bim:
        bid = k(b, "BankaImportID")
        A = n(b, "Isplata") + n(b, "Uplata")
        if not bid or storno(b) or A <= 0:
            continue
        direkt, spl, dvo = P.lanac(bid)
        akt = [x for x in direkt + spl if not x.st]
        if k(b, "Obradjeno") == "da" and not akt and not dvo:
            stor = [x for x in direkt if x.st]
            isplata = n(b, "Isplata") > 0
            D = P.bim_dan(b, direkt)
            pokusaji = []
            if stor:
                x0 = stor[0]
                pokusaji.append((("kooperant", x0.koop) if x0.koop else ("kupac", x0.pid) if x0.etip == "kupac"
                                 else ("om", x0.omid)) + ("storniran red",))
            pb = P.partner_bim(b)
            if pb:
                # izvodni red je cesto storniran BAS zato sto je mapiran na pogresnog partnera, pa je rucna
                # zamena na pravom partneru (konto -> PartnerMap -> ime)
                pokusaji.append(pb)
            zam, kako = [], ""
            for vrsta, pid, kako_p in pokusaji:
                zam = P.rucna_zamena(vrsta, pid, isplata, A, D) if (vrsta and pid) else []
                if zam:
                    kako = kako_p
                    break
            opis = "Stavka izvoda %s je Obradjeno = Da, a u knjizi nema aktivnog reda (%s)." % (
                _bim_txt(b), "stornirani: " + _lista(x.id for x in stor) if stor else "nijedan red sa BIM:%s" % bid.upper())
            if zam:
                opis += " Rucni red istog partnera i iznosa: %s (zamena rucnim unosom, partner po: %s)." % (
                    _lista(x.kratko() for x in zam), kako)
                predlog = "Povezati rucni red %s napomenom BIM:%s (nije greska ako je zamena namerna)." % (
                    zam[0].id, s(b, "BankaImportID"))
            else:
                predlog = "Obradjeno: Da -> '' da se stavka ponovo mapira u novoj svesci (ili Skip ako nije za knjigu)."
            P.ctx.nalaz(K, SUMNJA if zam else GRESKA, "tblBankaImport", s(b, "BankaImportID"), opis, predlog)
            continue
        if not akt or dvo:
            continue
        if any(x.st for x in direkt + spl):
            continue        # (a) / razdvojena raspodela
        zbir = sum((x.iznos() for x in akt), Decimal(0))
        if abs(zbir - A) > TOL:
            P.ctx.nalaz(K, GRESKA, "tblBankaImport", s(b, "BankaImportID"),
                        "Zbir aktivnog izvodnog lanca %s = %s RSD, a stavka %s glasi na %s RSD (razlika %s)." % (
                            _lista(x.id for x in akt), _rsd(zbir), _bim_txt(b), _rsd(A), _rsd(zbir - A)),
                        "Proveriti rucnu izmenu iznosa na %s; lanac mora dati iznos izvoda." % _lista(x.id for x in akt))


def _razdvojena(P):
    KR = "NOV_AVANS_RASPODELA_RAZDVOJENA_STORNOM"
    if not P.spremno(KR, [("tblNovac", NOV_BAZA + ["BrojDokumenta", "PartnerID", "Uplata"])]):
        return
    for kk, spl in P.av_spl.items():
        kor = P.av_kor.get(kk, [])
        if len(kor) != 1 or kor[0].st:
            continue
        R = kor[0]
        if R.idk in P.reported[KR]:
            continue
        stor = [y for y in spl if y.st]
        if not stor:
            continue
        akt = [y for y in spl if not y.st]
        # ponovna raspodela istog iznosa (noviji aktivan split) je legitimna
        stor = [y for y in stor if not any(z.idn > y.idn and abs(z.iznos() - y.iznos()) <= TOL for z in akt)]
        if not stor:
            continue
        if R.izvor == IZVOD:
            b = P.bim_po_id.get(R.bim)
            if b is None:
                continue
            A = n(b, "Isplata") + n(b, "Uplata")
            zbir = R.iznos() + sum((y.iznos() for y in akt), Decimal(0))
            if A - zbir <= TOL:
                continue
            ozb = GRESKA
            opis = ("Storniran split %s izvodnog avansa %s, original aktivan: lanac daje %s RSD, a izvod %s glasi na "
                    "%s RSD (nedostaje %s)." % (_lista(y.kratko() for y in stor), R.id, _rsd(zbir),
                                                s(b, "BankaImportID"), _rsd(A), _rsd(A - zbir)))
        else:
            ozb = SUMNJA
            opis = ("Storniran split %s, a original avans %s aktivan: iznos stornirane raspodele je nestao iz avansa "
                    "(original je pri raspodeli smanjen)." % (_lista(y.kratko() for y in stor), R.kratko()))
        P.ctx.nalaz(KR, ozb, "tblNovac", _lista(y.id for y in stor), opis,
                    "Originalu %s vratiti iznos stornirane raspodele (%s RSD) ili ponistiti storno; odluka operatera." % (
                        R.id, _rsd(sum((y.iznos() for y in stor), Decimal(0)))))


# =====================================================================================
# 12) NOV_DUPLI_RUCNI_UNOS
# =====================================================================================
def _dupli_rucni(P):
    K = "NOV_DUPLI_RUCNI_UNOS"
    if not P.spremno(K, [("tblNovac", NOV_BAZA + ["BrojDokumenta", "PartnerID", "OMID", "Uplata", "FakturaID"])]):
        return
    grupe = defaultdict(list)
    vec = P.reported.get("NOV_ISPLATA_OSIROCENA_STORNOM", set())     # siroce + naslednik: vec prijavljeno
    for x in P.akt:
        if x.izvor not in RUCNI or not x.dan or x.iznos() <= 0 or x.idk in vec:
            continue
        if x.tip == T_KES_FO:
            ent = x.omid
        else:
            ent = x.koop or x.pid
        if not ent:
            continue
        grupe[(x.tip, ent, round(x.iznos(), 2))].append(x)
    for kl, lst in grupe.items():
        if len(lst) < 2:
            continue
        lst.sort(key=lambda y: (y.dan, y.idn))
        par = defaultdict(set)
        for i, a in enumerate(lst):
            for b in lst[i + 1:]:
                if (b.dan - a.dan).days > 1:
                    break
                jak = None
                if a.otk and b.otk:
                    if a.otk == b.otk:
                        jak = "isti OtkupID %s" % a.otk_txt
                elif a.fak and b.fak:
                    if a.fak == b.fak:
                        jak = "isti FakturaID %s" % a.fak_txt
                if jak is None and not (a.otk and b.otk) and not (a.fak and b.fak and a.fak != b.fak):
                    if a.tip == T_KES_FO:
                        if (a.broj and a.broj == b.broj) or (a.nap and a.nap.casefold() == b.nap.casefold()):
                            jak = "isti broj/napomena"
                    elif a.broj and a.broj == b.broj:
                        jak = "isti BrojDokumenta %s" % a.broj_txt
                    elif bool(a.broj) != bool(b.broj):
                        jak = "jedan red bez BrojDokumenta"
                    elif 1 <= abs(a.idn - b.idn) <= 3:
                        jak = "uzastopni NovacID"
                if jak:
                    par[a.idk].add((b.idk, jak))
                    par[b.idk].add((a.idk, jak))
        if not par:
            continue
        # povezane komponente
        po = {x.idk: x for x in lst}
        vid = set()
        for x in lst:
            if x.idk not in par or x.idk in vid:
                continue
            komp, stek, razlozi = [], [x.idk], set()
            while stek:
                c = stek.pop()
                if c in vid:
                    continue
                vid.add(c)
                komp.append(po[c])
                for (dd, why) in par.get(c, ()):
                    razlozi.add(why)
                    stek.append(dd)
            komp.sort(key=lambda y: y.idn)
            ozb = SUMNJA
            otk = komp[0].otk
            if otk and all(y.otk == otk for y in komp):
                o = P.otk_po_id.get(otk)
                if o is not None:
                    if P.placeno.get(otk, Decimal(0)) - _vr(o) > TOL:
                        ozb = GRESKA
            fak = komp[0].fak
            if fak and all(y.fak == fak for y in komp):
                f = P.fak_po_id.get(fak)
                if f is not None and P.uplaceno.get(fak, Decimal(0)) - n(f, "Iznos") > TOL:
                    ozb = GRESKA
            P.ctx.nalaz(K, ozb, "tblNovac", _lista(y.id for y in komp),
                        "Ista rucna %s uneta %d puta (%s, %s RSD, +-1 dan): %s; %s%s." % (
                            "uplata" if komp[0].upl > 0 else "isplata", len(komp), komp[0].tip_txt,
                            _rsd(komp[0].iznos()), _lista(y.kratko() for y in komp), "; ".join(sorted(razlozi)),
                            "; dokument je time preplacen" if ozb == GRESKA else ""),
                        "Kandidat za storno: %s (veci NovacID); posle storna proveriti NOV_PREPLATA_OTKUPA." % _lista(
                            y.id for y in komp[1:]))


# =====================================================================================
# PROPUSTENO 1) IZVOD_VAN_KNJIGE
# =====================================================================================
def _izvod_van_knjige(P):
    K = "IZVOD_VAN_KNJIGE"
    if not P.bim or not P.spremno(K, [("tblBankaImport", ["BankaImportID", "Isplata", "Uplata", "Obradjeno"]),
                                      ("tblNovac", ["NovacID", "Napomena", "Stornirano"])]):
        return
    agregat = defaultdict(list)
    for b in P.bim:
        bid = k(b, "BankaImportID")
        A = n(b, "Isplata") + n(b, "Uplata")
        status = k(b, "Obradjeno")
        if not bid or storno(b) or A <= 0 or status not in ("skip", "error", ""):
            continue
        direkt, spl, _dvo = P.lanac(bid)
        if any(not x.st for x in direkt + spl):
            continue
        isplata = n(b, "Isplata") > 0
        D = P.bim_dan(b)
        pb = P.partner_bim(b)
        zam = P.rucna_zamena(pb[0], pb[1], isplata, A, D) if pb else []
        pm = P.mat()["pm"].get(k(b, "Partner")) if k(b, "Partner") else None
        st_txt = s(b, "Obradjeno") or "''"
        if zam:
            P.ctx.nalaz(K, SUMNJA, "tblBankaImport", s(b, "BankaImportID"),
                        "Stavka %s (Obradjeno %s) nije u knjizi, a postoji rucni red istog partnera %s i iznosa: %s "
                        "(zamena rucnim unosom)." % (_bim_txt(b), st_txt, pb[1].upper(), _lista(x.kratko() for x in zam)),
                        "Povezati %s napomenom BIM:%s i stavku oznaciti Da; NE mapirati ponovo (duplikat)." % (
                            zam[0].id, s(b, "BankaImportID")))
        elif pm is not None and status in ("skip", "error"):
            P.ctx.nalaz(K, SUMNJA, "tblBankaImport", s(b, "BankaImportID"),
                        "Stavka %s (Obradjeno %s) nije u knjizi, a '%s' postoji u tblPartnerMap -> %s %s: ispravno "
                        "mapiranje na drugog partnera je bilo blokirano (savePartnerMap greska 1039)." % (
                            _bim_txt(b), st_txt, s(b, "Partner"), s(pm, "EntitetTip"), s(pm, "PartnerID")),
                        "Proknjiziti stavku u novoj svesci na pravog partnera; proveriti zapis tblPartnerMap.")
        else:
            agregat[status].append(b)
    for status, lst in agregat.items():
        P.ctx.nalaz(K, INFO, "tblBankaImport", _lista(s(b, "BankaImportID") for b in lst),
                    "%d stavki izvoda sa Obradjeno %s nije u knjizi (ukupno %s RSD), bez traga zamene: %s." % (
                        len(lst), s(lst[0], "Obradjeno") or "''",
                        _rsd(sum((n(b, "Isplata") + n(b, "Uplata") for b in lst), Decimal(0))),
                        _lista(_bim_txt(b) for b in lst[:3])),
                    "Pregled: provizije i interni prenosi ostaju Skip; ostalo proknjiziti u novoj svesci.")


# =====================================================================================
# PROPUSTENO 2) IZVOD_BEZ_POZIVA_ODE_U_AVANS
# =====================================================================================
def _bez_poziva_avans(P):
    K = "IZVOD_BEZ_POZIVA_ODE_U_AVANS"
    if not P.bim or not P.otk or not P.spremno(K, [("tblNovac", NOV_BAZA + ["BrojDokumenta"]),
                                                  ("tblBankaImport", ["BankaImportID"]),
                                                  ("tblOtkup", OTK_BAZA)]):
        return
    dupla = P.reported.get("NOV_DUPLA_ISPLATA_RUCNO_I_IZVOD", set())
    for R in P.akt:
        if R.izvor != IZVOD or R.tip != T_VIR_AV or R.match != M_KOOP or not R.koop or not R.dan:
            continue
        if R.bim in dupla:
            continue
        b = P.bim_po_id.get(R.bim)
        poziv = polje(b, "PozivNaBroj") if b is not None else ""
        if P.otk_po_pozivu(poziv, R.koop, aktivni_samo=False):
            continue        # poziv pogadja blok kooperanta (zatvoren u trenutku mapiranja) -> NOV_DUPLA
        spl = P.splitovi(R)
        if spl is None:
            continue
        lanac = [R] + [y for y in spl if not y.st]
        iznos = sum((y.isp for y in lanac), Decimal(0))
        potr = [P.otk_po_id.get(y.otk) for y in lanac if y.otk]
        potr = [o for o in potr if o is not None]
        potr_ids = {k(o, "OtkupID") for o in potr}
        kasnije = [o for o in potr if d(o, "Datum") and d(o, "Datum") > R.dan]
        potr_dok = {_dok_kljuc(o) for o in potr}
        kand, vid_dok = [], set()
        for o in P.otk_idx()[2].get(R.koop, ()):
            if storno(o) or k(o, "OtkupID") in potr_ids or not d(o, "Datum") or d(o, "Datum") > R.dan:
                continue
            # odluka po DOKUMENTU: red Klase II bez kesa nije otvoren blok kad je kes na primarnom redu
            dk = _dok_kljuc(o)
            if dk in vid_dok or dk in potr_dok:
                continue
            vid_dok.add(dk)
            Vd, _pd, redovi_dok = P.dok(o)
            otv = P.otvoreno_dok(o)
            rucno_kasnije = [y for rr in redovi_dok for y in P.po_otk.get(k(rr, "OtkupID"), ())
                             if not y.st and y.izvor in (OM_ULAZ, OSTALO) and y.dan and y.dan > R.dan]
            iznos_ok = abs(iznos - otv) <= TOL or abs(iznos - Vd) <= TOL
            otvoren = otv > max(TOL, Vd * TOL_PCT)      # ostatak od zaokruzivanja nije otvoren blok
            if (otvoren and (kasnije or iznos_ok)) or (rucno_kasnije and iznos_ok):
                kand.append((not iznos_ok, _dani(d(o, "Datum"), R.dan), o, otv, rucno_kasnije))
        if not kand:
            continue
        kand.sort(key=lambda c: (c[0], c[1]))
        _x, _dd, o, otv, rk = kand[0]
        opis = ("Izvodni avans %s (Match:Kooperant, poziv '%s' ne pogadja blok) knjizen ceo kao avans, a na dan "
                "izvoda postoji stariji blok %s (otvoreno %s RSD%s)." % (
                    R.kratko(), tekst(poziv), _otk_txt(o), _rsd(otv),
                    ", kasnije placen rucno: " + _lista(y.id for y in rk) if rk else ""))
        if kasnije:
            opis += " Avans je potrosen na kasniji otkup %s." % _lista(s(c, "OtkupID") for c in kasnije)
        if len(kand) > 1:
            opis += " Drugi kandidati: %s." % _lista(s(c[2], "OtkupID") for c in kand[1:])
        P.ctx.nalaz(K, SUMNJA, "tblNovac", R.id, opis,
                    "Prevezati avans na stari blok %s (do %s RSD)%s; odluka operatera." % (
                        s(o, "OtkupID"), _rsd(min(iznos, otv if otv > 0 else _vr(o))),
                        "; rucnu isplatu %s proveriti kao duplikat" % _lista(y.id for y in rk) if rk else ""))


# =====================================================================================
# PROPUSTENO 3) REVERS_STORNIRAN_NOVAC_OSTAO
# =====================================================================================
def _revers(P):
    K = "REVERS_STORNIRAN_NOVAC_OSTAO"
    if not P.redovi("tblAmbalaza") or not P.spremno(K, [("tblNovac", NOV_BAZA + ["BrojDokumenta"]),
                                                       ("tblAmbalaza", ["DokumentID", "DokumentTip", "Stornirano"])]):
        return
    amb = P.ctx.amb_po_dokumentu
    for x in P.akt:
        if x.izvor != OM_ULAZ or not x.broj or x.isp <= 0:
            continue
        noge = [a for t in DOK_REVERS_KOOP for a in amb.get((t, x.broj), ())]
        noge = [a for a in noge if not d(a, "Datum") or not x.dan or d(a, "Datum") == x.dan]
        if not noge or not all(storno(a) for a in noge):
            continue
        ponovo = [y for y in P.akt if y.izvor == OM_ULAZ and y.idn > x.idn and y.koop == x.koop and y.tip == x.tip
                  and abs(y.isp - x.isp) <= TOL and y.dan and x.dan and 0 <= (y.dan - x.dan).days <= 30]
        # GRESKA samo kad je ponovljen CEO unos (novi Ulaz OM sa aktivnim reversom istog smera, do 3 dana):
        # isti iznos kasnije bez reversa moze biti redovan fiksni avans kooperantu.
        tipovi = {k(a, "DokumentTip") for a in noge}
        ceo = [y for y in ponovo if y.broj and y.broj != x.broj and (y.dan - x.dan).days <= 3
               and any(not storno(a) for t in tipovi for a in amb.get((t, y.broj), ()))]
        P.ctx.nalaz(K, GRESKA if ceo else SUMNJA, "tblNovac", x.id,
                    "Revers %s (%s) je storniran (noge %s), a novac unet istim Ulazom OM %s je aktivan%s." % (
                        x.broj_txt, "/".join(sorted({s(a, "DokumentTip") for a in noge})),
                        _lista(s(a, "AmbID") for a in noge), x.kratko(),
                        ("; ceo unos (revers + novac) je ponovljen: %s" % _lista(y.kratko() for y in ceo)) if ceo else
                        ("; isti iznos je kasnije ponovo isplacen (moguc i redovan avans): %s" % _lista(
                            y.kratko() for y in ponovo)) if ponovo else ""),
                    ("Storno %s (ponovljen unos)." % x.id) if ceo else
                    "Ako je ceo unos bio pogresan: storno %s; ako je samo ambalaza bila pogresna, novac ostaje." % x.id)


# =====================================================================================
# PROPUSTENO 4) UPLATA_OM_NEVIDLJIVA_U_SALDU
# =====================================================================================
def _uplata_om(P):
    K = "UPLATA_OM_NEVIDLJIVA_U_SALDU"
    if not P.spremno(K, [("tblNovac", NOV_BAZA + ["Uplata", "OMID"])]):
        return
    for x in P.akt:
        if x.izvor != IZVOD or x.etip != "om" or x.tip != T_KES_FO or x.upl <= 0 or x.koop:
            continue
        P.ctx.nalaz(K, SUMNJA, "tblNovac", x.id,
                    "Izvodna UPLATA %s od OM %s knjizena kao KesFirmaOtkupac (Match:%s): saldo OM (GetOMAvansSaldo) "
                    "sabira samo Isplata, pa povracaj nije smanjio raspolozivi avans OM-a." % (
                        x.kratko(), x.omid_txt, x.match or "-"),
                    "Potvrditi da je uplatilac stvarno OM %s (automatski fallback po nazivu stanice moze pogoditi "
                    "kupca); u novoj svesci povracaj OM-a smanjuje saldo OM." % x.omid_txt)


# =====================================================================================
# PROPUSTENO 5) POZICIONI_POMAK_TBLNOVAC
# =====================================================================================
def _kolone_tabele(P, ime):
    kol = P.ctx.kolone.get(ime)
    if kol is not None:
        return kol
    tb = tabela(getattr(P.ctx, "_tabele", {}) or {}, ime)
    return [c.casefold() for c in tb.kolone] if tb is not None else None


def _pozicioni(P):
    K = "POZICIONI_POMAK_TBLNOVAC"
    for ime, ocek, pk in (("tblNovac", OCEK_NOV, "NovacID"), ("tblPartnerMap", OCEK_PM, "BankaName")):
        redovi = P.redovi(ime)
        kol = _kolone_tabele(P, ime)
        if not redovi or kol is None:
            continue
        n_ocek = len(ocek)
        if ime == "tblNovac" and "osirocenood" not in kol:
            n_ocek -= 1          # bez OsirocenoOd AppendRow skracuje niz (modDataAccess.AppendRow)
        lose_poz = [(i + 1, kol[i] if i < len(kol) else "-", ocek[i]) for i in range(n_ocek)
                    if i >= len(kol) or kol[i] != ocek[i]]
        anom = defaultdict(list)
        if ime == "tblNovac":
            for x in P.nov:
                r = x.r
                if x.tip not in TIPOVI:
                    anom["Tip van 8 poznatih vrednosti"].append(x.id or "#%s" % r.get("_excel_red", "?"))
                if k(r, "Stornirano") not in ("", "da"):
                    anom["Stornirano van {'', 'Da'}"].append(x.id)
                if x.dan is None:
                    anom["Datum nije datum"].append(x.id)
                if x.otk and not _OTK_ID.match(x.otk) and x.otk not in P.otk_po_id:
                    anom["OtkupID ne lici na OTK-"].append(x.id)
                for kol_b in ("Isplata", "Uplata"):
                    if tekst(polje(r, kol_b)) and _broj(polje(r, kol_b)) is None:
                        anom["%s nije broj" % kol_b].append(x.id)
        else:
            pref = {"kooperant": "koop-", "kupac": "kup-", "om": "st-"}
            for r in redovi:
                et = k(r, "EntitetTip")
                if et not in pref:
                    anom["EntitetTip van {Kupac, Kooperant, OM}"].append(s(r, pk))
                elif k(r, "PartnerID") and not k(r, "PartnerID").startswith(pref[et]):
                    anom["PartnerID ne odgovara EntitetTip"].append(s(r, pk))
        if lose_poz:
            P.ctx.nalaz(K, GRESKA if anom else SUMNJA, ime, "",
                        "Zaglavlje %s ne prati pozicioni pisac (%s): %s%s." % (
                            ime, "SaveNovac" if ime == "tblNovac" else "savePartnerMap",
                            "; ".join("poz. %d = %s, ocekivano %s" % p for p in lose_poz[:6]),
                            "; u podacima: " + "; ".join("%s %d" % (a, len(v)) for a, v in anom.items())
                            if anom else ""),
                        "Redovi upisani posle pomeranja kolona imaju vrednosti u pogresnim kolonama; vratiti redosled "
                        "i premestiti vrednosti po pozicijama pisca.")
        for a, v in anom.items():
            if lose_poz:
                break
            P.ctx.nalaz(K, SUMNJA, ime, _lista(v),
                        "%s: %s (%d redova); zaglavlje je u redu, pa je verovatna rucna izmena." % (ime, a, len(v)),
                        "Ispraviti vrednosti pre migracije.")


# =====================================================================================
# PROPUSTENO 6) IZVOD_DEDUPE_NEPROVERLJIV (INFO)
# =====================================================================================
def _dedupe_info(P):
    K = "IZVOD_DEDUPE_NEPROVERLJIV"
    if not P.bim or not P.ctx.ima_kolonu("tblBankaImport", "BrojRacuna"):
        return
    po_rac = defaultdict(set)
    for b in P.bim:
        po_rac[s(b, "BrojRacuna") or "(prazno)"].add(s(b, "BrojDokumenta"))
    delovi = []
    for rac, brojevi in sorted(po_rac.items()):
        br = sorted({int(x) for x in brojevi if x.isdigit()})
        bs = set(br)
        rupe = [i for i in range(br[0], br[-1] + 1) if i not in bs] if 1 < len(br) and br[-1] - br[0] <= 5000 else []
        delovi.append("%s: %d izvoda%s" % (rac, len(brojevi), ", rupe u nizu %s" % _lista(str(i) for i in rupe)
                                            if rupe else ""))
    P.ctx.nalaz(K, INFO, "tblBankaImport", "",
                "DISTINCT BrojRacuna = %d (%s). Dedupe izvoda bez BrojRacuna (pre RF-09) je mogao tiho izbaciti "
                "transakcije drugog racuna; to se u podacima ne vidi." % (len(po_rac), "; ".join(delovi)),
                "Uporediti broj izvoda po racunu sa izvodima banke; rupe u nizu su kandidati za ponovni uvoz.")


# =====================================================================================
PROVERE = (_pozicioni, _dupla, _osirocena, _preplata, _isplaceno_posle_storna, _veza, _fak_uplata_storn,
           _uplata_kupca, _odvezana_storno_fakture, _izvod_kupac_om, _pogresan_kooperant, _tip_isplate,
           _zastareo_datum, _datum_lokal, _storno_bez_lanca, _razdvojena, _dupli_rucni, _izvod_van_knjige,
           _bez_poziva_avans, _revers, _uplata_om, _dedupe_info)


def provere(ctx):
    P = _P(ctx)
    if not P.nov and not P.bim:
        return
    for f in PROVERE:
        try:
            f(P)
        except Exception:  # noqa: BLE001 -- pad jedne provere ne sme da ugasi ostale
            ctx.nalaz("MODUL_PAO", GRESKA, "", "novac." + f.__name__.lstrip("_"), traceback.format_exc(limit=6))
