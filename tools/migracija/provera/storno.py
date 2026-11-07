"""Provere storna u staroj svesci (C002, vba-v2.2.3 .. vba-v2.28.4) -- SAMO CITA.

Stari storno je oznacavao redove (Stornirano = "Da") i birao ih PO BROJU, bez
vlasnika; veze (OtpremnicaID na otkupu, BrojZbirne na otkupu i prijemnici) nije
cistio, a sync je oslobodjene blokove ponovo pakovao. Klase (kod nalaza = ID
klase iz specifikacije, dokaz u starom kodu):

  OTP_AKTIVNA_SA_STORNIRANIM_BLOKOM   StornoOtkup oznaci red, OtpremnicaID ostaje,
                                      otpremnica se ne rekalkulise (modStorno.bas
                                      v2.2.3 :139-150; panel v2.28.4 modOtkupBlok :850-855).
  PRJ_STORNO_PO_BROJU_TUDJI_KUPAC     StornoPrijemnicaByBroj_TX bira po BrojPrijemnice,
                                      bez KupacID (v2.2.3 modStorno.bas :468-475).
  KOLATERALNI_STORNO_PO_BROJU         StornoOtpremnicaByBroj_TX / StornoZbirna po broju,
                                      bez stanice, vozaca i datuma (v2.11.0 :272-347).
  OTK_STORNO_PO_BROJU_TUDJI_BLOK      StornoOtkupByBrDok_TX po broju, bez stanice i
                                      kooperanta (v2.28.4 modStorno.bas :105-116).
  ZBR_PONISTENJE_OBARA_TUDJU_PRIJEMNICU  PONISTENJE ide kaskadom po BrojZbirne
                                      (PonistiZbirnaChain_TX, v2.28.4 modStornoFlow :1000-1040).
  AUTO_USKRSNUCE_POSLE_STORNA         DUPLI/PONISTENJE oslobodi blokove (FreeOtkupBloksInline),
                                      sync ih spakuje u NOVU otpremnicu (modMasterSync :666-667).
  DOK_DELIMICNO_STORNIRAN             dokument sa dva reda (Klasa I/II), storniran samo jedan.
  OTK_BROJZBIRNE_NA_STORNIRANOJ_ZBIRNOJ  StornoZbirna ne dira tblOtkup.BrojZbirne (INFO).
  FAK_OSIROCENO_ZASTARELO             Relink brise OsirocenoOd samo na liniji, ne na
                                      fakturi (v2.28.4 modDokumenta :1949-1950) (INFO).
  STORNO_KONTEKST_NEDOVRSEN           tblStornoVeze: PENDING / MANUAL_REQUIRED / FAILED /
                                      RESI_KASNIJE -- trag koji migracija ne prenosi.

Pravila: tabela koje nema ili je prazna -> provera se tiho preskace; tabela
postoji a kolona fali -> jedan INFO "<KOD>_PRESKOCENO". Jedan nalaz po
dokumentu po klasi (redovi Klase I i II istog dokumenta spojeni).
"""

from bisect import bisect_left, bisect_right
from collections import defaultdict
from datetime import date, datetime, timedelta
from decimal import Decimal

from citac import broj as _broj
from citac import kljuc, polje, tekst

from provera.zajednicko import GRESKA, INFO, SUMNJA, aktivni, cifre, d, dan_txt, k, kg_txt, n, s, storno

TOL = Decimal("0.01")
ISTI_TRENUTAK_S = 3          # storno po broju je jedna transakcija: redovi se oznace u istoj sekundi
ZAJEDNO_SNIMLJENO_S = 10     # SaveXMulti_TX pise redove Klase I i II odjednom
MAX_ID = 12
MASOVNI_PECAT = 1        # vise od 1 tudjeg reda sa istim ModifiedAt: masovna izmena (kopiranje, backfill), ne storno po broju

MODE_ISPRAVKA = "ispravka_odmah"
MODE_DUPLI = "dupli_fantom"
MODE_PONISTENJE = "ponistenje_bez_zamene"
MODE_KASNIJE = "resi_kasnije"
ST_PENDING = "pending"
ST_COMPLETED = "completed"
ST_FAILED = "failed"
ST_MANUAL = "manual_required"
ST_CANCELLED = "cancelled"

# tabela: (PK, kolone vlasnika bez broja, broj) -- isto kao tools/migracija/dokumenti.GRUPE
DOK = {
    "tblOtkup": ("OtkupID", ("KooperantID", "StanicaID", "Datum"), "BrojDokumenta"),
    "tblOtpremnica": ("OtpremnicaID", ("StanicaID", "VozacID", "Datum"), "BrojOtpremnice"),
    "tblZbirna": ("ZbirnaID", ("VozacID", "Datum"), "BrojZbirne"),
    "tblPrijemnica": ("PrijemnicaID", ("KupacID", "Datum"), "BrojPrijemnice"),
}
NAZIV = {"tblOtkup": "otkup", "tblOtpremnica": "otpremnica", "tblZbirna": "zbirna",
         "tblPrijemnica": "prijemnica"}

_TS_FORMATI = ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S.%f", "%d.%m.%Y %H:%M:%S",
               "%Y-%m-%d %H:%M", "%d.%m.%Y %H:%M", "%d/%m/%Y %H:%M:%S")
_EXCEL_NULA = datetime(1899, 12, 30)


# --- pomocne ---------------------------------------------------------------
def _vreme(v):
    """datetime SA vremenom ili None. Sam datum nije dovoljan za 'isti trenutak'."""
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return None
    t = tekst(v)
    if not t:
        return None
    b = _broj(t)
    if b is not None:
        if Decimal(20000) <= b < Decimal(100000) and b != b.to_integral_value():
            return _EXCEL_NULA + timedelta(days=float(b))
        return None
    for f in _TS_FORMATI:
        try:
            return datetime.strptime(t, f)
        except ValueError:
            continue
    return None


def _suma(redovi, kol):
    return sum((n(r, kol) for r in redovi), Decimal(0))


def _dan(r):
    return d(r, "Datum")


def _vrednost_kljuca(r, kol):
    if kol == "Datum":
        x = d(r, kol)
        return x.isoformat() if x else k(r, kol)
    return k(r, kol)


def _dok_kljuc(ime, r, ima_gen):
    """Logicki dokument kao dokumenti.identitet_dokumenata: GeneracijaID, inace
    vlasnik + Datum + broj. None kad red nema broj (ne grupise se)."""
    pk, vlasnik, br = DOK[ime]
    if ima_gen:
        g = k(r, "GeneracijaID")
        if g:
            return ("GEN", g)
    b = k(r, br)
    if not b:
        return None
    return ("HEUR",) + tuple(_vrednost_kljuca(r, c) for c in vlasnik) + (b,)


def _lista_id(redovi, pk):
    vidjeno, out = set(), []
    for r in redovi:
        v = s(r, pk)
        if v and v.casefold() not in vidjeno:
            vidjeno.add(v.casefold())
            out.append(v)
    if len(out) > MAX_ID:
        return ", ".join(out[:MAX_ID]) + " (+%d)" % (len(out) - MAX_ID)
    return ", ".join(out)


def _kg_gaj(kg, gaj=None):
    if gaj is None:
        return "%s kg" % kg_txt(kg)
    return "%s kg / %s gajbi" % (kg_txt(kg), kg_txt(gaj))


def _klase(redovi):
    out = []
    for r in redovi:
        x = s(r, "Klasa")
        if x and x not in out:
            out.append(x)
    return "/".join(out) or "?"


def _isti_trenutak(a, b):
    """Prvi trenutak (ModifiedAt) u kome su izmenjeni i neki red iz a i neki red iz b, ili None."""
    ta = [x for x in (_vreme(polje(r, "ModifiedAt")) for r in a) if x]
    tb = [x for x in (_vreme(polje(r, "ModifiedAt")) for r in b) if x]
    for x in ta:
        for y in tb:
            if abs((x - y).total_seconds()) <= ISTI_TRENUTAK_S:
                return x
    return None


def _opis_dok(ime, redovi):
    """Kratak opis dokumenta (jednog ili vise redova istog dokumenta)."""
    r0 = redovi[0]
    pk, _vl, br = DOK[ime]
    ids_ = _lista_id(redovi, pk)
    if ime == "tblOtkup":
        return "otkup %s (%s; %s, %s, %s, %s)" % (
            s(r0, br), ids_, s(r0, "KooperantID"), s(r0, "StanicaID"), dan_txt(polje(r0, "Datum")),
            _kg_gaj(_suma(redovi, "Kolicina")))
    if ime == "tblOtpremnica":
        return "otpremnica %s (%s; %s, vozac %s, %s, kl. %s, %s)" % (
            s(r0, br), ids_, s(r0, "StanicaID"), s(r0, "VozacID"), dan_txt(polje(r0, "Datum")), _klase(redovi),
            _kg_gaj(_suma(redovi, "Kolicina"), _suma(redovi, "KolAmbalaze")))
    if ime == "tblZbirna":
        return "zbirna %s (%s; vozac %s, kupac %s, %s, kl. %s, %s)" % (
            s(r0, br), ids_, s(r0, "VozacID"), s(r0, "KupacID"), dan_txt(polje(r0, "Datum")), _klase(redovi),
            _kg_gaj(_suma(redovi, "UkupnoKolicina"), _suma(redovi, "UkupnoAmbalaze")))
    return "prijemnica %s (%s; kupac %s, vozac %s, zbirna %s, %s, kl. %s, %s)" % (
        s(r0, br), ids_, s(r0, "KupacID"), s(r0, "VozacID"), s(r0, "BrojZbirne"), dan_txt(polje(r0, "Datum")),
        _klase(redovi), _kg_gaj(_suma(redovi, "Kolicina"), _suma(redovi, "KolAmbalaze")))


def _veliko(t):
    return t[:1].upper() + t[1:]


def _stanje(redovi):
    akt = aktivni(redovi)
    if not akt:
        return "storniran"
    return "aktivan" if len(akt) == len(redovi) else "storniran pa ponovo unet"


def _grupe(redovi, kol):
    g = defaultdict(list)
    for r in redovi:
        v = k(r, kol)
        if v:
            g[v].append(r)
    return g


class _Provera:
    def __init__(self, ctx):
        self.ctx = ctx
        self._preskoceno = set()
        self._prj_po_broju = None
        self._amb_po_dokid = None
        self._otk_po_broju = None
        self._sv = None
        self._mod_at = {}

    # --- citanje ---------------------------------------------------------
    def redovi(self, ime):
        return self.ctx.tab(ime) or []

    def spremno(self, kod, zahtevi):
        """True kad su sve tabele neprazne i imaju kolone. Prazna/nepostojeca
        tabela -> tiho False; kolona fali -> jedan INFO <kod>_PRESKOCENO."""
        for ime, _kolone in zahtevi:
            if not self.redovi(ime):
                return False
        fale = []
        for ime, kolone in zahtevi:
            for kol in kolone:
                if not self.ctx.ima_kolonu(ime, kol):
                    fale.append("%s.%s" % (ime, kol))
        if not fale:
            return True
        kl = (kod, tuple(fale))
        if kl not in self._preskoceno:
            self._preskoceno.add(kl)
            self.ctx.nalaz(kod + "_PRESKOCENO", INFO, fale[0].split(".")[0], "",
                           "Provera %s preskocena: nedostaje kolona %s." % (kod, ", ".join(fale)),
                           "Proveriti zaglavlje tabele u svesci.")
        return False

    def masovni_pecat(self, ime, trenutak, grupa):
        """True kad isti ModifiedAt (+-ISTI_TRENUTAK_S) nosi vise od MASOVNI_PECAT redova
        tabele van grupe: to je masovna izmena (kopiranje, backfill), a ne jedan storno."""
        if ime not in self._mod_at:
            self._mod_at[ime] = sorted(x for x in (_vreme(polje(r, "ModifiedAt")) for r in self.redovi(ime)) if x)
        lst = self._mod_at[ime]
        w = timedelta(seconds=ISTI_TRENUTAK_S)
        u_prozoru = bisect_right(lst, trenutak + w) - bisect_left(lst, trenutak - w)
        iz_grupe = sum(1 for r in grupa for x in (_vreme(polje(r, "ModifiedAt")),)
                       if x and abs((x - trenutak).total_seconds()) <= ISTI_TRENUTAK_S)
        return u_prozoru - iz_grupe > MASOVNI_PECAT

    def prj_po_broju(self):
        if self._prj_po_broju is None:
            self._prj_po_broju = _grupe(self.redovi("tblPrijemnica"), "BrojPrijemnice")
        return self._prj_po_broju

    def otk_po_broju(self):
        if self._otk_po_broju is None:
            self._otk_po_broju = _grupe(self.redovi("tblOtkup"), "BrojDokumenta")
        return self._otk_po_broju

    def amb_po_dokid(self):
        if self._amb_po_dokid is None:
            self._amb_po_dokid = _grupe(self.redovi("tblAmbalaza"), "DokumentID")
        return self._amb_po_dokid

    def sv_po_broju(self):
        """tblStornoVeze po (OldDocType, OldBroj) i po ParentBroj -- samo za opis dokaza."""
        if self._sv is None:
            po_old, po_parent = defaultdict(list), defaultdict(list)
            for r in self.redovi("tblStornoVeze"):
                if k(r, "OldBroj"):
                    po_old[(k(r, "OldDocType"), k(r, "OldBroj"))].append(r)
                if k(r, "ParentBroj"):
                    po_parent[k(r, "ParentBroj")].append(r)
            self._sv = (po_old, po_parent)
        return self._sv

    def sv_trag(self, tip, broj):
        rr = self.sv_po_broju()[0].get((tip, kljuc(broj)), [])
        if not rr:
            return ""
        return " tblStornoVeze: %s." % ", ".join(
            "%s %s/%s" % (s(r, "CorrectionID"), s(r, "Mode"), s(r, "Status")) for r in rr[:MAX_ID])

    def dok_po_broju(self, tip, broj, dok_id=""):
        """(ime tabele, PK, redovi) za dokument tblStornoVeze-a (OldDocType/NewDocType)."""
        kb = kljuc(broj)
        if tip == "otpremnica":
            return "tblOtpremnica", "OtpremnicaID", list(self.ctx.otp_po_broju.get(kb, [])) if kb else []
        if tip == "zbirna":
            return "tblZbirna", "ZbirnaID", list(self.ctx.zbr_po_broju.get(kb, [])) if kb else []
        if tip == "prijemnica":
            rr = list(self.prj_po_broju().get(kb, [])) if kb else []
            p0 = self.ctx.red("tblPrijemnica", dok_id) if dok_id else None
            if p0 is not None:
                rr = [r for r in rr if k(r, "KupacID") == k(p0, "KupacID")] or [p0]
            return "tblPrijemnica", "PrijemnicaID", rr
        if tip == "revers":
            return "tblAmbalaza", "AmbID", list(self.amb_po_dokid().get(kb or kljuc(dok_id), []))
        if tip == "otkup":
            return "tblOtkup", "OtkupID", list(self.otk_po_broju().get(kb, [])) if kb else []
        return "", "", []

    # --- OTP_AKTIVNA_SA_STORNIRANIM_BLOKOM -------------------------------
    def otp_sa_storniranim_blokom(self):
        KOD = "OTP_AKTIVNA_SA_STORNIRANIM_BLOKOM"
        ctx = self.ctx
        if not self.spremno(KOD, [("tblOtkup", ["OtkupID", "OtpremnicaID", "Kolicina", "Stornirano"]),
                                  ("tblOtpremnica", ["OtpremnicaID", "Kolicina", "Stornirano"])]):
            return
        gajbe = ctx.ima_kolonu("tblOtkup", "KolAmbalaze") and ctx.ima_kolonu("tblOtpremnica", "KolAmbalaze")
        nevezani = defaultdict(list)
        vezani_po_koop = defaultdict(list)
        for o in self.redovi("tblOtkup"):
            if storno(o):
                continue
            if k(o, "OtpremnicaID"):
                vezani_po_koop[k(o, "KooperantID")].append(o)
            else:
                nevezani[(k(o, "KooperantID"), k(o, "StanicaID"))].append(o)
        ima_gen = ctx.ima_kolonu("tblOtpremnica", "GeneracijaID")
        pogodci = defaultdict(list)
        for X in self.redovi("tblOtpremnica"):
            if storno(X):
                continue
            xid = k(X, "OtpremnicaID")
            clanovi = ctx.otk_po_otpremnici.get(xid, []) if xid else []
            S = [o for o in clanovi if storno(o)]
            if not S:
                continue
            A = [o for o in clanovi if not storno(o)]
            # stornirani blokovi koje nije zamenio aktivan blok istog kooperanta i kg na X
            slobodni, nez = list(A), []
            for o in S:
                par = next((a for a in slobodni if k(a, "KooperantID") == k(o, "KooperantID")
                            and abs(n(a, "Kolicina") - n(o, "Kolicina")) <= TOL), None)
                if par is None:
                    nez.append(o)
                else:
                    slobodni.remove(par)
            manjak = n(X, "Kolicina") - _suma(A, "Kolicina")
            if manjak <= TOL:
                continue
            izbor = next((sk for sk in (S, nez) if sk and abs(manjak - _suma(sk, "Kolicina")) <= TOL), None)
            if izbor is None:
                continue
            manjak_g = n(X, "KolAmbalaze") - _suma(A, "KolAmbalaze") if gajbe else None
            gajbe_ok = gajbe and abs(manjak_g - _suma(izbor, "KolAmbalaze")) <= TOL
            zamene, dupli = [], []
            for o in izbor:
                dn = _dan(o)
                for c in nevezani.get((k(o, "KooperantID"), k(o, "StanicaID")), []):
                    dc = _dan(c)
                    if abs(n(c, "Kolicina") - n(o, "Kolicina")) <= TOL and (
                            dn is None or dc is None or abs((dc - dn).days) <= 1):
                        zamene.append(c)
                for c in vezani_po_koop.get(k(o, "KooperantID"), []):
                    if abs(n(c, "Kolicina") - n(o, "Kolicina")) > TOL or k(c, "OtpremnicaID") == xid:
                        continue
                    Y = ctx.red("tblOtpremnica", s(c, "OtpremnicaID"))
                    if (Y is not None and not storno(Y) and k(Y, "StanicaID") == k(X, "StanicaID")
                            and _dan(Y) == _dan(X) and k(Y, "Klasa") == k(X, "Klasa")):
                        dupli.append((c, Y))
            kl = _dok_kljuc("tblOtpremnica", X, ima_gen) or ("ID", xid)
            pogodci[kl].append(dict(X=X, A=A, S=izbor, manjak=manjak, manjak_g=manjak_g,
                                    gajbe_ok=gajbe_ok, zamene=zamene, dupli=dupli))
        for hits in pogodci.values():
            delovi, predlozi = [], []
            for h in hits:
                X, S = h["X"], h["S"]
                g = h["manjak_g"] is not None
                bz = s(X, "BrojZbirne")
                delovi.append(
                    "Aktivna otpremnica %s (%s; %s, %s, kl. %s) ima %s, aktivni blokovi %s; razlika %s = "
                    "storniran(i) blok(ovi) %s koji i dalje nose OtpremnicaID%s." % (
                        s(X, "BrojOtpremnice"), s(X, "OtpremnicaID"), s(X, "StanicaID"), dan_txt(polje(X, "Datum")),
                        s(X, "Klasa") or "?",
                        _kg_gaj(n(X, "Kolicina"), n(X, "KolAmbalaze") if g else None),
                        _kg_gaj(_suma(h["A"], "Kolicina"), _suma(h["A"], "KolAmbalaze") if g else None),
                        _kg_gaj(h["manjak"], h["manjak_g"]),
                        "; ".join("%s (%s, %s)" % (s(o, "OtkupID"), s(o, "KooperantID"),
                                                    _kg_gaj(n(o, "Kolicina"), n(o, "KolAmbalaze") if g else None))
                                  for o in S),
                        ("; zbirna %s je precenjena za isti iznos" % bz) if bz else ""))
                if h["dupli"]:
                    delovi.append("Isti kooperant i kg su i na aktivnoj otpremnici %s (ista stanica, dan, klasa): "
                                  "ako je to ponovni unos storniranog bloka, roba je dva puta u lancu (moze biti i "
                                  "druga isporuka)." % ", ".join(
                                      "%s (%s)" % (s(Y, "BrojOtpremnice"), s(c, "OtkupID")) for c, Y in h["dupli"]))
                if h["zamene"]:
                    predlozi.append("Vezati nevezan otkup %s (isti kooperant i kg, dan +-1) za %s -- to je ponovni "
                                    "unos storniranog bloka." % (_lista_id(h["zamene"], "OtkupID"),
                                                                 s(X, "OtpremnicaID")))
                else:
                    novo_kg = n(X, "Kolicina") - h["manjak"]
                    t = "Kolicina %s: %s -> %s" % (s(X, "OtpremnicaID"), kg_txt(n(X, "Kolicina")), kg_txt(novo_kg))
                    if g:
                        t += ", KolAmbalaze: %s -> %s" % (kg_txt(n(X, "KolAmbalaze")),
                                                         kg_txt(n(X, "KolAmbalaze") - h["manjak_g"]))
                    if bz:
                        t += "; zbirnu %s umanjiti za %s" % (bz, _kg_gaj(h["manjak"], h["manjak_g"]))
                    predlozi.append(t + " (odluka operatera).")
            ozb = GRESKA if all(h["gajbe_ok"] for h in hits) else SUMNJA
            ctx.nalaz(KOD, ozb, "tblOtpremnica", _lista_id([h["X"] for h in hits], "OtpremnicaID"),
                      " ".join(delovi), " ".join(predlozi))

    # --- storno po broju (PRJ_..., KOLATERALNI_..., OTK_...) -------------
    def _po_broju(self, kod, ime, vlasnik_kol, zamena, ziv, obavezno_ziv, sv_tip, predlog, kg_kol="Kolicina"):
        """Grupe istog broja sa >= 2 vlasnika koji OBA imaju stornirane redove:
        storno po broju pogadja sve aktivne redove tog broja. Kandidat je vlasnik
        ciji je dokument ceo storniran, bez zamene (i sa zivim nastavkom lanca)."""
        pk, _vl, br = DOK[ime]
        grupe = _grupe(self.redovi(ime), br)
        for b, redovi in grupe.items():
            if len(redovi) < 2 or not any(storno(r) for r in redovi):
                continue
            po_vl = defaultdict(list)
            for r in redovi:
                po_vl[tuple(_vrednost_kljuca(r, c) for c in vlasnik_kol)].append(r)
            sa_stornom = [v for v, rr in po_vl.items() if any(storno(r) for r in rr)]
            if len(po_vl) < 2 or len(sa_stornom) < 2:
                continue
            kandidati = []
            for v in sa_stornom:
                rr = po_vl[v]
                if aktivni(rr) or zamena(rr):
                    continue
                # ispravka vlasnika pod istim brojem: aktivan red drugog vlasnika, isti dan i kg
                drugi_akt = [r for v2 in po_vl if v2 != v for r in aktivni(po_vl[v2])]
                if drugi_akt and all(any(_dan(c) == _dan(r) and abs(n(c, kg_kol) - n(r, kg_kol)) <= TOL
                                         for c in drugi_akt) for r in rr):
                    continue
                zivi = ziv(rr)
                if obavezno_ziv and not zivi:
                    continue
                kandidati.append((v, rr, zivi))
            for v, rr, zivi in kandidati:
                ostali = [r for v2 in sa_stornom if v2 != v for r in po_vl[v2] if storno(r)]
                trenutak = _isti_trenutak(rr, ostali)
                if trenutak is not None and self.masovni_pecat(ime, trenutak, redovi):
                    trenutak = None
                ozb = GRESKA if trenutak and len(kandidati) == 1 else SUMNJA
                drugi = [po_vl[v2] for v2 in po_vl if v2 != v]
                opis = "Storniran(a) %s; isti broj %s ima i %s -- storno po broju ne gleda vlasnika." % (
                    _opis_dok(ime, rr), s(rr[0], br),
                    "; ".join("%s [%s]" % (_opis_dok(ime, aktivni(x) or x), _stanje(x)) for x in drugi))
                if trenutak:
                    opis += " Oba su stornirana u istom trenutku (ModifiedAt), jednom operacijom."
                if zivi:
                    opis += " " + zivi
                if len(kandidati) == 1:
                    opis += " Drugi vlasnik ima ponovni unos ili zamenu, ovaj nema."
                opis += self.sv_trag(sv_tip, s(rr[0], br))
                self.ctx.nalaz(kod, ozb, ime, _lista_id(rr, pk), opis, predlog(rr))

    def storno_po_broju(self):
        ctx = self.ctx
        # -- prijemnica: storno po BrojPrijemnice, bez KupacID
        KOD = "PRJ_STORNO_PO_BROJU_TUDJI_KUPAC"
        if self.spremno(KOD, [("tblPrijemnica", ["PrijemnicaID", "BrojPrijemnice", "KupacID", "BrojZbirne",
                                                 "Stornirano"]),
                              ("tblZbirna", ["BrojZbirne", "Stornirano"])]):
            zbr_kupac = ctx.ima_kolonu("tblZbirna", "KupacID")

            def zbirne(rr):
                return sorted({k(r, "BrojZbirne") for r in rr if k(r, "BrojZbirne")})

            def zamena_prj(rr):
                kup = k(rr[0], "KupacID")
                return any(not storno(p) and k(p, "KupacID") == kup
                           for b in zbirne(rr) for p in ctx.prj_po_broju_zbirne.get(b, []))

            def ziv_prj(rr):
                kup = k(rr[0], "KupacID")
                for b in zbirne(rr):
                    zb = [z for z in aktivni(ctx.zbr_po_broju.get(b, []))
                          if not zbr_kupac or not k(z, "KupacID") or k(z, "KupacID") == kup]
                    if zb and not aktivni(ctx.prj_po_broju_zbirne.get(b, [])):
                        return "Zbirna %s (%s) je i dalje aktivna i nema nijednu aktivnu prijemnicu." % (
                            s(zb[0], "BrojZbirne"), _lista_id(zb, "ZbirnaID"))
                return ""

            def predlog_prj(rr):
                return ("Vratiti %s u aktivno (Stornirano = ''), uz njene stornirane redove tblAmbalaza "
                        "(DokumentID = PrijemnicaID) i fakture; odluka operatera." % _lista_id(rr, "PrijemnicaID"))

            self._po_broju(KOD, "tblPrijemnica", ("KupacID",), zamena_prj, ziv_prj, True, "prijemnica",
                           predlog_prj)

        # -- otpremnica i zbirna: storno po broju, bez stanice/vozaca i datuma
        KOD = "KOLATERALNI_STORNO_PO_BROJU"
        if self.spremno(KOD, [("tblOtpremnica", ["OtpremnicaID", "BrojOtpremnice", "StanicaID", "Datum",
                                                 "Stornirano"])]):
            akt = defaultdict(list)
            for r in self.redovi("tblOtpremnica"):
                if not storno(r):
                    akt[(k(r, "StanicaID"), _dan(r), k(r, "Klasa"))].append(r)

            def zamena_otp(rr):
                return all(any(abs(n(y, "Kolicina") - n(r, "Kolicina")) <= TOL
                               for y in akt.get((k(r, "StanicaID"), _dan(r), k(r, "Klasa")), []))
                           for r in rr)

            def ziv_otp(rr):
                blokovi = [o for r in rr for o in ctx.otk_po_otpremnici.get(k(r, "OtpremnicaID"), [])
                           if not storno(o)]
                if blokovi:
                    return "Aktivni blokovi %s i dalje su vezani za ovu otpremnicu." % _lista_id(blokovi, "OtkupID")
                prj = [p for r in rr for p in ctx.prj_po_broju_zbirne.get(k(r, "BrojZbirne"), [])
                       if not storno(p) and k(p, "VozacID") == k(r, "VozacID")]
                if prj:
                    return "Aktivna prijemnica %s je i dalje na njenoj zbirni %s." % (
                        _lista_id(prj, "PrijemnicaID"), s(rr[0], "BrojZbirne"))
                return ""

            def predlog_otp(rr):
                return ("Vratiti %s u aktivno (Stornirano = '') sa njenom ambalazom (tblAmbalaza DokumentID = "
                        "OtpremnicaID) i zbirnom %s, ili blokove prevezati na novu otpremnicu iste stanice; "
                        "odluka operatera." % (_lista_id(rr, "OtpremnicaID"), s(rr[0], "BrojZbirne") or "-"))

            self._po_broju(KOD, "tblOtpremnica", ("StanicaID", "Datum"), zamena_otp, ziv_otp, True, "otpremnica",
                           predlog_otp)

        if self.spremno(KOD, [("tblZbirna", ["ZbirnaID", "BrojZbirne", "VozacID", "Datum", "Stornirano"])]):
            akt_z = defaultdict(list)
            for z in self.redovi("tblZbirna"):
                if not storno(z):
                    akt_z[(k(z, "VozacID"), _dan(z), k(z, "Klasa"))].append(z)

            def zamena_zbr(rr):
                return all(any(abs(n(y, "UkupnoKolicina") - n(r, "UkupnoKolicina")) <= TOL
                               for y in akt_z.get((k(r, "VozacID"), _dan(r), k(r, "Klasa")), []))
                           for r in rr)

            def ziv_zbr(rr):
                b, voz = k(rr[0], "BrojZbirne"), k(rr[0], "VozacID")
                otp = [o for o in ctx.otp_po_broju_zbirne.get(b, []) if not storno(o) and k(o, "VozacID") == voz]
                prj = [p for p in ctx.prj_po_broju_zbirne.get(b, []) if not storno(p) and k(p, "VozacID") == voz]
                delovi = []
                if otp:
                    delovi.append("aktivne otpremnice %s" % _lista_id(otp, "OtpremnicaID"))
                if prj:
                    delovi.append("aktivna prijemnica %s" % _lista_id(prj, "PrijemnicaID"))
                return ("Na nju i dalje pokazuju %s (isti vozac)." % " i ".join(delovi)) if delovi else ""

            def predlog_zbr(rr):
                return ("Vratiti zbirnu %s u aktivno (Stornirano = '') i proveriti UkupnoKolicina prema aktivnim "
                        "otpremnicama; odluka operatera." % _lista_id(rr, "ZbirnaID"))

            self._po_broju(KOD, "tblZbirna", ("VozacID", "Datum"), zamena_zbr, ziv_zbr, True, "zbirna",
                           predlog_zbr, kg_kol="UkupnoKolicina")

        # -- otkup: StornoOtkupByBrDok_TX po broju, bez stanice i kooperanta
        KOD = "OTK_STORNO_PO_BROJU_TUDJI_BLOK"
        if self.spremno(KOD, [("tblOtkup", ["OtkupID", "BrojDokumenta", "KooperantID", "StanicaID", "Kolicina",
                                            "Stornirano"])]):
            akt_k = defaultdict(list)
            for o in self.redovi("tblOtkup"):
                if not storno(o):
                    akt_k[k(o, "KooperantID")].append(o)

            def zamena_otk(rr):
                return all(any(abs(n(c, "Kolicina") - n(r, "Kolicina")) <= TOL and _dan(c) == _dan(r)
                               for c in akt_k.get(k(r, "KooperantID"), []))
                           for r in rr)

            def ziv_otk(rr):
                otp = [ctx.red("tblOtpremnica", s(r, "OtpremnicaID")) for r in rr if k(r, "OtpremnicaID")]
                otp = [x for x in otp if x is not None and not storno(x)]
                if otp:
                    return "Blok je i dalje u aktivnoj otpremnici %s." % _lista_id(otp, "OtpremnicaID")
                return ""

            def predlog_otk(rr):
                return ("Vratiti %s u aktivno (Stornirano = '') sa ambalazom (tblAmbalaza DokumentID = OtkupID) i "
                        "vezom novca (tblNovac.OtkupID), ili ga uneti ponovo; odluka operatera."
                        % _lista_id(rr, "OtkupID"))

            self._po_broju(KOD, "tblOtkup", ("KooperantID", "StanicaID"), zamena_otk, ziv_otk, False, "otkup",
                           predlog_otk)

    # --- ZBR_PONISTENJE_OBARA_TUDJU_PRIJEMNICU ---------------------------
    def _dokumenti_prijemnica(self, prj):
        """(KupacID, BrojPrijemnice) -> redovi; isti broj istog kupca je jedan dokument
        i kad je storniran pa ponovo unet."""
        docs = defaultdict(list)
        for r in prj:
            docs[(k(r, "KupacID"), k(r, "BrojPrijemnice") or k(r, "PrijemnicaID"))].append(r)
        return docs

    def _prava_zbirna(self, rr, akt_po_vozacu):
        """Aktivne zbirne vozaca prijemnice, istog ili prethodnog dana, bez aktivne
        prijemnice -- rangirane po razlici kg."""
        p = rr[0]
        dan = _dan(p)
        kg = _suma(rr, "Kolicina")
        out = []
        for z in akt_po_vozacu.get(k(p, "VozacID"), []):
            dz = _dan(z)
            if dan is not None and dz is not None and not (0 <= (dan - dz).days <= 1):
                continue
            if aktivni(self.ctx.prj_po_broju_zbirne.get(k(z, "BrojZbirne"), [])):
                continue
            out.append((abs(n(z, "UkupnoKolicina") - kg), z))
        out.sort(key=lambda x: (x[0], s(x[1], "ZbirnaID")))
        return [(z, kg) for _r, z in out]

    def ponistenje(self):
        KOD = "ZBR_PONISTENJE_OBARA_TUDJU_PRIJEMNICU"
        ctx = self.ctx
        if not self.spremno(KOD, [("tblPrijemnica", ["PrijemnicaID", "BrojZbirne", "VozacID", "Stornirano"]),
                                  ("tblZbirna", ["BrojZbirne", "VozacID", "Stornirano"])]):
            return
        akt_po_vozacu = defaultdict(list)
        for z in self.redovi("tblZbirna"):
            if not storno(z):
                akt_po_vozacu[k(z, "VozacID")].append(z)

        def voz_zbirne(zb, b):
            v = {k(z, "VozacID") for z in zb if k(z, "VozacID")}
            if not v:
                st = ctx.stanica_iz_broja(b)
                v = {kljuc(st)} if st else set()
            return v

        def tudji_vozac(rr, vz, b):
            """Vozac prijemnice nije vozac zbirne B, a nije ni stanica iz prefiksa B
            (S<n>/..): otpremnica sa vozacem druge stanice (P4) daje zbirnu ciji se
            VozacID i prefiks razlikuju bez greske, pa jedno od dva mora da se poklopi."""
            v = k(rr[0], "VozacID")
            if not vz or not v or v in vz:
                return False
            return v != kljuc(ctx.stanica_iz_broja(b))

        def voz_prikaz(zb, b):
            v = sorted({s(z, "VozacID") for z in zb if s(z, "VozacID")})
            return ", ".join(v) or ctx.stanica_iz_broja(b) or "?"

        def predlog_prave(rr):
            kand = self._prava_zbirna(rr, akt_po_vozacu)
            if not kand:
                return ""
            return "%s vezati na pravu zbirnu: %s." % (_lista_id(rr, "PrijemnicaID"), "; ".join(
                "%s (%s, %s, kg %s vs %s)" % (s(z, "BrojZbirne"), s(z, "ZbirnaID"), dan_txt(polje(z, "Datum")),
                                             kg_txt(n(z, "UkupnoKolicina")), kg_txt(kg))
                for z, kg in kand[:3]))

        pokriveno = set()
        sv = self.redovi("tblStornoVeze")
        if sv and self.spremno(KOD, [("tblStornoVeze", ["Mode", "OldDocType", "OldBroj", "ParentBroj"])]):
            for R in sv:
                if k(R, "Mode") != MODE_PONISTENJE or k(R, "Status") == ST_CANCELLED:
                    continue
                tip = k(R, "OldDocType")
                if tip in ("prijemnica", "otpremnica"):
                    B = s(R, "ParentBroj")
                elif tip == "zbirna":
                    B = s(R, "OldBroj")
                else:
                    continue
                kb = kljuc(B)
                zb = ctx.zbr_po_broju.get(kb, []) if kb else []
                if not zb or aktivni(zb):
                    continue        # kaskada se nije desila (ili je zbirna kasnije ponovo uneta)
                pokriveno.add(kb)
                vz = voz_zbirne(zb, B)
                docs = self._dokumenti_prijemnica(ctx.prj_po_broju_zbirne.get(kb, []))

                def tudja(rr, vz=vz, B=B):
                    return tudji_vozac(rr, vz, B)

                stornirani = [dk for dk, rr in docs.items() if not aktivni(rr)]
                cid = s(R, "CorrectionID")
                vp = voz_prikaz(zb, B)
                uvod = "PONISTENJE %s (%s %s) oborilo je tok zbirne %s (vozac %s): zbirna i otpremnice stornirane." % (
                    cid, s(R, "OldDocType"), s(R, "OldBroj"), B, vp)
                if tip == "prijemnica":
                    oid = k(R, "OldDocID")
                    p0 = next((dk for dk, rr in docs.items() if oid and any(k(r, "PrijemnicaID") == oid for r in rr)),
                              None)
                    if p0 is None:
                        p0 = next((dk for dk in docs if dk[1] == kljuc(s(R, "OldBroj"))), None)
                    if p0 is None:
                        continue
                    P0 = docs[p0]
                    ostali = [dk for dk in docs if dk != p0]
                    c_i = tudja(P0)
                    c_ii = [dk for dk in ostali if aktivni(docs[dk])]
                    c_iii = len(stornirani) > 1
                    if not (c_i or c_ii or c_iii):
                        continue
                    vlasnici = [dk for dk in ostali if not tudja(docs[dk])]
                    pp = predlog_prave(P0) if c_i else ""
                    ozb = GRESKA if c_i and vlasnici and pp else SUMNJA
                    delovi = [uvod]
                    if c_i:
                        delovi.append("Ponistena je %s: vozac %s, a zbirna %s pripada vozacu %s -- prijemnica je "
                                      "bila pogresno vezana." % (_opis_dok("tblPrijemnica", P0), s(P0[0], "VozacID"),
                                                                 B, vp))
                    for dk in ostali:
                        delovi.append("Na istoj zbirni je i %s -- %s." % (
                            _opis_dok("tblPrijemnica", docs[dk]),
                            "OSTALA AKTIVNA na storniranoj zbirni" if aktivni(docs[dk]) else "stornirana kaskadom"))
                    meta = [docs[dk] for dk in (vlasnici or ostali)] or [P0]
                    predlog = []
                    if vlasnici:
                        predlog.append("Vratiti u aktivno %s, zbirnu %s i njene otpremnice (blokove ponovo vezati), "
                                       "uz ambalazu i palete." % (
                                           ", ".join(_lista_id(docs[dk], "PrijemnicaID") for dk in vlasnici), B))
                    if pp:
                        predlog.append(pp)
                    predlog.append("Odluka operatera.")
                    ctx.nalaz(KOD, ozb, "tblPrijemnica", _lista_id([r for rr in meta for r in rr], "PrijemnicaID"),
                              " ".join(delovi), " ".join(predlog))
                else:
                    tudje = [dk for dk in stornirani if tudja(docs[dk])]
                    if not tudje and len(stornirani) < 2:
                        continue
                    pogodjeni = tudje or stornirani
                    prave = [predlog_prave(docs[dk]) for dk in tudje]
                    ozb = GRESKA if any(prave) else SUMNJA
                    delovi = [uvod, "Kaskada je stornirala i: %s%s." % (
                        "; ".join(_opis_dok("tblPrijemnica", docs[dk]) for dk in stornirani),
                        (" -- vozac prijemnice nije vozac zbirne" if tudje else ""))]
                    predlog = ["Vratiti u aktivno %s." % ", ".join(_lista_id(docs[dk], "PrijemnicaID")
                                                                   for dk in pogodjeni)]
                    predlog += [x for x in prave if x]
                    predlog.append("Odluka operatera.")
                    ctx.nalaz(KOD, ozb, "tblPrijemnica",
                              _lista_id([r for dk in pogodjeni for r in docs[dk]], "PrijemnicaID"),
                              " ".join(delovi), " ".join(predlog))

        # bez traga u tblStornoVeze (nema tabele ili je zbirna oborena drugim putem)
        if not self.spremno(KOD, [("tblOtpremnica", ["BrojZbirne", "Stornirano"])]):
            return
        for kb, zb in ctx.zbr_po_broju.items():
            if kb in pokriveno or aktivni(zb):
                continue
            otp = ctx.otp_po_broju_zbirne.get(kb, [])
            if not otp or aktivni(otp):
                continue
            docs = self._dokumenti_prijemnica(ctx.prj_po_broju_zbirne.get(kb, []))
            if len(docs) < 2:
                continue
            bz0 = s(zb[0], "BrojZbirne")
            vz = voz_zbirne(zb, bz0)
            tudje = [dk for dk, rr in docs.items() if tudji_vozac(rr, vz, bz0)]
            if not tudje and len({(k(rr[0], "KupacID"), k(rr[0], "VozacID")) for rr in docs.values()}) < 2:
                continue        # ista stanica i kupac: storno pa ponovni unos prijemnice pod drugim brojem
            sumnjivi = [dk for dk, rr in docs.items() if aktivni(rr) or dk in tudje]
            if not sumnjivi:
                continue
            ctx.nalaz(KOD, SUMNJA, "tblPrijemnica",
                      _lista_id([r for dk in sumnjivi for r in docs[dk]], "PrijemnicaID"),
                      "Zbirna %s (vozac %s) i sve njene otpremnice su stornirane (bez traga u tblStornoVeze), a na "
                      "njoj su %d prijemnice: %s. Kaskada po BrojZbirne je mogla oboriti i prijemnicu druge "
                      "stanice." % (s(zb[0], "BrojZbirne"), voz_prikaz(zb, s(zb[0], "BrojZbirne")), len(docs),
                                    "; ".join("%s [%s]" % (_opis_dok("tblPrijemnica", rr),
                                                           "aktivna" if aktivni(rr) else "stornirana")
                                              for rr in docs.values())),
                      "Proveriti kojoj zbirni pripada svaka prijemnica (vozac, dan, kg); prijemnicu druge stanice "
                      "vratiti/vezati na njenu zbirnu. Odluka operatera.")

    # --- AUTO_USKRSNUCE_POSLE_STORNA -------------------------------------
    def auto_uskrsnuce(self):
        KOD = "AUTO_USKRSNUCE_POSLE_STORNA"
        ctx = self.ctx
        sv = self.redovi("tblStornoVeze")
        if not sv:
            return
        if not self.spremno(KOD, [("tblStornoVeze", ["Mode", "OldDocType", "OldBroj"]),
                                  ("tblOtpremnica", ["OtpremnicaID", "BrojOtpremnice", "StanicaID", "Datum",
                                                     "Kolicina", "Stornirano"])]):
            return
        otk_vozac = ctx.ima_kolonu("tblOtkup", "VozacID")
        akt = defaultdict(list)
        for A in self.redovi("tblOtpremnica"):
            if not storno(A):
                akt[(k(A, "StanicaID"), _dan(A), k(A, "Klasa"))].append(A)
        # aktivni nevezani blokovi po (stanica, dan, klasa): ako zbir jos daje kg stornirane
        # otpremnice, oslobodjeni blokovi nisu ponovo spakovani (sync jos nije prosao)
        slobodni_kg = defaultdict(Decimal)
        if ctx.ima_kolonu("tblOtkup", "OtpremnicaID"):
            for o in self.redovi("tblOtkup"):
                if not storno(o) and not k(o, "OtpremnicaID") and (not otk_vozac or k(o, "VozacID")):
                    slobodni_kg[(k(o, "StanicaID"), _dan(o), k(o, "Klasa"))] += n(o, "Kolicina")

        def manjak_drugih(A, iskljuci):
            """Aktivne otpremnice iste stanice, dana i klase sa manjkom blokova = kg od A."""
            out = []
            for y in akt.get((k(A, "StanicaID"), _dan(A), k(A, "Klasa")), []):
                if id(y) in iskljuci:
                    continue
                manjak = n(y, "Kolicina") - _suma(aktivni(ctx.otk_po_otpremnici.get(k(y, "OtpremnicaID"), [])),
                                                  "Kolicina")
                if manjak > TOL and abs(manjak - n(A, "Kolicina")) <= TOL:
                    out.append(y)
            return out
        ima_gen = ctx.ima_kolonu("tblOtpremnica", "GeneracijaID")
        hits = {}
        for R in sv:
            mode, tip = k(R, "Mode"), k(R, "OldDocType")
            if mode not in (MODE_DUPLI, MODE_PONISTENJE) or k(R, "Status") in (ST_CANCELLED, ST_FAILED):
                continue
            if tip == "otpremnica":
                S = ctx.otp_po_broju.get(k(R, "OldBroj"), [])
            elif tip == "zbirna" and mode == MODE_PONISTENJE:
                S = ctx.otp_po_broju_zbirne.get(k(R, "OldBroj"), [])
            elif tip == "prijemnica" and mode == MODE_PONISTENJE:
                S = ctx.otp_po_broju_zbirne.get(k(R, "ParentBroj"), [])
            else:
                continue
            sv_vreme = _vreme(polje(R, "CreatedAt"))
            for Sr in S:
                if not storno(Sr) or not k(Sr, "Datum"):
                    continue
                sid = cifre(s(Sr, "OtpremnicaID"))
                kljuc_s = (k(Sr, "StanicaID"), _dan(Sr), k(Sr, "Klasa"))
                if n(Sr, "Kolicina") > TOL and abs(slobodni_kg.get(kljuc_s, Decimal(0)) - n(Sr, "Kolicina")) <= TOL:
                    continue            # oslobodjeni blokovi su i dalje nevezani: nema uskrsnuca
                kand = []
                for A in akt.get(kljuc_s, []):
                    if abs(n(A, "Kolicina") - n(Sr, "Kolicina")) > TOL:
                        continue
                    if k(A, "BrojOtpremnice") == k(Sr, "BrojOtpremnice"):
                        continue
                    aid = cifre(s(A, "OtpremnicaID"))
                    if sid is None or aid is None or aid <= sid:
                        continue        # nova otpremnica je nastala POSLE stornirane
                    a_vreme = _vreme(polje(A, "CreatedAt"))
                    if sv_vreme and a_vreme and a_vreme < sv_vreme - timedelta(seconds=ISTI_TRENUTAK_S):
                        continue
                    blokovi = aktivni(ctx.otk_po_otpremnici.get(k(A, "OtpremnicaID"), []))
                    if not blokovi or abs(_suma(blokovi, "Kolicina") - n(A, "Kolicina")) > TOL:
                        continue        # sync pravi otpremnicu tacno od zbira blokova
                    if otk_vozac and any(not k(o, "VozacID") for o in blokovi):
                        continue        # sync pakuje samo blokove sa VozacID
                    bz_a = k(A, "BrojZbirne")
                    if bz_a and aktivni(ctx.prj_po_broju_zbirne.get(bz_a, [])) and not manjak_drugih(A, {id(A)}):
                        continue        # zbirna od A ima prijemnicu: tok je isporucen (npr. druga tura)
                    kand.append((aid, bool(sv_vreme and a_vreme), A, blokovi))
                if kand and not all(vr for _a, vr, _A, _b in kand):
                    kand = [max(kand, key=lambda x: x[0])]   # bez vremena: samo najnovija (sync je pravi posle)
                for _aid, _vr, A, blokovi in kand:
                    kl = _dok_kljuc("tblOtpremnica", A, ima_gen) or ("ID", k(A, "OtpremnicaID"))
                    h = hits.setdefault(kl, dict(A=[], S=[], R=[], blokovi=[]))
                    for lst, x in ((h["A"], A), (h["S"], Sr), (h["R"], R)):
                        if not any(x is y for y in lst):
                            lst.append(x)
                    h["blokovi"].extend(o for o in blokovi if not any(o is y for y in h["blokovi"]))
        for h in hits.values():
            A0 = h["A"][0]
            kg = _suma(h["A"], "Kolicina")
            uzeti = {id(x) for x in h["A"]}
            Y = []
            for A in h["A"]:
                Y.extend(y for y in manjak_drugih(A, uzeti) if not any(y is z for z in Y))
            bz = s(A0, "BrojZbirne")
            prj = aktivni(ctx.prj_po_broju_zbirne.get(kljuc(bz), [])) if bz else []
            modovi = sorted({s(R, "Mode") for R in h["R"]})
            opis = ("%s nastala je posle storna %s (%s), kojim je stornirana %s: blokovi %s, oslobodjeni tim "
                    "stornom, ponovo su spakovani u novu otpremnicu (sync AutoCreateOtpremniceFromPWA)." % (
                        _veliko(_opis_dok("tblOtpremnica", h["A"])), "/".join(modovi),
                        ", ".join("%s %s %s" % (s(R, "CorrectionID"), s(R, "OldDocType"), s(R, "OldBroj"))
                                  for R in h["R"]),
                        "; ".join(_opis_dok("tblOtpremnica", [x]) for x in h["S"]),
                        _lista_id(h["blokovi"], "OtkupID")))
            if bz:
                opis += " Njena zbirna %s %s." % (bz, "ima prijemnicu %s" % _lista_id(prj, "PrijemnicaID")
                                                  if prj else "nema aktivnu prijemnicu")
            if Y:
                opis += (" Aktivna otpremnica %s iste stanice, dana i klase ima manjak blokova od tacno %s: "
                         "roba je dva puta u lancu." % (_lista_id(Y, "OtpremnicaID"), kg_txt(kg)))
                predlog = ("Blokove %s prevezati na %s, a %s%s stornirati; odluka operatera." % (
                    _lista_id(h["blokovi"], "OtkupID"), _lista_id(Y, "OtpremnicaID"), _lista_id(h["A"], "OtpremnicaID"),
                    (" i zbirnu %s" % bz) if bz else ""))
            elif MODE_PONISTENJE.upper() in [m.upper() for m in modovi]:
                predlog = ("Ponistenje je trazilo da tok nestane, sync ga je vratio pod novim brojem. Odluka "
                           "operatera: stornirati %s%s (blokovi ostaju nevezani) ili ih zadrzati i dodati "
                           "prijemnicu." % (_lista_id(h["A"], "OtpremnicaID"), (" i zbirnu %s" % bz) if bz else ""))
            else:
                predlog = ("Blokovi su oslobodjeni radi prevezivanja na pravu otpremnicu: prevezati ih tamo i "
                           "stornirati %s%s, ili zadrzati ako je to prava otpremnica; odluka operatera." % (
                               _lista_id(h["A"], "OtpremnicaID"), (" i zbirnu %s" % bz) if bz else ""))
            ctx.nalaz(KOD, GRESKA if Y else SUMNJA, "tblOtpremnica", _lista_id(h["A"], "OtpremnicaID"), opis,
                      predlog)

    # --- DOK_DELIMICNO_STORNIRAN -----------------------------------------
    def delimicno(self):
        KOD = "DOK_DELIMICNO_STORNIRAN"
        ctx = self.ctx
        for ime, (pk, _vl, br) in DOK.items():
            if not self.spremno(KOD, [(ime, [pk, br, "Datum", "Klasa", "Stornirano"])]):
                continue
            ima_gen = ctx.ima_kolonu(ime, "GeneracijaID")
            grupe = defaultdict(list)
            for r in self.redovi(ime):
                kl = _dok_kljuc(ime, r, ima_gen)
                if kl is not None:
                    grupe[kl].append(r)
            for rr in grupe.values():
                if len(rr) < 2:
                    continue
                akt = aktivni(rr)
                st = [r for r in rr if storno(r)]
                if not akt or not st:
                    continue
                if {k(r, "Klasa") for r in akt} & {k(r, "Klasa") for r in st}:
                    continue        # ista klasa i aktivna i stornirana: storno pa ponovni unos istog broja
                zajedno = self._snimljeni_zajedno(rr, pk)
                if not zajedno:
                    continue        # nema dokaza da su redovi jedan unos (moze biti ponovni unos druge klase)
                ctx.nalaz(KOD, SUMNJA, ime, _lista_id(rr, pk),
                          "%s %s: storniran red kl. %s (%s), aktivan red kl. %s (%s); redovi su jedan unos (%s), "
                          "storniran je samo deo dokumenta." % (
                              NAZIV[ime].capitalize(), s(rr[0], br), _klase(st), _lista_id(st, pk), _klase(akt),
                              _lista_id(akt, pk), zajedno),
                          "Odluka operatera: stornirati i preostali red (%s) ili ga zadrzati kao samostalan dokument."
                          % _lista_id(akt, pk))

    @staticmethod
    def _snimljeni_zajedno(rr, pk):
        vremena = [_vreme(polje(r, "CreatedAt")) for r in rr]
        if all(vremena):
            if (max(vremena) - min(vremena)).total_seconds() <= ZAJEDNO_SNIMLJENO_S:
                return "isti CreatedAt"
            return ""
        brojevi = [cifre(s(r, pk)) for r in rr]
        if all(b is not None for b in brojevi) and len(set(brojevi)) == len(brojevi) \
                and max(brojevi) - min(brojevi) == len(brojevi) - 1:
            return "uzastopni ID-jevi"
        return ""

    # --- OTK_BROJZBIRNE_NA_STORNIRANOJ_ZBIRNOJ ---------------------------
    def otk_brojzbirne_storno(self):
        KOD = "OTK_BROJZBIRNE_NA_STORNIRANOJ_ZBIRNOJ"
        ctx = self.ctx
        if not self.spremno(KOD, [("tblOtkup", ["OtkupID", "BrojZbirne", "Stornirano"]),
                                  ("tblZbirna", ["BrojZbirne", "Stornirano"])]):
            return
        po_b = defaultdict(list)
        for o in self.redovi("tblOtkup"):
            b = k(o, "BrojZbirne")
            if storno(o) or not b:
                continue
            zb = ctx.zbr_po_broju.get(b, [])
            if not zb or aktivni(zb):
                continue
            P = ctx.red("tblOtpremnica", s(o, "OtpremnicaID")) if k(o, "OtpremnicaID") else None
            if P is not None and storno(P):
                continue        # blok na storniranoj otpremnici: to je klasa modula otkup
            po_b[b].append((o, P))
        for b, par in po_b.items():
            oo = [o for o, _P in par]
            otp = [P for _o, P in par if P is not None]
            zb = ctx.zbr_po_broju.get(b, [])
            opis = ("Aktivni otkupi %s nose BrojZbirne %s, a zbirna tog broja postoji samo stornirana (%s); "
                    "StornoZbirna ne cisti tblOtkup.BrojZbirne." % (_lista_id(oo, "OtkupID"), s(oo[0], "BrojZbirne"),
                                                                    _lista_id(zb, "ZbirnaID")))
            if otp:
                opis += " Otpremnica bloka: %s." % "; ".join(
                    "%s BrojZbirne '%s'" % (s(P, "OtpremnicaID"), s(P, "BrojZbirne")) for P in otp[:MAX_ID])
            ctx.nalaz(KOD, INFO, "tblOtkup", _lista_id(oo, "OtkupID"), opis,
                      "Denormalizovana kopija: uskladiti sa BrojZbirne otpremnice ili isprazniti (migracija je ne "
                      "prenosi, cilj.py). Dokaz zaostalog konteksta, ne roba.")

    # --- FAK_OSIROCENO_ZASTARELO -----------------------------------------
    def fak_osiroceno_zastarelo(self):
        KOD = "FAK_OSIROCENO_ZASTARELO"
        ctx = self.ctx
        if not self.spremno(KOD, [("tblFakture", ["FakturaID", "OsirocenoOd"]),
                                  ("tblFakturaStavke", ["FakturaID", "PrijemnicaID", "OsirocenoOd"]),
                                  ("tblPrijemnica", ["PrijemnicaID", "Stornirano"])]):
            return
        linije = defaultdict(list)
        for ln in self.redovi("tblFakturaStavke"):
            if not storno(ln) and k(ln, "FakturaID"):
                linije[k(ln, "FakturaID")].append(ln)
        for F in self.redovi("tblFakture"):
            if storno(F) or not s(F, "OsirocenoOd"):
                continue
            ll = linije.get(k(F, "FakturaID"), [])
            if not ll or any(s(ln, "OsirocenoOd") for ln in ll):
                continue
            prj = [ctx.red("tblPrijemnica", s(ln, "PrijemnicaID")) for ln in ll]
            if any(p is None or storno(p) for p in prj):
                continue
            ctx.nalaz(KOD, INFO, "tblFakture", s(F, "FakturaID"),
                      "Faktura %s nosi OsirocenoOd '%s', a sve njene aktivne linije (%d) su prevezane na aktivne "
                      "prijemnice (%s) i nemaju OsirocenoOd: oznaka na fakturi je zastarela (Relink je brise samo na "
                      "liniji)." % (s(F, "BrojFakture") or s(F, "FakturaID"), s(F, "OsirocenoOd"), len(ll),
                                    _lista_id(prj, "PrijemnicaID")),
                      "OsirocenoOd na fakturi isprazniti; nije greska u robi, nego izvor laznih nalaza za linije "
                      "fakture.")

    # --- STORNO_KONTEKST_NEDOVRSEN ---------------------------------------
    def kontekst_nedovrsen(self):
        KOD = "STORNO_KONTEKST_NEDOVRSEN"
        ctx = self.ctx
        sv = self.redovi("tblStornoVeze")
        if not sv or not self.spremno(KOD, [("tblStornoVeze", ["CorrectionID", "Mode", "Status"])]):
            return
        nas_trag = {c: (ctx.ima_kolonu(ime, "IspravkaOd"), ctx.ima_kolonu(ime, "CorrectionID"))
                    for c, ime in (("otpremnica", "tblOtpremnica"), ("zbirna", "tblZbirna"),
                                   ("prijemnica", "tblPrijemnica"), ("otkup", "tblOtkup"))}
        # zavrseni konteksti po dokumentu: kasniji COMPLETED nad istim dokumentom resava raniji otvoren
        zavrseni = defaultdict(list)
        for i, R in enumerate(sv):
            if k(R, "Status") == ST_COMPLETED and k(R, "OldBroj"):
                zavrseni[(k(R, "OldDocType"), k(R, "OldBroj"))].append((i, R))

        def kasniji_zavrsen(i, R):
            t0 = _vreme(polje(R, "CreatedAt"))
            c0 = cifre(s(R, "CorrectionID"))
            for j, R2 in zavrseni.get((k(R, "OldDocType"), k(R, "OldBroj")), []):
                if R2 is R:
                    continue
                t2 = _vreme(polje(R2, "CreatedAt"))
                c2 = cifre(s(R2, "CorrectionID"))
                if t0 and t2:
                    if t2 >= t0:
                        return R2
                elif c0 is not None and c2 is not None:
                    if c2 > c0:
                        return R2
                elif j > i:
                    return R2
            return None

        for i, R in enumerate(sv):
            mode, status = k(R, "Mode"), k(R, "Status")
            otvoren = status in (ST_PENDING, ST_MANUAL, ST_FAILED) or (mode == MODE_KASNIJE and status != ST_COMPLETED)
            if not otvoren:
                continue
            cid = s(R, "CorrectionID")
            tip = k(R, "OldDocType")
            ime, pk, stari = self.dok_po_broju(tip, s(R, "OldBroj"), s(R, "OldDocID"))
            st_akt = aktivni(stari)
            st_sto = [r for r in stari if storno(r)]
            # naslednik: NewBroj, ili trag ispravke (IspravkaOd / CorrectionID), ili ponovni unos istog broja
            ntip = k(R, "NewDocType") or tip
            nime, npk, novi = self.dok_po_broju(ntip, s(R, "NewBroj"), s(R, "NewDocID")) \
                if (s(R, "NewBroj") or s(R, "NewDocID")) else ("", "", [])
            naslednik = aktivni(novi)
            if not naslednik and tip in nas_trag and ime:
                isp, corr = nas_trag[tip]
                for r in self.redovi(ime):
                    if storno(r):
                        continue
                    if (isp and k(R, "OldBroj") and k(r, "IspravkaOd") == k(R, "OldBroj")) or \
                            (corr and cid and k(r, "CorrectionID") == kljuc(cid)):
                        naslednik.append(r)
                if not naslednik and st_sto and st_akt:
                    naslednik = st_akt
                nime, npk = (nime or ime), (npk or pk)
            # stanje i ozbiljnost
            R2 = kasniji_zavrsen(i, R)
            if R2 is not None:
                ozb, stanje = INFO, "kasnije reseno kontekstom %s [%s/%s] nad istim dokumentom; status nije azuriran" % (
                    s(R2, "CorrectionID") or "?", s(R2, "Mode"), s(R2, "Status"))
            elif status == ST_CANCELLED:
                ozb, stanje = INFO, "parkirano pa otkazano"
            elif status == ST_MANUAL:
                ozb, stanje = SUMNJA, "potreban rucni korak: %s" % (s(R, "RecoveryAction") or "?")
            elif status == ST_FAILED:
                if st_akt and not st_sto:
                    ozb, stanje = INFO, "pokusaj nije uspeo, dokument je ostao aktivan (rollback)"
                else:
                    ozb, stanje = SUMNJA, "pokusaj nije uspeo, a dokument je storniran"
            elif mode == MODE_KASNIJE:
                if st_akt and st_sto:
                    ozb, stanje = INFO, "parkirano za kasnije; dokument je u medjuvremenu storniran pa ponovo unet"
                elif st_akt:
                    ozb, stanje = SUMNJA, "parkirano za kasnije i nikad reseno; dokument je i dalje aktivan"
                else:
                    ozb, stanje = INFO, "parkirano za kasnije; dokument je u medjuvremenu storniran"
            elif mode == MODE_ISPRAVKA:
                if naslednik:
                    ozb, stanje = INFO, "u podacima postoji naslednik %s; status nije azuriran" % _lista_id(
                        naslednik, npk or pk)
                elif st_sto:
                    ozb, stanje = SUMNJA, "stari dokument storniran, novi nije snimljen (ispravka nije zavrsena)"
                else:
                    ozb, stanje = INFO, "kontekst otvoren, stari dokument nije storniran"
            else:
                ozb, stanje = SUMNJA, "storno tok prekinut pre kraja"
            if not stari:
                stari_txt = "nije nadjen"
            else:
                stari_txt = "%s [%s]" % (_lista_id(stari, pk), _stanje(stari))
            delovi = ["%s [%s/%s] %s %s: %s. Dokument: %s." % (
                cid, s(R, "Mode"), s(R, "Status"), s(R, "OldDocType"), s(R, "OldBroj"), stanje, stari_txt)]
            if s(R, "NewBroj"):
                delovi.append("Novi broj %s." % s(R, "NewBroj"))
            pb = s(R, "ParentBroj")
            povezani = {kljuc(s(r, pk)) for r in stari if pk} | {kljuc(s(r, npk)) for r in naslednik if npk}
            if pb:
                zb = ctx.zbr_po_broju.get(kljuc(pb), [])
                prj = ctx.prj_po_broju_zbirne.get(kljuc(pb), [])
                otp = ctx.otp_po_broju_zbirne.get(kljuc(pb), [])
                delovi.append("Roditelj %s %s: %s; aktivnih otpremnica %d, aktivnih prijemnica %d." % (
                    s(R, "ParentDocType") or "Zbirna", pb,
                    ("nema" if not zb else "aktivna" if aktivni(zb) else "stornirana"),
                    len(aktivni(otp)), len(aktivni(prj))))
                povezani |= {kljuc(s(z, "ZbirnaID")) for z in zb} | {kljuc(s(p, "PrijemnicaID")) for p in prj} \
                    | {kljuc(s(o, "OtpremnicaID")) for o in otp}
            povezani.discard("")
            kodovi = sorted({x.kod for x in ctx.nalazi if x.kod != KOD and not x.kod.endswith("_PRESKOCENO")
                             and {kljuc(t) for t in x.id.split(",")} & povezani})
            if kodovi:
                delovi.append("Povezani nalazi: %s." % ", ".join(kodovi))
            if s(R, "Message"):
                delovi.append("Poruka: %s" % s(R, "Message"))
            if mode == MODE_ISPRAVKA and ozb == SUMNJA:
                predlog = ("Snimiti novi dokument i dovrsiti prevezivanje (blokovi, prijemnica, palete) kao "
                           "Complete*Ispravka, ili vratiti stari u aktivno. Procitati PRE migracije: tblStornoVeze "
                           "se ne prenosi.")
            elif ozb == INFO:
                predlog = "Proveriti stvarno stanje i zatvoriti trag; tblStornoVeze se ne prenosi u novu svesku."
            else:
                predlog = ("Odluka operatera: dovrsiti ili ponistiti korak (%s). Procitati PRE migracije: "
                           "tblStornoVeze se ne prenosi." % (s(R, "RecoveryAction") or "vidi Message"))
            ctx.nalaz(KOD, ozb, "tblStornoVeze", cid, " ".join(delovi), predlog)


def provere(ctx):
    p = _Provera(ctx)
    p.otp_sa_storniranim_blokom()
    p.storno_po_broju()
    p.ponistenje()
    p.auto_uskrsnuce()
    p.delimicno()
    p.otk_brojzbirne_storno()
    p.fak_osiroceno_zastarelo()
    p.kontekst_nedovrsen()      # poslednji: ukrsta se sa nalazima ostalih klasa
