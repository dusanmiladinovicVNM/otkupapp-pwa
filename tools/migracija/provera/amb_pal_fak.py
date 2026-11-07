"""Provere ambalaze (reversi, OMUlaz, Kupci-Otpremnica), paleta i faktura -- SAMO CITA.

Mehanizmi u starom kodu (vba-v2.2.3 .. vba-v2.28.4), po specifikaciji modula:

  AMBALAZA  frmDokumenta posle snimanja reversa ne prazni cmbVozac,
            cmbOtkupnoMesto, tip ambalaze ni txtDatum (f11:1841-1850,
            f28:1965-1972); broj reversa se ne osvezava na promenu datuma
            (f11:529-534); Kupci-Otpremnica i OMUlaz nemaju storno
            (StornoAmbalazaByDokument / StornoOMKoopByBrDok filtriraju po
            DokumentTip, a ta dva tipa nisu na storno listi).
  PALETE    prevezivanje paleta i zbirne po BROJU prijemnice, bez kupca
            (ReassignPaleteToPrijemnica_TX p11:1203-1306,
            ReassignPrijemnicaToZbirna_TX d21:3188, Detach p28:1738-1776);
            storno prijemnice ne dira paletne stavke (s11:541-593).
  FAKTURE   RelinkFakturaStavke po broju, bez kupca i storna (d11:1859-1954);
            storno prijemnice ostavlja fakturu i liniju aktivne (s11:576-584).

Pravila: tabela ne postoji ili je prazna -> provera se tiho preskace; tabela
postoji a kolona fali -> jedan INFO "<KOD>_PRESKOCENO". Jedan nalaz po
dokumentu po klasi (redovi Klase I i II jednog dokumenta se spajaju).
"""

import re
from collections import Counter, defaultdict
from datetime import date, timedelta
from decimal import Decimal

from ambalaza import norm_tip
from citac import polje
from dokumenti import GRUPE

from provera.zajednicko import (GRESKA, INFO, SUMNJA, aktivni, cifre, d, dan_txt, ids, ima_broj, k, kg_txt, n,
                                norm_broj, s, storno)

# DokumentTip (casefold) -- modConfig.bas:214, 433-440 (vba-v2.28.4)
OMULAZ = "omulaz"
KUPCI = "kupci-otpremnica"
IZLAZ_KOOP = "om-izlaz-koop"
ULAZ_KOOP = "om-ulaz-koop"
ULAZ_FIRMA = "om-ulaz-firma"
IZLAZ_FIRMA = "om-izlaz-firma"
REV_FIRMA = (ULAZ_FIRMA, IZLAZ_FIRMA, OMULAZ)
REV_TOGGLE = (IZLAZ_KOOP, ULAZ_KOOP, ULAZ_FIRMA, IZLAZ_FIRMA)   # broj se auto-predlaze (f11:1780-1783)
RODITELJ = {"otkup": "tblOtkup", IZLAZ_KOOP: "tblOtkup", "otpremnica": "tblOtpremnica",
            "prijemnica": "tblPrijemnica"}
DAN_KOOP_SMER = date(2026, 6, 26)    # vba-v2.4.0: IZDAVANJE / PRIJEM kooperanta
PROZOR_TIP = 14
TOL_KG = Decimal("0.5")
TOL_IZNOS = Decimal("0.01")
TOL_LINIJA = Decimal("0.001")
_KANON = re.compile(r"^\d+/(\d{2})(\d{2})(\d{2})(?:-\d+)?$")
_AMB_ID = re.compile(r"^AMB-\d+$", re.IGNORECASE)


class _A:
    """Red tblAmbalaza, procitan jednom."""

    __slots__ = ("r", "id", "red", "pos", "dan", "tip", "tip_txt", "kol", "smer", "smer_txt", "etip", "etip_txt",
                 "eid", "eid_txt", "voz", "voz_txt", "dok", "dok_txt", "dtip", "dtip_txt", "st")

    def __init__(self, r):
        self.r = r
        self.id = s(r, "AmbID")
        self.dan = d(r, "Datum")
        self.tip_txt = s(r, "TipAmbalaze")
        self.tip = norm_tip(self.tip_txt)
        self.kol = n(r, "Kolicina")
        self.smer_txt, self.smer = s(r, "Smer"), k(r, "Smer")
        self.etip_txt, self.etip = s(r, "EntitetTip"), k(r, "EntitetTip")
        self.eid_txt, self.eid = s(r, "EntitetID"), k(r, "EntitetID")
        self.voz_txt, self.voz = s(r, "VozacID"), k(r, "VozacID")
        self.dok_txt, self.dok = s(r, "DokumentID"), k(r, "DokumentID")
        self.dtip_txt, self.dtip = s(r, "DokumentTip"), k(r, "DokumentTip")
        self.st = storno(r)

    def txt(self):
        return "%s %s %s %s %s x %s %s" % (self.id, self.dtip_txt, self.etip_txt, self.eid_txt, self.smer_txt,
                                           kg_txt(self.kol), self.tip_txt or "?")


class _Rad:
    """Jedan prolaz modula: kontekst, PRESKOCENO evidencija i indeksi koji se dele."""

    def __init__(self, ctx):
        self.ctx = ctx
        self._preskoceno = set()
        self._amb = None
        self._amb_dok = None
        self._saldo = None
        self._kolizije = None
        self._nasl = None
        self._nov = None
        self.otkup_id = set(ctx.po_id.get("tblOtkup", {}))
        self._po_id = {}
        self.prj = self.po_id("tblPrijemnica")
        self.pal = ctx.po_id.get("tblPaleta", {})
        self.fak = ctx.po_id.get("tblFakture", {})

    def po_id(self, ime):
        """ID -> red; kad isti ID ima vise redova (rucna kopija, sync), aktivan red ima prednost pred
        storniranim: 'roditelj storniran' znaci da je storniran SVAKI red tog ID-ja."""
        if ime not in self._po_id:
            pk = self.ctx.PK.get(ime)
            idx = {}
            for x in self.ctx.tab(ime) if pk else ():
                kl = k(x, pk)
                if kl not in idx or (storno(idx[kl]) and not storno(x)):
                    idx[kl] = x
            self._po_id[ime] = idx
        return self._po_id[ime]

    # --- kolone -------------------------------------------------------
    def kolone(self, kod, ime, kolone):
        """True ako tabela ima redove i sve kolone. Prazna/nepostojeca tabela ->
        tiho False; kolona fali -> jedan INFO <KOD>_PRESKOCENO."""
        if not self.ctx.tab(ime):
            return False
        fale = [c for c in kolone if not self.ctx.ima_kolonu(ime, c)]
        if not fale:
            return True
        if kod not in self._preskoceno:
            self._preskoceno.add(kod)
            self.ctx.nalaz(kod + "_PRESKOCENO", INFO, ime, "%s.%s" % (ime, "/".join(fale)),
                           "Provera %s preskocena: tabela %s nema kolonu %s." % (kod, ime, ", ".join(fale)),
                           "Proveriti zaglavlje tabele u svesci.")
        return False

    # --- ambalaza -----------------------------------------------------
    def amb(self):
        """Redovi tblAmbalaza po redosledu upisa (AmbID = MAX+1, GetNextID)."""
        if self._amb is None:
            redovi = self.ctx.tab("tblAmbalaza")
            out = [_A(r) for r in redovi]
            po_id = bool(out) and all(_AMB_ID.match(a.id) for a in out)
            for i, a in enumerate(out):
                a.red = (cifre(a.id), a.r.get("_redni", i)) if po_id else (a.r.get("_redni", i), 0)
            out.sort(key=lambda a: a.red)
            for i, a in enumerate(out):
                a.pos = i
            self._amb = out
        return self._amb

    def amb_dok(self):
        if self._amb_dok is None:
            g = defaultdict(list)
            for a in self.amb():
                if a.dok:
                    g[(a.dtip, a.dok)].append(a)
            self._amb_dok = g
        return self._amb_dok

    def je_revers(self, a):
        if a.dtip in REV_FIRMA or a.dtip == ULAZ_KOOP:
            return True
        return a.dtip == IZLAZ_KOOP and a.dok not in self.otkup_id

    def saldo(self):
        """(EntitetTip, EntitetID, tip) -> Ulaz - Izlaz nad aktivnim redovima."""
        if self._saldo is None:
            sal = defaultdict(Decimal)
            for a in self.amb():
                if a.st:
                    continue
                if a.smer == "ulaz":
                    sal[(a.etip, a.eid, a.tip)] += a.kol
                elif a.smer == "izlaz":
                    sal[(a.etip, a.eid, a.tip)] -= a.kol
            self._saldo = sal
        return self._saldo

    def novac_om(self):
        """BrojDokumenta (kljuc) -> aktivni redovi tblNovac OMUlaz sekcije (EntitetTip OM, KooperantID neprazan,
        f11:1287-1303). Prazno ako tblNovac nema potrebne kolone."""
        if self._nov is None:
            nov = defaultdict(list)
            if self.ctx.ima_kolonu("tblNovac", "BrojDokumenta") and self.ctx.ima_kolonu("tblNovac", "KooperantID"):
                for x in aktivni(self.ctx.tab("tblNovac")):
                    if not k(x, "KooperantID") or not k(x, "BrojDokumenta"):
                        continue
                    if k(x, "EntitetTip") and k(x, "EntitetTip") != "om":
                        continue
                    nov[k(x, "BrojDokumenta")].append(x)
            self._nov = nov
        return self._nov

    def kooperanti_novca(self, a):
        """KooperantID iz novca istog broja, iste stanice (OMID) i dana kao red a."""
        return {k(x, "KooperantID") for x in self.novac_om().get(a.dok, ()) if a.dok
                and (not k(x, "OMID") or k(x, "OMID") == a.eid) and (not d(x, "Datum") or d(x, "Datum") == a.dan)}

    def susedi_pre(self, a, koliko=5):
        amb = self.amb()
        return [amb[j] for j in range(a.pos - 1, max(-1, a.pos - 1 - koliko), -1)]

    # --- prijemnice / palete -------------------------------------------
    def kolizije(self):
        """BrojPrijemnice (norm) koji se javlja kod >= 2 razlicita KupacID (aktivni i stornirani)."""
        if self._kolizije is None:
            kupci = defaultdict(set)
            for p in self.ctx.tab("tblPrijemnica"):
                b = norm_broj(polje(p, "BrojPrijemnice"))
                if b and k(p, "KupacID"):
                    kupci[b].add(k(p, "KupacID"))
            self._kolizije = {b: v for b, v in kupci.items() if len(v) >= 2}
        return self._kolizije

    def paleta_aktivna(self, st):
        lp = self.pal.get(k(st, "PaletaID"))
        return lp is None or not storno(lp)

    def stavke_aktivne(self):
        """Aktivne paletne stavke na aktivnoj (ili nepoznatoj) paleti."""
        return [x for x in self.ctx.tab("tblPaletaStavka") if not storno(x) and self.paleta_aktivna(x)]


def _dok(ime, r):
    """Logicki dokument (dokumenti.GRUPE): vlasnik + datum + poslovni broj; bez broja -> ID."""
    pk, polja, broj = GRUPE[ime]
    if not k(r, broj):
        return (ime, k(r, pk))
    return (ime,) + tuple(dan_txt(d(r, c)) if c == "Datum" else k(r, c) for c in polja)


def _prj_txt(p):
    return "%s (broj %s, kupac %s, %s)" % (s(p, "PrijemnicaID"), s(p, "BrojPrijemnice") or "-", s(p, "KupacID") or "-",
                                          dan_txt(d(p, "Datum")) or "-")


def _jedinstveni(redovi, pk):
    out, vid = [], set()
    for r in redovi:
        if k(r, pk) not in vid:
            vid.add(k(r, pk))
            out.append(r)
    return out


def _razlicito(a, b):
    """Razlika dva teksta; prazno je nepoznato, nije razlika."""
    return bool(a) and bool(b) and a != b


def _fak_aktivna(f):
    return not storno(f) and k(f, "Status") != "stornirano"


def provere(ctx):
    r = _Rad(ctx)
    _amb_noga_storniranog(r)
    _amb_revers_vozac(r)
    _amb_kupci(r)
    _amb_revers_tip(r)
    _amb_omulaz_smer(r)
    _amb_bez_storna(r)
    _amb_kompenzacija(r)
    _amb_broj_reversa(r)
    _amb_datum_iz_ispravke(r)
    pokriveno = _pal_prevezivanje(r)
    _pal_zbirna_po_broju(r)
    _pal_broj_zbirne(r, pokriveno)
    _pal_gajbe_kg(r, pokriveno)
    _pal_identitet(r)
    _pal_mesani_kupci(r)
    _pal_na_storniranoj(r)
    _fak_sef(r)
    _fak_linije(r)
    _fak_status(r)


# =====================================================================
# AMBALAZA
# =====================================================================

AMB_OSNOVA = ["AmbID", "Datum", "TipAmbalaze", "Kolicina", "Smer", "EntitetID", "EntitetTip", "DokumentID",
              "DokumentTip"]


def _amb_noga_storniranog(r):
    """Aktivna noga (Otkup, OM-Izlaz-Koop uz otkup, Otpremnica, Prijemnica) ciji je roditelj storniran.
    Svaki storno dokumenta zove StornoAmbalazaByDokument (vba-v2.2.3:modStorno.bas:141, 195, 434)."""
    K = "AMB_AKTIVNA_NOGA_STORNIRANOG_DOKUMENTA"
    if not r.kolone(K, "tblAmbalaza", ["AmbID", "DokumentID", "DokumentTip", "Stornirano"]):
        return
    ctx = r.ctx
    grupe = defaultdict(list)
    for a in r.amb():
        ime = RODITELJ.get(a.dtip)
        if a.st or not ime or not a.dok:
            continue
        p = r.po_id(ime).get(a.dok)
        if p is None or not storno(p):
            continue
        grupe[_dok(ime, p)].append((a, p))
    for kl, lst in grupe.items():
        ime = kl[0]
        pk, _, bkol = GRUPE[ime]
        noge = [a for a, _ in lst]
        rod = _jedinstveni([p for _, p in lst], pk)
        ctx.nalaz(K, GRESKA, "tblAmbalaza", ", ".join(a.id for a in noge),
                  "Aktivne noge uz storniran dokument %s (broj %s, %s): %s." % (
                      ids(rod, pk), s(rod[0], bkol) or "-", dan_txt(d(rod[0], "Datum")),
                      "; ".join(a.txt() for a in noge)),
                  "Izuzeti noge %s (Stornirano = Da): dokument je storniran, njegova ambalaza ne sme ostati "
                  "u knjizi." % ", ".join(a.id for a in noge))


def _amb_revers_vozac(r):
    """Revers firma<->OM (OM-Ulaz-Firma, OM-Izlaz-Firma, OMUlaz) sa vozacem koji nije par-vozac stanice.
    cmbVozac i cmbOtkupnoMesto se ne prazne posle snimanja (f11:1841-1850, f28:1965-1972), pa je JEDNO od
    dva polja zastarelo; koje -- odlucuju otpremnice istog dana (K1 vozac / K2 stanica)."""
    K = "AMB_REVERS_VOZAC_NIJE_PAR_STANICE"
    K2 = "AMB_REVERS_STANICA_ZASTARELA"
    if not r.kolone(K, "tblAmbalaza", AMB_OSNOVA + ["VozacID"]):
        return
    ctx = r.ctx
    par = {k(v, "VozacID") for v in ctx.tab("tblVozaci")}
    if not par:
        # bez tblVozaci se ne zna koja stanica ima par-vozaca (stanica bez njega = rucni vozac): tiho preskoci
        return
    otp_dan = defaultdict(set)
    for o in aktivni(ctx.tab("tblOtpremnica")):
        if d(o, "Datum"):
            otp_dan[d(o, "Datum")].add(k(o, "StanicaID"))
    # stanica bez par-vozaca: auto-izbor ne radi, vozac je uvek rucni (lazni pozitiv)
    kand = [a for a in r.amb() if not a.st and a.etip == "stanica" and a.dtip in REV_FIRMA and a.voz
            and a.eid and a.voz != a.eid and a.eid in par]
    serija = Counter((a.dan, a.voz) for a in kand)
    vidjeno = set()
    for a in kand:
        dk = (a.dtip, a.dok) if a.dok else ("#", a.id)
        if dk in vidjeno:
            continue
        vidjeno.add(dk)
        stanice = otp_dan.get(a.dan, set())
        ima_e, ima_v = a.eid in stanice, a.voz in stanice
        k1 = "VozacID: %s -> %s" % (a.voz_txt, a.eid_txt)
        k2 = "EntitetID: %s -> %s" % (a.eid_txt, a.voz_txt)
        otp_txt = ", ".join(sorted(x.upper() for x in stanice)) or "nema"
        if not ctx.je_stanica(a.voz):
            kod, oz = K, INFO
            predlog = "%s ako vozac nije izabran namerno (%s nije par-vozac nijedne stanice)." % (k1, a.voz_txt)
        elif ima_e and not ima_v:
            kod, oz = K, SUMNJA
            predlog = "%s (vozac zastareo: tog dana otpremnicu ima %s, a %s nema)." % (k1, a.eid_txt, a.voz_txt)
        elif ima_v and not ima_e:
            kod, oz = K2, SUMNJA
            predlog = "%s (stanica zastarela: tog dana otpremnicu ima %s, a %s nema)." % (k2, a.voz_txt, a.eid_txt)
        else:
            kod, oz = K, SUMNJA
            predlog = "ZA ODLUKU: (1) %s ili (2) %s; otpremnice tog dana: %s." % (k1, k2, otp_txt)
        opis = "%s %s, %s: stanica %s, vozac %s (par-vozac stanice je %s)." % (
            a.dtip_txt, a.dok_txt or "(bez broja)", dan_txt(a.dan), a.eid_txt, a.voz_txt, a.eid_txt)
        if serija[(a.dan, a.voz)] > 1:
            opis += " Istog dana %d reversa sa istim tudjim vozacem." % serija[(a.dan, a.voz)]
        pre = [b for b in r.susedi_pre(a) if a.voz in (b.voz, b.eid)]
        if pre:
            opis += " Prethodni unos sa tim vozacem: %s %s (%s)." % (pre[0].id, pre[0].dtip_txt, dan_txt(pre[0].dan))
        ctx.nalaz(kod, oz, "tblAmbalaza", a.id, opis, predlog)


def _amb_kupci(r):
    """Kupci-Otpremnica (Kupac Izlaz + VozacID, SaveKupciIzlaz_TX d11:1225-1260): cmbKupac i cmbVozac se ne
    prazne posle unosa (f11:2808-2811). INFO: kupac bez prijemnice u Datum-30..Datum+14 (kupac pozajmljuje
    prazne i pred sezonu, pa prijemnica POSLE unosa iskljucuje nalaz); vozac bez prijemnice kod kupca u 7
    dana, a kupac je u tih 7 dana imao prijemnice sa drugim vozacem (inace nema poredjenja: prazne se
    vracaju danima kasnije)."""
    KK = "KUPAC_ZASTAREO_NA_KUPCI_OTPREMNICI_I_UPLATI"
    KV = "AMB_REVERS_VOZAC_NIJE_PAR_STANICE"
    if not (r.kolone(KK, "tblAmbalaza", AMB_OSNOVA + ["VozacID"])
            and r.kolone(KK, "tblPrijemnica", ["KupacID", "Datum"])):
        return
    ctx = r.ctx
    ima_voz = r.kolone(KV, "tblPrijemnica", ["VozacID"])
    po_kupcu = defaultdict(list)
    for p in aktivni(ctx.tab("tblPrijemnica")):
        if d(p, "Datum") and k(p, "KupacID"):
            po_kupcu[k(p, "KupacID")].append((d(p, "Datum"), k(p, "VozacID"), s(p, "VozacID")))
    vidjeno = set()
    bez_ikakve = defaultdict(list)     # kupac bez ijedne aktivne prijemnice: jedan nalaz po kupcu
    for a in r.amb():
        if a.st or a.dtip != KUPCI or a.etip != "kupac" or not a.dan or not a.eid:
            continue
        dk = a.dok or "#" + a.id
        if dk in vidjeno:
            continue
        vidjeno.add(dk)
        lst = po_kupcu.get(a.eid, [])
        if not lst:
            bez_ikakve[a.eid].append(a)
            continue
        u_prozoru = [x for x in lst if a.dan - timedelta(days=30) <= x[0] <= a.dan + timedelta(days=14)]
        if not u_prozoru:
            pre = [x[0] for x in lst if x[0] <= a.dan]
            ctx.nalaz(KK, INFO, "tblAmbalaza", a.id,
                      "Kupci-Otpremnica %s, %s: kupac %s nema aktivnu prijemnicu od 30 dana pre do 14 dana posle "
                      "(poslednja pre: %s); %s gajbi %s, vozac %s." % (
                          a.dok_txt or "(bez broja)", dan_txt(a.dan), a.eid_txt,
                          dan_txt(max(pre)) if pre else "nijedna", kg_txt(a.kol), a.tip_txt or "?",
                          a.voz_txt or "-"),
                      "Proveriti kupca (cmbKupac ostaje od prethodnog unosa); kandidat je kupac kod koga je vozac %s "
                      "imao prijemnicu tog ili prethodnih dana." % (a.voz_txt or "-"))
            continue
        if not ima_voz or not a.voz:
            continue
        u7 = [x for x in lst if a.dan - timedelta(days=7) <= x[0] <= a.dan]
        if not u7 or any(x[1] == a.voz for x in u7):
            continue
        kandidati = []
        for x in sorted(u7, reverse=True):
            if x[2] and x[2] not in kandidati:
                kandidati.append(x[2])
        ctx.nalaz(KV, INFO, "tblAmbalaza", a.id,
                  "Kupci-Otpremnica %s, %s: kupac %s, vozac %s nema prijemnicu kod tog kupca u 7 dana pre toga, "
                  "a kupac ih u tom roku ima sa vozacima %s." % (
                      a.dok_txt or "(bez broja)", dan_txt(a.dan), a.eid_txt, a.voz_txt, ", ".join(kandidati) or "-"),
                  "Proveriti vozaca (cmbVozac ostaje od prethodnog unosa); kandidati (prijemnica kod kupca u 7 dana): "
                  "%s." % (", ".join(kandidati) or "nema"))
    for lst in bez_ikakve.values():
        a = lst[0]
        ctx.nalaz(KK, INFO, "tblAmbalaza", ", ".join(x.id for x in lst),
                  "Kupac %s nema nijednu aktivnu prijemnicu u svesci, a ima %d Kupci-Otpremnica (%s, ukupno %s gajbi)."
                  % (a.eid_txt, len(lst), ", ".join("%s %s" % (x.dok_txt or "(bez broja)", dan_txt(x.dan)) for x in lst[:5])
                     + (" ..." if len(lst) > 5 else ""), kg_txt(sum(x.kol for x in lst))),
                  "Ako kupac nije partner za ambalazu: proveriti kupca (cmbKupac ostaje od prethodnog unosa); kandidat "
                  "je kupac kod koga je vozac imao prijemnicu tog ili prethodnih dana.")


def _revers_dokumenti(r):
    """[(familija, partner noga, sve noge)] za reverse i Kupci-Otpremnica, po redosledu upisa.
    Familija = combo tipa ambalaze: OMUlaz sekcija (cmbTipAmbOMUlaz) ili Izlaz kupci (cmbTipAmbIzlaz)."""
    grupe = defaultdict(list)
    for a in r.amb():
        if a.st:
            continue
        if r.je_revers(a) or a.dtip == KUPCI:
            grupe[(a.dtip, a.dok) if a.dok else ("#", a.id)].append(a)
    out = []
    for noge in grupe.values():
        dt = noge[0].dtip
        if dt == KUPCI:
            fam, et = "kupci", "kupac"
        elif dt in (IZLAZ_KOOP, ULAZ_KOOP):
            fam, et = "om", "kooperant"
        else:
            fam, et = "om", "stanica"
        partner = next((a for a in noge if a.etip == et), None)
        if partner is not None:
            out.append((fam, partner, noge))
    out.sort(key=lambda x: min(a.pos for a in x[2]))
    return out


def _amb_revers_tip(r):
    """Tip ambalaze na reversu / Kupci-Otpremnici ostao od prethodnog unosa (cmbTipAmbOMUlaz i
    cmbTipAmbIzlaz se pune samo pri prvom setup-u, f11:129-130, i ne prazne se)."""
    K = "AMB_REVERS_TIP_ZASTAREO"
    if not r.kolone(K, "tblAmbalaza", AMB_OSNOVA):
        return
    ctx = r.ctx
    if len({a.tip for a in r.amb() if not a.st and a.tip}) <= 1:
        return
    robni = defaultdict(list)
    for ime, kol, et in (("tblOtkup", "KooperantID", "kooperant"), ("tblOtkup", "StanicaID", "stanica"),
                         ("tblOtpremnica", "StanicaID", "stanica"), ("tblPrijemnica", "KupacID", "kupac")):
        if not ctx.tab(ime):
            continue
        if not r.kolone(K, ime, [kol, "TipAmbalaze", "Datum"]):
            return
        for x in aktivni(ctx.tab(ime)):
            tp = norm_tip(polje(x, "TipAmbalaze"))
            if d(x, "Datum") and tp and k(x, kol):
                robni[(et, k(x, kol))].append((d(x, "Datum"), tp))

    def tipovi(a):
        c = Counter()
        for dt, tp in robni.get((a.etip, a.eid), ()):
            if abs((dt - a.dan).days) <= PROZOR_TIP:
                c[tp] += 1
        return c

    saldo = r.saldo()
    po_fam = defaultdict(list)
    for fam, partner, noge in _revers_dokumenti(r):
        po_fam[fam].append((partner, noge))
    for fam, lst in po_fam.items():
        for i, (a, noge) in enumerate(lst):
            if not a.tip or not a.dan:
                continue
            tt = tipovi(a)
            if not tt or a.tip in tt:
                continue
            nasl = None
            for j in range(i - 1, max(-1, i - 11), -1):
                b = lst[j][0]
                if b.eid == a.eid:
                    continue
                if b.tip == a.tip and b.dan and b.tip in tipovi(b):
                    nasl = b
                break
            swap = None
            sb = saldo[(a.etip, a.eid, a.tip)]
            for ta, _ in tt.most_common():
                sa = saldo[(a.etip, a.eid, ta)]
                if sa and sb and (sa > 0) != (sb > 0) and a.kol <= abs(sa):
                    swap = (ta, sa, sb)
                    break
            dom = swap[0] if swap else tt.most_common(1)[0][0]
            opis = "%s %s, %s: %s %s, tip %s, a robni dokumenti tog naloga u +-%d dana nose samo %s." % (
                a.dtip_txt, a.dok_txt or "(bez broja)", dan_txt(a.dan), a.etip_txt, a.eid_txt, a.tip_txt,
                PROZOR_TIP, ", ".join("%s (%d)" % x for x in tt.most_common()))
            if nasl is not None:
                opis += " Prethodni revers %s (%s %s) ima isti tip %s i njemu odgovara: tip nasledjen." % (
                    nasl.id, nasl.etip_txt, nasl.eid_txt, nasl.tip_txt)
            if swap is not None:
                opis += " Saldo naloga: %s = %s, %s = %s (suprotan znak)." % (
                    swap[0], kg_txt(swap[1]), a.tip, kg_txt(swap[2]))
            ctx.nalaz(K, SUMNJA if (nasl is not None or swap is not None) else INFO, "tblAmbalaza",
                      ", ".join(x.id for x in noge), opis,
                      "TipAmbalaze: %s -> %s (dominantan tip naloga u prozoru); prihvatiti samo ako zamena zatvara "
                      "saldo." % (a.tip_txt, dom))


def _amb_omulaz_smer(r):
    """OMUlaz (Stanica Ulaz) kad je smer-toggle zaboravljen. Samo uz saldo-dokaz: kooperant iz novca istog
    broja drzi tacno +-Kolicina gajbi tog tipa (PRIJEM / IZDAVANJE bi ga sveo na nulu)."""
    K = "AMB_OMULAZ_ZABORAVLJEN_SMER"
    if not r.kolone(K, "tblAmbalaza", AMB_OSNOVA):
        return
    ctx = r.ctx
    if not r.kolone(K, "tblNovac", ["BrojDokumenta", "KooperantID"]):
        return
    nov = r.novac_om()     # novac OMUlaz sekcije: EntitetTip = "OM" (f11:1287-1303)
    saldo = r.saldo()
    for a in r.amb():
        if (a.st or a.dtip != OMULAZ or a.etip != "stanica" or not a.dok or not a.dan or a.dan < DAN_KOOP_SMER
                or a.kol <= 0):
            continue
        for x in nov.get(a.dok, ()):
            if k(x, "OMID") and k(x, "OMID") != a.eid:
                continue
            if d(x, "Datum") and d(x, "Datum") != a.dan:
                continue
            koop = k(x, "KooperantID")
            sk = saldo[("kooperant", koop, a.tip)]
            ss = saldo[("stanica", a.eid, a.tip)]
            if sk == a.kol:
                kand = ("OM-Ulaz-Koop (PRIJEM: kooperant Izlaz + stanica Ulaz, stanica ostaje ista)",
                        "kooperant nije razduzen")
            elif sk == -a.kol:
                kand = ("OM-Izlaz-Koop (IZDAVANJE: kooperant Ulaz + stanica Izlaz; stanica %s -> %s)" % (
                    kg_txt(ss), kg_txt(ss - 2 * a.kol)), "kooperant nije zaduzen")
            else:
                continue
            ctx.nalaz(K, SUMNJA, "tblAmbalaza", a.id,
                      "OMUlaz %s, %s: stanica %s Ulaz %s x %s, uz novac %s kooperantu %s; saldo tog kooperanta za "
                      "tip %s je %s (%s)." % (a.dok_txt, dan_txt(a.dan), a.eid_txt, kg_txt(a.kol), a.tip_txt,
                                              s(x, "NovacID"), s(x, "KooperantID"), a.tip, kg_txt(sk), kand[1]),
                      "Ako je smer zaboravljen: %s, ista kolicina/tip/datum/broj. Ako je OMUlaz stvaran (firma salje "
                      "prazne i novac za kooperanta preko vozaca), OMUlaz ostaje, a nedostaje poseban revers "
                      "kooperanta." % kand[0])
            break


def _amb_bez_storna(r):
    """Kupci-Otpremnica / OMUlaz nemaju storno: dupli unos ostaje aktivan, a Stornirano = Da na njima je
    rucna izmena (nijedan pisac to ne radi)."""
    K = "AMB_BEZ_STORNA_KUPCI_OTPREMNICA_OMULAZ"
    if not r.kolone(K, "tblAmbalaza", AMB_OSNOVA + ["VozacID", "Stornirano"]):
        return
    ctx = r.ctx
    grupe = defaultdict(list)
    storn = defaultdict(list)
    for a in r.amb():
        if a.dtip not in (KUPCI, OMULAZ):
            continue
        if a.st:
            storn[(a.dtip, a.dok) if a.dok else ("#", a.id)].append(a)
        elif a.dan and a.kol > 0:
            grupe[(a.dtip, a.eid, a.tip, a.kol, a.smer, a.dan, a.voz)].append(a)
    for lst in grupe.values():
        if len(lst) < 2:
            continue
        lst.sort(key=lambda a: a.pos)
        uz = any(lst[i + 1].pos - lst[i].pos == 1 for i in range(len(lst) - 1))
        prazan = any(not a.dok for a in lst)
        if not (uz or prazan) and len({a.dok for a in lst}) == len(lst):
            continue    # razliciti brojevi i drugi unosi izmedju: dve stvarne ture istog vozaca (legitimno)
        if lst[0].dtip == OMULAZ and not prazan and len({a.dok for a in lst}) == len(lst):
            koop = [r.kooperanti_novca(a) for a in lst]
            if all(koop) and len(set().union(*koop)) >= len(lst) and all(
                    not (koop[i] & koop[j]) for i in range(len(koop)) for j in range(i + 1, len(koop))):
                continue    # svaki broj nosi novac drugom kooperantu: odvojene isplate (izvorna namena sekcije)
        a = lst[0]
        ctx.nalaz(K, SUMNJA, "tblAmbalaza", ", ".join(x.id for x in lst),
                  "%d aktivna reda %s istog naloga, tipa, kolicine, smera, dana i vozaca: %s %s, %s x %s, %s, "
                  "vozac %s; brojevi %s%s." % (len(lst), a.dtip_txt, a.etip_txt, a.eid_txt, kg_txt(a.kol), a.tip_txt,
                                     dan_txt(a.dan), a.voz_txt or "-",
                                     ", ".join(x.dok_txt or "(prazan)" for x in lst),
                                     "; redovi su uzastopni" if uz else ""),
                  "Ako je dupli unos: izuzeti kasniji red %s (izuzeci.json, radnja 'izuzmi'); odluka operatera." %
                  ", ".join(x.id for x in lst[1:]))
    for noge in storn.values():
        a = noge[0]
        ctx.nalaz(K, INFO, "tblAmbalaza", ", ".join(x.id for x in noge),
                  "%s %s, %s ima Stornirano = Da, a stari kod taj tip ne stornira: rucna izmena sveske." % (
                      a.dtip_txt, a.dok_txt or "(bez broja)", dan_txt(a.dan)),
                  "Potvrditi da je storno nameran; u migraciji red ide u arhivu.")


def _amb_kompenzacija(r):
    """Pogresan Kupci-Otpremnica / OMUlaz 'storniran' kontra-unosom: isti nalog, tip i kolicina, suprotan
    smer, u 0-3 dana (isti tip dokumenta, ili OMUlaz pa OM-Izlaz-Firma iste stanice)."""
    K = "AMB_KOMPENZACIJA_BEZ_STORNA"
    if not r.kolone(K, "tblAmbalaza", AMB_OSNOVA):
        return
    ctx = r.ctx
    akt = [a for a in r.amb() if not a.st and a.dan and a.kol > 0 and a.smer in ("ulaz", "izlaz")]
    idx = defaultdict(list)
    for a in akt:
        if a.dtip in (KUPCI, OMULAZ) or (a.dtip == IZLAZ_FIRMA and a.etip == "stanica"):
            idx[(a.eid, a.tip, a.kol)].append(a)
    used = set()
    for a in akt:
        if a.dtip not in (KUPCI, OMULAZ) or a.pos in used:
            continue
        best = None
        for b in idx[(a.eid, a.tip, a.kol)]:
            if b is a or b.pos in used or b.smer == a.smer:
                continue
            delta = (b.dan - a.dan).days
            if b.dtip == a.dtip:
                if abs(delta) > 3:
                    continue
            elif a.dtip == OMULAZ and b.dtip == IZLAZ_FIRMA:
                if not (0 <= delta <= 3 and b.pos > a.pos):
                    continue
            else:
                continue
            ocena = (abs(delta), abs(b.pos - a.pos))
            if best is None or ocena < best[0]:
                best = (ocena, b)
        if best is None:
            continue
        b = best[1]
        used.update((a.pos, b.pos))
        isti_broj = bool(a.dok) and a.dok == b.dok
        ctx.nalaz(K, SUMNJA if isti_broj else INFO, "tblAmbalaza", "%s, %s" % (a.id, b.id),
                  "Par suprotnog smera iste kolicine: %s (%s, %s) i %s (%s, %s)%s." % (
                      a.txt(), a.dok_txt or "bez broja", dan_txt(a.dan), b.txt(), b.dok_txt or "bez broja",
                      dan_txt(b.dan), "; isti broj dokumenta" if isti_broj else ""),
                  "Spisak za odluku: ako je drugi red ponistavanje pogresnog prvog, oba izuzeti zajedno "
                  "(izuzeci.json, radnja 'izuzmi').")


def _amb_broj_reversa(r):
    """Broj reversa ne prati Datum: txtDatum_AfterUpdate ne osvezava predlog broja reversa (f11:529-534,
    f28:636-641). Zastareo je BROJ, ne datum."""
    K = "AMB_BROJ_REVERSA_NE_PRATI_DATUM"
    if not r.kolone(K, "tblAmbalaza", ["AmbID", "Datum", "DokumentID", "DokumentTip"]):
        return
    ctx = r.ctx
    dokovi = r.amb_dok()
    vidjeno = set()
    for a in r.amb():
        if a.st or a.dtip not in REV_TOGGLE or not a.dok or not a.dan or not r.je_revers(a):
            continue
        dk = (a.dtip, a.dok)
        if dk in vidjeno:
            continue
        vidjeno.add(dk)
        m = _KANON.match(a.dok_txt)
        if not m:
            continue
        try:
            bd = date(2000 + int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            continue
        if bd == a.dan:
            continue
        noge = [x for x in dokovi.get(dk, [a]) if not x.st]
        ctx.nalaz(K, INFO, "tblAmbalaza", ", ".join(x.id for x in noge),
                  "%s %s: broj nosi datum %s, a Datum reda je %s (stanica %s)." % (
                      a.dtip_txt, a.dok_txt, bd.isoformat(), dan_txt(a.dan),
                      next((x.eid_txt for x in noge if x.etip == "stanica"), "-")),
                  "Ako je broj auto-predlog: zastareo je broj (datum %s), ne Datum; menjati samo uz potvrdu "
                  "operatera." % a.dan.strftime("%d%m%y"))


def _ispravke(r):
    """Ispravke D2 (aktivna prijemnica/otpremnica) sa storniranim D1 istog datuma. Izvori: tblStornoVeze
    (OldDocID -> NewDocID), IspravkaOd (broj starog, od v2.28.0) i heuristika: storniran raniji red istog
    kupca i zbirne ili vozaca (prijemnica), odnosno iste stanice (otpremnica), istog dana."""
    ctx = r.ctx
    out = {}
    for ime, dt in (("tblPrijemnica", "prijemnica"), ("tblOtpremnica", "otpremnica")):
        redovi = ctx.tab(ime)
        if not redovi:
            continue
        pk, _, bkol = GRUPE[ime]
        if ime == "tblPrijemnica":
            # vozac + gajbe: isti dan i vozac sami nisu dovoljni (vise tura istog dana)
            kljucevi = [lambda x: ("z", k(x, "KupacID"), norm_broj(polje(x, "BrojZbirne"))),
                        lambda x: ("v", k(x, "KupacID"), k(x, "VozacID"), kg_txt(n(x, "KolAmbalaze")))]
        else:
            kljucevi = [lambda x: ("s", k(x, "StanicaID"))]
        stor = defaultdict(list)
        for x in redovi:
            if storno(x) and d(x, "Datum"):
                for f in kljucevi:
                    if all(f(x)[1:]):
                        stor[(f(x), d(x, "Datum"))].append(x)
        stor_broj = defaultdict(list)
        for x in redovi:
            if storno(x) and k(x, bkol):
                stor_broj[norm_broj(polje(x, bkol))].append(x)
        for x in aktivni(redovi):
            dx = d(x, "Datum")
            if not dx:
                continue
            nadjen = None
            for f in kljucevi:
                if not all(f(x)[1:]):
                    continue
                for y in stor.get((f(x), dx), ()):
                    if y.get("_redni", 0) < x.get("_redni", 0):
                        nadjen = y
                        break
                if nadjen is not None:
                    break
            if nadjen is None and k(x, "IspravkaOd"):
                nadjen = next((y for y in stor_broj.get(norm_broj(polje(x, "IspravkaOd")), ())
                               if d(y, "Datum") == dx), None)
            if nadjen is not None:
                out[(dt, k(x, pk))] = (ime, x, nadjen)
    for v in ctx.tab("tblStornoVeze"):
        for ime, dt in (("tblPrijemnica", "prijemnica"), ("tblOtpremnica", "otpremnica")):
            x = ctx.po_id.get(ime, {}).get(k(v, "NewDocID"))
            y = ctx.po_id.get(ime, {}).get(k(v, "OldDocID"))
            if x is not None and y is not None and not storno(x) and d(x, "Datum") and d(x, "Datum") == d(y, "Datum"):
                out.setdefault((dt, k(x, GRUPE[ime][0])), (ime, x, y))
    return out


def _amb_datum_iz_ispravke(r):
    """Datum nasledjen iz ispravke: Prefill ispravke prepisuje txtDatum datumom starog dokumenta
    (f21:2712; f28:2846, 2916, 2988), pa ga nasledjuju i sledeci unosi u istoj sesiji forme."""
    KR = "AMB_REVERS_DATUM_ZASTAREO"
    KN = "AMB_DATUM_NASLEDJEN_IZ_ISPRAVKE"
    if not r.kolone(KR, "tblAmbalaza", ["AmbID", "Datum", "DokumentID", "DokumentTip"]):
        return
    ctx = r.ctx
    isp = _ispravke(r)
    if not isp:
        return
    amb = r.amb()
    dokovi = r.amb_dok()
    vidjeno = set()
    for (dt, pid), (ime, x, y) in sorted(isp.items()):
        dan2 = d(x, "Datum")
        noge = [a for a in dokovi.get((dt, pid), ()) if not a.st]
        if not noge or not dan2:
            continue
        p0, p1 = min(a.pos for a in noge), max(a.pos for a in noge)
        pre = None
        for j in range(p0 - 1, max(-1, p0 - 6), -1):
            if amb[j].dan and amb[j].dan != dan2:
                pre = amb[j]
                break
        if pre is None or pre.dan <= dan2:
            continue        # ispravka nije usred kasnijeg unosa: nema dokaza zastarelog datuma
        kand, posle = [], None
        for j in range(p1 + 1, min(len(amb), p1 + 6)):
            b = amb[j]
            if not b.dan:
                continue
            if b.dan == dan2:
                if not b.st and (r.je_revers(b) or b.dtip == KUPCI):
                    kand.append(b)
                continue
            posle = b
            break
        if posle is None or posle.dan <= dan2:
            continue
        for b in kand:
            dk = (b.dtip, b.dok) if b.dok else ("#", b.id)
            if dk in vidjeno:
                continue
            vidjeno.add(dk)
            sve = [z for z in dokovi.get(dk, [b]) if not z.st] if b.dok else [b]
            ctx.nalaz(KR if r.je_revers(b) else KN, SUMNJA, "tblAmbalaza", ", ".join(z.id for z in sve),
                      "%s %s ima Datum %s, isti kao ispravka %s (zamena za storniran %s) upisana neposredno pre; "
                      "unos pre ispravke (%s) i posle (%s) nose %s / %s." % (
                          b.dtip_txt, b.dok_txt or "(bez broja)", dan_txt(dan2), s(x, GRUPE[ime][0]),
                          s(y, GRUPE[ime][0]), pre.id, posle.id, dan_txt(pre.dan), dan_txt(posle.dan)),
                      "Datum: %s -> %s (datum susednih unosa), ako papir nije stvarno od %s." % (
                          dan_txt(dan2), dan_txt(posle.dan), dan_txt(dan2)))


# =====================================================================
# PALETE
# =====================================================================

def _zbirne_po_broju(ctx):
    g = defaultdict(list)
    for z in aktivni(ctx.tab("tblZbirna")):
        if k(z, "BrojZbirne"):
            g[norm_broj(polje(z, "BrojZbirne"))].append(z)
    return g


def _istorija_detach(lp, broj):
    """tblPaleta.Istorija (od v2.21.0): 'act=DETACH;det=prij=<broj> gajb=..' (PaletaLog p28:1542-1563)."""
    if lp is None or not broj:
        return False
    pat = re.compile(r"act=detach;det=prij=\s*" + re.escape(broj.casefold()) + r"(?=\s|,|;|$)")
    return any(pat.search(e.strip().casefold()) for e in s(lp, "Istorija").split("|"))


def _pal_prevezivanje(r):
    """Prevezivanje paleta po BROJU prijemnice, bez kupca (ReassignPaleteToPrijemnica_TX p11:1203-1306,
    Detach p28:1738-1776). Samo za broj koji se javlja kod >= 2 kupca. Vraca PrijemnicaID (kljuc) koje su
    ovde prijavljene, da ih generalne provere zbirne/gajbi ne ponove."""
    K = "PAL_PREVEZIVANJE_PO_BROJU_TUDJI_KUPAC"
    KD = "PAL_DETACH_PO_BROJU_TUDJA_PRIJEMNICA"
    pokriveno = set()
    if not (r.kolone(K, "tblPaletaStavka", ["StavkaID", "PaletaID", "PrijemnicaID", "BrojPrijemnice", "BrojZbirne",
                                             "BrojGajbica"])
            and r.kolone(K, "tblPrijemnica", ["PrijemnicaID", "BrojPrijemnice", "KupacID", "KolAmbalaze",
                                              "BrojZbirne"])):
        return pokriveno
    ctx = r.ctx
    kol = r.kolizije()
    if not kol:
        return pokriveno
    po_broju = defaultdict(list)
    for p in ctx.tab("tblPrijemnica"):
        po_broju[norm_broj(polje(p, "BrojPrijemnice"))].append(p)
    akt_po_p = defaultdict(list)
    for st in r.stavke_aktivne():
        akt_po_p[k(st, "PrijemnicaID")].append(st)
    # (a') aktivna stavka broja iz kolizije ne odgovara svojoj prijemnici
    docs = defaultdict(list)
    for pid, stavke in akt_po_p.items():
        p = r.prj.get(pid)
        if p is None or storno(p):
            continue
        nb = norm_broj(polje(p, "BrojPrijemnice"))
        if nb not in kol and not any(norm_broj(polje(x, "BrojPrijemnice")) in kol for x in stavke):
            continue
        razlozi, teska = [], False
        bzp = norm_broj(polje(p, "BrojZbirne"))
        losi = [x for x in stavke if norm_broj(polje(x, "BrojZbirne")) != bzp]
        if losi:
            # obe kopije popunjene i razlicite: dokazano (kao PAL_BROJ_ZBIRNE_ZASTAREO)
            teska = teska or (bool(bzp) and any(k(x, "BrojZbirne") for x in losi))
            razlozi.append("BrojZbirne stavki %s = %s, prijemnice = %s" % (
                ids(losi, "StavkaID"), "/".join(sorted({s(x, "BrojZbirne") or "(prazno)" for x in losi})),
                s(p, "BrojZbirne") or "(prazno)"))
        for f in ("VrstaVoca", "SortaVoca"):
            raz = [x for x in stavke if _razlicito(k(x, f), k(p, f))]
            if raz:
                razlozi.append("%s stavki %s = %s, prijemnice = %s" % (f, ids(raz, "StavkaID"), s(raz[0], f), s(p, f)))
        g = sum(n(x, "BrojGajbica") for x in stavke)
        if ima_broj(p, "KolAmbalaze") and g != n(p, "KolAmbalaze"):
            razlozi.append("gajbi na paletama %s, na prijemnici %s" % (kg_txt(g), kg_txt(n(p, "KolAmbalaze"))))
            teska = True    # zbir gajbi po PrijemnicaID je po konstrukciji = KolAmbalaze (kao PAL_GAJBE_KG_NESAGLASNI)
        if not razlozi:
            continue
        bz_losi = {norm_broj(polje(x, "BrojZbirne")) for x in losi}
        alt = [q for q in po_broju.get(nb, ()) if k(q, "KupacID") != k(p, "KupacID")
               and norm_broj(polje(q, "BrojZbirne")) in bz_losi]
        pokriveno.add(pid)
        docs[_dok("tblPrijemnica", p)].append((p, stavke, razlozi, alt, teska))
    for lst in docs.values():
        p0 = lst[0][0]
        alt = _jedinstveni([q for _, _, _, a, _ in lst for q in a], "PrijemnicaID")
        stv = [x for _, st_, _, _, _ in lst for x in st_]
        kupci = sorted(x.upper() for x in kol.get(norm_broj(polje(p0, "BrojPrijemnice")), ()))
        opis = "Prijemnica %s, broj se javlja kod kupaca %s: %s." % (
            _prj_txt(p0), ", ".join(kupci), "; ".join(z for _, _, rz, _, _ in lst for z in rz))
        if alt:
            predlog = ("Stavke %s nose zbirnu prijemnice %s istog broja drugog kupca: PrijemnicaID stavki -> %s "
                       "(ako su gajbe njene), ili BrojZbirne stavki -> %s (ako je PrijemnicaID tacan)." % (
                           ids(stv, "StavkaID"), _prj_txt(alt[0]), s(alt[0], "PrijemnicaID"),
                           s(p0, "BrojZbirne") or "-"))
        else:
            predlog = ("Proveriti stavke %s: prevezivanje po broju %s je moglo da zahvati prijemnicu drugog kupca; "
                       "uskladiti sa prijemnicom %s." % (ids(stv, "StavkaID"), s(p0, "BrojPrijemnice"),
                                                         s(p0, "PrijemnicaID")))
        teska = bool(alt) or any(x[4] for x in lst)
        ctx.nalaz(K, GRESKA if teska else SUMNJA, "tblPaletaStavka", ids(stv, "StavkaID"), opis, predlog)
    # (b') aktivna prijemnica broja iz kolizije ostala bez stavki: sve su stornirane
    docs = defaultdict(list)
    for p in aktivni(ctx.tab("tblPrijemnica")):
        nb = norm_broj(polje(p, "BrojPrijemnice"))
        pid = k(p, "PrijemnicaID")
        if nb not in kol or n(p, "KolAmbalaze") <= 0 or akt_po_p.get(pid):
            continue
        stor = [x for x in ctx.pls_po_prijemnici.get(pid, ()) if storno(x)]
        if not stor:
            continue
        detach = [x for x in stor if _istorija_detach(r.pal.get(k(x, "PaletaID")), s(p, "BrojPrijemnice"))]
        # stavka na storniranoj paleti bez traga DETACH = StornoPaleta celog dokumenta (iskljuceno)
        ostale = [x for x in stor if x not in detach and r.paleta_aktivna(x)]
        if not detach and not ostale:
            continue
        docs[_dok("tblPrijemnica", p)].append((p, detach, ostale))
    for lst in docs.values():
        p0 = lst[0][0]
        detach = [x for _, dl, _ in lst for x in dl]
        stv = detach + [x for _, _, ol in lst for x in ol]
        prj_ids = ids(_jedinstveni([p for p, _, _ in lst], "PrijemnicaID"), "PrijemnicaID")
        drugi = [q for q in ctx.tab("tblPrijemnica") if norm_broj(polje(q, "BrojPrijemnice")) ==
                 norm_broj(polje(p0, "BrojPrijemnice")) and k(q, "KupacID") != k(p0, "KupacID")]
        opis = ("Aktivna prijemnica %s (%s gajbi) nema nijednu aktivnu paletnu stavku, a ima stornirane %s; "
                "isti broj ima i %s." % (_prj_txt(p0), kg_txt(sum(n(p, "KolAmbalaze") for p, _, _ in lst)),
                                         ids(stv, "StavkaID"), "; ".join(_prj_txt(q) for q in drugi[:3])))
        if detach:
            opis += " Istorija palete %s: DETACH po broju %s." % (
                ", ".join(sorted({s(x, "PaletaID") for x in detach})), s(p0, "BrojPrijemnice"))
        ctx.nalaz(KD if detach else K, SUMNJA, "tblPrijemnica", prj_ids, opis,
                  "Ako paleta fizicki postoji: vratiti storno stavki %s (Stornirano = '' i uvecati paletu) ili "
                  "ponovo paletizovati %s." % (ids(stv, "StavkaID"), prj_ids))
    return pokriveno


def _pal_zbirna_po_broju(r):
    """ReassignPrijemnicaToZbirna_TX (od v2.21.0, d21:3188) menja BrojZbirne svim AKTIVNIM prijemnicama i
    stavkama datog BROJA, bez kupca. Potpis: dve aktivne prijemnice istog broja, razlicitih kupaca, na istoj
    zbirni, koja pripada samo jednom od njih."""
    K = "PAL_ZBIRNA_PREVEZANA_PO_BROJU"
    if not (r.kolone(K, "tblPrijemnica", ["PrijemnicaID", "BrojPrijemnice", "KupacID", "BrojZbirne"])
            and r.kolone(K, "tblZbirna", ["BrojZbirne", "KupacID"])):
        return
    ctx = r.ctx
    zbr = _zbirne_po_broju(ctx)
    po_broju = defaultdict(list)
    for p in aktivni(ctx.tab("tblPrijemnica")):
        if k(p, "BrojPrijemnice") and k(p, "KupacID"):
            po_broju[norm_broj(polje(p, "BrojPrijemnice"))].append(p)
    docs = defaultdict(list)
    for nb, lst in po_broju.items():
        if len({k(p, "KupacID") for p in lst}) < 2:
            continue
        for p in lst:
            bz = norm_broj(polje(p, "BrojZbirne"))
            zk = {k(z, "KupacID") for z in zbr.get(bz, ()) if k(z, "KupacID")}
            if not bz or not zk or k(p, "KupacID") in zk:
                continue
            drugi = [q for q in lst if k(q, "KupacID") in zk and norm_broj(polje(q, "BrojZbirne")) == bz]
            if drugi:
                docs[_dok("tblPrijemnica", p)].append((p, drugi[0]))
    for lst in docs.values():
        p, q = lst[0]
        z = zbr[norm_broj(polje(p, "BrojZbirne"))][0]
        ctx.nalaz(K, SUMNJA, "tblPrijemnica", ids(_jedinstveni([x for x, _ in lst], "PrijemnicaID"), "PrijemnicaID"),
                  "Prijemnica %s je na zbirni %s kupca %s, kao i prijemnica %s istog broja: prevezivanje zbirne po "
                  "broju je zahvatilo obe." % (_prj_txt(p), s(p, "BrojZbirne"), s(z, "KupacID"), _prj_txt(q)),
                  "BrojZbirne prijemnice %s (i njenih paletnih stavki) vratiti na zbirnu kupca %s koju je imala pre "
                  "prevezivanja; prijemnica %s ostaje." % (s(p, "PrijemnicaID"), s(p, "KupacID"), s(q, "PrijemnicaID")))


def _pal_broj_zbirne(r, pokriveno):
    """BrojZbirne paletne stavke != BrojZbirne prijemnice: ReassignPrijemnicaToZbirna_TX do v2.21.0 menja
    samo prijemnicu (d11:3138-3202), pa paletni list nosi pogresne kooperante."""
    K = "PAL_BROJ_ZBIRNE_ZASTAREO"
    if not (r.kolone(K, "tblPaletaStavka", ["StavkaID", "PaletaID", "PrijemnicaID", "BrojZbirne"])
            and r.kolone(K, "tblPrijemnica", ["PrijemnicaID", "BrojZbirne"])):
        return
    ctx = r.ctx
    zbr = _zbirne_po_broju(ctx)

    def slaze(bz_txt, p):
        zl = zbr.get(norm_broj(bz_txt), [])
        if not bz_txt:
            return False, "prazno"
        if not zl:
            return False, "zbirna ne postoji ili je stornirana"
        z = next((x for x in zl if k(x, "VozacID") == k(p, "VozacID")), None)
        if z is None:
            return False, "zbirna drugog vozaca (%s)" % s(zl[0], "VozacID")
        return True, "zbirna vozaca prijemnice %s, %s" % (s(z, "VozacID"), dan_txt(d(z, "Datum")))

    grupe = defaultdict(list)
    for st in r.stavke_aktivne():
        p = r.prj.get(k(st, "PrijemnicaID"))
        if p is None or storno(p) or k(p, "PrijemnicaID") in pokriveno:
            continue
        if norm_broj(polje(st, "BrojZbirne")) == norm_broj(polje(p, "BrojZbirne")):
            continue
        grupe[_dok("tblPrijemnica", p)].append((st, p))
    for lst in grupe.values():
        stv = [x for x, _ in lst]
        p = lst[0][1]
        bs, bp = s(lst[0][0], "BrojZbirne"), s(p, "BrojZbirne")
        oba = any(k(x, "BrojZbirne") and k(q, "BrojZbirne") for x, q in lst)
        ok_s, txt_s = slaze(bs, p)
        ok_p, txt_p = slaze(bp, p)
        palete = ", ".join(sorted({s(x, "PaletaID") for x in stv}))
        if ok_s and not ok_p:
            predlog = ("Verovatno je pogresna PRIJEMNICA: BrojZbirne prijemnice %s -> %s (vidi oblast prijemnice); "
                       "stavke ostaju." % (bp or "(prazno)", bs))
        else:
            predlog = ("BrojZbirne stavki %s: %s -> %s, tek posle provere prijemnice; paleta %s -- odstampani "
                       "paletni list nosi pogresne kooperante." % (ids(stv, "StavkaID"), bs or "(prazno)",
                                                                   bp or "(prazno)", palete))
        ctx.nalaz(K, GRESKA if oba else SUMNJA, "tblPaletaStavka", ids(stv, "StavkaID"),
                  "Stavke %s (paleta %s) nose BrojZbirne %s [%s], a prijemnica %s nosi %s [%s]." % (
                      ids(stv, "StavkaID"), palete, bs or "(prazno)", txt_s, _prj_txt(p), bp or "(prazno)", txt_p),
                  predlog)


def _pal_gajbe_kg(r, pokriveno):
    """(a) zbir gajbi/neto aktivnih stavki jedne prijemnice (po PrijemnicaID) != prijemnica -- po konstrukciji
    su jednaki (PaletizePrijemnica deli tacno KolAmbalaze, p11:100-140); (b) zaglavlje palete != zbir stavki
    (RecomputePaletaFromStavke, p28:2198)."""
    K = "PAL_GAJBE_KG_NESAGLASNI"
    if not r.kolone(K, "tblPaletaStavka", ["StavkaID", "PaletaID", "PrijemnicaID", "BrojGajbica", "NetoKg"]):
        return
    ctx = r.ctx
    sve_akt = [x for x in ctx.tab("tblPaletaStavka") if not storno(x)]
    # (a)
    if r.kolone(K, "tblPrijemnica", ["PrijemnicaID", "KolAmbalaze", "Kolicina"]):
        po_p = defaultdict(list)
        for st in sve_akt:
            po_p[k(st, "PrijemnicaID")].append(st)
        grupe = defaultdict(list)
        for pid, stv in po_p.items():
            p = r.prj.get(pid)
            if p is None or storno(p) or pid in pokriveno:
                continue
            if not all(r.paleta_aktivna(x) for x in stv):
                continue    # deo stavki na storniranoj paleti: nejasno, ne prijavljuje se
            g = sum(n(x, "BrojGajbica") for x in stv)
            kg = sum(n(x, "NetoKg") for x in stv)
            raz_g = ima_broj(p, "KolAmbalaze") and g != n(p, "KolAmbalaze")
            raz_kg = (ima_broj(p, "Kolicina") and all(ima_broj(x, "NetoKg") for x in stv)
                      and abs(kg - n(p, "Kolicina")) > TOL_KG)
            if raz_g or raz_kg:
                grupe[_dok("tblPrijemnica", p)].append((p, stv, g, kg, raz_g))
        for lst in grupe.values():
            delovi = ["%s Klasa %s: gajbi %s / %s, neto %s / %s (stavke %s)" % (
                s(p, "PrijemnicaID"), s(p, "Klasa") or "-", kg_txt(g), kg_txt(n(p, "KolAmbalaze")), kg_txt(kg),
                kg_txt(n(p, "Kolicina")), ids(stv, "StavkaID")) for p, stv, g, kg, _ in lst]
            ctx.nalaz(K, GRESKA if any(x[4] for x in lst) else SUMNJA, "tblPrijemnica",
                      ids(_jedinstveni([x[0] for x in lst], "PrijemnicaID"), "PrijemnicaID"),
                      "Paletne stavke prijemnice %s ne daju njene gajbe/neto (palete / prijemnica): %s." % (
                          _prj_txt(lst[0][0]), "; ".join(delovi)),
                      "Kad se gajbe poklapaju: skalirati NetoKg stavki na Kolicina prijemnice; razliku gajbi "
                      "proveriti na paleti (relink/adjust/detach ili rucna izmena).")
    # (b)
    if not r.kolone(K, "tblPaleta", ["PaletaID", "BrojGajbica", "NetoKg"]):
        return
    po_pal = defaultdict(list)
    for st in sve_akt:
        po_pal[k(st, "PaletaID")].append(st)
    for lp in aktivni(ctx.tab("tblPaleta")):
        stv = po_pal.get(k(lp, "PaletaID"), [])
        sg = sum(n(x, "BrojGajbica") for x in stv)
        sn = sum(n(x, "NetoKg") for x in stv)
        sa = sum(n(x, "AmbalazaKg") for x in stv)
        raz, raz_g = [], False
        if ima_broj(lp, "BrojGajbica") and n(lp, "BrojGajbica") != sg:
            raz.append("BrojGajbica %s -> %s" % (kg_txt(n(lp, "BrojGajbica")), kg_txt(sg)))
            raz_g = True
        if ima_broj(lp, "NetoKg") and abs(n(lp, "NetoKg") - sn) > TOL_KG:
            raz.append("NetoKg %s -> %s" % (kg_txt(n(lp, "NetoKg")), kg_txt(sn)))
        if ima_broj(lp, "AmbalazaKg") and abs(n(lp, "AmbalazaKg") - sa) > TOL_KG:
            raz.append("AmbalazaKg %s -> %s" % (kg_txt(n(lp, "AmbalazaKg")), kg_txt(sa)))
        if ima_broj(lp, "BrutoKg") and ima_broj(lp, "NetoKg") and ima_broj(lp, "PaletaKg"):
            br = n(lp, "NetoKg") + n(lp, "AmbalazaKg") + n(lp, "PaletaKg")
            if abs(n(lp, "BrutoKg") - br) > TOL_KG:
                raz.append("BrutoKg %s != Neto+Ambalaza+Paleta %s" % (kg_txt(n(lp, "BrutoKg")), kg_txt(br)))
        if not raz:
            continue
        ctx.nalaz(K, GRESKA if raz_g else SUMNJA, "tblPaleta", s(lp, "PaletaID"),
                  "Zaglavlje palete %s (broj %s) ne odgovara zbiru %d aktivnih stavki: %s." % (
                      s(lp, "PaletaID"), s(lp, "BrojPalete") or "-", len(stv), "; ".join(raz)),
                  "Zaglavlje preracunati iz aktivnih stavki (kao RecomputePaletaFromStavke): %s." % "; ".join(raz))


def _pal_identitet(r):
    """Stavka/paleta druge vrste, sorte, klase ili tipa ambalaze od prijemnice: re-point stavke menja
    PrijemnicaID/BrojPrijemnice/BrojZbirne, a ne vrstu/sortu stavke ni zaglavlje palete (p11:1287-1306)."""
    K = "PAL_IDENTITET_RAZLICIT_OD_PRIJEMNICE"
    if not (r.kolone(K, "tblPaletaStavka", ["StavkaID", "PaletaID", "PrijemnicaID", "VrstaVoca", "SortaVoca",
                                             "Klasa"])
            and r.kolone(K, "tblPrijemnica", ["PrijemnicaID", "VrstaVoca", "SortaVoca", "Klasa"])):
        return
    ctx = r.ctx
    zaglavlje = bool(ctx.tab("tblPaleta")) and r.kolone(K, "tblPaleta", ["PaletaID", "VrstaVoca", "SortaVoca",
                                                                          "Klasa"])
    po_pal = defaultdict(list)
    for st in r.stavke_aktivne():
        p = r.prj.get(k(st, "PrijemnicaID"))
        if p is not None and k(st, "PaletaID"):
            po_pal[k(st, "PaletaID")].append((st, p))
    polja = ("VrstaVoca", "SortaVoca", "Klasa")
    for pal_id, lst in po_pal.items():
        lp = r.pal.get(pal_id)
        stavka_raz = []
        for st, p in lst:
            for f in polja:
                if _razlicito(k(st, f), k(p, f)):
                    stavka_raz.append("%s %s: stavka %s, prijemnica %s %s" % (
                        s(st, "StavkaID"), f, s(st, f), s(p, "PrijemnicaID"), s(p, f)))
        ident = {}
        for f in polja:
            ident[f] = {k(st, f) or k(p, f) for st, p in lst} - {""}
        ident["TipAmbalaze"] = {norm_tip(polje(p, "TipAmbalaze")) for _, p in lst} - {""}
        hetero = [f for f, v in ident.items() if len(v) > 1]
        zag_raz = []
        if lp is not None and zaglavlje and not hetero:
            for f, v in ident.items():
                if not v:
                    continue
                zv = norm_tip(polje(lp, f)) if f == "TipAmbalaze" else k(lp, f)
                if _razlicito(zv, next(iter(v))):
                    cilj = next((x for st, p in lst for x in ((s(st, f) if f != "TipAmbalaze" else ""), s(p, f))
                                 if x), "")
                    zag_raz.append("%s %s -> %s" % (f, s(lp, f), cilj))
        if not (stavka_raz or hetero or zag_raz):
            continue
        delovi = []
        if stavka_raz:
            delovi.append("stavka != prijemnica: " + "; ".join(stavka_raz))
        if hetero:
            delovi.append("heterogena paleta (%s): %s" % (", ".join(hetero), "; ".join(
                "%s %s/%s/%s" % (s(st, "StavkaID"), s(st, "VrstaVoca") or s(p, "VrstaVoca"),
                                 s(st, "SortaVoca") or s(p, "SortaVoca"), s(st, "Klasa") or s(p, "Klasa"))
                for st, p in lst)))
        if zag_raz:
            delovi.append("zaglavlje palete != prijemnice: " + "; ".join(zag_raz))
        jedna = len({k(p, "PrijemnicaID") for _, p in lst}) == 1 or not hetero
        predlog = ("Relabel palete na identitet prijemnice (%s)." % "; ".join(zag_raz) if zag_raz and not stavka_raz
                   else "Cela paleta jedne robe: relabel stavki/palete na identitet prijemnice." if jedna
                   else "Deljena paleta: stavku drugog identiteta skinuti i paletizovati ponovo.")
        pid = s(lst[0][0], "PaletaID")
        ctx.nalaz(K, SUMNJA, "tblPaleta", pid,
                  "Paleta %s (stavke %s): %s." % (pid, ids([x for x, _ in lst], "StavkaID"), " | ".join(delovi)),
                  predlog)


def _pal_mesani_kupci(r):
    """GetOrCreateOpenPaleta ne gleda kupca (p11:1007-1018): paleta moze da skupi prijemnice dva kupca."""
    K = "PAL_MESANI_KUPCI"
    if not (r.kolone(K, "tblPaletaStavka", ["PaletaID", "PrijemnicaID"])
            and r.kolone(K, "tblPrijemnica", ["PrijemnicaID", "KupacID"])):
        return
    po_pal = defaultdict(list)
    for st in r.stavke_aktivne():
        p = r.prj.get(k(st, "PrijemnicaID"))
        if p is not None and k(p, "KupacID") and k(st, "PaletaID"):
            po_pal[k(st, "PaletaID")].append((st, p))
    for lst in po_pal.values():
        kupci = sorted({s(p, "KupacID") for _, p in lst})
        if len({x.casefold() for x in kupci}) < 2:
            continue
        pid = s(lst[0][0], "PaletaID")
        r.ctx.nalaz(K, INFO, "tblPaleta", pid,
                    "Paleta %s ima stavke prijemnica %d kupaca: %s." % (pid, len(kupci), "; ".join(
                        "%s -> %s" % (s(st, "StavkaID"), _prj_txt(p)) for st, p in lst)),
                    "Nema automatske ispravke: operater potvrdjuje da je paleta fizicki jedna.")


def _naslednici(r):
    """tblStornoVeze: OldDocID -> [aktivna prijemnica NewDocID]."""
    if r._nasl is None:
        out = defaultdict(list)
        for v in r.ctx.tab("tblStornoVeze"):
            x = r.prj.get(k(v, "NewDocID"))
            if x is not None and not storno(x) and k(v, "OldDocID"):
                out[k(v, "OldDocID")].append(x)
        r._nasl = out
    return r._nasl


def _blizu(a, b):
    return a > 0 and abs(b - a) <= a / 10


def _zamene(r, p, klasa, kol=None):
    """Kandidati zamene stornirane prijemnice p: tblStornoVeze naslednik, pa aktivna prijemnica UNETA POSLE p
    (zamena se unosi posle storna), istog kupca, klase, vrste i sorte (prazno = nepoznato), i to iste zbirne,
    ili istog dana i vozaca sa kolicinom +-10% (isti dan sam nije dovoljan: vise tura istog dana).
    kol: kolicina linije fakture -- tada se +-10% trazi uvek."""
    ctx = r.ctx
    out = list(_naslednici(r).get(k(p, "PrijemnicaID"), []))
    bz, dan, red_p = norm_broj(polje(p, "BrojZbirne")), d(p, "Datum"), p.get("_redni", 0)
    for q in aktivni(ctx.tab("tblPrijemnica")):
        if q is p or k(q, "KupacID") != k(p, "KupacID") or k(q, "PrijemnicaID") == k(p, "PrijemnicaID"):
            continue
        if q.get("_redni", 0) <= red_p or _razlicito(k(q, "Klasa"), klasa):
            continue
        if any(_razlicito(k(q, f), k(p, f)) for f in ("VrstaVoca", "SortaVoca")):
            continue
        ista_zbirna = bool(bz) and norm_broj(polje(q, "BrojZbirne")) == bz
        isti_dan = (dan is not None and d(q, "Datum") == dan and bool(k(p, "VozacID"))
                    and k(q, "VozacID") == k(p, "VozacID") and _blizu(n(p, "Kolicina"), n(q, "Kolicina")))
        if not (ista_zbirna or isti_dan):
            continue
        if kol is not None and not _blizu(kol, n(q, "Kolicina")):
            continue
        out.append(q)
    return _jedinstveni(out, "PrijemnicaID")


def _pal_na_storniranoj(r):
    """Aktivne paletne stavke na storniranoj prijemnici: StornoPrijemnica ne dira tblPaletaStavka
    (vba-v2.11.0:modStorno.bas:541-593); od v2.21 i kroz OTKAZI ili neuspeo relink."""
    K = "PAL_STAVKA_NA_STORNIRANOJ_PRIJEMNICI"
    if not (r.kolone(K, "tblPaletaStavka", ["StavkaID", "PaletaID", "PrijemnicaID", "BrojGajbica"])
            and r.kolone(K, "tblPrijemnica", ["PrijemnicaID", "Stornirano"])):
        return
    ctx = r.ctx
    akt_po_p = defaultdict(list)
    for st in ctx.tab("tblPaletaStavka"):
        if not storno(st):
            akt_po_p[k(st, "PrijemnicaID")].append(st)
    grupe = defaultdict(list)
    for st in r.stavke_aktivne():
        p = r.prj.get(k(st, "PrijemnicaID"))
        if p is not None and storno(p):
            grupe[_dok("tblPrijemnica", p)].append((st, p))
    for lst in grupe.values():
        stv = [x for x, _ in lst]
        prj = _jedinstveni([p for _, p in lst], "PrijemnicaID")
        gajbe = sum(n(x, "BrojGajbica") for x in stv)
        dupl, izg = [], []
        for p in prj:
            for q in _zamene(r, p, k(p, "Klasa")):
                if akt_po_p.get(k(q, "PrijemnicaID")):
                    dupl.append(q)
                elif n(q, "KolAmbalaze") > 0:
                    izg.append(q)
        dupl, izg = _jedinstveni(dupl, "PrijemnicaID"), _jedinstveni(izg, "PrijemnicaID")
        prer = sorted({s(x, "PaletaID") for x in stv
                       if r.pal.get(k(x, "PaletaID")) is not None and k(r.pal[k(x, "PaletaID")], "Preradjeno") == "da"})
        opis = "Aktivne stavke %s (palete %s, %s gajbi) pokazuju na storniranu prijemnicu %s." % (
            ids(stv, "StavkaID"), ", ".join(sorted({s(x, "PaletaID") for x in stv})), kg_txt(gajbe), _prj_txt(prj[0]))
        if dupl:
            opis += " Zamena %s ima svoje aktivne stavke: gajbe se broje dvaput (visak %s)." % (
                ", ".join(s(q, "PrijemnicaID") for q in dupl), kg_txt(gajbe))
            predlog = "Skinuti osirocene stavke %s (Stornirano = Da + umanjiti paletu)." % ids(stv, "StavkaID")
        elif izg:
            opis += " Zamena %s nema nijednu aktivnu stavku (izgubljena paletizacija)." % ", ".join(
                s(q, "PrijemnicaID") for q in izg)
            predlog = "Prevezati stavke %s na %s iste klase (PrijemnicaID, BrojPrijemnice, BrojZbirne iz zamene)." % (
                ids(stv, "StavkaID"), _prj_txt(izg[0]))
        else:
            predlog = "Nema zamene: skinuti stavke %s (roba nije stigla) ili naci zamenu rucno." % ids(stv, "StavkaID")
        if prer:
            opis += " Palete %s su preradjene: u preradu je usla stornirana roba." % ", ".join(prer)
        ctx.nalaz(K, GRESKA, "tblPaletaStavka", ids(stv, "StavkaID"), opis, predlog)


# =====================================================================
# FAKTURE (C002 nema faktura: sve provere tiho rade nad praznim tabelama)
# =====================================================================

def _fak_sef(r):
    """Faktura stornirana/otkazana na SEF-u, a lokalno aktivna: StornoInvoiceOnSEF_TX i status sync
    azuriraju samo SEF polja (vba-v2.28.4:modSEFService.bas:546-600, modSEFStatusSync.bas:117-140)."""
    K = "FAK_SEF_STORNO_BEZ_LOKALNOG"
    if not r.kolone(K, "tblFakture", ["FakturaID", "SEFStatus"]):
        return
    for f in r.ctx.tab("tblFakture"):
        if not _fak_aktivna(f) or s(f, "SEFStatus").upper() not in ("STORNO", "CANCELLED", "CANCELED"):
            continue
        r.ctx.nalaz(K, GRESKA, "tblFakture", s(f, "FakturaID"),
                    "Faktura %s (broj %s, kupac %s, %s) ima SEFStatus %s, a lokalno je aktivna (SEFDocumentId %s)." % (
                        s(f, "FakturaID"), s(f, "BrojFakture") or "-", s(f, "KupacID") or "-",
                        dan_txt(d(f, "Datum")), s(f, "SEFStatus"), s(f, "SEFDocumentId") or "-"),
                    "Lokalni storno fakture (linije, oslobadjanje prijemnica, odvezivanje novca).")


def _aktivne_linije(r):
    """(linija, faktura) za aktivne linije aktivnih faktura."""
    out = []
    for ln in r.ctx.tab("tblFakturaStavke"):
        f = r.fak.get(k(ln, "FakturaID"))
        if f is not None and _fak_aktivna(f) and not storno(ln):
            out.append((ln, f))
    return out


def _fakture_po_prj(r):
    g = defaultdict(set)
    for ln, f in _aktivne_linije(r):
        if k(ln, "PrijemnicaID"):
            g[k(ln, "PrijemnicaID")].add(k(f, "FakturaID"))
    return g


def _fak_linije(r):
    """Linija aktivne fakture naspram prijemnice na koju pokazuje. RelinkFakturaStavke prevezuje po broju
    bez kupca (d11:1859-1954) i menja samo PrijemnicaID/OsirocenoOd; StornoPrijemnica ostavlja fakturu
    i liniju aktivne (s11:576-584)."""
    KT = "FAK_STAVKA_TUDJEG_KUPCA"
    KR = "FAK_RELINK_TUDJA_ILI_STORNIRANA"
    KZ = "FAK_STAVKA_ZASTARELA_POSLE_RELINKA"
    KA = "FAK_AKTIVNA_NA_STORNIRANOJ_PRIJEMNICI"
    KL = "FAK_LINIJA_NA_STORNIRANOJ_PRIJEMNICI"
    ctx = r.ctx
    if not ctx.tab("tblFakturaStavke") or not ctx.tab("tblPrijemnica"):
        return
    lin = ["StavkaID", "FakturaID", "PrijemnicaID"]
    ok_t = (r.kolone(KT, "tblFakture", ["FakturaID", "KupacID"]) and r.kolone(KT, "tblFakturaStavke", lin)
            and r.kolone(KT, "tblPrijemnica", ["PrijemnicaID", "KupacID"]))
    ok_z = (r.kolone(KZ, "tblFakture", ["FakturaID"])
            and r.kolone(KZ, "tblFakturaStavke", lin + ["Kolicina", "Cena", "Klasa", "BrojPrijemnice"])
            and r.kolone(KZ, "tblPrijemnica", ["PrijemnicaID", "Kolicina", "Cena", "Klasa", "BrojPrijemnice"]))
    ok_s = (r.kolone(KL, "tblFakture", ["FakturaID"]) and r.kolone(KL, "tblFakturaStavke", lin + ["Kolicina"])
            and r.kolone(KL, "tblPrijemnica", ["PrijemnicaID", "KupacID", "Stornirano"]))
    fak_po_prj = _fakture_po_prj(r)
    tudji, zast, osir = defaultdict(list), defaultdict(list), defaultdict(list)
    for ln, f in _aktivne_linije(r):
        p = r.prj.get(k(ln, "PrijemnicaID"))
        if p is None:
            continue
        g = (k(f, "FakturaID"),) + _dok("tblPrijemnica", p)
        if storno(p):
            if ok_s:
                osir[g].append((ln, f, p))
            continue
        if ok_t and _razlicito(k(p, "KupacID"), k(f, "KupacID")):
            tudji[g].append((ln, f, p))
            continue
        if not ok_z:
            continue
        raz, samo_broj = [], True
        for fl, fp in (("Kolicina", "Kolicina"), ("Cena", "Cena")):
            if ima_broj(ln, fl) and ima_broj(p, fp) and abs(n(ln, fl) - n(p, fp)) > TOL_LINIJA:
                raz.append("%s %s / %s" % (fl, kg_txt(n(ln, fl)), kg_txt(n(p, fp))))
                samo_broj = False
        if _razlicito(k(ln, "Klasa"), k(p, "Klasa")):
            raz.append("Klasa %s / %s" % (s(ln, "Klasa"), s(p, "Klasa")))
            samo_broj = False
        if _razlicito(norm_broj(polje(ln, "BrojPrijemnice")), norm_broj(polje(p, "BrojPrijemnice"))):
            raz.append("BrojPrijemnice %s / %s (rucna izmena: relink trazi isti broj)" % (
                s(ln, "BrojPrijemnice"), s(p, "BrojPrijemnice")))
        if raz:
            zast[g].append((ln, f, p, raz, samo_broj))
    for lst in tudji.values():
        ln, f, p = lst[0]
        relink = (not k(ln, "OsirocenoOd") and d(p, "Datum") is not None and d(f, "Datum") is not None
                  and d(p, "Datum") > d(f, "Datum"))
        kand = [q for q in aktivni(ctx.tab("tblPrijemnica"))
                if k(q, "KupacID") == k(f, "KupacID") and k(q, "BrojPrijemnice")
                and norm_broj(polje(q, "BrojPrijemnice")) == norm_broj(polje(ln, "BrojPrijemnice"))]
        opis = "Linije %s fakture %s (kupac %s, %s) pokazuju na prijemnicu %s drugog kupca%s." % (
            ids([x for x, _, _ in lst], "StavkaID"), s(f, "FakturaID"), s(f, "KupacID"), dan_txt(d(f, "Datum")),
            _prj_txt(p), "; prijemnica je mladja od fakture i linija nema OsirocenoOd -- znak relinka po broju"
            if relink else "")
        kupci_broja = r.kolizije().get(norm_broj(polje(ln, "BrojPrijemnice")))
        if kupci_broja:
            opis += " Broj %s se javlja kod kupaca %s." % (s(ln, "BrojPrijemnice"),
                                                         ", ".join(sorted(x.upper() for x in kupci_broja)))
        predlog = ("Vratiti liniju na prijemnicu kupca fakture: %s; prijemnicu %s osloboditi "
                   "(Fakturisano/FakturaID)." % (
                       _prj_txt(kand[0]) if kand else "istog BrojPrijemnice i KupacID (stornirana je cesto prava)",
                       s(p, "PrijemnicaID")))
        ctx.nalaz(KR if relink else KT, SUMNJA, "tblFakturaStavke", ids([x for x, _, _ in lst], "StavkaID"),
                  opis, predlog)
    for lst in zast.values():
        ln, f, p = lst[0][:3]
        ctx.nalaz(KZ, INFO if all(x[4] for x in lst) else SUMNJA, "tblFakturaStavke",
                  ids([x[0] for x in lst], "StavkaID"),
                  "Linije fakture %s ne odgovaraju prijemnici %s (linija / prijemnica): %s." % (
                      s(f, "FakturaID"), _prj_txt(p), "; ".join("%s: %s" % (s(x[0], "StavkaID"), ", ".join(x[3]))
                                                              for x in lst)),
                  "Prikazati obe vrednosti operateru; ako je faktura poslata (SEF), ispravka ide knjiznim "
                  "odobrenjem/zaduzenjem, ne prepisom.")
    for lst in osir.values():
        ln, f, p = lst[0]
        lnk = [x for x, _, _ in lst]
        kol = sum(n(x, "Kolicina") for x in lnk)
        zam = _zamene(r, p, k(ln, "Klasa") or k(p, "Klasa"), kol)
        # zamena na bilo kojoj aktivnoj fakturi (i na istoj, drugom linijom) = ista roba dvaput
        dvaput = [q for q in zam if fak_po_prj.get(k(q, "PrijemnicaID"))]
        osir_txt = "; OsirocenoOd = %s" % s(ln, "OsirocenoOd") if k(ln, "OsirocenoOd") else ""
        opis = "Aktivne linije %s aktivne fakture %s (kupac %s) pokazuju na storniranu prijemnicu %s, %s kg%s." % (
            ids(lnk, "StavkaID"), s(f, "FakturaID"), s(f, "KupacID"), _prj_txt(p), kg_txt(kol), osir_txt)
        if dvaput:
            q = dvaput[0]
            ctx.nalaz(KA, SUMNJA, "tblFakturaStavke", ids(lnk, "StavkaID"),
                      opis + " Zamena %s (%s kg) je fakturisana na %s: ista roba fakturisana dvaput." % (
                          _prj_txt(q), kg_txt(n(q, "Kolicina")),
                          ", ".join(sorted(x.upper() for x in fak_po_prj[k(q, "PrijemnicaID")]))),
                      "Jednu od faktura stornirati ili umanjiti knjiznim odobrenjem (odluka operatera).")
            continue
        ctx.nalaz(KL, SUMNJA, "tblFakturaStavke", ids(lnk, "StavkaID"), opis,
                  ("Prevezati liniju na nefakturisanu zamenu %s i uskladiti kolicinu/cenu." % _prj_txt(zam[0]))
                  if zam else "Nema zamene: storno fakture ili knjizno odobrenje, po odluci operatera.")
    # Iznos fakture != zbir linija (ApplyAvansToFaktura nije citan -- INFO)
    if ok_z and ctx.ima_kolonu("tblFakture", "Iznos"):
        sve, akt = defaultdict(Decimal), defaultdict(Decimal)
        for ln in ctx.tab("tblFakturaStavke"):
            v = n(ln, "Kolicina") * n(ln, "Cena")
            sve[k(ln, "FakturaID")] += v
            if not storno(ln):
                akt[k(ln, "FakturaID")] += v
        for f in ctx.tab("tblFakture"):
            fid = k(f, "FakturaID")
            if not _fak_aktivna(f) or not ima_broj(f, "Iznos") or fid not in sve:
                continue
            iz = n(f, "Iznos")
            if abs(iz - sve[fid]) > TOL_IZNOS and abs(iz - akt[fid]) > TOL_IZNOS:
                ctx.nalaz(KZ, INFO, "tblFakture", s(f, "FakturaID"),
                          "Iznos fakture %s je %s, a zbir Kolicina x Cena linija %s (aktivnih %s)." % (
                              s(f, "FakturaID"), kg_txt(iz), kg_txt(sve[fid]), kg_txt(akt[fid])),
                          "Proveriti: CreateFaktura upisuje Iznos = zbir linija; razlika znaci izmenu linija ili "
                          "rucnu izmenu (ApplyAvansToFaktura nije proveren).")


def _fak_status(r):
    """Fakturisano/FakturaID na prijemnici naspram linija aktivnih faktura (istina = linije). Poseban trag:
    linija stornirane fakture prevezana relinkom na novu prijemnicu (StornoFakturaStavke... ne brise
    OsirocenoOd, s11:1040-1050; Relink ne gleda storno linije, d11:1892-1937)."""
    K = "FAK_STATUS_PRIJEMNICE_NESAGLASAN"
    KR = "FAK_RELINK_TUDJA_ILI_STORNIRANA"
    ctx = r.ctx
    if not ctx.tab("tblPrijemnica"):
        return
    ima_fak = bool(ctx.tab("tblFakture") or ctx.tab("tblFakturaStavke"))
    if not (ctx.ima_kolonu("tblPrijemnica", "Fakturisano") and ctx.ima_kolonu("tblPrijemnica", "FakturaID")):
        if ima_fak:
            r.kolone(K, "tblPrijemnica", ["PrijemnicaID", "Fakturisano", "FakturaID"])
        return
    if ctx.tab("tblFakturaStavke") and not r.kolone(K, "tblFakturaStavke", ["FakturaID", "PrijemnicaID"]):
        return
    fak_po_prj = _fakture_po_prj(r)
    lin_po_prj = defaultdict(list)
    for ln in ctx.tab("tblFakturaStavke"):
        lin_po_prj[k(ln, "PrijemnicaID")].append(ln)
    stat, rel = defaultdict(list), defaultdict(list)
    for p in aktivni(ctx.tab("tblPrijemnica")):
        pid = k(p, "PrijemnicaID")
        fakt, fid = k(p, "Fakturisano") == "da", k(p, "FakturaID")
        a = fak_po_prj.get(pid, set())
        if not (fakt or fid or a):
            continue
        mrtve = [ln for ln in lin_po_prj.get(pid, ()) if fid and k(ln, "FakturaID") == fid
                 and (storno(ln) or r.fak.get(fid) is None or not _fak_aktivna(r.fak[fid]))]
        if mrtve and fid not in a:
            rel[_dok("tblPrijemnica", p)].append((p, mrtve))
            continue
        razlozi, teska = [], False
        if fid and (r.fak.get(fid) is None or not _fak_aktivna(r.fak[fid])):
            razlozi.append("FakturaID %s %s" % (s(p, "FakturaID"), "ne postoji" if r.fak.get(fid) is None
                                                 else "je stornirana"))
            teska = True
        if fakt and not a:
            razlozi.append("Fakturisano = Da, a nema aktivne linije na aktivnoj fakturi")
        if a and not fakt:
            razlozi.append("na aktivnoj fakturi %s, a Fakturisano nije Da" % ", ".join(sorted(x.upper() for x in a)))
        if a and fid not in a:
            razlozi.append("FakturaID %s nije medju fakturama linija (%s)" % (
                s(p, "FakturaID") or "(prazan)", ", ".join(sorted(x.upper() for x in a))))
        if len(a) >= 2:
            razlozi.append("na %d aktivne fakture: %s" % (len(a), ", ".join(sorted(x.upper() for x in a))))
            teska = True
        if razlozi:
            stat[_dok("tblPrijemnica", p)].append((p, razlozi, teska))
    for lst in rel.values():
        p, mrtve = lst[0]
        ctx.nalaz(KR, GRESKA, "tblPrijemnica", ids(_jedinstveni([x for x, _ in lst], "PrijemnicaID"), "PrijemnicaID"),
                  "Prijemnica %s ima Fakturisano/FakturaID = %s preko linije %s stornirane fakture ili stornirane "
                  "linije: relink je ozivio mrtvu liniju." % (_prj_txt(p), s(p, "FakturaID"), ids(mrtve, "StavkaID")),
                  "Fakturisano = '' i FakturaID = '' na prijemnici (otvorena je za fakturisanje); u novoj svesci je "
                  "nefakturisana.")
    for lst in stat.values():
        p = lst[0][0]
        ctx.nalaz(K, GRESKA if any(x[2] for x in lst) else SUMNJA, "tblPrijemnica",
                  ids(_jedinstveni([x[0] for x in lst], "PrijemnicaID"), "PrijemnicaID"),
                  "Prijemnica %s: %s." % (_prj_txt(p), "; ".join(
                      "%s: %s" % (s(x[0], "PrijemnicaID"), ", ".join(x[1])) for x in lst)),
                  "Istina su linije aktivnih faktura: uskladiti Fakturisano/FakturaID; operater dobija spisak da "
                  "prijemnicu ne fakturise ponovo nesvesno (main status izvodi iz linija).")
