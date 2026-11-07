"""Lanac dokumenata STARE sveske: OTK -> OTP -> ZBR -> PRJ.  SAMO CITA.

    python tools/migracija/migracija.py lanac --izvor "C:/putanja/stara.xlsm"

Ulaz za ispravke u staroj svesci PRE migracije (odluka operatera 07.11.2026).
Alat nista ne menja i nista ne bira umesto operatera: daje spisak sa dokazima.

1) PRIJEMNICA NA POGRESNOJ ZBIRNOJ (greska unosa u staroj formi)

   frmDokumenta (v2.2.3 - v2.28.4) posle "Unesi prijemnicu" nije praznila polje
   broja zbirne, a promena vozaca ga nije dirala. Pisac do v2.21.0 nije
   proveravao nista osim da broj nije prazan; od v2.21.0 samo da zbirna postoji
   i meko pitanje "1 zbirna = 1 prijemnica". SLEDECA prijemnica je zato
   nasledjivala zbirnu prethodne. Trag je uvek par:

     ZBR_VISE_PRIJEMNICA   aktivna zbirna sa dve (ili vise) aktivne prijemnice
     ZBR_BEZ_PRIJEMNICE    aktivna zbirna bez ijedne -- uvek alarm, jer sva roba
                           mora negde da ode

   Pogresan je KASNIJI unos (veci redni broj PrijemnicaID-a). Predlog je zbirna
   koja nema aktivnu prijemnicu, istog vozaca, kupca, vrste i sorte, u +-N dana,
   najbliza po kg pa po gajbama. Dve varijante, ista greska:
     * druga stanica: vozac promenjen, zbirna nije -> prefiks broja zbirne se ne
       slaze sa stanicom (PRJ_VOZAC_NIJE_VOZAC_ZBIRNE);
     * ista stanica (druga tura, sledeci dan): vozac i zbirna se slazu, pa je
       trag SAMO par gore. Migracija takvu prijemnicu veze na pogresnu zbirnu
       bez upozorenja (cilj._prijemnica_roditelj nadje tacno jednog kandidata).

2) STORNO NIJE PRESAO LANAC

   Aktivan dokument vezan za storniran dokument susednog nivoa (OTK_*, OTP_*,
   ZBR_*). Veze su iste kao u migraciji: otkup -> otpremnica po
   tblOtkup.OtpremnicaID, otpremnica -> zbirna po BrojZbirne (cilj._nadji_zbirnu,
   MIG-11), prijemnica -> zbirna po BrojZbirne.

   Napomena za OTP_SA_STORNIRANIM_OTKUPOM: cilj._clanstvo pravi clanstvo iz
   tblOtkup.OtpremnicaID i za STORNIRAN otkup, pa je u novoj svesci on aktivan
   clan aktivne otpremnice. main to stanje sam nikad ne pravi (StornoOtkup
   odbija storno izvora aktivne otpremnice), a cita ga: "U blokovima" ga sabira
   (VrednostIzvoraPoOtpremnici), a specifikacija pada ("Izvor otpremnice ... je
   storniran", modPrint). Ovde se samo meri; pravilo prenosa je odluka.

Dokument je red, osim kad vise redova nosi isti GeneracijaID i isti poslovni
broj (stari pisac: red po klasi) -- tada su ti redovi jedan dokument.
"""

import re
from collections import defaultdict
from decimal import Decimal

from citac import broj, datum, je_storno, kljuc, polje, tabela, tekst
from cilj import _nadji_zbirnu as nadji_zbirnu  # isto pravilo kao clanstvo migracije (MIG-11)

ALARM = "ALARM"
PROVERI = "PROVERI"

SPEC = {
    # tip: (tabela, PK, poslovni broj, kg, gajbe)
    "OTK": ("tblOtkup", "OtkupID", "BrojDokumenta", "Kolicina", "KolAmbalaze"),
    "OTP": ("tblOtpremnica", "OtpremnicaID", "BrojOtpremnice", "Kolicina", "KolAmbalaze"),
    "ZBR": ("tblZbirna", "ZbirnaID", "BrojZbirne", "UkupnoKolicina", "UkupnoAmbalaze"),
    "PRJ": ("tblPrijemnica", "PrijemnicaID", "BrojPrijemnice", "Kolicina", "KolAmbalaze"),
}

DANA = 3            # predlog zbirne: +- dana od datuma prijemnice
VISOKA_KG = 0.02    # relativna razlika kg do koje je predlog VISOKE pouzdanosti
SREDNJA_KG = 0.05


class Dok:
    """Logicki dokument stare sveske (jedan ili vise redova istog GeneracijaID-a)."""

    def __init__(self, tip, redovi):
        _, pk, br, kg, gaj = SPEC[tip]
        self.tip = tip
        self.redovi = redovi
        prvi = redovi[0]
        self.ids = [tekst(polje(r, pk)) for r in redovi]
        self.id = self.ids[0]
        self.broj = tekst(polje(prvi, br))
        self.datum = datum(polje(prvi, "Datum"))
        self.vozac = tekst(polje(prvi, "VozacID"))
        self.kupac = tekst(polje(prvi, "KupacID"))
        self.stanica = tekst(polje(prvi, "StanicaID"))
        self.vrsta = tekst(polje(prvi, "VrstaVoca"))
        self.sorta = tekst(polje(prvi, "SortaVoca"))
        self.broj_zbirne = tekst(polje(prvi, "BrojZbirne"))
        self.otpremnica_id = tekst(polje(prvi, "OtpremnicaID"))
        self.created = tekst(polje(prvi, "CreatedAt"))
        aktivni = [r for r in redovi if not je_storno(r)]
        self.aktivan = bool(aktivni)
        osnova = aktivni or redovi
        self.kg = sum((broj(polje(r, kg)) or Decimal(0) for r in osnova), Decimal(0))
        self.gajbe = sum((broj(polje(r, gaj)) or Decimal(0) for r in osnova), Decimal(0))
        self.redni = min(_redni(polje(r, pk), r.get("_redni", 0)) for r in redovi)

    def opis(self):
        return "%s %s, %s, %s kg, %s gajbi%s" % (
            self.id, self.broj, _dan(self.datum), _kg(self.kg), _kg(self.gajbe),
            "" if self.aktivan else ", STORNIRAN")


def _redni(idn, rezerva):
    m = re.search(r"(\d+)\s*$", tekst(idn))
    return int(m.group(1)) if m else int(rezerva or 0)


def _dan(d):
    return d.strftime("%d.%m.%Y") if d else ""


def _kg(v):
    return format(round(v, 2).normalize(), "f") if v is not None else ""


def prefiks_broja(broj_zbirne):
    """Broj stanice/vozaca iz broja zbirne: "S3/030726" -> 3, "5/250626-2" -> 5."""
    m = re.match(r"\s*[A-Za-z]*\s*(\d+)\s*/", tekst(broj_zbirne))
    return int(m.group(1)) if m else None


def broj_vozaca(vozac_id):
    """"ST-00002" -> 2; vozac bez brojnog sufiksa -> None."""
    m = re.match(r"\s*[A-Za-z]+-0*(\d+)\s*$", tekst(vozac_id))
    return int(m.group(1)) if m else None


def _isto(a, b):
    """Prazno polje ne protivreci (stara sveska ga nekad nema)."""
    return not tekst(a) or not tekst(b) or kljuc(a) == kljuc(b)


def _dokumenti(tabele, tip):
    t = tabela(tabele, SPEC[tip][0])
    if t is None:
        return [], {}
    pk, br = SPEC[tip][1], SPEC[tip][2]
    grupe, red = defaultdict(list), []
    for r in t.redovi:
        # Isti GeneracijaID I isti poslovni broj: klase jednog dokumenta. Sama
        # generacija nije dovoljna -- dve prijemnice razlicitog broja ostaju dva
        # dokumenta, inace bi se bas greska koju trazimo sakrila u jedan.
        gen = tekst(polje(r, "GeneracijaID"))
        k = ("G", kljuc(gen), kljuc(polje(r, br))) if gen else ("ID", kljuc(polje(r, pk)), r.get("_redni"))
        if k not in grupe:
            red.append(k)
        grupe[k].append(r)
    doks = [Dok(tip, grupe[k]) for k in red]
    po_id = {}
    for d in doks:
        for i in d.ids:
            po_id[kljuc(i)] = d
    return doks, po_id


class Lanac:
    """Dokumenti, veze i nalazi. Nalaz: dict(tezina, pravilo, tip, dok, opis, veze)."""

    def __init__(self, tabele, dana=DANA):
        self.dana = dana
        self.otk, otk_id = _dokumenti(tabele, "OTK")
        self.otp, self.otp_id = _dokumenti(tabele, "OTP")
        self.zbr, self.zbr_id = _dokumenti(tabele, "ZBR")
        self.prj, _ = _dokumenti(tabele, "PRJ")
        self.nalazi = []
        self.prijemnice = []          # predlozi za PRJ na pogresnoj zbirnoj
        self.bez_prijemnice = []      # aktivne zbirne bez aktivne prijemnice
        self._veze(tabele)

    def nalaz(self, tezina, pravilo, dok, opis, veze=()):
        self.nalazi.append(dict(tezina=tezina, pravilo=pravilo, tip=dok.tip, dok=dok, opis=opis,
                                veze=", ".join(v.id for v in veze)))

    # -- veze ---------------------------------------------------------------

    def _veze(self, tabele):
        self.otp_od_otk, self.otk_u_otp = {}, defaultdict(list)
        for o in self.otk:
            if o.otpremnica_id:
                p = self.otp_id.get(kljuc(o.otpremnica_id))
                self.otp_od_otk[o] = p
                if p is not None:
                    self.otk_u_otp[p].append(o)

        zbr_t = tabela(tabele, "tblZbirna")
        zbirne_redovi = list(zbr_t.redovi) if zbr_t is not None else []
        self.zbr_od_otp, self.otp_u_zbr, self.zbr_otp_zasto = {}, defaultdict(list), {}
        for p in self.otp:
            if not p.broj_zbirne:
                continue
            red, kako = nadji_zbirnu(p.broj_zbirne, p.redovi[0], zbirne_redovi)
            z = self.zbr_id.get(kljuc(polje(red, "ZbirnaID"))) if red is not None else None
            self.zbr_od_otp[p] = z
            self.zbr_otp_zasto[p] = kako
            if z is not None:
                self.otp_u_zbr[z].append(p)

        po_broju = defaultdict(list)
        for z in self.zbr:
            po_broju[kljuc(z.broj)].append(z)
        self.zbr_od_prj, self.prj_u_zbr, self.zbr_prj_zasto = {}, defaultdict(list), {}
        for pr in self.prj:
            if not pr.broj_zbirne:
                continue
            z, zasto = self._zbirna_prijemnice(pr, po_broju.get(kljuc(pr.broj_zbirne), []))
            self.zbr_od_prj[pr] = z
            self.zbr_prj_zasto[pr] = zasto
            if z is not None:
                self.prj_u_zbr[z].append(pr)

    @staticmethod
    def _zbirna_prijemnice(pr, kand):
        """Aktivna zbirna tog broja; tek kad aktivne nema, stornirana. Broj koji
        nosi vise zbirnih razresava vlasnik (vozac + kupac), kao pisac prijemnice."""
        if not kand:
            return None, "broj zbirne ne postoji"
        for skup in ([z for z in kand if z.aktivan], kand):
            if not skup:
                continue
            if len(skup) == 1:
                return skup[0], ""
            vl = [z for z in skup if kljuc(z.vozac) == kljuc(pr.vozac) and _isto(z.kupac, pr.kupac)]
            if len(vl) == 1:
                return vl[0], ""
            return None, "broj nosi %d zbirnih, vlasnik nije jednoznacan" % len(skup)
        return None, "broj zbirne ne postoji"

    # -- provere ------------------------------------------------------------

    def proveri(self):
        self._prijemnice_zbirnih()
        self._otkupi()
        self._otpremnice()
        self._zbirne()
        return self

    def _akt_sto(self, lista):
        return [d for d in lista if d.aktivan], [d for d in lista if not d.aktivan]

    def _prijemnice_zbirnih(self):
        prazne = []
        for z in self.zbr:
            if not z.aktivan:
                continue
            akt, sto = self._akt_sto(self.prj_u_zbr.get(z, []))
            if not akt:
                prazne.append(z)
                if not sto:
                    self.nalaz(ALARM, "ZBR_BEZ_PRIJEMNICE", z,
                               "%s: aktivna zbirna bez ijedne prijemnice -- roba nije nigde stigla"
                               % z.opis())
            elif len(akt) > 1:
                akt.sort(key=lambda d: d.redni)
                self.nalaz(ALARM, "ZBR_VISE_PRIJEMNICA", z,
                           "%s: %d aktivne prijemnice (%s) -- pravilo je 1 zbirna = 1 prijemnica"
                           % (z.opis(), len(akt), "; ".join(p.opis() for p in akt)), akt)
        for pr in self.prj:
            if not pr.aktivan:
                continue
            z = self.zbr_od_prj.get(pr)
            if not pr.broj_zbirne:
                self.nalaz(ALARM, "PRJ_BEZ_BROJA_ZBIRNE", pr, "%s: prijemnica bez broja zbirne" % pr.opis())
            elif z is None:
                self.nalaz(ALARM, "PRJ_ZBIRNA_NIJE_NADJENA", pr, "%s: BrojZbirne %s -- %s"
                           % (pr.opis(), pr.broj_zbirne, self.zbr_prj_zasto.get(pr, "")))
            elif not z.aktivan:
                self.nalaz(ALARM, "PRJ_ZBIRNA_STORNIRANA", pr, "%s: zbirna %s je stornirana"
                           % (pr.opis(), z.opis()), [z])
            elif pr.vozac and z.vozac and kljuc(pr.vozac) != kljuc(z.vozac):
                self.nalaz(ALARM, "PRJ_VOZAC_NIJE_VOZAC_ZBIRNE", pr,
                           "%s: vozac prijemnice %s, vozac zbirne %s %s -- migracija joj ne nalazi "
                           "roditelja (PRIJEMNICA_BEZ_JEDNOZNACNE_ZBIRNE)"
                           % (pr.opis(), pr.vozac, z.broj, z.vozac), [z])

        predlozi = defaultdict(list)
        for z in self.zbr:
            if not z.aktivan:
                continue
            akt = sorted(self._akt_sto(self.prj_u_zbr.get(z, []))[0], key=lambda d: d.redni)
            if len(akt) < 2:
                continue
            ostaje = akt[0]
            for pr in akt[1:]:
                red = self._predlog(pr, z, ostaje, [x for x in prazne if x is not z])
                self.prijemnice.append(red)
                if red["predlog"] is not None:
                    predlozi[red["predlog"]].append(pr)
        for red in self.prijemnice:
            p = red["predlog"]
            if p is not None and len(predlozi[p]) > 1:
                drugi = [x.id for x in predlozi[p] if x is not red["prj"]]
                red["signali"].append("ISTI PREDLOG kao za %s -- proveri rucno" % ", ".join(drugi))
                red["pouzdanost"] = "NISKA"
        za = {p: [x.id for x in lst] for p, lst in predlozi.items()}
        for z in prazne:
            self.bez_prijemnice.append(dict(zbr=z, otp=self.otp_u_zbr.get(z, []),
                                            stornirane_prj=self.prj_u_zbr.get(z, []),
                                            predlozena_za=za.get(z, [])))

    def _predlog(self, pr, sada, ostaje, prazne):
        d0 = pr.datum or sada.datum
        kand = []
        for z in prazne:
            if not _isto(z.kupac, pr.kupac) or not _isto(z.vrsta, pr.vrsta) or not _isto(z.sorta, pr.sorta):
                continue
            if not z.datum or not d0 or abs((z.datum - d0).days) > self.dana:
                continue
            nivo = 0 if kljuc(z.vozac) == kljuc(pr.vozac) else 1
            rel = float(abs(z.kg - pr.kg) / z.kg) if z.kg else 9.0
            kand.append((nivo, rel, abs(z.gajbe - pr.gajbe), abs((z.datum - d0).days), z.redni, z))
        kand.sort(key=lambda k: k[:5])

        signali = ["kasniji unos (%s posle %s)" % (pr.id, ostaje.id)]
        if pr.vozac and sada.vozac and kljuc(pr.vozac) != kljuc(sada.vozac):
            signali.append("vozac prijemnice %s nije vozac zbirne %s" % (pr.vozac, sada.vozac))
        pb, bv = prefiks_broja(sada.broj), broj_vozaca(pr.vozac)
        if pb is not None and bv is not None and pb != bv:
            signali.append("prefiks %s ne odgovara vozacu %s" % (sada.broj.split("/")[0], pr.vozac))
        if sada.kg:
            r_ost, r_pr = abs(sada.kg - ostaje.kg) / sada.kg, abs(sada.kg - pr.kg) / sada.kg
            if r_pr < r_ost:
                signali.append("PAZNJA: po kg zbirnoj bolje odgovara ovaj unos (%s%%) nego %s (%s%%)"
                               % (_kg(r_pr * 100), ostaje.id, _kg(r_ost * 100)))

        if not kand:
            pouzdanost, predlog = "NEMA KANDIDATA", None
        else:
            prvi = kand[0]
            drugi = kand[1] if len(kand) > 1 else None
            jasan = drugi is None or drugi[0] > prvi[0] or drugi[1] - prvi[1] >= VISOKA_KG
            if prvi[0] == 0 and prvi[1] <= VISOKA_KG and jasan:
                pouzdanost = "VISOKA"
            elif prvi[1] <= SREDNJA_KG and jasan:
                pouzdanost = "SREDNJA"
            else:
                pouzdanost = "NISKA"
            predlog = prvi[5]
            if prvi[0] != 0:
                signali.append("predlog je drugog vozaca (%s)" % predlog.vozac)
            if prvi[2] == 0:
                signali.append("gajbe se slazu")
        return dict(prj=pr, sada=sada, ostaje=ostaje, predlog=predlog, pouzdanost=pouzdanost,
                    signali=signali,
                    ostali=["%s %s (%s, %s kg, %s gajbi)" % (k[5].broj, k[5].vozac, _dan(k[5].datum),
                                                            _kg(k[5].kg), _kg(k[5].gajbe))
                            for k in kand[1:4]])

    def _otkupi(self):
        for o in self.otk:
            if not o.aktivan or not o.otpremnica_id:
                continue
            p = self.otp_od_otk.get(o)
            if p is None:
                self.nalaz(PROVERI, "OTK_OTPREMNICA_NE_POSTOJI", o, "%s: OtpremnicaID %s ne postoji"
                           % (o.opis(), o.otpremnica_id))
                continue
            if not p.aktivan:
                self.nalaz(ALARM, "OTK_U_STORNIRANOJ_OTPREMNICI", o, "%s: otpremnica %s" % (o.opis(), p.opis()), [p])
                continue
            z = self.zbr_od_otp.get(p)
            if z is None:
                continue
            if not z.aktivan:
                self.nalaz(ALARM, "OTK_U_STORNIRANOJ_ZBIRNOJ", o, "%s: otpremnica %s, zbirna %s"
                           % (o.opis(), p.broj, z.opis()), [p, z])
                continue
            akt, sto = self._akt_sto(self.prj_u_zbr.get(z, []))
            if sto and not akt:
                self.nalaz(ALARM, "OTK_PRIJEMNICA_STORNIRANA", o, "%s: zbirna %s ima samo stornirane prijemnice (%s)"
                           % (o.opis(), z.broj, ", ".join(x.id for x in sto)), [p, z] + sto)

    def _otpremnice(self):
        for p in self.otp:
            if not p.aktivan:
                continue
            akt, sto = self._akt_sto(self.otk_u_otp.get(p, []))
            if sto:
                self.nalaz(ALARM, "OTP_SA_STORNIRANIM_OTKUPOM", p,
                           "%s: stornirani blokovi %s (%s kg); aktivnih %d (%s kg) -- u novoj svesci "
                           "\"U blokovima\" ih sabira, a specifikacija pada"
                           % (p.opis(), ", ".join(x.id for x in sto), _kg(sum((x.kg for x in sto), Decimal(0))),
                              len(akt), _kg(sum((x.kg for x in akt), Decimal(0)))), sto)
            if not p.broj_zbirne:
                continue
            z = self.zbr_od_otp.get(p)
            if z is None:
                self.nalaz(PROVERI, "OTP_ZBIRNA_NIJE_NADJENA", p, "%s: BrojZbirne %s -- %s"
                           % (p.opis(), p.broj_zbirne, self.zbr_otp_zasto.get(p, "")))
            elif not z.aktivan:
                self.nalaz(ALARM, "OTP_U_STORNIRANOJ_ZBIRNOJ", p, "%s: zbirna %s" % (p.opis(), z.opis()), [z])
            else:
                pa, ps = self._akt_sto(self.prj_u_zbr.get(z, []))
                if ps and not pa:
                    self.nalaz(ALARM, "OTP_PRIJEMNICA_STORNIRANA", p,
                               "%s: zbirna %s ima samo stornirane prijemnice (%s)"
                               % (p.opis(), z.broj, ", ".join(x.id for x in ps)), [z] + ps)

    def _zbirne(self):
        for z in self.zbr:
            if not z.aktivan:
                continue
            otps = self.otp_u_zbr.get(z, [])
            akt, sto = self._akt_sto(otps)
            prj_akt, prj_sto = self._akt_sto(self.prj_u_zbr.get(z, []))
            if otps and not akt:
                paznja = ""
                if prj_akt:
                    paznja = " -- PAZNJA: aktivna prijemnica %s; storno zbirne bi je ostavio bez zbirne" \
                             % ", ".join(x.id for x in prj_akt)
                self.nalaz(ALARM, "ZBR_SVE_OTPREMNICE_STORNIRANE", z, "%s: otpremnice %s su stornirane%s"
                           % (z.opis(), ", ".join(x.broj or x.id for x in sto), paznja), sto + prj_akt)
            elif sto:
                self.nalaz(PROVERI, "ZBR_SA_STORNIRANOM_OTPREMNICOM", z, "%s: stornirane %s, aktivne %s"
                           % (z.opis(), ", ".join(x.broj or x.id for x in sto),
                              ", ".join(x.broj or x.id for x in akt)), sto)
            elif not otps:
                self.nalaz(PROVERI, "ZBR_BEZ_OTPREMNICE", z, "%s: nijedna otpremnica ne nosi ovaj broj zbirne"
                           % z.opis())
            sto_otk = [o for p in akt for o in self.otk_u_otp.get(p, []) if not o.aktivan]
            if sto_otk:
                self.nalaz(PROVERI, "ZBR_SA_STORNIRANIM_OTKUPOM", z,
                           "%s: stornirani blokovi u aktivnim otpremnicama: %s (v. OTP_SA_STORNIRANIM_OTKUPOM)"
                           % (z.opis(), ", ".join(o.id for o in sto_otk)), sto_otk)
            if prj_sto and not prj_akt:
                self.nalaz(ALARM, "ZBR_SAMO_STORNIRANE_PRIJEMNICE", z, "%s: prijemnice %s su stornirane, aktivne nema"
                           % (z.opis(), ", ".join(x.id for x in prj_sto)), prj_sto)


# -- izvestaj ---------------------------------------------------------------

LEGENDA = [
    ("ALARM", "stanje koje ne sme da ude u novu svesku bez odluke"),
    ("PROVERI", "moze biti ispravno; pogledati dokument"),
    ("Prijemnice", "kasniji unos na zbirnoj sa vise prijemnica + predlog prave zbirne"),
    ("ZbirneBezPrijemnice", "aktivne zbirne bez aktivne prijemnice -- svaka mora negde da ode"),
    ("Pouzdanost VISOKA", "isti vozac, kg do %d%%, drugi kandidat jasno losiji" % int(VISOKA_KG * 100)),
    ("Pouzdanost SREDNJA", "kg do %d%%, drugi kandidat jasno losiji" % int(SREDNJA_KG * 100)),
    ("Pouzdanost NISKA", "postoji kandidat, ali dokaz nije jednoznacan -- odlucuje operater"),
]


def redovi_prijemnica(lanac):
    out = []
    for r in lanac.prijemnice:
        p, s, o, z = r["prj"], r["sada"], r["ostaje"], r["predlog"]
        out.append([r["pouzdanost"], p.id, p.broj, p.datum, p.vozac, p.kg, p.gajbe,
                    s.broj, s.vozac, s.kg, s.gajbe, "%s (%s kg, %s gajbi)" % (o.id, _kg(o.kg), _kg(o.gajbe)),
                    z.broj if z else "", z.id if z else "", z.datum if z else "", z.vozac if z else "",
                    z.kg if z else "", z.gajbe if z else "",
                    (z.kg - p.kg) if z else "", (z.gajbe - p.gajbe) if z else "",
                    "; ".join(r["signali"]), "; ".join(r["ostali"])])
    return out


KOLONE_PRIJEMNICA = ["Pouzdanost", "PrijemnicaID", "BrojPrijemnice", "Datum", "VozacID", "Kg", "Gajbe",
                     "SadaBrojZbirne", "SadaVozacZbirne", "SadaKgZbirne", "SadaGajbeZbirne", "OstajeNaZbirnoj",
                     "PredlogBrojZbirne", "PredlogZbirnaID", "PredlogDatum", "PredlogVozac", "PredlogKg",
                     "PredlogGajbe", "RazlikaKg", "RazlikaGajbe", "Signali", "OstaliKandidati"]

KOLONE_BEZ = ["ZbirnaID", "BrojZbirne", "Datum", "VozacID", "KupacID", "Sorta", "Kg", "Gajbe",
              "Otpremnice", "StorniranePrijemnice", "PredlozenaZa"]

KOLONE_NALAZI = ["Tezina", "Pravilo", "Dokument", "ID", "Broj", "Datum", "Opis", "Veze"]


def redovi_bez(lanac):
    return [[b["zbr"].id, b["zbr"].broj, b["zbr"].datum, b["zbr"].vozac, b["zbr"].kupac, b["zbr"].sorta,
             b["zbr"].kg, b["zbr"].gajbe,
             ", ".join("%s%s" % (p.broj or p.id, "" if p.aktivan else " (storn.)") for p in b["otp"]),
             ", ".join(x.id for x in b["stornirane_prj"]), ", ".join(b["predlozena_za"])]
            for b in lanac.bez_prijemnice]


def redovi_nalaza(lanac):
    red = {ALARM: 0, PROVERI: 1}
    return [[n["tezina"], n["pravilo"], n["tip"], n["dok"].id, n["dok"].broj, n["dok"].datum, n["opis"], n["veze"]]
            for n in sorted(lanac.nalazi, key=lambda n: (red.get(n["tezina"], 9), n["pravilo"], n["dok"].redni))]


def sazetak(lanac):
    from collections import Counter
    c = Counter((n["tezina"], n["pravilo"]) for n in lanac.nalazi)
    poudz = Counter(r["pouzdanost"] for r in lanac.prijemnice)
    out = [("Dokumenata: otkup / otpremnica / zbirna / prijemnica",
            "%d / %d / %d / %d" % (len(lanac.otk), len(lanac.otp), len(lanac.zbr), len(lanac.prj))),
           ("Prijemnice na pogresnoj zbirnoj (kasniji unos)", len(lanac.prijemnice))]
    for k in ("VISOKA", "SREDNJA", "NISKA", "NEMA KANDIDATA"):
        if poudz.get(k):
            out.append(("  pouzdanost %s" % k, poudz[k]))
    out.append(("Aktivne zbirne bez aktivne prijemnice", len(lanac.bez_prijemnice)))
    out.append(("", ""))
    for (t, p), n in sorted(c.items()):
        out.append(("%s  %s" % (t, p), n))
    return out
