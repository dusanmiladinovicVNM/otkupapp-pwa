"""Stara knjiga ambalaze -> AMB-10 kretanja. U MEMORIJI; nista se ne pise.

Stari model (do 4d3e2caa) je knjiga NOGU: svaki red je jedan entitet (Kooperant,
Stanica, Kupac) sa smerom Ulaz/Izlaz, a vozac je zapisan samo kao ogledalo u
koloni VozacID. Saldo = Ulaz - Izlaz po entitetu. AMB-10 (main) je knjiga
KRETANJA: jedan red = Od -> Na, sa vrstom kretanja (docs/DOMEN/AMBALAZA.md).

Pravila su ista kao u lancu MIG-04 .. MIG-19 (grana chore/import-migration-tools),
samo sazeta u jedan prolaz i bez SHA zakljucavanja:

  stari DokumentTip    noge (entitet, smer)              -> kretanje            vrsta
  Otkup                Kooperant Izlaz + Stanica Ulaz       K -> S              AMBALAZA_UZ_ROBU
  OM-Izlaz-Koop        Kooperant Ulaz + Stanica Izlaz       S -> K              IZDATA_PRAZNA
                       (DokumentID = OtkupID -> uz otkup, inace revers)
  OM-Ulaz-Koop         Kooperant Izlaz + Stanica Ulaz       K -> S  (revers)    POVRAT_PRAZNE
  OM-Izlaz-Firma       Stanica Izlaz + VozacID              S -> V  (revers)    PRENOS_INTERNO
  OM-Ulaz-Firma        Stanica Ulaz + VozacID               V -> S  (revers)    PRENOS_INTERNO
  Otpremnica           Stanica Izlaz + VozacID              S -> V              AMBALAZA_UZ_ROBU
  Prijemnica           Kupac Ulaz + VozacID                 V -> Kupac          AMBALAZA_UZ_ROBU
                       Kupac Izlaz + VozacID                Kupac -> V          POVRAT_PRAZNE
  Kupci-Otpremnica     Kupac Izlaz + VozacID                Kupac -> V (revers  POVRAT_PRAZNE
                                                            partnera; NOVO, MIG ga nije imao)
  OMUlaz i sve ostalo  -> ZA ODLUKU (ili izuzetak u izuzeci.json)

Stornirani redovi idu u arhivu, ne u knjigu (MIG-05.7/05.8). Aktivna noga
storniranog dokumenta ide ZA ODLUKU (slucaj AMB-00700, MIG-19.12).

Istorija se prevodi VERNO: negativna stara salda ostaju negativna. Klijent nije
unosio pocetno stanje stanica, a tudja ambalaza tada nije postojala kao pojam
(odluka operatera 28.10.2026), pa se u istoriju nista ne dopisuje. Prelaz u
novu sezonu (main ne pusta realan nalog ispod nule, AMB-INV-07) je zaseban
dogadjaj na dan preseka -- pregled ga samo predlaze, iz salda na kraju.

Rezim ISPRAVNO (podrazumevan u komandi) dodaje ono sto main trazi za dalji
rad (odluka operatera 28.10.2026): kooperant -> politika MIG-16/17 po dnevnom
zakljucku (rast duga -> ULAZ_TUDJE na njegovom izlaznom dokumentu, pad duga ->
izdate prazne postaju VRACANJE_TUDJE); kupac -> ULAZ_TUDJE u trenutku manjka;
stanica/vozac -> jedna NABAVKA na prvi dan rada. Sve dodato nosi CreatedBy
MIGRACIJA i vidi se u pregledu kao zasebna kolona.

Svaki stari red dobija TACNO JEDAN ishod: KRETANJE, STORNO, IZUZETO ili
ZA_ODLUKU. To je osnovna tvrdnja i proverava je `proveri_pokrivenost`.
"""

import json
import uuid
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from decimal import Decimal

from citac import broj, datum, je_storno, kljuc, polje, tabela, tekst

GRANICA = "SpoljniSvet"
PARTNERI = ("Kooperant", "Kupac")
SOPSTVENI = ("Stanica", "Vozac", "Firma")
REALNI = PARTNERI + SOPSTVENI

ISHOD_KRETANJE = "KRETANJE"
ISHOD_STORNO = "STORNO"
ISHOD_IZUZETO = "IZUZETO"
ISHOD_ZA_ODLUKU = "ZA_ODLUKU"

UZ_ROBU = "AMBALAZA_UZ_ROBU"
IZDATA = "IZDATA_PRAZNA"
POVRAT = "POVRAT_PRAZNE"
PRENOS = "PRENOS_INTERNO"
ULAZ_TUDJE = "ULAZ_TUDJE_AMBALAZE"
VRACANJE_TUDJE = "VRACANJE_TUDJE_AMBALAZE"
NABAVKA = "NABAVKA"
VRSTE = {UZ_ROBU, IZDATA, POVRAT, PRENOS, ULAZ_TUDJE, VRACANJE_TUDJE, NABAVKA,
         "OTPIS", "POZAJMICA_PARTNERA"}

DOK_AMB = "AmbalazaDokument"

# verno:    istorija tacno kao u staroj svesci, negativna salda ostaju.
# ispravno: ista istorija + ono sto main trazi da bi sveska bila spremna za rad
#           (nijedan realan nalog ni jednog dana ispod nule, AMB-INV-07):
#           kooperant -> ULAZ_TUDJE / VRACANJE_TUDJE po danu (MIG-16/17),
#           kupac     -> ULAZ_TUDJE u trenutku manjka (kao zivi pisac),
#           stanica/vozac -> jedna NABAVKA na prvi dan rada (pocetno stanje).
VERNO = "verno"
ISPRAVNO = "ispravno"

MATICNE = {
    "Kooperant": ("tblKooperanti", "KooperantID"),
    "Stanica": ("tblStanice", "StanicaID"),
    "Kupac": ("tblKupci", "KupacID"),
    "Vozac": ("tblVozaci", "VozacID"),
}
RODITELJ = {
    "otkup": ("tblOtkup", "OtkupID", "Otkup", "BrojDokumenta"),
    "otpremnica": ("tblOtpremnica", "OtpremnicaID", "Otpremnica", "BrojOtpremnice"),
    "prijemnica": ("tblPrijemnica", "PrijemnicaID", "Prijemnica", "BrojPrijemnice"),
}

# Istorijsko znacenje starog tipa dokumenta (MIG-19.11), za kolonu Opis.
OPIS = {
    "otkup": "Kooperant predaje pune na stanicu",
    "om-izlaz-koop": "Stanica izdaje prazne kooperantu",
    "om-ulaz-koop": "Kooperant vraca prazne stanici",
    "otpremnica": "Stanica otprema ambalazu sa vozacem",
    "prijemnica": "Kupac prima pune (+) ili vraca prazne (-)",
    "kupci-otpremnica": "Kupac vraca prazne vozacu",
    "om-izlaz-firma": "Stanica predaje ambalazu vozacu",
    "om-ulaz-firma": "Stanica prima ambalazu od vozaca",
    "omulaz": "Jednostrani prijem na stanici",
}

# Fiksni prostor imena za deterministicke ID-jeve migracije: isti ulaz daje
# isti ID, pa ponovljen prolaz ne pravi nove redove.
NS_MIGRACIJA = uuid.UUID("5d6f4b0e-6c1e-5b8a-9a43-0d7c2f9e8a11")


def novi_id(prefiks, *delovi):
    """Isti oblik kao modDataAccess.NewEntityID: prefiks + 32 hex znaka."""
    return prefiks + uuid.uuid5(NS_MIGRACIJA, "|".join(str(d) for d in delovi)).hex.upper()


def norm_tip(v):
    """'0,4 kg' / '0.4KG' -> '0,4KG' (MIG-19.3 norm_type)."""
    import re
    t = re.sub(r"\s+", "", tekst(v).upper())
    return re.sub(r"(?<=\d)\.(?=\d)", ",", t)


def _kanon_nalog(v):
    k = kljuc(v)
    for ime in REALNI + (GRANICA,):
        if k == ime.casefold():
            return ime
    return tekst(v)


@dataclass
class Noga:
    redni: int
    excel_red: int
    amb_id: str
    datum: object
    datum_txt: str
    tip: str
    kolicina: object
    smer: str
    ent_tip: str
    ent_id: str
    vozac: str
    dok_tip: str
    dok_id: str
    storno: bool
    created_at: str = ""
    created_by: str = ""
    ishod: str = ""
    razlog: str = ""
    kretanje: str = ""

    @property
    def dt(self):
        return kljuc(self.dok_tip)

    @property
    def uloga(self):
        return (self.ent_tip, self.smer)

    def doprinos(self):
        """Doprinos staroj knjizi: [((tip naloga, id, tip ambalaze), +/-kolicina)].

        Isto pravilo kao stari saldo, ukljucujuci izvedeno ogledalo vozaca za
        noge Stanice i Kupca koje nose VozacID (MIG-04.1 / MIG-04.2).
        """
        if self.storno or self.kolicina is None or self.smer not in ("ulaz", "izlaz"):
            return []
        q = self.kolicina if self.smer == "ulaz" else -self.kolicina
        out = []
        if self.ent_tip and self.ent_id:
            out.append(((self.ent_tip, self.ent_id, self.tip), q))
        if self.ent_tip in ("Stanica", "Kupac") and self.vozac:
            out.append((("Vozac", self.vozac, self.tip), -q))
        return out


@dataclass
class Kretanje:
    amb_id: str
    datum: object
    tip: str
    kolicina: Decimal
    dok_tip: str
    dok_id: str
    od_tip: str
    od_id: str
    na_tip: str
    na_id: str
    vrsta: str
    pravilo: str
    stari_tip: str = ""
    stari_broj: str = ""
    noge: tuple = ()
    stari_amb: tuple = ()
    redosled: int = 0
    created_at: str = ""
    created_by: str = ""
    generisano: bool = False
    napomena: str = ""

    def doprinos(self):
        out = []
        if self.od_tip != GRANICA:
            out.append(((self.od_tip, self.od_id, self.tip), -self.kolicina))
        if self.na_tip != GRANICA:
            out.append(((self.na_tip, self.na_id, self.tip), self.kolicina))
        return out

    def par(self):
        return frozenset({(self.od_tip, kljuc(self.od_id)), (self.na_tip, kljuc(self.na_id))})


@dataclass
class Rezultat:
    oblik: str = ""
    noge: list = field(default_factory=list)
    kretanja: list = field(default_factory=list)
    zaglavlja: dict = field(default_factory=dict)
    za_odluku: list = field(default_factory=list)
    upozorenja: list = field(default_factory=list)
    izuzeci_primenjeni: list = field(default_factory=list)
    nazivi: dict = field(default_factory=dict)
    brojevi_dok: dict = field(default_factory=dict)
    tipovi_ambalaze: set = field(default_factory=set)
    politika: list = field(default_factory=list)
    knjizenje: str = VERNO


class GreskaAmbalaze(Exception):
    pass


# ---------------------------------------------------------------------------
# Ulaz: maticne tabele, roditelji, noge
# ---------------------------------------------------------------------------

def _ucitaj_maticne(tabele):
    ids = {}
    nazivi = {}
    for nalog, (tbl, pk) in MATICNE.items():
        t = tabela(tabele, tbl)
        ids[nalog] = None if t is None else {}
        if t is None:
            continue
        for r in t.redovi:
            i = tekst(polje(r, pk))
            if not i:
                continue
            ids[nalog][kljuc(i)] = i
            if nalog in ("Kooperant", "Vozac"):
                ime = (tekst(polje(r, "Ime")) + " " + tekst(polje(r, "Prezime"))).strip()
            else:
                ime = tekst(polje(r, "Naziv"))
            nazivi[(nalog, kljuc(i))] = ime
    return ids, nazivi


def _ucitaj_roditelje(tabele):
    roditelji = {}
    brojevi = {}
    for dt, (tbl, pk, novi_tip, broj_kol) in RODITELJ.items():
        t = tabela(tabele, tbl)
        roditelji[dt] = {}
        if t is None:
            continue
        for r in t.redovi:
            i = kljuc(polje(r, pk))
            if i:
                roditelji[dt][i] = r
                brojevi[(novi_tip, i)] = tekst(polje(r, broj_kol))
    return roditelji, brojevi


def _noge(t):
    out = []
    for r in t.redovi:
        smer = kljuc(polje(r, "Smer"))
        out.append(Noga(
            redni=r["_redni"], excel_red=r["_excel_red"],
            amb_id=tekst(polje(r, "AmbID")),
            datum=datum(polje(r, "Datum")), datum_txt=tekst(polje(r, "Datum")),
            tip=tekst(polje(r, "TipAmbalaze")), kolicina=broj(polje(r, "Kolicina")),
            smer=smer, ent_tip=_kanon_nalog(polje(r, "EntitetTip")),
            ent_id=tekst(polje(r, "EntitetID")), vozac=tekst(polje(r, "VozacID")),
            dok_tip=tekst(polje(r, "DokumentTip")), dok_id=tekst(polje(r, "DokumentID")),
            storno=je_storno(r),
            created_at=tekst(polje(r, "CreatedAt")), created_by=tekst(polje(r, "CreatedBy"))))
    return out


# ---------------------------------------------------------------------------
# Izuzeci (izuzeci.json) -- rucne odluke, svaka sa potpisom reda
# ---------------------------------------------------------------------------

def ucitaj_izuzetke(putanja):
    if not putanja:
        return []
    with open(putanja, encoding="utf-8") as f:
        podaci = json.load(f)
    lista = podaci.get("izuzeci", [])
    for i, iz in enumerate(lista, 1):
        if not tekst(iz.get("AmbID")):
            raise GreskaAmbalaze("izuzetak #%d nema AmbID" % i)
        if iz.get("radnja") not in ("izuzmi", "kretanje", "za_odluku"):
            raise GreskaAmbalaze("izuzetak %s: nepoznata radnja %r" % (iz["AmbID"], iz.get("radnja")))
        if iz["radnja"] == "kretanje":
            k = iz.get("kretanje") or {}
            if k.get("VrstaKretanja") not in VRSTE or len(k.get("Od", [])) != 2 or len(k.get("Na", [])) != 2:
                raise GreskaAmbalaze("izuzetak %s: kretanje nije potpuno" % iz["AmbID"])
    return lista


def _potpis_se_slaze(n, provera):
    """Izuzetak vazi samo za red koji i dalje izgleda tacno kao kad je odluka doneta."""
    razlike = []
    for kol, ocekivano in (provera or {}).items():
        k = kol.casefold()
        if k == "datum":
            stvarno = n.datum.isoformat() if n.datum else n.datum_txt
        elif k == "tipambalaze":
            stvarno, ocekivano = norm_tip(n.tip), norm_tip(ocekivano)
        elif k == "kolicina":
            stvarno, ocekivano = n.kolicina, broj(ocekivano)
        elif k == "dokumenttip":
            stvarno, ocekivano = n.dt, kljuc(ocekivano)
        elif k == "entitettip":
            stvarno, ocekivano = kljuc(n.ent_tip), kljuc(ocekivano)
        elif k == "entitetid":
            stvarno, ocekivano = kljuc(n.ent_id), kljuc(ocekivano)
        elif k == "smer":
            stvarno, ocekivano = n.smer, kljuc(ocekivano)
        elif k == "dokumentid":
            stvarno, ocekivano = kljuc(n.dok_id), kljuc(ocekivano)
        else:
            razlike.append("nepoznata kolona provere " + kol)
            continue
        if stvarno != ocekivano:
            razlike.append("%s: ocekivano %s, nadjeno %s" % (kol, ocekivano, stvarno))
    return razlike


# ---------------------------------------------------------------------------
# Glavni prolaz
# ---------------------------------------------------------------------------

def prevedi(tabele, izuzeci=(), knjizenje=VERNO, politika_partneri=None):
    t = tabela(tabele, "tblAmbalaza")
    if t is None:
        raise GreskaAmbalaze("Sveska nema tabelu tblAmbalaza")
    rez = Rezultat()
    ids, rez.nazivi = _ucitaj_maticne(tabele)
    roditelji, rez.brojevi_dok = _ucitaj_roditelje(tabele)
    tt = tabela(tabele, "tblTipAmbalaze")
    rez.tipovi_ambalaze = None if tt is None else {
        norm_tip(polje(r, "TipAmbalaze")) for r in tt.redovi if tekst(polje(r, "TipAmbalaze"))}

    if t.ima_kolonu("OdNalogTip"):
        rez.oblik = "AMB10"
        _postojeca_knjiga(t, rez)
        return rez
    if not (t.ima_kolonu("Smer") and t.ima_kolonu("EntitetTip")):
        raise GreskaAmbalaze("tblAmbalaza nije ni stari ni AMB-10 oblik (nema Smer/EntitetTip ni OdNalogTip)")
    rez.oblik = "STARI"
    rez.noge = _noge(t)

    _jedinstveni_amb_id(rez)
    for n in rez.noge:
        if n.storno:
            n.ishod = ISHOD_STORNO
    for n in _aktivne(rez):
        razlozi = _osnovna_provera(n, ids, rez.tipovi_ambalaze)
        if razlozi:
            _za_odluku(rez, [n], ";".join(razlozi))
    _primeni_izuzetke(rez, izuzeci)

    po_tipu = defaultdict(list)
    for n in _aktivne(rez):
        po_tipu[n.dt].append(n)
    _otkup(rez, po_tipu.pop("otkup", []), roditelji)
    _om_koop(rez, po_tipu.pop("om-izlaz-koop", []), roditelji, izlaz=True)
    _om_koop(rez, po_tipu.pop("om-ulaz-koop", []), roditelji, izlaz=False)
    _jednostrane(rez, po_tipu.pop("om-izlaz-firma", []), "om-izlaz-firma", roditelji, ids)
    _jednostrane(rez, po_tipu.pop("om-ulaz-firma", []), "om-ulaz-firma", roditelji, ids)
    _jednostrane(rez, po_tipu.pop("otpremnica", []), "otpremnica", roditelji, ids)
    _jednostrane(rez, po_tipu.pop("prijemnica", []), "prijemnica", roditelji, ids)
    _jednostrane(rez, po_tipu.pop("kupci-otpremnica", []), "kupci-otpremnica", roditelji, ids)
    for dt, noge in po_tipu.items():
        razlog = "OMULAZ_IZVOR_NEPOZNAT" if dt == "omulaz" else "NEPOZNAT_TIP_DOKUMENTA"
        for n in noge:
            _za_odluku(rez, [n], razlog)

    _proveri_efekte(rez)
    rez.knjizenje = knjizenje
    if politika_partneri is None:
        politika_partneri = ("Kooperant",) if knjizenje == ISPRAVNO else ()
    rez.politika = primeni_politiku_duga(rez, politika_partneri)
    if knjizenje == ISPRAVNO:
        rez.politika += pokrice_po_manjku(rez, ("Kupac",))
        rez.politika += nabavka_na_pocetku(rez, SOPSTVENI)
    return rez


def _aktivne(rez):
    return [n for n in rez.noge if not n.ishod]


def _jedinstveni_amb_id(rez):
    c = Counter(kljuc(n.amb_id) for n in rez.noge)
    for n in rez.noge:
        if not n.amb_id or c[kljuc(n.amb_id)] > 1:
            rez.upozorenja.append(dict(vrsta="AMBID_PRAZAN_ILI_DUPLI", redni=n.redni, amb_id=n.amb_id,
                                       opis="Stari AmbID nije jedinstven; za novi ID koristi se redni broj."))


def _izvorni_kljuc(n):
    return n.amb_id if n.amb_id else "R%d" % n.redni


def _osnovna_provera(n, ids, tipovi):
    r = []
    if n.kolicina is None or n.kolicina <= 0:
        r.append("NEISPRAVNA_KOLICINA")
    if n.smer not in ("ulaz", "izlaz"):
        r.append("NEISPRAVAN_SMER")
    if n.datum is None:
        r.append("NEISPRAVAN_DATUM")
    if not n.tip:
        r.append("NEMA_TIPA_AMBALAZE")
    elif tipovi is not None and norm_tip(n.tip) not in tipovi:
        r.append("TIP_AMBALAZE_NIJE_U_MATICNOJ")
    if n.ent_tip not in MATICNE:
        r.append("NEPOZNAT_NALOG")
    elif not n.ent_id:
        r.append("NEMA_NALOGA")
    elif ids.get(n.ent_tip) is not None and kljuc(n.ent_id) not in ids[n.ent_tip]:
        r.append("NALOG_NIJE_U_MATICNOJ")
    if n.vozac and ids.get("Vozac") is not None and kljuc(n.vozac) not in ids["Vozac"]:
        r.append("VOZAC_NIJE_U_MATICNOJ")
    return r


def _za_odluku(rez, noge, razlog, predlog="", napomena=""):
    for n in noge:
        n.ishod = ISHOD_ZA_ODLUKU
        n.razlog = razlog
    rez.za_odluku.append(dict(razlog=razlog, noge=[n.redni for n in noge], predlog=predlog,
                              napomena=napomena))


def _primeni_izuzetke(rez, izuzeci):
    po_id = defaultdict(list)
    for n in rez.noge:
        if n.amb_id:
            po_id[kljuc(n.amb_id)].append(n)
    for iz in izuzeci:
        nadjene = po_id.get(kljuc(iz["AmbID"]), [])
        if len(nadjene) != 1:
            rez.upozorenja.append(dict(vrsta="IZUZETAK_BEZ_REDA", amb_id=iz["AmbID"],
                                       opis="Izuzetak ne pogadja tacno jedan red (nadjeno %d)." % len(nadjene)))
            continue
        n = nadjene[0]
        if n.ishod == ISHOD_STORNO:
            rez.upozorenja.append(dict(vrsta="IZUZETAK_NAD_STORNOM", amb_id=iz["AmbID"], redni=n.redni,
                                       opis="Red je storniran; izuzetak se ne primenjuje."))
            continue
        razlike = _potpis_se_slaze(n, iz.get("provera"))
        if razlike:
            # Red se promenio posle odluke: odluka vise ne vazi, operater bira ponovo.
            if n.ishod != ISHOD_ZA_ODLUKU:
                _za_odluku(rez, [n], "IZUZETAK_POTPIS_SE_NE_SLAZE", napomena="; ".join(razlike))
            continue
        if n.ishod == ISHOD_ZA_ODLUKU:
            # Izuzetak sme da razresi i red koji osnovna provera zadrzava
            # (npr. nepoznat tip) -- tada ga vadimo iz liste za odluku.
            rez.za_odluku = [z for z in rez.za_odluku if n.redni not in z["noge"]]
            n.ishod = ""
        rez.izuzeci_primenjeni.append(dict(amb_id=n.amb_id, redni=n.redni, radnja=iz["radnja"],
                                           razlog=tekst(iz.get("razlog"))))
        if iz["radnja"] == "izuzmi":
            n.ishod = ISHOD_IZUZETO
            n.razlog = tekst(iz.get("razlog")) or "IZUZETAK"
        elif iz["radnja"] == "za_odluku":
            _za_odluku(rez, [n], "IZUZETAK_ZA_ODLUKU", napomena=tekst(iz.get("razlog")))
        else:
            if n.kolicina is None or n.kolicina <= 0 or n.datum is None or not n.tip:
                rez.izuzeci_primenjeni.pop()
                _za_odluku(rez, [n], "IZUZETAK_NAD_NEISPRAVNIM_REDOM",
                           napomena="Kretanje trazi ispravnu kolicinu, datum i tip ambalaze.")
                continue
            k = iz["kretanje"]
            od = (_kanon_nalog(k["Od"][0]), tekst(k["Od"][1]))
            na = (_kanon_nalog(k["Na"][0]), tekst(k["Na"][1]))
            vrsta_dok = k.get("Dokument", "REVERS")
            if vrsta_dok in ("REVERS", "NABAVKA"):
                # Ambalazni dokument (tblAmbalazaDokument); vlasnik broja je
                # sopstveni nalog -- za NABAVKU onaj koji gajbe prima.
                vlasnik = od if od[0] in SOPSTVENI else na
                broj_dok = n.dok_id or "MIG/" + _izvorni_kljuc(n)
                dok_id = _zaglavlje(rez, vrsta_dok, vlasnik, n.datum, broj_dok, (od, na),
                                    "Istorijski %s %s (izuzetak)" % (n.dok_tip, n.amb_id))
                dok_tip = DOK_AMB
            else:
                dok_tip, dok_id = k["Dokument"], n.dok_id
            _dodaj(rez, [n], od, na, k["VrstaKretanja"], dok_tip, dok_id, "IZUZETAK",
                   napomena=tekst(iz.get("razlog")))


def _zaglavlje(rez, vrsta, vlasnik, dan, broj_dok, par, napomena):
    """Nacrt zaglavlja tblAmbalazaDokument. Jedan par naloga po dokumentu
    (AMB-INV-10), pa isti broj reversa sa dva partnera daje dva zaglavlja."""
    partneri = sorted("%s:%s" % (p[0], kljuc(p[1])) for p in par)
    adk = novi_id("ADK-", vrsta, vlasnik[0], kljuc(vlasnik[1]), dan.isoformat() if dan else "",
                  kljuc(broj_dok), *partneri)
    if adk not in rez.zaglavlja:
        rez.zaglavlja[adk] = dict(AmbDokID=adk, Vrsta=vrsta, BrojDokumenta=broj_dok, Datum=dan,
                                  BrojOwnerTip=vlasnik[0], BrojOwnerID=vlasnik[1],
                                  Napomena=napomena, partneri=" / ".join(partneri), kretanja=0)
    rez.zaglavlja[adk]["kretanja"] += 1
    return adk


def _dodaj(rez, noge, od, na, vrsta, dok_tip, dok_id, pravilo, napomena=""):
    n0 = noge[0]
    amb = novi_id("AMB-", "kretanje", *sorted(_izvorni_kljuc(n) for n in noge))
    k = Kretanje(amb_id=amb, datum=n0.datum, tip=n0.tip, kolicina=n0.kolicina,
                 dok_tip=dok_tip, dok_id=dok_id, od_tip=od[0], od_id=od[1],
                 na_tip=na[0], na_id=na[1], vrsta=vrsta, pravilo=pravilo,
                 stari_tip=n0.dok_tip, stari_broj=n0.dok_id,
                 noge=tuple(n.redni for n in noge), stari_amb=tuple(n.amb_id for n in noge),
                 redosled=min(n.redni for n in noge),
                 created_at=n0.created_at, created_by=n0.created_by, napomena=napomena)
    rez.kretanja.append(k)
    for n in noge:
        n.ishod = ISHOD_KRETANJE
        n.kretanje = amb
    return k


def _upozori(rez, n, vrsta, opis):
    rez.upozorenja.append(dict(vrsta=vrsta, redni=n.redni, amb_id=n.amb_id, opis=opis))


# ---------------------------------------------------------------------------
# Sparivanje dvostranih upisa (Otkup, OM-*-Koop)
# ---------------------------------------------------------------------------

def _spari(noge, uloga_a, uloga_b):
    """Spari noge dve uloge po (datum, tip, kolicina). Vrati (parovi, ostatak, razlog).

    Kad u istoj grupi ima vise nogu na OBE strane i vise razlicitih naloga na
    obe strane, par bi bio pogodjen, a ne dokazan -- takva grupa ide za odluku.
    """
    a = [n for n in noge if n.uloga == uloga_a]
    b = [n for n in noge if n.uloga == uloga_b]
    strane = [n for n in noge if n.uloga not in (uloga_a, uloga_b)]
    if strane:
        return [], noge, "NEOCEKIVANA_ULOGA_NOGE"
    po_potpisu = defaultdict(lambda: ([], []))
    for n in a:
        po_potpisu[(n.datum, norm_tip(n.tip), n.kolicina)][0].append(n)
    for n in b:
        po_potpisu[(n.datum, norm_tip(n.tip), n.kolicina)][1].append(n)
    parovi, ostatak, razlog = [], [], ""
    for (aa, bb) in po_potpisu.values():
        aa.sort(key=lambda x: x.redni)
        bb.sort(key=lambda x: x.redni)
        if (len(aa) > 1 and len(bb) > 1 and len({kljuc(x.ent_id) for x in aa}) > 1
                and len({kljuc(x.ent_id) for x in bb}) > 1):
            ostatak += aa + bb
            razlog = "DVOSMISLENO_SPARIVANJE"
            continue
        m = min(len(aa), len(bb))
        parovi += list(zip(aa[:m], bb[:m]))
        ostatak += aa[m:] + bb[m:]
    if ostatak and not razlog:
        razlog = "NOGA_BEZ_PARA"
    return parovi, ostatak, razlog


def _roditelj_problem(rez, dt, dok_id, roditelji):
    r = roditelji[dt].get(kljuc(dok_id))
    if r is None:
        return "NEMA_DOKUMENTA_RODITELJA", None
    if je_storno(r):
        return "AKTIVNA_NOGA_STORNIRANOG_DOKUMENTA", r
    return "", r


def _otkup(rez, noge, roditelji):
    grupe = defaultdict(list)
    for n in noge:
        grupe[kljuc(n.dok_id)].append(n)
    for dok, g in sorted(grupe.items(), key=lambda x: min(n.redni for n in x[1])):
        problem, rod = _roditelj_problem(rez, "otkup", dok, roditelji)
        if problem:
            _za_odluku(rez, g, problem, predlog="izuzmi" if problem.startswith("AKTIVNA") else "")
            continue
        parovi, ostatak, razlog = _spari(g, ("Kooperant", "izlaz"), ("Stanica", "ulaz"))
        for koop, st in parovi:
            _proveri_vlasnika(rez, koop, rod, "KooperantID")
            _proveri_vlasnika(rez, st, rod, "StanicaID")
            _dodaj(rez, [koop, st], ("Kooperant", koop.ent_id), ("Stanica", st.ent_id),
                   UZ_ROBU, "Otkup", tekst(polje(rod, "OtkupID")), "OTKUP_PAR")
        if ostatak:
            _za_odluku(rez, ostatak, razlog)


def _om_koop(rez, noge, roditelji, izlaz):
    if izlaz:
        a, b = ("Kooperant", "ulaz"), ("Stanica", "izlaz")
    else:
        a, b = ("Kooperant", "izlaz"), ("Stanica", "ulaz")
    grupe = defaultdict(list)
    for n in noge:
        grupe[kljuc(n.dok_id)].append(n)
    for dok, g in sorted(grupe.items(), key=lambda x: min(n.redni for n in x[1])):
        rod = roditelji["otkup"].get(dok) if izlaz and dok else None
        parovi, ostatak, razlog = _spari(g, a, b)
        for koop, st in parovi:
            if rod is not None:
                # DokumentID je OtkupID: izdavanje praznih UZ OTKUP (Otkup.KolAmbIzdata).
                # Isti broj sa drugim nalozima nije dokaz ni otkupa ni reversa
                # (MIG-05.1 OTKUP_ID_PRESENT_CONTEXT_REVIEW) -- ide za odluku.
                if kljuc(polje(rod, "KooperantID")) != kljuc(koop.ent_id) \
                        or kljuc(polje(rod, "StanicaID")) != kljuc(st.ent_id):
                    _za_odluku(rez, [koop, st], "IZDAVANJE_UZ_OTKUP_NALOG_SE_NE_SLAZE")
                    continue
                if je_storno(rod):
                    _za_odluku(rez, [koop, st], "AKTIVNA_NOGA_STORNIRANOG_DOKUMENTA", predlog="izuzmi")
                    continue
                _dodaj(rez, [koop, st], ("Stanica", st.ent_id), ("Kooperant", koop.ent_id),
                       IZDATA, "Otkup", tekst(polje(rod, "OtkupID")), "OTKUP_IZDATE_PRAZNE")
                continue
            if not koop.dok_id:
                _za_odluku(rez, [koop, st], "REVERS_BEZ_BROJA")
                continue
            stanica = ("Stanica", st.ent_id)
            kooperant = ("Kooperant", koop.ent_id)
            od, na, vrsta = (stanica, kooperant, IZDATA) if izlaz else (kooperant, stanica, POVRAT)
            adk = _zaglavlje(rez, "REVERS", stanica, st.datum, koop.dok_id, (od, na),
                             "Migriran istorijski revers (%s, izvorni broj zadrzan)." % koop.dok_tip)
            _dodaj(rez, [koop, st], od, na, vrsta, DOK_AMB, adk,
                   "REVERS_IZDATE_PRAZNE" if izlaz else "REVERS_POVRAT_PRAZNE")
        if ostatak:
            _za_odluku(rez, ostatak, razlog)


# ---------------------------------------------------------------------------
# Jednostrani upisi sa vozacem (ogledalo)
# ---------------------------------------------------------------------------

JEDNOSTRANE = {
    # dt: {(entitet, smer): (od_je_entitet, vrsta, pravilo)}
    "om-izlaz-firma": {("Stanica", "izlaz"): (True, PRENOS, "REVERS_STANICA_VOZAC")},
    "om-ulaz-firma": {("Stanica", "ulaz"): (False, PRENOS, "REVERS_VOZAC_STANICA")},
    "otpremnica": {("Stanica", "izlaz"): (True, UZ_ROBU, "OTPREMNICA_VOZAC")},
    "prijemnica": {("Kupac", "ulaz"): (False, UZ_ROBU, "PRIJEMNICA_PUNE"),
                   ("Kupac", "izlaz"): (True, POVRAT, "PRIJEMNICA_POVRAT")},
    "kupci-otpremnica": {("Kupac", "izlaz"): (True, POVRAT, "KUPAC_VRACA_PRAZNE")},
}
VLASNIK_POLJE = {"otpremnica": "StanicaID", "prijemnica": "KupacID"}


def _jednostrane(rez, noge, dt, roditelji, ids):
    pravila = JEDNOSTRANE[dt]
    for n in sorted(noge, key=lambda x: x.redni):
        if n.uloga not in pravila:
            _za_odluku(rez, [n], "NEOCEKIVANA_ULOGA_NOGE")
            continue
        if not n.vozac:
            _za_odluku(rez, [n], "NEMA_VOZACA_ZA_OGLEDALO")
            continue
        od_je_ent, vrsta, pravilo = pravila[n.uloga]
        ent = (n.ent_tip, n.ent_id)
        voz = ("Vozac", n.vozac)
        od, na = (ent, voz) if od_je_ent else (voz, ent)
        if dt in RODITELJ:
            problem, rod = _roditelj_problem(rez, dt, n.dok_id, roditelji)
            if problem:
                _za_odluku(rez, [n], problem, predlog="izuzmi" if problem.startswith("AKTIVNA") else "")
                continue
            _proveri_vlasnika(rez, n, rod, VLASNIK_POLJE[dt])
            if kljuc(polje(rod, "VozacID")) and kljuc(polje(rod, "VozacID")) != kljuc(n.vozac):
                _upozori(rez, n, "VOZAC_SE_NE_SLAZE_SA_DOKUMENTOM",
                         "Vozac na nozi %s, na dokumentu %s." % (n.vozac, tekst(polje(rod, "VozacID"))))
            pk = RODITELJ[dt][1]
            _dodaj(rez, [n], od, na, vrsta, RODITELJ[dt][2], tekst(polje(rod, pk)), pravilo)
            continue
        if not n.dok_id:
            _za_odluku(rez, [n], "REVERS_BEZ_BROJA")
            continue
        if dt == "kupci-otpremnica":
            adk = _zaglavlje(rez, "REVERS_PARTNERA", ent, n.datum, n.dok_id, (od, na),
                             "Migriran istorijski Kupci-Otpremnica (kupcev broj zadrzan).")
        else:
            adk = _zaglavlje(rez, "REVERS", ent, n.datum, n.dok_id, (od, na),
                             "Migriran istorijski revers (%s, izvorni broj zadrzan)." % n.dok_tip)
        _dodaj(rez, [n], od, na, vrsta, DOK_AMB, adk, pravilo)


def _proveri_vlasnika(rez, n, rod, polje_ime):
    na_dok = tekst(polje(rod, polje_ime))
    if na_dok and kljuc(na_dok) != kljuc(n.ent_id):
        _upozori(rez, n, "NALOG_SE_NE_SLAZE_SA_DOKUMENTOM",
                 "%s na nozi %s, na dokumentu %s." % (n.ent_tip, n.ent_id, na_dok))


def _proveri_efekte(rez):
    """Svako kretanje mora imati TACNO isti efekat na salda kao njegove stare noge
    (MIG-05.2 ACCOUNTING_EFFECT_DIFFERS). Izuzetak sme da ga menja -- on je odluka."""
    po_rednom = {n.redni: n for n in rez.noge}
    for k in rez.kretanja:
        if k.pravilo == "IZUZETAK":
            continue
        staro = defaultdict(Decimal)
        for r in k.noge:
            for kl, v in po_rednom[r].doprinos():
                staro[_kl(kl)] += v
        novo = defaultdict(Decimal)
        for kl, v in k.doprinos():
            novo[_kl(kl)] += v
        staro = {a: b for a, b in staro.items() if b}
        novo = {a: b for a, b in novo.items() if b}
        if staro != novo:
            raise GreskaAmbalaze("Interna greska: kretanje %s (pravilo %s) menja salda drugacije od "
                                 "svojih nogu %s" % (k.amb_id, k.pravilo, k.noge))


def _kl(kl):
    return (kl[0], kljuc(kl[1]), norm_tip(kl[2]))


def _postojeca_knjiga(t, rez):
    """Sveska je vec AMB-10 oblika: kretanja su redovi; nema prevoda ni politike."""
    for r in t.redovi:
        od = (_kanon_nalog(polje(r, "OdNalogTip")), tekst(polje(r, "OdNalogID")))
        na = (_kanon_nalog(polje(r, "NaNalogTip")), tekst(polje(r, "NaNalogID")))
        rez.kretanja.append(Kretanje(
            amb_id=tekst(polje(r, "AmbID")), datum=datum(polje(r, "Datum")),
            tip=tekst(polje(r, "TipAmbalaze")), kolicina=broj(polje(r, "Kolicina")) or Decimal(0),
            dok_tip=tekst(polje(r, "DokumentTIP")), dok_id=tekst(polje(r, "DokumentID")),
            od_tip=od[0], od_id=od[1], na_tip=na[0], na_id=na[1],
            vrsta=tekst(polje(r, "VrstaKretanja")), pravilo="POSTOJECA_KNJIGA",
            redosled=r["_redni"], created_at=tekst(polje(r, "CreatedAt")),
            created_by=tekst(polje(r, "CreatedBy")),
            napomena=("storno od " + tekst(polje(r, "StornoOd"))) if tekst(polje(r, "StornoOd")) else ""))


# ---------------------------------------------------------------------------
# Politika duga prema partneru (MIG-16/17): dnevni zakljucak
# ---------------------------------------------------------------------------

def primeni_politiku_duga(rez, partneri):
    """Negativan stari saldo partnera = firma drzi njegove gajbe (obaveza firme).

    Po danu i tipu: rast obaveze -> ULAZ_TUDJE (SpoljniSvet -> partner) na
    partnerovim izlaznim kretanjima tog dana; pad obaveze -> izdavanje praznih
    tog dana postaje VRACANJE_TUDJE. Fizicki vektor se ne menja; novi saldo
    partnera na kraju dana = max(stari, 0), a obaveza = max(-stari, 0).
    """
    tragovi = []
    po_danu = defaultdict(list)
    for k in rez.kretanja:
        for strana, tipn, idn in (("Od", k.od_tip, k.od_id), ("Na", k.na_tip, k.na_id)):
            if tipn in partneri:
                po_danu[(tipn, kljuc(idn), norm_tip(k.tip))].append((k.datum, strana, k))
    nova = []
    for (tipn, idn, tip), stavke in sorted(po_danu.items(), key=lambda x: (x[0][0], x[0][1], x[0][2])):
        dani = defaultdict(list)
        for d, strana, k in stavke:
            dani[d].append((strana, k))
        stari = Decimal(0)
        for d in sorted(x for x in dani if x is not None):
            delta = sum((k.kolicina if s == "Na" else -k.kolicina) for s, k in dani[d])
            pre, posle = stari, stari + delta
            stari = posle
            dug_pre, dug_posle = max(Decimal(0), -pre), max(Decimal(0), -posle)
            rast, pad = max(Decimal(0), dug_posle - dug_pre), max(Decimal(0), dug_pre - dug_posle)
            if pad:
                izdavanja = sorted((k for s, k in dani[d] if s == "Na" and k.vrsta == IZDATA),
                                   key=lambda x: x.redosled)
                ostalo = pad
                for k in izdavanja:
                    if ostalo <= 0:
                        break
                    deo = min(k.kolicina, ostalo)
                    if deo == k.kolicina:
                        k.vrsta = VRACANJE_TUDJE
                        cilj = k
                    else:
                        # Podela: stare noge ostaju na izvornom kretanju (staro
                        # se broji jednom), odvojeni deo nosi samo novu kolicinu.
                        k.kolicina -= deo
                        cilj = Kretanje(**{**k.__dict__})
                        cilj.amb_id = novi_id("AMB-", "vracanje", k.amb_id)
                        cilj.kolicina = deo
                        cilj.vrsta = VRACANJE_TUDJE
                        cilj.generisano = True
                        cilj.pravilo = "PODELA_VRACANJE_TUDJE"
                        cilj.noge = ()
                        cilj.stari_amb = ()
                        cilj.napomena = "Deo kretanja %s (vracanje tudje ambalaze)." % k.amb_id
                        nova.append(cilj)
                    ostalo -= deo
                    tragovi.append(dict(uloga="VRACANJE_TUDJE", partner=(tipn, idn), tip=tip,
                                        datum=d, kolicina=deo, amb_id=cilj.amb_id))
                if ostalo > 0:
                    rez.za_odluku.append(dict(
                        razlog="PAD_OBAVEZE_BEZ_IZDAVANJA", noge=[], predlog="",
                        napomena="%s %s, %s, %s: obaveza pada za %s, a izdatih praznih tog dana nema dovoljno."
                                 % (tipn, idn, tip, d, ostalo)))
            if rast:
                izlazi = sorted((k for s, k in dani[d] if s == "Od"), key=lambda x: x.redosled)
                ostalo = rast
                for k in izlazi:
                    if ostalo <= 0:
                        break
                    deo = min(k.kolicina, ostalo)
                    pok = _pokrice(k, deo, "POKRICE_DUGA")
                    nova.append(pok)
                    ostalo -= deo
                    tragovi.append(dict(uloga="ULAZ_TUDJE", partner=(tipn, idn), tip=tip,
                                        datum=d, kolicina=deo, amb_id=pok.amb_id))
                if ostalo > 0:
                    rez.za_odluku.append(dict(
                        razlog="RAST_OBAVEZE_BEZ_IZLAZA", noge=[], predlog="",
                        napomena="%s %s, %s, %s: obaveza raste za %s, a partnerovih izlaza tog dana nema dovoljno."
                                 % (tipn, idn, tip, d, ostalo)))
    rez.kretanja.extend(nova)
    return tragovi


def _pokrice(k, deo, pravilo):
    """ULAZ_TUDJE (SpoljniSvet -> partner) na ISTOM dokumentu kao partnerov
    izlaz -- isto mesto na koje ga stavlja i zivi pisac (PrenesiAmbalazu)."""
    return Kretanje(
        amb_id=novi_id("AMB-", "pokrice", k.amb_id), datum=k.datum, tip=k.tip,
        kolicina=deo, dok_tip=k.dok_tip, dok_id=k.dok_id,
        od_tip=GRANICA, od_id="", na_tip=k.od_tip, na_id=k.od_id,
        vrsta=ULAZ_TUDJE, pravilo=pravilo, stari_tip=k.stari_tip,
        stari_broj=k.stari_broj, redosled=k.redosled, created_at=k.created_at,
        created_by="MIGRACIJA", generisano=True,
        napomena="Tudja ambalaza partnera (stari negativan saldo).")


def _po_nalogu_i_danu(kretanja, tipovi):
    """{(tip, id, tip amb.): (prikaz_id, prikaz_tip, {datum: [(strana, kretanje)]})}"""
    out = {}
    for k in kretanja:
        for strana, tipn, idn in (("Od", k.od_tip, k.od_id), ("Na", k.na_tip, k.na_id)):
            if tipn not in tipovi:
                continue
            kl = (tipn, kljuc(idn), norm_tip(k.tip))
            if kl not in out:
                out[kl] = (idn, k.tip, defaultdict(list))
            out[kl][2][k.datum].append((strana, k))
    return out


def _delta(stavke):
    return sum((k.kolicina if s == "Na" else -k.kolicina) for s, k in stavke)


def pokrice_po_manjku(rez, tipovi):
    """Partner koji ne prima izdate prazne (kupac): pokrice TACNO u trenutku manjka.

    Isto kao zivi pisac (AMB-INV-07): kad bi saldo pao ispod nule, manjak ulazi
    kao ULAZ_TUDJE na partnerovom izlaznom kretanju tog dana. Obaveza ostaje dok
    je ne zatvori VRACANJE_TUDJE -- i u zivom radu je tako. Pravilo MIG-16/17
    ovde ne vazi: ono obavezu smanjuje preko izdatih praznih, a kupcu se prazne
    ne izdaju.
    """
    tragovi, nova = [], []
    for (tipn, idn, tip), (_pid, _ptip, dani) in sorted(_po_nalogu_i_danu(rez.kretanja, tipovi).items()):
        b = Decimal(0)
        for d in sorted(x for x in dani if x is not None):
            b += _delta(dani[d])
            if b >= 0:
                continue
            potrebno = -b
            for k in sorted((k for s, k in dani[d] if s == "Od"), key=lambda x: x.redosled):
                if potrebno <= 0:
                    break
                deo = min(k.kolicina, potrebno)
                pok = _pokrice(k, deo, "POKRICE_MANJKA")
                nova.append(pok)
                potrebno -= deo
                b += deo
                tragovi.append(dict(uloga="ULAZ_TUDJE", partner=(tipn, idn), tip=tip,
                                    datum=d, kolicina=deo, amb_id=pok.amb_id))
            if potrebno > 0:
                rez.za_odluku.append(dict(
                    razlog="RAST_OBAVEZE_BEZ_IZLAZA", noge=[], predlog="",
                    napomena="%s %s, %s, %s: manjak %s bez partnerovog izlaza tog dana."
                             % (tipn, idn, tip, d, potrebno)))
    rez.kretanja.extend(nova)
    return tragovi


def nabavka_na_pocetku(rez, tipovi):
    """Sopstveni nalog (stanica, vozac) koji je ijednog dana u minusu dobija
    JEDNU NABAVKU na prvi dan svog rada, u iznosu najdubljeg minusa.

    Stara sveska nikad nije upisala koliko gajbi firma ima. AMBALAZA.md 6.7b:
    "firma ima gajbe" je NABAVKA SpoljniSvet -> Stanica. Posle nje nalog ni
    jednog dana nije ispod nule (main inace odbija izdavanje i storno nad
    nalogom u minusu), a saldo na kraju = stari saldo + nabavka.
    """
    tragovi = []
    for (tipn, idn, tip), (pid, ptip, dani) in sorted(_po_nalogu_i_danu(rez.kretanja, tipovi).items()):
        dati = sorted(x for x in dani if x is not None)
        if not dati:
            continue
        b = najm = Decimal(0)
        for d in dati:
            b += _delta(dani[d])
            najm = min(najm, b)
        if najm >= 0:
            continue
        prvi = dati[0]
        adk = _zaglavlje(rez, "NABAVKA", (tipn, pid), prvi, "MIG-POC/%s/%s" % (pid, ptip),
                         ((GRANICA, ""), (tipn, pid)),
                         "Pocetno stanje sezone (migracija): gajbe firme koje stara sveska nije upisala.")
        k = Kretanje(amb_id=novi_id("AMB-", "nabavka", tipn, idn, tip), datum=prvi, tip=ptip,
                     kolicina=-najm, dok_tip=DOK_AMB, dok_id=adk, od_tip=GRANICA, od_id="",
                     na_tip=tipn, na_id=pid, vrsta=NABAVKA, pravilo="NABAVKA_POCETNO_STANJE",
                     redosled=0, created_by="MIGRACIJA", generisano=True,
                     napomena="Najdublji minus u sezoni; posle nabavke nalog nije ispod nule.")
        rez.kretanja.append(k)
        tragovi.append(dict(uloga="NABAVKA", partner=(tipn, idn), tip=tip, datum=prvi,
                            kolicina=-najm, amb_id=k.amb_id))
    return tragovi


# ---------------------------------------------------------------------------
# Salda, provere, stavke po nalogu
# ---------------------------------------------------------------------------

def proveri_pokrivenost(rez):
    """Svaki stari red ima tacno jedan ishod; noga u kretanju je u tacno jednom."""
    bez = [n.redni for n in rez.noge if n.ishod not in (ISHOD_KRETANJE, ISHOD_STORNO,
                                                         ISHOD_IZUZETO, ISHOD_ZA_ODLUKU)]
    u_kretanju = Counter(r for k in rez.kretanja for r in k.noge)
    dvaput = [r for r, c in u_kretanju.items() if c > 1]
    kretanje_bez = [n.redni for n in rez.noge
                    if n.ishod == ISHOD_KRETANJE and u_kretanju.get(n.redni) != 1]
    van = [r for r in u_kretanju if r not in {n.redni for n in rez.noge if n.ishod == ISHOD_KRETANJE}]
    return dict(bez_ishoda=bez, u_dva_kretanja=dvaput, kretanje_bez_reda=kretanje_bez,
                red_van_kretanja=van)


def izracunaj_salda(rez):
    """Po nalogu i tipu: stari saldo, novi saldo, pokrice, nabavka, van knjige, izuzeci.

    stari - van_knjige + izuzeci + pokrice + nabavka == novi   (razlika mora biti 0)
    """
    s = defaultdict(lambda: defaultdict(Decimal))
    prikaz = {}
    po_rednom = {n.redni: n for n in rez.noge}

    def dodaj(kl, kol, v):
        k = _kl(kl)
        prikaz.setdefault(k, kl)
        s[k][kol] += v

    for n in rez.noge:
        for kl, v in n.doprinos():
            dodaj(kl, "stari", v)
            if n.ishod in (ISHOD_IZUZETO, ISHOD_ZA_ODLUKU):
                dodaj(kl, "van_knjige", v)
    for k in rez.kretanja:
        for kl, v in k.doprinos():
            dodaj(kl, "novi", v)
            if k.vrsta == ULAZ_TUDJE and k.generisano:
                dodaj(kl, "pokrice", v)
            if k.vrsta == NABAVKA and k.generisano:
                dodaj(kl, "nabavka", v)
        if k.pravilo == "IZUZETAK":
            for kl, v in k.doprinos():
                dodaj(kl, "izuzeci", v)
            for r in k.noge:
                for kl, v in po_rednom[r].doprinos():
                    dodaj(kl, "izuzeci", -v)
    if rez.oblik == "AMB10":
        for k, v in s.items():
            v["stari"] = v["novi"]

    dnevno = _dnevni_zakljucci(rez)
    out = []
    for k in sorted(s, key=lambda x: (x[0], x[1], x[2])):
        v = s[k]
        razlika = (v["stari"] - v["van_knjige"] + v["izuzeci"] + v["pokrice"] + v["nabavka"]
                   - v["novi"])
        d = dnevno.get(k, {})
        out.append(dict(nalog_tip=prikaz[k][0], nalog_id=prikaz[k][1], tip=prikaz[k][2],
                        naziv=rez.nazivi.get((prikaz[k][0], k[1]), ""),
                        stari=v["stari"], novi=v["novi"], pokrice=v["pokrice"], nabavka=v["nabavka"],
                        van_knjige=v["van_knjige"], izuzeci=v["izuzeci"], razlika=razlika,
                        najmanji=d.get("najmanji", Decimal(0)), najmanji_dan=d.get("najmanji_dan"),
                        obaveza=d.get("obaveza", Decimal(0)),
                        najmanja_obaveza=d.get("najmanja_obaveza", Decimal(0))))
    return out


def _dnevni_zakljucci(rez):
    """Novi saldo na kraju svakog dana; najmanji zakljucni saldo i obaveza partnera."""
    po = defaultdict(lambda: defaultdict(Decimal))
    obaveza = defaultdict(lambda: defaultdict(Decimal))
    for k in rez.kretanja:
        for kl, v in k.doprinos():
            po[_kl(kl)][k.datum] += v
        if k.na_tip in PARTNERI and k.vrsta in (ULAZ_TUDJE, "POZAJMICA_PARTNERA") and k.od_tip == GRANICA:
            obaveza[_kl((k.na_tip, k.na_id, k.tip))][k.datum] += k.kolicina
        if k.na_tip in PARTNERI and k.vrsta == VRACANJE_TUDJE:
            obaveza[_kl((k.na_tip, k.na_id, k.tip))][k.datum] -= k.kolicina
    out = {}
    for kl, dani in po.items():
        tek = Decimal(0)
        najm, najm_dan = None, None
        for d in sorted(x for x in dani if x is not None):
            tek += dani[d]
            if najm is None or tek < najm:
                najm, najm_dan = tek, d
        ob, najm_ob = Decimal(0), Decimal(0)
        for d in sorted(x for x in obaveza.get(kl, {}) if x is not None):
            ob += obaveza[kl][d]
            najm_ob = min(najm_ob, ob)
        out[kl] = dict(najmanji=najm if najm is not None else Decimal(0), najmanji_dan=najm_dan,
                       obaveza=ob, najmanja_obaveza=najm_ob)
    return out


def proveri_invarijante(rez):
    """AMB-INV-01/02/04/10/11/12 nad projektovanom knjigom (AMBALAZA.md 6.9).

    INV-07 (nijedan realan nalog ispod nule) i INV-09 (obaveza >= 0) se citaju
    iz dnevnih zakljucaka u izracunaj_salda -- tamo su i iznosi potrebnog pokrica.
    """
    nalazi = []
    ids = Counter(kljuc(k.amb_id) for k in rez.kretanja)
    for i, c in ids.items():
        if c > 1:
            nalazi.append(dict(inv="AMB-INV-11", opis="AmbID se ponavlja %d puta" % c, amb_id=i))
    identitet = defaultdict(list)
    parovi = defaultdict(set)
    for k in rez.kretanja:
        if not k.kolicina or k.kolicina <= 0:
            nalazi.append(dict(inv="AMB-INV-01", opis="Kolicina nije > 0", amb_id=k.amb_id))
        if k.od_tip == k.na_tip and kljuc(k.od_id) == kljuc(k.na_id):
            nalazi.append(dict(inv="AMB-INV-02", opis="Od = Na", amb_id=k.amb_id))
        if rez.tipovi_ambalaze is not None and norm_tip(k.tip) not in rez.tipovi_ambalaze:
            nalazi.append(dict(inv="AMB-INV-12", opis="Tip ambalaze '%s' nije u tblTipAmbalaze" % k.tip,
                               amb_id=k.amb_id))
        if k.vrsta not in VRSTE:
            nalazi.append(dict(inv="VRSTA", opis="Nepoznata vrsta kretanja %s" % k.vrsta, amb_id=k.amb_id))
        if k.pravilo == "POSTOJECA_KNJIGA" and k.napomena.startswith("storno od"):
            continue
        identitet[(kljuc(k.dok_tip), kljuc(k.dok_id), k.vrsta, norm_tip(k.tip))].append(k)
        if not (k.generisano and k.vrsta == ULAZ_TUDJE):
            parovi[(kljuc(k.dok_tip), kljuc(k.dok_id))].add(k.par())
    for ident, ks in identitet.items():
        if len(ks) > 1:
            nalazi.append(dict(inv="AMB-INV-04",
                               opis="%d kretanja dele identitet (DokumentTIP, DokumentID, Vrsta, Tip) = %s"
                                    % (len(ks), "/".join(ident)),
                               amb_id=";".join(k.amb_id for k in ks)))
    for dok, ps in parovi.items():
        if len(ps) > 1:
            nalazi.append(dict(inv="AMB-INV-10",
                               opis="Dokument %s nosi %d razlicita para naloga" % ("/".join(dok), len(ps)),
                               amb_id=""))
    return nalazi


def stavke_po_nalogu(rez):
    """Za svaki realan nalog: hronoloske stavke sa tekucim novim i starim saldom.

    Kretanje daje novi efekat; stari efekat je efekat njegovih starih nogu (0
    za generisano pokrice). Noga van knjige (izuzeta / za odluku) ima samo
    stari efekat. Tako poslednja stavka nosi oba konacna salda.
    """
    po_rednom = {n.redni: n for n in rez.noge}
    stavke = defaultdict(list)
    for k in rez.kretanja:
        staro = defaultdict(Decimal)
        if rez.oblik == "AMB10":
            for kl, v in k.doprinos():
                staro[_kl(kl)] += v
        for r in k.noge:
            for kl, v in po_rednom[r].doprinos():
                staro[_kl(kl)] += v
        novo = defaultdict(Decimal)
        for kl, v in k.doprinos():
            novo[_kl(kl)] += v
        for kl in set(novo) | set(staro):
            if kl[0] == GRANICA:
                continue
            druga = (k.na_tip, k.na_id) if (k.od_tip, kljuc(k.od_id)) == (kl[0], kl[1]) else (k.od_tip, k.od_id)
            stavke[kl].append(dict(datum=k.datum, redosled=k.redosled, kretanje=k, noga=None,
                                   novo=novo.get(kl, Decimal(0)), staro=staro.get(kl, Decimal(0)),
                                   druga=druga))
    for n in rez.noge:
        if n.ishod not in (ISHOD_IZUZETO, ISHOD_ZA_ODLUKU):
            continue
        for kl, v in n.doprinos():
            stavke[_kl(kl)].append(dict(datum=n.datum, redosled=n.redni, kretanje=None, noga=n,
                                        novo=Decimal(0), staro=v, druga=None))
    out = {}
    for kl, lst in stavke.items():
        lst.sort(key=lambda x: (x["datum"] is None, x["datum"] or 0, x["redosled"],
                                0 if x["kretanje"] is None or not x["kretanje"].generisano else 1))
        novi = stari = Decimal(0)
        for st in lst:
            novi += st["novo"]
            stari += st["staro"]
            st["saldo_novi"] = novi
            st["saldo_stari"] = stari
        out[kl] = lst
    return out
