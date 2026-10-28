"""Nova sveska U MEMORIJI: svaki red svake tabele, po kanonu schema/schema.json.

Nista se ne pise. Ovo je ulaz i za probu (`migracija.py proba`, paket za
pregled + provere) i za upis u svesku napravljenu iz gita.

Pravila (odluke operatera 28.10.2026; trag u docs/DOMEN/MIGRACIJA_LEGACY.md):

  DOKUMENTI  stari red = jedno zaglavlje + jedna stavka (klijent ima jednu klasu
             po dokumentu -- izmereno: 0 dokumenata sa vise redova). Zaglavlje
             zadrzava stari ID; stavka dobija nov deterministicki ID u obliku
             koji pravi i main (OKS-/OPS-/ZBS-/PRS- + 32 hex).
             Kolone koje danasnji pisac NE pise ostaju prazne (tblOtkup.
             OtpremnicaID/BrojOtpremnice/BrojZbirne/ZbirnaRoditeljID/VozacID,
             tblOtpremnica.BrojZbirne/ZbirnaRoditeljID, GeneracijaID...): popunjene
             bi ukljucile stare citaoce (dupli broj blokova u storno lancu, A1).
  CLANSTVO   otkup -> otpremnica iz starog tblOtkup.OtpremnicaID (eksplicitan ID,
             MIG-11); otpremnica -> zbirna iz tblOtpremnica.BrojZbirne, po broju
             pa po vozacu i datumu (MIG-11 lookup_zbr); prijemnica -> zbirna kao
             pisac prijemnice: aktivna zbirna istog broja, kupca i vozaca
             (ZbirnaRoditeljID) uz zadrzan BrojZbirne.
  STATUS     otpremnica i zbirna: IZDATO (stari program nije imao nacrt, MIG-14).
             Otkup: IZDATO, OSIM aktivnog otkupa bez aktivne otpremnice -- njega
             bi sledeci sync sam stavio u automatsku otpremnicu (modMasterSync
             GrupeZaAutoOtpremnicu, modAutoHladnjaca.AutoLanacZaNevezaneBlokove
             traze IZDATO), pa ostaje prazan dok operater ne odluci.
             Prijemnica: prazno (niko ga ne cita).
  AMBALAZA   ambalaza.prevedi(knjizenje=ISPRAVNO) -> tblAmbalaza + tblAmbalazaDokument.
  OSTALO     maticne, novac, banka, magacin, palete, prerada, fakture, SEF,
             utovar, korisnici: 1:1 po imenu kolone. Paleta i faktura stavka
             dobijaju PrijemnicaStavkaID (main ih vezuje preko nje).
  NE PRENOSI tblStornoVeze, tblStornoZurnal (pokazuju na kolone starog modela),
             tblPoruke (iz nove sveske).
  PODESAVANJA tblConfig / tblSEFConfig / tblLocalConfig: vrednost iz stare sveske
             prepisuje test vrednost nove; aktivacija licence se ne prenosi.
"""

import uuid
from collections import defaultdict

import ambalaza
from citac import broj, datum, je_storno, kljuc, polje, tabela, tekst

NS_CILJ = uuid.UUID("2b9e3d4a-71f0-5c8e-b2a6-4e1d9c7f0a35")

IZDATO = "IZDATO"


def novi_id(prefiks, *delovi):
    return prefiks + uuid.uuid5(NS_CILJ, "|".join(str(d) for d in delovi)).hex.upper()


DOKUMENTI = {
    "tblOtkup": dict(
        stavke="tblOtkupStavke", pk="OtkupID", spk="OtkupStavkaID", prefiks="OKS-",
        mapa={"Klasa": "Klasa", "Kolicina": "Kolicina", "Cena": "Cena",
              "KolAmbalaze": "KolAmbalaze", "BrutoKg": "BrutoKg"},
        prazno=("OtpremnicaID", "BrojOtpremnice", "BrojZbirne", "ZbirnaRoditeljID",
                "GeneracijaID", "VozacID", "CorrectionID")),
    "tblOtpremnica": dict(
        stavke="tblOtpremnicaStavke", pk="OtpremnicaID", spk="OtpremnicaStavkaID", prefiks="OPS-",
        mapa={"Klasa": "Klasa", "Kolicina": "Kolicina", "KolAmbalaze": "KolAmbalaze",
              "BrutoKg": "BrutoKg", "Cena": "PredlogCena"},
        prazno=("BrojZbirne", "ZbirnaRoditeljID", "GeneracijaID", "IspravkaOd", "ZamenjenSa",
                "CorrectionID")),
    "tblZbirna": dict(
        stavke="tblZbirnaStavke", pk="ZbirnaID", spk="ZbirnaStavkaID", prefiks="ZBS-",
        mapa={"Klasa": "Klasa", "UkupnoKolicina": "Kolicina", "UkupnoAmbalaze": "KolAmbalaze"},
        prazno=("UkupnoKolicina", "UkupnoAmbalaze", "Klasa", "GeneracijaID", "IspravkaOd",
                "ZamenjenSa", "CorrectionID")),
    "tblPrijemnica": dict(
        stavke="tblPrijemnicaStavke", pk="PrijemnicaID", spk="PrijemnicaStavkaID", prefiks="PRS-",
        mapa={"Klasa": "Klasa", "Kolicina": "Kolicina", "Cena": "Cena",
              "KolAmbalaze": "KolAmbalaze", "BrutoKg": "BrutoKg"},
        prazno=("CorrectionID", "IzdatoStatus")),
}

PREPIS = ("tblKooperanti", "tblStanice", "tblKupci", "tblVozaci", "tblKulture", "tblParcele",
          "tblTipAmbalaze", "tblTipPalete", "tblKutije", "tblKese", "tblVrstaGotovihProizvoda",
          "tblArtikli", "tblCenovnik", "tblPrevoznici", "tblPartnerMap", "tblMagacin",
          "tblNovac", "tblBankaImport", "tblPaleta", "tblPaletaStavka", "tblPrerada",
          "tblPreradaStavka", "tblFakture", "tblFakturaStavke", "tblSEFSubmission",
          "tblSEFEventLog", "tblUtovar", "tblUtovarStavke", "tblKorisnici", "tblMGMT")

NE_PRENOSI = {
    "tblStornoVeze": "pokazuje na kolone starog modela (Vrati storno / Zavrsi ispravku)",
    "tblStornoZurnal": "pokazuje na kolone starog modela (Vrati storno)",
    "tblPoruke": "katalog poruka dolazi iz nove sveske (EnsurePoruke)",
}

# Podesavanja (key/value): vrednost iz STARE sveske pobedjuje -- nova sveska iz
# gita nosi test vrednosti (make_fixture: MALINA_MODE=NO, "Test prodavac DOO").
# Aktivacija licence je vezana za masinu i ne prenosi se (kao modMigracija na
# drugoj masini): licenca se aktivira ponovo.
PODESAVANJA = {"tblConfig": ("Kljuc", "Vrednost"), "tblSEFConfig": ("ConfigKey", "ConfigValue"),
               "tblLocalConfig": ("Kljuc", "Vrednost")}
AKTIVACIJA_LICENCE = {"LICENSE_KEY", "LICENSE_TOKEN", "LICENSE_BOUND_PARTS", "LICENSE_HWM",
                      "LICENSE_STATUS", "LICENSE_NEXT_CHECK", "TRIAL_START", "TRIAL_HWM"}

# Kolone u kojima je broj zapravo TEKST (modSchema.SchemaReadyOrFail trazi "@").
TEKST_KOLONE = {"BrojDokumenta", "BrojOtpremnice", "BrojZbirne", "TipAmbalaze", "BrojPrijemnice"}


class Cilj:
    """Redovi nove sveske: tabele[ime] = list[dict kolona -> vrednost]."""

    def __init__(self, kanon):
        self.kanon = kanon
        self.tabele = {ime: [] for ime in kanon}
        self.nalazi = []              # (tezina, pravilo, tabela, id, opis)
        self.izvor_redova = {}        # tabela -> broj starih redova
        self.stavka_za = {}           # (tabela dokumenta, ID.casefold()) -> ID stavke
        self.amb = None
        self.podesavanja = {}         # tabela -> [(kljuc, {kolona: vrednost})]

    def kljuc_podesavanja(self, ime):
        return PODESAVANJA[ime][0]

    def nalaz(self, tezina, pravilo, tbl, idn, opis):
        self.nalazi.append(dict(tezina=tezina, pravilo=pravilo, tabela=tbl, id=idn, opis=opis))


def _prepisi_red(src, kolone):
    return {k: polje(src, k) for k in kolone}


def ucitaj_izuzetke_dokumenata(putanja):
    """Sekcija "dokumenti" iz izuzeci.json: otpremnice kojima se veza NE popravlja."""
    import json
    if not putanja:
        return []
    with open(putanja, encoding="utf-8") as f:
        lista = json.load(f).get("dokumenti", [])
    for iz in lista:
        if iz.get("radnja") != "bez_popravke_veze" or not tekst(iz.get("OtpremnicaID")):
            raise ValueError("izuzeci.json, dokumenti: nepoznat oblik %r" % (iz,))
    return lista


def izgradi(tabele, kanon, izuzeci=(), knjizenje=ambalaza.ISPRAVNO, izuzeci_dok=()):
    c = Cilj(kanon)
    for ime in PREPIS:
        t = tabela(tabele, ime)
        if t is None or ime not in kanon:
            continue
        c.izvor_redova[ime] = len(t.redovi)
        for r in t.redovi:
            c.tabele[ime].append(_prepisi_red(r, kanon[ime]))
        _visak_kolona(c, t, kanon[ime])
    for ime, razlog in NE_PRENOSI.items():
        t = tabela(tabele, ime)
        if t is not None and t.redovi:
            c.nalaz("INFO", "NE_PRENOSI_SE", ime, "", "%d red. -- %s" % (len(t.redovi), razlog))

    _dokumenti(c, tabele)
    _clanstvo(c, tabele)
    _status_dokumenata(c)
    _popravi_veze(c, {kljuc(iz["OtpremnicaID"]): iz for iz in izuzeci_dok})
    _otpremnice_bez_izvora(c)
    _status_otkupa(c)
    _prijemnica_roditelj(c, tabele)
    _kupac_hladnjaca(c, tabele)
    _stavke_nizvodno(c)
    _ambalaza(c, tabele, izuzeci, knjizenje)
    _tekst_kolone(c)
    _podesavanja(c, tabele)
    return c


def _podesavanja(c, tabele):
    for ime, (kk, kv) in PODESAVANJA.items():
        t = tabela(tabele, ime)
        if t is None or ime not in c.kanon:
            continue
        parovi, preskoceno = [], []
        for r in t.redovi:
            k = tekst(polje(r, kk))
            if not k:
                continue
            if k.upper() in AKTIVACIJA_LICENCE:
                preskoceno.append(k)
                continue
            if (k.upper().startswith("LICENSE_") or k.upper().startswith("TRIAL_")) and \
                    k.upper() not in ("LICENSE_ENABLED", "LICENSE_ENDPOINT", "TRIAL_ENABLED", "TRIAL_DAYS"):
                c.nalaz("UPOZORENJE", "NEPOZNAT_KLJUC_LICENCE", ime, k,
                        "prenosi se, ali proveri da li je vezan za masinu")
            parovi.append((k, {x: polje(r, x) for x in c.kanon[ime] if x in (kk, kv) or t.ima_kolonu(x)}))
        c.podesavanja[ime] = parovi
        if preskoceno:
            c.nalaz("ODLUKA", "AKTIVACIJA_LICENCE_SE_NE_PRENOSI", ime, "",
                    "%s -- licenca se aktivira ponovo na prvom pokretanju" % ", ".join(preskoceno))


def _tekst_kolone(c):
    """Broj u tekstualnoj koloni (BrojDokumenta...) ide kao tekst: main za te
    kolone trazi format "@" (modSchema.SchemaReadyOrFail) i poredi ih kao tekst."""
    from decimal import Decimal
    for redovi in c.tabele.values():
        for r in redovi:
            for k in TEKST_KOLONE:
                v = r.get(k)
                if isinstance(v, Decimal):
                    r[k] = format(v.normalize(), "f") if v != v.to_integral_value() else str(int(v))


def _visak_kolona(c, t, kolone):
    kanon_kf = {k.casefold() for k in kolone}
    visak = [k for k in t.kolone if k.casefold() not in kanon_kf]
    if visak:
        popunjene = [k for k in visak if any(tekst(polje(r, k)) for r in t.redovi)]
        if popunjene:
            c.nalaz("INFO", "STARA_KOLONA_SE_NE_PRENOSI", t.ime, "",
                    "kolone sa podacima kojih nema u kanonu: " + ", ".join(popunjene))


# ---------------------------------------------------------------------------
# Dokumenti: zaglavlje + jedna stavka
# ---------------------------------------------------------------------------

def _klasa(v):
    k = tekst(v).upper()
    return {"1": "I", "2": "II"}.get(k, k)


def _dokumenti(c, tabele):
    for ime, spec in DOKUMENTI.items():
        t = tabela(tabele, ime)
        if t is None:
            continue
        c.izvor_redova[ime] = len(t.redovi)
        kol_z = c.kanon[ime]
        kol_s = c.kanon[spec["stavke"]]
        mapa_kf = {k.casefold() for k in spec["mapa"]}
        prazno_kf = {k.casefold() for k in spec["prazno"]}
        for r in t.redovi:
            idn = tekst(polje(r, spec["pk"]))
            z = {}
            for k in kol_z:
                if k.casefold() in mapa_kf or k.casefold() in prazno_kf:
                    z[k] = ""
                else:
                    z[k] = polje(r, k)
            c.tabele[ime].append(z)
            s = {k: "" for k in kol_s}
            sid = novi_id(spec["prefiks"], ime, kljuc(idn))
            s[spec["spk"]] = sid
            s[spec["pk"]] = idn
            if "RedniBroj" in s:
                s["RedniBroj"] = 1
            for staro, novo in spec["mapa"].items():
                v = polje(r, staro)
                if novo == "Klasa":
                    v = _klasa(v)
                elif novo == "KolAmbalaze":
                    # Pisac uvek pise broj; prazno u starom redu znaci "bez gajbi".
                    b = broj(v)
                    v = b if b is not None else (0 if not tekst(v) else tekst(v))
                elif novo in ("BrutoKg", "PredlogCena"):
                    # Pisac pise bruto / predlog cene samo kad je unet i > 0
                    # (modDokUnos: PredlogCena > 0; MIG-14 OTP-00440).
                    b = broj(v)
                    v = b if b is not None and b > 0 else ""
                else:
                    b = broj(v)
                    v = b if b is not None else tekst(v)
                s[novo] = v
            for k in ("CreatedAt", "CreatedBy", "ModifiedAt", "ModifiedBy"):
                if k in s:
                    s[k] = polje(r, k)
            c.tabele[spec["stavke"]].append(s)
            c.stavka_za[(ime, kljuc(idn))] = sid
        _visak_kolona(c, t, list(kol_z) + list(spec["mapa"]) + ["Novac", "Isplaceno"])
        _veze_ispravke(c, ime, spec["pk"])


def _veze_ispravke(c, ime, pk):
    """IspravkaOdID / ZamenjenSaID moraju pokazivati na postojeci dokument."""
    ids = {kljuc(z[pk]) for z in c.tabele[ime]}
    for z in c.tabele[ime]:
        for k in ("IspravkaOdID", "ZamenjenSaID"):
            if k in z and tekst(z[k]) and kljuc(z[k]) not in ids:
                c.nalaz("UPOZORENJE", "VEZA_ISPRAVKE_BEZ_DOKUMENTA", ime, z[pk],
                        "%s='%s' ne postoji -- obrisano" % (k, z[k]))
                z[k] = ""


# ---------------------------------------------------------------------------
# Clanstvo: tblOtpremnicaIzvori i tblZbirnaIzvori
# ---------------------------------------------------------------------------

def _clanstvo(c, tabele):
    otk_src = tabela(tabele, "tblOtkup")
    otp_src = tabela(tabele, "tblOtpremnica")
    zbr_src = tabela(tabele, "tblZbirna")
    otp_ids = {kljuc(z["OtpremnicaID"]): z["OtpremnicaID"] for z in c.tabele["tblOtpremnica"]}
    if otk_src is not None:
        for r in otk_src.redovi:
            oid, pid = tekst(polje(r, "OtkupID")), tekst(polje(r, "OtpremnicaID"))
            if not pid:
                continue
            if kljuc(pid) not in otp_ids:
                c.nalaz("UPOZORENJE", "OTKUP_POKAZUJE_NA_NEPOSTOJECU_OTPREMNICU", "tblOtkup", oid,
                        "OtpremnicaID='%s' ne postoji -- veza se ne pravi" % pid)
                continue
            c.tabele["tblOtpremnicaIzvori"].append(_izvor(
                "OPI-", "OtpremnicaIzvorID", c.kanon["tblOtpremnicaIzvori"],
                OtpremnicaID=otp_ids[kljuc(pid)], OtkupID=oid,
                CreatedAt=polje(r, "CreatedAt"), CreatedBy=polje(r, "CreatedBy")))
    if otp_src is not None and zbr_src is not None:
        zbirne = list(zbr_src.redovi)
        for r in otp_src.redovi:
            pid, broj_zbr = tekst(polje(r, "OtpremnicaID")), tekst(polje(r, "BrojZbirne"))
            if not broj_zbr:
                continue
            zbr, kako = _nadji_zbirnu(broj_zbr, r, zbirne)
            if zbr is None:
                c.nalaz("UPOZORENJE", "ZBIRNA_PO_BROJU_NIJE_NADJENA", "tblOtpremnica", pid,
                        "BrojZbirne='%s': %s -- veza se ne pravi" % (broj_zbr, kako))
                continue
            c.tabele["tblZbirnaIzvori"].append(_izvor(
                "ZBI-", "ZbirnaIzvorID", c.kanon["tblZbirnaIzvori"],
                ZbirnaID=tekst(polje(zbr, "ZbirnaID")), OtpremnicaID=pid,
                CreatedAt=polje(r, "CreatedAt"), CreatedBy=polje(r, "CreatedBy")))


def _izvor(prefiks, pk, kolone, **vrednosti):
    red = {k: "" for k in kolone}
    red.update({k: v for k, v in vrednosti.items() if k in red})
    deo = [vrednosti.get(k, "") for k in ("OtpremnicaID", "OtkupID", "ZbirnaID")]
    red[pk] = novi_id(prefiks, *[kljuc(x) for x in deo])
    return red


def _nadji_zbirnu(broj_zbr, otp, zbirne):
    """MIG-11 lookup_zbr: jedinstven broj; ako se broj ponavlja -- vozac i datum.
    Protivrecan dokaz ne pogadja."""
    kand = [z for z in zbirne if kljuc(polje(z, "BrojZbirne")) == kljuc(broj_zbr)]
    if not kand:
        return None, "broj ne postoji"
    if len(kand) == 1:
        return kand[0], "jedinstven broj"
    voz, dan = kljuc(polje(otp, "VozacID")), datum(polje(otp, "Datum"))
    po_voz = [z for z in kand if voz and kljuc(polje(z, "VozacID")) == voz]
    po_dan = [z for z in kand if dan and datum(polje(z, "Datum")) == dan]
    oba = [z for z in po_voz if z in po_dan]
    if len(oba) == 1:
        return oba[0], "broj + vozac + datum"
    if len(oba) > 1:
        return None, "vise zbirnih istog broja, vozaca i datuma"
    if len(po_voz) == 1 and len(po_dan) == 1:
        if po_voz[0] is not po_dan[0]:
            return None, "vozac i datum pokazuju na razlicite zbirne"
        return po_voz[0], "broj + vozac i datum"
    if len(po_voz) == 1 and not po_dan:
        return po_voz[0], "broj + vozac"
    if len(po_dan) == 1 and not po_voz:
        vz = kljuc(polje(po_dan[0], "VozacID"))
        if voz and vz and vz != voz:
            return None, "vozac protivreci datumu"
        return po_dan[0], "broj + datum"
    return None, "broj se ponavlja, dokaz nije jednoznacan"


# ---------------------------------------------------------------------------
# Status, prijemnica -> zbirna, KupacHladnjaca
# ---------------------------------------------------------------------------

def aktivne_otpremnice_po_otkupu(c):
    aktivne = {kljuc(z["OtpremnicaID"]) for z in c.tabele["tblOtpremnica"] if not je_storno(_kf(z))}
    out = defaultdict(list)
    for v in c.tabele["tblOtpremnicaIzvori"]:
        if kljuc(v["OtpremnicaID"]) in aktivne:
            out[kljuc(v["OtkupID"])].append(v["OtpremnicaID"])
    return out


def _kf(red):
    """Red sa casefold kljucevima, za je_storno/polje."""
    return {k.casefold(): v for k, v in red.items()}


def _status_dokumenata(c):
    """Otpremnica i zbirna: IZDATO -- stari program nije imao nacrt (MIG-14);
    storno se ne dira, izdat pa storniran dokument je i dalje bio izdat."""
    for ime in ("tblOtpremnica", "tblZbirna"):
        for z in c.tabele[ime]:
            if "IzdatoStatus" in z:
                z["IzdatoStatus"] = IZDATO


def _stavke_po_klasi(c, ime, pk):
    out = defaultdict(lambda: defaultdict(lambda: [0, 0]))
    for s in c.tabele[ime]:
        kl = out[kljuc(s.get(pk))][tekst(s.get("Klasa"))]
        kl[0] += round(broj(s.get("Kolicina")) or 0, 2)
        kl[1] += round(broj(s.get("KolAmbalaze")) or 0, 2)
    return out


def _popravi_veze(c, bez_popravke):
    """MIG-21: aktivna izdata otpremnica bez aktivnog otkupa dobija vezu SAMO kad
    tacno jedan aktivan otkup ima isti broj, datum, stanicu, vrstu, sortu, tip
    ambalaze i iste kg i gajbe po klasi, a sam nije u aktivnoj otpremnici
    (njegova jedina veza, ako je ima, vodi na storniranu). Bez aproksimacije;
    para-potez iz MIG-21 se ne prenosi (u MIG-21 nikad nije radio)."""
    otp = {kljuc(z["OtpremnicaID"]): z for z in c.tabele["tblOtpremnica"]}
    otk = {kljuc(z["OtkupID"]): z for z in c.tabele["tblOtkup"]}
    akt_otp = {i for i, z in otp.items() if not je_storno(_kf(z))}
    akt_otk = {i for i, z in otk.items() if not je_storno(_kf(z))}
    veze_otk = defaultdict(list)
    izvori = defaultdict(list)
    for v in c.tabele["tblOtpremnicaIzvori"]:
        veze_otk[kljuc(v["OtkupID"])].append(v)
        izvori[kljuc(v["OtpremnicaID"])].append(kljuc(v["OtkupID"]))
    sk_otp = _stavke_po_klasi(c, "tblOtpremnicaStavke", "OtpremnicaID")
    sk_otk = _stavke_po_klasi(c, "tblOtkupStavke", "OtkupID")
    brojevi = defaultdict(list)
    for i in akt_otp:
        brojevi[kljuc(otp[i].get("BrojOtpremnice"))].append(i)
    polja = (("Datum", "Datum"), ("StanicaID", "StanicaID"), ("VrstaVoca", "VrstaVoca"),
             ("SortaVoca", "SortaVoca"), ("TipAmbalaze", "TipAmbalaze"))
    for pid in sorted(akt_otp):
        if any(o in akt_otk for o in izvori.get(pid, [])):
            continue
        z = otp[pid]
        if pid in bez_popravke:
            c.nalaz("INFO", "VEZA_SE_NE_POPRAVLJA_IZUZETAK", "tblOtpremnica", z["OtpremnicaID"],
                    tekst(bez_popravke[pid].get("razlog")))
            continue
        br = kljuc(z.get("BrojOtpremnice"))
        if not br or len(brojevi[br]) != 1:
            continue
        kand = []
        for oid in sorted(akt_otk):
            o = otk[oid]
            if kljuc(o.get("BrojDokumenta")) != br:
                continue
            if not all(kljuc(o.get(a)) and kljuc(o.get(a)) == kljuc(z.get(b)) for a, b in polja
                       if a != "Datum"):
                continue
            if datum(o.get("Datum")) is None or datum(o.get("Datum")) != datum(z.get("Datum")):
                continue
            if not sk_otk[oid] or dict(sk_otk[oid]) != dict(sk_otp[pid]):
                continue
            veze = veze_otk.get(oid, [])
            if len(veze) > 1 or any(kljuc(v["OtpremnicaID"]) in akt_otp for v in veze):
                continue
            kand.append((oid, veze[0] if veze else None))
        if len(kand) > 1:
            c.nalaz("UPOZORENJE", "VEZA_VISE_KANDIDATA", "tblOtpremnica", z["OtpremnicaID"],
                    "%d jednako dobrih otkupa -- veza se ne pogadja" % len(kand))
            continue
        if not kand:
            continue
        oid, veza = kand[0]
        if veza is not None:
            staro = veza["OtpremnicaID"]
            veza["OtpremnicaID"] = z["OtpremnicaID"]
            veza["OtpremnicaIzvorID"] = novi_id("OPI-", kljuc(z["OtpremnicaID"]), oid, "")
            opis = "premestena sa stornirane %s" % staro
        else:
            c.tabele["tblOtpremnicaIzvori"].append(_izvor(
                "OPI-", "OtpremnicaIzvorID", c.kanon["tblOtpremnicaIzvori"],
                OtpremnicaID=z["OtpremnicaID"], OtkupID=otk[oid]["OtkupID"], CreatedBy="MIGRACIJA"))
            opis = "dodata"
        izvori[pid].append(oid)
        c.nalaz("INFO", "VEZA_POPRAVLJENA_MIG21", "tblOtpremnica", z["OtpremnicaID"],
                "otkup %s %s: broj, datum, stanica, vrsta, sorta, ambalaza, kg i gajbe po klasi se slazu"
                % (otk[oid]["OtkupID"], opis))


def _otpremnice_bez_izvora(c):
    """Aktivna otpremnica bez ijednog aktivnog otkupa ne sme biti IZDATA: izvestaji
    robe bi pali na njoj (VrednostIzvoraZaOtpremnicu). Ostaje neizdata -- vidi
    se u listi, ne ulazi u izvestaje robe."""
    akt_otk = {kljuc(z["OtkupID"]) for z in c.tabele["tblOtkup"] if not je_storno(_kf(z))}
    izvori = defaultdict(list)
    for v in c.tabele["tblOtpremnicaIzvori"]:
        izvori[kljuc(v["OtpremnicaID"])].append(kljuc(v["OtkupID"]))
    for z in c.tabele["tblOtpremnica"]:
        if je_storno(_kf(z)) or z.get("IzdatoStatus") != IZDATO:
            continue
        if not any(o in akt_otk for o in izvori.get(kljuc(z["OtpremnicaID"]), [])):
            z["IzdatoStatus"] = ""
            c.nalaz("ODLUKA", "OTPREMNICA_BEZ_OTKUPA_OSTAJE_NEIZDATA", "tblOtpremnica", z["OtpremnicaID"],
                    "broj %s, %s: nijedan aktivan otkup se ne moze dokazati -- ne ulazi u izvestaje robe"
                    % (z.get("BrojOtpremnice", ""), _dan(z.get("Datum"))))


def _status_otkupa(c):
    po_otkupu = aktivne_otpremnice_po_otkupu(c)
    for z in c.tabele["tblOtkup"]:
        if je_storno(_kf(z)) or po_otkupu.get(kljuc(z["OtkupID"])):
            z["IzdatoStatus"] = IZDATO
        else:
            z["IzdatoStatus"] = ""
            c.nalaz("ODLUKA", "AKTIVAN_OTKUP_BEZ_OTPREMNICE", "tblOtkup", z["OtkupID"],
                    "datum %s, broj %s: ostaje NEIZDAT (sync ga inace sam stavlja u automatsku "
                    "otpremnicu); za nastavak u novoj sezoni treba odluka"
                    % (_dan(z.get("Datum")), z.get("BrojDokumenta", "")))


def _dan(v):
    d = datum(v)
    return d.isoformat() if d else tekst(v)


def _prijemnica_roditelj(c, tabele):
    """Kao pisac prijemnice (modDokumenta ~6922): aktivna izdata zbirna istog
    broja, kupca i vozaca. BrojZbirne ostaje -- mnogi citaoci biraju po njemu."""
    zbirne = [z for z in c.tabele["tblZbirna"] if not je_storno(_kf(z))]
    for p in c.tabele["tblPrijemnica"]:
        if "ZbirnaRoditeljID" not in p:
            continue
        p["ZbirnaRoditeljID"] = ""
        if not tekst(p.get("BrojZbirne")):
            continue
        kand = [z for z in zbirne
                if kljuc(z.get("BrojZbirne")) == kljuc(p.get("BrojZbirne"))
                and kljuc(z.get("KupacID")) == kljuc(p.get("KupacID"))
                and kljuc(z.get("VozacID")) == kljuc(p.get("VozacID"))]
        if len(kand) == 1:
            p["ZbirnaRoditeljID"] = kand[0]["ZbirnaID"]
        elif not je_storno(_kf(p)):
            c.nalaz("UPOZORENJE", "PRIJEMNICA_BEZ_JEDNOZNACNE_ZBIRNE", "tblPrijemnica",
                    p["PrijemnicaID"], "BrojZbirne='%s', kandidata: %d -- ZbirnaRoditeljID ostaje prazan"
                    % (p.get("BrojZbirne"), len(kand)))


def _config_vrednost(tabele, kljuc_cfg):
    for tbl, k, v in (("tblSEFConfig", "ConfigKey", "ConfigValue"), ("tblConfig", "Kljuc", "Vrednost")):
        t = tabela(tabele, tbl)
        if t is None:
            continue
        for r in t.redovi:
            if kljuc(polje(r, k)) == kljuc_cfg.casefold():
                return tekst(polje(r, v))
    return ""


def _kupac_hladnjaca(c, tabele):
    """ZBR-OWN-01: "Da" kad je kupac zbirne kupac hladnjace (MALINA_DEFAULT_KUPAC)."""
    h = _config_vrednost(tabele, "MALINA_DEFAULT_KUPAC")
    for z in c.tabele["tblZbirna"]:
        if "KupacHladnjaca" in z:
            z["KupacHladnjaca"] = "Da" if h and kljuc(z.get("KupacID")) == kljuc(h) else "Ne"
    if not h:
        c.nalaz("INFO", "NEMA_MALINA_DEFAULT_KUPAC", "tblZbirna", "",
                "u staroj svesci nema MALINA_DEFAULT_KUPAC -- sve zbirne dobijaju KupacHladnjaca=Ne")


def _stavke_nizvodno(c):
    """Paleta i faktura stavka dobijaju PrijemnicaStavkaID svoje prijemnice."""
    for ime in ("tblPaletaStavka", "tblFakturaStavke"):
        for r in c.tabele.get(ime, []):
            if "PrijemnicaStavkaID" not in r:
                continue
            pid = tekst(r.get("PrijemnicaID"))
            if not pid:
                continue
            sid = c.stavka_za.get(("tblPrijemnica", kljuc(pid)))
            if sid:
                r["PrijemnicaStavkaID"] = sid
            else:
                c.nalaz("BLOKIRA" if ime == "tblFakturaStavke" and not je_storno(_kf(r)) else "UPOZORENJE",
                        "STAVKA_BEZ_PRIJEMNICE", ime, r.get("StavkaID", ""),
                        "PrijemnicaID='%s' ne postoji u tblPrijemnica" % pid)


# ---------------------------------------------------------------------------
# Ambalaza
# ---------------------------------------------------------------------------

def _ambalaza(c, tabele, izuzeci, knjizenje):
    if tabela(tabele, "tblAmbalaza") is None:
        return
    rez = ambalaza.prevedi(tabele, izuzeci, knjizenje=knjizenje)
    c.amb = rez
    kol = c.kanon["tblAmbalaza"]
    for k in sorted(rez.kretanja, key=lambda x: (x.datum is None, x.datum or 0, x.redosled, x.generisano)):
        red = {x: "" for x in kol}
        red.update(AmbID=k.amb_id, Datum=k.datum, TipAmbalaze=k.tip, Kolicina=k.kolicina,
                   DokumentID=k.dok_id, DokumentTIP=k.dok_tip, CreatedAt=k.created_at,
                   CreatedBy=k.created_by or "MIGRACIJA", OdNalogTip=k.od_tip, OdNalogID=k.od_id,
                   NaNalogTip=k.na_tip, NaNalogID=k.na_id, VrstaKretanja=k.vrsta, StornoOd="")
        c.tabele["tblAmbalaza"].append({x: red[x] for x in kol})
    kol_d = c.kanon["tblAmbalazaDokument"]
    for z in sorted(rez.zaglavlja.values(), key=lambda x: (x["Datum"] is None, x["Datum"] or 0, x["AmbDokID"])):
        red = {x: "" for x in kol_d}
        red.update({x: z.get(x, "") for x in kol_d if x in z})
        red["CreatedBy"] = "MIGRACIJA"
        c.tabele["tblAmbalazaDokument"].append(red)
    for zz in rez.za_odluku:
        c.nalaz("ODLUKA", "AMBALAZA_" + zz["razlog"], "tblAmbalaza", "",
                "%d stara reda; %s" % (len(zz["noge"]), zz["napomena"]))
