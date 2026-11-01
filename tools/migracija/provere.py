"""Provere nove sveske PRE upisa -- ista pravila po kojima main CITA.

main cita tabelu dokumenta CELU: jedan los red obara svaki ekran, izvestaj i
storno koji tu tabelu cita, ne samo sebe (modOtkup.StavkeOtkupaRedovi,
modDokumenta.StavkeOtpremniceRedovi / StavkeZbirneRedovi / StavkePrijemniceRedovi,
AktivnoOtpClanstvoPoKanonu, AktivnoClanstvoPoKanonu, modFaktura.FakturaLinijePoStavci).
Zato se meri svaki red pre upisa.

Tezine:
  BLOKIRA     posle upisa bi main pao na citanju cele tabele -- upis se odbija
  ODLUKA      main ne pada, ali ishod zavisi od poslovne odluke operatera
  UPOZORENJE  akcija nad tim jednim dokumentom ne bi prosla (storno, stampa...)
  INFO        za znanje
"""

from collections import Counter, defaultdict
from decimal import Decimal

import ambalaza
from citac import broj, je_storno, kljuc, tekst
from cilj import DOKUMENTI, IZDATO, _kf

BLOKIRA, ODLUKA, UPOZORENJE, INFO = "BLOKIRA", "ODLUKA", "UPOZORENJE", "INFO"
KLASE = {"I", "II"}


def proveri(c):
    _stavke(c)
    _clanstvo(c)
    _izvori_otpremnice(c)
    _zbirne(c)
    _fakture(c)
    _novac(c)
    _palete(c)
    _hladnjaca(c)
    _ambalaza(c)
    return c.nalazi


def _ceo_broj(d):
    return d is not None and d == d.to_integral_value()


def _stavke(c):
    """Stavke*Redovi: zaglavlje tacno jednom, bar jedna stavka, jedna stavka po
    klasi, Klasa I/II, kolicina > 0, cena (otkup > 0, prijemnica >= 0), gajbe
    ceo broj >= 0, bruto prazan ili ispravan. Vazi i za STORNIRANA zaglavlja."""
    for ime, spec in DOKUMENTI.items():
        pk, stv = spec["pk"], spec["stavke"]
        ids = Counter(kljuc(z.get(pk)) for z in c.tabele[ime])
        for i, n in ids.items():
            if not i:
                c.nalaz(BLOKIRA, "ZAGLAVLJE_BEZ_ID", ime, "", "%d zaglavlja bez %s" % (n, pk))
            elif n > 1:
                c.nalaz(BLOKIRA, "DUPLO_ZAGLAVLJE", ime, i, "%s se ponavlja %d puta" % (pk, n))
        zag = {kljuc(z.get(pk)): z for z in c.tabele[ime]}
        po_zaglavlju = defaultdict(list)
        sids = Counter()
        for s in c.tabele[stv]:
            po_zaglavlju[kljuc(s.get(pk))].append(s)
            sids[kljuc(s.get(spec["spk"]))] += 1
            idn = s.get(pk)
            if kljuc(idn) not in zag:
                c.nalaz(BLOKIRA, "STAVKA_BEZ_ZAGLAVLJA", stv, idn, "")
            if tekst(s.get("Klasa")) not in KLASE:
                c.nalaz(BLOKIRA, "KLASA_NIJE_I_II", stv, idn, "Klasa='%s'" % tekst(s.get("Klasa")))
            kol = broj(s.get("Kolicina"))
            if kol is None or kol <= 0:
                storno = je_storno(_kf(zag.get(kljuc(idn), {})))
                c.nalaz(BLOKIRA, "KOLICINA_NIJE_POZITIVNA", stv, idn,
                        "Kolicina='%s'%s" % (tekst(s.get("Kolicina")),
                                             " (STORNIRAN dokument -- slucaj MIG-10)" if storno else ""))
            amb = broj(s.get("KolAmbalaze"))
            if amb is None or amb < 0 or not _ceo_broj(amb):
                c.nalaz(BLOKIRA, "GAJBE_NISU_CEO_BROJ", stv, idn, "KolAmbalaze='%s'" % tekst(s.get("KolAmbalaze")))
            if ime == "tblOtkup":
                cena = broj(s.get("Cena"))
                if cena is None or cena <= 0:
                    c.nalaz(BLOKIRA, "CENA_OTKUPA_NIJE_POZITIVNA", stv, idn, "Cena='%s'" % tekst(s.get("Cena")))
                bruto = broj(s.get("BrutoKg"))
                if bruto is not None and kol is not None and bruto < kol:
                    c.nalaz(BLOKIRA, "BRUTO_MANJI_OD_NETO", stv, idn, "BrutoKg=%s < Kolicina=%s" % (bruto, kol))
            if ime == "tblPrijemnica":
                cena = broj(s.get("Cena"))
                if cena is None or cena < 0:
                    c.nalaz(BLOKIRA, "CENA_PRIJEMNICE_NEISPRAVNA", stv, idn, "Cena='%s'" % tekst(s.get("Cena")))
                bruto = broj(s.get("BrutoKg"))
                if tekst(s.get("BrutoKg")) and (bruto is None or bruto <= 0):
                    c.nalaz(BLOKIRA, "BRUTO_NEISPRAVAN", stv, idn, "BrutoKg='%s'" % tekst(s.get("BrutoKg")))
        for i, n in sids.items():
            if not i or n > 1:
                c.nalaz(BLOKIRA, "ID_STAVKE_PRAZAN_ILI_DUPLI", stv, i, "")
        for i, z in zag.items():
            lst = po_zaglavlju.get(i, [])
            if not lst:
                c.nalaz(BLOKIRA, "ZAGLAVLJE_BEZ_STAVKE", ime, z.get(pk), "")
            klase = Counter(tekst(s.get("Klasa")) for s in lst)
            if any(n > 1 for n in klase.values()):
                c.nalaz(BLOKIRA, "DVE_STAVKE_ISTE_KLASE", ime, z.get(pk), "")


def _aktivni(c, ime, pk):
    return {kljuc(z.get(pk)) for z in c.tabele[ime] if not je_storno(_kf(z))}


def _clanstvo(c):
    """AktivnoOtpClanstvoPoKanonu / AktivnoClanstvoPoKanonu: veze na postojece
    dokumente, bez duplih parova, otkup u najvise jednoj aktivnoj otpremnici,
    otpremnica u najvise jednoj aktivnoj zbirnoj."""
    otk = {kljuc(z.get("OtkupID")) for z in c.tabele["tblOtkup"]}
    otp = {kljuc(z.get("OtpremnicaID")) for z in c.tabele["tblOtpremnica"]}
    zbr = {kljuc(z.get("ZbirnaID")) for z in c.tabele["tblZbirna"]}
    akt_otp = _aktivni(c, "tblOtpremnica", "OtpremnicaID")
    akt_zbr = _aktivni(c, "tblZbirna", "ZbirnaID")
    for ime, roditelj, dete, roditelji, deca, aktivni in (
            ("tblOtpremnicaIzvori", "OtpremnicaID", "OtkupID", otp, otk, akt_otp),
            ("tblZbirnaIzvori", "ZbirnaID", "OtpremnicaID", zbr, otp, akt_zbr)):
        parovi = Counter()
        u_aktivnom = defaultdict(set)
        for v in c.tabele[ime]:
            r, d = kljuc(v.get(roditelj)), kljuc(v.get(dete))
            if not r or not d or r not in roditelji or d not in deca:
                c.nalaz(BLOKIRA, "CLANSTVO_NA_NEPOSTOJECI_DOKUMENT", ime, v.get(dete),
                        "%s='%s', %s='%s'" % (roditelj, v.get(roditelj), dete, v.get(dete)))
            parovi[(r, d)] += 1
            if r in aktivni:
                u_aktivnom[d].add(r)
        for (r, d), n in parovi.items():
            if n > 1:
                c.nalaz(BLOKIRA, "ISTI_PAR_CLANSTVA_DVAPUT", ime, d, "%s/%s" % (r, d))
        for d, rs in u_aktivnom.items():
            if len(rs) > 1:
                c.nalaz(BLOKIRA, "CLAN_DVA_AKTIVNA_RODITELJA", ime, d, ", ".join(sorted(rs)))


def _izvori_otpremnice(c):
    """Izdata aktivna otpremnica bez izvora obara izvestaje robe
    (VrednostIzvoraZaOtpremnicu); klasa koje izvori nemaju obara stampu."""
    izvori = defaultdict(list)
    for v in c.tabele["tblOtpremnicaIzvori"]:
        izvori[kljuc(v.get("OtpremnicaID"))].append(kljuc(v.get("OtkupID")))
    akt_otk = _aktivni(c, "tblOtkup", "OtkupID")
    klase_otk = defaultdict(set)
    for s in c.tabele["tblOtkupStavke"]:
        klase_otk[kljuc(s.get("OtkupID"))].add(tekst(s.get("Klasa")))
    klase_otp = defaultdict(set)
    for s in c.tabele["tblOtpremnicaStavke"]:
        klase_otp[kljuc(s.get("OtpremnicaID"))].add(tekst(s.get("Klasa")))
    for z in c.tabele["tblOtpremnica"]:
        if je_storno(_kf(z)) or tekst(z.get("IzdatoStatus")) != IZDATO:
            continue
        pid = kljuc(z.get("OtpremnicaID"))
        aktivni = [o for o in izvori.get(pid, []) if o in akt_otk]
        if not aktivni:
            c.nalaz(BLOKIRA, "IZDATA_OTPREMNICA_BEZ_AKTIVNOG_OTKUPA", "tblOtpremnica", z.get("OtpremnicaID"),
                    "broj %s: izvestaji robe padaju na njoj (VrednostIzvoraZaOtpremnicu); "
                    "veza se mora dokazati ili otpremnica ostati neizdata" % z.get("BrojOtpremnice", ""))
            continue
        imaju = set().union(*(klase_otk[o] for o in aktivni))
        if klase_otp[pid] - imaju:
            c.nalaz(UPOZORENJE, "KLASA_OTPREMNICE_BEZ_IZVORA", "tblOtpremnica", z.get("OtpremnicaID"),
                    "klase %s nema u izvorima -- stampa otpremnice pada" % sorted(klase_otp[pid] - imaju))


def _zbirne(c):
    """Zbirna bez clanova i aktivna zbirna sa storniranim/neizdatim clanom
    obaraju akcije nad TOM zbirnom (IzvoriZbirne, ZbrClanoviPoStanju); zbir
    je samo izvestaj (ValidateZbirna), ne pada citanje."""
    clanovi = defaultdict(list)
    for v in c.tabele["tblZbirnaIzvori"]:
        clanovi[kljuc(v.get("ZbirnaID"))].append(kljuc(v.get("OtpremnicaID")))
    otp = {kljuc(z.get("OtpremnicaID")): z for z in c.tabele["tblOtpremnica"]}
    kol_otp = defaultdict(lambda: defaultdict(Decimal))
    for s in c.tabele["tblOtpremnicaStavke"]:
        kol_otp[kljuc(s.get("OtpremnicaID"))][tekst(s.get("Klasa"))] += broj(s.get("Kolicina")) or 0
    kol_zbr = defaultdict(lambda: defaultdict(Decimal))
    for s in c.tabele["tblZbirnaStavke"]:
        kol_zbr[kljuc(s.get("ZbirnaID"))][tekst(s.get("Klasa"))] += broj(s.get("Kolicina")) or 0
    for z in c.tabele["tblZbirna"]:
        zid = kljuc(z.get("ZbirnaID"))
        lst = clanovi.get(zid, [])
        storno = je_storno(_kf(z))
        if not lst:
            c.nalaz(UPOZORENJE, "ZBIRNA_BEZ_CLANOVA", "tblZbirna", z.get("ZbirnaID"),
                    "broj %s%s: storno/akcije nad njom padaju (IzvoriZbirne)"
                    % (z.get("BrojZbirne", ""), " (stornirana)" if storno else ""))
            continue
        if storno:
            continue
        losi = [p for p in lst if p in otp and (je_storno(_kf(otp[p])) or
                                                tekst(otp[p].get("IzdatoStatus")) != IZDATO)]
        if losi:
            c.nalaz(UPOZORENJE, "AKTIVNA_ZBIRNA_SA_STORNIRANIM_CLANOM", "tblZbirna", z.get("ZbirnaID"),
                    "clanovi %s -- storno tokovi nad zbirnom padaju (ZbrClanoviPoStanju)"
                    % ", ".join(otp[p].get("OtpremnicaID") for p in losi))
        zbir = defaultdict(Decimal)
        for p in lst:
            if p in otp and not je_storno(_kf(otp[p])):
                for k, v in kol_otp[p].items():
                    zbir[k] += v
        if dict(zbir) != dict(kol_zbr[zid]):
            c.nalaz(INFO, "ZBIRNA_RAZLICITA_OD_ZBIRA_OTPREMNICA", "tblZbirna", z.get("ZbirnaID"),
                    "zbirna %s, zbir aktivnih otpremnica %s" % (dict(kol_zbr[zid]), dict(zbir)))


def _fakture(c):
    """FakturaLinijePoStavci: linija prijemnice bez PrijemnicaStavkaID-a,
    faktura bez ID-a ili dupla, aktivna linija stornirane fakture -- pada."""
    fak = Counter(kljuc(f.get("FakturaID")) for f in c.tabele.get("tblFakture", []))
    for i, n in fak.items():
        if not i or n > 1:
            c.nalaz(BLOKIRA, "FAKTURA_ID_PRAZAN_ILI_DUPLI", "tblFakture", i, "")
    storno_fak = {kljuc(f.get("FakturaID")) for f in c.tabele.get("tblFakture", []) if je_storno(_kf(f))}
    for s in c.tabele.get("tblFakturaStavke", []):
        if je_storno(_kf(s)):
            continue
        if kljuc(s.get("FakturaID")) not in fak:
            c.nalaz(BLOKIRA, "FAKTURA_STAVKA_BEZ_FAKTURE", "tblFakturaStavke", s.get("StavkaID"), "")
        if kljuc(s.get("FakturaID")) in storno_fak:
            c.nalaz(BLOKIRA, "AKTIVNA_STAVKA_STORNIRANE_FAKTURE", "tblFakturaStavke", s.get("StavkaID"), "")
        if tekst(s.get("PrijemnicaID")) and not tekst(s.get("PrijemnicaStavkaID")):
            c.nalaz(BLOKIRA, "FAKTURA_STAVKA_BEZ_STAVKE_PRIJEMNICE", "tblFakturaStavke", s.get("StavkaID"),
                    "PrijemnicaID='%s'" % s.get("PrijemnicaID"))
    for f in c.tabele.get("tblFakture", []):
        b = tekst(f.get("BrojFakture"))
        if b and not _broj_fakture_ok(b):
            c.nalaz(UPOZORENJE, "BROJ_FAKTURE_NIJE_N_GGGG", "tblFakture", f.get("FakturaID"),
                    "BrojFakture='%s' -- nova numeracija iste godine bi krenula od 1" % b)


def _broj_fakture_ok(b):
    deo = b.split("/")
    return len(deo) == 2 and deo[0].strip().isdigit() and deo[1].strip().isdigit() and len(deo[1].strip()) == 4


AVANS_TIPOVI = {"virmanfirmakoop", "virmanavanskoop"}


def _novac(c):
    """Isplata kooperantu bez OtkupID-a je za main NEISKORISCEN AVANS: prvi
    sledeci otkup tog kooperanta ga sam potrosi (modNovac.JeAvansKooperanta).
    Otkup je placen samo preko tblNovac.Isplata sa njegovim OtkupID-om; stari
    Otkup.Isplaceno se vise ne cita."""
    otk = {kljuc(z.get("OtkupID")) for z in c.tabele["tblOtkup"]}
    for n in c.tabele.get("tblNovac", []):
        if je_storno(_kf(n)):
            continue
        isplata = broj(n.get("Isplata")) or 0
        if isplata > 0 and kljuc(n.get("Tip")) in AVANS_TIPOVI and not tekst(n.get("OtkupID")):
            c.nalaz(ODLUKA, "ISPLATA_BEZ_OTKUPA_POSTAJE_AVANS", "tblNovac", n.get("NovacID"),
                    "%s %s, kooperant %s: bez OtkupID-a se trosi na prvi sledeci otkup"
                    % (n.get("Tip"), isplata, n.get("KooperantID")))
        if tekst(n.get("OtkupID")) and kljuc(n.get("OtkupID")) not in otk:
            c.nalaz(UPOZORENJE, "NOVAC_NA_NEPOSTOJECI_OTKUP", "tblNovac", n.get("NovacID"),
                    "OtkupID='%s'" % n.get("OtkupID"))


def proveri_isplate_iz_starog(c, tabele):
    """Stari Otkup.Isplaceno / Novac bez reda u tblNovac: u novoj svesci otkup
    ostaje OTVOREN (neplacen), jer main placanje cita samo iz tblNovac."""
    from citac import polje, tabela
    t = tabela(tabele, "tblOtkup")
    if t is None:
        return
    placeni = {kljuc(n.get("OtkupID")) for n in c.tabele.get("tblNovac", [])
               if not je_storno(_kf(n)) and (broj(n.get("Isplata")) or 0) > 0}
    for r in t.redovi:
        oid = tekst(polje(r, "OtkupID"))
        if je_storno(r) or kljuc(oid) in placeni:
            continue
        isp = kljuc(polje(r, "Isplaceno"))
        novac = broj(polje(r, "Novac")) or 0
        if isp in ("da", "true", "1", "yes") or novac > 0:
            c.nalaz(ODLUKA, "STARI_ISPLACENO_BEZ_NOVCA", "tblOtkup", oid,
                    "staro Isplaceno='%s', Novac=%s, a u tblNovac nema isplate -- u novoj svesci "
                    "otkup je OTVOREN" % (tekst(polje(r, "Isplaceno")), novac))


def _palete(c):
    """Otvorena paleta stare sezone bi se punila gajbicama nove sezone
    (modPaletniList, ista vrsta/sorta/klasa/tip)."""
    for p in c.tabele.get("tblPaleta", []):
        if je_storno(_kf(p)) or kljuc(p.get("Preradjeno")) == "da":
            continue
        if kljuc(p.get("Status")) == "otvorena":
            c.nalaz(ODLUKA, "OTVORENA_PALETA", "tblPaleta", p.get("PaletaID"),
                    "broj %s: nova sezona bi je punila -- zatvoriti pri migraciji?" % p.get("BrojPalete", ""))
    prj = {kljuc(z.get("PrijemnicaID")) for z in c.tabele["tblPrijemnica"]}
    klasa_prj = {kljuc(s.get("PrijemnicaID")): tekst(s.get("Klasa")) for s in c.tabele["tblPrijemnicaStavke"]}
    for s in c.tabele.get("tblPaletaStavka", []):
        if je_storno(_kf(s)):
            continue
        pid = kljuc(s.get("PrijemnicaID"))
        if pid and pid in prj and tekst(s.get("Klasa")) and tekst(s.get("Klasa")).upper() != klasa_prj.get(pid):
            c.nalaz(UPOZORENJE, "KLASA_PALETE_NIJE_KLASA_PRIJEMNICE", "tblPaletaStavka", s.get("StavkaID"),
                    "paleta %s, prijemnica %s" % (s.get("Klasa"), klasa_prj.get(pid)))


def _hladnjaca(c):
    for z in c.tabele["tblZbirna"]:
        if "KupacHladnjaca" in z and z["KupacHladnjaca"] not in ("Da", "Ne"):
            c.nalaz(BLOKIRA, "KUPAC_HLADNJACA_NIJE_DA_NE", "tblZbirna", z.get("ZbirnaID"), "")


def _ambalaza(c):
    """Knjiga ambalaze po pravilima main-a posle #445-#450 (MIG-ODL-13..15).

    Invarijante knjige (AMB-INV-01/02/04/10/11/12/15, AMB-BIL-02), zatvaranje
    salda i pokrivenost su greska ALATA ako padnu -- BLOKIRA. U rezimu ispravno i:
    nijedan realan nalog ni jednog dana ispod nule (AMB-INV-07) i prazne
    stanice/vozaca nijednog dana ispod nule (AMB-INV-16, ODL-37): inace bi prvo
    izdavanje u novoj sezoni palo. PUNE ispod nule (nalog je poslao vise punih
    nego sto je primio) knjiga ne moze da popravi nabavkom -- ODLUKA (KI-065).
    """
    rez = c.amb
    if rez is None:
        return
    for x in ambalaza.proveri_invarijante(rez):
        c.nalaz(BLOKIRA, "AMB_" + x["inv"].replace("-", "_"), "tblAmbalaza", x.get("amb_id", ""), x["opis"])
    for kljuc_, lst in ambalaza.proveri_pokrivenost(rez).items():
        if lst:
            c.nalaz(BLOKIRA, "AMB_POKRIVENOST", "tblAmbalaza", "", "%s: %d starih redova" % (kljuc_, len(lst)))
    ispravno = rez.knjizenje == ambalaza.ISPRAVNO
    for x in ambalaza.izracunaj_salda(rez):
        nalog = "%s %s '%s'" % (x["nalog_tip"], x["nalog_id"], x["tip"])
        if x["razlika"]:
            c.nalaz(BLOKIRA, "AMB_SALDA_SE_NE_ZATVARAJU", "tblAmbalaza", "", "%s: razlika %s" % (nalog, x["razlika"]))
        if not ispravno or x["nalog_tip"] not in ambalaza.REALNI:
            continue
        if x["najmanji"] < 0:
            c.nalaz(BLOKIRA, "AMB_INV07_NALOG_ISPOD_NULE", "tblAmbalaza", "",
                    "%s: %s na dan %s" % (nalog, x["najmanji"], x["najmanji_dan"]))
        if x["najmanje_prazne"] is not None and x["najmanje_prazne"] < 0:
            c.nalaz(BLOKIRA, "AMB_INV16_PRAZNE_ISPOD_NULE", "tblAmbalaza", "",
                    "%s: prazne %s na dan %s" % (nalog, x["najmanje_prazne"], x["dan_prazne"]))
        if x["najmanje_pune"] is not None and x["najmanje_pune"] < 0:
            c.nalaz(ODLUKA, "AMB_PUNE_ISPOD_NULE", "tblAmbalaza", "",
                    "%s: pune %s na dan %s -- poslao je vise punih nego sto je primio, pa su mu "
                    "prazne za toliko iznad stvarnih (KI-065)" % (nalog, x["najmanje_pune"], x["dan_pune"]))
