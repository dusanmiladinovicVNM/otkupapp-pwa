"""Cista izmisljena stara sveska (C002 oblik, malina mod) za testove provere.

Pravila po kojima je sastavljena su pravila stare aplikacije (vba-v2.28.4):
  * otpremnica "x/ddmmyy[-n]", x = broj stanice; zbirna = "S" + broj otpremnice,
    VozacID = mirror stanice (ST-...), KupacID = MALINA_DEFAULT_KUPAC;
  * prijemnica "1/ddmmyy[-n]" (niz po kupcu i danu), BrojZbirne svoje zbirne,
    VozacID i KupacID kao na zbirnoj, isti dan;
  * otkup -> otpremnica po OtpremnicaID, otkup nosi BrojZbirne i VozacID otpremnice;
  * novac -> otkup po OtkupID; ambalaza -> dokument po DokumentTIP + DokumentID;
  * paletna stavka -> prijemnica po PrijemnicaID (+ BrojPrijemnice, BrojZbirne);
  * jedan red = jedan dokument (Klasa I), kao sto je izmereno u C002.

NAD OVOM SVESKOM NIJEDNA PROVERA NE SME DA PRIJAVI NIJEDAN NALAZ. Test u
test_zajednicko.py to drzi; svaki modul u svom testu polazi od nje, podmetne
gresku i trazi svoj kod.
"""

import copy
from datetime import date
from decimal import Decimal

import pregled

D5 = date(2026, 7, 5)
D6 = date(2026, 7, 6)
KUPAC = "KUP-00001"
CENA_OTK = 250
CENA_PRJ = 300

KOLONE = {
    "tblStanice": ["StanicaID", "Naziv", "Mesto", "Aktivan", "JeHladnjaca"],
    "tblVozaci": ["VozacID", "Ime", "Prezime", "Telefon", "KapacitetKG", "Aktivan"],
    "tblKupci": ["KupacID", "Naziv", "Mesto", "Aktivan", "Hladnjaca", "TekuciRacun"],
    "tblKooperanti": ["KooperantID", "Ime", "Prezime", "StanicaID", "Mesto", "Telefon", "BPGBroj",
                      "Adresa", "JMBG", "TekuciRacun", "Aktivan"],
    "tblSEFConfig": ["ConfigKey", "ConfigValue", "Opis", "Aktivan"],
    "tblOtkup": ["OtkupID", "Datum", "KooperantID", "StanicaID", "KulturaID", "VrstaVoca", "SortaVoca",
                 "Kolicina", "Cena", "TipAmbalaze", "KolAmbalaze", "VozacID", "BrojDokumenta", "Novac",
                 "PrimalacNovca", "Klasa", "Stornirano", "BrojZbirne", "Isplaceno", "DatumIsplate",
                 "OtpremnicaID", "ParcelaID", "ClientRecordID", "SyncSource", "KolAmbIzdata", "VremeUnosa",
                 "BrutoKg", "BrojOtpremnice", "CreatedAt", "CreatedBy", "ModifiedAt", "ModifiedBy"],
    "tblOtpremnica": ["OtpremnicaID", "Datum", "StanicaID", "VozacID", "BrojOtpremnice", "BrojZbirne",
                      "VrstaVoca", "SortaVoca", "Kolicina", "Cena", "TipAmbalaze", "KolAmbalaze", "Klasa",
                      "Stornirano", "BrutoKg", "CreatedAt", "CreatedBy", "ModifiedAt", "ModifiedBy"],
    "tblZbirna": ["ZbirnaID", "Datum", "VozacID", "BrojZbirne", "KupacID", "Hladnjaca", "Pogon", "VrstaVoca",
                  "SortaVoca", "UkupnoKolicina", "TipAmbalaze", "UkupnoAmbalaze", "Klasa", "Stornirano",
                  "ClientRecordID", "SyncSource", "CreatedAt", "CreatedBy", "ModifiedAt", "ModifiedBy"],
    "tblPrijemnica": ["PrijemnicaID", "Datum", "KupacID", "VozacID", "BrojPrijemnice", "BrojZbirne",
                      "VrstaVoca", "SortaVoca", "Kolicina", "Cena", "TipAmbalaze", "KolAmbalaze",
                      "KolAmbVracena", "Klasa", "Fakturisano", "FakturaID", "Stornirano", "BrutoKg",
                      "CreatedAt", "CreatedBy", "ModifiedAt", "ModifiedBy"],
    "tblNovac": ["NovacID", "BrojDokumenta", "Datum", "Partner", "PartnerID", "EntitetTip", "OMID",
                 "KooperantID", "FakturaID", "VrstaVoca", "Tip", "Uplata", "Isplata", "Napomena", "Stornirano",
                 "OtkupID", "OsirocenoOd", "CreatedAt", "CreatedBy", "ModifiedAt", "ModifiedBy"],
    "tblAmbalaza": ["AmbID", "Datum", "TipAmbalaze", "Kolicina", "Smer", "EntitetID", "EntitetTip", "VozacID",
                    "DokumentID", "DokumentTIP", "Stornirano", "CreatedAt", "CreatedBy"],
    "tblPaleta": ["PaletaID", "BrojPalete", "Godina", "Datum", "VrstaVoca", "SortaVoca", "Klasa", "TipAmbalaze",
                  "TipPalete", "KapacitetGajbica", "BrojGajbica", "NetoKg", "AmbalazaKg", "PaletaKg", "BrutoKg",
                  "Status", "Preradjeno", "CreatedAt", "Stornirano", "Istorija"],
    "tblPaletaStavka": ["StavkaID", "PaletaID", "PrijemnicaID", "BrojPrijemnice", "BrojZbirne", "Klasa",
                        "VrstaVoca", "SortaVoca", "BrojGajbica", "NetoKg", "AmbalazaKg", "CreatedAt", "Stornirano"],
    "tblFakture": ["FakturaID", "BrojFakture", "Datum", "KupacID", "Iznos", "Status", "DatumPlacanja",
                   "Stornirano"],
    "tblFakturaStavke": ["StavkaID", "FakturaID", "PrijemnicaID", "Kolicina", "Cena", "Klasa", "BrojPrijemnice",
                         "Stornirano", "OsirocenoOd"],
    "tblStornoVeze": ["CorrectionID", "Mode", "Status", "OldDocType", "OldDocID", "OldBroj", "NewDocType",
                      "NewDocID", "NewBroj", "ParentDocType", "ParentDocID", "ParentBroj", "CreatedAt",
                      "CreatedBy", "CompletedAt", "Message", "NeedsRecovery", "RecoveryAction"],
    "tblBankaImport": ["BankaImportID", "BrojRacuna", "BrojDokumenta", "DatumTransakcije", "Partner", "PartnerKonto",
                       "PozivNaBroj", "Isplata", "Uplata", "Obradjeno", "Stornirano"],
}

# (OTP, broj, stanica, dan, kg, gajbe, [(OTK, kooperant, kg, gajbe, broj otkupa)], ZBR, PRJ, broj prij., kg prij.)
LANCI = [
    ("OTP-00001", "1/050726", "ST-00001", D5, 100, 40,
     [("OTK-00001", "KOOP-00001", 60, 24, "1/050726"), ("OTK-00002", "KOOP-00002", 40, 16, "1/050726-2")],
     "ZBR-00001", "PRJ-00001", "1/050726", Decimal("99")),
    ("OTP-00002", "1/050726-2", "ST-00001", D5, 50, 20,
     [("OTK-00003", "KOOP-00001", 50, 20, "1/050726-3")],
     "ZBR-00002", "PRJ-00002", "1/050726-2", Decimal("49.5")),
    ("OTP-00003", "2/050726", "ST-00002", D5, 80, 32,
     [("OTK-00004", "KOOP-00003", 80, 32, "2/050726")],
     "ZBR-00003", "PRJ-00003", "1/050726-3", Decimal("79")),
    ("OTP-00004", "1/060726", "ST-00001", D6, 70, 28,
     [("OTK-00005", "KOOP-00002", 70, 28, "1/060726")],
     "ZBR-00004", "PRJ-00004", "1/060726", Decimal("69.3")),
    ("OTP-00005", "3/060726", "ST-00003", D6, 90, 36,
     [("OTK-00006", "KOOP-00004", 90, 36, "3/060726")],
     "ZBR-00005", "PRJ-00005", "1/060726-2", Decimal("89")),
]


def _ts(dan, minut):
    return "%s %02d:%02d:00" % (dan.isoformat(), 8 + minut // 60, minut % 60)


def osnovna():
    """dict ime_tabele -> list[dict kolona -> vrednost] (cist C002 oblik)."""
    t = {ime: [] for ime in KOLONE}
    t["tblStanice"] = [dict(StanicaID="ST-%05d" % i, Naziv="Stanica %d" % i, Mesto="Mesto", Aktivan="Da",
                            JeHladnjaca="") for i in (1, 2, 3)]
    t["tblVozaci"] = [dict(VozacID="ST-%05d" % i, Ime="Stanica %d" % i, Aktivan="Da") for i in (1, 2, 3)]
    t["tblVozaci"].append(dict(VozacID="VOZ-00001", Ime="Vlada", Prezime="Vozac", Aktivan="Da"))
    t["tblKupci"] = [dict(KupacID=KUPAC, Naziv="Hladnjaca Test", Aktivan="Da", Hladnjaca="Da")]
    t["tblKooperanti"] = [
        dict(KooperantID="KOOP-00001", Ime="Pera", Prezime="Peric", StanicaID="ST-00001", Aktivan="Da"),
        dict(KooperantID="KOOP-00002", Ime="Mika", Prezime="Mikic", StanicaID="ST-00001", Aktivan="Da"),
        dict(KooperantID="KOOP-00003", Ime="Zika", Prezime="Zikic", StanicaID="ST-00002", Aktivan="Da"),
        dict(KooperantID="KOOP-00004", Ime="Laza", Prezime="Lazic", StanicaID="ST-00003", Aktivan="Da"),
    ]
    t["tblSEFConfig"] = [dict(ConfigKey="MALINA_DEFAULT_KUPAC", ConfigValue=KUPAC),
                         dict(ConfigKey="MALINA_MODE", ConfigValue="YES")]
    amb = []
    minut = 0
    for (otp, bro, st, dan, kg, gaj, otkupi, zbr, prj, brp, kgp) in LANCI:
        bz = "S" + bro
        for (otk, koop, okg, ogaj, bdok) in otkupi:
            minut += 3
            t["tblOtkup"].append(dict(
                OtkupID=otk, Datum=dan, KooperantID=koop, StanicaID=st, KulturaID="KUL-00001",
                VrstaVoca="Malina", SortaVoca="Vilamet", Kolicina=okg, Cena=CENA_OTK, TipAmbalaze="12/1",
                KolAmbalaze=ogaj, VozacID=st, BrojDokumenta=bdok, Novac=0, PrimalacNovca="", Klasa="I",
                Stornirano="", BrojZbirne=bz, Isplaceno="", DatumIsplate="", OtpremnicaID=otp,
                KolAmbIzdata=0, BrutoKg=okg, BrojOtpremnice=bro, CreatedAt=_ts(dan, minut), CreatedBy="op",
                ModifiedAt=_ts(dan, minut), ModifiedBy="op"))
            amb.append((dan, ogaj, "Izlaz", koop, "Kooperant", "", otk, "Otkup"))
            amb.append((dan, ogaj, "Ulaz", st, "Stanica", "", otk, "Otkup"))
        minut += 5
        t["tblOtpremnica"].append(dict(
            OtpremnicaID=otp, Datum=dan, StanicaID=st, VozacID=st, BrojOtpremnice=bro, BrojZbirne=bz,
            VrstaVoca="Malina", SortaVoca="Vilamet", Kolicina=kg, Cena=0, TipAmbalaze="12/1", KolAmbalaze=gaj,
            Klasa="I", Stornirano="", BrutoKg=kg, CreatedAt=_ts(dan, minut), CreatedBy="op",
            ModifiedAt=_ts(dan, minut), ModifiedBy="op"))
        amb.append((dan, gaj, "Izlaz", st, "Stanica", st, otp, "Otpremnica"))
        t["tblZbirna"].append(dict(
            ZbirnaID=zbr, Datum=dan, VozacID=st, BrojZbirne=bz, KupacID=KUPAC, Hladnjaca="", Pogon="",
            VrstaVoca="Malina", SortaVoca="Vilamet", UkupnoKolicina=kg, TipAmbalaze="12/1", UkupnoAmbalaze=gaj,
            Klasa="I", Stornirano="", CreatedAt=_ts(dan, minut), CreatedBy="op", ModifiedAt=_ts(dan, minut),
            ModifiedBy="op"))
        minut += 30
        t["tblPrijemnica"].append(dict(
            PrijemnicaID=prj, Datum=dan, KupacID=KUPAC, VozacID=st, BrojPrijemnice=brp, BrojZbirne=bz,
            VrstaVoca="Malina", SortaVoca="Vilamet", Kolicina=kgp, Cena=CENA_PRJ, TipAmbalaze="12/1",
            KolAmbalaze=gaj, KolAmbVracena=gaj, Klasa="I", Fakturisano="", FakturaID="", Stornirano="",
            BrutoKg=kgp, CreatedAt=_ts(dan, minut), CreatedBy="op", ModifiedAt=_ts(dan, minut), ModifiedBy="op"))
        amb.append((dan, gaj, "Ulaz", KUPAC, "Kupac", st, prj, "Prijemnica"))
        amb.append((dan, gaj, "Izlaz", KUPAC, "Kupac", st, prj, "Prijemnica"))
    # novac: kes uz OTK-00001 pri unosu, virman za OTK-00004
    t["tblNovac"] = [
        dict(NovacID="NOV-00001", BrojDokumenta="1/050726", Datum=D5, Partner="Pera Peric", PartnerID="KOOP-00001",
             EntitetTip="Kooperant", OMID="ST-00001", KooperantID="KOOP-00001", VrstaVoca="Malina",
             Tip="KesOtkupacKoop", Uplata=0, Isplata=60 * CENA_OTK, Stornirano="", OtkupID="OTK-00001",
             CreatedAt=_ts(D5, 3), CreatedBy="op"),
        dict(NovacID="NOV-00002", BrojDokumenta="V-1", Datum=D6, Partner="Zika Zikic", PartnerID="KOOP-00003",
             EntitetTip="Kooperant", OMID="ST-00002", KooperantID="KOOP-00003", VrstaVoca="Malina",
             Tip="VirmanFirmaKoop", Uplata=0, Isplata=80 * CENA_OTK, Stornirano="", OtkupID="OTK-00004",
             CreatedAt=_ts(D6, 200), CreatedBy="op"),
    ]
    for r in t["tblOtkup"]:
        if r["OtkupID"] == "OTK-00001":
            r.update(Novac=60 * CENA_OTK, PrimalacNovca="Pera Peric", Isplaceno="Da", DatumIsplate=D5)
        if r["OtkupID"] == "OTK-00004":
            r.update(Isplaceno="Da", DatumIsplate=D6)
    for i, (dan, kol, smer, ent, etip, voz, dok, dtip) in enumerate(amb, 1):
        t["tblAmbalaza"].append(dict(AmbID="AMB-%05d" % i, Datum=dan, TipAmbalaze="12/1", Kolicina=kol, Smer=smer,
                                     EntitetID=ent, EntitetTip=etip, VozacID=voz, DokumentID=dok, DokumentTIP=dtip,
                                     Stornirano="", CreatedAt="", CreatedBy="op"))
    # paleta od PRJ-00001 i PRJ-00002 (ista klasa, isti kupac)
    t["tblPaleta"] = [dict(PaletaID="PAL-00001", BrojPalete="1", Godina=2026, Datum=D5, VrstaVoca="Malina",
                           SortaVoca="Vilamet", Klasa="I", TipAmbalaze="12/1", TipPalete="EUR", KapacitetGajbica=120,
                           BrojGajbica=60, NetoKg=Decimal("148.5"), Status="Zatvorena", Preradjeno="",
                           CreatedAt=_ts(D5, 300), Stornirano="")]
    t["tblPaletaStavka"] = [
        dict(StavkaID="PLS-00001", PaletaID="PAL-00001", PrijemnicaID="PRJ-00001", BrojPrijemnice="1/050726",
             BrojZbirne="S1/050726", Klasa="I", VrstaVoca="Malina", SortaVoca="Vilamet", BrojGajbica=40,
             NetoKg=Decimal("99"), CreatedAt=_ts(D5, 301), Stornirano=""),
        dict(StavkaID="PLS-00002", PaletaID="PAL-00001", PrijemnicaID="PRJ-00002", BrojPrijemnice="1/050726-2",
             BrojZbirne="S1/050726-2", Klasa="I", VrstaVoca="Malina", SortaVoca="Vilamet", BrojGajbica=20,
             NetoKg=Decimal("49.5"), CreatedAt=_ts(D5, 302), Stornirano=""),
    ]
    return t


def red(t, ime, kol, vrednost):
    """Prvi red tabele ciji je kol == vrednost (za izmene u testu)."""
    for r in t[ime]:
        if r.get(kol) == vrednost:
            return r
    raise KeyError("%s.%s = %r" % (ime, kol, vrednost))


def kopija(t):
    return copy.deepcopy(t)


def snimi(putanja, t, kolone=None):
    """Snimi kao .xlsx sa pravim Excel tabelama (citac.py ih cita isto kao .xlsm).
    kolone: opciono dict ime -> lista kolona (za test kad kolona fali)."""
    kol = dict(KOLONE)
    kol.update(kolone or {})
    sv = pregled.Sveska()
    for ime, redovi in t.items():
        kolone_t = kol[ime] if ime in kol else sorted({c for r in redovi for c in r})
        sv.dodaj_list(ime.replace("tbl", "")[:28], kolone_t,
                      [[r.get(c, "") for c in kolone_t] for r in redovi], tabela=ime)
    sv.snimi(putanja)
