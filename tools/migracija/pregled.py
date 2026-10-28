"""Pregled migracije kao Excel sveska (.xlsx) -- standardnom bibliotekom.

Svaki list je prava Excel tabela (filter na zaglavlju, zamrznut prvi red), pa
se pregled po kooperantu/stanici dobija filterom na koloni Naziv ili ID.

Isti pisac pravi i sinteticku svesku za testove (test_migracija.py): tabele
koje on napise, citac.py cita nazad kao ListObject-e.
"""

import re
import zipfile
from collections import Counter, defaultdict
from datetime import date, datetime
from decimal import Decimal
from xml.sax.saxutils import escape

from ambalaza import (DOK_AMB, ISHOD_IZUZETO, ISHOD_KRETANJE, ISHOD_STORNO, ISHOD_ZA_ODLUKU,
                      ISPRAVNO, NABAVKA, OPIS, PARTNERI, SOPSTVENI, ULAZ_TUDJE)
from citac import EXCEL_NULA, kljuc

_NEDOZVOLJENO = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

# ---------------------------------------------------------------------------
# Pisac .xlsx
# ---------------------------------------------------------------------------


def _kolona(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


class _Stringovi:
    """Deljeni stringovi (xl/sharedStrings.xml) -- oblik koji i Excel pise."""

    def __init__(self):
        self.indeks = {}
        self.lista = []

    def __call__(self, tekst):
        tekst = _NEDOZVOLJENO.sub("", str(tekst))
        if tekst not in self.indeks:
            self.indeks[tekst] = len(self.lista)
            self.lista.append(tekst)
        return self.indeks[tekst]

    def xml(self):
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'count="%d" uniqueCount="%d">%s</sst>'
                % (len(self.lista), len(self.lista), "".join(
                    '<si><t xml:space="preserve">%s</t></si>' % escape(t) for t in self.lista)))


def _celija(ref, v, sst):
    if v is None or v == "":
        return ""
    if isinstance(v, bool):
        v = "DA" if v else "NE"
    if isinstance(v, datetime):
        v = v.date()
    if isinstance(v, date):
        return '<c r="%s" s="1"><v>%d</v></c>' % (ref, (v - EXCEL_NULA.date()).days)
    if isinstance(v, (int, float, Decimal)):
        if isinstance(v, Decimal):
            txt = format(v.normalize(), "f")
        else:
            txt = repr(v)
        return '<c r="%s"><v>%s</v></c>' % (ref, txt)
    return '<c r="%s" t="s"><v>%d</v></c>' % (ref, sst(v))


class Sveska:
    def __init__(self):
        self.listovi = []

    def dodaj_list(self, ime, kolone, redovi, sirine=None, tabela=None):
        """kolone: nazivi; redovi: liste vrednosti; tabela: ime Excel tabele."""
        ime = re.sub(r"[\[\]:*?/\\]", "_", ime)[:31]
        self.listovi.append(dict(ime=ime, kolone=list(kolone), redovi=[list(r) for r in redovi],
                                 sirine=sirine, tabela=tabela))

    def snimi(self, putanja):
        with zipfile.ZipFile(putanja, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", self._tipovi())
            z.writestr("_rels/.rels",
                       '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                       '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                       '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                       'relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
            z.writestr("xl/workbook.xml", self._knjiga())
            z.writestr("xl/_rels/workbook.xml.rels", self._rel_knjige())
            z.writestr("xl/styles.xml", _STILOVI)
            sst = _Stringovi()
            br_tabele = 0
            for i, l in enumerate(self.listovi, 1):
                tabela_deo = None
                if l["tabela"]:
                    br_tabele += 1
                    tabela_deo = br_tabele
                    z.writestr("xl/tables/table%d.xml" % br_tabele, self._tabela(l, br_tabele))
                    z.writestr("xl/worksheets/_rels/sheet%d.xml.rels" % i,
                               '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                               '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/'
                               '2006/relationships/table" Target="../tables/table%d.xml"/></Relationships>'
                               % br_tabele)
                z.writestr("xl/worksheets/sheet%d.xml" % i, self._list(l, tabela_deo, sst))
            z.writestr("xl/sharedStrings.xml", sst.xml())

    def _tipovi(self):
        delovi = "".join(
            '<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/'
            'vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % i
            for i in range(1, len(self.listovi) + 1))
        n_tab = sum(1 for l in self.listovi if l["tabela"])
        delovi += "".join(
            '<Override PartName="/xl/tables/table%d.xml" ContentType="application/'
            'vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>' % i
            for i in range(1, n_tab + 1))
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Override PartName="/xl/workbook.xml" ContentType="application/'
                'vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                '<Override PartName="/xl/styles.xml" ContentType="application/'
                'vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                '<Override PartName="/xl/sharedStrings.xml" ContentType="application/'
                'vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' + delovi + '</Types>')

    def _knjiga(self):
        listovi = "".join('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (escape(l["ime"]), i, i)
                          for i, l in enumerate(self.listovi, 1))
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                '<sheets>' + listovi + '</sheets></workbook>')

    def _rel_knjige(self):
        n = len(self.listovi)
        rels = "".join('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/'
                       '2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>' % (i, i)
                       for i in range(1, n + 1))
        rels += ('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/'
                 '2006/relationships/styles" Target="styles.xml"/>' % (n + 1))
        rels += ('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/'
                 '2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' % (n + 2))
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                + rels + '</Relationships>')

    def _list(self, l, tabela_deo, sst):
        kolone, redovi = l["kolone"], l["redovi"]
        n_kol = max(1, len(kolone))
        poslednji = max(2, len(redovi) + 1) if l["tabela"] else len(redovi) + 1
        delovi = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
                  '<dimension ref="A1:%s%d"/>' % (_kolona(n_kol), poslednji),
                  '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" '
                  'activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>']
        sirine = l["sirine"] or [max(10, min(45, len(str(k)) + 3)) for k in kolone]
        delovi.append("<cols>" + "".join(
            '<col min="%d" max="%d" width="%s" customWidth="1"/>' % (i, i, w)
            for i, w in enumerate(sirine, 1)) + "</cols>")
        delovi.append("<sheetData>")
        stil = "" if l["tabela"] else ' s="2"'
        zag = "".join('<c r="%s1" t="s"%s><v>%d</v></c>' % (_kolona(i), stil, sst(k))
                      for i, k in enumerate(kolone, 1))
        delovi.append('<row r="1">%s</row>' % zag)
        for rn, red in enumerate(redovi, 2):
            delovi.append('<row r="%d">%s</row>' % (rn, "".join(
                _celija("%s%d" % (_kolona(i), rn), v, sst) for i, v in enumerate(red, 1))))
        delovi.append("</sheetData>")
        if tabela_deo:
            delovi.append('<tableParts count="1"><tablePart r:id="rId1"/></tableParts>')
        delovi.append("</worksheet>")
        return "".join(delovi)

    def _tabela(self, l, br):
        kolone = l["kolone"]
        poslednji = max(2, len(l["redovi"]) + 1)
        ref = "A1:%s%d" % (_kolona(len(kolone)), poslednji)
        kol = "".join('<tableColumn id="%d" name="%s"/>' % (i, escape(str(k)))
                      for i, k in enumerate(kolone, 1))
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="%d" '
                'name="%s" displayName="%s" ref="%s" totalsRowShown="0"><autoFilter ref="%s"/>'
                '<tableColumns count="%d">%s</tableColumns><tableStyleInfo name="TableStyleMedium2" '
                'showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>'
                % (br, escape(l["tabela"]), escape(l["tabela"]), ref, ref, len(kolone), kol))


_STILOVI = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<numFmts count="1"><numFmt numFmtId="164" formatCode="dd.mm.yyyy"/></numFmts>'
            '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>'
            '<font><b/><sz val="11"/><name val="Calibri"/></font></fonts>'
            '<fills count="2"><fill><patternFill patternType="none"/></fill>'
            '<fill><patternFill patternType="gray125"/></fill></fills>'
            '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
            '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
            '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>'
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
            '</styleSheet>')


# ---------------------------------------------------------------------------
# Izvestaj migracije
# ---------------------------------------------------------------------------

OBJASNJENJE = {
    "NEISPRAVNA_KOLICINA": "Kolicina nije broj veci od nule.",
    "NEISPRAVAN_SMER": "Smer nije Ulaz ni Izlaz.",
    "NEISPRAVAN_DATUM": "Datum se ne moze procitati.",
    "NEMA_TIPA_AMBALAZE": "Prazan tip ambalaze.",
    "TIP_AMBALAZE_NIJE_U_MATICNOJ": "Tip ambalaze ne postoji u tblTipAmbalaze (AMB-INV-12).",
    "NEPOZNAT_NALOG": "EntitetTip nije Kooperant, Stanica, Kupac ni Vozac.",
    "NEMA_NALOGA": "Prazan EntitetID.",
    "NALOG_NIJE_U_MATICNOJ": "EntitetID ne postoji u maticnoj tabeli.",
    "VOZAC_NIJE_U_MATICNOJ": "VozacID ne postoji u tblVozaci.",
    "NOGA_BEZ_PARA": "Dvostrani upis (Otkup / OM-*-Koop) nema drugu nogu iste kolicine, datuma i tipa.",
    "NEOCEKIVANA_ULOGA_NOGE": "Entitet ili smer ne odgovara tipu dokumenta.",
    "DVOSMISLENO_SPARIVANJE": "Vise nogu i vise naloga na obe strane -- par se ne moze dokazati.",
    "NEMA_DOKUMENTA_RODITELJA": "DokumentID ne postoji u tabeli dokumenta.",
    "AKTIVNA_NOGA_STORNIRANOG_DOKUMENTA": "Dokument je storniran, a red ambalaze je ostao aktivan (kao AMB-00700).",
    "IZDAVANJE_UZ_OTKUP_NALOG_SE_NE_SLAZE": "OM-Izlaz-Koop nosi OtkupID, ali kooperant/stanica nisu sa tog otkupa.",
    "REVERS_BEZ_BROJA": "Revers bez broja dokumenta -- nema zaglavlja kome bi pripao.",
    "NEMA_VOZACA_ZA_OGLEDALO": "Jednostrani upis bez VozacID -- druga strana kretanja je nepoznata.",
    "OMULAZ_IZVOR_NEPOZNAT": "OMUlaz: jednostrani prijem na stanici, izvor nije zapisan.",
    "NEPOZNAT_TIP_DOKUMENTA": "Tip dokumenta za koji ne postoji pravilo.",
    "IZUZETAK_POTPIS_SE_NE_SLAZE": "Red se promenio posle odluke u izuzeci.json -- odluka vise ne vazi.",
    "IZUZETAK_ZA_ODLUKU": "Rucno oznaceno u izuzeci.json kao otvoreno.",
    "IZUZETAK_NAD_NEISPRAVNIM_REDOM": "Izuzetak trazi kretanje, a red nema ispravnu kolicinu/datum/tip.",
    "PAD_OBAVEZE_BEZ_IZDAVANJA": "Dug prema partneru pada, a tog dana nema izdatih praznih da se oznace kao vracanje.",
    "RAST_OBAVEZE_BEZ_IZLAZA": "Dug prema partneru raste, a tog dana nema partnerovog izlaza da nosi pokrice.",
}


def _kom(d):
    return format(d.normalize(), "f")


def _naziv(rez, tip, idn):
    return rez.nazivi.get((tip, kljuc(idn)), "")


def _broj_dok(rez, dok_tip, dok_id):
    if dok_tip == DOK_AMB:
        z = rez.zaglavlja.get(dok_id)
        return z["BrojDokumenta"] if z else ""
    return rez.brojevi_dok.get((dok_tip, kljuc(dok_id)), "")


def _predlog_preseka(s):
    """Dogadjaj na dan preseka koji nalog sa minusom na kraju dovodi na nulu.

    Istorija se ne dira; main posle preseka ne pusta realan nalog ispod nule
    (AMB-INV-07). Sopstveni nalog (stanica, vozac): NABAVKA. Partner (kooperant,
    kupac): ULAZ_TUDJE -- firma drzi njegove gajbe, to je obaveza firme.
    """
    if s["novi"] >= 0:
        return "", Decimal(0)
    manjak = -s["novi"]
    if s["nalog_tip"] in SOPSTVENI:
        return "NABAVKA", manjak
    if s["nalog_tip"] in PARTNERI:
        return "ULAZ_TUDJE_AMBALAZE", manjak
    return "?", manjak


def napravi_izvestaj(putanja, izvor, rez, salda, stavke, invarijante, pokrivenost,
                     popis, dok_sazetak, dok_grupe, oblik_dok):
    sv = Sveska()
    po_rednom = {n.redni: n for n in rez.noge}
    ishodi = Counter(n.ishod for n in rez.noge)
    vrste = Counter(k.vrsta for k in rez.kretanja)
    pokrica = [k for k in rez.kretanja if k.generisano and k.vrsta == ULAZ_TUDJE]
    nabavke = [k for k in rez.kretanja if k.generisano and k.vrsta == NABAVKA]
    ispod_nule = [s for s in salda if s["najmanji"] < 0 and s["nalog_tip"] in PARTNERI + SOPSTVENI]
    minus = defaultdict(lambda: [0, Decimal(0)])
    bio = Counter()
    for s in salda:
        if s["novi"] < 0:
            minus[s["nalog_tip"]][0] += 1
            minus[s["nalog_tip"]][1] += -s["novi"]
        elif s["najmanji"] < 0:
            bio[s["nalog_tip"]] += 1
    razlike = [s for s in salda if s["razlika"] != 0]
    inv = Counter(x["inv"] for x in invarijante)
    pokr_ok = not any(pokrivenost.values())

    saz = [
        ("Izvor", izvor["ime"]),
        ("Izvor SHA256", izvor["sha"]),
        ("Izvor nepromenjen posle citanja", "DA" if izvor["nepromenjen"] else "NE -- PROVERI"),
        ("Vreme pregleda", izvor["vreme"]),
        ("Oblik knjige ambalaze", "stari (noge Ulaz/Izlaz)" if rez.oblik == "STARI" else "AMB-10 (vec preveden)"),
        ("Oblik dokumenata", "red po klasi (stari)" if oblik_dok == "STARI" else "zaglavlje + stavke"),
        ("", ""),
        ("Stari redovi ambalaze -- ukupno", len(rez.noge)),
        ("  u kretanjima", ishodi.get(ISHOD_KRETANJE, 0)),
        ("  stornirani (arhiva)", ishodi.get(ISHOD_STORNO, 0)),
        ("  izuzeti (izuzeci.json)", ishodi.get(ISHOD_IZUZETO, 0)),
        ("  ZA ODLUKU", ishodi.get(ISHOD_ZA_ODLUKU, 0)),
        ("Svaki red ima tacno jedan ishod", "DA" if pokr_ok else "NE -- GRESKA ALATA"),
        ("", ""),
        ("AMB-10 kretanja -- ukupno", len(rez.kretanja)),
    ]
    saz += [("  " + v, c) for v, c in sorted(vrste.items())]
    saz += [
        ("Knjizenje", "ISPRAVNO (istorija + dopuna za rad)" if rez.knjizenje == ISPRAVNO
         else "VERNO (tacno kao stara sveska)"),
        ("  dopisano ULAZ_TUDJE (kooperant, kupac)", "%d red., %s kom." % (
            len(pokrica), _kom(sum((k.kolicina for k in pokrica), Decimal(0))))),
        ("  dopisano NABAVKA (stanica, vozac)", "%d dok., %s kom." % (
            len(nabavke), _kom(sum((k.kolicina for k in nabavke), Decimal(0))))),
        ("Zaglavlja reversa (nacrt)", len(rez.zaglavlja)),
        ("", ""),
        ("Naloga (nalog x tip ambalaze)", len(salda)),
    ]
    for t in ("Kooperant", "Stanica", "Vozac", "Kupac"):
        if t in minus:
            saz.append(("  u minusu NA KRAJU: %s" % t, "%d, presek %s %s kom." % (
                minus[t][0], "NABAVKA" if t in SOPSTVENI else "ULAZ_TUDJE",
                format(minus[t][1].normalize(), "f"))))
    for t in ("Kooperant", "Stanica", "Vozac", "Kupac"):
        if bio[t]:
            saz.append(("  bio u minusu tokom sezone, na kraju >= 0: %s" % t, bio[t]))
    saz += [
        ("Nijedan nalog ni jednog dana ispod nule (AMB-INV-07)", "DA" if not ispod_nule else
         "NE -- %d naloga" % len(ispod_nule)),
        ("Salda: stari - van knjige + izuzeci + pokrice + nabavka = novi", "DA" if not razlike else
         "NE za %d naloga -- GRESKA ALATA" % len(razlike)),
        ("Nalazi nad invarijantama", ", ".join("%s: %d" % x for x in sorted(inv.items())) or "nema"),
        ("Stavki za odluku", len(rez.za_odluku)),
    ]
    po_razlogu = Counter()
    for z in rez.za_odluku:
        po_razlogu[z["razlog"]] += max(1, len(z["noge"]))
    saz += [("  " + r, "%d red." % c) for r, c in sorted(po_razlogu.items())]
    saz += [
        ("Upozorenja", len(rez.upozorenja)),
        ("Primenjeni izuzeci", len(rez.izuzeci_primenjeni)),
        ("", ""),
    ]
    for d in dok_sazetak:
        saz.append(("Dokumenti %s" % d["tabela"], "%d redova -> %d logickih, %d sa vise redova"
                    % (d["redova"], d["logickih"], d["sa_vise_redova"])))
    fali_tab = [p["tabela"] for p in popis if p["u_svesci"] == "NE"]
    visak_tab = [p["tabela"] for p in popis if "NEMA JE U KANONU" in p["napomena"]]
    saz += [
        ("Tabele kanona kojih nema u svesci", len(fali_tab)),
        ("Tabele sveske kojih nema u kanonu", ", ".join(visak_tab) or "nema"),
    ]
    sv.dodaj_list("Sazetak", ["Stavka", "Vrednost"], saz, sirine=[55, 70])

    # Salda
    red_s = []
    for s in salda:
        predlog, manjak = _predlog_preseka(s)
        status = "OK"
        if s["razlika"] != 0:
            status = "RAZLIKA"
        elif s["novi"] < 0:
            status = "MINUS NA KRAJU"
        elif s["najmanji"] < 0:
            status = "BIO U MINUSU"
        red_s.append([s["nalog_tip"], s["nalog_id"], s["naziv"], s["tip"], s["stari"], s["novi"],
                      s["pokrice"], s["nabavka"], s["obaveza"] if s["nalog_tip"] in PARTNERI else "",
                      s["van_knjige"], s["izuzeci"], s["razlika"], s["najmanji"], s["najmanji_dan"],
                      predlog, manjak or "", status])
    sv.dodaj_list("Salda", ["Tip naloga", "ID", "Naziv", "Tip ambalaze", "Stari saldo", "Novi saldo",
                            "Pokrice tudje", "Nabavka", "Obaveza firme", "Van knjige", "Izuzeci", "Razlika",
                            "Najmanji dnevni saldo", "Dan najmanjeg", "Presek: dogadjaj",
                            "Presek: kolicina", "Status"],
                  red_s, sirine=[11, 12, 28, 12, 11, 11, 12, 10, 12, 11, 10, 9, 13, 13, 22, 12, 15],
                  tabela="Salda")

    # Stavke po nalogu
    red_st = []
    for kl in sorted(stavke, key=lambda x: (x[0], x[1], x[2])):
        lst = stavke[kl]
        prvi = lst[0]
        if prvi["kretanje"] is not None:
            k0 = prvi["kretanje"]
            nalog = next(((t, i) for t, i in ((k0.od_tip, k0.od_id), (k0.na_tip, k0.na_id))
                          if (t, kljuc(i)) == (kl[0], kl[1])), (kl[0], kl[1]))
        else:
            nalog = (kl[0], prvi["noga"].vozac if kl[0] == "Vozac" else prvi["noga"].ent_id)
        naziv = _naziv(rez, nalog[0], nalog[1])
        for st in lst:
            k, n = st["kretanje"], st["noga"]
            if k is not None:
                dok, broj = k.dok_tip, _broj_dok(rez, k.dok_tip, k.dok_id)
                stari_tip, stari_broj = k.stari_tip, k.stari_broj
                vrsta, druga = k.vrsta, st["druga"]
                napomena = k.pravilo + (" -- " + k.napomena if k.napomena else "")
                novi_id, stari_id = k.amb_id, ";".join(k.stari_amb)
                opis = OPIS.get(kljuc(k.stari_tip), "")
                tip = k.tip
            else:
                dok, broj = "", ""
                stari_tip, stari_broj = n.dok_tip, n.dok_id
                vrsta, druga = "VAN KNJIGE (%s)" % n.ishod, None
                napomena = n.razlog
                novi_id, stari_id = "", n.amb_id
                opis = OPIS.get(n.dt, "")
                tip = n.tip
            red_st.append([
                nalog[0], nalog[1], naziv, tip, st["datum"], dok, broj, stari_tip, stari_broj, opis,
                vrsta, (druga[0] + " " + druga[1]).strip() if druga else "",
                _naziv(rez, druga[0], druga[1]) if druga else "",
                st["novo"] if st["novo"] > 0 else "", -st["novo"] if st["novo"] < 0 else "",
                st["saldo_novi"],
                st["staro"] if st["staro"] > 0 else "", -st["staro"] if st["staro"] < 0 else "",
                st["saldo_stari"], napomena, novi_id, stari_id])
    sv.dodaj_list("Stavke", ["Tip naloga", "ID", "Naziv", "Tip ambalaze", "Datum", "Dokument",
                             "Broj dokumenta", "Stari tip dokumenta", "Stari DokumentID", "Opis",
                             "Vrsta kretanja", "Druga strana", "Naziv druge strane", "Ulaz", "Izlaz",
                             "Saldo", "Stari ulaz", "Stari izlaz", "Stari saldo", "Napomena",
                             "Novi AmbID", "Stari AmbID"],
                  red_st, sirine=[11, 12, 26, 11, 11, 16, 14, 16, 16, 30, 24, 20, 24, 8, 8, 9, 9, 9,
                                  10, 40, 38, 14], tabela="Stavke")

    # Kretanja -- nacrt knjige u redosledu kolona kanona tblAmbalaza
    red_k = []
    for k in sorted(rez.kretanja, key=lambda x: (x.datum is None, x.datum or 0, x.redosled, x.generisano)):
        red_k.append([k.amb_id, k.datum, k.tip, k.kolicina, k.dok_id, k.dok_tip, k.created_at,
                      k.created_by, "", "", k.od_tip, k.od_id, k.na_tip, k.na_id, k.vrsta, "",
                      _broj_dok(rez, k.dok_tip, k.dok_id), k.pravilo, ";".join(k.stari_amb), k.napomena])
    sv.dodaj_list("Kretanja", ["AmbID", "Datum", "TipAmbalaze", "Kolicina", "DokumentID", "DokumentTIP",
                               "CreatedAt", "CreatedBy", "ModifiedAt", "ModifiedBy", "OdNalogTip",
                               "OdNalogID", "NaNalogTip", "NaNalogID", "VrstaKretanja", "StornoOd",
                               "Broj dokumenta", "Pravilo", "Stari AmbID", "Napomena"],
                  red_k, sirine=[38, 11, 11, 9, 38, 16, 18, 12, 10, 10, 11, 12, 11, 12, 24, 8, 14, 22,
                                 16, 40], tabela="Kretanja")

    sv.dodaj_list("Reversi", ["AmbDokID", "Vrsta", "BrojDokumenta", "Datum", "BrojOwnerTip",
                              "BrojOwnerID", "Napomena", "Partneri", "Kretanja"],
                  [[z["AmbDokID"], z["Vrsta"], z["BrojDokumenta"], z["Datum"], z["BrojOwnerTip"],
                    z["BrojOwnerID"], z["Napomena"], z["partneri"], z["kretanja"]]
                   for z in sorted(rez.zaglavlja.values(), key=lambda x: (x["Datum"] or date.min,
                                                                          x["BrojDokumenta"]))],
                  sirine=[38, 16, 14, 11, 12, 12, 50, 40, 9], tabela="Reversi")

    # Za odluku: jedan red po nozi (ili po stavci bez nogu)
    red_z = []
    for z in rez.za_odluku:
        objasnjenje = OBJASNJENJE.get(z["razlog"], "")
        if not z["noge"]:
            red_z.append([z["razlog"], objasnjenje, z["predlog"], z["napomena"]] + [""] * 14)
            continue
        for r in z["noge"]:
            n = po_rednom[r]
            rod_tip = {"otkup": "Otkup", "om-izlaz-koop": "Otkup", "otpremnica": "Otpremnica",
                       "prijemnica": "Prijemnica"}.get(n.dt, "")
            red_z.append([z["razlog"], objasnjenje, z["predlog"], z["napomena"], n.redni, n.excel_red,
                          n.amb_id, n.datum or n.datum_txt, n.dok_tip, n.dok_id,
                          _broj_dok(rez, rod_tip, n.dok_id) if rod_tip else "",
                          n.ent_tip, n.ent_id, _naziv(rez, n.ent_tip, n.ent_id), n.smer, n.vozac,
                          n.tip, n.kolicina])
    sv.dodaj_list("ZaOdluku", ["Razlog", "Objasnjenje", "Predlog", "Napomena", "Redni", "Excel red",
                               "AmbID", "Datum", "Stari tip dokumenta", "DokumentID", "Broj dokumenta",
                               "EntitetTip", "EntitetID", "Naziv", "Smer", "VozacID", "Tip ambalaze",
                               "Kolicina"],
                  red_z, sirine=[30, 50, 10, 40, 7, 8, 12, 11, 16, 14, 13, 11, 12, 24, 7, 10, 11, 9],
                  tabela="ZaOdluku")

    red_u = [["IZUZETAK_PRIMENJEN", i["redni"], i["amb_id"], "%s: %s" % (i["radnja"], i["razlog"])]
             for i in rez.izuzeci_primenjeni]
    red_u += [[u["vrsta"], u.get("redni", ""), u.get("amb_id", ""), u["opis"]] for u in rez.upozorenja]
    sv.dodaj_list("Upozorenja", ["Vrsta", "Redni", "AmbID", "Opis"], red_u,
                  sirine=[34, 8, 14, 90], tabela="Upozorenja")

    sv.dodaj_list("Invarijante", ["Invarijanta", "Opis", "AmbID"],
                  [[x["inv"], x["opis"], x["amb_id"]] for x in invarijante],
                  sirine=[13, 90, 60], tabela="Invarijante")

    sv.dodaj_list("Dokumenti", ["Tabela", "Redova", "Aktivnih", "Logickih dokumenata",
                                "Sa vise redova", "Po GeneracijaID", "Dupla klasa", "Delimicno storno",
                                "Bez poslovnog broja", "Ima GeneracijaID"],
                  [[d["tabela"], d["redova"], d["aktivnih"], d["logickih"], d["sa_vise_redova"],
                    d["po_generaciji"], d["dupla_klasa"], d["mesano_storno"], d["bez_broja"],
                    d["ima_generaciju"]] for d in dok_sazetak],
                  sirine=[16, 9, 9, 18, 14, 15, 11, 15, 17, 15], tabela="Dokumenti")

    sv.dodaj_list("DokumentiViseRedova", ["Tabela", "Osnov grupe", "Broj", "Datum", "Redova", "Aktivnih",
                                          "ID-jevi", "Klase", "Napomena"],
                  [[g["tabela"], g["osnov"], g["broj"], g["datum"], g["redova"], g["aktivnih"], g["id"],
                    g["klase"], g["napomena"]] for g in dok_grupe],
                  sirine=[15, 18, 14, 11, 8, 9, 40, 14, 40], tabela="DokumentiViseRedova")

    sv.dodaj_list("Tabele", ["Tabela", "U svesci", "Redova", "Fali kolona", "Visak kolona", "Redosled",
                             "Napomena"],
                  [[p["tabela"], p["u_svesci"], p["redova"], p["fali"], p["visak"], p["redosled"],
                    p["napomena"]] for p in popis],
                  sirine=[24, 9, 9, 50, 50, 14, 50], tabela="Tabele")
    sv.snimi(putanja)
    return saz
