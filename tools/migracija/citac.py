"""Citac Excel sveske bez Excela -- SAMO CITA.

Cita sve Excel tabele (ListObject) iz .xlsm/.xlsx preko OOXML-a, standardnom
bibliotekom. Excel se ne pokrece, makroi se ne izvrsavaju, fajl se otvara samo
za citanje. Sazeto iz grane chore/import-migration-tools: `legacy_audit.py`
(opseg tabele, zaglavlje, totals red) i MIG-19.1 `XlsmReadOnly` (relacije,
sharedStrings).

Kolone se traze BEZ OBZIRA NA VELIKA SLOVA, kao sto ih trazi i Excel/VBA: stara
sveska ima `DokumentTIP`, kod trazi `DokumentTip`. Python dict razlikuje
velika slova, pa je to vec jednom obaralo citanje prvog ispravnog reda (MIG-04.0).
"""

import hashlib
import posixpath
import re
import zipfile
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from xml.etree import ElementTree as ET

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
N = {"x": NS_MAIN, "r": NS_REL, "p": NS_PKG}

MAX_FAJL = 350 * 1024 * 1024
MAX_RASPAKOVANO = 900 * 1024 * 1024
EXCEL_NULA = datetime(1899, 12, 30)


class GreskaCitanja(Exception):
    pass


class Tabela:
    """Jedna Excel tabela: ime, list, kolone (originalni nazivi) i redovi.

    Red je dict ciji su kljucevi nazivi kolona u `casefold` obliku; `polje()`
    zato radi i za `DokumentTip` i za `DokumentTIP`. `excel_red` je broj reda na
    listu (za poruke operateru), `redni` je redni broj u tabeli, od 1.
    """

    def __init__(self, ime, list_ime, kolone):
        self.ime = ime
        self.list_ime = list_ime
        self.kolone = kolone
        self.redovi = []
        self.formule_bez_vrednosti = 0
        vidjeno = {}
        for k in kolone:
            kf = k.casefold()
            if kf in vidjeno:
                raise GreskaCitanja(
                    "%s: dve kolone se razlikuju samo po velikim slovima: %s / %s"
                    % (ime, vidjeno[kf], k))
            vidjeno[kf] = k

    def ima_kolonu(self, ime):
        return ime.casefold() in {k.casefold() for k in self.kolone}


def polje(red, ime):
    """Vrednost kolone iz reda, bez obzira na velika slova; '' ako je nema."""
    v = red.get(ime.casefold(), "")
    return "" if v is None else v


def tekst(v):
    return "" if v is None else str(v).strip()


def kljuc(v):
    return tekst(v).casefold()


def broj(v):
    """Decimal ili None. Prihvata i decimalni zarez (zatecen tekst u celiji)."""
    s = tekst(v).replace(" ", "").replace("\u00a0", "")
    if not s:
        return None
    if "," in s and "." not in s:
        s = s.replace(",", ".")
    try:
        d = Decimal(s)
    except (InvalidOperation, ValueError):
        return None
    return d if d.is_finite() else None


_FORMATI = ("%Y-%m-%d", "%d.%m.%Y", "%d.%m.%Y.", "%d/%m/%Y",
            "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%d.%m.%Y %H:%M:%S",
            "%d.%m.%Y %H:%M", "%Y-%m-%d %H:%M", "%Y-%m-%dT%H:%M:%S.%f")


def datum(v):
    """date ili None. Excel cuva datum kao broj dana (serijski broj)."""
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = tekst(v)
    if not s:
        return None
    d = broj(s)
    if d is not None:
        if Decimal(20000) <= d < Decimal(100000):
            return (EXCEL_NULA + timedelta(days=int(d))).date()
        return None
    for fmt in _FORMATI:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def je_storno(red):
    """Konvencija starog VBA (ExcludeStornirano): Stornirano = 'Da'."""
    return kljuc(polje(red, "Stornirano")) in ("da", "true", "1", "yes")


def sha256_fajla(putanja):
    h = hashlib.sha256()
    with open(putanja, "rb") as f:
        for deo in iter(lambda: f.read(1024 * 1024), b""):
            h.update(deo)
    return h.hexdigest()


def _tag(ime):
    return "{%s}%s" % (NS_MAIN, ime)


def _putanja(osnova, cilj):
    # Relacija je relativna u odnosu na folder dela koji je nosi.
    if cilj.startswith("/"):
        return posixpath.normpath(cilj.lstrip("/"))
    return posixpath.normpath(posixpath.join(posixpath.dirname(osnova), cilj))


def _relacije(z, deo):
    rels = posixpath.join(posixpath.dirname(deo), "_rels",
                          posixpath.basename(deo) + ".rels")
    if rels not in z.namelist():
        return {}
    out = {}
    for r in ET.fromstring(z.read(rels)):
        if r.get("TargetMode") == "External":
            continue
        out[r.get("Id", "")] = (r.get("Type", ""), _putanja(deo, r.get("Target", "")))
    return out


def _kolona_broj(slova):
    n = 0
    for ch in slova.upper():
        if not "A" <= ch <= "Z":
            raise GreskaCitanja("Neispravna adresa kolone: " + slova)
        n = n * 26 + ord(ch) - 64
    return n


def _opseg(ref):
    m = re.fullmatch(r"\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)", ref or "", re.I)
    if not m:
        raise GreskaCitanja("Nepodrzan opseg tabele: %r" % ref)
    return (_kolona_broj(m.group(1)), int(m.group(2)),
            _kolona_broj(m.group(3)), int(m.group(4)))


def _deljeni_stringovi(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    out = []
    for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall("x:si", N):
        t = si.find("x:t", N)
        if t is not None:
            out.append(t.text or "")
        else:
            # Bogat tekst: samo <r><t>, bez fonetskih <rPh> delova.
            out.append("".join((x.text or "") for r in si.findall("x:r", N)
                               for x in r.findall("x:t", N)))
    return out


# Ugradjeni Excel formati koji su datum/vreme (ECMA-376 18.8.30).
_DATUM_FORMATI = set(range(14, 23)) | set(range(27, 37)) | {45, 46, 47} | set(range(50, 59))


def _datumski_stilovi(z):
    """Skup indeksa cellXfs ciji je format datum/vreme.

    Excel cuva datum kao broj dana; samo format celije kaze da je to datum.
    Bez ovoga bi upis u novu svesku datum preneo kao goli broj (45200).
    """
    if "xl/styles.xml" not in z.namelist():
        return set()
    koren = ET.fromstring(z.read("xl/styles.xml"))
    svoji = {}
    for f in koren.findall("x:numFmts/x:numFmt", N):
        kod = re.sub(r'"[^"]*"|\\.|\[[^\]]*\]', "", f.get("formatCode", ""))
        svoji[int(f.get("numFmtId", "0"))] = bool(re.search(r"[dyh]", kod, re.I))
    out = set()
    for i, xf in enumerate(koren.findall("x:cellXfs/x:xf", N)):
        fid = int(xf.get("numFmtId", "0"))
        if fid in _DATUM_FORMATI or svoji.get(fid):
            out.add(i)
    return out


def _vrednost_celije(c, deljeni, datumski=frozenset()):
    """(vrednost, formula_bez_kesa). Tip se cuva: tekst -> str, broj -> Decimal,
    broj sa datumskim formatom -> datetime (date kad nema vremena), bool -> bool."""
    typ = c.get("t", "")
    if typ == "inlineStr":
        el = c.find("x:is", N)
        if el is None:
            return "", False
        return "".join((x.text or "") for x in el.iter(_tag("t"))), False
    v = c.find("x:v", N)
    if v is None or v.text is None:
        return "", c.find("x:f", N) is not None
    if typ == "s":
        try:
            return deljeni[int(v.text)], False
        except (ValueError, IndexError):
            raise GreskaCitanja("Pokvaren indeks u sharedStrings")
    if typ == "b":
        return v.text == "1", False
    if typ in ("str", "e"):
        return v.text, False
    try:
        d = Decimal(v.text)
    except (InvalidOperation, ValueError):
        return v.text, False
    if c.get("s") and int(c.get("s")) in datumski and Decimal(0) < d < Decimal(2958466):
        dt = EXCEL_NULA + timedelta(days=float(d))
        if dt.hour == dt.minute == dt.second == dt.microsecond == 0:
            return dt.date(), False
        return dt.replace(microsecond=0), False
    return d, False


def procitaj_svesku(putanja):
    """Vrati {ime_tabele.casefold(): Tabela} za sve Excel tabele u svesci."""
    import os
    if not os.path.isfile(putanja):
        raise GreskaCitanja("Sveska ne postoji: %s" % putanja)
    if os.path.getsize(putanja) > MAX_FAJL:
        raise GreskaCitanja("Sveska je veca od bezbednosnog ogranicenja")
    if not zipfile.is_zipfile(putanja):
        raise GreskaCitanja(
            "Nije .xlsm/.xlsx (OOXML). Binarni .xlsb i sifrovane sveske nisu podrzane.")
    with zipfile.ZipFile(putanja, "r") as z:
        if sum(i.file_size for i in z.infolist()) > MAX_RASPAKOVANO:
            raise GreskaCitanja("Raspakovana sveska je prevelika")
        if "xl/workbook.xml" not in z.namelist():
            raise GreskaCitanja("Nedostaje xl/workbook.xml")
        knjiga = ET.fromstring(z.read("xl/workbook.xml"))
        rel_knjige = _relacije(z, "xl/workbook.xml")
        listovi = []
        for sh in knjiga.findall("x:sheets/x:sheet", N):
            rid = sh.get("{%s}id" % NS_REL, "")
            if rid in rel_knjige:
                listovi.append((sh.get("name", ""), rel_knjige[rid][1]))
        deljeni = _deljeni_stringovi(z)
        datumski = _datumski_stilovi(z)
        tabele = {}
        for list_ime, list_deo in listovi:
            if list_deo not in z.namelist():
                continue
            rel_lista = _relacije(z, list_deo)
            meta = []
            for typ, deo in rel_lista.values():
                if not typ.endswith("/table"):
                    continue
                t = ET.fromstring(z.read(deo))
                ime = t.get("name") or t.get("displayName") or ""
                kolone = [c.get("name", "") for c in t.findall("x:tableColumns/x:tableColumn", N)]
                c1, r1, c2, r2 = _opseg(t.get("ref", ""))
                zaglavlje = int(t.get("headerRowCount", "1"))
                ukupno = int(t.get("totalsRowCount", "0"))
                if c2 - c1 + 1 != len(kolone):
                    raise GreskaCitanja("%s: sirina opsega i broj kolona se ne slazu" % ime)
                if not ime or ime.casefold() in tabele:
                    raise GreskaCitanja("Prazno ili duplo ime tabele: %r" % ime)
                tab = Tabela(ime, list_ime, kolone)
                tabele[ime.casefold()] = tab
                meta.append((tab, c1, c2, r1 + zaglavlje, r2 - ukupno))
            if not meta:
                continue
            koren = ET.fromstring(z.read(list_deo))
            prethodni_red = 0
            for row in koren.findall("x:sheetData/x:row", N):
                rn = int(row.get("r", prethodni_red + 1))
                prethodni_red = rn
                pogodjene = [m for m in meta if m[3] <= rn <= m[4]]
                if not pogodjene:
                    continue
                celije = {}
                bez_kesa = set()
                prethodna_kol = 0
                for c in row.findall("x:c", N):
                    ref = c.get("r")
                    if ref:
                        ci = _kolona_broj(re.match(r"\$?([A-Z]+)", ref, re.I).group(1))
                    else:
                        ci = prethodna_kol + 1
                    prethodna_kol = ci
                    vrednost, nema_kes = _vrednost_celije(c, deljeni, datumski)
                    celije[ci] = vrednost
                    if nema_kes:
                        bez_kesa.add(ci)
                for tab, c1, c2, _p, _k in pogodjene:
                    vrednosti = [celije.get(i, "") for i in range(c1, c2 + 1)]
                    if any(tekst(v) for v in vrednosti):
                        red = {k.casefold(): v for k, v in zip(tab.kolone, vrednosti)}
                        red["_excel_red"] = rn
                        red["_redni"] = len(tab.redovi) + 1
                        tab.redovi.append(red)
                    tab.formule_bez_vrednosti += sum(c1 <= ci <= c2 for ci in bez_kesa)
        return tabele


def tabela(tabele, ime):
    """Tabela po imenu (bez obzira na velika slova) ili None."""
    return tabele.get(ime.casefold())
