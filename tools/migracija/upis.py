"""Upis nove sveske preko Excela -- JEDINI deo alata koji pise.

Trazi Windows + Excel + pywin32 (isto kao tools/make_dev_workbook.py i
tools/run_vba.py). U Linux sesiji i CI-ju se izvrsavaju samo ciste funkcije
(pretvaranje vrednosti, mapiranje kolona, poredjenje posle upisa) -- one imaju
testove; sam COM deo je NEVERIFIKOVAN dok ga operater ne pokrene.

Bezbednost:
  * izvorna sveska se nikad ne otvara u Excelu (cita je citac.py, OOXML);
  * cilj je SVEZA sveska iz gita (make_dev_workbook); tabela u koju se pise mora
    biti PRAZNA, inace upis odbija -- alat nikad ne mesa podatke;
  * sveska se otvara sa ugasenim makroima i dogadjajima (Workbook_Open se ne
    pokrece), a cuva se tek kad su sve tabele upisane;
  * posle cuvanja sveska se procita ponovo (OOXML) i svaka celija se poredi sa
    onim sto je trebalo da bude upisano.

Ovaj masovni upis je zapisan izuzetak od AMB-INV-08 i od A11 (vlasnik upisa):
jednokratno punjenje PRAZNE sveske pre rada -- docs/DOMEN/MIGRACIJA_LEGACY.md.
"""

import os
import re
from datetime import date, datetime
from decimal import Decimal

from citac import procitaj_svesku, tabela
from cilj import TEKST_KOLONE

XL_AUTOMATION_FORCE_DISABLE = 3
DELOVI = 2000

# Tabele koje nova sveska vec nosi iz make_dev_workbook; ako su pune, ne diraju se.
PRESKOCI_AKO_PUNA = {"tblKorisnici"}


def com_vrednost(v):
    """Python vrednost -> vrednost koju Excel prima preko COM-a, bez promene smisla.

    Tekst koji bi Excel sam pretvorio u broj ili datum ("101", "2026-06-20 08:00",
    "012345") dobija apostrof -- Excel ga cuva kao tekst, a apostrof nije deo
    vrednosti. Isto za tekst koji pocinje sa = + - @ (inace postaje formula).
    """
    if v is None or (isinstance(v, str) and v == ""):
        return None
    if isinstance(v, bool):
        return v
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    if isinstance(v, Decimal):
        if v == v.to_integral_value() and abs(v) < 2 ** 53:
            return int(v)
        return float(v)
    if isinstance(v, (int, float)):
        return v
    s = str(v)
    if re.search(r"\d", s) or s[:1] in ("=", "+", "-", "@"):
        return "'" + s
    return s


def niz_za_tabelu(redovi, kolone_u_svesci):
    """Redovi (dict po kanonskom imenu) -> 2D niz u redosledu kolona U SVESCI.

    Kolona koja postoji u svesci a nema je u redu (npr. runtime kolona na kraju)
    ostaje prazna; vrednost reda bez kolone u svesci je greska -- ne gubi se tiho.
    Tekstualna kolona (TEKST_KOLONE) dobija tekst BEZ apostrofa: _upisi_tabelu joj
    pre upisa postavlja format "@" (isti koji trazi modSchema.SchemaReadyOrFail).
    """
    kf = {}
    for r in redovi[:1]:
        kf = {k.casefold(): k for k in r}
    u_svesci = {k.casefold() for k in kolone_u_svesci}
    nema = [k for c, k in kf.items() if c not in u_svesci and any(r.get(k) not in ("", None) for r in redovi)]
    if nema:
        raise ValueError("kolone sa podacima kojih nema u tabeli nove sveske: " + ", ".join(nema))
    tekst_kf = {k.casefold() for k in TEKST_KOLONE}
    out = []
    for r in redovi:
        red = []
        for k in kolone_u_svesci:
            v = r.get(kf.get(k.casefold(), k), "")
            if k.casefold() in tekst_kf and v not in ("", None):
                red.append(str(v))
            else:
                red.append(com_vrednost(v))
        out.append(tuple(red))
    return out


def norm(v):
    """Uporediv oblik vrednosti posle povratnog citanja (tekst/broj/datum)."""
    if v is None or v == "":
        return ""
    if isinstance(v, bool):
        return "1" if v else "0"
    if isinstance(v, datetime):
        if v.hour == v.minute == v.second == 0:
            return v.date().isoformat()
        return v.replace(microsecond=0).isoformat(sep=" ")
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, (int, float, Decimal)):
        d = Decimal(str(v))
        return "0" if d == 0 else format(d.normalize(), "f")
    return str(v).strip()


def uporedi(c, putanja, preskocene=()):
    """Procitaj upisanu svesku i uporedi svaku celiju. Vrati listu razlika.
    `preskocene`: tabele koje upis namerno nije dirao (vec pune u novoj svesci)."""
    t = procitaj_svesku(putanja)
    razlike = []
    for ime, redovi in sorted(c.tabele.items()):
        if not redovi or ime in preskocene:
            continue
        tb = tabela(t, ime)
        if tb is None:
            razlike.append("%s: tabele nema u upisanoj svesci" % ime)
            continue
        if len(tb.redovi) != len(redovi):
            razlike.append("%s: upisano %d redova, ocekivano %d" % (ime, len(tb.redovi), len(redovi)))
            continue
        for i, (ocek, proc) in enumerate(zip(redovi, tb.redovi), 1):
            for k, v in ocek.items():
                a, b = norm(v), norm(proc.get(k.casefold(), ""))
                if a != b:
                    razlike.append("%s red %d, %s: ocekivano %r, upisano %r" % (ime, i, k, a, b))
                    break
        if len(razlike) > 200:
            razlike.append("... (prekinuto posle 200 razlika)")
            break
    return razlike


# ---------------------------------------------------------------------------
# COM -- trazi Windows + Excel + pywin32
# ---------------------------------------------------------------------------

def _excel_putanja(p):
    # Excel trazi backslash; "C:/x/y.xlsm" mu je nevalidna putanja.
    return os.path.abspath(p).replace("/", "\\")


def upisi(putanja, c, log=print):
    """Upisi sve tabele i podesavanja; vrati skup tabela koje su namerno preskocene."""
    import win32com.client as w32

    xl = w32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AutomationSecurity = XL_AUTOMATION_FORCE_DISABLE
    wb = None
    try:
        wb = xl.Workbooks.Open(_excel_putanja(putanja), UpdateLinks=0)
        lo_po_imenu = {}
        for ws in wb.Worksheets:
            for lo in ws.ListObjects:
                lo_po_imenu[str(lo.Name).casefold()] = lo

        plan = []
        preskocene = set()
        for ime, redovi in sorted(c.tabele.items()):
            if not redovi:
                continue
            lo = lo_po_imenu.get(ime.casefold())
            if lo is None:
                raise RuntimeError("Nova sveska nema tabelu %s -- da li je napravljena iz gita?" % ime)
            if _ima_podataka(lo):
                if ime in PRESKOCI_AKO_PUNA:
                    log("  %s: vec ima redove u novoj svesci -- ne dira se" % ime)
                    preskocene.add(ime)
                    continue
                raise RuntimeError("Tabela %s u novoj svesci NIJE prazna -- upis odbijen "
                                   "(alat ne mesa podatke; napravi svezu svesku)." % ime)
            kolone = [str(x.Name) for x in lo.ListColumns]
            plan.append((ime, lo, kolone, niz_za_tabelu(redovi, kolone)))

        for ime, lo, kolone, niz in plan:
            _upisi_tabelu(lo, niz, kolone)
            log("  %-28s %6d red." % (ime, len(niz)))
        for ime, parovi in sorted(c.podesavanja.items()):
            lo = lo_po_imenu.get(ime.casefold())
            if lo is None or not parovi:
                continue
            samo_nova = upisi_podesavanja(lo, parovi, c.kljuc_podesavanja(ime))
            log("  %-28s %6d kljuceva iz stare sveske" % (ime, len(parovi)))
            for k, v in samo_nova:
                log("    samo u novoj svesci (proveri): %s = %s" % (k, v))
        wb.Save()
        log("Sacuvano: %s" % putanja)
        return preskocene
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass


def upisi_podesavanja(lo, parovi, kolona_kljuc):
    """Key/value tabela: kljuc iz stare sveske prepisuje vrednost nove (ili se
    dodaje); vrati kljuceve koji postoje SAMO u novoj svesci, sa vrednoscu --
    to su test vrednosti iz make_dev_workbook koje operater treba da proveri."""
    kolone = [str(x.Name) for x in lo.ListColumns]
    ci = {k.casefold(): i + 1 for i, k in enumerate(kolone)}
    ki = ci[kolona_kljuc.casefold()]
    postojeci = {}
    if lo.ListRows.Count > 0:
        for r in range(1, lo.ListRows.Count + 1):
            k = str(lo.DataBodyRange.Cells(r, ki).Value or "").strip()
            if k:
                postojeci[k.casefold()] = r
    stari = set()
    for k, vrednosti in parovi:
        stari.add(k.casefold())
        r = postojeci.get(k.casefold())
        if r is None:
            lo.ListRows.Add()
            r = lo.ListRows.Count
            postojeci[k.casefold()] = r
        for kol, v in vrednosti.items():
            if kol.casefold() in ci:
                lo.DataBodyRange.Cells(r, ci[kol.casefold()]).Value = com_vrednost(v)
    vi = None
    for x in ("vrednost", "configvalue"):
        if x in ci:
            vi = ci[x]
    samo_nova = []
    for k, r in sorted(postojeci.items()):
        if k not in stari:
            v = lo.DataBodyRange.Cells(r, vi).Value if vi else ""
            samo_nova.append((str(lo.DataBodyRange.Cells(r, ki).Value), v))
    return samo_nova


def _ima_podataka(lo):
    if lo.ListRows.Count == 0:
        return False
    v = lo.DataBodyRange.Value
    redovi = v if isinstance(v, tuple) else ((v,),)
    return any(any(x not in (None, "") for x in (r if isinstance(r, tuple) else (r,))) for r in redovi)


def _upisi_tabelu(lo, niz, kolone):
    n, m = len(niz), len(kolone)
    zag = lo.HeaderRowRange
    lo.Resize(zag.Resize(n + 1, m))
    telo = lo.DataBodyRange
    tekst_kf = {k.casefold() for k in TEKST_KOLONE}
    for j, k in enumerate(kolone, 1):
        if k.casefold() in tekst_kf:
            # Format PRE upisa: posle bi "101" vec bio broj.
            telo.Columns(j).NumberFormat = "@"
    for i in range(0, n, DELOVI):
        deo = niz[i:i + DELOVI]
        telo.Cells(i + 1, 1).Resize(len(deo), m).Value = tuple(deo)
