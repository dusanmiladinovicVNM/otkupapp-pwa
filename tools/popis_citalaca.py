#!/usr/bin/env python3
"""Popis citalaca starog modela dokumenta -- refaktor "dokument = header + stavke".

Samo CITA `src-vba/` (radno stablo ili git commit). Ne dira VBA, svesku ni git.
Nastao za PR7 (otpremnica cutover), docs/REFAKTOR_PR7_POPIS.md: isti instrument
meri populaciju na baznom commitu slajsa i prag "dual READ = 0" na njegovom kraju.

Sta meri
--------
1. ZATVOREN SPISAK MESTA -- grupe sidara (regex nad kodom bez komentara, bez
   modConfig/modSchema) -> fajl:linija:procedura. Osnovne grupe su iste kao u
   popisu od 15.09.2026 (uporedivost); prosirene grupe (prefiks `x_`) hvataju ono
   sto konstante ne vide: literale imena kolona, SaveOtkup*, VremeUnosa, trace
   kolone, prosledjene indekse kolona.
2. GRAF POZIVA -- imenovane ulazne tacke, ne pogadjanje:
     UI    dogadjaji frmOtkupUI, WithEvents handleri klasa, ugovor ekrana
           (Application.Run m & ".Scr_*" nad registrovanim modScr* modulima),
           registar panela ("KLJUC|modul|graditelj|...")
     WB    Workbook_* / Worksheet_* dogadjaji
     MAKRO javni Sub bez argumenata u produkcionom standardnom modulu (Alt+F8);
           ime Test_*/TestHook* je test hook, ne produkcija
   Kasno vezani pozivi se razresavaju iz stringova: "modul.Proc", "Proc"
   (OnTime/OnAction/QualifiedProc/CallOptional), ".Scr_X" (ugovor ekrana),
   "_Release"/"_ImaNesacuvano" (paneli), i string konstante modula.
3. KAPIJE PAUZE -- funkcije iz KAPIJE koje danas vracaju konstantno False, i
   izricito ugaseni pozivi iz UGASENI. Kapija se proverava pri svakom pokretanju:
   ako funkcija vise ne vraca False, ili poruka ugasenog poziva vise ne postoji,
   alat to prijavljuje kao UPOZORENJE i ne tretira mesto kao pauzirano.
4. STATUS -- procedure: ZIV_UI / ZIV_SYNC / ZIV_MAKRO / PAUZIRAN / SAMO_TEST /
   MRTAV (+ TEST za procedure test modula). Mesto nasledjuje status procedure,
   osim kad je iza kapije u samoj proceduri (tada PAUZIRAN).
5. DUAL READ -- zive reference na linijska polja zaglavlja (otkup: Kolicina,
   Cena, Klasa, KolAmbalaze, BrutoKg; otpremnica: Kolicina, Klasa, KolAmbalaze,
   BrutoKg). Referenca nije isto sto i citanje (pristup CITA/PISE/MAPA odredjuje
   klasifikacija u popisu), ali prag na kraju slajsa je strozi: nula zivih
   referenci, osim imenovanih izuzetaka.

Granice (imenovane, ne skrivene)
--------------------------------
- Razresavanje poziva je po imenu: lokalna procedura modula ima prednost, zatim
  javne procedure standardnih modula. Isto ime u vise modula daje ivice ka
  svima (oznaceno `visezn`). Clan objekta (`x.Metod`) vodi samo ka metodama klasa.
- Dinamicki sastavljeno ime koje nije ni u jednom stringu (npr.
  `Application.Run procName` sa imenom iz tabele) se NE vidi.
- Kapija se prepoznaje samo u oblicima: `If Not G() Then` sa Exit/Err.Raise/GoTo
  u grani (pauzira ostatak procedure ili do labele), grana `ElseIf`/`Else` posle
  `If [A And] Not G() Then`, i `If G() Then` (pauzira granu).

Upotreba (iz root-a repoa)
--------------------------
  python tools/popis_citalaca.py                    # sazetak nad radnim stablom
  python tools/popis_citalaca.py --commit a173c134  # sazetak nad commitom
  python tools/popis_citalaca.py --json izlaz.json  # pun izlaz (mesta, procedure, lanci)
  python tools/popis_citalaca.py --uporedi a173c134 # delta mesta: stari commit -> ovo stablo
  python tools/popis_citalaca.py --procedura modOtkupBlok.SumKolByOtp   # status + lanac
"""
import argparse
import collections
import json
import os
import re
import subprocess
import sys

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_DIR = "src-vba"
EXT = (".bas", ".cls", ".frm", ".doccls")
BEZ_SIDARA = {"modconfig", "modschema"}

# --- grupe sidara -----------------------------------------------------------
# Osnovne grupe: IDENTICNE popisu od 15.09.2026 (uporedivost brojeva).
GRUPE = collections.OrderedDict([
    ("otk_veze", re.compile(r"\bCOL_OTK_(OTPREMNICA_ID|BROJ_ZBIRNE|VOZAC|BROJ_OTPREMNICE)\b")),
    ("otk_linija", re.compile(r"\bCOL_OTK_(KOLICINA|CENA|KLASA|KOL_AMB|KOL_AMB_IZDATA|BRUTO|NOVAC|PRIMALAC|TIP_AMB)\b")),
    ("otp_linija", re.compile(r"\bCOL_OTP_(KOLICINA|KLASA|KOL_AMB|BRUTO|TIP_AMB|SORTA|VRSTA|KULTURA)\b")),
    ("otp_cena", re.compile(r"\bCOL_OTP_CENA\b")),
    ("otp_brojzbirne", re.compile(r"\bCOL_OTP_BROJ_ZBIRNE\b")),
    ("otp_stari_pisac", re.compile(r"\b(SaveOtpremnica(Multi)?(_TX)?|AutoLinkOtkupOtpremnica\w*)\b")),
    ("pauza", re.compile(r"\b(NapredakBlokaDostupan|IzvedeniLanacIzPwaDostupan)\b")),
    ("split_plus", re.compile(r"Split\([^)]*\"\s\+\s\"")),
])
# Prosirene grupe: ono sto konstante ne vide (kriticar pokrivenosti, 15.09).
GRUPE_X = collections.OrderedDict([
    ("x_saveotkup", re.compile(r"\bSaveOtkup(_TX)?\b")),
    ("x_vreme_unosa", re.compile(r"\bCOL_OTK_VREME_UNOSA\b")),
    ("x_trace", re.compile(r"\bCOL_TRACE_(ISPRAVKA_OD|ZAMENJEN_SA)\b")),
])
# Literal imena kolone u naredbi koja pominje tabelu starog modela.
LITERAL_KOLONE = {
    "otkup": ["OtpremnicaID", "BrojZbirne", "VozacID", "BrojOtpremnice", "Kolicina", "Cena",
              "Klasa", "KolAmbalaze", "BrutoKg", "Novac", "PrimalacNovca", "VremeUnosa"],
    "otpremnica": ["Kolicina", "Klasa", "KolAmbalaze", "BrutoKg", "Cena", "BrojZbirne",
                   "IspravkaOd", "ZamenjenSa"],
}
TABELA_U_NAREDBI = {
    "otkup": re.compile(r"\bTBL_OTKUP\b|\"tblOtkup\"", re.I),
    "otpremnica": re.compile(r"\bTBL_OTPREMNICA\b|\"tblOtpremnica\"", re.I),
}
# Indeks kolone starog modela koji se racuna u jednoj proceduri pa prosledjuje drugoj.
INDEKS_KOLONE = re.compile(
    r"^\s*(?:Set\s+)?(\w+)\s*=\s*(?:\w+\.)?(?:RequireColumnIndex|GetColumnIndex|ColIdx|RequireColIdx)\s*\(\s*"
    r"(TBL_OTKUP|TBL_OTPREMNICA)\s*,\s*(COL_OT[KP]_\w+)", re.I)
# Genericki pristupnici celiji (niz, red, kolona): indeks koji im se preda je
# isto sto i data(r, c) -- nije prosledjivanje ZNANJA o koloni drugoj proceduri.
PRISTUPNICI_CELIJI = {"modUiData.CellS", "modUiData.CellD", "modUiData.CellDate",
                      "modPaletniList.SafeCell", "modStornoFlow.NzTxC"}
KOLONE_STAROG_MODELA = re.compile(
    r"^COL_OTK_(OTPREMNICA_ID|BROJ_ZBIRNE|VOZAC|BROJ_OTPREMNICE|KOLICINA|CENA|KLASA|KOL_AMB|BRUTO|NOVAC|PRIMALAC|VREME_UNOSA)$"
    r"|^COL_OTP_(KOLICINA|KLASA|KOL_AMB|BRUTO|CENA|BROJ_ZBIRNE)$")

# Linijska polja zaglavlja -- prag dual READ.
DUAL_READ = re.compile(r"\bCOL_OTK_(KOLICINA|CENA|KLASA|KOL_AMB|BRUTO)\b|\bCOL_OTP_(KOLICINA|KLASA|KOL_AMB|BRUTO)\b")

# --- kapije pauze -----------------------------------------------------------
# (modul, funkcija): funkcija mora da vraca konstantno False, inace upozorenje.
KAPIJE = [("modMasterSync", "IzvedeniLanacIzPwaDostupan"),
          ("modOtkupBlok", "NapredakBlokaDostupan")]
# Izricito ugasen poziv: u modulu stoji poruka o pauzi umesto poziva.
# (modul sa porukom, kljuc poruke, ciljni modul, ciljna procedura)
UGASENI = [("modOtkupUnos", "OTKUNOS_MSG_LANAC_PAUZIRAN", "modAutoHladnjaca", "AutoChainHladnjaca")]

STATUS_RED = ["ZIV_UI", "ZIV_SYNC", "ZIV_MAKRO", "PAUZIRAN", "SAMO_TEST", "MRTAV", "TEST"]

KASNO_VEZANO = re.compile(r"\b(Run|OnTime|OnAction|CallByName|QualifiedProc|CallOptional|MacroOptions)\b", re.I)
PROC_RE = re.compile(r"^\s*(?:(Public|Private|Friend|Global)\s+)?(?:Static\s+)?"
                     r"(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)", re.I)
END_RE = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.I)
IDENT = re.compile(r"[A-Za-z_]\w*")
TEST_MODUL = re.compile(r"(^modTest|Tests?$)", re.I)
TEST_HOOK = re.compile(r"^(Test_|T_\d|Diag_)|TestHook|Test$|TestReset$|TestPrljav$", re.I)


def je_test_modul(name):
    return bool(TEST_MODUL.search(name))


# --- izvor ------------------------------------------------------------------
def ucitaj_izvor(commit=None):
    """Vraca {ime_fajla: [linije]}; iz git objekta kad je commit zadat."""
    out = {}
    if commit:
        ls = subprocess.run(["git", "-C", REPO, "ls-tree", "--name-only", commit, SRC_DIR + "/"],
                            capture_output=True, check=True).stdout.decode().split()
        for path in ls:
            f = path.split("/")[-1]
            if f.lower().endswith(EXT):
                raw = subprocess.run(["git", "-C", REPO, "show", "%s:%s" % (commit, path)],
                                     capture_output=True, check=True).stdout
                out[f] = raw.decode("latin-1").splitlines()
    else:
        d = os.path.join(REPO, SRC_DIR)
        for f in sorted(os.listdir(d)):
            if f.lower().endswith(EXT):
                with open(os.path.join(d, f), encoding="latin-1", newline="") as fh:
                    out[f] = fh.read().splitlines()
    return out


def bez_komentara(line):
    """Isto kao popis od 15.09: odseca komentar, stringove ZADRZAVA."""
    out, in_str = [], False
    for ch in line:
        if ch == '"':
            in_str = not in_str
        if ch == "'" and not in_str:
            break
        out.append(ch)
    s = "".join(out)
    return "" if re.match(r"^\s*Rem\b", s, re.I) else s


def razlozi(line):
    """(kod sa ispraznjenim stringovima, [stringovi]) -- komentar odsecen."""
    code, strs, cur, in_str, i, n = [], [], [], False, 0, len(line)
    while i < n:
        ch = line[i]
        if in_str:
            if ch == '"':
                if i + 1 < n and line[i + 1] == '"':
                    cur.append('"')
                    code.append("  ")
                    i += 2
                    continue
                in_str = False
                strs.append("".join(cur))
                cur = []
                code.append('"')
            else:
                cur.append(ch)
                code.append(" ")
            i += 1
            continue
        if ch == '"':
            in_str = True
            code.append('"')
        elif ch == "'":
            break
        else:
            code.append(ch)
        i += 1
    s = "".join(code)
    if re.match(r"^\s*Rem\b", s, re.I):
        return "", []
    return s, strs


class Proc(object):
    __slots__ = ("modul", "ime", "start", "end", "vid", "vrsta", "ima_arg", "lokalne", "kapije")

    def __init__(self, modul, ime, start, vid, vrsta):
        self.modul, self.ime, self.start, self.end = modul, ime, start, start
        self.vid, self.vrsta = (vid or "").lower(), vrsta.lower()
        self.ima_arg, self.lokalne, self.kapije = False, set(), []

    @property
    def kljuc(self):
        return "%s.%s" % (self.modul.ime, self.ime)


class Modul(object):
    def __init__(self, fajl, linije):
        self.fajl = fajl
        self.ime = os.path.splitext(fajl)[0]
        self.vrsta = os.path.splitext(fajl)[1].lower().lstrip(".")
        self.linije = linije
        self.test = je_test_modul(self.ime)
        self.procs = []
        self.po_liniji = {}           # linija -> Proc
        self.promenljive = set()      # modul-level promenljive i konstante (lower)
        self.withevents = set()       # imena WithEvents promenljivih (lower)
        self.konst_str = {}           # ime konstante (lower) -> string vrednost
        self.kod = []                 # (kod bez stringova, stringovi) po liniji
        for raw in linije:
            self.kod.append(razlozi(raw))
        self._procedure()

    def _procedure(self):
        cur = None
        for i, raw in enumerate(self.linije, 1):
            code = self.kod[i - 1][0]
            m = PROC_RE.match(code)
            if m and not re.match(r"^\s*(Public|Private)?\s*Declare\b", code, re.I):
                cur = Proc(self, m.group(3), i, m.group(1), re.sub(r"\s+", " ", m.group(2)))
                self.procs.append(cur)
                self.po_liniji[i] = cur
                # zaglavlje moze da se lomi na vise linija
                hdr, j = code, i
                while hdr.rstrip().endswith(" _") and j < len(self.linije):
                    hdr = hdr.rstrip()[:-1] + " " + self.kod[j][0]
                    self.po_liniji[j + 1] = cur
                    j += 1
                p = hdr.find("(", m.end(3))
                if p >= 0:
                    depth, k = 0, p
                    while k < len(hdr):
                        if hdr[k] == "(":
                            depth += 1
                        elif hdr[k] == ")":
                            depth -= 1
                            if depth == 0:
                                break
                        k += 1
                    args = hdr[p + 1:k]
                    cur.ima_arg = bool(args.strip())
                    for part in args.split(","):
                        w = re.sub(r"\b(Optional|ByVal|ByRef|ParamArray)\b", " ", part, flags=re.I).strip()
                        mm = IDENT.match(w)
                        if mm:
                            cur.lokalne.add(mm.group(0).lower())
                continue
            if cur is None:
                self._deklaracija(code, self.kod[i - 1][1])
            else:
                self.po_liniji[i] = cur
                dm = re.match(r"^\s*(Dim|Static|ReDim(?:\s+Preserve)?|Const)\s+(.*)$", code, re.I)
                if dm:
                    for part in _podeli_zarezom(dm.group(2)):
                        mm = IDENT.match(part.strip())
                        if mm:
                            cur.lokalne.add(mm.group(0).lower())
                fm = re.match(r"^\s*For\s+(?:Each\s+)?(\w+)", code, re.I)
                if fm:
                    cur.lokalne.add(fm.group(1).lower())
            if cur is not None and END_RE.match(code):
                cur.end = i
                self.po_liniji[i] = cur
                cur = None

    def _deklaracija(self, code, strs):
        m = re.match(r"^\s*(?:Public|Private|Global|Dim)\s+(?:(WithEvents)\s+)?(?:Const\s+)?(\w+)", code, re.I)
        if m and m.group(2).lower() not in ("sub", "function", "property", "declare", "type", "enum", "event"):
            self.promenljive.add(m.group(2).lower())
            if m.group(1):
                self.withevents.add(m.group(2).lower())
        c = re.match(r"^\s*(?:Public\s+|Private\s+|Global\s+)?Const\s+(\w+)\s+As\s+String\s*=\s*\"", code, re.I)
        if c and strs:
            self.konst_str[c.group(1).lower()] = strs[0]
        if m is None:
            m2 = re.match(r"^\s*Const\s+(\w+)", code, re.I)
            if m2:
                self.promenljive.add(m2.group(1).lower())


def _podeli_zarezom(s):
    out, depth, cur = [], 0, []
    for ch in s:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        if ch == "," and depth == 0:
            out.append("".join(cur))
            cur = []
        else:
            cur.append(ch)
    out.append("".join(cur))
    return out


def logicke_naredbe(modul, proc):
    """[(prva, poslednja, kod, [stringovi])] -- spaja ' _' nastavke."""
    out, i = [], proc.start
    while i <= proc.end:
        first, code, strs = i, modul.kod[i - 1][0], list(modul.kod[i - 1][1])
        while code.rstrip().endswith(" _") and i < proc.end:
            i += 1
            code = code.rstrip()[:-1] + " " + modul.kod[i - 1][0]
            strs += modul.kod[i - 1][1]
        out.append((first, i, code, strs))
        i += 1
    return out


# --- kapije -----------------------------------------------------------------
def _gate_rx(ime):
    return re.compile(r"(?<![\w.])(?:\w+\.)?" + ime + r"\b\s*(?:\(\s*\))?", re.I)


def nadji_kapije(moduli, upozorenja):
    """Proverava da kapije i dalje vracaju False; vraca {ime_lower: (modul, linija)}."""
    aktivne = {}
    for mod_ime, fn in KAPIJE:
        mod = moduli.get(mod_ime.lower())
        p = None
        if mod:
            p = next((x for x in mod.procs if x.ime.lower() == fn.lower()), None)
        if p is None:
            upozorenja.append("KAPIJA %s.%s ne postoji -- nije primenjena" % (mod_ime, fn))
            continue
        telo = [mod.kod[i - 1][0].strip() for i in range(p.start + 1, p.end)]
        telo = [t for t in telo if t]
        if telo != ["%s = False" % p.ime] and [t.lower() for t in telo] != [("%s = false" % p.ime).lower()]:
            upozorenja.append("KAPIJA %s.%s vise ne vraca konstantno False (%s) -- nije primenjena"
                              % (mod_ime, fn, " | ".join(telo)[:120]))
            continue
        aktivne[fn.lower()] = (mod.fajl, p.start)
    return aktivne


def oznaci_kapije(modul, proc, kapije):
    """Popunjava proc.kapije: [(od, do, opis)] -- linije iza kapije u proceduri."""
    if not kapije:
        return
    rxs = {k: _gate_rx(k) for k in kapije}
    nar = logicke_naredbe(modul, proc)
    labele = {}
    for (a, b, code, _) in nar:
        lm = re.match(r"^\s*(\w+):\s*$", code)
        if lm:
            labele[lm.group(1).lower()] = a

    def kapija_u(cond):
        for k, rx in rxs.items():
            for m in rx.finditer(cond):
                pre = cond[:m.start()].rstrip()
                neg = bool(re.search(r"\bNot$", pre, re.I))
                return k, neg
        return None, None

    def opis(k, line):
        return "%s (%s:%d)" % (k, modul.fajl, line)

    def izlaz_grane(stmts):
        for (a, b, code, _) in stmts:
            c = code.strip()
            if re.match(r"^(Exit\s+(Sub|Function|Property)|Err\.Raise\b)", c, re.I):
                return ("KRAJ", None)
            g = re.match(r"^GoTo\s+(\w+)", c, re.I)
            if g:
                return ("GOTO", g.group(1).lower())
        return (None, None)

    # blok-parser: stek If blokova; svaka grana pamti svoje naredbe na svom nivou
    BLOK_OTVARA = re.compile(r"^(For\b|Do\b|With\b|Select\s+Case\b|While\b)", re.I)
    BLOK_ZATVARA = re.compile(r"^(Next\b|Loop\b|End\s+With\b|End\s+Select\b|Wend\b)", re.I)
    stek = []   # {'grane': [[uslov, prva_linija, [naredbe]]], 'dubina': int}
    dubina = 0
    for (a, b, code, strs) in nar:
        c = code.strip()
        if not c or c.startswith("#"):
            continue
        if re.match(r"^If\b.*\bThen$", c, re.I):
            stek.append({"grane": [[c[2:-4].strip(), a, []]], "dubina": dubina})
            dubina += 1
            continue
        if re.match(r"^ElseIf\b.*\bThen$", c, re.I) and stek:
            stek[-1]["grane"].append([c[6:-4].strip(), a, []])
            continue
        if re.match(r"^Else$", c, re.I) and stek:
            stek[-1]["grane"].append(["", a, []])
            continue
        if re.match(r"^End\s+If$", c, re.I) and stek:
            blok = stek.pop()
            dubina -= 1
            grane = blok["grane"]
            prvi_uslov = grane[0][0]
            for idx, (uslov, prva, stmts) in enumerate(grane):
                kraj_grane = grane[idx + 1][1] - 1 if idx + 1 < len(grane) else a - 1
                k, neg = kapija_u(uslov)
                if k and not neg:
                    # uslov se izvrsava uvek; pauzirano je TELO grane
                    proc.kapije.append((prva + 1, kraj_grane, opis(k, prva)))
                if k and neg and idx == 0:
                    vrsta, lab = izlaz_grane(stmts)
                    if vrsta == "KRAJ":
                        proc.kapije.append((a + 1, proc.end, opis(k, prva)))
                    elif vrsta == "GOTO" and lab in labele:
                        proc.kapije.append((a + 1, labele[lab] - 1, opis(k, prva)))
                if idx > 0:
                    k0, neg0 = kapija_u(prvi_uslov)
                    if k0 and neg0:
                        norm0 = re.sub(r"\s+", " ", prvi_uslov).strip().lower()
                        normi = re.sub(r"\s+", " ", uslov).strip().lower()
                        m_and = re.match(r"^(.*)\s+and\s+not\s+(?:\w+\.)?" + k0 + r"\b\s*(?:\(\s*\))?$", norm0)
                        cist = re.match(r"^not\s+(?:\w+\.)?" + k0 + r"\b\s*(?:\(\s*\))?$", norm0)
                        if cist or (m_and and normi == m_and.group(1).strip()):
                            proc.kapije.append((prva + 1, kraj_grane, opis(k0, grane[0][1])))
            if stek:
                stek[-1]["grane"][-1][2].append((a, b, code, strs))
            continue
        if BLOK_OTVARA.match(c):
            dubina += 1
        elif BLOK_ZATVARA.match(c):
            dubina -= 1
        # jednolinijski If
        sm = re.match(r"^If\b(.*?)\bThen\b\s*(\S.*)$", c, re.I)
        if sm:
            k, neg = kapija_u(sm.group(1))
            if k and neg:
                vrsta, lab = izlaz_grane([(a, b, sm.group(2), [])])
                if vrsta == "KRAJ":
                    proc.kapije.append((b + 1, proc.end, opis(k, a)))
                elif vrsta == "GOTO" and lab in labele:
                    proc.kapije.append((b + 1, labele[lab] - 1, opis(k, a)))
            elif k:
                proc.kapije.append((a, b, opis(k, a)))
        if stek and dubina == stek[-1]["dubina"] + 1:
            stek[-1]["grane"][-1][2].append((a, b, code, strs))


def kapija_za_liniju(proc, linija):
    for (od, do, op) in proc.kapije:
        if od <= linija <= do:
            return op
    return ""


# --- graf poziva --------------------------------------------------------------
class Graf(object):
    def __init__(self, moduli, upozorenja):
        self.moduli = moduli
        self.upozorenja = upozorenja
        self.po_imenu = collections.defaultdict(list)     # ime lower -> [Proc]
        for m in moduli.values():
            for p in m.procs:
                self.po_imenu[p.ime.lower()].append(p)
        self.ivice = collections.defaultdict(list)        # kljuc -> [(ciljni kljuc, linija, kapija)]
        self.test_pozivaoci = collections.Counter()
        self.prod_pozivaoci = collections.Counter()
        self.registrovani = set()                         # moduli pomenuti u registru "A|mod|..."
        self.koreni = {}                                  # kljuc -> vrsta
        self.proc = {}
        for m in moduli.values():
            for p in m.procs:
                self.proc[p.kljuc.lower()] = p
        self._registri()
        self.kapije = nadji_kapije({k: v for k, v in moduli.items()}, upozorenja)
        for m in moduli.values():
            for p in m.procs:
                oznaci_kapije(m, p, self.kapije)
        self._ivice()
        self._ugaseni()
        self._koreni()

    def _registri(self):
        imena = {m.ime.lower() for m in self.moduli.values()}
        for m in self.moduli.values():
            if m.test:
                continue
            for code, strs in m.kod:
                for s in strs:
                    if "|" in s:
                        for t in s.split("|"):
                            if t.strip().lower() in imena:
                                self.registrovani.add(t.strip().lower())

    def _kandidati(self, ime, iz_modula):
        ime = ime.lower()
        lok = [p for p in iz_modula.procs if p.ime.lower() == ime]
        if lok:
            return lok
        return [p for p in self.po_imenu.get(ime, [])
                if p.modul.vrsta == "bas" and p.vid != "private"]

    def _string_ciljevi(self, s, iz_modula):
        # "'Sveska.xlsm'!Proc" -- literal je cesto samo "'!Proc" (ime sveske se lepi)
        s = re.sub(r"^'[^!]*!", "", s.strip())
        out = []
        m = re.match(r"^(\w+)\.(\w+)$", s)
        if m:
            mod = self.moduli.get(m.group(1).lower())
            if mod:
                out += [p for p in mod.procs if p.ime.lower() == m.group(2).lower()]
            return out
        if re.match(r"^\w+$", s):
            ime = s.lower()
            out += [p for p in self.po_imenu.get(ime, [])
                    if p.modul.vrsta in ("bas", "cls") and (p.vid != "private" or p.modul is iz_modula)]
            return out
        m = re.match(r"^\.(\w+)$", s)
        if m:
            for mod in self.moduli.values():
                if mod.ime.lower().startswith("modscr") and mod.ime.lower() in self.registrovani:
                    out += [p for p in mod.procs if p.ime.lower() == m.group(1).lower()]
            return out
        m = re.match(r"^_(\w+)$", s)
        if m:
            for mod in self.moduli.values():
                if mod.ime.lower() in self.registrovani:
                    cilj = (mod.ime[3:] + "_" + m.group(1)).lower()
                    out += [p for p in mod.procs if p.ime.lower() == cilj]
            return out
        if "|" in s:
            toks = [t.strip() for t in s.split("|")]
            for i, t in enumerate(toks):
                mod = self.moduli.get(t.lower())
                if mod:
                    for t2 in toks[i + 1:]:
                        out += [p for p in mod.procs if p.ime.lower() == t2.lower()]
        return out

    def _dodaj(self, p, cilj, linija):
        if cilj is p:
            return
        kap = kapija_za_liniju(p, linija)
        self.ivice[p.kljuc.lower()].append((cilj.kljuc.lower(), linija, kap))
        if p.modul.test:
            self.test_pozivaoci[cilj.kljuc.lower()] += 1
        else:
            self.prod_pozivaoci[cilj.kljuc.lower()] += 1

    def _ivice(self):
        for m in self.moduli.values():
            for p in m.procs:
                for (a, b, code, strs) in logicke_naredbe(m, p):
                    if a == p.start:
                        continue
                    self._ivice_naredbe(m, p, a, code, strs)

    def _ivice_naredbe(self, m, p, linija, code, strs):
        # String je poziv SAMO u kontekstu kasnog vezivanja ili kao red registra.
        # "modX.Proc" u LogError/Const SRC je ime za log, ne poziv.
        kasno = bool(KASNO_VEZANO.search(code))
        for s in strs:
            if kasno or "|" in s:
                for cilj in self._string_ciljevi(s, m):
                    self._dodaj(p, cilj, linija)
        toks = list(IDENT.finditer(code))
        for idx, t in enumerate(toks):
            ime = t.group(0)
            low = ime.lower()
            pre = code[:t.start()].rstrip()
            posle = code[t.end():].lstrip()
            if pre.endswith("."):
                prefiks = IDENT.findall(pre[:-1])
                pf = prefiks[-1].lower() if prefiks and re.search(r"\w$", pre[:-1]) else ""
                if pf == "me":
                    for c in m.procs:
                        if c.ime.lower() == low:
                            self._dodaj(p, c, linija)
                elif pf in self.moduli:
                    for c in self.moduli[pf].procs:
                        if c.ime.lower() == low:
                            self._dodaj(p, c, linija)
                else:
                    for c in self.po_imenu.get(low, []):
                        if c.modul.vrsta == "cls" and c.vid != "private":
                            self._dodaj(p, c, linija)
                continue
            if posle.startswith(".") and low in self.moduli:
                continue
            if low in p.lokalne or low in m.promenljive:
                if kasno and low in m.konst_str:
                    for cilj in self._string_ciljevi(m.konst_str[low], m):
                        self._dodaj(p, cilj, linija)
                continue
            if idx > 0 and toks[idx - 1].group(0).lower() == "new":
                cls = self.moduli.get(low)
                if cls:
                    # Instanca klase oziveljava njen konstruktor i WithEvents
                    # handlere -- handler NIJE ulazna tacka sam po sebi: klasa
                    # koju niko ziv ne instancira nema dogadjaja.
                    for c in cls.procs:
                        cl = c.ime.lower()
                        if cl in ("class_initialize", "class_terminate") or \
                                ("_" in cl and cl.split("_")[0] in cls.withevents):
                            self._dodaj(p, c, linija)
                continue
            for c in self._kandidati(ime, m):
                self._dodaj(p, c, linija)

    def _ugaseni(self):
        for (mod_ime, kljuc, cilj_mod, cilj_proc) in UGASENI:
            mod = self.moduli.get(mod_ime.lower())
            cmod = self.moduli.get(cilj_mod.lower())
            cilj = next((x for x in cmod.procs if x.ime.lower() == cilj_proc.lower()), None) if cmod else None
            nadjeno = False
            if mod and cilj:
                for i, (code, strs) in enumerate(mod.kod, 1):
                    if kljuc in strs and i in mod.po_liniji:
                        p = mod.po_liniji[i]
                        self.ivice[p.kljuc.lower()].append(
                            (cilj.kljuc.lower(), i, "ugasen poziv (%s:%d)" % (mod.fajl, i)))
                        self.prod_pozivaoci[cilj.kljuc.lower()] += 1
                        nadjeno = True
            if not nadjeno:
                self.upozorenja.append("UGASEN POZIV %s -> %s.%s: poruka %s nije nadjena -- nije primenjen"
                                       % (mod_ime, cilj_mod, cilj_proc, kljuc))

    def _koreni(self):
        for m in self.moduli.values():
            if m.test:
                continue
            for p in m.procs:
                low = p.ime.lower()
                if m.vrsta == "frm" and "_" in p.ime:
                    self.koreni[p.kljuc.lower()] = "UI"
                elif m.vrsta == "doccls" and re.match(r"^(workbook|worksheet)_", low):
                    self.koreni[p.kljuc.lower()] = "WB"
                elif (m.vrsta == "bas" and p.vrsta == "sub" and p.vid in ("", "public")
                      and not p.ima_arg):
                    # Makro je samo Sub na koji se ne poziva nijedan PRODUKCIONI kod
                    # (ni string). Pomocni Sub bez argumenata koji zove mrtav kod nije
                    # ulazna tacka -- inace bi ceo mrtav panel "ziveo" preko Alt+F8.
                    # Test pozivalac ne diskvalifikuje: test sme da zove pravi makro.
                    k = p.kljuc.lower()
                    if self.prod_pozivaoci[k]:
                        continue
                    self.koreni[k] = "MAKRO_TEST" if TEST_HOOK.search(p.ime) else "MAKRO"

    def statusi(self):
        """{kljuc: (status, lanac[(kljuc, linija)], kapija)}"""
        prod = {k for k, p in self.proc.items() if not p.modul.test}
        res = {}

        def bfs(koreni, sa_kapijom):
            pred = {}
            red = collections.deque()
            for k in koreni:
                pred[k] = (None, 0, "")
                red.append(k)
            while red:
                k = red.popleft()
                for (c, ln, kap) in self.ivice.get(k, []):
                    if c not in prod or c in pred:
                        continue
                    if kap and not sa_kapijom:
                        continue
                    pred[c] = (k, ln, kap if kap else pred[k][2])
                    red.append(c)
            return pred

        def lanac(pred, k):
            out, cur, seen = [], k, set()
            while cur is not None and cur not in seen:
                seen.add(cur)
                prev, ln, _ = pred[cur]
                out.append((cur, ln))
                cur = prev
            return list(reversed(out))

        redosled = [("UI", "ZIV_UI"), ("WB", "ZIV_SYNC"), ("MAKRO", "ZIV_MAKRO")]
        zivi = {}
        for vrsta, status in redosled:
            pred = bfs([k for k, v in self.koreni.items() if v == vrsta], False)
            for k in pred:
                if k not in zivi:
                    zivi[k] = (status, lanac(pred, k), "")
        sve = bfs([k for k, v in self.koreni.items() if v in ("UI", "WB", "MAKRO")], True)
        # SAMO_TEST: dostizno iz test modula ili test hook makroa (ukljucujuci
        # produkcione procedure koje zove samo test-only procedura).
        test_ulaz = [k for k, v in self.koreni.items() if v == "MAKRO_TEST"]
        for k, iv in self.ivice.items():
            if self.proc[k].modul.test:
                test_ulaz += [c for (c, _, _) in iv if c in prod]
        test_dost = bfs(test_ulaz, True)
        for k, p in self.proc.items():
            if p.modul.test:
                res[k] = ("TEST", [], "")
            elif k in zivi:
                res[k] = zivi[k]
            elif k in sve:
                res[k] = ("PAUZIRAN", lanac(sve, k), sve[k][2])
            elif k in test_dost:
                res[k] = ("SAMO_TEST", [], "")
            else:
                res[k] = ("MRTAV", [], "")
        return res


# --- mesta ------------------------------------------------------------------
def mesta(moduli, graf, sa_prosirenim=True):
    out = []
    for m in sorted(moduli.values(), key=lambda x: x.fajl.lower()):
        if m.ime.lower() in BEZ_SIDARA:
            continue
        tip = "TEST" if m.test else "PROD"
        for i, raw in enumerate(m.linije, 1):
            code15 = bez_komentara(raw)
            p = m.po_liniji.get(i)
            pime = p.ime if p else "(deklaracije)"
            for g, rx in GRUPE.items():
                if rx.search(code15):
                    out.append(_mesto(g, tip, m, i, pime, raw, rx.findall(code15) and _izrazi(rx, code15)))
            if not sa_prosirenim:
                continue
            for g, rx in GRUPE_X.items():
                if rx.search(code15):
                    out.append(_mesto(g, tip, m, i, pime, raw, _izrazi(rx, code15)))
        if sa_prosirenim:
            out += _literali(m, tip)
            out += _indeksi(m, tip, graf)
    return out


def _izrazi(rx, code):
    return sorted({mm.group(0) for mm in rx.finditer(code)})


def _mesto(g, tip, m, i, pime, raw, izrazi):
    return {"grupa": g, "tip": tip, "fajl": m.fajl, "modul": m.ime, "linija": i,
            "procedura": pime, "izrazi": izrazi or [], "kod": raw.strip()}


def _literali(m, tip):
    out = []
    for p in m.procs:
        for (a, b, code, strs) in logicke_naredbe(m, p):
            for tab, rx in TABELA_U_NAREDBI.items():
                if not rx.search(code) and not any(("tbl" + tab).lower() == s.lower() for s in strs):
                    continue
                for ln in range(a, b + 1):
                    lstrs = m.kod[ln - 1][1]
                    hit = sorted({s for s in lstrs if s in LITERAL_KOLONE[tab]})
                    if hit:
                        out.append(_mesto("x_literal", tip, m, ln, p.ime, m.linije[ln - 1],
                                          ["%s.\"%s\"" % (tab, h) for h in hit]))
    # ista linija moze da padne u obe tabele -- jedan red po liniji
    spojeno = collections.OrderedDict()
    for r in out:
        k = (r["fajl"], r["linija"])
        if k in spojeno:
            spojeno[k]["izrazi"] = sorted(set(spojeno[k]["izrazi"]) | set(r["izrazi"]))
        else:
            spojeno[k] = r
    return list(spojeno.values())


def _pozivi_sa_argumentima(code, graf, m, p):
    """[(ciljna procedura, [argumenti na vrhu])] za pozive korisnickih procedura u naredbi."""
    out = []
    for t in IDENT.finditer(code):
        ime, low = t.group(0), t.group(0).lower()
        pre = code[:t.start()].rstrip()
        if pre.endswith("."):
            pf = IDENT.findall(pre[:-1])
            pf = pf[-1].lower() if pf else ""
            cilj = [c for c in graf.moduli[pf].procs if c.ime.lower() == low] if pf in graf.moduli else []
        elif low in p.lokalne or low in m.promenljive:
            cilj = []
        else:
            cilj = graf._kandidati(ime, m)
        cilj = [c for c in cilj if c is not p]
        if not cilj:
            continue
        rest = code[t.end():]
        s = rest.lstrip()
        if s.startswith("("):
            depth, k = 0, 0
            for k, ch in enumerate(s):
                if ch == "(":
                    depth += 1
                elif ch == ")":
                    depth -= 1
                    if depth == 0:
                        break
            args = s[1:k]
        elif re.match(r"^\s*(Call\s+)?(\w+\.)?$", code[:t.start()], re.I):
            args = rest
        else:
            continue
        out.append((cilj[0], [a.strip() for a in _podeli_zarezom(args)]))
    return out


def _indeksi(m, tip, graf):
    """Indeks kolone starog modela prosledjen DRUGOJ proceduri kao argument."""
    out = []
    for p in m.procs:
        prom = {}
        nar = logicke_naredbe(m, p)
        for (a, b, code, strs) in nar:
            im = INDEKS_KOLONE.match(code)
            if im and KOLONE_STAROG_MODELA.match(im.group(3)):
                prom[im.group(1).lower()] = im.group(3)
        if not prom:
            continue
        for (a, b, code, strs) in nar:
            if INDEKS_KOLONE.match(code):
                continue
            nosi = []
            for cilj, args in _pozivi_sa_argumentima(code, graf, m, p):
                if "%s.%s" % (cilj.modul.ime, cilj.ime) in PRISTUPNICI_CELIJI:
                    continue
                for arg in args:
                    arg = re.sub(r"^(ByVal|ByRef)\s+", "", arg, flags=re.I).strip()
                    if arg.lower() in prom:
                        nosi.append("%s -> %s.%s" % (prom[arg.lower()], cilj.modul.ime, cilj.ime))
            if nosi:
                out.append(_mesto("x_indeks", tip, m, a, p.ime, m.linije[a - 1], sorted(set(nosi))))
    return out


def napravi(commit=None):
    izvor = ucitaj_izvor(commit)
    moduli = collections.OrderedDict()
    for f, linije in izvor.items():
        mod = Modul(f, linije)
        moduli[mod.ime.lower()] = mod
    upozorenja = []
    graf = Graf(moduli, upozorenja)
    st = graf.statusi()
    spisak = mesta(moduli, graf)
    for r in spisak:
        m = moduli[r["modul"].lower()]
        p = m.po_liniji.get(r["linija"])
        if p is None:
            r.update(status_procedure="DEKLARACIJA", status_mesta="DEKLARACIJA", kapija="")
            continue
        s, lan, kap = st[p.kljuc.lower()]
        # Linija koja SAMA proverava kapiju izvrsava se uvek.
        kap_m = "" if r["grupa"] == "pauza" else kapija_za_liniju(p, r["linija"])
        r["status_procedure"] = s
        r["status_mesta"] = "PAUZIRAN" if (kap_m and s.startswith("ZIV")) else s
        r["kapija"] = kap_m or kap
    return moduli, graf, st, spisak, upozorenja


# --- izvestaji --------------------------------------------------------------
def fmt_lanac(graf, lan):
    """Svaki korak je `modul.Procedura:linija` gde je linija POZIV sledeceg koraka
    u toj proceduri; poslednji korak (procedura koja sadrzi mesto) nema liniju."""
    out = []
    for i, (k, _) in enumerate(lan):
        p = graf.proc[k]
        ln = lan[i + 1][1] if i + 1 < len(lan) else 0
        out.append("%s.%s%s" % (p.modul.ime, p.ime, (":%d" % ln) if ln else ""))
    return out


def sazetak(moduli, graf, st, spisak, upozorenja, commit):
    print("popis_citalaca @ %s" % (commit or "radno stablo"))
    kor = collections.Counter(graf.koreni.values())
    print("ulazne tacke: " + ", ".join("%s=%d" % kv for kv in sorted(kor.items())))
    print("kapije: " + (", ".join("%s (%s:%d)" % (k, v[0], v[1]) for k, v in graf.kapije.items()) or "nema"))
    print()
    grupe = list(GRUPE) + list(GRUPE_X) + ["x_literal", "x_indeks"]
    print("%-16s %5s %5s | %s" % ("grupa", "PROD", "TEST", "PROD po statusu mesta"))
    for g in grupe:
        rows = [r for r in spisak if r["grupa"] == g]
        prod = [r for r in rows if r["tip"] == "PROD"]
        c = collections.Counter(r["status_mesta"] for r in prod)
        print("%-16s %5d %5d | %s" % (g, len(prod), len(rows) - len(prod),
                                      ", ".join("%s=%d" % (s, c[s]) for s in STATUS_RED + ["DEKLARACIJA"] if c[s])))
    osnovne = [r for r in spisak if r["grupa"] in GRUPE]
    print("%-16s %5d %5d" % ("osnovne ukupno", sum(1 for r in osnovne if r["tip"] == "PROD"),
                             sum(1 for r in osnovne if r["tip"] == "TEST")))
    print()
    dr = [r for r in spisak if r["tip"] == "PROD" and r["grupa"] in ("otk_linija", "otp_linija")
          and DUAL_READ.search(bez_komentara(r["kod"]))]
    ziv = [r for r in dr if r["status_mesta"].startswith("ZIV")]
    print("DUAL READ (reference na linijska polja zaglavlja, PROD): %d mesta; zivih %d u %d procedura"
          % (len(dr), len(ziv), len({(r["modul"], r["procedura"]) for r in ziv})))
    c = collections.Counter(r["status_mesta"] for r in dr)
    print("  po statusu: " + ", ".join("%s=%d" % (s, c[s]) for s in STATUS_RED if c[s]))
    if upozorenja:
        print()
        for u in upozorenja:
            print("UPOZORENJE: " + u)


def uporedi(stari, novi):
    def kljuc(r):
        return (r["grupa"], r["modul"], r["procedura"], re.sub(r"\s+", " ", r["kod"]))
    ca = collections.Counter(kljuc(r) for r in stari)
    cb = collections.Counter(kljuc(r) for r in novi)
    nestalo = ca - cb
    novo = cb - ca
    print("mesta: stari %d, novi %d, nestalo %d, novo %d, isto (moguce pomereno) %d"
          % (len(stari), len(novi), sum(nestalo.values()), sum(novo.values()), sum((ca & cb).values())))
    po = collections.defaultdict(lambda: [0, 0])
    for k, v in nestalo.items():
        po[(k[0], k[1], k[2])][0] += v
    for k, v in novo.items():
        po[(k[0], k[1], k[2])][1] += v
    for (g, mod, proc), (n0, n1) in sorted(po.items()):
        print("  %-15s %s.%s  -%d +%d" % (g, mod, proc, n0, n1))


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--commit", help="meri git commit umesto radnog stabla")
    ap.add_argument("--json", help="pun izlaz u JSON fajl")
    ap.add_argument("--uporedi", metavar="STARI", help="delta mesta: STARI commit -> merena verzija")
    ap.add_argument("--procedura", action="append", default=[], help="modul.Procedura: status, lanac, pozivaoci")
    a = ap.parse_args()

    moduli, graf, st, spisak, upozorenja = napravi(a.commit)
    if a.procedura:
        for k in a.procedura:
            s = st.get(k.lower())
            if s is None:
                print("%s: nema takve procedure" % k)
                continue
            p = graf.proc[k.lower()]
            print("%s (%s:%d-%d) status=%s kapija=%s prod_pozivaoci=%d test_pozivaoci=%d koren=%s"
                  % (k, p.modul.fajl, p.start, p.end, s[0], s[2] or "-", graf.prod_pozivaoci[k.lower()],
                     graf.test_pozivaoci[k.lower()], graf.koreni.get(k.lower(), "-")))
            if s[1]:
                print("  lanac: " + " -> ".join(fmt_lanac(graf, s[1])))
            if p.kapije:
                print("  kapije u proceduri: " + "; ".join("%d-%d %s" % x for x in p.kapije))
            poz = sorted({(q, ln, kap) for q, iv in graf.ivice.items() for (c, ln, kap) in iv if c == k.lower()})
            for q, ln, kap in poz[:40]:
                qp = graf.proc[q]
                print("  <- %s.%s:%d%s" % (qp.modul.ime, qp.ime, ln, (" [%s]" % kap) if kap else ""))
        return
    if a.uporedi:
        _, _, _, stari, _ = napravi(a.uporedi)
        uporedi(stari, spisak)
        return
    sazetak(moduli, graf, st, spisak, upozorenja, a.commit)
    if a.json:
        izlaz = {
            "commit": a.commit or "radno stablo",
            "grupe": {k: v.pattern for k, v in list(GRUPE.items()) + list(GRUPE_X.items())},
            "kapije": {k: "%s:%d" % v for k, v in graf.kapije.items()},
            "upozorenja": upozorenja,
            "procedure": {},
            "mesta": spisak,
        }
        potrebne = {(r["modul"].lower() + "." + r["procedura"].lower()) for r in spisak}
        for k, (s, lan, kap) in st.items():
            if k in potrebne or graf.koreni.get(k):
                p = graf.proc[k]
                izlaz["procedure"][p.kljuc] = {
                    "fajl": p.modul.fajl, "od": p.start, "do": p.end, "status": s, "kapija": kap,
                    "lanac": fmt_lanac(graf, lan), "koren": graf.koreni.get(k, ""),
                    "prod_pozivaoci": graf.prod_pozivaoci[k], "test_pozivaoci": graf.test_pozivaoci[k],
                    "kapije_u_proceduri": ["%d-%d %s" % x for x in p.kapije]}
        with open(a.json, "w", encoding="utf-8", newline="\n") as fh:
            json.dump(izlaz, fh, ensure_ascii=True, indent=1)
        print("\nJSON: %s (%d mesta, %d procedura)" % (a.json, len(spisak), len(izlaz["procedure"])))


if __name__ == "__main__":
    sys.exit(main())
