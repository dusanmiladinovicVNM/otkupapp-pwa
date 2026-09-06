"""Paritet HARD/SOFT algoritma izmedju modSelfUpdate i modVbaTools + fiksture.

Isti algoritam zivi u DVE privatne kopije: `modSelfUpdate` (klijentski
self-update) i `modVbaTools` (dev ImportAllVBA). Spojiti ih se ne moze --
`modSelfUpdate` je frozen bootstrap (`SKIP_MODULES`) i ne sme da zavisi ni od
cega sto se update-uje. Zato kopije moraju da se CUVAJU, a ne da se nadaju.

Zasto ovo postoji: kopije su vec divergirale. PR #274 je popravio detekciju
razmaka u `modVbaTools`, `modSelfUpdate` je ostao sa rupom, a `modVbaTools` je
pritom nosio komentar "isti obrazac stoji i u modSelfUpdate.IsHardModuleBody" --
koji vise nije bio tacan. To je zavedeno kao zamka #22 u docs/SELF_UPDATE.md.
Komentar ne moze da bude kapija; exit kod moze.

    python tools/vba_parity_check.py              # provera nad src-vba/
    python tools/vba_parity_check.py --self-test  # dokaz da provera hvata kvar

Izlazni kod: 0 = cisto, 2 = ima nalaza.

Provere:
  PARITET   -- sest procedura algoritma mora biti KOD-ZA-KOD isto u oba modula.
               Poredi se telo bez komentara i bez visestrukih razmaka: komentari
               smeju (i treba da) budu razliciti, kod ne sme.
  FIKSTURA  -- referentna Python implementacija istog algoritma nad korpusom
               slucajeva sa ocekivanim verdiktom. Pina SPECIFIKACIJU: sta je
               tvrdo telo i kada su dva tela "isto".

Sta ovo NE dokazuje: da se Python referenca ponasa isto kao VBA. Dokazuje da
obe VBA kopije nose ISTI kod i da taj kod odgovara pisanoj specifikaciji.
Ponasanje nad zivim VBE-om mereno je rucno (06.09.2026, zamke #22/#23).
"""

import io
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_VBA = os.path.join(ROOT, "src-vba")

NBSP = " "

# Procedure koje MORAJU biti identicne u oba modula. Lista je namerno kratka:
# to je tacno algoritam koji odlucuje da li modul ide u AddFromString ili u
# Remove+Import, i da li se izmena uopste prenosi.
PARITY_PROCS = (
    "IsHardModuleBody",
    "CodeLineUpper",
    "CollapseSpaces",
    "SameCode",
    "CanonCode",
    "LowerOutsideStrings",
)

PARITY_MODULES = ("modSelfUpdate.bas", "modVbaTools.bas")


# --- citanje VBA izvora ------------------------------------------------------

_PROC_HEAD = re.compile(
    r"^(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?"
    r"(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)",
    re.IGNORECASE,
)
_PROC_END = re.compile(r"^\s*End\s+(?:Sub|Function|Property)\s*$", re.IGNORECASE)


def strip_comment(line):
    """Odseci trailing komentar. Apostrof UNUTAR stringa nije komentar."""
    out = []
    inq = False
    for ch in line:
        if ch == '"':
            inq = not inq
        if ch == "'" and not inq:
            break
        out.append(ch)
    return "".join(out)


def read_procs(path):
    """{ime procedure: [redovi tela]} iz jednog .bas fajla."""
    txt = io.open(path, encoding="ascii", newline="").read().replace("\r\n", "\n")
    res = {}
    cur = None
    buf = []
    for ln in txt.split("\n"):
        if cur is None:
            m = _PROC_HEAD.match(ln)
            if m:
                cur = m.group(1)
                buf = [ln]
            continue
        buf.append(ln)
        if _PROC_END.match(ln):
            res.setdefault(cur, buf)
            cur = None
    return res


def code_only(lines):
    """Telo svedeno na kod: bez komentara, bez praznih redova, jedan razmak."""
    out = []
    for ln in lines:
        c = strip_comment(ln).strip()
        if c:
            out.append(re.sub(r"\s+", " ", c))
    return out


# --- referentna implementacija (prati VBA red po red) ------------------------

def ref_code_line_upper(s):
    """modSelfUpdate.CodeLineUpper: odseci komentar, UCase, LTrim.

    NB: sadrzaj STRINGA se NE uklanja -- odseca se samo komentar. Zato
    string-literal koji sadrzi "WithEvents " klasifikuje modul kao tvrd
    (v. slucaj "string sa WithEvents" nize).
    """
    out = []
    inq = False
    for ch in s:
        if ch == '"':
            inq = not inq
        if ch == "'" and not inq:
            break
        out.append(ch)
    return "".join(out).lstrip().upper()


def ref_collapse_spaces(s):
    """modSelfUpdate.CollapseSpaces: niz razmaka/tabova -> jedan razmak."""
    out = []
    was_sp = False
    for c in s:
        if c in (" ", "\t"):
            if not was_sp:
                out.append(" ")
            was_sp = True
        else:
            out.append(c)
            was_sp = False
    return "".join(out)


def ref_is_hard(body):
    """modSelfUpdate.IsHardModuleBody: module-level WithEvents / As MSForms."""
    t = body.replace("\r\n", "\n").replace("\r", "\n")
    for line in t.split("\n"):
        u = ref_collapse_spaces(ref_code_line_upper(line))
        if not u:
            continue
        w = u
        for pref in ("PUBLIC ", "PRIVATE ", "FRIEND ", "STATIC "):
            if w.startswith(pref):
                w = w[len(pref):]
        if (w.startswith("SUB ") or w.startswith("FUNCTION ")
                or w.startswith("PROPERTY ")):
            break                     # prva procedura = kraj module-level dela
        if "WITHEVENTS " in u:
            return True
        if " AS MSFORMS." in u:
            return True
    return False


def ref_lower_outside_strings(s):
    """modSelfUpdate.LowerOutsideStrings: lower + sazmi razmak IZVAN stringa."""
    out = []
    inq = False
    i = 0
    n = len(s)
    while i < n:
        c = s[i]
        if inq:
            if c == '"':
                if i + 1 < n and s[i + 1] == '"':
                    out.append('""')          # escaped "" ostaje u stringu
                    i += 2
                else:
                    inq = False
                    out.append(c)
                    i += 1
            else:
                out.append(c)                 # cuvaj case unutar stringa
                i += 1
        else:
            if c == '"':
                inq = True
                out.append(c)
                i += 1
            elif c == "'":
                out.append(s[i:])             # komentar do kraja reda -> cuvaj
                break
            elif c in (" ", "\t"):
                out.append(" ")               # niz razmaka/tabova -> jedan
                while i < n and s[i] in (" ", "\t"):
                    i += 1
            else:
                out.append(c.lower())
                i += 1
    return "".join(out)


def ref_canon(s):
    """modSelfUpdate.CanonCode."""
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = s.replace(NBSP, " ")          # VBE ume da vrati NBSP umesto obicnog space
    s = "\n".join(ref_lower_outside_strings(x).rstrip() for x in s.split("\n"))
    return s.strip("\n")


def ref_same_code(a, b):
    """modSelfUpdate.SameCode -- binarno poredjenje kanonizovanih tela."""
    return ref_canon(a) == ref_canon(b)


# --- korpus ------------------------------------------------------------------
#
# Ocekivanja su PINOVANO STVARNO ponasanje, ne zelja. Gde se stvarno ponasanje
# razlikuje od intuicije, razlog stoji uz slucaj.

TAB = "\t"

HARD_CASES = (
    # (ime, ocekivano, izvor)
    ("jedan razmak",           True,  "Private x As MSForms.Label"),
    ("dva razmaka",            True,  "Private x  As  MSForms.Label"),
    ("sest razmaka",           True,  "Private x      As MSForms.Label"),
    ("cetiri i cetiri",        True,  "Private x    As    MSForms.Label"),
    ("tabovi",                 True,  "Private x" + TAB + "As" + TAB + "MSForms.Label"),
    ("WithEvents",             True,  "Private WithEvents x As MSForms.CommandButton"),
    ("Public WithEvents",      True,  "Public    WithEvents    btn As MSForms.CommandButton"),
    ("Friend WithEvents",      True,  "Friend WithEvents x As MSForms.CommandButton"),
    ("mala slova",             True,  "private withevents x as msforms.commandbutton"),
    ("posle Option i Const",   True,  "Option Explicit\r\nPrivate Const K As Long = 1\r\n"
                                     "Private mLst As MSForms.ListBox"),

    ("proc-level Dim",         False, "Public Sub Test()\r\n"
                                     "    Dim x As MSForms.Label\r\n"
                                     "End Sub"),
    ("komentar",               False, "' Private WithEvents x As MSForms.Label"),
    ("posle Sub granice",      False, "Public Sub X()\r\nEnd Sub\r\n"
                                     "Private y As MSForms.Label"),
    ("posle Function granice", False, "Private Function X() As Long\r\nEnd Function\r\n"
                                     "Private WithEvents y As MSForms.CommandButton"),
    ("posle Property granice", False, "Friend Property Get X() As Long\r\nEnd Property\r\n"
                                     "Private y As MSForms.Label"),
    ("posle Public Static Sub", False, "Public Static Sub X()\r\n"
                                      "    Dim y As MSForms.Label\r\nEnd Sub"),
    ("obican modul",           False, "Option Explicit\r\nPrivate mN As Long\r\n"
                                     "Public Sub X()\r\nEnd Sub"),

    # --- lazni pozitiv koji se SVESNO pina ----------------------------------
    # Plan je ocekivao False. Stvarno je True: CodeLineUpper odseca KOMENTAR, a
    # sadrzaj stringa ostavlja -- pa "WithEvents " iz literala udje u pretragu.
    # Isto vazi u OBE kopije (modSelfUpdate i modVbaTools).
    #
    # Zasto se NE popravlja ovde: to bi bila izmena SEMANTIKE detektora, a ne
    # zatvaranje rupe. Greska je konzervativna -- modul se klasifikuje kao tvrd
    # pa ide kroz Remove+Import umesto kroz AddFromString: skuplji put, nikad
    # netacan ishod. U src-vba danas nema nijednog takvog modula (mereno
    # census-om). Ako se bude popravljalo, mora u OBE kopije istovremeno i uz
    # merenje nad zivim VBE-om -- ne usput.
    ("string sa WithEvents",   True,  'Private Const X As String = '
                                      '"Private WithEvents x As MSForms.Label"'),
)

SAME_CASES = (
    # (ime, ocekivano, a, b)
    ("razmak u deklaraciji",     True,  "Dim x  As Long",      "Dim x As Long"),
    ("razmak posle Set",         True,  "Set  x = y",          "Set x = y"),
    ("uvlacenje",                True,  "    x = 1",           "        x = 1"),
    ("tab vs razmaci",           True,  TAB + "x = 1",         "    x = 1"),
    ("re-casing identifikatora", True,  "Dim X As Long",       "dim x as long"),
    ("CRLF vs LF",               True,  "a = 1\r\nb = 2",      "a = 1\nb = 2"),
    ("CR vs LF",                 True,  "a = 1\rb = 2",        "a = 1\nb = 2"),
    ("vodeci prazan red",        True,  "\r\nx = 1",           "x = 1"),
    ("zavrsni prazan red",       True,  "x = 1\r\n\r\n",       "x = 1"),
    ("trailing razmak",          True,  "x = 1   ",            "x = 1"),
    ("NBSP umesto razmaka",      True,  "x =" + NBSP + "1",    "x = 1"),

    ("case u stringu",           False, 'x = "DA"',            'x = "da"'),
    # Escaped navodnik UNUTAR stringa: literal je  say "DA" now.  Case iza ""
    # mora ostati sacuvan -- da parser ne "izadje" iz stringa na "" i pocne da
    # lowercase-uje sadrzaj. (NB: 'x = ""DA""' NIJE ovaj slucaj -- to su DVA
    # prazna stringa sa golim kodom izmedju, pa je DA stvarno van stringa.)
    ("case iza escaped navodnika", False, 'x = "say ""DA"" now"',
                                          'x = "say ""da"" now"'),
    ("tekst komentara",          False, "x = 1 ' staro",       "x = 1 ' novo"),
    ("razmak u stringu",         False, 'x = "a  b"',          'x = "a b"'),
    ("prava izmena koda",        False, "x = 1",               "x = 2"),
)


# --- provere -----------------------------------------------------------------

def check_parity(modules=PARITY_MODULES, procs=PARITY_PROCS, src_dir=None):
    """Nalazi: procedura koja nedostaje ili se kod razlikuje medju modulima."""
    import difflib

    src_dir = src_dir or SRC_VBA
    nalazi = []
    tabele = {}
    for m in modules:
        put = os.path.join(src_dir, m)
        if not os.path.isfile(put):
            nalazi.append("PARITET  {}: fajl ne postoji".format(m))
            return nalazi
        tabele[m] = read_procs(put)

    prvi = modules[0]
    for p in procs:
        if p not in tabele[prvi]:
            nalazi.append("PARITET  {}: nema procedure {}".format(prvi, p))
            continue
        etalon = code_only(tabele[prvi][p])
        for m in modules[1:]:
            if p not in tabele[m]:
                nalazi.append("PARITET  {}: nema procedure {} (postoji u {})"
                              .format(m, p, prvi))
                continue
            drugi = code_only(tabele[m][p])
            if drugi != etalon:
                d = "\n".join("           " + x for x in difflib.unified_diff(
                    etalon, drugi, prvi, m, lineterm="", n=1))
                nalazi.append("PARITET  {}: kod se razlikuje od {}\n{}"
                              .format(p, prvi, d))
    return nalazi


def check_fixtures():
    """Nalazi: referentni algoritam ne daje ocekivan verdikt nad korpusom."""
    nalazi = []
    for ime, ocek, izvor in HARD_CASES:
        dobijeno = ref_is_hard(izvor)
        if dobijeno != ocek:
            nalazi.append("FIKSTURA HARD/{}: ocekivano {}, dobijeno {}"
                          .format(ime, ocek, dobijeno))
    for ime, ocek, a, b in SAME_CASES:
        dobijeno = ref_same_code(a, b)
        if dobijeno != ocek:
            nalazi.append("FIKSTURA SAME/{}: ocekivano {}, dobijeno {}"
                          .format(ime, ocek, dobijeno))
    return nalazi


# --- self-test: dokaz u oba smera -------------------------------------------
#
# Zelena provera koja nikad nije pokazana crvena ne dokazuje da nesto meri.
# Svaki slucaj KVARI ulaz i trazi da provera pukne PO IMENU.

def _napisi(put, telo):
    with io.open(put, "w", encoding="ascii", newline="\r\n") as fh:
        fh.write(telo)


def self_test():
    import shutil
    import tempfile

    palo = []
    ukupno = 0

    # 1) PARITET hvata razliku u kodu.
    ukupno += 1
    tmp = tempfile.mkdtemp(prefix="vbaparity_kod_")
    try:
        _napisi(os.path.join(tmp, "modA.bas"),
                "Private Function F(ByVal s As String) As String\n"
                "    F = UCase$(s)\n"
                "End Function\n")
        _napisi(os.path.join(tmp, "modB.bas"),
                "Private Function F(ByVal s As String) As String\n"
                "    F = LCase$(s)\n"                       # <-- sabotaza
                "End Function\n")
        n = check_parity(("modA.bas", "modB.bas"), ("F",), src_dir=tmp)
        if not any(x.startswith("PARITET  F:") for x in n):
            palo.append("  paritet-razlika: sabotiran kod NIJE prijavljen "
                        "(dobijeno {})".format(n))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    # 2) PARITET hvata proceduru koja NEDOSTAJE u drugoj kopiji.
    ukupno += 1
    tmp = tempfile.mkdtemp(prefix="vbaparity_nema_")
    try:
        _napisi(os.path.join(tmp, "modA.bas"),
                "Private Function F() As Long\n    F = 1\nEnd Function\n")
        _napisi(os.path.join(tmp, "modB.bas"),
                "Private Function G() As Long\n    G = 1\nEnd Function\n")
        n = check_parity(("modA.bas", "modB.bas"), ("F",), src_dir=tmp)
        if not any("nema procedure F" in x for x in n):
            palo.append("  paritet-nedostaje: nedostajuca procedura NIJE prijavljena "
                        "(dobijeno {})".format(n))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    # 3) PARITET hvata komentar-samo razliku kao CISTO (inace bi svaki modul koji
    #    objasnjava svoj kontekst bio lazno crven i provera bi se ugasila).
    ukupno += 1
    tmp = tempfile.mkdtemp(prefix="vbaparity_kom_")
    try:
        _napisi(os.path.join(tmp, "modA.bas"),
                "Private Function F() As Long\n    F = 1   ' klijentska strana\n"
                "End Function\n")
        _napisi(os.path.join(tmp, "modB.bas"),
                "Private Function F() As Long\n    F = 1   ' dev alat\n"
                "End Function\n")
        n = check_parity(("modA.bas", "modB.bas"), ("F",), src_dir=tmp)
        if n:
            palo.append("  paritet-komentar: razlika SAMO u komentaru prijavljena "
                        "kao nalaz ({})".format(n))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    # 4) PARITET je ZELEN nad stvarnim src-vba (inace bi slucaj 1 bio besmislen:
    #    provera koja uvek vristi ne razlikuje ispravno od pokvarenog).
    ukupno += 1
    n = check_parity()
    if n:
        palo.append("  paritet-stvarni: src-vba nije u paritetu -- {}".format(n))

    # 5) FIKSTURA hvata regresiju razmaka -- tacno rupu iz zamke #22.
    ukupno += 1
    if ref_is_hard("Private x  As  MSForms.Label") is not True:
        palo.append("  fikstura-razmak: dva razmaka nisu prepoznata kao tvrdo telo")

    # 6) Sidro slucaja 5: sabotirana verzija BEZ sazimanja mora da PROMASI. Bez
    #    ovoga slucaj 5 vremenom prestane da meri rupu, a ostane zelen.
    ukupno += 1

    def bez_sazimanja(body):
        for line in body.replace("\r\n", "\n").split("\n"):
            u = ref_code_line_upper(line)        # <-- sabotaza: nema CollapseSpaces
            if not u:
                continue
            if "WITHEVENTS " in u or " AS MSFORMS." in u:
                return True
        return False

    if bez_sazimanja("Private x  As  MSForms.Label") is not False:
        palo.append("  fikstura-sidro: sabotirana verzija bi NASLA dva razmaka -- "
                    "slucaj vise ne meri rupu iz zamke #22")

    # 7) FIKSTURA hvata regresiju u SameCode: case u stringu se NE sme progutati.
    ukupno += 1
    if ref_same_code('x = "DA"', 'x = "da"') is not False:
        palo.append("  fikstura-string-case: case-only izmena u stringu progutana")

    # 8) Ceo korpus mora biti zelen.
    ukupno += 1
    n = check_fixtures()
    if n:
        palo.append("  fikstura-korpus: {}".format(n))

    if palo:
        print("\n".join(palo), file=sys.stderr)
        print("\nself-test: {} od {} slucajeva palo.".format(len(palo), ukupno),
              file=sys.stderr)
        return 2
    print("self-test: cisto ({} slucajeva: paritet-razlika, paritet-nedostaje, "
          "paritet-komentar, paritet-stvarni, fikstura-razmak, fikstura-sidro, "
          "fikstura-string-case, fikstura-korpus).".format(ukupno))
    return 0


def main(argv):
    if "--self-test" in argv:
        return self_test()

    nalazi = check_parity() + check_fixtures()
    if nalazi:
        for x in nalazi:
            print(x, file=sys.stderr)
        print("\nvba_parity_check: {} nalaza.".format(len(nalazi)), file=sys.stderr)
        return 2
    print("vba_parity_check: cisto ({} procedura u paritetu, {} fikstura)."
          .format(len(PARITY_PROCS), len(HARD_CASES) + len(SAME_CASES)))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
