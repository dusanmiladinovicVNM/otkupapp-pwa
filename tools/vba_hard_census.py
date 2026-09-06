"""Census tvrdih (HARD) VBA komponenti + kapija protiv nove kontaminacije.

TVRDA komponenta = ima MODULE-LEVEL `WithEvents` ili `As MSForms.` deklaraciju
(pre prve procedure). Takva komponenta ne sme kroz `AddFromString` nego ide kroz
`Remove`+`Import` u FAZI 2 self-update-a -- najrizicniji deo celog toka: projekat
je izmedju faza nepotpun, a oporavak zavisi od `Application.OnTime` i durable
registry stanja.

Zato tvrda povrsina mora da bude MALA I NAMERNA. Do sada je jedina evidencija bio
komentar u `modSelfUpdate`, koji je nabrajao module koji odavno nisu tacni.

    python tools/vba_hard_census.py              # census + kapija nad src-vba/
    python tools/vba_hard_census.py --list       # samo ispis, bez kapije (exit 0)
    python tools/vba_hard_census.py --self-test  # dokaz da kapija hvata kvar

Izlazni kod: 0 = cisto, 2 = ima nalaza.

Pravila kapije:
  TVRDA_FORMA    -- `.frm` NIKAD ne sme biti tvrd. `AddFromString` takvog tela
                    diskonektuje CodeModule (zamka #3), a forma se ne sme
                    Remove+Import-ovati u runtime-u (zamka #1) -- jedini ishod je
                    "potreban reinstall", tj. update koji NIKAD ne moze da prodje.
  TVRDA_DOCCLS   -- isto za document module (sheet-ovi): ne mogu se Remove-ovati.
  TVRDA_BAS      -- standardni modul ne sme biti tvrd, bez izuzetka. Kontroler
                    drzi reference kontrola kao `Object`, a evente rutira kroz
                    sink klasu. Mala event-adapter klasa sme da bude tvrda;
                    veliki kontroler ne treba da bude.
  TVRDA_CLS      -- nova tvrda klasa mora biti u WHITELIST (svesna odluka).
  MRTAV_UNOS     -- klasa iz WHITELIST-a koja vise NIJE tvrda mora da izadje iz
                    liste. Inace whitelist truli u spisak imena bez znacenja i
                    prvo sledece stvarno zaprljanje prodje neopazeno.

Algoritam detekcije se NE duplira: uvozi se referentna implementacija iz
`vba_parity_check`, koja je pod paritetnom kapijom sa obe VBA kopije. Ovde bi
treca kopija bila treca stvar koja moze da divergira.
"""

import io
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from vba_parity_check import (           # noqa: E402
    HARD_CASES,
    ref_code_line_upper,
    ref_collapse_spaces,
    ref_is_hard,
)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_VBA = os.path.join(ROOT, "src-vba")

EKSTENZIJE = (".bas", ".cls", ".frm", ".doccls")

# Namerne event-sink klase. Lista je izvedena iz stvarnog stanja `main`-a u
# trenutku uvodjenja kapije (posle uklanjanja tockica misa i SOFT-ifikacije
# velikih kontrolera): tvrda povrsina je tada pala sa 10 na 5 komponenti.
# NIJE vecna -- MRTAV_UNOS pravilo tera da se odrzava.
WHITELIST = (
    "clsAdminBtn",
    "clsBlokUI",
    "clsConfigBtn",
    "clsFlatBtn",
    "clsUiSink",
)


def extract_module_code(path):
    """Kod bez izvoznog header-a: VERSION, Begin..End dizajn blok, Attribute redovi.

    Isti posao kao `modSelfUpdate.ExtractModuleCode`. Bez ovoga bi `.frm`
    dizajnerski blok (`Begin {..} frmX ... End`) usao u analizu i svaka forma sa
    ListBox-om bi izgledala kao tvrda.
    """
    txt = io.open(path, encoding="latin-1", newline="").read()
    body = []
    depth = 0
    in_header = True
    for ls in txt.replace("\r\n", "\n").split("\n"):
        u = ls.strip().upper()
        if in_header:
            if u.startswith("VERSION "):
                continue
            if u.startswith("BEGIN"):
                depth += 1
                continue
            if depth > 0:
                if u == "END":
                    depth -= 1
                continue
            if u.startswith("ATTRIBUTE "):
                continue
            in_header = False
        if u.startswith("ATTRIBUTE "):
            continue
        body.append(ls)
    return "\n".join(body)


def hard_reason(body):
    """Prvi MODULE-LEVEL red koji cini telo tvrdim, ili None.

    Mora davati isti verdikt kao `ref_is_hard` -- self-test to tvrdi nad celim
    korpusom fikstura, da se objasnjenje i odluka ne raziđu.
    """
    for line in body.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        u = ref_collapse_spaces(ref_code_line_upper(line))
        if not u:
            continue
        w = u
        for pref in ("PUBLIC ", "PRIVATE ", "FRIEND ", "STATIC "):
            if w.startswith(pref):
                w = w[len(pref):]
        if (w.startswith("SUB ") or w.startswith("FUNCTION ")
                or w.startswith("PROPERTY ")):
            return None                   # prva procedura = kraj module-level dela
        if "WITHEVENTS " in u:
            return ("module-level WithEvents", line.strip())
        if " AS MSFORMS." in u:
            return ("module-level As MSForms.", line.strip())
    return None


def census(src_dir=None):
    """{ime fajla: (ext, razlog, red)} za svaku tvrdu komponentu + spisak svih."""
    src_dir = src_dir or SRC_VBA
    tvrdi = {}
    svi = []
    for fn in sorted(os.listdir(src_dir)):
        ext = os.path.splitext(fn)[1].lower()
        if ext not in EKSTENZIJE:
            continue
        svi.append((fn, ext))
        r = hard_reason(extract_module_code(os.path.join(src_dir, fn)))
        if r:
            tvrdi[fn] = (ext, r[0], r[1])
    return tvrdi, svi


def ispis(tvrdi, svi):
    red = []
    red.append("TVRDE VBA KOMPONENTE")
    red.append("")
    if not tvrdi:
        red.append("  (nema)")
    for fn in sorted(tvrdi):
        ext, razlog, linija = tvrdi[fn]
        red.append("  %s" % fn)
        red.append("      razlog: %s" % razlog)
        red.append("      red:    %s" % linija)
    red.append("")
    for ext, naziv in ((".bas", "standardni moduli"), (".cls", "klase"),
                       (".frm", "forme"), (".doccls", "document moduli")):
        uk = sum(1 for _, e in svi if e == ext)
        tv = sum(1 for fn, (e, _, _) in tvrdi.items() if e == ext)
        red.append("  %-18s ukupno %3d, tvrdih %d" % (naziv, uk, tv))
    return "\n".join(red)


def check_census(tvrdi, whitelist=WHITELIST):
    nalazi = []
    wl = set(whitelist)
    videni_wl = set()

    for fn in sorted(tvrdi):
        ext, razlog, linija = tvrdi[fn]
        base = os.path.splitext(fn)[0]
        if ext == ".frm":
            nalazi.append(
                "TVRDA_FORMA   {}: {} ({}). Forma se ne sme ni AddFromString-ovati "
                "(zamka #3) ni Remove+Import-ovati u runtime-u (zamka #1) -- update "
                "bi zavrsavao kao 'potreban reinstall', dakle nikad ne bi prosao."
                .format(fn, razlog, linija))
        elif ext == ".doccls":
            nalazi.append(
                "TVRDA_DOCCLS  {}: {} ({}). Document modul se ne moze Remove-ovati, "
                "pa za njega ne postoji faza 2.".format(fn, razlog, linija))
        elif ext == ".bas":
            nalazi.append(
                "TVRDA_BAS     {}: {} ({}). Standardni modul ne sme biti tvrd -- drzi "
                "kontrole kao Object, a evente rutiraj kroz sink klasu."
                .format(fn, razlog, linija))
        elif ext == ".cls":
            if base in wl:
                videni_wl.add(base)
            else:
                nalazi.append(
                    "TVRDA_CLS     {}: {} ({}). Nova tvrda klasa mora biti svesna "
                    "odluka -- ako je stvarno event-sink adapter, dodaj je u WHITELIST "
                    "u tools/vba_hard_census.py; ako nije, drzi kontrolu kao Object."
                    .format(fn, razlog, linija))

    for base in sorted(wl - videni_wl):
        nalazi.append(
            "MRTAV_UNOS    {}: u WHITELIST-u a nije (vise) tvrd. Izbaci ga -- "
            "whitelist koji truli je spisak imena bez znacenja, i prvo sledece "
            "stvarno zaprljanje prolazi neopazeno.".format(base))
    return nalazi


# --- self-test: dokaz u oba smera -------------------------------------------

def _napisi(put, telo):
    with io.open(put, "w", encoding="ascii", newline="\r\n") as fh:
        fh.write(telo)


TELO_TVRDO = ('Attribute VB_Name = "X"\nOption Explicit\n'
              "Private mBtn As MSForms.CommandButton\n"
              "Public Sub Go()\nEnd Sub\n")
TELO_MEKO = ('Attribute VB_Name = "X"\nOption Explicit\n'
             "Private mBtn As Object\n"
             "Public Sub Go()\n    Dim l As MSForms.Label\nEnd Sub\n")


def self_test():
    import shutil
    import tempfile

    palo = []
    slucajevi = []

    def slucaj(naziv, fajlovi, whitelist, ocekuj_kod):
        slucajevi.append(naziv)
        tmp = tempfile.mkdtemp(prefix="hardcensus_")
        try:
            for fn, telo in fajlovi.items():
                _napisi(os.path.join(tmp, fn), telo)
            tvrdi, _ = census(tmp)
            n = check_census(tvrdi, whitelist)
            if ocekuj_kod is None:
                if n:
                    palo.append("  {}: ocekivano CISTO, dobijeno {}".format(naziv, n))
            elif not any(x.startswith(ocekuj_kod) for x in n):
                palo.append("  {}: ocekivan nalaz {}, dobijeno {}"
                            .format(naziv, ocekuj_kod, n))
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    # Sidro: cist skup mora biti ZELEN, inace su svi ostali slucajevi prazan hod.
    slucaj("cist-skup",
           {"modA.bas": TELO_MEKO, "clsSink.cls": TELO_TVRDO},
           ("clsSink",), None)

    slucaj("tvrda-forma", {"frmX.frm": TELO_TVRDO}, (), "TVRDA_FORMA")
    slucaj("tvrd-doccls", {"sX.doccls": TELO_TVRDO}, (), "TVRDA_DOCCLS")
    slucaj("tvrd-bas", {"modX.bas": TELO_TVRDO}, (), "TVRDA_BAS")
    slucaj("nova-tvrda-klasa", {"clsNova.cls": TELO_TVRDO}, (), "TVRDA_CLS")
    slucaj("mrtav-unos", {"clsX.cls": TELO_MEKO}, ("clsX",), "MRTAV_UNOS")

    # Dizajnerski blok forme NE sme da napravi laznu tvrdu formu: `.frm` header
    # nosi "Begin {..} ... ListBox ..." koji bez strip-a izgleda kao deklaracija.
    slucajevi.append("dizajner-nije-telo")
    frm = ('VERSION 5.00\n'
           'Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmX \n'
           '   Begin Forms.ListBox lst \n'
           '   End\n'
           'End\n'
           'Attribute VB_Name = "frmX"\n'
           "Option Explicit\n"
           "Public Sub Go()\nEnd Sub\n")
    tmp = tempfile.mkdtemp(prefix="hardcensus_frm_")
    try:
        _napisi(os.path.join(tmp, "frmX.frm"), frm)
        tvrdi, _ = census(tmp)
        if tvrdi:
            palo.append("  dizajner-nije-telo: dizajnerski blok procitan kao telo "
                        "-> lazna tvrda forma ({})".format(tvrdi))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    # hard_reason i ref_is_hard MORAJU da se slazu nad celim korpusom fikstura.
    # Inace bi census objasnjavao jedno a paritetna kapija merila drugo.
    slucajevi.append("saglasnost-sa-fiksturama")
    neslaganja = []
    for ime, ocek, izvor in HARD_CASES:
        if (hard_reason(izvor) is not None) != ref_is_hard(izvor):
            neslaganja.append(ime)
        if ref_is_hard(izvor) != ocek:
            neslaganja.append(ime + "(fikstura)")
    if neslaganja:
        palo.append("  saglasnost-sa-fiksturama: hard_reason i ref_is_hard se ne "
                    "slazu na {}".format(neslaganja))

    # Stvarni src-vba mora biti cist.
    slucajevi.append("stvarni-src-vba")
    tvrdi, _ = census()
    n = check_census(tvrdi)
    if n:
        palo.append("  stvarni-src-vba: {}".format(n))

    if palo:
        print("\n".join(palo), file=sys.stderr)
        print("\nself-test: {} od {} slucajeva palo."
              .format(len(palo), len(slucajevi)), file=sys.stderr)
        return 2
    print("self-test: cisto ({} slucajeva: {})."
          .format(len(slucajevi), ", ".join(slucajevi)))
    return 0


def main(argv):
    if "--self-test" in argv:
        return self_test()

    tvrdi, svi = census()
    print(ispis(tvrdi, svi))

    if "--list" in argv:
        return 0

    nalazi = check_census(tvrdi)
    if nalazi:
        print("", file=sys.stderr)
        for x in nalazi:
            print(x, file=sys.stderr)
        print("\nvba_hard_census: {} nalaza.".format(len(nalazi)), file=sys.stderr)
        return 2
    print("")
    print("vba_hard_census: cisto ({} tvrdih, sve u WHITELIST-u)."
          .format(len(tvrdi)))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
