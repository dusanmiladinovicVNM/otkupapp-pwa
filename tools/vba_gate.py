#!/usr/bin/env python3
"""Identitet onoga sto je dokazano: popis suita + otisak src-vba + marker zelenog.

Radi svuda -- nema Excela, nema COM-a. Odgovara na dva pitanja koja `run_vba.py`
ne moze da postavi sam sebi:

    1. POSTOJI LI SUITE KOJU NIJEDNA KAPIJA NE POKRECE?
    2. DA LI JE BAS OVAJ IZVOR PROSAO TESTOVE, ili neki drugi?

    python tools/vba_gate.py --popis              # suite u kodu vs katalog SUITES
    python tools/vba_gate.py --hash               # kanonski SHA256 src-vba
    python tools/vba_gate.py --status             # izvor vs marker, po suite-u
    python tools/vba_gate.py --require-green      # exit 0 samo ako je OVAJ izvor dokazan
    python tools/vba_gate.py --require-green --suite RunStornoTestSuite
    python tools/vba_gate.py --require-green --require-compile
    python tools/vba_gate.py --mark-compile       # operater potvrdio Debug > Compile
    python tools/vba_gate.py --clear
    python tools/vba_gate.py --self-test

Izlazni kod: 0 = ok, 2 = nalaz.

--- 1. POPIS SUITA --------------------------------------------------------

Suite se napise, prodje kad je neko rucno pozove iz Immediate prozora, i nikad se
ne prikljuci kapiji. Niko to ne primeti: katalog `SUITES` u run_vba.py ne zna da
procedura postoji, pa je ne pominje ni kao preskocenu. Zaglavlja takvih suita to
i kazu naglas -- "Pozivaj sa: ?RunHttpUtilsSmokeSuite / Ocekivano: PASS=18" --
dakle i verdikt i ocekivan broj zive u komentaru i u necijem pamcenju.

`SUITE_VAN_KAPIJA` ispod je zato REGISTAR, ne izuzetak:
  - prazan ili prekratak razlog je nalaz;
  - zastareo unos je nalaz (suite obrisana, ili je u medjuvremenu u `SUITES`);
  - razlog mora da kaze STA BI JE ZATVORILO, ne samo da stoji.

Isti oblik i isti razlog kao `MRTAV_UNOS` u `vba_hard_census`: lista koja se ne
odrzava postaje spisak imena bez znacenja, pa prvo sledece stvarno zaprljanje
prodje neopazeno.

--- 2. OTISAK I MARKER ----------------------------------------------------

"Suite su bile zelene" je tvrdnja o NEKOM izvoru. Koji -- do sada se znalo samo
iz recenice uz PR. Posle rebase-a, amend-a ili jedne usputne izmene ta recenica
i dalje stoji a vise ne vazi: prolaz je bio nad src-vba koji vise ne postoji.

Zato se izvor hesira, a marker pamti otisak i REZULTAT PO SUITE-U --
`--require-green` time ume da razlikuje "dokazano" od "dokazano nesto drugo", i
da imenuje koja suite nedostaje.

STO MARKER NE SME DA UPISE:
  - run koji je pao (rc != 0);
  - `--no-import` run: kod u svesci tada NIJE src-vba, pa bi otisak lagao o tome
    sta je izvrseno;
  - BLIND suite kao dokazanu. `gate: False` znaci "prosla bez greske", a to NIJE
    isto sto i "sve provere prosle" (v. `SUITES` u run_vba.py). Takva se pamti
    kao `BLIND` i `--require-green` je ne priznaje.

COMPILE je rucna kapija operatera (`Alt+F11 -> Debug -> Compile VBAProject`), pa
ima svoj upis: `--mark-compile` vezuje tu potvrdu za OTISAK. Potvrda data nad
jednim izvorom time prestaje da vazi za sledeci -- a bas to se vec desilo, kao P3
u review-u PR #400 ("compile evidence stale").

ZASTO NORMALIZACIJA PRELOMA
    `.gitattributes` danas drzi `eol=crlf` nad celim `src-vba/`, pa bi i sirov
    hash bajtova bio stabilan -- ali samo dok ta linija stoji i dok se ne pojavi
    fajl koji joj umakne (nov nastavak, worktree bez atributa, sandbox). Hash
    koji na dve masine znaci dve stvari je gori od hasha koji se racuna jednom
    vise, pa se tekst normalizuje na LF; `.frx` je binaran i ide sirov.

    `modBuildInfo.bas` je IZUZET: `tools/stamp-build` ga prepisuje PRE svakog
    `ImportAllVBA` (BUILD_SHA / BUILD_VERSION / BUILD_DATE) i ta izmena se ne
    commit-uje. Da je u hesu, stamp bi obarao marker bas u trenutku release-a --
    a modul ne nosi nikakvo ponasanje, tri konstante koje niko ne grana.

    OVAJ OTISAK NIJE `dokaz.py._otisak`, i ne treba da bude. Tamo se poredi
    stanje PRE i POSLE jedne mutacije u istom procesu, pa je sirov bajt po bajtu
    tacno ono sto se trazi -- ukljucujuci `modBuildInfo`. Ovde se poredi izvor
    izmedju dva procesa i dve masine. Dve namene, dve funkcije; spajanje bi
    jednoj oduzelo ono zbog cega postoji.
"""
from __future__ import annotations

import argparse
import hashlib
import importlib.util
import io
import json
import os
import platform
import re
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_VBA = os.path.join(ROOT, "src-vba")
MARKER = os.path.join(ROOT, "tests", "last_green.json")

# Modul koji stamp-build prepisuje pred svaki import -- v. docstring.
IZUZET_IZ_OTISKA = ("modBuildInfo.bas",)
TEKST_NASTAVCI = (".bas", ".cls", ".frm", ".doccls")

# Ulazna tacka suite-a po konvenciji imena. `RunProductionHealthCheck` joj ne
# odgovara, i ne mora: sve iz `SUITES` se proverava po IMENU (da procedura
# postoji), a konvencija sluzi samo da nadje ono sto u katalogu NIJE.
KONVENCIJA = re.compile(r"^Public Sub (Run\w*(?:Suite|Tests)|Test\w*_All)\(\s*\)\s*$",
                        re.M)
_JAVNA_PROC = re.compile(r"^Public Sub (\w+)\(([^)]*)\)", re.M)

MIN_RAZLOG = 40

# Suite koje nijedna kapija ne pokrece -- REGISTAR, ne izuzetak (v. docstring).
#
# Sve cetiri su nadjene `--popis`-om pri prvom pokretanju; do tada nisu bile
# zapisane nigde. Zajednicko im je da im verdikt ide u Immediate prozor: nijedna
# nema `Err.Raise` u svom telu, pa bi i prikljucena bila BLIND -- to je posao
# koji svaki unos zatvara, i zato stoji u razlogu.
SUITE_VAN_KAPIJA = {
    "RunHttpUtilsSmokeSuite": (
        "HTTP util helperi, bez mreze. Zaglavlje kaze 'Pozivaj sa: ?RunHttp... / "
        "Ocekivano: PASS=18' -- dakle i verdikt i broj zive u komentaru. Zatvara "
        "se tako sto telo dobije Err.Raise na kraju (inace bi u SUITES bila BLIND), "
        "pa unos u SUITES sa gate: True."),
    "RunSEFDocumentIdShapeSuite": (
        "Oblik SEF document ID-a, cist string test. Zaglavlje kaze 'Ocekivano: "
        "PASS=14'. Isti harness kao RunHttpUtilsSmokeSuite i isti posao: Err.Raise "
        "u telu, pa unos u SUITES sa gate: True."),
    "RunSEFStateTransitionSuite": (
        "Prelazi stanja SEF automata. Deo istih Test_* provera vec vrti "
        "RunSEFTestSuite (koja JE u SUITES i podize gresku), pa je prvo pitanje da "
        "li je ova zasebna suite uopste potrebna ili joj tvrdnje idu u onu."),
    "RunSEFClientParserSmokeSuite": (
        "Parser SEF odgovora, bez mreze, ali ZIVI U PRODUKCIONOM MODULU "
        "(modSEFClient.bas) i pad broji u lokalnu promenljivu. Zatvara se "
        "premestanjem u modSEFTests uz Err.Raise, ili brisanjem ako su te tvrdnje "
        "pokrivene drugde."),
}


# --- 1. popis suita --------------------------------------------------------


def _ucitaj_run_vba():
    put = os.path.join(ROOT, "tools", "run_vba.py")
    spec = importlib.util.spec_from_file_location("_run_vba_za_gate", put)
    modul = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modul)
    return modul


def katalog_suita() -> dict:
    """`SUITES` iz run_vba.py -- jedini izvor istine o tome koja suite postoji.

    Cita se, ne prepisuje: druga kopija kataloga bi bila druga stvar koja moze da
    se razidje (isti razlog zbog koga vba_hard_census uvozi referencu iz
    vba_parity_check umesto da je ponovi).
    """
    return _ucitaj_run_vba().SUITES


def skeniraj(src_dir: str = SRC_VBA) -> tuple:
    """(javne, po_konvenciji) iz JEDNOG prolaza kroz src-vba.

    Jedan prolaz, ne dva: provera ide kroz `vba_check`, dakle kroz PostToolUse
    hook, pa se citanje 190 fajlova ne placa dvaput.

    `javne` su sve `Public Sub` (ime -> fajl) i sluze da se proveri POSTOJANJE
    imena iz kataloga. `po_konvenciji` su kandidati za ulaznu tacku suite-a, i
    traze se samo u `.bas`: makro se po imenu moze pozvati jedino iz standardnog
    modula, pa `Public Sub RunFooSuite()` u klasi ili formi nije suite koju bi
    kapija mogla da pokrene.
    """
    javne, po_konvenciji = {}, {}
    if not os.path.isdir(src_dir):
        return javne, po_konvenciji
    for ime in sorted(os.listdir(src_dir)):
        if not ime.endswith(TEKST_NASTAVCI):
            continue
        with io.open(os.path.join(src_dir, ime), encoding="ascii",
                     errors="replace", newline="") as fh:
            tekst = fh.read().replace("\r\n", "\n")
        for m in _JAVNA_PROC.finditer(tekst):
            javne.setdefault(m.group(1), (ime, not m.group(2).strip()))
        if ime.endswith(".bas"):
            for m in KONVENCIJA.finditer(tekst):
                po_konvenciji.setdefault(m.group(1), ime)
    return javne, po_konvenciji


def popis_problemi(suites: dict = None, registar: dict = None,
                   src_dir: str = SRC_VBA, skenirano: tuple = None) -> list:
    """Nalazi iz popisa suita. Prazna lista = cisto."""
    suites = katalog_suita() if suites is None else suites
    registar = SUITE_VAN_KAPIJA if registar is None else registar
    javne, po_konvenciji = skeniraj(src_dir) if skenirano is None else skenirano

    nalazi = []

    # a) katalog imenuje proceduru koje nema -- run bi pukao na Run()
    for ime in sorted(suites):
        if ime not in javne:
            nalazi.append("FANTOM: SUITES['%s'] nema `Public Sub %s` u src-vba"
                          % (ime, ime))

    # b) suite po konvenciji koju nijedna kapija ne pokrece i koja nije zapisana
    for ime in sorted(set(po_konvenciji) - set(suites) - set(registar)):
        nalazi.append("NEPOKRETANA: `%s` (%s) nije ni u SUITES ni u "
                      "SUITE_VAN_KAPIJA -- napisana suite koju nijedna kapija ne "
                      "pokrece" % (ime, po_konvenciji[ime]))

    # c) registar koji je zastareo -- isti oblik kao MRTAV_UNOS u hard_census
    for ime in sorted(registar):
        if ime in suites:
            nalazi.append("MRTAV UNOS: `%s` je u medjuvremenu u SUITES -- obrisi "
                          "ga iz SUITE_VAN_KAPIJA" % ime)
        elif ime not in javne:
            nalazi.append("MRTAV UNOS: `%s` ne postoji u src-vba -- obrisi ga iz "
                          "SUITE_VAN_KAPIJA" % ime)
        elif len((registar[ime] or "").strip()) < MIN_RAZLOG:
            nalazi.append("BEZ RAZLOGA: SUITE_VAN_KAPIJA['%s'] ne kaze zasto je "
                          "van kapije ni sta bi je zatvorilo" % ime)
    return nalazi


# --- 2. otisak i marker ----------------------------------------------------


def otisak_izvora(src_dir: str = SRC_VBA) -> str:
    """Kanonski SHA256 celog src-vba: ime fajla + sadrzaj. V. docstring."""
    h = hashlib.sha256()
    if not os.path.isdir(src_dir):
        return ""
    for ime in sorted(os.listdir(src_dir)):
        if ime in IZUZET_IZ_OTISKA:
            continue
        h.update(ime.encode())
        with open(os.path.join(src_dir, ime), "rb") as fh:
            raw = fh.read()
        if ime.endswith(TEKST_NASTAVCI):
            raw = raw.replace(b"\r\n", b"\n")
        h.update(raw)
    return h.hexdigest()


def procitaj_marker(put: str = MARKER) -> dict:
    try:
        with io.open(put, encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return None


def upisi_marker(podaci: dict, put: str = MARKER) -> None:
    os.makedirs(os.path.dirname(put), exist_ok=True)
    with io.open(put, "w", encoding="utf-8") as fh:
        json.dump(podaci, fh, ensure_ascii=False, indent=2, sort_keys=True)
        fh.write("\n")


def _git_glava() -> str:
    try:
        r = subprocess.run(["git", "-C", ROOT, "rev-parse", "--short", "HEAD"],
                           capture_output=True, text=True, timeout=20)
        return r.stdout.strip() if r.returncode == 0 else ""
    except (OSError, subprocess.SubprocessError):
        return ""


def potrebne_suite(suites: dict = None) -> list:
    """Suite koje `--require-green` trazi kad mu se ne kaze drugacije.

    Podrazumevan set iz `SUITES`, ali SAMO one sa `gate: True`: blind suite se ne
    moze dokazati, pa je traziti znacilo bi traziti nesto sto marker po pravilu
    nikad nece sadrzati kao OK.
    """
    suites = katalog_suita() if suites is None else suites
    return sorted(k for k, v in suites.items()
                  if v.get("default") and v.get("gate"))


def zabelezi_prolaz(report: dict, rc: int, no_import: bool = False,
                    put: str = MARKER, src_dir: str = SRC_VBA) -> str:
    """Zapisi rezultat run-a u marker. Vraca poruku (sta je upisano ili zasto nije).

    Zove je `run_vba.py` na kraju run-a. Pravila su u docstring-u modula; ovde su
    kao kod, jer su upravo ona ono sto marker cini tvrdnjom a ne dekoracijom.
    """
    if no_import:
        return ("marker nije upisan: --no-import znaci da kod u svesci nije "
                "src-vba, pa otisak ne bi opisivao ono sto je izvrseno")
    if rc != 0:
        return "marker nije upisan: run nije zelen (rc=%s)" % rc

    otisak = otisak_izvora(src_dir)
    stari = procitaj_marker(put)
    if stari and stari.get("otisak") == otisak:
        podaci = stari                       # dopuni: delimicni run-ovi se slazu
    else:
        # Rezultati nad drugim izvorom nisu rezultati nad ovim. Brisu se svi,
        # ukljucujuci potvrdu compile-a.
        podaci = {"otisak": otisak, "suites": {}}

    podaci["git"] = _git_glava()
    podaci["platforma"] = platform.platform()
    podaci["kada"] = time.strftime("%Y-%m-%dT%H:%M:%S")
    rezultati = report.get("suite_results", {}) or {}
    for s in report.get("suites", []):
        t = rezultati.get(s["name"]) or {}
        podaci["suites"][s["name"]] = {
            "status": s.get("status"),
            "ukupno": t.get("total"),
            "palo": t.get("failed"),
            "kada": podaci["kada"],
        }
    upisi_marker(podaci, put)
    imena = sorted(podaci["suites"])
    return "marker upisan nad %s: %d suita (%s)" % (
        otisak[:12], len(imena), ", ".join(imena) or "nijedna")


def zabelezi_compile(put: str = MARKER, src_dir: str = SRC_VBA) -> str:
    """Operater je potvrdio `Debug > Compile` nad OVIM izvorom."""
    otisak = otisak_izvora(src_dir)
    podaci = procitaj_marker(put)
    if not podaci or podaci.get("otisak") != otisak:
        podaci = {"otisak": otisak, "suites": {}}
    podaci["compile"] = {"kada": time.strftime("%Y-%m-%dT%H:%M:%S"),
                         "git": _git_glava()}
    upisi_marker(podaci, put)
    return "compile potvrdjen nad %s" % otisak[:12]


def zahtevaj_zeleno(trazene: list = None, suites: dict = None,
                    trazi_compile: bool = False, put: str = MARKER,
                    src_dir: str = SRC_VBA) -> list:
    """Nalazi zbog kojih OVAJ izvor nije dokazan. Prazna lista = dokazan."""
    suites = katalog_suita() if suites is None else suites
    trazene = potrebne_suite(suites) if trazene is None else list(trazene)
    otisak = otisak_izvora(src_dir)
    marker = procitaj_marker(put)

    if not marker:
        return ["nema markera: nijedan prolaz nije zapisan (pusti "
                "`python tools/run_vba.py`)"]
    if marker.get("otisak") != otisak:
        return ["marker je nad DRUGIM izvorom: marker %s, izvor %s -- prolaz je "
                "bio nad kodom koji vise ne stoji"
                % ((marker.get("otisak") or "?")[:12], otisak[:12])]

    nalazi = []
    zapisane = marker.get("suites") or {}
    for ime in trazene:
        z = zapisane.get(ime)
        if not z:
            nalazi.append("%s: nije u markeru -- nije pustena nad ovim izvorom"
                          % ime)
        elif z.get("status") != "OK":
            nalazi.append("%s: zapisana kao %s, a to nije dokaz da su sve provere "
                          "prosle" % (ime, z.get("status")))
    if trazi_compile and not marker.get("compile"):
        nalazi.append("compile nije potvrdjen nad ovim izvorom "
                      "(`--mark-compile` posle Debug > Compile VBAProject)")
    return nalazi


def stanje_redovi(suites: dict = None, put: str = MARKER,
                  src_dir: str = SRC_VBA) -> list:
    suites = katalog_suita() if suites is None else suites
    otisak = otisak_izvora(src_dir)
    marker = procitaj_marker(put)
    redovi = ["izvor:  %s" % otisak[:16]]
    if not marker:
        redovi.append("marker: nema ga -- nijedan prolaz nije zapisan")
        return redovi
    redovi.append("marker: %s  %s  (git %s, %s)"
                  % ((marker.get("otisak") or "?")[:16],
                     "ISTI IZVOR" if marker.get("otisak") == otisak
                     else "DRUGI IZVOR",
                     marker.get("git") or "?", marker.get("kada") or "?"))
    c = marker.get("compile")
    redovi.append("compile: %s" % ("potvrdjen %s" % c.get("kada") if c
                                   else "nije potvrdjen"))
    zapisane = marker.get("suites") or {}
    trazene = set(potrebne_suite(suites))
    for ime in sorted(set(zapisane) | trazene):
        z = zapisane.get(ime) or {}
        oznaka = "*" if ime in trazene else " "
        if not z:
            redovi.append(" %s %-28s --" % (oznaka, ime))
        else:
            redovi.append(" %s %-28s %-6s %s/%s"
                          % (oznaka, ime, z.get("status"),
                             z.get("palo"), z.get("ukupno")))
    redovi.append("(* = trazi je --require-green)")
    return redovi


# --- self-test: dokaz u oba smera ------------------------------------------
#
# CLAUDE.md par. 5: kad se menja sam checker, trazi se dvosmerni dokaz. Nijedno
# pravilo ovde ne trazi Excel, pa sve ide kroz sinteticki `src-vba` u temp
# folderu i sinteticki marker -- nikad nad pravim repoom, da self-test ne moze
# da obrise ili prepise stvarni marker.


def _lazni_izvor(koren: str, fajlovi: dict) -> str:
    put = os.path.join(koren, "src-vba")
    os.makedirs(put, exist_ok=True)
    for ime, sadrzaj in fajlovi.items():
        with open(os.path.join(put, ime), "wb") as fh:
            fh.write(sadrzaj if isinstance(sadrzaj, bytes)
                     else sadrzaj.encode("ascii"))
    return put


def _self_test(tiho: bool = False) -> int:
    import shutil
    import tempfile

    nalazi = []

    def tvrdi(uslov, opis):
        if not uslov:
            nalazi.append(opis)

    SUITES = {"RunAllTests": {"gate": True, "default": True},
              "RunNovacSmokeSuite": {"gate": False, "default": True},
              "RunSEFTestSuite": {"gate": True, "default": False}}
    JAVNE = {"RunAllTests": ("modTest.bas", True),
             "RunNovacSmokeSuite": ("modNovacTests.bas", True),
             "RunSEFTestSuite": ("modSEFTests.bas", True)}

    # --- popis: svaki nalaz u oba smera ------------------------------------
    tmp = tempfile.mkdtemp(prefix="vbagate_")
    try:
        src = _lazni_izvor(tmp, {
            "modTest.bas": "Public Sub RunAllTests()\r\nEnd Sub\r\n",
            "modNovacTests.bas": "Public Sub RunNovacSmokeSuite()\r\nEnd Sub\r\n",
            "modSEFTests.bas": "Public Sub RunSEFTestSuite()\r\nEnd Sub\r\n",
        })
        RAZLOG = "x" * MIN_RAZLOG

        def popis(suites=SUITES, registar=None, src_dir=src):
            return popis_problemi(suites, registar or {}, src_dir)

        tvrdi(not popis(), "POPIS: cist izvor daje nalaz")

        tvrdi(any("FANTOM" in n for n in popis(
                  suites=dict(SUITES, RunNemaMe={"gate": True, "default": True}))),
              "POPIS: SUITES unos bez `Public Sub` nije FANTOM")

        with io.open(os.path.join(src, "modNov.bas"), "w", newline="") as fh:
            fh.write("Public Sub RunNovaSuite()\r\nEnd Sub\r\n")
        tvrdi(any("NEPOKRETANA" in n and "RunNovaSuite" in n for n in popis()),
              "POPIS: nova suite van kataloga nije NEPOKRETANA")
        tvrdi(not popis(registar={"RunNovaSuite": RAZLOG}),
              "POPIS: zapisana suite u registru i dalje daje nalaz")
        tvrdi(any("BEZ RAZLOGA" in n for n in popis(
                  registar={"RunNovaSuite": "kratko"})),
              "POPIS: prekratak razlog u registru nije nalaz")
        tvrdi(any("MRTAV UNOS" in n for n in popis(
                  registar={"RunAllTests": RAZLOG})),
              "POPIS: unos koji je u SUITES nije MRTAV UNOS")
        tvrdi(any("MRTAV UNOS" in n for n in popis(
                  registar={"RunNemaNigde": RAZLOG})),
              "POPIS: unos kog nema u src-vba nije MRTAV UNOS")
        os.remove(os.path.join(src, "modNov.bas"))

        # Procedura sa argumentima nije ulazna tacka suite-a.
        with io.open(os.path.join(src, "modP.bas"), "w", newline="") as fh:
            fh.write("Public Sub RunNestoSuite(ByVal x As Long)\r\nEnd Sub\r\n")
        tvrdi(not popis(), "POPIS: procedura SA argumentima se broji kao suite")
        os.remove(os.path.join(src, "modP.bas"))

        # Isto ime u KLASI nije suite: makro se po imenu zove samo iz .bas.
        with io.open(os.path.join(src, "clsX.cls"), "w", newline="") as fh:
            fh.write("Public Sub RunKlasaSuite()\r\nEnd Sub\r\n")
        tvrdi(not popis(), "POPIS: Public Sub u .cls se broji kao suite")
        os.remove(os.path.join(src, "clsX.cls"))

        # --- otisak -------------------------------------------------------
        a = otisak_izvora(src)
        tvrdi(len(a) == 64, "OTISAK: nije SHA256")
        with io.open(os.path.join(src, "modTest.bas"), "a", newline="") as fh:
            fh.write("' izmena\r\n")
        tvrdi(otisak_izvora(src) != a, "OTISAK: izmena u .bas ga ne menja")
        with io.open(os.path.join(src, "modTest.bas"), encoding="ascii",
                     newline="") as fh:
            crlf = fh.read()
        with open(os.path.join(src, "modTest.bas"), "wb") as fh:
            fh.write(crlf.replace("\r\n", "\n").encode())
        tvrdi(otisak_izvora(src) == otisak_izvora(src),
              "OTISAK: nije stabilan nad istim stanjem")
        lf = otisak_izvora(src)
        with open(os.path.join(src, "modTest.bas"), "wb") as fh:
            fh.write(crlf.encode())
        tvrdi(otisak_izvora(src) == lf,
              "OTISAK: CRLF i LF isti sadrzaj daju razlicit otisak")

        b = otisak_izvora(src)
        with io.open(os.path.join(src, "modBuildInfo.bas"), "w", newline="") as fh:
            fh.write('Public Const BUILD_SHA As String = "abc1234"\r\n')
        tvrdi(otisak_izvora(src) == b,
              "OTISAK: modBuildInfo ulazi u otisak -- stamp bi obarao marker")

        # --- marker -------------------------------------------------------
        put = os.path.join(tmp, "tests", "last_green.json")
        ZELEN = {"suites": [{"name": "RunAllTests", "status": "OK"}],
                 "suite_results": {"RunAllTests": {"total": 199, "failed": 0}}}

        def zahtevaj(**kw):
            # Pad je NALAZ, ne traceback: ugasena kapija `if not marker` inace
            # rusi self-test, pa se ne vidi kao sopstvena greska.
            kw.setdefault("trazene", ["RunAllTests"])
            try:
                return zahtevaj_zeleno(suites=SUITES, put=put, src_dir=src, **kw)
            except Exception as e:          # noqa: BLE001
                nalazi.append("MARKER: zahtevaj_zeleno je pukao -- %r" % (e,))
                return []

        tvrdi(any("nema markera" in n for n in zahtevaj()),
              "MARKER: bez markera je izvor 'dokazan'")

        poruka = zabelezi_prolaz(ZELEN, 2, put=put, src_dir=src)
        tvrdi("nije zelen" in poruka and not os.path.exists(put),
              "MARKER: PAO run upisuje marker")

        poruka = zabelezi_prolaz(ZELEN, 0, no_import=True, put=put, src_dir=src)
        tvrdi("no-import" in poruka and not os.path.exists(put),
              "MARKER: --no-import run upisuje marker")

        zabelezi_prolaz(ZELEN, 0, put=put, src_dir=src)
        tvrdi(not zahtevaj(), "MARKER: zelen prolaz nad istim izvorom nije dokaz")
        tvrdi(any("compile nije potvrdjen" in n
                  for n in zahtevaj(trazi_compile=True)),
              "MARKER: --require-compile prolazi bez potvrde")
        zabelezi_compile(put=put, src_dir=src)
        tvrdi(not zahtevaj(trazi_compile=True),
              "MARKER: potvrda compile-a se ne priznaje")

        tvrdi(any("nije u markeru" in n
                  for n in zahtevaj(trazene=["RunSEFTestSuite"])),
              "MARKER: suite koja nije pustena se racuna kao dokazana")

        BLIND = {"suites": [{"name": "RunNovacSmokeSuite", "status": "BLIND"}],
                 "suite_results": {}}
        zabelezi_prolaz(BLIND, 0, put=put, src_dir=src)
        tvrdi(any("BLIND" in n for n in zahtevaj(trazene=["RunNovacSmokeSuite"])),
              "MARKER: BLIND suite se priznaje kao dokaz")
        tvrdi(not zahtevaj(),
              "MARKER: drugi prolaz nad istim izvorom je obrisao prvi")

        # Izmena izvora obara i suite i compile -- bez toga marker lazhe.
        with io.open(os.path.join(src, "modTest.bas"), "a", newline="") as fh:
            fh.write("' jos jedna izmena\r\n")
        tvrdi(any("DRUGIM izvorom" in n for n in zahtevaj()),
              "MARKER: izmena izvora ne obara marker")
        zabelezi_prolaz(ZELEN, 0, put=put, src_dir=src)
        tvrdi(any("compile nije potvrdjen" in n
                  for n in zahtevaj(trazi_compile=True)),
              "MARKER: nov izvor nasledjuje staru potvrdu compile-a")
        tvrdi("RunNovacSmokeSuite" not in (procitaj_marker(put) or {}).get(
                  "suites", {}),
              "MARKER: nov izvor nasledjuje stare rezultate suita")

        # --- potrebne_suite -----------------------------------------------
        tvrdi(potrebne_suite(SUITES) == ["RunAllTests"],
              "POTREBNE: blind ili ne-default suite ulazi u trazeni set")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    if nalazi:
        print("vba_gate --self-test: %d nalaza" % len(nalazi), file=sys.stderr)
        for n in nalazi:
            print("  " + n, file=sys.stderr)
        return 1
    if not tiho:
        print("vba_gate --self-test: popis, otisak i marker grizu")
    return 0


def main(argv: list) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--popis", action="store_true",
                    help="suite u kodu vs katalog SUITES vs registar")
    ap.add_argument("--hash", action="store_true",
                    help="ispisi kanonski SHA256 src-vba")
    ap.add_argument("--status", action="store_true",
                    help="citljiv rezime: izvor vs marker, po suite-u")
    ap.add_argument("--require-green", action="store_true",
                    help="exit 0 samo ako je BAS OVAJ izvor dokazan")
    ap.add_argument("--require-compile", action="store_true",
                    help="uz --require-green: trazi i potvrdu Debug > Compile")
    ap.add_argument("--suite", action="append", default=[],
                    help="uz --require-green: trazi bas ovu suite (moze vise puta)")
    ap.add_argument("--mark-compile", action="store_true",
                    help="zapisi da je Debug > Compile prosao nad ovim izvorom")
    ap.add_argument("--clear", action="store_true", help="obrisi marker")
    ap.add_argument("--self-test", action="store_true",
                    help="dokazi da popis, otisak i marker zaista grizu")
    args = ap.parse_args(argv)

    if args.self_test:
        return _self_test()
    if args.hash:
        print(otisak_izvora())
        return 0
    if args.clear:
        if os.path.exists(MARKER):
            os.remove(MARKER)
            print("marker obrisan")
        else:
            print("markera nema")
        return 0
    if args.mark_compile:
        print(zabelezi_compile())
        return 0
    if args.status:
        for red in stanje_redovi():
            print(red)
        return 0
    if args.require_green:
        nalazi = zahtevaj_zeleno(args.suite or None,
                                 trazi_compile=args.require_compile)
        if not nalazi:
            print("izvor %s je dokazan" % otisak_izvora()[:12])
            return 0
        print("izvor NIJE dokazan:", file=sys.stderr)
        for n in nalazi:
            print("  " + n, file=sys.stderr)
        return 2
    if args.popis:
        nalazi = popis_problemi()
        if not nalazi:
            print("popis suita: ok (%d u SUITES, %d van kapija sa zapisanim "
                  "razlogom)" % (len(katalog_suita()), len(SUITE_VAN_KAPIJA)))
            return 0
        for n in nalazi:
            print(n, file=sys.stderr)
        print("\npopis suita: %d nalaza" % len(nalazi), file=sys.stderr)
        return 2

    ap.print_help()
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
