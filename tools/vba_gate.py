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
    python tools/vba_gate.py --require-green --sveska druga.xlsm
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

Zato se hesira, a marker pamti REZULTAT PO SUITE-U uz otiske pod kojima je
nastao -- `--require-green` time ume da razlikuje "dokazano" od "dokazano nesto
drugo", i da imenuje koja suite nedostaje.

DVA OTISKA, NE JEDAN. Otisak samog `src-vba` ne dokazuje da je dokaz izvrsen nad
OVIM test sistemom:

    izvor    = src-vba
    ugovor   = izvor + tools/run_vba.py + tools/make_fixture.py
               + tests/golden/* + verzija markera

`RunGoldenSuite` meri ishod protiv `tests/golden/*.txt`; `run_vba.py` odlucuje
koja suite postoji, da li je `gate` i kako se cita rezultat; `make_fixture.py`
odredjuje podatke nad kojima testovi rade. Promeni golden fajl, ne pusti nijedan
test, i marker vezan samo za izvor bi i dalje tvrdio "dokazano" -- a trenutni
golden ugovor nikad nije bio izvrsen nad tim izvorom.

Razdvojena su zato sto COMPILE pripada samo izvoru: `Debug > Compile` ne zna za
golden fajlove, pa potvrda ne sme da propadne zato sto se jedan promenio. Ugovor
se pritom racuna iz DELOVA, pa nalaz ume da kaze KOJI se deo promenio, a ne samo
"nesto".

KONTEKST SVESKE se pamti uz svaki rezultat, i to kao IDENTITET -- putanja plus
hes sadrzaja, ne ime. Basename ne razlikuje `C:\\A\\test.xlsm` od
`D:\\B\\test.xlsm`, a podrazumevani fixture je gitignored i regenerise se: bez
hesa sadrzaja bi zamena sveske ostavila prethodni GREEN na nogama, iako
`make_fixture.signature()` pokriva samo deklarativni seed, ne celu svesku
izvedenu iz donora. Zato `run_vba --workbook X.xlsm` ne zadovoljava podrazumevani
zahtev, a ni zamenjen fixture ne prolazi pod starim dokazom.

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

# VERZIJA MARKERA. Ako se znacenje polja promeni, stari upis NE SME da zadovolji
# nova pravila -- inace bi prelazak na strozija pravila tiho priznao dokaze
# napravljene pod slabijim. Zato verzija ulazi i u otisak ugovora.
# 4: upis je od tada fail-closed nad snimkom konteksta uzetim PRE run-a.
# Marker verzije 3 je mogao nastati bez te provere, pa mu se ne veruje.
MARKER_VERZIJA = 4

# Delovi TEST-UGOVORA koji zive van src-vba, IMENOVANI -- da nalaz kaze koji se
# deo promenio, a ne samo "nesto".
#
# `kapija` je ovaj fajl. Bez njega alat hvata promenu test RUNNERA, a ne i
# promenu VERIFIKATORA: `vba_gate.py` odlucuje sta znaci `--require-green`, koje
# se suite traze, kako se porede otisci i sta se priznaje kao OK. Promeni
# acceptance logiku, zaboravi da bumpujes MARKER_VERZIJA, i stari dokaz ostaje
# "vazeci" pod novim pravilima. MARKER_VERZIJA ostaje -- ali kao izricita
# oznaka namere, ne kao jedina brana koja zavisi od toga da se covek seti.
UGOVOR_FAJLOVI = {
    "runner": ("tools/run_vba.py",),
    "fixture": ("tools/make_fixture.py",),
    "kapija": ("tools/vba_gate.py",),
}
GOLDEN_PODFOLDER = "tests/golden"
PODRAZUMEVANA_SVESKA = "tests/fixtures/otkup_test.xlsm"

# Modul koji stamp-build prepisuje pred svaki import -- v. docstring.
IZUZET_IZ_OTISKA = ("modBuildInfo.bas",)
TEKST_NASTAVCI = (".bas", ".cls", ".frm", ".doccls")

# IME ulazne tacke suite-a po konvenciji. `RunProductionHealthCheck` joj ne
# odgovara, i ne mora: sve iz `SUITES` se proverava po IMENU (da procedura
# postoji), a konvencija sluzi samo da nadje ono sto u katalogu NIJE.
#
# Vidljivost i argumenti se NE proveravaju ovde nego kroz
# `vba_check.deklaracija_procedure` -- jedan razlagac deklaracije za ceo tooling
# sloj. Dva promasaja koja su se tu vec platila: izraz je trazio literalni
# "Public Sub" (a modifikator je opcion, default Public), pa literalne "()" (a
# zagrade su opcione, pa je `Sub RunFooSuite` validna javna suite bez
# argumenata). Oba su bila nevidljiva, i oba su obarala glavnu tvrdnju popisa.
IME_SUITE = re.compile(r"^(?:Run\w*(?:Suite|Tests)|Test\w*_All)$")

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
    "RunSEFOfflineSuite": (
        "SEF tok bez mreze. Nasao je razlagac, jer je deklarisana sa JEDNIM "
        "OPCIONIM argumentom (`Optional ByVal fakturaID As String = \"\"`) -- "
        "stari izraz je trazio praznu listu i preskakao je. Telo nema Err.Raise "
        "nego LogFatal, pa bi prikljucena bila BLIND: zatvara se Err.Raise-om na "
        "kraju, pa unosom u SUITES sa gate: True."),
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


def _razlagac():
    """`vba_check.deklaracija_procedure` -- jedan razlagac za ceo tooling sloj.

    Uvozi se LENJO, iz funkcije: `vba_check` sa svoje strane uvozi ovaj modul
    zbog kapije popisa, pa bi uvoz na nivou modula bio kruzan. Na toj (hook)
    putanji vba_check razlagac PREDAJE, pa se ne uvozi dvaput; ovaj put placa
    samo samostalni `--popis`.
    """
    put = os.path.join(ROOT, "tools", "vba_check.py")
    spec = importlib.util.spec_from_file_location("_vba_check_za_gate", put)
    modul = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modul)
    return modul.deklaracija_procedure


def skeniraj(src_dir: str = SRC_VBA, razlagac=None) -> tuple:
    """(javne, po_konvenciji) iz JEDNOG prolaza kroz src-vba.

    Jedan prolaz, ne dva: provera ide kroz `vba_check`, dakle kroz PostToolUse
    hook, pa se citanje 190 fajlova ne placa dvaput.

    `javne` su sve javne procedure modula (ime -> fajl) i sluze da se proveri
    POSTOJANJE imena iz kataloga -- ukljucujuci implicitno javne, bez
    modifikatora.

    `po_konvenciji` su kandidati za ulaznu tacku suite-a, i to su TRI uslova nad
    razlozenom deklaracijom, ne jedan izraz:

      javna (Public ili bez modifikatora)
      + NULA OBAVEZNIH argumenata -- `Sub Foo`, `Sub Foo()` i
        `Sub Foo(Optional x)` se sve zovu po imenu bez argumenta
      + ime po konvenciji, i samo u `.bas`

    `.bas` jer se makro po imenu zove jedino iz standardnog modula: `Sub
    RunFooSuite()` u klasi ili formi nije suite koju bi kapija mogla da pokrene.
    `Function` se prihvata uz `Sub`: i nju `Application.Run` zove po imenu, pa bi
    inace bila ista rupa drugog oblika.
    """
    razlagac = _razlagac() if razlagac is None else razlagac
    javne, po_konvenciji = {}, {}
    if not os.path.isdir(src_dir):
        return javne, po_konvenciji
    for ime in sorted(os.listdir(src_dir)):
        if not ime.endswith(TEKST_NASTAVCI):
            continue
        with io.open(os.path.join(src_dir, ime), encoding="ascii",
                     errors="replace", newline="") as fh:
            tekst = fh.read().replace("\r\n", "\n")
        for red in tekst.split("\n"):
            d = razlagac(red)
            if not d or not d["javna"] or d["vrsta"] == "property":
                continue
            bez_arg = d["obaveznih"] == 0
            javne.setdefault(d["ime"], (ime, bez_arg))
            if ime.endswith(".bas") and bez_arg and IME_SUITE.match(d["ime"]):
                po_konvenciji.setdefault(d["ime"], ime)
    return javne, po_konvenciji


def popis_problemi(suites: dict = None, registar: dict = None,
                   src_dir: str = SRC_VBA, skenirano: tuple = None,
                   razlagac=None) -> list:
    """Nalazi iz popisa suita. Prazna lista = cisto."""
    suites = katalog_suita() if suites is None else suites
    registar = SUITE_VAN_KAPIJA if registar is None else registar
    if skenirano is None:
        javne, po_konvenciji = skeniraj(src_dir, razlagac)
    else:
        javne, po_konvenciji = skenirano

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


def _hash_fajlova(putanje: list) -> str:
    """SHA256 nad skupom fajlova; nepostojeci fajl je deo odgovora, ne greska."""
    h = hashlib.sha256()
    for p in sorted(putanje):
        h.update(os.path.basename(p).encode())
        try:
            with open(p, "rb") as fh:
                h.update(fh.read().replace(b"\r\n", b"\n"))
        except OSError:
            h.update(b"<nema fajla>")
    return h.hexdigest()


def _hash_sirov(put: str) -> str:
    """SHA256 sirovog sadrzaja, BEZ normalizacije preloma.

    Za `.xlsm` (zip) normalizacija CRLF->LF nije samo nepotrebna nego i stetna:
    sazimala bi bajt-par koji u komprimovanom sadrzaju nema nikakvo znacenje
    kraja reda, i time bez razloga smanjivala otpornost na sudar.
    """
    h = hashlib.sha256()
    try:
        with open(put, "rb") as fh:
            for blok in iter(lambda: fh.read(1 << 20), b""):
                h.update(blok)
    except OSError:
        return ""
    return h.hexdigest()


def podrazumevana_sveska(koren: str = ROOT) -> str:
    return os.path.join(koren, *PODRAZUMEVANA_SVESKA.split("/"))


def kontekst_sveske(put: str, hes_iz: str = None) -> dict:
    """Identitet sveske nad kojom je dokaz nastao: PUTANJA i SADRZAJ.

    `hes_iz` je fajl iz koga se cita sadrzaj kad to nije ista putanja:
    `run_vba` predaje TEMP KOPIJU, jer je ona ono sto Excel stvarno otvara.
    Hes uzet iz nje nema prozor izmedju "procitao sam fixture radi hesa" i
    "kopirao sam ga Excelu" -- a putanja ostaje IZVORNA, da se dokaz moze
    uporediti sa svescom koja i dalje stoji na disku.

    Basename nije identitet. `C:\\A\\test.xlsm` i `D:\\B\\test.xlsm` se po
    njemu ne razlikuju; gore od toga, podrazumevani fixture je gitignored i
    regenerise se, pa je zamena sveske bez ijedne izmene u izvoru, runneru,
    generatoru ili golden-u ostavljala prethodni GREEN na nogama.

    `make_fixture.signature()` to ne pokriva i nije mu namena: on opisuje
    DEKLARATIVNI seed/config, a sveska je izvedena iz donora -- pa dve razlicite
    sveske mogu imati isti potpis generatora. Zato ide hes sadrzaja.

    Hesira se IZVORNA sveska, ne temp kopija koju Excel menja: `run_vba` je prvo
    kopira u temp, pa je izvor stabilan i posle run-a.
    """
    return {"ime": os.path.basename(put),
            "putanja": os.path.realpath(put),
            "otisak": _hash_sirov(hes_iz or put)}


def _razlika_sveske(zapisan: dict, trazen: dict) -> str:
    """Zasto zapisana sveska nije ona koja se trazi -- prazno ako jeste."""
    zapisan = zapisan or {}
    if zapisan.get("otisak") != trazen.get("otisak"):
        return ("sadrzaj sveske je drugi (zapisan %s, sada %s)"
                % ((zapisan.get("otisak") or "?")[:12],
                   (trazen.get("otisak") or "?")[:12]))
    if zapisan.get("putanja") != trazen.get("putanja"):
        return "sveska je sa druge putanje (%s)" % zapisan.get("putanja")
    return ""


def ugovor_delovi(src_dir: str = SRC_VBA, koren: str = ROOT) -> dict:
    """Delovi TEST-UGOVORA: sve od cega zavisi sta "zeleno" znaci.

    Otisak src-vba sam po sebi ne dokazuje da je dokaz izvrsen nad OVIM test
    sistemom. `RunGoldenSuite` meri ishod protiv `tests/golden/*.txt`;
    `run_vba.py` odlucuje koja suite postoji, da li je `gate` i kako se cita
    rezultat; `make_fixture.py` odredjuje podatke nad kojima testovi rade (i
    potpis koji `run_vba` proverava PRE Excela, cime sam priznaje da je stanje
    fixture-a deo preduslova). Promeni bilo sta od toga bez novog run-a, i stari
    marker bi i dalje tvrdio "dokazano".

    Racuna se iz DELOVA, a ne kao jedan veliki hes, iz dva razloga: `--status`
    ume da imenuje KOJI se deo promenio, i COMPILE se vezuje samo za `izvor` --
    potvrda `Debug > Compile` ne sme da propadne zato sto se promenio golden
    fajl, jer compile o golden fajlovima ne zna nista.

    `make_fixture.signature()` se ne zove posebno: on je cista funkcija SEED-a iz
    istog fajla, pa ga hes fajla vec pokriva -- a izbegava se uvoz modula od
    2500 linija na hook putanji.
    """
    golden = os.path.join(koren, *GOLDEN_PODFOLDER.split("/"))
    try:
        fajlovi = [os.path.join(golden, f) for f in sorted(os.listdir(golden))]
    except OSError:
        fajlovi = []
    delovi = {
        "verzija": MARKER_VERZIJA,
        "izvor": otisak_izvora(src_dir),
        "golden": _hash_fajlova(fajlovi),
    }
    for deo, putanje in UGOVOR_FAJLOVI.items():
        delovi[deo] = _hash_fajlova(
            [os.path.join(koren, *p.split("/")) for p in putanje])
    return delovi


def otisak_ugovora(delovi: dict = None, src_dir: str = SRC_VBA,
                   koren: str = ROOT) -> str:
    delovi = ugovor_delovi(src_dir, koren) if delovi is None else delovi
    return hashlib.sha256(
        json.dumps(delovi, sort_keys=True).encode()).hexdigest()


def _razlika_ugovora(stari: dict, novi: dict) -> list:
    """Imena delova ugovora koji se razlikuju -- da nalaz kaze STA se promenilo."""
    stari = stari or {}
    return sorted(k for k in set(stari) | set(novi)
                  if stari.get(k) != novi.get(k))


def snimi_kontekst(src_dir: str = SRC_VBA, koren: str = ROOT,
                   sveska: str = None, sveska_hes_iz: str = None) -> dict:
    """Nepromenljiv snimak konteksta dokaza, uzet PRE run-a.

    Otisci racunati POSLE run-a opisuju stanje diska na kraju, a ne ono sto je
    testirano. Prolaz traje 20-60 minuta i razvoj ide paralelno, pa je prozor
    stvaran: izmeni `src-vba`, golden ili fixture tokom run-a, i GREEN bi bio
    pripisan stanju koje Excel nikad nije video. Zato se kontekst snima na
    ulasku, a `zabelezi_prolaz` ga samo PRIMA i pre upisa tvrdi da se nije
    promenio.
    """
    delovi = ugovor_delovi(src_dir, koren)
    return {
        "verzija": MARKER_VERZIJA,
        "izvor": delovi["izvor"],
        "ugovor": otisak_ugovora(delovi),
        "ugovor_delovi": delovi,
        "sveska": kontekst_sveske(sveska or podrazumevana_sveska(koren),
                                 sveska_hes_iz),
    }


def _razlika_konteksta(pre: dict, posle: dict) -> list:
    """Sta se promenilo izmedju snimka i trenutnog stanja. Prazno = nista."""
    pre, posle = pre or {}, posle or {}
    razlike = []
    if pre.get("verzija") != posle.get("verzija"):
        razlike.append("verzija markera")
    if pre.get("izvor") != posle.get("izvor"):
        razlike.append("src-vba")
    if pre.get("ugovor") != posle.get("ugovor"):
        # `izvor` je DEO ugovora, i prijavljen je vec iznad. Ovde se imenuje samo
        # ono sto je ugovoru specificno -- inace bi ista promena bila prijavljena
        # dva puta, a poredjenje izvora iznad ne bi bilo nezavisno merljivo:
        # ugasis ga, a ugovor i dalje hvata istu promenu.
        delovi = [d for d in _razlika_ugovora(pre.get("ugovor_delovi"),
                                             posle.get("ugovor_delovi") or {})
                  if d != "izvor"]
        if delovi:
            razlike.append("test-ugovor (%s)" % ", ".join(delovi))
    if _razlika_sveske(pre.get("sveska"), posle.get("sveska") or {}):
        razlike.append("sveska")
    return razlike


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
                    kontekst: dict = None, put: str = MARKER,
                    src_dir: str = SRC_VBA, koren: str = ROOT) -> str:
    """Zapisi rezultat run-a u marker. Vraca poruku (sta je upisano ili zasto nije).

    Zove je `run_vba.py` na kraju run-a. Pravila su u docstring-u modula; ovde su
    kao kod, jer su upravo ona ono sto marker cini tvrdnjom a ne dekoracijom.

    Otisci se pamte PO SUITE-U, ne po markeru. Delimican run (`--suite X`) time
    ne brise tudje rezultate, ali ni ne pozajmljuje svoje otiske njima: svaki
    upis nosi izvor i ugovor pod kojim je nastao, pa `--require-green` posle
    izmene ume da kaze koja je suite zastarela, a koja nije.

    `kontekst` je snimak uzet PRE run-a (v. `snimi_kontekst`). Ne racuna se ovde
    iznova: otisci uzeti na kraju opisuju stanje diska posle testova, a ne ono
    sto je testirano. Pre upisa se TVRDI da se kontekst nije promenio -- ako se
    promenio, run moze biti zelen ali marker se NE UPISUJE. Fail-closed: bolje
    nema dokaza nego dokaz pripisan stanju koje nije mereno.
    """
    if no_import:
        return ("marker nije upisan: --no-import znaci da kod u svesci nije "
                "src-vba, pa otisak ne bi opisivao ono sto je izvrseno")
    if rc != 0:
        return "marker nije upisan: run nije zelen (rc=%s)" % rc
    if not kontekst:
        return ("marker nije upisan: nema snimka konteksta uzetog PRE run-a "
                "(v. snimi_kontekst) -- bez njega bi se potpisalo stanje diska "
                "posle testova, a ne ono sto je testirano")

    # `kontekst or {}`: ugasena kapija iznad ne sme da zavrsi kao AttributeError
    # nego kao ODBIJEN UPIS -- pad nije merenje. Prazan kontekst se razlikuje od
    # trenutnog stanja po verziji, pa upis pada fail-closed.
    k = kontekst or {}
    sada = snimi_kontekst(src_dir, koren,
                          (k.get("sveska") or {}).get("putanja"))
    razlike = _razlika_konteksta(k, sada)
    if razlike:
        return ("marker nije upisan: %s se promenilo TOKOM run-a -- prolaz je "
                "mozda zelen, ali mereno stanje vise ne stoji na disku"
                % ", ".join(razlike))

    delovi = k.get("ugovor_delovi") or {}
    ugovor = k.get("ugovor")
    podaci = procitaj_marker(put)
    if not podaci or podaci.get("verzija") != MARKER_VERZIJA:
        podaci = {"verzija": MARKER_VERZIJA, "suites": {}}
    podaci.setdefault("suites", {})

    kada = time.strftime("%Y-%m-%dT%H:%M:%S")
    ks = k.get("sveska") or {}
    rezultati = report.get("suite_results", {}) or {}
    for s in report.get("suites", []):
        r = rezultati.get(s["name"]) or {}
        podaci["suites"][s["name"]] = {
            "status": s.get("status"),
            "ukupno": r.get("total"),
            "palo": r.get("failed"),
            "izvor": delovi.get("izvor"),
            "ugovor": ugovor,
            "ugovor_delovi": delovi,
            "sveska": ks,
            "kada": kada,
            "git": _git_glava(),
            "platforma": platform.platform(),
        }
    upisi_marker(podaci, put)
    imena = sorted(s["name"] for s in report.get("suites", []))
    return ("marker upisan (izvor %s, ugovor %s, sveska %s/%s): %d suita (%s)"
            % ((delovi.get("izvor") or "?")[:12], (ugovor or "?")[:12],
               ks.get("ime"), (ks.get("otisak") or "?")[:8], len(imena),
               ", ".join(imena) or "nijedna"))


def zabelezi_compile(put: str = MARKER, src_dir: str = SRC_VBA) -> str:
    """Operater je potvrdio `Debug > Compile` nad OVIM IZVOROM.

    Vezuje se SAMO za izvor, ne za test-ugovor: compile ne zna za golden fajlove
    ni za runner, pa ne sme da izgubi potvrdu zato sto se jedan golden promenio.
    Zato i ne brise rezultate suita -- oni nose svoje otiske.
    """
    otisak = otisak_izvora(src_dir)
    podaci = procitaj_marker(put)
    if not podaci or podaci.get("verzija") != MARKER_VERZIJA:
        podaci = {"verzija": MARKER_VERZIJA, "suites": {}}
    podaci["compile"] = {"izvor": otisak,
                         "kada": time.strftime("%Y-%m-%dT%H:%M:%S"),
                         "git": _git_glava()}
    upisi_marker(podaci, put)
    return "compile potvrdjen nad izvorom %s" % otisak[:12]


def zahtevaj_zeleno(trazene: list = None, suites: dict = None,
                    trazi_compile: bool = False, sveska: str = None,
                    put: str = MARKER, src_dir: str = SRC_VBA,
                    koren: str = ROOT) -> list:
    """Nalazi zbog kojih OVAJ izvor nije dokazan. Prazna lista = dokazan.

    Trazi se poklapanje OBA otiska po suite-u: izvor (src-vba) i test-ugovor
    (runner + fixture generator + golden + verzija markera). Otisak izvora sam
    ne bi razlikovao "dokazano" od "dokazano pod drugim test sistemom".

    `sveska` je PUTANJA sveske nad kojom se dokaz priznaje; podrazumevano je to
    fixture. Poredi se identitet, ne ime: putanja i hes sadrzaja. Run nad tudjom
    svescom (`run_vba --workbook X.xlsm`) zato ne zadovoljava podrazumevani
    zahtev, a ni zamenjen fixture ne prolazi pod starim dokazom.
    """
    suites = katalog_suita() if suites is None else suites
    trazene = potrebne_suite(suites) if trazene is None else list(trazene)
    delovi = ugovor_delovi(src_dir, koren)
    ugovor = otisak_ugovora(delovi)
    trazena = kontekst_sveske(sveska or podrazumevana_sveska(koren))
    marker = procitaj_marker(put)

    if not marker:
        return ["nema markera: nijedan prolaz nije zapisan (pusti "
                "`python tools/run_vba.py`)"]
    if marker.get("verzija") != MARKER_VERZIJA:
        return ["marker je verzije %s, a trazi se %s -- znacenje polja se "
                "promenilo, pa stari upis ne vazi (pusti run ponovo)"
                % (marker.get("verzija"), MARKER_VERZIJA)]

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
        elif z.get("izvor") != delovi["izvor"]:
            nalazi.append("%s: dokazan je DRUGI izvor (%s, sada %s)"
                          % (ime, (z.get("izvor") or "?")[:12],
                             delovi["izvor"][:12]))
        elif z.get("ugovor") != ugovor:
            nalazi.append("%s: TEST-UGOVOR je promenjen posle dokaza (razlika u: "
                          "%s) -- suite nije pustena nad ovim test sistemom"
                          % (ime, ", ".join(_razlika_ugovora(
                              z.get("ugovor_delovi"), delovi)) or "?"))
        elif _razlika_sveske(z.get("sveska"), trazena):
            nalazi.append("%s: %s -- dokaz nije napravljen nad svescom koja se "
                          "trazi (%s)"
                          % (ime, _razlika_sveske(z.get("sveska"), trazena),
                             trazena["ime"]))
    if trazi_compile:
        c = marker.get("compile") or {}
        if not c:
            nalazi.append("compile nije potvrdjen nad ovim izvorom "
                          "(`--mark-compile` posle Debug > Compile VBAProject)")
        elif c.get("izvor") != delovi["izvor"]:
            nalazi.append("compile je potvrdjen nad DRUGIM izvorom (%s, sada %s)"
                          % ((c.get("izvor") or "?")[:12],
                             delovi["izvor"][:12]))
    return nalazi


def stanje_redovi(suites: dict = None, put: str = MARKER,
                  src_dir: str = SRC_VBA, koren: str = ROOT) -> list:
    suites = katalog_suita() if suites is None else suites
    delovi = ugovor_delovi(src_dir, koren)
    ugovor = otisak_ugovora(delovi)
    marker = procitaj_marker(put)
    sveska_sada = kontekst_sveske(podrazumevana_sveska(koren))
    redovi = ["izvor:   %s" % delovi["izvor"][:16],
              "ugovor:  %s  (%s, verzija %s)"
              % (ugovor[:16],
                 ", ".join("%s %s" % (k, (delovi[k] or "?")[:8])
                           for k in sorted(UGOVOR_FAJLOVI) + ["golden"]),
                 delovi["verzija"]),
              "sveska:  %s  %s" % (sveska_sada["ime"],
                                   sveska_sada["otisak"][:16] or "NEMA JE")]
    if not marker:
        redovi.append("marker:  nema ga -- nijedan prolaz nije zapisan")
        return redovi
    if marker.get("verzija") != MARKER_VERZIJA:
        redovi.append("marker:  verzija %s, trazi se %s -- stari upis ne vazi"
                      % (marker.get("verzija"), MARKER_VERZIJA))
        return redovi
    c = marker.get("compile") or {}
    redovi.append("compile: %s" % (
        "nije potvrdjen" if not c else
        "potvrdjen %s nad %s%s" % (c.get("kada"), (c.get("izvor") or "?")[:12],
                                   "" if c.get("izvor") == delovi["izvor"]
                                   else "  <-- DRUGI IZVOR")))
    zapisane = marker.get("suites") or {}
    trazene = set(potrebne_suite(suites))
    for ime in sorted(set(zapisane) | trazene):
        z = zapisane.get(ime) or {}
        oznaka = "*" if ime in trazene else " "
        if not z:
            redovi.append(" %s %-28s --" % (oznaka, ime))
            continue
        beleska = ""
        if z.get("izvor") != delovi["izvor"]:
            beleska = "DRUGI IZVOR"
        elif z.get("ugovor") != ugovor:
            beleska = "UGOVOR: " + ", ".join(
                _razlika_ugovora(z.get("ugovor_delovi"), delovi))
        elif _razlika_sveske(z.get("sveska"), sveska_sada):
            beleska = "SVESKA: " + _razlika_sveske(z.get("sveska"), sveska_sada)
        redovi.append(" %s %-28s %-6s %s/%s  %s"
                      % (oznaka, ime, z.get("status"), z.get("palo"),
                         z.get("ukupno"), beleska))
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
    # PRAVI deljeni razlagac, ne kopija: self-test time meri i to da je
    # definicija "javne procedure" stvarno jedna za ceo tooling sloj.
    RAZLAGAC = _razlagac()

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
            return popis_problemi(suites, registar or {}, src_dir,
                                  razlagac=RAZLAGAC)

        tvrdi(not popis(), "POPIS: cist izvor daje nalaz")

        # Razlagac odbija red koji NIJE deklaracija (ime pa nesto trece), a
        # prihvata deklaraciju sa tipom povratka bez zagrada.
        tvrdi(RAZLAGAC("Sub RunNijeSuite: Bar") is None,
              "RAZLAGAC: red koji nije deklaracija se razlaze kao deklaracija")
        tvrdi((RAZLAGAC("Function RunTipSuite As String") or {}).get(
                  "obaveznih") == 0,
              "RAZLAGAC: deklaracija sa tipom povratka bez zagrada se odbija")

        # IMPLICITNO JAVNA suite: `Sub RunX()` je u VBA Public po defaultu.
        # Prva verzija je trazila literalni "Public Sub" i ovo je bilo
        # nevidljivo -- validna javna suite, van kataloga, a CI zelen.
        with io.open(os.path.join(src, "modImp.bas"), "w", newline="") as fh:
            fh.write("Sub RunImplicitnaSuite()\r\nEnd Sub\r\n")
        tvrdi(any("NEPOKRETANA" in n and "RunImplicitnaSuite" in n
                  for n in popis()),
              "POPIS: implicitno javna suite (`Sub RunX()`) je nevidljiva")
        os.remove(os.path.join(src, "modImp.bas"))

        # PRIVATE nije javna, pa nije ni suite koju kapija moze da pokrene.
        with io.open(os.path.join(src, "modPriv.bas"), "w", newline="") as fh:
            fh.write("Private Sub RunPrivatnaSuite()\r\nEnd Sub\r\n")
        tvrdi(not popis(), "POPIS: `Private Sub` se broji kao javna suite")
        os.remove(os.path.join(src, "modPriv.bas"))

        # Function se takodje zove po imenu kroz Application.Run.
        with io.open(os.path.join(src, "modFun.bas"), "w", newline="") as fh:
            fh.write("Function RunFunkcijaSuite()\r\nEnd Function\r\n")
        tvrdi(any("RunFunkcijaSuite" in n for n in popis()),
              "POPIS: implicitno javna Function-suite je nevidljiva")
        os.remove(os.path.join(src, "modFun.bas"))

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

        # BEZ ZAGRADA. VBA: `Sub name [ ( arglist ) ]` -- zagrade su opcione, pa
        # je ovo validna javna suite bez argumenata. Izraz koji je trazio "()" ju
        # je potpuno promasivao.
        with io.open(os.path.join(src, "modBZ.bas"), "w", newline="") as fh:
            fh.write("Sub RunBezZagradaSuite\r\nEnd Sub\r\n")
        tvrdi(any("NEPOKRETANA" in n and "RunBezZagradaSuite" in n
                  for n in popis()),
              "POPIS: suite BEZ ZAGRADA (`Sub RunFooSuite`) je nevidljiva")
        os.remove(os.path.join(src, "modBZ.bas"))

        # Kontrola: procedura SA obaveznim argumentom nije ulazna tacka.
        with io.open(os.path.join(src, "modP.bas"), "w", newline="") as fh:
            fh.write("Public Sub RunNestoSuite(ByVal x As Long)\r\nEnd Sub\r\n")
        tvrdi(not popis(), "POPIS: procedura SA argumentima se broji kao suite")
        os.remove(os.path.join(src, "modP.bas"))

        # Isto i bez zagrada u kontroli: `Sub RunSaArgumentomSuite(ByVal x)`.
        with io.open(os.path.join(src, "modPA.bas"), "w", newline="") as fh:
            fh.write("Sub RunSaArgumentomSuite(ByVal x As Long)\r\n"
                     "End Sub\r\n")
        tvrdi(not popis(),
              "POPIS: implicitno javna procedura SA argumentom je kandidat")
        os.remove(os.path.join(src, "modPA.bas"))

        # OPCIONI argument nije obavezan -- takva se zove po imenu bez argumenta,
        # pa JE ulazna tacka.
        with io.open(os.path.join(src, "modOP.bas"), "w", newline="") as fh:
            fh.write("Sub RunOpcioniSuite(Optional ByVal x As Long = 1)\r\n"
                     "End Sub\r\n")
        tvrdi(any("RunOpcioniSuite" in n for n in popis()),
              "POPIS: suite sa samo OPCIONIM argumentom je nevidljiva")
        os.remove(os.path.join(src, "modOP.bas"))

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

        # --- marker: izvor I test-ugovor ----------------------------------
        #
        # Ugovor zivi van src-vba (runner, fixture generator, golden), pa self-test
        # gradi sopstveni KOREN u temp folderu: time se promena golden fajla moze
        # izmeriti bez diranja repoa.
        put = os.path.join(tmp, "tests", "last_green.json")
        os.makedirs(os.path.join(tmp, "tools"), exist_ok=True)
        os.makedirs(os.path.join(tmp, "tests", "golden"), exist_ok=True)
        os.makedirs(os.path.join(tmp, "tests", "fixtures"), exist_ok=True)
        for rel in ("tools/run_vba.py", "tools/make_fixture.py",
                    "tools/vba_gate.py", "tests/golden/G1.txt"):
            with io.open(os.path.join(tmp, *rel.split("/")), "w",
                         newline="") as fh:
                fh.write("prvo stanje " + rel + "\n")
        # Podrazumevani fixture mora da POSTOJI: bez njega bi hes bio prazan sa
        # obe strane, pa bi poredjenje sveske prolazilo ne merivsi nista.
        with open(podrazumevana_sveska(tmp), "wb") as fh:
            fh.write(b"fixture A")

        ZELEN = {"suites": [{"name": "RunAllTests", "status": "OK"}],
                 "suite_results": {"RunAllTests": {"total": 199, "failed": 0}}}

        def upisi(rep=None, rc=0, sveska=None, **kw):
            # Snimak se uzima neposredno pre upisa: normalna putanja, gde se
            # izmedju snimka i upisa nije nista promenilo.
            k = snimi_kontekst(src, tmp, sveska)
            return zabelezi_prolaz(rep or ZELEN, rc, kontekst=k, put=put,
                                   src_dir=src, koren=tmp, **kw)

        def zahtevaj(**kw):
            # Pad je NALAZ, ne traceback: ugasena kapija `if not marker` inace
            # rusi self-test, pa se ne vidi kao sopstvena greska.
            kw.setdefault("trazene", ["RunAllTests"])
            try:
                return zahtevaj_zeleno(suites=SUITES, put=put, src_dir=src,
                                       koren=tmp, **kw)
            except Exception as e:          # noqa: BLE001
                nalazi.append("MARKER: zahtevaj_zeleno je pukao -- %r" % (e,))
                return []

        tvrdi(any("nema markera" in n for n in zahtevaj()),
              "MARKER: bez markera je izvor 'dokazan'")

        poruka = upisi(rc=2)
        tvrdi("nije zelen" in poruka and not os.path.exists(put),
              "MARKER: PAO run upisuje marker")

        poruka = upisi(no_import=True)
        tvrdi("no-import" in poruka and not os.path.exists(put),
              "MARKER: --no-import run upisuje marker")

        upisi()
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
        upisi(BLIND)
        tvrdi(any("BLIND" in n for n in zahtevaj(trazene=["RunNovacSmokeSuite"])),
              "MARKER: BLIND suite se priznaje kao dokaz")
        tvrdi(not zahtevaj(),
              "MARKER: drugi prolaz je pokvario rezultat prvog")

        # ISTI IZVOR, PROMENJEN GOLDEN. Scenario iz review-a #402: RunGoldenSuite
        # meri ishod protiv tests/golden/*.txt, pa promena ocekivanja bez novog
        # run-a ne sme da ostavi stari dokaz na nogama.
        with io.open(os.path.join(tmp, "tests", "golden", "G1.txt"), "w",
                     newline="") as fh:
            fh.write("drugo stanje\n")
        poruke = zahtevaj()
        tvrdi(any("TEST-UGOVOR" in n and "golden" in n for n in poruke),
              "MARKER: promenjen GOLDEN ostavlja stari dokaz vazecim")
        tvrdi(not zahtevaj(trazi_compile=True) or all(
                  "compile" not in n for n in zahtevaj(trazi_compile=True)),
              "MARKER: promenjen golden obara i POTVRDU COMPILE-a (ne sme)")
        upisi()
        tvrdi(not zahtevaj(), "MARKER: nov run pod novim ugovorom nije dokaz")

        # ISTI IZVOR, PROMENJEN RUNNER. run_vba odlucuje koja suite postoji, da li
        # je gate i kako se cita rezultat -- dakle sta "zeleno" uopste znaci.
        with io.open(os.path.join(tmp, "tools", "run_vba.py"), "w",
                     newline="") as fh:
            fh.write("drugo stanje runnera\n")
        tvrdi(any("TEST-UGOVOR" in n and "runner" in n for n in zahtevaj()),
              "MARKER: promenjen RUNNER ostavlja stari dokaz vazecim")
        upisi()
        tvrdi(not zahtevaj(), "MARKER: nov run pod novim runnerom nije dokaz")

        # PROMENJENA KAPIJA. `vba_gate.py` odlucuje sta se priznaje kao dokaz --
        # koje suite se traze, kako se porede otisci, sta je OK. Njegova izmena
        # mora da obori stare dokaze SAMA, bez rucnog bumpa MARKER_VERZIJA: ta
        # rucna disciplina je bas ono sto alat zamenjuje.
        with io.open(os.path.join(tmp, "tools", "vba_gate.py"), "w",
                     newline="") as fh:
            fh.write("druga semantika kapije\n")
        tvrdi(any("TEST-UGOVOR" in n and "kapija" in n for n in zahtevaj()),
              "MARKER: promenjena KAPIJA ostavlja stari dokaz vazecim")
        upisi()
        tvrdi(not zahtevaj(), "MARKER: nov run pod novom kapijom nije dokaz")

        # TUDJA SVESKA, i to sa ISTIM BASENAME-om -- bas bypass iz review-a.
        # Identitet je putanja + sadrzaj, pa ime nije dovoljno.
        druga = os.path.join(tmp, "druga", "otkup_test.xlsm")
        treca = os.path.join(tmp, "treca", "otkup_test.xlsm")
        for p in (druga, treca):
            os.makedirs(os.path.dirname(p), exist_ok=True)
            with open(p, "wb") as fh:
                fh.write(b"tudja sveska")
        upisi(sveska=druga)
        tvrdi(any("sadrzaj sveske" in n for n in zahtevaj()),
              "MARKER: dokaz nad tudjom svescom zadovoljava podrazumevani zahtev")
        tvrdi(not zahtevaj(sveska=druga),
              "MARKER: izricito trazena sveska se ne priznaje")
        # Isti sadrzaj, druga putanja: `C:\\A\\test.xlsm` vs `D:\\B\\test.xlsm`.
        tvrdi(any("druge putanje" in n for n in zahtevaj(sveska=treca)),
              "MARKER: ista sveska sa DRUGE putanje se priznaje")
        upisi()
        tvrdi(not zahtevaj(),
              "MARKER: vracanje na fixture nije priznato")

        # ZAMENJEN FIXTURE. Podrazumevani fixture je gitignored i regenerise se;
        # `make_fixture.signature()` pokriva deklarativni seed, ne celu svesku
        # izvedenu iz donora -- pa dve razlicite sveske mogu imati isti potpis
        # generatora. Bez hesa sadrzaja zamena bi prosla neopazeno.
        with open(podrazumevana_sveska(tmp), "wb") as fh:
            fh.write(b"fixture B")
        tvrdi(any("sadrzaj sveske" in n for n in zahtevaj()),
              "MARKER: zamenjen fixture ostavlja stari dokaz vazecim")
        upisi()

        # IZMENA IZVORA obara dokaz suita, i potvrdu compile-a -- compile je vezan
        # za izvor, pa mu izvor i jeste jedina osa.
        with io.open(os.path.join(src, "modTest.bas"), "a", newline="") as fh:
            fh.write("' jos jedna izmena\r\n")
        tvrdi(any("DRUGI izvor" in n for n in zahtevaj()),
              "MARKER: izmena izvora ne obara dokaz suite")
        upisi()
        tvrdi(any("compile" in n and "DRUGIM izvorom" in n
                  for n in zahtevaj(trazi_compile=True)),
              "MARKER: nov izvor nasledjuje staru potvrdu compile-a")

        # VERZIJA MARKERA. Stroza pravila ne smeju da priznaju dokaz napravljen pod
        # slabijim, pa marker druge verzije ne vazi.
        podaci = procitaj_marker(put)
        podaci["verzija"] = MARKER_VERZIJA - 1
        upisi_marker(podaci, put)
        tvrdi(any("verzije" in n for n in zahtevaj()),
              "MARKER: marker stare verzije se priznaje")
        upisi()

        # --- TOCTOU: kontekst se snima PRE run-a --------------------------
        #
        # Otisci racunati na KRAJU opisuju stanje diska posle testova, a ne ono
        # sto je testirano. Prolaz traje 20-60 minuta i razvoj ide paralelno, pa
        # je prozor stvaran. Upis je zato fail-closed: run moze biti zelen, a
        # marker se ne upisuje.
        tvrdi("nema snimka konteksta" in zabelezi_prolaz(
                  ZELEN, 0, put=put, src_dir=src, koren=tmp),
              "MARKER: upis BEZ snimka konteksta se izvrsava")

        def toctou(izmena, opis):
            k = snimi_kontekst(src, tmp)      # snimak PRE "run-a"
            izmena()                          # ... pa se nesto promeni ...
            poruka = zabelezi_prolaz(ZELEN, 0, kontekst=k, put=put,
                                     src_dir=src, koren=tmp)
            tvrdi("TOKOM run-a" in poruka,
                  "MARKER: %s TOKOM run-a se upisuje kao dokaz" % opis)
            upisi()                           # vrati marker u zeleno stanje

        def dopisi(rel, tekst):
            def f():
                with io.open(os.path.join(tmp, *rel.split("/")), "a",
                             newline="") as fh:
                    fh.write(tekst)
            return f

        def izmeni_izvor():
            with io.open(os.path.join(src, "modTest.bas"), "a",
                         newline="") as fh:
                fh.write("' izmena tokom run-a\r\n")

        def zameni_svesku():
            with open(podrazumevana_sveska(tmp), "wb") as fh:
                fh.write(b"sveska zamenjena tokom run-a")

        toctou(izmeni_izvor, "izmenjen SRC-VBA")
        toctou(dopisi("tests/golden/G1.txt", "tokom run-a\n"),
               "izmenjen GOLDEN")
        toctou(dopisi("tools/run_vba.py", "tokom run-a\n"),
               "izmenjen RUNNER")
        toctou(dopisi("tools/vba_gate.py", "tokom run-a\n"),
               "izmenjena KAPIJA")
        toctou(zameni_svesku, "zamenjena SVESKA")

        # Sadrzaj sveske se cita iz TEMP KOPIJE (ono sto Excel otvara), a
        # putanja ostaje IZVORNA -- da se dokaz moze uporediti sa svescom koja i
        # dalje stoji na disku.
        kopija = os.path.join(tmp, "temp_kopija.xlsm")
        with open(kopija, "wb") as fh:
            fh.write(b"kopija koju Excel otvara")
        ks = kontekst_sveske(podrazumevana_sveska(tmp), kopija)
        tvrdi(ks["putanja"] == os.path.realpath(podrazumevana_sveska(tmp)),
              "SVESKA: putanja se uzima iz temp kopije umesto iz izvora")
        tvrdi(ks["otisak"] == _hash_sirov(kopija),
              "SVESKA: sadrzaj se ne cita iz temp kopije koju Excel otvara")

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
    ap.add_argument("--sveska", metavar="PUTANJA",
                    help="uz --require-green: priznaj dokaz napravljen nad TOM "
                         "svescom (putanja, ne ime; podrazumevano fixture)")
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
                                 trazi_compile=args.require_compile,
                                 sveska=args.sveska)
        if not nalazi:
            delovi = ugovor_delovi()
            print("dokazano: izvor %s, ugovor %s"
                  % (delovi["izvor"][:12], otisak_ugovora(delovi)[:12]))
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
