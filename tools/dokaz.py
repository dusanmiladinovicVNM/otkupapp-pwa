#!/usr/bin/env python3
"""Pun dvosmerni dokaz: svaka sabotaza iz kataloga mora da POKAZE crveno.

CLAUDE.md paragraf 5 trazi da se posle izmene pusti ceo dvosmerni dokaz i tvrdi
da je broj CRVENIH jednak broju sabotaza. Do sada je to bio rucni ritual --
skripta iz scratchpada, pa se u praksi vrteo samo podskup. Rezultat: deset sidara
je istrunulo neprimeceno, jedna sabotaza je prestala da obara ista, a "36 od 39"
se citalo kao zeleno (v. docs/engineering/postmortems/2026-08-verifikacija.md 10).

    python tools/dokaz.py                      # ceo katalog (satima!)
    python tools/dokaz.py modOtkupUI.bas       # samo sabotaze nad tim fajlom
    python tools/dokaz.py mreza-podnozje       # samo sabotaze sa tim prefiksom
    python tools/dokaz.py amb-pisac --grupe    # vise sabotaza u JEDNOM prolazu
    python tools/dokaz.py --grupe 8 --plan     # samo plan grupa, bez Excela

CETIRI STVARI KOJE OVAJ ALAT TVRDI, a ne samo gleda:

1. BAZA MORA BITI ZELENA PRE PRVE MUTACIJE. Bez toga alat ne dokazuje "mutacija
   je izazvala crveno" nego samo "posle mutacije postoji crveno" -- a to nisu
   iste tvrdnje. Test koji vec pada iz nekog treceg razloga proglasio bi svaku
   sabotazu nad sobom dokazanom, ukljucujuci onu koja ne radi nista.

2. IZVOR SE VRACA I KAD RUN PUKNE. Mutacija namerno kvari radni izvor, pa
   ciscenje ide kroz `finally`: timeout Excela ili Ctrl+C usred prolaza inace
   ostavlja namerno pokvaren `src-vba/`. Posle svakog vracanja se poredi potpis
   celog `src-vba` -- ako se ne poklopi, dokaz STAJE odmah, jer bi sve merene
   posle toga islo nad pokvarenim kodom.

3. PALA JE BAS NJENA TVRDNJA, ne samo njen test. Ime testa nije dovoljno:
   AssertEq puca na PRVOM padu, pa sabotaza koja usput obori raniju, uzgrednu
   tvrdnju ostavlja ciljanu NEIZVRSENOM -- a izlaz i dalje nosi ime pravog testa
   (zamka 6). Zato peti clan n-torke mora da se nadje u poruci koja je pala.
   To ga cini merenom vrednoscu, a ne komentarom: tekst koji vise ne opisuje
   ono sto pada je nalaz, ne sitnica.

4. GRUPNI PROLAZ MERI MANJE OD POJEDINACNOG, i tako se i prijavljuje. `--grupe`
   pusta vise mutacija u JEDNOM prolazu suite-a. Time dokazuje da je svaka
   tvrdnja osetljiva na mutacije grupe ZAJEDNO -- ne i da je bas njena mutacija
   oborila bas nju. Verdikt se zato zove `DOKAZANO (grupno)`, a ne `DOKAZANO`:
   pun pojedinacni dokaz je isti poziv bez `--grupe`. Pravila grupisanja i ono
   sto grupni prolaz NE sme da propusti stoje nad `naprav_grupe`.

Banka-suite ne ispisuje ime testa uz pad, ali svaka njena tvrdnja nosi stabilan
prefiks ("T21 izabran placen blok: ..."), pa se identitet vadi iz njega. Bez toga
bi za te sabotaze tvrdnja bila samo "nesto je palo".

JEFTINA POLOVINA ovoga je `python tools/sabotaza.py --proveri-sidra` (ide i kroz
vba_check, dakle kroz PostToolUse hook): hvata zastarela sidra i pogresna imena
testova za sekundu. Ovaj alat je jedini koji zna da li sabotaza STVARNO nesto
obara -- i traje satima nad celim katalogom, pa se pusta nad onim sto je izmena
dirala.
"""
import argparse
import hashlib
import importlib.util
import os
import re
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_VBA = os.path.join(ROOT, "src-vba")
SUITE_BANKA = "RunBankaImportTestSuite"
SUITE_ALL = "RunAllTests"
SUITE_BFP = "RunBusinessFlowProSuite"

# Podrazumevana velicina grupe kad se `--grupe` da bez broja. Nije izvedena iz
# merenja nego je granica rizika: sto je grupa veca, to grupni prolaz tvrdi
# manje (v. tacka 4 u docstring-u). Nad celim katalogom je 6 dovoljno -- pravila
# grupisanja ionako vezu grupu na ~5 clanova, pa veci broj skoro nista ne doda.
GRUPA_PODRAZUMEVANO = 6


def _modul_sabotaza():
    put = os.path.join(ROOT, "tools", "sabotaza.py")
    spec = importlib.util.spec_from_file_location("_sab_za_dokaz", put)
    modul = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modul)
    return modul


def _otisak() -> str:
    """Potpis celog src-vba -- da se vidi da je posle svakog vracanja isti."""
    h = hashlib.sha256()
    for ime in sorted(os.listdir(SRC_VBA)):
        h.update(ime.encode())
        with open(os.path.join(SRC_VBA, ime), "rb") as fh:
            h.update(fh.read())
    return h.hexdigest()[:16]


def _pusti(*a, timeout=1200):
    return subprocess.run(a, cwd=ROOT, capture_output=True, text=True,
                          timeout=timeout)


def _suite_za(test: str, sab=None) -> str:
    """Suite se bira po MODULU u kome test zivi, ne po obliku imena.

    Prefiks imena je lagao: sve sto ne pocinje sa "T_" islo je u banka-suite, pa
    bi test iz modBusinessFlowProTests bio trazen u suiti u kojoj ne postoji --
    ona prodje ZELENO, i sabotaza izgleda kao da nista ne meri. sabotaza.py to
    vec resava (_suita_testa nad _SUITA_PO_MODULU); ovde se samo koristi, da
    preslikavanje modul -> suite ima JEDNU definiciju.
    """
    if sab is not None:
        return sab._suita_testa(test)
    return SUITE_ALL if test.startswith("T_") else SUITE_BANKA


def _tokeni_banke(izlaz: str) -> list:
    """Imena testova iz banka-suite izlaza, preko stabilnog prefiksa tvrdnje.

    ReportResults ispisuje "PAO   T21 izabran placen blok: ...", jer svaka
    tvrdnja u toj suite nosi `Const S As String = "T21 ..."`. Ime testa se ne
    ispisuje, ali se broj vadi -- a on je isti onaj iz T21_....
    """
    out = []
    for red in re.findall(r"^\s*PAO\s+(.*)$", izlaz, re.M):
        m = re.match(r"\s*(T\d+)\b", red)
        out.append((m.group(1) if m else "?", red))
    return out


def _pali(izlaz: str, suite: str) -> list:
    if suite == SUITE_ALL:
        return re.findall(r"^\s*FAIL (\S+) -- (.*)$", izlaz, re.M)
    if suite == SUITE_BFP:
        # BFP ne ispisuje ime Sub-a nego NAZIV TVRDNJE (LogFail prima bas njega),
        # pa identitet nosi tvrdnja. Zato katalog za BFP mora da nosi TACAN
        # tekst tvrdnje, ne podniz -- podniz se ovde prijavi kao "NE OBARA SVOJ
        # TEST", dakle glasno, ne tiho.
        # Separator je " :: ", ne " -- ": tekst tvrdnje sme da sadrzi " -- ", pa
        # bi se ime na njemu odseklo i sabotaza koja radi savrseno bi bila
        # prijavljena kao "NE OBARA SVOJ TEST". modBusinessFlowProTests zato
        # pise " :: " -- isti separator koji vec koristi u Debug.Print.
        return [(t.strip(), t.strip())
                for t in re.findall(r"^\s*FAIL (.*?)(?: :: .*)?$", izlaz, re.M)]
    return _tokeni_banke(izlaz)


def _kljuc_testa(test: str, suite: str, tvrdnja: str = "") -> str:
    """Sta se poredi sa onim sto je palo."""
    if suite == SUITE_ALL:
        return test
    if suite == SUITE_BFP:
        return tvrdnja.strip()               # v. _pali: BFP pise tvrdnju
    m = re.match(r"(T\d+)_", test)          # T21_IzabranPlacenBlok... -> T21
    return m.group(1) if m else test


def _baza_zelena(suite: str) -> tuple:
    """(je_zelena, opis). Tvrdi se, ne gleda se."""
    try:
        r = _pusti(sys.executable, "tools/run_vba.py", "--suite", suite)
    except subprocess.TimeoutExpired:
        return False, f"{suite}: timeout"
    izlaz = r.stdout + r.stderr

    # Red je oznacen imenom suite-a, pa se dve suite u istom run-u ne mesaju.
    m = re.search(r"TESTS\s+" + re.escape(suite) + r": (\d+) ukupno, (\d+) palo",
                  izlaz)
    if not m:
        return False, f"{suite}: nema oznacenog reda TESTS u izlazu"
    if int(m.group(2)) != 0:
        return False, f"{suite}: {m.group(0)}"
    return True, m.group(0)


# --- gde sidro zivi ---------------------------------------------------------
#
# Procedura u kojoj je sidro je jedino sto se o "kom kodu sabotaza pripada" moze
# utvrditi BEZ Excela, a dovoljno je grubo da ne laze: dve mutacije u istoj
# proceduri su po pravilu dva susedna uslova iste odluke, pa jedna lako obori
# tvrdnju druge. Fajl je za to previse krupan (svih 20 amb-pisac sabotaza zivi u
# dva fajla), a red previse sitan.
_PROC_KES = {}


def _procedura(fajl: str, sidro: str, sab) -> str:
    """'fajl::ImeProcedure' u kojoj sidro zivi, ili None ako se ne nalazi.

    None je ovde ZNACAJAN ishod, ne greska: sidro koje se ne nalazi je nalaz za
    `--proveri-sidra`, a ovde znaci "ne znam gde je" -- pa se takva sabotaza
    nikad ne grupise, nego meri sama.
    """
    if fajl not in _PROC_KES:
        try:
            tekst = sab._procitaj(os.path.join(SRC_VBA, fajl))[0]
        except (OSError, UnicodeDecodeError, ValueError):
            _PROC_KES[fajl] = None
        else:
            _PROC_KES[fajl] = (tekst, sab._procedure_po_redu(tekst))
    if _PROC_KES[fajl] is None:
        return None

    tekst, procedure = _PROC_KES[fajl]
    poz = tekst.find("\n" + sidro)          # isto pravilo kao _pogodaka (zamka 2)
    if poz < 0:
        return None
    red = tekst[:poz + 1].count("\n")
    for ime, _vrsta, prvi, poslednji in procedure:
        if prvi <= red <= poslednji:
            return "%s::%s" % (fajl, ime)
    return "%s::<deklaracije>" % fajl


# --- grupisanje -------------------------------------------------------------


def naprav_grupe(stavke: list, cap: int) -> list:
    """Grupe koje JEDAN prolaz suite-a sme da meri zajedno.

    Cena pojedinacnog dokaza je prolaz suite-a po sabotazi (BFP ~150s), a ne
    sama mutacija. Vise mutacija u jednom prolazu zato skracuje dokaz onoliko
    koliko ih se sme spojiti. Sta se SME, odlucuju cetiri tvrda pravila -- i
    svako postoji zbog nacina na koji bi grupni prolaz inace LAGAO:

    1. ISTA SUITE. Prolaz je jedan poziv jedne suite; clan ciji test u njoj ne
       postoji ne bi bio ni izvrsen, a prolaz bi prosao zeleno.

    2. RAZLICIT KLJUC poredjenja (`_kljuc_testa`). Dve sabotaze koje se mere
       istim kljucem se iz izlaza ne mogu razluciti -- obe bi bile "priznate"
       zato sto je palo nesto. To vec postoji u katalogu kao priznat nalaz
       (POZNATI_NALAZI: dve sabotaze obore bas istu poruku).

    3. RAZLICIT TEST. U modTest `AssertEq` DIZE gresku, pa test staje na prvom
       padu: druga tvrdnja istog testa se ne bi ni izvrsila. U BFP tvrdnje ne
       prekidaju test (LogFail), ali ga prekida svaka greska iz tela -- a bas to
       vise ovih sabotaza i izaziva. Nad celim katalogom ovo pravilo ne kosta
       nista (118 prolaza sa njim, 119 bez), a na rezu sa malo testova i mnogo
       sabotaza kosta: amb-pisac je 9 prolaza sa njim, 5 bez. Placa se, jer je
       jedino pravilo koje iskljucuje clanove koji dele PUTANJU IZVRSAVANJA.

    4. RAZLICITA PROCEDURA (i poznata -- v. `_procedura`). Dva uslova iste
       odluke su najcesci oblik "druga kapija obori tvrdnju prve": na #400 su
       dva takva nalaza izasla iz jednog kruga review-a. Isti fajl nije dovoljno
       grubo sito, red nije dovoljno krupno.

    Sto grupni prolaz NE propusta, i zato ne trazi labavija pravila: clan koji
    nije oborio BAS SVOJU tvrdnju ne dobija priznanje iz grupe nego se ponavlja
    SAM (v. `_prolaz`). Istrunulo sidro, sabotaza koja ne obara nista i sabotaza
    koja obara tudju tvrdnju zato prolaze kroz pojedinacno merenje kao i danas.
    Ostaje tacno jedna stvar koju grupa ne ume: da razluci "njena mutacija je
    oborila njenu tvrdnju" od "neciji partner je oborio njenu tvrdnju". Zato
    verdikt nosi oznaku `(grupno)`.

    Redosled je iz kataloga, a punjenje "prvo mesto koje prima" -- plan je zato
    determinisan i dva poziva daju istu podelu.
    """
    # Precica radi jasnoce, ne mehanizam: i bez nje bi `len(g) >= cap` nad cap=1
    # odbio svako pridruzivanje, pa je ishod isti. Zato se ovaj red ne moze
    # pokazati crvenim -- i zato ne stoji kao pravilo iznad.
    if cap < 2:
        return [[s] for s in stavke]
    grupe = []
    for s in stavke:
        if not s["proc"]:
            grupe.append([s])                # ne znam gde je -> meri se sama
            continue
        for g in grupe:
            if len(g) >= cap:
                continue
            if g[0]["suite"] != s["suite"]:
                continue
            if any(x["kljuc"] == s["kljuc"] or x["test"] == s["test"]
                   or x["proc"] == s["proc"] or not x["proc"] for x in g):
                continue
            g.append(s)
            break
        else:
            grupe.append([s])
    return grupe


def _oceni(s: dict, pali: list, ocekivani: set) -> tuple:
    """(priznato, stanje) za jednu sabotazu iz izlaza JEDNOG prolaza.

    Ista funkcija za pojedinacan i za grupni prolaz, da se pravilo ne napise
    dva puta pa se razidje -- isti razlog zbog koga `_pogodaka` u sabotaza.py
    dele primena i staticka provera.

    `ocekivani` su kljucevi SVIH clanova prolaza. Kod pojedinacnog merenja je to
    jedan kljuc, pa je "uz jos N testa" isto sto i pre; u grupi se tako iz tog
    broja izuzmu partneri, koji nisu uzgredna steta nego ono sto je trazeno.
    """
    if not pali:
        return False, "NE OBARA NISTA"
    kljuc = s["kljuc"]
    imena = sorted({p0 for p0, _ in pali})
    # Poredi se SAMO ono sto je palo u njenom testu. Siroka sabotaza obori i
    # druge testove, pa bi tvrdnja iz TUDJEG testa inace mogla da je "potvrdi".
    poruke = " | ".join(p1 for p0, p1 in pali if p0 == kljuc)
    if kljuc not in imena:
        return False, "NE OBARA SVOJ TEST, nego: " + ", ".join(imena)
    if not s["tvrdnja"]:
        return False, "KATALOG NEMA TVRDNJU -- nema sta da se poredi"
    if s["tvrdnja"].lower() not in poruke.lower():
        # Pravi test a pogresna tvrdnja NIJE dokaz: ciljana tvrdnja mozda nije
        # ni izvrsena (AssertEq puca na prvom padu).
        return False, "PALA DRUGA TVRDNJA: " + poruke[:120]
    strani = [i for i in imena if i not in ocekivani]
    if strani:
        return True, "OK (uz jos %d testa)" % len(strani)
    return True, "OK"


def _prolaz(clanovi: list, pre: str, grupno: bool) -> dict:
    """Jedan prolaz suite-a nad mutacijama svih clanova.

    Grupa od jednog clana JE pojedinacno merenje, pa je funkcija jedna. Razlika
    je samo u tome sta znaci neuspela primena: kod jednog clana je to nalaz
    (istrunulo sidro), a u grupi moze biti i SUDAR SIDARA -- partner je pregazio
    tekst na koji je ovo sidro zakaceno. Takav clan se zato ne prijavljuje kao
    APPLY-FAIL nego se vraca na pojedinacno merenje, gde pitanja o partnerima
    nema. Isto za clana koji u grupi nije oborio svoju tvrdnju.

    Vraca {"stop": razlog ili "", "ishodi": [(stavka, priznato, stanje, crveno)],
           "ostatak": [(stavka, razlog)]}.
    """
    rez = {"stop": "", "ishodi": [], "ostatak": []}

    primenjeni = []
    for s in clanovi:
        p = _pusti(sys.executable, "tools/sabotaza.py", s["ime"])
        if p.returncode == 0:
            primenjeni.append(s)
        elif grupno:
            rez["ostatak"].append((s, "sidro se ne primenjuje uz partnere"))
        else:
            rez["ishodi"].append(
                (s, False, "APPLY-FAIL -- v. sabotaza.py --proveri-sidra", False))

    if not primenjeni:
        # Nista nije upisano u izvor, pa nema ni sta da se vraca ni sta da se
        # meri -- `primeni` je ili zamenio tacno jednom ili nije dirao fajl.
        return rez

    # Jedan primenjen clan je -- bez obzira na to kako je grupa planirana --
    # tacno pojedinacno merenje: u izvoru nema partnera, pa nema ni tvrdnje
    # koju bi partner mogao da obori. Zato se tako i ocenjuje (i ne vraca se
    # u red za ponovno merenje, sto bi bio isti prolaz drugi put).
    sam = len(primenjeni) == 1
    suite = primenjeni[0]["suite"]
    ocekivani = {s["kljuc"] for s in primenjeni}

    pali, greska = [], ""
    try:
        run = _pusti(sys.executable, "tools/run_vba.py", "--suite", suite)
        pali = _pali(run.stdout + run.stderr, suite)
    except subprocess.TimeoutExpired:
        greska = "TIMEOUT suite"
    except KeyboardInterrupt:
        greska = "PREKID"
    finally:
        # Izvor se vraca i kad je run pukao -- inace radni tree ostaje namerno
        # pokvaren. `--vrati` vraca SVE zatecene sabotaze, pa radi i nad grupom.
        v = _pusti(sys.executable, "tools/sabotaza.py", "--vrati")

    sada = _otisak()
    if v.returncode != 0 or sada != pre:
        rez["stop"] = "REVERT-FAIL (potpis %s)" % sada
        for s in primenjeni:
            rez["ishodi"].append((s, False, "REVERT-FAIL", False))
        return rez

    if greska:
        if grupno and not sam and greska != "PREKID":
            # Timeout grupe ne kaze KOJI clan ga je izazvao, a najcesci uzrok je
            # sabotaza koja obori COMPILE (zamka 4 u sabotaza.py): Excel ostane u
            # [break] i prolaz visi do timeout-a. Pojedinacno merenje to
            # lokalizuje -- krivac padne sam, a ostali se izmere. PREKID se ne
            # ponavlja: Ctrl+C znaci da operater hoce da stane.
            for s in primenjeni:
                rez["ostatak"].append((s, greska))
            return rez
        for s in primenjeni:
            rez["ishodi"].append((s, False, greska, False))
        if greska == "PREKID":
            rez["stop"] = greska
        return rez

    for s in primenjeni:
        priznato, stanje = _oceni(s, pali, ocekivani)
        if priznato or sam or not grupno:
            rez["ishodi"].append((s, priznato, stanje, bool(pali)))
        else:
            rez["ostatak"].append((s, stanje))
    return rez


def _izaberi(sab, filter_: list) -> list:
    """Stavke po filteru, sa svim sto grupisanje i ocena traze."""
    stavke = []
    for ime, (fajl, sidro, _novo, test, tvrdnja) in sab.SABOTAZE.items():
        if filter_ and not (fajl in filter_ or
                            any(ime.startswith(f) for f in filter_)):
            continue
        suite = _suite_za(test, sab)
        stavke.append({
            "ime": ime, "fajl": fajl, "test": test, "tvrdnja": tvrdnja,
            "suite": suite,
            "kljuc": _kljuc_testa(test, suite, tvrdnja),
            "proc": _procedura(fajl, sidro, sab),
        })
    return stavke


def _plan(stavke: list, cap: int) -> int:
    """Podela na grupe bez ijednog poziva Excela -- i tvrdnja da je potpuna."""
    grupe = naprav_grupe(stavke, cap)

    # Plan koji izgubi ili udvoji sabotazu bi tiho smanjio imenilac dokaza.
    # Ovo je jedina tvrdnja koju plan daje, pa stoji ovde, a ne u komentaru.
    pokriveni = [s["ime"] for g in grupe for s in g]
    assert sorted(pokriveni) == sorted(s["ime"] for s in stavke), \
        "plan ne pokriva svaku sabotazu tacno jednom"

    for i, g in enumerate(grupe, 1):
        print("grupa %3d  %-28s %d" % (i, g[0]["suite"], len(g)))
        for s in g:
            print("    %-46s %-40s %s"
                  % (s["ime"], s["test"], s["proc"] or "<sidro nije nadjeno>"))
    prolaza = len(grupe)
    print("\nsabotaza %d, prolaza %d (najvise %d po grupi) -- %.1fx manje "
          "prolaza suite-a" % (len(stavke), prolaza, cap,
                               len(stavke) / float(prolaza)))
    return 0


def main(argv: list) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("filter", nargs="*",
                    help="ime fajla (modX.bas) ili prefiks imena sabotaze")
    ap.add_argument("--grupe", nargs="?", type=int, const=GRUPA_PODRAZUMEVANO,
                    default=1, metavar="N",
                    help="najvise N sabotaza u jednom prolazu suite-a "
                         "(podrazumevano %d uz golo --grupe, 1 = bez grupisanja)"
                         % GRUPA_PODRAZUMEVANO)
    ap.add_argument("--plan", action="store_true",
                    help="ispisi podelu na grupe i stani (ne trazi Excel)")
    ap.add_argument("--self-test", action="store_true",
                    help="dokazi da pravila grupisanja, ocena i orkestracija prolaza grizu")
    args = ap.parse_args(argv)

    if args.self_test:
        return _self_test()
    if args.grupe < 1:
        print("--grupe mora biti >= 1", file=sys.stderr)
        return 2

    sab = _modul_sabotaza()
    poznati_spisak = getattr(sab, "POZNATI_NALAZI_DOKAZ", {})
    stavke = _izaberi(sab, args.filter)

    if not stavke:
        print("filter ne pogadja nijednu sabotazu", file=sys.stderr)
        return 2

    if args.plan:
        return _plan(stavke, args.grupe)

    print("sabotaza: %d" % len(stavke), flush=True)

    # --- KAPIJA: baza mora biti zelena pre prve mutacije --------------------
    potrebne = sorted({s["suite"] for s in stavke} | {SUITE_ALL})
    for suite in potrebne:
        ok, opis = _baza_zelena(suite)
        print("BAZNO: %s" % opis, flush=True)
        if not ok:
            print("STOP: baza nije zelena. Dokaz bi merio crveno koje sabotaza "
                  "nije izazvala.", file=sys.stderr)
            return 2

    pre = _otisak()
    print("potpis izvora: %s" % pre, flush=True)

    crvenih, lose, poznati = 0, [], []
    grupno_priznatih, stop = 0, ""

    def upisi(ishodi, uvlaka=""):
        """Zajednicki ispis i knjizenje za grupni i pojedinacni prolaz."""
        nonlocal crvenih
        for s, priznato, stanje, crveno in ishodi:
            if crveno:
                crvenih += 1
            if not priznato:
                lose.append((s["ime"], stanje))
            print("%s%-46s %s" % (uvlaka, s["ime"], stanje), flush=True)

    grupe = naprav_grupe(stavke, args.grupe)
    ponavljaju = []
    if args.grupe > 1:
        print("grupisanje: %d prolaza nad %d sabotaza (najvise %d po grupi); "
              "clan koji u grupi ne obori SVOJU tvrdnju meri se posle sam"
              % (len(grupe), len(stavke), args.grupe), flush=True)

    for i, g in enumerate(grupe, 1):
        if len(g) > 1:
            print("GRUPA %d/%d  %s  clanova %d"
                  % (i, len(grupe), g[0]["suite"], len(g)), flush=True)
        ishod = _prolaz(g, pre, len(g) > 1)
        if len(g) > 1:
            grupno_priznatih += sum(1 for _s, p, _st, _c in ishod["ishodi"] if p)
        upisi(ishod["ishodi"], "  " if len(g) > 1 else "")
        for s, razlog in ishod["ostatak"]:
            print("  %-46s -> ponavlja se sam (%s)" % (s["ime"], razlog),
                  flush=True)
            ponavljaju.append(s)
        if ishod["stop"]:
            stop = ishod["stop"]
            break

    if ponavljaju and not stop:
        print("PONAVLJA SE SAM: %d" % len(ponavljaju), flush=True)
        for s in ponavljaju:
            ishod = _prolaz([s], pre, False)
            upisi(ishod["ishodi"])
            for s2, razlog in ishod["ostatak"]:
                # Pojedinacno merenje nema partnere, pa nema ni sta da vrati u
                # red -- ali ako ga ikad vrati, sabotaza bi tiho ispala iz
                # imenioca. Zato je to nalaz, a ne tiho preskakanje.
                lose.append((s2["ime"], "NIJE IZMERENA: %s" % razlog))
                print("%-46s NIJE IZMERENA: %s" % (s2["ime"], razlog), flush=True)
            if ishod["stop"]:
                stop = ishod["stop"]
                break

    if stop:
        if stop.startswith("REVERT-FAIL"):
            print("STOP: izvor nije vracen u pocetno stanje. Sve mereno posle "
                  "ovoga islo bi nad pokvarenim kodom.", file=sys.stderr)

    # Priznat, zapisan nalaz sa vlasnikom ne obara gejt -- crven alat koji svi
    # nauce da preskoce ne cuva nista. Upis koji nista ne pokriva je isto nalaz.
    ostali = []
    pokriveni = set()
    for ime, sta in lose:
        if sab.poznat_nalaz(ime, sta, poznati_spisak):
            poznati.append((ime, sta))
            pokriveni.add(ime)
        else:
            ostali.append((ime, sta))
    izabrana_imena = {s["ime"] for s in stavke}
    for ime in sorted(set(poznati_spisak) & izabrana_imena - pokriveni):
        ostali.append((ime, "POZNATI_NALAZI_DOKAZ['%s'] ne pokriva nijedan "
                            "nalaz -- obrisi ga ili ispravi ime" % ime))
    lose = ostali

    posle = _otisak()
    # PRIZNAT NALAZ SE VADI I IZ IMENIOCA, ne samo iz spiska problema.
    #
    # Bez toga priznanje radi samo za polovinu vrsta nalaza: "PALA DRUGA TVRDNJA"
    # jeste crvena pa se broji, a "NE OBARA NISTA" po definiciji nikad nije -- pa
    # je crvenih uvek manje od ukupno i verdikt ostaje NIJE DOKAZANO ma sta pisalo
    # u POZNATI_NALAZI_DOKAZ. Zapisan nalaz sa vlasnikom je tako i dalje drzao
    # alat crvenim -- tacno ono sto komentar iznad zabranjuje.
    #
    # Zloupotreba je pokrivena: upis koji ne pokriva nijedan nalaz je i sam nalaz
    # (v. gore), pa priznanje ne moze da prezivi popravku koju opisuje.
    print("\ncrvenih: %d / sabotaza: %d%s"
          % (crvenih, len(stavke),
             " (priznatih: %d)" % len(pokriveni) if pokriveni else ""))
    print("izvor pre/posle: %s / %s -> %s"
          % (pre, posle, "IDENTICAN" if pre == posle else "RAZLIKA!"))
    if grupno_priznatih:
        print("grupno izmereno: %d / %d -- tvrdnja je sprovedena nad mutacijama "
              "grupe ZAJEDNO, ne nad svakom posebno; pun pojedinacni dokaz je "
              "isti poziv bez --grupe" % (grupno_priznatih, len(stavke)))
    for ime, sta in poznati:
        print(" POZNATO: %s -> %s" % (ime, sta))
    for ime, sta in lose:
        print(" PROBLEM: %s -> %s" % (ime, sta))
    ok = (not lose and not stop and pre == posle
          and crvenih >= len(stavke) - len(pokriveni))
    print("=== %s ===" % ("NIJE DOKAZANO" if not ok else
                          "DOKAZANO (grupno)" if grupno_priznatih else
                          "DOKAZANO"))
    return 0 if ok else 1


# --- self-test: dokaz u oba smera nad pravilima ------------------------------
#
# CLAUDE.md paragraf 5 trazi dvosmerni dokaz kad se menja SAM CHECKER. Za
# grupisanje taj dokaz ne moze da bude sabotaza u src-vba (sabotaze obaraju VBA
# testove, a ovo su pravila u Python-u), pa stoji ovde: svako pravilo se meri i
# u zelenom i u crvenom smeru -- par koji se razlikuje SAMO po toj osi mora da
# se grupise, a par koji se po njoj poklapa ne sme. Pravilo bez crvenog smera
# bi inace moglo da bude i konstanta True.
#
# Ne cita ni src-vba ni Excel, pa ide i kroz vba_check (PostToolUse hook).


def _s(ime, **kw) -> dict:
    osnova = {"ime": ime, "fajl": "modA.bas", "test": "T_A", "tvrdnja": "t",
              "suite": SUITE_ALL, "kljuc": "T_A", "proc": "modA.bas::P"}
    osnova.update(kw)
    return osnova


class _LazniOdgovor:
    def __init__(self, rc=0, stdout=""):
        self.returncode, self.stdout, self.stderr = rc, stdout, ""


def _lazni_prolaz(clanovi, izlaz_suite, grupno=True, pukli=(), potpis="P",
                  timeout=False):
    """Pokreni `_prolaz` bez Excela. Vraca (rez, pozivi).

    Pravila grupisanja i ocena su ciste funkcije i mere se direktno, ali
    ORKESTRACIJA nije: primeni sve -> pusti suite TACNO JEDNOM -> vrati sve ->
    oceni, pa clan koji nije priznat vrati na pojedinacno merenje. Bez ovoga bi
    taj redosled prvi put bio izvrsen nad pravim Excelom, gde jedan prolaz traje
    minutima i gde greska u njemu izgleda kao nalaz nad VBA kodom.
    """
    pozivi = []

    def lazni_pusti(*a, **kw):
        pozivi.append(tuple(a[1:]))
        if a[1].endswith("run_vba.py"):
            if timeout:
                raise subprocess.TimeoutExpired(a, 1)
            return _LazniOdgovor(0, izlaz_suite)
        if a[2] != "--vrati":
            return _LazniOdgovor(2 if a[2] in pukli else 0)
        return _LazniOdgovor(0)

    pravi = _pusti, _otisak
    globals()["_pusti"] = lazni_pusti
    globals()["_otisak"] = lambda: potpis
    try:
        rez = _prolaz(clanovi, "P", grupno)
    finally:
        globals()["_pusti"], globals()["_otisak"] = pravi
    return rez, pozivi


def _self_test(tiho: bool = False) -> int:
    nalazi = []

    def tvrdi(uslov, opis):
        if not uslov:
            nalazi.append(opis)

    def zajedno(a, b, cap=6):
        g = naprav_grupe([a, b], cap)
        return len(g) == 1 and len(g[0]) == 2

    # --- pravila grupisanja, svako u oba smera ------------------------------
    #
    # SLUCAJ SE RAZLIKUJE PO TACNO JEDNOJ OSI, i to je tvrdnja, ne stil. Par koji
    # se razlikuje po dve ose meri ono pravilo koje prvo odbije, pa ostaje zelen
    # kad se drugo ugasi -- prva verzija ovog self-testa je tako "merila" pravilo
    # o razlicitom testu, a stvarno merila pravilo o kljucu. Dvosmerni dokaz nad
    # samim pravilima je to pokazao; `sudar` ispod cini da se ne moze vratiti.
    def sudar(a, b) -> set:
        """Ose po kojima par NE SME u istu grupu -- citano iz samih stavki."""
        ose = set()
        if a["suite"] != b["suite"]:
            ose.add("suite")
        for osa in ("test", "kljuc", "proc"):
            if a[osa] == b[osa]:
                ose.add(osa)
        return ose

    cist = _s("b", test="T_B", kljuc="T_B", proc="modA.bas::Q")
    tvrdi(not sudar(_s("a"), cist), "GRUPE: osnovni par ima sudar")
    tvrdi(zajedno(_s("a"), cist),
          "GRUPE: par bez ijednog sudara se NE grupise")

    # Isti test a RAZLICIT kljuc postoji samo u BFP: tamo je kljuc tekst tvrdnje,
    # pa dve tvrdnje istog testa nose razlicite kljuceve. U modTest je kljuc IME
    # TESTA, pa se pravilo o razlicitom testu tamo ne moze ni izmeriti odvojeno
    # od pravila o kljucu. Zato slucaj za osu "test" ide kroz BFP.
    bfp = {"suite": SUITE_BFP, "test": "Test_X"}
    po_osi = {
        "test": (_s("a", kljuc="tvrdnja jedna", **bfp),
                 _s("b", kljuc="tvrdnja druga", proc="modA.bas::Q", **bfp)),
        "kljuc": (_s("a"), _s("b", test="T_B", proc="modA.bas::Q")),
        "proc": (_s("a"), _s("b", test="T_B", kljuc="T_B")),
        "suite": (_s("a"), _s("b", test="T_B", kljuc="T_B",
                              proc="modA.bas::Q", suite=SUITE_BFP)),
    }
    for osa, (a, b) in sorted(po_osi.items()):
        tvrdi(sudar(a, b) == {osa},
              "GRUPE: slucaj za osu '%s' se razlikuje po %s -- meri susedno "
              "pravilo, ne svoje" % (osa, sorted(sudar(a, b))))
        tvrdi(not zajedno(a, b), "GRUPE: sudar po osi '%s' se grupise" % osa)

    tvrdi(zajedno(_s("a", kljuc="tvrdnja jedna", **bfp),
                  _s("b", test="Test_Y", kljuc="tvrdnja druga",
                     proc="modA.bas::Q", suite=SUITE_BFP)),
          "GRUPE: dve BFP sabotaze bez sudara se NE grupisu")
    tvrdi(not zajedno(_s("a"),
                      _s("b", test="T_B", kljuc="T_B", proc=None)),
          "GRUPE: sabotaza sa NEPOZNATOM procedurom se grupise")
    tvrdi(not zajedno(_s("a", proc=None),
                      _s("b", test="T_B", kljuc="T_B", proc="modA.bas::Q")),
          "GRUPE: clan uz nepoznatu proceduru u grupi se dodaje")
    tvrdi(not zajedno(_s("a"), cist, cap=1),
          "GRUPE: cap=1 ne gasi grupisanje")
    tvrdi(len(naprav_grupe([_s("a"), cist,
                            _s("c", test="T_C", kljuc="T_C",
                               proc="modA.bas::R")], 2)) == 2,
          "GRUPE: cap se ne postuje")

    # Plan je potpun: nista se ne izgubi i nista ne udvoji.
    mnogo = [_s("i%d" % i, test="T_%d" % i, kljuc="T_%d" % i,
                proc="modA.bas::P%d" % (i % 3)) for i in range(17)]
    for cap in (1, 2, 5, 40):
        imena = [s["ime"] for g in naprav_grupe(mnogo, cap) for s in g]
        tvrdi(sorted(imena) == sorted(s["ime"] for s in mnogo),
              "GRUPE: plan (cap=%d) ne pokriva svaku sabotazu tacno jednom" % cap)

    # --- ocena jednog prolaza ----------------------------------------------
    s = _s("a", tvrdnja="pravi razlog")
    tvrdi(_oceni(s, [("T_A", "pravi razlog je pao")], {"T_A"})[0],
          "OCENA: pad svoje tvrdnje nije priznat")
    tvrdi(_oceni(s, [], {"T_A"}) == (False, "NE OBARA NISTA"),
          "OCENA: prazan izlaz je priznat")
    # Tvrdi se i RAZLOG, ne samo "nije priznato": bez imena razloga ovaj
    # slucaj obori provera tvrdnje ispod (poruke svog kljuca su prazne, pa se
    # tekst "ne nalazi"), i ostaje zelen kad se pravilo o kljucu ugasi.
    tvrdi(_oceni(s, [("T_B", "pravi razlog je pao")], {"T_A"})
          == (False, "NE OBARA SVOJ TEST, nego: T_B"),
          "OCENA: pad TUDJEG testa nije imenovan kao tudj test")
    tvrdi(_oceni(s, [("T_A", "neki drugi razlog")], {"T_A"})[1]
          .startswith("PALA DRUGA TVRDNJA"),
          "OCENA: pad DRUGE tvrdnje u svom testu nije imenovan")
    tvrdi(not _oceni(_s("a", tvrdnja=""), [("T_A", "bilo sta")], {"T_A"})[0],
          "OCENA: prazna tvrdnja u katalogu je priznata")
    tvrdi(_oceni(s, [("T_A", "pravi razlog"), ("T_X", "x")], {"T_A"})[1]
          == "OK (uz jos 1 testa)",
          "OCENA: uzgredna steta se ne prijavljuje")
    tvrdi(_oceni(s, [("T_A", "pravi razlog"), ("T_B", "x")],
                 {"T_A", "T_B"})[1] == "OK",
          "OCENA: partner iz grupe se broji kao uzgredna steta")

    # --- orkestracija prolaza (bez Excela) ----------------------------------
    g3 = [_s("a", test="T_A", kljuc="T_A", tvrdnja="razlog a", proc="modA.bas::P"),
          _s("b", test="T_B", kljuc="T_B", tvrdnja="razlog b", proc="modA.bas::Q"),
          _s("c", test="T_C", kljuc="T_C", tvrdnja="razlog c", proc="modA.bas::R")]
    SVE = ("FAIL T_A -- razlog a\n"
           "FAIL T_B -- razlog b\n"
           "FAIL T_C -- razlog c\n")

    def prolaz(clanovi, izlaz, **kw):
        try:
            return _lazni_prolaz(clanovi, izlaz, **kw)
        except Exception as e:              # pad je nalaz, ne goli traceback
            nalazi.append("PROLAZ: prolaz je pukao -- %r" % (e,))
            return {"stop": "", "ishodi": [], "ostatak": []}, []

    def suita(pozivi):
        return [p for p in pozivi if p and p[0].endswith("run_vba.py")]

    rez, pozivi = prolaz(g3, SVE)
    tvrdi(len(suita(pozivi)) == 1,
          "PROLAZ: grupa od tri clana ne pusta suite TACNO jednom")
    tvrdi(len(rez["ishodi"]) == 3 and all(p for _s2, p, _st, _c in rez["ishodi"])
          and not rez["ostatak"],
          "PROLAZ: grupa u kojoj su sve tvrdnje pale ne priznaje sve clanove")
    tvrdi(pozivi[-1][1:] == ("--vrati",),
          "PROLAZ: izvor se ne vraca posle prolaza")

    rez, _p = prolaz(g3, "FAIL T_A -- razlog a\nFAIL T_B -- razlog b\n")
    tvrdi(len(rez["ishodi"]) == 2 and all(p for _s2, p, _st, _c in rez["ishodi"]),
          "PROLAZ: clan koji nije oborio svoju tvrdnju se u grupi PRIZNAJE")
    tvrdi([s["ime"] for s, _r in rez["ostatak"]] == ["c"],
          "PROLAZ: clan koji nije oborio svoju tvrdnju se ne vraca na svoje merenje")

    rez, pozivi = prolaz(g3, SVE, pukli=("b",))
    tvrdi([s["ime"] for s, _r in rez["ostatak"]] == ["b"]
          and "partner" in rez["ostatak"][0][1],
          "PROLAZ: sudar sidara u grupi se prijavljuje kao nalaz, a ne kao sudar")
    tvrdi(len(rez["ishodi"]) == 2 and len(suita(pozivi)) == 1,
          "PROLAZ: sudar jednog clana obori merenje ostalih")

    rez, pozivi = prolaz([g3[0]], "", grupno=False, pukli=("a",))
    tvrdi(rez["ishodi"] and rez["ishodi"][0][2].startswith("APPLY-FAIL")
          and not rez["ostatak"],
          "PROLAZ: istrunulo sidro pojedinacnog merenja nije APPLY-FAIL")
    tvrdi(not suita(pozivi),
          "PROLAZ: suite se pusta i kad nijedna mutacija nije primenjena")

    # Grupa u kojoj se primenio samo JEDAN clan JE pojedinacno merenje: u izvoru
    # nema partnera, pa se clan ocenjuje tu i ne vraca u red -- inace bi isti
    # prolaz bio placen dva puta.
    rez, _p = prolaz(g3, "", pukli=("b", "c"))
    tvrdi([s["ime"] for s, _r in rez["ostatak"]] == ["b", "c"],
          "PROLAZ: sudareni clanovi se vracaju kao nalaz umesto na svoje merenje")
    tvrdi([(s["ime"], p) for s, p, _st, _c in rez["ishodi"]] == [("a", False)],
          "PROLAZ: jedini primenjen clan se vraca u red umesto da se oceni")

    rez, _p = prolaz(g3, SVE, timeout=True)
    tvrdi([s["ime"] for s, _r in rez["ostatak"]] == ["a", "b", "c"]
          and not rez["ishodi"],
          "PROLAZ: timeout grupe se ne lokalizuje pojedinacnim merenjem")
    rez, _p = prolaz([g3[0]], SVE, grupno=False, timeout=True)
    tvrdi(rez["ishodi"] and rez["ishodi"][0][2] == "TIMEOUT suite"
          and not rez["ostatak"],
          "PROLAZ: timeout pojedinacnog merenja nije nalaz")

    rez, _p = prolaz(g3, SVE, potpis="DRUGI")
    tvrdi(rez["stop"].startswith("REVERT-FAIL")
          and not any(p for _s2, p, _st, _c in rez["ishodi"]),
          "PROLAZ: nevracen izvor ne zaustavlja dokaz niti gasi priznanja")

    if nalazi:
        print("dokaz.py --self-test: %d nalaza" % len(nalazi), file=sys.stderr)
        for n in nalazi:
            print("  " + n, file=sys.stderr)
        return 1
    if not tiho:
        print("dokaz.py --self-test: pravila grupisanja, ocena i orkestracija prolaza grizu")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main(sys.argv[1:]))
    except KeyboardInterrupt:
        # Poslednja mreza: prekid izmedju dve mutacije ne sme da ostavi
        # pokvaren izvor.
        subprocess.run([sys.executable, "tools/sabotaza.py", "--vrati"], cwd=ROOT)
        print("\nprekinuto -- izvor vracen", file=sys.stderr)
        sys.exit(130)
